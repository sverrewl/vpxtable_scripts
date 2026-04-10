'For testing game:
'Press 2 to block outlane and center drain.  Press 2 again to cycle throug options or remove blocking walls
'Press and hold Y, U, or I to catch ball near right flipper.
'Release Y, U or I  to to shoot Ramp, Spinner or Right LaneBonus
'Press and hold F, G or H to catch ball and shoot at drop Targets
'

' ****************************************************************
'
' ****************************************************************
Option Explicit
Randomize

Const UseUltraDMD     = 0   '0 = Off.  1 = Enable UltraDMD. 2 - Try 2 if enabling UltraDMD crashes (Windows locale setting issue for non-English setting)
Const UseApronDMD     = 0   '0 = Off.  1 = Enable DMD on Apron
Const UseVPReelDMD      = 0
Const UseVPTextDMD      = 1


Const BallsPerGame      = 3   'Default: 3 balls
Const SongVolume      = 0.9
Const bFreePlay       = False 'True = FreePlay.  False = Coins
Const GIcolor         = "white"   'Colors:  "white", "blue"

Const BallSaverTime     = 12    'Default: 12 seconds
Const ChooseBall      = 0   ' *** Ball Settings **********
                  ' *** 0 = Normal Ball
                  ' *** 1 = White GlowBall
                  ' *** 2 = Magma GlowBall
                  ' *** 3 = Blue GlowBall
                  ' *** 4 = HDR Ball
                  ' *** 5 = Earth Ball
                  ' *** 6 = Green Glowball
                  ' *** 7 = Light blue Glowball
                  ' *** 8 = Red Glowball
                  ' *** 9 = Shiny Ball
Const FlipperPhysicsMode  = 2     '1 = VPX Flippers,   2 = NFozzy flipper tweaks
Const ResetHighScore    = 0   '0 = Keep Scores.  1 = Reset all high scores.  START TABLE ONCE to reset high scores and then set this back to 0
Const UltraDMDUpdateTime  = 5000  'UltraDMD update time (msec).  Increase value if you encounter stutter with UltraDMD on

'///////////////////////-----General Sound Options-----///////////////////////
'//  VolumeDial:
'//  VolumeDial is the actual global volume multiplier for the mechanical sounds.
'//  Values smaller than 1 will decrease mechanical sounds volume.
'//  Recommended values should be no greater than 1.
Const VolumeDial = 0.8

'============================
'  DOF Events for this table are listed at the bottom of this script
'============================


' ****************************************************************
Const Testmode        = 1   'Testing only
Const debugGeneral      = 1   'For debug only
Const debugHighScore    = 1   'For debug only
Const debugMultiball    = 1   'For debug only

Const debugTargets      = 1   'For debug only
'Const debugGrid      = 1   'For debug only
'Const debugDestroyRAD    = 1   'For debug only
'Const debugMysteryAward  = 1   'For debug only


' Improved directional sounds 2019 October : ' !! NOTE : Table not verified yet !!
' Volume devided by - lower gets higher sound
Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )
' The rest of the values are multipliers
'  .5 = lower volume
' 1.5 = higher volume
Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "lochness"
Const BallSize = 50 ' 50 is the normal size
Const BallMass = 1

' Load the core.vbs for supporting Subs and functions
LoadCoreVBS

Sub LoadCoreVBS
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Goto 0
End Sub

Sub startB2S(aB2S)
    If B2SOn Then
        Controller.B2SSetData 1, 0
        Controller.B2SSetData 2, 0
        Controller.B2SSetData 3, 0
        Controller.B2SSetData 4, 0
        Controller.B2SSetData 5, 0
        Controller.B2SSetData 6, 0
        Controller.B2SSetData 7, 0
        Controller.B2SSetData 8, 0
        Controller.B2SSetData aB2S, 1
    End If
End Sub

' Define any Constants
Const TableName = "Lochness"
Const MaxPlayers = 4
Const MaxMultiplier = 5 'limit to 4x in this game
Const MaxMultiballs = 2  ' max number of balls during multiballs

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim ReactorScore(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue
Dim HandsFreeSkillshotInsert
Dim bAutoPlunger
Dim bInstantInfo
Const Quotemode       = 0

' Define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock
Dim BallsInHole

' Define Game Flags
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim DrainBonusReady
Dim bMusicOn
Dim GIcolorOpposite

'Skillshot
Dim SkillshotReady  '0 = Off, 1 = Start, 2 = Plunged
Dim bSkillshotSelect 'used to select the skillshot you want

Dim bExtraBallWonThisBall
Dim bJustStarted

Dim plungerIM 'used mostly as an autofire plunger
'Dim ttable, cbleft, cbright
Dim Bank1Counter(4), Bank2Counter(4), Bank3Counter(4)
Dim Bank1Dropped(4,3), Bank2Dropped(4,3), Bank3Dropped(4,4)
Dim LNMCounter(4)
Dim NessieFrames, NessieBGCounter
Dim BallIsLocked
Dim BonusAmount(4)
Dim FlasherCounter1, FlasherCounter2
Dim HighScoresDisplayed
Dim obj

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************
Sub Table1_Init()
    Dim i

    Randomize
    LoadEM
  Credits=0
  If ResetHighScore = 1 then Reseths

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 43 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    'load saved values, highscore, names, jackpot
    Loadhs

  If ((bFreePlay = True) Or (Credits > 0)) Then DOF 140, DOFOn
    'Init main variables

  ' start the UltraDMD
  If UseUltraDMD > 0 Then LoadUltraDMD

    ' initalise the DMD display
    If UseVPReelDMD Then DMD_Init

    ' initialse any other flags
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bGameInPlay = False
    bAutoPlunger = False
  BallIsLocked = False
  HighScoresDisplayed=false
    bMusicOn = True
  BallLockedBlocker.IsDropped=1
    SetBallsOnPlayfield 0
    BallsInLock = 0
    BallsInHole = 0
  FlasherCounter1=0
  FlasherCounter2=0
  NessieBGCounter=0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bInstantInfo = False
  PlaySound "Bootup"

    'EndOfGame()
  'StartRainbow "all"
  'PlaySong "xxx.mp3", 2
  'ShowTableInfo
  'LightSeqAttract.Play SeqRandom, 40, 1000, 0
  'StartAttractMode 1

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' Remove the cabinet rails if in FS mode
    If Table1.ShowDT = False then
        lrail.Visible = False
        rrail.Visible = False
    Light3.visible = false
    ReelHighGame.visible=false
    ReelGameOver.visible=false
    ReelShootAgain.visible=false
    ReelTilt.visible=false
    for each obj in BIPReels
      obj.visible=false
    next
    for each obj in CreditReels
      obj.visible=false
    next
    for each obj in PlayerScore1
      obj.visible=false
    next
    for each obj in PlayerScore2
      obj.visible=false
    next
    for each obj in PlayerScore3
      obj.visible=false
    next
    for each obj in PlayerScore4
      obj.visible=false
    next
    for each obj in PlayerUpLights
      obj.visible=false
    next
    End If

  'Glowball
  ChangeBall(ChooseBall)
  If GlowBall Then GraphicsTimer.enabled = True End If
  NessieFrames=1
  AttractModeHSDisplay.enabled=true
  AttractNessieBG.enabled=true

End Sub

'******
' Section; Keys
'******
Dim CoopMode
Dim kickertest
kickertest = 8
Sub Table1_KeyDown(ByVal Keycode)
    If Keycode = ((AddCreditKey)  AND (bFreePlay = False))Then
        Credits = Credits + 1
    If B2SOn then
      Controller.B2SSetScorePlayer5 Credits
    end if
    SplitCreditsForDT Credits,CreditReels
        DOF 140, DOFOn
        If(Tilted = False)Then
      If hsbModeActive = False Then
'       DMDFlush
'       DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "CREDITS: " & Credits), eNone, 500, True, "fx_coin"
        PlaySound "fx_coin"
        UDMD "CREDITS: " & Credits, "", 500
        TextDMD "CREDITS: " & Credits

        If NOT bGameInPlay Then ShowTableInfo
      End If
        End If
    End If

  If keycode = PlungerKey Then
    Plunger.Pullback
    PlaySound "fx_plungerpull"
  End If

' ************ DEBUG KEYS Sect  *******
' ************ DEBUG KEYS Sect  *******
  If Keycode = 3 Then
    BlockerWalls
  End If
  If TestMode = 1 Then
    If keycode = 17 Then  'W key
      debug.print "*****SUB:" & "Table1_KeyDown, Testmode:"
    End If
    If keycode = 18 Then  'E key
      debug.print "*****SUB:" & "Table1_KeyDown, Testmode:"
    End If
    If keycode = 19 Then  'R key
      debug.print "*****SUB:" & "Table1_KeyDown, Testmode:"
    End If
    If keycode = 20 Then  'T key
      debug.print "*****SUB:" & "Table1_KeyDown, Testmode:"
    End If

    if keycode = 21 then kickerdebug.enabled = true 'Y key:  RAMP shot: press 3 to put up blocking walls.  press and hold to catch ball.  release to shoot ball
    if keycode = 22 then kickerdebug.enabled = true 'U key:  SPINNER shot: press 3 to put up blocking walls.  press and hold to catch ball.  release to shoot ball
    if keycode = 23 then kickerdebug.enabled = true 'I key:  TOPLANE shot: press 3 to put up blocking walls.  press and hold to catch ball.  release to shoot ball
    'if keycode = 25 then kickerdebug.enabled = true  'P key:
    if keycode = 33 then kickerdebug.enabled = true 'F key:  Left target bank
    if keycode = 34 then kickerdebug.enabled = true 'G key:  Middle Target Bank
    if keycode = 35 then kickerdebug.enabled = true 'H key:  Right Target Bank

    If keycode = 30 Then 'A key
      debug.print "*****SUB:" & "Table1_KeyDown, Testmode:"
    End If
    If keycode = 31 Then 'S key
      debug.print "*****SUB:" & "Table1_KeyDown, Testmode:"
    End If

  End If
' ************ DEBUG KEYS Sect  *******
' ************ DEBUG KEYS Sect  *******

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    ' Table specific
    ' Normal flipper action
    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then Nudge 90, 6:Playsound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 6:Playsound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 7:Playsound SoundFX("fx_nudge",0), 0, 1, 1, 0.25:CheckTilt
        If keycode = MechanicalTilt Then Playsound SoundFX("fx_nudge",0),0,1,1,0,25:CheckTilt

    If keycode = LeftFlipperKey Then SolLFlipper 1:InstantInfoTimer.Enabled = False
    If keycode = RightFlipperKey Then SolRFlipper 1:InstantInfoTimer.Enabled = False

    If keycode = StartGameKey Then
      If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = True))Then

        If(bFreePlay = True)Then
          PlayersPlayingGame = PlayersPlayingGame + 1
          If B2SOn Then Controller.B2SSetScorePlayer PlayersPlayingGame, 0
          TotalGamesPlayed = TotalGamesPlayed + 1
          PlaySound "AddX"
          'DMD "", eNone, "_", eNone, CenterLine(2, PlayersPlayingGame & " PLAYERS"), eBlink, "", eNone, 500, True, ""
          UDMD PlayersPlayingGame & " PLAYERS", "", 500
          TextDMD PlayersPlayingGame & " PLAYERS"
        Else
          If(Credits > 0)then
            PlayersPlayingGame = PlayersPlayingGame + 1
            If B2SOn Then Controller.B2SSetScorePlayer PlayersPlayingGame, 0
            TotalGamesPlayed = TotalGamesPlayed + 1
            Credits = Credits - 1
            If B2SOn then
              Controller.B2SSetScorePlayer5 Credits
            end if
            SplitCreditsForDT Credits,CreditReels
            If Credits = 0 Then DOF 140, DOFOff
            PlaySound "AddX"
            'DMD "", eNone, "_", eNone, CenterLine(2, PlayersPlayingGame & " PLAYERS"), eBlink, "", eNone, 500, True, ""
            UDMD PlayersPlayingGame & " PLAYERS", "", 500
            TextDMD PlayersPlayingGame & " PLAYERS"
          Else
            ' Not Enough Credits to start a game.
            DOF 140, DOFOff
            DMDFlush
            'DMD "", eNone, CenterLine(1, "CREDITS " & Credits), eNone, CenterLine(2, "INSERT COIN"), eNone, "", eNone, 500, True, "tna_electricity1"
            UDMD "INSERT COIN", "", 500
            TextDMD "INSERT COIN"
          End If
        End If
      End If
    End If
  Else ' If (GameInPlay) not started yet

    If keycode = StartGameKey Then
      If(bFreePlay = True)Then
        If(BallsOnPlayfield = 0)Then
          AttractModeHSDisplay.enabled=false
          AttractNessieBG.enabled=false
          If B2sOn then
            Controller.B2SSetData 99,0
            Controller.B2SSetGameOver 0
          end if
          ReelGameOver.Setvalue(0)
          ResetForNewGame()
        End If
      Else
        If(Credits > 0)Then
          If(BallsOnPlayfield = 0)Then
            Credits = Credits - 1
            If B2SOn then
              Controller.B2SSetScorePlayer5 Credits
            end if
            SplitCreditsForDT Credits,CreditReels
            If Credits = 0 Then DOF 140, DOFOff
            AttractModeHSDisplay.enabled=false
            AttractNessieBG.enabled=false
            If B2sOn then
              Controller.B2SSetData 99,0
              Controller.B2SSetGameOver 0
            end if
            ReelGameOver.Setvalue(0)
            ResetForNewGame()
          End If
        Else
          ' Not Enough Credits to start a game.
          DOF 140, DOFOff
          DMDFlush
'         DMD "", eNone, CenterLine(1, "CREDITS " & Credits), eNone, CenterLine(2, "INSERT COIN"), eBlink, "", eNone, 500, True, ""
'         UDMD "INSERT COIN", "", 500
          TextDMD "INSERT COIN"
          PlaySound "", 0, 1, -0.05, 0.05
          ShowTableInfo
        End If
      End If
    End If

    If keycode = LeftMagnaSave or  keycode = RightMagnaSave Then

    End If

    End If ' If (GameInPlay)
End Sub


Sub Table1_KeyUp(ByVal keycode)

    If keycode = PlungerKey Then
        Plunger.Fire
    PlaySoundAtVol "fx_plunger", Plunger, 1
    End If

  'Testkeys
  'targetbank2 = -3
  'ramp -20
  'spinner -12
  'to pops 14
  'target 3 = 23.5
  'target 1 = -26.5
  if keycode = 21 then  'Y key
    kickerdebug.enabled = false
    kickertest = -20
    kickerdebug.kick kickertest, 40
  end if
  if keycode = 22 then  'U key
    kickerdebug.enabled = false
    kickertest = -12
    kickerdebug.kick kickertest, 40
  end if
  if keycode = 23 then  'I key
    kickerdebug.enabled = false
    kickertest = 14
    kickerdebug.kick kickertest, 40
  end if
  if keycode = 25 then  'P key
  end if
  if keycode = 33 then  'F key
    kickerdebug.enabled = false
    kickertest = -26.5
    kickerdebug.kick kickertest, 40
  end if
  if keycode = 34 then  'G key
    kickerdebug.enabled = false
    kickertest = -3
    kickerdebug.kick kickertest, 40
  end if
  if keycode = 35 then  'H key
    kickerdebug.enabled = false
    kickertest = 23.5
    kickerdebug.kick kickertest, 40
  end if

    If hsbModeActive Then
        Exit Sub
    End If

    ' Table specific

    If (bGameInPLay AND NOT Tilted ) Then
        If keycode = LeftFlipperKey Then
            SolLFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                bInstantInfo = False
                DMDScoreNow
            End If
        End If
        If keycode = RightFlipperKey Then
            SolRFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                bInstantInfo = False
                DMDScoreNow
            End If

        End If
    End If

End Sub

Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    bInstantInfo = True
    Jackpot = 1000000 + Round(Score(CurrentPlayer) / 10, 0)
    'DMD "", eNone, CenterLine(1, "INSTANT INFO"), eNone, CenterLine(2, "JACKPOT"), eScrollLeft, CenterLine(3, FormatScore(Jackpot)), eScrollLeft, 800, False, ""
    'DMD "", eNone, CenterLine(1, "INSTANT INFO"), eNone, CenterLine(2, "SUPERJACKPOT"), eScrollLeft, CenterLine(3, FormatScore(SuperJackpot)), eScrollLeft, 800, False, ""
    'DMD "", eNone, CenterLine(1, "INSTANT INFO"), eNone, CenterLine(2, "BONUS MULTIPLIER"), eScrollLeft, CenterLine(3, BonusMultiplier(CurrentPlayer)), eScrollLeft, 800, False, ""
End Sub

Sub EndFlipperStatus
    If bInstantInfo Then
        bInstantInfo = False
        DMDScoreNow
    End If
End Sub

'*************
' Section; Pause Table
'*************
Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub table1_Exit
    Savehs
  If B2SOn Then Controller.Stop
End Sub

'********************
' Section; Flippers
'********************
Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFXDOF("fxz_flipperupL", 101, DOFOn, DOFFlippers), LeftFlipper, 1
    If FlipperPhysicsMode = 1 Then
      LeftFlipper.RotateToEnd
    Else
      LF.Fire 'LeftFlipper.RotateToEnd
    End If
        RotateLaneLightsLeft
    Else
        PlaySoundAtVol SoundFXDOF("fxz_flipperdownL", 101, DOFOff, DOFFlippers), LeftFlipper, 1
        LeftFlipper.RotateToStart
    End If
End Sub


Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFXDOF("fxz_flipperupr", 102, DOFOn, DOFFlippers), RightFlipper, 1
    If FlipperPhysicsMode = 1 Then
      RightFlipper.RotateToEnd
    Else
      RF.Fire 'RightFlipper.RotateToEnd
    End If
        RotateLaneLightsRight
    Else
        PlaySoundAtVol SoundFXDOF("fxz_flipperdownr", 102, DOFOff, DOFFlippers), RightFlipper, 1
        RightFlipper.RotateToStart
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_1", parm / 10
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_1", parm / 10
End Sub

Sub LeftFlipper2_Collide(parm)
    PlaySoundAtBallVol "flip_hit_1", parm / 10
End Sub

Sub RightFlipper2_Collide(parm)
    PlaySoundAtBallVol "flip_hit_1", parm / 10
End Sub


'*********
' Section; TILT
'*********
'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                    'Called when table is nudged
  If debugGeneral > 0 Then debug.print "CheckTilt"
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt > TiltSensitivity)AND(Tilt < 15)Then 'show a warning
        DOF 161, DOFPulse
    'DMD "", eNone, "_", eNone, CenterLine(2, "DANGER!"), eBlinkFast, "", eNone, 500, True, "tna_tilt"
    'UDMD " DANGER! ", "", 500
    End if
    If Tilt > 15 Then 'If more that 15 then TILT the table
        Tilted = True
    If B2SOn then Controller.B2SSetTilt 1
    ReelTilt.setvalue(1)
        'display Tilt
        DOF 162, DOFPulse:
    'DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "TILT!"), eBlinkFast, 500, False, "tna_tilt"
    'UDMD "  TILT!  ", "", 500
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt > 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        'turn off GI and turn off all the lights
        GiOff
        'LightSeqTilt.Play SeqAllOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart

        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'turn back on GI and the lights
        'GiOn
'        'LightSeqTilt.StopPlay
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        DMDFlush
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0)Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'********************
' Section; Music
'********************

Dim Song
Song = ""
Dim m_main

Sub StartBackgroundMusic
  Dim tmp
  If debugGeneral Then debug.print "StartBackgroundMusic"

  m_main = "lnm-theme1"
  PlaySong m_main, 1
End Sub

Sub StopBackgroundMusic
  If debugGeneral Then debug.print "StopBackgroundMusic"
  EndMusic
  StopSound "lnm-theme1"
  StopSound "lnm-theme2"
  Song = ""
End Sub

Dim MultiballMusicOn
Sub StartMultiballMusic
  If debugGeneral Then debug.print "StartMultiballMusic"
  If MultiballMusicOn = False Then
    StopBackgroundMusic
'   playsound "xxx", -1
    MultiballMusicOn = True
  End If
End Sub

Sub StopMultiballMusic
  If debugGeneral Then debug.print "StopMultiballMusic"
' stopsound "xxx"
  MultiballMusicOn = False
End Sub

Sub PlaySong(name, ttype)
  If debugGeneral Then debug.print "PlaySong " & name & " " & Song
    If bMusicOn Then
        If ((Song <> name) Or (ttype = 100)) Then
            StopSound Song
      StopBackgroundMusic
            Song = name
        If ttype = 1 Then 'Play wav file
              If Song = "m_end" Then
                PlaySound Song, 0, 0.1  'this last number is the volume, from 0 to 1
              Else
                PlaySound Song, -1, 1 'this last number is the volume, from 0 to 1
              End If
        Else    'Play mp3 file
              PlayMusic Song, SongVolume
        End If
        End If
    End If
End Sub

'**********************
' Section; GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'**********************
Const DOFuyellow= 144
Const DOFuwhite = 145
Const DOFublue  = 146
Const DOFured   = 147
Const DOFugreen = 148
Const DOFupurple= 149
Dim UndercabColor : UndercabColor = DOFublue

Dim OldGiState, CurrCol
Dim PrevCol, NextCol, NextPerm
OldGiState = -1   'start witht the Gi off

Sub ChangeGi(col, perm) 'changes the gi color, perm specifies if permanent base gi change (1) or temp change(0)
  If debugGeneral Then debug.print "*****SUB:ChangeGi " & col

  NextCol = col
  NextPerm = perm
  ChangeGiTimer.Interval = 500
  ChangeGITimer.Enabled = True

  GIOff
End Sub

Sub ChangeGiTimer_Timer
  If debugGeneral Then debug.print "*****SUB:ChangeGiTimer_Timer " & NextCol
  ChangeGITimer.Enabled = False

  ChangeGiImmediate NextCol, NextPerm
  GiOn
End Sub

Sub ChangeGiImmediate (col, perm) 'Perm=1 then save
  If debugGeneral Then debug.print "*****SUB:ChangeGiImmediate " & col
  If col = "white" Then col = "whitegi"

  If (perm = 1) Then
    CurrCol = col
  End If

    Dim bulb
    For each bulb in aGILights
        SetLight bulb, col, -1
    Next


  DOF UndercabColor, DOFOff
  If col = "white" Then
    UndercabColor = DOFuwhite
  ElseIf col = "whitegi" Then
    UndercabColor = DOFuwhite
  ElseIf col = "blue" Then
    UndercabColor = DOFublue
  ElseIf col = "red" Then
    UndercabColor = DOFured
  ElseIf col = "green" Then
    UndercabColor = DOFugreen
  ElseIf col = "purple" Then
    UndercabColor = DOFupurple
  ElseIf col = "yellow" Then
    UndercabColor = DOFuyellow
  End If
  DOF UndercabColor, DOFOn

End Sub


Sub PreviousGI
  ChangeGiImmediate CurrCol, 1
End Sub

Sub GiOn
  'If debugGeneral Then debug.print "*****SUB:GiOn"
    DOF UndercabColor, DOFOn
    Dim bulb
    For each bulb in Gi
        bulb.State = 1
    Next
  Dim xx
    For each xx in layer1col: xx.image = "layer1": Next
    For each xx in layer2col: xx.image = "layer2": Next
    For each xx in layer3col: xx.image = "layer3": Next
    For each xx in layer4col: xx.image = "layer4": Next
    For each xx in bracketscol: xx.image = "brackets": Next
    playfield_gi1.visible=1
End Sub

Sub GiOff
  'If debugGeneral Then debug.print "*****SUB:GiOff"
    DOF UndercabColor, DOFOff
    Dim bulb
    For each bulb in Gi
        bulb.State = 0
    Next
  Dim xx
    For each xx in layer1col: xx.image = "layer1off": Next
    For each xx in layer2col: xx.image = "layer2off": Next
    For each xx in layer3col: xx.image = "layer3off": Next
    For each xx in layer4col: xx.image = "layer4off": Next
    For each xx in bracketscol: xx.image = "bracketsoff": Next
    playfield_gi1.visible=0
End Sub
'
''Gi - Newer GI effects
'Dim EffectNum
'Sub GIGame (num, col)
' debug.print "GIGAme: " & num & " : " & col & "(" & CurrCol
' GIOff
' EffectNum = num
' NextCol = col
' GIGameTimer.Interval = 250
' GIGameTimer.Enabled = True
'End Sub
'
'Sub GIGameTimer_Timer
' GiGameImmediate EffectNum, NextCol
' GIGameTimer.Enabled = False
'End Sub
'
'Dim currGIGame
'Sub GiGameImmediate (num, col)  'temporary gi animations
' debug.print "GiGameImmediate: " & num & " : " & col & "(" & CurrCol
' GIOff
' currGIGame = col
' Select Case Num
''    Case 1 'Example
''      ChangeGIImmediate col, 0
''      LightSeqGame.UpdateInterval = 2
''      LightSeqGame.Play SeqDownOn, 50, 1
' End Select
'End Sub


'Sub LightSeqReady_PlayDone 'NExt step after LightSeq ends
' If CurrCol = "green" Then
'   GigameImmediate 8, "white"
' Else
'   PreviousGI
'   LightSeqReady.TimerInterval = 500
'   LightSeqReady.TimerEnabled = True
' End If
'End Sub
'
'Sub LightSeqReady_Timer
' GiOn
' LightSeqReady.TimerEnabled = False
'End Sub


' GI & light sequence effects
Dim GiNum
Sub GiEffect(n)
  Select case n
    Case 2:   'GI blink 5 times at 100 msec rate
      GiOff
      GiNum=0
      GiEffectTimer.Interval = 50
    Case 3:   'GI Blink 1 time for 100 msec
      GiOff
      GiNum = 10
      GiEffectTimer.Interval = 100
  End Select

  GiEffectTimer.Enabled = True
End Sub

Sub GiEffectTimer_Timer
  Select case GiNum
    Case 0, 2, 4, 6, 8: 'Toggle GI On
      GiOn
    Case 10   'Stop timer
      GiOn
      GiEffectTimer.Enabled = False
    Case 20,30,40,100,500,1000  'Stop timer
      GiOn
      GiEffectTimer.Enabled = False
    Case Else 'Toggle GO Off
      GiOff
  End Select
  GiNum = GiNum  + 1
End Sub



Sub LightEffect(n)

End Sub

' Flasher Effects

Dim FEStep, FEffect
FEStep = 0
FEffect = 0

Sub FlashEffect(n)
    Select case n
        Case 1:FEStep = 0:FEffect = 1:FlashEffectTimer.Enabled = 1 'all blink
        Case 2:FEStep = 0:FEffect = 2:FlashEffectTimer.Enabled = 1 'random
        Case 3:FEStep = 0:FEffect = 3:FlashEffectTimer.Enabled = 1 'upon
        Case 4:FEStep = 0:FEffect = 4:FlashEffectTimer.Enabled = 1 'ordered random :)
    End Select
End Sub

Sub FlashEffectTimer_Timer()
    FEStep = FEStep + 1
  FlashEffectTimer.Enabled = 0
End Sub


Sub ResetInserts
  Dim Obj

  For each obj in aLights
    obj.state = 0
  Next

' SetLight F1, "white", 0
' SetLight F2, "white", 0

End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' Sub PlaySoundAt(soundname, tableobj)
'   PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
' End Sub

'Set all as per ball position & speed.

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub PlaySoundAtBall(soundname)
'   PlaySoundAt soundname, ActiveBall
' End Sub

'Set position as table object and Vol manually.

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub PlaySoundAtVol(sound, tableobj, Volume)
'   PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
' End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' Sub PlaySoundAtBallVol(sound, VolMult)
'   PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

' *********************************************************************
' Section; Supporting Ball & Sound Functions
' *********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
  On Error Resume Next
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function


'*****************************************
' Section; JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 20 ' total number of balls
Const lob = 4   'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = 3 Then Exit Sub 'there are always 4 balls on this table

    ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b)) > 1 Then
            rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
      'debug.print BOT(b).velz
    End If
    Next
End Sub

'**********************
' Section; Ball Collision Sound
'**********************

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' Sub OnBallBallCollision(ball1, ball2, velocity)
'     PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
' End Sub

'******************************
' Section; Diverse Collection Hit Sounds
'******************************

Sub aTargets_Hit(idx):PlaySound "fx_Target_soft", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aBigTargets_Hit(idx):PlaySound "fx_Target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aYellowPins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aCaptiveWalls_Hit(idx):PlaySound "fx_collide", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub


'Sub PlayQuote_timer() 'one quote each 2 minutes
'    Dim Quote
'    Quote = "xxx" & INT(RND * 56) + 1
'    PlaySound Quote
'End Sub

' Ramp Sounds
Sub RHelp1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub RHelp2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
Dim RandomSongStart
Sub ResetForNewGame()
  If debugGeneral Then Debug.print "*****SUB:ResetForNewGame()"

    Dim i,j
    bGameInPLay = True

    'resets the score display, and turn off attrack mode
    StopAttractMode
  'StopRainbow
    GiOn

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
  SetPlayerUp
  Tilted=false
  DisableTable false
  'Resets all the led reels - hack
  If PlayersPlayingGame > 1 Then
    If B2SOn Then Controller.stop: Controller.Run
  End If

    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    LNMCounter(i)=0

    Next
  SplitScoreForDT 0,PlayerScore1
  SplitScoreForDT 0,PlayerScore2
  SplitScoreForDT 0,PlayerScore3
  SplitScoreForDT 0,PlayerScore4
  for i = 1 to 4
    Bank1Counter(i)=0
    Bank2Counter(i)=0
    Bank3Counter(i)=0
    for j=1 to 3
      Bank1Dropped(i,j)=0
      Bank2Dropped(i,j)=0
    next
    for j=1 to 4
      Bank3Dropped(i,j)=0
    next
  next
  If B2SOn Then Controller.B2SSetScorePlayer 1, 0
  If B2SOn Then Controller.B2SSetScorePlayer 2, 0
  If B2SOn Then Controller.B2SSetScorePlayer 3, 0
  If B2SOn Then Controller.B2SSetScorePlayer 4, 0

    ' initialise any other flags
    Tilt = 0

    ' initialise Game variables
    Game_Init()

    ' you may wish to start some music, play a sound, do whatever at this point
    ' set up the start delay to handle any Start of Game Attract Sequence
    vpmtimer.addtimer 1500, "FirstBall '"
End Sub

' This is used to delay the start of a game to allow any attract sequence to
' complete.  When it expires it creates a ball for the player to start playing with

Sub FirstBall
  If debugGeneral Then Debug.print "*****SUB:FirstBall()"

    ' reset the table for a new ball
    ResetForNewPlayerBall()
    ' create a new ball in the shooters lane
    CreateNewBall()

  playsound "FirstBall"
End Sub

' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))

Sub ResetForNewPlayerBall()
  If debugGeneral Then Debug.print "*****SUB:ResetForNewPlayerBall()"

    ' make sure the correct display is upto date
    AddScore 0

    ' set the current players bonus multiplier back down to 1X
    SetBonusMultiplier 1

    ' reset any drop targets, lights, game modes etc..

    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False

    ResetNewBallLights()

    'Reset any table specific
    ResetNewBallVariables
  ResetTopLane
  ResetMultiplier

    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True

    'and the skillshot
    SkillShotReady = 1

  'and Drain bonus Ready
  DrainBonusReady = 1
  lMystery.state=0
  lScore.state=0
  if lLock.state=2 then
    lLock.state=0
    bLockIsLit=false
  end if
  lArrow.state=0

  BumperLight001.state=0
  BumperLight002.state=0
  BumperLight003.state=0

  lOutlaneR.state=0
  lOutlaneL.state=0
  lInlaneL.state=0
  lInlaneR.state=0
  lWhite.state=0
  lBlue.state=0

  If BonusAmount(CurrentPlayer)>0 then
    BonusCount=BonusAmount(CurrentPlayer)
    SetBonusCount
  else
    BonusCount = 0
    SetBonusCount
  end if
  ResetTargetBankTimer001.enabled=true
  ResetTargetBankTimer002.enabled=true
  ResetTargetBankTimer003.enabled=true
  If B2SOn then
    Controller.B2SSetScorePlayer6 Balls
    Controller.B2SSetBallInPlay 1
    Controller.B2SSetMatch 0
  end if
  SplitBIPForDT Balls,BIPReels
'Change the music ?
End Sub

' Create a new ball on the Playfield
Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedball BallSize / 2
    ' There is a (or another) ball on the playfield
    AddBallsOnPlayfield 1

  If debugGeneral Then Debug.print "*****SUB:CreateNewBall, BallCnt = " & " : " & BallsOnPlayfield

    ' kick it out..
    PlaySound SoundFXDOF("fxz_Ballrel", 121, DOFPulse, DOFContactors), 0, 1, 0.1, 0.1, AudioFade(BallRelease)
    BallRelease.Kick 90, 4

  If bMultiBallMode = False Then
    StartBackgroundMusic
  Else

  End If
End Sub

Sub CreateNewBallAfterBallLock()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedball BallSize / 2

  If debugGeneral Then Debug.print "*****SUB:CreateNewBallAfterBallLock, BallCnt = " & " : " & BallsOnPlayfield

    ' kick it out..
    PlaySound SoundFXDOF("fx_Ballrel", 121, DOFPulse, DOFContactors), 0, 1, 0.1, 0.1, AudioFade(BallRelease)
    BallRelease.Kick 90, 4
    bAutoPlunger = True
End Sub


' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
  If debugGeneral Then Debug.print "*****SUB:AddMultiball()"

    mBalls2Eject = mBalls2Eject + nballs
  CreateMultiballTimer.Interval = 1000
    CreateMultiballTimer.Enabled = True
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
  If debugGeneral Then Debug.print "*****SUB:CreateMultiballTimer_Timer()"

    ' wait if there is a ball in the plunger lane
    If bBallInPlungerLane Then
'uuuu   debug.print "AAA"
        Exit Sub
    Else
        If BallsOnPlayfield < MaxMultiballs Then
            CreateNewBall()
            mBalls2Eject = mBalls2Eject -1
            If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
                Me.Enabled = False
            End If
        Else 'the max number of multiballs is reached, so stop the timer
            mBalls2Eject = 0
            Me.Enabled = False
        End If
    End If
End Sub

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded

Sub EndOfBall()
  Dim AwardPoints1, AwardPoints2, AwardPoints3, TotalBonus
  Dim tmp
  If debugGeneral Then Debug.print "*****SUB:EndOfBall()"

  ' bonus
    AwardPoints1 = 0
    AwardPoints2 = 0
    AwardPoints3 = 0
    TotalBonus = 0

    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False

  'Stop music
  StopBackgroundMusic

  bAutoPlunger = False

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)

    If(Tilted = False)Then

        ' Count the bonus. This table uses several bonus
        'dmdflush

    If(ExtraBallsAwards(CurrentPlayer) <> 0)Then 'Shoot again
'     DOF 163, DOFPulse: DMD "", eNone, "", eNone, "", eNone, CenterLine(3, ("STAND BY...")), eBlink, 2000, True, "tna_standbyforextraball"
'     UDMD "STAND BY...", "", 2000
    Else ' Normal Bouss MOde

    End If

    BonusCountDown

    ' add a bit of a delay to allow for the bonus points to be shown & added up
      'vpmtimer.addtimer 9000, "EndOfBall2 '"
    'EndofBall2 is called when Bonus CountDown complete

  Else
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub


Sub EndOfBall2()
  If debugGeneral Then Debug.print "*****SUB:EndOfBall2()"

  ChangeGi GIcolor, 1

    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilted = False
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots
  If B2SOn then Controller.B2SSetTilt 0
  ReelTilt.setvalue(0)
    ' has the player won an extra-ball ? (might be multiple outstanding)
    If(ExtraBallsAwards(CurrentPlayer) <> 0)Then
        debug.print "Extra Ball"

        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            lShootAgain.State = 0
      ReelShootAgain.setvalue(0)
      If B2SOn then
        Controller.B2SSetShootAgain 0
      end if
        End If

        ' You may wish to do a bit of a song AND dance at this point
        'DMD "", eNone, "", eNone, "", eNone, CenterLine(3, ("SHOOT AGAIN")), eBlink, 1000, True, "tna_shootagain"
    UDMD "SHOOT AGAIN", "", 1000

        ' Create a new ball in the shooters lane
    ResetForNewPlayerBall
        CreateNewBall()


    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0)Then
            debug.print "No More Balls, High Score Entry"

            ' Submit the CurrentPlayers score to the High Score system
            CheckHighScore()
        ' you may wish to play some music at this point

        Else

            ' not the last ball (for that player)
            ' if multiple players are playing then move onto the next one
            EndOfBallComplete()
        End If
    End If
End Sub

' This function is called when the end of bonus display
' (or high score entry finished) AND it either end the game or
' move onto the next player (or the next ball of the same player)
'
Sub EndOfBallComplete()
    Dim NextPlayer

    If debugGeneral Then debug.print "*****SUB: EndOfBallComplete()"

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame > 1)Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer > PlayersPlayingGame)Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

        ' set the machine into game over mode
        EndOfGame()

    ' you may wish to put a Game Over message on the desktop/backglass

    Else
    ' Save any additional Player data
    SavePlayerData

    ' Extra Save step for co-op mode.  Copies data to other players
    If CoopMode = 1 Then
      'Copy player data to all Players
      CopyPlayerData CurrentPlayer, 1
      CopyPlayerData CurrentPlayer, 2
      CopyPlayerData CurrentPlayer, 3
      CopyPlayerData CurrentPlayer, 4
    ElseIf CoopMode = 2 Then
      'Copy score to alternate Players
      Select Case CurrentPlayer
        Case 1
          CopyPlayerData CurrentPlayer, 3
        Case 2
          CopyPlayerData CurrentPlayer, 4
        Case 3
          CopyPlayerData CurrentPlayer, 1
        Case 4
          CopyPlayerData CurrentPlayer, 2
      End Select
    End If

        ' set the next player
        CurrentPlayer = NextPlayer
    SetPlayerUp
    ' Restore next player data or load default setting for ball 1
    RestorePlayerData

        ' make sure the correct display is up to date
        AddScore 0

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()

        ' play a sound if more than 1 player
        If PlayersPlayingGame > 1 Then
            'PlaySound "vo_player" &CurrentPlayer
            'DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "PLAYER " &CurrentPlayer), eNone, 800, True, ""
      UDMD "PLAYER " & CurrentPlayer, "", 800
      TextDMD "PLAYER " & CurrentPlayer
        End If
    End If

    If debugGeneral Then debug.print "Next Player = " & NextPlayer

End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    If debugGeneral Then debug.print "*****SUB:EndOfGame()"

    bGameInPLay = False
  BallSearchTimer.Enabled = False
  CoopMode = 0
    ' just ended your game then play the end of game tune
    If NOT bJustStarted Then
        'PlaySong "m_end", 1
    End If
    bJustStarted = False
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all modes - eject locked balls
    ' most of the modes/timers terminate at the end of the ball
    'PlayQuote.Enabled = 0

    ' set any lights for the attract mode
    GiOff
' StartRainbow "all"
'   StartAttractMode 1

  ' you may wish to light any Game Over Light you may have
  'Release any balls left on playfield at end of game
  'SetBallsOnPlayfield x
  If BallIsLocked=true then
    BallIsLocked=false
    bLockIsLit = False
    lBallRelease.state=0
    BallLockedBlocker.IsDropped=1
    lLock.state=0
    AddBallsOnPlayfield 1
    Tilted=true
    DisableTable True
    vpmTimer.addTimer 1500, "BallLockEject '"
  end if
  AttractModeHSDisplay.enabled=true
  AttractNessieBG.enabled=true
  If B2SOn then
    Controller.B2SSetScorePlayer6 0
    Controller.B2SSetGameOver 1
    Controller.B2SSetBallInPlay 0
    Controller.B2SSetPlayerUp 0

  end if
  PlayerUpLights(0).state=0
  PlayerUpLights(1).state=0
  PlayerUpLights(2).state=0
  PlayerUpLights(3).state=0
  ReelGameOver.Setvalue(1)
  SplitBIPForDT 0,BIPReels
  CheckMatch
End Sub

Sub CheckMatch
  Dim tempmatch,Match,i

  tempmatch=Int(Rnd*10)
  Match=tempmatch*10
  'MatchReel.SetValue(tempmatch+1)

  If B2SOn Then
    Controller.B2SSetMatch 1
    If Match = 0 Then
      Controller.B2SSetScorePlayer6 100
    Else
      Controller.B2SSetScorePlayer6 Match
    End If
  End if
  SplitBIPForDT Match,BIPReels
  for i = 1 to PlayersPlayingGame
    if Match=(Score(i) mod 100) then
      AwardSpecial
    end if
  next
end sub


Dim BallinPlay
Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
    BallinPlay = BallsPerGame
    Else
        Balls = tmp
    BallinPlay = tmp
    End If
End Function

' *********************************************************************
'  Section; Drain / Plunger Functions
' *********************************************************************

' lost a ball ;-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one
'
Sub Drain_Hit()
  DOF 116, DOFPulse
    startB2S(7)
    ' Destroy the ball
    Drain.DestroyBall
    ' Exit Sub ' only for debugging
    AddBallsOnPlayfield -1

  If debugGeneral Then Debug.print "Drain_Hit(), ballcnt=" & BallsOnPlayfield

    ' pretend to knock the ball into the ball storage mech
    PlaySoundAtVol "fxz_drain", Drain, 1

    'if Tilted the end Ball modes
    If Tilted Then
        StopEndOfBallModes
    End If

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True)AND(Tilted = False)Then

        ' is the ball saver active,
        If(bBallSaverActive = True) Then

            ' yep, create a new ball in the shooters lane
            ' we use the Addmultiball in case the multiballs are being ejected
            AddMultiball 1
            ' we kick the ball with the autoplunger
            bAutoPlunger = True
            ' you may wish to put something on a display or play a sound at this point
      if (bMultiBallMode = False) Then
'       DOF 165, DOFPulse: DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "BALL SAVED"), eBlinkfast, 800, True, "tna_ballsaved"
'       UDMD "BALL SAVED", "", 800
      End If

        Else


            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1)Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True)then
                    ' not in multiball mode any more
                    bMultiBallMode = False
          lWhenFlashing.state=0
          EndMultiball

                    ' you may wish to change any music over at this point and
                    ' turn off any multiball specific lights
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0)Then
                ' End Modes and timers
                'PlaySong "m_wait", 1
                StopEndOfBallModes
        BumperFlashingTimer.enabled=false
        BumperLight001.state=0
        BumperLight002.state=0
        BumperLight003.state=0
               ChangeGi "red", 0
                ' handle the end of ball (count bonus, change player, high score entry etc..)
        If DrainBonusReady = 1 Then
          DrainBonusReady = 0
          vpmtimer.addtimer 1500, "EndOfBall '"
        End If
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Dim BallRestingInPlungerLane
Const AutoPlungeDelay = 1
Sub swPlungerRest_Hit()
    If debugGeneral Then debug.print "*****SUB:swPlungerRest_Hit, Resting"

  BallRest.TimerInterval = AutoPlungeDelay*1000

    ' some sound according to the ball position
    PlaySound "fx_sensor", 0, 1, 0.15, 0.25
    bBallInPlungerLane = True

  If SkillshotReady = 2 Then  'Soft plunge failed.  Cancel Skillshot and Autolaunch
    bAutoPlunger = True
    CheckSkillShot 0
    PlaySound "tna_electricity1", 0, 1, -0.05, 0.05
    BallRest.TimerInterval = 500
    BallRest.TimerEnabled = False
    BallRest.TimerEnabled = True
  End If

    ' turn on Launch light is there is one
    'LaunchLight.State = 2
    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
    If debugGeneral Then debug.print "*****SUB:swPlungerRest_Hit, AutoPlunge"
    BallRest.TimerEnabled = False
    BallRest.TimerEnabled = True


        If mBalls2Eject = 0 Then bAutoPlunger = False

  End If


    'Start the Selection of the skillshot if ready
    If SkillShotReady = 1 Then
        StartSkillshot()
    End If

    ' remember last trigger hit by the ball.
    SetLastSwitchHit "swPlungerRest"

  uDMDScoreUpdate

End Sub

Sub BallRest_Timer
    If debugGeneral Then debug.print "*****SUB:BallRest_Timer, AutoPlunge"

        PlungerIM.AutoFire
    PlaySound SoundFXDOF("fxz_autoplunger",125,DOFPulse,DOFContactors)
        DOF 123, DOFPulse
    BallRest.TimerEnabled = False
    'LightSeqAutoLaunch.StopPlay
'   LightSeqAutoLaunch.Play SeqAllOff
'   LightSeqAutoLaunch.UpdateInterval = 4
'   LightSeqAutoLaunch.Play SeqUpOn, 25, 1

End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Dim bTimedSkillShot
Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
    If SkillShotReady  = 1 Then
        SkillShotReady = 2  'Ball plunged
        bSkillShotSelect = False  'Skillshot frozen
    swPlungerRest.TimerInterval = 5000
    swPlungerRest.TimerEnabled = True
    bTimedSkillShot = True
    End If

    If debugGeneral Then debug.print "*****SUB:swPlungerRest_UnHit, Ballcnt: " & BallsOnPlayfield

    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
  'msgbox bBallSaverReady & ":" & BallSaverTime & ":" & bBallSaverActive
    If(bBallSaverReady = True)AND(BallSaverTime <> 0)And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    End If


End Sub

Sub swPlungerRest_Timer 'Timed skillshot support
  bTimedSkillShot = False
  swPlungerRest.TimerEnabled = False
End Sub


'Version of Ballsaver using led display
Dim dbs1, dbs2, dbsdelta, dbstime, dbstens, dbsones, dbsdecimals
Sub EnableBallSaver(seconds)
  seconds = seconds + 0.3  'padding
    If debugGeneral Then debug.print "*****SUB:EnableBallSaver, seconds=" & seconds
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer

  BallSaverTimerExpired.Interval = 1000 * seconds
  BallSaverTimerExpired.Enabled = True

  lShootAgain.State = 2

End Sub

Sub StopBallSaver
  BallSaverTimer2Expired.Enabled = False
  If ExtraBallsAwards(CurrentPlayer) = 0 Then
    lShootAgain.State = 0
  Else
    lShootAgain.State = 1
  End If
  bBallSaverActive = False
End Sub


' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimerExpired_Timer()
    If debugGeneral Then debug.print "*****SUB:" & "BallSaverTimerExpired_Timer"
    BallSaverTimerExpired.Enabled = False

  If ExtraBallsAwards(CurrentPlayer) = 0 Then
    lShootAgain.State = 0
  Else
    lShootAgain.State = 1
  End If

    ' give extra 2 second
  BallSaverTimer2Expired.Interval = 2000
  BallSaverTimer2Expired.Enabled = True
End Sub


Sub BallSaverTimer2Expired_Timer()
    If debugGeneral Then debug.print "*****SUB:" & "BallSaverTimer2Expired_Timer"
    BallSaverTimer2Expired.Enabled = False

    ' clear the flag
    bBallSaverActive = False
End Sub


'Version of Ballsaver using light insert
'Sub EnableBallSaver(seconds)
'    If debugGeneral Then debug.print "*****SUB:EnableBallSaver, seconds=" & seconds
'    ' set our game flag
'    bBallSaverActive = True
'    bBallSaverReady = False
'    ' start the timer
'    BallSaverTimerExpired.Interval = 1000 * seconds
'    BallSaverTimerExpired.Enabled = True
'    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
'    BallSaverSpeedUpTimer.Enabled = True
'    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
'    lLightShootAgain.BlinkInterval = 160
'    lLightShootAgain.State = 2
'End Sub
'
'' The ball saver timer has expired.  Turn it off AND reset the game flag
''
'Sub BallSaverTimerExpired_Timer()
'       If debugGeneral Then debug.print "*****SUB:" & "BallSaverTimerExpired_Timer"
'    BallSaverTimerExpired.Enabled = False
'    ' clear the flag
'    bBallSaverActive = False
'    ' if you have a ball saver light then turn it off at this point
'    lLightShootAgain.State = 0
'End Sub
'
'Sub BallSaverSpeedUpTimer_Timer()
'    If debugGeneral Then debug.print "*****SUB:" & "BallSaverSpeedUpTimer_Timer"
'    BallSaverSpeedUpTimer.Enabled = False
'    ' Speed up the blinking
'    lLightShootAgain.BlinkInterval = 80
'    lLightShootAgain.State = 2
'End Sub
' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

Sub AddScore(points)
  Dim xmultiplier
  'Score multiplied during multiball
  xMultiplier = BallsOnPlayfield
  If xMultiplier = 0 then xMultiplier = 1

    If(Tilted = False)Then
        ' add the points to the current players score variable
        Score(CurrentPlayer) = Score(CurrentPlayer) + (points * xMultiplier)

        ' update the score displays
        DMDScore
    End If

   If debugGeneral Then debug.print "*****SUB:" & "AddScore(" & points * xMultiplier & "), Total=" & Score(CurrentPlayer)
  ' you may wish to check to see if the player has gotten a replay
End Sub


Sub AddScoreSpecial(points) 'Increase Score and do extra stuff
    If debugGeneral Then debug.print "*****SUB:" & "AddScoreSpecial(" & points & ")"

  AddScore (points)

  'increase bonus

End Sub

Sub AddScoreAndBonus (points)
  ' add the bonus to the current players bonus variable
    BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + 1
  Addscore (points)
End Sub


'Sub AddBonus(points)
'    If debugGeneral Then debug.print "*****SUB:" & "AddBonus(" & points & ")"
'
'    If(Tilted = False)Then
'        ' add the bonus to the current players bonus variable
'        BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
'        ' update the score displays
'        DMDScore
'    End if
'
'' you may wish to check to see if the player has gotten a replay
'End Sub
'
'' Add some points to the current Jackpot.
''
'Sub AddJackpot(points) 'not used in this table
'End Sub

Sub SetExtraBallDisplay
  If B2SOn then
    Controller.B2SSetShootAgain 1
  end if
  lShootAgain.state=1
  ReelShootAgain.setvalue(1)
end sub



Sub AwardExtraBall()
  If CoopMode = 0 Then
    If NOT bExtraBallWonThisBall Then 'Use if you want to limit to 1 extraball
      DOF 167, DOFPulse:
      'DMD "", eNone, "_", eNone, Centerline(2, "EXTRA BALL WON"), eBlink, "", eNone, 1000, True, "tna_extraball"
      UDMD "EXTRA BALL", "AWARDED", 1000
      ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
      debug.print "ExtraBall count " & ExtraBallsAwards(CurrentPlayer)
'     bExtraBallWonThisBall = True

      'Set Insert or Display
      SetExtraBallDisplay
    End If
  End If
End Sub

Sub AwardExtraBallNoCallout()
  If CoopMode = 0 Then
    If NOT bExtraBallWonThisBall Then 'Use if you want to limit to 1 extraball
      ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
      debug.print "ExtraBall count " & ExtraBallsAwards(CurrentPlayer)
'         bExtraBallWonThisBall = True

      'Set Insert or Display
      SetExtraBallDisplay
    End If
  End If
End Sub


Sub AwardSpecial()
    DMDFlush
'    DMD "", eNone, "_", eNone, Centerline(2, "EXTRA GAME WON"), eBlink, "", eNone, 1000, True, SoundFXDOF("fx_knocker", 129, DOFPulse, DOFKnocker)
    UDMD "EXTRA GAME", "AWARDED", 1000
  Credits = Credits + 1
  If B2SOn then
    Controller.B2SSetScorePlayer5 Credits
  end if
  SplitCreditsForDT Credits,CreditReels
  DOF 123, DOFPulse
    DOF 140, DOFOn
  Playsound "knocker"
End Sub

Sub AwardJackpot() 'award a normal jackpot, double or triple jackpot
  If debugGeneral Then Debug.print "*****SUB:AwardJackpot()"
    DMDFlush
    DOF 123, DOFPulse
' DMD "", eNone, Centerline(1, ("JACKPOT")), eNone, "", eNone, CenterLine(3, FormatScore(Jackpot)), eBlinkFast, 1000, True, "tna_jackpot"
    UDMD "JACKPOT", Jackpot, 1000
  AddScore Jackpot
  GIGame 1, "purple"

End Sub

Sub AwardDoubleJackpot() 'in this table the jackpot is always 1 million + 10% of your score
  If debugGeneral Then Debug.print "*****SUB:AwardDoubleJackpot()"
    DMDFlush
    DOF 123, DOFPulse
' DMD "", eNone, Centerline(1, ("DOUBLE JACKPOT")), eNone, "", eNone, CenterLine(3, FormatScore(DoubleJackpot)), eBlinkFast, 1000, True, "tna_doublejackpot"
    UDMD "DOUBLE JACKPOT", DoubleJackpot, 1000
    AddScore DoubleJackpot
  GIGame 9, "purple"
End Sub

Sub AwardTripleJackpot() 'in this table the jackpot is always 1 million + 10% of your score
  If debugGeneral Then Debug.print "*****SUB:AwardTripleJackpot()"
    DOF 132, DOFPulse
    DMDFlush
'   DMD "", eNone, Centerline(1, ("TRIPLE JACKPOT")), eNone, "", eNone, CenterLine(3, FormatScore(TripleJackpot)), eBlinkFast, 1000, True, "tna_triplejackpot"
    UDMD "TRIPLE JACKPOT", TripleJackpot, 1000
    AddScore TripleJackpot
  GIGame 10, "purple"
End Sub

Sub AwardSuperJackpot()
  If debugGeneral Then Debug.print "*****SUB:AwardSuperJackpot()"
    DOF 133, DOFPulse
    DMDFlush
'    DMD "", eNone, Centerline(1, ("SUPER JACKPOT")), eNone, "", eNone, CenterLine(3, FormatScore(SuperJackpot)), eBlinkFast, 1000, True, "tna_superjackpot"
    UDMD "SUPER JACKPOT", SuperJackpot, 1000
    AddScore SuperJackpot
  GIGame 11, "purple"
End Sub


Sub AwardSkillshot()
  If debugGeneral Then Debug.print "*****SUB:AwardSkillshot()"
    DMDFlush
  DOF 168, DOFPulse:
  DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "SKILLSHOT"), eBlinkfast, 1000, True, "tna_lanesaveincreased"
    AddScore SkillshotValue
    GiEffect 2

End Sub





'*****************************
' Section;    Load / Save / Highscore
'*****************************

Sub Reseths
  If debugHighScore Then debug.print "Sub:Reseths"
    Dim x
  x = ""
    If(x <> "")Then HighScore(1) = CDbl(x)Else HighScore(1) = 100000 End If

    If(x <> "")then Credits = CInt(x)Else Credits = 0 End If
'    If(x <> "")then TotalGamesPlayed = CInt(x)Else TotalGamesPlayed = 0 End If
  Savehs
End Sub

Sub Loadhs
  If debugHighScore Then debug.print "Sub:Loadhs"
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "")Then HighScore(1) = CDbl(x) Else HighScore(1) = 100000 End If

    x = LoadValue(TableName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0 End If
    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "")then TotalGamesPlayed = CInt(x)Else TotalGamesPlayed = 0 End If
End Sub

Sub Savehs
  If debugHighScore Then debug.print "Sub:Savehs"

    SaveValue TableName, "HighScore1", HighScore(1)
    SaveValue TableName, "Credits", Credits
    SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
End Sub

' ***********************************************************
'  High Score Initals Entry Functions - based on Black's code
' ***********************************************************

Dim hsbModeActive
Dim hsEnteredName
Dim hsEnteredDigits(3)
Dim hsCurrentDigit
Dim hsValidLetters
Dim hsCurrentLetter
Dim hsLetterFlash

Sub CheckHighscore()
  If debugHighScore Then debug.print "Sub:CheckHighscore"

  If CoopMode = 0 Then
    Dim tmp
'   tmp = Score(1)
'   If Score(2) > tmp Then tmp = Score(2)
'   If Score(3) > tmp Then tmp = Score(3)
'   If Score(4) > tmp Then tmp = Score(4)

    tmp = Score(CurrentPlayer)

    If tmp > HighScore(1)Then 'add 1 credit for beating the highscore
      HighScore(1)=tmp
      AwardSpecial
    End If

    EndOfBallComplete()
  Else 'Co op mode running so no high score allowed
    EndOfBallComplete()
  End If
End Sub

Sub HighScoreEntryInit()
  exit sub
End Sub

Sub EnterHighScoreKey(keycode)
  If debugHighScore Then debug.print "Sub:EnterHighScoreKey"

    If keycode = LeftFlipperKey Then
        Playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0)then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        Playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter > len(hsValidLetters))then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "`")then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3)then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit > 0)then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayNameNow()
        end if
    end if
End Sub

Sub HighScoreDisplayNameNow()
  exit sub
End Sub

Sub HighScoreDisplayName()
  exit sub
End Sub

Sub HighScoreFlashTimer_Timer()
  If debugHighScore Then debug.print "Sub:HighScoreFlashTimer_Timer"
    HighScoreFlashTimer.Enabled = False
End Sub

Sub HighScoreCommitName()
  If debugHighScore Then debug.print "Sub:HighScoreCommitName"
    HighScoreFlashTimer.Enabled = False
    'hsbModeActive = False

  DMD "", eNone, "", eNone, "", eNone, "", eNone, 1000, True, ""
  vpmtimer.addtimer 800, "HighscoreDelay '"

    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ")then
        hsEnteredName = "YOU"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
    EndOfBallComplete()
End Sub

Sub HighscoreDelay
  hsbModeActive = False
End Sub

Sub SortHighscore

End Sub

' *****************************************************************************
'   JP's Reduced Display Driver Functions for Slimer (based on script by Black)
' only 5 effects: none, scroll left, scroll right, blink and blinkfast
' 4 Lines, treats all 4 lines as text
' Example format:
' DMD "backgnd", eNone,"text1", eNone,"text2", eNone, "centertext", eNone, 250, True, "sound"
' Short names:
' dq = display queue
' de = display effect
' "_" in a line means: do nothing
' *****************************************************************************

Const eNone = 0        ' Instantly displayed
Const eScrollLeft = 1  ' scroll on from the right
Const eScrollRight = 2 ' scroll on from the left
Const eBlink = 3       ' Blink (blinks for 'TimeOn')
Const eBlinkFast = 4   ' Blink (blinks for 'TimeOn') at user specified intervals (fast speed)
Const dqSize = 64

Dim dqHead
Dim dqTail
Dim deSpeed
Dim deBlinkSlowRate
Dim deBlinkFastRate

Dim dCharsPerLine(3)
Dim dLine(3)
Dim deCount(3)
Dim deCountEnd(3)
Dim deBlinkCycle(3)

Dim dqText(3, 64)
Dim dqEffect(3, 64)
Dim dqTimeOn(64)
Dim dqbFlush(64)
Dim dqSound(64)

Sub DMD_Init() 'default/startup values
    Dim i, j
    DMDFlush()
    deSpeed = 20
    deBlinkSlowRate = 5
    deBlinkFastRate = 2
    dCharsPerLine(0) = 3
    dCharsPerLine(1) = 19
    dCharsPerLine(2) = 19
    dCharsPerLine(3) = 13
    For i = 0 to 3
        dLine(i) = Space(dCharsPerLine(i))
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
        dqTimeOn(i) = 0
        dqbFlush(i) = True
        dqSound(i) = ""
    Next
    For i = 0 to 3
        For j = 0 to 64
            dqText(i, j) = ""
            dqEffect(i, j) = eNone
        Next
    Next
    DMD "", eNone, "", eNone, "", eNone, "", eNone, 25, True, ""
End Sub

Sub DMDFlush()
  If (UseVPReelDMD > 0) Then
    Dim i
    DMDTimer.Enabled = False
    DMDEffectTimer.Enabled = False
    dqHead = 0
    dqTail = 0
    For i = 0 to 3
      deCount(i) = 0
      deCountEnd(i) = 0
      deBlinkCycle(i) = 0
    Next
  End If

  If UseUltraDMD > 0 Then UltraDMD.CancelRendering

End Sub

Sub DMDScoreNow()
    DMDFlush()
'    DMDScore()
End Sub

Sub DMDScore()
    Dim tmp0, tmp1, tmp2, tmp3,tempdump

  If UseVPReelDMD > 0 Then
    if(dqHead = dqTail)Then
      tmp0 = ""
      tmp1 = FillLine(1, " PLAYER " & CurrentPlayer, FormatScore(Score(CurrentPlayer)))
      tmp2 = FillLine(2, " RVAL:" & FormatScore(ReactorValue(CurrentPlayer)), "BALL" & Balls)
      tmp3 = ""
      DMD tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 25, True, ""
    End If
  End If
  tempdump=Score(CurrentPlayer)
  Select case CurrentPlayer
    case 1:
      SplitScoreForDT tempdump,PlayerScore1
    case 2:
      SplitScoreForDT tempdump,PlayerScore2
    case 3:
      SplitScoreForDT tempdump,PlayerScore3
    case 4:
      SplitScoreForDT tempdump,PlayerScore4
  end select

  If B2SOn Then
    Controller.B2SSetScorePlayer CurrentPlayer, (Score(CurrentPlayer))
  End If

  uDMDScoreUpdate

  If UseVPTextDMD > 0 Then
    ScoreText.Text = "B:" & Balls & " C: " & Credits & "   " & (Score(CurrentPlayer))
  End If
End Sub

Sub TextDMD (val)
  If UseVPTextDMD > 0 Then
    ScoreText.Text = val
  End If
End Sub

Sub DMD(Text0, Effect0, Text1, Effect1, Text2, Effect2, Text3, Effect3, TimeOn, bFlush, Sound) 'the lines are  background. top line, bottom line, and centerline
  If (UseVPReelDMD > 0) Then
    if(dqTail < dqSize)Then
      if(Text0 = "_")Then
        dqEffect(0, dqTail) = eNone
        dqText(0, dqTail) = "_"
      Else
        dqEffect(0, dqTail) = Effect0
        dqText(0, dqTail) = ExpandLine(Text0, 0)
      End If

      if(Text1 = "_")Then
        dqEffect(1, dqTail) = eNone
        dqText(1, dqTail) = "_"
      Else
        dqEffect(1, dqTail) = Effect1
        dqText(1, dqTail) = ExpandLine(Text1, 1)
      End If

      if(Text2 = "_")Then
        dqEffect(2, dqTail) = eNone
        dqText(2, dqTail) = "_"
      Else
        dqEffect(2, dqTail) = Effect2
        dqText(2, dqTail) = ExpandLine(Text2, 2)
      End If

      if(Text3 = "_")Then
        dqEffect(3, dqTail) = eNone
        dqText(3, dqTail) = "_"
      Else
        dqEffect(3, dqTail) = Effect3
        dqText(3, dqTail) = ExpandLine(Text3, 3)
      End If

      dqTimeOn(dqTail) = TimeOn
      dqbFlush(dqTail) = bFlush
      dqSound(dqTail) = Sound
      dqTail = dqTail + 1
      if(dqTail = 1)Then
        DMDHead()
      End If
    End If
  End If
End Sub

Sub DMDHead()
    Dim i
    deCount(0) = 0
    deCount(1) = 0
    deCount(2) = 0
    deCount(3) = 0
    DMDEffectTimer.Interval = deSpeed

    For i = 0 to 3
        Select Case dqEffect(i, dqHead)
            Case eNone:deCountEnd(i) = 1
            Case eScrollLeft:deCountEnd(i) = Len(dqText(i, dqHead))
            Case eScrollRight:deCountEnd(i) = Len(dqText(i, dqHead))
            Case eBlink:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
            Case eBlinkFast:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
        End Select
    Next



    if(dqSound(dqHead) <> "")Then
        PlaySound(dqSound(dqHead))
    End If
    DMDEffectTimer.Enabled = True
End Sub

Sub DMDEffectTimer_Timer()
    DMDEffectTimer.Enabled = False
    DMDProcessEffectOn()
End Sub

Sub DMDTimer_Timer()
    Dim Head
    DMDTimer.Enabled = False
    Head = dqHead
    dqHead = dqHead + 1
    if(dqHead = dqTail)Then
        if(dqbFlush(Head) = True)Then
            DMDFlush()
            DMDScore()
        Else
            dqHead = 0
            DMDHead()
        End If
    Else
        DMDHead()
    End If
End Sub

Sub DMDProcessEffectOn()
    Dim i
    Dim BlinkEffect
    Dim Temp

    BlinkEffect = False

    For i = 0 to 3
        if(deCount(i) <> deCountEnd(i))Then
            deCount(i) = deCount(i) + 1

            select case(dqEffect(i, dqHead))
                case eNone:
                    Temp = dqText(i, dqHead)
                case eScrollLeft:
                    Temp = Right(dLine(i), dCharsPerLine(i)- 1)
                    Temp = Temp & Mid(dqText(i, dqHead), deCount(i), 1)
                case eScrollRight:
                    Temp = Mid(dqText(i, dqHead), (dCharsPerLine(i) + 1)- deCount(i), 1)
                    Temp = Temp & Left(dLine(i), dCharsPerLine(i)- 1)
                case eBlink:
                    BlinkEffect = True
                    if((deCount(i)MOD deBlinkSlowRate) = 0)Then
                        deBlinkCycle(i) = deBlinkCycle(i)xor 1
                    End If

                    if(deBlinkCycle(i) = 0)Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i))
                    End If
                case eBlinkFast:
                    BlinkEffect = True
                    if((deCount(i)MOD deBlinkFastRate) = 0)Then
                        deBlinkCycle(i) = deBlinkCycle(i)xor 1
                    End If

                    if(deBlinkCycle(i) = 0)Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i))
                    End If
            End Select

            if(dqText(i, dqHead) <> "_")Then
                dLine(i) = Temp
                DMDUpdate i
            End If
        End If
    Next

    if(deCount(0) = deCountEnd(0))and(deCount(1) = deCountEnd(1))and(deCount(2) = deCountEnd(2))and(deCount(3) = deCountEnd(3))Then

        if(dqTimeOn(dqHead) = 0)Then
            DMDFlush()
        Else
            if(BlinkEffect = True)Then
                DMDTimer.Interval = 10
            Else
                DMDTimer.Interval = dqTimeOn(dqHead)
            End If

            DMDTimer.Enabled = True
        End If
    Else
        DMDEffectTimer.Enabled = True
    End If
End Sub

Function ExpandLine(TempStr, id) 'id is the number of the dmd line
    If TempStr = "" Then
        TempStr = Space(dCharsPerLine(id))
    Else
        if(Len(TempStr) > Space(dCharsPerLine(id)))Then
            TempStr = Left(TempStr, Space(dCharsPerLine(id)))
        Else
            if(Len(TempStr) < dCharsPerLine(id))Then
                TempStr = TempStr & Space(dCharsPerLine(id)- Len(TempStr))
            End If
        End If
    End If
    ExpandLine = TempStr
End Function

Function FormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
    dim i
    dim NumString

    NumString = CStr(abs(Num))

    For i = Len(NumString)-3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1))then
            NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1)) + 48) & right(NumString, Len(NumString)- i)
        end if
    Next
    FormatScore = NumString
End function

Function UDMDFormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
    dim i
    dim NumString

    NumString = CStr(abs(Num))

    For i = Len(NumString)-3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1))then
            NumString = left(NumString, i-1) & "," & right(NumString, Len(NumString)- i)
        end if
    Next
    UDMDFormatScore = NumString
End function

Function CenterLine(id, aString)
    Dim tmp, tmpStr
    tmp = (dCharsPerLine(id)- Len(aString)) \ 2
    If(tmp + tmp + Len(aString)) < dCharsPerLine(id)Then
        tmpStr = " " & Space(tmp) & aString & Space(tmp)
    Else
        tmpStr = Space(tmp) & aString & Space(tmp)
    End If
    CenterLine = tmpStr
End Function

Function FillLine(id, aString, bString)
    Dim tmp, tmpStr
    tmp = dCharsPerLine(id)- Len(aString)- Len(bString)
    tmpStr = aString & Space(tmp) & bString
    FillLine = tmpStr
End Function

Function RightLine(id, aString)
    Dim tmp, tmpStr
    tmp = dCharsPerLine(id)- Len(aString)
    tmpStr = Space(tmp) & aString
    RightLine = tmpStr
End Function

'*********************
' Section; Update DMD - reels
'*********************
Dim DesktopMode:DesktopMode = Table1.ShowDT

Dim Digits(3)

If UseVPReelDMD Then DMDReels_Init

Sub DMDReels_Init
    If DesktopMode Then
        'Desktop
        Digits(0) = Array(d0, d1, d2)                                                                                    'backdrop
        Digits(1) = Array(d3, d4, d5, d6, d7, d8, d9, d10, d11, d12, d13, d14, d15, d16, d17, d18, d19, d20, d21)        'upper line
        Digits(2) = Array(d22, d23, d24, d25, d26, d27, d28, d29, d30, d31, d32, d33, d34, d35, d36, d37, d38, d39, d40) 'lower line
        Digits(3) = Array(d41, d42, d43, d44, d45, d46, d47, d48, d49, d50, d51, d52, d53)                               ' center line
        d54.Visible = 0:d55.Visible = 0:d56.Visible = 0
    Else
        'FS
        Digits(0) = Array(d54, d55, d56)                                                                                 'backdrop
        Digits(1) = Array(d57, d58, d59, d60, d61, d62, d63, d64, d65, d66, d67, d68, d69, d70, d71, d72, d73, d74, d75) 'upper line
        Digits(2) = Array(d76, d77, d78, d79, d80, d81, d82, d83, d84, d85, d86, d87, d88, d89, d90, d91, d92, d93, d94) 'lower line
        Digits(3) = Array(d95, d96, d97, d98, d99, d100, d101, d102, d103, d104, d105, d106, d107)
        d0.Visible = 0:d1.Visible = 0:d2.Visible = 0
    End If
End Sub

Sub DMDUpdate(id)
    Dim digit, value

  If UseApronDMD = 1 or DesktopMode or hsbModeActive or ResetHighScore = 1 Then
    For digit = 0 to dCharsPerLine(id)-1
      value = ASC(mid(dLine(id), digit + 1, 1))-32
      Digits(id)(digit).SetValue value
    Next
  End If

End Sub

'****************************************
' Section; Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub GameTimer_Timer
    RollingUpdate
    ' add any other real time update subs, like gates or diverters
End Sub

'********************************************************************************************
' Only for VPX 10.2 and higher.
' Section; FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first myVersion

    If TypeName(MyLight) = "Light" Then

        If FinalState = 2 Then
            FinalState = MyLight.State 'Keep the current light state
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then

        Dim steps

        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
        If FinalState = 2 Then                      'Keep the current flasher state
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state

        ' Start blink timer and create timer subroutine
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub

Sub FlashForMsrgb(MyLight, TotalPeriod, BlinkPeriod, FinalState, Red, Green, Blue) 'thanks gtxjoe for the first myVersion

    If TypeName(MyLight) = "Light" Then

        If FinalState = 2 Then
            FinalState = MyLight.State 'Keep the current light state
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then

        Dim steps

        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
        If FinalState = 2 Then                      'Keep the current flasher state
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state

        ' Start blink timer and create timer subroutine
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub
'******************************************
' Section; Change light color - simulate color leds
' changes the light color and state
' colors: red, orange, yellow, green, blue, white
'******************************************

'Sub SetLight(n, col, stat)
Sub SetLight(n, col, stat)
  'SEt Color
    Select Case col
        Case "red"
            n.color = RGB(128, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case "orange"
            n.color = RGB(18, 3, 0)
            n.colorfull = RGB(255, 64, 0)
        Case "yellow"
            n.color = RGB(18, 18, 0)
            n.colorfull = RGB(255, 255, 0)
        Case "green"
            n.color = RGB(0, 32, 0)
            n.colorfull = RGB(0, 200, 0)
        Case "blue"
            n.color = RGB(0, 0, 128)
            n.colorfull = RGB(0, 0, 255)
        Case "white"
            n.color = RGB(255, 252, 224)
            n.colorfull = RGB(193, 91, 0)
        Case "whitegi"
            n.color = RGB(0, 0, 0)
            n.colorfull = RGB(225, 225, 225)
        Case "purple"
            n.color = RGB(128, 0, 128)
            n.colorfull = RGB(255, 0, 255)
        Case "amber"
            n.color = RGB(193, 49, 0)
            n.colorfull = RGB(255, 153, 0)
    Case ""
    End Select

  'Set State
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If

End Sub


' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo

'    If Score(1)Then
'   DMD "", eNone, "", eNone, "", eNone, "", eNone, 1000, True, ""
'   UDMD "", "", 1000
'   'info goes in a loop only stopped by the credits and the startkey
'   DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "GAME OVER"), eBlink, 2000, False, "":PuPEvent 50 'game over
'   UDMD "GAME OVER", "", 2000
' Else
'   DMD "", eNone, "", eNone, "", eNone, CenterLine(3, " "), eBlink, 5000, False, "" 'game over
' END IF
'
'    If Score(1)Then
'        DMD "", eNone, CenterLine(1, FormatScore(Score(1))), eNone, CenterLine(2, "PLAYER 1"), eNone, "", eNone, 3000, False, ""
'   UDMD "PLAYER 1", Score(1), 3000
'    End If
'    If Score(2)Then
'        DMD "", eNone, CenterLine(1, FormatScore(Score(2))), eNone, CenterLine(2, "PLAYER 2"), eNone, "", eNone, 3000, False, ""
'   UDMD "PLAYER 2", Score(2), 3000
' End If
'    If Score(3)Then
'        DMD "", eNone, CenterLine(1, FormatScore(Score(3))), eNone, CenterLine(2, "PLAYER 3"), eNone, "", eNone, 3000, False, ""
'   UDMD "PLAYER 3", Score(3), 3000
' End If
'    If Score(4)Then
'        DMD "", eNone, CenterLine(1, FormatScore(Score(4))), eNone, CenterLine(2, "PLAYER 4"), eNone, "", eNone, 3000, False, ""
'   UDMD "PLAYER 4", Score(4), 3000
'    End If
'
'    DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "GAME OVER"), eBlink, 2000, False, "" 'game over
' UDMD "GAME OVER", "", 2000
'
' If bFreePlay Then
'        DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "FREE PLAY"), eNone, 3000, False, ""
'   UDMD "FREE PLAY", "", 3000
'    Else
'        If Credits > 0 Then
'            DMD "", eNone, CenterLine(1, "CREDITS " & Credits), eNone, CenterLine(2, "PRESS START"), eNone, "", eNone, 2000, False, ""
'     UDMD "PRESS START", "", 2000
'        Else
'            DMD "", eNone, CenterLine(1, "CREDITS " & Credits), eNone, CenterLine(2, "INSERT COIN"), eNone, "", eNone, 2000, False, ""
'     UDMD "INSERT COIN", "", 2000
'        End If
'    End If
'    DMD "", eNone, "        VPX", eNone, "     PRESENTS", eNone, "", eNone, 2000, False, ""
' UDMD "VPX", "PRESENTS", 2000
'
'    DMD "", eNone, "   TOTAL  NUCLEAR", eNone, "    ANNIHILATION", eNone, "", eNone, 2000, False, ""
' UDMD "TOTAL NUCLEAR", "ANNIHILATION", 2000
'
'    DMD "", eNone, "   TRY CO-OP MODES", eNone, "   PRESS MAGNASAVE", eNone, "", eNone, 2000, False, ""
' UDMD "TRY CO-OP MODES", "USE MAGNASAVE", 2000
'
''    DMD "", eNone, CenterLine(1, "HIGH SCORES"), eScrollLeft, Space(dCharsPerLine(2)), eScrollLeft, "", eNone, 20, False, ""
''  UDMD "", "", 20
'
'    DMD "", eNone, CenterLine(1, "HIGH SCORES"), eBlinkFast, "", eNone, "", eNone, 2000, False, ""
' UDMD "HIGH SCORES", "", 2000
'
'    DMD "", eNone, CenterLine(1, "HIGH SCORE"), eNone, "  1  " &HighScoreName(0) & " " &FormatScore(HighScore(0)), eNone, "", eNone, 2000, False, ""
' UDMD "HIGH SCORE 1", HighScoreName(0) & "  " & HighScore(0), 2000
'
'    DMD "", eNone, "_", eNone, "  2  " &HighScoreName(1) & " " &FormatScore(HighScore(1)), eNone, "", eNone, 2000, False, ""
' UDMD "HIGH SCORE 2", HighScoreName(1) & "  " &HighScore(1), 2000
'
'    DMD "", eNone, "_", eNone, "  3  " &HighScoreName(2) & " " &FormatScore(HighScore(2)), eNone, "", eNone, 2000, False, ""
' UDMD "HIGH SCORE 3", HighScoreName(2) & "  " &HighScore(2), 2000
'
'    DMD "", eNone, "_", eNone, "  4  " &HighScoreName(3) & " " &FormatScore(HighScore(3)), eNone, "", eNone, 2000, False, ""
' UDMD "HIGH SCORE 4", HighScoreName(3) & "  " &HighScore(3), 2000
'
'    DMD "", eNone, Space(dCharsPerLine(1)), eNone, Space(dCharsPerLine(2)), eNone, "", eNone, 500, False, ""
' UDMD "", "", 500
End Sub

Sub StartAttractMode(dummy)
  AttractModeHSDisplay.enabled=true
  AttractNessieBG.enabled=true
  DOF 149, DOFOn
    StartLightSeq
    DMDFlush

  If ResetHighScore = 0 Then
    ShowTableInfo
  Else
    DMD "", eNone, "  HIGH SCORE RESET", eNone, "      EXIT GAME", eNone, "", eNone, 100000, False, ""
    UDMD "HIGH SCORE RESET", "EXIT GAME", 100000
  End If
'''    StartRainbow "arrows"
End Sub

Sub StopAttractMode
  AttractModeHSDisplay.enabled=false
  AttractNessieBG.enabled=false
  DOF 149, DOFOff
    DMDFlush
  DMD "", eNone, "", eNone, "", eNone, "", eNone, 500, True, ""
  UDMD "", "", 500
'    LightSeqAttract.StopPlay
'    StopRainbow
' StopSong
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqAttract.UpdateInterval = 15
    LightSeqAttract.Play SeqCircleInOn, 40, 1

    LightSeqAttract.UpdateInterval = 2
    LightSeqAttract.Play SeqRandom, 40, , 4000
    LightSeqAttract.Play SeqAllOff

    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqCircleOutOn, 25, 4

    LightSeqAttract.UpdateInterval = 4
    LightSeqAttract.Play SeqBlinking, , 5, 150

    LightSeqAttract.UpdateInterval = 4
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.Play SeqUpOn, 25, 1, 500
    LightSeqAttract.UpdateInterval = 4
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.Play SeqUpOn, 25, 1, 500

    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqCircleOutOn, 25, 4

    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 50, 1

    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqStripe2VertOn, 50, 4

    LightSeqAttract.UpdateInterval = 4
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.Play SeqUpOn, 25, 1, 500
    LightSeqAttract.UpdateInterval = 4
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.Play SeqUpOn, 25, 1, 500

    LightSeqAttract.UpdateInterval = 2
    LightSeqAttract.Play SeqScrewRightOn, 50, 8

    LightSeqAttract.UpdateInterval = 2
    LightSeqAttract.Play SeqBlinking, , 5, 150
End Sub

Sub LightSeqAttract_PlayDone()
    StartLightSeq()
End Sub

'*************************
' Section; Rainbow Changing Lights
'*************************

Dim RGBStep, RGBFactor, Red, Green, Blue, RainbowLights

Sub StartRainbow(n)
    RainbowLights = n
    RGBStep = 0
    RGBFactor = 10
    Red = 255
    Green = 0
    Blue = 0

    RainbowTimer.Interval = 40
    RainbowTimer.Enabled = 1
End Sub

Sub StopRainbow()
    Dim obj
    RainbowTimer.Enabled = 0
    If RainbowLights = "all" Then
        For each obj in aRGBLightsMinusSome
            SetLight obj, "white", 0
        Next
    ElseIf RainbowLights = "gi" Then
        For each obj in aGiLights
            SetLight obj, "white", 0
        Next
    End If
End Sub

Sub RainbowTimer_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep
        Case 0 'Green
            Green = Green + RGBFactor
            If Green > 255 then
                Green = 255
                RGBStep = 1
            End If
        Case 1 'Red
            Red = Red - RGBFactor
            If Red < 0 then
                Red = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            Blue = Blue + RGBFactor
            If Blue > 255 then
                Blue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            Green = Green - RGBFactor
            If Green < 0 then
                Green = 0
                RGBStep = 4
            End If
        Case 4 'Red
            Red = Red + RGBFactor
            If Red > 255 then
                Red = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            Blue = Blue - RGBFactor
            If Blue < 0 then
                Blue = 0
                RGBStep = 0
            End If
    End Select
    If RainbowLights = "all" Then
        For each obj in aRGBLightsMinusSome
            obj.color = RGB(Red \ 10, Green \ 10, Blue \ 10)
            obj.colorfull = RGB(Red, Green, Blue)
        Next
    ElseIf RainbowLights = "gi" Then
        For each obj in aGiLights
            obj.color = RGB(Red \ 10, Green \ 10, Blue \ 10)
            obj.colorfull = RGB(Red, Green, Blue)
        Next
    End If
End Sub

'***********************************************************************
' *********************************************************************
' Section; Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************

' droptargets, animations, etc
Sub VPObjects_Init

End Sub

' tables variables and modes init
Dim bSuperJackpot
Dim bDoubleSuperJackpot
Dim bTripleSuperJackpot
Dim bJackpot
Dim SpinnerValue
Dim SpinnerReactorValue
Dim Multiplier2x
Dim Multiplier3x
Dim LaneBonus
Dim BallLockScore

Dim InlaneScore
Dim CenterScore
Dim BumperScore
Dim BumperDoubleScore
Dim ToplaneScore
Dim OutlaneScore
Dim StarScore
Dim LowerSlingshotScore
Dim UpperSlingshotScore
Dim Jackpot
Dim DoubleJackpot
Dim TripleJackpot
Dim SuperJackpot
Dim BonusScore
Dim TargetScore
Const TargetBonusValue = 1000
Const LaneSaveBonusValue = 10000


Sub Game_Init() 'called at the start of a new game
  If debugGeneral Then Debug.print "*****SUB:Game_Init()"
    Dim i
    bExtraBallWonThisBall = False
  ResetInserts
    'TurnOffPlayfieldLights()

    'Init Variables
    bSkillshotSelect = False
  SkillshotReady = 0

'****************************
'SubSection; Scoring Values
'***************************
    SpinnerValue = 1000
    SpinnerReactorValue = 2000
  UpperSlingshotScore = 5000
  LowerSlingshotScore = 2000
  TargetScore = 1000
    SkillshotValue = 10000
  BallLockScore = 200000
  Jackpot = 150000
  DoubleJackpot = 300000
  TripleJackpot = 450000
  SuperJackpot = 750000
  BonusScore = 1000
  BumperScore = 100
  BumperDoubleScore = 5000
  ToplaneScore = 1000
  InlaneScore = 3000
  CenterScore = 5000
  OutlaneScore = 15000
  StarScore = 1000

  'Initialize Player Data
  InitializePlayerData

  'Call Reset routines here
' ResetCORE
' ResetSAVE

    bMultiBallMode = False
    bBonusHeld = False
    Multiplier2x = 1
    Multiplier3x = 1

    'Play some Music
    StartBackgroundMusic

  'GI color
  ChangeGI GIcolor, 1
End Sub

Sub StopEndOfBallModes() 'this sub is called after the last ball is drained

    'If Modes(0)Then StopMode Modes(0) 'a mode is active so stop it
End Sub

Sub ResetNewBallVariables()           'reset variables for a new ball or player
    bSuperJackpot = False
    bDoubleSuperJackpot = False
    bTripleSuperJackpot = False
    bJackpot = False
    LaneBonus = 0
End Sub

Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
'''    Bumper1Light.Visible = 0
End Sub

' ***********************************
'  Loop Combo Section / Last Switch Hit / Ball Search Routine
' ***********************************
Dim ComboLoopFlag
Sub SetLastSwitchHit (value)
  Dim comboscore

  'increase bonus
  If NOT((value = "Bumper1") or (value = "Bumper2") or (value = "Bumper3") or (value = "swPlungerRest") or (value = "LeftSlingShot") or (value = "RightSlingShot"))  Then
    AddBonusCount
  End If

  LastSwitchHit = value

  'Reset Ball Search Timer every time a switch is hit
  If (bGameInPlay AND NOT Tilted AND BallsOnPlayfield > 0) Then
    BallSearchTimer.Enabled = False
    BallSearchTimer.Interval = BallSearchTime * 1000
    BallSearchTimer.Enabled = True
  End If
End Sub


''*********************
'' Section; Bonus Lights
''*********************

Sub AddBonusCount
  BonusCount = BonusCount + 1

  SetBonusCount
End Sub

Sub DecreaseBonusCount
  BonusCount = BonusCount - 1
  SetBonusCount
End Sub

Sub ResetBonusCount
  BonusCount = 0
  SetBonusCount
End Sub

Sub BonusCountDown
  BonusAmount(CurrentPlayer)=BonusCount
  lBonus50.TimerInterval = 100
  lBonus50.TimerEnabled = True
  PlaySound "BonusCount"

End Sub

Sub lBonus50_Timer
  DecreaseBonusCount
  AddScore BonusScore

  if BonusCount <= 0 Then
    lBonus50.TimerEnabled = False
    StopSound "BonusCount"
    If BonusMultiplier(CurrentPlayer)<2 then
      If lWhite.state=0 then
        BonusAmount(CurrentPlayer)=0
      end if
      If lWhite.state=1 AND Balls=BallsPerGame then
        BonusCount=BonusAmount(CurrentPlayer)
        SetBonusCount
        lWhite.state=0
        BonusCountDown2.enabled=true
        exit Sub
      end if
      vpmtimer.addtimer 1500, "EndOfBall2 '"
    else
      NextMultipleBonus.enabled=true
    end if
  End If
End Sub

Sub BonusCountDown2_timer
  BonusCountDown2.enabled=false
  BonusCountDown
end sub

Sub NextMultipleBonus_timer
  NextMultipleBonus.enabled=false
  BonusCount=BonusAmount(CurrentPlayer)
  SetBonusCount
  BonusMultiplier(CurrentPlayer)=BonusMultiplier(CurrentPlayer)-1
  SetBonusMultiplier(BonusMultiplier(CurrentPlayer))
  lBonus50.TimerEnabled = True
  PlaySound "BonusCount"
end sub

Dim BonusCount
Sub SetBonusCount
  Dim tmp

  'Check Ones
  tmp = (BonusCount mod 10)
  Select Case tmp
    Case 0
      lBonus1.State = 0
      lBonus2.State = 0
      lBonus3.State = 0
      lBonus4.State = 0
      lBonus5.State = 0
      lBonus6.State = 0
      lBonus7.State = 0
      lBonus8.State = 0
      lBonus9.State = 0
    Case 1
      lBonus1.State = 1
      lBonus2.State = 0
      lBonus3.State = 0
      lBonus4.State = 0
      lBonus5.State = 0
      lBonus6.State = 0
      lBonus7.State = 0
      lBonus8.State = 0
      lBonus9.State = 0
    Case 2
      lBonus1.State = 1
      lBonus2.State = 1
      lBonus3.State = 0
      lBonus4.State = 0
      lBonus5.State = 0
      lBonus6.State = 0
      lBonus7.State = 0
      lBonus8.State = 0
      lBonus9.State = 0
    Case 3
      lBonus1.State = 1
      lBonus2.State = 1
      lBonus3.State = 1
      lBonus4.State = 0
      lBonus5.State = 0
      lBonus6.State = 0
      lBonus7.State = 0
      lBonus8.State = 0
      lBonus9.State = 0
    Case 4
      lBonus1.State = 1
      lBonus2.State = 1
      lBonus3.State = 1
      lBonus4.State = 1
      lBonus5.State = 0
      lBonus6.State = 0
      lBonus7.State = 0
      lBonus8.State = 0
      lBonus9.State = 0
    Case 5
      lBonus1.State = 1
      lBonus2.State = 1
      lBonus3.State = 1
      lBonus4.State = 1
      lBonus5.State = 1
      lBonus6.State = 0
      lBonus7.State = 0
      lBonus8.State = 0
      lBonus9.State = 0
    Case 6
      lBonus1.State = 1
      lBonus2.State = 1
      lBonus3.State = 1
      lBonus4.State = 1
      lBonus5.State = 1
      lBonus6.State = 1
      lBonus7.State = 0
      lBonus8.State = 0
      lBonus9.State = 0
    Case 7
      lBonus1.State = 1
      lBonus2.State = 1
      lBonus3.State = 1
      lBonus4.State = 1
      lBonus5.State = 1
      lBonus6.State = 1
      lBonus7.State = 1
      lBonus8.State = 0
      lBonus9.State = 0
    Case 8
      lBonus1.State = 1
      lBonus2.State = 1
      lBonus3.State = 1
      lBonus4.State = 1
      lBonus5.State = 1
      lBonus6.State = 1
      lBonus7.State = 1
      lBonus8.State = 1
      lBonus9.State = 0
    Case 9
      lBonus1.State = 1
      lBonus2.State = 1
      lBonus3.State = 1
      lBonus4.State = 1
      lBonus5.State = 1
      lBonus6.State = 1
      lBonus7.State = 1
      lBonus8.State = 1
      lBonus9.State = 1
  End Select

  'Check Tens
  If BonusCount < 10 Then
    lBonus10.State = 0
    lBonus20.State = 0
    lBonus30.State = 0
    lBonus40.State = 0
    lBonus50.State = 0
  ElseIf BonusCount >= 10 and BonusCount < 20 Then
    lBonus10.State = 1
    lBonus20.State = 0
    lBonus30.State = 0
    lBonus40.State = 0
    lBonus50.State = 0
  ElseIf BonusCount >= 20 and BonusCount < 30 Then
    lBonus10.State = 0
    lBonus20.State = 1
    lBonus30.State = 0
    lBonus40.State = 0
    lBonus50.State = 0
  ElseIf BonusCount >= 30 and BonusCount < 40 Then
    lBonus10.State = 0
    lBonus20.State = 0
    lBonus30.State = 1
    lBonus40.State = 0
    lBonus50.State = 0
  ElseIf BonusCount >= 40 and BonusCount < 50 Then
    lBonus10.State = 0
    lBonus20.State = 0
    lBonus30.State = 0
    lBonus40.State = 1
    lBonus50.State = 0
  ElseIf BonusCount >= 50 and BonusCount < 60 Then
    lBonus10.State = 0
    lBonus20.State = 0
    lBonus30.State = 0
    lBonus40.State = 0
    lBonus50.State = 1
  ElseIf BonusCount >= 60 and BonusCount < 70 Then
    lBonus10.State = 1
    lBonus20.State = 0
    lBonus30.State = 0
    lBonus40.State = 0
    lBonus50.State = 1
  ElseIf BonusCount >= 70 and BonusCount < 80 Then
    lBonus10.State = 0
    lBonus20.State = 1
    lBonus30.State = 0
    lBonus40.State = 0
    lBonus50.State = 1
  ElseIf BonusCount >= 80 and BonusCount < 90 Then
    lBonus10.State = 0
    lBonus20.State = 0
    lBonus30.State = 1
    lBonus40.State = 0
    lBonus50.State = 1
  ElseIf BonusCount >= 90 and BonusCount < 100 Then
    lBonus10.State = 0
    lBonus20.State = 0
    lBonus30.State = 0
    lBonus40.State = 1
    lBonus50.State = 1
  ElseIf BonusCount >= 100 and BonusCount < 110 Then
    lBonus10.State = 1
    lBonus20.State = 0
    lBonus30.State = 0
    lBonus40.State = 1
    lBonus50.State = 1
  ElseIf BonusCount >= 110 and BonusCount < 120 Then
    lBonus10.State = 0
    lBonus20.State = 1
    lBonus30.State = 0
    lBonus40.State = 1
    lBonus50.State = 1
  ElseIf BonusCount >= 120 and BonusCount < 130 Then
    lBonus10.State = 0
    lBonus20.State = 0
    lBonus30.State = 1
    lBonus40.State = 1
    lBonus50.State = 1
  ElseIf BonusCount >= 130 and BonusCount < 140 Then
    lBonus10.State = 1
    lBonus20.State = 0
    lBonus30.State = 1
    lBonus40.State = 1
    lBonus50.State = 1
  ElseIf BonusCount >= 140 and BonusCount < 150 Then
    lBonus10.State = 0
    lBonus20.State = 1
    lBonus30.State = 1
    lBonus40.State = 1
    lBonus50.State = 1
  Else
    lBonus10.State = 1
    lBonus20.State = 1
    lBonus30.State = 1
    lBonus40.State = 1
    lBonus50.State = 1
  End If

End Sub

Const BallSearchTime = 3000 'Seconds
Sub BallSearchTimer_Timer 'only triggered if no switches hit for x seconds

  If (bGameInPlay AND NOT Tilted AND BallsOnPlayfield > 0) Then
    If ((LeftFlipper.CurrentAngle > 90 ) AND (RightFlipper.CurrentAngle < -90)) Then  'If flippers are not up holding a ball
      If bBallInPlungerLane = 0 Then 'If ball not resting in plungerlane

'       DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "BALL SEARCH"), eNone, 2000, True, "tna_leftscoopawardeject"
'       UDMD "BALL SEARCH", "Ejecting ball", 2000
        vpmtimer.addtimer 1100, "BallSearchEject '"
      End If
    End If
  End If
End Sub

Sub BallSearchEject
' LeftScoop.Createball
' LeftScoop.Kick 165, LeftScoopStrength, 1.56
' LeftScoop.Enabled = True
End Sub


''*****************************
'' Section; Top Lane
''*****************************
Sub ResetTopLane()
  LTop1.State = 0
  LTop2.State = 0
  LTop3.State = 0
End Sub

Sub CheckTopLane()
  If (LTop1.State = 1 and LTop2.State = 1 and LTop3.State = 1) Then
    LTop1.State = 2
    LTop2.State = 2
    LTop3.State = 2
    vpmTimer.AddTimer 2000, "AwardBonusMultiplier '"
    BumperLight001.state=2
    BumperLight002.state=2
    BumperLight003.state=2
    BumperFlashingTimer.enabled=true
  End If
End Sub

Sub RotateLaneLightsLeft()
    Dim tmp
    tmp = lTop1.State
    lTop1.state = lTop2.State
    lTop2.State = lTop3.State
    lTop3.State = tmp

End Sub
'
Sub RotateLaneLightsRight()
    Dim tmp
    tmp = lTop3.State
    lTop3.State = lTop2.State
    lTop2.State = lTop1.State
    lTop1.State = tmp
End Sub

Sub AwardBonusMultiplier
  ResetTopLane
  AddBonusMultiplier 1
End Sub
''*****************************
'' Section; Bonus Multiplier
''*****************************
Sub ResetMultiplier
  BonusMultiplier(CurrentPlayer) = 1
  SetBonusMultiplier 1
End Sub


Sub AddBonusMultiplier (n)
    If debugGeneral Then debug.print "*****SUB:" & "AddBonusMultiplier(" & n & ")"

    ' if not at the maximum bonus level
    if(BonusMultiplier(CurrentPlayer) + n <= MaxMultiplier)then
        ' then add and set the lights
        BonusMultiplier(CurrentPlayer) = BonusMultiplier(CurrentPlayer) + n
        SetBonusMultiplier BonusMultiplier(CurrentPlayer)

    End if
End Sub

Sub SetBonusMultiplier (n)
    If debugGeneral Then debug.print "*****SUB:" & "SetBonusMultiplier(" & n & ")"

  Select Case n
    Case 1
      l2x.State = 0
      l3x.State = 0
      l4x.State = 0
      l5x.State = 0
    Case 2
      l2x.State = 1
      l3x.State = 0
      l4x.State = 0
      l5x.State = 0
    Case 3
      l2x.State = 1
      l3x.State = 1
      l4x.State = 0
      l5x.State = 0
    Case 4
      l2x.State = 1
      l3x.State = 1
      l4x.State = 1
      l5x.State = 0
    Case 5
      l2x.State = 1
      l3x.State = 1
      l4x.State = 1
      l5x.State = 1
  End Select
End Sub




'**************************
' Section; SAVE - Inlanes Outlanes
'**************************
Dim LaneSaveCount(4)
Dim PlayerLa1(4)
Dim PlayerLa2(4)
Dim PlayerLa3(4)
Dim PlayerLa4(4)
Const LaneSaveCountMax = 6
Dim bBallSaverSingleUse

Sub ResetSave() ' ClearSave on end ball or start ball
  LaneSaveCount(CurrentPlayer) = 0

  SetLight la1, "red", 0
  SetLight la2, "red", 0
  SetLight la3, "red", 0
  SetLight la4, "red", 0


  ttSaves.text = "S:" & LaneSaveCount(CurrentPlayer)
End Sub

Sub CheckSAVE()
  dim tmp
  tmp = 0

  if la1.state = 1 then tmp = tmp + 1
  if la2.state = 1 then tmp = tmp + 1
  if la3.state = 1 then tmp = tmp + 1
  if la4.state = 1 then tmp = tmp + 1

  If tmp = 3 then '3 targets lit, earn a SAVE
    AwardSAVE 1
  End If

End Sub

Sub AwardSAVE (value)
  Dim SaveInsertAlreadyLit, savecolor
  SaveInsertAlreadyLit = 0

  LaneSaveCount(CurrentPlayer) = LaneSaveCount(CurrentPlayer) + value
  If LaneSaveCount(CurrentPlayer) > LaneSaveCountMax Then
    LaneSaveCount(CurrentPlayer) = LaneSaveCountMax
  Else
    If LaneSaveCount(CurrentPlayer) = 1 Then
      DOF 170, DOFPulse: DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "LANE SAVE " & LaneSaveCount(CurrentPlayer)), eBlinkfast, 1500, True, "tna_lanesavelevelone"
      UDMD "LANE SAVE", "LEVEL 1", 1500
    Else
      DOF 170, DOFPulse: DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "LANE SAVE " & LaneSaveCount(CurrentPlayer)), eBlinkfast, 1500, True, "tna_lanesaveincreased"
      UDMD "LANE SAVE", "LEVEL " & LaneSaveCount(CurrentPlayer), 1500
    End If
    GIGame 1, "yellow"
  End If
  Debug.Print "LaneSaveCount(CurrentPlayer) = " & LaneSaveCount(CurrentPlayer)
  ttSaves.text = "S:" & LaneSaveCount(CurrentPlayer)


  'Select new Insert color
  SetSaveColor

  'Check if any Save insert already blinking and change color based on Save count (Red, Orange, Yellow,Green, Blue, Purple)
  if la1.state = 2 then
    SaveInsertAlreadyLit = 1
  elseif la2.state = 2 then
    SaveInsertAlreadyLit = 1
  elseif la3.state = 2 then
    SaveInsertAlreadyLit = 1
  elseif la4.state = 2 then
    SaveInsertAlreadyLit = 1
  end if

  'If First Save awarded, set Flashing save insert
  if SaveInsertAlreadyLit=0 Then
    if la1.state = 0 then
      la1.state = 2
    elseif la2.state = 0 then
      la2.state = 2
    elseif la3.state = 0 then
      la3.state = 2
    elseif la4.state = 0 then
      la4.state = 2
    end if
  End If

  'Reset the rest of the SAVE inserts
  if la1.state = 1 then
    la1.state = 0
  end if
  if la2.state = 1 then
    la2.state = 0
  end if
  if la3.state = 1 then
    la3.state = 0
  end if
  if la4.state = 1 then
    la4.state = 0
  end if

End Sub


Sub UseSAVE (sw) ' On outlane, check for ballsave condition
  dim savecolor
  dim tmp
  tmp = ""
  if LaneSaveCount(CurrentPlayer) > 0 Then
    'Check if SAVE insert is matching the drain outlane
    if la1.state = 2 then
      tmp = "swOutlaneL"
    elseif la4.state = 2 then
      tmp = "swOutlaneR"
    end if

    'debug.print "Entering usesave " & LaneSaveCount(CurrentPlayer) & " : " & sw & " ; " & tmp

    if (StrComp(sw, tmp) = 0)then 'Ball Saved!

      'debug.print "ballSavesingle!"
      bBallSaverSingleUse = 1
      LaneSaveCount(CurrentPlayer) = LaneSaveCount(CurrentPlayer) - 1

      Playsound "tna_ballsaved"
      DOF 165, DOFPulse: DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "BALL SAVED"), eBlinkfast, 800, True, ""
      UDMD "BALL SAVED", "", 800
      GIGameImmediate 12, "green"


      if LaneSaveCount(CurrentPlayer) <=0 Then ' Turn off Save insert
        LaneSaveCount(CurrentPlayer) = 0
        if la1.state = 2 then
          la1.state = 0
        elseif la2.state = 2 then
          la2.state = 0
        elseif la3.state = 2 then
          la3.state = 0
        elseif la4.state = 2 then
          la4.state = 0
        end if
      End If

      'Select new Insert color
      SetSaveColor

    End If
  End If
  ttSaves.text = "S:" & LaneSaveCount(CurrentPlayer)

    DMDScoreNow
End Sub


Sub SetSaveColor
  dim savecolor

  'Select new Insert color
  Select Case LaneSaveCount(CurrentPlayer)
    Case 0: savecolor = "red"
    Case 1: savecolor = "red"
    Case 2: savecolor = "orange"
    Case 3: savecolor = "yellow"
    Case 4: savecolor = "green"
    Case 5: savecolor = "blue"
    Case else:  savecolor = "purple"
  End Select

  'Set insert color
  SetLight la1, savecolor, -1
  SetLight la2, savecolor, -1
  SetLight la3, savecolor, -1
  SetLight la4, savecolor, -1
End Sub

Sub InitLaneSaveData
  Dim i
    For i = 1 To MaxPlayers
    LaneSaveCount(i) = 0
    PlayerLa1(i) = 0
    PlayerLa2(i) = 0
    PlayerLa3(i) = 0
    PlayerLa4(i) = 0
    Next
End Sub

Sub SaveLaneSaveData
  'LaneSaveCount(i) is already saved
  PlayerLa1(CurrentPlayer) = la1.state
  PlayerLa2(CurrentPlayer) = la2.state
  PlayerLa3(CurrentPlayer) = la3.state
  PlayerLa4(CurrentPlayer) = la4.state
End Sub

Sub RestoreLaneSaveData
  la1.state = PlayerLa1(CurrentPlayer)
  la2.state = PlayerLa2(CurrentPlayer)
  la3.state = PlayerLa3(CurrentPlayer)
  la4.state = PlayerLa4(CurrentPlayer)

  SetSaveColor
End Sub

Sub CopyLaneSaveData (p1, p2)
  LaneSaveCount(p2) = LaneSaveCount(p1)
  PlayerLa1(p2) = PlayerLa1(p1)
  PlayerLa2(p2) = PlayerLa2(p1)
  PlayerLa3(p2) = PlayerLa3(p1)
  PlayerLa4(p2) = PlayerLa4(p1)
End Sub

'**************************
' Section; Skillshot
'**************************
Sub StartSkillShot() 'Updates the DMD & lights with the chosen skillshots
    'DMDFlush
    bSkillShotSelect = True
End Sub

Sub SelectSkillshot(index)
End Sub

Sub CheckSkillShot (index) ' reset the skillshot lights & variables
    SkillShotReady = 0
    bSkillShotSelect = False

'    DMDScoreNow
End Sub



' *********************************************************************
'                        Table Object Hit Events
'
' Any target hit Sub will follow this:
' - play a sound
' - do some physical movement
' - add a score, bonus
' - check some variables/modes this trigger is a member of
' - set the "LastSwicthHit" variable in case it is needed later
' *********************************************************************

' Slingshots has been hit

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fxz_leftslingshot", 103, DOFPulse, DOFContactors), 0, 1, -0.05, 0.05
  PlaySound "sling"
    DOF 104, DOFPulse
    Lsling1.Visible = 1:Lsling.Visible = 0
    slingl.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 10
  FlipAltRelay
    ' add some effect to the table?
' GILeftSlingHit

    ' remember last trigger hit by the ball
    SetLastSwitchHit "LeftSlingShot"
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:Lsling1.Visible = 0:Lsling2.Visible = 1:slingl.TransZ = -5
        Case 2:Lsling2.Visible = 0:Lsling.Visible = 1:slingl.TransZ =0:LeftSlingShot.TimerEnabled = False
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fxz_rightslingshot", 105, DOFPulse, DOFContactors), 0, 1, 0.05, 0.05
  PlaySound "sling"
  DOF 106, DOFPulse
    Rsling1.Visible = 1:Rsling.Visible = 0
  slingr.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 10
  FlipAltRelay
    ' add some effect to the table?
' GIRightSlingHit

    ' remember last trigger hit by the ball
    SetLastSwitchHit "RightSlingShot"
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:Rsling1.Visible = 0:Rsling2.Visible = 1:slingr.TransZ = -5
        Case 2:Rsling2.Visible = 0:rsling.Visible = 1:slingr.TransZ = 0:RightSlingShot.TimerEnabled = False
    End Select
    RStep = RStep + 1
End Sub

'******************
' Section; leaf switches
'******************

Sub phys_leafs_hit
  If Tilted then Exit Sub
  AddScore 10
  FlipAltRelay
end sub


'******************
' Section; Ramp
'******************
Sub RampTrigger_Hit
  Dim tempextra
  if Tilted then Exit Sub
  If lScore.state=2 then
    tempextra=Int(RND(25)+1)
    Addscore (25000+(tempextra*1000))
    lScore.state=0
    lMystery.state=0
  else
    AddScore 25000
  end if
  If BallIsLocked=true then
    BallIsLocked=false
    flash2.timerenabled=true
    bLockIsLit = False
    lBallRelease.state=0
    NessieTimer.enabled=true
    NessieBGTimer.enabled=true
    PlaySound "Nessy_Roar"
    lWhenFlashing.state=2
    BallLockedBlocker.IsDropped=1
    lLock.state=0
    bMultiBallMode = True
    AddBallsOnPlayfield 1
    vpmTimer.addTimer 1500, "BallLockEject '"
  end if
end sub


'******************
' Section; Spinner
'******************
Dim SpinCount, SuperSpinnerValue
Sub Spinner1_Spin()
    If Tilted Then Exit Sub
  DOF 112, DOFPulse

  AddScore SpinnerValue
  Playsound "spin"

End Sub

'*********************
' Section; Inlanes - Outlanes
'*********************
Sub swOutlaneL_Hit()
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
  If lOutlaneL.state=1 then
    OutlaneScoreLNM
  else
    AddScore OutlaneScore
  end if

    ' change some light
  PlaySound "Outlane"
    SetLastSwitchHit "swOutlaneL"
  wire001.z=-2
End Sub

Sub swOutlaneL_UnHit()
  wire001.z=0
end sub

Sub swOutlaneR_Hit()
    PlaySoundAtVol "fx_sensor",ActiveBall, 1
    If Tilted Then Exit Sub
  If lOutlaneR.state=1 then
    OutlaneScoreLNM
  else
    AddScore OutlaneScore
  end if

    ' change some light
  PlaySound "Outlane"
  SetLastSwitchHit "swOutlaneR"
  wire004.z=-2
End Sub

Sub swOutlaneR_UnHit()
  wire004.z=0
end sub

Sub swInlaneL_Hit()
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    AddScore InlaneScore
  If lInlaneL.state=1 then
    lWhite.state=1
  end if
    ' change some light
  PlaySound "Return"
  SetLastSwitchHit "swInlaneL"
  wire002.z=-2
End Sub

Sub swInlaneL_UnHit()
  wire002.z=0
end sub

Sub swInlaneR_Hit()
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    AddScore InlaneScore
  If lInlaneR.state=1 then
    lWhite.state=1
  end if
    ' change some light
  PlaySound "Return"
  SetLastSwitchHit "swInlaneR"
  wire003.z=-2
End Sub

Sub swInlaneR_UnHit()
  wire003.z=0
end sub

Sub swCenter_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    AddScore CenterScore
  SetLastSwitchHit "swCenter"
  wire005.z=-2
End Sub

Sub swCenter_UnHit()
  wire005.z=0
end sub

Sub OutlaneScoreLNM
  If LNMCounter(CurrentPlayer)=0 then AddScore 50000
  If LNMCounter(CurrentPlayer)=1 then AddScore 100000
  If LNMCounter(CurrentPlayer)=2 then AwardExtraBall
  If LNMCounter(CurrentPlayer)=3 then AwardSpecial
  If LNMCounter(CurrentPlayer)=4 then AwardJackpot
end sub

'*********************
' Section; Top Lanes
'*********************
Sub swTop1_Hit()
  DOF 150, DOFPulse
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub

  PlaySound "toplane"
    SetLastSwitchHit "swTop1"
  If SkillShotReady <> 0 Then CheckSkillShot 1
  If BumperLight001.state=0 then
    BumperLight001.state=1
  end if
  If LTop1.state <> 1 then
    playsound "tna_toplane"
    Addscore TopLaneScore
  else
    AddScore 3000
  end if
  LTop1.state = 1


  wire006.z=-2
  CheckTopLane
End Sub

Sub swTop1_UnHit()
  wire006.z=0
end sub

Sub swTop2_Hit()
  DOF 151, DOFPulse
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub

  PlaySound "toplane"
    SetLastSwitchHit "swTop2"
  If SkillShotReady <> 0 Then CheckSkillShot 2
  If BumperLight002.state=0 then
    BumperLight002.state=1
  end if
  If lTop2.state <> 1 then
    playsound "tna_toplane"
    Addscore TopLaneScore
  else
    AddScore 3000
  end if
  lTop2.state = 1


  wire007.z=-2
  CheckTopLane
End Sub

Sub swTop2_UnHit()
  wire007.z=0
end sub

Sub swTop3_Hit()
  DOF 152, DOFPulse
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub

  PlaySound "toplane"
    SetLastSwitchHit "swTop3"
  If SkillShotReady <> 0 Then CheckSkillShot 3
  If BumperLight003.state=0 then
    BumperLight003.state=1
  end if
  If LTop3.state <> 1 then
    playsound "tna_toplane"
    Addscore TopLaneScore
  else
    AddScore 3000
  end if
  LTop3.state = 1


  wire008.z=-2
  CheckTopLane
End Sub

Sub swTop3_UnHit()
  wire008.z=0
end sub

''*********************
'' Section; Stand Up Target
''*********************
Sub Target001_Hit

  targetstandup.transY=-2
  If Tilted Then Exit Sub
  AddScore 10000

end sub

Sub Target001_UnHit
  targetstandup.transY=0
end sub

Sub tsw4_Hit
  targetstandup.transY=-2
end sub

Sub tsw4_UnHit
  targetstandup.transY=0
end sub

'Example routines:  Reset, Start, Stop, Set, Check, Save, Restore, Copy, Test


''*********************
'' Section; Star Rollovers
''*********************
Sub swStar1_hit
  AddScore StarScore
  SetLastSwitchHit "swStar1"
End Sub

Sub swStar2_hit
  AddScore StarScore
  if lArrow.state=1 then AwardExtraBall:lArrow.state=0
  SetLastSwitchHit "swStar2"
End Sub




''*********************
'' Section; Targets Bank 1
''*********************
Sub Target1_Hit
  If debugTargets Then debug.print "*****SUB:Target1_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 114, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore
  DOF 191, DOFPulse
  Target1.IsDropped=1:ptarget1.transz=-45:Target1.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target1"
  Bank1Dropped(CurrentPlayer,1)=1
  CheckTargetBank1
End Sub

Sub Target2_Hit
  If debugTargets Then debug.print "*****SUB:Target2_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 114, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore

  DOF 191, DOFPulse
  Target2.IsDropped=1:ptarget2.transz=-45:Target2.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target2"
  Bank1Dropped(CurrentPlayer,2)=1
  CheckTargetBank1
End Sub

Sub Target3_Hit
  If debugTargets Then debug.print "*****SUB:Target3_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 114, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore

  DOF 191, DOFPulse
  Target3.IsDropped=1:ptarget3.transz=-45:Target3.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target3"
  Bank1Dropped(CurrentPlayer,3)=1
  CheckTargetBank1
End Sub

'------------------------
Sub Target1_Dropped:  CheckTargetBank1: End Sub
Sub Target2_Dropped:  CheckTargetBank1: End Sub
Sub Target3_Dropped:  CheckTargetBank1: End Sub

'------------------------
Sub ResetTargetBankTimer001_timer
  If debugTargets Then debug.print "*****SUB:ResetTargetBank1"

  if Bank1Dropped(CurrentPlayer,1)=0 then
    Target1.IsDropped = 0:ptarget1.transz=0:Target1.Collidable=true
  else
    Target1.IsDropped = 1:ptarget1.transz=-45:Target1.Collidable=false
  end if

  if Bank1Dropped(CurrentPlayer,2)=0 then
    Target2.IsDropped = 0:ptarget2.transz=0:Target2.Collidable=true
  else
    Target2.IsDropped = 1:ptarget2.transz=-45:Target2.Collidable=false
  end if

  if Bank1Dropped(CurrentPlayer,3)=0 then
    Target3.IsDropped = 0:ptarget3.transz=0:Target3.Collidable=true
  else
    Target3.IsDropped = 1:ptarget3.transz=-45:Target3.Collidable=false
  end if
    select case Bank1Counter(CurrentPlayer)
      case 0:
        lLochL.state=0
        lLochO.state=0
        lLochC.state=0
        lLochH.state=0
      case 1:
        lLochL.state=1
        lLochO.state=0
        lLochC.state=0
        lLochH.state=0
      case 2:
        lLochL.state=1
        lLochO.state=1
        lLochC.state=0
        lLochH.state=0
      case 3:
        lLochL.state=1
        lLochO.state=1
        lLochC.state=1
        lLochH.state=0
      case 4,5,6,7:
        lLochL.state=1
        lLochO.state=1
        lLochC.state=1
        lLochH.state=1
    end select
  CheckLNMSequence
' SetLight lt1, "red", 0
' SetLight lt2, "red", 0
' SetLight lt3, "red", 0
  ResetTargetBankTimer001.enabled=false
End Sub

Sub StartTargetBank1
  If debugTargets Then debug.print "*****SUB:StartTargetBank1"
' SetLight lt1, "red", 2
' SetLight lt2, "red", 0
' SetLight lt3, "red", 0
End Sub

Sub CheckTargetBank1
  Dim j
  If debugTargets Then debug.print "*****SUB:CheckTargetBank1"

  If ((ptarget1.transz=-45) and (ptarget2.transz=-45) and (ptarget3.transz=-45)) Then
    'vpmTimer.addTimer 2000, "ResetTargetBank1 '"
    Bank1Counter(CurrentPlayer)=Bank1Counter(CurrentPlayer)+1
    If Bank1Counter(CurrentPlayer)>7 then Bank1Counter(CurrentPlayer)=7
    for j = 1 to 3
      Bank1Dropped(CurrentPlayer,j)=0
    next
    lMystery.state=2
    lScore.state=2
    lLoch.state=1
    If Bank1Counter(CurrentPlayer)>1 then
      lInlaneL.state=1
      lInlaneR.state=0
    end if
    CollectLNMSequence

    ResetTargetBankTimer001.enabled=true
  End If
End Sub

Sub SaveTargetBank1
  If debugTargets Then debug.print "*****SUB:SaveTargetBank1"

End Sub

Sub RestorTargetBank1
  If debugTargets Then debug.print "*****SUB:RestorTargetBank1"
End Sub

''*********************
'' Section; Targets Bank 2
''*********************
Sub Target4_Hit
  If Target4.IsDropped=1 then exit sub
  If debugTargets Then debug.print "*****SUB:Target4_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 115, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore

  DOF 191, DOFPulse
  Target4.IsDropped=1:ptarget4.transz=-45:Target4.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target4"
  Bank2Dropped(CurrentPlayer,1)=1
  CheckTargetBank2
End Sub

Sub Target5_Hit
  If Target5.IsDropped=1 then exit sub
  If debugTargets Then debug.print "*****SUB:Target5_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 115, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore

  DOF 191, DOFPulse
  Target5.IsDropped=1:ptarget5.transz=-45:Target5.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target5"
  Bank2Dropped(CurrentPlayer,2)=1
  CheckTargetBank2
End Sub

Sub Target6_Hit
  If Target6.IsDropped=1 then exit sub
  If debugTargets Then debug.print "*****SUB:Target6_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 115, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore

  DOF 191, DOFPulse
  Target6.IsDropped=1:ptarget6.transz=-45:Target6.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target6"
  Bank2Dropped(CurrentPlayer,3)=1
  CheckTargetBank2
End Sub

'------------------------
Sub Target4_Dropped:  CheckTargetBank2: End Sub
Sub Target5_Dropped:  CheckTargetBank2: End Sub
Sub Target6_Dropped:  CheckTargetBank2: End Sub

'------------------------
Sub ResetTargetBankTimer002_timer
  If debugTargets Then debug.print "*****SUB:ResetTargetBank2"

  if Bank2Dropped(CurrentPlayer,1)=0 then
    Target4.IsDropped=0:ptarget4.transz=0:Target4.Collidable=true
  else
    Target4.IsDropped=1:ptarget4.transz=-45:Target4.Collidable=false
  end if

  if Bank2Dropped(CurrentPlayer,2)=0 then
    Target5.IsDropped=0:ptarget5.transz=0:Target5.Collidable=true
  else
    Target5.IsDropped=1:ptarget5.transz=-45:Target5.Collidable=false
  end if

  if Bank2Dropped(CurrentPlayer,3)=0 then
    Target6.IsDropped=0:ptarget6.transz=0:Target6.Collidable=true
  else
    Target6.IsDropped=1:ptarget6.transz=-45:Target6.Collidable=false
  end if
    select case Bank2Counter(CurrentPlayer)
      case 0:
        lNessN.state=0
        lNessE.state=0
        lLessS1.state=0
        lLessL2.state=0
        l1000Plus.state=1
        l2000Plus.state=0
        l4000Plus.state=0
        SpinnerValue=1000
      case 1:
        lNessN.state=1
        lNessE.state=0
        lLessS1.state=0
        lLessL2.state=0
        l1000Plus.state=0
        l2000Plus.state=1
        l4000Plus.state=0
        SpinnerValue=2000
      case 2:
        lNessN.state=1
        lNessE.state=1
        lLessS1.state=0
        lLessL2.state=0
        l1000Plus.state=1
        l2000Plus.state=1
        l4000Plus.state=0
        SpinnerValue=3000
      case 3:
        lNessN.state=1
        lNessE.state=1
        lLessS1.state=1
        lLessL2.state=0
        l1000Plus.state=0
        l2000Plus.state=0
        l4000Plus.state=1
        SpinnerValue=4000
      case 4:
        lNessN.state=1
        lNessE.state=1
        lLessS1.state=1
        lLessL2.state=1
        l1000Plus.state=1
        l2000Plus.state=0
        l4000Plus.state=1
        SpinnerValue=5000
      case 5:
        lNessN.state=1
        lNessE.state=1
        lLessS1.state=1
        lLessL2.state=1
        l1000Plus.state=0
        l2000Plus.state=1
        l4000Plus.state=1
        SpinnerValue=6000
      case 6,7:
        lNessN.state=1
        lNessE.state=1
        lLessS1.state=1
        lLessL2.state=1
        l1000Plus.state=1
        l2000Plus.state=1
        l4000Plus.state=1
        SpinnerValue=7000
    end select
  CheckLNMSequence
' SetLight lt1, "red", 0
' SetLight lt2, "red", 0
' SetLight lt3, "red", 0
  ResetTargetBankTimer002.enabled=false
End Sub

Sub StartTargetBank2
  If debugTargets Then debug.print "*****SUB:StartTargetBank2"
' SetLight lt1, "red", 2
' SetLight lt2, "red", 0
' SetLight lt3, "red", 0
End Sub

Sub CheckTargetBank2
  Dim j
  If debugTargets Then debug.print "*****SUB:CheckTargetBank2"

  If ((ptarget4.transz=-45) and (ptarget5.transz=-45) and (ptarget6.transz=-45)) Then
    'vpmTimer.addTimer 2000, "ResetTargetBank2 '"
    Bank2Counter(CurrentPlayer)=Bank2Counter(CurrentPlayer)+1
    If Bank2Counter(CurrentPlayer)>7 then Bank2Counter(CurrentPlayer)=7
    If Bank2Counter(CurrentPlayer)>3 then lLock.state=2:bLockIsLit=true
    for j = 1 to 3
      Bank2Dropped(CurrentPlayer,j)=0
    next
    lNess.state=1
    CollectLNMSequence
    ResetTargetBankTimer002.enabled=true
  End If
End Sub

''*********************
'' Section; Targets Bank 3
''*********************
Sub Target7_Hit
  If debugTargets Then debug.print "*****SUB:Target7_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 116, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore

  DOF 191, DOFPulse
  Target7.IsDropped=1:ptarget7.transz=-45:Target7.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target7"
  Bank3Dropped(CurrentPlayer,1)=1
  CheckTargetBank3
End Sub

Sub Target8_Hit
  If debugTargets Then debug.print "*****SUB:Target8_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 116, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore

  DOF 191, DOFPulse
  Target8.IsDropped=1:ptarget8.transz=-45:Target8.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target8"
  Bank3Dropped(CurrentPlayer,2)=1
  CheckTargetBank3
End Sub

Sub Target9_Hit
  If debugTargets Then debug.print "*****SUB:Target9_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 116, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore

  DOF 191, DOFPulse
  Target9.IsDropped=1:ptarget9.transz=-45:Target9.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target9"
  Bank3Dropped(CurrentPlayer,3)=1
  CheckTargetBank3
End Sub

Sub Target10_Hit
  If debugTargets Then debug.print "*****SUB:Target10_Hit"

  PlaySoundAtVol SoundFXDOF("fx_droptarget", 116, DOFPulse, DOFDropTargets), Me, 1
  AddScore TargetScore

  DOF 191, DOFPulse
  Target10.IsDropped=1:ptarget10.transz=-45:Target10.Collidable=false
  Playsound "Drop"
    SetLastSwitchHit "Target10"
  Bank3Dropped(CurrentPlayer,4)=1
  CheckTargetBank3
End Sub

'------------------------
Sub Target7_Dropped:  CheckTargetBank3: End Sub
Sub Target8_Dropped:  CheckTargetBank3: End Sub
Sub Target9_Dropped:  CheckTargetBank3: End Sub
Sub Target10_Dropped:   CheckTargetBank3: End Sub

'------------------------
Sub ResetTargetBankTimer003_timer
  If debugTargets Then debug.print "*****SUB:ResetTargetBank3"

  if Bank3Dropped(CurrentPlayer,1)=0 then
    Target7.IsDropped=0:ptarget7.transz=0:Target7.Collidable=true
  else
    Target7.IsDropped=1:ptarget7.transz=-45:Target7.Collidable=false
  end if

  if Bank3Dropped(CurrentPlayer,2)=0 then
    Target8.IsDropped=0:ptarget8.transz=0:Target8.Collidable=true
  else
    Target8.IsDropped=1:ptarget8.transz=-45:Target8.Collidable=false
  end if

  if Bank3Dropped(CurrentPlayer,3)=0 then
    Target9.IsDropped=0:ptarget9.transz=0:Target9.Collidable=true
  else
    Target9.IsDropped=1:ptarget9.transz=-45:Target9.Collidable=false
  end if

  if Bank3Dropped(CurrentPlayer,4)=0 then
    Target10.IsDropped=0:ptarget10.transz=0:Target10.Collidable=true
  else
    Target10.IsDropped=1:ptarget10.transz=-45:Target10.Collidable=false
  end if
    select case Bank3Counter(CurrentPlayer)
      case 0:
        lMonM.state=0
        lMonO.state=0
        lMonN.state=0
        lMonS.state=0
        lMonT.state=0
        lMonE.state=0
        lMonR.state=0
      case 1:
        lMonM.state=1
        lMonO.state=0
        lMonN.state=0
        lMonS.state=0
        lMonT.state=0
        lMonE.state=0
        lMonR.state=0

      case 2:
        lMonM.state=1
        lMonO.state=1
        lMonN.state=0
        lMonS.state=0
        lMonT.state=0
        lMonE.state=0
        lMonR.state=0
      case 3:
        lMonM.state=1
        lMonO.state=1
        lMonN.state=1
        lMonS.state=0
        lMonT.state=0
        lMonE.state=0
        lMonR.state=0

      case 4:
        lMonM.state=1
        lMonO.state=1
        lMonN.state=1
        lMonS.state=1
        lMonT.state=0
        lMonE.state=0
        lMonR.state=0

      case 5:
        lMonM.state=1
        lMonO.state=1
        lMonN.state=1
        lMonS.state=1
        lMonT.state=1
        lMonE.state=0
        lMonR.state=0

      case 6:
        lMonM.state=1
        lMonO.state=1
        lMonN.state=1
        lMonS.state=1
        lMonT.state=1
        lMonE.state=1
        lMonR.state=0

      case 7:
        lMonM.state=1
        lMonO.state=1
        lMonN.state=1
        lMonS.state=1
        lMonT.state=1
        lMonE.state=1
        lMonR.state=1

    end select
  CheckLNMSequence
' SetLight lt1, "red", 0
' SetLight lt2, "red", 0
' SetLight lt3, "red", 0
  ResetTargetBankTimer003.enabled=false
End Sub

Sub StartTargetBank3
  If debugTargets Then debug.print "*****SUB:StartTargetBank3"
' SetLight lt1, "red", 2
' SetLight lt2, "red", 0
' SetLight lt3, "red", 0
End Sub

Sub CheckTargetBank3
  dim j
  If debugTargets Then debug.print "*****SUB:CheckTargetBank3"

  If ((ptarget7.transz=-45) and (ptarget8.transz=-45) and (ptarget9.transz=-45) and (ptarget10.transz=-45)) Then
    'vpmTimer.addTimer 2000, "ResetTargetBank3 '"
    Bank3Counter(CurrentPlayer)=Bank3Counter(CurrentPlayer)+1
    If Bank3Counter(CurrentPlayer)>7 then Bank3Counter(CurrentPlayer)=7
    for j = 1 to 4
      Bank3Dropped(CurrentPlayer,j)=0
    next
    LMonster.state=1
    lOutlaneR.state=1
    lOutlaneL.state=0
    lBlue.state=1
    if Bank3Counter(CurrentPlayer)>2 then lArrow.state=1
    CollectLNMSequence
    ResetTargetBankTimer003.enabled=true
  End If
End Sub

Sub CollectLNMSequence
  Dim LowestCount
  LowestCount=Bank1Counter(CurrentPlayer)
  if Bank2Counter(CurrentPlayer)<LowestCount then LowestCount=Bank2Counter(CurrentPlayer)
  if Bank3Counter(CurrentPlayer)<LowestCount then LowestCount=Bank3Counter(CurrentPlayer)

  If LowestCount=1 AND LNMCounter(CurrentPlayer)=0 then AddScore 50000:LNMCounter(CurrentPlayer)=LNMCounter(CurrentPlayer)+1:lLoch.state=0:lNess.state=0:LMonster.state=0:NessieTimer.enabled=true:NessieBGTimer.enabled=true:PlaySound "Nessy_Roar"
  If LowestCount=2 AND LNMCounter(CurrentPlayer)=1 then AddScore 50000:LNMCounter(CurrentPlayer)=LNMCounter(CurrentPlayer)+1:lLoch.state=0:lNess.state=0:LMonster.state=0:NessieTimer.enabled=true:NessieBGTimer.enabled=true:PlaySound "Nessy_Roar"
  If LowestCount=3 AND LNMCounter(CurrentPlayer)=2 then AddScore 100000:LNMCounter(CurrentPlayer)=LNMCounter(CurrentPlayer)+1:lLoch.state=0:lNess.state=0:LMonster.state=0:NessieTimer.enabled=true:NessieBGTimer.enabled=true:PlaySound "Nessy_Roar"
  If LowestCount=4 AND LNMCounter(CurrentPlayer)=3 then AddScore 100000:LNMCounter(CurrentPlayer)=LNMCounter(CurrentPlayer)+1:lLoch.state=0:lNess.state=0:LMonster.state=0:NessieTimer.enabled=true:NessieBGTimer.enabled=true:PlaySound "Nessy_Roar"
  If LowestCount=5 AND LNMCounter(CurrentPlayer)=4 then AwardExtraBall:LNMCounter(CurrentPlayer)=LNMCounter(CurrentPlayer)+1:lLoch.state=0:lNess.state=0:LMonster.state=0:NessieTimer.enabled=true:NessieBGTimer.enabled=true:PlaySound "Nessy_Roar"
  If LowestCount=6 AND LNMCounter(CurrentPlayer)=5 then AwardSpecial:LNMCounter(CurrentPlayer)=LNMCounter(CurrentPlayer)+1:lLoch.state=0:lNess.state=0:LMonster.state=0:NessieTimer.enabled=true:NessieBGTimer.enabled=true:PlaySound "Nessy_Roar"
  If LowestCount=7 AND LNMCounter(CurrentPlayer)=6 then AwardJackpot :LNMCounter(CurrentPlayer)=LNMCounter(CurrentPlayer)+1:lLoch.state=0:lNess.state=0:LMonster.state=0:NessieTimer.enabled=true:NessieBGTimer.enabled=true:PlaySound "Nessy_Roar"
end sub


Sub CheckLNMSequence
  Dim LowestCount,HighestCount
  LowestCount=Bank1Counter(CurrentPlayer)
  if Bank2Counter(CurrentPlayer)<LowestCount then LowestCount=Bank2Counter(CurrentPlayer)
  if Bank3Counter(CurrentPlayer)<LowestCount then LowestCount=Bank3Counter(CurrentPlayer)
  HighestCount=Bank1Counter(CurrentPlayer)
  if Bank2Counter(CurrentPlayer)<HighestCount then HighestCount=Bank2Counter(CurrentPlayer)
  if Bank3Counter(CurrentPlayer)<HighestCount then HighestCount=Bank3Counter(CurrentPlayer)
  l50k.state=2
  l100k.state=0
  lextraball.state=0
  lSpecial.state=0
  lJackpot.state=0
  If LowestCount>=2 then l50k.state=1:l100k.state=2
  If LowestCount>=4 then l100k.state=1:lextraball.state=2
  If LowestCount>=5 then lextraball.state=1:lSpecial.state=2
  If LowestCount>=6 then lSpecial.state=1:lJackpot.state=2
  lLoch.state=0:lNess.state=0:LMonster.state=0

  If Bank1Counter(CurrentPlayer)=HighestCount AND Bank2Counter(CurrentPlayer)=HighestCount AND Bank3Counter(CurrentPlayer)=HighestCount then Exit sub
  If Bank1Counter(CurrentPlayer)>LNMCounter(CurrentPlayer) then lLoch.state=1
  If Bank2Counter(CurrentPlayer)>LNMCounter(CurrentPlayer) then lNess.state=1
  If Bank3Counter(CurrentPlayer)>LNMCounter(CurrentPlayer) then LMonster.state=1
end sub

''*********************
' Section; Multiball kicker
''*********************
Dim bLockIsLit
Dim MultiballStartScore

Sub BallLock_Hit
  PlaySound "fxz_kicker_enter", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)

  If (bLockIsLit = False) OR (BallsOnPlayfield>1) Then 'Eject ball
    vpmTimer.addTimer 1500, "BallLockEject '"
  Else
    BallIsLocked=true
    lBallRelease.state=2
    lLock.state=1
    flash1.timerenabled=true
    BallLockedBlocker.IsDropped=0
    CreateNewBallAfterBallLock
  End If
End Sub

Sub BallLockEject
  Dim vary
  vary = (Rnd * 4)
  BallLock.kick 0, 50 + vary
  flash1.timerenabled=true
  PlaySoundAtVol SoundFXDOF("fx_kicker", 136, DOFPulse, DOFContactors), BallLock, 1

  PlaySound "LockEject"



End Sub

Sub StartMultiball
  If debugMultiball Then debug.print "*****SUB:StartMultiball"

  bMultiBallMode = True
  DOF 115, DOFPulse
  MultiballTargetReset
  StartMultiballMusic
  MultiballStartScore = Score(CurrentPlayer)

End Sub

Sub EndMultiball  'Multiball ending
  Dim mballtotal

  mballtotal = Score(CurrentPlayer) - MultiballStartScore


End Sub


'********
' Section; Bumper
'********
Sub Bumper1_Hit
  If debugGeneral Then debug.print "*****SUB:Bumper1_Hit"
    If Tilted Then Exit Sub
    PlaySoundAtVol SoundFXDOF("fxz_topbumper_hit", 107, DOFPulse, DOFContactors), Bumper1, 1
  Playsound "Pop"
  DOF 117, DOFPulse
  If BumperLight001.state=0 then
    AddScore BumperScore
  elseif BumperLight001.state=1 then
    AddScore 1000
  else
    AddScore 5000
  end if

    SetLastSwitchHit "Bumper1"
End Sub


Sub Bumper2_Hit
  If debugGeneral Then debug.print "*****SUB:Bumper2_Hit"
    If Tilted Then Exit Sub
    PlaySoundAtVol SoundFXDOF("fxz_topbumper_hit", 107, DOFPulse, DOFContactors), Bumper2, 1
  Playsound "Pop"
  DOF 118, DOFPulse
  If BumperLight002.state=0 then
    AddScore BumperScore
  elseif BumperLight002.state=1 then
    AddScore 1000
  else
    AddScore 5000
  end if

    SetLastSwitchHit "Bumper2"
End Sub

Sub Bumper3_Hit
  If debugGeneral Then debug.print "*****SUB:Bumper3_Hit"
    If Tilted Then Exit Sub
    PlaySoundAtVol SoundFXDOF("fxz_topbumper_hit", 107, DOFPulse, DOFContactors), Bumper3, 1
  Playsound "Pop"
  DOF 119, DOFPulse
  If BumperLight003.state=0 then
    AddScore BumperScore
  elseif BumperLight003.state=1 then
    AddScore 1000
  else
    AddScore 5000
  end if

    SetLastSwitchHit "Bumper3"
End Sub


''*********************
'' Section; Debug routines
''*********************
'BlockerWall for debug testing
Const BlockerWallsEnabled = 1
Dim BLState

If BlockerWallsEnabled = 1 Then
  BLW1.IsDropped=1:BLP1.Visible=0:BLR1.Visible=0: BLW2.IsDropped=1:BLP2.Visible=0:BLR2.Visible=0: BLW3.IsDropped=1:BLP3.Visible=0:BLR3.Visible=0
Else
  BLW1.IsDropped=1:BLP1.Visible=0:BLR1.Visible=0: BLW2.IsDropped=1:BLP2.Visible=0:BLR2.Visible=0: BLW3.IsDropped=1:BLP3.Visible=0:BLR3.Visible=0
End If

Sub BlockerWalls
  BLState = (BLState + 1) Mod 4
  debug.print "BlockerWalls"
  Playsound "fx_next"
  Select Case BLState:
    Case 0
      BLW1.IsDropped=1:BLP1.Visible=0:BLR1.Visible=0: BLW2.IsDropped=1:BLP2.Visible=0:BLR2.Visible=0: BLW3.IsDropped=1:BLP3.Visible=0:BLR3.Visible=0
    Case 1:
      BLW1.IsDropped=0:BLP1.Visible=1:BLR1.Visible=1: BLW2.IsDropped=0:BLP2.Visible=1:BLR2.Visible=1: BLW3.IsDropped=0:BLP3.Visible=1:BLR3.Visible=1
    Case 2:
      BLW1.IsDropped=0:BLP1.Visible=1:BLR1.Visible=1: BLW2.IsDropped=0:BLP2.Visible=1:BLR2.Visible=1: BLW3.IsDropped=1:BLP3.Visible=0:BLR3.Visible=0
    Case 3:
      BLW1.IsDropped=1:BLP1.Visible=0:BLR1.Visible=0: BLW2.IsDropped=1:BLP2.Visible=0:BLR2.Visible=0: BLW3.IsDropped=0:BLP3.Visible=1:BLR3.Visible=1
  End Select
End Sub


Sub SetBallsOnPlayfield (value)
  If debugGeneral Then debug.print "*****SUB:SetBallsOnPlayfield"

  If value < 0 Then
    value = 0
  End If

  BallsOnPlayfield = value
' Select Case BallsOnPlayfield
'   Case 1: 'Not multiball so resume based current X multiplier
'     Select Case BonusMultiplier(CurrentPlayer):
'       Case 1:
'         SetLight l2x, "white", 0
'         SetLight l3x, "white", 0
'         SetLight l4x, "white", 0
'
'       Case 2:
'         SetLight l2x, "white", 1
'         SetLight l3x, "white", 0
'         SetLight l4x, "white", 0
'
'       Case 3:
'         SetLight l2x, "white", 1
'         SetLight l3x, "white", 1
'         SetLight l4x, "white", 0
'
'       Case 4:
'         SetLight l2x, "white", 1
'         SetLight l3x, "white", 1
'         SetLight l4x, "white", 1
'
'       Case Else:
'         SetLight l2x, "white", 1
'         SetLight l3x, "white", 1
'         SetLight l4x, "white", 1
'     End Select
'
'
'   Case 2:
'     SetLight l2x, "red", 2
'     SetLight l3x, "red", 0
'     SetLight l4x, "red", 0
'   Case 3:
'     SetLight l2x, "red", 0
'     SetLight l3x, "red", 2
'     SetLight l4x, "red", 0
'   Case 4:
'     SetLight l2x, "red", 0
'     SetLight l3x, "red", 0
'     SetLight l4x, "red", 2
' End Select

End Sub

Sub AddBallsOnPlayfield (value)
  If debugGeneral Then debug.print "*****SUB:AddBallsOnPlayfield"

  Dim tmp
  tmp = BallsOnPlayfield + value

  SetBallsOnPlayfield tmp
End Sub


'Section; Save Player Data
'1.  Create player array
'2. Create Save routine
'3. Create Restore routine
'4.  Add to Master Save routine
'5.  Add to Master Restore routine

Sub InitializePlayerData
  If debugGeneral Then debug.print "*****SUB:InitializePlayerData"
' InitLaneSaveData
' InitMysteryAwardData
' InitGridData
' InitReactorData
' InitiReactorDestroyData
End Sub

Sub SavePlayerData
  If debugGeneral Then debug.print "*****SUB:SavePlayerData"
' SaveLaneSaveData
' SaveMysteryAwardData
' SaveGridData
' 'SaveReactorData - already saved
' SaveReactorDestroyData
End Sub

Sub RestorePlayerData
  If debugGeneral Then debug.print "*****SUB:RestorePlayerData"
' RestoreLaneSaveData
' RestoreMysteryAwardData
' RestoreGridData
' RestoreReactorData
' RestoreReactorDestroyData
End Sub

Sub CopyPlayerData (p1, p2)
  If debugGeneral Then debug.print "*****SUB:CopyPlayerData"
' CopyLaneSaveData p1, p2
' CopyMysteryAwardData p1, p2
' CopyGridData p1, p2
' CopyReactorData p1, p2
' CopyReactorDestroyData p1, p2
End Sub


Sub InitReactorData
  If debugGeneral Then debug.print "*****SUB:InitReactorData"
' Dim i
'    For i = 1 To MaxPlayers
'   ReactorState(i) = 0
'   ReactorLevel(i) = 1
'   ReactorTNAAchieved(i) = 0
'   ReactorReactorTotalReward(i) = 0
'   ReactorPercent(i) = -1
'   ReactorDestroyCount(i) = 0
'   ReactorValue(i) = ReactorValue1
'   ReactorValueMax(i) = ReactorValue1 * ReactorMaxMultiplier
'   ResetReactorLoopInserts
'    Next
End Sub

'Sub SaveReactorData
' Line 5071: Dim ReactorState(4)
' Line 5072: Dim ReactorLevel(4)
'        Dim ReactorTNAAchieved(4)
'        Dim ReactorReactorTotalReward(4)
' Line 5073: Dim ReactorPercent(4)
' Line 5074: Dim ReactorDestroyCount(4)
' Line 5075: Dim ReactorValue(4)
' Line 5076: Dim ReactorValueMax(4)
'End Sub

Sub CopyReactorData (p1, p2)
' ReactorState(p2) = ReactorState(p1)
' ReactorLevel(p2) = ReactorLevel(p1)
' ReactorTNAAchieved(p2) = ReactorTNAAchieved(p1)
' ReactorReactorTotalReward(p2) = ReactorReactorTotalReward(p1)
' ReactorPercent(p2) = ReactorPercent(p1)
' ReactorDestroyCount(p2) = ReactorDestroyCount(p1)
' ReactorValue(p2) = ReactorValue(p1)
' ReactorValueMax(p2) = ReactorValueMax(p1)
End Sub

Sub RestoreReactorData
' Dim tmpcolor
' Dim i
'
' 'Set all inserts off, light destroyed reactor green, then current reactor blinking red or green
' If ReactorState(CurrentPlayer) = 3 Then 'Critical
'   tmpcolor = "red"
' Else
'   tmpcolor = "green"
' End If
' For i = 0 to 8
'   SetLight aReactorLevelInserts(i), "red", 0
' Next
' For i = 0 to (ReactorLevel(CurrentPlayer) - 2)
'   If i >=0 Then SetLight aReactorLevelInserts(i), "red", 1
' Next
' i = ReactorLevel(CurrentPlayer) - 1
' If i >=0 AND i < 9 Then
'   SetLight aReactorLevelInserts(i), tmpcolor, 2
' End If
'
' 'Set table to default state
' ResetReactorLoopInserts
' lStart.State = 0
' lScoopEjectUpdate
' SetReactorInserts 0
' SetReactorPercent ReactorPercent(CurrentPlayer)
' fRones.TimerEnabled = False
' GIReactorStoppedImmediate
' StopReactorCriticalMusic
'
' 'Set up active items
' Select Case ReactorState(CurrentPlayer)
'   Case 0: 'Targeted
'
'   Case 1: 'Ready
'     SetReactorReady
'
'   Case 2: 'Started
'     SetReactorInserts 2
'     GIReactorStarted
'
'     If ReactorLevel(CurrentPlayer) > LastReactorBeforeDifficultyKicksIn Then  'Enable Reactor Percentage drop logic and only one loop
'       fRones.TimerInterval = ReactorPercentLossTime * 1000
'       fRones.TimerEnabled = True
'       StartReactorRightLoopInserts
'     Else 'ReactorLevel(CurrentPlayer) 1 and 2 and 3
'       StartReactorLoopInserts
'     End If
'
'   Case 3: 'Critical
'     SetReactorInserts 1
'     GiReactorCritical
''      ChangeGiImmediate "red", 1
''      GIGameImmediate 5, "red"
'     StartReactorCriticalMusic
'
'     'Need to restore Destroy targets count and inserts
'
' End Select
End Sub


Dim PlayeraDTGT1(4)
Dim PlayeraDTGT2(4)
Dim PlayeraDTGT3(4)
Dim PlayeraDTGT4(4)
Dim PlayeraDTGT5(4)
Dim PlayeraDTGT6(4)
Dim PlayeraDTGT7(4)

Sub InitiReactorDestroyData
  Dim i
    For i = 1 To MaxPlayers
    PlayeraDTGT1(i) = 0
    PlayeraDTGT2(i) = 0
    PlayeraDTGT3(i) = 0
    PlayeraDTGT4(i) = 0
    PlayeraDTGT5(i) = 0
    PlayeraDTGT6(i) = 0
    PlayeraDTGT7(i) = 0
    Next
End Sub

Sub SaveReactorDestroyData
  If ReactorState(CurrentPlayer) = 3 Then
    PlayeraDTGT1(CurrentPlayer) = lRAD1.State
    PlayeraDTGT2(CurrentPlayer) = lRAD2.State
    PlayeraDTGT3(CurrentPlayer) = lRAD3.State
    PlayeraDTGT4(CurrentPlayer) = lD1.State
    PlayeraDTGT5(CurrentPlayer) = lD2.State
    PlayeraDTGT6(CurrentPlayer) = lD3.State
    PlayeraDTGT7(CurrentPlayer) = lD4.State
  Else
    PlayeraDTGT1(CurrentPlayer) = 0
    PlayeraDTGT2(CurrentPlayer) = 0
    PlayeraDTGT3(CurrentPlayer) = 0
    PlayeraDTGT4(CurrentPlayer) = 0
    PlayeraDTGT5(CurrentPlayer) = 0
    PlayeraDTGT6(CurrentPlayer) = 0
    PlayeraDTGT7(CurrentPlayer) = 0
  End If
End Sub

Sub RestoreReactorDestroyData

  SetLight aDTGT(0), "white", PlayeraDTGT1(CurrentPlayer)
  SetLight aDTGT(1), "white", PlayeraDTGT2(CurrentPlayer)
  SetLight aDTGT(2), "white", PlayeraDTGT3(CurrentPlayer)
  SetLight aDTGT(3), "white", PlayeraDTGT4(CurrentPlayer)
  SetLight aDTGT(4), "white", PlayeraDTGT5(CurrentPlayer)
  SetLight aDTGT(5), "white", PlayeraDTGT6(CurrentPlayer)
  SetLight aDTGT(6), "white", PlayeraDTGT7(CurrentPlayer)


  SetLight aDFTGT(0), "white", PlayeraDTGT1(CurrentPlayer)
  SetLight aDFTGT(1), "white", PlayeraDTGT2(CurrentPlayer)
  SetLight aDFTGT(2), "white", PlayeraDTGT3(CurrentPlayer)
  SetLight aDFTGT(3), "white", PlayeraDTGT4(CurrentPlayer)
  SetLight aDFTGT(4), "white", PlayeraDTGT5(CurrentPlayer)
  If PlayeraDTGT6(CurrentPlayer) = 2 Then
    SetLight aDFTGT(5), "white", PlayeraDTGT6(CurrentPlayer)  'bumper shared with dtgt6 and 7
  Else
    SetLight aDFTGT(5), "white", PlayeraDTGT7(CurrentPlayer)  'bumper shared with dtgt6 and 7
  End If
End Sub

Sub CopyReactorDestroyData (p1, p2)
  PlayeraDTGT1(p2) = PlayeraDTGT1(p1)
  PlayeraDTGT2(p2) = PlayeraDTGT2(p1)
  PlayeraDTGT3(p2) = PlayeraDTGT3(p1)
  PlayeraDTGT4(p2) = PlayeraDTGT4(p1)
  PlayeraDTGT5(p2) = PlayeraDTGT5(p1)
  PlayeraDTGT6(p2) = PlayeraDTGT6(p1)
  PlayeraDTGT7(p2) = PlayeraDTGT7(p1)

End Sub


'==================================
'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart, FlipperEnd, LR, PartialFlipCoef
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
  Public Property Let EndPoint(aInput) : if IsObject(aInput) then FlipperEnd = aInput.x else FlipperEnd = aInput : end if : End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
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
        if DebugOn then StickL.visible = True : StickL.x = balldata(x).x    'debug TODO
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 30 Then
      PartialFlipCoef = 0
    End If
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
      dim teststr : teststr = "Cutoff"
      tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
      if tmp < 0.1 then 'if real ball position is behind flipper, exit Sub to prevent stucks  'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
        if DebugOn then TestStr = "real pos < 0.1 ( " & round(tmp,2) & ")" : tbpl.text = Teststr
        'RemoveBall aBall
        'Exit Sub
      end if

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        if DebugOn then teststr = "y velocity: " & round(aBall.vely, 3) & "exit sub" : tbpl.text = teststr
        RemoveBall aBall
        exit Sub
      end if
      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          'TB.TEXT = balldata(x).id & " " & BALLDATA(X).X & VBNEWLINE & FLIPPERSTART & " " & FLIPPEREND
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        if DebugOn then set tmp = new spoofball : tmp.data = aBall : End If
        if IsEmpty(BallData(idx).id) and aBall.VelY < -12 then 'if tip hit with no collected data, do vel correction anyway
          if PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1) > 1.1 then 'adjust plz
            VelCoef = LinearEnvelope(5, VelocityIn, VelocityOut)
            if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
            if Enabled then aBall.Velx = aBall.Velx*VelCoef'VelCoef
            if Enabled then aBall.Vely = aBall.Vely*VelCoef'VelCoef
            if DebugOn then teststr = "tip protection" & vbnewline & "velcoef: " & round(velcoef,3) & vbnewline & round(PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1),3) & vbnewline
            'debug.print teststr
          end if
        Else
     :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
          if Enabled then aBall.Velx = aBall.Velx*VelCoef
          if Enabled then aBall.Vely = aBall.Vely*VelCoef
        end if
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      'debug
      if DebugOn then
        TestStr = teststr & "%pos:" & round(BallPos,2)
        if IsEmpty(PolarityOut(0) ) then
          teststr = teststr & vbnewline & "(Polarity Disabled)" & vbnewline
        else
          teststr = teststr & "+" & round(1 *(AddX*ycoef*PartialFlipcoef),3)
          if BallPos >= PolarityOut(uBound(PolarityOut) ) then teststr = teststr & "(MAX)" & vbnewline else teststr = teststr & vbnewline end if
          if Ycoef < 1 then teststr = teststr &  "ycoef: " & ycoef & vbnewline
          if PartialFlipcoef < 1 then teststr = teststr & "PartialFlipcoef: " & round(PartialFlipcoef,4) & vbnewline
        end if

        teststr = teststr & vbnewline & "Vel: " & round(BallSpeed(tmp),2) & " -> " & round(ballspeed(aBall),2) & vbnewline
        teststr = teststr & "%" & round(ballspeed(aBall) / BallSpeed(tmp),2)
        tbpl.text = TestSTR
      end if
    Else
      'if DebugOn then tbpl.text = "td" & timedelay
    End If
    RemoveBall aBall
  End Sub
End Class

'================================
'Helper Functions


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
'   dim x, aCount : aCount = 0
'   redim a(uBound(aArray) )
'   for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
'     if not IsEmpty(aArray(x) ) Then
'       if IsObject(aArray(x)) then
'         Set a(aCount) = aArray(x)
'       Else
'         a(aCount) = aArray(x)
'       End If
'       aCount = aCount + 1
'     End If
'   Next
'   if offset < 0 then offset = 0
'   redim aArray(aCount-1+offset) 'Resize original array
'   for x = 0 to aCount-1   'set objects back into original array
'     if IsObject(a(x)) then
'       Set aArray(x) = a(x)
'     Else
'       aArray(x) = a(x)
'     End If
'   Next
' End Sub


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub ShuffleArrays(aArray1, aArray2, offset)
'   ShuffleArray aArray1, offset
'   ShuffleArray aArray2, offset
' End Sub


Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

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


Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'Clamp if on the boundry lines
  'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
  'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function



dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

If FlipperPhysicsMode = 2 Then InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    x.TimeDelay = 44
  Next

  '"Polarity" Profile
  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.368, -4
  AddPt "Polarity", 2, 0.451, -3.7
  AddPt "Polarity", 3, 0.493, -3.88
  AddPt "Polarity", 4, 0.65, -2.3
  AddPt "Polarity", 5, 0.71, -2
  AddPt "Polarity", 6, 0.785,-1.8
  AddPt "Polarity", 7, 1.18, -1
  AddPt "Polarity", 8, 1.2, 0


  '"Velocity" Profile
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() :   If FlipperPhysicsMode = 2 Then LF.Addball activeball End If: End Sub
Sub TriggerLF_UnHit() : If FlipperPhysicsMode = 2 Then LF.PolarityCorrect activeball End If: End Sub
Sub TriggerRF_Hit() :   If FlipperPhysicsMode = 2 Then RF.Addball activeball End If: End Sub
Sub TriggerRF_UnHit() : If FlipperPhysicsMode = 2 Then RF.PolarityCorrect activeball End If: End Sub


Sub lScoopEjectUpdate
  if ((lmystery.state = 2) Or (lStart.state = 2)) Then
    lScoopEject.state = 2
  Else
    lScoopEject.state = 0
  End If
End Sub

'****************************************************
' Glowball routine
'****************************************************
Const GlowballEnabled = 0
'XXXXX NEED TO UNCOMMENT IF USING Set Glowing(0) = Glowball1 : Set Glowing(1) = Glowball2 : Set Glowing(2) = Glowball3 : Set Glowing(3) = Glowball4


Dim ChooseBats
' *** 0=default flipper, 1=primitive flipper, 2=glow green, 3=glow blue, 4=glow orange ****
ChooseBats = 1

'*********** Glowball
Dim GlowBall, CustomBulbIntensity(10)
Dim  GBred(10)
Dim GBgreen(10), GBblue(10)
Dim CustomBallImage(10), CustomBallLogoMode(10), CustomBallDecal(10), CustomBallGlow(10)

' default Ball
CustomBallGlow(0) =     False
CustomBallImage(0) =    "TTMMball"
CustomBallLogoMode(0) =   False
CustomBallDecal(0) =    "scratches"
CustomBulbIntensity(0) =  0.01
GBred(0) = 0 : GBgreen(0) = 0 : GBblue(0) = 0

' white GlowBall
CustomBallGlow(1) =     True
CustomBallImage(1) =    "white"
CustomBallLogoMode(1) =   True
CustomBallDecal(1) =    ""
CustomBulbIntensity(1) =  0
GBred(1) = 255 : GBgreen(1) = 255 : GBblue(1) = 255

' Magma GlowBall
CustomBallGlow(2) =     True
CustomBallImage(2) =    "ballblack"
CustomBallLogoMode(2) =   True
CustomBallDecal(2) =    "magma6"
CustomBulbIntensity(2) =  0
GBred(2) = 255 : GBgreen(2) = 20 : GBblue(2) = 20

' Blue ball
CustomBallGlow(3) =     True
CustomBallImage(3) =    "blueball2"
CustomBallLogoMode(3) =   False
CustomBallDecal(3) =    ""
CustomBulbIntensity(3) =  0
GBred(3) = 30 : GBgreen(3)  = 40 : GBblue(3) = 200

' HDR ball
CustomBallGlow(4) =     False
CustomBallImage(4) =    "ball_HDR"
CustomBallLogoMode(4) =   False
CustomBallDecal(4) =    "JPBall-Scratches"
CustomBulbIntensity(4) =  0.01
GBred(4) = 0 : GBgreen(4) = 0 : GBblue(4) = 0

' Earth
CustomBallGlow(5) =     True
CustomBallImage(5) =    "ballblack"
CustomBallLogoMode(5) =   True
CustomBallDecal(5) =    "earth"
CustomBulbIntensity(5) =  0
GBred(5) = 100 : GBgreen(5) = 100 : GBblue(5) = 100

' green GlowBall
CustomBallGlow(6) =     True
CustomBallImage(6) =    "glowball green"
CustomBallLogoMode(6) =   True
CustomBallDecal(6) =    ""
CustomBulbIntensity(6) =  0
GBred(6) = 100 : GBgreen(6) = 255 : GBblue(6) = 100

' blue GlowBall
CustomBallGlow(7) =     True
CustomBallImage(7) =    "glowball blue"
CustomBallLogoMode(7) =   True
CustomBallDecal(7) =    ""
CustomBulbIntensity(7) =  0
GBred(7) = 50 : GBgreen(7)  = 50 : GBblue(7) = 255
'GBred(7) = 100 : GBgreen(7)  = 100 : GBblue(7) = 255

' red GlowBall
CustomBallGlow(8) =     True
CustomBallImage(8) =    "glowball orange"
CustomBallLogoMode(8) =   True
CustomBallDecal(8) =    ""
CustomBulbIntensity(8) =  0
GBred(8) = 255 : GBgreen(8) = 0 : GBblue(8) = 000
'GBred(8) = 255 : GBgreen(8)  = 255 : GBblue(8) = 100  'orange

' shiny Ball
CustomBallGlow(9) =     False
CustomBallImage(9) =    "pinball3"
CustomBallLogoMode(9) =   False
CustomBallDecal(9) =    "JPBall-Scratches"
CustomBulbIntensity(9) =  0.01
GBred(9) = 0 : GBgreen(9) = 0 : GBblue(9) = 0

' *** prepare the variable with references to three lights for glow ball ***
Dim Glowing(10)


'*** change ball appearance ***
Sub ChangeBall(ballnr)
  Dim BOT, ii, col
  table1.BallDecalMode = CustomBallLogoMode(ballnr)
  table1.BallFrontDecal = CustomBallDecal(ballnr)
  table1.DefaultBulbIntensityScale = CustomBulbIntensity(ballnr)
  table1.BallImage = CustomBallImage(ballnr)
  GlowBall = CustomBallGlow(ballnr)
  If GlowballEnabled > 0 Then
    For ii = 0 to 3
      col = RGB(GBred(ballnr), GBgreen(ballnr), GBblue(ballnr))
      Glowing(ii).color = col : Glowing(ii).colorfull = col
    Next
  End If
End Sub



' *** Ball Shadow code / Glow Ball code / Primitive Flipper Update ***

Dim BallShadowArray
BallShadowArray = Array (BallShadow1, BallShadow2, BallShadow3)
Const anglecompensate = 15

Sub GraphicsTimer_Timer()
' Dim BOT, b
'    BOT = GetBalls
'
' ' switch off glowlight for removed Balls
' IF GlowBall Then
'   For b = UBound(BOT) + 1 to 3
'     If GlowBall and Glowing(b).state = 1 Then Glowing(b).state = 0 End If
'   Next
' End If
'
'    For b = 0 to UBound(BOT)
'   ' *** move ball shadow for max 3 balls ***
'   If b < 3 Then
'     If BOT(b).X < table1.Width/2 Then
'       BallShadowArray(b).X = ((BOT(b).X) - (50/6) + ((BOT(b).X - (table1.Width/2))/7)) + 10
'     Else
'       BallShadowArray(b).X = ((BOT(b).X) + (50/6) + ((BOT(b).X - (table1.Width/2))/7)) - 10
'     End If
'     BallShadowArray(b).Y = BOT(b).Y + 20 : BallShadowArray(b).Z = 1
'     If BOT(b).Z > 20 Then BallShadowArray(b).visible = 1 Else BallShadowArray(b).visible = 0 End If
'   End If
'   ' *** move glowball light for max 3 balls ***
'   If GlowBall and b < 4 Then
'     If Glowing(b).state = 0 Then Glowing(b).state = 1 end if
'     Glowing(b).BulbHaloHeight = BOT(b).z + 51
'     Glowing(b).x = BOT(b).x : Glowing(b).y = BOT(b).y + anglecompensate
'   End If
' Next
' If ChooseBats = 1 Then
'   ' *** move primitive bats ***
'   batleft.objrotz = LeftFlipper.CurrentAngle + 1
'   batleftshadow.objrotz = batleft.objrotz
'   batright.objrotz = RightFlipper.CurrentAngle - 1
'   batrightshadow.objrotz  = batright.objrotz
' Else
'   If ChooseBats > 1 Then
'     ' *** move glowbats ***
'     GlowBatLightLeft.y = 1720 - 121 + LeftFlipper.CurrentAngle
'     glowbatleft.objrotz = LeftFlipper.CurrentAngle
'     GlowBatLightRight.y =1720 - 121 - RightFlipper.CurrentAngle
'     glowbatright.objrotz = RightFlipper.CurrentAngle
'   End If
' End If
  batleftshadow.RotZ = LeftFlipper.currentangle
  batrightshadow.RotZ = RightFlipper.currentangle
End Sub

' ****************************************************************
' Ultradmd support
' ****************************************************************
Dim UltraDMD
Sub LoadUltraDMD
    Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    UltraDMD.Init
  uDMDScoreTimer.Interval = UltraDMDUpdateTime
  uDMDScoreTimer.Enabled = 1
  uDMDScoreUpdate
End Sub

Sub uDMDScoreTimer_Timer
  uDMDScoreUpdate
End Sub

Sub uDMDScoreUpdate
  If UseUltraDMD = 1 Then
    If TestMode = 0 Then
      UltraDMD.DisplayScoreboard00 PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), " ", "Ball " & Balls
    Else
      UltraDMD.DisplayScoreboard00 PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "DEBUG MODE" & ":BP:" & BallsOnPlayfield, "Ball " & Balls
    End If
  ElseIf UseUltraDMD = 2 Then
    If TestMode = 0 Then
      UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Reactor Val:" & ReactorValue(CurrentPlayer), "Ball " & Balls
    Else
      UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "RV:" & ReactorValue(CurrentPlayer) & ":BP:" & BallsOnPlayfield, "Ball " & Balls
    End If
  End If
End Sub


Sub UDMD (toptext, bottomtext, utime)
  If UseUltraDMD > 0 Then UltraDMD.DisplayScene00Ex "", toptext, 8, 14, bottomtext, 8,14, 14, utime, 14
End Sub





'*****************************
' AUTO TESTING
' by:NailBuster
' Global variable "AutoQA" below will switch all this on/off during testing.
'
'*****************************
' NailBusters AutoQA Code and triggers..
' this to do for ROM based:  timeout on keydown.  if 30 seconds, then assume game is over and you add coins/start game key.
' add a timer called AutoQAStartGame.  you can run every 10000 interval.

Dim AutoQA
Dim QACoinStartSec:QACoinStartSec=60   'timeout seconds for AutoCoinStartSec
Dim QANumberOfCoins:QANumberOfCoins=3 + AutoQa 'number of coins to add for each start
Dim QASecondsDiff

Dim QALastFlipperTime:QALastFlipperTime=Now()
Dim AutoFlipperLeft:AutoFlipperLeft=false
Dim AutoFlipperRight:AutoFlipperRight=false

AutoQa = 0      '0 = off, 1, 2,3,4 = 1 or 2 or 3 or 4 player test.   Main QA Testing FLAG setting to false will disable all this stuff.
If AutoQa > 0 Then AutoQAStartGame.Enabled = True


Sub AutoQAStartGame_Timer()                 'this is a timeout when sitting in attract with no flipper presses for 60 seconds, then add coins and start game.
 if AutoQA=0 Then Exit Sub

 QASecondsDiff = DateDiff("s",QALastFlipperTime,NOW())

 if QASecondsDiff>QACoinStartSec Then

    'simulate quarters and start game keys
    Dim fx : fx=0
    Dim keydelay : keydelay=100
  Do While fx<QANumberOfCoins
    vpmtimer.addtimer keydelay,"Table1_KeyDown(keyInsertCoin1) '"
        vpmtimer.addtimer keydelay+200,"Table1_KeyUp(keyInsertCoin1) '"
        keydelay=keydelay+500
    fx=fx+1
  Loop

  fx=0
  Do While fx<AutoQa
    vpmtimer.addtimer keydelay,"Table1_KeyDown(StartGameKey) '"
    vpmtimer.addtimer keydelay+200,"Table1_KeyUp(StartGameKey) '"
        keydelay=keydelay+500
    fx=fx+1
  Loop


    QALastFlipperTime=Now()
  AutoFlipperLeft=false
    AutoFlipperRight=false
 End if

 if QASecondsDiff>30 Then   'safety of stuck up flipers.
   AutoFlipperLeft=false
   AutoFlipperRight=false
  End if
End Sub


Sub TriggerAutoPlunger_Hit()          'add a trigger in front of plunger.  adjust the delay timings if needed.
    if AutoQA=0 Then Exit Sub
  vpmtimer.addtimer 10,"Table1_KeyDown(PlungerKey) '"
    vpmtimer.addtimer 900+RND(400),"Table1_KeyUp(PlungerKey) '"
End Sub



Sub FlipperUP(which)  'which=1 left 2 right
QALastFlipperTime=Now()
if which=1 Then
   Table1_KeyDown(LeftFlipperKey)
   vpmtimer.addtimer 200+Rnd(200),"Table1_KeyUP(LeftFlipperKey):AutoFlipperLeft=false  '"
Else
   Table1_KeyDown(RightFlipperKey)
   vpmtimer.addtimer 200+Rnd(200),"Table1_KeyUP(RightFlipperKey):AutoFlipperRight=false  '"
end If

End Sub



Sub TriggerLeftAuto_Hit()
  if AutoQA>0 And AutoFlipperLeft=false then vpmtimer.addtimer 20+Rnd(20),"FlipperUP(1) '"
    AutoFlipperLeft=true
End Sub

Sub TriggerRightAuto_Hit()
  if AutoQA>0 and AutoFlipperRight=false then vpmtimer.addtimer 20+Rnd(20),"FlipperUP(2) '"
    AutoFlipperRight=true
End Sub

Sub TriggerLeftAuto2_Hit()
  TriggerLeftAuto_Hit()
End Sub

Sub TriggerRightAuto2_Hit()
  TriggerRightAuto_Hit()
End Sub




'*****************************************************



'================================================================
'DOF Events -   "-- Means DOF event is used
'================================================================
'101 - Left Flipper
'102 - Right Flipper
'103 - Left SlingShot Solenoid
'104 - Left SlingShot Flasher
'105 - Right SlingShot Solenoid
'106 - Right SlingShot Flasher
'107 - Bumper 1 Solenoid
'108 - Bumper 2 Solenoid
'109 - Bumper 3 Solenoid
'108 - RAD Left Standup Target bank
'109 - Grid Targetx,y,z standup target
'110 - Destroy Left standup
'111 - Destroy Right Standup
'112 - Left Spinner Flasher
'113 - Right Spinner Flasher
'114 - Reactor Stand Up targets
'115 -- Multiball - Consider Dof On and Off when MB ends
'116 - Drain
'117 - Bumper 1 Flasher
'118 - Bumper 2 Flasher
'119 - Bumper 3 Flasher
'121 - Ball Trough
'122 - Left Scoop Solenoid and Right Scoop Solenoid - Consider separating
'123 - Strobe used for Autoplunge , Extra Game, Left Scoop, Right Scoop, Single Jackpot, Double Jackpot
'125 - Autoplunge 'Solenoid
'129 - Knocker
'132 - Triple Jackpot Strobe
'133 - Super Jackpot Strobe
'136 - Drop Target Upper Solenoid
'137 - Drop Target MIddle Solenoid
'138 - Drop Target Lower Solenoid
'140 - Credits/Free Play 'Start Button
'Undercab E145 White/E146 Blue/E147 Red/E148 Green/E149 Purple
  '--Const DOFuyellow = 144
  '--Const DOFuwhite = 145
  '--Const DOFublue  = 146
  '--Const DOFured   = 147
  '--Const DOFugreen = 148
  '--Const DOFupurple= 149
'150 -- Top Lane 1
'151 -- Top Lane 2
'152 -- Top Lane 3
'153 -- Top Lane 4
'160 - Combo Loop
'161 - Tilt warning
'162 - Tilted
'163 - Shoot again
'164 - Player Bonus Count (8 seconds)
'165 - Ball Saved
'166 -- Bonus Multiplier Awarded (1.5 seconds)
'167 - Extra Ball Earned
'168 - Skill Shot
'169 - Handsfree Skill Shot
'170 - Lane Save Earned
'171 - Super Spinner Awarded
'172 - Locks are Lit
'173 -- Ball Lock 1
'173 -- Ball Lock 2 (using 173 for same Effect)
'175 - Reactor Grid Jackpot
'176 - Reactor Ready
'177 - Reactor Started
'178 -- Reactor Critical
'179 -- Reactor Destroyed
'180 -- Total Annihilation Achieved
'181 - Reactor Value Maxed
'182 - Mystery is Lit
'183 - Mystery Awarded (4 seconds)
'184 -- Grid insert Left Green
'185 -- Grid insert Middle Green
'186 -- Grid insert Right Green
'187 -- Grid insert Left Purple
'188 -- Grid insert Middle Purple
'189 -- Grid insert Right Purple
'191 - Target1
'192 - Target2
'193 - Target3
'194 - Target4
'195 - Target5
'196 - Target6
'197 - Target7
'198 - Target8
'199 - Target9
'200 - Target10
'================================================================
'DOF Events
'================================================================

'***********Rotate Spinner
Dim Angle

Sub SpinnerTimer_Timer
  Angle = (sin (spinner1.CurrentAngle-180))
  SpinnerRod.TransX = sin( (spinner1.CurrentAngle+180) * (2*PI/360)) * 12
  SpinnerRod.TransZ = sin( (spinner1.CurrentAngle- 90) * (2*PI/360)) * 3.5
  Dim SpinnerRadius: SpinnerRadius=7
End Sub

'*** PI returns the value for PI
Function PI()
  PI = 4*Atn(1)
End Function

'******************************************************
'   RUBBER POST AND SLEEVE DAMPENERS
'******************************************************
'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.

Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub

dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"

'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub

End Class

'******************************************************
'   TRACK ALL BALL VELOCITIES
'     FOR RUBBER DAMPENER AND DROP TARGETS
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

Sub RDampen_Timer()
  Cor.Update
End Sub

'******************************************************
'   FLIPPER POLARITY, DAMPENER, AND DROP TARGET
'       SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
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
Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

' Used for drop targets
Function Atn2(dy, dx)
  dim pi
  pi = 4*Atn(1)

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
'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

' Used for drop targets
Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

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


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

' using table width and height in script slows down the performance
dim tablewidth: tablewidth = Table1.width
dim tableheight: tableheight = Table1.height

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
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound ("Nudge_1"), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
    Case 2 : PlaySound ("Nudge_2"), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
    Case 3 : PlaySound ("Nudge_3"), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
  End Select
End Sub

Sub SoundNudgeRight()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound ("Nudge_1"), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
    Case 2 : PlaySound ("Nudge_2"), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
    Case 3 : PlaySound ("Nudge_3"), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
  End Select
End Sub

Sub SoundNudgeCenter()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound ("Nudge_1"), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
    Case 2 : PlaySound ("Nudge_2"), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
    Case 3 : PlaySound ("Nudge_3"), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
  End Select
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
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtLevelStatic ("Drain_1"), DrainSoundLevel, drainswitch
    Case 2 : PlaySoundAtLevelStatic ("Drain_2"), DrainSoundLevel, drainswitch
    Case 3 : PlaySoundAtLevelStatic ("Drain_3"), DrainSoundLevel, drainswitch
    Case 4 : PlaySoundAtLevelStatic ("Drain_4"), DrainSoundLevel, drainswitch
    Case 5 : PlaySoundAtLevelStatic ("Drain_5"), DrainSoundLevel, drainswitch
    Case 6 : PlaySoundAtLevelStatic ("Drain_6"), DrainSoundLevel, drainswitch
    Case 7 : PlaySoundAtLevelStatic ("Drain_7"), DrainSoundLevel, drainswitch
    Case 8 : PlaySoundAtLevelStatic ("Drain_8"), DrainSoundLevel, drainswitch
    Case 9 : PlaySoundAtLevelStatic ("Drain_9"), DrainSoundLevel, drainswitch
    Case 10 : PlaySoundAtLevelStatic ("Drain_10"), DrainSoundLevel, drainswitch
    Case 11 : PlaySoundAtLevelStatic ("Drain_11"), DrainSoundLevel, drainswitch
  End Select
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBallRelease(drainswitch)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("BallRelease1",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 2 : PlaySoundAtLevelStatic SoundFX("BallRelease2",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 3 : PlaySoundAtLevelStatic SoundFX("BallRelease3",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 4 : PlaySoundAtLevelStatic SoundFX("BallRelease4",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 5 : PlaySoundAtLevelStatic SoundFX("BallRelease5",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 6 : PlaySoundAtLevelStatic SoundFX("BallRelease6",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 7 : PlaySoundAtLevelStatic SoundFX("BallRelease7",DOFContactors), BallReleaseSoundLevel, drainswitch
  End Select
End Sub



'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Sling_L1",DOFContactors), SlingshotSoundLevel, Sling
    Case 2 : PlaySoundAtLevelStatic SoundFX("Sling_L2",DOFContactors), SlingshotSoundLevel, Sling
    Case 3 : PlaySoundAtLevelStatic SoundFX("Sling_L3",DOFContactors), SlingshotSoundLevel, Sling
    Case 4 : PlaySoundAtLevelStatic SoundFX("Sling_L4",DOFContactors), SlingshotSoundLevel, Sling
    Case 5 : PlaySoundAtLevelStatic SoundFX("Sling_L5",DOFContactors), SlingshotSoundLevel, Sling
    Case 6 : PlaySoundAtLevelStatic SoundFX("Sling_L6",DOFContactors), SlingshotSoundLevel, Sling
    Case 7 : PlaySoundAtLevelStatic SoundFX("Sling_L7",DOFContactors), SlingshotSoundLevel, Sling
    Case 8 : PlaySoundAtLevelStatic SoundFX("Sling_L8",DOFContactors), SlingshotSoundLevel, Sling
    Case 9 : PlaySoundAtLevelStatic SoundFX("Sling_L9",DOFContactors), SlingshotSoundLevel, Sling
    Case 10 : PlaySoundAtLevelStatic SoundFX("Sling_L10",DOFContactors), SlingshotSoundLevel, Sling
  End Select
End Sub

Sub RandomSoundSlingshotRight(sling)
  Select Case Int(Rnd*8)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Sling_R1",DOFContactors), SlingshotSoundLevel, Sling
    Case 2 : PlaySoundAtLevelStatic SoundFX("Sling_R2",DOFContactors), SlingshotSoundLevel, Sling
    Case 3 : PlaySoundAtLevelStatic SoundFX("Sling_R3",DOFContactors), SlingshotSoundLevel, Sling
    Case 4 : PlaySoundAtLevelStatic SoundFX("Sling_R4",DOFContactors), SlingshotSoundLevel, Sling
    Case 5 : PlaySoundAtLevelStatic SoundFX("Sling_R5",DOFContactors), SlingshotSoundLevel, Sling
    Case 6 : PlaySoundAtLevelStatic SoundFX("Sling_R6",DOFContactors), SlingshotSoundLevel, Sling
    Case 7 : PlaySoundAtLevelStatic SoundFX("Sling_R7",DOFContactors), SlingshotSoundLevel, Sling
    Case 8 : PlaySoundAtLevelStatic SoundFX("Sling_R8",DOFContactors), SlingshotSoundLevel, Sling
  End Select
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
  End Select
End Sub

Sub RandomSoundBumperMiddle(bump)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
  End Select
End Sub

Sub RandomSoundBumperBottom(Bump)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
  End Select
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
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_L01",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_L02",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_L07",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_L08",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_L09",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_L10",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_L12",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 8 : PlaySoundAtLevelStatic SoundFX("Flipper_L14",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 9 : PlaySoundAtLevelStatic SoundFX("Flipper_L18",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 10 : PlaySoundAtLevelStatic SoundFX("Flipper_L20",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 11 : PlaySoundAtLevelStatic SoundFX("Flipper_L26",DOFFlippers), FlipperLeftHitParm, Flipper
  End Select
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_R01",DOFFlippers), FlipperRightHitParm, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_R02",DOFFlippers), FlipperRightHitParm, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_R03",DOFFlippers), FlipperRightHitParm, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_R04",DOFFlippers), FlipperRightHitParm, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_R05",DOFFlippers), FlipperRightHitParm, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_R06",DOFFlippers), FlipperRightHitParm, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_R07",DOFFlippers), FlipperRightHitParm, Flipper
    Case 8 : PlaySoundAtLevelStatic SoundFX("Flipper_R08",DOFFlippers), FlipperRightHitParm, Flipper
    Case 9 : PlaySoundAtLevelStatic SoundFX("Flipper_R09",DOFFlippers), FlipperRightHitParm, Flipper
    Case 10 : PlaySoundAtLevelStatic SoundFX("Flipper_R10",DOFFlippers), FlipperRightHitParm, Flipper
    Case 11 : PlaySoundAtLevelStatic SoundFX("Flipper_R11",DOFFlippers), FlipperRightHitParm, Flipper
  End Select
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L01",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L02",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L03",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
  End Select
End Sub

Sub RandomSoundReflipUpRight(flipper)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R01",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R02",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R03",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
  End Select
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_1",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_2",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_3",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_4",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_5",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_6",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_7",DOFFlippers), FlipperDownSoundLevel, Flipper
  End Select
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  Select Case Int(Rnd*8)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_1",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_2",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_3",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_4",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_5",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_6",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_7",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 8 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_8",DOFFlippers), FlipperDownSoundLevel, Flipper
  End Select
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
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_1"), parm  * RubberFlipperSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_2"), parm  * RubberFlipperSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_3"), parm  * RubberFlipperSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_4"), parm  * RubberFlipperSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_5"), parm  * RubberFlipperSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_6"), parm  * RubberFlipperSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_7"), parm  * RubberFlipperSoundFactor
  End Select
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rollover_1"), RolloverSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("Rollover_2"), RolloverSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("Rollover_3"), RolloverSoundLevel
    Case 4 : PlaySoundAtLevelActiveBall ("Rollover_4"), RolloverSoundLevel
  End Select
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
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_1"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_2"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_5"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_6"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_7"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_8"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_9"), Vol(ActiveBall) * RubberWeakSoundFactor
  End Select
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
  Select Case Int(Rnd*13)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Metal_Touch_1"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Metal_Touch_2"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Metal_Touch_3"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Metal_Touch_4"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall ("Metal_Touch_5"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall ("Metal_Touch_6"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall ("Metal_Touch_7"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall ("Metal_Touch_8"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall ("Metal_Touch_9"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall ("Metal_Touch_10"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 11 : PlaySoundAtLevelActiveBall ("Metal_Touch_11"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 12 : PlaySoundAtLevelActiveBall ("Metal_Touch_12"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 13 : PlaySoundAtLevelActiveBall ("Metal_Touch_13"), Vol(ActiveBall) * MetalImpactSoundFactor
  End Select
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
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_2"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
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
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_1"), BottomArchBallGuideSoundFactor * 0.25
    Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_2"), BottomArchBallGuideSoundFactor * 0.25
    Case 3 : PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_3"), BottomArchBallGuideSoundFactor * 0.25
  End Select
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
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Medium_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_2"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*7)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Soft_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Soft_2"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall ("Apron_Soft_3"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 4 : PlaySoundAtLevelActiveBall ("Apron_Soft_4"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 5 : PlaySoundAtLevelActiveBall ("Apron_Soft_5"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 6 : PlaySoundAtLevelActiveBall ("Apron_Soft_6"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 7 : PlaySoundAtLevelActiveBall ("Apron_Soft_7"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_5",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_6",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_7",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_8",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
  End Select
End Sub

Sub RandomSoundTargetHitWeak()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_1",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_2",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_3",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_4",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
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
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_3"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_4"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_6"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
  End Select
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
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySoundAtLevelStatic ("Gate_FastTrigger_1"), GateSoundLevel, Activeball
    Case 2 : PlaySoundAtLevelStatic ("Gate_FastTrigger_2"), GateSoundLevel, Activeball
  End Select
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
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Arch_L1"), Vol(ActiveBall) * ArchSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Arch_L2"), Vol(ActiveBall) * ArchSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Arch_L3"), Vol(ActiveBall) * ArchSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Arch_L4"), Vol(ActiveBall) * ArchSoundFactor
  End Select
End Sub

Sub RandomSoundRightArch()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Arch_R1"), Vol(ActiveBall) * ArchSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Arch_R2"), Vol(ActiveBall) * ArchSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Arch_R3"), Vol(ActiveBall) * ArchSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Arch_R4"), Vol(ActiveBall) * ArchSoundFactor
  End Select
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
  Select Case Int(Rnd*2)+1
    Case 1: PlaySoundAtLevelStatic ("Saucer_Enter_1"), SaucerLockSoundLevel, Activeball
    Case 2: PlaySoundAtLevelStatic ("Saucer_Enter_2"), SaucerLockSoundLevel, Activeball
  End Select
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


Sub lArrow_Init()

End Sub

Sub NessieTimer_timer
  If NessieFrames<10 then
    Nessie.image="nessie000"&NessieFrames
  else
    Nessie.image="nessie00"&NessieFrames
  end if
  NessieFrames=NessieFrames+1
  if NessieFrames>50 then NessieFrames=1:NessieTimer.enabled=false:Nessie.image="nessie0001"
end sub

Sub BumperFlashingTimer_timer
  BumperFlashingTimer.enabled=false
  BumperLight001.state=0
  BumperLight002.state=0
  BumperLight003.state=0
end sub

Sub flash1_timer
  FlasherCounter1=FlasherCounter1+1
  If flash1.state=0 then
    flash1.state=1
    PLAYFIELD_flasher1.visible=1
  else
    flash1.state=0
    PLAYFIELD_flasher1.visible=0
  end if

  If FlasherCounter1>9 then
    flash1.state=0
    PLAYFIELD_flasher1.visible=0
    flash1.timerenabled=false
    FlasherCounter1=0
  end if
end sub

sub flash2_timer
  FlasherCounter2=FlasherCounter2+1
  If flash2.state=0 then
    flash2.state=1
    PLAYFIELD_flasher2.visible=1
  else
    flash2.state=0
    PLAYFIELD_flasher2.visible=0
  end if
  If FlasherCounter2>9 then
    flash2.state=0
    PLAYFIELD_flasher2.visible=0
    flash2.timerenabled=false
    FlasherCounter2=0
  end if
end sub

sub AttractModeHSDisplay_timer
  if HighScoresDisplayed=false then
    HighScoresDisplayed=true
    If B2SOn then
      Controller.B2SSetScorePlayer1 HighScore(1)
      Controller.B2SSetScorePlayer2 HighScore(1)
      Controller.B2SSetScorePlayer3 HighScore(1)
      Controller.B2SSetScorePlayer4 HighScore(1)
      Controller.B2SSetData 99,1
      Controller.B2SSetScorePlayer5 Credits
      Controller.B2SSetGameOver 1
    end if
    ReelGameOver.Setvalue(1)
    ReelHighGame.SetValue(1)
    SplitScoreForDT HighScore(1),PlayerScore1
    SplitScoreForDT HighScore(1),PlayerScore2
    SplitScoreForDT HighScore(1),PlayerScore3
    SplitScoreForDT HighScore(1),PlayerScore4
    SplitCreditsForDT Credits,CreditReels
  else
    HighScoresDisplayed=false
    If B2SOn then
      Controller.B2SSetScorePlayer1 Score(1)
      Controller.B2SSetScorePlayer2 Score(2)
      Controller.B2SSetScorePlayer3 Score(3)
      Controller.B2SSetScorePlayer4 Score(4)
      Controller.B2SSetData 99,0
      Controller.B2SSetScorePlayer5 Credits

    end if
    ReelHighGame.SetValue(0)
    SplitScoreForDT Score(1),PlayerScore1
    SplitScoreForDT Score(2),PlayerScore2
    SplitScoreForDT Score(3),PlayerScore3
    SplitScoreForDT Score(4),PlayerScore4
    SplitCreditsForDT Credits,CreditReels
  end if
end sub

Sub NessieBGTimer_Timer
  If (NessieBGCounter Mod 2)=0 then
    If B2SOn then Controller.B2SSetData 98,1
    Light3.state=1
  else
    If B2SOn then Controller.B2SSetData 98,0
    Light3.state=0
  end if
  NessieBGCounter=NessieBGCounter+1
  If NessieBGCounter=30 then
    NessieBGCounter=0
    NessieBGTimer.enabled=false
    If B2SOn then Controller.B2SSetData 98,0
  end if
end sub

Sub AttractNessieBG_timer
  NessieBGTimer.enabled=true
  PlaySound "Nessy_Roar"
end sub

Sub SplitScoreForDT(tempbgscore,unit)
  dim q1,q2,q3,q4,q5,q6,q7
  q7=0
  q6=int((tempbgscore mod 100)/10)
  q5=int((tempbgscore mod 1000)/100)
  q4=int((tempbgscore mod 10000)/1000)
  q3=int((tempbgscore mod 100000)/10000)
  q2=int((tempbgscore mod 1000000)/100000)
  q1=int((tempbgscore mod 10000000)/1000000)

  unit(0).setvalue(1)
  If tempbgscore<1 then unit(1).setvalue(0):unit(2).setvalue(0):unit(3).setvalue(0):unit(4).setvalue(0):unit(5).setvalue(0):unit(6).setvalue(0):EXIT SUB
  if tempbgscore>0 then
    unit(1).setvalue(q6+1)
  else
    unit(1).setvalue(0)
  end if
  if tempbgscore>99 then
    unit(2).setvalue(q5+1)
  else
    unit(2).setvalue(0)
  end if
  if tempbgscore>999 then
    unit(3).setvalue(q4+1)
  else
    unit(3).setvalue(0)
  end if
  if tempbgscore>9999 then
    unit(4).setvalue(q3+1)
  else
    unit(4).setvalue(0)
  end if
  if tempbgscore>99999 then
    unit(5).setvalue(q2+1)
  else
    unit(5).setvalue(0)
  end if
  if tempbgscore>999999 then
    unit(6).setvalue(q1+1)
  else
    unit(6).setvalue(0)
  end if
end sub

Sub SplitCreditsForDT(tempbgcredits,unit)
  dim q1,q2,q3,q4,q5,q6,q7
  q7=int((tempbgcredits mod 10))
  q6=int((tempbgcredits mod 100)/10)


  If tempbgcredits<1 then unit(1).setvalue(0):unit(0).setvalue(1):EXIT SUB
  unit(0).setvalue(q7+1)
  if tempbgcredits>9 then
    unit(1).setvalue(q6+1)
  else
    unit(1).setvalue(0)
  end if
end sub

Sub SplitBIPForDT(tempbgbip,unit)
  dim q1,q2,q3,q4,q5,q6,q7
  q7=int((tempbgbip mod 10))
  q6=int((tempbgbip mod 100)/10)


  If tempbgbip<1 then unit(1).setvalue(0):unit(0).setvalue(0):EXIT SUB
  unit(0).setvalue(q7+1)
  if tempbgbip>9 then
    unit(1).setvalue(q6+1)
  else
    unit(1).setvalue(0)
  end if
end sub

Sub SetPlayerUp
  dim qq
  for qq=0 to 3
    PlayerUpLights(qq).state=0
  next
  PlayerUpLights(CurrentPlayer-1).state=1
  If B2SOn then
    Controller.B2SSetPlayerUp CurrentPlayer
  end if
end sub

Sub FlipAltRelay
  If lOutlaneL.state=1 OR lOutlaneR.state=1 then
    If lOutlaneR.state=1 then
      lOutlaneR.state=0
      lOutlaneL.state=1
    else
      lOutlaneR.state=1
      lOutlaneL.state=0
    end if
  end if
  if lInlaneL.state=1 OR lInlaneR.state=1 then
    If lInlaneR.state=1 then
      lInlaneR.state=0
      lInlaneL.state=1
    else
      lInlaneR.state=1
      lInlaneL.state=0
    end if
  end if
end sub

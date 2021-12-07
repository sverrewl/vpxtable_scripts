' ****************************************************************
'                       VISUAL PINBALL X
'                 JPSalas Diablo Pinball Script
'                         Version 1.0.0
' ****************************************************************

'DOF Config by Arngrim
'101 Left Flipper
'102 Right Flipper
'103 Left Slingshot
'104 Left Slingshot Shake
'105 Right Slingshot
'106 Right Slingshot Shake
'107 Bumper Left
'108 Bumper Center
'109 Bumper Right
'110 Ball Release
'111 AutoPlunger
'112 White Flash on Strobe Toy for autoplunger, Chestout kicker, extra ball, party award
'113 Drain Light
'114 Start Button Light, when credit is bigger than 0
'115 Chestout kicker
'116 ChestDoor Drop or Reset
'117 Strobe 250ms
'118 Drop targets left hit
'119 Drop targets center hit
'120 Drop targets right hit
'121 Knocker for extra ball and award special
'122 Triggertop1 Light
'123 Triggertop2 Light
'124 Triggertop3 Light
'125 Strobe 500ms
'126 Beacon when mutliball start
'127 small shaker effect when magnets are ON
'128 beacon effect when ball is on chest hole
'129 Left Bumper Flash
'130 Center Bumper Flash
'131 Right Bumper Flash
'132 Outlane Left Light
'133 Inlanes Left Light
'134 Inlane Riht Light
'135 Outlane Right Light
'136 Chest flasher 2000 ms
'137 Chest flasher 1000 ms
'138 Bumper Left Blink Flasher
'139 Bumper Center Blink Flasher
'140 Bumper Right Blink Flasher
'141 RGB GI White
'142 RGB GI Purple
'143 RGB GI Darkblue
'144 RGB GI Blue
'145 RGB GI Green
'146 RGB GI Darkgreen
'147 RGB GI Yellow
'148 RGB GI Amber
'149 RGB GI Orange
'150 RGB GI Red
'151 Flasher Center Purple Blink
'152 Flasher Left Purple Blink
'153 Flasher Right Purple Blink
'154 Rear Flashers Blink 1000ms
'155 Rear Flashers Blink 500ms

Option Explicit
Randomize

' Thalamus - converted from wmv to mp4 - UltraDMD
' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!
' As this is a original - most sounds should be moved to the backglass.

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolFlip   = 1    ' Flipper volume.


Const BallSize = 50 ' 50 is the normal size
Const cGameName = "diablo"

' Load the core.vbs for supporting Subs and functions
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
    On Error Goto 0
End Sub

' Define any Constants
Const TableName = "Diablo"
Const myVersion = "1.0.0"
Const MaxPlayers = 4     ' from 1 to 4
Const BallSaverTime = 20 ' in seconds
Const MaxMultiplier = 5 ' limit to 5x in this game
Const BallsPerGame = 3   ' 3 or 5
Const MaxMultiballs = 5  ' max number of balls during multiballs

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
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot
Dim SuperJackpot
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue(4)
Dim bAutoPlunger
Dim bInstantInfo
Dim bAttractMode
Dim ao,bsnr 'Stats B2S use it
Dim ActiveDOFCol

' Define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock
Dim BallsInHole

' Define Game Flags
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim bMusicOn
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim bJustStarted

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger
Dim LMagnet, RMagnet

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    LoadEM
    Dim i
    Randomize

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    Set LMagnet = New cvpmTurnTable
    With LMagnet
        .InitTurnTable Magnet1, 90
        .spinCW = False
        .SpinUp = 90
        .SpinDown = 90
        .MotorOn = False
        .CreateEvents "LMagnet"
    End With

    Set RMagnet = New cvpmTurnTable
    With RMagnet
        .InitTurnTable Magnet2, 90
        .spinCW = True
        .SpinUp = 90
        .SpinDown = 90
        .MotorOn = False
        .CreateEvents "RMagnet"
    End With

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' load saved values, highscore, names, jackpot
    Loadhs

    ' Initalise the DMD display
    DMD_Init

    ' freeplay or coins
    bFreePlay = True 'we dont want coins

    ' Init main variables and any other flags
    bAttractMode = False
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bGameInPlay = False
    bAutoPlunger = False
    bMusicOn = True
    BallsOnPlayfield = 0
    BallsInLock = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bInstantInfo = False
    ' set any lights for the attract mode
    GiOff
    StartAttractMode

    ' Remove the cabinet rails if in FS mode
    If Table1.ShowDT = False then
        lrail.Visible = False
        rrail.Visible = False
    End If
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
    If Keycode = AddCreditKey Then
        Credits = Credits + 1
        DOF 114, DOFOn
        If(Tilted = False) Then
            DMDFlush
            DMD "black.jpg", "", "CREDITS " &credits, 500
            PlaySoundAtVol "fx_coin", drain, 1
            If NOT bGameInPlay Then ShowTableInfo
        End If
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
    End If

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    ' Table specific

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 6:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

        If keycode = LeftFlipperKey Then SolLFlipper 1:InstantInfoTimer.Enabled = True
        If keycode = RightFlipperKey Then SolRFlipper 1:InstantInfoTimer.Enabled = True

    If keycode = MechanicalTilt Then Tilted = 1:MechCheckTilt:End If

        If keycode = StartGameKey Then
            If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    DMDFlush
                    DMD "black.jpg", " ", PlayersPlayingGame & " PLAYERS", 500
                    PlaySound "so_fanfare1"
                Else
                    If(Credits> 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
                        DMDFlush
                        DMD "black.jpg", " ", PlayersPlayingGame & " PLAYERS", 500
                        PlaySound "so_fanfare1"
                    Else
                        ' Not Enough Credits to start a game.
                        DOF 114, DOFOff
                        DMDFlush
                        DMD "black.jpg", "CREDITS " &credits, "INSERT COIN", 500
                        PlaySound "so_nocredits"
                   End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey Then
                If(bFreePlay = True) Then
                    If(BallsOnPlayfield = 0) Then
                        ResetForNewGame()
                    End If
                Else
                    If(Credits> 0) Then
                        If(BallsOnPlayfield = 0) Then
                            Credits = Credits - 1
                            ResetForNewGame()
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        DOF 114, DOFOff
                        DMDFlush
                        DMD "black.jpg", "CREDITS " &credits, "INSERT COIN", 500
                        ShowTableInfo
                    End If
                End If
            End If
    End If ' If (GameInPlay)

if keycode = "3" then StartNextMode
if keycode = "4" then winmode mode(0)

End Sub

Sub Table1_KeyUp(ByVal keycode)

    If keycode = PlungerKey Then
        Plunger.Fire
    End If

    If hsbModeActive Then
        Exit Sub
    End If

    ' Table specific

    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then
            SolLFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
            End If
        End If
        If keycode = RightFlipperKey Then
            SolRFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
            End If
        End If
    End If
End Sub

Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    bInstantInfo = True
    DMDFlush
    UltraDMDTimer.Enabled = 1
End Sub

Sub InstantInfo
    Jackpot = 1000000 + Round(Score(CurrentPlayer) / 10, 0)
    DMD "black.jpg", "", "INSTANT INFO", 500
    DMD "black.jpg", "JACKPOT", Jackpot, 800
    DMD "black.jpg", "LEVEL", Level(CurrentPlayer), 800
    DMD "black.jpg", "BONUS MULT", BonusMultiplier(CurrentPlayer), 800
    DMD "black.jpg", "ORBIT BONUS", OrbitHits, 800
    DMD "black.jpg", "LANE BONUS", LaneBonus, 800
    DMD "black.jpg", "TARGET BONUS", TargetBonus, 800
    DMD "black.jpg", "RAMP BONUS", RampBonus, 800
    DMD "black.jpg", "MONSTERS KILLED", Monsters(CurrentPlayer), 800
End Sub

'*************
' Pause Table
'*************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub table1_Exit
  Savehs
  if B2SOn Then Controller.stop
  If Not UltraDMD is Nothing Then
    If UltraDMD.IsRendering Then
      UltraDMD.CancelRendering
    End If
  UltraDMD.Uninit
    UltraDMD = NULL
  End If
End Sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
        If bSkillshotReady = False Then
            RotateLaneLightsLeft
        End If
    Else
        PlaySoundAtVol SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
        If bSkillshotReady = False Then
            RotateLaneLightsRight
        End If
    Else
        PlaySoundAtVol SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.05, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.05, 0.25
End Sub

Sub RotateLaneLightsLeft
    Dim TempState
    'flipper lanes
    TempState = l15.State
    l15.State = l16.State
    l16.State = l17.State
    l17.State = l18.State
    l18.State = l19.State
    l19.State = TempState
    'top lanes
    TempState = l20.State
    l20.State = l21.State
    l21.State = l22.State
    l22.State = TempState
    If Mode(0) <> 2 AND bBumperFrenzy = False Then UpdateBumperLights
End Sub

Sub RotateLaneLightsRight
    Dim TempState
    'flipperlanes
    TempState = l19.State
    l19.State = l18.State
    l18.State = l17.State
    l17.State = l16.State
    l16.State = l15.State
    l15.State = TempState
    'top lanes
    TempState = l22.State
    l22.State = l21.State
    l21.State = l20.State
    l20.State = TempState
    If Mode(0) <> 2 AND bBumperFrenzy = False Then UpdateBumperLights
End Sub

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                    'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity) AND(Tilt <15) Then 'show a warning
        DMDFlush
        DMD "black.jpg", " ", "CAREFUL!", 800
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        DMDFlush
        DMD "black.jpg", " ", "TILT!", 99999
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
    End If
End Sub

Sub MechCheckTilt                                    'Called when table is nudged
        Tilted = True
        'display Tilt
        DMDFlush
        DMD "black.jpg", " ", "TILT!", 99999
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
End Sub


Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt> 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        'turn off GI and turn off all the lights
        GiOff
        LightSeqTilt.Play SeqAllOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        'Bumper1.Force = 0

        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        'Bumper1.Force = 6
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        DMDFlush
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0) Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'********************
' Music as wav sounds
'********************

Dim Song
Song = ""

Sub PlaySong(name)
    If bMusicOn Then
        If Song <> name Then
            StopSound Song
            Song = name
            If Song = "mu_end" Then
                PlaySound Song, 0, 0.1  'this last number is the volume, from 0 to 1
            Else
                PlaySound Song, -1, 0.1 'this last number is the volume, from 0 to 1
            End If
        End If
    End If
End Sub

'**********************
'     GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'**********************

Dim OldGiState
OldGiState = -1   'start witht the Gi off

Sub ChangeGi(col) 'changes the gi color
  Dim i
  ActiveDOFCol = 140 + col
  DOF ActiveDOFCol, DOFOn
  debug.print ActiveDOFCol
  For i = 140 to 150
    If i <> ActiveDOFCol Then DOF i, DOFOff
  Next
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = 3 Then 'we have 4 captive balls on the table (-1 means no balls, 0 is the first ball, 1 is the second..)
            GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    DOF ActiveDOFCol, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next

    StartFire
End Sub

Sub GiOff
    DOF ActiveDOFCol, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next

    StartFire
End Sub

Dim Fire1Pos, Fire2Pos, Fire3Pos, Fire4Pos, Fire5Pos, Flames
Flames = Array("fire01", "fire02", "fire03", "fire04", "fire05", "fire06", "fire07", "fire08", "fire09", _
    "fire10", "fire11", "fire12", "fire13", "fire14", "fire15", "fire16", "fire17", "fire18", "fire19",  _
    "fire20", "fire21", "fire22", "fire23", "fire24", "fire25", "fire26", "fire27", "fire28", "fire29", "fire30")

Sub StartFire
    Fire1Pos = 0
    Fire2Pos = 6
    Fire3Pos = 12
    Fire4Pos = 18
    Fire5Pos = 24
    FireTimer.Enabled = 1
End Sub

Sub FireTimer_Timer
    'debug.print fire1pos
    Fire1.ImageA = Flames(Fire1Pos)
    Fire2.ImageA = Flames(Fire2Pos)
    Fire3.ImageA = Flames(Fire3Pos)
    Fire4.ImageA = Flames(Fire4Pos)
    Fire5.ImageA = Flames(Fire5Pos)
    Fire1Pos = (Fire1Pos + 1) MOD 30
    Fire2Pos = (Fire2Pos + 1) MOD 30
    Fire3Pos = (Fire3Pos + 1) MOD 30
    Fire4Pos = (Fire4Pos + 1) MOD 30
    Fire5Pos = (Fire5Pos + 1) MOD 30
End Sub

' GI & light sequence effects

Sub GiEffect(n)
    Dim ii
    Select Case n
        Case 0 'all off
            LightSeqGi.Play SeqAlloff
        Case 1 'all blink
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 10, 10
        Case 2 'random
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqRandom, 50, , 1000
        Case 3 'upon
            LightSeqGi.UpdateInterval = 4
            LightSeqGi.Play SeqUpOn, 5, 1
        Case 4 ' left-right-left
            LightSeqGi.UpdateInterval = 5
            LightSeqGi.Play SeqLeftOn, 10, 1
            LightSeqGi.UpdateInterval = 5
            LightSeqGi.Play SeqRightOn, 10, 1
    End Select
End Sub

Sub LightEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqInserts.Play SeqAlloff
        Case 1 'all blink
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 10, 10
        Case 2 'random
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqRandom, 50, , 1000
        Case 3 'upon
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 10, 1
        Case 4 ' left-right-left
            LightSeqInserts.UpdateInterval = 5
            LightSeqInserts.Play SeqLeftOn, 10, 1
            LightSeqInserts.UpdateInterval = 5
            LightSeqInserts.Play SeqRightOn, 10, 1
    End Select
End Sub

' Flasher Effects using lights

Dim FEStep, FEffect
FEStep = 0
FEffect = 0

Sub FlashEffect(n)
    Dim ii
    Select case n
        Case 0 ' all off
            LightSeqFlasher.Play SeqAlloff
        Case 1 'all blink
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqBlinking, , 10, 10
        Case 2 'random
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqRandom, 50, , 1000
        Case 3 'upon
            LightSeqFlasher.UpdateInterval = 4
            LightSeqFlasher.Play SeqUpOn, 10, 1
        Case 4 ' left-right-left
            LightSeqFlasher.UpdateInterval = 5
            LightSeqFlasher.Play SeqLeftOn, 10, 1
            LightSeqFlasher.UpdateInterval = 5
            LightSeqFlasher.Play SeqRightOn, 10, 1
    End Select
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Random quotes from the game

Sub PlayQuote_timer() 'one quote each 2 minutes
    Dim Quote
    Quote = "di_quote" & INT(RND * 56) + 1
    PlaySound Quote
End Sub

' Ramp Soundss
Sub REnd1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub REnd2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
    Dim i

    bGameInPLay = True

    'resets the score display, and turn off attrack mode
    StopAttractMode
    GiOn

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    Next

    ' initialise any other flags
    Tilt = 0

    ' initialise Game variables
    Game_Init()

    ' you may wish to start some music, play a sound, do whatever at this point

    vpmtimer.addtimer 1500, "FirstBall '"
End Sub

' This is used to delay the start of a game to allow any attract sequence to
' complete.  When it expires it creates a ball for the player to start playing with

Sub FirstBall
    ' reset the table for a new ball
    ResetForNewPlayerBall()
    ' create a new ball in the shooters lane
    CreateNewBall()
End Sub

' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))

Sub ResetForNewPlayerBall()
    ' make sure the correct display is upto date
    AddScore 0

    ' set the current players bonus multiplier back down to 1X
    SetBonusMultiplier 1

    ' reset any drop targets, lights, game Mode etc..

    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False
    ResetNewBallLights()
    'Reset any table specific
    ResetNewBallVariables

    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True

    'and the skillshot
    bSkillShotReady = True

'Change the music ?
End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedball BallSize / 2

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    PlaySoundAtVol SoundFXDOF("fx_Ballrel", 110, DOFPulse, DOFContactors), BallRelease, VolKick
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield> 1 Then
        DOF 111, DOFPulse
    DOF 112, DOFPulse
    DOF 126, DOFPulse
        bMultiBallMode = True
        bAutoPlunger = True
    End If
End Sub

' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
    'and eject the first ball
    CreateMultiballTimer_Timer
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
    ' wait if there is a ball in the plunger lane
    If bBallInPlungerLane Then
        Exit Sub
    Else
        If BallsOnPlayfield <MaxMultiballs Then
            CreateNewBall()
            mBalls2Eject = mBalls2Eject -1
            If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
                CreateMultiballTimer.Enabled = False
            End If
        Else 'the max number of multiballs is reached, so stop the timer
            mBalls2Eject = 0
            CreateMultiballTimer.Enabled = False
        End If
    End If
End Sub

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded

Sub EndOfBall()
    Dim AwardPoints, TotalBonus, ii
    AwardPoints = 0
    TotalBonus = 0
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)

    If NOT Tilted Then

        'Count the bonus. This table uses several bonus
        'Lane Bonus
        AwardPoints = LaneBonus * 1000
        TotalBonus = AwardPoints
        DMD "bonus-background.mp4", "LANE BONUS", AwardPoints, 700
        vpmtimer.addtimer 1300, "PlaySound ""bonus"", 0, 1, 0, 0, 0, 0, 0 '"

        'Number of Target banks completed
        AwardPoints = TargetBonus * 100000
        TotalBonus = TotalBonus + AwardPoints
        DMD "bonus-background.mp4", "TARGET BONUS", AwardPoints, 700
        vpmtimer.addtimer 2200, "PlaySound ""bonus"", 0, 1, 0, 0, 1000, 0, 0 '"

        'Number of Ramps completed
        AwardPoints = RampBonus * 10000
        TotalBonus = TotalBonus + AwardPoints
        DMD "bonus-background.mp4", "RAMP BONUS", AwardPoints, 700
        vpmtimer.addtimer 2900, "PlaySound ""bonus"", 0, 1, 0, 0, 2000, 0, 0 '"

        'Number of Orbits registered
        AwardPoints = OrbitHits * 32260
        TotalBonus = TotalBonus + AwardPoints
        DMD "bonus-background.mp4", "ORBIT BONUS", AwardPoints, 700
        vpmtimer.addtimer 3600, "PlaySound ""bonus"", 0, 1, 0, 0, 3000, 0, 0 '"

        'Number of monsters defeated
        AwardPoints = Monsters(CurrentPlayer) * 25130
        DMD "bonus-background.mp4", "MONSTERS KILLED", monsters(CurrentPlayer), 700
        vpmtimer.addtimer 4400, "PlaySound ""bonus"", 0, 1, 0, 0, 4000, 0, 0 '"

        'Player Level
        AwardPoints = Level(CurrentPlayer) * 50000
        TotalBonus = TotalBonus + AwardPoints
        DMD "bonus-background.mp4", "LEVEL BONUS", Level(CurrentPlayer), 700
        vpmtimer.addtimer 5000, "PlaySound ""bonus"", 0, 1, 0, 0, 5000, 0, 0 '"

        ' calculate the totalbonus
        TotalBonus = TotalBonus * BonusMultiplier(CurrentPlayer) + BonusHeldPoints(CurrentPlayer)

        ' handle the bonus held
        ' reset the bonus held value since it has been already added to the bonus
        BonusHeldPoints(CurrentPlayer) = 0

        ' the player has won the bonus held award so do something with it :)
        If bBonusHeld Then
            If Balls = BallsPerGame Then ' this is the last ball, so if bonus held has been awarded then double the bonus
                TotalBonus = TotalBonus * 2
            End If
        Else ' this is not the last ball so save the bonus for the next ball
            BonusHeldPoints(CurrentPlayer) = TotalBonus
        End If
        bBonusHeld = False

        ' Add the bonus to the score
        DMD "bonus-background.mp4", "TOTAL BONUS X " &BonusMultiplier(CurrentPlayer), TotalBonus, 2200
        vpmtimer.addtimer 6200, "PlaySound ""goblin"" '"
        AddScore TotalBonus

        ' add a bit of a delay to allow for the bonus points to be shown & added up
        vpmtimer.addtimer 8500, "EndOfBall2 '"
    Else 'if tilted then only add a short delay
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the currentplayer)
'
Sub EndOfBall2()
    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilted = False
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

    ' has the player won an extra-ball ? (might be multiple outstanding)
    If(ExtraBallsAwards(CurrentPlayer) <> 0) Then
        'debug.print "Extra Ball"

        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
            LightShootAgain.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
        DMD "extra-ball.mp4", "", "", 5000

        ' In this table an extra ball will have the skillshot and ball saver, so we reset the playfield for the new ball
        ResetForNewPlayerBall()

        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0) Then
            'debug.print "No More Balls, High Score Entry"

            ' Submit the currentplayers score to the High Score system
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

    'debug.print "EndOfBall - Complete"

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame> 1) Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer> PlayersPlayingGame) Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

        ' set the machine into game over mode
        EndOfGame()

    ' you may wish to put a Game Over message on the desktop/backglass

    Else
        ' set the next player
        CurrentPlayer = NextPlayer

        ' make sure the correct display is up to date
        AddScore 0

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()

        ' play a sound if more than 1 player
        If PlayersPlayingGame> 1 Then
            PlaySound "vo_player" &CurrentPlayer
            DMD "black.jpg", " ", "PLAYER " &CurrentPlayer, 800
        End If
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    'debug.print "End Of Game"
    bGameInPLay = False
    ' just ended your game then play the end of game tune
    If NOT bJustStarted Then
    'PlaySong "m_end"
    End If

    bJustStarted = False
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball
    'PlayQuote.Enabled = 0
    ' show game over on the DMD
    DMD "game-over.mp4", "", "", 11000

    ' set any lights for the attract mode
    GiOff
    StartAttractMode
' you may wish to light any Game Over Light you may have
End Sub

Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp> BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' lost a ball ;-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one
'
Sub Drain_Hit()
    ' Destroy the ball
    Drain.DestroyBall
    ' Exit Sub ' only for debugging - this way you can add balls from the debug window

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' pretend to knock the ball into the ball storage mech
    PlaySoundAtVol "fx_drain", drain, 1
  DOF 113, DOFPulse
    'if Tilted the end Ball Mode
    If Tilted Then
        StopEndOfBallMode
    End If

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True) AND(Tilted = False) Then

        ' is the ball saver active,
        If(bBallSaverActive = True) Then

            ' yep, create a new ball in the shooters lane
            ' we use the Addmultiball in case the multiballs are being ejected
            AddMultiball 1
            ' we kick the ball with the autoplunger
            bAutoPlunger = True
            ' you may wish to put something on a display or play a sound at this point
            DMD "ball-save.mp4", "", "", 5000
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1) Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True) then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ' you may wish to change any music over at this point and
                    ' turn off any multiball specific lights
                    ' ResetJackpotLights
          l42.State = 0
          l43.State = 0
          l44.State = 0
          l45.State = 0
          l46.State = 0
                    PlaySong "mu_main"
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0) Then
                ' End Mode and timers
                StopSound Song:Song = ""
                ChangeGi white
                ' Show the end of ball animation
                ' and continue with the end of ball
                DMDFlush
                Select Case Mode(0)
                    Case 0, 1, 3, 5, 7
                        DMD "ball-lost-2.mp4", "", "", 5000
                        vpmtimer.addtimer 4500, "EndOfBall '"
                    Case 2, 4, 6, 8, 10
                        DMD "ball-lost-1.mp4", "", "", 10000
                        vpmtimer.addtimer 9500, "EndOfBall '"
                End Select
                StopEndOfBallMode
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    PlaySoundAtVol "fx_sensor", plunger, 1
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
    'LaunchLight.State = 2

    'be sure to update the Scoreboard after the animations, if any
    UltraDMDScoreTimer.Enabled = 1

    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        PlungerIM.AutoFire
        DOF 111, DOFPulse
    DOF 112, DOFPulse
        bAutoPlunger = False
    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    End If
    'Start the Selection of the skillshot if ready
    If bSkillShotReady Then
        swPlungerRest.TimerEnabled = 1 ' this is a new ball, so show the launch ball if inactive for 6 seconds
        UpdateSkillshot()
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
        ResetSkillShotTimer.Enabled = 1
    End If
    If NOT bMultiballMode Then
        PlaySong "mu_main"
    End If
' turn off LaunchLight
' LaunchLight.State = 0
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds

Sub swPlungerRest_Timer
    DMD "start.mp4", "", "", 17000
    swPlungerRest.TimerEnabled = 0
End Sub

Sub EnableBallSaver(seconds)
    'debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
    BallSaverTimerExpired.Interval = 1000 * seconds
    BallSaverTimerExpired.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    SetLightColor LightShootAgain, amber, 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimerExpired_Timer()
    'debug.print "Ballsaver ended"
    BallSaverTimerExpired.Enabled = False
    ' clear the flag
    bBallSaverActive = False
    ' if you have a ball saver light then turn it off at this point
    LightShootAgain.State = 0
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    'debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
End Sub

' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board
' In this table we use SecondRound variable to double the score points in the second round after killing Malthael
Sub AddScore(points)
    If(Tilted = False) Then
        ' add the points to the current players score variable
        Score(CurrentPlayer) = Score(CurrentPlayer) + points * SecondRound
        ' update the score displays
        DMDScore
    End if

' you may wish to check to see if the player has gotten a replay
End Sub

' Add bonus to the bonuspoints AND update the score board

Sub AddBonus(points) 'not used in this table, since there are many different bonus items.
    If(Tilted = False) Then
        ' add the bonus to the current players bonus variable
        BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
        ' update the score displays
        DMDScore
    End if

' you may wish to check to see if the player has gotten a replay
End Sub

' Add some points to the current Jackpot.
'
Sub AddJackpot(points) 'not used in this table
' Jackpots only generally increment in multiball mode AND not tilted
' but this doesn't have to be the case
'If(Tilted = False)Then

' If(bMultiBallMode = True) Then
' Jackpot = Jackpot + points
' you may wish to limit the jackpot to a upper limit, ie..
' If (Jackpot >= 6000) Then
'   Jackpot = 6000
'   End if
'End if
'End if
End Sub

Sub AddSuperJackpot(points)
    If(Tilted = False) Then

    ' If(bMultiBallMode = True) Then
    '   SuperJackpot = SuperJackpot + points
    ' you may wish to limit the jackpot to a upper limit, ie..
    ' If (Jackpot >= 6000) Then
    '   Jackpot = 6000
    '   End if
    'End if
    End if
End Sub

Sub AddBonusMultiplier(n)
    Dim NewBonusLevel
    ' if not at the maximum bonus level
    if(BonusMultiplier(CurrentPlayer) + n <MaxMultiplier) then
        ' then add and set the lights
        NewBonusLevel = BonusMultiplier(CurrentPlayer) + n
        SetBonusMultiplier(NewBonusLevel)
    Else
        l2.State = 2:l3.State = 2:l4.State = 2:l5.State = 2
        AddScore 5000000
        DMD "black.jpg", " ", "5.000.000", 1000
    End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly
' There is no bonus multiplier lights in this table

Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level
    ' Update the lights
    Select Case Level
        Case 1:l2.State = 0:l3.State = 0:l4.State = 0:l5.State = 0
        Case 2:l2.State = 1:l3.State = 0:l4.State = 0:l5.State = 0
        Case 3:l2.State = 0:l3.State = 1:l4.State = 0:l5.State = 0
        Case 4:l2.State = 0:l3.State = 0:l4.State = 1:l5.State = 0
        Case 5:l2.State = 0:l3.State = 0:l4.State = 0:l5.State = 1
    End Select
End Sub

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        DMDBlink "black.jpg", " ", "EXTRA BALL WON", 100, 10
        'DMD "extra-ball.mp4", "", "", 5000
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    DOF 121, DOFPulse
    DOF 112, DOFPulse
        bExtraBallWonThisBall = True
        GiEffect 1
        LightEffect 2
    END If
End Sub

Sub AwardSpecial()
    DMDBlink "black.jpg", " ", "EXTRA GAME WON", 100, 10
    Credits = Credits + 1
    DOF 114, DOFOn
  DOF 121, DOFPulse
  DOF 112, DOFPulse
    GiEffect 1
    LightEffect 1
End Sub

'in this table the jackpot is always 1 million + 10% of your score

Sub AwardJackpot() 'award a normal jackpot, double or triple jackpot
    Jackpot = 1000000 + Round(Score(CurrentPlayer) / 10, 0)
    DMDBlink "black.jpg", "JACKPOT", Jackpot, 100, 10
    PlaySound "criticalhit"
    AddScore Jackpot
    GiEffect 1
  DOF 125, DOFPulse
    LightEffect 2
    FlashEffect 2
  DOF 154, DOFPulse
End Sub

Sub AwardDoubleJackpot() 'in this table the jackpot is always 1 million + 10% of your score
    Jackpot = (1000000 + Round(Score(CurrentPlayer) / 10, 0) ) * 2
    DMDBlink "black.jpg", "DOUBLE JACKPOT", Jackpot, 100, 10
    PlaySound "criticalhit"
    AddScore Jackpot
    GiEffect 1
  DOF 125, DOFPulse
    LightEffect 2
    FlashEffect 2
  DOF 154, DOFPulse
End Sub

Sub AwardSuperJackpot() 'this is actually a tripple jackpot
    SuperJackpot = (1000000 + Round(Score(CurrentPlayer) / 10, 0) ) * 3
    DMDBlink "black.jpg", "SUPER JACKPOT", SuperJackpot, 100, 10
    PlaySound "criticalhit"
    AddScore SuperJackpot
    GiEffect 1
  DOF 125, DOFPulse
    LightEffect 2
    FlashEffect 2
  DOF 154, DOFPulse
End Sub

Sub AwardSkillshot()
    Dim i
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMDFlush
    Select case SkillShotValue(CurrentPlayer)
        case 1000000
            DMD "skillshot1.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case 2000000
            DMD "skillshot2.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case 3000000
            DMD "skillshot3.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case 4000000
            DMD "skillshot4.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case 5000000
            DMD "skillshot5.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case ELSE
            DMD "skillshot.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
    End Select
    ' increment the skillshot value with 1 million
    SkillShotValue(CurrentPLayer) = SkillShotValue(CurrentPLayer) + 1000000
    'do some light show
    GiEffect 1
  DOF 117, DOFPulse
    LightEffect 2
    'enable the start act/battle by opening the chest door
    DropChestDoor
End Sub

Sub Congratulation()
    Dim tmp
    tmp = "vo_congrat" & INT(RND * 21 + 1)
    PlaySound tmp
End Sub
'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(cGameName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 100000 End If

    x = LoadValue(cGameName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If

    x = LoadValue(cGameName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 100000 End If

    x = LoadValue(cGameName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If

    x = LoadValue(cGameName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 100000 End If

    x = LoadValue(cGameName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If

    x = LoadValue(cGameName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 100000 End If

    x = LoadValue(cGameName, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If

    x = LoadValue(cGameName, "Credits")
    If(x <> "") then Credits = CInt(x):DOF 114, DOFOn:Else Credits = 0 End If

    'x = LoadValue(cGameName, "Jackpot")
    'If(x <> "") then Jackpot = CDbl(x) Else Jackpot = 200000 End If
    x = LoadValue(cGameName, "TotalGamesPlayed")
    If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If
End Sub

Sub Savehs
    SaveValue cGameName, "HighScore1", HighScore(0)
    SaveValue cGameName, "HighScore1Name", HighScoreName(0)
    SaveValue cGameName, "HighScore2", HighScore(1)
    SaveValue cGameName, "HighScore2Name", HighScoreName(1)
    SaveValue cGameName, "HighScore3", HighScore(2)
    SaveValue cGameName, "HighScore3Name", HighScoreName(2)
    SaveValue cGameName, "HighScore4", HighScore(3)
    SaveValue cGameName, "HighScore4Name", HighScoreName(3)
    SaveValue cGameName, "Credits", Credits
    'SaveValue cGameName, "Jackpot", Jackpot
    SaveValue cGameName, "TotalGamesPlayed", TotalGamesPlayed
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
    Dim tmp
    tmp = Score(1)

    If Score(2)> tmp Then tmp = Score(2)
    If Score(3)> tmp Then tmp = Score(3)
    If Score(4)> tmp Then tmp = Score(4)

    If tmp> HighScore(1) Then 'add 1 credit for beating the highscore
        AwardSpecial
    End If

    If tmp> HighScore(3) Then
        vpmtimer.addtimer 2000, "PlaySound ""vo_contratulationsgreatscore"" '"
        HighScore(3) = tmp
        'enter player's name
        HighScoreEntryInit()
    Else
        EndOfBallComplete()
    End If
End Sub

Sub HighScoreEntryInit()
    hsbModeActive = True
    PlaySound "vo_enteryourinitials"
    hsLetterFlash = 0

    hsEnteredDigits(0) = "A"
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ<+-0123456789" ' < is used to delete the last letter
    hsCurrentLetter = 1
    DMDFlush
    DMDId "hsc", "black.jpg", "YOUR NAME:", " ", 999999
    HighScoreDisplayName()
End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        Playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayName()
    End If

    If keycode = RightFlipperKey Then
        Playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayName()
    End If

    If keycode = StartGameKey OR keycode = PlungerKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<") then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) then
                HighScoreCommitName()
            else
                HighScoreDisplayName()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit> 0) then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayName()
        end if
    end if
End Sub

Sub HighScoreDisplayName()
    Dim i, TempStr

    TempStr = " >"
    if(hsCurrentDigit> 0) then TempStr = TempStr & hsEnteredDigits(0)
    if(hsCurrentDigit> 1) then TempStr = TempStr & hsEnteredDigits(1)
    if(hsCurrentDigit> 2) then TempStr = TempStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3) then
        if(hsLetterFlash <> 0) then
            TempStr = TempStr & "_"
        else
            TempStr = TempStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit <1) then TempStr = TempStr & hsEnteredDigits(1)
    if(hsCurrentDigit <2) then TempStr = TempStr & hsEnteredDigits(2)

    TempStr = TempStr & "< "
    DMDMod "hsc", "YOUR NAME:", Mid(TempStr, 2, 5), 999999
End Sub

Sub HighScoreCommitName()
    hsbModeActive = False
    'PlaySong "m_end"
    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ") then
        hsEnteredName = "YOU"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
    DMDFlush
    EndOfBallComplete()
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 0 to 3
        For j = 0 to 2
            If HighScore(j) <HighScore(j + 1) Then
                tmp = HighScore(j + 1)
                tmp2 = HighScoreName(j + 1)
                HighScore(j + 1) = HighScore(j)
                HighScoreName(j + 1) = HighScoreName(j)
                HighScore(j) = tmp
                HighScoreName(j) = tmp2
            End If
        Next
    Next
End Sub

'****************************************
' Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub GameTimer_Timer
    RollingUpdate
' add any other real time update subs, like gates or diverters
End Sub

'********************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version

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
' Change light color - simulate color leds
' changes the light color and state
' 10 colors: red, orange, amber, yellow...
'******************************************
' in this table this colors are use to keep track of the progress during the acts and battles

'colors
Dim red, orange, amber, yellow, darkgreen, green, blue, darkblue, purple, white

red = 10
orange = 9
amber = 8
yellow = 7
darkgreen = 6
green = 5
blue = 4
darkblue = 3
purple = 2
white = 1

Sub SetLightColor(n, col, stat)
    Select Case col
        Case 0
            n.color = RGB(18, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case red
            n.color = RGB(18, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case orange
            n.color = RGB(18, 3, 0)
            n.colorfull = RGB(255, 64, 0)
        Case amber
            n.color = RGB(193, 49, 0)
            n.colorfull = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(18, 18, 0)
            n.colorfull = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 8, 0)
            n.colorfull = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 18, 0)
            n.colorfull = RGB(0, 255, 0)
        Case blue
            n.color = RGB(0, 18, 18)
            n.colorfull = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 8, 8)
            n.colorfull = RGB(0, 64, 64)
        Case purple
            n.color = RGB(128, 0, 128)
            n.colorfull = RGB(255, 0, 255)
        Case white
            n.color = RGB(255, 252, 224)
            n.colorfull = RGB(193, 91, 0)
        Case white
            n.color = RGB(255, 252, 224)
            n.colorfull = RGB(193, 91, 0)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

Sub ResetAllLightsColor ' Called at a new game
    'shoot again
    SetLightColor LightShootAgain, amber, -1
    ' bonus
    SetLightColor l2, yellow, -1
    SetLightColor l3, yellow, -1
    SetLightColor l4, yellow, -1
    SetLightColor l5, yellow, -1
    ' Acts
    SetLightColor l7, red, -1
    SetLightColor l8, red, -1
    SetLightColor l10, red, -1
    SetLightColor l12, red, -1
    SetLightColor l13, red, -1
    ' Lords
    SetLightColor l6, red, -1
    SetLightColor l9, red, -1
    SetLightColor l11, red, -1
    SetLightColor l14, red, -1
    SetLightColor l47, red, -1
    ' flipper lanes
    SetLightColor l15, orange, -1
    SetLightColor l16, orange, -1
    SetLightColor l17, orange, -1
    SetLightColor l18, orange, -1
    SetLightColor l19, orange, -1
    ' bash
    SetLightColor l28, blue, -1
    SetLightColor l29, blue, -1
    SetLightColor l30, blue, -1
    SetLightColor l31, blue, -1
    ' skill
    SetLightColor l35, yellow, -1
    SetLightColor l36, yellow, -1
    SetLightColor l37, yellow, -1
    SetLightColor l40, yellow, -1
    SetLightColor l41, yellow, -1
    ' life - extra ball
    SetLightColor l23, purple, -1
    SetLightColor l24, purple, -1
    SetLightColor l25, purple, -1
    SetLightColor l26, purple, -1
    SetLightColor l27, purple, -1
    ' attack arrows
    SetLightColor l32, green, -1
    SetLightColor l33, green, -1
    SetLightColor l34, green, -1
    SetLightColor l38, green, -1
    SetLightColor l39, green, -1
    ' critical - jackpot
    SetLightColor l42, red, -1
    SetLightColor l43, red, -1
    SetLightColor l44, green, -1
    SetLightColor l45, red, -1
    SetLightColor l46, red, -1
    ' level
    SetLightColor l20, orange, -1
    SetLightColor l21, orange, -1
    SetLightColor l22, orange, -1
End Sub

Sub UpdateBonusColors
End Sub

'*************************
' Rainbow Changing Lights
'*************************

Dim RGBStep, RGBFactor, rRed, rGreen, rBlue, RainbowLights

Sub StartRainbow(n)
    set RainbowLights = n
    RGBStep = 0
    RGBFactor = 5
    rRed = 255
    rGreen = 0
    rBlue = 0
    RainbowTimer.Enabled = 1
End Sub

Sub StopRainbow()
    Dim obj
    RainbowTimer.Enabled = 0
    RainbowTimer.Enabled = 0
End Sub

Sub RainbowTimer_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep
        Case 0 'Green
            rGreen = rGreen + RGBFactor
            If rGreen> 255 then
                rGreen = 255
                RGBStep = 1
            End If
        Case 1 'Red
            rRed = rRed - RGBFactor
            If rRed <0 then
                rRed = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            rBlue = rBlue + RGBFactor
            If rBlue> 255 then
                rBlue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            rGreen = rGreen - RGBFactor
            If rGreen <0 then
                rGreen = 0
                RGBStep = 4
            End If
        Case 4 'Red
            rRed = rRed + RGBFactor
            If rRed> 255 then
                rRed = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            rBlue = rBlue - RGBFactor
            If rBlue <0 then
                rBlue = 0
                RGBStep = 0
            End If
    End Select
    For each obj in RainbowLights
        obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
        obj.colorfull = RGB(rRed, rGreen, rBlue)
    Next
End Sub

'***********************************************************************************
'               JPS DMD - very simple DMD routines using UltraDMD
'***********************************************************************************

Dim UltraDMD
Dim FlexDMD

' DMD using UltraDMD calls

Sub DMD(background, toptext, bottomtext, duration)
    UltraDMD.DisplayScene00 background, toptext, 15, bottomtext, 15, 14, duration, 14
    UltraDMDScoreTimer.Enabled = 1                               'to show the score after the animation/message
End Sub

Sub DMDBlink(background, toptext, bottomtext, duration, nblinks) 'blinks the lower text nblinks times
    Dim i
    For i = 1 to nblinks
        UltraDMD.DisplayScene00 background, toptext, 15, bottomtext, 15, 14, duration, 14
    Next

    UltraDMDScoreTimer.Enabled = 1 'to show the score after the animation/message
End Sub

Sub DMDScore
    If NOT UltraDMD.IsRendering Then
        UltraDMD.SetScoreboardBackgroundImage "scoreboard-background.jpg", 15, 7
        Select case Mode(0)
            Case 0, 11 'no mode is active, then show the 4 players, current player, Level and ball number
                UltraDMD.SetScoreboardBackgroundImage "scoreboard-background.jpg", 15, 7
                UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer & " Level " & Level(CurrentPlayer), "Ball " & Balls
            Case 1, 3, 5, 7, 9 'during acts, show the 4 players, player number, Level and number of Monsters hit and the total number of Monsters
                UltraDMD.SetScoreboardBackgroundImage "scoreboard-background.jpg", 15, 7
                UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer & " Level " & Level(CurrentPlayer), "Act " & MonsterHits & "/" & MonsterTotal
            Case 2 'during monster battles show the monsters life and the score
                UltraDMD.SetScoreboardBackgroundImage Butcher(ColorCode), 15, 7
                UltraDMD.DisplayScoreboard 2, 1, Score(CurrentPlayer), LordStrength, 0, 0, "", ""
            Case 4 'during monster battles show the monsters life and the score
                UltraDMD.SetScoreboardBackgroundImage Belial(ColorCode), 15, 7
                UltraDMD.DisplayScoreboard 2, 1, Score(CurrentPlayer), LordStrength, 0, 0, "", ""
            Case 6 'during monster battles show the monsters life and the score
                UltraDMD.SetScoreboardBackgroundImage Azmodan(ColorCode), 15, 7
                UltraDMD.DisplayScoreboard 2, 1, Score(CurrentPlayer), LordStrength, 0, 0, "", ""
            Case 8 'during monster battles show the monsters life and the score
                UltraDMD.SetScoreboardBackgroundImage Diablo(ColorCode), 15, 7
                UltraDMD.DisplayScoreboard 2, 1, Score(CurrentPlayer), LordStrength, 0, 0, "", ""
            Case 10 'during monster battles show the monsters life and the score
                UltraDMD.SetScoreboardBackgroundImage Malthael(ColorCode), 15, 7
                UltraDMD.DisplayScoreboard 2, 1, Score(CurrentPlayer), LordStrength, 0, 0, "", ""
        End Select
    Else
        UltraDMDScoreTimer.Enabled = 1
    End If
End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMDFLush
    UltraDMDTimer.Enabled = 0
    UltraDMDScoreTimer.Enabled = 0
    UltraDMD.CancelRendering
    UltraDMD.Clear
End Sub

Sub DMDScrollCredits(background, text, duration)
    UltraDMD.ScrollingCredits background, text, 15, 14, duration, 14
End Sub

Sub DMDId(id, background, toptext, bottomtext, duration)
    UltraDMD.DisplayScene00ExwithID id, False, background, toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
End Sub

Sub DMDMod(id, toptext, bottomtext, duration)
    UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
End Sub

Sub UltraDMDTimer_Timer() 'used for repeating the attrack mode and the instant info.
    If bInstantInfo Then
        InstantInfo
    ElseIf bAttractMode Then
        ShowTableInfo
    End If
End Sub

Sub UltraDMDScoreTimer_Timer()
    If NOT UltraDMD.IsRendering Then
        DMDScoreNow
    End If
End Sub

Dim Butcher, Belial, Azmodan, Diablo, Malthael
Butcher = Array("butch-0.jpg", "butch-10.jpg", "butch-20.jpg", "butch-30.jpg", "butch-40.jpg", "butch-50.jpg", "butch-60.jpg", "butch-70.jpg", "butch-80.jpg", "butch-90.jpg", "butch-100.jpg")
Belial = Array("bel-0.jpg", "bel-10.jpg", "bel-20.jpg", "bel-30.jpg", "bel-40.jpg", "bel-50.jpg", "bel-60.jpg", "bel-70.jpg", "bel-80.jpg", "bel-90.jpg", "bel-100.jpg")
Azmodan = Array("az-0.jpg", "az-10.jpg", "az-20.jpg", "az-30.jpg", "az-40.jpg", "az-50.jpg", "az-60.jpg", "az-70.jpg", "az-80.jpg", "az-90.jpg", "az-100.jpg")
Diablo = Array("dia-0.jpg", "dia-10.jpg", "dia-20.jpg", "dia-30.jpg", "dia-40.jpg", "dia-50.jpg", "dia-60.jpg", "dia-70.jpg", "dia-80.jpg", "dia-90.jpg", "dia-100.jpg")
Malthael = Array("mal-0.jpg", "mal-10.jpg", "mal-20.jpg", "mal-30.jpg", "mal-40.jpg", "mal-50.jpg", "mal-60.jpg", "mal-70.jpg", "mal-80.jpg", "mal-90.jpg", "mal-100.jpg")

' Update the embedded DMD inside VPX. This requires the table to have a timer object, named 'FlexDMDTimer', enabled, with a timer interval of -1 (VPX > 10.2)
Sub FlexDMDTimer_Timer
  Dim DMDp
  If UseDMD Then
    DMDp = FlexDMD.DmdPixels
    If Not IsEmpty(DMDp) Then
      DMDWidth = FlexDMD.Width
      DMDHeight = FlexDMD.Height
      DMDPixels = DMDp
    End If
  ElseIf UseColoredDMD Then
    DMDp = FlexDMD.DmdColoredPixels
    If Not IsEmpty(DMDp) Then
      DMDWidth = FlexDMD.Width
      DMDHeight = FlexDMD.Height
      DMDColoredPixels = DMDp
    End If
  End If
End Sub


Sub DMD_Init
  UseColoredDMD = true
    'Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
    If FlexDMD is Nothing Then
        MsgBox "No UltraDMD found.  This table will NOT run without it."
        Exit Sub
    End If

    FlexDMD.GameName = cGameName
  FlexDMD.RenderMode = 2
  Set UltraDMD = FlexDMD.NewUltraDMD()
    UltraDMD.Init
    If Not UltraDMD.GetMajorVersion = 1 Then
        MsgBox "Incompatible Version of UltraDMD found."
        Exit Sub
    End If

    If UltraDMD.GetMinorVersion <1 Then
        MsgBox "Incompatible Version of UltraDMD found. Please update to version 1.1 or newer."
        Exit Sub
    End If

    Dim fso:Set fso = CreateObject("Scripting.FileSystemObject")
    Dim curDir:curDir = fso.GetAbsolutePathName(".")

    Dim DirName
    DirName = curDir& "\" &cGameName& ".UltraDMD"

    If Not fso.FolderExists(DirName) Then _
            Msgbox "UltraDMD userfiles directory '" & DirName & "' does not exist." & CHR(13) & "No graphic images will be displayed on the DMD"
    UltraDMD.SetProjectFolder DirName

    ' wait for the animation to end
    While UltraDMD.IsRendering = True
    WEnd

    ' Show ROM version number
    DMD "black.jpg", "DIABLO", "ROM VERS " &myVersion, 1000
End Sub

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo
    Dim i
    'info goes in a loop only stopped by the credits and the startkey
    If Score(1) Then
        DMD "black.jpg", "PLAYER 1", Score(1), 3000
    End If
    If Score(2) Then
        DMD "black.jpg", "PLAYER 2", Score(2), 3000
    End If
    If Score(3) Then
        DMD "black.jpg", "PLAYER 3", Score(3), 3000
    End If
    If Score(4) Then
        DMD "black.jpg", "PLAYER 4", Score(4), 3000
    End If

    'coins or freeplay
    If bFreePlay Then
        DMD "black.jpg", " ", "FREE PLAY", 2000
        DMD "intro-freeplay.mp4", "", "", 63000
    Else
        If Credits> 0 Then
            DMD "black.jpg", "CREDITS " &credits, "PRESS START", 2000
        Else
            DMD "black.jpg", "CREDITS " &credits, "INSERT COIN", 2000
        End If
        DMD "intro-coins.mp4", "", "", 65000
    End If

    DMD "black.jpg", "HIGHSCORES", "1> " & HighScoreName(0) & " " & FormatNumber(HighScore(0), 0, , , -1), 3000
    DMD "black.jpg", "HIGHSCORES", "2> " & HighScoreName(1) & " " & FormatNumber(HighScore(1), 0, , , -1), 3000
    DMD "black.jpg", "HIGHSCORES", "3> " & HighScoreName(2) & " " & FormatNumber(HighScore(2), 0, , , -1), 3000
    DMD "black.jpg", "HIGHSCORES", "4> " & HighScoreName(3) & " " & FormatNumber(HighScore(3), 0, , , -1), 3000
End Sub

Sub StartAttractMode()
    bAttractMode = True
    UltraDMDTimer.Enabled = 1
    StartLightSeq
    ShowTableInfo
    StartRainbow aLights
End Sub

Sub StopAttractMode()
    bAttractMode = False
    DMDScoreNow
    LightSeqAttract.StopPlay
    LightSeqFlasher.StopPlay
    StopRainbow
    ResetAllLightsColor
'StopSong
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
    LightSeqAttract.UpdateInterval = 25
    LightSeqAttract.Play SeqBlinking, , 5, 150
    LightSeqAttract.Play SeqRandom, 40, , 4000
    LightSeqAttract.Play SeqAllOff
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 40, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 40, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqRightOn, 30, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqLeftOn, 30, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqStripe1VertOn, 50, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
End Sub

Sub LightSeqAttract_PlayDone()
    StartLightSeq()
End Sub

Sub LightSeqTilt_PlayDone()
    LightSeqTilt.Play SeqAllOff
End Sub

Sub LightSeqSkillshot_PlayDone()
    LightSeqSkillshot.Play SeqAllOff
End Sub

'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************

' droptargets, animations, etc
Sub VPObjects_Init

End Sub

' tables variables and Mode init
Dim Level(4)
Dim demons(4)
Dim monsters(4)
Dim LaneBonus
Dim TargetBonus
Dim RampBonus
Dim OrbitHits
Dim MonsterHits
Dim MonsterTotal
Dim bActReady
Dim LordStrength
Dim LordLifeStep
Dim LordLifeInterval
Dim bBumperFrenzy
Dim BumperValue(4)
Dim SecondRound
Dim bJackpot
Dim ComboValue

Sub Game_Init() 'called at the start of a new game
    Dim i
    bExtraBallWonThisBall = False
    TurnOffPlayfieldLights()
    'Play some Music
    PlaySong "mu_main"
    'Init Variables
    For i = 0 to 4
        Level(i) = 0
        demons(i) = 0
        monsters(i) = 0
        SkillshotValue(i) = 1000000 ' increases by 1000000 each time it is collected
    Next

    LaneBonus = 0                   'it gets deleted when a new ball is launched
    TargetBonus = 0
    RampBonus = 0
    OrbitHits = 0
    MonsterHits = 0
    MonsterTotal = 15
    bActReady = False
    LordStrength = 100
    LordLifeStep = 1
    LordLifeInterval = 10000
    BumperValue(1) = 150
    BumperValue(2) = 150
    BumperValue(3) = 150
    BumperValue(4) = 150
    bBumperFrenzy = False
    SecondRound = 1
    bJackpot = False
    ComboValue = 0
    'Init Delays/Timers
    'PlayQuote.Enabled = 1
    'MainMode Init()
    For i = 0 to 10
        Mode(i) = 0
    Next
'Init lights
End Sub

Sub StopEndOfBallMode()              'this sub is called after the last ball is drained
    ResetSkillShotTimer_Timer
    If Mode(0) Then StopMode Mode(0) 'a mode is active so stop it
    Gate1.Open = False               'close the top gates
    Gate2.Open = False
    StopBumperFrenzy
    ResetChestDoor
    MalthaelAttackTimer.Enabled = 0
    StopJackpots
End Sub

Sub ResetNewBallVariables() 'reset variables for a new ball or player
    LaneBonus = 0
    TargetBonus = 0
    RampBonus = 0
    OrbitHits = 0
    MalthaelAttackTimer.Enabled = 1
End Sub

Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released
    'top Level lights
    l20.State = 0
    l21.State = 0
    l22.State = 0
    SetLightColor bumperLight1, white, 0
    SetLightColor bumperLight1a, white, 0
    SetLightColor bumperLight2, white, 0
    SetLightColor bumperLight2a, white, 0
    SetLightColor bumperLight3, white, 0
    SetLightColor bumperLight3a, white, 0
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub UpdateSkillShot() 'Updates the skillshot light
    LightSeqSkillshot.Play SeqAllOff
    l21.State = 2
'DMDFlush
End Sub

Sub SkillshotOff_Hit 'trigger to stop the skillshot due to a weak plunger shot
    If bSkillShotReady Then
        ResetSkillShotTimer_Timer
    End If
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    LightSeqSkillshot.StopPlay
    If l21.State = 2 Then l21.State = 0
    SetLightColor l34, blue, 2 'enable the start act/battle at the chest after the skillshot
    bActReady = True
'DMDScoreNow
End Sub

'**********
' Jackpots
'**********
' One of the red arrows blinks, hit it for a Jackpot

Sub StartJackpots
    JackpotTimerExpired.Enabled = 1
    ' always start the jackpot with the left ramp & right ramp
    SetLightColor l43, red, 2
    SetLightColor l45, red, 2
    bJackpot = true
End Sub

Sub StopJackpots
    JackpotTimerExpired.Enabled = 0
    l42.State = 0
    l43.State = 0
    l45.State = 0
    l46.State = 0
    bJackpot = False
End Sub

Sub JackpotTimerExpired_Timer
    StopJackpots
End Sub

' *********************************************************************
'                        Table Object Hit Events
'
' Any target hit Sub will follow this:
' - play a sound
' - do some physical movement
' - add a score, bonus
' - check some variables/Mode this trigger is a member of
' - set the "LastSwitchHit" variable in case it is needed later
' *********************************************************************

' Slingshots has been hit

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAtVol SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors), lemk, 1
    DOF 104, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    DragonLS.RotZ = -15
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    ' add some effect to the table?
    'FlashForMs l20, 1000, 50, 0:FlashForMs l20f, 1000, 50, 0
    'FlashForMs l21, 1000, 50, 0:FlashForMs l21f, 1000, 50, 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14:dragonLS.RotZ = -10
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2:dragonLS.RotZ = -5
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:dragonLS.RotZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAtVol SoundFXDOF("fx_slingshot", 105, DOFPulse, DOFContactors), remk, 1
    DOF 106, DOFPulse
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    DragonRS.RotZ = -15
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    ' add some effect to the table?
    'FlashForMs l22, 1000, 50, 0:FlashForMs l22f, 1000, 50, 0
    'FlashForMs l23, 1000, 50, 0:FlashForMs l23f, 1000, 50, 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14:dragonRS.RotZ = -10
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2:dragonRS.RotZ = -5
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:dragonRS.RotZ = 0:RightSlingShot.TimerEnabled = 0
    End Select

    RStep = RStep + 1
End Sub

'*********
' Bumpers
'*********

Sub Bumper1_Hit
    If NOT Tilted Then
        PlaySoundAtVol SoundFXDOF("fx_Bumper",107,DOFPulse,DOFContactors), Bumper1, VolBump
    DOF 129, DOFPulse
        PlaySound "di_sword1"
        ' add some points
        If bBumperFrenzy Then
            AddScore BumperValue(CurrentPLayer) * 50
        Else
            AddScore BumperValue(CurrentPLayer) + 1000 * BumperLight1.State ' during the Butcher's fight they will score double (the lights are blinking, so the light.State has a value of 2)
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper1"
        If Mode(0) = 2 Then
            LordStrength = LordStrength - 1 - Level(CurrentPLayer)
            FlashEffect 2
      DOF 154, DOFPulse
            CheckLordStrength
            ' Restart the Lords life restore timer
            LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
            RestartLordLifeTimer
        End If
    End If
End Sub

Sub Bumper2_Hit
    If NOT Tilted Then
        PlaySoundAtVol SoundFXDOF("fx_Bumper",108,DOFPulse,DOFContactors), Bumper2, VolBump
    DOF 130, DOFPulse
        PlaySound "di_sword2"
        ' add some points
        If bBumperFrenzy Then
            AddScore BumperValue(CurrentPLayer) * 50
        Else
            AddScore BumperValue(CurrentPLayer) + 1000 * BumperLight1.State ' during the Butcher's fight they will score double (the lights are blinking, so the light.State has a value of 2)
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper2"
        If Mode(0) = 2 Then
            LordStrength = LordStrength - 1 - Level(CurrentPLayer)
            FlashEffect 2
      DOF 154, DOFPulse
            CheckLordStrength
            ' Restart the Lords life restore timer
            LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
            RestartLordLifeTimer
        End If
    End If
End Sub

Sub Bumper3_Hit
    If NOT Tilted Then
        PlaySoundAtVol SoundFXDOF("fx_Bumper",109,DOFPulse,DOFContactors), Bumper3, VolBump
    DOF 131, DOFPulse
        PlaySound "di_sword3"
        ' add some points
        If bBumperFrenzy Then
            AddScore BumperValue(CurrentPLayer) * 50
        Else
            AddScore BumperValue(CurrentPLayer) + 1000 * BumperLight1.State ' during the Butcher's fight they will score double (the lights are blinking, so the light.State has a value of 2)
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper3"
        If Mode(0) = 2 Then
            LordStrength = LordStrength - 1 - Level(CurrentPLayer)
            FlashEffect 2
      DOF 154, DOFPulse
            CheckLordStrength
            ' Restart the Lords life restore timer
            LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
            RestartLordLifeTimer
        End If
    End If
End Sub

Sub Bumper4_Hit
    If NOT Tilted Then
        PlaySoundAtVol "fx_Bumper", Bumper4, VolBump
        PlaySound "di_sword3"
        ' add some points
        If bBumperFrenzy Then
            AddScore BumperValue(CurrentPLayer) * 50
        Else
            AddScore BumperValue(CurrentPLayer) + 1000 * BumperLight1.State ' during the Butcher's fight they will score double (the lights are blinking, so the light.State has a value of 2)
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper4"
        If Mode(0) = 2 Then
            LordStrength = LordStrength - 1 - Level(CurrentPLayer)
            FlashEffect 2
      DOF 154, DOFPulse
            CheckLordStrength
            ' Restart the Lords life restore timer
            LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
            RestartLordLifeTimer
        End If
    End If
End Sub

Sub UpdateBumperLights
    SetLightColor bumperLight1, white, 0
    SetLightColor bumperLight1a, white, 0
    SetLightColor bumperLight2, white, 0
    SetLightColor bumperLight2a, white, 0
    SetLightColor bumperLight3, white, 0
    SetLightColor bumperLight3a, white, 0
    bumperLight1.State = l20.State
    bumperLight2.State = l21.State
    bumperLight3.State = l22.State
    bumperLight1a.State = l20.State
    bumperLight2a.State = l21.State
    bumperLight3a.State = l22.State
End Sub

Sub StartBumperFrenzy
    bBumperFrenzy = True
    BumperFrenzyTimer.Enabled = 1
    ' turn on the bumper lights
    SetLightColor bumperLight1, white, 2
    SetLightColor bumperLight1a, white, 2
    SetLightColor bumperLight2, white, 2
    SetLightColor bumperLight2a, white, 2
    SetLightColor bumperLight3, white, 2
    SetLightColor bumperLight3a, white, 2
    ' Start the rainbow lights on the bumpers
    StartRainbow aBumperLights
End Sub

Sub StopBumperFrenzy
    Dim ii
    bBumperFrenzy = False
    ' turn off the timer
    BumperFrenzyTimer.Enabled = 0
    ' restore the bumper lights
    StopRainbow
    For each ii in aBumperLights
        SetLightColor ii, White, 0
    Next

    UpdateBumperLights
End Sub

Sub BumperFrenzyTimer_Timer ' the timer is up, so stop the mode
    StopBumperFrenzy
End Sub

'***********
' Level Lanes
'***********

Sub triggertop1_Hit
    ' PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
    PlaySoundAtVol "fx_sensor", triggertop1, 1
  DOF 122, DOFPulse
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    If bskillshotReady Then ResetSkillShotTimer_Timer
    l20.State = 1
    If(Mode(0) <> 2) AND(bBumperFrenzy = False) Then UpdateBumperLights
    AddScore 2870
    ' Do some sound or light effect: turn on bumper 1 lights
    If Mode(0) <> 2 Then 'during the Butcher's fight the bumpers are blinking
        FlashForms BumperLight1, 1000, 40, 1:FlashForms BumperLight1a, 1000, 40, 1
    DOF 138, DOFPulse
    End If
    LastSwitchHit = "triggertop1"
    ' do some check
    CheckLevelLanes
End Sub

Sub triggertop2_Hit 'this is the skillshot lane
    ' PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
    PlaySoundAtVol "fx_sensor", triggertop2, 1
  DOF 123, DOFPulse
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    If bskillshotReady Then
        AwardSkillShot
    End If
    l21.State = 1
    If(Mode(0) <> 2) AND(bBumperFrenzy = False) Then UpdateBumperLights
    AddScore 2870 'normal score
    ' Do some sound or light effect: turn on bumper 2 lights
    If Mode(0) <> 2 Then 'during the Butcher's fight the bumpers are blinking
        FlashForms BumperLight2, 1000, 40, 1:FlashForms BumperLight2a, 1000, 40, 1
    DOF 139, DOFPulse
    End If
    LastSwitchHit = "triggertop2"
    ' do some check
    CheckLevelLanes
End Sub

Sub triggertop3_Hit
    ' PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
    PlaySoundAtVol "fx_sensor", triggertop3, 1
  DOF 124, DOFPulse
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    If bskillshotReady Then ResetSkillShotTimer_Timer
    l22.State = 1
    If(Mode(0) <> 2) AND(bBumperFrenzy = False) Then UpdateBumperLights
    AddScore 2870
    ' Do some sound or light effect: turn on bumper 3 lights
    If Mode(0) <> 2 Then 'during the Butcher's fight the bumpers are blinking
        FlashForms BumperLight3, 1000, 40, 1:FlashForms BumperLight3a, 1000, 40, 1
    DOF 140, DOFPulse
    End If
    LastSwitchHit = "triggertop3"
    ' do some check
    CheckLevelLanes
End Sub

Sub CheckLevelLanes() 'use the lane lights
    If l20.State + l21.State + l22.State = 3 Then
        ' turn off only the Level lights, not the bumpers
        l20.State = 0
        l21.State = 0
        l22.State = 0
        'up one level to the current player
        Level(CurrentPlayer) = Level(CurrentPlayer) + 1
        'Increase bonus multiplier
        AddBonusMultiplier 1
        'flash and Light effect
        LightEffect 3
        FlashEffect 3
        'DMD
        DMD "levelup.mp4", "", "", 3000
        ' open the gates and enable the orbit shots
        Gate1.Open = true:Gate2.Open = True
        OrbitCloseTimer.Enabled = 1
    End If
End Sub

Sub OrbitCloseTimer_Timer
    Gate1.Open = False:Gate2.Open = False
    OrbitCloseTimer.Enabled = 0
End Sub

'**********************
' Flipper Lanes: QUEST
'**********************

Sub sw1_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
  DOF 132, DOFPulse
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    l15.State = 1
    AddScore 50050
    ' Do some sound or light effect
    ' do some check
    CheckQUESTLanes
End Sub

Sub sw2_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
  DOF 133, DOFPulse
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    l16.State = 1
    AddScore 10010
    ' Do some sound or light effect
    ' do some check
    CheckQUESTLanes
End Sub

Sub sw3_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
  DOF 133, DOFPulse
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    l17.State = 1
    AddScore 10010
    ' Do some sound or light effect
    ' do some check
    CheckQUESTLanes
End Sub

Sub sw4_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
  DOF 134, DOFPulse
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    l18.State = 1
    AddScore 10010
    ' Do some sound or light effect
    ' do some check
    CheckQUESTLanes
End Sub

Sub sw5_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
  DOF 135, DOFPulse
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    l19.State = 1
    AddScore 50050
    ' Do some sound or light effect
    ' do some check
    CheckQUESTLanes
End Sub

Sub CheckQUESTLanes() 'use the lane lights
    If l15.State + l16.State + l17.State + l18.State + l19.State = 5 Then
        DMD "quest-lights.mp4", "", "", 4000
        'Increase bonus multiplier
        AddBonusMultiplier 1
        'flash and Light effect
        l15.State = 0
        l16.State = 0
        l17.State = 0
        l18.State = 0
        l19.State = 0
        LightEffect 1
        GIEffect 1
    DOF 117, DOFPulse
        FlashEffect 1
    DOF 155, DOFPulse
        DropChestDoor  'open the door to the chest
        AddMultiball 1 'add a multiball
    End If
End Sub

'*****************
'  BASH Targets
'*****************

Sub dt1_Hit
    PlaySoundAtVol SoundFXDOF("fx_target",118,DOFPulse,DOFTargets), ActiveBall, 1
    PlaySound "bash"
    If Tilted Then Exit Sub
    Select Case Mode(0)
        Case 4                    ' Belial
            If l28.State = 1 Then 'the light is lit so hit Belial
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                GiEffect 2
        DOF 117, DOFPulse
                CheckLordStrength
                ' Restart the Lords life restore timer
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
            Else ' the light is off so give back life to Belial
                LordStrength = LordStrength + 5
                If LordStrength> 100 Then
                    LordStrength = 100
                End If
            End If
        Case Else
            l28.State = 1
            ' do some check
            CheckBASHTargets
    End Select
    AddScore 25010
    ' Do some sound or light effect
    LightEffect 1
    LastSwitchHit = "dt1"
End Sub

Sub dt2_Hit
    PlaySoundAtVol SoundFXDOF("fx_target",118,DOFPulse,DOFTargets), ActiveBall, 1
    PlaySound "bash"
    If Tilted Then Exit Sub
    Select Case Mode(0)
        Case 4                    ' Belial
            If l29.State = 1 Then 'the light is lit so hit Belial
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                GiEffect 2
        DOF 117, DOFPulse
                CheckLordStrength
                ' Restart the Lords life restore timer
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
            Else ' the light is off so give back life to Belial
                LordStrength = LordStrength + 5
                If LordStrength> 100 Then
                    LordStrength = 100
                End If
            End If
        Case Else
            l29.State = 1
            ' do some check
            CheckBASHTargets
    End Select
    AddScore 25010
    ' Do some sound or light effect
    LightEffect 1
    LastSwitchHit = "dt2"
End Sub

Sub dt3_Hit
    PlaySoundAtVol SoundFXDOF("fx_target",118,DOFPulse,DOFTargets), ActiveBall, 1
    PlaySound "bash"
    If Tilted Then Exit Sub
    Select Case Mode(0)
        Case 4                    ' Belial
            If l30.State = 1 Then 'the light is lit so hit Belial
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                GiEffect 2
        DOF 117, DOFPulse
                CheckLordStrength
            Else ' the light is off so give back life to Belial
                LordStrength = LordStrength + 5
                If LordStrength> 100 Then
                    LordStrength = 100
                End If
            End If
        Case Else
            l30.State = 1
            ' do some check
            CheckBASHTargets
    End Select
    AddScore 25010
    ' Do some sound or light effect
    LightEffect 1
    LastSwitchHit = "dt3"
End Sub

Sub dt4_Hit
    PlaySoundAtVol SoundFXDOF("fx_target",118,DOFPulse,DOFTargets), ActiveBall, 1
    PlaySound "bash"
    If Tilted Then Exit Sub
    Select Case Mode(0)
        Case 4                    ' Belial
            If l31.State = 1 Then 'the light is lit so hit Belial
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                GiEffect 2
        DOF 117, DOFPulse
                CheckLordStrength
                ' Restart the Lords life restore timer
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
            Else ' the light is off so give back life to Belial
                LordStrength = LordStrength + 5
                If LordStrength> 100 Then
                    LordStrength = 100
                End If
            End If
        Case Else
            l31.State = 1
            ' do some check
            CheckBASHTargets
    End Select
    AddScore 25010
    ' Do some sound or light effect
    LightEffect 1
    LastSwitchHit = "dt4"
End Sub

Sub CheckBASHTargets
    If l28.State + l29.State + l30.State + l31.State = 4 Then
        FlashEffect 2
    DOF 154, DOFPulse
        ' increment the target bonus
        TargetBonus = TargetBonus + 1
        If TargetBonus = 3 Then 'lit extra ball at the chest
            l25.State = 2
            DMDBlink "black.jpg", " ", "EXTRA BALL IS LIT", 50, 20
        End If
        Select Case Mode(0)
            Case 1, 3, 5, 7, 9
                DMD "bash-lights.mp4", "", "", 4000
                MonsterHits = MonsterHits + 5 + Level(CurrentPlayer)
                CheckActMonsters
                MonsterHitTimer_Timer
                ResetBashLights
            Case 2, 4, 6, 8, 10
                LordStrength = LordStrength -15 + Level(CurrentPlayer)
                CheckLordStrength
                ResetBashLights
        End Select
    End If
End Sub

Sub ResetBashLights
    l28.State = 0
    l29.State = 0
    l30.State = 0
    l31.State = 0
End Sub

'*****************
'  SKILL Targets
'*****************

Sub dt5_Hit
    PlaySoundAtVol SoundFXDOF("fx_target",119,DOFPulse,DOFTargets), ActiveBall, 1
    PlaySound "chains"
    If Tilted Then Exit Sub
    Select Case Mode(0)
        Case 4                    ' Belial
            If l35.State = 1 Then 'the light is lit so hit Belial
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                GiEffect 2
        DOF 117, DOFPulse
                CheckLordStrength
                ' Restart the Lords life restore timer
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
            Else ' the light is off so give back life to Belial
                LordStrength = LordStrength + 5
                If LordStrength> 100 Then
                    LordStrength = 100
                End If
            End If
        Case Else
            l35.State = 1
            ' do some check
            CheckSKILLTargets
    End Select
    AddScore 25010
    ' Do some sound or light effect
    LightEffect 1
    LastSwitchHit = "dt5"
End Sub

Sub dt6_Hit
    PlaySoundAtVol SoundFXDOF("fx_target",119,DOFPulse,DOFTargets), ActiveBall, 1
    PlaySound "chains"
    If Tilted Then Exit Sub
    Select Case Mode(0)
        Case 4                    ' Belial
            If l36.State = 1 Then 'the light is lit so hit Belial
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                GiEffect 2
        DOF 117, DOFPulse
                CheckLordStrength
                ' Restart the Lords life restore timer
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
            Else ' the light is off so give back life to Belial
                LordStrength = LordStrength + 5
                If LordStrength> 100 Then
                    LordStrength = 100
                End If
            End If
        Case Else
            l36.State = 1
            ' do some check
            CheckSKILLTargets
    End Select
    AddScore 25010
    ' Do some sound or light effect
    LightEffect 1
    LastSwitchHit = "dt6"
End Sub

Sub dt7_Hit
    PlaySoundAtVol SoundFXDOF("fx_target",119,DOFPulse,DOFTargets), ActiveBall, 1
    PlaySound "chains"
    If Tilted Then Exit Sub
    Select Case Mode(0)
        Case 4                    ' Belial
            If l37.State = 1 Then 'the light is lit so hit Belial
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                GiEffect 2
        DOF 117, DOFPulse
                CheckLordStrength
                ' Restart the Lords life restore timer
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
            Else ' the light is off so give back life to Belial
                LordStrength = LordStrength + 5
                If LordStrength> 100 Then
                    LordStrength = 100
                End If
            End If
        Case Else
            l37.State = 1
            ' do some check
            CheckSKILLTargets
    End Select
    AddScore 25010
    ' Do some sound or light effect
    LightEffect 1
    LastSwitchHit = "dt7"
End Sub

Sub dt8_Hit
    PlaySoundAtVol SoundFXDOF("fx_target",120,DOFPulse,DOFTargets), ActiveBall, 1
    PlaySound "chains"
    If Tilted Then Exit Sub
    Select Case Mode(0)
        Case 4                    ' Belial
            If l40.State = 1 Then 'the light is lit so hit Belial
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                GiEffect 2
        DOF 117, DOFPulse
                CheckLordStrength
                ' Restart the Lords life restore timer
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
            Else ' the light is off so give back some life to Belial
                LordStrength = LordStrength + 5
                If LordStrength> 100 Then
                    LordStrength = 100
                End If
            End If
        Case Else
            l40.State = 1
            ' do some check
            CheckSKILLTargets
    End Select
    AddScore 25010
    ' Do some sound or light effect
    LightEffect 1
    LastSwitchHit = "dt8"
End Sub

Sub dt9_Hit
    PlaySoundAtVol SoundFXDOF("fx_target",120,DOFPulse,DOFTargets), ActiveBall, 1
    PlaySound "chains"
    If Tilted Then Exit Sub
    Select Case Mode(0)
        Case 4                    ' Belial
            If l41.State = 1 Then 'the light is lit so hit Belial
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                GiEffect 2
        DOF 117, DOFPulse
                CheckLordStrength
                ' Restart the Lords life restore timer
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
            Else ' the light is off so give back life to Belial
                LordStrength = LordStrength + 5
                If LordStrength> 100 Then
                    LordStrength = 100
                End If
            End If
        Case Else
            l41.State = 1
            ' do some check
            CheckSKILLTargets
    End Select
    AddScore 25010
    ' Do some sound or light effect
    LightEffect 1
    LastSwitchHit = "dt9"
End Sub

Sub CheckSKILLTargets
    If l35.State + l36.State + l37.State + l40.State + l41.State = 5 Then
        FlashEffect 2
    DOF 154, DOFPulse
        ' increment the target bonus
        TargetBonus = TargetBonus + 1
        If TargetBonus = 3 Then 'lit extra ball at the chest
            l25.State = 2
            DMDBlink "black.jpg", " ", "EXTRA BALL IS LIT", 50, 20
        End If
        Select Case Mode(0)
            Case 1, 3, 5, 7, 9
                DMD "skill-lights.mp4", "", "", 3000
                MonsterHits = MonsterHits + 7 + Level(CurrentPlayer)
                CheckActMonsters
                MonsterHitTimer_Timer
                ResetSkillLights
            Case 2, 4, 6, 8, 10
                LordStrength = LordStrength -20 + Level(CurrentPlayer)
                CheckLordStrength
                ResetSkillLights
        End Select
    End If
End Sub

Sub ResetSkillLights
    l35.State = 0
    l36.State = 0
    l37.State = 0
    l40.State = 0
    l41.State = 0
End Sub

'*******************
' Resplendent Chest
'*******************

Sub ChestDoor_Hit
    PlaySoundAtVol SoundFXDOF("fx_droptarget",116,DOFPulse,DOFContactors), ActiveBall, 1
    PlaySound "centershot"
    If Tilted Then Exit Sub
    DropChestDoor
End Sub

Sub DropChestDoor
    ChestDoor.IsDropped = 1
  DOF 116, DOFPulse
    'Play a sound, do some light/flash effect
    GiEffect 1
  DOF 117, DOFPulse
End Sub

Sub ResetChestDoor
    ChestDoor.IsDropped = 0
  DOF 116, DOFPulse
    PlaySound "fx_resetdrop" ' TODO
End Sub

Sub ChestHole_Hit
  DOF 128, DOFOn
    PlaySoundAtVol "fx_kicker_enter", ChestHole, 1
    ResetChestDoor
    If NOT Tilted Then
        FlashForMs f16, 2000, 50, 0
    DOF 136, DOFPulse
        If Mode(0) = 0 Then 'we could check here too for the variable bActReady
            StartNextMode
            vpmtimer.addtimer 6000, "ChestExit '"
        Else
            If Mode(0) = 11 Then 'the end
                WinMalthael2
                vpmtimer.addtimer 80000, "ChestExit '"
            Else
                GiveRandomAward
                vpmtimer.addtimer 4400, "ChestExit '"
            End If
        End If
        If l25.State = 2 Then 'give the extra ball
            l25.State = 0
            AwardExtraBall
        End If
        If bJackpot AND l44.State = 2 Then 'give the SuperJackpot & enable the next jackpot
            l44.State = 0
            SetLightColor l43, red, 2
            SetLightColor l45, red, 2
            AwardSuperJackpot
        End If
    Else 'if tilted then just exit the ball
        vpmtimer.addtimer 100, "ChestExit '"
    End If
End Sub

Sub ChestExit()
  DOF 128, DOFOff
  If B2SOn Then Controller.B2SSetData bsnr,1:Controller.B2SSetData 20,0:Controller.B2SSetData 80,1
    FlashForMs f16, 1000, 50, 0
  DOF 137, DOFPulse
    ChestHole.DestroyBall
    ChestOut.CreateBall
    PlaySoundAtVol SoundFXDOF("fx_popper",115,DOFPulse,DOFContactors), ChestHole, VolKick
  DOF 112, DOFPulse
    ChestOut.kick 65, 30, 1.56
End Sub

Sub GiveRandomAward() 'from the Chest and the Orbits
    Dim tmp, tmp2
    PlaySound "gold"
    ' show some random values on the dmd
    DMD "black.jpg", " ", "EXTRA POINTS", 20
    DMD "black.jpg", " ", "EXTRA BALL IS LIT", 20
    DMD "black.jpg", " ", "EXTRA POINTS", 20
    DMD "black.jpg", " ", "BONUS MULTIPLIER", 20
    DMD "black.jpg", " ", "EXTRA POINTS", 20
    DMD "black.jpg", " ", "LEVEL UP", 20
    DMD "black.jpg", " ", "EXTRA POINTS", 20
    DMD "black.jpg", " ", "BUMPER VALUE", 20
    DMD "black.jpg", " ", "LEVEL UP", 20
    DMD "black.jpg", " ", "EXTRA POINTS", 20
    DMD "black.jpg", " ", "EXTRA BALL IS LIT", 20
    DMD "black.jpg", " ", "EXTRA POINTS", 20
    DMD "black.jpg", " ", "BONUS MULTIPLIER", 20
    DMD "black.jpg", " ", "EXTRA POINTS", 20
    DMD "black.jpg", " ", "LEVEL UP", 20
    DMD "black.jpg", " ", "EXTRA POINTS", 20
    DMD "black.jpg", " ", "BUMPER VALUE", 20
    DMD "black.jpg", " ", "LEVEL UP", 20
    DMD "black.jpg", " ", " ", 200

    tmp = INT(RND(1) * 80)
    Select Case tmp
        Case 1         'Light Extra Ball
            l23.State = 2:DMDBlink "black.jpg", " ", "EXTRA BALL IS LIT", 50, 20
        Case 2         'Light Extra Ball
            l24.State = 2:DMDBlink "black.jpg", " ", "EXTRA BALL IS LIT", 50, 20
        Case 3         'Light Extra Ball
            l25.State = 2:DMDBlink "black.jpg", " ", "EXTRA BALL IS LIT", 50, 20
        Case 4         'Light Extra Ball
            l26.State = 2:DMDBlink "black.jpg", " ", "EXTRA BALL IS LIT", 50, 20
        Case 5         'Light Extra Ball
            l27.State = 2:DMDBlink "black.jpg", " ", "EXTRA BALL IS LIT", 50, 20
        Case 6, 41, 42 'Start Bumper frenzy, where each bumper hit will award 50x the bumper value. The bumpers will lit blue during this mode.
            DMDBlink "black.jpg", " ", "BUMPER FRENZY", 50, 20
            StartBumperFrenzy
        Case 7, 8 '2,000,000 points
            DMDBlink "black.jpg", "BIG POINTS", "2000000", 50, 20
            AddScore 2000000
        Case 9, 10, 11, 12 'Hold Bonus
            DMDBlink "black.jpg", "BONUS HELD", "ACTIVATED", 50, 20
            bBonusHeld = True
        Case 13, 14, 15 '1,000,000 points
            DMDBlink "black.jpg", "BIG POINTS", "1000000", 50, 20
            AddScore 2000000
        Case 16, 17, 18 'Increase Bonus Multiplier
            DMDBlink "black.jpg", "INCREASED", "BONUS MULTIPLIER", 50, 20
            AddBonusMultiplier 1
        Case 19, 20, 21 'Level up 1
            DMDBlink "black.jpg", " ", "1 LEVEL UP", 50, 20
            Level(CurrentPlayer) = Level(CurrentPlayer) + 1
        Case 22, 23 'Level up 2
            DMDBlink "black.jpg", " ", "2 LEVELS UP", 50, 20
            Level(CurrentPlayer) = Level(CurrentPlayer) + 2
        Case 24, 25, 26, 27, 28 '500,000 points
            DMDBlink "black.jpg", "BIG POINTS", "500000", 50, 20
            AddScore 500000
        Case 29, 30, 31, 32, 33, 34, 35, 36, 37, 38 'Increase Bumper value
            BumperValue(CurrentPlayer) = BumperValue(CurrentPlayer) + 300
            DMDBlink "black.jpg", "BUMPER VALUE", BumperValue(CurrentPlayer), 50, 20
        Case 39, 40, 43, 44 'extra multiball
            DMDBlink "black.jpg", "EXTRA", "MULTIBALL", 50, 20
            AddMultiball 1
    Case 45, 46, 47, 48 ' Ball Save
      EnableBallSaver 20
            DMDBlink "black.jpg", "BALL SAVE", "ACTIVATED", 50, 20
        Case 50 'Level up 3
            DMDBlink "black.jpg", " ", "3 LEVELS UP", 50, 20
            Level(CurrentPlayer) = Level(CurrentPlayer) + 3
        Case ELSE 'Add a Random score from 100.000 to 300,000 points
            tmp2 = INT((RND) * 3) * 100000 + 100000
            DMDBlink "black.jpg", "EXTRA POINTS", tmp2, 50, 20
            AddScore tmp2
    End Select
End Sub

'*****************
'   The Orbit
'*****************

Sub OrbitTrigger1_Hit
    If Tilted Then Exit Sub
    If LastSwitchHit = "OrbitTrigger2" Then 'this is a completed orbit
        OrbitHits = OrbitHits + 1
        GiveRandomAward
    Else
        If l32.State = 2 Then
            Select Case Mode(0)
                Case 1, 3, 5, 7, 9
                    GiEffect 2
          DOF 125, DOFPulse
                    DMD "leftloop.mp4", "", "", 2000
                    Monsters(CurrentPlayer) = Monsters(CurrentPlayer) + 1 'bonus
                    MonsterHits = MonsterHits + 1 + Level(CurrentPlayer)
                    CheckActMonsters
                    MonsterHitTimer_Timer
                    UpdateModeLights
                Case 4, 6, 10
                    FlashEffect 2
          DOF 154, DOFPulse
                    LordStrength = LordStrength -10 - Level(CurrentPLayer)
                    CheckLordStrength
                    ' Restart the Lords life restore timerted)
                    LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                    RestartLordLifeTimer
                Case 8 'diablo
                    l32.State = 0:l39.State = 2
                    FlashEffect 2
          DOF 154, DOFPulse
                    LordStrength = LordStrength -10 - Level(CurrentPLayer)
                    CheckLordStrength
                    ' Restart the Lords life restore timerted)
                    LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                    RestartLordLifeTimer
            End Select
        End If
        If bJackpot AND l42.State = 2 Then 'give the Jackpot & enable the next one
            l42.State = 0
            SetLightColor l46, red, 2
            AwardJackpot
        End If
        If l23.State = 2 Then 'give the extra ball
            l23.State = 0
            AwardExtraBall
        End If
    End If
    LastSwitchHit = "OrbitTrigger1"
End Sub

Sub OrbitTrigger2_Hit
    If Tilted Then Exit Sub
    If LastSwitchHit = "OrbitTrigger1" Then 'this is a completed orbit
        OrbitHits = OrbitHits + 1
        GiveRandomAward
    Else
        If l39.State = 2 Then
            Select Case Mode(0)
                Case 1, 3, 5, 7, 9
                    GiEffect 2
          DOF 125, DOFPulse
                    DMD "rightloop.mp4", "", "", 2000
                    Monsters(CurrentPlayer) = Monsters(CurrentPlayer) + 1 'bonus
                    MonsterHits = MonsterHits + 1 + Level(CurrentPlayer)
                    CheckActMonsters
                    MonsterHitTimer_Timer
                    UpdateModeLights
                Case 4, 6, 10
                    FlashEffect 2
          DOF 154, DOFPulse
                    LordStrength = LordStrength -10 - Level(CurrentPLayer)
                    CheckLordStrength
                    ' Restart the Lords life restore timerted)
                    LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                    RestartLordLifeTimer
                Case 8 'diablo
                    l32.State = 2:l39.State = 0
                    FlashEffect 2
          DOF 154, DOFPulse
                    LordStrength = LordStrength -10 - Level(CurrentPLayer)
                    CheckLordStrength
                    ' Restart the Lords life restore timerted)
                    LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                    RestartLordLifeTimer
            End Select
        End If
        If bJackpot AND l46.State = 2 Then 'give the Jackpot & enable the next one
            l46.State = 0
            SetLightColor l44, red, 2
            AwardJackpot
        End If
        If l27.State = 2 Then 'give the extra ball
            l27.State = 0
            AwardExtraBall
        End If
    End If
    LastSwitchHit = "OrbitTrigger2"
End Sub

'****************
'     Ramps
'****************

Sub LeftRampDone_Hit
    PlaySoundAtVol "fx_metalrolling", ActiveBall, 1
    If Tilted Then Exit Sub
    'increase the ramp bonus
    RampBonus = RampBonus + 1
    If l33.State = 2 Then
        Select Case Mode(0)
            Case 1, 3, 5, 7, 9
                GiEffect 2
        DOF 125, DOFPulse
                DMD "leftramp.mp4", "", "", 2000
                Monsters(CurrentPlayer) = Monsters(CurrentPlayer) + 1 'bonus
                MonsterHits = MonsterHits + 1 + Level(CurrentPlayer)
                CheckActMonsters
                MonsterHitTimer_Timer
                UpdateModeLights
            Case 4, 6, 8, 10
                FlashEffect 2
        DOF 154, DOFPulse
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                CheckLordStrength
                ' Restart the Lords life restore timerted)
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
        End Select
    End If
    If LastSwitchHit = "RightRampDone" Then 'give combo
    DOF 117, DOFPulse
        ComboValue = 100000 + Round(Score(CurrentPlayer) / 10, 0)
        DMDBlink "black.jpg", "COMBO", ComboValue, 100, 10
        AddScore ComboValue
    End If
    If bJackpot AND l43.State = 2 Then 'give the Jackpot
        AwardJackpot
        l43.State = 0
        l45.State = 0
        SetLightColor l42, red, 2
    End If
    If l24.State = 2 Then 'give the extra ball
        l24.State = 0
        AwardExtraBall
    End If
    LastSwitchHit = "LeftRampDone"
End Sub

Sub RightRampDone_Hit
    PlaySoundAtVol "fx_metalrolling", ActiveBall, 1
    If Tilted Then Exit Sub
    'increase the ramp bonus
    RampBonus = RampBonus + 1
    If l38.State = 2 Then
        Select Case Mode(0)
            Case 1, 3, 5, 7, 9
                GiEffect 2
        DOF 125, DOFPulse
                DMD "rightramp.mp4", "", "", 2000
                Monsters(CurrentPlayer) = Monsters(CurrentPlayer) + 1 'bonus
                MonsterHits = MonsterHits + 1 + Level(CurrentPlayer)
                CheckActMonsters
                MonsterHitTimer_Timer
                UpdateModeLights
            Case 4, 6, 8, 10
                FlashEffect 2
        DOF 154, DOFPulse
                LordStrength = LordStrength -10 - Level(CurrentPLayer)
                CheckLordStrength
                ' Restart the Lords life restore timerted)
                LordLifeTimer.Interval = LordLifeInterval + Level(CurrentPlayer) * 200
                RestartLordLifeTimer
        End Select
    End If
    If LastSwitchHit = "LeftRampDone" Then 'give combo
    DOF 117, DOFPulse
        ComboValue = 100000 + Round(Score(CurrentPlayer) / 10, 0)
        DMDBlink "black.jpg", "COMBO", ComboValue, 100, 10
        AddScore ComboValue
    End If
    If bJackpot AND l45.State = 2 Then 'give the Jackpot
        AwardJackpot
        l43.State = 0
        l45.State = 0
        SetLightColor l42, red, 2
    End If
    If l26.State = 2 Then 'give the extra ball
        l26.State = 0
        AwardExtraBall
    End If
    LastSwitchHit = "RightRampDone"
End Sub

'************************
'  Mode: Acts & Battles
'************************

' Mode are played one after each other
' This is a cooperative table where each player will continue right after where the last player finished
' If you loose the ball you'll need to restart the mode :)
'

Dim Mode(10) 'this is for easier programming

' Mode(0) will have the current mode number
' when a mode is not active Mode(n) = 0
' when a mode is active Mode(n) = 2
' when a mode is completed Mode(n) = 1
' this is to be consistent with the light states

' Only one mode can be active at a time.

' Mode(1): Act 1
' Mode(2): end of Act 1: Battle the Butcher
' Mode(3): Act 2
' Mode(4): end of Act 2: Battle Belial
' Mode(5): Act 3
' Mode(6): end of Act 3: Battle Azmodan
' Mode(7): Act 4
' Mode(8): end of Act 4: Battle Diablo
' Mode(9): Act 5
' Mode(10): end of Act 5: Battle Malthael

Sub StartNextMode
    ' stop the start act lights & variables
    l34.State = 0
    l44.State = 0
    bActReady = False

    If Mode(1) = 0 Then
        StartMode 1
    ElseIf Mode(2) = 0 Then
        StartMode 2
    ElseIf Mode(3) = 0 Then
        StartMode 3
    ElseIf Mode(4) = 0 Then
        StartMode 4
    ElseIf Mode(5) = 0 Then
        StartMode 5
    ElseIf Mode(6) = 0 Then
        StartMode 6
    ElseIf Mode(7) = 0 Then
        StartMode 7
    ElseIf Mode(8) = 0 Then
        StartMode 8
    ElseIf Mode(9) = 0 Then
        StartMode 9
    ElseIf Mode(10) = 0 Then
        StartMode 10
    End If
End Sub

Sub StartMode(n)
    l34.State = 0
    DMDFlush
    Select Case n
        Case 1:StartAct1
        Case 2:StartButcher
        Case 3:StartAct2
        Case 4:StartBelial
        Case 5:StartAct3
        Case 6:StartAzmodan
        Case 7:StartAct4
        Case 8:StartDiablo
        Case 9:StartAct5
        Case 10:StartMalthael
    End Select
  If B2sOn Then
    Controller.B2SSetData 80,0:Controller.B2SSetData 20,1
    for ao=1 to 11:Controller.B2SSetData ao,0:next
    bsnr = n
  End If
End Sub

Sub StopMode(n) 'called at the end of a ball
    l34.State = 2
    Select Case n
        Case 1:StopAct1
        Case 2:StopButcher
        Case 3:StopAct2
        Case 4:StopBelial
        Case 5:StopAct3
        Case 6:StopAzmodan
        Case 7:StopAct4
        Case 8:StopDiablo
        Case 9:StopAct5
        Case 10:StopMalthael
    End Select
  If B2sOn Then
    Controller.B2SSetData bsnr,0
  End If
End Sub

Sub WinMode(n) 'called after completing a mode
    DMDFlush
    Select Case n
        Case 1:WinAct1
        Case 2:WinButcher
        Case 3:WinAct2
        Case 4:WinBelial
        Case 5:WinAct3
        Case 6:WinAzmodan
        Case 7:WinAct4
        Case 8:WinDiablo
        Case 9:WinAct5
        Case 10:WinMalthael
    End Select
End Sub

Sub UpdateModeLights()
    'update the lights according to the mode state
    l7.State = Mode(1)
    l6.State = Mode(2)
    l8.State = Mode(3)
    l9.State = Mode(4)
    l10.State = Mode(5)
    l11.State = Mode(6)
    l12.State = Mode(7)
    l14.State = Mode(8)
    l13.State = Mode(9)
    l47.State = Mode(10)
    'change the color of the active mode depending of how far the act or battle is finished
    Select Case Mode(0)
        Case 1:SetLightColor l7, ColorCode, 2
        Case 2:SetLightColor l6, ColorCode, 2
        Case 3:SetLightColor l8, ColorCode, 2
        Case 4:SetLightColor l9, ColorCode, 2
        Case 5:SetLightColor l10, ColorCode, 2
        Case 6:SetLightColor l11, ColorCode, 2
        Case 7:SetLightColor l12, ColorCode, 2
        Case 8:SetLightColor l14, ColorCode, 2
        Case 9:SetLightColor l13, ColorCode, 2
        Case 10:SetLightColor l47, ColorCode, 2
    End Select
End Sub

Function ColorCode ' gives the percent completed, this will be a number from 0 to 10
    Select Case Mode(0)
        Case 1, 3, 5, 7, 9
            ColorCode = INT((MonsterHits / MonsterTotal) * 10)
        Case 2, 4, 6, 8, 10
            ColorCode = LordStrength \ 10
    End Select
If ColorCode < 0 Then ColorCode = 0
If ColorCode > 10 Then ColorCode = 10
End Function

'********************
' Monsters Hit Timer
'********************

Dim mStep:mStep = False

Sub MonsterHitTimer_Timer() ' will change the lights differently in each act
    Dim tmp
    Select Case Mode(0)
        Case 1 ' all 4 lights are turned on
            SetLightColor l32, White, 2
            SetLightColor l33, White, 2
            SetLightColor l38, White, 2
            SetLightColor l39, White, 2
        Case 3 ' 3 lights are turned on
            tmp = INT(RND * 4)
            SetLightColor l32, White, 2
            SetLightColor l33, White, 2
            SetLightColor l38, White, 2
            SetLightColor l39, White, 2
            Select Case tmp
                Case 0:l32.State = 0
                Case 1:l33.State = 0
                Case 2:l38.State = 0
                Case 3:l39.State = 0
            End Select
        Case 5 ' 2 lights, alternate between orbits and ramps
            mStep = NOT mStep
            If mStep Then
                SetLightColor l32, White, 2
                SetLightColor l33, White, 0
                SetLightColor l38, White, 0
                SetLightColor l39, White, 2
            Else
                SetLightColor l32, White, 0
                SetLightColor l33, White, 2
                SetLightColor l38, White, 2
                SetLightColor l39, White, 0
            End If
        Case 7 ' left or right lights
            mStep = NOT mStep
            If mStep Then
                SetLightColor l32, White, 2
                SetLightColor l33, White, 2
                SetLightColor l38, White, 0
                SetLightColor l39, White, 0
            Else
                SetLightColor l32, White, 0
                SetLightColor l33, White, 0
                SetLightColor l38, White, 2
                SetLightColor l39, White, 2
            End If
        Case 9 ' 1 random light at a time
            tmp = INT(RND * 4)
            l32.State = 0
            l33.State = 0
            l38.State = 0
            l39.State = 0
            Select Case tmp
                Case 0:SetLightColor l32, White, 2
                Case 1:SetLightColor l33, White, 2
                Case 2:SetLightColor l38, White, 2
                Case 3:SetLightColor l39, White, 2
            End Select
    End Select
End Sub

Sub StartMonsterTimer
    MonsterHitTimer_Timer 'start one of the lights
    MonsterHitTimer.Enabled = 1
End Sub

Sub StopMonsterTimer
    MonsterHitTimer.Enabled = 0
    l32.State = 0
    l33.State = 0
    l38.State = 0
    l39.State = 0
End Sub

Sub CheckActMonsters
    If MonsterHits >= MonsterTotal Then
        WinMode Mode(0)
    End If
    UpdateModeLights()
End Sub

Sub CheckLordStrength
    If LordStrength <= 0 Then
        WinMode Mode(0)
    End If
    UpdateModeLights()
    DMDScore
End Sub

Sub RestartLordLifeTimer()
    LordLifeTimer.Enabled = 0
    LordLifeTimer.Enabled = 1
End Sub

Sub LordLifeTimer_Timer ' increases the Lords strength if not hit
    LordLifeTimer.Interval = LordLifeInterval
    If LordStrength <100 then
        LordStrength = LordStrength + LordLifeStep
        UpdateModeLights()
        DMDScore
    End If
End Sub

'****************
' Mode 1: Act 1
'****************
' All the acts use these 2 variables:
' MonsterHits, this the numbers of monsters to kill to end the Act and start the battle against the Lord
' MonsterTotal, the total number of monsters of the act, or the total life of the Lord during battles

Sub StartAct1
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act1.mp4", " ", "         ", 6500
    ' start music after the intro
    vpmtimer.AddTimer 6000, "PlaySong ""mu_act"" '"
    'setup variables
    MonsterTotal = 10 + INT(RND * 5) ' from 10 to 15 monsters
    MonsterHits = 0
    Mode(0) = 1                      ' this is the active mode
    Mode(1) = 2                      ' this means it is started, and the corresponding light will be turned on and start blinking
    'update lights
    UpdateModeLights()
    ChangeGi white
    ' Start a timer to light a random arrow light and to change it from time to time
    StartMonsterTimer
End Sub

Sub StopAct1
    StopMonsterTimer
    Mode(0) = 0
    Mode(1) = 0
    UpdateModeLights()
    ChangeGi white
End Sub

Sub WinAct1
    StopMonsterTimer
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act-complete.mp4", "", "", 12000
    FlashEffect 2
  DOF 154, DOFPulse
    'setup variables
    Mode(0) = 0
    Mode(1) = 1
    'update lights
    UpdateModeLights()
    ChangeGi white
    SetLightColor l7, purple, 1
    ' lit the boss light
    SetLightColor l44, green, 2
    vpmtimer.AddTimer 12000, "PlaySong ""mu_main"" '"
End Sub

'*********************
' Mode 2: The Butcher
'*********************
' The fights with the Lords use just one variable LordStrength

Sub StartButcher
    StopSound Song:Song = ""
    DMD "butcher-start.mp4", "", "", 12000
    ' start music after the intro
    vpmtimer.AddTimer 12000, "PlaySong ""mu_lordfight"" '"
    vpmtimer.AddTimer 8000, "PlaySound ""battle"" '"
    LordStrength = 100
    Mode(0) = 2
    Mode(2) = 2
    'update lights
    UpdateModeLights()
    ChangeGi red
    ' you need to hit the bumpers to defeat the Butcher so turn them red and blinking
    SetLightColor bumperLight1, red, 2
    SetLightColor bumperLight1a, red, 2
    SetLightColor bumperLight2, red, 2
    SetLightColor bumperLight2a, red, 2
    SetLightColor bumperLight3, red, 2
    SetLightColor bumperLight3a, red, 2
    ' add 1 multiball after the video is finished
    vpmtimer.addtimer 12000, "AddMultiBall 1 '"
    ' Start the life restore timer
    LordLifeStep = 1
    LordLifeInterval = 3000
    LordLifeTimer.Interval = LordLifeInterval
    RestartLordLifeTimer
    StartJackpots
End Sub

Sub StopButcher
    Mode(0) = 0
    Mode(2) = 0
    'update lights
    SetLightColor bumperLight1, white, 0
    SetLightColor bumperLight1a, white, 0
    SetLightColor bumperLight2, white, 0
    SetLightColor bumperLight2a, white, 0
    SetLightColor bumperLight3, white, 0
    SetLightColor bumperLight3a, white, 0
    UpdateModeLights()
    SetLightColor l6, white, 0
    ChangeGi white
End Sub

Sub WinButcher
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "butcher-defeated.mp4", "", "", 14000
    FlashEffect 2
  DOF 154, DOFPulse
    LightEffect 2
    GiEffect 2
  DOF 125, DOFPulse
    'setup variables
    Mode(0) = 0
    Mode(2) = 1
    'update lights
    SetLightColor bumperLight1, white, 0
    SetLightColor bumperLight1a, white, 0
    SetLightColor bumperLight2, white, 0
    SetLightColor bumperLight2a, white, 0
    SetLightColor bumperLight3, white, 0
    SetLightColor bumperLight3a, white, 0
    UpdateModeLights()
    SetLightColor l6, purple, 1
    ' lit the start act light to show the player that he can start the next mode
    SetLightColor l34, blue, 2
    ChangeGi white
    vpmtimer.AddTimer 14000, "PlaySong ""mu_main"" '"
End Sub

'****************
' Mode 3: Act 2
'****************
' All the acts use these 2 variables:
' MonsterHits, this the numbers of monsters to kill to end the Act and start the battle against the Lord
' MonsterTotal, the total number of monsters of the act, or the total life of the Lord during battles

Sub StartAct2
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act2.mp4", " ", "         ", 6500
    ' start music after the intro
    vpmtimer.AddTimer 6000, "PlaySong ""mu_act2"" '"
    'setup variables
    MonsterTotal = 15 + INT(RND * 5) ' from 15 to 20 monsters
    MonsterHits = 0
    Mode(0) = 3                      ' this is the active mode
    Mode(3) = 2                      ' this means it is started, and the corresponding light will be turned on and start blinking
    'update lights
    UpdateModeLights()
    ChangeGi white
    ' Start a timer to light a random arrow light and to change it from time to time
    StartMonsterTimer
End Sub

Sub StopAct2
    StopMonsterTimer
    Mode(0) = 0
    Mode(3) = 0
    UpdateModeLights()
    ChangeGi white
End Sub

Sub WinAct2
    StopMonsterTimer
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act-complete.mp4", "", "", 12000
    FlashEffect 2
  DOF 154, DOFPulse
    'setup variables
    Mode(0) = 0
    Mode(3) = 1
    'update lights
    UpdateModeLights()
    ChangeGi white
    SetLightColor l8, purple, 1
    ' lit the boss light
    SetLightColor l44, green, 2
    vpmtimer.AddTimer 12000, "PlaySong ""mu_main"" '"
End Sub

'**********************
' Mode 4: Battle Belial
'**********************
' Belial hides behind the targets.
' Hit the lit targets to kill him
' Do not hit the the off targets, they will increase the life of Belial
Dim TargetsPos

Sub StartBelial
    StopSound Song:Song = ""
    DMD "belial-start.mp4", "", "", 12000
    ' start music after the intro
    vpmtimer.AddTimer 12000, "PlaySong ""mu_lordfight"" '"
    vpmtimer.AddTimer 8000, "PlaySound ""battle"" '"
    LordStrength = 100
    Mode(0) = 4
    Mode(4) = 2
    'update lights
    UpdateModeLights()
    ChangeGi yellow
    ' you need to hit the lit targets to defeat Belial so turn start
    TargetsPos = 0
    BelialTargets.Enabled = 1
    ' add 2 multiball after the video is finished
    vpmtimer.addtimer 12000, "AddMultiBall 2 '"
    ' Start the life restore timer
    LordLifeStep = 2
    LordLifeInterval = 2500
    LordLifeTimer.Interval = LordLifeInterval
    RestartLordLifeTimer
    StartJackpots
End Sub

Sub StopBelial
    Mode(0) = 0
    Mode(4) = 0
    BelialTargets.Enabled = 0
    UpdateModeLights()
    l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 0
    ChangeGi white
End Sub

Sub WinBelial
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "belial-defeated.mp4", "", "", 14000
    FlashEffect 2
  DOF 154, DOFPulse
    LightEffect 2
    GiEffect 2
  DOF 125, DOFPulse
    'setup variables
    Mode(0) = 0
    Mode(4) = 1
    'update lights
    BelialTargets.Enabled = 0
    UpdateModeLights()
    l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 0
    ChangeGi white
    SetLightColor l9, purple, 1
    ' lit the start act light to show the player that he can start the next mode
    SetLightColor l34, blue, 2
    vpmtimer.AddTimer 14000, "PlaySong ""mu_main"" '"
End Sub

Sub BelialTargets_Timer()
    Select Case TargetsPos
        Case 0:l28.State = 1:l29.State = 1:l30.State = 1:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 0
        Case 1:l28.State = 0:l29.State = 1:l30.State = 1:l31.State = 1:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 0
        Case 2:l28.State = 0:l29.State = 0:l30.State = 1:l31.State = 1:l35.State = 1:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 0
        Case 3:l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 1:l35.State = 1:l36.State = 1:l37.State = 0:l40.State = 0:l41.State = 0
        Case 4:l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 1:l36.State = 1:l37.State = 1:l40.State = 0:l41.State = 0
        Case 5:l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 1:l37.State = 1:l40.State = 1:l41.State = 0
        Case 6:l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 1:l40.State = 1:l41.State = 1
        Case 7:l28.State = 1:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 1:l41.State = 1
        Case 8:l28.State = 1:l29.State = 1:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 1
        Case 9:l28.State = 1:l29.State = 1:l30.State = 1:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 0
        Case 10:l28.State = 0:l29.State = 1:l30.State = 1:l31.State = 1:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 0
        Case 11:l28.State = 0:l29.State = 0:l30.State = 1:l31.State = 1:l35.State = 1:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 0
        Case 12:l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 1:l35.State = 1:l36.State = 1:l37.State = 0:l40.State = 0:l41.State = 0
        Case 13:l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 1:l36.State = 1:l37.State = 1:l40.State = 0:l41.State = 0 ' the 3 lights "SKI" stay lit for a few seconds
        Case 25:l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 1:l37.State = 1:l40.State = 1:l41.State = 0
        Case 26:l28.State = 0:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 1:l40.State = 1:l41.State = 1
        Case 27:l28.State = 1:l29.State = 0:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 1:l41.State = 1
        Case 28:l28.State = 1:l29.State = 1:l30.State = 0:l31.State = 0:l35.State = 0:l36.State = 0:l37.State = 0:l40.State = 0:l41.State = 1:TargetsPos = -1
    End Select

    TargetsPos = TargetsPos + 1
End Sub

'****************
' Mode 5: Act 3
'****************
' All the acts use these 2 variables:
' MonsterHits, this the numbers of monsters to kill to end the Act and start the battle against the Lord
' MonsterTotal, the total number of monsters of the act, or the total life of the Lord during battles

Sub StartAct3
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act3.mp4", " ", "         ", 6500
    ' start music after the intro
    vpmtimer.AddTimer 6000, "PlaySong ""mu_act"" '"
    'setup variables
    MonsterTotal = 20 + INT(RND * 5) ' from 20 to 25 monsters
    MonsterHits = 0
    Mode(0) = 5                      ' this is the active mode
    Mode(5) = 2                      ' this means it is started, and the corresponding light will be turned on and start blinking
    'update lights
    UpdateModeLights()
    ChangeGi white
    ' Start a timer to light a random arrow(s) light and to change it from time to time
    StartMonsterTimer
End Sub

Sub StopAct3
    StopMonsterTimer
    Mode(0) = 0
    Mode(5) = 0
    UpdateModeLights()
    ChangeGi white
End Sub

Sub WinAct3
    StopMonsterTimer
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act-complete.mp4", "", "", 12000
    FlashEffect 2
  DOF 154, DOFPulse
    'setup variables
    Mode(0) = 0
    Mode(5) = 1
    'update lights
    UpdateModeLights()
    ChangeGi white
    SetLightColor l10, purple, 1
    ' lit the boss light
    SetLightColor l44, green, 2
    vpmtimer.AddTimer 12000, "PlaySong ""mu_main"" '"
End Sub

'************************
' Mode 6: Battle Azmodan
'************************
' The fights with the Lords use just one variable LordStrength
' you need to hit at least 10 times the orbits left or right

Sub StartAzmodan
    StopSound Song:Song = ""
    DMD "azmodan-start.mp4", "", "", 20000
    ' start music after the intro
    vpmtimer.AddTimer 20000, "PlaySong ""mu_lordfight"" '"
    vpmtimer.AddTimer 8000, "PlaySound ""battle"" '"
    LordStrength = 100
    Mode(0) = 6
    Mode(6) = 2
    'update lights
    UpdateModeLights()
    ChangeGi green
    ' turn on the arrow lights
    SetLightColor l32, white, 2
    SetLightColor l39, white, 2
    ' add 3 multiball after the video is finished
    vpmtimer.addtimer 20000, "AddMultiBall 3 '"
    ' Start the life restore timer
    LordLifeStep = 2
    LordLifeInterval = 2000
    LordLifeTimer.Interval = LordLifeInterval
    RestartLordLifeTimer
    StartJackpots
End Sub

Sub StopAzmodan
    Mode(0) = 0
    Mode(6) = 0
    UpdateModeLights()
    l32.state = 0
    l39.State = 0
    ChangeGi white
End Sub

Sub WinAzmodan
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "azmodan-defeated.mp4", "", "", 14000
    FlashEffect 2
  DOF 154, DOFPulse
    LightEffect 2
    GiEffect 2
  DOF 125, DOFPulse
    'setup variables
    Mode(0) = 0
    Mode(6) = 1
    'update lights
    UpdateModeLights()
    l32.state = 0
    l39.State = 0
    ChangeGi white
    SetLightColor l11, purple, 1
    ' lit the start act light to show the player that he can start the next mode
    SetLightColor l34, blue, 2
    vpmtimer.AddTimer 14000, "PlaySong ""mu_main"" '"
End Sub

'****************
' Mode 7: Act 4
'****************
' All the acts use these 2 variables:
' MonsterHits, this the numbers of monsters to kill to end the Act and start the battle against the Lord
' MonsterTotal, the total number of monsters of the act, or the total life of the Lord during battles

Sub StartAct4
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act4.mp4", " ", "         ", 6500
    ' start music after the intro
    vpmtimer.AddTimer 6000, "PlaySong ""mu_act2"" '"
    'setup variables
    MonsterTotal = 25 + INT(RND * 5) ' from 25 to 30 monsters
    MonsterHits = 0
    Mode(0) = 7                      ' this is the active mode
    Mode(7) = 2                      ' this means it is started, and the corresponding light will be turned on and start blinking
    'update lights
    UpdateModeLights()
    ChangeGi white
    ' Start a timer to light a random arrow light and to change it from time to time
    StartMonsterTimer
End Sub

Sub StopAct4
    StopMonsterTimer
    Mode(0) = 0
    Mode(7) = 0
    UpdateModeLights()
End Sub

Sub WinAct4
    StopMonsterTimer
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act-complete.mp4", "", "", 12000
    FlashEffect 2
  DOF 154, DOFPulse
    'setup variables
    Mode(0) = 0
    Mode(7) = 1
    'update lights
    UpdateModeLights()
    SetLightColor l12, purple, 1
    ' lit the boss light
    SetLightColor l44, green, 2
    vpmtimer.AddTimer 12000, "PlaySong ""mu_main"" '"
End Sub

'************************
' Mode 8: Battle Diablo
'************************
' The fights with the Lords use just one variable LordStrength

Sub StartDiablo
    StopSound Song:Song = ""
    DMD "diablo-start.mp4", "", "", 19000
    ' start music after the intro
    vpmtimer.AddTimer 19000, "PlaySong ""mu_lordfight"" '"
    vpmtimer.AddTimer 8000, "PlaySound ""battle"" '"
    LordStrength = 100
    Mode(0) = 8
    Mode(8) = 2
    'update lights
    UpdateModeLights()
    ChangeGi blue
    ' you need to hit the top lanes and ramps, only one light will be active
    SetLightColor l32, white, 2
    ' Start the life restore timer
    LordLifeStep = 3
    LordLifeInterval = 1500
    LordLifeTimer.Interval = LordLifeInterval
    RestartLordLifeTimer
    ' add 4 multiball after the video is finished
    vpmtimer.addtimer 20000, "AddMultiBall 4 '"
    StartJackpots
End Sub

Sub StopDiablo
    Mode(0) = 0
    Mode(8) = 0
    UpdateModeLights()
    l32.State = 0
    l39.State = 0
    ChangeGi white
End Sub

Sub WinDiablo
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "diablo-defeated.mp4", "", "", 14000
    FlashEffect 2
  DOF 154, DOFPulse
    LightEffect 2
    GiEffect 2
  DOF 125, DOFPulse
    'setup variables
    Mode(0) = 0
    Mode(8) = 1
    'update lights
    UpdateModeLights()
    l32.State = 0
    l39.State = 0
    ChangeGi white
    SetLightColor l14, purple, 1
    ' lit the start act light to show the player that he can start the next mode
    SetLightColor l34, blue, 2
    vpmtimer.AddTimer 14000, "PlaySong ""mu_main"" '"
End Sub

'****************
' Mode 9: Act 5
'****************
' All the acts use these 2 variables:
' MonsterHits, this the numbers of monsters to kill to end the Act and start the battle against the Lord
' MonsterTotal, the total number of monsters of the act, or the total life of the Lord during battles

Sub StartAct5
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act5.mp4", " ", "         ", 6500
    ' start music after the intro
    vpmtimer.AddTimer 6000, "PlaySong ""mu_act"" '"
    'setup variables
    MonsterTotal = 30 + INT(RND * 5) ' from 10 to 15 monsters
    MonsterHits = 0
    Mode(0) = 9                      ' this is the active mode
    Mode(9) = 2                      ' this means it is started, and the corresponding light will be turned on and start blinking
    'update lights
    UpdateModeLights()
    ChangeGi white
    ' Start a timer to light a random arrow light and to change it from time to time
    StartMonsterTimer
End Sub

Sub StopAct5
    StopMonsterTimer
    Mode(0) = 0
    Mode(9) = 0
    UpdateModeLights()
End Sub

Sub WinAct5
    StopMonsterTimer
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "act-complete.mp4", "", "", 12000
    FlashEffect 2
  DOF 154, DOFPulse
    ' Score an extra 5 million points
    DMDBlink "black.jpg", "EXTRA SCORE", "5.000.000", 100, 10
    Addscore 5000000
    'setup variables
    Mode(0) = 0
    Mode(9) = 1
    'update lights
    UpdateModeLights()
    SetLightColor l13, purple, 1
    ' lit the boss light
    SetLightColor l44, green, 2
    vpmtimer.AddTimer 12000, "PlaySong ""mu_main"" '"
End Sub

'************************
' Mode 10: Battle Malthael
'************************
' The fights with the Lords use just one variable LordStrength

Sub StartMalthael
    StopSound Song:Song = ""
    DMD "malthael-start.mp4", "", "", 18000
    ' start music after the intro
    vpmtimer.AddTimer 18000, "PlaySong ""mu_lordfight"" '"
    vpmtimer.AddTimer 8000, "PlaySound ""battle"" '"
    LordStrength = 100
    Mode(0) = 10
    Mode(10) = 2
    'update lights
    UpdateModeLights()
    ChangeGi purple
    ' you need to hit the ramps
    SetLightColor l33, white, 2
    SetLightColor l38, white, 2
    ' add 4 multiball after the video is finished
    vpmtimer.addtimer 18000, "AddMultiBall 4 '"
    ' Start the life restore timer
    LordLifeStep = 4
    LordLifeInterval = 1000
    LordLifeTimer.Interval = LordLifeInterval
    RestartLordLifeTimer
    StartJackpots
End Sub

Sub StopMalthael
    Mode(0) = 0
    Mode(10) = 0
    UpdateModeLights()
    l33.State = 0
    l38.State = 0
    ChangeGi white
End Sub

Sub WinMalthael
    'setup variables
    Mode(0) = 11 'the end
    Mode(10) = 1
    'update lights
    l33.State = 0
    l38.State = 0
    UpdateModeLights()
    ChangeGi white
    SetLightColor l47, purple, 1
    DMD "black.jpg", "SHOOT THE CHEST", "TO KILL MALTHAEL", 5000
    MalthaelAttackTimer.Enabled = 0
    ChestDoor.IsDropped = 1
  DOF 116, DOFPulse
    ' lit the 3 chest lights
    SetLightColor l44, purple, 2
    SetLightColor l34, purple, 2
    SetLightColor l25, purple, 2
    FlashEffect 2
  DOF 154, DOFPulse
End Sub

Sub WinMalthael2
    Dim i
    'play a sound - video -dmd
    StopSound Song:Song = ""
    DMD "malthael-defeated.mp4", "", "", 80000
    'light show
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 20, , 80000
    LightSeqInserts.UpdateInterval = 150
    LightSeqInserts.Play SeqRandom, 50, , 80000
    ' Score an extra 5 million points
    DMDBlink "black.jpg", "EXTRA SCORE", "5.000.000", 100, 10
    Addscore 5000000
    'reset all lights and modes
    For i = 0 to 10
        Mode(i) = 0
    Next

    SecondRound = 2 'used to multiply by 2 all the scores in the scond round
    UpdateModeLights()
    l25.State = 0
    l34.State = 0
    l44.State = 0
    ' lit the start act light to show the player that he can start the next mode
    SetLightColor l34, blue, 2
    vpmtimer.AddTimer 80000, "PlaySong ""mu_main"" '"
End Sub

' this is the Malthael attack timer, starts at 2 attacks per minute and increases the frecuency on each act or battle
' In the last battle he will attack each 10 seconds. Each attack is 5 seconds long.

Sub MalthaelAttackTimer_Timer
    Dim tmp
    tmp = Mode(0)
    MalthaelAttackTimer.Interval = 30000 - tmp * 2000
    ' play a sound
    PlaySound "di_thunder"
    ' flash the flashers
    FlashForMs f2, 5000, 50, 0
  DOF 151, DOFPulse
    FlashForMs f14, 5000, 50, 0
  DOF 152, DOFPulse
    FlashForMs f15, 5000, 50, 0
  DOF 153, DOFPulse
    ' enable the magnets for a few seconds
    EnableMagnets
    vpmtimer.AddTimer 5000, "DisableMagnets '"
End Sub

Sub DisableMagnets
    'debug.print "magnets off"
  DOF 127, DOFOff
    LMagnet.MotorOn = False
    RMagnet.MotorOn = False
End Sub

Sub EnableMagnets
  DOF 127, DOFOn
    LMagnet.MotorOn = True
    RMagnet.MotorOn = True
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 20 ' total number of balls
Const lob = 0   'number of locked balls
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
    If UBound(BOT) = -1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball

    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

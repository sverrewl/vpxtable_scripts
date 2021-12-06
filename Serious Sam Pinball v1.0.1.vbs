' ****************************************************************
'                       VISUAL PINBALL X
'                 JPSalas Serious Sam Pinball Script
'   plain VPX script using core.vbs for supporting functions
'                         Version 1.0.0
' ****************************************************************

Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

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
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Dim Controller: Set Controller = CreateObject("B2S.Server"): Controller.Run

Const BallSize = 50 ' 50 is the normal size

' Load the core.vbs for supporting Subs and functions
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
End Sub

' Define any Constants
Const TableName = "SeriousSam"
Const myVersion = "1.0.0"
Const MaxPlayers = 4     ' from 1 to 4
Const BallSaverTime = 20 ' in seconds
Const MaxMultiplier = 5  ' limit to 5x in this game, both bonus multiplier and playfield multiplier
Const BallsPerGame = 5   ' usually 3 or 5
Const MaxMultiballs = 5  ' max number of balls during multiballs

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim PlayfieldMultiplier(4)
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot(4)
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

' Define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock(4)
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
Dim bJackpot

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger during multiballs
Dim cbRight   'captive ball

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    Dim i
    Randomize

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 36 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd "fx_kicker", "fx_solenoid"
        .CreateEvents "plungerIM"
    End With

    Set cbRight = New cvpmCaptiveBall
    With cbRight
        .InitCaptive CapTrigger1, CapWall1, Array(CapKicker1, CapKicker1a), 0
        .NailedBalls = 1
        .ForceTrans = .9
        .MinForce = 3.5
        '.CreateEvents "cbRight"
        .Start
    End With
    CapKicker1.CreateBall

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
    BallsInLock(1) = 0
    BallsInLock(2) = 0
    BallsInLock(3) = 0
    BallsInLock(4) = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bJackpot = False
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
'******************
' Captive Ball Subs
'******************
Sub CapTrigger1_Hit:cbRight.TrigHit ActiveBall:End Sub
Sub CapTrigger1_UnHit:cbRight.TrigHit 0:End Sub
Sub CapWall1_Hit:cbRight.BallHit ActiveBall:PlaySound "fx_collide", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub CapKicker1a_Hit:cbRight.BallReturn Me:End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
    If Keycode = AddCreditKey Then
        Credits = Credits + 1
        If(Tilted = False) Then
            DMDFlush
            DMD "", "CREDITS " &credits, 500
            PlaySoundAtVol "fx_coin", drain, 1
            If NOT bGameInPlay Then ShowTableInfo
        End If
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAtVol "fx_plungerpull", Plunger, 1
        PlaySound "fx_reload", 0, 1, 0.05, 0.05
    End If

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    ' Table specific

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

        If keycode = LeftFlipperKey Then SolLFlipper 1:InstantInfoTimer.Enabled = True
        If keycode = RightFlipperKey Then SolRFlipper 1:InstantInfoTimer.Enabled = True


    ' Thalamus, added instant mechanical Tilt

    If keycode = MechanicalTilt Then Tilt = 16:Tilted = 1:CheckTilt

        If keycode = StartGameKey Then
            If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    DMDFlush
                    DMD " ", PlayersPlayingGame & " PLAYERS", 500
                    PlaySound "so_fanfare1"
                Else
                    If(Credits> 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
                        DMDFlush
                        DMD " ", PlayersPlayingGame & " PLAYERS", 500
                        PlaySound "so_fanfare1"
                    Else
                        ' Not Enough Credits to start a game.
                        DMDFlush
                        DMD "CREDITS " &credits, "INSERT COIN", 500
                        PlaySound "so_nocredits"
                    End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey Then
                DMDFlush
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
                        DMD "CREDITS " &credits, "INSERT COIN", 500
                        ShowTableInfo
                    End If
                End If
            End If
    End If ' If (GameInPlay)

    'test keys
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If keycode = PlungerKey Then
        Plunger.Fire
        PlaySoundAtVol "fx_plunger", Plunger, 1
        If bBallInPlungerLane Then PlaySound "fx_fire", 0, 1, 0.05, 0.05
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
If NOT hsbModeActive Then
    bInstantInfo = True
    DMDFlush
    UltraDMDTimer.Enabled = 1
End If
End Sub

Sub InstantInfo
    DMD "", "INSTANT INFO", 1500
    DMD "JACKPOT VALUE", Jackpot(CurrentPlayer), 1500
    DMD "SPINNER VALUE", spinnervalue(CurrentPlayer), 1500
    DMD "BUMPER VALUE", bumpervalue(CurrentPlayer), 1500
    DMD "BONUS X", BonusMultiplier(CurrentPlayer), 1500
    DMD "PLAYFIELD X", PlayfieldMultiplier(CurrentPlayer), 1500
    DMD "LOCKED BALLS", BallsInLock(CurrentPlayer), 1500
    DMD "LANE BONUS", LaneBonus, 1500
    DMD "TARGET BONUS", TargetBonus, 1500
    DMD "RAMP BONUS", RampBonus, 1500
    DMD "MONSTERS KILLED", MonstersKilled(CurrentPlayer), 1500
    DMD "HIGHEST SCORE", HighScoreName(0) & " " & HighScore(0), 1500
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
  Controller.Stop
  If Not UltraDMD is Nothing Then
    If UltraDMD.IsRendering Then
      UltraDMD.CancelRendering
    End If
    UltraDMD = NULL
  End If
End Sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol "fx_flipperup", LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol "fx_flipperdown", LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol "fx_flipperup", RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol "fx_flipperdown", RightFlipper, VolFlip
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

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                    'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity) AND(Tilt <15) Then 'show a warning
        DMDFlush
        DMD " ", "CAREFUL!", 2000
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        DMDFlush
        DMD " ", "TILT!", 99999
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
    End If
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
            PlaySound Song, -1, 0.7 'this last number is the volume, from 0 to 1
        End If
    End If
End Sub

Sub PlayBattleSong
    Dim tmp
    tmp = INT(RND * 6)
    Select Case tmp
        Case 0:PlaySong "mu_battle1"
        Case 1:PlaySong "mu_battle2"
        Case 2:PlaySong "mu_battle3"
        Case 3:PlaySong "mu_battle4"
        Case 4:PlaySong "mu_battle5"
        Case 5:PlaySong "mu_battle6"
    End Select
End Sub

Sub PlayMultiballSong
    Dim tmp
    tmp = INT(RND * 4)
    Select Case tmp
        Case 0:PlaySong "mu_war1"
        Case 1:PlaySong "mu_war2"
        Case 2:PlaySong "mu_war3"
        Case 3:PlaySong "mu_war4"
    End Select
End Sub

Sub ChangeSong
    If(BallsOnPlayfield = 0) Then
        PlaySong "mu_end"
    Else
        Select Case Battle(CurrentPlayer, 0)
            Case 0
                PlaySong "mu_main"
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
                PlayBattleSong
            Case 13, 14, 15
                PlayMultiballSong
        End Select
    End If
End Sub

'********************
' Play random quotes
'********************

Sub PlayQuote
    Dim tmp
    tmp = INT(RND * 115) + 1
    PlaySound "quote_" &tmp
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
        If UBound(tmp) = 1 Then 'we have 2 captive balls on the table (-1 means no balls, 0 is the first ball, 1 is the second..)
            GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

' GI, light & flashers sequence effects

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
        Case 1 'all blink fast
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 5, 10
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
        Case 3 'all blink fast
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 5, 10
    End Select
End Sub

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
        Case 1 'all blink fast
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqBlinking, , 5, 10
    End Select
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Ramp Soundss
Sub RHelp1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub RHelp2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

'***********************
' Idols follow the ball
'***********************
' they do not move during multiball

Sub aIdolPos_Hit(idx)
    If NOT bMultiBallMode Then
        Idol1.Rotz = -(idx * 2) -32
        Idol2.RotZ = idx * 2 + 32
    End If
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
        PlayfieldMultiplier(i) = 1
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

    ' reduce the playfield multiplier
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
    PlaySoundAtVol "fx_Ballrel", BallRelease, 1
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield> 1 Then
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
        DMD "LANE BONUS", AwardPoints, 1000
        'PlaySound "fx_bonus", 0, 1, 0, 0, 0, 0, 0

        'Number of Target hits
        AwardPoints = TargetBonus * 2000
        TotalBonus = TotalBonus + AwardPoints
        DMD "TARGET BONUS", AwardPoints, 1000
        'vpmtimer.addtimer 1000, "PlaySound ""fx_bonus"", 0, 1, 0, 0, 1000, 0, 0 '"

        'Number of Ramps completed
        AwardPoints = RampBonus * 10000
        TotalBonus = TotalBonus + AwardPoints
        DMD "RAMP BONUS", AwardPoints, 1000
        'vpmtimer.addtimer 2000, "PlaySound ""fx_bonus"", 0, 1, 0, 0, 2000, 0, 0 '"

        'Number of Monsters Killed
        AwardPoints = MonstersKilled(CurrentPlayer) * 25000
        TotalBonus = TotalBonus + AwardPoints
        DMD "Monsters killed", AwardPoints, 1000
        'vpmtimer.addtimer 3000, "PlaySound ""fx_bonus"", 0, 1, 0, 0, 3000, 0, 0 '"

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
        DMD "TOTAL BONUS X " &BonusMultiplier(CurrentPlayer), TotalBonus, 2000
        'vpmtimer.addtimer 4000, "PlaySound ""fx_bonus"", 0, 1, 0, 0, 5000, 0, 0 '"
        AddScore TotalBonus

        ' add a bit of a delay to allow for the bonus points to be shown & added up
        vpmtimer.addtimer 6000, "EndOfBall2 '"
    Else 'if tilted then only add a short delay
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the CurrentPlayer)
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
        DMD " ", "EXTRA BALL", 2000

        ' In this table an extra ball will have the skillshot and ball saver, so we reset the playfield for the new ball
        ResetForNewPlayerBall()

        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0) Then
            'debug.print "No More Balls, High Score Entry"

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
            DMD " ", "PLAYER " &CurrentPlayer, 800
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
        ChangeSong
    End If

    bJustStarted = False
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball

    ' show game over on the DMD
    DMD " ", "GAME OVER", 11000

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
            DMD " ", "BALL SAVE", 2000
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1) Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True) then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ' you may wish to change any music over at this point and
                    ' turn off any multiball specific lights
                    ResetJackpotLights
                    Select Case Battle(CurrentPlayer, 0)
                        Case 13, 14, 15:WinBattle
                    End Select

                    ChangeGi white
                    ChangeSong
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0) Then
                ' End Mode and timers
                ChangeSong
                ChangeGi white
                ' Show the end of ball animation
                ' and continue with the end of ball
                DMDFlush
                ' DMD something?
                StopEndOfBallMode
                vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, since there is no animation then move to the end of ball
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
    'LaunchLight.State = 2

    'be sure to update the Scoreboard after the animations, if any
    UltraDMDScoreTimer.Enabled = 1

    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        PlungerIM.AutoFire
        PlaySoundAtVol "fx_fire", ActiveBall, 1
        bAutoPlunger = False
    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    Else
        ' show the message to shoot the ball in case the player has fallen sleep
        swPlungerRest.TimerEnabled = 1
    End If
    'Start the Selection of the skillshot if ready
    If bSkillShotReady Then
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
    ChangeSong
' turn off LaunchLight
' LaunchLight.State = 0
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds

Sub swPlungerRest_Timer
    DMD "", "SHOOT THE BALL", 2000
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
    LightShootAgain.State = 2
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
        Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer)
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
Sub AddJackpot(points)
    ' Jackpots only generally increment in multiball mode AND not tilted
    ' but this doesn't have to be the case
    If(Tilted = False) Then

        ' If(bMultiBallMode = True) Then
        Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + points
        DMD "INCREASED JACKPOT", Jackpot(CurrentPlayer), 1000
    ' you may wish to limit the jackpot to a upper limit, ie..
    ' If (Jackpot >= 6000) Then
    '   Jackpot = 6000
    '   End if
    'End if
    End if
End Sub

Sub AddSuperJackpot(points) 'not used in this table
    If(Tilted = False) Then
    End if
End Sub

Sub AddBonusMultiplier(n)
    Dim NewBonusLevel
    ' if not at the maximum bonus level
    if(BonusMultiplier(CurrentPlayer) + n <= MaxMultiplier) then
        ' then add and set the lights
        NewBonusLevel = BonusMultiplier(CurrentPlayer) + n
        SetBonusMultiplier(NewBonusLevel)
        DMD " ", "BONUS X " &NewBonusLevel, 2000
        PlaySound "fx_bonux"
    Else
        AddScore 50000
        DMD " ", "50000", 1000
    End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly

Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level
    UPdateBonusXLights(Level)
End Sub

Sub UpdateBonusXLights(Level)
    ' Update the lights
    Select Case Level
        Case 1:light54.State = 0:light55.State = 0:light56.State = 0:light57.State = 0
        Case 2:light54.State = 1:light55.State = 0:light56.State = 0:light57.State = 0
        Case 3:light54.State = 0:light55.State = 1:light56.State = 0:light57.State = 0
        Case 4:light54.State = 0:light55.State = 0:light56.State = 1:light57.State = 0
        Case 5:light54.State = 0:light55.State = 0:light56.State = 0:light57.State = 1
    End Select
End Sub

Sub AddPlayfieldMultiplier(n)
    Dim NewPFLevel
    ' if not at the maximum level x
    if(PlayfieldMultiplier(CurrentPlayer) + n <= MaxMultiplier) then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) + n
        SetPlayfieldMultiplier(NewPFLevel)
        DMD " ", "PLAYFIELD X " &NewPFLevel, 2000
        PlaySound "fx_bonux"
    Else 'if the 5x is already lit
        AddScore 50000
        DMD " ", "50000", 2000
    End if
    'Start the timer to reduce the playfield x every 30 seconds
    pfxtimer.Enabled = 0
    pfxtimer.Enabled = 1
End Sub

' Set the Playfield Multiplier to the specified level AND set any lights accordingly

Sub SetPlayfieldMultiplier(Level)
    ' Set the multiplier to the specified level
    PlayfieldMultiplier(CurrentPlayer) = Level
    UpdatePFXLights(Level)
End Sub

Sub UpdatePFXLights(Level)
    ' Update the lights
    Select Case Level
        Case 1:light3.State = 0:light2.State = 0:light1.State = 0:light4.State = 0
        Case 2:light3.State = 1:light2.State = 0:light1.State = 0:light4.State = 0
        Case 3:light3.State = 0:light2.State = 1:light1.State = 0:light4.State = 0
        Case 4:light3.State = 0:light2.State = 0:light1.State = 1:light4.State = 0
        Case 5:light3.State = 0:light2.State = 0:light1.State = 0:light4.State = 1
    End Select
' show the multiplier in the DMD
' DMD
End Sub

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        DMD " ", "EXTRA BALL", 2000
        PlaySound "vo_extraball"
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        GiEffect 2
        LightEffect 2
    END If
End Sub

Sub AwardSpecial()
    DMD " ", "EXTRA GAME", 2000
    PlaySound "fx_fanfare2"
    Credits = Credits + 1
    LightEffect 2
    FlashEffect 2
End Sub

Sub AwardJackpot() 'award a normal jackpot, double or triple jackpot
Dim tmp
    DMD "JACKPOT", Jackpot(CurrentPlayer), 2000
tmp = INT(RND *2)
Select Case tmp
Case 0:PlaySound "vo_Jackpot"
Case 0:PlaySound "vo_Jackpot2"
Case 0:PlaySound "vo_Jackpot3"
End Select
    AddScore Jackpot(CurrentPlayer)
    LightEffect 2
    FlashEffect 2
    'sjekk for superjackpot
    EnableSuperJackpot
End Sub

Sub AwardSuperJackpot() 'this is actually 4 times a jackpot
    SuperJackpot = Jackpot(CurrentPlayer) * 4
    DMD "SUPER JACKPOT", SuperJackpot, 2000
    PlaySound "vo_superjackpot"
    AddScore SuperJackpot
    LightEffect 2
    FlashEffect 2
    'enabled jackpots again
    StartJackpots
End Sub

Sub AwardSkillshot()
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMDFlush
    DMD "SKILLSHOT", SkillShotValue(CurrentPlayer), 2000
    PlaySound "fx_fanfare2"
  Addscore SkillShotValue(CurrentPlayer)
    ' increment the skillshot value with 250.000
    SkillShotValue(CurrentPlayer) = SkillShotValue(CurrentPlayer) + 250000
    'do some light show
    GiEffect 2
    LightEffect 2
End Sub

'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 100000 End If

    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If

    x = LoadValue(TableName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 100000 End If

    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If

    x = LoadValue(TableName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 100000 End If

    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If

    x = LoadValue(TableName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 100000 End If

    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If

    x = LoadValue(TableName, "Credits")
    If(x <> "") then Credits = CInt(x) Else Credits = 0 End If

    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If
End Sub

Sub Savehs
    SaveValue TableName, "HighScore1", HighScore(0)
    SaveValue TableName, "HighScore1Name", HighScoreName(0)
    SaveValue TableName, "HighScore2", HighScore(1)
    SaveValue TableName, "HighScore2Name", HighScoreName(1)
    SaveValue TableName, "HighScore3", HighScore(2)
    SaveValue TableName, "HighScore3Name", HighScoreName(2)
    SaveValue TableName, "HighScore4", HighScore(3)
    SaveValue TableName, "HighScore4Name", HighScoreName(3)
    SaveValue TableName, "Credits", Credits
    SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
End Sub

Sub Reseths
    HighScoreName(0) = "AAA"
    HighScoreName(1) = "BBB"
    HighScoreName(2) = "CCC"
    HighScoreName(3) = "DDD"
    HighScore(0) = 100000
    HighScore(1) = 100000
    HighScore(2) = 100000
    HighScore(3) = 100000
    Savehs
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
        If NOT bFreePlay Then
            AwardSpecial
        End If
    End If

    If tmp> HighScore(3) Then
        vpmtimer.addtimer 2000, "PlaySound ""fx_fanfare8"" '"
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
    DMDId "hsc", "YOUR NAME:", " <A  > ", 999999
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
            n.color = RGB(0, 16, 0)
            n.colorfull = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 18, 18)
            n.colorfull = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 8, 8)
            n.colorfull = RGB(0, 64, 64)
        Case purple
            n.color = RGB(64, 0, 96)
            n.colorfull = RGB(128, 0, 192)
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
'        JPS DMD - very very simple DMD text only routines using UltraDMD
'***********************************************************************************

Dim UltraDMD

' Available effects:
' FadeIn 0
' FadeOut 1
' ZoomIn 2
' ZoomOut 3
' ScrollOffLeft 4
' ScrollOffRight 5
' ScrollOnLeft 6
' ScrollOnRight 7
' ScrollOffUp 8
' ScrollOffDown 9
' ScrollOnUp 10
' ScrollOnDown 11
' None 14

' text DMD using UltraDMD calls

Sub DMD(toptext, bottomtext, duration)
    UltraDMD.DisplayScene00 "", toptext, 15, bottomtext, 15, 14, duration, 14
    UltraDMDScoreTimer.Enabled = 1 'to show the score after the animation/message
End Sub

Sub DMDScore
    If NOT UltraDMD.IsRendering Then
        Select Case Battle(CurrentPlayer, 0)
            Case 0:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer, "Ball " & Balls
            Case 1:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Spinners left", 200-SpinCount
            Case 2:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Bumper hits left", 100-SuperBumperHIts
            Case 3:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Ramp hits left", 10-ramphits
            Case 4:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Orbit Hits left", 10-orbithits
            Case 5:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Shoot the lights", "Ball " & Balls
            Case 6:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Shoot the lights", "Ball " & Balls
            Case 7:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Hit the targets", 50-TargetHits
            Case 8:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Hit the targets", 20-TargetHits
            Case 9:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Shoot the captive ball", 10-CaptiveBallHits
            Case 10:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Shoot ramps and orbits", 10-RampHits
            Case 11:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Hit the blue targets", 20-TargetHits
            Case 12:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Hit the loops", 5-loopCount
            Case 13:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Hit the lit light", "Ball " & Balls
            Case 14:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Hit the lit light", "Ball " & Balls
            Case 15:UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Battle Mordekai", "Ball " & Balls
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

Sub DMDId(id, toptext, bottomtext, duration) 'used in the highscore entry routine
    UltraDMD.DisplayScene00ExwithID id, False, "", toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
End Sub

Sub DMDMod(id, toptext, bottomtext, duration) 'used in the highscore entry routine
    UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
End Sub

Sub UltraDMDTimer_Timer() 'used for repeating the attrack mode and the instant info.
    If bInstantInfo Then
        InstantInfo
    ElseIf bAttractMode Then
        ShowTableInfo
    End If
End Sub

Sub UltraDMDScoreTimer_Timer() 'call the score after finished rendering
    If NOT UltraDMD.IsRendering Then
        DMDScoreNow
        UltraDMDScoreTimer.Enabled = 0
    End If
End Sub

Sub DMD_Init
    Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    If UltraDMD is Nothing Then
        MsgBox "No UltraDMD found.  This table will NOT run without it."
        Exit Sub
    End If

    UltraDMD.Init
    If Not UltraDMD.GetMajorVersion = 1 Then
        MsgBox "Incompatible Version of UltraDMD found."
        Exit Sub
    End If

    ' wait for the animation to end
    While UltraDMD.IsRendering = True
    WEnd

    ' Show ROM version number
    DMD "SERIOUS SAM", "ROM VERS " &myVersion, 3000
End Sub

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo
    Dim i
    'info goes in a loop only stopped by the credits and the startkey
    If Score(1) Then
        DMD "PLAYER 1", Score(1), 3000
    End If
    If Score(2) Then
        DMD "PLAYER 2", Score(2), 3000
    End If
    If Score(3) Then
        DMD "PLAYER 3", Score(3), 3000
    End If
    If Score(4) Then
        DMD "PLAYER 4", Score(4), 3000
    End If

    'coins or freeplay
    If bFreePlay Then
        DMD " ", "FREE PLAY", 2000
    Else
        If Credits> 0 Then
            DMD "CREDITS " &credits, "PRESS START", 2000
        Else
            DMD "CREDITS " &credits, "INSERT COIN", 2000
        End If
    End If
    ' some info about the table
    DMD "JPSALAS PRESENTS", "SERIOUS SAM", 3000
    ' Highscores
    DMD "HIGHSCORE 1", HighScoreName(0) & " " & HighScore(0), 3000
    DMD "HIGHSCORE 2", HighScoreName(1) & " " & HighScore(1), 3000
    DMD "HIGHSCORE 3", HighScoreName(2) & " " & HighScore(2), 3000
    DMD "HIGHSCORE 4", HighScoreName(3) & " " & HighScore(3), 3000
End Sub

Sub StartAttractMode()
    bAttractMode = True
    UltraDMDTimer.Enabled = 1
    StartLightSeq
    ShowTableInfo
'StartRainbow aBattleLights
End Sub

Sub StopAttractMode()
    bAttractMode = False
    DMDScoreNow
    LightSeqAttract.StopPlay
    LightSeqFlasher.StopPlay
'StopRainbow
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
Dim LaneBonus
Dim TargetBonus
Dim RampBonus
Dim BumperValue(4)
Dim BumperHits
Dim SuperBumperHits
Dim SpinnerValue(4)
Dim MonstersKilled(4)
Dim SpinCount
Dim RampHits
Dim OrbitHits
Dim TargetHits
Dim CaptiveBallHits
Dim loopCount
Dim BattlesWon(4)
Dim Battle(4, 15) '12 battles, 2 wizard modes, 1 final battle
Dim NewBattle
Dim MachineGunHits

Sub Game_Init()   'called at the start of a new game
    Dim i, j
    bExtraBallWonThisBall = False
    TurnOffPlayfieldLights()
    'Play some Music
    ChangeSong
    'Init Variables
    LaneBonus = 0 'it gets deleted when a new ball is launched
    TargetBonus = 0
    RampBonus = 0
    BumperHits = 0
    For i = 1 to 4
        SkillshotValue(i) = 500000
        Jackpot(i) = 100000
        MonstersKilled(i) = 0
        BallsInLock(i) = 0
        SpinnerValue(i) = 1000
        BumperValue(i) = 210 'start at 210 and every 30 hits its value is increased by 500 points
    Next

    ResetBattles
    SpinCount = 0
    SuperBumperHits = 0
    RampHits = 0
    OrbitHits = 0
    TargetHits = 0
    CaptiveBallHits = 0
    loopCount = 0
  MachineGunHits = 0
'Init Delays/Timers
'MainMode Init()
'Init lights
End Sub

Sub StopEndOfBallMode() 'this sub is called after the last ball is drained
    ResetSkillShotTimer_Timer
    StopBattle

End Sub

Sub ResetNewBallVariables() 'reset variables for a new ball or player
    Dim i
    LaneBonus = 0
    TargetBonus = 0
    RampBonus = 0
    BumperHits = 0
    ' select a battle
    SelectBattle
End Sub

Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub UpdateSkillShot() 'Setup and updates the skillshot lights
    LightSeqSkillshot.Play SeqAllOff
    Light21.State = 2
    Light29.State = 2
    Gate2.Open = 1
    Gate3.Open = 1
    DMD "Hit the lit light", "for the skillshot", 3000
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    LightSeqSkillshot.StopPlay
    If Light21.State = 2 Then Light21.State = 0
    Light29.State = 0
    Gate2.Open = 0
    Gate3.Open = 0
    DMDScoreNow
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

' Tree animation
Dim MyPi, TreeStep, TreeDir
MyPi = Round(4 * Atn(1), 6) / 90
TreeStep = 0

Sub Trees_Timer()
    TreeDir = SIN(TreeStep * MyPi)
    TreeStep = (TreeStep + 1) MOD 360
    Tree1.RotY = - TreeDir
    Tree2.RotY = TreeDir
    Tree3.RotY = - TreeDir
    Tree4.RotY = TreeDir
    Tree5.RotY = - TreeDir
    Tree6.RotY = TreeDir
    Tree7.RotY = - TreeDir
    Tree8.RotY = TreeDir
    Tree9.RotY = - TreeDir
    Tree10.RotY = TreeDir
    Tree11.RotY = - TreeDir
    Tree12.RotY = TreeDir
    Tree13.RotY = - TreeDir
    Tree14.RotY = TreeDir
    Tree15.RotY = - TreeDir
End Sub

'*********************************************************
' Slingshots has been hit
' In this table the slingshots change the outlanes lights

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAtVol "fx_slingshot", Lemk, 1
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 210
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
    ChangeOutlanes
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAtVol "fx_slingshot", Remk, 1
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 210
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
    ChangeOutlanes
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select

    RStep = RStep + 1
End Sub

Sub ChangeOutlanes
    Dim tmp
    tmp = light5.State
    light5.State = light9.State
    light9.State = tmp
End Sub

'*********
' Bumpers
'*********
' after each 30 hits the bumpers increase theis score value by 500 points up to 3210
' and they increase the playfield multiplier. The playfield multiplier is reduced after every 30 seconds

Sub Bumper1_Hit
    If NOT Tilted Then
        PlaySoundAtVol "fx_Bumper", Bumper1, VolBump
        ' add some points
        AddScore BumperValue(CurrentPlayer)
        If Battle(CurrentPlayer, 0) = 2 Then
            SuperBumperHits = SuperBumperHits + 1
            Addscore 5000
            CheckWinBattle
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper1"
    End If
    CheckBumpers
End Sub

Sub Bumper2_Hit
    If NOT Tilted Then
        PlaySoundAtVol "fx_Bumper", Bumper2, VolBump
        ' add some points
        AddScore BumperValue(CurrentPlayer)
        If Battle(CurrentPlayer, 0) = 2 Then
            SuperBumperHits = SuperBumperHits + 1
            Addscore 5000
            CheckWinBattle
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper2"
    End If
    CheckBumpers
End Sub

Sub Bumper3_Hit
    If NOT Tilted Then
        PlaySoundAtVol "fx_Bumper", Bumper3, VolBump
        ' add some points
        AddScore BumperValue(CurrentPlayer)
        If Battle(CurrentPlayer, 0) = 2 Then
            SuperBumperHits = SuperBumperHits + 1
            Addscore 5000
            CheckWinBattle
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper3"
    End If
    CheckBumpers
End Sub

' Check the bumper hits

Sub CheckBumpers()
    ' increase the bumper hit count and increase the bumper value after each 30 hits
    BumperHits = BumperHits + 1
    If BumperHits MOD 30 = 0 Then
        If BumperValue(CurrentPlayer) <3210 Then
            BumperValue(CurrentPlayer) = BumperValue(CurrentPlayer) + 500
        End If
        ' lit the playfield multiplier light
        light53.State = 1
    End If
End Sub

'*************************
' Top & Inlanes: Bonus X
'*************************
' lit the 2 top lane lights and the 2 inlane lights to increase the bonus multiplier

Sub sw8_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    Light20.State = 1
    FlashForMs f5, 1000, 50, 0
    If bSkillShotReady Then
        ResetSkillShotTimer_Timer
    Else
        CheckBonusX
    End If
End Sub

Sub sw9_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    Light21.State = 1
    FlashForMs f5, 1000, 50, 0
    If bSkillShotReady Then
        Awardskillshot
    Else
        CheckBonusX
    End If
End Sub

Sub sw2_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    Light6.State = 1
    FlashForMs f8, 1000, 50, 0
    AddScore 5000
    CheckBonusX
'do something
' Do some sound or light effect
End Sub

Sub sw3_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    Light8.State = 1
    FlashForMs f9, 1000, 50, 0
    AddScore 5000
    CheckBonusX
'do something
' Do some sound or light effect
End Sub

Sub CheckBonusX
    If Light20.State + Light21.State + Light6.State + Light8.State = 4 Then
        AddBonusMultiplier 1
        GiEffect 1
        FlashForMs Light20, 1000, 50, 0
        FlashForMs Light21, 1000, 50, 0
        FlashForMs Light6, 1000, 50, 0
        FlashForMs Light8, 1000, 50, 0
    End IF
End Sub

'************************************
' Flipper OutLanes: Virtual kickback
'************************************
' if the light is lit then activate the ballsave

Sub sw1_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    AddScore 50000
    ' Do some sound or light effect
    ' do some check
    If light5.State = 1 Then
        EnableBallSaver 5
    End If
End Sub

Sub sw4_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    AddScore 50000
    ' Do some sound or light effect
    ' do some check
    If Light9.State = 1 Then
        EnableBallSaver 5
    End If
End Sub

'*******************************
' 3Bank Targets: Jungle Targets
'*******************************

Sub Target1_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    ' Do some sound or light effect
    Light12.State = 1
    FlashForMs f7, 1000, 50, 0
    ' do some check
    Check3BankTargets
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Light44.State = 2 Then
                Light44.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light44.State = 2 Then
                Light50.State = 2
                Light44.State = 0
                Addscore 100000
            End If
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 8:TargetHits = TargetHits + 1:Addscore 25000:CheckWinBattle
        Case 13
            If Light44.State = 2 Then
                AddScore 100000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case 14
            If Light44.State = 2 Then
                AddScore 120000
                FlashEffect 3
                DMD " ", "120000", 1000
            End If
    End Select
    LastSwitchHit = "Target1"
End Sub

Sub Target2_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    ' Do some sound or light effect
    Light13.State = 1
    FlashForMs f7, 1000, 50, 0
    ' do some check
    Check3BankTargets
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Light44.State = 2 Then
                Light44.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light44.State = 2 Then
                Light44.State = 0
                Light50.State = 2
                Addscore 100000
            End If
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 8:TargetHits = TargetHits + 1:Addscore 25000:CheckWinBattle
        Case 13
            If Light44.State = 2 Then
                AddScore 100000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case 14
            If Light44.State = 2 Then
                AddScore 120000
                FlashEffect 3
                DMD " ", "120000", 1000
            End If
    End Select
    LastSwitchHit = "Target2"
End Sub

Sub Target3_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    ' Do some sound or light effect
    Light14.State = 1
    FlashForMs f7, 1000, 50, 0
    ' do some check
    Check3BankTargets
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Light44.State = 2 Then
                Light44.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light44.State = 2 Then
                Light50.State = 2
                Light44.State = 0
                Addscore 100000
            End If
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 8:TargetHits = TargetHits + 1:Addscore 25000:CheckWinBattle
        Case 13
            If Light44.State = 2 Then
                AddScore 100000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case 14
            If Light44.State = 2 Then
                AddScore 120000
                FlashEffect 3
                DMD " ", "120000", 1000
            End If
    End Select
    LastSwitchHit = "Target3"
End Sub

Sub Check3BankTargets
    If light12.state + light13.state + light14.state = 3 Then
        light12.state = 0
        light13.state = 0
        light14.state = 0
        LightEffect 1
        FlashEffect 2
        Addscore 30000
        If Light5.State = 0 AND Light9.State = 0 Then 'light one of the outlanes lights
            Light5.State = 1
        Else
            Addscore 20000
        End If
        ' increment the spinner value
        spinnervalue(CurrentPlayer) = spinnervalue(CurrentPlayer) + 500
    End If
End Sub

'************
'  Spinners
'************

Sub spinner1_Spin
    If Tilted Then Exit Sub
    Addscore spinnervalue(CurrentPlayer)
    PlaySoundAtVol "fx_spinner", Spinner1, VolSpin
    Select Case Battle(CurrentPlayer, 0)
        Case 1
            Addscore 3000
            SpinCount = SpinCount + 1
            CheckWinBattle
    End Select
End Sub

Sub spinner2_Spin
    If Tilted Then Exit Sub
    PlaySoundAtVol "fx_spinner", Spinner2, VolSpin
    Addscore spinnervalue(CurrentPlayer)
    Select Case Battle(CurrentPlayer, 0)
        Case 1
            Addscore 3000
            SpinCount = SpinCount + 1
            CheckWinBattle
    End Select
End Sub

'*********************************
' 2Bank targets: The Lock Targets
'*********************************

Sub Target10_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    ' Do some sound or light effect
    Light11.State = 1
    FlashForMs f6, 1000, 50, 0
    ' do some check
    Check2BankTargets
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Light50.State = 2 Then
                Light50.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light50.State = 2 Then
                Light46.State = 2
                Light50.State = 0
                Addscore 100000
            End If
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 8:TargetHits = TargetHits + 1:Addscore 25000:CheckWinBattle
        Case 13
            If Light50.State = 2 Then
                AddScore 100000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case 14
            If Light50.State = 2 Then
                AddScore 120000
                FlashEffect 3
                DMD " ", "120000", 1000
            End If
    End Select
    LastSwitchHit = "Target10"
End Sub

Sub Target11_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    ' Do some sound or light effect
    Light10.State = 1
    FlashForMs f6, 1000, 50, 0
    ' do some check
    Check2BankTargets
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Light50.State = 2 Then
                Light50.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light50.State = 2 Then
                Light46.State = 2
                Light50.State = 0
                Addscore 100000
            End If
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 8:TargetHits = TargetHits + 1:Addscore 25000:CheckWinBattle
        Case 13
            If Light50.State = 2 Then
                AddScore 100000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case 14
            If Light50.State = 2 Then
                AddScore 120000
                FlashEffect 3
                DMD " ", "120000", 1000
            End If
    End Select
    LastSwitchHit = "Target11"
End Sub

Sub Check2BankTargets
    If light11.state + light10.state = 2 Then
        light11.state = 0
        light10.state = 0
        LightEffect 1
        FlashEffect 1
        Addscore 20000
        If(Light27.State = 0) AND(bMultiballMode = FALSE) Then 'lit the lock light if it is off, rise the lock post and activate the lock switch
            Light27.State = 1
            'PlaySound "vo_lockislit"
            DMD " ", "LOCK IS LIT", 2000
        ElseIf light52.State = 0 Then 'lit the increase jackpot light if the lock light is lit
            light52.State = 1
        'PlaySound "vo_IncreaseJakpot"
        Else
            Addscore 30000
        End If
    End If
End Sub

'**************************
' The Lock: Main Multiball
'**************************
' the lock is a virtual lock, where the locked balls are simply counted

Sub lock_Hit
    Dim delay
    delay = 500
    StopSound "fx_metalrolling"
    PlaySoundAtVol "fx_kicker_enter", lock, VolKick
    If light27.State = 1 Then 'lock the ball
        BallsInLock(CurrentPlayer) = BallsInLock(CurrentPlayer) + 1
        delay = 4000
        Select Case BallsInLock(CurrentPlayer)
            Case 1:PlaySound "vo_ball1locked":DMD "BALL 1 LOCKED", "", 2000
            Case 2:PlaySound "vo_ball2locked":DMD "BALL 2 LOCKED", "", 2000
            Case 3:PlaySound "vo_ball3locked":DMD "BALL 3 LOCKED", "", 2000
        End Select

        light27.State = 0
        If BallsInLock(CurrentPlayer) = 3 Then 'start multiball
            vpmtimer.addtimer 2000, "StartMainMultiball '"
        End If
    End If
    If(Battle(CurrentPlayer, NewBattle) = 2) AND(Battle(CurrentPlayer, 0) = 0) Then 'the battle is ready, so start it
        vpmtimer.addtimer 2000, "StartBattle '"
        delay = 6000
    End If
    vpmtimer.addtimer delay, "ReleaseLockedBall '"
End Sub

Sub ReleaseLockedBall 'release locked ball
    FlashForMs f6, 1000, 50, 0
    lockpost.isdropped = 1:PlaySoundAtVol "fx_solenoid", lock, 1
    lock.kick 190, 4
    lockpost.TimerInterval = 400
    lockpost.TimerEnabled = 1
End Sub

Sub lockpost_Timer
    lockpost.TimerEnabled = 0
    lockpost.isdropped = 0:PlaySound "fx_solenoidoff", 0, 1, 0.05 'TODO
End Sub

Sub StartMainMultiball
    AddMultiball 3
    PlaySound "vo_multiball"
    DMD " ", "MULTIBALL", 2000
    StartJackpots
    ChangeGi 5
  'reset BallsInLock variable
  BallsInLock(CurrentPlayer) = 0
End Sub

'**********
' Jackpots
'**********
' Jackpots are enabled during the Main multiball and the wizard battles

Sub StartJackpots
    bJackpot = true
    'turn on the jackpot lights
    Select Case Battle(CurrentPlayer, 0)
        Case 13 'first wizard mode
            light28.State = 2
            light30.State = 2
        Case 14 'second wizard mode
            light24.State = 2
            light28.State = 2
            light30.State = 2
        Case 15 'final battle
            light24.State = 2
            light25.State = 2
            light28.State = 2
            light30.State = 2
            light32.State = 2
    Case Else
      If bMultiballMode Then
                light28.State = 2
                light30.State = 2
      End If
    End Select
End Sub

Sub ResetJackpotLights 'when multiball is finished, resets jackpot and superjackpot lights
    bJackpot = False
    light24.State = 0
    light25.State = 0
    light28.State = 0
    light30.State = 0
    light32.State = 0
    light29.State = 0
    light51.State = 0
End Sub

Sub EnableSuperJackpot
    If bJackpot = True Then
        If light24.State + light25.State + light28.State + light30.State + light32.State = 0 Then
            'PlaySound "vo_superjackpotislit"
            light29.State = 2
            light51.State = 2
        End If
    End If
End Sub

'***********************************
' 5Bank Targets:  The Idols Targets
'***********************************

Sub Target4_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target4"
    ' Do some sound or light effect
    Light15.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 11:TargetHits = TargetHits + 1:Addscore 50000:CheckWinBattle
    End Select

    Check5BankTargets
End Sub

Sub Target5_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target5"
    ' Do some sound or light effect
    Light16.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 11:TargetHits = TargetHits + 1:Addscore 50000:CheckWinBattle
    End Select

    Check5BankTargets
End Sub

Sub Target7_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target7"
    ' Do some sound or light effect
    Light17.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 11:TargetHits = TargetHits + 1:Addscore 50000:CheckWinBattle
    End Select

    Check5BankTargets
End Sub

Sub Target8_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target8"
    ' Do some sound or light effect
    Light18.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 11:TargetHits = TargetHits + 1:Addscore 50000:CheckWinBattle
    End Select

    Check5BankTargets
End Sub

Sub Target9_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target9"
    ' Do some sound or light effect
    Light19.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits = TargetHits + 1:Addscore 10000:CheckWinBattle
        Case 11:TargetHits = TargetHits + 1:Addscore 50000:CheckWinBattle
    End Select

    Check5BankTargets
End Sub

Sub Check5BankTargets
    Dim tmp
    FlashForMs f7, 1000, 50, 0
    FlashForMs f1, 1000, 50, 0
    FlashForMs f2, 1000, 50, 0
    FlashForMs f5, 1000, 50, 0
    FlashForMs f6, 1000, 50, 0
    tmp = INT(RND * 27) + 1
    PlaySound "enemy_" &tmp, 0, 1, pan(ActiveBall), 0.05
    ' if all 5 targets are hit then kill a monster & activate the mystery light
    If light15.state + light16.state + light17.state + light18.state + light19.state = 5 Then
        ' kill a monster & play enemy sound
        MonstersKilled(CurrentPlayer) = MonstersKilled(CurrentPlayer) + 1
        DMD "MONSTERS KILLED", MonstersKilled(CurrentPlayer), 2000
        LightEffect 1
        FlashEffect 1
        ' Lit the Mystery light if it is off
        If Light61.State = 1 Then
            AddScore 50000
        Else
            Light61.State = 1
            AddScore 25000
        End If
        ' reset the lights
        light15.state = 0
        light16.state = 0
        light17.state = 0
        light18.state = 0
        light19.state = 0
    End If
End Sub

' Playfiel Multiplier timer: reduces the multiplier after 30 seconds

Sub pfxtimer_Timer
    If PlayfieldMultiplier(CurrentPlayer)> 1 Then
        PlayfieldMultiplier(CurrentPlayer) = PlayfieldMultiplier(CurrentPlayer) -1
        SetPlayfieldMultiplier PlayfieldMultiplier(CurrentPlayer)
    Else
        pfxtimer.Enabled = 0
    End If
End Sub

'*****************
'  Captive Target
'*****************

Sub Target6_Hit
    PlaySoundAtVol "fx_target", ActiveBall, VolTarg
    If Tilted Then Exit Sub
    If bSkillShotReady Then
        Awardskillshot
        Exit Sub
    End If
    AddScore 5000 'all targets score 5000
    ' Do some sound or light effect
    ' do some check
    If(bJackpot = True) AND(light51.State = 2) Then
        AwardSuperJackpot
        light51.State = 0
        light29.State = 0
        StartJackpots
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 0:SelectBattle 'no battle is active then change to another battle
        Case 9:CaptiveBallHits = CaptiveBallHits + 1:Addscore 25000:CheckWinBattle
    End Select

    ' increase the playfield multiplier for 30 seconds
    If light53.State = 1 Then
        AddPlayfieldMultiplier 1
        light53.State = 0
    End If

    ' increase Jackpot
    If light52.State = 1 Then
        AddJackpot 50000
        light52.State = 0
    End If
End Sub

'****************************
'  Pyramid Hole Hit & Awards
'****************************

Sub PyramidKicker_Hit
    Dim Delay
    Delay = 200
    PlaySoundAtVol "fx_kicker_enter", PyramidKicker, VolKick
    If NOT Tilted Then
        ' do something
        If(bJackpot = True) AND(light24.State = 2) Then
            light24.State = 0
            AwardJackpot
            Delay = 2000
        End If
        If light61.State = 1 Then ' mystery light is lit
            light61.State = 0
            GiveRandomAward
            Delay = 5000
        End If
        If light23.State = 1 Then ' extra ball is lit
            light23.State = 0
            AwardExtraBall
            Delay = 2000
        End If
        Select Case Battle(CurrentPlayer, 0)
            Case 5
                If Light46.State = 2 Then
                    Light46.State = 0
                    Addscore 100000
                    CheckWinBattle
                    Delay = 1000
                End If
            Case 6
                If Light46.State = 2 Then
                    Light45.State = 2
                    Light46.State = 0
                    Addscore 100000
                    Delay = 1000
                End If
            Case 13
                If Light46.State = 2 Then
                    AddScore 100000
                    FlashEffect 3
                    DMD " ", "100000", 1000
                End If
            Case 14
                If Light46.State = 2 Then
                    AddScore 120000
                    FlashEffect 3
                    DMD " ", "120000", 1000
                End If
        End Select
    End If
    vpmtimer.addtimer Delay, "PyramidExit '"
End Sub

Sub PyramidExit()
    FlashForMs f3, 1000, 50, 0
    PlaySoundAtVol "fx_kicker", PyramidKicker, VolKick
    PyramidKicker.kick 180, 35
    PlaySoundAtVol "fx_cannon", PyramidKicker, VolKick
End Sub

Sub GiveRandomAward() 'from the Pyramid
    Dim tmp, tmp2

    ' show some random values on the dmd
    DMD " ", "EXTRA POINTS", 20
    DMD " ", "PLAYFIELD X", 20
    DMD " ", "BUMPER VALUE", 20
    DMD " ", "EXTRA POINTS", 20
    DMD " ", "EXTRA BALL", 20
    DMD " ", "EXTRA POINTS", 20
    DMD " ", "BONUS X", 20
    DMD " ", "EXTRA POINTS", 20
    DMD " ", "SPINNER VALUE", 20
    DMD " ", "BUMPER VALUE", 20
    DMD " ", "EXTRA POINTS", 20
    DMD " ", "PLAYFIELD X", 20
    DMD " ", "BUMPER VALUE", 20
    DMD " ", "EXTRA POINTS", 20
    DMD " ", "EXTRA BALL", 20
    DMD " ", " ", 200

    vpmtimer.addtimer 2200, "PlaySound ""fx_fanfare1"" '"

    tmp = INT(RND(1) * 80)
    Select Case tmp
        Case 1, 2, 3, 4, 5, 6 'Lit Extra Ball
            DMD "EXTRA BALL IS LIT", " ", 2000
            light23.State = 2
        Case 7, 8, 13, 14, 15 '100,000 points
            DMD "BIG POINTS", "100000", 2000
            AddScore 100000
        Case 9, 10, 11, 12 'Hold Bonus
            DMD "BONUS HELD", "ACTIVATED", 2000
            bBonusHeld = True
        Case 16, 17, 18 'Increase Bonus Multiplier
            DMD "INCREASED", "BONUS X", 2000
            AddBonusMultiplier 1
        Case 19, 20, 21 'Complete Battle
            If Battle(CurrentPlayer, 0)> 0 AND Battle(CurrentPlayer, 0) <13 Then
                DMD "BATTLE", "COMPLETED", 2000
                WinBattle
            Else
                DMD "BIG POINTS", "100000", 2000
                AddScore 100000
            End If
        Case 22, 23, 36, 37, 38 'PlayField multiplier
            DMD "INCREASED", "PLAYFIELD X", 2000
            AddPlayfieldMultiplier 1
        Case 24, 25, 26, 27, 28 '100,000 points
            DMD "BIG POINTS", "100000", 2000
            AddScore 500000
        Case 29, 30, 31, 32, 33, 34, 35 'Increase Bumper value
            BumperValue(CurrentPlayer) = BumperValue(CurrentPlayer) + 500
            DMD "BUMPER VALUE", BumperValue(CurrentPlayer), 2000
        Case 39, 40, 43, 44 'extra multiball
            DMD "EXTRA", "MULTIBALL", 2000
            AddMultiball 1
        Case 45, 46, 47, 48 ' Ball Save
            DMD "BALL SAVE", "ACTIVATED", 2000
            EnableBallSaver 20
        Case ELSE 'Add a Random score from 10.000 to 100,000 points
            tmp2 = INT((RND) * 9) * 10000 + 10000
            DMD "EXTRA POINTS", tmp2, 2000
            AddScore tmp2
    End Select
End Sub

'*******************
'   The Orbit lanes
'*******************

Sub sw5_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    If(bJackpot = True) AND(light25.State = 2) Then
        light25.State = 0
        AwardJackpot
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 4:OrbitHits = OrbitHits + 1:Addscore 70000:CheckWinBattle
        Case 5
            If Light45.State = 2 Then
                Light45.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light45.State = 2 Then
                Light49.State = 2
                Light45.State = 0
                Addscore 100000
            End If
        Case 10
            If Light45.State = 2 Then
                RampHits = RampHits + 1
                Light48.State = 2
                Light47.State = 2
                Light45.State = 0
                Light49.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 12
            If LastSwitchHit = "sw6" Then
                LastSwitchHit = ""
                loopCount = loopCount + 1
                Addscore 140000
                CheckWinBattle
            End If
        Case 13
            If Light45.State = 2 Then
                AddScore 100000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case 14
            If Light45.State = 2 Then
                AddScore 120000
                FlashEffect 3
                DMD " ", "120000", 1000
            End If
    End Select
    LastSwitchHit = "sw5"
End Sub

Sub sw6_Hit
    PlaySoundAtVol "fx_sensor", ActiveBall, 1
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    If(bJackpot = True) AND(light32.State = 2) Then
        light32.State = 0
        AwardJackpot
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 4:OrbitHits = OrbitHits + 1:Addscore 70000:CheckWinBattle
        Case 5
            If Light49.State = 2 Then
                Light49.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light49.State = 2 Then
                Light47.State = 2
                Light49.State = 0
                Addscore 100000
            End If
        Case 10
            If Light49.State = 2 Then
                RampHits = RampHits + 1
                Light48.State = 2
                Light47.State = 2
                Light45.State = 0
                Light49.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 12
            If LastSwitchHit = "sw5" Then
                LastSwitchHit = ""
                loopCount = loopCount + 1
                Addscore 140000
                CheckWinBattle
            End If
        Case 13
            If Light49.State = 2 Then
                AddScore 100000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case 14
            If Light49.State = 2 Then
                AddScore 120000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
    End Select
    LastSwitchHit = "sw6"
End Sub

'****************
'     Ramps
'****************

Sub LeftRampDone_Hit
    Dim tmp
    PlaySoundAtVol "fx_metalrolling", ActiveBall, 1
    If Tilted Then Exit Sub
    'increase the ramp bonus
    RampBonus = RampBonus + 1
    If(bJackpot = True) AND(light28.State = 2) Then
        light28.State = 0
        AwardJackpot
    End If
  'Machine Gun - left ramp only counts the variable
  MachineGunHits = MachineGunHits + 1
  CheckMachineGun
    'Battles
    Select Case Battle(CurrentPlayer, 0)
        Case 3:RampHits = RampHits + 1:Addscore 100000:CheckWinBattle
        Case 5
            If Light47.State = 2 Then
                Light47.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light47.State = 2 Then
                Light48.State = 2
                Light47.State = 0
                Addscore 100000
            End If
        Case 10
            If Light47.State = 2 Then
                RampHits = RampHits + 1
                Light48.State = 0
                Light47.State = 0
                Light45.State = 2
                Light49.State = 2
                Addscore 100000
                CheckWinBattle
            End If
        Case 13
            If Light47.State = 2 Then
                AddScore 100000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case 14
            If Light47.State = 2 Then
                AddScore 120000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case else
            ' play ss quote
            PlayQuote
    End Select
    'check for combos
    if LastSwitchHit = "LeftRampDone" Then
        Addscore jackpot(CurrentPlayer)
        DMD "COMBO", jackpot(CurrentPlayer), 1000
    End If
    LastSwitchHit = "LeftRampDone"
End Sub

Sub RightRampDone_Hit
    Dim tmp
    PlaySoundAtVol "fx_metalrolling", ActiveBall, 1
    If Tilted Then Exit Sub
    'increase the ramp bonus
    RampBonus = RampBonus + 1
    If(bJackpot = True) AND(light30.State = 2) Then
        light30.State = 0
        AwardJackpot
    End If
  'Machine Gun - rightt ramp counts the variable and give the jackpot if light31 is lit
  If light31.State = 2 Then
    DMD "Machine GUN", Jackpot(CurrentPlayer), 2000
    PlaySound "vo_Jackpot"
    AddScore Jackpot(CurrentPlayer)
    LightEffect 2
    FlashEffect 2
  Else
  MachineGunHits = MachineGunHits + 1
  CheckMachineGun
  End If
    'Battles
    Select Case Battle(CurrentPlayer, 0)
        Case 3:RampHits = RampHits + 1:Addscore 100000:CheckWinBattle
        Case 5
            If Light48.State = 2 Then
                Light48.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light48.State = 2 Then
                Light48.State = 0
                Addscore 100000
                WinBattle
            End If
        Case 10
            If Light48.State = 2 Then
                RampHits = RampHits + 1
                Light48.State = 0
                Light47.State = 0
                Light45.State = 2
                Light49.State = 2
                Addscore 100000
                CheckWinBattle
            End If
        Case 13
            If Light48.State = 2 Then
                AddScore 100000
                FlashEffect 3
                DMD " ", "100000", 1000
            End If
        Case 14
            If Light48.State = 2 Then
                AddScore 120000
                FlashEffect 3
                DMD " ", "120000", 1000
            End If
        Case else
            ' play ss quote
            PlayQuote
    End Select

    'check for combos
    if LastSwitchHit = "RightRampDone" Then
        Addscore jackpot(CurrentPlayer)
        DMD "COMBO", jackpot(CurrentPlayer), 1000
    End If
    LastSwitchHit = "RightRampDone"
End Sub

'************************
'       Battles
'************************

' This table has 12 main battles, 2 wizard battles, and a final battle
' you may choose any the 12 main battles you want to play
' the first wizard mode is played after completing 4 battles
' the second wizard battle is played after completing 8 battles
' After completing all 12 battles you play the final battle

' current active battle number is stored in Battle(CurrentPlayer,0)

Sub SelectBattle 'select a new random battle if none is active
    Dim i
    If Battle(CurrentPlayer, 0) = 0 Then
        ' reset the battles that are not finished
        For i = 1 to 15
            If Battle(CurrentPlayer, i) = 2 Then Battle(CurrentPlayer, i) = 0
        Next
        Select Case BattlesWon(CurrentPlayer)
            Case 4:NewBattle = 13:Battle(CurrentPlayer, NewBattle) = 2:UpdateBattleLights:StartBattle  '4 battles = start wizard mode
            Case 9:NewBattle = 14:Battle(CurrentPlayer, NewBattle) = 2:UpdateBattleLights:StartBattle  '8 battles + wizard mode = start 2nd wizard mode
            Case 14:NewBattle = 15:Battle(CurrentPlayer, NewBattle) = 2:UpdateBattleLights:StartBattle '12 battles + 2 wizard modes = start final battle
            Case Else
                NewBattle = INT(RND * 12 + 1)
                do while Battle(CurrentPlayer, NewBattle) <> 0
                    NewBattle = INT(RND * 12 + 1)
                loop
                Battle(CurrentPlayer, NewBattle) = 2
                Light26.State = 2
                UpdateBattleLights
        End Select
    'debug.print "newbatle " & newbattle
    End If
End Sub

' Update the lights according to the battle's state
Sub UpdateBattleLights
    Light38.State = Battle(CurrentPlayer, 1)
    Light7.State = Battle(CurrentPlayer, 2)
    Light36.State = Battle(CurrentPlayer, 3)
    Light40.State = Battle(CurrentPlayer, 4)
    Light33.State = Battle(CurrentPlayer, 5)
    Light37.State = Battle(CurrentPlayer, 6)
    Light42.State = Battle(CurrentPlayer, 7)
    Light34.State = Battle(CurrentPlayer, 8)
    Light39.State = Battle(CurrentPlayer, 9)
    Light43.State = Battle(CurrentPlayer, 10)
    Light35.State = Battle(CurrentPlayer, 11)
    Light41.State = Battle(CurrentPlayer, 12)
    Light64.State = Battle(CurrentPlayer, 13)
    Light65.State = Battle(CurrentPlayer, 14)
    Light58.State = Battle(CurrentPlayer, 15)
End Sub

' Starting a battle means to setup some lights and variables, maybe timers
' Battle lights will always blink during an active battle
Sub StartBattle
    Battle(CurrentPlayer, 0) = NewBattle
    Light26.State = 0
    ChangeSong
    PlaySound "fx_alarm"
    Select Case NewBattle
        Case 1 'Serpent Yards = Super Spinners
            DMD "Serpent Yards", "started", 2000
            DMD "Shoot the", "spinners", 2000
            Light45.State = 2
            Light49.State = 2
            SpinCount = 0
        Case 2 'Calakmul = Super Pop Bumpers
            DMD "Calakmul", "started", 2000
            DMD "Shoot the", "pop bumpers", 2000
            Light22.State = 2
            LightSeqBumpers.Play SeqRandom, 10, , 1000
            SuperBumperHits = 0
        Case 3 'Sierra de Chiapas = Ramps
            DMD "Sierra de Chiapas", "started", 2000
            DMD "Shoot the", "ramps", 2000
            Light47.State = 2
            Light48.State = 2
            RampHits = 0
        Case 4 'Karnak = Orbits
            DMD "karnak", "started", 2000
            DMD "Shoot the", "orbits", 2000
            OrbitHits = 0
            Light45.State = 2
            Light49.State = 2
        Case 5 'Copán = Shoot the lights 2
            DMD "copan", "started", 2000
            DMD "Shoot the", "lights", 2000
            Light44.State = 2
            Light45.State = 2
            Light46.State = 2
            Light47.State = 2
            Light48.State = 2
            Light49.State = 2
            Light50.State = 2
        Case 6 'The Pit = Shoot the lights 1
            DMD "the pit", "started", 2000
            DMD "Shoot the", "lights", 2000
            Light44.State = 2
        Case 7 'Palenque=  Target Frenzy
            DMD "palenque", "started", 2000
            DMD "Shoot the", "targets", 2000
            TargetHits = 0
            LightSeqAllTargets.Play SeqRandom, 10, , 1000
        Case 8 'Chichén Itzá = Left & Right Targets
            DMD "Chichen Itza", "started", 2000
            DMD "Shoot the", "targets", 2000
            TargetHits = 0
            Light44.State = 2
            Light50.State = 2
        Case 9 'Bonampak = Captive Ball
            DMD "Bonampak", "started", 2000
            DMD "Shoot the", "captive ball", 2000
            CaptiveBallHits = 0
            Light29.State = 2
        Case 10 'Monte Albán = Ramps and Orbits
            'use the ramphits to count the hits
            DMD "monte alban", "started", 2000
            DMD "Shoot the", "ramps and orbits", 2000
            RampHits = 0
            Light47.State = 2
            Light48.State = 2
        Case 11 'Uxmal = Blue Targets
            DMD "uxmal", "started", 2000
            DMD "Shoot the", "blue targets", 2000
            TargetHits = 0
            LightSeqBlueTargets.Play SeqRandom, 10, , 1000
        Case 12 'Valley of the Jaguar = Super Loops
            DMD "valley od the jaguar", "started", 2000
            DMD "Shoot the", "loops", 2000
            loopCount = 0
            Light45.State = 2
            Light49.State = 2
            Gate2.Open = 1
            Gate3.Open = 1
        Case 13 'The Grand Cathedral = Follow the Lights 1
            DMD "the grand cathedral", "started", 2000
            DMD "Shoot the", "lit light", 2000
            FollowTheLights.Enabled = 1
            AddMultiball 2
            StartJackpots
            ChangeGi 5
        Case 14 'The City of the Gods = Follow the Lights 2
            DMD "the city of the gods", "started", 2000
            DMD "Shoot the", "lit light", 2000
            FollowTheLights.Enabled = 1
            AddMultiball 2
            StartJackpots
            ChangeGi 5
        Case 15 'Mordekai the Summoner - the final battle
            DMD "Mordekai", "started", 2000
            DMD "Shoot the", "jackpots", 2000
            AddMultiball 4
            StartJackpots
            ChangeGi 5
    End Select
End Sub

' wizard modes can't be won, you simply play them.
' check if the battle is completed
Sub CheckWinBattle
    dim tmp
    tmp = INT(RND * 8) + 1
    PlaySound "fx_thunder" & tmp
    LightSeqInserts.StopPlay 'stop the light effects before starting again so they don't play too long.
    LightEffect 3
    Select Case NewBattle
        Case 1
            If SpinCount = 200 Then WinBattle:End if
        Case 2
            If SuperBumperHits = 100 Then WinBattle:End if
        Case 3
            If RampHits = 10 Then WinBattle:End if
        Case 4
            If OrbitHits = 8 Then WinBattle:End if
        Case 5
            If Light44.State + Light45.State + Light46.State + Light47.State + Light48.State + Light49.State + Light50.State = 0 Then WinBattle:End if
        Case 6
        Case 7
            If TargetHits = 50 Then WinBattle:End if
        Case 8
            If TargetHits = 20 Then WinBattle:End if
        Case 9
            If CaptiveBallHits = 10 Then WinBattle:End if
        Case 10:
            If RampHits = 10 Then WinBattle
        Case 11
            If TargetHits = 20 Then WinBattle:End if
        Case 12
            If loopCount = 5 Then WinBattle:End if
    End Select
End Sub

Sub StopBattle 'called at the end of a ball
    Dim i
    Battle(CurrentPlayer, 0) = 0
    For i = 0 to 15
        If Battle(CurrentPlayer, i) = 2 Then Battle(CurrentPlayer, i) = 0
    Next
    UpdateBattleLights
    StopBattle2
    NewBattle = 0
End Sub

'called after completing a battle
Sub WinBattle
    Dim tmp
    DMDFlush
    BattlesWon(CurrentPlayer) = BattlesWon(CurrentPlayer) + 1
    Battle(CurrentPlayer, 0) = 0
    Battle(CurrentPlayer, NewBattle) = 1
    UpdateBattleLights
    FlashEffect 2
    LightEffect 2
  GiEffect 2
    DMD " ", "BATTLE COMPLETED", 2000
  PlaySound"fx_Explosion01"
    tmp = INT(RND * 4)
    Select Case tmp
        Case 0:vpmtimer.addtimer 1500, "PlaySound ""vo_excelent"" '"
        Case 1:vpmtimer.addtimer 1500, "PlaySound ""vo_impressive"" '"
        Case 2:vpmtimer.addtimer 1500, "PlaySound ""vo_welldone"" '"
        Case 3:vpmtimer.addtimer 1500, "PlaySound ""vo_YouWon"" '"
    End Select

    StopBattle2
    NewBattle = 0
    SelectBattle 'automatically select a new battle
    ChangeSong
End Sub

Sub StopBattle2
    'Turn off the bomb lights
    Light44.State = 0
    Light45.State = 0
    Light46.State = 0
    Light47.State = 0
    Light48.State = 0
    Light49.State = 0
    Light50.State = 0
    ' stop some timers or reset battle variables
    Select Case NewBattle
        Case 1:SpinCount = 0
        Case 2:Light22.State = 0:LightSeqBumpers.StopPlay:SuperBumperHits = 0
        Case 3:RampHits = 0
        Case 4:OrbitHits = 0
        Case 5:LoopCount = 0
        Case 6:
        Case 7:LightSeqAllTargets.StopPlay:TargetHits = 0
        Case 8:Light44.State = 0:Light50.State = 0:TargetHits = 0
        Case 9:CaptiveBallHits = 0:Light29.State = 0
        Case 10:RampHits = 0
        Case 11:TargetHits = 0:LightSeqBlueTargets.StopPlay
        Case 12:LoopCount = 0:Gate2.Open = 0:Gate3.Open = 0
        Case 13:FollowTheLights.Enabled = 0:SelectBattle
        Case 14:FollowTheLights.Enabled = 0:SelectBattle
        Case 15:ResetBattles:SelectBattle
    End Select
End Sub

Sub ResetBattles
    Dim i, j
    For j = 0 to 4
        BattlesWon(j) = 0
        For i = 0 to 15
            Battle(CurrentPlayer, i) = 0
        Next
    Next

    NewBattle = 0
End Sub

'Extra subs for the battles

Sub LightSeqAllTargets_PlayDone()
    LightSeqAllTargets.Play SeqRandom, 10, , 1000
End Sub

Sub LightSeqBumpers_PlayDone()
    LightSeqBumpers.Play SeqRandom, 10, , 1000
End Sub

Sub LightSeqBlueTargets_PlayDone()
    LightSeqBlueTargets.Play SeqRandom, 10, , 1000
End Sub

' Wizards modes timer
Dim FTLstep:FTLstep = 0

Sub FollowTheLights_Timer
    Light44.State = 0
    Light45.State = 0
    Light46.State = 0
    Light47.State = 0
    Light48.State = 0
    Light49.State = 0
    Light50.State = 0
    Select Case Battle(CurrentPlayer, 0)
        Case 13 '1st wizard mode
            Select case FTLstep
                Case 0:FTLstep = 1:Light44.State = 2
                Case 1:FTLstep = 2:Light45.State = 2
                Case 2:FTLstep = 3:Light46.State = 2
                Case 3:FTLstep = 4:Light47.State = 2
                Case 4:FTLstep = 5:Light48.State = 2
                Case 5:FTLstep = 6:Light49.State = 2
                Case 6:FTLstep = 7:Light50.State = 2
                Case 7:FTLstep = 8:Light49.State = 2
                Case 8:FTLstep = 9:Light48.State = 2
                Case 9:FTLstep = 10:Light47.State = 2
                Case 10:FTLstep = 11:Light46.State = 2
                Case 11:FTLstep = 0:Light45.State = 2
            End Select
        Case 14 '2nd wizard mode
            FTLstep = INT(RND * 7)
            Select case FTLstep
                Case 0:Light44.State = 2
                Case 1:Light45.State = 2
                Case 2:Light46.State = 2
                Case 3:Light47.State = 2
                Case 4:Light48.State = 2
                Case 5:Light49.State = 2
                Case 6:Light50.State = 2
            End Select
    End Select
End Sub

'**********************
' Machine Gun Jackpot
'**********************
' 30 seconds hurry up with jackpots on the right ramp
' uses variable MachineGunHits and the light31

Sub CheckMachineGun
If light31.State = 0 Then
  If MachineGunHits MOD 10 = 0 Then
    EnableMachineGun
  End If
End If
End Sub

Sub EnableMachineGun
    ' start the timers
    MachineGunTimerExpired.Enabled = True
    MachineGunSpeedUpTimer.Enabled = True
    ' turn on the light
    Light31.BlinkInterval = 160
    Light31.State = 2
End Sub

Sub MachineGunTimerExpired_Timer()
    MachineGunTimerExpired.Enabled = False
    ' turn off the light
    Light31.State = 0
End Sub

Sub MachineGunSpeedUpTimer_Timer()
    MachineGunSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    Light31.BlinkInterval = 80
    Light31.State = 2
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
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
Const lob = 1   'number of locked balls
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
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

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


'
'        ***** **                                                                          ***** **              ***     ***
'     ******  ***                                                                       ******  ***               ***     ***
'   **    *  * ***                                                                     **   *  * **                **      **
'  *     *  *   ***                                                                   *    *  *  **                **      **
'       *  *     *** ***  ****                                ****                        *  *   *                 **      **
'      ** **      **  **** **** *    ****        ****        * ***  * ***  ****          ** **  *        ****      **      **
'      ** **      **   **   ****    * ***  *    *  ***  *   *   ****   **** **** *       ** ** *        * ***  *   **      **
'      ** **      **   **          *   ****    *    ****   **    **     **   ****        ** ***        *   ****    **      **
'      ** **      **   **         **    **    **     **    **    **     **    **         ** ** ***    **    **     **      **
'      ** **      **   **         **    **    **     **    **    **     **    **         ** **   ***  **    **     **      **
'      *  **      **   **         **    **    **     **    **    **     **    **         *  **     ** **    **     **      **
'         *       *    **         **    **    **     **    **    **     **    **            *      ** **    **     **      **
'    *****       *     ***        **    **    **     **     ******      **    **        ****     ***  **    **     **      **
'   *   *********       ***        ***** **    ********      ****       ***   ***      *  ********     ***** **    *** *   *** *
'  *       ****                     ***   **     *** ***                 ***   ***    *     ****        ***   **    ***     ***
'  *                                                  ***                             *
'   **                                          ****   ***                             **
'                                             *******  **
'                                            *     ****
'
'
'        *******                                                             *******                                                                       ***** **         **
'      *       ***                                                         *       ***               *                                                  ******  **** *       **
'     *         **                                                        *         **              ***                                                **   *  * ****        **
'     **        *                                                         **        *                *                                                *    *  *   **         **
'      ***          **   ****        ****              ***  ****           ***                             **   ****                                      *  *               **
'     ** ***         **    ***  *   * ***  *    ***     **** **** *       ** ***           ****    ***      **    ***  *     ****    ***  ****           ** **           *** **
'      *** ***       **     ****   *   ****    * ***     **   ****         *** ***        * ***  *  ***     **     ****     * ***  *  **** **** *        ** **          *********
'        *** ***     **      **   **    **    *   ***    **                  *** ***     *   ****    **     **      **     *   ****    **   ****         ** ******     **   ****
'          *** ***   **      **   **    **   **    ***   **                    *** ***  **    **     **     **      **    **    **     **    **          ** *****      **    **
'            ** ***  **      **   **    **   ********    **                      ** *** **    **     **     **      **    **    **     **    **          ** **         **    **
'             ** **  **      **   **    **   *******     **                       ** ** **    **     **     **      **    **    **     **    **          *  **         **    **
'              * *   **      **   **    **   **          **                        * *  **    **     **     **      **    **    **     **    **             *          **    **
'    ***        *     ******* **  *******    ****    *   ***             ***        *   **    **     **      *********    **    **     **    **         ****         * **    **
'   *  *********       *****   ** ******      *******     ***           *  *********     ***** **    *** *     **** ***    ***** **    ***   ***       *  ***********   *****     **
'  *     *****                    **           *****                   *     *****        ***   **    ***            ***    ***   **    ***   ***     *     ******       ***     ****
'  *                              **                                   *                                      *****   ***                             *                           **
'   **                            **                                    **                                  ********  **                               **
'                                  **                                                                      *      ****
'
'
' ***************************************************************************************
'                          Dragon Ball: Super Saiyan Edition
'                       v1.0 built for VISUAL PINBALL X 10.8
'              based off of JP's Wrath of Olympus prototype v5.5 table
' ***************************************************************************************


Option Explicit
Randomize

Const BallSize = 50    ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1     ' standard ball mass in JP's VPX Physics 3.0.1
Const SongVolume = 0.3 ' 1 is full volume, but I set it quite low to listen better the other sounds since I use headphones, adjust to your setup :)

'FlexDMD in high or normal quality
'change it to True if you have an LCD screen, 256x64
'or keep it False if you have a real DMD at 128x32 in size
Const FlexDMDHighQuality = False

'**************  USER OPTIONS   *********************************************************
' If TimedModes is set to False then it will turn off the battle timers so
' you can play the mode until you win or loose the ball
' If it is set to True then the battles/trainings will have duration of about 2 minutes

Const TimedModes = False

' If FastBonusCount is True then the bonus count will you only the Total Bonus awarded
' otherwise the individual bonus will be shown

Const FastBonusCount = False

'Change 0 to 1 for alternate attract track
AttractMusicAlt = 0

'Change 0 to 1 for alternate ball (Orange DragonBall)
BallAlt = 0
'set to true for slight orange glow from ALT dragonball
Dim aBallGlowing:aBallGlowing = False 'True

' Pin between flippers
Const FlipperPin = True

'Table has 6 different desktop wallpapers.  TO SET: go to Table > Image Manager, Rename backdrop to backdrop1. Select new backdrop file and rename to 'backdrop'.

'********   END OF  USER OPTIONS   ******************************************************

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
Const cGameName = "DragonBallSSE"
Const myVersion = "1.0.0"
Const MaxPlayers = 4          ' from 1 to 4
Const MaxMultiplier = 10      ' limit playfield multiplier
Const MaxBonusMultiplier = 10 'limit Bonus multiplier
Const BallsPerGame = 3        ' usually 3 or 5
Const MaxMultiballs = 6       ' max number of balls during multiballs

' Use FlexDMD if in FS mode
Dim UseFlexDMD
If Table1.ShowDT = True then
    UseFlexDMD = False
Else
    UseFlexDMD = True
End If


' Define Global Variables
Dim BallSaverTime ' in seconds of the first ball and during the game
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim PlayfieldMultiplier(4)
Dim PFxSeconds
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot(4)
Dim SuperJackpot(4)
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue(4)
Dim SuperSkillshotValue(4)
Dim bAutoPlunger
Dim bInstantInfo
Dim bAttractMode
Dim FatesCounter
Dim AttractMusicAlt
Dim BallAlt
Dim x      'used in loops
Dim LF, RF 'they are 0 when the flippers are down and 1 when they are up, used to skip the bonus count

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
Dim bJackpot

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger during multiballs
Dim cbLeft    'captive ball at the magnet

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    LoadEM
    Dim i
    Randomize

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 0.5        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFXDOF("fx_kicker", 141, DOFPulse, DOFContactors), SoundFXDOF("fx_solenoid", 141, DOFPulse, DOFContactors)
        .CreateEvents "plungerIM"
    End With

    Set cbLeft = New cvpmCaptiveBall
    With cbLeft
        .InitCaptive CapTrigger, CapWall, CapKicker, 0
        .ForceTrans = .7
        .MinForce = 3.5
        .CreateEvents "cbLeft"
        .Start
    End With

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' load saved values, highscore, names, jackpot
    Credits = 0
    Loadhs

    ' Initalise the DMD display
    DMD_Init

    ' freeplay or coins
    bFreePlay = False 'we want coins

    if bFreePlay Then DOF 121, DOFOn

    ' Init main variables and any other flags
  bAttractMode = False
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    PFxSeconds = 0
    bGameInPlay = False
    bAutoPlunger = False
    bMusicOn = True
    BallsOnPlayfield = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJackpot = False
    bInstantInfo = False
    ' set any lights for the attract mode
    GiOff
    StartAttractMode

    ' Start the RealTime timer
    RealTime.Enabled = 1

    ' Load table color
    LoadLut
  'LoadBall
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25


    If keycode = LeftMagnaSave Then bLutActive = True:SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If
    If Keycode = AddCreditKey OR Keycode = AddCreditKey2 Then
        If Credits <30 Then Credits = Credits + 1
        if bFreePlay = False Then DOF 121, DOFOn
        If(Tilted = False) Then
            DMDFlush
            DMD "", CL("CREDITS " & Credits), "", eNone, eNone, eNone, 500, True, "fx_coin"
            If NOT bGameInPlay Then ShowTableInfo
        End If
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then CheckTilt 'only check the tilt during game
        If keycode = RightTiltKey Then CheckTilt
        If keycode = CenterTiltKey Then CheckTilt
        If keycode = MechanicalTilt Then CheckTilt

        If keycode = LeftFlipperKey Then lf = 1:SolLFlipper 1:InstantInfoTimer.Enabled = True:RotateLaneLights 1:LeftdiverterTimer.Enabled = 1
        If keycode = RightFlipperKey Then rf = 1:SolRFlipper 1:InstantInfoTimer.Enabled = True:RotateLaneLights 0:RightdiverterTimer.Enabled = 1
        If keycode = RightMagnaSave Then kickBallOut 'sometimes all the balls don't come out of the scoop (!?)

        If keycode = StartGameKey Then
            If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, True, ""
                Else
                    If(Credits> 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
                        DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, True, ""
                        If Credits <1 And bFreePlay = False Then DOF 121, DOFOff
                        Else
                            ' Not Enough Credits to start a game.
                            DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 1000, True, "vo_MrSatanLaugh_" &RndNbr(3)
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
                            If Credits <1 And bFreePlay = False Then DOF 121, DOFOff
                            ResetForNewGame()
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        DMDFlush
                        DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 1000, True, "vo_MrSatanLaugh_" &RndNbr(3)
                        ShowTableInfo
                    End If
                End If
            End If
    End If ' If (GameInPlay)

  'TEST KEYS FOR CHEESE
  'If keycode = "3" Then
  ' CurrentMode(CurrentPlayer) = CurrentMode(CurrentPlayer) + 0 '1=cell, 2=zamasu, 3=goku black, 4=frieza, 5=trunks, 6=vegeta, 7=piccolo, 8=goku, 9=kai mode
  ' StartNextMode                    'start next mode, press again for next mode after
  'End If

  'If keycode = "4" Then
  ' ModeStep = ModeStep +1           'Complete battle/training step 1, All battles have at least 2 step, sometimes up to 5.
  ' li061.State = 1                  'Upper goku lights
  ' li062.State = 1                  'Upper goku lights
  ' li063.State = 1                  'Upper goku lights
  ' li064.State = 1                  'Upper goku lights
  ' OrbitHits = OrbitHits + 1        'Add 1 to orbit count
  ' CheckWinMode
  'End If

    'if keycode = "7" Then
  ' BallsInLock(CurrentPlayer) = 3   'Dragon radar
  ' CheckMinotaurMB                  'Dragon rader
  'End If

    'if keycode = "8" Then
  ' bMedusaMBStarted = True          'Level 3 mb (goku blue)
  ' StartMedussaMB                   'Level 3 mb (goku blue)
  'end If

    'if keycode = "9" Then
  ' ZeusCount(CurrentPlayer) = 3     'level 4 mb (Goku Ultra Instinct)
  ' CheckStartModes                  'level 4 mb (Goku Ultra Instinct)
  ' StartZeusMB                      'level 4 mb (Goku Ultra Instinct)
  'End If

End Sub

Sub Table1_KeyUp(ByVal keycode)

    If keycode = LeftMagnaSave Then bLutActive = False:HideLUT

    If keycode = PlungerKey Then
        Plunger.Fire
        PlaySoundAt "fx_plunger", plunger
    End If

    If hsbModeActive Then
        Exit Sub
    End If

    ' Table specific

    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then
            lf = 0
            SolLFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo = True Then
                DMDScoreNow
                bInstantInfo = False
            End If
            LeftdiverterTimer.Enabled = 0
            diverter1.Isdropped = 0
            StopSound "vo_kamehamehaR1"
            StopSound "vo_kamehamehaR2"
        End If
        If keycode = RightFlipperKey Then
            rf = 0
            SolRFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo = True Then
                DMDScoreNow
                bInstantInfo = False
            End If
            RightdiverterTimer.Enabled = 0
            diverter2.Isdropped = 0
            StopSound "vo_kamehamehaL1"
            StopSound "vo_kamehamehaL2"
        End If
    End If
End Sub

Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    If NOT hsbModeActive Then
        bInstantInfo = True
        DMDFlush
        InstantInfo
    End If
End Sub

'*************
' Pause Table
'*************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub Table1_Exit
    Savehs
    If UseFlexDMD Then FlexDMD.Run = False
    If B2SOn = true Then Controller.Stop
End Sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper2.RotateToEnd
        LeftFlipper001.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper2.RotateToStart
        LeftFlipper001.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper2.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper001_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper001_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'*********************************************************
' Real Time Flipper adjustments - by JLouLouLou & JPSalas
'        (to enable flipper tricks)
'*********************************************************

Dim FlipperPower
Dim FlipperElasticity
Dim SOSTorque, SOSAngle
Dim FullStrokeEOS_Torque, LiveStrokeEOS_Torque
Dim LeftFlipperOn
Dim RightFlipperOn

Dim LLiveCatchTimer
Dim RLiveCatchTimer
Dim LiveCatchSensivity

FlipperPower = 5000
FlipperElasticity = 0.8
FullStrokeEOS_Torque = 0.3 ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2 ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.1
SOSAngle = 6

LiveCatchSensivity = 10

LLiveCatchTimer = 0
RLiveCatchTimer = 0

LeftFlipper.TimerInterval = 1
LeftFlipper.TimerEnabled = 1

Sub LeftFlipper_Timer 'flipper's tricks timer
    'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If LeftFlipperOn = 1 Then
        If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle then
            LeftFlipper.EOSTorque = FullStrokeEOS_Torque
            LLiveCatchTimer = LLiveCatchTimer + 1
            If LLiveCatchTimer <LiveCatchSensivity Then
                LeftFlipper.Elasticity = 0
            Else
                LeftFlipper.Elasticity = FlipperElasticity
                LLiveCatchTimer = LiveCatchSensivity
            End If
        End If
    Else
        LeftFlipper.Elasticity = FlipperElasticity
        LeftFlipper.EOSTorque = LiveStrokeEOS_Torque
        LLiveCatchTimer = 0
    End If


    'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If RightFlipperOn = 1 Then
        If RightFlipper.CurrentAngle = RightFlipper.EndAngle Then
            RightFlipper.EOSTorque = FullStrokeEOS_Torque
            RLiveCatchTimer = RLiveCatchTimer + 1
            If RLiveCatchTimer <LiveCatchSensivity Then
                RightFlipper.Elasticity = 0
            Else
                RightFlipper.Elasticity = FlipperElasticity
                RLiveCatchTimer = LiveCatchSensivity
            End If
        End If
    Else
        RightFlipper.Elasticity = FlipperElasticity
        RightFlipper.EOSTorque = LiveStrokeEOS_Torque
        RLiveCatchTimer = 0
    End If
End Sub

Sub RotateLaneLights(n) 'n is the direction, 1 = left or 0 = right
    Dim tmp
    If bRotateLights Then
        If n = 1 Then
            tmp = li002.State
            li002.State = li003.State
            li003.State = li004.State
            li004.State = li005.State
            li005.State = tmp
        Else
            tmp = li005.State
            li005.State = li004.State
            li004.State = li003.State
            li003.State = li002.State
            li002.State = tmp
        End If
    End If
End Sub

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt 'Called when table is nudged
    Dim BOT
    BOT = GetBalls

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub
    Tilt = Tilt + TiltSensitivity                  'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity) AND(Tilt <= 15) Then 'show a warning
        DMD "_", CL("CAREFUL"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
    End if
    If(NOT Tilted) AND Tilt> 15 Then 'If more that 15 then TILT the table
        'display Tilt
        InstantInfoTimer.Enabled = False
        DMDFlush
        DMD CL("YOU"), CL("TILTED"), "", eNone, eNone, eNone, 200, False, ""
        StopSound "vo_kamehamehaL1"
        StopSound "vo_kamehamehaL2"
        StopSound "vo_kamehamehaR1"
        StopSound "vo_kamehamehaR2"
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
        bMultiBallMode = False           'normally disabled in the drain sub
        StopMBmodes
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
        Tilted = True
        'turn off GI and turn off all the lights
        GiOff
        LightSeqTilt.Play SeqAllOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        LeftFlipper001.RotateToStart
        RightFlipper.RotateToStart
        Bumper1.Threshold = 100
        Bumper2.Threshold = 100
        Bumper3.Threshold = 100
        Bumper4.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    TurnOffPFLayers               'turn off playfield layers
    Else
        Tilted = False
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        Bumper1.Threshold = 1
        Bumper2.Threshold = 1
        Bumper3.Threshold = 1
        Bumper4.Threshold = 1
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
        vpmtimer.Addtimer 2000, "EndOfBall() '"
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'*****************************************
'         Music as wav sounds
' in VPX 10.7 you may use also mp3 or ogg
'*****************************************

Dim Song
Dim BaseSong
Song = ""

Sub PlaySong(name)
    If bMusicOn Then
    Debug.print name
    If Song <> name Then
            StopSound Song
            Song = name
            PlaySound Song, -1, SongVolume
    End If
    End If
End Sub

Sub ChangeSong
    If bZeusMBStarted = True Then
    If Song = "mu_GOKUmultiball" Then 'Do not restart song
        Else
      PlaySong "mu_GOKUmultiball"
    End If
  Else If bMedusaMBStarted = True Then
    If Song = "mu_SSBmultiball" Then 'Do not restart song
        Else
      PlaySong "mu_SSBmultiball"
    End If
  Else If bMinotaurMBStarted = True Then
    If Song = "mu_SSmultiball" Then 'Do not restart song
    Else
      PlaySong "mu_SSmultiball"
    End If
  Else If bCCorpMBStarted = True Then
    If Song = "mu_CCmultiball" Then 'Do not restart song
    Else
      PlaySong "mu_CCmultiball"
    End If
  Else
    Select Case Mode(CurrentPlayer, 0)
      Case 0
        If PlayNextSong.Enabled = 0 Then
          BaseSong = RndNbr(7)
          If BaseSong = 1 Then
            PlaySong "mu_base1":PlayNextSong.Interval = 136000:PlayNextSong.Enabled = 1
          Else If BaseSong = 2 Then
            PlaySong "mu_base2":PlayNextSong.Interval = 271000:PlayNextSong.Enabled = 1
          Else If BaseSong = 3 Then
            PlaySong "mu_base3":PlayNextSong.Interval = 195000:PlayNextSong.Enabled = 1
          Else If BaseSong = 4 Then
            PlaySong "mu_base4":PlayNextSong.Interval = 202000:PlayNextSong.Enabled = 1
          Else If BaseSong = 5 Then
            PlaySong "mu_base5":PlayNextSong.Interval = 294000:PlayNextSong.Enabled = 1
          Else If BaseSong = 6 Then
            PlaySong "mu_base6":PlayNextSong.Interval = 192000:PlayNextSong.Enabled = 1
          Else
            PlaySong "mu_base7":PlayNextSong.Interval = 224000:PlayNextSong.Enabled = 1
          End If
          End If
          End If
          End If
          End If
          End If
        Else
          If PlayNextSong.Interval = 224000 Then
            PlaySong "mu_base1":PlayNextSong.Interval = 136000:PlayNextSong.Enabled = 1
          Else If PlayNextSong.Interval = 136000 Then
            PlaySong "mu_base2":PlayNextSong.Interval = 271000:PlayNextSong.Enabled = 1
          Else If PlayNextSong.Interval = 271000 Then
            PlaySong "mu_base3":PlayNextSong.Interval = 195000:PlayNextSong.Enabled = 1
          Else If PlayNextSong.Interval = 195000 Then
            PlaySong "mu_base4":PlayNextSong.Interval = 202000:PlayNextSong.Enabled = 1
          Else If PlayNextSong.Interval = 202000 Then
            PlaySong "mu_base5":PlayNextSong.Interval = 294000:PlayNextSong.Enabled = 1
          Else If PlayNextSong.Interval = 294000 Then
            PlaySong "mu_base6":PlayNextSong.Interval = 192000:PlayNextSong.Enabled = 1
          Else
            PlaySong "mu_base7":PlayNextSong.Interval = 224000:PlayNextSong.Enabled = 1
          End If
          End If
          End If
          End If
          End If
          End If
        End If
      Case 1:PlaySong "mu_minotaur"  ' CELL/Minotaur
      Case 2:PlaySong "mu_hydra"     ' ZAMASU/Hydra
      Case 3:PlaySong "mu_cerberus"  ' GOKU BLACK/Cerberus
      Case 4:PlaySong "mu_medusa"    ' GOLDEN FRIEZA/Medusa
      Case 5:PlaySong "mu_ares"      ' TRUNKS/Ares
      Case 6:PlaySong "mu_poseidon"  ' VEGETA/Poseidon
      Case 7:PlaySong "mu_hades"     ' PICCOLO/Hades
      Case 8:PlaySong "mu_zeus":PlayNextSong.Interval = 171000:PlayNextSong.Enabled = 1   'GOKU/Zeus track with intro
      Case 9                         ' Supreme Kai/God or King Kai/Demi-God WIZARD mode
        If bGokuCompleted (CurrentPlayer) = True Then
          PlaySong "mu_KaiMode120"
        Else
          PlaySong "mu_KaiMode60"
        End If
    End Select
  End If
  End If
  End If
  End If
End Sub

Sub PlayNextSong_Timer
  Me.Enabled = 0
  Select Case Mode(CurrentPlayer, 0)
    Case 0
      If song = "mu_base1" Then
        PlaySong "mu_base2":PlayNextSong.Interval = 271000:PlayNextSong.Enabled = 1
        'Debug.print BaseSong &" mu_base2. Next loop"
      Else If song = "mu_base2" Then
        PlaySong "mu_base3":PlayNextSong.Interval = 195000:PlayNextSong.Enabled = 1
        'Debug.print BaseSong &" mu_base3. Next loop"
      Else If song = "mu_base3" Then
        PlaySong "mu_base4":PlayNextSong.Interval = 202000:PlayNextSong.Enabled = 1
        'Debug.print BaseSong &" mu_base4. Next loop"
      Else If song = "mu_base4" Then
        PlaySong "mu_base5":PlayNextSong.Interval = 294000:PlayNextSong.Enabled = 1
        'Debug.print BaseSong &" mu_base5. Next loop"
      Else If song = "mu_base5" Then
        PlaySong "mu_base6":PlayNextSong.Interval = 192000:PlayNextSong.Enabled = 1
        'Debug.print BaseSong &" mu_base6. Next loop"
      Else If song = "mu_base6" Then
        PlaySong "mu_base7":PlayNextSong.Interval = 224000:PlayNextSong.Enabled = 1
        'Debug.print BaseSong &" mu_base7. Next loop"
      Else If song = "mu_base7" Then
        PlaySong "mu_base1":PlayNextSong.Interval = 136000:PlayNextSong.Enabled = 1
        'Debug.print BaseSong &" mu_base1. Next loop"
      End If
      End If
      End If
      End If
      End If
      End If
      End If
        Case 8:PlaySong "mu_zeus2"
  End Select
End Sub

Sub StopSong(name)
    StopSound name
End Sub

'******************************
' Play random quotes & sounds
'******************************

Sub PlayThunder
    PlaySound "sfx_thunder" &RndNbr(9)
End Sub

Sub PlayLightning
    PlaySound "sfx_lightning" &RndNbr(6)
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

Dim GiIntensity
GiIntensity = 1   'used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
    SetFlashColor GiFlash, col, 1
End Sub

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = 0 Then '-1 means no balls, 0 is the first captive ball, 1 is the second captive ball...)
            GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    PlaySoundAt "fx_GiOn", li008 'about the center of the table
    DOF 118, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    PlaySoundAt "fx_GiOff", li008 'about the center of the table
    DOF 118, DOFOff
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
            LightSeqGi.UpdateInterval = 40
            LightSeqGi.Play SeqBlinking, , 15, 25
        Case 2 'random
            LightSeqGi.UpdateInterval = 25
            LightSeqGi.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqGi.UpdateInterval = 20
            LightSeqGi.Play SeqBlinking, , 10, 10
        Case 4 'seq up
            LightSeqGi.UpdateInterval = 3
            LightSeqGi.Play SeqUpOn, 25, 3
        Case 5 'seq down
            LightSeqGi.UpdateInterval = 3
            LightSeqGi.Play SeqDownOn, 25, 3
    End Select
End Sub

Sub LightEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqInserts.Play SeqAlloff
        Case 1 'all blink
            LightSeqInserts.UpdateInterval = 40
            LightSeqInserts.Play SeqBlinking, , 15, 25
        Case 2 'random
            LightSeqInserts.UpdateInterval = 25
            LightSeqInserts.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqInserts.UpdateInterval = 20
            LightSeqInserts.Play SeqBlinking, , 10, 10
        Case 4 'center
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqCircleOutOn, 15, 2
        Case 5 'top down
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqDownOn, 15, 1
        Case 6 'down to top
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 15, 1
    End Select
End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp> 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v4.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 1     'number of locked balls
Const maxvel = 45 'max ball velocity
ReDim rolling(tnob)



InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
        aBallGlow(i).Visible = 0
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
        aBallShadow(b).Y = 2072 'hide it under the apron
        aBallGlow(b).Visible = 0
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y
        aBallShadow(b).Height = BOT(b).Z - Ballsize / 2

        If aBallGlowing Then
            aBallGlow(b).Visible = 1
            aBallGlow(b).X = BOT(b).X
            aBallGlow(b).Y = BOT(b).Y +10
            aBallGlow(b).BulbHaloHeight = BOT(b).Z + 26
        Else
            aBallGlow(b).Visible = 0
        End If

        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 3
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        End If

        ' jps ball speed & spin control
        BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound "fx_collide", 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'************************************
' Diverse Collection Hit Sounds v3.0
'************************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aRubber_Pegs_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub
Sub aCBallsHit_Hit(idx):PlaySoundAt "fx_collide", CapKicker:End Sub 'just the sound of the ball hitting the captive ball

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
    Dim i

    bGameInPLay = True

    Cell_DeathHalo.visible = 0 'revives all dead enemies
    Zamasu_DeathHalo.visible = 0
    GokuBlack_DeathHalo.visible = 0
    Frieza_DeathHalo.visible = 0
  Trunks_PoweredUp.visible = 0 'resets heroes power levels
  Vegeta_PoweredUp.visible = 0
  Piccolo_PoweredUp.visible = 0
  LightSeqGoku.StopPlay
  DragonRadar.image = "Radar_4"

    'resets the score display, and turn off attract mode
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

    ' initialise specific Game variables
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
    DMDScoreNow
  LayerCheck

  ' set the current players bonus multiplier back down to 1X
    SetBonusMultiplier 1

    ' set the playfield multiplier
    SetPlayfieldMultiplier 1

    ' reset any drop targets, lights, game Mode etc..

    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False

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
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    UpdateBallImage

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    PlaySoundAt SoundFXDOF("fx_Ballrel", 107, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield> 1 Then
        DOF 131, DOFPulse
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

Sub LayerCheck
  If bTrunksCompleted(CurrentPlayer) Then
    Trunks_PoweredUp.visible = 1
  Else
    Trunks_PoweredUp.visible = 0
  End If
  If bVegetaCompleted(CurrentPlayer) Then
    Vegeta_PoweredUp.visible = 1
  Else
    Vegeta_PoweredUp.visible = 0
  End If
  If bPiccoloCompleted(CurrentPlayer) Then
    Piccolo_PoweredUp.visible = 1
  Else
    Piccolo_PoweredUp.visible = 0
  End If
  If bCellCompleted(CurrentPlayer) Then
    Cell_DeathHalo.visible = 1
  Else
    Cell_DeathHalo.visible = 0
  End If
  If bZamasuCompleted(CurrentPlayer) Then
    Zamasu_DeathHalo.visible = 1
  Else
    Zamasu_DeathHalo.visible = 0
  End If
  If bGokuBlackCompleted(CurrentPlayer) Then
    GokuBlack_DeathHalo.visible = 1
  Else
    GokuBlack_PoweredUp.visible = 0
  End If
  If bFriezaCompleted(CurrentPlayer) Then
    Frieza_DeathHalo.visible = 1
  Else
    Frieza_DeathHalo.visible = 0
  End If
End Sub

Sub EndOfBall()
    Dim AwardPoints, TotalBonus, ii
    AwardPoints = 0
    TotalBonus = 10 'yes 10 points :)
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False
    TurnOffPFLayers 'turns off any pf layers left on

    ' only process any of this if the table is not tilted.
    '(the tilt recovery mechanism will handle any extra balls or end of game)

    If NOT Tilted Then
        PlaySong "mu_plunger"
        'Count the bonus. This table uses several bonus750
        DMD CL("BONUS"), "", "", eNone, eNone, eNone, 750, True, ""

        'Targets Hit x 3,500
        AwardPoints = BonusTargets(CurrentPlayer) * 3500
        TotalBonus = TotalBonus + AwardPoints
        If NOT FastBonusCount Then DMD CL("TARGETS HIT " & BonusTargets(CurrentPlayer) ), CL(FormatScore(AwardPoints) ), "", eNone, eBlinkFast, eNone, 750, True, ""

        'Ramps Hit x 9,000
        AwardPoints = BonusRamps(CurrentPlayer) * 9000
        TotalBonus = TotalBonus + AwardPoints
        If NOT FastBonusCount Then DMD CL("RAMPS HIT " & BonusRamps(CurrentPlayer) ), CL(FormatScore(AwardPoints) ), "", eNone, eBlinkFast, eNone, 750, True, ""

        'Orbits Hit x 12,500
        AwardPoints = BonusOrbits(CurrentPlayer) * 12500
        TotalBonus = TotalBonus + AwardPoints
        If NOT FastBonusCount Then DMD CL("ORBITS HIT " & BonusOrbits(CurrentPlayer) ), CL(FormatScore(AwardPoints) ), "", eNone, eBlinkFast, eNone, 750, True, ""

        'Trees collected x 12,500
        AwardPoints = TreeHits(CurrentPlayer) * 12500
        TotalBonus = TotalBonus + AwardPoints
        If NOT FastBonusCount Then DMD CL("LIT IN-LANES " & TreeHits(CurrentPlayer) ), CL(FormatScore(AwardPoints) ), "", eNone, eBlinkFast, eNone, 750, True, ""

        'Combo Hits x 25,000
        AwardPoints = ComboHits(CurrentPlayer) * 25000
        TotalBonus = TotalBonus + AwardPoints
        If NOT FastBonusCount Then DMD CL("COMBO HITS " & ComboHits(CurrentPlayer) ), CL(FormatScore(AwardPoints) ), "", eNone, eBlinkFast, eNone, 750, True, ""

        'X Hits x 50,000
        AwardPoints = BonusXHits(CurrentPlayer) * 50000
        TotalBonus = TotalBonus + AwardPoints
        If NOT FastBonusCount Then DMD CL("X HITS " & BonusXHits(CurrentPlayer) ), CL(FormatScore(AwardPoints) ), "", eNone, eBlinkFast, eNone, 750, True, ""

        'Monsters defeated x 300,000
        AwardPoints = TotalMonsters(CurrentPlayer) * 300000
        TotalBonus = TotalBonus + AwardPoints
        If NOT FastBonusCount Then DMD CL("ENEMIES DEFEATED " & TotalMonsters(CurrentPlayer) ), CL(FormatScore(AwardPoints) ), "", eNone, eBlinkFast, eNone, 750, True, ""

        'Gods defeated x 300,000
        AwardPoints = TotalGods(CurrentPlayer) * 300000
        TotalBonus = TotalBonus + AwardPoints
        If NOT FastBonusCount Then DMD CL("TRAININGS COMPLETED " & TotalGods(CurrentPlayer) ), CL(FormatScore(AwardPoints) ), "", eNone, eBlinkFast, eNone, 750, True, ""

        If NOT FastBonusCount Then DMD CL("BONUS X MULTIPLIER"), CL(FormatScore(TotalBonus) & " X " & BonusMultiplier(CurrentPlayer) ), "", eNone, eNone, eNone, 1500, True, ""
        TotalBonus = TotalBonus * BonusMultiplier(CurrentPlayer)
        DMD CL("TOTAL BONUS"), CL(FormatScore(TotalBonus) ), "", eNone, eNone, eNone, 2000, True, ""
        AddScore2 TotalBonus

        ' add a bit of a delay to allow for the bonus points to be shown & added up
        If FastBonusCount Then
            vpmtimer.addtimer 3000, "EndOfBall2 '"
        Else
            vpmtimer.addtimer 11000, "EndOfBall2 '"
        End If
    Else 'if tilted then only add a short delay and move to the 2nd part of the end of the ball
        vpmtimer.addtimer 200, "EndOfBall2 '"
    End If
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the CurrentPlayer)
'
Sub EndOfBall2()
    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots
    TurnOffPFLayers    'turns off any pf layers left on


    ' has the player won an extra-ball ? (might be multiple outstanding)
    If ExtraBallsAwards(CurrentPlayer)> 0 Then
        'debug.print "Extra Ball"

        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1

        ' if no more EB's then turn off any Extra Ball light if there was any
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
            LightShootAgain.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
        DMD CL("EXTRA BALL"), CL("SHOOT AGAIN"), "", eNone, eBlink, eNone, 1500, True, "vo_live_again"

        ' In this table an extra ball will have the skillshot and ball saver, so we reset the playfield for the new ball
        ResetForNewPlayerBall()

        ' Create a new ball in the shooters lane
        'CreateNewBall()
         AddMultiball 1
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0) Then
            ' debug.print "No More Balls, High Score Entry"
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
    TurnOffPFLayers 'turns off any pf layers left on
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
        DMDScoreNow

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()

        ' play a sound if more than 1 player
        If PlayersPlayingGame> 1 Then
            Select Case CurrentPlayer
                Case 1:DMD "", CL("PLAYER 1"), "", eNone, eNone, eNone, 1000, True, "vo_player1"
                Case 2:DMD "", CL("PLAYER 2"), "", eNone, eNone, eNone, 1000, True, "vo_player2"
                Case 3:DMD "", CL("PLAYER 3"), "", eNone, eNone, eNone, 1000, True, "vo_player3"
                Case 4:DMD "", CL("PLAYER 4"), "", eNone, eNone, eNone, 1000, True, "vo_player4"
            End Select
        Else
            DMD "", CL("PLAYER 1"), "", eNone, eNone, eNone, 1000, True, ""
        End If
    End If
End Sub

Sub TurnOffPFLayers
    Goku_PoweredUp.visible = 0        'Turns off Goku pf overlays/ sounds/ lights
  StopSound "vo_AuraPoweredUp"        '
  LightSeqGokuPU.StopPlay             '
    Goku_SS2.visible = 0                '
  StopSound "vo_AuraSS2"            '
  LightSeqGokuSS2.StopPlay          '
    Goku_Blue.visible = 0             '
  StopSound "vo_auraBlueWithVocals"   '
  LightSeqGokuBlue.StopPlay         '
    Goku_Blue.visible = 0                   '
    Goku_UltraInstinct.visible = 0          '
    Goku_UltraInstinct20.visible = 0        '
  StopSound "vo_GokuScreamPowerUp"        '
    If bTrunksCompleted (CurrentPlayer) = False Then
    Trunks_PoweredUp.visible = 0        'Layers only turn off if mode not completed
  End If
    If bVegetaCompleted (CurrentPlayer) = False Then
    Vegeta_PoweredUp.visible = 0        'Layers only turn off if mode not completed
  End If
    If bPiccoloCompleted (CurrentPlayer) = False Then
    Piccolo_PoweredUp.visible = 0       'Layers only turn off if mode not completed
  End If
  Cell_L.State = 0                       'Turns off OTHER pf overlay
  Zamasu_L.State = 0            '
  GokuBlack_L.State = 0         '
  Frieza_L.State = 0           '
    pf_GrateApeTransformation.visible = 0   '
    plastic_KaiMode.visible = 0
  Magnet.Enabled = False:DOF 132, DOFOFF  'Turn off magnet
    LightMOON.STATE = 0                     'Turn off moon light
    LoadLUT                                 'reset LUT
  If bKaiModeStarted = True Then
    Goku_UltraInstinct.visible = 1      'All hero layers on for KaiMode, they turn back off after mode
    Trunks_PoweredUp.visible = 1    '
    Vegeta_PoweredUp.visible = 1    '
    Piccolo_PoweredUp.visible = 1   '
    plastic_KaiMode.visible = 1     '
    li036.State = 2
    li040.State = 2
    li042.State = 2
    li044.State = 2
  End If
End Sub
' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    'debug.print "End Of Game"
    bGameInPLay = False
    ' just ended your game then play the end of game tune
    ' vpmtimer.AddTimer 2500, "PlayEndQuote '"
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball

    ' set any lights for the attract mode
    GiOff
    StartAttractMode
' you may wish to light any Game Over Light you may have
End Sub

'this calculates the ball number in play
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
  StopSound "vo_kamehamehaL1"
    StopSound "vo_kamehamehaL2"
    StopSound "vo_kamehamehaR1"
    StopSound "vo_kamehamehaR2"
  'debug.print "Multiball variable: " & bMultiBallMode
  'debug.print "Current mode: " & CurrentMode(CurrentPlayer)
    Drain.DestroyBall

    BallsOnPlayfield = BallsOnPlayfield - 1
  'debug.print "balls left " & BallsOnPlayfield
    ' pretend to knock the ball into the ball storage mech
    PlaySoundAt "fx_drain", Drain
    DOF 109, DOFPulse
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
            ' we kick the ball with the autoplunger
            bAutoPlunger = True
            AddMultiball 1
            ' you may wish to put something on a display or play a sound at this point
            ' stop the ballsaver timer during the launch ball saver time, but not during multiballs
            If NOT bMultiBallMode Then
                DMD "_", CL("BALL SAVED"), "_", eNone, eBlinkfast, eNone, 2500, True, "vo_senzu" &RndNbr(5)
            'BallSaverTimerExpired_Timer 'enable this line to stop the ballsaver timer
            End If
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
                ' AND in a multi-ball??
            If(BallsOnPlayfield = 1) AND CurrentMode(CurrentPlayer) <> 9 Then
                If(bMultiBallMode = True) then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ' turn off any multiball specific lights
                    If Mode(CurrentPlayer, 0) = 0 Then
                        ChangeGi white
                        ChangeGIIntensity 1
                    End If
                    'stop any multiball modes of this game
                    StopMBmodes
                    ' you may wish to change any music over at this point
                    changesong
                End If
            End If

            ' was that the last ball on the playfield
            If BallsOnPlayfield = 0 Then
        If CurrentMode(CurrentPlayer) = 8 Then
          StartNextMode
        Else If CurrentMode(CurrentPlayer) = 9 Then
          ChangeGi white
          ChangeGIIntensity 1
          StopMode
          StopEndOfBallMode
          vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, if there is no animation then move to the end of ball
        Else
          ' End Mode and timers
          'StopSong Song
          ChangeGi white
          ChangeGIIntensity 1
          ' Show the end of ball animation
          ' and continue with the end of ball
          ' DMD something?
          StopEndOfBallMode
          vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, if there is no animation then move to the end of ball
        End If
        End If
      End If
      'If BallsOnPlayfield = 2 Then
      ' 'Do nothing, wait for loopback for 1
      'End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    If bPlayIntro Then PlaySound "vo_game_start":bPlayIntro = False
    PlaySoundAt "fx_sensor", swPlungerRest
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
    'LaunchLight.State = 2
    ' be sure to update the Scoreboard after the animations, if any
    'Start the skillshot lights & variables if any
    If bSkillShotReady Then
        PlaySong "mu_plunger"
        UpdateSkillshot()
        ' show the message to shoot the ball in case the player has fallen sleep
        swPlungerRest.TimerEnabled = 1
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

Sub swPLunger2_Hit 'extra trigger to detect a ball resting down on the plunger
    ' kick the ball in play if the bAutoPlunger flag is on
    bBallInPlungerLane = True 'unsure this flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        vpmtimer.addtimer 1200, "PlungerIM.AutoFire:DOF 113, DOFPulse:DOF 130, DOFPulse:PlaySoundAt ""fx_kicker"", swPlunger '"
    End If
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    lighteffect 6
    bBallInPlungerLane = False
    bAutoPlunger = False           'disable the autoplunger as the ball has left the plunger lane
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
        ChangeSong
        ResetSkillShotTimer.Enabled = 1
    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    End If
' turn off LaunchLight
' LaunchLight.State = 0
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds
Sub swPlungerRest_Timer
    IF bOnTheFirstBall Then
        Select Case RndNbr(5)
            Case 1:DMD CL("YOU WILL"), CL("NEVER WIN"), "_", eNone, eNone, eNone, 2800, True, ""
            Case 2:DMD CL("WHATS THE MATTER"), CL("LOST YA DRAGONBALLS"), "_", eNone, eNone, eNone, 2800, True, ""
            Case 3:DMD CL("I AM VEGETA"), CL("BOW BEFORE ME"), "_", eNone, eNone, eNone, 2800, True, "sfx_thunder" &RndNbr(9):FlashEffect 2:ZeusF
            Case 4:DMD CL("HEY YOU"), CL("SHOOT THE BALL"), "_", eNone, eNone, eNone, 2800, True, ""
            Case 5:DMD CL("WHATS HIS"), CL("POWER LEVEL RADITZ"), "_", eNone, eNone, eNone, 2800, True, ""
        End Select
    Else
        Select Case RndNbr(4)
            Case 1:DMD CL("ARE YOU"), CL("SLEEPING"), "_", eNone, eNone, eNone, 2000, True, ""
            Case 2:DMD CL("WHAT ARE"), CL("YOU WAITING FOR"), "_", eNone, eNone, eNone, 2000, True, ""
            Case 3:DMD CL("HEY"), CL("PULL THE PLUNGER"), "_", eNone, eNone, eNone, 2000, True, ""
            Case 4:DMD CL("ARE YOU PLAYING"), CL("THIS GAME"), "_", eNone, eNone, eNone, 2000, True, ""
        End Select
    End If
End Sub

Sub EnableBallSaver(seconds)
    'debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
    BallSaverTimerExpired.Enabled = False
    BallSaverSpeedUpTimer.Enabled = False
    BallSaverTimerExpired.Interval = 1000 * seconds
    BallSaverTimerExpired.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.State = 2
    li006.BlinkInterval = 160
    BallSave_Text.visible = 1 'SenzuBeanLayerON
    li006.State = 2           'God mode light
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimerExpired_Timer()
    'debug.print "Ballsaver ended"
    BallSaverTimerExpired.Enabled = False
    BallSaverSpeedUpTimer.Enabled = False 'ensure this timer is also stopped
    ' clear the flag
    bBallSaverActive = False
    ' if you have a ball saver light then turn it off at this point
    LightShootAgain.State = 0
    li006.State = 0           'God mode light
    BallSave_Text.visible = 0 'SenzuBeanLayerOFF
    ' if the table uses the same lights for the extra ball or replay then turn them on if needed
    If ExtraBallsAwards(CurrentPlayer)> 0 Then
        LightShootAgain.State = 1
    End If
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    'debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
    li006.BlinkInterval = 80
    li006.State = 2
End Sub

' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board

Sub AddScore(points) 'normal score routine
    If Tilted Then Exit Sub
    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer) * FatesMultiplier
' you may wish to check to see if the player has gotten a replay
End Sub

Sub AddScore2(points) 'used in jackpots, skillshots, combos, and bonus as they doe not use the PlayfieldMultiplier
    If Tilted Then Exit Sub
    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + points
End Sub

' Add bonus to the bonuspoints AND update the score board

Sub AddBonus(points) 'not used in this table, since there are many different bonus items.
    If Tilted Then Exit Sub
    ' add the bonus to the current players bonus variable
    BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
End Sub

' Add some points to the current Jackpot.
'
Sub AddJackpot(points)
    ' Jackpots only generally increment in multiball mode AND not tilted
    ' but this doesn't have to be the case
    If Tilted Then Exit Sub

    ' If(bMultiBallMode = True) Then
    Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + points
    DMD "_", CL("INCREASED JACKPOT"), "_", eNone, eNone, eNone, 1000, True, ""
' you may wish to limit the jackpot to a upper limit, ie..
' If (Jackpot >= 6000000) Then
'   Jackpot = 6000000
'   End if
'End if
End Sub

Sub AddSuperJackpot(points) 'not used in this table
    If Tilted Then Exit Sub
End Sub

Sub AddBonusMultiplier(n)
    Dim NewBonusLevel
    ' if not at the maximum bonus level
    if(BonusMultiplier(CurrentPlayer) + n <= MaxBonusMultiplier) then
        ' then add and set the lights
        NewBonusLevel = BonusMultiplier(CurrentPlayer) + n
        SetBonusMultiplier(NewBonusLevel)
        DMD "_", CL("BONUS X " &NewBonusLevel), "_", eNone, eBlink, eNone, 2000, True, ""
    Else
        AddScore2 500000
        DMD "_", CL("500000"), "_", eNone, eNone, eNone, 1000, True, ""
    End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly

Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level
    UpdateBonusXLights(Level)
End Sub

Sub UpdateBonusXLights(Level) 'no lights in this table
    ' Update the lights
    Select Case Level
    '        Case 1:li021.State = 0:li022.State = 0:li023.State = 0:li024.State = 0
    '        Case 2:li021.State = 1:li022.State = 0:li023.State = 0:li024.State = 0
    '        Case 3:li021.State = 1:li022.State = 1:li023.State = 0:li024.State = 0
    '        Case 4:li021.State = 1:li022.State = 1:li023.State = 1:li024.State = 0
    '        Case 5:li021.State = 1:li022.State = 1:li023.State = 1:li024.State = 1
    End Select
End Sub

Sub AddPlayfieldMultiplier(n)
    Dim NewPFLevel
    ' if not at the maximum level x
    if(PlayfieldMultiplier(CurrentPlayer) + n <= MaxMultiplier) then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) + n
        SetPlayfieldMultiplier(NewPFLevel)
        DMD "_", CL("PLAYFIELD X " &NewPFLevel), "_", eNone, eBlink, eNone, 2000, True, "sfx_thunder" &RndNbr(7)
        LightEffect 4
    ' Play a voice sound
    Else 'if the max is already lit
        AddScore2 500000
        DMD "_", CL("500000"), "_", eNone, eNone, eNone, 2000, True, ""
    End if
    ' restart the PlayfieldMultiplier timer to reduce the multiplier
    PFXTimer.Enabled = 0
    PFXTimer.Enabled = 1
End Sub

Sub PFXTimer_Timer
    DecreasePlayfieldMultiplier
End Sub

Sub DecreasePlayfieldMultiplier 'reduces by 1 the playfield multiplier
    Dim NewPFLevel
    ' if not at 1 already
    if(PlayfieldMultiplier(CurrentPlayer)> 1) then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) - 1
        SetPlayfieldMultiplier(NewPFLevel)
    Else
        PFXTimer.Enabled = 0
    End if
End Sub

' Set the Playfield Multiplier to the specified level AND set any lights accordingly

Sub SetPlayfieldMultiplier(Level)
    ' Set the multiplier to the specified level
    PlayfieldMultiplier(CurrentPlayer) = Level
    UpdatePFXLights(Level)
End Sub

Sub UpdatePFXLights(Level) 'no lights in this table
    ' Update the playfield multiplier lights
    Select Case Level
    '        Case 1:li025.State = 0:li026.State = 0:li027.State = 0:li027.State = 0
    '        Case 2:li025.State = 1:li026.State = 0:li027.State = 0:li027.State = 0
    '        Case 3:li025.State = 0:li026.State = 1:li027.State = 0:li027.State = 0
    '        Case 4:li025.State = 0:li026.State = 0:li027.State = 1:li027.State = 0
    '        Case 5:li025.State = 0:li026.State = 0:li027.State = 0:li027.State = 1
    End Select
' perhaps show also the multiplier in the DMD?
End Sub

Sub AwardExtraBall()
    '   If NOT bExtraBallWonThisBall Then 'in this table you can win several extra balls
    DMD "_", CL("EXTRA BALL WON"), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("fx_Knocker", 108, DOFPulse, DOFKnocker)
    DOF 130, DOFPulse
    PLaySound "vo_extra_ball"
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    'bExtraBallWonThisBall = True
    light009.State = 0        'turn off extra ball light
    LightShootAgain.State = 1 'light the shoot again lamp
    GiEffect 1
    LightEffect 2
'    END If
End Sub

Sub AwardSpecial()
    DMD "_", CL("EXTRA GAME WON"), "_", eNone, eBlink, eNone, 2000, True, SoundFXDOF("fx_Knocker", 108, DOFPulse, DOFKnocker)
    DOF 130, DOFPulse
    Credits = Credits + 1
    AddScore2 3000000 '3 mill only for this table
    If bFreePlay = False Then DOF 121, DOFOn
    LightEffect 2
    GiEffect 1
    Light010.State = 0 'turn off special light
End Sub

Sub AwardJackpot()     'only used for the final mode
    DMD CL("JACKPOT"), CL(FormatScore(Jackpot(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_Jackpot"
    DOF 137, DOFPulse
    AddScore2 Jackpot(CurrentPlayer)
    Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + 100000
    LightEffect 2
    GiEffect 1
    FlashEffect 1
End Sub

Sub AwardSuperJackpot() 'not used in this table as there are several superjackpots but I keep it as a reference
    DMD CL("SUPER JACKPOT"), CL(FormatScore(SuperJackpot(CurrentPlayer) ) ), "d_border", eNone, eBlink, eNone, 2000, True, "vo_super_jackpot"
    DOF 137, DOFPulse
    AddScore2 SuperJackpot(CurrentPlayer)
    LightEffect 2
    GiEffect 1
End Sub

Sub AwardSkillshot()
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL("SKILLSHOT"), CL(FormatScore(SkillshotValue(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_skillshot"
    DOF 127, DOFPulse
    Addscore2 SkillShotValue(CurrentPlayer)
    ' increment the skillshot value with 50.000
    SkillShotValue(CurrentPlayer) = SkillShotValue(CurrentPlayer) + 50000
    'do some light show
    GiEffect 1
    LightEffect 2
End Sub

Sub AwardSuperSkillshot()
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL("SUPER SKILLSHOT"), CL(FormatScore(SuperSkillshotValue(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_superskillshot"
    DOF 138, DOFPulse
    Addscore2 SuperSkillshotValue(CurrentPlayer)
    ' increment the skillshot value with 500.000
    SuperSkillshotValue(CurrentPlayer) = SuperSkillshotValue(CurrentPlayer) + 500000
    'do some light show
    GiEffect 1
    LightEffect 2
End Sub

Sub AwardFreakySkillshot()
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL("LEGENDARY SKILLSHOT"), CL(FormatScore(FreakySkillshotValue(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_freakyskillshot"
    DOF 138, DOFPulse
    Addscore2 FreakySkillshotValue(CurrentPlayer)
    ' increment the skillshot value with 500.000
    FreakySkillshotValue(CurrentPlayer) = FreakySkillshotValue(CurrentPlayer) + 500000
    'do some light show
    GiEffect 1
    LightEffect 2
End Sub

Sub aSkillshotTargets_Hit(idx) 'stop the skillshot if any other target/switch is hit
    If bSkillshotReady then ResetSkillShotTimer_Timer
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
    If(x <> "") then Credits = CInt(x) Else Credits = 0:If bFreePlay = False Then DOF 121, DOFOff:End If
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
    SaveValue cGameName, "TotalGamesPlayed", TotalGamesPlayed
End Sub

Sub Reseths
    HighScoreName(0) = "AAA"
    HighScoreName(1) = "BBB"
    HighScoreName(2) = "CCC"
    HighScoreName(3) = "DDD"
    HighScore(0) = 1500000
    HighScore(1) = 1400000
    HighScore(2) = 1300000
    HighScore(3) = 1200000
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
    tmp = Score(CurrentPlayer)

    If tmp> HighScore(0) Then 'add 1 credit for beating the highscore
        Credits = Credits + 1
        DOF 121, DOFOn
    End If

    If tmp> HighScore(3) Then
        PlaySound SoundFXDOF("fx_Knocker", 108, DOFPulse, DOFKnocker)
        DOF 130, DOFPulse
        HighScore(3) = tmp
        'Play HighScore sound
        'enter player's name
        HighScoreEntryInit()
    Else
        EndOfBallComplete()
    End If
End Sub

Sub HighScoreEntryInit()
    hsbModeActive = True
    PlaySound "vo_enterinitials"
    hsLetterFlash = 0

    hsEnteredDigits(0) = " "
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<" ' < is back arrow
    hsCurrentLetter = 1
    DMDFlush()
    HighScoreDisplayNameNow()

    HighScoreFlashTimer.Interval = 250
    HighScoreFlashTimer.Enabled = True
End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<") then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit> 0) then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayNameNow()
        end if
    end if
End Sub

Sub HighScoreDisplayNameNow()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreDisplayName()
    Dim i
    Dim TempTopStr
    Dim TempBotStr

    TempTopStr = "YOUR NAME:"
    dLine(0) = ExpandLine(TempTopStr)
    DMDUpdate 0

    TempBotStr = "    > "
    if(hsCurrentDigit> 0) then TempBotStr = TempBotStr & hsEnteredDigits(0)
    if(hsCurrentDigit> 1) then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit> 2) then TempBotStr = TempBotStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3) then
        if(hsLetterFlash <> 0) then
            TempBotStr = TempBotStr & "_"
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit <1) then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit <2) then TempBotStr = TempBotStr & hsEnteredDigits(2)

    TempBotStr = TempBotStr & " <    "
    dLine(1) = ExpandLine(TempBotStr)
    DMDUpdate 1
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2) then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = False

    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ") then
        hsEnteredName = "YOU"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
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

'************************************
'       LUT - Darkness control
' 10 normal level & 10 warmer levels
'************************************

Dim bLutActive, LUTImage

Sub LoadLUT
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1) MOD 22:UpdateLUT:SaveLUT:SetLUTLine "Color LUT image " & table1.ColorGradeImage:End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT6"
        Case 7:table1.ColorGradeImage = "LUT7"
        Case 8:table1.ColorGradeImage = "LUT8"
        Case 9:table1.ColorGradeImage = "LUT9"
        Case 10:table1.ColorGradeImage = "LUT10"
        Case 11:table1.ColorGradeImage = "LUT Warm 0"
        Case 12:table1.ColorGradeImage = "LUT Warm 1"
        Case 13:table1.ColorGradeImage = "LUT Warm 2"
        Case 14:table1.ColorGradeImage = "LUT Warm 3"
        Case 15:table1.ColorGradeImage = "LUT Warm 4"
        Case 16:table1.ColorGradeImage = "LUT Warm 5"
        Case 17:table1.ColorGradeImage = "LUT Warm 6"
        Case 18:table1.ColorGradeImage = "LUT Warm 7"
        Case 19:table1.ColorGradeImage = "LUT Warm 8"
        Case 20:table1.ColorGradeImage = "LUT Warm 9"
        Case 21:table1.ColorGradeImage = "LUT Warm 10"
    End Select
End Sub



' New LUT postit
Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetLUTLine(String)
    Dim Index
    Dim xFor
    Index = 1
    LUBack.imagea = "PostItNote"
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
    SetLUTLine ""
    LUBack.imagea = "PostitBL"
End Sub

' *************************************************************************
'   JP's Reduced Display Driver Functions (based on script by Black)
' only 5 effects: none, scroll left, scroll right, blink and blinkfast
' 3 Lines, treats all 3 lines as text.
' 1st and 2nd lines are 20 characters long
' 3rd line is just 1 character
' Example format:
' DMD "text1","text2","backpicture", eNone, eNone, eNone, 250, True, "sound"
' Short names:
' dq = display queue
' de = display effect
' *************************************************************************

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

Dim dLine(2)
Dim deCount(2)
Dim deCountEnd(2)
Dim deBlinkCycle(2)

Dim dqText(2, 64)
Dim dqEffect(2, 64)
Dim dqTimeOn(64)
Dim dqbFlush(64)
Dim dqSound(64)

Dim FlexDMD
Dim DMDScene

Sub DMD_Init() 'default/startup values
    Dim i, j
    If UseFlexDMD Then
        Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
        If Not FlexDMD is Nothing Then
            If FlexDMDHighQuality Then
                FlexDMD.TableFile = Table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 256
                FlexDMD.Height = 64
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                For i = 0 to 40
                    DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
                    Digits(i).Visible = False
                Next
                digitgrid.Visible = False
                For i = 0 to 19 ' Top
                    DMDScene.GetImage("Dig" & i).SetBounds 8 + i * 12, 6, 12, 22
                Next
                For i = 20 to 39 ' Bottom
                    DMDScene.GetImage("Dig" & i).SetBounds 8 + (i - 20) * 12, 34, 12, 22
                Next
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            Else
                FlexDMD.TableFile = Table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 128
                FlexDMD.Height = 32
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                For i = 0 to 40
                    DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
                    Digits(i).Visible = False
                Next
                digitgrid.Visible = False
                For i = 0 to 19 ' Top
                    DMDScene.GetImage("Dig" & i).SetBounds 4 + i * 6, 3, 6, 11
                Next
                For i = 20 to 39 ' Bottom
                    DMDScene.GetImage("Dig" & i).SetBounds 4 + (i - 20) * 6, 17, 6, 11
                Next
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            End If
        End If
    End If

    DMDFlush()
    deSpeed = 20
    deBlinkSlowRate = 10
    deBlinkFastRate = 5
    For i = 0 to 2
        dLine(i) = Space(20)
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
        dqTimeOn(i) = 0
        dqbFlush(i) = True
        dqSound(i) = ""
    Next
    dLine(2) = " "
    For i = 0 to 2
        For j = 0 to 64
            dqText(i, j) = ""
            dqEffect(i, j) = eNone
        Next
    Next
    DMD dLine(0), dLine(1), dLine(2), eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDFlush()
    Dim i
    DMDTimer.Enabled = False
    DMDEffectTimer.Enabled = False
    dqHead = 0
    dqTail = 0
    For i = 0 to 2
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
    Next
End Sub

Sub DMDScore()
    Dim tmp, tmp1, tmp1a, tmp1b, tmp2
    if(dqHead = dqTail) Then
        ' default when no modes are active
        tmp = RL(FormatScore(Score(Currentplayer) ) )
        tmp1 = FL("PLAYER " &CurrentPlayer, "BALL " & Balls)
        tmp2 = "d_border"
        'info on the second line & background
    Select Case Mode(CurrentPlayer, 0)
            Case 0 'no battle active
                If bMedusaMBStarted Then
                    tmp2 = "d_zeus"
                End If
            Case 1 ' CELL/Minotaur
                tmp2 = "d_minotaur"
                Select Case ModeStep
          Case 1, 2, 3:tmp1 = RL("SHOOT LIT SHOT")
                    Case 4:tmp1 = RL("SHOOT THE SCOOP")
                End Select
            Case 2 ' ZAMASU/Hydra
                tmp2 = "d_hydra"
                Select Case ModeStep
                    Case 1, 2, 3, 4:tmp1 = RL("SHOOT LIT SHOT")
                    Case 5:tmp1 = RL("SHOOT THE SCOOP")
                End Select
            Case 3 ' GOKU BLACK/Cerberus
                tmp2 = "d_cerberus"
                Select Case ModeStep
                    Case 1:tmp1 = RL("SHOOT LIT SHOTS")
                    Case 2:tmp1 = RL("SHOOT THE SCOOP")
                End Select
            Case 4 ' GOLDEN FRIEZA/Medusa
                tmp2 = "d_medusa"
                Select Case ModeStep
                    Case 1, 2:tmp1 = RL("SHOOT LIT SHOTS")
                    Case 3:tmp1 = RL("SHOOT THE SCOOP")
                End Select
            Case 5 ' TRUNKS/Ares
                tmp2 = "d_ares"
                Select Case ModeStep
                    Case 1:tmp1 = RL("SPINNERS LEFT " & (100-SpinnerHits) )
                    Case 2:tmp1 = RL("SHOOT THE SCOOP")
                End Select
            Case 6 ' VEGETA/Poseidon
                tmp2 = "d_poseidon"
                Select Case ModeStep
                    Case 1:tmp1 = RL("HIT LIT TARGETS")
                    Case 2:tmp1 = RL("SHOOT THE SCOOP")
                End Select
            Case 7 ' PICCOLO/Hades
                tmp2 = "d_hades"
                tmp1 = RL("HITS LEFT " & (5-TrapDoorHits) )
            Case 8 ' GOKU/Zeus
                tmp2 = "d_zeus"
                Select Case ModeStep
                    Case 1:tmp1 = RL("HIT GOKU TARGETS")
                    Case 2:tmp1 = RL("SHOOT THE SCOOP")
                End Select
            Case 9 ' KING KAI/God or SUPREME KAI/Demi-God mode
                tmp2 = "d_kai"
                tmp1 = RL("SHOOT JACKPOTS")
        End Select

    End If
    DMD tmp, tmp1, tmp2, eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
    if(dqTail <dqSize) Then
        if(Text0 = "_") Then
            dqEffect(0, dqTail) = eNone
            dqText(0, dqTail) = "_"
        Else
            dqEffect(0, dqTail) = Effect0
            dqText(0, dqTail) = ExpandLine(Text0)
        End If
        if(Text1 = "_") Then
            dqEffect(1, dqTail) = eNone
            dqText(1, dqTail) = "_"
        Else
            dqEffect(1, dqTail) = Effect1
            dqText(1, dqTail) = ExpandLine(Text1)
        End If

        if(Text2 = "_") Then
            dqEffect(2, dqTail) = eNone
            dqText(2, dqTail) = "_"
        Else
            dqEffect(2, dqTail) = Effect2
            dqText(2, dqTail) = Text2 'it is always 1 letter in this table
        End If

        dqTimeOn(dqTail) = TimeOn
        dqbFlush(dqTail) = bFlush
        dqSound(dqTail) = Sound
        dqTail = dqTail + 1
        if(dqTail = 1) Then
            DMDHead()
        End If
    End If
End Sub

Sub DMDHead()
    Dim i
    deCount(0) = 0
    deCount(1) = 0
    deCount(2) = 0
    DMDEffectTimer.Interval = deSpeed

    For i = 0 to 2
        Select Case dqEffect(i, dqHead)
            Case eNone:deCountEnd(i) = 1
            Case eScrollLeft:deCountEnd(i) = Len(dqText(i, dqHead) )
            Case eScrollRight:deCountEnd(i) = Len(dqText(i, dqHead) )
            Case eBlink:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
            Case eBlinkFast:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
        End Select
    Next
    if(dqSound(dqHead) <> "") Then
        PlaySound(dqSound(dqHead) )
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
    if(dqHead = dqTail) Then
        if(dqbFlush(Head) = True) Then
            DMDScoreNow()
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

    For i = 0 to 2
        if(deCount(i) <> deCountEnd(i) ) Then
            deCount(i) = deCount(i) + 1

            select case(dqEffect(i, dqHead) )
                case eNone:
                    Temp = dqText(i, dqHead)
                case eScrollLeft:
                    Temp = Right(dLine(i), 19)
                    Temp = Temp & Mid(dqText(i, dqHead), deCount(i), 1)
                case eScrollRight:
                    Temp = Mid(dqText(i, dqHead), 21 - deCount(i), 1)
                    Temp = Temp & Left(dLine(i), 19)
                case eBlink:
                    BlinkEffect = True
                    if((deCount(i) MOD deBlinkSlowRate) = 0) Then
                        deBlinkCycle(i) = deBlinkCycle(i) xor 1
                    End If

                    if(deBlinkCycle(i) = 0) Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(20)
                    End If
                case eBlinkFast:
                    BlinkEffect = True
                    if((deCount(i) MOD deBlinkFastRate) = 0) Then
                        deBlinkCycle(i) = deBlinkCycle(i) xor 1
                    End If

                    if(deBlinkCycle(i) = 0) Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(20)
                    End If
            End Select

            if(dqText(i, dqHead) <> "_") Then
                dLine(i) = Temp
                DMDUpdate i
            End If
        End If
    Next

    if(deCount(0) = deCountEnd(0) ) and(deCount(1) = deCountEnd(1) ) and(deCount(2) = deCountEnd(2) ) Then

        if(dqTimeOn(dqHead) = 0) Then
            DMDFlush()
        Else
            if(BlinkEffect = True) Then
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

Function ExpandLine(TempStr) 'id is the number of the dmd line
    If TempStr = "" Then
        TempStr = Space(20)
    Else
        if Len(TempStr)> Space(20) Then
            TempStr = Left(TempStr, Space(20) )
        Else
            if(Len(TempStr) <20) Then
                TempStr = TempStr & Space(20 - Len(TempStr) )
            End If
        End If
    End If
    ExpandLine = TempStr
End Function

Function FormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
    dim i
    dim NumString

    NumString = CStr(abs(Num) )

    For i = Len(NumString) -3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1) ) then
            NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1) ) + 48) & right(NumString, Len(NumString) - i)
        end if
    Next
    FormatScore = NumString
End function

Function FL(NumString1, NumString2) 'Fill line
    Dim Temp, TempStr
    Temp = 20 - Len(NumString1) - Len(NumString2)
    TempStr = NumString1 & Space(Temp) & NumString2
    FL = TempStr
End Function

Function CL(NumString) 'center line
    Dim Temp, TempStr
    If Len(NumString)> 20 Then NumString = left(NumString, 20)
    Temp = (20 - Len(NumString) ) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(NumString) 'right line
    Dim Temp, TempStr
    Temp = 20 - Len(NumString)
    TempStr = Space(Temp) & NumString
    RL = TempStr
End Function

'**************
' Update DMD
'**************

Sub DMDUpdate(id)
    Dim digit, value
    If UseFlexDMD Then FlexDMD.LockRenderThread
    Select Case id
        Case 0 'top text line
            For digit = 0 to 19
                DMDDisplayChar mid(dLine(0), digit + 1, 1), digit
            Next
        Case 1 'bottom text line
            For digit = 20 to 39
                DMDDisplayChar mid(dLine(1), digit -19, 1), digit
            Next
        Case 2 ' back image - back animations
            If dLine(2) = "" OR dLine(2) = " " Then dLine(2) = "d_border"
            Digits(40).ImageA = dLine(2)
            If UseFlexDMD Then DMDScene.GetImage("Back").Bitmap = FlexDMD.NewImage("", "VPX." & dLine(2) & "&dmd=2").Bitmap
    End Select
    If UseFlexDMD Then FlexDMD.UnlockRenderThread
End Sub

Sub DMDDisplayChar(achar, adigit)
    If achar = "" Then achar = " "
    achar = ASC(achar)
    Digits(adigit).ImageA = Chars(achar)
    If UseFlexDMD Then DMDScene.GetImage("Dig" & adigit).Bitmap = FlexDMD.NewImage("", "VPX." & Chars(achar) & "&dmd=2&add").Bitmap
End Sub

'****************************
' JP's new DMD using flashers
'****************************

Dim Digits, Chars(255), Images(255)

DMDInit

Sub DMDInit
    Dim i
    Digits = Array(digit001, digit002, digit003, digit004, digit005, digit006, digit007, digit008, digit009, digit010, _
        digit011, digit012, digit013, digit014, digit015, digit016, digit017, digit018, digit019, digit020,            _
        digit021, digit022, digit023, digit024, digit025, digit026, digit027, digit028, digit029, digit030,            _
        digit031, digit032, digit033, digit034, digit035, digit036, digit037, digit038, digit039, digit040,            _
        digit041)
    For i = 0 to 255:Chars(i) = "d_empty":Next

    Chars(32) = "d_empty"
    Chars(33) = ""        '!
    Chars(34) = ""        '"
    Chars(35) = ""        '#
    Chars(36) = ""        '$
    Chars(37) = ""        '%
    Chars(38) = ""        '&
    Chars(39) = ""        ''
    Chars(40) = ""        '(
    Chars(41) = ""        ')
    Chars(42) = ""        '*
    Chars(43) = "d_plus"  '+
    Chars(44) = ""        '
    Chars(45) = "d_minus" '-
    Chars(46) = "d_dot"   '.
    Chars(47) = ""        '/
    Chars(48) = "d_0"     '0
    Chars(49) = "d_1"     '1
    Chars(50) = "d_2"     '2
    Chars(51) = "d_3"     '3
    Chars(52) = "d_4"     '4
    Chars(53) = "d_5"     '5
    Chars(54) = "d_6"     '6
    Chars(55) = "d_7"     '7
    Chars(56) = "d_8"     '8
    Chars(57) = "d_9"     '9
    Chars(60) = "d_less"  '<
    Chars(61) = ""        '=
    Chars(62) = "d_more"  '>
    Chars(64) = ""        '@
    Chars(65) = "d_a"     'A
    Chars(66) = "d_b"     'B
    Chars(67) = "d_c"     'C
    Chars(68) = "d_d"     'D
    Chars(69) = "d_e"     'E
    Chars(70) = "d_f"     'F
    Chars(71) = "d_g"     'G
    Chars(72) = "d_h"     'H
    Chars(73) = "d_i"     'I
    Chars(74) = "d_j"     'J
    Chars(75) = "d_k"     'K
    Chars(76) = "d_l"     'L
    Chars(77) = "d_m"     'M
    Chars(78) = "d_n"     'N
    Chars(79) = "d_o"     'O
    Chars(80) = "d_p"     'P
    Chars(81) = "d_q"     'Q
    Chars(82) = "d_r"     'R
    Chars(83) = "d_s"     'S
    Chars(84) = "d_t"     'T
    Chars(85) = "d_u"     'U
    Chars(86) = "d_v"     'V
    Chars(87) = "d_w"     'W
    Chars(88) = "d_x"     'X
    Chars(89) = "d_y"     'Y
    Chars(90) = "d_z"     'Z
    Chars(94) = "d_up"    '^
    '    Chars(95) = '_
    Chars(96) = "d_0a"  '0.
    Chars(97) = "d_1a"  '1. 'a
    Chars(98) = "d_2a"  '2. 'b
    Chars(99) = "d_3a"  '3. 'c
    Chars(100) = "d_4a" '4. 'd
    Chars(101) = "d_5a" '5. 'e
    Chars(102) = "d_6a" '6. 'f
    Chars(103) = "d_7a" '7. 'g
    Chars(104) = "d_8a" '8. 'h
    Chars(105) = "d_9a" '9. 'i
    Chars(106) = ""     'j
    Chars(107) = ""     'k
    Chars(108) = ""     'l
    Chars(109) = ""     'm
    Chars(110) = ""     'n
    Chars(111) = ""     'o
    Chars(112) = ""     'p
    Chars(113) = ""     'q
    Chars(114) = ""     'r
    Chars(115) = ""     's
    Chars(116) = ""     't
    Chars(117) = ""     'u
    Chars(118) = ""     'v
    Chars(119) = ""     'w
    Chars(120) = ""     'x
    Chars(121) = ""     'y
    Chars(122) = ""     'z
    Chars(123) = ""     '{
    Chars(124) = ""     '|
    Chars(125) = ""     '}
    Chars(126) = ""     '~
End Sub

'********************
' Real Time updates
'********************
'used for all the real time updates

Sub Realtime_Timer
    RollingUpdate
End Sub

' flippers top animations

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub
Sub LeftFlipper001_Animate:LeftFlipperTop001.RotZ = LeftFlipper001.CurrentAngle:End Sub
Sub RightFlipper001_Animate:RightFlipperTop001.RotZ = RightFlipper001.CurrentAngle:End Sub
Sub LeftFlipper2_Animate:LeftFlipperTop002.RotZ = LeftFlipper2.CurrentAngle:End Sub
Sub RightFlipper2_Animate:RightFlipperTop002.RotZ = RightFlipper2.CurrentAngle:End Sub

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
' 12 colors: red, orange, amber, yellow...
'******************************************

'colors
Const red = 1
Const orange = 2
Const amber = 3
Const yellow = 4
Const darkgreen = 5
Const green = 6
Const blue = 7
Const darkblue = 8
Const purple = 9
Const white = 10
Const teal = 11
Const ledwhite = 12

Sub SetLightColor(n, col, stat) 'stat 0 = off, 1 = on, 2 = blink, -1= no change
    Select Case col
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
        Case white 'bulb
            n.color = RGB(193, 91, 0)
            n.colorfull = RGB(255, 197, 143)
        Case teal
            n.color = RGB(1, 64, 62)
            n.colorfull = RGB(2, 128, 126)
        Case ledwhite
            n.color = RGB(255, 197, 143)
            n.colorfull = RGB(255, 252, 224)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

Sub SetFlashColor(n, col, stat) 'stat 0 = off, 1 = on, -1= no change - no blink for the flashers, use FlashForMs
    Select Case col
        Case red
            n.color = RGB(255, 0, 0)
        Case orange
            n.color = RGB(255, 64, 0)
        Case amber
            n.color = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 64, 64)
        Case purple
            n.color = RGB(128, 0, 192)
        Case white 'bulb
            n.color = RGB(255, 197, 143)
            stat = 0
        Case teal
            n.color = RGB(2, 128, 126)
        Case ledwhite
            n.color = RGB(255, 252, 224)
    End Select
    If stat <> -1 Then
        n.Visible = stat
    End If
End Sub

'*************************
' Rainbow Changing Lights
'*************************

Dim RGBStep, RGBFactor, rRed, rGreen, rBlue, RainbowLights

Sub StartRainbow(n) 'n is a collection
    set RainbowLights = n
    RGBStep = 0
    RGBFactor = 5
    rRed = 255
    rGreen = 0
    rBlue = 0
    RainbowTimer.Enabled = 1
End Sub

Sub StopRainbow()
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

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo
    Dim ii
    'info goes in a loop only stopped by the credits and the startkey
    If Score(1) Then
        DMD CL("LAST SCORE"), CL("PLAYER 1 " &FormatScore(Score(1) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(2) Then
        DMD CL("LAST SCORE"), CL("PLAYER 2 " &FormatScore(Score(2) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(3) Then
        DMD CL("LAST SCORE"), CL("PLAYER 3 " &FormatScore(Score(3) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(4) Then
        DMD CL("LAST SCORE"), CL("PLAYER 4 " &FormatScore(Score(4) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    DMD "", CL("GAME OVER"), "", eNone, eBlink, eNone, 2000, False, ""
    If bFreePlay Then
        DMD "", CL("FREE PLAY"), "", eNone, eBlink, eNone, 2000, False, ""
    Else
        If Credits> 0 Then
            DMD CL("CREDITS " & Credits), CL("PRESS START"), "", eNone, eBlink, eNone, 2000, False, ""
        Else
            DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 2000, False, ""
        End If
    End If
    DMD "        CHEESE AND", "          JPSALAS", "d_jppresents", eNone, eNone, eNone, 2000, False, ""
    DMD "", "", "d_title", eNone, eNone, eNone, 4000, False, ""
    DMD "ORIGINAL", "DESIGN T-800", "d_t800", eNone, eNone, eNone, 2000, False, ""
    DMD "", CL("ROM VERSION " &myversion), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("HIGHSCORES"), Space(20), "", eScrollLeft, eScrollLeft, eNone, 20, False, ""
    DMD CL("HIGHSCORES"), "", "", eBlinkFast, eNone, eNone, 1000, False, ""
    DMD CL("HIGHSCORES"), "1> " &HighScoreName(0) & " " &FormatScore(HighScore(0) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "2> " &HighScoreName(1) & " " &FormatScore(HighScore(1) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "3> " &HighScoreName(2) & " " &FormatScore(HighScore(2) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "4> " &HighScoreName(3) & " " &FormatScore(HighScore(3) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD Space(20), Space(20), "", eScrollLeft, eScrollLeft, eNone, 500, False, ""
End Sub

Sub StartAttractMode
    StartLightSeq
    DMDFlush
    ShowTableInfo
  If AttractMusicAlt = 1 Then
    PlaySong "mu_game_start_2"
  Else
    PlaySong "mu_game_start"
  End If
End Sub

Sub StopAttractMode
    StopRainbow
    DMDScoreNow
    LightSeqAttract.StopPlay
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqDiagUpRightOn, 25, 2
    LightSeqAttract.Play SeqStripe1VertOn, 25
    LightSeqAttract.Play SeqClockRightOn, 180, 2
    LightSeqAttract.Play SeqFanLeftUpOn, 50, 2
    LightSeqAttract.Play SeqFanRightUpOn, 50, 2
    LightSeqAttract.Play SeqScrewRightOn, 50, 2
    LightSeqAttract.Play SeqDiagDownLeftOn, 25, 2
    LightSeqAttract.Play SeqStripe2VertOn, 25, 2
    LightSeqAttract.Play SeqFanLeftDownOn, 50, 2
    LightSeqAttract.Play SeqFanRightDownOn, 50, 2
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

' droptargets, animations, timers, etc
Sub VPObjects_Init
    TrapdoorDown
    If FlipperPin Then
       Ppin.Visible = 1
       RubberPin.Visible = 1
       rpin.Collidable = 1
    Else
       Ppin.Visible = 0
       RubberPin.Visible = 0
       rpin.Collidable = 0
    End If

  'TESTING ONLY - BLOCKED OUTLANES - CHEESE
  Rubber005.Visible = 0
  Rubber010.Visible = 0
  rpin002.Collidable = 0
  rpin003.Collidable = 0
  rpin002.Visible = 0
  rpin003.Visible = 0
End Sub

' tables variables and Mode init
Dim bRotateLights
Dim bPlayIntro
Dim Mode(4, 9) '4 players, 8 modes
Dim BumperHits 'used for the skillshot and
Dim FreakySkillshotValue(4)
Dim ComboValue(4)
Dim ComboHits(4)
Dim ComboCount
Dim BumperAward
Dim OrbitHits
Dim RampHits
Dim CurrentMode(4)       'the current selected mode, used to increase the modes or battles
Dim bModeReady           'used to enable the mode at the start battle hole
Dim ModeStep             'use for the different steps during the modes/battles
Dim EndModeCountdown
Dim ZeusTargetsCompleted 'used in Zeus mode to count the time the Zeus targets has been completed
Dim ZeusHits
Dim ZeusHitsNeeded
Dim ZeusCount(4)  'hits needed to start zeus multiball
Dim bMedusaMBStarted
Dim MedussaX      'multiplier during medusa MB
Dim ExtraBallHits 'used in medusa multiball
Dim SpinnerHits   'used in Ares mode
Dim TrapDoorHits  'used in Hades mode
Dim bJackpotsEnabled
Dim bLockEnabled
Dim MinotaurJackpot(4)              'jackpot value
Dim bMinotaurMBStarted
Dim bKaiModeStarted
Dim bKaiWizardModeRan
Dim bTrunksCompleted (4)
Dim bVegetaCompleted (4)
Dim bPiccoloCompleted (4)
Dim bGokuCompleted (4)
Dim bCellCompleted (4)
Dim bZamasuCompleted (4)
Dim bGokuBlackCompleted (4)
Dim bFriezaCompleted (4)
Dim Minotaur1, Minotaur2, Minotaur3 'to check if all 3 holes has been hit
Dim bZeusMBStarted
Dim bZeusStarted
Dim bCCorpMBStarted
Dim PandoraEnabled                  'variable to check if this is already enabled
Dim PandorasHitsRight(4)            'targets hits
Dim PandorasHitsLeft(4)
Dim PandorasNeeded(4)               'number of hits needed to start the random award
Dim TreeHits(4)
Dim KickbackEnabled                 'variable to check if this is already enabled
Dim kickbackHitsRight(4)
Dim kickbackHitsLeft(4)
Dim kickbackNeeded(4)
Dim GodModeHits(4)   'the number of times the god mode targets has been hit
Dim GodModeNeeded(4) 'number of hits required to start
Dim FatesHits(4)     'number of hits to start the Fates mode: PlayfieldMultiplier
Dim bFatesStarted
Dim FatesMultiplier
Dim CBHits(4) 'captive ball hits
Dim BonusTargets(4)
Dim BonusRamps(4)
Dim BonusOrbits(4)
Dim BonusXHits(4)
Dim TotalMonsters(4)
Dim TotalGods(4)
Dim NeededScoopHits(4) 'number of hits to enable the lock, starts with 1 hit
Dim ScoopHits(4)

Sub Game_Init()        'called at the start of a new game
    Dim i, j
    'Init Variables
    KickbackEnabled = False
    PandoraEnabled = False
    bPlayIntro = True
    BallSaverTime = 20
    bExtraBallWonThisBall = False
    BumperHits = 0
    bRotateLights = True
    BumperAward = 5000
    ComboCount = 0
    OrbitHits = 0
    RampHits = 0
    EndModeCountdown = 0
    ZeusHits = 0
    ZeusHitsNeeded = 12
    ModeStep = 0
    bModeReady = False
    bMedusaMBStarted = False
    MedussaX = 1
    ExtraBallHits = 0
    SpinnerHits = 0
    TrapDoorHits = 0
    ZeusTargetsCompleted = 0
    bJackpotsEnabled = False
    bLockEnabled = False
    Minotaur1 = 0
    Minotaur2 = 0
    Minotaur3 = 0
    bMinotaurMBStarted = False
  bKaiModeStarted = False
  bKaiWizardModeRan = False
  bCCorpMBStarted = False
    bZeusMBStarted = False
  bZeusStarted = False
    bFatesStarted = False
    FatesMultiplier = 1
    For i = 0 to 4
        SkillshotValue(i) = 100000
        SuperSkillshotValue(i) = 1000000
        FreakySkillshotValue(i) = 1500000
        CurrentMode(i) = 0
        ComboValue(i) = 250000
        ZeusCount(i) = 0
        MinotaurJackpot(i) = 250000
        BallsInLock(i) = 0
        SuperJackpot(i) = 5000000
        Jackpot(i) = 1000000
        PandorasHitsRight(i) = 0
        PandorasHitsLeft(i) = 0
        PandorasNeeded(i) = 2
        TreeHits(i) = 0
        kickbackHitsRight(i) = 0
        kickbackHitsLeft(i) = 0
        kickbackNeeded(i) = 2
        GodModeHits(i) = 0
        GodModeNeeded(i) = 1
        FatesHits(i) = 0
        CBHits(i) = 0
        BonusTargets(i) = 0
        BonusRamps(i) = 0
        BonusOrbits(i) = 0
        BonusXHits(i) = 0
        ComboHits(i) = 0
        TotalMonsters(i) = 0
        TotalGods(i) = 0
        NeededScoopHits(i) = 1
        ScoopHits(i) = 0
    bTrunksCompleted(i) = False
    bVegetaCompleted(i) = False
    bPiccoloCompleted(i) = False
    bGokuCompleted(i) = False
    bCellCompleted(i) = False
    bZamasuCompleted(i) = False
    bGokuBlackCompleted(i) = False
    bFriezaCompleted(i) = False
    Next
    For i = 0 to 4
        For j = 0 to 9
            Mode(i, j) = 0
        Next
    Next
    TurnOffPlayfieldLights()
End Sub

Sub InstantInfo
    Dim tmp
    DMD CL("INSTANT INFO"), "", "", eNone, eNone, eNone, 1000, True, ""
    Select Case Mode(CurrentPlayer, 0)
        Case 0 ' no Battle active
        Case 1 ' CELL/Minotaur
            DMD CL("CURRENT MODE"), CL("CELL"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("SHOOT THE LIGHTS"), CL("AND SCOOP TO FINISH"), "", eNone, eNone, eNone, 3000, False, ""
        Case 2 ' ZAMASU/Hydra
            DMD CL("CURRENT MODE"), CL("ZAMASU"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("SHOOT THE LIGHTS"), CL("AND SCOOP TO FINISH"), "", eNone, eNone, eNone, 3000, False, ""
        Case 3 ' GOKU BLACK/Cerberus
            DMD CL("CURRENT MODE"), CL("GOKU BLACK"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("SHOOT THE LIGHTS"), CL("AND SCOOP TO FINISH"), "", eNone, eNone, eNone, 3000, False, ""
        Case 4 ' GOLDEN FRIEZA/Medusa
            DMD CL("CURRENT MODE"), CL("GOLDEN FRIEZA"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("SHOOT THE LIGHTS"), CL("AND SCOOP TO FINISH"), "", eNone, eNone, eNone, 3000, False, ""
        Case 5 ' TRUNKS/Ares
            DMD CL("CURRENT MODE"), CL("TRUNKS"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("SHOOT THE SPINNERS"), CL("AND SCOOP TO FINISH"), "", eNone, eNone, eNone, 3000, False, ""
        Case 6 ' VEGETA/Poseidon
            DMD CL("CURRENT MODE"), CL("VEGETA"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("SHOOT LIT TARGETS"), CL("AND SCOOP TO FINISH"), "", eNone, eNone, eNone, 3000, False, ""
        Case 7 ' PICCOLO/Hades
            DMD CL("CURRENT MODE"), CL("PICCOLO"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("SHOOT THE WORLD"), CL("TOURNAMENT ENTRANCE"), "", eNone, eNone, eNone, 3000, False, ""
        Case 8 ' GOKU/Zeus
            DMD CL("CURRENT MODE"), CL("GOKU"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("HIT ALL UPPER GOKU"), CL("TARGETS THEN SCOOP"), "", eNone, eNone, eNone, 3000, False, ""
        Case 9 ' God or Demi-God mode
            DMD CL("CURRENT MODE"), CL("KAI WIZARD MODE"), "", eNone, eNone, eNone, 2000, False, ""
            DMD CL("TIMED BONUS ROUND"), CL("COLLECT JACKPOTS"), "", eNone, eNone, eNone, 3000, False, ""
    End Select

    DMD CL("YOUR SCORE"), CL(FormatScore(Score(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("EXTRA BALLS"), CL(ExtraBallsAwards(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("BONUS MULTIPLIER"), CL(FormatScore(BonusMultiplier(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("PLAYFIELD MULTIPLIER"), CL(FormatScore(PlayfieldMultiplier(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("SKILLSHOT VALUE"), CL(FormatScore(SkillshotValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("SUPR SKILLSHOT VALUE"), CL(FormatScore(SuperSkillshotValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("TARGETS HIT"), CL(FormatScore(BonusTargets(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("RAMPS HIT"), CL(FormatScore(BonusRamps(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("ORBITS HIT"), CL(FormatScore(BonusOrbits(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("COMBO VALUE"), CL(FormatScore(ComboValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("X HITS"), CL(FormatScore(BonusXHits(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("ENEMIES DEFEATED"), CL(TotalMonsters(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("TRAININGS COMPLETED"), CL(TotalGods(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    If Score(1) Then
        DMD CL("PLAYER 1 SCORE"), CL(FormatScore(Score(1) ) ), "", eNone, eNone, eNone, 2000, False, ""
    End If
    If Score(2) Then
        DMD CL("PLAYER 2 SCORE"), CL(FormatScore(Score(2) ) ), "", eNone, eNone, eNone, 2000, False, ""
    End If
    If Score(3) Then
        DMD CL("PLAYER 3 SCORE"), CL(FormatScore(Score(3) ) ), "", eNone, eNone, eNone, 2000, False, ""
    End If
    If Score(4) Then
        DMD CL("PLAYER 4 SCORE"), CL(FormatScore(Score(4) ) ), "", eNone, eNone, eNone, 2000, False, ""
    End If
End Sub

 Sub StopMBmodes 'stop multiball modes after loosing the last multibal
    bCCorpMBStarted = False
  Goku_PoweredUp.visible = 0         'Turns off GokuPowerUP overlays/ sounds/ lights
  StopSound "vo_AuraPoweredUp"       '
  LightSeqGokuPU.StopPlay            '
  bMedusaMBStarted = False
    Goku_Blue.visible = 0              'Turns off GokuBlue overlays/ sounds/ lights
  StopSound "vo_auraBlueWithVocals"    '
  LightSeqGokuBlue.StopPlay          '
    MedussaX = 1
    If bMinotaurMBStarted Then
        bMinotaurMBStarted = False
        DragonRadar.image = "Radar_" & NeededScoopHits(CurrentPlayer) + 3
    Goku_SS2.visible = 0           'Turns off GokuSS overlays/ sounds/ lights
    StopSound "vo_AuraSS2"         '
    LightSeqGokuSS2.StopPlay       '
        If Mode(CurrentPlayer, 0) <> 7 Then 'Not in the Hades mode
            TrapdoorDown
        End If
    End If
  Goku_UltraInstinct20.visible = 0
  StopSound "vo_GokuScreamPowerUp" '        '
    bZeusMBStarted = False
  li042.state = 0
  li044.state = 0
    ZeusMBFlashTimer.Enabled = 0
  If Mode(CurrentPlayer, 0) = 9 OR bKaiWizardModeRan = True Then
    'StopMode 'stop the God or Demi God multiball   'now stays on until timer or end of ball
    'LightSeqGoku.UpdateInterval = 100             'Keep lights on goku for the rest of the game
    'LightSeqGoku.Play SeqRandom, 50, , 10000000
  Else
    LightSeqGoku.StopPlay
  End If
End Sub

Sub StopEndOfBallMode()      'this sub is called after the last ball in play is drained, modes, timers
    StopMode                 'stop current mode
    TrapdoorDown
    LightSeqBumpers.StopPlay 'in case it was on.
End Sub

Sub ResetNewBallVariables()  'reset variables and lights for a new ball or player
    'turn on or off the needed lights before a new ball is released
    TurnOffPlayfieldLights 'this also turn off extra ball and special lights
    BumperLights 0
    'set up the lights according to the player achievments
    UpdateModeLights
    BonusMultiplier(CurrentPlayer) = 1
    PlayfieldMultiplier(CurrentPlayer) = 1
    'Light007.State = 1
    bJackpotsEnabled = False
    If BallsInLock(CurrentPlayer) Then
        li039.State = 2
        bLockEnabled = True
    End If
    BumperHits = 0       'prepare for skillshot
    BumperAward = 5000
    gatekf.RotateToStart 'close the kickback
    KickbackEnabled = False
    kickbackHitsLeft(CurrentPlayer) = 0
    kickbackHitsRight(CurrentPlayer) = 0
    PandorasHitsLeft(CurrentPlayer) = 0
    PandorasHitsRight(CurrentPlayer) = 0
  li015.state = 0
  li016.state = 0
  li017.state = 0
  li018.state = 0
    leftoutlane.Enabled = 1
    If bFatesStarted Then
        bFatesStarted = False
        FatesHits(CurrentPlayer) = 0
        FatesMultiplier = 1
    End If
End Sub

Sub TurnOffPlayfieldLights()
  Dim a
  For each a in aLights
        If TypeName(a) = "Light" Then a.State = 0: End If
        If TypeName(a) = "Flasher" then a.Visible = 0: End If
  Next
End Sub

Sub BumperLights(stat)
    Dim x
    For each x in aBumperLights
        x.State = stat
    Next
End Sub

Sub TurnOffXlights 'turn off the other lights after selecting a X light
    If li024.State = 2 Then li024.State = 0
    If li025.State = 2 Then li025.State = 0
    If li026.State = 2 Then li026.State = 0
    If li027.State = 2 Then li027.State = 0
    If li028.State = 2 Then li028.State = 0
End Sub

Sub UpdateSkillShot() 'Setup and updates the skillshot lights
    LightSeqSkillshot.Play SeqAllOff
    DMD CL("HIT LIT LIGHT"), CL("FOR SKILLSHOT"), "", eNone, eNone, eNone, 3000, True, ""
    'li044.State = 2
    li040.State = 2
    BumperLights 2            'blinking
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    bRotateLights = True
    LightSeqSkillshot.StopPlay
    'li044.State = 0
    li040.State = 0
    BumperLights 1 'on
    DMDScoreNow
End Sub

Sub CheckSkillshot
    If bSkillShotReady Then
        If BumperHits >= 3 Then
            AwardSkillshot
        End If
    End If
End Sub

Sub CheckSuperSkillshot
    If bSkillShotReady Then
        If LastSwitchHit = "hurrican" Then
            AwardSuperSkillshot
        End If
    End If
End Sub

Sub CheckFreakySkillshot
    If bSkillShotReady AND LastSwitchHit = "upperloop3" Then
        If BumperHits = 0 Then
            AwardFreakySkillshot
        End If
    End If
End Sub

'********************
' Flasher light seq.
'********************

Sub RBF 'right bottom Flasher
    LightSeqRBF.Play SeqBlinking, , 8, 40
    DOF 301, DOFPulse
End Sub

Sub RMF
    LightSeqRMF.Play SeqBlinking, , 8, 40
    DOF 304, DOFPulse
End Sub

Sub RTF
    LightSeqRTF.Play SeqBlinking, , 8, 40
    DOF 307, DOFPulse
End Sub

Sub LBF 'left bottom Flasher
    LightSeqLBF.Play SeqBlinking, , 8, 40
    DOF 310, DOFPulse
End Sub

Sub LMF
    LightSeqLMF.Play SeqBlinking, , 8, 40
    DOF 313, DOFPulse
End Sub

Sub LTF
    LightSeqLTF.Play SeqBlinking, , 8, 40
    DOF 316, DOFPulse
End Sub

Sub HelmetF
    LightSeqHelmet.Play SeqBlinking, , 8, 40
    DOF 319, DOFPulse
End Sub

Sub ZeusF
    LightSeqZeusF.Play SeqRandom, 1, , 1500
    DOF 322, DOFPulse
End Sub

Sub FlashEffect(n)
    Select Case n
        Case 1 'all blink
            LightSeqRBF.Play SeqBlinking, , 8, 40:DOF 301, DOFPulse
            LightSeqRMF.Play SeqBlinking, , 8, 40:DOF 304, DOFPulse
            LightSeqRTF.Play SeqBlinking, , 8, 40:DOF 307, DOFPulse
            LightSeqLBF.Play SeqBlinking, , 8, 40:DOF 310, DOFPulse
            LightSeqLMF.Play SeqBlinking, , 8, 40:DOF 313, DOFPulse
            LightSeqLTF.Play SeqBlinking, , 8, 40:DOF 316, DOFPulse
            LightSeqHelmet.Play SeqBlinking, , 8, 40:DOF 319, DOFPulse
        Case 2 'random
            vpmtimer.addtimer RndNbr(6) * 200, "LightSeqRBF.Play SeqBlinking, , 5, 40: DOF 302, DOFPulse '"
            vpmtimer.addtimer RndNbr(6) * 200, "LightSeqRMF.Play SeqBlinking, , 5, 40: DOF 305, DOFPulse '"
            vpmtimer.addtimer RndNbr(6) * 200, "LightSeqRTF.Play SeqBlinking, , 5, 40: DOF 308, DOFPulse '"
            vpmtimer.addtimer RndNbr(6) * 200, "LightSeqLBF.Play SeqBlinking, , 5, 40: DOF 311, DOFPulse '"
            vpmtimer.addtimer RndNbr(6) * 200, "LightSeqLMF.Play SeqBlinking, , 5, 40: DOF 314, DOFPulse '"
            vpmtimer.addtimer RndNbr(6) * 200, "LightSeqLTF.Play SeqBlinking, , 5, 40: DOF 317, DOFPulse '"
            vpmtimer.addtimer RndNbr(6) * 200, "LightSeqHelmet.Play SeqBlinking, , 5, 40: DOF 320, DOFPulse '"
        Case 3 'all blink fast
            LightSeqRBF.Play SeqBlinking, , 4, 30:DOF 302, DOFPulse
            LightSeqRMF.Play SeqBlinking, , 4, 30:DOF 305, DOFPulse
            LightSeqRTF.Play SeqBlinking, , 4, 30:DOF 308, DOFPulse
            LightSeqLBF.Play SeqBlinking, , 4, 30:DOF 311, DOFPulse
            LightSeqLMF.Play SeqBlinking, , 4, 30:DOF 314, DOFPulse
            LightSeqLTF.Play SeqBlinking, , 4, 30:DOF 317, DOFPulse
            LightSeqHelmet.Play SeqBlinking, , 4, 30:DOF 320, DOFPulse
        Case 4 'center
            vpmtimer.addtimer 800, "LightSeqRBF.Play SeqBlinking, , 6, 30: DOF 302, DOFPulse '"
            vpmtimer.addtimer 400, "LightSeqRMF.Play SeqBlinking, , 6, 30: DOF 305, DOFPulse '"
            vpmtimer.addtimer 800, "LightSeqRTF.Play SeqBlinking, , 6, 30: DOF 308, DOFPulse '"
            vpmtimer.addtimer 800, "LightSeqLBF.Play SeqBlinking, , 6, 30: DOF 311, DOFPulse '"
            vpmtimer.addtimer 400, "LightSeqLMF.Play SeqBlinking, , 6, 30: DOF 314, DOFPulse '"
            vpmtimer.addtimer 800, "LightSeqLTF.Play SeqBlinking, , 6, 30: DOF 317, DOFPulse '"
            LightSeqHelmet.Play SeqBlinking, , 6, 30:DOF 320, DOFPulse
        Case 5 'top down
            vpmtimer.addtimer 200, "LightSeqRBF.Play SeqBlinking, , 2, 40: DOF 303, DOFPulse '"
            vpmtimer.addtimer 100, "LightSeqRMF.Play SeqBlinking, , 2, 40: DOF 306, DOFPulse '"
            LightSeqRTF.Play SeqBlinking, , 2, 40:DOF 309, DOFPulse
            vpmtimer.addtimer 200, "LightSeqLBF.Play SeqBlinking, , 2, 40: DOF 312, DOFPulse '"
            vpmtimer.addtimer 100, "LightSeqLMF.Play SeqBlinking, , 2, 40: DOF 315, DOFPulse '"
            LightSeqLTF.Play SeqBlinking, , 2, 40:DOF 318, DOFPulse
            vpmtimer.addtimer 50, "LightSeqHelmet.Play SeqBlinking, , 2, 40: DOF 321, DOFPulse '"
        Case 6 'down to top
            LightSeqRBF.Play SeqBlinking, , 2, 40:DOF 303, DOFPulse
            vpmtimer.addtimer 100, "LightSeqRMF.Play SeqBlinking, , 2, 40: DOF 306, DOFPulse '"
            vpmtimer.addtimer 200, "LightSeqRTF.Play SeqBlinking, , 2, 40: DOF 309, DOFPulse '"
            LightSeqLBF.Play SeqBlinking, , 2, 40:DOF 312, DOFPulse
            vpmtimer.addtimer 100, "LightSeqLMF.Play SeqBlinking, ,2, 40: DOF 315, DOFPulse '"
            vpmtimer.addtimer 200, "LightSeqLTF.Play SeqBlinking, , 2, 40: DOF 318, DOFPulse '"
            vpmtimer.addtimer 150, "LightSeqHelmet.Play SeqBlinking, , 2, 40: DOF 321, DOFPulse '"
        Case 7 'circle 2 rounds
            vpmtimer.addtimer 250, "LightSeqRBF.Play SeqBlinking, , 1, 40: DOF 303, DOFPulse '"
            vpmtimer.addtimer 200, "LightSeqRMF.Play SeqBlinking, , 1, 40: DOF 306, DOFPulse '"
            vpmtimer.addtimer 150, "LightSeqRTF.Play SeqBlinking, , 1, 40: DOF 309, DOFPulse '"
            LightSeqLBF.Play SeqBlinking, , 1, 40:DOF 312, DOFPulse
            vpmtimer.addtimer 50, "LightSeqLMF.Play SeqBlinking, , 1, 40: DOF 315, DOFPulse '"
            vpmtimer.addtimer 100, "LightSeqLTF.Play SeqBlinking, , 1, 40: DOF 318, DOFPulse '"
            vpmtimer.addtimer 550, "LightSeqRBF.Play SeqBlinking, , 1, 40: DOF 303, DOFPulse '"
            vpmtimer.addtimer 500, "LightSeqRMF.Play SeqBlinking, , 1, 40: DOF 306, DOFPulse '"
            vpmtimer.addtimer 450, "LightSeqRTF.Play SeqBlinking, , 1, 40: DOF 309, DOFPulse '"
            vpmtimer.addtimer 300, "LightSeqLBF.Play SeqBlinking, , 1, 40: DOF 312, DOFPulse '"
            vpmtimer.addtimer 350, "LightSeqLMF.Play SeqBlinking, , 1, 40: DOF 315, DOFPulse '"
            vpmtimer.addtimer 400, "LightSeqLTF.Play SeqBlinking, , 1, 40: DOF 318, DOFPulse '"
    End Select
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

'*********************************************************
' Slingshots has been hit
' In this table the slingshots change the outlanes lights

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFcontactors), Lemk
    DOF 144, DOFPulse
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 530
    ' check modes
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_slingshot", 104, DOFPulse, DOFcontactors), Remk
    DOF 145, DOFPulse
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 530
    ' check modes
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub SlingTimer_Timer
    Select case SlingCount
        Case 0, 2, 4, 6, 8:Controller.B2SSetData 10, 1
        Case 1, 3, 5, 7, 9:Controller.B2SSetData 10, 0
        Case 10:SlingTimer.Enabled = 0
    End Select
    SlingCount = SlingCount + 1
End Sub

'***********************
'        Bumper
'***********************

Sub Bumper1_Hit
    If Tilted Then Exit Sub
    Dim tmp
    PlaySoundAt SoundFXDOF("fx_bumper", 105, DOFPulse, DOFContactors), Bumper1
    DOF 147, DOFPulse
    BumperHits = BumperHits + 1
    AddScore BumperAward
    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper1"
    'checkmodes this switch is part of
    CheckSkillshot
    CheckBumperHits
End Sub

Sub Bumper2_Hit
    If Tilted Then Exit Sub
    Dim tmp
    PlaySoundAt SoundFXDOF("fx_bumper", 106, DOFPulse, DOFContactors), Bumper2
    DOF 147, DOFPulse
    BumperHits = BumperHits + 1
    AddScore BumperAward
    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper2"
    'checkmodes this switch is part of
    CheckSkillshot
    CheckBumperHits
End Sub

Sub Bumper3_Hit
    If Tilted Then Exit Sub
    Dim tmp
    PlaySoundAt SoundFXDOF("fx_bumper", 106, DOFPulse, DOFContactors), Bumper3
    DOF 147, DOFPulse
    BumperHits = BumperHits + 1
    AddScore BumperAward
    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper3"
    'checkmodes this switch is part of
    CheckSkillshot
    CheckBumperHits
End Sub

Sub Bumper4_Hit
    If Tilted Then Exit Sub
    Dim tmp
    PlaySoundAt SoundFXDOF("fx_bumper", 106, DOFPulse, DOFContactors), Bumper4
    DOF 147, DOFPulse
    BumperHits = BumperHits + 1
    AddScore BumperAward
    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper4"
    'checkmodes this switch is part of
    CheckSkillshot
    CheckBumperHits
End Sub

Sub CheckBumperHits
    If BumperHits MOD 20 = 0 Then 'add a ball if in multiball Mode
        If bMultiBallMode Then
            DMD "_", CL("ADD A BALL"), "_", eNone, eNone, eNone, 1500, True, ""
            AddMultiball 1
        End If
    End If
    If BumperHits MOD 25 = 0 Then 'activate chain lightning, bumpers score 4x
        DMD CL("GREAT APE"), CL("TRANSFORMATION"), "_", eNone, eNone, eNone, 5000, True, "vo_chain_lightningApe"
    DMD CL("SUPER BUMPERS"), CL("X4 VALUE OF " &FormatScore(BumperAward) ), "_", eNone, eNone, eNone, 5000, True, "vo_chain_lightning"
    PlaySound "sfx_roar" &RndNbr(5)
        FlashEffect RndNbr(7)
        LightSeqBumpers.Play SeqRandom, 10, , 1000
    BumperAward = BumperAward * 4
        pf_GrateApeTransformation.visible = 1 '4x overlay
    LightMOON.STATE = 2
    LightMOON.BlinkInterval = 1000        'slow blink MOON
    Table1.ColorGradeImage = "LUT14"
        li035.State = 2 'ape light
        Magnet.Enabled = True:DOF 132, DOFOn
    End If
End Sub

Sub Magnet_Hit
    vpmTimer.AddTimer 1500, "MagnetKickBall '"
End Sub

'Kicks balls in a random direction
Sub MagnetKickBall
    Magnet.Kick 260 + RndNbr(60), 20
  DMD CL("GREAT APE"), CL("HIT"), "_", eNone, eNone, eNone, 1500, True, "sfx_roar" &RndNbr(5)
End Sub

Sub LightSeqBumpers_PlayDone()
    LightSeqBumpers.Play SeqRandom, 10, , 1000
End Sub
'*********
' Lanes
'*********
' in and outlanes
Sub leftoutlane_Hit
    PlaySoundAt "fx_sensor", leftoutlane
    If Tilted Then Exit Sub
    'score & bonus
    AddScore 250000
    If activeball.VelY> 0 Then 'only when going down
        If li002.State Then    'if the skull is lit then reduce the ball saver time
            'PlaySound "vo_evilmuah"
            BallSaverTime = BallSaverTime -2
        End If
    End If
    LastSwitchHit = "leftoutlane"
End Sub

Sub leftinlane_Hit
    PlaySoundAt "fx_sensor", leftinlane
    If Tilted Then Exit Sub
    'score & bonus
    AddScore 100000
    If li003.State = 0 Then 'count the trees that are being lit up
        li003.State = 1
        TreeHits(CurrentPlayer) = TreeHits(CurrentPlayer) + 1
        CheckTrees
    End If
'LastSwitchHit = "leftinlane"
End Sub

Sub rightinlane_Hit
    PlaySoundAt "fx_sensor", rightinlane
    If Tilted Then Exit Sub
    'score & bonus
    AddScore 100000
    If li004.State = 0 Then 'count the trees that are being lit up
        li004.State = 1
        TreeHits(CurrentPlayer) = TreeHits(CurrentPlayer) + 1
        CheckTrees
    End If
'LastSwitchHit = "rightinlane"
End Sub

Sub rightoutlane_Hit
    PlaySoundAt "fx_sensor", rightoutlane
    If Tilted Then Exit Sub
    'score & bonus
    AddScore 250000
    If li005.State Then 'if the skull is lit then reduce the ball saver time
        PlaySound "vo_outlane"
        BallSaverTime = BallSaverTime -2
        DMD CL("BALLSAVE DECREASED"), CL("IT IS NOW " &BallSaverTime& " SEC"), "", eNone, eNone, eNone, 2000, True, ""
    End If
    LastSwitchHit = "rightoutlane"
End Sub

Sub CheckTrees                                'check for 6 treehits and if all the lights are lit
    If TreeHits(CurrentPlayer) MOD 6 = 0 Then 'every 6 trees adds 3 seconds to the ball saver value
        If BallSaverTime <50 then
            BallSaverTime = BallSaverTime + 3
            DMD CL("BALLSAVE INCREASED"), CL("IT IS NOW " &BallSaverTime& " SEC"), "", eNone, eNone, eNone, 2000, True, ""
        End If
    End If
    If li002.State + li003.State + li004.State + li005.State = 4 Then 'all lights are lit then turn them off
        li002.State = 0
        li003.State = 0
        li004.State = 0
        li005.State = 0
        LightSeqLanes.Play SeqRandom, 4, , 1000
    End If
End Sub

' loops

Sub hurrican_Hit
    PlaySoundAt "fx_sensor", hurrican
    If Tilted Then Exit Sub
    'score & bonus
    AddScore 100000 * MedussaX
    If li024.State = 1 Then
        Addscore 100000 * MedussaX 'double the score
    End If
    If li024.State = 2 Then
        li024.State = 1
        TurnOffXlights
    End If
    If li030.State = 0 Then
        li030.State = 1
        CheckBonusX
    End If
    'Modes
    If activeBall.VelY <0 Then                               'ball going up
        Select Case Mode(CurrentPlayer, 0)
            Case 0:OrbitHits = OrbitHits + 1:CheckStartModes 'added 4 lines
            Case 1, 4
                if li060.State = 2 Then
                    li060.state = 0
                    CheckWinMode
                End If
        End Select
        'Combos
        If LastSwitchHit = "upperloop2" OR LastSwitchHit = "upperloop4" Then
            AwardCombo
        Else
            ComboCount = 1
        End If
    End If
    LastSwitchHit = "hurrican"
End Sub

Sub upperloop1_Hit
    PlaySoundAt "fx_sensor", upperloop1
    If Tilted Then Exit Sub
    'score & bonus
    AddScore 100000 * MedussaX
    'Modes
    If activeBall.VelY <0 Then                               'ball going up
        Select Case Mode(CurrentPlayer, 0)
            Case 0:OrbitHits = OrbitHits + 1:CheckStartModes 'added 4 lines
            Case 1, 4
                if li037.State = 2 Then
                    li037.state = 0
                    CheckWinMode
                End If
        End Select
        'Combos
        If LastSwitchHit = "upperloop4" OR LastSwitchHit = "hurrican" Then
            AwardCombo
        Else
            ComboCount = 1
        End If
    End If
    LastSwitchHit = "upperloop1"
End Sub

Sub upperloop2_Hit
    PlaySoundAt "fx_sensor", upperloop2
    If Tilted Then Exit Sub
    'score & bonus
    AddScore 100000 * MedussaX
    ' check modes

    'Select Case Mode(CurrentPlayer, 0)                   '3 lines commented out
    '   Case 0:OrbitHits = OrbitHits + 1:CheckStartModes
    'End Select

    'Medusa MB
    If bMedusaMBStarted Then
        ExtraBallHits = ExtraBallHits + 1
        CheckExtraBallHits
    End If
    LastSwitchHit = "upperloop2"
End Sub

Sub upperloop3_Hit
    PlaySoundAt "fx_sensor", upperloop3
    If Tilted Then Exit Sub
    'score & bonus
    AddScore 100000 * MedussaX
    LastSwitchHit = "upperloop3"
    CheckFreakySkillshot
End Sub

Sub upperloop4_Hit
    PlaySoundAt "fx_sensor", upperloop4
    If Tilted Then Exit Sub
    'score & bonus
    AddScore 100000 * MedussaX
    'Modes
    If li028.State = 1 Then
        Addscore 100000 'double the score
    End If
    If li028.State = 2 Then
        li028.State = 1
        TurnOffXlights
    End If
    If li034.State = 0 Then
        li034.State = 1
        CheckBonusX
    End If
    If activeBall.VelY <0 Then                               'ball going up
        Select Case Mode(CurrentPlayer, 0)
            Case 0:OrbitHits = OrbitHits + 1:CheckStartModes 'added 4 lines
            Case 1, 4
                if li041.State = 2 Then
                    li041.state = 0
                    CheckWinMode
                End If
        End Select
        'Combos
        If LastSwitchHit = "upperloop2" OR LastSwitchHit = "hurrican" Then
            AwardCombo
        Else
            ComboCount = 1
        End If
    End If
    LastSwitchHit = "upperloop4"
End Sub

'ramps completed
Sub lramp_Hit
    PlaySoundAt "fx_sensor", lramp
    If Tilted Then Exit Sub
    'score & bonus
    If Light009.State Then
        AwardExtraBall
    End If
    If bMedusaMBStarted Then
        AddScore 50000
    Else
        AddScore 100000
    End If
    If bJackpotsEnabled Then AwardJackpot
    'Fates
    If bFatesStarted Then 'change the scoring multiplier
    light006.State = 0
        Light007.State = 0
        Light008.State = 0
    FatesCounter = FatesCounter + 1
        Select Case FatesCounter
            Case 1:FatesMultiplier = 1.5:light006.State = 1
            Case 2:FatesMultiplier = 2:light007.State = 1
            Case else:FatesMultiplier = 3:light008.State = 1
        End Select
    Else
        FatesHits(CurrentPlayer) = FatesHits(CurrentPlayer) + 1
        CheckFates
    End If
    'fire & X lights
    If li025.State = 1 Then
        AddScore 100000
    End If
    If li025.State = 2 Then
        li025.State = 1
        TurnOffXlights
    End If
    If li031.State = 0 Then
        li031.State = 1
        CheckBonusX
    End If
    ' check modes
    Select Case Mode(CurrentPlayer, 0)
        Case 0
            RampHits = RampHits + 1:CheckStartModes
        Case 2
            If ModeStep = 1 OR ModeStep = 3 Then
                CheckWinMode
            End If
        Case 3:li036.State = 0:CheckWinMode
        Case 4
            If ModeStep = 2 Then
                CheckWinMode
            End If
        Case 9
            AwardJackpot
    End Select
    'Combos
    If LastSwitchHit = "tramp" OR LastSwitchHit = "rramp" Then
        AwardCombo
    Else
        ComboCount = 1
    End If
    LastSwitchHit = "lramp"
End Sub

Sub rramp_Hit
    PlaySoundAt "fx_sensor", rramp
    If Tilted Then Exit Sub
    'score & bonus
    CheckSuperSkillshot
    If Light010.State Then
        AwardSpecial
    End If
    If bMedusaMBStarted Then
        AddScore 50000
    Else
        AddScore 100000
    End If
    If bJackpotsEnabled Then AwardJackpot
    If li027.State = 1 Then
        AddScore 100000 'double the score
    End If
    If li027.State = 2 Then
        li027.State = 1
        TurnOffXlights
    End If
    If li033.State = 0 Then
        li033.State = 1
        CheckBonusX
    End If
    ' check modes
    Select Case Mode(CurrentPlayer, 0)
        Case 0
            RampHits = RampHits + 1:CheckStartModes
        Case 2
            If ModeStep = 2 OR ModeStep = 4 Then
                CheckWinMode
            End If
        Case 3:li040.State = 0:CheckWinMode
        Case 9
            AwardJackpot
    End Select
    'Combos
    If LastSwitchHit = "tramp" OR LastSwitchHit = "lramp" Then
        AwardCombo
    Else
        ComboCount = 1
    End If
    LastSwitchHit = "rramp"
End Sub

Sub tramp_Hit
    PlaySoundAt "fx_sensor", tramp
    If Tilted Then Exit Sub
    diverter1.Isdropped = 0
    LTF:RTF
'score & bonus
If bMedusaMBStarted Then
    AddScore 250000
Else
    AddScore 500000
End If
    If bJackpotsEnabled Then AwardJackpot
    If bZeusMBStarted Then AwardSuperJackpot
    ' check modes
    Select Case Mode(CurrentPlayer, 0)
        Case 0
            RampHits = RampHits + 1:CheckStartModes
        Case 3:li044.State = 0:CheckWinMode
        Case 9
            AwardJackpot
    End Select
    LastSwitchHit = "tramp"
End Sub

Sub tramp2_Hit
    PlaySoundAt "fx_sensor", tramp2
    If Tilted Then Exit Sub
    diverter2.Isdropped = 0
    LTF:RTF
'score & bonus
If bMedusaMBStarted Then
    AddScore 250000
Else
    AddScore 500000
End If
    If bJackpotsEnabled Then AwardJackpot
    If bZeusMBStarted Then AwardSuperJackpot
    ' check modes
    Select Case Mode(CurrentPlayer, 0)
        Case 0
            RampHits = RampHits + 1:CheckStartModes
        Case 3:li042.State = 0:CheckWinMode
        Case 9
            AwardJackpot
    End Select
    LastSwitchHit = "tramp2"
End Sub

'Effect triggers

Sub Trigger001_Hit:LMF:End Sub
Sub Trigger002_Hit:LBF:End Sub
Sub Trigger003_Hit:RMF:End Sub
Sub Trigger004_Hit:RBF:End Sub

'***********
' Targets
'***********

' kickback targets
Sub leftkb_Hit 'left lower kickback target
    PlaySoundAtBall SoundFXDOF("fx_Target", 124, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 10000

    Select Case Mode(CurrentPlayer, 0)
        Case 6:
            If li015.State Then
                PlayLightning
                GiEffect 3
                li015.State = 0
                AddScore 15000
                CheckWinMode
            End If
        Case Else
            kickbackHitsLeft(CurrentPlayer) = kickbackHitsLeft(CurrentPlayer) + 1
            If kickbackHitsLeft(CurrentPlayer) >= kickbackNeeded(CurrentPlayer) Then
                FlashForMs li015, 1000, 50, 1
                If kickbackHitsRight(CurrentPlayer) >= kickbackNeeded(CurrentPlayer) Then
                    Checkkickback
                End If
            Else
                FlashForMs li015, 1000, 50, 0
            End If
    End Select
    LastSwitchHit = "leftkb"
End Sub

Sub rightkb_Hit 'left upper kickback target
    PlaySoundAtBall SoundFXDOF("fx_Target", 125, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 10000

    Select Case Mode(CurrentPlayer, 0)
        Case 6:
            If li016.State Then
                PlayLightning
                GiEffect 3
                li016.State = 0
                AddScore 15000
                CheckWinMode
            End If
        Case Else
            kickbackHitsRight(CurrentPlayer) = kickbackHitsRight(CurrentPlayer) + 1
            If kickbackHitsRight(CurrentPlayer) >= kickbackNeeded(CurrentPlayer) Then
                FlashForMs li016, 1000, 50, 1
                If kickbackHitsLeft(CurrentPlayer) >= kickbackNeeded(CurrentPlayer) Then
                    Checkkickback
                End If
            Else
                FlashForMs li016, 1000, 50, 0
            End If
    End Select
    LastSwitchHit = "rightkb"
End Sub

'Pandora's targets

Sub lmyst_Hit 'right upper pandora target
    PlaySoundAtBall SoundFXDOF("fx_Target", 126, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 10000
    Select Case Mode(CurrentPlayer, 0)
        Case 6:
            If li018.State Then
                PlayLightning
                GiEffect 3
                li018.State = 0
                AddScore 15000
                CheckWinMode
            End If
        Case Else
            PandorasHitsRight(CurrentPlayer) = PandorasHitsRight(CurrentPlayer) + 1
            if PandorasHitsRight(CurrentPlayer) >= PandorasNeeded(CurrentPlayer) Then
                FlashForMs li018, 1000, 50, 1
                If PandorasHitsLeft(CurrentPlayer) >= PandorasNeeded(CurrentPlayer) Then
                    CheckPandora
                End If
            Else
                FlashForMs li018, 1000, 50, 0
            End If
    End Select
    LastSwitchHit = "lmyst"
End Sub

Sub rmyst_Hit 'right lower pandora target
    PlaySoundAtBall SoundFXDOF("fx_Target", 127, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 10000

    Select Case Mode(CurrentPlayer, 0)
        Case 6:
            If li017.State Then
                PlayLightning
                GiEffect 3
                li017.State = 0
                AddScore 15000
                CheckWinMode
            End If
        Case Else
            PandorasHitsLeft(CurrentPlayer) = PandorasHitsLeft(CurrentPlayer) + 1
            if PandorasHitsLeft(CurrentPlayer) >= PandorasNeeded(CurrentPlayer) Then
                FlashForMs li017, 1000, 50, 1
                If PandorasHitsRight(CurrentPlayer) >= PandorasNeeded(CurrentPlayer) Then
                    CheckPandora
                End If
            Else
                FlashForMs li017, 1000, 50, 0
            End If
    End Select
    LastSwitchHit = "rmyst"
End Sub

' captive ball
Sub cball_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 119, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 75000
    LightEffect 3
    'Modes
    CBHits(CurrentPlayer) = CBHits(CurrentPlayer) + 1
    If CBHits(CurrentPlayer) MOD 10 = 0 Then 'open trapdoor for jackpot
        vpmTimer.AddTimer 1500, "PLaySound""vo_trapdoor"":TrapdoorUp '"
    End If
    If CBHits(CurrentPlayer) MOD 25 = 0 Then 'ball breaker
        DMD CL("WIN WORLD TOURNAMENT"), CL(FormatScore(750000) ), "_", eNone, eNone, eNone, 3000, True, "vo_ball_breaker"
        Addscore2 750000
        li045.State = 1
    End If
    LastSwitchHit = "cball"
End Sub

' thin targets
Sub liup1_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 142, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 10000
    Select Case Mode(CurrentPlayer, 0)
        Case 6:
            If li019.State Then
                PlayLightning
                GiEffect 3
                li019.State = 0
                AddScore 15000
                CheckWinMode
            End If
        Case Else
            If li019.State = 0 Then
                PlayLightning
                ZeusF
                li019.State = 1
                CheckGodMode
            End If
    End Select
    LastSwitchHit = "liup1"
End Sub

Sub liup2_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 142, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 10000
    Select Case Mode(CurrentPlayer, 0)
        Case 6:
            If li020.State Then
                PlayLightning
                GiEffect 3
                li020.State = 0
                AddScore 15000
                CheckWinMode
            End If
        Case Else
            If li020.State = 0 Then
                PlayLightning
                ZeusF
                li020.State = 1
                CheckGodMode
            End If
    End Select
    LastSwitchHit = "liup2"
End Sub

Sub liup3_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 142, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 10000
    Select Case Mode(CurrentPlayer, 0)
        Case 6:
            If li021.State Then
                PlayLightning
                GiEffect 3
                li021.State = 0
                AddScore 15000
                CheckWinMode
            End If
        Case Else
            If li021.State = 0 Then
                PlayLightning
                ZeusF
                li021.State = 1
                CheckGodMode
            End If
    End Select
    LastSwitchHit = "liup3"
End Sub

Sub liup4_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 142, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 10000
    Select Case Mode(CurrentPlayer, 0)
        Case 6:
            If li022.State Then
                PlayLightning
                GiEffect 3
                li022.State = 0
                AddScore 15000
                CheckWinMode
            End If
        Case Else
            If li022.State = 0 Then
                PlayLightning
                ZeusF
                li022.State = 1
                CheckGodMode
            End If
    End Select
    LastSwitchHit = "liup4"
End Sub

Sub liup5_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 143, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 10000
    Select Case Mode(CurrentPlayer, 0)
        Case 6:
            If li023.State Then
                PlayLightning
                GiEffect 3
                li023.State = 0
                AddScore 15000
                CheckWinMode
            End If
        Case Else
            If li023.State = 0 Then
                PlayLightning
                ZeusF
                li023.State = 1
                CheckGodMode
            End If
    End Select
    LastSwitchHit = "liup5"
End Sub

Sub TargetLightsAll(stat) 'same as light.state
    For each x in aTargetsAll
        x.State = stat
    Next
End Sub

Sub TargetLightsThunder(stat) 'same as light.state
    For each x in aTargetsThunder
        x.State = stat
    Next
End Sub

' Zeus targets
Sub ztgt_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 133, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 50000
    li061.State = 1
    ZeusF
    CheckStartModes
    LastSwitchHit = "ztgt"
End Sub

Sub etgt_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 134, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 50000
    li062.State = 1
    ZeusF
    CheckStartModes
    LastSwitchHit = "etgt"
End Sub

Sub utgt_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 135, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 50000
    li063.State = 1
    ZeusF
    CheckStartModes
    LastSwitchHit = "utgt"
End Sub

Sub stgt_Hit
    PlaySoundAtBall SoundFXDOF("fx_Target", 136, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 50000
    li064.State = 1
    ZeusF
    CheckStartModes
    LastSwitchHit = "stgt"
End Sub

'*************
'  Spinners
'*************

Sub rspin_Spin 'right
    PlaySoundAt "fx_spinner", rspin
    DOF 129, DOFPulse
    If Tilted Then Exit Sub
    Addscore 10000
    ' check modes
    Select Case Mode(CurrentPlayer, 0)
        Case 5:SpinnerHits = SpinnerHits + 1:CheckWinMode
    End Select
End Sub

Sub lspin_Spin 'left
    PlaySoundAt "fx_spinner", lspin
    DOF 128, DOFPulse
    If Tilted Then Exit Sub
    Addscore 10000
    ' check modes
    Select Case Mode(CurrentPlayer, 0)
        Case 5:SpinnerHits = SpinnerHits + 1:CheckWinMode
    End Select
End Sub

'*********
' scoops
'*********

Sub scoop1_Hit 'DRAGON RADAR/helmet scoop
    Dim tmp
    PlaySoundAt "fx_hole_enter", scoop1
    scoop1.Destroyball
    BallsinHole = BallsInHole + 1
    If Tilted Then vpmtimer.addtimer 500, "kickBallOut '":Exit Sub
    FlashEffect 7
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    ' Modes
    If bMinotaurMBStarted AND Minotaur1 = 0 Then
        AwardMinotaurJackpot
        Minotaur1 = 1
        CheckMinotaurMBHits
    Else 'lock balls when not in multiball
        IF NOT bMinotaurMBStarted Then
            If bLockEnabled = False Then
                ScoopHits(CurrentPlayer) = ScoopHits(CurrentPlayer) + 1
                If ScoopHits(CurrentPlayer) = NeededScoopHits(CurrentPlayer) Then 'enable the lock and light
                    DMD "_", CL("LOCK IS LIT"), "_", eNone, eNone, eNone, 1500, True, "fx_DragonRadar_3"
                    DragonRadar.image = "Radar_3"
                    bLockEnabled = True
                    li039.State = 2
                Else
                    tmp = NeededScoopHits(CurrentPlayer) - ScoopHits(CurrentPlayer)
                    DMD "_", CL("NEED " &tmp& " MORE HITS"), "_", eNone, eNone, eNone, 1500, True, "fx_DragonRadar_4"
                    DragonRadar.image = "Radar_" & (tmp + 3)
                End If
            Else 'add a ball to the lock
                BallsInLock(CurrentPlayer) = BallsInLock(CurrentPlayer) + 1
                CheckMinotaurMB
            End If
        End If
    End If
    Addscore 5000
    ' check modes
    Select Case Mode(CurrentPlayer, 0)
        Case 3, 5, 6, 8
            If ModeStep = 2 Then
                WinMode
            Else
                vpmtimer.addtimer 1500, "kickBallOut '"
            End If
        Case 4
            If ModeStep = 3 Then
                WinMode
            Else
                vpmtimer.addtimer 1500, "kickBallOut '"
            End If
        Case 1
            If ModeStep = 4 Then
                WinMode
            Else
                vpmtimer.addtimer 1500, "kickBallOut '"
            End If
        Case 2
            If ModeStep = 5 Then
                WinMode
            Else
                vpmtimer.addtimer 1500, "kickBallOut '"
            End If
        Case Else
            ' Nothing left to do, so kick out the ball
            vpmtimer.addtimer 1500, "kickBallOut '"
    End Select
End Sub

Sub kickBallOut 'from all the holes
  debug.print BallsinHole
    If BallsinHole> 0 Then
        BallsinHole = BallsInHole - 1
        PlaySoundAt SoundFXDOF("fx_popper", 111, DOFPulse, DOFcontactors), scoopexit
        DOF 130, DOFPulse
        scoopexit.CreateSizedBallWithMass BallSize / 2, BallMass
        UpdateBallImage
        scoopexit.kick 196, 28
        HelmetF
        LightEffect 5
    debug.print "kicking ball out now"
        vpmtimer.addtimer 1500, "kickBallOut '" 'kick out the rest of the balls, if any
    End If
End Sub

' hole2 - Start Battle

Sub scoop2_Hit 'Start Battle
    PlaySoundAt "fx_hole_enter", scoop2
    scoop2.Destroyball
    BallsinHole = BallsInHole + 1
    If Tilted Then vpmtimer.addtimer 500, "kickBallOut '":Exit Sub
    ' Modes
    Addscore 25000
    If li026.State = 1 Then
        AddScore 25000 'double the score
    End If
    If li026.State = 2 Then
        li026.State = 1
        TurnOffXlights
    End If
    If li032.State = 0 Then
        li032.State = 1
        CheckBonusX
    End If
    If bMinotaurMBStarted AND Minotaur2 = 0 Then
        AwardMinotaurJackpot
        Minotaur2 = 1
        CheckMinotaurMBHits
    End If
    If bModeReady Then
        bModeReady = False
        li001.State = 0
        li038.State = 0
        StartNextMode
    Else
        vpmtimer.addtimer 500, "kickBallOut '"
    End If
End Sub

'***********
' WORLD TOURNAMNET/Trapdoor
'***********
'hole3 - HadesTrapdoor

Sub TrapDoorK_Hit 'Hades Battle
    PlaySoundAt "fx_hole_enter", TrapDoorK
    TrapDoorK.Destroyball
    BallsinHole = BallsInHole + 1
    If Tilted Then vpmtimer.addtimer 500, "kickBallOut '":Exit Sub
    ' Modes
    If bMinotaurMBStarted AND Minotaur3 = 0 Then
        AwardMinotaurJackpot
        Minotaur3 = 1
        CheckMinotaurMBHits
    End If
    Select Case Mode(CurrentPlayer, 0)
        Case 7
            TrapDoorHits = TrapDoorHits + 1
            AddScore2 500000
            CheckWinMode
        Case Else
            If NOT bMinotaurMBStarted Then
                'score a Jackpot
                AwardJackpot
                ' and close the trap door
                vpmTimer.AddTimer 1000, "TrapdoorDown '"
            End If
    End Select
    vpmtimer.addtimer 2000, "kickBallOut '"
End Sub

Sub TrapdoorUp
    PlaySoundAt SoundFXDOF("fx_SolenoidOn", 117, DOFPulse, DOFContactors), TrapDoorK
    TrapdoorA.IsDropped = 1
    TrapdoorB.IsDropped = 0
    TrapdoorC.IsDropped = 0
    TrapdoorD.Collidable = 0
    TrapDoorK.Enabled = 1
    Light012.State = 2
End Sub

Sub TrapdoorDown
    PlaySoundAt SoundFXDOF("fx_SolenoidOff", 117, DOFPulse, DOFContactors), TrapDoorK
    TrapdoorA.IsDropped = 0
    TrapdoorB.IsDropped = 1
    TrapdoorC.IsDropped = 1
    TrapdoorD.Collidable = 1
    TrapDoorK.Enabled = 0
    Light012.State = 0
End Sub

'***********
' Kick back
'***********

Sub Checkkickback
    'check if enough hits to enable kickback
    if KickbackEnabled = false Then
        DMD "_", CL("KICKBACK IS LIT"), "", eNone, eNone, eNone, 1500, True, "vo_kickbackislit"
        PlaySoundAt "fx_diverter", gatekf
        gatekf.RotateToEnd
        leftoutlane.Enabled = 0 'disable the outlane kicker so it doesn't remove seconds in case of the skull light being lit.
        LMF:LBF
        KickbackEnabled = true
    End If
End Sub

Sub Kickback_Hit
    PlaySoundAt "fx_kicker_enter", Kickback
    vpmtimer.addtimer 2000, "DOF 110, DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors),Kickback:kickback.kick 0,45 '"
    vpmtimer.addtimer 2500, "PlaySoundAt""fx_diverter"", gatekf:leftoutlane.Enabled =1:gatekf.RotateToStart '"
    kickbackNeeded(CurrentPlayer) = kickbackNeeded(CurrentPlayer) + 1
    kickbackHitsRight(CurrentPlayer) = 0
    kickbackHitsLeft(CurrentPlayer) = 0
    li015.STATE = 0 'turn off kickback lights
    li016.STATE = 0 'turn off kickback lights
    KickbackEnabled = False
End Sub

'******************************
' CAPSULE CORP/Pandora's Box - Extra awards
'******************************
'hole4 - Pandora's box

Sub pandorak_Hit
    PlaySoundAt "fx_hole_enter", pandorak
    pandorak.Destroyball
    BallsinHole = BallsInHole + 1
    If Tilted Then vpmtimer.addtimer 500, "kickBallOut '":Exit Sub
    RMF:RBF
If li029.State Then
    StartPandora
    PandorasHitsRight(CurrentPlayer) = 0
    PandorasHitsLeft(CurrentPlayer) = 0
    li017.STATE = 0 'turn off Pandora lights
    li018.STATE = 0 'turn off Pandora lights
    PandorasHitsLeft(CurrentPlayer) = 0
    PandorasHitsRight(CurrentPlayer) = 0
    If PandorasNeeded(CurrentPlayer) <10 Then
        PandorasNeeded(CurrentPlayer) = PandorasNeeded(CurrentPlayer) + 1
    End If
    li029.State = 0
Else
    vpmtimer.addtimer 1500, "kickBallOut '"
End If
End Sub

Sub CheckPandora
    'check if enough hits to light Pandora's light
    If li029.State = 0 Then
        DMD "_", CL("CAPSULE CORP IS LIT"), "", eNone, eNone, eNone, 1500, True, "vo_mystery"
        li029.State = 1
    End If
End Sub

Sub StartPandora
    'do some animation
    DMDFlush
    DMD CL("CAPSULE CORP."), CL("LIGHT EXTRABALL"), "", eNone, eNone, eNone, 200, False, ""
    DMD "_", CL("1 MILLION"), "", eNone, eNone, eNone, 200, False, ""
    DMD "_", CL("LIT SPECIAL"), "", eNone, eNone, eNone, 200, False, ""
    DMD "_", CL("5 MILLION"), "", eNone, eNone, eNone, 200, False, ""
    DMD "_", CL("START MODE"), "", eNone, eNone, eNone, 200, False, ""
    'give award
    Dim tmp
    Select case RndNbr(16)
        Case 1, 8 'light extraball
            If Light009.State = 0 Then
                DMD CL("CAPSULE CORP."), CL("EXTRA BALL IS LIT"), "_", eNone, eBlink, eNone, 2500, True, "vo_extraballislit"
                Light009.State = 2
            Else
                DMD CL("CAPSULE CORP."), CL(FormatScore(250000) ), "_", eNone, eBlink, eNone, 2000, True, ""
                Addscore 250000
            End If
        Case 2, 9 'light special
            If Light010.State = 0 Then
                DMD CL("CAPSULE CORP."), CL("SPECIAL IS LIT"), "_", eNone, eBlink, eNone, 2500, True, "vo_specialislit"
                Light010.State = 2
            Else
                DMD CL("CAPSULE CORP."), CL(FormatScore(250000) ), "_", eNone, eBlink, eNone, 2000, True, ""
                Addscore2 250000
            End If
        Case 3, 10 'start 3 balls multiball, double playfield scores
      bCCorpMBStarted = True
      If bMedusaMBStarted = True OR bZeusMBStarted = True OR bMinotaurMBStarted = True OR bKaiModeStarted Then  'Higher layer already on
        DMD CL("DOUBLE MULTIBALL"), CL("PLAYFIELD X 2"), "_", eBlink, eNone, eNone, 3000, True, "vo_kaiokenx2"
      Else If CurrentMode(CurrentPlayer) = 8 OR CurrentMode(CurrentPlayer) = 9 Then          'Higher layer already on
        DMD CL("CAPSULE CORP MBALL"), CL("PLAYFIELD X 2"), "_", eNone, eNone, eNone, 2500, True, "vo_multiball"
      Else
        DMD CL("CAPSULE CORP MBALL"), CL("PLAYFIELD X 2"), "_", eNone, eNone, eNone, 2500, True, "vo_multiball"
        Goku_PoweredUp.visible = 1
        PlaySound "vo_AuraPoweredUp"
        LightSeqGokuPU.UpdateInterval = 100
        LightSeqGokuPU.Play SeqRandom, 50, , 10000000
        ChangeSong
        EnableBallSaver (8)
      End If
      End If
      AddPlayfieldMultiplier 1
            AddMultiball 2
        Case 4, 11 'award from 250k to 5 million
            tmp = 250000 * RndNbr(20)
            DMD CL("CAPSULE CORP."), CL(FormatScore(tmp) ), "_", eNone, eBlink, eNone, 2000, True, ""
            Addscore2 tmp
        Case 5, 12 'increment bonus multiplier
            AddBonusMultiplier 1
            DMD CL("CAPSULE CORP."), CL("BONUS X " & BonusMultiplier(CurrentPlayer) ), "_", eNone, eBlinkFast, eNone, 2000, True, ""
        Case 6, 13 'add 10 pop bumper values
            BumperAward = BumperAward * 10
            DMD CL("CAPSULE CORP."), CL("BUMPERS 10X VALUE"), "_", eNone, eNone, eNone, 1000, True, ""
            DMD CL("BUMPERS VALUE"), CL(FormatScore(BumperAward) ), "_", eNone, eNone, eNone, 1500, True, ""
        Case 7, 14 '20 or more seconds ball saver
            DMD CL("CAPSULE CORP."), CL("SENZU BALL SAVE"), "_", eNone, eNone, eNone, 1000, True, ""
            StartGodMode
        Case Else 'DUD!  Only 1000 to 25000 points
            tmp = 1000 * RndNbr(25)
            DMD CL("CAPSULE CORP. DUD"), CL(FormatScore(tmp) ), "_", eNone, eBlink, eNone, 2000, True, "vo_DUD" &RndNbr(11)
            Addscore2 tmp
    End Select
    vpmtimer.addtimer 4500, "kickBallOut '"
End Sub

Sub StartGodMode
    ' SetB2SData 6, 1 'backglass gods flashing
    DMD CL("SENZU BALL SAVE"), CL("FOR " &BallSaverTime& " SECONDS"), "_", eNone, eNone, eNone, 2000, True, "vo_godmode"
    EnableBallSaver BallSaverTime
    PlayThunder
    FlashEffect RndNbr(7)
End Sub

Sub CheckGodMode
    If li019.State + li020.State + li021.State + li022.State + li023.State = 5 Then
        GodModeHits(CurrentPlayer) = GodModeHits(CurrentPlayer) + 1
        If GodModeHits(CurrentPlayer) = GodModeNeeded(CurrentPlayer) Then
            StartGodMode
            GodModeHits(CurrentPlayer) = 0
            GodModeNeeded(CurrentPlayer) = GodModeNeeded(CurrentPlayer) + 1 'increase the number of times needed to start God Mode
            If GodModeNeeded(CurrentPlayer)> 5 Then GodModeNeeded(CurrentPlayer) = 5
        End If
        li019.State = 0
        li020.State = 0
        li021.State = 0
        li022.State = 0
        li023.State = 0
        LightEffect 2
    End If
End Sub

'***********************************
' Modes - ENEMIES/Monsters BATTLES and HEROES/Gods TRAININGS
'***********************************
' only one mode can be played at one time
' This table has 8 modes or battles, with a nr. 9 being the end Wizard mode

Sub CheckStartModes 'check for the requirements to start a new mode, and activate the "start battle" light/scoop
    Dim tmp
    tmp = li061.State + li062.State + li063.State + li064.State
    Select Case Mode(CurrentPlayer, 0)
        Case 0 'no battle is active
            'are all GOKU/zeus targets lit? Or there is enough ramp/orbits hits to start a new battle?
            If tmp = 4 OR(OrbitHits + RampHits = 4) Then
                DMD "_", CL("START BATTLE IS LIT"), "_", eNone, eBlink, eNone, 2500, True, ""
                bModeReady = True
                li001.State = 2 'the start battle lights
                li038.State = 2
            End If
      'reset the GOKU/zeus lights and count the number of times they have been lit to start the GOKU/Zeus multiball
            If tmp = 4 Then
                AddPlayfieldMultiplier 1
                li061.State = 0
                li062.State = 0
                li063.State = 0
                li064.State = 0
                ZeusCount(CurrentPlayer) = ZeusCount(CurrentPlayer) + 1
        If ZeusCount(CurrentPlayer) = 1 Then
          DMD CL("GOKU TARGETS HIT"), CL("TWICE MORE"), "_", eNone, eBlink, eNone, 3500, True, ""
        Else If ZeusCount(CurrentPlayer) = 2 Then
          DMD CL("GOKU TARGETS HIT"), CL("ONCE MORE"), "_", eNone, eBlink, eNone, 3500, True, ""
        End If
        End If
                If ZeusCount(CurrentPlayer) MOD 3 = 0 then 'start GOKU/zeus multiball
          GokuVisuals
                    StartZeusMB
                End If
            End If
        Case 8 'GOKU/Zeus training
            If tmp = 4 Then
                AddPlayfieldMultiplier 1
                li061.State = 0
                li062.State = 0
                li063.State = 0
                li064.State = 0
                ZeusTargetsCompleted = ZeusTargetsCompleted + 1
                If Light009.State = 0 Then
                    DMD "_", CL("EXTRA BALL IS LIT"), "_", eNone, eBlink, eNone, 2500, True, "vo_extraballislit"
                    Light009.State = 2
                End If
                CheckWinMode
            End If
        Case Else 'check also for GOKU/Zeus multiball during other battles
'     reset the GOKU/Zeus lights and count the number of times they have been lit to start the GOKU/Zeus multiball
            If tmp = 4 Then
                AddPlayfieldMultiplier 1
                li061.State = 0
                li062.State = 0
                li063.State = 0
                li064.State = 0
                ZeusCount(CurrentPlayer) = ZeusCount(CurrentPlayer) + 1
        If ZeusCount(CurrentPlayer) = 1 Then
          DMD CL("GOKU TARGETS HIT"), CL("TWICE MORE"), "_", eNone, eBlink, eNone, 3500, True, ""
        Else If ZeusCount(CurrentPlayer) = 2 Then
          DMD CL("GOKU TARGETS HIT"), CL("ONCE MORE"), "_", eNone, eBlink, eNone, 3500, True, ""
        End If
        End If
                If ZeusCount(CurrentPlayer) MOD 3 = 0 then 'start GOKU/zeus multiball
          GokuVisuals
                    StartZeusMB
                End If
            End If
    End Select
End Sub

Sub StartNextMode 'starts the mode in the Mode(CurrentPlayer, 0) variable
    Dim tmp
  'If CurrentMode(CurrentPlayer) = 8
    CurrentMode(CurrentPlayer) = CurrentMode(CurrentPlayer) + 1
    Mode(CurrentPlayer, 0) = CurrentMode(CurrentPlayer)
    ChangeSong
  EnableBallSaver (13)  '10 seconds + 3 second kickout timer

    Select Case Mode(CurrentPlayer, 0)
    Case 1 ' Cell
            'OPTIONAL 120 second timer, using TimedModes variable in UserOption at top
            'REQUIRED SHOTS: left orbit, right orbit, center orbit, final shot being the center scoop.
      If bCellCompleted (CurrentPlayer) = True Then
        StartNextMode ' Skip to next battle.  Used in game loop situtations.
      Else
        Cell_L.State = 1
        DMD RL("CELL BATTLE"), RL("SHOOT LIT SHOTS"), "d_minotaur", eNone, eBlink, eNone, 4000, True, "vo_cell" &RndNbr(9)
        li013.BlinkInterval = 100
        Mode(CurrentPlayer, 1) = 2
        UpdateModeLights
        EndModeCountdown = 120
        If TimedModes Then EndModeTimer.Enabled = 1
        ModeStep = 1
        'arrows used in the mode
        li060.State = 2
        ChangeGi green
        ChangeGIIntensity 3.0
        vpmtimer.addtimer 3000, "kickBallOut '"
      End If
        Case 2 ' Zamasu
            'OPTIONAL 120 second timer, using TimedModes variable in UserOption at top
            'REQUIRED SHOTS: left ramp, right ramp, left ramp, right ramp, final shot being the center scoop.
      If bZamasuCompleted (CurrentPlayer) = True Then
        StartNextMode ' Skip to next battle.  Used in game loop situtations.
      Else
        Zamasu_L.State = 1
        DMD RL("ZAMASU BATTLE"), RL("SHOOT LIT SHOTS"), "d_hydra", eNone, eBlink, eNone, 4000, True, "vo_zamasu" &RndNbr(9)
        li011.BlinkInterval = 100
        Mode(CurrentPlayer, 2) = 2
        UpdateModeLights
        EndModeCountdown = 120
        If TimedModes Then EndModeTimer.Enabled = 1
        ModeStep = 1
        'arrows used in the mode
        li036.State = 2
        ChangeGi purple
        ChangeGIIntensity 3.0
        vpmtimer.addtimer 3000, "kickBallOut '"
      End If
        Case 3 ' Goku Black
      'OPTIONAL 120 second timer, using TimedModes variable in UserOption at top
      'REQUIRED SHOTS: left ramp, right ramp, and both upper ramps, (any order) final shot being the center scoop.
      If bGokuBlackCompleted (CurrentPlayer) = True Then
        StartNextMode ' Skip to next battle.  Used in game loop situtations.
      Else
        GokuBlack_L.State = 1
        DMD RL("GOKU BLACK BATTLE"), RL("SHOOT LIT SHOTS"), "d_cerberus", eNone, eBlink, eNone, 4000, True, "vo_gokublack" &RndNbr(5)
        li012.BlinkInterval = 100
        Mode(CurrentPlayer, 3) = 2
        UpdateModeLights
        EndModeCountdown = 120
        If TimedModes Then EndModeTimer.Enabled = 1
        ModeStep = 1
        'arrows used in the mode
        li036.State = 2
        li040.State = 2
        li042.State = 2
        li044.State = 2
        ChangeGi red
        ChangeGIIntensity 2.2
        vpmtimer.addtimer 3000, "kickBallOut '"
      End If
        Case 4 ' Golden Frieza
      'OPTIONAL 120 second timer, using TimedModes variable in UserOption at top
      'REQUIRED SHOTS: left orbit X2, right orbit X2, and center orbit X2, final shot being the center scoop.
      'Completing the mode awards a five ball Medusa Multi-ball. 15 seconds ball saver.
      If bFriezaCompleted (CurrentPlayer) = True Then
        StartNextMode ' Skip to next battle.  Used in game loop situtations.
      Else
        Frieza_L.State = 1
        DMD RL("FRIEZA BATTLE"), RL("SHOOT LIT SHOTS"), "d_medusa", eNone, eBlink, eNone, 4000, True, "vo_frieza" &RndNbr(9)
        li014.BlinkInterval = 100
        Mode(CurrentPlayer, 4) = 2
        UpdateModeLights
        EndModeCountdown = 120
        If TimedModes Then EndModeTimer.Enabled = 1
        ModeStep = 1
        'arrows used in the mode
        li060.State = 2
        li041.State = 2
        li037.State = 2
        ChangeGi yellow
        ChangeGIIntensity 2.0
        vpmtimer.addtimer 3000, "kickBallOut '"
        EnableBallSaver (20)
      End If
        Case 5 ' Trunks
            'OPTIONAL 120 second timer, using TimedModes variable in UserOption at top
            'REQUIRED SHOTS: 100 spinners and hit the scoop to finish.
            'check for extra jackpots
      If bTrunksCompleted (CurrentPlayer) = True Then
        StartNextMode ' Skip to next battle.  Used in game loop situtations.
      Else
        If li014.State = 1 AND li012.State = 1 Then  'Check surrounding lights for immediate reward
          DMD "JACKPOTS ENABLED", "HIT RAMPS TO COLLECT", "_", eNone, eNone, eNone, 4000, True, ""
          li042.State = 2
          li044.State = 2
          li036.State = 2
          li040.State = 2
          bJackpotsEnabled = True
        End If
        Trunks_PoweredUp.visible = 1
        DMD RL("TRUNKS TRAINING"), RL("SHOOT SPINNERS"), "d_ares", eNone, eBlink, eNone, 4000, True, "vo_Trunks" &RndNbr(7)
        li009.BlinkInterval = 100
        Mode(CurrentPlayer, 5) = 2
        UpdateModeLights
        EndModeCountdown = 120
        If TimedModes Then EndModeTimer.Enabled = 1
        ModeStep = 1
        SpinnerHits = 0
        'arrows used in the mode
        li037.State = 2
        li041.State = 2
        ChangeGi purple
        ChangeGIIntensity 3.0
        vpmtimer.addtimer 3000, "kickBallOut '"
      End If
        Case 6 ' Vegeta
            'OPTIONAL 120 second timer, using TimedModes variable in UserOption at top
            'REQUIRED SHOTS: All the lit targets and finish at the scoop.
            'check for extra ball
            If bVegetaCompleted (CurrentPlayer) = True Then
        StartNextMode ' Skip to next battle.  Used in game loop situtations.
      Else
        If li013.State = 1 AND li011.State = 1 Then 'Check surrounding lights for immediate reward
          If Light009.State = 0 Then
            DMD "_", CL("EXTRA BALL IS LIT"), "_", eNone, eBlink, eNone, 4000, True, "vo_extraballislit"
            Light009.State = 2
          End If
        End If
        If bJackpotsEnabled Then
          bJackpotsEnabled = False 'stops the jackpots if they were enabled
          li036.State = 0
          li040.State = 0
          li042.State = 0
          li044.State = 0
        End If
        Vegeta_PoweredUp.visible = 1
        DMD RL("VEGETA TRAINING"), RL("HIT LIT TARGETS"), "d_poseidon", eNone, eBlink, eNone, 4000, True, "vo_vegeta" &RndNbr(8)
        li007.BlinkInterval = 100
        Mode(CurrentPlayer, 6) = 2
        UpdateModeLights
        EndModeCountdown = 120
        If TimedModes Then EndModeTimer.Enabled = 1
        ModeStep = 1
        'lights used in the mode
        TargetLightsAll 2
        ChangeGi blue
        ChangeGIIntensity 3.5
        vpmtimer.addtimer 3000, "kickBallOut '"
      End If
        Case 7 ' Piccolo
            'OPTIONAL 120 second timer, using TimedModes variable in UserOption at top
            'REQUIRED SHOTS: Hit trap door 5 times
            'check for Shoot Again
            If bPiccoloCompleted (CurrentPlayer) = True Then
        StartNextMode ' Skip to next battle.  Used in game loop situtations.
      Else
        Piccolo_PoweredUp.visible = 1
        If li012.State = 1 AND li011.State = 1 Then 'Check surrounding lights for immediate reward
          AwardExtraBall
        End If
        DMD RL("PICCOLO TRAINING"), RL("HIT TRAPDOOR"), "d_hades", eNone, eBlink, eNone, 4000, True, "vo_piccolo" &RndNbr(9)
        li010.BlinkInterval = 100
        Mode(CurrentPlayer, 7) = 2
        UpdateModeLights
        EndModeCountdown = 120
        If TimedModes Then EndModeTimer.Enabled = 1
        ModeStep = 1
        TrapDoorHits = 0
        TrapdoorUp
        ChangeGi green
        ChangeGIIntensity 3.6
        vpmtimer.addtimer 3000, "kickBallOut '"
      End If
        Case 8 ' Goku
            'OPTIONAL 180 second timer, using TimedModes variable in UserOption at top
            'REQUIRED SHOTS: Complete ZEUS targets 3 times and finish at the scoop.
            'check for special
            If li012.State = 1 AND li014.State = 1 Then 'Check surrounding lights for immediate reward
                DMD CL("SPECIAL"), CL("IS LIT"), "_", eNone, eNone, eNone, 2500, True, "vo_specialislit"
                Light010.State = 2
            End If
            DMD RL("GOKU TRAINING"), RL("SPELL GOKU 3 TIMES"), "d_zeus", eNone, eBlink, eNone, 4000, True, ""
            li008.BlinkInterval = 100
            Mode(CurrentPlayer, 8) = 2
            UpdateModeLights
            EndModeCountdown = 180
            If TimedModes Then EndModeTimer.Enabled = 1
            ModeStep = 1
            'lights used in the mode
      li042.state = 2 'point player to correct direction
      li044.state = 2 'point player to correct direction
      li061.state = 0 'Start with being off
      li062.state = 0
      li063.state = 0
      li064.state = 0
            ChangeGi yellow
            ChangeGIIntensity 2.0
            vpmtimer.addtimer 19500, "GokuVisuals '"    'Table shake and Ultra visuals
            ZeusTargetsCompleted = 0 'reset the count
      If bMultiBallMode = True Then
        bMultiBallMode = False
        SolLFlipper 0
        SolRFlipper 0
        StopSound "vo_kamehamehaL1"
        StopSound "vo_kamehamehaL2"
        StopSound "vo_kamehamehaR1"
        StopSound "vo_kamehamehaR2"
        bInstantInfo = False
        instantInfoTimer.Enabled = False
        bGameInPlay = False
        bBallSaverActive = False
        bBallSaverReady = False
        TurnOffPFLayers
      End If
      vpmtimer.addtimer 20000, "DelayGoku '"        '"DMD CL(""    GET READY""), CL(""    GET READY""), ""_"", eBlink, eBlink, eNone, 2500, True, """" '"
        Case 9  ' King Kai or Supreme Kai WIZARD mode
            '60 / 120 seconds timer for wizard mode
            'REQUIRED SHOTS: None,
      '5 ball multiball, jackpots enabled on the ramps and orbits.
      bKaiWizardModeRan = True
      bKaiModeStarted = True
            If bGokuCompleted (CurrentPlayer) = True Then
                DMD RL("SUPREME KAI MODE"), RL("SHOOT JACKPOTS"), "d_kai", eNone, eBlink, eNone, 2500, True, ""
                EndModeCountdown = 120
        DMD RL("MODE ENDS IN 120 SEC"), RL("SHOOT JACKPOTS"), "d_kai", eNone, eBlink, eNone, 2500, True, ""
        EnableBallSaver (25)
            Else
                DMD RL("KING KAI MODE"), RL("SHOOT JACKPOTS"), "d_kai", eNone, eBlink, eNone, 2500, True, ""
                EndModeCountdown = 60
        DMD RL("MODE ENDS IN 60 SEC"), RL("SHOOT JACKPOTS"), "d_kai", eNone, eBlink, eNone, 2500, True, ""
        EnableBallSaver (15)
            End If
            Mode(CurrentPlayer, 9) = 2
            EndModeTimer.Enabled = 1
            'lights used in the mode - jackpots on the ramps
            li044.State = 2
            li042.State = 2
            li036.State = 2
            li040.State = 2
            StartRainbow aGiLights
            ChangeGIIntensity 2
      GokuVisualStep = 0 'Reset step
      Gokuvisualbluetimer.Enabled = 1  'Table shake
      Goku_UltraInstinct20.visible = 0
      Goku_UltraInstinct.visible = 1
        LightSeqGoku.UpdateInterval = 100
      LightSeqGoku.Play SeqRandom, 50, , 10000000
            Vegeta_PoweredUp.visible = 1
            Trunks_PoweredUp.visible = 1
            Piccolo_PoweredUp.visible = 1
      plastic_KaiMode.visible = 1
      bAutoPlunger = True
            AddMultiball 3
    End Select
End Sub

Sub DelayGoku
  bGameInPlay = True
  BallsinHole = 1
  DMD CL("    GET READY"), CL("    GET READY"), "_", eBlink, eBlink, eNone, 2500, True, ""
  EnableBallSaver 24 '20 seconds default + 4 for the kickout delay
  vpmtimer.addtimer 4000, "kickBallOut '"
  bZeusStarted = True 'Used for music loop
End Sub

Dim GokuVisualStep

Sub GokuVisuals
    GokuVisualStep = 0
    Goku_UltraInstinct20.visible = 0 'this is mostly when testing
    gokuvisualtimer.Enabled = 1
    LightSeqGoku.UpdateInterval = 100
    LightSeqGoku.Play SeqRandom, 50, , 10000000
End sub

Sub ShakerMotor
    ShakerStep = 0
    shakertimer.Enabled = 1
End Sub

Sub gokuvisualtimer_Timer    'Table shake with Ultra Instinct visuauls flashing
    Select Case GokuVisualStep
        Case 0, 3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 39, 42, 45, 48, 51, 54, 57, 60, 63, 66, 69, 72
            shakertimer.Enabled = 1
            GiOff
            Goku_UltraInstinct40.visible = 0
            Goku_UltraInstinct80.visible = 1
        Case 1, 4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37, 40, 43, 46, 49, 52, 55, 58, 61, 64, 67, 70, 73, 76, 79, 82, 85, 88, 91, 94, 97, 100, 103, 106, 109, 112, 115, 118, 221, 124, 127, 130, 133, 136, 139, 142, 145, 148, 151, 154, 157
            ShakerMotor
            GiOff
            Goku_UltraInstinct80.visible = 0
            Goku_UltraInstinct60.visible = 1
        Case 2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44, 47, 50, 53, 56, 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, 89, 92, 95, 98, 101, 104, 107, 110, 113, 116, 119, 222, 125, 128, 131, 134, 137, 140, 143, 146, 149, 152, 155, 158
            GiOn
            Goku_UltraInstinct60.visible = 0
            Goku_UltraInstinct40.visible = 1
        Case 75, 78, 81, 84, 87, 90, 93, 96, 99, 102, 105, 108, 111, 114, 117, 120, 123, 126, 129, 132, 135, 138, 141, 144, 147, 150, 153, 156, 159
            GiOn
            Goku_UltraInstinct40.visible = 0
            Goku_UltraInstinct20.visible = 1
        Case 166
            gokuvisualtimer.Enabled = 0
    End Select
    GokuVisualStep = GokuVisualStep + 1
End Sub

Sub gokuvisualBLUEtimer_Timer    'Table shake with NO visuauls
    Select Case GokuVisualStep
        Case 0, 3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 39, 42, 45, 48, 51, 54, 57, 60, 63, 66, 69, 72
            shakertimer.Enabled = 1
            GiOff
        Case 1, 4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37, 40, 43, 46, 49, 52, 55, 58, 61, 64, 67, 70, 73, 76, 79, 82, 85, 88, 91, 94, 97, 100, 103, 106, 109, 112, 115, 118, 221, 124, 127, 130, 133, 136, 139, 142, 145, 148, 151, 154, 157
            ShakerMotor
            GiOff
        Case 2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44, 47, 50, 53, 56, 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, 89, 92, 95, 98, 101, 104, 107, 110, 113, 116, 119, 222, 125, 128, 131, 134, 137, 140, 143, 146, 149, 152, 155, 158
            GiOn
        Case 75, 78, 81, 84, 87, 90, 93, 96, 99, 102, 105, 108, 111, 114, 117, 120, 123, 126, 129, 132, 135, 138, 141, 144, 147, 150, 153, 156, 159
            GiOn
        Case 166
            gokuvisualBLUEtimer.Enabled = 0
    End Select
    GokuVisualStep = GokuVisualStep + 1
End Sub

Dim ShakerStep
Sub shakertimer_timer
    Select Case ShakerStep
        Case 0:Nudge RndNbr(350), 1
        Case 1:Nudge RndNbr(350), 1
        Case 2:Nudge RndNbr(350), 1
        Case 3:Nudge RndNbr(350), 1
        Case 4:Nudge RndNbr(350), 1
        Case 5:Nudge RndNbr(350), 1
        Case 6:shakertimer.Enabled = 0
    End Select
    ShakerStep = ShakerStep + 1
End Sub

' Update the lights according to the mode's state, 0 not started, 1 finished, 2 started
Sub UpdateModeLights
    li013.State = Mode(CurrentPlayer, 1)
    li011.State = Mode(CurrentPlayer, 2)
    li012.State = Mode(CurrentPlayer, 3)
    li014.State = Mode(CurrentPlayer, 4)
    li009.State = Mode(CurrentPlayer, 5)
    li007.State = Mode(CurrentPlayer, 6)
    li010.State = Mode(CurrentPlayer, 7)
    li008.State = Mode(CurrentPlayer, 8)
End Sub

Sub CheckWinMode
    Dim tmp    'when you complete one the tasks
    Select Case Mode(CurrentPlayer, 0)
        Case 1 ' CELL/Minotaur
            Select Case ModeStep
                Case 1:li060.State = 0:li041.State = 2:ModeStep = 2
                Case 2:li041.State = 0:li037.State = 2:ModeStep = 3
                Case 3:
                    li037.State = 0
                    TurnOnHelmet
                    li039.state = 2
                    ModeStep = 4
            End Select
        Case 2 ' ZAMASU/Hydra
            Select Case ModeStep
                Case 1:li036.State = 0:li040.State = 2:ModeStep = 2
                Case 2:li040.State = 0:li036.State = 2:ModeStep = 3
                Case 3:li036.State = 0:li040.State = 2:ModeStep = 4
                Case 4:
                    li040.State = 0
                    TurnOnHelmet
                    li039.state = 2
                    ModeStep = 5
            End Select
        Case 3 ' GOKU BLACK/Cerberus
            If ModeStep = 1 Then
                If li036.State + li040.State + li044.State + li042.State = 0 Then 'ALL ramps lights are out then move to step 2
                    ModeStep = 2
                    TurnOnHelmet
                    li039.state = 2
                End If
            End If
        Case 4 ' GOLDEN FRIEZA/Medusa
            Select Case ModeStep
                Case 1
                    If li060.State + li041.State + li037.State = 0 Then 'turn on the lights again for a second round
                        li060.State = 2
                        li041.State = 2
                        li037.State = 2
                        ModeStep = 2
                    End If
                Case 2
                    If li060.State + li041.State + li037.State = 0 Then
                        TurnOnHelmet
                        li039.state = 2
                        ModeStep = 3
                    End If
            End Select
        Case 5                             ' TRUNKS/Ares
            If ModeStep = 1 Then
                If SpinnerHits >= 100 Then 'move to step 2
                    li037.State = 0
                    li041.State = 0
                    ModeStep = 2
                    TurnOnHelmet
                    li039.state = 2
                End If
            End If
        Case 6 ' VEGETA/Poseidon
            If ModeStep = 1 Then
                tmp = 0
                For each x in aTargetsAll
                    tmp = tmp + x.State
                Next
                If tmp = 0 Then 'move to step 2
                    ModeStep = 2
                    TurnOnHelmet
                    li039.state = 2
                End If
            End If
        Case 7 ' PICCOLO/Hades
            If TrapDoorHits = 5 Then
                WinMode
            Else
                vpmtimer.addtimer 2000, "kickBallOut '"
            End If
        Case 8 ' Goky/Zeus
            If ModeStep = 1 Then
                If ZeusTargetsCompleted = 3 Then
                    ModeStep = 2
                    TurnOnHelmet
                    li039.state = 2
                End If
            End If
        Case 9 ' This is a bonus stage for collecting jackpots.  Mode is active for 60 or 120 seconds and until time is up or loose all the balls.
    End Select
End Sub

Sub WinMode 'when you complete all the tasks
    GiEffect 1
    LightEffect 2
    FlashEffect RndNbr(7)
    ZeusF
    Select Case Mode(CurrentPlayer, 0)
        Case 1                         ' CELL/Minotaur
            Mode(CurrentPlayer, 1) = 1 'set the mode as finished, and it will make the light solid lit (UpdateModeLights)
            TotalMonsters(CurrentPlayer) = TotalMonsters(CurrentPlayer) + 1
            DMDFlush
            DMD "_", RL("CELL DEFEATED"), "_", eNone, eScrollLeft, eNone, 20, True, "sfx_thunder" & RndNbr(9)
            DMD "_", RL("CELL DEFEATED"), "_", eNone, eBlinkFast, eNone, 2000, True, ""
            DMD "_", SPACE(20), "_", eNone, eScrollLeft, eNone, 20, True, ""
            ModeStep = 0
            AddScore2 1000000
      bCellCompleted (CurrentPlayer) = True
            Cell_L.State = 0 'Turns off pf overlay
            Cell_DeathHalo.visible = 1 'Turns ON pf overlay for death halo
            TurnOffHelmet
            vpmtimer.addtimer 2500, "kickBallOut '"
        Case 2 ' ZAMASU/Hydra
            Mode(CurrentPlayer, 2) = 1
            TotalMonsters(CurrentPlayer) = TotalMonsters(CurrentPlayer) + 1
            DMDFlush
            DMD "_", RL("ZAMASU DEFEATED"), "_", eNone, eScrollLeft, eNone, 20, True, "sfx_thunder" & RndNbr(9)
            DMD "_", RL("ZAMASU DEFEATED"), "_", eNone, eBlinkFast, eNone, 2000, True, ""
            DMD "_", SPACE(20), "_", eNone, eScrollLeft, eNone, 20, True, ""
            ModeStep = 0
            AddScore2 1000000
      bZamasuCompleted (CurrentPlayer) = True
            Zamasu_L.State = 0 'Turns off pf overlay
            Zamasu_DeathHalo.visible = 1 'Turns ON pf overlay for death halo
            TurnOffHelmet
            vpmtimer.addtimer 2500, "kickBallOut '"
        Case 3 ' GOKU BLACK/Cerberus
            Mode(CurrentPlayer, 3) = 1
            TotalMonsters(CurrentPlayer) = TotalMonsters(CurrentPlayer) + 1
            DMDFlush
            DMD "_", RL("GOKU BLACK DEFEATED"), "_", eNone, eScrollLeft, eNone, 20, True, "sfx_thunder" & RndNbr(9)
            DMD "_", RL("GOKU BLACK DEFEATED"), "_", eNone, eBlinkFast, eNone, 2000, True, ""
            DMD "_", SPACE(20), "_", eNone, eScrollLeft, eNone, 20, True, ""
            ModeStep = 0
            AddScore2 1000000
      bGokuBlackCompleted (CurrentPlayer) = True
            GokuBlack_L.State = 0 'Turns off pf overlay
            GokuBlack_DeathHalo.visible = 1 'Turns ON pf overlay for death halo
            TurnOffHelmet
            vpmtimer.addtimer 2500, "kickBallOut '"
        Case 4 ' GOLDEN FRIEZA/Medusa
            Mode(CurrentPlayer, 4) = 1
            TotalMonsters(CurrentPlayer) = TotalMonsters(CurrentPlayer) + 1
            DMDFlush
            DMD "_", RL("FRIEZA DEFEATED"), "_", eNone, eScrollLeft, eNone, 20, True, "sfx_thunder" & RndNbr(9)
            DMD "_", RL("FRIEZA DEFEATED"), "_", eNone, eBlinkFast, eNone, 2000, True, ""
            DMD "_", SPACE(20), "_", eNone, eScrollLeft, eNone, 20, True, ""
            ModeStep = 0
            AddScore2 1000000
      bFriezaCompleted (CurrentPlayer) = True
            Frieza_L.State = 0             'Turns off pf overlay
            Frieza_DeathHalo.visible = 1             'Turns ON pf overlay for death halo
            TurnOffHelmet
            If TotalMonsters(CurrentPlayer) = 4 Then 'all monsters have been defeated, turn on the beast master light
                DMD "YOU DEFEATED", CL("ALL ENEMIES"), "_", eNone, eNone, eNone, 2000, True, "vo_beast_master"
                li047.State = 1
                AddScore2 1000000
            End If
      StartMedussaMB                           'start medusa multiball
            vpmtimer.addtimer 4000, "kickBallOut '"
        Case 5 ' TRUNKS/Ares
            Mode(CurrentPlayer, 5) = 1
            TotalGods(CurrentPlayer) = TotalGods(CurrentPlayer) + 1
            DMDFlush
            DMD "_", RL("TRAINING COMPLETED"), "_", eNone, eScrollLeft, eNone, 20, True, "sfx_thunder" & RndNbr(9)
            DMD "_", RL("TRAINING COMPLETED"), "_", eNone, eBlinkFast, eNone, 2000, True, ""
            DMD "_", SPACE(20), "_", eNone, eScrollLeft, eNone, 20, True, ""
            ModeStep = 0
            AddScore2 1000000
      bTrunksCompleted (CurrentPlayer) = True
            TurnOffHelmet
            vpmtimer.addtimer 3000, "kickBallOut '"
        Case 6 ' VEGETA/Poseidon
            Mode(CurrentPlayer, 6) = 1
            TotalGods(CurrentPlayer) = TotalGods(CurrentPlayer) + 1
            DMDFlush
            DMD "_", RL("TRAINING COMPLETED"), "_", eNone, eScrollLeft, eNone, 20, True, "sfx_thunder" & RndNbr(9)
            DMD "_", RL("TRAINING COMPLETED"), "_", eNone, eBlinkFast, eNone, 2000, True, ""
            DMD "_", SPACE(20), "_", eNone, eScrollLeft, eNone, 20, True, ""
            ModeStep = 0
            AddScore2 1000000
      bVegetaCompleted (CurrentPlayer) = True
            TurnOffHelmet
            vpmtimer.addtimer 3000, "kickBallOut '"
        Case 7 ' PICCOLO/Hades
            Mode(CurrentPlayer, 7) = 1
            TotalGods(CurrentPlayer) = TotalGods(CurrentPlayer) + 1
            DMDFlush
            DMD "_", RL("TRAINING COMPLETED"), "_", eNone, eScrollLeft, eNone, 20, True, "sfx_thunder" & RndNbr(9)
            DMD "_", RL("TRAINING COMPLETED"), "_", eNone, eBlinkFast, eNone, 2000, True, ""
            DMD "_", SPACE(20), "_", eNone, eScrollLeft, eNone, 20, True, ""
            ModeStep = 0
            AddScore2 5000000
            bPiccoloCompleted (CurrentPlayer) = True
            TurnOffHelmet
            vpmtimer.addtimer 3000, "kickBallOut '"
        Case 8 ' GOKU/Zeus
            Mode(CurrentPlayer, 8) = 1
            TotalGods(CurrentPlayer) = TotalGods(CurrentPlayer) + 1
            DMDFlush
            DMD "_", RL("TRAINING COMPLETED"), "_", eNone, eScrollLeft, eNone, 20, True, "sfx_thunder" & RndNbr(9)
            DMD "_", RL("TRAINING COMPLETED"), "_", eNone, eBlinkFast, eNone, 2000, True, ""
            DMD "_", SPACE(20), "_", eNone, eScrollLeft, eNone, 20, True, ""
            ModeStep = 0
            AddScore2 5000000
      bGokuCompleted (CurrentPlayer) = True
            Goku_UltraInstinct20.visible = 0 'Turns off pf overlay and sound
      StopSound "vo_GokuScreamPowerUp" '
            TurnOffHelmet
            vpmtimer.addtimer 3000, "kickBallOut '"
            If TotalGods(CurrentPlayer) = 4 Then 'all gods has been defeated
                DMD "YOU COMPLETED", CL("ALL TRAININGS"), "_", eNone, eNone, eNone, 2000, True, "vo_god_of_gods"
                AddScore2 2000000
            End If
        Case 9 ' God or Demi-God mode - (the mode is a multiball, it stops when you run out of time or lose the balls)
    End Select
    StopMode
  ChangeSong
End Sub

Sub StopMode                 'called after a win or at the end of a ball to stop the current mode variables and timers
    EndModeTimer.Enabled = 0 'ensure it is stopped
    TurnOffArrows
    Select Case Mode(CurrentPlayer, 0)
        Case 1                         ' CELL/Minotaur
            li013.BlinkInterval = 1000 'slow blink in case the mode is not finished
        Case 2                         ' ZAMASU/Hydra
            li011.BlinkInterval = 1000
        Case 3                         ' GOKU BLACK/Cerberus
            li012.BlinkInterval = 1000
        Case 4                         ' GOLDEN FRIEZA/Medusa
            li014.BlinkInterval = 1000
        Case 5                         ' TRUNKS/Ares
            li009.BlinkInterval = 1000
        Case 6                         ' VEGETA/Poseidon
            li007.BlinkInterval = 1000
            TargetLightsAll 0
      If KickbackEnabled then li015.State = 1: li016.State = 1
        Case 7                         ' PICCOLO/Hades
            li010.BlinkInterval = 1000
            vpmtimer.addtimer 2500, "TrapdoorDown '"

        Case 8                         ' GOKU/Zeus
            li008.BlinkInterval = 1000
      PlayNextSong.Enabled = 0
      bZeusStarted = False
            ' and start the wizard mode
            vpmtimer.addtimer 1500, "EnableBallSaver 15:StartNextMode '" 'wait a little before start the wizard mode, so the last mode stops and all the variables are set up right
    Case 9                                                          ' God or Demi-God mode -
            bKaiModeStarted = False
      Goku_UltraInstinct.visible = 0
      Goku_UltraInstinct20.visible = 0
      If bTrunksCompleted (CurrentPlayer) = False Then
        Trunks_PoweredUp.visible = 0        'Layers only turn off, if mode not completed
      End If
      If bVegetaCompleted (CurrentPlayer) = False Then
        Vegeta_PoweredUp.visible = 0        'Layers only turn off, if mode not completed
      End If
      If bPiccoloCompleted (CurrentPlayer) = False Then
        Piccolo_PoweredUp.visible = 0       'Layers only turn off, if mode not completed
      End If
      plastic_KaiMode.visible = 0
      ResetModes
    End Select
    ' reset variables
    ModeStep = 0
    UpdateModeLights
    OrbitHits = 0 'start counting again for the next mode
    RampHits = 0
    Mode(CurrentPlayer, 0) = 0
    bModeReady = False
    ChangeGi white
    ChangeGIIntensity 1
End Sub

Sub ResetModes 'called after the last wizard mode to start all over again
    Dim i, j
    For i = 0 to 9
        Mode(CurrentPlayer, i) = 0
    Next
    StopRainbow
    UpdateModeLights
    'reset Mode variables
    CurrentMode(CurrentPlayer) = 0
    bModeReady = False
End Sub

Sub EndModeTimer_Timer '1 second timer to count down to end the timed modes
    EndModeCountdown = EndModeCountdown - 1
    Select Case EndModeCountdown
    Case 30
      If Mode(CurrentPlayer, 0) = 9 Then
        DMD "_", RL("30 SECONDS"), "d_kai", eNone, eBlinkFast, eNone, 2000, True, ""
      End If
        Case 16:DMD "_", RL("TIME IS RUNNING OUT"), "_", eNone, eNone, eNone, 1000, True, ""
        Case 10:DMD "_", CL("10"), "_", eNone, eNone, eNone, 500, True, ""
        Case 9:DMD "_", CL("9"), "_", eNone, eNone, eNone, 500, True, ""
        Case 8:DMD "_", CL("8"), "_", eNone, eNone, eNone, 500, True, ""
        Case 7:DMD "_", CL("7"), "_", eNone, eNone, eNone, 500, True, ""
        Case 6:DMD "_", CL("6"), "_", eNone, eNone, eNone, 500, True, ""
        Case 5:DMD "_", CL("5"), "_", eNone, eNone, eNone, 500, True, ""
        Case 4:DMD "_", CL("4"), "_", eNone, eNone, eNone, 500, True, ""
        Case 3:DMD "_", CL("3"), "_", eNone, eNone, eNone, 500, True, ""
        Case 2:DMD "_", CL("2"), "_", eNone, eNone, eNone, 500, True, ""
        Case 1:DMD "_", CL("1"), "_", eNone, eNone, eNone, 500, True, ""
        Case 0
            If Mode(CurrentPlayer, 0) = 8 Then
                DMD RL("GET READY FOR"), RL("KAI WIZARD MODE"), "d_kai", eNone, eBlinkFast, eNone, 2000, True, ""
            Else If Mode(CurrentPlayer, 0) = 9 Then
        DMD RL("KAI MODE COMPLETE"), RL("WELL DONE"), "d_kai", eNone, eBlinkFast, eNone, 1000, True, ""
      Else
                DMD RL("TIME IS UP"), RL("YOU FAILED"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
            End If
      End If
            StopMode
    End Select
End Sub

Sub TurnOnHelmet 'to end a battle
    Light003.State = 2
End Sub

Sub TurnOffHelmet
    Light003.State = 0
End Sub

Sub TurnOffArrows 'at the end of the ball or timed mode
    li060.State = 0
    li036.State = 0
    li038.State = 0
    li041.State = 0
    li037.State = 0
    li044.State = 0
    li040.State = 0
    li042.State = 0
    li039.State = 0
End Sub

'*************
' Magnet
'*************

Sub ReleaseMagnetBalls 'mMagnet off and release the ball if any
    Dim ball
    mMagnet.MagnetOn = False
    DOF 132, DOFOff
    For Each ball In mMagnet.Balls
        With ball
            .VelX = 1
            .VelY = 1
        End With
    Next
End Sub

'**************
'   COMBOS
'**************

Sub AwardCombo
    ComboCount = ComboCount + 1
    DOF 130, DOFPulse
    Select Case ComboCount
        Case 0, 1:Exit Sub 'this should never happen though
        Case 2
            DMD CL("COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 1500, True, "vo_combo"
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 3
            DMD CL("2X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * ComboCount) ), "", eNone, eNone, eNone, 1500, True, "vo_combo"
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 4
            DMD CL("3X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * ComboCount) ), "", eNone, eNone, eNone, 1500, True, "vo_combo"
            ComboValue(CurrentPlayer) = ComboValue(CurrentPlayer) + 100000:ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 5
            DMD CL("FLYING NIMBUS"), CL(FormatScore(ComboValue(CurrentPlayer) * ComboCount) ), "", eNone, eNone, eNone, 1500, True, "vo_hurricane"
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            If Light009.State = 0 Then
                DMD "_", CL("EXTRA BALL IS LIT"), "_", eNone, eBlink, eNone, 2500, True, "vo_extraballislit"
                Light009.State = 2
            End If
        Case 6
            DMD CL("FLYING NIMBUS X2"), CL(FormatScore(ComboValue(CurrentPlayer) * ComboCount) ), "", eNone, eNone, eNone, 1500, True, "vo_combo_king"
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            li046.State = 1
            If Light010.State = 0 Then
                DMD CL("_"), CL("SPECIAL IS LIT"), "_", eNone, eBlink, eNone, 2000, True, "vo_specialislit"
                Light010.State = 2
            End If
        Case 7
            DMD CL("FLYING NIMBUS X3"), CL(FormatScore(ComboValue(CurrentPlayer) * ComboCount) ), "", eNone, eNone, eNone, 1500, True, "vo_wind_rider"
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            li048.State = 1
        Case Else 'enjoy the award because there is no prize. :p
            DMD CL("MASTER ROSHI AWARD"), CL(FormatScore(ComboValue(CurrentPlayer) * ComboCount) ), "", eNone, eNone, eNone, 1500, True, "vo_combo"
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
    End Select
    AddScore2 ComboValue(CurrentPlayer) * ComboCount
    ComboValue(CurrentPlayer) = ComboValue(CurrentPlayer) + 100000
End Sub

Sub aComboTargets_Hit(idx) 'reset the combo count if the ball hits another target/trigger
    ComboCount = 0
End Sub

'    MULTIBALLS

'********************
' SUPER SAIYAN BLUE/MEDUSA MB - LOOPS
'********************
' Starts after the GOLDEN FRIEZA/medusa monster Mode
' orbit shots doubles their values
' ramp shots are worth 1/2 of their value
' orbithits + ramphits are 5 or more then light the extra ball

Sub StartMedussaMB
  bMedusaMBStarted = True
  If bZeusMBStarted = True OR bKaiModeStarted = True Then 'Higher layer already on
    DMD CL("DOUBLE MULTIBALL"), CL("SHOOT THE LOOPS"), "d_zeus", eBlink, eNone, eNone, 3000, True, "vo_kaiokenx10"
  Else If bMinotaurMBStarted = True OR bCCorpMBStarted = True Then 'Higher layer already on
    DMD CL("DOUBLE MULTIBALL"), CL("SHOOT THE LOOPS"), "d_zeus", eBlink, eNone, eNone, 3000, True, "vo_kaiokenx2"
    StopSound "vo_AuraPoweredUp"          'Turn off lower layer first
    Goku_PoweredUp.visible = 0            '
    LightSeqGokuPU.StopPlay               '
    Goku_SS2.visible = 0                '
    StopSound "vo_AuraSS2"                '
    LightSeqGokuSS2.StopPlay   '
    Goku_Blue.visible = 1             'Turns on Blue
    GokuVisualStep = 0
    gokuvisualBLUEtimer.Enabled = 1               '
    PlaySound "vo_auraBlueWithVocals"         '
    LightSeqGokuBlue.UpdateInterval = 100     '
    LightSeqGokuBlue.Play SeqRandom, 50, , 10000000 '
    ChangeSong
  Else If CurrentMode(CurrentPlayer) = 8 OR CurrentMode(CurrentPlayer) = 9 Then 'Higher layer already on
    DMD CL("SUPER SAIYAN MBALL"), CL("SHOOT THE LOOPS"), "d_zeus", eNone, eNone, eNone, 2500, True, "vo_medusa_multiball"
  Else
    DMD CL("SUPER SAIYAN MBALL"), CL("SHOOT THE LOOPS"), "d_zeus", eNone, eNone, eNone, 2500, True, "vo_multiball"
    Goku_Blue.visible = 1
    GokuVisualStep = 0
    gokuvisualBLUEtimer.Enabled = 1
    PlaySound "vo_auraBlueWithVocals"
    LightSeqGokuBlue.UpdateInterval = 100
    LightSeqGokuBlue.Play SeqRandom, 50, , 10000000
    ChangeSong
  End If
  End If
  End If
    ExtraBallHits = 0
    AddMultiball 4
    EnableBallSaver 15
End Sub

' during medusa MB check for 5 orbits and light extra ball
Sub CheckExtraBallHits
    If bMedusaMBStarted Then
        If ExtraBallHits MOD 5 = 0 Then
            If Light009.State = 0 Then
                DMD "_", CL("EXTRA BALL IS LIT"), "_", eNone, eBlink, eNone, 2500, True, "vo_extraballislit"
                Light009.State = 2
            End If
        End If
    End If
End Sub

'****************************************************
' DRAGON RADAR/Minotaur MB - holes & lock system at the DRAGON RADAR/Ares hole
'****************************************************
' lock 3 balls, and MB starts
' holes score Jackpots
' value doubles each time all three holes has been Hit
' each hole gives just 1 jackpot until all three has been hit again.

Sub CheckMinotaurMB
    If BallsInLock(CurrentPlayer) = 1 Then
        DMD "_", CL("BALL 1 LOCKED"), "_", eNone, eNone, eNone, 1500, True, "fx_DragonRadar_2"
        PlaySound "vo_lock1"
        DragonRadar.image = "Radar_2"
    End If
    If BallsInLock(CurrentPlayer) = 2 Then
        DMD "_", CL("BALL 2 LOCKED"), "_", eNone, eNone, eNone, 1500, True, "fx_DragonRadar_1"
        PlaySound "vo_lock2"
        DragonRadar.image = "Radar_1"
    End If
    If BallsInLock(CurrentPlayer) = 3 Then
    bMinotaurMBStarted = True
        DragonRadar.image = "Radar_0"
    If bMedusaMBStarted = True OR bZeusMBStarted = True Or bKaiModeStarted = True Then  'Higher layer already on
      DMD CL("DOUBLE MULTIBALL"), CL("SHOOT THE HOLES"), "_", eBlink, eNone, eNone, 3000, True, "vo_kaiokenx2"
    Else If bCCorpMBStarted = True Then
      DMD CL("DOUBLE MULTIBALL"), CL("SHOOT THE HOLES"), "_", eBlink, eNone, eNone, 3000, True, "vo_kaiokenx2"
      StopSound "vo_AuraPoweredUp"    'Turn off lower layer first
      Goku_PoweredUp.visible = 0      '
      LightSeqGokuPU.StopPlay         '
      Goku_SS2.visible = 1          'Turn on correct layer
      PlaySound "vo_AuraSS2"          '
      LightSeqGokuSS2.UpdateInterval = 100 '
      LightSeqGokuSS2.Play SeqRandom, 50, , 10000000
      ChangeSong
    Else If CurrentMode(CurrentPlayer) = 8 OR CurrentMode(CurrentPlayer) = 9 Then  'Higher layer already on
      DMD CL("DRAGON RADAR MBALL"), CL("SHOOT THE HOLES"), "_", eNone, eNone, eNone, 2500, True, "vo_minotaur_multiball"
    Else
      DMD CL("DRAGON RADAR MBALL"), CL("SHOOT THE HOLES"), "_", eNone, eNone, eNone, 2500, True, "vo_minotaur_multiball"
      Goku_SS2.visible = 1
      PlaySound "vo_AuraSS2"
      LightSeqGokuSS2.UpdateInterval = 100
      LightSeqGokuSS2.Play SeqRandom, 50, , 10000000
      ChangeSong
    End If
    End If
    End If
        NeededScoopHits(CurrentPlayer) = NeededScoopHits(CurrentPlayer) + 1
        If NeededScoopHits(CurrentPlayer)> 5 Then NeededScoopHits(CurrentPlayer) = 5
        ScoopHits(CurrentPlayer) = 0
        bLockEnabled = False
        li039.State = 0
        Minotaur1 = 0
        Minotaur2 = 0
        Minotaur3 = 0
        AddMultiball 2
        BallsInLock(CurrentPlayer) = 0
        TrapdoorUp
        EnableBallSaver 15
    End If
End Sub

Sub CheckMinotaurMBHits
    If Minotaur1 + Minotaur2 + Minotaur3 = 3 Then 'all 3 holes has been hit so double the jackpot
        MinotaurJackpot(CurrentPlayer) = MinotaurJackpot(CurrentPlayer) * 2
        DMD CL("RADAR JACKPOT IS"), CL(FormatScore(MinotaurJackpot(CurrentPlayer) ) ), "_", eNone, eNone, eNone, 1500, True, ""
        Minotaur1 = 0
        Minotaur2 = 0
        Minotaur3 = 0
    End If
End Sub

Sub AwardMinotaurJackpot()
    DOF 130, DOFPulse
    DMD CL("DRAGON RADAR JACKPOT"), CL(FormatScore(MinotaurJackpot(CurrentPlayer) ) ), "_", eNone, eBlinkFast, eNone, 2000, True, "vo_Jackpot"
    DOF 126, DOFPulse
    AddScore2 Jackpot(CurrentPlayer)
    MinotaurJackpot(CurrentPlayer) = MinotaurJackpot(CurrentPlayer) + 100000
    LightEffect 2
    GiEffect 1
    FlashEffect 1
End Sub

'**********************
' GOKU/Zeus MB - Upper Ramp
'**********************

 Sub StartZeusMB
  If bMedusaMBStarted = True OR bMinotaurMBStarted = True OR bCCorpMBStarted = True OR bKaiModeStarted = True Then  'Multiball already running
    DMD RL("DOUBLE MULTIBALL"), RL("SHOOT UPPER RAMPS"), "d_zeus", eNone, eNone, eNone, 3000, True, "vo_kaiokenx10"
    Goku_PoweredUp.visible = 0        'turns off lower power levels first
    StopSound "vo_AuraPoweredUp"      '
    LightSeqGokuPU.StopPlay           '
    Goku_SS2.visible = 0              '
    StopSound "vo_AuraSS2"            '
    LightSeqGokuSS2.StopPlay          '
    Goku_Blue.visible = 0             '
    StopSound "vo_auraBlueWithVocals" '
    LightSeqGokuBlue.StopPlay         '
  Else
    DMD RL("GOKU MULTIBALL"), RL("SHOOT UPPER RAMPS"), "d_zeus", eNone, eNone, eNone, 2500, True, "vo_GokuScreamPowerUp"
  End If
  bZeusMBStarted = True
    ChangeSong
  li044.State = 2
  li042.State = 2
  AddMultiball 3
  EnableBallSaver 20
  ZeusMBFlashTimer.Enabled = 1
  GokuVisuals
End Sub

Sub ZeusMBFlashTimer_Timer
    LTF
    vpmtimer.addtimer 250, "RTF '" 'delay a little the right flasher
End Sub

'*******************
' BONUS MULTIPLIER
'*******************
' fire shots

Sub CheckBonusX
    If li030.State + li031.State + li032.State + li033.State + li034.State = 5 Then 'all the fire lights are on
        AddBonusMultiplier 1
        FlashEffect 5
        LightEffect 5
        li030.State = 0
        li031.State = 0
        li032.State = 0
        li033.State = 0
        li034.State = 0
        'blink the X lights fast but only the ones that were off
        If li024.State = 0 Then li024.State = 2
        If li025.State = 0 Then li025.State = 2
        If li026.State = 0 Then li026.State = 2
        If li027.State = 0 Then li027.State = 2
        If li028.State = 0 Then li028.State = 2
    End If
End Sub

'************
' HYPERBOLIC TIME CHAMBER/The Fates
'************

Sub CheckFates                                                                             'checks for hits and start mode
    If FatesHits(CurrentPlayer) MOD 10 = 0 Then
        DMD CL("HYPERBOLIC"), CL("TIME CHAMBER"), "_", eNone, eNone, eNone, 1500, True, "" 'vo_The_fates"
        bFatesStarted = True
    End If
End Sub

'*****************
' BONUS HIT SUBS
'*****************

Sub aBonusTargets_Hit(idx):BonusTargets(CurrentPlayer) = BonusTargets(CurrentPlayer) + 1:End Sub
Sub aBonusRamps_Hit(idx):BonusRamps(CurrentPlayer) = BonusRamps(CurrentPlayer) + 1:End Sub
Sub aBonusOrbits_Hit(idx):BonusOrbits(CurrentPlayer) = BonusOrbits(CurrentPlayer) + 1:End Sub


' Diverters, activated after 3 seconds flipper Update

Sub LeftdiverterTimer_Timer
    PlaySound "vo_KamehamehaL" &RndNbr(2)
    LeftdiverterTimer.Enabled = 0
    diverter2.Isdropped = 1
End Sub

Sub rightdiverterTimer_Timer
    PlaySound "vo_KamehamehaR" &RndNbr(2)
    rightdiverterTimer.Enabled = 0
    diverter1.Isdropped = 1
End Sub

'Change Ball types

Dim BallType: 'BallType = 1


Sub UpdateBallImage()
    Dim BOT, b
    BOT = GetBalls
    ' exit the Sub if no balls on the table
    If UBound(BOT) = 0 Then Exit Sub

    ' change the image for each ball
    For b = 1 to UBound(BOT)
        Select Case BallType
            Case 0
                If BallAlt = 0 Then
          BOT(b).FrontDecal = "JPBall-Scratches"
          BOT(b).Image = "JP Chrome_Ball"
        Else
          BOT(b).Image = "Ball_Back"
          BOT(b).FrontDecal = "BallOrangeStar"
        End If
        End Select
    Next
End Sub


' DMD CL(""), CL(""), "", eNone, eNone, eNone, 3000, True, ""

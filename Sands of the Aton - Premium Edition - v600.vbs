' ****************************************************************
'         Sands of the Aton - Premium Edition
'             for VISUAL PINBALL X 10.8
'         Uses FlexDMD for cabinet / FS mode
'       WPX, art, rules, music & sounds by dea TEE
'            WPX & script by jpsalas - 2024
' ****************************************************************

'DOF Config by Arngrim
'101 Left Flipper
'102 Right Flipper
'103 Left Slingshot
'104 Right Slingshot
'105 Not Used
'106 See Script
'107 See Script
'108 Bumper Left
'109 Bumper Right
'110 Bumper Center
'111 Not Used
'112 Not Used
'113 Not Used
'114 Spinner Right (Shaker)
'115 Spinner Left (Shaker)
'116 RightHole Kicker
'117 TopHole Kicker
'118 CenterHole Kicker
'119 Not Used
'120 AutoFire
'121 See Script
'122 Knocker
'123 Ball Release
'124 Not Used
'125 See Script
'126 Not Used
'127 See Script
'128 Not Used
'129 Not Used
'141 kicker
'142 RiseRightBank (Resetdrop)
'143 RiseTopBank (Resetdrop)
'144 RiseTombBank (Resetdrop)
'145 RiseLeftBank (Resetdrop)
'146 RiseGuardians (Resetdrop) X8
'147 RiseForceGuardians
'148 RiseAt 1,3,5,7
'149 RiseAt 2,4,6,8
'159 See Script

Option Explicit
Randomize

Const BallSize = 50 ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1  ' standard ball mass

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
Const cGameName = "SOTA"     'used for DOF and saving some values
Const myVersion = "2.00"
Const MaxPlayers = 4         ' from 1 to 4
Const MaxMultiplier = 9      ' limit playfield multiplier
Const MaxBonusMultiplier = 1 'limit Bonus multiplier
Const MaxMultiballs = 6      ' max number of balls during multiballs

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim PlayfieldMultiplier(4)
Dim BallSaverTime ' in seconds of the first ball
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot(4)
Dim SuperJackpot(4)
Dim LockedBalls
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim bAutoPlunger
Dim bInstantInfo
Dim bAttractMode
Dim LoopCount
Dim ComboCount
Dim ComboHits(4)
Dim ComboValue(4)
Dim x

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
Dim bMultiBallStarted
Dim bMusicOn
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim bJackpot

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger during multiballs
Dim LeftMagnet

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

    ' Left Magnet
    Set LeftMagnet = New cvpmMagnet
    With LeftMagnet
        .InitMagnet Magnet1, 20
        .GrabCenter = True
        .CreateEvents "LeftMagnet"
    End With

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' load saved values, highscore, names, jackpot
    Credits = 0
    Loadhs

    ' Initalise the DMD display
    DMD_Init

    ' Init main variables and any other flags
    bAttractMode = False
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    BallSaverTime = 10 'at the start and multiball
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bMultiBallStarted = False
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
    LockedBalls = 0
    ' set any lights for the attract mode
    vpmtimer.addtimer 2000, "GiOn '"
    StartAttractMode
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound "fx_nudge", 0, 1, 1, 0.25
    If ( Keycode = AddCreditKey OR Keycode = AddCreditKey2 ) AND NOT bFreePlay Then
        Credits = Credits + 1
        'if bFreePlay = False Then DOF 125, DOFOn
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

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then CheckTilt 'only check the tilt during game
        If keycode = RightTiltKey Then CheckTilt
        If keycode = CenterTiltKey Then CheckTilt
        If keycode = MechanicalTilt Then CheckTilt

        If keycode = LeftFlipperKey Then SolLFlipper 1:InstantInfoTimer.Enabled = True
        If keycode = RightFlipperKey Then SolRFlipper 1:InstantInfoTimer.Enabled = True

        If keycode = StartGameKey Then
            If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, True, ""
                Else
                    If(Credits> 0) Then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
                        DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, True, ""
                        If Credits <1 And bFreePlay = False Then DOF 125, DOFOff
                        Else
                            ' Not Enough Credits to start a game.
                            DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 1000, True, ""
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
                            If Credits <1 And bFreePlay = False Then DOF 125, DOFOff
                            ResetForNewGame()
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        DMDFlush
                        DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 1000, True, ""
                        ShowTableInfo
                    End If
                End If
            End If
    End If ' If (GameInPlay)
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If hsbModeActive Then
        Exit Sub
    End If

    If keycode = PlungerKey Then
        Plunger.Fire
        PlaySoundAt "fx_plunger", plunger
        DOF 159, DOFpulse
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
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBall "fx_rubber_flipper"
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBall "fx_rubber_flipper"
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

FlipperPower = 3600
FlipperElasticity = 0.6
FullStrokeEOS_Torque = 0.6 ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.3 ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.2
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

'********************
' Real Time updates
'   done by VPX8
'********************
'used for all the real time updates

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub
Sub at1f_Animate:at1p.Z = at1f.CurrentAngle:End Sub
Sub at2f_Animate:at2p.Z = at2f.CurrentAngle:End Sub
Sub at3f_Animate:at3p.Z = at3f.CurrentAngle:End Sub
Sub at4f_Animate:at4p.Z = at4f.CurrentAngle:End Sub
Sub at5f_Animate:at5p.Z = at5f.CurrentAngle:End Sub
Sub at6f_Animate:at6p.Z = at6f.CurrentAngle:End Sub
Sub at7f_Animate:at7p.Z = at7f.CurrentAngle:End Sub
Sub at8f_Animate:at8p.Z = at8f.CurrentAngle:End Sub
Sub PRIESTESSf_Animate:PRIESTESSp.Z = PRIESTESSf.Currentangle:End Sub
Sub PRIESTf_Animate:PRIESTp.Z = PRIESTf.Currentangle:End Sub
Sub Rad1F_Animate:Rad1P.Z = Rad1F.Currentangle:End Sub
Sub Rad2F_Animate:Rad2P.Z = Rad2F.Currentangle:End Sub

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
    If(NOT Tilted) AND Tilt> 15 Then 'If more that 15 Then TILT the table
        'display Tilt
        InstantInfoTimer.Enabled = False
        PlaySound "s_tilt"
        DMDFlush
        DMD CL("YOU"), CL("TILTED"), "", eNone, eNone, eNone, 3000, True, ""
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
        bMultiBallMode = False
        StopMBmodes
        BallSaverTimerExpired_Timer
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
        RightFlipper.RotateToStart
        Bumper001.Threshold = 100
        Bumper002.Threshold = 100
        Bumper003.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
        LeftMagnet.MagnetOn = False
    Else
        Tilted = False
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        Bumper001.Threshold = 1
        Bumper002.Threshold = 1
        Bumper003.Threshold = 1
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        DMDFlush
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained Then..
    If(BallsOnPlayfield = LockedBalls) Then
        bMultiBallMode = False
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        vpmtimer.Addtimer 2000, "EndOfBall() '"
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'*****************************************
'         Internal Music
'*****************************************

Dim Song
Song = ""

Sub PlaySong(name)
    If bMusicOn Then
        If Song <> name Then
            StopSound Song
            Song = name
            PlaySound Song, -1, SongVolume
        End If
    End If
End Sub

Sub ChangeSong
    If bRoamDesert Then
        PlaySong "s_roam the desert"
    ElseIf bAtonMB8 Then
        PlaySong "s_HELPME"
    ElseIf bAtonMB Then
        PlaySong "s_ATOMMUSICLOOP"
    Elseif bPriestMB Then
        PlaySong "s_priestessmcut"
    Elseif bPriestessMB Then
        PlaySong "s_priestessmcut"
    Elseif bGuardianMB Then
        PlaySong "s_guardianmusic"
    Else
        PlaySong "s_Music_Egyptian Loop"
    End If
End Sub

Sub StopSong
    StopSound Song
    Song = ""
End Sub

'**********************
'     GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'**********************

Dim GiIsOn
GiIsOn = False

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub ChangeGIIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGILights
        bulb.IntensityScale = factor
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) = -1 Then '-1 means no balls, 0 is the first captive ball, 1 is the second captive ball...)
        GiOff                ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
    Else
        Gion
    End If
End Sub

Sub GiOn
    Dim bulb
    If GiIsOn = False Then
        PlaySoundAt "fx_GiOn", GiRelay 'about the center of the table
        For each bulb in aGiLights
            bulb.State = 1
        Next
        GiIsOn = True
    End If
End Sub

Sub GiOff
    Dim bulb
    If GiIsOn = True Then
        PlaySoundAt "fx_GiOff", GiRelay 'about the center of the table
        For each bulb in aGiLights
            bulb.State = 0
        Next
        GiIsOn = False
    End If
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
            LightSeqGi.UpdateInterval = 40
            LightSeqGi.Play SeqBlinking, , 10, 20
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
            LightSeqInserts.UpdateInterval = 5
            LightSeqInserts.Play SeqCircleOutOn, 15, 1
        Case 5 'top down
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqDownOn, 15, 1
        Case 6 'down to top
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 15, 1
        Case 7 'random All Lights
            LightSeqAllLights.UpdateInterval = 25
            LightSeqAllLights.Play SeqRandom, 50, , 1000
        Case 8 'center in All Lights
            LightSeqAllLights.UpdateInterval = 5
            LightSeqAllLights.Play SeqCircleInOn, 15, 1
        Case 9 'all Priest lights blink
            LightSeqPriest.UpdateInterval = 30
            LightSeqPriest.Play SeqBlinking, , 15, 25
        Case 10 'random Priest lights
            LightSeqPriest.UpdateInterval = 25
            LightSeqPriest.Play SeqRandom, 50, , 1000
        Case 11 'all Priest lights at the end of Aton
            LightSeqPriest.UpdateInterval = 30
            LightSeqPriest.Play SeqBlinking, , 10, 25
    End Select
End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v3.0
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
    Pitch = BallVel(ball) * 200
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
    PlaySound soundname, 0, 1, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics, with a greater random pitch
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v3.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 19   'total number of balls, 20 balls, from 0 to 19
Const lob = 0     'number of locked balls
Const maxvel = 45 'max ball velocity
ReDim rolling(tnob)

InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
    RollingTimer.Enabled = 1
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls and hide the shadow
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 5
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
' Diverse Collection Hit Sounds v4.0
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

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
    Dim i

    bGameInPLay = True

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
        PlayfieldMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    Next

    ' initialise any other flags
    Tilt = 0

    ' initialise specific Game variables
    Game_Init()
    UpdateBallInPlay

    ' you may wish to start some music, play a sound, do whatever at this point
    PlaySound ""

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
    'DMDScoreNow

    ' set the current players bonus multiplier back down to 1X
    ' SetBonusMultiplier 1

    ' Set the playfield multiplier to 1
    SetPlayfieldMultiplier 1

    ' reset any drop targets, lights, game Mode etc..

    'BonusPoints(CurrentPlayer) = 0
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

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1
    UpdateBallInPlay

    ' kick it out..
    PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4

' if there is 2 or more balls Then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield - LockedBalls> 1 Then
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
            If mBalls2Eject = 0 Then 'if there are no more balls to eject Then stop the timer
                CreateMultiballTimer.Enabled = False
            End If
        Else 'the max number of multiballs is reached, so stop the timer
            mBalls2Eject = 0
            CreateMultiballTimer.Enabled = False
        End If
    End If
End Sub

Sub UpdateBallInPlay
    If B2SOn Then
        if BallsOnPlayfield = 0 Then
            Controller.B2sSetData 50, 0:Controller.B2sSetData 51, 0:Controller.B2sSetData 52, 0
        Else
            select case BallsRemaining(CurrentPlayer)
                Case 5:Controller.B2sSetData 50, 1:Controller.B2sSetData 51, 0:Controller.B2sSetData 52, 0
                Case 4:Controller.B2sSetData 50, 1:Controller.B2sSetData 51, 0:Controller.B2sSetData 52, 0
                Case 3:Controller.B2sSetData 50, 1:Controller.B2sSetData 51, 0:Controller.B2sSetData 52, 0
                Case 2:Controller.B2sSetData 50, 0:Controller.B2sSetData 51, 1:Controller.B2sSetData 52, 0
                Case 1:Controller.B2sSetData 50, 0:Controller.B2sSetData 51, 0:Controller.B2sSetData 52, 1
                Case 0:Controller.B2sSetData 50, 0:Controller.B2sSetData 51, 0:Controller.B2sSetData 52, 0
            end select
        end if
    End If
End Sub

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded

Sub EndOfBall()
    Dim TotalBonus
    ToTalBonus = 0
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False
    GiOff
    ' only process any of this if the table is not tilted.
    '(the tilt recovery mechanism will handle any extra balls or end of game)
    StopSong
    PlaySound "s_end of ball"

    If NOT Tilted Then
        'Add the bonus
        DMD CL("CURRENT AFTERLIFE"), CL(BonusPoints(CurrentPlayer) ), "", eNone, eNone, eNone, 1500, True, "s_AM_collectBonus"
        ToTalBonus = TotalBonus + BonusPoints(CurrentPlayer)

        DMD CL("LOOP BONUS"), "     2000 X " &LoopCount, "", eNone, eNone, eNone, 1500, True, "s_AM_collectBonus"
        ToTalBonus = TotalBonus + 2000 * LoopCount

        DMD CL("COMBO BONUS"), "     2000 X " & ComboCount, "", eNone, eNone, eNone, 1500, True, "s_AM_collectBonus"
        ToTalBonus = TotalBonus + 2000 * ComboCount

        DMD CL("TOTAL BONUS"), CL(TotalBonus), "", eNone, eBlink, eNone, 2000, True, "s_AM_collectBonus"
        Score(CurrentPlayer) = Score(CurrentPlayer) + TotalBonus

        ' add a bit of a delay to allow for the bonus points to be shown & added up
        vpmtimer.addtimer 6500, "EndOfBall2 '"
    Else 'if tilted Then only add a short delay and move to the 2nd part of the end of the ball
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, Then check to see if this was the last ball (of the CurrentPlayer)
'
Sub EndOfBall2()
    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

    ' Check for the last ball and the Afterlife extra Ball
    If BallsRemaining(CurrentPlayer) = 1 And bAfterLifeBallAwarded = False AND BonusPoints(CurrentPlayer) >= 6000 Then
        bAfterLifeBallAwarded = True
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        BonusPoints(CurrentPlayer) = 0 'consume afterlife bonus on the 4th ball
    End If
    ' has the player won an extra-ball ? (might be multiple outstanding)
    If ExtraBallsAwards(CurrentPlayer)> 0 Then
        'debug.print "Extra Ball"

        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1

        ' if no more EB's Then turn off any Extra Ball light if there was any
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
        ' LightShootAgain.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
        If bAfterLifeBallAwarded Then
            bAfterLifeBallAwarded = false
            DMD CL("4TH BALL-AFTERLIFE"), CL("SHOOT AGAIN"), "", eNone, eBlink, eNone, 1500, True, "s_4th ball afterlife"
        Else
            DMD CL("EXTRA BALL"), CL("SHOOT AGAIN"), "", eNone, eBlink, eNone, 1500, True, ""
        End If

        ' In this table an extra ball will have the skillshot and ball saver, so we reset the playfield for the new ball
        ResetForNewPlayerBall()

        ' Create a new ball in the shooters lane
        CreateNewBall()
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
            ' if multiple players are playing Then move onto the next one
            EndOfBallComplete()
        End If
    End If
End Sub

' This function is called when the end of bonus display
' (or high score entry finished) AND it either ends the game or
' move onto the next player (or the next ball of the same player)
'
Sub EndOfBallComplete()
    Dim NextPlayer

    'debug.print "EndOfBall - Complete"

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame> 1) Then
        ' Then move to the next player
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
                Case 1:DMD "", CL("PLAYER 1"), "", eNone, eNone, eNone, 1000, True, ""
                Case 2:DMD "", CL("PLAYER 2"), "", eNone, eNone, eNone, 1000, True, ""
                Case 3:DMD "", CL("PLAYER 3"), "", eNone, eNone, eNone, 1000, True, ""
                Case 4:DMD "", CL("PLAYER 4"), "", eNone, eNone, eNone, 1000, True, ""
            End Select
        Else
            DMD "", CL("PLAYER 1"), "", eNone, eNone, eNone, 1000, True, ""
        End If
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    'debug.print "End Of Game"
    bGameInPLay = False
    ' just ended your game Then play the end of game tune
    PlaySound "s_GameOver"
    ' ChangeSong
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all Mode - eject locked balls
    If LockedBalls Then
        EjectRightHole
        EjectTopHole
    End If
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
' if only one Then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) Then kill of the ball but don't create
' a new one
'
Sub Drain_Hit()
    ' Destroy the ball
    Drain.DestroyBall
    ' Exit Sub ' only for debugging - this way you can add balls from the debug window

    If BallsOnPlayfield> 0 Then
        BallsOnPlayfield = BallsOnPlayfield - 1
    End If

    ' pretend to knock the ball into the ball storage mech
    PlaySoundAt "fx_drain", Drain

    If bGameInPLay = False Then Exit Sub 'don't do anything, just delete the ball

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
            ' stop the ballsaver timer during the launch ball saver time, but not during multiballs
            If NOT bMultiBallMode Then
                DMD "_", CL("ETERNAL LIFE"), "_", eNone, eBlinkfast, eNone, 2500, True, ""
                BallSaverTimerExpired_Timer
            End If
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield - lockedballs = 1) Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True) Then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ' you may wish to change any music over at this point and
                    changesong
                    ' turn off any multiball specific lights
                    ChangeGIIntensity 1
                    'stop any multiball modes of this game
                    StopMBmodes
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield - LockedBalls = 0) Then
                ' End Mode and timers
                StopSong
                ChangeGIIntensity 1
                UpdateBallInPlay
                ' Show the end of ball animation
                ' and continue with the end of ball
                ' DMD something?
                StopEndOfBallMode
                vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, if there is no animation Then move to the end of ball
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    PlaySoundAt "fx_sensor", swPlungerRest
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
    'LaunchLight.State = 2
    ' be sure to update the Scoreboard after the animations, if any
    ' if the ball goes into the plunger lane during a multiball Then activate the autoplunger
    If bMultiBallMode Then
        bAutoPlunger = True ' kick the ball in play if the bAutoPlunger flag is on
    End If
    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        vpmtimer.addtimer 1500, "PlungerIM.AutoFire:DOF 120, DOFPulse:PlaySoundAt ""fx_kicker"", swPlungerRest:bAutoPlunger = False '"
    End If
    'Start the skillshot lights & variables if any
    ChangeSong
    If bSkillShotReady Then
        UpdateSkillshot()
        ' show the message to shoot the ball in case the player has fallen sleep
        swPlungerRest.TimerEnabled = 1
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    LightEffect 6
    bBallInPlungerLane = False
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
        ChangeSong
        ResetSkillShotTimer.Enabled = 1
    End If
    ' if there is a need for a ball saver, Then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    End If
' turn off LaunchLight
' LaunchLight.State = 0
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball after a while
Sub swPlungerRest_Timer
'PlaySound "wakeup" &RndNbr(6) 'there are only 3 sounds in the table, so it will play a sound about 50% of times
End Sub

Sub EnableBallSaver(seconds)
    ' do not start the timer if extra ball has been awarded
    'If ExtraBallsAwards'(CurrentPlayer) > 0 Then
    '    BallSaverTimerExpired.Enabled = False
    '    BallSaverSpeedUpTimer.Enabled = False
    '    LightShootAgain.State = 1
    '    Exit Sub
    'End If
    'debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' stop the ball saver grace period if it was started
    BallSaverGracePeriod.Enabled = 0
    ' start the timers
    BallSaverTimerExpired.Interval = 1000 * seconds
    BallSaverTimerExpired.Enabled = False
    BallSaverTimerExpired.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * (seconds) - 3000 'the last 3 seconds will the light blink faster
    BallSaverSpeedUpTimer.Enabled = False
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.State = 0
    LightShootAgain.State = 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimerExpired_Timer()
    'debug.print "Ballsaver ended"
    BallSaverTimerExpired.Enabled = False
    BallSaverSpeedUpTimer.Enabled = False 'ensure this timer is also stopped
    ' clear the flag after a few seconds of grace period
    BallSaverGracePeriod.Interval = 1500
    BallSaverGracePeriod.Enabled = 0 'restart the grace period timer
    BallSaverGracePeriod.Enabled = 1
    ' if you have a ball saver light Then turn it off at this point
    LightShootAgain.State = 0
' if the table uses the same lights for the extra ball or replay Then turn them on if needed
' If ExtraBallsAwards(CurrentPlayer) > 0 Then
' LightShootAgain.State = 1
' End If
End Sub

Sub BallSaverGracePeriod_Timer 'grace period 1,5 seconds
    bBallSaverActive = False '"
    BallSaverGracePeriod.Enabled = 0
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

Sub AddScore(points) 'normal score routine; points x playfieldmultiplier
    If Tilted Then Exit Sub
    If bSkillshotReady AND(Score(CurrentPlayer)> 0) Then ResetSkillShotTimer_Timer
    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer)
' you may wish to check to see if the player has gotten a replay
End Sub

Sub AddBonus(points)
    If Tilted Then Exit Sub
    ' add the bonus to the current players bonus variable
    BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
End Sub

' Set the Playfield Multiplier to the specified level AND set any lights accordingly

Sub SetPlayfieldMultiplier(Level)
    ' Set the multiplier to the specified level
    PlayfieldMultiplier(CurrentPlayer) = Level
    UpdatePFXLights(Level)
End Sub

Sub UpdatePFXLights(Level)
    ' Update the playfield multiplier lights
    Select Case Level
        Case 1:light022.State = 0:light027.State = 0:light023.State = 0
        Case 3:light022.State = 1:light027.State = 0:light023.State = 0
        Case 6:light022.State = 0:light027.State = 1:light023.State = 0
        Case 9:light022.State = 0:light027.State = 0:light023.State = 1
    End Select
' perhaps show also the multiplier in the DMD?
End Sub

Sub AwardExtraBall()
    '   If NOT bExtraBallWonThisBall Then 'uncomment this If in case you want to give just one extra ball per ball
    DMD "_", " EXTRA BALL AWARDED", "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
    DOF 121, DOFPulse
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    '    bExtraBallWonThisBall = True
    'LightShootAgain.State = 1 'light the shoot again lamp
    GiEffect 3
    LightEffect 2
'    END If
End Sub

Sub AwardSkillshot(points)
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL("SKILLSHOT"), CL(FormatScore(points) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "s_AM_SkillShot"
    DOF 127, DOFPulse
    AddScore points
    'do some light show
    GiEffect 3
    LightEffect 2
End Sub

Sub AwardSuperSkillshot(points)
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL("SUPER SKILLSHOT"), CL(FormatScore(points) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "s_SuperSKillShot"
    DOF 127, DOFPulse
    AddScore points
    'do some light show
    GiEffect 3
    LightEffect 2
End Sub

Sub AwardSuperJackpot
    'show dmd animation
    DMD CL("SUPER JACKPOT"), CL("500.000"), "d_border", eNone, eBlinkFast, eNone, 2000, True, "s_SuperJackpot"
    DOF 127, DOFPulse
    AddScore 500000
    'do some light show
    GiEffect 3
    LightEffect 2
End Sub

'*****************************
'    Load / Save / Highscore
'*****************************

Dim MyTable
MyTable = "SOTA"

Sub Loadhs
    Dim x
    x = LoadValue(MyTable, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 100000 End If
    x = LoadValue(MyTable, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If
    x = LoadValue(MyTable, "HighScore2")
    If(x <> "") Then HighScore(1) = CDbl(x) Else HighScore(1) = 100000 End If
    x = LoadValue(MyTable, "HighScore2Name")
    If(x <> "") Then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If
    x = LoadValue(MyTable, "HighScore3")
    If(x <> "") Then HighScore(2) = CDbl(x) Else HighScore(2) = 100000 End If
    x = LoadValue(MyTable, "HighScore3Name")
    If(x <> "") Then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If
    x = LoadValue(MyTable, "HighScore4")
    If(x <> "") Then HighScore(3) = CDbl(x) Else HighScore(3) = 100000 End If
    x = LoadValue(MyTable, "HighScore4Name")
    If(x <> "") Then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If
    x = LoadValue(MyTable, "Credits")
    If(x <> "") Then Credits = CInt(x) Else Credits = 0:If bFreePlay = False Then DOF 125, DOFOff:End If
    x = LoadValue(MyTable, "TotalGamesPlayed")
    If(x <> "") Then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If
    SortHighscore 'to be sure they are in the right order.
End Sub

Sub Savehs
    SaveValue MyTable, "HighScore1", HighScore(0)
    SaveValue MyTable, "HighScore1Name", HighScoreName(0)
    SaveValue MyTable, "HighScore2", HighScore(1)
    SaveValue MyTable, "HighScore2Name", HighScoreName(1)
    SaveValue MyTable, "HighScore3", HighScore(2)
    SaveValue MyTable, "HighScore3Name", HighScoreName(2)
    SaveValue MyTable, "HighScore4", HighScore(3)
    SaveValue MyTable, "HighScore4Name", HighScoreName(3)
    SaveValue MyTable, "Credits", Credits
    SaveValue MyTable, "TotalGamesPlayed", TotalGamesPlayed
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
    'DOF 125, DOFOn
    End If

    If tmp> HighScore(3) Then
        PlaySound SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
        DOF 121, DOFPulse
        HighScore(3) = tmp
        'enter player's name
        HighScoreEntryInit()
    Else
        EndOfBallComplete()
    End If
End Sub

Sub HighScoreEntryInit()
    hsbModeActive = True
    PlaySound "" &RndNbr(6)
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
        if(hsCurrentLetter = 0) Then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters) ) Then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<") Then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) Then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit> 0) Then
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
    if(hsCurrentDigit> 0) Then TempBotStr = TempBotStr & hsEnteredDigits(0)
    if(hsCurrentDigit> 1) Then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit> 2) Then TempBotStr = TempBotStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3) Then
        if(hsLetterFlash <> 0) Then
            TempBotStr = TempBotStr & "_"
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit <1) Then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit <2) Then TempBotStr = TempBotStr & hsEnteredDigits(2)

    TempBotStr = TempBotStr & " <    "
    dLine(1) = ExpandLine(TempBotStr)
    DMDUpdate 1
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2) Then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = False

    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ") Then
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
    savehs 'save the highscore in case the table is forced quit or it crashes.
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
    Else
        digitgrid.Visible = True
        For i = 0 to 40
            Digits(i).Visible = True
        Next
    End If

    Dim i, j
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
    'info on the second line: tmp1
    'Select Case Mode(CurrentPlayer, 0)
    'Case 0 'no Mode active
    'Case 1
    '    tmp = FL("PLAYER " &CurrentPlayer, FormatScore(Score(Currentplayer)))
    '    tmp1 = CL("SPINNERS LEFT " & SpinsToCount1-SpinCount1)
    'End Select
    End If
    DMD tmp, tmp1, tmp2, eNone, eNone, eNone, 10, True, ""
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
    DMDEffectTimer.Interval = deSpeed
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
                        If i = 2 Then
                            Temp = ""
                        End If
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
                        If i = 2 Then
                            Temp = ""
                        End If
                    End If
                case eLongScrollLeft:
                    Temp = Right(dLine(i), 19)
                    Temp = Temp & Mid(dqText(i, dqHead), deCount(i), 1)
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

Function ExpandLine(TempStr)
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
        if IsNumeric(mid(NumString, i, 1) ) Then
            NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1) ) + 128) & right(NumString, Len(NumString) - i)
        end if
    Next
    FormatScore = NumString
End function

Function FL(NumString1, NumString2) 'Fill line
    Dim Temp, TempStr
    If Len(NumString1) + Len(NumString2) <20 Then
        Temp = 20 - Len(NumString1) - Len(NumString2)
        TempStr = NumString1 & Space(Temp) & NumString2
        FL = TempStr
    End If
End Function

Function CL(NumString) 'center line
    Dim Temp, TempStr
    If Len(NumString)> 20 Then NumString = Left(NumString, 20)
    Temp = (20 - Len(NumString) ) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(NumString) 'right line
    Dim Temp, TempStr
    If Len(NumString)> 20 Then NumString = Left(NumString, 20)
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

'************************************
'    JP's new DMD using flashers
' two text lines and 1 backdrop image
'************************************

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
    'Chars(33) = ""        '!
    'Chars(34) = ""        '"
    'Chars(35) = ""        '#
    'Chars(36) = ""        '$
    'Chars(37) = ""        '%
    'Chars(38) = ""        '&
    'Chars(39) = ""        ''
    'Chars(40) = ""        '(
    'Chars(41) = ""        ')
    Chars(42) = "d_star" '*
    Chars(43) = "d_plus" '+
    'Chars(44) = ""        '
    Chars(45) = "d_minus" '-
    Chars(46) = "d_dot"   '.
    'Chars(47) = ""        '/
    Chars(48) = "d_0"    '0
    Chars(49) = "d_1"    '1
    Chars(50) = "d_2"    '2
    Chars(51) = "d_3"    '3
    Chars(52) = "d_4"    '4
    Chars(53) = "d_5"    '5
    Chars(54) = "d_6"    '6
    Chars(55) = "d_7"    '7
    Chars(56) = "d_8"    '8
    Chars(57) = "d_9"    '9
    Chars(60) = "d_less" '<
    'Chars(61) = ""        '=
    Chars(62) = "d_more" '>
    'Chars(64) = ""        '@
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
    Chars(95) = "d_under" '_
    'used in the FormatScore function
    Chars(176) = "d_0a" '0.
    Chars(177) = "d_1a" '1.
    Chars(178) = "d_2a" '2.
    Chars(179) = "d_3a" '3.
    Chars(180) = "d_4a" '4.
    Chars(181) = "d_5a" '5.
    Chars(182) = "d_6a" '6.
    Chars(183) = "d_7a" '7.
    Chars(184) = "d_8a" '8.
    Chars(185) = "d_9a" '9.
End Sub

'********************************************************************************************
' FlashForMs VPX8 will blink light for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Blink, -1 keep the original state
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version
    If FinalState = -1 Then
        FinalState = MyLight.State                            'Keep the current light state
    End If
    MyLight.BlinkInterval = BlinkPeriod
    MyLight.Duration 2, TotalPeriod, FinalState
End Sub

'******************************************
' Change light color - simulate color leds
' changes the light color and state
' 11 colors: red, orange, amber, yellow...
'******************************************

'colors
Const red = 5
Const orange = 4
Const amber = 6
Const yellow = 3
Const darkgreen = 7
Const green = 2
Const blue = 1
Const darkblue = 8
Const purple = 9
Const white = 11
Const teal = 10

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
        Case white
            n.color = RGB(193, 91, 0)
            n.colorfull = RGB(255, 197, 143)
        Case teal
            n.color = RGB(1, 64, 62)
            n.colorfull = RGB(2, 128, 126)
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
        Case white
            n.color = RGB(255, 197, 143)
        Case teal
            n.color = RGB(2, 128, 126)
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
            If rGreen> 255 Then
                rGreen = 255
                RGBStep = 1
            End If
        Case 1 'Red
            rRed = rRed - RGBFactor
            If rRed <0 Then
                rRed = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            rBlue = rBlue + RGBFactor
            If rBlue> 255 Then
                rBlue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            rGreen = rGreen - RGBFactor
            If rGreen <0 Then
                rGreen = 0
                RGBStep = 4
            End If
        Case 4 'Red
            rRed = rRed + RGBFactor
            If rRed> 255 Then
                rRed = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            rBlue = rBlue - RGBFactor
            If rBlue <0 Then
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
    DMD CL("DEA TEE AND"), CL("JPSALAS"), "", eNone, eNone, eNone, 3000, False, ""
    DMD "", CL("PRESENTS"), "", eNone, eNone, eNone, 3000, False, ""
    DMD "", "", "d_title", eNone, eNone, eNone, 4000, False, ""
    DMD CL("PREMIUM EDITION"), CL("ROM VERSION " &myversion), "", eNone, eNone, eNone, 2000, False, ""
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
    PlaySong "s_Music_Idle TableSOTA"
End Sub

Sub StopAttractMode
    DMDScoreNow
    LightSeqAttract.StopPlay
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqAttract.UpdateInterval = 40
    LightSeqAttract.Play SeqBlinking, , 5, 150
    LightSeqAttract.Play SeqRandom, 40, , 6000
    LightSeqAttract.Play SeqAllOff
    LightSeqAttract.UpdateInterval = 15
    LightSeqAttract.Play SeqUpOn, 50, 1
    LightSeqAttract.UpdateInterval = 15
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 15
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 15
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 15
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 15
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 15
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 15
    LightSeqAttract.Play SeqLeftOn, 50, 1
End Sub

Sub LightSeqAttract_PlayDone()
    StartLightSeq()
End Sub

Sub LightSeqTilt_PlayDone()
    LightSeqTilt.Play SeqAllOff
End Sub

Sub LightSeqTopFlashers_PlayDone()
    FlashEffect 7
End Sub

'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************

' droptargets, animations, timers, etc
Sub VPObjects_Init
    WallHelp.IsDropped = 1
End Sub

' tables variables and Mode init
Dim BumperHits(4)
Dim BumperValue(4)
Dim BumperText(4)
Dim BallinTopHole
Dim BallInRightHole
Dim AerialValue(4)
Dim LidalValue(4)
Dim bPriestMBReady
Dim bPriestessMBReady
Dim bGuardianMBReady
Dim bAtonMBReady
Dim bPriestMB
Dim bPriestessMB
Dim bGuardianMB
Dim bAtonMB
Dim bAtonMB8 'last target - added for change of the music
Dim GoldHits(4, 4)  '4 players, 4 targets
Dim EgyptHits(4, 5) '4 players, 5 targets
Dim EgyptCount
Dim EgyptLetters
Dim PriestHits(4)
Dim PriestessHits(4)
Dim bLeftCheatDeath
Dim bRightCheatDeath
Dim bNorthLock 'during AtonMB
Dim bEastLock  'during AtonLock
Dim bRoamDesert
Dim bRoamDesert2
Dim bAfterLifeBallAwarded
Dim HiddenPassHits

' Modes variables

Sub Game_Init() 'called at the start of a new game
    TurnOffPlayfieldLights()
    DropAllAT_Solenoid
    DropPRIESTESS_Solenoid
    DropPriest
    DropGuardians
    DropForceGuardians_Solenoid
    DropRadiTargets
    RiseTopBank
    RiseTombBank
    'Reset variables
    For x = 1 to 4
        BumperHits(x) = 0
        BumperValue(x) = 1000
        BumperText(x) = "  BUMPER HIT"
        AerialValue(x) = 0
        LidalValue(x) = 0
        PriestHits(x) = 0
        PriestessHits(x) = 0
        GoldHits(x, 1) = 0
        GoldHits(x, 2) = 0
        GoldHits(x, 3) = 0
        GoldHits(x, 4) = 0
        EgyptHits(x, 1) = 0
        EgyptHits(x, 2) = 0
        EgyptHits(x, 3) = 0
        EgyptHits(x, 4) = 0
        EgyptHits(x, 5) = 0
    Next
    BallinTopHole = False
    BallInRightHole = False
    bPriestMBReady = False
    bPriestessMBReady = False
    bGuardianMBReady = False
    bAtonMBReady = False
    bPriestMB = False
    bPriestessMB = False
    bGuardianMB = False
    bAtonMB = False
    bAtonMB8 = False
    LoopCount = 0
    ComboCount = 0
    EgyptCount = 0
    bLeftCheatDeath = False
    bRightCheatDeath = False
    bNorthLock = False
    bEastLock = False
    bRoamDesert = False
    bRoamDesert2 = False
    bAfterLifeBallAwarded = False
    HiddenPassHits = 0
    ' play a welcome Sound
    StopSound "s_GameOver"
    'PLaySound "Start" &RndNbr(10)
End Sub

Sub InstantInfo
    Dim tmp
    DMD CL("INSTANT INFO"), "", "", eNone, eNone, eNone, 1000, False, ""
    'Show some info on the current Mode

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
    DMD CL("BUMPER HITS"), CL(BumperHits(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("EXTRA BALLS"), CL(ExtraBallsAwards(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("CURRENT AFTERLIFE"), CL(BonusPoints(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
End Sub

Sub StopMBmodes 'stop multiball modes after loosing the last multiball
    If bGuardianMB Then StopGuardianMB
    If bPriestMB AND bRoamDesert2= False Then StopPriestMB
    If bPriestessMB AND bRoamDesert2= False Then StopPriestessMB
    If bRoamDesert Then RoamDesertTimer_timer 'to stop it.
    If bAtonMB Then StopAtonMB
    CheckCobraLight
    ATONLight.State = 0 'to be usre this light is off
End Sub

Sub StopEndOfBallMode()                       'this sub is called after the last ball in play is drained, reset skillshot, modes, timers
    If bRoamDesert Then RoamDesertTimer_timer 'to stop it.
    ResetSkillShotTimer_Timer                 'it should never be needed, but just to be sure
End Sub

Sub ResetNewBallVariables()                   'reset variables and lights for a new ball or player
    'turn on or off the needed lights before a new ball is released
    TurnOffPlayfieldLights
    If BumperHits(CurrentPlayer)> 250 Then 'reset the bumpers value
        BumperHits(CurrentPlayer) = 0
        BumperValue(CurrentPlayer) = 1000
        BumperText(CurrentPlayer) = "  BUMPER HIT"
    End If
    bPriestMBReady = False
    bPriestessMBReady = False
    bGuardianMBReady = False
    If bAtonMBReady Then Light004.State = 2 'players can steal the Aton Multiball from other players.
    LoopCount = 0
    ComboCount = 0
    DropGuardians
    'set up the lights/targets according to the player achievments
    If EgyptHits(currentPlayer, 0) = 5 Then
        vpmtimer.Addtimer 1200, "RiseRightBank '"
    Else
        UpdateRightBank
    End If
    If GoldHits(currentPlayer, 0) = 4 Then
        vpmTimer.AddTimer 1200, "RiseLeftBank '"
    Else
        UpdateLeftBank
    End If
    UpdateCURSE
    UpdateWOKEN
    SetPlayfieldMultiplier 1:EgyptCount = 0
    bLeftCheatDeath = False
    bRightCheatDeath = False
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub UpdateSkillShot() 'Setup and updates the skillshot lights
    DMD CL("HIT LIT BUMPER"), CL("FOR SKILLSHOT"), "", eNone, eNone, eNone, 3000, True, ""
    Select Case RndNbr(3)
        Case 1:Lbumper1a.State = 2:Lbumper1b.State = 2:Lbumper2a.State = 0:Lbumper2b.State = 0:Lbumper3a.State = 0:Lbumper3b.State = 0
        Case 2:Lbumper1a.State = 0:Lbumper1b.State = 0:Lbumper2a.State = 2:Lbumper2b.State = 2:Lbumper3a.State = 0:Lbumper3b.State = 0
        Case 3:Lbumper1a.State = 0:Lbumper1b.State = 0:Lbumper2a.State = 0:Lbumper2b.State = 0:Lbumper3a.State = 2:Lbumper3b.State = 2
    End Select
    Light029.State = 2
    Light026.State = 2
    Light025.State = 2
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    Lbumper1a.State = 1
    Lbumper1b.State = 1
    Lbumper2a.State = 1
    Lbumper2b.State = 1
    Lbumper3a.State = 1
    Lbumper3b.State = 1
    Light029.State = 0
    Light026.State = 0
    Light025.State = 0
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

'*********************************************************
' Slingshots has been hit

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFcontactors), Lemk
    DOF 106, DOFPulse 'DOF Solenoid/MX
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 50
    ' check modes
    ' add some effect to the table?
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
    DOF 107, DOFPulse 'DOF Solenoid/MX
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 50
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

'***********************
'        Bumpers
'***********************

Sub Bumper001_Hit
    If Tilted Then Exit Sub
    If bSkillshotReady AND Lbumper1a.State = 2 Then
        AwardSkillshot 25000
    ElseIf bSkillShotReady Then ResetSkillShotTimer_Timer
    End If
    PlaySoundAt SoundFXDOF("fx_bumper", 108, DOFPulse, DOFContactors), Bumper001
    CheckBumperHits
' check for modes
' remember last trigger hit by the ball
'LastSwitchHit = "Bumper"
End Sub

Sub Bumper002_Hit
    If Tilted Then Exit Sub
    If bSkillshotReady AND Lbumper2a.State = 2 Then
        AwardSkillshot 25000
    ElseIf bSkillShotReady Then ResetSkillShotTimer_Timer
    End If
    PlaySoundAt SoundFXDOF("fx_bumper", 109, DOFPulse, DOFContactors), Bumper002
    CheckBumperHits
' check for modes
' remember last trigger hit by the ball
'LastSwitchHit = "Bumper"
End Sub

Sub Bumper003_Hit
    If Tilted Then Exit Sub
    If bSkillshotReady AND Lbumper3a.State = 2 Then
        AwardSkillshot 25000
    ElseIf bSkillShotReady Then ResetSkillShotTimer_Timer
    End If
    PlaySoundAt SoundFXDOF("fx_bumper", 110, DOFPulse, DOFContactors), Bumper003
    CheckBumperHits
' check for modes
' remember last trigger hit by the ball
'LastSwitchHit = "Bumper"
End Sub

Sub CheckBumperHits 'addscore and increases value sccording to number of hits
    BumperHits(CurrentPlayer) = BumperHits(CurrentPlayer) + 1
    If BumperHits(CurrentPlayer) <51 Then BumperValue(CurrentPlayer) = 1000:BumperText(CurrentPlayer) = "BUMPER HIT"
    If BumperHits(CurrentPlayer)> 50 Then BumperValue(CurrentPlayer) = 2000:BumperText(CurrentPlayer) = "RELIC BUMPERS"
    If BumperHits(CurrentPlayer)> 100 Then BumperValue(CurrentPlayer) = 3000:BumperText(CurrentPlayer) = "COSMIC BUMPERS"
    If BumperHits(CurrentPlayer)> 150 Then BumperValue(CurrentPlayer) = 5000:BumperText(CurrentPlayer) = "ATON BUMPERS"
    If BumperHits(CurrentPlayer)> 250 Then BumperValue(CurrentPlayer) = 10000:BumperText(CurrentPlayer) = "SUPER ATON BUMPERS"
    If NOT bMultiBallMode Then
        DMD BumperValue(CurrentPlayer), BumperText(CurrentPlayer) & " " &BumperHits(CurrentPlayer), "_", eNone, eNone, eNone, 200, True, ""
    End If
    AddScore BumperValue(CurrentPlayer)
End Sub

''*********
' Lanes
'*********
' in and outlanes
Sub LeftOutlane_Hit
    PlaySoundAt "fx_sensor", LeftOutlane
    If Tilted Then Exit Sub
    Addscore 50
    ' check modes
    if Light015.State Then AddBonus 2000:Light015.State = 0
    bLeftCheatDeath = False
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftOutlane"
End Sub

Sub LeftInlane_Hit
    PlaySoundAt "fx_sensor", LeftInlane
    If Tilted Then Exit Sub
    Addscore 100
    ' Cheat Death Loops
    'enable it if the ball goes upwards and disable it if it goes downwards
    If ActiveBall.VelY <0 Then 'going up
        bLeftCheatDeath = True
    Else
        bLeftCheatDeath = False
    End If
    ' only remember the switch name if the ball is not coming down from the ramp, due to combos
    If LastSwitchHit <> "LidarTrigger" Then LastSwitchHit = "LeftInlane"
End Sub

Sub RightInlane_Hit
    PlaySoundAt "fx_sensor", RightInlane
    If Tilted Then Exit Sub
    Addscore 100
    ' Cheat Death Loops
    'enable it if the ball goes upwards and disable it if it goes downwards
    If ActiveBall.VelY <0 Then 'going up
        bRightCheatDeath = True
    Else
        bRightCheatDeath = False
    End If
    ' only remember the switch name if the ball is not coming down from the ramp, due to combos
    If LastSwitchHit <> "AerialTrigger" Then LastSwitchHit = "RightInlane"
End Sub

Sub RightOutlane_Hit
    PlaySoundAt "fx_sensor", RightOutlane
    If Tilted Then Exit Sub
    Addscore 50
    ' check modes
    if Light014.State Then AddBonus 2000:Light014.State = 0
    bRightCheatDeath = False
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightOutlane"
End Sub

'other lanes
Sub Trigger001_Hit 'Serpent Delta: top left lane
    PlaySoundAt "fx_sensor", Trigger001
    If Tilted Then Exit Sub
    Addscore 1500
    'If bMultiBallMode Then
        PlaySound "s_AM_serpent delta"
    'Else
     '   DMD "_", CL("SERPENT DELTA"), "_", eNone, eBlink, eNone, 1000, True, "s_AM_serpent delta"
    'End If
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger001"
End Sub

Sub Trigger003_Hit 'right top gate
    PlaySoundAt "fx_sensor", Trigger003
    If Tilted Then Exit Sub
    ' check modes
    If Score(CurrentPlayer) = 0 Then 'this is the first ball of the PlayersPlayingGame
        AddScore 500
        DMD "", " WELLCOME TO EGYPT", "", eNone, eNone, eNone, 1000, True, "s_AirplaneEntrance"
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger003"
End Sub

Dim SecretHit
SecretHit = False

Sub Trigger004_Hit 'Secret Passage: hidden gate
    PlaySoundAt "fx_sensor", Trigger004
    If Tilted Then Exit Sub
    If SecretHit = False Then
        SecretHit = True
        Addscore 50000
        If bMultiBallMode Then
            PlaySound "s_SecretPass"
        Else
            DMD CL("50.000"), CL("SECRET PASSAGE"), "", eBlink, eBlink, eNone, 1500, True, "s_SecretPass"
        End If
        vpmTimer.AddTimer 2000, "SecretHit = False '"
    End If
    ' check modes
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger004"
End Sub

' Shooting Stars: Top left 11 triggers

Sub aShootingStars_Hit(idx)
    If Tilted Then Exit Sub
    Addscore 100
    aShootingLights(idx).Duration 1, 300, 0
    DMD "_", CL("SHOOTING STAR"), "", eNone, eBlinkFast, eNone, 100, True, ""
End Sub

' Columns

Sub dp1_Hit
    PlaysoundAtBall "fx_woodhit"
    If Tilted Then Exit Sub
    If bMultiBallMode Then
        PlaySound "s_sci-fi-insect-plague"
    Else
        DMD CL("1.250"), CL("DESERT PILLARS"), "", eBlink, eNone, eNone, 1500, True, "s_sci-fi-insect-plague"
    End If
    FlashForMs dpl1, 1000, 50, 0
    Addscore 1250
End Sub

Sub dp2_Hit
    PlaysoundAtBall "fx_woodhit"
    If Tilted Then Exit Sub
    'If bMultiBallMode Then
        PlaySound "s_sci-fi-insect-plague"
    'Else
    '    DMD CL("1.250"), CL("DESERT PILLARS"), "", eBlink, eNone, eNone, 1500, True, "s_sci-fi-insect-plague"
    'End If
    FlashForMs dpl1, 1000, 50, 0
    Addscore 1250
End Sub

'***********
' Targets
'***********
' left and right sides

Sub Target003_Hit 'right red target
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 100
    FlashForMs f2, 1000, 50, 0
    ' check modes
    Light014.State = 1
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target003"
End Sub

Sub Target004_Hit 'left red target
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 100
    FlashForMs f1, 1000, 50, 0
    ' check modes
    Light015.State = 1
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target004"
End Sub

Sub Target005_Hit 'left white target
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 100
    FlashForMs f3, 1000, 50, 0
    If bBallSaverActive = False AND bRoamDesert2 = False Then EnableBallSaver 5
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "Target004"
End Sub

Sub Target006_Hit 'right white target
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 100
    FlashForMs f4, 1000, 50, 0
    If bBallSaverActive = False AND bRoamDesert2 = False Then EnableBallSaver 5
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "Target006"
End Sub

' Red targets

Sub TargetRed1_Hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 50
    'Rise South East Guardian
    If GuardianTarget4.IsDropped Then
        GuardianTarget4.IsDropped = 0
        If NOT bGuardianMB Then GuardianTarget4.TimerEnabled = 1
        PlaySoundAt SoundFXDOF("fx_resetdrop", 146, DOFPulse, DOFContactors), GuardianTarget4
        GTL4.State = 2
        If NOT bMultiBallMode Then
            DMD "_", CL("ALARM TARGET"), "_", eNone, eBlink, eNone, 1000, True, ""
        End If
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "TargetBlue1"
End Sub

Sub TargetRed2_Hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 50
    'Rises the South West Guardian Target
    If GuardianTarget2.IsDropped Then
        GuardianTarget2.IsDropped = 0
        If NOT bGuardianMB Then GuardianTarget2.TimerEnabled = 1
        PlaySoundAt SoundFXDOF("fx_resetdrop", 146, DOFPulse, DOFContactors), GuardianTarget4
        GTL2.State = 2
        If NOT bMultiBallMode Then
            DMD "_", CL("ALARM TARGET"), "_", eNone, eBlink, eNone, 1000, True, ""
        End If
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "TargetBlue2"
End Sub

Sub TargetRed3_Hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 50
    'Rises the Center Guardian Target
    If GuardianTarget3.IsDropped Then
        GuardianTarget3.IsDropped = 0
        If NOT bGuardianMB Then GuardianTarget3.TimerEnabled = 1
        PlaySoundAt SoundFXDOF("fx_resetdrop", 146, DOFPulse, DOFContactors), GuardianTarget4
        GTL3.State = 2
        If NOT bMultiBallMode Then
            DMD "_", CL("ALARM TARGET"), "_", eNone, eBlink, eNone, 1000, True, ""
        End If
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "TargetBlue3"
End Sub

'***********
'  Spinner
'***********

Sub Spinner001_Spin
    If Tilted Then Exit Sub
    PlaySoundAt "fx_spinner", spinner001
    DOF 114, DOFpulse
    Addscore 500
End Sub

Sub Spinner002_Spin
    If Tilted Then Exit Sub
    PlaySoundAt "fx_spinner", spinner002
    DOF 115, DOFpulse
    Addscore 500
End Sub

' DropTargets

' Top Bank: 3 droptargets
Sub Tdt1_Hit
    PlaySoundAt "fx_droptarget", TopHole
    If Tilted Then Exit Sub
    Addscore 2000
    Tdt1.IsDropped = 1
    If bMultiBallMode Then
        PlaySound "s_AM_StoneDoor"
    Else
        DMD "_", CL("TOMB DOOR"), "_", eNone, eNone, eNone, 1000, True, "s_AM_StoneDoor"
    End If
    ATONLight.duration 2, 4500, 0
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "Tdt1"
End Sub

Sub Tdt2_Hit
    PlaySoundAt "fx_droptarget", TopHole
    If Tilted Then Exit Sub
    Addscore 3000
    Tdt2.IsDropped = 1
    If bMultiBallMode Then
        PlaySound "s_AM_StoneDoor"
    Else
        DMD "_", CL("TOMB DOOR"), "_", eNone, eNone, eNone, 1000, True, "s_AM_StoneDoor"
    End If
    ATONLight.duration 2, 4500, 0
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "Tdt2"
End Sub

Sub Tdt3_Hit
    PlaySoundAt "fx_droptarget", TopHole
    If Tilted Then Exit Sub
    Addscore 5000
    If NOT BallinTopHole Then
        Tdt3.IsDropped = 1
        If bMultiBallMode Then
            PlaySound "s_AM_StoneDoor"
        Else
            DMD "_", CL("TOMB DOOR"), "_", eNone, eNone, eNone, 1000, True, "s_AM_StoneDoor"
        End If
        ATONLight.duration 2, 4500, 0
    End If
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "Tdt3"
End Sub

Sub RiseTopBank
    PlaySoundAt SoundFXDOF("fx_resetdrop", 143, DOFPulse, DOFContactors), TopHole
    Tdt1.IsDropped = 0
    Tdt2.IsDropped = 0
    Tdt3.IsDropped = 0
End SUb

'Left Bank: 4 droptargets
Sub Ldt1_Hit
    PlaySoundAt "fx_droptarget", Ldt1
    If Tilted Then Exit Sub
    Addscore 1000
    Light010.State = 1
    GoldHits(CurrentPlayer, 1) = 1
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "Ldt1"
End Sub

Sub Ldt2_Hit
    PlaySoundAt "fx_droptarget", Ldt2
    If Tilted Then Exit Sub
    Addscore 1000
    Light011.State = 1
    GoldHits(CurrentPlayer, 2) = 1
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "Ldt2"
End Sub

Sub Ldt3_Hit
    PlaySoundAt "fx_droptarget", Ldt3
    If Tilted Then Exit Sub
    Addscore 1000
    Light012.State = 1
    GoldHits(CurrentPlayer, 3) = 1
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "Ldt3"
End Sub

Sub Ldt4_Hit
    PlaySoundAt "fx_droptarget", Ldt4
    If Tilted Then Exit Sub
    PlaySound "s_AM_ GOLD coin drop"
    Addscore 1000
    Light013.State = 1
    GoldHits(CurrentPlayer, 4) = 1
    GoldHits(CurrentPlayer, 0) = 4
    ' check modes
    Loop3.Enabled = 1
    bGuardianMBReady = True
    If NOT bMultiBallMode Then
        Light004.State = 2
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Ldt4"
End Sub

' EGYPT: Right Bank: 5 droptargets
Sub Rdt1_Hit
    PlaySoundAt "fx_droptarget", Rdt1
    If Tilted Then Exit Sub
    Addscore 1000
    Light009.State = 1
    BlinkEgypt
    ' check modes
    EgyptHits(CurrentPlayer, 1) = 1
    CheckEgypt
    ' remember last trigger hit by the ball
    LastSwitchHit = "Rdt1"
End Sub

Sub Rdt2_Hit
    PlaySoundAt "fx_droptarget", Rdt2
    If Tilted Then Exit Sub
    Addscore 1000
    Light008.State = 1
    BlinkEgypt
    ' check modes
    EgyptHits(CurrentPlayer, 2) = 1
    CheckEgypt
    ' remember last trigger hit by the ball
    LastSwitchHit = "Rdt2"
End Sub

Sub Rdt3_Hit
    PlaySoundAt "fx_droptarget", Rdt3
    If Tilted Then Exit Sub
    Addscore 1000
    Light007.State = 1
    BlinkEgypt
    ' check modes
    EgyptHits(CurrentPlayer, 3) = 1
    CheckEgypt
    ' remember last trigger hit by the ball
    LastSwitchHit = "Rdt3"
End Sub

Sub Rdt4_Hit
    PlaySoundAt "fx_droptarget", Rdt4
    If Tilted Then Exit Sub
    Addscore 1000
    Light006.State = 1
    BlinkEgypt
    ' check modes
    EgyptHits(CurrentPlayer, 4) = 1
    CheckEgypt
    ' remember last trigger hit by the ball
    LastSwitchHit = "Rdt4"
End Sub

Sub Rdt5_Hit
    PlaySoundAt "fx_droptarget", Rdt5
    If Tilted Then Exit Sub
    Addscore 1000
    Light005.State = 1
    BlinkEgypt
    ' check modes
    EgyptHits(CurrentPlayer, 5) = 1
    CheckEgypt
    ' remember last trigger hit by the ball
    LastSwitchHit = "Rdt5"
End Sub

Sub BlinkEgypt
    If Light005.State Then EgyptLetters = " " Else EgyptLetters = "E"
    If Light006.State Then EgyptLetters = EgyptLetters & " " Else EgyptLetters = EgyptLetters & "G"
    If Light007.State Then EgyptLetters = EgyptLetters & " " Else EgyptLetters = EgyptLetters & "Y"
    If Light008.State Then EgyptLetters = EgyptLetters & " " Else EgyptLetters = EgyptLetters & "P"
    If Light009.State Then EgyptLetters = EgyptLetters & " " Else EgyptLetters = EgyptLetters & "T"
    DMD CL("1000"), CL("EGYPT"), "_", eNone, eNone, eNone, 100, True, "s_AM_Egypt target"
    DMD CL("1000"), CL(EgyptLetters), "_", eNone, eNone, eNone, 100, True, ""
    DMD CL("1000"), CL("EGYPT"), "_", eNone, eNone, eNone, 100, True, ""
    DMD CL("1000"), CL(EgyptLetters), "_", eNone, eNone, eNone, 100, True, ""
    DMD CL("1000"), CL("EGYPT"), "_", eNone, eNone, eNone, 100, True, ""
    DMD CL("1000"), CL(EgyptLetters), "_", eNone, eNone, eNone, 100, True, ""
    DMD CL("1000"), CL("EGYPT"), "_", eNone, eNone, eNone, 100, True, ""
    DMD CL("1000"), CL(EgyptLetters), "_", eNone, eNone, eNone, 100, True, ""
    DMD CL("1000"), CL("EGYPT"), "_", eNone, eNone, eNone, 100, True, ""
    DMD CL("1000"), CL(EgyptLetters), "_", eNone, eNone, eNone, 100, True, ""
End Sub

Sub CheckEgypt 'check if all EGYPT lights are on
    EgyptHits(CurrentPlayer, 0) = EgyptHits(CurrentPlayer, 1) + EgyptHits(CurrentPlayer, 2) + EgyptHits(CurrentPlayer, 3) + EgyptHits(CurrentPlayer, 4) + EgyptHits(CurrentPlayer, 5)
    If EgyptHits(CurrentPlayer, 0) = 5 Then
        EgyptCount = EgyptCount + 1
        'Rise the green guardian after 1/2 second so the ball doesn't hit the target on its way up
        vpmTimer.AddTimer 500, "GuardianTarget5.IsDropped = 0:GTL5.State = 2 '"
        DOF 146, DOFPulse
        Select Case EgyptCount
            Case 1:SetPlayfieldMultiplier 3
            Case 2:SetPlayfieldMultiplier 6
            Case 3:SetPlayfieldMultiplier 9
        End Select
    End If
End Sub

'Force Bank: 4 droptargets

Sub force1_Hit
    PlaySoundAt "fx_droptarget", force1
    fl1.State = 0
    If Tilted Then Exit Sub
    Addscore 50
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "force1"
End Sub

Sub force2_Hit
    PlaySoundAt "fx_droptarget", force2
    fl2.State = 0
    If Tilted Then Exit Sub
    Addscore 50
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "force2"
End Sub

Sub force3_Hit
    PlaySoundAt "fx_droptarget", force3
    fl3.State = 0
    If Tilted Then Exit Sub
    Addscore 50
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "force3"
End Sub

Sub force4_Hit
    PlaySoundAt "fx_droptarget", force4
    fl4.State = 0
    If Tilted Then Exit Sub
    Addscore 50
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "force4"
End Sub

' Tomb bank: 2 droptargets

Sub tomb1_Hit
    PlaySoundAt "fx_droptarget", RightHole
    If Tilted Then Exit Sub
    Addscore 2000
    tomb1.IsDropped = 1
    If bMultiBallMode Then
        PlaySound "s_AM_StoneDoor"
    Else
        DMD "_", CL("TOMB DOOR"), "_", eNone, eNone, eNone, 1000, True, "s_AM_StoneDoor"
    End If
    ATONLight.duration 2, 4500, 0
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "tomb1"
End Sub

Sub tomb2_Hit
    PlaySoundAt "fx_droptarget", RightHole
    If Tilted Then Exit Sub
    Addscore 3000
    If NOT BallInRightHole Then
        tomb2.IsDropped = 1
        If bMultiBallMode Then
            PlaySound "s_AM_StoneDoor"
        Else
            DMD "_", CL("TOMB DOOR"), "_", eNone, eNone, eNone, 1000, True, "s_AM_StoneDoor"
        End If
        ATONLight.duration 2, 4500, 0
    End If
    ' check modes

    ' remember last trigger hit by the ball
    LastSwitchHit = "tomb2"
End Sub

Sub RiseTombBank
    PlaySoundAt SoundFXDOF("fx_resetdrop", 144, DOFPulse, DOFContactors), RightHole
    tomb1.IsDropped = 0
    tomb2.IsDropped = 0
End Sub

' Holes - saucers

Sub TopHole_Hit 'top hole
    PlaySoundAt "fx_kicker_enter", TopHole
    DMDFlush
    If Tilted Then
        EjectTopHole
        Exit Sub
    End If
    If Not bAtonMB Then
        BallInTopHole = True
        LockFlash1.State = 2
        LockedBalls = LockedBalls + 1
        DMD "_", CL("BALL TRAPPED"), "_", eNone, eNone, eNone, 1200, True, "s_AM_BallTrapped"
        AddScore 10000
        Tdt3.IsDropped = 0:PlaySoundAt SoundFXDOF("fx_resetdrop", 143, DOFPulse, DOFContactors), TopHole
        CheckLockedBalls
        vpmTimer.AddTimer 1000, "AddMultiball 1 '"
    Else 'during AtonMB
        bNorthLock = True
        LockFlash1.State = 2
        If bEastLock Then
            AwardSuperJackpot
            bNorthLock = False
            bEastLock = False
            LockFlash1.State = 0
            LockFlash2.State = 0
            FlashForMs LockFlash1, 3000, 50, 0
            FlashForMs LockFlash2, 3000, 50, 0
        Else
            DMD "_", CL("NORTH LOCK"), "_", eNone, eNone, eNone, 1200, True, "s_LOOP"
        End If
        vpmTimer.AddTimer 3000, "EjectTopHole '"
    End If
End Sub

Sub RightHole_Hit 'top hole
    PlaySoundAt "fx_kicker_enter", RightHole
    DMDFlush
    If Tilted Then
        EjectRightHole
        Exit Sub
    End If
    If Not bAtonMB Then
        BallInRightHole = True
        LockFlash2.State = 2
        LockedBalls = LockedBalls + 1
        DMD "_", CL("BALL TRAPPED"), "_", eNone, eNone, eNone, 1200, True, "s_AM_BallTrapped"
        AddScore 10000
        tomb2.IsDropped = 0:PlaySoundAt SoundFXDOF("fx_resetdrop", 144, DOFPulse, DOFContactors), RightHole
        CheckLockedBalls
        vpmTimer.AddTimer 1000, "AddMultiball 1 '"
    Else 'during AtonMB
        bEastLock = True
        LockFlash2.State = 2
        If bNorthLock Then
            AwardSuperJackpot
            bNorthLock = False
            bEastLock = False
            LockFlash1.State = 0
            LockFlash2.State = 0
            FlashForMs LockFlash1, 3000, 50, 0
            FlashForMs LockFlash2, 3000, 50, 0
        Else
            DMD "_", CL("EAST LOCK"), "_", eNone, eNone, eNone, 1200, True, "s_LOOP"
        End If
        vpmTimer.AddTimer 3000, "EjectRightHole '"
    End If
End Sub

Sub TopHole_Unhit 'be sure all the droptargets are down
    Tdt1.IsDropped = 1
    Tdt2.IsDropped = 1
    Tdt3.IsDropped = 1
End Sub

Sub RightHole_Unhit 'be sure all the droptargets are down
    Tomb1.IsDropped = 1
    Tomb2.IsDropped = 1
End Sub

Sub EjectTopHole
    If bAtonMB Then
        TopHole.Kick 220, 8
        PlaySoundAt SoundFXDOF("fx_kicker", 118, DOFPulse, DOFContactors), TopHole
    End If
    If BallinTopHole Then
        BallInTopHole = False
        LockFlash1.State = 0
        TopHole.Kick 220, 8
        PlaySoundAt SoundFXDOF("fx_kicker", 118, DOFPulse, DOFContactors), TopHole
        PlaySound "s_AM_BallTrapsReleased"
        LockedBalls = LockedBalls -1
        LightEffect 5
    End If
End Sub

Sub EjectRightHole
    If bAtonMB Then
        RightHole.Kick 200, 8
        PlaySoundAt SoundFXDOF("fx_kicker", 116, DOFPulse, DOFContactors), RightHole
    End If
    If BallInRightHole Then
        BallInRightHole = False
        LockFlash2.State = 0
        RightHole.Kick 200, 8
        PlaySoundAt SoundFXDOF("fx_kicker", 116, DOFPulse, DOFContactors), RightHole
        LockedBalls = LockedBalls -1
        LightEffect 5
    End If
End Sub

Sub CheckLockedBalls
    If BallInRightHole AND BallInTopHole Then
        bAtonMBReady = True
        If NOT bMultiBallMode Then
            Light004.State = 2
        End If
    End If
End Sub

Sub CenterHole_Hit 'Center Hole: Cobra Vault
    Dim delay
    delay = 500
    PlaySoundAt "fx_kicker_enter", CenterHole
    If Tilted Then EjectCenterHole:Exit Sub
    Light004.State = 0
    DMDFlush
    'check for multiball Jackpots
    If bGuardianMB OR bPriestMB OR bPriestessMB OR bAtonMB Then
        DMD CL("COBRA JACKPOT"), CL("15.000"), "", eNone, eBlinkFast, eNone, 2000, True, "s_JACKPOT_Cobra"
        AddScore 15000
        Delay = 2500
    Else
        'check for starting multiballs
        If bGuardianMBReady Then 'start guardian multiball
            StartGuardianMB
            delay = 3500
        ElseIf bPriestMBReady Then
            StartPriestMB
            delay = 6500
        ElseIf bPriestessMBReady Then
            StartPriestessMB
            delay = 6500
        ElseIf bAtonMBReady Then
            StartAtonMB
            delay = 15658 'wait for the locked balls to be released
        ElseIf bRoamDesert Then
            PlaySound"s_cobragasp"
            AddScore 5000
            delay = 1000
        ' else normal play
        Else
            CobraVaultAward 'give an award, it takes about 4 seconds, speech max 7-8 seconds
            delay = 7000
        End If
    End If
    vpmTimer.AddTimer delay, "FlashForMs Light024, 1000, 50, 0 '"
    vpmTimer.AddTimer delay + 1000, "EjectCenterHole '"
End Sub

Sub EjectCenterHole
    CenterHole.Kick 182, 20
    PlaySoundAt SoundFXDOF("fx_kicker", 117, DOFPulse, DOFContactors), CenterHole
    LightEffect 5
End Sub

Sub CobraVaultAward
    Dim tmp, tmp2
    tmp2 = RndNbr(31) 'from 1 to 31 - Becca voice
    tmp = RndNbr(100) 'from 1 to 100 - Award
    'dmd
    DMD CL("COBRA VAULT"), "", "d_border", eNone, eNone, eNone, 1000, True, ""
    If tmp2 = 31 Then FlashForMs ForceLight037, 7000, 50, 0 'blink the aton flasher
    ' do not sound Becca on some awards
    If tmp<> 31 AND tmp<> 32 AND tmp<> 33 Then PlaySound "s_bs" &tmp2
    DMD "_", CL("SEARCHING"), "_", eNone, eBlink, eNone, 1500, True, ""
    'award
    Select Case True
        Case tmp <= 30:DMD "_", CL("NO TREASURE FOUND"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 50
            DMD "_", CL("50 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 31:DMD CL("STARTING"), CL("TOMB DYNAMITE"), "_", eNone, eNone, eNone, 2000, True, ""
            'start tomb dynamite
            StartTombDynamite
        Case tmp <= 33:
            'start roam the desert
            DMD CL("STARTING"), CL("ROAM THE DESERT"), "_", eNone, eNone, eNone, 2000, True, "s_Roam_Speech"
            DMD "SUPER LOOPS", CL("DOUBLE SCORING"), "_", eNone, eBlink, eNone, 2000, True, ""
            DMD "HIT 3 PURPLE ARROWS", CL("IN 60 SECONDS"), "_", eNone, eNone, eNone, 2000, True, ""
            StartRoamDesert
        Case tmp <= 35:DMD "_", CL("TOURIST PAINTING"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 100
            DMD "_", CL("100 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 37:DMD "_", CL("TOY HIPPO"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 100
            DMD "_", CL("100 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 39:DMD "_", CL("SNAKE ANTI-VENOM"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 1000
            DMD "_", CL("1000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 41:DMD "_", CL("HIBISCUS TEA"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 500
            DMD "_", CL("500 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 43:DMD "_", CL("LAPIS LAZULI RING"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 1000
            DMD "_", CL("1000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 45:DMD "_", CL("MONKEYS PAW"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 1500
            DMD "_", CL("1500 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 47:DMD "_", CL("CROCODILE CARVING"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 2500
            DMD "_", CL("2500 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 49:DMD "_", CL("METEORITE DAGGER"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 3000
            DMD "_", CL("3000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 53:DMD "_", CL("USED SOCKS"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 100
            DMD "_", CL("100 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 56:DMD "_", CL("SUNGLASSES"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 100
            DMD "_", CL("100 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 59:DMD "_", CL("FAKE MAP"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 100
            DMD "_", CL("100 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 62:DMD "_", CL("OLD CANTEEN"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 500
            DMD "_", CL("500 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 65:DMD "_", CL("BULLET CASING"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 500
            DMD "_", CL("500 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 68:DMD "_", CL("COMPASS"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 1000
            DMD "_", CL("1.000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 71:DMD "_", CL("FLASHLIGHT"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 1000
            DMD "_", CL("1.000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 74:DMD "_", CL("E.T. CARTRIDGE"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 1500
            DMD "_", CL("1.500 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 77:DMD "_", CL("ANCIENT ARTIFACT"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 1500
            DMD "_", CL("1.500 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 80:DMD "_", CL("PAPYRUS FRAGMENT"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 2500
            DMD "_", CL("2.500 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 83:DMD "_", CL("SILVER STATUE"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 2500
            DMD "_", CL("2.500 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 86:DMD "_", CL("IBIS MUMMY"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 3000
            DMD "_", CL("3.000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 89:DMD "_", CL("IVORY CANE"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 3000
            DMD "_", CL("3.000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 92:DMD "_", CL("SCARAB NECKLACE"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 5000
            DMD "_", CL("5.000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 95:DMD "_", CL("COPPER SCROLL"), "_", eNone, eNone, eNone, 2000, True, "":AddScore 5000
            DMD "_", CL("5.000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
        Case tmp <= 98:DMD "_", CL("GOLD COINS"), "_", eNone, eNone, eNone, 2000, True, "s_AM_ GOLD coin drop":AddScore 10000
            DMD "_", CL("10.000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
            DropLeftBank
        Case tmp <= 100:DMD "_", CL("HIDDEN JACKPOT"), "_", eNone, eNone, eNone, 2000, True, "s_SecretPass":AddScore 500000
            DMD "_", CL("500.000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
    End Select
End Sub

' Vault Modes
' Tomb Dynamite ' drop one of the top tomb's targets
Sub StartTombDynamite
    If Tdt1.IsDropped = 0 Then
       vpmTimer.AddTimer 1000, "Tdt1.IsDropped = 1: PlaySoundAt ""fx_droptarget"", TopHole '"
       vpmTimer.AddTimer 1500, "PlaySound ""s_AM_StoneDoor"" '"
       PlaySound"s_dynamite"
       LightEffect 5
       FlashForMs LockFlash1, 1000, 50, 0
       StartAtonLightShowFaster
    ElseIf Tdt2.IsDropped = 0 Then
       vpmTimer.AddTimer 1000, "Tdt2.IsDropped = 1: PlaySoundAt ""fx_droptarget"", TopHole '"
       vpmTimer.AddTimer 1500, "PlaySound ""s_AM_StoneDoor"" '"
       PlaySound"s_dynamite"
       LightEffect 5
       FlashForMs LockFlash1, 1000, 50, 0
       StartAtonLightShowFaster
    ElseIf Tdt3.IsDropped = 0 Then
       vpmTimer.AddTimer 1000, "Tdt3.IsDropped = 1: PlaySoundAt ""fx_droptarget"", TopHole '"
       vpmTimer.AddTimer 1500, "PlaySound ""s_AM_StoneDoor"" '"
       PlaySound"s_dynamite"
       LightEffect 5
       FlashForMs LockFlash1, 1000, 50, 0
       StartAtonLightShowFaster
    Else
        Addscore 5000
    End If
End Sub

' Roam the Desert - double scoring on the loops for 30 seconds
Sub StartRoamDesert 'from the Cobra Vault award
    bRoamDesert = True
    ChangeSong
    EnableBallSaver 60
    Light001.State = 2
    Light002.State = 2
    Light003.State = 2
    HiddenPassHits = 0
    Select Case RndNbr(2)
        Case 1: pinkarrow1.State = 2
        Case 2: pinkarrow2.State = 2
    End Select
    ChangeGi purple
    ChangeGIIntensity 2
    rdl.state = 2
    'HiddenPF1. Visible = 1
    RoamDesertTimer.Enabled = 1 'to stop the mode
End Sub

Sub RoamDesertTimer_timer 'from an cobra vault award
    If bRoamDesert2 Then
        DMDFlush
        DMD CL("ATOMIC LEVELS"), CL("NORMALIZED"), "_", eNone, eBlink, eNone, 2000, True, "s_deenergized"
        DropRadiTargets
        bPriestessMB = False
        bPriestMB = False
    End If
    bRoamDesert = False
    bRoamDesert2 = False
    ChangeSong
    RoamDesertTimer.Enabled = 0
    Light001.State = 0
    Light002.State = 0
    Light003.State = 0
    pinkarrow1.State = 0
    pinkarrow2.State = 0
    rdl.State = 0
    'HiddenPF1.Visible = 0
    ChangeGIIntensity 1
    ChangeGi white
End Sub

Sub StartRoamDesert2 'part of the wizard mode after defeating a multiball mode, runs until multiball stops
    RiseRadiTargets
    bRoamDesert = True
    bRoamDesert2= True
    ChangeSong
    Light001.State = 2
    Light002.State = 2
    Light003.State = 2
    pinkarrow1.State = 0
    pinkarrow2.State = 0
    rdl.State = 0
    'HiddenPF1.Visible = 0
    RoamDesertTimer.Enabled = 0 'stop the timer so the mode ends when the multiball ends.
End Sub

' Guards

Sub GuardianTarget1_Hit 'quick sand
    PlaySoundAt "fx_droptarget", GuardianTarget1
    GTL1.State = 0
    If Tilted Then Exit Sub
    LightEffect 4
    If bMultiBallMode Then
        PlaySound "s_shot " &RndNbr(2)
    Else
        DMD "_", " GUARDIAN TAKEDOWN", "_", eNone, eBlinkFast, eNone, 1000, True, "s_shot " &RndNbr(2)
    End If
    ' check modes
    If bAtonMB Then
        Addscore 5000
    ElseIf bGuardianMB Then
        Addscore 5000
        CheckGuardiansDown
    Else
        Addscore 2200
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "GuardianTarget1"
End Sub

Sub GuardianTarget1_Timer
    Me.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", GuardianTarget1
    GuardianTarget1.IsDropped = 1
    GTL1.State = 0
End Sub

Sub GuardianTarget2_Hit 'Sand 1
    PlaySoundAt "fx_droptarget", GuardianTarget2
    GTL2.State = 0
    If Tilted Then Exit Sub
    LightEffect 4
    If bMultiBallMode Then
        PlaySound "s_shot " &RndNbr(2)
    Else
        DMD "_", " GUARDIAN TAKEDOWN", "_", eNone, eBlinkFast, eNone, 1000, True, "s_shot " &RndNbr(2)
    End If
    ' check modes
    If bAtonMB Then
        Addscore 5000
    ElseIf bGuardianMB Then
        Addscore 5000
        CheckGuardiansDown
    Else
        Addscore 2200
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "GuardianTarget2"
End Sub

Sub GuardianTarget2_Timer
    Me.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", GuardianTarget2
    GuardianTarget2.IsDropped = 1
    GTL2.State = 0
End Sub

Sub GuardianTarget3_Hit 'Sand 2
    PlaySoundAt "fx_droptarget", GuardianTarget3
    GTL3.State = 0
    If Tilted Then Exit Sub
    LightEffect 4
    If bMultiBallMode Then
        PlaySound "s_shot " &RndNbr(2)
    Else
        DMD "_", " GUARDIAN TAKEDOWN", "_", eNone, eBlinkFast, eNone, 1000, True, "s_shot " &RndNbr(2)
    End If
    ' check modes
    If bAtonMB Then
        Addscore 5000
    ElseIf bGuardianMB Then
        Addscore 5000
        CheckGuardiansDown
    Else
        Addscore 2200
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "GuardianTarget3"
End Sub

Sub GuardianTarget3_Timer
    Me.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", GuardianTarget3
    GuardianTarget3.IsDropped = 1
    GTL3.State = 0
End Sub

Sub GuardianTarget4_Hit 'Sand 3
    PlaySoundAt "fx_droptarget", GuardianTarget4
    GTL4.State = 0
    If Tilted Then Exit Sub
    LightEffect 4
    If bMultiBallMode Then
        PlaySound "s_shot " &RndNbr(2)
    Else
        DMD "_", " GUARDIAN TAKEDOWN", "_", eNone, eBlinkFast, eNone, 1000, True, "s_shot " &RndNbr(2)
    End If
    ' check modes
    If bAtonMB Then
        Addscore 5000
    ElseIf bGuardianMB Then
        Addscore 5000
        CheckGuardiansDown
    Else
        Addscore 2200
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "GuardianTarget4"
End Sub

Sub GuardianTarget4_Timer
    Me.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", GuardianTarget4
    GuardianTarget4.IsDropped = 1
    GTL4.State = 0
End Sub

Sub GuardianTarget5_Hit 'River
    PlaySoundAt "fx_droptarget", GuardianTarget5
    GTL5.State = 0
    If Tilted Then Exit Sub
    LightEffect 2
    ' check modes
    DMD "_", CL("RIVER GUARDIAN"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_shot " &RndNbr(2)
    Addscore 3500
    vpmTimer.AddTimer 1500, "RiseRightBank '"
    ' remember last trigger hit by the ball
    LastSwitchHit = "GuardianTarget5"
End Sub

Sub GuardianTarget6_Hit 'Right ramp
    PlaySoundAt "fx_droptarget", GuardianTarget6
    GTL6.State = 0
    If Tilted Then Exit Sub
    LightEffect 4
    If bMultiBallMode Then
        PlaySound "s_shot " &RndNbr(2)
    Else
        DMD "_", " GUARDIAN TAKEDOWN", "_", eNone, eBlinkFast, eNone, 1000, True, "s_shot " &RndNbr(2)
    End If
    ' check modes
    If bAtonMB Then
        Addscore 5000
    ElseIf bGuardianMB Then
        Addscore 5000
        CheckGuardiansDown
    Else
        Addscore 2200
    vpmTimer.AddTimer 1500, "RiseRightBank '"
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "GuardianTarget6"
End Sub

Sub GuardianTarget6_Timer
    Me.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", GuardianTarget6
    GuardianTarget6.IsDropped = 1
    GTL6.State = 0
End Sub

Sub GuardianTarget7_Hit 'Right hole
    PlaySoundAt "fx_droptarget", GuardianTarget7
    GTL7.State = 0
    If Tilted Then Exit Sub
    LightEffect 4
    If bMultiBallMode Then
        PlaySound "s_shot " &RndNbr(2)
    Else
        DMD "_", " GUARDIAN TAKEDOWN", "_", eNone, eBlinkFast, eNone, 1000, True, "s_shot " &RndNbr(2)
    End If
    ' check modes
    If bAtonMB Then
        Addscore 5000
    ElseIf bGuardianMB Then
        Addscore 5000
        CheckGuardiansDown
    Else
        Addscore 2200
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "GuardianTarget7"
End Sub

Sub GuardianTarget7_Timer
    Me.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", GuardianTarget7
    GuardianTarget7.IsDropped = 1
    GTL7.State = 0
End Sub

Sub GuardianTarget8_Hit 'left ramp
    PlaySoundAt "fx_droptarget", GuardianTarget8
    GTL8.State = 0
    If Tilted Then Exit Sub
    LightEffect 4
    If bMultiBallMode Then
        PlaySound "s_shot " &RndNbr(2)
    Else
        DMD "_", " GUARDIAN TAKEDOWN", "_", eNone, eBlinkFast, eNone, 1000, True, "s_shot " &RndNbr(2)
    End If
    ' check modes
    If bAtonMB Then
        Addscore 5000
    ElseIf bGuardianMB Then
        Addscore 5000
        CheckGuardiansDown
    Else
        Addscore 2200
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "GuardianTarget8"
End Sub

Sub GuardianTarget8_Timer
    Me.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", GuardianTarget8
    GuardianTarget8.IsDropped = 1
    GTL8.State = 0
End Sub

Sub GuardianTarget9_Hit 'top left lane
    PlaySoundAt "fx_droptarget", GuardianTarget9
    GTL9.State = 0
    If Tilted Then Exit Sub
    LightEffect 4
    If bMultiBallMode Then
        PlaySound "s_shot " &RndNbr(2)
    Else
        DMD "_", " GUARDIAN TAKEDOWN", "_", eNone, eBlinkFast, eNone, 1000, True, "s_shot " &RndNbr(2)
    End If
    ' check modes
    If bAtonMB Then
        Addscore 5000
    ElseIf bGuardianMB Then
        Addscore 5000
        CheckGuardiansDown
    Else
        Addscore 2200
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "GuardianTarget9"
End Sub

Sub GuardianTarget9_Timer
    Me.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", GuardianTarget9
    GuardianTarget9.IsDropped = 1
    GTL9.State = 0
End Sub

' AT targets: 8 drop targets

Sub At1_Hit
    PlaySoundAt "fx_droptarget", at1
    DropAt1
    If Tilted Then Exit Sub
    ' check modes
    FlashForMs ForceLight037, 800, 40, 0
    DMDFlush
    DMD CL("ATON HIT"), CL("250.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_ATON HIT 1"
    AddScore 250000
    Force1.TimerInterval = 9000
    vpmTimer.AddTimer 1500, "RiseForceGuardians '"
    ' remember last trigger hit by the ball
    LastSwitchHit = "at1"
End Sub

Sub At2_Hit
    PlaySoundAt "fx_droptarget", at2
    DropAt2
    If Tilted Then Exit Sub
    ' check modes
    If Force1.TimerEnabled Then DropForceGuardians
    FlashForMs ForceLight037, 800, 40, 0
    DMDFlush
    DMD CL("ATON HIT"), CL("250.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_ATON HIT 2"
    AddScore 250000
    If Force1.TimerEnabled Then DropForceGuardians
    vpmTimer.AddTimer 1000, "Playsound Song, -1, 0.1,0, 0, 0, 1, 0, 0 '" 'reduce the volume of the playing song
    vpmTimer.AddTimer 2000, "DMD CL(""WARNING""), CL(""ANTIGRAVITY ATTACK""), ""_"",eBlinkFast,eNone,eNone,1500,True,""s_ATON_antigravityATTACK1"" '"
    vpmTimer.AddTimer 8500, "AntiGravityAttack '"
    vpmTimer.AddTimer 13230, "Playsound Song, -1, SongVolume,0, 0, 0, 1, 0, 0 '" 'restore the volume of the playing song
    ' remember last trigger hit by the ball
    LastSwitchHit = "at2"
End Sub

Sub At3_Hit
    PlaySoundAt "fx_droptarget", at3
    DropAt3
    If Tilted Then Exit Sub
    ' check modes
    FlashForMs ForceLight037, 800, 40, 0
    DMDFlush
    DMD CL("ATON HIT"), CL("250.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_ATON HIT 1"
    AddScore 250000
    Force1.TimerInterval = 7000
    vpmTimer.AddTimer 1000, "RiseForceGuardians '"
    ' remember last trigger hit by the ball
    LastSwitchHit = "at3"
End Sub

Sub At4_Hit
    PlaySoundAt "fx_droptarget", at4
    DropAt4
    If Tilted Then Exit Sub
    ' check modes
    FlashForMs ForceLight037, 800, 40, 0
    DMDFlush
    DMD CL("ATON HIT"), CL("250.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_ATON HIT 2"
    AddScore 250000
    If Force1.TimerEnabled Then DropForceGuardians
    ' remember last trigger hit by the ball
    LastSwitchHit = "at4"
End Sub

Sub At5_Hit
    PlaySoundAt "fx_droptarget", at5
    DropAt5
    If Tilted Then Exit Sub
    ' check modes
    FlashForMs ForceLight037, 800, 40, 0
    DMDFlush
    DMD CL("ATON HIT"), CL("250.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_ATON HIT 1"
    AddScore 250000
    Force1.TimerInterval = 5000
    vpmTimer.AddTimer 1000, "RiseForceGuardians '"
    ' remember last trigger hit by the ball
    LastSwitchHit = "at5"
End Sub

Sub At6_Hit
    PlaySoundAt "fx_droptarget", at6
    DropAt6
    If Tilted Then Exit Sub
    ' check modes
    FlashForMs ForceLight037, 800, 40, 0
    DMDFlush
    DMD CL("ATON HIT"), CL("250.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_ATON HIT 2"
    AddScore 250000
    If Force1.TimerEnabled Then DropForceGuardians
    vpmTimer.AddTimer 1000, "Playsound Song, -1, 0.1,0, 0, 0, 1, 0, 0 '" 'reduce the volume of the playing song
    vpmTimer.AddTimer 2000, "DMD CL(""WARNING""), CL(""ANTIGRAVITY ATTACK""), ""_"",eBlinkFast,eNone,eNone,1500,True,""s_ATON_antigravityATTACK1"" '"
    vpmTimer.AddTimer 8500, "AntiGravityAttack '"
    vpmTimer.AddTimer 13230, "Playsound Song, -1, SongVolume,0, 0, 0, 1, 0, 0 '" 'restore the volume of the playing song
    ' remember last trigger hit by the ball
    LastSwitchHit = "at6"
End Sub

Sub At7_Hit
    PlaySoundAt "fx_droptarget", at7
    DropAt7
    If Tilted Then Exit Sub
    ' check modes
    FlashForMs ForceLight037, 800, 40, 0
    DMDFlush
    DMD CL("ATON HIT"), CL("250.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_ATON HIT 1"
    AddScore 250000
    bAtonMB8= True
    ChangeSong
    RiseGuardians
    ' remember last trigger hit by the ball
    LastSwitchHit = "at7"
End Sub

Sub At8_Hit
    PlaySoundAt "fx_droptarget", at8
    DropAt8
    bAtonMB8 = False
    If Tilted Then Exit Sub
    ' check modes
    FlashForMs ForceLight037, 800, 40, 0
    DMDFlush
    DMD CL("ATON DEFEATED"), CL("1.100.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_ATON HIT FINAL"
    AddScore 1000000
    AwardExtraBall
    WinAtonMB
    DropForceGuardians_Solenoid
    ' remember last trigger hit by the ball
    LastSwitchHit = "at8"
End Sub

Sub AntiGravityAttack
    PlaySound "s_ATON_antigravityATTACK"
    Dim BOT, b
    BOT = GetBalls
    For b = lob to UBound(BOT)
        BOT(b).VelY = -20 - RndNbr(20)
    Next
End Sub

' Priests: 2 targets

Sub Priest_Hit
    PlaySoundAt "fx_target", Priestp
    If Tilted Then Exit Sub
    ' check modes
    If bPriestMB Then
        LightEffect 9
        PriestHits(CurrentPlayer) = PriestHits(CurrentPlayer) + 1
        AerialValue(CurrentPlayer) = AerialValue(CurrentPlayer) - 1 'turn off one of the CURSE lights
        UpdateCURSE
        Select Case PriestHits(CurrentPlayer)
            Case 5:Addscore 200000:WinPriestMB
            Case Else:AddScore 25000:DMDFlush:DMD CL("PRIEST HIT"), CL("25.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_Priest Hit"
        End Select
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Priest"
End Sub

Sub PRIESTESS_Hit
    PlaySoundAt "fx_target", PRIESTESSp
    If Tilted Then Exit Sub
    ' check modes
    If bPriestessMB Then
        LightEffect 9
        PriestessHits(CurrentPlayer) = PriestessHits(CurrentPlayer) + 1
        LidalValue(CurrentPlayer) = LidalValue(CurrentPlayer) - 1 'turn off one of the WOKEN lights
        UpdateWOKEN
        Select Case PriestessHits(CurrentPlayer)
            Case 5:Addscore 200000:WinPriestessMB
            Case Else:AddScore 25000:DMDFlush:DMD CL("PRIESTESS HIT"), CL("25.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_Priest Hit"
        End Select
    End If

    ' remember last trigger hit by the ball
    LastSwitchHit = "PRIESTESS"
End Sub

' Animations

'AT targets & priests
Sub RiseAt1:PlaySoundAt SoundFXDOF("fx_resetdrop", 148, DOFPulse, DOFContactors), at1p:at1f.RotatetoEnd:at1.IsDropped = 0:End Sub
Sub DropAt1:PlaySoundAt "fx_droptarget", at1p:at1f.RotatetoStart:at1.IsDropped = 1:End Sub

Sub RiseAt2:PlaySoundAt SoundFXDOF("fx_resetdrop", 149, DOFPulse, DOFContactors), at2p:at2f.RotatetoEnd:at2.IsDropped = 0:End Sub
Sub DropAt2:PlaySoundAt "fx_droptarget", at2p:at2f.RotatetoStart:at2.IsDropped = 1:End Sub

Sub RiseAt3:PlaySoundAt SoundFXDOF("fx_resetdrop", 148, DOFPulse, DOFContactors), at3p:at3f.RotatetoEnd:at3.IsDropped = 0:End Sub
Sub DropAt3:PlaySoundAt "fx_droptarget", at3p:at3f.RotatetoStart:at3.IsDropped = 1:End Sub

Sub RiseAt4:PlaySoundAt SoundFXDOF("fx_resetdrop", 149, DOFPulse, DOFContactors), at4p:at4f.RotatetoEnd:at4.IsDropped = 0:End Sub
Sub DropAt4:PlaySoundAt "fx_droptarget", at4p:at4f.RotatetoStart:at4.IsDropped = 1:End Sub

Sub RiseAt5:PlaySoundAt SoundFXDOF("fx_resetdrop", 148, DOFPulse, DOFContactors), at5p:at5f.RotatetoEnd:at5.IsDropped = 0:End Sub
Sub DropAt5:PlaySoundAt "fx_droptarget", at5p:at5f.RotatetoStart:at5.IsDropped = 1:End Sub

Sub RiseAt6:PlaySoundAt SoundFXDOF("fx_resetdrop", 149, DOFPulse, DOFContactors), at6p:at6f.RotatetoEnd:at6.IsDropped = 0:End Sub
Sub DropAt6:PlaySoundAt "fx_droptarget", at6p:at6f.RotatetoStart:at6.IsDropped = 1:End Sub

Sub RiseAt7:PlaySoundAt SoundFXDOF("fx_resetdrop", 148, DOFPulse, DOFContactors), at7p:at7f.RotatetoEnd:at7.IsDropped = 0:End Sub
Sub DropAt7:PlaySoundAt "fx_droptarget", at7p:at7f.RotatetoStart:at7.IsDropped = 1:End Sub

Sub RiseAt8:PlaySoundAt SoundFXDOF("fx_resetdrop", 149, DOFPulse, DOFContactors), at8p:at8f.RotatetoEnd:at8.IsDropped = 0:End Sub
Sub DropAt8:PlaySoundAt "fx_droptarget", at8p:at8f.RotatetoStart:at8.IsDropped = 1:End Sub

Sub DropAllAT
    vpmTimer.AddTimer 100, "DropAt1 '"
    vpmTimer.AddTimer 200, "DropAt2 '"
    vpmTimer.AddTimer 300, "DropAt3 '"
    vpmTimer.AddTimer 400, "DropAt4 '"
    vpmTimer.AddTimer 500, "DropAt5 '"
    vpmTimer.AddTimer 600, "DropAt6 '"
    vpmTimer.AddTimer 700, "DropAt7 '"
    vpmTimer.AddTimer 800, "DropAt8 '"
End Sub

Sub DropAllAT_Solenoid
    PlaySoundAt "fx_droptarget_solenoid", at1p
    at1f.RotatetoStart:at1.IsDropped = 1
    at2f.RotatetoStart:at2.IsDropped = 1
    at3f.RotatetoStart:at3.IsDropped = 1
    at4f.RotatetoStart:at4.IsDropped = 1
    at5f.RotatetoStart:at5.IsDropped = 1
    at6f.RotatetoStart:at6.IsDropped = 1
    at7f.RotatetoStart:at7.IsDropped = 1
    at8f.RotatetoStart:at8.IsDropped = 1
End Sub

Sub RiseAllAT
    vpmTimer.AddTimer 240, "RiseAt8 '"
    vpmTimer.AddTimer 405, "RiseAt7 '"
    vpmTimer.AddTimer 570, "RiseAt6 '"
    vpmTimer.AddTimer 735, "RiseAt5 '"
    vpmTimer.AddTimer 900, "RiseAt4 '"
    vpmTimer.AddTimer 1065, "RiseAt3 '"
    vpmTimer.AddTimer 1230, "RiseAt2 '"
    vpmTimer.AddTimer 1395, "RiseAt1 '"
End Sub

Sub RisePRIESTESS
    PlaySoundAt "fx_resetdrop", PRIESTESSp
    PlaySound "s_Priestess_UP"
    SetLightColor Light035, red, 2
    PriestessTlight.State = 2
    PRIESTESSf.RotatetoEnd
    PRIESTESS.IsDropped = 0
    PriestessWall.IsDropped = 0
End Sub

Sub DropPRIESTESS
    PlaySoundAt "fx_droptarget", PRIESTESSp
    PlaySound "s_Priestess_DOWN"
    Light035.State = 0
    PriestessTlight.State = 0
    PRIESTESSf.RotatetoStart
    PRIESTESS.IsDropped = 1
    PRIESTESSWall.IsDropped = 1
End Sub

Sub DropPRIESTESS_Solenoid
    PlaySoundAt "fx_droptarget_solenoid", PRIESTESSp
    Light035.State = 0
    PriestessTlight.State = 0
    PRIESTESSf.RotatetoStart
    PRIESTESS.IsDropped = 1
    PRIESTESSWall.IsDropped = 1
End Sub

Sub RisePriest
    PlaySoundAt "fx_resetdrop", Priestp
    PlaySound "s_Priest_UP"
    SetLightColor Light036, purple, 2
    PriestTlight.State = 2
    Priestf.RotatetoEnd
    Priest.IsDropped = 0
    PriestWall.IsDropped = 0
End Sub

Sub DropPriest
    PlaySoundAt "fx_droptarget", Priestp
    PlaySound "s_Priest_DOWN"
    Light036.State = 0
    PriestTlight.State = 0
    Priestf.RotatetoStart
    Priest.IsDropped = 1
    PriestWall.IsDropped = 1
End Sub

Sub DropPriest_Solenoid
    PlaySoundAt "fx_droptarget_solenoid", Priestp
    Light036.State = 0
    PriestTlight.State = 0
    Priestf.RotatetoStart
    Priest.IsDropped = 1
    PriestWall.IsDropped = 1
End Sub

' Guards

Sub DropForceGuardians
    force1.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", force4
    force1.IsDropped = 1:FL1.State = 0
    force2.IsDropped = 1:FL2.State = 0
    force4.IsDropped = 1:FL4.State = 0
    force3.IsDropped = 1:FL3.State = 0
    PlaySound "s_forcefdown"
End Sub

Sub DropForceGuardians_Solenoid
    force1.TimerEnabled = 0
    PlaySoundAt "fx_droptarget_solenoid", force4
    force1.IsDropped = 1:FL1.State = 0
    force2.IsDropped = 1:FL2.State = 0
    force4.IsDropped = 1:FL4.State = 0
    force3.IsDropped = 1:FL3.State = 0
End Sub

Sub RiseForceGuardians
    PlaySoundAt SoundFXDOF("fx_resetdrop", 147, DOFPulse, DOFContactors), force4
    force1.IsDropped = 0:FL1.State = 1
    force2.IsDropped = 0:FL2.State = 1
    force4.IsDropped = 0:FL4.State = 1
    force3.IsDropped = 0:FL3.State = 1
    PlaySound "s_forcefup"
    Force1.TimerEnabled = 1
End Sub

Sub Force1_Timer
    DropForceGuardians
End Sub

Sub DropGuardians
    PlaySoundAt "fx_droptarget_solenoid", GuardianTarget2
    GuardianTarget1.IsDropped = 1:GTL1.State = 0: GuardianTarget1.TimerEnabled = 0
    GuardianTarget2.IsDropped = 1:GTL2.State = 0: GuardianTarget2.TimerEnabled = 0
    GuardianTarget3.IsDropped = 1:GTL3.State = 0: GuardianTarget3.TimerEnabled = 0
    GuardianTarget4.IsDropped = 1:GTL4.State = 0: GuardianTarget4.TimerEnabled = 0
    GuardianTarget6.IsDropped = 1:GTL6.State = 0: GuardianTarget6.TimerEnabled = 0
    GuardianTarget7.IsDropped = 1:GTL7.State = 0: GuardianTarget7.TimerEnabled = 0
    GuardianTarget8.IsDropped = 1:GTL8.State = 0: GuardianTarget8.TimerEnabled = 0
    GuardianTarget9.IsDropped = 1:GTL9.State = 0: GuardianTarget9.TimerEnabled = 0
End Sub

Sub RiseGuardians
    PlaySoundAt SoundFXDOF("fx_resetdrop", 146, DOFPulse, DOFContactors), GuardianTarget2
    GuardianTarget1.IsDropped = 0:GTL1.State = 2: GuardianTarget1.TimerEnabled = 0
    GuardianTarget2.IsDropped = 0:GTL2.State = 2: GuardianTarget2.TimerEnabled = 0
    GuardianTarget3.IsDropped = 0:GTL3.State = 2: GuardianTarget3.TimerEnabled = 0
    GuardianTarget4.IsDropped = 0:GTL4.State = 2: GuardianTarget4.TimerEnabled = 0
    GuardianTarget6.IsDropped = 0:GTL6.State = 2: GuardianTarget6.TimerEnabled = 0
    GuardianTarget7.IsDropped = 0:GTL7.State = 2: GuardianTarget7.TimerEnabled = 0
    GuardianTarget8.IsDropped = 0:GTL8.State = 2: GuardianTarget8.TimerEnabled = 0
    GuardianTarget9.IsDropped = 0:GTL9.State = 2: GuardianTarget9.TimerEnabled = 0
End Sub

Sub CheckGuardiansDown
    If GTL1.State + GTL2.State + GTL3.State + GTL4.State + GTL6.State + GTL7.State + GTL8.State + GTL9.State = 0 Then
        WinGuardianMB
    End If
End Sub

' Left bank - Gold coins
Sub RiseLeftBank
    PlaySoundAt SoundFXDOF("fx_resetdrop", 145, DOFPulse, DOFContactors), Ldt2
    Ldt1.IsDropped = 0
    Ldt2.IsDropped = 0
    Ldt3.IsDropped = 0
    Ldt4.IsDropped = 0
    Light010.State = 0
    Light011.State = 0
    Light012.State = 0
    Light013.State = 0
    Loop3.Enabled = 0
    bGuardianMBReady = False
    GoldHits(CurrentPlayer, 0) = 0
    GoldHits(CurrentPlayer, 1) = 0
    GoldHits(CurrentPlayer, 2) = 0
    GoldHits(CurrentPlayer, 3) = 0
    GoldHits(CurrentPlayer, 4) = 0
End Sub

Sub DropLeftBank 'manually drop all targets
    PlaySoundAt "fx_droptarget_solenoid", Ldt2
    Ldt1.IsDropped = 1
    Ldt2.IsDropped = 1
    Ldt3.IsDropped = 1
    Ldt4.IsDropped = 1
    Light010.State = 1
    Light011.State = 1
    Light012.State = 1
    Light013.State = 1
    GoldHits(CurrentPlayer, 0) = 4
    Loop3.Enabled = 1
    bGuardianMBReady = True
    If NOT bMultiBallMode Then
        Light004.State = 2
    End If
End Sub

Sub UpdateLeftBank
    Ldt1.IsDropped = GoldHits(CurrentPlayer, 1)
    Ldt2.IsDropped = GoldHits(CurrentPlayer, 2)
    Ldt3.IsDropped = GoldHits(CurrentPlayer, 3)
    Ldt4.IsDropped = GoldHits(CurrentPlayer, 4)
    Light010.State = GoldHits(CurrentPlayer, 1)
    Light011.State = GoldHits(CurrentPlayer, 2)
    Light012.State = GoldHits(CurrentPlayer, 3)
    Light013.State = GoldHits(CurrentPlayer, 4)
    If Light013.State Then bGuardianMBReady = True:Light004.State = 2
End Sub

'Egypt target: right bank
Sub RiseRightBank
    PlaySoundAt SoundFXDOF("fx_resetdrop", 142, DOFPulse, DOFContactors), Rdt3
    Rdt1.IsDropped = 0
    Rdt2.IsDropped = 0
    Rdt3.IsDropped = 0
    Rdt4.IsDropped = 0
    Rdt5.IsDropped = 0
    Light005.State = 0
    Light006.State = 0
    Light007.State = 0
    Light008.State = 0
    Light009.State = 0
    EgyptLetters = "EGYPT"
    EgyptHits(CurrentPlayer, 0) = 0
    EgyptHits(CurrentPlayer, 1) = 0
    EgyptHits(CurrentPlayer, 2) = 0
    EgyptHits(CurrentPlayer, 3) = 0
    EgyptHits(CurrentPlayer, 4) = 0
    EgyptHits(CurrentPlayer, 5) = 0
    'ensure the guardian is down
    GTL5.State = 0
    GuardianTarget5.IsDropped = 1
End Sub

Sub UpdateRightBank
    Rdt1.IsDropped = EgyptHits(CurrentPlayer, 1)
    Rdt2.IsDropped = EgyptHits(CurrentPlayer, 2)
    Rdt3.IsDropped = EgyptHits(CurrentPlayer, 3)
    Rdt4.IsDropped = EgyptHits(CurrentPlayer, 4)
    Rdt5.IsDropped = EgyptHits(CurrentPlayer, 5)
    Light005.State = EgyptHits(CurrentPlayer, 5)
    Light006.State = EgyptHits(CurrentPlayer, 4)
    Light007.State = EgyptHits(CurrentPlayer, 3)
    Light008.State = EgyptHits(CurrentPlayer, 2)
    Light009.State = EgyptHits(CurrentPlayer, 1)
End Sub

' Magnet
Sub Magnet2_Hit 'trigger at the top of the quick sand to activate the magnet
    If Tilted Then Exit Sub
    LeftMagnet.MagnetOn = True
    vpmTimer.AddTimer 2000, " LeftMagnet.MagnetOn = False '"
    If bMultiBallMode Then
        PlaySound "s_quicksand"
    Else
        DMD "_", CL("QUICKSAND"), "_", eNone, eNone, eNone, 1000, True, "s_quicksand"
    End If
    AddScore 5
    'Rise the Guardian
    If GuardianTarget1.IsDropped Then
        GuardianTarget1.IsDropped = 0
        If NOT bGuardianMB Then GuardianTarget1.TimerEnabled = 1
        GTL1.State = 2
        PlaySoundAt SoundFXDOF("fx_resetdrop", 146, DOFPulse, DOFContactors), GuardianTarget1
    End If
End Sub

' Ramp Triggers

Sub AerialTrigger_Hit 'complete the ramp shot trigger to light CURSE
    If Tilted Then Exit Sub
    If bGuardianMB OR bPriestMB OR bPriestessMB OR bAtonMB Then
        DMD CL("SURVEY JACKPOT"), CL("25.000"), "", eNone, eBlinkFast, eNone, 1500, True, "s_JACKPOT_SURVEY"
        AddScore 25000
    ElseIf bRoamDesert Then
        If pinkarrow2.State Then
            DMD CL("HIDDEN PASS"), CL("33.000"), "", eNone, eBlinkFast, eNone, 1500, True, "s_hiddenpass"
            AddScore 33000
            pinkarrow2.State = 0
            pinkarrow1.State = 2
            HiddenPassHits = HiddenPassHits +1
            CheckHiddenPassHits
        Else
            DMD "_", CL("5.000"), "", eNone, eBlinkFast, eNone, 1500, True, "s_sandstorm"
            AddScore 5000
        End If
        Else 'normal scoring
        PlaySound "s_cursed whisper"
        LightSeqCurse.UpdateInterval = 25
        LightSeqCurse.Play SeqRandom, 50, , 1000
        AerialValue(CurrentPlayer) = AerialValue(CurrentPlayer) + 1
        Select Case AerialValue(CurrentPlayer)
            Case 0
            Case 1:Addscore 1500
            Case 2:AddScore 3000
            Case 3:AddScore 6000
            Case 4:AddScore 12000
            Case 5:AddScore 25000:bPriestMBReady = True
            Case Else
                If bPriestMBReady Then
                    AddScore 1500
                Else
                    bPriestMBReady = True
                End If
        End Select
        UpdateCURSE
    End If
    'combos
    If LastSwitchHit = "AerialTrigger" OR LastSwitchHit = "LidarTrigger" Then
        DMD "_", CL("COMBO"), "_", eNone, eBlink, eNone, 1000, True, "s_LOOP"
        ComboCount = ComboCount + 1
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "AerialTrigger"
End Sub

Sub UpdateCURSE 'curse lights
    If AerialValue(CurrentPlayer)> 5 Then AerialValue(CurrentPlayer) = 5
    Select Case AerialValue(CurrentPlayer)
        Case 0:cl1.State = 0:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 0
        Case 1:cl1.State = 1:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 0
        Case 2:cl1.State = 1:cl2.State = 1:cl3.State = 0:cl4.State = 0:cl5.State = 0
        Case 3:cl1.State = 1:cl2.State = 1:cl3.State = 1:cl4.State = 0:cl5.State = 0
        Case 4:cl1.State = 1:cl2.State = 1:cl3.State = 1:cl4.State = 1:cl5.State = 0
        Case 5:cl1.State = 1:cl2.State = 1:cl3.State = 1:cl4.State = 1:cl5.State = 1
            If NOT bMultiBallMode Then
                Light004.State = 2 'cobra vault light, Priest multiball is ready to start
                bPriestMBReady = True
            End If
    End Select
End Sub

Sub LidarTrigger_Hit 'complete the ramp shot trigger to light WOKEN
    If Tilted Then Exit Sub
    If bGuardianMB OR bPriestMB OR bPriestessMB OR bAtonMB Then
        DMD CL("LIDAR JACKPOT"), CL("25.000"), "", eNone, eBlinkFast, eNone, 1500, True, "s_JACKPOT_LIDAR"
        AddScore 25000
    ElseIf bRoamDesert Then
        If pinkarrow1.State Then
            DMD CL("HIDDEN PASS"), CL("33.000"), "", eNone, eBlinkFast, eNone, 1500, True, "s_hiddenpass"
            AddScore 33000
            pinkarrow1.State = 0
            pinkarrow2.State = 2
            HiddenPassHits = HiddenPassHits +1
            CheckHiddenPassHits
        Else
            DMD "_", CL("5.000"), "", eNone, eBlinkFast, eNone, 1500, True, "s_sandstorm"
            AddScore 5000
        End If
    Else 'normal scoring
        PlaySound "s_woken whisper"
        LightSeqWoken.UpdateInterval = 25
        LightSeqWoken.Play SeqRandom, 50, , 1000
        LidalValue(CurrentPlayer) = LidalValue(CurrentPlayer) + 1
        Select Case LidalValue(CurrentPlayer)
            Case 0
            Case 1:Addscore 1500
            Case 2:AddScore 3000
            Case 3:AddScore 6000
            Case 4:AddScore 12000
            Case 5:AddScore 25000:bPriestessMBReady = True
            Case Else
                If bPriestessMBReady Then
                    AddScore 1500
                Else
                    bPriestessMBReady = True
                End If
        End Select
        UpdateWOKEN
    End If
    'combos
    If LastSwitchHit = "AerialTrigger" OR LastSwitchHit = "LidarTrigger" Then
        DMD "_", CL("COMBO"), "_", eNone, eBlink, eNone, 1000, True, "s_LOOP"
        ComboCount = ComboCount + 1
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "LidarTrigger"
End Sub

Sub CheckHiddenPassHits
    If HiddenPassHits = 1 Then DMD "_", CL("2 MORE HITS"), "_", eNone, eBlink, eNone, 1000, True, ""
    If HiddenPassHits = 2 Then DMD "_", CL("1 MORE HIT"), "_", eNone, eBlink, eNone, 1000, True, ""
    If HiddenPassHits = 3 Then
        'Wizard mode
        DMD CL("STARTING"), CL("ROAM THE DESERT"), "d_border", eNone, eNone, eNone, 2000, True, ""
        DMD CL("RADIATION TARGETS"), CL("AND SUPER LOOPS"), "d_border", eNone, eNone, eNone, 2000, True, ""
        AddMultiball 1
        StartRoamDesert2
        ChangeGIIntensity 1
        ChangeGi white
    End If
End Sub

Sub UpdateWOKEN 'woken lights
    If LidalValue(CurrentPlayer)> 5 Then LidalValue(CurrentPlayer) = 5
    Select Case LidalValue(CurrentPlayer)
        Case 0:wl1.State = 0:wl2.State = 0:wl3.State = 0:wl4.State = 0:wl5.State = 0
        Case 1:wl1.State = 1:wl2.State = 0:wl3.State = 0:wl4.State = 0:wl5.State = 0
        Case 2:wl1.State = 1:wl2.State = 1:wl3.State = 0:wl4.State = 0:wl5.State = 0
        Case 3:wl1.State = 1:wl2.State = 1:wl3.State = 1:wl4.State = 0:wl5.State = 0
        Case 4:wl1.State = 1:wl2.State = 1:wl3.State = 1:wl4.State = 1:wl5.State = 0
        Case 5:wl1.State = 1:wl2.State = 1:wl3.State = 1:wl4.State = 1:wl5.State = 1
            If NOT bMultiBallMode Then
                Light004.State = 2 'cobra vault light, Priest multiball is ready to start
                bPriestessMBReady = True
            End If
    End Select
End Sub

'***************
' LOOP triggers
'***************

Sub MarketTrigger_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_sensor", MarketTrigger
    If bSkillshotReady Then
        AwardSuperSkillshot 50000
    End If
    Addscore 500
    ' check loops: only award when the ball is going down
    If LastSwitchHit = "Loop1" AND ActiveBall.VelY> 0 Then
        If bRoamDesert Then
            Addscore 25000
            DMD "SUPER LOOP", CL("25.000"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_super loops"
        Else
            If bMultiBallMode Then
                PlaySound "s_LOOP"
            Else
                DMD "_", CL("MARKET LOOP"), "_", eNone, eNone, eNone, 1000, True, "s_LOOP"
            End If
        End If
        If bRoamDesert Then
            LoopCount = LoopCount + 2
        Else
            LoopCount = LoopCount + 1
        End If
        FlashForMs Light003, 1200, 50, 0
        FlashForMs Light002, 1200, 50, 0
    ElseIf LastSwitchHit = "Loop6" AND ActiveBall.VelY> 0 Then
        If bRoamDesert Then
            Addscore 25000
            DMD "SUPER LOOP", CL("25.000"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_super loops"
        Else
            If bMultiBallMode Then
                PlaySound "s_LOOP"
            Else
                DMD "_", CL("SHORT CUT LOOP"), "_", eNone, eNone, eNone, 1000, True, "s_LOOP"
            End If
        End If
        If bRoamDesert Then
            LoopCount = LoopCount + 2
        Else
            LoopCount = LoopCount + 1
        End If
        FlashForMs Light003, 1200, 50, 0
        FlashForMs Light002, 1200, 50, 0
    ElseIf LastSwitchHit = "Loop2" AND ActiveBall.VelY> 0 Then
        If bRoamDesert Then
            Addscore 25000
            DMD "SUPER LOOP", CL("25.000"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_super loops"
        Else
            If bMultiBallMode Then
                PlaySound "s_LOOP"
            Else
                DMD "_", CL("BLACK MARKET LOOP"), "_", eNone, eNone, eNone, 1000, True, "s_LOOP"
            End If
        End If
        LoopCount = LoopCount + 1
        FlashForMs Light001, 1200, 50, 0
        FlashForMs Light002, 1200, 50, 0
    Else
        PlaySound "s_backmarketrollover", 0, 0.3
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "MarketTrigger"
End Sub

Sub Loop1_Hit 'city
    If Tilted Then Exit Sub
    ' check loops
    If LastSwitchHit = "Loop2" AND ActiveBall.VelY> 0 Then
        If bRoamDesert Then
            Addscore 25000
            DMD "SUPER LOOP", CL("25.000"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_super loops"
        Else
            If bMultiBallMode Then
                PlaySound "s_LOOP"
            Else
                DMD "_", CL("NILE LOOP"), "_", eNone, eNone, eNone, 1000, True, "s_LOOP"
            End If
        End If
        If bRoamDesert Then
            LoopCount = LoopCount + 2
        Else
            LoopCount = LoopCount + 1
        End If
        FlashForMs Light001, 1200, 50, 0
        FlashForMs Light002, 1200, 50, 0
    End If
    ' remember last trigger hit by the ball
    If Activeball.VelY <0 Then 'only set the name when going up due to loops being awarded when they shouldn't
        LastSwitchHit = "Loop1"
    End If
End Sub

Sub Loop2_Hit 'Nile
    If Tilted Then Exit Sub
    ' check loops
    If LastSwitchHit = "Loop1" AND ActiveBall.VelY> 0 Then
        If bRoamDesert Then
            Addscore 25000
            DMD "SUPER LOOP", CL("25.000"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_super loops"
        Else
            If bMultiBallMode Then
                PlaySound "s_LOOP"
            Else
                DMD "_", CL("CITY LOOP"), "_", eNone, eNone, eNone, 1000, True, "s_LOOP"
            End If
        End If
        If bRoamDesert Then
            LoopCount = LoopCount + 2
        Else
            LoopCount = LoopCount + 1
        End If
        FlashForMs Light001, 1200, 50, 0
        FlashForMs Light003, 1200, 50, 0
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Loop2"
End Sub

Sub Loop3_Hit 'Gold
    If Tilted Then Exit Sub
    ' check loops
    If bRoamDesert Then
        Addscore 25000
        DMD "SUPER LOOP", CL("25.000"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_super loops"
    Else
        If bMultiBallMode Then
            PlaySound "s_LOOP"
        Else
            DMD "_", CL("GOLD LOOP"), "_", eNone, eNone, eNone, 1000, True, "s_LOOP"
        End If
    End If
    If bRoamDesert Then
        LoopCount = LoopCount + 2
    Else
        LoopCount = LoopCount + 1
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Loop3"
End Sub

Sub Loop4_Hit 'Left Inlane
    If Tilted Then Exit Sub
    ' check loops
    If bLeftCheatDeath Then 'going up
        If bRoamDesert Then
            Addscore 25000
            DMD "SUPER LOOP", CL("25.000"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_super loops"
        Else
            If bMultiBallMode Then
                PlaySound "s_LOOP"
            Else
                DMD "_", CL("CHEAT DEATH LOOP"), "_", eNone, eNone, eNone, 1000, True, "s_LOOP"
            End If
        End If
        If bRoamDesert Then
            LoopCount = LoopCount + 2
        Else
            LoopCount = LoopCount + 1
        End If
        bLeftCheatDeath = False
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Loop4"
End Sub

Sub Loop5_Hit 'Right Inlane
    If Tilted Then Exit Sub
    ' check loops
    If bRightCheatDeath Then 'going up
        If bRoamDesert Then
            Addscore 25000
            DMD "SUPER LOOP", CL("25.000"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_super loops"
        Else
            If bMultiBallMode Then
                PlaySound "s_LOOP"
            Else
                DMD "_", CL("CHEAT DEATH LOOP"), "_", eNone, eNone, eNone, 1000, True, "s_LOOP"
            End If
        End If
        If bRoamDesert Then
            LoopCount = LoopCount + 2
        Else
            LoopCount = LoopCount + 1
        End If
        bRightCheatDeath = False
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Loop5"
End Sub

Sub Loop6_Hit 'on top of cobra vault
    If Tilted Then Exit Sub
    ' check loops
    If LastSwitchHit = "MarketTrigger" AND ActiveBall.VelY> 0 Then
        If bRoamDesert Then
            Addscore 25000
            DMD "SUPER LOOP", CL("25.000"), "_", eNone, eBlinkFast, eNone, 1000, True, "s_super loops"
        Else
            If bMultiBallMode Then
                PlaySound "s_LOOP"
            Else
                DMD "_", CL("SHORT CUT LOOP"), "_", eNone, eNone, eNone, 1000, True, "s_LOOP"
            End If
        End If
        If bRoamDesert Then
            LoopCount = LoopCount + 2
        Else
            LoopCount = LoopCount + 1
        End If
        FlashForMs Light003, 1200, 50, 0
        FlashForMs Light002, 1200, 50, 0
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Loop6"
End Sub

' MULTIBALLS

'Guardian multiball

Sub StartGuardianMB
    bGuardianMBReady = False
    bGuardianMB = True
    DMDFlush
    DMD " GUARDIAN MULTIBALL", CL("STARTING"), "d_border", eNone, eBlinkFast, eNone, 2000, True, ""
    vpmTimer.AddTimer 1000, "ChangeSong:AddMultiball 1 '"
    EnableBallSaver 20
    RiseGuardians
End Sub

Sub StopGuardianMB
    bGuardianMB = False
    DMDFlush
    DMD " GUARDIAN MULTIBALL", CL("NO ESCAPE FOR YOU"), "d_border", eNone, eBlinkFast, eNone, 2000, True, ""
    DropGuardians
    DropRadiTargets
    If EgyptHits(currentPlayer, 0) = 5 Then vpmtimer.Addtimer 1500, "RiseRightBank '"
    ChangeSong
    vpmTimer.AddTimer 1500, "RiseLeftBank '"
End Sub

Sub WinGuardianMB
    'bGuardianMB = False
    DMDFlush
    DMD CL("SUPER TAKEDOWN"), CL("450.000"), "d_border", eNone, eBlinkFast, eNone, 2000, True, "s_AM_SUPERTAKEDOWN"
    Addscore 450000
    'Wizard mode
    DMD CL("STARTING"), CL("ROAM THE DESERT"), "d_border", eNone, eNone, eNone, 2000, True, ""
    DMD CL("RADIATION TARGETS"), CL("AND SUPER LOOPS"), "d_border", eNone, eNone, eNone, 2000, True, ""
    vpmTimer.AddTimer 1500, "StartRoamDesert2 '"
End Sub

' Priest Multiball

Sub StartPriestMB
    bPriestMBReady = False
    bPriestMB = True
    PriestHits(CurrentPlayer) = 0
    RisePriest
    DMD CL("PRIEST MULTIBALL"), CL("STARTING"), "d_border", eNone, eBlinkFast, eNone, 2000, True, "s_PriestsSpeaks" &RndNbr(4)
    vpmTimer.AddTimer 1000, "AddMultiball 2 '"
    ChangeSong
    EnableBallSaver 20
End Sub

Sub StopPriestMB
    bPriestMB = False
    DMDFlush
    DMD CL("PRIEST MULTIBALL"), CL("I WILL BE BACK"), "d_border", eNone, eBlinkFast, eNone, 2000, True, ""
    DropPriest
    DropRadiTargets                       'in case they are up
    If PriestHits(CurrentPlayer) = 0 Then 'no hits Then reset Curse
        AerialValue(CurrentPlayer) = 0
        UpdateCURSE
    End If
    ChangeSong
End Sub

Sub WinPriestMB
    'bPriestMB = False
    DMDFlush
    DMD CL("PRIEST DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, "s_AM_PriestDefeated"
    DMD "", CL("PRIEST DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("PRIEST DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, ""
    DMD "", CL("PRIEST DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("PRIEST DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, ""
    DMD "", CL("PRIEST DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("PRIEST DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, ""
    DMD "", CL("PRIEST DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("PRIEST DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, ""
    DMD "", CL("PRIEST DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("KEEP SHOOTING"), CL("JACKPOTS"), "d_border", eNone, eNone, eNone, 2000, True, ""
    DropPriest
    'Wizard mode
    DMD CL("STARTING"), CL("ROAM THE DESERT"), "d_border", eNone, eNone, eNone, 2000, True, ""
    DMD CL("RADIATION TARGETS"), CL("AND SUPER LOOPS"), "d_border", eNone, eNone, eNone, 2000, True, ""
    vpmTimer.AddTimer 1500, "StartRoamDesert2 '"
'CheckCobraLight
End Sub

' Priestess Multiball

Sub StartPriestessMB
    bPriestessMBReady = False
    bPriestessMB = True
    PriestessHits(CurrentPlayer) = 0
    RisePRIESTESS
    DMD CL("PRIESTESS MULTIBALL"), CL("STARTING"), "d_border", eNone, eBlinkFast, eNone, 2000, True, "s_PriestessSpeaks" &RndNbr(4)
    vpmTimer.AddTimer 1000, "AddMultiball 2 '"
    ChangeSong
    EnableBallSaver 20
End Sub

Sub StopPriestessMB
    bPriestessMB = False
    DMDFlush
    DMD CL("PRIESTESS MULTIBALL"), CL("I WILL BE BACK"), "d_border", eNone, eBlinkFast, eNone, 2000, True, ""
    DropPRIESTESS
    DropRadiTargets                          'in case they are up
    If PriestessHits(CurrentPlayer) = 0 Then 'no hits Then reset Woken
        LidalValue(CurrentPlayer) = 0
        UpdateWOKEN
    End If
    ChangeSong
End Sub

Sub WinPriestessMB
    'bPriestessMB = False
    DMDFlush
    DMD CL("PRIESTESS DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, "s_AM_PriestessDefeated"
    DMD "", CL("PRIESTESS DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("PRIESTESS DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, ""
    DMD "", CL("PRIESTESS DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("PRIESTESS DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, ""
    DMD "", CL("PRIESTESS DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("PRIESTESS DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, ""
    DMD "", CL("PRIESTESS DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("PRIESTESS DEFEATED"), "", "d_border", eNone, eNone, eNone, 200, True, ""
    DMD "", CL("PRIESTESS DEFEATED"), "d_border", eNone, eNone, eNone, 200, True, ""
    DMD CL("KEEP SHOOTING"), CL("JACKPOTS"), "d_border", eNone, eNone, eNone, 2000, True, ""
    DropPriestess
    'Wizard mode
    DMD CL("STARTING"), CL("ROAM THE DESERT"), "d_border", eNone, eNone, eNone, 2000, True, ""
    DMD CL("RADIATION TARGETS"), CL("AND SUPER LOOPS"), "d_border", eNone, eNone, eNone, 2000, True, ""
    vpmTimer.AddTimer 1500, "StartRoamDesert2 '"
'CheckCobraLight
End Sub

' ATON Multiball

Sub StartAtonMB
    StartAtonLightShow
    bAtonMBReady = False
    bAtonMB = True
    bNorthLock = False
    bEastLock = False
    ForceLight037.State = 2
    ATONLight.State = 2
    atl1.State = 2
    atl2.State = 2
    atl3.State = 2
    atl4.State = 2
    atl5.State = 2
    atl6.State = 2
    atl7.State = 2
    atl8.State = 2
    DMDFlush
    DMD CL("ATON MULTIBALL"), CL("STARTING..."), "d_border", eNone, eNone, eNone, 1500, True, "s_ATONSPEECH" 'speech is 16 seconds
    DMD CL("ATON MULTIBALL"), CL("YOU SHOULD NOT HAVE"), "d_border", eNone, eNone, eNone, 1500, True, ""
    DMD CL("ATON MULTIBALL"), CL("COME HERE"), "d_border", eNone, eNone, eNone, 2000, True, ""
    DMD CL("ATON MULTIBALL"), "I TRIED TO WARN YOU", "d_border", eNone, eNone, eNone, 4000, True, ""
    DMD CL("ATON MULTIBALL"), "AND NOW YOU RE GONNA", "d_border", eNone, eNone, eNone, 1500, True, ""
    DMD CL("ATON MULTIBALL"), CL("DIE"), "d_border", eNone, eBlinkFast, eNone, 6000, True, ""
    vpmTimer.AddTimer 14000, "RiseAllAT '"
    vpmTimer.AddTimer 16658, "ChangeSong:lighteffect 11 '"
    vpmTimer.AddTimer 17000, "EnableBallSaver 30:EjectTopHole '"
    vpmTimer.AddTimer 18000, "EjectRightHole '"
    vpmTimer.AddTimer 19000, "AddMultiball 2 '"
End Sub

Sub StopAtonMB
    DMDFlush
    DMD CL("ATON MULTIBALL"), CL("I WILL RISE AGAIN"), "d_border", eNone, eBlinkFast, eNone, 2000, True, ""
    DropGuardians
    DropRadiTargets
    bAtonMB = False
    bAtonMB8 = False
    ResetAton
    CheckCobraLight
End Sub

Sub WinAtonMB
    DMDFlush
    DMD CL("ATON MULTIBALL"), CL("ATON DEFEATED"), "d_border", eNone, eBlinkFast, eNone, 2000, True, "s_AM_ATONDefeated"
    bAtonMB = False
    bAtonMB8 = False
    DropGuardians
    ResetAton
    'Wizard mode
    DMD CL("STARTING"), CL("ROAM THE DESERT"), "d_border", eNone, eNone, eNone, 2000, True, ""
    DMD CL("RADIATION TARGETS"), CL("AND SUPER LOOPS"), "d_border", eNone, eNone, eNone, 2000, True, ""
    vpmTimer.AddTimer 1500, "StartRoamDesert2 '"
End Sub

Sub ResetAton
    ChangeSong
    ForceLight037.State = 0
    ATONLight.State = 0
    DropAllAT
    RiseTombBank
    RiseTopBank
    atl1.State = 0
    atl2.State = 0
    atl3.State = 0
    atl4.State = 0
    atl5.State = 0
    atl6.State = 0
    atl7.State = 0
    atl8.State = 0
    bNorthLock = False
    bEastLock = False
    LockFlash1.State = 0
    LockFlash2.State = 0
End Sub

Sub CheckCobraLight
    If bGuardianMBReady OR bPriestMBReady OR bPriestessMBReady OR bAtonMBReady Then
        Light004.State = 2
    End If
End Sub

Sub StartAtonLightShow
    Dim i
    For each i in aAlienFlashers:i.State = 2:Next
    'stop the hands
    vpmTimer.AddTimer 12020, "Dim i: For each i in aAlienFlashers: i.State = 0: Next '"
End Sub

Sub StartAtonLightShowFaster
    Dim i
    For each i in aAlienFlashers:i.State = 2:Next
    'stop the hands
    vpmTimer.AddTimer 5000, "Dim i: For each i in aAlienFlashers: i.State = 0: Next '"
End Sub

' Radiation targets

Sub RiseRadiTargets
    PlaySoundAt "fx_resetdrop", Rad1p
    SetLightColor Light035, green, 2
    PriestessTlight.State = 2
    Rad1F.RotatetoEnd
    Rad1.IsDropped = 0
    Rad1Wall.IsDropped = 0
    SetLightColor Light036, green, 2
    PriestTlight.State = 2
    Rad2F.RotatetoEnd
    Rad2.IsDropped = 0
    Rad2Wall.IsDropped = 0
    BallSaverTimerExpired_Timer
    EnableBallSaver BallSaverTime
End Sub

Sub DropRadiTargets
    PlaySoundAt "fx_droptarget", Rad1p
    Light035.State = 0
    PriestessTlight.State = 0
    Rad1F.RotatetoStart
    Rad1.IsDropped = 1
    Rad1Wall.IsDropped = 1
    Light036.State = 0
    PriestTlight.State = 0
    Rad2F.RotatetoStart
    Rad2.IsDropped = 1
    Rad2Wall.IsDropped = 1
End Sub

Sub Rad1_Hit
    PlaySoundAt "fx_target", Rad1P
    If Tilted Then Exit Sub
    AddScore 25000
    DMDFlush
    DMD CL("RESTRICTED AREA"), CL("25.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_Radi_Hit"
End Sub

Sub Rad2_Hit
    PlaySoundAt "fx_target", Rad2P
    If Tilted Then Exit Sub
    AddScore 25000
    DMDFlush
    DMD CL("RESTRICTED AREA"), CL("25.000"), "_", eNone, eBlinkFast, eNone, 1500, True, "s_Radi_Hit"
End Sub

' Extra wall to make the ball flow better behind the right droptargets
Sub WallHelpTrigger_Hit
If ActiveBall.VelY > 0 Then 'ball is going down
    WallHelp.IsDropped = 0
End If
End Sub

Sub WallHelpTrigger_UnHit: WallHelp.IsDropped = 1: End Sub

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, BallsPerGame, UseFlexDMD, OldUseFlex, FlexDMDHighQuality, SongVolume
UseFlexDMD = False 'initialize variable
OldUseFlex = False

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Desktop DMD
    x = Table1.Option("DMD Type", 0, 1, 1, 1, 0, Array("Desktop DMD", "FlexDMD") )
    If UseFlexDMD AND x = 0 Then FlexDMD.Run = False
    If X then UseFlexDMD = True Else UseFlexDMD = False

    ' FlexDMD Quality
    x = Table1.Option("FlexDMD Quality", 0, 1, 1, 1, 0, Array("Low", "High") )
    If x Then FlexDMDHighQuality = True Else FlexDMDHighQuality = False
    If OldUseFlex <> UseFlexDMD Then
        DMD_Init
        If NOT bGameInPlay Then ShowTableInfo
        OldUseFlex = UseFlexDMD
    End If

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next

    ' Side Blades
    x = Table1.Option("Side Blades", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aSideBlades:y.SideVisible = x:next

    ' Balls per Game
    x = Table1.Option("Balls per Game", 0, 1, 1, 0, 0, Array("3 Balls", "5 Balls") )
    If x = 1 Then BallsPerGame = 5 Else BallsPerGame = 3

    ' FreePlay
    x = Table1.Option("Free Play", 0, 1, 1, 0, 0, Array("No", "Yes") )
    If x then bFreePlay = True Else bFreePlay = False

    ' Music  On/Off
    x = Table1.Option("Music", 0, 1, 1, 1, 0, Array("OFF", "ON") )
    If x Then bMusicOn = True Else bMusicOn = False

    ' Music Volume
    SongVolume = Table1.Option("Music Volume", 0, 1, 0.1, 0.3, 0)
    If bMusicOn Then
        PlaySound Song, -1, SongVolume, , , , 1, 0
    Else
        StopSound Song
    End If
End Sub

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

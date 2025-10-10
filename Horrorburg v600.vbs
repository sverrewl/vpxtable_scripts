' ****************************************************************
'                    Horrorburg
'          aka "NOT Welcome to Horrorburg"
'            for VISUAL PINBALL X 10.8
'         Uses FlexDMD for cabinet / FS mode
'            script by jpsalas - 2023
'         table based on an idea by MaxCore
'     table's theme colors based on Maxcore's cabinet
' ****************************************************************

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
Const cGameName = "horrorburg" 'used for DOF
Const myVersion = "6.00"
Const MaxPlayers = 4           ' from 1 to 4
Const MaxMultiplier = 2        ' limit playfield multiplier
Const MaxBonusMultiplier = 10  'limit Bonus multiplier
Const MaxMultiballs = 7        ' max number of balls during multiballs

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim PlayfieldMultiplier(4)
Dim PFxSeconds
Dim BallSaverTime ' in seconds of the first ball
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
Dim bAutoPlunger
Dim bInstantInfo
Dim bAttractMode
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
Dim bSkillShotSelect
Dim bExtraBallWonThisBall
Dim bJackpot

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger during multiballs
Dim cbRight

'*******************************************
'     Hauntfreak's directb2s effects
'*******************************************

Dim b2sstep
b2sstep = 0
'b2sflash.enabled = 0
Dim b2satm

Sub startB2S(aB2S)
    b2sflash.enabled = 1
    b2satm = ab2s
End Sub

Sub b2sflash_timer
    If B2SOn Then
        b2sstep = b2sstep + 1
        Select Case b2sstep
            Case 0
                Controller.B2SSetData b2satm, 0
            Case 1
                Controller.B2SSetData b2satm, 1
            Case 2
                Controller.B2SSetData b2satm, 0
            Case 3
                Controller.B2SSetData b2satm, 1
            Case 4
                Controller.B2SSetData b2satm, 0
            Case 5
                Controller.B2SSetData b2satm, 1
            Case 6
                Controller.B2SSetData b2satm, 0
            Case 7
                Controller.B2SSetData b2satm, 1
            Case 8
                Controller.B2SSetData b2satm, 0
                b2sstep = 0
                b2sflash.enabled = 0
        End Select
    End If
End Sub

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

    Set cbRight = New cvpmCaptiveBall
    With cbRight
        .InitCaptive CapTrigger1, CapWall1, Array(CapKicker1, CapKicker1a), 0
        .NailedBalls = 1
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbRight"
        .Start
    End With
    CapKicker1.CreateSizedBallWithMass BallSize / 2, BallMass

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
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bMultiBallStarted = False
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
    vpmtimer.addtimer 2000, "GiOn '"
    StartAttractMode
    StartFire
    ChangeGIIntensity 1 'default is 1
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

' Balloon animation
Dim MyPi, BalloonStep, BalloonDir
MyPi = Round(4 * Atn(1), 6) / 90
BalloonStep = 0

Sub Balloons_Timer()
    BalloonDir = SIN(BalloonStep * MyPi)
    BalloonStep = (BalloonStep + 1) MOD 360
    Balloon.RotY = - BalloonDir * 4
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    If bSkillShotSelect Then
        If keycode = LeftFlipperKey Then SkillshotType = 2:UpdateSkillShot
        If keycode = RightFlipperKey Then SkillshotType = 3:UpdateSkillShot
    End If

    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound "fx_nudge", 0, 1, 1, 0.25
    If keycode = MechanicalTilt Then CheckTilt

    If Keycode = AddCreditKey OR Keycode = AddCreditKey2 Then
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

        If keycode = LeftFlipperKey Then SolLFlipper 1:InstantInfoTimer.Enabled = True:RotateNunLeft:RotateKnivesLeft
        If keycode = RightFlipperKey Then SolRFlipper 1:InstantInfoTimer.Enabled = True:RotateNunRight:RotateKnivesRight

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
'test Keys
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If hsbModeActive Then
        Exit Sub
    End If

    If bSkillShotSelect Then
        If keycode = LeftFlipperKey Then SkillshotType = 1:UpdateSkillShot
        If keycode = RightFlipperKey Then SkillshotType = 1:UpdateSkillShot
    End If

    If keycode = PlungerKey Then
        Plunger.Fire
        PlaySoundAt "fx_plunger", plunger
        DOF 147, DOFpulse
    End If

    ' Table specific

    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then
            SolLFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
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

' test Keys
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
    PlaySoundAtBAll "fx_rubber_flipper"
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBall "fx_rubber_flipper"
End Sub

' Flippers top animation
Sub LeftFlipper_Animate()
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
End Sub

Sub RightFlipper_Animate()
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
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
        DMD "_", CL("CAREFUL"), "_", eNone, eBlinkFast, eNone, 1000, True, "vo_careful"
    End if
    If(NOT Tilted) AND Tilt> 15 Then 'If more that 15 then TILT the table
        'display Tilt
        InstantInfoTimer.Enabled = False
        DMDFlush
        DMD CL("YOU"), CL("TILTED"), "", eNone, eNone, eNone, 3000, True, "vo_youtilted"
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
        bMultiBallMode = False
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
        RightFlipper.RotateToStart
        Bumper001.Threshold = 100
        Bumper002.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        Tilted = False
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        Bumper001.Threshold = 1
        Bumper002.Threshold = 1
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        DMDFlush
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0) Then
        bMultiBallMode = False
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        If bRestorePower Then
            vpmtimer.Addtimer 4000, "EndOfBall() '"
            LightSeqFlashers.StopPlay
            bRestorePower = False
        Else
            vpmtimer.Addtimer 2000, "EndOfBall() '"
        End If
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
    If bGameInPlay Then
        PlaySong "m_main"
    Else
        PlaySong "m_gameover"
    End If
End Sub

Sub StopSong
    StopSound Song
End Sub

'********************
' Play random sounds
'********************

Sub PlaySfx
    PlaySound "sfx" &RndNbr(11)
End Sub

'**********************
'     GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'**********************

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

Sub GiOn
    PlaySoundAt "fx_GiOn", GiRelay 'about the center of the table
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    PlaySoundAt "fx_GiOff", GiRelay 'about the center of the table
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

Sub GiRedOn
    PlaySoundAt "fx_GiOn", GiRelay 'about the center of the table
    Dim bulb
    For each bulb in aGiLightsRED
        bulb.State = 1
    Next
End Sub

Sub GiRedOff
    PlaySoundAt "fx_GiOff", GiRelay 'about the center of the table
    Dim bulb
    For each bulb in aGiLightsRED
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
        Case 4 'center - used in the bonus count
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqCircleOutOn, 15, 1
        Case 5 'top down
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqDownOn, 15, 2
        Case 6 'down to top
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 15, 1
        Case 7
            LightSeqFlashers.UpdateInterval = 25
            LightSeqFlashers.Play SeqBlinking, , 15, 25
            PlaySound "sfx_thunder" &RndNbr(11)
    End Select
End Sub

' Fire lamps
Dim Fire1Pos, Fire2Pos, Fire3Pos, Fire4Pos, Fire5Pos, Fire6Pos, Flames
Flames = Array("fire01", "fire02", "fire03", "fire04", "fire05", "fire06", "fire07", "fire08", "fire09", _
    "fire10", "fire11", "fire12", "fire13", "fire14", "fire15", "fire16")

Sub StartFire
    Fire1Pos = 0
    Fire2Pos = 2
    Fire3Pos = 5
    Fire4Pos = 8
    Fire5Pos = 11
    Fire6Pos = 14
    FireTimer.Enabled = 1
End Sub

Sub FireTimer_Timer
    'debug.print fire1pos
    Fire1.ImageA = Flames(Fire1Pos)
    Fire2.ImageA = Flames(Fire2Pos)
    Fire3.ImageA = Flames(Fire3Pos)
    Fire4.ImageA = Flames(Fire4Pos)
    Fire5.ImageA = Flames(Fire5Pos)
    Fire6.ImageA = Flames(Fire6Pos)
    Fire1Pos = (Fire1Pos + 1) MOD 16
    Fire2Pos = (Fire2Pos + 1) MOD 16
    Fire3Pos = (Fire3Pos + 1) MOD 16
    Fire4Pos = (Fire4Pos + 1) MOD 16
    Fire5Pos = (Fire5Pos + 1) MOD 16
    Fire6Pos = (Fire6Pos + 1) MOD 16
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
    Pitch = BallVel(ball) * 50
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

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'******************************
'   JP's VPX Rolling Sounds
'******************************

Const tnob = 19   'total number of balls, 20 balls, from 0 to 19
Const lob = 2     'number of locked balls
Const maxvel = 45 'max ball velocity
ReDim rolling(tnob)

InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
    ' Start the ball rolling timer
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

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 2
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
Sub aDroptargets_Hit(idx):PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, Pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall):End Sub
Sub aTargets_Hit(idx):PlaySound SoundFX("fx_target", DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall):End Sub
Sub aTriggers_Hit(idx):PlaySoundAt "fx_sensor", aTriggers(idx):End Sub

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
    UpdateBallInPlay

    ' you may wish to start some music, play a sound, do whatever at this point
    ' PlaySound "vo_start" &RndNbr(3)

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

    ' set the current players bonus multiplier back down to 1X
    SetBonusMultiplier 1

    ' reduce the playfield multiplier by 1
    DecreasePlayfieldMultiplier

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
    bSkillShotSelect = True

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

Sub UpdateBallInPlay
    If B2SOn Then
        if BallsOnPlayfield = 0 then
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
    Dim AwardPoints, TotalBonus
    AwardPoints = 0
    TotalBonus = 0
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False
    GiOff
    ' only process any of this if the table is not tilted.
    '(the tilt recovery mechanism will handle any extra balls or end of game)

    If NOT Tilted Then
        PlaySong "m_bonus"
        'Count the bonus. This table uses several bonus
        DMD CL("BONUS"), "", "", eNone, eNone, eNone, 1000, True, ""
        AwardPoints = Switches * 300:TotalBonus = TotalBonus + AwardPoints
        DMD CL("SWITCH BONUS"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 1000, True, ""
        AwardPoints = Jumps(CurrentPlayer) * 10000:TotalBonus = TotalBonus + AwardPoints
        DMD CL("JUMP BONUS"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 1000, True, ""
        AwardPoints = Weapons(CurrentPlayer) * 50000:TotalBonus = TotalBonus + AwardPoints
        DMD CL("WEAPONS BONUS"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 1000, True, ""
        AwardPoints = HostagesRescued(CurrentPlayer) * 25000:TotalBonus = TotalBonus + AwardPoints
        DMD CL("HOSTAGES RESCUED"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 1000, True, ""
        TotalBonus = TotalBonus * BonusMultiplier(CurrentPlayer)
        DMD CL("TOTAL BONUS"), CL(FormatScore(TotalBonus) ), "", eNone, eBlinkFast, eNone, 2000, True, ""
        Score(CurrentPlayer) = Score(CurrentPlayer)

        ' add a bit of a delay to allow for the bonus points to be shown & added up
        vpmtimer.addtimer 7500, "EndOfBall2 '"
    Else 'if tilted then only add a short delay and move to the 2nd part of the end of the ball
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
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

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
        DMD CL("EXTRA BALL"), CL("SHOOT AGAIN"), "", eNone, eBlink, eNone, 1500, True, "vo_extraball"

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
            ' if multiple players are playing then move onto the next one
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
            DMD "", CL("PLAYER 1"), "", eNone, eNone, eNone, 1000, True, "vo_youareup"
        End If
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    'debug.print "End Of Game"
    bGameInPLay = False
    ' just ended your game then play the end of game tune
    PlaySound "vo_gameover"
    ChangeSong
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
                DMD "_", CL("BALL SAVED"), "_", eNone, eBlinkfast, eNone, 2500, True, "vo_ballsaved"
            'BallSaverTimerExpired_Timer 'uncomment the line if you want to stop the ballsaver
            End If
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1) Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True) then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ' you may wish to change any music over at this point and
                    ' changesong
                    ' turn off any multiball specific lights
                    'ChangeGIIntensity 1
                    'ChangeGi white
                    'stop any multiball modes of this game
                    StopMBmodes
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0) Then
                ' End Mode and timers
                'ChangeGIIntensity 1
                'ChangeGi white
                UpdateBallInPlay
                ' Show the end of ball animation
                ' and continue with the end of ball
                ' DMD something?
                StopEndOfBallMode
                vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, if there is no animation then move to the end of ball
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
    ' if the ball goes into the plunger lane during a multiball then activate the autoplunger
    If bMultiBallMode Then
        bAutoPlunger = True ' kick the ball in play if the bAutoPlunger flag is on
    End If
    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        vpmtimer.addtimer 1500, "PlungerIM.AutoFire:DOF 120, DOFPulse:PlaySoundAt ""fx_kicker"", swPlungerRest:bAutoPlunger = False '"
    End If
    'Start the skillshot lights & variables if any
    If bSkillShotReady Then
        ChangeSong
        SkillshotType = 1:UpdateSkillshot()
    ' show the message to shoot the ball in case the player has fallen sleep
    ' swPlungerRest.TimerEnabled = 1
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    lighteffect 6
    bBallInPlungerLane = False
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    bSkillShotSelect = False
    If bSkillShotReady Then
        ChangeSong
        ResetSkillShotTimer.Enabled = 0
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

' swPlungerRest timer to play a sound if the player has not shot the ball after a while
Sub swPlungerRest_Timer
    Dim i
    i = RndNbr(8) 'there are only 4 sounds in the table, so it will play a sound about 50% of times
    Select case i
        Case 1:PlaySound "vo_areyougoingtoplay"
        Case 2:PlaySound "vo_areyouplayingthisgame"
        Case 3:PlaySound "vo_pressthestartbutton"
        Case 4:PlaySound "vo_whatareyouwaitingfor"
    End Select
End Sub

Sub EnableBallSaver(seconds)
    ' do not start the timer if extra ball has been awarded
    If ExtraBallsAwards(CurrentPlayer)> 0 Then
        BallSaverTimerExpired.Enabled = False
        BallSaverSpeedUpTimer.Enabled = False
        LightShootAgain.State = 1
        Exit Sub
    End If
    'debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' stop the timers
    BallSaverTimerExpired.Enabled = False
    BallSaverSpeedUpTimer.Enabled = False
    ' restart the timers
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
    BallSaverSpeedUpTimer.Enabled = False 'ensure this timer is also stopped
    ' clear the flag
    bBallSaverActive = False
    ' if you have a ball saver light then turn it off at this point
    LightShootAgain.State = 0
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
End Sub

' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board

Sub AddScore(points) 'normal score routine; points x playfieldmultiplier
    If Tilted Then Exit Sub
    If bSkillshotReady Then ResetSkillShotTimer_Timer
    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer)
    If Mode <> 0 Then ModeScore = ModeScore + points * PlayfieldMultiplier(CurrentPlayer)
' you may wish to check to see if the player has gotten a replay
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

    If(bMultiBallMode = True) Then
        Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + points
        ' DMD "_", CL("INCREASED JACKPOT"), "_", eNone, eNone, eNone, 1000, True, ""
        ' you may wish to limit the jackpot to a upper limit, ie..
        If(Jackpot(CurrentPlayer) >= 1000000) Then
            Jackpot(CurrentPlayer) = 1000000
        End if
    End if
End Sub

Sub AddSuperJackpot(points)
    If Tilted Then Exit Sub
    If(bMultiBallMode = True) Then
        SuperJackpot(CurrentPlayer) = SuperJackpot(CurrentPlayer) + points
        ' DMD "_", "INCREASED SP.JACKPOT", "_", eNone, eNone, eNone, 1000, True, ""
        ' you may wish to limit the jackpot to a upper limit, ie..
        If(SuperJackpot(CurrentPlayer) >= 9000000) Then
            SuperJackpot(CurrentPlayer) = 9000000
        End if
    End if
End Sub

Sub AddBonusMultiplier(n) 'adapted to this table
    Knives(0) = Knives(0) + 1
    Select Case Knives(0)
        Case 1:SetBonusMultiplier 2
        Case 2:SetBonusMultiplier 3
        Case 3:SetBonusMultiplier 5
        Case 4:SetBonusMultiplier 7
        Case 5:SetBonusMultiplier 8
        Case 6:SetBonusMultiplier 10
        Case Else
            AddScore 50000
            DMD "_", CL("50.000 POINTS"), "_", eNone, eBlink, eNone, 1000, True, ""
    End Select
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly

Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level
    UpdateBonusXLights(Level)
    If level> 1 Then
        DMD "_", CL("BONUS X " &Level), "_", eNone, eBlink, eNone, 2000, True, ""
        GiEffect 1
    End If
End Sub

Sub UpdateBonusXLights(Level) '4 lights in this table, from 2x to 5x
    ' Update the lights
    Select Case Level
        Case 1:Light060.State = 0:Light059.State = 0:Light061.State = 0
        Case 2:Light060.State = 1:Light059.State = 0:Light061.State = 0
        Case 3:Light060.State = 0:Light059.State = 1:Light061.State = 0
        Case 5:Light060.State = 0:Light059.State = 0:Light061.State = 1
        Case 7:Light060.State = 1:Light059.State = 0:Light061.State = 1
        Case 8:Light060.State = 0:Light059.State = 1:Light061.State = 1
        Case 10:Light060.State = 1:Light059.State = 1:Light061.State = 1
    End Select
End Sub

Sub AddPlayfieldMultiplier(n)
    Dim NewPFLevel
    ' if not at the maximum level x
    if(PlayfieldMultiplier(CurrentPlayer) + n <= MaxMultiplier) then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) + n
        SetPlayfieldMultiplier(NewPFLevel)
        DMD "_", CL("PLAYFIELD X " &NewPFLevel), "_", eNone, eBlink, eNone, 2000, True, ""
        GiEffect 1
    Else 'if the max is already lit
        AddScore 50000
        DMD "_", CL("50.000 POINTS"), "_", eNone, eBlink, eNone, 2000, True, ""
    End if
    ' restart the PlayfieldMultiplier timer in case it was already started
    PFXTimer.Enabled = 0
    PFXTimer.Enabled = 1
    PFXTimerSpeedUp.Enabled = 0
    PFXTimerSpeedUp.Enabled = 1
End Sub

Sub PFXTimer_Timer
    DecreasePlayfieldMultiplier
End Sub

Sub PFXTimerSpeedUp_Timer 'speed up the blink light for the last 10 seconds
    Light058.BlinkInterval = 200:Light058.State = 2
    PFXTimerSpeedUp.Enabled = 0
End Sub

Sub DecreasePlayfieldMultiplier 'reduces by 1 the playfield multiplier, this will stop the timer as this table only has a 2x multiplier
    Dim NewPFLevel
    ' if not at 1 already
    if(PlayfieldMultiplier(CurrentPlayer)> 1) then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) - 1
        SetPlayfieldMultiplier(NewPFLevel)
        PFXTimer.Enabled = 0
        PFXTimer.Enabled = 1
        PFXTimerSpeedUp.Enabled = 0
        PFXTimerSpeedUp.Enabled = 1
    Else
        PFXTimer.Enabled = 0
        PFXTimerSpeedUp.Enabled = 0
    End if
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
        Case 1:Light058.State = 0
        Case 2:Light058.BlinkInterval = 400:Light058.State = 2
    End Select
' perhaps show also the multiplier in the DMD?
End Sub

Sub ExtraBallIsLit
    If Light048.State = 0 Then
        DMD "_", CL("EXTRA BALL IS LIT"), "", eNone, eNone, eNone, 1500, True, "vo_extraballislit"
        Light048.State = 1
        XtraBalisLit(CurrentPlayer) = 1
    End If
End Sub

Sub AwardExtraBall()
    '   If NOT bExtraBallWonThisBall Then 'uncomment this If in case you want to give just one extra ball per ball
    DMD "_", CL("EXTRA BALL WON"), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
    DOF 121, DOFPulse
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    ApronDMDUpdate
    '    bExtraBallWonThisBall = True
    LightShootAgain.State = 1 'light the shoot again lamp
    GiEffect 3
    LightEffect 2
'    END If
End Sub

Sub AwardSpecial()
    DMD "_", CL("EXTRA GAME WON"), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
    DOF 121, DOFPulse
    Credits = Credits + 1
    If bFreePlay = False Then DOF 125, DOFOn
    LightEffect 2
    GiEffect 3
End Sub

Sub AwardJackpot()
    DMDFlush
    DMD CL("JACKPOT"), CL(FormatScore(Jackpot(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 1500, True, "vo_Jackpot"
    DOF 126, DOFPulse
    AddScore Jackpot(CurrentPlayer)
    LightEffect 2
    GiEffect 3
    ' modes handling
    Select Case Mode
        Case 4 'Dracula MB 'after 5 jackpots turn on the super jackpot light
            JackpotCount = JackpotCount + 1
            If JackpotCount >= 5 Then
                Light077.State = 2
            End If
    End Select
End Sub

Sub AwardSuperJackpot()
    DMDflush
    SuperJackpot(CurrentPlayer) = Jackpot(CurrentPlayer) * JackpotCount '250.000 or more
    DMD CL("SUPER JACKPOT"), CL(FormatScore(SuperJackpot(CurrentPlayer) ) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_superjackpot"
    DOF 126, DOFPulse
    AddScore SuperJackpot(CurrentPlayer)
    LightEffect 2
    GiEffect 3
End Sub

Sub AwardSkillshot(points)
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL("SKILLSHOT"), CL(FormatScore(points) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_skillshot"
    DOF 127, DOFPulse
    AddScore points
    'do some light show
    GiEffect 3
    LightEffect 2
End Sub

Sub AwardSuperSkillshot(points)
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL("SUPER SKILLSHOT"), CL(FormatScore(points) ), "d_border", eNone, eBlinkFast, eNone, 2000, True, "vo_superskillshot"
    DOF 127, DOFPulse
    AddScore points
    'do some light show
    GiEffect 3
    LightEffect 2
End Sub

'**************
'   COMBOS
'**************

Sub AwardCombo
    DOF 128, DOFPulse 'Combo
    ComboCount = ComboCount + 1
    Select Case ComboCount
        Case 1:DMD CL("COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 1500, True, "vo_combo"
        Case 2:DMD CL("2X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 2) ), "", eNone, eNone, eNone, 1500, True, "vo_doublecombo"
        Case 3:DMD CL("3X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 3) ), "", eNone, eNone, eNone, 1500, True, "vo_triplecombo"
        Case 4:DMD CL("4X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 4) ), "", eNone, eNone, eNone, 1500, True, "vo_supercombo"
        Case 5:DMD CL("5X COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 5) ), "", eNone, eNone, eNone, 1500, True, "vo_supercombo"
        Case Else:DMD CL("SUPER COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * ComboCount) ), "", eNone, eNone, eNone, 1500, True, "vo_supercombo"
    End Select
    AddScore ComboValue(CurrentPlayer) * ComboCount
End Sub

'*****************************
'    Load / Save / Highscore
'*****************************

Dim MyTable
MyTable = "horrorburg"

Sub Loadhs
    Dim x
    x = LoadValue(MyTable, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 100000 End If
    x = LoadValue(MyTable, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If
    x = LoadValue(MyTable, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 100000 End If
    x = LoadValue(MyTable, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If
    x = LoadValue(MyTable, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 100000 End If
    x = LoadValue(MyTable, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If
    x = LoadValue(MyTable, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 100000 End If
    x = LoadValue(MyTable, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If
    x = LoadValue(MyTable, "Credits")
    If(x <> "") then Credits = CInt(x) Else Credits = 0:If bFreePlay = False Then DOF 125, DOFOff:End If
    x = LoadValue(MyTable, "TotalGamesPlayed")
    If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If
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
        vpmTimer.AddTimer 2000, "PlaySound ""vo_taunt"" &RndNbr(9) '"
        EndOfBallComplete()
    End If
End Sub

Sub HighScoreEntryInit()
    Dim tmp
    tmp = RndNbr(3)
    Select Case tmp
        Case 1:PlaySound "vo_nicescore"
        Case 2:PlaySound "vo_enterinitials"
        Case 3:playSound "vo_excellentscore"
    End Select
    hsbModeActive = True
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
        playsound "sfx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        playsound "sfx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<") then
            playsound "sfx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "sfx_Esc"
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
    Savehs
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
        If bRestorePowerReady OR bEscapeHWReady Then tmp1 = "  SHOOT THE SCOOP"
        Select Case Mode
            Case 0 'no Mode active
            Case 1:tmp1 = "   RESTORE POWER"
            Case 2:tmp1 = " ESCAPE HORRORBURG"
            Case 3:tmp1 = "PENNYWISE MULTIBALL"
            Case 4:tmp1 = " DRACULA MULTIBALL"
        End Select
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
                        If i = 2 then
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
                        If i = 2 then
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
        if IsNumeric(mid(NumString, i, 1) ) then
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
        digit041, digit042, digit043, digit044, digit045, digit046)
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
    Chars(42) = "d_star"  '*
    'Chars(43) = ""        '+
    Chars(44) = "d_comma" ',
    Chars(45) = "d_minus" '-
    Chars(46) = "d_dot"   '.
    'Chars(47) = ""        '/
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
    'Chars(61) = ""        '=
    Chars(62) = "d_more"  '>
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
    'Chars(94) = ""        '^
    'Chars(95) = "" '_
    'Chars(96) = ""
    'Chars(97) = ""  'a
    'Chars(98) = ""  'b
    'Chars(99) = ""  'c
    'Chars(100) = "" 'd
    'Chars(101) = "" 'e
    'Chars(102) = "" 'f
    'Chars(103) = "" 'g
    'Chars(104) = "" 'h
    'Chars(105) = "" 'i
    'Chars(106) = "" 'j
    'Chars(107) = "" 'k
    'Chars(108) = "" 'l
    'Chars(109) = "" 'm
    'Chars(110) = "" 'n
    'Chars(111) = "" 'o
    'Chars(112) = "" 'p
    'Chars(113) = "" 'q
    'Chars(114) = "" 'r
    'Chars(115) = "" 's
    'Chars(116) = "" 't
    'Chars(117) = "" 'u
    'Chars(118) = "" 'v
    'Chars(119) = "" 'w
    'Chars(120) = "" 'x
    'Chars(121) = "" 'y
    'Chars(122) = "" 'z
    'Chars(123) = "" '{
    'Chars(124) = "" '|
    'Chars(125) = "" '}
    'Chars(126) = "" '~
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
    DMD "         JPSALAS", "         PRESENTS", "d_jpsalas", eNone, eNone, eNone, 3000, False, ""
    DMD CL("HORRORBURG"), CL("ROM VERSION " &myversion), "", eNone, eNone, eNone, 4000, False, ""
    DMD CL("HIGHSCORES"), Space(20), "", eScrollLeft, eScrollLeft, eNone, 20, False, ""
    DMD CL("HIGHSCORES"), "", "", eBlinkFast, eNone, eNone, 1000, False, ""
    DMD CL("HIGHSCORES"), "1> " &HighScoreName(0) & " " &FormatScore(HighScore(0) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "2> " &HighScoreName(1) & " " &FormatScore(HighScore(1) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "3> " &HighScoreName(2) & " " &FormatScore(HighScore(2) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "4> " &HighScoreName(3) & " " &FormatScore(HighScore(3) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD Space(20), Space(20), "", eScrollLeft, eScrollLeft, eNone, 1000, False, ""
    DMD "     HOW TO PLAY    ", "                    ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD " HIT THE MAIN SHOTS ", "TO CAPTURE 5 KILLERS", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD " CAPTURE 2 KILLERS  ", "TO LIGHT EXTRA BALL ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD "CAPTURE 5 KILLERS TO", "START RESTORE POWER ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD "  HIT ANNABELLE TO  ", "  START A HURRY UP  ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD "DESTROY THE TOMBS TO", " DRACULA MULTI-BALL ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD " HIT PENNYWISE FOR  ", "PENNYWISE MULTI-BALL", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD " HIT POP BUMPERS TO ", " BUILD CHUCKY VALUE ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD " HIT CHUCKY TARGET  ", "     TO COLLECT     ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD "   HITTING JIGSAW   ", " RELEASES HOSTAGES  ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD "COLLECT THE 4 KNIFES", "FOR BONUS MULTIPLIER", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD " COMPLETE NUN LANES ", " FOR DOUBLE SCORING ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD " RESCUE 25 HOSTAGES ", "TO ESCAPE HORRORBURG", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD "   SPELL CHAOS TO   ", "  COLLECT WEAPONS   ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD "  WEAPONS INCREASE  ", "SCORE IN MULTIBALLS ", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD "", "", "", eScrollLeft, eScrollLeft, eNone, 2000, False, ""
    DMD Space(20), Space(20), "", eScrollLeft, eScrollLeft, eNone, 1000, False, ""
End Sub

Sub StartAttractMode
    StartLightSeq
    DMDFlush
    ShowTableInfo
    PlaySong "m_gameover"
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
End Sub

' tables variables and Mode init
Dim BalloonsLeft(4)
Dim HostagesLeft(4)
Dim HostagesRescued(4)
Dim HostagesLights(4, 5) '5 Lights
Dim ChaosLights(4, 5)    '5 lights
Dim KillerHits(4, 5)     '5 killers
Dim KillersCompleted(4)
Dim SkillshotType
Dim Nun(3)             'top lanes
Dim Knives(4)          '4 inlanes & outlanes knifes Knives(0) contains the current bonus multiplier
Dim Weapons(4)         'weapons collected for each player, chaos awards weapons
Dim ChuckyValue(4)     'players Chucky value, it is increased by the bumpers
Dim ChainSawHits(4)    'players chainsaw hits
Dim JigSawHits(4)      'players jigsaw hits
Dim DT(4, 3)           'droptargets state
Dim Switches           'nr of switches hit by each ball
Dim Jumps(4)           'total number of ramp jumps for each player
Dim AnnabelleHits(4)   'captive ball hits for each player
Dim LeatherfaceHits(4) 'Leatherface hits
Dim RPlights(11)       'number of hits on each Power line
Dim ModeScore          'points earned during the Restore Power
Dim Mode               'Current Wizard and multiball modes
Dim JackpotLights(5)   'the state of the jackpot Lights
Dim JackpotCount       'count the jackpots to enable the super jackpot
Dim SpinnerHits(4)
Dim XtraBalisLit(4)

Dim bRestorePower
Dim bRestorePowerReady
Dim bEscapeHW
Dim bEscapeHWReady
Dim bPennywise
Dim bDracula
' Modes variables

Sub Game_Init() 'called at the start of a new game
    Dim i, j
    ' play a welcome Sound
    ' PLaySound "Start" &RndNbr(10)
    For i = 0 to 4
        Jackpot(i) = 50000
        SuperJackpot(i) = 250000
        BalloonsLeft(i) = 50
        HostagesLeft(i) = 25
        HostagesRescued(i) = 0
        Knives(i) = 0
        Weapons(i) = 0
        ChuckyValue(i) = 1000
        ChainSawHits(i) = 0
        JigSawHits(i) = 0
        Jumps(i) = 0
        AnnabelleHits(i) = 0
        LeatherfaceHits(i) = 0
        SpinnerHits(i) = 0
        ComboHits(i) = 0
        ComboValue(i) = 100000
        KillersCompleted(i) = 0
        XtraBalisLit(i) = 0
        BallsInLock(i) = 0
        For j = 0 to 5
            HostagesLights(i, j) = 0
            ChaosLights(i, j) = 0
            KillerHits(i, j) = 0
            JackpotLights(j) = 0
        Next
        For j = 0 to 3
            DT(i, j) = 0
        Next
    Next
    ' set the first chaos Light for each player
    For i = 0 to 4
        ChaosLights(i, 1) = 2
    Next
    ReleaseHostage 'release 1 hostage
    SkillshotType = 1
    BallSaverTime = 20
    bRestorePower = False
    bRestorePowerReady = False
    bEscapeHW = False
    bEscapeHWReady = False
    bPennywise = False
    bDracula = False
    Mode = 0
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
    DMD CL("CHUCKY VALUE"), CL(FormatScore(ChuckyValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("JACKPOT VALUE"), CL(FormatScore(Jackpot(CurrentPlayer) ) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("SPINS 2 UP WEAPON"), CL(250-SpinnerHits(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("WEAPONS COLLECTED"), CL(Weapons(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("JUMPS COMPLETED"), CL(Jumps(CurrentPlayer) ), "", eNone, eNone, eNone, 2000, False, ""
End Sub

Sub StopMBmodes 'stop multiball modes after loosing the last multiball
    If bEscapeHW Then StopEscapeHW
    If bPennywise Then StopPennywiseMB
    If bDracula Then StopDraculaMB
End Sub

Sub StopEndOfBallMode() 'this sub is called after the last ball in play is drained, reset skillshot, modes, timers
    ResetSkillShotTimer_Timer
    DecreasePlayfieldMultiplier
    If bRestorePowerReady Then
        LightSeqFlashers.StopPlay
        bRestorePowerReady = False
    End If
End Sub

Sub ResetNewBallVariables() 'reset variables and lights for a new ball or player
    Dim i
    'turn on or off the needed lights before a new ball is released
    TurnOffPlayfieldLights
    'set up the lights according to the player achievments
    UpdateLights 'chaos and hostages lights
    UpdateKillerLights
    ApronDMDUpdate
    UpdateJigSaw
    'reset NUN variables
    For i = 0 to 3
        Nun(i) = 0
    Next
    'reset knives and bonus multiplier
    For i = 0 to 4
        Knives(i) = 0
    Next
    SetBonusMultiplier 1
    If AnnabelleHits(CurrentPlayer) = 9 Then AnnabelleHits(CurrentPlayer) = 8 'you need an extra hit to activate the jackpot
    Switches = 0
    Mode = 0
    ResetDT 'reset Dracula & drop targets
    ComboCount = 0
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub UpdateSkillShot() 'Setup and updates the skillshot lights
    LightSeqTilt.Play SeqAlloff
    LightSeqSkillshot.StopPlay
    LightSeqSkillshot2.StopPlay
    LightSeqSkillshot3.StopPlay
    Select Case SkillshotType
        Case 1:
            LightSeqSkillshot.PLay SeqBlinking, , 50, 300
            DMD CL("HIT LIT LIGHT"), CL("FOR SKILLSHOT"), "", eNone, eNone, eNone, 3000, True, ""
        Case 2:
            LeftGate.Open = True
            LightSeqSkillshot2.PLay SeqBlinking, , 50, 300
            DMD CL("HIT THE RAMP"), CL("FOR SUPERSKILLSHOT"), "", eNone, eNone, eNone, 3000, True, ""
        Case 3:
            LeftGate.Open = True
            LightSeqSkillshot3.PLay SeqBlinking, , 50, 300
            DMD CL("HIT PENNYWISE"), CL("FOR SUPERSKILLSHOT"), "", eNone, eNone, eNone, 3000, True, ""
    End Select
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    bSkillShotSelect = False
    LeftGate.Open = False
    LightSeqTilt.StopPLay
    LightSeqSkillshot.StopPlay
    LightSeqSkillshot2.StopPlay
    LightSeqSkillshot3.StopPlay
    DMDScoreNow
End Sub

Sub UpdateLights 'chaos and hostages lights
    Light048.State = XtraBalisLit(CurrentPlayer)
    UpdateHLights
    UpdateCLights
End Sub

Sub UpdateHLights 'hostages lights for the current player
    'Hostages
    Light071.State = HostagesLights(CurrentPlayer, 1)
    Light072.State = HostagesLights(CurrentPlayer, 2)
    Light073.State = HostagesLights(CurrentPlayer, 3)
    Light074.State = HostagesLights(CurrentPlayer, 4)
    Light075.State = HostagesLights(CurrentPlayer, 5)
End Sub

Sub UpdateCLights 'chaos lights for the current player
    'Chaos
    Light070.State = ChaosLights(CurrentPlayer, 1)
    Light069.State = ChaosLights(CurrentPlayer, 2)
    Light068.State = ChaosLights(CurrentPlayer, 3)
    Light067.State = ChaosLights(CurrentPlayer, 4)
    Light066.State = ChaosLights(CurrentPlayer, 5)
End Sub

Sub UpdateKillerLights
    Select Case KillerHits(CurrentPlayer, 1) 'Jason
        Case 0:Light053.State = 0:Light003.State = 0:Light004.State = 0:Light005.State = 0:Light006.State = 0:Light007.State = 0:Light008.State = 0:Light082.State = 0
        Case 1:Light053.State = 0:Light003.State = 1:Light004.State = 1:Light005.State = 0:Light006.State = 0:Light007.State = 0:Light008.State = 0:Light082.State = 0
        Case 2:Light053.State = 0:Light003.State = 1:Light004.State = 1:Light005.State = 1:Light006.State = 1:Light007.State = 0:Light008.State = 0:Light082.State = 0
        Case 3:Light053.State = 2:Light003.State = 0:Light004.State = 0:Light005.State = 0:Light006.State = 0:Light007.State = 0:Light008.State = 0:Light082.State = 0
        Case 4:Light053.State = 2:Light003.State = 1:Light004.State = 0:Light005.State = 0:Light006.State = 0:Light007.State = 0:Light008.State = 0:Light082.State = 2
        Case 5:Light053.State = 2:Light003.State = 1:Light004.State = 1:Light005.State = 0:Light006.State = 0:Light007.State = 0:Light008.State = 0:Light082.State = 2
        Case 6:Light053.State = 2:Light003.State = 1:Light004.State = 1:Light005.State = 1:Light006.State = 0:Light007.State = 0:Light008.State = 0:Light082.State = 2
        Case 7:Light053.State = 2:Light003.State = 1:Light004.State = 1:Light005.State = 1:Light006.State = 1:Light007.State = 0:Light008.State = 0:Light082.State = 2
        Case 8:Light053.State = 2:Light003.State = 1:Light004.State = 1:Light005.State = 1:Light006.State = 1:Light007.State = 1:Light008.State = 0:Light082.State = 2
        Case 9:Light053.State = 1:Light003.State = 1:Light004.State = 1:Light005.State = 1:Light006.State = 1:Light007.State = 1:Light008.State = 1:Light082.State = 0
        Case 10:KillerHits(CurrentPlayer, 1) = 9
    End Select
    Select Case KillerHits(CurrentPlayer, 2) 'Ghostface
        Case 0:Light056.State = 0:Light022.State = 0:Light023.State = 0:Light024.State = 0:Light025.State = 0:Light026.State = 0:Light021.State = 0:Light081.State = 0
        Case 1:Light056.State = 0:Light022.State = 1:Light023.State = 1:Light024.State = 0:Light025.State = 0:Light026.State = 0:Light021.State = 0:Light081.State = 0
        Case 2:Light056.State = 0:Light022.State = 1:Light023.State = 1:Light024.State = 1:Light025.State = 1:Light026.State = 0:Light021.State = 0:Light081.State = 0
        Case 3:Light056.State = 2:Light022.State = 0:Light023.State = 0:Light024.State = 0:Light025.State = 0:Light026.State = 0:Light021.State = 0:Light081.State = 0
        Case 4:Light056.State = 2:Light022.State = 1:Light023.State = 0:Light024.State = 0:Light025.State = 0:Light026.State = 0:Light021.State = 0:Light081.State = 2
        Case 5:Light056.State = 2:Light022.State = 1:Light023.State = 1:Light024.State = 0:Light025.State = 0:Light026.State = 0:Light021.State = 0:Light081.State = 2
        Case 6:Light056.State = 2:Light022.State = 1:Light023.State = 1:Light024.State = 1:Light025.State = 0:Light026.State = 0:Light021.State = 0:Light081.State = 2
        Case 7:Light056.State = 2:Light022.State = 1:Light023.State = 1:Light024.State = 1:Light025.State = 1:Light026.State = 0:Light021.State = 0:Light081.State = 2
        Case 8:Light056.State = 2:Light022.State = 1:Light023.State = 1:Light024.State = 1:Light025.State = 1:Light026.State = 1:Light021.State = 0:Light081.State = 2
        Case 9:Light056.State = 1:Light022.State = 1:Light023.State = 1:Light024.State = 1:Light025.State = 1:Light026.State = 1:Light021.State = 1:Light081.State = 0
        Case 10:KillerHits(CurrentPlayer, 2) = 9
    End Select
    Select Case KillerHits(CurrentPlayer, 3) 'Freddy
        Case 0:Light057.State = 0:Light028.State = 0:Light029.State = 0:Light030.State = 0:Light031.State = 0:Light032.State = 0:Light027.State = 0:Light080.State = 0
        Case 1:Light057.State = 0:Light028.State = 1:Light029.State = 1:Light030.State = 0:Light031.State = 0:Light032.State = 0:Light027.State = 0:Light080.State = 0
        Case 2:Light057.State = 0:Light028.State = 1:Light029.State = 1:Light030.State = 1:Light031.State = 1:Light032.State = 0:Light027.State = 0:Light080.State = 0
        Case 3:Light057.State = 2:Light028.State = 0:Light029.State = 0:Light030.State = 0:Light031.State = 0:Light032.State = 0:Light027.State = 0:Light080.State = 0
        Case 4:Light057.State = 2:Light028.State = 1:Light029.State = 0:Light030.State = 0:Light031.State = 0:Light032.State = 0:Light027.State = 0:Light080.State = 2
        Case 5:Light057.State = 2:Light028.State = 1:Light029.State = 1:Light030.State = 0:Light031.State = 0:Light032.State = 0:Light027.State = 0:Light080.State = 2
        Case 6:Light057.State = 2:Light028.State = 1:Light029.State = 1:Light030.State = 1:Light031.State = 0:Light032.State = 0:Light027.State = 0:Light080.State = 2
        Case 7:Light057.State = 2:Light028.State = 1:Light029.State = 1:Light030.State = 1:Light031.State = 1:Light032.State = 0:Light027.State = 0:Light080.State = 2
        Case 8:Light057.State = 2:Light028.State = 1:Light029.State = 1:Light030.State = 1:Light031.State = 1:Light032.State = 1:Light027.State = 0:Light080.State = 2
        Case 9:Light057.State = 1:Light028.State = 1:Light029.State = 1:Light030.State = 1:Light031.State = 1:Light032.State = 1:Light027.State = 1:Light080.State = 0
        Case 10:KillerHits(CurrentPlayer, 3) = 9
    End Select
    Select Case KillerHits(CurrentPlayer, 4) 'Pinhead
        Case 0:Light054.State = 0:Light010.State = 0:Light011.State = 0:Light012.State = 0:Light013.State = 0:Light014.State = 0:Light009.State = 0:Light002.State = 0
        Case 1:Light054.State = 0:Light010.State = 1:Light011.State = 1:Light012.State = 0:Light013.State = 0:Light014.State = 0:Light009.State = 0:Light002.State = 0
        Case 2:Light054.State = 0:Light010.State = 1:Light011.State = 1:Light012.State = 1:Light013.State = 1:Light014.State = 0:Light009.State = 0:Light002.State = 0
        Case 3:Light054.State = 2:Light010.State = 0:Light011.State = 0:Light012.State = 0:Light013.State = 0:Light014.State = 0:Light009.State = 0:Light002.State = 0
        Case 4:Light054.State = 2:Light010.State = 1:Light011.State = 0:Light012.State = 0:Light013.State = 0:Light014.State = 0:Light009.State = 0:Light002.State = 2
        Case 5:Light054.State = 2:Light010.State = 1:Light011.State = 1:Light012.State = 0:Light013.State = 0:Light014.State = 0:Light009.State = 0:Light002.State = 2
        Case 6:Light054.State = 2:Light010.State = 1:Light011.State = 1:Light012.State = 1:Light013.State = 0:Light014.State = 0:Light009.State = 0:Light002.State = 2
        Case 7:Light054.State = 2:Light010.State = 1:Light011.State = 1:Light012.State = 1:Light013.State = 1:Light014.State = 0:Light009.State = 0:Light002.State = 2
        Case 8:Light054.State = 2:Light010.State = 1:Light011.State = 1:Light012.State = 1:Light013.State = 1:Light014.State = 1:Light009.State = 0:Light002.State = 2
        Case 9:Light054.State = 1:Light010.State = 1:Light011.State = 1:Light012.State = 1:Light013.State = 1:Light014.State = 1:Light009.State = 1:Light002.State = 0
        Case 10:KillerHits(CurrentPlayer, 4) = 9
    End Select
    Select Case KillerHits(CurrentPlayer, 5) 'Michael
        Case 0:Light055.State = 0:Light016.State = 0:Light017.State = 0:Light018.State = 0:Light019.State = 0:Light020.State = 0:Light015.State = 0:Light078.State = 0
        Case 1:Light055.State = 0:Light016.State = 1:Light017.State = 1:Light018.State = 0:Light019.State = 0:Light020.State = 0:Light015.State = 0:Light078.State = 0
        Case 2:Light055.State = 0:Light016.State = 1:Light017.State = 1:Light018.State = 1:Light019.State = 1:Light020.State = 0:Light015.State = 0:Light078.State = 0
        Case 3:Light055.State = 2:Light016.State = 0:Light017.State = 0:Light018.State = 0:Light019.State = 0:Light020.State = 0:Light015.State = 0:Light078.State = 0
        Case 4:Light055.State = 2:Light016.State = 1:Light017.State = 0:Light018.State = 0:Light019.State = 0:Light020.State = 0:Light015.State = 0:Light078.State = 2
        Case 5:Light055.State = 2:Light016.State = 1:Light017.State = 1:Light018.State = 0:Light019.State = 0:Light020.State = 0:Light015.State = 0:Light078.State = 2
        Case 6:Light055.State = 2:Light016.State = 1:Light017.State = 1:Light018.State = 1:Light019.State = 0:Light020.State = 0:Light015.State = 0:Light078.State = 2
        Case 7:Light055.State = 2:Light016.State = 1:Light017.State = 1:Light018.State = 1:Light019.State = 1:Light020.State = 0:Light015.State = 0:Light078.State = 2
        Case 8:Light055.State = 2:Light016.State = 1:Light017.State = 1:Light018.State = 1:Light019.State = 1:Light020.State = 1:Light015.State = 0:Light078.State = 2
        Case 9:Light055.State = 1:Light016.State = 1:Light017.State = 1:Light018.State = 1:Light019.State = 1:Light020.State = 1:Light015.State = 1:Light078.State = 0
        Case 10:KillerHits(CurrentPlayer, 5) = 9
    End Select
End Sub

' Other animations
Sub coffinf_Animate:coffin.Z = coffinf.CurrentAngle:coffindoor.Z = coffinf.CurrentAngle + 10:End Sub
Sub coffindoorf_Animate:coffindoor.RotZ = coffindoorf.CurrentAngle:End Sub
Sub dgatef_Animate:dgate1.RotZ = - dgatef.CurrentAngle:dgate2.RotZ = dgatef.CurrentAngle:End Sub
Sub dtf1_Animate:dt1.Z = dtf1.CurrentAngle:End Sub
Sub dtf2_Animate:dt2.Z = dtf2.CurrentAngle:End Sub
Sub dtf3_Animate:dt3.Z = dtf3.CurrentAngle:End Sub

' Apron digits display

Sub ApronDMDUpdate
    'apron Digits
    dim digit, tmp
    'hostages left
    tmp = HostagesLeft(CurrentPlayer)
    CStr(abs(tmp) )
    If len(tmp) = 1 then tmp = "0" &tmp
    For digit = 41 to 42
        ApronDMDDisplayChar mid(tmp, digit -40, 1), digit
    Next
    'balloons left
    tmp = Balloonsleft(CurrentPlayer)
    CStr(abs(tmp) )
    If len(tmp) = 1 then tmp = "0" &tmp
    For digit = 43 to 44
        ApronDMDDisplayChar mid(tmp, digit -42, 1), digit
    Next
    'extra balls
    tmp = ExtraBallsAwards(CurrentPlayer)
    CStr(abs(tmp) )
    ApronDMDDisplayChar mid(tmp, 1, 1), 45
End Sub

Sub ApronDMDDisplayChar(achar, adigit)
    achar = ASC(achar)
    Digits(adigit).ImageA = Chars(achar)
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
    startB2S(12)
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 530
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
    startB2S(13)
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

'***********************
'        Bumpers
'***********************

Sub Bumper001_Hit
    If Tilted Then Exit Sub
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    PlaySoundAt SoundFXDOF("fx_bumper", 108, DOFPulse, DOFContactors), Bumper001
    ' check for modes
    Lbumper1a.Duration 1, 100, 0
    Lbumper1b.Duration 1, 100, 0
    AddScore 1000
    Switches = Switches + 1
    ChuckyValue(CurrentPlayer) = INT(ChuckyValue(CurrentPlayer) + 500)
    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper001"
End Sub

Sub Bumper002_Hit
    If Tilted Then Exit Sub
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    PlaySoundAt SoundFXDOF("fx_bumper", 109, DOFPulse, DOFContactors), Bumper002
    ' check for modes
    Lbumper2a.Duration 1, 100, 0
    Lbumper2b.Duration 1, 100, 0
    AddScore 1000
    Switches = Switches + 1
    ChuckyValue(CurrentPlayer) = INT(ChuckyValue(CurrentPlayer) + 500)
    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper002"
End Sub

'*********
' Lanes
'*********
' in and outlanes
Sub Trigger001_Hit
    PLaySoundAt "fx_sensor", Trigger001
    If Tilted Then Exit Sub
    Switches = Switches + 1
    Select Case Mode
        Case 0 'normal scoring
            Knives(1) = 1:CheckKnives
            Light044.State = Knives(1)
            Addscore 5000
        Case 1 'Restore Power
            RPlights(1) = RPlights(1) + 1
            CheckRP
            If RPlights(1) <4 Then Addscore 1000 * RPlights(1)
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger001"
End Sub

Sub Trigger002_Hit
    PLaySoundAt "fx_sensor", Trigger002
    If Tilted Then Exit Sub
    Switches = Switches + 1
    Select Case Mode
        Case 0 'normal scoring
            Knives(2) = 1:CheckKnives
            Light045.State = Knives(2)
            Addscore 1000
        Case 1 'Restore Power
            RPlights(1) = RPlights(1) + 1
            CheckRP
            If RPlights(1) <4 Then Addscore 1000 * RPlights(1)
    End Select
    ' remember last trigger hit by the ball
    If LastSwitchHit <> "Trigger011" Then
        LastSwitchHit = "Trigger002"
    End If
End Sub

Sub Trigger003_Hit
    PLaySoundAt "fx_sensor", Trigger003
    If Tilted Then Exit Sub
    Switches = Switches + 1
    Select Case Mode
        Case 0 'normal scoring
            Knives(3) = 1:CheckKnives
            Light046.State = Knives(3)
            Addscore 1000
        Case 1 'Restore Power
            RPlights(1) = RPlights(1) + 1
            CheckRP
            If RPlights(1) <4 Then Addscore 1000 * RPlights(1)
    End Select
    ' remember last trigger hit by the ball
    If LastSwitchHit <> "Trigger010" Then
        LastSwitchHit = "Trigger003"
    End If
End Sub

Sub Trigger004_Hit
    PLaySoundAt "fx_sensor", Trigger004
    If Tilted Then Exit Sub
    Switches = Switches + 1
    Select Case Mode
        Case 0 'normal scoring
            Knives(4) = 1:CheckKnives
            Light047.State = Knives(4)
            Addscore 5000
        Case 1 'Restore Power
            RPlights(1) = RPlights(1) + 1
            CheckRP
            If RPlights(1) <4 Then Addscore 1000 * RPlights(1)
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger004"
End Sub

Sub CheckKnives
    Dim tmp
    tmp = Knives(1) + Knives(2) + Knives(3) + Knives(4)
    If tmp = 4 Then
        AddBonusMultiplier 1
        Knives(1) = 0:Light044.State = 0
        Knives(2) = 0:Light045.State = 0
        Knives(3) = 0:Light046.State = 0
        Knives(4) = 0:Light047.State = 0
        LightEffect 2
    End If
End Sub

Sub RotateKnivesLeft
    Dim tmp
    tmp = Knives(1)
    Knives(1) = Knives(2)
    Knives(2) = Knives(3)
    Knives(3) = Knives(4)
    Knives(4) = tmp
    Light044.State = Knives(1)
    Light045.State = Knives(2)
    Light046.State = Knives(3)
    Light047.State = Knives(4)
End Sub

Sub RotateKnivesRight
    Dim tmp
    tmp = Knives(4)
    Knives(4) = Knives(3)
    Knives(3) = Knives(2)
    Knives(2) = Knives(1)
    Knives(1) = tmp
    Light044.State = Knives(1)
    Light045.State = Knives(2)
    Light046.State = Knives(3)
    Light047.State = Knives(4)
End Sub

'top lanes
Sub Trigger006_Hit 'top left
    PLaySoundAt "fx_sensor", Trigger006
    If Tilted Then Exit Sub
    Switches = Switches + 1
    Select Case Mode
        Case 0 'normal scoring
            Nun(1) = 1:CheckNun
            Light050.State = Nun(1)
            Addscore 1000
        Case 1 'Restore Power
            RPlights(10) = RPlights(10) + 1
            CheckRP
            If RPlights(10) <4 Then Addscore 1000 * RPlights(10)
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger006"
End Sub

Sub Trigger007_Hit 'top center
    PLaySoundAt "fx_sensor", Trigger007
    If Tilted Then Exit Sub
    If bSkillshotReady AND SkillshotType = 1 Then 'award Skillshot
        AwardSkillshot 250000
        Exit Sub
    End If
    Switches = Switches + 1
    Select Case Mode
        Case 0 'normal scoring
            Nun(2) = 1:CheckNun
            Light051.State = Nun(2)
            Addscore 1000
        Case 1 'Restore Power
            RPlights(10) = RPlights(10) + 1
            CheckRP
            If RPlights(10) <4 Then Addscore 1000 * RPlights(10)
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger007"
End Sub

Sub Trigger008_Hit 'top right
    PLaySoundAt "fx_sensor", Trigger008
    If Tilted Then Exit Sub
    Switches = Switches + 1
    Select Case Mode
        Case 0 'normal scoring
            Nun(3) = 1:CheckNun
            Light052.State = Nun(3)
            Addscore 1000
        Case 1 'Restore Power
            RPlights(10) = RPlights(10) + 1
            CheckRP
            If RPlights(10) <4 Then Addscore 1000 * RPlights(10)
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger008"
End Sub

Sub CheckNuN
    Dim tmp
    tmp = Nun(1) + Nun(2) + Nun(3)
    If tmp = 3 Then
        AddPlayfieldMultiplier 1
        Nun(1) = 0:Light050.State = 0
        Nun(2) = 0:Light051.State = 0
        Nun(3) = 0:Light052.State = 0
        LightEffect 2
    End If
End Sub

Sub RotateNunLeft
    Dim tmp
    tmp = Nun(1)
    Nun(1) = Nun(2)
    Nun(2) = Nun(3)
    Nun(3) = tmp
    Light050.State = Nun(1)
    Light051.State = Nun(2)
    Light052.State = Nun(3)
End Sub

Sub RotateNunRight
    Dim tmp
    tmp = Nun(3)
    Nun(3) = Nun(2)
    Nun(2) = Nun(1)
    Nun(1) = tmp
    Light050.State = Nun(1)
    Light051.State = Nun(2)
    Light052.State = Nun(3)
End Sub

' 5 killer switches
Sub Trigger005_Hit 'Jason - left orbit
    PLaySoundAt "fx_sensor", Trigger005
    If Tilted OR bSkillShotReady Then Exit Sub
    Switches = Switches + 1
    'Hostages - can be rescued on all modes
    If HostagesLights(CurrentPlayer, 1) = 2 Then 'the light is blinking, so rescue the HostagesLeft
        HostagesLights(CurrentPlayer, 1) = 0
        HostagesRescued(CurrentPlayer) = HostagesRescued(CurrentPlayer) + 1
        HostagesLeft(CurrentPlayer) = HostagesLeft(CurrentPlayer) - 1
        AddScore 10000
        CheckHostages
    End If
    Select Case Mode
        Case 0 'normal scoring
            KillerHits(CurrentPlayer, 1) = KillerHits(CurrentPlayer, 1) + 1
            UpdateKillerLights
            CheckKillers 1
            'Chaos letter can only be collected during Mode 0: standard mode
            If ChaosLights(CurrentPlayer, 1) = 2 Then 'collect the letter and light the next letter
                PLaySound "sfx_bomb1"
                ChaosLights(CurrentPlayer, 1) = 1
                ChaosLights(CurrentPlayer, 2) = 2
                UpdateLights
            End If
        Case 1 'Restore Power
            RPlights(3) = RPlights(3) + 1
            CheckRP
            If RPlights(3) <4 Then Addscore 1000 * RPlights(3)
        Case 2 'Escape HW
            If JackpotLights(1) Then
                JackpotLights(1) = 0
                AwardJackpot
                SetupJackpots
            End If
        Case 4 'Dracula MB
            If JackpotLights(1) Then
                JackpotLights(1) = 0
                AwardJackpot
                SetupJackpots
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger005"
End Sub

Sub Trigger009_Hit 'Ghostface - left ramp done
    PLaySoundAt "fx_sensor", Trigger009
    If Tilted Then Exit Sub
    Switches = Switches + 1 'always counting
    'combo
    If LastSwitchHit = "Trigger011" OR LastSwitchHit = "Trigger010" Then
        AwardCombo
    Else
        ComboCount = 0
    End If
    'Hostages
    If HostagesLights(CurrentPlayer, 2) = 2 Then 'the light is blinking, so rescue the HostagesLeft
        HostagesLights(CurrentPlayer, 2) = 0
        HostagesRescued(CurrentPlayer) = HostagesRescued(CurrentPlayer) + 1
        HostagesLeft(CurrentPlayer) = HostagesLeft(CurrentPlayer) - 1
        AddScore 10000
        CheckHostages
    End If
    Select Case Mode
        Case 0 'normal scoring
            KillerHits(CurrentPlayer, 2) = KillerHits(CurrentPlayer, 2) + 1
            UpdateKillerLights
            CheckKillers 2
            'Chaos letter can only be collected during Mode 0: standard mode
            If ChaosLights(CurrentPlayer, 2) = 2 Then 'collect the letter and light the next letter
                PLaySound "sfx_bomb1"
                ChaosLights(CurrentPlayer, 2) = 1
                ChaosLights(CurrentPlayer, 3) = 2
                UpdateLights
            End If
        Case 1 'Restore Power
            RPlights(4) = RPlights(4) + 1
            CheckRP
            If RPlights(4) <4 Then Addscore 1000 * RPlights(4)
        Case 2 'Escape HW
            If JackpotLights(2) Then
                JackpotLights(2) = 0
                AwardJackpot
                SetupJackpots
            End If
        Case 3 'Pennywise MB
            If JackpotLights(2) Then
                AwardJackpot
            End If
        Case 4 'Dracula MB
            If JackpotLights(2) Then
                JackpotLights(2) = 0
                AwardJackpot
                SetupJackpots
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger009"
End Sub

Sub Trigger013_Hit 'Freddy - inner loop
    PLaySoundAt "fx_sensor", Trigger013
    If Tilted Then Exit Sub
    Switches = Switches + 1
    If LastSwitchHit = "Trigger013" Then
        AwardCombo
    Else
        ComboCount = 0
    End If
    'Hostages
    If HostagesLights(CurrentPlayer, 3) = 2 Then 'the light is blinking, so rescue the HostagesLeft
        HostagesLights(CurrentPlayer, 3) = 0
        HostagesRescued(CurrentPlayer) = HostagesRescued(CurrentPlayer) + 1
        HostagesLeft(CurrentPlayer) = HostagesLeft(CurrentPlayer) - 1
        AddScore 10000
        CheckHostages
    End If
    Select Case Mode
        Case 0 'normal scoring
            KillerHits(CurrentPlayer, 3) = KillerHits(CurrentPlayer, 3) + 1
            UpdateKillerLights
            CheckKillers 3
            'Chaos letter can only be collected during Mode 0: standard mode
            If ChaosLights(CurrentPlayer, 3) = 2 Then 'collect the letter and light the next letter
                PLaySound "sfx_bomb1"
                ChaosLights(CurrentPlayer, 3) = 1
                ChaosLights(CurrentPlayer, 4) = 2
                UpdateLights
            End If
        Case 1 'Restore Power
            RPlights(5) = RPlights(5) + 1
            CheckRP
            If RPlights(5) <4 Then Addscore 1000 * RPlights(5)
        Case 2 'Escape HW
            If JackpotLights(3) Then
                JackpotLights(3) = 0
                AwardJackpot
                SetupJackpots
            End If
        Case 4 'Dracula MB
            If JackpotLights(3) Then
                JackpotLights(3) = 0
                AwardJackpot
                SetupJackpots
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger013"
End Sub

Sub Trigger011_Hit 'Pinhead - center ramp
    PLaySoundAt "fx_sensor", Trigger011
    If Tilted Then Exit Sub
    If bSkillshotReady AND SkillshotType = 2 Then 'award SuperSkillshot
        AwardSuperSkillshot 500000
        Exit Sub
    End If
    Switches = Switches + 1
    If LastSwitchHit = "Trigger011" OR LastSwitchHit = "Trigger010" Then
        AwardCombo
    Else
        ComboCount = 0
    End If
    'Hostages
    If HostagesLights(CurrentPlayer, 4) = 2 Then 'the light is blinking, so rescue the HostagesLeft
        HostagesLights(CurrentPlayer, 4) = 0
        HostagesRescued(CurrentPlayer) = HostagesRescued(CurrentPlayer) + 1
        HostagesLeft(CurrentPlayer) = HostagesLeft(CurrentPlayer) - 1
        AddScore 10000
        CheckHostages
    End If
    Select Case Mode
        Case 0 'normal scoring
            KillerHits(CurrentPlayer, 4) = KillerHits(CurrentPlayer, 4) + 1
            UpdateKillerLights
            CheckKillers 4
            'Chaos letter can only be collected during Mode 0: standard mode
            If ChaosLights(CurrentPlayer, 4) = 2 Then 'collect the letter and light the next letter
                PLaySound "sfx_bomb1"
                ChaosLights(CurrentPlayer, 4) = 1
                ChaosLights(CurrentPlayer, 5) = 2
                UpdateLights
            End If
        Case 1 'Restore Power
            RPlights(6) = RPlights(6) + 1
            CheckRP
            If RPlights(6) <4 Then Addscore 1000 * RPlights(6)
        Case 2 'Escape HW
            If JackpotLights(4) Then
                JackpotLights(4) = 0
                AwardJackpot
                SetupJackpots
            End If
        Case 3 'Pennywise MB
            If JackpotLights(4) Then
                AwardJackpot
            End If
        Case 4 'Dracula MB
            If JackpotLights(4) Then
                JackpotLights(4) = 0
                AwardJackpot
                SetupJackpots
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger011"
End Sub

Sub Trigger012_Hit 'Michael - right orbit
    PLaySoundAt "fx_sensor", Trigger012
    If Tilted OR bSkillShotReady Then Exit Sub
    Switches = Switches + 1
    'Hostages
    If HostagesLights(CurrentPlayer, 5) = 2 Then 'the light is blinking, so rescue the HostagesLeft
        HostagesLights(CurrentPlayer, 5) = 0
        HostagesRescued(CurrentPlayer) = HostagesRescued(CurrentPlayer) + 1
        HostagesLeft(CurrentPlayer) = HostagesLeft(CurrentPlayer) - 1
        AddScore 10000
        CheckHostages
    End If
    Select Case Mode
        Case 0 'normal scoring
            KillerHits(CurrentPlayer, 5) = KillerHits(CurrentPlayer, 5) + 1
            UpdateKillerLights
            CheckKillers 5
            'Chaos letter can only be collected during Mode 0: standard mode
            If ChaosLights(CurrentPlayer, 5) = 2 Then 'collect the letter and light the next letter
                PLaySound "vo_weaponsupgraded"
                Weapons(CurrentPlayer) = Weapons(CurrentPlayer) + 1
                LightSeqChaos.Play SeqRandom, 50, , 1000
                ChaosLights(CurrentPlayer, 1) = 2
                ChaosLights(CurrentPlayer, 2) = 0
                ChaosLights(CurrentPlayer, 3) = 0
                ChaosLights(CurrentPlayer, 4) = 0
                ChaosLights(CurrentPlayer, 5) = 0
                UpdateLights
            End If
        Case 1 'Restore Power
            RPlights(7) = RPlights(7) + 1
            CheckRP
            If RPlights(7) <4 Then Addscore 1000 * RPlights(7)
        Case 2 'Escape HW
            If JackpotLights(5) Then
                JackpotLights(5) = 0
                AwardJackpot
                SetupJackpots
            End If
        Case 4 'Dracula MB
            If JackpotLights(5) Then
                JackpotLights(5) = 0
                AwardJackpot
                SetupJackpots
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger012"
End Sub

'jump switch

Sub Trigger010_Hit 'jump ramp
    PLaySoundAt "fx_sensor", Trigger010
    If Tilted Then Exit Sub
    Addscore 4000
    Jumps(CurrentPlayer) = Jumps(CurrentPlayer) + 1 'only used in the bonus
    LightEffect 2
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger010"
End Sub

'***********
' Targets
'***********
Sub Target001_Hit 'left - jigsaw - 3 hits releases hostages
    PLaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Switches = Switches + 1
    Select Case Mode
        Case 0 'normal scoring
            Addscore 1000
            JigSawHits(CurrentPlayer) = JigSawHits(CurrentPlayer) + 1
            UpdateJigSaw
        Case 1 'Restore Power
            RPlights(2) = RPlights(2) + 1
            CheckRP
            If RPlights(2) <4 Then Addscore 1000 * RPlights(2)
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target001"
End Sub

Sub UpdateJigSaw '3 target Lights
    Select Case JigSawHits(CurrentPlayer)
        Case 0:Light065.State = 2:Light064.State = 2:Light063.State = 2
        Case 1:Light065.State = 1:Light064.State = 2:Light063.State = 2
        Case 2:Light065.State = 1:Light064.State = 1:Light063.State = 2
        Case 3:Light065.State = 2:Light064.State = 2:Light063.State = 2:LightEffect 2:JigSawHits(CurrentPlayer) = 0:ReleaseHostage
    End Select
End Sub

Sub ReleaseHostage 'lights one random hostage light
    Dim i, tmp
    tmp = 0
    For i = 1 to 5
        tmp = tmp + HostagesLights(CurrentPlayer, i) 'the lights are blinkning or off, so the value for each light is 2 or 0
    Next
    If tmp <10 Then                                  'there are some light/s off (state is 2 and there are 5 Lights)
        i = RndNbr(5)
        do while HostagesLights(CurrentPlayer, i) <> 0
            i = RndNbr(5)
        Loop
        HostagesLights(CurrentPlayer, i) = 2
        UpdateLights
    Else
        Addscore 10000 '10000 points if all the Hostages lights are lit
    End If
End Sub

Sub Target002_Hit 'top right - chucky
    PLaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    If LastSwitchHit = "Target007" Then
        LightEffect 7
        AwardCombo
    Else
        ComboCount = 0
    End If
    DMD " CHUCKY BONUS SCORE", CL(FormatScore(ChuckyValue(CurrentPlayer) ) ), "_", eNone, eBlink, eNone, 1500, True, ""
    Addscore ChuckyValue(CurrentPlayer)
    Switches = Switches + 1
    ChuckyValue(CurrentPlayer) = 1000 'reset the Chucky value
    ' check modes
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target002"
End Sub

Sub Target003_Hit ' dracula left
    PLaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 1000
    Switches = Switches + 1
    RotateHostagesLeft
    ' check modes
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target003"
End Sub

Sub Target004_Hit 'dracula right
    PLaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 1000
    Switches = Switches + 1
    RotateHostagesRight
    ' check modes
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target004"
End Sub

Sub RotateHostagesLeft 'rotate hostages lights to the left
    Dim tmp
    tmp = Light071.State
    Light071.State = Light072.State
    Light072.State = Light073.State
    Light073.State = Light074.State
    Light074.State = Light075.State
    Light075.State = tmp
End Sub

Sub RotateHostagesRight 'rotate hostages lights to the right
    Dim tmp
    tmp = Light075.State
    Light075.State = Light074.State
    Light074.State = Light073.State
    Light073.State = Light072.State
    Light072.State = Light071.State
    Light071.State = tmp
End Sub

'collect balloons at pennywise sewer
Sub Target005_Hit 'pennywise
    PLaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    If bSkillshotReady AND SkillshotType = 3 Then 'award SuperSkillshot 2
        AwardSuperSkillshot 750000
        Exit Sub
    End If

    Switches = Switches + 1
    ' check modes
    Select Case Mode
        Case 0 'normal scoring
            FlashForMs Light084, 500, 80, 0
            Addscore 1000
            GetaBalloon
            If XtraBalisLit(CurrentPlayer) Then
                AwardExtraBall
                XtraBalisLit(CurrentPlayer) = 0
                Light048.State = 0
            End If
        Case 1 'Restore Power
            RPlights(8) = RPlights(8) + 1
            CheckRP
            If RPlights(8) <4 Then Addscore 1000 * RPlights(8)
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target005"
End Sub

Sub rlband007_Hit 'pennywise rubber
    If Tilted Then Exit Sub
    Switches = Switches + 1
    ' check modes
    Select Case Mode
        Case 0 'normal scoring
            FlashForMs Light084, 500, 80, 0
            Addscore 1000
            GetaBalloon
    End Select
End Sub

Sub rsband005_Hit 'pennywise rubber
    If Tilted Then Exit Sub
    If bRestorePower Then Exit Sub
    Switches = Switches + 1
    ' check modes
    Select Case Mode
        Case 0 'normal scoring
            FlashForMs Light084, 500, 80, 0
            Addscore 1000
            GetaBalloon
    End Select
End Sub

Sub GetaBalloon
    BalloonsLeft(CurrentPlayer) = BalloonsLeft(CurrentPlayer) - 1
    If BalloonsLeft(CurrentPlayer) <0 Then BalloonsLeft(CurrentPlayer) = 0
    ApronDMDUpdate
    If BalloonsLeft(CurrentPlayer) = 0 Then 'Start Pennywise multiball
        StartPennywiseMB
    End If
End Sub

Sub Target006_Hit 'Annabelle - captive ball target
    PLaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Switches = Switches + 1
    ' check modes
    Select Case Mode
        Case 0 'normal scoring
            AnnabelleHits(CurrentPlayer) = AnnabelleHits(CurrentPlayer) + 1
            Addscore 1000 * AnnabelleHits(CurrentPlayer)
            Select Case AnnabelleHits(CurrentPlayer)
                Case 1
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL(" NNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL(" NNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL(" NNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                Case 2
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("  NABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("  NABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("  NABELLE"), "", eNone, eNone, eNone, 200, True, ""
                Case 3
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("   ABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("   ABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("   ABELLE"), "", eNone, eNone, eNone, 200, True, ""
                Case 4
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("    BELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("    BELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("    BELLE"), "", eNone, eNone, eNone, 200, True, ""
                Case 5
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("     ELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("     ELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("     ELLE"), "", eNone, eNone, eNone, 200, True, ""
                Case 6
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("      LLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("      LLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("      LLE"), "", eNone, eNone, eNone, 200, True, ""
                Case 7
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("       LE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("       LE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("       LE"), "", eNone, eNone, eNone, 200, True, ""
                Case 8
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("        E"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("        E"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("        E"), "", eNone, eNone, eNone, 200, True, ""
                Case 9
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("         "), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("         "), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("ANNABELLE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("         "), "", eNone, eNone, eNone, 200, True, ""
                    DMD " ANNABELLE JACKPOT", CL("IS READY"), "", eNone, eNone, eNone, 2000, True, "vo_annabellebonusisready"
                    Light079.BlinkInterval = 400
                    Light079.State = 2
                    Capkicker1.TimerEnabled = 1  '15 seconds
                    Capkicker1a.TimerEnabled = 1 '10 seconds
                Case 10
                    DMD CL("ANNABELLE JACKPOT"), CL("50.000"), "", eNone, eBlink, eNone, 2000, True, "vo_jackpot"
                    Addscore 50000
                    Light079.State = 0
                    LightEffect 3
                    AnnabelleHits(CurrentPlayer) = 0
                    Capkicker1.TimerEnabled = 0
                    Capkicker1a.TimerEnabled = 0
            End Select
        Case 1 'Restore Power
            RPlights(11) = RPlights(11) + 1
            CheckRP
            If RPlights(11) <4 Then Addscore 1000 * RPlights(11)
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target006"
End Sub

Sub Capkicker1_Timer 'turn off the light and stop the hurry-up
    Capkicker1.TimerEnabled = 0
    Light079.State = 0
    AnnabelleHits(CurrentPlayer) = 0
End Sub

Sub Capkicker1a_Timer 'speed up the light
    Capkicker1a.TimerEnabled = 0
    Light079.State = 0
    Light079.BlinkInterval = 150
    Light079.State = 2
End Sub

Sub Target007_Hit 'Leatherface - chainsaw
    PLaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Switches = Switches + 1
    ' check modes
    Select Case Mode
        Case 0 'normal scoring
            LeatherfaceHits(CurrentPlayer) = LeatherfaceHits(CurrentPlayer) + 1
            Addscore 1000 * LeatherfaceHits(CurrentPlayer)
            Select Case LeatherfaceHits(CurrentPlayer)
                Case 1
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL(" EATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL(" EATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL(" EATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                Case 2
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("  ATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("  ATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("  ATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                Case 3
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("   THERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("   THERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("   THERFACE"), "", eNone, eNone, eNone, 200, True, ""
                Case 4
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("    HERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("    HERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("    HERFACE"), "", eNone, eNone, eNone, 200, True, ""
                Case 5
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("     ERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("     ERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("     ERFACE"), "", eNone, eNone, eNone, 200, True, ""
                Case 6
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("      RFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("      RFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("      ERFACE"), "", eNone, eNone, eNone, 200, True, ""
                Case 7
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("       FACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("       FACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("       FACE"), "", eNone, eNone, eNone, 200, True, ""
                Case 8
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("        ACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("        ACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("        ACE"), "", eNone, eNone, eNone, 200, True, ""
                Case 9
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("         CE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("         CE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("         CE"), "", eNone, eNone, eNone, 200, True, ""
                Case 10
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("          E"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("          E"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("          E"), "", eNone, eNone, eNone, 200, True, ""
                Case 11
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("         "), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("         "), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("LEATHERFACE"), "", eNone, eNone, eNone, 200, True, ""
                    DMD "_", CL("         "), "", eNone, eNone, eNone, 200, True, ""
                    DMD "LEATHERFACE JACKPOT", CL("IS READY"), "", eNone, eNone, eNone, 2000, True, "vo_leatherfacejackpotisready"
                    LightEffect 7
                    Light076.BlinkInterval = 300
                    Light076.State = 2
                    Target007.TimerEnabled = 1 '20 seconds
                Case 12
                    DMDFlush
                    DMD "LEATHERFACE JACKPOT", CL("75.000"), "", eNone, eBlink, eNone, 2000, True, "vo_jackpot"
                    Addscore 75000
                    LightEffect 7
                    LightEffect 3
                    LeatherfaceHits(CurrentPlayer) = 0
                    Target007.TimerEnabled = 0
                    Light076.State = 0
                    LeatherfaceHits(CurrentPlayer) = 0
            End Select
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target007"
End Sub

Sub Target007_Timer 'turn off the light and stop the hurry-up
    Target007.TimerEnabled = 0
    Light076.State = 0
    LeatherfaceHits(CurrentPlayer) = 0
End Sub

'rubbers

Sub rlband004_Hit:AddScore 110:Switches = Switches + 1:End Sub
Sub rlband005_Hit:AddScore 110:Switches = Switches + 1:End Sub

'***********
'  Spinner
'***********

Sub Spinner001_Spin 'left
    If Tilted Then Exit Sub
    PlaySoundAt "fx_spinner", spinner001
    DOF 116, DOFPulse
    Addscore 1000
    Light049.Duration 1, 100, 0
    'check modes
    Select case Mode
        Case 0 'standad Mode
            SpinnerHits(CurrentPlayer) = SpinnerHits(CurrentPlayer) + 1
            CheckSpinnerHits
        Case 2 'Escape HW
            Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + 1000
    End Select
End Sub

Sub Spinner002_Spin 'right
    If Tilted Then Exit Sub
    PlaySoundAt "fx_spinner", spinner002
    DOF 115, DOFPulse
    Addscore 1000
    Light001.Duration 1, 100, 0
    'check modes
    Select case Mode
        Case 0 'standad Mode
            SpinnerHits(CurrentPlayer) = SpinnerHits(CurrentPlayer) + 1
            CheckSpinnerHits
        Case 2 'Escape HW
            Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + 1000
    End Select
End Sub

Sub CheckSpinnerHits
    If SpinnerHits(CurrentPlayer) = 250 Then
        SpinnerHits(CurrentPlayer) = 0
        Weapons(CurrentPlayer) = Weapons(CurrentPlayer) + 1
        PlaySound "vo_weaponsupgraded"
    End If
End Sub

'*******************
' Draculas hole
'*******************

Sub Kicker001_Hit
    Dim delay
    delay = 1500
    PlaySoundAt "fx_hole_enter", Kicker001
    'Kicker001.Destroyball 'do not delete the ball, just close the door
    coffindoorf.RotateToStart
    BallsinHole = BallsInHole + 1
    If NOT Tilted Then
        If bRestorePowerReady Then StartRestorePower2:delay = 8000
        If bEscapeHWReady Then StartEscapeHW2:delay = 8000
        Select Case Mode
            Case 0 'normal scoring
                BallsInLock(CurrentPlayer) = BallsInLock(CurrentPlayer) + 1
                Select Case BallsInLock(CurrentPlayer)
                    Case 1
                        DMD "", CL("BALL 1 LOCKED"), "_", eNone, eBlink, eNone, 1500, True, "vo_ball1locked":Addscore 25000:Delay = 2500
                    Case 2
                        DMD "", CL("BALL 2 LOCKED"), "_", eNone, eBlink, eNone, 1500, True, "vo_ball2locked":Addscore 25000:Delay = 2500
                    Case 3
                        DMD "", CL("BALL 3 LOCKED"), "_", eNone, eBlink, eNone, 1500, True, "vo_ball3locked":Addscore 25000
                        vpmtimer.addtimer 3000, "StartDraculaMB '"
                        delay = 5500
                End Select
            Case 1 'Restore Power
                RPlights(9) = RPlights(9) + 1
                CheckRP
                If RPlights(9) <4 Then Addscore 1000 * RPlights(9)
            Case 4 'Dracula MB 'check for super jackpot
                If Light077.State = 2 Then
                    AwardSuperJackpot
                    Light077.State = 0
                    JackpotCount = 0
                Else
                    Addscore 1000
                End If
        End Select
        LightEffect 7
        vpmtimer.addtimer delay, "kickBallOut '"
    Else 'if Tilted kick the ball fast
        vpmtimer.addtimer 500, "kickBallOut '"
    End If
End Sub

Sub kickBallOut
    If BallsinHole> 0 Then
        BallsinHole = BallsInHole - 1
        PlaySoundAt SoundFXDOF("fx_popper", 111, DOFPulse, DOFcontactors), Kicker001
        coffindoorf.RotateToEnd
        'Kicker001.CreateSizedBallWithMass BallSize / 2, BallMass
        vpmTimer.AddTimer 200, "Kicker001.kick 170, 22 '"
        LightEffect 5
    End If
End Sub

' Drop targets

Sub Tomb1_Hit 'left dt
    If Tilted Then Exit Sub
    Switches = Switches + 1
    ' check modes
    LightEffect 7
    Addscore 1000
    Light083.Duration 1, 200, 0
    DT(CurrentPlayer, 1) = 1
    UpdateDT
    ' remember last trigger hit by the ball
    LastSwitchHit = "Tomb1"
End Sub

Sub Tomb2_Hit 'center dt
    If Tilted Then Exit Sub
    Switches = Switches + 1
    ' check modes
    LightEffect 7
    Addscore 1000
    Light083.Duration 1, 200, 0
    DT(CurrentPlayer, 2) = 1
    UpdateDT
    ' remember last trigger hit by the ball
    LastSwitchHit = "Tomb2"
End Sub

Sub Tomb3_Hit 'right dt
    If Tilted Then Exit Sub
    LightEffect 7
    Addscore 1000
    Light083.Duration 1, 200, 0
    Switches = Switches + 1
    DT(CurrentPlayer, 3) = 1
    UpdateDT
    ' check modes
    ' remember last trigger hit by the ball
    LastSwitchHit = "Tomb3"
End Sub

Sub ResetDT ' all
    DT(CurrentPlayer, 1) = 0
    DT(CurrentPlayer, 2) = 0
    DT(CurrentPlayer, 3) = 0
    UpdateDT
End Sub
'    PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFcontactors)
Sub UpdateDT
    If DT(CurrentPlayer, 1) = 1 Then
        PLaySoundAt "fx_Droptarget", dt1
        startB2S(10)
        tomb1.IsDropped = 1
        tomb1b.IsDropped = 1
        dtf1.RotateToEnd
    Else
        PLaySoundAt SoundFXDOF("fx_ResetDrop", 113, DOFPulse, DOFcontactors), dt1
        tomb1.IsDropped = 0
        tomb1b.IsDropped = 0
        dtf1.RotateToStart
    End If

    If DT(CurrentPlayer, 2) = 1 Then
        PLaySoundAt "fx_Droptarget", dt2
        startB2S(11)
        tomb2.IsDropped = 1
        tomb2b.IsDropped = 1
        dtf2.RotateToEnd
        dgatef.RotateToEnd
        PlaySound "sfx_tomb"
        vpmTimer.AddTimer 200, "coffinf.RotateToEnd '"
        vpmTimer.AddTimer 2000, "coffindoorf.RotateToEnd '"
        draculagate.IsDropped = 1
        Kicker001.Enabled = 1
    Else
        PLaySoundAt SoundFXDOF("fx_ResetDrop", 114, DOFPulse, DOFcontactors), dt2
        tomb2.IsDropped = 0
        tomb2b.IsDropped = 0
        dtf2.RotateToStart
        dgatef.RotateToStart
        PlaySound "sfx_tomb"
        vpmTimer.AddTimer 500, "coffinf.RotateToStart '"
        vpmTimer.AddTimer 200, "coffindoorf.RotateToStart '"
        draculagate.IsDropped = 0
        Kicker001. Enabled = 0
    End If

    If DT(CurrentPlayer, 3) = 1 Then
        PLaySoundAt "fx_Droptarget", dt3
        startB2S(10)
        tomb3.IsDropped = 1
        tomb3b.IsDropped = 1
        dtf3.RotateToEnd
    Else
        PLaySoundAt SoundFXDOF("fx_ResetDrop", 112, DOFPulse, DOFcontactors), dt3
        tomb3.IsDropped = 0
        tomb3b.IsDropped = 0
        dtf3.RotateToStart
    End If
End Sub

'********************************
'   Wizard Modes & Multiballs
'********************************

Dim Counter

Sub StartCountDown(n)
    Counter = n
    CountDown.Enabled = 1
End Sub

Sub CountDown_Timer
    Counter = Counter -1
    If Counter = 0 Then Me.Enabled = 0
End Sub

'********************************
'        Restore Power:
'         Wizard mode
'  after 5 killers are captured
'********************************
' this is not a multiball,
' but a 2 minutes scoring feast
' Mode 1

Sub CheckKillers(n)                          'n is the number of the current killer hit, and it is used to score
    Select Case KillerHits(CurrentPlayer, n) 'number of hits 1 to 9
        Case 1, 2, 3:Addscore 10000
        Case 4, 5, 6:Addscore 25000:PlaySfx
        Case 7, 8:Addscore 50000:PlaySfx
        Case 9:Addscore 200000:PlaySfx ' last Hit. extra hits do not score anymore
            KillersCompleted(CurrentPlayer) = KillersCompleted(CurrentPlayer) + 1
            If KillersCompleted(CurrentPlayer) MOD 2 = 0 Then
                ExtraBallIsLit
            End If
    End Select
    dim i, tmp
    For i = 1 to 5
        tmp = tmp + KillerHits(CurrentPlayer, i)
    Next
    If tmp >= 45 AND Mode = 0 Then '9 hits for each killer = 45, then all the killers are captured so start the wizard Restore Power
        StartRestorePower
    End If
End Sub

Sub StartRestorePower
    Dim i
    For Each i in aTiltLights:i.State = 0:Next 'turn all lights off
    GiOff
    Light083.BlinkInterval = 150
    Light083.State = 2
    Light077.BlinkInterval = 150
    Light077.State = 2
    Light062.State = 2
    bRestorePowerReady = True
    Mode = -1 'stop all modes even the normal scoring
    'Drop the droptargets
    DT(CurrentPlayer, 1) = 1
    DT(CurrentPlayer, 2) = 1
    DT(CurrentPlayer, 3) = 1
    UpdateDT
    LightSeqFlashers.UpdateInterval = 100
    LightSeqFlashers.Play SeqBlinking, , 150, 250
    DMDFlush
    DMD CL("RESTORE POWER"), CL("IS READY"), "_", eNone, eNone, eNone, 2500, True, "vo_restorepower"
    DMD CL("SHOOT THE SCOOP"), CL("TO START"), "_", eNone, eNone, eNone, 2500, True, ""
End Sub

Sub StartRestorePower2
    dim i
    PlaySong "m_RestorePower"
    DMD CL("STARTING"), CL("RESTORE POWER"), "_", eNone, eNone, eNone, 2500, True, ""
    DMD CL("COMPLETE"), CL("THE 11 POWERLINES"), "_", eNone, eNone, eNone, 2500, True, ""
    DMD CL("YOU HAVE"), CL("2 MINUTES"), "_", eNone, eNone, eNone, 2500, True, ""
    LightSeqFlashers.StopPlay
    bRestorePowerReady = False
    bRestorePower = True
    Mode = 1
    'setup lights - all lights off
    For Each i in aTiltLights:i.State = 0:Next
    'init the hit array
    For i = 0 to 11:RPlights(i) = 0:Next
    UpdateRPLights
    GiOff
    GiRedOn
    ModeScore = 0
    'Start the timers
    EnableBallSaver 120
    StartCountDown 120
    RestorePowerTimer.Enabled = 1  '120 seconds
    RestorePowerTimer2.Enabled = 1 '15 seconds to reduce the power on all powerlines that are not completed.
End Sub

Sub CheckRP
    Dim i, tmp
    tmp = 0
    UpdateRPLights
    For i = 1 to 11
        tmp = tmp + RPlights(i)
    Next
    If tmp >= 55 Then '11 power lines, value 5 is completed, all the completed so...
        WinRestorePower
    End If
End Sub

Sub RestorePowerTimer_Timer
    Me.Enabled = 0
    StopRestorePower
End Sub

Sub RestorePowerTimer2_Timer
    Dim i
    For i = 1 to 11 'check the power lines, and if they are not online (value 5) then reduce them
        If RPlights(i)> 0 AND RPlights(i) <5 Then
            RPlights(i) = RPlights(i) - 1
        End If
    Next
    UpdateRPLights
End Sub

Sub StopRestorePower
    Dim i
    For i = 1 to 5 'reset the killers hits count
        KillerHits(CurrentPlayer, i) = 0
    Next
    'stop the timers
    BallSaverTimerExpired_Timer 'stop the ball saver
    RestorePowerTimer.Enabled = 0
    RestorePowerTimer2.Enabled = 0
    CountDown.Enabled = 0
    For each i in aTiltLights:i.BlinkInterval = 400:Next
    For each i in aHostagesLights:i.BlinkInterval = 350:Next
    DisableTable True
    TiltRecoveryTimer.Enabled = True 'this will check for all the balls being drained and it will continue the game.
    GiOn
    GiRedOff
    Mode = 0
    DMDFlush
    DMD "RESTORE POWER SCORE", CL(FormatScore(ModeScore) ), "_", eNone, eNone, eNone, 3000, True, ""
    DMD CL("PLEASE WAIT"), CL("COLLECTING BALLS"), "_", eNone, eNone, eNone, 2500, True, ""
    ResetDT
    bRestorePower = False
End Sub

Sub WinRestorePower
    AwardExtraBall
    LightEffect 2
    GiEffect 2
    'DMD
    DMDFlush
    DMD CL("CONGRATULATIONS"), "", "_", eBlink, eNone, eNone, 2500, True, "vo_welldone" &RndNbr(4)
    DMD CL("YOU RESTORED"), " THE VILLAGES POWER", "_", eNone, eNone, eNone, 2500, True, ""
    StopRestorePower
End Sub

Sub UpdateRPLights 'update the lights blinking according to the number of hits
    Dim i
    Select Case RPlights(1)
        Case 0:For each i in aRPL1:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next:Addscore 1000
        Case 1:For each i in aRPL1:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next:Addscore 2000
        Case 2:For each i in aRPL1:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next:Addscore 3000
        Case 3:For each i in aRPL1:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next:Addscore 4000
        Case 4:For each i in aRPL1:i.State = 1:Next:RPlights(1) = 5:i = 5000 * Counter:DMD " POWERLINE 1 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(1) = 5 ' power line is on
    End Select
    Select Case RPlights(2)
        Case 0:For each i in aRPL2:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL2:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL2:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL2:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL2:i.State = 1:Next:RPlights(2) = 5:i = 5000 * Counter:DMD " POWERLINE 2 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(2) = 5 ' power line is on
    End Select
    Select Case RPlights(3)
        Case 0:For each i in aRPL3:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL3:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL3:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL3:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL3:i.State = 1:Next:RPlights(3) = 5:i = 5000 * Counter:DMD " POWERLINE 4 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(3) = 5 ' power line is on
    End Select
    Select Case RPlights(4)
        Case 0:For each i in aRPL4:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL4:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL4:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL4:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL4:i.State = 1:Next:RPlights(4) = 5:i = 5000 * Counter:DMD " POWERLINE 4 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(4) = 5 ' power line is on
    End Select
    Select Case RPlights(5)
        Case 0:For each i in aRPL5:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL5:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL5:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL5:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL5:i.State = 1:Next:RPlights(5) = 5:i = 5000 * Counter:DMD " POWERLINE 5 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(5) = 5 ' power line is on
    End Select
    Select Case RPlights(6)
        Case 0:For each i in aRPL6:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL6:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL6:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL6:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL6:i.State = 1:Next:RPlights(6) = 5:i = 5000 * Counter:DMD " POWERLINE 6 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(6) = 5 ' power line is on
    End Select
    Select Case RPlights(7)
        Case 0:For each i in aRPL7:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL7:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL7:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL7:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL7:i.State = 1:Next:RPlights(7) = 5:i = 5000 * Counter:DMD " POWERLINE 7 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(7) = 5 ' power line is on
    End Select
    Select Case RPlights(8)
        Case 0:For each i in aRPL8:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL8:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL8:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL8:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL8:i.State = 1:Next:RPlights(8) = 5:i = 5000 * Counter:DMD " POWERLINE 8 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(8) = 5 ' power line is on
    End Select
    Select Case RPlights(9)
        Case 0:For each i in aRPL9:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL9:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL9:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL9:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL9:i.State = 1:Next:RPlights(9) = 5:i = 5000 * Counter:DMD " POWERLINE 9 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(9) = 5 ' power line is on
    End Select
    Select Case RPlights(10)
        Case 0:For each i in aRPL10:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL10:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL10:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL10:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL10:i.State = 1:Next:RPlights(10) = 5:i = 5000 * Counter:DMD " POWERLINE 10 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(10) = 5 ' power line is on
    End Select
    Select Case RPlights(11)
        Case 0:For each i in aRPL11:i.BlinkInterval = 1500:i.State = 0:i.State = 2:Next
        Case 1:For each i in aRPL11:i.BlinkInterval = 1000:i.State = 0:i.State = 2:Next
        Case 2:For each i in aRPL11:i.BlinkInterval = 500:i.State = 0:i.State = 2:Next
        Case 3:For each i in aRPL11:i.BlinkInterval = 250:i.State = 0:i.State = 2:Next
        Case 4:For each i in aRPL11:i.State = 1:Next:RPlights(11) = 5:i = 5000 * Counter:DMD " POWERLINE 11 IS ON", CL(FormatScore(i) ), "_", eNone, eNone, eNone, 2000, True, "sfx_Electricity":Addscore i
        Case Else:RPlights(11) = 5 ' power line is on
    End Select
End Sub

'********************************
'       Escape Horrorburg
' after 25 hostages are rescued
'********************************
' Mode 2
' 3 ball multiball, ballsaver 20 seconds

Sub CheckHostages
    PlaySound "vo_rescuedhostage"
    UpdateHLights
    ApronDMDUpdate
    If HostagesLeft(CurrentPlayer) <= 0 AND Mode = 0 Then
        vpmTimer.AddTimer 2000, "StartEscapeHW '"
    End If
End Sub

Sub StartEscapeHW
    Dim i
    For Each i in aTiltLights:i.State = 0:Next 'turn all lights off
    GiOff
    Light083.BlinkInterval = 150
    Light083.State = 2
    Light077.BlinkInterval = 150
    Light077.State = 2
    Light085.State = 2
    bEscapeHWReady = True
    Mode = -1 'stop all modes even the normal scoring
    'Drop the droptargets
    DT(CurrentPlayer, 1) = 1
    DT(CurrentPlayer, 2) = 1
    DT(CurrentPlayer, 3) = 1
    UpdateDT
    LightSeqFlashers.UpdateInterval = 100
    LightSeqFlashers.Play SeqBlinking, , 150, 250
    DMDFlush
    DMD CL("ESCAPE HORRORBURG"), CL("IS READY"), "_", eNone, eNone, eNone, 2500, True, "vo_escapehorrorburg"
    DMD CL("SHOOT THE SCOOP"), CL("TO START"), "_", eNone, eNone, eNone, 2500, True, ""
End Sub

Sub StartEscapeHW2
    dim i
    PlaySong "m_mb1"
    DMD CL("STARTING"), CL("ESCAPE HORRORBURG"), "_", eNone, eNone, eNone, 2500, True, ""
    DMD CL("SHOOT THE JACKPOTS"), CL("AND THE SPINNERS"), "_", eNone, eNone, eNone, 2500, True, ""
    LightSeqFlashers.StopPlay
    bEscapeHWReady = False
    bEscapeHW = True
    Mode = 2
    'setup lights - all lights off - all Gi blinks
    For Each i in aTiltLights:i.State = 0:Next
    For Each i In aGiLights:i.State = 2:Next
    For Each i In aGiLightsRED:i.State = 2:Next
    Jackpot(CurrentPlayer) = 50000 'reset to 50000 - spinners increase value
    'setup jackpot lights according to the nr of weapons collected
    SetupJackpots
    vpmTimer.AddTimer 600, "AddMultiball 2 '"
    EnableBallSaver 20
    ModeScore = 0
End Sub

Sub StopEscapeHW 'when the multiball is over
    bEscapeHW = False
    Mode = 0
    GiOn
    GiRedOff
    UpdateLights 'chaos and hostages lights
    UpdateKillerLights
    ChangeSong
    DT(CurrentPlayer, 1) = 0
    DT(CurrentPlayer, 2) = 0
    DT(CurrentPlayer, 3) = 0
    UpdateDT
    HostagesLeft(CurrentPlayer) = 25 'reset the hostages to start rescuing again
    ApronDMDUpdate
    Light077.State = 0               'be sure the super jackpot is off
    DMD "ESCAPE HW SCORE", CL(FormatScore(ModeScore) ), "_", eNone, eNone, eNone, 3000, True, ""
End Sub

Sub SetupJackpots
    Dim i, j, tmp
    tmp = Weapons(CurrentPlayer)
    If tmp> 5 then tmp = 5
    If tmp = 0 Then tmp = 1
    'reset the Jackpots
    for i = 1 to 5
        JackpotLights(i) = 0
    Next
    'setup random jackpots according to the nr of weapons
    j = RndNbr(5)
    For i = 1 to tmp
        do while JackpotLights(j) <> 0
            j = RndNbr(5)
        Loop
        JackpotLights(j) = 2
    Next
    UpdateJackpotLights
End Sub

Sub UpdateJackpotLights
    Light082.State = JackpotLights(1)
    Light081.State = JackpotLights(2)
    Light080.State = JackpotLights(3)
    Light002.State = JackpotLights(4)
    Light078.State = JackpotLights(5)
End Sub

'********************************
'       Pennywise Fight
' after all required ballons
'       are collected
'********************************
' Mode 3
' jackpot on the ramps only
' aim for ramp combos

Sub StartPennywiseMB
    bPennywise = True
    Mode = 3
    Dim i
    For Each i in aTiltLights:i.State = 0:Next 'turn all lights off
    GiOff
    GiRedOn
    UpdateHLights
    Jackpot(CurrentPlayer) = 50000
    'setup jackpot lights on the ramps
    JackpotLights(2) = 2
    JackpotLights(4) = 2
    UpdateJackpotLights
    ModeScore = 0
    PlaySong "m_mb2"
    DMD CL("PENNYWISE"), CL("MULTIBALL"), "_", eNone, eNone, eNone, 3000, True, "vo_pennywisemb"
    vpmTimer.AddTimer 3500, "PlaySound""vo_shootramps"" '"
    AddMultiball 1
End Sub

Sub StopPennywiseMB
    bPennywise = False
    Mode = 0
    GiOn
    GiRedOff
    UpdateLights 'chaos and hostages lights
    UpdateKillerLights
    ChangeSong
    BalloonsLeft(CurrentPlayer) = 99 'set the balloons needed for next mb
    ApronDMDUpdate
    DMD "  PENNYWISE SCORE", CL(FormatScore(ModeScore) ), "_", eNone, eNone, eNone, 3000, True, ""
End Sub

'********************************
'         Dracula Fight
'     after locking 3 balls
'********************************
' Mode 4

Sub StartDraculaMB
    dim i
    PlaySong "m_mb3"
    DMD CL("STARTING"), CL("DRACULA MULTIBALL"), "_", eNone, eNone, eNone, 2500, True, "vo_draculamb"
    DMD CL("SHOOT 5 JACKPOTS"), "AND THE SUPERJACKPOT", "_", eNone, eNone, eNone, 2500, True, ""
    bDracula = True
    Mode = 4
    'setup lights - all lights off - all Gi RED blinks
    For Each i in aTiltLights:i.State = 0:Next
    GiOff
    For Each i In aGiLightsRED:i.State = 2:Next
    Jackpot(CurrentPlayer) = 50000       'reset to 50000
    SuperJackpot(CurrentPlayer) = 250000 'reset to 250000
    'setup jackpot lights according to the nr of weapons collected
    SetupJackpots
    vpmTimer.AddTimer 3500, "AddMultiball 3 '"
    EnableBallSaver 20
    JackpotCount = 0
    ModeScore = 0
    Light077.State = 0
End Sub

Sub StopDraculaMB
    bDracula = False
    Mode = 0
    GiOn
    GiRedOff
    UpdateLights 'chaos and hostages lights
    UpdateKillerLights
    ChangeSong
    ApronDMDUpdate
    DMD "  DRACULA MB SCORE", CL(FormatScore(ModeScore) ), "_", eNone, eNone, eNone, 3000, True, ""
    'reset locked balls
    DT(CurrentPlayer, 1) = 0
    DT(CurrentPlayer, 2) = 0
    DT(CurrentPlayer, 3) = 0
    UpdateDT
    BallsInLock(CurrentPlayer) = 0
End Sub

'******************************************
' check for balls trapped behind the gates
'******************************************

Sub Trigger014_Hit
    Me.TimerEnabled = 0
    Me.TimerEnabled = 1
End Sub

Sub Trigger014_UnHit
    Me.TimerEnabled = 0
End Sub

Sub Trigger014_Timer
    Me.TimerEnabled = 0
    dgatef.RotateToEnd
    draculagate.IsDropped = 1
    draculagate.TimerEnabled = 1
End Sub

Sub draculagate_Timer
    Me.TimerEnabled = 0
    dgatef.RotateToStart
    draculagate.IsDropped = 0
End Sub

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
    'x = Table1.Option("Side Blades", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    'For each y in aSideBlades:y.SideVisible = x:next

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
    SongVolume = Table1.Option("Music Volume", 0, 1, 0.1, 0.1, 0)
    If bMusicOn Then
        ChangeSong
    Else
        StopSong
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

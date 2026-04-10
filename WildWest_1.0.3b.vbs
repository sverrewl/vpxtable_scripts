' ***********************************************************************
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'        ____  ____  ____  ____  _________  ____  ____  ____  ____
'       ||W ||||I ||||L ||||D ||||       ||||W ||||E ||||S ||||T ||
'       ||__||||__||||__||||__||||__ By __||||__||||__||||__||||__||
'       |/__\||/__\||/__\||/__\||/_TOMBG_\||/__\||/__\||/__\||/__\|
'
'     .oo ooo.   o    o o     ooooo .oPYo.     .oPYo. o    o o     o   o
'    .P 8 8  `8. 8    8 8       8   8          8    8 8b   8 8     `b d'
'   .P  8 8   `8 8    8 8       8   `Yooo.     8    8 8`b  8 8      `b'
'  oPooo8 8    8 8    8 8       8       `8     8    8 8 `b 8 8       8
' .P    8 8   .P 8    8 8       8        8     8    8 8  `b8 8       8
'.P     8 8ooo'  `YooP' 8oooo   8   `YooP'     `YooP' 8   `8 8oooo   8
'..:::::.......:::.....:......::..:::.....::::::.....:..:::........::..::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ***********************************************************************

'For DOF Do Not Use same Numbers that are being used by Controller.B2SSetData (See Bottom of Script)

'DOF Solenoid Config by Outhere
'100 Do Not Use (Controller.B2SSetData Interseres With DOF) At Bottom of Script
'101 Left Flipper
'102 Right Flipper
'103 Left Slingshot
'104 Right Slingshot
'105
'106
'107
'108
'109 Bumper Right
'110 Do Not Use (Controller.B2SSetData Interseres With DOF) At Bottom of Script
'111 Bumper Left
'112 Bumper Center
'113
'114
'115
'116 GiOn and GiOff
'117
'118
'119
'120 Do Not Use (Controller.B2SSetData Interseres With DOF) At Bottom of Script
'121 Kicker003
'122 Knocker
'123 BallRelease
'124
'125
'126
'127
'128
'129
'130 Do Not Use (Controller.B2SSetData Interseres With DOF) At Bottom of Script
'140 Do Not Use (Controller.B2SSetData Interseres With DOF) At Bottom of Script
'150 Do Not Use (Controller.B2SSetData Interseres With DOF) At Bottom of Script
'160 Do Not Use (Controller.B2SSetData Interseres With DOF) At Bottom of Script
'220
'221 Kicker003
'222
'223 .
'224
'225 Kicker005
'226 Kicker006
'227 Kicker009
'228
'229
'230 Kicker010
'231 Kicker007
'232 Kicker002
'233
'234
'235
'236
'*************************************************************************

Option Explicit

'*************************************************************************
' // User Settings in F12 menu //
'*************************************************************************

Dim LUTimage, ModeNudeActive
Dim ModeDifficulty
Dim ModeCheatActive

'----- VR Room -----
Dim VRRoomChoice : VRRoomChoice = 3       ' 1 - Cab Only, 2 - Minimal Room, 3 - MEGA room
Dim VRTest : VRTest = False

Sub Table1_OptionEvent(ByVal eventId)
  ' 10.8 only : called when options are tweaked by the player.
  '... 0: game has started, good time to load options and adjust accordingly
  '... 1: an option has changed
  '... 2: options have been reseted
  '... 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments... option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional Array for Menu Text fot that Option/Setting (strings)
'
  ' -- pause processes not needed to run --
    If eventId = 1 Then DisableStaticPreRendering = True
  DMDTimer.Enabled = False ' stop FlexDMD timer
  Controller.Pause = True

  ' --- Table LUT ---
  LUTimage = Table1.Option("Table Color Setting (LUT)", 0, 8, 1, 0, 0, Array("Natural","Natural Mid ","Natural Dark","Natural Very Dark","Natural Bright","Blue","Green","Red","Yellow"))
  Table1.ColorGradeImage = "LUT" & LUTimage
    '0 = Natural
    '1 = Natural Mid
    '2 = Natural Dark
    '3 = Natural Very Dark
    '4 = Natural Bright
    '5 = Blue
    '6 = Green
    '7 = Red
    '8 = Yellow

  ModeNudeActive = Table1.Option("Adult mode", 0, 1, 1, 0, 0, Array("Disabled","Enabled"))
  ModeDifficulty = Table1.Option("Mode Easy", 0, 1, 1, 0, 0, Array("Yes","No"))
  SetDifficulty ModeDifficulty
  ModeCheatActive = Table1.Option("Mode Cheat", 0, 1, 1, 0, 0, Array("No","Yes"))
  SetShowTriche ModeCheatActive

   ' VRRoom
  VRRoomChoice = Table1.Option("VR Room", 1, 3, 1, 3, 0, Array("CabOnly", "Minimal", "MEGA"))
  LoadVRRoom

  ' -- start paused processes --
  DMDTimer.Enabled = True
  Controller.Pause = False
    If eventId = 3 Then DisableStaticPreRendering = False

End Sub

Sub SetDifficulty(Opt)
  Select Case Opt
    Case 0:
      postrubber.Visible = 1
      zCol_Rubber_Post008.Collidable = 1
      post.visible = 1
    Case 1:
      postrubber.Visible = 0
      zCol_Rubber_Post008.Collidable = 0
      post.visible = 0
  End Select
End Sub

Sub SetShowTriche(Opt)
  Select Case Opt
    Case 0:
      Soutif.Visible = 1
      Soutif2.Visible = 1
      ModeCheatActive = False
      FlasherCheatMode.visible = 0
    Case 1:
      Soutif.Visible = 1
      If ModeNudeActive = 1 Then
      Soutif2.Visible = 0
      Elseif ModeNudeActive = 0 Then
      Soutif2.Visible = 1
      End If
      ModeCheatActive = True
      FlasherCheatMode.visible = 1
  End Select
End Sub
'****** PuP Variables ******

Dim usePUP: Dim cPuPPack: Dim PuPlayer: Dim PUPStatus: PUPStatus=false ' dont edit this line!!!

'*************************** PuP Settings for this table ********************************

usePUP   = true               ' enable Pinup Player functions for this table
cPuPPack = "WildWest"    ' name of the PuP-Pack / PuPVideos folder for this table

'//////////////////// PINUP PLAYER: STARTUP & CONTROL SECTION //////////////////////////

Sub PuPStart(cPuPPack)
    If PUPStatus=true then Exit Sub
    If usePUP=true then
        Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
        If PuPlayer is Nothing Then
            usePUP=false
            PUPStatus=false
        Else
            PuPlayer.B2SInit "",cPuPPack 'start the Pup-Pack
            PUPStatus=true
        End If
    End If
End Sub

Sub pupevent(EventNum)
    if (usePUP=false or PUPStatus=false) then Exit Sub
    PuPlayer.B2SData "E"&EventNum,1  'send event to Pup-Pack
End Sub

'************ PuP-Pack Startup **************

PuPStart(cPuPPack) 'Check for PuP - If found, then start Pinup Player / PuP-Pack

'///////////////////////////////////////////////

Randomize

Const BallSize = 50    ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1
Const SongVolume = 0.5 ' 1 is full volume.
'Const ModeNudeActive = 0 ' 1= Yes or 0=No  ' moved to Options Menu above.

' Define any Constants
Const cGameName = "WildWest"
Const myVersion = "1.0.3"
Const MaxPlayers = 4     ' from 1 to 4
Const BallSaverTime = 15 ' in seconds
Const MaxMultiplier = 3  ' limit to 3x in this game, both bonus multiplier and playfield multiplier
Const BallsPerGame = 5   ' usually 3 or 5

Const Special1 = 600000  ' High score to obtain an extra ball/game
Const Special2 = 1000000
Const Special3 = 1500000

' Use FlexDMD if in FS mode
Dim UseFlexDMD
If Table1.ShowDT = True then
    UseFlexDMD = False
Else
    UseFlexDMD = True
End If

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

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)

'Dim  KickerBonus(4)
Dim BonusMultiplier(4)
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Special1Awarded(4)
Dim Special2Awarded(4)
Dim Special3Awarded(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim bAttractMode
Dim x

' Define Game Control Variables
Dim BallsOnPlayfield
Dim D3TargetsDown
Dim D3TargetsMilieu

' Define Game Flags
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMusicOn
Dim bJustStarted
Dim Mission1_RODEO(4) 'Update ok
Dim Mission2_DUEL(4) 'Update ok
Dim Mission3_DYNAMITE(4)'Update ok
Dim Mission4_BANK(4) 'Update ok
Dim Mission5_MINE(4) 'Update ok
Dim Mission6_POKER(4) 'Update ok
Dim Coeur_Sav(4) 'Update ok
Dim Saloon_Sav(4) 'Update ok
Dim Hunder_Sav(4) 'Update ok
Dim MBLock_Sav(4)
Dim Mine_Sav(4) 'Update ok
Dim Bank_Sav(4) 'Update ok
Dim Star_Sav(4) 'Update ok
Dim Dyna_Sav(4) 'Update ok
Dim ModePokerActive(4) 'Update ok
Dim ModeRodeoActive(4)'Update ok
Dim ModeBankActive(4) 'Update ok
Dim ModeDynamiteActive(4) 'Update ok
Dim ModeDuelActive(4) 'Update ok
Dim ModeMineActive(4) 'Update ok
Dim ModeSpinnersActive(4) 'Update ok
Dim SpinnerCount 'Update ok
Dim ModeMultiballActive
Dim MusicActive : MusicActive = 1
Dim Song : Song = ""
Dim LastUpMusicTime : LastUpMusicTime = 0
Const TOTAL_SONGS = 12
Const DEBOUNCE_DELAY = 250
' core.vbs variables
Dim TopMagnet

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    LoadEM
  Dim i
  Light001.state=0
  Light002.state=0
  Light003.state=0
  li0RM1.state=2
  li0RR1.state=2
  Wall020.IsDropped = True
  FlasherOpenDoor2.visible = False
  ResetABC
  ResetMystery
    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' load saved values, highscore, names, jackpot
    Credits = 1
    Loadhs

    ' Initalise the DMD display
    DMD_Init

    ' freeplay or coins
    bFreePlay = True 'we don't want coins

    if bFreePlay Then DOF 125, DOFOn

    ' Init main variables and any other flags
    bAttractMode = False
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
  EnableBallSaver 15
    bGameInPlay = False
    bMusicOn = True
    BallsOnPlayfield = 0
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bJustStarted = True
    ' set any lights for the attract mode
    GiOff
    StartAttractMode

    ' Start the RealTime timer
    RealTime.Enabled = 1

End Sub

'****************************************
' Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub Realtime_Timer
    'RollingUpdate
    ' add any other real time update subs, like gates or diverters
    LeftFlipperTop.Rotz = LeftFlipper.CurrentAngle
    RightFlipperTop.Rotz = RightFlipper.CurrentAngle
  RightFlipperTop2.Rotz = RightFlipper2.CurrentAngle
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    If Keycode = AddCreditKey Then
    DOF 202, DOFPulse
        Credits = Credits + 1
        if bFreePlay = False Then
            DOF 125, DOFOn
            If(Tilted = False) Then
                DMDFlush
                DMD "_", CL(1, "CREDITS: " & Credits), "", eNone, eNone, eNone, 500, True, "fx_coin"
            End If
        End If
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
        If keycode = MechanicalTilt Then CheckTilt

'   If keycode = RightMagnaSave Then UpMusic
'   If keycode = LeftMagnasave Then DownMusic 'ShowTriche
    Select Case Keycode
      Case RightMagnaSave: UpMusic
      Case LeftMagnaSave: DownMusic
        ' Ajoutez d'autres touches ici si besoin
    End Select

        If keycode = LeftFlipperKey Then
             FlipperActivate LeftFlipper, LFPress 'nFozzy
             SolLFlipper 1
        End If
        If keycode = RightFlipperKey Then
             FlipperActivate RightFlipper, RFPress 'nFozzy
             SolRFlipper 1
        End If

        If keycode = StartGameKey Then
            If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 500, True, "so_fanfare1"
                Else
                    If(Credits> 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
                        DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 500, True, "so_fanfare1"
                        If Credits <1 And bFreePlay = False Then DOF 125, DOFOff
                        Else
                            ' Not Enough Credits to start a game.
                            DMD CL(0, "CREDITS " & Credits), CL(1, "INSERT COIN"), "", eNone, eBlink, eNone, 500, True, "so_nocredits"
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
                        DMD CL(0, "CREDITS " & Credits), CL(1, "INSERT COIN"), "", eNone, eBlink, eNone, 500, True, "so_nocredits"
                    End If
                End If
            End If
    End If ' If (GameInPlay)

'table keys


End Sub

Sub Table1_KeyUp(ByVal keycode)

    If hsbModeActive Then
        Exit Sub
    End If

    If keycode = PlungerKey Then
        Plunger.Fire
        PlaySoundAt "fx_plunger", plunger
    End If

    ' Table specific

    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then
            FlipperDeActivate LeftFlipper, LFPress 'nFozzy
            SolLFlipper 0
        End If
        If keycode = RightFlipperKey Then
            FlipperDeActivate RightFlipper, RFPress 'nFozzy
            SolRFlipper 0
        End If
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
    If B2SOn = true Then Controller.Stop
End Sub

'********************
'     Flippers
'********************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire
    DOF 101, DOFOn
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
            'PlaySoundAt SoundFXDOF("", 101, DOFOn, DOFFlippers), LeftFlipper
      RandomSoundReflipUpLeft LeftFlipper
            LeftFlipperOn = 1
        RotateLaneLightsLeftUp
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    DOF 101, DOFOff
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
            LeftFlipperOn = 0
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire
    DOF 102, DOFOn
    RightFlipper2.RotateToEnd
        RightFlipperOn = 1
    RotateLaneLightsRightUp
        'PlaySoundAt SoundFXDOF("", 102, DOFOn, DOFFlippers), RightFlipper
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
            'PlaySoundAt SoundFXDOF("", 102, DOFOn, DOFFlippers), RightFlipper
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    RightFlipper2.RotateToStart
    DOF 102, DOFOff
        RightFlipperOn = 0
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm 'nFozzy
    'PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm 'nFozzy
    'PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub

Sub RightFlipper2_Collide(parm)
  RightFlipperCollide parm
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
FullStrokeEOS_Torque = 0.3  ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2  ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.1
SOSAngle = 6

LiveCatchSensivity = 10

LLiveCatchTimer = 0
RLiveCatchTimer = 0

LeftFlipper.TimerInterval = 1
'LeftFlipper.TimerEnabled = 1 'nFozzy

Sub LeftFlipper_Timer 'flipper's tricks timer
'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If LeftFlipperOn = 1 Then
    If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle then
      LeftFlipper.EOSTorque = FullStrokeEOS_Torque
      LLiveCatchTimer = LLiveCatchTimer + 1
      If LLiveCatchTimer < LiveCatchSensivity Then
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
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If RightFlipperOn = 1 Then
    If RightFlipper.CurrentAngle = RightFlipper.EndAngle Then
      RightFlipper.EOSTorque = FullStrokeEOS_Torque
      RLiveCatchTimer = RLiveCatchTimer + 1
      If RLiveCatchTimer < LiveCatchSensivity Then
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

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                    'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity) AND(Tilt <15) Then 'show a warning
        DMD "_", CL(1, "CAREFUL"), "_", eNone, eBlinkFast, eNone, 500, True, ""
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        DMDFlush
        DMD "", "", "TILT", eNone, eNone, eBlink, 200, False, ""
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
        Bumper1.Force = 0
        Bumper2.Force = 0
    Bumper3.Force = 0
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        Bumper1.Force = 10
        Bumper2.Force = 10
    Bumper3.Force = 10
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

'*****************************
' Music as wav/mp3/ogg sounds
'*****************************
Dim SongDurations
SongDurations = Array(0, 251, 217, 152, 218, 279, 132, 306, 352, 220, 279, 253, 200)

Sub PlaySong(name)
    ' On arrête proprement la musique précédente
    StopSound Song

    ' On définit la nouvelle
    Song = name
    PlaySound Song, 0, SongVolume ' 0 = pas de boucle / -1 Loop / 1 play x2

    ' On règle le Timer sur la durée prévue dans l'Array
  If MusicActive >= 1 Then
    MusicTimer.Interval = SongDurations(MusicActive) * 1000
    MusicTimer.Enabled = True
    Debug.Print "Lecture : " & Song & " (Suivant dans " & SongDurations(MusicActive) & "s)"
  End If
End Sub

Sub StopSong
    If Song <> "" Then
        StopSound Song
        Song = ""
    End If
End Sub

Sub ChangeSong
    If BallsOnPlayfield = 0 Then
        PlaySong "m_end"
        MusicActive = 0
        Exit Sub
    End If

    If bAttractMode Then
        PlaySong "m_Attrack"
        MusicActive = 0
        Exit Sub
    End If

  If ModeMultiballActive = True Then
    Playsong "m_MuMb"
        Exit Sub
  End If
    ' Sélection aléatoire entre 1 et TOTAL_SONGS
    MusicActive = Int(Rnd * TOTAL_SONGS) + 1
    PlaySong "Mu_" & MusicActive
End Sub

Sub MusicTimer_Timer()
        Debug.Print "Fin de durée : Passage auto."
        UpMusic
End Sub

Sub UpMusic
    If (GameTime - LastUpMusicTime) < DEBOUNCE_DELAY Then Exit Sub
    LastUpMusicTime = GameTime

    If BallsOnPlayfield >= 1 And ModeMultiballActive = False Then
        MusicActive = MusicActive + 1
        If MusicActive > TOTAL_SONGS Then MusicActive = 1
        Debug.Print "Action : Monter Musique"
        PlaySong "Mu_" & MusicActive ' On appelle PlaySong avec un G
    End If
End Sub

Sub DownMusic
    If (GameTime - LastUpMusicTime) < DEBOUNCE_DELAY Then Exit Sub
    LastUpMusicTime = GameTime

    If BallsOnPlayfield >= 1 And ModeMultiballActive = False Then
        MusicActive = MusicActive - 1
        If MusicActive < 1 Then MusicActive = TOTAL_SONGS
        Debug.Print "Action : Descendre Musique"
        PlaySong "Mu_" & MusicActive ' On appelle PlaySong avec un G
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
    DOF 118, DOFOn
  PlaySound"fx_gion"
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
    For each bulb in aBumperLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    DOF 118, DOFOff
  PlaySound"fx_gioff"
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    For each bulb in aBumperLights
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
            LightSeqGi.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 10, 10
'        Case 4 'all blink once
'            LightSeqGi.UpdateInterval = 10
'            LightSeqGi.Play SeqBlinking, , 4, 1
    End Select
End Sub

Sub LightEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqInserts.Play SeqAlloff
        Case 1 'all blink
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 10, 10
        Case 4 'up 1 time
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 8, 1
        Case 5 'up 2 times
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 8, 2
        Case 6 'down 1 time
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqDownOn, 8, 1
        Case 7 'down 2 times
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqDownOn, 8, 2
    End Select
End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v4.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 34 'max ball velocity
ReDim rolling(tnob)
InitRolling

'******************************
' Diverse Collection Hit Sounds
'******************************

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
  StopSong
  PlaySound "Vo_WelcomeHG"
  If B2SOn Then
    CheckB2S
  End If
' pupevent 114
    GiOn
  ResetABC
  ResetMystery

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
'   KickerBonus(i) = 0
        Score(i) = 0
        BonusPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
        Special1Awarded(i) = False
        Special2Awarded(i) = False
        Special3Awarded(i) = False
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
' CheckLightEvent
  CheckMissionActive
    ' create a new ball in the shooters lane
    CreateNewBall()
End Sub

' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))

Sub ResetForNewPlayerBall()
    ' make sure the correct display is upto date
    AddScore 0
  ResetABC
  ResetMystery
  EnableBallSaver 15
' CheckLightEvent
  CheckMissionActive
    ' set the current players bonus multiplier back down to 1X
    BonusMultiplier(CurrentPlayer) = 1

    ' reset any drop targets, lights, game Mode etc..

    BonusPoints(CurrentPlayer) = 0

    'Reset any table specific
  Dim i
    For i = 1 To MaxPlayers
'   KickerBonus(i) = 0
    Next
    ResetNewBallVariables
''    ResetNewBallLights()
  ResetLightBall2345
End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    RandomSoundBallRelease BallRelease
    PlaySoundAt SoundFXDOF("", 123, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4

  'only this table
' ChangeBallImage
End Sub

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded

Sub EndOfBall()
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False
  ResetABC
  ResetMystery

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)

    If NOT Tilted Then
        BonusCountTimer.Interval = 300
        BonusCountTimer.Enabled = 1
    Else 'Si hay falta simplemente espera un momento y va directo a la segunta parte después de perder la bola
        vpmtimer.addtimer 400, "EndOfBall2 '"
    End If
End Sub

Sub BonusCountTimer_Timer 'Add bonus and update the bonus lights
    'debug.print "BonusCount_Timer"
    If BonusPoints(CurrentPlayer)> 0 Then
        BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) -1
        AddScore 1000 * BonusMultiplier(CurrentPlayer)
'        UpdateBonusLights
    Else
        ' end of bonus, go to end of ball
        BonusCountTimer.Enabled = 0
        vpmtimer.addtimer 1000, "EndOfBall2 '"
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
  ResetABC
  ResetMystery
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
        DMD CL(0, "EXTRA BALL"), CL(1, "SHOOT AGAIN"), "", eNone, eNone, eBlink, 1000, True, ""

        ' reset the playfield for the new ball
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
    ShowTableInfo
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
            PlaySound "vo_player" &CurrentPlayer
            DMD "_", CL(1, "PLAYER " &CurrentPlayer), "_", eNone, eNone, eNone, 800, True, ""
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

    ' set any lights for the attract mode
    GiOff

  ResetABC
  ResetMystery
    StartAttractMode
  If ModeCheatActive = False Then
  ResetLightEvent
  Elseif ModeCheatActive = True Then
  ResetLightEvent2
  End If
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

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' pretend to knock the ball into the ball storage mech
    'PlaySoundAt "fx_drain", Drain
    RandomSoundDrain drain
    'if Tilted then end Ball Mode
    If Tilted Then
        StopEndOfBallMode
    End If

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True) AND(Tilted = False) Then

        ' is the ball saver active,
'        If(bBallSaverActive = True) Then
'   If LightBallSave.state = 1 or LightBallSave.state = 2 Then
    If BallsOnPlayfield = 1 And ModeMultiballActive = True Then
      Mode_Multiball_End
      li0TD001.state = 2 'MB
      li0TD002.state = 2
      li0TD003.state = 2
      li0TD004.state = 2
      li0TD005.state = 2
      liLOCK.state=0
      liMB.state=0
    End If
    If BallsOnPlayfield = 0 and LightBallSave.State = 2 or BallsOnPlayfield = 0 and LightBallSave.State = 1 Then
            ' yep, create a new ball in the shooters lane
            ' we use the Addmultiball in case the multiballs are being ejected
'     CreateNewBall()
      SaveBall
            ' you may wish to put something on a display or play a sound at this point
            DMD CL(0, "BALL SAVED"), CL(1, "SHOOT AGAIN"), "_", eBlink, eBlink, eNone, 800, True, ""
        Else
            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0) Then
                ' End Mode and timers
        StopSong
        ShakeJailON
        li0TD001.state = 2 'MB
        li0TD002.state = 2
        li0TD003.state = 2
        li0TD004.state = 2
        li0TD005.state = 2
        liLOCK.state=0
        liMB.state=0
        ChangeBall(0)
'         If ModeMultiballActive = True Then
'         Mode_Multiball_End
'         li0TD001.state = 2 'MB
'         li0TD002.state = 2
'         li0TD003.state = 2
'         li0TD004.state = 2
'         li0TD005.state = 2
'         liLOCK.state=0
'         liMB.state=0
'         End If
        PlaySoundEndBall
                StopEndOfBallMode
                vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, since there is no animation then move to the end of ball
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
  DOF 203, DOFOn
  StopSong
    DMDScoreNow
    bBallInPlungerLane = True
End Sub

' The ball is released from the plunger

Sub swPlungerRest_UnHit()
  DOF 203, DOFOff
    DOF 215, DOFPulse
    bBallInPlungerLane = False
    LightEffect 4
  ChangeSong
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds

Sub swPlungerRest_Timer
    DMD "_", CL(1, "SHOOT THE BALL"), "_", eNone, eBlink, eNone, 800, True, ""
    swPlungerRest.TimerEnabled = 0
End Sub
'********************
' BALL SAVE
'********************

Sub SaveBall
  DOF 204, DOFPulse
  KickerAuto.CreateBall
  PlaySound "fx_kicker"
  KickerAuto.Kick 10, 30
' pupevent 110
  BallsOnPlayfield = 1
  LightShootAgain.State = 0
End Sub


Sub EnableBallSaver(seconds)
  bBallSaverActive = True
  bBallSaverReady = False
  BallSaverTimerExpired.Interval = 1000 * seconds
  BallSaverTimerExpired.Enabled = True
  BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
  BallSaverSpeedUpTimer.Enabled = True
  LightBallSave.BlinkInterval = 160
  LightBallSave.State = 2
End Sub

Sub SkillShootTimerExpired_Timer()
  liSskillShoot.State = 0
  SkillShootTimerExpired.Enabled = False
End Sub

Sub BallSaverTimerExpired_Timer()
  BallSaverTimerExpired.Enabled = False
  bBallSaverActive = False
  LightBallSave.State = 0
  liSskillShoot.State = 0
End Sub

Sub BallSaverSpeedUpTimer_Timer()
  BallSaverSpeedUpTimer.Enabled = False
  LightBallSave.BlinkInterval = 80
  LightBallSave.State = 2
End Sub
' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board
Sub AddScore(points)
    If Tilted Then Exit Sub

    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + points

    ' play a sound for each score

    ' you may wish to check to see if the player has gotten an extra ball by a high score
    If Score(CurrentPlayer) >= Special1 AND Special1Awarded(CurrentPlayer) = False Then
        AwardExtraBall
        Special1Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special2 AND Special2Awarded(CurrentPlayer) = False Then
        AwardExtraBall
        Special2Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special3 AND Special3Awarded(CurrentPlayer) = False Then
        AwardExtraBall
        Special3Awarded(CurrentPlayer) = True
    End If
End Sub

' Add bonus to the bonuspoints AND update the score board
Sub AddBonus(points) 'not used in this table, since there are many different bonus items.
    If Tilted Then Exit Sub
    ' add the bonus to the current players bonus variable
    BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
    If BonusPoints(CurrentPlayer)> 15 Then
        BonusPoints(CurrentPlayer) = 15
    End If
    ' Update the lights
End Sub

Sub AwardExtraBall()
  DMD "","","EXTRABALL (1)",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","EXTRABALL (2)",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","EXTRABALL (3)",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","EXTRABALL (4)",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","EXTRABALL (5)",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","EXTRABALL (6)",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","EXTRABALL (7)",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","EXTRABALL (8)",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","EXTRABALL (9)",eNone, eNone, eNone, 2600, TRUE, SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
  DOF 121, DOFPulse
' pupevent 111
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    LightShootAgain.State = 1
    LightEffect 2
End Sub

Sub AwardALLMISSION()
  DMD "ALL MISSIONS","    COMPLETED","",eNone, eNone, eNone, 3000, TRUE, SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
  DOF 121, DOFPulse
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    LightShootAgain.State = 1
    LightEffect 2
End Sub
'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(cGameName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 200000 End If
    x = LoadValue(cGameName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If
    x = LoadValue(cGameName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 200000 End If
    x = LoadValue(cGameName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If
    x = LoadValue(cGameName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 200000 End If
    x = LoadValue(cGameName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If
    x = LoadValue(cGameName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 200000 End If
    x = LoadValue(cGameName, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If
    x = LoadValue(cGameName, "Credits")
    If(x <> "") then Credits = CInt(x) Else Credits = 0:If bFreePlay = False Then DOF 125, DOFOff:End If
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
    HighScore(0) = 200000
    HighScore(1) = 300000
    HighScore(2) = 400000
    HighScore(3) = 600000
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
    ChangeSong
    hsLetterFlash = 0

    hsEnteredDigits(0) = " "
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<" ' ` is back arrow
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
    dLine(0) = ExpandLine(TempTopStr, 0)
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
    dLine(1) = ExpandLine(TempBotStr, 1)
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
    ChangeSong
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

' *************************************************************************
'   JP's Reduced Display Driver Functions (based on script by Black)
' only 5 effects: none, scroll left, scroll right, blink and blinkfast
' 3 Lines, treats all 3 lines as text. 3rd line is just 1 character
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

Dim dCharsPerLine(2)
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
            FlexDMD.TableFile = Table1.Filename & ".vpx"
            FlexDMD.RenderMode = 2
            FlexDMD.Width = 128
            FlexDMD.Height = 32
            FlexDMD.Clear = True
            FlexDMD.GameName = cGameName
            FlexDMD.Run = True
            Set DMDScene = FlexDMD.NewGroup("Scene")
            DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.bkempty")
            DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
            For i = 0 to 40
                DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.dempty&dmd=2")
                Digits(i).Visible = False
            Next
            digitgrid.Visible = False
            For i = 0 to 19 ' Top
                DMDScene.GetImage("Dig" & i).SetBounds 4 + i * 6, 3, 7, 11
            Next
            For i = 20 to 39 ' Bottom
                DMDScene.GetImage("Dig" & i).SetBounds 4 + (i - 20) * 6, 3 + 12 + 2, 7, 11
            Next
            FlexDMD.LockRenderThread
            FlexDMD.Stage.AddActor DMDScene
            FlexDMD.UnlockRenderThread
        End If
    End If

    Dim i, j
    DMDFlush()
    deSpeed = 20
    deBlinkSlowRate = 5
    deBlinkFastRate = 2
    dCharsPerLine(0) = 20 'characters lower line
    dCharsPerLine(1) = 20 'characters top line
    dCharsPerLine(2) = 1  'characters back line
    For i = 0 to 2
        dLine(i) = Space(dCharsPerLine(i))
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
        dqTimeOn(i) = 0
        dqbFlush(i) = True
        dqSound(i) = ""
    Next
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
    Dim tmp, tmp1, tmp2
    if(dqHead = dqTail) Then
        tmp = RL(0, FormatScore(Score(Currentplayer) ) )
        tmp1 = CL(1, "PLAYER " & CurrentPlayer & "  BALL " & Balls)
        tmp2 = "bkborder"
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
            dqText(0, dqTail) = ExpandLine(Text0, 0)
        End If

        if(Text1 = "_") Then
            dqEffect(1, dqTail) = eNone
            dqText(1, dqTail) = "_"
        Else
            dqEffect(1, dqTail) = Effect1
            dqText(1, dqTail) = ExpandLine(Text1, 1)
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
                    Temp = Right(dLine(i), dCharsPerLine(i) - 1)
                    Temp = Temp & Mid(dqText(i, dqHead), deCount(i), 1)
                case eScrollRight:
                    Temp = Mid(dqText(i, dqHead), (dCharsPerLine(i) + 1) - deCount(i), 1)
                    Temp = Temp & Left(dLine(i), dCharsPerLine(i) - 1)
                case eBlink:
                    BlinkEffect = True
                    if((deCount(i) MOD deBlinkSlowRate) = 0) Then
                        deBlinkCycle(i) = deBlinkCycle(i) xor 1
                    End If

                    if(deBlinkCycle(i) = 0) Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i) )
                    End If
                case eBlinkFast:
                    BlinkEffect = True
                    if((deCount(i) MOD deBlinkFastRate) = 0) Then
                        deBlinkCycle(i) = deBlinkCycle(i) xor 1
                    End If

                    if(deBlinkCycle(i) = 0) Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i) )
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

Function ExpandLine(TempStr, id) 'id is the number of the dmd line
    If TempStr = "" Then
        TempStr = Space(dCharsPerLine(id) )
    Else
        if(Len(TempStr)> Space(dCharsPerLine(id) ) ) Then
            TempStr = Left(TempStr, Space(dCharsPerLine(id) ) )
        Else
            if(Len(TempStr) <dCharsPerLine(id) ) Then
                TempStr = TempStr & Space(dCharsPerLine(id) - Len(TempStr) )
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

Function CL(id, NumString)
    Dim Temp, TempStr
    Temp = (dCharsPerLine(id) - Len(NumString) ) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(id, NumString)
    Dim Temp, TempStr
    Temp = dCharsPerLine(id) - Len(NumString)
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
            If dLine(2) = "" OR dLine(2) = " " Then dLine(2) = "bkempty"
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
    For i = 0 to 255:Chars(i) = "dempty":Next

    Chars(32) = "dempty"
    Chars(35) = "jp1"
    Chars(36) = "jp2"
    Chars(37) = "jp3"
    Chars(38) = "title1"
    Chars(39) = "title2"
    Chars(40) = "title3"
    Chars(46) = "dot"     '.
    Chars(48) = "d0"      '0
    Chars(49) = "d1"      '1
    Chars(50) = "d2"      '2
    Chars(51) = "d3"      '3
    Chars(52) = "d4"      '4
    Chars(53) = "d5"      '5
    Chars(54) = "d6"      '6
    Chars(55) = "d7"      '7
    Chars(56) = "d8"      '8
    Chars(57) = "d9"      '9
    Chars(60) = "dless"   '<
    Chars(61) = "dequal"  '=
    Chars(62) = "dmore"   '>
    Chars(64) = "bkempty" '@
    Chars(65) = "da"      'A
    Chars(66) = "db"      'B
    Chars(67) = "dc"      'C
    Chars(68) = "dd"      'D
    Chars(69) = "de"      'E
    Chars(70) = "df"      'F
    Chars(71) = "dg"      'G
    Chars(72) = "dh"      'H
    Chars(73) = "di"      'I
    Chars(74) = "dj"      'J
    Chars(75) = "dk"      'K
    Chars(76) = "dl"      'L
    Chars(77) = "dm"      'M
    Chars(78) = "dn"      'N
    Chars(79) = "do"      'O
    Chars(80) = "dp"      'P
    Chars(81) = "dq"      'Q
    Chars(82) = "dr"      'R
    Chars(83) = "ds"      'S
    Chars(84) = "dt"      'T
    Chars(85) = "du"      'U
    Chars(86) = "dv"      'V
    Chars(87) = "dw"      'W
    Chars(88) = "dx"      'X
    Chars(89) = "dy"      'Y
    Chars(90) = "dz"      'Z
    Chars(94) = "dup"     '^
    '    Chars(95) = '_
    Chars(96) = "d0a"  '0.
    Chars(97) = "d1a"  '1. 'a
    Chars(98) = "d2a"  '2. 'b
    Chars(99) = "d3a"  '3. 'c
    Chars(100) = "d4a" '4. 'd
    Chars(101) = "d5a" '5. 'e
    Chars(102) = "d6a" '6. 'f
    Chars(103) = "d7a" '7. 'g
    Chars(104) = "d8a" '8. 'h
    Chars(105) = "d9a" '9  'i
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

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo
    Dim ii
    'info goes in a loop only stopped by the credits and the startkey
    If Score(1) Then
        DMD CL(0, "LAST SCORE"), CL(1, "PLAYER 1 " &FormatScore(Score(1) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(2) Then
        DMD CL(0, "LAST SCORE"), CL(1, "PLAYER 2 " &FormatScore(Score(2) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(3) Then
        DMD CL(0, "LAST SCORE"), CL(1, "PLAYER 3 " &FormatScore(Score(3) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(4) Then
        DMD CL(0, "LAST SCORE"), CL(1, "PLAYER 4 " &FormatScore(Score(4) ) ), "", eNone, eNone, eNone, 3000, False, ""
    End If
    DMD "", "", "gameover", eNone, eNone, eBlink, 2000, False, ""
    If bFreePlay Then
        DMD "", CL(1, "    FREE PLAY"), "", eNone, eNone, eNone, 2000, False, ""
    Else
        If Credits> 0 Then
            DMD CL(0, "CREDITS " & Credits), CL(1, "PRESS START"), "", eNone, eBlink, eNone, 2000, False, ""
        Else
            DMD CL(0, "CREDITS " & Credits), CL(1, "INSERT COIN"), "", eNone, eBlink, eNone, 2000, False, ""
        End If
    End If
    DMD "", "", "presents", eNone, eNone, eNone, 3000, False, ""
    DMD "", "", "WW_001", eNone, eNone, eNone, 120, False, ""
    DMD "", "", "WW_002", eNone, eNone, eNone, 120, False, ""
    DMD "", "", "WW_003", eNone, eNone, eNone, 120, False, ""
    DMD "", "", "WW_004", eNone, eNone, eNone, 120, False, ""
    DMD "", "", "WW_005", eNone, eNone, eNone, 120, False, ""
    DMD "", "", "WW_006", eNone, eNone, eNone, 120, False, ""
    DMD "", "", "WW_007", eNone, eNone, eNone, 120, False, ""
    DMD "", "", "WW_008", eNone, eNone, eNone, 3100, False, ""
    DMD CL(0, "HIGHSCORES"), Space(dCharsPerLine(1) ), "", eScrollLeft, eScrollLeft, eNone, 20, False, ""
    DMD CL(0, "HIGHSCORES"), "", "", eBlinkFast, eNone, eNone, 1000, False, ""
    DMD CL(0, "HIGHSCORES"), "1> " &HighScoreName(0) & " " &FormatScore(HighScore(0) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "2> " &HighScoreName(1) & " " &FormatScore(HighScore(1) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "3> " &HighScoreName(2) & " " &FormatScore(HighScore(2) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "4> " &HighScoreName(3) & " " &FormatScore(HighScore(3) ), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD Space(dCharsPerLine(0) ), Space(dCharsPerLine(1) ), "", eScrollLeft, eScrollLeft, eNone, 500, False, ""
End Sub

Sub StartAttractMode
  DOF 201, DOFOn
    ChangeSong
    StartLightSeq
    DMDFlush
    ShowTableInfo
' pupevent 110
End Sub

Sub StopAttractMode
  DOF 201, DOFOff
    LightSeqAttract.StopPlay
    DMDScoreNow
End Sub

Sub StartLightSeq()
    'lights sequences
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



Dim GiIntensity
GiIntensity = 1   'can be used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************

'Dim bFlipperPost
Dim bRedBall

' droptargets, animations, etc
Sub VPObjects_Init
End Sub

' tables variables and Mode init
Sub Game_Init() 'called at the start of a new game
    Dim i, j
  For i = 1 To MaxPlayers
'   KickerBonus(i) = 0
    Next
    'Play some Music
'    ChangeSong
    'Init Variables
' bFlipperPost = True
  bRedBall = False
  If ModeCheatActive = False Then
  ResetLightEvent
  Elseif ModeCheatActive = True Then
  ResetLightEvent2
  End If
End Sub

Sub StopEndOfBallMode()     'this sub is called after the last ball is drained
  DOF 205, DOFPulse
End Sub

Sub ResetNewBallVariables() 'reset variables for a new ball or player
Dim i
End Sub


' *********************************************************************
'                        Table Object Hit Events
'
' Any target hit Sub should do something like this:
' - play a sound
' - do some physical movement
' - add a score, bonus
' - check some variables/Mode this trigger is a member of
' - set the "LastSwitchHit" variable in case it is needed later
' *********************************************************************

'************
' Slingshots
'************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    LS.VelocityCorrect(Activeball) 'nFozzy
  RandomSoundSlingshotLeft Lemk
    PlaySoundAt SoundFXDOF("", 103, DOFPulse, DOFcontactors), Lemk
    DOF 105, DOFPulse
  GiEffect 4
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    AddScore 100
  If ModeDuelActive(CurrentPlayer) = 2 or ModeMultiballActive = True Then
  ShakeGunL1
  AddScore 400
  Elseif ModeDuelActive(CurrentPlayer) <= 1 And ModeMultiballActive = False Then
  ShakeLeftGun
  End If
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
    RS.VelocityCorrect(Activeball) 'nFozzy
  RandomSoundSlingshotRight Remk
    PlaySoundAt SoundFXDOF("", 104, DOFPulse, DOFcontactors), Remk
    DOF 106, DOFPulse
  GiEffect 4
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    AddScore 100
  If ModeDuelActive(CurrentPlayer) = 2 or ModeMultiballActive = True Then
  ShakeGunR1
  AddScore 400
  Elseif ModeDuelActive(CurrentPlayer) <= 1 And ModeMultiballActive = False Then
  ShakeRightGun
  End If
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'*********
' Bumpers
'*********

Sub Bumper1_Hit
    If NOT Tilted Then
        RandomSoundBumperMiddle Bumper1
        PlaySoundAt SoundFXDOF("", 109, DOFPulse, DOFContactors), Bumper1
        DOF 138, DOFPulse
    GiEffect 4
    DMDFlush
        DMD "  BUMPER", "   HITS", "Bumper_001", eNone, eBlink, eNone, 450, True, ""
        AddScore 100
    End If
End Sub

Sub Bumper2_Hit
    If NOT Tilted Then
        RandomSoundBumperTop Bumper2
        PlaySoundAt SoundFXDOF("", 111, DOFPulse, DOFContactors), Bumper2
        DOF 140, DOFPulse
    GiEffect 4
    DMDFlush
        DMD "  BUMPER", "   HITS", "Bumper_001", eNone, eBlink, eNone, 450, True, ""
        AddScore 100
    End If
End Sub

Sub Bumper3_Hit
    If NOT Tilted Then
        RandomSoundBumperBottom Bumper3
        PlaySoundAt SoundFXDOF("", 112, DOFPulse, DOFContactors), Bumper3
        DOF 141, DOFPulse
    GiEffect 4
    DMDFlush
        DMD "  BUMPER", "   HITS", "Bumper_001", eNone, eBlink, eNone, 450, True, ""
        AddScore 100
    End If
End Sub

'*********
' Rubbers
'*********

'50 points
Sub RubberBand003_Hit:Addscore 50:End Sub
Sub RubberBand006_Hit:Addscore 50:End Sub

'10 points
Sub RubberBand007_Hit:Addscore 10:End Sub
Sub RubberBand008_Hit:Addscore 10:End Sub

'************
' ABC - Bonus X
'************

Sub CheckBonusX
  If FlasherBonus.ImageA = "Bonus3" Then
    FlasherBonus.ImageA = "Bonus5"
    FlasherBonusB.ImageA = "Bonus5b"
    dmdflush
    DMD "             BONUS", "              X 5", "", eScrollRight, eBlink, eNone, 4000, True, ""
'   PlaySound ""
        BonusMultiplier(CurrentPlayer) = 5
    ResetABC
    End If

  If FlasherBonus.ImageA = "Bonus2" Then
    FlasherBonus.ImageA = "Bonus3"
    FlasherBonusB.ImageA = "Bonus3b"
    dmdflush
    DMD "             BONUS", "              X 3", "", eScrollRight, eBlink, eNone, 4000, True, ""
        BonusMultiplier(CurrentPlayer) = 3
    ResetABC
    End If

  If FlasherBonus.ImageA = "Bonus1" Then
    FlasherBonus.ImageA = "Bonus2"
    FlasherBonusB.ImageA = "Bonus2b"
    dmdflush
    DMD "             BONUS", "              X 2", "", eScrollRight, eBlink, eNone, 4000, True, ""
        BonusMultiplier(CurrentPlayer) = 2
    ResetABC
    End If
End Sub

Sub CheckABC
  If li016.State = 1 And li017.State = 1 And li018.State = 1 Then
  CheckBonusX
  AddScore 25000
    PLaySound"Bonus0"&RndNbr(2)
    LightEffect 7
  End If
End Sub

Sub ResetABC
    li016.State = 0
    li017.State = 0
    li018.State = 0
End Sub

'***********************
' Outlanes - Extra Ball
'***********************

Sub Trigger001_Hit
  DOF 206, DOFPulse
  leftInlaneSpeedLimit
    PlaySoundAt "fx_sensor", Trigger001
    If Tilted Then Exit Sub
    AddScore 500
End Sub

Sub Trigger004_Hit
  DOF 209, DOFPulse
  RightInlaneSpeedLimit
    PlaySoundAt "fx_sensor", Trigger004
    If Tilted Then Exit Sub
    AddScore 500
End Sub

'**********
' Inlanes
'**********

Sub Trigger002_Hit
  DOF 207, DOFPulse
  leftInlaneSpeedLimit
    PlaySoundAt "fx_sensor", Trigger002
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
End Sub

Sub Trigger003_Hit
  DOF 208, DOFPulse
  RightInlaneSpeedLimit
    PlaySoundAt "fx_sensor", Trigger003
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
End Sub
'#################################
'FX SOUND
'#################################
Sub PlaySoundGun
  Select Case Int(Rnd * 6) + 1
    Case 1
      PlaySound "fx_Gun01"
    Case 2
      PlaySound "fx_Gun02"
    Case 3
      PlaySound "fx_Gun03"
    Case 4
      PlaySound "fx_Gun04"
    Case 5
      PlaySound "fx_Gun05"
    Case 6
      PlaySound "fx_Gun06"
  End Select
End Sub

Sub PlaySoundRire
  Select Case Int(Rnd * 3) + 1
    Case 1
      PlaySound "fx_RireLady001"
    Case 2
      PlaySound "fx_RireLady002"
    Case 3
      PlaySound "fx_RireLady003"
  End Select
End Sub

Sub PlaySoundEndBall
  Select Case Int(Rnd * 4) + 1
    Case 1
      PlaySound "m_endball001"
    Case 2
      PlaySound "m_endball002"
    Case 3
      PlaySound "m_endball003"
    Case 4
      PlaySound "m_endball004"
  End Select
End Sub
'*********
' Targets
'*********

Sub Target001_Hit
  DOF 216, DOFPulse
  PlaySoundGun
    If Tilted Then Exit Sub
    li016.State = 1
    AddScore 1000
    AddBonus 1
    CheckABC
End Sub

Sub Target002_Hit
  DOF 217, DOFPulse
  PlaySoundGun
    If Tilted Then Exit Sub
  li017.State = 1
    AddScore 1000
    AddBonus 1
    CheckABC
End Sub

Sub Target003_Hit
  DOF 218, DOFPulse
  PlaySoundGun
    If Tilted Then Exit Sub
    li018.State = 1
    AddScore 1000
    AddBonus 1
    CheckABC
End Sub

Sub Trigger008_Hit()
  DOF 219, DOFPulse
  AddScore 500
  StartOpenDoor
End Sub

Sub Trigger009_Hit()
  WireRampOff
  DOF 220, DOFPulse
    DOF 229, DOFPulse
  Dim xx: For each xx in aGiLights: xx.color = RGB(0,0,255): Next 'Bleu
  Trigger009.DestroyBall
  If Wall020.IsDropped = True And Wall018.IsDropped = False Then
  Kicker001.CreateBall
  Kicker001.TimerInterval = 2000
  Kicker001.TimerEnabled = True
  Elseif Wall020.IsDropped = False And Wall018.IsDropped = True Then
  Kicker004.CreateBall
  Kicker004.TimerInterval = 2000
  Kicker004.TimerEnabled = True
  End If
End Sub

Sub Kicker003_Hit() 'Wall20 ouvert - False visible
    SoundSaucerLock
  'DOF 221, DOFPulse
  AddScore 500
  If Wall020.IsDropped = True And Wall018.IsDropped = False And li0Switch2.state = 1 Then
  Wall020.IsDropped = False
  FlasherOpenDoor2.visible = True
  Wall018.IsDropped = True
  PlaySound "OpenDoor"
' pupevent 101
  li0Switch1.state = 1
  li0Switch2.state = 0
  DMDFlush
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon2", eBlink, eNone, eNone, 200, False, ""
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon2", eBlink, eNone, eNone, 200, False, ""
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon2", eBlink, eNone, eNone, 200, False, ""
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon2", eBlink, eNone, eNone, 200, False, ""
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon2", eBlinK, eNone, eNone, 2000, True, ""
  Else
  Wall020.IsDropped = True
  FlasherOpenDoor2.visible = False
  Wall018.IsDropped = False
  DMDFlush
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon3", eBlink, eNone, eNone, 200, False, ""
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon3", eBlink, eNone, eNone, 200, False, ""
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon3", eBlink, eNone, eNone, 200, False, ""
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon3", eBlink, eNone, eNone, 200, False, ""
  DMD "", "", "saloon1", eNone, eNone, eNone, 200, False, ""
  DMD "", "", "saloon3", eBlinK, eNone, eNone, 2000, True, ""
  PlaySound "CloseDoor"
' pupevent 102
  li0Switch1.state = 0
  li0Switch2.state = 1
  End If
  Kicker003.TimerInterval = 2000
  Kicker003.TimerEnabled = True
End Sub

Sub CheckLightMB
  If Light001.state=0 And Light002.state=0 And Light003.state=0 Then
  light001.state=1
  MBLock_Sav(CurrentPlayer) = 1
' pupevent 109
  PlaySound "Vo_BallLock"
  DMD "","","BallLock",eBlink, eNone, eNone, 3500, True, ""
  DOF 222, DOFPulse
  liLOCK.State = 0
  Elseif Light001.state=1 And Light002.state=0 And Light003.state=0 Then
  Light002.state=1
  MBLock_Sav(CurrentPlayer) = 2
  PlaySound "Vo_BallLock"
' pupevent 109
  DMD "","","BallLock",eBlink, eNone, eNone, 3500, True, ""
   DOF 223, DOFPulse
  liLOCK.State = 0
  Elseif Light001.state=1 And Light002.state=1 And Light003.state=0 Then
  light003.state=1
  liMB.State = 0
  Mode_Multiball_Start
  DOF 224, DOFPulse
  ShakeChasseurON
' pupevent 105
  PlaySound "Vo_MB"
  Elseif Light001.state=1 And Light002.state=1 And Light003.state=1 Then
  light003.state=0
  liLOCK.State = 0
  End IF
End Sub

Sub Kicker005_Hit()
  DOF 225, DOFPulse
  PlaySound "fx_hole_enter"
  AddScore 500
  ResetLightJail
  CheckLightMB
  EnableBallSaver 15
' li020.state=1
' li021.state=1
  ActiveOutlineSave
  Kicker005.TimerInterval = 2000
  Kicker005.TimerEnabled = True
End Sub
Sub Kicker006_Hit()
    SoundSaucerLock
  PlaySound "fx_hole_enter"
  'DOF 226, DOFPulse
  AddScore 500
  li0START.state = 0
  If ModeBankActive(CurrentPlayer) = 1 Then StartMissionBANK
  If ModeMineActive(CurrentPlayer) = 1 Then StartMissionMINE
  If ModePokerActive(CurrentPlayer) = 1 Then StartMissionPOKER : ActiveOutlineSave 'li020.state=1 : li021.state=1
  If ModeDynamiteActive(CurrentPlayer) = 1 Then StartMissionDYNAMITE : ActiveOutlineSave 'li020.state=1 : li021.state=1
  If ModeDuelActive(CurrentPlayer) = 1 Then StartSHERIFFtarget : ActiveOutlineSave 'li020.state=1 : li021.state=1
  If ModeRodeoActive(CurrentPlayer) = 1 Then StartMissionRODEO

  If liSskillShoot.state = 2 Then
    AddScore 50000
    DOF 127, DOFPulse
    PlaySound "Vo_SkillShot"
'   pupevent 112
' Elseif liSskillShoot.state = 2 And ModeMineActive(CurrentPlayer) = False or liSskillShoot.state = 2 And liMISS005.state = 2 Then
' AddScore 5000
' PlaySound "Vo_SkillShot"
' pupevent 112
    DMD "","","SkillS_001",eNone, eNone, eNone, 150, FALSE, ""
    DMD "","","SkillS_002",eNone, eNone, eNone, 150, FALSE, ""
    DMD "","","SkillS_003",eNone, eNone, eNone, 150, FALSE, ""
    DMD "","","SkillS_004",eNone, eNone, eNone, 150, FALSE, ""
    DMD "","","SkillS_005",eNone, eNone, eNone, 150, FALSE, ""
    DMD "","","SkillS_006",eNone, eNone, eNone, 150, FALSE, ""
    DMD "","","SkillS_007",eNone, eNone, eNone, 150, FALSE, ""
    DMD "","","SkillS_008",eNone, eNone, eNone, 150, FALSE, ""
    DMD "","","SkillS_009",eNone, eNone, eNone, 3500, TRUE, ""
  End if
  If li0MYST.state = 2 Then
  Dim ModeChosen
    ModeChosen = False
    Do While ModeChosen = False
      Select Case RndNbr(9)
      Case 1
        DMDFlush
        DMD "  MYSTERY AWARD","  ALL LIGHTS HOT","",eNone, eBlink, eNone, 3500, TRUE, ""
        DOF 126, DOFPulse
        li0RR1.State = 1
        li0RR2.State = 1
        li0RR3.State = 1
        li0RR4.State = 2
        ModeChosen = True
      Case 2
        DMDFlush
        DMD "    MYSTERY AWARD","ALL LIGHTS SALOON","",eNone, eBlink, eNone, 3500, TRUE, ""
        DOF 126, DOFPulse
        Saloon_Sav(CurrentPlayer) = 6
        CheckLightSALOON
        ModeChosen = True
      Case 3
        DMDFlush
        DMD "  MYSTERY AWARD"," UP LIGHTS BANK","",eNone, eBlink, eNone, 3500, TRUE, ""
        DOF 126, DOFPulse
        Bank_Sav(CurrentPlayer) = Bank_Sav(CurrentPlayer) + 1
        CheckLightBANK
        ModeChosen = True
      Case 4
        DMDFlush
        DMD " MYSTERY AWARD","","",eBlink, eNone, eNone, 500, FALSE, ""
        DMD "","","jackpot",eNone, eNone, eNone, 3000, TRUE, ""
        DOF 126, DOFPulse
        AddScore 25000
        ModeChosen = True
      Case 5
        DMDFlush
        DMD " MYSTERY AWARD","  ","",eBlink, eNone, eNone, 500, FALSE, ""
        DMD "BONUS", "X3", "", eNone, eBlink, eNone, 3000, True, ""
        DOF 126, DOFPulse
        FlasherBonus.ImageA = "Bonus3"
        FlasherBonusB.ImageA = "Bonus3b"
        BonusMultiplier(CurrentPlayer) = 3
        ResetABC
        ModeChosen = True
      Case 6
        DMDFlush
        DMD " MYSTERY AWARD","  LOCK IS LIT","",eNone, eBlink, eNone, 3500, TRUE, ""
        DOF 126, DOFPulse
        liLOCK.state=2
        ShakeJailOFF
        ModeChosen = True
      Case 7
        DMDFlush
        DMD "  MYSTERY AWARD","  OUTLANE SAVE","",eNone, eBlink, eNone, 3500, TRUE, ""
        DOF 126, DOFPulse
'       li020.state=1
'       li021.state=1
        ActiveOutlineSave
        ModeChosen = True
      Case 8
        DMDFlush
        DMD "  MYSTERY AWARD","   ADD 1 BALL","",eNone, eBlink, eNone, 3500, TRUE, ""
        DOF 126, DOFPulse
        AddMBmission
        ModeChosen = True
      Case 9
        DMDFlush
        DMD "    MYSTERY AWARD","ALL LIGHTS UNDERTAKER","",eNone, eBlink, eNone, 3500, TRUE, ""
        DOF 126, DOFPulse
        Hunder_Sav(CurrentPlayer) = 10
        CheckLightHUNDERTAKER
        ModeChosen = True
      End Select
    Loop
  ResetMystery
  CheckB2S
  End If
  Kicker006.TimerInterval = 3000
  Kicker006.TimerEnabled = True
  CheckB2S
End Sub

Sub Kicker008_Hit()
    SoundSaucerLock
  PlaySound "fx_hole_enter"
  'DOF 227, DOFPulse
  AddScore 500
  Kicker008.TimerInterval = 2000
  Kicker008.TimerEnabled = True
End Sub

Sub TriggerTEST_Hit()
' CheckMissionActive
  If LightBallSave.State = 2 Then
  liSskillShoot.State = 2
  SkillShootTimerExpired.Enabled = True
  End If
  If ModeMultiballActive = False Then
    FlasherBonus.Visible = 1
    FlasherBonusB.Visible = 0
  Elseif ModeMultiballActive = True Then
    FlasherBonus.Visible = 0
    FlasherBonusB.Visible = 1
  End If
End Sub
'***********
' MbModeMission
'***********
Sub AddMBmission
  EnableBallSaver 15
  PlaySound "fx_kicker"
  KickerAuto.CreateBall
  KickerAuto.Kick 10, 30
  BallsOnPlayfield = BallsOnPlayfield +1
' li020.state=1 'Outline save
' li021.state=1
  ActiveOutlineSave
End Sub
Sub AddMBmission2 'Switch
  Kicker003.CreateBall
  Kicker003.TimerInterval = 4000
  Kicker003.TimerEnabled = True
  BallsOnPlayfield = BallsOnPlayfield +1
End Sub
Sub AddMBmission3 'HUNDERTEKER
  Kicker001.CreateBall
  Kicker001.TimerInterval = 6000
  Kicker001.TimerEnabled = True
  BallsOnPlayfield = BallsOnPlayfield +1
End Sub
Sub Mode_Multiball_Start
  DOF 228, DOFOn
  AddMBmission2
  AddMBmission3
  EnableBallSaver 15
  AddScore 5000
  ModeMultiballActive = True
  ChangeSong
' Playsong "m_MuMb"
  FlasherBonus.Visible = 0
  FlasherBonusB.Visible = 1
  If ModeNudeActive = 1 Then
  Soutif.Visible = 0
  Soutif2.Visible = 0
  Soutif3.Visible = 0
  Elseif ModeNudeActive = 0 Then
  Soutif.Visible = 1
  Soutif2.Visible = 1
  Soutif3.Visible = 1
  End If
  FlasherPFALT.visible = 1
  ChangeBall(3)
' pupevent 105
  Light001.state=2
  Light002.state=2
  Light003.state=2
' li020.state = 1
' li021.state = 1
  ActiveOutlineSave
  liSskillShoot.state = 0
  DMD "","","MB_001",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "","","MB_002",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "","","MB_003",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "","","MB_004",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "","","MB_005",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "","","MB_006",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "","","MB_007",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "","","MB_008",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "","","MB_009",eNone, eNone, eNone, 3500, TRUE, ""
End Sub
Sub Mode_Multiball_End
  DOF 228, DOFOff
  ModeMultiballActive = False
  ChangeSong
  FlasherBonus.Visible = 1
  FlasherBonusB.Visible = 0
  FlasherPFALT.visible = 0
  Soutif3.Visible = 0
  ChangeBall(0)
  Light001.state=0
  Light002.state=0
  Light003.state=0
  MBLock_Sav(CurrentPlayer) = 0
End Sub
'*****************
'Flasher DOOR
'*****************
Dim OpenDoorPos, FramesOpenDoor
FramesOpenDoor = Array("PS001", "PS002", "PS003", "PS004", "PS005", "PS006", "PS007", "PS008", "PS009", "PS010", "PS011", "PS012", "PS013", "PS014", "PS015", "PS016", "PS017", "PS018", "PS019", "PS020", "PS021")

Sub StartOpenDoor
  Playsound "fx_doorSaloon"
  FlasherOpenDoor.visible = 1
  OpenDoorPos = 0
    OpenDoorTimer.Enabled = 1
  StopOpenDoorTimer.Enabled = 1
End Sub
Sub OpenDoorTimer_Timer
    FlasherOpenDoor.ImageA = FramesOpenDoor (OpenDoorPos)
    OpenDoorPos = (OpenDoorPos + 1) MOD 21
End Sub
Sub StopOpenDoor
  FlasherOpenDoor.visible = 1
  OpenDoorTimer.Enabled = 0
End Sub
Sub StopOpenDoorTimer_Timer
  StopOpenDoor
Me.Enabled = 0
End Sub
'*****************
'Flasher GIRL
'*****************
Dim GirlPos, FramesGirl
FramesGirl = Array("Girl001", "Girl002", "Girl003", "Girl004", "Girl005", "Girl006", "Girl007", "Girl008", "Girl009", "Girl010", "Girl011", "Girl012", "Girl013", "Girl014", "Girl015", "Girl016", "Girl017", "Girl018", "Girl019", "Girl020", "Girl021", "Girl022")

Sub StartGirl
  PlaySound "CanCan"
  GirlLOOK 0
' pupevent 115
  FlasherGirl.visible = 1
  liGIRL002.state = 1
  GirlPos = 0
    GirlTimer.Enabled = 1
  StopGirlTimer.Enabled = 1
End Sub
Sub GirlTimer_Timer
  GirlLOOK 0
    FlasherGirl.ImageA = FramesGirl (GirlPos)
    GirlPos = (GirlPos + 1) MOD 22
End Sub
Sub StopGirl
  FlasherGirl.visible = 0
  GirlLOOK 1
  CheckCoeur_Sav
  liGIRL002.state = 1
  If ModeNudeActive = 1 Then
  Soutif.Visible = 0
  Elseif ModeNudeActive = 0 Then
  Soutif.Visible = 1
  End If
  GirlTimer.Enabled = 0
End Sub
Sub StopGirlTimer_Timer
  StopGirl
Me.Enabled = 0
End Sub
'*****************
'Flasher GlassBrok
'*****************
Dim GlassPos, FramesGlass
FramesGlass = Array("Glass001", "Glass002", "Glass003", "Glass004", "Glass005", "Glass006", "Glass007", "Glass008", "Glass009", "Glass010", "Glass011")

Sub StartGlass
  PlaySound "GlassB"
  GlassBrok.visible = 1
  GlassPos = 0
    GlassTimer.Enabled = 1
  StopGlassTimer.Enabled = 1
End Sub
Sub GlassTimer_Timer
    GlassBrok.ImageA = FramesGlass (GlassPos)
    GlassPos = (GlassPos + 1) MOD 11
End Sub
Sub StopGlass
  GlassBrok.visible = 0
  Flasher008.visible = 1
  Flasher007.visible = 1
  GlassTimer.Enabled = 0
End Sub
Sub StopGlassTimer_Timer
  StopGlass
Me.Enabled = 0
End Sub
'********************
' Shake Gun 360
'********************
Sub GunL1Timer_timer()
GunL.RotY = GunL.RotY - 52
if GunL.RotY <= -520 then GunL1Timer.enabled = 0
End Sub
Sub GunR1Timer_timer()
GunR.RotY = GunR.RotY + 52
if GunR.RotY >= 520 then GunR1Timer.enabled = 0
End Sub

Sub ShakeGunL1
  If GunL1Timer.enabled = 0 Then
  StartCanon6
  GunL.RotY = -160
  GunL1Timer.enabled = 1
  End If
End Sub
Sub ShakeGunR1
  If GunR1Timer.enabled = 0 Then
  StartCanon7
  GunR.RotY = 160
  GunR1Timer.enabled = 1
  End If
End Sub
'********************
' Shake Gun
'********************
Dim GunLPos, GunRPos

Sub ShakeLeftGun
  DOF 229, DOFPulse
    GunLPos = 8
    GunLTimer.Enabled = 1
End Sub

Sub GunLTimer_Timer
    GunL.TransY = GunLPos
    If GunLPos = 0 Then Me.Enabled = 0:Exit Sub
    If GunLPos < 0 Then
        GunLPos = ABS(GunLPos) - 1
    Else
        GunLPos = - GunLPos + 1
    End If
End Sub

Sub ShakeRightGun
  DOF 229, DOFPulse
    GunRPos = 8
    GunRTimer.Enabled = 1
End Sub

Sub GunRTimer_Timer
    GunR.TransY = GunRPos
    If GunRPos = 0 Then Me.Enabled = 0:Exit Sub
    If GunRPos < 0 Then
        GunRPos = ABS(GunRPos) - 1
    Else
        GunRPos = - GunRPos + 1
    End If
End Sub
'***********
'Open JAIL
'************
Sub ResetLightJail
  li0TD005.State = 2
  li0TD004.State = 2
  li0TD003.State = 2
  li0TD002.State = 2
  li0TD001.State = 2
End Sub
Sub CheckLightJAIL
  If li0TD005.State = 1 And li0TD004.State = 1 And li0TD003.State = 1 And li0TD002.State = 1 And li0TD001.State = 1 Then
  ShakeJailOFF
  AddScore 2000
    If Light001.state=1 And Light002.state=1 And Light003.state=0 Then
    liLOCK.State = 0
    liMB.State = 2
    Else
    liLOCK.State = 2
    DMDFlush
    DMD "   LOCK IS LIT","   GO TO JAIL","LockIsLit",eBlink, eNone, eNone, 4000, True, ""
    liMB.State = 0
    End If
 '   LightEffect 7
  End If
End Sub

Sub Target004_Hit
    PlaySoundAt "fx_droptarget", Target004
    If Tilted Then Exit Sub
    AddScore 500
    li0TD005.State = 1
  CheckLightJAIL
End Sub
Sub Target005_Hit
    PlaySoundAt "fx_droptarget", Target005
    If Tilted Then Exit Sub
    AddScore 500
    li0TD004.State = 1
  CheckLightJAIL
End Sub
Sub Target006_Hit
    PlaySoundAt "fx_droptarget", Target006
    If Tilted Then Exit Sub
    AddScore 500
    li0TD003.State = 1
  CheckLightJAIL
End Sub
Sub Target007_Hit
    PlaySoundAt "fx_droptarget", Target007
    If Tilted Then Exit Sub
    AddScore 500
    li0TD002.State = 1
  CheckLightJAIL
End Sub
Sub Target008_Hit
    PlaySoundAt "fx_droptarget", Target008
    If Tilted Then Exit Sub
    AddScore 500
    li0TD001.State = 1
  CheckLightJAIL
End Sub
'***********
' Spinners
'***********

Sub Spinner001_Spin
  DOF 235, DOFPulse
    PlaySoundAt "fx_spinner", Spinner001
    If Tilted Then Exit Sub
    Addscore 100
End Sub

Sub Spinner002_Spin
  DOF 236, DOFPulse
    PlaySoundAt "fx_spinner", Spinner002
    If Tilted Then Exit Sub
    Addscore 100
End Sub

'***********
' Light Events
'***********
Sub ResetLightEvent
  Dim xx: For each xx in aGiLights: xx.color = RGB(255,180,100): Next 'Standard
  Tilted = False
  DisableTable False
  FlasherBonus.Visible = 0
  FlasherBonusB.Visible = 0
  FlasherETG1A.Visible = 0
  FlasherETG1B.Visible = 0
  FlasherETG2A.Visible = 0
  FlasherETG2B.Visible = 0
  FlasherETG3.Visible = 0
  FlasherETG4A.Visible = 0
  FlasherETG4B.Visible = 0
  FlasherRDC1.Visible = 1
  FlasherRDC2A.Visible = 1
  FlasherRDC2B.Visible = 1
  li0Switch.state = 2
  li0Switch1.state = 1
  li0Switch2.state = 0
  Wall020.IsDropped = False
  FlasherOpenDoor2.visible = True
  Wall018.IsDropped = True
  Wall039.SideImage = "Bank"
  li0AS4.state = 0
  li0AS3.state = 0
  li0AS2.state = 0
  li0AS1.state = 0
  liSpinnerL.state = 1
  liSpinnerR.state = 1
  liRampR.state = 1
  liRamp2M1.state = 1
  liRampM.state = 1
  liRampL.state = 1
  ResetLightJail
  ResetLightHUNDERTAKER
  ResetLightSALOON
  ResetLightHOT
  ResetLightBANK
  ResetLightMINE
  Mission1_RODEO(CurrentPlayer) = 0
  Mission2_DUEL(CurrentPlayer) = 0
  Mission3_DYNAMITE(CurrentPlayer) = 0
  Mission4_BANK(CurrentPlayer) = 0
  Mission5_MINE(CurrentPlayer) = 0
  Mission6_POKER(CurrentPlayer) = 0
  GirlLOOK 0
  ShakeJailON
  ShakeChasseurOFF
  Target011.IsDropped = True
  Target012.IsDropped = True
  Target013.IsDropped = True
  Target014.IsDropped = True
  liGIRL001.state = 0 'Coeur GIRL
  liGIRL002.state = 0
  liGIRL003.state = 0
  Coeur_Sav(CurrentPlayer) = 0
  MBLock_Sav(CurrentPlayer) = 0
  Star_Sav(CurrentPlayer) = 0
  Saloon_Sav(CurrentPlayer) = 0
  Hunder_Sav(CurrentPlayer) = 0
  Mine_Sav(CurrentPlayer) = 0
  Bank_Sav(CurrentPlayer) = 0
  Star_Sav(CurrentPlayer) = 0
  Dyna_Sav(CurrentPlayer) = 0
  liMISS001.state = 1
  liMISS002.state = 1
  liMISS003.state = 1
  liMISS004.state = 1
  liMISS005.state = 1
  liMISS006.state = 1
  li0SHERIFF.state = 0
  FlasherDyna001.Visible = 0
  FlasherDyna002.Visible = 0
  FlasherDyna003.Visible = 0
  FlasherDyna004.Visible = 0
  FlasherDyna005.Visible = 0
  FlasherMISS001.Visible = 0
  FlasherMISS002.Visible = 0
  FlasherMISS003.Visible = 0
  FlasherMISS004.Visible = 0
  FlasherMISS005.Visible = 0
  FlasherMISS006.Visible = 0
  Soutif.Visible = 1
  Soutif2.Visible = 1
  ModeBankActive(CurrentPlayer) = 0
  ModeDuelActive(CurrentPlayer) = 0
  ModeDynamiteActive(CurrentPlayer) = 0
  ModePokerActive(CurrentPlayer) = 0
  ModeRodeoActive(CurrentPlayer) = 0
  ModeMultiballActive = False
  ModeSpinnersActive(CurrentPlayer) = True
  ModeMineActive(CurrentPlayer) = 0
  ModeMultiballActive = False
  li0START.state = 0
  SpinnerCount = 25
  li0TD001.state = 2 'MB
  li0TD002.state = 2
  li0TD003.state = 2
  li0TD004.state = 2
  li0TD005.state = 2
  Light001.state=0 'MB
  Light002.state=0
  Light003.state=0
  liLOCK.state=0
  liMB.state=0
' li020.state=1 'Outline save
' li021.state=1
  ActiveOutlineSave
End Sub

Sub ActiveOutlineSave
  If ModeDifficulty = 0 Then
    li020.state=1 'Outline save
    li021.state=1
  Elseif ModeDifficulty = 1 Then
    li020.state=2 'Outline save
    li021.state=2
  End If
End Sub

Sub ResetLightEvent2 'ModeTriche
  Dim xx: For each xx in aGiLights: xx.color = RGB(255,180,100): Next 'Standard
  Tilted = False
  DisableTable False
  FlasherBonus.Visible = 0
  FlasherBonusB.Visible = 0
  FlasherETG1A.Visible = 0
  FlasherETG1B.Visible = 0
  FlasherETG2A.Visible = 0
  FlasherETG2B.Visible = 0
  FlasherETG3.Visible = 0
  FlasherETG4A.Visible = 0
  FlasherETG4B.Visible = 0
  FlasherRDC1.Visible = 1
  FlasherRDC2A.Visible = 1
  FlasherRDC2B.Visible = 1
  li0Switch.state = 2
  li0Switch1.state = 1
  li0Switch2.state = 0
  Wall020.IsDropped = False
  FlasherOpenDoor2.visible = True
  Wall018.IsDropped = True
  Wall039.SideImage = "Bank"
  li0AS4.state = 0
  li0AS3.state = 0
  li0AS2.state = 0
  li0AS1.state = 0
  liSpinnerL.state = 1
  liSpinnerR.state = 1
  liRampR.state = 1
  liRamp2M1.state = 1
  liRampM.state = 1
  liRampL.state = 1
' ResetLightHUNDERTAKER
  li0R3M1.state = 0
  li0R3M2.state = 1
  li0R3M3.state = 1
  li0R3M4.state = 1
  li0R3M5.state = 1
  li0R3M6.state = 1
  li0R3M7.state = 1
  li0R3M8.state = 1
  li0R3M9.state = 1
  li0R3M10.state = 1
' ResetLightSALOON
  li0R2M6.state = 1
  li0R2M5.state = 1
  li0R2M4.state = 1
  li0R2M3.state = 1
  li0R2M2.state = 1
  li0R2M1.state = 0
  ResetLightHOT
  liGIRL001.state = 1 'Coeur GIRL
  liGIRL002.state = 0
  liGIRL003.state = 0
  Coeur_Sav(CurrentPlayer) = 1
  MBLock_Sav(CurrentPlayer) = 0
  Star_Sav(CurrentPlayer) = 1
  Saloon_Sav(CurrentPlayer) = 5
  Hunder_Sav(CurrentPlayer) = 6
  Mine_Sav(CurrentPlayer) = 3
  Bank_Sav(CurrentPlayer) = 3
  Star_Sav(CurrentPlayer) = 1
  Dyna_Sav(CurrentPlayer) = 0
  GirlLOOK 2
' ResetLightBANK
  li0RM1.state = 1
  li0RM2.state = 1
  li0RM3.state = 1
  li0RM4.state = 0
' ResetLightMINE
  li0RL1.state = 1
  li0RL2.state = 1
  li0RL3.state = 1
  li0RL4.state = 0
  Mission1_RODEO(CurrentPlayer) = 0
  Mission2_DUEL(CurrentPlayer) = 0
  Mission3_DYNAMITE(CurrentPlayer) = 0
  Mission4_BANK(CurrentPlayer) = 0
  Mission5_MINE(CurrentPlayer) = 0
  Mission6_POKER(CurrentPlayer) = 0
  ShakeJailOFF
  ShakeChasseurON
' ResetLightJail
  Target011.IsDropped = True
  Target012.IsDropped = True
  Target013.IsDropped = True
  Target014.IsDropped = True
  liMISS001.state = 1
  liMISS002.state = 1
  liMISS003.state = 1
  liMISS004.state = 1
  liMISS005.state = 1
  liMISS006.state = 1
  li0SHERIFF.state = 2
  FlasherDyna001.Visible = 0
  FlasherDyna002.Visible = 0
  FlasherDyna003.Visible = 0
  FlasherDyna004.Visible = 0
  FlasherDyna005.Visible = 0
  FlasherMISS001.Visible = 0
  FlasherMISS002.Visible = 0
  FlasherMISS003.Visible = 0
  FlasherMISS004.Visible = 0
  FlasherMISS005.Visible = 0
  FlasherMISS006.Visible = 0
  ModeBankActive(CurrentPlayer) = 0
  ModeDuelActive(CurrentPlayer) = 0
  ModeDynamiteActive(CurrentPlayer) = 0
  ModePokerActive(CurrentPlayer) = 0
  ModeRodeoActive(CurrentPlayer) = 0
  ModeMultiballActive = False
  ModeSpinnersActive(CurrentPlayer) = True
  ModeMineActive(CurrentPlayer) = 0
  ModeMultiballActive = False
  li0START.state = 0
  SpinnerCount = 15
  li0TD001.state = 1 'MB
  li0TD002.state = 1
  li0TD003.state = 1
  li0TD004.state = 1
  li0TD005.state = 1
  liLOCK.state = 2
  liMB.state = 0
  Light001.state=0 'MB
  Light002.state=0
  Light003.state=0
' li020.state=1 'Outline save
' li021.state=1
  ActiveOutlineSave
End Sub

Sub CheckMissionActive 'players mission ET lights
  ResetABC
  ResetMystery
  CheckLightMINE
  CheckLightSALOON
  If ModePokerActive(CurrentPlayer) = 2 Then StartMissionPOKER : Li0AS2.state = 1 : li0AS4.state = 1
  If ModePokerActive(CurrentPlayer) = 1 Then li0START.State = 2
  If ModePokerActive(CurrentPlayer) = 0 Then li0AS1.state = 0 : li0AS2.state = 0 : li0AS3.state = 0 : li0AS4.state = 0 : liSpinnerL.state = 1 : liSpinnerR.state = 1
  CheckLightHUNDERTAKER
  CheckLightBANK
  CheckBankState
  If Mission5_MINE(CurrentPlayer) = 1 Then FlasherMISS005.Visible = 1 : liMISS005.state = 0
  If Mission5_MINE(CurrentPlayer) = 0 Then FlasherMISS005.Visible = 0
  CheckLightPoker
  If Mission6_POKER(CurrentPlayer) = 1 Then FlasherMISS006.Visible = 1 : liMISS006.state = 0
  If Mission6_POKER(CurrentPlayer) = 0 Then FlasherMISS006.Visible = 0
  If Mission3_DYNAMITE(CurrentPlayer) = 1 Then FlasherMISS003.Visible = 1 : liMISS003.state = 0
  If Mission3_DYNAMITE(CurrentPlayer) = 0 Then
    FlasherMISS003.Visible = 0
    liMISS003.state = 1
    If ModeDynamiteActive(CurrentPlayer) = 0 And ModeCheatActive = True Then ModeSpinnersActive(CurrentPlayer) = True : SpinnerCount = 15 : FlasherDyna001.visible = 0 : FlasherDyna002.visible = 0 : FlasherDyna003.visible = 0 : FlasherDyna004.visible = 0
    If ModeDynamiteActive(CurrentPlayer) = 0 And ModeCheatActive = False Then ModeSpinnersActive(CurrentPlayer) = True : SpinnerCount = 25 : FlasherDyna001.visible = 0 : FlasherDyna002.visible = 0 : FlasherDyna003.visible = 0 : FlasherDyna004.visible = 0
    If ModeDynamiteActive(CurrentPlayer) = 1 Then li0START.State = 2 : FlasherDyna001.visible = 0 : FlasherDyna002.visible = 0 : FlasherDyna003.visible = 0 : FlasherDyna004.visible = 0
    If ModeDynamiteActive(CurrentPlayer) = 2 Then StartMissionDYNAMITE
  End If
  If Mission2_DUEL(CurrentPlayer) = 1 Then FlasherMISS002.Visible = 1 : liMISS002.state = 0
  If Mission2_DUEL(CurrentPlayer) = 0 Then
    FlasherMISS002.Visible = 0
    liMISS002.state = 1
    If ModeDuelActive(CurrentPlayer) = 2 Then StartSHERIFFtarget
    If ModeDuelActive(CurrentPlayer) = 1 Then li0START.State = 2 : Target011.IsDropped = True : Target012.IsDropped = True : Target013.IsDropped = True : Target014.IsDropped = True
    If ModeDuelActive(CurrentPlayer) = 0 Then Target011.IsDropped = True : Target012.IsDropped = True : Target013.IsDropped = True : Target014.IsDropped = True
  End If
  CheckB2S
  CheckStar_Sav
  CheckMBLock_Save
  CheckCoeur_Sav
  If ModeRodeoActive(CurrentPlayer) = 2 Then StartMissionRODEO
  If ModeDynamiteActive(CurrentPlayer) = 1 or ModePokerActive(CurrentPlayer) = 1 or ModeMineActive(CurrentPlayer) = 1 or ModeDuelActive(CurrentPlayer) = 1 or ModeRodeoActive(CurrentPlayer) = 1 or ModeSpinnersActive(CurrentPlayer) = 1 Then li0START.state = 2 Else li0START.state = 0
End Sub

Sub CheckCoeur_Sav
  If Mission1_RODEO(CurrentPlayer) = 1 Then
    If ModeNudeActive = 1 Then GirlLOOK 4
    If ModeNudeActive = 0 Then GirlLOOK 3
    liGIRL001.state = 1
    liGIRL002.state = 1
    liGIRL003.state = 1
    FlasherMISS001.Visible = 1
    liMISS001.state = 0
  Elseif Mission1_RODEO(CurrentPlayer) = 0 Then
    FlasherMISS001.Visible = 0
    liMISS001.state = 1
    If Coeur_Sav(CurrentPlayer) = 0 Then GirlLOOK 0 : liGIRL001.state = 0 : liGIRL002.state = 0 : liGIRL003.state = 0
    If Coeur_Sav(CurrentPlayer) = 1 Then GirlLOOK 2 : liGIRL001.state = 1 : liGIRL002.state = 0 : liGIRL003.state = 0
    If Coeur_Sav(CurrentPlayer) = 2 Then GirlLOOK 1 : liGIRL001.state = 1 : liGIRL002.state = 1 : liGIRL003.state = 0
    If Coeur_Sav(CurrentPlayer) = 3 Then GirlLOOK 3 : liGIRL001.state = 1 : liGIRL002.state = 1 : liGIRL003.state = 1
    If Coeur_Sav(CurrentPlayer) >= 4 And ModeNudeActive = 0 Then GirlLOOK 3 : liGIRL001.state = 1 : liGIRL002.state = 1 : liGIRL003.state = 1
    If Coeur_Sav(CurrentPlayer) >= 4 And ModeNudeActive = 1 Then GirlLOOK 4 : liGIRL001.state = 1 : liGIRL002.state = 1 : liGIRL003.state = 1
    If ModeRodeoActive(CurrentPlayer) = 1 Then li0START.State = 2
  End If
End Sub

Sub CheckMBLock_Save
  If MBLock_Sav(CurrentPlayer) = 0 Then
  Light001.state=0
  Light002.state=0
  Light003.state=0
  Elseif MBLock_Sav(CurrentPlayer) = 1 Then
  Light001.state=1
  Light002.state=0
  Light003.state=0
  Elseif MBLock_Sav(CurrentPlayer) = 2 Then
  Light001.state=1
  Light002.state=1
  Light003.state=0
  End If
End Sub
'-------- MISSION TARGET STAR

Sub CheckStar_Sav
  If Star_Sav(CurrentPlayer) = 1 And ModeCheatActive = True Then
  li0SHERIFF.state = 2
  Elseif Star_Sav(CurrentPlayer) = 0 And ModeCheatActive = False Then
  li0SHERIFF.state = 1
  Elseif Star_Sav(CurrentPlayer) = 1 And ModeCheatActive = False Then
  li0SHERIFF.state = 1
  Elseif Star_Sav(CurrentPlayer) = 1 And ModeCheatActive = True Then
  li0SHERIFF.state = 1
  End If
End Sub

Sub StartSHERIFFtarget
  DOF 237, DOFPulse    'GET EM
  D3TargetsMilieu = 4
' pupevent 106
  DMD " SHOOT TARGETS","TRUANDS  WANTED","ModDuel",eBlink, eNone, eNone, 5000, True, ""
  liMISS002.State = 2
  ModeDuelActive(CurrentPlayer) = 2
  Target011.IsDropped = False
  Target012.IsDropped = False
  Target013.IsDropped = False
  Target014.IsDropped = False
End Sub

Sub StopSHERIFFtarget
  AddScore 45000
    DOF 238, DOFPulse    'OUR HERO
  DMDFlush
  DMD "    THE GUYS","  ARE  ARRESTED","ModDuel",eNone, eNone, eNone, 5000, True, ""
  ModeDuelActive(CurrentPlayer) = 0
  Mission2_DUEL(CurrentPlayer) = 1
  Star_Sav(CurrentPlayer) = 0
  li0SHERIFF.State = 1
  liMISS002.State = 0
  FlasherMISS002.visible = 1
  Target011.IsDropped = True
  Target012.IsDropped = True
  Target013.IsDropped = True
  Target014.IsDropped = True
  StartCanon5
  CheckB2S
End Sub
Sub CheckTargetsMode ()
' D3TargetsMilieu = 4
  D3TargetsMilieu = D3TargetsMilieu - D3TargetsDown
  If D3TargetsDown => 4 Then
    StopSHERIFFtarget
  Else
'   DMDFlush
'   DMD "NUMBERS REMAINING", formatscore(D3TargetsMilieu) & " OF 4", "", eNone, eNone, eNone, 3000, True, ""
  End If
End Sub

Sub DMDkillT
  DMD "", "", "KillT001", eNone, eNone, eNone, 50, False, ""
  DMD "", "", "KillT002", eNone, eNone, eNone, 50, False, ""
  DMD "", "", "KillT003", eNone, eNone, eNone, 50, False, ""
  DMD "", "", "KillT004", eNone, eNone, eNone, 50, False, ""
  DMD "", "", "KillT005", eNone, eNone, eNone, 50, False, ""
  DMD "", "", "KillT006", eNone, eNone, eNone, 50, False, ""
  DMD "", "", "KillT007", eNone, eNone, eNone, 50, False, ""
  DMD "", "", "KillT008", eNone, eNone, eNone, 50, False, ""
  DMD "", "", "KillT009", eNone, eNone, eNone, 50, False, ""
  DMD "     NICE", "       SHOOT", "KillT010", eBlink, eNone, eNone, 300, False, ""
  DMD "     NICE", "       SHOOT", "KillT011", eBlinK, eNone, eNone, 2000, True, ""
End Sub

Sub Target011_Hit
  DOF 239, DOFPulse
  AddScore 5500
  D3TargetsDown = D3TargetsDown + 1
  CheckTargetsMode()
  DMDkillT
  PlaySoundGun
End Sub
Sub Target012_Hit
  DOF 239, DOFPulse
  AddScore 5500
  D3TargetsDown = D3TargetsDown + 1
  CheckTargetsMode()
  DMDkillT
  PlaySoundGun
End Sub
Sub Target013_Hit
  DOF 239, DOFPulse
  AddScore 5500
  D3TargetsDown = D3TargetsDown + 1
  CheckTargetsMode()
  DMDkillT
  PlaySoundGun
End Sub
Sub Target014_Hit
  DOF 239, DOFPulse
  AddScore 5500
  D3TargetsDown = D3TargetsDown + 1
  CheckTargetsMode()
  DMDkillT
  PlaySoundGun
End Sub

Sub ResetLightBall2345
  FlasherBonus.ImageA = "Bonus1"
  FlasherBonusB.ImageA = "Bonus1b"
End Sub
Sub ResetLightHUNDERTAKER
  li0R3M1.state = 0
  li0R3M2.state = 0
  li0R3M3.state = 0
  li0R3M4.state = 0
  li0R3M5.state = 0
  li0R3M6.state = 0
  li0R3M7.state = 0
  li0R3M8.state = 0
  li0R3M9.state = 0
  li0R3M10.state = 0
  li0OUTLL2.State = 0
  LightShootAgain.state=0
  Hunder_Sav(CurrentPlayer) = 0
  CheckB2S
End Sub

Sub CheckLightHUNDERTAKER
  If ModeCheatActive = True And Hunder_Sav(CurrentPlayer) <= 6 Then
    Hunder_Sav(CurrentPlayer) = 6
    li0R3M1.state = 0
    li0R3M2.state = 0
    li0R3M3.state = 0
    li0R3M4.state = 0
    li0R3M5.state = 1
    li0R3M6.state = 1
    li0R3M7.state = 1
    li0R3M8.state = 1
    li0R3M9.state = 1
    li0R3M10.state = 1
    li0OUTLL2.State = 0
    LightShootAgain.state=0
  Else
    If Hunder_Sav(CurrentPlayer) = 0 Then ResetLightHUNDERTAKER
    If Hunder_Sav(CurrentPlayer) = 1 Then li0R3M1.state = 0 : li0R3M2.state = 0 : li0R3M3.state = 0 : li0R3M4.state = 0 : li0R3M5.state = 0 : li0R3M6.state = 0 : li0R3M7.state = 0 : li0R3M8.state = 0 : li0R3M9.state = 0 : li0R3M10.state = 1 : li0OUTLL2.State = 0 : LightShootAgain.state=0
    If Hunder_Sav(CurrentPlayer) = 2 Then li0R3M1.state = 0 : li0R3M2.state = 0 : li0R3M3.state = 0 : li0R3M4.state = 0 : li0R3M5.state = 0 : li0R3M6.state = 0 : li0R3M7.state = 0 : li0R3M8.state = 0 : li0R3M9.state = 1 : li0R3M10.state = 1 : li0OUTLL2.State = 0 : LightShootAgain.state=0
    If Hunder_Sav(CurrentPlayer) = 3 Then li0R3M1.state = 0 : li0R3M2.state = 0 : li0R3M3.state = 0 : li0R3M4.state = 0 : li0R3M5.state = 0 : li0R3M6.state = 0 : li0R3M7.state = 0 : li0R3M8.state = 1 : li0R3M9.state = 1 : li0R3M10.state = 1 : li0OUTLL2.State = 0 : LightShootAgain.state=0
    If Hunder_Sav(CurrentPlayer) = 4 Then li0R3M1.state = 0 : li0R3M2.state = 0 : li0R3M3.state = 0 : li0R3M4.state = 0 : li0R3M5.state = 0 : li0R3M6.state = 0 : li0R3M7.state = 1 : li0R3M8.state = 1 : li0R3M9.state = 1 : li0R3M10.state = 1 : li0OUTLL2.State = 0 : LightShootAgain.state=0
    If Hunder_Sav(CurrentPlayer) = 5 Then li0R3M1.state = 0 : li0R3M2.state = 0 : li0R3M3.state = 0 : li0R3M4.state = 0 : li0R3M5.state = 0 : li0R3M6.state = 1 : li0R3M7.state = 1 : li0R3M8.state = 1 : li0R3M9.state = 1 : li0R3M10.state = 1 : li0OUTLL2.State = 0 : LightShootAgain.state=0
    If Hunder_Sav(CurrentPlayer) = 6 Then li0R3M1.state = 0 : li0R3M2.state = 0 : li0R3M3.state = 0 : li0R3M4.state = 0 : li0R3M5.state = 1 : li0R3M6.state = 1 : li0R3M7.state = 1 : li0R3M8.state = 1 : li0R3M9.state = 1 : li0R3M10.state = 1 : li0OUTLL2.State = 0 : LightShootAgain.state=0
    If Hunder_Sav(CurrentPlayer) = 7 Then li0R3M1.state = 0 : li0R3M2.state = 0 : li0R3M3.state = 0 : li0R3M4.state = 1 : li0R3M5.state = 1 : li0R3M6.state = 1 : li0R3M7.state = 1 : li0R3M8.state = 1 : li0R3M9.state = 1 : li0R3M10.state = 1 : li0OUTLL2.State = 0 : LightShootAgain.state=0
    If Hunder_Sav(CurrentPlayer) = 8 Then li0R3M1.state = 0 : li0R3M2.state = 0 : li0R3M3.state = 1 : li0R3M4.state = 1 : li0R3M5.state = 1 : li0R3M6.state = 1 : li0R3M7.state = 1 : li0R3M8.state = 1 : li0R3M9.state = 1 : li0R3M10.state = 1 : li0OUTLL2.State = 0 : LightShootAgain.state=0
    If Hunder_Sav(CurrentPlayer) = 9 Then li0R3M1.state = 0 : li0R3M2.state = 1 : li0R3M3.state = 1 : li0R3M4.state = 1 : li0R3M5.state = 1 : li0R3M6.state = 1 : li0R3M7.state = 1 : li0R3M8.state = 1 : li0R3M9.state = 1 : li0R3M10.state = 1 : li0OUTLL2.State = 0 : LightShootAgain.state=0
    If Hunder_Sav(CurrentPlayer) >= 10 Then li0R3M1.state = 2 : li0R3M2.state = 2 : li0R3M3.state = 2 : li0R3M4.state = 2 : li0R3M5.state = 2 : li0R3M6.state = 2 : li0R3M7.state = 2 : li0R3M8.state = 2 : li0R3M9.state = 2 : li0R3M10.state = 2 : li0OUTLL2.State = 2 : LightShootAgain.state=2
  End If
End Sub

Sub CheckAwardHUNDERTAKER
  If Hunder_Sav(CurrentPlayer) <= 9 Then
'   pupevent 102
    AddScore 2000
    Hunder_Sav(CurrentPlayer) = Hunder_Sav(CurrentPlayer) + 1
  Elseif Hunder_Sav(CurrentPlayer) >= 10 Then
    AddScore 30000
    LightEffect 7
    AwardExtraBall
'   pupevent 111
    ResetLightHUNDERTAKER
  End If
  CheckLightHUNDERTAKER
End Sub

Sub Trigger012_Hit
  CheckAwardHUNDERTAKER
  DMD "", "", "Under", eNone, eNone, eNone, 3500, True, ""
  PlaySound "fx_PasHorse"
End Sub

Sub ResetLightSALOON
  li0R2M1.state = 0
  li0R2M2.state = 0
  li0R2M3.state = 0
  li0R2M4.state = 0
  li0R2M5.state = 0
  li0R2M6.state = 0
  Saloon_Sav(CurrentPlayer) = 0
End Sub

Sub CheckLightSALOON
  If ModeCheatActive = True And Saloon_Sav(CurrentPlayer) <= 5 And Mission6_POKER(CurrentPlayer) = 0 Then
    Saloon_Sav(CurrentPlayer) = 5
    li0R2M1.state = 0
    li0R2M2.state = 1
    li0R2M3.state = 1
    li0R2M4.state = 1
    li0R2M5.state = 1
    li0R2M6.state = 1
  Else
    If Saloon_Sav(CurrentPlayer) = 0 Then ResetLightSALOON
    If Saloon_Sav(CurrentPlayer) = 1 Then li0R2M1.State = 0 : li0R2M2.State = 0 : li0R2M3.State = 0 : li0R2M4.State = 0 : li0R2M5.State = 0 : li0R2M6.State = 1
    If Saloon_Sav(CurrentPlayer) = 2 Then li0R2M1.State = 0 : li0R2M2.State = 0 : li0R2M3.State = 0 : li0R2M4.State = 0 : li0R2M5.State = 1 : li0R2M6.State = 1
    If Saloon_Sav(CurrentPlayer) = 3 Then li0R2M1.State = 0 : li0R2M2.State = 0 : li0R2M3.State = 0 : li0R2M4.State = 1 : li0R2M5.State = 1 : li0R2M6.State = 1
    If Saloon_Sav(CurrentPlayer) = 4 Then li0R2M1.State = 0 : li0R2M2.State = 0 : li0R2M3.State = 1 : li0R2M4.State = 1 : li0R2M5.State = 1 : li0R2M6.State = 1
    If Saloon_Sav(CurrentPlayer) = 5 Then li0R2M1.State = 0 : li0R2M2.State = 1 : li0R2M3.State = 1 : li0R2M4.State = 1 : li0R2M5.State = 1 : li0R2M6.State = 1
    If Saloon_Sav(CurrentPlayer) = 6 Then li0R2M1.State = 2 : li0R2M2.State = 2 : li0R2M3.State = 2 : li0R2M4.State = 2 : li0R2M5.State = 2 : li0R2M6.State = 2
    If Saloon_Sav(CurrentPlayer) >= 7 Then li0R2M1.State = 1 : li0R2M2.State = 1 : li0R2M3.State = 1 : li0R2M4.State = 1 : li0R2M5.State = 1 : li0R2M6.State = 1
  End If
End Sub

Sub CheckAwardSaloon
  AddScore 2000
  If Saloon_Sav(CurrentPlayer) <= 5 Then Saloon_Sav(CurrentPlayer) = Saloon_Sav(CurrentPlayer) + 1
  If Saloon_Sav(CurrentPlayer) = 6 Then
    AddScore 10000
    If Mission6_POKER(CurrentPlayer) = 1 Then
      PlaySound"Vo_Combo"
      DOF 128, DOFPulse
    ' pupevent 112
      dmdflush
      DMD "", "", "Combo", eNone, eNone, eNone, 3500, True, ""
    Elseif Mission6_POKER(CurrentPlayer) = 0 Then
      ModePokerActive(CurrentPlayer) = 1
      li0START.state=2
    End If
    LightEffect 7
    ResetLightSALOON
  End If
  CheckLightSALOON
End Sub

Sub StartMissionPOKER
  ModePokerActive(CurrentPlayer) = 2
  Saloon_Sav(CurrentPlayer) = 7
' pupevent 107
  DMD "SHOOT SPINNERS","COLLECT 4 AS","ModPoker",eBlink, eNone, eNone, 5000, True, ""
  ChangeBall(1)
  li0AS1.state = 2
  li0AS2.state = 2
  li0AS3.state = 2
  li0AS4.state = 2
  liSpinnerL.state = 2
  liSpinnerR.state = 2
  liMISS006.state = 2
  li0R2M1.state = 1
  li0R2M2.state = 1
  li0R2M3.state = 1
  li0R2M4.state = 1
  li0R2M5.state = 1
  li0R2M6.state = 1
End Sub

Sub StopMissionPOKER
  ModePokerActive(CurrentPlayer) = 0
  Mission6_POKER(CurrentPlayer) = 1
  DMD "   WELL DONE","YOU WIN POKER","ModPoker",eBlink, eNone, eNone, 5000, True, ""
  li0AS1.state = 0
  li0AS2.state = 0
  li0AS3.state = 0
  li0AS4.state = 0
  ChangeBall(0)
  liSpinnerL.state = 1
  liSpinnerR.state = 1
  liMISS006.state = 0
  FlasherMISS006.visible = 1
  StartCanon5
  ResetLightSALOON
  CheckB2S
End Sub

Sub Trigger013_Hit
  CheckLightPoker
  If li0AS3.state = 2 And li0AS4.state = 2 Then
  li0AS4.state = 1
  dmdflush
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerP", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerP", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerP", eNone, eNone, eNone, 3500, True, ""
  AddScore 3000
  Elseif li0AS3.state = 2 And li0AS4.state = 1 Then
  li0AS3.state = 1
  dmdflush
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerCO", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerCO", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerCO", eNone, eNone, eNone, 3500, True, ""
  AddScore 3000
  End IF
End Sub
Sub Trigger014_Hit
  CheckLightPoker
  If li0AS2.state = 2 And li0AS1.state = 2 Then
  li0AS2.state = 1
  dmdflush
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerT", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerT", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerT", eNone, eNone, eNone, 3500, True, ""
  AddScore 3000
  Elseif li0AS2.state = 1 And li0AS1.state = 2 Then
  li0AS1.state = 1
  dmdflush
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerCA", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerCA", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPoker", eNone, eNone, eNone, 300, False, ""
  DMD " ADD AS", "", "ModPokerCA", eNone, eNone, eNone, 3500, True, ""
  AddScore 3000
  End If
End Sub
Sub CheckLightPoker
  If li0AS1.state = 1 And li0AS2.state = 1 And li0AS3.state = 1 And li0AS4.state = 1 Then
  StopMissionPOKER
  End If
End Sub


Sub Trigger007_Hit
  CheckAwardSaloon
' pupevent 101
  If ModeNudeActive = 1 Then
  DMD "", "", "saloonL", eNone, eNone, eNone, 3500, True, ""
  Elseif ModeNudeActive = 0 Then
  DMD "", "", "saloonL", eNone, eNone, eNone, 3500, True, ""
  End If
  PlaySound "fx_PasHorse"
End Sub

Sub ResetLightHOT
  li0RR1.state = 0
  li0RR2.state = 0
  li0RR3.state = 0
  li0RR4.State = 1
End Sub

Sub CheckLightHOT 'COUNT GIRL
  If ModeRodeoActive(CurrentPlayer) <1 or ModeRodeoActive(CurrentPlayer) >=2 Then
    If li0RR1.State = 0 And li0RR2.State = 0 And li0RR3.State = 0 Then
    li0RR1.State = 1
    PlaySoundRire
    Elseif li0RR1.State = 1 And li0RR2.State = 0 And li0RR3.State = 0 Then
    li0RR2.State = 1
    PlaySoundRire
    Elseif li0RR1.State = 1 And li0RR2.State = 1 And li0RR3.State = 0 Then
    li0RR3.State = 1
    li0RR4.State = 2
    PlaySoundRire
    Elseif li0RR4.State = 2 Then
    AddScore 2000
    LightEffect 7
    If Mission1_RODEO(CurrentPlayer) = 0 And ModeRodeoActive(CurrentPlayer) <1 Then Coeur_Sav(CurrentPlayer) = Coeur_Sav(CurrentPlayer) + 1
    CheckAwardGIRL
    ResetLightHOT
    End If
  Else
    AddScore 3150
    PlaySoundRire
  End If
End Sub

Sub CheckAwardGIRL 'AWARD
  If Mission1_RODEO(CurrentPlayer) = 0 Then
    If Coeur_Sav(CurrentPlayer) = 1 Then
      GirlLOOK 2
      CheckCoeur_Sav
      DMDShowGirl
      AddMBmission
      PlaySound "Vo_MB"
    Elseif Coeur_Sav(CurrentPlayer) = 2 Then
      GirlLOOK 0
      AddMBmission
      StartGirl 'ajout CheckCoeur_Sav
      DMDShowGirl
    Elseif Coeur_Sav(CurrentPlayer) = 3 Then
      GirlLOOK 3
      DMDShowGirl
      AddMBmission
      PlaySound "Vo_MB"
      ModeRodeoActive(CurrentPlayer) = 1
      li0START.State = 2
      CheckCoeur_Sav
    End If
    If ModeRodeoActive(CurrentPlayer) = 2 Then
      AddMBmission
      GirlLOOK 4
      StopMissionRODEO
      CheckCoeur_Sav
    End If
  End If
End Sub

Sub DMDShowGirl
  If ModeNudeActive = 1 Then
  DMD "","","ModeHot_001",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","ModeHot_002",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","ModeHot_003",eNone, eNone, eNone, 120, FALSE, ""
  DMD "","","ModeHot_004",eNone, eNone, eNone, 1000, FALSE, ""
  DMD "","","ModeHot_005",eNone, eNone, eNone, 150, FALSE, ""
  DMD "","","ModeHot_006",eNone, eNone, eNone, 150, FALSE, ""
  DMD "","","ModeHot_007",eNone, eNone, eNone, 150, FALSE, ""
  DMD "","","ModeHot_008",eNone, eNone, eNone, 160, FALSE, ""
  DMD "","","ModeHot_009",eNone, eNone, eNone, 160, FALSE, ""
  DMD "","","ModeHot_010",eNone, eNone, eNone, 160, FALSE, ""
  DMD "","","ModeHot_011",eNone, eNone, eNone, 160, FALSE, ""
  DMD "","","ModeHot_012",eNone, eNone, eNone, 160, FALSE, ""
  DMD "","","ModeHot_013",eNone, eNone, eNone, 160, FALSE, ""
  DMD "","","ModeHot_014",eNone, eNone, eNone, 160, FALSE, ""
  DMD "YOU SEDUCE","","ModeHot_015",eNone, eNone, eNone, 800, FALSE, ""
  DMD "YOU SEDUCE","","ModeHot_016",eNone, eNone, eNone, 160, FALSE, ""
  DMD "YOU SEDUCE","","ModeHot_017",eNone, eNone, eNone, 160, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_018",eBlink, eNone, eNone, 120, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_019",eBlink, eNone, eNone, 120, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_020",eBlink, eNone, eNone, 120, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_021",eBlink, eNone, eNone, 120, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_022",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_023",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_024",eBlink, eNone, eNone, 150, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_025",eBlink, eNone, eNone, 160, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_026",eBlink, eNone, eNone, 160, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_027",eBlink, eNone, eNone, 160, FALSE, ""
  DMD "YOU SEDUCE","    THE LADY","ModeHot_028",eNone, eNone, eNone, 3000, TRUE, ""
  Elseif ModeNudeActive = 0 Then
  DMD "YOU SEDUCE","    THE LADY","ModRodeoOK",eNone, eNone, eNone, 4000, TRUE, ""
  End If
End Sub

'-----MISSION 4 BANK
Sub ResetLightBANK
  li0RM1.state = 0
  li0RM2.state = 0
  li0RM3.state = 0
  li0RM4.state = 0
  Bank_Sav(CurrentPlayer) = 0
End Sub

Sub StartMissionBank
  DMD "  SHOOT RAMP BANK","  EMPTY THE TRUNK","ModBank",eBlink, eNone, eNone, 5000, True, ""
  Bank_Sav(CurrentPlayer) = 5
  ModeBankActive(CurrentPlayer) = 2
  FlasherDyna005.visible = 1
  spinner3.MotorOn = True
  spinning3.enabled = True
  liRampM.state=2
  liMISS004.state = 2
  li0RM1.state = 2
  li0RM2.state = 2
  li0RM3.state = 2
  li0RM4.state = 2
' li020.state=1
' li021.state=1
  ActiveOutlineSave
End Sub

Sub AwardBANK1
  AddScore 5000
  StartCanon4
  PlaySound "fx_Booum"
  FlasherDyna005.visible = 0
  Bank_Sav(CurrentPlayer) = 6
  ModeBankActive(CurrentPlayer) = 3
  CheckBankState
End Sub

Sub AwardBANK2
' pupevent 113
  AddScore 85000
  ModeBankActive(CurrentPlayer) = 4
  Mission4_BANK(CurrentPlayer) = 1
  DMD "        HOLD UP","       SUCCESSFUL","ModBank",eBlink, eNone, eNone, 5000, True, ""
  ResetLightBANK
  CheckBankState
  StartCanon5
  CheckB2S
End Sub

Sub CheckLightBANK
  If ModeCheatActive = True And Bank_Sav(CurrentPlayer) <= 3 And Mission4_BANK(CurrentPlayer) = 0 Then
    Bank_Sav(CurrentPlayer) = 3
    li0RM1.state = 1
    li0RM2.state = 1
    li0RM3.state = 1
    li0RM4.state = 0
  Else
    If Bank_Sav(CurrentPlayer) = 0 Then ResetLightBANK
    If Bank_Sav(CurrentPlayer) = 1 Then li0RM1.State = 1 : li0RM2.State = 0 : li0RM3.State = 0 : li0RM4.State = 0
    If Bank_Sav(CurrentPlayer) = 2 Then li0RM1.State = 1 : li0RM2.State = 1 : li0RM3.State = 0 : li0RM4.State = 0
    If Bank_Sav(CurrentPlayer) = 3 Then li0RM1.State = 1 : li0RM2.State = 1 : li0RM3.State = 1 : li0RM4.State = 0
    If Bank_Sav(CurrentPlayer) = 4 And Mission4_BANK(CurrentPlayer) = 0 Then ModeBankActive(CurrentPlayer) = 1
    If Bank_Sav(CurrentPlayer) = 4 And Mission4_BANK(CurrentPlayer) = 1 Then ResetLightMINE
    If Bank_Sav(CurrentPlayer) >= 5 Then li0RM1.State = 2 : li0RM2.State = 2 : li0RM3.State = 2 : li0RM4.State = 2 : liRampM.state=2
  End If
  CheckBankState
End Sub

Sub CheckBankState
  If Mission4_BANK(CurrentPlayer) = 0 Then
    FlasherMISS004.visible = 0
    If ModeBankActive(CurrentPlayer) = 0 Then Wall039.SideImage = "Bank" : liMISS004.state = 1 : liRampM.state=1 :  : FlasherDyna005.visible = 0
    If ModeBankActive(CurrentPlayer) = 1 Then Wall039.SideImage = "Bank" : li0START.state = 2 : liMISS004.state = 1 : liRampM.state=1 : FlasherDyna005.visible = 0
    If ModeBankActive(CurrentPlayer) = 2 Then Wall039.SideImage = "Bank" : liMISS004.state = 2 : FlasherDyna005.visible = 1':StartMissionBank
    If ModeBankActive(CurrentPlayer) = 3 Then Wall039.SideImage = "Bank2" : Flasher005.visible = 0 : liMISS004.state = 2 : liRampM.state=2
    If ModeBankActive(CurrentPlayer) = 4 Then Wall039.SideImage = "Bank3" : Flasher005.visible = 0
  Elseif Mission4_BANK(CurrentPlayer) = 1 Then
    ModeBankActive(CurrentPlayer) = 4
    Wall039.SideImage = "Bank3"
    liMISS004.state = 0
    Flasher005.visible = 1
    FlasherMISS004.visible = 1
    liRampM.state=1
  End If
End Sub

Sub CheckAwardBank
  If Mission4_BANK(CurrentPlayer) = 0 Then
    AddScore 1000
    If Bank_Sav(CurrentPlayer) <=3 Then
      Bank_Sav(CurrentPlayer) = Bank_Sav(CurrentPlayer) + 1
    Elseif Bank_Sav(CurrentPlayer) >=4 And ModeBankActive(CurrentPlayer) = 0  Then
      ModeBankActive(CurrentPlayer) = 1
      Bank_Sav(CurrentPlayer) = 5
      li0START.State = 2
'   Elseif Bank_Sav(CurrentPlayer) >=4 And ModeBankActive(CurrentPlayer) = 1  Then
'     StartMissionBank
    Elseif Bank_Sav(CurrentPlayer) >=5 And ModeBankActive(CurrentPlayer) = 2  Then
      AwardBANK1
    Elseif Bank_Sav(CurrentPlayer) >=6 And ModeBankActive(CurrentPlayer) = 3  Then
      AwardBANK2
    End If
  Elseif Mission4_BANK(CurrentPlayer) = 1 Then
    Bank_Sav(CurrentPlayer) = Bank_Sav(CurrentPlayer) + 1
    AddScore 5000
    If Bank_Sav(CurrentPlayer) >=4 Then
      AddScore 25000
      DOF 128, DOFPulse
      PLaySound"Vo_Combo"
    ' pupevent 112
      dmdflush
        DMD "", "", "Combo", eNone, eNone, eNone, 3500, True, ""
      LightEffect 7
      ResetLightBANK
    End If
  End If
  CheckLightBANK
End Sub

Sub Trigger006_Hit
  DMD "      SCOUT","       BANK","ModBank",eBlink, eNone, eNone, 4000, True, ""
  PlaySound "fx_PasHorse"
End Sub


'''' MISSION MINE

Sub StartMissionMINE
  DMD "  SHOOT RAMP MINE","FILL WAGON WITH GOLD","ModMine",eBlink, eNone, eNone, 5000, True, ""
  ModeMineActive(CurrentPlayer) = 2
  Mine_Sav(CurrentPlayer) = 5
  liMISS005.state = 2
  li0RL1.state = 2
  li0RL2.state = 2
  li0RL3.state = 2
  li0RL4.state = 2
End Sub

Sub ResetLightMINE
  li0RL1.state = 0
  li0RL2.state = 0
  li0RL3.state = 0
  li0RL4.state = 0
  Mine_Sav(CurrentPlayer) = 0
End Sub

Sub CheckLightMINE
  If ModeCheatActive = True And Mine_Sav(CurrentPlayer) <= 3 And Mission5_MINE(CurrentPlayer) = 0 Then
    Mine_Sav(CurrentPlayer) = 3
    li0RL1.State = 1
    li0RL2.State = 1
    li0RL3.State = 1
    li0RL4.State = 0
  Else
    If Mine_Sav(CurrentPlayer) = 0 Then ResetLightMINE
    If Mine_Sav(CurrentPlayer) = 1 Then li0RL1.State = 1 : li0RL2.State = 0 : li0RL3.State = 0 : li0RL4.State = 0
    If Mine_Sav(CurrentPlayer) = 2 Then li0RL1.State = 1 : li0RL2.State = 1 : li0RL3.State = 0 : li0RL4.State = 0
    If Mine_Sav(CurrentPlayer) = 3 Then li0RL1.State = 1 : li0RL2.State = 1 : li0RL3.State = 1 : li0RL4.State = 0
    If Mine_Sav(CurrentPlayer) = 4 And Mission5_MINE(CurrentPlayer) = 0 Then li0RL1.State = 1 : li0RL2.State = 1 : li0RL3.State = 1 : li0RL4.State = 1 : ModeMineActive(CurrentPlayer) = 1 : li0START.State = 2
    If Mine_Sav(CurrentPlayer) = 4 And Mission5_MINE(CurrentPlayer) = 1 Then ResetLightMINE
    If Mine_Sav(CurrentPlayer) >= 5 Then StartMissionMINE
  End If
End Sub

Sub CheckAwardMine
  If Mission5_MINE(CurrentPlayer) = 0 Then
    AddScore 1000
    If Mine_Sav(CurrentPlayer) <=3 Then Mine_Sav(CurrentPlayer) = Mine_Sav(CurrentPlayer) + 1
    If Mine_Sav(CurrentPlayer) >=4 And ModeMineActive(CurrentPlayer) = 1  Then li0START.State = 2
    If Mine_Sav(CurrentPlayer) >=4 And ModeMineActive(CurrentPlayer) = 2  Then AwardMINE
  Elseif Mission5_MINE(CurrentPlayer) = 1 Then
    Mine_Sav(CurrentPlayer) = Mine_Sav(CurrentPlayer) + 1
    AddScore 5000
    If Mine_Sav(CurrentPlayer) >=4 Then
      AddScore 25000
      DOF 128, DOFPulse
      PLaySound"Vo_Combo"
    ' pupevent 112
      dmdflush
        DMD "", "", "Combo", eNone, eNone, eNone, 3500, True, ""
      LightEffect 7
      ResetLightMINE
    End If
  End If
  CheckLightMINE
End Sub

Sub AwardMine
  AddScore 20000
  Mission5_MINE(CurrentPlayer) = 1
  DMD "     WELL DONE","      WEALTH","ModMine",eBlink, eNone, eNone, 5000, True, ""
  ModeMineActive(CurrentPlayer) = 0
  ChangeBall(0)
  ResetLightMINE
  liMISS005.state = 0
  Flasher005.visible = 1
  FlasherMISS005.visible = 1
  StartCanon5
  CheckB2S
End Sub

Sub Trigger005_Hit
  ChangeDOOR
  Dim xx: For each xx in aGiLights: xx.color = RGB(255,180,100): Next 'Standard
  If ModeMineActive(CurrentPlayer) = 0 Then
  ChangeBall (0)
  End If
  If liSskillShoot.state = 0 Then
  CheckAwardMine
  End If
End Sub

Sub Trigger010_Hit
  Dim xx: For each xx in aGiLights: xx.color = RGB(255,255,0): Next 'Jaune
  PlaySound "fx_Train"'
  If Mission5_MINE(CurrentPlayer) = 0 And (ModeMineActive(CurrentPlayer) = 0 or ModeMineActive(CurrentPlayer) = 2) Then
    ShakeTrainON
    ChangeBall(4)
    Wall045.Collidable = True
    OpenMineTimer.enabled = 1
  Else
    Wall045.Collidable = False
  End If
End Sub

'***********
' Kicker
'***********
Sub KickerAuto_Timer()
  PlaySound "fx_kicker"
  KickerAuto.Kick 10, 30
  KickerAuto.timerEnabled = False
End Sub
Sub Kicker001_Timer()
  PlaySound "fx_kicker"
  Kicker001.Kick 90, 20
  Kicker001.timerEnabled = False
End Sub
Sub Kicker003_Timer()
    PlaySound SoundFXDOF("fx_kicker", 221, DOFPulse, DOFContactors)
  'PlaySound "fx_kicker"
  Kicker003.Kick 190, 4
  Kicker003.timerEnabled = False
End Sub
Sub Kicker004_Timer()
  PlaySound "fx_kicker"
  Kicker004.Kick 180, 5
  Kicker004.timerEnabled = False
End Sub
Sub Kicker005_Timer()
  ShakeJailON
  Kicker005.Destroyball
  ShakeGunL1
  ShakeGunR1
  PlaySound "fx_kicker"
  KickerAuto.CreateBall
  KickerAuto.Kick 10, 30
  Kicker005.timerEnabled = False
End Sub
Sub Kicker006_Timer()
    PlaySound SoundFXDOF("fx_kicker", 226, DOFPulse, DOFContactors)
  'PlaySound "fx_kicker"
  Kicker006.Kick 150, 5
  Kicker006.timerEnabled = False
End Sub

Sub Kicker008_Timer()
  If Mission2_DUEL(CurrentPlayer) = 0 or ModeDynamiteActive(CurrentPlayer) = 2 Then
    If ModeDuelActive(CurrentPlayer) = 0 Then
      If li0SHERIFF.state = 0 Then
        li0SHERIFF.state = 1
      Elseif li0SHERIFF.state = 1 Then
        li0SHERIFF.state = 2
      Elseif li0SHERIFF.state = 2 Then
        li0SHERIFF.state = 0
        ModeDuelActive(CurrentPlayer) = 1
        li0START.state = 2
      End If
    End If
    StartGlass
    Kicker008.Destroyball
    Kicker009.CreateBall
'   Kicker009.Kick -55, 95
    PlaySound SoundFXDOF("fx_kicker", 227, DOFPulse, DOFContactors)
    'PlaySound "fx_kicker"
    Kicker009.Kick -72, 39
    Kicker008.timerEnabled = False
  Else
    PlaySound "fx_electro"
    Kicker008.Destroyball
    Kicker010.CreateBall
    Kicker010.Kick 5, 10
        DOF 230, DOFPulse
    Kicker008.timerEnabled = False
  End If
End Sub


'***********************
' Glowball Section
'***********************
Dim GlowBall, CustomBulbIntensity(4)
Dim  GBred(4)
Dim GBgreen(4)
Dim GBblue(4)
Dim CustomBallImage(4), CustomBallLogoMode(4), CustomBallDecal(4), CustomBallGlow(4)
Dim GlowAura,GlowIntensity,ChooseBall

GlowAura=230 'GlowBlob Auroa radius
GlowIntensity=25'Glowblob intensity

' blue GlowBall
CustomBallGlow(1) =     True
GBred(1) = 1 : GBgreen(1) = 33 : GBblue(1) = 105
' Vert GlowBall
CustomBallGlow(2) =     True
GBred(2) = 1 : GBgreen(2) = 33 : GBblue(2) = 5
' Rouge GlowBall
CustomBallGlow(3) =     True
GBred(3) = 130 : GBgreen(3) = 1 : GBblue(3) = 1
' Orange GlowBall
CustomBallGlow(4) =     True
GBred(4) = 250 : GBgreen(4) = 130 : GBblue(4) = 1

Dim Glowing(10)
Set Glowing(0) = Glowball1 : Set Glowing(1) = Glowball2 : Set Glowing(2) = Glowball3 : Set Glowing(3) = Glowball4 : Set Glowing(4) = Glowball4

'*** change ball appearance ***

Sub ChangeBall(ballnr)
  Dim BOT, ii, col
  GlowBall = CustomBallGlow(ballnr)
  For ii = 0 to 4
    col = RGB(GBred(ballnr), GBgreen(ballnr), GBblue(ballnr))
    Glowing(ii).color = col : Glowing(ii).colorfull = col
  Next
End Sub

' *** Ball Shadow code / Glow Ball code / Primitive Flipper Update ***

Dim BallShadowArray
BallShadowArray = Array (BallShadow1, BallShadow2, BallShadow3,BallShadow4,BallShadow5)

Sub GraphicsTimer_Timer()
  Dim BOT, b
    BOT = GetBalls

  ' switch off glowlight for removed Balls
  IF GlowBall Then
    For b = UBound(BOT) + 1 to 3
      If GlowBall and Glowing(b).state = 1 Then Glowing(b).state = 0 End If
    Next
  End If

    For b = 0 to UBound(BOT)
    If GlowBall and b < 4 Then
      If Glowing(b).state = 0 Then Glowing(b).state = 1 end if
      Glowing(b).BulbHaloHeight = BOT(b).z + 32
      Glowing(b).x = BOT(b).x
      Glowing(b).y = BOT(b).Y+10
      Glowing(b).falloff=GlowAura 'GlowBlob Auroa radius
      Glowing(b).intensity=GlowIntensity 'Glowblob intensity
    End If
  Next
End Sub
'************
' Rotation Light Bonus
'************
Sub RotateLaneLightsLeftUp
    Dim TempState
    TempState = li016.State
    li016.State = li017.State
    li017.State = li018.State
    li018.State = TempState
End Sub

Sub RotateLaneLightsRightUp
    Dim TempState
    TempState = li018.State
    li018.State = li017.State
    li017.State = li016.State
    li016.State = TempState
End Sub

'*********
'Outlane
'*********
Sub Kicker007_Hit
  If li020.state=1 Then
  kicker007.Kick -1, 45
    PlaySound SoundFXDOF("fx_kicker", 231, DOFPulse, DOFContactors)
  'PlaySound "fx_kicker"
  li020.state=2
  Elseif li020.state=2 Then
  kicker007.Kick -1, 45
  'PlaySound "fx_kicker"
    PlaySound SoundFXDOF("fx_kicker", 231, DOFPulse, DOFContactors)
  li020.state=0
  Elseif li020.state=0 Then
  kicker007.Kick 90, 1
  End If
End Sub
Sub Kicker002_Hit
  If li021.state=1 Then
  PlaySound "fx_kicker"
    PlaySound SoundFXDOF("fx_kicker", 232, DOFPulse, DOFContactors)
  kicker002.Kick -1, 45
  li021.state=2
  Elseif li021.state=2 Then
  PlaySound "fx_kicker"
    PlaySound SoundFXDOF("fx_kicker", 232, DOFPulse, DOFContactors)
  kicker002.Kick -1, 45
  li021.state=0
  Elseif li021.state=0 Then
  kicker002.Kick -90, 1
  End If
End Sub

'**********
'MysterylIGHT
'**********
Sub Trigger029_Hit
    DOF 213, DOFPulse
  LaneLeftTimer.Enabled = True
  LaneLeftTimerEnd.Enabled = True
End Sub
Sub LaneLeftTimerEnd_Timer
  LaneLeftTimer.Enabled = False
End Sub
Sub Trigger030_Hit
  ChangeDOOR
  If LaneLeftTimer.Enabled = True And li0OUTLL.state = 0 Then
  li0OUTLL.state=2
  Elseif LaneLeftTimer.Enabled = True And li0OUTLL.state = 2 Then
  li0OUTLL.state=0
  li0MYST.state=2
  End If
End Sub

Sub Trigger018_Hit
  LaneRightTimer.Enabled = True
  LaneRightTimerEnd.Enabled = True
End Sub
Sub LaneRightTimerEnd_Timer
  LaneRightTimer.Enabled = False
End Sub
Sub Trigger019_Hit
    DOF 214, DOFPulse
  ChangeDOOR
  If LaneRightTimer.Enabled = True And li0OUTLL1.state = 0 Then
  li0OUTLL1.state=2
  Elseif LaneRightTimer.Enabled = True And li0OUTLL1.state = 2 Then
  li0OUTLL1.state=0
  li0MYST.state=2
  End If
End Sub
Sub ResetMystery
  li0OUTLL.state=0
  li0OUTLL1.state=0
  li0MYST.state=0
End Sub
'***********
' Sheriff STAR - Animation Girl
'***********
sub timer1_Timer()
StarGold.rotY=rainspin1.currentangle
MecaBull.rotY=rainspin2.currentangle
GirlbMeca.rotY=rainspin2.currentangle
GirlhMeca.rotY=rainspin2.currentangle
GirlCMeca.rotY=rainspin2.currentangle
GirlbMecaNude.rotY=rainspin2.currentangle
GirlCMecaNude.rotY=rainspin2.currentangle
Roue2.rotY=rainspin3.currentangle
Roue3.rotY=rainspin4.currentangle
Roue1.rotY=rainspin5.currentangle
end Sub

Sub Rainspin1_Spin
    PlaySoundAt "fx_spinner", rainspin1
    If Tilted Then Exit Sub
    Addscore 100
  If ModeSpinnersActive(CurrentPlayer) = true Then
    SpinnerCount = SpinnerCount - 1
    dmdflush
    DMD "TURN STAR - DYNAMITE", FormatScore(SpinnerCount), "ModDyna", eNone, eNone, eNone, 5000, True, ""
    If SpinnerCount = 0 Then
      li0START.state=2
      ModeDynamiteActive(CurrentPlayer) = 1
      ModeSpinnersActive(CurrentPlayer) = False
    end If
  end If
End Sub

'-------- Mission1_RODEO

dim Girlswitch
Girlswitch=0

sub GirlLOOK(Girlswitch)
select case Girlswitch
case 0 'off
  GirlbMeca.Visible = 0
  GirlhMeca.Visible = 0
  GirlCMeca.Visible = 0
  GirlbMecaNude.Visible = 0
  GirlCMecaNude.Visible = 0
  MecaBull.Visible = 0
  MecaBull2.Visible = 0
  Girl1b.Visible = 0
  Girl1h.Visible = 0
  Girl2b.Visible = 0
  Girl2h.Visible = 0
case 1 'Mode debout
  Coeur_Sav(CurrentPlayer) = 2
  GirlbMeca.Visible = 0
  GirlhMeca.Visible = 0
  GirlCMeca.Visible = 0
  GirlbMecaNude.Visible = 0
  GirlCMecaNude.Visible = 0
  MecaBull.Visible = 0
  MecaBull2.Visible = 0
  Girl1b.Visible = 0
  Girl1h.Visible = 0
  Girl2b.Visible = 1
  Girl2h.Visible = 1
case 2 'Mode Assise
  Coeur_Sav(CurrentPlayer) = 1
  GirlbMeca.Visible = 0
  GirlhMeca.Visible = 0
  GirlCMeca.Visible = 0
  GirlbMecaNude.Visible = 0
  GirlCMecaNude.Visible = 0
  MecaBull.Visible = 0
  MecaBull2.Visible = 0
  Girl1b.Visible = 1
  Girl1h.Visible = 1
  Girl2b.Visible = 0
  Girl2h.Visible = 0
  If ModeNudeActive = 1 Then
  Soutif2.Visible = 0
  Elseif ModeNudeActive = 0 Then
  Soutif2.Visible = 1
  End If
case 3 'Mode MecaBull
  Coeur_Sav(CurrentPlayer) = 3
  GirlbMeca.Visible = 1
  GirlhMeca.Visible = 1
  GirlCMeca.Visible = 1
  GirlbMecaNude.Visible = 0
  GirlCMecaNude.Visible = 0
  MecaBull.Visible = 1
  MecaBull2.Visible = 1
  Girl1b.Visible = 0
  Girl1h.Visible = 0
  Girl2b.Visible = 0
  Girl2h.Visible = 0
case 4 'Mode MecaBull NUDE
  If ModeNudeActive = 1 Then
  GirlbMeca.Visible = 0
  GirlbMecaNude.Visible = 1
  GirlhMeca.Visible = 1
  GirlCMeca.Visible = 0
  GirlCMecaNude.Visible = 1
  MecaBull.Visible = 1
  MecaBull2.Visible = 1
  Girl1b.Visible = 0
  Girl1h.Visible = 0
  Girl2b.Visible = 0
  Girl2h.Visible = 0
  Elseif ModeNudeActive = 0 Then
  GirlLOOK 3
  End If
end Select
end Sub

Sub StartMissionRODEO
  ModeRodeoActive(CurrentPlayer) = 2
  Coeur_Sav(CurrentPlayer) = 4
  li0START.State = 2
  liMISS001.state = 2
End Sub

Sub StopMissionRODEO
  AddScore 75200
  ModeRodeoActive(CurrentPlayer) = 0
  Mission1_RODEO(CurrentPlayer) = 1
  DMD "   SHE IS IN"," LOVE WITH YOU","ModRodeoOK",eNone, eNone, eNone, 1000, False, ""
  DMD "   SHE IS IN"," LOVE WITH YOU","ModRodeoOK2",eNone, eNone, eNone, 4000, True, ""
  liMISS001.state = 0
  FlasherMISS001.Visible = 1
  li0RR1.State=0
  li0RR2.State=0
  li0RR3.State=0
  li0RR4.State=0
  liGIRL001.State=1
  liGIRL002.State=1
  liGIRL003.State=1
  StartCanon5
  CheckB2S
End Sub
Sub Trigger011_Hit
    DOF 212, DOFPulse
  Dim xx: For each xx in aGiLights: xx.color = RGB(255,0,0): Next 'Rouge
  CheckLightHOT
  If GirlbMeca.Visible = True or GirlbMecaNude.Visible = True Then
  AddScore 2000
  PlaySound "Vo_GirlRodeo"
  End If
End Sub
Sub Trigger016_Hit
  ChangeDOOR
  Dim xx: For each xx in aGiLights: xx.color = RGB(255,180,100): Next 'Standard
  AddScore 300
End Sub
'***********
' Moove Chasseur
'***********
Sub ChasseurOFFTimer_timer()
Chasseur.RotY = Chasseur.RotY - 1
Revolver.RotY = Revolver.RotY - 1
Carabine.RotY = Carabine.RotY - 1
if Chasseur.RotY <= 75 then ChasseurOFFTimer.enabled = 0
End Sub
Sub ChasseurONTimer_timer()
Chasseur.RotY = Chasseur.RotY + 1
Revolver.RotY = Revolver.RotY + 1
Carabine.RotY = Carabine.RotY + 1
if Chasseur.RotY >= 155 then ChasseurONTimer.enabled = 0
End Sub

Sub ShakeChasseurOFF
  If ChasseurONTimer.enabled = 0 Then
  ChasseurOFFTimer.enabled = 1
  End If
End Sub
Sub ShakeChasseurON
  If ChasseurOFFTimer.enabled = 0 And liLOCK.state = 0 And liMB.state = 0 Then
  ChasseurONTimer.enabled = 1
  PlaySound "Vo_EnemieAhead"
  End If
End Sub
'***********
' Moove Jail
'***********
Sub JailOFFTimer_timer()
GrilleJail.TransY = GrilleJail.TransY - 1
if GrilleJail.TransY <= 1 then JailOFFTimer.enabled = 0
End Sub
Sub JailONTimer_timer()
GrilleJail.TransY = GrilleJail.TransY + 1
if GrilleJail.TransY >= 55 then JailONTimer.enabled = 0
End Sub

Sub ShakeJailOFF 'open
  Wall044.Collidable = False
  ShakeChasseurON
  If JailONTimer.enabled = 0 Then
  JailOFFTimer.enabled = 1
  End If
End Sub
Sub ShakeJailON 'Close
  Wall044.Collidable = True
  ShakeChasseurOFF
  If JailOFFTimer.enabled = 0 Then
  JailONTimer.enabled = 1
  AddScore 1000
  End If
End Sub
'***********
' Moove PAL
'***********
Sub PalOFFTimer_timer()
MoulinHELI.RotY = MoulinHELI.RotY + 1
MoulinPAL.RotY = MoulinPAL.RotY + 1
MoulinPAL2.RotY = MoulinPAL2.RotY + 1
if MoulinHELI.RotY >= -30 then PalOFFTimer.enabled = 0
End Sub
Sub PalONTimer_timer()
MoulinHELI.RotY = MoulinHELI.RotY - 5
MoulinPAL.RotY = MoulinPAL.RotY - 5
MoulinPAL2.RotY = MoulinPAL2.RotY - 5
if MoulinHELI.RotY <= -160 then
PalONTimer.enabled = 0
PalOFFTimer.enabled = 1
End If
End Sub

Sub ShakePal 'Open
  If PalOFFTimer.enabled = 0 Then
  PalONTimer.enabled = 1
  AddScore 1000
  End If
End Sub
Sub Trigger015_Hit
  ShakePal
  ChangeDOOR
  PlaySound "fx_Moulin"
  Dim xx: For each xx in aGiLights: xx.color = RGB(255,180,100): Next 'Standard
End Sub
'***********
' Moove Train
'***********
Sub TrainOFFTimer_timer()
Train2.TransX = Train2.TransX - 2
TrainGold.visible=1
TrainGold.TransX = TrainGold.TransX - 2
if Train2.TransX <= 1 then TrainOFFTimer.enabled = 0
End Sub
Sub TrainONTimer_timer()
Train2.TransX = Train2.TransX + 2
TrainGold.visible=0
TrainGold.TransX = TrainGold.TransX + 2
if Train2.TransX >= 930 then TrainONTimer.enabled = 0
End Sub

Sub ShakeTrainOFF
  If TrainONTimer.enabled = 0 Then
  TrainOFFTimer.enabled = 1
  End If
End Sub
Sub ShakeTrainON
  If TrainOFFTimer.enabled = 0 Then
  TrainONTimer.enabled = 1
  AddScore 10000
  End If
End Sub

Sub OpenMineTimer_timer()
OpenMine2Timer.enabled = 1
OpenMineTimer.enabled = 0
End Sub
Sub OpenMine2Timer_timer()
ShakeTrainOFF
OpenMine3Timer.enabled = 1
OpenMine2Timer.enabled = 0
DMDFlush
DMD "  COLLECT GOLD","","ModMine001",eBlink, eNone, eNone, 2100, FALSE, ""
DMD "  COLLECT GOLD","","ModMine002",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine003",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine004",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine005",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine006",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine007",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine008",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine009",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine010",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine011",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine012",eBlink, eNone, eNone, 150, FALSE, ""
DMD "  COLLECT GOLD","","ModMine013",eBlink, eNone, eNone, 150, True, ""
End Sub
Sub OpenMine3Timer_timer()
Wall045.Collidable = False
OpenMine3Timer.enabled = 0
End Sub

'****************************************
'StartWheelSpinner DYNAMITE MODE
'****************************************
Sub StartMissionDYNAMITE
  DMD "   SHOOT","   DYNAMITE","ModDyna",eBlink, eNone, eNone, 5000, True, ""
  ModeDynamiteActive(CurrentPlayer) = 2
  ModeSpinnersActive(CurrentPlayer) = False
  SpinnerCount = 25
  FlasherDyna001.visible = 1
  FlasherDyna002.visible = 1
  FlasherDyna003.visible = 1
  FlasherDyna004.visible = 1
  liMISS003.state = 2
  spinner3.MotorOn = True
  spinning3.enabled = True
End Sub

Sub StopMissionDYNAMITE
  Mission3_DYNAMITE(CurrentPlayer) = 1
  DMD " MISSION","SUCCESSFUL","ModDyna",eBlink, eNone, eNone, 5000, True, ""
  ModeDynamiteActive(CurrentPlayer) = 0
  ModeSpinnersActive(CurrentPlayer) = False
  SpinnerCount = 25
  liMISS003.state = 0
  FlasherMISS003.visible = 1
  StartCanon5
  AddScore 30000
  FlasherDyna001.visible = 0
  FlasherDyna002.visible = 0
  FlasherDyna003.visible = 0
  FlasherDyna004.visible = 0
  spinner3.MotorOn = False
  spinning3.enabled = False
  CheckB2S
End Sub

Sub CheckDYNA
  If ModeDynamiteActive(CurrentPlayer) = 2 And Dyna_Sav(CurrentPlayer) = 4 Then
  StopMissionDYNAMITE
  End If
End Sub

Sub TriggerDYNA001_Hit
  If FlasherDyna001.visible = True And ModeDynamiteActive(CurrentPlayer) = 2 Then
  Dyna_Sav(CurrentPlayer) = Dyna_Sav(CurrentPlayer) + 1
  AddScore 3000
  FlasherDyna001.visible = 0
  PlaySound "fx_Booum"
  StartCanon1
  CheckDYNA
  End If
End Sub

Sub TriggerDYNA002_Hit
  ChangeDOOR
  Dim xx: For each xx in aGiLights: xx.color = RGB(255,180,100): Next 'Standard
  If FlasherDyna002.visible = True And ModeDynamiteActive(CurrentPlayer) = 2 Then
  Dyna_Sav(CurrentPlayer) = Dyna_Sav(CurrentPlayer) + 1
  AddScore 3000
  PlaySound "fx_Booum"
  FlasherDyna002.visible = 0
  StartCanon2
  CheckDYNA
  End If
End Sub
Sub TriggerDYNA003_Hit
  Star_Sav(CurrentPlayer) = 1
  If FlasherDyna003.visible = True And ModeDynamiteActive(CurrentPlayer) = 2 Then
  Dyna_Sav(CurrentPlayer) = Dyna_Sav(CurrentPlayer) + 1
  AddScore 3000
  PlaySound "fx_Booum"
  FlasherDyna003.visible = 0
  li0SHERIFF.state = 1
  StartCanon3
  CheckDYNA
  End If
End Sub
Sub TriggerDYNA004_Hit
  DOF 210, DOFPulse
  Dim xx: For each xx in aGiLights: xx.color = RGB(0,255,0): Next 'Vert
  CheckAwardBank
  If FlasherDyna004.visible = True And ModeDynamiteActive(CurrentPlayer) = 2 Then
  Dyna_Sav(CurrentPlayer) = Dyna_Sav(CurrentPlayer) + 1
  AddScore 3000
  PlaySound "fx_Booum"
  FlasherDyna004.visible = 0
  StartCanon4
  CheckDYNA
  End If
End Sub

Sub WheelImageSpin3_Timer '
  FlasherDyna001.rotz = FlasherDyna001.rotz - 1
  FlasherDyna002.rotz = FlasherDyna002.rotz + 1
  FlasherDyna003.rotz = FlasherDyna003.rotz + 1
  FlasherDyna004.rotz = FlasherDyna004.rotz - 1
  FlasherDyna005.rotz = FlasherDyna005.rotz - 1
End Sub
Sub spinning3_timer
  WheelImageSpin3.Enabled=False
  FlasherDyna001.rotz = FlasherDyna001.rotz - 5
  FlasherDyna002.rotz = FlasherDyna002.rotz + 5
  FlasherDyna003.rotz = FlasherDyna003.rotz + 5
  FlasherDyna004.rotz = FlasherDyna004.rotz - 5
  FlasherDyna005.rotz = FlasherDyna005.rotz - 5
End Sub

dim spinner3
Set spinner3 = New cvpmTurntable
With spinner3
  .InitTurntable spindisc_wheel3, 100
  .SpinDown = 20
  .CreateEvents "spinner3"
End With
spinner3.MotorOn = false

'*************
'Flasher Canon
'*************
Dim CanonPos, Canon2Pos, Canon3Pos, Canon4Pos, Canon5Pos, Canon6Pos, Canon7Pos, FramesCanon
FramesCanon = Array("Canon1", "Canon2", "Canon3", "Canon4", "Canon5", "Canon6", "Canon7", "Canon8", "Canon9")

Sub StartCanon1
  Canon.visible = 1
  CanonPos = 0
    CanonTimer.Enabled = 1
  StopCanonTimer.Enabled = 1
End Sub
Sub StartCanon2
  Canon2.visible = 1
  Canon2Pos = 0
    CanonTimer.Enabled = 1
  StopCanonTimer.Enabled = 1
End Sub
Sub StartCanon3
  Canon3.visible = 1
  Canon3Pos = 0
    CanonTimer.Enabled = 1
  StopCanonTimer.Enabled = 1
End Sub
Sub StartCanon4
  Canon4.visible = 1
  Canon4Pos = 0
    CanonTimer.Enabled = 1
  StopCanonTimer.Enabled = 1
End Sub
Sub StartCanon5
  Canon5.visible = 1
  Canon5Pos = 0
  PlaySound "fx_gun06"
    CanonTimer.Enabled = 1
  StopCanonTimer.Enabled = 1
End Sub
Sub StartCanon6
  Canon6.visible = 1
  Canon6Pos = 0
  PlaySound "fx_gun06"
    CanonTimer.Enabled = 1
  StopCanonTimer.Enabled = 1
End Sub
Sub StartCanon7
  Canon7.visible = 1
  Canon7Pos = 0
  PlaySound "fx_gun06"
    CanonTimer.Enabled = 1
  StopCanonTimer.Enabled = 1
End Sub
Sub CanonTimer_Timer
    Canon.ImageA = FramesCanon (CanonPos)
    CanonPos = (CanonPos + 1) MOD 9
  Canon2.ImageA = FramesCanon (Canon2Pos)
    Canon2Pos = (Canon2Pos + 1) MOD 9
  Canon3.ImageA = FramesCanon (Canon3Pos)
    Canon3Pos = (Canon3Pos + 1) MOD 9
  Canon4.ImageA = FramesCanon (Canon4Pos)
    Canon4Pos = (Canon4Pos + 1) MOD 9
  Canon5.ImageA = FramesCanon (Canon5Pos)
    Canon5Pos = (Canon5Pos + 1) MOD 9
  Canon6.ImageA = FramesCanon (Canon6Pos)
    Canon6Pos = (Canon6Pos + 1) MOD 9
  Canon7.ImageA = FramesCanon (Canon7Pos)
    Canon7Pos = (Canon7Pos + 1) MOD 9
  vpmtimer.addtimer 720, "StopCanon '"
End Sub
Sub StopCanon
  Canon.visible = 0
  Canon2.visible = 0
  Canon3.visible = 0
  Canon4.visible = 0
  Canon5.visible = 0
  Canon6.visible = 0
  Canon7.visible = 0
  CanonTimer.Enabled = 0
End Sub
Sub StopCanonTimer_Timer
  StopCanon
Me.Enabled = 0
End Sub

Sub Trigger017_Hit 'UPDATE event DEBUG
  ' Bank
  CheckB2S
  ' Poker SALOON
  If ModeDynamiteActive(CurrentPlayer) = 1 or ModePokerActive(CurrentPlayer) = 1 or ModeMineActive(CurrentPlayer) = 1 or ModeDuelActive(CurrentPlayer) = 1 or ModeRodeoActive(CurrentPlayer) = 1 or ModeSpinnersActive(CurrentPlayer) = 1 Then li0START.state = 2 Else li0START.state = 0
  If li0R2M1.State = 2 And li0R2M2.State = 2 And li0R2M3.State = 2 And li0R2M4.State = 2 And li0R2M5.State = 2 And li0R2M6.State = 2 And Mission5_MINE(CurrentPlayer) = 0 And ModePokerActive(CurrentPlayer) = 0 Then
  li0START.state = 2
  ModePokerActive(CurrentPlayer) = 1
  CheckLightPoker
  End If
  If Mission1_RODEO(CurrentPlayer) = 1 And Mission2_DUEL(CurrentPlayer) = 1 And Mission3_DYNAMITE(CurrentPlayer) = 1 And Mission4_BANK(CurrentPlayer) = 1 And Mission5_MINE(CurrentPlayer) = 1 And Mission6_POKER(CurrentPlayer) = 1 Then
  AwardExtraBall
  AddMBmission2
  AddMBmission3
  EnableBallSaver 15
  ResetLightEvent
  DMDFlush
  DMD " ALL MISSIONS","SUCCESSFUL","",eBlink, eNone, eNone, 5000, True, ""
  AddScore 100000
  End If
End Sub

Sub CheckB2S
  If B2SOn Then
    If Mission1_RODEO(CurrentPlayer) = 1 Then
    Controller.B2SSetData 12, 0
    Controller.B2SSetData 120, 0
    Controller.B2SSetData 32, 1
      Elseif Mission1_RODEO(CurrentPlayer) = 0 And ModeRodeoActive(CurrentPlayer) >= 1 Then
      Controller.B2SSetData 12, 1
      Controller.B2SSetData 120, 1
      Controller.B2SSetData 32, 0
      Elseif Mission1_RODEO(CurrentPlayer) = 0 And ModeRodeoActive(CurrentPlayer) = 0 Then
      Controller.B2SSetData 12, 0
      Controller.B2SSetData 120, 0
      Controller.B2SSetData 32, 0
    End If
    If Mission2_DUEL(CurrentPlayer) = 1 Then
      Controller.B2SSetData 11, 0
      Controller.B2SSetData 110, 0
      Controller.B2SSetData 31, 1
      Elseif Mission2_DUEL(CurrentPlayer) = 0 And ModeDuelActive(CurrentPlayer) >= 1 Then
      Controller.B2SSetData 11, 1
      Controller.B2SSetData 110, 1
      Controller.B2SSetData 31, 0
      Elseif Mission2_DUEL(CurrentPlayer) = 0 And ModeDuelActive(CurrentPlayer) = 0 Then
      Controller.B2SSetData 11, 0
      Controller.B2SSetData 110, 0
      Controller.B2SSetData 31, 0
    End If
    If Mission3_DYNAMITE(CurrentPlayer) = 1 Then
    Controller.B2SSetData 14, 0
    Controller.B2SSetData 140, 0
    Controller.B2SSetData 34, 1
      Elseif Mission3_DYNAMITE(CurrentPlayer) = 0 And ModeDynamiteActive(CurrentPlayer) >= 1 Then
      Controller.B2SSetData 14, 1
      Controller.B2SSetData 140, 1
      Controller.B2SSetData 34, 0
      Elseif Mission3_DYNAMITE(CurrentPlayer) = 0 And ModeDynamiteActive(CurrentPlayer) = 0 Then
      Controller.B2SSetData 14, 0
      Controller.B2SSetData 140, 0
      Controller.B2SSetData 34, 0
    End If
    If Mission4_BANK(CurrentPlayer) = 1 Then
    Controller.B2SSetData 10, 0
    Controller.B2SSetData 100, 0
    Controller.B2SSetData 30, 1
      Elseif Mission4_BANK(CurrentPlayer) = 0 And ModeBankActive(CurrentPlayer) >= 1 Then
      Controller.B2SSetData 10, 1
      Controller.B2SSetData 100, 1
      Controller.B2SSetData 30, 0
      Elseif Mission4_BANK(CurrentPlayer) = 0 And ModeBankActive(CurrentPlayer) = 0 Then
      Controller.B2SSetData 10, 0
      Controller.B2SSetData 100, 0
      Controller.B2SSetData 30, 0
    End If
    If Mission5_MINE(CurrentPlayer) = 1 Then
    Controller.B2SSetData 13, 0
    Controller.B2SSetData 130, 0
    Controller.B2SSetData 33, 1
      Elseif Mission5_MINE(CurrentPlayer) = 0 And ModeMineActive(CurrentPlayer) >= 1 Then
      Controller.B2SSetData 13, 1
      Controller.B2SSetData 130, 1
      Controller.B2SSetData 33, 0
      Elseif Mission5_MINE(CurrentPlayer) = 0 And ModeMineActive(CurrentPlayer) = 0 Then
      Controller.B2SSetData 13, 0
      Controller.B2SSetData 130, 0
      Controller.B2SSetData 33, 0
    End If
    If Mission6_POKER(CurrentPlayer) = 1 Then
    Controller.B2SSetData 15, 0
    Controller.B2SSetData 150, 0
    Controller.B2SSetData 35, 1
      Elseif Mission6_POKER(CurrentPlayer) = 0 And ModePokerActive(CurrentPlayer) >= 1 Then
      Controller.B2SSetData 15, 1
      Controller.B2SSetData 150, 1
      Controller.B2SSetData 35, 0
      Elseif Mission6_POKER(CurrentPlayer) = 0 And ModePokerActive(CurrentPlayer) = 0 Then
      Controller.B2SSetData 15, 0
      Controller.B2SSetData 150, 0
      Controller.B2SSetData 35, 0
    End If
    If li0OUTLL2.state = 2 Then
      Controller.B2SSetData 16, 1
      Controller.B2SSetData 160, 1
      Elseif li0OUTLL2.state = 1 or li0OUTLL2.state = 0 Then
      Controller.B2SSetData 16, 0
      Controller.B2SSetData 160, 0
    End If
  Controller.B2SSetData 20, 1
  End If
End Sub

Sub Kicker011_Hit
PlaySound "fx_kicker"
Kicker011.Kick -45, 21
Kicker011.timerEnabled = False
End Sub

Sub ChangeDOOR
Dim ModeChosenDOOR
    ModeChosenDOOR = False
    Do While ModeChosenDOOR = False
      Select Case RndNbr(12)
      Case 1
        FlasherETG1A.Visible = 0
        FlasherETG1B.Visible = 0
        ModeChosenDOOR = True
      Case 2
        FlasherETG2A.Visible = 0
        FlasherETG2B.Visible = 0
        ModeChosenDOOR = True
      Case 3
        FlasherETG3.Visible = 0
        ModeChosenDOOR = True
      Case 4
        FlasherETG4A.Visible = 0
        FlasherETG4B.Visible = 0
        ModeChosenDOOR = True
      Case 5
        FlasherRDC1.Visible = 0
        ModeChosenDOOR = True
      Case 6
        FlasherRDC2A.Visible = 0
        FlasherRDC2B.Visible = 0
        ModeChosenDOOR = True
      Case 7
        FlasherETG1A.Visible = 1
        FlasherETG1B.Visible = 1
        ModeChosenDOOR = True
      Case 8
        FlasherETG2A.Visible = 1
        FlasherETG2B.Visible = 1
        ModeChosenDOOR = True
      Case 9
        FlasherETG3.Visible = 1
        ModeChosenDOOR = True
      Case 10
        FlasherETG4A.Visible = 1
        FlasherETG4B.Visible = 1
        ModeChosenDOOR = True
      Case 11
        FlasherRDC1.Visible = 1
        ModeChosenDOOR = True
      Case 12
        FlasherRDC2A.Visible = 1
        FlasherRDC2B.Visible = 1
        ModeChosenDOOR = True
      End Select
    Loop
End SUB

'Flipper Correction Initialization early 90’s and after

  Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

'
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
    x.AddPt "Polarity", 1, 0.05, - 5.5
    x.AddPt "Polarity", 2, 0.16, - 5.5
    x.AddPt "Polarity", 3, 0.20, - 0.75
    x.AddPt "Polarity", 4, 0.25, - 1.25
    x.AddPt "Polarity", 5, 0.3, - 1.75
    x.AddPt "Polarity", 6, 0.4, - 3.5
    x.AddPt "Polarity", 7, 0.5, - 5.25
    x.AddPt "Polarity", 8, 0.7, - 4.0
    x.AddPt "Polarity", 9, 0.75, - 3.5
    x.AddPt "Polarity", 10, 0.8, - 3.0
    x.AddPt "Polarity", 11, 0.85, - 2.5
    x.AddPt "Polarity", 12, 0.9, - 2.0
    x.AddPt "Polarity", 13, 0.95, - 1.5
    x.AddPt "Polarity", 14, 1, - 1.0
    x.AddPt "Polarity", 15, 1.05, -0.5
    x.AddPt "Polarity", 16, 1.1, 0
    x.AddPt "Polarity", 17, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.23, 0.85
    x.AddPt "Velocity", 2, 0.27, 1
    x.AddPt "Velocity", 3, 0.3, 1
    x.AddPt "Velocity", 4, 0.35, 1
    x.AddPt "Velocity", 5, 0.6, 1 '0.982
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
        If FlipperTrigger(gBOT(b).x, BOT(b).y, Flipper2) Then
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
'  Const EOSReturn = 0.045  'late 70's to mid 80's
' Const EOSReturn = 0.035  'mid 80's to early 90's
   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

'Sub FlipperDeactivate(Flipper, FlipperPress)
' FlipperPress = 0
' Flipper.eostorqueangle = EOSA
' Flipper.eostorque = EOST * EOSReturn / FReturn

' If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
'   Dim b', BOT
    '   BOT = GetBalls

'   For b = 0 To UBound(gBOT)
'     If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
'       If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
'     End If
'   Next
' End If
'End Sub

'Sub FlipperDeactivate(Flipper, FlipperPress)
' FlipperPress = 0
' Flipper.eostorqueangle = EOSA
' Flipper.eostorque = EOST * EOSReturn / FReturn
'
' If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
'   Dim b', BOT
'   '   BOT = GetBalls   ' if you use LF.ReProcessBalls / gBOT elsewhere, this can stay commented
'
'   For b = 0 To UBound(gBOT)
'           ' make sure this slot actually holds a ball object
'            If Not (gBOT(b) Is Nothing) Then
'         If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
'           If gBOT(b).vely >= -0.4 Then gBOT(b).vely = -0.4
'         End If
'            End If
'   Next
' End If
'End Sub


Sub FlipperDeactivate(Flipper, FlipperPress)
    Dim BOT, b

    FlipperPress = 0
    Flipper.eostorqueangle = EOSA
    Flipper.eostorque = EOST * EOSReturn / FReturn

    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
        BOT = GetBalls   ' fresh list of valid balls on the table

        If IsArray(BOT) Then
            For b = 0 To UBound(BOT)
                If IsObject(BOT(b)) Then
                    ' check for cradle near this flipper
                    If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then
                        ' clamp downward speed a bit
                        If BOT(b).VelY >= -0.4 Then BOT(b).VelY = -0.4
                    End If
                End If
            Next
        End If
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
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub


Dim LS: Set LS = New SlingshotCorrection
Dim RS: Set RS = New SlingshotCorrection
'Dim ULS: Set ULS = New SlingshotCorrection
'Dim URS: Set URS = New SlingshotCorrection

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

' The following sub are needed, however they exist in the ZMAT maths section of the script. Uncomment below if needed
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
'****  END FLIPPER CORRECTIONS
'******************************************************
Sub leftInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  if activeball.vely < 0 then exit sub              'don't affect upwards movement
    activeball.AngMomZ = -abs(activeball.AngMomZ) * RndNum(3,6)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

    if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


Sub rightInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  if activeball.vely < 0 then exit sub              'don't affect upwards movement
    activeball.AngMomZ = abs(activeball.AngMomZ) * RndNum(2,4)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

  if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub
'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

' Put all the Post and Pin objects in dPosts collection. Make sure dPosts fires hit events.
Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

' This collection contains the bottom sling posts. They are not in the dPosts collection so that the TargetBouncer is not applied to them, but they should still have dampening applied
' If you experience airballs with posts or targets, consider adding them to this collection
Sub NoTargetBouncer_Hit
    RubbersD.dampen ActiveBall
End Sub

' Put all the Sleeve objects in dSleeves collection. Make sure dSleeves fires hit events.
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
    aBall.velz = aBall.velz * coef
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
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7-1.0

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

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



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************

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

Dim gBot
Sub GameTimer_Timer()
  Cor.Update
  gBOT = GetBalls
  RollingUpdate
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
RollingSoundFactor = 3 / 5

Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1

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

Sub Posts_Hit(idx)
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

'/////////////////////////////  GI AND FLASHER   ////////////////////////////

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
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************


'******************************************************
' ZBRL: BALL ROLLING AND DROP SOUNDS
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
  Dim b', BOT
' BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ < -1 and gBOT(b).z < 55 and gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
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



'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
dim RampBalls(5,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(5)

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)  'Add ball
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  dim x : for x = 1 to uBound(RampBalls)  'Check, don't add balls twice
    if RampBalls(x, 1) = input.id then
      if Not IsEmpty(RampBalls(x,1) ) then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 to uBound(RampBalls)
    if IsEmpty(RampBalls(x, 1)) then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      exit Sub
    End If
    if x = uBound(RampBalls) then   'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
      RampBalls(0, 0) & vbnewline & _
      Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
      Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
      Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
      Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
      Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
      " "
    End If
  next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)   'Remove ball
  'Debug.Print "In WRemoveBall() + Remove ball from loop array"
  dim ballcount : ballcount = 0
  dim x : for x = 1 to Ubound(RampBalls)
    if ID = RampBalls(x, 1) then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
  next
  if BallCount = 0 then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end if
  next
  if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
  "1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
  "2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
  "3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
  "4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
  "5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
  "6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
  " "
End Sub


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function


'Ramp triggers
sub StartRamp_Hit()
  If activeball.velx < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

sub StartRamp001_Hit()
  If activeball.velx < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

sub StartWire1_Hit()
    WireRampOn False
End Sub

sub StartRamp2_Hit()
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

sub StartRamp3_Hit()
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub StartRamp4_Hit()
    WireRampOn False
End Sub

sub StartRamp5_Hit()
  If activeball.velx < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub StartRamp6_Hit()
    WireRampOff
    WireRampOn False
End Sub

sub StartWireRamp_Hit()
  If activeball.vely < 0 Then
        WireRampOff
    WireRampOn False
  Else
        WireRampOn True
    WireRampOff
  End If
End Sub

sub StartWireRamp2_Hit()
  If activeball.vely < 0 Then
        WireRampOff
    WireRampOn False
  Else
        WireRampOn True
    WireRampOff
  End If
End Sub

Sub Rampend_hit()
    WireRampOff
End Sub

Sub Rampend3_hit()
    WireRampOff
End Sub

Sub Rampend2_hit()
  If activeball.vely < 0 Then
    WireRampOff
    Else
    WireRampOn True
  End If
End Sub

Sub Rampend4_hit()
    WireRampOff
End Sub

Sub Rampend5_hit()
    WireRampOff
End Sub

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

'*************
'VR Stuff
'*************
Dim VRThings
Dim VRRoom
Dim DMD_Initialised : DMD_Initialised = False ' Flag pour ne le faire qu'une fois

Sub LoadVRRoom
    ' 1. Cacher les objets VR par défaut
    For Each VRThings In VR_Cab : VRThings.Visible = 0 : Next
    For Each VRThings In VR_Min : VRThings.Visible = 0 : Next
    For Each VRThings In VR_Mega: VRThings.Visible = 0 : Next

    If RenderingMode = 2 Or VRTest Then
        VRRoom = VRRoomChoice

        ' Objets à masquer en VR
        Flasher6.Visible = 0
        Flasher7.Visible = 0
        ramp1.Visible = 0
        ramp003.Visible = 0
        wall34.SideVisible = 0
        wall35.SideVisible = 0
        rrail.Visible = 0
        lrail.Visible = 0
        wall21.SideVisible = 0
        wall21.Visible = 0

        '=========================================================
        ' POSITIONNEMENT DU DMD (Exécuté une seule fois)
        '=========================================================
        If Not DMD_Initialised Then
            ' On applique les offsets de position ici
            digitgrid.rotx = -86
            digitgrid.x = digitgrid.x + 1020
            digitgrid.y = digitgrid.y - 365
            digitgrid.height = digitgrid.height + 280

            digit041.rotx = -86
            digit041.x = digit041.x + 1020
            digit041.y = digit041.y - 365
            digit041.height = digit041.height + 280

            Dim VrObj
            For Each VrObj In DMDUpper
                VrObj.rotx = -86
                VrObj.x = VrObj.x + 1020
                VrObj.y = VrObj.y - 345
                VrObj.height = VrObj.height + 297
            Next
            For Each VrObj In DMDLower
                VrObj.rotx = -86
                VrObj.x = VrObj.x + 1020
                VrObj.y = VrObj.y - 384
                VrObj.height = VrObj.height + 272
            Next

            DMD_Initialised = True ' On verrouille pour les prochains appels (F12)
        End If
        '=========================================================

    Else
        VRRoom = 0
    End If

    ' Gestion des types de salles VR
    If VRRoom = 1 Then
        For Each VRThings In VR_Cab : VRThings.Visible = 1 : Next
    ElseIf VRRoom = 2 Then
        For Each VRThings In VR_Cab : VRThings.Visible = 1 : Next
        For Each VRThings In VR_Min : VRThings.Visible = 1 : Next
    ElseIf VRRoom = 3 Then
        For Each VRThings In VR_Cab : VRThings.Visible = 1 : Next
        For Each VRThings In VR_Mega: VRThings.Visible = 1 : Next
    End If
End Sub

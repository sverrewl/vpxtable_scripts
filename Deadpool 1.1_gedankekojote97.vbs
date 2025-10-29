' ****************************************************************
'             JP's Deadpool for VISUAL PINBALL X 10.7
'              Including JP's Arcade Physics 3.0.1
' ****************************************************************

'On the DOF website put Exxx
'DOF some updating by outhere
'101 Left Flipper
'102 Right Flipper
'103 left slingshot
'104 right slingshot
'105
'106
'107 Center Bumper
'108 RIGHT Bumper
'109 Left bumper
'110
'111 HellHouse
'118
'119 Reset drop Targets
'120 AutoFire
'122 knocker
'123 ballrelease


Option Explicit
Randomize

'//////////////////////////////////////////////////////////////////////
'// OPTIONS
'//////////////////////////////////////////////////////////////////////

Const MusicAttractMode = 1          '0 - No Music in Attract (recommended with PUP), 1 - Play Music in Attract Mode

Const Siderails =  1                '0 - No Siderails, 1 - Siderails
Const BallBright = 1        '0 - Normal, 1 - Bright
Const SongVolume = 0.15
Const VolumeDial = 0.8
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

'//////////////////////////////////////////////////////////////////////
'// PUP OPTIONS
'//////////////////////////////////////////////////////////////////////

Const FPS = 1     ' Slow PC Recommended "0" Fast PC Recommended 1

Const PupScreen6 = 0    'LeftApron  0=Off - 1= On - 2= On High Quality Slow PC 0 or 1
Const PupScreen7 = 0    'RightApron  0=Off - 1= On - 2= On High Quality Slow PC 0 or 1

'//////////////////////////////////////////////////////////////////////


Const BallSize = 50
Const BallMass = 1
Const tnob = 5
Const lob = 0

Const AmbientBallShadowOn = 1
Const DynamicBallShadowsOn = 1
Const RubberizerEnabled = 1
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

'//////////////////////////////////////////////////////////////////////
'// LUT
'//////////////////////////////////////////////////////////////////////

'//////////////---- LUT (Colour Look Up Table) ----//////////////
'0 = Fleep Natural Dark 1
'1 = Fleep Natural Dark 2
'2 = Fleep Warm Dark
'3 = Fleep Warm Bright
'4 = Fleep Warm Vivid Soft
'5 = Fleep Warm Vivid Hard
'6 = Skitso Natural and Balanced
'7 = Skitso Natural High Contrast
'8 = 3rdaxis Referenced THX Standard
'9 = CalleV Punchy Brightness and Contrast
'10 = HauntFreaks Desaturated
'11 = Tomate Washed Out
'12 = VPW Original 1 to 1
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book

Dim LUTset, DisableLUTSelector, LutToggleSound, bLutActive
LutToggleSound = True
LoadLUT
'LUTset = 0     ' Override saved LUT for debug
SetLUT
DisableLUTSelector = 0  ' Disables the ability to change LUT option with magna saves in game when set to 1

Sub SetLUT  'AXS
  Table1.ColorGradeImage = "LUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  Select Case LUTSet
    Case 0: LUTBox.text = "Fleep Natural Dark 1"
    Case 1: LUTBox.text = "Fleep Natural Dark 2"
    Case 2: LUTBox.text = "Fleep Warm Dark"
    Case 3: LUTBox.text = "Fleep Warm Bright"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard"
    Case 6: LUTBox.text = "Skitso Natural and Balanced"
    Case 7: LUTBox.text = "Skitso Natural High Contrast"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast"
    Case 10: LUTBox.text = "HauntFreaks Desaturated"
      Case 11: LUTBox.text = "Tomate washed out"
        Case 12: LUTBox.text = "VPW original 1on1"
        Case 13: LUTBox.text = "bassgeige"
        Case 14: LUTBox.text = "blacklight"
        Case 15: LUTBox.text = "B&W Comic Book"
    Case 16: LUTBox.text = "Skitso New ColorLut"
  End Select
  LUTBox.TimerEnabled = 1
End Sub

Sub SaveLUT
  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if

  if LUTset = "" then LUTset = 0 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "DeadpoolLUT.txt",True)
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub
Sub LoadLUT
    bLutActive = False
  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    LUTset=0
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "DeadpoolLUT.txt") then
    LUTset=0
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "DeadpoolLUT.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      LUTset=0
      Exit Sub
    End if
    LUTset = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

'//////////////////////////////////////////////////////////////////////

'FlexDMD in high or normal quality
'change it to True if you have an LCD screen, 256x64
'or False if you have a real DMD at 128x32 in size
Const FlexDMDHighQuality = False

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
Const cGameName = "jpsdeadpool"
Const myVersion = "1.08"
Const MaxPlayers = 4          ' from 1 to 4
Const BallSaverTime = 20      ' in seconds of the first ball
Const MaxMultiplier = 5       ' limit playfield multiplier
Const MaxBonusMultiplier = 50 'limit Bonus multiplier
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
Dim bMusicOn
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bJackpot

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger during multiballs

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
LoadLUT
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
        .InitExitSnd SoundFX("Popper", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
        .CreateEvents "plungerIM"
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

    if bFreePlay Then DOF 125, DOFOn

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
    bJustStarted = True
    bJackpot = False
    bInstantInfo = False
    ' set any lights for the attract mode
    GiOff
    StartAttractMode
pupevent 800

    ' Start the RealTime timer
    RealTime.Enabled = 1

    ' Load table color
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
'LUT controls
If keycode = LeftMagnaSave Then bLutActive = True
    If keycode = RightMagnaSave Then
            If bLutActive Then
                    if DisableLUTSelector = 0 then
                LUTSet = LUTSet  - 1
                if LutSet < 0 then LUTSet = 16
                SetLUT
                ShowLUT
            End If
        End If
        End If


        If keycode = RightMagnaSave or keycode = LockBarKey Then
            BoomHit
        End If

If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft() :lilshake 5
If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight() :lilshake 5
If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter() :lilshake 6


    If Keycode = AddCreditKey Then
PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
        Credits = Credits + 1
        if bFreePlay = False Then DOF 125, DOFOn
        If(Tilted = False)Then
            DMDFlush
            DMD "_", CL(1, "CREDITS " & Credits), "", eNone, eNone, eNone, 500, True, ""
            If NOT bGameInPlay Then ShowTableInfo
        End If
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
        SoundPlungerPull()
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

        If keycode = LeftFlipperKey Then SolLFlipper 1 : FlipperActivate LeftFlipper, RFPress:InstantInfoTimer.Enabled = True:RotateLaneLights 1:UpdateGates 1
        If keycode = RightFlipperKey Then SolRFlipper 1 :FlipperActivate RightFlipper, RFPress:InstantInfoTimer.Enabled = True:RotateLaneLights 0

        If bChooseBattle Then
            ChooseBattle(keycode)
            Exit Sub
        End If

        If keycode = StartGameKey Then
            If((PlayersPlayingGame <MaxPlayers)AND(bOnTheFirstBall = True))Then

                If(bFreePlay = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, True, ""
                Else
                    If(Credits> 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
                        DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, True, ""
                        If Credits <1 And bFreePlay = False Then DOF 125, DOFOff
                        Else
                            ' Not Enough Credits to start a game.
                            DMD CL(0, "CREDITS " & Credits), CL(1, "INSERT COIN"), "", eNone, eBlink, eNone, 1000, True, "vo_nocredits"
                    End If
                End If
            End If
        End If
    Else ' If (GameInPlay)

        If keycode = StartGameKey Then
            If(bFreePlay = True)Then
                If(BallsOnPlayfield = 0)Then
                    ResetForNewGame()
          pupevent 801
                End If
            Else
                If(Credits> 0)Then
                    If(BallsOnPlayfield = 0)Then
                        Credits = Credits - 1
                        If Credits <1 And bFreePlay = False Then DOF 125, DOFOff
                        ResetForNewGame()
                    End If
                Else
                    ' Not Enough Credits to start a game.
                    DMDFlush
                    DMD CL(0, "CREDITS " & Credits), CL(1, "INSERT COIN"), "", eNone, eBlink, eNone, 1000, True, "vo_nocredits" : pupevent 802
                    ShowTableInfo
                End If
            End If
        End If
    End If ' If (GameInPlay)
End Sub

Sub Table1_KeyUp(ByVal keycode)

If keycode = LeftMagnaSave Then bLutActive = False

    If keycode = PlungerKey Then
        Plunger.Fire
        SoundPlungerReleaseBall()
    End If

    If hsbModeActive Then
        Exit Sub
    End If

    ' Table specific

    If bGameInPLay AND NOT Tilted Then

'LUT controls
If keycode = LeftMagnaSave Then bLutActive = False
        If keycode = LeftFlipperKey Then
SolLFlipper 0
            FlipperDeActivate LeftFlipper, LFPress
            UpdateGates 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
            End If
        End If
        If keycode = RightFlipperKey Then
SolRFlipper 0
            FlipperDeActivate RightFlipper, RFPress
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
    SaveLUT
    Savehs
  pupflasher3.VideoCapUpdate=""
  pupflasher6.VideoCapUpdate=""
  pupflasher7.VideoCapUpdate=""
    If UseFlexDMD Then FlexDMD.Run = False
    If B2SOn = true Then Controller.Stop
End Sub


'********************
'     Flippers
'********************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend
 DOF  101, DOFOn
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
  DOF  101, DOFOff
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire 'rightflipper.rotatetoend
 DOF  102, DOFOn
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
  DOF  102, DOFOff
    RightFlipper.RotateToStart
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
Sub FlipperVisualUpdate
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -2.7
'        AddPt "Polarity", 2, 0.33, -2.7
'        AddPt "Polarity", 3, 0.37, -2.7
'        AddPt "Polarity", 4, 0.41, -2.7
'        AddPt "Polarity", 5, 0.45, -2.7
'        AddPt "Polarity", 6, 0.576,-2.7
'        AddPt "Polarity", 7, 0.66, -1.8
'        AddPt "Polarity", 8, 0.743, -0.5
'        AddPt "Polarity", 9, 0.81, -0.5
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -3.7
'        AddPt "Polarity", 2, 0.33, -3.7
'        AddPt "Polarity", 3, 0.37, -3.7
'        AddPt "Polarity", 4, 0.41, -3.7
'        AddPt "Polarity", 5, 0.45, -3.7
'        AddPt "Polarity", 6, 0.576,-3.7
'        AddPt "Polarity", 7, 0.66, -2.3
'        AddPt "Polarity", 8, 0.743, -1.5
'        AddPt "Polarity", 9, 0.81, -1
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'


'*******************************************
'  Late 80's early 90's

'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'   x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
' Next
'
' AddPt "Polarity", 0, 0, 0
' AddPt "Polarity", 1, 0.05, -5
' AddPt "Polarity", 2, 0.4, -5
' AddPt "Polarity", 3, 0.6, -4.5
' AddPt "Polarity", 4, 0.65, -4.0
' AddPt "Polarity", 5, 0.7, -3.5
' AddPt "Polarity", 6, 0.75, -3.0
' AddPt "Polarity", 7, 0.8, -2.5
' AddPt "Polarity", 8, 0.85, -2.0
' AddPt "Polarity", 9, 0.9,-1.5
' AddPt "Polarity", 10, 0.95, -1.0
' AddPt "Polarity", 11, 1, -0.5
' AddPt "Polarity", 12, 1.1, 0
' AddPt "Polarity", 13, 1.3, 0
'
' addpt "Velocity", 0, 0,         1
' addpt "Velocity", 1, 0.16, 1.06
' addpt "Velocity", 2, 0.41,         1.05
' addpt "Velocity", 3, 0.53,         1'0.982
' addpt "Velocity", 4, 0.702, 0.968
' addpt "Velocity", 5, 0.95,  0.968
' addpt "Velocity", 6, 1.03,         0.945
'
' LF.Object = LeftFlipper
' LF.EndPoint = EndPointLp
' RF.Object = RightFlipper
' RF.EndPoint = EndPointRp
'End Sub



'
''*******************************************
'' Early 90's and after
'
Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 60
        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -5.5
        AddPt "Polarity", 2, 0.4, -5.5
        AddPt "Polarity", 3, 0.6, -5.0
        AddPt "Polarity", 4, 0.65, -4.5
        AddPt "Polarity", 5, 0.7, -4.0
        AddPt "Polarity", 6, 0.75, -3.5
        AddPt "Polarity", 7, 0.8, -3.0
        AddPt "Polarity", 8, 0.85, -2.5
        AddPt "Polarity", 9, 0.9,-2.0
        AddPt "Polarity", 10, 0.95, -1.5
        AddPt "Polarity", 11, 1, -1.0
        AddPt "Polarity", 12, 1.05, -0.5
        AddPt "Polarity", 13, 1.1, 0
        AddPt "Polarity", 14, 1.3, 0

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


' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub




'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
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
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
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
'  FLIPPER TRICKS
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
  Dim b, BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        end If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
  End If
End Sub

'*****************
' Maths
'*****************
Dim PI: PI = 4*Atn(1)

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
'  Check ball distance from Flipper for Rem
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
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
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

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b, BOT
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
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
  End If
End Sub

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                  'Called when table is nudged
    Tilt = Tilt + TiltSensitivity              'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity)AND(Tilt <15)Then 'show a warning
        DMD "_", CL(1, "CAREFUL"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        DMDFlush
        DMD "", "", "d_TILT", eNone, eNone, eBlink, 200, False, "" : pupevent 803
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
        DOF 101, DOFOff
        RightFlipper.RotateToStart
        DOF 102, DOFOff
        Bumper1.Threshold = 100
        Bumper2.Threshold = 100
        Bumper3.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        Bumper1.Threshold = 1
        Bumper2.Threshold = 1
        Bumper3.Threshold = 1
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
    Dim a
    Select Case Battle(CurrentPlayer, 0)
        Case 0 'no battle active so play the standard songs or the multiball songs
            If bBerserker Then
                PLaySong "mu_beserkerrage"
            ElseIf bDiscoLoopsEnabled Then
                PlaySong "mu_discoloops"
            ElseIf bDiscoMBEnabled Then
                PlaySong "mu_discomb"
            ElseIf bMechSuitMBStarted Then
                PlaySong "mu_mechsuitmb"
            ElseIf bNinjaMB Then
                PlaySong "mu_ninjamb"
            Else 'play default music
                PlaySong "mu_hellhouse" &Balls
            End If
        Case 1 'Juggernaut
            PlaySong "mu_juggernaut"
        Case 2 'Mystique
            PlaySong "mu_mystique"
        Case 3 'Sabretooth
            PlaySong "mu_sabretooth"
        Case 4 'T-Rex
            PlaySong "mu_trex"
        Case 5 'Megalodon
            PlaySong "mu_megalodon"
        Case 6 'Sauron
            PlaySong "mu_sauronmb"
        Case 7 'MrSinister
            PlaySong "mu_mrsinister"
    End Select
End Sub

'********************
' Play random quotes
'********************

Sub PlayQuote
    Dim tmp
    tmp = RndNbr(42)
    PlaySound "quote" &tmp
End Sub

Sub PlaySfx
    Dim tmp
    tmp = RndNbr(39)
    PlaySound "sfx" &tmp
    FlashForms flasher007, 1500, 50, 0
    DOF 130, DOFPulse
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
        If UBound(tmp) = -1 Then '-1 means no balls, 0 is the first captive ball, 1 is the second captive ball...)
            GiOff                ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

dim gilvl:gilvl = 1

Sub GiOn
SidewallGIONL.sidematerial= "Plastic"
SidewallGIONR.sidematerial= "Plastic"
SidewallGIOFFL.sidematerial= "colormaxnoreflectionquarter"
SidewallGIOFFR.sidematerial= "colormaxnoreflectionquarter"

Ramp003.image= "LiRamp"
Ramp004.image= "LiRamp4"

Ramp010.image= "LiRamp4"
Ramp009.image= "LiRamp"

apron.image= "plastics2"
Wall035.image= "plastics2"
Wall034.image= "plastics2"
Wall032.image= "plastics2"
Wall033.image= "plastics2"
Wall059.image= "plastics2"
bumpertop004.image= "plastics2"
bumpertop003.image= "plastics2"
bumpertop002.image= "plastics2"
bumpertop005.image= "plastics2"
bumpertop001.image= "plastics2"
bumpertop006.image= "plastics2"

Backwall2On.visible =1
Backwall2Off.visible =0

BackwallOn.visible =1
BackwallOff.visible =0

GIShadows.visible = 1

Ramp028.image = "Trucktop_On"
Ramp027.image = "Truck2_On"

    gilvl = 1
    Sound_GI_Relay 1, li036 'about the center of the table
    DOF 118, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
    Fi001.Visible = 1
    Fi002.Visible = 1
    Fi003.Visible = 1
    Fi004.Visible = 1
    Fi005.Visible = 1
    Fi006.Visible = 1
End Sub

Sub GiOff
SidewallGIOFFL.sidematerial= "Plastic"
SidewallGIOFFR.sidematerial= "Plastic"
SidewallGIONL.sidematerial= "colormaxnoreflectionquarter"
SidewallGIONR.sidematerial= "colormaxnoreflectionquarter"


Ramp003.image= "LiRampOff "
Ramp004.image= "LiRamp4Off"

Ramp010.image= "LiRamp4Off"
Ramp009.image= "LiRampOff "

apron.image= "plastics2_Off"
Wall035.image= "plastics2_Off"
Wall034.image= "plastics2_Off"
Wall032.image= "plastics2_Off"
Wall033.image= "plastics2_Off"
Wall059.image= "plastics2_Off"
bumpertop004.image= "plastics2_Off"
bumpertop003.image= "plastics2_Off"
bumpertop002.image= "plastics2_Off"
bumpertop005.image= "plastics2_Off"
bumpertop001.image= "plastics2_Off"
bumpertop006.image= "plastics2_Off"

Backwall2On.visible =0
Backwall2Off.visible =1

BackwallOn.visible =0
BackwallOff.visible =1

GIShadows.visible = 0

Ramp028.image = "Trucktop_Off"
Ramp027.image = "Truck2_Off"

    gilvl = 0
    Sound_GI_Relay 1, li036 'about the center of the table
    DOF 118, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    Fi001.Visible = 0
    Fi002.Visible = 0
    Fi003.Visible = 0
    Fi004.Visible = 0
    Fi005.Visible = 0
    Fi006.Visible = 0
End Sub

' GI, light & flashers sequence effects

Sub GiEffect(n)
    Dim ii
    Select Case n
        Case 0 'all off
            LightSeqGi.Play SeqAlloff
        Case 1 'all blink
            LightSeqGi.UpdateInterval = 20
            LightSeqGi.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqGi.UpdateInterval = 20
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
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 10, 10
        Case 4 'center fra lilDP
            LightSeqlilDP.UpdateInterval = 4
            LightSeqlilDP.Play SeqCircleOutOn, 15, 2
        Case 5 'top down
            LightSeqPlayfield.UpdateInterval = 4
            LightSeqPlayfield.Play SeqDownOn, 15, 2
        Case 6 'down to top
            LightSeqPlayfield.UpdateInterval = 4
            LightSeqPlayfield.Play SeqUpOn, 15, 1
    End Select
End Sub

Sub FlashEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqFlashers.Play SeqAlloff
        Case 1 'all blink
            LightSeqFlashers.UpdateInterval = 40
            LightSeqFlashers.Play SeqBlinking, , 15, 30
        Case 2 'random
            LightSeqFlashers.UpdateInterval = 30
            LightSeqFlashers.Play SeqRandom, 25, , 1500
        Case 3 'all blink fast
            LightSeqFlashers.UpdateInterval = 20
            LightSeqFlashers.Play SeqBlinking, , 10, 30
        Case 4 'center
            LightSeqFlashers.UpdateInterval = 4
            LightSeqFlashers.Play SeqCircleOutOn, 15, 2
        Case 5 'top down
            LightSeqFlashers.UpdateInterval = 4
            LightSeqFlashers.Play SeqDownOn, 15, 2
        Case 6 'down to top
            LightSeqFlashers.UpdateInterval = 4
            LightSeqFlashers.Play SeqUpOn, 15, 1
    End Select
End Sub

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

    ' kick it out..
DOF  123,DOFPulse
    RandomSoundBallRelease Ballrelease
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield> 1 Then
        DOF 143, DOFPulse
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
    AwardPoints = 0 : pupevent 805
    TotalBonus = 10 'yes 10 points :)
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False

    ' only process any of this if the table is not tilted.
    '(the tilt recovery mechanism will handle any extra balls or end of game)

    If NOT Tilted Then
        PLaySong "mu_bonuscount"
        'Count the bonus. This table uses several bonus
        DMD CL(0, "BONUS"), "", "", eBlink, eNone, eNone, 1000, True, ""

        'Ninja Stars
        AwardPoints = NinjaStars(CurrentPlayer) * 100000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL(0, FormatScore(AwardPoints)), CL(1, "NINJA STARS " & NinjaStars(CurrentPlayer)), "", eBlink, eNone, eNone, 1000, True, ""

        'Chimichangas
        AwardPoints = Chimichangas(CurrentPlayer) * 250000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL(0, FormatScore(AwardPoints)), CL(1, "CHIMICHANGAS " & Chimichangas(CurrentPlayer)), "", eBlink, eNone, eNone, 1000, True, ""

        'Chimichangas
        AwardPoints = Weapons(CurrentPlayer) * 100000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL(0, FormatScore(AwardPoints)), CL(1, "WEAPONS " & Weapons(CurrentPlayer)), "", eBlink, eNone, eNone, 1000, True, ""

        ' calculate the totalbonus
        DMD CL(0, FormatScore(TotalBonus)), CL(1, "TOTAL BONUS " & " X" & BonusMultiplier(CurrentPlayer)), "", eNone, eNone, eNone, 1500, True, ""
        TotalBonus = TotalBonus * BonusMultiplier(CurrentPlayer)
        ' Add the bonus to the score
        AddScore TotalBonus

        ' add a bit of a delay to allow for the bonus points to be shown & added up
        vpmtimer.addtimer 8000, "EndOfBall2 '"
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
    Tilted = False
  pupevent 806
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

    ' has the player won an extra-ball ? (might be multiple outstanding)
    If(ExtraBallsAwards(CurrentPlayer) <> 0)Then
        'debug.print "Extra Ball"

        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            LightShootAgain.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
        DMD CL(0, "EXTRA BALL"), CL(1, "SHOOT AGAIN"), "", eNone, eNone, eBlink, 1000, True, ""

        ' In this table an extra ball will have the skillshot and ball saver, so we reset the playfield for the new ball
        ResetForNewPlayerBall()

        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0)Then
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
    If(PlayersPlayingGame> 1)Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer> PlayersPlayingGame)Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

        ' set the machine into game over mode
        EndOfGame() : pupevent 808

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
            Select Case CurrentPlayer
                Case 1:DMD "", "", "d_player1", eNone, eNone, eNone, 1000, True, "vo_player1" : pupevent 809
                Case 2:DMD "", "", "d_player2", eNone, eNone, eNone, 1000, True, "vo_player2" : pupevent 810
                Case 3:DMD "", "", "d_player3", eNone, eNone, eNone, 1000, True, "vo_player3" : pupevent 811
                Case 4:DMD "", "", "d_player4", eNone, eNone, eNone, 1000, True, "vo_player4" : pupevent 812
            End Select
        Else
            DMD "", "", "d_player1", eNone, eNone, eNone, 1000, True, ""
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

    End If
    bJustStarted = False
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball

    ' set any lights for the attract mode
    GiOff
    StartAttractMode
pupevent 880
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
    If bGameInPLay = False Then Exit Sub 'don't do anything, just delete the ball
    ' Exit Sub ' only for debugging - this way you can add balls from the debug window

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' pretend to knock the ball into the ball storage mech
    RandomSoundDrain Drain
    'if Tilted the end Ball Mode
    If Tilted Then
        StopEndOfBallMode
    End If

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True)AND(Tilted = False)Then

        ' is the ball saver active,
        If(bBallSaverActive = True)Then

            ' yep, create a new ball in the shooters lane
            ' we use the Addmultiball in case the multiballs are being ejected
            AddMultiball 1
            ' we kick the ball with the autoplunger
            bAutoPlunger = True
            ' you may wish to put something on a display or play a sound at this point
            DMD "_", CL(1, "BALL SAVED"), "_", eNone, eBlinkfast, eNone, 1000, True, "" : pupevent 843
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1)Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True)then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ' you may wish to change any music over at this point and

                    ' turn off any multiball specific lights
                    ChangeGi white
                    'stop any multiball modes
                    StopMBmodes
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0)Then
                ' End Mode and timers

                ChangeGi white
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
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
    'LaunchLight.State = 2

    'be sure to update the Scoreboard after the animations, if any

    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        vpmtimer.addtimer 1500, "DOF 121, DOFPulse:PlungerIM.AutoFire:DOF 120, DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), swPlungerRest:bAutoPlunger = False '"
    End If
    'Start the Selection of the skillshot if ready
    If bSkillShotReady Then
        PlaySong "mu_shooterlane" : pupevent 813
        UpdateSkillshot()
        ' show the message to shoot the ball in case the player has fallen sleep
        SwPlungerCount = 0
        swPlungerRest.TimerEnabled = 1
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    lighteffect 6
    bBallInPlungerLane = False : pupevent 814
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
        ChangeSong
        ResetSkillShotTimer.Enabled = 1
    End If
    If bLilDPMB AND gate3.open = False Then 'if in LilDP MultiBall then open the left top gate if it was closed
        gate3.open = True
        vpmtimer.addtimer 2000, "gate3.open = False '"
    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True)AND(BallSaverTime <> 0)And(bBallSaverActive = False)Then
        EnableBallSaver BallSaverTime
    End If
' turn off LaunchLight
' LaunchLight.State = 0
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds
Dim SwPlungerCount
Sub swPlungerRest_Timer
    Select Case SwPlungerCount
        Case 0
            If bInfoNeeded1(CurrentPLayer)Then
                DMD "USE FLIPPER BUTTONS", "TO SELECT SKILLSHOT", "_", eNone, eNone, eNone, 2000, True, "vo_useflipperselecttoplaneskillshot" : pupevent 815
            End If
        Case 1
            If bInfoNeeded1(CurrentPLayer)Then
                DMD "HOLD LEFT FLIPPER", "FOR SUPERSKILLSHOT", "_", eNone, eNone, eNone, 2000, True, "vo_holdflipper-sk" : pupevent 816
                bInfoNeeded1(CurrentPLayer) = False
            End If
        Case 4
            DMD CL(0, "HEY YOU"), CL(1, "SHOOT THE BALL"), "_", eNone, eNone, eNone, 2000, True, "vo_whatareyouwaitingfor"
            swPlungerRest.TimerEnabled = 0
    End Select
    SwPlungerCount = SwPlungerCount + 1
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
    If Tilted Then Exit Sub
    ' add the points to the current players score variable
    If Battle(CurrentPLayer, 0)> 0 Then
        Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer) * ColossusValue
    Else
        Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer)
    End if
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

    ' If(bMultiBallMode = True) Then
    Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + points
    DMD "_", CL(1, "INCREASED JACKPOT"), "_", eNone, eNone, eNone, 1000, True, ""
' you may wish to limit the jackpot to a upper limit, ie..
' If (Jackpot >= 6000) Then
'   Jackpot = 6000
'   End if
'End if
End Sub

Sub AddSuperJackpot(points) 'not used in this table
    If Tilted Then Exit Sub
End Sub

Sub AddBonusMultiplier(n)
    Dim NewBonusLevel
    ' if not at the maximum bonus level
    if(BonusMultiplier(CurrentPlayer) + n <= MaxBonusMultiplier)then
        ' then add and set the lights
        NewBonusLevel = BonusMultiplier(CurrentPlayer) + n
        SetBonusMultiplier(NewBonusLevel)
        'DMD "_", CL(1, "BONUS X " &NewBonusLevel), "_", eNone, eBlink, eNone, 2000, True, ""
    Else
        AddScore 50000
        DMD "_", CL(1, "50000"), "_", eNone, eNone, eNone, 1000, True, ""
    End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly

Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level
    UPdateBonusXLights(Level)
End Sub

Sub UpdateBonusXLights(Level) 'no lights in this table
    ' Update the lights
    Select Case Level
    'Case 1:light54.State = 0:light55.State = 0:light56.State = 0:light57.State = 0
    'Case 2:light54.State = 1:light55.State = 0:light56.State = 0:light57.State = 0
    'Case 3:light54.State = 0:light55.State = 1:light56.State = 0:light57.State = 0
    'Case 4:light54.State = 0:light55.State = 0:light56.State = 1:light57.State = 0
    'Case 5:light54.State = 0:light55.State = 0:light56.State = 0:light57.State = 1
    End Select
End Sub

Sub AddPlayfieldMultiplier(n)
    Dim snd
    Dim NewPFLevel
    ' if not at the maximum level x
    if(PlayfieldMultiplier(CurrentPlayer) + n <= MaxMultiplier)then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) + n
        SetPlayfieldMultiplier(NewPFLevel)
        snd = "sfx_thunder" &NewPFLevel
        DMD "_", CL(1, "PLAYFIELD X " &NewPFLevel), "_", eNone, eBlink, eNone, 2000, True, snd
    Else 'if the 5x is already lit
        AddScore 50000
        DMD "_", CL(1, "50000"), "_", eNone, eNone, eNone, 2000, True, ""
    End if
'Start the timer to reduce the playfield x every 30 seconds
pfxtimer.Enabled = 0
pfxtimer.Enabled = 1
End Sub

Sub PFXTimer_Timer
    DecreasePlayfieldMultiplier
End Sub

Sub DecreasePlayfieldMultiplier 'reduces by 1 the playfield multiplier
    Dim NewPFLevel
    ' if not at 1 already
    if(PlayfieldMultiplier(CurrentPlayer) > 1) then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) - 1
        SetPlayfieldMultiplier(NewPFLevel)
        PFXTimer.Enabled = 1
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

Sub UpdatePFXLights(Level)
    ' Update the playfield multiplier lights
    Select Case Level
        Case 1:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0
        Case 2:li002.State = 1:li003.State = 0:li004.State = 0:li005.State = 0
        Case 3:li002.State = 1:li003.State = 1:li004.State = 0:li005.State = 0
        Case 4:li002.State = 1:li003.State = 1:li004.State = 1:li005.State = 0
        Case 5:li002.State = 1:li003.State = 1:li004.State = 1:li005.State = 1
    End Select
' show the multiplier in the DMD?
End Sub

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        DMD "_", CL(1, ("EXTRA BALL WON")), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("Knocker", 122, DOFPulse, DOFKnocker) : pupevent 807
        DOF 121, DOFPulse
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        LightShootAgain.State = 1 'light the shoot again lamp
        GiEffect 2
        LightEffect 2
    END If
End Sub

Sub AwardSpecial()
    DMD "_", CL(1, ("EXTRA GAME WON")), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("Knocker", 122, DOFPulse, DOFKnocker)
    DOF 121, DOFPulse
    Credits = Credits + 1
    If bFreePlay = False Then DOF 125, DOFOn
    LightEffect 2
    GiEffect 2
End Sub

Sub AwardJackpot() 'award a normal jackpot
    If bLilDPMB Then
        Select case RndNbr(4)
            Case 1:DMD CL(0, FormatScore(Jackpot(CurrentPlayer))), CL(1, "LILDP JACKPOT"), "bkborder", eBlinkFast, eBlinkFast, eNone, 1000, True, "vo_lildp_jackpot1" : pupevent 807
            Case 2:DMD CL(0, FormatScore(Jackpot(CurrentPlayer))), CL(1, "LILDP JACKPOT"), "bkborder", eBlinkFast, eBlinkFast, eNone, 1000, True, "vo_lildp_jackpot2" : pupevent 807
            Case 3:DMD CL(0, FormatScore(Jackpot(CurrentPlayer))), CL(1, "LILDP JACKPOT"), "bkborder", eBlinkFast, eBlinkFast, eNone, 1000, True, "vo_lildp_jackpot3" : pupevent 807
            Case 4:DMD CL(0, FormatScore(Jackpot(CurrentPlayer))), CL(1, "LILDP JACKPOT"), "bkborder", eBlinkFast, eBlinkFast, eNone, 1000, True, "vo_lildp_jackpot4" : pupevent 807
        End Select
        LightEffect 4
        GiEffect 4
    Else
        DMD CL(0, FormatScore(Jackpot(CurrentPlayer))), CL(1, "JACKPOT"), "bkborder", eBlinkFast, eBlinkFast, eNone, 1000, True, "vo_Jackpot" : pupevent 807
        LightEffect 2
        GiEffect 2
    End If
    DOF 126, DOFPulse
    AddScore Jackpot(CurrentPlayer)
End Sub

Sub AwardMegaJackpot() 'award a Megalodon Jackpot
    DMD CL(0, FormatScore(Jackpot(CurrentPlayer))), CL(1, "MEGALODON JACKPOT"), "d_megalodon", eBlinkFast, eBlinkFast, eNone, 1000, True, "vo_Jackpot" : pupevent 818
    DOF 126, DOFPulse
    AddScore MegaJackpot(CurrentPlayer)
    MegaJackpot(CurrentPlayer) = MegaJackpot(CurrentPlayer) + 10000
    MegaSuperJCount = MegaSuperJCount + 1
    LightEffect 2
    GiEffect 2
End Sub

Sub AwardSuperMegaJackpot() 'award a SuperMegalodon Jackpot
    MegaSuperJValue = MegaJackpot(CurrentPlayer) * MegaSuperJCount
    DMD CL(0, FormatScore(MegaSuperJValue)), "MEGALODON SUPJACKPOT", "d_megalodon", eBlinkFast, eBlinkFast, eNone, 1000, True, "vo_superjackpot" : pupevent 819
    DOF 126, DOFPulse
    AddScore MegaSuperJValue
    LightEffect 2
    GiEffect 2
End Sub

Sub AwardSuperJackpot()
    DMD CL(0, FormatScore(SuperJackpot(CurrentPlayer))), CL(1, "SUPER JACKPOT"), "bkborder", eBlinkFast, eBlinkFast, eNone, 1000, True, "vo_superjackpot" : pupevent 820
    DOF 126, DOFPulse
    AddScore SuperJackpot(CurrentPlayer)
    LightEffect 2
    GiEffect 2
End Sub

Sub AwardSkillshot()
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL(0, FormatScore(SkillshotValue(CurrentPlayer))), CL(1, ("SKILLSHOT")), "bkborder", eBlinkFast, eBlink, eNone, 1000, True, "fanfare" : pupevent 821
    DOF 127, DOFPulse
    Addscore SkillShotValue(CurrentPlayer)
    ' increment the skillshot value with 250.000
    SkillShotValue(CurrentPlayer) = SkillShotValue(CurrentPlayer) + 250000
    'do some light show
    GiEffect 2
    LightEffect 2
End Sub

Sub AwardSuperSkillshot()
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL(0, FormatScore(SuperSkillshotValue(CurrentPlayer))), CL(1, ("SUPER SKILLSHOT")), "bkborder", eBlinkFast, eBlink, eNone, 1000, True, "fanfare" : pupevent 822
    DOF 127, DOFPulse
    Addscore SuperSkillshotValue(CurrentPlayer)
    ' increment the skillshot value with 500.000
    SuperSkillshotValue(CurrentPlayer) = SuperSkillshotValue(CurrentPlayer) + 500000
    'do some light show
    GiEffect 2
    LightEffect 2
End Sub

'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(cGameName, "HighScore1")
    If(x <> "")Then HighScore(0) = CDbl(x)Else HighScore(0) = 100000 End If
    x = LoadValue(cGameName, "HighScore1Name")
    If(x <> "")Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If
    x = LoadValue(cGameName, "HighScore2")
    If(x <> "")then HighScore(1) = CDbl(x)Else HighScore(1) = 100000 End If
    x = LoadValue(cGameName, "HighScore2Name")
    If(x <> "")then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If
    x = LoadValue(cGameName, "HighScore3")
    If(x <> "")then HighScore(2) = CDbl(x)Else HighScore(2) = 100000 End If
    x = LoadValue(cGameName, "HighScore3Name")
    If(x <> "")then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If
    x = LoadValue(cGameName, "HighScore4")
    If(x <> "")then HighScore(3) = CDbl(x)Else HighScore(3) = 100000 End If
    x = LoadValue(cGameName, "HighScore4Name")
    If(x <> "")then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If
    x = LoadValue(cGameName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0:If bFreePlay = False Then DOF 125, DOFOff:End If
    x = LoadValue(cGameName, "TotalGamesPlayed")
    If(x <> "")then TotalGamesPlayed = CInt(x)Else TotalGamesPlayed = 0 End If
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
    HighScore(0) = 150000
    HighScore(1) = 140000
    HighScore(2) = 130000
    HighScore(3) = 120000
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

    If tmp> HighScore(0)Then 'add 1 credit for beating the highscore
        Credits = Credits + 1
        DOF 125, DOFOn
    End If

    If tmp> HighScore(3)Then
        PlaySound SoundFXDOF("Knocker", 122, DOFPulse, DOFKnocker)
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
    PlaySound "vo_enterinitials" : pupevent 823
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
        if(hsCurrentLetter = 0)then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters))then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<")then
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
            if(hsCurrentDigit> 0)then
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
    if(hsCurrentDigit> 0)then TempBotStr = TempBotStr & hsEnteredDigits(0)
    if(hsCurrentDigit> 1)then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit> 2)then TempBotStr = TempBotStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3)then
        if(hsLetterFlash <> 0)then
            TempBotStr = TempBotStr & "_"
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit <1)then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit <2)then TempBotStr = TempBotStr & hsEnteredDigits(2)

    TempBotStr = TempBotStr & " <    "
    dLine(1) = ExpandLine(TempBotStr, 1)
    DMDUpdate 1
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2)then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = False

    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ")then
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
            If HighScore(j) <HighScore(j + 1)Then
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
            If FlexDMDHighQuality Then
                FlexDMD.TableFile = Table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 256
                FlexDMD.Height = 64
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.bkempty")
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
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.bkempty")
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
    Dim tmp, tmp1, tmp1a, tmp1b, tmp2
    if(dqHead = dqTail)Then
        tmp = FL(0, "P" &CurrentPlayer& "B" &Balls, FormatScore(Score(Currentplayer)))
        'tmp = CL(0, FormatScore(Score(Currentplayer) ) )
        'tmp1 = CL(1, "PLAYER " & CurrentPlayer & " BALL " & Balls)
        'tmp1 = FormatScore(Bonuspoints(Currentplayer) ) & " X" &BonusMultiplier(Currentplayer)
        tmp2 = "bkborder"
        Select Case Battle(CurrentPlayer, 0)
            Case 0:              'no battle active then display the Team-up state
                If bLilDPMB Then 'LilDP multiball is active then show the jackpot value
                    tmp = RL(0, FormatScore(Score(Currentplayer)))
                    tmp1 = RL(1, FormatScore(LilDPJackpot(CurrentPlayer)) & " P" &CurrentPlayer& "B" &Balls)
                    tmp2 = "d_lildpmb"
                Else
                    tmp1a = Chr(33)&Chr(37 + DazzlePower(CurrentPlayer))&Chr(34)&Chr(37 + ColossusPower(CurrentPlayer))&Chr(35)&Chr(37 + WolverinePower(CurrentPlayer))&Chr(36)&Chr(37 + DominoPower(CurrentPlayer))
                    tmp1b = tmp1 & chr(42) & NinjaStars(CurrentPlayer) & chr(44) & Chimichangas(CurrentPlayer) & chr(47) & Weapons(CurrentPlayer)
                    tmp1 = FL(1, tmp1a, tmp1b)
                End If
            Case 1 'Juggernaut
                tmp = RL(0, FormatScore(Score(Currentplayer)))
                tmp1a = ShowLife(LifeLeft(CurrentPlayer, 1))
                tmp1b = "P" &CurrentPlayer& "B" &Balls
                if DazzlePower(CurrentPlayer) = 4 Then tmp1b = Chr(33)&tmp1b
                if ColossusPower(CurrentPlayer) = 4 Then tmp1b = Chr(34)&tmp1b
                if WolverinePower(CurrentPlayer) = 4 Then tmp1b = Chr(35)&tmp1b
                if DominoPower(CurrentPlayer) = 4 Then tmp1b = Chr(36)&tmp1b
                tmp1 = FL(1, tmp1a, tmp1b)
                tmp2 = "d_juggernaut"
            Case 2 'Mystique
                tmp = RL(0, FormatScore(Score(Currentplayer)))
                tmp1a = ShowLife(LifeLeft(CurrentPlayer, 2))
                tmp1b = "P" &CurrentPlayer& "B" &Balls
                if DazzlePower(CurrentPlayer) = 4 Then tmp1b = Chr(33)&tmp1b
                if ColossusPower(CurrentPlayer) = 4 Then tmp1b = Chr(34)&tmp1b
                if WolverinePower(CurrentPlayer) = 4 Then tmp1b = Chr(35)&tmp1b
                if DominoPower(CurrentPlayer) = 4 Then tmp1b = Chr(36)&tmp1b
                tmp1 = FL(1, tmp1a, tmp1b)
                tmp2 = "d_mystique"
            Case 3 'Sabretooth
                tmp = RL(0, FormatScore(Score(Currentplayer)))
                tmp1a = ShowLife(LifeLeft(CurrentPlayer, 3))
                tmp1b = "P" &CurrentPlayer& "B" &Balls
                if DazzlePower(CurrentPlayer) = 4 Then tmp1b = Chr(33)&tmp1b
                if ColossusPower(CurrentPlayer) = 4 Then tmp1b = Chr(34)&tmp1b
                if WolverinePower(CurrentPlayer) = 4 Then tmp1b = Chr(35)&tmp1b
                if DominoPower(CurrentPlayer) = 4 Then tmp1b = Chr(36)&tmp1b
                tmp1 = FL(1, tmp1a, tmp1b)
                tmp2 = "d_sabretooth"
            Case 4 'T-Rex
                tmp = RL(0, FormatScore(Score(Currentplayer)))
                tmp1a = ShowLife(LifeLeft(CurrentPlayer, 4))
                tmp1b = "P" &CurrentPlayer& "B" &Balls
                if DazzlePower(CurrentPlayer) = 4 Then tmp1b = Chr(33)&tmp1b
                if ColossusPower(CurrentPlayer) = 4 Then tmp1b = Chr(34)&tmp1b
                if WolverinePower(CurrentPlayer) = 4 Then tmp1b = Chr(35)&tmp1b
                if DominoPower(CurrentPlayer) = 4 Then tmp1b = Chr(36)&tmp1b
                tmp1 = FL(1, tmp1a, tmp1b)
                tmp2 = "d_trex"
            Case 5 'Megalodon
                tmp = RL(0, FormatScore(Score(Currentplayer)))
                tmp1a = ShowLife(LifeLeft(CurrentPlayer, 5))
                tmp1b = "P" &CurrentPlayer& "B" &Balls
                if DazzlePower(CurrentPlayer) = 4 Then tmp1b = Chr(33)&tmp1b
                if ColossusPower(CurrentPlayer) = 4 Then tmp1b = Chr(34)&tmp1b
                if WolverinePower(CurrentPlayer) = 4 Then tmp1b = Chr(35)&tmp1b
                if DominoPower(CurrentPlayer) = 4 Then tmp1b = Chr(36)&tmp1b
                tmp1 = FL(1, tmp1a, tmp1b)
                tmp2 = "d_megalodon"
            Case 6 'Sauron
                tmp = RL(0, FormatScore(Score(Currentplayer)))
                tmp1a = ShowLife(LifeLeft(CurrentPlayer, 6))
                tmp1b = "P" &CurrentPlayer& "B" &Balls
                if DazzlePower(CurrentPlayer) = 4 Then tmp1b = Chr(33)&tmp1b
                if ColossusPower(CurrentPlayer) = 4 Then tmp1b = Chr(34)&tmp1b
                if WolverinePower(CurrentPlayer) = 4 Then tmp1b = Chr(35)&tmp1b
                if DominoPower(CurrentPlayer) = 4 Then tmp1b = Chr(36)&tmp1b
                tmp1 = FL(1, tmp1a, tmp1b)
                tmp2 = "d_sauron"
            Case 7 'MrSinister
                tmp = RL(0, FormatScore(Score(Currentplayer)))
                tmp1a = ShowLife(LifeLeft(CurrentPlayer, 7))
                tmp1b = "P" &CurrentPlayer& "B" &Balls
                if DazzlePower(CurrentPlayer) = 4 Then tmp1b = Chr(33)&tmp1b
                if ColossusPower(CurrentPlayer) = 4 Then tmp1b = Chr(34)&tmp1b
                if WolverinePower(CurrentPlayer) = 4 Then tmp1b = Chr(35)&tmp1b
                if DominoPower(CurrentPlayer) = 4 Then tmp1b = Chr(36)&tmp1b
                tmp1 = FL(1, tmp1a, tmp1b)
                tmp2 = "d_mrsinister"
        End Select
    End If
    DMD tmp, tmp1, tmp2, eNone, eNone, eNone, 25, True, ""
End Sub

Function Showlife(n)
    Dim tmp, a
    tmp = INT(n)
    Select Case tmp
        Case 0:a = "mnnnnnnnno"
        Case 1:a = "jnnnnnnnno"
        Case 2:a = "jknnnnnnno"
        Case 3:a = "jkknnnnnno"
        Case 4:a = "jkkknnnnno"
        Case 5:a = "jkkkknnnno"
        Case 6:a = "jkkkkknnno"
        Case 7:a = "jkkkkkknno"
        Case 8:a = "jkkkkkkkno"
        Case 9:a = "jkkkkkkkko"
        Case 10:a = "jkkkkkkkkl"
    End Select
    Showlife = a
End Function

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
    if(dqTail <dqSize)Then
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
            dqText(2, dqTail) = Text2 'it is always 1 letter in this table
        End If

        dqTimeOn(dqTail) = TimeOn
        dqbFlush(dqTail) = bFlush
        dqSound(dqTail) = Sound
        dqTail = dqTail + 1
        if(dqTail = 1)Then
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

    if(deCount(0) = deCountEnd(0))and(deCount(1) = deCountEnd(1))and(deCount(2) = deCountEnd(2))Then

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
        if(Len(TempStr)> Space(dCharsPerLine(id)))Then
            TempStr = Left(TempStr, Space(dCharsPerLine(id)))
        Else
            if(Len(TempStr) <dCharsPerLine(id))Then
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

Function FL(id, NumString1, NumString2) 'Fill line
    Dim Temp, TempStr
    Temp = dCharsPerLine(id)- Len(NumString1)- Len(NumString2)
    TempStr = NumString1 & Space(Temp) & NumString2
    FL = TempStr
End Function

Function CL(id, NumString) 'center line
    Dim Temp, TempStr
    Temp = (dCharsPerLine(id)- Len(NumString)) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(id, NumString) 'right line
    Dim Temp, TempStr
    Temp = dCharsPerLine(id)- Len(NumString)
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
    For i = 0 to 255:Chars(i) = "d_empty":Next

    Chars(32) = "d_empty"
    Chars(33) = "d_daz"    '! dazzle icon
    Chars(34) = "d_col"    '" colosus icon
    Chars(35) = "d_wol"    '# wonverine icon
    Chars(36) = "d_dom"    '$ domino icon
    Chars(37) = "d_Power0" '%
    Chars(38) = "d_Power1" '&
    Chars(39) = "d_Power2" ''
    Chars(40) = "d_Power3" '(
    Chars(41) = "d_Power4" ')
    Chars(42) = "d_star"   '* ninja star
    Chars(43) = "d_plus"   '+
    Chars(44) = "d_chimi"  ' chimichangas
    Chars(45) = "d_minus"  '-
    Chars(46) = "d_dot"    '.
    Chars(47) = "d_gun"    '/ weapons
    Chars(48) = "d_0"      '0
    Chars(49) = "d_1"      '1
    Chars(50) = "d_2"      '2
    Chars(51) = "d_3"      '3
    Chars(52) = "d_4"      '4
    Chars(53) = "d_5"      '5
    Chars(54) = "d_6"      '6
    Chars(55) = "d_7"      '7
    Chars(56) = "d_8"      '8
    Chars(57) = "d_9"      '9
    Chars(60) = "d_less"   '<
    Chars(61) = "d_equal"  '=
    Chars(62) = "d_more"   '>
    Chars(64) = "bkempty"  '@
    Chars(65) = "d_a"      'A
    Chars(66) = "d_b"      'B
    Chars(67) = "d_c"      'C
    Chars(68) = "d_d"      'D
    Chars(69) = "d_e"      'E
    Chars(70) = "d_f"      'F
    Chars(71) = "d_g"      'G
    Chars(72) = "d_h"      'H
    Chars(73) = "d_i"      'I
    Chars(74) = "d_j"      'J
    Chars(75) = "d_k"      'K
    Chars(76) = "d_l"      'L
    Chars(77) = "d_m"      'M
    Chars(78) = "d_n"      'N
    Chars(79) = "d_o"      'O
    Chars(80) = "d_p"      'P
    Chars(81) = "d_q"      'Q
    Chars(82) = "d_r"      'R
    Chars(83) = "d_s"      'S
    Chars(84) = "d_t"      'T
    Chars(85) = "d_u"      'U
    Chars(86) = "d_v"      'V
    Chars(87) = "d_w"      'W
    Chars(88) = "d_x"      'X
    Chars(89) = "d_y"      'Y
    Chars(90) = "d_z"      'Z
    Chars(94) = "d_up"     '^
    '    Chars(95) = '_
    Chars(96) = "d_0a"        '0.
    Chars(97) = "d_1a"        '1. 'a
    Chars(98) = "d_2a"        '2. 'b
    Chars(99) = "d_3a"        '3. 'c
    Chars(100) = "d_4a"       '4. 'd
    Chars(101) = "d_5a"       '5. 'e
    Chars(102) = "d_6a"       '6. 'f
    Chars(103) = "d_7a"       '7. 'g
    Chars(104) = "d_8a"       '8. 'h
    Chars(105) = "d_9a"       '9. 'i
    Chars(106) = "d_LifeLon"  'j
    Chars(107) = "d_LifeMon"  'k
    Chars(108) = "d_LifeRon"  'l
    Chars(109) = "d_LifeLoff" 'm
    Chars(110) = "d_LifeMoff" 'n
    Chars(111) = "d_LifeRoff" 'o
    Chars(112) = ""           'p
    Chars(113) = ""           'q
    Chars(114) = ""           'r
    Chars(115) = ""           's
    Chars(116) = ""           't
    Chars(117) = ""           'u
    Chars(118) = ""           'v
    Chars(119) = ""           'w
    Chars(120) = ""           'x
    Chars(121) = ""           'y
    Chars(122) = ""           'z
    Chars(123) = ""           '{
    Chars(124) = ""           '|
    Chars(125) = ""           '}
    Chars(126) = ""           '~
End Sub

'********************
' Real Time updates
'********************
'used for all the real time updates

Sub Realtime_Timer
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
' add any other real time update subs, like gates or diverters, flippers
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
' 11 colors: red, orange, amber, yellow...
'******************************************
' in this table this colors are use to keep track of the progress during the modes

'colors
Const red = 10
Const orange = 9
Const amber = 8
Const yellow = 7
Const darkgreen = 6
Const green = 5
Const blue = 4
Const darkblue = 3
Const purple = 2
Const white = 1
Const teal = 0

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
            n.color = RGB(255, 197, 143)
            n.colorfull = RGB(255, 252, 224)
        Case teal
            n.color = RGB(1, 64, 62)
            n.colorfull = RGB(2, 128, 126)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

Sub SetFlashColor(n, col, stat) 'stat 0 = off, 1 = on, -1= no change - no blink for the flashers
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
            n.color = RGB(255, 252, 224)
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
    TurnOffArrows
End Sub

Sub TurnOffArrows() 'during battles when changing modes
    For each x in aArrows
        x.State = 0
    Next
    For x = 1 to 6
        BattleLights(x) = 0
    Next
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
    If Score(1)Then
        DMD CL(0, "LAST SCORE"), CL(1, "PLAYER 1 " &FormatScore(Score(1))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(2)Then
        DMD CL(0, "LAST SCORE"), CL(1, "PLAYER 2 " &FormatScore(Score(2))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(3)Then
        DMD CL(0, "LAST SCORE"), CL(1, "PLAYER 3 " &FormatScore(Score(3))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(4)Then
        DMD CL(0, "LAST SCORE"), CL(1, "PLAYER 4 " &FormatScore(Score(4))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    DMD "", "", "d_gameover", eNone, eNone, eBlink, 2000, False, ""
    If bFreePlay Then
        DMD "", CL(1, "FREE PLAY"), "", eNone, eBlink, eNone, 2000, False, ""
    Else
        If Credits> 0 Then
            DMD CL(0, "CREDITS " & Credits), CL(1, "PRESS START"), "", eNone, eBlink, eNone, 2000, False, ""
        Else
            DMD CL(0, "CREDITS " & Credits), CL(1, "INSERT COIN"), "", eNone, eBlink, eNone, 2000, False, ""
        End If
    End If
    DMD "", "", "d_jppresents", eNone, eNone, eNone, 3000, False, ""
    DMD "", "", "d_title", eNone, eNone, eNone, 4000, False, "vo_deadpool"
    DMD "", CL(1, "ROM VERSION " &myversion), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL(0, "HIGHSCORES"), Space(dCharsPerLine(1)), "", eScrollLeft, eScrollLeft, eNone, 20, False, ""
    DMD CL(0, "HIGHSCORES"), "", "", eBlinkFast, eNone, eNone, 1000, False, ""
    DMD CL(0, "HIGHSCORES"), "1> " &HighScoreName(0) & " " &FormatScore(HighScore(0)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "2> " &HighScoreName(1) & " " &FormatScore(HighScore(1)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "3> " &HighScoreName(2) & " " &FormatScore(HighScore(2)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "4> " &HighScoreName(3) & " " &FormatScore(HighScore(3)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD Space(dCharsPerLine(0)), Space(dCharsPerLine(1)), "", eScrollLeft, eScrollLeft, eNone, 500, False, ""
End Sub

Sub StartAttractMode
    Dim a
    StartRainbow aArrows
    StartLightSeq
    DMDFlush
    ShowTableInfo

If MusicAttractMode = 1 then
    a = RndNbr(2)
    Select Case a
        Case 1:PlaySong "mu_gameover"
        case 2:PlaySong "mu_boom-dp-rap"
    End Select
End If
End Sub

Sub StopAttractMode
    StopRainbow
    DMDScoreNow
    LightSeqAttract.StopPlay
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqAttract.UpdateInterval = 25
    LightSeqAttract.Play SeqBlinking, , 5, 150
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

' droptargets, animations, timers, etc
Sub VPObjects_Init
End Sub

' tables variables and Mode init
Dim bRotateLights
Dim UpdateArrowsCount
Dim Battle(4, 7)        'the first 4 is the current player, contains status of the battle, 0 not started, 1 won, 2 started
Dim LifeLeft(4, 7)      'life left of the enemies on each battle
Dim Dead(4, 4)          'dead lights
Dim Pool(4, 4)          'pool lights
Dim Boom(4, 4)          'boom lights
Dim BoomCount
Dim WeaponLights(4, 6)  'weapon lights
Dim DiamondLights(4, 6) 'diamond lights
Dim BumperHits(4)
Dim lilPos
Dim bChooseBattle
Dim bBattleReady
Dim Battlenr
Dim WolverinePower(4)
Dim DazzlePower(4)
Dim ColossusPower(4)
Dim DominoPower(4)
Dim NinjaStars(4)   'bonus, collected Stars
Dim Chimichangas(4) 'bonus. collected chimichangas
Dim Weapons(4)      'bonus, collected weapons
Dim OldCombo
Dim ComboCount
Dim ComboValue
Dim AttackPower
Dim WolverineValue
Dim bEndBattleJackpot
Dim bFirstBattle(4)
Dim BattlesWon(4)
Dim BattlesToChoose
Dim MegaJackpot(4)
Dim MegaSuperJCount
Dim MegaSuperJValue
Dim bLockEnabled
Dim DominoValue
Dim ColossusValue
Dim DazzlerValue
Dim bLilDPMB
Dim LilDPHitsNeeded(4)
Dim LilDPHits
Dim LilDPJackpot(4)
Dim LeftSpinnerCount(4)
Dim bDiscoEnabled
Dim bDiscoLoopsEnabled
Dim DiscoLoopsValue(4)
Dim bDiscoMBEnabled
Dim DiscoSuperJackpot(4)
Dim NinjaMBJackpot(4)
Dim NinjaMBCount         'count the Ninja Multiball jackpots to calculate the super jackpot
Dim NinjaApocJackpot(4)
Dim BattleLights(6)      'Arrows
Dim StarLights(6)        'Arrows
Dim ChimiChangaLights(6) 'Arrows
Dim LilDPLights(6)       'Arrows
Dim NinjaMBLights(6)     'Arrows
Dim MechSuitLights(6)    'Arrows
Dim bNinjaMB
Dim bNinjaMBSJackpot
Dim bMechSuitMBLight(4)
Dim bMechSuitMBStarted
Dim MechMBJackpot(4)
Dim bMechSuitMBSJackpot
Dim ChimichangaHits(4)
Dim ChimichangaHitsNeeded(4)
Dim ChimichangaValue(4)
Dim bBerserker
Dim BerserkerCount(4)
Dim BerserkerValue(4)
Dim bKatanaramaTime
Dim KatanaramaValue(4)
Dim KatanaramaCount(4)
Dim bInfoNeeded1(4) 'skillshot
Dim bInfoNeeded2(4) 'dead
Dim bInfoNeeded3(4) 'pool
Dim bInfoNeeded4(4) 'deadpool mystery
Dim bInfoNeeded5(4) 'deadpool regenerate
Dim bInfoNeeded6(4) 'top naes bonus x
Dim bBallLilDPLocked

Sub Game_Init()     'called at the start of a new game
    Dim i, j
    bExtraBallWonThisBall = False
    TurnOffPlayfieldLights()
    'Play some Music

    'Init Variables
    bRotateLights = True
    UpdateArrowsCount = 0
    lilPos = 0
    Battlenr = 1
    bChooseBattle = False
    bEndBattleJackpot = False
    BattlesToChoose = 3
    MegaSuperJCount = 0
    BoomCount = 0
    OldCombo = ""
    ComboCount = 0
    ComboValue = 500000
    AttackPower = 1
    WolverineValue = 1
    bBattleReady = True
    bLockEnabled = True:SwordEffect 1
    DominoValue = 0
    ColossusValue = 1
    DazzlerValue = 0
    bLilDPMB = False 'True during LilDP multiball
    bDiscoEnabled = False
    bDiscoLoopsEnabled = False
    bDiscoMBEnabled = False
    bNinjaMB = False
    bNinjaMBSJackpot = False
    bMechSuitMBStarted = False
    bMechSuitMBSJackpot = False
    bBerserker = False
    bKatanaramaTime = False
    bBallLilDPLocked = False
    For i = 0 to 4
        BonusMultiplier(i) = 1
        BumperHits(i) = 0
        WolverinePower(i) = 2
        DazzlePower(i) = 2
        ColossusPower(i) = 2
        DominoPower(i) = 2
        BallsInLock(i) = 0 'current player
        SkillshotValue(i) = 250000
        SuperSkillshotValue(i) = 5000000
        PlayfieldMultiplier(i) = 1
        NinjaStars(i) = 0
        Chimichangas(i) = 0
        Weapons(i) = 0
        bFirstBattle(i) = True
        BattlesWon(i) = 0
        MegaJackpot(i) = 50000
        LilDPHitsNeeded(i) = 1
        LilDPJackpot(i) = 500000
        LeftSpinnerCount(i) = 0
        DiscoLoopsValue(i) = 350000
        DiscoSuperJackpot(i) = 500000
        NinjaMBJackpot(i) = 500000
        NinjaApocJackpot(i) = 5000000
        bMechSuitMBLight(i) = False
        MechMBJackpot(i) = 5000000
        ChimichangaHits(i) = 0
        ChimichangaHitsNeeded(i) = 1
        ChimichangaValue(i) = 500000
        BerserkerCount(i) = 0
        BerserkerValue(i) = 50000
        KatanaramaValue(i) = 1000000
        KatanaramaCount(i) = 0
        bInfoNeeded1(i) = True
        bInfoNeeded2(i) = True
        bInfoNeeded3(i) = True
        bInfoNeeded4(i) = True
        bInfoNeeded5(i) = True
        bInfoNeeded6(i) = True
    Next
    For i = 0 to 4
        For j = 0 to 7
            Battle(i, j) = 0
            LifeLeft(i, j) = 10
        Next
    Next
    For i = 0 to 4
        For j = 0 to 4
            Dead(i, j) = 0
            Pool(i, j) = 0
            Boom(i, j) = 0
        Next
    Next
    For i = 0 to 4
        For j = 0 to 6
            DiamondLights(i, j) = 0
            WeaponLights(i, j) = 1
        Next
    Next
    For i = 0 to 6
        BattleLights(i) = 0
        StarLights(i) = 0
        ChimiChangaLights(i) = 0
        LilDPLights(i) = 0
        NinjaMBLights(i) = 0
        MechSuitLights(i) = 0
    Next
    'Init Delays/Timers
    DiamondLights(1, 1) = 2 'enable the first Diamond light
    DiamondLights(2, 1) = 2
    DiamondLights(3, 1) = 2
    DiamondLights(4, 1) = 2
    UpdateArrows.Enabled = 1
'MainMode Init()
End Sub

Sub InstantInfo
    Dim tmp
    DMD CL(0, "INSTANT INFO"), "", "", eNone, eNone, eNone, 1000, False, ""
    If Score(1)Then
        DMD CL(0, "PLAYER 1 SCORE"), CL(1, FormatScore(Score(1))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(2)Then
        DMD CL(0, "PLAYER 2 SCORE"), CL(1, FormatScore(Score(2))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(3)Then
        DMD CL(0, "PLAYER 3 SCORE"), CL(1, FormatScore(Score(3))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(4)Then
        DMD CL(0, "PLAYER 4 SCORE"), CL(1, FormatScore(Score(4))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    DMD CL(0, "COLLECTED"), CL(1, NinjaStars(CurrentPlayer) & " NINJA STARS"), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "COLLECTED"), CL(1, Chimichangas(CurrentPlayer) & " CHIMICHANGAS"), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "COLLECTED"), CL(1, Weapons(CurrentPlayer) & " WEAPONS"), "", eNone, eNone, eNone, 1000, False, ""
    tmp = 25 - NinjaStars(CurrentPlayer)MOD 25
    DMD CL(0, "COLLECT " &tmp& " STARS"), CL(1, "TO INCREASE PFX"), "", eNone, eNone, eNone, 1000, False, ""
    tmp = 50 - NinjaStars(CurrentPlayer)MOD 50
    DMD CL(0, "COLLECT " &tmp& " STARS"), CL(1, "TO LIT EXTRA BALL"), "", eNone, eNone, eNone, 1000, False, ""
    tmp = 10 - Chimichangas(CurrentPlayer)MOD 10
    DMD CL(0, "COLLECT " &tmp& " CHIMICHA"), CL(1, "TO INCREASE PFX"), "", eNone, eNone, eNone, 1000, False, ""
    tmp = 30 - Weapons(CurrentPlayer)MOD 30
    DMD CL(0, "COLLECT " &tmp& " WEAPONS"), CL(1, "TO LIT EXTRA BALL"), "", eNone, eNone, eNone, 1000, False, ""
    tmp = 50 - Weapons(CurrentPlayer)MOD 50
    DMD CL(0, "COLLECT " &tmp& " WEAPONS"), CL(1, "TO INCREASE PFX"), "", eNone, eNone, eNone, 1000, False, ""
    tmp = 75 - Weapons(CurrentPlayer)MOD 75
    DMD CL(0, "COLLECT " &tmp& " WEAPONS"), CL(1, "TO MECHSUIT MB"), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "BONUS X"), CL(1, BonusMultiplier(CurrentPlayer)), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "BALLS LOCKED"), CL(1, BallsInLock(CurrentPLayer)), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "JACKPOT VALUES"), "", "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "NINJA MB JACKPOT"), CL(1, FormatScore(NinjaMBJackpot(CurrentPlayer))), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "MEGALODON JACKPOT"), CL(1, FormatScore(MegaJackpot(CurrentPlayer))), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "MEGALODON SUPJACKPOT"), CL(1, FormatScore(MegaJackpot(CurrentPlayer) * MegaSuperJCount)), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "LIL DP JACKPOT"), CL(1, FormatScore(LilDPJackpot(CurrentPlayer))), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "DISCO MB JACKPOT"), CL(1, FormatScore(DiscoSuperJackpot(CurrentPlayer))), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "MECHSUIT MB JACKPOT"), CL(1, FormatScore(MechMBJackpot(CurrentPlayer))), "", eNone, eNone, eNone, 1000, False, ""
    DMD CL(0, "HIGHEST SCORE"), CL(1, HighScoreName(0) & " " & HighScore(0)), "", eNone, eNone, eNone, 1000, False, ""
End Sub

Sub StopMBmodes                                   'stop multiball modes after loosing the last multibal
    If bLilDPMB Then StopLilDPMB:ResetDropTargets 'this stops the LilDP multiball and reset the targets and lights
    If bDiscoMBEnabled Then StopDiscoMB           'discoMB is on so stop it
    If bNinjaMB Then StopNinjaMB
    If bMechSuitMBStarted Then StopMechSuitMB
    If bBallLilDPLocked Then 'a ball is locked in the lilDP and no MB modes are running anymore, then just release the ball
        DropDownTargets
        vpmtimer.addtimer 1500, "ResetDropTargets '"
    End If
    Trigger017.TimerEnabled = 1 'last check to release a locked ball
    changesong
End Sub

Sub StopEndOfBallMode() 'this sub is called after the last ball in play is drained, reset skillshot, modes, timers
    pfxtimer.Enabled = 0
    StopCombo
    StopBattle
    If bDiscoLoopsEnabled Then DiscoLoopsTimer_Timer 'stops the discolopp mode
    If bBerserker Then StopBerserker
    If bKatanaramaTime Then StopKatanarama
End Sub

Sub ResetNewBallVariables() 'reset variables for a new ball or player
    dim i
    'turn on or off the needed lights before a new ball is released
    TurnOffPlayfieldLights
    'lilDp droptargets
    ResetDropTargets
    LilDPHits = 0
    CloseGates
    'reset playfield multipiplier
    SetPlayfieldMultiplier 1
    If Balls = 1 then 'only on the first ball
        bBattleReady = True
        bLockEnabled = True:SwordEffect 1
    End If
    'update dead, pool, boom & other lights
    TurnOffArrows
    For i = 0 to 6
        BattleLights(i) = 0
        StarLights(i) = 0
        ChimiChangaLights(i) = 0
        LilDPLights(i) = 0
    Next
    TurnOnTeamUpLights
End Sub

Sub TurnOnTeamUpLights
    li073.State = 2
    li040.State = 2
    li074.State = 2
    li072.State = 2
End Sub

Sub TurnOffTeamUpLights
    li073.State = 0
    li040.State = 0
    li074.State = 0
    li072.State = 0
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub UpdateLights
    Dim i
    li021.State = Dead(CurrentPlayer, 1)
    li022.State = Dead(CurrentPlayer, 2)
    li023.State = Dead(CurrentPlayer, 3)
    li024.State = Dead(CurrentPlayer, 4)
    li017.State = Pool(CurrentPlayer, 1)
    li018.State = Pool(CurrentPlayer, 2)
    li019.State = Pool(CurrentPlayer, 3)
    li020.State = Pool(CurrentPlayer, 4)
    li006.State = Boom(CurrentPlayer, 1)
    li014.State = Boom(CurrentPlayer, 2)
    li015.State = Boom(CurrentPlayer, 3)
    li016.State = Boom(CurrentPlayer, 4)
    li045.State = WeaponLights(CurrentPlayer, 1)
    li028.State = WeaponLights(CurrentPlayer, 2)
    li042.State = WeaponLights(CurrentPlayer, 3)
    li046.State = WeaponLights(CurrentPlayer, 4)
    li032.State = WeaponLights(CurrentPlayer, 5)
    li033.State = WeaponLights(CurrentPlayer, 6)
    li025.State = DiamondLights(CurrentPlayer, 1)
    li029.State = DiamondLights(CurrentPlayer, 2)
    li026.State = DiamondLights(CurrentPlayer, 3)
    li030.State = DiamondLights(CurrentPlayer, 4)
    li027.State = DiamondLights(CurrentPlayer, 5)
    li031.State = DiamondLights(CurrentPlayer, 6)
    li008.State = Battle(CurrentPLayer, 1)
    li010.State = Battle(CurrentPLayer, 2)
    li011.State = Battle(CurrentPLayer, 3)
    li007.State = Battle(CurrentPLayer, 4)
    li009.State = Battle(CurrentPLayer, 5)
    li012.State = Battle(CurrentPLayer, 6)
    li013.State = Battle(CurrentPLayer, 7)
    If bBattleReady Then
        li056.State = 2
        li070.State = 2
    Else
        li056.State = 0
        li070.State = 0
    End If
    If bLockEnabled Then
        li048.State = 2
    Else
        li048.State = 0
    End If
'    If BoomCount> 0 then
'        BoomLight.State = 2
'    Else
'        BoomLight.State = 0
'    End If
     If BoomCount> 0 then
       BoomLight.State = 2
       If BoomCount <= 3 then
         DOF 200, DOFOn: DOF 201, DOFOff 'White Fire button
       Else
         DOF 200, DOFOff: DOF 201, DOFOn 'Red Fire button
       End If
     Else
       BoomLight.State = 0
       DOF 200, 0: DOF 201, 0	'Fire button off
     End If
    If bNinjaMBSJackpot Then
        SetLightColor li058, teal, 2
    ElseIf bSkillShotReady Then
        SetLightColor li058, red, 2
    End If
    If bMultiBallMode Then
        li060.state = 2:li061.State = 2
    Else
        li060.state = 0:li061.State = 0
    End If
    If bMechSuitMBLight(CurrentPlayer)Then
        li047.State = 1
    End If
End Sub

Sub UpdateArrows_Timer 'timer change the color of the blinking lights according to the Battle or Mode active
    'Update other lights
    UpdateLights
    Select Case UpdateArrowsCount
        Case 0 'Battle lights
            Select Case Battle(CurrentPlayer, 0)
                Case 1
                    If BattleLights(1) = 2 Then SetlightColor li054, red, 2
                    If BattleLights(2) = 2 Then SetlightColor li057, red, 2
                    If BattleLights(3) = 2 Then SetlightColor li055, red, 2
                    If BattleLights(4) = 2 Then SetlightColor li071, red, 2
                    If BattleLights(5) = 2 Then SetlightColor li058, red, 2
                    If BattleLights(6) = 2 Then SetlightColor li059, red, 2
                Case 2
                    If BattleLights(1) = 2 Then SetlightColor li054, white, 2
                    If BattleLights(2) = 2 Then SetlightColor li057, white, 2
                    If BattleLights(3) = 2 Then SetlightColor li055, white, 2
                    If BattleLights(4) = 2 Then SetlightColor li071, white, 2
                    If BattleLights(5) = 2 Then SetlightColor li058, white, 2
                    If BattleLights(6) = 2 Then SetlightColor li059, white, 2
                Case 3
                    If BattleLights(1) = 2 Then SetlightColor li054, yellow, 2
                    If BattleLights(2) = 2 Then SetlightColor li057, yellow, 2
                    If BattleLights(3) = 2 Then SetlightColor li055, yellow, 2
                    If BattleLights(4) = 2 Then SetlightColor li071, yellow, 2
                    If BattleLights(5) = 2 Then SetlightColor li058, yellow, 2
                    If BattleLights(6) = 2 Then SetlightColor li059, yellow, 2
                Case 4
                    If BattleLights(1) = 2 Then SetlightColor li054, darkgreen, 2
                    If BattleLights(2) = 2 Then SetlightColor li057, darkgreen, 2
                    If BattleLights(3) = 2 Then SetlightColor li055, darkgreen, 2
                    If BattleLights(4) = 2 Then SetlightColor li071, darkgreen, 2
                    If BattleLights(5) = 2 Then SetlightColor li058, darkgreen, 2
                    If BattleLights(6) = 2 Then SetlightColor li059, darkgreen, 2
                Case 5
                    If BattleLights(1) = 1 Then SetlightColor li054, blue, 1
                    If BattleLights(2) = 1 Then SetlightColor li057, blue, 1
                    If BattleLights(3) = 1 Then SetlightColor li055, blue, 1
                    If BattleLights(4) = 1 Then SetlightColor li071, blue, 1
                    If BattleLights(5) = 1 Then SetlightColor li058, blue, 1
                    If BattleLights(6) = 1 Then SetlightColor li059, blue, 1
                    If BattleLights(1) = 2 Then SetlightColor li054, red, 2
                    If BattleLights(2) = 2 Then SetlightColor li057, red, 2
                    If BattleLights(3) = 2 Then SetlightColor li055, red, 2
                    If BattleLights(4) = 2 Then SetlightColor li071, red, 2
                    If BattleLights(5) = 2 Then SetlightColor li058, red, 2
                    If BattleLights(6) = 2 Then SetlightColor li059, red, 2
                Case 6
                    If BattleLights(1) = 2 Then SetlightColor li054, green, 2
                    If BattleLights(2) = 2 Then SetlightColor li057, green, 2
                    If BattleLights(3) = 2 Then SetlightColor li055, green, 2
                    If BattleLights(4) = 2 Then SetlightColor li071, green, 2
                    If BattleLights(5) = 2 Then SetlightColor li058, green, 2
                    If BattleLights(6) = 2 Then SetlightColor li059, green, 2
                Case 7
                    If BattleLights(1) = 2 Then SetlightColor li054, darkblue, 2
                    If BattleLights(2) = 2 Then SetlightColor li057, darkblue, 2
                    If BattleLights(3) = 2 Then SetlightColor li055, darkblue, 2
                    If BattleLights(4) = 2 Then SetlightColor li071, darkblue, 2
                    If BattleLights(5) = 2 Then SetlightColor li058, darkblue, 2
                    If BattleLights(6) = 2 Then SetlightColor li059, darkblue, 2
            End Select
        Case 1 'Chimichangas
            If ChimiChangaLights(1) = 2 Then SetlightColor li054, orange, 2
            If ChimiChangaLights(2) = 2 Then SetlightColor li057, orange, 2
            If ChimiChangaLights(3) = 2 Then SetlightColor li055, orange, 2
            If ChimiChangaLights(4) = 2 Then SetlightColor li071, orange, 2
            If ChimiChangaLights(5) = 2 Then SetlightColor li058, orange, 2
            If ChimiChangaLights(6) = 2 Then SetlightColor li059, orange, 2
        Case 2 'Lil Deadpool
            If LilDPLights(1) = 2 Then SetlightColor li054, amber, 2
            If LilDPLights(2) = 2 Then SetlightColor li057, amber, 2
            If LilDPLights(3) = 2 Then SetlightColor li055, amber, 2
            If LilDPLights(4) = 2 Then SetlightColor li071, amber, 2
            If LilDPLights(5) = 2 Then SetlightColor li058, amber, 2
            If LilDPLights(6) = 2 Then SetlightColor li059, amber, 2
        Case 3 'Ninja Multiball jackpot lights
            If NinjaMBLights(1) = 2 Then SetlightColor li054, teal, 2
            If NinjaMBLights(2) = 2 Then SetlightColor li057, teal, 2
            If NinjaMBLights(3) = 2 Then SetlightColor li055, teal, 2
            If NinjaMBLights(4) = 2 Then SetlightColor li071, teal, 2
            If NinjaMBLights(5) = 2 Then SetlightColor li058, teal, 2
            If NinjaMBLights(6) = 2 Then SetlightColor li059, teal, 2
        Case 4 'Stars
            If StarLights(1) = 2 Then SetlightColor li054, purple, 2
            If StarLights(2) = 2 Then SetlightColor li057, purple, 2
            If StarLights(3) = 2 Then SetlightColor li055, purple, 2
            If StarLights(4) = 2 Then SetlightColor li071, purple, 2
            If StarLights(5) = 2 Then SetlightColor li058, purple, 2
            If StarLights(6) = 2 Then SetlightColor li059, purple, 2
        Case 5 'Mech Suit Multiball
            If MechSuitLights(1) = 2 Then SetlightColor li054, red, 2
            If MechSuitLights(2) = 2 Then SetlightColor li057, red, 2
            If MechSuitLights(3) = 2 Then SetlightColor li055, red, 2
            If MechSuitLights(4) = 2 Then SetlightColor li071, red, 2
            If MechSuitLights(5) = 2 Then SetlightColor li058, red, 2
            If MechSuitLights(6) = 2 Then SetlightColor li059, red, 2
    End Select
    UpdateArrowsCount = (UpdateArrowsCount + 1)MOD 6
End Sub

Sub UpdateSkillShot() 'Setup and updates the skillshot lights
    LightSeqSkillshot.Play SeqAllOff
    DMD CL(0, "HIT LIT LIGHT"), CL(1, "FOR SKILLSHOT"), "", eNone, eNone, eNone, 1500, True, ""
    li050.State = 2
    SetLightColor li058, Red, 2
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    bRotateLights = True
    LightSeqSkillshot.StopPlay
    If Li049.State = 2 Then Li049.State = 0
    If Li050.State = 2 Then Li050.State = 0
    If Li051.State = 2 Then Li051.State = 0
    If Li052.State = 2 Then Li052.State = 0
    Li058.State = 0
    CloseGates
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
' In this table the slingshots change the outlanes lights

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
    If Tilted Then Exit Sub
DOF 103, DOFPulse
    RandomSoundSlingshotLeft Lemk
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 530
    ' check modes
    StopCombo
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
    If Tilted Then Exit Sub
DOF 104, DOFPulse
    RandomSoundSlingshotRight Remk
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 530
    ' check modes
    StopCombo
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'*********
' Rubbers
'*********

Sub rlband004_Hit:Addscore 110:End Sub
Sub rlband003_Hit:Addscore 110:End Sub
Sub rlband007_Hit:Addscore 110:End Sub

'***********************
' Bumpers - Spawn Ninja
'***********************

Sub Bumper1_Hit
    If Tilted Then Exit Sub
DOF  107,DOFPulse
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    RandomSoundBumperTop Bumper1
  pupevent 879
    FlashForms Flasher008, 1500, 75, 0
    ' add some points
    AddScore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPLayer, 0)
        Case 0:PlaySound "sfx_sword1"
        Case 3 'Sabretooth
            Addscore DominoValue
            PlaySound ""
            CheckBattle
    End Select

    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper1"
    CheckBumpers
End Sub

Sub Bumper2_Hit
    If Tilted Then Exit Sub
DOF  108,DOFPulse
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    RandomSoundBumperMiddle Bumper2
    FlashForms Flasher009, 1500, 75, 0
    ' add some points
    AddScore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPLayer, 0)
        Case 0:PlaySound "sfx_sword2"
        Case 3 'Sabretooth
            Addscore DominoValue
            PlaySound ""
            CheckBattle
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper2"
    CheckBumpers
End Sub

Sub Bumper3_Hit
    If Tilted Then Exit Sub
DOF  109,DOFPulse
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    RandomSoundBumperBottom Bumper3
    FlashForms Flasher010, 1500, 75, 0
    ' add some points
    AddScore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPLayer, 0)
        Case 0:PlaySound "sfx_sword3"
        Case 3 'Sabretooth
            Addscore DominoValue
            PlaySound ""
            CheckBattle
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Bumper3"
    CheckBumpers
End Sub

' Check the bumper hits - spawn Ninjas

Sub CheckBumpers()
    ' increase the bumper hit count and increase the bumper value after each 30 hits
    BumperHits(CurrentPlayer) = BumperHits(CurrentPlayer) + 1
    If BumperHits(CurrentPlayer)MOD 10 = 0 Then ' spawn a Ninja Star - purple Arrow light
        SpawnNinja
    End If
End Sub

Sub SpawnNinja
    Dim tmp
    tmp = RndNbr(6)
    StarLights(tmp) = 2
End Sub

Sub CheckNinjaStars
    NinjaStars(CurrentPlayer) = NinjaStars(CurrentPlayer) + 1
    If NinjaStars(CurrentPlayer)MOD 50 = 0 Then 'light extra ball
        DMD "_", CL(1, "EXTRA BALL IS LIT"), "", eNone, eNone, eNone, 1000, True, "vo_lildp_xtraball"
        li038.State = 2
    End If
    If NinjaStars(CurrentPlayer)MOD 25 = 0 Then 'increase pfx
        AddPlayfieldMultiplier 1
    End If
End Sub

'*******************************
' Top lanes: Bonus X /Skillshot
'*******************************
' score 2.500.000, 250.000 or 1000

Sub Trigger006_Hit
    If Tilted Then Exit Sub
    ' check for modes
    StopCombo
    If bSkillShotReady AND li049.State = 2 Then
        AwardSkillshot
    Else
        Addscore 1000
        li049.State = 1
        CheckBonusX
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger006"
End Sub

Sub Trigger007_Hit
    If Tilted Then Exit Sub
    ' check for modes
    StopCombo
    If bSkillShotReady AND li050.State = 2 Then
        AwardSkillshot
    Else
        Addscore 1000
        li050.State = 1
        CheckBonusX
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger007"
End Sub

Sub Trigger008_Hit
    If Tilted Then Exit Sub
    ' check for modes
    StopCombo
    If bSkillShotReady AND li051.State = 2 Then
        AwardSkillshot
    Else
        Addscore 1000
        li051.State = 1
        CheckBonusX
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger008"
End Sub

Sub Trigger009_Hit
    If Tilted Then Exit Sub
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPLayer, 0)
        Case 1
            If li054.State = 2 Then CheckBattle
    End Select
    If bSkillShotReady AND li052.State = 2 Then
        AwardSkillshot
    Else
        Addscore 1000
        li052.State = 1
        CheckBonusX
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger009"
End Sub

Sub CheckBonusX
    If Li049.State + Li050.State + Li051.State + Li052.State = 4 Then
        AddBonusMultiplier 1
        DMD CL(0, "BONUS MULTIPLIER"), CL(1, BonusMultiplier(CurrentPlayer) & "X"), "", eNone, eNone, eNone, 1500, True, "Fanfare"
        Addscore 15000
        GiEffect 1
        FlashForMs Li049, 1500, 100, 0
        FlashForMs Li050, 1500, 100, 0
        FlashForMs Li051, 1500, 100, 0
        FlashForMs Li052, 1500, 100, 0
        FlashForms Flasher008, 1500, 100, 0
        FlashForms Flasher009, 1500, 100, 0
        FlashForms Flasher010, 1500, 100, 0
    ElseIf bInfoNeeded6(CurrentPLayer)Then
        PlaySound "vo_completetoplanes"
        bInfoNeeded6(CurrentPLayer) = False
    End IF
End Sub

'**************************
' Out/inLanes : BOOM lanes
'**************************

Sub Trigger001_Hit
    If Tilted Then Exit Sub
    ' check for modes
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    If li075.State Then                   'Regenerate
        If NOT bExtraBallWonThisBall Then 'same as extra ball
            DMD "_", CL(1, ("REGENERATE")), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("Knocker", 122, DOFPulse, DOFKnocker)
            DOF 121, DOFPulse
            ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
            bExtraBallWonThisBall = True
            GiEffect 2
            LightEffect 2
        END If
    End If
    Addscore 5000
    Boom(CurrentPlayer, 1) = 1
    CheckBOOM
  pupevent 877
End Sub

Sub Trigger002_Hit
    If Tilted Then Exit Sub
    ' check for modes
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    Addscore 1000
    Boom(CurrentPlayer, 2) = 1
    CheckBOOM
  pupevent 877
End Sub

Sub Trigger003_Hit
    If Tilted Then Exit Sub
    ' check for modes
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    Addscore 1000
    Boom(CurrentPlayer, 3) = 1
    CheckBOOM
  pupevent 877
End Sub

Sub Trigger004_Hit
    If Tilted Then Exit Sub
    ' check for modes
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    Addscore 5000
    Boom(CurrentPlayer, 4) = 1
    CheckBOOM
  pupevent 877
End Sub

Sub CheckBOOM
    Dim tmp, i
    tmp = 0
    For i = 1 to 4:tmp = tmp + Boom(CurrentPlayer, i):Next
    If tmp = 4 Then
        PlaySound "vo_boombuttonisready" : pupevent 846
        BoomCount = BoomCount + 1
        For i = 1 to 4:Boom(CurrentPlayer, i) = 0:Next
'        If BoomCount> 3 Then
'            SetLightColor BoomLight, red, 2
'        Else
'            SetLightColor BoomLight, white, 2
'        End If
        If BoomCount> 3 Then
          SetLightColor BoomLight, red, 2
          DOF 200, DOFOff: DOF 201, DOFOn
        Else
          SetLightColor BoomLight, white, 2
          DOF 200, DOFOn: DOF 201, DOFOff
        End If
    End If
End Sub

Sub BoomHit
    If BoomCount> 0 Then
        'PlaySound "vo_boom"
         DMD " ", " ", "d_boom", eNone, eNone, eNone, 1500, True, "vo_boom" : pupevent 845
        If Battle(CurrentPlayer, 0)> 0 Then
            LifeLeft(CurrentPlayer, Battle(CurrentPlayer, 0)) = LifeLeft(CurrentPlayer, Battle(CurrentPlayer, 0))- 4
            CheckBattle
            'DMDScoreNow
        ElseIf BoomCount> 3 Then 'not in a battle
            Addscore 5000000
        Else
            Addscore 500000
        End If
        BoomCount = BoomCount - 1
    End If
End Sub

'*************
' Battle lanes
'*************
' also used for combos and other modes

Sub Trigger011_Hit 'colossus loop
    If Tilted Then Exit Sub
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    Addscore 10000
    If li039.State = 2 Then
        li039.State = 0
        AwardSpecial
    End If
    ' check for modes
    If ChimiChangaLights(3) = 2 Then 'collect chimichanga
        ChimiChangaLights(3) = 0:Li055.State = 0
        Collectchimichanga
    End If
    If MechSuitLights(3) = 2 Then 'collect jackpot
        MechSuitLights(3) = 0:Li055.State = 0
        CheckMechSuitMBHits
    End If
    If WeaponLights(CurrentPlayer, 3) = 1 Then 'collect weapon
        WeaponLights(CurrentPlayer, 3) = 0:Li042.State = 0
        CheckWeapons
    End If
    If StarLights(3) = 2 Then 'a ninja star is at that position
        StarLights(3) = 0:Li055.State = 0
        CheckNinjaStars
    End If
    If bNinjaMB AND NinjaMBLights(3) = 2 Then 'Ninja Multiball
        NinjaMBLights(3) = 0:li055.State = 0
        NinjaMBCheckHits
    End If
    If bLilDPMB AND LilDPHits> 8 AND LilDPLights(3) = 2 Then 'second part of the LilDP multiball
        LilDPLights(3) = 0:Li055.State = 0
        LilDPCheckHits
    End If
    Select Case Battle(CurrentPLayer, 0)
        Case 0
            ColossusPower(CurrentPlayer) = ColossusPower(CurrentPlayer) + 1
            CheckColossus
        Case 1, 4, 6, 7
            If BattleLights(3) = 2 Then PlaySound "":CheckBattle
        Case 5
            If BattleLights(3) = 2 Then
                PlaySound "":AwardSuperMegaJackpot:CheckBattle
            ElseIf BattleLights(3) = 1 Then
                BattleLights(3) = 0:li055.State = 0:AwardMegaJackpot
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger011"
End Sub

Sub Trigger010_Hit 'left loop
    If Tilted Then Exit Sub
    Addscore 10000
    ' check for modes
    If bKatanaramaTime Then
        Addscore KatanaramaValue(CurrentPlayer)
        DMD " KATANARAMA JACKPOT", CL(1, FormatScore(KatanaramaValue(CurrentPlayer))), "", eNone, eBlink, eNone, 1500, True, "vo_jackpot" : pupevent 824
        KatanaramaValue(CurrentPlayer) = KatanaramaValue(CurrentPlayer) + 250000
    End If
    If ChimiChangaLights(1) = 2 Then 'collect chimichanga
        ChimiChangaLights(1) = 0:Li054.State = 0
        Collectchimichanga
    End If
    If MechSuitLights(1) = 2 Then 'collect jackpot
        MechSuitLights(1) = 0:Li054.State = 0
        CheckMechSuitMBHits
    End If
    If WeaponLights(CurrentPlayer, 1) = 1 Then 'collect weapon
        WeaponLights(CurrentPlayer, 1) = 0:li045.State = 0
        CheckWeapons
    End If
    If StarLights(1) = 2 Then 'a ninja star is at that position
        StarLights(1) = 0:li054.State = 0
        CheckNinjaStars
    End If
    If bNinjaMB AND NinjaMBLights(1) = 2 Then 'Ninja Multiball
        NinjaMBLights(1) = 0:li054.State = 0
        NinjaMBCheckHits
    End If
    If bLilDPMB AND LilDPHits> 8 AND LilDPLights(1) = 2 Then 'second part of the LilDP multiball
        LilDPCheckHits
        LilDPLights(1) = 0:li054.State = 0
    End If
    If bDiscoLoopsEnabled AND LastSwitchHit = "Trigger005" Then
        Jackpot(CurrentPlayer) = DiscoLoopsValue(CurrentPlayer)
        DiscoLoopsValue(CurrentPlayer) = DiscoLoopsValue(CurrentPlayer) + 150000
        AwardJackpot
    End If
    If ActiveBall.VelY <0 Then 'ball is going up
        Select Case Battle(CurrentPLayer, 0)
            Case 0
                DazzlePower(CurrentPlayer) = DazzlePower(CurrentPlayer) + 1
                CheckDazzler
            Case 1, 4, 6, 7
                If BattleLights(1) = 2 Then PlaySound "":CheckBattle
            Case 5
                If BattleLights(1) = 2 Then
                    PlaySound "":AwardSuperMegaJackpot:CheckBattle
                ElseIf BattleLights(1) = 1 Then
                    li054.State = 0:BattleLights(1) = 0:AwardMegaJackpot
                End If
        End Select
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger010"
End Sub

Sub Trigger005_Hit 'right loop
    If Tilted Then Exit Sub
    If bSkillshotReady Then
        bRotateLights = False
        Exit Sub
    Else
        Addscore 10000
    End If
    ' check for modes
    If bKatanaramaTime Then
        Addscore KatanaramaValue(CurrentPlayer)
        DMD " KATANARAMA JACKPOT", CL(1, FormatScore(KatanaramaValue(CurrentPlayer))), "", eNone, eBlink, eNone, 1500, True, "vo_jackpot"
        KatanaramaValue(CurrentPlayer) = KatanaramaValue(CurrentPlayer) + 250000
    End If
    If ChimiChangaLights(6) = 2 Then 'if the chimichanga light is on then collect it
        ChimiChangaLights(6) = 0:Li059.State = 0
        CollectChimichanga
    Else 'count the chimichanga hits
        CheckChimichangaHits
    End If
    If MechSuitLights(6) = 2 Then 'collect jackpot
        MechSuitLights(6) = 0:Li059.State = 0
        CheckMechSuitMBHits
    End If
    If WeaponLights(CurrentPlayer, 6) = 1 Then 'collect weapon
        WeaponLights(CurrentPlayer, 6) = 0:li033.State = 0
        CheckWeapons
    End If
    If StarLights(6) = 2 Then 'a ninja star is at that position
        StarLights(6) = 0:li059.State = 0
        CheckNinjaStars
    End If
    If bNinjaMB AND NinjaMBLights(6) = 2 Then 'Ninja Multiball
        NinjaMBLights(6) = 0:li059.State = 0
        NinjaMBCheckHits
    End If
    If bLilDPMB AND LilDPHits> 8 AND LilDPLights(6) = 2 Then 'second part of the LilDP multiball
        LilDPCheckHits
        li059.State = 0:LilDPLights(6) = 0
    End If
    If bDiscoLoopsEnabled AND LastSwitchHit = "Trigger005" Then
        Jackpot(CurrentPlayer) = DiscoLoopsValue(CurrentPlayer)
        DiscoLoopsValue(CurrentPlayer) = DiscoLoopsValue(CurrentPlayer) + 150000
        AwardJackpot
    End If
    Select Case Battle(CurrentPLayer, 0)
        Case 0
            If bSkillshotReady = False Then
                DominoPower(CurrentPlayer) = DominoPower(CurrentPlayer) + 1
                CheckDomino
            End If
        Case 1, 4, 6, 7
            If BattleLights(6) = 2 Then PlaySound "":CheckBattle
        Case 5
            If BattleLights(6) = 2 Then
                PlaySound "":AwardSuperMegaJackpot:CheckBattle
            ElseIf BattleLights(6) = 1 Then
                li059.State = 0:BattleLights(6) = 0:AwardMegaJackpot
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger005"
End Sub

Sub Trigger012_Hit 'left ramp
    If Tilted Then Exit Sub
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    Addscore 10000
    ' check for modes
    If bKatanaramaTime Then
        Addscore KatanaramaValue(CurrentPlayer)
        DMD " KATANARAMA JACKPOT", CL(1, FormatScore(KatanaramaValue(CurrentPlayer))), "", eNone, eNone, eNone, 1500, True, "vo_jackpot"
        KatanaramaValue(CurrentPlayer) = KatanaramaValue(CurrentPlayer) + 250000
    ElseIf DiamondLights(CurrentPLayer, 1) = 2 OR DiamondLights(CurrentPLayer, 3) = 2 OR DiamondLights(CurrentPLayer, 5) = 2 Then
        CheckKatanarama
    End If
    If ChimiChangaLights(2) = 2 Then 'collect chimichanga
        ChimiChangaLights(2) = 0:Li057.State = 0
        Collectchimichanga
    End If
    If MechSuitLights(2) = 2 Then 'collect jackpot
        MechSuitLights(2) = 0:Li057.State = 0
        CheckMechSuitMBHits
    End If
    If WeaponLights(CurrentPlayer, 2) = 1 Then 'collect weapon
        WeaponLights(CurrentPlayer, 2) = 0:li028.State = 0
        CheckWeapons
    End If
    If StarLights(2) = 2 Then 'a ninja star is at that position
        StarLights(2) = 0:li057.State = 0
        CheckNinjaStars
    End If
    If bNinjaMB AND NinjaMBLights(2) = 2 Then 'Ninja Multiball
        NinjaMBLights(2) = 0:li057.State = 0
        NinjaMBCheckHits
    End If
    If bLilDPMB AND LilDPHits> 8 AND LilDPLights(2) = 2 Then 'second part of the LilDP multiball
        LilDPCheckHits
        li057.State = 0:LilDPLights(2) = 0
    End If
    If bDiscoMBEnabled Then
        SuperJackpot(CurrentPlayer) = DiscoSuperJackpot(CurrentPlayer)
        AwardSuperjackpot
    End If
    CheckCombo "Trigger012"
    Select Case Battle(CurrentPLayer, 0)
        Case 1, 2, 4, 6
            If BattleLights(2) = 2 Then PlaySound "":CheckBattle
        Case 5
            If BattleLights(2) = 2 Then
                PlaySound "":AwardSuperMegaJackpot:CheckBattle
            ElseIf BattleLights(2) = 1 Then
                li057.State = 0:BattleLights(2) = 0:AwardMegaJackpot
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger012"
End Sub

Sub Trigger016_Hit 'just a quote trigger
    Playquote
End Sub

Sub Trigger013_Hit 'right ramp
    If Tilted Then Exit Sub
    If bSkillShotReady Then
        AwardSuperSkillShot
    Else
        Addscore 10000
    End If
    ' check for modes
    If LastSwitchHit = "Target014" Then 'Increase Playfield multiplier
        AddPlayfieldMultiplier 1
    End If
    If bDiscoMBEnabled Then
        SuperJackpot(CurrentPlayer) = DiscoSuperJackpot(CurrentPlayer)
        AwardSuperjackpot
    End If
    CheckCombo "Trigger013"
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger013"
End Sub

Sub Trigger014_Hit 'wolverine loop
    If Tilted Then Exit Sub
    If bSkillShotReady Then
        AwardSuperSkillShot
    Else
        Addscore 10000
    End If
    ' check for modes
    If ChimiChangaLights(4) = 2 Then 'collect chimichanga
        ChimiChangaLights(4) = 0:Li071.State = 0
        Collectchimichanga
    End If
    If MechSuitLights(4) = 2 Then 'collect jackpot
        MechSuitLights(4) = 0:Li071.State = 0
        CheckMechSuitMBHits
    End If
    If WeaponLights(CurrentPlayer, 4) = 1 Then 'collect weapon
        WeaponLights(CurrentPlayer, 4) = 0:li046.State = 0
        CheckWeapons
    End If
    If StarLights(4) = 2 Then 'a ninja star is at that position
        StarLights(4) = 0:li071.State = 0
        CheckNinjaStars
    End If
    If bNinjaMB AND NinjaMBLights(4) = 2 Then 'Ninja Multiball
        NinjaMBLights(4) = 0:li071.State = 0
        NinjaMBCheckHits
    End If
    If bLilDPMB AND LilDPHits> 8 AND LilDPLights(4) = 2 Then 'second part of the LilDP multiball
        LilDPCheckHits
        li071.State = 0:LilDPLights(4) = 0
    End If
    CheckCombo "Trigger014"
    Select Case Battle(CurrentPLayer, 0)
        Case 0
            WolverinePower(CurrentPlayer) = WolverinePower(CurrentPlayer) + 1
            CheckWolverine
        Case 2, 4, 6, 7
            If BattleLights(4) = 2 Then PlaySound "":CheckBattle
        Case 5
            If BattleLights(4) = 2 Then
                PlaySound "":AwardSuperMegaJackpot:CheckBattle
            ElseIf BattleLights(4) = 1 Then
                li071.State = 0:BattleLights(4) = 0:AwardMegaJackpot
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Trigger014"
End Sub

Sub Trigger015_Hit 'cross shot
    If Tilted Then Exit Sub
    Addscore 10000
    ' check for modes
    If bKatanaramaTime Then
        Addscore KatanaramaValue(CurrentPlayer)
        DMD " KATANARAMA JACKPOT", CL(1, FormatScore(KatanaramaValue(CurrentPlayer))), "", eNone, eNone, eNone, 1500, True, "vo_jackpot" : pupevent 824
        KatanaramaValue(CurrentPlayer) = KatanaramaValue(CurrentPlayer) + 250000
    ElseIf DiamondLights(CurrentPLayer, 2) = 2 OR DiamondLights(CurrentPLayer, 4) = 2 OR DiamondLights(CurrentPLayer, 6) = 2 Then
        CheckKatanarama
    End If
    If ChimiChangaLights(5) = 2 Then 'collect chimichanga
        ChimiChangaLights(5) = 0:Li058.State = 0
        Collectchimichanga
    End If
    If MechSuitLights(5) = 2 Then 'collect jackpot
        MechSuitLights(5) = 0:Li058.State = 0
        CheckMechSuitMBHits
    End If
    If WeaponLights(CurrentPlayer, 5) = 1 Then 'collect weapon
        WeaponLights(CurrentPlayer, 5) = 0:li032.State = 0
        CheckWeapons
    End If
    If StarLights(5) = 2 Then 'a ninja star is at that position
        StarLights(5) = 0:li058.State = 0
        CheckNinjaStars
    End If
    If bNinjaMB AND NinjaMBLights(5) = 2 Then 'Ninja Multiball
        NinjaMBLights(5) = 0:li058.State = 0
        NinjaMBCheckHits
    End If
    If bLilDPMB AND LilDPHits> 8 AND LilDPLights(5) = 2 Then 'second part of the LilDP multiball
        LilDPCheckHits
        li058.State = 0:LilDPLights(5) = 0
    End If
    Select Case Battle(CurrentPLayer, 0)
        Case 1, 2, 4, 6
            If BattleLights(5) = 2 Then PlaySound "":CheckBattle
        Case 5
            If BattleLights(5) = 2 Then
                PlaySound "":AwardSuperMegaJackpot:CheckBattle
            ElseIf BattleLights(5) = 1 Then
                li058.State = 0:BattleLights(5) = 0:AwardMegaJackpot
            End If
    End Select
End Sub

Sub Trigger017_Hit 'LilDP Ball locked
    bBallLilDPLocked = True
End Sub

Sub Trigger017_UnHit 'LilDP Ball locked has been released
    bBallLilDPLocked = False
End Sub

Sub Trigger017_Timer 'LilDP Ball lock last check
    If bBallLilDPLocked Then
        DropDownTargets
        vpmtimer.addtimer 1500, "ResetDropTargets '"
    Else
        Trigger017.TimerEnabled = 0
    End If
End Sub

'**************
' DEAD targets
'**************

Sub Target001_Hit 'lower d
    If Tilted Then Exit Sub
    Addscore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPlayer, 0)
        Case 0 'no battle active
            Dead(CurrentPlayer, 1) = 1
            CheckDead
        Case 1 'Juggernaut
            If LifeLeft(CurrentPlayer, 1)> 6 Then
                Addscore DominoValue
                PlaySound ""
                CheckBattle
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target001"
End Sub

Sub Target002_Hit 'e
    If Tilted Then Exit Sub
    Addscore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPlayer, 0)
        Case 0 'no battle active
            Dead(CurrentPlayer, 2) = 1
            CheckDead
        Case 1 'Juggernaut
            If LifeLeft(CurrentPlayer, 1)> 6 Then
                Addscore DominoValue
                PlaySound ""
                CheckBattle
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target002"
End Sub

Sub Target003_Hit 'a
    If Tilted Then Exit Sub
    Addscore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPlayer, 0)
        Case 0 'no battle active
            Dead(CurrentPlayer, 3) = 1
            CheckDead
        Case 1 'Juggernaut
            If LifeLeft(CurrentPlayer, 1)> 6 Then
                Addscore DominoValue
                PlaySound ""
                CheckBattle
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target003"
End Sub

Sub Target004_Hit 'd
    If Tilted Then Exit Sub
    Addscore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPlayer, 0)
        Case 0 'no battle active
            Dead(CurrentPlayer, 4) = 1
            CheckDead
        Case 1 'Juggernaut
            If LifeLeft(CurrentPlayer, 1)> 6 Then
                Addscore DominoValue
                PlaySound ""
                CheckBattle
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target004"
End Sub

Sub CheckDead
    Dim tmp, i
    tmp = 0
    For i = 1 to 4:tmp = tmp + Dead(CurrentPLayer, i):Next
    If tmp = 2 AND bInfoNeeded2(CurrentPLayer)AND bBattleready = False Then PlaySound "vo_completedeadtargets":bInfoNeeded2(CurrentPLayer) = False
    If tmp = 4 then 'enable Battle at the scoop
        DMD "   BATTLE IS READY  ", "  AT THE HELLHOUSE  ", "", eNone, eNone, eNone, 1000, True, "vo_battleislit" : pupevent 826
        Flashforms flasher012, 800, 50, 1
        DOF 128, DOFPulse
        Addscore 250000
        li056.State = 2
        li070.BlinkInterval = 300:li070.State = 2
        bBattleready = True
        'reset targets
        For i = 1 to 4:Dead(CurrentPLayer, i) = 0:Next
        ' count how many times Pool targets has been all hit
        Pool(CurrentPLayer, 0) = Pool(CurrentPLayer, 0) + 1
        CheckDeadPool
    End If
End Sub

'**************
' POOL targets
'**************

Sub Target005_Hit 'l
    If Tilted Then Exit Sub
    Addscore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPlayer, 0)
        Case 0 'no battle active
            Pool(CurrentPlayer, 1) = 1
            CheckPool
        Case 1 'Juggernaut
            If LifeLeft(CurrentPlayer, 1)> 6 Then
                Addscore DominoValue
                CheckBattle
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target005"
End Sub

Sub Target006_Hit 'lower o
    If Tilted Then Exit Sub
    Addscore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPlayer, 0)
        Case 0 'no battle active
            Pool(CurrentPlayer, 2) = 1
            CheckPool
        Case 1 'Juggernaut
            If LifeLeft(CurrentPlayer, 1)> 6 Then
                Addscore DominoValue
                CheckBattle
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target006"
End Sub

Sub Target007_Hit 'o
    If Tilted Then Exit Sub
    Addscore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPlayer, 0)
        Case 0 'no battle active
            Pool(CurrentPlayer, 3) = 1
            CheckPool
        Case 1 'Juggernaut
            If LifeLeft(CurrentPlayer, 1)> 6 Then
                Addscore DominoValue
                CheckBattle
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target007"
End Sub

Sub Target008_Hit 'p
    If Tilted Then Exit Sub
    Addscore 1000
    ' check for modes
    StopCombo
    Select Case Battle(CurrentPlayer, 0)
        Case 0 'no battle active
            Pool(CurrentPlayer, 4) = 1
            CheckPool
        Case 1 'Juggernaut
            If LifeLeft(CurrentPlayer, 1)> 6 Then
                Addscore DominoValue
                CheckBattle
            End If
    End Select
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target008"
End Sub

Sub CheckPool
    Dim tmp, i
    tmp = 0
    For i = 1 to 4:tmp = tmp + Pool(CurrentPLayer, i):Next
    If tmp = 2 AND bInfoNeeded3(CurrentPLayer)AND bLockEnabled = False Then PlaySound "vo_completepool-lock":bInfoNeeded3(CurrentPLayer) = False
    If tmp = 4 then 'enable Lock again
        Addscore 250000
        If bLockEnabled = False Then
            DMD "", "", "d_lockislit", eNone, eNone, eBlink, 1500, True, "vo_lockislit"
            bLockEnabled = True
            SwordEffect 1
        End If
        'reset targets
        For i = 1 to 4:Pool(CurrentPLayer, i) = 0:Next
        ' count how many times Pool targets has been all hit
        Pool(CurrentPLayer, 0) = Pool(CurrentPLayer, 0) + 1
        CheckDeadPool
    End if
End Sub

Sub CheckDeadPool 'all targets has been hit
    IF Dead(CurrentPlayer, 0) = 1 OR Pool(CurrentPLayer, 0) = 1 AND bInfoNeeded4(CurrentPLayer)Then bInfoNeeded4(CurrentPLayer) = False:vpmtimer.addtimer 3000, "PlaySound ""vo_completedptargets-mystery"" '"
    IF Dead(CurrentPlayer, 0) = 2 OR Pool(CurrentPLayer, 0) = 2 AND bInfoNeeded5(CurrentPLayer)Then bInfoNeeded5(CurrentPLayer) = False:vpmtimer.addtimer 3000, "PlaySound ""vo_completedptargets-regenerate"" '"
    IF Dead(CurrentPlayer, 0)> 2 AND Pool(CurrentPLayer, 0)> 2 Then
        'regeneration light
        DMD "_", "  REGENERATE IS LIT", "", eNone, eNone, eNone, 1500, True, "vo_regenerateislit" : pupevent 827
        li075.State = 1
        ' and reset the counter
        Dead(CurrentPlayer, 0) = 0
        Pool(CurrentPlayer, 0) = 0
    ElseIf Dead(CurrentPlayer, 0)> 1 AND Pool(CurrentPLayer, 0)> 1 Then
        'enable Mystery light
        DMD "_", "   MYSTERY IS LIT", "", eNone, eNone, eNone, 1500, True, "vo_mysteryislit" : pupevent 828
        li037.State = 1
    End If
End Sub

'*********************
' Lil Deadpool targets
'*********************

Sub Target009_Dropped 'lil drop 1
    PlaySoundAt SoundFXDOF("Drop_Target_Down_1",110,DOFPulse,DOFDropTargets), Target009
    If Tilted Then Exit Sub
    Addscore 1000
    'drop down also the right droptarget to avoid ball getting stuck
    Target011.IsDropped = 1
    li036.State = 0
    li034.State = 0
    ' check for modes
    StopCombo
    setlightcolor li044, blue, 2
    setlightcolor li044a, blue, 2
    If bLilDPMB then
        DropDownTargets
        setlightcolor li044, red, 2
        setlightcolor li044a, red, 2
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target009"
End Sub

Sub Target010_Dropped 'lil drop 2
    PlaySoundAt SoundFXDOF("Drop_Target_Down_2",110,DOFPulse,DOFDropTargets), Target010
    If Tilted Then Exit Sub
    Addscore 1000
    'drop down also the right droptarget to avoid ball getting stuck
    Target011.IsDropped = 1
    li035.State = 0
    li034.State = 0
    ' check for modes
    StopCombo
    setlightcolor li044, blue, 2
    setlightcolor li044a, blue, 2
    If bLilDPMB then
        DropDownTargets
        setlightcolor li044, red, 2
        setlightcolor li044a, red, 2
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target010"
End Sub

Sub Target011_Dropped 'lil drop 3
    PlaySoundAt SoundFXDOF("Drop_Target_Down_3",110,DOFPulse,DOFDropTargets), Target011
    If Tilted Then Exit Sub
    Addscore 1000
    li034.State = 0
    ' check for modes
    StopCombo
    setlightcolor li044, blue, 2
    setlightcolor li044a, blue, 2
    If bLilDPMB then
        DropDownTargets
        setlightcolor li044, red, 2
        setlightcolor li044a, red, 2
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target011"
End Sub

Sub ResetDropTargets
    ' PlaySoundAt "fx_resetdrop", Target010
    If Target009.IsDropped OR Target010.IsDropped OR Target011.IsDropped Then
        PlaySoundAt SoundFXDOF("Drop_Target_Reset_1", 119, DOFPulse, DOFcontactors), Target010
        Target009.IsDropped = 0
        Target010.IsDropped = 0
        Target011.IsDropped = 0
        li034.State = 2:li035.State = 2:li036.State = 2 'lilDP
        li044a.State = 0:li044.State = 0
    End If
End Sub

Sub DropDownTargets 'after the hurry up timer or after the first hit to release the trapped ball
    Target009.IsDropped = 1
    Target010.IsDropped = 1
    Target011.IsDropped = 1
    li034.State = 2:li035.State = 2:li036.State = 2 'lilDP
    LilDPTimer.Enabled = 0
End Sub

Sub Target012_Hit 'lil DP Stand up target
    PlaySoundAtBall SoundFXDOF("",115,DOFPUlse,DOFShaker)
    If Tilted Then Exit Sub
    If Spots.Enabled = 0 Then FlashForms spot1, 1000, 50, 0:FlashForms spot2, 1000, 50, 0:DOF 131, DOFPulse
    Addscore 10000
    ' check for modes
    StopCombo
    If LilDPHits <10 Then LilDPCheckHits
    lilshake activeball.velY
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target012"
End Sub

'***********************
'Lil Deadpool Multiball
'***********************
' variabes used bLilDPMB, LilDPHits and LilDPHitsNeeded(4), LilDPJackpot(4)

Sub LilDPTimer_Timer 'hurry up timer
    LilDPJackpot(CurrentPlayer) = LilDPJackpot(CurrentPlayer)- 5000 * LilDPHitsNeeded(CurrentPlayer)
    If LilDPJackpot(CurrentPlayer) = 250000 Then PlaySound "vo_lildp_multibal" : pupevent 848
    If LilDPJackpot(CurrentPlayer) <= 100000 OR bLilDPMB = False Then 'release the locked ball
        PlaySound "vo_lildp_timeup"
        LilDPJackpot(CurrentPlayer) = 100000
        Me.Enabled = 0
        DropdownTargets
        If bMultiBallMode = False Then 'not in a multibal anymore then reset the droptargets after a short pause to let the ball out
            vpmtimer.addtimer 1500, "ResetDropTargets '"
        End If
    End If
End Sub

Sub PlaylilDPHitSound
    Dim tmp
    tmp = RndNbr(44)
    PlaySound "vo_lildp" &tmp
End Sub

Sub LilDPCheckHits
    LilDPHits = LilDPHits + 1
    'debug.print "lilDPHits " & LilDPHits
    If bLilDPMB = False AND LilDPHits> 4 Then StopLilDPMB:Exit Sub 'this should never happen, but just in case :)
    If LilDPHits <LilDPHitsNeeded(CurrentPlayer)Then
        PlaylilDPHitSound
        lighteffect 4
        setlightcolor li044, blue, 2
        setlightcolor li044a, blue, 2
        Exit Sub
    ElseIf LilDPHits = LilDPHitsNeeded(CurrentPlayer)Then
        lighteffect 4
        setlightcolor li044, green, 2
        setlightcolor li044a, green, 2
        LilDPHits = 4 'to compensate for required hits
    End If

    Select Case LilDPHits
        Case 0, 1, 2, 3 'do nothing
        Case 4          'start multiball
            DMD "       LIL DEADPOOL ", "         MULTIBALL  ", "d_lildpmb", eNone, eNone, eNone, 1500, True, "vo_lildp_multiball2" : pupevent 848
            ResetDroptargets
            AddMultiball 1
            bLilDPMB = True
            LilDPJackpot(CurrentPlayer) = 500000 * LilDPHitsNeeded(CurrentPlayer)
            LilDPHitsNeeded(CurrentPlayer) = LilDPHitsNeeded(CurrentPlayer) + 1
            If LilDPHitsNeeded(CurrentPlayer)> 4 Then LilDPHitsNeeded(CurrentPlayer) = 4
            LilDPTimer.Interval = 350 + DazzlerValue * 10
            LilDPTimer.Enabled = 1
            setlightcolor li044, red, 2
            setlightcolor li044a, red, 2
        Case 5, 6, 7, 8 '5 jackpots at the LilDP target
            setlightcolor li044, red, 2
            setlightcolor li044a, red, 2
            lighteffect 4:PlaySfx
            Jackpot(CurrentPLayer) = LilDPJackpot(CurrentPlayer)
            AwardJackpot
        Case 9
            lighteffect 4:PlaySfx
            Jackpot(CurrentPLayer) = LilDPJackpot(CurrentPlayer)
            AwardJackpot
            'stop the lights at the lilDP target and start the arrow light jackpots
            li034.State = 0:li035.State = 0:li036.State = 0 'lilDP
            li044a.State = 0:li044.State = 0
            'arrow lights
            For x = 1 to 6:LilDPLights(x) = 2:Next
        Case 10, 11, 12, 13 '5 jackpots
            Jackpot(CurrentPLayer) = LilDPJackpot(CurrentPlayer)
            AwardJackpot
        Case 14 'last jackpot and enabled hurry up on the spinners
            Jackpot(CurrentPLayer) = LilDPJackpot(CurrentPlayer)
            AwardJackpot
            For x = 1 to 6:LilDPLights(x) = 0:Next
            li054.State = 0
            li057.State = 0
            li055.State = 0
            li071.State = 0
            li058.State = 0
            li059.State = 0
            LilDPLights(1) = 2
            LilDPLights(4) = 2
            LilDPSpinner.Interval = 20000 + dazzlerValue * 1000
            LilDPSpinner.Enabled = 1
        Case 15 'called after the hurry up spinner timer is out
            setlightcolor li044, red, 2
            setlightcolor li044a, red, 2
            LilDPHits = 4
            AddPlayFieldMultiplier 1
        Case Else 'should never happen
            StopLilDPMB
    End Select
End Sub

Sub StopLilDPMB 'end of LilDP multiball
    LilDPHits = 0
    bLilDPMB = False
    LilDPTimer.Enabled = 0
    LilDPSpinner.Enabled = 0
    For x = 1 to 6:LilDPLights(x) = 0:Next
    li054.State = 0
    li057.State = 0
    li055.State = 0
    li071.State = 0
    li058.State = 0
    li059.State = 0
End Sub

Sub LilDPSpinner_Timer
    Me.Enabled = 0
    For x = 1 to 6:LilDPLights(x) = 0:Next
    li054.State = 0
    li057.State = 0
    li055.State = 0
    li071.State = 0
    li058.State = 0
    li059.State = 0
    LilDPHits = 14 'ensure to go to the right step
    LilDPCheckHits 'go to the next step in LilDP multiball
End Sub

'***********************
'Lil Deadpool animation
'   (simple jiggle)
'***********************

Sub lilshake(strength)
    lilPos = strength
    lilshakeTimer.Enabled = True
End Sub

Sub lilshakeTimer_Timer()
    lildp.Transy = lilPos
    If lilPos = 0 Then Me.Enabled = False:Exit Sub
    If lilPos <0 Then
        lilPos = ABS(lilPos)- 0.1
    Else
        lilPos = - lilPos + 0.1
    End If
End Sub

'***************
' other targets
'***************

Sub Target013_Hit 'colossal jackpot
    If Tilted Then Exit Sub
    ' check for modes
    StopCombo
    If li069.State = 2 Then
        li069.State = 0
        SuperJackpot(CurrentPLayer) = Score(CurrentPlayer) * 0.5 ' 20% of the score
        AwardSuperJackpot
    Else
        Addscore 50000
    End If
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target013"
End Sub

Sub Target014_Hit 'snitch
    If Tilted Then Exit Sub
    PlaySound"vo_snitchshot"
    Addscore 1000
    ' check for modes
    StopCombo
    ' remember last trigger hit by the ball
    LastSwitchHit = "Target014"
End Sub

'**********
' spinners
'**********

Sub Spinner001_Spin 'dazzler
    SoundSpinner Spinner001
    If Tilted Then Exit Sub
    If bBerserker Then
        Addscore BerserkerValue(CurrentPlayer)
        BerserkerValue(CurrentPlayer) = BerserkerValue(CurrentPlayer) + 10000
    Else
        Addscore 1000
        CheckBerserker                 'check for the number of hits to start Berserker rage
    End If
    If bLilDPMB AND LilDPHits> 14 Then 'second part of the LilDP multiball
        Addscore 500000
    End If
    If bDiscoEnabled = False Then 'count the spins when not in Disco Mode
        LeftSpinnerCount(CurrentPlayer) = LeftSpinnerCount(CurrentPlayer) + 1
        CheckDisco
    End If
    If bDiscoMBEnabled Then 'increase the Disco Super Jackpot
        DiscoSuperJackpot(CurrentPlayer) = DiscoSuperJackpot(CurrentPlayer) + 100000
        DMD " DISCO SUPERJACKPOT", CL(1, "IS " &FormatScore(DiscoSuperJackpot(CurrentPlayer))), "", eNone, eNone, eNone, 200, True, ""
    End if
End Sub

Sub Spinner002_Spin 'wolverine
    DOF 124, DOFPulse
    SoundSpinner Spinner002
    If Tilted Then Exit Sub
    If bBerserker Then
        Addscore BerserkerValue(CurrentPlayer)
        BerserkerValue(CurrentPlayer) = BerserkerValue(CurrentPlayer) + 10000
    Else
        Addscore 1000
        CheckBerserker                 'check for the number of hits to start Berserker rage
    End If
    If bLilDPMB AND LilDPHits> 14 Then 'second part of the LilDP multiball
        Addscore 500000
    End If
    If bDiscoMBEnabled Then 'increase the Disco Super Jackpot
        DiscoSuperJackpot(CurrentPlayer) = DiscoSuperJackpot(CurrentPlayer) + 100000
        DMD " DISCO SUPERJACKPOT", CL(1, "IS " &FormatScore(DiscoSuperJackpot(CurrentPlayer))), "", eNone, eNone, eNone, 200, True, ""
    End if
    Select Case Battle(CurrentPlayer, 0)
        Case 2 'Mystique
            Addscore DominoValue
    End Select
End Sub

'****************
' Top lane Gates
'****************

Sub OpenGates
    gate3.open = True
    gate4.open = True
End Sub

Sub CloseGates
    gate3.open = False
    gate4.open = False
End Sub

'*****************
'    Hell House
'*****************

Sub HellHouse_Hit
    PlaySoundAt "VUKEnter", HellHouse
    If bSkillShotReady Then ResetSkillShotTimer_Timer
    ' check for modes
    StopCombo
    Addscore 25000
    HellHouse.destroyball
    BallsinHole = BallsInHole + 1
    li056.State = 0
    Flashforms flasher012, 800, 50, 0
    DOF 128, DOFPulse
    ' check modes
    If bChooseBattle Then Exit Sub 'do nothing if you already are selecing a battle, due to multiball
    If bMechSuitMBSJackpot Then
        SuperJackpot(CurrentPLayer) = MechMBJackpot(CurrentPlayer) * 6
        AwardSuperJackpot
        MechMBJackpot(CurrentPlayer) = MechMBJackpot(CurrentPlayer) + 1000000 'increase by 1 million
        ContinueMechSuitMB
        Exit Sub
    End If
    If bMechSuitMBLight(CurrentPlayer)Then 'Start Mech Suit Multiball
        StartMechSuitMB
    pupevent 838
        Exit Sub
    End If
    If bBattleReady AND bMechSuitMBStarted = False then
        StartChooseBattle 'choose battle will kick out the ball after choosing a battle
        bBattleReady = False
        PlaySong "mu_shooterlane"
        Exit Sub
    End If
    If bEndBattleJackpot Then
        bEndBattleJackpot = False
        Jackpot(CurrentPLayer) = 3000000:AwardJackpot
    End If
    If li041.State Then 'weapon jackpot
        li041.State = 0
        DMD CL(0, "WEAPONS JACKPOT"), CL(1, "1 MILLION"), "", eNone, eBlink, eNone, 2000, True, "fanfare"
        Addscore 1000000
        RelitWeapons
    End If
    If li038.State Then 'extra ball
        li038.State = 0
        AwardExtraBall
    End If
    If li037.State Then 'mystery
        li037.State = 0 : pupevent 847
        StartMystery
        Exit Sub 'mystery will kick the ball out after it is finished
    End If
    ' Nothing left to do, so kick out the ball
    Flashforms li070, 1500, 50, 0
    vpmtimer.addtimer 1500, "kickBallOut '"
End Sub

Sub kickBallOut
    If BallsinHole> 0 Then
        BallsinHole = BallsInHole - 1
        '    PlaySoundAt "fx_popper", HellHouse
        PlaySoundAt SoundFXDOF("VUKOut", 111, DOFPulse, DOFcontactors), HellHouse
        DOF 121, DOFPulse
        HellHouse.CreateSizedBallWithMass BallSize / 2, BallMass
        HellHouse.kick 164, 50, 1
        Flashforms li070, 500, 50, 0
        Flashforms flasher012, 500, 50, 0
        DOF 129, DOFPulse
        vpmtimer.addtimer 350, "kickBallOut '" 'kick out the rest of the balls, if any
    End If
End Sub

'************************
'   DISCO lights & Modes
'************************
Sub StartDisco
    DOF 141, DOFOn
    discolight.visible = 1
    DiscoTimer.Enabled = 1
    bDiscoEnabled = True
  pupevent 844
End Sub

Sub StopDisco
    DOF 141, DOFOff
    discolight.visible = 0
    DiscoTimer.Enabled = 0
    bDiscoEnabled = False
End Sub

Sub DiscoTimer_Timer
    discoball.rotz = (discoball.rotz + 1)mod 360
    discolight.rotz = (discolight.rotz + 1)mod 360
End Sub

Sub CheckDisco
    Select Case LeftSpinnerCount(CurrentPlayer)
        Case 50 'Disco Loops
            DMD "_", CL(1, "DISCO LOOPS"), "", eNone, eNone, eNone, 1500, True, "vo_discoloops" : pupevent 878
            StartDisco
            StartSpots
            OpenGates
            DiscoLoopsTimer.Interval = 20000 + DazzlerValue * 1000
            DiscoLoopsTimer.Enabled = 1
            bDiscoLoopsEnabled = True
            ChangeSong
            li053.State = 2
        Case 100                                'Disco MB
            DMD "_", CL(1, "DISCO MULTIBALL"), "", eNone, eNone, eNone, 1500, True, "vo_discoloops" : pupevent 869
            LeftSpinnerCount(CurrentPlayer) = 0 'and start again counting towards Disco Loops
            StartDisco
            StartSpots
            bDiscoMBEnabled = True
            ChangeSong
            li053.State = 2
            SetlightColor li057, red, 2
            SetlightColor li058, red, 0
            AddMultiball 1
    End Select
End Sub

Sub DiscoLoopsTimer_Timer 'the discolopp is over
    DiscoLoopsTimer.Enabled = 0
    StopDisco
    StopSpots
    bDiscoLoopsEnabled = False
    ChangeSong
    If Battle(CurrentPlayer, 0) = 3 OR Battle(CurrentPlayer, 0) = 0 Then CloseGates 'Sabretooth because of the bumpers
    li053.State = 0
End Sub

Sub StopDiscoMB 'stop the Disco Multiball
    StopDisco
    StopSpots
    bDiscoMBEnabled = False
    ChangeSong
    li053.State = 0
    li057.State = 0
    li058.State = 0
End Sub

'**************
'   SWORD
'***************

Sub SwordEffect(n)
    Select Case n
        Case 0 ' all off
            For each x in aSwordLights:x.State = 0:Next
        Case 1 'updown ball lock activated
            For each x in aSwordLights:x.BlinkInterval = 250:x.BlinkPattern = 10:x.State = 2:Next
        Case 2 'all blink when a ball is locked
            LightSeqSword.UpdateInterval = 20
            LightSeqSword.Play SeqRandom, 7, , 1000
        Case 3 'top down eject balls
            LightSeqSword.UpdateInterval = 8
            LightSeqSword.Play SeqDownOn, 7, 1
        Case 4
            li062.BlinkPattern = "000000100000"
            li063.BlinkPattern = "000001010000"
            li064.BlinkPattern = "000010001000"
            li065.BlinkPattern = "000100000100"
            li066.BlinkPattern = "001000000010"
            li067.BlinkPattern = "010000000001"
            li068.BlinkPattern = "100000000000"
            For each x in aSwordLights:x.BlinkInterval = 125:x.State = 2:Next
    End Select
End Sub

'*******************
' Rotate Lane Lights
'*******************
' BAM! lights and BOOM lights

Sub RotateLaneLights(n) 'n is the direction, 1 or 0, left or right
    Dim tmp
    If bRotateLights Then
        If n = 1 Then
            tmp = li049.State
            li049.State = li050.State
            li050.State = li051.State
            li051.State = li052.State
            li052.State = tmp
            tmp = li006.State
            li006.State = li014.State
            li014.State = li015.State
            li015.State = li016.State
            li016.State = tmp
            'rotate the Boom array too
            tmp = Boom(CurrentPlayer, 1)
            Boom(CurrentPlayer, 1) = Boom(CurrentPlayer, 2)
            Boom(CurrentPlayer, 2) = Boom(CurrentPlayer, 3)
            Boom(CurrentPlayer, 3) = Boom(CurrentPlayer, 4)
            Boom(CurrentPlayer, 4) = tmp
        Else
            tmp = li052.State
            li052.State = li051.State
            li051.State = li050.State
            li050.State = li049.State
            li049.State = tmp
            tmp = li016.State
            li016.State = li015.State
            li015.State = li014.State
            li014.State = li006.State
            li006.State = tmp
            'rotate the Boom array too
            tmp = Boom(CurrentPlayer, 4)
            Boom(CurrentPlayer, 4) = Boom(CurrentPlayer, 3)
            Boom(CurrentPlayer, 3) = Boom(CurrentPlayer, 2)
            Boom(CurrentPlayer, 2) = Boom(CurrentPlayer, 1)
            Boom(CurrentPlayer, 1) = tmp
        End If
    End If
End Sub

Sub UpdateGates(n) '1 is open 0 is closed
    If bskillshotready then
        Gate3.Open = n
        Gate4.Open = n
    End If
End Sub

'*********************
'      BATTLES
'*********************
' Battles((CurrentPlayer), x) = value
' x being from 1 to 7, this is 7 battles
' x = 0 current running battle number
' values 0: not started, 1 finished, 2 ready to start
' 7 Battles
' 5 Battles to choose
' complete 3 to battle Sauron
' complete 5 to battle Mr.Sinister

Sub StartChooseBattle
    If NOT bChooseBattle Then
        DMD " CHOOSE YOUR BATTLE", "", "", eNone, eNone, eNone, 2000, True, "vo_chooseyourbattle" : pupevent 830
        bChooseBattle = True
        If BattlesWon(CurrentPlayer)> 2 Then BattlesToChoose = 6
        If BattlesWon(CurrentPlayer)> 5 Then BattlesToChoose = 7
        vpmtimer.addtimer 2000, "UpdateDMDBattle '"
    End If
End Sub

Sub ChooseBattle(keycode) '5 first battles
    If(keycode = PlungerKey OR keycode = StartGameKey OR keycode = Lockbarkey)AND LifeLeft(CurrentPlayer, Battlenr)> 0 Then
        bChooseBattle = False
        StartBattle Battlenr
    End If
    If keycode = LeftFlipperKey Then
        Battlenr = (Battlenr - 1)
        If Battlenr <1 Then Battlenr = BattlesToChoose
        UpdateDMDBattle
    End If
    If keycode = RightFlipperKey Then
        Battlenr = (Battlenr + 1)
        If Battlenr> BattlesToChoose Then Battlenr = 1
        UpdateDMDBattle
    End If
End Sub

Sub UpdateDMDBattle
    Dim tmp
    tmp = ShowLife(LifeLeft(CurrentPlayer, Battlenr))
    Select Case Battlenr
        Case 1:DMDFlush:DMD "      JUGGERNAUT", tmp, "d_juggernaut", eNone, eNone, eNone, 10000, True, "vo_juggernaut" : pupevent 831
        case 2:DMDFlush:DMD "       MYSTIQUE", tmp, "d_mystique", eNone, eNone, eNone, 10000, True, "vo_mystique" : pupevent 832
        Case 3:DMDFlush:DMD "      SABRETOOTH", tmp, "d_sabretooth", eNone, eNone, eNone, 10000, True, "vo_sabretooth" : pupevent 833
        Case 4:DMDFlush:DMD "        T-REX", tmp, "d_trex", eNone, eNone, eNone, 10000, True, "vo_trex" : pupevent 834
        Case 5:DMDFlush:DMD "       MEGALODON", tmp, "d_megalodon", eNone, eNone, eNone, 10000, True, "vo_megalodon" : pupevent 835
        Case 6:DMDFlush:DMD "         SAURON", tmp, "d_sauron", eNone, eNone, eNone, 10000, True, "vo_sauron" : pupevent 836
        Case 7:DMDFlush:DMD "     MR-SINISTER", tmp, "d_mrsinister", eNone, eNone, eNone, 10000, True, "vo_mrsinister" : pupevent 837
    End Select
End Sub

Sub StartBattle(n)
    DMDFlush
    TurnOffTeamUpLights
    Battle(CurrentPlayer, 0) = n
    Battle(CurrentPlayer, n) = 2 '0 battle not started, 1 battle finished, 2 battle started
    If bFirstBattle(CurrentPlayer)Then
        AttackPower = 1
    Else
        AttackPower = 0.5 'the battles are now double as hard to complete
    End If
    TurnOffArrows
    ChangeSong
    Select Case Battle(CurrentPlayer, 0)
        Case 1 'Juggernaut
            DMD "  SHOOT DP TARGETS", "TO ATTACK JUGGERNAUT", "d_juggernaut", eNone, eNone, eNone, 3000, True, "vo_shootflashingshots" : pupevent 831
            If LifeLeft(CurrentPlayer, 1)> 6 Then
                LightSeqDPtargets.Play SeqRandom, 40, , 4000
            Else
                BattleLights(2) = 2:BattleLights(5) = 2
            End If
            Flasher003.visible = 1
            OpenGates
        Case 2 'Mystique
            DMD " SHOOT WHITE SHOTS", "  TO ATTACK MYSTIQUE", "d_mystique", eNone, eNone, eNone, 3000, True, "vo_shoot-mystique" : pupevent 832
            BattleLights(4) = 2
            Flasher004.visible = 1
            OpenGates
        Case 3 'Sabretooth
            DMD " SHOOT POP BUMPERS", "TO ATTACK SABRETOOTH", "d_sabretooth", eNone, eNone, eNone, 3000, True, "vo_shoot-sabretooth" : pupevent 833
            BattleLights(1) = 2:BattleLights(6) = 2
            li001.State = 2
            Flasher005.visible = 1
            CloseGates
        Case 4 'T-Rex
            DMD " SHOOT GREEN SHOTS", " TO ATTACK T-REX", "d_trex", eNone, eNone, eNone, 3000, True, "vo_shootgreenshots-trex" : pupevent 834
            Select case RndNbr(6)
                Case 1:BattleLights(1) = 2:BattleLights(2) = 2
                Case 2:BattleLights(2) = 2:BattleLights(4) = 2
                Case 3:BattleLights(4) = 2:BattleLights(5) = 2
                Case 4:BattleLights(5) = 2:BattleLights(6) = 2
                Case 5:BattleLights(6) = 2:BattleLights(3) = 2
                Case 6:BattleLights(3) = 2:BattleLights(1) = 2
            End Select
            Flasher002.visible = 1
            OpenGates
        Case 5 'Megalodon
            DMD "   SHOOT RED SHOTS", " TO ATTACK MEGALODON", "d_megalodon", eNone, eNone, eNone, 3000, True, "vo_shootflashingshots" : pupevent 835
            For each x in aArrows
                SetLightColor x, blue, 1
            Next
            BattleLights(1) = 1
            BattleLights(2) = 1
            BattleLights(3) = 1
            BattleLights(4) = 1
            BattleLights(5) = 1
            BattleLights(6) = 1
            Select case RndNbr(6)
                Case 1:BattleLights(1) = 2
                Case 2:BattleLights(2) = 2
                Case 3:BattleLights(3) = 2
                Case 4:BattleLights(4) = 2
                Case 5:BattleLights(5) = 2
                Case 6:BattleLights(6) = 2
            End Select
            OpenGates
        Case 6 'Sauron
            DMD " SHOOT GREEN SHOTS", " TO ATTACK SAURON", "d_sauron", eNone, eNone, eNone, 3000, True, "vo_shootgreenshots-sauron" : pupevent 836
            AddMultiball 2
            BattleLights(2) = 2
            BattleLights(5) = 2
            OpenGates
        Case 7 'MrSinister
            DMD "SHOOT FLASHING SHOTS", "TO ATTACK MRSINISTER", "d_mrsinister", eNone, eNone, eNone, 3000, True, "vo_shootflashingshots-ifyoudare" : pupevent 837
            Select case RndNbr(4)
                Case 1:BattleLights(1) = 2
                Case 2:BattleLights(4) = 2
                Case 3:BattleLights(6) = 2
                Case 4:BattleLights(3) = 2
            End Select
            AddMultiball 1
            OpenGates
            MrSinisterTimer.Interval = 120 + dazzlerValue
            MrSinisterTimer.Enabled = 1
    End Select
    Flashforms li070, 3000, 50, 0
    vpmtimer.addtimer 3000, "KickBallOut '"
End Sub

Sub LightSeqDPtargets_PlayDone()
    LightSeqDPtargets.Play SeqRandom, 40, , 4000
End Sub

Sub CheckBattle                                     'called after each target or lane hit to change lights and check for the end of the battle
DMD "", "", "d_bam", eNone, eNone, eBlink, 1000, True, "sfx19"
    Select Case Battle(CurrentPlayer, 0)
        Case 1                                      'Juggernaut
            LifeLeft(CurrentPlayer, 1) = LifeLeft(CurrentPlayer, 1)- AttackPower * WolverineValue
            If LifeLeft(CurrentPlayer, 1) <= 0 Then 'life is empty, then enabled the kicker to finish the battle
                TurnOffArrows
                SetLightColor li056, red, 2
                WinBattle
            ElseIF LifeLeft(CurrentPlayer, 1) <7 Then 'end the target hits and start the lane hits
                LightSeqDPtargets.StopPlay
                If BattleLights(3) = 2 Then
                    TurnOffArrows
                    BattleLights(2) = 2:BattleLights(5) = 2
                Else
                    TurnOffArrows
                    BattleLights(1) = 2:BattleLights(4) = 2:BattleLights(6) = 2
                End If
            End If
        Case 2                                      'Mystique
            LifeLeft(CurrentPlayer, 2) = LifeLeft(CurrentPlayer, 2)- AttackPower * WolverineValue
            If LifeLeft(CurrentPlayer, 2) <= 0 Then 'life is empty, then enabled the kicker to finish the battle
                TurnOffArrows
                SetLightColor li056, red, 2
                WinBattle
            ElseIF LifeLeft(CurrentPlayer, 2) <8 Then
                BattleLights(2) = 2:BattleLights(4) = 2:BattleLights(5) = 2
            ElseIF LifeLeft(CurrentPlayer, 2) <10 Then
                BattleLights(2) = 2:li071.State = 0:BattleLights(5) = 2
            End If
        Case 3                                      'Sabretooth
            LifeLeft(CurrentPlayer, 3) = LifeLeft(CurrentPlayer, 3)-(AttackPower / 4) * WolverineValue
            If LifeLeft(CurrentPlayer, 3) <= 0 Then 'life is empty, then enabled the kicker to finish the battle
                TurnOffArrows
                SetLightColor li056, red, 2
                WinBattle
            End If
        Case 4                                      'T-Rex
            LifeLeft(CurrentPlayer, 4) = LifeLeft(CurrentPlayer, 4)- AttackPower * WolverineValue
            If LifeLeft(CurrentPlayer, 4) <= 0 Then 'life is empty, then enabled the kicker to finish the battle
                TurnOffArrows
                SetLightColor li056, red, 2
                WinBattle
            ElseIF LifeLeft(CurrentPlayer, 4) <10 Then 'change the green arrow
                TurnOffArrows
                Select case RndNbr(6)
                    Case 1:BattleLights(1) = 2:BattleLights(2) = 2
                    Case 2:BattleLights(2) = 2:BattleLights(4) = 2
                    Case 3:BattleLights(4) = 2:BattleLights(5) = 2
                    Case 4:BattleLights(5) = 2:BattleLights(6) = 2
                    Case 5:BattleLights(6) = 2:BattleLights(3) = 2
                    Case 6:BattleLights(3) = 2:BattleLights(1) = 2
                End Select
            End If
        Case 5                                      'Megalodon
            LifeLeft(CurrentPlayer, 5) = LifeLeft(CurrentPlayer, 5)- AttackPower * WolverineValue
            If LifeLeft(CurrentPlayer, 5) <= 0 Then 'life is empty, then enabled the kicker to finish the battle
                turnOffarrows
                SetLightColor li056, red, 2
                WinBattle
            ElseIF LifeLeft(CurrentPlayer, 5) <10 Then 'change the red blinking arrow
                For each x in aArrows
                    SetLightColor x, blue, 1
                Next
                BattleLights(1) = 1
                BattleLights(2) = 1
                BattleLights(3) = 1
                BattleLights(4) = 1
                BattleLights(5) = 1
                BattleLights(6) = 1
                Select case RndNbr(6)
                    Case 1:BattleLights(1) = 2
                    Case 2:BattleLights(2) = 2
                    Case 3:BattleLights(3) = 2
                    Case 4:BattleLights(4) = 2
                    Case 5:BattleLights(5) = 2
                    Case 6:BattleLights(6) = 2
                End Select
            End If
        Case 6                                      'Sauron
            LifeLeft(CurrentPlayer, 6) = LifeLeft(CurrentPlayer, 6)- AttackPower * WolverineValue
            If LifeLeft(CurrentPlayer, 6) <= 0 Then 'life is empty, then enabled the kicker to finish the battle
                turnOffarrows
                SetLightColor li056, red, 2
                WinBattle
            ElseIF LifeLeft(CurrentPlayer, 5) <3 Then 'enable superjackpot at the colossus target
                li069.State = 2
            ElseIF LifeLeft(CurrentPlayer, 5) <7 Then 'enable superjackpot at the colossus target
                li069.State = 2
            End If
        Case 7                                      'MrSinister
            LifeLeft(CurrentPlayer, 7) = LifeLeft(CurrentPlayer, 7)- AttackPower * WolverineValue
            If LifeLeft(CurrentPlayer, 7) <= 0 Then 'life is empty, then enabled the kicker to finish the battle
                turnOffarrows
                SetLightColor li056, red, 2
                WinBattle
            ElseIF LifeLeft(CurrentPlayer, 7) <10 Then 'change the darkblue blinking arrow
                TurnOffArrows
                Select case RndNbr(4)
                    Case 1:BattleLights(1) = 2
                    Case 2:BattleLights(4) = 2
                    Case 3:BattleLights(6) = 2
                    Case 4:BattleLights(3) = 2
                End Select
            End If
    End Select
End Sub

Sub StopBattle 'stops the battle, mostly when you loose the ball, it can be continued
    bEndBattleJackpot = False
    TurnOffArrows
    li056.State = 0
    li070.State = 0
    DMDScoreNow
    Select Case Battle(CurrentPlayer, 0)
        Case 1 'Juggernaut
            LightSeqDPtargets.StopPlay
            Battle(CurrentPlayer, 1) = 0
            Flasher003.visible = 0
        Case 2 'Mystique
            Battle(CurrentPlayer, 2) = 0
            Flasher004.visible = 0
        Case 3 'Sabretooth
            Battle(CurrentPlayer, 3) = 0
            Flasher005.visible = 0
        Case 4 'T-Rex
            Battle(CurrentPlayer, 4) = 0
            Flasher002.visible = 0
        Case 5 'Megalodon
            Battle(CurrentPlayer, 5) = 0
            Flasher006.visible = 0
        Case 6 'Sauron
            Battle(CurrentPlayer, 6) = 0
        Case 7 'MrSinister
            Battle(CurrentPlayer, 7) = 0
            ResetForNewRound
            MrSinisterTimer.Enabled = 0
    End Select
    ResetTeamUps
    Battle(CurrentPlayer, 0) = 0
    CloseGates
End Sub

Sub WinBattle
    BattlesWon(CurrentPlayer) = BattlesWon(CurrentPlayer) + 1
    Jackpot(CurrentPlayer) = 500000 * BattlesWon(CurrentPlayer)
    Select Case Battle(CurrentPlayer, 0)
        Case 1 'Juggernaut
            Battle(CurrentPlayer, 1) = 1
            Flasher003.visible = 0
            DMD "         JUGGERNAUT", "          DEFEATED", "d_juggernaut", eBlink, eNone, eNone, 2000, False, "vo_juggernautdefeated" : pupevent 870
        Case 2 'Mystique
            Battle(CurrentPlayer, 2) = 1
            Flasher004.visible = 0
            DMD "          MYSTIQUE", "          DEFEATED", "d_mystique", eBlink, eNone, eNone, 2000, False, "vo_mystiquedefeated" : pupevent 871
        Case 3 'Sabretooth
            Battle(CurrentPlayer, 3) = 1
            li001.State = 0
            Flasher005.visible = 0
            DMD "         SABRETOOTH", "          DEFEATED", "d_sabretooth", eBlink, eNone, eNone, 2000, False, "vo_sabretoothdefeated" : pupevent 872
        Case 4 'T-Rex
            Battle(CurrentPlayer, 4) = 1
            DMD "           T-REX", "          DEFEATED", "d_trex", eBlink, eNone, eNone, 2000, False, "vo_trexdefeated" : pupevent 873
            Flasher002.visible = 0
        Case 5 'Megalodon
            Battle(CurrentPlayer, 5) = 1
            Flasher006.visible = 0
            DMD "          MEGALODON", "          DEFEATED", "d_megalodon", eBlink, eNone, eNone, 2000, False, "vo_megalodondefeated" : pupevent 874
        Case 6 'Sauron
            Battle(CurrentPlayer, 6) = 1
            DMD "           SAURON", "          DEFEATED", "d_sauron", eBlink, eNone, eNone, 2000, False, "vo_saurondefeated" : pupevent 875
        Case 7 'MrSinister
            Jackpot(CurrentPlayer) = 1000000 * BattlesWon(CurrentPlayer)
            Battle(CurrentPlayer, 7) = 1
            DMD "         MRSINISTER", "          DEFEATED", "d_mrsinister", eBlink, eNone, eNone, 2000, False, "vo_mrsinisterdefeated" : pupevent 876
            MrSinisterTimer.Enabled = 0
            ResetForNewRound
            'turn on the Special light
            li039.State = 2
    End Select
    ResetTeamUps
    Battle(CurrentPlayer, 0) = 0
    bEndBattleJackpot = True
    DMD "  SHOOT THE SCOOP", " FOR EXTRA JACKPOT", "", eBlink, eNone, eNone, 2500, True, "vo_shootthescoop"
    li070.BlinkInterval = 300:li070.State = 2
    CloseGates
    ChangeSong
End Sub

Sub ResetForNewRound 'reset battles after finishing them
    Dim j
    bFirstBattle(CurrentPlayer) = False
    For j = 0 to 7
        Battle(CurrentPlayer, j) = 0
        LifeLeft(CurrentPlayer, j) = 10
    Next
End Sub

Sub MrSinisterTimer_Timer 'timed battle
    StopBattle
    MrSinisterTimer.Enabled = 0
End Sub

'***********************
' spotflasher animation
'***********************

Dim MyPi, SpotStep, SpotDir
Dim sRGBStep, sRGBFactor, sRed, sGreen, sBlue

Sub StartSpots
    Spot1.visible = 1
    Spot2.visible = 1
    MyPi = Round(4 * Atn(1), 6) / 90
    SpotStep = 0
    sRGBStep = 0
    sRGBFactor = 5
    sRed = 255
    sGreen = 0
    sBlue = 0
    Spots.Enabled = 1
End Sub

Sub StopSpots
    Spot1.visible = 0
    Spot2.visible = 0
    Spots.Enabled = 0
    Spot1.RotZ = 210
    Camera1.RotZ = 210
    Spot2.RotZ = 170
    Camera2.RotZ = 170
    Spot1.color = RGB(255, 252, 224)
    Spot2.color = RGB(255, 252, 224)
End Sub

Sub Spots_Timer()
    Spot1.visible = 1
    Spot2.visible = 1
    'rotate spots
    SpotDir = SIN(SpotStep * MyPi) * 50
    SpotStep = (SpotStep + 1)MOD 360
    Spot1.RotZ = 210 - SpotDir
    Camera1.RotZ = 210 - SpotDir
    Spot2.RotZ = 170 + SpotDir
    Camera2.RotZ = 170 + SpotDir
    ' color the spotlights
    Select Case sRGBStep
        Case 0 'Green
            sGreen = sGreen + sRGBFactor
            If sGreen> 255 then
                sGreen = 255
                sRGBStep = 1
            End If
        Case 1 'Red
            sRed = sRed - sRGBFactor
            If sRed <0 then
                sRed = 0
                sRGBStep = 2
            End If
        Case 2 'Blue
            sBlue = sBlue + sRGBFactor
            If sBlue> 255 then
                sBlue = 255
                sRGBStep = 3
            End If
        Case 3 'Green
            sGreen = sGreen - sRGBFactor
            If sGreen <0 then
                sGreen = 0
                sRGBStep = 4
            End If
        Case 4 'Red
            sRed = sRed + sRGBFactor
            If sRed> 255 then
                sRed = 255
                sRGBStep = 5
            End If
        Case 5 'Blue
            sBlue = sBlue - sRGBFactor
            If sBlue <0 then
                sBlue = 0
                sRGBStep = 0
            End If
    End Select
    Spot1.color = RGB(sRed, sGreen, sBlue)
    Spot2.color = RGB(sRed, sGreen, sBlue)
End Sub

'*******************
'      COMBOS
'*******************
' don't time out
' starts at 500K for a 2 way combo and it is doubled on each combo
' shots that count as combos:
' Left Ramp 1
' Wolverine loop  2
' Right Ramp  3
'   5 combos, supercombo and superdupercombo

Sub CheckCombo(n)
    If Battle(CurrentPlayer, 0) <> 0 then Exit Sub             'no combos during battles
    If n = OldCombo Then StopCombo:Exit Sub                   'repeated shot so stop the combos

    If n = "Trigger013" AND LastSwitchHit = "Trigger010" Then 'NinjaApocalypse
        DMD CL(0, "NINJA"), CL(1, "APOCAYPSE"), "d_ninjamb", eNone, eNone, eNone, 1500, True, "vo_supercombo" : pupevent 854
        For x = 1 to 10
            CheckNinjaStars
        Next
        Addscore NinjaApocJackpot(CurrentPLayer)
        NinjaApocJackpot(CurrentPLayer) = NinjaApocJackpot(CurrentPLayer) + 500000
        StopCombo
    Else
        OldCombo = n
        ComboCount = ComboCount + 1
        Select Case ComboCount
            Case 1: 'just starting
            Case 2:DMD "", "", "d_combo", eNone, eNone, eBlink, 1500, True, "vo_combo" : pupevent 849
            Case 3:DMD "", "", "d_combo2", eNone, eNone, eBlink, 1500, True, "vo_2xcombo" : pupevent 850
            Case 4:DMD "", "", "d_combo3", eNone, eNone, eBlink, 1500, True, "vo_3xcombo" : pupevent 851
            Case 5:DMD "", "", "d_combo4", eNone, eNone, eBlink, 1500, True, "vo_4xcombo" : pupevent 852
            Case 6:DMD "", "", "d_combo5", eNone, eNone, eBlink, 1500, True, "vo_5xcombo" : pupevent 853
            Case 7:DMD CL(0, "SUPER"), CL(1, "COMBO"), "", eNone, eBlink, eNone, 1500, True, "vo_supercombo" : pupevent 854
            Case Else:DMD CL(0, "SUPERDUPER"), CL(1, "COMBO"), "", eNone, eBlink, eNone, 1500, True, "vo_superdupercombo" : pupevent 855
        End Select
        AddScore ComboValue * ComboCount
    End If
End Sub

Sub StopCombo
    ComboCount = 0
    ComboValue = 500000
    OldCombo = ""
End Sub

'************
'  Team ups
'************

'reset team-ups
Sub ResetTeamUps
    If DazzlePower(CurrentPlayer) = 4 Then DazzlePower(CurrentPlayer) = 0:DazzlerValue = 0
    If ColossusPower(CurrentPlayer) = 4 Then ColossusPower(CurrentPlayer) = 0:ColossusValue = 1
    If WolverinePower(CurrentPlayer) = 4 Then WolverinePower(CurrentPlayer) = 0:WolverineValue = 1
    If DominoPower(CurrentPlayer) = 4 Then DominoPower(CurrentPlayer) = 0:DominoValue = 0
    Flasher013.Visible = 0
    Flasher014.Visible = 0
    Flasher016.Visible = 0
    Flasher017.Visible = 0
End Sub

'dazzler
Sub CheckDazzler
    If DazzlePower(CurrentPlayer) = 4 Then
        DMD "  DAZZLER", "  TEAM-UP", "d_dazzler", eNone, eNone, eNone, 1500, True, "vo_dazzlerteamup" : pupevent 856
        DazzlerValue = 20 'add 20 extra seconds to any timer or hurry up
        FlashForms Flasher014, 1500, 50, 1
        DOF 132, DOFPulse
    End If
    If DazzlePower(CurrentPlayer)> 4 Then DazzlePower(CurrentPlayer) = 4
End Sub

'colossus
Sub CheckColossus
    If ColossusPower(CurrentPlayer) = 4 Then
        DMD " COLOSSUS", "  TEAM-UP", "d_colossus", eNone, eNone, eNone, 1500, True, "vo_colossusteamup" : pupevent 857
        ColossusValue = 2 'doubles points during battles
        FlashForms Flasher017, 1500, 50, 1
        DOF 133, DOFPulse
    End If
    If ColossusPower(CurrentPlayer)> 4 Then ColossusPower(CurrentPlayer) = 4
End Sub

'wolverine
Sub CheckWolverine
    If WolverinePower(CurrentPlayer) = 4 Then
        DMD " WOLVERINE", "  TEAM-UP", "d_wolverine", eNone, eNone, eNone, 1500, True, "vo_wolverineteamup" : pupevent 858
        WolverineValue = 2 'doubles damage during battles
        FlashForms Flasher013, 1500, 50, 1
        DOF 134, DOFPulse
    End If
    If WolverinePower(CurrentPlayer)> 4 Then WolverinePower(CurrentPlayer) = 4
End Sub

'domino
Sub CheckDomino
    If DominoPower(CurrentPlayer) = 4 Then
        DMD " DOMINO", "  TEAM-UP", "d_domino", eNone, eNone, eNone, 1500, True, "vo_dominoteamup" : pupevent 859
        DominoValue = 10000 '10000 extra points during some battles
        FlashForms Flasher016, 1500, 50, 1
        DOF 135, DOFPulse
    End If
    If DominoPower(CurrentPlayer)> 4 Then DominoPower(CurrentPlayer) = 4
End Sub

'*************
'  Weapons x
'*************
'30 weapons = light extra ball
'75 weapons = light MechSuit multiball
'50 weapons = playfield X + 1

Sub Checkweapons                                            'increase and check the number of weapons collected
    Weapons(CurrentPlayer) = Weapons(CurrentPlayer) + 1
    If Weapons(CurrentPlayer)MOD 6 = 0 Then li041.State = 2 'jackpot at the Hell house scoop
    If Weapons(CurrentPlayer)MOD 30 = 0 Then                'lit extra ball
        li038.State = 2
        DMD "_", CL(1, "EXTRA BALL IS LIT"), "", eNone, eNone, eNone, 2000, True, "vo_shootscoop-extraball" : pupevent 860
    End If
    If Weapons(CurrentPlayer)MOD 75 = 0 Then bMechSuitMBLight(CurrentPlayer) = True 'lit Mech Suit MB, starts at the Hellhouse
    If Weapons(CurrentPlayer)MOD 50 = 0 Then AddPlayfieldMultiplier 1
    FlashForms Flasher007, 1500, 50, 0
    DOF 130, DOFPulse
End Sub

Sub RelitWeapons
    Dim i
    For i = 1 to 6:WeaponLights(CurrentPlayer, i) = 1:Next
End Sub
'********************
' Mech Suit multiball
'********************

Sub StartMechSuitMB
    Dim i
    DMD CL(0, "MECH SUIT"), CL(1, "MULTIBALL"), "", eNone, eNone, eNone, 1500, True, "vo_multiball" : pupevent 838
    vpmtimer.addtimer 4000, "AddMultiball 3 '"
    bMechSuitMBLight(CurrentPlayer) = False:li047.State = 2
    bMechSuitMBStarted = true
    ChangeSong
    For i = 1 to 6:MechSuitLights(i) = 2:Next
    vpmtimer.addtimer 1500, "kickBallOut '"
End Sub

Sub ContinueMechSuitMB 'after super jackpot
    Dim i
    bMechSuitMBSJackpot = False
    bMechSuitMBLight(CurrentPlayer) = False
    li047.BlinkInterval = 250:li047.State = 0:li047.State = 2
    bMechSuitMBStarted = true
    For i = 1 to 6:MechSuitLights(i) = 2:Next
    vpmtimer.addtimer 1500, "kickBallOut '"
End Sub

Sub CheckMechSuitMBHits
    Dim i, tmp
    i = 0:tmp = 0
    Jackpot(Currentplayer) = MechMBJackpot(CurrentPlayer)
    AwardJackpot
    MechMBJackpot(CurrentPlayer) = MechMBJackpot(CurrentPlayer) + 100000 'increment the mech suit mb jackpot
    For i = 1 to 6:tmp = tmp + MechSuitLights(i):next
    If tmp = 0 Then                                                      'turn on the Super Jackpot at HellHouse scoop
        bMechSuitMBSJackpot = true
        li047.BlinkInterval = 125:li047.State = 0:li047.State = 2
    End If
End Sub

Sub StopMechSuitMB
    Dim i
    For i = 1 to 6:MechSuitLights(i) = 0:Next
    bMechSuitMBStarted = False
    ChangeSong
    li047.State = 0
    li054.State = 0
    li057.State = 0
    li055.State = 0
    li071.State = 0
    li058.State = 0
    li059.State = 0
End Sub

'************************
' Ninja Lock & Multiball
'************************

' Uses a virtual lock for easier handling when several players are playing

Sub Lock_Hit
    Dim i, tmp               'time to kick the ball
    tmp = 1500
    IF bNinjaMBSJackpot Then 'Award Ninja MB Super Jackpot
        SuperJackpot(CurrentPlayer) = NinjaMBJackpot(CurrentPlayer) * 6
        NinjaMBJackpot(CurrentPlayer) = NinjaMBJackpot(CurrentPlayer) + 1000000
        AwardSuperJackpot
        'reset normal jackpots
        bNinjaMBSJackpot = False
        li058.State = 0
        SwordEffect 0
        For i = 1 to 6:NinjaMBLights(i) = 2:Next
    ElseIf bLockEnabled Then
        BallsInLock(CurrentPLayer) = BallsInLock(CurrentPLayer) + 1
        SwordEffect 2
        Select Case BallsInLock(CurrentPLayer)
            Case 1:DMD "", "", "d_lock1", eNone, eNone, eNone, 1500, True, "vo_ball1locked" : pupevent 839
            Case 2:DMD "", "", "d_lock2", eNone, eNone, eNone, 1500, True, "vo_ball2locked " : pupevent 840
            Case 3:DMD "", "", "d_lock3", eNone, eNone, eNone, 1500, True, "vo_ball3locked" : pupevent 841
                'Start Ninja Multiball
                DMD CL(0, "NINJA"), CL(1, "MULTIBALL"), "", eNone, eNone, eNone, 1500, True, "vo_multiball" : pupevent 825
                vpmtimer.addtimer 4000, "AddMultiball 2 '"
                BallsInLock(CurrentPLayer) = 0
                bLockEnabled = False
                SwordEffect 0
                bNinjaMB = True
        pupevent 825
                ChangeSong
                NinjaMBJackpot(CurrentPlayer) = 500000 + 50000 * NinjaStars(CurrentPLayer)
                'Turn On the Ninja Jacpot Arrows in a teal color
                For i = 1 to 6:NinjaMBLights(i) = 2:Next
                tmp = 3000
        End Select
    End If
    vpmtimer.addtimer tmp, "ExitLock '"
End Sub

Sub NinjaMBCheckHits
    Dim i, tmp:i = 0:tmp = 0
    Jackpot(Currentplayer) = NinjaMBJackpot(CurrentPlayer)
    AwardJackpot
    For i = 1 to 6:tmp = tmp + NinjaMBLights(i):next
    If tmp = 0 Then 'turn on the Super Jackpot at the right ramp
        bNinjaMBSJackpot = true
        SwordEffect 4
    End If
End Sub

Sub StopNinjaMB 'when loose last multiball
    bNinjaMB = False
    bNinjaMBSJackpot = False
    ChangeSong
    For x = 1 to 6:NinjaMBLights(x) = 0:Next
    li054.State = 0
    li057.State = 0
    li055.State = 0
    li071.State = 0
    li058.State = 0
    li059.State = 0
    SwordEffect 0
End Sub

Sub ExitLock
    SwordEffect 3
    PLaySoundAt "fx_kicker", lock
    Lock.kick 180, 3
End Sub

'*****************
'  Chimichangas
'*****************
' chimichangas will also increase the playfield multiplier

Sub CheckChimichangaHits 'called from the right orbit
    Dim tmp
    ChimichangaHits(CurrentPlayer) = ChimiChangaHits(CurrentPlayer) + 1
    If ChimichangaHits(CurrentPlayer) = ChimichangaHitsNeeded(CurrentPlayer)Then 'spot a chimichanga as an  orange light
Flash1 True
        tmp = RndNbr(6)
        ChimiChangaLights(tmp) = 2
        ChimichangaHits(CurrentPlayer) = 0                                              'reset count and
        ChimichangaHitsNeeded(CurrentPlayer) = ChimichangaHitsNeeded(CurrentPlayer) + 2 'increase the needed hits to spot a new chimichanga
    Else
Flash1 True
        DOF 136, DOFPulse
    End If
End Sub

Sub CollectChimichanga
    DMD CL(0, "COLLECTED ONE"), CL(1, "CHIMICHANGA"), "d_chimi2", eNone, eNone, eNone, 1000, True, "sfx17" : pupevent 861
    DMD "", CL(1, FormatScore(ChimichangaValue(CurrentPlayer))), "d_chimi2", eNone, eBlink, eNone, 1000, True, ""
    Chimichangas(CurrentPlayer) = Chimichangas(CurrentPlayer) + 1
    ChimichangaValue(CurrentPlayer) = ChimichangaValue(CurrentPlayer) + 100000
    Flash1 True
    DOF 137, DOFPulse
    If Chimichangas(CurrentPlayer)MOD 10 = 0 Then AddPlayfieldMultiplier 1
End Sub

'*****************
'  Berserker Rage
'*****************
' spinners timed mode

Sub StartBerserker
    BerserkerTime.Interval = 20000 + dazzlerValue * 1000
    BerserkerTime.Enabled = 1
    bBerserker = True
    ChangeSong
End Sub

Sub StopBerserker
    BerserkerTime.Enabled = 0
    bBerserker = False
    ChangeSong
    BerserkerCount(CurrentPLayer) = 0
End Sub

Sub CheckBerserker
    BerserkerCount(CurrentPlayer) = BerserkerCount(CurrentPlayer) + 1
    If BerserkerCount(CurrentPLayer) = 25 Then 'Start Berserker Rage
        DMD CL(0, "STARTING"), CL(1, "BERSERKER RAGE"), "", eNone, eNone, eNone, 1500, True, "" : pupevent 842
        DMD CL(0, "BERSERKER RAGE"), CL(1, "SHOOT THE SPINNERS"), "", eNone, eNone, eNone, 1500, True, ""
        StartBerserker
    End If
End Sub

Sub BerserkerTime_Timer 'stop the Berserker Rage mode
    StopBerserker
End Sub

'*******************
'  Katanarama Time
'*******************
' orbits and ramps timed mode

Sub StartKatanarama
    KatanaramaTime.Interval = 20000 + dazzlerValue * 1000
    KatanaramaTime.Enabled = 1
    bKatanaramaTime = True
    DiamondLights(CurrentPlayer, 1) = 2
    DiamondLights(CurrentPlayer, 2) = 2
    DiamondLights(CurrentPlayer, 3) = 2
    DiamondLights(CurrentPlayer, 4) = 2
    DiamondLights(CurrentPlayer, 5) = 2
    DiamondLights(CurrentPlayer, 6) = 2
End Sub

Sub StopKatanarama
    KatanaramaTime.Enabled = 0
    bKatanaramaTime = False
    KatanaramaCount(CurrentPlayer) = 0
    'reset diamond lights
    KatanaramaCount(CurrentPlayer) = 0
    DiamondLights(CurrentPlayer, 1) = 2
    DiamondLights(CurrentPlayer, 2) = 0
    DiamondLights(CurrentPlayer, 3) = 0
    DiamondLights(CurrentPlayer, 4) = 0
    DiamondLights(CurrentPlayer, 5) = 0
    DiamondLights(CurrentPlayer, 6) = 0
End Sub

Sub CheckKatanarama 'increases the counter and check the lights to start KatanaramaTime
    KatanaramaCount(CurrentPlayer) = KatanaramaCount(CurrentPlayer) + 1
    Select Case KatanaramaCount(CurrentPlayer)
        Case 1
            DiamondLights(CurrentPlayer, 1) = 1
            DiamondLights(CurrentPlayer, 2) = 2
            DMD "1 DIAMOND COLLECTED", "", "d_diamond1", eNone, eNone, eNone, 1000, True, "fanfare" : pupevent 862
        Case 2
            DiamondLights(CurrentPlayer, 2) = 1
            DiamondLights(CurrentPlayer, 3) = 2
            DMD "2 DIAMONDS COLLECTED", "", "d_diamond2", eNone, eNone, eNone, 1000, True, "fanfare" : pupevent 863
        Case 3
            DiamondLights(CurrentPlayer, 3) = 1
            DiamondLights(CurrentPlayer, 4) = 2
            DMD "3 DIAMONDS COLLECTED", "", "d_diamond3", eNone, eNone, eNone, 1000, True, "fanfare" : pupevent 864
        Case 4
            DiamondLights(CurrentPlayer, 4) = 1
            DiamondLights(CurrentPlayer, 5) = 2
            DMD "4 DIAMONDS COLLECTED", "", "d_diamond4", eNone, eNone, eNone, 1000, True, "fanfare" : pupevent 865
        Case 5
            DiamondLights(CurrentPlayer, 5) = 1
            DiamondLights(CurrentPlayer, 6) = 2
            DMD "5 DIAMONDS COLLECTED", "", "d_diamond5", eNone, eNone, eNone, 1000, True, "fanfare" : pupevent 866
        Case 6 'all diamonds are on, so start katanarama time
            DiamondLights(CurrentPlayer, 6) = 1
            DMD "6 DIAMONDS COLLECTED", "", "d_diamond6", eNone, eNone, eNone, 1000, True, "vo_katanaramaisready" : pupevent 867
            DMD CL(0, "STARTING"), CL(1, "KATANARAMA TIME"), "", eNone, eNone, eNone, 1500, True, "" : pupevent 868
            DMD CL(0, "KATANARAMA TIME"), CL(1, "SHOOT ORBITS-RAMPS"), "", eNone, eNone, eNone, 1500, True, "" : pupevent 868
            StartKatanarama
    End Select
End Sub

Sub KatanaramaTime_Timer
    StopKatanarama
End Sub

'****************
' Mystery award
'****************

Sub StartMystery
    DMD "   MYSTERY AWARD    ", CL(1, "LIGHT LOCK"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "DEADPOOL WUZ HERE"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "SMALL POINTS"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "INFINITE BALLS"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "COLLECT X STARS"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "NOTHING"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "LIGHT EXTRA BALL"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "10 PESETAS"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "COLLECT X WEAPONS"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "SOME POPCORN"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "BIG POINTS"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "5 POINTS"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "LIGHT ADD-A-BALL"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "50 MILLIONS"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "ADVANCE CHIMICHANGA"), "", eNone, eNone, eNone, 100, False, ""
    DMD "   MYSTERY AWARD    ", CL(1, "DEFEAT EVERYONE"), "", eNone, eNone, eNone, 100, False, ""
    vpmtimer.addtimer 3000, "ChooseMysteryAward '"
End Sub

Sub ChooseMysteryAward
    Dim tmp, tmp2
    tmp = RndNbr(20)
    DMDFlush
    Select case tmp
        Case 1
            DMD "   MYSTERY AWARD    ", CL(1, "LIGHT LOCK"), "", eNone, eBlink, eNone, 1000, True, "fanfare"
            DMD "", "", "d_lockislit", eNone, eNone, eBlink, 1500, True, "vo_lockislit"
            bLockEnabled = True
            SwordEffect 1
        Case 3
            DMD "   MYSTERY AWARD    ", CL(1, "COLLECT X STARS"), "", eNone, eBlink, eNone, 1000, True, "fanfare"
            tmp2 = RndNbr(5)
            DMD "   MYSTERY AWARD    ", CL(1, "COLLECTED " &tmp2& " STARS"), "", eNone, eBlink, eNone, 1500, True, ""
            For x = 1 to tmp2
                CheckNinjaStars
            Next
        Case 4
            DMD "   MYSTERY AWARD    ", CL(1, "LIGHT EXTRA BALL"), "", eNone, eBlink, eNone, 2500, True, "vo_extraballislit"
            li038.State = 1
        Case 5
            DMD "   MYSTERY AWARD    ", CL(1, "COLLECT X WEAPONS"), "", eNone, eBlink, eNone, 1000, True, "fanfare"
            tmp2 = RndNbr(10)
            DMD "   MYSTERY AWARD    ", CL(1, "COLLECTED " &tmp2& " WEAPONS"), "", eNone, eBlink, eNone, 1500, True, ""
            For x = 1 to tmp2
                CheckWeapons
            Next
        Case 6, 12
            DMD "   MYSTERY AWARD    ", CL(1, "BIG POINTS"), "", eNone, eBlink, eNone, 1000, True, "fanfare"
            tmp2 = 500000 * RndNbr(5)
            DMD "   MYSTERY AWARD    ", CL(1, FormatScore(tmp2)), "", eNone, eBlink, eNone, 1500, True, ""
            Addscore tmp2
        Case 8
            DMD "   MYSTERY AWARD    ", CL(1, "COLLECT 1 CHIMICHANGA"), "", eNone, eBlink, eNone, 1000, True, ""
            Collectchimichanga
        Case Else
            DMD "   MYSTERY AWARD    ", CL(1, "SMALL POINTS"), "", eNone, eBlink, eNone, 1000, True, "fanfare"
            tmp2 = 50000 * RndNbr(5)
            DMD "   MYSTERY AWARD    ", CL(1, FormatScore(tmp2)), "", eNone, eBlink, eNone, 1500, True, ""
            Addscore tmp2
    End Select
    'kickout the ball
    Flashforms li070, 2500, 50, 0
    vpmtimer.addtimer 2600, "kickBallOut '"
End Sub

'DMD "                    ", "                    ", "", eNone, eNone, eNone, 1000, True, ""

'//////////////////////////////////////////////////////////////////////
'// TIMERS
'//////////////////////////////////////////////////////////////////////


' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
  RollingUpdate         'update rolling sounds
End Sub


' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub


'//////////////////////////////////////////////////////////////////////
'// TargetBounce
'//////////////////////////////////////////////////////////////////////

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

' Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit
  TargetBouncer activeball, 1
End Sub

'//////////////////////////////////////////////////////////////////////
'// PHYSICS DAMPENERS
'//////////////////////////////////////////////////////////////////////

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

    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
      aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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

'//////////////////////////////////////////////////////////////////////
'// TRACK ALL BALL VELOCITIES FOR RUBBER DAMPENER AND DROP TARGETS
'//////////////////////////////////////////////////////////////////////

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

'//////////////////////////////////////////////////////////////////////
'// RAMP ROLLING SFX
'//////////////////////////////////////////////////////////////////////

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
dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)

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

'//////////////////////////////////////////////////////////////////////
'// RAMP TRIGGERS
'//////////////////////////////////////////////////////////////////////

Sub ramptrigger01_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger02_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger02_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger03_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger03_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger03
End Sub



Sub ramptrigger001_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger002_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger002_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger003_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub


Sub ramptrigger0001_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub


'//////////////////////////////////////////////////////////////////////
'// Ball Rolling
'//////////////////////////////////////////////////////////////////////

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
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
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
        BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
      Else
        BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
      End If
      BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
      BallShadowA(b).X = BOT(b).X
      BallShadowA(b).visible = 1
    End If
  Next
End Sub

'//////////////////////////////////////////////////////////////////////
'// Mechanic Sounds
'//////////////////////////////////////////////////////////////////////

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
SpinnerSoundLevel = 0.5                                       'volume level; range [0, 1]

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

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function



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


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


'/////////////////////////////////////////////////////////////////
'         Blades
'/////////////////////////////////////////////////////////////////
Dim Sideblades
Sideblades = 1                      '1 - Image, 0 - Black Steel


If Sideblades = 1 then
cabblades.visible = 0
else
cabblades.visible = 1
End If

'//////////////////////////////////////////////////////////////////////
'// Ball
'//////////////////////////////////////////////////////////////////////

If BallBright Then
  table1.BallImage = "ball_HDR_brighter"
Else
  table1.BallImage = "MRBallDark2b"
End If

'//////////////////////////////////////////////////////////////////////
'// Dynamic Ball Shadows
'//////////////////////////////////////////////////////////////////////

Const fovY          = 0   'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source


' *** This segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' ' stop the sound of deleted balls
' For b = UBound(BOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
'   ...rolling(b) = False
'   ...StopSound("BallRoll_" & b)
' Next
'
'...rolling and drop sounds...

'   If DropCount(b) < 5 Then
'     DropCount(b) = DropCount(b) + 1
'   End If
'
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If BOT(b).Z > 30 Then
'       BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
'     Else
'       BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
'     End If
'     BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
'     BallShadowA(b).X = BOT(b).X
'     BallShadowA(b).visible = 1
'   End If

' *** Required Functions, enable these if they are not already present elswhere in your table
Function DistanceFast(x, y)
  dim ratio, ax, ay
  ax = abs(x)         'Get absolute value of each vector
  ay = abs(y)
  ratio = 1 / max(ax, ay)   'Create a ratio
  ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
  if ratio > 0 then     'Quickly determine if it's worth using
    DistanceFast = 1/ratio
  Else
    DistanceFast = 0
  End if
end Function

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

'Dim PI: PI = 4*Atn(1)

'Function Atn2(dy, dx)
' If dx > 0 Then
'   Atn2 = Atn(dy / dx)
' ElseIf dx < 0 Then
'   If dy = 0 Then
'     Atn2 = pi
'   Else
'     Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
'   end if
' ElseIf dx = 0 Then
'   if dy = 0 Then
'     Atn2 = 0
'   else
'     Atn2 = Sgn(dy) * pi / 2
'   end if
' End If
'End Function
'
'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******
Dim sourcenames, currentShadowCount, DSSources(33), numberofsources, numberofsources_hold
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(12), objrtx2(12)
dim objBallShadow(12)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9,BallShadowA10,BallShadowA11)

DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 0.01
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 0.02
    objrtx2(iii).visible = 0

    currentShadowCount(iii) = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = iii/1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    iii = iii + 1
  Next
  numberofsources = iii
  numberofsources_hold = iii
end sub


Sub DynamicBSUpdate
  Dim falloff:  falloff = 150     'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, currentMat, AnotherSource, BOT, iii
  BOT = GetBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 to tnob
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub    'No balls in play, exit

'The Magic happens now
  For s = lob to UBound(BOT)

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If BOT(s).Z > 30 Then             'The flasher follows the ball up ramps while the primitive is on the pf
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + BallSize/10 + fovY
        objBallShadow(s).visible = 1

        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf, primitive only
        objBallShadow(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + fovY
        BallShadowA(s).visible = 0
      Else                      'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If BOT(s).Z > 30 Then             'In a ramp
        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          BallShadowA(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        BallShadowA(s).Y = BOT(s).Y + Ballsize/10 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/2 + 5
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 Then 'And BOT(s).Y < (TableHeight - 200) Then 'Or BOT(s).Z > 105 Then    'Defining when and where (on the table) you can have dynamic shadows
        For iii = 0 to numberofsources - 1
          LSd=DistanceFast((BOT(s).x-DSSources(iii)(0)),(BOT(s).y-DSSources(iii)(1))) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then               'If the ball is within the falloff range of a light and light is on (we will set numberofsources to 0 when GI is off)
            currentShadowCount(s) = currentShadowCount(s) + 1   'Within range of 1 or 2
            if currentShadowCount(s) = 1 Then           '1 dynamic shadow source
              sourcenames(s) = iii
              currentMat = objrtx1(s).material
              objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01            'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-LSd)/falloff                 'Sets opacity/darkness of shadow by distance to light
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness           'Scales shape of shadow with distance/opacity
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^2,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-ShadowOpacity)
              End If

            Elseif currentShadowCount(s) = 2 Then
                                  'Same logic as 1 shadow, but twice
              currentMat = objrtx1(s).material
              AnotherSource = sourcenames(s)
              objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(AnotherSource)(0),DSSources(AnotherSource)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-DistanceFast((BOT(s).x-DSSources(AnotherSource)(0)),(BOT(s).y-DSSources(AnotherSource)(1))))/falloff
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

              currentMat = objrtx2(s).material
              objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
  '           objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx2(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity2 = (falloff-LSd)/falloff
              objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2))
              End If
            end if
          Else
            currentShadowCount(s) = 0
            BallShadowA(s).Opacity = 100*AmbientBSFactor
          End If
        Next
      Else                  'Hide dynamic shadows everywhere else
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub

'//////////////////////////////////////////////////////////////////////
'//
'//////////////////////////////////////////////////////////////////////


'DMD "                    ", "                    ", "", eNone, eNone, eNone, 1000, True, ""

' COPY EVERYTHING BELOW TO THE TOP OF YOUR TABLE SCRIPT UNDER OPTION EXPLICIT                             Start Pup Pack

'****** PuP Variables ******

Const PupScreen3 = 0

Dim usePUP: Dim cPuPPack: Dim PuPlayer: Dim PUPStatus: PUPStatus=false ' dont edit this line!!!

'*************************** PuP Settings for this table ********************************

usePUP   = true               ' enable Pinup Player functions for this table
cPuPPack = "jpsdeadpool"    ' name of the PuP-Pack / PuPVideos folder for this table

'//////////////////// PINUP PLAYER: STARTUP & CONTROL SECTION //////////////////////////

' This is used for the startup and control of Pinup Player

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

' ******* How to use PUPEvent to trigger / control a PuP-Pack *******

' Usage: pupevent(EventNum)

' EventNum = PuP Exxx trigger from the PuP-Pack

' Example: pupevent 102

' This will trigger E102 from the table's PuP-Pack

' DO NOT use any Exxx triggers already used for DOF (if used) to avoid any possible confusion

'************ PuP-Pack Startup **************

PuPStart(cPuPPack) 'Check for PuP - If found, then start Pinup Player / PuP-Pack
'**********************************************APRON LEFT***************************
Sub pupflasher6_Timer()
    pupflasher6.VideoCapUpdate="PUPSCREEN18"
end Sub

If PupScreen6 = 1 Then
  pupflasher6.VideoCapWidth=200
  pupflasher6.VideoCapHeight=200
  pupflasher6.visible=true
    pupflasher6.TimerEnabled=true
    pupflasher6.TimerInterval=60
    If FPS = 1 Then pupflasher6.timerinterval=40
End If

If PupScreen6 = 2 Then
  pupflasher6.VideoCapWidth=500
  pupflasher6.VideoCapHeight=500
  pupflasher6.visible=true
    pupflasher6.TimerEnabled=true
    pupflasher6.TimerInterval=60
    If FPS = 1 Then pupflasher6.timerinterval=40
End If

if PuPScreen6 = 0 Then
    pupflasher6.visible=false
    pupflasher6.TimerEnabled=false
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 18, ""FN"":12 }"
End If
'*********************************************APRON RIGHT********************************
Sub pupflasher7_Timer()
    pupflasher7.VideoCapUpdate="PUPSCREEN19"
end Sub

If PupScreen7 = 1 Then
  pupflasher7.VideoCapWidth=200
  pupflasher7.VideoCapHeight=200
  pupflasher7.visible=true
    pupflasher7.TimerEnabled=true
    pupflasher7.TimerInterval=60
    If FPS = 1 Then pupflasher7.timerinterval=40
End If

If PupScreen7 = 2 Then
  pupflasher7.VideoCapWidth=500
  pupflasher7.VideoCapHeight=500
  pupflasher7.visible=true
    pupflasher7.TimerEnabled=true
    pupflasher7.TimerInterval=60
    If FPS = 1 Then pupflasher7.timerinterval=40
End If

if PuPScreen7 = 0 Then
    pupflasher7.visible=false
    pupflasher7.TimerEnabled=false
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 19, ""FN"":12 }"
End If
'***********************************************BackWall********************************
Sub pupflasher3_Timer()
    pupflasher3.VideoCapUpdate="PUPSCREEN17"
end Sub

If PupScreen3 = 1 Then
  pupflasher3.VideoCapWidth=400
  pupflasher3.VideoCapHeight=400
  pupflasher3.visible=true
    pupflasher3.TimerEnabled=true
    pupflasher3.TimerInterval=60
    If FPS = 1 Then pupflasher3.timerinterval=40
End If

If PupScreen3 = 2 Then
  pupflasher3.VideoCapWidth=600
  pupflasher3.VideoCapHeight=600
  pupflasher3.visible=true
    pupflasher3.TimerEnabled=true
    pupflasher3.TimerInterval=60
    If FPS = 1 Then pupflasher3.timerinterval=40
End If

if PuPScreen3 = 0 Then
    pupflasher3.visible=false
    pupflasher3.TimerEnabled=false
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 17, ""FN"":12 }"
End If

'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************
' To add these slingshot corrections:
'   - On the table, add the endpoint primitives that define the two ends of the Slingshot
' - Initialize the SlingshotCorrection objects in InitSlingCorrection
'   - Call the .VelocityCorrect methods from the respective _Slingshot event sub


dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection

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
  AddSlingsPt 0, 0.00,  -4
  AddSlingsPt 1, 0.45,  -7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4

End Sub


Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LS, RS)
  dim x : for each x in a
    x.addpoint idx, aX, aY
  Next
End Sub

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function

Class SlingshotCorrection
  Public DebugOn, Enabled
  private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): Enabled = True : End Sub

  Public Property let Object(aInput) : Set Slingshot = aInput : End Property
  Public Property Let EndPoint1(aInput) : SlingX1 = aInput.x: SlingY1 = aInput.y: End Property
  Public Property Let EndPoint2(aInput) : SlingX2 = aInput.x: SlingY2 = aInput.y: End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 then Report
  End Sub

  Public Sub Report()         'debug, reports all coords in tbPL.text
    If not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1 : YL = SlingY1 : XR = SlingX2 : YR = SlingY2
    Else
      XL = SlingX2 : YL = SlingY2 : XR = SlingX1 : YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If ABS(XR-XL) > ABS(YR-YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If not IsEmpty(ModIn(0) ) then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      'debug.print " BallPos=" & BallPos &" Angle=" & Angle
      'debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled then aBall.Velx = RotVxVy(0)
      If Enabled then aBall.Vely = RotVxVy(1)
      'debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      'debug.print " "
    End If
  End Sub

End Class

'******************************************************
'*****   FLUPPER DOMES
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
' Example: RotateFlasher 1, 180     'where 1 is the flasher number and 180 is the angle of Z rotation

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
'   Objlevel(1) = 1 : FlasherFlash1_Timer
' End If
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
' Objlevel(1) = level/255 : FlasherFlash1_Timer
' End Sub


 Sub Flash1(Enabled)
  If Enabled Then
    Objlevel(1) = 1 : FlasherFlash1_Timer
  End If
  Sound_Flash_Relay enabled, Flasherbase1
 End Sub

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.2   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.4   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.2   ' *** lower this, if the blooms are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.6    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"

InitFlasher 1, "red"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 1,17 : RotateFlasher 2,0 : RotateFlasher 3,90 : RotateFlasher 4,90


Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 60
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 then
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
  objlight(nr).bulbhaloheight = objbase(nr).z -10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  dim xthird, ythird
  xthird = tablewidth/3
  ythird = tableheight/3

  If objbase(nr).x >= xthird and objbase(nr).x <= xthird*2 then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  end if

  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objbloom(nr).color = RGB(4,120,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4) : objbloom(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4) : objbloom(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255) : objbloom(nr).color = RGB(230,49,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50) : objbloom(nr).color = RGB(200,173,25)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59) : objbloom(nr).color = RGB(255,240,150)
    Case "orange" :  objlight(nr).color = RGB(255,70,0) : objflasher(nr).color = RGB(255,70,0) : objbloom(nr).color = RGB(255,70,0)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub


'******************************************************
'******  END FLUPPER DOMES
'******************************************************


'/////////////////////////////////////////////////////////////////
'         Siderails
'/////////////////////////////////////////////////////////////////

If Siderails = 1 Then
rrail.visible = 1
lrail.visible = 1
Else
rrail.visible = 0
lrail.visible = 0
End if

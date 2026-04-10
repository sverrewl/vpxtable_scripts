Option Explicit       ' Force explicit variable declaration
Randomize

'Const BallSize = 49
'Const BallMass = 1.1

Const BallSize = 50
Const BallMass = 1


' Load the FPVPX.vbs for Future Tables Comvert to VPX
'On Error Resume Next
'ExecuteGlobal GetTextFile("FPVPX.vbs")
'If Err Then MsgBox "you need the fpvpx.vbs for the proper functioning of the table"
'If FPVPXv> "" Then If FPVPXv < 1.01 or Err Then msgbox "FPVPX.vbs needs updated version for the proper functioning of the table"
On Error Goto 0

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Load the core.vbs for supporting Subs and functions
LoadCoreFiles

  Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Goto 0
  End Sub


if Table1.ShowDT Then
   hdisplay1.visible =1
   display1.visible = 1
   hdisplay2.visible = 1
  display2.visible = 1
   DispDmd1.visible = 0
 Else
   hdisplay1.visible = 0
   display1.visible = 0
   hdisplay2.visible = 0
  display2.visible = 0
   DispDmd1.visible = 0
End If

Const cGameName = "goonies"

' Define any Constants
Const TableName = "The Goonies"
Const myVersion = "1.03"
Const MaxPlayers = 1     ' from 1 to 4
Const BallSaverTime = 10 ' in seconds
Const MaxMultiplier = 10  ' limit to 10x in this game, both bonus multiplier and playfield multiplier
Const BallsPerGame = 3   ' usually 3 or 5
Const MaxMultiballs = 5  ' max number of balls during multiballs



'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.8



'**************************
'   PinUp Player USER Config
'**************************

dim PuPDMDDriverType: PuPDMDDriverType=1   ' 0=LCD DMD, 1=RealDMD 2=FullDMD (large LCD)
dim useRealDMDScale : useRealDMDScale=1    ' 0 or 1 for RealDMD scaling.  Choose which one you prefer. 1 is usually best for 128x32
dim useDMDVideos    : useDMDVideos=true    ' true or false to use DMD splash videos.
Dim pGameName       : pGameName="TheGoonies"  'pupvideos foldername, probably set to cGameName in realworld





Dim Credits
Dim BonusHeldPoints(4)
Dim PlayfieldMultiplier(4)
Dim bBonusHeld
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim specialhighscorename(4)
Dim specialhighscore(4)
Dim specialscore(4)
Dim replayscore
Dim replayscored
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

' HighScore
dim ScoreChecker
dim CheckAllScores
Dim HighScoreReward
Dim SpecialHighScoreReward
dim sortscores(4)
dim sortplayers(4)
Dim TextStr, TextStr2
Dim debugscore
Dim InProgress

' Define Game Control Variables
Dim BallsInHole

' Define Game Flags
Dim bGameInPlay
Dim bBallSaverReady
Dim bMusicOn
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bJackpot

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger during multiballs

dim FastFlips
Dim BallsInDrain

' Define Game Flags
Dim bFreePlay         ' Either in Free Play or Handling Credits
Dim bOnTheFirstBall     ' First Ball (player one). Used for Adding New Players
Dim bBallInPlungerLane    ' is there a ball in the plunger lane
Dim bBallSaverActive      ' is the ball saver active
Dim bMultiBallMode      ' multiball mode active ?
Dim bEnteringAHighScore   ' player is entering their name into the high score table
Dim turnonultradmd

Dim nvR1, nvR2, nvR3, nvR4, nvR5
Dim highscoreOn
Dim LevelBall
Dim mLevelMagnet

' *********************************************************************
' **                                                                 **
' **               Future Pinball Defined Script Events              **
' **                                                                 **
' *********************************************************************


' The Method Is Called Immediately the Game Engine is Ready to
' Start Processing the Script.
'
Sub Table1_Init()
    LoadEM

  If HasPup Then
    PuPInit  'Initialize PUPSystem AFTER you init B2s if applicable
  End If
    Dim i
    Randomize

    'Animation Ball Level
    Set mLevelMagnet = New cvpmMagnet
    With mLevelMagnet
        .InitMagnet LevelMagnet, 5
    .GrabCenter = 0
        .CreateEvents "mLevelMagnet"
    End With
    mLevelMagnet.MagnetOn = 1

    Set LevelBall = BallLevel.Createsizedball(8):LevelBall.Image = "BolaAgua":BallLevel.Kick 0, 0

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFX("nri-cannon", DOFContactors), SoundFX("KickerKick", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    'Boot Pinball
    BootE = True
    Boot.enabled = 1

    'Captive Balls
  kicker3.CreateSizedballWithMass Ballsize/2, BallMass
  kicker3.kick 0,0
  kicker4.CreateSizedballWithMass Ballsize/2, BallMass
  kicker4.kick 0,0
  kicker5.CreateSizedballWithMass Ballsize/2, BallMass
  kicker5.kick 0,0

  ' kill the last switch hit (this this variable is very usefull for game control)
  LastSwitchHit = ""'DummyTrigger

    'load saved values, highscore, names, jackpot
    Loadhs

    'Reset HighScore
    'Reseths

    HighScoreReward = 0
    SpecialHighScoreReward = 0


    ' freeplay or coins
    bFreePlay = True 'we dont want coins

    ' initialse any other flags
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bGameInPlay = False
    bAutoPlunger = False
    bMusicOn = True
    'BallsOnPlayfield = 0
    BallsInLock = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bInstantInfo = False
  BallsInDrain = 5

    'Replay
  replayscore = 7500000

    If Credits > 0 or bFreePlay Then DOF 136, DOFOn

End Sub

Dim LevelB
Sub LevelAnim()
    LevelB = True
 If LevelB = True Then
    mLevelMagnet.MagnetOn = 0
    vpmtimer.addtimer 500, "mLevelMagnet.MagnetOn = 1 '"
 end If
End Sub

Sub BallLevelKick_Hit
    BallLevelKick.kick 0, 3
End Sub


Dim bootT
Dim BootE

Sub Boot_Timer
 bootT = bootT + 1
 Select Case bootT
        Case 0:
   Playsound "fx_sys11_bootup"
        Case 1:
   Playsound "fx_sys11_bootup"
        Case 2:
   Playsound "fx_sys11_bootup"
        Case 3:
   bootT = 0
   BootE = False
   boot.enabled = 0
   DoorUp 0, 60
   FuturePinball_BeginPlay
End Select
End Sub


dim ballindrain

Sub swPlunger_Hit
  bulb1.State = 2
End Sub


'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                    'Called when table is nudged

    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity) AND(Tilt <15) Then 'show a warning
        'DMDFlush
        hdisplay1.text = "CAREFUL!"
        pupDMDDisplay "default", "CAREFUL!!", "", 3, 1, 10
        Playsound "jycsfx26"
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
        Tilted = True
        Playsound "tilt"
        PlaySong "mu_end"
        'display Tilt
        DMDUpdate.Enabled = 0
        hdisplay1.text = "TILT!"
        pupDMDDisplay "default", "TILT!!", "@tilt.mp4", 4, 1, 10
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
    DOF 101, DOFOff
    DOF 102, DOFOff
        Bumper1.Force = 0
        Bumper2.Force = 0
        Bumper3.Force = 0
        Bumper4.Force = 0
        LeftSlingshotRubber.Disabled = 1
        RightSlingshotRubber.Disabled = 1
        LeftSlingshotRubber1.Disabled = 1
        RightSlingShotRubber1.Disabled = 1
        skilltimer1.Enabled = false
        skilltimer2.Enabled = false
        skilltimer3.Enabled = false
        skilltimer4.Enabled = false
        skilltimer5.Enabled = false
    Else
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        Bumper1.Force = 8
        Bumper2.Force = 8
        Bumper3.Force = 8
        Bumper4.Force = 8
        LeftSlingshotRubber.Disabled = 0
        RightSlingshotRubber.Disabled = 0
        LeftSlingshotRubber1.Disabled = 0
        RightSlingShotRubber1.Disabled = 0
    End If
End Sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' Sub TiltRecoveryTimer_Timer()
'     ' if all the balls have been drained then..
'     If(BallsOnPlayfield = 0) Then
'         ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
'         EndOfBall()
'         TiltRecoveryTimer.Enabled = False
'     End If
' ' else retry (checks again in another second or so)
' End Sub







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
        If UBound(tmp) = 3 Then 'we have 4 captive balls on the table (-1 means no balls, 0 is the first ball, 1 is the second..)
            GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    DOF 126, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
    Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
    Luz_Mesh.visible = 1
End Sub

Sub GiOff
    DOF 126, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSatDark"
    Luz_Mesh.visible = 0
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


'****************************************
' Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub GameTimer_Timer
    'RollingTimer
    'RollingUpdate
    BallShadowUpdate

  'leftdiverterP.objRotZ = leftdiverter.CurrentAngle
   ' rightdiverterP.objRotZ = rightdiverter.CurrentAngle


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
  Dim red, orange, amber, yellow, darkgreen, green, blue, darkblue, purple, white, base

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
  base = 11

  Sub SetLightColor(n, col, stat)
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
      Case base
        n.color = RGB(255, 197, 143)
        n.colorfull = RGB(255, 255, 236)
    End Select
    If stat <> -1 Then
      n.State = 0
      n.State = stat
    End If
  End Sub

  Sub ResetAllLightsColor ' Called at a new game
    SetLightColor toplane1l, green, -1
    SetLightColor toplane2l, green, -1
    SetLightColor toplane3l, green, -1

    SetLightColor topbank1l, red, -1
    SetLightColor topbank2l, red, -1
    SetLightColor topbank3l, red, -1

    SetLightColor locollect, amber, -1
    SetLightColor lojackpot, amber, -1

    SetLightColor rocollect, amber, -1
    SetLightColor rojackpot, amber, -1

    SetLightColor tavernstashl, red, -1
    SetLightColor tavern1l, red, -1

    SetLightColor rain1, red, -1
    SetLightColor rain2, orange, -1
    SetLightColor rain3, yellow, -1
    SetLightColor rain4, green, -1
    SetLightColor rain5, blue, -1
    SetLightColor rain6, purple, -1

    SetLightColor findkey, amber, -1
    SetLightColor keyl, amber, -1

    SetLightColor pointsbank, green, -1
    SetLightColor pointsbanka, green, -1
    SetLightColor points10, darkgreen, -1
    SetLightColor points20, darkgreen, -1
    SetLightColor points30, darkgreen, -1
    SetLightColor points40, darkgreen, -1
    SetLightColor points50, darkgreen, -1
    SetLightColor points60, darkgreen, -1
    SetLightColor points70, darkgreen, -1
    SetLightColor points80, darkgreen, -1
    SetLightColor points90, darkgreen, -1

    SetLightColor lep51l, white, -1
    SetLightColor lep52l, amber, -1

    SetLightColor lep41l, white, -1
    SetLightColor lep42l, amber, -1

    SetLightColor lep31l, white, -1
    SetLightColor lep32l, amber, -1

    SetLightColor lep21l, white, -1
    SetLightColor lep22l, amber, -1

    SetLightColor lep11l, white, -1
    SetLightColor lep12l, amber, -1

    SetLightColor twoxl, red, -1
    SetLightColor threexl, red, -1

    SetLightColor score1l, yellow, -1
    SetLightColor score2l, yellow, -1
    SetLightColor score3l, yellow, -1
    SetLightColor score4l, yellow, -1
    SetLightColor score5l, yellow, -1

    SetLightColor ebl, green, -1
    SetLightColor ebr, green, -1

    SetLightColor lane1l, orange, -1
    SetLightColor lane2l, orange, -1
    SetLightColor lane3l, orange, -1
    SetLightColor lane4l, orange, -1

    SetLightColor kickbackll, darkgreen, -1
    SetLightColor kickbackrl, darkgreen, -1

    SetLightColor LightShootAgain, red, -1

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

  Dim RGBStep2, RGBFactor2, rRed2, rGreen2, rBlue2, RainbowLights2
  Sub StartRainbow2(n)
    set RainbowLights2 = n
    RGBStep2 = 0
    RGBFactor2 = 5
    rRed2 = 255
    rGreen2 = 0
    rBlue2 = 0
    RainbowTimer1.Enabled = 1
  End Sub


  Dim RGBStep3, RGBFactor3, rRed3, rGreen3, rBlue3, RainbowLights3
  Sub StartRainbow3(n)
    set RainbowLights3 = n
    RGBStep3 = 0
    RGBFactor3 = 5
    rRed3 = 255
    rGreen3 = 0
    rBlue3 = 0
    RainbowTimer2.Enabled = 1
  End Sub

  Sub StopRainbow(n)
    Dim obj
    RainbowTimer.Enabled = 0
    RainbowTimer.Enabled = 0
      For each obj in RainbowLights
        SetLightColor obj, "white", 0
      Next
  End Sub

  Sub StopRainbow2(n)
    Dim obj
    RainbowTimer1.Enabled = 0
      For each obj in RainbowLights2
        SetLightColor obj, "white", 0
        obj.state = 1
        obj.Intensity = 12
      Next
  End Sub

  Sub RainbowTimer_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep
      Case 0 'Green
        rGreen = rGreen + RGBFactor
        If rGreen > 255 then
          rGreen = 255
          RGBStep = 1
        End If
      Case 1 'Red
        rRed = rRed - RGBFactor
        If rRed < 0 then
          rRed = 0
          RGBStep = 2
        End If
      Case 2 'Blue
        rBlue = rBlue + RGBFactor
        If rBlue > 255 then
          rBlue = 255
          RGBStep = 3
        End If
      Case 3 'Green
        rGreen = rGreen - RGBFactor
        If rGreen < 0 then
          rGreen = 0
          RGBStep = 4
        End If
      Case 4 'Red
        rRed = rRed + RGBFactor
        If rRed > 255 then
          rRed = 255
          RGBStep = 5
        End If
      Case 5 'Blue
        rBlue = rBlue - RGBFactor
        If rBlue < 0 then
          rBlue = 0
          RGBStep = 0
        End If
    End Select
      For each obj in RainbowLights
        obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
        obj.colorfull = RGB(rRed, rGreen, rBlue)
      Next
  End Sub

  Sub RainbowTimer1_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep2
      Case 0 'Green
        rGreen2 = rGreen2 + RGBFactor2
        If rGreen2 > 255 then
          rGreen2 = 255
          RGBStep2 = 1
        End If
      Case 1 'Red
        rRed2 = rRed2 - RGBFactor2
        If rRed2 < 0 then
          rRed2 = 0
          RGBStep2 = 2
        End If
      Case 2 'Blue
        rBlue2 = rBlue2 + RGBFactor2
        If rBlue2 > 255 then
          rBlue2 = 255
          RGBStep2 = 3
        End If
      Case 3 'Green
        rGreen2 = rGreen2 - RGBFactor2
        If rGreen2 < 0 then
          rGreen2 = 0
          RGBStep2 = 4
        End If
      Case 4 'Red
        rRed2 = rRed2 + RGBFactor2
        If rRed2 > 255 then
          rRed2 = 255
          RGBStep2 = 5
        End If
      Case 5 'Blue
        rBlue2 = rBlue2 - RGBFactor2
        If rBlue2 < 0 then
          rBlue2 = 0
          RGBStep2 = 0
        End If
    End Select
      For each obj in RainbowLights2
        obj.color = RGB(rRed2 \ 10, rGreen2 \ 10, rBlue2 \ 10)
        obj.colorfull = RGB(rRed2, rGreen2, rBlue2)
      Next
  End Sub


  Sub RainbowTimer2_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep3
      Case 0 'Green
        rGreen3 = rGreen3 + RGBFactor3
        If rGreen3 > 255 then
          rGreen3 = 255
          RGBStep3 = 1
        End If
      Case 1 'Red
        rRed3 = rRed3 - RGBFactor3
        If rRed3 < 0 then
          rRed3 = 0
          RGBStep3 = 2
        End If
      Case 2 'Blue
        rBlue3 = rBlue3 + RGBFactor3
        If rBlue3 > 255 then
          rBlue3 = 255
          RGBStep3 = 3
        End If
      Case 3 'Green
        rGreen3 = rGreen3 - RGBFactor3
        If rGreen3 < 0 then
          rGreen3 = 0
          RGBStep3 = 4
        End If
      Case 4 'Red
        rRed3 = rRed3 + RGBFactor3
        If rRed3 > 255 then
          rRed3 = 255
          RGBStep3 = 5
        End If
      Case 5 'Blue
        rBlue3 = rBlue3 - RGBFactor3
        If rBlue3 < 0 then
          rBlue3 = 0
          RGBStep3 = 0
        End If
    End Select
      For each obj in RainbowLights3
        obj.color = RGB(rRed3 \ 10, rGreen3 \ 10, rBlue3 \ 10)
        obj.colorfull = RGB(rRed3, rGreen3, rBlue3)
      Next
  End Sub








' *********************************************************************
'               Funciones para los sonidos de la mesa
' *********************************************************************



'Function Pan(ball) ' Calculala posición estéreo de la bola (izquierda a derecha)
'    Dim tmp
'    tmp = ball.x * 2 / table1.width-1
'    If tmp > 0 Then
'        Pan = Csng(tmp ^10)
'    Else
'        Pan = Csng(-((- tmp) ^10))
'    End If
'End Function

'Function Pitch(ball) ' Calcula el tono según la velocidad de la bola
'    Pitch = BallVel(ball) * 20
'End Function

'Function BallVel(ball) 'Calcula la velocidad de la bola
'    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
'End Function

'Function AudioFade(ball) 'Solo para VPX 10.4 y siguentes: calcula la posición de arriba/abajo de la bola, para mesas con el sonido dolby
'    Dim tmp
'    tmp = ball.y * 2 / Table1.height-1
'    If tmp > 0 Then
'        AudioFade = Csng(tmp ^10)
'    Else
'        AudioFade = Csng(-((- tmp) ^10))
'    End If
'End Function

'Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
'    Dim tmp
'    tmp = tableobj.x * 2 / table1.width-1
'    If tmp > 0 Then
'        AudioPan = Csng(tmp ^10)
'    Else
'        AudioPan = Csng(-((- tmp) ^10) )
'    End If
'End Function

'Sub PlaySoundAt(soundname, tableobj) ' Hace sonar un sonido en la posición de un objeto, como bumpers y flippers
'    PlaySound soundname, 0, 1, Pan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
'End Sub
'
'Sub PlaySoundAtBall(soundname) ' hace sonar un sonido en la posición de la bola
'    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
' PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
'End Sub



'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5, BallShadow6, BallShadow7, BallShadow8, BallShadow9, BallShadow10, BallShadow11, BallShadow12)

Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' render the shadow for each ball
    For b = 0 to Ubound(BOT)
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) - 10
    End If
      ballShadow(b).Y = BOT(b).Y + 15
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub



'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

End Sub



'******************************
' Diverse Collection Hit Sounds
'******************************

'Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aPlastics_Hit(idx):PlaySound "fx_plastichit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aDropTargets_Hit(idx):PlaySound "fx_target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

Sub TriggerRamp_Hit :PuPEvent 806
    StopSound "wirerolling"
    PlaySoundAtLevelActiveBall ("wireramp_stop"), 1 * VolumeDial
End Sub

Sub RHelp2_Hit() :PuPEvent 807
    StopSound "wirerolling"
    PlaySoundAtLevelActiveBall ("wireramp_stop"), 1 * VolumeDial
End Sub

Sub LHelp1_Hit() :PuPEvent 808
    StopSound "wirerolling"
    PlaySoundAtLevelActiveBall ("wireramp_stop"), 1 * VolumeDial
End Sub



'Sub Trigger2_hit()
'    PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall)
'End Sub

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'

Sub StopGameOverSong
    PlaySong "mu_end"
    StopSound Song:Song = ""
    StopAllSounds
End Sub






'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
  If BootE = True Then Exit Sub
    If Keycode = AddCreditKey Then
        Credits = Credits + 1
    DOF 136, DOFOn
        If(Tilted = False) Then

      Select Case Int(rnd*3)
        Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
        Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
        Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
      End Select
            'startB2S (31)
           ' DisplayB2SText2 " CREDITS " &credits
            hdisplay1.Text = " CREDITS " &credits
          'PuPlayer.playlistplayex pDMD,"DMDBackground","backdmd.jpg",0,1
            'pDMDSplashBonus "CREDITS " & credits ,"PRESS START BUTTON",10,33023
            pupDMDDisplay "credits", "CREDITS "&credits&"^PRESS START", "", 3, 1, 20
          'PuPlayer.playlistplayex pDMD,"DMDBackground","GooniesLogoDMD.png",0,1
            PlaySound "go"
       '     If NOT bGameInPlay Then ShowTableInfo:
        End If
    End If

    If keycode = RightFlipperKey Then pAttractNext 'if in attract go to next attract

    If keycode = PlungerKey Then
      PlungerIM.AutoFire
    If bBallInPlungerLane Then DOF 111, DOFPulse: DOF 121, DOFPulse
    End If

    If bGameInPlay AND NOT Tilted Then
        If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft():CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight():CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter():CheckTilt
        If keycode = MechanicalTilt Then CheckTilt

        ' nFozzy - Start
    If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
    If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1
    ' nFozzy - End

        If keycode = StartGameKey And bAttractMode = True Then
      soundStartButton()
            If((PlayersPlayingGame < MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                Else

                    If(Credits > 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
            If Credits < 1 and Not bFreePlay Then DOF 126, DOFOff
                    Else
                        ' Not Enough Credits to start a game.

                        'DMD CenterLine(0, "CREDITS " & Credits), CenterLine(1, "INSERT COIN"), 0, eNone, eBlink, eNone, 500, True, ""
                    End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey And bAttractMode = True Then
        soundStartButton()
                If(bFreePlay = True) Then
                    If(BallsOnPlayfield = 0) Then
                        ResetForNewGame()
                    End If
                Else
                    If(Credits > 0) Then
                        If(BallsOnPlayfield = 0) Then
                            Credits = Credits - 1
              If Credits < 1 And Not bFreePlay Then DOF 126, DOFOff
                            ResetForNewGame()
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        hdisplay1.Text = " CREDITS " &credits&"   INSERT COIN"

                      'PuPlayer.playlistplayex pDMD,"DMDBackground","backdmd.jpg",0,1
               pupDMDDisplay "default", "CREDITS "&credits&"^INSERT COIN", "", 3, 1, 10
                      'PuPlayer.playlistplayex pDMD,"DMDBackground","GooniesLogoDMD.png",0,1

                    End If
                End If
            End If
    End If ' If (GameInPlay)


    If hsbModeActive Then EnterHighScoreKey(keycode)

    If SpecialhsbModeActive Then EnterSpecialHighScoreKey(keycode)

' Table specific
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If bGameInPLay AND NOT Tilted Then
        ' nFozzy - Start
    If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
    If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
        If keycode = LeftFlipperKey Then SolLFlipper 0
        If keycode = RightFlipperKey Then SolRFlipper 0
        ' nFozzy - End
    End If
End Sub




'********************
'     Flippers
'********************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend
    leftlight
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
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
    rightlight
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  RightFlipperCollide parm
End Sub


' flippers hit Sound

'Sub LeftFlipper_Collide(parm)
' CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
'    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.05, 0.25
'End Sub
'
'Sub RightFlipper_Collide(parm)
' CheckLiveCatch Activeball, RightFlipper, RFCount, parm
'    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.05, 0.25
'End Sub






Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    bInstantInfo = True
'    DMDFlush
'    UltraDMDTimer.Enabled = 1
End Sub

Sub InstantInfo
     hdisplay1.Text = " POW " &POWBonusCount
 '   Jackpot = 1000000 + Round(Score(CurrentPlayer) / 10, 0)
 '   DMD "black.jpg", "", "INSTANT INFO", 500
 '   DMD "black.jpg", "JACKPOT", Jackpot, 800
 '   DMD "black.jpg", "LEVEL", Level(CurrentPlayer), 800
 '   DMD "black.jpg", "BONUS MULT", BonusMultiplier(CurrentPlayer), 800
 '   DMD "black.jpg", "ORBIT BONUS", OrbitHits, 800
  '  DMD "black.jpg", "LANE BONUS", LaneBonus, 800
   ' DMD "black.jpg", "TARGET BONUS", TargetBonus, 800
  '  DMD "black.jpg", "RAMP BONUS", RampBonus, 800
 '   DMD "black.jpg", "MONSTERS KILLED", Monsters(CurrentPlayer), 800
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
    If B2SOn Then Controller.Stop
End Sub

Sub AwardSpecial()
    'DMD " ", "EXTRA GAME", 2000
    hdisplay1.text = "EXTRA GAME"& "    CREDITS:  "& Credits
    'PlaySound SoundFXDOF("fx_knocker",124,DOFPulse,DOFKnocker)
  PlaySound SoundFXDOF("Knocker_1",135,DOFPulse,DOFKnocker)
  DOF 121, DOFPulse

    Credits = Credits + 1
    DisplayB2SText2 "   EXTRA GAME   " & "    CREDITS:  "& Credits
    LightEffect 2
    FlashEffect 2
End Sub

'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 9000000 End If

    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If

    x = LoadValue(TableName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 8500000 End If

    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If

    x = LoadValue(TableName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 8000000 End If

    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If

    x = LoadValue(TableName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 7500000 End If

    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If



    x = LoadValue(TableName, "SpecialHighScore1")
    If(x <> "") Then SpecialHighScore(0) = CDbl(x) Else SpecialHighScore(0) = 4 End If

    x = LoadValue(TableName, "SpecialHighScore1Name")
    If(x <> "") Then SpecialHighScoreName(0) = x Else SpecialHighScoreName(0) = "AAA" End If

    x = LoadValue(TableName, "SpecialHighScore2")
    If(x <> "") then SpecialHighScore(1) = CDbl(x) Else SpecialHighScore(1) = 3 End If

    x = LoadValue(TableName, "SpecialHighScore2Name")
    If(x <> "") then SpecialHighScoreName(1) = x Else SpecialHighScoreName(1) = "BBB" End If

    x = LoadValue(TableName, "SpecialHighScore3")
    If(x <> "") then SpecialHighScore(2) = CDbl(x) Else SpecialHighScore(2) = 2 End If

    x = LoadValue(TableName, "SpecialHighScore3Name")
    If(x <> "") then SpecialHighScoreName(2) = x Else SpecialHighScoreName(2) = "CCC" End If

    x = LoadValue(TableName, "SpecialHighScore4")
    If(x <> "") then SpecialHighScore(3) = CDbl(x) Else SpecialHighScore(3) = 1 End If

    x = LoadValue(TableName, "SpecialHighScore4Name")
    If(x <> "") then SpecialHighScoreName(3) = x Else SpecialHighScoreName(3) = "DDD" End If



    x = LoadValue(TableName, "Credits")
    If(x <> "") then Credits = CInt(x) Else Credits = 0 End If

    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If

  '  x = LoadValue(TableName, "Score")
  '  If(x <> "") then Score(CurrentPlayer) = CInt(x) Else Score(CurrentPlayer) = 0 End If
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


    SaveValue TableName, "SpecialHighScore1", SpecialHighScore(0)
    SaveValue TableName, "SpecialHighScore1Name", SpecialHighScoreName(0)
    SaveValue TableName, "SpecialHighScore2", SpecialHighScore(1)
    SaveValue TableName, "SpecialHighScore2Name", SpecialHighScoreName(1)
    SaveValue TableName, "SpecialHighScore3", SpecialHighScore(2)
    SaveValue TableName, "SpecialHighScore3Name", SpecialHighScoreName(2)
    SaveValue TableName, "SpecialHighScore4", SpecialHighScore(3)
    SaveValue TableName, "SpecialHighScore4Name", SpecialHighScoreName(3)

    SaveValue TableName, "Credits", Credits
    SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
   ' SaveValue TableName, "Score(CurrentPlayer)", Score(CurrentPlayer)

End Sub

Sub Reseths
    HighScoreName(0) = "AAA"
    HighScoreName(1) = "BBB"
    HighScoreName(2) = "CCC"
    HighScoreName(3) = "DDD"
    HighScore(0) = 90000
    HighScore(1) = 85000
    HighScore(2) = 80000
    HighScore(3) = 75000

    SpecialHighScoreName(0) = "AAA"
    SpecialHighScoreName(1) = "BBB"
    SpecialHighScoreName(2) = "CCC"
    SpecialHighScoreName(3) = "DDD"
    SpecialHighScore(0) = 4
    SpecialHighScore(1) = 3
    SpecialHighScore(2) = 2
    SpecialHighScore(3) = 1
    Savehs
End Sub

' ***********************************************************
'  High Score Initals Entry Functions - based on Black's code
' ***********************************************************

Dim hsbModeActive, SpecialhsbModeActive
Dim hsEnteredName, hsSpecialEnteredName
Dim hsEnteredDigits(3)
Dim hsSpecialEnteredDigits(3)
Dim hsCurrentDigit, hsSpecialCurrentDigit
Dim hsValidLetters, hsSpecialValidLetters
Dim hsCurrentLetter, hsSpecialCurrentLetter
Dim hsLetterFlash, hsSpecialLetterFlash


Sub CheckHighscore()
    DMDUpdate.enabled = 0
    display2.Text = " "
    display1.Text = " "
    'PuPlayer.LabelShowPage pDMD,1,0,""
    'pDMDSetPage 0 'blank text overlay
    Dim tmp

    'Score(1)=HighScore(3) +1000

    tmp = Score(1)

    If Score(2)> tmp Then tmp = Score(2)
    If Score(3)> tmp Then tmp = Score(3)
    If Score(4)> tmp Then tmp = Score(4)

    If tmp> HighScore(1) Then 'add 1 credit for beating the highscore
       HighScoreReward = 1
    End If

    If tmp> HighScore(3) Then
        PlaySound "spchkittenme"
        HighScore(3) = tmp
        display1.text = "  GREAT SCORE   " & "       "
        pupDMDDisplay "default", "GREAT SCORE^"&Score(CurrentPlayer), "", 2, 0, 10
        'enter player's name
        GiOff
        vpmtimer.addtimer 2500, "HighScoreEntryInit '"
       ' HighScoreEntryInit()
    Else
        'EndOfBallComplete()
        CheckSpecialHighscore
    End If
End Sub


Sub HighScoreEntryInit()
    hsbModeActive = True
    PlaySound "mu_jychighscore"
    hsLetterFlash = 0

    hsEnteredDigits(0) = "A"
    hsEnteredDigits(1) = "-"
    hsEnteredDigits(2) = "-"
    hsCurrentDigit = 0

    hsValidLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ<+-0123456789" ' < is used to delete the last letter
    hsCurrentLetter = 1
   ' DMDFlush
    display1.text = "YOUR NAME:" & " "
     'pDMDSplashHighScore "YOUR NAME" , " ", 2,  33023
    'DMDId "hsc", "", "YOUR NAME:", " ", 999999
    HighScoreDisplayName()
End Sub


Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        Playsound "jycsfx12"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayName()
    End If

    If keycode = RightFlipperKey Then
        Playsound "jycsfx12"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayName()
    End If

    If keycode = StartGameKey OR keycode = PlungerKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<") then
            playsound "jycsfx9"
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
    DMDUpdate.enabled = 0
    Dim i, TempStr
    display1.text = TempStr
    'pDMDSplashHighScore "ENTER YOUR NAME" , " " & TempStr, 2,  33023
    'pupDMDDisplay "HSEnter", "ENTER NAME^"& TempStr, "", 20, 1, 10


    TempStr = " >"
    if(hsCurrentDigit> 0) then TempStr = TempStr & hsEnteredDigits(0)
    if(hsCurrentDigit> 1) then TempStr = TempStr & hsEnteredDigits(1)
    if(hsCurrentDigit> 2) then TempStr = TempStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3) then
        if(hsLetterFlash <> 0) then
            TempStr = TempStr & "_"
            'DisplayB2SText TempStr & "_"
            display2.TEXT = TempStr & "_"
            'pDMDSplashHighScore " ^"& TempStr&"_", 2,  33023
            pupDMDDisplay "HSEnter", " ^"& TempStr&"_", "", 20, 0, 10
        else
            TempStr = TempStr & mid(hsValidLetters, hsCurrentLetter, 1)
            'DisplayB2SText TempStr & mid(hsValidLetters, hsCurrentLetter, 1)
            display2.TEXT = TempStr & mid(hsValidLetters, hsCurrentLetter, 1)
            'pDMDSplashHighScore " " , " " & TempStr & mid(hsValidLetters, hsCurrentLetter, 1), 99,  33023
            pupDMDDisplay "HSEnter", " ^" & TempStr & mid(hsValidLetters, hsCurrentLetter, 1), "", 20, 0, 10
        end if
    end if

    if(hsCurrentDigit <1) then TempStr = TempStr & hsEnteredDigits(1)
    if(hsCurrentDigit <2) then TempStr = TempStr & hsEnteredDigits(2)

    TempStr = TempStr & "< "
   ' DMDMod "hsc", "YOUR NAME:", Mid(TempStr, 2, 5), 999999
     'pDMDSplashHighScore "ENTER HIGHSCORE NAME" , " " & Mid(TempStr, 2, 5), 99,  33023
     pupDMDDisplay "HSEnter", "ENTER HIGHSCORE NAME^"&" "&Mid(TempStr, 2, 5), "", 20, 0, 10
    display1.TEXT = "ENTER HIGHSCORE NAME "
    display2.TEXT = "    " & Mid(TempStr, 2, 5)
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
   ' DMDFlush
    DMDUpdate.enabled = 1
    'EndOfBallComplete()
    CheckSpecialHighscore()
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



'*****************
' Claw Special HighScore
'*****************


Sub CheckSpecialHighscore()

    DMDUpdate.enabled = 0
    display1.Text = " "
    display2.Text = " "
    'pDMDSplashHighScore " " , " ", 2,  33023
    Dim tmp
    tmp = SpecialScore(1)

    If SpecialScore(2)> tmp Then tmp = SpecialScore(2)
    If SpecialScore(3)> tmp Then tmp = SpecialScore(3)
    If SpecialScore(4)> tmp Then tmp = SpecialScore(4)

    If tmp> SpecialHighScore(1) Then 'add 1 credit for beating the highscore
       SpecialHighScoreReward = 1
    End If

    If tmp> SpecialHighscore(3) Then
        SpecialHighscore(3) = tmp
        'enter player's name
        PlaySound "spchnotoveryet"
        display1.text = "  GOOD GAME  "
        display2.text = " PUSSY CAT  "
        'pDMDSplashHighScore "SPECIAL HIGHSCORE " , " ", 2,  33023
        pupDMDDisplay "HSEnter", "SPECIAL", "HIGHSCORE", 2, 1, 10
        vpmtimer.addtimer 2000, "SpecialHighScoreEntryInit '"
        'SpecialHighScoreEntryInit()
    Else
        SolLFlipper 0
        SolRFlipper 0
        CreditsHighScoreReward.Enabled = 1
        'EndOfBallComplete()
    End If
End Sub


Sub SpecialHighScoreEntryInit()
    SpecialhsbModeActive = True
    StopSound "mu_jychighscore"
    PlaySound "mu_jychighscore"
    hsSpecialLetterFlash = 0


    hsSpecialEnteredDigits(0) = "A"
    hsSpecialEnteredDigits(1) = "-"
    hsSpecialEnteredDigits(2) = "-"
    hsSpecialCurrentDigit = 0

    hsSpecialValidLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ<+-0123456789" ' < is used to delete the last letter
    hsSpecialCurrentLetter = 1
   ' DMDFlush
    'display1.text = "YOUR NAME:" & " "
    'DMDId "hsc", "", "YOUR NAME:", " ", 999999
    SpecialHighScoreDisplayName()
End Sub


Sub EnterSpecialHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        Playsound "jycsfx12"
        hsSpecialCurrentLetter = hsSpecialCurrentLetter - 1
        if(hsSpecialCurrentLetter = 0) then
            hsSpecialCurrentLetter = len(hsSpecialValidLetters)
        end if
        SpecialHighScoreDisplayName()
    End If

    If keycode = RightFlipperKey Then
        Playsound "jycsfx12"
        hsSpecialCurrentLetter = hsSpecialCurrentLetter + 1
        if(hsSpecialCurrentLetter> len(hsSpecialValidLetters) ) then
            hsSpecialCurrentLetter = 1
        end if
        SpecialHighScoreDisplayName()
    End If

    If keycode = StartGameKey OR keycode = PlungerKey Then
        if(mid(hsSpecialValidLetters, hsSpecialCurrentLetter, 1) <> "<") then
            playsound "jycsfx9"
            hsSpecialEnteredDigits(hsSpecialCurrentDigit) = mid(hsSpecialValidLetters, hsSpecialCurrentLetter, 1)
            hsSpecialCurrentDigit = hsSpecialCurrentDigit + 1
            if(hsSpecialCurrentDigit = 3) then
                SpecialHighScoreCommitName()
            else
                SpecialHighScoreDisplayName()
            end if
        else
            playsound "fx_Esc"
            hsSpecialEnteredDigits(hsSpecialCurrentDigit) = " "
            if(hsSpecialCurrentDigit> 0) then
                hsSpecialCurrentDigit = hsSpecialCurrentDigit - 1
            end if
            SpecialHighScoreDisplayName()
        end if
    end if
End Sub

Sub SpecialHighScoreDisplayName()
    DMDUpdate.enabled = 0
    Dim i, TempStr
    display2.text = TempStr
    'pDMDSplashHighScore " " , " " & TempStr, 2,  33023
    'pupDMDDisplay "HSEnter", "^"& TempSTr, "", 20, 1, 10
    TempStr = " >"
    if(hsSpecialCurrentDigit> 0) then TempStr = TempStr & hsSpecialEnteredDigits(0)
    if(hsSpecialCurrentDigit> 1) then TempStr = TempStr & hsSpecialEnteredDigits(1)
    if(hsSpecialCurrentDigit> 2) then TempStr = TempStr & hsSpecialEnteredDigits(2)

    if(hsSpecialCurrentDigit <> 3) then
        if(hsSpecialLetterFlash <> 0) then
            TempStr = TempStr & "_"
            'DisplayB2SText TempStr & "_"
            display2.TEXT = TempStr & "_"
            'pDMDSplashHighScore " " , "_" & TempStr, 2,  33023
            pupDMDDisplay "HSEnter", "^"& TempSTr, "", 20, 0, 10
        else
            TempStr = TempStr & mid(hsSpecialValidLetters, hsSpecialCurrentLetter, 1)
            'DisplayB2SText TempStr & mid(hsSpecialValidLetters, hsSpecialCurrentLetter, 1)
            display2.TEXT = TempStr & mid(hsSpecialValidLetters, hsSpecialCurrentLetter, 1)
            'pDMDSplashHighScore " " , " " & TempStr & mid(hsSpecialValidLetters, hsSpecialCurrentLetter, 1), 2,  33023
            pupDMDDisplay "HSEnter", "^"&" "& TempStr & mid(hsSpecialValidLetters, hsSpecialCurrentLetter, 1), "", 20, 0, 10
        end if
    end if

    if(hsSpecialCurrentDigit <1) then TempStr = TempStr & hsSpecialEnteredDigits(1)
    if(hsSpecialCurrentDigit <2) then TempStr = TempStr & hsSpecialEnteredDigits(2)

    TempStr = TempStr & "< "
   ' DMDMod "hsc", "YOUR NAME:", Mid(TempStr, 2, 5), 999999
    'pDMDSplashHighScore "MASTER NAME" , " " & Mid(TempStr), 2,  33023
    pupDMDDisplay "HSEnter","MASTER NAME^ " & Mid(TempStr, 2, 5), "", 20, 0, 10
    'else UNKNOWN
    display1.TEXT = "CLAW MASTER NAME: "
    display2.TEXT = "    " & Mid(TempStr, 2, 5)
End Sub

Sub SpecialHighScoreCommitName()
    SpecialhsbModeActive = False
    'PlaySong "m_end"
    hsSpecialEnteredName = hsSpecialEnteredDigits(0) & hsSpecialEnteredDigits(1) & hsSpecialEnteredDigits(2)
    if(hsSpecialEnteredName = "   ") then
        hsSpecialEnteredName = "YOU"
    end if

    SpecialHighScoreName(3) = hsSpecialEnteredName
    SortSpecialHighscore
    DMDUpdate.enabled = 1
    SpecialReward = True
    CreditsHighScoreReward.Enabled = True
    'EndOfBallComplete()
End Sub

Sub SortSpecialHighscore
    Dim tmp, tmp2, i, j
    For i = 0 to 3
        For j = 0 to 2
            If SpecialHighScore(j) <SpecialHighScore(j + 1) Then
                tmp = SpecialHighScore(j + 1)
                tmp2 = SpecialHighScoreName(j + 1)
                SpecialHighScore(j + 1) = SpecialHighScore(j)
                SpecialHighScoreName(j + 1) = SpecialHighScoreName(j)
                SpecialHighScore(j) = tmp
                SpecialHighScoreName(j) = tmp2
            End If
        Next
    Next
End Sub



Dim SpecialReward
Sub CreditsHighScoreReward_Timer()
    CreditsHighScoreReward.Enabled = False
    SolLFlipper 0
    SolRFlipper 0
 If SpecialReward = True Then
    SpecialReward = False

  If HighScoreReward = 1 And SpecialHighScoreReward = 1  Then
     Credits = Credits + 1
    'PlaySound SoundFXDOF("fx_knocker",124,DOFPulse,DOFKnocker)
  PlaySound SoundFXDOF("Knocker_1",135,DOFPulse,DOFKnocker)
  DOF 121, DOFPulse
     display1.text = "     CREDITS" & "  "&(Credits) &"  " & "      "
     HighScoreReward = 0: SpecialHighScoreReward = 0

     pupDMDDisplay "default", "CREDITS^"&Credits&"|16744448", "", 3, 1, 10
     vpmtimer.addtimer 1000, "UnCreditoMas '"
  End If


   If HighScoreReward = 1 And SpecialHighScoreReward = 0 Then
      Credits = Credits + 1
  DOF 136, DOFOn
    'PlaySound SoundFXDOF("fx_knocker",124,DOFPulse,DOFKnocker)
  PlaySound SoundFXDOF("Knocker_1",135,DOFPulse,DOFKnocker)
  DOF 121, DOFPulse
      display1.text = "     CREDITS" & "  "&(Credits) &"  " & "      "
      HighScoreReward = 0
      pupDMDDisplay "default", "CREDITS^"&Credits&"|16744448", "", 3, 21, 10
     vpmtimer.addtimer 1000, "EndOfBallComplete() '"
   End If

   If HighScoreReward = 0 And SpecialHighScoreReward = 1 Then
    Credits = Credits + 1
  DOF 136, DOFOn
    'PlaySound SoundFXDOF("fx_knocker",124,DOFPulse,DOFKnocker)
  PlaySound SoundFXDOF("Knocker_1",135,DOFPulse,DOFKnocker)
  DOF 121, DOFPulse
      display1.text = "     CREDITS" & "  "&(Credits) &"  " & "      "
      SpecialHighScoreReward = 0
      'pupDMDDisplay "credits", "CREDITS", " "&(Credits), "", "", false, 2, 1, 1
     pupDMDDisplay "default", "CREDITS^"&Credits&"|16744448", "", 2, 1, 10

      vpmtimer.addtimer 2500, "EndOfBallComplete() '"
   End If

 ' If HighScoreReward = 0 And SpecialHighScoreReward = 0 Then

 ' End If

 End If
     Playsound "spchyeahthatsit"
     EndOfBallComplete()
End Sub

Sub UnCreditoMas
    Credits = Credits + 1
  DOF 136, DOFOn
    'PlaySound SoundFXDOF("fx_knocker",124,DOFPulse,DOFKnocker)
  PlaySound SoundFXDOF("Knocker_1",135,DOFPulse,DOFKnocker)
  DOF 121, DOFPulse
    display1.text = "     CREDITS" & "  "&(Credits) &"  " & "      "
    pupDMDDisplay "credits", "CREDITS", " "&(Credits), "", "", false, 2, 1, 1
    vpmtimer.addtimer 2000, "EndOfBallComplete() '"
End Sub


'********************
' Music sounds
'********************

Dim Song
Song = ""

Sub xxxPlaySong(name)
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


Sub PlaySong(name)
    If bMusicOn and HasPup Then
          PuPlayer.playlistplayex pMusic,"MUSIC",name&".ogg",100,10
          PuPlayer.SetLoop pMusic,1
          If name = "mu_end" Then PuPlayer.playstop pMusic
    End If
End Sub



'-----------------------------
'-----  FS Display Code  -----
'-----------------------------

'If You want to hide a display, set the reel value of every reel to 44. This picture is transparent
'This is best done using collection:
'
' If HideDisplay then
'   For Each obj in ReelsCollection:obj.setvalue(44):next
' end if


 Dim Char(32),i,TempText                    'increase dimension if You need larger displays



'-----------------------------------------------
'-----  B2S section, not used in the demo  -----
'-----------------------------------------------

Sub DisplayB2SText(TextPar)             'Procedure to display Text on a 32 digit B2S LED reel. Assuming that it is display 1 with internal digit numbers 1-32
  If B2SOn Then
  TempText = TextPar
  for i = 1 to 32
    if i <= len(TextPar) then
      Char(i) = left(TempText,1)
      TempText = right(Temptext,len(TempText)-1)
    else
      Char(i) = " "
    end if
  next
  if B2SOn Then
  for i = 1 to 32
    controller.B2SSetLED i,B2SLEDValue(Char(i))
  next
  end if
  End If
End Sub


Sub DisplayB2SText2(TextPar)              'Procedure to display Text on a 30 digit B2S LED reel. Assuming that it is display 1 with internal digit numbers 1-32
 If B2SOn Then
  TempText = TextPar
  for i = 1 to 32
    if i <= len(TextPar) then
      Char(i) = left(TempText,1)
      TempText = right(Temptext,len(TempText)-1)
    else
      Char(i) = " "
    end if
  next

  for i = 1 to 32
    controller.B2SSetLED i,B2SLEDValue(Char(i))

  next
End If

DMDUpdate.interval = 2000
DMDUpdate.enabled = 1

End Sub



Function B2SLEDValue(CharPar)           'to be used with dB2S 15-segments-LED used in Herweh's Designer
  B2SLEDValue = 0                 'default for unknown characters
  select case CharPar
    Case "","": B2SLEDValue = 0
    Case "0": B2SLEDValue = 63
    Case "1": B2SLEDValue = 8704
    Case "2": B2SLEDValue = 2139
    Case "3": B2SLEDValue = 2127
    Case "4": B2SLEDValue = 2150
    Case "5": B2SLEDValue = 2157
    Case "6": B2SLEDValue = 2172
    Case "7": B2SLEDValue = 7
    Case "8": B2SLEDValue = 2175
    Case "9": B2SLEDValue = 2159
    Case "A": B2SLEDValue = 2167
    Case "B": B2SLEDValue = 10767
    Case "C": B2SLEDValue = 57
    Case "D": B2SLEDValue = 8719
    Case "E": B2SLEDValue = 121
    Case "F": B2SLEDValue = 2161
    Case "G": B2SLEDValue = 2109
    Case "H": B2SLEDValue = 2166
    Case "I": B2SLEDValue = 8713
    Case "J": B2SLEDValue = 31
    Case "K": B2SLEDValue = 5232
    Case "L": B2SLEDValue = 56
    Case "M": B2SLEDValue = 1334
    Case "N": B2SLEDValue = 4406
    Case "O": B2SLEDValue = 63
    Case "P": B2SLEDValue = 2163
    Case "Q": B2SLEDValue = 4287
    Case "R": B2SLEDValue = 6259
    Case "S": B2SLEDValue = 2157
    Case "T": B2SLEDValue = 8705
    Case "U": B2SLEDValue = 62
    Case "V": B2SLEDValue = 17456
    Case "W": B2SLEDValue = 20534
    Case "X": B2SLEDValue = 21760
    Case "Y": B2SLEDValue = 9472
    Case "Z": B2SLEDValue = 17417
    Case "<": B2SLEDValue = 5120
    Case ">": B2SLEDValue = 16640
    Case "^": B2SLEDValue = 17414
    Case ".": B2SLEDValue = 8
    Case "!": B2SLEDValue = 0
    Case ".": B2SLEDValue = 128
    Case "*": B2SLEDValue = 32576
    Case "/": B2SLEDValue = 17408
    Case "\": B2SLEDValue = 4352
    Case "|": B2SLEDValue = 8704
    Case "=": B2SLEDValue = 2120
    Case "+": B2SLEDValue = 10816
    Case "-": B2SLEDValue = 2112
  end select
  B2SLEDValue = cint(B2SLEDValue)
End Function









Dim Ball


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same - and empty
' Sub DMDScore
' End Sub





Sub DMDUpdate_Timer
pUpdateScores
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






'**********
' UltraDMD
'**********

  Dim UltraDMD

  ' DMD using UltraDMD calls

  Sub DMD(background, toptext, bottomtext, duration)
    If turnonultradmd = 0 then exit sub
    UltraDMD.DisplayScene00 background, toptext, 15, bottomtext, 15, 14, duration, 14
    UltraDMDTimer.Enabled = 1 'to show the score after the animation/message
  End Sub

  Sub DMDScore
  End Sub

  sub dmdscoretime
    If turnonultradmd = 0 then exit sub
    UltraDMD.SetScoreboardBackgroundImage "scoreboard-DMD.jpg", 15, 7
    UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer, "Ball " & Balls
  end sub

  Sub DMDScoreNow
    If turnonultradmd = 0 then exit sub
    DMDFlush
    DMDScore
  End Sub

  Sub DMDFLush
    If turnonultradmd = 0 then exit sub
    UltraDMDTimer.Enabled = 0
    UltraDMD.CancelRendering
  End Sub

  Sub DMDScrollCredits(background, text, duration)
    If turnonultradmd = 0 then exit sub
    UltraDMD.ScrollingCredits background, text, 15, 14, duration, 14
  End Sub

  Sub DMDId(id, background, toptext, bottomtext, duration)
    If turnonultradmd = 0 then exit sub
    UltraDMD.DisplayScene00ExwithID id, False, background, toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
  End Sub

  Sub DMDMod(id, toptext, bottomtext, duration)
    If turnonultradmd = 0 then exit sub
    UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
  End Sub

  Sub UltraDMDTimer_Timer 'used for the attrackmode and the instant info.
    If turnonultradmd = 0 then exit sub
    If bInstantInfo Then
      InstantInfo
    ElseIf bAttractMode Then
    ElseIf NOT UltraDMD.IsRendering Then
      DMDScoreNow
    ElseIf bromconfig Then
      romconfig
    End If
  End Sub

  Sub DMD_Init
    If turnonultradmd = 0 then exit sub
    Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    If UltraDMD is Nothing Then
      MsgBox "No UltraDMD found.  This table will NOT run without it."
      Exit Sub
    End If

    UltraDMD.Init
    If turnonultradmd = 0 then exit sub
    If Not UltraDMD.GetMajorVersion = 1 Then
      MsgBox "Incompatible Version of UltraDMD found."
      Exit Sub
    End If

    If UltraDMD.GetMinorVersion <1 Then
    If turnonultradmd = 0 then exit sub
      MsgBox "Incompatible Version of UltraDMD found. Please update to version 1.1 or newer."
      Exit Sub
    End If

    Dim fso:Set fso = CreateObject("Scripting.FileSystemObject")
    Dim curDir:curDir = fso.GetAbsolutePathName(".")

    Dim DirName
    DirName = curDir& "\" &TableName& ".UltraDMD"

    If Not fso.FolderExists(DirName) Then _
        Msgbox "UltraDMD userfiles directory '" & DirName & "' does not exist." & CHR(13) & "No graphic images will be displayed on the DMD"
    UltraDMD.SetProjectFolder DirName

    ' wait for the animation to end
    While UltraDMD.IsRendering = True
    WEnd

  End Sub
















Sub StartAttractMode()
    bGameInPLay = False
  bAttractMode = true
    StartLightSeq
    'nevertimer1.Interval = 2500
    'nevertimer1.Enabled = true
    'attractmessage = attractmessage - attractmessage
    'messagetimer.enabled=true
    pAttractStart
End Sub

Sub StopAttractMode()
    messagetimer.enabled=False
    ShowHighScores.Enabled = 0
    bAttractMode = False
    addscore(0)
    LightSeqAttract.StopPlay
    LightSeqFlasher.StopPlay
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






dim attractmessage

sub messagetimer_Timer()
    attractmessage = attractmessage + 1

end sub



Sub ShowHighScores_Timer

End Sub














' *********************************************************************
' **                                                                 **
' **                        FUTURE PINBALL                           **
' **                                                                 **
' **                   Generic Table Script v1.0                     **
' **                                                                 **
' **     (C) 2007 Chris Leathley - BSP Software Design Solutions     **
' **                                                                 **
' ** This script should provide a simple frame work you can use in   **
' ** your own tables as a starting point.                            **
' **                                                                 **
' *********************************************************************

'Option Explicit        ' Force explicit variable declaration
  ' DispDmd1.AddFont 1, "dmd05x05p"
' DispDmd1.AddFont 2, "dmd06x07p"
' DispDmd1.AddFont 3, "dmd08x09p"
' DispDmd1.AddFont 4, "dmd08x13p"
' DispDmd1.AddFont 5, "dmd09x11po"
' DispDmd1.AddFont 6, "dmd09x15po"
' DispDmd1.AddFont 7, "indyfont"
' DispDmd1.AddFont 8, "isdmdfont1"
' DispDmd1.AddFont 9, "gooniesf1"
' DispDmd1.AddFont 10, "gooniesf2"
' DispDmd1.AddFont 11, "gooniesf3"
' DispDmd1.AddFont 12, "gooniesf4"
' DispDmd1.AddFont 13, "gooniesf5"
' DispDmd1.AddFont 14, "gooniesf6"

' Define any Constants
Const constMaxPlayers     = 4     ' Maximum number of players per game (between 1 and 4)
Const constBallSaverTime  = 7000  ' Time in which a free ball is given if it lost very quickly
                        ' Set this to 0 if you don't want this feature
Const constMaxMultiplier  = 6   ' Defines the maximum bonus multiplier level

' Define Global Variables
'
Dim PlayersPlayingGame    ' number of players playing the current game
Dim CurrentPlayer       ' current player (1-4) playing the game
Dim BonusPoints(4)      ' Bonus Points for the current player
Dim BonusMultiplier(4)    ' Bonus Multiplier for the current player
Dim BallsRemaining(4)   ' Balls remaining to play (inclusive) for each player
Dim ExtraBallsAwards(4)   ' number of EB's out-standing (for each player)

' Define Game Control Variables
Dim LastSwitchHit       ' Id of last switch hit
Dim BallsOnPlayfield      ' number of balls on playfield (multiball exclusive)
Dim BallsInLock       ' number of balls in multi-ball lock

' Define Game Flags
'Dim bFreePlay          ' Either in Free Play or Handling Credits
'Dim bOnTheFirstBall      ' First Ball (player one). Used for Adding New Players
'Dim bBallInPlungerLane   ' is there a ball in the plunger lane
'Dim bBallSaverActive     ' is the ball saver active
'Dim bMultiBallMode     ' multiball mode active ?
'Dim bEnteringAHighScore    ' player is entering their name into the high score table
dim liftno
dim ballno
dim scoreupdate
dim lastscore
dim spellarkframe
dim organhitsleft
dim tempsound
dim dataD
dim dataA
dim dataT
dim dataA2
dim slothS
dim slothL
dim slothO
dim slothT
dim slothH
dim oframestart
dim oframeend
dim ospeed
dim raidersjp
dim templejp
dim crusadejp
dim crystaljp
dim modename
dim mapmulti
dim swordhit
dim swordneeded
dim dmdfont
dim dmdspeed
dim currentframe
dim startframe
dim endframe
dim repeatdmd
dim jpscore
dim nextscenescore
dim indyscoring
dim jonesscoring
dim indyhitsleft
dim joneshitsleft
dim ballbonus
dim swordscore
dim totalbonus
dim arkhit
dim templehit
dim grailhit
dim skullhit
dim hitskull
dim mysterymode
dim hurrytotal
dim endtext
dim xmarks
dim xmarkshits
dim stonesleft
dim traphitsleft
dim stonekicker
dim indymatch
dim playermatch
dim rjackpothits
dim rjackpothitsleft
dim jackpotanim
dim templeopen
dim stonescollected
dim rampshots
dim rampshotsleft
dim loopd
dim keytotal
dim doubloon
dim slothmulti
dim jackpotscore
dim marbles
dim ballslocked
dim autoball
dim wellhit
dim wellsound
dim modetime
dim keysearch
dim boneplay
dim datachr
dim tchr
'dim temptext
dim comboon
dim combovalue
dim matchchr
dim endmatchchr
dim modejp
' ************************
'*********************************************
' **                                                                 **
' **               Future Pinball Defined Script Events              **
' **                                                                 **
' *********************************************************************


' The Method Is Called Immediately the Game Engine is Ready to
' Start Processing the Script.
'
Sub FuturePinball_BeginPlay()
  ' seed the randomiser (rnd(1) function)
  Randomize
nevertimer1.Interval = 2500
nevertimer1.Enabled = true' , 2500
lastscore =0
'LightSeq1.Play SeqRandomEffect,20,9999990
'dispdmd1.text = "[f9][xc][yc]" &chr(32)
 'overlay1.fadeout()

'Overlay1.Frame 1, 1395,1


' initalise the player display reel to the last known player scores
  'Player1Reel.SetValue(nvScore1)
  'Player2Reel.SetValue(nvScore2)
  'Player3Reel.SetValue(nvScore3)
  'Player4Reel.SetValue(nvScore4)

  ' We want the player to put in credits (Change this
  'bFreePlay = FALSE

  ' kill the last switch hit (this this variable is very usefull for game control)
  'set LastSwitchHit = DummyTrigger

  ' initialse any other flags
  bOnTheFirstBall = FALSE
  bEnteringAHighScore = FALSE
  BallsOnPlayfield = 0
  BallsInLock = 0
      SolLFlipper 0
      SolRFlipper 0
'matchgame()
  EndOfGame()
End Sub







' *********************************************************************
' **                                                                 **
' **                     User Defined Script Events                  **
' **                                                                 **
' *********************************************************************

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
  Dim i
    StopAttractMode
    bGameInPLay = True
  'AddDebugText "ResetForNewGame"

  ' get Future Pinball to zero out the nvScore (and nvSpecialScore) Variables
  ' aswell and ensure the camera is looking in the right direction.
  ' this also Sets the fpGameInPlay flag

  'BeginGame()

  'Player1Reel.ResetToZero
  'Player2Reel.ResetToZero
  'Player3Reel.ResetToZero
  'Player4Reel.ResetToZero

  ' increment the total number of games played
   TotalGamesPlayed = TotalGamesPlayed + 1

  ' Start with player 1
  CurrentPlayer = 1

  ' Single player (for now, more can be added in later)
  PlayersPlayingGame = 1

  ' We are on the First Ball (for Player One)
  bOnTheFirstBall = TRUE

  ' initialise all the variables which are used for the duration of the game
  ' (do all players incase any new ones start a game)
  For i = 1 To constMaxPlayers
    Score(i) = 0
    ' Bonus Points for the current player
    BonusPoints(i)  = 0
    ' Initial Bonus Multiplier
    BonusMultiplier(i) = 1
    ' Balls Per Game
    BallsRemaining(i) = BallsPerGame
    ' Number of EB's out-standing
    ExtraBallsAwards(i) = 0
  Next


  ' initialise any other flags
  bMultiBallMode = FALSE


    Game_Init
  ' you may wish to start some music, play a sound, do whatever at this point

  ' set up the start delay to handle any Start of Game Attract Sequence
  FirstBallDelayTimer.Interval = 500
  FirstBallDelayTimer.Enabled = TRUE

  pDMDStartGame ' setup PuP
End Sub

Sub Game_Init
modejp = false
PlaySong "mu_end"
DOF 122, DOFPulse
target10.Isdropped = 0
target11.Isdropped = 0
target12.Isdropped = 0
target13.Isdropped = 0
target14.Isdropped = 0
target15.Isdropped = 0
target16.Isdropped = 0
tchr = 34
modetime = 0
keysearch = 0
wellsound = true
wellhit = 0
autoball = false
ballslocked = 0
loopd = ""
titletimer.Enabled = false
dmdanimtimer.Enabled = false
slothmulti = 1
jackpotscore = 1000000
marbles = 0
comboon = false
combovalue = 50000

if captiveP.RotX > 15 then
stonetarget.Isdropped = false
stonetarget1.Isdropped = false
stonetarget2.Isdropped = false
'stonetarget.render = false
'captiveP.RotX = 0
captiveDown 35, 0
end if

resetxmarks.Enabled = false
resetlights()
LightSeq1.stopplay()
nevertimer1.Enabled = false
nevertimer2.Enabled = false
nevertimer3.Enabled = false
nevertimer4.Enabled = false
nevertimer5.Enabled = false
nevertimer6.Enabled = false
nevertimer7.Enabled = false
nevertimer8.Enabled = false
nevertimer9.Enabled = false
nevertimer10.Enabled = false
nevertimer11.Enabled = false

neverreset.Enabled = false
nevern.state = 0: NeverNP.visible = 0
nevere1.state = 0: NeverEP.visible = 0
neverv.state = 0: NeverVP.visible = 0
nevere2.state = 0: NeverE2P.visible = 0
neverr.state = 0: NeverRP.visible = 0
nevers.state= 0: NeverSP.visible = 0
nevera.state = 0: NeverAP.visible = 0
nevery.state = 0: NeverYP.visible = 0
neverd.state = 0: NeverDP.visible = 0
neveri.state = 0: NeverIP.visible = 0
nevere3.state = 0: NeverEP3.visible = 0
playSound "ballready"',true
light27.state = 2
light28.state = 2
light29.state = 2
light30.state = 2
light31.state = 2
light32.state = 2
light33.state = 2
light34.state = 2
light35.state = 2
light51.state = 2
light48.state = 2
liftno = 1
scoreupdate = true
lastscore =0
spellarkframe = 42
organhitsleft = 7
tempsound = ""
raidersjp = 1000000
templejp = 1000000
crystaljp = 1000000
crusadejp = 1000000
modename = ""
mapmulti = 1
swordhit = 0
swordneeded = 5
nextscenescore = 1000000
replayscored = false
indyscoring = false
jonesscoring = false
indyhitsleft = 0
joneshitsleft = 0
ballbonus = 0
totalbonus = 0
arkhit = 0
templehit = 0
grailhit = 0
skullhit = 0
hitskull = true
mysterymode = ""
xmarks = 1
xmarkshits = 3
traphitsleft = 2
stonesleft = 3
stonekicker = false
jackpotanim = false
templeopen = false
stonescollected = 0
rampshots = 0
rampshotsleft = 5
dataD = " "
dataA = " "
dataT = " "
dataA2 = " "
slothS = " "
slothL = " "
slothO = " "
slothT = " "
slothH = " "
if fakedoor.isdropped = true then
resetdoor.Enabled = true ', 100
end if

End Sub


sub checkreplay()

    if (Score(CurrentPlayer) > replayscore) and (replayscored = false) then
      pupDMDDisplay "replay", "REPLAY", "", 4, 1, 10
            scorereplay()
    end if

end sub

sub scorereplay()
  Credits = Credits + 1
  DOF 136, DOFOn
    'PlaySound SoundFXDOF("fx_knocker",124,DOFPulse,DOFKnocker)
  PlaySound SoundFXDOF("fx_knocker",136,DOFPulse,DOFKnocker)
  DOF 121, DOFPulse
  replayscored = true
end sub


' This Timer is used to delay the start of a game to allow any attract sequence to
' complete.  When it expires it creates a ball for the player to start playing with
'
Sub FirstBallDelayTimer_Timer()
  ' stop the timer
  FirstBallDelayTimer.Enabled = FALSE

  ' reset the table for a new ball
  ResetForNewPlayerBall()

  ' create a new ball in the shooters lane
  CreateNewBall()
End Sub


' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))
'
Sub ResetForNewPlayerBall()
combovalue = 50000
autoball = false
AddScore(0)
light49.state = 0
arkhit = 0
templehit = 0
grailhit = 0
skullhit = 0
hitskull = true
scoreupdate = true
marbles = 0
light50.state = 2
light40.state = 0
light37.state = 0
light26.state = 0
totalbonus = 0
AddScore(0)
PlaySong "plungerloop"
mapmulti = 1
light20.state = 0
light21.state = 0
light22.state = 0
xmarks = 1
xmarkshits = 3
jackpotanim = false
  ' set the current players bonus multiplier back down to 1X
  SetBonusMultiplier(1)
startskill()
  ' reset any drop targets, lights, game modes etc..
  ShootAgainLight.State = 0
keytotal = 0
doubloon = 0
rampshots = 0
if currentplayer = 1 then
ballbonus = ballbonus + 125000
end if
End Sub



' Create a new ball on the Playfield
'
Sub CreateNewBall()
  ' create a ball in the plunger lane kicker.
  PlungerKicker.CreateBall
    LastSwitchHit  = ""
  ' There is a (or another) ball on the playfield
  BallsOnPlayfield = BallsOnPlayfield + 1
  'Ballsindrain = Ballsindrain - 1

  'ballindrain = false

  ' kick it out..
  PlungerKicker.kick 90, 8
    RandomSoundBallRelease PlungerKicker
    pDMDStartBall 'PuPDMD
End Sub


' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded
'
Sub EndOfBall()
  Dim BonusDelayTime

  'AddDebugText "EndOfBall"

  ' the first ball has been lost. From this point on no new players can join in
  bOnTheFirstBall = FALSE

  ' only process any of this if the table is not tilted.  (the tilt recovery
  ' mechanism will handle any extra balls or end of game)
  If (Tilted = FALSE) Then
    Dim AwardPoints

    ' add in any bonus points (multipled by the bunus multiplier)
    AwardPoints = BonusPoints(CurrentPlayer) * BonusMultiplier(CurrentPlayer)
    AddScore(AwardPoints)
    'AddDebugText "Bonus Points = " & AwardPoints

    ' you may wish to do some sort of display effects which the bonus is
    ' being added to the players score

    ' add a bit of a delay to allow for the bonus points to be added up
    BonusDelayTime = 4500
      if keytotal > 0 then
    bonusdelaytime = bonusdelaytime + 1500
    end if
if doubloon > 0 then
    bonusdelaytime = bonusdelaytime + 1500
    end if

if grailhit > 0 then
    bonusdelaytime = bonusdelaytime + 1500
    end if
      starteob()
      if rampshots > 0 then
    bonusdelaytime = bonusdelaytime + 1500
    end if

if marbles > 0 then
    bonusdelaytime = bonusdelaytime + 1500
end if


      starteob()


  Else
    ' no bonus, so move to the next state quickly
    BonusDelayTime = 20
  End If

  ' start the end of ball timer which allows you to add a delay at this point
  EndOfBallTimer.Interval = BonusDelayTime
  EndOfBallTimer.Enabled = TRUE
End Sub


' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the currentplayer)
'
Sub EndOfBallTimer_Timer()
  ' disable the timer
  EndOfBallTimer.Enabled = FALSE
dim skilldelay
skilldelay = 100
  ' if were tilted, reset the internal tilted flag (this will also
  ' set fpTiltWarnings back to zero) which is useful if we are changing player LOL
  Tilted = FALSE

  ' has the player won an extra-ball ? (might be multiple outstanding)
  If (ExtraBallsAwards(CurrentPlayer) <> 0) Then

    'AddDebugText "Extra Ball"


    ' yep got to give it to them
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1
pupDMDDisplay "shoot", "SHOOT AGAIN!", "@shootagain.mp4", 3, 1, 10

skilldelay = skilldelay +2000
    ' if no more EB's then turn off any shoot again light
    If (ExtraBallsAwards(CurrentPlayer) = 0) Then
      ShootAgainLight.State = 0


    End If
PlaySong "plungerloop"',true
startskilltimer.Enabled = skilldelay
    ' You may wish to do a bit of a song and dance at this point

    ' Create a new ball in the shooters lane


  Else  ' no extra balls

    BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

    ' was that the last ball ?
    If (BallsRemaining(CurrentPlayer) <= 0) Then

      'AddDebugText "No More Balls, High Score Entry"

      ' Submit the currentplayers score to the High Score system built into Future Pinball
      ' if they have gotten a high score then it will ask them for their initials.  If not
      ' it will call the FuturePinball_NameEntryComplete right away
      'bEnteringAHighScore = TRUE
      'EnterHighScore(CurrentPlayer)
            CheckHighscore
      ' you may wish to play some music at this point

    Else

      ' not the last ball (for that player)
      ' if multiple players are playing then move onto the next one
      EndOfBallComplete()

    End If
  End If
End Sub

sub startskilltimer_Timer()
startskilltimer.Enabled = false
startskill()
CreateNewBall()
end sub

' This function is called when the end of bonus display
' (or high score entry finished) and it either end the game or
' move onto the next player (or the next ball of the same player)
'
Sub EndOfBallComplete()
  Dim NextPlayer

  'AddDebugText "EndOfBall - Complete"

  ' are there multiple players playing this game ?
  If (PlayersPlayingGame > 1) Then
    ' then move to the next player
    NextPlayer = CurrentPlayer + 1
    ' are we going from the last player back to the first
    ' (ie say from player 4 back to player 1)
    If (NextPlayer > PlayersPlayingGame) Then
      NextPlayer = 1
    End If
  Else
    NextPlayer = CurrentPlayer
  End If

  'AddDebugText "Next Player = " & NextPlayer

   ' is it the end of the game ? (all balls been lost for all players)
  If ((BallsRemaining(CurrentPlayer) <= 0) And (BallsRemaining(NextPlayer) <= 0)) Then
    ' you may wish to do some sort of Point Match free game award here
    ' generally only done when not in free play mode

    ' set the machine into game over mode

      matchgame() :PuPEvent 809
     ' EndOfGame()

    ' you may wish to put a Game Over message on the

  Else
    ' set the next player
    CurrentPlayer = NextPlayer

    ' make sure the correct display is upto date
    AddScore(0)

    ' reset the playfield for the new player (or new ball)
    ResetForNewPlayerBall()

    ' and create a new ball
    CreateNewBall()

  End If
End Sub


' This frunction is called at the End of the Game, it should reset all
' Drop targets, and eject any 'held' balls, start any attract sequences etc..
Sub EndOfGame()
  'AddDebugText "End Of Game"
  ' let Future Pinball know that the game has finished.
  ' This also clear the fpGameInPlay flag.
  'EndGame()

if pINAttract=false Then pupDMDDisplay "GAMEOVER", "GAME OVER", "@gameover.mp4", 3, 1, 50   'this endofgame is called in attract for some reason

Select case (RandomNumber(4))
case 1: PlaySong  "intro"
case 2: PlaySong  "goodenough"
case 3: PlaySong  "intro"
case 4: PlaySong  "goodenough"
end select

'LightSeq1.Play SeqRandom,20,9999990
nevertimer1.interval = 2000
nevertimer1.Enabled = true
'dispdmd1.text = "[f9][xc][yc]" &chr(32)
matchchr = 32
endmatchchr = 89
dmdfont = " "
dmdspeed = 80
'dmdanimtimer.set true , 100
'endtext = "[f1] "
  ' ensure that the flippers are down
' LeftFlipper.RotateToStart
' RightFlipper.RotateToStart

if ballslocked > 0 then
SolLFlipper 0
SolRFlipper 0
LeftFlipper.RotateToStart
RightFlipper.RotateToStart
ballsonplayfield = ballslocked
popup1.IsDropped = 1
Popup.TransY = -25
PlaySound SoundFXDOF("DiverterOn", 130, DOFPulse, DOFContactors)
end if

  ' turn off the reel bulbs
  'Player1Reel.State = BulbOff
  'Player2Reel.State = BulbOff
  'Player3Reel.State = BulbOff
  'Player4Reel.State = BulbOff

  ' set any lights for the attract mode
  SetAllLightsForAttractMode()
StartAttractMode
  ' you may wish to light any Game Over Light you may have
End Sub


' The tilt recovery timer waits for all the balls to drain before continuing on
' as per normal
'
Sub TiltRecoveryTimer_Timer()
  ' disable the timer
  TiltRecoveryTimer.Enabled = FALSE
  ' if all the balls have been drained then..
  If (BallsOnPlayfield = 0) Then
    ' do the normal end of ball thing (this dosn't give a bonus if the table is tilted)
    EndOfBall()
  Else
    ' else retry (checks again in another second)
    TiltRecoveryTimer.Interval = 1000
    TiltRecoveryTimer.Enabled = TRUE
  End If
End Sub


' Set any lights for the Attract Mode.
'
Sub SetAllLightsForAttractMode()
  'ShootAgainLight.State = BulbBlink

End Sub

Sub Endattractmode()
    bAttractMode = False
End Sub

' *********************************************************************
' **                                                                 **
' **                   Drain / Plunger Functions                     **
' **                                                                 **
' *********************************************************************

' lost a ball ;-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count and test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one
'
Sub Drain_Hit()
  DOF 212, DOFPulse
  ' Destroy the ball
  Drain.DestroyBall
  BallsOnPlayfield = BallsOnPlayfield - 1
  ' pretend to knock the ball into the ball storage mech
  RandomSoundDrain Drain
if (BallsRemaining(CurrentPlayer) <= 0) then
exit sub
end if

  ' if there is a game in progress and
  If (bGameInPlay = TRUE) And (Tilted = FALSE) Then

    ' is the ball saver active,
    If (bBallSaverActive = TRUE) Then

      ' yep, create a new ball in the shooters lane
      CreateNewBall()
         'plunger.pull()
         plungertimer.Interval = 1000
         plungertimer.Enabled = true', 1000
         pupDMDDisplay "ballsave", "BALL SAVED!", "@ballsaved.mp4", 3, 1, 10 :PuPEvent 802
         playsound "wait"
    Else

      ' cancel any multiball if on last ball (ie. lost all other balls)
      '
      If (BallsOnPlayfield = 1) Then
        ' and in a multi-ball??
        If (bMultiBallMode = True) then
if modename = "trapmulti" then

light39.state = 0
light49.state = 0
light44.state = 0
end if
if modename = "wizard" then
endwizard
end if
                restartmusic.Interval = 500
                restartmusic.Enabled = true', 500
                flushdmdtimer.Interval = 500
                flushdmdtimer.Enabled = true' ,500


                modename = ""
                resetrjackpot.Enabled = false
          ' not in multiball mode any more
          bMultiBallMode = False
               light36.state = 0
light37.state = 0
'light38.state = bulboff
light39.state = 0
light50.state = 0
light51.state = 0
          ' you may wish to change any music over at this point and
          ' turn off any multiball specific lights
        End If
      End If

      ' was that the last ball on the playfield
      If (BallsOnPlayfield = 0) Then
MainSeq.play seqalloff
if modename = "key" or modename = "boulder" or modename = "water" or modename = "fight" then
endmode
end if
if modename = "bone" then
resetdoor.Interval = 1000
resetdoor.Enabled = true ', 1000
endmode
end if
if slothmulti = 2 then


endsloth
end if
if modename = "truffle" then
endtruffle
modename = ""
end if


if modename = "sword" then
modename = ""
swordmodetimer.Enabled = false

light47.state = 0
swordmodetimer2.Enabled = false

swordneeded = swordneeded + 3
swordhit = 0

end if

if modename = "rats" or modename = "ants" or modename = "snakes" then
snakestimer1.Enabled = false
ratstimer1.Enabled = false
antstimer1.Enabled = false
snakestimer2.Enabled = false
snakesendtimer.Enabled = false
light50.state = 0
light51.state = 0
ratsendtimer.Enabled = false
ratstimer2.Enabled = false
antstimer2.Enabled = false
antsendtimer.Enabled = false

scoreupdate = true
modename = ""
restartmusic.Enabled = true' , 200
end if

if indyscoring = true then
endindy()
end if
if jonesscoring = true then
endjones()
end if


        ' handle the end of ball (change player, high score entry etc..)
        EndOfBall()
      End If

    End If
  End If
End Sub


' A ball is pressing down the trigger in the shooters lane
'
Sub PlungerLaneTrigger_Hit()
  bBallInPlungerLane = True
    Bulb1.State = 2
  ' remember last trigger hit by the ball
  LastSwitchHit = "PlungerLaneTrigger"
End Sub

Sub PlungerLaneTrigger_Unhit()
  bBallInPlungerLane = False
End Sub

' A Ball may of rolled into the Plunger Lane Kicker, if so then kick it
' back out again
'
'Sub PlungerKicker_Hit()
' PlungerKicker.Kick 90, 15
'playsound "fx_kicker"
'End Sub


' The Ball has rolled out of the Plunger Lane.  Check to see if a ball saver machanisim
' is needed and if so fire it up.
'
Sub PlungerLaneGate_Hit()
  ' if there is a need for a ball saver, then start off a timer
  ' only start if it is currently not running, else it will reset the time period
  If (constBallSaverTime <> 0) And (bBallSaverActive <> TRUE) Then
    ' and only if the last trigger hit was the plunger wire.
    ' (ball in the shooters lane)
    If (LastSwitchHit.Name = "PlungerLaneTrigger") Then
      ' set our game flag
      bBallSaverActive = TRUE
      ' start the timer
      BallSaverTimer.Enabled = FALSE
      BallSaverTimer.Interval = constBallSaverTime
      BallSaverTimer.Enabled = TRUE
      ' if you have a ball saver light you might want to turn it on at this
      ' point (or make it flash)
    End If
  End If
End Sub


' The ball saver timer has expired.  Turn it off and reset the game flag
'
Sub BallSaverTimer_Timer()
  ' stop the timer from repeating
  BallSaverTimer.Enabled = FALSE
  ' clear the flag
  bBallSaverActive = FALSE
  ' if you have a ball saver light then turn it off at this point

if (ExtraBallsAwards(CurrentPlayer) > 0) Then
shootagainlight.state = 1
Else
shootagainlight.state = 0
End If
End Sub



' *********************************************************************
' **                                                                 **
' **                   Supporting Score Functions                    **
' **                                                                 **
' *********************************************************************

' Add points to the score and update the score board
'
Sub AddScore(points)
  If (Tilted = FALSE) Then

    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + ((points)*slothmulti)
lastscore = points
    ' update the score displays

if scoreupdate = true and modename <> "truffle" then
    ' add the points to the correct display and light the current players display
      DispDmd1.Text = "[f7][xc][y3]" & formatscore(score(currentplayer)) & "BALL:" & (6 -BallsRemaining(CurrentPlayer)) & " PLAYER:" & currentplayer & "CREDITS:" & credits  & " "
            'DispDmd1.State = BulbOn


  End if
if scoreupdate = true and modename = "truffle" then
    ' add the points to the correct display and light the current players display
      DispDmd1.Text = "[f2][x5][y12]" & formatscore(score(currentplayer)) & "TRUFFLE SHUFFLE" & "TARGETS WORTH 50,000" & " " & " "
            pupDMDDisplay "truffle", "TRUFFLE SHUFFLE", "", 3, 1, 10

            'DispDmd1.State = BulbOn
  End if
if scoreupdate = true and slothmulti = 2 then
    ' add the points to the correct display and light the current players display
DispDmd1.Text =  "2X SCORING" & " " & chr(42) & " " & formatscore(score(currentplayer))
            'DispDmd1.State = BulbOn
  End if
end if
        DMDScore
        checkreplay()
  ' you may wish to check to see if the player has gotten a replay
End Sub


' Add some points to the current Jackpot.
'
Sub AddJackpot(points)
  ' Jackpots only generally increment in multiball mode and not tilted
  ' but this dosn't have to tbe the case
  If (Tilted = False) Then

    If (bMultiBallMode = TRUE) Then
      Jackpot = Jackpot + points
      ' you may wish to limit the jackpot to a upper limit, ie..
      ' If (nvJackpot >= 6000) Then
      '   nvJackpot = 6000
      '   End if
    End if
  End if
End Sub


' Will increment the Bonus Multiplier to the next level
'
Sub IncrementBonusMultiplier()
  Dim NewBonusLevel

  ' if not at the maximum bonus level
  if (BonusMultiplier(CurrentPlayer) < constMaxMultiplier) then
    ' then set it the next next one and set the lights
    NewBonusLevel = BonusMultiplier(CurrentPlayer) + 1
    SetBonusMultiplier(NewBonusLevel)
   End if
End Sub


' Set the Bonus Multiplier to the specified level and set any lights accordingly
'
Sub SetBonusMultiplier(Level)
  ' Set the multiplier to the specified level
  BonusMultiplier(CurrentPlayer) = Level

  ' If the multiplier is 1 then turn off all the bonus lights
  If (BonusMultiplier(CurrentPlayer) = 1) Then
    ' insert your own code here
  Else
    ' there is a bonus, turn on all the lights upto the current level
    If (BonusMultiplier(CurrentPlayer) >= 2) Then
      ' insert your own code here
    End If
    ' etc..
  End If
End Sub



' *********************************************************************
' **                                                                 **
' **                     Table Object Script Events                  **
' **                                                                 **
' *********************************************************************

' The Left Slingshot has been Hit, Add Some Points and Flash the Slingshot Lights
'
Dim LStep,  LStep1, RStep,  RStep1
Sub LeftSlingshotRubber_Slingshot()
  ' add some points
playsound "unknownnotneeded4"
if modename = "rscene1" or modename = "rscene2" then
raidershit()
exit sub
end if
  AddScore(1000)
  ' flash the lights around the slingshot
  FlashForMs LeftSlingshotBulb1, 100, 50, 1
  FlashForMs LeftSlingshotBulb2, 100, 50, 1

  'flasher3.FlashForMs 200, 50, BulbOff
'flasher3b.FlashForMs 200, 50, BulbOff
'if light17.state = bulbon then
'light17.state = bulboff
'light19.state = bulbon
'exit sub
'end if
'if light19.state = bulbon then
'light19.state = bulboff
'light17.state = bulbon
'end if
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingshotRubber.TimerEnabled = True
End Sub

Sub LeftSlingshotRubber_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14: PlaySoundAt SoundFXDOF("fx_slings", 103, DOFPulse, DOFContactors), LeftInLaneTrigger:DOF 200, DOFPulse
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingshotRubber.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub



Sub LeftSlingshotRubber1_Slingshot()
  ' add some points
    LeftSling004.Visible = 1
    Lemk001.RotX = 26
    LStep1 = 0
  RandomSoundSlingshotLeft Lemk001
    LeftSlingshotRubber1.TimerEnabled = True
End Sub

Sub LeftSlingshotRubber1_Timer
    Select Case LStep1
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk001.RotX = 14: PlaySoundAt SoundFXDOF("fx_slings", 104, DOFPulse, DOFContactors), Trigger3
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk001.RotX = 2
        Case 3:LeftSLing002.Visible = 0:Lemk001.RotX = -10:LeftSlingshotRubber1.TimerEnabled = 0
    End Select

    LStep1 = LStep1 + 1
End Sub

' The Right Slingshot has been Hit, Add Some Points and Flash the Slingshot Lights
'
Sub RightSlingshotRubber_Slingshot()
  ' add some points
playsound "unknownnotneeded3"
if modename = "rscene1" or modename = "rscene2" then
raidershit()
exit sub
end if
  AddScore(1000)
  ' flash the lights around the slingshot
  FlashForMs RightSlingshotBulb1, 100, 50, 1
  FlashForMs RightSlingshotBulb2, 100, 50, 1
  'flasher2.FlashForMs 200, 50, BulbOff
'flasher2b.FlashForMs 200, 50, BulbOff
'if light17.state = bulbon then
'light17.state = bulboff
'light19.state = bulbon
'exit sub
'end if
'if light19.state = bulbon then
'light19.state = bulboff
'light17.state = bulbon
'end if
  RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
  RandomSoundSlingshotRight Remk
    RightSlingShotRubber.TimerEnabled = True
End Sub

Sub RightSlingshotRubber_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14: PlaySoundAt SoundFXDOF("fx_slings", 106, DOFPulse, DOFContactors), RightInLaneTrigger:DOF 201, DOFPulse
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShotRubber.TimerEnabled = 0
    End Select

    RStep = RStep + 1
End Sub


Sub RightSlingshotRubber1_Slingshot()
  ' add some points
    RightSling004.Visible = 1
    Remk001.RotX = 26
    RStep1 = 0
  RandomSoundSlingshotRight Remk001
    RightSlingShotRubber1.TimerEnabled = True
End Sub

Sub RightSlingshotRubber1_Timer
    Select Case RStep1
        Case 1:RightSling004.Visible = 0:RightSLing003.Visible = 1:Remk001.RotX = 14: PlaySoundAt SoundFXDOF("fx_slings", 105, DOFPulse, DOFContactors), Trigger003
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk001.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk001.RotX = -10:RightSlingShotRubber1.TimerEnabled = 0
    End Select

    RStep1 = RStep1 + 1
End Sub


'
' The Left InLane trigger has been Hit
'
Sub LeftInLaneTrigger_Hit()
 DOF 207, DOFPulse
 PlaySoundAt "fx_sensor", LeftInLaneTrigger
if light42.state = 0 then
light42.state = 1
checkrich
end if

if modename = "rscene1" or modename = "rscene2" then
raidershit()
exit sub
end if
  ' add some points
  AddScore(1000)
  ' remember last trigger hit by the ball
  set LastSwitchHit = LeftInLaneTrigger
End Sub


' The Right InLane trigger has been Hit
'
Sub RightInLaneTrigger_Hit()
 DOF 208, DOFPulse
 PlaySoundAt "fx_sensor", RightInLaneTrigger
  ' add some points
'if modename = "rscene1" or modename = "rscene2" then
'raidershit()
'exit sub
'end if
  AddScore(1000)
if light18.state = 0 then
light18.state = 1
checkrich
exit sub
end if
'if light18.state = bulbon then
'light18.state = bulboff
'bulb1.state = bulbon
'lightseq1.Play Seqblinking, , 5,50
'mystery lit
'playsound "unknownnotneeded10"
'exit sub
'end if
  ' remember last trigger hit by the ball
  set LastSwitchHit = RightInLaneTrigger
End Sub


' The Left OutLane trigger has been Hit
'
Sub LeftOutLaneTrigger_Hit()
DOF 206, DOFPulse
Select case (RandomNumber(2))
case 1: PlaySoundAt "jerkalert", LeftOutLaneTrigger
case 2: PlaySoundAt "chunkawshit", LeftOutLaneTrigger
end select

'playsound "ballrollinga"
  ' add some points
if light17.state = 0 then
light17.state = 1
checkrich
end if

'if modename = "rscene1" or modename = "rscene2" then
'raidershit()
'exit sub
'end if
  AddScore(10000)
'if light17.state = bulbon then
'collectspecial()
'exit sub
'end if

select case (randomnumber(7))
case 1: PlaySoundAt "adiossenoir", LeftOutLaneTrigger
case 2: PlaySoundAt "aghhaghh", LeftOutLaneTrigger
case 3: PlaySoundAt "andwhosgonnacometosaveyou", LeftOutLaneTrigger
case 4: PlaySoundAt "dontcallmejunior", LeftOutLaneTrigger
case 5: PlaySoundAt "fxthenscream", LeftOutLaneTrigger
case 6: PlaySoundAt "youcallthisarcheaology", LeftOutLaneTrigger
case 7: PlaySoundAt "thisidhowwesaygoodbye", LeftOutLaneTrigger
end select
  ' remember last trigger hit by the ball
  set LastSwitchHit = LeftOutLaneTrigger
End Sub


' The Right OutLane trigger has been Hit
'
Sub RightOutLaneTrigger_Hit()
DOF 209, DOFPulse
Select case (RandomNumber(4))
case 1: PlaySoundAt "jerkalert", RightOutLaneTrigger
case 2: PlaySoundAt "chunkawshit", RightOutLaneTrigger
end select


if light19.state = 0 then
light19.state = 1
checkrich
end if

'if modename = "rscene1" or modename = "rscene2" then
'raidershit()
'exit sub
'end if
  ' add some points
  AddScore(10000)
'if light19.state = bulbon then
'collectspecial()
'exit sub
'end if
select case (randomnumber(7))
case 1: PlaySoundAt "adiossenoir", RightOutLaneTrigger
case 2: PlaySoundAt "aghhaghh", RightOutLaneTrigger
case 3: PlaySoundAt "andwhosgonnacometosaveyou", RightOutLaneTrigger
case 4: PlaySoundAt "dontcallmejunior", RightOutLaneTrigger
case 5: PlaySoundAt "fxthenscream", RightOutLaneTrigger
case 6: PlaySoundAt "youcallthisarcheaology", RightOutLaneTrigger
case 7: PlaySoundAt "thisidhowwesaygoodbye", RightOutLaneTrigger
end select
  ' remember last trigger hit by the ball
  set LastSwitchHit = RightOutLaneTrigger
End Sub



'kicker3.createball  255,255,159
'kicker4.createball 255,255,159
'kicker5.createball 255,255,159
'kicker7.createball()
'kicker8.createball()
'kicker9.createball()
'kicker10.createball()



function FormatScore(num)
    Dim n, f, s
    n = CStr(num)
    f = ""

    do while len(n)>3
        if len(f)>0 then
            f = Right(n, 3) & "," & f
        else
            f = Right(n, 3)
        end if
        n = Left(n, Len(n)-3)
    loop
    if len(n)>0 then
        if len(f) > 0 then
            f = n & "," & f
        else
            f = n
        end if
    end if
    FormatScore = f
End Function

function RandomNumber(ByVal max)
  RandomNumber = Int(max * Rnd + 1)
end function
sub delayintro_Timer()
delayintro.Enabled = false
intro()
end sub


'Skill shot gate

sub trigger5_hit()
'fakegate1hit
If Tilted Then exit sub
'playsound "fx_wire"
balldropsound.Enabled = true ', 500
if modename <>"" Or Tilted then
exit sub
end if
if modename = "" and autoball = false then
restartmusic.Enabled = true',200


skilltimer1.Enabled = false
skilltimer2.Enabled = false
skilltimer3.Enabled = false
skilltimer4.Enabled = false
skilltimer5.Enabled = false
addscore(1020)
end if
autoball = false
      bBallSaverActive = TRUE
      ' start the timer
      BallSaverTimer.Enabled = FALSE
      BallSaverTimer.Interval = constBallSaverTime
      BallSaverTimer.Enabled = TRUE
         shootagainlight.state = 2
if light52.state = 1 then
'indy jones scoring

startsloth
light52.state = 0

lightseq1.Play Seqblinking, , 5,50
end if

if light53.state = 1 then
startsword()

lightseq1.Play Seqblinking, , 5,50
light53.state = 0
end if
if light54.state = 1 then
' light mystery
playsound "unknownnotneeded10"
light41.state = 2
scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true' , 2000

lightseq1.Play Seqblinking, , 5,50
DispDmd1.Text = " DATA GADGET" & " LIT" & " " & chr(36) :PuPEvent 810
pupDMDDisplay "default", "DATA GADGET LIT", "@DataGadgetLit.mp4", 3, 1, 10
light54.state = 0
light27.state = 1
light28.state = 1
light29.state = 1
light30.state = 1

end if
if light55.state = 1 then
' light superpops
playsound "elephant"
scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true' , 2000

lightseq1.Play Seqblinking, , 5,50
DispDmd1.Text = " SUPER" & " POPS "
'pDMDSplashBig "SUPER POPS", 3, 33023
pupDMDDisplay "default", "SUPER POPS", "@SuperPops.mp4", 3, 1, 10
Bumper1L.state = 2
bumper2l.state = 2
bumper3l.state = 2
bumper4l.state = 2
light55.state = 0
end if
if light56.state = 1 then

DispDmd1.Text = " TRAP" & " OPEN "
pupDMDDisplay "default", "TRAP^OPEN", "@TrapOpen.mp4", 3, 1, 10
if templeopen = false then
opencaptive()
end if

'next scene
'playsound "youcheatdrjones"
'scoreupdate = false
'dmdspeed = 100
'startframe = 2
'endframe = 6
'repeatdmd = true
'dmdfont = "[f8]"
'showjptimer.set true, 2000
'flushdmdtimer.set true ,3500
'startanimation()
'jpscore = nextscenescore
'nextscenescore = nextscenescore  + 1000000
'nextscene()
light56.state = 0

lightseq1.Play Seqblinking, , 5,50
end if

end sub



sub restartmusic_Timer()
restartmusic.Enabled = false
PlaySong "mu_end"

if modename = "" and light41.state = 0 then
PlaySong "mu_end"
PlaySong "maintheme"',true
end if

if slothmulti = 2 and modename = "" then
PlaySong "mu_end"
PlaySong "slothmulti"', true
end if

if light41.state = 2 and modename = "" then
PlaySong "mu_end"
PlaySong "data" ', true
end if

if modename = "boulder" then
PlaySong "mu_end"
PlaySong "boulder"' , true
end if

if modename = "bone" then
PlaySong "mu_end"
PlaySong "bonemusic"' , true
end if

if modename = "fight" then
PlaySong "mu_end"
PlaySong "fratfight" ' , true
end if

if modename = "key" then
PlaySong "mu_end"
PlaySong "keymusic"', true
end if

if modename = "truffle" then
PlaySong "mu_end"
PlaySong "slothmulti"' , true
end if

if modename = "water" then
PlaySong "mu_end"
PlaySong "waterslide" ' , true
end if

if modename = "sword" then
PlaySong "mu_end"
PlaySong "fratelies"', true
end if

if modename = "trapmulti" then
PlaySong "mu_end"
PlaySong "multiball"', true
end if

if modename = "wizard" then
PlaySong "mu_end"
PlaySong "goodenough"', true
end if

end sub

sub openark()
'lift1.dropped = false
'lift1.collidable = true
'lifttimer.set true , 500
'if lid.angleyz = 0 then
'lid.rotateyz 50,30
'end if
'if lid.angleyz = 30 then
'lid.rotateyz -50,0
'end if
end sub


'hit angle = 335
'open angle 245
'swordsman.anglexz = 245




'skillshot
sub startskill()
light53.state = 0
light54.state = 0
light55.state = 0
light56.state = 0


DispDmd1.Text =  "SKILLSHOT" & "SHOOT HERE FOR" & "SLOTH"

pupDMDDisplay "default", "SKILLSHOT^SHOOT HERE^FOR SLOTH", "", 2, 0, 10
light52.state = 1
skilltimer1.Interval = 500
skilltimer1.Enabled = true',500
end sub

sub skilltimer1_Timer()
skilltimer1.Enabled = false
'swordsman
DispDmd1.Text =  "SKILLSHOT"  &" SHOOT HERE FOR" & " FRATELLIS"

pupDMDDisplay "default", "SKILLSHOT^SHOOT HERE^FOR FRATELLIS", "", 2, 0, 10

light52.state = 0
light53.state = 1
skilltimer2.Interval = 500
skilltimer2.Enabled = true',500
end sub
sub skilltimer2_Timer()
skilltimer2.Enabled = false
'mystery

DispDmd1.Text =  "SKILLSHOT"  &"SHOOT HERE FOR" & " DATA"
pupDMDDisplay "default", "SKILLSHOT^SHOOT HERE^FOR DATA", "", 2, 0, 10

light53.state = 0
light54.state = 1
skilltimer3.Interval = 500
skilltimer3.Enabled = true',500
end sub
sub skilltimer3_Timer()
skilltimer3.Enabled = false
'super pops

DispDmd1.Text =  "SKILLSHOT"  &" SHOOT HERE FOR" & " SUPER POPS"
pupDMDDisplay "default", "SKILLSHOT^SHOOT HERE^FOR SUPER POPS", "", 2, 0, 10

light54.state = 0
light55.state = 1
skilltimer4.Interval = 500
skilltimer4.Enabled = true',500
end sub

sub skilltimer4_Timer()
skilltimer4.Enabled = false
'next scene

DispDmd1.Text =  "SKILLSHOT"  &" SHOOT HERE FOR" & " OPEN TRAP"
pupDMDDisplay "default", "SKILLSHOT^SHOOT HERE^FOR OPEN TRAP", "", 2, 0, 10

light55.state = 0
light56.state = 1
skilltimer5.Interval = 500
skilltimer5.Enabled = true',500
end sub

sub skilltimer5_Timer()
skilltimer5.Enabled = false
'indy jones
DispDmd1.Text =  "SKILLSHOT"  &" SHOOT HERE FOR" & " SLOTH"
pupDMDDisplay "default", "SKILLSHOT^SHOOT HERE^FOR SLOTH", "", 2, 0, 10

light52.state = 1
light56.state = 0
skilltimer1.Enabled = true',500
end sub

sub flushdmdtimer_Timer()
modejp = false
repeatdmd = false
scoreupdate = true
dmdtimer.Enabled = false
flushdmdtimer.Enabled = false
'DispDmd1.Text = "[xc] [y3] [f7]" & formatscore(nvscore(currentplayer)) & "[f1][x0][y26]BALL:" & (6 -BallsRemaining(CurrentPlayer)) & "[f1][x40][y26]PLAYER:" & currentplayer & "[f1][x86][y26]CREDITS:" & nvcredits
dispdmd1.text = " "
addscore(0)
endtext = ""
end sub

'************************************************************************
'transmap
'************************************************************************
sub kicker7_hit()
PlaySoundAt SoundFXDOF("fx15", 114, DOFPulse, DOFContactors), Kicker7
Kicker7Fake.enabled = 0
if light46.state = 2 then
startwizard
exit sub
end if

if bulb4.state = 2 and modename ="bone" then
bonehit
exit sub
end if
if modename <> "" or slothmulti = 2 then
kicker7solenoidpulse()
exit sub
end if


if light43.state = 2 then
starttransmap


exit sub


end if
end sub


Sub kicker7solenoidpulse()
    PlaySound "salidadebola"
    Kicker7.Kick 0, 30, 1.5
  DOF 112, DOFPulse
  DOF 121, DOFPulse
    Kicker7FakeTimer.interval = 800
    Kicker7FakeTimer.enabled = True
End Sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
'Sub Kicker1FakeTimer_Timer
'    Kicker7FakeTimer.enabled = False
'    Kicker7Fake.enabled = 1
'End Sub

sub starttransmap()
PlaySong "mu_end"
restartmusic.Enabled = true ', 15000
kicker7timer.Enabled = true ', 15000
temptext = "TRANSLATE MAP"' & " "' & chr(40)
DispDmd1.Text =  temptext
pupDMDDisplay "default", "TRANSLATE^MAP", "@TranslateMap.mp4", 3, 1, 10

transdisptimer.Enabled = true ', 500

DOF 122, DOFPulse
target10.Isdropped = 0
target11.Isdropped = 0
target12.Isdropped = 0
target13.Isdropped = 0
target14.Isdropped = 0
target15.Isdropped = 0
target16.Isdropped = 0
if neverA.state = 0 then
neverA.state = 1
playsound "i_have_the_key"
'checknever
'startwatermode.set true , 12000
startkeymode.Interval = 8000
startkeymode.Enabled = true ', 8000
restartmusic.Interval = 15500
restartmusic.Enabled = true', 15500
kicker7timer.Interval = 17000
kicker7timer.Enabled = true', 17000
exit sub
end if

if nevery.state = 0 then
playsound "mouthintruders"
'dodge boulders
nevery.state = 1
'checknever
startbouldermode.Interval = 12000
startbouldermode.Enabled = true ', 12000
restartmusic.Interval = 12500
restartmusic.Enabled = true', 12500
kicker7timer.Interval = 18000
kicker7timer.Enabled = true', 18000
exit sub
end if
if neverd.state = 0 then
neverd.state = 1
playsound "boneorgan1"

startbonemode.Interval = 8000
startbonemode.Enabled = true ', 8000
restartmusic.Interval = 20500
restartmusic.Enabled = true', 20500
kicker7timer.Interval = 21000
kicker7timer.Enabled = true', 21000
'checknever
exit sub
end if
if neveri.state = 0 then
neveri.state = 1

playsound "mouthintruders"

startwatermode.Interval = 12000
startwatermode.Enabled = true ', 12000
restartmusic.Interval = 12500
restartmusic.Enabled = true', 12500
kicker7timer.Interval = 18000
kicker7timer.Enabled = true', 18000
'water slide
'checknever
exit sub
end if
if nevere3.state = 0 then
nevere3.state = 1
playsound "mouthintruders"
startfightmode.Interval = 12000
startfightmode.Enabled = true ', 12000
restartmusic.Interval = 34000'
restartmusic.Enabled = true', 34000
kicker7timer.Interval = 40000
kicker7timer.Enabled = true', 40000
'fight fratellis
'checknever
exit sub
end if
end sub
sub kicker7timer_Timer()
transdisptimer.Enabled = false
'playsound "mouthintruders"
kicker7timer.Enabled = false
kicker7solenoidpulse
if modename <> "bone" then
resetdoor.Interval = 2000
resetdoor.Enabled = true' , 1000
end if

end sub

'-------------------------------------
'Key Mode
'-------------------------------------
sub startkeymode_Timer() :PuPEvent 805


transdisptimer.Enabled = false
light43.state = 0
light51.state = 2
light36.state = 2
light45.state = 2
light23.state = 2
light24.state = 2
light25.state = 2
startkeymode.Enabled = false
modename = "key"

modetime = 61
modetimer.Interval = 8000
modetimer.Enabled = true ', 8000
keysearch = 20
scoreupdate = false
playsound "keyintro"
DispDmd1.Text =  "[xc][y5][f5]HIT SPINNER TO[xc][y18]SEARCH FOR KEY[edge4]"
pupDMDDisplay "default", "HIT SPINNER TO^SEARCH FOR KEY", "", 2, 1, 10

end sub

sub keyfound()
addscore(jackpotscore)
restartmusic.Interval = 2000
restartmusic.Enabled = true ', 2000
playsound "keyfound"
DispDmd1.Text =  "[xc][yc][f5]KEY FOUND![edge4]" & "[f9][xc][yc]" & chr(45)
pupDMDDisplay "default", "KEY FOUND!!", "@KeyFound.mp4", 3, 1, 10

flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true' , 2000
endmode
end sub

'------------------------------------------------------------------
'boulder mode
'----------------------
sub startbouldermode_Timer() :PuPEvent 803


transdisptimer.Enabled false
light43.state = 0
light50.state = 2
light40.state = 2
light37.state = 2
light26.state = 2
light51.state = 2
light36.state = 2
light45.state = 2
light23.state = 2
light24.state = 2
light25.state = 2
startbouldermode.Enabled false
modename = "boulder"
playsound "boulders"
modetime = 45
modetimer.Interval = 8000
modetimer.Enabled true ', 8000
'keysearch = 30
scoreupdate = false
'playsound "i_have_the_key"
DispDmd1.Text =  "[xc][y5][f5]HIT LOOPS TO[xc][y18]DODGE BOULDERS[edge4]"
pupDMDDisplay "default", "HIT LOOPS TO^DODGE BOULDERS", "", 2, 1, 10
end sub

sub boulderjp()
addscore(jackpotscore)
DispDmd1.Text =  "[xc][yc][f5]JACKPOT[edge4]"
pupDMDDisplay "default", "JACKPOT!!", "@jackpot.mp4", 3, 1, 10
playsound "jackpot"
'flushdmdtimer.set true , 2000
modetimer.Interval = 2000
modetimer.Enabled true ', 2000

end sub
'---------------------------------------------
'bone organ mode


sub startbonemode_Timer() :PuPEvent 804

transdisptimer.Enabled false
light43.state = 0
bulb4.state = 2
startbonemode.Enabled false
modename = "bone"
modetime = 45
modetimer.Interval = 14000
modetimer.Enabled = true ', 14000
'keysearch = 30
scoreupdate = false
boneplay = 0
playsound "boneorgan2"
DispDmd1.Text =  "[xc][y5][f5]PLAY THE[xc][y18]BONES[edge4]"
pupDMDDisplay "default", "PLAY THE^BONES", "3", 3, 1, 10
end sub

sub bonehit()
boneplay = boneplay + 1

if boneplay = 1 or boneplay = 2 then
playsound "boneorgan6"
addscore(jackpotscore)
DispDmd1.Text =  "[xc][yc][f5]JACKPOT[edge4]"
pupDMDDisplay "default", "JACKPOT!!", "@jackpot.mp4", 3, 1, 20
'flushdmdtimer.set true , 2000
modetimer.Interval = 2000
modetimer.Enabled = true ', 2000
kicker7timer.Interval = 2000
kicker7timer.Enabled true ', 2000
'resetdoor,set true , 3000
end if

if boneplay = 3 then
playsound "boneorgan3"
addscore(jackpotscore)
DispDmd1.Text =  " SUPER JACKPOT "
pupDMDDisplay "default", "SUPER JACKPOT", "@superjackpot.mp4", 3, 1, 20
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true' , 2000
delayendmode.Interval = 2000
delayendmode.Enabled true ', 2000
kicker7timer.Interval = 2000
kicker7timer.Enabled true ', 2000
resetdoor.Interval = 3000
resetdoor.Enabled true ', 3000
restartmusic.Interval = 3000
restartmusic.Enabled true ', 2000
end if
end sub

'-------------------------------------------------------------
'Water slide mode

sub startwatermode_Timer()

transdisptimer.Enabled false
light43.state = 0
light44.state = 2
light39.state = 2
light49.state = 2

startwatermode.Enabled = false
modename = "water"
modetime = 45
modetimer.Interval = 8000
modetimer.Enabled = true ' 8000
'keysearch = 30
scoreupdate = false
DispDmd1.Text =  "WATER SLIDE"
pupDMDDisplay "default", "WATER SLIDE", "@waterslide.mp4", 3, 1, 10
end sub


sub waterjp()
addscore(jackpotscore)
modejp = true
flushdmdtimer.Interval = 6500
flushdmdtimer.Enabled = true' , 6500
playsound "water_slide"
dmdfont = ""

matchchr = 32
endmatchchr = 91
dmdanimtimer.Enabled = true' , 100
dmdspeed = 70

endtext =  "JACKPOT"
'flushdmdtimer.set true , 2000
modetimer.Interval = 2000
modetimer.Enabled = true ', 2000

end sub

'--------------------------------------------------------------------
'fight fratellis


sub startfightmode_Timer()

transdisptimer.Enabled = false
light43.state = 0
light20.state = 2
light21.state = 2
light22.state = 2

playsound "fratfightstart"
startfightmode.set false
modename = "fight"
modetime = 45
modetimer.Interval = 8000
modetimer.Enabled = true ' 8000
'keysearch = 30
scoreupdate = false
DispDmd1.Text =  "FIGHT THE FRATELLIS"
'pDMDSplashBig "FIGHT THE FRATELLIS", 4, 33023
pupDMDDisplay "default", "FIGHT THE^FRATELLIS", "@FightFratellis.mp4", 3, 1, 10
end sub


sub fighthit()
addscore(jackpotscore)
DispDmd1.Text =  "JACKPOT"
pupDMDDisplay "default", "JACKPOT!!", "@jackpot.mp4", 3, 1, 20
playsound "jackpot"
'flushdmdtimer.set true , 2000
modetimer.Interval = 2000
modetimer.Enabled = true ' 2000

end sub


sub delayendmode_Timer()
delayendmode.set false
endmode
end sub


sub transdisptimer_Timer()
pupDMDDisplay "default", "TRANSLATE^MAP", "@translatemap.mp4", 3, 1, 10
end sub


sub modetimer_Timer()
modetime = modetime - 1
modetimer.Interval = 1000
modetimer.Enabled = true ' 1000
if modetime < 0 then
restartmusic.Interval = 1000
restartmusic.Enabled = true', 500
flushdmdtimer.Interval = 500
flushdmdtimer.Enabled = true' , 500
if modename = "bone" then

kicker7timer.Enabled = true ', 1000
resetdoor.Enabled = true' , 2000
end if

endmode
end if
if modejp = false then
if modename = "key" then
DispDmd1.Text =  "SEARCH FOR KEY" & keysearch & " " & modetime & " " & " " & chr(45)
pupDMDDisplay "default", "SEARCH FOR KEY^"&keysearch&" " & modetime & " " & " ", "3", 2, 1, 10
end if
if modename = "boulder" then
DispDmd1.Text =  "HIT LOOPS" & " " & modetime & " "
pupDMDDisplay "default", "HIT LOOPS^"& modetime, "3", 3, 1, 10
end if


if modename = "bone" then
DispDmd1.Text =  "HIT BONE ORGAN" & " " & modetime & " "
pupDMDDisplay "default", "HIT BONE ORGAN^"& modetime, "3", 2, 1, 10
end if
if modename = "water" then
DispDmd1.Text =  "HIT RAMP" & " " & modetime & " "
pupDMDDisplay "default", "HIT RAMP^"& modetime, "3", 2, 1, 10
end if
if modename = "fight" then
DispDmd1.Text =  "HIT FRATELLIS" & " " & modetime & " "
pupDMDDisplay "default", "HIT FRATELLIS^"& modetime, "3", 2, 1, 10
end if
end if

end sub

sub endmode()
checknever
organhitsleft = 7
if modename = "key" then
modename = ""
modetimer.Enabled = false
light51.state = 1
light36.state = 1
light45.state = 1
light23.state = 1
light24.state = 2
light25.state = 0
'scoreupdate = true
end if
if modename = "boulder" then
modename = ""
modetimer.Enabled = false
light50.state = 2
light40.state = 0
light37.state = 0
light26.state = 0
light51.state = 1
light36.state = 1
light45.state = 1
light23.state = 1
light24.state = 2
light25.state = 0
'scoreupdate = true
end if
if modename = "bone" then
modename = ""
bulb4.state = 0
modetimer.Enabled = false

end if
if modename = "water" then
modename = ""
light49.state = 0
light39.state = 0
light44.state = 0

modetimer.Enabled = false

end if

if modename = "fight" then
modename = ""
light20.state = 0
light21.state = 0
light22.state = 0
playsound "fratfightend"
modetimer.Enabled = false

end if
end sub



'need to add hits left display
sub target10_hit()
PlaySoundAt SoundFXDOF("DTDrop", 119, DOFPulse, DOFDropTargets), target10
'playsound "fx_droptarget"
if modename = "truffle" then
trufflehit
end if

if light43.state = 2 then
exit sub
end if

if modename <> "" then
DOF 119, DOFPulse
target10.Isdropped = 0
end if
if modename = "" then
'flasher4.state = bulbblink
'flasher5.state = bulbblink
'bulb11.state = bulbblink
'bulb19.state = bulbblink
'flasheroff.set true , 500
organhitsleft = organhitsleft - 1
'spellarkframe = spellarkframe + 1
addscore (5000)
if organhitsleft = 0 then
startmap
exit sub
end if

if scoreupdate = true then
scoreupdate = false
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(spellarkframe)
dispdmd1.text = "[f3][xc][y3]BONE ORGAN" & "[xc][y19]" &  organhitsleft & " MORE FOR MAP[edge4]"
pupDMDDisplay "shownum", organhitsleft&"^MORE FOR MAP^BONE ORGAN", "", 2, 0, 10
flushdmdtimer.Interval = 2000
 flushdmdtimer.Enabled = true' , 2000
end if
exit sub
end if
end sub
sub target11_hit()
PlaySoundAt SoundFXDOF("DTDrop", 119, DOFPulse, DOFDropTargets), Target11
'playsound "fx_droptarget"
if modename = "truffle" then
trufflehit
target11.Isdropped = 0
end if
if light43.state = 2 then
exit sub
end if
if modename <> "" then
DOF 119, DOFPulse
target11.Isdropped = 0
end if
if modename = "" then
'flasher4.state = bulbblink
'flasher5.state = bulbblink
'bulb11.state = bulbblink
'bulb19.state = bulbblink
'flasheroff.set true , 500
organhitsleft = organhitsleft - 1
'spellarkframe = spellarkframe + 1
addscore (5000)
if organhitsleft = 0 then
startmap
exit sub
end if

if scoreupdate = true then
scoreupdate = false
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(spellarkframe)
dispdmd1.text = "[f3][xc][y3]BONE ORGAN" & "[xc][y19]" &  organhitsleft & " MORE FOR MAP[edge4]"
pupDMDDisplay "shownum", organhitsleft&"^MORE FOR MAP^BONE ORGAN", "", 2, 0, 10
flushdmdtimer.Interval = 2000
 flushdmdtimer.Enabled = true' , 2000
end if
exit sub
end if
end sub



sub target12_hit()
PlaySoundAt SoundFXDOF("DTDrop", 119, DOFPulse, DOFDropTargets), Target12
'playsound "fx_droptarget"
if modename = "truffle" then
trufflehit
DOF 119, DOFPulse
target12.Isdropped = 0
end if
if light43.state = 2 then
exit sub
end if
if modename <> "" then

target12.Isdropped = 0
end if
if modename = "" then
'flasher4.state = bulbblink
'flasher5.state = bulbblink
'bulb11.state = bulbblink
'bulb19.state = bulbblink
'flasheroff.set true , 500
organhitsleft = organhitsleft - 1
'spellarkframe = spellarkframe + 1
addscore (5000)
if organhitsleft = 0 then
startmap
exit sub
end if

if scoreupdate = true then
scoreupdate = false
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(spellarkframe)
dispdmd1.text = "[f3][xc][y3]BONE ORGAN" & "[xc][y19]" &  organhitsleft & " MORE FOR MAP[edge4]"
pupDMDDisplay "shownum", organhitsleft&"^MORE FOR MAP^BONE ORGAN", "", 2, 0, 10
flushdmdtimer.Interval = 2000
 flushdmdtimer.Enabled = true' , 2000
end if
exit sub
end if
end sub




sub target13_hit()
PlaySoundAt SoundFXDOF("DTDrop", 120, DOFPulse, DOFDropTargets), Target13
'playsound "fx_droptarget"
if modename = "truffle" then
trufflehit
DOF 120, DOFPulse
target13.Isdropped = 0
end if
if light43.state = 2 then
exit sub
end if
if modename <> "" then

target13.Isdropped = 0
end if
if modename = "" then
'flasher4.state = bulbblink
'flasher5.state = bulbblink
'bulb11.state = bulbblink
'bulb19.state = bulbblink
'flasheroff.set true , 500
organhitsleft = organhitsleft - 1
'spellarkframe = spellarkframe + 1
addscore (5000)
if organhitsleft = 0 then
startmap
exit sub
end if

if scoreupdate = true then
scoreupdate = false
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(spellarkframe)
dispdmd1.text = "[f3][xc][y3]BONE ORGAN" & "[xc][y19]" &  organhitsleft & " MORE FOR MAP[edge4]"
pupDMDDisplay "shownum", organhitsleft&"^MORE FOR MAP^BONE ORGAN", "", 2, 0, 10
flushdmdtimer.Interval = 2000
 flushdmdtimer.Enabled = true' , 2000
end if
exit sub
end if
end sub



sub target14_hit()
PlaySoundAt SoundFXDOF("DTDrop", 120, DOFPulse, DOFDropTargets), Target14
'playsound "fx_droptarget"
if modename = "truffle" then
trufflehit
DOF 120, DOFPulse
target14.Isdropped = 0
end if
if light43.state = 2 then
exit sub
end if
if modename <> "" then

target14.Isdropped = 0
end if
if modename = "" then
'flasher4.state = bulbblink
'flasher5.state = bulbblink
'bulb11.state = bulbblink
'bulb19.state = bulbblink
'flasheroff.set true , 500
organhitsleft = organhitsleft - 1
'spellarkframe = spellarkframe + 1
addscore (5000)
if organhitsleft = 0 then
startmap
exit sub
end if

if scoreupdate = true then
scoreupdate = false
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(spellarkframe)
dispdmd1.text = "[f3][xc][y3]BONE ORGAN" & "[xc][y19]" &  organhitsleft & " MORE FOR MAP[edge4]"
pupDMDDisplay "shownum", organhitsleft&"^MORE FOR MAP^BONE ORGAN", "", 2, 0, 10
flushdmdtimer.Interval = 2000
 flushdmdtimer.Enabled = true' , 2000
end if
exit sub
end if
end sub



sub target15_hit()
PlaySoundAt SoundFXDOF("DTDrop", 120, DOFPulse, DOFDropTargets), Target15
'playsound "fx_droptarget"
if modename = "truffle" then
trufflehit
DOF 120, DOFPulse
target15.Isdropped = 0
end if
if light43.state = 2 then
exit sub
end if
if modename <> "" then

target15.Isdropped = 0
end if
if modename = "" then
'flasher4.state = bulbblink
'flasher5.state = bulbblink
'bulb11.state = bulbblink
'bulb19.state = bulbblink
'flasheroff.set true , 500
organhitsleft = organhitsleft - 1
'spellarkframe = spellarkframe + 1
addscore (5000)
if organhitsleft = 0 then
startmap
exit sub
end if

if scoreupdate = true then
scoreupdate = false
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(spellarkframe)
dispdmd1.text = "[f3][xc][y3]BONE ORGAN" & "[xc][y19]" &  organhitsleft & " MORE FOR MAP[edge4]"
pupDMDDisplay "shownum", organhitsleft&"^MORE FOR MAP^BONE ORGAN", "", 2, 0, 10
flushdmdtimer.Interval = 2000
 flushdmdtimer.Enabled = true' , 2000
end if
exit sub
end if

end sub



sub target16_hit()
PlaySoundAt SoundFXDOF("DTDrop", 120, DOFPulse, DOFDropTargets), Target16
'playsound "fx_droptarget"
if modename = "truffle" then
trufflehit
DOF 120, DOFPulse
target16.Isdropped = 0
end if
if light43.state = 2 then
target16.Isdropped = 0
exit sub
end if
if modename <> "" then

target16.Isdropped = 0
end if

if modename = "" then
'flasher4.state = bulbblink
'flasher5.state = bulbblink
'bulb11.state = bulbblink
'bulb19.state = bulbblink
'flasheroff.set true , 500
organhitsleft = organhitsleft - 1
'spellarkframe = spellarkframe + 1
addscore (5000)
if organhitsleft = 0 then
startmap
exit sub
end if

if scoreupdate = true then
scoreupdate = false
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(spellarkframe)
dispdmd1.text = "[f3][xc][y3]BONE ORGAN" & "[xc][y19]" &  organhitsleft & " MORE FOR MAP[edge4]"
pupDMDDisplay "shownum", organhitsleft&"^MORE FOR MAP^BONE ORGAN", "", 2, 0, 10
flushdmdtimer.Interval = 2000
 flushdmdtimer.Enabled = true' , 2000
end if
exit sub
end if

end sub


sub startmap()
light43.state = 2
fakedoor.Isdropped = true
bulb4.state = 2
'door.MoveTo door.Tx, door.Ty-40, door.Tz , 50
DoorDown 60, 0
end sub

Sub Fakedoor_Hit
    Playsound "tack"
End Sub


sub startraiderscene1()
jackpotanim = false
lightseq1.play seqalloff
flushdmdtimer.set false
rjackpothits = 10
rjackpothitsleft = 10
magnet.Isdropped = false
magnet.render = false
'overlay1.fadein()
'Overlay1.Frame 73, 274,73
ospeed = 77
oframestart =73
oframeend = 274

scoreupdate = false
endtext  = "" & score(currentplayer) & " " & lastscore &" " & " " & chr(58) & " " & rjackpothitsleft & " MORE" & "SWITCHES" & "FOR JACKPOT"

'startoverlay()
'ball release
stopmusic 1
sounddelay.set true, 820
tempsound =  "raidersintro"
'light1.state = bulbon
'light2.state = bulbblink
spellarkframe = 38
organhitsleft = 19

delayark.set true,16000
modename = "rscene1"
bMultiBallMode = True
end sub
sub delayark_Timer()
lightseq1.stopplay()
delayark.set false
openark()
'if modename = "rscene1" then
fourballtimer.set true, 500
'end if
delayark2.set true,1000
'overlay1.fadeout()
'overlay1.frame 1,1,1
end sub
sub delayark2_Timer()
delayark2.set false
openark()
restartmusic.Interval = 15000
restartmusic.Enabled = true
end sub
sub fourballtimer_Timer()
fourballtimer.set false
'kicker7.createball
'kicker8.createball
'kicker9.createball
'kicker10.createball
balldropsound.set true ,1500
magnet.Isdropped = true
if modename = "rscene1" then
createnewball()
'plunger.pull
threeballtimer.set true,1500
end if
end sub

sub threeballtimer_Timer()
threeballtimer.set false
plungersolenoidpulse
ballcreatetimer.set true,500
twoballtimer.set true, 1500
end sub
sub twoballtimer_Timer()
twoballtimer.set false
plungersolenoidpulse
ballcreatetimer.set true,500
oneballtimer.set true, 1500
end sub

sub oneballtimer_Timer()
oneballtimer.set false
plungersolenoidpulse
end sub
 sub ballcreatetimer_Timer()
createnewball()
'plunger.pull
ballcreatetimer.set false
end sub

sub startraiderscene2()
jackpotanim = false
lightseq1.play seqalloff
flushdmdtimer.set false
rjackpothits = 10
rjackpothitsleft = 10
magnet.Isdropped = false
magnet.render = false
'overlay1.fadein()
'Overlay1.Frame 73, 274,73
ospeed = 77
oframestart =73
oframeend = 274

scoreupdate = false
endtext  = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]" & rjackpothitsleft & " MORE" & "[f5][x90][y25]TO LIGHT" & "[f5][x90][y45]JACKPOT"

startoverlay()
'ball release
stopmusic 1
sounddelay.set true, 820
tempsound =  "raidersintro"
light2.state = 1
light3.state = 2
spellarkframe = 38
organhitsleft = 19
delayark.set true,16000
modename = "rscene2"
bMultiBallMode = True
restartmusic.Interval = 15000
restartmusic.Enabled = true
end sub
sub startraiderscene3()
jackpotanim = false
lightseq1.play seqalloff
rjackpothits = 3
rjackpothitsleft = 3
flushdmdtimer.set false
magnet.Isdropped = false
magnet.render = false
'overlay1.fadein()
'Overlay1.Frame 73, 274,73
ospeed = 77
oframestart =73
oframeend = 274

scoreupdate = false
endtext = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]" & rjackpothitsleft & " MORE" & "[f5][x90][y25]TO LIGHT" & "[f5][x90][y45]JACKPOT"

startoverlay()
'ball release
stopmusic 1
sounddelay.set true, 820
tempsound =  "raidersintro"
light3.state = 1
light4.state = 2
spellarkframe = 38
organhitsleft = 19
delayark.set true,16000
modename = "rscene3"
bMultiBallMode = True
restartmusic.Interval = 15000
restartmusic.Enabled = true
end sub
sub startraiderscene4()
jackpotanim = false
lightseq1.play seqalloff
rjackpothits = 2
rjackpothitsleft = 2
flushdmdtimer.set false
magnet.Isdropped = false
magnet.render = false
'overlay1.fadein()
'Overlay1.Frame 73, 274,73
ospeed = 77
oframestart =73
oframeend = 274
endtext = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]" & rjackpothitsleft & " MORE" & "[f5][x90][y25]TO LIGHT" & "[f5][x90][y45]JACKPOT"

scoreupdate = false
light36.state = 2
light37.state = 2
'light38.state = bulbblink
light39.state = 2
light50.state = 2
light51.state = 2

startoverlay()
'ball release
stopmusic 1
sounddelay.set true, 820
tempsound =  "raidersintro"
light4.state = 1
light23.state = 2
spellarkframe = 38
organhitsleft = 19
delayark.set true,16000
modename = "rscene4"
bMultiBallMode = True
restartmusic.Interval = 15000
restartmusic.Enabled = true
end sub

sub raidershit()
rjackpothitsleft = rjackpothitsleft - 1
addscore(5000)
if rjackpothitsleft = 0 and modename = "rscene1"then
raiderjackpotgot()
if rjackpothits = 40 or rjackpothits = 60 or rjackpothits = 80 or rjackpothits = 100 then
lightrsuper()
exit sub
end if

if rjackpothits = 10 and modename = "rscene1" then
rjackpothits = 30
else
rjackpothits = rjackpothits + 5
end if

rjackpothitsleft = rjackpothits
exit sub
end if
if rjackpothitsleft = 0 and modename = "rscene2"then
raiderjackpotlight1()
rjackpothits = rjackpothits + 10
rjackpothitsleft = rjackpothits
exit sub
end if
if rjackpothitsleft = 0 and modename = "rscene3"then
lightrsuper3()
exit sub
end if
if rjackpothitsleft = 0 and modename = "rscene4"then
lightjackpot4()
exit sub
end if
if jackpotanim = false and modename = "rscene1" then
'dispdmd1.text = " " & score(currentplayer) & "" & lastscore &"" & "" & chr(58) & "" & rjackpothitsleft & " MORE" & "SWITCHES" & "FOR JACKPOT"
pupDMDDisplay "shownum", rjackpothitsleft &"^MORE SWITCHES^FOR JACKPOT", "", 2, 0, 10
end if
if jackpotanim = false and modename = "rscene2" then
'dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]" & rjackpothitsleft & " MORE" & "[f5][x90][y25]TO LIGHT" & "[f5][x90][y45]JACKPOT"
pupDMDDisplay "shownum", rjackpothitsleft &"^MORE HITS TO^LIGHT JACKPOT", "", 2, 0, 10

end if
if jackpotanim = false and modename = "rscene3" then
'dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]" & rjackpothitsleft & " MORE" & "[f5][x90][y25]TO LIGHT" & "[f5][x90][y45]JACKPOT"
pupDMDDisplay "shownum", rjackpothitsleft &"^MORE HITS TO^LIGHT JACKPOT", "", 2, 0, 10
end if
if jackpotanim = false and modename = "rscene4" then
'dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]" & rjackpothitsleft & " MORE" & "[f5][x90][y25]TO LIGHT" & "[f5][x90][y45]JACKPOT"
pupDMDDisplay "shownum", rjackpothitsleft &"^MORE HITS TO^LIGHT JACKPOT", "", 2, 0, 10
end if
end sub

sub lightrsuper()
'light38.state = bulbblink
modename = "rsuper"
'dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x85][y7]HIT ARK FOR" & "[f5][x95][y25]SUPER" & "[f5][x90][y45]JACKPOT"
pupDMDDisplay "default", "HIT ARK FOR^SUPER JACKPOT", "", 2, 1, 10
end sub

sub getrsuper()

lightseq1.Play Seqblinking, , 20,50
jackpotanim = true
ospeed = 77
oframestart =354
oframeend = 399
startoverlay()
jpscore = raidersjp*5
endtext = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64][f8][x60][yc]" & chr(59) & "[f7][x80][yc]" & formatscore(jpscore)
pupDMDDisplay "default", score(currentplayer) & " " & lastscore & "^" & formatscore(jpscore), "", 3, 1, 10
addscore(jpscore)
resetrjackpot.set true ,5000
sounddelay.set true,3500
tempsound = "superjackpot2"
modename = "rscene1"
rjackpothits = rjackpothits + 5
'light38.state = bulboff
rjackpothitsleft = rjackpothits
end sub
sub raiderjackpotgot()

lightseq1.Play Seqblinking, , 20,50
jackpotanim = true
ospeed = 77
oframestart =275
oframeend = 306
startoverlay()
jpscore = raidersjp
endtext = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64][f8][x60][yc]" & chr(59) & "[f7][x80][yc]" & formatscore(jpscore)
addscore(jpscore)
if modename = "rscene1" then
resetrjackpot.set true ,4000
end if

sounddelay.set true,2500
tempsound = "jackpot2"
end sub

sub resetrjackpot_Timer()
jackpotanim = false
resetrjackpot.set false
if modename = "rsuper" then
dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x85][y7]HIT ARK FOR" & "[f5][x95][y25]SUPER" & "[f5][x90][y45]JACKPOT"
pupDMDDisplay "default", "HIT ARK FOR^SUPER JACKPOT", "", 2, 1, 10
else
dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]" & rjackpothitsleft & " MORE" & "[f5][x90][y25]SWITCHES" & "[f5][x80][y45]FOR JACKPOT"
pupDMDDisplay "shownum", rjackpothitsleft &"^MORE SWITCHES^FOR JACKPOT", "", 2, 0, 10
end if
if modename = "rscene2" or modename = "rscene3" or modename = "rscene4" then
dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]" & rjackpothitsleft & " MORE" & "[f5][x90][y25]TO LIGHT" & "[f5][x90][y45]JACKPOT"
pupDMDDisplay "shownum", rjackpothitsleft &"^MORE HITS TO^LIGHT JACKPOT", "", 2, 0, 10
end if
if modename = "rfinal" then
dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]HIT ARK" & "[f5][x105][y25]FOR" & "[f5][x90][y45]JACKPOT"
pupDMDDisplay "default", "HIT ARK FOR^SUPER JACKPOT", "", 2, 1, 10
end if


end sub

sub raiderjackpotlight1()
'light38.state = bulbblink
modename = "r2jackpot"
dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x85][y7]HIT ARK FOR" & "[f5][x90][y45]JACKPOT"
pupDMDDisplay "default", "HIT ARK FOR^SUPER JACKPOT", "", 2, 1, 10
end sub
'NAILBUSTER
sub getr2jackpot()
lightseq1.Play Seqblinking, , 20,50
jackpotanim = true
ospeed = 77
oframestart =307
oframeend = 334
startoverlay()
jpscore = raidersjp
endtext = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64][f8][x60][yc]" & chr(59) & "[f7][x80][yc]" & formatscore(jpscore)
pupDMDDisplay "shoot", ""& score(currentplayer) & "" & lastscore &"^"&formatscore(jpscore), "", 3, 1, 10
'pupDMDDisplay "default", "DATA GADGET LIT", "@DataGadgetLit.mp4", 3, 1, 10
addscore(jpscore)
modename = "rscene2"
resetrjackpot.set true ,4000
'light38.state = bulboff

sounddelay.set true,2500
tempsound = "jackpot2"

end sub
sub raiderjackpotgot3()
if rjackpothits =7 or rjackpothits =13 or rjackpothits =18 or rjackpothits = 24 then
getrsuper3()
exit sub
end if
rjackpothits = rjackpothits + 1

rjackpothitsleft = rjackpothits
lightseq1.Play Seqblinking, , 20,50
jackpotanim = true
ospeed = 77
oframestart =335
oframeend = 353
startoverlay()
jpscore = raidersjp
endtext = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64][f8][x60][yc]" & chr(59) & "[f7][x80][yc]" & formatscore(jpscore)
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", " "& formatscore(jpscore) , 2, 0
addscore(jpscore)
modename = "rscene3"
resetrjackpot.set true ,4000
'light38.state = bulboff
sounddelay.set true,2500
tempsound = "jackpot2"
end sub
sub getrsuper3()

lightseq1.Play Seqblinking, , 20,50
jackpotanim = true
ospeed = 77
oframestart =354
oframeend = 399
startoverlay()
jpscore = raidersjp*5
endtext = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64][f8][x60][yc]" & chr(59) & "[f7][x80][yc]" & formatscore(jpscore)
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", " "& formatscore(jpscore) , 2, 0
addscore(jpscore)
resetrjackpot.set true ,5000
sounddelay.set true,3500
tempsound = "superjackpot2"
modename = "rscene3"
rjackpothits = rjackpothits + 1
'light38.state = bulboff
rjackpothitsleft = rjackpothits
end sub

sub lightrsuper3()
'light38.state = bulbblink
modename = "rsuper3"
dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x85][y7]HIT ARK FOR" & "[f5][x95][y25]SUPER" & "[f5][x90][y45]JACKPOT"
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", "HIT ARK FOR SUPER JACKPOT" , 2, 0
end sub


sub lightjackpot4()
'light38.state = bulbblink
light36.state = 0
light37.state = 0
light39.state = 0
light50.state = 0
light51.state = 0
modename = "r4jackpot"
dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x85][y7]HIT ARK FOR" & "[f5][x90][y45]JACKPOT"
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", "HIT ARK FOR JACKPOT" , 2, 0
end sub


sub getr4jackpot()
lightseq1.Play Seqblinking, , 20,50
jackpotanim = true

rjackpothits = rjackpothits + 1

rjackpothitsleft = rjackpothits
select case(randomnumber(3))
case 1:
ospeed = 77
oframestart =307
oframeend = 334

case 2:
ospeed = 77
oframestart =335
oframeend = 353
case 3:
ospeed = 77
oframestart =275
oframeend = 306
end select
startoverlay()
jpscore = raidersjp
endtext = "" & score(currentplayer) & "" & lastscore &"" & chr(59) & "" & formatscore(jpscore)
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", " " , 2, 0
addscore(jpscore)
modename = "rscene4"
resetrjackpot.set true ,4500
light36.state = 2
light37.state = 2
'light38.state = bulbblink
light39.state = 2
light50.state = 2
light51.state = 2

sounddelay.set true,2500
tempsound = "jackpot2"

end sub



sub startraidersfinal()
jackpotanim = false
lightseq1.play seqalloff
flushdmdtimer.set false
rjackpothits = 7
rjackpothitsleft = 7
magnet.dropped = false
magnet.render = false
'overlay1.fadein()
'Overlay1.Frame 73, 274,73
ospeed = 77
oframestart =73
oframeend = 274

scoreupdate = false
endtext  = "" & score(currentplayer) & "" & lastscore &"" & "" & chr(58) & "HIT ARK" & "FOR" & "JACKPOT"
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", "HIT ARK FOR JACKPOT" , 2, 0

startoverlay()
'ball release
stopmusic 1
sounddelay.set true, 820
tempsound =  "raidersintro"
light23.state = 1
spellarkframe = 38
organhitsleft = 19
delayark.set true,16000
modename = "rfinal"
bMultiBallMode = True
end sub

sub rfinaljackpot()

lightseq1.Play Seqblinking, , 20,50

jpscore = raidersjp

rjackpothitsleft = rjackpothitsleft - 1
if rjackpothitsleft = 0 then
lightseq1.Play Seqblinking, , 20,50
jackpotanim = true
ospeed = 77
oframestart =354
oframeend = 399
startoverlay()
jpscore = raidersjp*5
endtext = "" & score(currentplayer) & "" & lastscore &"" & chr(59) & "" & formatscore(jpscore)
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", " " , 2, 0
addscore(jpscore)
resetrjackpot.set true ,5000
sounddelay.set true,3500
tempsound = "superjackpot2"
modename = "rfinal"
resetrjackpot.set true ,4500

rjackpothits = rjackpothits + 1
'light38.state = bulboff
rjackpothitsleft = rjackpothits
exit sub
end if
'rjackpothitsleft = rjackpothits
if jackpotanim = false then
jackpotanim = true
select case(randomnumber(3))
case 1:
ospeed = 77
oframestart =307
oframeend = 334

case 2:
ospeed = 77
oframestart =335
oframeend = 353
case 3:
ospeed = 77
oframestart =275
oframeend = 306
end select

startoverlay()
end if

endtext = "" & score(currentplayer) & "" & lastscore &"" & chr(59) & "[f7][x80][yc]" & formatscore(jpscore)
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", " "& formatscore(jpscore) , 2, 0
addscore(jpscore)
modename = "rfinal"
resetrjackpot.set true ,4500

'light38.state = bulbblink


sounddelay.set true,2500
tempsound = "jackpot2"
end sub


sub raidersbonus()
spellarkframe = 38
organhitsleft = 19
jpscore = raidersjp*5
if scoreupdate = true then
dispdmd1.text= "" & score(currentplayer) & "" & lastscore &"" & chr(59) & "" & formatscore(jpscore)
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", " "& formatscore(jpscore) , 2, 0
flushdmdtimer.set true ,1500
end if
playsound "superjackpot2"

addscore(jpscore)
end sub

'*******************************************************
'STONE HIT
'*******************************************************

sub stonetarget_hit()
playsound "ballclick"
if modename <> "" then
exit sub
end if

if modename = "truffle" then
trufflehit
end if
if ballsinlock = 2 and modename <> "" then
exit sub
end if

traphitsleft = traphitsleft - 1
if scoreupdate = true and modename = "" then
scoreupdate = false
dispdmd1.text = " " & traphitsleft & " MORE" & " TO OPEN"
pupDMDDisplay "shownum", traphitsleft & "^ MORE^TO OPEN", "", 2, 1, 10
flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true ',1500
end if
if traphitsleft = 0 then
opencaptive()
traphitsleft = 2
scoreupdate = false
DispDmd1.Text = " TRAP" & " OPEN "
pupDMDDisplay "default", "TRAP^OPEN", "@trapopen.mp4", 3, 1, 10

flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true ',1500
playsound "molarram"
end if
'end if
end sub



'sub Target20_hit()
'If HPos <= 0 Then
'if modename <> "" then
'exit sub
'end if

'if modename = "truffle" then
'trufflehit
'end if
'if ballsinlock = 2 and modename <> "" then
'exit sub
'end if

'traphitsleft = traphitsleft - 1
'if scoreupdate = true and modename = "" then
'scoreupdate = false
'dispdmd1.text = " " & traphitsleft & " MORE" & " TO OPEN"
'flushdmdtimer.Interval = 1500
'flushdmdtimer.Enabled = true ',1500
'end if
'if traphitsleft = 0 then
'opencaptive()
'traphitsleft = 2
'scoreupdate = false
'DispDmd1.Text = "TRAP" & " OPEN "
'flushdmdtimer.Interval = 1500
'flushdmdtimer.Enabled = true ',1500
'playsound "molarram"
'end if
'end if
'end sub

Sub kicker6_hit()
PlaySoundAt "fx_drain", kicker6
kicker6.kick 10, 120
nevere2.state = 1
checknever
DispDmd1.Text = "BALL" & " TRAPPED "
pupDMDDisplay "default", "BALL^TRAPPED", "@trap.mp4", 3, 1, 10
opencaptive()
templeopen = false
ballslocked = ballslocked + 1
PlaySong "mu_end"
if ballslocked = 1 then
playsound "lock1"
restartmusic.Interval = 8000
restartmusic.Enabled = true ', 8000
end if
if ballslocked = 2 then
playsound "lock2"
restartmusic.Interval = 8000
restartmusic.Enabled = true ', 8000
end if

if ballslocked = 3 then
startmultiball
neverr.state = 1
checknever
'vuk1.createball 255,255,159
 'vpmtimer.addtimer 1000, "Vuk1SolenoidPulse '"
Vuk1SolenoidPulse.Interval = 800
Vuk1SolenoidPulse.Enabled = True
else
locktrap.Interval = 2500
locktrap.Enabled = true ', 2500
end if
end sub

Sub kicker6_Unhit()
PlaySound "subway2"
End Sub


sub startmultiball()
DOF 118, DOFPulse
ballslocked = 0
ballsonplayfield = 3
modename = "trapmulti"
'stopmusic 1
PlaySong "mu_end"
delaymulti.Interval = 5000
delaymulti.Enabled = true ', 5000
scoreupdate = false
flushdmdtimer.Interval = 30000
flushdmdtimer.Enabled = true ', 30000
dispdmd1.text = " MULTIBALL "
'pDMDSplashBig "MULTIBALL", 6, 33023
pupDMDDisplay "default", "MULTIBALL", "@multiball.mp4", 3, 1, 10
'animation
playsound "lock3"
end sub

sub delaymulti_Timer()
light39.state = 2
light49.state = 2
light44.state = 2
delaymulti.Enabled = false
restartmusic.Interval = 500
restartmusic.Enabled = true ', 500
popup1.IsDropped = 1
Popup.TransY = -25

PlaySound SoundFXDOF("DiverterOn", 130, DOFPulse, DOFContactors)
bmultiballmode = true
end sub

sub trapjp()

addscore(jackpotscore)
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true' , 2000
dispdmd1.text = "JACKPOT"
pupDMDDisplay "default", "JACKPOT!!", "@jackpot.mp4", 3, 1, 10 :PuPEvent 811
end if
end sub



sub locktrap_Timer()
popup1.IsDropped = 0
Popup.TransY = 0
PlaySound SoundFXDOF("DiverterOff", 130, DOFPulse, DOFContactors)
'vuk1.createball 255,255,159
ballsonplayfield = ballsonplayfield - 1
'vuk1.createball 255,255,159
Vuk1SolenoidPulse.Interval = 800
Vuk1SolenoidPulse.Enabled = True
createnewball
locktrap.Enabled = false
autoball = true
vpmtimer.addtimer 3000, "TiraBolaAutomatico.Enabled = 1 '"
end sub

Sub TiraBolaAutomatico_Timer()
TiraBolaAutomatico.Enabled = 0
DOF 111, DOFPulse
DOF 121, DOFPulse
PlungerIM.AutoFire
End Sub

'sub collectstone()
'stonescollected = stonescollected + 1
'if stonescollected < 3 then
'playsound "preparetomeetcaleem"

'overlay1.fadein()
'overlay1.frame 782
'stonedisplaytimer.set true , 2000
'end if
'end sub

'sub stonedisplaytimer_expired()
'stonedisplaytimer.set false
''overlay1.fadeout()
'scoreupdate = false
'addscore(125000)
'dispdmd1.text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64][f8][x60][yc]" & chr(75) & "[f6][x110][y15]" & stonescollected & " OF 3[f6][x95][y30]COLLECTED"
'flushdmdtimer.set true ,2000
'end sub


sub starttemplescene()
if light24.state = 1 then
scoreupdate = true
templebonus()
exit sub
end if
stopmusic 1
stonescollected = 0
ospeed = 77
oframestart =401
oframeend = 639

scoreupdate = false

'endtext  = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(58) & "[f5][x95][y7]" & rjackpothitsleft & " MORE" & "[f5][x90][y25]SWITCHES" & "[f5][x80][y45]FOR JACKPOT"
if light24.state = 2 then
light24.state = 1
modename = "tfinal"
end if
if light8.state = 2 then
light8.state = 1
light24.state = 2
modename = "tscene4"
end if
if light7.state = 2 then
light7.state = 1
light8.state = 2
modename = "tscene3"
end if
if light6.state = 2 then
light6.state = 1
light7.state = 2
modename = "tscene2"
end if
if light5.state = 2 then
light5.state = 1
light6.state = 2
modename = "tscene1"
endtext  = "" & score(currentplayer) & "" & lastscore &"" & "" & chr(75) & "TEMPLE OF DOOM"  & "SHOOT FLASHING" & "ARROWS"
pDMDSplashLines  "TEMPLE OF DOOM", "SHOOT^FLASHING ARROWS" , 2, 0
light36.state = 2
light37.state = 2
'light38.state = bulbblink
light39.state = 2
light50.state = 2
light51.state = 2
end if


starttemplescenedelay.set true , 21000
startoverlay()
bmultiballmode = true
end sub

sub starttemplescenedelay_Timer()
starttemplescenedelay.set false
kicker1.solenoidpulse()
playsound SoundFX("fx_kicker", DOFContactors)
restartmusic.set true, 500
createnewball()
'plunger.pull
threeballtimer.set true,1500
end sub


sub templebonus()
jpscore = templejp*5
if scoreupdate = true then
dispdmd1.text= "" & score(currentplayer) & "" & lastscore &"" & chr(59) & "" & formatscore(jpscore)
pDMDSplashLines " "& score(currentplayer) & " " & lastscore &"" & "", " " & formatscore(jpscore) , 2, 0

flushdmdtimer.set true ,1500
end if
playsound "superjackpot2"

addscore(jpscore)
end sub

'*******************************************************
'grail hit
'*******************************************************

sub kicker1_hit()
If Tilted Then
Kicker1Fake.enabled = 0
kicker1timer2.Interval = 500
kicker1timer2.Enabled =  true ', 1800
Else
Kicker1Fake.enabled = 0
PlaySoundAt "scoopenter", kicker1
if light41.state = 2  and modename = "" then
pDMDSplash3Lines "MYSTERY", "", "START", 5, 0
startmystery()
light41.state = 0
exit sub
end if
kicker1timer.Interval = 1500
kicker1timer.Enabled = true ', 1500
'grailhit = grailhit + 1
addscore(10000)
kicker1timer2.Interval = 1800
kicker1timer2.Enabled =  true ', 1800
end if
end sub

sub kicker1timer_Timer()
kicker1timer2.Enabled =  false
kicker1solenoidpulse()
playsound SoundFX("fx_kicker",DOFContactors)
kicker1timer.Enabled =  false
end sub

sub kicker1timer2_Timer()
kicker1timer2.Enabled =  false
kicker1solenoidpulse()
playsound SoundFX("fx_kicker",DOFContactors)
kicker1timer.Enabled =  false
end sub
sub resettempledisplay_Timer()
'resettempledisplay.Enabled =  false
end sub



Sub kicker1solenoidpulse()
    PlaySound SoundFX("salidadebola",DOFContactors)
    Kicker1.Kick 0, 30, 1.5
  DOF 114, DOFPulse
  DOF 121, DOFPulse
    Kicker1FakeTimer.interval = 800
    Kicker1FakeTimer.enabled = True
End Sub

Sub Kicker1FakeTimer_Timer
    Kicker1FakeTimer.enabled = False
    Kicker1Fake.enabled = 1
End Sub

'*******************************************************
'skull ramp
'*******************************************************
sub trigger7_hit()
DOF 117, DOFPulse
'playsound "fx_PlasticRamp"
if modename = "wizard" then
wizardjp
exit sub
end if
if modename = "water" then
waterjp
exit sub
end if
if modename = "trapmulti" then
trapjp
exit sub
end if

if comboon = false then
comboon = true
combotimer.Interval = 8000
combotimer.Enabled = true ', 8000
else
addcombo
combotimer.Interval = 8000
combotimer.Enabled = true ', 8000
end if

if light49.state = 2 then
addscore(jackpotscore)
playsound "chunkamazing"
light39.state = 0
light49.state = 0
light44.state = 0
end if
addscore(5000)
rampshots = rampshots + 1
if rampshots = 10 then
light1.state = 2
end if

if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true ', 1500
if rampshots = 1 then
dispdmd1.text = "[edge4][x20][y5][f6]" & rampshots & " RAMP" & "[f9][xc][yc]" & chr(33) & "[f1][x15][y19]9 MORE TO[x2][y25]LIGHT EXTRA BALL"
pupDMDDisplay "shownum", rampshots & "RAMP MORE TO^LIGHT EXTRA BALL", "", 3, 0, 10
exit sub
end if
if rampshots > 1 and rampshots < 10 then

dispdmd1.text = "[edge4][x20][y5][f6]" & rampshots & " RAMPS" & "[f9][xc][yc]" & chr(33) & "[f1][x15][y19]" & (10 - rampshots) & " MORE TO[x2][y25]LIGHT EXTRA BALL"
pDMDSplashLines " " & rampshots & " RAMPS"& (10 - rampshots)&" MORE TO", "LIGHT EXTRA BALL", 2, 0
exit sub
end if
if rampshots = 10 then
dispdmd1.text = "[edge4][x20][y5][f6]" & rampshots & " RAMP" & "[f9][xc][yc]" & chr(33) & "[f1][x18][y19]GET THE[x15][y25]EXTRA BALL"
pupDMDDisplay "shownum", rampshots & "RAMP MORE TO^LIGHT EXTRA BALL", "", 3, 0, 10
light1.state = 2
exit sub
end if

dispdmd1.text = "[edge4][x20][y5][f6]" & rampshots & " RAMP" & "[f9][xc][yc]" & chr(33) & "[f1][x20][y23]10,000"
pDMDSplashBig " " & rampshots & " RAMP ", 2, 0

addscore(5000)

'if rampshots = rampshotsleft then
'startskullscene()
'end if

end if



'if hitskull = true then
'skullhit = skullhit +1
'resetskull.set true,1500
addscore(20000)
'else
'exit sub
'end if


'hitskull = false

end sub

Sub Trigger002_Hit
   StopSound "fx_PlasticRamp"
End Sub

Sub Trigger001_Hit
  StopSound "fx_PlasticRamp"
  Playsound "fx_RampPlasticHit2"
End Sub

sub startskullscene()
rampshots = 0
rampshotsleft = rampshotsleft + 1
end sub

sub resetskull_Timer()
resetskull.set false
hitskull = true
end sub



'*******************************************************
'DATA Targets
'*******************************************************

sub target6_hit()
DOF 128, DOFPulse
Select case (RandomNumber(2))
case 1: PlaySoundAt "boing4", Target6
case 2: PlaySoundAt "boing6", Target6
end select
if modename = "truffle" then
trufflehit
exit sub
end if

if modename <> "" then
exit sub
end if
'if modename = "rscene1" or modename = "rscene2" or modename = "rscene3" then
'raidershit()
'exit sub
'end if
'if light27.state = bulbon and indyscoring = true then
'indyhit()
'exit sub
'end if
'if light28.state = bulbon and indyscoring = true then
'indyhit2()
'exit sub
'end if
'if light29.state = bulbon and indyscoring = true then
'indyhit3()
'exit sub
'end if

if light27.state = 2 then
light27.state = 1
addscore(5000)

'if scoreupdate = true then
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(60) & indyi & indyn & indyd & indyy & "[f1][x65][y57]COMPLETE FOR INDY SCORING"
'scoreupdate = false
'flushdmdtimer.set true ,2000
'end if
dataD = "D"' & chr(46)
checkdata
else
addscore(1000)
end if

end sub

sub target7_hit()
DOF 128, DOFPulse
Select case (RandomNumber(2))
case 1: PlaySoundAt "boing4", Target7
case 2: PlaySoundAt "boing6", Target7
end select

if modename = "truffle" then
trufflehit
exit sub
end if

if modename <> "" then
exit sub
end if
'if modename = "rscene1" or modename = "rscene2" or modename = "rscene3" then
'raidershit()
'exit sub
'end if
'if light28.state = bulbon and indyscoring = true then
'indyhit()
'exit sub
'end if
'if light29.state = bulbon and indyscoring = true then
'indyhit2()
'exit sub
'end if
'if light30.state = bulbon and indyscoring = true then
'indyhit3()
'exit sub
'end if


if light28.state = 2 then
light28.state = 1
'indyN = "[f8][yc][x60]" & chr(62)

dataA = "A"' & chr(47)
addscore(5000)

'if scoreupdate = true then
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(60) & indyi & indyn & indyd & indyy & "[f1][x65][y57]COMPLETE FOR INDY SCORING"
'scoreupdate = false
'flushdmdtimer.set true ,2000
'end if
checkdata
else
addscore(1000)
end if
end sub

sub target8_hit()
DOF 128, DOFPulse
Select case (RandomNumber(2))
case 1: PlaySoundAt "boing4", Target8
case 2: PlaySoundAt "boing6", Target8
end select
if modename = "truffle" then
trufflehit
exit sub
end if

if modename <> "" then
exit sub
end if
'if modename = "rscene1" or modename = "rscene2" or modename = "rscene3" then
'raidershit()
'exit sub
'end if
'if light29.state = bulbon and indyscoring = true then
'indyhit()
'exit sub
'end if

'if light30.state = bulbon and indyscoring = true then
'indyhit2()
'exit sub
'end if
'if light27.state = bulbon and indyscoring = true then
'indyhit3()
'exit sub
'end if

if light29.state = 2 then
light29.state = 1
'indyD = "[f8][yc][x60]" & chr(63)
addscore(5000)

'if scoreupdate = true then
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(60) & indyi & indyn & indyd & indyy & "[f1][x65][y57]COMPLETE FOR INDY SCORING"
'scoreupdate = false
'flushdmdtimer.set true ,2000
'end if
dataT = "T"' & chr(48)
checkdata
else
addscore(1000)
end if
end sub

sub target9_hit()
DOF 128, DOFPulse
Select case (RandomNumber(2))
case 1: PlaySoundAt "boing4", Target9
case 2: PlaySoundAt "boing6", Target9
end select
if modename = "truffle" then
trufflehit
exit sub
end if

if modename <> "" then
exit sub
end if
'if modename = "rscene1" or modename = "rscene2" or modename = "rscene3" then
'raidershit()
'exit sub
'end if
'if light30.state = bulbon and indyscoring = true then
'indyhit()
'exit sub
'end if
'if light27.state = bulbon and indyscoring = true then
'indyhit2()
'exit sub
'end if
'if light28.state = bulbon and indyscoring = true then
'indyhit3()
'exit sub
'end if
if light30.state = 2 then
light30.state = 1
'indyY = "[f8][yc][x60]" & chr(64)
addscore(5000)
'if scoreupdate = true then
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(60) & indyi & indyn & indyd & indyy & "[f1][x65][y57]COMPLETE FOR INDY SCORING"
'scoreupdate = false
'flushdmdtimer.set true ,2000
'end if
dataA2 = "A"' & chr(49)
checkdata
else
addscore(1000)
end if
end sub

'********************************************************************************
'Sloth 2x scoring
'*******************************************************************************

sub target1_hit()
DOF 127, DOFPulse
Select case (RandomNumber(4))
case 1: PlaySoundAt "sloth_laugh1", Target1
case 2: PlaySoundAt "sloth_laugh2", Target1
case 3: PlaySoundAt "sloth_laugh3", Target1
case 4: PlaySoundAt "slothsloth", Target1
end select
playsound "fx14"
if modename = "truffle" then
trufflehit
exit sub
end if

'if modename <> "" then
'exit sub
'end if
'if modename = "rscene1" or modename = "rscene2" or modename = "rscene3" then
'raidershit()
'exit sub
'end if
'if light31.state = bulbon and jonesscoring = true then
'joneshit()
'exit sub
'end if
'if light32.state = bulbon and jonesscoring = true then
'joneshit2()
'exit sub
'end if
'if light33.state = bulbon and jonesscoring = true then
'joneshit3()
'exit sub
'end if



if light31.state = 2 then
light31.state = 1
'JonesJ = "[f8][yc][x60]" & chr(66)

slothS = "S"' & chr(50)
addscore(5000)
'if scoreupdate = true then
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(65) & jonesJ & JonesO & jonesN & JonesE & JonesS & "[f1][x65][y57]COMPLETE FOR JONES SCORING"
'scoreupdate = false
'flushdmdtimer.set true ,2000
'end if
checksloth()
else
addscore(1000)
end if
end sub

sub target2_hit()
DOF 127, DOFPulse
Select case (RandomNumber(4))
case 1: PlaySoundAt "sloth_laugh1", Target2
case 2: PlaySoundAt "sloth_laugh2", Target2
case 3: PlaySoundAt "sloth_laugh3", Target2
case 4: PlaySoundAt "slothsloth", Target2
end select
playsound "fx14"
if modename = "truffle" then
trufflehit
exit sub
end if

'if modename <> "" then
'exit sub
'end if
'if modename = "rscene1" or modename = "rscene2" or modename = "rscene3" then
'raidershit()
'exit sub
'end if
'if light32.state = bulbon and jonesscoring = true then
'joneshit()
'exit sub
'end if
'if light33.state = bulbon and jonesscoring = true then
'joneshit2()
'exit sub
'end if
'if light34.state = bulbon and jonesscoring = true then
'joneshit3()
'exit sub
'end if

if light32.state = 2 then
light32.state = 1
'JonesO = "[f8][yc][x60]" & chr(67)
addscore(5000)
'if scoreupdate = true then
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(65) & jonesJ & JonesO & jonesN & JonesE & JonesS & "[f1][x65][y57]COMPLETE FOR JONES SCORING"
'scoreupdate = false
'flushdmdtimer.set true ,2000
'end if
slothL = "L"' & chr(51)
checksloth()
else
addscore(1000)
end if
end sub


sub target3_hit()
DOF 127, DOFPulse
Select case (RandomNumber(4))
case 1: PlaySoundAt "sloth_laugh1", Target3
case 2: PlaySoundAt "sloth_laugh2", Target3
case 3: PlaySoundAt "sloth_laugh3", Target3
case 4: PlaySoundAt "slothsloth", Target3
end select


playsound "fx14"
if modename = "truffle" then
trufflehit
exit sub
end if

'if modename <> "" then
'exit sub
'end if
'if modename = "rscene1" or modename = "rscene2" or modename = "rscene3" then
'raidershit()
'exit sub
'end if
'if light33.state = bulbon and jonesscoring = true then
'joneshit()
'exit sub
'end if
'if light34.state = bulbon and jonesscoring = true then
'joneshit2()
'exit sub
'end if
'if light35.state = bulbon and jonesscoring = true then
'oneshit3()
'exit sub
'end if

if light33.state = 2 then
light33.state = 1
'JonesN = "[f8][yc][x60]" & chr(68)
addscore(5000)
'if scoreupdate = true then
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(65) & jonesJ & JonesO & jonesN & JonesE & JonesS & "[f1][x65][y57]COMPLETE FOR JONES SCORING"
'scoreupdate = false
'flushdmdtimer.set true ,2000
'end if
slothO = "O"' & chr(52)
checksloth()

else
addscore(1000)
end if
end sub


sub target4_hit()
DOF 127, DOFPulse
Select case (RandomNumber(4))
case 1: PlaySoundAt "sloth_laugh1", Target4
case 2: PlaySoundAt "sloth_laugh2", Target4
case 3: PlaySoundAt "sloth_laugh3", Target4
case 4: PlaySoundAt "slothsloth", Target4
end select

playsound "fx14"
if modename = "truffle" then
trufflehit
exit sub
end if

'if modename <> "" then
'exit sub
'end if
'if modename = "rscene1" or modename = "rscene2" or modename = "rscene3" then
'raidershit()
'exit sub
'end if

'if light34.state = bulbon and jonesscoring = true then
'joneshit()
'exit sub
'end if
'if light35.state = bulbon and jonesscoring = true then
'joneshit2()
'exit sub
'end if
'if light31.state = bulbon and jonesscoring = true then
'joneshit3()
'exit sub
'end if

if light34.state = 2 then
light34.state = 1
'JonesE = "[f8][yc][x60]" & chr(69)
addscore(5000)
'if scoreupdate = true then
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(65) & jonesJ & JonesO & jonesN & JonesE & JonesS & "[f1][x65][y57]COMPLETE FOR JONES SCORING"
'scoreupdate = false
'flushdmdtimer.set true ,2000
'end if
slothT = "T"' & chr(53)
checksloth()
else
addscore(1000)
end if

end sub


sub target5_hit()
DOF 127, DOFPulse
Select case (RandomNumber(4))
case 1: PlaySoundAt "sloth_laugh1", Target5
case 2: PlaySoundAt "sloth_laugh2", Target5
case 3: PlaySoundAt "sloth_laugh3", Target5
case 4: PlaySoundAt "slothsloth", Target5
end select

playsound "fx14"
if modename = "truffle" then
trufflehit
exit sub
end if

'if modename <> "" then
'exit sub
'end if
'if modename = "rscene1" or modename = "rscene2" or modename = "rscene3" then
'raidershit()
'exit sub
'end if
'if light35.state = bulbon and jonesscoring = true then
'joneshit()
'exit sub
'end if
'if light31.state = bulbon and jonesscoring = true then
'joneshit2()
'exit sub
'end if
'if light32.state = bulbon and jonesscoring = true then
'joneshit3()
'exit sub
'end if
if light35.state = 2 then
light35.state = 1
'JonesS = "[f8][yc][x60]" & chr(70)
addscore(5000)
'if scoreupdate = true then
'DispDmd1.Text = "[f1][x0] [y0]" & nvscore(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f8][x60][yc]" & chr(65) & jonesJ & JonesO & jonesN & JonesE & JonesS & "[f1][x65][y57]COMPLETE FOR JONES SCORING"
'scoreupdate = false
'flushdmdtimer.set true ,2000
'end if
slothH = "H"' & chr(54)
checksloth()

else
addscore(1000)
end if
end sub


sub checkdata()
DIM AlphaState
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true ', 1500
dispdmd1.text = dataD & dataA & DataT & dataA2
'pDMDSplashBig " " & dataD & dataA & DataT & dataA2, 3, 33023
pupDMDDisplay "target", "DATA^1111", "datagadget.mp4", 3, 0, 12
end if

if light27.state = 1 and light28.state = 1 and light29.state = 1 and light30.state = 1 then
startindy()
exit sub

end if

if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true ', 1500
dispdmd1.text = dataD & dataA & DataT & dataA2
if light27.state=1 then AlphaState=AlphaState&"1" else AlphaState=AlphaState&"0"
if light28.state=1 then AlphaState=AlphaState&"1" else AlphaState=AlphaState&"0"
if light29.state=1 then AlphaState=AlphaState&"1" else AlphaState=AlphaState&"0"
if light30.state=1 then AlphaState=AlphaState&"1" else AlphaState=AlphaState&"0"
pupDMDDisplay "target", "DATA^"&AlphaState, "datagadget.mp4", 3, 0, 12
end if
end sub



sub checksloth()
DIM AlphaState

if light31.state = 1 and light32.state = 1 and light33.state = 1 and light34.state = 1 and light35.state = 1 then
pupDMDDisplay "target", "SLOTH^11111", "sloth.mp4", 3, 0, 12
startsloth()
end if


if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true ', 1500
dispdmd1.text = slothS & slothL & slothO & slothT & slothH

if light31.state=1 then AlphaState=AlphaState&"1" else AlphaState=AlphaState&"0"
if light32.state=1 then AlphaState=AlphaState&"1" else AlphaState=AlphaState&"0"
if light33.state=1 then AlphaState=AlphaState&"1" else AlphaState=AlphaState&"0"
if light34.state=1 then AlphaState=AlphaState&"1" else AlphaState=AlphaState&"0"
if light35.state=1 then AlphaState=AlphaState&"1" else AlphaState=AlphaState&"0"


'pDMDSplashBig " " & slothS & slothL & slothO & slothT & slothH, 2, 33023
pupDMDDisplay "target", "SLOTH^"&AlphaState, "sloth.mp4", 3, 0, 12

end if
end sub


sub startindy()
playsound "unknownnotneeded10"
light41.state = 2
if modename = "" then
restartmusic.Enabled = true ', 100
end if
scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true ', 2000

lightseq1.Play Seqblinking, , 5,50
DispDmd1.Text = "[x2][y1][f5]DATA GADGET" & "[xc][y15][f6]LIT" & "[f9][xc][yc]"' & chr(36)
pDMDSplashBig "DATA GADGET LIT", 2, 33023
light54.state = 0
light27.state = 1
light28.state = 1
light29.state = 1
light30.state = 1
end sub


sub indylightstimer_Timer()
if light30.state = 1 then
light30.state = 0
light27.state = 1
exit sub
end if
if light29.state = 1 then
light30.state = 1
light29.state = 0
exit sub
end if
if light28.state = 1 then
light29.state = 1
light28.state = 0
exit sub
end if
if light27.state = 1 then
light28.state = 1
light27.state = 0
exit sub
end if
end sub



sub indyscoringtimer_Timer()

'flushdmdtimer.set true , 1500

' show number of indy hits left.
indyhitsleft = indyhitsleft + 1
if modename = "" then
DispDmd1.Text = " " & score(currentplayer) & " " & lastscore &" " & " " & chr(60) &" " & chr(61) &" " & chr(62) &" " & chr(63) &" " & chr(64) & "SUPER SCORE #" & indyhitsleft & " OF 6"
pDMDSplashLines " "& score(currentplayer) & "" & lastscore &"" & "", "SUPER SCORE #" & indyhitsleft & " OF 6", 2, 0
end if
indyscoringtimer.Enabled = false
if indyhitsleft = 6 then
endindy()
if modename = "" then
DispDmd1.Text = "" & score(currentplayer) & "" & lastscore &"" & "" & chr(60) &"" & chr(61) &"" & chr(62) &"" & chr(63) &"" & chr(64) & "BONUS " & formatscore(2000000)
pDMDSplashLines "BONUS", " " & formatscore(2000000), 3, 0
end if
addscore(2000000)
end if

end sub

sub endindy()
'indyscoringtimer.set false
indylightstimer.Enabled = false

indyscoring = false
light27.state = 2
light28.state = 2
light29.state = 2
light30.state = 2
indyI = "[f1] "
indyN = "[f1] "
indyD = "[f1] "
indyY = "[f1] "
indyhitsleft = 0
end sub

sub indyhit()

lightseq1.Play Seqblinking, , 2,50
addscore(750000)
if modename = "" then
DispDmd1.Text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f7][x90][yc]" & formatscore(750000)
pDMDSplashLines  " ", " " & formatscore(750000), 2, 0
scoreupdate = false

indyscoringtimer.Interval = 2000
indyscoringtimer.Enabled = true', 2000
flushdmdtimer.Interval = 3500
flushdmdtimer.Enabled = true ', 3500
end if
end sub
sub indyhit2()

lightseq1.Play Seqblinking, , 2,50
addscore(650000)
if modename = "" then
DispDmd1.Text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f7][x90][yc]" & formatscore(750000)
pDMDSplashLines  " ", " " & formatscore(750000), 2, 0
scoreupdate = false

indyscoringtimer.Interval = 2000
indyscoringtimer.Enabled = true', 2000
flushdmdtimer.Interval = 3500
flushdmdtimer.Enabled = true' , 3500
end if
end sub
sub indyhit3()

lightseq1.Play Seqblinking, , 2,50
addscore(500000)
if modename = "" then
DispDmd1.Text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f7][x90][yc]" & formatscore(750000)
pDMDSplashLines  " ", " " & formatscore(750000), 2, 0
scoreupdate = false

indyscoringtimer.Interval = 2000
indyscoringtimer.Enabled = true', 2000
flushdmdtimer.Interval = 3500
flushdmdtimer.Enabled = true ', 3500
end if
end sub




sub startsloth()
nevern.state = 1:NeverNP.visible = 1
checknever
playsound "slothhey"
slothmulti = 2
flushdmdtimer.Interval = 30000
flushdmdtimer.Enabled = true ', 30000
if modename = "" then
'modename = "sloth"
PlaySong "mu_end"
restartmusic.Enabled = true' , 100
DispDmd1.Text = "[f4][x10][y0]SLOTH" & "[f4][y18][x0]2X SCORING" & "[f9][xc][yc]" & chr(42)
pDMDSplashBonus "SLOTH", "X2 SCORING", 2, 33023
scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true ', 2000
end if
'jonesscoring = true
light38.state = 2
light31.state = 1
light32.state = 1
light33.state = 1
light34.state = 1
light35.state = 1
'joneslightstimer.set true,500
'indyscoringtimer.set true,20000
end sub

sub slothtimer_Timer()
endsloth
playsound "slothrockyroad"
if modename = "" then
PlaySong "mu_end"
restartmusic.Enabled = true ', 50
end if
end sub





sub endsloth()
'indyscoringtimer.set false


slothS = "S"
slothL = "L"
slothO = "O"
slothT = "T"
slothH = "H"

slothtimer.Enabled = false
slothmulti = 2 = false
light31.state = 2
light32.state = 2
light33.state = 2
light34.state = 2
light35.state = 2
light38.state = 0

end sub

sub joneshit()

lightseq1.Play Seqblinking, , 2,50
addscore(750000)
if modename = "" then
DispDmd1.Text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f7][x90][yc]" & formatscore(750000)
pDMDSplashLines " ", " " & formatscore(750000), 2, 0
scoreupdate = false

'jonesscoringtimer.Enabled = true', 2000
flushdmdtimer.Interval = 3500
flushdmdtimer.Enabled = true ', 3500
end if

end sub
sub joneshit2()

lightseq1.Play Seqblinking, , 2,50
addscore(650000)
if modename = "" then
DispDmd1.Text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f7][x90][yc]" & formatscore(750000)
pDMDSplashLines " ", " " & formatscore(750000), 2, 0
scoreupdate = false

'jonesscoringtimer.Enabled = true, 2000
flushdmdtimer.Interval = 3500
flushdmdtimer.Enabled = true ', 3500
end if
end sub
sub joneshit3()

lightseq1.Play Seqblinking, , 2,50
addscore(500000)
if modename = "" then
DispDmd1.Text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f7][x90][yc]" & formatscore(750000)
pDMDSplashLines " ", " " & formatscore(750000), 2, 0
scoreupdate = false

'jonesscoringtimer.Enabled = true, 2000
flushdmdtimer.Interval = 3500
flushdmdtimer.Enabled = true ', 3500
end if
end sub


'**********************************************************************************
'Bumpers
'**********************************************************************************
sub bumper2_hit() :PuPEvent 812 ' raiders
RandomSoundBumperTop Bumper2
DOF 204, DOFPulse
If Tilted Then Exit Sub
'playsound "gunshot6"
if bumper2l.state = 2 then
addscore(25000)
else
addscore(1000)
end if
if modename = "rscene1" or modename = "rscene2" then
raidershit()
exit sub
end if
resetbumpers.Interval = 10000
resetbumpers.Enabled = true', 10000



end sub

sub bumper1_hit() :PuPEvent 808' temple
RandomSoundBumperMiddle Bumper1
DOF 205, DOFPulse
If Tilted Then Exit Sub
'playsound "gunshot6"
if bumper1l.state = 2 then
addscore(25000)
else
addscore(1000)
end if
if modename = "rscene1" or modename = "rscene2" then
raidershit()
exit sub
end if
resetbumpers.Interval = 10000
resetbumpers.Enabled = true', 10000


end sub

sub bumper4_hit():PuPEvent 812 'crusade
RandomSoundBumperBottom Bumper4
DOF 203, DOFPulse
If Tilted Then Exit Sub
'playsound "gunshot6"
if bumper4l.state = 2 then
addscore(25000)
else
addscore(1000)
end if
if modename = "rscene1" or modename = "rscene2" then
raidershit()
exit sub
end if
resetbumpers.Interval = 10000
resetbumpers.Enabled = true', 10000


resetbumperscore.Enabled = true',200


end sub

sub bumper3_hit() :PuPEvent 808 'crystal
RandomSoundBumperTop Bumper3
DOF 202, DOFPulse
If Tilted Then Exit Sub
'playsound "gunshot6"
if bumper3l.state = 2 then
addscore(25000)
else
addscore(1000)
end if
if modename = "rscene1" or modename = "rscene2" then
raidershit()
exit sub
end if
resetbumpers.Interval = 10000
resetbumpers.Enabled = true', 10000



resetbumperscore.Enabled = true',200

end sub

sub resetbumpers_Timer()
resetbumpers.Enabled = false
bumper1l.state = 0
bumper2l.state = 0
bumper3l.state = 0
bumper4l.state = 0
end sub
sub resetbumperscore_Timer()
resetbumperscore.Enabled = false
end sub
'*********************************************************************************
'map room
'*********************************************************************************
sub target17_hit()
PlaySoundAt SoundFX("fx3", DOFContactors), Target17
If Tilted Then Exit Sub
if modename = "truffle" then
trufflehit
end if
'if modename = "rscene1" or modename = "rscene2" then
'raidershit()
'exit sub
'end if
if light20.state = 0 or light20.state = 2 then
light20.state = 1
addscore(5000)
checkfrat()
else
addscore(1000)
end if

end sub
sub target18_hit()
PlaySoundAt SoundFX("fx3", DOFContactors), Target17
If Tilted Then Exit Sub
if modename = "truffle" then
trufflehit
end if
'if modename = "rscene1" or modename = "rscene2" then
'raidershit()
'exit sub
'end if
if light21.state = 0 or light21.state = 2 then
light21.state = 1
addscore(5000)
checkfrat()
else
addscore(1000)
end if

end sub
sub target19_hit()
PlaySoundAt SoundFX("fx3", DOFContactors), Target17
If Tilted Then Exit Sub
if modename = "truffle" then
trufflehit
end if
'if modename = "rscene1" or modename = "rscene2" then
'raidershit()
'exit sub
'end if
if light22.state = 0 or light22.state = 2 then
light22.state = 1
addscore(5000)
checkfrat()
else
addscore(1000)
end if

end sub
sub checkfrat()
if light20.state = 1 and light21.state = 1 and light22.state = 1 then

if modename = "fight" then

light20.state = 2
light21.state = 2
light22.state = 2
fighthit
exit sub
end if

jackpotscore = jackpotscore + 250000

Select case (RandomNumber(3))
case 1: playsound "frat1"
case 2:  playsound "frat2"
case 3:  playsound "frat3"
end select

if jackpotscore > 10000000 then
jackpotscore = 10000000
end if
lightseq1.Play Seqblinking, , 3,50
if scoreupdate = true and jackpotscore < 10000000 then
DispDmd1.Text = "[edge4][f5][xc][y3]JACKPOT GROWS" & "[f4][xc][y18]" & formatscore(jackpotscore)
pDMDSplashLines "JACKPOT GROWS", " " & formatscore(jackpotscore), 2, 0

scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true ', 2000
end if
if jackpotscore = 10000000 and scoreupdate = true then

DispDmd1.Text = "[edge4][f5][xc][y3]JACKPOT AT MAX" & "[f4][xc][y18]" & formatscore(jackpotscore)
pDMDSplashLines "JACKPOT AT MAX", " " & formatscore(jackpotscore), 2, 0

scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true ', 2000
playsound "jingle15"
end if
'light20.flashforms 2000, , bulboff
FlashForMs light20, 2000, 50, 0
'light21.flashforms 2000, , bulboff
FlashForMs light21, 2000, 50, 0
'light22.flashforms 2000, , bulboff
FlashForMs light22, 2000, 50, 0
addscore(10000)
end if
 end sub

sub kicker2_hit()
'if flasher1.state = bulboff then

'end if

if modename = "sword" then
swordsmanhit()
exit sub
end if



addscore(10000)
kicker2.Kick 215, 15
playsound SoundFXDOF("fx_kicker", 115, DOFPulse, DOFContactors)
end sub


sub resetxmarks_Timer()
resetxmarks.set false
xmarks = 1
xmarkshits = 3
'flasher1.state = bulboff
'flasher1b.state = bulboff

end sub



'*******************************************************************************
'loops
'*******************************************************************************
sub trigger3_hit() :PuPEvent 807
DOF 211, DOFPulse
If Tilted Then Exit Sub
loopd = ""
playsound "fx_sensor"
addscore(1000)
if modename = "rscene1" or modename = "rscene2" then
raidershit()
exit sub
end if
end sub

sub trigger4_hit() :PuPEvent 807
If Tilted Then Exit Sub
if loopd = "r" then
'right loop hit "MIKEY"

if comboon = false then
comboon = true
combotimer.Interval = 8000
combotimer.Enabled = true' , 8000
else
addcombo
combotimer.Interval = 8000
combotimer.Enabled = true' , 8000
end if
if modename = "wizard" then
wizardjp
exit sub
end if
if modename = "boulder" then
boulderjp
exit sub
end if

addmarble
end if


addscore(1000)
loopd = "l"
end sub

sub trigger2_hit() :PuPEvent 812
playsound "fx_sensor"
if loopd = "l" then
if comboon = false then
comboon = true
combotimer.Interval = 8000
combotimer.Enabled = true' , 8000
else
addcombo
combotimer.Interval = 8000
combotimer.Enabled = true' , 8000
end if
if modename = "wizard" then
wizardjp
exit sub
end if
if modename = "boulder" then
boulderjp
exit sub
end if
addchunk
end if

addscore(1000)
loopd = "r"
end sub

sub trigger1_hit()
DOF 210, DOFPulse
loopd = ""
if modename = "rscene1" or modename = "rscene2" then
raidershit()
exit sub
end if
addscore(1000)
end sub


sub addchunk()
'disp
if modename = "truffle" then
trufflejackpot
exit sub
end if

if modename <> "" then
exit sub
end if
if light51.state = 2 then
light51.state = 1
light36.state = 2
playsound "truf1"
exit sub
end if
if light36.state = 2 then
light36.state = 1
light45.state = 2
playsound "truf2"
exit sub
end if
if light45.state = 2 then
light45.state = 1
light23.state = 2
playsound "truf3"
exit sub
end if
if light23.state = 2 then
light23.state = 1
light24.state = 2
playsound "truf4"
exit sub
end if
if light24.state = 2 then
light24.state = 1
light25.state = 2
playsound "truf5"
exit sub
end if

if light25.state = 2 then
light25.state = 1
starttruffle
exit sub
end if
end sub

sub starttruffle()
modename = "truffle"
playsound "truf6"
nevers.state = 1
PlaySong "mu_end"
restartmusic.Interval = 500
restartmusic.Enabled = true ', 500
trufflechrtimer.Interval = 300
trufflechrtimer.Enabled = true ', 300
'sound of chunk
light51.state = 2
light36.state = 2
light45.state = 2
light23.state = 2
light24.state = 2
light25.state = 2
truffletimer.Interval = 30000
truffletimer.Enabled = true ', 30000
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true ', 1000
dispdmd1.text = "TRUFFLE SHUFFLE" & "ALL TARGETS" & "WORTH 50,000"
pDMDsplash3Lines "TRUFFLE SHUFFLE", "ALL TARGETS WORTH", "50,000", 2, 0
end if
end sub
sub endtruffle()
trufflechrtimer.Enabled = false
'chunk sound
truffletimer.Enabled = false
light51.state = 2
light36.state = 0
light45.state = 0
light23.state = 0
light24.state = 0
light25.state = 0
end sub

sub trufflehit()
addscore(50000)
'animation
'sound

end sub

sub trufflejackpot()
addscore(jackpotscore)
playsound "trufhit"
'disp
'sound
end sub
sub truffletimer_Timer()
endtruffle
modename = ""
restartmusic.Interval = 500
restartmusic.Enabled = true ', 500
addscore(0)
end sub

sub trufflechrtimer_Timer()
if tchr = 34 then
tchr = 35
else
tchr = 34
end if

DispDmd1.Text = "[f2][x5][y12]" & formatscore(score(currentplayer)) & "[f2][x2][y0]TRUFFLE SHUFFLE" & "[f1][x2][y25]TARGETS WORTH 50,000" & "[f9][xc][yc]" & chr(tchr) & "[f1] "
pDMDSplash3Lines "TRUFFLE SHUFFLE", "ALL TARGETS WORTH", "50,000", 2, 0
end sub

'----------------------------------
sub addmarble()
if modename <> "" then
exit sub
end if

'disp
if light50.state = 2 then
light50.state = 1
light40.state = 2
marbles = marbles + 1000
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1000
flushdmdtimer.Enabled = true ', 1000
dispdmd1.text = "MARBLE BONUS" & "" & formatscore(marbles) & "" & chr(39)
pDMDSplashLines "MARBLE BONUS", " "& formatscore(marbles), 3, 0
end if
exit sub
end if
if light40.state = 2 then
light40.state = 1
light26.state = 2
marbles = marbles + 2000
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1000
flushdmdtimer.Enabled = true ', 1000
dispdmd1.text = "MARBLE BONUS" & "[f4][xc][y18]" & formatscore(marbles) & " " & chr(39)
pDMDSplashLines "MARBLE BONUS", " "& formatscore(marbles), 3,0
end if
exit sub
end if
if light26.state = 2 then
light26.state = 1
light37.state = 2
marbles = marbles + 3000
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1000
flushdmdtimer.Enabled = true ', 1000
dispdmd1.text = "MARBLE BONUS" & "" & formatscore(marbles) & " " & chr(39)
pDMDSplashLines "MARBLE BONUS", " "& formatscore(marbles), 3, 0
end if
exit sub
end if
if light37.state = 2 then
playsound "marble_bag"
end if
if light37.state = 2 or light37.state = 1 then
light37.state = 1
marbles = marbles + 10000
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1000
flushdmdtimer.Enabled = true ', 1000
dispdmd1.text = "MARBLE BONUS" & " " & formatscore(marbles) & " " & chr(39)
pDMDSplashLines "MARBLE BONUS", " "& formatscore(marbles), 3, 0
end if
exit sub
end if

end sub
'**********************************************************************************
'Spinners
'**********************************************************************************
Sub spinner1_Spin() 'Key
DOF 213, DOFPulse
playsound "fx_spinner"
If Tilted Then Exit Sub
if modename = "key" then
addscore(10000)
keysearch = keysearch - 1
if keysearch = 0 then
keyfound
exit sub
end if
DispDmd1.Text =  "SEARCH FOR KEY" & keysearch & " " & modetime & " " & " " & chr(45)
pDMDSplashLines "SEARCH FOR KEY"& keysearch, " "& modetime, 3, 0
exit sub
end if

keytotal = keytotal + 250
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 500
flushdmdtimer.Enabled = true ', 500
'dispdmd1.text = "KEY","" & formatscore(keytotal) & " " & chr(45)
pDMDSplashLines "KEY","" & formatscore(keytotal),  3,0
end if


addscore(250)



end sub

sub spinner2_Spin()
DOF 214, DOFPulse
playsound "fx_spinner"
If Tilted Then Exit Sub
doubloon = doubloon + 250
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 500
flushdmdtimer.Enabled = true ', 500
dispdmd1.text = "DOUBLOON " & formatscore(doubloon) & " " & chr(44)
pDMDSplashLines "DOUBLOON", "" & formatscore(doubloon), 3, 0
end if

addscore(250)

end sub

'**********************************************************************************
'Swordsman
'**********************************************************************************
sub trigger6_hit()
If Tilted Then Exit Sub
playsound "ballclick"
'if modename = "rscene1" or modename = "rscene2" then
'raidershit()
'exit sub
'end if
playsound "ughh2"
if modename = "" then
swordhit = swordhit + 1
if swordhit = swordneeded then
startsword()
exit sub
end if
end if
if modename <>"" then
exit sub
end if

if scoreupdate = true then
DispDmd1.Text = " "&(swordneeded-swordhit)& " MORE FOR" & " HURRY UP"
pupDMDDisplay "shownum", " " & (swordneeded-swordhit) & "^MORE FOR^HURRY UP", "", 3, 0, 10
scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true ', 500
end if
end sub

sub startsword()
if templeopen = true then
traphitsleft = 2
opencaptive
templeopen = false
end if
playsound "masizefives"
light47.state = 2
nevere1.state = 1
checknever
swordscore = 1500000
DispDmd1.Text = "[f4][x20][y18]" & formatscore(swordscore) & "[f2][x2][y2]HIDE FROM THE" &"[f2][x10][y10]FRATELLIS" & "[f9][xc][yc]" & chr(38)
'pDMDShow3Lines "Hide from", ""&formatscore(swordscore), "the Fratellis", 3, 33023
pupDMDDisplay "default","Hide from^"&formatscore(swordscore)&"^the Fratellis", "", 3, 0, 10
scoreupdate = false
modename = "sword"
swordmodetimer.Interval = 60000
swordmodetimer.Enabled = true', 60000
swordmodetimer2.Interval = 5000
swordmodetimer2.Enabled = true', 5000
restartmusic.Interval = 200
restartmusic.Enabled = true', 200
swordmove()
end sub

sub swordmodetimer2_Timer()
swordscore = swordscore - 1520
if swordscore <250000 then
swordscore = 250000
end if
DispDmd1.Text = "[f4][x20][y18]" & formatscore(swordscore) & "[f2][x2][y2]HIDE FROM THE" &"[f2][x10][y10]FRATELLIS" & "[f9][xc][yc]" & chr(38)
pupDMDDisplay "default","Hide from^"&formatscore(swordscore)&"^the Fratellis", "", 3, 0, 10
swordmodetimer2.Interval = 200
swordmodetimer2.Enabled = true ',60
end sub



sub swordmodetimer_Timer()
light47.state = 0
swordmodetimer2.Enabled = false
DispDmd1.Text =  "[f9][xc][yc]" & chr(43)
playsound "fratlose"
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true ', 500
swordneeded = swordneeded + 3
swordhit = 0
modename = ""
swordmodetimer.Enabled = false
restartmusic.Interval = 200
restartmusic.Enabled = true', 200
swordmove()
end sub
sub swordsmanhit()
'swordman shot animation and score
kicker2timer.Interval = 3000
kicker2timer.Enabled = true ', 3000
lightseq1.Play Seqblinking, , 5,50
playsound "gunshot"
DispDmd1.Text = "[edge4][f7][xc][yc]" & formatscore(swordscore)
pDMDSplashLines " "&formatscore(swordscore), " ", 3,0
flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true ', 500
addscore(swordscore)
swordneeded = swordneeded + 3
swordhit = 0
modename = ""
swordmodetimer.Enabled = false
swordmodetimer2.Enabled = false
restartmusic.Enabled = true', 200

light47.state = 0
end sub
sub kicker2timer_Timer()
kicker2solenoidpulse()
playsound "fx_kicker"
kicker2timer.Enabled = false
end sub

Sub kicker2solenoidpulse()
Kicker2.kick 225, 10
DOF 115, DOFPulse
End Sub

sub swordmove()
'if swordsman.anglexz = 245 then
'swordsman.rotatexz 150,335
'end if
'if swordsman.anglexz = 335 then

'swordsman.RotateXZ -150,245
'end if
end sub

'*********************************************************************************
'Mystery
'*********************************************************************************

Sub startmystery()
neverv.state = 1
checknever
PlaySong "mu_end"
restartmusic.Interval = 5500
restartmusic.Enabled = true', 5200
playsound "databoody"
Select case (RandomNumber(8))
case 1: mysterymode = "SPY EYES" ' bonus at max DONE
pUpdateMystery "DATA GADGET", "", "SPY EYES", 2, 0
case 2: mysterymode = "SLICK SHOES" ' jackpot at max DONE
pUpdateMystery "DATA GADGET", "", "SLICK SHOES", 2, 0
case 3: mysterymode = "THATS NOT A CANDLE" ' shoot ramp for jackpot
pUpdateMystery "DATA GADGET", "", "THATS NOT A CANDLE", 2, 0
case 4: mysterymode = "WINGS OF FLIGHT" ' award jackpot DONE
pUpdateMystery "DATA GADGET", "", "WINGS OF FLIGHT", 2, 0
case 5: mysterymode = "STICKY DART" ' extra ball DONE
pUpdateMystery "DATA GADGET", "", "STICKY DART", 2, 0
case 6: mysterymode = "PINCERS OF PERIL" ' open trap DONE
pUpdateMystery "DATA GADGET", "", "PINCERS OF PERIL", 2, 0
case 7: mysterymode = "BULLY BLINDERS" ' start truffle DONE
pUpdateMystery "DATA GADGET", "", "BULLY BLINDERS", 2, 0
case 8: mysterymode = "BULLY BUSTER" ' marble bonus boost DONE
pUpdateMystery "DATA GADGET", "", "BULLY BUSTER", 2, 0
end select
scoreupdate = false
DispDmd1.Text = " DATA GADGET" & " " & mysterymode & " " & chr(36)
pUpdateMystery "DATA GADGET", "", ""& mysterymode, 3, 0
mysteryend.Interval = 4000
mysteryend.Enabled = true ',4000
mysterychange.Interval = 300
mysterychange.Enabled = true', 200

lightseq1.Play Seqblinking, , 10,50

light27.state = 2
light28.state = 2
light29.state = 2
light30.state = 2

dataD = "D"
dataA = "A"
dataT = "T"
dataA2 = "A"
end sub

sub mysteryend_Timer()
addscore(10000)
'mysterymode = "SLICK SHOES"
mysterychange.Enabled = false
'kicker1.solenoidpulse()
'playsound "fx_kicker"
mysteryend.Enabled = false
kicker1timer.Interval = 1000
kicker1timer.Enabled = true ',1000
'test
'startsnakes()
'exit sub
'mysterymode = "WINGS OF FLIGHT"
if mysterymode = "SPY EYES" then
kicker1timer.Interval = 2500
kicker1timer.Enabled = true ', 2500
flushdmdtimer.Interval = 2500
flushdmdtimer.Enabled = true ', 2500
scoreupdate = false

DispDmd1.Text = "SPY EYES" & " BONUS X AT MAX"
pDMDSplashLines "SPY EYES", "BONUS X AT MAX", 3, 0
mapmulti = 10

exit sub
end if

if mysterymode = "SLICK SHOES" then
kicker1timer.Interval = 2500
kicker1timer.Enabled = true ', 2500
flushdmdtimer.Interval = 2500
flushdmdtimer.Enabled = true ', 2500
scoreupdate = false

DispDmd1.Text = "[xc][y1][f5]SLICK SHOES" & "[xc][y18][f3]JACKPOT AT MAX[edge4]"
pDMDSplashLines "SLICK SHOES", "JACKPOT AT MAX", 3, 0
'light17.state = bulbon
jackpotscore = 10000000
playsound "dataslickshoes"
exit sub
end if

if mysterymode = "THATS NOT A CANDLE" then
kicker1timer.Interval = 2500
kicker1timer.Enabled = true ', 2500
flushdmdtimer.Interval = 2500
flushdmdtimer.Enabled = true ', 2500
scoreupdate = false
playsound "candle"
DispDmd1.Text = "[xc][y1][f4]THATS NOT A CANDLE" & "[xc][y18][f2]HIT RAMP FOR JACKPOT[edge4]"
pDMDSplashLines "THATS NOT A CANDLE", "HIT RAMP FOR JACKPOT", 3, 0
'to do add flag for ramp jackpot
light39.state = 2
light49.state = 2
light44.state = 2


exit sub
end if
if mysterymode = "WINGS OF FLIGHT" then
kicker1timer.Interval = 8500
kicker1timer.Enabled = true ', 8500
flushdmdtimer.Interval = 9500
flushdmdtimer.Enabled = true' , 9500
restartmusic.Interval = 8500
restartmusic.Enabled = true', 8500
scoreupdate = false
playsound "datawings"
dmdfont = "WINGS OF FLIGHT"

matchchr = 32
endmatchchr = 125
dmdanimtimer.Interval = 100
dmdanimtimer.Enabled = true ', 100
dmdspeed = 60
endtext = "WINGS OF FLIGHT" & " JACKPOT AWARDED"
pDMDSplashLines "WINGS OF FLIGHT", "JACKPOT AWARDED", 3, 0
addscore(jackpotscore)
exit sub
end if
if mysterymode = "STICKY DART" then
kicker1timer.Interval = 8500
kicker1timer.Enabled = true' , 8500
flushdmdtimer.Interval = 9500
flushdmdtimer.Enabled = true ', 9500
restartmusic.Interval = 8500
restartmusic.Enabled = true', 8500
scoreupdate = false
playsound "datadart"
dmdfont = "EXTRA BALL"

matchchr = 32
endmatchchr = 126
dmdanimtimer.Interval = 100
dmdanimtimer.Enabled = true ', 100
dmdspeed = 60
endtext = "STICKY DART" & " EXTRA BALL!"
pDMDSplashLines "STICKY DART", "EXTRA BALL", 5, 33023

ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
shootagainlight.state = 1
exit sub
end if
if mysterymode = "PINCERS OF PERIL" then
kicker1timer.Interval = 2500
kicker1timer.Enabled = true ', 2500
flushdmdtimer.Interval = 2500
flushdmdtimer.Enabled = true ', 2500
scoreupdate = false
playsound "datapincers"
DispDmd1.Text = "PINCERS OF PERIL" & " TRAP OPEN "
pDMDSplashLines "PINCERS OF PERIL", "TRAP OPEN", 3, 0

if templeopen = false then
opencaptive()
end if
exit sub
end if
if mysterymode = "BULLY BLINDERS" then
playsound "databullyblinders"
kicker1timer.Interval = 2500
kicker1timer.Enabled = true ', 2500
flushdmdtimer.Interval = 2500
flushdmdtimer.Enabled = true ', 2500
scoreupdate = false

DispDmd1.Text = "BULLY BLINDERS" & "START TRUFFLE SHUFFLE"
pDMDSplashLines "BULLY BLINDERS", "START TRUFFLE SHUFFLE", 3, 0
delaytruffle.Interval = 2500
delaytruffle.Enabled = true', 2500
exit sub
end if
if mysterymode = "BULLY BUSTER" then

kicker1timer.Interval = 2500
kicker1timer.Enabled = true ', 2500
flushdmdtimer.Interval = 2500
flushdmdtimer.Enabled = true ', 2500
scoreupdate = false

DispDmd1.Text = "BULLY BUSTER" & " BOOST MARBLE BONUS"
pDMDSplashLines "BULLY BUSTER", "BOOST MARBLE BONUS", 3, 0
marbles = marbles + 10000
exit sub
end if

flushdmdtimer.Interval = 100
flushdmdtimer.Enabled = true',100
end sub

sub delaytruffle_Timer()
delaytruffle.Enabled = false
starttruffle
end sub

sub mysterychange_Timer()
Select case (RandomNumber(8))
case 1: mysterymode = "SPY EYES"
 : pUpdateMystery "", "","SPY EYES", 4, 0 ':datachr = 36:
case 2: mysterymode = "SLICK SHOES"
: pUpdateMystery "", "", "SLICK SHOES", 4, 0 ': datachr = 37
case 3: mysterymode = "THATS NOT A CANDLE"
: pUpdateMystery "", "", "THATS NOT A CANDLE", 4, 0' :datachr = 36
case 4: mysterymode = "WINGS OF FLIGHT"
: pUpdateMystery "", "", "WINGS OF FLIGHT", 4, 0' :datachr = 36
case 5: mysterymode = "STICKY DART"
: pUpdateMystery "", "", "STICKY DART", 4, 0 ':datachr = 37
case 6: mysterymode = "PINCERS OF PERIL"
: pUpdateMystery "", "", "PINCERS OF PERIL", 4, 0 ':datachr = 36
case 7: mysterymode = "BULLY BLINDERS"
: pUpdateMystery "", "", "BULLY BLINDERS", 4, 0 ':datachr = 37
case 8: mysterymode = "BULLY BUSTER"
: pUpdateMystery "", "", "BULLY BUSTER", 4, 0 ':datachr = 36
end select
DispDmd1.Text = "DATA GADGET" & " " & mysterymode' & " " & chr(datachr)
pUpdateMystery "DATA GADGET" , "", "" & mysterymode , 4, 0
end sub

sub startrats()
ratsendtimer.Interval = 20000
ratsendtimer.Enabled = true', 20000
modename = "rats"
hurrytotal = 2500000
flushdmdtimer.Enabled = false
scoreupdate = false
DispDmd1.Text = " " & score(currentplayer) & " " & lastscore &" " & "RATS" & "Shoot orbits now!" & " " & formatscore(hurrytotal)
pDMDSplashLines "RATS" , "SHOOT ORBITS NOW!" & " " & formatscore(hurrytotal), 2 ,0
light50.state = 2
light51.state = 2
ratstimer1.Interval = 3000
ratstimer1.Enabled = true', 3000
end sub

sub ratstimer1_Timer()
ratstimer1.Enabled = false
ratstimer2.Interval = 70
ratstimer2.Enabled = true', 70
end sub

sub ratstimer2_Timer()
hurrytotal = hurrytotal - 10020
if hurrytotal < 500000 then
hurrytotal = 500000
end if

DispDmd1.Text = " " & score(currentplayer) & " " & lastscore &" " & "RATS" & "Shoot orbits now!" & " " & formatscore(hurrytotal)
pDMDSplashLines "RATS" , "SHOOT ORBITS NOW!" & " " & formatscore(hurrytotal), 2 ,0
end sub

sub ratsendtimer_Timer()
light50.state = 0
light51.state = 0
ratsendtimer.Enabled = false
ratstimer2.Enabled = false
scoreupdate = true
flushdmdtimer.Interval = 200
flushdmdtimer.Enabled = true', 200
modename = ""
restartmusic.Interval = 200
restartmusic.Enabled = true ', 200
end sub


sub startants()
antsendtimer.Interval = 20000
antsendtimer.Enabled = true', 20000
modename = "ants"
hurrytotal = 2500000
flushdmdtimer.Enabled = false
scoreupdate = false
DispDmd1.Text = " " & score(currentplayer) & " " & lastscore &" " & "ANTS" & " Shoot orbits now!" & " " & formatscore(hurrytotal)
pDMDSplashLines "ANTS" , "SHOOT ORBITS NOW!" & " " & formatscore(hurrytotal), 2 ,0
light50.state = 2
light51.state = 2
antstimer1.Interval = 3000
antstimer1.Enabled = true', 3000
end sub

sub antstimer1_Timer()
antstimer1.Enabled = false
antstimer2.Interval = 70
antstimer2.Enabled = true', 70
end sub

sub antstimer2_Timer()
hurrytotal = hurrytotal - 10020
if hurrytotal < 500000 then
hurrytotal = 500000
end if

DispDmd1.Text = " " & score(currentplayer) & " " & lastscore &" " & "ANTS" & " Shoot orbits now!" & "" & formatscore(hurrytotal)
pDMDSplashLines "ANTS" , "SHOOT ORBITS NOW!" & " " & formatscore(hurrytotal), 2 ,0
end sub

sub antsendtimer_Timer()
light50.state = 0
light51.state = 0
antsendtimer.Enabled = false
antstimer2.Enabled = false
scoreupdate = true
flushdmdtimer.Interval = 200
flushdmdtimer.Enabled = true', 200
modename = ""
restartmusic.Interval = 200
restartmusic.Enabled = true' , 200
end sub


sub startsnakes()

snakesendtimer.Interval = 20000
snakesendtimer.Enabled = true', 20000
modename = "snakes"
hurrytotal = 2500000
flushdmdtimer.Enabled = false
scoreupdate = false
DispDmd1.Text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f4][x105][y10]SNAKES" & "[f4][x70][y45]Shoot orbits now!" & "[f4][x100][y28]" & formatscore(hurrytotal)
pDMDSplashLines "SNAKES" , "SHOOT ORBITS NOW!" & " " & formatscore(hurrytotal), 2 ,0
light50.state = 2
light51.state = 2
snakestimer1.Interval = 3000
snakestimer1.Enabled = true', 3000
end sub
sub snakestimer1_Timer()
snakestimer1.Enabled = false
snakestimer2.Interval = 70
snakestimer2.Enabled = true', 70
end sub

sub snakestimer2_Timer()
hurrytotal = hurrytotal - 10020
if hurrytotal < 500000 then
hurrytotal = 500000
end if

DispDmd1.Text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & "[f4][x105][y10]SNAKES" & "[f4][x70][y45]Shoot orbits now!" & "[f4][x100][y28]" & formatscore(hurrytotal)
pDMDSplashLines "SNAKES" , "SHOOT ORBITS NOW!" & " " & formatscore(hurrytotal), 2 ,0
end sub


sub snakesendtimer_Timer()
light50.state = 0
light51.state = 0
snakesendtimer.Enabled = false
snakestimer2.Enabled = false
scoreupdate = true
flushdmdtimer.Interval = 500
flushdmdtimer.Enabled = true', 500
modename = ""
restartmusic.Interval = 200
restartmusic.Enabled = true ', 200
end sub


'**********************************************************************************
sub startoverlay()
'overlay1.fadein()
'overlay1.frame oframestart
'otimer.set true,ospeed
end sub
sub otimer_Timer()
oframestart = oframestart + 1
'overlay1.frame oframestart
if oframestart = oframeend then
otimer.Enabled = false
'overlay1.fadeout()
dispdmd1.text = endtext
endtext = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]0" &"[line3,59,0,59,64] [x70] [yc] [f7]" & formatscore(score(currentplayer)) & "[f1][x66][y57]BALL:" & (6 -BallsRemaining(CurrentPlayer)) & "[f1][x105][y57]PLAYER:" & currentplayer & "[f1][x146][y57]CREDITS:" & nvcredits


exit sub
end if
otimer.Enabled = true ', ospeed
end sub

sub endtexttimer_Timer()
endtexttimer.Enabled = false
dispdmd1.text = endtext
end sub





sub sounddelay_Timer()
playsound tempsound
sounddelay.Enabled = false
end sub

sub resetlights()
Dim x
 for x= 20 to 56
  Execute "Light" & x & ".State=BulbOff"
 next
light1.state = 0

end sub



'**********************************************************************************
'DMD animations
'**********************************************************************************

sub startanimation()
currentframe = startframe + 31
endframe = endframe + 31
dispdmd1.text =  "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & dmdfont & "[x60][yc]" & chr(currentframe)
'dmdanimtimer.Interval = dmdspeed
dmdanimtimer.Enabled = true ', dmdspeed
end sub

sub dmdtimer_Timer()
dmdtimer.Enabled = false
currentframe = currentframe + 1
dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64]" & dmdfont & "[x60][yc]" & chr(currentframe)
if currentframe = endframe then
if repeatdmd = true then

currentframe = startframe + 31
'dmdanimtimer.Interval = dmdspeed
dmdanimtimer.Enabled = true ', dmdspeed
exit sub
end if
exit sub
end if
'dmdanimtimer.Interval = dmdspeed
dmdanimtimer.Enabled = true ', dmdspeed
end sub

'****************************************************************************
'show jackpot
'****************************************************************************
sub showjptimer_Timer()
dmdtimer.Enabled = false
'flushdmdtimer.set true , 1500
playsound "excellent"
scoreupdate = false
dispdmd1.text = "[f1][x0] [y0]" & score(currentplayer) & "[f1][x0] [y6]" & lastscore &"[line3,59,0,59,64][f8][x60][yc]" & chr(59) & "[f7][x80][yc]" & formatscore(jpscore)
pDMDSplashLines " " & score(currentplayer) & " " & lastscore & " " & chr(59) , " " & formatscore(jpscore), 2 ,0
showjptimer.Enabled = false
addscore(jpscore)
end sub


'*****************************************************************************
'next scene
'*****************************************************************************

sub nextscene()






end sub


'******************************************************************************
'eob bonus
'******************************************************************************

sub starteob()
pDMDSetPage(pDMDBlank)
flushdmdtimer.Enabled = false
PlaySong "Mu_End" :PuPEvent 801
playsound "bonus1"
scoreupdate = false
dispdmd1.text = "BONUS"
pupDMDDisplay "BONUS", "BONUS", "", 4, 0, 10
'pNote "BONUS", ""
totalbonus =  ballbonus
if keytotal > 0 then
eobark.Interval = 1500
eobark.Enabled = true',1500
exit sub
end if

if doubloon > 0 then
eobtemple.Interval = 1500
eobtemple.Enabled = true',1500
exit sub
end if

if grailhit > 0 then
eobgrail.Interval = 1500
eobgrail.Enabled = true',1500
exit sub
end if

if rampshots > 0 then
eobskull.Interval = 1500
eobskull.Enabled = true',1500
exit sub
end if

if marbles > 0 then
eobmarble.Interval = 1500
eobmarble.Enabled = true',1500
exit sub
end if

eobmulti.Interval = 1500
eobmulti.Enabled = true',1500
end sub

sub eobark_Timer()
playsound "bonus1"
eobark.Enabled = false
dispdmd1.text = "KEY " & formatscore(keytotal) & " " & chr(45)
pupDMDDisplay "default", "KEYS^"&formatscore(keytotal), "", 2, 0, 10
'pNote "KEY " & formatscore(keytotal) & " " & chr(45), ""

totalbonus = totalbonus + (keytotal)

if doubloon > 0 then
eobtemple.Interval = 1500
eobtemple.Enabled = true',1500
exit sub
end if
if grailhit > 0 then
eobgrail.Interval = 1500
eobgrail.Enabled = true',1500
exit sub
end if
if rampshots > 0 then
eobskull.Interval = 1500
eobskull.Enabled = true',1500
exit sub
end if

if marbles > 0 then
eobmarble.Interval = 1500
eobmarble.Enabled = true',1500
exit sub
end if

eobmulti.Interval = 1500
eobmulti.Enabled = true', 1500
end sub

sub eobtemple_Timer()
playsound "bonus1"
eobtemple.Enabled = false
dispdmd1.text = "DOUBLOON " & formatscore(doubloon) & " " & chr(44)
'pDMDSplashBonus "DOUBLOON" , formatscore(keytotal) & "", 2 ,33023  doubloon or keytotal? bug?
pupDMDDisplay "default", "DOUBLOON^"&formatscore(doubloon), "", 2, 0, 10
'pNote "DOUBLOON " & formatscore(doubloon) & " " & chr(44), ""
totalbonus = totalbonus + (doubloon)

if grailhit > 0 then
eobgrail.Interval = 1500
eobgrail.Enabled = true',1500
exit sub
end if
if rampshots > 0 then
eobskull.Interval = 1500
eobskull.Enabled = true',1500
exit sub
end if

if marbles > 0 then
eobmarble.Interval = 1500
eobmarble.Enabled = true','1500
exit sub
end if

eobmulti.Interval = 1500
eobmulti.Enabled = true', 1500
end sub

sub eobgrail_Timer()
playsound "bonus1"
eobgrail.Enabled = false

dispdmd1.text = " " & score(currentplayer) & " " & lastscore &" " & chr(72) & " " & grailhit & " GRAIL HITS " & formatscore(grailhit*75000)
pDMDSplashBonus " GRAIL HITS ", "" & grailhit & "" & formatscore(grailhit*75000) , 2 ,33023
'pNote " " & score(currentplayer) & " " & lastscore &" " & chr(72) & " " & grailhit & " GRAIL HITS " & formatscore(grailhit*75000), ""

totalbonus = totalbonus + (grailhit*75000)

if rampshots > 0 then
eobskull.Interval = 1500
eobskull.Enabled = true',1500
exit sub'
end if


if marbles > 0 then
eobmarble.Interval = 1500
eobmarble.Enabled = true',1500
exit sub
end if
eobmulti.Interval = 1500
eobmulti.Enabled = true', 1500
end sub

sub eobskull_Timer()
playsound "bonus1"
eobskull.Enabled = false

if rampshots = 1 then
dispdmd1.text = " " & rampshots & " RAMP HIT" & " " & chr(33)
pDMDSplashBonus "RAMP HITS" ,""&rampshots, 2, 33023
'pNote " " & rampshots & " RAMP HIT" & " " & chr(33) , ""
else
dispdmd1.text = " " & rampshots & " RAMP HITS"  & " " & chr(33)
pDMDSplashBonus "RAMP HITS" , "" & rampshots, 2, 33023
'pNote " " & rampshots & " RAMP HITS"  & " " & chr(33), ""
end if
totalbonus = totalbonus + (rampshots*100000)

if marbles > 0 then
eobmarble.Interval = 2000
eobmarble.Enabled = true',1500
exit sub
end if
eobmulti.Interval = 2000
eobmulti.Enabled = true', 1500
end sub

sub eobmarble_Timer()
playsound "bonus1"
eobmarble.Enabled = false



dispdmd1.text = "MARBLE BONUS"  & " " & formatscore(marbles) & " " & chr(39)
pDMDSplashBonus "MARBLE BONUS"  , "" & formatscore(marbles), 1, 33023
totalbonus = totalbonus + (marbles)

eobmulti.Interval = 1500
eobmulti.Enabled = true', 1500
end sub
sub eobmulti_Timer()
playsound "bonus1"
dispdmd1.text = " " & mapmulti & "X " & formatscore(totalbonus)
pDMDSplashBonus mapmulti & " X"   , "" & formatscore(totalbonus) , 1, 33023
eobmulti.Enabled = false
totalbonus = (totalbonus * mapmulti)

eobtotal.Interval = 1500
eobtotal.Enabled = true',1500
end sub

sub eobtotal_Timer()
playsound "bonus2"
eobtotal.Enabled = false
dispdmd1.text = "TOTAL" &" " & formatscore(totalbonus)
pupDMDDisplay "default", "TOTAL^" & formatscore(totalbonus) , "", 3, 0, 10
addscore(totalbonus)
checkreplay()
flushdmdtimer.Enabled = false
end sub


'captive.ty = captive.ty-35


sub opencaptive()
'if captive.angleyz <>0 or captive.angleyz<>30 then
'exit sub
'end if
'if captiveP.RotX = 20 then
if captiveP.RotX > 15 then
  DOF 129, DOFPulse
  stonetarget.Isdropped = false
  stonetarget1.Isdropped = false
  stonetarget2.Isdropped = false
  'stonetarget.render = false
  'captiveP.RotX = 0',0
  captiveDown 35, 0
else
  DOF 129, DOFPulse
  templeopen = true
  traphitsleft = 2
  droptimer.Interval = 800
  droptimer.Enabled = true',800
  'captiveP.RotX = 20',20
  captiveUp 0, 35
end if

end sub

sub droptimer_Timer()
droptimer.Enabled = false
stonetarget.Isdropped = true
stonetarget1.Isdropped = true
stonetarget2.Isdropped = true
end sub

 Dim HPos, HPosEnd
 Sub captivePAnim_timer()
    captiveP.RotX = HPos
  If Hpos < HposEnd Then
     HPos = HPos + 1
  Else
   'Popupsolenoidpulse.enabled = 0
  End If
end Sub

 Sub captivePAnim2_timer()
    captiveP.RotX = HPos
  If Hpos > HposEnd Then
     HPos = HPos - 1
  Else
   captivePAnim2.enabled = 0
  End If
end Sub

Sub captiveUp(FrameStart, FrameEnd)
    'popup.Isdropped = 0
    HPos = FrameStart
    HPosEnd = FrameEnd
  captivePAnim.enabled = 1
End Sub

Sub captiveDown(FrameStart, FrameEnd)
    'popup.Isdropped = 1
    HPos = FrameStart
    HPosEnd = FrameEnd
  captivePAnim2.enabled = 1
End Sub

sub plungertimer_Timer()
  Plungersolenoidpulse
  PlaySound "cannon"
    'playsound "gunshot"
autoball = true
plungertimer.Enabled = false
end sub

Sub Plungersolenoidpulse()
  DOF 111, DOFPulse
  DOF 121, DOFPulse
    plungerIM.AutoFire
    LaserKickP.TransY = 90
    vpmtimer.addtimer 400, "LaserKickRes '"
End Sub

Sub LaserKickRes()
    LaserKickP.TransY = 0
End Sub

sub titletimer_Timer()
titletimer.Enabled = false

endtext = "  "
matchchr = 32
endmatchchr = 89
dmdfont = "  "
dmdspeed = 80
dmdanimtimer.Interval = 100
dmdanimtimer.Enabled = true ', 100
end sub


sub matchgame()
SollFlipper 0
SolrFlipper 0
LeftFlipper.RotateToStart
RightFlipper.RotateToStart
matchchr = 32
endmatchchr = 61
playermatch =  right(Score(CurrentPlayer),2)
dispdmd1.text = " " & playermatch & " " & chr(matchchr)
pDMDSplash3Lines " " & playermatch , " "&(matchchr)," ",  2, 0
matchdmdtimer.Interval = 150
matchdmdtimer.Enabled = true ', 150
select case (randomnumber(10))
case 1: indymatch = "00"
case 2: indymatch = "10"
case 3: indymatch = "20"
case 4: indymatch = "30"
case 5: indymatch = "40"
case 6: indymatch = "50"
case 7: indymatch = "60"
case 8: indymatch = "70"
case 9: indymatch = "80"
case 10: indymatch = "90"
end select
'playermatch = indymatch

PlaySong "mu_end"
endofgamedelay.Interval = 7500
endofgamedelay.Enabled = true' ,10000
endmatchdelay.Interval = 6500
endmatchdelay.Enabled = true ', 6500
'startoverlay()
end sub
sub matchdmdtimer_Timer()
matchdmdtimer.Interval = 150
matchdmdtimer.Enabled = true ', 150
matchchr = matchchr + 1
if matchchr > endmatchchr then
matchchr = endmatchchr
matchdmdtimer.Enabled = false
end if
if matchchr > 55 then

dispdmd1.text = " " & indymatch & " " & playermatch & " " & chr(matchchr)
pDMDSplash3Lines " " & indymatch & " " & playermatch , " " &(matchchr)," ",  2, 0
else

dispdmd1.text = " " & playermatch & " " & chr(matchchr)
pDMDSplash3Lines " " & playermatch , " " &(matchchr)," ",  2, 0
end if

end sub


sub endofgamedelay_Timer()
matchdmdtimer.Enabled = false
endofgamedelay.Enabled = false
endofgame()
end sub

sub endmatchdelay_Timer()
endmatchdelay.Enabled = false
if playermatch = indymatch then
credits = credits +1
DOF 136, DOFOn
'PlaySound SoundFXDOF("fx_knocker",124,DOFPulse,DOFKnocker)
PlaySound SoundFXDOF("fx_knocker",135,DOFPulse,DOFKnocker)
DOF 121, DOFPulse
else
select case (randomnumber(2))
case 1: playsound "seeyoutomorrowindy"
case 2: playsound "ikeeptellingyou"
end select

end if
end sub

sub collectspecial()
'light17.state = bulboff
'light19.state = bulboff
playsound "excellentfanfare"
addscore(50000000)
end sub




sub balldropsound_Timer()
balldropsound.Enabled = false
playsound "ballfall"
end sub
sub trigger8_hit()
balldropsound.Enabled = true' , 150
end sub

sub fakegate1hit()
if not fakegate1.IsDropped then
          fakegate1.popdown
        exit sub
      end if

      if fakegate1.IsDropped then
        fakegate1.SolenoidPulse
        exit sub
      end if
end sub

sub trigger9_hit()
PlaySoundAt "fx_PlasticRamp", trigger9
if not fakegate2.IsDropped then
          fakegate2.IsDropped = 1
        exit sub
      end if

      if fakegate2.IsDropped then
        fakegate2.IsDropped = 0
        exit sub
      end if
end sub

sub trigger10_hit()
playsound "ballrollinga"
end sub


sub vuk1_hit()
'vuk1.kick 0, 35, 1.5
'PlaySoundAt "fx_vukout_LAH" ,vuk1
End Sub

Sub Vuk1SolenoidPulse_Timer()
Vuk1SolenoidPulse.Enabled = False
vuk1.kick 0, 55, 1.5
Playsound SoundFXDOF("fx_vukout_LAH",116,DOFPulse,DOFContactors)
End Sub



sub rightlight()
dim tempstate
tempstate = light19.state
light19.state = light18.state
light18.state = light42.state
light42.state = light17.state
light17.state = tempstate
end sub

sub leftlight()
dim tempstate
tempstate = light17.state
light17.state = light42.state
light42.state = light18.state
light18.state = light19.state
light19.state = tempstate
end sub

sub checkrich()

if light17.state = 1 and light42.state = 1 and light18.state = 1 and light19.state = 1 then
mapmulti = mapmulti + 1
playsound "rich_stuff"
if mapmulti >= 10 then
mapmulti = 10
end if

LightSeqInserts.Play Seqblinking, , 3,50
if scoreupdate = true then
DispDmd1.Text = " " & mapmulti & "X"
scoreupdate = false
flushdmdtimer.Interval = 2000
flushdmdtimer.Enabled = true ',2000
playsound "jingle15"
end if

flashforms light17, 2000, 50, 0
flashforms light18, 2000, 50, 0
flashforms light19, 2000, 50, 0
flashforms Light42, 2000, 50, 0

light17.state = 0
light18.state = 0
light19.state = 0
Light42.state = 0
addscore(10000)
end if


end sub
'door.MoveTo door.Tx, door.Ty-8, door.Tz , 1000


Sub nevertimer1_Timer()
nevertimer1.Enabled = false
nevern.state = 1:NeverNP.visible = 1
nevertimer2.Interval = 1100
nevertimer2.Enabled =  true
end sub
Sub nevertimer2_Timer()
nevertimer2.Enabled =  false
nevere1.state = 1: NeverEP.visible = 1
nevertimer3.Interval = 800
nevertimer3.Enabled =  true
end sub
Sub nevertimer3_Timer()
nevertimer3.Enabled =  false
neverv.state = 1: NeverVP.visible = 1
nevertimer4.Interval = 800
nevertimer4.Enabled =  true
end sub
Sub nevertimer4_Timer()
nevertimer4.Enabled =  false
nevere2.state = 1: NeverE2P.visible = 1
nevertimer5.Interval = 800
nevertimer5.Enabled =  true
end sub
Sub nevertimer5_Timer()
nevertimer5.Enabled =  false
neverr.state = 1: NeverRP.visible = 1
nevertimer6.Interval = 800
nevertimer6.Enabled =  true
end sub
Sub nevertimer6_Timer()
nevertimer6.Enabled =  false
nevers.state = 1: NeverSP.visible = 1
nevertimer7.Interval = 1500
nevertimer7.Enabled =  true
end sub
Sub nevertimer7_Timer()
nevertimer7.Enabled =  false
neverA.state = 1: NeverAP.visible = 1
nevertimer8.Interval = 800
nevertimer8.Enabled =  true
end sub
Sub nevertimer8_Timer()
nevertimer8.Enabled =  false
nevery.state = 1: NeverYP.visible = 1
nevertimer9.Interval = 800
nevertimer9.Enabled =  true
end sub
Sub nevertimer9_Timer()
nevertimer9.Enabled =  false
neverd.state = 1: NeverDP.visible = 1
nevertimer10.Interval = 1300
nevertimer10.Enabled =  true
end sub
Sub nevertimer10_Timer()
nevertimer10.Enabled =  false
neveri.state = 1: NeverIP.visible = 1
nevertimer11.Interval = 800
nevertimer11.Enabled =  true
end sub
Sub nevertimer11_Timer()
nevertimer11.Enabled =  false
nevere3.state = 1: NeverEP3.visible = 1
neverreset.Interval = 2500
neverreset.Enabled =  true
end sub
sub neverreset_Timer()
neverreset.Enabled =  false
nevern.state = 0: NeverNP.visible = 0
nevere1.state = 0: NeverEP.visible = 0
neverv.state = 0: NeverVP.visible = 0
nevere2.state = 0: NeverE2P.visible = 0
neverr.state = 0: NeverRP.visible = 0
nevers.state= 0: NeverSP.visible = 0
nevera.state = 0: NeverAP.visible = 0
nevery.state = 0: NeverYP.visible = 0
neverd.state = 0: NeverDP.visible = 0
neveri.state = 0: NeverIP.visible = 0
nevere3.state = 0: NeverEP3.visible = 0
nevertimer1.Interval = 2500
nevertimer1.Enabled =  true
end sub


sub checknever()
if nevern.state = 1 and nevere1.state = 1 and neverv.state = 1 and nevere2.state = 1 and neverr.state = 1 and nevers.state = 1 and nevera.state = 1 and nevery.state = 1 and neverd.state = 1 and neveri.state = 1 and nevere3.state = 1 then
light46.state = 2
playsound "neversaydie"
end if

end sub

sub resetdoor_Timer()
resetdoor.Enabled = false

bulb4.state = 0
'light43.state = bulboff
'fakedoor.isdropped = false
'fakedoor.render = false
'door.MoveTo door.Tx, door.Ty+40, door.Tz , 50
DoorUp 0, 60
end sub


'Bonus Target
sub target20_hit()
DOF 128, DOFPulse
if modename = "key" or modename = "bone" or modename = "boulder" or modename = "water" or modename = "fight" then
modetime = modetime + 5
exit sub
end if

if light1.state = 2 then
ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
shootagainlight.state = 1
light1.state = 0
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true', 1500
dispdmd1.text = "EXTRA BALL"
pDMDSplashBig "EXTRA BALL", 3, 33023
end if
exit sub
end if

if slothmulti = 2 then
slothtimer.Interval = 30000
slothtimer.Enabled = true ', 30000
playsound "slothhey"
exit sub
end if

if wellsound = true then
wellsound = false
resetwell.Interval = 5000
resetwell.Enabled = true ', 5000
wellhit = wellhit + 1'
if wellhit = 1 then
playsound "wishing_well"
end if
if wellhit = 2 then
playsound "andybucket"
end if
if wellhit = 3 then
playsound "andyyougoonie"
'reward
addscore(500000)
if scoreupdate = true then
scoreupdate = false
flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true', 1500
dispdmd1.text = "WELL BONUS"
pDMDSplashBig "WELL BONUS" , 3, 33023
end if
wellhit = 0
end if
end if


end sub


Dim HPos2, HPosEnd2
 Sub DoorUpSolenoidpulse_timer()
    FakedoorP.Z = HPos2
  If Hpos2 < HposEnd2 Then
     HPos2 = HPos2 + 1
  Else
   DoorUpSolenoidpulse.enabled = 0
  End If
end Sub

 Sub DoorDownSolenoidpulse_timer()
    FakedoorP.Z = HPos2
  If Hpos2 > HposEnd2 Then
     HPos2 = HPos2 - 1
  Else
   DoorDownSolenoidpulse.enabled = 0
  End If
end Sub

Sub DoorUp(FrameStart, FrameEnd)
    Fakedoor.Isdropped = 0
    HPos2 = FrameStart
    HPosEnd2 = FrameEnd
  DoorUpSolenoidpulse.enabled = 1
End Sub

Sub DoorDown(FrameStart, FrameEnd)
    Fakedoor.Isdropped = 1
    HPos2 = FrameStart
    HPosEnd2 = FrameEnd
  DoorDownSolenoidpulse.enabled = 1
End Sub









sub resetwell_Timer()
resetwell.Enabled = false
wellsound = true
end sub

sub trigger13_hit()
bulb1.state = 2
end sub

sub trigger12_hit()
bulb1.state = 0
end sub

sub combotimer_Timer()
comboon = false
combotimer.Enabled = false
end sub

sub addcombo()
addscore(combovalue)
if scoreupdate = true then
flushdmdtimer.Interval = 1500
flushdmdtimer.Enabled = true ', 1500
scoreupdate = false
dispdmd1.text = " COMBO" & " " & formatscore(combovalue)
pDMDSplashLines " COMBO", ""& formatscore(combovalue), 1, 1
end if
combovalue = combovalue + 50000
end sub

sub dmdanimtimer_Timer()
'dmdanimtimer.Interval = dmdspeed
dmdanimtimer.Enabled = true ', dmdspeed
matchchr = matchchr + 1
if matchchr => endmatchchr then
matchchr = endmatchchr
dmdanimtimer.Enabled = false
dispdmd1.text = dmdfont & " " & chr(matchchr) & endtext & "  "
pDMDSplash3Lines dmdfont & " ",  chr(matchchr) & endtext &"  "," ", 1, 0
exit sub
end if

dispdmd1.text = dmdfont & " " & chr(matchchr) & "  "
pDMDSplash3Lines dmdfont & " ",  chr(matchchr) & "  "," ", 1, 0
end sub

sub resettable()
modejp = false

DOF 122, DOFPulse
target10.Isdropped = 0
target11.Isdropped = 0
target12.Isdropped = 0
target13.Isdropped = 0
target14.Isdropped = 0
target15.Isdropped = 0
target16.Isdropped = 0
tchr = 34
modetime = 0
keysearch = 0
wellsound = true
wellhit = 0
autoball = false
'ballslocked = 0
loopd = ""
titletimer.Enabled = false
dmdanimtimer.Enabled = false
slothmulti = 1
jackpotscore = 1000000
marbles = 0
comboon = false
combovalue = 50000
  Dim i
if captiveP.RotX > 15 then

stonetarget.Isdropped = false
stonetarget1.Isdropped = false
stonetarget2.Isdropped = false
'stonetarget.render = false
'captiveP.RotX = 0
captiveDown 35, 0
end if

resetxmarks.Enabled = false
resetlights()
LightSeq1.stopplay()
nevertimer1.Enabled = false
nevertimer2.Enabled = false
nevertimer3.Enabled = false
nevertimer4.Enabled = false
nevertimer5.Enabled = false
nevertimer6.Enabled = false
nevertimer7.Enabled = false
nevertimer8.Enabled = false
nevertimer9.Enabled = false
nevertimer10.Enabled = false
nevertimer11.Enabled = false

neverreset.Enabled = false
nevern.state = 0: NeverNP.visible = 0
nevere1.state = 0: NeverEP.visible = 0
neverv.state = 0: NeverVP.visible = 0
nevere2.state = 0: NeverE2P.visible = 0
neverr.state = 0: NeverRP.visible = 0
nevers.state= 0: NeverSP.visible = 0
nevera.state = 0: NeverAP.visible = 0
nevery.state = 0: NeverYP.visible = 0
neverd.state = 0: NeverDP.visible = 0
neveri.state = 0: NeverIP.visible = 0
nevere3.state = 0: NeverEP3.visible = 0
'playmusic 1 ,"ballready",true
light27.state = 2
light28.state = 2
light29.state = 2
light30.state = 2
light31.state = 2
light32.state = 2
light33.state = 2
light34.state = 2
light35.state = 2
light51.state = 2
light48.state = 2
liftno = 1
scoreupdate = true
lastscore =0
spellarkframe = 42
organhitsleft = 7
tempsound = ""
raidersjp = 1000000
templejp = 1000000
crystaljp = 1000000
crusadejp = 1000000
modename = ""
mapmulti = 1
swordhit = 0
swordneeded = 5
nextscenescore = 1000000
indyscoring = false
jonesscoring = false
indyhitsleft = 0
joneshitsleft = 0
ballbonus = 0
totalbonus = 0
arkhit = 0
templehit = 0
grailhit = 0
skullhit = 0
hitskull = true
mysterymode = ""
xmarks = 1
xmarkshits = 3
traphitsleft = 2
stonesleft = 3
stonekicker = false
jackpotanim = false
templeopen = false
stonescollected = 0
rampshots = 0
rampshotsleft = 5
dataD = "  "
dataA = "  "
dataT = "  "
dataA2 = "  "
slothS = "  "
slothL = "  "
slothO = "  "
slothT = "  "
slothH = "  "

end sub


sub startwizard()
PlaySong "mu_end"
restartmusic.Interval = 2000
restartmusic.Enabled = true', 2000
dispdmd1.text = " WIZARD "
pDMDSplashBig "WIZARD", 3, 33023
modename = "wizard"
light43.state = 0
light46.state = 0
light50.state = 2
light40.state = 2
light37.state = 2
light26.state = 2
light51.state = 2
light36.state = 2
light45.state = 2
light23.state = 2
light24.state = 2
light25.state = 2

light44.state = 2
light39.state = 2
light49.state = 2
plungertimer.Interval = 1500
plungertimer.Enabled = true ', 1500
createnewball
bmultiballmode = true
kicker7timer.Interval = 5000
kicker7timer.Enabled = true ', 5000
end sub

sub endwizard()
modename = ""
resettable
end sub

sub wizardjp()
addscore(jackpotscore)
modejp = true
scoreupdate = false
flushdmdtimer.Interval = 2500
flushdmdtimer.Enabled = true ',2500
dispdmd1.text =  "JACKPOT"
pDMDSplashBig "JACKPOT" , 2, 33023
'flushdmdtimer.set true , 2000

end sub

'********************* START OF PUPDMD FRAMEWORK v1.0 *************************
'******************** DO NOT MODIFY STUFF BELOW   THIS LINE!!!! ***************
'******************************************************************************
'*****   Create a PUPPack within PUPPackEditor for layout config!!!  **********
'******************************************************************************

Const HasPuP = True

Const pTopper=0
Const pDMD=1
Const pBackglass=2
Const pPlayfield=3
Const pMusic=4
Const pMusic2=5
Const pCallouts=6
Const pBackglass2=7
Const pTopper2=8
Const pPopUP=9
Const pPopUP2=10


'pages
Const pDMDBlank=0
Const pScores=1
Const pBigLine=2
Const pThreeLines=3
Const pTwoLines=4
Const pTargerLetters=5

'dmdType
Const pDMDTypeLCD=0
Const pDMDTypeReal=1
Const pDMDTypeFULL=2






Dim PuPlayer
dim PUPDMDObject  'for realtime mirroring.
Dim pDMDlastchk: pDMDLastchk= -1    'performance of updates
Dim pDMDCurPage: pDMDCurPage= 0     'default page is empty.
Dim pInAttract : pInAttract=false   'pAttract mode




'*************  starts PUP system,  must be called AFTER b2s/controller running so put in last line of table1_init
Sub PuPInit

Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
PuPlayer.B2SInit "", pGameName

if (PuPDMDDriverType=pDMDTypeReal) and (useRealDMDScale=1) Then
       PuPlayer.setScreenEx pDMD,0,0,128,32,0  'if hardware set the dmd to 128,32
End if

PuPlayer.LabelInit pDMD




if PuPDMDDriverType=pDMDTypeReal then

Set PUPDMDObject = CreateObject("PUPDMDControl.DMD")
PUPDMDObject.DMDOpen
PUPDMDObject.DMDPuPMirror
PUPDMDObject.DMDPuPTextMirror
PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 1, ""FN"":33 }"             'set pupdmd for mirror and hide behind other pups
PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 1, ""FN"":32, ""FQ"":3 }"   'set no antialias on font render if real
PuPEvent(800) 'startup pupdmd framework E800 so so an image in case a dmd needs to be pushed back.
END IF



pSetPageLayouts


pDMDSetPage(pDMDBlank)   'set blank text overlay page.
pAttractStart


End Sub 'end PUPINIT



'PinUP Player DMD Helper Functions

Sub pDMDLabelHide(labName)
PuPlayer.LabelSet pDMD,labName,"",0,""
end sub




Sub pDMDScrollBig(msgText,timeSec,mColor)
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,2,timeSec,""
    PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':" & (timeSec*1000000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
  End If
end sub

Sub pDMDScrollBigV(msgText,timeSec,mColor)
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,2,timeSec,""
    PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':2,'yps':1,'ype':-1,'len':" & (timeSec*1000000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
  End If
end sub


Sub pDMDSplashScore(msgText,timeSec,mColor)
  If HasPup Then
    PuPlayer.LabelSet pDMD,"MsgScore",msgText,0,"{'mt':1,'at':1,'fq':250,'len':"& (timeSec*1000) &",'fc':" & mColor & "}"
  End If
end Sub

Sub pDMDSplashScoreScroll(msgText,timeSec,mColor)
  If HasPup Then
    PuPlayer.LabelSet pDMD,"MsgScore",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':"& (timeSec*1000) &", 'mlen':"& (timeSec*1000) &",'tt':0, 'fc':" & mColor & "}"
  End If
end Sub

Sub pDMDZoomBig(msgText,timeSec,mColor)  'new Zoom
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,2,timeSec,""
    PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':3,'hstart':5,'hend':80,'len':" & (timeSec*1000) & ",'mlen':" & (timeSec*500) & ",'tt':5,'fc':" & mColor & "}"
  End If
end sub

Sub pDMDTargetLettersInfo(msgText,msgInfo, timeSec)  'msgInfo = '0211'  0= layer 1, 1=layer 2, 2=top layer3.
'this function is when you want to hilite spelled words.  Like B O N U S but have O S hilited as already hit markers... see example.
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,5,timeSec,""  'show page 5
  End If
Dim backText
Dim middleText
Dim flashText
Dim curChar
Dim i
Dim offchars:offchars=0
Dim spaces:spaces=" "  'set this to 1 or more depends on font space width.  only works with certain fonts
                          'if using a fixed font width then set spaces to just one space.

For i=1 To Len(msgInfo)
    curChar="" & Mid(msgInfo,i,1)
    if curChar="0" Then
            backText=backText & Mid(msgText,i,1)
            middleText=middleText & spaces
            flashText=flashText & spaces
            offchars=offchars+1
    End If
    if curChar="1" Then
            backText=backText & spaces
            middleText=middleText & Mid(msgText,i,1)
            flashText=flashText & spaces
    End If
    if curChar="2" Then
            backText=backText & spaces
            middleText=middleText & spaces
            flashText=flashText & Mid(msgText,i,1)
    End If
Next

if offchars=0 Then 'all litup!... flash entire string
   backText=""
   middleText=""
   FlashText=msgText
end if

If HasPup Then
  PuPlayer.LabelSet pDMD,"Back5"  ,backText  ,1,""
  PuPlayer.LabelSet pDMD,"Middle5",middleText,1,""
  PuPlayer.LabelSet pDMD,"Flash5" ,flashText ,0,"{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) & "}"
End If
end Sub


Sub pDMDSetPage(pagenum)
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,pagenum,0,""   'set page to blank 0 page if want off
    PDMDCurPage=pagenum
  End If
end Sub

Sub pHideOverlayText(pDisp)
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "& pDisp &", ""FN"": 34 }"             'hideoverlay text during next videoplay on DMD auto return
end Sub



Sub pDMDShowLines3(msgText,msgText2,msgText3,timeSec)
  Dim vis:vis=1
  if pLine1Ani<>"" Then vis=0
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,3,timeSec,""
    PuPlayer.LabelSet pDMD,"Splash3a",msgText,vis,pLine1Ani
    PuPlayer.LabelSet pDMD,"Splash3b",msgText2,vis,pLine2Ani
    PuPlayer.LabelSet pDMD,"Splash3c",msgText3,vis,pLine3Ani
  End If
end Sub


Sub pDMDShowLines2(msgText,msgText2,timeSec)
  Dim vis:vis=1
  If pLine1Ani<>"" Then vis=0
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,4,timeSec,""
    PuPlayer.LabelSet pDMD,"Splash4a",msgText,vis,pLine1Ani
    PuPlayer.LabelSet pDMD,"Splash4b",msgText2,vis,pLine2Ani
  End If
end Sub

Sub pDMDShowCounter(msgText,msgText2,msgText3,timeSec)
  Dim vis:vis=1
  If pLine1Ani<>"" Then vis=0
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,6,timeSec,""
    PuPlayer.LabelSet pDMD,"Splash6a",msgText,vis, pLine1Ani
    PuPlayer.LabelSet pDMD,"Splash6b",msgText2,vis,pLine2Ani
    PuPlayer.LabelSet pDMD,"Splash6c",msgText3,vis,pLine3Ani
  End If
end Sub


Sub pDMDShowBig(msgText,timeSec, mColor)
  Dim vis:vis=1
  If pLine1Ani<>"" Then vis=0
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,2,timeSec,""
    PuPlayer.LabelSet pDMD,"Splash",msgText,vis,pLine1Ani
  End If
end sub


Sub pDMDShowHS(msgText,msgText2,msgText3,timeSec) 'High Score
  Dim vis:vis=1
  if pLine1Ani<>"" Then vis=0
  If HasPup Then
    PuPlayer.LabelShowPage pDMD,7,timeSec,""
    PuPlayer.LabelSet pDMD,"Splash7a",msgText,vis,pLine1Ani
    PuPlayer.LabelSet pDMD,"Splash7b",msgText2,vis,pLine2Ani
    PuPlayer.LabelSet pDMD,"Splash7c",msgText3,vis,pLine3Ani
  End If
end Sub


Sub pDMDSetBackFrame(fname)
  If HasPup Then
    PuPlayer.playlistplayex pDMD,"PUPFrames",fname,0,1
  End If
end Sub

Sub pDMDStartBackLoop(fPlayList,fname)
  If HasPup Then
  PuPlayer.playlistplayex pDMD,fPlayList,fname,0,1
  PuPlayer.SetBackGround pDMD,1
  End If
end Sub

Sub pDMDStopBackLoop
  If HasPup Then
    PuPlayer.SetBackGround pDMD,0
    PuPlayer.playstop pDMD
  End If
end Sub


Dim pNumLines

'Theme Colors for Text (not used currenlty,  use the |<colornum> in text labels for colouring.
Dim SpecialInfo
Dim pLine1Color : pLine1Color=8454143
Dim pLine2Color : pLine2Color=8454143
Dim pLine3Color :  pLine3Color=8454143
Dim curLine1Color: curLine1Color=pLine1Color  'can change later
Dim curLine2Color: curLine2Color=pLine2Color  'can change later
Dim curLine3Color: curLine3Color=pLine3Color  'can change later


Dim pDMDCurPriority: pDMDCurPriority =-1
Dim pDMDDefVolume: pDMDDefVolume = 0   'default no audio on pDMD

Dim pLine1
Dim pLine2
Dim pLine3
Dim pLine1Ani
Dim pLine2Ani
Dim pLine3Ani

Dim PriorityReset:PriorityReset=-1
DIM pAttractReset:pAttractReset=-1
DIM pAttractBetween: pAttractBetween=2000 '1 second between calls to next attract page


'************************ where all the MAGIC goes,  pretty much call this everywhere  ****************************************
'*************************                see docs for examples                ************************************************
'****************************************   DONT TOUCH THIS CODE   ************************************************************

Sub pupDMDDisplay(pEventID, pText, VideoName,TimeSec, pAni,pPriority)
' pEventID = reference if application,
' pText = "text to show" separate lines by ^ in same string
' VideoName "gameover.mp4" will play in background  "@gameover.mp4" will play and disable text during gameplay.
' also global variable useDMDVideos=true/false if user wishes only TEXT
' TimeSec how long to display msg in Seconds
' animation if any 0=none 1=Flasher
' also,  now can specify color of each line (when no animation).  "sometext|12345"  will set label to "sometext" and set color to 12345

DIM curPos
if pDMDCurPriority>pPriority then Exit Sub  'if something is being displayed that we don't want interrupted.  same level will interrupt.
pDMDCurPriority=pPriority
if timeSec=0 then timeSec=1 'don't allow page default page by accident


pLine1=""
pLine2=""
pLine3=""
pLine1Ani=""
pLine2Ani=""
pLine3Ani=""


if pAni=1 Then  'we flashy now aren't we
pLine1Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"
pLine2Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"
pLine3Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"
end If

curPos=InStr(pText,"^")   'Lets break apart the string if needed
if curPos>0 Then
   pLine1=Left(pText,curPos-1)
   pText=Right(pText,Len(pText) - curPos)

   curPos=InStr(pText,"^")   'Lets break apart the string
   if curPOS>0 Then
      pLine2=Left(pText,curPos-1)
      pText=Right(pText,Len(pText) - curPos)

      curPos=InStr("^",pText)   'Lets break apart the string
      if curPos>0 Then
         pline3=Left(pText,curPos-1)
      Else
        if pText<>"" Then pline3=pText
      End if
   Else
      if pText<>"" Then pLine2=pText
   End if
Else
  pLine1=pText  'just one line with no break
End if


'lets see how many lines to Show
pNumLines=0
if pLine1<>"" then pNumLines=pNumlines+1
if pLine2<>"" then pNumLines=pNumlines+1
if pLine3<>"" then pNumLines=pNumlines+1


if (VideoName<>"") and (useDMDVideos) Then  'we are showing a splash video instead of the text.
  If HasPup Then
    PuPlayer.playlistplayex pDMD,"DMDSplash",VideoName,pDMDDefVolume,pPriority  'should be an attract background (no text is displayed)
    End If
end if 'if showing a splash video with no text




if StrComp(pEventID,"shownum",1)=0 Then              'check eventIDs
    pDMDShowCounter pLine1,pLine2,pLine3,timeSec
Elseif StrComp(pEventID,"target",1)=0 Then              'check eventIDs
    pDMDTargetLettersInfo pLine1,pLine2,timeSec
Elseif StrComp(pEventID,"highscore",1)=0 Then              'check eventIDs
    pDMDShowHS pLine1,pLine2,pline3,timeSec
Elseif (pNumLines=3) Then                'depends on # of lines which one to use.  pAni=1 will flash.
    pDMDShowLines3 pLine1,pLine2,pLine3,TimeSec
Elseif (pNumLines=2) Then
    pDMDShowLines2 pLine1,pLine2,TimeSec
Elseif (pNumLines=1) Then
    pDMDShowBig pLine1,timeSec, curLine1Color
Else
    pDMDShowBig pLine1,timeSec, curLine1Color
End if

PriorityReset=TimeSec*1000
End Sub 'pupDMDDisplay message



Sub pupDMDupdate_Timer()
  pUpdateScores

    if PriorityReset>0 Then  'for splashes we need to reset current prioirty on timer
       PriorityReset=PriorityReset-pupDMDUpdate.interval
       if PriorityReset<=0 Then
            pDMDCurPriority=-1
            if pInAttract then pAttractReset=pAttractBetween ' pAttractNext  call attract next after 1 second
      End if
    End if

    if pAttractReset>0 Then  'for splashes we need to reset current prioirty on timer
       pAttractReset=pAttractReset-pupDMDUpdate.interval
       if pAttractReset<=0 Then
            pAttractReset=-1
            if pInAttract then pAttractNext
      End if
    end if
End Sub


Sub PuPEvent(EventNum)
if hasPUP=false then Exit Sub
PuPlayer.B2SData "E"&EventNum,1  'send event to puppack driver
End Sub

'********************* END OF PUPDMD FRAMEWORK v1.0 *************************
'******************** DO NOT MODIFY STUFF ABOVE THIS LINE!!!! ***************
'****************************************************************************

'*****************************************************************
'   **********  PUPDMD  MODIFY THIS SECTION!!!  ***************
'PUPDMD Layout for each Table1
'Setup Pages.  Note if you use fonts they must be in FONTS folder of the pupVideos\tablename\FONTS  "case sensitive exact naming fonts!"
'*****************************************************************

Sub pSetPageLayouts

DIM dmddef
DIM dmdalt
DIM dmdscr
DIM dmdfixed

'labelNew <screen#>, <Labelname>, <fontName>,<size%>,<colour>,<rotation>,<xalign>,<yalign>,<xpos>,<ypos>,<PageNum>,<visible>
'***********************************************************************'
'<screen#>, in standard wed set this to pDMD ( or 1)
'<Labelname>, your name of the label. keep it short no spaces (like 8 chars) although you can call it anything really. When setting the label you will use this labelname to access the label.
'<fontName> Windows font name, this must be exact match of OS front name. if you are using custom TTF fonts then double check the name of font names.
'<size%>, Height as a percent of display height. 20=20% of screen height.
'<colour>, integer value of windows color.
'<rotation>, degrees in tenths   (900=90 degrees)
'<xAlign>, 0= horizontal left align, 1 = center horizontal, 2= right horizontal
'<yAlign>, 0 = top, 1 = center, 2=bottom vertical alignment
'<xpos>, this should be 0, but if you want to force a position you can set this. it is a % of horizontal width. 20=20% of screen width.
'<ypos> same as xpos.
'<PageNum> IMPORTANT this will assign this label to this page or group.
'<visible> initial state of label. visible=1 show, 0 = off.



if PuPDMDDriverType=pDMDTypeReal Then 'using RealDMD Mirroring.  **********  128x32 Real Color DMD
  dmdalt="PKMN Pinball"
    dmdfixed="Instruction"
    dmdscr="Impact"    'main scorefont
  dmddef="Zig"

  'Page 1 (default score display)
       PuPlayer.LabelNew pDMD,"Credits" ,dmddef,20,33023   ,0,2,2,95,0,1,0
     PuPlayer.LabelNew pDMD,"Play1"   ,dmdalt,21,33023   ,1,0,0,15,0,1,0
     PuPlayer.LabelNew pDMD,"Ball"    ,dmdalt,21,33023   ,1,2,0,85,0,1,0
     PuPlayer.LabelNew pDMD,"MsgScore",dmddef,20,33023   ,0,1,0, 0,40,1,0
     PuPlayer.LabelNew pDMD,"CurScore",dmdscr,60,8454143   ,0,1,1, 0,0,1,1


  'Page 2 (default Text Splash 1 Big Line)
     PuPlayer.LabelNew pDMD,"Splash"  ,dmdalt,40,33023,0,1,1,0,0,2,0

  'Page 3 (default Text Splash 2 and 3 Lines)
     PuPlayer.LabelNew pDMD,"Splash3a",dmddef,24,8454143,0,1,0,0,2,3,0
     PuPlayer.LabelNew pDMD,"Splash3b",dmdalt,30,33023,0,1,0,0,30,3,0
       PuPlayer.LabelNew pDMD,"Splash3c",dmdalt,25,33023,0,1,0,0,55,3,0


  'Page 4 (2 Line Gameplay DMD)
     PuPlayer.LabelNew pDMD,"Splash4a",dmddef,24,8454143,0,1,0,0,0,4,0
       PuPlayer.LabelNew pDMD,"Splash4b",dmddef,30,33023,0,1,2,0,75,4,0

  'Page 5 (3 layer large text for overlay targets function,  must you fixed width font!
    PuPlayer.LabelNew pDMD,"Back5"    ,dmdfixed,50,8421504,0,1,1,0,0,5,0
    PuPlayer.LabelNew pDMD,"Middle5"  ,dmdfixed,50,65535  ,0,1,1,0,0,5,0
    PuPlayer.LabelNew pDMD,"Flash5"   ,dmdfixed,50,65535  ,0,1,1,0,0,5,0

  'Page 6 (3 Lines for big # with two lines,  "19^Orbits^Count")
    PuPlayer.LabelNew pDMD,"Splash6a",dmddef,50,65280,0,0,1,1,0,6,0
    PuPlayer.LabelNew pDMD,"Splash6b",dmddef,20,33023,0,1,0,60,0,6,0
    PuPlayer.LabelNew pDMD,"Splash6c",dmddef,20,33023,0,1,0,60,50,6,0

  'Page 7 (Show High Scores Fixed Fonts)
    PuPlayer.LabelNew pDMD,"Splash7a",dmddef,20,8454143,0,1,0,0,2,7,0
    PuPlayer.LabelNew pDMD,"Splash7b",dmdalt,40,33023,0,1,0,0,20,7,0
    PuPlayer.LabelNew pDMD,"Splash7c",dmdalt,40,33023,0,1,0,0,50,7,0

END IF  ' use PuPDMDDriver

if PuPDMDDriverType=pDMDTypeLCD THEN  'Using 4:1 Standard ratio LCD PuPDMD  ************ lcd **************

  'dmddef="Impact"
  dmdalt="PKMN Pinball"
    dmdfixed="Instruction"
  dmdscr="Impact"  'main score font
  dmddef="Impact"

  'Page 1 (default score display)
    PuPlayer.LabelNew pDMD,"Credits" ,dmddef,20,33023   ,0,2,2,95,0,1,0
    PuPlayer.LabelNew pDMD,"Play1"   ,dmdalt,20,33023   ,1,0,0,15,0,1,0
    PuPlayer.LabelNew pDMD,"Ball"    ,dmdalt,20,33023   ,1,2,0,85,0,1,0
    PuPlayer.LabelNew pDMD,"MsgScore",dmddef,45,33023   ,0,1,0, 0,40,1,0
    PuPlayer.LabelNew pDMD,"CurScore",dmdscr,60,8454143   ,0,1,1, 0,0,1,0


  'Page 2 (default Text Splash 1 Big Line)
    PuPlayer.LabelNew pDMD,"Splash"  ,dmdalt,40,33023,0,1,1,0,0,2,0

  'Page 3 (default Text 3 Lines)
    PuPlayer.LabelNew pDMD,"Splash3a",dmddef,30,8454143,0,1,0,0,2,3,0
    PuPlayer.LabelNew pDMD,"Splash3b",dmdalt,30,33023,0,1,0,0,30,3,0
    PuPlayer.LabelNew pDMD,"Splash3c",dmdalt,25,33023,0,1,0,0,57,3,0


  'Page 4 (default Text 2 Line)
    PuPlayer.LabelNew pDMD,"Splash4a",dmddef,40,8454143,0,1,0,0,0,4,0
    PuPlayer.LabelNew pDMD,"Splash4b",dmddef,30,33023,0,1,2,0,75,4,0

  'Page 5 (3 layer large text for overlay targets function,  must you fixed width font!
    PuPlayer.LabelNew pDMD,"Back5"    ,dmdfixed,80,8421504,0,1,1,0,0,5,0
    PuPlayer.LabelNew pDMD,"Middle5"  ,dmdfixed,80,65535  ,0,1,1,0,0,5,0
    PuPlayer.LabelNew pDMD,"Flash5"   ,dmdfixed,80,65535  ,0,1,1,0,0,5,0

  'Page 6 (3 Lines for big # with two lines,  "19^Orbits^Count")
    PuPlayer.LabelNew pDMD,"Splash6a",dmddef,90,65280,0,0,0,15,1,6,0
    PuPlayer.LabelNew pDMD,"Splash6b",dmddef,40,33023,0,1,0,60,0,6,0
    PuPlayer.LabelNew pDMD,"Splash6c",dmddef,40,33023,0,1,0,60,50,6,0

  'Page 7 (Show High Scores Fixed Fonts)
    PuPlayer.LabelNew pDMD,"Splash7a",dmddef,20,8454143,0,1,0,0,2,7,0
    PuPlayer.LabelNew pDMD,"Splash7b",dmdfixed,40,33023,0,1,0,0,20,7,0
    PuPlayer.LabelNew pDMD,"Splash7c",dmdfixed,40,33023,0,1,0,0,50,7,0


END IF  ' use PuPDMDDriver

if PuPDMDDriverType=pDMDTypeFULL THEN  'Using FULL BIG LCD PuPDMD  ************ lcd **************

  'dmddef="Impact"
  dmdalt="PKMN Pinball"
    dmdfixed="Instruction"
  dmdscr="Impact"  'main score font
  dmddef="Impact"

  'Page 1 (default score display)
    PuPlayer.LabelNew pDMD,"Credits" ,dmddef,20,33023   ,0,2,2,95,0,1,0
    PuPlayer.LabelNew pDMD,"Play1"   ,dmdalt,20,33023   ,1,0,0,15,0,1,0
    PuPlayer.LabelNew pDMD,"Ball"    ,dmdalt,20,33023   ,1,2,0,85,0,1,0
    PuPlayer.LabelNew pDMD,"MsgScore",dmddef,45,33023   ,0,1,0, 0,40,1,0
    PuPlayer.LabelNew pDMD,"CurScore",dmdscr,60,8454143   ,0,1,1, 0,0,1,0
    PuPlayer.LabelNew pDMD,"CurImage",dmddef,50,33023 ,0,1,1,0, 0,1,1  'new image type


  'Page 2 (default Text Splash 1 Big Line)
    PuPlayer.LabelNew pDMD,"Splash"  ,dmdalt,40,33023,0,1,1,0,0,2,0

  'Page 3 (default Text 3 Lines)
    PuPlayer.LabelNew pDMD,"Splash3a",dmddef,30,8454143,0,1,0,0,2,3,0
    PuPlayer.LabelNew pDMD,"Splash3b",dmdalt,30,33023,0,1,0,0,30,3,0
    PuPlayer.LabelNew pDMD,"Splash3c",dmdalt,25,33023,0,1,0,0,57,3,0


  'Page 4 (default Text 2 Line)
    PuPlayer.LabelNew pDMD,"Splash4a",dmddef,40,8454143,0,1,0,0,0,4,0
    PuPlayer.LabelNew pDMD,"Splash4b",dmddef,30,33023,0,1,2,0,75,4,0

  'Page 5 (3 layer large text for overlay targets function,  must you fixed width font!
    PuPlayer.LabelNew pDMD,"Back5"    ,dmdfixed,80,8421504,0,1,1,0,0,5,0
    PuPlayer.LabelNew pDMD,"Middle5"  ,dmdfixed,80,65535  ,0,1,1,0,0,5,0
    PuPlayer.LabelNew pDMD,"Flash5"   ,dmdfixed,80,65535  ,0,1,1,0,0,5,0

  'Page 6 (3 Lines for big # with two lines,  "19^Orbits^Count")
    PuPlayer.LabelNew pDMD,"Splash6a",dmddef,90,65280,0,0,0,15,1,6,0
    PuPlayer.LabelNew pDMD,"Splash6b",dmddef,40,33023,0,1,0,60,0,6,0
    PuPlayer.LabelNew pDMD,"Splash6c",dmddef,40,33023,0,1,0,60,50,6,0

  'Page 7 (Show High Scores Fixed Fonts)
    PuPlayer.LabelNew pDMD,"Splash7a",dmddef,20,8454143,0,1,0,0,2,7,0
    PuPlayer.LabelNew pDMD,"Splash7b",dmdfixed,40,33023,0,1,0,0,20,7,0
    PuPlayer.LabelNew pDMD,"Splash7c",dmdfixed,40,33023,0,1,0,0,50,7,0


END IF  ' use PuPDMDDriver




end Sub 'page Layouts


'*****************************************************************
'        PUPDMD Custom SUBS/Events for each Table1
'     **********    MODIFY THIS SECTION!!!  ***************
'*****************************************************************
'
'
'  we need to somewhere in code if applicable
'
'   call pDMDStartGame,pDMDStartBall,pGameOver,pAttractStart
'
'
'
'
'


Sub pDMDStartGame
pInAttract=false
pDMDSetPage(pScores)   'set blank text overlay page.

end Sub


Sub pDMDStartBall
pDMDSetPage(pScores)
end Sub

Sub pDMDGameOver
pAttractStart
end Sub

Sub pAttractStart
'pupDMDDisplay "attract","Welcome^To Goonies","@welcome.mp4",5,1,10
if pInAttract then Exit SUB
pDMDSetPage(pDMDBlank)
pCurAttractPos=0 'reset position
pInAttract=True 'Startup in AttractMode
pAttractReset=1000  'start attract timer in 1 second

end Sub

DIM pCurAttractPos: pCurAttractPos=0


'********************** gets called auto each page next and timed already in DMD_Timer.  make sure you use pupDMDDisplay or it wont advance auto.
Sub pAttractNext
if pInAttract=false Then exit SUB
pCurAttractPos=pCurAttractPos+1

  Select Case pCurAttractPos

  Case 1
       pupDMDDisplay "attract","Welcome to^The Goonies", "@welcome.mp4",3,0,10
  Case 2
       pupDMDDisplay "attract", "The^Goonies", "@GooniesLogo.mp4", 3, 0, 10
  Case 3 pupDMDDisplay "attract","REPLAY AT^7,500,000","",3,1,10

  Case 4 pupDMDDisplay "attract","HIGHSCORES", "",2,0,10
  Case 5 pupDMDDisplay "highscore","High Scores^1> " & HighScoreName(0) & "  " & HighScore(0)&"^2> " & HighScoreName(1) & "  " & HighScore(1) , "", 3, 0, 10
  Case 6 pupDMDDisplay "highscore","High Scores^3> " & HighScoreName(2) & "  " & HighScore(2)&"^4> " & HighScoreName(3) & "  " & HighScore(3) , "", 3, 0, 10
  Case 7
    If score(currentplayer) > 0 Then
       pupDMDDisplay "GAMEOVER", "GAME OVER^Last Score "&(score(currentplayer)), "", 3, 1, 10
       pupDMDDisplay "GAMEOVER", "GAME OVER^Last Score "&(score(currentplayer)), "", 3, 1, 10
       Else
       pupDMDDisplay "GAMEOVER", "GAME OVER", "@gameover.mp4", 3, 1, 10
    end if
  Case 8 pupDMDDisplay "attract","CREDITS","@authorcredits.mp4",10,0,10
  Case 9 pupDMDDisplay "attract","WINNERS DON'T^USE DRUGS","",3,0,10
  Case 10
     if Credits = 0 then
       pupDMDDisplay "attract", "CREDITS 0^INSERT COIN", "", 3, 0, 10
       Else
       pupDMDDisplay "attract", "CREDITS "&(Credits)&"^PRESS START", "", 3, 0, 10
    End If
  Case 11 pupDMDDisplay "attract","VPX TABLE^BY JAVIER","",3,0,10
  Case Else
    pCurAttractPos=0
    pAttractNext 'reset to beginning
  end Select
'note if you want flipper keys to advance PriorityReset=1 will do it.
end Sub


'************************ called during gameplay to update Scores ***************************
Sub pUpdateScores
if pDMDCurPage <> pScores then Exit Sub
puPlayer.LabelSet pDMD,"CurScore","" & FormatNumber(Score(CurrentPlayer),0) ,1,""
'puPlayer.LabelSet pDMD,"Credits","CREDITS " & ""& Credits ,1,""
puPlayer.LabelSet pDMD,"Play1","Player " & CurrentPlayer,1,""
puPlayer.LabelSet pDMD,"Ball","Ball " & ""& balls ,1,""
end Sub






'backward compatiblity with old methods. to new pupdmd framework.
Sub pDMDSplashLines(msgText,msgText2,timeSec,Ani)
  if msgText="" Then msgText=" "
  if msgText2="" Then msgText2="  "
   pupDMDDisplay "default", msgText&"^"&msgText2, "", timesec, Ani, 10
end Sub

SUB pDMDSplashBonus(msgText,msgText2,timeSec,Ani)
if msgText="" Then msgText=" "
if msgText2="" Then msgText2="  "
pupDMDDisplay "default", msgText&"^"&msgText2, "", timesec, Ani, 10
end Sub


SUB pDMDSplash3Lines(msgText,msgText2,msgText3,timeSec,Ani)
if msgText="" Then msgText=" "
if msgText2="" Then msgText2="  "
if msgText3="" Then msgText3="  "
pupDMDDisplay "default", msgText&"^"&msgText2&"^"&msgText3, "", timesec, Ani, 10
end Sub


SUB pDMDSplashBig(msgText,timeSec,Ani)
pupDMDDisplay "default",msgText, "", timesec, Ani, 10
end Sub

SUB pUpdateMystery(msgText,msgText2,msgText3,timeSec,Ani)
  if msgText="" Then msgText=" "
  if msgText2="" Then msgText2="  "
  if msgText3="" Then msgText3="  "
  If HasPup Then
    PuPlayer.LabelSet pDMD,"Splash3a",msgText,1,""
    PuPlayer.LabelSet pDMD,"Splash3b",msgText2,1,""
    PuPlayer.LabelSet pDMD,"Splash3c",msgText3,1,""
  End If
end Sub

'Sub pupDMDDisplay(pEventID, pText, VideoName,TimeSec, pAni,pPriority)
' pEventID = reference if application,
' pText = "text to show" separate lines by ^ in same string
' VideoName "gameover.mp4" will play in background  "@gameover.mp4" will play and disable text during gameplay.
' also global variable useDMDVideos=true/false if user wishes only TEXT
' TimeSec how long to display msg in Seconds
' animation if any 0=none 1=Flasher
' also,  now can specify color of each line (when no animation).  "sometext|12345"  will set label to "sometext" and set color to 12345
'Samples
'pupDMDDisplay "balllock", "Ball^Locked|16744448", "", 5, 1, 10             '  5 seconds,  1=flash, 10=priority, ball is first line, locked on second and locked has custom color |
'pupDMDDisplay "balllock","Ball 2^is^Locked", "balllocked2.mp4",3, 1,10     '  3 seconds,  1=flash, play balllocked2.mp4 from dmdsplash folder,
'pupDMDDisplay "balllock","Ball^is^Locked", "@balllocked.mp4",3, 1,10       '  3 seconds,  1=flash, play @balllocked.mp4 from dmdsplash folder, because @ text by default is hidden unless useDmDvideos is disabled.
'pupDMDDisplay "shownum", "3^More To|616744448^GOOOO", "", 5, 1, 10         ' "shownum" is special.  layout is line1=BIG NUMBER and line2,line3 are side two lines.  "4^Ramps^Left"
'pupDMDDisplay "target", "POTTER^110120", "blank.mp4", 10, 0, 10            ' 'target'...  first string is line,  second is 0=off,1=already on, 2=flash on for each character in line (count must match)
'pupDMDDisplay "highscore", "High Score^AAA   2451654^BBB   2342342", "", 5, 0, 10            ' highscore is special  line1=text title like highscore, line2, line3 are fixed fonts to show AAA 123,123,123
'pupDMDDisplay "highscore", "High Score^AAA   2451654|616744448^BBB   2342342", "", 5, 0, 10  ' sames as above but notice how we use a custom color for text |



' nFozzy - Start


dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

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

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub


'******************************************************
'                        FLIPPER CORRECTION FUNCTIONS
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
'                FLIPPER POLARITY AND RUBBER DAMPENER
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
'                        FLIPPER TRICKS
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
' Maths
'*****************
Const PI = 3.1415927

Function dSin(degrees)
        dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
        dcos = cos(degrees * Pi/180)
End Function

'*************************************************
' Check ball distance from Flipper for Rem
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

' Used for drop targets and stand up targets
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

'*************************************************
' End - Check ball distance from Flipper for Rem
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
Const EOSTnew = 0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025

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



'****************************************************************************
'PHYSICS DAMPENERS
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

' Thalamus - patched : ' Thalamus - patched :                 aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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
'                TRACK ALL BALL VELOCITIES
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

Sub RDampen_Timer()
  Cor.Update
End Sub


'***** Fleep Start

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
'                     Fleep  Supporting Ball & Sound Functions
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


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 18 ' total number of balls
Const lob = 4 ' number of enclosed balls
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
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = lob - 1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = lob to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
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


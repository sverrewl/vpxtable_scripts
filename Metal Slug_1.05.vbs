' *********************************************************************
' **                                                                 **
' **                        FUTURE PINBALL                           **
' **                       Metal Slug  V2.0                          **
' **                                                                 **
' **                2006 Brendan Bailey (Pinwizkid)                  **
' **           Update by Dominique Camus (DoCam)                     **
' **                                                                 **
' **           Last update: 2.0  - November 1st, 2009                **
' **                                                                 **
' *********************************************************************


' *** WHAT'S NEW IN 2.0 BETA ? ***
'
' - Some textures are now correct (for FP 1.9), using TGA with alpha channel for transparency (instead of BMP), resized to "power of 2".
' - Reworked kickback with an invisible moveable guide, more reliable. Increased delay to avoid FP bug (ball outside cabinet !).
' - Updated delay for auto plunger.
' - Reworked displayed messages, formatted scores (comma as separator), text alignments, etc.
' - Added feature: table setup. Press & hold both flipper keys during attract mode to enter setup. You can set number of balls
'   (3/5/7/9), ball saver duration (15/30/45/60 sec), POW bonus value, how POW bonus is held or not, Jackpot value and tilt sensitivity.
'   You can reset to default (factory) settings only (as defined by Brendan in v1.0: 5 balls, 15 sec. ball saver, POW bonus = 10,000,
'   Jackpot = 300,000, tilt sensitivity set as medium = 2 warnings), high score table only, or both settings/high scores.
' - Using left or right flip key during attract advances quickly to the next message.
' - Added random POW names (got from real Metal Slug Super Vehicle-001) and ranks (when you get a POW bonus).
' - Unlit missions (1/2/3/4/5/!) at the end of final mission.
' - Reworked ball saver (normal blink, then fast blink during last 5 seconds), grace delay when hitting outlane trigger.
' - Fixed ball saver reactivation when the current played ball returns to the plunger lane.
' - Reworked the "ball saver" for 5-ball multiball during final mission.
' - Reworked the bonus countdown sequence, character sound moved to other unusued channel.
' - Boredom effect set to 0 (means disabled) for both backglass and 'HUD' displays.
' - Added myself (Dominique Camus a.k.a. DoCam) in special thanks section during attract mode (Brendan's request :-) Thanks !
' - Reworked tilt subroutine.
'
' *** TODO ***
'
' - Cleaning VB code, doing script and variables optimizations, adding more comments ASAP.
' - Timer optimizations (remove unused, multipurpose if possible by using .UserData timer property).
' - Reworking entirely the display (bad boredom effects, message priorities).
' - Explosion textures (JPG) conversion to "power of 2" size.



' *********************************************************************
' **                                                                 **
' **  VARIABLES & CONSTANTS DECLARATIONS                             **
' **                                                                 **
' *********************************************************************

Option Explicit
Randomize


' Thalamus - converted from wmv to mp4 - Converted to FlexDMD
' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Its a original - move sounds not directional referenced to backglass !
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
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolFlip   = 1    ' Flipper volume.


'---------- UltraDMD Unique Table Color preference -------------
Dim DMDColor, DMDColorSelect, UseFullColor
Dim DMDPosition, DMDPosX, DMDPosY, DMDSize, DMDWidth, DMDHeight


UseFullColor = "True" '                           "True" / "False"
DMDColorSelect = "Red"            ' Rightclick on UDMD window to get full list of colours

DMDPosition = False                               ' Use Manual DMD Position, True / False
DMDPosX = 100                                   ' Position in Decimal
DMDPosY = 40                                     ' Position in Decimal

DMDSize = False                                     ' Use Manual DMD Size, True / False
DMDWidth = 512                                    ' Width in Decimal
DMDHeight = 128                                   ' Height in Decimal

'Note open Ultradmd and right click on window to get the various sizes in decimal

GetDMDColor
Sub GetDMDColor
Dim WshShell,filecheck,directory
Set WshShell = CreateObject("WScript.Shell")
If DMDSize then
WshShell.RegWrite "HKCU\Software\UltraDMD\w",DMDWidth,"REG_DWORD"
WshShell.RegWrite "HKCU\Software\UltraDMD\h",DMDHeight,"REG_DWORD"
End if
If DMDPosition then
WshShell.RegWrite "HKCU\Software\UltraDMD\x",DMDPosX,"REG_DWORD"
WshShell.RegWrite "HKCU\Software\UltraDMD\y",DMDPosY,"REG_DWORD"
End if
WshShell.RegWrite "HKCU\Software\UltraDMD\fullcolor",UseFullColor,"REG_SZ"
WshShell.RegWrite "HKCU\Software\UltraDMD\color",DMDColorSelect,"REG_SZ"
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' //////////////////////
 ' B2S Light Show
 ' cause i mean everyone loves a good light show
 ' 1 = Background
 ' 2 = Tank
 ' 3 = Marco
 ' 4 = Gun
 ' 5 = Logo
 ' 6 = All Lights
 ' /////////////////////
 ' example B2S call
 ' startB2S(#)

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

' /////////////////
' END b2s
' well that was quick, tacos anyone?
' /////////////////
Const BallSize = 50
Const BallMass = 1.2

ExecuteGlobal GetTextFile("FPVPX.vbs")
If Err Then MsgBox "you need the fpvpx.vbs for the proper functioning of the table"

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

Const cGameName = "metalslug"


if Table1.ShowDT Then
   D1.visible = 1
 Else
   D1.visible = 0
End If

' Define any Constants
Const TableName = "Metal Slug"
Const myVersion = "1.0.0"
Const MaxPlayers = 4     ' from 1 to 4
Const BallSaverTime = 15 ' in seconds
Const MaxMultiplier = 5  ' limit to 5x in this game, both bonus multiplier and playfield multiplier
Const BallsPerGame = 3   ' usually 3 or 5
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
Dim HighScore(10)
Dim HighScoreName(10)
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





' Define global arrays (for table setup).
'
Const stMaxParam = 8    ' Setup: number of parameters.
Dim stParam(8)        ' Setup: parameter names.
Dim stNumberValues(8) ' Setup: number of possible value for each parameter.
Dim stValue(8, 6)     ' Setup: list of values (for each parameter).
Dim stCurrentValue(8) ' Setup: current selected (or default) value for each parameter.



' Define global arrays (for POW names and ranks).
'
Dim POWName(100)      ' Prisoner names (got from Metal Slug Super Vehicle-001, NeoGeo version).
Dim POWRank(12)     ' Prisoner ranks (from PVT to GEN).



' Initialize POW names array (will be randomly shuffled).
'
POWName(1) = "HYMENEUS" : POWName(2) = "FUKUMINI" : POWName(3) = "MARSYAS" : POWName(4) = "PLUTO" : POWName(5) = "FABRE"
POWName(6) = "JASON" : POWName(7) = "LAOCOON" : POWName(8) = "GARFIELD" : POWName(9) = "HABARA" : POWName(10) = "PRIAPUS"
POWName(11) = "MERCURY" : POWName(12) = "SILVANUS" : POWName(13) = "OEDIPUS" : POWName(14) = "CHRISTIAN" : POWName(15) = "OSBORN"
POWName(16) = "SATYR" : POWName(17) = "HECTOR" : POWName(18) = "THETIS" : POWName(19) = "AKIMOTO" : POWName(20) = "WOEIKOV"
POWName(21) = "FREDERIK" : POWName(22) = "JAMES" : POWName(23) = "PHOEBE" : POWName(24) = "BIGPOWER" : POWName(25) = "SARPEDON"
POWName(26) = "EASTWOOD" : POWName(27) = "URANFF" : POWName(28) = "WITTFOGEL" : POWName(29) = "RICCI" : POWName(30) = "WATSUK"
POWName(31) = "OKUMA" : POWName(32) = "CADMUS" : POWName(33) = "PHILEMON" : POWName(34) = "GANYMEDE" : POWName(35) = "ACHATES"
POWName(36) = "GRIFFIS" : POWName(37) = "KHUNG" : POWName(38) = "ICARUS" : POWName(39) = "MIDAS" : POWName(40) = "ALMAGRO"
POWName(41) = "SANTAYANA" : POWName(42) = "MAUBONT" : POWName(43) = "KARIBE" : POWName(44) = "SUZUKI" : POWName(45) = "BUCHHOLTZ"
POWName(46) = "GOODYEAR" : POWName(47) = "FORTUNA" : POWName(48) = "POLLUX" : POWName(49) = "KARIMEN" : POWName(50) = "APPLETON"
POWName(51) = "POPLIN" : POWName(52) = "BAILEY" : POWName(53) = "KRUPSKAIA" : POWName(54) = "HOWOOD" : POWName(55) = "KENDALL"
POWName(56) = "HANEDA" : POWName(57) = "BIROLINEN" : POWName(58) = "WATTS" : POWName(59) = "PALES" : POWName(60) = "VIGNEAUD"
POWName(61) = "QUEZON" : POWName(62) = "PERSEUS" : POWName(63) = "COOPER" : POWName(64) = "SANDE" : POWName(65) = "ALARCON"
POWName(66) = "DAIMYO" : POWName(67) = "EPIGONI" : POWName(68) = "ALCMANEON" : POWName(69) = "NYANDABA" : POWName(70) = "YOSHII"
POWName(71) = "ANCHISES" : POWName(72) = "SEMIRAMIS" : POWName(73) = "LANDE" : POWName(74) = "SHIINE" : POWName(75) = "COOLEY"
POWName(76) = "ORIKASA" : POWName(77) = "BORODIN" : POWName(78) = "VALCAN" : POWName(79) = "ALDRINGEN" : POWName(80) = "GULIEH"
POWName(81) = "ARCHANGEL" : POWName(82) = "PUSHKIN" : POWName(83) = "TAUNA" : POWName(84) = "MURRAI" : POWName(85) = "NEMESIS"
POWName(86) = "DOROTHY" : POWName(87) = "JACKSON" : POWName(88) = "HOLMES" : POWName(89) = "NYMPHUS" : POWName(90) = "MEDUSA"
POWName(91) = "MORGA" : POWName(92) = "KOYAMA" : POWName(93) = "BONNET" : POWName(94) = "KOUROGI" : POWName(95) = "KIRIYAMA"
POWName(96) = "CENTAUR" : POWName(97) = "VESPUCCI" : POWName(98) = "ARACHNE" : POWName(99) = "PELEUS" : POWName(100) = "INOMATA"



' Initialize POW ranks array.
'
POWRank(1) = "PVT" : POWRank(2) = "PFC" : POWRank(3) = "CPL" : POWRank(4) = "SGT" : POWRank(5) = "FSG" : POWRank(6) = "SMA"
POWRank(7) = "2LT" : POWRank(8) = "LT" : POWRank(9) = "CPT" : POWRank(10) = "LCL" : POWRank(11) = "COL" : POWRank(12) = "GEN"



' Define global variables.
'
Dim stCurrentParam      ' Setup: current parameter (actually displayed).
Dim Ball              ' Ball in play.
'Dim BallsOnPlayfield     ' Number of balls on playfield, including locked balls.
Dim BallsLocked       ' Number of balls locked in kickers.
Dim OutLaneSound        ' Pointer to sound when a ball hit an inlane/outlane (three possible sounds).
Dim POWBonus          ' POW bonus collected during each ball or since begin of game.
Dim POWBonusBall        ' POW bonus collected during each ball (always resetted).
Dim POWBonusCount       ' POW bonus (used for bonus countdown sequence).
Dim POWNamePtr          ' Pointer to POW name array (from 1 to 100).
Dim TempState         ' This temporary variable is used for POW lane light rotations (when pressing left or right flipper key).
Dim WeaponTargetsMove   ' Pointer to move WEAPON targets while playing missions 2 and 5.
Dim SlugSound         ' Pointer to next sound to be played on SLUG target hit.
Dim RemainingJets       ' Remaining jets, while playing mission 3.
Dim RemainingPOW        ' Remaining POW while playing mission 4.
Dim RemainingAllen      ' Remaining Allen's energy, during mission 5.
Dim CurrentWeapon       ' Pointer to current weapon.
Dim Mission           ' Current (or next) mission.
Dim NextMessage       ' Pointer to next message to be displayed.



' Define game flags - also known as boolean (TRUE/FALSE) variables.
'
Dim bPlayingBall        ' TRUE while the player is playing ball (useful for display)
Dim bLeftInLaneIsLit      ' Left inlane light (TRUE means lit).
Dim bRightInLaneIsLit   ' Right inlane light (TRUE means lit).
Dim bBallSaverIsActive    ' TRUE if the ball saver is active.
Dim bCanActivateBallSaver  ' TRUE if ball saver can be activated (only at the begin of each ball)
Dim bExtraBallAward     ' TRUE if an extra ball is awarded.
Dim bKickbackIsLit      ' Kickback light (TRUE means lit).
'Dim bBallInPlungerLane   ' TRUE if a ball is in the plunger lane.
Dim bBallOnLeftOutLaneTrigger ' TRUE while a ball press the outlane trigger.
Dim bLockIsLit          ' TRUE if a ball can be locked in kicker.
Dim bBallInKicker1      ' TRUE if a ball is in kicker 1 (left side kicker).
Dim bBallLockedInKicker1  ' TRUE if a ball is locked in kicker 1 (left side kicker).
Dim bBallInKicker2      ' TRUE if a ball is in kicker 2 (right side kicker).
Dim bBallLockedInKicker2  ' TRUE if a ball is locked in kicker 2 (right side kicker).
Dim bBallInKicker3      ' TRUE if a ball is in kicker 3 (center kicker, near bumpers).
Dim bBallLockedInKicker3  ' TRUE if a ball is locked in kicker 3 (center kicker, near bumpers).
Dim bBallInKicker4      ' TRUE if a ball is in kicker 4 (top left kicker).
Dim bBallLockedInKicker4  ' TRUE if a ball is locked in kicker 4 (top left kicker).
'Dim bMultiBallMode     ' Is a multiball mode active ?
Dim bJackpotIsLit       ' TRUE if JACKPOT is lit.
Dim bJackpotScored      ' TRUE if JACKPOT already scored.
Dim bW1TargetIsDown     ' WEAPON drop targets states (used to fix after mission).
Dim bW2TargetIsDown
Dim bW3TargetIsDown
Dim bW4TargetIsDown
Dim bW5TargetIsDown
Dim bW6TargetIsDown
Dim bWallIsDown       ' TRUE when all WEAPON drop targets (wall) is down.
Dim bMissionIsLit       ' TRUE if the next mission can be started.
Dim bFinalMissionIsReady  ' TRUE if the final mission is ready.
Dim bPlayingMission     ' TRUE when a mission is playing.
Dim bPlayingMission1      ' TRUE while playing mission 1 (or scoring mission 1 during final mission).
Dim bPlayingMission2      ' TRUE while playing mission 2 (weapon targets practice).
Dim bPlayingMission3      ' TRUE while playing mission 3 (bumpers super jets).
Dim bPlayingMission4      ' TRUE while playing mission 4 (POW).
Dim bPlayingMission5      ' TRUE while playing mission 5 (Defeat Allen O'Neill).
Dim bPlayingFinalMission  ' TRUE while playing final mission.
Dim bMissionSuspended   ' TRUE if the current mission is suspended.
Dim bInterruptedDisplay   ' TRUE when display is interrupted (ie: cannot be updated).
Dim bCanDisplayMessage    ' TRUE allows ShowNextMessage() process.
Dim bEnteringAHighScore   ' TRUE while the player is entering their name into the high score table.
Dim bAllowEnterSetup      ' TRUE if setup mode is allowed (during attract mode)
Dim bRunningSetup       ' TRUE during setup.
Dim bLeftFlipPressed      ' Used to entering setup mode (TRUE while left flipper key is pressed during game over mode).
Dim bRightFlipPressed   ' Also used to entering setup mode (TRUE while right flipper key is pressed during game over mode).

Dim nvR1, nvR2, nvR3, nvR4, nvR5
Dim nvS1, nvS2, nvS3, nvS4, nvS5

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    LoadEM
    Dim i
    Randomize


    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 50 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
        .CreateEvents "plungerIM"
    End With


   DisplayB2SText "           GAME OVER            "

    Kicker5.CreateBall
    Kicker5.kick 0, 2

  ' Display "GAME OVER" message.
  D1.Text = "           GAME OVER            "

  FixTargetTimer.Enabled = TRUE

    'StopSound Song:Song = ""
    PlaySound "mu_MSmissioncomplete"

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' load saved values, highscore, names, jackpot
    Loadhs

    ' Initalise the DMD display
    DMD_Init

    ' freeplay or coins
    bFreePlay = False 'we dont want coins

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
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bJackpot = False
    bInstantInfo = False
    ' set any lights for the attract mode
    GiOff
    EndAttractMode
    'StartAttractMode
    vpmtimer.addtimer 8000, "EnterAttractMode '"
 '   FuturePinball_BeginPlay

    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    Next

    If Credits > 0 Then DOF 136, DOFOn

End Sub


' *********************************************************************
' **                                                                 **
' **  FUTURE PINBALL DEFINED SCRIPT EVENTS                           **
' **                                                                 **
' *********************************************************************



' The subroutine is called immediately the game engine is ready to start processing the script.
'
' *** DOCAM OPTIMIZED ***
'
Sub FuturePinball_BeginPlay()
' Randomize
  ' Check if fpRAM is valid (regardling 'signature').
' If nvS1 <> "METALSLUGv2" Then
    ' If the 'signature' doesn't match, reset to default (factory) table settings.
'   SetupMenuResetSettings()
    ' Write the valid 'signature' to fpRAM.
'   nvS1 = "METALSLUGv2"
' End If
  ' Set the tilt sensitivity.
' Select Case nvR5
'   Case 1: nvTiltWarnings = 9
'   Case 2: nvTiltWarnings = 4
'   Case 3: nvTiltWarnings = 2
'   Case 4: nvTiltWarnings = 0
' End Select
  ' Display "GAME OVER" message.
' D1.Text = "           GAME OVER            "
' D2.Text = "           GAME OVER            "
  ' Create the captive ball.
  'Kicker5.CreateCaptiveBall()
  'Kicker5.SolenoidPulse(100)
  ' Give a small delay before raising C1/C2/C3 drop targets.
' FixTargetTimer.Enabled = TRUE
  ' Play default music.
' PlayMusic 1, "MSmissioncomplete", FALSE
  ' Force player to put quarter (coin key) in the machine.
' nvCredits = 0
  ' Reset player's score to 0.
' nvScore(1) = 0
  ' Initialize any other variables and flags.
' bEnteringAHighScore = FALSE
' BallsOnPlayfield = 0
  ' Initialize setup mode flags.
' bAllowEnterSetup = FALSE
' bRunningSetup = FALSE
' bLeftFlipPressed = FALSE
' bRightFlipPressed = FALSE
  ' Set the GAME OVER state (attract mode).
' EndOfGame()
End Sub



' This subroutine is called when the user has exited the game by pressing Escape key (Metal Slug table doesn't use this subroutine).
'
' *** DOCAM OPTIMIZED ***
'
'Sub FuturePinball_EndPlay()
'End Sub



' The user has pressed a key.
'
' *** DOCAM MUST BE CHECKED ***
'



'******
' Keys
'******


Sub Table1_KeyDown(ByVal Keycode)
    If Keycode = AddCreditKey Then
        Credits = Credits + 1
    DOF 136, DOFOn
        If(Tilted = False) Then
            PlaySoundAtVol "fx_coin", drain, 1
            DisplayB2SText " CREDITS " &credits
            D1.Text = " CREDITS " &credits
            PlaySound "go"
            If NOT bGameInPlay Then ShowTableInfo:
        End If
    End If

    If keycode = PlungerKey Then
    if bBallInPlungerLane Then
      DOF 125, DOFPulse
      DOF 114, DOFPulse
    end if
        PlungerIM.AutoFire
      ' If allowed, activate the ball saver (during 15, 30, 45 or 60 seconds, regardling table setup).
      If bCanActivateBallSaver Then
        bCanActivateBallSaver = FALSE
        bBallSaverIsActive = TRUE
        BallSaverClockTimer.Uservalue = 1
        BallSaverClockTimer.enabled = TRUE
        ShootAgainLight.BlinkInterval = 150
        ShootAgainLight.State = BulbBlink
      End If
    End If

    If bGameInPlay AND NOT Tilted Then
        If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 6:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
        If keycode = MechanicalTilt Then PlaySound "fx_nudge",0,1,1,0,25:CheckTilt
        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        If keycode = StartGameKey Then
            If((PlayersPlayingGame < MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                Else
                    If(Credits > 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
            If Credits < 1 Then DOF 136, DOFOff
                    Else
                        ' Not Enough Credits to start a game.

                        'DMD CenterLine(0, "CREDITS " & Credits), CenterLine(1, "INSERT COIN"), 0, eNone, eBlink, eNone, 500, True, ""
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
                    If(Credits > 0) Then
                        If(BallsOnPlayfield = 0) Then
                            Credits = Credits - 1
              If Credits < 1 Then DOF 136, DOFOff
                            ResetForNewGame()
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        D1.Text = " CREDITS " &credits&"   INSERT COIN"
                        DisplayB2SText " CREDITS " &credits &"   INSERT COIN "
                    End If
                End If
            End If
    End If ' If (GameInPlay)

    If hsbModeActive Then EnterHighScoreKey(keycode)

' Table specific
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then SolLFlipper 0
        If keycode = RightFlipperKey Then SolRFlipper 0
    End If
End Sub


Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    bInstantInfo = True
    DMDFlush
    UltraDMDTimer.Enabled = 1
End Sub

Sub InstantInfo
     D1.Text = " POW " &POWBonusCount
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

Sub Table1_Exit():
' Thalamus : Exit in a clean and proper way
  savehs
  Controller.Pause = False
  Controller.Stop
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
  startB2S(4)
    If Enabled Then
        PlaySoundAtVol SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
            RotateLaneLightsLeft

    Else
        PlaySoundAtVol SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
  startB2S(4)
    If Enabled Then
        PlaySoundAtVol SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
            RotateLaneLightsRight

    Else
        PlaySoundAtVol SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", VolFlip
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", VolFlip
End Sub

Sub RotateLaneLightsLeft
    Dim TempState
    TempState = POWLaneLight1.State
    POWLaneLight1.State = POWLaneLight2.State
    POWLaneLight2.State = POWLaneLight3.State
    POWLaneLight3.State = TempState
End Sub

Sub RotateLaneLightsRight
    Dim TempState
    TempState = POWLaneLight3.State
    POWLaneLight3.State = POWLaneLight2.State
    POWLaneLight2.State = POWLaneLight1.State
    POWLaneLight1.State = TempState
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
        'DMD "", " ", "CAREFUL!", 800
        DisplayB2SText "CAREFUL!"
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        'DMDFlush
        'DisplayB2SText "TILT!"
        DMD "", " ", "TILT!", 99999
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

        LeftSlingshotRubber.Disabled = 1
        RightSlingshotRubber.Disabled = 1
    Else
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        'Bumper1.Force = 6
        LeftSlingshotRubber.Disabled = 0
        RightSlingshotRubber.Disabled = 0
        'clean up the buffer display
        DMDFlush
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
        If UBound(tmp) = 0 Then 'we have 4 captive balls on the table (-1 means no balls, 0 is the first ball, 1 is the second..)
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
End Sub

Sub GiOff
    DOF 126, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSatDark"
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

'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5, BallShadow6)

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

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_plastichit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'

Sub StopGameOverSong
    PlaySong "mu_end"
    StopSound Song:Song = "":
  Stopsound "MSgameover"
  Stopsound "MShighscore"
End Sub

Sub ResetForNewGame()
    Dim i

    'D2.text = "juego nuevo"

    EndMultiball()

    LeftSlingshotRubber.Disabled = 0
    RightSlingshotRubber.Disabled = 0
    Bumper1.Force = 7
    Bumper2.Force = 7
    Bumper3.Force = 7

    bGameInPLay = True
    StopSound Song:Song = "":
    PlaySong "mu_Metalslug3"

   DisplayB2SText "           METAL SLUG           "
   D1.Text = "           METAL SLUG           "

  EndAttractMode()
  Ball = 0

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

    DisplayScore

    ' initialise any other flags
    Tilt = 0


    ' you may wish to start some music, play a sound, do whatever at this point
    FirstBallDelayTimer.enabled = 1
  ' Pick randomly the first POW name (between 1 to 100 inclusive).
  POWNamePtr = Int(Rnd(1) * 100) + 1

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

    bCanActivateBallSaver = TRUE


  bPlayingBall = TRUE


    ' set the current players bonus multiplier back down to 1X
    'SetBonusMultiplier 1

    ' reset any drop targets, lights, game Mode etc..

    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False

    'Reset any table specific


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
    PlaySoundAtVol SoundFXDOF("fx_Ballrel", 121, DOFPulse, DOFContactors), Ballrelease, 1
    BallRelease.Kick 95, 4
  DOF 109, DOFPulse

'PlungerKickerSolenoidPulse()

' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield> 1 Then
        DOF 129, DOFPulse
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
  Playsound "Marcodie"
  Playsound "explode3"
  FlashOff()
  ' Assume the player isn't playing a ball.
  bPlayingBall = FALSE
' D1.QueueText "                                ", seScrollUpOver, 10
' D2.QueueText "                                ", seScrollUpOver, 10
  If bPlayingMission Then EndMission()
  If Not(Tilted) Then
'   D1.FlushQueue()
'   D2.FlushQueue()
    InterruptDisplay(1600)
    D1.Text = "<<<<++  KILLED IN ACTION  ++>>>>"
        DisplayB2SText  "<<<<++  KILLED IN ACTION  ++>>>>"
'   D1.QueueText "<<<<++  KILLED IN ACTION  ++>>>>", seWipeIn, 800
'   D2.QueueText "<<<<++  KILLED IN ACTION  ++>>>>", seWipeIn, 800
  End If
  EndOfBallTimer2.Enabled = TRUE
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
            ShootAgainLight.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
        'DMD "extra-ball.mp4", "", "", 5000

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
           ' CheckHighScore()
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
'Sub EndOfBallComplete()
'    Dim NextPlayer

    'debug.print "EndOfBall - Complete"

    ' are there multiple players playing this game ?
'    If(PlayersPlayingGame> 1) Then
        ' then move to the next player
'        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
'        If(NextPlayer> PlayersPlayingGame) Then
'            NextPlayer = 1
'        End If
'    Else
'        NextPlayer = CurrentPlayer
'    End If

    'debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
'    If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

        ' set the machine into game over mode
'        EndOfGame()

    ' you may wish to put a Game Over message on the desktop/backglass

'    Else
        ' set the next player
'        CurrentPlayer = NextPlayer

        ' make sure the correct display is up to date
 '       AddScore 0

        ' reset the playfield for the new player (or new ball)
 '       ResetForNewPlayerBall()

        ' AND create a new ball
 '       CreateNewBall()

        ' play a sound if more than 1 player
 '       If PlayersPlayingGame> 1 Then
 '           PlaySound "vo_player" &CurrentPlayer
 '           D1.Text = "                         PLAYER " &CurrentPlayer
 '           DMD "", " ", "PLAYER " &CurrentPlayer, 800
 '       End If
 '   End If
'End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    bGameInPLay = False
    ' just ended your game then play the end of game tune
    If NOT bJustStarted Then
    'PlaySong "m_end"
    End If

  Dim i, j, iswap
  Dim POWNameTemp
  Dim Rand(100)

    EndMission()

    'D2.text = "game over"

    PlayersPlayingGame = 0

    bJustStarted = False
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    LeftSlingshotRubber.Disabled = 1
    RightSlingshotRubber.Disabled = 1
    Bumper1.Force = 0
    Bumper2.Force = 0
    Bumper3.Force = 0

    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball
    'PlayQuote.Enabled = 0
    ' show game over on the DMD
    'DMD "game-over.mp4", "", "", 11000

    ' set any lights for the attract mode
    GiOff
    FlashEffect 0
    FlashEffectMissionTimer.enabled = 0
    'StartAttractMode

' you may wish to light any Game Over Light you may have





  Kicker1SolenoidPulse()
  Kicker2SolenoidPulse()
  Kicker3SolenoidPulse()
  Kicker4SolenoidPulse()
   bBallLockedInKicker1 = 0
   bBallLockedInKicker2 = 0
   bBallLockedInKicker3 = 0
   bBallLockedInKicker4 = 0
  bBallInKicker1 = 0
  bBallInKicker2 = 0
  bBallInKicker3 = 0
  bBallInKicker4 = 0
   BallsLocked = 0
   FlashOff()
  ' BallsOnPlayfield = 0
  PlayfieldTimer.Enabled = TRUE
  ' Let Future Pinball know that the game has finished. This also clear the fpGameInPlay flag.
  'EndGame()
  ' LookAtBackbox()
   D1.Text = "           GAME OVER            "
   DisplayB2SText "           GAME OVER            "
  Playsound "msmissioncomplete"
  ' Ensure that the flippers are down.
  'LeftFlipper.SolenoidOff()
  'RightFlipper.SolenoidOff()
  ' Entering attract mode (lights and messages).
  EnterAttractMode()
  ' Doing POW name table randomly shuffled.
  For i = 1 To 100
    Rand(i) = Int(Rnd(1) * 10000)
  Next
  For i = 1 To 99
    For j = i + 1 To 100
      If Rand(i) > Rand(j) Then
        iswap = Rand(i)
        Rand(i) = Rand(j)
        Rand(j) = iswap
        POWNameTemp = POWName(i)
        POWName(i) = POWName(j)
        POWName(j) = POWNameTemp
      End If
    Next
  Next


FixTargetTimer.enabled = 1


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



' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
  'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    PlaySoundAtVol "fx_sensor", Primitive34, 1
    bBallInPlungerLane = TRUE
    ' turn on Launch light is there is one
    'LaunchLight.State = 2

    'be sure to update the Scoreboard after the animations, if any
    UltraDMDScoreTimer.Enabled = 1

    ' kick the ball in play if the bAutoPlunger flag is on
 '   If bAutoPlunger Then
        'debug.print "autofire the ball"
 '       PlungerIM.AutoFire
'   DOF 131, DOFPulse
'        DOF 125, DOFPulse
'        bAutoPlunger = False
'    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    End If
    'Start the Selection of the skillshot if ready
 '   If bSkillShotReady Then
 '       swPlungerRest.TimerEnabled = 1 ' this is a new ball, so show the launch ball if inactive for 6 seconds
 '       UpdateSkillshot()
 '   End If
    ' remember last trigger hit by the ball.

End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
 '   If bSkillShotReady Then
 '       ResetSkillShotTimer.Enabled = 1
 '   End If
 '   If NOT bMultiballMode Then
      '  PlaySong "mu_msmain"
 '   End If
    DMDFLush
    DMD "Explosion.gif", "", "", 2800
' turn off LaunchLight
' LaunchLight.State = 0
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
    ShootAgainLight.BlinkInterval = 160
    SetLightColor ShootAgainLight, amber, 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimerExpired_Timer()
    'debug.print "Ballsaver ended"
    BallSaverTimerExpired.Enabled = False
    ' clear the flag
    bBallSaverActive = False
    ' if you have a ball saver light then turn it off at this point
    ShootAgainLight.State = 0
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    'debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    ShootAgainLight.BlinkInterval = 80
    ShootAgainLight.State = 2
End Sub

' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board
' In this table we use SecondRound variable to double the score points in the second round after killing Malthael

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub AddScore(points)
'     If(Tilted = False) Then
'         ' add the points to the current players score variable
'         Score(CurrentPlayer) = Score(CurrentPlayer) + points * SecondRound
'         ' update the score displays
'         'DMDScore
'     End if
'
' ' you may wish to check to see if the player has gotten a replay
' End Sub

' Add bonus to the bonuspoints AND update the score board

Sub AddBonus(points) 'not used in this table, since there are many different bonus items.
    If(Tilted = False) Then
        ' add the bonus to the current players bonus variable
        BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
        ' update the score displays
       ' DMDScore
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



' Set the Bonus Multiplier to the specified level AND set any lights accordingly
' There is no bonus multiplier lights in this table



Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
       ' DMDBlink "", " ", "EXTRA BALL WON", 100, 10
        'DMD "extra-ball.mp4", "", "", 5000
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        GiEffect 1
        LightEffect 2
    END If
End Sub

Sub AwardSpecial()
    DisplayB2SText "  CONGRATULATION GREAT SCORE  "
    Credits = Credits + 1
    DOF 140, DOFOn
    DisplayB2SText "           CREDITS           " &Credits
    GiEffect 1
    LightEffect 1
End Sub

'in this table the jackpot is always 1 million + 10% of your score

Sub AwardJackpot() 'award a normal jackpot, double or triple jackpot
    Jackpot = 1000000 + Round(Score(CurrentPlayer) / 10, 0)
   ' DMDBlink "", "JACKPOT", Jackpot, 100, 10
    PlaySound "criticalhit"
    AddScore Jackpot
    GiEffect 1
    LightEffect 2
    FlashEffect 2
End Sub

Sub AwardDoubleJackpot() 'in this table the jackpot is always 1 million + 10% of your score
    Jackpot = (1000000 + Round(Score(CurrentPlayer) / 10, 0) ) * 2
   ' DMDBlink "", "DOUBLE JACKPOT", Jackpot, 100, 10
    PlaySound "criticalhit"
    AddScore Jackpot
    GiEffect 1
    LightEffect 2
    FlashEffect 2
End Sub

Sub AwardSuperJackpot() 'this is actually a tripple jackpot
    DOF 133, DOFPulse
    SuperJackpot = (1000000 + Round(Score(CurrentPlayer) / 10, 0) ) * 3
   ' DMDBlink "", "SUPER JACKPOT", SuperJackpot, 100, 10
    PlaySound "criticalhit"
    AddScore SuperJackpot
    GiEffect 1
    LightEffect 2
    FlashEffect 2
End Sub

Sub AwardSkillshot()
    Dim i
    DOF 131, DOFPulse
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMDFlush
    Select case SkillShotValue(CurrentPlayer)
        case 1000000
           ' DMD "skillshot1.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case 2000000
           ' DMD "skillshot2.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case 3000000
           ' DMD "skillshot3.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case 4000000
          '  DMD "skillshot4.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case 5000000
          '  DMD "skillshot5.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
        case ELSE
          '  DMD "skillshot.mp4", " ", " ", 5000
            AddScore SkillshotValue(CurrentPLayer)
    End Select
    ' increment the skillshot value with 1 million
    SkillShotValue(CurrentPLayer) = SkillShotValue(CurrentPLayer) + 1000000
    'do some light show
    GiEffect 1
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

    'x = LoadValue(TableName, "Jackpot")
    'If(x <> "") then Jackpot = CDbl(x) Else Jackpot = 200000 End If
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
    'SaveValue TableName, "Jackpot", Jackpot
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
    Dim tmp
    tmp = Score(1)

    If Score(2)> tmp Then tmp = Score(2)
    If Score(3)> tmp Then tmp = Score(3)
    If Score(4)> tmp Then tmp = Score(4)

    If tmp> HighScore(1) Then 'add 1 credit for beating the highscore
        AwardSpecial()
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
    hsEnteredDigits(1) = "A"
    hsEnteredDigits(2) = "A"
    hsCurrentDigit = 0

    hsValidLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ<+-0123456789" ' < is used to delete the last letter
    hsCurrentLetter = 1
    DMDFlush
    DMDId "hsc", "", "YOUR NAME:", " ", 999999
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
    BallShadowUpdate

  If KickbackLight.state = 1 And bGameInPlay And Not Tilted Then
     LeftOutLaneTrigger.enabled = 1
   Else
     LeftOutLaneTrigger.enabled = 0
  End If
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
    SetLightColor ShootAgainLight, amber, -1
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
      Dim i
        ' ShowTableInfo
        'UltraDMD.SetScoreboardBackgroundImage "metalslug.png", 15, 15
        DMD "metalslug.jpg", "", "", 5000
       ' UltraDMD.DisplayScene00 "MetalSluglogo.png", "", 14, 2, "", -1, -1', UltraDMD_Animation_None, 100, UltraDMD_Animation_None
      '  UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer, "Ball " & Balls
      Else
        UltraDMDScoreTimer.Enabled = 1
      End If

End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
    DisplayScore
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


Sub DMD_Init
    'Set UltraDMD = CreateObject("UltraDMD.DMDObject")
  Dim FlexDMD
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
    DirName = curDir& "\" &TableName& ".UltraDMD"

    If Not fso.FolderExists(DirName) Then _
            Msgbox "UltraDMD userfiles directory '" & DirName & "' does not exist." & CHR(13) & "No graphic images will be displayed on the DMD"
    UltraDMD.SetProjectFolder DirName

    ' wait for the animation to end
    ' While UltraDMD.IsRendering = True
    ' WEnd

    ' Show ROM version number
    DMD "", "Metal Slug", "ROM VERS " &myVersion, 2000
End Sub

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo

    Dim i
    'info goes in a loop only stopped by the credits and the startkey
    If Score(1) Then
        DMD "", "PLAYER 1", Score(1), 3000
    End If
    If Score(2) Then
        DMD "", "PLAYER 2", Score(2), 3000
    End If
    If Score(3) Then
        DMD "", "PLAYER 3", Score(3), 3000
    End If
    If Score(4) Then
        DMD "", "PLAYER 4", Score(4), 3000
    End If


    'coins or freeplay
    If bFreePlay Then
        DMD " ", "FREE PLAY", 2000
        DMD "IntroMS.mp4", "", "", 24500
    Else
        If Credits> 0 Then
            DMD "", "CREDITS " &credits, "PRESS START", 2000
        Else
            DMD "", "CREDITS " &credits, "INSERT COIN", 2000
        End If
        DMD "IntroMS.mp4", "", "", 24500
    End If
    ' some info about the table
    DMD "", "Javier And Pinwizkid", "PRESENTS", 3000
    DMD "", "", " Metal Slug ", 3000
    ' Highscores
    DMD "", "HIGHSCORE 1", HighScoreName(0) & " " & HighScore(0), 3000
    DMD "", "HIGHSCORE 2", HighScoreName(1) & " " & HighScore(1), 3000
    DMD "", "HIGHSCORE 3", HighScoreName(2) & " " & HighScore(2), 3000
    DMD "", "HIGHSCORE 4", HighScoreName(3) & " " & HighScore(3), 3000
End Sub

Sub StartAttractMode()
    bAttractMode = True
    UltraDMDTimer.Enabled = 1
    StartLightSeq
    ShowTableInfo
    StopGameOverSong
End Sub

Sub StopAttractMode()
    bAttractMode = False
    DMDScoreNow
    LightSeqAttract.StopPlay
    LightSeqFlasher.StopPlay
  '  StopRainbow
 '   ResetAllLightsColor
'StopSong
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
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




Sub Game_Init() 'called at the start of a new game
    Dim i
    bExtraBallWonThisBall = False
    TurnOffPlayfieldLights()



End Sub


Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub ResetNewBallLights()
'    LightArrow1.State = 2
'    LightArrow6.State = 2
'    l53.State = 2
End Sub


Sub UpdateSkillShot() 'Updates the skillshot light
    LightSeqSkillshot.Play SeqAllOff
'DMDFlush
End Sub

































' The user has paused the game (Metal Slug table doesn't use this subroutine).
'
' *** DOCAM OPTIMIZED ***




' The user has resumed the paused game (Metal Slug table doesn't use this subroutine).
'
' *** DOCAM OPTIMIZED ***
'



' The played has nudged the table a little too hard/much and a "Warning" must be given to the player.
' Metal Slug table uses to clear warnings if tilt is disabled.
'
' *** DOCAM OPTIMIZED ***
'




' The player has tilted the machine (too many "Warnings").
'
' *** DOCAM MUST BE CHECKED ***
'
Sub FuturePinball_Tilted()
  If Not(bPlayingBall) Then
    ' Ignore the tilt when the player is not playing (e.g. during bonus countdown sequence).
    Tilted = FALSE
    Exit Sub
  End If
  PlaySound "Tilt"
  ' Terminate current mission.
   EndMission()
  ' Stop theme.
   StopMusic 1
   Playsound "mordenlaugh", FALSE
   If bPlayingFinalMission Then ingamereset()
  ' Ensure all the flippers are down (as the keys won't work from now on).
  LeftFlipper.SolenoidOff()
  RightFlipper.SolenoidOff()
   TopRightFlipper.SolenoidOff()
  ' Ensure they're are not any ball near kickback.
  bKickbackIsLit = FALSE
  KickbackLight.State = 0
  KickbackGuide.isDropped = 1
  ' Disable ball saver.
  BallSaverClockTimerUserData = 0
  bBallSaverIsActive = FALSE
  ShootAgainLight.State = 0
  ShootAgainLight.BlinkInterval = 150
  ' Ensure the wall (WEAPON drop targets) is down during locating ball(s).
  W1.Isdropped = 1
  W2.Isdropped = 1
  W3.Isdropped = 1
  W4.Isdropped = 1
  W5.Isdropped = 1
  W6.Isdropped = 1
  ' Turn off general illumination (GI).
   GIOff()
    FlashEffect 0
    FlashEffectMissionTimer.enabled = 0
  ' Start the tilt recovery timer which waits until all balls have drained before doing the end of ball sequence (or end of game).
  TiltRecoveryTimer.Interval = 1000
  TiltRecoveryTimer.Enabled = TRUE
End Sub



' A music channel has finished playing. Channel is set to the channel number that has finished (Metal Slug table doesn't use this subroutine).
'
' *** DOCAM OPTIMIZED ***
'




' High Score entry has been completed by the player.
' Position is set to the position in the high score table (1 - 10), or 0 when no new high score.
' Special is set to 1 if the Special high score was beaten (Metal Slug table doesn't use Special high score).
'
' *** DOCAM OPTIMIZED ***
'
Sub FuturePinball_NameEntryComplete(ByVal Position, ByVal Special)
  bEnteringAHighScore = FALSE
  EndOfBallComplete()
End Sub



' *********************************************************************
' **                                                                 **
' **  TABLE OBJECT SCRIPT EVENTS (HIT & UNHIT)                       **
' **                                                                 **
' *********************************************************************



' *********************************************************************
' **  PLUNGER LANE TRIGGER                                           **
' *********************************************************************



' A ball is pressing down the trigger in the plunger lane.
'
' *** DOCAM OPTIMIZED ***
'
Sub PlungerLaneTrigger_Hit()
  bBallInPlungerLane = TRUE
End Sub



' The ball quits the plunger lane trigger.
'
' *** DOCAM OPTIMIZED ***
'
Sub PlungerLaneTrigger_UnHit()
  bBallInPlungerLane = FALSE
End Sub



' *********************************************************************
' **  "POW" LANES TRIGGERS                                           **
' *********************************************************************



' The "P" letter trigger has been hit (POW lanes).
'
' *** DOCAM OPTIMIZED ***
'
Sub POWLane1_Hit()
  startB2S(3)
  DOF 127, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  AddScore(100)
  PlaySoundAtVol "Lane1", ActiveBall, 1
  ' Turn on "P" letter light.
  POWLaneLight1.State = 1
  ' Check if all POW letters are completed.
  CheckPOWLanes()
End Sub



' The "O" letter trigger has been hit (POW lanes).
'
' *** DOCAM OPTIMIZED ***
'
Sub POWLane2_Hit()
  startB2S(2)
  DOF 128, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  AddScore(100)
  PlaySoundAtVol "Lane1", ActiveBall, 1
  ' Turn on "O" letter light.
  POWLaneLight2.State = 1
  ' Check if all POW letters are completed.
  CheckPOWLanes()
End Sub



' The "W" letter trigger has been hit (POW lanes).
'
' *** DOCAM OPTIMIZED ***
'
Sub POWLane3_Hit()
  startB2S(1)
  DOF 129, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  AddScore(100)
  PlaySoundAtVol "Lane1", ActiveBall, 1
  ' Turn on "W" letter light.
  POWLaneLight3.State = 1
  ' Check if all POW letters are completed.
  CheckPOWLanes()
End Sub



' *********************************************************************
' **  THE LOOP TRIGGER                                               **
' *********************************************************************



' The loop trigger has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub LoopTrigger_Hit()
  startB2S(6)
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  ResetWEAPONTargets()
  PlaySoundAtVol "explode3", ActiveBall, 1
  AddScore(2500)
End Sub



' *********************************************************************
' **  OUTLANES & INLANES TRIGGERS                                    **
' *********************************************************************



' The left inlane trigger has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub LeftInLaneTrigger_Hit()
  startB2S(3)
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  DOF 111, DOFPulse
  ' Add 1000 points.
  AddScore(1000)
  ' Play specific sound.
   PlaySoundAtVol "go", ActiveBall, 1
  ' Lit the left inlane light.
   LeftInLaneLight.State = 1
   bLeftInLaneIsLit = TRUE
  If bRightInLaneIsLit Then
    ' If the right inlane light is also lit, lit the kickback...
    LightKickback()
    ' ...and the red right outlane light.
    RightOutLaneLight.State = 1
  End If
End Sub



' The right inlane trigger has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub RightInLaneTrigger_Hit()
  startB2S(2)
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  DOF 112, DOFPulse
  ' Add 1000 points.
  AddScore(1000)
  ' Play specific sound.
   PlaySoundAtVol "go", ActiveBall, 1
  ' Lit the right inlane light.
   RightInLaneLight.State = 1
   bRightInLaneIsLit = TRUE
  If bLeftInLaneIsLit Then
    ' If the left inlane light is also lit, lit the kickback...
    LightKickback()
    ' ...and the red right outlane light.
    RightOutLaneLight.State = 1
  End If
End Sub



' The left outlane trigger has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub LeftOutLaneTrigger_Hit()
  startB2S(4)
  If fpTilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  DOF 110, DOFPulse
  ' This flag indicates the trigger is pressed.
  bBallOnLeftOutLaneTrigger = TRUE
   If bKickbackIsLit Then
    KickbackTimer.Uservalue = 1 ' UserData = 1 means ready for kickback.
    KickbackTimer.Enabled = TRUE
    KickbackLight.state = BulbOff', "10", 100
    FlashForMs KickbackLight, 500, 50, 0
  Else
    ' Kickback is disabled.
    ' Play outlane sound.
    PlayOutLaneSound()
    If bBallSaverIsActive Then
      ' Give 3 seconds saver grace delay, only if ball saver is active.
      BallSaverClockTimer.Uservalue = 3
      BallSaverClockTimer.Enabled = FALSE
      BallSaverClockTimer.interval = 3000
      ShootAgainLight.BlinkInterval = 50
      ShootAgainLight.State = BulbBlink
    End If
    ' Add 2000 points.
    AddScore(2000)
  End If
End Sub



' The right outlane trigger has been unhit.
'
' *** DOCAM OPTIMIZED ***
'
Sub LeftOutLaneTrigger_UnHit()
  startB2S(4)
  If fpTilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  ' Disable the kickback mechanism.
  bBallOnLeftOutLaneTrigger = FALSE
  KickbackTimer.Enabled = FALSE
  KickbackTimer.Uservalue = 0
End Sub



' The right outlane trigger has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Dim BallSaverClockTimerUserData

Sub RightOutLaneTrigger_Hit()
  startB2S(1)
  If fpTilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  DOF 113, DOFPulse
  ' Add 2000 points.
  AddScore(2000)

      If ShootAgainLight.State = 0 Then
         bBallSaverIsActive = FALSE
      End If

  If bBallSaverIsActive = True Then
    ' Give 3 seconds saver grace delay, only if ball saver is active.
    BallSaverClockTimer.Uservalue = 3
    BallSaverClockTimer.Enabled = FALSE
    BallSaverClockTimer.interval = 3000
    ShootAgainLight.BlinkInterval = 50
    ShootAgainLight.State = BulbBlink
  End If
  ' Play outlane sound.
  PlayOutLaneSound()
  ' Give additional 2000 points only if (red) right outlane is lit.
   If RightOutLaneLight.State = BulbOn Then AddScore(20000)
End Sub


' *********************************************************************
' **  RUBBERS & SLINGSHOTS                                           **
' *********************************************************************



' The left slingshot has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Dim LStep, RStep
Sub LeftSlingshotRubber_Slingshot()
  startB2S(4)
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.

    PlaySoundAtVol SoundFXDOF("LeftSlingShot", 103, DOFPulse, DOFContactors), lemk, 1
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingshotRubber.TimerEnabled = True

  ' Play specific sound.
  Playsound "Rocket1"
  ' Add 110 points.
  AddScore(110)
  ' Flash the lights around the lime slingshot.
  FlashForMs LeftSlingshotBulb1, 100, 50, 1
  FlashForMs LeftSlingshotBulb2, 100, 50, 1
  DOF 140, DOFPulse
End Sub


Sub LeftSlingshotRubber_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingshotRubber.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub

' The right slingshot has been hit,
'
' *** DOCAM OPTIMIZED ***
'
Sub RightSlingshotRubber_Slingshot()
  startB2S(4)
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.

    PlaySoundAtVol SoundFXDOF("RightSlingShot", 104, DOFPulse, DOFContactors), remk, 1

    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShotRubber.TimerEnabled = True

  ' Play specific sound.
  Playsound "Rocket1"
  ' Add 110 points.
  AddScore(110)
  ' Flash the lights around the right slingshot.
  FlashForMs RightSlingshotBulb1, 100, 50, 1
  FlashForMs RightSlingshotBulb2, 100, 50, 1
  DOF 141, DOFPulse
End Sub

Sub RightSlingShotRubber_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShotRubber.TimerEnabled = 0
    End Select

    RStep = RStep + 1
End Sub


' The left slingshot has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Rubber6_Hit()
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  AddScore(10)
  PlaysoundAtVol "go", ActiveBall, 1
End Sub



' *********************************************************************
' **  METALSLUG DROP TARGETS BANKS                                   **
' *********************************************************************



' The drop target "M" (from METAL drop targets bank) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Slug1_Hit()
  startB2S(1)
  DOF 122, DOFPulse
  ' Raise the drop target and ignore if the table is tilted.
  If Tilted Then
    Slug1.SolenoidPulse()
    Exit Sub
  End If
  AddScore(500)
  Slug1Light.State = 1
  CheckSLUGTimer.Enabled = TRUE
  PlaySlugSound()
  PlaysoundAtVol "gun8", ActiveBall, 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The drop target "E" (from METAL drop targets bank) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Slug2_Hit()
  startB2S(1)
  DOF 122, DOFPulse
  ' Raise the drop target and ignore if the table is tilted.
  If Tilted Then
    Slug2.SolenoidPulse()
    Exit Sub
  End If
  AddScore(500)
  Slug2Light.State = 1
  CheckSLUGTimer.Enabled = TRUE
  PlaySlugSound()
  PlaysoundAtVol "gun9", ActiveBall, 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The drop target "T" (from METAL drop targets bank) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Slug3_Hit()
  startB2S(1)
  DOF 122, DOFPulse
  ' Raise the drop target and ignore if the table is tilted.
  If Tilted Then
    Slug3.SolenoidPulse()
    Exit Sub
  End If
  AddScore(500)
  Slug3Light.State = 1
  CheckSLUGTimer.Enabled = TRUE
  PlaySlugSound()
  PlaysoundAtVol "gun8", ActiveBall, 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The drop target "A" (from METAL drop targets bank) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Slug4_Hit()
  startB2S(1)
  DOF 122, DOFPulse
  ' Raise the drop target and ignore if the table is tilted.
  If Tilted Then
    Slug4.SolenoidPulse()
    Exit Sub
  End If
  AddScore(500)
  Slug4Light.State = 1
  CheckSLUGTimer.Enabled = TRUE
  PlaySlugSound()
  PlaysoundAtVol "gun9", ActiveBall, 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The drop target "L" (from METAL drop targets bank) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Slug5_Hit()
  startB2S(1)
  DOF 122, DOFPulse
  ' Raise the drop target and ignore if the table is tilted.
  If Tilted Then
    Slug5.SolenoidPulse()
    Exit Sub
  End If
  AddScore(500)
  Slug5Light.State = 1
  CheckSLUGTimer.Enabled = TRUE
  PlaySlugSound()
  PlaysoundAtVol "gun8", ActiveBall, 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The drop target "S" (from SLUG drop targets bank) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Slug6_Hit()
  startB2S(1)
  DOF 132, DOFPulse
  ' Raise the drop target and ignore if the table is tilted.
  If Tilted Then
    Slug6.SolenoidPulse()
    Exit Sub
  End If
  AddScore(500)
  Slug6Light.State = 1
  CheckSLUGTimer.Enabled = TRUE
  PlaySlugSound()
  PlaysoundAtVol "gun9", ActiveBall, 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The drop target "L" (from SLUG drop targets bank) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Slug7_Hit()
  startB2S(1)
  DOF 132, DOFPulse
  ' Raise the drop target and ignore if the table is tilted.
  If Tilted Then
    Slug7.SolenoidPulse()
    Exit Sub
  End If
  AddScore(500)
  Slug7Light.State = 1
  CheckSLUGTimer.Enabled = TRUE
  PlaySlugSound()
  PlaysoundAtVol "gun8", ActiveBall, 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The drop target "U" (from SLUG drop targets bank) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Slug8_Hit()
  startB2S(1)
  DOF 132, DOFPulse
  ' Raise the drop target and ignore if the table is tilted.
  If Tilted Then
    Slug8.SolenoidPulse()
    Exit Sub
  End If
  AddScore(500)
  Slug8Light.State = 1
  CheckSLUGTimer.Enabled = TRUE
  PlaySlugSound()
  PlaysoundAtVol "gun9", ActiveBall, 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The drop target "G" (from SLUG drop targets bank) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Slug9_Hit()
  startB2S(1)
  DOF 132, DOFPulse
  ' Raise the drop target and ignore if the table is tilted.
  If Tilted Then
    Slug9.SolenoidPulse()
    Exit Sub
  End If
  AddScore(500)
  Slug9Light.State = 1
  CheckSLUGTimer.Enabled = TRUE
  PlaySlugSound()
  PlaysoundAtVol "gun8", ActiveBall, 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' *********************************************************************
' **  WEAPON DROP TARGETS BANK                                       **
' *********************************************************************



' The drop target "W" (from WEAPON drop targets bank, also called "the wall") has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub W1_Hit()
  startB2S(1)
  DOF 134, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
   If Not(bMultiBallMode) Then
    If bPlayingMission Then
      If bPlayingMission1 Then
        ScoreMission1()
        W1Light.State = 1
        bW1TargetIsDown = TRUE
        CheckWeapon()
      ElseIf bPlayingMission2 Then
        FlashForMs W1Light, 500, 100, 0
        ScoreMission2()
      ElseIf bPlayingMission5 Then
        FlashForMs W1Light, 500, 100, 0
        ScoreMission5()
      ElseIf bPlayingMission3 Or bPlayingMission4 Or bPlayingFinalMission Then
        PlaysoundAtVol "steel", ActiveBall, 1
        W1Light.State = 1
        bW1TargetIsDown = TRUE
        CheckWeapon()
        AddScore(500)
      End If
    Else
      PlaysoundAtVol "steel", ActiveBall, 1
      AddScore(500)
      W1Light.State = 1
      bW1TargetIsDown = TRUE
      CheckWeapon()
    End If
   End If
End Sub



' The drop target "E" (from WEAPON drop targets bank, also called "the wall") has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub W2_Hit()
  startB2S(1)
  DOF 134, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
   If Not(bMultiBallMode) Then
    If bPlayingMission Then
      If bPlayingMission1 Then
        ScoreMission1()
        W2Light.State = 1
        bW2TargetIsDown = TRUE
        CheckWeapon()
      ElseIf bPlayingMission2 Then
        FlashForMs W2Light, 500, 100, 0
        ScoreMission2()
      ElseIf bPlayingMission5 Then
        FlashForMs W2Light, 500, 100, 0
        ScoreMission5()
      ElseIf bPlayingMission3 Or bPlayingMission4 Or bPlayingFinalMission Then
        PlaysoundAtVol "steel", ActiveBall, 1
        W2Light.State = 1
        bW2TargetIsDown = TRUE
        CheckWeapon()
        AddScore(500)
      End If
    Else
      PlaysoundAtVol "steel", ActiveBall, 1
      AddScore(500)
      W2Light.State = 1
      bW2TargetIsDown = TRUE
      CheckWeapon()
    End If
   End If
End Sub



' The drop target "A" (from WEAPON drop targets bank, also called "the wall") has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub W3_Hit()
  startB2S(1)
  DOF 134, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
   If Not(bMultiBallMode) Then
    If bPlayingMission Then
      If bPlayingMission1 Then
        ScoreMission1()
        W3Light.State = 1
        bW3TargetIsDown = TRUE
        CheckWeapon()
      ElseIf bPlayingMission2 Then
        FlashForMs W3Light, 500, 100, BulbOff
        ScoreMission2()
      ElseIf bPlayingMission5 Then
        FlashForMs W3Light, 500, 100, BulbOff
        ScoreMission5()
      ElseIf bPlayingMission3 Or bPlayingMission4 Or bPlayingFinalMission Then
        PlaysoundAtVol "steel", ActiveBall, 1
        W3Light.State = 1
        bW3TargetIsDown = TRUE
        CheckWeapon()
        AddScore(500)
      End If
    Else
      PlaysoundAtVol "steel", ActiveBall, 1
      AddScore(500)
      W3Light.State = 1
      bW3TargetIsDown = TRUE
      CheckWeapon()
    End If
   End If
End Sub



' The drop target "P" (from WEAPON drop targets bank, also called "the wall") has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub W4_Hit()
  startB2S(1)
  DOF 134, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
   If Not(bMultiBallMode) Then
    If bPlayingMission Then
      If bPlayingMission1 Then
        ScoreMission1()
        W4Light.State = 1
        bW4TargetIsDown = TRUE
        CheckWeapon()
      ElseIf bPlayingMission2 Then
        FlashForMs W4Light, 500, 100, 0
        ScoreMission2()
      ElseIf bPlayingMission5 Then
        FlashForMs W4Light, 500, 100, 0
        ScoreMission5()
      ElseIf bPlayingMission3 Or bPlayingMission4 Or bPlayingFinalMission Then
        PlaysoundAtVol "steel", ActiveBall, 1
        W4Light.State = 1
        bW4TargetIsDown = TRUE
        CheckWeapon()
        AddScore(500)
      End If
    Else
      PlaysoundAtVol "steel", ActiveBall, 1
      AddScore(500)
      W4Light.State = 1
      bW4TargetIsDown = TRUE
      CheckWeapon()
    End If
   End If
End Sub



' The drop target "O" (from WEAPON drop targets bank, also called "the wall") has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub W5_Hit()
  startB2S(1)
  DOF 134, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
   If Not(bMultiBallMode) Then
    If bPlayingMission Then
      If bPlayingMission1 Then
        ScoreMission1()
        W5Light.State = 1
        bW5TargetIsDown = TRUE
        CheckWeapon()
      ElseIf bPlayingMission2 Then
        FlashForMs W5Light, 500, 100, 0
        ScoreMission2()
      ElseIf bPlayingMission5 Then
        FlashForMs W5Light, 500, 100, 0
        ScoreMission5()
      ElseIf bPlayingMission3 Or bPlayingMission4 Or bPlayingFinalMission Then
        PlaysoundAtVol "steel", ActiveBall, 1
        W5Light.State = 1
        bW5TargetIsDown = TRUE
        CheckWeapon()
        AddScore(500)
      End If
    Else
      PlaysoundAtVol "steel", ActiveBall, 1
      AddScore(500)
      W5Light.State = 1
      bW5TargetIsDown = TRUE
      CheckWeapon()
    End If
   End If
End Sub



' The drop target "N" (from WEAPON drop targets bank, also called "the wall") has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub W6_Hit()
  startB2S(1)
  DOF 134, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
   If Not(bMultiBallMode) Then
    If bPlayingMission Then
      If bPlayingMission1 Then
        ScoreMission1()
        W6Light.State = 1
        bW6TargetIsDown = TRUE
        CheckWeapon()
      ElseIf bPlayingMission2 Then
        FlashForMs W6Light, 500, 100, 0
        ScoreMission2()
      ElseIf bPlayingMission5 Then
        FlashForMs W6Light, 500, 100, 0
        ScoreMission5()
      ElseIf bPlayingMission3 Or bPlayingMission4 Or bPlayingFinalMission Then
        PlaysoundAtVol "steel", ActiveBall, 1
        W6Light.State = 1
        bW6TargetIsDown = TRUE
        CheckWeapon()
        AddScore(500)
      End If
    Else
      PlaysoundAtVol "steel", ActiveBall, 1
      AddScore(500)
      W6Light.State = 1
      bW6TargetIsDown = TRUE
      CheckWeapon()
    End If
   End If
End Sub



' *********************************************************************
' **  LEFT-SIDE DROP TARGETS BANK                                    **
' *********************************************************************



' The first left-side target has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub S1_Hit()
  startB2S(1)
  DOF 118, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  AddScore(1000)
  PlaysoundAtVol "explode4", ActiveBall, 1
  S1Light.State = 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The second left-side target has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub S2_Hit()
  startB2S(1)
  DOF 118, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  AddScore(2000)
  PlaysoundAtVol "explode4", ActiveBall, 1
  S2Light.State = 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' The last left-side target has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub S3_Hit()
  startB2S(1)
  DOF 118, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  AddScore(3000)
  PlaysoundAtVol "explode4", ActiveBall, 1
  S3Light.State = 1
  If bPlayingMission1 Then ScoreMission1()
End Sub



' *********************************************************************
' **  CAPTIVE BALL DROP TARGETS BANK                                 **
' *********************************************************************



' The captive drop target 1 (bottom) has been hit.
'
' *** DOCAM OPTIMIZED ****
'
Sub C1_Hit()
  startB2S(1)
  DOF 119, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  If bGameInPlay Then
    C1Light.State = 1
    AddScore(2000)
    If bPlayingMission Then
      If bPlayingMission1 Then ScoreMission1()
      If bPlayingMission4 Then ScoreMission4()
    Else
      LightMission()
      PlaysoundAtVol "Thankyou", ActiveBall, 1
    End If
  End If
End Sub



' The captive drop target 2 (middle) has been hit.
'
' *** DOCAM OPTIMIZED ****
'
Sub C2_Hit()
  startB2S(1)
  DOF 119, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  If bGameInPlay Then
    C2Light.State = 1
    AddScore(2000)
    If bPlayingMission Then
      If bPlayingMission1 Then ScoreMission1()
      If bPlayingMission4 Then ScoreMission4()
    Else
      LightMission()
      PlaysoundAtVol "Thankyou", ActiveBall, 1
    End If
  End If
End Sub



' The captive drop target 3 (top) has been hit.
'
' *** DOCAM OPTIMIZED ****
'
Sub C3_Hit()
  startB2S(1)
  DOF 119, DOFPulse
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  If bGameInPlay Then
    C1Light.State = 2
    C2Light.State = 2
    C3Light.State = 2
    AddScore(2000)
    If bPlayingMission Then
      If bPlayingMission1 Then ScoreMission1()
      If bPlayingMission4 Then ScoreMission4()
    Else
      LightMission()
      PlaysoundAtVol "Thankyou", ActiveBall, 1
    End If
  End If
End Sub



' *********************************************************************
' **  CAPTIVE BALL TARGET                                            **
' *********************************************************************



' The captive target has been hit.
'
' *** DOCAM OPTIMIZED ****
'
Sub C4_Hit()
  startB2S(6)
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
  ' This statement drops the next side target (useful to open access to kicker 4).
  If S1.IsDropped = 1 Then
    If S2.IsDropped = 1 Then
      If Not(S3.IsDropped) Then
        S3.isdropped = 1
        S3Light.State = 1
      End If
    Else
      S2.isdropped = 1
      S2Light.State = 1
    End If
  Else
    S1.isdropped = 1
    S1Light.State = 1
  End If
  PlaysoundAtVol "Ok", ActiveBall, 1
  AdvancePOW()
  AddScore(8000)
  If Not(bPlayingMission) And Not(bMissionSuspended) Then LightMission()
  If bPlayingMission Then
    If bPlayingMission1 Then ScoreMission1()
    If bPlayingMission4 Then ScoreMission4()
  End If
  If bPlayingFinalMission Then
    InterruptDisplay(600)
    AddScore(50000)
    'D1.FlushQueue()
    'D2.FlushQueue()
    D1.Text = "           * 40,000 *            "
    DisplayB2SText "           * 40,000 *            "
    'D1.QueueText "             40,000              ", seBlinkMask, 300
    'D2.QueueText "             40,000              ", seBlinkMask, 300
  End If
End Sub



' *********************************************************************
' **  DRAIN & KICKERS                                                **
' *********************************************************************



' Lost a ball (falling in drain).
'
' *** DOCAM OPTIMIZED ***
'
Sub Drain_Hit()
  startB2S(6)
  DOF 130, DOFPulse
  ' Destroy the ball.
  Drain.DestroyBall()
  BallsOnPlayfield = BallsOnPlayfield - 1
  ' Pretend to knock the ball into the ball storage mechanism.
  PlaySoundAtVol "Drain", drain, 1
   If fpTilted Then FixWeapon()
  ' If there is a game in progress and not tilted.
  If bGameInPlay And Not(fpTilted) Then
    If Not(bPlayingFinalMission) Then

      If ShootAgainLight.State = 0 Then
         bBallSaverIsActive = FALSE
      End If

      If bBallSaverIsActive = True Then
        bBallSaverIsActive = FALSE
        BallSaverClockTimer.Enabled = FALSE
        BallSaverClockTimer.Uservalue = 0
        ShootAgainLight.State = BulbOff
        ShootAgainLight.BlinkInterval = 150
        AutoBall()
        'D1.FlushQueue()
        'D2.FlushQueue()

        D1.Text = "           BALL SAVED           "
        DisplayB2SText "           BALL SAVED           "
        'D2.QueueText "           BALL SAVED           ", seBlinkMask, 150
      Else
        If (BallsOnPlayfield = 0) Then EndOfBall()
        If (BallsOnPlayfield = 1) And (BallsLocked = 1) Then EndOfBall()
        If (BallsOnPlayfield = 2) And (BallsLocked = 2) Then EndOfBall()
        If (BallsOnPlayfield = 3) And (BallsLocked = 3) Then EndOfBall()
        If (BallsOnPlayfield = 1) And (BallsLocked = 0) Then EndMultiBall()
        If (BallsOnPlayfield = 3) And (BallsLocked = 0) Then PlaysoundAtVol "Fiodie", ActiveBall, 1
        If (BallsOnPlayfield = 2) And (BallsLocked = 0) Then PlaysoundAtVol "Eridie", ActiveBall, 1
        If (BallsOnPlayfield = 1) And (BallsLocked = 0) Then PlaysoundAtVol "Tarmadie", ActiveBall, 1
      End If
    Else
      If Not(FiveBallOverTimer.Enabled) Then
        If BallsOnPlayfield = 0 Then
          ' Terminate the final mission (mission all over).
        ' LookAtBackbox()
        ' PlayMusic 1, "Msmissionallover", FALSE
        ' PlayMusic 4, "missionallover", FALSE
          MissionAllOverTimer.Enabled = TRUE
          FinalMissionShowTimer.Uservalue = 11
          FinalMissionShowTimer.Enabled = TRUE
          bPlayingMission = FALSE
          bPlayingFinalMission = FALSE
          bPlayingMission1 = FALSE
          NextMessage = 9
          bCanDisplayMessage = TRUE
          BallsLocked = 0
                    DisplayB2SText "        MISSION ALL OVER        "
          D1.Text = "        MISSION ALL OVER        "
        End If
      End If
    End If
  End If
End Sub



' A ball may of rolled into the plunger lane kicker, if so, kick it back out again.
'
' *** DOCAM OPTIMIZED ***
'
Sub PlungerKicker_Hit()
  PlungerKickerSolenoidPulse()
End Sub



' The kicker 1 (bottom left side) has been hit.
'
' *** DOCAM OPTIMIZED ****
'
Sub Kicker1_Hit()
  startB2S(3)
  If bBallLockedInKicker1 Then
    ' Prevent the _Hit() event when a ball is already locked (when a locked ball is hit by another ball, this generate _Hit event !).
    Exit Sub
  Else
    If Tilted Then
      ' Eject the ball if the table is tilted.
      If bBallInKicker1 And Not(bBallLockedInKicker1) Then
        Kicker1SolenoidPulse()
        bBallInKicker1 = FALSE
        Exit Sub
      End If
    End If
    bBallInKicker1 = TRUE
    AddScore(5000)
    If bFinalMissionIsReady Then
      StartMission()
    Else
      If bLockIsLit Then
        LockBall()
        If bMissionIsLit Then StartMissionSoon()
      Else
        If bMissionIsLit And Not(bMultiBallMode) Then StartMission()
        Kicker1Timer.Enabled = TRUE
      End If
    End If
    If bPlayingFinalMission Then AddScore 20000
  End If
End Sub

Sub Kicker1SolenoidPulse()
    Kicker1.kick 170, 15
    PlaysoundAtVol "fx_kicker", kicker1, VolKick
End Sub


' The kicker 2 (bottom right side) has been hit.
'
' *** DOCAM WORKING ****
'
Sub Kicker2_Hit()
  startB2S(2)
  If bBallLockedInKicker2 Then
    ' Prevent the _Hit() event when a ball is already locked (when a locked ball is hit by another ball, this generate _Hit event !).
    Exit Sub
  Else
    If Tilted Then
      ' Eject the ball if the table is tilted.
      If bBallInKicker2 And Not(bBallLockedInKicker2) Then
        Kicker2SolenoidPulse()
        bBallInKicker2 = FALSE
        Exit Sub
      End If
    End If
    bBallInKicker2 = TRUE
    AddScore(5000)
    If bFinalMissionIsReady Then
      StartMission()
    Else
      If bLockIsLit Then
        LockBall()
        If bMissionIsLit Then StartMissionSoon()
      Else
        If bMissionIsLit And Not(bMultiBallMode) Then StartMission()
        Kicker2Timer.Enabled = TRUE
      End If
    End If
    If bPlayingFinalMission Then AddScore(20000)
  End If
End Sub


Sub Kicker2SolenoidPulse()
    Kicker2.kick 220, 15
    PlaysoundAtVol "fx_kicker", kicker2, VolKick
End Sub

' The kicker 3 (center) has been hit.
'
' *** DOCAM WORKING ****
'
Sub Kicker3_Hit()
  startB2S(5)
  If bBallLockedInKicker3 Then
    ' Prevent the _Hit() event when a ball is already locked (when a locked ball is hit by another ball, this generate _Hit event !).
    Exit Sub
  Else
    If Tilted Then
      ' Eject the ball if the table is tilted.
      If bBallInKicker3 And Not(bBallLockedInKicker3) Then
        Kicker3SolenoidPulse()
        bBallInKicker3 = FALSE
        Exit Sub
      End If
    End If
    bBallInKicker3 = TRUE
    AddScore(5000)
    If bFinalMissionIsReady Then
      StartMission()
    Else
      If bLockIsLit Then
        LockBall()
        If bMissionIsLit Then StartMissionSoon()
      Else
        If bMissionIsLit And Not(bMultiBallMode) Then StartMission()
        Kicker3Timer.Enabled = TRUE
      End If
    End If
    If bPlayingFinalMission Then AddScore(20000)
  End If
End Sub

Sub Kicker3SolenoidPulse()
    Kicker3.kick 165, 15
    PlaysoundAtVol "fx_kicker", kicker3, VolKick
End Sub

' The kicker 4 (top of the table) has been hit.
'
' *** DOCAM OPTIMIZED ****
'

Sub Kicker4_Hit()
  startB2S(4)
  If bBallLockedInKicker4 Then
    ' Prevent the _Hit() event when a ball is already locked (when a locked ball is hit by another ball, this generate _Hit event !).
    Exit Sub
   Else
    ' Eject the ball if the table is tilted.
    If Tilted Then
      If bBallInKicker4 And Not(bBallLockedInKicker4) Then
        Kicker4SolenoidPulse()
        bBallInKicker4 = FALSE
        Exit Sub
      End If
    End If
    bBallInKicker4 = TRUE
    AddScore(5000)
    If bFinalMissionIsReady Then
      StartMission()
    Else
      If bLockIsLit Then
        LockBall()
      Else
        Kicker4Timer.Enabled = TRUE
      End If
    End If
    If bPlayingFinalMission Then AddScore(20000)
   End If
   If bJackpotIsLit Then
    ' Scoring jackpot.
    GIOff()
        FlashEffect 0
        FlashEffectMissionTimer.enabled = 0
    AddScore(nvR4 * 100000)
    JackpotLight.State = 1
    bJackpotScored = TRUE
    bJackpotIsLit = FALSE
    If Not(bPlayingFinalMission) Then StopMusic 1
    JackpotDisplayTimerUserData = 0
    JackpotDisplayTimer.Enabled = TRUE
  End If
End Sub

Sub Kicker4SolenoidPulse()
    Kicker4.kick 270, 25
    PlaysoundAtVol "fx_kicker", kicker4, VolKick
End Sub

' *********************************************************************
' **  BUMPERS                                                        **
' *********************************************************************



' The bumper 1 (top left) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Bumper1_Hit()
  startB2S(3)
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
   AddScore(1000)
  If bPlayingMission Then
    FlashForMs F9, 100, 50, 2
    DOF 144, DOFPulse
    If bPlayingMission3 Then
      RemainingJets = RemainingJets - 1
      CheckJets()
    Else
      'Playsound "gun9"
       playsoundAtVol SoundFXDOF("gun9",106,DOFPulse,DOFContactors), Bumper1, VolBump
    End If
   Else
    FlashForMs F9, 100, 50, 0
    DOF 144, DOFPulse
    'Playsound "gun9"
     playsound SoundFXDOF("gun9",106,DOFPulse,DOFContactors)
   End If
  If bPlayingFinalMission Then AddScore(5000)
End Sub



' The bumper 2 (top right) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Bumper2_Hit()
  startB2S(2)
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
   AddScore(1000)
  If bPlayingMission Then
    FlashForMs F10, 100, 50, 2
    DOF 142, DOFPulse
    If bPlayingMission3 Then
      RemainingJets = RemainingJets - 1
      CheckJets()
    Else
      'Playsound "gun8"
       playsoundAtVol SoundFXDOF("gun8",105,DOFPulse,DOFContactors), Bumper2, VolBump
    End If
   Else
    FlashForMs F10, 100, 50, 0
    DOF 142, DOFPulse
    'Playsound "gun8"
     playsoundAtVol SoundFXDOF("gun8",105,DOFPulse,DOFContactors), Bumper2, VolBump
   End If
  If bPlayingFinalMission Then AddScore(5000)
End Sub



' The bumper 2 (bottom) has been hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub Bumper3_Hit()
  startB2S(5)
  If Tilted Then Exit Sub ' Ignore this _Hit() event if the table is tilted.
   AddScore(1000)
  If bPlayingMission Then
    FlashForMs F6, 100, 50, 2
    DOF 143, DOFPulse
    If bPlayingMission3 Then
      RemainingJets = RemainingJets - 1
      CheckJets()
    Else
      'Playsound "gun9"
       playsoundAtVol SoundFXDOF("gun9",107,DOFPulse,DOFContactors), Bumper3, VolBump
    End If
   Else
    FlashForMs F6, 100, 50, 0
    DOF 143, DOFPulse
    'Playsound "gun9"
     playsoundAtVol SoundFXDOF("gun9",107,DOFPulse,DOFContactors), Bumper3, VolBump
   End If
  If bPlayingFinalMission Then AddScore(5000)
End Sub



' *********************************************************************
' **                                                                 **
' **  TIMER EVENTS                                                   **
' **                                                                 **
' *********************************************************************



' This timer is used to delay the start of a game to allow any attract sequence to
' complete. When it expires, it creates a ball for the player to start playing.
'
' *** DOCAM OPTIMIZED ***
'

Sub FirstBallDelayTimer_timer()
  ' Stop the timer.
  FirstBallDelayTimer.Enabled = FALSE
  ' Be sure the camera is correctly looking the table.
' LookAtPlayfield()
  ' Main theme.
   PlaySong "mu_MSMain"
  ' Turn on GI.
  GIOn()
  ' Set other flags and required lits.
  bPlayingMission = FALSE
  bMissionIsLit = FALSE
  bMissionSuspended = FALSE
  bLockIsLit = FALSE
  bFinalMissionIsReady = FALSE
  ' Turn on "mission ready" lites.
  LightMission()
  ' Enable the kickback.
  LightKickback()
  ' Raise drop targets.
  ResetSLUGTargets()
  ResetCaptiveTargets()
  ResetSideTargets()
  CurrentWeapon = 0
  Score(1) = 0
  ' Reset mission lights.
  Mission1Light.State = BulbBlink', "10", 150
  Mission2Light.State = BulbOff', "10", 150
  Mission3Light.State = BulbOff', "10", 150
  Mission4Light.State = BulbOff', "10", 150
  Mission5Light.State = BulbOff', "10", 150
  Mission6Light.State = BulbOff', "10", 150
  ' Raise the SLUG drop targets.
  ResetSLUGTargets()
  ' Reset mission.
  Mission = 0
  ' Raise WEAPON drop targets.
  bWallIsDown = TRUE
  ResetWEAPONTargets()
  bWallIsDown = FALSE
   Weapon1.State = BulbBlink', "10", 150
  ' Set the initial POW bonus (regardling table setup).
  If nvR3 = 3 Then
    ' If POW BONUS HOLD set to ALWAYS (bonus held until game over), bonus starts at 0, then kept during whole game.
    POWBonus = 0
    POWBonusBall = 0
  Else
    ' Otherwise, bonus starts at 1.
    POWBonus = 1
    POWBonusBall = 1
  End If
  ' Update score (at 0) and Ball in play.
  Ball = Ball + 1
  AddScore(0)
  ' Reset the table for a new ball.
  ResetForNewPlayerBall()
  ' Create a new ball in the shooters lane.
  CreateNewBall()
End Sub



' The tilt recovery timer waits for all the balls to drain before continuing on as per normal.
'
' *** DOCAM MUST BE CHECKED ***
'
Sub TiltRecoveryTimer_Timer()
  ' disable the timer.
  TiltRecoveryTimer.Enabled = FALSE
  If (BallsOnPlayfield - BallsLocked) = 0 Then
    ' Restore the general illumation (GI).
    GIOn()
    ' Restore the WEAPON drop targets state.
    FixWeapon()
    ' If all non locked ball(s) have been drained, then do the normal end of ball.
    EndOfBall()
  Else
    ' otherwise retry in another second (wait for locating ball(s)).
    TiltRecoveryTimer.Interval = 1000
    TiltRecoveryTimer.Enabled = TRUE
  End If
End Sub



' Called when the 'delay to protect' display is expired.
'
' *** DOCAM MUST BE CHECKED ***
'
Sub InterruptDisplayTimer_Timer()
  InterruptDisplayTimer.Enabled = FALSE
  bInterruptedDisplay = FALSE
  If bPlayingBall Then AddScore(0)
End Sub



' *** DOCAM MUST BE CHECKED ***
'
Sub EndOfBallTimer2_Timer()
  EndOfBallTimer2.Enabled = FALSE
  ' Only process any of this if the table is not tilted.
  ' The tilt recovery mechanism will handle any extra ball or end of game.
  If Not(Tilted) Then
    ' Table is not tilted: add in POW bonus points.
    POWBonusCount = POWBonus
    If POWBonusCount = 0 Then
      EndOfBallTimer.Interval = 3200
    Else
      EndOfBallTimer.Interval = 5000
    End If
    POWBonusTimer.Enabled = TRUE
  Else
    ' Table is tilted: no bonus, so move to the next state quickly.
    EndOfBallTimer.Interval = 20
    POWBonusCount = 0
  End If
  ' start the end of ball timer which allows you to add a delay at this point
  EndOfBallTimer.Enabled = TRUE
End Sub



' This which delays the machine to allow any bonus points to be added up has expired.
' Check to see if there are any extra ball for the player.
' if not, check to see if this was the last ball.
'
' *** DOCAM MUST BE CHECKED ***
'
Sub EndOfBallTimer_Timer()
  Dim PosInTable
  Dim Suffix
  ' disable the timer
  EndOfBallTimer.Enabled = FALSE
  ' if were tilted, reset the internal tilted flag (this will also
  ' set fpTiltWarnings back to zero) which is useful if we are changing player LOL
  Tilted = FALSE
  ' Has the player won an extra-ball ?
  If bExtraBallAward Then
    ' Reset the POW bonus if required (POW BONUS HOLD set to NEVER).
    If nvR3 = 1 Then
      POWBonus = 1
      POWBonusBall = 1
    Else
      POWBonusBall = 0
    End If
    ' Turn on Shoot Again light.
    ShootAgainLight.State = 1
    ' Display SHOOT AGAIN.
    ExtraBallShowTimer.Enabled = TRUE
       ' DMDBlink "", " ", "  *   S H O O T  A G A I N   *  ", 50, 20
    D1.Text = "  *   S H O O T  A G A I N   *  "
        DisplayB2SText "  *   S H O O T  A G A I N   *  "
'   D1.QueueText "  *   S H O O T  A G A I N   *  ", seBlinkMask, 1000
'   D2.QueueText "  *   S H O O T  A G A I N   *  ", seBlinkMask, 1000
    Playsound "cheers"
    CurrentWeapon = 0
    Weapon1.State = 2
    Weapon2.State = 0
    Weapon3.State = 0
    Weapon4.State = 0
    Weapon5.State = 0
    Weapon6.State = 0
    ExtraBall.State = 0
  Else
      Ball = Ball + 1
    If BallsPerGame - Ball < 0 Then
      ' It was the last ball.
      If Score(1) > HighScore(10) Then
        ' 10st score was beaten.
        PosInTable = 10
        ' Check the position in high score table.
        If Score(1) > HighScore(9) Then PosInTable = 9
        If Score(1) > HighScore(8) Then PosInTable = 8
        If Score(1) > HighScore(7) Then PosInTable = 7
        If Score(1) > HighScore(6) Then PosInTable = 6
        If Score(1) > HighScore(5) Then PosInTable = 5
        If Score(1) > HighScore(4) Then PosInTable = 4
        If Score(1) > HighScore(3) Then PosInTable = 3
        If Score(1) > HighScore(2) Then PosInTable = 2
        If Score(1) > HighScore(1) Then PosInTable = 1
        ' Define the suffix like st, nd, rd or th.
        If PosInTable = 1 Then
          Suffix = "ST"
        ElseIf PosInTable = 2 Then
          Suffix = "ND"
        ElseIf PosInTable = 3 Then
          Suffix = "RD"
        Else
          Suffix = "TH"
        End If
        ' Display final score.
        If Score(1) < 1000000 Then
          D1.Text = FormatScore(Score(1)) & String(10 - Len(PosInTable), " ") & "HIGH SCORE <" & PosInTable & Suffix & ">"
            DisplayB2SText (Score(1)) & String(10 - Len(PosInTable), " ") & "HIGH SCORE <" & PosInTable & Suffix & ">"
        Else
          D1.Text = FormatScore(Score(1)) & String(16 - Len(Score(1)) - Len(PosInTable), " ") & "HIGH SCORE <" & PosInTable & Suffix & ">"
            DisplayB2SText (Score(1)) & String(16 - Len(Score(1)) - Len(PosInTable), " ") & "HIGH SCORE <" & PosInTable & Suffix & ">"
        End If
        ' Submit the player score to the High Score system built into Future Pinball if required.
        'bEnteringAHighScore = TRUE
        ' Play the high score music.
        'Playsound "MShighscore"
        CheckHighscore()': hsbModeActive = True
      Else
        ' Player's score doesn't fit in the "top ten".
        If Score(1) < 1000000 Then
          D1.Text = FormatScore(Score(1)) & "                GAME OVER"
                    DisplayB2SText (Score(1)) & "                GAME OVER"
        Else
          D1.Text = FormatScore(Score(1)) & String(22 - Len(Score(1)), " ") & "GAME OVER"
            DisplayB2SText (Score(1)) & String(22 - Len(Score(1)), " ") & "GAME OVER"
        End If
        bEnteringAHighScore = FALSE
        ' Play the game over music.
        Playsound "MSgameover"
        ' Delay the game over during 7 seconds.
        GameOverTimer.Enabled = TRUE
      End If
    Else
      ' Reset the POW bonus to 1 if required (POW BONUS HOLD set to NEVER or EXTRA-BALL from the setup menu).
      If nvR3 <> 3 Then
        POWBonus = 1
        POWBonusBall = 1
      Else
        POWBonusBall = 0
      End If
      ' It is not the last ball.
      EndOfBallComplete()
    End If
  End If
End Sub



' This timer is used to add a 7 seconds delay before game over state (when no entry in high score table).
'
' *** DOCAM OPTIMIZED ***
'
Sub GameOverTimer_Timer()
    StopSound Song:Song = ""
  GameOverTimer.Enabled = FALSE
  EndOfBallComplete()
End Sub



' This timer disable the kickback mechanism.
'
' *** DOCAM OPTIMIZED ***
'

Sub KickbackTimer_timer()
  KickbackTimer.Enabled = FALSE
  DOF 124, DOFPulse
  If KickbackTimer.Uservalue = 1 Then
    ' Play rocket explosion sound.
    Playsound "rocket2"
    ' Add 2000 points.
    AddScore(2000)
    ' Unlit kickback orange light.
    bKickbackIsLit = FALSE
    KickbackLight.State = 0
    ' Open the left outlane.
    KickbackGuide.isDropped = TRUE
    ' Eject the ball.
        KickbackPlungerSolenoidPulse()
  End If
End Sub


Sub KickbackPlungerSolenoidPulse()
    LeftOutLaneTrigger.kick 0, 40
    LaserKickP1.TransY = 90
    vpmtimer.addtimer 800, "LaserKickP1.TransY = 0 '"
    Playsound "bumper_retro"
End sub

' This timer handles the regular ball saver mechanism.
'
' *** DOCAM OPTIMIZED ***
'
Sub BallSaverClockTimer_Timer()
  BallSaverClockTimer.Enabled = FALSE
  If BallSaverClockTimer.Uservalue = 1 Then
    ' At this point: remain 10 seconds.
    BallSaverClockTimer.Uservalue = 2
    BallSaverClockTimer.Interval = 5000
    ShootAgainLight.BlinkInterval = 100
    ShootAgainLight.State = BulbBlink
  ElseIf BallSaverClockTimer.Uservalue = 2 Then
    ' At this point: remain 5 seconds.
    BallSaverClockTimer.Uservalue = 3
    BallSaverClockTimer.Interval = 5000
    ShootAgainLight.BlinkInterval = 50
    ShootAgainLight.State = BulbBlink
  ElseIf BallSaverClockTimer.Uservalue = 3 Then
    ' Ball saver delay has been expired.
    BallSaverClockTimer.Uservalue = 0
    ShootAgainLight.State = BulbOff
    bBallSaverIsActive = FALSE
    ShootAgainLight.BlinkInterval = 150
  End If


End Sub



' This timer is used to display the POW bonus countdown sequence.
'
' *** DOCAM OPTIMIZED ***
'
Dim POWBonusTimerUserData
Sub POWBonusTimer_Timer()
  Dim lenPOWBonus
  Dim lenPOWBase
  Dim lenPOWScore
  POWBonusTimer.Uservalue = POWBonusTimer.Uservalue + 1
  If POWBonusCount = 0 Then
    Select Case POWBonusTimer.Uservalue
      Case 12:
        'D1.FlushQueue()
        'D2.FlushQueue()
                DisplayB2SText String(32, " ")
        D1.Text = String(32, " ")
        'D1.UpdateInterval = 20
        'D2.UpdateInterval = 20
        D1.Text = "          NO POW BONUS          "', seScrollIn, 1000, 0, TRUE
                DisplayB2SText "          NO POW BONUS          "
      Case 14:
        Playsound "cmonboy"
      Case 22:
        'D1.FlushQueue()
        'D2.FlushQueue()
        'D1.UpdateInterval = 10
        'D2.UpdateInterval = 10
        POWBonusTimer.Enabled = FALSE
        POWBonusTimer.Uservalue = 0
        POWBonusCount = 0
    End Select
  Else
    Select Case POWBonusTimer.Uservalue
      Case 8:
        lenPOWBonus = Len(POWBonus)
        lenPOWBase = Len(1000 * nvR2)
        lenPOWScore = Len(1000 * nvR2 * POWBonus)
                DisplayB2SText (POWBonus) & " POW     " & String(lenPOWBase, " ") & "     " & String(lenPOWScore, " ") & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        D1.Text = FormatScore(POWBonus) & " POW     " & String(lenPOWBase, " ") & "     " & String(lenPOWScore, " ") & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        Playsound "Thankyou"
      Case 15:
        lenPOWBonus = Len(POWBonus)
        lenPOWBase = Len(1000 * nvR2)
        lenPOWScore = Len(1000 * nvR2 * POWBonus)
                DisplayB2SText (POWBonus) & " POW  X  " & String(lenPOWBase, " ") & "     " & String(lenPOWScore, " ") & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        D1.Text = FormatScore(POWBonus) & " POW  X  " & String(lenPOWBase, " ") & "     " & String(lenPOWScore, " ") & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        Playsound "Steel"
      Case 20:
        lenPOWBonus = Len(POWBonus)
        lenPOWBase = Len(1000 * nvR2)
        lenPOWScore = Len(1000 * nvR2 * POWBonus)
                DisplayB2SText (POWBonus) & " POW  X  " & FormatScore(nvR2 * 1000) & "     " & String(lenPOWScore, " ") & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        D1.Text = FormatScore(POWBonus) & " POW  X  " & FormatScore(nvR2 * 1000) & "     " & String(lenPOWScore, " ") & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        Playsound "Steel"
      Case 25:
        lenPOWBonus = Len(POWBonus)
        lenPOWBase = Len(1000 * nvR2)
        lenPOWScore = Len(1000 * nvR2 * POWBonus)
                DisplayB2SText (POWBonus) & " POW  X  " & FormatScore(nvR2 * 1000) & "  =  " & String(lenPOWScore, " ") & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        D1.Text = FormatScore(POWBonus) & " POW  X  " & FormatScore(nvR2 * 1000) & "  =  " & String(lenPOWScore, " ") & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        Playsound "Steel"
      Case 30:
        lenPOWBonus = Len(POWBonus)
        lenPOWBase = Len(1000 * nvR2)
        lenPOWScore = Len(1000 * nvR2 * POWBonus)
                DisplayB2SText (POWBonus) & " POW  X  " & FormatScore(nvR2 * 1000) & "  =  " & FormatScore(1000 * nvR2 * POWBonus) & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        D1.Text = FormatScore(POWBonus) & " POW  X  " & FormatScore(nvR2 * 1000) & "  =  " & FormatScore(1000 * nvR2 * POWBonus) & String(Int((18 - lenPOWBonus - lenPOWBase - lenPOWScore) / 2), " ")
        If POWBonusBall < 10 Then
          Playsound "Ok"
        Else
          Playsound "youregreat"
        End If
      Case 42:
        POWBonusTimer.Enabled = FALSE
        POWBonusTimer.Uservalue = 0
        Score(1) = Score(1) + (1000 * nvR2 * POWBonus)
    End Select
  End If
End Sub



' This timer is called every second (like a stopwatch) during missions 1 and 2.
'
' *** DOCAM OPTIMIZED ***
'
Sub MissionTimer_Timer()
  MissionTimerUserData = MissionTimerUserData - 1
  If MissionTimerUserData = 0 Then
    MissionTimer.Enabled = FALSE
    CompleteMission()
  End If
  If bPlayingMission1 Then
    If MissionTimerUserData < 32 then
      NextMessage = 2
      bCanDisplayMessage = TRUE
      ShowNextMessage()
    End If
  ElseIf bPlayingMission2 Then
    If MissionTimerUserData < 32 Then
      NextMessage = 4
      bCanDisplayMessage = TRUE
      ShowNextMessage()
    End If
  End If
End Sub



' This timer is called to animate WEAPON targets while playing mission 2 or mission 5.
'
' *** DOCAM OPTIMIZED ***
'
Dim AnimateWeaponTimerUserData
Sub AnimateWeaponTimer_Timer()
  AnimateWeaponTimer.Uservalue = AnimateWeaponTimer.Uservalue + 1
  Select Case Mission
    Case 2:
      Select Case AnimateWeaponTimer.Uservalue
        Case 1:
          W1.Isdropped = 0
          W6.Isdropped = 1
                    W1Light.state = 2
        Case 2:
          W2.Isdropped = 0
          W1.Isdropped = 1
                    W1Light.state = 0
                    W2Light.state = 2
        Case 3:
          W3.Isdropped = 0
          W2.Isdropped = 1
                    W2Light.state = 0
                    W3Light.state = 2
        Case 4:
          W4.Isdropped = 0
          W3.Isdropped = 1
                    W3Light.state = 0
                    W4Light.state = 2
        Case 5:
          W5.Isdropped = 0
          W4.Isdropped = 1
                    W4Light.state = 0
                    W5Light.state = 2
        Case 6:
          W6.Isdropped = 0
          W5.Isdropped = 1
                    W5Light.state = 0
                    W6Light.state = 2
        Case 7:
          W1.Isdropped = 0
          W6.Isdropped = 1
                    W6Light.state = 0
                    W1Light.state = 2
        Case 8:
          W3.Isdropped = 0
          W1.Isdropped = 1
                    W1Light.state = 0
                    W3Light.state = 2
        Case 9:
          W5.Isdropped = 0
          W3.Isdropped = 1
                    W3Light.state = 0
                    W5Light.state = 2
        Case 10:
          W2.Isdropped = 0
          W5.Isdropped = 1
                    W5Light.state = 0
                    W2Light.state = 2
        Case 11:
          W4.Isdropped = 0
          W2.Isdropped = 1
                    W2Light.state = 0
                    W4Light.state = 2
        Case 12:
          W6.Isdropped = 0
          W4.Isdropped = 1
                    W4Light.state = 0
                    W6Light.state = 2
          AnimateWeaponTimer.Uservalue = 0
      End Select
    Case 5:
      Select Case AnimateWeaponTimer.Uservalue
        Case 1:
          W1.Isdropped = 0
          W6.Isdropped = 1
                    W6Light.state = 0
                    W1Light.state = 2
          Playsound "allenlaugh"
        Case 2:
          W6.Isdropped = 0
          W1.Isdropped = 1
                    W1Light.state = 0
                    W6Light.state = 2
          Playsound "allenlaugh"
        Case 3:
          W2.Isdropped = 0
          W6.Isdropped = 1
                    W6Light.state = 0
                    W2Light.state = 2
          Playsound "cmonboy"
        Case 4:
          W5.Isdropped = 0
          W2.Isdropped = 1
                    W2Light.state = 0
                    W5Light.state = 2
          Playsound "gohometomommy"
        Case 5:
          W3.Isdropped = 0
          W5.Isdropped = 1
                    W3Light.state = 0
                    W5Light.state = 2
          Playsound "allenlaugh"
        Case 6:
          W4.Isdropped = 0
          W3.Isdropped = 1
                    W4Light.state = 0
                    W3Light.state = 2
          Playsound "cmonboy"
        Case 7:
          W3.Isdropped = 0
          W4.Isdropped = 1
                    W4Light.state = 0
                    W3Light.state = 2
          Playsound "youremincemeat"
        Case 8:
          W5.Isdropped = 0
          W3.Isdropped = 1
                    W3Light.state = 0
                    W5Light.state = 2
          Playsound "cmonboy"
        Case 9:
          W2.Isdropped = 0
          W5.Isdropped = 1
                    W5Light.state = 0
                    W2Light.state = 2
          Playsound "allenlaugh"
        Case 10:
          W6.Isdropped = 0
          W2.Isdropped = 1
                    W2Light.state = 0
                    W6Light.state = 2
          Playsound "gohometomommy"
          AnimateWeaponTimer.Uservalue = 0
      End Select
  End Select
End Sub





' *** DOCAM OPTIMIZED ***
'
Sub PlayfieldTimer_Timer()
  PlayfieldTimer.Enabled = FALSE
' LookAtPlayfield()
End Sub



' *** DOCAM OPTIMIZED ***
'
Sub ExtraBallShowTimer_Timer()
  ExtraBallShowTimer.Enabled = FALSE
  bExtraBallAward = FALSE
  EndOfBallComplete()
End Sub



' This timer is used to autoplung a ball.
'
' *** DOCAM OPTIMIZED ***
'
Sub AutoBallTimer_Timer()
  AutoBallTimer.Enabled = FALSE
  PlungerSolenoidPulse.Interval = 2000
    PlungerSolenoidPulse.enabled = 1
  Playsound "gun6"
  FlashAnimate3()
End Sub

Sub PlungerSolenoidPulse_Timer
    PlungerSolenoidPulse.enabled = 0
    plungerIM.AutoFire
  DOF 125, DOFPulse
  DOF 114, DOFPulse
End Sub

' *** DOCAM OPTIMIZED ****
'
Sub Kicker1Timer_Timer()
  Kicker1Timer.Enabled = FALSE
  Kicker1SolenoidPulse()
  bBallInKicker1 = FALSE
  L1.State = 0
  Playsound "Rocket2"
End Sub



' *** DOCAM OPTIMIZED ****
'
Sub Kicker2Timer_Timer()
  Kicker2Timer.Enabled = FALSE
  Kicker2SolenoidPulse()
  bBallInKicker2 = FALSE
  L2.State = 0
  Playsound "Rocket2"
End Sub



' *** DOCAM OPTIMIZED ****
'
Sub Kicker3Timer_Timer()
  Kicker3Timer.Enabled = FALSE
  Kicker3SolenoidPulse()
  bBallInKicker3 = FALSE
  L3.State = 0
  Playsound "Rocket2"
End Sub



' *** DOCAM OPTIMIZED ****
'
Sub Kicker4Timer_Timer()
  Kicker4Timer.Enabled = FALSE
  Kicker4SolenoidPulse()
  bBallInKicker4 = FALSE
  L4.State = 0
  Playsound "Rocket2"
End Sub



' *** DOCAM OPTIMIZED ****
'
Sub KickOutBallTimer_Timer()
  KickOutBallTimer.Enabled = FALSE
  If Not(bLockIsLit) Then
    If bBallInKicker1 Then Kicker1Timer.Enabled = TRUE
    If bBallInKicker2 Then Kicker2Timer.Enabled = TRUE
    If bBallInKicker3 Then Kicker3Timer.Enabled = TRUE
    If bBallInKicker4 Then Kicker4Timer.Enabled = TRUE
  End If
End Sub



' This timer is called by FuturePinball_BeginPlay (raising C1/C2/C3 drop targets after captive ball creation/ejection).
'
' *** DOCAM OPTIMIZED ***
'
Sub FixTargetTimer_Timer()
  ' Disable this timer.
  FixTargetTimer.Enabled = FALSE
  ' Raise the three drop targets behind the captive ball.
  C1.Isdropped = 0
  C2.Isdropped = 0
  C3.Isdropped = 0
End Sub



' This timer checks the METALSLUG drop targets, doing actions if all targets are dropped.
'
' *** DOCAM OPTIMIZED ***
'
Sub CheckSLUGTimer_Timer()
  CheckSLUGTimer.Enabled = FALSE
  If Slug1.IsDropped = 1 And Slug2.IsDropped = 1 And Slug3.IsDropped = 1 And Slug4.IsDropped = 1 And Slug5.IsDropped = 1 And Slug6.IsDropped = 1 And Slug7.IsDropped = 1 And Slug8.IsDropped = 1 And Slug9.IsDropped = 1 Then
    ResetSLUGTargets()
    AddScore(20000)
    LightKickback()
    If Not(bLockIsLit) And Not(bMultiBallMode) And Not(bPlayingFinalMission) Then
      FlashEffectMissionTimer.enabled = 1'GI_Flash()
      InterruptDisplay(1200)
    ' D1.FlushQueue()
    ' D2.FlushQueue()
            DisplayB2SText "          LOCK IS LIT           "
      D1.Text = "          LOCK IS LIT           "
      'D1.QueueText "          LOCK IS LIT           "', seBlinkMask, 400
      'D2.QueueText "          LOCK IS LIT           "', seBlinkMask, 400
      FlashAnimate3()
      Playsound "explode5"
      bLockIsLit = TRUE
      'L1.state = 2
      'L2.state = 2
      'L3.state = 2
      'L4.state = 2
      L1.State = 2
      L2.State = 2
      L3.State = 2
      L4.State = 2
      Marco.State = 0
      Tarma.State = 0
      Eri.State = 0
      Fio.State = 0
      If Not(bPlayingMission) Then Playsound "MSlocklit"
    End If
    If bBallLockedInKicker1 And bBallLockedInKicker2 And bBallLockedInKicker3 And bMissionIsLit And (Mission < 5) Then StartMission(): FlashEffectMissionTimer.enabled = 0
  End If
End Sub



' *** DOCAM OPTIMIZED ***
'
Dim FiveBallOverTimerUserData
Dim FinalMissionShowTimerUserData
Sub FinalMissionShowTimer_Timer()
  FinalMissionShowTimer.Uservalue = FinalMissionShowTimer.Uservalue + 1
  Select Case FinalMissionShowTimer.Uservalue
    Case 1:
            DisplayB2SText "   STARTING 5-BALL MULTIBALL    "
      D1.Text = "   STARTING 5-BALL MULTIBALL    "
      Playsound "explode3"
      ExplodeAnimate()
      ' Cancels regular ball saver.
      bBallSaverIsActive = FALSE
      BallSaverClockTimer.Enabled = FALSE
      BallSaverClockTimer.Uservalue = 1
      ShootAgainLight.State = 0
      ShootAgainLight.BlinkInterval = 150
    Case 2:
            DisplayB2SText " DROP TARGETS SCORE 20,000 EACH  "
      D1.Text = " DROP TARGETS SCORE 20,000 EACH  "
      Playsound "steel"
      ' Destroy any ball in kickers.
      Kicker1SolenoidPulse
      Kicker2SolenoidPulse()
      Kicker3SolenoidPulse()
      Kicker4SolenoidPulse()
    Case 3:
            DisplayB2SText "      POW BONUS IS DOUBLED      "
      D1.Text = "      POW BONUS IS DOUBLED      "
      Playsound "thankyou"
    Case 4:
            DisplayB2SText "    KICKERS SCORE 25,000 EACH    "
      D1.Text = "    KICKERS SCORE 25,000 EACH    "
      Playsound "rocket2"
    Case 5:
            DisplayB2SText "       CAPTIVE BALL IS LIT      "
      D1.Text = "       CAPTIVE BALL IS LIT      "
      Playsound "ok"
    Case 6:
            DisplayB2SText "       JET BUMPERS ARE LIT      "
      D1.Text = "       JET BUMPERS ARE LIT      "
      Playsound "mordenlaugh"
    Case 7:
            DisplayB2SText "         JACKPOT IS LIT         "
      D1.Text = "         JACKPOT IS LIT         "
      Playsound "yell6"
      ' Create locked balls in kickers.
      bBallLockedInKicker1 = TRUE
      bBallInKicker1 = TRUE
      Kicker1SolenoidPulse()
      L1.State = 1
      bBallLockedInKicker2 = TRUE
      bBallInKicker2 = TRUE
      Kicker2.CreateBall()
      L2.State = 1
      bBallLockedInKicker3 = TRUE
      bBallInKicker3 = TRUE
      Kicker3.CreateBall()
      L3.State = 1
      bBallLockedInKicker4 = TRUE
      bBallInKicker4 = TRUE
      Kicker4.CreateBall()
      L4.State = 1
      ' Now they're four locked balls.
      BallsOnPlayfield = 4
      BallsLocked = 4
      ' Drop the side targets.
      S1.Isdropped = 1
      S1Light.State = 1
      S2.Isdropped = 1
      S2Light.State = 1
      S3.Isdropped = 1
      S3Light.State = 1
    Case 8:
            DisplayB2SText "        EVERYTHING IS LIT       "
      D1.Text = "        EVERYTHING IS LIT       "
    ' D1.QueueText "        EVERYTHING IS LIT       ", seBlinkMask, 1000
    ' D2.QueueText "        EVERYTHING IS LIT       ", seBlinkMask, 1000
      Playsound "explode3"
      ExplodeAnimate()
    Case 10:
      FinalMissionShowTimer.Enabled = FALSE
      FinalMissionShowTimer.Uservalue = 0
    ' LookAtPlayfield()
      AutoBall()
      CheckFiveBallTimer.interval = 3000
      ReleaseBallsTimer.Enabled = TRUE
      FiveBallOverTimer.Uservalue = 0
      ' Regardling table setup, ball saver duration is 2 x regular ball saver (except if 15 seconds: 45 seconds).
      Select Case nvR1
        Case 15: FiveBallOverTimer.Interval = 25000 ' 45 seconds
        Case 30: FiveBallOverTimer.Interval = 40000 ' 2 x 30 - 20 = 60 seconds.
        Case 45: FiveBallOverTimer.Interval = 70000 ' 2 x 45 - 20 = 90 seconds.
        Case 60: FiveBallOverTimer.Interval = 100000 ' 2 x 60 - 20 = 120 seconds.
      End Select
      ShootAgainLight.State = 2
      LightJackpot()
    Case 12:
      ' Mission all over.
            DisplayB2SText "        MISSION ALL OVER        "
      D1.Text = "        MISSION ALL OVER        "
    Case 13:
      FinalMissionShowTimer.Interval = 700
            DisplayB2SText "     MARCO                      "
      D1.Text = "     MARCO                      "
      Playsound "steel"
    Case 14:
            DisplayB2SText "     MARCO, TARMA                "
      D1.Text = "     MARCO, TARMA                "
      Playsound "steel"
    Case 15:
            DisplayB2SText "     MARCO, TARMA, ERI            "
      D1.Text = "     MARCO, TARMA, ERI            "
      Playsound "steel"
    Case 16:
            DisplayB2SText "     MARCO, TARMA, ERI & FIO      "
      D1.Text = "     MARCO, TARMA, ERI & FIO      "
      Playsound "steel"
    Case 17:
      FinalMissionShowTimer.Interval = 2300
            DisplayB2SText "       HAVE DONE IT AGAIN       "
      D1.Text = "       HAVE DONE IT AGAIN       "
    ' D1.QueueText "       HAVE DONE IT AGAIN       ", seBlinkMask, 2300
    ' D2.QueueText "       HAVE DONE IT AGAIN       ", seBlinkMask, 2300
      Playsound "explode3"
    Case 18:
            DisplayB2SText "      SEE YOU NEXT MISSION      "
      D1.Text = "      SEE YOU NEXT MISSION      "
    ' D1.QueueText "      SEE YOU NEXT MISSION      ", seBlinkMask, 2300
    ' D2.QueueText "      SEE YOU NEXT MISSION      ", seBlinkMask, 2300
      Playsound "metalslug1"
    Case 19:
            DisplayB2SText "      SEE YOU NEXT MISSION      "
      D1.Text = "      SEE YOU NEXT MISSION      "
      FinalMissionShowTimer.Enabled = FALSE
      FinalMissionShowTimer.Uservalue = 0
  End Select
End Sub



' *** DOCAM OPTIMIZED ****
'
Sub StartMissionSoonTimer_Timer()
  StartMissionSoonTimer.Enabled = FALSE
  StartMission()
End Sub



' During 5-ball play (final mission) this timer checks they're five balls on playfield, until the 45 seconds timer over.
' Called every second while they're 5 balls on playfield.
' Called every 3 seconds if they're are less than 5 balls.
'
' *** DOCAM OPTIMIZED ***
'
Sub CheckFiveBallTimer_Timer()
  CheckFiveBallTimer.Enabled = FALSE
  If BallsOnPlayfield = 5 Then
    CheckFiveBallTimer.interval = 1000
  Else
    CheckFiveBallTimer.interval = 3000
  End If
  If BallsOnPlayfield < 5 Then AutoBall()
End Sub



' Called during a total of 45 seconds (from final mission start), this timer works as dedicated ball saver for final mission.
' Called first time after 25 seconds, then 10 seconds, then two times 5 seconds. At every call, the 'Shoot Again' blinks more faster.
'
' *** DOCAM OPTIMIZED ***
'
Sub FiveBallOverTimer_Timer()
  FiveBallOverTimer.Uservalue = FiveBallOverTimer.Uservalue + 1
  Select Case FiveBallOverTimer.Uservalue
    Case 1:
      ' At this point, remains 20 seconds.
      ShootAgainLight.BlinkInterval = 100
      ShootAgainLight.State = 2
      FiveBallOverTimer.Enabled = FALSE
      FiveBallOverTimer.interval = 10000
    Case 2:
      ' At this point, remains 10 seconds.
      ShootAgainLight.BlinkInterval = 50
      ShootAgainLight.State = 2
      FiveBallOverTimer.Enabled = FALSE
      FiveBallOverTimer.interval = 5000
    Case 3:
      ' At this point, remains 5 seconds.
      ShootAgainLight.BlinkInterval = 25
      ShootAgainLight.State = 2
      FiveBallOverTimer.Enabled = FALSE
      FiveBallOverTimer.interval = 5000
    Case 4:
      ' At this point, cancels the 5-ball mode (disabling ball saver feature).
      FiveBallOverTimer.Enabled = FALSE
      CheckFiveBallTimer.Enabled = FALSE
      ShootAgainLight.State = 0
      ShootAgainLight.BlinkInterval = 150
      FiveBallOverTimer.Uservalue = 0
  End Select
End Sub



' This timer is used to play the correct theme.
'
' *** DOCAM OPTIMIZED ***
'
Sub ResetMusicTimer_Timer()
  ResetMusicTimer.Enabled = FALSE
' Effectmusic 1, FadeVolume, 1.0, 5000
  ResetMusic()
End Sub



' This timer give some delay before turning off flashers.
'
' *** DOCAM OPTIMIZED ***
'
Sub FlashOffTimer_Timer()
  FlashOffTimer.Enabled = FALSE
    FlashEffect 0
    FlashEffectMissionTimer.enabled = 0
  FlashOff()
End Sub



' This timer displays messages at the start of regular 4-ball multi ball.
'
' *** DOCAM OPTIMIZED ***
'
Dim MultiballDisplayTimerUserData
Sub MultiballDisplayTimer_Timer()
  MultiballDisplayTimerUserData = MultiballDisplayTimerUserData + 1
  Select Case MultiballDisplayTimerUserData
    Case 1:
            DisplayB2SText "M U L T I                B A L L"
      D1.Text = "M U L T I                B A L L"
    ' D1.FlushQueue()
    ' D2.FlushQueue()
    Case 2:
            DisplayB2SText " M U L T I              B A L L "
      D1.Text = " M U L T I              B A L L "
    Case 3:
            DisplayB2SText "  M U L T I            B A L L  "
      D1.Text = "  M U L T I            B A L L  "
    Case 4:
            DisplayB2SText "   M U L T I          B A L L   "
      D1.Text = "   M U L T I          B A L L   "
    Case 5:
            DisplayB2SText "    M U L T I        B A L L    "
      D1.Text = "    M U L T I        B A L L    "
    Case 6:
            DisplayB2SText "     M U L T I      B A L L     "
      D1.Text = "     M U L T I      B A L L     "
    Case 7:
            DisplayB2SText "      M U L T I    B A L L      "
      D1.Text = "      M U L T I    B A L L      "
    Case 8:
            DisplayB2SText "       M U L T I  B A L L       "
      D1.Text = "       M U L T I  B A L L       "
    Case 9:
            DisplayB2SText "  *    M U L T I  B A L L    *  "
      D1.Text = "  *    M U L T I  B A L L    *  "
    Case 10:
            DisplayB2SText "  *    M U L T I  B A L L    *  "
      D1.Text = "  *    M U L T I  B A L L    *  "', seBlinkMask, 1000
    Case 14:
      MultiballDisplayTimer.Enabled = FALSE
      MultiballDisplayTimerUserData = 0
  End Select
End Sub



' This timer is used to start multi ball play.
'
' *** DOCAM OPTIMIZED ***
'
Sub StartMultiballTimer_Timer()
  StartMultiballTimer.Enabled = FALSE
  If Not(bFinalMissionIsReady) Then
    BallsLocked = 0
    FlashEffectMissionTimer.enabled = 1'GI_Flash()
    If bPlayingMission Then
      ' If required, suspend the current mission (will be resumed at the end of multiball).
      bPlayingMission = FALSE
      bMissionSuspended = TRUE
      MissionTimer.Enabled = FALSE
      Select Case Mission
        Case 2:
          AnimateWeaponTimer.Enabled = FALSE
          DropWeapon()
        Case 3:
          Bumper1L.State = 1
          Bumper2l.State = 1
          Bumper3L.State = 1
        Case 4:
          CaptiveLight.State = 0
        Case 5
          AnimateWeaponTimer.Enabled = FALSE
          DropWeapon()
      End Select
    End If
    MultiballDisplayTimerUserData = 0
    MultiballDisplayTimer.Enabled = TRUE
    bMultiBallMode = TRUE
    bLockIsLit = FALSE
    StopSound Song:Song = "": PlaySong "mu_MSmultiball"
    Playsound "explode5"
    FlashAnimate6()
    ReleaseBallsTimer.Enabled = TRUE
    LightJackpot()
    DropWeapon()
    M1.State = 0
    M2.State = 0
    M3.State = 0
  Else
    StartMission()
  End If
End Sub



' This timer is called at the start of multi ball modes (regular and 5-ball) to eject balls from kickers.
'
' *** DOCAM OPTIMIZED ***
'
Dim ReleaseBallsTimerUserData
Sub ReleaseBallsTimer_Timer()
  ReleaseBallsTimerUserData = ReleaseBallsTimerUserData + 1
  Select Case ReleaseBallsTimerUserData
    Case 1:
      Kicker1Timer.Enabled = TRUE
      bBallLockedInKicker1 = FALSE
    Case 2:
      Kicker2Timer.Enabled = TRUE
      bBallLockedInKicker2 = FALSE
    Case 3:
      Kicker3Timer.Enabled = TRUE
      bBallLockedInKicker3 = FALSE
    Case 4:
      Kicker4Timer.Enabled = TRUE
      bBallLockedInKicker4 = FALSE
      ReleaseBallsTimer.Enabled = FALSE
      ReleaseBallsTimerUserData = 0
      MultiballDisplayTimer.Enabled = FALSE
  End Select
End Sub



' Called at the end of multi ball, this timer resets the side and captive drop targets, and restore WEAPON wall if required.
'
' *** DOCAM OPTIMIZED ***
'
Sub EndMultiballTimer_Timer()
  EndMultiballTimer.Enabled = FALSE
  ResetSideTargets()
  ResetCaptiveTargets()
  If Not(bPlayingMission) Then FixWeapon()
End Sub



' This timer is called at the end of final mission.
'
' *** DOCAM CLEANED ***
'
Sub MissionAllOverTimer_Timer()
  MissionAllOverTimer.Enabled = FALSE
' LookAtPlayfield()
  FlashOff()
  ResetMusic()
  CreateNewBall()
  ingamereset()
  AddScore(0)
End Sub



' This timer displays messages during jackpot award.
'
' *** DOCAM OPTIMIZED ***
'
Dim JackpotDisplayTimerUserData
Sub JackpotDisplayTimer_Timer()
  JackpotDisplayTimerUserData = JackpotDisplayTimerUserData + 1
  Select Case JackpotDisplayTimerUserData
    Case 1:
      D1.Text = "         J                      "
            DisplayB2SText "         J                      "
      bInterruptedDisplay = TRUE
      Playsound "steel"
    ' D1.FlushQueue()
    ' D2.FlushQueue()
    Case 2:
      D1.Text = "         J A                    "
             DisplayB2SText  "         J A                    "
      bInterruptedDisplay = TRUE
      Playsound "steel"
    Case 3:
      D1.Text = "         J A C                  "
            DisplayB2SText  "         J A C                  "
      bInterruptedDisplay = TRUE
      Playsound "steel"
    Case 4:
      D1.Text = "         J A C K                "
            DisplayB2SText "         J A C K                "
      bInterruptedDisplay = TRUE
      Playsound "steel"
    Case 5:
      D1.Text = "         J A C K P              "
            DisplayB2SText "         J A C K P              "
      bInterruptedDisplay = TRUE
      Playsound "steel"
    Case 6:
      D1.Text = "         J A C K P O            "
            DisplayB2SText "         J A C K P O            "
      bInterruptedDisplay = TRUE
      Playsound "steel"
    Case 7:
      D1.Text = "         J A C K P O T          "
            DisplayB2SText "         J A C K P O T          "
      bInterruptedDisplay = TRUE
      Playsound "steel"
    Case 8:
      D1.Text = "  *      J A C K P O T       *  "', seBlinkMask, 2000
            DisplayB2SText "  *      J A C K P O T       *  "
      bInterruptedDisplay = TRUE
      Playsound "explode5"
      Playsound "Cheers"
      FlashEffectMissionTimer.enabled = 1'GI_Flash()
      If Not(bPlayingFinalMission) Then Playsound "MSJackpot"
    Case 12:
      InterruptDisplayTimer.interval = 2000
      JackpotDisplayTimer.Enabled = FALSE
      JackpotDisplayTimerUserData = 0
  End Select
End Sub



' This timer displays messages during attract mode.
'
' *** DOCAM OPTIMIZED ***
'
Dim AttractMessagesTimerUserData
Sub AttractMessagesTimer_Timer()
    FixTargetTimer.enabled = 1
  AttractMessagesTimerUserData = AttractMessagesTimerUserData + 1
  Select Case AttractMessagesTimerUserData
    Case 1:
      D1.Text = "           GAME OVER            "
            DisplayB2SText "           GAME OVER            "
      ' From this point, you can enter in setup mode (by pressing both left & right flipper keys).
      bAllowEnterSetup = TRUE
    Case 2:
      D1.Text = "    BAILEY PINBALL PRESENTS     "
            DisplayB2SText "    BAILEY PINBALL PRESENTS     "
    Case 3:
      'D1.FlushQueue()
      'D2.FlushQueue()
      D1.Text = "           METAL SLUG           "', seScrollIn, 1000, 0, TRUE
            DisplayB2SText "           METAL SLUG           "
    Case 4:
      D1.Text = "       BY BRENDAN BAILEY        "
            DisplayB2SText "       BY BRENDAN BAILEY        "
    Case 5:
      D1.Text = "   BASED ON THE VIDEO GAMES BY  "
            DisplayB2SText "   BASED ON THE VIDEO GAMES BY  "
    Case 6:
      D1.Text = "          NAZCA & SNK           "
            DisplayB2SText "          NAZCA & SNK           "
    Case 7:
      If (Score(1) = 0) And (Credits = 0) Then
        D1.Text = "      PLEASE INSERT CREDIT      "
                DisplayB2SText "      PLEASE INSERT CREDIT      "
      ElseIf (Score(1) = 0) And (Credits > 0) Then
        D1.Text = "    PUSH START BUTTON TO PLAY   "
                DisplayB2SText "    PUSH START BUTTON TO PLAY   "
      Else
                DisplayB2SText " LAST SCORE WAS " & FormatScore(Score(1)) & String(8 - Int((Len(Score(1)) - 1) / 2), " ")
        D1.Text = " LAST SCORE WAS " & FormatScore(Score(1)) & String(8 - Int((Len(Score(1)) - 1) / 2), " ")

      End If
    Case 8:
      D1.Text = "       SPECIAL THANKS TO        "
            DisplayB2SText "       SPECIAL THANKS TO        "
    Case 9:
      D1.Text = "    THE METAL SLUG DATABASE     "
            DisplayB2SText "    THE METAL SLUG DATABASE     "
    Case 10:
      D1.Text = " DOMINIQUE CAMUS  A.K.A. DOCAM   "
            DisplayB2SText " DOMINIQUE CAMUS  A.K.A. DOCAM   "
    Case 11:
      D1.Text = " COLLECT WEAPONS FOR EXTRA BALL "
            DisplayB2SText " COLLECT WEAPONS FOR EXTRA BALL "
    Case 12:
      D1.Text = "     COLLECT POW FOR BONUS      "
            DisplayB2SText "     COLLECT POW FOR BONUS      "
    Case 13:
      D1.Text = " SHOOT KICKERS TO START MISSION "
            DisplayB2SText " SHOOT KICKERS TO START MISSION "
    Case 14:
      D1.Text = "COMPLETE METALSLUG TO LIGHT LOCK"
            DisplayB2SText "COMPLETE METALSLUG TO LIGHT LOCK"
    Case 15:
      D1.Text = "LOCK 4 BALLS TO START MULTI BALL"
            DisplayB2SText "LOCK 4 BALLS TO START MULTI BALL"
    Case 16:
      D1.Text = "** GOOD LUCK ON YOUR MISSIONS **"
            DisplayB2SText "** GOOD LUCK ON YOUR MISSIONS **"
    Case 18:
             If Score(1) Then
                D1.Text = "PLAYER 1"& Score(1)
                DisplayB2SText "PLAYER 1"& Score(1)
             End If
             If Score(2) Then
               D1.Text = "PLAYER 2"& Score(2)
                DisplayB2SText "PLAYER 1"& Score(2)
             End If
             If Score(3) Then
               D1.Text = "PLAYER 3"& Score(3)
                DisplayB2SText "PLAYER 1"& Score(3)
             End If
             If Score(4) Then
               D1.Text = "PLAYER 4"& Score(4)
                DisplayB2SText "PLAYER 1"& Score(4)
             End If

          'coins or freeplay
             If bFreePlay Then
                D1.Text = "FREE PLAY"
                DisplayB2SText "FREE PLAY"
              Else
             If Credits> 0 Then
               DisplayB2SText "    CREDITS " &credits& "     PRESS START"
               D1.Text = "CREDITS " &credits& " PRESS START"
              Else
               DisplayB2SText "    CREDITS " &credits& "     INSERT COIN"
               D1.Text = "CREDITS " &credits& " INSERT COIN"
             End If
             End If
    Case 19:
             DisplayB2SText " HIGHSCORES  "& "1> " & HighScoreName(0) & " :" & HighScore(0)', 0, , , -1
             D1.Text =  " HIGHSCORES"& " 1> " & HighScoreName(0) & " " & FormatNumber(HighScore(0), 0, , , -1)
    Case 20:
             DisplayB2SText " HIGHSCORES  "& "2> " & HighScoreName(1) & " :" & HighScore(1)', 0, , , -1
             D1.Text =  " HIGHSCORES"& " 2> " & HighScoreName(1) & " " & FormatNumber(HighScore(1), 0, , , -1)
    Case 21:
             DisplayB2SText " HIGHSCORES  "& "3> " & HighScoreName(2) & " :" & HighScore(2)', 0, , , -1
             D1.Text =  " HIGHSCORES"& " 3> " & HighScoreName(2) & " " & FormatNumber(HighScore(2), 0, , , -1)
    Case 22:
             DisplayB2SText " HIGHSCORES  "& "4> " & HighScoreName(3) & " :" & HighScore(3)', 0, , , -1
             D1.Text =  " HIGHSCORES"& " 4> " & HighScoreName(3) & " " & FormatNumber(HighScore(3), 0, , , -1)
       AttractMessagesTimerUserData = 0
  End Select
End Sub



' This timer enables the light sequencer when entering attract mode.
'
' *** DOCAM OPTIMIZED ***
'
Sub AttractLightsTimer_Timer()
' AttractLightsTimer.Enabled = FALSE
' MainSeq.Play SeqUpOn, 80, 4
End Sub



' This timer check if entering setup mode.
'
' *** DOCAM OPTIMIZED ***
'
Dim EnteringSetupTimerUserData
Sub EnteringSetupTimer_Timer()
  EnteringSetupTimer.Enabled = FALSE
  If bLeftFlipPressed And bRightFlipPressed And bAllowEnterSetup And Not(bRunningSetup) Then
    ' Stop the messages during attract mode.
    AttractMessagesTimer.Enabled = FALSE
    AttractMessagesTimerUserData = 0
    ' Set flags.
    bAllowEnterSetup = FALSE
    bRunningSetup = TRUE
    bLeftFlipPressed = FALSE
    bRightFlipPressed = FALSE
    D1.Text = "         ENTERING SETUP         "
    EnteringSetupTimerUserData = 1
    EnteringSetupTimer.interval = 1100
    Exit Sub
  End If
  If EnteringSetupTimerUserData = 1 Then
    D1.Text = " LEFT FLIP KEY = NEXT PARAMETER "

    EnteringSetupTimerUserData = 2
    EnteringSetupTimer.interval = 1500
  ElseIf EnteringSetupTimerUserData = 2 Then
    D1.Text = " RIGHT FLIP KEY = CHANGE VALUE  "

    EnteringSetupTimerUserData = 3
    EnteringSetupTimer.interval = 1500
  ElseIf EnteringSetupTimerUserData = 3 Then
    D1.Text = "                                "

    SetupMenuEntry()
    EnteringSetupTimerUserData = 4
    EnteringSetupTimer.interval = 800
  ElseIf EnteringSetupTimerUserData = 4 Then
    ' Reset default for this timer.
    EnteringSetupTimer.Interval = 2500
    EnteringSetupTimerUserData = 5
    ' Display first parameter: balls per game
    SetupMenuDisplayCurrentSetting()
    Playsound "steel"
  End If
End Sub



' *********************************************************************
' **                                                                 **
' **  USER DEFINED SPECIFIC SUBROUTINES & FUNCTIONS                  **
' **                                                                 **
' *********************************************************************



' *********************************************************************
' **  BALLS AND GAME HANDLING                                        **
' *********************************************************************



' Initialize the table for a new game.
'
' *** DOCAM MUST BE REWORKED !!! ***
'

'ResetForNewGame


' Initialize the table for a new ball (either a new ball after the player has lost one).
'
'Sub ResetForNewPlayerBall()
  ' This flag avoids ball saver restart during play.
' bCanActivateBallSaver = TRUE
  ' Assume the player is now playing ball.
' bPlayingBall = TRUE
  ' Make sure the correct display is upto date.
' AddScore(0)
'End Sub



' Create a new ball on the playfield (ejected to the plunger lane).
'
' *** DOCAM OPTIMIZED ***
'
'Sub CreateNewBall()
  ' Create a ball in the plunger lane kicker.
' PlungerKicker.CreateBall()
  ' There is a (or another) ball on the playfield
' BallsOnPlayfield = BallsOnPlayfield + 1
  ' Kick the ball in plunger lane.
' PlungerKickerSolenoidPulse()
'End Sub


' The player has lost his ball (there are no more balls on the playfield).
'
' *** DOCAM OPTIMIZED ***
'
'Sub EndOfBall()
' StopMusic 1
' PlayMusic 7, "Marcodie", FALSE
' PlayMusic 8, "explode3", FALSE
' FlashOff()
  ' Assume the player isn't playing a ball.
' bPlayingBall = FALSE
' D1.QueueText "                                ", seScrollUpOver, 10
' D2.QueueText "                                ", seScrollUpOver, 10
' If bPlayingMission Then EndMission()
' If Not(fpTilted) Then
'   D1.FlushQueue()
'   D2.FlushQueue()
'   InterruptDisplay(1600)
'   D1.Text = "                                "
'   D2.Text = "                                "
'   D1.QueueText "<<<<++  KILLED IN ACTION  ++>>>>", seWipeIn, 800
'   D2.QueueText "<<<<++  KILLED IN ACTION  ++>>>>", seWipeIn, 800
' End If
' EndOfBallTimer2.Enabled = TRUE
'End Sub



' This function is called when the end of bonus display (or high score entry finished) and it either end the game or move onto the next ball.
'
' *** DOCAM OPTIMIZED ***
'
Sub EndOfBallComplete()
    Dim NextPlayer
   ' Is it the end of the game ? (all balls has been lost for player).
  If BallsRemaining(CurrentPlayer) - Ball < 0 Then
    ' Set the machine into game over/attract mode.
      EndOfGame()
  Else
    ' Make sure the correct display is upto date.
    AddScore 0
    ' Reset the playfield for the new ball.
    ResetForNewPlayerBall()
    ' Create a new ball.
    CreateNewBall()
     LightKickback()
      ResetCaptiveTargets()
      AddScore 0
    ResetMusic()
    LightMission()
      POWLaneLight1.State = 0
      POWLaneLight2.State = 0
      POWLaneLight3.State = 0
    ' Unlit left inlane.
      LeftInLaneLight.State = 0
      bLeftInLaneIsLit = FALSE
    ' Unlit right inlane.
      RightInLaneLight.State = 0
      bRightInLaneIsLit = FALSE
    ' Unlit right outlane.
      RightOutLaneLight.State = 0
  End If
End Sub



' This function is called at the end of the game, and start the attract mode.
'
' *** DOCAM OPTIMIZED ***
'
'Sub EndOfGame()
' Dim i, j, iswap
' Dim POWNameTemp
' Dim Rand(100)
' Kicker1.DestroyBall()
' Kicker2.DestroyBall()
' Kicker3.DestroyBall()
' Kicker4.DestroyBall()
'   bBallLockedInKicker1 = FALSE
'   bBallLockedInKicker2 = FALSE
'   bBallLockedInKicker3 = FALSE
'   bBallLockedInKicker4 = FALSE
' bBallInKicker1 = FALSE
' bBallInKicker2 = FALSE
' bBallInKicker3 = FALSE
' bBallInKicker4 = FALSE
'   BallsLocked = 0
'   FlashOff()
'   BallsOnPlayfield = 0
' PlayfieldTimer.Enabled = TRUE
  ' Let Future Pinball know that the game has finished. This also clear the fpGameInPlay flag.
' EndGame()
  ' LookAtBackbox()
'   D1.Text = "           GAME OVER            "
'   D2.Text = "           GAME OVER            "
' Playsound "msmissioncomplete"
  ' Ensure that the flippers are down.
  'LeftFlipper.SolenoidOff()
  'RightFlipper.SolenoidOff()
  ' Entering attract mode (lights and messages).
' EnterAttractMode()
  ' Doing POW name table randomly shuffled.
' For i = 1 To 100
'   Rand(i) = Int(Rnd(1) * 10000)
' Next
' For i = 1 To 99
'   For j = i + 1 To 100
'     If Rand(i) > Rand(j) Then
'       iswap = Rand(i)
'       Rand(i) = Rand(j)
'       Rand(j) = iswap
'       POWNameTemp = POWName(i)
'       POWName(i) = POWName(j)
'       POWName(j) = POWNameTemp
'     End If
'   Next
' Next
'End Sub



' *********************************************************************
' **  SCORING FUNCTIONS                                              **
' *********************************************************************



' Add points to the score.
'
Sub AddScore(points)
  If Not(Tilted) Then
    ' Add the points to the current players score variable.
        Score(CurrentPlayer) = Score(CurrentPlayer) + points
    If Not(bInterruptedDisplay) And (Not(bPlayingMission) Or bPlayingFinalMission) Then
      ' Display score and ball in play.
      If Score(1) < 1000000 Then
                DisplayScore
        D1.Text = FormatScore(Score(1)) & "                        BALL " & Ball
      Else
                DisplayScore
        D1.Text = FormatScore(Score(1)) & String(25 - Len(Score(1)), " ")
      End If
      Exit Sub
    End If
    If Not(bInterruptedDisplay) And bPlayingMission And Not(bMissionSuspended) And Not(bPlayingFinalMission) Then
      If bPlayingMission3 Then
        ' Specific display for mission 3 (Super Jets).
        If Score(1) < 1000000 Then
                    DisplayB2SText (cstr(FormatScore(Score(1)))) & "        REMAINING JETS " & Right("0" & RemainingJets, 2)
          D1.Text = FormatScore(Score(1)) & "        REMAINING JETS " & Right("0" & RemainingJets, 2)
        Else
                    DisplayB2SText (cstr(FormatScore(Score(1)))) & String(14 - Len(Score(1)), " ") & "REMAINING JETS " & Right("0" & RemainingJets, 2)
          D1.Text = FormatScore(Score(1)) & String(14 - Len(Score(1)), " ") & "REMAINING JETS " & Right("0" & RemainingJets, 2)
        End If
      ElseIf bPlayingMission4 Then
        ' Specific display for mission 4 (POW).
        If Score(1) < 1000000 Then
                    DisplayB2SText (cstr(FormatScore(Score(1))))
          D1.Text = FormatScore(Score(1)) & "          REMAINING POW " & RemainingPOW
        Else
                    DisplayB2SText (cstr(FormatScore(Score(1)))) & String(16 - Len(Score(1)), " ") & "REMAINING POW " & RemainingPOW
          D1.Text = FormatScore(Score(1)) & String(16 - Len(Score(1)), " ") & "REMAINING POW " & RemainingPOW
        End If
      ElseIf bPlayingMission5 Then
        ' Specific display for mission 5 (Defeat Allen O'Neill).
        If Score(1) < 1000000 Then
                    DisplayB2SText (cstr(FormatScore(Score(1)))) & "         ALLEN'S ENERGY " & RemainingAllen
          D1.Text = FormatScore(Score(1)) & "         ALLEN'S ENERGY " & RemainingAllen
        Else
                    DisplayB2SText (cstr(FormatScore(Score(1)))) & String(15 - Len(Score(1)), " ") & "ALLEN'S ENERGY " & RemainingAllen
          D1.Text = FormatScore(Score(1)) & String(15 - Len(Score(1)), " ") & "ALLEN'S ENERGY " & RemainingAllen
        End If
      End If
    End If
  End If

End Sub



' This function add a comma as separator for three digits groups (useful for best score legibility).
'
' *** DOCAM OPTIMIZED ***
'
Function FormatScore(ByVal sc)
  Dim fsc
  Dim fdone
  fdone = ""
  While Len(sc) > 3
    fsc = Right(sc, 3)
    If fdone = "" Then
      fdone = "," & fsc
    Else
      fdone = "," & fsc & fdone
    End If
    sc = Left(sc, Len(sc)-3)
  Wend
  fdone = sc & fdone
  FormatScore = fdone
End Function



' Scoring during mission 1, when hitting any target.
'
' *** DOCAM OPTIMIZED ***
'
Sub ScoreMission1()
  Playsound "ok"
  InterruptDisplay(600)
  AddScore(20000)
  'D1.FlushQueue()
  'D2.FlushQueue()
    DisplayB2SText "           * 20,000 *            "
  D1.Text = "           * 20,000 *            "
' D1.QueueText "             20,000              "', seBlinkMask, 300
' D2.QueueText "             20,000              ", seBlinkMask, 300
  Playsound "explode3"
End Sub



' This subroutine add 40,000 points when a WEAPON target is hit during mission 2.
'
' *** DOCAM OPTIMIZED ***
'
Sub ScoreMission2()
  Playsound "Gun4"
  InterruptDisplay(600)
  AddScore(40000)
' D1.FlushQueue()
' D2.FlushQueue()
    DisplayB2SText "           * 40,000 *            "
  D1.Text = "           * 40,000 *            "
' D1.QueueText "             40,000              ", seBlinkMask, 300
' D2.QueueText "             40,000              ", seBlinkMask, 300
End Sub



' This subroutine checks the remaining jets during mission 3.
'
' *** DOCAM OPTIMIZED ***
'
Sub CheckJets()
  AddScore(4500)
  Playsound "Rocket1"
  If RemainingJets = 0 Then CompleteMission()
End Sub



' This subroutine decrements remaining POW during mission 4.
'
' *** DOCAM OPTIMIZED ***
'
Sub ScoreMission4()
  RemainingPOW = RemainingPOW - 1
  InterruptDisplay(600)
' D1.FlushQueue()
' D2.FlushQueue()
    DisplayB2SText "         * POW SAVED *          "
  D1.Text = "         * POW SAVED *          "
' D1.QueueText "           POW SAVED            ", seBlinkMask, 300
' D2.QueueText "           POW SAVED            ", seBlinkMask, 300
  AddScore(15000)
  Playsound "thankyou"
  Playsound "chopper"
  If RemainingPOW = 0 Then CompleteMission()
End Sub



' This subroutine decrements Allen's health during mission 5.
'
' *** DOCAM OPTIMIZED ***
'
Sub ScoreMission5()
  RemainingAllen = RemainingAllen - 1
  AddScore(7500)
  Playsound "gun9", FALSE
  If RemainingAllen = 0 Then CompleteMission()
End Sub



' *********************************************************************
' **  MUSIC & SOUND FUNCTIONS                                        **
' *********************************************************************



' This subroutine restarts the correct theme.
'
' *** DOCAM OPTIMIZED ***
'
Sub ResetMusic()
  If bPlayingMission Then
    Select Case Mission
      Case 1: StopSound Song:Song = "": PlaySong "mu_MSMode1"' Playsound "MSMode1"
      Case 2: StopSound Song:Song = "": PlaySong "mu_MSMode2"'Playsound "MSMode2"
      Case 3: StopSound Song:Song = "": PlaySong "mu_MSMode3"'Playsound "MSMode3"
      Case 4: StopSound Song:Song = "": PlaySong "mu_MSMode2"'Playsound "MSMode2"
      Case 5: StopSound Song:Song = "": PlaySong "mu_MSMode1"'Playsound "MSMode1"
      Case 6: StopSound Song:Song = "": PlaySong "mu_MSmultiball"'Playsound "MSmultiball"
    End Select
  Else
    If bLockIsLit Then
      Playsound "MSLocklit2"
    Else
      If bJackpotScored Then
        Playsound "MSjackpot"
      Else
                StopSound Song:Song = "": PlaySong "mu_MSMain"'
        'Playsound "MSMain"
      End If
    End If
  End If
End Sub



' This subroutine selects next sound and play it on SLUG target hit.
'
' *** DOCAM OPTIMIZED ***
'
Sub PlaySlugSound()
  SlugSound = SlugSound + 1
  Select Case SlugSound
    Case 1: Playsound "yell1"
    Case 2: Playsound "yell2"
    Case 3: Playsound "yell3"
    Case 4: Playsound "yell4"
    Case 5: Playsound "yell5"
    Case 6:
      Playsound "yell6"
      SlugSound = 0
  End Select
End Sub



' Play different sound when the ball roll on outlane/inlane.
'
' *** DOCAM OPTIMIZED ***
'
Sub PlayOutLaneSound()
  OutLaneSound = OutLaneSound + 1
  FlashAnimate1()
  Select Case OutLaneSound
    Case 1: Playsound "mordenlaugh"
    Case 2: Playsound "yell7"
    Case 3:
      Playsound "allenlaugh"
      OutLaneSound = 0
  End Select
End Sub



' *********************************************************************
' **  ATTRACT MODE SPECIFIC FUNCTIONS                                **
' *********************************************************************



' Entering the attract mode.
'
' *** DOCAM OPTIMIZED ***
'
Sub EnterAttractMode()
   StartAttractMode
   'AttractLightsTimer.Enabled = TRUE
   AttractMessagesTimer.Enabled = TRUE
End Sub



' Prepare the lights for the attract mode.
'
' *** DOCAM OPTIMIZED ***
'
Sub SetupLightsAttract()
   ' Characters.
   Marco.state = 2
  Tarma.state = 2
  Eri.state = 2
  Fio.state = 2
   ' Missions.
   Mission1Light.state = 2
  Mission2Light.state = 2
  Mission3Light.state = 2
  Mission4Light.state = 2
  Mission5Light.state = 2
  Mission6Light.state = 2
  ' Weapons and ExtraBall lights.
   Weapon1.state = 2
  Weapon2.state = 2
  Weapon3.state = 2
  Weapon4.state = 2
  Weapon5.state = 2
  Weapon6.state = 2
   ExtraBall.state = 2
  ' WEAPON target lights.
   W1Light.Set 2, "100100100100", 75
  W2Light.Set 2, "010010010010", 75
  W3Light.Set 2, "001001001001", 75
  W4Light.Set 2, "100100100100", 75
  W5Light.Set 2, "010010010010", 75
  W6Light.Set 2, "001001001001", 75
  ' METALSLUG target lights.
   Slug1Light.Set 2, "1000011111", 50
  Slug2Light.Set 2, "1100001111", 50
  Slug3Light.Set 2, "1110000111", 50
  Slug4Light.Set 2, "1111000011", 50
  Slug5Light.Set 2, "1111100001", 50
   Slug6Light.Set 2, "1111100001", 50
  Slug7Light.Set 2, "1111000011", 50
  Slug8Light.Set 2, "1110000111", 50
  Slug9Light.Set 2, "1100001111", 50
  ' Side target lights.
   S1Light.Set      2, "1000011000", 75
  S2Light.Set     2, "1100001100", 75
  S3Light.Set     2, "0110000110", 75
  L4.Set        2, "0011000011", 75
  JackpotLight.Set  2, "1001100001", 75
  ' POW lanes.
  POWLaneLight1.Set 2, "1000001010", 75
  POWLaneLight2.Set 2, "0010001010", 75
  POWLaneLight3.Set 2, "0000101010", 75
  ' Outlanes and inlanes lights.
   RightOutLaneLight.Set  2, "100100100", 75
   RightInLaneLight.Set   2, "010010010", 75
   ShootAgainLight.Set    2, "001001001", 75
  LeftInLaneLight.Set   2, "100100100", 75
   KickbackLight.Set      2, "010010010", 75
  ' Kicker lights.
   L1.Set 2, "10101111111101010000", 75
   M1.Set 2, "10101111111101010000", 75
   L2.Set 2, "00001010111111101010", 75
   M2.Set 2, "00001010111111101010", 75
   L3.Set 2, "01000000101011111101", 75
   M3.Set 2, "01000000101011111101", 75
  ' Captive ball.
   C1Light.Set      2, "10000010100000", 75
  C2Light.Set     2, "00100010100010", 75
  C3Light.Set     2, "00001010101000", 75
  CaptiveLight.Set  2, "00000010100000", 75
  ' Loop.
  LoopLight.Set   2, "0101010100101", 75
End Sub



' Called when the light sequencer queue is empty to repeat the attract mode sequence (lights only).
'
' *** DOCAM OPTIMIZED ***
'
Sub MainSeq_Empty()
' If Not(bGameInPlay) Then
'   SetupLightsAttract()
'   AttractLightsTimer.Enabled = TRUE
' End If
End Sub



' End the attract mode.
'
' *** DOCAM OPTIMIZED ***
'
Sub EndAttractMode()
  AllLightsOff()
    StopAttractMode
 '   LightSeqAttract.StopPlay
 '   LightSeqFlasher.StopPlay
  AttractLightsTimer.Enabled = FALSE
  AttractMessagesTimer.Enabled = FALSE
  AttractMessagesTimerUserData = 0
End Sub














Sub AllLightsOff()
  Marco.State = 0
  Tarma.State = 0
  Eri.State = 0
  Fio.State = 0
   Mission1Light.State = 0
  Mission2Light.State = 0
  Mission3Light.State = 0
  Mission4Light.State = 0
  Mission5Light.State = 0
  Mission6Light.State = 0
   Weapon1.State = 0
  Weapon2.State = 0
  Weapon3.State = 0
  Weapon4.State = 0
  Weapon5.State = 0
  Weapon6.State = 0
   ExtraBall.State = 0
   W1Light.State = 0
  W2Light.State = 0
  W3Light.State = 0
  W4Light.State = 0
  W5Light.State = 0
  W6Light.State = 0
   Slug1Light.State = 0
  Slug2Light.State = 0
  Slug3Light.State = 0
  Slug4Light.State = 0
  Slug5Light.State = 0
   Slug6Light.State = 0
  Slug7Light.State = 0
  Slug8Light.State = 0
  Slug9Light.State = 0
   S1Light.State = 0
  S2Light.State = 0
  S3Light.State = 0
  L4.State = 0
  JackpotLight.State = 0
  POWLaneLight1.State = 0
  POWLaneLight2.State = 0
  POWLaneLight3.State = 0
   RightOutLaneLight.State = 0
   RightInLaneLight.State = 0
   bRightInLaneIsLit = FALSE
   ShootAgainLight.State = 0
  LeftInLaneLight.State = 0
   bLeftInLaneIsLit = FALSE
   KickbackLight.State = 0
   L1.State = 0
   M1.State = 0
   L2.State = 0
   M2.State = 0
   L3.State = 0
   M3.State = 0
   C1Light.State = 0
  C2Light.State = 0
  C3Light.State = 0
  CaptiveLight.State = 0
  LoopLight.State = 0
End Sub









' *** DOCAM OPTIMIZED ***
'
Sub ResetSLUGTargets()
  DOF 123, DOFPulse
  Slug1Light.State =  2', "1000011111", 50
  Slug2Light.State =  2', "1100001111", 50
  Slug3Light.State =  2', "1110000111", 50
  Slug4Light.State =  2', "1111000011", 50
  Slug5Light.State =  2', "1111100001", 50
  Slug6Light.State =  2', "1111100001", 50
  Slug7Light.State =  2', "1111000011", 50
  Slug8Light.State =  2', "1110000111", 50
  Slug9Light.State =  2', "1100001111", 50
  Slug1.Isdropped = 0
  Slug2.Isdropped = 0
  Slug3.Isdropped = 0
  Slug4.Isdropped = 0
  Slug5.Isdropped = 0
  Slug6.Isdropped = 0
  Slug7.Isdropped = 0
  Slug8.Isdropped = 0
  Slug9.Isdropped = 0
End Sub



' This subroutine activate the kickback.
'
' *** DOCAM OPTIMIZED ***
'
Sub LightKickback()
  ' Enable (arm) the kickback.
  KickbackLight.State = 1
  bKickbackIsLit = TRUE
  ' Blocks the left outlane (raise the invisible guide, useful to guide the ball correctly prior to kickback).
  KickbackGuide.isDropped = FALSE
End Sub



' This subroutine lits every red arrow(s) - for free kicker(s) - means ready for playing mission.
'
' *** DOCAM OPTIMIZED ***
'
Sub LightMission()
   If Not(bMissionIsLit) And Not(bPlayingMission) And Not(bMissionSuspended) Then
    bMissionIsLit = TRUE
    If bBallLockedInKicker1 And bBallLockedInKicker2 And bBallLockedInKicker3 Then
      M1.State = 0
      M2.State = 0
      M3.State = 0
    End If
    CaptiveLight.State = 0
    If Not(bMultiBallMode) Then
      M1.State =  2', "10", 100
      M2.State =  2', "10", 100
      M3.State =  2', "10", 100
    End If
    If bBallLockedInKicker1 Then
      M1.State = 1
    Else
      M1.State = 2
    End If
    If bBallLockedInKicker2 Then
      M2.State = 1
    Else
      M2.State = 2
    End If
    If bBallLockedInKicker3 Then
      M3.State = 1
    Else
      M3.State = 2
    End If
   End If
End Sub



' *** DOCAM CLEANED ***
'
Dim MissionTimerUserData
Sub StartMission()
  If Not(bMultiBallMode) Then
    AddScore(5000)
    bMissionIsLit = FALSE
    Mission = Mission + 1
    M1.State = 1
    M2.State = 1
    M3.State = 1
    LoopLight.State = 0
    bPlayingMission = TRUE
    Select Case Mission
      Case 1:
        ' Starting mission 1 (any targets).
        Mission1Light.State = 1
        bPlayingMission1 = TRUE
        StopSound Song:Song = "": PlaySong "mu_MSMode1"'Playsound "MSMode1"
        'Playsound "Mission1"
                DMDFLush
                DMD "Mision1Start.mp4", "", "", 3950
        'StopMusic 4
        KickOutBallTimer.Enabled = TRUE
        CaptiveLight.state = 0
        FlashAnimate5()
        MissionTimerUserData = 35
        MissionTimer.Enabled = TRUE
        NextMessage = 1
        bCanDisplayMessage = TRUE
                DisplayB2SText "         MISSION 1 START        "
        D1.text = "         MISSION 1 START        "', seScrollUp, 500, 0, TRUE
      Case 2:
        ' Starting mission 2 (weapon practice).
        Mission2Light.State = 1
        bPlayingMission2 = TRUE
        StopSound Song:Song = "": PlaySong "mu_MSMode2"'Playsound "MSMode2"
        'Playsound "Mission2"
                DMDFLush
                DMD "Mision2Start.mp4", "", "", 3950
        'StopMusic 4
        FlashAnimate5()
        MissionTimerUserData = 35
        MissionTimer.Enabled = TRUE
        DropWeapon()
        AnimateWeaponTimer.Uservalue = 0
        AnimateWeaponTimer.enabled = 1
        NextMessage = 5
        bCanDisplayMessage = TRUE
                DisplayB2SText "         MISSION 2 START        "
        D1.text = "         MISSION 2 START        "', seScrollUp, 500, 0, TRUE
      Case 3:
        ' Starting mission 3 (super jets).
        Mission3Light.State = 1
        bPlayingMission3 = TRUE
        StopSound Song:Song = "": PlaySong "mu_MSMode3"'Playsound "MSMode3"
        'Playsound "Mission3"
                DMDFLush
                DMD "Mision3Start.mp4", "", "", 3950
        FlashAnimate5()
        Bumper1l.State = 2
        Bumper2l.State = 2
        Bumper3l.State = 2
        NextMessage = 6
        bCanDisplayMessage = TRUE
                DisplayB2SText "         MISSION 3 START        "
        D1.text = "         MISSION 3 START        "', seScrollUp, 500, 0, TRUE
        RemainingJets = 20
      Case 4:
        ' Starting mission 4 (POW).
        Mission4Light.State = 1
        bPlayingMission4 = TRUE
        StopSound Song:Song = "": PlaySong "mu_MSMode2"'Playsound "MSMode2"
        'Playsound "Mission4"
                DMDFLush
                DMD "Mision4Start.mp4", "", "", 3950
        FlashAnimate5()
        CaptiveLight.State = 2
                DisplayB2SText "         MISSION 4 START        "
        D1.text = "         MISSION 4 START        "', seScrollUp, 500, 0, TRUE
        RemainingPOW = 3
        NextMessage = 7
        bCanDisplayMessage = TRUE
        Playsound "chopper"
      Case 5:
        ' Starting mission 5 (defeat Allen O'Neill).
        Mission5Light.State = 1
        bPlayingMission5 = TRUE
        StopSound Song:Song = "": PlaySong "mu_MSMode1"'Playsound "MSMode1"
        'Playsound "Mission5"
                'DMDFLush
                DMD "", "", "Mission 5 Start", 3950
        DropWeapon()
        FlashAnimate5()
                DisplayB2SText "         MISSION 5 START        "
        D1.text = "         MISSION 5 START        "', seScrollUp, 500, 0, TRUE
        RemainingAllen = 5
        WeaponTargetsMove = 0
        DropWeapon()
        AnimateWeaponTimer.Uservalue = 0
        AnimateWeaponTimer.enabled = 1
        NextMessage = 8
        bCanDisplayMessage = TRUE
      Case 6:
        ' Starting mission 6 (final mission).
        Mission6Light.State = 1
        bFinalMissionIsReady = FALSE
        bPlayingFinalMission = TRUE
        bPlayingMission1 = TRUE
      ' LookAtBackbox()
        FinalMissionShowTimer.Interval 1300
                DMDFLush
                DMD "FinalMision.mp4", "", "", 3950
        NextMessage = 10
        bCanDisplayMessage = TRUE
        StopSound Song:Song = "": PlaySong "mu_MSWizard"'Playsound "MSWizard"
        'D1.FlushQueue()
        'D2.FlushQueue()
                DisplayB2SText "         FINAL MISSION          "
        D1.text = "         FINAL MISSION          "', seScrollUp, 500, 0, TRUE
        FlashAnimate5()
        bLockIsLit = FALSE
    End Select
  End If
End Sub



' Terminates any mission (silent).
'
' *** DOCAM OPTIMIZED ***
'
Sub EndMission()
  bPlayingMission = FALSE
   bPlayingMission1 = FALSE
   bPlayingMission2 = FALSE
   bPlayingMission3 = FALSE
   bPlayingMission4 = FALSE
   bPlayingMission5 = FALSE
  bPlayingFinalMission = FALSE
   M1.State = 0
   M2.State = 0
   M3.State = 0
   EndCurrentMission()
   If Not(Tilted) Then FixWeapon()
  CaptiveLight.State = 2
   FlashOff()
   If bWallIsDown Then LoopLight.State = 2
End Sub



' Mission is completed.
'
' *** DOCAM OPTIMIZED ***
'
Sub CompleteMission()
  AddScore(5000)
   bPlayingMission = FALSE
   bPlayingMission1 = FALSE
   bPlayingMission2 = FALSE
   bPlayingMission3 = FALSE
   bPlayingMission4 = FALSE
   bPlayingMission5 = FALSE
   bPlayingFinalMission = FALSE
  FlashEffectMissionTimer.enabled = 1'GI_Flash()
   Playsound "MSMissioncomplete"
   If Mission < 5 Then
    Playsound "MissionComplete"
   ElseIf Mission = 5 Then
    Playsound "seeyouinhell"
   End If
   bInterruptedDisplay = FALSE
   NextMessage = 3
  bCanDisplayMessage = TRUE
   ShowNextMessage()
  ' EffectMusic 1, SetVolume, 0.0
   ResetMusicTimer.Enabled = TRUE
   M1.State = 0
   M2.State = 0
   M3.State = 0
   EndCurrentMission()
   FixWeapon()
   CaptiveLight.State = 2
   FlashOff()
   If bWallIsDown Then LoopLight.State = 2
End Sub



' Terminates a specific mission.
'
' *** DOCAM OPTIMIZED ***
'
Sub EndCurrentMission()
  Select Case Mission
    Case 1:
      CaptiveLight.State = 0
      Mission2Light.State = 2
      MissionTimer.Enabled = FALSE
    Case 2:
      AnimateWeaponTimer.Enabled = FALSE
      Mission3Light.State = 2
      MissionTimer.Enabled = FALSE
      If bWallIsDown Then
        DropWeapon()
      Else
        FixWeapon()
      End If
    Case 3:
      Bumper1l.State = 1
      Bumper2l.State = 1
      Bumper3l.State = 1
      Mission4Light.State = 2
    Case 4:
      Mission5Light.State = 2
      CaptiveLight.State = 0
    Case 5:
      AnimateWeaponTimer.Enabled = FALSE
      Mission6Light.State = 2
      bFinalMissionIsReady = TRUE
      LightMission()
  End Select
End Sub



' *** DOCAM CLEANED ****
'
Sub StartMissionSoon()
  If Not(bBallLockedInKicker1) Or Not(bBallLockedInKicker2) Or Not(bBallLockedInKicker3) Or Not(bBallLockedInKicker4) Then StartMissionSoonTimer.Enabled = TRUE
End Sub



' *** DOCAM OPTIMIZED ****
'
Sub LockBall()
  'LightSeqAttract.Play SeqUpOn, 80
     LightEffect 3
  If bBallInKicker1 Then
    bBallLockedInKicker1 = TRUE
    L1.State = 1
  End If
  If bBallInKicker2 Then
    bBallLockedInKicker2 = TRUE
    L2.State = 1
  End If
  If bBallInKicker3 Then
    bBallLockedInKicker3 = TRUE
    L3.State = 1
  End If
  If bBallInKicker4 Then
    bBallLockedInKicker4 = TRUE
    L4.State = 1
  End If
  BallsLocked = BallsLocked + 1
  If BallsLocked = 1 Then
    DisplayBallLocked()
    AutoBall()
    Playsound "Marco"
        DMDFLush
        DMD "Marco.jpg", "", "", 2000
    Marco.State = 1
    Tarma.State = 0
    Eri.State = 0
    Fio.State = 0
  ElseIf BallsLocked = 2 Then
    DisplayBallLocked()
    AutoBall()
    Playsound "Tarma"
        DMDFLush
        DMD "Tarma.jpg", "", "", 2000
    Marco.State = 1
    Tarma.State = 1
    Eri.State = 0
    Fio.State = 0
  ElseIf BallsLocked = 3 Then
    DisplayBallLocked()
    AutoBall()
    Playsound "Eri"
        DMDFLush
        DMD "Eri.jpg", "", "", 2000
    Marco.State = 1
    Tarma.State = 1
    Eri.State = 1
    Fio.State = 0
  ElseIf BallsLocked = 4 Then
    DisplayBallLocked()
    StartMultiballTimer.Enabled = TRUE
    Playsound "Fio"
        DMDFLush
        DMD "Fio.jpg", "", "", 2000
    Marco.State = 1
    Tarma.State = 1
    Eri.State = 1
    Fio.State = 1
  End If
End Sub



Sub AutoBall()
  CreateNewBall()
  AutoBallTimer.Enabled = TRUE
End Sub




Sub LightJackpot()
  bJackpotIsLit = TRUE
  JackpotLight.State = 2
End Sub



Sub UnLightJackpot()
  bJackpotIsLit = FALSE
  If bJackpotScored Then
    JackpotLight.State = 1
  Else
    JackpotLight.State = 0
  End If
End Sub



Sub EndMultiball()
  If Not(bJackpotScored) Then ResetMusic()
  bMultiBallMode = FALSE
  UnLightJackpot()
  FlashOff()
  ' This timer add delay to reset side and captive drop targets.
  EndMultiballTimer.Enabled = TRUE
  If bMissionSuspended Then
    ' Resume the suspended mission, if required.
    bPlayingMission = TRUE
    bMissionSuspended = FALSE
    ResetMusic()
    FlashAnimate5()
    Select Case Mission
      Case 1:
        MissionTimer.Enabled = TRUE
      Case 2:
        MissionTimer.Enabled = TRUE
        AnimateWeaponTimer.Enabled = TRUE
      Case 3:
        Bumper1l.State = 2
        Bumper2l.State = 2
        Bumper3l.State = 2
        'D1.FlushQueue())
                'DisplayUpdate
        AddScore(0)
      Case 4:
        CaptiveLight.State = 2
        'D1.FlushQueue()
                'DisplayUpdate
        AddScore(0)
      Case 5:
        AnimateWeaponTimer.Enabled = TRUE
        'D1.FlushQueue()
                'DisplayUpdate
        AddScore(0)
    End Select
  End If
  ResetSLUGTargets()
  ResetSideTargets()
  Marco.State = 0
  Tarma.State = 0
  Eri.State = 0
  Fio.State = 0
  L1.State = 0
  L2.State = 0
  L3.State = 0
  L4.State = 0
  If bMissionIsLit Then
    M1.State = 2
    M2.State = 2
    M3.State = 2
  Else
    CaptiveLight.State = 2
  End If
End Sub



Sub ResetCaptiveTargets()
  DOF 121, DOFPulse
  C1.Isdropped = 0'()
  C2.Isdropped = 0'()
  C3.Isdropped = 0'()
  C1Light.State = 0
  C2Light.State = 0
  C3Light.State = 0
  If bMissionIsLit Then
    CaptiveLight.State = 0
  Else
    CaptiveLight.State = 2
  End If
End Sub



Sub ResetSideTargets()
  DOF 120, DOFPulse
  S1.Isdropped = 0'()
  S2.Isdropped = 0'()
  S3.Isdropped = 0'()
  S1Light.State = 0
  S2Light.State = 0
  S3Light.State = 0
End Sub



' *** DOCAM OPTIMIZED ***
'
Sub CheckWeapon()
  If bW1TargetIsDown And bW2TargetIsDown And bW3TargetIsDown And bW4TargetIsDown And bW5TargetIsDown And bW6TargetIsDown Then
    FlashEffectMissionTimer.enabled = 1'GI_Flash()
    CurrentWeapon = CurrentWeapon + 1
    Select Case CurrentWeapon
      Case 1:
        Weapon1.State = 1
        Weapon2.State = 2
        PlaySound "Heavymachinegun"
        InterruptDisplay(1000)
        'D1.FlushQueue()
        'D2.FlushQueue()
                DisplayB2SText "  WEAPON IS HEAVY MACHINE GUN   "
                DisplayB2SText "            HEAVY MACHINE GUN   "
        D1.Text = "  WEAPON IS HEAVY MACHINE GUN   "
        D1.Text =  "            HEAVY MACHINE GUN   "', seBlinkMask, 400
        AddScore(1000)
      Case 2:
        Weapon2.State = 1
        Weapon3.State = 2
        PlaySound "Flameshot"
        InterruptDisplay(1000)
      ' D1.FlushQueue()
      ' D2.FlushQueue()
                DisplayB2SText "      WEAPON IS FLAME SHOT      "
                DisplayB2SText "                FLAME SHOT      "
        D1.Text = "      WEAPON IS FLAME SHOT      "
        D1.Text =  "                FLAME SHOT      "', seBlinkMask, 400
        AddScore(2000)
      Case 3:
        Weapon3.State = 1
        Weapon4.State = 2
        PlaySound "Rocketlauncher"
        InterruptDisplay(1000)
      ' D1.FlushQueue()
      ' D2.FlushQueue()
                DisplayB2SText "    WEAPON IS ROCKER LAUNCHER   "
                DisplayB2SText "              ROCKET LAUNCHER   "
        D1.Text = "    WEAPON IS ROCKER LAUNCHER   "
        D1.Text =  "              ROCKET LAUNCHER   "', seBlinkMask, 400
        AddScore(3000)
      Case 4:
        Weapon4.State = 1
        Weapon5.State = 2
        PlaySound "Shotgun"
        InterruptDisplay(1000)
      ' D1.FlushQueue()
      ' D2.FlushQueue()
                DisplayB2SText "        WEAPON IS SHOTGUN       "
                DisplayB2SText "                  SHOTGUN       "
        D1.Text = "        WEAPON IS SHOTGUN       "
        D1.Text =  "                  SHOTGUN       "', seBlinkMask, 400
        AddScore(4000)
      Case 5:
        Weapon5.State = 1
        Weapon6.State = 2
        PlaySound "EnemyChaser"
        InterruptDisplay(1000)
      ' D1.FlushQueue()
      ' D2.FlushQueue()
                DisplayB2SText "     WEAPON IS ENEMY CHASER     "
                DisplayB2SText "               ENEMY CHASER     "
        D1.Text = "     WEAPON IS ENEMY CHASER     "
        D1.Text =  "               ENEMY CHASER     "', seBlinkMask, 400
        AddScore(5000)
      Case 6:
        Weapon6.State = 1
        ExtraBall.State = 2
        PlaySound "Lasergun"
        InterruptDisplay(1000)
      ' D1.FlushQueue()
      ' D2.FlushQueue()
                DisplayB2SText "       WEAPON IS LASER GUN      "
                DisplayB2SText "                 LASER GUN      "
        D1.Text = "       WEAPON IS LASER GUN      "
        D1.Text =  "                 LASER GUN      "', seBlinkMask, 400
        AddScore(6000)
      Case 7:
        ExtraBall.State = 1
        PlaySound SoundFXDOF("knocker",131,DOFPulse,DOFKnocker)
        DOF 114, DOFPulse
        PlaySound "cheers"
        bExtraBallAward = TRUE
        InterruptDisplay(1000)
      ' D1.FlushQueue()
      ' D2.FlushQueue()
                DisplayB2SText "     * E X T R A  B A L L *     "
        D1.Text = "     * E X T R A  B A L L *     "
      ' D1.QueueText "     * E X T R A  B A L L *     ", seBlinkMask, 600
      Case 8:
        CurrentWeapon = 7
        PlaySound "ok"
        InterruptDisplay(1000)
        AddScore(40000)
      ' D1.FlushQueue()
      ' D2.FlushQueue()
                DisplayB2SText "           * 40,000 *            "
        D1.Text = "           * 40,000 *            "
      ' D1.QueueText "             40,000              ", seBlinkMask, 300
    End Select
    FlashAnimate1()
    bWallIsDown = TRUE
    'LoopLight.state = 2
    LoopLight.State = 2
  End If
End Sub



' *** DOCAM OPTIMIZED ***
'
Sub DropWeapon()
  W1.Isdropped = 1
  W2.Isdropped = 1
  W3.Isdropped = 1
  W4.Isdropped = 1
  W5.Isdropped = 1
  W6.Isdropped = 1
  W1Light.State = 0
  W2Light.State = 0
  W3Light.State = 0
  W4Light.State = 0
  W5Light.State = 0
  W6Light.State = 0
End Sub



' *** DOCAM OPTIMIZED ***
'
Sub FixWeapon()
  LoopLight.State = 0
  If bW1TargetIsDown Then
    W1.Isdropped = 1
    W1Light.State = 1
  Else
    W1.Isdropped = 0
    W1Light.State = 0
  End If
  If bW2TargetIsDown Then
    W2.Isdropped = 1
    W2Light.State = 1
  Else
    W2.Isdropped = 0
    W2Light.State = 0
  End If
  If bW3TargetIsDown Then
    W3.Isdropped = 1
    W3Light.State = 1
  Else
    W3.Isdropped = 0
    W3Light.State = 0
  End If
  If bW4TargetIsDown Then
    W4.Isdropped = 1
    W4Light.State = 1
  Else
    W4.Isdropped = 0
    W4Light.State = 0
  End If
  If bW5TargetIsDown Then
    W5.Isdropped = 1
    W5Light.State = 1
  Else
    W5.Isdropped = 0
    W5Light.State = 0
  End If
  If bW6TargetIsDown Then
    W6.Isdropped = 1
    W6Light.State = 1
  Else
    W6.Isdropped = 0
    W6Light.State = 0
  End If
  If bW1TargetIsDown And bW2TargetIsDown And bW3TargetIsDown And bW4TargetIsDown And bW5TargetIsDown And bW6TargetIsDown Then
    bWallIsDown = TRUE
    LoopLight.State = 2
  End If
End Sub



' *** DOCAM CHECKED ***
'
Sub ResetWEAPONTargets()
  If Not(bMultiBallMode) And Not(bPlayingMission) And bWallIsDown Then
    DOF 135, DOFPulse
    W1.Isdropped = 0
    W1Light.State = 0
    bW1TargetIsDown = FALSE
    W2.Isdropped = 0
    W2Light.State = 0
    bW2TargetIsDown = FALSE
    W3.Isdropped = 0
    W3Light.State = 0
    bW3TargetIsDown = FALSE
    W4.Isdropped = 0
    W4Light.State = 0
    bW4TargetIsDown = FALSE
    W5.Isdropped = 0
    W5Light.State = 0
    bW5TargetIsDown = FALSE
    W6.Isdropped = 0
    W6Light.State = 0
    bW6TargetIsDown = FALSE
    LoopLight.State = 0
    bWallIsDown = FALSE
  End If
End Sub



Sub CheckPOWLanes()
  If (POWLaneLight1.State = 1) And (POWLaneLight2.State = 1) And (POWLaneLight3.State = 1) Then
    AdvancePOW()
    AddScore(7000)
    POWLaneLight1.State = 0
    POWLaneLight2.State = 0
    POWLaneLight3.State = 0
    Playsound "ThankYou"
    If bPlayingMission And (Mission = 4) Then ScoreMission4()
  End If
End Sub



Sub AdvancePOW()
  Dim POWRankRnd
  FlashEffectMissionTimer.enabled = 1'GI_Flash()
  ' Add a POW bonus.
  POWBonus = POWBonus + 1
  POWBonusBall = POWBonusBall + 1
  ' Next POW name (if > 100, return to first).
  POWNamePtr = POWNamePtr + 1
  If POWNamePtr > 100 Then POWNamePtr = 1
  ' Pick a random value (from 0 to 99).
  POWRankRnd = Int(Rnd(1) * 100)
  ' Then convert to rank.
  If POWRankRnd < 25 Then
    POWRankRnd = 1  ' PVT.
  ElseIf POWRankRnd < 40 Then
    POWRankRnd = 2  ' PFC.
  ElseIf POWRankRnd < 50 Then
    POWRankRnd = 3  ' CPL.
  ElseIf POWRankRnd < 60 Then
    POWRankRnd = 4  ' SGT.
  ElseIf POWRankRnd < 70 Then
    POWRankRnd = 5  ' FSG.
  ElseIf POWRankRnd < 75 Then
    POWRankRnd = 6  ' SMA.
  ElseIf POWRankRnd < 80 Then
    POWRankRnd = 7  ' 2LT.
  ElseIf POWRankRnd < 85 Then
    POWRankRnd = 8  ' LT.
  ElseIf POWRankRnd < 90 Then
    POWRankRnd = 9  ' CPT.
  ElseIf POWRankRnd < 94 Then
    POWRankRnd = 10 ' LCL.
  ElseIf POWRankRnd < 98 Then
    POWRankRnd = 11 ' COL.
  Else
    POWRankRnd = 12 ' GEN.
  End If
  ' Add another POW bonus during final mission (because POW is double).
  If bPlayingFinalMission Then
    POWBonus = POWBonus + 1
    POWBonusBall = POWBonusBall + 1
  End If
  ' Maximise POW bonus at... 9999 !!!
  If POWBonus > 9999 Then POWBonus = 9999
  If POWBonusBall > 9999 Then POWBonusBall = 9999
  FlashAnimate2()
  InterruptDisplay(1200)
  ' Display collected POW bonus.
    DisplayB2SText (POWBonus) & " POW SAVED   " & String(Len(POWRank(POWRankRnd)), " ") & " " & String(Len(POWName(POWNamePtr)), " ") & String(Int((18 - Len(POWBonus) - Len(POWRank(POWRankRnd)) - Len(POWName(POWNamePtr))) / 2), " ")
  D1.Text = FormatScore(POWBonus) & " POW SAVED   " & String(Len(POWRank(POWRankRnd)), " ") & " " & String(Len(POWName(POWNamePtr)), " ") & String(Int((18 - Len(POWBonus) - Len(POWRank(POWRankRnd)) - Len(POWName(POWNamePtr))) / 2), " ")
' D1.QueueText FormatScore(POWBonus) & " POW SAVED - " & POWRank(POWRankRnd) & ". " & POWName(POWNamePtr) & String(Int((18 - Len(POWBonus) - Len(POWRank(POWRankRnd)) - Len(POWName(POWNamePtr))) / 2), " ")', seBlinkMask, 300
' D2.QueueText FormatScore(POWBonus) & " POW SAVED - " & POWRank(POWRankRnd) & ". " & POWName(POWNamePtr) & String(Int((18 - Len(POWBonus) - Len(POWRank(POWRankRnd)) - Len(POWName(POWNamePtr))) / 2), " ")', seBlinkMask, 300
End Sub



Sub FlashAnimate1()
    FlashEffect 3
End Sub



Sub FlashAnimate2()
    FlashEffect 3
End Sub



Sub FlashAnimate3()
    FlashEffect 3
End Sub



Sub FlashAnimate4()
    FlashEffect 3
End Sub



Sub FlashAnimate5()
    FlashEffect 3

End Sub



Sub FlashAnimate6()
    FlashEffect 3

End Sub




Sub FlashOff()
  F1.State = 0
  F2.State = 0
  F3.State = 0
  F4.State = 0
  F5.State = 0
  F6.State = 0
  F7.State = 0
  F8.State = 0
  F9.State = 0
  F10.State = 0
  If bPlayingMission Or bMultiBallMode Then FlashAnimate5()
End Sub



' *** DOCAM CLEANED ***
'
Sub D1_Empty()
  If bGameInPlay Then
    If bPlayingMission And bCanDisplayMessage Then
      ShowNextMessage()
    End If
    If bPlayingBall Then AddScore(0)
  End If
End Sub



' *** DOCAM CLEANED ***
'
Sub ShowNextMessage()
  bCanDisplayMessage = FALSE
  If Not(bInterruptedDisplay) Then
    Select Case NextMessage
      Case 1:
        ' Starting mission 1.
        D1.Text = " ANY DROP TARGET SCORES 20,000  "', seScrollUp, 1300, 0, TRUE
        DisplayB2SText " ANY DROP TARGET SCORES 20,000  "', seScrollUp, 1300, 0, TRUE
      Case 2:
        ' Remaining time for mission 1.
        D1.Text = "SHOOT ANY DROP TARGET    TIME " & Right("0" & MissionTimerUserData, 2)
        DisplayB2SText "SHOOT ANY DROP TARGET    TIME " & Right("0" & MissionTimerUserData, 2)
      Case 3:
        ' Mission completed.
        D1.Text =  "        MISSION COMPLETE        "', seScrollUp, 2000, 0, TRUE
        DisplayB2SText  "        MISSION COMPLETE        "', seScrollUp, 2000, 0, TRUE
      Case 4:
        ' Remaining time for mission 2.
        D1.Text = "WEAPON PRACTICE          TIME " & Right("0" & MissionTimerUserData, 2)
        DisplayB2SText "WEAPON PRACTICE          TIME " & Right("0" & MissionTimerUserData, 2)
      Case 5:
        ' Starting mission 2.
        D1.Text =  "SHOOT WEAPON TARGETS FOR 40,000 "', seScrollUp, 2000, 0, TRUE
        DisplayB2SText  "SHOOT WEAPON TARGETS FOR 40,000 "', seScrollUp, 2000, 0, TRUE
      Case 6:
        ' Starting mission 3.
        D1.Text =  "        SUPER JET BUMPERS       "', seScrollUp, 2000, 0, TRUE
        DisplayB2SText  "        SUPER JET BUMPERS       "', seScrollUp, 2000, 0, TRUE
      Case 7:
        ' Starting mission 4.
        D1.Text =  "SHOOT CAPTIVE BALL OR POW LANES "', seScrollUp, 2000, 0, TRUE
        DisplayB2SText  "SHOOT CAPTIVE BALL OR POW LANES "', seScrollUp, 2000, 0, TRUE
      Case 8:
        ' Starting mission 5.
        D1.Text =  "      DEFEAT ALLEN O'NEIL       "', seScrollUp, 2000, 0, TRUE
        DisplayB2SText  "      DEFEAT ALLEN O'NEIL       "', seScrollUp, 2000, 0, TRUE
      Case 9:
        ' End of final mission.
        D1.Text = "      SEE YOU NEXT MISSION      "
        DisplayB2SText "      SEE YOU NEXT MISSION      "
      Case 10:
        ' Starting final mission.
        D1.Text =  "         FINAL MISSION          "', seScrollUp, 1300, 0, TRUE
        DisplayB2SText  "         FINAL MISSION          "', seScrollUp, 1300, 0, TRUE
    End Select
  End If
End Sub



'*** DOCAM OPTIMIZED ***
'
Sub InterruptDisplay(aDelay)
  bInterruptedDisplay = TRUE
  InterruptDisplayTimer.enabled = TRUE
End Sub



' This subroutine display the number of locked ball(s) (from 1 to 4).
'
' *** DOCAM OPTIMIZED ***
'
Sub DisplayBallLocked()
  InterruptDisplay(800)
' D1.FlushQueue()
' D2.FlushQueue()
    DisplayB2SText "          BALL " & BallsLocked & " LOCKED         "
  D1.Text = "          BALL " & BallsLocked & " LOCKED         "
' D1.QueueText "               " & BallsLocked & "                ", seBlinkMask, 400
' D2.QueueText "               " & BallsLocked & "                ", seBlinkMask, 400
End Sub



Sub ExplodeAnimate()
  exp1.frame 1, 26
  exp2.frame 1, 26
  exp3.frame 1, 26
  exp4.frame 1, 26
  exp5.frame 1, 26
  exp6.frame 1, 26
  exp7.frame 1, 26
  exp8.frame 1, 26
  exp9.frame 1, 26
  exp10.frame 1, 26
  exp11.frame 1, 26
  exp12.frame 1, 26
  exp13.frame 1, 26
  exp14.frame 1, 26
  exp15.frame 1, 26
End Sub



Sub ingamereset()
  Mission = 0
  ResetCaptiveTargets()
  ResetSideTargets()
  ResetSLUGTargets()
  bLockIsLit = FALSE
  LightMission()
  bWallIsDown = TRUE
  ResetWEAPONTargets()
  bWallIsDown = FALSE
  UnLightJackpot()
  Mission1Light.State = 2
  Mission2Light.State = 0
  Mission3Light.State = 0
  Mission4Light.State = 0
  Mission5Light.State = 0
  Mission6Light.State = 0
End Sub



' *** DOCAM OPTIMIZED ***
'
Sub GI_Flash()

'GiEffect 3
' Bulb1.FlashForMs 1000, 100, 1
' Bulb2.FlashForMs 1000, 100, 1
' Bulb3.FlashForMs 1000, 100, 1
' Bulb4.FlashForMs 1000, 100, 1
' Bulb5.FlashForMs 1000, 100, 1
' Bulb6.FlashForMs 1000, 100, 1
' Bulb7.FlashForMs 1000, 100, 1
' Bulb8.FlashForMs 1000, 100, 1
' Bulb9.FlashForMs 1000, 100, 1
' Bulb10.FlashForMs 1000, 100, 1
' Bulb11.FlashForMs 1000, 100, 1
' Bulb12.FlashForMs 1000, 100, 1
' Bulb13.FlashForMs 1000, 100, 1
' Bulb14.FlashForMs 1000, 100, 1
' Bulb15.FlashForMs 1000, 100, 1
' Bulb16.FlashForMs 1000, 100, 1
' Bulb17.FlashForMs 1000, 100, 1
' Bulb18.FlashForMs 1000, 100, 1
' Bulb19.FlashForMs 1000, 100, 1
' Bulb20.FlashForMs 1000, 100, 1
' Bulb21.FlashForMs 1000, 100, 1
' Bulb22.FlashForMs 1000, 100, 1
' Bulb23.FlashForMs 1000, 100, 1
' Bulb24.FlashForMs 1000, 100, 1
' Bulb25.FlashForMs 1000, 100, 1
' Bulb26.FlashForMs 1000, 100, 1
' Bulb28.FlashForMs 1000, 100, 1
' Bulb29.FlashForMs 1000, 100, 1
' Bulb30.FlashForMs 1000, 100, 1
' Bulb31.FlashForMs 1000, 100, 1
' Bulb32.FlashForMs 1000, 100, 1
' Bulb33.FlashForMs 1000, 100, 1
' Bulb34.FlashForMs 1000, 100, 1
' Bulb35.FlashForMs 1000, 100, 1
' Bulb36.FlashForMs 1000, 100, 1
' Bulb37.FlashForMs 1000, 100, 1
' Bulb38.FlashForMs 1000, 100, 1
' Bulb39.FlashForMs 1000, 100, 1
' Bulb40.FlashForMs 1000, 100, 1
' Bulb41.FlashForMs 1000, 100, 1
' Bulb42.FlashForMs 1000, 100, 1
' Bulb43.FlashForMs 1000, 100, 1
' Bulb44.FlashForMs 1000, 100, 1
  If Not(bPlayingMission3) Then
    FlashForMs Bumper1l, 1000, 100, 1
    FlashForMs Bumper2l, 1000, 100, 1
    FlashForMs Bumper3l, 1000, 100, 1
  End If
End Sub



' *** DOCAM OPTIMIZED ***
'
Sub GI_Off()
GiOff

' Bulb1.State = 0
' Bulb2.State = 0
' Bulb3.State = 0
' Bulb4.State = 0
' Bulb5.State = 0
' Bulb6.State = 0
' Bulb7.State = 0
' Bulb8.State = 0
' Bulb9.State = 0
' Bulb10.State = 0
' Bulb11.State = 0
' Bulb12.State = 0
' Bulb13.State = 0
' Bulb14.State = 0
' Bulb15.State = 0
' Bulb16.State = 0
' Bulb17.State = 0
' Bulb18.State = 0
' Bulb19.State = 0
' Bulb20.State = 0
' Bulb21.State = 0
' Bulb22.State = 0
' Bulb23.State = 0
' Bulb24.State = 0
' Bulb25.State = 0
' Bulb26.State = 0
' Bulb28.State = 0
' Bulb29.State = 0
' Bulb30.State = 0
' Bulb31.State = 0
' Bulb32.State = 0
' Bulb33.State = 0
' Bulb34.State = 0
' Bulb35.State = 0
' Bulb36.State = 0
' Bulb37.State = 0
' Bulb38.State = 0
' Bulb39.State = 0
' Bulb40.State = 0
' Bulb41.State = 0
' Bulb42.State = 0
' Bulb43.State = 0
' Bulb44.State = 0
  Bumper1l.State = 0
  Bumper2l.State = 0
  Bumper3l.State = 0
End Sub



' *** DOCAM OPTIMIZED ***
'
Sub GI_On()
GiOn
' Bulb1.State = 1
' Bulb2.State = 1
' Bulb3.State = 1
' Bulb4.State = 1
' Bulb5.State = 1
' Bulb6.State = 1
' Bulb7.State = 1
' Bulb8.State = 1
' Bulb9.State = 1
' Bulb10.State = 1
' Bulb11.State = 1
' Bulb12.State = 1
' Bulb13.State = 1
' Bulb14.State = 1
' Bulb15.State = 1
' Bulb16.State = 1
' Bulb17.State = 1
' Bulb18.State = 1
' Bulb19.State = 1
' Bulb20.State = 1
' Bulb21.State = 1
' Bulb22.State = 1
' Bulb23.State = 1
' Bulb24.State = 1
' Bulb25.State = 1
' Bulb26.State = 1
' Bulb28.State = 1
' Bulb29.State = 1
' Bulb30.State = 1
' Bulb31.State = 1
' Bulb32.State = 1
' Bulb33.State = 1
' Bulb34.State = 1
' Bulb35.State = 1
' Bulb36.State = 1
' Bulb37.State = 1
' Bulb38.State = 1
' Bulb39.State = 1
' Bulb40.State = 1
' Bulb41.State = 1
' Bulb42.State = 1
' Bulb43.State = 1
' Bulb44.State = 1
  Bumper1l.State = 1
  Bumper2l.State = 1
  Bumper3l.State = 1
End Sub



' *********************************************************************
' **  SETUP MENU SPECIFIC FUNCTIONS                                  **
' *********************************************************************



' This subroutine initializes variables and parameters on setup menu entry.
'
' *** DOCAM OPTIMIZED ***
'
Sub SetupMenuEntry()
  ' Always select first setting (number of balls) on setup menu entry.
  stCurrentParam = 1
  ' Parameter 1: Number of balls per game.
  stParam(1) = "BALLS PER GAME"
  stNumberValues(1) = 4
  stValue(1, 1) = "3"
  stValue(1, 2) = "5"
  stValue(1, 3) = "7"
  stValue(1, 4) = "9"
  Select Case BallsPerGame
    Case 3: stCurrentValue(stCurrentParam) = 1
    Case 5: stCurrentValue(stCurrentParam) = 2
    Case 7: stCurrentValue(stCurrentParam) = 3
    Case 9: stCurrentValue(stCurrentParam) = 4
  End Select
  ' Parameter 2: ball saver duration.
  stParam(2) = "BALL SAVER"
  stNumberValues(2) = 4
  stValue(2, 1) = "15 SECONDS"
  stValue(2, 2) = "30 SECONDS"
  stValue(2, 3) = "45 SECONDS"
  stValue(2, 4) = "60 SECONDS"
  Select Case nvR1
    Case 15: stCurrentValue(2) = 1
    Case 30: stCurrentValue(2) = 2
    Case 45: stCurrentValue(2) = 3
    Case 60: stCurrentValue(2) = 4
  End Select
  ' Parameter 3: POW bonus value.
  stParam(3) = "POW BONUS VALUE"
  stNumberValues(3) = 6
  stValue(3, 1) = "5,000"
  stValue(3, 2) = "10,000"
  stValue(3, 3) = "15,000"
  stValue(3, 4) = "20,000"
  stValue(3, 5) = "25,000"
  stValue(3, 6) = "50,000"
  Select Case nvR2
    Case 5: stCurrentValue(3) = 1
    Case 10: stCurrentValue(3) = 2
    Case 15: stCurrentValue(3) = 3
    Case 20: stCurrentValue(3) = 4
    Case 25: stCurrentValue(3) = 5
    Case 50: stCurrentValue(3) = 6
  End Select
  ' Parameter 4: POW bonus hold.
  stParam(4) = "POW BONUS HOLD"
  stNumberValues(4) = 3
  stValue(4, 1) = "NEVER"
  stValue(4, 2) = "X-BALL ONLY"
  stValue(4, 3) = "ALWAYS"
  stCurrentValue(4) = nvR3
  ' Parameter 5: jackpot value.
  stParam(5) = "JACKPOT VALUE"
  stNumberValues(5) = 4
  stValue(5, 1) = "100,000"
  stValue(5, 2) = "300,000"
  stValue(5, 3) = "500,000"
  stValue(5, 4) = "1,000,000"
  Select Case nvR4
    Case 1: stCurrentValue(5) = 1
    Case 3: stCurrentValue(5) = 2
    Case 5: stCurrentValue(5) = 3
    Case 10: stCurrentValue(5) = 4
  End Select
  ' Parameter 6: Tilt.
  stParam(6) = "TILT SENSITIVITY"
  stNumberValues(6) = 4
  stValue(6, 1) = "NEVER TILT"
  stValue(6, 2) = "LOW"
  stValue(6, 3) = "REGULAR"
  stValue(6, 4) = "HIGH"
  stCurrentValue(6) = nvR5
  ' Parameter 7: Reset settings/high scores.
  stParam(7) = "RESET"
  stNumberValues(7) = 4
  stValue(7, 1) = "NO"
  stValue(7, 2) = "TABLE SETTINGS ONLY"
  stValue(7, 3) = "HIGH SCORES ONLY"
  stValue(7, 4) = "ALL (FACTORY)"
  stCurrentValue(7) = 1
  ' Parameter 8: Exit/Save.
  stParam(8) = "EXIT"
  stNumberValues(8) = 3
  stValue(8, 1) = "NO"
  stValue(8, 2) = "YES & SAVE"
  stValue(8, 3) = "YES, BUT DON'T SAVE"
  stCurrentValue(8) = 1
End Sub



' This subroutine saves settings when exiting the setup menu.
'
' *** DOCAM OPTIMIZED ***
'
Sub SetupMenuSaveExit()
  ' Saving parameter 1: balls per game.
  BallsPerGame = stCurrentValue(1) * 2 + 1
  ' Saving parameter 2: ball saver duration.
  Select Case stCurrentValue(2)
    Case 1: nvR1 = 15
    Case 2: nvR1 = 30
    Case 3: nvR1 = 45
    Case 4: nvR1 = 60
  End Select
  ' Saving parameter 3: POW bonus value.
  Select Case stCurrentValue(3)
    Case 1: nvR2 = 5
    Case 2: nvR2 = 10
    Case 3: nvR2 = 15
    Case 4: nvR2 = 20
    Case 5: nvR2 = 25
    Case 6: nvR2 = 50
  End Select
  ' Saving parameter 4: POW bonus hold.
  nvR3 = stCurrentValue(4)
  ' Saving parameter 5: jackpot value.
  Select Case stCurrentValue(5)
    Case 1: nvR4 = 1
    Case 2: nvR4 = 3
    Case 3: nvR4 = 5
    Case 4: nvR4 = 10
  End Select
  ' Saving parameter 6: Tilt sensitivity.
  nvR5 = stCurrentValue(6)
  Select Case stCurrentValue(6)
    Case 1: TiltWarnings = 9
    Case 2: TiltWarnings = 4
    Case 3: TiltWarnings = 2
    Case 4: TiltWarnings = 0
  End Select
End Sub



' This subroutine reset to default (factory) settings.
' Note: High score table isn't altered, however.
'
' *** DOCAM OPTIMIZED ***
'
Sub SetupMenuResetSettings()
  ' Reset default settings and clear the highscore table.
  stCurrentValue(1) = 2
  BallsPerGame = 5
  stCurrentValue(2) = 1
  nvR1 = 15
  stCurrentValue(3) = 2
  nvR2 = 10
  stCurrentValue(4) = 3
  nvR3 = 3
  stCurrentValue(5) = 2
  nvR4 = 3
  stCurrentValue(6) = 3
  nvR5 = 3
  TiltWarnings = 2
  stCurrentValue(7) = 1
End Sub



' This subroutine reset (clean) the high score table.
'
' *** DOCAM OPTIMIZED ***
'
Sub SetupMenuResetHighScores()
  ' Reset the high scores table.
  HighScore(1) = 0
  HighScoreName(1) = "..."
  HighScore(2) = 0
  HighScoreName(2) = "..."
  HighScore(3) = 0
  HighScoreName(3) = "..."
  HighScore(4) = 0
  HighScoreName(4) = "..."
  HighScore(5) = 0
  HighScoreName(5) = "..."
  HighScore(6) = 0
  HighScoreName(6) = "..."
  HighScore(7) = 0
  HighScoreName(7) = "..."
  HighScore(8) = 0
  HighScoreName(8) = "..."
  HighScore(9) = 0
  HighScoreName(9) = "..."
  HighScore(10) = 0
  HighScoreName(10) = "..."
End Sub



' This subroutine select the next item (in list) for a specific settings.
'
' *** DOCAM OPTIMIZED ***
'
Sub SetupMenuNextValue()
  Dim i
  Dim LenParam, LenValue
  stCurrentValue(stCurrentParam) = stCurrentValue(stCurrentParam) + 1
  If stCurrentValue(stCurrentParam) > stNumberValues(stCurrentParam) Then stCurrentValue(stCurrentParam) = 1
  LenParam = 0
  For i = 1 To Len(stParam(stCurrentParam))
    If (Mid(stParam(stCurrentParam), i, 1) <> ".") And (Mid(stParam(stCurrentParam), i, 1) <> ",") Then LenParam = LenParam + 1
  Next
  LenValue = 0
  For i = 1 To Len(stValue(stCurrentParam, stCurrentValue(stCurrentParam)))
    If (Mid(stValue(stCurrentParam, stCurrentValue(stCurrentParam)), i, 1) <> ".") And (Mid(stValue(stCurrentParam, stCurrentValue(stCurrentParam)), i, 1) <> ",") Then LenValue = LenValue + 1
  Next
  'D1.FlushQueue()
  D1.Text = " " & stParam(stCurrentParam) & String(30 - LenParam - LenValue, " ") & stValue(stCurrentParam, stCurrentValue(stCurrentParam)) & " "
  D1.QueueText " " & String(30 - LenValue, " ") & stValue(stCurrentParam, stCurrentValue(stCurrentParam)) & " ", seBlinkMask, 500
  Playsound "Lane1", FALSE
End Sub



' This subroutine display the current setting.
'
' *** DOCAM OPTIMIZED ***
'
Sub SetupMenuDisplayCurrentSetting()
  Dim i
  Dim LenParam, LenValue
  LenParam = 0
  For i = 1 To Len(stParam(stCurrentParam))
    If (Mid(stParam(stCurrentParam), i, 1) <> ".") And (Mid(stParam(stCurrentParam), i, 1) <> ",") Then LenParam = LenParam + 1
  Next
  LenValue = 0
  For i = 1 To Len(stValue(stCurrentParam, stCurrentValue(stCurrentParam)))
    If (Mid(stValue(stCurrentParam, stCurrentValue(stCurrentParam)), i, 1) <> ".") And (Mid(stValue(stCurrentParam, stCurrentValue(stCurrentParam)), i, 1) <> ",") Then LenValue = LenValue + 1
  Next
  D1.Text = " " & stParam(stCurrentParam) & String(30 - LenParam - LenValue, " ") & stValue(stCurrentParam, stCurrentValue(stCurrentParam)) & " "
End Sub


Dim FlashEffectMissionTimerUserData
Sub FlashEffectMissionTimer_Timer

    FlashEffectMissionTimerUserData = FlashEffectMissionTimerUserData + 1
    FlashEffectMissionTimer.interval = 800
    FlashEffectMissionTimer.enabled = 1
  Select Case FlashEffectMissionTimerUserData

         Case 1:  FlashEffect 3

         Case 2:  FlashEffect 2

         Case 3:  FlashEffect 3

         Case 4:  FlashEffect 4

         Case 5:  FlashEffect 2

         Case 6:  FlashEffect 3

         Case 7:  FlashEffect 4
                  FlashEffectMissionTimer.enabled = 0

  End Select


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


Dim Char(32),i,TempText                     'increase dimension if You need larger displays



'-----------------------------------------------
'-----  B2S section, not used in the demo  -----
'-----------------------------------------------

Sub DisplayB2SText(TextPar)             'Procedure to display Text on a 30 digit B2S LED reel. Assuming that it is display 1 with internal digit numbers 1-30
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








Sub DisplayScore
  If Score(1) < 1000000 Then
   DisplayB2SText (cstr(Score(1))) & "" & "                   BALL " & Ball
  Else
   DisplayB2SText (cstr(Score(1))) & (Score(1)) & String(32 - Len(Score(1)), " ")
  End If
End Sub


Sub DisplayUpdate
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


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


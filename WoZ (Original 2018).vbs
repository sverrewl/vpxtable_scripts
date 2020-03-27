' ****************************
'     WoZ -  Original 2018
' ****************************
'
' All the material making up this emulation (pictures, rules of operation, graphics, sounds, etc.)
' were obtained from information available in the public domain.  The entire endeavour was undertaken
' for private use, educational purposes, and as a "technology" parody attempting to accentuate the
' differences between what is possible using a virtual emulation when compared to a physical pinball machine.
' As this material finds its way into the public domain, its use for any commercial purposes is NOT
' in anyway sanctioned, supported, or encouraged. Just to be even more explicit; this work is NOT
' to be used for ANY commercial purpose. It is not intended to do "harm" to any legally produced and
' commercially available product or going concern involved in the production of the product.
'
' Acknowledgements and thanks to:
' The people who develop and are developing VPX, the B2S backglass designer, VPinMAME and the
' many tables, backglasses and other support works (e.g. DOF, PinballX, PinballY, PinUPsystem, the Forums)
' that are the source of much learning and enjoyment . THANKS!  More specifically for this work ...
' JPSalas - most of this table design as well as a large portion of the code and playfield items
'           came from his tables especially Slimer & Diablo. He is a prolific, master Table Builder;
'           Thanks to him and the others who have and continue to help him.
' Wildman's many backglass examples and shared implementation instructions in the forums inspired this attempt.
' mfuegemann & STAT - the backglass innovations (e.g. diablo, moon light) demonstrate by example
'                     how much of what has been implemented below is done.
' The Walking Dead's wellzombie primitive arms were used in the bumper tree primitives.
' Used the Environment Emission Image from Hot Tip table - 12/2017
' Sliderpoint provided high resolution playfields with blueprints
' Nailbuster provided guidance and support with setting up the PUPSystem and its use.
' EmBee has provided higher resolution images for some of the playfield elements for the 2018 releases.
' Tarcisio provided many improved backglass images and PUPpack videos and provided an approach to using some of those
' videos while still having visible backglass timers.
' The 3D Witch model was done by SnakeEyes Friex with enhancements made by Dark
' Bigus1 worked on latest (20200220 Release) lighting changes
' EmBee updated the playfields, plastics, backwall and target images, plunger decal, apron decal, provided
' updated and additional PUPpack videos and testing support for the 2020 release.
'
'
' Release Notes
' 20200314.01 - Replaced Playfields (main, castle & munchkin), replaced Witch with 3D Model, back wall images, plastics,
'               shooter lane decal, munchkin house roof, soldier and rainbow target images, spinner image,
'               apron and plunger decal images, added OZ lane sign, Hologram rework including images, lighting changes,
'               Pup pack video changes and additions, DOF gear and DOF MX Effect support, lightsequence changes,
'               sound additions, playfield walls added and plenty of fault fixes.  The backglass has
'               also been changed (missing digit fault) and is included in the release.
'             - Added a castle wall in front of doors and adjusted area to correct ball stuck issues and
'               added Glinda image
' 20181230.02 - Added SOTR mode; lighting changes; more PUPpack triggers; table tweaks and fault fixes
' 20181025.01 - Added support for PUPpack movie clips; new plunger plate, State Fair decal, & Castle Doors;
'               fixes for Ball stuck, ramp flyover, & Munchkin mode music restart; added a Test mode constant
'               changed the table rom name to woz_original to match DOF ConfigTool entry.
' 20181004.02 - Updated to Hi-Res Playfield (larger size); Tilt fix; Audio Adjustments; Castle Door Pictures
' 20180804.02 - DOF and sound file adjustments
' 20180122.01 - VPX 10.4.0 Final (built using a 10.3 new table and the Slimer table script as a base)
'               First external release - playable; however caveats lsted below
'
' Install Notes
' 20200314 - Included in this release package are: vpx table, backglass, ultraDMD directory, and a
'            PUPpack. The DOF Configtool als has been modified and the new version has been submitted for
'            public consideration. The backglass file is large and will not load unless the 4G patch
'            is applied to the backglass server.  The backglass is a very significant part of the gameplay
'            experience and should not be removed when the PUPpack is installed.
'            There is an option in the table script below (under GAME OPTIONS) to turn off the PUPpack triggers
'            if you don't install the PUPpack. Images and sound files will replace the PUPpack videos but the
'            PUPpack videos significantly enhance the gameplay experience and play well with the backglass.
'            To reiterate ... Whether the PUPpack is installed or not, the backglass should still be included.
' ********************************************************************************************************************
'
' **********************************************************************************
'     RULE CHANGES FROM
'  WOZ Rulesheet Version 1.26.pdf
'  Available at the following link (link still active but content unchanged 20200314)
' ***********************************************************************************
'
' https://papa.org/wp-content/uploads/WOZ-Rulesheet-Version-1-26.pdf
'
'
' NOT IMPLEMENTED
' - Match not implemented (no code and also will need a full screen animation)
' - No Witch Legs popping out of house (need primitive and supporting code)
' - No Competition/Tournament Mode (not even contemplated)
' - Multiple Player Support (needs more testing)
' - Desktop mode has not been tested.  All development and testing done in cabinet mode. The gameplay
'   experience (and code) is deeply intwined with the backglass. Consequently the backglass should remain
'   even when a PUPpack is active.
'
' SKILL SHOTS
' - Skill shot not activated when ball locks at Munchkin House and new ball created.
' - Skill shot remains active if active before Monkey capture
' - Haunted Forest skill shot increases by 2,500 after successful hit
' - Crystal Ball letter skill shot increases by 5,000 after successful hit

' HURRY-UP
' - The value last collected becomes the "floor" for the next award.  The "top" is the "floor"
'   plus a constant point value (1500) that keeps the duration of the HurryUp the same.
' - The Hurry Up pot is reset to the base value (1500) whenever a hurry up is not made.
' - The Hurry Up pot is the same (there's only one pot) for all types of "Inlane" Hurry Ups.
'   Separate pots are not set up and maintained for each type of Hurry Up.
' - Only one inlane HurryUp is active at any one time.
'
' GLINDA AWARDS
' - During Crystal Ball Lights Out or On, the lights may be returned to normal (2x scoring remains)
' - During Crystal Ball No-Hold or Linked Flippers or any multiball mode, an extra ball may be added
' - During timed Haunted Modes & Munchkin Modes 25 to 35 seconds up to a maximum of 99 seconds may be added.
' - There is no ball resurrection after any multiball; however adding a balls is possible.
' - The Glinda Awards are random; however, the same back-to-back awards are still possible.
'
' WIZARD AWARDS
' - The letters for WIZARD are displayed above the image of the wizard
' - If the "Special" Light is lit as a result of the WIZARD award, then an extra ball will be awarded
'   during free play.
'
' HORSE OF A DIFFERENT COLOR (HOADC)
' - A "white" horse is not displayed as a cow.
' - A "rainbow" horse has the light inset cycle through all colors except white
'
' THERE'S NO PLACE LIKE HOME (TNPLH) BALL SAVE
' - When all stages are completed, the playfield reverts to the state prior to the drain event
'   with the TNPLH lights cleared.
' - During stage 1 (RAINBOW Lights), a hit target spots the light of an adjacent unlit target if an
'   unlit one is available. This reduces the difficulty.
'
' TOTO BALL SAVE
' - Toto escapes only count ("index") if you are successful.  Only three (3) successful escapes are
'   allowed for each player during a single game.
'
' EMERALD CITY MULTIBALL
' - Each player's locked balls are displayed and restored at the start of their turn.
'
' RESCUE MULTIBALL
' - The Capture Dorothy Light is yellow when Dorothy can be captured and it is red once she is captured.
' - The Winkie target only ever requires one hit to knock down.
' - The castle door only requires two hits to open  - can be changed below in Game Option constants
' - RESCUE does not need to be spelled again (relit) if a locked ball is added; however only two
'   additional locked balls can be added.
' - Hitting a blinking letter before entering the kicker behind the door is not needed to collect the Super Jackpot.
' - When spelling RESCUE, hitting the saucer behind the door will lite up to two unlit letters.
' - The concept of "advances" for multiple activations of RESCUE multiball was not implemented.
' - Lighting a letter requires only one hit.
'
' TWISTER MODE
' - Each rainbow target requires one hit to light the associated light.
' - TWISTER mode has a grace period; however, the timed Munchkin modes do not.
'
' WICKED WITCH ATTACK (HURRY UP)
' - Hit the witch target three times to start a Wicked Witch Hurry Up - can be changed in Game Options
' - Hit the witch target twice to stop a Wicked Witch Hurry Up - can be changed in Game Options
'
' CRYSTAL BALL MODES
' - Weak Flippers is replaced with LINKED Flippers (all flippers active when either flipper button pressed.
'   (Was not able to figure out how to implement a "weak" flipper during the game)
' - There is no five second ball saver available at the end of Flipper Frenzy.
'
' SOMEWHERE OVER THE RAINBOW (SOTR)
' - Reaching this mode requires accumulating eight (8) jewels; however, they can include multiples
'   of the same jewel. (This mode is difficult enough to reach:).
' - Details on scoring are sparse in the PAPA document so there may need to be changes made here.
' - This mode remains active even when there is only one ball on the playfield (unlike other multiball
'   modes) since it is so hard to reach.
' - When SOTR is ready but not yet activated, if TOTO and/or the TNPLH lights are all lit, the mode will not be
'   entered on ball drain. However, the state of the TOTO & TNPLH lights will be held until after SOTR.
'
' MAJOR APPEARANCE DIFFERENCES
' - Some of the scenes played on the JJP backglass were not implemented; others were either reproduced
'   using animations or a static screen.  The game progress diagram is not implemented.
' - The playfield is a standard size emulation not a widebody.
' - Lights were added to the OZ lane insets
' - In general, the lighting is less active (chaotic) due to light changing speed limitations (as the
'   twister mode circling demonstrates on slower machines).
'   UPDATE: On a Windows 10 machine, i7, 32GB, GTX1080 the speed of the lighting increases dramatically even on a
'   4k playfield.
' - No witch legs pop out of the house
' - The Munchkin house has Leds on the roof to indicate when a ball is locked because the last two locked balls are
'   not visible when locked.
' - Static lights were added to tree bumpers for eyes. (tree arms come from TWD; thanks)
' - Walls added to castle and although you can't see them clear walls extend up to keep ball from
'   being hit out of castle playfield or over lanes within the castle playfield.
' - Fire ring added to witch with no plastic enclosure
' - Crystal Ball simulation is autonomous and images are random (don't relate to gameplay)
' - Some voice prompts and game related sounds have been changed or omitted while others have been added.
' - There are videos (when the PUPpack option is active) along with still images. There is full game coverage
'   with backglass images and some animations even when there is no PUPpack.
'
' PUPpack (Movie Clip) Additions
' - If the PUPpack constant below under Game Options is set to True then movie clips will be played.
'   A PUPpack is part of the release. Some of the clips will not be played during multiball modes and
'   some others will also cause the ball to stop while the clip plays.
'   This was done both to help with synchronization issues and to provide game play information.
'   Nearly all the PuPpack videos have been supplied or upgraded by EmBee.
'   There is also an awesome Attact mode video provided by Tarcisio.
' PUPpack Notes:
'   The screen designations (that are not off) are: Topper-0 and Backglass-2.
'   Screens  5, 6, 7, 8, 9 & 10 all use custom positioning that displays on the backglass
'   Screens  7, 8, 9 & 10 are the quadrants while 5 is full screen and 6 is almost full screen
'   Screen 11 uses custom positioning that displays the topper
'   No videos are ever displayed on screen 2
'   Screens 1, 3 & 4 (DMD, Playfield & Music) are all in the off mode
'   This information is available in the PUPpack editor
'
' *****************************************************************************************************************
' KNOWN FAULTS OR ISSUES
' 20200307
' - The amount of testing done has verified most of the implemented rules and modes. A regression test involving
'   automated play has been run many times and for multiplayer games with up to 10 balls. The testing
'   code was left in the table script but deactivated with flags.  There are screen display
'   synchronization issues when multiple multiball modes are active.  This has been handled by replacing
'   some full screen PUPpack videos with short quarter screen versions of the videos during multiball modes.
'   However, even these images may not be displayed if a Witch Hurry Up is in progress.
' - Multiplayer operation has only had basic testing done. So the lack of faults found here is a direct
'   result of a lack of testing.
' - Backglass screens and displayed digits sometimes (usually when many modes are active) step on
'   other screens (remain on when they shouldn't).  Code has been added when the ball drains to again clear
'   screens that should've cleared. This seems to help the game correct itself. The layers and initial
'   loading sequence of the illumination snippets (in the backglass) was carefully planned and has seen significant
'   revisions. Scoring doesn't seem affected by these issues although it may lag switch events.
'   A backglass spreadsheet is available on request that lists all the screens, the layering and ids.
' - The 4G patch MUST be applied to B2SBackglassServer.exe for the backglass to load.  The backglass file is well beyond
'   the normal backglass size.
' ******************************************************************************************************************

Option Explicit
Randomize

Const BallSize = 50 ' 50 is the normal size
Const cGameName = "woz_original"
Const cTableName = "Wizard of Oz, The"
Const tempcGameName = "WoZ"

' ***********************************************************************************************
'                                        GAME OPTIONS
' ***********************************************************************************************
Const PUPpack = True            ' Plays videos using PinUP system replacing backglass animations
                                ' and some images
Const BallsPerGame = 5          ' 3 or 5 or whatever you need up to 99
Const maxPlayers = 4            ' More than 1 player has not been tested very well
Const MaxMultiballs = 4         ' Max number of balls on playfield during multiball mode(s)
Const ExtraBallScore = 100000
Const ReplayScore = 250000
Const MegaJackpot = 80000
Const SuperJackpot = 25000
Const CastleDoorHits = 2
Const WitchHurryUpStartHits = 3
Const WitchHurryUpStopHits = 2

' UltraDMD Option
Const UltraDMD_Option = True    ' True turns on UltraDMD

' DMD on Apron Option
Const DMD_Option = True         ' True places a DMD on the Apron
' ***********************************************************************************************


' Forces the creation and running of the Backglass server if B2SON is False
Const B2SOverride = False

' Load the core.vbs for supporting Subs and functions
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
    On Error Goto 0
    If B2SOverride=True AND B2SOn=False Then
        Set Controller = CreateObject("B2S.Server")
    End If
End Sub

' Thalamus 2020 March : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 800     ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


' Define Constants
Const TableName = "WoZ"
Const cBallImages = 13

' Define Global Variables
Dim CurrentPlayer
Dim Credits
Dim PlayersPlayingGame
Dim TotalGamesPlayed

Dim Tilt
Dim Tilted
Dim TiltSensitivity

'  Storage for multiple player game in progress states
Dim BallsRemaining(4)
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim ExtraBallsAwards(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Score(4)

' Define General Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock
Dim BallsInHole
Dim FreezeCnt

' Define General Game Flags
Dim bFreeze
Dim bDisableTable
Dim bAttractMode
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim bMusicOn
Dim bSkillshotReady
Dim bSkillshotSelect
Dim bSkillShotPlayedOnce
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bAutoPlunger
Dim bBonusHeld
Dim bFreePlay
Dim bGameInPlay
Dim bInstantInfo
Dim bBallJustLocked
Dim bLFlipper, bRFlipper
Dim bFlipFlag
Dim bScoreExBall
Dim bScoreReplay

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger
Dim LMagnet, RMagnet

bFreeze = 0
bDisableTable = 0

Sub Table1_Init
	LoadEM
    If B2SOverride=True AND B2SOn=False Then
        Controller.B2SName = cGameName
        Controller.Run()
        B2SOn = True
    End If
    Randomize
    ' Load saved values, highscore, names
    Loadhs
    If Credits>0 OR bFreePlay=True Then DOF 114, DOFOn

    ' DMD display Setup
    If DMD_Option=True Then
        DMDReels_Init
        DMD_Init
        Ramp001.Visible = True
        Ramp002.Visible = True
    End If

    ' UltraDMD Add
    If UltraDMD_Option=True Then LoadUltraDMD

    '***********************************************
    '           IMPULSE PLUNGER - AutoPlunger
    '***********************************************
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .CreateEvents "plungerIM"
    End With

    '***********************************************
    '               MAGNETS
    '***********************************************
	' Magnet  (Witch)
    Set LMagnet = New cvpmTurnTable
    With LMagnet
        .InitTurnTable Magnet1, 90
        .spinCW = False
        .SpinUp = 90
        .SpinDown = 90
        .MotorOn = False
        .CreateEvents "LMagnet"
    End With

	' Magnet  (Munchkin Hole)
    Set RMagnet = New cvpmTurnTable
    With RMagnet
        .InitTurnTable Magnet2, 90
        .spinCW = True
        .SpinUp = 90
        .SpinDown = 90
        .MotorOn = False
        .CreateEvents "RMagnet"
    End With
    '***********************************************

    ' Initialize any other flags
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bAutoPlunger = False
    bGameInPlay = False
    bFreePlay = True ' coins not needed
    bMusicOn = True
    bBonusHeld = False
    bJustStarted = True
    bInstantInfo = False

    bOnTheFirstBall = 0
    bMultiBallMode = 0
    bLFlipper=0
    bRFlipper=0
    BallsOnPlayfield = 0
    BallsInLock = 0
    BallsInHole = 0
    LastSwitchHit = 0
    Tilt = 0
    Tilted = False
    TiltSensitivity = 6

    ' UltraDMD Add
    If UltraDMD_Option=True Then
        UltraDMD_DisplayScene "woz_0.png", "", 13, "", 10, UltraDMD_Animation_None, 3000, UltraDMD_Animation_ScrollOffRight
        ' Show ROM version number
'        UltraDMD_simple "black.png", "WoZ", "NO ROM REQUIRED", 1000
    End If

    ' Clear BackGlass
    ClearBackGlass

	' Table Object Initialization
    TableObject_Setup

    ' GI off
    GiOff

    ' Lower Witch
    WitchLower

    ' Attract Mode Activation
    StartAttractMode

    ' Remove the cabinet rails if in FS mode
    If Table1.ShowDT = False then
        lrail.Visible = False
        rrail.Visible = False
    End If

    ' Topper Activate
    If PUPpack=True Then vpmtimer.addtimer 2000, "DOF 324, DOFPulse '"

    sw5.enabled = True
End Sub

'***********************************************
'           PAUSE TABLE
'***********************************************
Sub Table1_Paused
End Sub

Sub Table1_unPaused
End Sub

'***********************************************
'               EXIT
'***********************************************
Sub Table1_Exit
    Savehs
End Sub

'***********************************************
'               Keys
'***********************************************
Sub Table1_KeyDown(ByVal keycode)
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
	If keycode = PlungerKey Then
        Plunger.PullBack: PlaySoundAtVol "plungerpull", Plunger, 1
        Exit Sub
 	End If

    If Keycode = AddCreditKey Then
        Credits = Credits + 1
        PlaySoundAtVol "fx_coin", Drain, 1
		DOF 114, DOFOn
        If(Tilted = False)Then
            ' DMD Display
            If DMD_Option=True Then
                DMDFlush
                DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "CREDITS: " & Credits), eNone, 500, True, ""
            End If
            ' UltraDMD Add
            If UltraDMD_Option=True Then UltraDMDScoreNow
            If NOT bGameInPlay Then ShowTableInfo
        End If
    End If

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    If Tilted Then Exit Sub

    If bGameInPlay Then
        If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 6:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
        If keycode = MechanicalTilt Then CheckTilt

        If keycode = LeftFlipperKey Then
            ' Turn off Tree Bumpers if both flippers held up
            bLFlipper=1
            If bRFlipper=1 Then
                Bumper1.HasHitEvent=0
                Bumper2.HasHitEvent=0
                Bumper3.HasHitEvent=0
            End If
            ' Switch Flippers during Flipper Frenzy
            If bCB_FlipFrenzy=1 Then SolRFlipper 1 Else SolLFlipper 1
            InstantInfoTimer.Enabled = True
        End If

        If keycode = RightFlipperKey Then
            ' Turn off Tree Bumpers if both flippers held up
            bRFlipper=1
            If bLFlipper=1 Then
                Bumper1.HasHitEvent=0
                Bumper2.HasHitEvent=0
                Bumper3.HasHitEvent=0
            End If
            ' Switch Flippers during Flipper Frenzy
            If bCB_FlipFrenzy=1 Then SolLFlipper 1 Else SolRFlipper 1
            InstantInfoTimer.Enabled = True
        End If

        If keycode = StartGameKey Then
             If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = 1))Then
                If(bFreePlay = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    ' DMD Display
                    If DMD_Option=True Then
                        DMD "", eNone, "_", eNone, CenterLine(2, PlayersPlayingGame & " PLAYERS"), eBlink, "", eNone, 500, True, ""
                    End If
                Else
                    If(Credits > 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
 			            If Credits < 1 and bFreePlay = False Then DOF 114, DOFOff
                        ' DMD Display
                        If DMD_Option=True Then
                             DMD "", eNone, "_", eNone, CenterLine(2, PlayersPlayingGame & " PLAYERS"), eBlink, "", eNone, 500, True, ""
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        ' DMD Display
                        If DMD_Option=True Then
                            DMDFlush
                            DMD "", eNone, CenterLine(1, "CREDITS " & Credits), eNone, CenterLine(2, "INSERT COIN"), eNone, "", eNone, 500, True, ""
                        End If
                    End If
                End If
                ' UltraDMD Add
                If UltraDMD_Option=True Then UltraDMDScoreNow
            End If
        End If
    Else ' If (NOT GameInPlay)
        If keycode = StartGameKey Then

            EndOfGameCleanUp.Enabled = False
            If(bFreePlay = True)Then
                PlayStartupSound
                If(BallsOnPlayfield = 0)Then ResetForNewGame
            Else
                If(Credits > 0)Then
                    PlayStartupSound
                    If(BallsOnPlayfield = 0)Then
                        Credits = Credits - 1
                        If Credits < 1 and bFreePlay = False Then DOF 114, DOFOff
                        ResetForNewGame
                    End If
                Else
                    ' Not Enough Credits to start a game.
                    ' DMD Display
                    If DMD_Option=True Then
                        DMDFlush
                        DMD "", eNone, CenterLine(1, "CREDITS " & Credits), eNone, CenterLine(2, "INSERT COIN"), eBlink, "", eNone, 500, True, ""
                    End If
                    ' UltraDMD Add
'                   ' If UltraDMD_Option=True Then UltraDMDScoreNow
                    ShowTableInfo
                End If
            End If
        End If
    End If ' If (GameInPlay)

'  Used for debugging - requires modifying Test Code appropriately
    If TEST_MODE=True Then TestCode_KeyDown(keycode)

End Sub

Sub Table1_KeyUp(ByVal keycode)
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
	If keycode = PlungerKey Then
        Plunger.Fire: PlaySoundAtVol "plunger",Plunger, 1
    End If

    If hsbModeActive Then
        Exit Sub
    End If

    If Tilted Then Exit Sub

    If bGameInPlay Then

	    If keycode = LeftFlipperKey Then
            bLFlipper=0
            Bumper1.HasHitEvent=1
            Bumper2.HasHitEvent=1
            Bumper3.HasHitEvent=1
            SolLFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                bInstantInfo = False
                ' DMD Display
                If DMD_Option=True Then DMDScoreNow
                ' UltraDMD Add
                If UltraDMD_Option=True Then UltraDMDScoreNow
            End If
        End If

	    If keycode = RightFlipperKey Then
            bRFlipper=0
            Bumper1.HasHitEvent=1
            Bumper2.HasHitEvent=1
            Bumper3.HasHitEvent=1
            SolRFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                bInstantInfo = False
                ' DMD Display
                If DMD_Option=True Then DMDScoreNow
                ' UltraDMD Add
                If UltraDMD_Option=True Then UltraDMDScoreNow
            End If
        End If

    End If    '  If (GamePlay)

'  Used for debugging - requires modifying Test Code appropriately
    If TEST_MODE=True Then TestCode_KeyUp(keycode)

End Sub


'***********************************************
'              FLIPPERS
'***********************************************
Dim bLFlipActive, bRFlipActive
bLFlipActive=0
bRFlipActive=0
Sub SolLFlipper(Enabled)
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
    If Enabled Then
        bLFlipActive = 1
        LeftFlipper.RotateToEnd
        MunchkinFlipper.RotateToEnd
        PlaySoundAtVol SoundFXDOF("fx_flipperup",101,DOFOn,DOFFlippers), LeftFlipper, VolFlip
        PlaySoundAtVol SoundFXDOF("fx_flipperup",101,DOFOn,DOFFlippers), MunchkinFlipper, VolFlip
        If bSkillshotReady = 0 AND bSkillshotSelect = 1 AND bTOTOMode=0 AND bTNPLHMode=0 Then
            MoveLaneLightLeft
        End If
        If bCB_NoHoldFlip=1 Then NoHold.enabled=True
    Else
        bLFlipActive = 0
        LeftFlipper.RotateToStart
        MunchkinFlipper.RotateToStart
        PlaySoundAtVol SoundFXDOF("fx_flipperdown",101,DOFOff,DOFFlippers), LeftFlipper, VolFlip
        PlaySoundAtVol SoundFXDOF("fx_flipperdown",101,DOFOff,DOFFlippers), MunchkinFlipper, VolFlip
        LeftFlipper.RotateToStart
        MunchkinFlipper.RotateToStart
    End If
    If bCB_LinkedFlip=1 AND bFlipFlag=0 Then bFlipFlag=1: SolRFlipper(Enabled): Exit Sub
    bFlipFlag=0
End Sub


Sub SolRFlipper(Enabled)
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
    If Enabled Then
        bRFlipActive = 1
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
        CastleFlipper.RotateToEnd
        PlaySoundAtVol SoundFXDOF("fx_flipperup",102,DOFOn,DOFFlippers), RightFlipper, VolFlip
        PlaySoundAtVol SoundFXDOF("fx_flipperup",102,DOFOn,DOFFlippers), RightFlipper1, VolFlip
        PlaySoundAtVol SoundFXDOF("fx_flipperup",102,DOFOn,DOFFlippers), CastleFlipper, VolFlip
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
        CastleFlipper.RotateToEnd
        If bSkillshotReady = 0 AND bSkillshotSelect = 1 AND bTOTOMode=0 AND bTNPLHMode=0 Then
            MoveLaneLightRight
        End If
        If bCB_NoHoldFlip=1 Then NoHold.enabled=True
    Else
        bRFlipActive = 0
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
        CastleFlipper.RotateToStart
        PlaySoundAtVol SoundFXDOF("fx_flipperdown",102,DOFOff,DOFFlippers), RightFlipper, VolFlip
        PlaySoundAtVol SoundFXDOF("fx_flipperdown",102,DOFOff,DOFFlippers), RightFlipper1, VolFlip
        PlaySoundAtVol SoundFXDOF("fx_flipperdown",102,DOFOff,DOFFlippers), CastleFlipper, VolFlip
    End If
    If bCB_LinkedFlip=1 AND bFlipFlag=0 Then bFlipFlag=1: SolLFlipper(Enabled): Exit Sub
    bFlipFlag=0
End Sub

' Flipper overlay (slippers) motion control
Sub FlipperTimer_Timer
	flipperL.RotY = LeftFlipper.CurrentAngle
    flipperR.RotY = RightFlipper.CurrentAngle
End Sub

' flippers hit Sound
Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

Sub NoHold_Timer
    Me.Enabled = 0
    SolLFlipper 0
    SolRFlipper 0
End Sub


Sub MoveLaneLightLeft
    If l253.state=1 Then
        Flashers(flash_dict.Item(l104))=1:Flashers(flash_dict.Item(l105))=0
        SetLightColor l253,"white",0:SetLightColor l254,"white",0:LightState(105)=0:Flashers(flash_dict.Item(l105))=0
        SetLightColor l252, "white", 1:LightState(104)=1:Flashers(flash_dict.Item(l104))=1
    End If
End Sub

Sub MoveLaneLightRight
    If l252.state=1 Then
        Flashers(flash_dict.Item(l104))=0:Flashers(flash_dict.Item(l105))=1
        SetLightColor l253,"white",1:SetLightColor l254,"white",1:LightState(105)=1:Flashers(flash_dict.Item(l105))=1
        SetLightColor l252, "white", 0:LightState(104)=0:Flashers(flash_dict.Item(l104))=0
    End If
End Sub

'***********************************************
'                  SLING SHOT ANIMATIONS
'***********************************************
' Rstep, Lstep and Tstep  are the variables that
' increment the animation
'***********************************************
Dim RStep, Lstep, Tstep
Sub TopSlingShot_Slingshot
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
    If Tilted Then Exit Sub
    PlaySoundAtVol SoundFXDOF("fx_slingshot", 105, DOFPulse, DOFContactors), sling3, 1
    DOF 106, DOFPulse
    SwitchHitEvent 3
    TSling.Visible = 0
    TSling1.Visible = 1
    sling3.TransZ = -20
    TStep = 0
    TopSlingShot.TimerEnabled = 1
End Sub

Sub TopSlingShot_Timer
    Select Case TStep
        Case 3:TSLing1.Visible = 0:TSLing2.Visible = 1:sling3.TransZ = -10
        Case 4:TSLing2.Visible = 0:TSLing.Visible = 1:sling3.TransZ = 0:TopSlingShot.Timerenabled = 0
    End Select
    TStep = TStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
    If Tilted Then Exit Sub
    PlaySoundAtVol SoundFXDOF("fx_slingshot", 105, DOFPulse, DOFContactors), sling1, 1
    DOF 106, DOFPulse
    SwitchHitEvent 3
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.Timerenabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAtVol SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors), sling2, 1
    DOF 104, DOFPulse
    SwitchHitEvent 3
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.Timerenabled = 0
    End Select
    LStep = LStep + 1
End Sub


'***********************************************
'                  BUMPERS
'***********************************************
Dim b1Step, b2Step, b3Step
Sub Bumper1_Hit
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
    If Tilted Then Exit Sub
    TreeBumper1.TransZ= -15 : b1Step = 0
    b1StepMove.Enabled = True
    PlaySoundAtVol SoundFXDOF("fx_Bumper4",117,DOFPulse,DOFContactors), ActiveBall, 1
    SwitchHitEvent 49
	B1L1.State = 1:B1L2. State = 1
	Me.Timerenabled = 1
End Sub

Sub Bumper1_Timer:B1L1.State = 0:B1L2. State = 0:Me.Timerenabled = 0:End Sub
Sub b1StepMove_Timer
    Select Case b1Step
        Case 1:TreeBumper1.TransZ = -10
        Case 2:TreeBumper1.TransZ = -5
        Case 3:TreeBumper1.TransZ = 0:b1StepMove.Enabled = False
    End Select
    b1Step = b1Step + 1
End Sub


Sub Bumper2_Hit
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
    If Tilted Then Exit Sub
    TreeBumper2.TransZ= -15 : b2Step = 0
    b2StepMove.Enabled = True
    PlaySoundAtVol SoundFXDOF("fx_Bumper4",107,DOFPulse,DOFContactors), ActiveBall, 1
    SwitchHitEvent 50
	B2L1.State = 1:B2L2. State = 1
	Me.Timerenabled = 1
End Sub

Sub Bumper2_Timer:B2L1.State = 0:B2L2. State = 0:Me.Timerenabled = 0:End Sub
Sub b2StepMove_Timer
    Select Case b2Step
        Case 1:TreeBumper2.TransZ = -10
        Case 2:TreeBumper2.TransZ = -5
        Case 3:TreeBumper2.TransZ = 0:b2StepMove.Enabled = False
    End Select
    b2Step = b2Step + 1
End Sub


Sub Bumper3_Hit
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
    If Tilted Then Exit Sub
    TreeBumper3.TransZ= -15 : b3Step = 0
    b3StepMove.Enabled = True
    PlaySoundAtVol SoundFXDOF("fx_Bumper4",115,DOFPulse,DOFContactors), ActiveBall, 1
    SwitchHitEvent 51
	B3L1.State = 1:B3L2. State = 1
	Me.Timerenabled = 1
End Sub

Sub Bumper3_Timer:	B3L1.State = 0:B3L2. State = 0:Me.Timerenabled = 0:End Sub
Sub b3StepMove_Timer
    Select Case b3Step
        Case 1:TreeBumper3.TransZ = -10
        Case 2:TreeBumper3.TransZ = -5
        Case 3:TreeBumper3.TransZ = 0:b3StepMove.Enabled = False
    End Select
    b3Step = b3Step + 1
End Sub


Sub Bumper4_Hit
    If bFreeze=1 OR bDisableTable=1 Then Exit Sub
    If Tilted Then Exit Sub
    PlaySoundAtVol SoundFXDOF("fx_Bumper4",103,DOFPulse,DOFContactors), ActiveBall, 1
    ' Used for MX Effect
    DOF 326,DOFPulse
    SwitchHitEvent 41
	B4L1.State = 1:B4L2. State = 1
	Me.Timerenabled = 1
End Sub

Sub Bumper4_Timer:B4L1.State = 0:B4L2. State = 0:Me.Timerenabled = 0:End Sub


'***********************************************
'              SW PLUNGER - AutoFire
'***********************************************
Sub swPlungerRest_Hit
    ' turn on Launch light if there is one
    DOF 156, DOFOn
    SetLightColor L001, "red", 2
    bBallInPlungerLane = True
    PlaySoundAtVol "fx_sensor", Plunger, 1

    'be sure to update the Scoreboard after the animations, if any
    If UltraDMD_Option=True Then UltraDMDScoreNow

    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        PlungerIM.AutoFire
        AutoFireCheck.Enabled = True
        DOF 111, DOFPulse	' AutoPlunger
		DOF 112, DOFPulse	' Strobe
        bAutoPlunger = False
    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    End If
    'Start the Selection of the skillshot if ready
    If bSkillShotReady Then
        swPlungerRest.Timerenabled = 1 ' this is a new ball, so show launch ball message if inactive for 6 seconds
    End If
End Sub


' The ball is released from the plunger turn off some flags and check for skillshot
Sub swPlungerRest_UnHit
    DOF 156, DOFOff
    SetLightColor L001, "red", 0
    bBallInPlungerLane = False
    swPlungerRest.Timerenabled = 0 'stop the launch ball timer if active
    If bMultiballMode=0 Then
        ' New ball message?
    End If
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds
Sub swPlungerRest_Timer
    ' Inactivity Message when Ball @ Plunger
    swPlungerRest.Timerenabled = 0
End Sub

'***********************************************
'              BALL  SAVER
'***********************************************
Sub EnableBallSaver(seconds)
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
    BallSaverTimerExpired.Interval = 1000 * seconds
    BallSaverTimerExpired.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    L49.BlinkInterval = 160
    SetLightColor L49, "orange", 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
Sub BallSaverTimerExpired_Timer
    BallSaverTimerExpired.Enabled = False
    ' clear the flag
    bBallSaverActive = False
    ' if you have a ball saver light then turn it off at this point
    L49.State = 0
End Sub

Sub BallSaverSpeedUpTimer_Timer
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    L49.BlinkInterval = 80
    L49.State = 2
End Sub

'***********************************************
'               TILT
'***********************************************
Sub CheckTilt                                       ' Called when table is nudged
    Tilt = Tilt + TiltSensitivity                   ' Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt > TiltSensitivity)AND(Tilt < 15)Then    ' Display warning
        If UltraDMD_Option=True Then UltraDMD_DisplayImage("Tilt_Warning.png")
        If DMD_Option=True Then DMD "", eNone, "_", eNone, _
           CenterLine(2, "CAREFUL!"), eBlinkFast, "", eNone, 500, True, ""
        DOF 400, DOFOn
        vpmtimer.addtimer 500, "DOF 400, DOFOff '"
    End if
    If Tilt > 15 Then                               ' If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        If UltraDMD_Option=True Then UltraDMD_DisplayImage("Tilt.png"):UltraDMD_Freeze
        If DMD_Option=True Then DMDFlush:DMD "", eNone, "", eNone, "", eNone, _
           CenterLine(3, "TILT!"), eBlinkFast, 200, False, ""
        DisableTable True
        DOF 401, DOFOn
        TiltRecoveryTimer.Enabled = True            ' Start the Tilt delay to check for all the balls drained
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' Decrease Tilt
    If Tilt > 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        'turn off GI
        GiOff
        ' PF lights - Blink Red
        If Tilted Then StartTiltSequence
        'Disable slings, bumpers etc
        SolLFlipper 0
        SolRFlipper 0
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
        bDisableTable = 1
        If bBallInPlungerLane=True Then PlungerIM.AutoFire:AutoFireCheck.Enabled = True
    Else
        'turn back on GI and the lights
        GiOn
        bDisableTable = 0
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        If UltraDMD_Option=True Then UltraDMD_Freeze
    End If
End Sub

Sub TiltRecoveryTimer_Timer
    ' if all the balls have been drained then..
    If (BallsOnPlayfield = 0) Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table was tilted)

        ' Reset the internal tilted flag (this will also
        ' set TiltWarnings back to zero) which is useful if we are changing players
        Tilted = False
        Tilt = 0
        DOF 401, DOFoff

        EndOfBall
        ' Reset the internal tilted flag (this will also
        ' set TiltWarnings back to zero) which is useful if we are changing player LOL

        If DMD_Option=True Then DMDScoreNow
        If UltraDMD_Option=True Then UltraDMD_UnFreeze
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in ten seconds)
End Sub


' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************
' lost a ball ;-( Check to see how many balls are on the playfield.
' If only one then decrement the remaining count AND test for End of game
' If more than 1 ball (multi-ball) then destroy ball but don't create a new one
Sub Drain_Hit
    Dim i

    PlaySoundAtVol "drain", Drain, 1
	Drain.DestroyBall

    If bSOTR=0 Then StopSound Song

    ' Disable the Table only lasts until ball drains
    bDisableTable = 0

    ' only for debugging - Add balls from the debug window
    If bDrainFlag Then Exit Sub

    ' Lock up Table for BTWW Finish
    If bFreeze=1 Then Exit Sub

    If bGameInPlay=False Then Exit Sub

    If bSkillshotReady=1 Then SkillShotStop

    BallsOnPlayfield = BallsOnPlayfield - 1
	DOF 113, DOFPulse

   ' Table Tilted
	If Tilted Then Exit Sub

    ' Check for BTWW Mode Melting in progress
    If BTWWCnt=7 Then
        bBTWW_check = 1
        Exit Sub
    Else
        If BTWWCnt>0 Then
            BTWW_end
        Else
            If bBTWW_check=1 Then ClearBG_BTWW:bBTWW_check = 0
        End If
    End If
    If bBTWW=1 AND BTWWCnt=0 Then BTWW_end


    ' Check for TOTO Mode Activated
    If l37.State=1 AND l38.State=1 AND l39.State=1 AND l40.State=1 _
                   AND bDingDong=0 AND bBTWW=0 AND SOTR_capture=0 Then
        If bTOTOMode=0 Then
            TotoMode
        Else
            Wizard_restore
            RestoreOZLaneFlashers
        End If
        If bTOTOExhausted=0 Then Exit Sub
    End If

    ' Check for TNPLH Mode Activated
    If l59.State=1 AND l60.State=1 AND l61.State=1 AND l62.State=1 AND l63.State=1 _
                   AND bTNPLHMode=0 AND bDingDong=0 AND bBTWW=0 AND SOTR_capture=0 Then
        If bTNPLHMode=0 Then
            TNPLHMode
        Else
            Wizard_restore
            RestoreOZLaneFlashers
        End If
        Exit Sub
    End If

    If BallsOnPlayfield>1 AND bMultiBallMode=1 Then Exit Sub

    ' Is the ball saver active?
    If bBallSaverActive=True Then
        ' create a new ball in the shooters lane
        ' use the Addmultiball in case other multiballs are queued to be ejected
        AddMultiball 1
        ' Activate autoplunger
        bAutoPlunger = True
        ' Display message and play sound
        If DMD_Option=True Then DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "BALL SAVED"), eBlinkfast, 800, True, ""
        If UltraDMD_Option=True Then UltraDMD_DisplayImage("Ball_Saved.png")
        Exit Sub
    End If

    ' Cancel multiball mode if on last ball (ie. lost all other balls)
    If BallsOnPlayfield=1 AND mBalls2Eject=0 AND bMultiBallMode=1 Then
        ' not in multiball mode any more
        ' turn off any multiball specific lights
        If bSOTR=1 AND bSOTR_init=1 Then Exit Sub
        bMultiBallMode = 0
        If bDingDong=0 Then
            If bMunchkinLand=1 OR bMunchkinFrenzy=1 OR bMunchkinLollipop=1 Then
                ChangeGi "red"
            Else
                If bHauntedShots=1 OR bHauntedHoles=1 OR bHauntedTargets=1 OR bHauntedBumpers=1 Then
                    ChangeGi "purple"
                Else
                    ChangeGi "white"
                End If
            End If
        End If
        If bECMultiB=1 Then ECMB_Done
        If bMunchkinMultiBall=1 Then
            bMunchkinMultiBall=0:bMunchkinMB=1
            SetLightColor l161,"pink",0
            SetLightColor l96,"white",0
            RestoreBG_Rainbow
        End If
        If bRescueMB=1 Then
            bRescueMB=0
            ClearBG_RescueMB
            RescueMB_End
            If bCastlePFActive=0 Then
               RestoreBG_Castle
            Else
               RestoreBG_Rescue
            End If
        End If
        If bCB_LightsON=1 OR bCB_LightsOFF=1 OR bCB_LinkedFlip=1 OR bCB_NoHoldFlip=1 OR bCB_FlipFrenzy=1 Then
            CrystalBallModeTimer.interval = 50
            CrystalBallModeTimer.enabled = True
        End If
        If bHauntedMB=1 Then
            hauntedTimeLeft=1
            HauntedModeTimer.interval = 50
            HauntedModeTimer.enabled = True
        End If
        Exit Sub
    End If
    If bSOTR=1 AND bSOTR_init=1 Then SOTR_end

    vpmtimer.addtimer 250, "Drain_Settled '"
    debug.print "Drain Settling Delay"

End Sub


Sub Drain_Settled
    Dim obj

    ' Last Ball On Playfield?
    If BallsOnPlayfield<1 Then
        ' MODE CHECKING - Send it to the Associated End of Mode Timer or Routine

        If bDingDong=1 Then DingDong_End
        If bTOTOMode=1 Then
            TOTOMode_End
            If bSOTR=0 Then
                Kicker_TOTO.enabled=False
                debug.print "Kicker.TOTO OFF - Drain-Settled"
            Else
                Kicker_TOTO.enabled=True: hold_time=30000
            End If
            If totoCnt>0 Then totoCnt = totoCnt - 1
        End If
        If bTNPLHMode=1 Then
            TNPLHMode_End
            If bSOTR=0 Then
                Kicker_TOTO.enabled=False
            Else
                Kicker_TOTO.enabled=True: hold_time=30000
            End If
        End If
        If bHouseSpin=1 Then
            TwisterModeHold.interval = 50
            TwisterModeHold.enabled = True
        End If
        If bHauntedMode=1 Then
            hauntedTimeLeft=1
            HauntedModeTimer.interval = 50
            HauntedModeTimer.enabled = True
        End If
        If bMunchkinMode=1 Then
            bMunchkinMode=0: munchkinTimeLeft=1
            MunchkinModeTimer.interval = 50
            MunchkinModeTimer.enabled = True
        End If
        If bFireballFrenzy=1 AND bFireballFrenzyStarted=1 Then FireballFrenzyStop
'        If bWitchHU=1 Then
        WitchHUStop
        If bHU=1 Then HUStop
        ' MODE CHECKING END

        ' Clear Haunted Lights
        For each obj in Lights_Haunted
            SetLightColor obj,"purple",0
            Controller.B2SSetData bg_dict.Item(obj), obj.State
        Next

        BallsOnPlayfield=0
        ' Clear OZ Lane LEDs (Flashers) - Need Some Time To Complete
        ClearOZLaneFlashers
        SetLightColor l252,"white",0
        SetLightColor l253,"white",0:SetLightColor l254,"white",0
        ChangeGi "white"

        ' BackGlass Cleanup - compensates for asynchronouse issues (shouldn't be needed)
        vpmtimer.addtimer 1500, "ClearBG_NewBall '"
        EndOfBall

    End If
End Sub


'***********************************************
'               GAME  HANDLING
'***********************************************
' Setup for a new game
Sub ResetForNewGame
    Dim i

    ' Variable/Flag setup
    bGameInPlay = True
    bOnTheFirstBall = 1
    mBalls2Eject = 0
    bExtraBallWonThisBall = False

    StopAttractMode	' turn off attract mode
    TurnOffPlayfieldLights
    GiOn            ' turn on GI
    BackGlass_AttractMode_OFF

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
    PlayersPlayingGame = 1

    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        If TEST_AUTOPLAY=True Then
            BallsRemaining(i) = TEST_BallsPerGame
        Else
            BallsRemaining(i) = BallsPerGame
        End If
        ExtraBallsAwards(i) = 0

    Next

    ' initialise any other flags
    Tilt = 0

    ' UltraDMD
    tickCount = 0
    delay = 0
    bWait = False

    ' initialise Game variables
    Game_Init

    ' This is used to delay the start of a game to allow any attract sequence to
    ' complete.  When it expires it creates a ball for the player to start playing with
    vpmtimer.addtimer 200, "FirstBall '"
End Sub

' First Ball of a New Game
Sub FirstBall
    Dim i
    bSkillShotPlayedOnce = False
    ' reset the table for a new ball
    ResetForNewPlayerBall
    ResetFirstBallModes
    ' Capture initial Player Specific State Information
    For i=1 To 4
        Capture_PlayerState i
    Next
    ' create a new ball in the shooters lane
    CreateNewBall
   ' UltraDMD Add
   If UltraDMD_Option=True Then UltraDMD_Timer1.Enabled = True
End Sub

' Player with New Ball
' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))
Sub ResetForNewPlayerBall
    ' Zero Score Display
    AddScore 0

    ' set the current players bonus multiplier back down to 1X
    BonusMultiplier(CurrentPlayer) = 1

    ' reset Player specific items
    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False

    'Set any table specific items for new ball - (skill shots, ...)
    NewBall_Init

    'This is a new ball, so activate the ballsaver if ON
     If BallSaveOption_ON = True Then bBallSaverReady = True
End Sub


' Create a new ball on the Playfield
Sub CreateNewBall
	BallRelease.CreateBall
'    debug.print "bTOTOMode = " & bTOTOMode
'    debug.print "bTNPLHMode = " & bTNPLHMode

    ' Has a ball been locked?
    If bBallJustLocked>0 Then
        bBallJustLocked = bBallJustLocked - 1
    Else
        BallsOnPlayfield = BallsOnPlayfield + 1
    End If

    ' kick it out.
	BallRelease.Kick 90,6.
    PlaySoundAtVol SoundFXDOF("fx_Ballrel",110,DOFPulse,DOFContactors), BallRelease, 1

    ' if there is 2 or more balls then set the multiball flag
    ' (remember to check for locked balls and other balls used for animations)
    ' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield > 1 Then
        DOF 111, DOFPulse 	' AutoPlunger
		DOF 112, DOFPulse	' Strobe
'		DOF 126, DOFPulse	' Beacon for multiball
        bMultiBallMode = 1
        bAutoPlunger = True
'        PlaySong "m_multiball"
    End If
End Sub

' Checks for loss of Ball
Dim bBallLossCheck
Sub BallLossCheck_Timer
    debug.print "Ball Loss Timer Fired"
    If bBallLossCheck=1 Then
        If BallsOnPlayfield>0 AND bBallInPlungerLane=False _
           AND bLFlipActive=0 AND bRFlipActive=0 Then
            ClearBG_NewBall
            Drain_Hit
            debug.print "BALL LOSS DETECTED"
            If UltraDMD_Option=True Then
                UltraDMDBlink "black.png", " ", "BALL LOSS DETECTED", 100, 10
            End If
        End If
        If TEST_MODE=True Then
            CenterPost(0)
            debug.print "CenterPost Down"
        End If
        bBallLossCheck=0
     Else
        bBallLossCheck=1
     End If
End Sub


' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table
Dim mBalls2Eject
Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
    'and eject the first ball
    CreateMultiballTimer_Timer
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer
    ' Check to see if there are more balls to eject, if not stop the timer
    If mBalls2Eject<=0 Then
        CreateMultiballTimer.Enabled = False
    Else
        ' Check for too many balls on playfield
        If BallsOnPlayfield>=MaxMultiballs Then Exit Sub
        ' Wait if there is a ball in the plunger lane.
        If bBallInPlungerLane=False Then
            CreateNewBall
            mBalls2Eject = mBalls2Eject - 1
        End If
        SwitchFlags(125) = 0
        PlungerIM.AutoFire
        AutoFireCheck.Enabled = True
    End If
End Sub

Sub AutoFireCheck_Timer
    If bBallInPlungerLane=True AND bMultiBallMode=1 Then
        PlungerIM.AutoFire
    Else
        If SwitchFlags(125)>0 Then Me.enabled = 0
'        If SwitchFlags(125)>0 OR bMultiBallMode=0 Then Me.enabled = 0
    End If
End Sub


Sub EndOfBall
    Dim secdelay
    ' UltraDMD Add
    If UltraDMD_Option=True Then UltraDMD_Freeze
    Dim TotalBonus, tmpBonus
    TotalBonus = 0
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = 0

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)

    If(Tilted = False)Then

        ' Calculate the Bonuses. This table uses several bonuses
        ' Yellow Brick Road Bonus
        YBRBonus = modeYBR_cnt * 25
        TotalBonus = TotalBonus + YBRBonus

        ' Munchkin BONUS
        tmpBonus = twisterbonus + munchkinBonus
        munchkinMB_Bonus = munchkinMB_Bonus + tmpBonus
        munchkinMB_Bonus = munchkinX * munchkinMB_Bonus
        If bMunchkinMB=1 Then tmpbonus = munchkinMB_Bonus: munchkinMB_Bonus = 0: bMunchkinMB=0: munchkinX=1
        TotalBonus = TotalBonus + tmpBonus

        ' SOTR Rainbow Bonus
        SOTRBonus = SOTRBonus * SOTRmultiplier
        TotalBonus = TotalBonus + SOTRBonus

        ' Haunted Bonus
        ' TOTAL BONUS
        TotalBonus = TotalBonus*Multipliers(1) + hauntedBonus + hauntedMB_Bonus
        BonusMultiplier(CurrentPlayer)=Multipliers(1)

        If DMD_Option=True Then
           dmdflush
           DMD "", eNone, CenterLine(1, "BONUS"), eNone, "", eNone, "", eNone, 400, False, ""
           If YBRBonus>0 Then
               DMD "", eNone, CenterLine(1, "BONUS"), eNone, _
                              CenterLine(2, "YELLOW BRICK ROAD"), eNone, _
                              CenterLine(3, FormatScore(YBRBonus)), eBlinkFast, 400, False, "gb_Bonus2"
           End If
           DMD "", eNone, CenterLine(1, "BONUS"), eNone, _
                          CenterLine(2, "BONUS MULTIPLIER"), eNone, _
                          CenterLine(3, FormatScore(Multipliers(1))), eBlinkFast, 400, False, "gb_Bonus3"
           If tmpbonus>0 Then
               DMD "", eNone, CenterLine(1, "BONUS"), eNone,_
                              CenterLine(2, "MUNCHKIN"), eNone, _
                              CenterLine(3, FormatScore(tmpBonus)), eBlinkFast, 400, False, "gb_Bonus4"
           End If
           If hauntedBonus>0 Then
               DMD "", eNone, CenterLine(1, "BONUS"), eNone, _
                              CenterLine(2, "HAUNTED"), eNone, _
                              CenterLine(3, FormatScore(hauntedBonus)), eBlinkFast, 400, False, "gb_Bonus4"
           End If
           If hauntedMB_Bonus>0 Then
               DMD "", eNone, CenterLine(1, "BONUS"), eNone, _
                              CenterLine(2, "HAUNTED MBALL"), eNone, _
                              CenterLine(3, FormatScore(hauntedMB_Bonus)), eBlinkFast, 400, False, "gb_Bonus4"
           End If
           If SOTRBonus>0 Then
               DMD "", eNone, CenterLine(1, "BONUS"), eNone, _
                              CenterLine(2, "RAINBOW"), eNone, _
                              CenterLine(3, FormatScore(SOTRBonus)), eBlinkFast, 400, False, "gb_Bonus4"
           End If
           ' TOTAL BONUS
           DMD "", eNone, CenterLine(1, "BONUS"), eNone, _
           CenterLine(2, "TOTAL BONUS X" &BonusMultiplier(CurrentPlayer)), eNone, _
           CenterLine(3, FormatScore(TotalBonus)), eBlinkFast, 1000, True, "gb_Bonus5"
        End If
        secdelay=1
        If UltraDMD_Option=True Then
            If YBRBonus>0 Then UltraDMD_simple "black.png", "YELLOW BRICK ROAD", YBRBonus, 1500: secdelay=secdelay+1
            If tmpbonus>0 Then UltraDMD_simple "black.png", "MUNCHKIN BONUS", tmpbonus, 1500: secdelay=secdelay+1
            If hauntedBonus>0 Then UltraDMD_simple "black.png", "HAUNTED BONUS", hauntedBonus, 1500: secdelay=secdelay+1
            If hauntedMB_Bonus>0 Then UltraDMD_simple "black.png", "HAUNTED MULTIBALL", hauntedMB_Bonus, 1500: secdelay=secdelay+1
            If SOTRBonus>0 Then UltraDMD_simple "black.png", "RAINBOW", SOTRBonus, 1500: secdelay=secdelay+1
            If TotalBonus>0 Then UltraDMD_simple "black.png", "TOTAL BONUS X" &BonusMultiplier(CurrentPlayer), TotalBonus, 1500
        End If
        secdelay = secdelay*1500+1000

        debug.print "Current Player Score = " & Score(CurrentPlayer)
        debug.print "Total Bonus to be added: " & TotalBonus
        AddScore_points TotalBonus

        ' Clear Single Ball Bonuses
        YBRBonus = 0
        hauntedBonus=0:hauntedMB_Bonus=0
        twisterbonus=0: munchkinBonus=0
        SOTRbonus = 0
        TotalBonus = 0

        ' add a bit of a delay to allow for the bonus points to be shown & read
        vpmtimer.addtimer secdelay, "EndOfBall2 '"

        ' UltraDMD Add
        If UltraDMD_Option=True Then vpmtimer.addtimer secdelay, "UltraDMD_UnFreeze '"
    Else
        vpmtimer.addtimer 100, "EndOfBall2 '"
        ' UltraDMD Add
        If UltraDMD_Option=True Then vpmtimer.addtimer 50, "UltraDMD_UnFreeze '"
    End If
End Sub


' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the currentplayer)
Sub EndOfBall2

    DisableTable False 'enable again bumpers and slingshots

    ' has the player won an extra-ball ? (might be multiple outstanding)
    If(ExtraBallsAwards(CurrentPlayer) <> 0)Then
        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            L49.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
        If PUPpack=True Then
            DOF 81,DOFON
            vpmtimer.addtimer 3000, "DOF 81,DOFOff '"
        Else
            PlaySound "voc2_Extra_Ball"
        End If
        If DMD_Option=True Then DMD "", eNone, "", eNone, "", eNone, CenterLine(3, ("EXTRA BALL")), eBlink, 1000, True, "vo_extraball"
        If UltraDMD_Option=True Then UltraDMD_simple "Extra_Ball.png", " ", "", 2500
        ' End Of Ball Processing with NextPlayer=CurrentPlayer & Game Not Over
        AddScore 0
        ResetForNewPlayerBall

        ' Create a new ball in the shooters lane but give time for the Extra Ball Message
        vpmtimer.addtimer 2000, "CreateNewBall '"
    Else ' no extra balls
        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

        ' Check for last ball
        If(BallsRemaining(CurrentPlayer) = 0) Then
            'No More Balls for current player, Check for High Score
            CheckHighScore
        End If
        If hsbModeActive=False Then EndOfBallComplete
    End If
End Sub


' This function is called when the end of bonus display
' (or high score entry finished) AND it is either the end of the game or
' time to move on to the next player (or the next ball of the same player)
Sub EndOfBallComplete
    Dim NextPlayer

    ' are there multiple players playing this game ?
    If PlayersPlayingGame>1 Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If NextPlayer>PlayersPlayingGame Then
            NextPlayer = 1
        End If
        bNoCapture=0
    Else
        NextPlayer = CurrentPlayer
        bNoCapture=1
    End If

    debug.print "Balls Remaining = " & BallsRemaining(NextPlayer)
    ' is it the end of the game ? (all balls been lost for all players)
    If BallsRemaining(CurrentPlayer)<=0 AND BallsRemaining(NextPlayer)<=0 Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

        ' set the machine into game over mode
        ' put a Game Over message on the desktop/backglass
        If UltraDMD_Option=True Then UltraDMDFlush
        If PUPpack=True Then
            DOF 300,DOFpulse
        Else
            Controller.B2SSetData 200,1
            PlaySound "voc16_WritingWitch"
        End If
        vpmtimer.addtimer 9000, "EndOfGame '"
    Else
        ' Capture Player State - Lights, Modes and other inter-Ball Progress
        If bNoCapture=0 Then Capture_PlayerState CurrentPlayer

        ' set the next player
        CurrentPlayer = NextPlayer

        ' Restore Player's Lights & Mode States from last ball if not on first ball
        If bOnTheFirstBall=0 AND bNoCapture=0 Then
            vpmtimer.addtimer 500, "Restore_Delayed '"
        Else
            ' Reset the playfield for the new player (or new ball)
            ResetForNewPlayerBall

            ' Create a new ball
            CreateNewBall
        End If

        ' Play a sound if more than 1 player
        If PlayersPlayingGame > 1 Then
            If (BallsRemaining(CurrentPlayer)MOD 2=1) AND (Int(Rnd*2+0.35)=1) Then
                PlaySound "vo_playerlong" &CurrentPlayer
            Else
                PlaySound "vo_player" &CurrentPlayer
            End If
            If DMD_Option=True Then DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "PLAYER " &CurrentPlayer), eNone, 800, True, ""
        End If

    End If
End Sub

Sub Restore_Delayed
    ' Restore Player's Lights & Mode States from last ball if not on first ball
    If bOnTheFirstBall=0 AND bNoCapture=0 Then Restore_PlayerState

    ' reset the playfield for the new player (or new ball)
    ResetForNewPlayerBall

    ' Create a new ball
    CreateNewBall
End Sub


' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..
Sub EndOfGame
    bGameInPLay = False
    bJustStarted = False
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all modes - eject locked balls
    ' most of the modes/timers terminate at the end of the ball
    ModeTerminate_EndOfGame

    ' Clear BackGlass
    ClearBackGlass

    ' set any lights for the attract mode
    GiOff
    WitchLower
    If TEST_MODE=True Then
        CenterPost(0)
        debug.print "CenterPost Down"
    End If
    StartAttractMode
    ' you may wish to light any Game Over Light you may have
    ' UltraDMD Add
    '    If UltraDMD_Option=True Then UltraDMD_Timer1.Enabled = True
    debug.print "End Of Game Done"
End Sub


Function Balls
    Dim tmp

    If TEST_AUTOPLAY=True Then
        tmp = TEST_BallsPerGame - BallsRemaining(CurrentPlayer) + 1
        If tmp > TEST_BallsPerGame Then
            Balls = TEST_BallsPerGame
        Else
            Balls = tmp
        End If
    Else
        tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
        If tmp > BallsPerGame Then
            Balls = BallsPerGame
        Else
            Balls = tmp
        End If
    End If
End Function


'**********************************************************************************************
'                  INPUTS - SWITCHES, KICKERS & TRIGGERS
'**********************************************************************************************
'***********************************************
'             SWITCHES
'***********************************************
Sub sw4_Hit:SwitchHitEvent 4: End Sub      ' Entry Ramp to Munchkin Playfield
Sub sw9_Hit:SwitchHitEvent 9: End Sub      ' Spinner entry to Crystal Ball VUK
Sub sw16_Hit:SwitchHitEvent 16: End Sub    ' Right Orbit Enter
Sub sw20_Hit:SwitchHitEvent 20: End Sub    ' Castle Exit  ( On Wire Ramp )
Sub sw24_Hit:SwitchHitEvent 24: End Sub    ' Left Orbit Enter
Sub sw27_Hit:SwitchHitEvent 27: End Sub    ' Left Return Lane
Sub sw28_Hit:SwitchHitEvent 28: End Sub    ' Left Outlane Lane
Sub sw29_Hit:SwitchHitEvent 29: End Sub    ' Crystal BALL Target - B
Sub sw30_Hit:SwitchHitEvent 30: End Sub    ' Crystal BALL Target - A
Sub sw31_Hit:SwitchHitEvent 31: End Sub    ' Crystal BALL Target - L
Sub sw32_Hit:SwitchHitEvent 32: End Sub    ' Crystal BALL Target - L
Sub sw35_Hit:SwitchHitEvent 35: End Sub    ' Right Return Lane Lane
Sub sw37_Hit:SwitchHitEvent 37: End Sub    ' Right Outer Lane - TOTO Top
Sub sw38_Hit:SwitchHitEvent 38: End Sub    ' Right Outer Lane - TOTO Ctr
Sub sw39_Hit:SwitchHitEvent 39: End Sub    ' Right Outer Lane - TOTO Ctr
Sub sw40_Hit:SwitchHitEvent 40: End Sub    ' Right Outer Lane - TOTO Bottom
Sub sw42_Hit:SwitchHitEvent 42: End Sub    ' State Fair Rubber
Sub sw43_Hit:SwitchHitEvent 43: End Sub    ' TNPLH Target - THERE'S
Sub sw44_Hit:SwitchHitEvent 44: End Sub    ' TNPLH Target - NO
Sub sw45_Hit:SwitchHitEvent 45: End Sub    ' TNPLH Target - PLACE
Sub sw46_Hit:SwitchHitEvent 46: End Sub    ' TNPLH Target - LIKE
Sub sw47_Hit:SwitchHitEvent 47: End Sub    ' TNPLH Target - HOME
Sub sw52_Hit:SwitchHitEvent 52: End Sub    ' Bumper Entry Rubber - Lights Glinda Star
Sub sw53_Hit:SwitchHitEvent 53: End Sub    ' Bumper Exit Lane
Sub sw55_Hit:SwitchHitEvent 55: End Sub    ' Skill Target
Sub sw56_Hit:SwitchHitEvent 56: End Sub    ' TIN MAN Rollover
Sub sw61_Hit:SwitchHitEvent 61: End Sub	   ' Winged Monkey Target
Sub sw62_Hit:SwitchHitEvent 62: End Sub	   ' Winged Monkey Target
Sub sw64_Hit:SwitchHitEvent 64: End Sub    ' Left Orbit Made
Sub sw65_Hit:SwitchHitEvent 65: End Sub    ' RESCUE Target - S
Sub sw66_Hit:SwitchHitEvent 66: End Sub    ' RESCUE Target - E
Sub sw67_Hit:SwitchHitEvent 67: End Sub    ' RESCUE Target - R
Sub sw68_Hit:SwitchHitEvent 68: End Sub    ' RESCUE Target - C
Sub sw69_Hit:SwitchHitEvent 69: End Sub    ' RESCUE Target - U
Sub sw70_Hit:SwitchHitEvent 70: End Sub    ' RESCUE Target - E
Sub sw71_Hit:SwitchHitEvent 71: End Sub    ' Castle Loop
Sub sw73_Hit:SwitchHitEvent 73: End Sub    ' Winkie Soldier Drop Target
Sub sw74_Hit:SwitchHitEvent 74: End Sub    ' Glinda Target

Sub sw76_Hit:SwitchHitEvent 76: End Sub    ' Emerald Ramp Made
Sub sw77_Hit:SwitchHitEvent 77: End Sub    ' Castle Door Bash
Sub sw81_Hit:SwitchHitEvent 81: End Sub    ' Left OZ lane
Sub sw82_Hit:SwitchHitEvent 82: End Sub    ' Right OZ Lane
Sub sw87_Hit:SwitchHitEvent 87: End Sub    ' Right Orbit Made
Sub sw89_Hit:SwitchHitEvent 89: End Sub    ' RAINBOW Target - R
Sub sw90_Hit:SwitchHitEvent 90: End Sub    ' RAINBOW Target - A
Sub sw91_Hit:SwitchHitEvent 91: End Sub    ' RAINBOW Target - I
Sub sw92_Hit:SwitchHitEvent 92: End Sub    ' RAINBOW Target - N
Sub sw93_Hit:SwitchHitEvent 93: End Sub    ' RAINBOW Target - B
Sub sw94_Hit:SwitchHitEvent 94: End Sub    ' RAINBOW Target - O
Sub sw95_Hit:SwitchHitEvent 95: End Sub    ' RAINBOW Target - W
Sub sw96_Hit:SwitchHitEvent 96: End Sub    ' SCARECROW Rollover
Sub sw98_Hit:SwitchHitEvent 98: End Sub    ' Horse of Different Color Collect
Sub sw101_Hit:SwitchHitEvent 101: End Sub  ' Munchkin PF Loop lower
Sub sw102_Hit:SwitchHitEvent 102: End Sub  ' Munchkin PF Loop upper
Sub sw103_Hit:SwitchHitEvent 103: End Sub  ' Throne Room Rubber
Sub sw104_Hit:SwitchHitEvent 104: End Sub  ' LION Rollover
Sub sw125_Hit:SwitchHitEvent 125: End Sub  ' Shooter Lane - Cleared Collect Flipper Return
Sub sw126_Hit:SwitchHitEvent 126: End Sub  ' Behind Right Middle Flipper - Collect Lane

' Emerald Ramp Enter
' Timer and flag need to avoid disabling table on other timers expiring resulting
' in ball getting stuck in Munchkin House
Dim bSW75_RampEnter
bSW75_RampEnter = 0

Sub sw75_Hit
    bSW75_RampEnter=1
    sw75Hit.enabled = 1
    SwitchHitEvent 75
End Sub

Sub sw75Hit_Timer
    Me.enabled = 0
    bSW75_RampEnter = 0
End Sub


'***********************************************
'                  KICKERS
'***********************************************
' Munchkin House Ball Handler
' After several iterations (e.g. vlock, catching balls, kicking balls) this approach
' seems most reliable. It is critical to consolidate this handling in one place.
Sub MHouse_BHandler
        If bECLock=0 OR bTOTOMode=1 OR bTNPLHMode=1 OR bBTWW=1 OR bDingDong=1 Then
            Select Case bECBallsLocked
              case 0
                sw6.createBall
                sw6.Kick 180,5
              case Else
                sw6.Kick 180,5
                sw6.createBall
            End Select
        Else
            BallsInLock=BallsInLock+1
            bECBallsLocked=bECBallsLocked+1
            bECLock=bECLock-1
            If PUPpack=False Then ClearBG_EMC1
            Select Case bECBallsLocked
              case 1
                If bMultiBallMode=0 Then Display_MX "LOCK1"
                If PUPpack=True AND bMultiBallMode=0 Then
                    DOF 72, DOFOn
                    vpmtimer.addtimer 3000, "DOF 72, DOFOff '"
                Else
                    Controller.B2SSetData "EMC_Ball1",1
                    vpmtimer.addtimer 3000, "RestoreBG_EMC1 '"
                End If
                PlaySound "voc30_silencewsnap"
                sw6.createBall
                LightState(150) = 1:Flashers(flash_dict.Item(l150))=1
                PlaySound "fx_woodhit"
              case 2
                If bMultiBallMode=0 Then Display_MX "LOCK2"
                If PUPpack=True AND bMultiBallMode=0 Then
                    DOF 75, DOFOn
                    vpmtimer.addtimer 3000, "DOF 75, DOFOff '"
                Else
                    Controller.B2SSetData "EMC_Ball2",1
                    vpmtimer.addtimer 3000, "RestoreBG_EMC1 '"
                End If
                PlaySound "voc44_ComeBackTomorrow"
                vpmtimer.addtimer 3000, "RestoreBG_EMC1 '"
                sw7.createBall
                LightState(151) = 1:Flashers(flash_dict.Item(l151))=1
                PlaySound "fx_collide"
              case 3
                If bMultiBallMode=0 Then Display_MX "LOCK3"
                ClearBG_EMC1
                If PUPpack=True AND bMultiBallMode=0 Then
                    DOF 69, DOFOn
                    vpmTimer.AddTimer 3000,"DOF 69, DOFOff '"
                End If
                Controller.B2SSetData "EMC_MB1",1
                vpmtimer.addtimer 5000, "EmeraldCityMB '"
                PlaySound "voc21_paynoattention"
                sw8.createBall
                LightState(152) = 1:Flashers(flash_dict.Item(l152))=1
                PlaySound "fx_collide"
                SetLightColor L137,"green",2:
                Exit Sub
            End Select
            bBallJustLocked = bBallJustLocked + 1
            vpmtimer.addtimer 200, "CreateNewBall '"
            If bECLock=0 Then
                SetLightsOFF 137
                If bTwisterFlag=1 Then
                    DivRight.RotateToEnd
                    SetLightColor L134,"yellow",2:Controller.B2SSetData 233,1
                End If
            End If
        End If
End Sub

' Sets up for locking a ball in the Munchkin House
Sub ELockProcess
    SetLights 137
    MunchkinPost True
    bECLock=bECLock+1
    SetLightsOFF 134:Controller.B2SSetData 233,0:DivRight.RotateToStart
End Sub

' Munchkin House - Entrance
' Staging kicker - catches ball and destroys
Sub sw5_Hit
    vpmtimer.addtimer 300, "SwitchHitEvent 5 '"
	Me.Timerenabled = 1
End Sub

Sub sw5_Timer
	Me.Timerenabled = 0
    sw5.DestroyBall
End Sub

' Munchkin House
Sub sw6_Hit
    PlaySoundAtVol "fx_woodhit", sw6, 1
    LightState(150) = 1:Flashers(flash_dict.Item(l150))=1
    If bECLock>0 Then sw7.enabled=True
End Sub

Sub sw7_Hit
    PlaySoundAtVol "fx_collide", sw7, 1
    LightState(151) = 1:Flashers(flash_dict.Item(l151))=1
    If bECLock>0 Then sw8.enabled=True
End Sub

Sub sw8_Hit
    PlaySoundAtVol "fx_collide", sw8, 1
    LightState(152) = 1:Flashers(flash_dict.Item(l152))=1
End Sub

' Scoop - Right Middle (Throne Room)
Sub sw15_Hit
    SwitchHitEvent 15
    If bTOTODone=1 AND PUPpack=True Then Me.TimerInterval = 8000
	Me.Timerenabled = 1
    If bTNPLHMode=1 OR bTOTOMode=1 OR bBTWW=1 OR bDingDong=1 OR bSOTR=1Then Exit Sub
    Display_MX "WIZARD"
End Sub

Sub sw15_Timer
    Me.TimerInterval = 1500
	Me.Timerenabled = 0
    PlaySoundAtVol SoundFXDOF("fx_kicker",166,DOFPulse,DOFContactors),sw15, VolKick
	sw15.Kick 225, 20
End Sub

' VUK - Top Center (Soldier) to Castle Playfield
Sub sw17_Hit
    SwitchHitEvent 17
	Me.Timerenabled = 1
End Sub

Sub sw17_Timer
	Me.Timerenabled = 0
	sw17.DestroyBall
	sw17out.CreateBall
    PlaySoundAtVol SoundFXDOF("fx_kicker",167,DOFPulse,DOFContactors), sw17, VolKick
	sw17out.Kick 180, 10
End Sub

' Scoop - Left Top (Castle)
Sub sw18_Hit
    SwitchHitEvent 18
	Me.Timerenabled = 1
End Sub

Sub sw18_Timer
	Me.Timerenabled = 0
    PlaySoundAtVol SoundFXDOF("fx_kicker",168,DOFPulse,DOFContactors), sw18, VolKick
	sw18.Kick 135, 10
End Sub

' VUK - Left Middle (Owl)
Sub sw21_Hit
    SwitchHitEvent 21
    If bTOTODone=1 Then Me.TimerInterval = 7000
	Me.Timerenabled = 1
End Sub

Sub sw21_Timer
    Me.TimerInterval = 1000
	Me.Timerenabled = 0
	sw21.DestroyBall
	sw21out.CreateBall
    PlaySoundAtVol SoundFXDOF("fx_kicker",168,DOFPulse,DOFContactors), sw21, 1
	sw21out.Kick 0, 20
End Sub

' Magnet Grab Simulator - Monkey Magnet Top
Sub Kicker_Mag_Hit
    PlaysoundAtVol "fx_kicker_enter", Kicker_Mag, 1
    MonkeyMoverStart.enabled = 1
End Sub

' Fall Through - Munchkin Playfield
Sub Kicker_Munchkin_Hit
    PlaySoundAtVol "fx_ballrampdrop", Kicker_Munchkin, 1
    If bTNPLHMode=1 OR bTOTOMode=1 OR bBTWW=1 OR bSOTR=1 Then Exit Sub
    ballsOnMunchkinPF = ballsOnMunchkinPF - 1
    If ballsOnMunchkinPF<=0 Then
        LightSeqTwister.StopPlay
        vpmtimer.AddTimer 1000, "TNPLH_TOTO_LightsReset '"
        StopSound "voc20_Twister"
        If bMultiBallMode=0 Then GiOn
        bOnMunchkinPF=0
        If bCastlePFActive=1 Then
            Multipliers(4) = 2
        Else
            Multipliers(4) = 1
        End If
        SuperXcolor Multipliers(4)*Multipliers(6)
        munchkinMB_Rcvd=0
        If bMunchkinMode=1 Then Kicker_Mag1.enabled=True: Exit Sub
        If bMunchkinMode=0 AND bBTWW=0 Then
            ' enable the magnets for a few seconds
            EnableRMagnet
            TwisterModeHold.interval = 30000    ' Twister Mode Hold Interval when Munchkim Mode Not Made
            TwisterModeHold.enabled = True
'            Twister_ClearCounts
            vpmtimer.AddTimer 1000, "PlaySong ""voc18_OffToSee"" '"
            vpmtimer.AddTimer 2000, "DisableMagnets '"
        End If
    End If
End Sub

' Magnet Grab Simulator - Below Munchkin Hole
Sub Kicker_Mag1_Hit
    Me.Timerenabled = 0
    If PUPpack=True AND bMultiBallMode=0 Then
        Select Case munchkinModeCnt
          case 0
            Me.TimerInterval = 10000
          case 1
            Me.TimerInterval = 8000
          case 2
            Me.TimerInterval = 6000
          case Else
            Me.TimerInterval = 2000
        End Select
    Else
        Select Case munchkinModeCnt
          case 0
            Me.TimerInterval = 5000
          case 1
            Me.TimerInterval = 4000
          case 2
            Me.TimerInterval = 3000
          case Else
            Me.TimerInterval = 2000
        End Select
    End If
 	Me.Timerenabled = 1
    MunchkinMode
End Sub

Sub Kicker_Mag1_Timer
	Me.Timerenabled = 0
    Me.TimerInterval = 2000
    PlaySoundAtVol "fx_kicker", Kicker_Mag1, VolKick
	Kicker_Mag1.Kick 180, 8
    Kicker_Mag1.enabled=False
    vpmtimer.AddTimer 1000, "PlaySong ""voc18_OffToSee"" '"
End Sub

' TOTO Mode (1st Round Only) Ball Hold On Complete
' Set Timer interval to TOTO Back video length
Dim hold_time, bTOTO_kicker
hold_time = 8000
bTOTO_kicker=0

Sub Kicker_TOTO_Hit
    Kicker_TOTO.TimerInterval = hold_time
    bTOTO_kicker=1
	Me.Timerenabled = 1
    If bTOTOMode=0 AND bTNPLHMode=0 AND SOTR_capture=1 Then
        SOTR_capture=0
        SOTR_Mode
    End If
End Sub

Sub Kicker_TOTO_Timer
	Me.Timerenabled = 0
	Kicker_TOTO.Kick 135, 5
    Kicker_TOTO.enabled=False
'    debug.print "Kicker.TOTO OFF - Kicker_TOTO_Timer"
    hold_time = 8000
    bTOTO_kicker=0
End Sub


'***********************************************
'               MAGNET TRIGGERS
'***********************************************
' Witch Magnets
Sub sw57_Hit
    SwitchHitEvent 57
    EnableLMagnet
    DOF 127,DOFPulse
    vpmtimer.AddTimer 3000, "DisableMagnets '"
End Sub

Sub sw58_Hit
    SwitchHitEvent 58
    EnableLMagnet
    DOF 127,DOFPulse
    vpmtimer.AddTimer 3000, "DisableMagnets '"
End Sub

Sub DisableMagnets
    LMagnet.MotorOn = False
    RMagnet.MotorOn = False
End Sub

Sub EnableLMagnet
    LMagnet.MotorOn = True
End Sub

Sub EnableRMagnet
    RMagnet.MotorOn = True
End Sub

'***********************************************
'               MONKEY MOVERS
'***********************************************
Dim x_pos, y_pos, z_pos, delta_X, delta_Y, delta_Z, size_delta_X, size_delta_Y
Dim x_pos_B, y_pos_B, z_pos_B
Dim x_step, y_step, z_step, x_size_step, y_size_step
Dim start_move_cnt, x_size, y_size
Sub MonkeyMoverStart_Timer
     Me.enabled = 0
     FlyingMonkey_Ball.visible = 0
     FlyingMonkey_NB.visible = 0

     FlyingMonkey_NB.X = 358
     FlyingMonkey_NB.Y = 85
     FlyingMonkey_NB.Z = 180
     FlyingMonkey_NB.Size_X = 130
     FlyingMonkey_NB.Size_Y = 150
     FlyingMonkey_NB.visible = 1

     FlyingMonkey_Ball.X = 741
     FlyingMonkey_Ball.Y = 105
     FlyingMonkey_Ball.Z = 100
     FlyingMonkey_Ball.Size_X = 165
     FlyingMonkey_Ball.Size_Y = 230


     ' Position of Flying Monkey without Ball (Start)
     x_pos = 358: y_pos = 85: z_pos = 180
     ' Position of Flying Monkey with Ball (Where Ball Captured)
     x_pos_B = 741: y_pos_B = 105: z_pos_B = 100
     ' Distances to travel
     delta_X = x_pos_B - x_pos
     delta_Y = y_pos_B - y_pos
     delta_Z = z_pos_B - z_pos
     x_size = 130
     y_size = 150
     size_delta_X = 35
     size_delta_Y = 80

     x_step = delta_X/142
     y_step = delta_Y/142
     z_step = delta_Z/142
     x_size_step = size_delta_X/142
     y_size_step = size_delta_Y/142

     start_move_cnt = delta_X + x_step
     MoveMonkeyNB.enabled = 1
     DOF 352, DOFOn   ' Turn Gear On
     If PUPpack=False Then PlaySound "voc6_nowfly"
     If PUPpack=True AND bMultiBallMode=0 Then
         DOF 89, DOFPulse
     Else
         If PUPpack=True AND bMultiBallMode=1 AND bWitchHU=0 Then
             DOF 79, DOFPulse
         End If
     End If
End Sub


Sub MoveMonkeyNB_Timer
     start_move_cnt = start_move_cnt - x_step

     x_pos = x_pos + x_step
     y_pos = y_pos + y_step
     z_pos = z_pos + z_step
     x_size = x_size + x_size_step
     y_size = y_size + y_size_step

     If start_move_cnt > 0 Then
         FlyingMonkey_NB.X = x_pos
         FlyingMonkey_NB.Y = y_pos
         FlyingMonkey_NB.Z = z_pos
         FlyingMonkey_NB.Size_X = x_size
         FlyingMonkey_NB.Size_Y = y_size
    Else
         Me.enabled = 0
         FlyingMonkey_NB.visible = 0
         FlyingMonkey_Ball.visible = 1
	     Kicker_Mag.DestroyBall
         Kicker_Mag.enabled = False
         x_pos = 741: y_pos = 105: z_pos = 100

         delta_Y = 20
         y_step = delta_Y/142
         x_step = -x_step
         y_step = -y_step
         z_step = -z_step
         x_size = 165
         y_size = 230
         x_size_step = -x_size_step
         y_size_step = -y_size_step
         start_move_cnt =  -(delta_X + x_step)
         MoveMonkeyBall.Interval = 1000
         MoveMonkeyBall.enabled = 1
         sw72_out.enabled = True
    End If
End Sub


Sub MoveMonkeyBall_Timer
     Me.Interval = 10
     start_move_cnt = start_move_cnt - x_step

     x_pos = x_pos + x_step
     y_pos = y_pos + y_step
     z_pos = z_pos + z_step
     x_size = x_size + x_size_step
     y_size = y_size + y_size_step

     If start_move_cnt < 0 Then
         FlyingMonkey_Ball.X = x_pos
         FlyingMonkey_Ball.Y = y_pos
         FlyingMonkey_Ball.Z = z_pos
         FlyingMonkey_Ball.Size_X = x_size
         FlyingMonkey_Ball.Size_Y = y_size
    Else
         Me.enabled = 0
         DOF 352, DOFOff    ' Turn Gear Off
         FlyingMonkey_NB.visible = 0
         FlyingMonkey_Ball.visible = 0
	     sw72_out.CreateBall
         SetSwitchFlags 72
         BallsInLock=BallsInLock+1
         bBallJustLocked = bBallJustLocked + 1
'         bBallJustLocked=1
         CreateNewBall
    End If
End Sub

Sub sw72_out_Timer
	Me.Timerenabled = 0
    PlaySoundAtVol "fx_kicker", sw72, VolKick
	sw72_out.Kick 180, 5
    sw72_out.enabled = False
End Sub


'***********************************************
'      DOORS & DIVERTERS (FLIPPERs)
'***********************************************
Sub DiverterRight(enabled)
    If enabled Then
        DivRight.RotatetoEnd
    Else
        DivRight.RotateToStart
    End If
End Sub

'************************************************
'              POSTS
'************************************************
' lock post - auto drop
Sub lock1_Hit:lock.Isdropped = 1:PlaySoundAtVol "metalhit2",lock1,VolMetal:End Sub
Sub lock1_UnHit:lock.Isdropped = 0:End Sub

Sub MunchkinPost(enabled)
	If Enabled Then
        If lock.IsDropped Then
		    Lock.IsDropped = 0
		    PlaySound SoundFX("fx_PostUp",DOFContactors) ' TODO
        End If
	Else
		Lock.IsDropped = 0
		PlaySound SoundFX("fx_PostDown",0) ' TODO
	End If
End Sub

'************************************************
'              CASTLE DOORS
'************************************************
Sub CastleDoorsOpen
    CastleFlipper1.RotateToEnd
    CastleFlipper2.RotateToEnd
    CastleDoor_Left.RotY = CastleFlipper2.CurrentAngle + 5
    CastleDoor_Right.RotY = CastleFlipper1.CurrentAngle + 46
    Wall108.IsDropped = 1
End Sub

Sub CastleDoorsClose
    CastleFlipper1.RotateToStart
    CastleFlipper2.RotateToStart
    CastleDoor_Left.RotY = CastleFlipper2.CurrentAngle + 15
    CastleDoor_Right.RotY = CastleFlipper1.CurrentAngle + 40
    Wall108.IsDropped = 0
End Sub

'************************************************
'          CRYSTAL BALL Spinner
'************************************************
Sub Spinner_Spin
    PlaySoundAtVol "fx_spinner", Spinner, VolSpin
    AddScore 2   ' Treat like Switch Number 2
    SetDOF 2     ' Treat like Switch Number 2
End Sub

'************************************************
'      CRYSTAL BALL - PASSIVE AUTONOMOUS DISPLAY
' Crystal Ball Hologram animation
' Code and associated playfield items copied
' or adapted from JPSalas' Slimer
'************************************************
Dim HoloDir, HoloSize, HoloStep, HoloCnt
Dim HoloSizeX, HoloSizeY, HoloImage, HoloPreShrink
Dim tFlag
Sub HologramStart
    Hologram.visible = 0
    Primitive6.visible = 1

    HoloDir = 0
    HoloCnt = 0
    HoloSize = 20
    HoloPreShrink = 80
    HoloStep = 0
    Hologram.TransX = 0
    Hologram.TransY = 0
    Hologram.Rotz = 0
    Hologram.Roty = 0
    HoloSizeX = Hologram.Size_X
    HoloSizeY = Hologram.Size_Y
    tFlag = False


    ' Load Fog and shrink
    Hologram.Image = "cball0"
    Hologram.Size_X = HoloGram.Size_X + 20
    Hologram.Size_Y = HoloGram.Size_Y + 20

'    HologramWait.Enabled = 0
    HologramTimer.Enabled = 1

    ' Queue up first image
    HoloImage = "cball" &INT(RND * cBallImages + 1)

End Sub

Sub HologramTimer_Timer
    Select Case HoloStep
        Case 0 'zoom
            Hologram.visible = 1
            HoloCnt = HoloCnt + 1
            HoloDir = 1
            If Hologram.Image = "cball0" Then
                Hologram.Rotz = Hologram.Rotz + 0.25
            Else
                Hologram.Size_X = HoloGram.Size_X + HoloDir
                Hologram.Size_Y = HoloGram.Size_Y + HoloDir
            End If
            If HoloCnt > (HoloPreShrink + 20) Then
               HoloCnt = 0
               If Hologram.Image = "cball0" Then
                   HoloStep = 3
               Else
                   HoloStep = 2
               End If
            End If
        Case 1 'shrink
            HoloCnt = HoloCnt + 0.75
            HoloDir = -0.75
            If Hologram.Image = "cball0" Then
                Hologram.Rotz = Hologram.Rotz + 0.25
            Else
                Hologram.Size_X = HoloGram.Size_X + HoloDir
                Hologram.Size_Y = HoloGram.Size_Y + HoloDir
            End If
            If HoloCnt > HoloPreShrink  Then
               HoloCnt = 0
               Hologram.visible = 0
               Hologram.Image = HoloImage
               If Hologram.Image = "cball0" Then
                   Hologram.Size_X = HoloSizeX + 20
                   Hologram.Size_Y =  HoloSizeY + 20
               Else
                   Hologram.Size_X = HoloSizeX
                   Hologram.Size_Y =  HoloSizeY
                   Hologram.Rotz = 0
               End If
               If Hologram.Image <> "cball0" Then
                   Hologram.Size_X = HoloGram.Size_X - HoloPreShrink
                   Hologram.Size_Y = HoloGram.Size_Y - HoloPreShrink
               End If
               If HoloImage = "cball0" Then
                   HoloImage = "cball" &INT(RND * cBallImages + 1)
               Else
                   HoloImage = "cball0"
               End If
               HoloStep = 0
            End If
        Case 2 'display of full image hold time
            If tFlag = False Then
                HologramTimerExpired.Interval = 7000
                HologramTimerExpired.Enabled = 1
                HologramWait.Interval = 8000
                HologramWait.Enabled = 1
                tflag = True
                Primitive6B.visible = 1
                Primitive6.visible = 0
            End If
        Case 3 'display of full fog image hold time and slot for fog motion
            If tFlag = False Then
                HologramTimerExpired.Interval = 1500
                HologramTimerExpired.Enabled = 1
                tflag = True
            Else
                If Hologram.Image = "cball0" Then Hologram.Rotz = Hologram.Rotz + 0.25
            End If
    End Select
End Sub

Sub HologramTimerExpired_Timer
    HologramTimer.Enabled = 0
    Me.enabled = 0
    HoloStep = 1
'    Primitive6.visible = 1
'    Primitive6B.visible = 0
    tFlag = False
    HologramTimer.Enabled = 1
End Sub

Sub HologramWait_Timer
    Me.enabled = 0
    Primitive6.visible = 1
    Primitive6B.visible = 0
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************
'************************************
' What you need to add to your table
'************************************
' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc

'******************************************
' Explanation of the rolling sound routine
'******************************************
' sounds are played based on the ball speed and position
' the routine checks first for deleted balls and stops the rolling sound.
' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.
' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.

'**************************************
' Explanation of the collision routine
'**************************************
' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

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

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

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
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
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
'      JP's VP10 Rolling Sounds
'*****************************************
Const tnob = 20 ' total number of balls
Const lob = 0   'number of locked balls on table at game start

ReDim rolling(tnob)

InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer
    Dim BOT, b, ballpitch
    BOT = GetBalls

	' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

	' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

	' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) ) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) * 50
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'**********************
' Ramp Sounds
'**********************
Sub RWireStart_Hit
    PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub LWireStart_Hit
    PlaySound "fx_metalrolling_long", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub LWireStart2_Hit
    PlaySound "fx_metalrolling_short", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub RHelp1_Hit
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub RHelp2_Hit
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub RHelp3_Hit
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************
Sub aCaptiveWalls_Hit(idx):PlaySound "fx_collide", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub a_Metals_Medium_Hit (idx):PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
Sub a_Metals_Thin_Hit (idx):PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
Sub a_Metals2_Hit (idx):PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
Sub a_Pins_Hit (idx):PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub a_Spinner_Spin:PlaySound "fx_spinner",0,.25,0,0.25:End Sub
Sub a_Targets_Hit (idx):PlaySound "fx_target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

Sub a_Posts_Hit(idx):
    dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber
 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber
 	End If
End sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber
 	End If
End Sub

Sub RandomSoundRubber
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'********************************************
' Random quotes or music clips from the movie
'********************************************
Dim Quote
Sub PlayQuote_timer
    ' Check for Game in Play ... Synchronization Issue
    If bGameInPlay = True Then
        Me.Enabled = False
        Exit Sub
    End If
    Quote = "quote_" & INT(RND *6) + 1
    PlaySound Quote
    PlayQuote.Interval = 35000 'one quote or music clip every 45 seconds
End Sub

'*****************************
'     Music as wav sounds
'*****************************
Dim Song
Song = ""

Sub PlaySong(name)
    If bSOTR_init=1 Then Exit Sub
    If bMusicOn Then
        If Song <> name Then
            StopSound Song
            Song = name
            Select Case Song
              case "voc13_winkiemarch", "voc20_Twister"
                PlaySound Song, -1, 1 'this last number is the volume, from 0 to 1
              case "voc18_OffToSee", "voc25_Hahaha"
                PlaySound Song, 1, 0.1 'this last number is the volume, from 0 to 1
              case "voc8_Rainbow", "voc78_Marvel", "voc79_ToShelter", "voc80_ToHouse"
                PlaySound Song, 1, 0.1 'this last number is the volume, from 0 to 1
              case Else
                PlaySound Song, -1, 0.1 'this last number is the volume, from 0 to 1
            End Select
        End If
    End If
End Sub

'******************************************************
'                       GI effects
'******************************************************
' Change GI Color
Sub ChangeGi(col)
    Dim bulb
    Dim i
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
    ' PF Left & Right Sides MX Effect - Match to mode
    For i = 310 to 317
        DOF i, DOFOff
    Next
    Select Case col
        Case "amber"
            DOF 310, DOFOn
        Case "blue"
            DOF 311, DOFOn
        Case "green"
            DOF 312, DOFOn
        Case "orange"
            DOF 313, DOFOn
        Case "pink"
            DOF 314, DOFOn
        Case "purple"
            DOF 315, DOFOn
        Case "red"
            DOF 316, DOFOn
        Case "yellow"
            DOF 317, DOFOn
        Case Else
            For i = 310 to 317
                DOF i, DOFOff
            Next
    End Select
End Sub

Sub GiOn
    Dim bulb, obj, get_color
    For each bulb in aGiLights
        bulb.State = 1
    Next
    ' SCOOP lights
    For each obj in aScoopLights
        get_color = lcolor_dict.Item(obj)
        SetLightColor obj,get_color,1
    Next
End Sub

Sub GiOff
    Dim bulb, obj, get_color
    Dim i
    For each bulb in aGiLights
        bulb.State = 0
    Next
    ' SCOOP lights
    For each obj in aScoopLights
        get_color = lcolor_dict.Item(obj)
        SetLightColor obj,get_color,0
    Next
    ' PF Left & Right Sides MX Effect - Turn Off
    For i = 310 to 317
        DOF i, DOFOff
    Next
End Sub

' GI light sequence effects
Sub GiEffect(n, rpt)
    Select Case n
        Case 0 'all off
            LightSeqGI.Play SeqAlloff
        Case 1 'all blink
            LightSeqGI.UpdateInterval = 10
            LightSeqGI.Play SeqBlinking, , rpt, 100
        Case 2 'random
            LightSeqGI.UpdateInterval = 15
            LightSeqGI.Play SeqRandom, 5, , 1000
        Case 3 'upon
            LightSeqGI.UpdateInterval = 10
            LightSeqGI.Play SeqUpOn, 5, 1
        Case 4 ' left-right-left
            LightSeqGI.UpdateInterval = 10
            LightSeqGI.Play SeqLeftOn, 10, 1
            LightSeqGI.UpdateInterval = 10
            LightSeqGI.Play SeqRightOn, 10, 1
    End Select
End Sub

'******************************************************
'                 INSERT LIGHT effects
'******************************************************
' Insert lights are all non-passive PF Lights
Sub LightEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqInserts.Play SeqAlloff
        Case 1 'all blink
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 5, 100
        Case 2 'random
            LightSeqInserts.UpdateInterval = 15
            LightSeqInserts.Play SeqRandom, 5, , 1000
        Case 3 'upon
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqUpOn, 10, 1
        Case 4 ' left-right-left
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqLeftOn, 10, 1
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqRightOn, 10, 1
        Case 5  ' all On
            LightSeqInserts.Play SeqAllon
    End Select
End Sub

'******************************************************
'                 FLASHER effects
'******************************************************
Sub FlashEffect(n)
    Select case n
        Case 0 ' all off
            LightSeqFlasher.Play SeqAlloff
        Case 1 'all blink
            LightSeqFlasher.UpdateInterval = 15
        Case 2 'random
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqRandom, 50, , 1000
        Case 3 'upon
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqUpOn, 10, 1
        Case 4 ' left-right-left
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqLeftOn, 10, 1
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqRightOn, 10, 1
    End Select
End Sub

'******************************************************
'        JP's VP10 Fading Lamps & Flashers
'  very reduced, mostly for rom activated flashers
' if you need to turn a light on or off then use:
'	LightState(lightnumber) = 0 or 1
'        Based on PD's Fading Light System
'******************************************************
Dim LightState(200), FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitFlashers ' turn off the lights and flashers and reset them to the default parameters

LampTimer.Interval = 200 'lamp fading speed
LampTimer.Enabled = 1

Sub LampTimer_timer
    Dim chgLamp, x
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For x = 0 To UBound(chgLamp)
            LightState(chgLamp(x, 0) ) = chgLamp(x, 1) 'light state as set by the rom
        Next
    End If

' Lights & Flashers
'    LightX 1, l1     ' Example

    Flash 104, l104   ' Oz Lane Left
    Flash 105, l105   ' Oz Lane Right
    Flash 115, l115   ' Capture Dorothy Arrow
    Flash 150, l150   ' Munchkin Ramp Lock #1
    Flash 151, l151   ' Munchkin Ramp Lock #2
    Flash 152, l152   ' Munchkin Ramp Lock #3
    Flash 153, l153   ' Dorothy Captured Arrow

End Sub

Sub SetLamp(nr, value)
    If value <> LightState(nr) Then
        LightState(nr) = value
    End If
End Sub

' lamp subs
Sub InitFlashers
    Dim x
    For x = 0 to 200
        LightState(x) = 0     	 ' light state: 0=off, 1=on, -1=no change (on or off)
        FlashSpeedUp(x) = 0.5    ' Fade Speed Up
        FlashSpeedDown(x) = 0.25 ' Fade Speed Down
        FlashMax(x) = 1          ' the maximum intensity when on, usually 1
        FlashMin(x) = 0          ' the minimum intensity when off, usually 0
        FlashLevel(x) = 0        ' the intensity/fading of the flashers
    Next
End Sub

' VPX Lights, just turn them on or off
Sub LightX(nr, object)
    Select Case LightState(nr)
        Case 0, 1:object.state = LightState(nr):LightState(nr) = -1
    End Select
End Sub

Sub LightXm(nr, object) 'multiple lights
    Select Case LightState(nr)
        Case 0, 1:object.state = LightState(nr)
    End Select
End Sub

' VPX Flashers, changes the intensity
Sub Flash(nr, object)
    Select Case LightState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                LightState(nr) = -1 'completely off, so stop the fading loop
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                LightState(nr) = -1 'completely on, so stop the fading loop
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the intensity
    Object.IntensityScale = FlashLevel(nr)
End Sub

'*************************
' Light Sequencing
' Code adapted
' from JPSalas' Slimer
'*************************
Dim sequence, collection
Const SequenceDuration = 3600 ' About 2 hours (120min*60sec/min / 2sec/sequence)

Sub StartAttractMode
    DOF 323, DOFOn   'DOF MX - Attract Mode ON
    If PUPpack=True Then
        PUP_Attract.Interval = 130000  ' Wait ~2 minutes - need 30sec for backglass load to finish
        PUP_Attract.Enabled = 1
    End If
    PlayQuote.Interval = 33000
    PlayQuote.Enabled = 1
    bAttractMode = True
    StartLightSeq
    If DMD_Option Then DMDFlush
    ShowTableInfo: ShowTableInfoTimer.enabled = True
    StartRainbow "aLEDLights"        ' Change Playfield LED light colors
    ChangeGi "darkgreen":GiOn
End Sub

Sub PUP_Attract_Timer
    ClearBackGlass
    PlayQuote.Enabled = 0
    StopSound Quote
    PUP_Attract.Enabled = 0
    DOF 100, DOFPulse
    PUP_Attract.Interval = 240000  ' Repeat in 4 minutes
    PUP_Attract.Enabled = 1
    ' Need to wait for video to finish but can cause synchronization issue
    vpmtimer.addtimer 120000, "PlayQuote.Enabled = 1 '"
End Sub

Sub StopAttractMode
    DOF 323, DOFOff   'DOF MX - Attract Mode OFF
    If PUPpack=True Then
        PUP_Attract.Enabled = 0
        DOF 301,DOFpulse
    End If
    PlayQuote.Enabled = 0
    StopSound Quote
    bAttractMode = False
    If DMD_Option=True Then DMDFlush:DMD "", eNone, "", eNone, "", eNone, "", eNone, 500, True, ""
    If UltraDMD_Option=True Then UltraDMDScoreNow
    LightSeqAttract.StopPlay
    LightSeqFlasher.StopPlay
    StopRainbow
    ShowTableInfoTimer.enabled = False
    ChangeGi "white"

End Sub

Sub ShowTableInfoTimer_Timer
    ShowTableInfo
End Sub


Sub StartSequence(collectionChoice, sequenceChoice)
    collection = collectionChoice
    sequence=sequenceChoice
     On Error Resume Next
    If collection = "aLEDLights"  Then
        Select Case sequence
            Case 0  'Sequence Down
                LightSeq_choice.UpdateInterval = 8
                LightSeq_choice.Play SeqDownOn, 25, SequenceDuration
            Case 1  'Sequence Up
                LightSeq_choice.UpdateInterval = 12
                LightSeq_choice.Play SeqUpOn, 75, SequenceDuration
        End Select
     End if
     If collection = "aGILights"  Then
         Select Case sequence
             Case 0  'Sequence Down
                 LightSeq_choice.UpdateInterval = 8
                 LightSeq_choice.Play SeqDownOn, 25, SequenceDuration
             Case 1  'Sequence Up
                 LightSeq_choice.UpdateInterval = 12
                 LightSeq_choice.Play SeqUpOn, 75, SequenceDuration
         End Select
     End if
End Sub

Sub StopSequence
    StopRainbow
    If collection = "aLEDLights" Then
        LightSeq_choice.StopPlay
    End If
    If collection = "aGILights" Then
        LightSeq_choice.StopPlay
    End If
    InitFlashers
End Sub


Sub StartTiltSequence
    Dim obj
    StopRainbow
    If collection = "aLEDLights" Then
        LightSeq_choice.StopPlay
        For each obj in aLEDLights
            SetLightColor obj,"red",1
        Next
        LightSeq_choice.Play SeqBlinking,,SequenceDuration,300
    End If
    If collection = "aGILights" Then
        LightSeq_choice.StopPlay
        For each obj in aLEDLights
            SetLightColor obj,"red",1
        Next
        LightSeq_choice.Play SeqBlinking,,SequenceDuration,300
    End If
    InitFlashers
End Sub

Sub SequenceTimer_Timer
End Sub

Sub StartLightSeq
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqRandom,10,,500
    LightSeqAttract.Play SeqDownOn,25,3
    LightSeqAttract.Play SeqBlinking,,5,150
    LightSeqAttract.Play SeqUpOn,50,3
    LightSeqAttract.Play SeqCircleOutOn,15,5
    LightSeqAttract.Play SeqRightOn,50,3
    LightSeqAttract.Play SeqLeftOn,40,4
    LightSeqAttract.Play SeqStripe1HorizOn,50,6
    LightSeqAttract.Play SeqStripe2VertOn,50,3
    LightSeqAttract.Play SeqDiagUpRightOn,25,3
    LightSeqAttract.Play SeqDiagUpRightOff,25,3
    LightSeqAttract.Play SeqDiagUpLeftOn,25,2
    LightSeqAttract.Play SeqDiagUpLeftOff,25,4
    LightSeqAttract.Play SeqDiagDownRightOn,25,3
    LightSeqAttract.Play SeqDiagDownRightOff,25,3
    LightSeqAttract.Play SeqDiagDownLeftOn,25,2
    LightSeqAttract.Play SeqDiagDownLeftOff,25,2
    LightSeqAttract.Play SeqHatch2HorizOn,50,4
    LightSeqAttract.Play SeqHatch1VertOff,50,2
    LightSeqAttract.Play SeqClockRightOn,50,2
    LightSeqAttract.Play SeqRadarLeftOn,50,4
    LightSeqAttract.Play SeqWiperLeftOn,45,4
    LightSeqAttract.Play SeqFanLeftUpOn,45,2
    LightSeqAttract.Play SeqArcBottomRightUpOn,90,3
    LightSeqAttract.Play SeqScrewRightOn,25,3
End Sub


Sub TStartLightSeq
    LightSeqAttract.Play SeqScrewRightOn,5
End Sub


Sub LightSeqAttract_PlayDone
    StartLightSeq
End Sub


'******************************************
' Change light color - simulate color leds
' changes the light color and state
' colors: red, orange, yellow, green, blue, white
'******************************************
Sub SetLightColor(n, col, stat)
    If col = "rainbow" Then SetLightRainbow n, col, stat
    Select Case col
        Case "amber"
            n.color = RGB(193, 49, 0)
            n.colorfull = RGB(255, 153, 0)
        Case "red"
            n.color = RGB(18, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case "orange"
            n.color = RGB(18, 3, 0)
            n.colorfull = RGB(255, 64, 0)
        Case "yellow"
            n.color = RGB(18, 18, 0)
            n.colorfull = RGB(255, 255, 0)
        Case "ylwgrn"
            n.color = RGB(9, 18, 0)
            n.colorfull = RGB(128, 255, 0)
        Case "green"
            n.color = RGB(0, 18, 0)
            n.colorfull = RGB(0, 255, 0)
        Case "darkgreen"
            n.color = RGB(0, 8, 0)
            n.colorfull = RGB(0, 64, 0)
        Case "blue"
            n.color = RGB(0, 18, 18)
            n.colorfull = RGB(0, 255, 255)
        Case "darkblue"
            n.color = RGB(0, 8, 8)
            n.colorfull = RGB(0, 64, 64)
        Case "cyan"
            n.color = RGB(0, 18, 15)
            n.colorfull = RGB(0, 255, 200)
        Case "white"
            n.color = RGB(255, 252, 224)
            n.colorfull = RGB(193, 91, 0)
        Case "purple"
            n.color = RGB(128, 0, 128)
            n.colorfull = RGB(255, 0, 255)
        Case "pink"
            n.color = RGB(255, 171, 142)
            n.colorfull = RGB(255, 227, 215)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub


Sub TurnOffPlayfieldLights
    Dim a
    For each a in aLEDLights
        SetLightColor a,"white",0
    Next
End Sub

'*************************
' Rainbow Changing Lights
' Code copied from JPSalas' Slimer
'*************************
Dim RGBStep, RGBFactor, Red, Green, Blue
Dim RainbowLights
Dim hold
Dim light_obj

' This subroutine "SetLightRainbow" assumes Start Rainbow will not be used when game active
Sub SetLightRainbow(n, col, stat)
    RainbowLights = "none"
    Set light_obj = n
    RGBStep = 0
    RGBFactor = 20
    Red = 255
    Green = 0
    Blue = 0
    RainbowTimer.Enabled = 1
End Sub

Sub StartRainbow(collectionChoice)
    RainbowLights = collectionChoice
    RGBStep = 0
    RGBFactor = 20
    Red = 255
    Green = 0
    Blue = 0
    RainbowTimer.Enabled = 1
End Sub

Sub StopRainbow
    Dim obj
    RainbowTimer.Enabled = 0

    Select Case RainbowLights
        Case "aLEDLights"  ' aLEDLights
            For each obj in aLEDLights
                SetLightColor obj, "white", 0
            Next
        Case "aGILights"  ' aGILights
            For each obj in aGiLights
                SetLightColor obj, "white", 0
            Next
        Case Else
             SetLightColor light_obj, "white", 0
    End Select
End Sub

' Step Transitions
Sub RainbowTimer_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep
        Case 0 'Green
            Green = 255
            hold=hold + RGBFactor
            If hold > 255 then
                hold = 255
                RGBStep = 1
            End If
        Case 1 'Red
            Red = 0
            hold=hold - RGBFactor
            If hold < 0 then
                hold = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            Blue = 255
            hold=hold + RGBFactor
            If hold > 255 then
                hold = 255
                RGBStep = 3
            End If
        Case 3 'Green
            Green = 0
            hold=hold - RGBFactor
            If hold < 0 then
                hold = 0
                RGBStep = 4
            End If
        Case 4 'Red
            Red = 255
            hold=hold + RGBFactor
            If hold > 255 then
                hold = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            Blue = 0
            hold=hold - RGBFactor
            If hold < 0 then
                hold = 0
                RGBStep = 0
            End If
    End Select

    Select Case RainbowLights
        Case "aLEDLights"  ' aLEDLights
            For each obj in aLEDLights
                obj.color = RGB(0, 0, 0)
                obj.colorfull = RGB(Red, Green, Blue)
            Next
        Case "aGILights"  ' aGILights
            For each obj in aGiLights
                obj.color = RGB(0, 0, 0)
                obj.colorfull = RGB(Red, Green, Blue)
            Next
        Case Else
                light_obj.color = RGB(0, 0, 0)
                light_obj.colorfull = RGB(Red, Green, Blue)
    End Select
End Sub

' Fading Transitions
Sub TRainbowTimer 'rainbow led light color changing
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

    Select Case RainbowLights
        Case "aLEDLights"  ' aLEDLights
            For each obj in aLEDLights
                obj.color = RGB(0, 0, 0)
                obj.colorfull = RGB(Red, Green, Blue)
            Next
        Case "aGILights"  ' aGILights
            For each obj in aGiLights
                obj.color = RGB(0, 0, 0)
                obj.colorfull = RGB(Red, Green, Blue)
            Next
    End Select
End Sub


'*****************************
'    Load / Save / Highscore
'*****************************
Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "")Then HighScore(0) = CDbl(x)Else HighScore(0) = 1000 End If
    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "")Then HighScoreName(0) = x Else HighScoreName(0) = "ME" End If
    x = LoadValue(TableName, "HighScore2")
    If(x <> "")then HighScore(1) = CDbl(x)Else HighScore(1) = 1000 End If
    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "")then HighScoreName(1) = x Else HighScoreName(1) = "YOU" End If
    x = LoadValue(TableName, "HighScore3")
    If(x <> "")then HighScore(2) = CDbl(x)Else HighScore(2) = 1000 End If
    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "")then HighScoreName(2) = x Else HighScoreName(2) = "HIM" End If
    x = LoadValue(TableName, "HighScore4")
    If(x <> "")then HighScore(3) = CDbl(x)Else HighScore(3) = 1000 End If
    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "")then HighScoreName(3) = x Else HighScoreName(3) = "HER" End If
    x = LoadValue(TableName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0 End If

    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "")then TotalGamesPlayed = CInt(x)Else TotalGamesPlayed = 0 End If
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

Sub CheckHighscore
    Dim tmp
    tmp = Score(CurrentPlayer)

    If tmp> HighScore(1) Then 'add 1 credit for beating the highscore
        AwardSpecial
    End If

    If tmp> HighScore(3) Then
        hsbModeActive = True
        PlaySound "vo_HiScore"
        BallLossCheck.Enabled=False
        HighScore(3) = tmp
        'enter player's name
        vpmtimer.addtimer 1000, "HighScoreEntryInit '"
        HiScoreEntryTimer.Enabled=True
    End If
End Sub

' If a letter is not selected within 1 minute this will exit the mode
Sub HiScoreEntryTimer_Timer
    Me.Enabled = False
    hsEnteredDigits(0) = "W"
    hsEnteredDigits(1) = "O"
    hsEnteredDigits(2) = "Z"
    HighScoreCommitName
    If DMD_Option=True Then
        DMD2Flush
        If (dqHead2 = dqTail2) Then
            DMD2 "", eNone, "HISCORE INITIALS", eNone, "TIMEOUT EXPIRED", eNone, "", eNone, 5000, True, ""
        End If
    End If
End Sub

Sub HighScoreEntryInit
    PlaySound "vo_enteryourinitials"
    hsLetterFlash = 0

    hsEnteredDigits(0) = " "
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ<+-0123456789" ' < is used to delete the last letter
    hsCurrentLetter = 1
    'UltraDMD Display
    If UltraDMD_Option=True Then
        UltraDMDFlush
        UltraDMDId "hsc", "black.png", "YOUR NAME:", " ", 999999
    End If
    HighScoreDisplayName
End Sub

Sub EnterHighScoreKey(keycode)
    HiScoreEntryTimer.Enabled=False
    HiScoreEntryTimer.Interval=60000
    HiScoreEntryTimer.Enabled=True
    If keycode = LeftFlipperKey Then
        Playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayName
    End If

    If keycode = RightFlipperKey Then
        Playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayName
    End If

    If keycode = StartGameKey OR keycode = PlungerKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<") then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) then
                HighScoreCommitName
            else
                HighScoreDisplayName
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit> 0) then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayName
        end if
    end if
End Sub

Sub HighScoreDisplayName
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
    ' UltraDMD Display
    If UltraDMD_Option=True Then
        UltraDMDMod "hsc", "YOUR NAME:", Mid(TempStr, 2, 5), 999999
    End If
End Sub

Sub HighScoreCommitName
    hsbModeActive = False
    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if hsEnteredName="   " Then
        hsEnteredName = "WOZ"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
    ' UltraDMD Display
    If UltraDMD_Option=True Then UltraDMDFlush
    EndOfBallComplete
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


'********************************************************************************
'*                 UltraDMD SETUP & SUPPORT
'*
'* ckpin - used online manual; however, comma formatting not supported
'********************************************************************************
Dim UltraDMD

' Const UltraDMD_bComma = True
Const UltraDMD_VideoMode_Stretch = 0
Const UltraDMD_VideoMode_Top = 1
Const UltraDMD_VideoMode_Middle = 2
Const UltraDMD_VideoMode_Bottom = 3
Const UltraDMD_Animation_FadeIn = 0
Const UltraDMD_Animation_FadeOut = 1
Const UltraDMD_Animation_ZoomIn = 2
Const UltraDMD_Animation_ZoomOut = 3
Const UltraDMD_Animation_ScrollOffLeft = 4
Const UltraDMD_Animation_ScrollOffRight = 5
Const UltraDMD_Animation_ScrollOnLeft = 6
Const UltraDMD_Animation_ScrollOnRight = 7
Const UltraDMD_Animation_ScrollOffUp = 8
Const UltraDMD_Animation_ScrollOffDown = 9
Const UltraDMD_Animation_ScrollOnUp = 10
Const UltraDMD_Animation_ScrollOnDown = 11
Const UltraDMD_Animation_None = 14

'---------- UltraDMD Unique Table Color preference -------------
Dim DMDColor, DMDColorSelect, UseFullColor
Dim DMDPosition, DMDPosX, DMDPosY, DMDSize, DMDWidth, DMDHeight


UseFullColor = "True"       '          "True" / "False"
DMDColorSelect = "cyan"     ' Rightclick on UDMD window to get full list of colors

DMDPosition = False         ' Use Manual DMD Position, True / False
DMDPosX = 100                ' Position in Decimal
DMDPosY = 40                 ' Position in Decimal

DMDSize = False             ' Use Manual DMD Size, True / False
DMDWidth = 1024               ' Width in Decimal
DMDHeight = 256              ' Height in Decimal

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


Sub LoadUltraDMD
    Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    If UltraDMD is Nothing Then
        MsgBox "No UltraDMD found.  Set UltraDMD_Option to False; can only use DMD."
        Exit Sub
    End If
    UltraDMD.Init

    ' A Major version change indicates the version is no longer backward compatible
    If Not UltraDMD.GetMajorVersion = 1 Then
        MsgBox "Incompatible Version of UltraDMD found."
        Exit Sub
    End If

    'A Minor version change indicates new features that are all backward compatible
    If UltraDMD.GetMinorVersion < 1 Then
        MsgBox "Incompatible Version of UltraDMD found.  Please update to version 1.1 or newer."
        Exit Sub
    End If

    Dim fso, curDir
    Set fso = CreateObject("Scripting.FileSystemObject")
    curDir = fso.GetAbsolutePathName(".")

    Dim DirName
    DirName = curDir& "\" &tempcGameName& ".UltraDMD"

    If Not fso.FolderExists(DirName) Then _
            Msgbox "UltraDMD userfiles directory '" & DirName & "' does not exist." & CHR(13) & "No graphic images will be displayed on the DMD"
    UltraDMD.SetProjectFolder DirName

    ' UltraDMD.DisplayScoresWithComma UltraDMD_bComma  - DOESN'T WORK


End Sub

Sub UltraDMD_simple(background, toptext, bottomtext, duration)
    If Not UltraDMD is Nothing Then
         UltraDMD.DisplayScene00 background, toptext, 15, bottomtext, 15, 14, duration, 14
    End If
End Sub

Sub UltraDMDBlink(background, toptext, bottomtext, duration, nblinks) 'blinks the lower text nblinks times
    Dim i
    For i = 1 to nblinks
        UltraDMD.DisplayScene00 background, toptext, 15, bottomtext, 15, 14, duration, 14
    Next
End Sub


Sub UltraDMD_DisplayScene(bkgnd,toptext,topBrightness, bottomtext, bottomBrightness, animateIn, pauseTime, animateOut)
    If Not UltraDMD is Nothing Then
        UltraDMD.DisplayScene00 bkgnd, toptext, topBrightness, bottomtext, bottomBrightness, animateIn, pauseTime, animateOut
        If pauseTime > 0 OR animateIn < 14 OR animateOut < 14 Then
            UltraDMD_Timer1.Enabled = True
        End If
    End If
End Sub


Sub UltraDMDId(id, background, toptext, bottomtext, duration)
    UltraDMD.DisplayScene00ExwithID id, False, background, toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
End Sub

Sub UltraDMDMod(id, toptext, bottomtext, duration)
    UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
End Sub


Dim UltraDMD_numPlayers
Dim UltraDMD_curPlayer
Dim UltraDMD_p1Score
Dim UltraDMD_p2Score
Dim UltraDMD_p3Score
Dim UltraDMD_p4Score
Dim UltraDMD_numCredits
Dim UltraDMD_ballNum
Dim Ultra_DMD_gameinplay
Dim UltraDMD_Tilted

Sub OnScoreboardChanged
    Dim i
    UltraDMD_numPlayers = PlayersPlayingGame
    UltraDMD_curPlayer = CurrentPlayer
    ' Display can't handle scores in excess of 9,999,999
    For i =1 To 4
        If Score(i)>9999999 Then Score(i)=9999999
    Next
    UltraDMD_p1Score = Score(1)
    UltraDMD_p2Score = Score(2)
    UltraDMD_p3Score = Score(3)
    UltraDMD_p4Score = Score(4)
    UltraDMD_numCredits = Credits
    UltraDMD_ballNum = Balls
    Ultra_DMD_gameinplay = bGameInPLay

    If bWait=True Or bWait_Freeze=True Then Exit Sub

    'If we're in the middle of a scene, don't interrupt it to display the scoreboard.
    If Not UltraDMD.IsRendering Then
        If Ultra_DMD_gameinplay = False Then
            UltraDMD.SetScoreboardBackgroundImage "WOZ_score.png", 13, -1
            UltraDMD.DisplayScoreboard 0, 0, 0, 0, 0, 0, "Credits " & CStr(UltraDMD_numCredits), "GAMEOVER"
        Else
            If bSOTR=0 Then
                UltraDMD.SetScoreboardBackgroundImage "woz_1.png", 13, -1
            Else
                UltraDMD.SetScoreboardBackgroundImage "SOTR_Rainbow1.png", 13, -1
            End If
            UltraDMD.DisplayScoreboard UltraDMD_numPlayers, UltraDMD_curPlayer, UltraDMD_p1Score, _
              UltraDMD_p2Score, UltraDMD_p3Score, UltraDMD_p4Score,  _
             "Player " & CStr(UltraDMD_curPlayer), "Ball " & CStr(UltraDMD_ballNum)
        End If
    End If
End Sub


Sub UltraDMD_DisplayImage(UltraDMD_image)
    UltraDMD.SetScoreboardBackgroundImage UltraDMD_image, 13, -1
    UltraDMD.DisplayScoreboard 0, 0, 0, 0, 0, 0, "", ""
    ' Add delay to allow time to see image
    UltraDMD_Delay(4)
End Sub

Sub UltraDMDScoreNow
    UltraDMDFlush
    UltraDMD_Timer1.Enabled = True
End Sub

Sub UltraDMDFlush
    UltraDMD.CancelRendering
    UltraDMD.Clear
End Sub

Dim tickCount
Dim delay
Dim bWait, bWait_Freeze
tickCount = 0
delay = 0
bWait = False

Sub UltraDMD_Delay(sec)
    delay = sec * 2
    bWait = True
    UltraDMD_Timer2.Enabled = True
End Sub

Sub UltraDMD_Freeze
    bWait_Freeze = True
End Sub

Sub UltraDMD_UnFreeze
    bWait_Freeze = False
    OnScoreboardChanged
End Sub

Sub UltraDMD_Timer1_Timer
    If Not UltraDMD.IsRendering Then
        'When the scene finishes rendering, then immediately display the scoreboard
        UltraDMD_Timer1.Enabled = False
        OnScoreboardChanged
    End If
End Sub

Sub UltraDMD_Timer2_Timer
    tickCount = tickCount + 1
    If tickCount > delay Then
        bWait = False
        tickCount = 0
        delay = 0
        UltraDMD_Timer2.Enabled = False
    End If
End Sub

'***** End UltraDMD ************


' *****************************************************************************
'                           DMD  DISPLAYS ON PLAYFIELD
' *****************************************************************************
' *****************************************************************************
' Uses JP's Reduced Display Driver Functions (based on script by Black)
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
Dim dqHead2
Dim dqTail2
Dim deSpeed
Dim deBlinkSlowRate
Dim deBlinkFastRate

Dim dCharsPerLine(7)
Dim dLine(7)
Dim deCount(7)
Dim deCountEnd(7)
Dim deBlinkCycle(7)

Dim dqText(7, 64)
Dim dqEffect(7, 64)
Dim dqTimeOn(64)
Dim dqbFlush(64)
Dim dqSound(64)
Dim dqTimeOn2(64)
Dim dqbFlush2(64)
Dim dqSound2(64)

Sub DMD_Init 'default/startup values
    Dim i, j
    DMDFlush
    DMD2Flush
    deSpeed = 20
    deBlinkSlowRate = 5
    deBlinkFastRate = 2
    dCharsPerLine(0) = 3
    dCharsPerLine(1) = 16
    dCharsPerLine(2) = 16
    dCharsPerLine(3) = 11

    dCharsPerLine(4) = 3
    dCharsPerLine(5) = 16
    dCharsPerLine(6) = 16
    dCharsPerLine(7) = 11

    For i = 0 to 7
        dLine(i) = Space(dCharsPerLine(i))
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
    Next
    For j = 0 To 64
        dqTimeOn(j) = 0
        dqbFlush(j) = True
        dqSound(j) = ""
        dqTimeOn2(j) = 0
        dqbFlush2(j) = True
        dqSound2(j) = ""
    Next
    For i = 0 to 7
        For j = 0 to 64
            dqText(i, j) = ""
            dqEffect(i, j) = eNone
        Next
    Next
End Sub

Sub DMDFlush
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
End Sub


Sub DMD2Flush
    Dim i
    DMD2Timer.Enabled = False
    DMD2EffectTimer.Enabled = False
    dqHead2 = 0
    dqTail2 = 0
    For i = 4 to 7
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
    Next
End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMDScore
    bDisplay=0
    Dim tmp0, tmp1, tmp2, tmp3
    if(dqHead = dqTail)Then
        tmp0 = ""
        tmp1 = FillLine(1, "SCORE", FormatScore(Score(Currentplayer)))
        tmp2 = FillLine(2, "PLAYER " & CurrentPlayer, " BALL " & Balls)
        tmp3 = ""
        DMD tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 25, True, ""
    End If

    'UltraDMD Add
    If UltraDMD_Option=True Then UltraDMD_Timer1.Enabled = True
End Sub

Sub DMD(Text0, Effect0, Text1, Effect1, Text2, Effect2, Text3, Effect3, TimeOn, bFlush, Sound)
'       The lines are  background. top line, bottom line, and centerline
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
            DMDHead
        End If
    End If
End Sub

Sub DMD2(Text0, Effect0, Text1, Effect1, Text2, Effect2, Text3, Effect3, TimeOn, bFlush, Sound)
'       The lines are  background. top line, bottom line, and centerline

    if(dqTail2 < dqSize)Then
        if(Text0 = "_")Then
            dqEffect(4, dqTail2) = eNone
            dqText(4, dqTail2) = "_"
        Else
            dqEffect(4, dqTail2) = Effect0
            dqText(4, dqTail2) = ExpandLine(Text0, 4)
        End If

        if(Text1 = "_")Then
            dqEffect(5, dqTail2) = eNone
            dqText(5, dqTail2) = "_"
        Else
            dqEffect(5, dqTail2) = Effect1
            dqText(5, dqTail2) = ExpandLine(Text1, 5)
        End If

        if(Text2 = "_")Then
            dqEffect(6, dqTail2) = eNone
            dqText(6, dqTail2) = "_"
        Else
            dqEffect(6, dqTail2) = Effect2
            dqText(6, dqTail2) = ExpandLine(Text2, 6)
        End If

        if(Text3 = "_")Then
            dqEffect(7, dqTail2) = eNone
            dqText(7, dqTail2) = "_"
        Else
            dqEffect(7, dqTail2) = Effect3
            dqText(7, dqTail2) = ExpandLine(Text3, 7)
        End If

        dqTimeOn2(dqTail2) = TimeOn
        dqbFlush2(dqTail2) = bFlush
        dqSound2(dqTail2) = Sound
        dqTail2 = dqTail2 + 1
        if(dqTail2 = 1)Then
            DMD2Head
        End If
    End If
End Sub

Sub DMDHead
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

Sub DMD2Head
    Dim i
    deCount(4) = 0
    deCount(5) = 0
    deCount(6) = 0
    deCount(7) = 0

    DMD2EffectTimer.Interval = deSpeed

    For i = 4 to 7
        Select Case dqEffect(i, dqHead2)
            Case eNone:deCountEnd(i) = 1
            Case eScrollLeft:deCountEnd(i) = Len(dqText(i, dqHead2))
            Case eScrollRight:deCountEnd(i) = Len(dqText(i, dqHead2))
            Case eBlink:deCountEnd(i) = int(dqTimeOn2(dqHead2) / deSpeed)
                deBlinkCycle(i) = 0
            Case eBlinkFast:deCountEnd(i) = int(dqTimeOn2(dqHead2) / deSpeed)
                deBlinkCycle(i) = 0
        End Select
    Next
    if(dqSound2(dqHead2) <> "")Then
        PlaySound(dqSound2(dqHead2))
    End If
    DMD2EffectTimer.Enabled = True
End Sub

Sub DMDEffectTimer_Timer
    DMDEffectTimer.Enabled = False
    DMDProcessEffectOn
End Sub


Sub DMD2EffectTimer_Timer
    DMD2EffectTimer.Enabled = False
    DMD2ProcessEffectOn
End Sub

Sub DMDTimer_Timer
    Dim Head
    DMDTimer.Enabled = False
    Head = dqHead
    dqHead = dqHead + 1
    if(dqHead = dqTail)Then
        if(dqbFlush(Head) = True)Then
            DMDFlush
            DMDScore
        Else
            dqHead = 0
            DMDHead
        End If
    Else
        DMDHead
    End If
End Sub

Sub DMD2Timer_Timer
    Dim Head
    DMD2Timer.Enabled = False
    Head = dqHead2
    dqHead2 = dqHead2 + 1
    if(dqHead2 = dqTail2)Then
        if(dqbFlush2(Head) = True)Then
            DMD2Flush
'            DMDScoreNow
        Else
            dqHead2 = 0
            DMD2Head
        End If
    Else
        DMD2Head
    End If
End Sub


Sub DMDProcessEffectOn
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

    if(deCount(0) = deCountEnd(0)) AND (deCount(1) = deCountEnd(1)) AND (deCount(2) = deCountEnd(2)) _
                    AND (deCount(3) = deCountEnd(3)) Then
        if(dqTimeOn(dqHead) = 0)Then
            DMDFlush
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

Sub DMD2ProcessEffectOn
    Dim i
    Dim BlinkEffect
    Dim Temp

    BlinkEffect = False

    For i = 4 to 7
        if(deCount(i) <> deCountEnd(i))Then
            deCount(i) = deCount(i) + 1

            select case(dqEffect(i, dqHead2))
                case eNone:
                    Temp = dqText(i, dqHead2)
                case eScrollLeft:
                    Temp = Right(dLine(i), dCharsPerLine(i)- 1)
                    Temp = Temp & Mid(dqText(i, dqHead2), deCount(i), 1)
                case eScrollRight:
                    Temp = Mid(dqText(i, dqHead2), (dCharsPerLine(i) + 1)- deCount(i), 1)
                    Temp = Temp & Left(dLine(i), dCharsPerLine(i)- 1)
                case eBlink:
                    BlinkEffect = True
                    if((deCount(i)MOD deBlinkSlowRate) = 0)Then
                        deBlinkCycle(i) = deBlinkCycle(i)xor 1
                    End If

                    if(deBlinkCycle(i) = 0)Then
                        Temp = dqText(i, dqHead2)
                    Else
                        Temp = Space(dCharsPerLine(i))
                    End If
                case eBlinkFast:
                    BlinkEffect = True
                    if((deCount(i)MOD deBlinkFastRate) = 0)Then
                        deBlinkCycle(i) = deBlinkCycle(i)xor 1
                    End If

                    if(deBlinkCycle(i) = 0)Then
                        Temp = dqText(i, dqHead2)
                    Else
                        Temp = Space(dCharsPerLine(i))
                    End If
            End Select

            if(dqText(i, dqHead2) <> "_")Then
                dLine(i) = Temp
                DMD2Update i
            End If
        End If
    Next

    if(deCount(4) = deCountEnd(4)) AND (deCount(5) = deCountEnd(5)) AND (deCount(6) = deCountEnd(6)) _
                    AND (deCount(7) = deCountEnd(7)) Then
        if(dqTimeOn2(dqHead2) = 0)Then
            DMD2Flush
        Else
            if(BlinkEffect = True)Then
                DMD2Timer.Interval = 10
            Else
                DMD2Timer.Interval = dqTimeOn2(dqHead2)
            End If

            DMD2Timer.Enabled = True
        End If
    Else
        DMD2EffectTimer.Enabled = True
    End If
End Sub

Sub DMD2_Clear
    DMD2Flush
    if dqHead2=dqTail2 Then DMD2 "", eNone, "", eNone, "", eNone, "", eNone, 1000, True, ""
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

Function CenterLine(id, aString)
    Dim tmp, tmpStr
    tmp = (dCharsPerLine(id)- Len(aString)) \ 2
    if tmp<1 Then tmp=0
    If(tmp + tmp + Len(aString)) < dCharsPerLine(id) Then
        tmpStr = " " & Space(tmp) & aString & Space(tmp)
    Else
        tmpStr = Space(tmp) & aString & Space(tmp)
    End If
    CenterLine = tmpStr
End Function

Function FillLine(id, aString, bString)
    Dim tmp, tmpStr
    tmp = dCharsPerLine(id)- Len(aString)- Len(bString)
    If tmp<1 Then tmp=0
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
' Update DMD - reels
'*********************

Dim DesktopMode:DesktopMode = Table1.ShowDT

Dim Digits(7)

Sub DMDReels_Init
    Dim obj
    If DesktopMode Then
        'Desktop
        For each obj in DT_EMReel
                obj.Visible = 1
        Next
        For each obj in DT2_EMReel
                obj.Visible = 1
        Next
        For each obj in FS_EMReel
                obj.Visible = 0
        Next
        For each obj in FS2_EMReel
                obj.Visible = 0
        Next
        Digits(0) = Array(d0, d1, d2)          'backdrop
        Digits(1) = Array(d3, d4, d5, d6, d7, d8, d9, d10, d11, d12, d13, d14, d15, d16, d17, d18, d19, d20, d21)        'upper line
        Digits(2) = Array(d22, d23, d24, d25, d26, d27, d28, d29, d30, d31, d32, d33, d34, d35, d36, d37, d38, d39, d40) 'lower line
        Digits(3) = Array(d41, d42, d43, d44, d45, d46, d47, d48, d49, d50, d51, d52, d53)                               ' center line
        d0.Visible = 0:d1.Visible = 0:d2.Visible = 0

        Digits(4) = Array(d162, d163, d164)    'backdrop
        Digits(5) = Array(d165, d166, d167, d168, d169, d170, d171, d172, d173, d174, d175, d176, d177, d178, d179, d180, d181, d182, d183) 'upper line
        Digits(6) = Array(d184, d185, d186, d187, d188, d189, d190, d191, d192, d193, d194, d195, d196, d197, d198, d199, d200, d201, d202) 'lower line
        Digits(7) = Array(d203, d204, d205, d206, d207, d208, d209, d210, d211, d212, d213, d214, d215)
        d162.Visible = 0:d163.Visible = 0:d164.Visible = 0

    Else
        'FS
        For each obj in FS_EMReel
                obj.Visible = 1
        Next
        For each obj in FS2_EMReel
                obj.Visible = 1
        Next
        For each obj in DT_EMReel
                obj.Visible = 0
        Next
        For each obj in DT2_EMReel
                obj.Visible = 0
        Next
        Digits(0) = Array(d54, d55, d56)       'backdrop
        Digits(1) = Array(d57, d58, d59, d60, d61, d62, d63, d64, d65, d66, d67, d68, d69, d70, d71, d72, d73, d74, d75) 'upper line
        Digits(2) = Array(d76, d77, d78, d79, d80, d81, d82, d83, d84, d85, d86, d87, d88, d89, d90, d91, d92, d93, d94) 'lower line
        Digits(3) = Array(d95, d96, d97, d98, d99, d100, d101, d102, d103, d104, d105, d106, d107)
        ' Shorter display width and hide three large background digits used for images
        d73.Visible=0:d74.Visible=0:d75.Visible=0:d92.Visible=0:d93.Visible=0:d94.Visible=0:d106.Visible=0:d107.Visible=0
        d54.Visible = 0:d55.Visible = 0:d56.Visible = 0

        Digits(4) = Array(d108, d109, d110)    'backdrop
        Digits(5) = Array(d111, d112, d113, d114, d115, d116, d117, d118, d119, d120, d121, d122, d123, d124, d125, d126, d127, d128, d129) 'upper line
        Digits(6) = Array(d130, d131, d132, d133, d134, d135, d136, d137, d138, d139, d140, d141, d142, d143, d144, d145, d146, d147, d148) 'lower line
        Digits(7) = Array(d149, d150, d151, d152, d153, d154, d155, d156, d157, d158, d159, d160, d161)
        d108.Visible = 0:d109.Visible = 0:d110.Visible = 0
        d127.Visible=0:d128.Visible=0:d129.Visible=0:d146.Visible=0:d147.Visible=0:d148.Visible=0:d160.Visible=0:d161.Visible=0

    End If
    If DMD_Option=False Then
        For each obj in FS_EMReel
                obj.Visible = 0
        Next
        For each obj in FS2_EMReel
                obj.Visible = 0
        Next
        For each obj in DT_EMReel
                obj.Visible = 0
        Next
        For each obj in DT2_EMReel
                obj.Visible = 0
        Next
    End If
End Sub

Dim bDisplay  '0 - Apron Left, 1 - Apron Right
bDisplay=0
Sub DMDUpdate(id)
    Dim digit, value
    For digit = 0 to dCharsPerLine(id)-1
        value = ASC(mid(dLine(id), digit + 1, 1))-32
        Digits(id)(digit).SetValue value
    Next
End Sub

Sub DMD2Update(id)
    Dim digit, value
    For digit = 0 to dCharsPerLine(id)-1
        value = ASC(mid(dLine(id), digit + 1, 1))-32
        Digits(id)(digit).SetValue value
    Next
End Sub

'***** End DMD ************

'******************************************************************************
'          INSTANT INFO
'******************************************************************************
Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    bInstantInfo = True

    BonusMultiplier(CurrentPlayer)=Multipliers(1)
    If DMD_Option=True Then
        DMD "", eNone, CenterLine(1, "INSTANT INFO"), eNone, CenterLine(2, "YELLOW BRICK RD"), eScrollLeft, CenterLine(3, FormatScore(modeYBR_cnt)), eScrollLeft, 800, False, ""
        DMD "", eNone, CenterLine(1, "INSTANT INFO"), eNone, CenterLine(2, "BONUS MULTIPLIER"), eScrollLeft, CenterLine(3, BonusMultiplier(CurrentPlayer)), eScrollLeft, 800, False, ""
        DMD "", eNone, CenterLine(1, "INSTANT INFO"), eNone, CenterLine(2, "JEWELS COLLECTED"), eScrollLeft, CenterLine(3, FormatScore(SOTR)), eScrollLeft, 800, False, ""
        HOADC_Score_Calc
        DMD "", eNone, CenterLine(1, "INSTANT INFO"), eNone, CenterLine(2, "  HOADC SO FAR"), eScrollLeft, CenterLine(3, FormatScore(HOADC_points)), eScrollLeft, 800, False, ""
    End If

    If UltraDMD_Option=True Then
        UltraDMD_simple "black.png", "", "INSTANT INFO", 1500
        UltraDMD_simple "black.png", "YELLOW BRICK ROAD", modeYBR_cnt, 1500
        UltraDMD_simple "black.png", "BONUS MULTIPLIER", BonusMultiplier(CurrentPlayer), 1500
        UltraDMD_simple "black.png", "JEWELS COLLECTED", SOTR, 1500
        HOADC_Score_Calc
        UltraDMD_simple "black.png", "HOADC SO FAR", HOADC_Points, 1500
    End If

End Sub

Sub EndFlipperStatus
    If bInstantInfo Then
        bInstantInfo = False
        If DMD_Option=True Then DMDScoreNow
        ' UltraDMD Add
        If UltraDMD_Option=True Then UltraDMD_Timer1.Enabled = True
    End If
End Sub

Sub ShowTableInfo
    'info goes in a loop only stopped by the credits and the startkey
    If DMD_Option=True Then
        If Score(1)Then
            DMD "", eNone, CenterLine(1, FormatScore(Score(1))), eNone, CenterLine(2, "PLAYER 1"), eNone, "", eNone, 3000, False, ""
        End If
        If Score(2)Then
            DMD "", eNone, CenterLine(1, FormatScore(Score(2))), eNone, CenterLine(2, "PLAYER 2"), eNone, "", eNone, 3000, False, ""
        End If
        If Score(3)Then
            DMD "", eNone, CenterLine(1, FormatScore(Score(3))), eNone, CenterLine(2, "PLAYER 3"), eNone, "", eNone, 3000, False, ""
        End If
        If Score(4)Then
            DMD "", eNone, CenterLine(1, FormatScore(Score(4))), eNone, CenterLine(2, "PLAYER 4"), eNone, "", eNone, 3000, False, ""
        End If
        DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "GAME OVER "), eBlink, 2000, False, "" 'game over
        If bFreePlay=True Then
            DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "FREE PLAY"), eNone, 3000, False, ""
        Else
            If Credits > 0 Then
                DMD "", eNone, CenterLine(1, "CREDITS " & Credits), eNone, CenterLine(2, "PRESS START"), eNone, "", eNone, 2000, False, ""
            Else
                DMD "", eNone, CenterLine(1, "CREDITS " & Credits), eNone, CenterLine(2, "INSERT COIN"), eNone, "", eNone, 2000, False, ""
            End If
        End If

        DMD "   ", eScrollLeft, "", eNone, "", eNone, "", eNone, 500, False, ""
        DMD "", eNone, CenterLine(1, "HIGHSCORES  "), eScrollLeft, Space(dCharsPerLine(2)), eScrollLeft, "", eNone, 20, False, ""
        DMD "", eNone, CenterLine(1, "HIGHSCORES  "), eBlinkFast, "", eNone, "", eNone, 1000, False, ""
        DMD "", eNone, CenterLine(1, "HIGHSCORES  "), eNone, " 1> " &HighScoreName(0) & " " &FormatScore(HighScore(0)), eScrollLeft, "", eNone, 2000, False, ""
        DMD "", eNone, "_", eNone, " 2> " &HighScoreName(1) & " " &FormatScore(HighScore(1)), eScrollLeft, "", eNone, 2000, False, ""
        DMD "", eNone, "_", eNone, " 3> " &HighScoreName(2) & " " &FormatScore(HighScore(2)), eScrollLeft, "", eNone, 2000, False, ""
        DMD "", eNone, "_", eNone, " 4> " &HighScoreName(3) & " " &FormatScore(HighScore(3)), eScrollLeft, "", eNone, 2000, False, ""
        DMD "", eNone, Space(dCharsPerLine(1)), eScrollLeft, Space(dCharsPerLine(2)), eScrollLeft, "", eNone, 500, False, ""
    End If
    If UltraDMD_Option=True Then
        If Score(1) Then
            UltraDMD_simple "black.png", "PLAYER 1", Score(1), 3000
        End If
        If Score(2) Then
            UltraDMD_simple "black.png", "PLAYER 2", Score(2), 3000
        End If
        If Score(3) Then
            UltraDMD_simple "black.png", "PLAYER 3", Score(3), 3000
        End If
        If Score(4) Then
            UltraDMD_simple "black.png", "PLAYER 4", Score(4), 3000
        End If

        'coins or freeplay
        If bFreePlay=True Then
            UltraDMD_simple "black.png", " ", "FREE PLAY", 2000
'            UltraDMD_simple "intro-freeplay.wmv", "", "", 63000
        Else
            If Credits> 0 Then
                UltraDMD_simple "black.png", "CREDITS " & Credits, "PRESS START", 2000
            Else
                UltraDMD_simple "black.png", "CREDITS " & Credits, "INSERT COIN", 2000
            End If
'            UltraDMD_simple "intro-coins.wmv", "", "", 65000
        End If

        UltraDMD_simple "black.png", "HIGHSCORES", " 1> " & HighScoreName(0) & "  " & FormatNumber(HighScore(0),0,,,0), 3000
        UltraDMD_simple "black.png", "HIGHSCORES", " 2> " & HighScoreName(1) & "  " & FormatNumber(HighScore(1),0,,,0), 3000
        UltraDMD_simple "black.png", "HIGHSCORES", " 3> " & HighScoreName(2) & "  " & FormatNumber(HighScore(2),0,,,0), 3000
        UltraDMD_simple "black.png", "HIGHSCORES", " 4> " & HighScoreName(3) & "  " & FormatNumber(HighScore(3),0,,,0), 3000
        UltraDMD_simple "GameOver.png", "", "", 5000

    End If
End Sub

'******************************************************************************
'                     REALLY SPECIFIC TABLE "STUFF"
'******************************************************************************

'*******************************************************************************
'                    PLAYER STATE CAPTURE  - NOT FULLY TESTED YET
'*******************************************************************************
' Game Specific Player State Snapshots
' ASSUMES MaxPlayers is 4
Dim PL_Lights(4,200)
Dim PL_Flashers(4,8)
Dim PL_Flags(4,16)
Dim PL_SwitchFlags(4,200)
Dim PL_Mode_HU(4,16)
Dim PL_Mode_WitchHU(4,16)
Dim PL_Mode_FireballF(4,16)
Dim PL_Mode_Throne(4,16)
Dim PL_Mode_YBR(4,16)
Dim PL_Mode_HOADC(4,16)
Dim PL_Mode_TOTO(4, 16)
Dim PL_Mode_TNPLH(4, 16)
Dim PL_Mode_Emerald(4,16)
Dim PL_Mode_TWISTER(4,16)
Dim PL_Mode_CRYSTALBALL(4,16)
Dim PL_Mode_MUNCHKIN(4,16)
Dim PL_Mode_Haunted(4,16)
Dim PL_Mode_Dorothy(4,16)
Dim PL_Mode_BTWW(4,16)
Dim PL_Mode_Rescue(4)
Dim PL_Jewels(4)
Dim PL_Mode_TinMan(4)
Dim PL_Mode_Lion(4)
Dim PL_Mode_Scarecrow(4)

Dim bNoCapture

' Capture Player Lights at end of ball
Sub Capture_PlayerState(player)
    Dim obj, bg_id, flash_id, j

    debug.print "CAPTURE PlayerState - Player = " & player
    j = 1
    For each obj in aLEDLights
        PL_Lights(player,j)= obj.State
        j = j+1
    Next
    j = 1
    For each obj in aFlashers
        flash_id = flash_dict.Item(obj)
        PL_Flashers(player,j)= Flashers(flash_id)
        j = j+1
    Next
    For j=1 To 200
        PL_SwitchFlags(player,j) = SwitchFlags(j)
    Next

    PL_Flags(player,1) = bScoreExBall
    PL_Flags(player,2) = bScoreReplay
    PL_Flags(player,3) = bOnTheFirstBall
    PL_Flags(player,4) = bNoCapture
    PL_Flags(player,5) = bBallSaverActive
    PL_Flags(player,6) = bBallSaverReady
    PL_Flags(player,7) = bMusicOn
    PL_Flags(player,8) = bJustStarted
    PL_Flags(player,9) = bBonusHeld
    PL_Flags(player,10) = bFreePlay
    PL_Flags(player,11) = bGameInPlay



    PL_Mode_HU(player,1) = HUPot
    PL_Mode_HU(player,2) = HULow
    PL_Mode_HU(player,3) = cnt_HU_Sound

    PL_Mode_WitchHU(player,1) = WitchHUPot
    PL_Mode_WitchHU(player,2) = WitchHULow
    PL_Mode_WitchHU(player,3) = WitchHUCnt

    PL_Mode_FireballF(player,1) = bFireballFrenzy
    PL_Mode_FireballF(player,2) = bFireballFrenzyStarted
    PL_Mode_FireballF(player,3) = blueShots
    PL_Mode_FireballF(player,4) = blueshotcnt
    PL_Mode_FireballF(player,5) = redshotcnt

    PL_Mode_Throne(player,1) = throneHits
    Wizard_reset

    PL_Mode_YBR(player,1) = modeYBR_cnt
    PL_Mode_YBR(player,2) = YBRBonus

    For j=1 To 6
        PL_Mode_HOADC(player,j) = HOADC_Color(j)
    Next
    For j=7 To 12
        PL_Mode_HOADC(player,j) = HOADC_Queue(j-6)
    Next
    PL_Mode_HOADC(player,13) = HOADC_Base
    PL_Mode_HOADC(player,14) = HOADC_FirstX
    PL_Mode_HOADC(player,15) = HOADC_SecondX
    PL_Mode_HOADC(player,16) = HOADC_collects
    SetLightColor L28,"white",0

    PL_Mode_Emerald(player,1) = bECLock
    PL_Mode_Emerald(player,2) = bECBallsLocked
    PL_Mode_Emerald(player,3) = qualLion
    PL_Mode_Emerald(player,4) = qualTinman
    PL_Mode_Emerald(player,5) = qualScarecrow
    PL_Mode_Emerald(player,6) = qualCnt
    PL_Mode_Emerald(player,7) = giftCnt
    PL_Mode_Emerald(player,8) = ECMB_NextLevel
    Select Case bECBallsLocked
      case 1
        sw6.DestroyBall
        LightState(150)=0:Flashers(flash_dict.Item(l150))=0
      case 2
        sw6.DestroyBall
        LightState(150)=0:Flashers(flash_dict.Item(l150))=0
        sw7.DestroyBall
        LightState(151)=0:Flashers(flash_dict.Item(l151))=0
    End Select

    PL_Mode_TOTO(player,1) = totoCnt
    PL_Mode_TOTO(player,2) = bTOTOMode
    PL_Mode_TOTO(player,3) = bTOTOExhausted

    PL_Mode_TNPLH(player,1) = tnplhCnt
    PL_Mode_TNPLH(player,2) = tnplhStage
    PL_Mode_TNPLH(player,3) = bTNPLH
    PL_Mode_TNPLH(player,4) = bTNPLHMode

    PL_Mode_BTWW(player,1) = bBTWW
    PL_Mode_BTWW(player,2) = bBTWW_1
    PL_Mode_BTWW(player,3) = bBTWW_2
    PL_Mode_BTWW(player,4) = bBTWW_3
    PL_Mode_BTWW(player,5) = bBTWW_4
    PL_Mode_BTWW(player,6) = BTWWCnt
    PL_Mode_BTWW(player,7) = BTWWjackpot
    Controller.B2sSetData "Quadrant_Lt1",0
    Controller.B2sSetData "Quadrant_Lt2",0
    Controller.B2sSetData "Quadrant_Lt3",0
    Controller.B2sSetData "Quadrant_Lt4",0


    PL_Mode_TWISTER(player,1) = twisterloops
    PL_Mode_TWISTER(player,2) = twisterincrement
    PL_Mode_TWISTER(player,3) = twisterbonus
    PL_Mode_TWISTER(player,4) = twisterModeCnt
    PL_Mode_TWISTER(player,5) = loopsleft
    PL_Mode_TWISTER(player,6) = bTwisterFlag

    PL_Mode_CRYSTALBALL(player,1) = bCrystalBallMode
    PL_Mode_CRYSTALBALL(player,2) = bCB_LightsON
    PL_Mode_CRYSTALBALL(player,3) = bCB_LightsOFF
    PL_Mode_CRYSTALBALL(player,4) = bCB_LinkedFlip
    PL_Mode_CRYSTALBALL(player,5) = bCB_NoHoldFlip
    PL_Mode_CRYSTALBALL(player,6) = bCB_FlipFrenzy
    PL_Mode_CRYSTALBALL(player,7) = crystalBallModeCnt
    PL_Mode_CRYSTALBALL(player,8) = cb_Color
    PL_Mode_CRYSTALBALL(player,9) = CBBonus
    PL_Mode_CRYSTALBALL(player,10) = CBMB_Bonus

    PL_Mode_MUNCHKIN(player,1) = bMunchkinMode
    PL_Mode_MUNCHKIN(player,2) = bMunchkinLand
    PL_Mode_MUNCHKIN(player,3) = bMunchkinFrenzy
    PL_Mode_MUNCHKIN(player,4) = bMunchkinLollipop
    PL_Mode_MUNCHKIN(player,5) = bMunchkinMultiBall
    PL_Mode_MUNCHKIN(player,6) = bMunchkinMB
    PL_Mode_MUNCHKIN(player,7) = munchkinModeCnt
    PL_Mode_MUNCHKIN(player,8) = munchkinBonus
    PL_Mode_MUNCHKIN(player,9) = munchkinMB_Bonus
    PL_Mode_MUNCHKIN(player,10) = munchkinX

    PL_Mode_Haunted(player,1) = bHauntedMode
    PL_Mode_Haunted(player,2) = bHauntedShots
    PL_Mode_Haunted(player,3) = bHauntedTargets
    PL_Mode_Haunted(player,4) = bHauntedHoles
    PL_Mode_Haunted(player,5) = bHauntedBumpers
    PL_Mode_Haunted(player,6) = bHauntedMB
    PL_Mode_Haunted(player,7) = HauntedModeCnt
    PL_Mode_Haunted(player,8) = HauntedBonus
    PL_Mode_Haunted(player,9) = HauntedMB_Bonus
    PL_Mode_Haunted(player,10) = hauntedLtrCnt

    PL_Mode_Rescue(player) = ballRescueLock

    PL_Jewels(player) = SOTR

    PL_Mode_TinMan(player) = cnt_Tinman
    PL_Mode_Lion(player) = cnt_Lion
    PL_Mode_Scarecrow(player) = cnt_Scarecrow

    PL_Mode_Dorothy(player,1) = bDorothyCaught
    PL_Mode_Dorothy(player,2) = bDorothyCaptureRdy
    PL_Mode_Dorothy(player,3) = BallsInLock
    PL_Mode_Dorothy(player,4) = bBallJustLocked
    If bDorothyCaught=1 Then
        Controller.B2SSetData 180,0
        sw72_out.DestroyBall
        LightState(153)=0:Flashers(flash_dict.Item(l153))=0
    End If
    If bDorothyCaught=0 AND Flashers(flash_dict.Item(l115))=1 Then
        LightState(115)=0:Flashers(flash_dict.Item(l115))=0
        Kicker_Mag.enabled=False
    End If
End Sub

Sub Restore_PlayerState
    Dim obj, bg_id, flash_id, player, j, cb_State

    player = CurrentPlayer
    debug.print "RESTORE PlayerState - Player = " & player
    j = 1
    For each obj in aLEDLights
        obj.State = PL_Lights(player,j)
        j = j+1
    Next
    j = 1
    For each obj in aFlashers
        flash_id = flash_dict.Item(obj)
        Flashers(flash_id) = PL_Flashers(player,j)
        j = j+1
    Next
    For j=1 To 200
        SwitchFlags(j) = PL_SwitchFlags(player,j)
    Next

    bScoreExBall = PL_Flags(player,1)
    bScoreReplay = PL_Flags(player,2)
    bOnTheFirstBall = PL_Flags(player,3)
    bNoCapture = PL_Flags(player,4)
    bBallSaverActive = PL_Flags(player,5)
    bBallSaverReady = PL_Flags(player,6)
    bMusicOn = PL_Flags(player,7)
    bJustStarted = PL_Flags(player,8)
    bBonusHeld = PL_Flags(player,9)
    bFreePlay = PL_Flags(player,10)
    bGameInPlay = PL_Flags(player,11)


    GiOff
    HUPot = PL_Mode_HU(player,1)
    HULow = PL_Mode_HU(player,2)
    cnt_HU_Sound = PL_Mode_HU(player,3)

    WitchLower
    WitchHUPot = PL_Mode_WitchHU(player,1)
    WitchHULow = PL_Mode_WitchHU(player,2)
    WitchHUCnt = PL_Mode_WitchHU(player,3)

    bFireballFrenzy = PL_Mode_FireballF(player,1)
    bFireballFrenzyStarted = PL_Mode_FireballF(player,2)
    blueShots = PL_Mode_FireballF(player,3)
    blueshotcnt = PL_Mode_FireballF(player,4)
    redshotcnt = PL_Mode_FireballF(player,5)

    throneHits = PL_Mode_Throne(player,1)
    Select Case throneHits
      case 1
        w_sign.visible=1
      case 2
        w_sign.visible=1: i_sign.visible=1
      case 3
        w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
      case 4
        w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
        a_sign.visible=1
     case 5
       w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
       a_sign.visible=1:r_sign.visible=1
     case 6
       w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
       a_sign.visible=1: r_sign.visible=1: d_sign.visible=1
    End Select

    modeYBR_cnt = PL_Mode_YBR(player,1)
    YBRBonus = PL_Mode_YBR(player,2)

    For j=1 To 6
        HOADC_Color(j) = PL_Mode_HOADC(player,j)
    Next
    For j=7 To 12
        HOADC_Queue(j-6) = PL_Mode_HOADC(player,j)
    Next
    HOADC_Base = PL_Mode_HOADC(player,13)
    HOADC_FirstX = PL_Mode_HOADC(player,14)
    HOADC_SecondX = PL_Mode_HOADC(player,15)
    HOADC_collects = PL_Mode_HOADC(player,16)
    If HOADC_Queue(4 + HOADC_collects)<>"none" Then SetLightColor L28,"white",2

    bECLock = PL_Mode_Emerald(player,1)
    bECBallsLocked = PL_Mode_Emerald(player,2)
    qualLion = PL_Mode_Emerald(player,3)
    qualTinman = PL_Mode_Emerald(player,4)
    qualScarecrow = PL_Mode_Emerald(player,5)
    qualCnt = PL_Mode_Emerald(player,6)
    giftCnt = PL_Mode_Emerald(player,7)
    ECMB_NextLevel = PL_Mode_Emerald(player,8)
    Select Case bECBallsLocked
      case 0
        LightState(150)=0:LightState(151)=0
        Flashers(flash_dict.Item(l150))=0:Flashers(flash_dict.Item(l151))=0
      case 1
        sw6.createball
        LightState(150)=1:Flashers(flash_dict.Item(l150))=1
      case 2
        sw6.createball
        LightState(150)=1:Flashers(flash_dict.Item(l150))=1
        sw7.createball
        LightState(151)=1:Flashers(flash_dict.Item(l151))=1
    End Select
    totoCnt = PL_Mode_TOTO(player,1)
    bTOTOMode = PL_Mode_TOTO(player,2)
    bTOTOExhausted = PL_Mode_TOTO(player,3)

    tnplhCnt = PL_Mode_TNPLH(player,1)
    tnplhStage = PL_Mode_TNPLH(player,2)
    bTNPLH = PL_Mode_TNPLH(player,3)
    bTNPLHMode = PL_Mode_TNPLH(player,4)

    TNPLH_TOTO_LightsReset

    bBTWW = PL_Mode_BTWW(player,1)
    bBTWW_1 = PL_Mode_BTWW(player,2)
    bBTWW_2 = PL_Mode_BTWW(player,3)
    bBTWW_3 = PL_Mode_BTWW(player,4)
    bBTWW_4 = PL_Mode_BTWW(player,5)
    BTWWCnt = PL_Mode_BTWW(player,6)
    BTWWjackpot = PL_Mode_BTWW(player,7)
    If bBTWW_1=1 Then Controller.B2sSetData "Quadrant_Lt1",1
    If bBTWW_2=1 Then Controller.B2sSetData "Quadrant_Lt2",1
    If bBTWW_3=1 Then Controller.B2sSetData "Quadrant_Lt3",1
    If bBTWW_4=1 Then Controller.B2sSetData "Quadrant_Lt4",1

    twisterloops = PL_Mode_TWISTER(player,1)
    twisterincrement = PL_Mode_TWISTER(player,2)
    twisterbonus = PL_Mode_TWISTER(player,3)
    twisterModeCnt = PL_Mode_TWISTER(player,4)
    loopsleft = PL_Mode_TWISTER(player,5)
    bTwisterFlag = PL_Mode_TWISTER(player,6)

    bCrystalBallMode = PL_Mode_CRYSTALBALL(player,1)
    bCB_LightsON = PL_Mode_CRYSTALBALL(player,2)
    bCB_LightsOFF = PL_Mode_CRYSTALBALL(player,3)
    bCB_LinkedFlip = PL_Mode_CRYSTALBALL(player,4)
    bCB_NoHoldFlip = PL_Mode_CRYSTALBALL(player,5)
    bCB_FlipFrenzy = PL_Mode_CRYSTALBALL(player,6)
    crystalBallModeCnt = PL_Mode_CRYSTALBALL(player,7)
    cb_Color = PL_Mode_CRYSTALBALL(player,8)
    CBBonus = PL_Mode_CRYSTALBALL(player,9)
    CBMB_Bonus = PL_Mode_CRYSTALBALL(player,10)
'    For each obj in Lights_CrystalBall
'        cb_State = obj.State
'        SetLightColor obj, cb_Color, cb_State
'    Next

    bMunchkinMode = PL_Mode_MUNCHKIN(player,1)
    bMunchkinLand = PL_Mode_MUNCHKIN(player,2)
    bMunchkinFrenzy = PL_Mode_MUNCHKIN(player,3)
    bMunchkinLollipop = PL_Mode_MUNCHKIN(player,4)
    bMunchkinMultiBall = PL_Mode_MUNCHKIN(player,5)
    bMunchkinMB = PL_Mode_MUNCHKIN(player,6)
    munchkinModeCnt = PL_Mode_MUNCHKIN(player,7)
    munchkinBonus = PL_Mode_MUNCHKIN(player,8)
    munchkinMB_Bonus = PL_Mode_MUNCHKIN(player,9)
    munchkinX = PL_Mode_MUNCHKIN(player,10)

    bHauntedMode = PL_Mode_Haunted(player,1)
    bHauntedShots = PL_Mode_Haunted(player,2)
    bHauntedTargets = PL_Mode_Haunted(player,3)
    bHauntedHoles = PL_Mode_Haunted(player,4)
    bHauntedBumpers = PL_Mode_Haunted(player,5)
    bHauntedMB = PL_Mode_Haunted(player,6)
    HauntedModeCnt = PL_Mode_Haunted(player,7)
    HauntedBonus = PL_Mode_Haunted(player,8)
    HauntedMB_Bonus = PL_Mode_Haunted(player,9)
    hauntedLtrCnt = PL_Mode_Haunted(player,10)

    ballRescueLock = PL_Mode_Rescue(player)

    SOTR = PL_Jewels(player)

    cnt_Tinman = PL_Mode_TinMan(player)
    cnt_Lion = PL_Mode_Lion(player)
    cnt_Scarecrow = PL_Mode_Scarecrow(player)

    RestoreBG_Rainbow
    If l134.State=2 Then DivRight.RotateToEnd: SetLightColor L134,"yellow",2:Controller.B2SSetData 233,1:bTwistAgain=0

    For each obj in Lights_Haunted
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
    For each obj in Lights_Scarecrow
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
    For each obj in Lights_Lion
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
    For each obj in Lights_Tinman
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next

    Select Case SwitchFlags(15)
      case 1
        w_sign.visible=1
      case 2
        w_sign.visible=1: i_sign.visible=1
      case 3
        w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
      case 4
        w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
        a_sign.visible=1
      case 5
        w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
        a_sign.visible=1: r_sign.visible=1
     End Select

    bDorothyCaught = PL_Mode_Dorothy(player,1)
    bDorothyCaptureRdy = PL_Mode_Dorothy(player,2)
    BallsInLock = PL_Mode_Dorothy(player,3)
    bBallJustLocked = PL_Mode_Dorothy(player,4)
    If bDorothyCaught = 0 Then
        Controller.B2SSetData 180,0
        If Flashers(flash_dict.Item(l115))=1 Then
            Kicker_Mag.Enabled=True: LightState(115)=1
        End If
    Else
        sw72_out.CreateBall
        SetSwitchFlags 72
        LightState(153)=1:Flashers(flash_dict.Item(l153))=1
    End If
End Sub

'*******************************************************************************
'                            TABLE OBJECT SETUP
'*******************************************************************************
' Game Specific Flags
Dim bDorothyCaught, bDorothyCaptureRdy
Dim bOZLeft, bOZRight
Dim bBLit, bALit, bL1Lit, bL2Lit
Dim bECMultiB, bECLock, bECBallsLocked

Dim throneHits


Dim Multipliers(6)
' Multiplier 0 - Not Used
' Multiplier 1 - OZ lanes
' Multiplier 2 - Horse Of A Different Color (HOADC) Horses or Failure
' Multiplier 3 - Horse Of A Different Color (HOADC) Color Factor
' Multiplier 4 - Number of Active Play Fields
' Multiplier 5 - Munchkin Mode
' Multiplier 6 - Crystal Ball Modes

'*******************************************************************************
'   CALLED AT TABLE INITIALIZATION
'*******************************************************************************
Sub TableObject_Setup

    DiverterRight False    ' Directs Ball to Munchkin House
    MunchkinPost False     ' Will NOT capture balls routed to Munchkin House

    ' Crystal Ball Setup and Run - Autonomous to Game activity at present
    Hologram.visible = 0
    Primitive6.visible = 1
    HologramStart

    ClearOZLaneFlashers

End Sub

'*******************************************************************************
'   CALLED AT THE START OF A NEW GAME
'*******************************************************************************
'Initialize Flags & Variables with Game Scope that are specific to this game
Sub Game_Init
    Dim i, j, z, obj
    Dim get_color

    ' Arrays for Holding Player Specific State Information
    For i = 1 To MaxPlayers
        For j = 0 to 200:PL_Lights(i,j) = 0:Next
        For j = 0 to 8:PL_Flashers(i,j) = 0:Next
        For j = 0 to 16:PL_Flags(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_HU(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_WitchHU(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_FireballF(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_YBR(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_HOADC(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_TOTO(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_TNPLH(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_TWISTER(i,j) =0:Next
        For j = 0 to 16:PL_Mode_CRYSTALBALL(i,j) =0:Next
        For j = 0 to 16:PL_Mode_MUNCHKIN(i,j) =0:Next
        For j = 0 to 16:PL_Mode_Haunted(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_Dorothy(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_Emerald(i,j) = 0:Next
        For j = 0 to 16:PL_Mode_BTWW(i,j) = 0:Next
        PL_Mode_Rescue(i) = 0
        PL_Jewels(i) = 0
        PL_Mode_TinMan(i) = 0
        PL_Mode_Lion(i) = 0
        PL_Mode_Scarecrow(i) = 0
    Next

    ' Zero Switch Hits
    For z = 0 to UBound(SwitchFlags):SwitchFlags(z) = 0:Next
    bEndOfGame=0

    ' Clear Flasher State Array
    For z = 0 to UBound(Flashers):Flashers(z) = 0:Next

    ' Clear Score Flags
    bScoreExBall=0
    bScoreReplay=0

    ' Flag for Capturing Player State
    bNoCapture=0

    ' SCOOP lights
    For each obj in aScoopLights
        get_color = lcolor_dict.Item(obj)
        SetLightColor obj,get_color,1
    Next

    ' Skill Shot
    bSkillshotSelect = 1  ' Support for OZ Lane SkillShot Movement
    SkillshotValue(1) = 5000  ' OZ lane light
    SkillshotValue(2) = 5000  ' Witch
    SkillshotValue(3) = 7500  ' Haunted Forest Target
    SkillshotValue(4) = 10000 ' Crystal BALL light

    ' GameOver Animation Reset
    If PUPpack = False Then Controller.B2SSetData 200,0

    ' InLane HurryUps
    bHU = 0
    HUPot = 0
    HULow = 0
    cnt_HU_Sound = 0
    activeHU = ""

    ' Witch HurryUp & Position
    bWitchHU = 0
    bWitchDown=1
    witchHitCnt = 0
    WitchHUCnt = 0
    witchHUPot = 0
    witchHULow = 0
    witchStep = 5
    witchDirection = 1

    ' OZ Lane States

    ' Throne Room - wizard
    throneHits=0

    ' Munchkin House Kickers
'    sw6.enabled = False:sw7.enabled = False:sw8.enabled = False

    ' Emerald City
    bECMultiB = 0
    bECLock = 0
    bECBallsLocked = 0
    ' Lock Pin Not Visible
    MunchkinPost False
    qualLion=0:qualTinman=0:qualScarecrow=0: qualCnt=0: giftCnt=0: ECMB_NextLevel=0
    cnt_Lion=0:cnt_Scarecrow=0:cnt_Tinman=0

    ' Twister
    bTwisterFlag = 0
    bTwistAgain=0
    twisterloops=0
    twisterbonus = 0
    twisterModeCnt = 0
    loopsleft = 0
    L_Ones = "L1_0": L_Tens = "Ld_1"
    Y_Ones = "Y1_0": Y_Tens = "Yd_1"
    T_Ones="M1_0":T_Tens="Md_1"
    hStep=0

    ' Munchkin
    bMunchkinMode=0: bMunchkinLand=0: bMunchkinFrenzy=0: bMunchkinLollipop=0: bMunchkinMultiBall=0
    bMunchkinMB=0
    bOnMunchkinPF=0
    munchkinTimeLeft=0: munchkinModeCnt=0: munchkinBonus=0:munchkinMB_Bonus=0: munchkinX=1
    munchkinMB_Rcvd=0

    ' Crystal Ball
    bFlipFlag=0
    bCrystalBallMode=0
    bCB_LightsON=0: bCB_LightsOFF=0: bCB_LinkedFlip=0: bCB_NoHoldFlip=0: bCB_FlipFrenzy=0
    crystalBallModeCnt=1: CBBonus=0: CBMB_Bonus=0
    cb_Color = "blue"

    ' Haunted
    bHauntedMode = 0:bHauntedLtrMode = 0
    bHauntedShots=0:bHauntedTargets=0:bHauntedHoles=0:bHauntedBumpers=0:bHauntedMB=0
    hauntedLtrCnt=0:hauntedModeCnt=0:hauntedTimeLeft=0:hauntedBonus=0:hauntedMB_Bonus=0
    H_Ones="R1_0":H_Tens="Rd_1"

    ' Castle & Rescue & Dorothy
    Kicker_Mag.Enabled=False
    bDorothyCaught = 0
    bDorothyCaptureRdy = 0
    bRescueMB=0
    bCastlePFActive = 0
    bCastleDoorOpen = 0
    castleDoorBashCnt = 0
    ballRescueLock=0
    sw73.IsDropped=0:SetLightsOff(73)
    Wall71.IsDropped=0
    Wall108.IsDropped = 0

    ' WIZARD
    Wizard_reset
    wiz_points = 2500

    ' Fireball Frenzy
    bFireballFrenzy = 0
    bFireballFrenzyStarted = 0
    blueShots = 0
    blueshotcnt = 0
    redshotcnt = 0
    J_Ones="J1_0":J_Tens="Jd_1"

    ' Battle The Wicked Witch - BTWW
    bBTWW = 0
    bBTWW_1 = 0
    bBTWW_2 = 0
    bBTWW_3 = 0
    bBTWW_4 = 0
    bBTWW_check = 0

    ' RESCUE
    sw17.TimerInterval = 1000     ' Castle VUK Timer Set
    bRescueMB=0
    bCastlePFActive = 0
    bCastleDoorOpen = 0
    bMegaPot = 0
    castleDoorBashCnt = 0
    ballRescueLock=0

    ' Ding Dong
    bDingDong = 0
    ddWhiteShots = 5000

    ' Yellow Brick Road
    modeYBR_cnt = 0
    YBRBonus = 0

    ' TOTO Mode
    totoCnt=0
    bTOTOMode=0
    bTOTODone=0
    bTOTOExhausted=0
    TOTOTimeLeft=0

    ' TNPLH Mode
    bTNPLHMode=0
    bTNPLH = 0
    TNPLHTimeLeft=0
    tnplhCnt=0
    tnplhStage=1

    ' HOADC
    For i = 0 To 6
        HOADC_Queue(i) = "none"
    Next
    HOADC_Base=0:HOADC_FirstX=1:HOADC_SecondX=1
    HOADC_Collects = 0
    bHOADC_Score = 0

    ' SOTR
    SOTR = 0
    SOTR_capture = 0

    ' Jewels
    For each obj in Lights_Jewels
        SetLightColor obj,"white",0
    Next

    ' Glinda Light
    L135.State=0

End Sub

Sub ResetFirstBallModes           'reset stuff for first ball
    ' TOTO Mode Lights
    SetLightColor l37,"red",2
    SetLightColor l38,"red",2
    SetLightColor l39,"red",2
    SetLightColor l40,"red",2
    ' TNPLH Mode Lights
    SetLightColor l59,"green",2
    SetLightColor l60,"green",2
    SetLightColor l61,"green",2
    SetLightColor l62,"green",2
    SetLightColor l63,"green",2
    ' YBR
    YBRBonus = 0
    modeYBR_cnt=0
    YBR_Setup          ' Reset YBR lights
End Sub

'*******************************************************************************
'   CALLED AT THE START OF A NEW BALL
'*******************************************************************************
' Called at the start of a new ball
Sub NewBall_Init
    Dim obj, get_color
    Dim i

    mBalls2Eject = 0
    bBallJustLocked = 0

    ' Ball Loss Check
    bBallLossCheck = 1
    BallLossCheck.Interval = 40000  ' Too long a time between switch hits
    BallLossCheck.Enabled = True

    ' SuperX light reset
    SetLightColor l48,"white",0

    ' Song Reset
    Song = ""

    ' "COWARDLY" Lights - Always ON
    For each obj in acowardly
        SetLightColor obj, "green", 1
    Next

    ' SCOOP lights
    For each obj in aScoopLights
        get_color = lcolor_dict.Item(obj)
        SetLightColor obj,get_color,1
    Next

    ' Yellow Brick Road Start Lights - Always ON
    For each obj in Lights_YBR_Start
        SetLightColor obj,"yellow",1
    Next

    ' TOTO Mode Lights
    If l37.State=0 AND l38.State=0 AND l39.State=0 AND l40.State= 0 Then
        SetLightColor l37,"red",2
        SetLightColor l38,"red",2
        SetLightColor l39,"red",2
        SetLightColor l40,"red",2
    End If
    ' TNPLH Mode Lights
    If l59.State=0 AND l60.State=0 AND l61.State=0 AND l62.State= 0 AND l63.State=0 Then
        SetLightColor l59,"green",2
        SetLightColor l60,"green",2
        SetLightColor l61,"green",2
        SetLightColor l62,"green",2
        SetLightColor l63,"green",2
    End If

   ' Init Variables - Ball Scope
    For i = 0 to UBound(SwitchFlags):SwitchFlags(i) = 0:Next
    For i = 0 to UBound(Multipliers):Multipliers(i) = 1:Next
    twisterbonus=0:twisterloops=0
    bPreviousPosted = 1

   ' Init Flags - Ball Scope
   ' Activate the ballsaver
    If bBallSaverActive=True Then bBallSaverReady = True

   ' Activate the skillshot
    bSkillShotReady = 1



    ' Choose Skill Shots
    ChooseSkillShot

    ' Choose Hurry Up
    HU_Select

    ' HOADC - Select Lights and color
    HOADC_Select

    ' YBR Choose Shot
    YBR_ChooseShot

    ' Raise Winkie
    sw73.IsDropped = 0: SwitchFlags(73)=0:SetLightsOff(73):Controller.B2SSetData 179,0

    ' Balls on Munchkin & Castle PF
    ballsOnMunchkinPF = 0
    ballsOnCastlePF = 0

    ' Fireball Frenzy
    ff_lock = 0

End Sub

'*******************************************************************************
'   SCORING
'*******************************************************************************
' Add points to the score and update score displays by using switch number
Sub AddScore(switch)
    Dim i, x, Multiplier, rainbowhits
    Dim h_title, hits_score

    If Tilted Then Exit Sub
    points = BasePoints(switch)
    Multiplier = 1
    For x = 0 to UBound(Multipliers)
        If Multipliers(x)<>0 Then Multiplier = Multiplier*Multipliers(x)
    Next
    ' Add the points to the current players score
    Score(CurrentPlayer) = Score(CurrentPlayer) + points * Multiplier
    ' Update score displays
    If DMD_Option=True Then DMDScoreNow
    'UltraDMD Add
    If UltraDMD_Option=True Then UltraDMD_Timer1.Enabled = True

    ' Check for special mode handling
    If bHauntedMode=1 Then
        Select Case Switch
          Case 9,16,17,24,71,73,75,101
            If bHauntedShots=1 Then
                hauntedBonus = hauntedBonus + hauntedShots
                hauntedShots = hauntedShots + hauntedIncrease
                h_title = " HAUNTED SHOTS": hits_score = hauntedShots
                HauntedDMD2Display h_title, hits_score
            End If
            If bHauntedMB=1 Then hauntedMB_Bonus = hauntedMB_Bonus + hauntedShots
          Case 29,30,31,32,43,44,45,46,47,52,57,58,61,62,65,66,67,68,68,70,71,73,74,77,89,90,91,92,93,94,95
            If bHauntedTargets=1 Then
                hauntedBonus = hauntedBonus + hauntedTargets
                hauntedTargets = hauntedTargets + hauntedIncrease
                h_title = "HAUNTED TARGETS": hits_score = hauntedTargets
                HauntedDMD2Display h_title, hits_score
            End If
            If bHauntedMB=1 Then hauntedMB_Bonus = hauntedMB_Bonus + hauntedTargets
          Case 15,17,18,21
            If bHauntedHoles=1 Then
                hauntedBonus = hauntedBonus + hauntedHoles
                hauntedHoles = hauntedHoles + hauntedIncrease
                h_title = "  HAUNTED HOLES": hits_score = hauntedHoles
                HauntedDMD2Display h_title, hits_score
            End If
            If bHauntedMB=1 Then hauntedMB_Bonus = hauntedMB_Bonus + hauntedHoles
          Case 3,41,49,50,51
            If bHauntedBumpers=1 Then
                hauntedBonus = hauntedBonus + hauntedBumpers
                hauntedBumpers = hauntedBumpers + hauntedIncrease
                h_title = "HAUNTED BUMPERS": hits_score = hauntedBumpers
                HauntedDMD2Display h_title, hits_score
            End If
            If bHauntedMB=1 Then hauntedMB_Bonus = hauntedMB_Bonus + hauntedBumpers
        End Select
    End If
    If bMunchkinLand=1 Then
        rainbowhits = 0
        For i=89 To 95
            rainbowHits = rainbowHits + SwitchFlags(i)
        Next
        MunchkinBonus = MunchkinBonus + 1000 + 100*rainbowHits
    End If
    If bMunchkinFrenzy=1 Then
        rainbowhits = 0
        For i=89 To 95
            rainbowHits = rainbowHits + SwitchFlags(i)
        Next
        MunchkinBonus = MunchkinBonus + 100 + 10*rainbowHits
    End If

    ' Check for replay or extraball
    If Score(CurrentPlayer)>ExtraBallScore AND bScoreExBall=0 Then AwardExtraBall:bScoreExBall=1
    If Score(CurrentPlayer)>ReplayScore AND bScoreReplay=0 Then AwardSpecial:bScoreReplay=1

End Sub

' Add points to the score and update score by specifying points
Dim points
Sub AddScore_points(points)

    If Tilted Then Exit Sub

    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + points
    debug.print "Current Player Score = " & Score(CurrentPlayer)
    ' update the score displays
    If DMD_Option=True Then DMDScoreNow
    'UltraDMD Add
    If UltraDMD_Option=True Then UltraDMD_Timer1.Enabled = True

    ' Check for replay or extraball
    If Score(CurrentPlayer)>ExtraBallScore AND bScoreExBall=0 Then AwardExtraBall:bScoreExBall=1
    If Score(CurrentPlayer)>ReplayScore AND bScoreReplay=0 Then AwardSpecial:bScoreReplay=1

End Sub

Sub AwardExtraBall
' This results in only one extra ball award per ball; however, it is not explicitly stated
' in the rules or information reviewed that this is a valid rule. Remove outside If statement
' if multiple extraball awards per ball are allowed. Tested both ways.
    If NOT bExtraBallWonThisBall Then
        ' DMD Display
        If UltraDMD_Option=True AND bSOTR=0 Then
            UltraDMDBlink "black.png", " ", "EXTRA BALL WON", 100, 10
        End If
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
		DOF 121, DOFPulse
		DOF 112, DOFPulse
        bExtraBallWonThisBall = True
        If bCB_LightsOFF=0 AND bCB_LightsON=0 Then
            GiEffect 1, 5
            LightEffect 2
        End If
        L49.BlinkInterval = 160
        SetLightColor L49, "orange", 2
    End If
    SetLightColor l29,"white",0
End Sub

Sub AwardSpecial
    If bFreePlay=True Then Exit Sub
    If PUPpack=True Then DOF 82,DOFpulse
    ' DMD Display
    If UltraDMD_Option=True AND bSOTR=0 Then
        UltraDMDBlink "black.png", " ", "EXTRA GAME WON", 100, 10
    End If
    If UltraDMD_Option=True AND bSOTR=1 Then
        UltraDMDBlink "black.png", " ", "THREE RAINBOWS!", 100, 10
    End If
    Credits = Credits + 1
    DOF 114, DOFOn
	DOF 121, DOFPulse
	DOF 112, DOFPulse
    If bCB_LightsOFF=0 AND bCB_LightsON=0 Then
        GiEffect 1, 10
        LightEffect 1
    End If
    SetLightColor l31,"white",0
End Sub

Sub AwardJewel
    SOTR = SOTR + 1

    ' DMD Display
    If UltraDMD_Option=True Then
'        UltraDMDBlink "black.png", " ", "JEWEL AWARDED", 100, 10
        UltraDMD_simple "black.png", "JEWEL AWARDED", "TOTAL JEWELS: " &Cstr(SOTR), 1500
    End If

    Dim tmp0, tmp1, tmp2, tmp3
    If DMD_Option=True Then
        DMD2Flush
        if(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "  JEWEL AWARDED"
            tmp2 = FillLine(2, "TOTAL JEWELS:", FormatScore(SOTR))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
    End If

    If SOTR=8 Then
        Controller.B2SSetData 201,1   ' Blink Center
        If PUPpack=True Then DOF 341, DOFPulse
        SOTR_capture=1:Kicker_TOTO.enabled=True: hold_time=30000
    End If
End Sub

'*******************************************************************************
'                        SWITCH EVENT PROCESSING
'*******************************************************************************
' TRYING TO ROUTE ALL GAME SPECIFIC SWITCH EVENTS HERE
' This isolates and helps control the interactions of the game inputs; however,
' need to watch for performance and asynchronous issues.
Sub SwitchHitEvent(switch)
    bBallLossCheck = 0
    If Tilted Then Exit Sub
    If bGameInPLay=False Then Exit Sub
    SoundPlay switch
    If bFreeze=1 Then Exit Sub
    If bTOTOMode=1 Then TOTOMode_Check switch: Exit Sub
    If bTNPLHMode=1 Then TNPLHMode_Check switch: Exit Sub
    If bDisableTable=1 AND bSW75_RampEnter=0 Then Exit Sub
    If bBTWW=1 Then BTWWMode_Check switch: Exit Sub
    If bDingDong=1 Then DingDongMode_Check switch: Exit Sub
    If bSOTR=1 Then SOTRMode_Check switch: Exit Sub
    If bCB_LightsOFF=0 AND bCB_LightsON=0 Then SetLights switch
    SetSwitchFlags switch
    AddScore switch
    SetDOF switch
End Sub

'*******************************************************************************
'                       SWITCH ASSOCIATED SOUNDS
'*******************************************************************************
Sub SoundPlay(switch)
    Dim get_sound
    If sound_dict.Exists(switch) Then
        get_sound = sound_dict.Item(switch)
        Playsound get_sound
    End If
End Sub

'*******************************************************************************
'                    SWITCH ASSOCIATED LIGHT EFFECTS
'*******************************************************************************
Sub SetLights(switch)
    Dim get_light
    Dim get_color
    If Tilted Then Exit Sub
    If bOnMunchkinPF=1 Then  ' Needed to prevent lighting Rainbow during Twister mode
        Select Case Switch
          Case 89, 90 ,91, 92, 93, 94, 95
            Exit Sub
        End Select
    End If
    If light_dict.Exists(switch) Then
        Set get_light = light_dict.Item(switch)
        If lcolor_dict.Exists(get_light) Then
            get_color = lcolor_dict.Item(get_light)
            SetLightColor get_light,get_color,1
        End If
    End If
End Sub

Sub SetLightsOFF(switch)
    Dim get_light
    If Tilted Then Exit Sub
    If light_dict.Exists(switch) Then
        Set get_light = light_dict.Item(switch)
        SetLightColor get_light,BaseColor(switch),0
    End If
End Sub

Sub SetLightsBlink(switch)
    Dim get_light
    If Tilted Then Exit Sub
    If light_dict.Exists(switch) Then
        Set get_light = light_dict.Item(switch)
        SetLightColor get_light,BaseColor(switch),2
    End If
End Sub

'*******************************************************************************
'                     SWITCH EVENT ACTIONS
'*******************************************************************************
Sub SetSwitchFlags(switch)
    Dim i, bg_id, cnt, color, g_award, wiz_award_cnt
    Dim tmp0, tmp1, tmp2, tmp3
    Dim obj, get_light

    SwitchFlags(switch) = SwitchFlags(switch) + 1
    Select Case switch
      case 3
        HOADC_Select
        YBR_ChooseShot
        If bHU=0 Then HU_Select
      case 4              '  Munchkin Playfield Entrance
        If bMultiBallMode=1 Then
            If bCastlePFActive=1 Then
                Multipliers(4) = 3
            Else
                If Multipliers(4)<2 Then Multipliers(4) = 2
            End If
            SuperXcolor Multipliers(4)*Multipliers(6)
        End If
        SetLightColor l96,"white",2
        bOnMunchkinPF=1
        ballsOnMunchkinPF = ballsOnMunchkinPF + 1
        If bTwisterFlag=1 AND bMunchkinMultiBall=0 Then
            TwisterModeHold.Enabled=False
            If bTwistAgain=0 Then
                AddScore_points 1000
                loopsleft = 10 + munchkinModeCnt*5
            End If
            ' Rainbow Switch Hit counts set to zero
            SwitchFlags(89)=0:SwitchFlags(90)=0:SwitchFlags(91)=0:SwitchFlags(92)=0
            SwitchFlags(93)=0:SwitchFlags(94)=0:SwitchFlags(95)=0
            For each obj in Lights_Rainbow
                SetLightColor obj, "white", 2
            Next
            bHouseSpin=1: HouseSpin.Enabled = True:  DOF 351, DOFOn      ' Gear Motor On
            Display_MX "TWISTER"
            PlaySong "voc20_Twister"
            If bMultiBallMode=0 Then GiOff
            If bCB_LightsOFF=0 AND bCB_LightsON=0 AND bMultiBallMode=0 Then
'                LightSeqTwister.Play SeqClockLeftOn,30,100
                LightSeqTwister.UpdateInterval=5
                LightSeqTwister.Play SeqScrewLeftOn,30,100
            End If
            bTwistAgain=1
            ClearBG_TRainbow
            Controller.B2SSetData 236,1
        End If
      case 5             '  Munchkin House Entrance
'        debug.print "SW5 - L137 State = " & l137.State
'        debug.print "SW5 - bECBallsLocked = " & bECBallsLocked
'        debug.print "SW5 - bECLock = " & bECLock
'        debug.print "Balls On Playfield = " & BallsOnPlayfield
        MHouse_BHandler
      case 9             ' Spinner entry to Crystal Ball VUK
        If bCrystalBallMode=0 AND bHauntedLtrMode=0 Then SwitchFlags(9) = 0
        If bRescueMB=1 Then RescueMB_Update switch
        If bFireballFrenzyStarted=1 AND ff_Lock=0 Then ff_Lock=1: FireballFrenzyUpdate l76
        If bMunchkinLollipop=1 Then Lollipop_Update switch
        If L84.State=1 Then HOADC_Hit switch
        If bHU=1 AND activeHU="CRYSTAL" Then HUDone
      case 15            ' Throne Room - Wizard
        If bHU=1 AND activeHU="WIZARD" Then HUDone
        If bECMultiB=1 Then ECMultiB_Throne
        If bRescueMB=1 Then RescueMB_Update switch
        If bFireballFrenzyStarted=1 AND ff_Lock=0 Then ff_Lock=1: FireballFrenzyUpdate l30
        If bMunchkinLollipop=1 Then Lollipop_Update switch
        If L29.state=2 Then AwardExtraBall:SetLightColor l29,"orange",0
        If L31.state=2 Then
           If bFreePlay Then
              AwardExtraBall
           Else
              AwardSpecial
           End If
           SetLightColor l31,"orange",0
        End If
        throneHits = throneHits + 1 : If throneHits>6 Then throneHits=6
        ' DOF PF Back Effects MX
        DOF 166, DOFOn:vpmTimer.AddTimer 5000,"DOF 166, DOFOff '"
        Select Case throneHits
          case 1
            w_sign.visible=1
            PlaySound "voc66_IamOZ"
          case 2
            w_sign.visible=1: i_sign.visible=1
            PlaySound "voc70_WhoAreYou"
          case 3
            w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
            PlaySound "voc74_WhyComeBack"
          case 4
            w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
            a_sign.visible=1
            PlaySound "voc75_NotSoFast1"
          case 5
            w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
            a_sign.visible=1:r_sign.visible=1
            PlaySound "voc76_NotSoFast2"
          case 6
            w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
            a_sign.visible=1: r_sign.visible=1: d_sign.visible=1
            vpmTimer.AddTimer 2000,"Wizard_reset '"
            PlaySound "voc23_comeforward"
            ' Can't award character lock or Emerald City Multiball while in
            ' Emerald City Multiball; however, these awards are removed during
            ' any multiball to limit complexity. Using Do while loops
            ' when asynchronous events come in rapid but intermittent bunches makes
            ' for a testing nightmare and strange lockups
            If bMultiBallMode=1 Then
                wiz_award_cnt=INT(RND * 5 + 1)
            Else
                wiz_award_cnt=INT(RND * 7 + 1)
            End If
            Select Case wiz_award_cnt
              case 1    ' Light Extra Ball
                SetLightColor l29,"orange",2
                tmp3 = "EXTRA BALL LIGHT"
              case 2    ' Award Points
                AddScore_points wiz_points
                wiz_points = wiz_points + 2500
                tmp3 = Cstr(wiz_points) & " POINTS"
              case 3    ' Light Special
                SetLightColor l31,"orange",2
                tmp3 = "  SPECIAL LIGHT"
              case 4    ' Add Rescue Lock
                If ballRescueLock < 2 Then ballRescueLock = ballRescueLock + 1
                PlaySound "voc62_Xylo"
                tmp3 = "  RESCUE LOCK"
              case 5    ' Add Tilt Warning
                Tilt = 6
                CheckTilt
                tmp3 = "  TILT WARNING"
              case 6    ' Light Lock for One Character
                If L23.state=1 Then
                   If L83.state=1 Then
                       If L10.state =1 Then
                           AddScore_points wiz_points
                       Else
                           Do While L10.state=0
                               SetSwitchFlags(96)
                           Loop
                       End If
                    Else
                        Do While L83.state=0
                            SetSwitchFlags(56)
                        Loop
                    End If
                Else
                   Do While L23.state=0
                       SetSwitchFlags(104)
                   Loop
                End If
                tmp3 = " CHARACTER LOCK"
              case 7    ' Start Emerald City MultiBall
                If Flashers(flash_dict.Item(l150))=0 Then sw6.createBall
                If Flashers(flash_dict.Item(l151))=0 Then sw7.createBall
                EmeraldCityMB
                tmp3 = "   MULTIBALL"
            End Select
            If DMD_Option=True Then
                DMD2Flush
                If (dqHead2 = dqTail2) Then
                    DMD2 "", eNone, " WIZARD AWARDS", eNone, tmp3, eNone, "", eNone, 2500, True, ""
                End If
            End If
            throneHits=0
        End Select
        If L35.State=1 Then HOADC_Hit switch
      case 16            ' Right Orbit Start
        If bRescueMB=1 Then RescueMB_Update switch
        If bFireballFrenzyStarted=1 AND ff_Lock=0 Then ff_Lock=1: FireballFrenzyUpdate l25
        If L27.State = 2 Then AdvanceYBR
      case 17            ' Castle Playfield VUK
        If bMultiBallMode=1 Then
            If bOnMunchkinPF=1 Then
                Multipliers(4) = 3
            Else
                If Multipliers(4)<2 Then Multipliers(4) = 2
            End If
            SuperXcolor Multipliers(4)*Multipliers(6)
        End If
        bCastlePFActive=1
        ballsOnCastlePF = ballsOnCastlePF + 1
        PlaySong "voc13_winkiemarch"
        If bRescueMB=1 Then RescueMB_Update switch
        If bFireballFrenzyStarted=1 AND ff_Lock=0 Then ff_Lock=1: FireballFrenzyUpdate l132
        If bRescueMB=0 Then
            Controller.B2SSetData 179,0
            Controller.B2SSetData 180,0
            Controller.B2SSetData 205,0
            Controller.B2SSetData 206,1
            For each obj in Lights_Castle
                bg_id = bg_dict.Item(obj)
                Controller.B2SSetData bg_id, obj.State
            Next
            If ballRescueLock>0 Then Controller.B2SSetData 241,1
            If ballRescueLock=2 Then Controller.B2SSetData 242,1
        End If
        If L106.State=1 AND L107.State=1 AND L108.State=1 AND L111.State=1 AND L112.State=1 AND L113.State=1 Then
             If bDorothyCaught=1 Then SetLightColor l110,"yellow",2
             SetLightColor l109,"yellow",2
        End If
        If L136.State=1 Then HOADC_Hit(switch)
      case 18            ' Castle Playfield Kicker Behind Doors
        If bRescueMB=1 Then RescueMB_Jackpot: Exit Sub
        cnt=0
        For each obj in Lights_RESCUE
            If obj.State=0 Then
                If cnt<2 Then
                    SetLightColor obj,"orange",1
                    cnt=cnt+1
                    bg_id = bg_dict.Item(obj)
                    Select Case bg_id
                      case 207  ' R
                        vpmtimer.addtimer 150, "SetSwitchFlags(67) '"
                      case 208  ' E
                        vpmtimer.addtimer 150, "SetSwitchFlags(66) '"
                      case 209  ' S
                        vpmtimer.addtimer 150, "SetSwitchFlags(65) '"
                      case 210  ' C
                        vpmtimer.addtimer 150, "SetSwitchFlags(68) '"
                      case 211  ' U
                        vpmtimer.addtimer 150, "SetSwitchFlags(69) '"
                      case 212  ' E
                        vpmtimer.addtimer 150, "SetSwitchFlags(70) '"
                    End Select
                End If
            End If
        Next
        If cnt=0 AND bDorothyCaught=1 Then
            bDorothyCaught = 0: bDorothyCaptureRdy = 0
            LightState(153)=0:Flashers(flash_dict.Item(l153))=0
            SwitchFlags(61) = 0:SwitchFlags(62) = 0
            Wall8.IsDropped = 1
            Wall71.IsDropped=0

            castleDoorBashCnt=0
            ClearBG_Rescue
            Controller.B2SSetData 243,1
            bBTWW_2 = 1
            Controller.B2sSetData "Quadrant_Lt2",1
            sw72_out.Timerenabled = 1
'            sw72_out.Kick 180,15
            BallsInLock=BallsInLock-1
            BallsOnPlayfield=BallsOnPlayfield+1
            ballsOnCastlePF = ballsOnCastlePF + 1
            RescueMBClock.enabled = 1
            If ballRescueLock>0 Then
                AddMultiball ballRescueLock
            Else
                DOF 112, DOFPulse	' Strobe
'		        DOF 126, DOFPulse	' Beacon for multiball
            End If
            bMultiBallMode=1
            Display_MX "RESCUE"
            If PUPpack=True Then
                DOF 80, DOFOn
                vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
                DOF 57, DOFOn
                vpmTimer.AddTimer 3000,"DOF 57, DOFOff '"
            End If
            ChangeGi "blue"
            bRescueMB=1
            ballRescueLock=0
            For i= 65 To 70
                SwitchFlags(i)=0
                SetLightsOFF i
            Next
        End If
      case 20            ' Leaving Castle Playfield
        ballsOnCastlePF = ballsOnCastlePF - 1
        If ballsOnCastlePF<=0 Then
'            sw73.IsDropped = 0: SwitchFlags(73)=0:SetLightsOff(73):Controller.B2SSetData 179,0
            SwitchFlags(71)=0
            bCastlePFActive=0
            If bOnMunchkinPF=1 Then
                Multipliers(4) = 2
            Else
                Multipliers(4) = 1
            End If
            SuperXcolor Multipliers(4)*Multipliers(6)
            PlaySong "voc18_OffToSee"
            If bCastleDoorOpen Then
                CastleDoorsClose
                PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
                bCastleDoorOpen=0
                Wall8.IsDropped=0
            End If
            SetLightColor L110,"yellow",0
            SetLightColor l109,"yellow",0
            ClearBG_Rescue
            RestoreBG_Castle
            ballsOnCastlePF=0
        End If
      case 21            ' Left Side VUK (Crystal Ball & Haunted)
        If bCrystalBallMode=1 AND SwitchFlags(9)>0 AND  _
           bCB_FlipFrenzy=0 AND bCB_LightsOFF=0 AND bCB_LightsON=0 AND bCB_NoHoldFlip=0 AND bCB_LinkedFlip=0 Then
             PlaySong "voc9_Gulch"
             Display_MX "CRYSTAL"
             Select Case crystalBallModeCnt
              case 1  ' Lights OFF  - 2x
                CB_LightsOFF
              case 2  ' Lights On - 2x
                CB_LightsON
              case 3  ' Linked Flippers - 2x
                CB_LinkedFlip
              case 4  ' No Hold Flippers - 2x
                CB_NoHoldFlip
              case 5  ' Flipper Frenzy - 3x
                crystalBallModeCnt=0
                CB_FlipFrenzy
            End Select
            crystalBallModeCnt = crystalBallModeCnt + 1
        Else
            If bHauntedLtrMode=1 AND SwitchFlags(9)=0 Then
                HauntedModes
            Else
                PlaySound "voc55_CB_VUK"
            End If
            If bHauntedMB=1 AND l142.State=1 Then
               hauntedMB_Bonus = hauntedMB_Bonus + hauntedShots + hauntedHoles + hauntedTargets + hauntedBumpers
               SetLightColor L97,"purple",0
               SetLightColor l142,"red",0
            End If
        End If
        SwitchFlags(9)=0
      case 24            ' Left Orbit
        If bRescueMB=1 Then RescueMB_Update switch
        If L100.State = 2 Then AdvanceYBR
      case 27            ' Left InLane
        If L68.State=2 OR L69.State=2 OR L70.State=2 Then HurryUp.Enabled = True
      case 28            ' TNPLH Outlane
        HOADC_BG_clr
        If PUPpack=True AND bMultiBallMode=0 Then
            DOF 28,DOFPulse
        Else
            PlaySound "voc12_theresnoplace"
        End If
        Controller.B2SSetData "TNPLHS0_0",1
        If bMultiBallMode=1 Then
            vpmtimer.addtimer 2000, "Controller.B2sSetData ""TNPLHS0_0"",0 '"
        Else
            vpmtimer.addtimer 4000, "Controller.B2sSetData ""TNPLHS0_0"",0 '"
        End If

      case 29,30,31,32    ' Crystal Ball - BALL
        If light_dict.Exists(switch) Then
            Set get_light = light_dict.Item(switch)
            If bSkillshotReady=1 AND get_light.State=2 Then SkillShotMade CRYSTAL_SKILLSHOT
            Select Case crystalBallModeCnt
              case 1
                cb_Color = "blue"
              case 2
                cb_Color = "yellow"
              case 3
                cb_Color = "green"
              case 4
                cb_Color = "red"
              case 5
               cb_Color = "purple"
            End Select
            If bCrystalBallMode=0 Then SetLightColor get_light,cb_Color,1
        End If
        If bCrystalBallMode=0 AND l72.State=1 AND l73.State=1 AND l74.State=1 AND l75.State=1 Then CrystalBallMode

      case 35            ' Right InLane
        If L32.State=2 OR L33.State=2 OR L34.State=2 Then HurryUp.Enabled = True
      case 49, 50, 51    ' Haunted - Bumpers
        If bFireballFrenzyStarted=1 AND ff_Lock=0 Then ff_Lock=1: FireballFrenzyUpdate l80
        If bHauntedLtrMode=0 AND bHauntedMode=0 Then HauntedLights
        If bHauntedMode=1 Then
            If SwitchFlags(49)+SwitchFlags(50)+SwitchFlags(51)>6 Then
                 If bHauntedMB=1 AND l97.State=0 Then
                     SetLightColor L97,"purple",2
                     SetLightColor l142,"red",1
                 Else
                     Controller.B2SSetData H_Tens,0:Controller.B2SSetData H_Ones,0
                     hauntedTimeLeft=hauntedTimeLeft+10
                 End If
                 SwitchFlags(49)=0:SwitchFlags(50)=0:SwitchFlags(51)=0
            End If
        End If
      case 55            ' Haunted Forest Skill Shot Target
        If bSkillshotReady=1 Then SkillShotMade HAUNTED_SKILLSHOT
      case 56            ' Tin Man Rollover
        If bECMultiB=1 Then
            Select Case qualTinman
              Case 0
                qualTinman=1
                If qualLion=1 Then qualLion=0
                If qualScarecrow=1 Then qualScarecrow=0
                AddScore_points 2500
              Case 3
                qualTinman=4
                AddScore_points 2500
            End Select
            Exit Sub
        End If
        cnt_Tinman = cnt_Tinman + 1
        If cnt_Tinman<6 Then
            SetLightsBlink(Switch)
        Else
            SetLights(Switch)
        End If
        Select Case cnt_Tinman
          Case 0
            ' Covers the unexpected
          Case 1
            SetSwitchFlags 119
            SetLights 119
          Case 2
            SetSwitchFlags 120
            SetLights 120 :SetLights 119
          Case 3
            SetSwitchFlags 121
            SetLights 121 :SetLights 120 :SetLights 119
          Case 4
            SetSwitchFlags 122
            SetLights 122 :SetLights 121 :SetLights 120
            SetLights 119
          Case 5
            SetSwitchFlags 123
            SetLights 123 :SetLights 122 :SetLights 121
            SetLights 120 :SetLights 119
          Case 6
            SetSwitchFlags 124
            SetLightColor l137,"green",1
            SetLights 124 :SetLights 123 :SetLights 122
            SetLights 121 :SetLights 120 :SetLights 119
            ELockProcess
          Case Else
            SetLights 119 :SetLights 120 :SetLights 121
            SetLights 122 :SetLights 123 :SetLights 124
        End Select
        For each obj in Lights_TinMan
            bg_id = bg_dict.Item(obj)
            Controller.B2SSetData bg_id, obj.State
        Next
'        debug.print "cnt_Tinman = " & cnt_Tinman
      case 57, 58            ' Witch Switches
        If bMultiBallMode=0 AND bFireballFrenzy=0 AND bMunchkinMode=0 _
          AND bBTWW_1=1 AND bBTWW_2=1 AND bBTWW_3=1 AND bBTWW_4=1 Then BTWW: Exit Sub
        If bFireballFrenzy=1 Then FireballFrenzy: Exit Sub
        witchHitCnt = witchHitCnt + 1
        If bWitchHU=0 AND bSkillshotReady=1 Then
            SkillShotMade WITCH_SKILLSHOT
        Else
            If bWitchHU=0 AND bWitchDown=1 AND witchHitCnt=>WitchHurryUpStartHits Then
                WitchRaise
                bWitchHU=1: witchHitCnt = WitchHurryUpStartHits
                SetLightColor L131,"orange",1
                Controller.B2SSetData 189,1
                WitchHurryUp.Enabled = True
                SetLightColor l8,"red",2
            Else
                If bWitchHU=1 AND bWitchDown=0 AND witchHitCnt=>WitchHurryUpStartHits+WitchHurryUpStopHits Then
                    AddScore_points witchHUPot
                    blueShots = blueShots + witchHUPot
                    If DMD_Option=True Then
                        DMD2Flush
                        If(dqHead2 = dqTail2)Then
                            tmp0 = ""
                            tmp1 = " WITCH HURRY UP"
                            tmp2 = FillLine(2, "POT SCORE", FormatScore(witchHUPot))
                            tmp3 = ""
                            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
                        End If
                        vpmTimer.AddTimer 2500,"DMD2_Clear '"
                    End If
                    WitchHUCnt=WitchHUCnt+1
                    If witchHUCnt=1 Then PlaySound "voc10_illbeback"
                    If witchHUCnt=2 Then PlaySound "voc68_WatchingYouChild"
                    If WitchHUCnt=3 Then
                        PlaySound "voc69_LuckyShot"
                        SetLightColor L130,"orange",2:SetLightColor L131,"red",2:SetLightColor L133,"red",2
                        bFireballFrenzy = 1
                        WitchHUCnt = 0
                        bWitchHU = 0
                    End If
                    WitchHULow=WitchHULow+witchHUPot
                    If WitchHULow>4500 Then WitchHULow=4500
                    WitchHUStop
                End If
            End If
        End If
      case 61    ' Winged Monkey
        If SwitchFlags(62)=1 AND bDorothyCaptureRdy=0 Then
            LightState(115)=1:Flashers(flash_dict.Item(l115))=1:Kicker_Mag.Enabled=True
            bDorothyCaptureRdy=1
        End If
      case 62    ' Winged Monkey
        If SwitchFlags(61)=1 AND bDorothyCaptureRdy=0 Then
            LightState(115)=1:Flashers(flash_dict.Item(l115))=1:Kicker_Mag.Enabled=True
            bDorothyCaptureRdy=1
        End If
     case 64
        If bFireballFrenzyStarted=1 AND ff_Lock=0 Then ff_Lock=1: FireballFrenzyUpdate l80
        If bMunchkinLollipop=1 Then Lollipop_Update switch
        If L99.State=1 Then HOADC_Hit(switch)
        If bHU=1 AND activeHU="OZ LEFT" Then HUDone
      case 65, 66, 67, 68, 69, 70    ' Castle PF Target Switches
        If bCastlePFActive=1 Then
            Set obj = light_dict.Item(switch)
            bg_id = bg_dict.Item(obj)
            Controller.B2SSetData bg_id,1
        End If
        If L106.State=1 AND L107.State=1 AND L108.State=1 AND L111.State=1 AND L112.State=1 AND L113.State=1 Then
             If bDorothyCaught=1 Then SetLightColor l110,"yellow",2
             SetLightColor l109,"yellow",2
        End If
      case 71    ' Castle PF Loop
        If bRescueMB=1 Then Exit Sub
        If bCastlePFActive=1 Then
            If L109.State=0 Then      ' Spot a RESCUE light
                For i = 65 To 70
                    If SwitchFlags(i)=0 Then
                        SwitchHitEvent i
                        Exit For
                    End If
                Next
            Else                      ' Add a locked ball
                If ballRescueLock < 2 Then ballRescueLock = ballRescueLock + 1: PlaySound "voc62_Xylo"
                If ballRescueLock > 0 Then Controller.B2SSetData 241,1
                If ballRescueLock > 1 Then Controller.B2SSetData 242,1
            End If
        End If
      case 72    ' Dorothy Captured
        bDorothyCaught = 1
        Wall71.IsDropped=1
        SwitchFlags(61) = 0:SwitchFlags(62) = 0
        SetLightColor L101,"blue",0:SetLightColor L102,"blue",0
        Kicker_Mag.Enabled = False
        LightState(115)=0: Flashers(flash_dict.Item(l115))=0
        LightState(153)=1: Flashers(flash_dict.Item(l153))=1
        Controller.B2SSetData 180,1
        If L106.State=1 AND L107.State=1 AND L108.State=1 AND L111.State=1 AND L112.State=1 AND L113.State=1 Then
             SetLightColor l110,"yellow",2
             SetLightColor l109,"yellow",2
        End If
      case 73    ' Winkie Guard Down
        SetLights(Switch)
        If bHU=1 AND activeHU="WINKIE" Then HUDone
        If bRescueMB=1 Then
            RescueMB_Update switch
        Else
            Controller.B2SSetData 179,1
            If PUPpack=True AND bMultiBallMode=0 AND bWitchHU=0 Then DOF 325,DOFPulse
        End If
        If bFireballFrenzyStarted=1 AND ff_Lock=0 Then ff_Lock=1: FireballFrenzyUpdate l132
        If bMunchkinLollipop=1 Then Lollipop_Update switch
        If L136.State=1 Then HOADC_Hit(switch)
      case 74    ' Glinda Target
        If L135.State=0 Then Exit Sub
        Controller.B2SSetData "Glinda",1
        vpmTimer.AddTimer 2000,"Controller.B2SSetData ""Glinda"",0 '"
        PlaySound "voc11_witchofthenorth"
        L135.State=0
        Display_MX "GLINDA"
        ' Random Selection of mode check sequence
        Dim rnd_seq(4), rnd_index
        For i = 0 To 4: rnd_seq(i) = i: Next
        For i = 4 To 1 step -1
            rnd_index = INT(RND*i + 1)
            Select Case rnd_seq(rnd_index)
              case 1
                If bCB_LightsON OR bCB_LightsOFF Then
                    LightSeqInserts.StopPlay
                    ChangeGi "pink"
                    If UltraDMD_Option=True Then UltraDMD_simple "Thwart_Witch.png", " ", "", 2000
                    If DMD_Option=True Then
                      DMD2Flush
                        If (dqHead2 = dqTail2) Then
                            DMD2 "", eNone, "  THWART WITCH", eNone, "  LIGHTS FIXED", eNone, "", eNone, 2500, True, ""
                        End If
                    End If
                    Exit Sub
                End If
                If bCB_NoHoldFlip OR bCB_LinkedFlip OR bMultiBallMode Then
                    If UltraDMD_Option=True Then UltraDMD_simple "Thwart_Witch.png", " ", "", 2000
                    AddMultiball 1
                    Exit Sub
                End If
              case 2
                If bWitchHU Then
                    If UltraDMD_Option=True Then UltraDMD_simple "Thwart_Witch.png", " ", "", 2000
                    If DMD_Option=True Then
                      DMD2Flush
                        If (dqHead2 = dqTail2) Then
                            DMD2 "", eNone, "  THWART WITCH", eNone, "    ADD TIME", eNone, "", eNone, 2500, True, ""
                        End If
                    End If
                    witchHUPot = witchHULow + 1500
                    Controller.B2SSetData Ones,0
                    Controller.B2SSetData Tens,0
                    Controller.B2SSetData Hundreds,0
                    Controller.B2SSetData Thousands,0
                    BG_HUScore
                    Exit Sub
                End If
              case 3
                If bMunchkinLand OR bMunchkinFrenzy OR bMunchkinLollipop Then
                    If UltraDMD_Option=True Then UltraDMD_simple "Add_Time.png", " ", "", 2000
                    Controller.B2SSetData T_Ones,0:Controller.B2SSetData T_Tens,0
                    g_award = INT(RND * 10 + 1)
                    munchkinTimeLeft = munchkinTimeLeft + 25 + g_award
                    If munchkinTimeLeft>99 Then munchkinTimeLeft = 99
                    Exit Sub
                End If
              case 4
                If bHauntedBumpers OR bHauntedHoles OR bHauntedShots OR bHauntedTargets Then
                   If UltraDMD_Option=True Then UltraDMD_simple "Add_Time.png", " ", "", 2000
                    Controller.B2SSetData H_Ones,0:Controller.B2SSetData H_Tens,0
                    g_award = INT(RND * 10 + 1)
                    hauntedTimeLeft = hauntedTimeLeft + 15 + g_award
                    If hauntedTimeLeft>99 Then hauntedTimeLeft = 99
                    Exit Sub
                End If
            End Select
            If rnd_index<i Then rnd_seq(rnd_index) = rnd_seq(i)
        Next
        ' No Mode Active
        cnt = 0
        g_award = INT(RND * 5 + 1)
        Select Case g_award
          Case 1    ' Spot up to two Rainbow Lights
            If UltraDMD_Option=True Then UltraDMD_simple "Spot_Lights.png", " ", "", 2000
            cnt = INT(RND + 0.5)
            For i = 89 To 95
                If SwitchFlags(i)=0 Then
                    cnt = cnt + 1
                    SwitchHitEvent i
                End If
                If cnt>1 Then Exit For
            Next
          Case 2
            AddScore_points 2000
            If UltraDMD_Option=True Then UltraDMD_simple "Add_Points.png", " ", "", 2000
          Case 3    ' Spot up to two Crystal Ball Lights
            If UltraDMD_Option=True Then UltraDMD_simple "Spot_Lights.png", " ", "", 2000
            cnt = INT(RND + 0.5)
            For i = 29 To 32
                If SwitchFlags(i)=0 Then
                    cnt = cnt + 1
                    SwitchHitEvent i
                End If
                If cnt>1 Then Exit For
            Next
          Case 4
            AddScore_points 3000
            If UltraDMD_Option=True Then UltraDMD_simple "Add_Points.png", " ", "", 2000
          Case 5
            HOADC_Select
         End Select
      case 75
        If L138.State=1 Then HOADC_Hit(switch)
        If l143.State=2 Then AdvanceYBR
      case 76    ' Upper Ramp Made
        If bECMultiB=1 Then ECMultiB_Ramp
        If bRescueMB=1 Then RescueMB_Update switch
        If bFireballFrenzyStarted=1 AND ff_Lock=0 Then ff_Lock=1: FireballFrenzyUpdate l139
        If bMunchkinLollipop=1 Then Lollipop_Update switch
      case 77    ' Castle Door Bash
         castleDoorBashCnt = castleDoorBashCnt + 1
         If castleDoorBashCnt=>CastleDoorHits AND bCastleDoorOpen=0 Then
             CastleDoorsOpen
             PlaySound SoundFXDOF("flipperup",102,DOFOn,DOFContactors), 0, .67, 0.05, 0.05
             bCastleDoorOpen = 1
         End If
      case 81    ' OZ Left
        If bOZLeft=1 AND bSkillshotReady=1 Then SkillShotMade OZ_SKILLSHOT: Exit Sub
        PlaySound "voc81_fastbeat"
        If l253.state=1 Then
            Multipliers(1)=Multipliers(1)+1: SwitchFlags(81)=0:SwitchFlags(82)=0
            If Multipliers(1)>10 Then Multipliers(1)=10
            If UltraDMD_Option=True Then UltraDMD_simple "black.png", "OZ MULTIPLIER", "X " &Cstr(Multipliers(1)), 1500
            SetLightColor l252,"white",0:LightState(104)=0:Flashers(flash_dict.Item(l104))=0
            SetLightColor l253,"white",0:SetLightColor l254,"white",0:LightState(105)=0:Flashers(flash_dict.Item(l105))=0
            If DMD_Option=True Then OZ_DMD2
        Else
            SetLightColor l252, "white", 1:LightState(104)=1:Flashers(flash_dict.Item(l104))=1
        End If
      case 82    ' OZ Right
        If bOZRight=1 AND bSkillshotReady=1 Then SkillShotMade OZ_SKILLSHOT: Exit Sub
        PlaySound "voc81_fastbeat"
        If l252.state=1 Then
            Multipliers(1)=Multipliers(1)+1: SwitchFlags(81)=0:SwitchFlags(82)=0
            If Multipliers(1)>10 Then Multipliers(1)=10
            If UltraDMD_Option=True Then UltraDMD_simple "black.png", "OZ MULTIPLIER", "X " &Cstr(Multipliers(1)), 1500
            SetLightColor l252,"white",0:LightState(104)=0:Flashers(flash_dict.Item(l104))=0
            SetLightColor l253,"white",0:SetLightColor l254,"white",0:LightState(105)=0:Flashers(flash_dict.Item(l105))=0
            If DMD_Option=True Then OZ_DMD2
        Else
            SetLightColor l253, "white", 1:SetLightColor l254, "white", 1
            LightState(105)=1:Flashers(flash_dict.Item(l105))=1
        End If
      case 87
        If bFireballFrenzyStarted=1 AND ff_Lock=0 Then ff_Lock=1: FireballFrenzyUpdate l25
        If bHU=1 AND activeHU="OZ RIGHT" Then HUDone
        If bMunchkinLollipop=1 Then Lollipop_Update switch
        If L26.State=1 Then HOADC_Hit(switch)
      case 89, 90, 91, 92, 93, 94, 95   ' Rainbow targets
        set obj = light_dict.Item(Switch)
        bg_id = bg_dict.Item(obj)
        If bMunchkinMode=0 Then
            If bTwisterFlag=0 Then Controller.B2SSetData bg_id,1
            If bMultiBallMode=0 Then Display_MX "RAINBOW"
            If L1.State=1 AND L2.State=1 AND L3.State=1 AND L4.State=1 AND L5.State=1 AND L6.State=1 AND L7.State=1 Then
                bTwisterFlag=1
                If L137.State=0 Then
                    DivRight.RotateToEnd: SetLightColor L134,"yellow",2:Controller.B2SSetData 233,1:bTwistAgain=0
                End If
            End If
        Else
            SetLightColor obj,"white",1
        End If
        If bHU=1 AND activeHU="RAINBOW" Then HUDone
        If bMunchkinLollipop=1 Then Lollipop_Update switch
      case 96             ' Scarecrow Rollover
        If bECMultiB=1 Then
            Select Case qualScarecrow
              Case 0
                qualScarecrow=1
                If qualTinman=1 Then qualTinman=0
                If qualLion=1 Then qualLion=0
                AddScore_points 2500
              Case 3
                qualScarecrow=4
                AddScore_points 2500
            End Select
            Exit Sub
        End If
        cnt_Scarecrow = cnt_Scarecrow + 1
        If cnt_Scarecrow<9 Then
            SetLightsBlink(Switch)
        Else
            SetLights(Switch)
        End If
        Select Case cnt_Scarecrow
          Case 0
            ' Covers the unexpected
          Case 1
            SetSwitchFlags 118
            SetLights 118
          Case 2
            SetSwitchFlags 117
            SetLights 118 :SetLights 117
          Case 3
            SetSwitchFlags 116
            SetLights 118 :SetLights 117  :SetLights 116
          Case 4
            SetSwitchFlags 115
            SetLights 118 :SetLights 117  :SetLights 116
            SetLights 115
          Case 5
            SetSwitchFlags 114
            SetLights 118 :SetLights 117  :SetLights 116
            SetLights 115 :SetLights 114
          Case 6
            SetSwitchFlags 113
            SetLights 118 :SetLights 117  :SetLights 116
            SetLights 115 :SetLights 114  :SetLights 113
          Case 7
            SetSwitchFlags 112
            SetLights 118 :SetLights 117  :SetLights 116
            SetLights 115 :SetLights 114  :SetLights 113
            SetLights 112
          Case 8
            SetSwitchFlags 111
            SetLights 118 :SetLights 117  :SetLights 116
            SetLights 115 :SetLights 114  :SetLights 113
            SetLights 112 :SetLights 111
          Case 9
            SetSwitchFlags 110
            SetLightColor l137,"green",1
            SetLights 118 :SetLights 117  :SetLights 116
            SetLights 115 :SetLights 114  :SetLights 113
            SetLights 112 :SetLights 111  :SetLights 110
            ELockProcess
          Case Else
            SetLights 118 :SetLights 117  :SetLights 116
            SetLights 115 :SetLights 114  :SetLights 113
            SetLights 112 :SetLights 111  :SetLights 110
        End Select
        For each obj in Lights_Scarecrow
            bg_id = bg_dict.Item(obj)
            Controller.B2SSetData bg_id, obj.State
        Next
'        debug.print "cnt_Scarecrow = " & cnt_Scarecrow
      case 98             ' Collect Trigger
        If L28.State=2 AND bHOADC_Score=1 Then HOADC_Score
        bHOADC_Score=0
      case 101, 102             ' Munchkin PF Upper & Lower - Increments Bounus
        If bMunchkinMultiBall=1 Then
            If munchkinMB_Rcvd=0 Then munchkinMB_Bonus = 2*munchkinMB_Bonus: munchkinMB_Rcvd=1
            SetLightColor l134,"yellow",0
            DivRight.RotateToStart
            Exit Sub
        End If
        twisterbonus = twisterbonus + twisterloops*BasePoints(switch)

        ' TESTING - Stops excessive twisterloops - auto mode sometimes exceeds 100 loops
        If bAutoFlip=1 AND twisterloops>31 Then Auto_Regression_Flip: vpmTimer.AddTimer 1000,"Auto_Regression_Flip '"

        ' Clear previous counts
        Controller.B2SSetData L_Ones,0:Controller.B2SSetData L_Tens,0
        Controller.B2SSetData Y_Ones,0:Controller.B2SSetData Y_Tens,0
        ' Loops related text
        Controller.B2SSetData 236,2
        If twisterloops=1 Then Controller.B2SSetData 235,1 Else Controller.B2SSetData 235,2

        If twisterloops<99 Then twisterloops = twisterloops + 1
        If loopsleft>0 Then loopsleft=loopsleft - 1
        If loopsleft=0 Then bMunchkinMode=1
        If bPreviousPosted=1 Then Display_Loops
      case 104             ' Lion Rollover
        If bECMultiB=1 Then
            Select Case qualLion
              Case 0
                qualLion=1
                If qualTinman=1 Then qualTinman=0
                If qualScarecrow=1 Then qualScarecrow=0
                AddScore_points 2500
              Case 3
                qualLion=4
                AddScore_points 2500
            End Select
            Exit Sub
        End If
        cnt_Lion = cnt_Lion + 1
        If cnt_Lion<4 Then
            SetLightsBlink(Switch)
        Else
            SetLights(Switch)
        End If
        Select Case cnt_Lion
          Case 0
            ' Covers the unexpected
          Case 1
            SetSwitchFlags 106
            SetLights 106
          Case 2
            SetSwitchFlags 107
            SetLights 106 :SetLights 107
          Case 3
            SetSwitchFlags 108
            SetLights 106 :SetLights 107 :SetLights 108
          Case 4
            SetSwitchFlags 109
            SetLightColor l137,"green",1
            SetLights 106 :SetLights 107 :SetLights 108
            SetLights 109
            ELockProcess
          Case Else
            SetLights 106 :SetLights 107 :SetLights 108
            SetLights 109
        End Select
        For each obj in Lights_Lion
            bg_id = bg_dict.Item(obj)
            Controller.B2SSetData bg_id, obj.State
        Next
'        debug.print "cnt_Lion = " & cnt_Lion
      Case 125
        vpmTimer.AddTimer 2000,"PlaySong ""voc18_OffToSee"" '"
        If(bBallSaverReady = True) AND (BallSaverTime <> 0) AND (bBallSaverActive = False) Then
            EnableBallSaver BallSaverTime
        End If
      Case 126
        bHOADC_Score=1
    End Select
    If bSkillshotReady=1 AND switch<>125 AND switch<>98 AND switch<>87 AND switch<>72 AND switch<>16 Then
        SkillShotStop
    End If
    LastSwitchHit = switch
End Sub

Dim bPreviousPosted

Sub Display_Loops
    bPreviousPosted=0
    L_Ones = "L1_" & CStr(twisterloops MOD 10):Controller.B2SSetData L_Ones,1
    If twisterloops>9 Then L_Tens = "Ld_" & CStr(INT(twisterloops/10) MOD 10):Controller.B2SSetData L_Tens,1
    Controller.B2SSetData Y_Ones,0:Controller.B2SSetData Y_Tens,0
    Y_Ones = "Y1_" & CStr(loopsleft MOD 10):Controller.B2SSetData Y_Ones,1
    If loopsleft>9 Then Y_Tens = "Yd_" & CStr(INT(loopsleft/10) MOD 10):Controller.B2SSetData Y_Tens,1
    vpmTimer.AddTimer 100,"bPreviousPosted=1 '"
End Sub

'*******************************************************************************
'                SKILL SHOT
'*******************************************************************************
Const OZ_SKILLSHOT=1
Const WITCH_SKILLSHOT= 2
Const HAUNTED_SKILLSHOT=3
Const CRYSTAL_SKILLSHOT=4

Dim skillShot_TimeLeft

Sub ChooseSkillShot
    Dim obj, cb_State
    'Choose Skill Shots
    ' OZ Lanes
    Select Case INT(RND * 2 + 1)
      case 1
        SetLightColor l252, "white", 2
        LightState(104)=1:LightState(105)=0:bOZLeft=1:bOZRight=0
        Flashers(flash_dict.Item(l104))=1:Flashers(flash_dict.Item(l105))=0
      case 2
        SetLightColor l253, "white", 2:SetLightColor l254, "white", 2
        LightState(104)=0:LightState(105)=1:bOZLeft=0:bOZRight=1
        Flashers(flash_dict.Item(l104))=0:Flashers(flash_dict.Item(l105))=1
    End Select
    ' Witch
    Smoke.Visible=1
    SetLightColor w1,"red",2
    SetLightColor w2,"red",2
    SetLightColor w3,"red",2
    SetLightColor w4,"red",2
'    WitchRaise
    SetLightColor L131,"orange",2
    ' Haunted Forest Target
    SetLightColor L19,"white",2:SetLightColor L90,"white",2
    ' Crystal Ball lights
    Select Case crystalBallModeCnt
      case 1
        cb_Color = "blue"
      case 2
        cb_Color = "yellow"
      case 3
        cb_Color = "green"
      case 4
        cb_Color = "red"
      case 5
        cb_Color = "purple"
    End Select
    For each obj in Lights_CrystalBall
        cb_State = obj.State
        SetLightColor obj, cb_Color, cb_State
    Next
    Select Case INT(RND * 4 + 1)
      case 1
        bBLit=L72.State
        SetLightColor L72,cb_Color,2
      case 2
        bALit=L73.State
        SetLightColor L73,cb_Color,2
      case 3
        bL1Lit=L74.State
        SetLightColor L74,cb_Color,2
      case 4
        bL2Lit=L75.State
        SetLightColor L75,cb_Color,2
    End Select
    skillShot_TimeLeft=30
    SkillShotTimer.Interval=1000
    SkillShotTimer.Enabled=1
    SkillShotDisplay
End Sub

Sub SkillShotDisplay
    Dim tmp0, tmp1, tmp2, tmp3
    If DMD_Option=True Then
        DMD2Flush
        if(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "SHOOT SKILL SHOT"
            tmp2 = FillLine(2, "TIME LEFT", FormatScore(skillShot_TimeLeft))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eBlink, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
    End If
End Sub

Sub SkillShotClear
    If DMD_Option=True Then DMD2_Clear
End Sub

Sub SkillShotMade(skillshot_nbr)
    Dim tmp0, tmp1, tmp2, tmp3
    Dim i
    SkillShotTimer.Enabled = 0
    bSkillShotReady = 0
    PlaySound "voc83_BeautifulShot"
    If UltraDMD_Option=True Then UltraDMD_simple "Skill_Shot.png", "SKILL SHOT MADE", _
       "SCORE: " &Cstr(SkillshotValue(skillshot_nbr)), 1500
    If DMD_Option=True Then
        DMD2Flush
        If(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "SKILL SHOT MADE"
            tmp2 = FillLine(2, "SCORE", FormatScore(SkillshotValue(skillshot_nbr)))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2500, True, ""
        End If
    End If
    AddScore_points SkillshotValue(skillshot_nbr)
    SkillshotValue(skillshot_nbr) = SkillshotValue(skillshot_nbr) + SkillshotIncrease(skillshot_nbr)
    ' Clear Skill Shot Not Made
    If skillshot_nbr<>OZ_SKILLSHOT Then
        LightState(104)=0:LightState(105)=0
        Flashers(flash_dict.Item(l104))=0:Flashers(flash_dict.Item(l105))=0
        SetLightColor l252, "white", 0:SetLightColor l253, "white", 0:SetLightColor l254, "white", 0
    End If
    If skillshot_nbr<>WITCH_SKILLSHOT Then SetLightColor L131,"orange",0: WitchLower
    If skillshot_nbr<>HAUNTED_SKILLSHOT Then SetLightColor L19,"white",0:SetLightColor L90,"white",0
    If skillshot_nbr<>CRYSTAL_SKILLSHOT AND bCrystalBallMode=0 Then
        If L72.State=2 Then SetLightColor L72,"blue",bBLit: bBLit=0
        If L73.State=2 Then SetLightColor L73,"blue",bALit: bALit=0
        If L74.State=2 Then SetLightColor L74,"blue",bL1Lit:bL1Lit=0
        If L75.State=2 Then SetLightColor L75,"blue",bL2Lit:bL2Lit=0
    End If
    ' Specific Skill Shot Made Awards
    Select Case skillshot_nbr
      case OZ_SKILLSHOT    ' OZ Lane Skill Shot
        If bOZLeft=1 Then
           SetLightColor l252,"white",1
           LightState(104)=1:Flashers(flash_dict.Item(l104))=1
        End If
        If bOZRight=1 Then
           SetLightColor l253,"white",1:SetLightColor l254,"white",1
           LightState(105)=1:Flashers(flash_dict.Item(l105))=1
        End If
      case WITCH_SKILLSHOT    ' Witch Skill Shot
        If bFireballFrenzyStarted=0 Then vpmTimer.AddTimer 2500,"SkillShotWitchHU '"
      case HAUNTED_SKILLSHOT    ' Haunted Forest Target Skill Shot
        If bHauntedMode=0 Then bHauntedLtrMode=1:SwitchFlags(9)=0:SetSwitchFlags 21
      case CRYSTAL_SKILLSHOT    ' Crystal Ball Skill Shot
        For i = 0 to 3:SwitchHitEvent (29+i):Next  ' Crystal Ball Targets
    End Select
    vpmTimer.AddTimer 3500,"SkillShotClear '"
End Sub

Sub SkillShotWitchHU
    WitchRaise
    bWitchHU=1: witchHitCnt = 2
    SetLightColor L131,"orange",1
    Controller.B2SSetData 189,1
    WitchHurryUp.Enabled = True
    SetLightColor l8,"red",2
End Sub

'  Skill Shot Timer
Sub SkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    skillShot_TimeLeft=skillShot_TimeLeft-1
    If skillShot_TimeLeft>0 Then
        SkillShotDisplay
    Else
        SkillShotStop
        SkillShotClear
        Me.Enabled=0
    End If
End Sub

Sub SkillShotStop
    SkillShotTimer.Enabled = 0
    bSkillShotReady = 0
    ' Turn Off Skill Shots
    ' OZ lanes
    If bOZLeft=1 Then
        LightState(104)=0:Flashers(flash_dict.Item(l104))=0:SetLightColor l252, "white", 0
    Else
        LightState(105)=0:Flashers(flash_dict.Item(l105))=0
        SetLightColor l253, "white", 0:SetLightColor l254, "white", 0
    End If
    ' Witch
    SetLightColor L131,"orange",0
    WitchLower
    Smoke.Visible=0
    SetLightColor w1,"red",0
    SetLightColor w2,"red",0
    SetLightColor w3,"red",0
    SetLightColor w4,"red",0
    ' Haunted Forest Target
    SetLightColor L19,"white",0:SetLightColor L90,"white",0
    ' Crystal Ball lights
    If bCrystalBallMode=0 Then
        If L72.State=2 Then SetLightColor L72,cb_Color,bBLit: bBLit=0
        If L73.State=2 Then SetLightColor L73,cb_Color,bALit: bALit=0
        If L74.State=2 Then SetLightColor L74,cb_Color,bL1Lit:bL1Lit=0
        If L75.State=2 Then SetLightColor L75,cb_Color,bL2Lit:bL2Lit=0
    End If
    SkillShotClear
End Sub

' ******************************************************************************
'                       GAME MODE PROCESSING
' ******************************************************************************
' SOTR - Somewhere Over the rainbow jewels
Dim SOTR
SOTR = 0

' OZ Lanes
Sub OZ_DMD2
    Dim tmp0, tmp1, tmp2, tmp3

    If bHU=0 AND DMD_Option=True Then
        DMD2Flush
        If(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = " OZ LANE BONUS"
            tmp2 = FillLine(2, "ADVANCE", " X" & FormatScore(Multipliers(1)))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
        vpmTimer.AddTimer 3000,"DMD2_Clear '"
    End If
End Sub

Sub ClearOZLaneFlashers
    LightState(104) = 0
    LightState(105) = 0
    Flashers(flash_dict.Item(l104))=0:Flashers(flash_dict.Item(l105))=0
End Sub

Sub RestoreOZLaneFlashers
    If l253.state=1 Then
        LightState(105)=1:Flashers(flash_dict.Item(l105))=1
    End If
    If l252.state=1 Then
        LightState(104)=1:Flashers(flash_dict.Item(l104))=1
    End If
End Sub



' HOADC Mode
Dim HOADC_Color(6), HOADC_Queue(6)
Dim HOADC_Base, HOADC_FirstX, HOADC_SecondX, HOADC_collects, HOADC_points
Dim bHOADC_Score

Dim color_dict
Set color_dict = CreateObject("Scripting.Dictionary")

color_dict.Add 1, "purple"
color_dict.Add 2, "blue"
color_dict.Add 3, "green"
color_dict.Add 4, "white"
color_dict.Add 5, "yellow"
color_dict.Add 6, "orange"
color_dict.Add 7, "red"
color_dict.Add 8, "rainbow"

Sub HOADC_Select
    If bBTWW=1 OR bDingDong=1 OR bSOTR=1 Then Exit Sub
    Dim i, nbr_lights
    For i = 1 To 6
        nbr_lights = INT(RND * 8 + 1)    ' Choose a color
        If HOADC_Color(i)="rainbow" Then StopRainbow
        HOADC_Color(i) = color_dict.Item(nbr_lights)
    Next
    nbr_lights = 2 + INT(RND * 3 + 1)    ' Select at least two lights & five at most
    L26.State=0:L35.State=0:L84.State=0:L99.State=0:L136.State=0:L138.State=0
    i = 1
    Do While i<nbr_lights
        Select Case INT(RND * 5 + 1)
          case 1  ' Only One Orbit Horse Light can be On
            If L26.State<>1 AND L99.State<>1 Then
                Select Case INT(RND*2)
                  case 0
                    SetLightColor L26,HOADC_Color(1),1
                  case 1
                    SetLightColor L99,HOADC_Color(4),1
                End Select
                i = i + 1
            End If
          case 2
            If L35.State<>1 Then SetLightColor L35,HOADC_Color(2),1: i = i + 1
          case 3
            If L84.State<>1 Then SetLightColor L84,HOADC_Color(3),1: i = i + 1
          case 4
            If L136.State<>1 Then SetLightColor L136,HOADC_Color(5),1: i = i + 1
          case 5
            If L138.State<>1 Then SetLightColor L138,HOADC_Color(6),1: i = i + 1
        End Select
    Loop
    vpmTimer.AddTimer 2000,"HOADC_BG_clr '"
End Sub

Sub HOADC_Hit(switch)
    If bBTWW=1 OR bDingDong=1 OR bSOTR=1 Then Exit Sub
    Dim h_color, i
    If switch<>LastSwitchHit Then
        Select Case switch
          case 87
            h_color = HOADC_Color(1)
          case 15
            h_color = HOADC_Color(2)
          case 9
            h_color = HOADC_Color(3)
          case 64
            h_color = HOADC_Color(4)
          case 17, 73
            h_color = HOADC_Color(5)
          case 75
            h_color = HOADC_Color(6)
        End Select
        For i = 1 To 6
            If i = 6 Then
                HOADC_Queue(1) = h_color
            Else
                HOADC_Queue(7 - i) = HOADC_Queue(6 - i)
            End If
        Next
        If HOADC_Queue(4 + HOADC_collects)<>"none" Then SetLightColor L28,"white",2
        HOADC_BG_On
    End If
End Sub

Sub HOADC_Score
    Dim i, horse_nbr
'    Dim HOADC_repeat(9)
    Dim tmp0, tmp1, tmp2, tmp3

    If bEndOfGame=1 Then Exit Sub
    HOADC_BG_On
    HOADC_Score_Calc
    PlaySound "voc38_HorseOfADifferentColor"
    If UltraDMD_Option=True Then UltraDMD_simple "hoadc.png", " ", "SCORE: " &Cstr(HOADC_points), 5000
    If bHU=0 AND DMD_Option=True Then
        DMD2Flush
        If(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "     HOADC"
            tmp2 = FillLine(2, "COLLECTED", FormatScore(HOADC_points))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
        vpmTimer.AddTimer 3000,"DMD2_Clear '"
    End If
    AddScore_points HOADC_points
    For i = 1 To 6
        HOADC_Queue(i) = "none"
    Next
    HOADC_Base=0:HOADC_FirstX=1:HOADC_SecondX=1
    If HOADC_collects=0 AND horse_nbr>5 Then AwardExtraBall
    Select Case horse_nbr
      case 4
        munchkinX = munchkinX + 0.5
      case 5
        munchkinX = munchkinX + 1
      case 6
        munchkinX = munchkinX + 1.5
    End Select
    If munchkinX > 10 Then munchkinX = 10
    HOADC_Collects = HOADC_collects + 1
    If HOADC_collects>2 Then HOADC_collects = 2
    SetLightColor L28,"white",0
    HOADC_Select
End Sub


Sub HOADC_Score_Calc
    Dim i, horse_nbr
    Dim HOADC_repeat(9)

    HOADC_Base = 0
    For i = 1 To 9
        HOADC_repeat(i) = 0
    Next
    For i = 1 To 6
        Select Case HOADC_Queue(i)
          case "purple"
            HOADC_Base = HOADC_Base + 100
            HOADC_repeat(1) = HOADC_repeat(1) + 1
          case "blue"
            HOADC_Base = HOADC_Base + 150
            HOADC_repeat(2) = HOADC_repeat(2) + 1
          case "green"
            HOADC_Base = HOADC_Base + 200
            HOADC_repeat(3) = HOADC_repeat(3) + 1
          case "white"
            HOADC_Base = HOADC_Base + 500
            HOADC_repeat(4) = HOADC_repeat(4) + 1
          case "yellow"
            HOADC_Base = HOADC_Base + 250
            HOADC_repeat(5) = HOADC_repeat(5) + 1
          case "orange"
            HOADC_Base = HOADC_Base + 300
            HOADC_repeat(6) = HOADC_repeat(6) + 1
          case "red"
            HOADC_Base = HOADC_Base + 350
            HOADC_repeat(7) = HOADC_repeat(7) + 1
          case "rainbow"
            HOADC_Base = HOADC_Base + 400
            HOADC_repeat(8) = HOADC_repeat(8) + 1
          case "none"
            HOADC_repeat(9) = HOADC_repeat(9) + 1
        End Select
    Next
    horse_nbr = 6 - HOADC_repeat(9)
    For i = 1 To 7
        If HOADC_repeat(i)>1 AND HOADC_repeat(i)<horse_nbr Then
            HOADC_FirstX = 0.5
        Else
            If HOADC_repeat(i)=horse_nbr Then HOADC_FirstX = horse_nbr
        End If
    Next
    If HOADC_repeat(8)>1 AND HOADC_repeat(8)<horse_nbr Then
        HOADC_SecondX = 0.5
    Else
        If HOADC_repeat(8)=horse_nbr Then HOADC_SecondX = 5
    End If
    HOADC_points = HOADC_Base * HOADC_FirstX * HOADC_SecondX
    HOADC_points = INT(HOADC_points)
End Sub


Sub HOADC_BG_On
    If bFireballFrenzyStarted=1 Then Exit Sub
    If bBTWW=1 OR bDingDong=1 OR bSOTR=1 Then Exit Sub
    Dim i, str_tmp
    Controller.B2SSetData "HOADC_Base",1
    For i = 1 To 6
        str_tmp = "Horse" & Cstr(i) & "_" & HOADC_Queue(i)
        Controller.B2SSetData str_tmp,1
    Next
    vpmTimer.AddTimer 500,"HOADC_Select '"
End Sub

Sub HOADC_BG_clr
'    If bBTWW=1 OR bDingDong=1 OR bSOTR=1 Then Exit Sub
    Dim i, j, str_tmp
    Controller.B2SSetData "HOADC_Base",0
    For i = 1 To 6
        For j = 1 to 8
            str_tmp = "Horse" & Cstr(i) & "_" & color_dict.Item(j)
            Controller.B2SSetData str_tmp,0
        Next
    Next
End Sub


' TOTO Mode
Dim totoCnt, bTOTOMode, bTOTOExhausted, bTOTODone
Dim l30_state, l41_state, l49_state, l76_state, l134_state, TOTOTimeLeft

Sub TOTOMode
    Dim i
    SetLightColor l37,"red",2
    SetLightColor l38,"red",2
    SetLightColor l39,"red",2
    SetLightColor l40,"red",2
    If bMultiBallMode=1 Then
        If DMD_Option=True Then DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "BALL SAVED"), eBlinkfast, 800, True, ""
        If UltraDMD_Option=True Then UltraDMD_DisplayImage("Ball_Saved.png")
        AddMultiball 1
        Exit Sub
    End If
    bTOTOMode=1
    bTOTODone=0
    totoCnt = totoCnt + 1
    l30_state = l30.State: SetLightColor l30,"white",0
    l41_state = l41.State: SetLightColor l41,"white",0
    l49_state = l49.State: SetLightColor l49,"orange",0
    l76_state = l76.State: SetLightColor l76,"white",0
    l134_state = l134.State: SetLightColor l134,"white",0
    HOADC_BG_clr
    WitchHUStop
    HouseSpinStop
    Select Case totoCnt
      case 1
        SetLightColor l134,"white",2
        Kicker_TOTO.enabled=True: hold_time=8000
'        debug.print "Kicker.TOTO Enabled - case 1"
      case 2
        SetLightColor l134,"white",2
        SetLightColor l30,"white",2
        SetLightColor l41,"white",2
      case 3
        SetLightColor l134,"white",2
        SetLightColor l30,"white",2
        SetLightColor l41,"white",2
        SetLightColor l76,"white",2
        SetLightColor l65,"white",2
      case Else
        SetLightColor l37,"red",2
        SetLightColor l38,"red",2
        SetLightColor l39,"red",2
        SetLightColor l40,"red",2
        bTOTOExhausted=1
        'If PUPpack=True Then DOF 98, DOFPulse
        Controller.B2sSetData "Gulch_3",1
        vpmtimer.addtimer 5000, "Controller.B2sSetData ""Gulch_3"",0 '"
        Exit Sub
    End Select
    Controller.B2sSetData "Gulch_2",1
    If PUPpack=True Then DOF 99, DOFPulse
    PlaySound "voc9_Gulch", -1
    If bDorothyCaptureRdy=1 AND bDorothyCaught=0 Then
        Kicker_Mag.Enabled=False
        LightState(115)=0:Flashers(flash_dict.Item(l115))=0
    End If
    If bDorothyCaptureRdy=1 AND bDorothyCaught=1 Then
        LightState(153)=0:Flashers(flash_dict.Item(l153))=0
    End If
    LightSeqTOTO.Play SeqAlloff
    ' PF MX Side Effect - All Off
    For i = 310 to 317
        DOF i, DOFOff
    Next
    Wizard_reset
    ClearOZLaneFlashers
    CreateNewBall
    TOTOTimeLeft=31
    TOTOModeTimer.enabled = True
End Sub

Sub TOTOMode_Check(switch)
    Select Case switch
      case 5  ' Munchkin House Entrance
        MHouse_BHandler
        Exit Sub
      case 9  ' Crystal Ball VUK
        l76.State=0: l65.State=0
        If totoCnt=3 AND l134.State=2 AND l30.State=0 Then
            Kicker_TOTO.enabled=True: hold_time=8000
'            debug.print "Kicker.TOTO Enabled - sw9 Hit"
        End If
      case 15 ' Throne
        l30.State=0: l41.State=0
        If totoCnt=2 AND l134.State=2 Then
            Kicker_TOTO.enabled=True: hold_time=8000
'            debug.print "Kicker.TOTO Enabled - sw15/2 Hit"
        Else
            If totoCnt=3 AND l134.State=2 AND l76.State=0 Then
                Kicker_TOTO.enabled=True: hold_time=8000
'                debug.print "Kicker.TOTO Enabled - sw15/3 Hit"
            End If
        End If
      case 76 ' Ramp
        l134.State=0
    End Select
    If l30.State=0 AND l76.State=0 AND l134.State=0 Then
        bTOTODone=1
        TOTOModeTimer.enabled = False
        StopSound "voc9_Gulch"
        If PUPpack=True Then
            vpmtimer.addtimer 500, "DOF 97, DOFPulse '"
        Else
            PlaySound "voc7_TotoBack"
        End If
        Controller.B2sSetData "Gulch_2",0
        Controller.B2sSetData "Gulch_4",1
        vpmtimer.addtimer 2000, "Controller.B2sSetData ""Gulch_4"",0 '"
        BallSaveDisplayClear
        TOTOMode_End
    End If
End Sub

Sub TOTOMode_End
    Dim obj, bg_id
    bTOTOMode=0
    LightSeqTOTO.StopPlay
    TOTOModeTimer.enabled = False
    If PUPpack=True Then DOF 98, DOFPulse
    StopSound "voc9_Gulch"
    vpmtimer.addtimer 2000, "ClearBG_TOTO '"

    If bDorothyCaptureRdy=1 AND bDorothyCaught=0 Then
        Kicker_Mag.Enabled=True
        LightState(115)=1:Flashers(flash_dict.Item(l115))=1
    End If
    If bDorothyCaptureRdy=1 AND bDorothyCaught=1 Then
        LightState(153)=1:Flashers(flash_dict.Item(l153))=1
    End If
    HouseSpinRestore
'    HouseSpinStop
    l30.State = l30_state
    l41.State = l41_state
    l49.State = l49_state
    l76.State = l76_state
    l134.State = l134_state
    If bFireballFrenzy=1 AND bFireballFrenzyStarted=1 Then WitchRaise
    If SOTR_capture=1 OR bSOTR=1 Then Kicker_TOTO.enabled=True
    Wizard_restore
    RestoreOZLaneFlashers
    ' Rainbow Backglass Lights Restored
    For each obj in Lights_Rainbow
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
    If L1.State=1 AND L2.State=1 AND L3.State=1 AND L4.State=1 AND L5.State=1 AND L6.State=1 AND L7.State=1 Then
        bTwisterFlag=1
        If L137.State=0 Then
            DivRight.RotateToEnd: SetLightColor L134,"yellow",2:Controller.B2SSetData 233,1:bTwistAgain=0
        End If
    End If
    vpmtimer.addtimer 10000, "bTOTODone=0 '"
    vpmtimer.AddTimer 12000, "PlaySong ""voc18_OffToSee"" '"
End Sub

Sub TOTOModeTimer_Timer
    Dim obj, get_color, i
    If TOTOTimeLeft>0 Then TOTOTimeLeft = TOTOTimeLeft - 1
'    *** Uses T_ for time storage & digits for display  ***
    Controller.B2SSetData T_Ones,0:Controller.B2SSetData T_Tens,0
    If TOTOTimeLeft>9 Then T_Tens = "Td_" & CStr(INT(TOTOTimeLeft/10) MOD 10):Controller.B2SSetData T_Tens,1
    T_Ones = "T1_" & CStr(TOTOTimeLeft MOD 10):Controller.B2SSetData T_Ones,1
    SaveTOTOTimeDisplay
    If TOTOTimeLeft<1 AND bSW75_RampEnter=0 Then
        If TEST_MODE=True Then
            CenterPost(0)
            debug.print "CenterPost Down"
        End If
        If PUPpack=True Then
            ClearBG_TOTO
            DOF 98, DOFPulse: DOF 59, DOFPulse
        End If
        StopSound "voc9_Gulch"
        If SOTR_capture=1 OR bSOTR=1 Then
            Kicker_TOTO.enabled=True
        Else
            Kicker_TOTO.enabled=False
'            debug.print "Kicker.TOTO OFF - TOTOMode_Timer"
        End If
        Me.enabled = 0
'       *** Uses T_ for time storage & digits for display  ***
        Controller.B2SSetData T_Ones,0:Controller.B2SSetData T_Tens,0
        SolLFlipper 0
        SolRFlipper 0
        bDisableTable = 1
        BallSaveDisplayClear
        If bBallInPlungerLane=True Then PlungerIM.AutoFire:AutoFireCheck.Enabled = True
    End If
End Sub

Sub SaveTOTOTimeDisplay
    Dim tmp0, tmp1, tmp2, tmp3
    If DMD_Option=True Then
        DMD2Flush
        if(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "HELP TOTO ESCAPE"
            tmp2 = FillLine(2, "TIME LEFT", FormatScore(TOTOTimeLeft))
            tmp3 = " BALL SAVE"
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
    End If
End Sub

Sub BallSaveDisplayClear
    If DMD_Option=True Then DMD2_Clear
End Sub

' TNPLH Mode
Dim tnplhCnt, tnplhStage, bTNPLH, bTNPLHMode
Dim TNPLHTimeLeft
Dim TNPLHStates(32)

Sub TNPLHMode
    Dim obj, i
    SetLightColor l59,"green",2
    SetLightColor l60,"green",2
    SetLightColor l61,"green",2
    SetLightColor l62,"green",2
    SetLightColor l63,"green",2
    If bMultiBallMode=1 Then
        If DMD_Option=True Then DMD "", eNone, "", eNone, "", eNone, CenterLine(3, "BALL SAVED"), eBlinkfast, 800, True, ""
        If UltraDMD_Option=True Then UltraDMD_DisplayImage("Ball_Saved.png")
        AddMultiball 1
        Exit Sub
    End If
    bTNPLHMode = 1
    Kicker_TOTO.enabled=False
    HOADC_BG_clr
    WitchHUStop
    Twister_ClearCounts
    HouseSpinStop
    LightSeqTNPLH.Play SeqAlloff
    ' PF MX Side Effect - All Off
    For i = 310 to 317
        DOF i, DOFOff
    Next
    Wizard_reset
    ClearOZLaneFlashers
    If PUPpack=True Then
        DOF 329, DOFPulse
        vpmtimer.addtimer 2000, "DOF 322, DOFPulse '"
    End If
    Controller.B2SSetData "TNPLHS0_0",0
    Controller.B2sSetData "TNPLHS0_1",1
    vpmtimer.addtimer 2500, "Controller.B2sSetData ""TNPLHS0_1"",0 '"
    i = 0
    For each obj in Lights_TNPLHStages
        i = i + 1
        TNPLHStates(i) = obj.State
        SetLightColor obj,"white",0
    Next
    ' What was the last completed stage?
    Select Case tnplhStage
      case 1  ' Rainbow Lights
        For each obj in Lights_Rainbow
            SetLightColor obj,"white",2
        Next
        If PUPpack=True Then
            vpmtimer.addtimer 1800, "DOF 94, DOFPulse '"
        Else
            vpmtimer.addtimer 2100, "PlaySong ""voc8_Rainbow"" '"
        End If
        vpmtimer.addtimer 2000, "Controller.B2sSetData ""TNPLHS1_1"",1 '"
      case 2  ' Crystal & Throne Room
        SetLightColor l76,"white",2
        SetLightColor l65,"white",2
        SetLightColor l30,"white",2
        SetLightColor l41,"white",2
        If PUPpack=True Then
            vpmtimer.addtimer 1800, "DOF 93, DOFPulse '"
        Else
            vpmtimer.addtimer 2100, "PlaySong ""voc78_Marvel"" '"
        End If
        vpmtimer.addtimer 2000, "Controller.B2sSetData ""TNPLHS2_1"",1 '"
      case 3  ' Left & Right Orbits
        SetLightColor l100,"white",2
        SetLightColor l27,"white",2
        SetLightColor l98,"white",2
        SetLightColor l64,"white",2
        SetLightColor l80,"white",2
        SetLightColor l66,"white",2
        SetLightColor l25,"white",2
        SetLightColor l36,"white",2
        If PUPpack=True Then
           vpmtimer.addtimer 1800, "DOF 92, DOFPulse '"
        Else
            vpmtimer.addtimer 2100, "PlaySong ""voc79_ToShelter"" '"
        End If
        vpmtimer.addtimer 2000, "Controller.B2sSetData ""TNPLHS3_1"",1 '"
      case 4  ' Ramp
        SetLightColor l134,"white",2
        SetLightColor l143,"white",2
        Kicker_TOTO.enabled=True
        If PUPpack=True Then
            vpmtimer.addtimer 1800, "DOF 91, DOFPulse '"
            hold_time=10000
        Else
            vpmtimer.addtimer 2100, "PlaySong ""voc80_ToHouse"" '"
            hold_time=3000
        End If
        vpmtimer.addtimer 2400, "Controller.B2sSetData ""TNPLHS4_1"",1 '"
    End Select
    If bDorothyCaptureRdy=1 AND bDorothyCaught=0 Then
        Kicker_Mag.Enabled=False
        LightState(115)=0:Flashers(flash_dict.Item(l115))=0
    End If
    If bDorothyCaptureRdy=1 AND bDorothyCaught=1 Then
        LightState(153)=0:Flashers(flash_dict.Item(l153))=0
    End If
    CreateNewBall
    TNPLHTimeLeft=31
    TNPLHModeTimer.enabled = True
End Sub


Sub TNPLHMode_Check(switch)
    Dim obj, lite, tmpsw
    If switch=5 Then         '  Munchkin House Entrance
        MHouse_BHandler
        Exit Sub
    End If
    Select Case tnplhStage
      case 1    '   RAINBOW
        Select Case switch
          case 89, 90, 91, 92, 93, 94, 95
            set obj = light_dict.Item(switch)
            SetLightColor obj,"white",1
            tmpsw = switch + 1
            If tmpsw>95 Then tmpsw = 95
            set obj = light_dict.Item(tmpsw)
            If obj.State=1 Then
                tmpsw = switch - 1
                If tmpsw<89 Then tmpsw = 89
                set obj = light_dict.Item(tmpsw)
            End If
            SetLightColor obj,"white",1
            TNPLHTimeLeft = TNPLHTimeLeft + 5
            tnplhCnt=0
            For each obj in Lights_Rainbow
              If obj.State = 1 Then tnplhCnt = tnplhCnt + 1
            Next
            If tnplhCnt=7 Then
              tnplhStage = 2
              vpmtimer.addtimer 800, "Controller.B2sSetData ""TNPLHS1_1"",0 '"
              vpmtimer.addtimer 800, "Controller.B2sSetData ""TNPLHS2_1"",1 '"
              If PUPpack=True Then
                  DOF 95, DOFPulse
                  vpmtimer.addtimer 300, "DOF 93, DOFPulse '"
              Else
                  vpmtimer.addtimer 300, "PlaySong ""voc78_Marvel"" '"
              End If
              For each lite in Lights_Rainbow
                  SetLightColor lite,"white",0
              Next
              SetLightColor l76,"white",2
              SetLightColor l65,"white",2
              SetLightColor l30,"white",2
              SetLightColor l41,"white",2
              TNPLHTimeLeft = 31
            End If
        End Select
      case 2    '  Marvel
        Select case switch
          case 9
            SetLightColor l76,"white",0
            SetLightColor l65,"white",0
            L76.State=0
            TNPLHTimeLeft = TNPLHTimeLeft + 5
          case 15
            SetLightColor l30,"white",0
            SetLightColor l41,"white",0
            L30.State=0
            TNPLHTimeLeft = TNPLHTimeLeft + 5
        End Select
        If L30.State=0 AND L76.State=0 Then
            tnplhStage = 3
            vpmtimer.addtimer 2000, "Controller.B2sSetData ""TNPLHS2_1"",0 '"
            vpmtimer.addtimer 2000, "Controller.B2sSetData ""TNPLHS3_1"",1 '"
            If PUPpack=True Then
'                DOF 96, DOFPulse
                vpmtimer.addtimer 1250, "DOF 96, DOFPulse '"
                vpmtimer.addtimer 1500, "DOF 92, DOFPulse '"
            Else
                vpmtimer.addtimer 1000, "PlaySong ""voc79_ToShelter"" '"
            End If
            SetLightColor l76,"white",0
            SetLightColor l65,"white",0
            SetLightColor l30,"white",0
            SetLightColor l41,"white",0
            SetLightColor l100,"white",2
            SetLightColor l27,"white",2
            SetLightColor l98,"white",2
            SetLightColor l64,"white",2
            SetLightColor l80,"white",2
            SetLightColor l66,"white",2
            SetLightColor l25,"white",2
            SetLightColor l36,"white",2
            TNPLHTimeLeft = 31
        End If
      case 3    ' To House
        Select case switch
          case 64
            SetLightColor l98,"white",1
            SetLightColor l64,"white",1
            SetLightColor l80,"white",1
            SetLightColor l66,"white",1
            SetLightColor l100,"white",1
            TNPLHTimeLeft = TNPLHTimeLeft + 5
          case 87
            SetLightColor l25,"white",1
            SetLightColor l36,"white",1
            SetLightColor l27,"white",1
            TNPLHTimeLeft = TNPLHTimeLeft + 5
        End Select
        If l98.State=1 AND l25.State=1 Then
            tnplhStage = 4
            Kicker_TOTO.enabled=True
            vpmtimer.addtimer 1500, "Controller.B2sSetData ""TNPLHS3_1"",0 '"
            vpmtimer.addtimer 1500, "Controller.B2sSetData ""TNPLHS4_1"",1 '"
            If PUPpack=True Then
                vpmtimer.addtimer 500,  "DOF 90, DOFPulse '"
                vpmtimer.addtimer 1000, "DOF 91, DOFPulse '"
            Else
                vpmtimer.addtimer 300, "PlaySong ""voc80_ToHouse"" '"
            End If

            SetLightColor l100,"white",0
            SetLightColor l27,"white",0
            SetLightColor l98,"white",0
            SetLightColor l64,"white",0
            SetLightColor l80,"white",0
            SetLightColor l66,"white",0
            SetLightColor l25,"white",0
            SetLightColor l36,"white",0
            SetLightColor l134,"white",2
            SetLightColor l143,"white",2
            TNPLHTimeLeft = 31
        End If
      case 4     '  Final Stage - Ramp
        If switch=76 Then
            tnplhStage=1
            If PUPpack=True Then
                DOF 86, DOFPulse
                vpmtimer.addtimer 300, "DOF 327, DOFPulse '"
            Else
                vpmtimer.addtimer 300, "PlaySound ""voc4_notinkansas"" '"
            End If
            SetLightColor l134,"white",0
            SetLightColor l143,"white",0

            BallSaveDisplayClear
            vpmtimer.addtimer 800, "Controller.B2SSetData ""TNPLHS4_1"",0 '"
            vpmtimer.addtimer 800, "Controller.B2SSetData ""TNPLHS5_1"",1 '"
            vpmtimer.addtimer 5000, "Controller.B2sSetData ""TNPLHS5_1"",0 '"
            TNPLHMode_End
        End If
    End Select
End Sub


SUB TNPLHMode_End
    Dim obj, i, bg_id, get_color

    TNPLHModeTimer.enabled = False
    bTNPLHMode=0
    LightSeqTNPLH.StopPlay
    If PUPpack=True Then DOF 95,DOFPulse: DOF 96,DOFPulse: DOF 90,DOFPulse: DOF 86,DOFPulse
    vpmtimer.addtimer 2000, "StopSound Song '"
'    ClearBG_TNPLH
    Twister_ClearCounts
    BallSaveDisplayClear
    i=0
    For each obj in Lights_TNPLHStages
        i = i + 1
        get_color = lcolor_dict.Item(obj)
        SetLightColor obj, get_color, TNPLHStates(i)
'        obj.State = TNPLHStates(i)
    Next
    ' Rainbow Backglass Lights Restored
    For each obj in Lights_Rainbow
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
    If L1.State=1 AND L2.State=1 AND L3.State=1 AND L4.State=1 AND L5.State=1 AND L6.State=1 AND L7.State=1 Then
        bTwisterFlag=1
        If L137.State=0 Then
            DivRight.RotateToEnd: SetLightColor L134,"yellow",2:Controller.B2SSetData 233,1:bTwistAgain=0
        End If
    End If
'    If SOTR_capture=1 OR bSOTR=1 Then
'        Kicker_TOTO.enabled=True
'    Else
'        Kicker_TOTO.enabled=False
'    End If
    If bDorothyCaptureRdy=1 AND bDorothyCaught=0 Then
        Kicker_Mag.Enabled=True
        LightState(115)=1:Flashers(flash_dict.Item(l115))=1
    End If
    If bDorothyCaptureRdy=1 AND bDorothyCaught=1 Then
        LightState(153)=1:Flashers(flash_dict.Item(l153))=1
    End If
    HouseSpinRestore
    If bFireballFrenzy=1 AND bFireballFrenzyStarted=1 Then WitchRaise
    Wizard_restore
    RestoreOZLaneFlashers
End Sub


Sub TNPLHModeTimer_Timer
    Dim obj, get_color, i
    TNPLHTimeLeft = TNPLHTimeLeft - 1
    TNPLHTimeDisplay
    If TNPLHTimeLeft<1 AND bSW75_RampEnter=0 Then
        If TEST_MODE=True Then
            CenterPost(0)
            debug.print "CenterPost Down"
        End If
        Kicker_TOTO.enabled=False
        If PUPpack=True Then DOF 95, DOFPulse: DOF 96,DOFPulse: DOF 90,DOFPulse: DOF 86,DOFPulse
        StopSound Song
        Me.enabled = False
        BallSaveDisplayClear
'    *** Uses T_ for time storage & digits for display  ***
        Controller.B2SSetData T_Ones,0:Controller.B2SSetData T_Tens,0
        SolLFlipper 0
        SolRFlipper 0
        bDisableTable = 1
        If bBallInPlungerLane=True Then PlungerIM.AutoFire:AutoFireCheck.Enabled = True
    Else
'    *** Uses T_ for time storage & digits for display  ***
        Controller.B2SSetData T_Ones,0:Controller.B2SSetData T_Tens,0
        If TNPLHTimeLeft>9 Then T_Tens = "Td_" & CStr(INT(TNPLHTimeLeft/10) MOD 10):Controller.B2SSetData T_Tens,1
        T_Ones = "T1_" & CStr(TNPLHTimeLeft MOD 10):Controller.B2SSetData T_Ones,1
    End If

End Sub

Sub TNPLHTimeDisplay
    Dim tmp0, tmp1, tmp2, tmp3
    If DMD_Option=True Then
        DMD2Flush
        if(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "  HELP DOROTHY"
            tmp2 = FillLine(2, "TIME LEFT", FormatScore(TNPLHTimeLeft))
            tmp3 = " BALL SAVE"
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
    End If
End Sub

Sub TNPLH_TOTO_LightsReset
    ' TOTO Mode Lights
    If l37.State=2 OR l37.State=0 Then SetLightColor l37,"red",2 Else SetLightColor l37,"white",1
    If l38.State=2 OR l38.State=0 Then SetLightColor l38,"red",2 Else SetLightColor l38,"white",1
    If l39.State=2 OR l39.State=0 Then SetLightColor l39,"red",2 Else SetLightColor l39,"white",1
    If l40.State=2 OR l40.State=0 Then SetLightColor l40,"red",2 Else SetLightColor l40,"white",1

    ' TNPLH Mode Lights
    If l59.State=2 OR l59.State=0 Then SetLightColor l59,"green",2 Else SetLightColor l59,"white",1
    If l60.State=2 OR l60.State=0 Then SetLightColor l60,"green",2 Else SetLightColor l60,"white",1
    If l61.State=2 OR l61.State=0 Then SetLightColor l61,"green",2 Else SetLightColor l61,"white",1
    If l62.State=2 OR l62.State=0 Then SetLightColor l62,"green",2 Else SetLightColor l62,"white",1
    If l63.State=2 OR l63.State=0 Then SetLightColor l63,"green",2 Else SetLightColor l63,"white",1
End Sub

' INLANE HURRY UPS
Dim HUPot, HUCnt, HULow
Dim cnt_HU_Sound
Dim activeHU
Dim bHU

Dim hu_dict
Set hu_dict = CreateObject("Scripting.Dictionary")

hu_dict.Add L32, "WINKIE"
hu_dict.Add L33, "CRYSTAL"
hu_dict.Add L34, "OZ RIGHT"
hu_dict.Add L68, "RAINBOW"
hu_dict.Add L69, "WIZARD"
hu_dict.Add L70, "OZ LEFT"

Sub HU_Select
    Dim obj, get_color
    For each obj in Lights_HU
        get_color = lcolor_dict.Item(obj)
        SetLightColor obj, get_color, 0
    Next
    Select Case INT(RND * 6 + 1)
      case 1
        L32.State=2:L33.State=0:L34.State=0:L68.State=0:L69.State=0:L70.State=0
      case 2
        L32.State=0:L33.State=2:L34.State=0:L68.State=0:L69.State=0:L70.State=0
      case 3
        L32.State=0:L33.State=0:L34.State=2:L68.State=0:L69.State=0:L70.State=0
      case 4
        L32.State=0:L33.State=0:L34.State=0:L68.State=2:L69.State=0:L70.State=0
      case 5
        L32.State=0:L33.State=0:L34.State=0:L68.State=0:L69.State=2:L70.State=0
      case 6
        L32.State=0:L33.State=0:L34.State=0:L68.State=0:L69.State=0:L70.State=2
    End Select
End Sub

Sub HurryUp_Timer
    Dim obj

    For each obj in Lights_HU
        If obj.State=2 Then activeHU = hu_dict.Item(obj):obj.State=1
    Next
    Me.enabled = 0
    bHU=1
    HUPot = 1500 + HULow
    If HUPot>9999 Then HUPot=9999:HULow = HUPot - 1500
    If (cnt_HU_Sound MOD 5)=0 Then
        Select Case activeHU
          case "OZ_LEFT", "OZ_RIGHT", "WIZARD"
            If INT(Rnd + 0.5)=1 Then
                PlaySound "voc17_MakeHaste"
            Else
                PlaySound "voc54_Propel"
            End If
          case "WINKIE", "CRYSTAL", "RAINBOW"
            PlaySound "voc40_HurryPlease"
        End Select
    Else
        PlaySound "voc24_HU"
    End If
    cnt_HU_Sound = cnt_HU_Sound + 1
    HUScoreDisplay
    HUScore.Interval = 1000
    HUScore.Enabled = 1
'    HU_ClearScore
End Sub

Sub HUScoreDisplay
    Dim tmp0, tmp1, tmp2, tmp3
    If DMD_Option=True Then
        DMD2Flush
        if(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "INLANE HURRY UP"
            tmp2 = FillLine(2, activeHU, FormatScore(HUPot))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eBlink, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
    End If
End Sub

Sub HUScore_Timer
    HUPot = HUPot - 5
    If HUPot<=HULow Then
       HULow = 0
       HUStop
       Me.enabled = 0
       Exit Sub
    End If
    HUScoreDisplay
    Me.Interval = 10
End Sub

Sub HUDone
    Dim tmp0, tmp1, tmp2, tmp3
    HUScore.Enabled = 0
    AddScore_points HUPot
    HULow = HULow + HUPot
    If DMD_Option=True Then
        DMD2Flush
        If(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = " HURRY UP DONE"
            tmp2 = FillLine(2, activeHU, FormatScore(HUPot))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
        vpmTimer.AddTimer 4000,"HUstop '"
    End If
End Sub

Sub HUStop
    HurryUp.Enabled = 0
    HUScore.Enabled = 0
    HUPot = 0
    bHU=0
    If INT(Rnd*4 + 1)=1 Then PlaySound "voc58_ReturnConquer"
    HU_Select
    HU_ClearScore
End Sub

Sub HU_ClearScore
    If DMD_Option=True Then
        DMD2Flush
        if dqHead2=dqTail2 Then DMD2 "", eNone, "", eNone, "", eNone, "", eNone, 1000, True, ""
        DMDScoreNow
    End If
End Sub


' WITCH HURRYUP Modes
Dim witchHitCnt, witchHUPot, WitchHUCnt, witchHULow
Dim witchStep, witchDirection
Dim bWitchHU, bWitchDown

Sub WitchHurryUp_Timer
    Dim w_HU_pause
    w_HU_pause = 1000
    Controller.B2SSetData 189,0
    Me.enabled = 0
    Controller.B2SSetData "HU_Witch_1",0
    Controller.B2SSetData "HU_Witch_2",0
    Controller.B2SSetData "HU_Witch_3",0
    Select Case WitchHUCnt
      case 0
        Select Case INT(RND*3 + 1)
          case 1: PlaySound "voc41_SoonOver"
          case 2: PlaySound "voc56_MoreTrouble"
          case 3: PlaySound "voc57_Well"
        End Select
        Controller.B2SSetData "HU_Witch_1",1
      case 1
        Controller.B2SSetData "HU_Witch_2",1
      case 2
        Controller.B2SSetData "HU_Witch_3",1
      case Else
        vpmTimer.AddTimer w_HU_pause,"Controller.B2SSetData ""HU_Witch_0"",1 '" 'waits on animation
    End Select
    If WitchHUCnt>0 Then PlaySound "voc42_WitchHU"
    witchHUPot = 1500 + witchHULow
    If witchHUPot>9999 Then witchHUPot=9999
    BG_HUScore
    WitchHUScore.Interval = 1000
    WitchHUScore.Enabled = 1
    WitchHU_ClearScore
End Sub

Sub WitchHUStop
    Controller.B2SSetData 189,0
    If WitchHUCnt>0 Then StopSound "voc42_WitchHU"
    WitchLower
    WitchHurryUp.Enabled = 0
    WitchHUScore.Enabled = 0
    witchHUPot = 0
    witchHitCnt=0:bWitchHU=0
    SetLightColor L131,"orange",0
    Controller.B2SSetData "HU_Witch_1",0
    Controller.B2SSetData "HU_Witch_2",0
    Controller.B2SSetData "HU_Witch_3",0
    Controller.B2SSetData "HU_Witch_0",0
    SetLightColor l8,"red",0
    WitchHU_ClearScore
End Sub

Dim Ones, Tens, Hundreds, Thousands, TenK

Sub WitchHU_ClearScore
    D_DigitsClear
End Sub

Sub BG_HUScore
    D_DigitsDisplay witchHUPot
End Sub

Sub D_DigitsDisplay(d_digits)
    Ones = "D1_" & CStr(d_digits MOD 10):Controller.B2SSetData Ones,1
    Tens = "Dd_" & CStr(INT(d_digits/10) MOD 10):Controller.B2SSetData Tens,1
    If d_digits>99 Then Hundreds = "Dc_" & Cstr(INT(d_digits/100) MOD 10): Controller.B2SSetData Hundreds,1
    If d_digits>999 Then Thousands = "Dk_" & Cstr(INT(d_digits/1000) MOD 10) :Controller.B2SSetData Thousands,1
    If d_digits>9999 Then TenK = "Dtk_" & Cstr(INT(d_digits/10000) MOD 10) :Controller.B2SSetData TenK,1
End Sub

Sub D_DigitsClear
    Dim i
    For i = 0 To 9
        Ones = "D1_" & Cstr(i)
        Controller.B2SSetData Ones,0
        Tens = "Dd_" & Cstr(i)
        Controller.B2SSetData Tens,0
        Hundreds = "Dc_" & Cstr(i)
        Controller.B2SSetData Hundreds,0
        Thousands = "Dk_" & Cstr(i)
        Controller.B2SSetData Thousands,0
    Next
    For i = 1 To 9
        TenK = "Dtk_" & Cstr(i)
        Controller.B2SSetData TenK,0
    Next

End Sub

Sub WitchHUScore_Timer
    witchHUPot = witchHUPot - 1
    If witchHUPot<=WitchHULow Then
       WitchHUStop
       Me.enabled = 0
       Exit Sub
    End If
    Controller.B2SSetData Ones,0
    Controller.B2SSetData Tens,0
    Controller.B2SSetData Hundreds,0
    Controller.B2SSetData Thousands,0
    BG_HUScore
    Me.Interval = 20
End Sub

' FIREBALL FRENZY
Dim blueShots, blueshotcnt, redshotcnt
Dim bFireballFrenzy, bFireballFrenzyStarted
Dim J_Ones, J_Tens
Dim ff_lock


Dim ff_dict
Set ff_dict = CreateObject("Scripting.Dictionary")

Sub FireballFrenzy
    If bFireballFrenzyStarted=0 Then
        If bSkillshotReady=1 Then SkillShotMade WITCH_SKILLSHOT
        If bwitchHU<>0 Then
            WitchHUStop
            vpmTimer.AddTimer 2000,"FireballFrenzyStart '"
        Else
            FireballFrenzyStart
        End If
    End If
End Sub

Sub FireballFrenzyStart
    WitchRaise
    bBTWW_3=1
    If PUPpack=False Then PlaySound "voc5_playball"
    If PUPpack=True Then
        If bMultiBallMode=0 Then DOF 88, DOFPulse
        If bMultiBallMode=1 AND bWitchHU=0 Then DOF 78, DOFPulse
    End If
    Controller.B2sSetData "Quadrant_Lt3",1
    Controller.B2SSetData "FireballFrenzy",1
    FFClock.enabled = 1
    blueShots = INT(blueShots/2)
    redshotcnt = 0
    FF_ChooseShots
    bFireballFrenzyStarted=1
    Display_MX "FRENZY"
End Sub


Sub FF_ChooseShots
    Dim obj, obj1, cnt, jcolor, tmpflag
    tmpflag=0
    cnt = 2 + INT(RND * 2 + 1)
    ff_dict.RemoveAll
    For each obj in Lights_FF
        SetLightColor obj,"white",0
    Next
    Do While cnt<>0
        If cnt = 1 Then
           jcolor = "blue"
        Else
           jcolor = "red"
        End If
        Select Case INT(RND * 6 + 1)
          Case 1
            Set obj = l80: Set obj1 = l66
          Case 2
            Set obj = l132: Set obj1 = l58
          Case 3
            Set obj = l139: Set obj1 = l42
            tmpflag=1
          Case 4
            Set obj = l30: Set obj1 = l41
          Case 5
            Set obj = l25: Set obj1 = l36
          Case 6
            Set obj = l76: Set obj1 = l65
        End Select
        If obj.State=2 Then
            cnt = cnt + 1
        Else
            SetLightColor obj,jcolor,2
            SetLightColor obj1,jcolor,2
            If tmpflag=1 Then SetLightColor l98,jcolor,2:SetLightColor l64,jcolor,2:tmpflag=0
            ff_dict.Add obj, jcolor
        End If
        cnt = cnt - 1
    Loop
End Sub

Sub FireballFrenzyUpdate(obj)
    Dim get_color
    Dim tmp0, tmp1, tmp2, tmp3
    If ff_dict.Exists(obj) Then
        get_color = ff_dict.Item(obj)
    End If
    If get_color="red" Then
        redshotcnt = redshotcnt + 1
        If redshotcnt>2 Then FireballFrenzyStop: ff_lock=0: Exit Sub
        blueShots = blueShots - 500
    Else
        Controller.B2SSetData J_Ones,0:Controller.B2SSetData J_Tens,0
        AddScore_points blueShots
        If DMD_Option=True Then
            DMD2Flush
            If(dqHead2 = dqTail2)Then
                tmp0 = ""
                tmp1 = "FIREBALL FRENZY"
                tmp2 = FillLine(2, "JACKPOT", FormatScore(blueShots))
                tmp3 = ""
                DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
            End If
            vpmTimer.AddTimer 5000,"DMD2_Clear '"
        End If
        blueshotcnt = blueshotcnt + 1
        Select Case blueshotcnt
          case 10, 25, 45, 70, 100
            AwardExtraBall
            AwardJewel
        End Select
        blueShots = blueShots + 500
        If blueshotcnt>1 AND bRescueMB=0 Then
            Display_MX "JACKPOT"
        End If
        vpmTimer.AddTimer 500,"PlaySound ""voc61_jackpot"" '"
    End If
    ' Jackpots collected
    If blueshotcnt>9 Then J_Tens = "Jd_" & CStr(INT(blueshotcnt/10) MOD 10):Controller.B2SSetData J_Tens,1
    J_Ones = "J1_" & CStr(blueshotcnt MOD 10):Controller.B2SSetData J_Ones,1
    FF_ChooseShots
    ff_lock=0
End Sub

Sub FireballFrenzyStop
    Dim i, obj
    redshotcnt = 0
    blueShots = 0
    bFireballFrenzy=0
    bFireballFrenzyStarted = 0
    WitchLower
    bWitchHU=0:witchHUCnt=0:witchHitCnt=0:WitchHULow = 0
    Controller.B2SSetData "FireballFrenzy",0
    Controller.B2SSetData "FireballFrenzy1",0
    For i = 0 To 9
        J_Ones = "J1_" & Cstr(i)
        Controller.B2SSetData J_Ones,0
        J_Tens = "Jd_" & Cstr(i)
        Controller.B2SSetData J_Tens,0
    Next
    FFClock.enabled = 0
    For each obj in Lights_FF
        SetLightColor obj,"white",0
    Next
    SetLightColor l130,"white",0:SetLightColor l131,"orange",0:SetLightColor l133,"orange",0
End Sub

Sub FFClock_Timer
    Me.enabled = 0
    Controller.B2SSetData "FireballFrenzy",0
    Controller.B2SSetData "FireballFrenzy1",1
    ' Jackpots collected
    If blueshotcnt>9 Then J_Tens = "Jd_" & CStr(INT(blueshotcnt/10) MOD 10):Controller.B2SSetData J_Tens,1
    J_Ones = "J1_" & CStr(blueshotcnt MOD 10):Controller.B2SSetData J_Ones,1
End Sub


' BATTLE THE WICKED WITCH - BTWW
Dim bBTWW, bBTWW_1, bBTWW_2, bBTWW_3, bBTWW_4
Dim BTWWCnt, BTWWjackpot, bBTWW_check, bBTWW_init
Dim BTWW_sw_color(4)
Dim BTWWLiteState(150)


Sub BTWW
    Dim obj, i
    bBTWW = 1: bBTWW_init = 0
    BTWWCnt = 0: BTWWjackpot = 6250

    SkillShotStop:skillShot_TimeLeft=0
    WitchHUStop
    HouseSpinStop
    Kicker_TOTO.enabled=False
    If bDorothyCaptureRdy=1 AND bDorothyCaught=0 Then
        Kicker_Mag.Enabled=False
        LightState(115)=0:Flashers(flash_dict.Item(l115))=0
    End If
    If bDorothyCaptureRdy=1 AND bDorothyCaught=1 Then
        LightState(153)=0:Flashers(flash_dict.Item(l153))=0
    End If
    WitchRaise
    SetLightColor w1,"red",0
    SetLightColor w2,"red",0
    SetLightColor w3,"red",0
    SetLightColor w4,"red",0
    AddScore_points 10000
    For i = 1 To 4
        BTWW_sw_color(i) = "white"
    Next
    Controller.B2sSetData "Quadrant_Lt1",0
    Controller.B2sSetData "Quadrant_Lt2",0
    Controller.B2sSetData "Quadrant_Lt3",0
    Controller.B2sSetData "Quadrant_Lt4",0
    i=0
    For each obj in Lights_BTWW
        i = i + 1
        BTWWliteState(i) = obj.State
    Next
    For each obj in Lights_BTWW
        SetLightColor obj, "green", 1
    Next
    SetLightColor l25,"white",1
    SetLightColor l26,"white",1
    SetLightColor l30,"white",1
    SetLightColor l35,"white",1
    SetLightColor l36,"white",1
    SetLightColor l41,"white",1
    SetLightColor l64,"white",1
    SetLightColor l65,"white",1
    SetLightColor l66,"white",1
    SetLightColor l76,"white",1
    SetLightColor l80,"white",1
    SetLightColor l84,"white",1
    SetLightColor l98,"white",1
    SetLightColor l99,"white",1
    bMultiBallMode=1
    AddMultiball 1
    ChangeGi "green"
    If PUPpack=True Then DOF 70,DOFPulse
    PlaySound "voc39_MyLittleParty"
    Controller.B2sSetData "WW_01",1
    btoggle = True
    BTWWShots.Enabled = True
    Controller.B2SSetData "BTWW_6",1
    vpmTimer.AddTimer 3000,"Controller.B2sSetData ""WW_03"",1 '"
    vpmTimer.AddTimer 3000,"Controller.B2sSetData ""WW_01"",0 '"
End Sub

' A substitue for a backglass animation
Dim btoggle
btoggle = True

Sub BTWWShots_Timer
    bBTWW_init=1
    If btoggle=True Then
        Controller.B2sSetData "WW_Redshots",1
        Controller.B2sSetData "WW_WhiteShots",0
        btoggle = False
    Else
        Controller.B2sSetData "WW_WhiteShots",1
        Controller.B2sSetData "WW_Redshots",0
        btoggle = True
    End If
End Sub

Sub BTWWShots_Off
    BTWWShots.Enabled = False
    Controller.B2sSetData "WW_Redshots",0
    Controller.B2sSetData "WW_WhiteShots",0
End Sub

Sub BTWW_end
    Dim obj, i, get_color, tmp_str
    bBTWW=0:bBTWW_1=0:bBTWW_2=0:bBTWW_3=0:bBTWW_4=0
    WitchLower
    bMultiBallMode=0
    mBalls2Eject=0
    i=0
    If bDingDong=0 Then
        For each obj in Lights_BTWW
            i = i + 1
            get_color = lcolor_dict.Item(obj)
            SetLightColor obj, get_color, BTWWliteState(i)
        Next
        TNPLH_TOTO_LightsReset
        HouseSpinRestore
        If bDorothyCaptureRdy=1 AND bDorothyCaught=0 Then
            Kicker_Mag.Enabled=True
            LightState(115)=1:Flashers(flash_dict.Item(l115))=1
        End If
        If bDorothyCaptureRdy=1 AND bDorothyCaught=1 Then
            LightState(153)=1:Flashers(flash_dict.Item(l153))=1
        End If
    End If
    BTWWShots_Off
    If PUPpack=True Then DOF 71,DOFPulse
    If SOTR_capture=1 OR bSOTR=1 Then
        Kicker_TOTO.enabled=True
    Else
        Kicker_TOTO.enabled=False
    End If
    If l134.State=0 Then
        DivRight.RotateToStart
    Else
        DivRight.RotateToEnd
    End If
    ChangeGi "white"
    ClearBG_BTWW
    GiOn
    sw73.IsDropped=0:SetLightsOff(73)
    sw17.TimerInterval = 1000
    vpmtimer.AddTimer 500, "PlaySong ""voc18_OffToSee"" '"
    BTWWCnt=0
End Sub

Sub BTWW_Jackpot_Clear
    Controller.B2SSetData "BTWW_6",0
    Controller.B2SSetData "BTWW_12",0
    Controller.B2SSetData "BTWW_18",0
    Controller.B2SSetData "BTWW_25",0
    Controller.B2SSetData "BTWW_31",0
    Controller.B2SSetData "BTWW_37",0
    Controller.B2SSetData "BTWW_43",0
    Controller.B2SSetData "BTWW_50",0
    Controller.B2SSetData "BTWW_56",0
    Controller.B2SSetData "BTWW_62",0
    Controller.B2SSetData "BTWW_75",0
    Controller.B2SSetData "BTWW_100",0
End Sub

Sub BTWWMode_check(switch)
    Dim obj, i, tmp_str

    If switch=5 Then         '  Munchkin House Entrance
        MHouse_BHandler
        Exit Sub
    End If
    If bBTWW_init=0 Then Exit Sub

    If BTWWCnt<4 Then
        Select Case switch
          case 9    ' Crystal Ball VUK
            BTWWCnt = BTWWCnt + 1
            If BTWW_sw_color(1)="white" Then
                BTWW_sw_color(1) = "red"
                SetLightColor l76,"red",1:SetLightColor l65,"red",1:SetLightColor l84,"red",1
                BTWWjackpot = 2*BTWWjackpot
            Else
                BTWWjackpot = BTWWjackpot + 6250
            End If
          case 15    ' Throne
            BTWWCnt = BTWWCnt + 1
            If BTWW_sw_color(2)="white" Then
                BTWW_sw_color(2) = "red"
                SetLightColor l35,"red",1:SetLightColor l30,"red",1:SetLightColor l41,"red",1
                BTWWjackpot = 2*BTWWjackpot
            Else
                BTWWjackpot = BTWWjackpot + 6250
            End If
          case 64    ' Left Orbit
            BTWWCnt = BTWWCnt + 1
            If BTWW_sw_color(3)="white" Then
                BTWW_sw_color(3) = "red"
                SetLightColor l80,"red",1:SetLightColor l98,"red",1:SetLightColor l99,"red",1
                SetLightColor l66,"red",1:SetLightColor l64,"red",1
                BTWWjackpot = 2*BTWWjackpot
            Else
                BTWWjackpot = BTWWjackpot + 6250
            End If
          case 87    'Right Orbit
            BTWWCnt = BTWWCnt + 1
            If BTWW_sw_color(4)="white" Then
                BTWW_sw_color(4) = "red"
                SetLightColor l25,"red",1:SetLightColor l36,"red",1:SetLightColor l26,"red",1
                BTWWjackpot = 2*BTWWjackpot
            Else
                BTWWjackpot = BTWWjackpot + 6250
            End If
          case Else
            AddScore Switch
        End Select
        Select Case BTWWCnt
          case 1
            Controller.B2SSetData "WW_03",0
            Controller.B2SSetData "WW_04",1
            AddScore_points 2500
          case 2
            Controller.B2SSetData "WW_03",0
            Controller.B2SSetData "WW_04",0
            Controller.B2SSetData "WW_05",1
            AddScore_points 5000
          case 3
            Controller.B2SSetData "WW_03",0
            Controller.B2SSetData "WW_04",0
            Controller.B2SSetData "WW_05",0
            Controller.B2SSetData "WW_06",1
            AddScore_points 7500
          case 4
            BTWWShots.Enabled = False
            BTWWShots_Off
            Controller.B2SSetData "WW_03",0
            Controller.B2SSetData "WW_04",0
            Controller.B2SSetData "WW_05",0
            Controller.B2SSetData "WW_06",0
            Controller.B2SSetData "WW_07",1
            AddScore_points 10000
            SetLightColor l25,"green",1
            SetLightColor l36,"green",1
            SetLightColor l26,"green",1
            SetLightColor l30,"green",1
            SetLightColor l41,"green",1
            SetLightColor l35,"green",1
            SetLightColor l76,"green",1
            SetLightColor l65,"green",1
            SetLightColor l80,"green",1
            SetLightColor l66,"green",1
            SetLightColor l84,"green",1
            SetLightColor l98,"green",1
            SetLightColor l64,"green",1
            SetLightColor l99,"green",1
            SetLightColor l132,"white",1
            SetLightColor l58,"white",1
            SetLightColor l136,"white",1
            sw17.TimerInterval = 30000
            sw73.IsDropped=1:SetLights(73)
        End Select
        Select Case BTWWjackpot
          case 12500
            Controller.B2SSetData "BTWW_6",0
            Controller.B2SSetData "BTWW_12",1
          case 18750
            Controller.B2SSetData "BTWW_12",0
            Controller.B2SSetData "BTWW_18",1
          case 25000
            Controller.B2SSetData "BTWW_12",0
            Controller.B2SSetData "BTWW_18",0
            Controller.B2SSetData "BTWW_25",1
          case 31250
            Controller.B2SSetData "BTWW_25",0
            Controller.B2SSetData "BTWW_31",1
          case 37500
            Controller.B2SSetData "BTWW_18",0
            Controller.B2SSetData "BTWW_31",0
            Controller.B2SSetData "BTWW_37",1
          case 43750
            Controller.B2SSetData "BTWW_37",0
            Controller.B2SSetData "BTWW_43",1
          case 50000
            Controller.B2SSetData "BTWW_25",0
            Controller.B2SSetData "BTWW_50",1
          case 56250
            Controller.B2SSetData "BTWW_50",0
            Controller.B2SSetData "BTWW_56",1
          case 62500
            Controller.B2SSetData "BTWW_31",0
            Controller.B2SSetData "BTWW_62",1
          case 75000
            Controller.B2SSetData "BTWW_37",0
            Controller.B2SSetData "BTWW_75",1
          case 100000
            Controller.B2SSetData "BTWW_50",0
            Controller.B2SSetData "BTWW_100",1
        End Select
    Else
        Select Case switch
          case 17        ' Castle VUK
            If BTWWCnt=4 Then
                BTWWCnt = 5
                AddScore_points 10000
                SetLightColor l132,"green",1
                SetLightColor l58,"green",1
                SetLightColor l136,"green",1
                SetLightColor l134,"white",1
                SetLightColor l137,"white",1
                SetLightColor l138,"white",1
                SetLightColor l139,"white",1
                SetLightColor l42,"white",1
                SetLightColor l96,"white",1
                SetLightColor l143,"white",1
                SetLightColor w1,"red",2
                SetLightColor w2,"red",2
                SetLightColor w3,"red",2
                SetLightColor w4,"red",2
                DivRight.RotateToEnd
                Controller.B2SSetData "WW_07",0
                Controller.B2SSetData "WW_08",1
            End If
          case 57, 58    ' Witch Hit
            If BTWWCnt=6 Then
                BTWWCnt=7
                If TEST_MODE=True Then
                    CenterPost(0)
                    debug.print "CenterPost Down"
                End If
                AddScore_points 10000
                For each obj in Lights_BTWW
                    SetLightColor obj, "green", 2
                Next
                ClearBG_BTWW
                ' Melting Animation
                Controller.B2SSetData 178,1
                SetLightColor w1,"blue",1
                SetLightColor w2,"blue",1
                SetLightColor w3,"blue",1
                SetLightColor w4,"blue",1
                witchDirection=-1
                ' WitchMove.Interval = 800
                witchStep = 0.25
                WitchMove.Enabled = 1
                bFreeze=1
                FreezeCnt=9
                Freeze.Interval = 3450
                Freeze.Enabled = 1
                sw17.TimerInterval = 1000
                If PUPpack=True Then DOF 87,DOFPulse Else PlaySound "Melting1"
            End If
          case 76        ' Ramp Complete
            If BTWWCnt=5 Then
                BTWWCnt = 6
                AddScore_points 10000
                For each obj in Lights_BTWW
                    SetLightColor obj, "red", 1
                Next
                smoke.visible=0
                Controller.B2SSetData "WW_08",0
                Controller.B2SSetData "WW_09",1
            End If
        End Select
    End If
End Sub

Sub Freeze_Timer
    Dim obj
    FreezeCnt = FreezeCnt - 1
    Select Case FreezeCnt
      case 8
        L63.State=0:L62.State=0:L61.State=0:L60.State=0:L59.State=0
        L40.State=0:L39.State=0:L38.State=0:L37.State=0:L34.State=0
        L252.State=0:L253.State=0:L254.State=0
        ClearOZLaneFlashers
        gi1.State=0:gi3.State=0:gi5.State=0:gi11.State=0:gi14.State=0
      case 6
        L30.State=0:L31.State=0:L32.State=0:L33.State=0:L35.State=0:L41.State=0
        L46.State=0:L47.State=0:L48.State=0:L49.State=0:L50.State=0
        L51.State=0:L52.State=0:L71.State=0:L70.State=0
        gi2.State=0:gi4.State=0:gi5.State=0:gi12.State=0:gi19.State=0
      case 5
        L27.State=0:L28.State=0:L29.State=0:L33.State=0:L35.State=0
        L54.State=0:L53.State=0:L44.State=0:L45.State=0
        L55.State=0:L56.State=0:L57.State=0
        L68.State=0:L69.State=0:L72.State=0:L73.State=0:L134.State=0
        L161.State=0:L162.State=0:L163.State=0:L164.State=0
        L109.State=0:L113.State=0:L108.State=0:L112.State=0:L110.State=0
        gi8.State=0:gi9.State=0:gi10.State=0:gi6.State=0:gi7.State=0
        ScoopLight1.State=0:ScoopLight2.State=0:ScoopLight3.State=0
      case 4
        L20.State=0:L21.State=0:L22.State=0:L23.State=0:L24.State=0
        L25.State=0:L26.State=0:L10.State=0:L11.State=0:L12.State=0:L28.State=0:L36.State=0
        L43.State=0:L77.State=0:L78.State=0:L83.State=0:L8.State=0
        L82.State=0:L76.State=0:L84.State=0:L74.State=0:L75.State=0:L65.State=0
        L106.State=0:L107.State=0:L111.State=0
        L192.State=0:L193.State=0:L194.State=0:L195.State=0:L196.State=0
        L197.State=0:L198.State=0:L199.State=0
        gi18.State=0:gi23.State=0:gi21.State=0:gi22.State=0
      case 3
        L1.State=0:L2.State=0:L3.State=0:L13.State=0:L14.State=0:L15.State=0
        L17.State=0:L18.State=0:L19.State=0:L16.State=0:L127.State=0
        L129.State=0:L9.State=0:L79.State=0:L80.State=0:L66.State=0:L81.State=0:L142.State=0
        L85.State=0:L86.State=0:L87.State=0:L97.State=0:L98.State=0:L99.State=0:L134.State=0:L64.State=0
      case 2
        L4.State=0:L5.State=0:L6.State=0:L7.State=0:L96.State=0:L141.State=0
        L32.State=0:L101.State=0:L100.State=0:L102.State=0:L90.State=0
        L88.State=0:L89.State=0:L91.State=0:L92.State=0:L93.State=0:L95.State=0:L94.State=0
        L128.State=0:L129.State=0:L130.State=0:L133.State=0:L139.State=0:L138.State=0:L42.State=0
        gi13.State=0:gi17.State=0
      case 1
        L135.State=0:L136.State=0:L137.State=0:L132.State=0:L131.State=0:L143.State=0:L58.State=0
        gi16.State=0
      case 0
        Me.Enabled = 0
        bFreeze = 0
        SetLightColor w1,"red",0
        SetLightColor w2,"red",0
        SetLightColor w3,"red",0
        SetLightColor w4,"red",0

        vpmTimer.AddTimer 1000,"DingDongDead '"

    End Select
End Sub

' DING DONG mode
Dim bDingDong, ddWhiteShots
Dim DD_sw_color(6)

Sub DingDongDead
    Dim obj
    For each obj in aLights
        SetLightColor obj, "white", 0
    Next
    ChangeGi "green"
    GiEffect 1, 50
    Controller.B2SSetData 178,0
    Controller.B2SSetData "melt_00",0
    Controller.B2SSetData "DingDong",1
    StopSound Song: Song = " "
    If PUPpack=True Then DOF 63, DOFPulse
    PlaySound "voc14_WitchIsDead"
'    LightSeqAttract.UpdateInterval=250
'    LightSeqAttract.Play SeqRandom,1,,9000
    BTWWCnt=0
    BTWW_end
    AwardJewel
    bDingDong = 1
    vpmTimer.AddTimer 10000,"DingDong '"
End Sub

Sub DingDong
    Dim obj
    LightSeqGI.StopPlay
    AddScore_points BTWWjackpot
    BallsOnPlayfield=0
    mBalls2Eject=0
    bMultiBallMode=1
    If PUPpack=True Then
        DOF 80, DOFOn
        vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
    End If
    AddMultiball 4
    ChangeGi "green"
    GiOn
    For each obj in Lights_BTWW
        SetLightColor obj, "yellow", 1
    Next
    ' Ramp - sw75
    SetLightColor l137,"white",1:SetLightColor l139,"white",1:SetLightColor l42,"white",1
    SetLightColor l143,"white",1:SetLightColor l138,"white",1
    DD_sw_color(1) = "white"
    ' Castle VUK - sw17
    SetLightColor l132,"white",1:SetLightColor l58,"white",1:SetLightColor l136,"white",1
    DD_sw_color(2) = "white"
    ' Throne - sw15
    SetLightColor l30,"white",1:SetLightColor l41,"white",1:SetLightColor l35,"white",1
    DD_sw_color(3) = "white"
    ' Left Orbit - sw64
    SetLightColor l80,"white",1:SetLightColor l98,"white",1:SetLightColor l66,"white",1
    SetLightColor l64,"white",1:SetLightColor l99,"white",1
    DD_sw_color(4) = "white"
    ' Right Orbit - sw87
    SetLightColor l25,"white",1:SetLightColor l36,"white",1:SetLightColor l26,"white",1
    DD_sw_color(5) = "white"
    ' Crystal Ball VUK - sw 9
    SetLightColor l76,"white",1:SetLightColor l65,"white",1:SetLightColor l84,"white",1
    DD_sw_color(6) = "white"
    Controller.B2SSetData "DingDong",0
    Controller.B2sSetData "DingDong1",1
    ddWhiteShots = 5000
    DD_Score
    DDScore.Interval = 1000
    DDScore.Enabled = 1
End Sub

Sub DD_Score
    D_DigitsDisplay ddWhiteShots
End Sub

Sub DDScore_Timer
    If bLFlipper=0 AND bRFlipper=0 Then DDScore.Interval = 250: Exit Sub
    ddWhiteShots = ddWhiteShots - INT(RND * 5 + 1)
    If ddWhiteShots<=0 Then
       DingDong_End
       Me.enabled = 0
       Exit Sub
    End If
    Controller.B2SSetData Ones,0
    Controller.B2SSetData Tens,0
    Controller.B2SSetData Hundreds,0
    Controller.B2SSetData Thousands,0
    DD_Score
    Me.Interval = 15
End Sub

Sub DingDong_End
    Dim obj, i, get_color
    bDingDong = 0
    bMultiBallMode=0
    mBalls2Eject=0
    AutoFireCheck.Enabled = False
    DDScore.Enabled=0
    D_DigitsClear
    Controller.B2sSetData "DingDong1",0
    If PUPpack=True Then DOF 340, DOFPulse
    StopSound "voc14_WitchIsDead"
    ChangeGi "white"
    GiOn
    For each obj in Lights_BTWW
        i = i + 1
        get_color = lcolor_dict.Item(obj)
        SetLightColor obj, get_color, BTWWliteState(i)
    Next
    TNPLH_TOTO_LightsReset
    If SOTR_capture=1 OR bSOTR=1 Then
        Kicker_TOTO.enabled=True
    Else
        Kicker_TOTO.enabled=False
    End If
    HouseSpinRestore
    If bDorothyCaptureRdy=1 AND bDorothyCaught=0 Then
        Kicker_Mag.Enabled=True
        LightState(115)=1:Flashers(flash_dict.Item(l115))=1
    End If
    If bDorothyCaptureRdy=1 AND bDorothyCaught=1 Then
        LightState(153)=1:Flashers(flash_dict.Item(l153))=1
    End If
    vpmtimer.AddTimer 500, "PlaySong ""voc18_OffToSee"" '"
End Sub

Sub DingDongMode_Check(switch)
    Select Case switch
      case 5  ' Munchkin House Entrance
        MHouse_BHandler
        Exit Sub
      case 9  ' Crystal Ball VUK
        If DD_sw_color(6)="white" Then
            DD_sw_color(6) = "green"
            SetLightColor l76,"green",1:SetLightColor l65,"green",1:SetLightColor l84,"green",1
            If DD_sw_color(1) = "green" Then
                SetLightColor l137,"white",1:SetLightColor l139,"white",1:SetLightColor l42,"white",1
                SetLightColor l143,"white",1:SetLightColor l138,"white",1:DD_sw_color(1) = "white"
            End If
        Else
            Exit Sub
        End If
      case 15  ' Throne
        If DD_sw_color(3)="white" Then
            DD_sw_color(3) = "green"
            SetLightColor l30,"green",1:SetLightColor l41,"green",1:SetLightColor l35,"green",1
            If DD_sw_color(4) = "green" Then
                SetLightColor l80,"white",1:SetLightColor l98,"white",1
                SetLightColor l66,"white",1:SetLightColor l64,"white",1
                SetLightColor l99,"white",1
                DD_sw_color(4) = "white"
            End If
        Else
            Exit Sub
        End If
      case 17  ' Castle VUK
        If DD_sw_color(2)="white" Then
            DD_sw_color(2) = "green"
            SetLightColor l132,"green",1:SetLightColor l58,"green",1:SetLightColor l136,"green",1
            If DD_sw_color(5) = "green" Then
                SetLightColor l25,"white",1:SetLightColor l36,"white",1:SetLightColor l26,"white",1
                DD_sw_color(5) = "white"
            End If
        Else
            Exit Sub
        End If
      case 64  ' Left Orbit
        If DD_sw_color(4)="white" Then
            DD_sw_color(4) = "green"
            SetLightColor l80,"green",1:SetLightColor l98,"green",1:SetLightColor l99,"green",1
            SetLightColor l66,"green",1:SetLightColor l64,"green",1
            If DD_sw_color(3) = "green" Then
                SetLightColor l30,"white",1:SetLightColor l41,"white",1:SetLightColor l35,"white",1
                DD_sw_color(3) = "white"
            End If
        Else
            Exit Sub
        End If
      case 75   ' Ramp
        If DD_sw_color(1)="white" Then
            DD_sw_color(1) = "green"
            SetLightColor l137,"green",1:SetLightColor l139,"green",1:SetLightColor l42,"green",1
            SetLightColor l143,"green",1:SetLightColor l138,"green",1
            If DD_sw_color(6) = "green" Then
                SetLightColor l76,"white",1:SetLightColor l65,"white",1:SetLightColor l84,"white",1
                DD_sw_color(6) = "white"
            End If
        Else
            Exit Sub
        End If
      case 87  ' Right Orbit
        If DD_sw_color(5)="white" Then
            DD_sw_color(5) = "green"
            SetLightColor l25,"green",1:SetLightColor l36,"green",1:SetLightColor l26,"green",1
            If DD_sw_color(2) = "green" Then
                SetLightColor l132,"white",1:SetLightColor l58,"white",1:SetLightColor l136,"white",1
                DD_sw_color(2) = "white"
            End If
        Else
            Exit Sub
        End If
      case Else
        AddScore Switch
        Exit Sub
    End Select
    AddScore_points ddWhiteShots
    ddWhiteShots = ddWhiteShots + INT(RND * 10 + 1)
    If ddWhiteShots>6999 Then ddWhiteShots = 6999
    DD_Score
End Sub

' WITCH SUPPORT
Sub WitchRaise
    witch.visible=1
    smoke.visible=1
    witchDirection=1
    WitchMove.Interval = 20
    WitchMove.Enabled = 1
    DOF 350, DOFOn      ' Gear Motor On
End Sub

Sub WitchLower
    witch.visible=1
    witchDirection=-1
    WitchMove.Interval = 40
    WitchMove.Enabled = 1
    DOF 350, DOFOn      ' Gear Motor On
End Sub

Sub WitchMove_Timer
    Select Case witchDirection
      case -1    ' Lower Witch
        If Witch.z>-120 Then
            Witch.z = Witch.z - witchStep
        Else
            Witch.z = -120
            Witch.visible=0
            smoke.visible=0
            SetLightColor w1,"red",0
            SetLightColor w2,"red",0
            SetLightColor w3,"red",0
            SetLightColor w4,"red",0
            Me.enabled=0
            bWitchDown=1
            DOF 350, DOFOff      ' Gear Motor Off
            witchStep=5
        End If
      case 1     ' Raise Witch
        If Witch.z<250 Then
            Witch.z = Witch.z + witchStep
        Else
            Me.enabled=0
            Witch.z = 250
            Witch.visible=1
            SetLightColor w1,"red",2
            SetLightColor w2,"red",2
            SetLightColor w3,"red",2
            SetLightColor w4,"red",2
            bWitchDown=0
            DOF 350, DOFOff      ' Gear Motor Off
            witchStep=5
        End If
    End Select
End Sub

' SOMEWHERE OVER THE RAINBOW (SOTR) Mode
' Similar to BTWW this mode takes full control of playfield
Dim bSOTR, bSOTR_init
Dim SOTR_1, SOTR_2, SOTR_3, SOTR_capture
Dim SOTRCnt, SOTRbonus, SOTRmultiplier, SOTR_DblColor, SOTRColor
Dim SOTR_RBW_X(3)
Dim SOTR_color(7), SOTRHitColor(7)
Dim SOTRLiteState(150)

Dim SOTRlite_dict
Set SOTRlite_dict = CreateObject("Scripting.Dictionary")

SOTRlite_dict.Add l106, "red"
SOTRlite_dict.Add l7, "red"
SOTRlite_dict.Add l65, "red"
SOTRlite_dict.Add l76, "red"

SOTRlite_dict.Add l107, "orange"
SOTRlite_dict.Add l6, "orange"
SOTRlite_dict.Add l66, "orange"
SOTRlite_dict.Add l80, "orange"
SOTRlite_dict.Add l64, "orange"
SOTRlite_dict.Add l98, "orange"

SOTRlite_dict.Add l108, "yellow"
SOTRlite_dict.Add l5, "yellow"
SOTRlite_dict.Add l58, "yellow"
SOTRlite_dict.Add l132, "yellow"

SOTRlite_dict.Add l109, "ylwgrn"
SOTRlite_dict.Add l4, "ylwgrn"
SOTRlite_dict.Add l135, "ylwgrn"

SOTRlite_dict.Add l113, "green"
SOTRlite_dict.Add l3, "green"
SOTRlite_dict.Add l139, "green"
SOTRlite_dict.Add l42, "green"

SOTRlite_dict.Add l112, "blue"
SOTRlite_dict.Add l2, "blue"
SOTRlite_dict.Add l25, "blue"
SOTRlite_dict.Add l36, "blue"

SOTRlite_dict.Add l111, "purple"
SOTRlite_dict.Add l1, "purple"
SOTRlite_dict.Add l30, "purple"
SOTRlite_dict.Add l41, "purple"

Sub SOTR_Mode
    Dim obj, i
    bSOTR = 1: bSOTR_init=0
    SOTR_1=0:SOTR_2=0:SOTR_3=0
    SOTRColor="white"
    SOTRCnt = 0: SOTRbonus = 1000: SOTRmultiplier = 1
    bFreeze=1

    SkillShotStop:skillShot_TimeLeft=0
    WitchHUStop
    HOADC_BG_clr
    Twister_ClearCounts
    HouseSpinStop
    ' PF MX Side Effect - All Off
    For i = 310 to 317
        DOF i, DOFOff
    Next
    Wizard_reset
    ClearOZLaneFlashers

    SetLightColor w1,"red",0
    SetLightColor w2,"red",0
    SetLightColor w3,"red",0
    SetLightColor w4,"red",0
    DivRight.RotateToEnd
    If bDorothyCaptureRdy=1 AND bDorothyCaught=0 Then
        Kicker_Mag.Enabled=False
        LightState(115)=0:Flashers(flash_dict.Item(l115))=0
    End If
    If bDorothyCaptureRdy=1 AND bDorothyCaught=1 Then
        LightState(153)=0:Flashers(flash_dict.Item(l153))=0
    End If

    AddScore_points 100000

    SOTR_color(1) = "red"
    SOTR_color(2) = "orange"
    SOTR_color(3) = "yellow"
    SOTR_color(4) = "ylwgrn"
    SOTR_color(5) = "green"
    SOTR_color(6) = "blue"
    SOTR_color(7) = "purple"

    For i = 1 to 7
    SOTRHitColor(i)=0
    Next

    For i = 1 to 3
    SOTR_RBW_X(i)=1
    Next

    i=0
    For each obj in Lights_BTWW
        i = i + 1
        SOTRliteState(i) = obj.State
    Next
    ' Setup Playfield Lights
    For each obj in Lights_BTWW
        SetLightColor obj, "white", 0
    Next
    Controller.B2sSetData 201,0
    ChangeGi "purple"
    If UltraDMD_Option=True Then UltraDMD_simple "SOTR.png", "", "", 30000
    StopSound Song
    If PUPpack=True Then
        DOF 60,DOFPulse
    Else
        Controller.B2sSetData "SOTR_01",1  ' Screen Image if no PUPpack
        PlaySound "voc97_SOTR_Intro1"
    End If
    LightSeqAll.Play SeqAlloff
    LightSeqAll.Play SeqCircleOutOn,10,10,0
    DOF 112, DOFPulse	' Strobe
	DOF 126, DOFPulse	' Beacon
    SOTR = 0
    SOTRImage.Interval = 12000
    SOTRImage.Enabled = True
End Sub

Sub SOTRImage_Timer
    Me.Interval = 6000
    SOTRCnt = SOTRCnt + 1
    If PUPpack=True AND SOTRcnt<4 Then Exit Sub
    Select Case SOTRCnt
      Case 1
        Controller.B2sSetData "SOTR_01",0
        Controller.B2sSetData "SOTR_02",1
      Case 2
        Controller.B2sSetData "SOTR_02",0
        Controller.B2sSetData "SOTR_03",1
      Case 3
        Controller.B2sSetData "SOTR_03",0
        Controller.B2sSetData "SOTR_04",1
        SOTR_Light_Init
      Case 4
        ClearBG_BTWW
        SOTRImage.Enabled = False
        SOTRCnt = 0
        Controller.B2sSetData "SOTR_04",0
        Controller.B2sSetData "SOTR_08",1
        Controller.B2sSetData "SOTR_X1",1
        SOTR_DblColor = 5
        Controller.B2sSetData "SOTR_Ball_" + SOTR_color(SOTR_DblColor),1
        ChangeGi SOTR_color(SOTR_DblColor)
        bFreeze=0
        bMultiBallMode=1
        AddMultiball 7
        bSOTR_init=1
        SOTR_ClearRBonus
        SOTR_RBonus
        vpmTimer.AddTimer 5000,"PlaySound ""voc8_Rainbow"", -1, 1 '"
    End Select
End Sub

Sub SOTR_Light_Init
    Dim obj
    Dim get_color

    For each obj in Lights_SOTR
        get_color = SOTRlite_dict.Item(obj)
        SetLightColor obj,get_color,2
    Next
End Sub

Sub SOTRMode_check(switch)
    Dim obj, i, tmp_str
    Dim get_light
    Dim get_color
    Dim get_idx

    ' This shouldn't happen but just in case :)
    If switch=5 Then         '  Munchkin House Entrance
        MHouse_BHandler
        Exit Sub
    End If

    ' Extra balls will drain while Instructions Play
    If bSOTR_init=0 Then Exit Sub

    AddScore switch
    SetDOF switch
    Select Case switch
      Case 101  '  Munchkin House Loop - Change color
        tmp_str = "SOTR_Ball_" & SOTR_color(SOTR_DblColor)
        Controller.B2sSetData tmp_str,0
        SOTR_DblColor = (SOTR_DBLCOLOR + 1) MOD 8
        If SOTR_DblColor=0 Then SOTR_DblColor = 1
        tmp_str = "SOTR_Ball_" & SOTR_color(SOTR_DblColor)
        Controller.B2sSetData tmp_str,1
        ChangeGi SOTR_color(SOTR_DblColor)
      Case 65, 66, 67, 68, 69, 70, 71   ' CASTLE Rainbow
        Set get_light = light_dict.Item(switch)
        get_light.State = 1
        If switch=67 Then SOTRColor="red":i=1
        If switch=66 Then SOTRColor="orange":i=2
        If switch=65 Then SOTRColor="yellow":i=3
        If switch=71 Then SOTRColor="ylwgrn":i=4
        If switch=68 Then SOTRColor="green":i=5
        If switch=69 Then SOTRColor="blue":i=6
        If switch=70 Then SOTRColor="purple":i=7
        SOTRHitColor(i) = SOTRHitColor(i) + BasePoints(switch): AddScore_points SOTRHitColor(i)
        SOTRBonus = SOTRBonus + SOTRHitColor(i)
        If SOTRColor=SOTR_color(SOTR_DblColor) Then
            SOTRBonus = SOTRBonus + SOTRmultiplier * BasePoints(switch) * 2 + SOTRHitColor(i)
            AddScore switch
            AddScore_points SOTRHitColor(SOTR_DblColor)
        Else
            SOTRBonus = SOTRBonus + SOTRmultiplier * BasePoints(switch)
        End If
        SOTR_RBonus
        If L106.State=1 AND L107.State=1 AND L108.State=1 AND L109.State=1 AND _
                                  L111.State=1 AND L112.State=1 AND L113.State=1 Then
            ' Blink RESCUE & Loop Lights
            SetLightColor l106,"red",2
            SetLightColor l107,"orange",2
            SetLightColor l108,"yellow",2
            SetLightColor l109,"ylwgrn",2
            SetLightColor l113,"green",2
            SetLightColor l112,"blue",2
            SetLightColor l111,"purple",2
            AddScore_points (SOTRBonus * SOTRmultiplier)
            SOTR_RBW_X(1) = SOTR_RBW_X(1) + 1
            Controller.B2sSetData "SOTR_RBW_01",1
            If SOTR_RBW_X(1)<=SOTR_RBW_X(2) AND SOTR_RBW_X(1)<=SOTR_RBW_X(3) Then
                AddScore_points SOTRBonus
                SOTRmultiplier = SOTRmultiplier + 1
                If SOTRmultiplier<5 Then
                    tmp_str = "SOTR_X" & Cstr(SOTRmultiplier - 1)
                    Controller.B2sSetData tmp_str,0
                    tmp_str = "SOTR_X" & Cstr(SOTRmultiplier)
                    Controller.B2sSetData tmp_str,1
                End If
                vpmTimer.AddTimer 2500,"SOTR_3Rainbows '"
            End If
            SOTR_Rainbow_Fanfare
        End If
      Case 89, 90, 91, 92, 93, 94, 95    ' TARGET Rainbow
        Set get_light = light_dict.Item(switch)
        get_color = lcolor_dict.Item(get_light)
        SetLightColor get_light,get_color,1
        get_idx = switch - 88
        SOTRHitColor(get_idx) = SOTRHitColor(get_idx) + BasePoints(switch): AddScore_points SOTRHitColor(get_idx)
        SOTRBonus = SOTRBonus + SOTRHitColor(i)
        If get_color=SOTR_color(SOTR_DblColor) Then
            SOTRBonus = SOTRBonus + SOTRmultiplier * BasePoints(switch) * 2 + SOTRHitColor(i)
            AddScore switch
            AddScore_points SOTRHitColor(SOTR_DblColor)
        Else
            SOTRBonus = SOTRBonus + SOTRmultiplier * BasePoints(switch)
        End If
        SOTR_RBonus
        If L1.State=1 AND L2.State=1 AND L3.State=1 AND L4.State=1 AND _
                                 L5.State=1 AND L6.State=1 AND L7.State=1 Then
            ' Blink Rainbow Lights
            For i = 89 To 95
                Set obj = light_dict.Item(i)
                get_color = lcolor_dict.Item(obj)
                SetLightColor obj, get_color, 2
            Next
            AddScore_points (SOTRBonus * SOTRmultiplier)
            SOTR_RBW_X(2) = SOTR_RBW_X(2) + 1
            Controller.B2sSetData "SOTR_RBW_02",1
            If SOTR_RBW_X(2)<=SOTR_RBW_X(1) AND SOTR_RBW_X(2)<=SOTR_RBW_X(3) Then
                AddScore_points SOTRBonus
                SOTRmultiplier = SOTRmultiplier + 1
                If SOTRmultiplier<5 Then
                    tmp_str = "SOTR_X" & Cstr(SOTRmultiplier - 1)
                    Controller.B2sSetData tmp_str,0
                    tmp_str = "SOTR_X" & Cstr(SOTRmultiplier)
                    Controller.B2sSetData tmp_str,1
                End If
                vpmTimer.AddTimer 2500,"SOTR_3Rainbows '"
            End If
            SOTR_Rainbow_Fanfare
        End If
      Case 9, 15, 16, 17, 24, 73, 74, 76    ' SHOTS Rainbow
        If switch=9 Then SetLightColor l65,"red",1: SetLightColor l76,"red",1: SOTRColor="red":i=1
        If switch=24 Then SetLightColor l80,"orange",1: SetLightColor l66,"orange",1:SetLightColor l98,"orange",1:SetLightColor l64,"orange",1: SOTRColor="orange":i=2
        If switch=73 OR SWITCH=17 Then SetLightColor l132,"yellow",1:SetLightColor l58,"yellow",1: SOTRColor="yellow":i=3
        If switch=74 Then SetLightColor l135,"ylwgrn",1: SOTRColor="ylwgrn":i=4
        If switch=76 Then SetLightColor l42,"green",1: SetLightColor l139,"green",1:SOTRColor="green":i=5
        If switch=16 Then SetLightColor l25,"blue",1: SetLightColor l36,"blue",1: SOTRColor="blue":i=6
        If switch=15 Then SetLightColor l30,"purple",1:SetLightColor l41,"purple",1: SOTRColor="purple":i=7
        SOTRHitColor(i) = SOTRHitColor(i) + BasePoints(switch): AddScore_points SOTRHitColor(i)
        SOTRBonus = SOTRBonus + SOTRHitColor(i)
        If SOTRColor=SOTR_color(SOTR_DblColor) Then
            SOTRBonus = SOTRBonus + SOTRmultiplier * 210 + SOTRHitColor(i)
            AddScore_points 105  ' Average of all seven
            AddScore_points SOTRHitColor(SOTR_DblColor)
        Else
            SOTRBonus = SOTRBonus + SOTRmultiplier * 105
        End If
        SOTR_RBonus
        If L76.State=1 AND L80.State=1 AND L98.State=1 AND L132.State=1 AND L135.State=1 AND _
                                 L139.State=1 AND L25.State=1 AND L30.State=1 Then
            ' Blink Shot Lights
            SetLightColor l65,"red",2: SetLightColor l76,"red",2
            SetLightColor l80,"orange",2: SetLightColor l66,"orange",2:SetLightColor l98,"orange",2:SetLightColor l64,"orange",2
            SetLightColor l132,"yellow",2:SetLightColor l58,"yellow",2
            SetLightColor l135,"ylwgrn",2
            SetLightColor l42,"green",2:SetLightColor l139,"green",2
            SetLightColor l25,"blue",2: SetLightColor l36,"blue",2
            SetLightColor l30,"purple",2:SetLightColor l41,"purple",2
            AddScore_points (SOTRBonus * SOTRmultiplier)
            SOTR_RBW_X(3) = SOTR_RBW_X(3) + 1
            Controller.B2sSetData "SOTR_RBW_03",1
            If SOTR_RBW_X(3)<=SOTR_RBW_X(1) AND SOTR_RBW_X(3)<=SOTR_RBW_X(2) Then
                AddScore_points SOTRBonus
                SOTRmultiplier = SOTRmultiplier + 1
                If SOTRmultiplier<5 Then
                    tmp_str = "SOTR_X" & Cstr(SOTRmultiplier - 1)
                    Controller.B2sSetData tmp_str,0
                    tmp_str = "SOTR_X" & Cstr(SOTRmultiplier)
                    Controller.B2sSetData tmp_str,1
                End If
                vpmTimer.AddTimer 2500,"SOTR_3Rainbows '"
            End If
            SOTR_Rainbow_Fanfare
        End If
      Case Else
        AddScore Switch
    End Select
End Sub

Sub SOTR_ClearRBonus
    D_DigitsClear
End Sub

Sub SOTR_RBonus
    Controller.B2SSetData Ones,0
    Controller.B2SSetData Tens,0
    Controller.B2SSetData Hundreds,0
    Controller.B2SSetData Thousands,0
    Controller.B2SSetData TenK,0
    D_DigitsDisplay SOTRBonus
    SOTR_RBonusDisplay
End Sub

Sub SOTR_RBonusDisplay
    Dim tmp0, tmp1, tmp2, tmp3
    If DMD_Option=True Then
        DMD2Flush
        if(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "  RAINBOW BONUS"
            tmp2 = FillLine(2, "POINTS", FormatScore(SOTRBonus))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
    End If
End Sub


Sub SOTR_Rainbow_Fanfare
    AddScore_points 100000
    If PUPpack=True Then DOF 82,DOFPulse
    LightSeqAll.Play SeqAlloff
    LightSeqAll.Play SeqCircleOutOn,10,5,0
    DOF 112, DOFPulse	' Strobe
	DOF 126, DOFPulse	' Beacon
End Sub

Sub SOTR_3Rainbows
    Dim i, tmp_str
    ' The game should go bonkers at this point
    If PUPpack=True Then DOF 58,DOFPulse
    LightSeqAll.StopPlay
    AddScore_points 500000
    For i = 1 To 7
        AddScore_points (SOTRHitColor(i) * SOTRmultiplier)
    Next
    AwardSpecial
    For i = 1 to 3
       tmp_str = "SOTR_RBW_0" & Cstr(i)
       Controller.B2SSetData tmp_str,0
    Next
End Sub

Sub SOTR_end
    Dim obj, i, get_color, tmp_str
    bSOTR=0
    SOTR_1=0:SOTR_2=0:SOTR_3=0
    WitchLower
    bMultiBallMode=0
    mBalls2Eject=0
    i=0
    For each obj in Lights_BTWW
        i = i + 1
        get_color = lcolor_dict.Item(obj)
        SetLightColor obj, get_color, SOTRliteState(i)
    Next
    TNPLH_TOTO_LightsReset
    Wizard_restore
    HouseSpinRestore
    If bDorothyCaptureRdy=1 AND bDorothyCaught=0 Then
        Kicker_Mag.Enabled=True
        LightState(115)=1:Flashers(flash_dict.Item(l115))=1
    End If
    If bDorothyCaptureRdy=1 AND bDorothyCaught=1 Then
        LightState(153)=1:Flashers(flash_dict.Item(l153))=1
    End If

    If PUPpack=True Then DOF 62,DOFPulse
    SOTR_ClearRBonus
    ClearBG_SOTR
    If l134.State=0 Then
        DivRight.RotateToStart
    Else
        DivRight.RotateToEnd
    End If
    ChangeGi "white"
    GiOn
    sw73.IsDropped=0:SetLightsOff(73)
    sw17.TimerInterval = 1000
    StopSound "voc8_Rainbow"
    StopSound "voc96_SOTR_IntroAdd"
    vpmtimer.AddTimer 500, "PlaySong ""voc18_OffToSee"" '"
    SOTRCnt=0: bSOTR_init=0
End Sub


' CRYSTAL BALL Mode
Dim bCrystalBallMode
Dim bCB_LightsON, bCB_LightsOFF, bCB_LinkedFlip, bCB_NoHoldFlip, bCB_FlipFrenzy
Dim crystalBallModeCnt, cb_Color
Dim CBBonus, CBMB_Bonus


Sub CrystalBallMode
    If bCB_LightsON=1 OR bCB_LightsOFF=1 OR bCB_LinkedFlip=1 OR bCB_NoHoldFlip=1 OR bCB_FlipFrenzy=1 Then Exit Sub
    bCrystalBallMode=1
    Controller.B2sSetData 244,1
    Select Case crystalBallModeCnt
      case 1
        cb_Color = "blue"
      case 2
        cb_Color = "yellow"
      case 3
        cb_Color = "green"
      case 4
        cb_Color = "red"
      case 5
        cb_Color = "purple"
    End Select
    SetLightColor l72,cb_Color,2:SetLightColor l73,cb_Color,2:SetLightColor l74,cb_Color,2:SetLightColor l75,cb_Color,2
    If PUPpack=True AND bMultiBallMode=0 Then
        DOF 85,DOFpulse
        vpmTimer.AddTimer 8000,"Controller.B2SSetData 244,0 '"
    Else
        PlaySound "voc59_TakeWarning"
        vpmTimer.AddTimer 4000,"Controller.B2SSetData 244,0 '"
    End If
End Sub

Sub CB_LightsOFF
    Multipliers(6) = 2
    SuperXcolor Multipliers(6)
    Controller.B2sSetData 244,0
    Controller.B2SSetData "CrystalBall_OFF",1
    bCB_LightsOFF=1
    LightEffect 0
    CrystalBallLights.Enabled = 1
    bMultiBallMode=1
    If PUPpack=True Then
        DOF 80, DOFOn
        vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
    End If
    GiOff
    cb_Color = "yellow"
    AddMultiball 1
    PlaySound "voc60_TakeCareYouNow"
End Sub

Sub CB_LightsON
    Multipliers(6) = 2
    SuperXcolor Multipliers(6)
    Controller.B2sSetData 244,0
    Controller.B2SSetData "CrystalBall_ON",1
    bCB_LightsON=1
    LightEffect 5
    CrystalBallLights.Enabled = 1
    bMultiBallMode=1
    If PUPpack=True Then
        DOF 80, DOFOn
        vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
    End If
    ChangeGi "yellow"
    cb_Color = "green"
    AddMultiball 1
    PlaySound "voc60_TakeCareYouNow"
End Sub

Sub CB_LinkedFlip
    Multipliers(6) = 2
    SuperXcolor Multipliers(6)
    Controller.B2sSetData 244,0
    Controller.B2SSetData "CrystalBall_Linked",1
    bCB_LinkedFlip=1
    bMultiBallMode=1
    If PUPpack=True Then
        DOF 80, DOFOn
        vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
    End If
    ChangeGi "green"
    cb_Color = "red"
    AddMultiball 1
    PlaySound "voc60_TakeCareYouNow"
End Sub

Sub CB_NoHoldFlip
    Multipliers(6) = 2
    SuperXcolor Multipliers(6)
    Controller.B2sSetData 244,0
    Controller.B2SSetData "CrystalBall_NoHold",1
    bCB_NoHoldFlip=1
    bMultiBallMode=1
    If PUPpack=True Then
        DOF 80, DOFOn
        vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
    End If
    ChangeGi "red"
    cb_Color = "purple"
    AddMultiball 1
    PlaySound "voc60_TakeCareYouNow"
End Sub


Sub CB_FlipFrenzy
    Multipliers(6) = 3
    SuperXcolor Multipliers(6)
    Controller.B2sSetData 244,0
    Controller.B2SSetData "CrystalBall_Frenzy",1
    bCB_FlipFrenzy=1
    AwardJewel
    bMultiBallMode=1
    If PUPpack=True Then
        DOF 80, DOFOn
        vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
    End If
    ChangeGi "purple"
    cb_Color = "blue"
    AddMultiball 1
    PlaySound "voc60_TakeCareYouNow"
End Sub

Sub CrystalBallModeTimer_Timer
    Dim i
    Dim obj
    Multipliers(6) = 1
    SuperXcolor Multipliers(6)
    Me.enabled=0
    CrystalBallLights.enabled = 0
    bCrystalBallMode=0
    If bCB_LightsON=1 OR bCB_LightsOFF=1 Then LightSeqInserts.StopPlay
    If bCB_LightsOFF=1 Then ChangeGi "white": GiOn
    bCB_LightsON=0: bCB_LightsOFF=0: bCB_LinkedFlip=0: bCB_NoHoldFlip=0: bCB_FlipFrenzy=0
    Controller.B2sSetData 244,0
    Controller.B2SSetData "CrystalBall_ON",0
    Controller.B2SSetData "CrystalBall_OFF",0
    Controller.B2SSetData "CrystalBall_Linked",0
    Controller.B2SSetData "CrystalBall_NoHold",0
    Controller.B2SSetData "CrystalBall_Frenzy",0
    Controller.B2sSetData 202,1
    ' Reset Crystal Ball Switches to zero
    For i = 29 To 32
        SwitchFlags(i)=0
    Next
    For each obj in Lights_CrystalBall
        SetLightColor obj,cb_Color,0
    Next
    HauntedLights
End Sub

Sub CyrstalBallLights_Timer
    LightSeqInserts.StopPlay
    If bCB_LightsON=1 Then vpmTimer.AddTimer 1000,"LightEffect 5 '"
    If bCB_LightsOFF=1 Then
        vpmTimer.AddTimer 1000,"LightEffect 0 '"
        vpmTimer.AddTimer 1500,"ChangeGi ""white"" '"
        ' "COWARDLY" Lights - Always ON
        For each obj in acowardly
            SetLightColor obj, "green", 1
        Next
    End If
End Sub


' HAUNTED Mode
Dim bHauntedMode, bHauntedLtrMode
Dim bHauntedShots, bHauntedTargets, bHauntedHoles, bHauntedBumpers, bHauntedMB
Dim hauntedLtrCnt, hauntedModeCnt
Dim hauntedTimeLeft, hauntedBonus, hauntedMB_Bonus
Dim hauntedShots, hauntedTargets, hauntedHoles, hauntedBumpers, hauntedIncrease

Dim modeHaunted
Set modeHaunted = CreateObject("System.Collections.ArrayList")

modeHaunted.Add 49
modeHaunted.Add 50
modeHaunted.Add 51
modeHaunted.Add 53

Sub HauntedLights
'    If bHauntedMode=1 Then Exit Sub
    Dim obj
    Dim bg_id
    For each obj in Lights_Haunted
        SetLightColor obj,"purple",0
    Next
    hauntedLtrCnt = SwitchFlags(49) + SwitchFlags(50) + SwitchFlags(51)
    If hauntedLtrCnt>0 Then SetLightColor l86,"purple",1
    If hauntedLtrCnt>1 Then SetLightColor l87,"purple",1
    If hauntedLtrCnt>2 Then SetLightColor l88,"purple",1
    If hauntedLtrCnt>3 Then SetLightColor l89,"purple",1
    If hauntedLtrCnt>4 Then SetLightColor l91,"purple",1
    If hauntedLtrCnt>5 Then SetLightColor l92,"purple",1
    If hauntedLtrCnt>6 Then
        SetLightColor l93,"purple",1
        SetLightColor l95,"red",2
        SetLightColor l97,"purple",2
        SetLightColor l142,"red",1
        bHauntedLtrMode=1
    End If
    HauntedLightsCheck.Enabled=True
End Sub

Sub HauntedLightsCheck_Timer
    Dim obj
    Dim bg_id
    Me.enabled=0
    For each obj in Lights_Haunted
        bg_id = bg_dict.Item(obj)
        If obj.State=0 Then
            Controller.B2SSetData bg_id, 0
        Else
            Controller.B2SSetData bg_id, 1
        End If
    Next
End Sub

Sub HauntedModes
    Dim obj
    Dim h_title, hits_score
    bHauntedMode=1
    bHauntedLtrMode=0
    hauntedModeCnt = hauntedModeCnt + 1
    If bCB_LightsOFF=0 AND bCB_LightsON=0 Then ChangeGi "purple"
    SwitchFlags(49)=0:SwitchFlags(50)=0:SwitchFlags(51)=0
    For each obj in Lights_Haunted
        SetLightColor obj,"purple",2
    Next
    SetLightColor l95,"green",2
    SetLightColor l97,"purple",0
    SetLightColor l142,"red",0
    ClearBG_Haunted
    Controller.B2SSetData 246,1
    hauntedTimeLeft=31
    HauntedModeTimer.interval = 1000
    HauntedModeTimer.enabled = True
    Display_MX "HAUNTED"
    Select Case hauntedModeCnt
      case 1  ' Shots
        Haunted_Shots
        h_title = " HAUNTED SHOTS": hits_score = hauntedShots
      case 2  ' Targets
        Haunted_Targets
        h_title = "HAUNTED TARGETS": hits_score = hauntedTargets
      case 3  ' Holes
        Haunted_Holes
        h_title = " HAUNTED HOLES": hits_score = hauntedHoles
      case 4  ' Bumpers
        Haunted_Bumpers
        h_title = "HAUNTED BUMPERS": hits_score = hauntedBumpers
      case 5  ' Haunted MultiBall
        hauntedModeCnt=0
        Haunted_MB
    End Select
    HauntedDMD2Display h_title, hits_score
End Sub

Sub Haunted_Shots
    Controller.B2SSetData 245,1
    Controller.B2SSetData 246,1
    If PUPpack=True Then DOF 65, DOFPulse
    PlaySound "voc43_Spooks"
    bHauntedShots=1
    hauntedShots = 500: hauntedIncrease = 50
End Sub

Sub Haunted_Targets
    Controller.B2SSetData 245,1
    Controller.B2SSetData 246,0
    Controller.B2SSetData "HF_Shots_Done",1
    Controller.B2SSetData 247,1
    If PUPpack=True Then DOF 66, DOFPulse
    PlaySound "voc48_ThatsSilly"
    bHauntedTargets=1
    hauntedTargets = 250: hauntedIncrease = 25
End Sub

Sub Haunted_Holes
    Controller.B2SSetData 245,1
    Controller.B2SSetData 247,0
    Controller.B2SSetData "HF_Shots_Done",1
    Controller.B2SSetData "HF_Targets_Done",1
    Controller.B2SSetData 248,1
    If PUPpack=True Then DOF 67, DOFPulse
    PlaySound "voc49_DontYouBelieve"
    bHauntedHoles=1
    hauntedHoles = 750: hauntedIncrease = 75
End Sub

Sub Haunted_Bumpers
    Controller.B2SSetData 245,1
    Controller.B2SSetData 248,0
    Controller.B2SSetData "HF_Shots_Done",1
    Controller.B2SSetData "HF_Targets_Done",1
    Controller.B2SSetData "HF_Holes_Done",1
    Controller.B2SSetData 249,1
    If PUPpack=True Then DOF 68, DOFPulse
    PlaySound "voc50_IDoIDoIDo"
    bHauntedBumpers=1
    hauntedBumpers = 100: hauntedIncrease = 10
End Sub

Sub Haunted_MB
    HauntedModeTimer.enabled = False
    Controller.B2SSetData 245,0
    Controller.B2SSetData 249,0
    Controller.B2SSetData "HF_Shots_Done",0
    Controller.B2SSetData "HF_Targets_Done",0
    Controller.B2SSetData "HF_Holes_Done",0
    Controller.B2SSetData "HF_Bumpers_Done",0
    Controller.B2SSetData 250,1
    PlaySound "voc51_WhoIs"
    SetLightColor l97,"purple",2
    SetLightColor l142,"red",1
    AwardJewel
    bHauntedMB=1
    bMultiBallMode=1
    If PUPpack=True Then
        DOF 302, DOFPulse
        DOF 80, DOFOn
        vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
    End If:
    AddMultiball 2
End Sub

Dim H_Ones, H_Tens

Sub HauntedModeTimer_Timer
    If hauntedTimeLeft>0 Then hauntedTimeLeft = hauntedTimeLeft - 1
    Controller.B2SSetData H_Ones,0:Controller.B2SSetData H_Tens,0
    If hauntedTimeLeft>9 Then H_Tens = "Rd_" & CStr(INT(hauntedTimeLeft/10) MOD 10):Controller.B2SSetData H_Tens,1
    H_Ones = "R1_" & CStr(hauntedTimeLeft MOD 10):Controller.B2SSetData H_Ones,1
    If hauntedTimeLeft=0 Then
        Me.enabled=0
        bHauntedMode=0
        bHauntedShots=0:bHauntedTargets=0:bHauntedHoles=0:bHauntedBumpers=0:bHauntedMB=0
        SetLightColor l95,"red",0
        Controller.B2SSetData H_Ones,0
        ClearBG_HauntedModes
        If bMultiBallMode=0 AND bMunchkinMode=0 AND bCrystalBallMode=0 Then ChangeGI "white"
        If bHU=0 AND DMD_Option=True Then DMD2_Clear
        Controller.B2sSetData 202,1
        SwitchFlags(49)=0:SwitchFlags(50)=0:SwitchFlags(51)=0
        HauntedLights
    End If
End Sub

Sub HauntedDMD2Display(h_title, hits_score)
    Dim tmp0, tmp1, tmp2, tmp3
    If bHU=0 AND DMD_Option=True Then
        DMD2Flush
        If (dqHead2 = dqTail2) Then
            tmp0 = ""
            tmp1 = h_title
            tmp2 = FillLine(2, "HITS SCORE",FormatScore(hits_score))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
    End If
End Sub


' YELLOW BRICK ROAD Mode
Dim modeYBR_cnt, YBRBonus

Sub YBR_Setup
    Dim obj
    For each obj in Lights_YBR
        SetLightColor obj,"yellow",0
    Next
    For each obj in Lights_YBR_Start
        SetLightColor obj,"yellow",1
    Next
End Sub

Sub YBR_ChooseShot
    SetLightColor L27,"yellow",0:SetLightColor L100,"yellow",0:SetLightColor l143,"yellow",0
    Select Case INT(RND * 3 + 1)
      case 1
        SetLightColor L27,"yellow",2
      case 2
        SetLightColor L100,"yellow",2
      case 3
        SetLightColor l143,"yellow",2
    End Select
End Sub

Sub AdvanceYBR
    Dim obj
    Dim tmp0, tmp1, tmp2, tmp3
    modeYBR_cnt = modeYBR_cnt + 1
    If bHU=0 AND DMD_Option=True Then
        DMD2Flush
        If(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "  YELLOW BRICK"
            tmp2 = FillLine(2, "ROAD ADVANCE", " X" & FormatScore(modeYBR_cnt))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
        vpmTimer.AddTimer 3000,"DMD2_Clear '"
    End If
    For each obj in Lights_YBR_Start
       SetLightColor obj,"yellow",1
    Next
    Select Case modeYBR_cnt
      case 1
         PlaySound "voc29_followtheyellow"
      case 10    ' Extra Ball
        AwardExtraBall
      case 15
         PlaySound "voc29_followtheyellow"
      case 20    ' Big Points - 10000
        Score(CurrentPlayer) = Score(CurrentPlayer) + 10000
        PlaySound "voc53_YBR"
      case 30    ' Lights special or extra ball
        If bFreePlay=0 Then
            SetLightColor l31,"orange",2
        Else
            SetLightColor l29,"orange",2
        End If
        PlaySound "voc53_YBR"
      case 35
         PlaySound "voc29_followtheyellow"
      case 40    ' Big Points - 20000
        Score(CurrentPlayer) = Score(CurrentPlayer) + 20000
        PlaySound "voc53_YBR"
      case 50    ' Jewel towards SOTR and Extra Ball
        AwardExtraBall
        AwardJewel
      case Else
        PlaySound "voc53_YBR"
    End Select
    Select Case (modeYBR_cnt MOD 10)
      case 0
        For each obj in Lights_YBR
           SetLightColor obj,"yellow",1
        Next
      case 1
        For each obj in Lights_YBR
           SetLightColor obj,"yellow",0
        Next
        SetLightColor L51,"yellow",1
      case 2
        SetLightColor L50,"yellow",1
      case 3
        SetLightColor L47,"yellow",1
      case 4
        SetLightColor L46,"yellow",1
      case 5
        SetLightColor L45,"yellow",1
      case 6
        SetLightColor L44,"yellow",1
      case 7
        SetLightColor L43,"yellow",1
      case 8
        SetLightColor L9,"yellow",1
      case 9
        SetLightColor L129,"yellow",1
    End Select
    If UltraDMD_Option=True Then UltraDMD_simple "black.png", "YELLOW BRICK ROAD", "x " &Cstr(modeYBR_cnt), 1500
End Sub


' TWISTER Mode
Dim bTwisterFlag, bTwistAgain, twisterloops, twisterincrement, twisterbonus
Dim loopsleft, twisterModeCnt
Dim L_Ones, L_Tens, Y_Ones, Y_Tens
Dim bHouseSpin, bHouseSpinState
Dim hStep
Dim ballsOnMunchkinPF

Dim modeTwister
Set modeTwister = CreateObject("System.Collections.ArrayList")

modeTwister.Add 74
modeTwister.Add 89
modeTwister.Add 90
modeTwister.Add 91
modeTwister.Add 92
modeTwister.Add 93
modeTwister.Add 94
modeTwister.Add 95


Sub HouseSpin_Timer
    hStep = (hStep + 1) MOD 72
    house.RotY = hstep*(-5)
    If bHouseSpin=0 AND hstep=0 Then
        Me.enabled = 0
        DOF 351, DOFOff      ' Gear Motor Off
    End If
End Sub

Sub HouseSpinStop
    bHouseSpinState=bHouseSpin
    If bHouseSpin=1 Then
        bHouseSpin=0
        TwisterModeHold.enabled = False
    End If
End Sub

Sub HouseSpinRestore
    If bHouseSpinState=1 Then
        bHouseSpin=1
        HouseSpin.Enabled = True
        DOF 351, DOFOn      ' Gear Motor On
        Display_MX "TWISTER"
        TwisterModeHold.Enabled = True
    End If
End Sub

Sub TwisterModeHold_Timer
    Dim i, obj, get_color
    ClearBG_Twister
    DivRight.RotateToStart: L134.State=0:Controller.B2SSetData 233,0
    ' Rainbow Lights Off
    For each obj in Lights_Rainbow
        get_color = lcolor_dict.Item(obj)
        SetLightColor obj, get_color, 0
    Next
    For i = 181 to 187
        Controller.B2SSetData i,0
    Next
    SetLightColor l96,"white",0
    SwitchFlags(89)=0:SwitchFlags(90)=0:SwitchFlags(91)=0:SwitchFlags(92)=0
    SwitchFlags(93)=0:SwitchFlags(94)=0:SwitchFlags(95)=0:SwitchFlags(76)=0
    bHouseSpin = 0
    bTwisterFlag = 0
    bTwistAgain = 0
    twisterloops = 0
    RestoreBG_Rainbow
    Me.enabled = 0
End Sub

' MUNCHKIN Modes
Dim bMunchkinMode, bMunchkinLand, bMunchkinFrenzy, bMunchkinLollipop, bMunchkinMultiBall, bMunchkinMB, bOnMunchkinPF
Dim munchkinMB_Rcvd
Dim bLollipop, bLullaby
Dim munchkinTimeLeft, munchkinModeCnt, munchkinBonus, munchkinMB_Bonus, munchkinX
Dim lullabyX, lollipopX, lollipopCnt

Sub MunchkinMode
    Dim obj
    ClearBG_Twister
    DivRight.RotateToStart: L134.State=0:Controller.B2SSetData 233,0
    ' Rainbow Lights White
    For each obj in Lights_Rainbow
        SetLightColor obj,"white",1
    Next
    If bCB_LightsOFF=0 AND bCB_LightsON=0 Then ChangeGi "red"
    SwitchFlags(89)=0:SwitchFlags(90)=0:SwitchFlags(91)=0:SwitchFlags(92)=0
    SwitchFlags(93)=0:SwitchFlags(94)=0:SwitchFlags(95)=0:SwitchFlags(76)=0
    bHouseSpin = 0
    bTwisterFlag = 0
    bTwistAgain = 0
    twisterloops = 0
    bBTWW_4 = 1
    Controller.B2sSetData "Quadrant_Lt4",1
    munchkinModeCnt = munchkinModeCnt + 1
    munchkinTimeLeft=61
    MunchkinModeTimer.interval = 1000
    MunchkinModeTimer.enabled = True
    Select Case munchkinModeCnt
      case 1
        WelcomeMunchkinLand
      case 2
        LollipopLullaby
      case 3
        MunchkinFrenzy
      case 4
        MunchkinMultiBall
    End Select
End Sub

Sub WelcomeMunchkinLand
    Controller.B2SSetData 237,1
    Display_MX "MUNCHKIN"
    If PUPpack=False Then PlaySound "voc31_WelcomeRegally"
    If PUPpack=True AND bMultiBallMode=0 Then
        DOF 84, DOFPulse
    Else
        If PUPpack=True AND bMultiBallMode=1 AND bWitchHU=0 Then
            DOF 74, DOFPulse   ' Short Video
            vpmTimer.AddTimer 1000,"PlaySound ""voc31_WelcomeRegally"" '"
        End If
    End If
    SetLightColor l164,"red",2
    SetLightColor l96,"white",0
    bMunchkinLand=1
    ChangeGi "red"
End Sub

Sub MunchkinFrenzy
    Controller.B2SSetData 238,1
    Display_MX "MUNCHKIN"
    If PUPpack=False Then PlaySound "voc52_FrenzyIntro"
    If PUPpack=True AND bMultiBallMode=0 Then
        DOF 77, DOFPulse
    Else
        If PUPpack=True AND bMultiBallMode=1 AND bWitchHU=0 Then
            DOF 76, DOFPulse   ' Short Video
            vpmTimer.AddTimer 1000,"PlaySound ""voc52_FrenzyIntro"" '"
        End If
    End If
    SetLightColor l163,"blue",2
    SetLightColor l96,"white",0
    bMunchkinFrenzy=1
    ChangeGi "red"
End Sub

Sub LollipopLullaby
    Dim tmp0, tmp1, tmp2, tmp3
    Controller.B2SSetData 239,1
    Display_MX "MUNCHKIN"
    If PUPpack=False Then PlaySound "voc32_LollipopGuild"
    If PUPpack=True AND bMultiBallMode=0 Then
        DOF 83, DOFPulse
    Else
        If PUPpack=True AND bMultiBallMode=1 AND bWitchHU=0 Then
            DOF 73, DOFPulse   ' Short Video
            vpmTimer.AddTimer 1000,"PlaySound ""voc32_LollipopGuild"" '"
        End If
    End If
    SetLightColor l162,"green",2
    SetLightColor l96,"white",0
    bMunchkinLollipop=1
    ChangeGi "red"
    lullabyX = 1: lollipopX = 1
    lollipopCnt = 0
    If INT(RND + 0.5)=1 Then
        bLollipop = 1: bLullaby = 0
    Else
        bLollipop = 0: bLullaby = 1
    End If
    LollipopLullabyLights
    If DMD_Option=True Then
        DMD2Flush
        If(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "LOLLIPOP LULLABY"
            tmp2 = "SHOOT FLASHING"
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2000, True, ""
        End If
        vpmTimer.AddTimer 3000,"DMD2_Clear '"
    End If
    If UltraDMD_Option=True Then UltraDMD_simple "Lollipop.png", "", "LOLLIPOP LULLABY", 5000
End Sub

Sub LollipopLullabyLights
    If bLollipop=1 Then
        SetLightColor L76,"blue",2: SetLightColor L65,"blue",2
        SetLightColor L98,"blue",2: SetLightColor L64,"blue",2
        SetLightColor L80,"blue",2: SetLightColor L66,"blue",2
        SetLightColor L132,"blue",2: SetLightColor L58,"blue",2
        SetLightColor L25,"pink",1: SetLightColor L36,"pink",1
        SetLightColor L30,"pink",1: SetLightColor L41,"pink",1
        SetLightColor L139,"pink",1: SetLightColor L42,"pink",1
    Else
        SetLightColor L25,"blue",2: SetLightColor L36,"blue",2
        SetLightColor L30,"blue",2: SetLightColor L41,"blue",2
        SetLightColor L139,"blue",2: SetLightColor L42,"blue",2
        SetLightColor L76,"pink",1: SetLightColor L65,"pink",1
        SetLightColor L98,"pink",1: SetLightColor L64,"pink",1
        SetLightColor L80,"pink",1: SetLightColor L66,"pink",1
        SetLightColor L132,"pink",1: SetLightColor L58,"pink",1
    End If
End Sub

Sub Lollipop_Update(switch)
    Select Case switch
      case 89, 90, 91, 92, 93, 94, 95
        lollipopCnt = lollipopCnt + 1
      case 9, 64, 73
        If bLollipop=1 Then
            lollipopX = lollipopX + 1
            munchkinBonus = munchkinBonus + (1000 + 100 * lollipopCnt) * lollipopX
        Else
            lollipopX = 1: bLollipop = 1: bLullaby = 0
            munchkinBonus = munchkinBonus + (1000 + 100 * lollipopCnt) * 0.5
            LollipopLullabyLights
        End If
      case 15, 76, 87
        If bLullaby=1 Then
            lullabyX = lullabyX + 1
            munchkinBonus = munchkinBonus + (1000 + 100 * lollipopCnt) * lullabyX
        Else
            lullabyX = 1: bLullaby = 1: bLollipop = 0
            munchkinBonus = munchkinBonus + (1000 + 100 * lollipopCnt) * 0.5
            LollipopLullabyLights
        End If
    End Select
End Sub

Sub MunchkinMultiBall
    munchkinModeCnt = 0
    Controller.B2SSetData 240,1
    Display_MX "MUNCHKIN"
    MunchkinModeTimer.enabled = False
    DivRight.RotateToEnd
    SetLightColor l134,"yellow",2
    SetLightColor l161,"pink",2
    SetLightColor l96,"pink",2
    bMunchkinMultiBall=1
    AwardJewel
    bMultiBallMode=1
    If PUPpack=True Then
        DOF 303, DOFPulse
        DOF 80, DOFOn
        vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
    End If
    ChangeGi "red"
    AddMultiball 2
End Sub

Dim T_Ones, T_Tens

Sub MunchkinModeTimer_Timer
    Dim obj, get_color, i
    If munchkinTimeLeft>0 Then munchkinTimeLeft = munchkinTimeLeft - 1
    Controller.B2SSetData T_Ones,0:Controller.B2SSetData T_Tens,0
    If munchkinTimeLeft>9 Then T_Tens = "Md_" & CStr(INT(munchkinTimeLeft/10) MOD 10):Controller.B2SSetData T_Tens,1
    T_Ones = "M1_" & CStr(munchkinTimeLeft MOD 10):Controller.B2SSetData T_Ones,1
    If bMunchkinLollipop=1 AND (munchkinTimeLeft+3) MOD 8 = 0 Then
        If bLollipop = 1 Then
            bLollipop = 0: bLullaby = 1
        Else
            bLollipop = 1: bLullaby = 0
        End If
        LollipopLullabyLights
    End If
    If munchkinTimeLeft=0 Then
        bMunchkinMode=0:Me.enabled = 0
        bMunchkinLand=0:bMunchkinFrenzy=0
        Controller.B2SSetData 237,0
        Controller.B2SSetData 238,0
        Controller.B2SSetData 239,0
        Controller.B2SSetData 240,0
        For each obj in Lights_Munchkin
            get_color = lcolor_dict.Item(obj)
            SetLightColor obj, get_color, 0
        Next
        ClearBG_Twister
        ' Turn off Rainbow Lights that did not get hit during munchkin mode
        For i = 89 To 95
            set obj = light_dict.Item(i)
            get_color = lcolor_dict.Item(obj)
            If SwitchFlags(i)=0 Then
                SetLightColor obj, get_color, 0
            Else
                SetLightColor obj, get_color, 1
            End If
        Next
        If L1.State=1 AND L2.State=1 AND L3.State=1 AND L4.State=1 AND L5.State=1 AND L6.State=1 AND L7.State=1 Then
            bTwisterFlag=1
            If L137.State=0 Then
                DivRight.RotateToEnd: SetLightColor L134,"yellow",2:Controller.B2SSetData 233,1:bTwistAgain=0
            End If
        End If
        If bMunchkinLollipop=1 Then
            For each obj in Lights_Jewels
                get_color = lcolor_dict.Item(obj)
                SetLightColor obj, get_color, 0
            Next
        End If
        bMunchkinLollipop=0
        If bMultiBallMode=0 Then ChangeGI "white"
        Twister_ClearCounts
        Munchkin_ClearDigits
        RestoreBG_Rainbow
    End If
End Sub


' EMERALD CITY mode
Dim cnt_Tinman, cnt_Lion, cnt_Scarecrow
Dim qualLion, qualTinman, qualScarecrow, qualCnt, giftCnt, ECMB_NextLevel

Sub EmeraldCityMB
    Dim i
    MunchkinPost False
    sw6.Timerenabled = 1
    '  Give Credit toward Battle The Witch
    bBTWW_1 = 1
    Controller.B2sSetData "Quadrant_Lt1",1
    BallsInLock=BallsInLock-3
    BallsOnPlayfield=BallsOnPlayfield+2
    bMultiBallMode=1
    If PUPpack=True Then
        DOF 80, DOFOn
        vpmTimer.AddTimer 8000,"DOF 80, DOFOff '"
    End If
    If bCB_LightsOFF=0 AND bCB_LightsON=0 Then ChangeGi "amber"
    bECMultiB = 1
    For i = 106 To 124
        SetLightsOFF i
        SwitchFlags(i) = 0
    Next
    SetLightsBlink(24):SetLightsBlink(85):SetLightsBlink(127)
    If bTwisterFlag=1 Then DivRight.RotateToEnd: SetLightColor L134,"yellow",2:Controller.B2SSetData 233,1
    cnt_Lion=0:cnt_Scarecrow=0:cnt_Tinman=0
    bECLock=0
    bECBallsLocked=0
    DOF 112, DOFPulse	' Strobe for multiball
'	DOF 126, DOFPulse	' Beacon for multiball
    Display_MX "EMERALD"
End Sub


Sub ECMB_Done
    Dim i, obj
    bECMultiB=0
    SetLightsOFF 137
    qualLion=0:qualTinman=0:qualScarecrow=0: qualCnt=0: giftCnt=0: ECMB_NextLevel=0
    For i = 106 To 124
        SetLightsOFF i
        SwitchFlags(i) = 0
    Next
    cnt_Lion=0:cnt_Scarecrow=0:cnt_Tinman=0
    RestoreBG_EMC1
End Sub


Sub ECMultiB_Ramp
    If ECMB_NextLevel>0 Then
        ECMB_NextLevel = ECMB_NextLevel + 1
        If ECMB_NextLevel MOD 2 = 0 Then AddScore_points 12500 + ECMB_NextLevel*2500
        Exit Sub
    End If
    If qualLion=1 Then
        qualLion=2
        Controller.B2sSetData "ECMB_Lion",1
        PlaySound "voc35_hadthenerve"
        If qualCnt<3 Then qualCnt = qualCnt + 1
        AddScore_points 5000 + qualCnt*2500
    End If
    If  qualScarecrow=1 Then
        qualScarecrow=2
        Controller.B2sSetData "ECMB_Scarecrow",1
        PlaySound "voc34_haveabrain"
        If qualCnt<3 Then qualCnt = qualCnt + 1
        AddScore_points 5000 + qualCnt*2500
    End If
    If qualTinman=1 Then
        qualTinman=2
        Controller.B2sSetData "ECMB_Tinman",1
        PlaySound "voc33_haveaheart"
        If qualCnt<3 Then qualCnt = qualCnt + 1
        AddScore_points 5000 + qualCnt*2500
    End If
End Sub

Sub ECMultiB_Throne
    If ECMB_NextLevel>0 Then
        ECMB_NextLevel = ECMB_NextLevel + 1
        If ECMB_NextLevel MOD 2 = 0 Then AddScore_points 40000 + ECMB_NextLevel*10000
        Exit Sub
    End If
    If qualLion=2 AND qualScarecrow=2 AND qualTinman=2 Then
        qualLion=3: qualScarecrow=3: qualTinman=3
        AddScore_points 15000
    Else
       If qualLion=4 Then
           qualLion=5
           Controller.B2sSetData "ECMB_Medal",1
           PlaySound "voc46_ShucksImSpeechless"
           If giftCnt<3 Then giftCnt = giftCnt+1
           AddScore_points 10000 + giftCnt*10000
       End If
       If qualScarecrow=4 Then
           qualScarecrow=5
           Controller.B2sSetData "ECMB_Diploma",1
           PlaySound "voc45_Igotabrain"
           If giftCnt<3 Then giftCnt = giftCnt+1
           AddScore_points 10000 + giftCnt*10000
       End If
       If qualTinman=4 Then
           qualTinman=5
           Controller.B2sSetData "ECMB_Heart",1
           PlaySound "voc47_ItTicks"
           If giftCnt<3 Then giftCnt = giftCnt+1
           AddScore_points 10000 + giftCnt*10000
       End If
       If qualLion=5 AND qualScarecrow=5 AND qualTinman=5 Then
           giftCnt=0
           AwardJewel
           qualLion=6: qualScarecrow=6: qualTinman=6
           ECMB_NextLevel=1
       End If
    End If
End Sub

Sub sw6_Timer
	Me.Timerenabled = 0
    PlaySound "fx_kicker"
	sw6.Kick 180, 5
    LightState(150) = 0:Flashers(flash_dict.Item(l150))=0
    sw7.Timerenabled = 1
End Sub

Sub sw7_Timer
	Me.Timerenabled = 0
    PlaySound "fx_kicker"
	sw7.Kick 180, 5
    LightState(151) = 0:Flashers(flash_dict.Item(l151))=0
    sw8.Timerenabled = 1
End Sub

Sub sw8_Timer
	Me.Timerenabled = 0
    PlaySound "fx_kicker"
	sw8.Kick 180, 5
    LightState(152) = 0:Flashers(flash_dict.Item(l152))=0
End Sub


' RESCUE Mode
Dim bCastlePFActive, bCastleDoorOpen, bRescueMB, bMegaPot
Dim castleDoorBashCnt, ballRescueLock
Dim ballsOnCastlePF
Dim RescueMB_X(6)

Sub RescueMBClock_Timer
    Dim i, obj
    Me.enabled = 0
    Controller.B2SSetData 243,0
    Controller.B2SSetdata "Rescue_MB",0
    Controller.B2SSetData "Rescue_1",1
    For i = 1 To 6
        Controller.B2SSetData "1X_" & CStr(i),1
        RescueMB_X(i) = 1
        SwitchFlags(64+i)=0
    Next
    bMegaPot = 0
    For each obj in Lights_Castle
        SetLightColor obj,"orange",0
    Next
    For each obj in Lights_Jewels
        SetLightColor obj,"blue",2
    Next
End Sub

Sub RescueMB_Update(switch)
    Dim i, obj, pts
    Dim tmp0, tmp1, tmp2, tmp3
    Select Case switch
      case 9
        i = 1
        Set obj = l106
      case 24
        i = 2
        Set obj = l107
      case 17,73
        i = 3
        Set obj = l108
      case 76
        i = 4
        Set obj = l113
      case 16
        i = 5
        Set obj = l112
      case 15
        i = 6
        Set obj = l111
    End Select
    If RescueMB_X(i)=3 Then SetLightColor obj,"orange",2
    If PUPpack=True Then
        DOF 64, DOFOn
        vpmTimer.AddTimer 3000,"DOF 64, DOFOff '"
    End If
    PlaySound "voc63_jackpot"
    pts = RescueMB_X(i)*200
    AddScore_Points pts
    If DMD_Option=True Then
        DMD2Flush
        If(dqHead2 = dqTail2)Then
            tmp0 = ""
            tmp1 = "RESCUE JACKPOT"
            tmp2 = FillLine(2, "POINTS", FormatScore(pts))
            tmp3 = ""
            DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2500, True, ""
        End If
    End If
    RescueMB_X(i) = RescueMB_X(i) + 1
    If RescueMB_X(i)>3 Then RescueMB_X(i)=0.5
    If RescueMB_X(i)<>0.5 Then
        Controller.B2SSetData "1X_" & CStr(i),0
        Controller.B2SSetData "2X_" & CStr(i),0
        Controller.B2SSetData Cstr(RescueMB_X(i)) & "X_" & CStr(i),1
    Else
        Controller.B2SSetData "3X_" & CStr(i),0
    End If
    If l106.State=2 AND l107.State=2 AND l108.State=2 AND l111.State=2 AND _
       l112.State=2 AND l113.State=2 Then
        SwitchFlags(71) = 0
        bMegaPot=1
        SetLightColor l110,"white",2:SetLightColor l109,"white",1
    End If
End Sub

Sub RescueMB_Jackpot
    Dim cnt, i, obj
    Dim tmp0, tmp1, tmp2, tmp3
    If bMegaPot=1 AND SwitchFlags(71)>0 Then
        AwardJewel
        AddScore_Points MegaJackpot
        PlaySound "voc65_megajackpot"
        If DMD_Option=True Then
          DMD2Flush
            If(dqHead2 = dqTail2)Then
                tmp0 = ""
                tmp1 = "MEGA  JACKPOT!!!"
                tmp2 = FillLine(2, "POINTS", FormatScore(MegaJackpot))
                tmp3 = ""
                DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2500, True, ""
            End If
        End If
        PlaySound "fx_fanfare5"
    Else
        cnt = 0
        For i = 1 to 6
            If RescueMB_X(i)=3 OR RescueMB_X(i)=0.5 Then cnt = cnt + 1
        Next
        If cnt=0 Then Exit Sub
        PlaySound "voc64_superjackpot"
        AddScore_Points SuperJackpot*cnt
        If DMD_Option=True Then
          DMD2Flush
            If(dqHead2 = dqTail2)Then
                tmp0 = ""
                tmp1 = "  SUPER JACKPOT!"
                tmp2 = FillLine(2, "POINTS", FormatScore(SuperJackpot*cnt))
                tmp3 = ""
                DMD2 tmp0, eNone, tmp1, eNone, tmp2, eNone, tmp3, eNone, 2500, True, ""
            End If
        End If
        PlaySound "fx_fanfare1"
    End If
    RescueMB_End
    For i = 1 To 6
        Controller.B2SSetData "1X_" & CStr(i),1
    Next
    For each obj in Lights_Jewels
        SetLightColor obj,"blue",2
    Next
End Sub

Sub RescueMB_End
    Dim i, obj
    For i = 1 To 6
        SwitchFlags(64+i)=0
        Controller.B2SSetData "3X_" & CStr(i),0
        Controller.B2SSetData "2X_" & CStr(i),0
        Controller.B2SSetData "1X_" & CStr(i),0
        RescueMB_X(i) = 1
    Next
    For each obj in Lights_Castle
         SetLightColor obj,"orange",0
    Next
    For each obj in Lights_Jewels
        SetLightColor obj,"blue",0
    Next
    SetLightColor l110,"orange",0:SetLightColor l109,"orange",0
    SwitchFlags(71)=0
    bMegaPot = 0
End Sub

' WIZARD Mode
Dim wiz_points
Sub Wizard_reset
   w_sign.visible=0: i_sign.visible=0: z_sign.visible=0
   a_sign.visible=0: r_sign.visible=0: d_sign.visible=0
   SwitchFlags(15)=0
End Sub

'WIZARD Restore
Sub Wizard_restore
    Select Case throneHits
      case 1
        w_sign.visible=1
      case 2
        w_sign.visible=1: i_sign.visible=1
      case 3
        w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
      case 4
        w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
        a_sign.visible=1
     case 5
       w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
       a_sign.visible=1:r_sign.visible=1
     case 6
       w_sign.visible=1: i_sign.visible=1: z_sign.visible=1
       a_sign.visible=1: r_sign.visible=1: d_sign.visible=1
    End Select
End Sub


' SUPER X
Sub SuperXcolor(superX)
    Select Case superX
      case 2
        SetLightColor l48,"cyan",2
      case 3
        SetLightColor l48,"green",2
      case 4
        SetLightColor l48,"yellow",2
      case 6
        SetLightColor l48,"orange",2
      case 9
        SetLightColor l48,"red",2
      case Else
        SetLightColor l48,"white",0
     End Select
End Sub


' SOUND Support
Sub PlayStartupSound
    Select Case INT(RND * 5 + 1)
      case 1
        PlaySound "voc0_WitchCackle"
      case 2
        PlaySound "voc4_notinkansas"
      case 3
        PlaySound "voc3_trytostay"
      case 4
        PlaySound "voc1_illgetyou"
      case 5
        PlaySound "voc27_GoodWitchBad"
    End Select
End Sub


' END OF GAME Support
Dim bEndOfGame

Sub ModeTerminate_EndOfGame
    bEndOfGame=1
    BallLossCheck.Enabled=False
    'Cleanup
    Wizard_reset
    'Flashers
    LightState(104)=0:Flashers(flash_dict.Item(l104))=0
    LightState(105)=0:Flashers(flash_dict.Item(l105))=0
    LightState(115)=0:Flashers(flash_dict.Item(l115))=0
    LightState(150)=0:Flashers(flash_dict.Item(l150))=0
    LightState(151)=0:Flashers(flash_dict.Item(l151))=0
    LightState(152)=0:Flashers(flash_dict.Item(l152))=0
    LightState(153)=0:Flashers(flash_dict.Item(l153))=0
    ' Locked Ball Release or Destroy
    BallsInLock = 0
    sw72_out.DestroyBall
    EndOfGameCleanUp.Interval = 1000
    EndOfGameCleanUp.Enabled = True
End Sub

Sub EndOfGameCleanUp_Timer
    If bAutoFlip=1 Then Auto_Regression_Flip
    If bECBallsLocked=0 AND bBallInPlungerLane=False Then
       If Me.interval=1000 Then
           Me.interval = 10000
       Else
           Me.enabled = 0
           ClearBackGlass
       End If
    End If
    If bECBallsLocked > 0 Then
        Select Case bECBallsLocked
          case 1
            sw6.Kick 180,5
          case 2
            sw6.Kick 180,5
            sw7.Kick 180,5
            sw6.Kick 180,5
        End Select
        LightState(150) = 0:LightState(151) = 0:LightState(152) = 0
        Flashers(flash_dict.Item(l150))=0:Flashers(flash_dict.Item(l151))=0:Flashers(flash_dict.Item(l152))=0
    End If
    If bBallInPlungerLane=True Then PlungerIM.AutoFire:AutoFireCheck.Enabled = True
End Sub


'******************************************************************************
'       TABLE SPECIFIC MAPPINGS - THINGS THAT DON"T CHANGE
'******************************************************************************
Dim SwitchFlags(200)
Dim BasePoints(200)
Dim BaseColor(200)
Dim Flashers(16)
Dim SkillshotValue(4)
Dim SkillshotIncrease(4)

Dim z
' Clear Arrays and set to default
For z = 0 to UBound(SwitchFlags):SwitchFlags(z) = 0:Next
For z = 0 to UBound(BasePoints):BasePoints(z) = 0:Next
For z = 0 to UBound(BaseColor):BaseColor(z) = "white":Next
For z = 0 to UBound(Flashers):Flashers(z) = 0:Next

SkillshotValue(1) = 5000 ' OZ lane light
SkillshotValue(2) = 5000 ' Witch
SkillshotValue(3) = 7500 ' Haunted Forest Target
SkillshotValue(4) = 10000 ' Crystal BALL light
SkillshotIncrease(1) = 1000 ' OZ lane light
SkillshotIncrease(2) = 1000 ' Witch
SkillshotIncrease(3) = 2500 ' Haunted Forest Target
SkillshotIncrease(4) = 5000 ' Crystal BALL light

' Base points awarded when switch hit
' Array index is switch number
BasePoints(0)  = 0		'empty
BasePoints(1)  = 1  	'Handles 1 point awards
BasePoints(2)  = 2  	'Handles 2 points for spinner
BasePoints(3)  = 3  	'Handles 3 point awards - e.g. Slings
BasePoints(15) = 165	'Throne Kicker
BasePoints(16) = 5		'Right Orbit Enter
BasePoints(17) = 50		'VUK to Castle Kicker
BasePoints(18) = 150	'Castle Kicker
BasePoints(21) = 250	'VUK - Crystal Ball
BasePoints(24) = 5  	'Left Orbit Enter
BasePoints(27) = 3		'Left Return Lane
BasePoints(28) = 30		'Left Outlane
BasePoints(29) = 115	'Crystal Ball Target - B
BasePoints(30) = 115	'Crystal Ball Target - A
BasePoints(31) = 115	'Crystal Ball Target - L
BasePoints(32) = 115	'Crystal Ball Target - L
BasePoints(37) = 7		'Toto Rollover - T
BasePoints(38) = 7		'Toto Rollover - O
BasePoints(39) = 7		'Toto Rollover - T
BasePoints(40) = 7		'Toto Rollover - O
BasePoints(41) = 60		'Bumper - State Fair
BasePoints(42) = 30		'TNPLH - Rubber
BasePoints(43) = 25		'TNPLH - THERE'S
BasePoints(44) = 25		'TNPLH - NO
BasePoints(45) = 25		'TNPLH - PLACE
BasePoints(46) = 25		'TNPLH - LIKE
BasePoints(47) = 25		'TNPLH - HOME
BasePoints(49) = 104	'Bumper - Left- MidTop
BasePoints(50) = 104	'Bumper - Left- MidCtr
BasePoints(51) = 104	'Bumper - Left- MidBot
BasePoints(52) = 50		'Bumper Entry Rubber - Lights Glinda
BasePoints(53) = 15		'Bumper Exit Lane
BasePoints(56) = 100	'Tin Man Rollover
BasePoints(57) = 50		'Witch - Top
BasePoints(58) = 75		'Witch - Bottom
BasePoints(62) = 20		'Castle PF Exit
BasePoints(64) = 7		'Left Orbit Made
BasePoints(65) = 115	'Castle RESCUE - S
BasePoints(66) = 115	'Castle RESCUE - E
BasePoints(67) = 115	'Castle RESCUE - R
BasePoints(68) = 115	'Castle RESCUE - E
BasePoints(69) = 115	'Castle RESCUE - U
BasePoints(70) = 115	'Castle RESCUE - C
BasePoints(71) = 50		'Castle - Loop Trigger
BasePoints(72) = 0		'Castle - Dorothy Trigger
BasePoints(73) = 250	'Winkie Guard DT
BasePoints(74) = 20		'Glinda Target (Unlit)
BasePoints(75) = 6		'Emerald Ramp Enter
BasePoints(76) = 25		'Emerald Ramp Made
BasePoints(77) = 15		'Castle Door - Left
BasePoints(78) = 15		'Castle Door - Right
BasePoints(81) = 20		'Left OZ   (Each Way)
BasePoints(82) = 20		'Right OZ (Each Way)
BasePoints(87) = 7		'Right Orbit Made
BasePoints(89) = 25 	'Rainbow - R
BasePoints(90) = 25 	'Rainbow - A
BasePoints(91) = 25 	'Rainbow - I
BasePoints(92) = 25 	'Rainbow - N
BasePoints(93) = 25 	'Rainbow - B
BasePoints(94) = 25 	'Rainbow - O
BasePoints(95) = 25 	'Rainbow - W
BasePoints(96) = 50 	'Scarecrow Rollover
BasePoints(101) = 50 	'Munchkin - Loop Upper
BasePoints(102) = 0 	'Munchkin - Loop Lower
BasePoints(103) = 300 	'Throne Rubber
BasePoints(104) = 150 	'Lion Rollover

' Color Map array for switch associated lights
' Array index is switch number
BaseColor(37) = "white"
BaseColor(38) = "white"
BaseColor(39) = "white"
BaseColor(40) = "white"
BaseColor(43) = "white"
BaseColor(44) = "white"
BaseColor(45) = "white"
BaseColor(46) = "white"
BaseColor(47) = "white"
BaseColor(52) = "white"
BaseColor(56) = "red"
BaseColor(61) = "blue"
BaseColor(62) = "blue"
BaseColor(65) = "orange"
BaseColor(66) = "orange"
BaseColor(67) = "orange"
BaseColor(68) = "orange"
BaseColor(69) = "orange"
BaseColor(70) = "orange"
BaseColor(73) = "blue"
BaseColor(89) = "red"
BaseColor(90) = "orange"
BaseColor(91) = "yellow"
BaseColor(92) = "ylwgrn"
BaseColor(93) = "green"
BaseColor(94) = "blue"
BaseColor(95) = "purple"
BaseColor(96) = "yellow"
BaseColor(104) = "green"
BaseColor(106) = "green"
BaseColor(107) = "green"
BaseColor(108) = "green"
BaseColor(109) = "green"
BaseColor(110) = "yellow"
BaseColor(111) = "yellow"
BaseColor(112) = "yellow"
BaseColor(113) = "yellow"
BaseColor(114) = "yellow"
BaseColor(115) = "yellow"
BaseColor(116) = "yellow"
BaseColor(117) = "yellow"
BaseColor(118) = "yellow"
BaseColor(119) = "red"
BaseColor(120) = "red"
BaseColor(121) = "red"
BaseColor(122) = "red"
BaseColor(123) = "red"
BaseColor(124) = "red"
BaseColor(192) = "green"
BaseColor(193) = "green"
BaseColor(194) = "green"
BaseColor(195) = "green"
BaseColor(196) = "green"
BaseColor(197) = "green"
BaseColor(198) = "green"
BaseColor(199) = "green"


' Set up Switch to Light Map Dictionary
' The switch number will be used as the key with the associated light object as the item
' Dictionary used because lights are objects, switch numbers are not strings & simple
Dim light_dict
Set light_dict = CreateObject("Scripting.Dictionary")

light_dict.Add 29 ,L72
light_dict.Add 30 ,L73
light_dict.Add 31 ,L74
light_dict.Add 32 ,L75
light_dict.Add 37 ,L37
light_dict.Add 38 ,L38
light_dict.Add 39 ,L39
light_dict.Add 40 ,L40
light_dict.Add 43 ,l59
light_dict.Add 44 ,l60
light_dict.Add 45 ,l61
light_dict.Add 46 ,l62
light_dict.Add 47 ,l63
light_dict.Add 52 ,L135
light_dict.Add 61 ,L101
light_dict.Add 62 ,L102
light_dict.Add 65 ,L108
light_dict.Add 66 ,L107
light_dict.Add 67 ,l106
light_dict.Add 68 ,l113
light_dict.Add 69 ,l112
light_dict.Add 70 ,l111
light_dict.Add 71 ,l109
light_dict.Add 56 ,L85
light_dict.Add 73 ,L141
light_dict.Add 89 ,L7
light_dict.Add 90 ,L6
light_dict.Add 91 ,L5
light_dict.Add 92 ,L4
light_dict.Add 93 ,L3
light_dict.Add 94 ,L2
light_dict.Add 95 ,L1
light_dict.Add 96 ,L127
light_dict.Add 104 ,L24
light_dict.Add 106 ,L20
light_dict.Add 107 ,L21
light_dict.Add 108 ,L22
light_dict.Add 109 ,L23
light_dict.Add 110 ,L10
light_dict.Add 111 ,L11
light_dict.Add 112 ,L12
light_dict.Add 113 ,L13
light_dict.Add 114 ,L14
light_dict.Add 115 ,L15
light_dict.Add 116 ,L16
light_dict.Add 117 ,L17
light_dict.Add 118 ,L18
light_dict.Add 119 ,L79
light_dict.Add 120 ,L78
light_dict.Add 121 ,L77
light_dict.Add 122 ,L81
light_dict.Add 123 ,L82
light_dict.Add 124 ,L83
light_dict.Add 134 ,l134
light_dict.Add 137 ,l137


' Set up Light to Color Map Dictionary
' The light object will be used as the key with the associated light color as the item
' Dictionary used because lights are objects and random access common
Dim lcolor_dict
Set lcolor_dict = CreateObject("Scripting.Dictionary")

lcolor_dict.Add L1, "purple"
lcolor_dict.Add L2, "blue"
lcolor_dict.Add L3, "green"
lcolor_dict.Add L4, "ylwgrn"
lcolor_dict.Add L5, "yellow"
lcolor_dict.Add L6, "orange"
lcolor_dict.Add L7, "red"
lcolor_dict.Add L9, "yellow"
lcolor_dict.Add L10, "yellow"
lcolor_dict.Add L11, "yellow"
lcolor_dict.Add L12, "yellow"
lcolor_dict.Add L13, "yellow"
lcolor_dict.Add L14, "yellow"
lcolor_dict.Add L15, "yellow"
lcolor_dict.Add L16, "yellow"
lcolor_dict.Add L17, "yellow"
lcolor_dict.Add L18, "yellow"
lcolor_dict.Add L20, "green"
lcolor_dict.Add L21, "green"
lcolor_dict.Add L22, "green"
lcolor_dict.Add L23, "green"
lcolor_dict.Add L24, "green"
lcolor_dict.Add L25, "white"
lcolor_dict.Add L27, "white"
lcolor_dict.Add L28, "yellow"
lcolor_dict.Add L29, "yellow"
lcolor_dict.Add L30, "white"
lcolor_dict.Add L31, "orange"
lcolor_dict.Add L32, "orange"
lcolor_dict.Add L33, "blue"
lcolor_dict.Add L34, "green"
lcolor_dict.Add L36, "white"
lcolor_dict.Add L37, "white"
lcolor_dict.Add L38, "white"
lcolor_dict.Add L39, "white"
lcolor_dict.Add L40, "white"
lcolor_dict.Add L41, "white"
lcolor_dict.Add L42, "white"
lcolor_dict.Add L43, "yellow"
lcolor_dict.Add L44, "yellow"
lcolor_dict.Add L45, "yellow"
lcolor_dict.Add L46, "yellow"
lcolor_dict.Add L47, "yellow"
lcolor_dict.Add L49, "orange"
lcolor_dict.Add L50, "yellow"
lcolor_dict.Add L51, "yellow"
lcolor_dict.Add L52, "yellow"
lcolor_dict.Add L53, "yellow"
lcolor_dict.Add L54, "yellow"
lcolor_dict.Add L55, "yellow"
lcolor_dict.Add L56, "yellow"
lcolor_dict.Add L57, "yellow"
lcolor_dict.Add l58, "orange"
lcolor_dict.Add l59, "white"
lcolor_dict.Add l60, "white"
lcolor_dict.Add l61, "white"
lcolor_dict.Add l62, "white"
lcolor_dict.Add l63, "white"
lcolor_dict.Add l64, "white"
lcolor_dict.Add l65, "white"
lcolor_dict.Add l66, "white"
lcolor_dict.Add l68, "blue"
lcolor_dict.Add l69, "yellow"
lcolor_dict.Add l70, "green"
lcolor_dict.Add L76, "white"
lcolor_dict.Add L79, "red"
lcolor_dict.Add L78, "red"
lcolor_dict.Add L77, "red"
lcolor_dict.Add L81, "red"
lcolor_dict.Add L82, "red"
lcolor_dict.Add L83, "red"
lcolor_dict.Add L85, "red"
lcolor_dict.Add L86, "purple"
lcolor_dict.Add L87, "purple"
lcolor_dict.Add L88, "purple"
lcolor_dict.Add L89, "purple"
lcolor_dict.Add L90, "purple"
lcolor_dict.Add L91, "purple"
lcolor_dict.Add L92, "purple"
lcolor_dict.Add L93, "purple"
lcolor_dict.Add L95, "red"
lcolor_dict.Add L97, "purple"
lcolor_dict.Add l98, "white"
lcolor_dict.Add l100, "white"
lcolor_dict.Add l101, "blue"
lcolor_dict.Add l102, "blue"
lcolor_dict.Add l106, "orange"
lcolor_dict.Add L107, "orange"
lcolor_dict.Add L108, "orange"
lcolor_dict.Add L109, "orange"
lcolor_dict.Add L110, "orange"
lcolor_dict.Add l111, "orange"
lcolor_dict.Add l112, "orange"
lcolor_dict.Add l113, "orange"
lcolor_dict.Add L127, "yellow"
lcolor_dict.Add L128, "yellow"
lcolor_dict.Add L129, "yellow"
lcolor_dict.Add l134, "yellow"
lcolor_dict.Add L135, "pink"
lcolor_dict.Add l137, "green"
lcolor_dict.Add L139, "white"
lcolor_dict.Add L141, "white"
lcolor_dict.Add L142, "red"
lcolor_dict.Add l143, "white"
lcolor_dict.Add l161, "pink"
lcolor_dict.Add l162, "green"
lcolor_dict.Add l163, "blue"
lcolor_dict.Add l164, "red"
lcolor_dict.Add L192, "green"
lcolor_dict.Add L193, "green"
lcolor_dict.Add L194, "green"
lcolor_dict.Add L195, "green"
lcolor_dict.Add L196, "green"
lcolor_dict.Add L197, "green"
lcolor_dict.Add L198, "green"
lcolor_dict.Add L199, "green"
lcolor_dict.Add ScoopLight1, "darkblue"
lcolor_dict.Add ScoopLight2, "green"
lcolor_dict.Add ScoopLight3, "purple"

' Set up Flasher Object to  Flasher State Map Dictionary
Dim flash_dict
Set flash_dict = CreateObject("Scripting.Dictionary")

' OZ Flashers
flash_dict.Add l104, 1
flash_dict.Add l105, 2
' Dorothy Flashers
flash_dict.Add l115, 3
flash_dict.Add l153, 4
' Munchkin House Flashers
flash_dict.Add l150, 5
flash_dict.Add l151, 6
flash_dict.Add l152, 7

' Set up Switch to Sound Map Dictionary
' The switch number will be used as the key with the associated sound as the item
' Dictionary used because it allows a check for presence. If not present, do local processing.
Dim sound_dict
Set sound_dict = CreateObject("Scripting.Dictionary")

sound_dict.Add 15, "fx_kicker_enter"
sound_dict.Add 17, "fx_kicker_enter"
sound_dict.Add 18, "fx_kicker_enter"
sound_dict.Add 21, "fx_kicker_enter"
sound_dict.Add 29, "fx_target"
sound_dict.Add 30, "fx_target"
sound_dict.Add 31, "fx_target"
sound_dict.Add 32, "fx_target"
sound_dict.Add 43, "fx_target"
sound_dict.Add 44, "fx_target"
sound_dict.Add 45, "fx_target"
sound_dict.Add 46, "fx_target"
sound_dict.Add 47, "fx_target"
sound_dict.Add 55, "fx_target"
sound_dict.Add 61, "fx_target"
sound_dict.Add 62, "fx_target"
sound_dict.Add 65, "fx_target"
sound_dict.Add 66, "fx_target"
sound_dict.Add 67, "fx_target"
sound_dict.Add 68, "fx_target"
sound_dict.Add 69, "fx_target"
sound_dict.Add 70, "fx_target"
sound_dict.Add 73, "fx_target"
sound_dict.Add 74, "fx_target"
sound_dict.Add 89, "fx_target"
sound_dict.Add 90, "fx_target"
sound_dict.Add 91, "fx_target"
sound_dict.Add 92, "fx_target"
sound_dict.Add 93, "fx_target"
sound_dict.Add 94, "fx_target"
sound_dict.Add 95, "fx_target"


'******************************************************************************
'       BACK GLASS PROCESSING
'******************************************************************************
' Set up Light to BackGlass Map Dictionary
' The Light label will be used as the key with the associated Backglass ID as the item
' Dictionary used because light label is an object
Dim bg_dict
Set bg_dict = CreateObject("Scripting.Dictionary")

' Rainbow Lights
bg_dict.Add l7, 181
bg_dict.Add l6, 182
bg_dict.Add l5, 183
bg_dict.Add l4, 184
bg_dict.Add l3, 185
bg_dict.Add l2, 186
bg_dict.Add l1, 187
' Scarecrow Lights
bg_dict.Add l10, 221
bg_dict.Add l11, 220
bg_dict.Add l12, 219
bg_dict.Add l13, 218
bg_dict.Add l14, 217
bg_dict.Add l15, 216
bg_dict.Add l16, 215
bg_dict.Add l17, 214
bg_dict.Add l18, 213
' Lion Lights
bg_dict.Add l20, 228
bg_dict.Add l21, 229
bg_dict.Add l22, 230
bg_dict.Add l23, 231
' Tin Man Lights
bg_dict.Add l77, 224
bg_dict.Add l78, 223
bg_dict.Add l79, 222
bg_dict.Add l81, 225
bg_dict.Add l82, 226
bg_dict.Add l83, 227
'  HAUNTED Lights
bg_dict.Add l86, 190
bg_dict.Add l87, 191
bg_dict.Add l88, 192
bg_dict.Add l89, 193
bg_dict.Add l91, 194
bg_dict.Add l92, 195
bg_dict.Add l93, 196
' RESCUE Lights
bg_dict.Add l106, 207
bg_dict.Add l107, 208
bg_dict.Add l108, 209
bg_dict.Add l113, 210
bg_dict.Add l112, 211
bg_dict.Add l111, 212


Sub SetBackGlass(obj)
    Dim get_id
    If bg_dict.Exists(obj) Then
        get_id = bg_dict.Item(obj)
        Controller.B2SSetData get_id, 1
    End If
End Sub

Sub BackGlass_AttractMode_OFF
	Dim i
    ' Display Screen Panels
    For i = 202 to 205
        Controller.B2SSetData i,1
    Next
    Controller.B2SSetData "Quadrant_Small",1
End Sub

Sub BackGlass_AttractMode_ON
    ClearBackGlass
End Sub

Sub ClearBackGlass
    Dim i, tmpStr
    ' Dorothy Captured, Storm Castle Off & Rainbow Lights
    For i = 179 to 187
        Controller.B2SSetData i,0
    Next
    ' Explosion Animation
    Controller.B2SSetData 189,0
    ' Haunted Lights Off
    For i = 190 to 196
        Controller.B2SSetData i,0
    Next
    ' Hurry Up Witch
    Controller.B2SSetData "HU_Witch_1",0
    Controller.B2SSetData "HU_Witch_2",0
    Controller.B2SSetData "HU_Witch_3",0
    Controller.B2SSetData "HU_Witch_0",0
    For i = 197 to 200
    Controller.B2SSetData i,0
    Next
    Controller.B2SSetData 188,0
    ' Screen Panels
    For i = 202 to 206
        Controller.B2SSetData i,0
    Next
     Controller.B2SSetData "Quadrant_Small",0
     Controller.B2sSetData "Quadrant_Lt1",0
     Controller.B2sSetData "Quadrant_Lt2",0
     Controller.B2sSetData "Quadrant_Lt3",0
     Controller.B2sSetData "Quadrant_Lt4",0
    ' Rescue, Scarecrow, TinMan & Lion Lights
    For i = 207 to 231
        Controller.B2SSetData i,0
    Next
    ' EMC Lock & Multiball screens
    Controller.B2SSetData "EMC_Ball1",0:Controller.B2SSetData "EMC_Ball2",0:Controller.B2SSetData "EMC_MB1",0
    Controller.B2sSetData "ECMB_Tinman",0
    Controller.B2sSetData "ECMB_Lion",0
    Controller.B2sSetData "ECMB_Scarecrow",0
    Controller.B2sSetData "ECMB_Heart",0
    Controller.B2sSetData "ECMB_Medal",0
    Controller.B2sSetData "ECMB_Diploma",0
    ' Twister Mode Light Off
    Controller.B2SSetData 233,0
    ' Twister Screens
    ClearBG_Twister
    ' Munchkin Screens
    Controller.B2SSetData T_Ones,0:Controller.B2SSetData T_Tens,0
    Controller.B2SSetData 237,0
    Controller.B2SSetData 238,0
    Controller.B2SSetData 239,0
    Controller.B2SSetData 240,0
    Munchkin_ClearDigits
    ' Rescue MB
    ClearBG_RescueMB
    ' HOADC
    HOADC_BG_clr
    ' Fireball Frenzy
    ClearBG_FF
    ' BTWW
    ClearBG_BTWW
    ' TOTO
    ClearBG_TOTO
    ' TNPLH
    ClearBG_TNPLH
    ' SOTR
    ClearBG_SOTR
End Sub

Sub ClearBG_NewBall
    Dim i
    ' Hurry Up Witch
    For i = 197 to 200
        Controller.B2SSetData i,0
    Next
    ' HU Witch Explosion Animation
    Controller.B2SSetData 189,0
    ' Haunted
    ClearBG_HauntedModes
    ' EMC Lock & Multiball screens
    Controller.B2SSetData "EMC_Ball1",0:Controller.B2SSetData "EMC_Ball2",0:Controller.B2SSetData "EMC_MB1",0
    Controller.B2sSetData "ECMB_Tinman",0
    Controller.B2sSetData "ECMB_Lion",0
    Controller.B2sSetData "ECMB_Scarecrow",0
    Controller.B2sSetData "ECMB_Heart",0
    Controller.B2sSetData "ECMB_Medal",0
    Controller.B2sSetData "ECMB_Diploma",0
    ' Crystal Ball Modes
    ClearBG_CB
    ' Twister & Munchkin Screens including digits
    ClearBG_Twister
    ' Rescue MB
    ClearBG_RescueMB
    ' HOADC
    HOADC_BG_clr
    ' Fireball Frenzy
    ClearBG_FF
    ' BTWW
    ClearBG_BTWW
    ' DingDong
    WitchHU_ClearScore
    Controller.B2sSetData "DingDong1",0
    ' TOTO
    ClearBG_TOTO
    ' TNPLH
    ClearBG_TNPLH
End Sub

Sub ClearBG_TOTO
    Dim i
'    Controller.B2sSetData "Gulch_1",0
    Controller.B2sSetData "Gulch_2",0
    Controller.B2sSetData "Gulch_3",0
    Controller.B2sSetData "Gulch_4",0
    For i = 0 To 9
        T_Ones = "T1_" & Cstr(i)
        Controller.B2SSetData T_Ones,0
    Next
    For i = 1 To 9
        T_Tens = "Td_" & Cstr(i)
        Controller.B2SSetData T_Tens,0
    Next
End Sub

Sub ClearBG_TNPLH
    Controller.B2SSetData "TNPLHS0_0",0
    Controller.B2SSetData "TNPLHS0_1",0
    Controller.B2SSetData "TNPLHS1_1",0
    Controller.B2SSetData "TNPLHS2_1",0
    Controller.B2SSetData "TNPLHS3_1",0
    Controller.B2SSetData "TNPLHS4_1",0
End Sub

Sub ClearBG_EMC1
    Dim i
    ' Rescue, Scarecrow, TinMan & Lion Lights Off
    For i = 207 to 231
        Controller.B2SSetData i,0
    Next
    Controller.B2SSetData 204,0
End Sub

Sub RestoreBG_EMC1
    Dim obj
    Dim bg_id
    Controller.B2SSetData "EMC_Ball1",0:Controller.B2SSetData "EMC_Ball2",0:Controller.B2SSetData "EMC_MB1",0
    Controller.B2sSetData "ECMB_Tinman",0
    Controller.B2sSetData "ECMB_Lion",0
    Controller.B2sSetData "ECMB_Scarecrow",0
    Controller.B2sSetData "ECMB_Heart",0
    Controller.B2sSetData "ECMB_Medal",0
    Controller.B2sSetData "ECMB_Diploma",0
    Controller.B2SSetData 204,1
    ' Rescue, Scarecrow, TinMan & Lion Lights Restored
    For each obj in Lights_Lion
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
    For each obj in Lights_Scarecrow
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
    For each obj in Lights_TinMan
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
End Sub

Sub ClearBG_TRainbow
    Dim i
    ' Rainbow Lights Off
    For i = 181 to 187
        Controller.B2SSetData i,0
    Next
    Controller.B2SSetData 233,0
    Controller.B2SSetData 203,0
End Sub

Sub RestoreBG_Rainbow
    Dim obj
    Dim bg_id
    Controller.B2SSetData 237,0
    Controller.B2SSetData 238,0
    Controller.B2SSetData 239,0
    Controller.B2SSetData 240,0
    Controller.B2SSetData 203,1
    ' Rainbow Lights Restored
    For each obj in Lights_Rainbow
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
End Sub

Sub ClearBG_Twister
    Twister_ClearCounts
    Controller.B2SSetData 235,0
    Controller.B2SSetData 236,0
    Controller.B2SSetData 237,0
    Controller.B2SSetData 238,0
    Controller.B2SSetData 239,0
    Controller.B2SSetData 240,0
End Sub

Sub Twister_ClearCounts
    Dim i
    For i = 0 To 9
        L_Ones = "L1_" & Cstr(i)
        Controller.B2SSetData L_Ones,0
        Y_Ones = "Y1_" & Cstr(i)
        Controller.B2SSetData Y_Ones,0
        T_Ones = "T1_" & Cstr(i)
        Controller.B2SSetData T_Ones,0
        If i<>0 Then
            L_Tens = "Ld_" & Cstr(i)
            Controller.B2SSetData L_Tens,0
            Y_Tens = "Yd_" & Cstr(i)
            Controller.B2SSetData Y_Tens,0
            T_Tens = "Td_" & Cstr(i)
            Controller.B2SSetData T_Tens,0
        End If
    Next
End Sub

Sub Munchkin_ClearDigits
    Dim i
        For i = 0 To 9
            T_Ones = "M1_" & Cstr(i)
            Controller.B2SSetData T_Ones,0
        Next
        For i = 1 To 9
            T_Tens = "Md_" & Cstr(i)
            Controller.B2SSetData T_Tens,0
        Next
End Sub

Sub ClearBG_Rescue
    Dim i
    For i = 206 to 212
        Controller.B2SSetData i,0
    Next
    Controller.B2SSetData 241,0
    Controller.B2SSetData 242,0
    Controller.B2SSetData 243,0
    Controller.B2SSetData "Rescue_MB",0

End Sub

Sub ClearBG_RescueMB
    Dim i
    Controller.B2SSetData "Rescue_1",0
    For i = 1 To 6
        Controller.B2SSetData "3X_" & CStr(i),0
        Controller.B2SSetData "2X_" & CStr(i),0
        Controller.B2SSetData "1X_" & CStr(i),0
    Next
End Sub

Sub ClearBG_FF
    Dim i
    Controller.B2SSetData "FireballFrenzy1",0
    For i = 0 To 9
        J_Ones = "J1_" & Cstr(i)
        Controller.B2SSetData J_Ones,0
    Next
    For i = 1 To 9
        J_Tens = "Ld_" & Cstr(i)
        Controller.B2SSetData J_Tens,0
    Next
End Sub


Sub RestoreBG_Castle
    Controller.B2SSetData 205,1
    Controller.B2SSetData 179,1
    If bDorothyCaught=1 Then  Controller.B2SSetData 180,1
End Sub

Sub RestoreBG_Rescue
    Dim obj, bg_id
    Controller.B2SSetData 179,0
    Controller.B2SSetData 180,0
    Controller.B2SSetData 205,0
    Controller.B2SSetData 206,1
    For each obj in Lights_Castle
        bg_id = bg_dict.Item(obj)
        Controller.B2SSetData bg_id, obj.State
    Next
    If ballRescueLock>1 Then Controller.B2SSetData 241,1
    If ballRescueLock=2 Then Controller.B2SSetData 242,1
End Sub

Sub ClearBG_CB
    Controller.B2sSetData 244,0
    Controller.B2SSetData "CrystalBall_ON",0
    Controller.B2SSetData "CrystalBall_OFF",0
    Controller.B2SSetData "CrystalBall_Linked",0
    Controller.B2SSetData "CrystalBall_NoHold",0
    Controller.B2SSetData "CrystalBall_Frenzy",0
    Controller.B2sSetData 202,1
End Sub

Sub ClearBG_Haunted
    Dim i
    Controller.B2sSetData 202,0
    For i = 190 to 196
        Controller.B2SSetData i,0
    Next
End Sub

Sub ClearBG_HauntedModes
    Dim i
    For i = 245 to 250
        Controller.B2SSetData i,0
    Next
    Controller.B2SSetData "HF_Shots_Done",0
    Controller.B2SSetData "HF_Targets_Done",0
    Controller.B2SSetData "HF_Holes_Done",0
    Controller.B2SSetData "HF_Bumpers_Done",0
    ClearBG_HauntedDigits
End Sub

Sub ClearBG_HauntedDigits
    Dim i
    For i = 0 To 9
        H_Ones = "R1_" & Cstr(i)
        Controller.B2SSetData H_Ones,0
        If i<>0 Then
            H_Tens = "Rd_" & Cstr(i)
            Controller.B2SSetData H_Tens,0
        End If
    Next
End Sub

Sub ClearBG_BTWW
    Dim i, tmp_str
    BTWWShots_Off
    Controller.B2SSetData "WW_01",0
    For i = 3 To 9
       tmp_str = "WW_0" & Cstr(i)
       Controller.B2SSetData tmp_str,0
    Next
    BTWW_Jackpot_Clear
    Controller.B2SSetData 178,0
End Sub

Sub ClearBG_SOTR
    Dim i, tmp_str
    For i = 1 To 4
       tmp_str = "SOTR_0" & Cstr(i)
       Controller.B2SSetData tmp_str,0
    Next
    Controller.B2SSetData "SOTR_08",0
    For i = 1 to 3
       tmp_str = "SOTR_RBW_0" & Cstr(i)
       Controller.B2SSetData tmp_str,0
    Next
    For i = 1 to 4
       tmp_str = "SOTR_X" & Cstr(i)
       Controller.B2SSetData tmp_str,0
    Next
    For i = 1 to 7
       tmp_str = "SOTR_Ball_" & SOTR_color(i)
       Controller.B2SSetData tmp_str,0
    Next
    Controller.B2SSetdata 201,0
End Sub

'******************************************************************************
'       DISPLAY PF BACK MX EFFECT WORDS
'******************************************************************************
' Must be centralized to prevent multiple words being displayed at the same time

Dim current_word_str
current_word_str = "none"

Sub Display_MX(word_str)
    Dim i
    Select Case current_word_str
        Case "FRENZY"
            DOF 318,DOFOff
         Case "GLINDA"
            DOF 309,DOFOff
        Case "RESCUE"
            DOF 320,DOFOff
        Case "TWISTER"
            DOF 321,DOFOff
        Case "HAUNTED"
            DOF 304,DOFOff
        Case "EMERALD"
            DOF 306,DOFOff
        Case "CRYSTAL"
            DOF 305,DOFOff
        Case "JACKPOT"
            DOF 307,DOFOff
        Case "WIZARD"
            DOF 308,DOFOff
        Case "MUNCHKIN"
            DOF 319,DOFOff
        Case "RAINBOW"
            For i = 330 to 336
                DOF i, DOFOff
            Next
        Case "LOCK1"
            DOF 337,DOFOff
        Case "LOCK2"
            DOF 338,DOFOff
        Case "LOCK3"
            DOF 339,DOFOff
    End Select
    Select Case word_str
        Case "FRENZY"
            DOF 318,DOFON: vpmTimer.AddTimer 3000,"DOF 318, DOFOff '"
         Case "GLINDA"
            DOF 309,DOFON: vpmTimer.AddTimer 3000,"DOF 309, DOFOff '"
        Case "RESCUE"
            DOF 320,DOFON: vpmTimer.AddTimer 3000,"DOF 320, DOFOff '"
        Case "TWISTER"
            DOF 321,DOFON: vpmTimer.AddTimer 3000,"DOF 321, DOFOff '"
        Case "HAUNTED"
            DOF 304,DOFON: vpmTimer.AddTimer 3000,"DOF 304, DOFOff '"
        Case "EMERALD"
            DOF 306,DOFON: vpmTimer.AddTimer 3000,"DOF 306, DOFOff '"
        Case "CRYSTAL"
            DOF 305,DOFON: vpmTimer.AddTimer 3000,"DOF 305, DOFOff '"
        Case "JACKPOT"
            DOF 307,DOFON: vpmTimer.AddTimer 3000,"DOF 307, DOFOff '"
        Case "WIZARD"
            DOF 308,DOFON: vpmTimer.AddTimer 3000,"DOF 308, DOFOff '"
        Case "MUNCHKIN"
            DOF 319,DOFON: vpmTimer.AddTimer 3000,"DOF 319, DOFOff '"
        Case "RAINBOW"
            If L1.State=1 Then  DOF 336, DOFOn: vpmTimer.AddTimer 3000,"DOF 336, DOFOff '"
            If l2.State=1 Then  DOF 335, DOFOn: vpmTimer.AddTimer 3000,"DOF 335, DOFOff '"
            If L3.State=1 Then  DOF 334, DOFOn: vpmTimer.AddTimer 3000,"DOF 334, DOFOff '"
            If L4.State=1 Then  DOF 333, DOFOn: vpmTimer.AddTimer 3000,"DOF 333, DOFOff '"
            If L5.State=1 Then  DOF 332, DOFOn: vpmTimer.AddTimer 3000,"DOF 332, DOFOff '"
            If L6.State=1 Then  DOF 331, DOFOn: vpmTimer.AddTimer 3000,"DOF 331, DOFOff '"
            If L7.State=1 Then  DOF 330, DOFOn: vpmTimer.AddTimer 3000,"DOF 330, DOFOff '"
        Case "LOCK1"
            DOF 337,DOFON: vpmTimer.AddTimer 3000,"DOF 337, DOFOff '"
        Case "LOCK2"
            DOF 338,DOFON: vpmTimer.AddTimer 3000,"DOF 338, DOFOff '"
        Case "LOCK3"
            DOF 339,DOFON: vpmTimer.AddTimer 3000,"DOF 339, DOFOff '"
    End Select
    current_word_str = word_str
End Sub


'******************************************************************************
'       DOF PROCESSING
'******************************************************************************
' Set up Switch to DOF Map Dictionary
' The Switch can be used as the key with the associated DOF ID as the item
' The objective is to place all DOF assignments in one place (at least for most switches)
Dim dof_dict
Set dof_dict = CreateObject("Scripting.Dictionary")

' DOF Switch Assignments
'  28 ' Used for PUPpack - TNPLH_1 - Left Outlane Switch Start
'  57 ' Used for PUPpack - Rescue - Multiball
'  58 ' Used for PUPpack - SOTR Mode - 3 Rainbows!!!
'  59 ' Used for PUPpack - TOTO Late
'  60 ' Used for PUPpack - SOTR Instructions
'  61 ' Used for PUPpack - SOTR Main Screen
'  62 ' Used for PUPpack - SOTR Mode Ends
'  63 ' Used for PUPpack - DING DONG Witch is Dead
'  64 ' Used for PUPpack - Topper Jackpot
'  65 ' Used for PUPpack - Haunted Mode - Shots
'  66 ' Used for PUPpack - Haunted Mode - Targets
'  67 ' Used for PUPpack - Haunted Mode - Holes
'  68 ' Used for PUPpack - Haunted Mode - Bumpers
'  69 ' Used for PUPpack - EMC - Multiball
'  70 ' Used for PUPpack - BTWW - Starting Screen (WW_01)
'  71 ' Used for PUPpack - BTWW - BTWW_End
'  72 ' Used for PUPpack - EMC - Ball 1 Lock
'  73 ' Used for PUPpack - Munchkin Lullabye - MultiBall active
'  74 ' Used for PUPpack - MunchkinLand Welcome - MultiBall active
'  75 ' Used for PUPpack - EMC - Ball 2 Lock
'  78 ' Used for PUPpack - Fireball Frenzy - Multiball active
'  79 ' Used for PUPpack - Monkey Fly - MultiBall active
'  80 ' Used for PUPpack - Topper_Multiball
'  81 ' Used for PUPpack - Extra ball
'  82 ' Used for PUPpack - Special Award (Extra Game & SOTR Rainbow)
'  83 ' Used for PUPpack - Munchkin Lullabye - MultiBall NOT active
'  84 ' Used for PUPpack - MunchkinLand Welcome - MultiBall NOT active
'  85 ' Used for PUPpack - CB - Crystal Ball Mode Active (Take Warning)
'  86 ' Used for PUPpack - TNPLH_6 ToHouse Stop
'  87 ' Used for PUPpack - Melting
'  88 ' Used for PUPpack - Fireball Frenzy
'  89 ' Used for PUPpack - Monkey Fly - MultiBall NOT active
'  90 ' Used for PUPpack - TNPLH_5 ToShelter Stop
'  91 ' Used for PUPpack - TNPLH_6 ToHouse
'  92 ' Used for PUPpack - TNPLH_5 ToShelter
'  93 ' Used for PUPpack - TNPLH_4 Marvel Start
'  94 ' Used for PUPpack - TNPLH_3 Rainbow Start
'  95 ' Used for PUPpack - TNPLH_3 Rainbow Stop
'  96 ' Used for PUPpack - TNPLH_4 Marvel Stop
'  97 ' Used for PUPpack - TOTO Escape
'  98 ' Used for PUPpack - TOTO Stop
'  99 ' Used for PUPpack - TOTO Start
' 100 ' Used for PUPpack - Attract Mode Start

' 101 ' Left Flipper
' 102 ' Right Flipper
' 103 ' Left Slingshot
' 104 ' Left Slingshot Shake
' 105 ' Right Slingshot
' 106 ' Right Slingshot Shake
' 107 ' Bumper Back Left
' 108 ' Bumper Back Center
' 109 ' Bumper Back Right
' 110 ' Ball Release
' 111 ' AutoPlunger
' 112 ' White Flash on Strobe Toy for autoplunger, extra ball or extra game
' 113 ' Drain Light & PUPpack - TNPLH_1 Stop
' 114 ' Start Button Light, when credits is greater than 0
' 115 ' Bumper Middle Left
' 116 ' Bumper Front Left
' 117 ' Bumper Middle Center
dof_dict.Add 30, 118 ' Drop targets left hit
dof_dict.Add 31, 118 ' Drop targets left hit
dof_dict.Add 32, 118 ' Drop targets left hit
'dof_dict.Add 43, 118 ' Drop targets left hit
'dof_dict.Add 44, 118 ' Drop targets left hit
'dof_dict.Add 45, 118 ' Drop targets left hit
'dof_dict.Add 46, 118 ' Drop targets left hit
'dof_dict.Add 47, 118 ' Drop targets left hit
dof_dict.Add 55, 118 ' Drop targets left hit
dof_dict.Add 61, 118 ' Drop targets left hit
dof_dict.Add 62, 118 ' Drop targets left hit
dof_dict.Add 65, 118 ' Drop targets left hit
dof_dict.Add 66, 118 ' Drop targets left hit
dof_dict.Add 67, 118 ' Drop targets left hit
dof_dict.Add 68, 119 ' Drop targets center hit
dof_dict.Add 69, 119 ' Drop targets center hit
dof_dict.Add 70, 119 ' Drop targets center hit
dof_dict.Add 73, 119 ' Drop targets center hit
dof_dict.Add 89, 120 ' Drop targets right hit
dof_dict.Add 90, 120 ' Drop targets right hit
dof_dict.Add 91, 120 ' Drop targets right hit
dof_dict.Add 92, 120 ' Drop targets right hit
dof_dict.Add 93, 120 ' Drop targets right hit
dof_dict.Add 94, 120 ' Drop targets right hit
dof_dict.Add 95, 120 ' Drop targets right hit
' 121 ' Knocker for extra ball and award special
' 122 ' Triggertop1 Light
' 123 ' Triggertop2 Light
' 124 ' Triggertop3 Light
' 125 ' Strobe 500ms
' 126 ' Beacon when mutliball start
dof_dict.Add 57, 127 ' small shaker effect when magnets are ON
dof_dict.Add 58, 127 ' small shaker effect when magnets are ON
' 128 ' beacon effect
dof_dict.Add 41, 129 ' Left Bumper Flash
dof_dict.Add 50, 129 ' Left Bumper Flash
dof_dict.Add 51, 130 ' Center Bumper Flash
dof_dict.Add 49, 131 ' Right Bumper Flash
dof_dict.Add 28, 132 ' Outlane Left Light
dof_dict.Add 27, 133 ' Inlanes Left Light
dof_dict.Add 35, 134 ' Inlane Riht Light
dof_dict.Add 37, 135 ' Outlane Right Light
dof_dict.Add 38, 135 ' Outlane Right Light
dof_dict.Add 39, 135 ' Outlane Right Light
dof_dict.Add 40, 135 ' Outlane Right Light
' 136 ' flasher 2000 ms
' 137 ' flasher 1000 ms
' 138 ' Bumper Left Blink Flasher
' 139 ' Bumper Center Blink Flasher
' 140 ' Bumper Right Blink Flasher
' 141 ' RGB GI White
' 142 ' RGB GI Purple
' 143 ' RGB GI Darkblue
' 144 ' RGB GI Blue
' 145 ' RGB GI Green
' 146 ' RGB GI Darkgreen
' 147 ' RGB GI Yellow
' 148 ' RGB GI Amber
' 149 ' RGB GI Orange
' 150 ' RGB GI Red
' 151 ' Flasher Center Purple Blink
dof_dict.Add 43, 152 ' Drop targets left bottom hit
dof_dict.Add 44, 152 ' Drop targets left bottom hit
dof_dict.Add 45, 152 ' Drop targets left bottom hit
dof_dict.Add 46, 152 ' Drop targets left bottom hit
dof_dict.Add 47, 152 ' Drop targets left bottom hit
' 153 ' Flasher Right Purple Blink
' 154 ' Rear Flashers Blink 1000ms
' 155 ' Rear Flashers Blink 500ms
' 156 ' Launch Button Blink
dof_dict.Add 56, 157 ' TinMan Rollover
dof_dict.Add 42, 158 ' State Fair Rubber
dof_dict.Add 2, 158  ' Spinner
dof_dict.Add 71, 159 ' Castle Loop
dof_dict.Add 77, 160 ' Castle Door Bash
dof_dict.Add 81, 161 ' Left OZ lane
dof_dict.Add 82, 162 ' Right OZ Lane
dof_dict.Add 104, 162 ' Right OZ Lane & LION Rollover
dof_dict.Add 96, 163 ' SCARECROW Rollover & Munchkin PF Loop lower
dof_dict.Add 101, 163 ' SCARECROW Rollover & Munchkin PF Loop lower
dof_dict.Add 102, 164 ' Munchkin PF Loop upper
dof_dict.Add 103, 165 ' Throne Room Rubber
dof_dict.Add 15, 166 ' Scoop - Right Middle (Throne Room)
dof_dict.Add 17, 167 ' VUK - Top Center (Soldier) to Castle Playfield
dof_dict.Add 18, 168 ' Scoop - Left Top (Castle) & Left Middle (Owl)
dof_dict.Add 21, 168 ' Scoop - Left Top (Castle) & Left Middle (Owl)

' 178-231, 233, 235-250 Used by backglass
' Consequently 178 through 250 is reserved for the backglass stuff

' 300 ' Used for PUPpack - Game Over
' 301 ' Used for PUPpack - Attract Mode Stop
' 302 ' Used for PUPpack - Haunted Mode - Multiball
' 303 ' Used for PUPpack - Munchkin Mode - Multiball
' 304 ' Used for PF Back Effects MX - Haunted Modes
' 305 ' Used for PF Back Effects MX - Crystal Modes
' 306 ' Used for PF Back Effects MX - Emerald City MultiBall Mode
' 307 ' Used for PF Back Effects MX - Fireball Frenzy Jackpot
' 308 ' Used for PF Back Effects MX - WIZARD Mode (sw15)
' 309 ' Used for PF Back Effects MX - Glinda Mode

' 310 ' Used for PF Back Effects MX - GiColor amber
' 311 ' Used for PF Back Effects MX - GiColor blue
' 312 ' Used for PF Back Effects MX - GiColor green
' 313 ' Used for PF Back Effects MX - GiColor orange
' 314 ' Used for PF Back Effects MX - GiColor pink
' 315 ' Used for PF Back Effects MX - GiColor purple
' 316 ' Used for PF Back Effects MX - GiColor red
' 317 ' Used for PF Back Effects MX - GiColor yellow

' 318 ' Used for PF Back Effects MX - Fireball Frenzy Mode
' 319 ' Used for PF Back Effects MX - Munchkin Modes
' 320 ' Used for PF Back Effects MX - Rescue Mode
' 321 ' Used for PF Back Effects MX - TWISTER

' 322 ' Used for PUPpack - TNPLH_2  Stop

' 323 ' Used for PF Back Effects MX - Attract Mode
' 324 ' Used for PUPpack - Topper Background Activate
' 325 ' Used for PUPpack - Storm the Castle
' 326 ' Used for PUPpack - State Fair Bumper#4
' 327 ' Used for PUPpack - TNPLH_7 Complete Start
' 328 ' Used for PUPpack - TNPLH_7 Complete Stop
' 329 ' Used for PUPpack - TNPLH_2  Start

' 330 ' Used for PF Back Effects MX - "R" of RAINBOW
' 331 ' Used for PF Back Effects MX - "A" of RAINBOW
' 332 ' Used for PF Back Effects MX - "I" of RAINBOW
' 333 ' Used for PF Back Effects MX - "N" of RAINBOW
' 334 ' Used for PF Back Effects MX - "B" of RAINBOW
' 335 ' Used for PF Back Effects MX - "O" of RAINBOW
' 336 ' Used for PF Back Effects MX - "W" of RAINBOW
' 337 ' Used for PF Back Effects MX - "LOCK1"
' 338 ' Used for PF Back Effects MX - "LOCK2"
' 339 ' Used for PF Back Effects MX - "LOCK3"

' 340 ' Used for PUPpack - Ding Dong Mode - Witch Is Dead Stop
' 341 ' Used for PUPpack - SOTR is Ready (Seventh Jewel Awarded)

' 350 ' Witch Raise & Descend - intended for gear motor
' 351 ' House Rotate - intended for gear motor
' 352 ' Monkey Move - intended for gear motor
' 400 ' TILT
' 401 ' TILT Warning

'*******************************************************************************
'                       SWITCH ASSOCIATED DOF
'*******************************************************************************
Sub SetDOF(switch)
    Dim get_DOF
    If dof_dict.Exists(switch) Then
        get_dof = dof_dict.Item(switch)
        DOF get_dof, DOFPulse
    End If
End Sub


'************************************
' GAME OPTIONS  - NOT IMPLEMENTED
'************************************
' Ball Saver Option - Not Available ... Enabling will cause errors
Const BallSaveOption_ON = False
Const BallSaverTime = 30        ' seconds

'*******************************************************************************
'********************** END OF TABLE "STUFF" ***********************************
'*******************************************************************************


'*******************************************************************************
'       FEATURE TESTING SUPPORT  - ckpin
'*******************************************************************************
' TESTING OPTION - This is only used for feature and regression testing
' When activated (TEST_AUTOPLAY = True and Left MagnaSave Pressed), the number
' of balls per game increses to 10 and switches near some of the flippers activate
' the flippers.  If a ball is in the shooter lane, it will activate autoplunge.
' The automatic flippers may be turned off and on by pressing the Left MagnaSave.
' Despite the primitive approach it nearly always plays better than I do :)
' The right magnasave and some of the number keys are used to setup certain states
' A center post can also be activated with the 6 key..

Const TEST_MODE = False
Const TEST_AUTOPLAY = False
Const TEST_BallsPerGame = 10  ' Change to 10 or more during regression tests

Dim f1_Test, f2_Test,BTWW_tmp_cnt
Dim bDrainFlag, bAutoFlip
Dim SOTR_counter

bDrainFlag=0
bAutoFlip = 0
f1_Test=False
f2_Test=True
BTWW_tmp_cnt=0
SOTR_counter=0

CenterPost(0)
CenterPostP.Visible=False

' Action only occur when button released
Sub TestCode_KeyDown(keycode)

End Sub


Sub TestCode_KeyUp(keycode)
    Dim i

    debug.print "keycode = " & keycode

    ' Turn on Auto Regression Testing
	If keycode = LeftMagnaSave AND TEST_AUTOPLAY=True Then Auto_Regression_Flip


    ' Various testing support switch and mode activations
 	If keycode = RightMagnaSave AND TEST_MODE=True Then
'       AwardExtraBall
'       For i = 1 To 4:SwitchHitEvent (36+i):Next  ' TOTO
'       For i = 0 to 4:SwitchHitEvent (43+i):Next  ' TNPLH
        For i = 0 To 6:SwitchHitEvent (89+i):Next  ' RAINBOW
        For i = 0 To 5:SwitchHitEvent (65+i):Next  ' RESCUE
       SwitchHitEvent 61:SwitchHitEvent 62        ' Dorothy Capture Setup
       SwitchHitEvent 52                          ' Glinda Reward prep
       For i = 1 To 6:SwitchHitEvent 56:Next      ' TinMan Lights
       For i = 1 To 9:SwitchHitEvent 96:Next      ' Scarecrow Lights
       For i = 1 To 4:SwitchHitEvent 104:Next     ' Lion Lights

'       If UltraDMD_Option=True Then UltraDMD_simple "MultiballGreen.gif", "", "", 4000
'       bDisplay=1
'       DMD2Flush
'       DMD2 "", eNone, "HURRY UP", eNone, "THIS IS A TEST", eNone, _
'            "CENTER", eNone, 3000, True, ""
'       LightSeqAll.Play SeqAlloff
'       LightSeqAll.Play SeqCircleOutOn,10,5,0
'       LightSeqTwister.Play SeqScrewLeftOn,30
'       GiEffect 1, 10
'       LightEffect 1
    End If


    ' The 4 key  BTWW Test - Enter BTWW Mode
    If keycode = 5 AND TEST_MODE=True Then
      bBTWW_1=1:bBTWW_2=1:bBTWW_3=1:bBTWW_4=1
      BTWW
      CenterPost(1)
'    ' Turn on Flag controlling No Effect Drain Activate/DeActivate with the 6 key
'    ' Drain area blinks red when active and will blink blue after it is deactivated
'        If f1_Test Then
'            bDrainFlag = 0
'            SetLightColor gi20,"blue",2
'        Else
'            bDrainFlag = 1
'            SetLightColor gi20,"red",2
'       End If
'       f1_Test = Not(f1_Test)
    End If

    ' The 5 key  BTWW Mode Switches - Once BTWW Mode is Active
    If keycode = 6 AND TEST_MODE=True Then
        BTWW_tmp_cnt = BTWW_tmp_cnt + 1
        If BTWW_tmp_cnt>14 Then BTWW_tmp_cnt=1
        Select Case BTWW_tmp_cnt
          ' High Score Check - All White Shots
          Case 1    ' Witch Spinner - White
            SwitchHitEvent 9
          Case 2    ' Wizard
            SwitchHitEvent 15
          Case 3    ' Left Orbit
            SwitchHitEvent 64
          Case 4    ' Right Orbit
            SwitchHitEvent 87
          Case 5     ' Castle VUK
            SwitchHitEvent 17
          Case 6     ' Ramp
            SwitchHitEvent 76
          Case 7     ' Witch
            SwitchHitEvent 57
          ' Low Score Check - Red Shots
          Case 8    ' Witch Spinner - White
            SwitchHitEvent 9
          Case 9    ' Witch Spinner - White
            SwitchHitEvent 9
          Case 10    ' Witch Spinner - White
            SwitchHitEvent 9
          Case 11    ' Witch Spinner - White
            SwitchHitEvent 9
          Case 12    ' Castle VUK
            SwitchHitEvent 17
          Case 13    ' Ramp
            SwitchHitEvent 76
          Case 14     ' Witch
            SwitchHitEvent 57
        End Select
    End If


    ' The 6 key  -  CENTER POST TOGGLE
    If keycode = 7 AND TEST_MODE=True Then
        If CenterPostP.Visible=False Then
            CenterPost(1)
        Else
            CenterPost(0)
        End If
    End If


    ' The 7 key   SOTR MODE Test
    If keycode = 8 AND TEST_MODE=True Then
        SOTR_counter = SOTR_counter + 1
        If SOTR_counter>4 Then SOTR_counter=1
        Select Case SOTR_counter
          Case 1
            SOTR = 7: AwardJewel
          Case 2    ' Castle Rainbow
            For i = 0 To 5:SwitchHitEvent (65+i):Next  ' RESCUE
            ' OZ Lanes
            SwitchHitEvent 81: SwitchHitEvent 82
          Case 3    ' Target Rainbow
            For i = 0 To 6:SwitchHitEvent (89+i):Next  ' RAINBOW
          Case 4    ' Shot Rainbow
            SwitchHitEvent 9:SwitchHitEvent 24:SwitchHitEvent 73:SwitchHitEvent 74
            SwitchHitEvent 76:SwitchHitEvent 15:SwitchHitEvent 16
        End Select
    End If


End Sub


'******************************************************
'       Center Post (only for testing)
'******************************************************
Sub CenterPost(enabled)
	If enabled then
        CenterPostP.Visible=True
		CenterPostP.TransZ = 26
		Playsound SoundFX("Centerpost_Up",DOFContactors)
		CPC.IsDropped = 0
	Else
        CenterPostP.Visible=False
		Playsound SoundFX("Centerpost_Down",DOFContactors)
		CenterPostP.TransZ = 0
		CPC.IsDropped = 1
	End If
End Sub


'******************************************************************************
'                REGRESSION TESTING SUPPORT - ckpin
'******************************************************************************
' Set the OPTION TEST_AUTOPLAY to True
' Activate / DeActivate with Left MagnaSave then watch for correct responses
' and verify game doesn't fail.
' The switches below are used for Testing only; they are used to detect the
' ball passing by a flipper
' Auto Main flippers
Sub sw127_Hit:Auto_Flipper "R": End Sub    ' Auto Right Flipper
Sub sw130_Hit:Auto_Flipper "R": End Sub    ' Auto Right Flipper
Sub sw128_Hit:Auto_Flipper "L": End Sub    ' Auto Left Flipper
Sub sw129_Hit:Auto_Flipper "L": End Sub    ' Auto Left Flipper
Sub sw131_Hit     ' Auto Castle Flipper
   If bTOTOMode=0 AND bTNPLHMode=0 Then Auto_Flipper "R"
End Sub
Sub sw132_Hit    ' Auto Munchkin Flipper
    If bTOTOMode=0 AND bTNPLHMode=0 Then Auto_Flipper "L"
End Sub
Sub sw133_Hit:Auto_Flipper "R": End Sub    ' Auto COLLECT Flipper
Sub sw134_Hit    ' Auto Munchkin Flipper
    If bTOTOMode=0 AND bTNPLHMode=0 Then Auto_Flipper "L"
End Sub          ' Auto Castle Flipper
Sub sw135_Hit
    If bTOTOMode=0 AND bTNPLHMode=0 Then Auto_Flipper "R"
End Sub

Sub Auto_Regression_Flip
    If f2_Test Then
        AutoFireTest.enabled=True
        bAutoFlip = 1
    Else
        AutoFireTest.enabled=False
        bAutoFlip = 0
    End If
    f2_Test = Not f2_Test
End Sub

Sub Auto_Flipper(side)
    If bAutoFlip = 0 Then Exit Sub
    NoHold.enabled=True
    If side = "L" Then
        SolLFlipper 1
    Else
        SolRFlipper 1
    End If
End Sub


Sub AutoFireTest_Timer
    If bBallInPlungerLane=True Then
        PlungerIM.AutoFire
    End If
End Sub

' *****************************  END  TESTING SUPPORT **********************************************************

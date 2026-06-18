Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("core.vbs")
If Err Then MsgBox "Can't open core.vbs"
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

' Game Options
Const bUseDMD = True ' Use Freezy / FlexDMD display
Const bUseBackdropDisplay = False ' Use old backdrop display

'----- Virtual Reality Options -----
Dim bUseVR
' If you use an earlier version of VPX than 10.7.2, you must enable VR manually here
bUseVR = False
If version >= 10702 Then
  If RenderingMode = 2 Then bUseVR = True
End If

Dim VRRoomChoice : VRRoomChoice = 1
'1 - Minimal Room
'2 - Ultra minimal room


Const MusicVol = 0.25 ' Music Volume
'----- General Sound Options -----
Const VolumeDial = 0.8        'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

' CONTENT CONTRIBUTION CREDITS
' Rubberizer: iaakki
' Target Bouncer: iaakki, wrd1972, apophis
' Flipper and physics corrections: nFozzy, Rothbauerw
' Sound effects package: Fleep
' Ramp rolling sounds: nFozzy
' Bumpers: Flupper
' VR Cabinet & Room: Sixtoe, Flupper, 3rdaxis

Dim anScore       ' Array of players' scores
Dim nPlayers      ' Count of the number of players in the game
Dim nPlayersLastGame  ' Count of players last game
Dim nCurPlayer      ' Index of the player currently active
Dim nBall       ' Count of which ball is currently in play
Dim BIP         ' Number of balls in play
Dim nExtraBalls     ' Count of how many extra balls the current player has
Dim anExtraBallsLit   ' Array of counts of extra balls lit, per player
Dim abBossXBallWon    ' Array of Booleans: Has the player won an extra ball from bosses?
Dim bDrainingBalls    ' Is the table draining balls after a tilt or timed out multiball?
Dim anBonus       ' Array of incidental bonus, per player
Dim abBonusHeld     ' Is a player's bonus held from the preivous ball?
Dim nBonusX       ' bonus multiplier
Dim nTiltWarnings   ' Number of tilt warning used this ball
Dim nTiltLevel      ' How much has the game been nudged?
Dim nSkillShotState   ' Current state of the skill shot
Dim nColorCycleIndex  ' Index of which color the lights are during attract mode
Dim bPhysLockLeft   ' Is a ball is physically locked in the left saucer?
Dim bPhysLockRight    ' Is a ball is physically locked in the right saucer?
Dim avLocksLit      ' Array of counts of how many locks are lit for each player
Dim anLocksMade     ' Array of counts of how many balls are (virtually) locked for each player
Const cLockLeft = 0
Const cLockRight = 1
Const cLockBoth = 3
Dim bBallLaunched   ' Has the ball been plunged?
Dim bBallAtPlunger    ' Is a ball resting at the plunger lane?
Dim bPlungerAutoResting ' Has the plunger been pulled for an autoplunge?
Dim nAutoPlungeBalls  ' Count of balls to autoplunge
Dim nTimeGateLoopL    ' The time when the left loop switch was last triggered
Dim nTimeGateLoopR    ' The time when the right loop switch was last triggered
Dim bUpperHoleHit   ' Was the shop kicker hit through the upper hole?
Dim avBottomLanes   ' Array of bitsets of which bottom lanes are hit per player
Dim aoBottomLaneLights  ' Array of light objects for the bottom lanes
Dim anKnightMBalls    ' Array of counts of how many Shovel Knight Mballs have been played
Dim anKnightSJacks    ' Array of counts of Shovel Knight Super Jackpots, per player
Dim sSJMultiplier   ' String to display for super jackpot multiplier (2X, 3X, etc)
Dim nSJValue      ' Value of super jackpot, after multiplier
Dim avXUsed       ' Array of bitsets of which shots have expired 2X, per player
Dim aanXTimers(3, 9)  ' Array of arrays of timers for 2X. Index 1: player, index 2: shot
Dim anXCompleted    ' Array of how many times each players have completed every multiplier
Dim vMultipliersLit   ' Bitset of which shots will activate its shot multiplier when hit
Dim bBallHeld     ' Is the ball grabbed by a saucer (pausing timers)?
Dim bCountingBonus    ' Is the game currently displaying the bonus count?
Dim nBonusScreen    ' Count of which screen of the bonus count is currently showing
Dim bUsingShop      ' Is the player currently using the shop?
Dim aanBossMedals(3, 9) ' Array of arrays of boss medals. Index 1: player, index 2: boss
Dim aanTotalMedals(3, 9)' Array of arrays of boss medals that have been reset by Enchantress mode
Dim anModeTimeLeft(18)  ' Array of time (in milliseconds) remaining on modes
Dim anGracePeriod(18) ' Array of times (in ms) of grace period remaining on modes
Dim asModeProgress(18)  ' Array of strings to use to display the progress of each mode
Dim anModeScore(18)   ' Array of points scored during each mode
Dim anBossShots(9)    ' Array of shots made for each boss
Dim nBallSaveTimer    ' Time (in milliseconds) remaining on ball saver
Dim bBallSaveUsed   ' Has the ball saver been used?
Dim bOutlaneSave    ' Boolean of if a ball has been saved after an outlane drain
Dim vModesRunning   ' Bitset of which modes are currently running
Dim vPrevModesRunning ' Bitset of which modes were running last display text update
Dim anMostRecentLevel ' Array of the levels most recently played, per player
Dim anMostRecentBoss  ' Array of the boss most recently lit, per player
Dim anLevelSelected   ' Array of the level selected with the right flipper, per player
Dim avBossesLit     ' Array of bitsets of which bosses are ready to Start, per player
Dim bBossStarted    ' Boolean: Are we kicking out the ball after a boss was started?
Dim vTreasureHits   ' Bitset of which shots have been hit in Treasure Knight boss battle
Dim nBlackKnightCycle ' Index of the cycle of the K-N-I-G-H-T targets during Black Knight boss
Dim anLevelShotsToBeat  ' Array of counts of shots required to finish a level, per player
Dim anLevelShotsMade  ' Array of counts of shots made towards finishing a level, per player
Dim dLevelCountUp   ' Where the counting up of the level progress is on the display
Dim nCountUpShots   ' Number of shots to display in count up (may not be actual # of shots)
Dim nCountUpGoal    ' Goal for shots to display in count up (may not be actual # of shots)
Dim abPlayingLevel    ' Is player is working on lighting a mode? Array, per player
Dim aanWandererJPs(3,4) ' Array of arrays of wanderer jackpots. index 1 = player, 2 = wanderer
Dim avWanderersBeaten ' Array of bitsets of wanderers that have jackpots scored, per player
Dim anWanderersPlayed ' Array of counts of wanderer multiballs played, per player
Dim anWandererHits    ' Array of hits left to light wanderer, per player
Dim anWandererSelected  ' Array of which wanderer will be started next (or is running), per player
Dim avSecretProgress  ' Array of bitsets of goals toward secret battle, per player
Dim bShowingInstantInfo ' Is the display currently showing instant info?
Dim nInstantInfoPage    ' Count of which page instant info is currently showing
Const cLastInfoPage = 29
Const cMedalNotPlayed = 0
Const cMedalConsolation = 1
Const cMedalBronze = 2
Const cMedalSilver = 3
Const cMedalGold = 4
Dim avShovelTargetsHit  ' Array of bitsets of which SHOVEL KNIGHT targets have been hit, per player
Dim avShovelLightsMask  ' Array of Bitsets of which SHOVEL KNIGHT targets will do nothing, per player
Dim aoShovelLights    ' Array of light objects for SHOVEL KNIGHT targets
Dim AoShovelTargets   ' Array of target objects for SHOVEL KNIGHT targets
Dim nUpperPFHits    ' Count of how many upper playfield switches have been hit w/o a lower
Dim avTrouppleHits    ' Array of bitsets of which troupple king targets are hit, per player
Dim anTrouppleLit   ' Array of how many mystery awards are lit, per player
Dim abTrouppleXBall   ' Array of Booleans of if mystery has given an extra ball this game, per player
Dim anTrouppleAwards  ' Array of counts of mystery awards collected, per player
Dim bAddABallUsed   ' Boolean: Was an add-a-ball awarded this mutilball?
Dim aoShotLights    ' Array of light objects for shots
Dim anShovelJackpots  ' Array of indeces for shots that have a Shovel Knight Mball Jackpot lit
Dim bShovelSuperJP    ' Is the Shovel Knight Mball super Jackpot lit?
Dim aoXLights     ' Array of light objects for 2X shot multipliers
Dim anColorCycleLight ' Array of colors of lights turned on in the color cycle
Dim anColorCycleDark  ' Array of colors of lights turned off in the color cycle
Dim avShotColors    ' Array of bitsets for which colors to blink per shot
Dim anShotIdleColor(9)  ' Array of which color to show solid when there are no blinks, per shot
Dim nTimeDisplay    ' Time in milliseconds of how long the display has shown a message cycle
Dim nScoreBlinks    ' Count of which digit of the score to blink on the scoreboard
Dim sTextTop      ' The text on the top line of the display
Dim sTextBottom     ' The text on the bottom line of the display
Dim aoTopRowDigits    ' Top row of flasher objects in VR alphanumeric display
Dim aoBottomRowDigits ' Bottom row of flasher objects in VR alphanumeric display
Dim nTextDuration   ' Time in ms of how much longer to display the current message
Dim nTextPriority ' a low number is a high priority, 1 means top priority
          ' 99 means "only display if otherwise idle"
Dim oQueueTextTop   ' Queue for messages to show on top row of display
Dim oQueueTextBottom  ' Queue for messages to show on bottom row of display
Dim oQueueTextDuration  ' Queue for how long, in ms, to display the message in oQueueText*
Dim oQueueStatusTop   ' Queue for status messages to show on top row of display
Dim oQueueStatusBottom  ' Queue for status messages to show on bottom row of display
Dim oQueueStatusDuration' Queue for how long, in ms, to display the message in oQueueStatus*
Dim sModeProgress   ' Mode-specific progress to display
Dim nModeTime     ' Time remaining on mode to display
Dim anComboShots(2)   ' The three last shots made as part of a combo
Dim nComboCurrent   ' Count of the number of shots in the current combo
Dim nComboJPCurrent   ' Point value of the current combo jackpot
Dim anComboCount    ' Array of counts of combo shots made, per player
Dim anComboJPSum    ' Array of total points awarded from combo jackpots, per player
Dim anComboTimers(9)  ' Array of time in milliseconds remaining to hit a combo, per shot
Dim anJuggleTimers(9) ' Array of time in ms remaining for each shot in juggling mode
Dim nJuggleLives    ' Count of drains remaining before juggling mode is lost
Dim aDisplayTopLine   ' Array of EMReel objects for letters on top row of display
Dim aDisplayBottomLine  ' Array of EMReel objects for letters on top row of display
const PlayUntilHit = -1 ' If a text's duration is set to this, display the text until a switch hit
Dim aSegments(128)    ' Array of display segments for backglass, index = ASCII value of letter
Dim nMusicCurrentTrack  ' Index of current music track
Dim nMusicNextTrack   ' Index of next music track to play, usually loops the current track
Dim aoMusicTimers   ' Array of timer objects with loop times for the music tracks
Dim asMusicFiles    ' Array of file names of the music tracks
Dim bIsMusicLoading   ' Is the music preloading?
Dim nTimeLowerLeftFlip  ' The time when the lower left flipper was last raised
Dim nTimeLowerRightFlip ' The time when the lower right flipper was last raised
Dim nTimeUpperLeftFlip  ' The time when the upper left flippers were last raised
Dim nTimeUpperRightFlip ' The time when the upper right flipper was last raised
Dim nTimeBumperHit    ' the time when a bumper was last hit
Dim nTimeLastSwitch   ' The time when any switch was last hit
Dim nTimeStartPressed ' The time when the start button was pressed
Dim oCaptiveBall    ' The captive ball object
Dim bEnteringHighScore  ' Are we currently entering a high score?

aoMusicTimers = Array(MusicTimer00, MusicTimer01, MusicTimer02, MusicTimer03, _
  MusicTimer04, MusicTimer05, MusicTimer06, MusicTimer07, MusicTimer08, MusicTimer09,_
  MusicTimer10, MusicTimer11, MusicTimer12, MusicTimer13, MusicTimer14, MusicTimer15,_
  MusicTimer16, MusicTimer17, MusicTimer18, MusicTimer19, MusicTimer20, MusicTimer21,_
  MusicTimer22, MusicTimer23, MusicTimer24, MusicTimer25, MusicTimer26, MusicTimer27,_
  MusicTimer28, MusicTimer29, MusicTimer30, MusicTimer31, MusicTimer32, MusicTimer33,_
  MusicTimer34)
asMusicFiles = Array("shovelknight00", "shovelknight01", "shovelknight02",_
  "shovelknight03", "shovelknight04", "shovelknight05",_
  "shovelknight06", "shovelknight07", "shovelknight08",_
  "shovelknight09", "shovelknight10", "shovelknight11",_
  "shovelknight12", "shovelknight13", "shovelknight14",_
  "shovelknight15", "shovelknight16", "shovelknight17",_
  "shovelknight18", "shovelknight19", "shovelknight20",_
  "shovelknight21", "shovelknight22", "shovelknight23",_
  "shovelknight24", "shovelknight25", "shovelknight26",_
  "shovelknight27", "shovelknight28", "shovelknight29",_
  "shovelknight30", "shovelknight31", "shovelknight32",_
  "shovelknight33", "shovelknight34")
Const StopMusic = -1  'Call InitMusic with this param to stop all music
Dim sDebugText

Const cHighScoreAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ <"
Const cGameName = "shovelknight"
Const nHighScoreSlots = 4
Dim HighScore(3)
Dim HighScoreName(3)
Dim SortedHighscores()
Dim SortedHighscoreNames()
Dim aScoreRanks
Dim nHighScoreIter
Dim nTimePressedLeft  ' For entering highscores: when L button was pressed
Dim nTimePressedRight ' For entering highscores: when R button was pressed
Dim HighscoreLetter   ' The index of the current letter in the cHighScoreAlphabet
Dim InitialsEntered     ' The initials being entered in the highscore

Dim aPowersOfTwo
' Note that &h8000 is -32768 in VBScript. WHAT. THE. FUCK.
aPowersOfTwo = Array(1, 2, 4, 8, &h10, &h20, &h40, &h80,_
        &h100, &h200, &h400, &h800, &h1000, &h2000, &h4000, 32768,_
        &h10000, &h20000, &h40000, &h80000, &h100000, &h200000, &h400000, &h800000,_
        &h1000000, &h2000000, &h4000000, &h8000000, &h10000000, &h20000000, &h40000000)
const cAllShovelTargets = &h3f  '000000111111
const cAllKnightTargets = &hfc0 '111111000000
aoShovelLights = Array(Light_S, LightH1, LightO, LightV, LightE, LightL,_
          LightK, LightN, LightI, LightG, LightH2, LightT)
AoShovelTargets = Array(StandupS, StandupH1, StandupO, StandupV, StandupE, StandupL, _
  StandupK, StandupN, StandupI, StandupG, StandupH2, StandupT)
aoShotLights = Array(LightLLoop, LightLRamp, LightCaptive, LightCRamp, LightShop,_
  LightSuperJP, LightVUK, LightRSaucer, LightRLoop, LightRRamp)
aoXLights = Array(LightXLoopL, LightXRampL, LightXCaptive, LightXRampC, LightXShop,_
  LightXUpperHole, LightXVUK, LightXSaucer, LightXLoopR, LightXRampR)
aoBottomLaneLights = Array(LightBottomLane1, LightBottomLane2, LightBottomLane3,_
  LightBottomLane4, LightBottomLane5)
Dim aoBossLights
aoBossLights = Array(LightBlack, LightKing, LightSpectre, LightPlague, LightTreasure, _
  LightMole, LightTinker, LightPolar, LightPropeller)
Dim aoTrouppleLights
aoTrouppleLights = Array(LightTroupple1, LightTroupple2, LightTroupple3)
Const cShotNone = -1
Const cShotLoopL = 0
Const cShotRampL = 1
Const cShotCaptive = 2
Const cShotRampC = 3
Const cShotShop = 4
Const cShotUpperHole = 5
Const cShotVUK = 6
Const cShotSaucerR = 7
Const cShotLoopR = 8
Const cShotRampR = 9
Dim asShotDescriptions ' names of shots to use on display
asShotDescriptions = Array("LEFT LOOP", "LEFT RAMP", "CAPTIVE BALL", "CENTER RAMP", _
  "CENTER LANE", "UPPER HOLE", "VERTICAL KICKER", "RIGHT SAUCER", "RIGHT LOOP", _
  "RIGHT RAMP")
Dim anJugglingShots
anJugglingShots = Array(cShotRampL, cShotRampC, cShotShop, cShotVUK, cShotSaucerR)

Const cSwitchShot = 1
Const cSwitchTarget = 2
Const cSwitchLane = 3
Const cSwitchBumper = 4
Const cSwitchSpinner = 5

Const cSkillShotDisabled = 1
Const cSkillShotNoneHit = 2
Const cSkillShotHiddenHit = 3
Const cSkillShotHiddenMissed = 4

Const cMystery5k = 1
Const cMystery10k = 2
Const cMysteryBonusX = 3
Const cMysteryLightLock = 4
Const cMysteryShotX = 5
Const cMysteryBonusHeld = 6
Const cMysteryXBall = 7
Const cMysteryLightBoss = 8
Const cMysteryMoreTime = 9
Const cMysteryBallSaver = 10
Const cMysteryAddABall = 11

Const cModeNone = -1
Const cModeBlack = 0
Const cModeKing = 1
Const cModeSpectre = 2
Const cModePlague = 3
Const cModeTreasure = 4
Const cModeMole = 5
Const cModeTinker = 6
Const cModePolar = 7
Const cModePropeller = 8
Const cModeEnchantress = 9
Const cModeShovel = 10 ' Shovel Knight multiball
Const cModeWanderer = 11
Const cModeChampions = 12 ' Hall of Champions challenge
Const cModeJuggling = 13 ' Shop sold out: juggling challenge
Const cModeSecret = 14
Const cModeCombo = 15
Const cModeSuperCombo = 16
Const cModeGoldTargets = 17
Const cModeGoldRamps = 18
Dim anTimedModes
anTimedModes = Array(cModeBlack, cModeKing, cModeSpectre, cModePlague, cModeTreasure,_
  cModeMole, cModeTinker, cModePolar, cModePropeller, cModeEnchantress, cModeChampions,_
  cModeSuperCombo)
Dim asLevelNames ' Array of level names to display
asLevelNames = Array("PLAINS / PASSAGE", "PRIDEMOOR KEEP", "LICH YARD", _
  "EXPLODATORIUM", "IRON WHALE", "THE LOST CITY", "CLOCKWORK TOWER", _
  "STRANDED SHIP", "FLYING MACHINE", "FINAL TOWER")
Dim asKnightNames ' Array of boss names to display
asKnightNames = Array("BLACK KNIGHT", "KING KNIGHT", "SPECTRE KNIGHT", _
  "PLAGUE KNIGHT", "TREASURE KNIGHT", "MOLE KNIGHT", "TINKER KNIGHT", _
  "POLAR KNIGHT", "PROPELLER KNIGHT", "ENCHANTRESS")
Dim anBossLightOrder ' Order of boss lights on playfield, starting from Black Knight
           ' and going down
anBossLightOrder = Array(cModeBlack, cModePlague, cModePropeller, cModeTreasure, cModeMole, _
  cModeTinker, cModeSpectre, cModePolar, cModeKing)
Dim anLightOfBoss ' Opposite of anBossLightOrder - array index is boss, element
           ' content is the index the boss' light has in the aoBossLights array
anLightOfBoss = Array(0, 8, 6, 1, 3, 4, 5, 7, 2)
Dim anSilverScores ' Array of shots required to get a silver medal for a boss
anSilverScores = Array(3, 16, 2, 8, 2, 2, 3, 3, 2)
Dim anGoldScores ' Array of shots required to get a gold medal for a boss
anGoldScores = Array(5, 32, 4, 16, 3, 4, 5, 6, 5)
Dim anBossTimers ' Array of default timer for boss
anBossTimers = Array(32000, 32000, 32000, 32000, 27000, 32000, 32000, 32000, 32000)
Dim asWandererNames
asWandererNames = Array("BLACK KNIGHT", "PHANTOM STRIKER", "RIEZE", "BAZ", "SUPER WANDERER")
Dim aoWandererLights
aoWandererLights = Array(LightWanderer1, LightWanderer2, LightWanderer3, LightWanderer4)

Const cWandererNone = -1
Const cWandererBlack = 0
Const cWandererPhantom = 1
Const cWandererReize = 2
Const cWandererBaz = 3
Const cWandererAll = 4
Const cSecret3xSuper = 0
Const cSecret8xJackpot = 1
Const cSecretEnchantress = 2
Const cSecretWanderer = 3
Const cSecretJuggling = 4
Const cSecretChampions = 5

aoBottomLaneLights = Array(LightBottomLane1, LightBottomLane2, LightBottomLane3,_
  LightBottomLane4, LightBottomLane5)
Const Color_LtGreen = &h33FF33
Const Color_LtRed = &h3333FF
Const Color_LtBlue = &hFF3333
Const Color_LtYellow = &h33FFFF
Const Color_Yellow = &h01FFF2
Const Color_LtAmber = &h33BBFF
Const Color_Amber = 42751 ' &hA6FF
Const Color_LtOrange = &h3377FF
Const Color_Orange = &h55FF
Const Color_LtPink = &ha366FF
Const Color_Pink = &ha366FF
Const Color_LtPurple = &hFF9933
Const Color_Purple = &hFF0080
Const Color_Magenta = &hCC00FF
Const Color_LtMagenta = &hF033FF
Const Color_LtCyan = &hFFFF33
Const Color_Gray = &h999999
anColorCycleLight = Array(Color_LtRed, Color_LtOrange, Color_LtAmber, Color_LtYellow,_
  Color_LtGreen, Color_LtCyan, Color_LtBlue, Color_LtPurple, Color_LtMagenta, Color_LtPink, vbWhite)
anColorCycleDark = Array(vbRed, Color_Orange, Color_Amber, vbYellow,_
  vbGreen, vbCyan, vbBlue, Color_Purple, Color_Magenta, Color_Pink, Color_Gray)
Const cBlinkRed = 0
Const cBlinkOrange = 1
Const cBlinkAmber = 2
Const cBlinkYellow = 3
Const cBlinkGreen = 4
Const cBlinkCyan = 5
Const cBlinkBlue = 6
Const cBlinkPurple = 7
Const cBlinkMagenta = 8
Const cBlinkPink = 9
Const cBlinkWhite = 10
Const nColorsInCycle = 11

Const cBaseShovelJP = 4000 ' Base Shovel Knight Multiball jackpot
Const cBaseSuperJP = 30000 ' Base Shovel Knight Multiball super jackpot
Const cJugglingGoal = 60000 ' Sum of combo points needed to light juggling mode
Const cBallSaveStart = 12000 ' Ball save timer at start of ball
Const cBallSaveMball = 9000 ' Ball save timer at start of multiball
Dim BallSize : BallSize = 50
Dim BallMass : BallMass = 1
Const tnob = 7            'Total number of balls
Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Sub InitGame(bFastRestart)
  Dim i, j
  anScore = Array(0, 0, 0, 0)
  nPlayers = 1
  nCurPlayer = 0
  nBall = 1
  nExtraBalls = 0
  anExtraBallsLit = Array(0, 0, 0, 0)
  abBossXBallWon = Array(False, False, False, False)
  anBonus = Array(0, 0, 0, 0)
  abBonusHeld = Array(False, False, False, False)

  anKnightMBalls = Array(0, 0, 0, 0)
  anKnightSJacks = Array(0, 0, 0, 0)
  avShovelTargetsHit = Array(0, 0, 0, 0)
  avShovelLightsMask = Array(0, 0, 0, 0)
  avTrouppleHits = Array(0, 0, 0, 0)
  anTrouppleLit = Array(0, 0, 0, 0)
  abTrouppleXBall = Array(False, False, False, False)
  anTrouppleAwards = Array(0, 0, 0, 0)
  LightLRamp.color = Color_Gray
  LightLRamp.colorFull = vbWhite
  LightCRamp.color = Color_Gray
  LightCRamp.colorFull = vbWhite
  bBallAtPlunger = bFastRestart
  If Not bFastRestart Then
    bPhysLockLeft = False
    bPhysLockRight = False
  End If
  avLocksLit = Array(0, 0, 0, 0)
  anLocksMade = Array(0, 0, 0, 0)
  avBottomLanes = Array(0, 0, 0, 0)
  avBossesLit = Array(0, 0, 0, 0)
  avSecretProgress = Array(0, 0, 0, 0)
  bBossStarted = False
  abPlayingLevel = Array(False, False, False, False)
  anLevelShotsToBeat = Array(0, 0, 0, 0)
  anLevelShotsMade = Array(0, 0, 0, 0)
  avXUsed = Array(0, 0, 0, 0)
  anXCompleted = Array(0, 0, 0, 0)
  anShovelJackpots = Array(-1, -1, -1)
  anComboCount = Array(0, 0, 0, 0)
  anComboJPSum = Array(0, 0, 0, 0)
  avWanderersBeaten = Array(0, 0, 0, 0)
  anWanderersPlayed = Array(0, 0, 0, 0)
  anWandererHits = Array(60, 60, 60, 60)
  anWandererSelected = Array(cWandererBlack, cWandererBlack, cWandererBlack, cWandererBlack)
  bShovelSuperJP = False
  bBallHeld = False
  Set oQueueTextTop = New Queue
  Set oQueueTextBottom = New Queue
  Set oQueueTextDuration = New Queue
  Set oQueueStatusTop = New Queue
  Set oQueueStatusBottom = New Queue
  Set oQueueStatusDuration = New Queue
  sModeProgress = ""
  nModeTime = 0
  anMostRecentLevel = Array(cModeNone, cModeNone, cModeNone, cModeNone)
  anMostRecentBoss = Array(cModeNone, cModeNone, cModeNone, cModeNone)
  anLevelSelected = Array(cModeBlack, cModeBlack, cModeBlack, cModeBlack)
  bShowingInstantInfo = False
  nInstantInfoPage = 0

  nTimeGateLoopR = -1
  nTimeBumperHit = 0
  nTimeLastSwitch = 0
  bUpperHoleHit = False
  DropShop.isDropped = True
  bEnteringHighScore = False
  for i = 0 to 3
    for j = 0 to 9
      aanBossMedals(i, j) = cMedalNotPlayed
      aanTotalMedals(i ,j) = 0
      aanXTimers(i, j) = 0
    next
    for j = 0 to 4
      aanWandererJPs(i, j) = 0
    next
  next

  RaiseDrops

  LightSeq1.StopPlay
  avShotColors = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
  for i = 0 to 9
    aoShotLights(i).color = Color_Gray
    aoShotLights(i).colorFull = vbWhite
    aoXLights(i).color = Color_Gray
    aoXLights(i).colorFull = vbWhite
    anShotIdleColor(i) = -1
  next
  for i = 0 to 4
    aoBottomLaneLights(i).color = Color_Gray
    aoBottomLaneLights(i).colorFull = vbWhite
  next
  for i = 0 to 11
    aoShovelLights(i).color = vbGreen
    aoShovelLights(i).colorFull = Color_LtGreen
  next

  InitBall bFastRestart
End Sub

Sub Table1_KeyDown(ByVal keycode)
  Dim i

  If keycode = PlungerKey Then
    SoundPlungerPull
    Plunger.PullBack
  End If

  If nPlayers > 0 And not bDrainingBalls Then

    If keycode = LeftFlipperKey Then
      If Not bCountingBonus Then
        FlipperActivate LeftFlipper, LFPress
        SolLFlipper True
        DOF 101, DOFOn
        UpperLeftFlipper.RotateToEnd
        If False = BitSetContains(vModesRunning, cModeJuggling) Then
          LeftSmallFlipper.RotateToEnd
        End If
      End If
      RotLeft avBottomLanes(nCurPlayer), 5
      UpdateLightsLanes
      nTimeUpperLeftFlip = GameTime
      ' double flip skip if flips are less than 50 ms apart
      If (GameTime - nTimeUpperRightFlip) < 50 Then
        DoubleFlipHurry
      End If
      If bShowingInstantInfo Then
        nInstantInfoPage = (nInstantInfoPage + 1) mod (cLastInfoPage + 1)
        nTimeDisplay = 0
      End If
    End If

    If keycode = RightFlipperKey Then
      If Not bCountingBonus Then
        FlipperActivate RightFlipper, RFPress
        SolRFlipper True
        If False = BitSetContains(vModesRunning, cModeJuggling) Then
          RightSmallFlipper.RotateToEnd
        End If
        DOF 102, DOFOn
      End If
      RotRight avBottomLanes(nCurPlayer), 5
      UpdateLightsLanes
      ChangeLevelToStart
      nTimeUpperRightFlip = GameTime
      ' double flip skip if flips are less than 50 ms apart
      If (GameTime - nTimeUpperLeftFlip) < 50 Then
        DoubleFlipHurry
      End If
      If bShowingInstantInfo Then
        nInstantInfoPage = (nInstantInfoPage + 1) mod (cLastInfoPage + 1)
        nTimeDisplay = 0
      End If
    End If

    If keycode = LeftMagnaSave Then
      nTimeLowerLeftFlip = GameTime
      RotLeft avBottomLanes(nCurPlayer), 5
      UpdateLightsLanes
      If nTimeUpperLeftFlip = -1 And bCountingBonus = False Then
        FlipperActivate LeftFlipper, LFPress
        SolLFlipper True
        DOF 101, DOFOn
      End If
    End If

    If keycode = RightMagnaSave Then
      nTimeLowerRightFlip = GameTime
      RotRight avBottomLanes(nCurPlayer), 5
      UpdateLightsLanes
      If nTimeUpperRightFlip = -1 And bCountingBonus = False Then
        FlipperActivate RightFlipper, RFPress
        SolRFlipper True
        DOF 102, DOFOn
      End If
    End If

    If keycode = LeftTiltKey Then
      Nudge 90, 2
      SoundNudgeLeft
      If BIP > 0 and not bDrainingBalls Then HandleTilt
    End If

    If keycode = RightTiltKey Then
      Nudge 270, 2
      SoundNudgeRight
      If BIP > 0 and not bDrainingBalls Then HandleTilt
    End If

    If keycode = CenterTiltKey Then
      Nudge 0, 3
      SoundNudgeCenter
      If BIP > 0 and not bDrainingBalls Then HandleTilt
    End If

    If keycode = MechanicalTilt Then
      If BIP > 0 and not bDrainingBalls Then HandleTilt
    End If

'   if keycode = AddCreditKey Then
'     avSecretProgress(nCurPlayer) = BitSetAdd(avSecretProgress(nCurPlayer), cSecret3xSuper)
'     avSecretProgress(nCurPlayer) = BitSetAdd(avSecretProgress(nCurPlayer), cSecretChampions)
'     avSecretProgress(nCurPlayer) = BitSetAdd(avSecretProgress(nCurPlayer), cSecretEnchantress)
'     avSecretProgress(nCurPlayer) = BitSetAdd(avSecretProgress(nCurPlayer), cSecretJuggling)
'     avSecretProgress(nCurPlayer) = BitSetAdd(avSecretProgress(nCurPlayer), cSecretWanderer)
'     UpdateLightsBoss
'   End If

    If keycode = StartGameKey Then
      SoundStartButton
      If nBall = 1 and nPlayers < 4 Then
        PlaySound "fx-start"
        nPlayers = nPlayers + 1
      End If
      nTimeStartPressed = GameTime
    End If

  Else
    If keycode = LeftFlipperKey And bEnteringHighScore Then
      HighScoreLeftPressed
    End If

    If keycode = RightFlipperKey And bEnteringHighScore Then
      HighScoreRightPressed
    End If

    If keycode = PlungerKey And bEnteringHighScore Then
      HighScoreEnterPressed
    End If

    If keycode = StartGameKey And bIsMusicLoading = False Then
      SoundStartButton
      If bEnteringHighScore Then
        HighScoreEnterPressed
      ElseIf bDrainingBalls = False Then
        InitGame False
      End If
    End If
  End If

End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    If bBallAtPlunger Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If

  If keycode = LeftFlipperKey Then
    If bEnteringHighScore Then
      HighScoreLeftReleased
    Else
      nTimeUpperLeftFlip = -1
      UpperLeftFlipper.RotateToStart
      LeftSmallFlipper.RotateToStart
      If nTimeLowerLeftFlip = -1 Then
        FlipperDeActivate LeftFlipper, LFPress
        SolLFlipper False
      End If
      If nPlayers > 0 And bCountingBonus = False And _
      bEnteringHighScore = False And bDrainingBalls = False Then
        DOF 101, DOFOff
      End If
    End If
  End If

  If keycode = RightFlipperKey Then
    If bEnteringHighScore Then
      HighScoreRightReleased
    Else
      nTimeUpperRightFlip = -1
      RightSmallFlipper.RotateToStart
      If nTimeLowerRightFlip = -1 Then
        FlipperDeActivate RightFlipper, RFPress
        SolRFlipper False
      End If
      If nPlayers > 0 And bCountingBonus = False And _
      bEnteringHighScore = False And bDrainingBalls = False Then
        DOF 102, DOFOff
      End If
    End If
  End If

  If keycode = LeftMagnaSave Then
    nTimeLowerLeftFlip = -1
    If nTimeUpperLeftFlip = -1 And bCountingBonus = False And _
    bEnteringHighScore = False And bDrainingBalls = False Then
      FlipperDeActivate LeftFlipper, LFPress
      SolLFlipper False
      DOF 101, DOFOff
    End If
  End If

  If keycode = RightMagnaSave Then
    nTimeLowerRightFlip = -1
    If nTimeUpperRightFlip = -1 And bCountingBonus = False And _
    bEnteringHighScore = False And bDrainingBalls = False Then
      FlipperDeActivate RightFlipper, RFPress
      SolRFlipper False
      DOF 102, DOFOff
    End If
  End If

  If keycode = StartGameKey Then
    nTimeStartPressed = 0
  End If

End Sub

Sub DoubleFlipHurry
  If bCountingBonus Then
    nBonusScreen = 6
    TimerBonusCount.enabled = False
    TimerBonusCount_Timer
  Elseif bBossStarted Then
    ' Skip messages
    nTextDuration = 0
    ClearTextQueue
    ' kick out from right saucer
    TimerEjectBallR_Timer
  End If
End Sub

Sub Table1_Init
  Dim i
  Randomize
  LoadEM
    Set oCaptiveBall = New cvpmCaptiveBall
    With oCaptiveBall
        .InitCaptive CapTrigger1, CapWall1, Array(CapKicker1, CapKicker1a), 345
        .NailedBalls = 1
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "oCaptiveBall"
        .Start
    End With
    CapKicker1.CreateBall

  InitVR
  InitDisplay
  If bUseDMD And (Not bUseVR) Then InitFlexAlpha

  nPlayers = 0
  Diverter1R.isDropped = True
  Diverter1L.isDropped = False
  nTimeLowerLeftFlip = -1
  nTimeLowerRightFlip = -1
  nTimeUpperLeftFlip = -1
  nTimeUpperRightFlip = -1
  nMusicCurrentTrack = -1
  bIsMusicLoading = False
  InitMusic 2
  nColorCycleIndex = 0
  FlBumperFadeTarget(1) = 0
  FlBumperFadeTarget(2) = 0
  FlBumperFadeTarget(3) = 0
  LightSeq1.UpdateInterval = 25
  LightSeq1.Play SeqCircleOutOn, 40
  anScore = Array(0, 0, 0, 0)
  nCurPlayer = 0
  nPlayersLastGame = 4
  TimerDisplay.enabled = True
  Loadhs
End Sub

Sub Table1_Exit
  If b2son then Controller.Stop
End Sub

Sub LightSeq1_PlayDone
  Dim i
  dim colorDark, colorFull
  If nPlayers < 1 Then
    nColorCycleIndex = (nColorCycleIndex + 1) mod nColorsInCycle
    colorDark = anColorCycleDark(nColorCycleIndex)
    colorFull = anColorCycleLight(nColorCycleIndex)
    for i = 0 to 9
      aoShotLights(i).color = colorDark
      aoShotLights(i).colorFull = colorFull
      aoXLights(i).color = colorDark
      aoXLights(i).colorFull = colorFull
    next
    for i = 0 to 4
      aoBottomLaneLights(i).color = colorDark
      aoBottomLaneLights(i).colorFull = colorFull
    next
  for i = 0 To 11
    aoShovelLights(i).color = colorDark
    aoShovelLights(i).colorFull = colorFull
  next
    LightSeq1.Play SeqCircleOutOn, 40
  End If
End Sub

Sub InitBall(bFastRestart)
  Dim i
  bDrainingBalls = False
  nTiltWarnings = 0
  nTiltLevel = 0
  bCountingBonus = False
  bUsingShop = False

  If abBonusHeld(nCurPlayer) Then
    abBonusHeld(nCurPlayer) = False
  Else
    anBonus(nCurPlayer) = 10
  End If
  nBonusX = 1

  KickerAutoPlunge.enabled = False
  If False = bFastRestart Then
    BallRelease.CreateSizedBallWithMass Ballsize / 2, BallMass
    BallRelease.Kick 90, 7
    RandomSoundBallRelease BallRelease
    DOF 123, DOFPulse
  End If
  BIP = 1
  nTimeGateLoopL = -1
  nTimeGateLoopR = -1
  bBallLaunched = False
  bPlungerAutoResting = False
  bBallSaveUsed = False
  bOutlaneSave = False
  nSkillShotState = cSkillShotNoneHit
  AddShotColor cShotVUK, cBlinkRed
  nAutoPlungeBalls = 0
  timerMode.enabled = true
  timerCombo.enabled = true
  vMultipliersLit = 0
  vModesRunning = 0
  If abPlayingLevel(nCurPlayer) = True Then
    nCountUpShots = anLevelShotsMade(nCurPlayer)
    nCountUpGoal = anLevelShotsToBeat(nCurPlayer)
  Else
    nCountUpShots = 0
    nCountUpGoal = 0
  End If
  dLevelCountUp = anLevelShotsMade(nCurPlayer)
  for i = 0 to 2 : anComboShots(i) = cShotNone : Next
  for i = 0 to 9 : anComboTimers(i) = 0 : Next
  nUpperPFHits = 0
  nComboCurrent = 0
  nComboJPCurrent = 0
  vPrevModesRunning = 0
  SelectMusic False

  ' Set lights and diverters to proper state when switching players
  RaiseDrops
  UpdateLightsBoss
  UpdateLightsLanes
  UpdateLightsStandups
  UpdateLightsLocks
  UpdateLightsMultipliers
  UpdateLightsWanderer
  UpdateRampC
  UpdateRampL
  UpdateDiverter

  for i = 0 to 18
    anModeTimeLeft(i) = 0
    anGracePeriod(i) = 0
    asModeProgress(i) = ""
    anModeScore(i) = 0
  Next
End Sub

Sub TimerRestart_Timer
  If nPlayers > 0 And nBall > 1 And True = bBallAtPlunger And BIP < 2 Then
    If nTimeStartPressed > 0 And ((GameTime - nTimeStartPressed) > 1999) Then
      PlaySound "fx-start"
      InitGame True
    End If
  End If
End Sub

Sub AutoPlunge(nBalls)
  If nAutoPlungeBalls < 1 Then
    nAutoPlungeBalls = nAutoPlungeBalls + nBalls
    KickerAutoPlunge.enabled = True
    AutoPlungePull
  Else
    nAutoPlungeBalls = nAutoPlungeBalls + nBalls
  End If
End Sub

Sub AutoPlungePull
  BallRelease.CreateSizedBallWithMass Ballsize / 2, BallMass
  BallRelease.Kick 90, 7
  RandomSoundBallRelease BallRelease
  DOF 123, DOFPulse
  Plunger.PullBack
  bPlungerAutoResting = False
  vpmTimer.AddTimer 1000, "AutoPlungeFire '"
End Sub

Sub AutoPlungeFire
  Plunger.Fire
  bBallLaunched = False
  bPlungerAutoResting = True
  nAutoPlungeBalls = nAutoPlungeBalls - 1
  If nAutoPlungeBalls > 0 Then
    vpmTimer.AddTimer 500, "AutoPlungePull '"
  Else
    nAutoPlungeBalls = 0
  End If
End Sub

Sub KickerAutoPlunge_Hit
  KickerAutoPlunge.TimerEnabled = True
End Sub

Sub KickerAutoPlunge_Timer
  KickerAutoPlunge.TimerEnabled = False
  SoundSaucerKick 1, KickerAutoPlunge
  DOF 121, DOFPulse
  DOF 122, DOFPulse
  KickerAutoPlunge.kick 0, 45
End Sub

Sub Drain_Hit()
  Dim i
  dim vModes
  If BIP < 2 And nPlayers > 0 And (Not bOutlaneSave) Then PlaySound "fx-death",0,1,0,0.05
  Drain.DestroyBall

  ' Handle ball that drained in an outlane during ball save
  If bOutlaneSave Then
    BIP = BIP + 1
    bOutlaneSave = False
  ' handle ball saver
  ElseIf nBallSaveTimer > 0 Then
    If BIP < 2 Then
      If nExtraBalls > 0 Then
        LightShootAgain.state = LightStateOn
      Else
        LightShootAgain.state = LightStateOff
      End If
      bBallSaveUsed = True
      If nSkillShotState <> cSkillShotDisabled Then
        RemoveShotColor cShotVUK, cBlinkRed
      End If
      If nSkillShotState <> cSkillShotDisabled Then
        RemoveShotColor cShotVUK, cBlinkRed
      End If
      nSkillShotState = cSkillShotDisabled
      nBallSaveTimer = 0
      ShowText "BALL SAVED", "PLAYER " & (nCurPlayer + 1), 2500, 2
      PlaySound "fx-swoosh"
    End If
    AutoPlunge 1
    Exit Sub
  Elseif BitSetContains(vModesRunning, cModeJuggling) Then
    nJuggleLives = nJuggleLives - 1
    If nJuggleLives < 1 Then
      BIP = BIP - 1
      EndModes BitSetAdd(0, cModeJuggling)
    Elseif 1 = nJuggleLives Then
      ShowText "BALL MISSED", "LAST CHANCE", 3000, 2
      AutoPlunge 1
    Else
      ShowText "BALL MISSED", nJuggleLives & " MISSES LEFT", 3000, 2
      AutoPlunge 1
    End If
    Exit Sub
  End If

  BIP = BIP - 1
  ' Handle balls kicked out from locks after a game over
  If nPlayers = 0 Then
    If BIP <= 0 Then
      BIP = 0
      bDrainingBalls = False
    End If
    Exit Sub
  End If

  If Not BitSetContains(vModesRunning, cModeSecret) Then
    If BIP = 1 And BitSetContains(vModesRunning, cModeShovel) = True Then
      nSkillShotState = cSkillShotDisabled
      vModes = 0
      vModes = BitSetAdd(vModes, cModeShovel)
      EndModes(vModes)
    End If

    If BIP = 1 And BitSetContains(vModesRunning, cModeWanderer) = True Then
      nSkillShotState = cSkillShotDisabled
      EndModes(BitSetAdd(0, cModeWanderer))
    End If
  End If

  If BIP = 0 then
    If bDrainingBalls Then
      Leftslingshot.disabled = False
      Rightslingshot.disabled = False
      Bumper1.HasHitEvent = True
      Bumper2.HasHitEvent = True
      Bumper3.HasHitEvent = True
      bDrainingBalls = False
      ' handle tilt
      If nTiltWarnings > 2 Then
        ClearTextQueue
        ClearStatus
        ' Skip to end of bonus, which is when going to the next player
        ' and ending the game are handled
        nBonusScreen = 6
        TimerBonusCount.enabled = False
        TimerBonusCount_Timer
      Else
        BIP = 1
        KickerAutoPlunge.enabled = False
        BallRelease.CreateSizedBallWithMass Ballsize / 2, BallMass
        BallRelease.Kick 90, 7
        RandomSoundBallRelease BallRelease
        DOF 123, DOFPulse
        bBallLaunched = False
      End If
      Exit Sub
    End If

    ' Disable shot multipliers
    For i = 0 to 9
      If BitSetContains(vMultipliersLit, i) Then
        RemoveShotColor i, cBlinkWhite
        vMultipliersLit = BitSetRemove(vMultipliersLit, i)
      End If

      If aanXTimers(nCurPlayer, i) > 0 Then
        aanXTimers(nCurPlayer, i) = 0
        EndMultiplier(i)
      End If
    Next

    ' End boss modes
    anComboTimers(cShotLoopL) = 0
    ClearStatus
    ClearTextQueue
    For each i in anTimedModes
      if BitSetContains(vModesRunning, i) Then
        anModeTimeLeft(i) = 0
        vModes = BitSetAdd(0, i)
        EndModes(vModes)
      End If
    Next
    If BitSetContains(vModesRunning, cModeSecret) Then
      vModes = BitSetAdd(0, cModeChampions)
      vModes = BitSetAdd(vModes, cModeWanderer)
      vModes = BitSetAdd(vModes, cModeShovel)
      vModes = BitSetAdd(vModes, cModeEnchantress)
      EndModes vModes
      EndModes BitSetAdd(0, cModeSecret)
    End If
    EndCombos
    RemoveShotColor cShotLoopL, cBlinkOrange

    ' Kill flippers
    LeftFlipper.RotateToStart
    UpperLeftFlipper.RotateToStart
    LeftSmallFlipper.RotateToStart
    RightFlipper.RotateToStart
    RightSmallFlipper.RotateToStart

    ' Count up bonus
    AllLightsOff
    bCountingBonus = True
    nBonusScreen = 1
    TimerBonusCount.Enabled = True
    ' show first screen right now
    ClearTextQueue
    SelectMusic False
    TimerBonusCount_Timer
  End If
End Sub

Sub TimerBonusCount_Timer
  Dim nMedalBonus
  Dim nMballBonus
  Dim nShotXBonus
  Dim nTotalBonus
  Dim i

  nMedalBonus = 0
  nMedalBonus = nMedalBonus + 10 * aanTotalMedals(nCurPlayer, cMedalConsolation)
  nMedalBonus = nMedalBonus + 50 * aanTotalMedals(nCurPlayer, cMedalBronze)
  nMedalBonus = nMedalBonus + 150 * aanTotalMedals(nCurPlayer, cMedalSilver)
  nMedalBonus = nMedalBonus + 300 * aanTotalMedals(nCurPlayer, cMedalGold)

  nMballBonus = 200 * anKnightMBalls(nCurPlayer)

  nShotXBonus = 1000 * anXCompleted(nCurPlayer)
  for i = 0 to 9
    If BitSetContains(avXUsed(nCurPlayer), i) Then
      nShotXBonus = nShotXBonus + 100
    End If
  Next

  ' Round bonus up to nearest mutiple of 10
  If anBonus(nCurPlayer) mod 10 <> 0 Then
    anBonus(nCurPlayer) = Int((anBonus(nCurPlayer) + 10) / 10) * 10
  End If

  nTotalBonus = nMedalBonus + nMballBonus + nShotXBonus + anBonus(nCurPlayer)

  Select Case nBonusScreen
    Case 1
      ShowText "MEDAL BONUS", WilliamsFormatNum(nMedalBonus), 1200, 1
    Case 2
      PlaySound "fx-chargeup"
      ShowText "MULTIBALL", "BONUS " & WilliamsFormatNum(nMballBonus), 1200, 1
    Case 3
      PlaySound "fx-chargeup"
      ShowText "SHOT MULTIPLIER", "BONUS " & WilliamsFormatNum(nShotXBonus), 1200, 1
    Case 4
      PlaySound "fx-chargeup"
      ShowText "TARGET BONUS", WilliamsFormatNum(anBonus(nCurPlayer)), 1200, 1
    Case 5
      PlaySound "fx-chargeup"
      ShowText "TOTAL " & WilliamsFormatNum(nTotalBonus), _
      "X" & nBonusX & " = " & WilliamsFormatNum(nBonusX * nTotalBonus), 1200, 1
    Case 6
      PlaySound "fx-spellcast"
      anScore(nCurPlayer) = anScore(nCurPlayer) + (nBonusX * nTotalBonus)
      ShowText "P" & (nCurPlayer + 1) & " TOTAL SCORE", WilliamsFormatNum(anScore(nCurPlayer)), 2000, 1
      bCountingBonus = False
      TimerBonusCount.enabled = False
      ' next Player
      ' Now that the bonus sequence is done, end the ball
      if nCurPlayer = nPlayers - 1 and nBall = 3 and nExtraBalls = 0 then
        nPlayersLastGame = nPlayers
        nPlayers = 0
        QueueText "GAME OVER", "", 5000
        CheckHighScore(anScore)
        If bPhysLockLeft Then
          RaiseRampL
          KickerLRamp.timerEnabled = True
          BIP = BIP + 1
        End If
        If bPhysLockRight Then
          RaiseRampR
          TimerEjectBallC.enabled = True
          BIP = BIP + 1
        End If
        If BIP > 0 Then bDrainingBalls = True
        exit sub
      elseif nExtraBalls > 0 then
        nExtraBalls = nExtraBalls - 1
        If nExtraBalls < 1 Then
          LightShootAgain.state = LightStateOff
        End If
        ShowText "PLAYER " + cStr(nCurPlayer + 1), "SHOOT AGAIN", PlayUntilHit, 1
      else
        if nCurPlayer = nPlayers - 1 then
          nCurPlayer = 0
          nBall = nBall + 1
        else
          nCurPlayer = nCurPlayer + 1
        end if
      end if
      InitBall False
    Case Else
      ShowText "BONUS IS BUGGED", "FIX THIS NOW", 2000, 1
      bCountingBonus = False
      TimerBonusCount.enabled = False
  End Select

  nBonusScreen = nBonusScreen + 1
End Sub

Sub HandleTilt
  Dim i

  If BIP < 1 Then Exit Sub
  nTiltLevel = nTiltLevel + 10
  If nTiltLevel > 15 Then
    nTiltWarnings = nTiltWarnings + 1
    Select Case nTiltWarnings
      Case 1
        PlaySound "fx-tiltwarning"
        ShowText "DANGER", "", 2000, 1
      Case 2
        PlaySound "fx-tiltwarning"
        ShowText "DANGER", "DANGER", 2000, 1
      Case 3
        bDrainingBalls = True
        nBallSaveTimer = 0

        For i = 0 to 9
          If BitSetContains(vMultipliersLit, i) Then
            RemoveShotColor i, cBlinkWhite
            vMultipliersLit = BitSetRemove(vMultipliersLit, i)
          End If

          If aanXTimers(nCurPlayer, i) > 0 Then
            aanXTimers(nCurPlayer, i) = 0
            EndMultiplier(i)
          End If
        Next

        ' End Wizard modes
        if BitSetContains(vModesRunning, cModeEnchantress) Then
          anModeTimeLeft(cModeEnchantress) = 1
        End If

        ' End boss modes
        ClearTextQueue
        For each i in anTimedModes
          if BitSetContains(vModesRunning, i) Then
            anModeTimeLeft(i) = 0
            EndModes BitSetAdd(0, i)
          End If
        Next
        EndCombos
        anComboTimers(cShotLoopL) = 0
        RemoveShotColor cShotLoopL, cBlinkOrange
        ClearTextQueue
        nTextDuration = 0
        InitMusic StopMusic
        ShowText "TILT", "", 500, 1
        QueueText "", "TILT", 500
        QueueText "TILT", "", 500
        QueueText "*loop*", "", 500
        PlaySound "fx-crash"

        AllLightsOff
        LightSeq1.play SeqDownOff
        LeftSlingshot.disabled = True
        RightSlingshot.disabled = True
        Bumper1.HasHitEvent = False
        Bumper2.HasHitEvent = False
        Bumper3.HasHitEvent = False
        LeftFlipper.RotateToStart
        UpperLeftFlipper.RotateToStart
        LeftSmallFlipper.RotateToStart
        RightFlipper.RotateToStart
        RightSmallFlipper.RotateToStart
    End Select
  End If
End Sub

Sub Tilt_Timer
  Tilt.Enabled = True
  If nTiltLevel > 0 Then nTiltLevel = nTiltLevel - 1
End Sub


Sub AnyBumperHit
  If bDrainingBalls Then Exit Sub

  Dim nPoints
  If aanXTimers(nCurPlayer, cShotLoopL) > 0 Then
    anScore(nCurPlayer) = anScore(nCurPlayer) + 50
  Else
    anScore(nCurPlayer) = anScore(nCurPlayer) + 100
  End If

  nTimeBumperHit = GameTime
  If bBallLaunched = False Then
    bBallLaunched = True
    If bBallSaveUsed = False And BIP < 2 Then
      If abPlayingLevel(nCurPlayer) = false And anLevelShotsToBeat(nCurPlayer) < 2 Then
        StartLevel(anLevelSelected(nCurPlayer))
      End If
      nBallSaveTimer = cBallSaveStart
      LightShootAgain.state = LightStateBlinking
    End If
  End If

  ' Wanderer countdown
  If False = BitSetContains(vModesRunning, cModeEnchantress) _
  And False = BitSetContains(vModesRunning, cModeSecret) _
  And False = BitSetContains(vModesRunning, cModeJuggling) _
  And False = BitSetContains(vModesRunning, cModeWanderer) Then
    ChangeWanderer
    If anWandererHits(nCurPlayer) > 0 Then
      ShowText anWandererHits(nCurPlayer) & " POPS / SPINS", "TO LITE WANDERER", 1000, 9
      anWandererHits(nCurPlayer) = anWandererHits(nCurPlayer) - 1
      If anWandererHits(nCurPlayer) < 1 Then
        PlaySound "fx-fanfare1"
        ShowText "WANDERER IS LIT", "SHOOT UPPER HOLE", 2500, 3
        UpdateLightsWanderer
      End If
    End If
  End If

  ' Bosses: Plague Knight
  If BitSetContains(vModesRunning, cModePlague) Or _
  BitSetContains(vModesRunning, cModeEnchantress) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    nPoints = 600
    If aanXTimers(nCurPlayer, cShotLoopL) > 0 Then
      nPoints = nPoints * 2
    End If
    If BitSetContains(vModesRunning, cModePlague) = False Then
      nPoints = nPoints * 0.6
    End If
    anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
    anModeScore(cModePlague) = anModeScore(cModePlague) + nPoints
    anBossShots(cModePlague) = anBossShots(cModePlague) + 1
    If BitSetContains(vModesRunning, cModeSecret) = False Then
      ShowText asKnightNames(cModePlague), _
        "T=%t TOTAL=" & WilliamsFormatNum(anModeScore(cModePlague)), 2500, 4
    Else
      anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
      CheckSecret
    End If
    UpdateModeDisplay False
  End If
End Sub

Sub Bumper1_Hit
  If bDrainingBalls Then Exit Sub

  If BitSetContains(vModesRunning, cModePlague) Then
    PlaySound "fx-thud", 0, 0.4
  Else
    FlBumperFadeTarget(1) = 1   'Flupper bumper
    Bumper1.TimerEnabled = 1
  End If
  RandomSoundBumperBottom Bumper1
  DOF 111, DOFPulse
  DOF 112, DOFPulse
  AnyBumperHit
  AnySwitchHit cSwitchBumper, 0
End Sub

Sub Bumper1_Timer
  FlBumperFadeTarget(1) = 0   'Flupper bumper
  Bumper1.Timerenabled = 0
End Sub

Sub Bumper2_Hit
  If bDrainingBalls Then Exit Sub
  If BitSetContains(vModesRunning, cModePlague) Then
    PlaySound "fx-thud", 0, 0.4
  Else
    FlBumperFadeTarget(2) = 1   'Flupper bumper
    Bumper2.TimerEnabled = 1
  End If
  RandomSoundBumperTop Bumper2
  DOF 109, DOFPulse
  DOF 110, DOFPulse
  AnyBumperHit
  AnySwitchHit cSwitchBumper, 1
End Sub

Sub Bumper2_Timer
  FlBumperFadeTarget(2) = 0   'Flupper bumper
  Bumper2.Timerenabled = 0
End Sub

Sub Bumper3_Hit
  If bDrainingBalls Then Exit Sub
  If BitSetContains(vModesRunning, cModePlague) Then
    PlaySound "fx-thud", 0, 0.4
  Else
    FlBumperFadeTarget(3) = 1   'Flupper bumper
    Bumper3.TimerEnabled = 1
  End If
  RandomSoundBumperMiddle Bumper3
  DOF 107, DOFPulse
  DOF 110, DOFPulse
  AnyBumperHit
  AnySwitchHit cSwitchBumper, 2
End Sub

Sub Bumper3_Timer
  FlBumperFadeTarget(3) = 0   'Flupper bumper
  Bumper3.Timerenabled = 0
End Sub

Sub Spinner1_Spin
  Dim nPoints

  If aanXTimers(nCurPlayer, cShotLoopL) > 0 Then
    anScore(nCurPlayer) = anScore(nCurPlayer) + 30
  Else
    anScore(nCurPlayer) = anScore(nCurPlayer) + 60
  End If

  If bDrainingBalls Then Exit Sub

  PlaySound "fx-jump", 0, 0.3
  SoundSpinner Spinner1
  AnySwitchHit cSwitchSpinner, 0

  ' Wanderer countdown
  If False = BitSetContains(vModesRunning, cModeEnchantress) _
  And False = BitSetContains(vModesRunning, cModeSecret) _
  And False = BitSetContains(vModesRunning, cModeJuggling) _
  And False = BitSetContains(vModesRunning, cModeWanderer) Then
    ChangeWanderer
    If anWandererHits(nCurPlayer) > 0 Then
      ShowText anWandererHits(nCurPlayer) & " POPS / SPINS", "TO LITE WANDERER", 1000, 9
      anWandererHits(nCurPlayer) = anWandererHits(nCurPlayer) - 1
      If anWandererHits(nCurPlayer) < 1 Then
        PlaySound "fx-fanfare1"
        ShowText "WANDERER IS LIT", "SHOOT UPPER HOLE", 2500, 3
        UpdateLightsWanderer
      End If
    End If
  End If

  ' King Knight boss mode
  If BitSetContains(vModesRunning, cModeKing) Or _
  BitSetContains(vModesRunning, cModeEnchantress) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    nPoints = 320
    If aanXTimers(nCurPlayer, cShotLoopR) > 0 Then
      nPoints = nPoints * 2
    End If
    If BitSetContains(vModesRunning, cModeKing) = False Then
      nPoints = nPoints * 0.6
      nPoints = (nPoints \ 10) * 10 ' round down to multiple of 10
    End If
    anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
    anModeScore(cModeKing) = anModeScore(cModeKing) + nPoints
    anBossShots(cModeKing) = anBossShots(cModeKing) + 1
      If BitSetContains(vModesRunning, cModeSecret) = False Then
        ShowText asKnightNames(cModeKing), _
          "T=%t TOTAL=" & WilliamsFormatNum(anModeScore(cModeKing)), 2500, 4
    Else
      anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
      CheckSecret
    End If
    UpdateModeDisplay False
  End If
End Sub

'****Targets

Sub StandupCaptiveBall_Hit
  ShotHit cShotCaptive
End Sub

Sub ShovelTargetHit(targetIndex)
  Dim bSoundPlayed : bSoundPlayed = false
  Dim nPoints

  If bDrainingBalls Then Exit Sub
  anScore(nCurPlayer) = anScore(nCurPlayer) + 30

  ' Check if drop bank is complete
  If StandupS.isDropped And StandupH1.isDropped And StandupO.isDropped Then
    PlaySound "fx-spellcast", 0, 0.8
    nBonusX = nBonusX + 1
    If nBonusX > 10 Then
      nBonusX = 10
      anScore(nCurPlayer) = anScore(nCurPlayer) + 5000
      ShowText "BONUS MULTIPLIER", "MAXED   e000", 2000, 5
    Else
      ShowText "BONUS MULTIPLIER", nBonusX & " X", 2000, 5
    End If
    RaiseDrops
  End If

  AnySwitchHit cSwitchTarget, targetIndex
  If Not BitSetContains(avShovelLightsMask(nCurPlayer), targetIndex) Then
    If not BitSetContains(avShovelTargetsHit(nCurPlayer), targetIndex) Then
      If targetIndex < 3 Then
        DOF 115, DOFPulse
        SoundDropTargetDrop AoShovelTargets(targetIndex)
        PlaySound "fx-shovelswing", 0, 0.3
      Else
        PlaySound "fx-shovelswing", 0, 0.3
      End If
      bSoundPlayed = true
      anScore(nCurPlayer) = anScore(nCurPlayer) + 70
    End if
    avShovelTargetsHit(nCurPlayer) = BitSetAdd(avShovelTargetsHit(nCurPlayer), targetIndex)
    UpdateLightsStandups
    CheckSKStandups
  End If

  Dim nWanderer, nWandererGoal, nWandererPrevValue
  nWanderer = anWandererSelected(nCurPlayer)
  If nWanderer = cWandererAll Then
    nWandererGoal = 40000
  Else
    nWandererGoal = 10000
  End If
  ' Black Knight Wanderer
  If (BitSetContains(vModesRunning, cModeWanderer) Or _
  BitSetContains(vModesRunning, cModeSecret)) And _
  nUpperPFHits < 7 Then
    If nWanderer = cWandererAll Or nWanderer = cWandererBlack Then
      nWandererPrevValue = aanWandererJPs(nCurPlayer, nWanderer)
      nPoints = 200
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      if BitSetContains(vModesRunning, cModeSecret) then
        anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
        CheckSecret
      end if
      aanWandererJPs(nCurPlayer, nWanderer) = _
        aanWandererJPs(nCurPlayer, nWanderer) + 750

      If aanWandererJPs(nCurPlayer, nWanderer) >= nWandererGoal _
      And nWandererPrevValue < nWandererGoal Then
        ShowText asWandererNames(nWanderer), "JACKPOT IS LIT", 2500, 3
        PlaySound "fx-fanfare2"
        UpdateModeDisplay True
      End If

      UpdateLightsWanderer
    End If
  End If

  ' Black Knight boss
  If BitSetContains(vModesRunning, cModeBlack) Or _
  BitSetContains(vModesRunning, cModeEnchantress) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    If aoShovelLights(targetIndex).state = LightStateBlinking And targetIndex > 5 Then
      anBossShots(cModeBlack) = anBossShots(cModeBlack) + 1
      nPoints = 1200 * anBossShots(cModeBlack)
      If BitSetContains(vModesRunning, cModeBlack) = False Then
        nPoints = nPoints * 0.6
      End If
      StopSound "fx-shovelswing"
      PlaySound "fx-thud", 0, 0.6
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      anModeScore(cModeBlack) = anModeScore(cModeBlack) + nPoints
      If BitSetContains(vModesRunning, cModeSecret) = False Then
        ShowText asKnightNames(cModeBlack), _
          "T=%t TOTAL=" & WilliamsFormatNum(anModeScore(cModeBlack)), 2500, 4
      Else
        anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
        CheckSecret
      End If
      UpdateModeDisplay False
      If anBossShots(cModeBlack) = 5 Then
        If BitSetContains(vModesRunning, cModeBlack) Then
          ' time out mode instantly as no more targets can be hit
          anModeTimeLeft(cModeBlack) = 1
        Else
          anBossShots(cModeBlack) = 0
        End If
      End If
      TimerKnightCycle_Timer
    End If
  End If

  ' Polar Knight Boss
  If BitSetContains(vModesRunning, cModePolar) Or _
  BitSetContains(vModesRunning, cModeEnchantress) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    If targetIndex < 6 Then
      anBossShots(cModePolar) = anBossShots(cModePolar) + 1
      nPoints = 2500
      If BitSetContains(vModesRunning, cModePolar) = False Then
        nPoints = nPoints * 0.6
      End If
      StopSound "fx-shovelswing"
      PlaySound "fx-thud", 0, 0.6
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      anModeScore(cModePolar) = anModeScore(cModePolar) + nPoints
      If BitSetContains(vModesRunning, cModeSecret) = False Then
        ShowText asKnightNames(cModePolar), _
          "T=%t TOTAL=" & WilliamsFormatNum(anModeScore(cModePolar)), 2500, 4
      Else
        anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
        CheckSecret
      End If
      UpdateModeDisplay False
      UpdateLightsStandups
    End If
  End If

  If Not bSoundPlayed Then
    If targetIndex < 3 Then
      DOF 115, DOFPulse
      SoundDropTargetDrop AoShovelTargets(targetIndex)
    End If
  End If
End Sub

Sub StandupS_Dropped
  ShovelTargetHit 0
End Sub

Sub StandupH1_Dropped
  ShovelTargetHit 1
End Sub

Sub StandupO_Dropped
  ShovelTargetHit 2
End Sub

Sub StandupV_Hit
  ShovelTargetHit 3
End Sub

Sub StandupE_Hit
  ShovelTargetHit 4
End Sub

Sub StandupL_Hit
  ShovelTargetHit 5
End Sub

Sub StandupK_Hit
  ShovelTargetHit 6
End Sub

Sub StandupN_Hit
  ShovelTargetHit 7
End Sub

Sub StandupI_Hit
  ShovelTargetHit 8
End Sub

Sub StandupG_Hit
  ShovelTargetHit 9
End Sub

Sub StandupH2_Hit
  ShovelTargetHit 10
End Sub

Sub StandupT_Hit
  ShovelTargetHit 11
End Sub

Sub DropShop_Dropped
  If nPlayers > 0 Then
    UpdateDiverter
    If anKnightSJacks(nCurPlayer) mod 3 = 2 _
    And BitSetContains(vModesRunning, cModeShovel) Then
      ShotHit cShotShop
    End If
  End If
End Sub

Sub AnyTrouppleHit
  Dim nPoints

  If bDrainingBalls Then Exit Sub

  anScore(nCurPlayer) = anScore(nCurPlayer) + 30

  ' Mole Knight boss mode
  If BitSetContains(vModesRunning, cModeMole) Or _
  BitSetContains(vModesRunning, cModeEnchantress) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    nPoints = 4000
    If BIP > 1 Then
      nPoints = nPoints * 0.5
    End If
    If BitSetContains(vModesRunning, cModemole) = False Then
      nPoints = nPoints * 0.6
      nPoints = (nPoints \ 10) * 10 ' round down to multiple of 10
    End If

    anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
    anModeScore(cModeMole) = anModeScore(cModeMole) + nPoints
    anBossShots(cModeMole) = anBossShots(cModeMole) + 1
    If BitSetContains(vModesRunning, cModeSecret) = False Then
      ShowText asKnightNames(cModeMole), _
        "T=%t TOTAL=" & WilliamsFormatNum(anModeScore(cModeMole)), 2500, 4
    Else
      anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
      CheckSecret
    End If
    UpdateModeDisplay False
    PlaySound "fx-thud"
  End If

  ' Light mystery if bank completed
  If BitSetIsFull(avTrouppleHits(nCurPlayer), 3) Then
    PlaySound "fx-fanfare2", 0, 0.8
    ShowText "MYSTERY IS LIT", "HIT PINK SHOT", 2000, 3
    avTrouppleHits(nCurPlayer) = 0
    anTrouppleLit(nCurPlayer) = anTrouppleLit(nCurPlayer) + 1
    AddShotColor cShotShop, cBlinkPink
  End If
End Sub

Sub StandupTroupple1_Hit
  If False = BitSetContains(vModesRunning, cModeEnchantress) And _
  False = BitSetContains(vModesRunning, cModeJuggling) And _
  False = BitSetContains(vModesRunning, cModeSecret) Then
    If anTrouppleLit(nCurPlayer) < 1 Then
      If Not BitSetContains(avTrouppleHits(nCurPlayer), 0) Then
        avTrouppleHits(nCurPlayer) = BitSetAdd(avTrouppleHits(nCurPlayer), 0)
      Elseif Not BitSetContains(avTrouppleHits(nCurPlayer), 1) Then
        avTrouppleHits(nCurPlayer) = BitSetAdd(avTrouppleHits(nCurPlayer), 1)
      End If
    End If
  End If
  AnyTrouppleHit
  AnySwitchHit cSwitchTarget, 12
End Sub

Sub StandupTroupple2_Hit
  If False = BitSetContains(vModesRunning, cModeEnchantress) And _
  False = BitSetContains(vModesRunning, cModeJuggling) And _
  False = BitSetContains(vModesRunning, cModeSecret) Then
    If anTrouppleLit(nCurPlayer) < 1 Then
      If Not BitSetContains(avTrouppleHits(nCurPlayer), 1) Then
        avTrouppleHits(nCurPlayer) = BitSetAdd(avTrouppleHits(nCurPlayer), 1)
      Elseif Not BitSetContains(avTrouppleHits(nCurPlayer), 0) Then
        avTrouppleHits(nCurPlayer) = BitSetAdd(avTrouppleHits(nCurPlayer), 0)
      Elseif Not BitSetContains(avTrouppleHits(nCurPlayer), 2) Then
        avTrouppleHits(nCurPlayer) = BitSetAdd(avTrouppleHits(nCurPlayer), 2)
      End If
    End If
  End If
  AnyTrouppleHit
  AnySwitchHit cSwitchTarget, 13
End Sub

Sub StandupTroupple3_Hit
  If False = BitSetContains(vModesRunning, cModeEnchantress) And _
  False = BitSetContains(vModesRunning, cModeJuggling) And _
  False = BitSetContains(vModesRunning, cModeSecret) Then
    If anTrouppleLit(nCurPlayer) < 1 Then
      If Not BitSetContains(avTrouppleHits(nCurPlayer), 2) Then
        avTrouppleHits(nCurPlayer) = BitSetAdd(avTrouppleHits(nCurPlayer), 2)
      Elseif Not BitSetContains(avTrouppleHits(nCurPlayer), 1) Then
        avTrouppleHits(nCurPlayer) = BitSetAdd(avTrouppleHits(nCurPlayer), 1)
      End If
    End If
  End If
  AnyTrouppleHit
  AnySwitchHit cSwitchTarget, 14
End Sub

' *** Saucers

Sub KickerUpperHole_Hit
  bUpperHoleHit = true
End Sub

Sub KickerShop_Hit
  Dim nSkillShotValue

  If False = BitSetContains(vModesRunning, cModeJuggling) Then
    Me.TimerEnabled = 1
  End If

  If aanXTimers(nCurPlayer, cShotShop) = 0 Then
    nSkillShotValue = 5000
  Else
    nSkillShotValue = 10000
  End If

  If BIP < 2 Then bBallHeld = True
  If bUpperHoleHit Then
    ShotHit cShotUpperHole
    bUpperHoleHit = False
  Else
    SoundSaucerLock
    ShotHit cShotShop
    If nSkillShotState = cSkillShotNoneHit Then
      PlaySound "fx-fanfare1"
      If nSkillShotValue = 5000 Then
        ShowText "HIDDEN SKILL", "SHOT " & WilliamsFormatNum(5000) & " POINTS", 2000, 3
      Else
        ShowText "HIDDEN SKILL", "SHOT " & WilliamsFormatNum(10000) & " PTS", 2000, 3
      End If
      anScore(nCurPlayer) = anScore(nCurPlayer) + nSkillShotValue
      nSkillShotState = cSkillShotHiddenHit
    End If
  End If
End Sub

Sub KickerShop_UnHit
  SoundSaucerKick 1, KickerShop
  DOF 126, DOFPulse
  DOF 122, DOFPulse
  Me.TimerEnabled = 0
  bBallHeld = False
End Sub

Sub KickerShop_Timer
  Me.Kick 335, 35
  KickerShop.TimerEnabled = False
  KickerShop.TimerInterval = 500
End Sub

Sub KickerVUK_Hit()
  Dim nSkillShotValue
  Dim sSkillShotText

  If False = BitSetContains(vModesRunning, cModeJuggling) Then
    KickerVUK.TimerEnabled = True
  End If

  If bDrainingBalls Then Exit Sub

  If aanXTimers(nCurPlayer, cShotVUK) = 0 Then
    nSkillShotValue = 5000
    sSkillShotText = WilliamsFormatNum(5000) & " POINTS"
  Else
    nSkillShotValue = 10000
    sSkillShotText = WilliamsFormatNum(10000) & " PTS"
  End If
  SoundSaucerLock
  If BIP < 2 Then bBallHeld = True
  If nSkillShotState <> cSkillShotDisabled Then
    If nSkillShotState = cSkillShotHiddenHit Then
      ShowText "SUPER SKILL SHOT", sSkillShotText, 2000, 2
      PlaySound "fx-fanfare2"
    Else
      ShowText "FLIPPER SKILL", "SHOT " & sSkillShotText, 2000, 2
      PlaySound "fx-fanfare1"
    End If
    anScore(nCurPlayer) = anScore(nCurPlayer) + nSkillShotValue
    RemoveShotColor cShotVUK, cBlinkRed
    nSkillShotState = cSkillShotDisabled
  End If
  ShotHit cShotVUK
End Sub

Sub KickerVUK_Timer()
  SoundSaucerKick 1, KickerVUK
  DOF 127, DOFPulse
  DOF 122, DOFPulse
  KickerVUK.Kick  0, 60, 1.56
  bBallHeld = False
  KickerVUK.TimerEnabled = False
  KickerVUK.TimerInterval = 500
End Sub

Sub KickerLRamp_Hit
  ShotHit cShotRampL
  EndCombos
  SoundSaucerLock
  If BitSetContains(vModesRunning, cModeJuggling) Then
    ' do nothing
  ElseIf BitSetContains(avLocksLit(nCurPlayer), cLockLeft) Then
    PlaySound "fx-start", 0, 1
    anLocksMade(nCurPlayer) = BitSetAdd(anLocksMade(nCurPlayer), cLockLeft)
    ShowText "BALL " & CountLocksMade & " LOCKED", "", 2000, 4
    bPhysLockLeft = true
    LowerRampL
    UpdateLightsStandups
    RemoveShotColor cShotRampL, cBlinkGreen
    anShotIdleColor(cShotRampL) = cBlinkGreen
    If BitSetContains(anLocksMade(nCurPlayer), cLockRight) Then
      AddShotColor cShotSaucerR, cBlinkGreen
    End If
    If BIP < 2 Then
      KickerAutoPlunge.enabled = False
      BallRelease.CreateSizedBallWithMass Ballsize / 2, BallMass
      BallRelease.Kick 90, 7
      DOF 123, DOFPulse
      bBallLaunched = False
      bBallSaveUsed = True
      nSkillShotState = cSkillShotNoneHit
      AddShotColor cShotVUK, cBlinkRed
    Else
      AutoPlunge 1
    End If
  Else
    RaiseRampL
    Me.TimerInterval = 500
    Me.TimerEnabled = 1
  End If
End Sub

Sub KickerLRamp_UnHit
  SoundSaucerKick 1, KickerLRamp
  DOF 122, DOFPulse
  DOF 124, DOFPulse
  Me.TimerEnabled = 0
  bPhysLockLeft = false
  TimerLowerRampL.enabled = True
End Sub

Sub KickerLRamp_Timer
  Me.Kick -90, 8
End Sub

Sub TimerLowerRampL_Timer
  LowerRampL
  TimerLowerRampL.enabled = False
  UpdateRampL
End Sub

Sub TimerEjectBallC_Timer
  KickerRRamp.Kick -80, 12
  TimerEjectBallC.enabled = False
End Sub

Sub KickerRRamp_Hit
  ShotHit cShotRampC
  EndCombos
  SoundSaucerLock
  If BitSetContains(vModesRunning, cModeJuggling) Then
    ' do nothing
  ElseIf BitSetContains(avLocksLit(nCurPlayer), cLockRight) Then
    PlaySound "fx-start", 0, 1
    anLocksMade(nCurPlayer) = BitSetAdd(anLocksMade(nCurPlayer), cLockRight)
    ShowText "BALL " & CountLocksMade & " LOCKED", "", 2000, 4
    bPhysLockRight = true
    LowerRampR
    UpdateLightsStandups
    RemoveShotColor cShotRampC, cBlinkGreen
    anShotIdleColor(cShotRampC) = cBlinkGreen
    If BitSetContains(anLocksMade(nCurPlayer), cLockLeft) Then
      AddShotColor cShotSaucerR, cBlinkGreen
    End If
    If BIP < 2 Then
      KickerAutoPlunge.enabled = False
      BallRelease.CreateSizedBallWithMass Ballsize / 2, BallMass
      BallRelease.Kick 90, 7
      DOF 123, DOFPulse
      bBallLaunched = False
      bBallSaveUsed = True
      nSkillShotState = cSkillShotNoneHit
      AddShotColor cShotVUK, cBlinkRed
    Else
      AutoPlunge 1
    End If
  Else
    RaiseRampR
    Me.TimerInterval = 500
    Me.TimerEnabled = 1
  End If
End Sub

Sub KickerRRamp_UnHit
  SoundSaucerKick 1, KickerRRamp
  DOF 125, DOFPulse
  Me.TimerEnabled = 0
  bPhysLockRight = False
  TimerLowerRampC.enabled = True
End Sub

Sub KickerRRamp_Timer
  Me.Kick -80, 12
End Sub

Sub TimerLowerRampC_Timer
  LowerRampR
  TimerLowerRampC.enabled = False
  UpdateRampC
End Sub

Sub KickerSaucerR_Hit
  If BIP < 2 Then bBallHeld = True
  SoundSaucerLock
  If False = BitSetContains(vModesRunning, cModeJuggling) Then
    Me.TimerInterval = 500
    Me.TimerEnabled = 1
  End If
  ShotHit cShotSaucerR
End Sub

Sub KickerSaucerR_UnHit
  SoundSaucerKick 1, KickerSaucerR
  DOF 128, DOFPulse
  DOF 122, DOFPulse
  bBallHeld = False
  Me.TimerEnabled = 0
End Sub

Sub KickerSaucerR_Timer
  Dim vModes
  Dim nKickedLocks
  Dim i

  KickerSaucerR.timerEnabled = False
  KickerSaucerR.timerInterval = 500

  If bBossStarted Then
    bBossStarted = False
    If (BitSetContains(vModesRunning, cModeEnchantress) And _
    BitSetContains(avBossesLit(nCurPlayer), cModeEnchantress)) Or _
    (BitSetContains(vModesRunning, cModeSecret) And _
    BitSetContains(avBossesLit(nCurPlayer), cModeSecret)) Then
      avBossesLit(nCurPlayer) = BitSetRemove(avBossesLit(nCurPlayer), cModeEnchantress)
      avBossesLit(nCurPlayer) = BitSetRemove(avBossesLit(nCurPlayer), cModeSecret)
      nKickedLocks = 0
      If bPhysLockLeft Then
        RaiseRampL
        KickerLRamp.timerEnabled = True
        nKickedLocks = nKickedLocks + 1
      End If
      If bPhysLockRight Then
        RaiseRampR
        TimerEjectBallC.enabled = True
        nKickedLocks = nKickedLocks + 1
      End If
      LightSeq1.StopPlay
      AutoPlunge 5 - (nKickedLocks + BIP)
      BIP = 5
      LightShootAgain.state = LightStateBlinking
    ElseIf BitSetContains(vModesRunning, cModeJuggling) Then
      avBossesLit(nCurPlayer) = BitSetRemove(avBossesLit(nCurPlayer), cModeJuggling)
      nKickedLocks = 0
      If bPhysLockLeft Then
        KickerLRamp.timerEnabled = True
        nKickedLocks = nKickedLocks + 1
      End If
      If bPhysLockRight Then
        TimerEjectBallC.enabled = True
        nKickedLocks = nKickedLocks + 1
      End If
      If nKickedLocks = 0 Then
        AutoPlunge 1
        BIP = 2
      Else
        BIP = nKickedLocks + 1
      End if
      LightShootAgain.state = LightStateBlinking
      UpdateRampL
      UpdateRampC
    End If
    UpdateModeDisplay True
  End If

  If BitSetContains(anLocksMade(nCurPlayer), cLockLeft) = True and _
  BitSetContains(anLocksMade(nCurPlayer), cLockRight) = True and _
  BitSetContains(vModesRunning, cModeEnchantress) = False and _
  BitSetContains(vModesRunning, cModeJuggling) = False and _
  BitSetContains(vModesRunning, cModeSecret) = False and _
  bDrainingBalls = False Then
    vModes = BitSetAdd(0, cModeShovel)
    StartModes vModes
  Else
    KickerSaucerR.Kick 0, 40, 1.56
  End If
End Sub

Sub TimerEjectBallR_Timer
  TimerEjectBallR.enabled = False
  If bBossStarted Then
    KickerSaucerR_Timer
  Else
    KickerSaucerR.Kick 0, 40, 1.56
  End If
End Sub

' *** ramps and loops

Sub GateLoopL_Hit
  ' the game registers a loop if the right and left loop gates are Hit
  ' within 2 seconds of each other. Also, do not register a right Loop
  ' on balls launched from the plunger.
  If GateLoopL.timerEnabled = true then Exit Sub
  If (nTimeGateLoopR > 0) And (GameTime - nTimeGateLoopR < 2000) Then
    nTimeGateLoopR = -1
    ShotHit cShotLoopR
  Else
    If nSkillShotState = cSkillShotNoneHit Then
      nSkillShotState = cSkillShotHiddenMissed
    End If
    anScore(nCurPlayer) = anScore(nCurPlayer) + 30
    nTimeGateLoopL = gameTime
    GateLoopL.timerEnabled = true
  End If
End Sub

Sub GateLoopL_Timer
  Me.timerEnabled = false
End Sub

Sub GateLoopR_Hit
  If GateLoopR.timerEnabled = true then Exit Sub
  If (nTimeGateLoopL > 0) And (GameTime - nTimeGateLoopL < 2000) Then
    nTimeGateLoopL = -1
    ShotHit cShotLoopL
  Else
    GateLoopR.timerEnabled = true
    anScore(nCurPlayer) = anScore(nCurPlayer) + 30
    If Not bBallLaunched Then
      nTimeGateLoopR = -1
    Else
      nTimeGateLoopR = gameTime
    End If

    If bBallLaunched = False Then
      bBallLaunched = True
      If abPlayingLevel(nCurPlayer) = false And anLevelShotsToBeat(nCurPlayer) < 2 Then
        StartLevel(anLevelSelected(nCurPlayer))
      End If
      If bBallSaveUsed = False And BIP < 2 Then
        nBallSaveTimer = cBallSaveStart
        LightShootAgain.state = LightStateBlinking
      End If
    End If
  End If
End Sub

Sub GateLoopR_Timer
  Me.timerEnabled = false
End Sub

Sub TriggerLaunchRampEnter_Hit
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub TriggerLeftRampEnter_Hit
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub GateRampL_Hit
  If GateRampL.timerEnabled = false Then
    GateRampL.timerEnabled = true
    ShotHit cShotRampL
    WireRampOff
    WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
  end If
End Sub

Sub GateRampL_Timer
  GateRampL.timerEnabled = false
End Sub

Sub HoleRightInlane1_Hit
  WireRampOff
  If Not HoleRightInlane1.timerEnabled Then
    HoleRightInlane1.timerEnabled = True
    PlaySoundAtVol "WireRamp_Stop", Primitive33, Vol(ActiveBall) * MetalImpactSoundFactor
  End If
End Sub

Sub HoleRightInlane1_Timer
  HoleRightInlane1.timerEnabled = False
End Sub

Sub TriggerCenterRampEnter_Hit
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub GateRampC_Hit
  If GateRampC.timerEnabled = false Then
    GateRampC.timerEnabled = true
    ShotHit cShotRampC
  end If
  WireRampOff
End Sub

Sub GateRampC_Timer
  GateRampC.timerEnabled = false
End Sub

Sub TriggerCenterRampEnd_Hit
  WireRampOff
  If Not TriggerCenterRampEnd.timerEnabled Then
    TriggerCenterRampEnd.timerEnabled = True
    PlaySoundAtVol "WireRamp_Stop", Primitive28, Vol(ActiveBall) * MetalImpactSoundFactor
  End If
End Sub

Sub TriggerCenterRampEnd_Timer
  TriggerCenterRampEnd.timerEnabled = False
End Sub

Sub GateRampR_Hit
  If GateRampR.timerEnabled = false Then
    GateRampR.timerEnabled = true
    ShotHit cShotRampR
    WireRampOn True 'Play Plastic Ramp Sound
  end If
End Sub

Sub GateRampR_Timer
  GateRampR.timerEnabled = false
End Sub

Sub TriggerRightRampEnd_Hit
  WireRampOff
' PlaySoundAt "WireRamp_Stop", TriggerRightRampEnd
End Sub

'*** Bottom Lanes
Sub TriggerPlunger_Hit
  bBallAtPlunger = True
End Sub

Sub TriggerPlunger_UnHit
  bBallAtPlunger = False
End Sub

Sub LeftOutlane_Hit
  If bDrainingBalls Then Exit Sub
  If LeftOutlane.timerEnabled = false Then
    AnySwitchHit cSwitchLane, 0
    LeftOutlane.timerEnabled = true
    If nBallSaveTimer > 0 and BIP < 2 and (not (bOutlaneSave or bDrainingBalls)) Then
      OutlaneSave
    End If
    If vMultipliersLit = 0 Then
      avBottomLanes(nCurPlayer) = BitSetAdd(avBottomLanes(nCurPlayer), 0)
      UpdateLightsLanes
    End If
  end If
End Sub

Sub LeftOutlane_Timer
  LeftOutlane.timerEnabled = false
End Sub

Sub LeftInlane_Hit
  If bDrainingBalls Then Exit Sub
  If LeftInlane.timerEnabled = false Then
    AnySwitchHit cSwitchLane, 1
    LeftInlane.timerEnabled = true
    If vMultipliersLit = 0 Then
      avBottomLanes(nCurPlayer) = BitSetAdd(avBottomLanes(nCurPlayer), 1)
      UpdateLightsLanes
    End If
  end If
End Sub

Sub LeftInlane_Timer
  LeftInlane.timerEnabled = false
End Sub

Sub RightInlane1_Hit
  If bDrainingBalls Then Exit Sub
  If RightInlane1.timerEnabled = false Then
    AnySwitchHit cSwitchLane, 2
    RightInlane1.timerEnabled = true
    If vMultipliersLit = 0 Then
      avBottomLanes(nCurPlayer) = BitSetAdd(avBottomLanes(nCurPlayer), 2)
      UpdateLightsLanes
    End if
  end If
End Sub

Sub RightInlane1_Timer
  RightInlane1.timerEnabled = false
End Sub

Sub RightInlane2_Hit
  If bDrainingBalls Then Exit Sub
  If RightInlane2.timerEnabled = false Then
    AnySwitchHit cSwitchLane, 3
    RightInlane2.timerEnabled = true
    If vMultipliersLit = 0 Then
      avBottomLanes(nCurPlayer) = BitSetAdd(avBottomLanes(nCurPlayer), 3)
      UpdateLightsLanes
    end if
  end If
End Sub

Sub RightInlane2_Timer
  RightInlane2.timerEnabled = false
End Sub

Sub RightOutlane_Hit
  If bDrainingBalls Then Exit Sub
  If RightOutlane.timerEnabled = false Then
    AnySwitchHit cSwitchLane, 4
    RightOutlane.timerEnabled = true
    If nBallSaveTimer > 0 and BIP < 2 and (not (bOutlaneSave or bDrainingBalls)) Then
      OutlaneSave
    End If
    If vMultipliersLit = 0 Then
      avBottomLanes(nCurPlayer) = BitSetAdd(avBottomLanes(nCurPlayer), 4)
      UpdateLightsLanes
    end If
  end If
End Sub

Sub RightOutlane_Timer
  RightOutlane.timerEnabled = false
End Sub

Sub OutlaneSave
  AutoPlunge 1
  bOutlaneSave = True
  If nExtraBalls > 0 Then
    LightShootAgain.state = LightStateOn
  Else
    LightShootAgain.state = LightStateOff
  End If
  nBallSaveTimer = 0
  bBallSaveUsed = True
  If nSkillShotState <> cSkillShotDisabled Then
    RemoveShotColor cShotVUK, cBlinkRed
  End If
  If nSkillShotState <> cSkillShotDisabled Then
    RemoveShotColor cShotVUK, cBlinkRed
  End If
  nSkillShotState = cSkillShotDisabled
  ShowText "BALL SAVED", "PLAYER " & (nCurPlayer + 1), 2500, 2
  PlaySound "fx-swoosh"
End Sub

Sub UpdateLightsLanes
  Dim i
  If BitSetIsFull(avBottomLanes(nCurPlayer), 5) Then
    ' light multipliers
    ShowText "SHOT MULTIPLIERS", "ARE LIT", 2500, 3
    PlaySound "fx-pause"
    avBottomLanes(nCurPlayer) = 0
    LightBottomLane1.timerEnabled = True
    for i = 0 to 4
      aoBottomLaneLights(i).state = LightStateBlinking
    next
    vMultipliersLit = 0
    for i = 0 to 9
      if not BitSetContains(avXUsed(nCurPlayer), i) Then
        vMultipliersLit = BitSetAdd(vMultipliersLit, i)
        AddShotColor i, cBlinkWhite
      end If
    Next
  End If
  if LightBottomLane1.timerEnabled = false Then
    for i = 0 to 4
      if BitSetContains(avBottomLanes(nCurPlayer), i) Then
        aoBottomLaneLights(i).state = LightStateOn
      else
        aoBottomLaneLights(i).state = LightStateOff
      end if
    next
  end if
End Sub

Sub LightBottomLane1_Timer
  LightBottomLane1.timerEnabled = False
  UpdateLightsLanes
End Sub

'*** Functions to support levels and bosses

Function nCountBossesPlayed
  Dim i
  Dim nModeCount

  nModeCount = 0
  for i = 0 to 8
    If aanBossMedals(nCurPlayer, i) <> cMedalNotPlayed Then
      nModeCount = nModeCount + 1
    End If
  Next

  nCountBossesPlayed = nModeCount
End Function

Function nCountBossesLit
  Dim i
  Dim nModeCount

  nModeCount = 0
  for i = 0 to 9
    If BitSetContains(avBossesLit(nCurPlayer), i) Then
      nModeCount = nModeCount + 1
    End If
  Next

  nCountBossesLit = nModeCount
End Function

' * Calculate how many shots will be needed to finish the next level
Function nShotsToBeatLevel
  Dim nLevelsCompleted
  Dim nShots
  Dim nBossesLit

  nBossesLit = nCountBossesLit
  nLevelsCompleted = nCountBossesPlayed + nBossesLit
  Select Case nLevelsCompleted
    Case 1, 2
      nShots = 4
    Case 3, 4
      nShots = 5
    Case 5, 6
      nShots = 6
    Case 7, 8
      nShots = 7
    Case 9
      nShots = 10
    Case Else
      nShots = 3
  End Select

  ' To go from completing a level to starting a second level for a second boss
  ' to stack, requires a number of shots equal to half the shots needed To
  ' complete the level, rounded up. To go from completing a second level to
  ' starting a third level for a third boss to stack, requires a number of
  ' shots equal to 75% of the shots needed to complete the level, rounded up.
  If abPlayingLevel(nCurPlayer) = False Then
    If nBossesLit = 0 Then
      nShotsToBeatLevel = 0
    ElseIf nBossesLit = 1 Then
      nShotsToBeatLevel = (nShots + 1) \ 2
    Else
      nShotsToBeatLevel = ((3 * nShots) + 3) \ 4
    End If
  Else
    nShotsToBeatLevel = nShots
  End If

End Function

Sub StartLevel(nLevel)
  abPlayingLevel(nCurPlayer) = True
  anLevelShotsToBeat(nCurPlayer) = nShotsToBeatLevel
  anLevelShotsMade(nCurPlayer) = 0
  dLevelCountUp = 0
  anMostRecentLevel(nCurPlayer) = nLevel
  If TimerCountUpLevel.enabled = False Then
    ShowText asLevelNames(nLevel), "LEVEL STARTED", 2500, 3
  End If
  SelectMusic False
End Sub

' Returns whether level progress can be made. Level progress cannot be made:
' * When any boss battle is running
' * When the Shovel Knight multiball is running
' * When any wizard mode except Hall of Champions is running
' * When three bosses are already lit
' The following modes, in particular, allow level progress while they run:
' * Golden awards
' * Wanderer multiball
' * Hall of Champions
' The following is equivalent to doing BitSetAdd(vModesStackingWithLevels, nMode)
' for each mode that can be stacked.
Const vModesStackingWithLevels = 497664 '1111001100000000000
Function bCanProgressLevel
  Dim vBitMaskResult
  Dim bMoreBossesAvailable

  If nCountBossesPlayed > 8 Then
    ' Once the last boss is fought, the final tower level can be played
    ' (unless it is already completed)
    If Not BitSetContains(avBossesLit(nCurPlayer), cModeEnchantress)  Then
      bMoreBossesAvailable = True
    Else
      bMoreBossesAvailable = False
    End If
  ElseIf nCountBossesPlayed + nCountBossesLit > 8 Then
    ' But do not allow progess on the final tower before that
    bMoreBossesAvailable = False
  Elseif nCountBossesLit > 2 Then
    ' Only allow stacking 3 bosses at most
    bMoreBossesAvailable = False
  Else
    bMoreBossesAvailable = True
  End If

  vBitMaskResult = vModesRunning And Not vModesStackingWithLevels
  bCanProgressLevel = (vBitMaskResult = 0) And (bMoreBossesAvailable = True)
End Function

' Version of above that only checks that a forbidden mode is not running, not
' whether more levels can be started (OR THAT ANY BOSS IS LIT!)
Function bCanStartBoss
  Dim vBitMaskResult
  vBitMaskResult = vModesRunning And Not vModesStackingWithLevels
  bCanStartBoss = (vBitMaskResult = 0)
End Function

Function bCanStartJuggling
  If BIP < 2 And False = BitSetContains(avBossesLit(nCurPlayer), cModeEnchantress) _
  And False = BitSetContains(avBossesLit(nCurPlayer), cModeSecret) _
  And CountLocksMade < 2 And anComboJPSum(nCurPlayer) >= cJugglingGoal Then
    bCanStartJuggling = True
  Else
    bCanStartJuggling = False
  End If
End Function

Function bCanStartSecret
  If BIP < 2 And False = BitSetContains(avBossesLit(nCurPlayer), cModeEnchantress) _
  And False = BitSetContains(avBossesLit(nCurPlayer), cModeJuggling) _
  And CountLocksMade < 2 And avSecretProgress(nCurPlayer) = 61 Then
    bCanStartSecret = True
  Else
    bCanStartSecret = False
  End If
End Function

Sub ChangeLevelToStart
  Dim i
  Dim nModeCount
  Dim bUnplayedLevelFound
  Dim nBossLight
  Dim nLoops
  ' Make sure a ball is in play or in the shooter lane
  If nPlayers < 1 Then Exit Sub
  If bCountingBonus Or bDrainingBalls Or bEnteringHighScore Then Exit Sub
  ' Make sure there is an unplayed level to change To
  nModeCount = nCountBossesPlayed + nCountBossesLit
  If nModeCount > 8 Then
    If nCountBossesLit < 1 Then
      anLevelSelected(nCurPlayer) = cModeEnchantress
    End If
    Exit Sub
  End If
  ' Make sure the level is not already started
  If bCanProgressLevel = False Or abPlayingLevel(nCurPlayer) = True Then Exit Sub
  If anLevelSelected(nCurPlayer) = cModeNone _
  Or anLevelSelected(nCurPlayer) = cModeEnchantress Then
    Exit Sub
  End if
  ' Now change the level to start
  i = anLightOfBoss(anLevelSelected(nCurPlayer))
  bUnplayedLevelFound = False
  nLoops = 0
  While bUnplayedLevelFound = False And nLoops < 9
    nLoops = nLoops + 1
    i = i + 1
    If i > 8 Then i = 0
    nBossLight = anBossLightOrder(i)
    If aanBossMedals(nCurPlayer, nBossLight) = cMedalNotPlayed And _
    BitSetContains(avBossesLit(nCurPlayer), nBossLight) = False Then
      'update Lights, unless there is only 1 boss left
      If nModeCount < 8 Then
        aoBossLights(anLevelSelected(nCurPlayer)).state = LightStateOff
        'this will blink once LightTimer triggers
        aoBossLights(nBossLight).state = LightStateOn
      End If

      anLevelSelected(nCurPlayer) = nBossLight
      bUnplayedLevelFound = True
    End If
  Wend
End Sub

Sub ChangeWanderer
  Dim i
  Dim bUnplayedWandererFound
  Dim nLoops

  If bCountingBonus Or bDrainingBalls Or bEnteringHighScore Then Exit Sub
  ' Make sure there is an unplayed wanderer to change To
  if BitSetIsFull(avWanderersBeaten(nCurPlayer), 4) Then
    anWandererSelected(nCurPlayer) = cWandererAll
  End If
  If anWandererSelected(nCurPlayer) = cWandererAll Then Exit Sub

  ' Now change the wanderer to start
  bUnplayedWandererFound = False
  nLoops = 0
  i = anWandererSelected(nCurPlayer)
  While bUnplayedWandererFound = False And nLoops < 5
    nLoops = nLoops + 1
    i = i + 1
    If i > 3 Then i = 0
    If BitSetContains(avWanderersBeaten(nCurPlayer), i) = False Then
      'update Lights
      anWandererSelected(nCurPlayer) = i
      bUnplayedWandererFound = True
      UpdateLightsWanderer
    End If
  Wend
End Sub

' show bars for level progress
' the letter n represents a bar segment at 80% - 100%
' the letter m represents a bar segment at 60% - 79%
' the letter l represents a bar segment at 40% - 59%
' the letter k represents a bar segment at 20% - 39%
' the letter j represents a bar segment at 1% - 19%
' an empty bar segment is just a space
Function ProgressBar(dFract)
  dim dRemainder
  dim nFullSegments
  dim temp

  nFullSegments = Int(dFract * 12)
  dRemainder = (dFract * 12) - nFullSegments
  if nFullSegments < 0 Then nFullSegments = 0 : dRemainder = 0.0
  temp = String(nFullSegments, "n")
  If dRemainder > 0.8 Then
    temp = temp & "n"
  Elseif dRemainder > 0.6 Then
    temp = temp & "m"
  Elseif dRemainder > 0.4 Then
    temp = temp & "l"
  Elseif dRemainder > 0.2 Then
    temp = temp & "k"
  Elseif dRemainder > 0.0001 Then
    temp = temp & "j"
  End If

  temp = Left(temp, 12)

  ' pad out to 12 characters
  temp = temp & Pad(12 - len(temp))

  ProgressBar = temp
End Function

Sub TimerCountUpLevel_timer
  Dim dAmountToIncrement

  ' Only change the goal when the animation has finished (or no
  ' animation has been started)
  If nCountUpGoal < 2 Then
    nCountUpGoal = anLevelShotsToBeat(nCurPlayer)
    nCountUpShots = anLevelShotsMade(nCurPlayer)
    dLevelCountUp = 0
  End If

  ' If the number of shots goes down, assume the goal was
  ' reached and then reset. so play out the animation until 100%
  If anLevelShotsMade(nCurPlayer) < nCountUpShots Then
    If abPlayingLevel(nCurPlayer) = True Then
      nCountUpGoal = anLevelShotsToBeat(nCurPlayer)
      nCountUpShots = anLevelShotsMade(nCurPlayer)
      If nCountUpShots > 0 Then
        dLevelCountUp = nCountUpShots - 1
      Else
        dLevelCountUp = 0
      End If
    Else
      nCountUpShots = nCountUpGoal
    End If
  Else
    nCountUpShots = anLevelShotsMade(nCurPlayer)
  End If

  ' Count up 1 shot in 2000 ms
  dAmountToIncrement = TimerCountUpLevel.interval / 2000
  dLevelCountUp = dLevelCountUp + dAmountToIncrement
  If dLevelCountUp > nCountUpShots Then
    dLevelCountUp = nCountUpShots
  End If

  If oQueueTextTop.Size = 0 Then
    ShowText asLevelNames(anMostRecentLevel(nCurPlayer)), _
       ProgressBar(dLevelCountUp / nCountUpGoal) & "GOAL", _
       5, TimerDisplay.interval
  End If

  If dLevelCountUp = nCountUpShots Then
    TimerCountUpLevel.enabled = False
  End If
  If dLevelCountUp >= nCountUpGoal Then
    PlaySound "fx-fanfare2"
    ShowText asKnightNames(anMostRecentLevel(nCurPlayer)), _
         "BATTLE IS LIT", 2500, 3
    QueueText "HIT RIGHT SAUCER", "TO START BATTLE", 2500
    nCountUpShots = 0
    nCountUpGoal = 0
  End If

End Sub

Sub PlagueBlink_Timer
  If FlBumperFadeTarget(1) = 0 Then
    FlBumperFadeTarget(1) = 1
    FlBumperFadeTarget(2) = 1
    FlBumperFadeTarget(3) = 1
  Else
    FlBumperFadeTarget(1) = 0
    FlBumperFadeTarget(2) = 0
    FlBumperFadeTarget(3) = 0
  End If
End Sub

'*** Combos

Sub StartCombo(aShots)
  Dim i, j
  Dim nBannedShots
  Dim bWasBanned
  Dim vComboShots
  Dim temp

  vComboShots = 0
  for each i in aShots
    vComboShots = BitSetAdd(vComboShots, i)
  Next

  'when a combo is made, unlight ongoing combos
  For i = cShotRampL to cShotRampR
    If BitSetContains(vComboShots, i) Then
      bWasBanned = False

      If nComboCurrent > 4 Then
        nBannedShots = 3
      Else
        nBannedShots = Int(nComboCurrent / 2)
      End If

      temp = ""
      for j = 1 to nBannedShots
        temp = temp & anComboShots((1 + j + nComboCurrent) Mod 3) & " "
        If anComboShots((1 + j + nComboCurrent) Mod 3) = i Then
          bWasBanned = True
        End If
      Next

      ' Do not allow combos into raised ramps
      If i = cShotRampC And False = ramp3.Collidable Then
        bWasBanned = True
      End If
      If i = cShotRampL And False = ramp7.Collidable Then
        bWasBanned = True
      End If

      If False = bWasBanned Then
        anComboTimers(i) = 4000
        AddShotColor i, cBlinkYellow
      End If
    Else
      If anComboTimers(i) > 0 Then
        anComboTimers(i) = 0
      End If
      RemoveShotColor i, cBlinkYellow
    End If
  Next
End Sub

Sub EndCombos
  Dim i

  nComboCurrent = 0

  for i = 0 to 2
    anComboShots(i) = cShotNone
  Next

  For i = cShotRampL to cShotRampR ' left loop awards combo rather than building it
    If anComboTimers(i) > 0 Then
      anComboTimers(i) = 0
    End If
    RemoveShotColor i, cBlinkYellow
  Next
End Sub

Sub TimerCombo_Timer
  Dim i

  If nPlayers > 0 And False = bCountingBonus Then

    ' Do not count down combos unless a ball is in play
    If BIP < 2 And (True = bBallHeld Or False = bBallLaunched) Then
      Exit Sub
    End If

    ' Handle left Loop
    If anComboTimers(cShotLoopL) > 0 Then
      anComboTimers(cShotLoopL) = anComboTimers(cShotLoopL) - TimerCombo.Interval
      If anComboTimers(cShotLoopL) < 2000 + TimerCombo.Interval _
      And anComboTimers(cShotLoopL) > 2000 - TimerCombo.Interval Then
        RemoveShotColor cShotLoopL, cBlinkOrange
      End If
      If anComboTimers(cShotLoopL) < 1 Then
        anComboTimers(cShotLoopL) = 0
        nComboJPCurrent = 0
        If BitSetContains(vModesRunning, cModeSuperCombo) Then
          UpdateModeDisplay False
        Else
          UpdateModeDisplay True
        End If
        RemoveShotColor cShotLoopL, cBlinkOrange
      End If
    End If

    For i = cShotRampL to cShotRampR ' left loop awards combo rather than building it
      If anComboTimers(i) > 0 Then
        anComboTimers(i) = anComboTimers(i) - TimerCombo.Interval
        If anComboTimers(i) < 1 Then
          anComboTimers(i) = 0
          EndCombos
          Exit For
        End If
      End If
    Next

    if BitSetContains(vModesRunning, cModeJuggling) Then
      for each i in anJugglingShots
        if anJuggleTimers(i) > 0 Then
          anJuggleTimers(i) = anJuggleTimers(i) - TimerCombo.Interval
          If anJuggleTimers(i) < 1 Then
            anJuggleTimers(i) = 0
            Select case i
              case cShotRampL
                KickerLRamp.Kick -90, 8
              case cShotRampC
                KickerRRamp.Kick -80, 12
              case cShotShop
                KickerShop.Kick 335, 35
                DropShop.IsDropped = True
              case cShotVUK
                DOF 127, DOFPulse
                SoundSaucerKick 1, KickerVUK
                KickerVUK.Kick  0, 60, 1.56
              case cShotSaucerR
                KickerSaucerR.Kick 0, 40, 1.56
            End Select
            UpdateModeDisplay False
          end if
        end If
      Next
      UpdateLightsLocks
    end if
  End If
End Sub

' extend combos from right ramp when upper playfield exit is Hit
Sub GateUpperPF_Hit
  If anComboTimers(cShotUpperHole) > 0 Then
    anComboTimers(cShotUpperHole) = 0
    RemoveShotColor cShotUpperHole, cBlinkYellow
    StartCombo Array(cShotRampL, cShotCaptive, cShotRampC)
  End If
  WireRampOn False
End Sub

Sub JugglingLock(nShot)
  Dim i
  Dim nLockedBalls
  Dim nPoints
  Dim nMultiplier

  If aanXTimers(nCurPlayer, nShot) > 0 Then
    nMultiplier = 2
  Else
    nMultiplier = 1
  End If

  nLockedBalls = 0
  If nShot = cShotVUK Then
    anJuggleTimers(nShot) = 30000
  Else
    anJuggleTimers(nShot) = 20000
  end If

  for each i in anJugglingShots
    if anJuggleTimers(i) > 0 Then
      nLockedBalls = nLockedBalls + 1
    end If
  Next
  if nShot = cShotShop And nLockedBalls < BIP Then DropShop.IsDropped = False

  If nLockedBalls = BIP Then
    Select Case BIP
      Case 2
        nPoints = 3750
      Case 3
        nPoints = 7500
      Case 4
        nPoints = 15000
      Case 5
        nPoints = 30000
    End Select
    nPoints = nPoints * nMultiplier
    PlaySound "fx-fanfare2"
    ShowText BIP & " BALLS JUGGLED", WilliamsFormatNum(nPoints) & " POINTS", 2500, 1
    anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints

    If BIP > 4 Then
      avSecretProgress(nCurPlayer) = BitSetAdd(avSecretProgress(nCurPlayer), cSecretJuggling)
      EndModes BitSetAdd(0, cModeJuggling)
    Else
      LightShootAgain.state = LightStateBlinking
      nBallSaveTimer = cBallSaveMball
      AutoPlunge 1
      BIP = BIP + 1
      ' kick out balls
      for each i in anJugglingShots
        if anJuggleTimers(i) > 0 Then
          anJuggleTimers(i) = 200
        end If
      Next
    End If
  End If
  UpdateLightsLocks
End Sub

'*** Functions called when any switch or shot is hit

Sub AnySwitchHit(nSwitchType, nSwitchIndex)
  If bDrainingBalls Then Exit Sub

  nTimeLastSwitch = GameTime

  if nTextDuration = PlayUntilHit then
    nTextDuration = 0
    UpdateModeDisplay True
  End If

  If nSkillShotState = cSkillShotNoneHit Then
    If Not((nSwitchType = cSwitchShot) And (nSwitchIndex = cShotShop)) Then
      nSkillShotState = cSkillShotDisabled
      RemoveShotColor cShotVUK, cBlinkRed
    End If
  ElseIf (nSkillShotState = cSkillShotHiddenMissed Or _
          nSkillShotState = cSkillShotHiddenHit) Then
    If Not(((nSwitchType = cSwitchShot) And (nSwitchIndex = cShotVUK)) Or _
         ((nSwitchType = cSwitchTarget) And (nSwitchIndex = 5)) Or _
         ((nSwitchType = cSwitchTarget) And (nSwitchIndex = 12))) Then
      nSkillShotState = cSkillShotDisabled
      RemoveShotColor cShotVUK, cBlinkRed
    End If
  End If

  ' Increment incidental bonus
  Dim bUpperPFSwitch : bUpperPFSwitch = False
  Select Case nSwitchType
    Case cSwitchSpinner
      anBonus(nCurPlayer) = anBonus(nCurPlayer) + 2
    Case cSwitchBumper
      anBonus(nCurPlayer) = anBonus(nCurPlayer) + 4
    Case cSwitchLane
      anBonus(nCurPlayer) = anBonus(nCurPlayer) + 10
      anScore(nCurPlayer) = anScore(nCurPlayer) + 90
    Case cSwitchTarget
      anBonus(nCurPlayer) = anBonus(nCurPlayer) + 10
      If nSwitchIndex > 5 And nSwitchIndex < 12 Then bUpperPFSwitch = True
    Case cSwitchShot
      anBonus(nCurPlayer) = anBonus(nCurPlayer) + 30
  End Select
  If bUpperPFSwitch Then nUpperPFHits = nUpperPFHits + 1 else nUpperPFHits = 0

End Sub

Sub ShotHit(nShot)
  Dim i
  Dim nJPCount
  Dim nMultiplier
  Dim nPoints
  Dim sSoundToPlay
  Dim bBossStarted
  Dim nShotCount
  Dim vStartingModes
  Dim bCanStartChamps
  Dim nMysteryAward
  Dim oUsefulAwards
  Dim nNextComboAward
  Dim nShovelTargets
  Dim nSoundPriority : nSoundPriority = 99
  Dim nWanderer
  Dim nWandererGoal
  Dim nWandererPrevValue

  If bDrainingBalls Then Exit Sub

  If aanXTimers(nCurPlayer, nShot) > 0 Then
    nMultiplier = 2
  Else
    nMultiplier = 1
  End If

  AnySwitchHit cSwitchShot, nShot
  anScore(nCurPlayer) = anScore(nCurPlayer) + 190

  ' extra Balls
  If nShot = cShotUpperHole And anExtraBallsLit(nCurPlayer) > 0 Then
    nExtraBalls = nExtraBalls + anExtraBallsLit(nCurPlayer)
    anExtraBallsLit(nCurPlayer) = 0
    ShowText "EXTRA BALL", "EXTRA BALL", 2500, 1
    If nBallSaveTimer = 0 Then
      LightShootAgain.state = LightStateOn
    End If
    ' play jingle, unless uper jackpot is lit at the upper hole
    If False = BitSetContains(vModesRunning, cModeShovel) Or False = bShovelSuperJP Then
      InitMusic StopMusic
      sSoundToPlay = "fx-superjackpot3x"
      nSoundPriority = 1
      TimerSJackJingle.enabled = True
      LightSeq1.play SeqDownOn, 40, 1
    End If
    UpdateLightsWanderer
  End If

  ' Shovel Knight Multiball
  If BitSetContains(vModesRunning, cModeShovel) Or anGracePeriod(cModeShovel) > 0 Then

    ' Handle super jackpot
    If nShot = cShotUpperHole and bShovelSuperJP Then
      ' Adjust diverters and ramps.
      ' 1X Super = Upper diverter goes to upper playfield, both ramps down
      ' 2X Super = Upper diverter goes to right flipper, both ramps up
      ' 3X Super = Upper diverter goes to right flipper, both ramps down,
      '      shop shot is blocked
      if anKnightSJacks(nCurPlayer) mod 3 = 0 Then
        InitMusic StopMusic
        sSoundToPlay = "fx-superjackpot1x"
        nSoundPriority = 1
        LightShowSpin 5, 3
        TimerSJackJingle.enabled = True
        nSJValue = cBaseSuperJP * 1 * nMultiplier
        If nMultiplier = 1 Then
          sSJMultiplier = ""
        Else
          sSJMultiplier = "2X "
        End If
      Elseif anKnightSJacks(nCurPlayer) mod 3 = 1 Then
        InitMusic StopMusic
        nSoundPriority = 1
        sSoundToPlay = "fx-superjackpot2x"
        LightShowSpin 5, 3
        TimerSJackJingle.enabled = True
        nSJValue = cBaseSuperJP * 2 * nMultiplier
        sSJMultiplier = CStr(2 * nMultiplier) & "X "
        DropShop.isDropped = False
      Else
        InitMusic StopMusic
        sSoundToPlay = "fx-superjackpot3x"
        nSoundPriority = 1
        LightShowSpin 5, 3
        TimerSJackJingle.enabled = True
        nSJValue = cBaseSuperJP * 3 * nMultiplier
        sSJMultiplier = CStr(3 * nMultiplier) & "X "
        avSecretProgress(nCurPlayer) = _
          BitSetAdd(avSecretProgress(nCurPlayer), cSecret3xSuper)
      end If

      anKnightSJacks(nCurPlayer) = anKnightSJacks(nCurPlayer) + 1
      anScore(nCurPlayer) = anScore(nCurPlayer) + nSJValue
      If BitSetContains(vModesRunning, cModeSecret) Then
        anModeScore(cModeSecret) = anModeScore(cModeSecret) + nSJValue
        CheckSecret
      End If
      bShovelSuperJP = False
      RemoveShotColor cShotUpperHole, cBlinkPurple
      LightSJLit.state = LightStateOff
      UpdateLightsStandups

      UpdateDiverter
      UpdateRampL
      UpdateRampC

      ShowText sSJMultiplier & "SUPER JACKPOT", WilliamsFormatNum(nSJValue), 5000, 1
      If anGracePeriod(cModeShovel) < 1 Then
        avShovelLightsMask(nCurPlayer) = 0
        avShovelTargetsHit(nCurPlayer) = 0
        UpdateModeDisplay True
      End If
    End If

    ' Handle regular shovel knight multiball jackpot
    nJPCount = 0
    for i = 0 to 2
      If nShot = anShovelJackpots(i) Then nJPCount = nJPCount + 1
    Next

    If nJPCount > 0 Then
      If nSoundPriority > 1 Then
        sSoundToPlay = "fx-breakstuff"
        nSoundPriority = 2
      End If
      ' Light show for jackpot
      LightSeq1.UpdateInterval = 5
      Select Case nShot
        Case cShotLoopL, cShotRampL, cShotCaptive
          LightSeq1.play SeqDiagUpLeftOn, 40, 1
        Case cShotRampC, cShotShop
          LightSeq1.play SeqUpOn, 40, 1
        Case cShotVUK
          LightSeq1.play SeqRightOn, 40, 1
        Case cShotSaucerR, cShotLoopR, cShotRampR
          LightSeq1.play SeqDiagUpRightOn, 40, 1
      End Select
    End If
    If nJPCount = 1 And nMultiplier = 1 Then
      nPoints = (cBaseShovelJP * 1)
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      ShowText "JACKPOT", WilliamsFormatNum(cBaseShovelJP), 2500, 2
    Elseif nJPCount + nMultiplier = 3 Then
      nPoints = (cBaseShovelJP * 2)
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      ShowText "DOUBLE JACKPOT", WilliamsFormatNum(cBaseShovelJP * 2), 2500, 2
    Elseif nJPCount + nMultiplier = 4 Then
      nPoints = (cBaseShovelJP * 4)
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      ShowText "QUADRUPLE", "JACKPOT " & WilliamsFormatNum(cBaseShovelJP * 4), 2500, 2
    Elseif nJPCount + nMultiplier = 5 Then
      nPoints = (cBaseShovelJP * 8)
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      ShowText "OCTUPLE JACKPOT", WilliamsFormatNum(cBaseShovelJP * 8), 2500, 2
    End If
    If BitSetContains(vModesRunning, cModeSecret) Then
      anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
      CheckSecret
    End If

    If nShot = anShovelJackpots(0) Then
      RemoveShotColor nShot, cBlinkAmber
      anShovelJackpots(0) = MoveSKJackpot(nShot, 1)
      If anGracePeriod(cModeShovel) < 1 Then
        AddShotColor anShovelJackpots(0), cBlinkAmber
      End If
    End If
    If nShot = anShovelJackpots(1) Then
      RemoveShotColor nShot, cBlinkOrange
      anShovelJackpots(1) = MoveSKJackpot(nShot, 9)
      If anGracePeriod(cModeShovel) < 1 Then
        AddShotColor anShovelJackpots(1), cBlinkOrange
      End If
    End If
    If nShot = anShovelJackpots(2) Then
      RemoveShotColor nShot, cBlinkRed
      anShovelJackpots(2) = MoveSKJackpot(nShot, 2)
      If anGracePeriod(cModeShovel) < 1 Then
        AddShotColor anShovelJackpots(2), cBlinkRed
      End If
    End If
  End If

  ' Bosses: Spectre Knight
  If BitSetContains(vModesRunning, cModeSpectre) Or _
  BitSetContains(vModesRunning, cModeEnchantress) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    If nShot = cShotRampL Or nshot = cShotRampR Then
      nPoints = 4000
      If aanXTimers(nCurPlayer, nshot) > 0 Then
        nPoints = nPoints * 2
      End If
      If BitSetContains(vModesRunning, cModeSpectre) = False Then
        nPoints = nPoints * 0.6
      End If
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      anModeScore(cModeSpectre) = anModeScore(cModeSpectre) + nPoints
      anBossShots(cModeSpectre) = anBossShots(cModeSpectre) + 1
      If BitSetContains(vModesRunning, cModeSecret) = False Then
        ShowText asKnightNames(cModeSpectre), _
          "T=%t TOTAL=" & WilliamsFormatNum(anModeScore(cModeSpectre)), 2500, 4
      Else
        anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
        CheckSecret
      End If
      UpdateModeDisplay False
      If nSoundPriority > 3 Then
        sSoundToPlay = "fx-thud"
        nSoundPriority = 4
      End If
    End If
  End If

  ' Bosses: Tinker Knight
  If BitSetContains(vModesRunning, cModeTinker) Or _
  BitSetContains(vModesRunning, cModeEnchantress) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    If nShot = cShotCaptive Then
      nPoints = 4000
      If BitSetContains(vModesRunning, cModeTinker) = False Then
        nPoints = nPoints * 0.6
      End If
      If aanXTimers(nCurPlayer, nshot) > 0 Then
        nPoints = nPoints * 2
      End If
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      anModeScore(cModeTinker) = anModeScore(cModeTinker) + nPoints
      anBossShots(cModeTinker) = anBossShots(cModeTinker) + 1
      If BitSetContains(vModesRunning, cModeSecret) = False Then
        ShowText asKnightNames(cModeTinker), _
          "T=%t TOTAL=" & WilliamsFormatNum(anModeScore(cModeTinker)), 2500, 4
      Else
        anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
        CheckSecret
      End If
      UpdateModeDisplay False
      If nSoundPriority > 3 Then
        sSoundToPlay = "fx-thud"
        nSoundPriority = 4
      End If
    End If
  End If

  ' Bosses: Propeller Knight
  If BitSetContains(vModesRunning, cModePropeller) Or _
  BitSetContains(vModesRunning, cModeEnchantress) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    If nShot = cShotRampC Then
      nPoints = 3500
      If BitSetContains(vModesRunning, cModePropeller) = False Then
        nPoints = nPoints * 0.6
        nPoints = (nPoints \ 10) * 10 ' round down to multiple of 10
      End If
      If aanXTimers(nCurPlayer, nshot) > 0 Then
        nPoints = nPoints * 2
      End If
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      anModeScore(cModePropeller) = anModeScore(cModePropeller) + nPoints
      anBossShots(cModePropeller) = anBossShots(cModePropeller) + 1
      If BitSetContains(vModesRunning, cModeSecret) = False Then
        ShowText asKnightNames(cModePropeller), _
          "T=%t TOTAL=" & WilliamsFormatNum(anModeScore(cModePropeller)), 2500, 4
      Else
        anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
        CheckSecret
      End If
      UpdateModeDisplay False
      If nSoundPriority > 3 Then
        sSoundToPlay = "fx-thud"
        nSoundPriority = 4
      End If
    End If
  End If

  ' Bosses: treasure Knight
  If BitSetContains(vModesRunning, cModeTreasure) Or _
  BitSetContains(vModesRunning, cModeEnchantress) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    If nShot = cShotSaucerR Or nshot = cShotVUK Or nShot = cShotShop Then
      If not BitSetContains(vTreasureHits, nShot) Then
        vTreasureHits = BitSetAdd(vTreasureHits, nShot)
        RemoveShotColor nshot, cBlinkBlue
        nPoints = 3000 * (1 + (anBossShots(cModeTreasure) mod 3))
        anBossShots(cModeTreasure) = anBossShots(cModeTreasure) + 1
        If aanXTimers(nCurPlayer, nshot) > 0 Then
          nPoints = nPoints * 2
        End If
        If BitSetContains(vModesRunning, cModeTreasure) = False Then
          nPoints = nPoints * 0.6
          nPoints = (nPoints \ 10) * 10 ' round down to multiple of 10
        End If
        anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
        anModeScore(cModeTreasure) = anModeScore(cModeTreasure) + nPoints
        If BitSetContains(vModesRunning, cModeSecret) = False Then
          ShowText asKnightNames(cModeTreasure), _
            "T=%t TOTAL=" & WilliamsFormatNum(anModeScore(cModeTreasure)), 2500, 4
        Else
          anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
          CheckSecret
        End If
        UpdateModeDisplay False
        If nSoundPriority > 3 Then
          sSoundToPlay = "fx-thud"
          nSoundPriority = 4
        End If
        If anBossShots(cModeTreasure) = 3 Then
          If BitSetContains(vModesRunning, cModeTreasure) Then
            ' time out mode instantly as no more targets can be hit
            anModeTimeLeft(cModeTreasure) = 1
          Else
            vTreasureHits = 0
          End If
        End If
      End If
    End If
  End If

  ' juggling
  If BitSetContains(vModesRunning, cModeJuggling) Then
    for each i in anJugglingShots
      if nShot = i Then
        JugglingLock nShot
        exit For
      end If
    Next
  End If

  ' Boss / Hall of Champs / juggling Start
  bBossStarted = False
  nShotCount = 0
  For i = 0 to 9
    If aanXTimers(nCurPlayer, i) > 0 _
    Or true = BitSetContains(avXUsed(nCurPlayer), i) Then
      nShotCount = nShotCount + 1
    End If
  Next
  if 10 = nShotCount _
  And BitSetContains(vModesRunning, cModeJuggling) = False _
  And BitSetContains(vModesRunning, cModeChampions) = False _
  And BitSetContains(vModesRunning, cModeSecret) = False Then
    bCanStartChamps = True
  Else
    bCanStartChamps = False
  End If

  If nShot = cShotSaucerR And true = bCanStartSecret Then
    StartModes BitSetAdd(0, cModeSecret)
    bBossStarted = True
  Elseif nShot = cShotSaucerR And ((bCanStartBoss = True And nCountBossesLit > 0) _
  Or (bCanStartJuggling = True) Or (bCanStartChamps = True)) Then
    vStartingModes = avBossesLit(nCurPlayer)
    If bCanStartChamps Then
      vStartingModes = BitSetAdd(vStartingModes, cModeChampions)
    End If
    If bCanStartJuggling Then
      vStartingModes = BitSetAdd(vStartingModes, cModeJuggling)
    End If
    StartModes vStartingModes
    bBossStarted = True
  End If

  ' wanderers
  If BitSetContains(vModesRunning, cModeWanderer) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    nWanderer = anWandererSelected(nCurPlayer)
    If nWanderer = cWandererAll Then
      nWandererGoal = 40000
    Else
      nWandererGoal = 10000
    End If
    nWandererPrevValue = aanWandererJPs(nCurPlayer, nWanderer)

    ' Rieze
    If nWanderer = cWandererReize Or nWanderer = cWandererAll Then
      If nShot = cShotSaucerR Or nShot = cShotShop _
      Or nShot = cShotUpperHole Then
        If nSoundPriority > 7 Then
          sSoundToPlay = "fx-crash"
          nSoundPriority = 8
        End If
        nPoints = 750 * nMultiplier
        anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
        If BitSetContains(vModesRunning, cModeSecret) Then
          anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
          CheckSecret
        End If
        aanWandererJPs(nCurPlayer, nWanderer) = _
          aanWandererJPs(nCurPlayer, nWanderer) + (2500 * nMultiplier)
      End If
    End If
    ' baz
    If nWanderer = cWandererBaz Or nWanderer = cWandererAll Then
      If nShot = cShotRampL Or nShot = cShotRampC _
      Or nShot = cShotRampR Then
        If nSoundPriority > 7 Then
          sSoundToPlay = "fx-crash"
          nSoundPriority = 8
        End If
        nPoints = 750 * nMultiplier
        anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
        If BitSetContains(vModesRunning, cModeSecret) Then
          anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
          CheckSecret
        End If
        aanWandererJPs(nCurPlayer, nWanderer) = _
          aanWandererJPs(nCurPlayer, nWanderer) + (2500 * nMultiplier)
      End If
    End If
    ' phantom striker
    If nWanderer = cWandererPhantom Or nWanderer = cWandererAll Then
      If nShot = cShotLoopL Or nShot = cShotLoopR Then
        If nSoundPriority > 7 Then
          sSoundToPlay = "fx-crash"
          nSoundPriority = 8
        End If
        nPoints = 750 * nMultiplier
        anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
        If BitSetContains(vModesRunning, cModeSecret) Then
          anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
          CheckSecret
        End If
        aanWandererJPs(nCurPlayer, nWanderer) = _
          aanWandererJPs(nCurPlayer, nWanderer) + (2500 * nMultiplier)
      End If
    End If
    ' light jackpot
    If aanWandererJPs(nCurPlayer, nWanderer) >= nWandererGoal _
    And nWandererPrevValue < nWandererGoal Then
      ShowText asWandererNames(nWanderer), "JACKPOT IS LIT", 2500, 3
      If nSoundPriority > 2 Then
        sSoundToPlay = "fx-fanfare2"
        nSoundPriority = 3
      End If
      UpdateModeDisplay True
      UpdateLightsWanderer
    End If

    ' award jackpot
    If nShot = cShotVUK And aanWandererJPs(nCurPlayer, nWanderer) >= nWandererGoal Then
      nPoints = aanWandererJPs(nCurPlayer, nWanderer) * nMultiplier
      If BitSetContains(vModesRunning, cModeSecret) Then
        nPoints = nPoints / 2
        anModeScore(cModeSecret) = anModeScore(cModeSecret) + nPoints
        CheckSecret
      End If
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      ShowText asWandererNames(nWanderer), "JACKPOT " & nPoints, 2500, 2
      If nSoundPriority > 1 Then
        sSoundToPlay = "fx-breakstuff"
        nSoundPriority = 2
      End If
      aanWandererJPs(nCurPlayer, nWanderer) = 0
      UpdateLightsWanderer
      UpdateModeDisplay True
      If Not BitSetContains(vModesRunning, cModeSecret) Then
        If nWanderer = cWandererAll Then
          avSecretProgress(nCurPlayer) = BitSetAdd(avSecretProgress(nCurPlayer), cSecretWanderer)
          avWanderersBeaten(nCurPlayer) = 0
        Else
          avWanderersBeaten(nCurPlayer) = BitSetAdd(avWanderersBeaten(nCurPlayer), nWanderer)
        end if
      End If
    End If
  End If

  ' wanderer Start
  if nShot = cShotUpperHole _
  And False = BitSetContains(vModesRunning, cModeSecret) _
  And False = BitSetContains(vModesRunning, cModeJuggling) _
  And False = BitSetContains(vModesRunning, cModeEnchantress) _
  And False = BitSetContains(vModesRunning, cModeWanderer) Then
    If anWandererHits(nCurPlayer) > 0 Then
      ' Wanderer countdown
      anWandererHits(nCurPlayer) = anWandererHits(nCurPlayer) - 20
      If anWandererHits(nCurPlayer) < 0 then anWandererHits(nCurPlayer) = 0
      ShowText anWandererHits(nCurPlayer) & " POPS / SPINS", "TO LITE WANDERER", 1500, 9
      If anWandererHits(nCurPlayer) < 1 Then
        PlaySound "fx-fanfare1"
        ShowText "WANDERER IS LIT", "SHOOT UPPER HOLE", 2500, 3
        UpdateLightsWanderer
      End If
    Else
      StartModes BitSetAdd(0, cModeWanderer)
    End If
  End If

  ' Virtual ball locks
  If nShot = cShotRampL And bPhysLockLeft = True and _
  BitSetContains(avLocksLit(nCurPlayer), cLockLeft) = True and _
  BitSetContains(anLocksMade(nCurPlayer), cLockLeft) = False Then
    If nSoundPriority > 3 Then
      sSoundToPlay = "fx-start"
      nSoundPriority = 4
    End If
    VirtualLock nShot
  End If
  If nShot = cShotRampC And bPhysLockRight = True and _
  BitSetContains(avLocksLit(nCurPlayer), cLockRight) = True and _
  BitSetContains(anLocksMade(nCurPlayer), cLockRight) = False Then
    If nSoundPriority > 3 Then
      sSoundToPlay = "fx-start"
      nSoundPriority = 4
    End If
    VirtualLock nShot
  End If

  ' Combos
  If False = BitSetContains(vModesRunning, cModeEnchantress) And _
  False = BitSetContains(vModesRunning, cModeShovel) And _
  False = BitSetContains(vModesRunning, cModeSecret) And _
  False = BitSetContains(vModesRunning, cModeJuggling) Then
    If anComboTimers(nShot) > 0 And nShot <> cShotLoopL Then
      If nSoundPriority > 4 Then
        sSoundToPlay = "fx-gold"
        nSoundPriority = 5
      End If
      anComboCount(nCurPlayer) = anComboCount(nCurPlayer) + 1
      nComboCurrent = nComboCurrent + 1
      anComboTimers(cShotLoopL) = 12000
      AddShotColor cShotLoopL, cBlinkOrange

      nPoints = (aPowersOfTwo(nComboCurrent - 1)) * 300
      If nComboCurrent > 4 Then nPoints = 300
      nPoints = nPoints * nMultiplier
      if anModeTimeLeft(cModeSuperCombo) > 0 Then nPoints = nPoints * 3
      anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
      anComboJPSum(nCurPlayer) = anComboJPSum(nCurPlayer) + nPoints

      ' Combo award at even multiples of 4 combos
      If 0 = (anComboCount(nCurPlayer) Mod 4) Then
        Select Case ((anComboCount(nCurPlayer) \ 4) Mod 3)
          Case 0 '.12, 24. 36, 48...
            StartModes BitSetAdd(0, cModeSuperCombo)
          Case 1 ' 4, 16, 28, 40...
            nPoints = 5000 + ((anComboCount(nCurPlayer) \ 12) * 2500)
            nPoints = nPoints * nMultiplier
            anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
            anComboJPSum(nCurPlayer) = anComboJPSum(nCurPlayer) + nPoints
            ShowText anComboCount(nCurPlayer) & " COMBOS", _
              WilliamsFormatNum(nPoints) & " POINTS", 1500, 4
            If anComboJPSum(nCurPlayer) < cJugglingGoal Then
              QueueText "COMBO SUM=" & WilliamsFormatNum(anComboJPSum(nCurPlayer)), _
                    "JUGGLE AT " & WilliamsFormatNum(cJugglingGoal), 1500
            Else
              QueueText "COMBO SUM=" & WilliamsFormatNum(anComboJPSum(nCurPlayer)), _
                    "JUGGLING IS LIT", 1500
              UpdateLightsBoss
              UpdateModeDisplay True
            End If
          Case 2 ' 8, 20, 32, 44
            anTrouppleLit(nCurPlayer) = anTrouppleLit(nCurPlayer) + 1
            UpdateLightsBoss
            ShowText anComboCount(nCurPlayer) & " COMBOS", _
              "MYSTERY LIT", 1500, 4
        End Select
      Else
        nNextComboAward = ((anComboCount(nCurPlayer) + 3) \ 4) * 4
        Select Case ((anComboCount(nCurPlayer) \ 4) Mod 3)
          Case 0 ' 4, 16, 28, 40...
            nPoints = 5000 + ((anComboCount(nCurPlayer) \ 12) * 2500)
            ShowText anComboCount(nCurPlayer) & " COMBOS", _
              WilliamsFormatNum(nPoints) & " AT " & nNextComboAward, 1500, 4
          Case 1 ' 4, 16, 28, 40...
            ShowText anComboCount(nCurPlayer) & " COMBOS", _
              "MYSTERY AT " & nNextComboAward, 1500, 4
          Case 2 ' 8, 20, 32, 44
            ShowText anComboCount(nCurPlayer) & " COMBOS", _
              "SUPERCOMBO AT " & nNextComboAward, 1500, 4
        End Select
      End If

      ' Super combo
      if anModeTimeLeft(cModeSuperCombo) > 0 Then
        nComboJPCurrent = (aPowersOfTwo(nComboCurrent - 1)) * 3600
        If nComboJPCurrent > 28800 Then
          nComboJPCurrent = 28800
        End If
        UpdateModeDisplay False
      Else
      nComboJPCurrent = (aPowersOfTwo(nComboCurrent - 1)) * 1200
        If nComboJPCurrent > 9600 Then
          nComboJPCurrent = 9600
        End If
        UpdateModeDisplay True
      end if
    End If

    ' light combo shots
    anComboShots(nComboCurrent mod 3) = nShot
    Select Case nShot
      Case cShotLoopL
        If anComboTimers(nShot) > 1 Then
          nPoints = nComboJPCurrent * nMultiplier
          anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
          anComboJPSum(nCurPlayer) = anComboJPSum(nCurPlayer) + nPoints
          ShowText "COMBO AWARD", WilliamsFormatNum(nPoints), 2000, 3
          If anComboJPSum(nCurPlayer) < cJugglingGoal Then
            QueueText "COMBO SUM=" & WilliamsFormatNum(anComboJPSum(nCurPlayer)), _
                  "JUGGLE AT " & WilliamsFormatNum(cJugglingGoal), 1500
          Else
            QueueText "COMBO SUM=" & WilliamsFormatNum(anComboJPSum(nCurPlayer)), _
                  "JUGGLING IS LIT", 1500
          End If
          nComboJPCurrent = 0
          anComboTimers(nShot) = 0
        ' light juggling
          UpdateLightsBoss
          UpdateModeDisplay True
          RemoveShotColor nShot, cBlinkOrange
        End If
        EndCombos
      Case cShotSaucerR
        EndCombos
      Case cShotRampL
        StartCombo Array(cShotShop, cShotRampC, cShotLoopR, cShotRampR)
      Case cShotCaptive
        StartCombo Array(cShotShop, cShotRampL, cShotLoopR, cShotRampR, _
                 cShotRampC)
      Case cShotRampR
        StartCombo Array(cShotRampL, cShotCaptive, cShotRampC)
      Case cShotRampC, cShotVUK
        StartCombo Array(cShotUpperHole, cShotRampL, cShotCaptive)
      Case cShotVUK
        StartCombo Array(cShotUpperHole, cShotRampL, cShotCaptive, cShotRampC)
      Case cShotShop, cShotUpperHole, cShotLoopR
        StartCombo Array(cShotShop, cShotRampC, cShotVUK, cShotLoopR, cShotRampR)
    End Select
  End If

  ' Level Progress
  Dim nShotsLeft
  If bCanProgressLevel And Not bBossStarted Then
    If nSoundPriority > 5 Then
      sSoundToPlay = "fx-progress"
      nSoundPriority = 6
    End If
    anLevelShotsMade(nCurPlayer) = anLevelShotsMade(nCurPlayer) + 1
    ' Advance through levels twice as much if a multiplier is active
    If aanXTimers(nCurPlayer, nShot) > 0 And _
    anLevelShotsMade(nCurPlayer) < anLevelShotsToBeat(nCurPlayer) Then
      anLevelShotsMade(nCurPlayer) = anLevelShotsMade(nCurPlayer) + 1
    End If
    ' Animate level progress only if we are playing a level
    If abPlayingLevel(nCurPlayer) Then TimerCountUpLevel.enabled = True

    If anLevelShotsMade(nCurPlayer) >= anLevelShotsToBeat(nCurPlayer) Then
      If abPlayingLevel(nCurPlayer) Then
        anMostRecentBoss(nCurPlayer) = anMostRecentLevel(nCurPlayer)
        avBossesLit(nCurPlayer) = BitSetAdd(avBossesLit(nCurPlayer), anMostRecentLevel(nCurPlayer))
        abPlayingLevel(nCurPlayer) = False
        anLevelShotsMade(nCurPlayer) = 0
        anLevelShotsToBeat(nCurPlayer) = nShotsToBeatLevel
        AddShotColor cShotSaucerR, cBlinkBlue
        ChangeLevelToStart
      Else
        StartLevel(anLevelSelected(nCurPlayer))
      End If
    Elseif Not abPlayingLevel(nCurPlayer) Then
      nShotsLeft = anLevelShotsToBeat(nCurPlayer) - anLevelShotsMade(nCurPlayer)
      If nShotsLeft = 1 Then
        ShowText "ONE MORE SHOT", "TO START LEVEL", 2000, 5
      Else
        ShowText nShotsLeft & " MORE SHOTS", "TO START LEVEL", 2000, 5
      End If
    End If
  End If

  ' Multipliers. Note that the shot that activates a multiplier is NOT affected
  ' by the multiplier - the multiplier is only active AFTER the shot is made.
  If BitSetContains(vMultipliersLit, nShot) Then
    If nSoundPriority > 2 Then
      sSoundToPlay = "fx-pause"
      nSoundPriority = 3
    End If
    ' Hall of Champions mode
    If BitSetContains(vModesRunning, cModeChampions) Then
      ' Check that we did not just start the mode - the first saucer
      ' shot to start the mode should not count
      If bCanStartChamps = False Then
        ' turn off white blinks for shot made
        vMultipliersLit = BitSetRemove(vMultipliersLit, nShot)
        RemoveShotColor nShot, cBlinkWhite
        ' Special case for juggling / champions stack. Because the
        ' upper hole is not accessible during juggling, center lane
        ' shots also count as upper hole shots
        If BitSetContains(vModesRunning, cModeJuggling) _
        And cShotShop = nShot Then
          vMultipliersLit = BitSetRemove(vMultipliersLit, cShotUpperHole)
          RemoveShotColor cShotUpperHole, cBlinkWhite
        End If

        ' Count remaining shots
        nShotCount = 0
        For i = 0 to 9
          If BitSetContains(vMultipliersLit, i) Then
            nShotCount = nShotCount + 1
          End If
        Next

        ' award points and update mode total value
        nPoints = 2000 * (10 - nShotCount)
        anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
        anModeScore(cModeChampions) = anModeScore(cModeChampions) + nPoints
        ShowText "CHAMPS " & WilliamsFormatNum(nPoints), "T=%t SHOTS=%m", 1000, 2
        UpdateModeDisplay False

        ' reset timers
        anModeTimeLeft(cModeChampions) = 22000
        for i = 0 to 9
          aanXTimers(nCurPlayer, i) = 60000
        next

        ' check if all shots have been made
        If nShotCount = 0 Then
          If BitSetContains(vModesRunning, cModeSecret) Then
            ' Start over again!
            vMultipliersLit = 0
            for i = 0 to 9
              vMultipliersLit = BitSetAdd(vMultipliersLit, i)
              AddShotColor i, cBlinkWhite
              aanXTimers(nCurPlayer, i) = 60000
            next
            ' start multipliers, update lights
            UpdateLightsMultipliers
          Else
            avSecretProgress(nCurPlayer) = _
              BitSetAdd(avSecretProgress(nCurPlayer), cSecretChampions)
            anModeTimeLeft(cModeChampions) = 1
          End If
        End If
      End If
    Else
      ShowText asShotDescriptions(nShot), "2X FOR 60 SEC", 2000, 3
      ' It is possible to reset the timer of a running multiplier
      aanXTimers(nCurPlayer, nShot) = 60000
      ' turn off white blinks for shots
      for i = 0 to 9
        if BitSetContains(vMultipliersLit, i) Then
          RemoveShotColor i, cBlinkWhite
        end if
      next
      ' Check if Hall of Champions is lit
      nShotCount = 0
      For i = 0 to 9
        If aanXTimers(nCurPlayer, i) > 0 _
        Or true = BitSetContains(avXUsed(nCurPlayer), i) Then
          nShotCount = nShotCount + 1
        End If
      Next
      if 10 = nShotCount _
      And BitSetContains(vModesRunning, cModeChampions) = False _
      And BitSetContains(vModesRunning, cModeSecret) = False Then
        AddShotColor cShotSaucerR, cBlinkWhite
      End If

      UpdateLightsMultipliers
      vMultipliersLit = 0
    End If
  End If

  ' Mystery
  if nShot = cShotShop And anTrouppleLit(nCurPlayer) > 0 And _
  False = BitSetContains(vModesRunning, cModeEnchantress) And _
  False = BitSetContains(vModesRunning, cModeJuggling) And _
  False = BitSetContains(vModesRunning, cModeSecret)  Then
    ' extend kickout Timer
    KickerShop.TimerEnabled = False
    KickerShop.TimerInterval = 0
    KickerShop.TimerInterval = 2000
    KickerShop.TimerEnabled = True

    ' award mystery
    anTrouppleAwards(nCurPlayer) = anTrouppleAwards(nCurPlayer) + 1

    If nSoundPriority > 2 Then
      sSoundToPlay = "fx-mystery"
      nSoundPriority = 3
    End If

    nMysteryAward = 0
    ' If we are one target away from lighting a lock, always light lock
    nShovelTargets = 0
    for i = 0 to 11
      if False = BitSetContains(vModesRunning, cModeShovel) _
      And BitSetContains(avShovelTargetsHit(nCurPlayer), i) Then
        nShovelTargets = nShovelTargets + 1
      end If
    next
    If nShovelTargets = 11 And avLocksLit(nCurPlayer) <> cLockBoth Then
      nMysteryAward = cMysteryLightLock
    End If

    ' If we are in shovel knight multiball, add a ball unless that Has
    ' already been awarded
    If True = BitSetContains(vModesRunning, cModeShovel) And _
    False = bAddABallUsed Then
      nMysteryAward = cMysteryAddABall
    End If

    ' Guarantee an extra ball on the fourth mystery award
    If anTrouppleAwards(nCurPlayer) = 4 And False = abTrouppleXBall(nCurPlayer) Then
      nMysteryAward = cMysteryXBall
    End If

    ' Add all awards that can be used to a Queue
    Set oUsefulAwards = new Queue
    ' Points can always be awarded
    oUsefulAwards.Enqueue cMystery5k
    oUsefulAwards.Enqueue cMystery10k
    ' Bonus is maxed at 10x
    If nBonusX < 10 Then oUsefulAwards.Enqueue cMysteryBonusX
    ' bonus can be held if it is not already
    If False = abBonusHeld(nCurPlayer) Then oUsefulAwards.Enqueue cMysteryBonusHeld
    ' Only one extra ball from mystery per game
    If False = abTrouppleXBall(nCurPlayer) Then oUsefulAwards.Enqueue cMysteryXBall
    ' See if lighting a lock is possible
    If False = BitSetContains(vModesRunning, cModeEnchantress) And _
    False = BitSetContains(vModesRunning, cModeShovel) And _
    False = BitSetContains(vModesRunning, cModeSecret) And _
    False = BitSetContains(vModesRunning, cModeJuggling) And _
    avLocksLit(nCurPlayer) <> cLockBoth Then
      oUsefulAwards.Enqueue cMysteryLightLock
    End If
    ' If we are playing a level, we can light a boss
    if true = abPlayingLevel(nCurPlayer) And true = bCanProgressLevel Then
      oUsefulAwards.Enqueue cMysteryLightBoss
    End If
    ' If we are plaing a timed mode, we can add more time
    for each i in anTimedModes
      if BitSetContains(vModesRunning, i) and anModeTimeLeft(i) > 0 Then
        If i <> cModeEnchantress Then
          oUsefulAwards.Enqueue cMysteryMoreTime
          exit For
        End If
      end If
    Next
    ' Ball saver is not useful in enchantress mode
    If False = BitSetContains(vModesRunning, cModeEnchantress) Then
      oUsefulAwards.Enqueue cMysteryBallSaver
    End If
    ' Shot Multipliers can be lit if they are not already
    If vMultipliersLit = 0 Then oUsefulAwards.Enqueue cMysteryShotX

    ' If we didn't force an award, pick a random award from the queue
    If nMysteryAward = 0 Then
      for i = 0 to Int(Rnd * oUsefulAwards.Size)
        nMysteryAward = oUsefulAwards.Dequeue
      Next
    End If

    ' Give out the chosen award
    Select Case nMysteryAward
      ' light lock
      Case cMystery5k
        nPoints = 5000
        If aanXTimers(nCurPlayer, cShotShop) > 0 Then nPoints = nPoints * 2
        anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
        ShowText "MYSTERY AWARD", WilliamsFormatNum(nPoints) & " POINTS", 2000, 2
      Case cMystery10k
        nPoints = 10000
        If aanXTimers(nCurPlayer, cShotShop) > 0 Then nPoints = nPoints * 2
        anScore(nCurPlayer) = anScore(nCurPlayer) + nPoints
        ShowText "MYSTERY AWARD", WilliamsFormatNum(nPoints) & " POINTS", 2000, 2
      Case cMysteryBonusX
        nBonusX = nBonusX + 1
        ShowText "MYSTERY AWARD", "BONUS NOW AT " & nBonusX & "X", 2000, 2
      Case cMysteryLightLock
        ' Set shovel knight targets as completed
        avShovelTargetsHit(nCurPlayer) = cAllKnightTargets or cAllShovelTargets
        CheckSKStandups
      Case cMysteryShotX
        avBottomLanes(nCurPlayer) = 31
        UpdateLightsLanes
        ShowText "MYSTERY AWARD", "MULTIPLIERS LIT", 2000, 2
      Case cMysteryBonusHeld
        abBonusHeld(nCurPlayer) = True
        ShowText "MYSTERY AWARD", "BONUS HELD", 2000, 2
      Case cMysteryXBall
        abTrouppleXBall(nCurPlayer) = True
        anExtraBallsLit(nCurPlayer) = anExtraBallsLit(nCurPlayer) + 1
        PlaySound "fx-savepoint"
        ShowText "MYSTERY AWARD", "LIGHT EXTRA BALL", 2000, 2
        QueueText "SHOOT UPPER HOLE", "FOR EXTRA BALL", 2000
        UpdateLightsWanderer
      Case cMysteryAddABall
        ShowText "MYSTERY AWARD", "ADD-A-BALL", 2000, 2
        bAddABallUsed = True
        BIP = BIP + 1
        AutoPlunge 1
        nBallSaveTimer = nBallSaveTimer + cBallSaveStart
        LightShootAgain.state = LightStateBlinking
      Case cMysteryLightBoss
        anMostRecentBoss(nCurPlayer) = anMostRecentLevel(nCurPlayer)
        avBossesLit(nCurPlayer) = BitSetAdd(avBossesLit(nCurPlayer), anMostRecentLevel(nCurPlayer))
        abPlayingLevel(nCurPlayer) = False
        anLevelShotsMade(nCurPlayer) = 0
        anLevelShotsToBeat(nCurPlayer) = nShotsToBeatLevel
        AddShotColor cShotSaucerR, cBlinkBlue
        ChangeLevelToStart
        ShowText "MYSTERY AWARD", "BOSS LIT", 2000, 2
      Case cMysteryMoreTime
        for each i in anTimedModes
          if BitSetContains(vModesRunning, i) and anModeTimeLeft(i) > 0 Then
            If i <> cModeEnchantress Then
              anModeTimeLeft(i) = anModeTimeLeft(i) + 8000
            End If
          end If
        Next
        ShowText "MYSTERY AWARD", "MORE TIME ADDED", 2000, 2
      Case cMysteryBallSaver
        nBallSaveTimer = nBallSaveTimer + cBallSaveStart
        LightShootAgain.state = LightStateBlinking
        ShowText "MYSTERY AWARD", "BALL SAVER", 2000, 2
      Case Else
        ShowText "MYSTERY AWARD", "IS BROKEN " & nMysteryAward _
        & "/" & oUsefulAwards.Size, 2000, 3
    End Select

    anTrouppleLit(nCurPlayer) = anTrouppleLit(nCurPlayer) - 1
    If anTrouppleLit(nCurPlayer) < 1 then RemoveShotColor cShotShop, cBlinkPink
  End If

  ' Play sound
  If sSoundToPlay = "fx-mystery" Then
    InitMusic StopMusic
    TimerMysteryJingle.enabled = True
  End If
  PlaySound sSoundToPlay
End Sub

Function MoveSKJackpot(nShot, nSteps)
  Dim nMovedShot
  nMovedShot = nShot + nSteps
  ' The Upper hole is skipped
  If nMovedShot = cShotUpperHole And nSteps > 0 Then nMovedShot = nMovedShot + 1
  If nMovedShot = cShotUpperHole And nSteps < 0 Then nMovedShot = nMovedShot - 1
  ' Check if we moved past the first or last shot
  MoveSKJackpot = nMovedShot Mod 10
End Function

Sub ReleaseLockL_Timer
  LowerRampL
  me.Enabled = False
End Sub

Sub ReleaseLockR_Timer
  LowerRampR
  me.Enabled = False
End Sub

Sub UpdateLightsStandups
  Dim i, j, nLightCount
  Dim bDropsCleared
  ' SHOVEL KNIGHT Standups

  For i = 0 to 11
    If i > 5 And (BitSetContains(vModesRunning, cModeBlack) Or _
    BitSetContains(vModesRunning, cModeEnchantress) Or _
    BitSetContains(vModesRunning, cModeSecret)) Then
      ' If any mode is running that uses the KNIGHT lights, let
      ' TimerKnightCycle handle them
      Exit For
    End If

    If i < 6 and (BitSetContains(vModesRunning, cModePolar) Or _
    BitSetContains(vModesRunning, cModeEnchantress) Or _
    BitSetContains(vModesRunning, cModeSecret)) Then
      ' Handle Polar Knight Mode
      aoShovelLights(i).state = LightStateOff

      bDropsCleared = False
      If StandupS.isDropped And StandupH1.isDropped And StandupO.isDropped Then
        bDropsCleared = True
      End If

      If i = 2 Then
        If bDropsCleared Or Not StandupS.isDropped Then
          Light_S.state = LightStateBlinking
        End If
        If bDropsCleared Or Not StandupH1.isDropped Then
          LightH1.state = LightStateBlinking
        End If
        If bDropsCleared Or Not StandupO.isDropped Then
          LightO.state = LightStateBlinking
        End If
      ElseIf i > 2 Then
        aoShovelLights(i).state = LightStateBlinking
      End If
    Else
      If BitSetContains(avShovelLightsMask(nCurPlayer), i) Then
        aoShovelLights(i).state = LightStateOff
      ElseIf BitSetContains(avShovelTargetsHit(nCurPlayer), i) Then
        aoShovelLights(i).state = LightStateOn
      Else
        aoShovelLights(i).state = LightStateOff 'sync blinks
        aoShovelLights(i).state = LightStateBlinking
      End If
    end If
  Next

  ' Black Knight wanderer mode
  If BitSetContains(vModesRunning, cModeWanderer) _
  And (anWandererSelected(nCurPlayer) = cWandererBlack _
    Or anWandererSelected(nCurPlayer) = cWandererAll) Then
    for i = 0 to 1
      nLightCount = 0
      for j = 0 to 5
        If aoShovelLights(i * 6 + j).state = LightStateOff Then
          nLightCount = nLightCount + 1
        End If
      Next
      if nLightCount = 6 Then
        For j = 0 to 5
          aoShovelLights(i * 6 + j).state = LightStateOff 'sync blinks
          aoShovelLights(i * 6 + j).state = LightStateBlinking
        Next
      End If
    Next
    Exit Sub
  End If

  If BitSetContains(vModesRunning, cModeJuggling) Then
    For i = 0 to 11
      aoShovelLights(i).state = LightStateOff
    Next
  End If

End Sub

' Cycle the lit lights in black knight boss battle
Sub TimerKnightCycle_Timer
  Dim i
  Dim nTarget
  nBlackKnightCycle = (nBlackKnightCycle + 1) Mod 6

  for i = 6 to 11
    aoShovelLights(i).state = LightStateOff
  Next

  If anBossShots(cModeBlack) < 5 Then
    for i = 0 to (4 - anBossShots(cModeBlack))
      nTarget = 6 + ((nBlackKnightCycle + i) mod 6)
      aoShovelLights(nTarget).state = LightStateBlinking
    Next
  End If
End Sub

Sub CheckSKStandups
  Dim bShovelCompleted
  Dim bKnightCompleted
  Dim vLocksToLight
  Dim vMaskedBits

  bShovelCompleted = False
  bKnightCompleted = False
  vLocksToLight = 0
  ' Check if SHOVEL targets are completed with bitwise arithmetic
  vMaskedBits = avShovelTargetsHit(nCurPlayer) and cAllShovelTargets
  If vMaskedBits = cAllShovelTargets Then bShovelCompleted = true
  ' Check if KNIGHT targets are completed with bitwise arithmetic
  vMaskedBits = avShovelTargetsHit(nCurPlayer) and cAllKnightTargets
  If vMaskedBits = cAllKnightTargets Then bKnightCompleted = true

  If StandupS.isDropped <> 0 Then
    avShovelTargetsHit(nCurPlayer) = BitSetAdd(avShovelTargetsHit(nCurPlayer), 0)
  End If
  If StandupH1.isDropped <> 0 Then
    avShovelTargetsHit(nCurPlayer) = BitSetAdd(avShovelTargetsHit(nCurPlayer), 1)
  End If
  If StandupO.isDropped <> 0 Then
    avShovelTargetsHit(nCurPlayer) = BitSetAdd(avShovelTargetsHit(nCurPlayer), 2)
  End If

  ' Check if drop bank is complete
  If StandupS.isDropped And StandupH1.isDropped And StandupO.isDropped Then
    RaiseDrops
  End If

  If BitSetContains(vModesRunning, cModeShovel) Or _
  BitSetContains(vModesRunning, cModeSecret) Then
    ' Light super jackpot
    If bShovelCompleted And bKnightCompleted Then
      PlaySound "fx-fanfare2"
      if anKnightSJacks(nCurPlayer) mod 3 = 0 Then
        sSJMultiplier = ""
      Elseif anKnightSJacks(nCurPlayer) mod 3 = 1 Then
        sSJMultiplier = "2X "
      Else
        sSJMultiplier = "3X "
      end If

      avShovelTargetsHit(nCurPlayer) = 0
      ' unlight all SHOVEL KNIGHT lights until super is collected
      avShovelLightsMask(nCurPlayer) = cAllShovelTargets or cAllKnightTargets

      bShovelSuperJP = True
      AddShotColor cShotUpperHole, cBlinkPurple
      LightSJLit.state = LightStateBlinking

      ShowText sSJMultiplier & "SUPER JACKPOT", "IS LIT", 2000, 1
      UpdateModeDisplay True
    End If
  ElseIf BitSetContains(vModesRunning, cModeEnchantress) _
  Or BitSetContains(vModesRunning, cModeJuggling) Then
    Exit Sub
  Elseif anKnightMBalls(nCurPlayer) = 0 Then
    If bShovelCompleted or bKnightCompleted Then
      vLocksToLight = BitSetAdd(vLocksToLight, cLockLeft)
      vLocksToLight = BitSetAdd(vLocksToLight, cLockRight)
      LightLocks(vLocksToLight)
      ' All locks are lit, so unlight all SHOVEL KNIGHT lights
      avShovelLightsMask(nCurPlayer) = cAllShovelTargets or cAllKnightTargets
    End If

    If bShovelCompleted Then
      ' Clear SHOVEL standups
      avShovelTargetsHit(nCurPlayer) = avShovelTargetsHit(nCurPlayer) and not cAllShovelTargets
    Elseif bKnightCompleted Then
      ' Clear KNIGHT standups
      avShovelTargetsHit(nCurPlayer) = avShovelTargetsHit(nCurPlayer) and not cAllKnightTargets
    End If
  Elseif anKnightMBalls(nCurPlayer) = 1 Then
    If bShovelCompleted or bKnightCompleted Then
      If Not BitSetContains(avLocksLit(nCurPlayer), cLockLeft) Then
        vLocksToLight = BitSetAdd(vLocksToLight, cLockLeft)
        LightLocks(vLocksToLight)
      Else
        vLocksToLight = BitSetAdd(vLocksToLight, cLockRight)
        LightLocks(vLocksToLight)
        ' All locks are lit, so unlight all SHOVEL KNIGHT lights
        avShovelLightsMask(nCurPlayer) = cAllShovelTargets or cAllKnightTargets
      End If
    End If

    If bShovelCompleted Then
      ' Clear SHOVEL standups
      avShovelTargetsHit(nCurPlayer) = avShovelTargetsHit(nCurPlayer) and not cAllShovelTargets
    Elseif bKnightCompleted Then
      ' Clear KNIGHT standups
      avShovelTargetsHit(nCurPlayer) = avShovelTargetsHit(nCurPlayer) and not cAllKnightTargets
    End If
  Elseif anKnightMBalls(nCurPlayer) = 2 Then
    If bShovelCompleted And Not BitSetContains(avLocksLit(nCurPlayer), cLockLeft) Then
      vLocksToLight = BitSetAdd(vLocksToLight, cLockLeft)
      LightLocks(vLocksToLight)
      ' Clear SHOVEL standups
      avShovelTargetsHit(nCurPlayer) = avShovelTargetsHit(nCurPlayer) and not cAllShovelTargets
      ' unlight SHOVEL standups
      avShovelLightsMask(nCurPlayer) = avShovelLightsMask(nCurPlayer) or cAllShovelTargets
    Elseif bKnightCompleted Then
      vLocksToLight = BitSetAdd(vLocksToLight, cLockRight)
      LightLocks(vLocksToLight)
      ' Clear KNIGHT standups
      avShovelTargetsHit(nCurPlayer) = avShovelTargetsHit(nCurPlayer) and not cAllKnightTargets
      ' unlight KNIGHT standups
      avShovelLightsMask(nCurPlayer) = avShovelLightsMask(nCurPlayer) or cAllKnightTargets
    End If
  Elseif anKnightMBalls(nCurPlayer) >= 3 Then
    If bShovelCompleted And bKnightCompleted Then
      ' Clear all standups
      RaiseDrops
      avShovelTargetsHit(nCurPlayer) = 0
      If Not BitSetContains(avLocksLit(nCurPlayer), cLockLeft) Then
        vLocksToLight = BitSetAdd(vLocksToLight, cLockLeft)
        LightLocks(vLocksToLight)
      Else
        vLocksToLight = BitSetAdd(vLocksToLight, cLockRight)
        LightLocks(vLocksToLight)
        ' All locks are lit, so unlight all SHOVEL KNIGHT lights
        avShovelLightsMask(nCurPlayer) = cAllShovelTargets or cAllKnightTargets
      End If
    End If
  End If
  UpdateLightsStandups
End Sub

Function CountLocksMade
  Dim nLocks : nLocks = 0
  If BitSetContains(anLocksMade(nCurPlayer), cLockLeft) Then nLocks = nLocks + 1
  If BitSetContains(anLocksMade(nCurPlayer), cLockRight) Then nLocks = nLocks + 1
  CountLocksMade = nLocks
End Function

Sub LightLocks(vLocks)
  If BitSetIsFull(vLocks, 2) Then
    ShowText "LOCKS ARE LIT", "", 2000, 4
  Else
    ShowText "LOCK IS LIT", "", 2000, 4
  End If
  PlaySound "fx-fanfare1"
  If BitSetContains(vLocks, cLockLeft) Then
    If not bPhysLockLeft Then RaiseRampL
    avLocksLit(nCurPlayer) = BitSetAdd(avLocksLit(nCurPlayer), cLockLeft)
    AddShotColor cShotRampL, cBlinkGreen
  End If
  If BitSetContains(vLocks, cLockRight) Then
    If not bPhysLockRight Then RaiseRampR
    avLocksLit(nCurPlayer) = BitSetAdd(avLocksLit(nCurPlayer), cLockRight)
    AddShotColor cShotRampC, cBlinkGreen
  End If
  UpdateDiverter
End Sub

Sub VirtualLock(nShot)
  If nShot = cShotRampL Then
    anLocksMade(nCurPlayer) = BitSetAdd(anLocksMade(nCurPlayer), cLockLeft)
  Else
    anLocksMade(nCurPlayer) = BitSetAdd(anLocksMade(nCurPlayer), cLockRight)
  End If
  ShowText "BALL " & CountLocksMade & " LOCKED", "", 2000, 4
  UpdateLightsStandups

  RemoveShotColor nShot, cBlinkGreen
  anShotIdleColor(nShot) = cBlinkGreen
  If CountLocksMade = 2 Then
    AddShotColor cShotSaucerR, cBlinkGreen
  End If
End Sub

Sub RaiseRampR
  ramp3.HeightBottom = 80
  ramp3.Collidable = False
  RefreshLight.state = 1 : RefreshLight.state = 0
  If anComboTimers(cShotRampC) > 0 Then
    anComboTimers(cShotRampC) = 0
    RemoveShotColor cShotRampC, cBlinkYellow
  End If
End Sub

Sub RaiseRampL
  ramp7.HeightBottom = 80
  ramp7.Collidable = False
  RefreshLight.state = 1 : RefreshLight.state = 0
  If anComboTimers(cShotRampL) > 0 Then
    anComboTimers(cShotRampL) = 0
    RemoveShotColor cShotRampL, cBlinkYellow
  End If
End Sub

Sub LowerRampR
  ramp3.HeightBottom = 0
  ramp3.Collidable = True
  RefreshLight.state = 1 : RefreshLight.state = 0
End Sub

Sub LowerRampL
  ramp7.HeightBottom = 0
  ramp7.Collidable = True
  RefreshLight.state = 1 : RefreshLight.state = 0
End Sub

Sub RaiseDrops
  If StandupS.isDropped Or StandupH1.isDropped Or StandupO.isDropped Then
    vpmTimer.AddTimer 200, "PlayDropResetSound '"
  End If
  StandupS.isDropped = False
  StandupH1.isDropped = False
  StandupO.isDropped = False
End Sub

Sub PlayDropResetSound
  RandomSoundDropTargetReset StandupH1
  DOF 115, DOFPulse
End Sub

Sub UpdateRampL
  If KickerLRamp.timerEnabled or TimerLowerRampL.enabled Then
    Exit Sub
  End If

  ' Juggling mode
  If BitSetContains(vModesRunning, cModeJuggling) Then
    RaiseRampL
  ElseIf bPhysLockLeft Then
    LowerRampL
  Elseif BitSetContains(vModesRunning, cModeShovel) or _
  BitSetContains(vModesRunning, cModeSecret) Then
    if anKnightSJacks(nCurPlayer) mod 3 = 0 Then
      LowerRampL
    Elseif anKnightSJacks(nCurPlayer) mod 3 = 1 Then
      RaiseRampL
    Else
      LowerRampL
    end If
  elseIf BitSetContains(vModesRunning, cModeEnchantress) Then
      LowerRampL
  Else
    If BitSetContains(avLocksLit(nCurPlayer), cLockLeft) = True And _
    BitSetContains(anLocksMade(nCurPlayer), cLockLeft) = False Then
      RaiseRampL
    Else
      LowerRampL
    End If
  End If
End Sub

Sub UpdateRampC
  If KickerRRamp.timerEnabled or TimerLowerRampC.enabled Then
    Exit Sub
  End If

  ' Juggling mode
  If BitSetContains(vModesRunning, cModeJuggling) Then
    RaiseRampR
  ElseIf bPhysLockRight Then
    LowerRampR
  Elseif BitSetContains(vModesRunning, cModeShovel) or _
  BitSetContains(vModesRunning, cModeSecret) Then
    if anKnightSJacks(nCurPlayer) mod 3 = 0 Then
      LowerRampR
    Elseif anKnightSJacks(nCurPlayer) mod 3 = 1 Then
      RaiseRampR
    Else
      LowerRampR
    end If
  elseIf BitSetContains(vModesRunning, cModeEnchantress) Then
      LowerRampR
  else
    If BitSetContains(avLocksLit(nCurPlayer), cLockRight) = true And _
    BitSetContains(anLocksMade(nCurPlayer), cLockRight) = False Then
      RaiseRampR
    Else
      LowerRampR
    End If
  End If
End Sub

Sub UpdateDiverter
  Dim nLocksLitCount
  nLocksLitCount = 0
  If BitSetContains(avLocksLit(nCurPlayer), cLockLeft) Then
    nLocksLitCount = nLocksLitCount + 1
  End If
  If BitSetContains(avLocksLit(nCurPlayer), cLockRight) Then
    nLocksLitCount = nLocksLitCount + 1
  End If

  If BitSetContains(vModesRunning, cModePropeller) Then
    Diverter1R.isDropped = False
    Diverter1L.isDropped = True
    DropShop.isDropped = True
  Elseif BitSetContains(vModesRunning, cModeShovel) or _
  BitSetContains(vModesRunning, cModeSecret) Then
    if anKnightSJacks(nCurPlayer) mod 3 = 0 Then
      Diverter1R.isDropped = True
      Diverter1L.isDropped = False
      DropShop.isDropped = True
    Elseif anKnightSJacks(nCurPlayer) mod 3 = 1 Then
      Diverter1R.isDropped = False
      Diverter1L.isDropped = True
      DropShop.isDropped = True
    Else
      Diverter1R.isDropped = False
      Diverter1L.isDropped = True
      DropShop.isDropped = False
    end If
  Elseif ((CountLocksMade + nLocksLitCount) < 2) And _
  ((anKnightMBalls(nCurPlayer) = 4 And nLocksLitCount > 0) Or _
  (anKnightMBalls(nCurPlayer) > 4)) Then
    Diverter1R.isDropped = False
    Diverter1L.isDropped = True
    DropShop.isDropped = True
  Else
    Diverter1R.isDropped = True
    Diverter1L.isDropped = False
    DropShop.isDropped = True
  End If

  If BitSetContains(vModesRunning, cModeJuggling) _
  And anJuggleTimers(cShotShop) > 0 Then
    DropShop.IsDropped = False
  End If
End Sub

Sub StartModes(Byval vModesToStart)
  Dim i, j
  Dim nStolenLocks

  If BitSetContains(vModesToStart, cModeShovel) Then
    ' Remove green indicator that lock is lit or made
    RemoveShotColor cShotSaucerR, cBlinkGreen
    anShotIdleColor(cShotRampL) = -1
    anShotIdleColor(cShotRampC) = -1
    ' Eject balls from locks under ramps
    ' Handle stolen locks
    nStolenLocks = 0
    If bPhysLockLeft Then
      RaiseRampL
      KickerLRamp.timerEnabled = True
    Else
      nStolenLocks = nStolenLocks + 1
    End If
    If bPhysLockRight Then
      RaiseRampR
      TimerEjectBallC.enabled = True
    Else
      nStolenLocks = nStolenLocks + 1
    End If
    BIP = BIP + 2
    If nStolenLocks > 0 Then AutoPlunge nStolenLocks
    ' Update count of locks lit and made
    avLocksLit(nCurPlayer) = 0
    anLocksMade(nCurPlayer) = 0
    TimerEjectBallR.enabled = False
    TimerEjectBallR.interval = 0
    TimerEjectBallR.interval = 1000
    TimerEjectBallR.enabled = True
    ' Update drop and standup targets
    avShovelTargetsHit(nCurPlayer) = 0
    RaiseDrops
    avShovelLightsMask(nCurPlayer) = 0
    UpdateLightsStandups
    ' Set add-a-ball as unused for this multiball
    bAddABallUsed = False
    ' Start ball saver
    nBallSaveTimer = cBallSaveMball
    LightShootAgain.state = LightStateBlinking
    ' Update which modes are running
    EndCombos
    vModesRunning = BitSetAdd(vModesRunning, cModeShovel)
    ' Update boss Light
    UpdateLightsBoss
    ' Update how many SK multiballs have been played
    anKnightMBalls(nCurPlayer) = anKnightMBalls(nCurPlayer) + 1
    ' Add jackpots and color indicators to shots
    anShovelJackpots(0) = cShotRampL
    AddShotColor cShotRampL, cBlinkAmber
    anShovelJackpots(1) = cShotRampR
    AddShotColor cShotRampR, cBlinkOrange
    anShovelJackpots(2) = cShotRampC
    AddShotColor cShotRampC, cBlinkRed
    bShovelSuperJP = False
    ' Play proper Music
    SelectMusic False
    ' display proper text
    UpdateModeDisplay False
    ' Set ramps and diverter to correct state
    UpdateDiverter
    UpdateRampL
    UpdateRampC
    ' Have the lights do a barrel roll
    LightShowSpin 5, 1
  End If

  If BitSetContains(vModesToStart, cModeJuggling) Then
    vModesRunning = BitSetAdd(vModesRunning, cModeJuggling)
    EndCombos
    anComboJPSum(nCurPlayer) = 0
    RemoveShotColor cShotSaucerR, cBlinkPurple
    RemoveShotColor cShotShop, cBlinkPink
    LightSJLit.state = LightStateOff
    ClearTextQueue
    QueueText "JUGGLING", "CHALLENGE", 1500
    QueueText "LOCK EACH BALL", "IN A SAUCER", 1500
    for each i in anJugglingShots
      anJuggleTimers(i) = 0
    next
    nJuggleLives = 3
    UpdateLightsLocks
    UpdateLightsBoss
    UpdateLightsWanderer
    UpdateLightsStandups
    SelectMusic False
    ' start ball saver
    nBallSaveTimer = cBallSaveMball
    LightShootAgain.state = LightStateBlinking
    LeftSmallFlipper.RotateToStart
    RightSmallFlipper.RotateToStart
    ' Extend the saucer's kickout Timer
    KickerSaucerR.timerEnabled = False
    TimerEjectBallR.interval = 3000
    TimerEjectBallR.enabled = True
    bBossStarted = True
  End If

  Dim nWanderer
  ' FIXME: secret battle
  If BitSetContains(vModesToStart, cModeWanderer) Then
    vModesRunning = BitSetAdd(vModesRunning, cModeWanderer)
    nWanderer = anWandererSelected(nCurPlayer)
    anWanderersPlayed(nCurPlayer) = anWanderersPlayed(nCurPlayer) + 1
    RemoveShotColor cShotUpperHole, cBlinkCyan
    QueueText asWandererNames(nWanderer), "MULTIBALL", 2000
    If nWanderer = cWandererBlack Or nWanderer = cWandererAll Then
      for i = 0 to 11
        aoShovelLights(i).color = vbCyan
        aoShovelLights(i).colorFull = Color_LtCyan
      Next
      UpdateLightsStandups
    End If
    If nWanderer = cWandererPhantom Or nWanderer = cWandererAll Then
      AddShotColor cShotLoopL, cBlinkCyan
      AddShotColor cShotLoopR, cBlinkCyan
    End If
    If nWanderer = cWandererReize Or nWanderer = cWandererAll Then
      AddShotColor cShotShop, cBlinkCyan
      AddShotColor cShotSaucerR, cBlinkCyan
      AddShotColor cShotUpperHole, cBlinkCyan
    End If
    If nWanderer = cWandererBaz Or nWanderer = cWandererAll Then
      AddShotColor cShotRampL, cBlinkCyan
      AddShotColor cShotRampC, cBlinkCyan
      AddShotColor cShotRampR, cBlinkCyan
    End If
    UpdateLightsLocks
    UpdateLightsBoss
    UpdateLightsWanderer
    SelectMusic False
    ' start ball saver
    nBallSaveTimer = cBallSaveMball
    LightShootAgain.state = LightStateBlinking
    AutoPlunge 1
    BIP = BIP + 1
    UpdateModeDisplay True
  End If

  Dim nWizardTimer
  Dim nBadMedals
  Dim nBronzeMedals
  Dim nSilverMedals
  Dim nGoldMedals
  Dim nTextTime
  Dim nHoldBallTime
  Dim nBossExtraTime

  ' Boss battles
  Dim nBossesStarted : nBossesStarted = 0

  If BitSetContains(vModesToStart, cModeEnchantress) Then
    vModesRunning = BitSetAdd(vModesRunning, cModeEnchantress)
    nWizardTimer = 2000
    nBadMedals = 0
    nBronzeMedals = 0
    nSilverMedals = 0
    nGoldMedals = 0
    EndCombos
    for i = cModeBlack to cModePropeller
      anModeScore(i) = 0
      anModeTimeLeft(i) = 0
      anBossShots(i) = 0
      Select Case aanBossMedals(nCurPlayer, i)
        Case cMedalConsolation
          nBadMedals = nBadMedals + 1
          nWizardTimer = nWizardTimer + 1000
        Case cMedalBronze
          nBronzeMedals = nBronzeMedals + 1
          nWizardTimer = nWizardTimer + 2000
        Case cMedalSilver
          nSilverMedals = nSilverMedals + 1
          nWizardTimer = nWizardTimer + 5000
        Case cMedalGold
          nGoldMedals = nGoldMedals + 1
          nWizardTimer = nWizardTimer + 8000
      End Select
    Next

    nTextDuration = 0
    ClearTextQueue
    QueueText "ENCHANTRESS", "MULTIBALL", 2000
    If nBadMedals = 1 Then
      QueueText "1 PARTICIPATION", "TROPHY +1 SEC", 1500
    Else
      QueueText nBadMedals & " PARTICIPATION", "TROPHIES +" & _
      nBadMedals & " SEC", 1500
    End If
    If nBronzeMedals = 1 Then
      QueueText "1 BRONZE MEDAL", "2 SECONDS ADDED", 1500
    Else
      QueueText nBronzeMedals & " BRONZE MEDALS", _
      (nBronzeMedals * 2) & " SECONDS ADDED", 1500
    End If
    If nSilverMedals = 1 Then
      QueueText "1 SILVER MEDAL", "5 SECONDS ADDED", 1500
    Else
      QueueText nSilverMedals & " SILVER MEDALS", _
      (nSilverMedals * 5) & " SECONDS ADDED", 1500
    End If
    If nGoldMedals = 1 Then
      QueueText "1 GOLD MEDAL", "8 SECONDS ADDED", 1500
    Else
      QueueText nGoldMedals & " GOLD MEDALS", _
      (nGoldMedals * 8) & " SECONDS ADDED", 1500
    End If
    QueueText "BATTLE EVERY", "KNIGHT AT ONCE", 1500
    QueueText "UNLIMITED BALLS", "FOR " & CInt((nWizardTimer / 1000) - 2)  & " SECONDS", 2000

    nBlackKnightCycle = 0
    for i = 0 to 5
      aoShovelLights(i).state = LightStateOff
      If i Mod 2 = 1 Then
        aoShovelLights(i).BlinkPattern = "01"
      End If
      aoShovelLights(i).BlinkInterval = 75
      aoShovelLights(i).state = LightStateBlinking
    Next
    for i = 6 to 11
      aoShovelLights(i).BlinkInterval = 150
    Next
    TimerKnightCycle.enabled = True
    TimerKnightCycle_Timer
    RemoveShotColor cShotShop, cBlinkPink
    RemoveShotColor cShotUpperHole, cBlinkCyan
    AddShotColor cShotLoopR, cBlinkBlue
    AddShotColor cShotRampR, cBlinkBlue
    AddShotColor cShotRampL, cBlinkBlue
    AddShotColor cShotSaucerR, cBlinkBlue
    AddShotColor cShotShop, cBlinkBlue
    AddShotColor cShotVUK, cBlinkBlue
    AddShotColor cShotCaptive, cBlinkBlue
    AddShotColor cShotRampC, cBlinkBlue
    vTreasureHits = 0
    PlagueBlink.Enabled = True
    LightTroupple1.BlinkInterval = 75
    LightTroupple1.state = LightStateOff
    LightTroupple1.state = LightStateBlinking
    LightTroupple2.BlinkInterval = 75
    LightTroupple2.state = LightStateOff
    LightTroupple2.BlinkPattern = "01"
    LightTroupple2.state = LightStateBlinking
    LightTroupple3.BlinkInterval = 75
    LightTroupple3.state = LightStateOff
    LightTroupple3.state = LightStateBlinking
    RaiseDrops
    UpdateRampL
    UpdateRampC
    UpdateDiverter
    UpdateLightsWanderer

    anModeTimeLeft(cModeEnchantress) = nWizardTimer
    nBallSaveTimer = nWizardTimer - 2000

    LightShowSpin 10, 5
    ' Extend the saucer's kickout Timer
    KickerSaucerR.timerEnabled = False
    TimerEjectBallR.interval = 11500
    TimerEjectBallR.enabled = True
    bBossStarted = True

    SelectMusic False
    Exit Sub
  End If

  for i = cModeBlack to cModePropeller
    If BitSetContains(vModesToStart, i) Then
      nBossesStarted = nBossesStarted + 1

      anModeScore(i) = 0
      anBossShots(i) = 0
      avBossesLit(nCurPlayer) = BitSetRemove(avBossesLit(nCurPlayer), i)
    End If
  Next

  ' Show boss descriptions  - faster if there are more bosses
  If nBossesStarted = 1 Then
    nTextTime = 3000
    nHoldBallTime = 3000
    nBossExtraTime = 0
  ElseIf nBossesStarted = 2 Then
    nTextTime = 2000
    nHoldBallTime = 4000
    nBossExtraTime = 6000
  ElseIf nBossesStarted = 3 Then
    nTextTime = 1500
    nHoldBallTime = 4500
    nBossExtraTime = 9000
  End If

  If nBossesStarted > 0 Then
'   nTextDuration = 0
'   ClearTextQueue
  End If

  If BitSetContains(vModesToStart, cModeBlack) Then
    QueueTextFirst asKnightNames(cModeBlack), "HIT K-N-I-G-H-T", nTextTime
    nBlackKnightCycle = 0
    anBossShots(cModeBlack) = 0
    for i = 6 to 11
      aoShovelLights(i).BlinkInterval = 150
    Next
    TimerKnightCycle.enabled = True
    ' Start the first update right now rather that when te timer ticks
    TimerKnightCycle_Timer
  End If
  If BitSetContains(vModesToStart, cModeKing) Then
    QueueTextFirst "KING KNIGHT", "HIT THE SPINNER", nTextTime
    AddShotColor cShotLoopR, cBlinkBlue
  End If
  If BitSetContains(vModesToStart, cModeSpectre) Then
    QueueTextFirst "SPECTRE KNIGHT", "LEFT+RIGHT RAMP", nTextTime
    AddShotColor cShotRampR, cBlinkBlue
    AddShotColor cShotRampL, cBlinkBlue
  End If
  If BitSetContains(vModesToStart, cModeTreasure) Then
    QueueTextFirst asKnightNames(cModeTreasure), "HIT THE SAUCERS", nTextTime
    vTreasureHits = 0
    AddShotColor cShotSaucerR, cBlinkBlue
    AddShotColor cShotShop, cBlinkBlue
    AddShotColor cShotVUK, cBlinkBlue
  End If
  If BitSetContains(vModesToStart, cModePlague) Then
    QueueTextFirst asKnightNames(cModePlague), "HIT THE BUMPERS", nTextTime
    PlagueBlink.Enabled = True
  End If
  If BitSetContains(vModesToStart, cModeMole) Then
    QueueTextFirst asKnightNames(cModeMole), "HIT RED TARGETS", nTextTime
    LightTroupple1.BlinkInterval = 75
    LightTroupple1.state = LightStateOff
    LightTroupple1.state = LightStateBlinking
    LightTroupple2.BlinkInterval = 75
    LightTroupple2.state = LightStateOff
    LightTroupple2.BlinkPattern = "01"
    LightTroupple2.state = LightStateBlinking
    LightTroupple3.BlinkInterval = 75
    LightTroupple3.state = LightStateOff
    LightTroupple3.state = LightStateBlinking
  End If
  If BitSetContains(vModesToStart, cModePolar) Then
    RaiseDrops
    QueueTextFirst asKnightNames(cModePolar), "HIT S-H-O-V-E-L", nTextTime
    for i = 0 to 5
      aoShovelLights(i).state = LightStateOff
      If i Mod 2 = 1 Then
        aoShovelLights(i).BlinkPattern = "01"
      End If
      aoShovelLights(i).BlinkInterval = 75
      aoShovelLights(i).state = LightStateBlinking
    Next
  End If
  If BitSetContains(vModesToStart, cModeTinker) Then
    QueueTextFirst asKnightNames(cModeTinker), "CAPTIVE BALL", nTextTime
    AddShotColor cShotCaptive, cBlinkBlue
  End If
  If BitSetContains(vModesToStart, cModePropeller) Then
    QueueTextFirst asKnightNames(cModePropeller), "HIT CENTER RAMP", nTextTime
    AddShotColor cShotRampC, cBlinkBlue
  End If

  ' Set mode timers - more time if more bosses are stacked
  For i = 0 to 8
    If BitSetContains(vModesToStart, i) Then
      anModeTimeLeft(i) = anBossTimers(i) + nBossExtraTime
      vModesRunning = BitSetAdd(vModesRunning, i)
    End If
  Next

  If nBossesStarted > 0 Then
    ' Extend the saucer's kickout Timer
    KickerSaucerR.timerEnabled = False
    TimerEjectBallR.interval = nHoldBallTime
    TimerEjectBallR.enabled = True
    bBossStarted = True
    UpdateDiverter

    UpdateLightsBoss
    SelectMusic False
  End If

  ' Hall of Champions
  If BitSetContains(vModesToStart, cModeChampions) Then
    vModesRunning = BitSetAdd(vModesRunning, cModeChampions)
    ClearTextQueue
    QueueText "ENTER THE HALL", "OF CHAMPIONS", 1500
    QueueText "MAKE ALL WHITE", "SHOTS IN TIME", 1500
    QueueText "ALL SHOTS ARE", "WORTH DOUBLE", 1500
    ' disable inlanes
    avBottomLanes(nCurPlayer) = 0
    UpdateLightsLanes
    ' Set multiplier Timers
    anModeTimeLeft(cModeChampions) = 22000
    vMultipliersLit = 0
    for i = 0 to 9
      aanXTimers(nCurPlayer, i) = 60000
      vMultipliersLit = BitSetAdd(vMultipliersLit, i)
      AddShotColor i, cBlinkWhite
    next
    ' start multipliers, update lights
    UpdateLightsMultipliers
    SelectMusic False

    ' Extend the saucer's kickout Timer
    KickerSaucerR.timerEnabled = False
    TimerEjectBallR.interval = 4500
    TimerEjectBallR.enabled = True
    bBossStarted = True
  End If

  If BitSetContains(vModesToStart, cModeSuperCombo) Then
    vModesRunning = BitSetAdd(vModesRunning, cModeSuperCombo)
    ShowText anComboCount(nCurPlayer) & " COMBOS", "SUPER COMBO", 1500, 4
    anModeTimeLeft(cModeSuperCombo) = 32000
    SelectMusic False
  End If

  If BitSetContains(vModesToStart, cModeSecret) Then
    vModesRunning = BitSetAdd(vModesRunning, cModeSecret)
    avSecretProgress(nCurPlayer) = 0
    ClearTextQueue
    QueueText "SECRET FIGHT", "VS BATTLETOADS", 1500
    QueueText "SHOVEL MULTIBALL", "IS RUNNING", 1500
    QueueText "ENCHANTRESS", "IS RUNNING", 1500
    QueueText "HALL OF CHAMPS", "IS RUNNING", 1500
    QueueText "SUPER WANDERER", "IS RUNNING", 1500
    QueueText "SCORE " & WilliamsFormatNum(500000), "POINTS TO WIN", 1500
    anModeScore(cModeSecret) = 0

    ' Shovel multiball
    vModesRunning = BitSetAdd(vModesRunning, cModeShovel)
    anShotIdleColor(cShotRampL) = -1
    anShotIdleColor(cShotRampC) = -1
    anShovelJackpots(0) = cShotRampL
    AddShotColor cShotRampL, cBlinkAmber
    anShovelJackpots(1) = cShotRampR
    AddShotColor cShotRampR, cBlinkOrange
    anShovelJackpots(2) = cShotRampC
    AddShotColor cShotRampC, cBlinkRed
    bShovelSuperJP = False

    'Hall of Champions
    vModesRunning = BitSetAdd(vModesRunning, cModeChampions)
    ' disable inlanes
    avBottomLanes(nCurPlayer) = 0
    UpdateLightsLanes
    ' Set multiplier Timers
    anModeTimeLeft(cModeChampions) = 22000
    vMultipliersLit = 0
    for i = 0 to 9
      aanXTimers(nCurPlayer, i) = 60000
      vMultipliersLit = BitSetAdd(vMultipliersLit, i)
      AddShotColor i, cBlinkWhite
    next

    'Super Wanderer
    vModesRunning = BitSetAdd(vModesRunning, cModeWanderer)
    anWandererSelected(nCurPlayer) = cWandererAll
    for i = 0 to 11
      aoShovelLights(i).color = vbCyan
      aoShovelLights(i).colorFull = Color_LtCyan
    Next
    AddShotColor cShotLoopL, cBlinkCyan
    AddShotColor cShotLoopR, cBlinkCyan
    AddShotColor cShotShop, cBlinkCyan
    AddShotColor cShotSaucerR, cBlinkCyan
    AddShotColor cShotUpperHole, cBlinkCyan
    AddShotColor cShotRampL, cBlinkCyan
    AddShotColor cShotRampC, cBlinkCyan
    AddShotColor cShotRampR, cBlinkCyan

    'Enchantress
    vModesRunning = BitSetAdd(vModesRunning, cModeEnchantress)
    anModeTimeLeft(cModeEnchantress) = 999999999
    nBlackKnightCycle = 0
    for i = 0 to 5
      aoShovelLights(i).state = LightStateOff
      If i Mod 2 = 1 Then
        aoShovelLights(i).BlinkPattern = "01"
      End If
      aoShovelLights(i).BlinkInterval = 75
      aoShovelLights(i).state = LightStateBlinking
    Next
    for i = 6 to 11
      aoShovelLights(i).BlinkInterval = 150
    Next
    TimerKnightCycle.enabled = True
    TimerKnightCycle_Timer
    RemoveShotColor cShotShop, cBlinkPink
    RemoveShotColor cShotUpperHole, cBlinkCyan
    AddShotColor cShotLoopR, cBlinkBlue
    AddShotColor cShotRampR, cBlinkBlue
    AddShotColor cShotRampL, cBlinkBlue
    AddShotColor cShotSaucerR, cBlinkBlue
    AddShotColor cShotShop, cBlinkBlue
    AddShotColor cShotVUK, cBlinkBlue
    AddShotColor cShotCaptive, cBlinkBlue
    AddShotColor cShotRampC, cBlinkBlue
    vTreasureHits = 0
    PlagueBlink.Enabled = True
    LightTroupple1.BlinkInterval = 75
    LightTroupple1.state = LightStateOff
    LightTroupple1.state = LightStateBlinking
    LightTroupple2.BlinkInterval = 75
    LightTroupple2.state = LightStateOff
    LightTroupple2.BlinkPattern = "01"
    LightTroupple2.state = LightStateBlinking
    LightTroupple3.BlinkInterval = 75
    LightTroupple3.state = LightStateOff
    LightTroupple3.state = LightStateBlinking

    nBallSaveTimer = 30000
    LightShootAgain.state = LightStateBlinking
    LightShowSpin 10, 5
    ' Extend the saucer's kickout Timer
    KickerSaucerR.timerEnabled = False
    TimerEjectBallR.interval = 12000
    TimerEjectBallR.enabled = True
    bBossStarted = True

    SelectMusic False
    UpdateLightsStandups
    UpdateLightsBoss
    UpdateDiverter
    UpdateRampL
    UpdateRampC
  End if
End Sub

Sub EndModes(ByVal vModesToEnd)
  Dim i, j
  Dim nBossesEnded : nBossesEnded = 0
  Dim bLongerBossRunning
  Dim nWizardTotal

  If BitSetContains(vModesToEnd, cModeShovel) Then
    ClearTextQueue
    vModesRunning = BitSetRemove(vModesRunning, cModeShovel)
    avShovelTargetsHit(nCurPlayer) = 0
    RaiseDrops
    avShovelLightsMask(nCurPlayer) = 0
    UpdateLightsStandups
    UpdateLightsBoss
    UpdateDiverter
    UpdateRampL
    UpdateRampC
    RemoveShotColor anShovelJackpots(0), cBlinkAmber
    RemoveShotColor anShovelJackpots(1), cBlinkOrange
    RemoveShotColor anShovelJackpots(2), cBlinkRed
    If bShovelSuperJP Then
      RemoveShotColor cShotUpperHole, cBlinkPurple
      LightSJLit.state = LightStateOff
    End If
    If Not bDrainingBalls Then
      SelectMusic False
      anGracePeriod(cModeShovel) = 3000
    End If
    UpdateModeDisplay True
    ChangeLevelToStart
  End If

  If BitSetContains(vModesToEnd, cModeChampions) Then
    vModesRunning = BitSetRemove(vModesRunning, cModeChampions)
    ShowText "HALL OF CHAMPS", _
      "TOTAL " & WilliamsFormatNum(anModeScore(cModeChampions)), 2500, 2
    If anModeScore(cModeChampions) > 109000 Then
      InitMusic StopMusic
      PlaySound "fx-superjackpot3x"
      TimerSJackJingle.enabled = True
    End If
    ' Set multiplier Timers
    vMultipliersLit = 0
    avXUsed(nCurPlayer) = 0
    anXCompleted(nCurPlayer) = anXCompleted(nCurPlayer) + 1
    for i = 0 to 9
      aanXTimers(nCurPlayer, i) = 0
      vMultipliersLit = BitSetRemove(vMultipliersLit, i)
      RemoveShotColor i, cBlinkWhite
    next
    ' start multipliers, update lights
    UpdateLightsMultipliers
    ClearTextQueue
    UpdateModeDisplay True
    SelectMusic False
  End If

  If BitSetContains(vModesToEnd, cModeSuperCombo) Then
    vModesRunning = BitSetRemove(vModesRunning, cModeSuperCombo)
    UpdateModeDisplay True
    SelectMusic False
  End If

  Dim nWanderer
  ' FIXME: secret battle
  If BitSetContains(vModesToEnd, cModeWanderer) Then
    vModesRunning = BitSetRemove(vModesRunning, cModeWanderer)
    nWanderer = anWandererSelected(nCurPlayer)
    If nWanderer = cWandererBlack Or nWanderer = cWandererAll Then
      for i = 0 to 11
        aoShovelLights(i).color = vbGreen
        aoShovelLights(i).colorFull = Color_LtGreen
      Next
    End If
    If nWanderer = cWandererPhantom Or nWanderer = cWandererAll Then
      RemoveShotColor cShotLoopL, cBlinkCyan
      RemoveShotColor cShotLoopR, cBlinkCyan
    End If
    If nWanderer = cWandererReize Or nWanderer = cWandererAll Then
      RemoveShotColor cShotShop, cBlinkCyan
      RemoveShotColor cShotSaucerR, cBlinkCyan
      RemoveShotColor cShotUpperHole, cBlinkCyan
    End If
    If nWanderer = cWandererBaz Or nWanderer = cWandererAll Then
      RemoveShotColor cShotRampL, cBlinkCyan
      RemoveShotColor cShotRampC, cBlinkCyan
      RemoveShotColor cShotRampR, cBlinkCyan
    End If
    anWandererHits(nCurPlayer) = 60 + (anWanderersPlayed(nCurPlayer) * 10)
    If anWandererHits(nCurPlayer) > 150 then anWandererHits(nCurPlayer) = 150
    If nWanderer = cWandererAll Then
      anWandererSelected(nCurPlayer) = cWandererBlack
    End If
    UpdateLightsLocks
    UpdateLightsBoss
    ChangeWanderer
    UpdateLightsWanderer
    UpdateLightsStandups
    SelectMusic False
    UpdateModeDisplay True
  End If

  If BitSetContains(vModesToEnd, cModeJuggling) Then
    vModesRunning = BitSetRemove(vModesRunning, cModeJuggling)

    KickerLRamp.Kick -90, 8
    KickerRRamp.Kick -80, 12
    KickerShop.Kick 335, 35
    DropShop.IsDropped = True
    KickerVUK.Kick  0, 60, 1.56
    KickerSaucerR.Kick 0, 40, 1.56

    QueueText "JUGGLING ENDED", "DRAINING BALLS", 4000
    UpdateLightsStandups
    UpdateRampC
    UpdateRampL
    UpdateDiverter
    For each i in anJugglingShots
      anJuggleTimers(i) = 0
      RemoveShotColor i, cBlinkPurple
      anShotIdleColor(i) = -1
    Next

    nJuggleLives = 0
    bDrainingBalls = True
    Rightslingshot.disabled = True
    Leftslingshot.disabled = True
    Bumper1.HasHitEvent = False
    Bumper2.HasHitEvent = False
    Bumper3.HasHitEvent = False
    LeftFlipper.RotateToStart
    UpperLeftFlipper.RotateToStart
    LeftSmallFlipper.RotateToStart
    RightFlipper.RotateToStart
    RightSmallFlipper.RotateToStart

    If anTrouppleLit(nCurPlayer) < 1 Then
      AddShotColor cShotShop, cBlinkPink
    End If
    UpdateLightsLocks
    UpdateLightsBoss
    UpdateLightsWanderer
    UpdateLightsStandups
    UpdateModeDisplay True
    SelectMusic False
  End If

  bLongerBossRunning = False
  For i = 0 to 8
    If anModeTimeLeft (i) > 1000 And BitSetContains(vModesRunning, i) Then
      bLongerBossRunning = True
      Exit For
    End If
  Next

  For i = 0 to 8
    If BitSetContains(vModesToEnd, i) Then
      vModesRunning = BitSetRemove(vModesRunning, i)
      anModeTimeLeft(i) = 0
      ' clear blinking lights for mode shots
      If i = cModeKing Then
        RemoveShotColor cShotLoopR, cBlinkBlue
      Elseif i = cModeSpectre Then
        RemoveShotColor cShotRampR, cBlinkBlue
        RemoveShotColor cShotRampL, cBlinkBlue
      Elseif i = cModeTreasure Then
        RemoveShotColor cShotSaucerR, cBlinkBlue
        RemoveShotColor cShotShop, cBlinkBlue
        RemoveShotColor cShotVUK, cBlinkBlue
      Elseif i = cModeMole Then
        LightTroupple1.BlinkInterval = 300
        LightTroupple1.state = LightStateOff
        LightTroupple2.BlinkInterval = 300
        LightTroupple2.state = LightStateOff
        LightTroupple2.BlinkPattern = "10"
        LightTroupple3.BlinkInterval = 300
        LightTroupple3.state = LightStateOff
        ' Lights will be set correctly by LightTimer
      Elseif i = cModePlague Then
        PlagueBlink.Enabled = False
        FlBumperFadeTarget(1) = 0
        FlBumperFadeTarget(2) = 0
        FlBumperFadeTarget(3) = 0
      ElseIf i = cModeBlack Then
        for j = 6 to 11
          aoShovelLights(j).state = LightStateOff
          aoShovelLights(j).BlinkInterval = 250
        Next
        TimerKnightCycle.enabled = False
        UpdateLightsStandups
      Elseif i = cModePolar Then
        for j = 0 to 5
          aoShovelLights(j).state = LightStateOff
          aoShovelLights(j).BlinkPattern = "10"
          aoShovelLights(j).BlinkInterval = 250
        Next
        UpdateLightsStandups
      Elseif i = cModeTinker Then
        RemoveShotColor cShotCaptive, cBlinkBlue
      Elseif i = cModePropeller Then
        UpdateDiverter
        RemoveShotColor cShotRampC, cBlinkBlue
      End If

      If anBossShots(i) = 0 Then
        aanBossMedals(nCurPlayer, i) = cMedalConsolation
      ElseIf anBossShots(i) < anSilverScores(i) Then
        aanBossMedals(nCurPlayer, i) = cMedalBronze
      ElseIf anBossShots(i) < anGoldScores(i) Then
        aanBossMedals(nCurPlayer, i) = cMedalSilver
      Else
        aanBossMedals(nCurPlayer, i) = cMedalGold
      End If
      If BIP > 0 Then
        If nBossesEnded < 1 And bDrainingBalls = False Then
          If Not bLongerBossRunning Then
            InitMusic StopMusic
            PlaySound "fx-victory", 0, 0.6
            TimerBossEndJingle.enabled = True
          End If
        End If
        If anBossShots(i) = 0 Then
          aanTotalMedals(nCurPlayer, cMedalConsolation) = aanTotalMedals(nCurPlayer, cMedalConsolation) + 1
          QueueTextFirst "PARTICIPATION", "PRIZE AWARDED", 1500
        ElseIf anBossShots(i) < anSilverScores(i) Then
          aanTotalMedals(nCurPlayer, cMedalBronze) = aanTotalMedals(nCurPlayer, cMedalBronze) + 1
          QueueTextFirst "BRONZE MEDAL", "AWARDED", 1500
        ElseIf anBossShots(i) < anGoldScores(i) Then
          aanTotalMedals(nCurPlayer, cMedalSilver) = aanTotalMedals(nCurPlayer, cMedalSilver) + 1
          QueueTextFirst "SILVER MEDAL", "AWARDED", 1500
        Else
          aanTotalMedals(nCurPlayer, cMedalGold) = aanTotalMedals(nCurPlayer, cMedalGold) + 1
          QueueTextFirst "GOLD MEDAL", "AWARDED", 1500
        End If
        QueueTextFirst asKnightNames(i), "TOTAL " & WilliamsFormatNum(anModeScore(i)), 1500
      End If
      nBossesEnded = nBossesEnded + 1
    End If
  Next
  If nBossesEnded > 0 Then
    ' The most recent boss might be Treasure knight, which has a shorter
    ' timer, and when it ends another boss may be running
    If bLongerBossRunning Then
      For i = 0 to 8
        If BitSetContains(vModesRunning, i) Then
          anMostRecentBoss(nCurPlayer) = i
        End If
      Next
    Else
      anMostRecentBoss(nCurPlayer) = cModeNone
    End If

    If False = abBossXBallWon(nCurPlayer) Then
      If nCountBossesPlayed > 3 Then
        abBossXBallWon(nCurPlayer) = True
        anExtraBallsLit(nCurPlayer) = anExtraBallsLit(nCurPlayer) + 1
        PlaySound "fx-savepoint"
        ShowText nCountBossesPlayed & " BOSSES FOUGHT", "LIGHT EXTRA BALL", 2500, 1
        UpdateLightsWanderer
      Else
        If nCountBossesPlayed = 1 Then
          QueueText nCountBossesPlayed & " BOSS FOUGHT", "EXTRA BALL AT 4", 2000
        Else
          QueueText nCountBossesPlayed & " BOSSES FOUGHT", "EXTRA BALL AT 4", 2000
        End If
      End If
    End If

    UpdateModeDisplay True
    If not abPlayingLevel(nCurPlayer) Then
      anLevelShotsMade(nCurPlayer) = 0
      anLevelShotsToBeat(nCurPlayer) = 1
      nCountUpGoal = 0
    End If
    UpdateLightsBoss
    ChangeLevelToStart
  End If

  If BitSetContains(vModesToEnd, cModeEnchantress) Then
    vModesRunning = BitSetRemove(vModesRunning, cModeEnchantress)
    anModeTimeLeft(cModeEnchantress) = 0
    nWizardTotal = 0

    If Not BitSetContains(vModesRunning, cModeSecret) Then
      For i = 0 to 8
        aanBossMedals(nCurPlayer, i) = cMedalNotPlayed
        nWizardTotal = nWizardTotal + anModeScore(i)
      Next
      anLevelShotsMade(nCurPlayer) = 0
      anLevelShotsToBeat(nCurPlayer) = 1
      nCountUpGoal = 0
      ChangeLevelToStart

      ClearTextQueue
      If False = BitSetContains(vModesRunning, cModeSecret) Then
      ShowText "ENCHANTRESS", "TOTAL " & WilliamsFormatNum(nWizardTotal), PlayUntilHit, 2
      If nWizardTotal >= 125000 then
        avSecretProgress(nCurPlayer) = _
          BitSetAdd(avSecretProgress(nCurPlayer), cSecretEnchantress)
      End If
      End If
    End if

    If anTrouppleLit(nCurPlayer) < 1 Then
      AddShotColor cShotShop, cBlinkPink
    End If
    RemoveShotColor cShotLoopR, cBlinkBlue
    RemoveShotColor cShotRampR, cBlinkBlue
    RemoveShotColor cShotRampL, cBlinkBlue
    RemoveShotColor cShotSaucerR, cBlinkBlue
    RemoveShotColor cShotShop, cBlinkBlue
    RemoveShotColor cShotVUK, cBlinkBlue
    RemoveShotColor cShotCaptive, cBlinkBlue
    RemoveShotColor cShotRampC, cBlinkBlue
    LightTroupple1.BlinkInterval = 300
    LightTroupple1.state = LightStateOff
    LightTroupple2.BlinkInterval = 300
    LightTroupple2.state = LightStateOff
    LightTroupple2.BlinkPattern = "10"
    LightTroupple3.BlinkInterval = 300
    LightTroupple3.state = LightStateOff
    ' Lights will be set correctly by LightTimer
    PlagueBlink.Enabled = False
    FlBumperFadeTarget(1) = 0
    FlBumperFadeTarget(2) = 0
    FlBumperFadeTarget(3) = 0
    for j = 0 to 11
      aoShovelLights(j).state = LightStateOff
      aoShovelLights(j).BlinkPattern = "10"
      aoShovelLights(j).BlinkInterval = 250
    Next
    TimerKnightCycle.enabled = False
    UpdateLightsStandups
    UpdateLightsWanderer
    UpdateLightsLocks
    UpdateRampC
    UpdateRampL
    UpdateDiverter

    If BIP > 0 Then
      bDrainingBalls = True
      Rightslingshot.disabled = True
      Leftslingshot.disabled = True
      Bumper1.HasHitEvent = False
      Bumper2.HasHitEvent = False
      Bumper3.HasHitEvent = False
      LeftFlipper.RotateToStart
      UpperLeftFlipper.RotateToStart
      LeftSmallFlipper.RotateToStart
      RightFlipper.RotateToStart
      RightSmallFlipper.RotateToStart
    End If
    If Not BitSetContains(vModesRunning, cModeSecret) Then
      anLevelSelected(nCurPlayer) = cModeBlack
      InitMusic 13
    End If
  End If

  If BitSetContains(vModesToEnd, cModeSecret) Then
    vModesRunning = BitSetRemove(vModesRunning, cModeSecret)
    If anModeScore(cModeSecret) > 499999 Then
      anScore(nCurPlayer) = anScore(nCurPlayer) + 500000
      ShowText "BATTLETOADS", "DEFEATED " & WilliamsFormatNum(500000), PlayUntilHit, 1
    Else
      ShowText "BATTLETOADS", "TOTAL " & WilliamsFormatNum(anModeScore(cModeSecret)), PlayUntilHit, 1
    End If
    nBallSaveTimer = 0
    If nExtraBalls > 0 Then
      LightShootAgain.state = LightStateOn
    Else
      LightShootAgain.state = LightStateOff
    End If
    UpdateDiverter
    UpdateRampL
    UpdateRampC
    SelectMusic False
  End If
End Sub

Sub EndMultiplier(nShot)
  avXUsed(nCurPlayer) = BitSetAdd(avXUsed(nCurPlayer), nShot)
  ' Check if shot multiplier timed out while we have had the option of
  ' extending it
  If BitSetContains(vMultipliersLit, nShot) Then
    vMultipliersLit = BitSetRemove(vMultipliersLit, nShot)
    RemoveShotColor nShot, cBlinkWhite
  End If
  ' Check if Hall of Champions will be lit
  If False = BitSetContains(vModesRunning, cModeChampions) And _
  True = BitSetIsFull(avXUsed(nCurPlayer), 10) Then
    UpdateLightsBoss
  End If
  UpdateLightsMultipliers
End Sub

Sub CheckSecret
  Dim nModes
  If anModeScore(cModeSecret) > 499999 Then
    nModes = BitSetAdd(0, cModeChampions)
    nModes = BitSetAdd(nModes, cModeWanderer)
    nModes = BitSetAdd(nModes, cModeShovel)
    nModes = BitSetAdd(nModes, cModeEnchantress)
    EndModes nModes
    EndModes BitSetAdd(0, cModeSecret)
  End If
End Sub

Sub TimerMode_Timer
  Dim i, j
  Dim vModesTimedOut
  If nPlayers < 1 Then Exit Sub

  ' Halt timers while draining Balls
  If bDrainingBalls Then Exit Sub

  ' Do not count down timers while in the Bumpers, unless Plague Knight boss
  ' or a multiball is running
  If BIP < 2 And (GameTime - nTimeBumperHit < 500) And _
  False = BitSetContains(vModesRunning, cModePlague) Then
    Exit Sub
  End If

  If bBallLaunched And (Not bBallHeld) Then
    If False = BitSetContains(vModesRunning, cModeSecret) Then
      for i = 0 to 9
        ' Count down time for multipliers, and end multplier if time is out
        If aanXTimers(nCurPlayer, i) > 0 Then
          aanXTimers(nCurPlayer, i) = aanXTimers(nCurPlayer, i) - TimerMode.interval
          If aanXTimers(nCurPlayer, i) <= 0 Then
            aanXTimers(nCurPlayer, i) = 0
            EndMultiplier(i)
          End If
        End If
      Next
    End If

    for i = 0 to 18
      ' Count down time for grace periods
      If anGracePeriod(i) > 0 Then
        anGracePeriod(i) = anGracePeriod(i) - TimerMode.interval
        If anGracePeriod(i) <= 0 Then anGracePeriod(i) = 0
      End If
    Next

    vModesTimedOut = 0
    If true = BitSetContains(vModesRunning, cModeSecret) _
    And anModeTimeLeft(cModeChampions) < 1 Then
      ' Start over again!
      vMultipliersLit = 0
      anModeTimeLeft(cModeChampions) = 22000
      for i = 0 to 9
        vMultipliersLit = BitSetAdd(vMultipliersLit, i)
        AddShotColor i, cBlinkWhite
        aanXTimers(nCurPlayer, i) = 60000
      next
      ' start multipliers, update lights
      UpdateLightsMultipliers
    End If

    for each j in anTimedModes
      ' Count down time for modes, and end mode if time is out
      If BitSetContains(vModesRunning, j) Then
        anModeTimeLeft(j) = anModeTimeLeft(j) - TimerMode.interval
        If anModeTimeLeft(j) <= 0 Then
          anModeTimeLeft(j) = 0
          vModesTimedOut = BitSetAdd(vModesTimedOut, j)
        End If
      End If
    next
    If vModesTimedOut <> 0 Then
      EndModes vModesTimedOut
    End If

    If nBallSaveTimer > 0 Then
      nBallSaveTimer = nBallSaveTimer - TimerMode.interval
      ' Turn light off before the saver runs out, as a grace period
      If nBallSaveTimer < 3000 Then
        If nExtraBalls > 0 Then
          LightShootAgain.state = LightStateOn
        Else
          LightShootAgain.state = LightStateOff
        End If
      End If
      ' Make sure timer does not go negative
      If nBallSaveTimer <= 0 then nBallSaveTimer = 0
    End If
  End If
End Sub

'*****GI Lights On
dim xx

For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  RandomSoundSlingshotRight SLING1
  DOF 104, DOFPulse
  DOF 106, DOFPulse
  anScore(nCurPlayer) = anScore(nCurPlayer) + 70
  AnySwitchHit cSwitchBumper, 5
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  RandomSoundSlingshotLeft SLING2
  DOF 103, DOFPulse
  DOF 105, DOFPulse
  anScore(nCurPlayer) = anScore(nCurPlayer) + 70
  AnySwitchHit cSwitchBumper, 4
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

' Music

sub InitMusic(nTrack)
  Dim i

  if nTrack = nMusicCurrentTrack _
  Or (nTrack = 2 and nMusicCurrentTrack = 3) Then
    Exit Sub
  End if
  for each i in aoMusicTimers
    i.enabled = False
  Next
  MuteMusic
  nMusicCurrentTrack = nTrack
  if nTrack = StopMusic Then Exit Sub
  aoMusicTimers(nTrack).enabled = True
  PlaySound asMusicFiles(nTrack), 0, MusicVol
end sub

Sub MuteMusic
  StopSound "fx-victory"
  StopSound "fx-superjackpot1x"
  StopSound "fx-superjackpot2x"
  StopSound "fx-superjackpot3x"
  If nMusicCurrentTrack > -1 Then StopSound asMusicFiles(nMusicCurrentTrack)
End Sub

sub MusicTimer00_Timer
  InitMusic 1
end Sub

sub MusicTimer01_Timer
  PlaySound asMusicFiles(1), 0, MusicVol
end Sub

sub MusicTimer02_Timer
  InitMusic 3
end Sub

sub MusicTimer03_Timer
  PlaySound asMusicFiles(3), 0, MusicVol
end Sub

sub MusicTimer04_Timer
  PlaySound asMusicFiles(4), 0, MusicVol
end Sub

sub MusicTimer05_Timer
  PlaySound asMusicFiles(5), 0, MusicVol
end Sub

sub MusicTimer06_Timer
  PlaySound asMusicFiles(6), 0, MusicVol
end Sub

sub MusicTimer07_Timer
  PlaySound asMusicFiles(7), 0, MusicVol
end Sub

sub MusicTimer08_Timer
  PlaySound asMusicFiles(8), 0, MusicVol
end Sub

sub MusicTimer09_Timer
  PlaySound asMusicFiles(9), 0, MusicVol
end Sub

sub MusicTimer10_Timer
  PlaySound asMusicFiles(10), 0, MusicVol
end Sub

sub MusicTimer11_Timer
  PlaySound asMusicFiles(11), 0, MusicVol
end Sub

sub MusicTimer12_Timer
  PlaySound asMusicFiles(12), 0, MusicVol
end Sub

sub MusicTimer13_Timer
  PlaySound asMusicFiles(13), 0, MusicVol
end Sub

sub MusicTimer14_Timer
  PlaySound asMusicFiles(14), 0, MusicVol
end Sub

sub MusicTimer15_Timer
  PlaySound asMusicFiles(15), 0, MusicVol
end Sub

sub MusicTimer16_Timer
  PlaySound asMusicFiles(16), 0, MusicVol
end Sub

sub MusicTimer17_Timer
  InitMusic 18
end Sub

sub MusicTimer18_Timer
  PlaySound asMusicFiles(18), 0, MusicVol
end Sub

sub MusicTimer19_Timer
  PlaySound asMusicFiles(19), 0, MusicVol
end Sub

sub MusicTimer20_Timer
  PlaySound asMusicFiles(20), 0, MusicVol
end Sub

sub MusicTimer21_Timer
  PlaySound asMusicFiles(21), 0, MusicVol
end Sub

sub MusicTimer22_Timer
  PlaySound asMusicFiles(22), 0, MusicVol
end Sub

sub MusicTimer23_Timer
  PlaySound asMusicFiles(23), 0, MusicVol
end Sub

sub MusicTimer24_Timer
  PlaySound asMusicFiles(24), 0, MusicVol
end Sub

sub MusicTimer25_Timer
  PlaySound asMusicFiles(25), 0, MusicVol
end Sub

sub MusicTimer26_Timer
  PlaySound asMusicFiles(26), 0, MusicVol
end Sub

sub MusicTimer27_Timer
  PlaySound asMusicFiles(27), 0, MusicVol
end Sub

sub MusicTimer28_Timer
  PlaySound asMusicFiles(28), 0, MusicVol
end Sub

sub MusicTimer29_Timer
  PlaySound asMusicFiles(29), 0, MusicVol
end Sub

sub MusicTimer30_Timer
  PlaySound asMusicFiles(30), 0, MusicVol
end Sub

sub MusicTimer31_Timer
  PlaySound asMusicFiles(31), 0, MusicVol
end Sub

sub MusicTimer32_Timer
  PlaySound asMusicFiles(32), 0, MusicVol
end Sub

sub MusicTimer33_Timer
  PlaySound asMusicFiles(33), 0, MusicVol
end Sub

sub MusicTimer34_Timer
  PlaySound asMusicFiles(34), 0, MusicVol
end Sub

Sub TimerSJackJingle_Timer
  SelectMusic true
  TimerSJackJingle.enabled = False
End Sub

Sub TimerBossEndJingle_Timer
  SelectMusic true
  TimerBossEndJingle.enabled = False
End Sub

Sub TimerMysteryJingle_Timer
  SelectMusic true
  TimerMysteryJingle.enabled = False
End Sub

Sub SelectMusic(bResumeTrack)
  Dim nTrackChosen

  ' Highest priority goes first
  ' Bonus Count
  If bCountingBonus Then
    nTrackChosen = 14

  ' Wizard and mini-wizard modes
  ElseIf BitSetContains(vModesRunning, cModeSecret) Then
    nTrackChosen = 33
  Elseif BitSetContains(vModesRunning, cModeJuggling) Then
    nTrackChosen = 30
  Elseif BitSetContains(vModesRunning, cModeEnchantress) Then
    nTrackChosen = 29
  Elseif BitSetContains(vModesRunning, cModeChampions) Then
    nTrackChosen = 32

  ' Multiballs
  Elseif BitSetContains(vModesRunning, cModeShovel) Then
    If bResumeTrack Then nTrackChosen = 3 Else nTrackChosen = 2
  Elseif BitSetContains(vModesRunning, cModeWanderer) Then
    If anWandererSelected(nCurPlayer) = cWandererBlack Then
      nTrackChosen = 12
    Else
      nTrackChosen = 11
    End If

  ' Music when the ball is not in play (high score, bonus, shop)
  ElseIf bEnteringHighScore Then
    nTrackChosen = 31
  ElseIf bCountingBonus Then
    nTrackChosen = 14
  ElseIf bUsingShop Then
    nTrackChosen = 10

  ' If multiple modes are started, play the music for the most recently lit
  ElseIf BitSetContains(vModesRunning, cModeBlack) And anMostRecentBoss(nCurPlayer) = cModeBlack Then
    nTrackChosen = 4
  ElseIf BitSetContains(vModesRunning, cModeKing) And anMostRecentBoss(nCurPlayer) = cModeKing Then
    nTrackChosen = 6
  ElseIf BitSetContains(vModesRunning, cModeSpectre) And anMostRecentBoss(nCurPlayer) = cModeSpectre Then
    nTrackChosen = 8
  ElseIf BitSetContains(vModesRunning, cModePlague) And anMostRecentBoss(nCurPlayer) = cModePlague Then
    nTrackChosen = 16
  ElseIf BitSetContains(vModesRunning, cModeTreasure) And anMostRecentBoss(nCurPlayer) = cModeTreasure Then
    nTrackChosen = 19
  ElseIf BitSetContains(vModesRunning, cModeMole) And anMostRecentBoss(nCurPlayer) = cModeMole Then
    nTrackChosen = 34
  ElseIf BitSetContains(vModesRunning, cModeTinker) And anMostRecentBoss(nCurPlayer) = cModeTinker Then
    nTrackChosen = 22
  ElseIf BitSetContains(vModesRunning, cModePolar) And anMostRecentBoss(nCurPlayer) = cModePolar Then
    nTrackChosen = 24
  ElseIf BitSetContains(vModesRunning, cModePropeller) And anMostRecentBoss(nCurPlayer) = cModePropeller Then
    nTrackChosen = 26
  ElseIf BitSetContains(vModesRunning, cModeSuperCombo) Then
    nTrackChosen = 10

  ' Out of level music
  ElseIf abPlayingLevel(nCurPlayer) = False And nCountBossesLit = 0 Then
    If anMostRecentLevel(nCurPlayer) = cModeNone then
  ' Start of game music when ball has not been plunged
      nTrackChosen = 13
    Elseif nCountBossesPlayed > 8 Then
      nTrackChosen = 27
    Else
      nTrackChosen = 9
    End If

  ' out of mode, play the music of the level for the boss we're working toward
  ElseIf anMostRecentLevel(nCurPlayer) = cModeBlack Then
    If bResumeTrack Then nTrackChosen = 1 Else nTrackChosen = 0
  ElseIf anMostRecentLevel(nCurPlayer) = cModeKing Then
    nTrackChosen = 5
  ElseIf anMostRecentLevel(nCurPlayer) = cModeSpectre Then
    nTrackChosen = 7
  ElseIf anMostRecentLevel(nCurPlayer) = cModePlague Then
    nTrackChosen = 15
  ElseIf anMostRecentLevel(nCurPlayer) = cModeTreasure Then
    nTrackChosen = 17
  ElseIf anMostRecentLevel(nCurPlayer) = cModeMole Then
    nTrackChosen = 20
  ElseIf anMostRecentLevel(nCurPlayer) = cModeTinker Then
    nTrackChosen = 21
  ElseIf anMostRecentLevel(nCurPlayer) = cModePolar Then
    nTrackChosen = 23
  ElseIf anMostRecentLevel(nCurPlayer) = cModePropeller Then
    nTrackChosen = 25
  Elseif anMostRecentLevel(nCurPlayer) = cModeEnchantress Then
    nTrackChosen = 28
  End If

  InitMusic nTrackChosen
End Sub

Dim oFlexAlpha
const FlexDMD_RenderMode_SEG_2x16Alpha = 3

Sub InitFlexAlpha
  Dim i, AlphaNumSegs(31)
  Set oFlexAlpha = CreateObject("FlexDMD.FlexDMD")
  With oFlexAlpha
    .TableFile = Table1.Filename & ".vpx"
    .Color = RGB(255, 88, 32)
    .Width = 128
    .Height = 32
    .Clear = True
    .Run = True
    .GameName = cGameName
    .RenderMode = FlexDMD_RenderMode_SEG_2x16Alpha
  End With
  For i = 0 To 31
    AlphaNumSegs(i) = 0
  Next
  oFlexAlpha.Segments = AlphaNumSegs
End Sub

Sub InitDisplay
  Dim i
  Dim sElemName
  Dim oElem
  Dim aHide

  If bUseVR Then
    aoTopRowDigits = Array(digit001, digit002, digit003, digit004, digit005, digit006, _
      digit007, digit008, digit009, digit010, digit011, digit012, digit013, _
      digit014, digit015, digit016)
    aoBottomRowDigits = Array(digit017, digit018, digit019, digit020, digit021, _
      digit022, digit023, digit024, digit025, digit026, digit027, digit028, _
      digit029, digit030, digit031, digit032)
    For Each i in aoTopRowDigits
      i.visible = True
    Next
    For Each i in aoBottomRowDigits
      i.visible = True
    Next
  End If

  If B2SOn Or bUseDMD or bUseVR Then
    For i = 0 to 127 : aSegments(i) = 0 : Next
    aSegments(Asc("A")) = &h0877
    aSegments(Asc("B")) = &h2a0f
    aSegments(Asc("C")) = &h0039
    aSegments(Asc("D")) = &h220f
    aSegments(Asc("E")) = &h0879
    aSegments(Asc("F")) = &h0871
    aSegments(Asc("G")) = &h083d
    aSegments(Asc("H")) = &h0876
    aSegments(Asc("I")) = &h2209
    aSegments(Asc("J")) = &h001e
    aSegments(Asc("K")) = &h1470
    aSegments(Asc("L")) = &h0038
    aSegments(Asc("M")) = &h0536
    aSegments(Asc("N")) = &h1136
    aSegments(Asc("O")) = &h003f
    aSegments(Asc("P")) = &h0873
    aSegments(Asc("Q")) = &h103f
    aSegments(Asc("R")) = &h1873
    aSegments(Asc("S")) = &h086d
    aSegments(Asc("T")) = &h2201
    aSegments(Asc("U")) = &h003e
    aSegments(Asc("V")) = &h4430
    aSegments(Asc("W")) = &h5036
    aSegments(Asc("X")) = &h5500
    aSegments(Asc("Y")) = &h2500
    aSegments(Asc("Z")) = &h4409
    aSegments(Asc("0")) = &h003f
    aSegments(Asc("1")) = &h2200
    aSegments(Asc("2")) = &h085b
    aSegments(Asc("3")) = &h084f
    aSegments(Asc("4")) = &h0866
    aSegments(Asc("5")) = &h086d
    aSegments(Asc("6")) = &h087d
    aSegments(Asc("7")) = 7
    aSegments(Asc("8")) = &h087f
    aSegments(Asc("9")) = &h086f
    aSegments(39) = &h0400 ' apostrophe
    aSegments(34) = &h0202 ' double quote
    aSegments(Asc("&")) = &h135d
    aSegments(Asc("<")) = &h1400
    aSegments(Asc(">")) = &h4100
    aSegments(Asc("^")) = &h4406
    aSegments(Asc("*")) = &h7f40
    aSegments(Asc(".")) = &h0080
    aSegments(Asc("/")) = &h4400
    aSegments(Asc("\")) = &h1100
    aSegments(Asc("|")) = &h2200
    aSegments(Asc("+")) = &h2a40
    aSegments(Asc("-")) = &h0840
    aSegments(Asc("=")) = &h0848
    aSegments(Asc("_")) = 8
    aSegments(96) = &h00bf ' backtick, maps to "0,"
    aSegments(Asc("a")) = &h2280 '1,
    aSegments(Asc("b")) = &h08db '2,
    aSegments(Asc("c")) = &h08cf '3,
    aSegments(Asc("d")) = &h08e6 '4,
    aSegments(Asc("e")) = &h08ed  '5,
    aSegments(Asc("f")) = &h08fd  '6,
    aSegments(Asc("g")) = &h0087  '7,
    aSegments(Asc("h")) = &h08ff  '8,
    aSegments(Asc("i")) = &h08ef  '9,
    aSegments(Asc("j")) = &h0030 ' 20% progress
    aSegments(Asc("k")) = &h4170 ' 40% progress
    aSegments(Asc("l")) = &h6379 ' 60% progress
    aSegments(Asc("m")) = &h7779 ' 80% progress
    aSegments(Asc("n")) = &h7f7f ' 100% progress

    If bUseVR or (Not bUseBackdropDisplay) Then
      aHide = Array(EMReel1, EMReel2, EMReel3, EMReel4, EMReel5, EMReel6, EMReel7, EMReel8, _
                EMReel9, EMReel10, EMReel11, EMReel12, EMReel13, EMReel14, EMReel15, EMReel16)
      For i = 0 to 15
        aHide(i).visible = False
      Next
      aHide = Array(EMReel17, EMReel18, EMReel19, EMReel20, EMReel21, EMReel22, EMReel23, _
        EMReel24, EMReel25, EMReel26, EMReel27, EMReel28, EMReel29, EMReel30, EMReel31, EMReel32)
      For i = 0 to 15
        aHide(i).visible = False
      Next
      aHide = Array(EMReel33, EMReel34, EMReel35, EMReel36, EMReel37, EMReel38, EMReel39, _
        EMReel40, EMReel41, EMReel42, EMReel43, EMReel44, EMReel45, EMReel46, EMReel47, EMReel48)
      For i = 0 to 15
        aHide(i).visible = False
      Next
      aHide = Array(EMReel49, EMReel50, EMReel51, EMReel52, EMReel53, EMReel54, EMReel55, _
        EMReel56, EMReel57, EMReel58, EMReel59, EMReel60, EMReel61, EMReel62, EMReel63, EMReel64)
      For i = 0 to 15
        aHide(i).visible = False
      Next
      Exit Sub
    End If
  End If
  If Table1.ShowDT then
    aDisplayTopLine = Array(EMReel1, EMReel2, EMReel3, EMReel4, EMReel5, EMReel6, EMReel7, EMReel8, _
              EMReel9, EMReel10, EMReel11, EMReel12, EMReel13, EMReel14, EMReel15, EMReel16)
    aDisplayBottomLine = Array(EMReel17, EMReel18, EMReel19, EMReel20, EMReel21, EMReel22, EMReel23, _
      EMReel24, EMReel25, EMReel26, EMReel27, EMReel28, EMReel29, EMReel30, EMReel31, EMReel32)
    aHide = Array(EMReel33, EMReel34, EMReel35, EMReel36, EMReel37, EMReel38, EMReel39, _
      EMReel40, EMReel41, EMReel42, EMReel43, EMReel44, EMReel45, EMReel46, EMReel47, EMReel48)
    For i = 0 to 15
      aHide(i).visible = False
    Next
    aHide = Array(EMReel49, EMReel50, EMReel51, EMReel52, EMReel53, EMReel54, EMReel55, _
      EMReel56, EMReel57, EMReel58, EMReel59, EMReel60, EMReel61, EMReel62, EMReel63, EMReel64)
    For i = 0 to 15
      aHide(i).visible = False
    Next
  Else
    aDisplayTopLine = Array(EMReel33, EMReel34, EMReel35, EMReel36, EMReel37, EMReel38, EMReel39, _
      EMReel40, EMReel41, EMReel42, EMReel43, EMReel44, EMReel45, EMReel46, EMReel47, EMReel48)
    aDisplayBottomLine = Array(EMReel49, EMReel50, EMReel51, EMReel52, EMReel53, EMReel54, EMReel55, _
      EMReel56, EMReel57, EMReel58, EMReel59, EMReel60, EMReel61, EMReel62, EMReel63, EMReel64)
    aHide = Array(EMReel1, EMReel2, EMReel3, EMReel4, EMReel5, EMReel6, EMReel7, EMReel8, _
              EMReel9, EMReel10, EMReel11, EMReel12, EMReel13, EMReel14, EMReel15, EMReel16)
    For i = 0 to 15
      aHide(i).visible = False
    Next
    aHide = Array(EMReel17, EMReel18, EMReel19, EMReel20, EMReel21, EMReel22, EMReel23, _
      EMReel24, EMReel25, EMReel26, EMReel27, EMReel28, EMReel29, EMReel30, EMReel31, EMReel32)
    For i = 0 to 15
      aHide(i).visible = False
    Next
  End If
End Sub

Sub UpdateDisplay(sTopLine, sBottomLine)
  Dim i, j
  Dim sChar
  Dim nReelValue
  Dim aTextLines
  Dim sImage

  aTextLines = Array(sTopLine, sBottomLine)

  If bUseVR Then
    for i = 0 to 1
      ' Truncate long strings to 16 characters
      If len(aTextLines(i)) > 16 Then aTextLines(i) = Left(aTextLines(i), 16)
      ' Pad short strings to 16 characters
      If len(aTextLines(i)) < 16 Then
        aTextLines(i) = aTextLines(i) & Space(16 - len(aTextLines(i)))
      End If
      for j = 0 to 15
        sChar = Mid(aTextLines(i), j + 1, 1)
        ' Handle characters that don't have images
        if sChar = "." or sChar = ":" or sChar = ";" or sChar = "?" Then sChar = " "
        If Asc(sChar) < 32 Then sChar = " "
        If Asc(sChar) > 110 Then sChar = " "

        If Asc(sChar) < 100 Then
          sImage = "d0" & Asc(sChar)
        Else
          sImage = "d" & Asc(sChar)
        End If
        If i = 0 Then
          aoTopRowDigits(j).ImageA = sImage
        Else
          aoBottomRowDigits(j).ImageA = sImage
        End If
      next
    Next
    ' Do not use any other displays if VR is on
    Exit Sub
  End If

  If B2SOn Then
    for i = 0 to 1
      If len(aTextLines(i)) > 16 Then aTextLines(i) = Left(aTextLines(i), 16)
      for j = 0 to (len(aTextLines(i)) - 1)
        sChar = Mid(aTextLines(i), j + 1, 1)
        if Asc(sChar) > 127 Then sChar = " "
        Controller.B2SSetLED i * 16 + j + 1, aSegments(Asc(sChar))
      next
    Next
  End If

  If bUseDMD and isObject(oFlexAlpha) Then
    If Not (oFlexAlpha is Nothing) Then
      Dim AlphaNumSegs(31)
      For i = 0 To 31
        AlphaNumSegs(i) = 0
      Next
      for i = 0 to 1
        If len(aTextLines(i)) > 16 Then aTextLines(i) = Left(aTextLines(i), 16)
        for j = 0 to (len(aTextLines(i)) - 1)
          sChar = Mid(aTextLines(i), j + 1, 1)
          if Asc(sChar) > 127 Then sChar = " "
          AlphaNumSegs(i * 16 + j) = aSegments(Asc(sChar))
        next
      Next
      oFlexAlpha.Segments = AlphaNumSegs
    End If
  End If

  If bUseBackdropDisplay Then
    for i = 0 to 1
      If len(aTextLines(i)) > 16 Then aTextLines(i) = Left(aTextLines(i), 16)
      for j = 0 to (len(aTextLines(i)) - 1)
        sChar = Mid(aTextLines(i), j + 1, 1)
        ' The reels use a non-standard character map, so make these substitutions
        if sChar = "_" Then sChar = ";"
        if sChar = "|" Then sChar = "."
        if sChar = "^" Then sChar = "?"
        nReelValue = Asc(sChar) - 32
        If i = 0 Then
          aDisplayTopLine(j).setValue nReelValue
        Else
          aDisplayBottomLine(j).setValue nReelValue
        End If
      next
    Next
  End If
End Sub

Sub UpdateModeDisplay(bForceUpdate)
  Dim i, nModes
  Dim nWanderer, nWandererGoal

  ' Highest priority goes first
  ' Wizard and mini-wizard modes
  If BitSetContains(vModesRunning, cModeSecret) Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus "SECRET FIGHT", "%m", 1000
    End If
    Exit Sub
  End If
  if BitSetContains(vModesRunning, cModeJuggling) Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus "JUGGLE CHALLENGE", "%m", 1500
      QueueStatus "JUGGLING " & BIP & " BALLS", "%m", 1500
      QueueStatus "LOCK ALL BALLS", "%m", 1000
      QueueStatus "IN SAUCERS", "%m", 1000
      QueueStatus "*loop*", "", 1000
    End If
    exit sub
  End If
  if BitSetContains(vModesRunning, cModeEnchantress) Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus "ENCHANTRESS", "T=%t   %m", 1000
      QueueStatus "*loop*", "", 1000
    End If
  Elseif BitSetContains(vModesRunning, cModeChampions) Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus "HALL OF CHAMPS", "T=%t SHOTS=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If

  ' Multiballs
  Elseif BitSetContains(vModesRunning, cModeShovel) Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      if anKnightSJacks(nCurPlayer) mod 3 = 0 Then
        nSJValue = cBaseSuperJP
        sSJMultiplier = ""
      Elseif anKnightSJacks(nCurPlayer) mod 3 = 1 Then
        nSJValue = cBaseSuperJP * 2
        sSJMultiplier = "2X "
      Else
        nSJValue = cBaseSuperJP * 3
        sSJMultiplier = "3X "
      end If

      ClearStatus
      If bShovelSuperJP Then
        QueueStatus sSJMultiplier & "SUPER JACKPOT", "IS LIT", 2000
        QueueStatus "HIT PURPLE SHOT", "FOR " & WilliamsFormatNum(nSJValue), 2000
        QueueStatus "%s", "JACKPOT = " & WilliamsFormatNum(cBaseShovelJP), 1500
        QueueStatus "*loop*", "", 1500
      Else
        QueueStatus "%s", "SHOVEL MULTIBALL", 1500
        QueueStatus "%s", "HIT RED SHOT", 1500
        QueueStatus "%s", "JACKPOT = " & WilliamsFormatNum(cBaseShovelJP), 1500
        QueueStatus "%s", "HIT ORANGE SHOT", 1500
        QueueStatus "%s", "JACKPOT = " & WilliamsFormatNum(cBaseShovelJP), 1500
        QueueStatus "%s", "HIT AMBER SHOT", 1500
        QueueStatus "%s", "JACKPOT = " & WilliamsFormatNum(cBaseShovelJP), 1500
        QueueStatus "COMPLETE SHOVEL", "KNIGHT TARGETS", 2000
        QueueStatus "TO LIGHT " & WilliamsFormatNum(nSJValue), sSJMultiplier & "SUPER JACKPOT", 2000
        QueueStatus "*loop*", "", 1500
      End If
    End If
  Elseif BitSetContains(vModesRunning, cModeWanderer) Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      nWanderer = anWandererSelected(nCurPlayer)
      If nWanderer = cWandererAll Then
        nWandererGoal = 40000
      Else
        nWandererGoal = 10000
      End If

      ClearStatus
      QueueStatus asWandererNames(nWanderer), "VALUE = %m", 1500
      Select Case nWanderer
        Case cWandererBlack
          QueueStatus "HIT TARGETS", "VALUE = %m", 1500
        Case cWandererPhantom
          QueueStatus "HIT LOOPS", "VALUE = %m", 1500
        Case cWandererBaz
          QueueStatus "HIT RAMPS", "VALUE = %m", 1500
        Case cWandererReize
          QueueStatus "HIT SAUCERS", "VALUE = %m", 1500
        Case cWandererAll
          QueueStatus "HIT ANYTHING", "VALUE = %m", 1500
      End Select
      If aanWandererJPs(nCurPlayer, nWanderer) >= nWandererGoal Then
        QueueStatus "JACKPOT IS LIT", "VALUE = %m", 1500
        QueueStatus "HIT THE VUK", "VALUE = %m", 1500
      Else
        QueueStatus "INCREASE VALUE", "VALUE = %m", 1500
        QueueStatus "TO " & WilliamsFormatNum(nWandererGoal) & " TO", "VALUE = %m", 1500
        QueueStatus "LIGHT JACKPOT", "VALUE = %m", 1500
      End If
      QueueStatus "*loop*", "", 1000
    End If

  ' Music when the ball is not in play (high score, bonus, shop)
  ElseIf bEnteringHighScore Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ' TODO: Update text queue
    End If
  ElseIf bCountingBonus Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ' TODO: Update text queue
    End If

  ' If multiple modes are started, play the music for the most recently lit
  ElseIf BitSetContains(vModesRunning, cModeBlack) And _
  anMostRecentBoss(nCurPlayer) = cModeBlack Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus asKnightNames(cModeBlack), "T=%t K-N-I-G-H-T", 2000
      QueueStatus asKnightNames(cModeBlack), "T=%t TOTAL=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If
  ElseIf BitSetContains(vModesRunning, cModeKing) And _
  anMostRecentBoss(nCurPlayer) = cModeKing Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus "KING KNIGHT", "T=%t HIT SPINNER", 2000
      QueueStatus "KING KNIGHT", "T=%t TOTAL=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If
  ElseIf BitSetContains(vModesRunning, cModeSpectre) And _
  anMostRecentBoss(nCurPlayer) = cModeSpectre Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus "SPECTRE KNIGHT", "T=%t L+R RAMPS", 2000
      QueueStatus "SPECTRE KNIGHT", "T=%t TOTAL=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If
  ElseIf BitSetContains(vModesRunning, cModePlague) And _
  anMostRecentBoss(nCurPlayer) = cModePlague Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus asKnightNames(cModePlague), "T=%t HIT BUMPERS", 2000
      QueueStatus asKnightNames(cModePlague), "T=%t TOTAL=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If
  ElseIf BitSetContains(vModesRunning, cModeTreasure) And _
  anMostRecentBoss(nCurPlayer) = cModeTreasure Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus asKnightNames(cModeTreasure), "T=%t HIT SAUCERS", 2000
      QueueStatus asKnightNames(cModeTreasure), "T=%t TOTAL=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If
  ElseIf BitSetContains(vModesRunning, cModeMole) And _
  anMostRecentBoss(nCurPlayer) = cModeMole Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus asKnightNames(cModeMole), "T=%t RED TARGETS", 2000
      QueueStatus asKnightNames(cModeMole), "T=%t TOTAL=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If
  ElseIf BitSetContains(vModesRunning, cModeTinker) And _
  anMostRecentBoss(nCurPlayer) = cModeTinker Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus asKnightNames(cModeTinker), "T=%t HIT CAPTIVE", 2000
      QueueStatus asKnightNames(cModeTinker), "T=%t TOTAL=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If
  ElseIf BitSetContains(vModesRunning, cModePolar) And _
  anMostRecentBoss(nCurPlayer) = cModePolar Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus asKnightNames(cModePolar), "T=%t S-H-O-V-E-L", 2000
      QueueStatus asKnightNames(cModePolar), "T=%t TOTAL=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If
  ElseIf BitSetContains(vModesRunning, cModePropeller) And _
  anMostRecentBoss(nCurPlayer) = cModePropeller Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus asKnightNames(cModePropeller), "T=%t CENTER RAMP", 2000
      QueueStatus asKnightNames(cModePropeller), "T=%t TOTAL=%m", 1000
      QueueStatus "*loop*", "", 1000
    End If

  ' golden awards
  ElseIf anComboTimers(cShotLoopL) > 1 Or _
  True = BitSetContains(vModesRunning, cModeSuperCombo) Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus

      If BitSetContains(vModesRunning, cModeSuperCombo) Then
        QueueStatus "3X COMBO VALUES", "FOR %t SECONDS", 2000
      End If

      If anComboTimers(cShotLoopL) > 1 Then
        If (anModeTimeLeft(cModeSuperCombo) > 0 And nComboJPCurrent < 28800) _
        Or (anModeTimeLeft(cModeSuperCombo) < 1 And nComboJPCurrent < 9600) Then
          QueueStatus "LEFT LOOP AWARDS", "COMBO AT %m", 1000
        Else
          QueueStatus "LEFT LOOP AWARDS", "MAX COMBO %m", 1000
        End If
      End If

      QueueStatus "*loop*", "", 1000
    End If
  ElseIf BitSetContains(vModesRunning, cModeGoldTargets) Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus "GOLDEN TARGETS", "%t SECONDS", 1500
      QueueStatus "SHOOT TARGETS", "%t SECONDS", 1500
      QueueStatus "TO ADD MORE GOLD", "%t SECONDS", 1500
      QueueStatus "*loop*", "", 1500
    End If
  ElseIf BitSetContains(vModesRunning, cModeGoldRamps) Then
    If bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
      ClearStatus
      QueueStatus "GOLDEN RAMPS", "%t SECONDS", 1500
      QueueStatus "RAMPS ADD GOLD", "%t SECONDS", 1500
      QueueStatus "*loop*", "", 1500
    End If
  ElseIf bForceUpdate Or (vPrevModesRunning <> vModesRunning) Then
    ClearStatus
  End If
End Sub

'*** Substitutions to make in display text:
'  %s : current player's (formatted) score
'    %p : number of the current player
'    %m : value of the sModeProgress variable - set in UpdateModeDisplay
'  %t : value of the nModeTime variable
' "*loop*" : If the top row is excatly "*loop*", loop the message queue

Function FormatText(sText)
  Dim sReplacedText
  Dim nWizardScore
  Dim nShotCount
  Dim i

  nModeTime = 0
  If IsArray(anModeTimeLeft) And nPlayers > 0 Then
    If BitSetContains(vModesRunning, cModeSecret) Then
      sModeProgress = WilliamsFormatNum(anModeScore(cModeSecret)) & _
        "/" & WilliamsFormatNum(500000)
    ElseIf BitSetContains(vModesRunning, cModeJuggling) Then
      sModeProgress = ""
      for each i in anJugglingShots
        If anJuggleTimers(i) > 0 Then
          nModeTime = (CInt(anJuggleTimers(i) / 1000) - 2)
          If nModeTime < 0 Then nModeTime = 0
          sModeProgress = sModeProgress & nModeTime
        Else
          sModeProgress = sModeProgress & "--"
        End If
        If i <> cShotSaucerR Then sModeProgress = sModeProgress & " "
      Next
    ElseIf BitSetContains(vModesRunning, cModeChampions) Then
      nModeTime = CInt(anModeTimeLeft(cModeChampions) / 1000) - 2
      If nModeTime < 0 Then nModeTime = 0
      nShotCount = 0
      For i = 0 to 9
        If BitSetContains(vMultipliersLit, i) Then
          nShotCount = nShotCount + 1
        End If
      Next
      sModeProgress = CStr(nShotCount)
    ElseIf BitSetContains(vModesRunning, cModeWanderer) Then
      sModeProgress = CStr(aanWandererJPs(nCurPlayer, anWandererSelected(nCurPlayer)))
    ElseIf False = bCanStartBoss And anMostRecentBoss(nCurPlayer) <> cModeNone Then
      ' Instead of a grace period, show less time than we actually have
      nModeTime = CInt(anModeTimeLeft(anMostRecentBoss(nCurPlayer)) / 1000) - 2
      If nModeTime < 0 Then nModeTime = 0
      If anMostRecentBoss(nCurPlayer) = cModeEnchantress Then
        nWizardScore = 0
        for i = cModeBlack to cModePropeller
          nWizardScore = nWizardScore + anModeScore(i)
        Next
        sModeProgress = WilliamsFormatNum(nWizardScore)
      Else
        sModeProgress = WilliamsFormatNum(anModeScore(anMostRecentBoss(nCurPlayer)))
      End If
    ElseIf anComboTimers(cShotLoopL) > 1 Or _
    True = BitSetContains(vModesRunning, cModeSuperCombo) Then
      nModeTime = CInt(anModeTimeLeft(cModeSuperCombo) / 1000) - 2
      If nModeTime < 0 Then nModeTime = 0
      sModeProgress = WilliamsFormatNum(nComboJPCurrent)
    End If
  End If
  sReplacedText = sText
  sReplacedText = Replace(sReplacedText, "%s", WilliamsFormatNum(anScore(nCurPlayer)), vbTextCompare)
  sReplacedText = Replace(sReplacedText, "%p", CStr(nCurPlayer + 1), vbTextCompare)
  sReplacedText = Replace(sReplacedText, "%m", sModeProgress, vbTextCompare)
  sReplacedText = Replace(sReplacedText, "%t", CStr(nModeTime), vbTextCompare)
  FormatText = sReplacedText
End Function

Sub ShowText(sTopRow, sBottomRow, nDuration, nPriority)
  if nPriority <= nTextPriority or nTextDuration = 0 then
    sTextTop = sTopRow
    sTextBottom = sBottomRow
    nTextDuration = nDuration
    nTextPriority = nPriority
  end if
End Sub

Sub QueueText(sTopRow, sBottomRow, nDuration)
  oQueueTextTop.Enqueue sTopRow
  oQueueTextBottom.Enqueue sBottomRow
  oQueueTextDuration.Enqueue nDuration
End Sub

Sub QueueTextFirst(sTopRow, sBottomRow, nDuration)
  oQueueTextTop.PushFirst sTopRow
  oQueueTextBottom.PushFirst sBottomRow
  oQueueTextDuration.PushFirst nDuration
End Sub

Sub ClearTextQueue
  oQueueTextTop.Clear
  oQueueTextBottom.Clear
  oQueueTextDuration.Clear
End Sub

Sub QueueStatus(sTopRow, sBottomRow, nDuration)
  oQueueStatusTop.Enqueue sTopRow
  oQueueStatusBottom.Enqueue sBottomRow
  oQueueStatusDuration.Enqueue nDuration
End Sub

Sub ClearStatus
  oQueueStatusTop.Clear
  oQueueStatusBottom.Clear
  oQueueStatusDuration.Clear
End Sub

Sub TimerDisplay_Timer
  dim sTopRow
  dim sBotRow
  dim sScore
  dim sBallText
  dim sReplacedTop, sReplacedBottom
  dim i

  nTimeDisplay = nTimeDisplay + TimerDisplay.interval

  If isObject(oQueueTextTop) Then
    If oQueueTextTop.Size > 0 And nTextDuration = 0 Then
      nTextPriority = 98 'Queued messages are second lowest priority
      sTextTop = oQueueTextTop.Dequeue
      ' handle looping queue
      If sTextTop = "*loop*" Then
        oQueueTextTop.Rewind
        oQueueTextBottom.Rewind
        oQueueTextDuration.Rewind
        sTextTop = oQueueTextTop.Dequeue
      End If

      sTextBottom = oQueueTextBottom.Dequeue
      nTextDuration = oQueueTextDuration.Dequeue

      If oQueueTextTop.Size = 0 Then
        ClearTextQueue
      End If
    Elseif oQueueStatusTop.Size > 0 And nTextDuration = 0 Then
      nTextPriority = 99 'Queued status is lowest priority
      sTextTop = oQueueStatusTop.Dequeue
      ' handle looping queue
      If sTextTop = "*loop*" Then
        oQueueStatusTop.Rewind
        oQueueStatusBottom.Rewind
        oQueueStatusDuration.Rewind
        sTextTop = oQueueStatusTop.Dequeue
      End If

      sTextBottom = oQueueStatusBottom.Dequeue
      nTextDuration = oQueueStatusDuration.Dequeue

      If oQueueStatusTop.Size = 0 Then
        ClearStatus
      End If
    end if
  End If

  sReplacedTop = FormatText(sTextTop)
  sReplacedBottom = FormatText(sTextBottom)

  if nTextDuration <> 0 then
    UpdateDisplay CenterTextOnDisplay(sReplacedTop), CenterTextOnDisplay(sReplacedBottom)
    if nTextDuration <> PlayUntilHit then
      nTextDuration = nTextDuration - TimerDisplay.interval
      If nTextDuration < 0 Then nTextDuration = 0
    End If
    Exit Sub
  End If

  If bDrainingBalls and nTiltWarnings > 2 Then
    if nTimeDisplay < 1251 then
      sTopRow = Pad(16)
      sBotRow = CenterTextOnDisplay("TILT")
    elseif nTimeDisplay < 2501 then
      sTopRow = CenterTextOnDisplay("TILT")
      sBotRow = Pad(16)
    else
      nTimeDisplay = 0
    end if
    UpdateDisplay sTopRow, sBotRow
    Exit Sub
  End If

  ' Instant info
  If nPlayers > 0 And nTimeUpperLeftFlip = -1 And nTimeUpperRightFlip = -1 Then
    bShowingInstantInfo = False
  End If

  Dim nBadMedals: nBadMedals = 0
  Dim nBronzeMedals : nBronzeMedals = 0
  Dim nSilverMedals : nSilverMedals = 0
  Dim nGoldMedals : nGoldMedals = 0
  Dim nLocksLitCount : nLocksLitCount = 0
  Dim nXUsed : nXused = 0
  Dim nPoints, nNextComboAward, temp

  If nPlayers > 0 And nTimeLastSwitch > 0 And (GameTime - nTimeLastSwitch > 2000) Then
    If ((nTimeUpperLeftFlip > 0) And (GameTime - nTimeUpperLeftFlip > 2000)) Or _
      ((nTimeUpperRightFlip > 0) And (GameTime - nTimeUpperRightFlip > 2000)) Then
      If bShowingInstantInfo = False Then nTimeDisplay = 0
      bShowingInstantInfo = True
      If nTimeDisplay > 2000 Then
        nInstantInfoPage = (nInstantInfoPage + 1) mod (cLastInfoPage + 1)
        nTimeDisplay = 0
      End If
      ' Count medals and lock to display on instant info
      for i = cModeBlack to cModePropeller
        Select Case aanBossMedals(nCurPlayer, i)
          Case cMedalConsolation
            nBadMedals = nBadMedals + 1
          Case cMedalBronze
            nBronzeMedals = nBronzeMedals + 1
          Case cMedalSilver
            nSilverMedals = nSilverMedals + 1
          Case cMedalGold
            nGoldMedals = nGoldMedals + 1
        End Select
      Next

      sTopRow = CenterTextOnDisplay("INSTANT INFO")
      Select Case nInstantInfoPage
        Case 0
          If bCanProgressLevel = False And bCanStartBoss = True Then
            temp = "START A BOSS TO"
          ElseIf abPlayingLevel(nCurPlayer) Then
            temp = "LEVEL PROGRESS"
          Else
            temp = "NEXT LEVEL"
          End If
        Case 1
          If bCanProgressLevel = False And bCanStartBoss = True Then
            temp = "START NEXT LEVEL"
          ElseIf abPlayingLevel(nCurPlayer) Then
            temp = anLevelShotsMade(nCurPlayer) & " OF " & _
            anLevelShotsToBeat(nCurPlayer) & " SHOTS"
          Else
            temp = anLevelShotsMade(nCurPlayer) & " OF " & _
            anLevelShotsToBeat(nCurPlayer) & " SHOTS"
          End If
        Case 2
          for i = 0 to 9
            If BitSetContains(avXUsed(nCurPlayer), i) Then
              nXUsed = nXUsed + 1
            End If
          Next
          temp = nXUsed & " SHOT 2X USED"
        Case 3
          If nCountBossesLit = 1 Then
            temp = nCountBossesLit & " BOSS LIT"
          Else
            temp = nCountBossesLit & " BOSSES LIT"
          End If
        Case 4
          If nCountBossesPlayed = 1 Then
            temp = nCountBossesPlayed & " BOSS FOUGHT"
          Else
            temp = nCountBossesPlayed & " BOSSES FOUGHT"
          End If
        Case 5
          temp = nBadMedals & " " & Chr(34) & "YOU TRIED" & Chr(34)
        Case 6
          If nBronzeMedals = 1 Then
            temp = nBronzeMedals & " BRONZE MEDAL"
          Else
            temp = nBronzeMedals & " BRONZE MEDALS"
          End If
        Case 7
          If nSilverMedals = 1 Then
            temp = nSilverMedals & " SILVER MEDAL"
          Else
            temp = nSilverMedals & " SILVER MEDALS"
          End If
        Case 8
          If nGoldMedals = 1 Then
            temp = nGoldMedals & " GOLD MEDAL"
          Else
            temp = nGoldMedals & " GOLD MEDALS"
          End If
        Case 9
          If anKnightMBalls(nCurPlayer) = 1 Then
            temp = anKnightMBalls(nCurPlayer) & " MULTIBALL"
          Else
            temp = anKnightMBalls(nCurPlayer) & " MULTIBALLS"
          End If
        Case 10
          If BitSetContains(avLocksLit(nCurPlayer), cLockLeft) = True And _
          BitSetContains(anLocksMade(nCurPlayer), cLockLeft) = False Then
            nLocksLitCount = nLocksLitCount + 1
          End If
          If BitSetContains(avLocksLit(nCurPlayer), cLockRight) = True And _
          BitSetContains(anLocksMade(nCurPlayer), cLockRight) = False Then
            nLocksLitCount = nLocksLitCount + 1
          End If
          If nLocksLitCount = 1 Then
            temp = nLocksLitCount & " LOCK LIT"
          Else
            temp = nLocksLitCount & " LOCKS LIT"
          End If
        Case 11
          If CountLocksMade = 1 Then
            temp = CountLocksMade & " BALL LOCKED"
          Else
            temp = CountLocksMade & " BALLS LOCKED"
          End If
        Case 12
          If nExtraBalls = 1 Then
            temp = nExtraBalls & " EXTRA BALL"
          Else
            temp = nExtraBalls & " EXTRA BALLS"
          End If
        Case 13
          If anComboCount(nCurPlayer) = 1 Then
            temp = anComboCount(nCurPlayer) & " COMBO"
          Else
            temp = anComboCount(nCurPlayer) & " COMBOS"
          End If
        Case 14
          nNextComboAward = ((anComboCount(nCurPlayer) + 4) \ 4) * 4
          Select Case ((anComboCount(nCurPlayer) \ 4) Mod 3)
            Case 0 ' 4, 16, 28, 40...
              nPoints = 5000 + ((anComboCount(nCurPlayer) \ 12) * 2500)
              temp = WilliamsFormatNum(nPoints) & " AT " & nNextComboAward
            Case 1 ' 4, 16, 28, 40...
              temp = "MYSTERY AT " & nNextComboAward
            Case 2 ' 8, 20, 32, 44
              temp = "3X COMBO AT " & nNextComboAward
          End Select
        Case 15
          temp = "COMBO SUM=" & WilliamsFormatNum(anComboJPSum(nCurPlayer))
        Case 16
          If anComboJPSum(nCurPlayer) < cJugglingGoal Then
            temp = "JUGGLE AT " & WilliamsFormatNum(cJugglingGoal)
          Else
            temp = "JUGGLING IS LIT"
          End If
        Case 17
          If anWandererHits(nCurPlayer) > 0 Then
            temp = anWandererHits(nCurPlayer) & " POPS / SPINS"
          Else
            temp = "WANDERER IS LIT"
          End If
        Case 18
          If anWandererHits(nCurPlayer) > 0 Then
            temp = "TO LITE WANDERER"
          Else
            temp = "SHOOT UPPER HOLE"
          End If
        Case 19
          temp = "BONUS " & nBonusX & "X"
        Case 20
          temp = "1ST " & HighScoreName(0) & " " & WilliamsFormatNum(HighScore(0))
        Case 21
          temp = "2ND " & HighScoreName(1) & " " & WilliamsFormatNum(HighScore(1))
        Case 22
          temp = "3RD " & HighScoreName(2) & " " & WilliamsFormatNum(HighScore(2))
        Case 23
          temp = "4TH " & HighScoreName(3) & " " & WilliamsFormatNum(HighScore(3))
        Case 24
          temp = "SECRET FIGHT " & BitSetCount(avSecretProgress(nCurPlayer), 6) & "/5"
        Case 25
          If BitSetContains(avSecretProgress(nCurPlayer), cSecret3xSuper) Then
            temp = "TRIPLE SUPER YES"
          Else
            temp = "TRIPLE SUPER NO"
          End If
        Case 26
          If BitSetContains(avSecretProgress(nCurPlayer), cSecretChampions) Then
            temp = "HALL/CHAMPS YES"
          Else
            temp = "HALL/CHAMPS NO"
          End If
        Case 27
          If BitSetContains(avSecretProgress(nCurPlayer), cSecretWanderer) Then
            temp = "WANDERER SJ YES"
          Else
            temp = "WANDERER SJ NO"
          End If
        Case 28
          If BitSetContains(avSecretProgress(nCurPlayer), cSecretEnchantress) Then
            temp = "125K ENCHANT YES"
          Else
            temp = "125K ENCHANT NO"
          End If
        Case 29
          If BitSetContains(avSecretProgress(nCurPlayer), cSecretJuggling) Then
            temp = "JUGGLING SJ YES"
          Else
            temp = "JUGGLING SJ NO"
          End If
      End Select
      sBotRow = CenterTextOnDisplay(temp)
      UpdateDisplay sTopRow, sBotRow
      Exit Sub
    End If
  End If

  if nPlayers = 0 then
    if nTimeDisplay < 2001 then
      UpdateDisplay CenterTextOnDisplay("SHOVEL KNIGHT"), _
        CenterTextOnDisplay("FREE PLAY")
    elseif nTimeDisplay < 4001 then
      UpdateDisplay CenterTextOnDisplay("SHOVEL KNIGHT"), _
        CenterTextOnDisplay("PRESS START")
    elseif nTimeDisplay < 5501 then
      UpdateDisplay CenterTextOnDisplay("LAYOUT AND CODE"), _
        CenterTextOnDisplay("BY WIZBALL")
    elseif nTimeDisplay < 7001 then
      UpdateDisplay CenterTextOnDisplay("SOME ELEMENTS"), _
        CenterTextOnDisplay("BY JPSALAS")
    elseif nTimeDisplay < 8501 then
      UpdateDisplay CenterTextOnDisplay("SHOVEL KNIGHT BY"), _
        CenterTextOnDisplay("YACHT CLUB GAMES")
    elseif nTimeDisplay < 10001 then
      temp = WilliamsFormatNum(HighScore(0)) & " " & HighScoreName(0)
      UpdateDisplay CenterTextOnDisplay("GRAND CHAMPION"), _
        CenterTextOnDisplay(temp)
    elseif nTimeDisplay < 11501 then
      temp = WilliamsFormatNum(HighScore(1)) & " " & HighScoreName(1)
      UpdateDisplay CenterTextOnDisplay("HIGH SCORE 1"), _
        CenterTextOnDisplay(temp)
    elseif nTimeDisplay < 13001 then
      temp = WilliamsFormatNum(HighScore(2)) & " " & HighScoreName(2)
      UpdateDisplay CenterTextOnDisplay("HIGH SCORE 2"), _
        CenterTextOnDisplay(temp)
    elseif nTimeDisplay < 14501 then
      temp = WilliamsFormatNum(HighScore(3)) & " " & HighScoreName(3)
      UpdateDisplay CenterTextOnDisplay("HIGH SCORE 3"), _
        CenterTextOnDisplay(temp)
    elseIf nTimeDisplay < 16501 Then
      Select Case nPlayersLastGame
        Case 1
          sTopRow = getFormattedScore(0) & Pad(8)
          sBotRow = Pad(16)
        Case 2
          sTopRow = getFormattedScore(0) & getFormattedScore(1)
          sBotRow = Pad(16)
        Case 3
          sTopRow = getFormattedScore(0) & getFormattedScore(1)
          sBotRow = getFormattedScore(2) & Pad(8)
        Case Else
          sTopRow = getFormattedScore(0) & getFormattedScore(1)
          sBotRow = getFormattedScore(2) & getFormattedScore(3)
      End Select
      UpdateDisplay sTopRow, sBotRow
    Else
      nTimeDisplay = 0
    end if
  End If
  If nPlayers = 1 then
    sScore = blinkScore(nTimeDisplay, nCurPlayer)
    if len(sScore) < 8 then
      sTopRow = Pad(7 - len(sScore)) & sScore & Pad(9)
    else
      sTopRow = sScore & Pad(16 - len(sScore))
    end if

    if nBall < 10 then
      sBotRow = Pad(10) & "BALL " & cStr(nBall)
    else
      sBotRow = Pad(10) & "BALL" & cStr(nBall)
    end if

    UpdateDisplay sTopRow, sBotRow

    if nTimeDisplay >= 2160 then nTimeDisplay = 0
  Elseif nPlayers = 2 Then
    If nCurPlayer = 0 Then
      sTopRow = blinkScore(nTimeDisplay, 0) & getFormattedScore(1)
    Else
      sTopRow = getFormattedScore(0) & blinkScore(nTimeDisplay, 1)
    End If

    if nBall < 10 then
      sBotRow = Pad(10) & "BALL " & cStr(nBall)
    else
      sBotRow = Pad(10) & "BALL" & cStr(nBall)
    end if

    UpdateDisplay sTopRow, sBotRow

    if nTimeDisplay >= 2160 then nTimeDisplay = 0
  Elseif nPlayers = 3 Then
    if nBall < 10 then
      sBallText = Pad(2) & "BALL " & cStr(nBall)
    else
      sBallText = Pad(1) & "BALL " & cStr(nBall)
    end if

    If nCurPlayer = 0 Then
      sTopRow = blinkScore(nTimeDisplay, 0) & getFormattedScore(1)
      sBotRow = getFormattedScore(2) & sBallText
    ElseIf nCurPlayer = 1 Then
      sTopRow = getFormattedScore(0) & blinkScore(nTimeDisplay, 1)
      sBotRow = getFormattedScore(2) & sBallText
    Else
      sTopRow = getFormattedScore(0) & getFormattedScore(1)
      sBotRow = blinkScore(nTimeDisplay, 2) & sBallText
    End If

    UpdateDisplay sTopRow, sBotRow

    if nTimeDisplay >= 2160 then nTimeDisplay = 0
  Elseif nPlayers = 4 Then
    If nCurPlayer = 0 Then
      sTopRow = blinkScore(nTimeDisplay, 0) & getFormattedScore(1)
      sBotRow = getFormattedScore(2) & getFormattedScore(3)
    ElseIf nCurPlayer = 1 Then
      sTopRow = getFormattedScore(0) & blinkScore(nTimeDisplay, 1)
      sBotRow = getFormattedScore(2) & getFormattedScore(3)
    ElseIf nCurPlayer = 2 Then
      sTopRow = getFormattedScore(0) & getFormattedScore(1)
      sBotRow = blinkScore(nTimeDisplay, 2) & getFormattedScore(3)
    Else
      sTopRow = getFormattedScore(0) & getFormattedScore(1)
      sBotRow = getFormattedScore(2) & blinkScore(nTimeDisplay, 3)
    End If

    UpdateDisplay sTopRow, sBotRow

    If nScoreBlinks < 3 And nPlayers = 4 Then
      if nTimeDisplay >= 2160 Then
        nTimeDisplay = 0
        nScoreBlinks = nScoreBlinks + 1
      end if
    Else
      if nTimeDisplay >= 3200 And nPlayers = 4 then
        nTimeDisplay = 0
        nScoreBlinks = 0
      End If
    End If
  end if
End Sub


Function getFormattedScore(nPlayer)
  dim sScore
  sScore = WilliamsFormatNum(anScore(nPlayer))
  If len(sScore) < 8 Then
    sScore =  Pad(7 - len(sScore)) & sScore & " "
  End If

  getFormattedScore = sScore
End Function

Function blinkScore(nFrameNo, nPlayer)
  dim sBeforeBlank
  dim sAfterBlank
  dim sScore
  dim nFigureToBlink

  sScore = getFormattedScore(nPlayer)

  if nTimeDisplay > 999 And nPlayers > 0 Then
    If nTimeDisplay < 1601 Then
      nFigureToBlink = (nTimeDisplay - 1000) \ 80
      sBeforeBlank = Left(sScore, 7 - nFigureToBlink)
      sAfterBlank = Right(sScore, nFigureToBlink)
      sScore = sBeforeBlank & " " & sAfterBlank
    ElseIf nTimeDisplay < 2161 Then
      nFigureToBlink = (nTimeDisplay - 1600) \ 80
      sBeforeBlank = Left(sScore, nFigureToBlink)
      sAfterBlank = Right(sScore, 7 - nFigureToBlink)
      sScore = sBeforeBlank & " " & sAfterBlank
    ElseIf nTimeDisplay < 3201 Then ' Only show ball nr if 4 players
      if nBall < 10 then
        sScore = Pad(2) & "BALL " & cStr(nBall)
      else
        sScore = Pad(1) & "BALL " & cStr(nBall)
      end if
    End If
  End If

  blinkScore = sScore
End Function

' Format an integer as a string, using the special characters in the display
' reels to display a number and a comma in the same character. The
' "0," to "9," special characters are at code points 96 - 105.
Function WilliamsFormatNum(nNumber)
  Dim i
  dim sNum
  dim sPrefix
  dim sPostfix
  dim sDigitWithComma
  dim nLen
  if not IsNumeric(nNumber) then
    WilliamsFormatNum = "NAN"
    exit function
  end if
  WilliamsFormatNum = cStr(nNumber)
  nLen = len(WilliamsFormatNum)

  if nLen < 4 then exit function
  for i = nLen to 3 step -1
    if i mod 3 = 1 then
      sPrefix = left(WilliamsFormatNum, nLen - i)
      sPostfix = right(WilliamsFormatNum, i - 1)
      sDigitWithComma = chr(asc(mid(WilliamsFormatNum, nLen + 1 - i, 1)) + 48)
      WilliamsFormatNum = sPrefix & sDigitWithComma & sPostfix
    end if
  next
End Function

Function CenterTextOnDisplay(sText)
  dim nLen
  nLen = len(sText)
  if nLen > 16 then
    CenterTextOnDisplay = Left(sText, 16)
  Else
    CenterTextOnDisplay = Pad((16 - nLen) \ 2) & sText &_
                Pad((17 - nLen) \ 2)
  end if
End Function

Function Pad(nChars)
  if True Then
    Pad = Space(nChars)
  Else
    Pad = String(nChars, "*")
  End If
End Function

'*** Functions for color LEDs for shots

Sub AddShotColor(nShot, nColor)
  avShotColors(nShot) = BitSetAdd(avShotColors(nShot), nColor)
End Sub

Sub RemoveShotColor(nShot, nColor)
  avShotColors(nShot) = BitSetRemove(avShotColors(nShot), nColor)
End Sub

Sub UpdateLightsMultipliers
  Dim i
  for i = 0 to 9
    If BitSetContains(vModesRunning, cModeChampions) Then
      aoXLights(i).color = Color_Gray
      aoXLights(i).colorFull = vbWhite
      aoXLights(i).state = LightStateOn
    Else
      If BitSetContains(avXUsed(nCurPlayer), i) Then
        aoXLights(i).color = vbRed
        aoXLights(i).colorFull = Color_LtRed
        aoXLights(i).state = LightStateOn
      Else
        aoXLights(i).color = Color_Gray
        aoXLights(i).colorFull = vbWhite
      End If

      If aanXTimers(nCurPlayer, i) > 0 And aoXLights(i).state = LightStateOff Then
        aoXLights(i).state = LightStateOn
      ElseIf Not BitSetContains(avXUsed(nCurPlayer), i) Then
        aoXLights(i).state = LightStateOff
      End If
    End If
  Next
End Sub

' Update lights for locks and orange lights during juggling challenge

Sub UpdateLightsLocks
  Dim i
  For Each i in Array(cShotRampL, cShotRampC, cShotSaucerR)
    RemoveShotColor i, cBlinkGreen
    anShotIdleColor(i) = -1
  Next
  ' Don't display lock lights during sold out, secret, or enchantress wizard modes
  ' FIXME: secret mode
  If BitSetContains(vModesRunning, cModeJuggling) Then
    for each i in anJugglingShots
      if anJuggleTimers(i) > 0 Then
        RemoveShotColor i, cBlinkPurple
        anShotIdleColor(i) = cBlinkPurple
      Else
        AddShotColor i, cBlinkPurple
      end If
    Next
  ElseIf BitSetContains(vModesRunning, cModeEnchantress) Then
    ' Keep green blinks disabled
  ElseIf BitSetContains(vModesRunning, cModeSecret) Then
    ' Keep green blinks disabled
  Else
    If BitSetIsFull(anLocksMade(nCurPlayer), 2) Then
      anShotIdleColor(cShotRampL) = cBlinkGreen
      anShotIdleColor(cShotRampC) = cBlinkGreen
      AddShotColor cShotSaucerR, cBlinkGreen
    ElseIf BitSetContains(anLocksMade(nCurPlayer), cLockLeft) Then
      anShotIdleColor(cShotRampL) = cBlinkGreen
      If BitSetContains(avLocksLit(nCurPlayer), cLockRight) Then
        AddShotColor cShotRampC, cBlinkGreen
      End If
    ElseIf BitSetContains(anLocksMade(nCurPlayer), cLockRight) Then
      anShotIdleColor(cShotRampC) = cBlinkGreen
      If BitSetContains(avLocksLit(nCurPlayer), cLockLeft) Then
        AddShotColor cShotRampL, cBlinkGreen
      End If
    Else
      If BitSetContains(avLocksLit(nCurPlayer), cLockRight) Then
        AddShotColor cShotRampC, cBlinkGreen
      End If
      If BitSetContains(avLocksLit(nCurPlayer), cLockLeft) Then
        AddShotColor cShotRampL, cBlinkGreen
      End If
    End If

    ' If enchantress wizard mode is lit at the right saucer, shooting
    ' it will not start a multiball
    If BitSetContains(vModesRunning, cModeEnchantress) _
    Or BitSetContains(avBossesLit(nCurPlayer), cModeEnchantress) Then
      RemoveShotColor cShotSaucerR, cBlinkGreen
    End If
  End If
End Sub

Sub UpdateLightsWanderer
  Dim i
  Dim nWanderer, nWandererGoal

  nWanderer = anWandererSelected(nCurPlayer)
  If nWanderer = cWandererAll Then
    nWandererGoal = 40000
  Else
    nWandererGoal = 10000
  End If

  for i = 0 to 3
    If BitSetContains(avWanderersBeaten(nCurPlayer), i) Then
      aoWandererLights(i).state = LightStateOn
    Else
      aoWandererLights(i).state = LightStateOff
    End If
  Next

  If anWandererSelected(nCurPlayer) = cWandererAll Then
    for i = 0 to 3
      aoWandererLights(i).state = LightStateBlinking
    Next
  Else
    aoWandererLights(nWanderer).state = LightStateBlinking
  End If

  If False = BitSetContains(vModesRunning, cModeEnchantress) _
  And False = BitSetContains(vModesRunning, cModeSecret) _
  And False = BitSetContains(vModesRunning, cModeJuggling) _
  And False = BitSetContains(vModesRunning, cModeWanderer) _
  And anWandererHits(nCurPlayer) < 1 Then
    AddShotColor cShotUpperHole, cBlinkCyan
  End if

  If BitSetContains(vModesRunning, cModeWanderer) Then
    If aanWandererJPs(nCurPlayer, nWanderer) >= nWandererGoal Then
      AddShotColor cShotVUK, cBlinkCyan
    Else
      RemoveShotColor cShotVUK, cBlinkCyan
    End If
  Else
    RemoveShotColor cShotVUK, cBlinkCyan
  End If

  If anExtraBallsLit(nCurPlayer) > 0 Then
    AddShotColor cShotUpperHole, cBlinkMagenta
    LightXBallLit.state = LightStateBlinking
  Else
    RemoveShotColor cShotUpperHole, cBlinkMagenta
    LightXBallLit.state = LightStateOff
  End If
End Sub

' Also update white light for Hall of Champions mode and pink mystery light
Sub UpdateLightsBoss
  Dim nShotCount
  Dim i

  nShotCount = 0
  For i = 0 to 9
    If aanXTimers(nCurPlayer, i) > 0 _
    Or true = BitSetContains(avXUsed(nCurPlayer), i) Then
      nShotCount = nShotCount + 1
    End If
  Next
  If 10 = nShotCount And _
  False = BitSetContains(vModesRunning, cModeChampions) Then
    If False = BitSetContains(avBossesLit(nCurPlayer), cModeSecret) Then
      AddShotColor cShotSaucerR, cBlinkWhite
    End If
  Elseif BitSetContains(vMultipliersLit, cShotSaucerR) Then
    AddShotColor cShotSaucerR, cBlinkWhite
  Else
    RemoveShotColor cShotSaucerR, cBlinkWhite
  End If

  If BitSetContains(vModesRunning, cModeTreasure) Then
    AddShotColor cShotSaucerR, cBlinkBlue
  ElseIf Not bCanStartBoss Then
    RemoveShotColor cShotSaucerR, cBlinkBlue
  ElseIf nCountBossesLit > 0 Then
    AddShotColor cShotSaucerR, cBlinkBlue
  Else
    RemoveShotColor cShotSaucerR, cBlinkBlue
  End If

  If anTrouppleLit(nCurPlayer) > 0 Then
    AddShotColor cShotShop, cBlinkPink
  Else
    RemoveShotColor cShotShop, cBlinkPink
  End If

  If True = bCanStartJuggling Then
    AddShotColor cShotSaucerR, cBlinkPurple
  End If

  If True = bCanStartSecret Then
    AddShotColor cShotSaucerR, cBlinkOrange
    avBossesLit(nCurPlayer) = BitSetAdd(avBossesLit(nCurPlayer), cModeSecret)
  Else
    RemoveShotColor cShotSaucerR, cBlinkOrange
  End If

End Sub

Sub LightTimer_Timer
  Dim i, j
  Dim aActiveColors(9)
  Dim nColorCount
  Dim nColorToUse
  If nPlayers = 0 Or True = bCountingBonus then exit sub ' Don't update colors during attract mode
  for i = 0 to 9 ' shots
    aoShotLights(i).state = LightStateOff
    nColorCount = 0
    for j = 0 to nColorsInCycle - 1 ' colors
      if BitSetContains(avShotColors(i), j) Then
        aActiveColors(nColorCount) = j
        nColorCount = nColorCount + 1
      end If
    next
    If nColorCount > 0 Then
      nColorToUse = nColorCycleIndex mod nColorCount
      aoShotLights(i).color = anColorCycleDark(aActiveColors(nColorToUse))
      aoShotLights(i).colorFull = anColorCycleLight(aActiveColors(nColorToUse))
      aoShotLights(i).state = LightStateBlinking
    else
      if anShotIdleColor(i) = -1 Then
        aoShotLights(i).color = Color_Gray
        aoShotLights(i).colorFull = vbWhite
        ' Keep the light state to off
      Else
        aoShotLights(i).color = anColorCycleDark(anShotIdleColor(i))
        aoShotLights(i).colorFull = anColorCycleLight(anShotIdleColor(i))
        aoShotLights(i).state = LightStateOn
      end If
    end if

    ' Also update shot multiplier lights so the blinks are synced
    If BitSetContains(vModesRunning, cModeChampions) Then
      aoXLights(i).state = LightStateOff

      If anModeTimeLeft(cModeChampions) > 7000 Then
        aoXLights(i).BlinkInterval = 300
      ElseIf anModeTimeLeft(cModeChampions) > 4000 Then
        aoXLights(i).BlinkInterval = 150
      Else
        aoXLights(i).BlinkInterval = 75
      End If
      aoXLights(i).state = LightStateBlinking
    ElseIf aanXTimers(nCurPlayer, i) > 0 Then
      aoXLights(i).state = LightStateOff

      if aanXTimers(nCurPlayer, i) > 12000 Then
        aoXLights(i).BlinkInterval = 300
      ElseIf aanXTimers(nCurPlayer, i) > 4000 Then
        aoXLights(i).BlinkInterval = 150
      Else
        aoXLights(i).BlinkInterval = 75
      End If

      aoXLights(i).state = LightStateBlinking
    End If
  next

  ' Update mode lights, on the same blink cycle
  for i = 0 to 8
    if aanBossMedals(nCurPlayer, i) <> cMedalNotPlayed Then
      aoBossLights(i).state = LightStateOn
    Else
      aoBossLights(i).state = LightStateOff
    End If
    aoBossLights(i).BlinkInterval = 300
    aoBossLights(i).BlinkPattern = "10"
  Next
  If BitSetContains(avBossesLit(nCurPlayer), cModeEnchantress) Then
    for i = 0 to 8
      aoBossLights(i).state = LightStateBlinking
    Next
  Else
    If bCanProgressLevel Then
      for i = 0 to 8
        if BitSetContains(avBossesLit(nCurPlayer), i) Then
          aoBossLights(i).state = LightStateBlinking
        Elseif i = anLevelSelected(nCurPlayer) Then
          aoBossLights(i).BlinkInterval = 75
          aoBossLights(i).BlinkPattern = "10001000"
          aoBossLights(i).state = LightStateBlinking
        End If
      Next
    Else
      for i = 0 to 8
        If BitSetContains(vModesRunning, i) Then
          aoBossLights(i).state = LightStateBlinking
        ElseIf BitSetContains(avBossesLit(nCurPlayer), i) Then
          aoBossLights(i).state = LightStateBlinking
        End If
      next
    End If
  End If

  ' update Troupple king Lights
  If BitSetContains(vModesRunning, cModeJuggling) Then
    for i = 0 to 2
      aoTrouppleLights(i).state = LightStateOff
    Next
  Elseif BitSetContains(vModesRunning, cModeMole) = False Then
    for i = 0 to 2
      aoTrouppleLights(i).state = LightStateOff
      If anTrouppleLit(nCurPlayer) < 1 Then
        If BitSetContains(avTrouppleHits(nCurPlayer), i) Then
          aoTrouppleLights(i).state = LightStateOn
        Else
          aoTrouppleLights(i).state = LightStateBlinking
        End If
      End If
    Next
  end If

  nColorCycleIndex = nColorCycleIndex + 1
  nColorCycleIndex = nColorCycleIndex mod 60
End Sub

Sub AllLightsOff
  Dim i

  for i = 0 to 11
    aoShovelLights(i).state = LightStateOff
  Next
  for i = 0 to 2
    aoTrouppleLights(i).state = LightStateOff
  Next
  for i = 0 to 8
    aoBossLights(i).state = LightStateOff
  Next
  for i = 0 to 9
    avShotColors(i) = 0
    aoShotLights(i).state = LightStateOff
    aoShotLights(i).color = Color_Gray
    aoShotLights(i).colorFull = vbWhite
  Next
  PlagueBlink.Enabled = False
  FlBumperFadeTarget(1) = 0
  FlBumperFadeTarget(2) = 0
  FlBumperFadeTarget(3) = 0
  for i = 0 to 4
    aoBottomLaneLights(i).state = LightStateOff
  Next
  for i = 0 to 9
    aoXLights(i).state = LightStateOff
    aoXLights(i).color = Color_Gray
    aoXLights(i).colorFull = vbWhite
  Next
  for i = 0 to 3
    aoWandererLights(i).state = LightStateOff
  next
  LightXBallLit.state = LightStateOff
  LightSJLit.state = LightStateOff
  LightShootAgain.state = LightStateOff
End Sub

'*** Light shows
Sub LightShowSpin(nDelay, nLoops)
  LightSeqModes.UpdateInterval = nDelay
  LightSeqMultipliers.UpdateInterval = nDelay
  LightSeqShovels.UpdateInterval = nDelay
  LightSeqTargets.UpdateInterval = nDelay

  LightSeqModes.Play SeqScrewRightOn, 60, nLoops
  LightSeqMultipliers.Play SeqScrewLeftOn, 60, nLoops
  LightSeqShovels.Play SeqScrewRightOn, 60, nLoops
  LightSeqTargets.Play SeqScrewLeftOn, 60, nLoops
End Sub

'*****Functions for bitsets

' Given a bitset and the index of an item, return the bitset with the item added
Function BitSetAdd (ByVal nBitSet, ByVal nIndex)
  BitSetAdd = nBitSet Or aPowersOfTwo(nIndex)
End Function

' Given a bitset and the index of an item, return the bitset with the item removed
Function BitSetRemove (ByVal nBitSet, ByVal nIndex)
  BitSetRemove = nBitSet And Not aPowersOfTwo(nIndex)
End Function

' Given a bitset and the index of an item, return True if the bitset contains the
' item, and False if it doesn't
Function BitSetContains (ByVal nBitSet, ByVal nIndex)
  Dim temp
  temp = nBitSet And aPowersOfTwo(nIndex)
  BitSetContains = Temp <> 0
End Function

' Given a bitset and its capacity, return True if the amount of items that the
' bitset contains is equal to its capacity, and return False otherwise.
Function BitSetIsFull (ByVal nBitSet, ByVal nCapacity)
  Dim temp
  temp = aPowersOfTwo(nCapacity) - 1
  BitSetIsFull = Not (nBitSet <> temp)
End Function

' Given a bitset and its capacity, return the number of items that the bitset
' contains.
Function BitSetCount (ByVal nBitSet, ByVal nCapacity)
  Dim i, nCount
  nCount = 0
  for i = 0 to (nCapacity - 1)
    If BitSetContains(nBitSet, i) Then nCount = nCount + 1
  Next
  BitSetCount = nCount
End Function

' Given a bitset and its capacity, return the bitset rotated one bit to the right
Sub RotLeft(nBitSet, nSize)
  Dim lsb
' save least significant bit
  lsb = nBitSet And 1
' shift all bits one to the right
  nBitSet = Int(nBitSet / 2)
' Set the most significant bit to what the least significant bit was
  If lsb <> 0 Then
    nBitSet = nBitSet Or aPowersOfTwo(nSize - 1)
  Else
    nBitSet = nBitSet And (Not aPowersOfTwo(nSize - 1))
  End If
End Sub

' Given a bitset and its capacity, return the bitset rotated one bit to the left
Sub RotRight(nBitSet, nSize)
  Dim msb
' save most significant bit
  msb = nBitSet And aPowersOfTwo(nSize - 1)
' shift all bits one to the left
  nBitSet = nBitSet * 2
' Clear anything past the most significant bit
  nBitSet = nBitSet And (Not aPowersOfTwo(nSize))
' Set the least significant bit to what the most significant bit was
  If msb <> 0 Then
    nBitSet = nBitSet Or aPowersOfTwo(0)
  Else
    nBitSet = nBitSet And (Not aPowersOfTwo(0))
  End If
End Sub

' *** Message Queue
Class Queue

  private maQueue()
  private mnSize
  private mnCapacity
  private mnIndex
  private i

  private sub Class_Initialize
    mnCapacity = 10
    mnIndex = 0
    mnSize = 0
    redim maQueue(mnCapacity)
  end Sub

  public property get Size
    Size = mnSize - mnIndex
  end property

  public Sub Enqueue(element)
    If mnSize >= mnCapacity Then
      mnCapacity = mnCapacity * 2
      Redim preserve maQueue(mnCapacity)
    End If
    maQueue(mnSize) = element
    mnSize = mnSize + 1
  end Sub

  public Sub PushFirst(element)
    If mnSize >= mnCapacity Then
      mnCapacity = mnCapacity * 2
      Redim preserve maQueue(mnCapacity)
    End If
    for i = mnSize to (mnIndex + 1) step -1
      maQueue(i) = maQueue(i - 1)
    Next
    maQueue(mnIndex) = element
    mnSize = mnSize + 1
  End Sub

  public function Dequeue()
    If mnSize - mnIndex < 1 Then
      Dequeue = null
    Else
      Dequeue = maQueue(mnIndex)
      mnIndex = mnIndex + 1
    End If
  end Function

  public sub Rewind
    mnIndex = 0
  end Sub

  public sub Clear
    mnCapacity = 10
    mnIndex = 0
    mnSize = 0
    redim maQueue(mnCapacity)
  end Sub

End Class

'***** High Score handling
Dim nHighScoreUpdates
Dim sLineOne
Dim sInitialsDisplayed

Sub HighscoreUpdate_Timer
  Dim i
  If HighscoreLetter = -1 Then
    SortedHighscoreNames(aScoreRanks(nHighScoreIter)) = InitialsEntered
    If nHighScoreIter < 3 Then
      For i = (nHighScoreIter + 1) to 3
        nHighScoreIter = i
        If aScoreRanks(i) < nHighScoreSlots Then
          InitHighscore
          Exit For
        End If
      Next
    Else
      HighscoreUpdate.enabled = False
      bEnteringHighScore = false
      savehs
      nTextDuration = 0
      LightSeq1.UpdateInterval = 25
      LightSeq1.Play SeqCircleOutOn, 40
      InitMusic 2
    End If
    Exit Sub
  End If
  If (nTimePressedRight > 0) And (gameTime - nTimePressedRight > 749) Then
    ChangeHighScoreLetter(1)
  End If
  If (nTimePressedLeft > 0) And (gameTime - nTimePressedLeft > 749) Then
    ChangeHighScoreLetter(-1)
  End If
  If nHighScoreUpdates < 5 Then
    sInitialsDisplayed = InitialsEntered & Mid(cHighScoreAlphabet, HighscoreLetter, 1) &_
      Space(2 - len(InitialsEntered))
  Else
    sInitialsDisplayed = InitialsEntered & "_" & Space(2 - len(InitialsEntered))
  End If
  ShowText sLineOne, getFormattedScore(nHighScoreIter) & " <" & _
    sInitialsDisplayed & ">", PlayUntilHit, 1
  nHighScoreUpdates = (nHighScoreUpdates + 1) mod 10
End Sub

Sub InitHighscore()
  bEnteringHighScore = True
  InitMusic 31
  If 0 = aScoreRanks(nHighScoreIter) Then
    sLineOne = "P" & (nHighScoreIter + 1) & " GRAND CHAMP"
  Else
    sLineOne = "P" & (nHighScoreIter + 1) & " HIGH SCORE " & aScoreRanks(nHighScoreIter)
  End If
  ShowText sLineOne, getFormattedScore(nHighScoreIter) & " <   >", PlayUntilHit, 1
  InitialsEntered = ""
  HighscoreLetter = 1
  nHighScoreUpdates = 0
  nTimePressedLeft = -1
  nTimePressedRight = -1
  HighscoreUpdate.enabled = True
End Sub

Sub HighScoreRightPressed
  If nTimePressedRight < 0 Then nTimePressedRight = GameTime
End Sub

Sub HighScoreRightReleased
  If GameTime - nTimePressedRight < 750 Then ChangeHighScoreLetter(1)
  nTimePressedRight = -1
End Sub

Sub HighScoreLeftPressed
  If nTimePressedLeft < 0 Then nTimePressedLeft = GameTime
End Sub

Sub HighScoreLeftReleased
  If GameTime - nTimePressedLeft < 750 Then ChangeHighScoreLetter(-1)
  nTimePressedLeft = -1
End Sub

Sub ChangeHighScoreLetter(offset)
  Dim nLetterCount
  nLetterCount = len(cHighScoreAlphabet)
  ' Do not allow deleting if no letter has been entered
  If "" = InitialsEntered Then nLetterCount = nLetterCount - 1
  If HighscoreLetter = -1 then Exit Sub
  PlaySound "fx-move_cursor", 0, 0.3
  HighscoreLetter = HighscoreLetter + offset
  If HighscoreLetter > nLetterCount Then HighscoreLetter = 1
  If HighscoreLetter < 1 Then HighscoreLetter = nLetterCount
  nHighScoreUpdates = 0
End Sub

Sub HighScoreEnterPressed
  ' letter "<" means delete
  Dim Letter
  If HighscoreLetter = -1 then Exit Sub
  PlaySound "fx-start", 0, 1
  Letter = Mid(cHighScoreAlphabet, HighscoreLetter, 1)
  If Letter = "<" And Len(InitialsEntered) > 0 Then
    Letter = Right(InitialsEntered, 1)
    InitialsEntered = Left(InitialsEntered, Len(InitialsEntered) - 1)
    sInitialsDisplayed = InitialsEntered & Letter & Space(2 - len(InitialsEntered))
    HighscoreLetter = inStr(1, cHighScoreAlphabet, Letter, vbTextCompare)
  Else
    InitialsEntered = InitialsEntered & Letter
    If Len(InitialsEntered) > 2 Then
      HighscoreLetter = -1
      sInitialsDisplayed = InitialsEntered
    Else
      sInitialsDisplayed = InitialsEntered & Letter & Space(2 - len(InitialsEntered))
    End If
  End If
  ShowText sLineOne, getFormattedScore(nHighScoreIter) & " (" & _
    sInitialsDisplayed & ")", PlayUntilHit, 1
End Sub

Sub Loadhs
    Dim x
    x = LoadValue(cGameName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 250000 End If
    x = LoadValue(cGameName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "SHO" End If
    x = LoadValue(cGameName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 150000 End If
    x = LoadValue(cGameName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "VEL" End If
    x = LoadValue(cGameName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 100000 End If
    x = LoadValue(cGameName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "KNI" End If
    x = LoadValue(cGameName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 60000 End If
    x = LoadValue(cGameName, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "GHT" End If
End Sub

Sub Savehs
  Dim i
  For i = 0 to nHighScoreSlots - 1
    HighScore(i) = SortedHighscores(i)
    HighScoreName(i) = SortedHighscoreNames(i)
  Next
    SaveValue cGameName, "HighScore1", SortedHighscores(0)
    SaveValue cGameName, "HighScore1Name", SortedHighscoreNames(0)
    SaveValue cGameName, "HighScore2", SortedHighscores(1)
    SaveValue cGameName, "HighScore2Name", SortedHighscoreNames(1)
    SaveValue cGameName, "HighScore3", SortedHighscores(2)
    SaveValue cGameName, "HighScore3Name", SortedHighscoreNames(2)
    SaveValue cGameName, "HighScore4", SortedHighscores(3)
    SaveValue cGameName, "HighScore4Name", SortedHighscoreNames(3)
End Sub

' Input: an array of scores of player 1, 2, etc.
' Output: an array of the position in the highscores of the player
Sub CheckHighScore(aScores)
  Dim aIndices()
  Dim scoreIndices()
  Dim i,j
  Dim highestScore
  Dim highestIndex
  Dim bWillEnterHighScore
  Dim temp

  Redim SortedHighscoreNames(nHighScoreSlots + UBound(aScores))
  Redim SortedHighscores(nHighScoreSlots + UBound(aScores))
  Redim aIndices(nHighScoreSlots + UBound(aScores))
  Redim scoreIndices(UBound(aScores))
  bWillEnterHighScore = False

  For i = 0 to UBound(SortedHighscores)
    aIndices(i) = i
    If i < nHighScoreSlots Then
      SortedHighscores(i) = HighScore(i)
      SortedHighscoreNames(i) = HighScoreName(i)
    Else
      temp = aScores(i - nHighScoreSlots)
      SortedHighscores(i) = temp
      SortedHighscoreNames(i) = "NEW"
    End If
  Next

  ' sort the scores and adjust the aIndices array after how the scores
  ' were sorted
  For i = 0 to UBound(SortedHighscores)
    highestScore = SortedHighscores(i)
    highestIndex = i
    For j = i + 1 to UBound(SortedHighscores)
      If SortedHighscores(j) > highestScore Then
        highestScore = SortedHighscores(j)
        highestIndex = j
      End If
    Next
    temp = SortedHighscoreNames(i)
    SortedHighscoreNames(i) = SortedHighscoreNames(highestIndex)
    SortedHighscoreNames(highestIndex) = temp
    temp = SortedHighscores(i)
    SortedHighscores(i) = SortedHighscores(highestIndex)
    SortedHighscores(highestIndex) = temp
    temp = aIndices(i)
    aIndices(i) = aIndices(highestIndex)
    aIndices(highestIndex) = temp
  Next

  ' return an array of the positions of the new scores in the sorted scores
  For i = nHighScoreSlots to UBound(aIndices)
    For j = 0 to UBound(SortedHighscores)
      If aIndices(j) = i Then scoreIndices(i - nHighScoreSlots) = j
    Next
  Next

  aScoreRanks = scoreIndices

  For i = 0 to 3
    nHighScoreIter = i
    If aScoreRanks(i) < nHighScoreSlots Then
      nTextDuration = 0
      bWillEnterHighScore = True
      InitHighscore
      Exit For
    End If
  Next

  If bWillEnterHighScore = False Then
'   nTextDuration = 0
    LightSeq1.UpdateInterval = 25
    LightSeq1.Play SeqCircleOutOn, 40
    InitMusic 2
  End If
End Sub

'******************************************************
'******  FLUPPER BUMPERS
'******************************************************
' Based on FlupperBumpers 0.145 final

' Explanation of how these bumpers work:
' There are 10 elements involved per bumper:
' - the shadow of the bumper ( a vpx flasher object)
' - the bumper skirt (primitive)
' - the bumperbase (primitive)
' - a vpx light which colors everything you can see through the bumpertop
' - the bulb (primitive)
' - another vpx light which lights up everything around the bumper
' - the bumpertop (primitive)
' - the VPX bumper object
' - the bumper screws (primitive)
' - the bulb highlight VPX flasher object
' All elements have a special name with the number of the bumper at the end, this is necessary for the fading routine and the initialisation.
' For the bulb and the bumpertop there is a unique material as well per bumpertop.
' To use these bumpers you have to first copy all 10 elements to your table.
' Also export the textures (images) with names that start with "Flbumper" and "Flhighlight" and materials with names that start with "bumper".
' Make sure that all the ten objects are aligned on center, if possible with the exact same x,y coordinates
' After that copy the script (below); also copy the BumperTimer vpx object to your table
' Every bumper needs to be initialised with the FlInitBumper command, see example below;
' Colors available are red, white, blue, orange, yellow, green, purple and blacklight.
' In a GI subroutine you can then call set the bumperlight intensity with the "FlBumperFadeTarget(nr) = value" command
' where nr is the number of the bumper, value is between 0 (off) and 1 (full on) (so you can also use 0.3 0.4 etc).

' Notes:
' - There is only one color for the disk; you can photoshop it to a different color
' - The bumpertops are angle independent up to a degree; my estimate is -45 to + 45 degrees horizontally, 0 (topview) to 70-80 degrees (frontview)
' - I built in correction for the day-night slider; this might not work perfectly, depending on your table lighting
' - These elements, textures and materials do NOT integrate with any of the lighting routines I have seen in use in many VPX tables
'   (just find the GI handling routine and insert the FlBumperFadeTarget statement)
' - If you want to use VPX native bumperdisks just copy my bumperdisk but make it invisible



' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0 : DNA45 = (NightDay-10)/20 : DNA90 = 0 : DayNightAdjust = 0.4
Else
  DNA30 = (NightDay-10)/30 : DNA45 = (NightDay-10)/45 : DNA90 = (NightDay-10)/90 : DayNightAdjust = NightDay/25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt : For cnt = 1 to 6 : FlBumperActive(cnt) = False : Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight

FlInitBumper 1, "blue"
FlInitBumper 2, "blue"
FlInitBumper 3, "blue"

' ### uncomment the statement below to change the color for all bumpers ###
' Dim ind : For ind = 1 to 5 : FlInitBumper ind, "green" : next

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True
  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1 : FlBumperFadeTarget(nr) = 1.1: FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr) : FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr) : Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr) : Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr) : FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr): FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
  ' set the color for the two VPX lights
  select case col
    Case "red"
      FlBumperSmallLight(nr).color = RGB(255,4,0) : FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
      FlBumperBigLight(nr).color = RGB(255,32,0) : FlBumperBigLight(nr).colorfull = RGB(255,32,0)
      FlBumperHighlight(nr).color = RGB(64,255,0)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
      FlBumperSmallLight(nr).TransmissionScale = 0
    Case "blue"
      FlBumperBigLight(nr).color = RGB(32,80,255) : FlBumperBigLight(nr).colorfull = RGB(32,80,255)
      FlBumperSmallLight(nr).color = RGB(0,80,255) : FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
      FlBumperSmallLight(nr).TransmissionScale = 0 : MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
      FlBumperHighlight(nr).color = RGB(255,16,8)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    Case "green"
      FlBumperSmallLight(nr).color = RGB(8,255,8) : FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
      FlBumperBigLight(nr).color = RGB(32,255,32) : FlBumperBigLight(nr).colorfull = RGB(32,255,32)
      FlBumperHighlight(nr).color = RGB(255,32,255) : MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
      FlBumperSmallLight(nr).TransmissionScale = 0.005
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    Case "orange"
      FlBumperHighlight(nr).color = RGB(255,130,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).color = RGB(255,130,0) : FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
      FlBumperBigLight(nr).color = RGB(255,190,8) : FlBumperBigLight(nr).colorfull = RGB(255,190,8)
    Case "white"
      FlBumperBigLight(nr).color = RGB(255,230,190) : FlBumperBigLight(nr).colorfull = RGB(255,230,190)
      FlBumperHighlight(nr).color = RGB(255,180,100) :
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
    Case "blacklight"
      FlBumperBigLight(nr).color = RGB(32,32,255) : FlBumperBigLight(nr).colorfull = RGB(32,32,255)
      FlBumperHighlight(nr).color = RGB(48,8,255) :
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    Case "yellow"
      FlBumperSmallLight(nr).color = RGB(255,230,4) : FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
      FlBumperBigLight(nr).color = RGB(255,240,50) : FlBumperBigLight(nr).colorfull = RGB(255,240,50)
      FlBumperHighlight(nr).color = RGB(255,255,220)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
    Case "purple"
      FlBumperBigLight(nr).color = RGB(80,32,255) : FlBumperBigLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).color = RGB(80,32,255) : FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).TransmissionScale = 0 :
      FlBumperHighlight(nr).color = RGB(32,64,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
  end select
End Sub

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
' UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
'               OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
'               float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 )* DayNightAdjust

  select case FlBumperColor(nr)

    Case "blue" :
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(38-24*Z,130 - 98*Z,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20  + 500 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z +0.97 * Z^3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 10000 * (Z^3) / (0.5 + DNA90)

    Case "green"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(16 + 16 * sin(Z*3.14),255,16 + 16 * sin(Z*3.14)), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
      FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 6000 * (Z^3) / (1 + DNA90)

    Case "red"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 16 - 11*Z + 16 * sin(Z*3.14),0), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z*4,8-Z*8)

    Case "orange"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 100 - 22*z  + 16 * sin(Z*3.14),Z*32), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z*50, 0)

    Case "white"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
      FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 8 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255,255 - 20*Z,255-65*Z) : FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20*Z,255-65*Z)
      MaterialColor "bumpertopmat" & nr, RGB(255,235 - z*36,220 - Z*90)

    Case "blacklight"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 1, RGB(30-27*Z^0.03,30-28*Z^0.01, 255), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z^3
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255-240*(Z^0.1),255 - 240*(Z^0.1),255) : FlBumperSmallLight(nr).colorfull = RGB(255-200*z,255 - 200*Z,255)
      MaterialColor "bumpertopmat" & nr, RGB(255-190*Z,235 - z*180,220 + 35*Z)

    Case "yellow"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 180 + 40*z, 48* Z), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

    Case "purple" :
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(128-118*Z - 32 * sin(Z*3.14), 32-26*Z ,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 15  + 200 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z +0.97 * Z^3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (0.5 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(128-60*Z,32,255)


  end select
End Sub

Sub BumperTimer_Timer
  dim nr
  For nr = 1 to 6
    If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
      If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
    If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
      If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
  next
End Sub


'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

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
Sub CaptiveWalls_Hit (idx)
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
  PlaySoundAt snd, CapKicker1
End Sub


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

'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
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
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    ' If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
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

  Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************




'******************************************************
'**** RAMP ROLLING SFX
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



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'*******************************************
'  Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend

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

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub


'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

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



'******************************************************
'****  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Flippers:   https://www.youtube.com/watch?v=FWvM9_CdVHw
' Dampeners:  https://www.youtube.com/watch?v=tqsxx48C6Pg
' Physics:    https://www.youtube.com/watch?v=UcRMG-2svvE
'
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25)
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |



'******************************************************
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we’ll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
' 4. and, special scripting
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity


'*******************************************
'  Late 80's early 90's

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
  Next

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -5
  AddPt "Polarity", 2, 0.4, -5
  AddPt "Polarity", 3, 0.6, -4.5
  AddPt "Polarity", 4, 0.65, -4.0
  AddPt "Polarity", 5, 0.7, -3.5
  AddPt "Polarity", 6, 0.75, -3.0
  AddPt "Polarity", 7, 0.8, -2.5
  AddPt "Polarity", 8, 0.85, -2.0
  AddPt "Polarity", 9, 0.9,-1.5
  AddPt "Polarity", 10, 0.95, -1.0
  AddPt "Polarity", 11, 1, -0.5
  AddPt "Polarity", 12, 1.1, 0
  AddPt "Polarity", 13, 1.3, 0

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
Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

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


'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************








'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
'
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

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
' Thalamus - patched :       aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
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


'****************************************
' Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub Realtime_Timer
    RollingUpdate
    propeller.Roty = Spinner1.CurrentAngle
End Sub


'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

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

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
' Cor.Update
'End Sub

' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking (this sometimes goes in the RDampen_Timer sub)
  RollingUpdate         'update rolling sounds
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************

'*******************************************
' VR Room
'*******************************************

Sub InitVR
  DIM VRThings
  If bUseVR Then
    If VRRoomChoice = 1 Then
      for each VRThings in VR_Cab:VRThings.visible = 1:Next
    End If
    If VRRoomChoice = 2 Then
      for each VRThings in VR_Cab:VRThings.visible = 0:Next
      PinCab_Backglass.visible = 1
    End If
  Else
    for each VRThings in VR_Cab:VRThings.visible = 0:Next
  End if
End Sub

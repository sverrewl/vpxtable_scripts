' Team One (Gottlieb 1977)
' https://www.ipdb.org/machine.cgi?id=2507

' CHANGELOG:
' 0.01 - mcarter78 - 4k VLM batch, new scratch build (new playfield & plastics scans) using Loserman's script, updated physics code & params, most movables animated, update knocker sequence & add EM buzz to rollovers
' 0.02 - mcarter78 - new 4k batch with visual fixes (pf, plastics, right flipper, etc..), animate bumpers & slings, fix misnamed lights
' 0.03 - mcarter78 - Fix physical objects alignment, reshape flipper triggers, fix materials & collections, add all options to F12 menu, hook up instruction/replay cards & high score post-it, update LUT & DT POV
' 0.04 - mcarter78 - Make translucent insert cups shallower, fix apron walls, lower flipper elasticity, increase drop target resistance, add under pf wall, add hit sound to both sides of gate, add cabinet mode
' 0.05 - mcarter78 - new 4k batch with updated pf image (thanks tomate!), fixed visual drop targets and reworked inserts & lights strength.  Added knocker sound to 10,000 pt reel advance.
' 0.06 - mcarter78 - apron model fixed (thanks again tomate!), increase drop targets elasticity, increase bumper radius, tweak rubber animation, fix drop reset sound, fix scoring routine to count 10,000's place first
' 0.07 - mcarter78 - adjust rubber animation timers, fix physical gate position
' 0.08 - mcarter78 - new DT backdrop, add ball release sound
' 0.09 - bthlonewolf - very minor tweak to left/right slingshots
' 0.10 - bthlonewolf - update A/B sling animation (testing)
' 0.11 - mcarter78 - 4k batch with new rubber objects for sling animation. Fixed apron edge artifacts & lit targets blurriness.
' 0.12 - bthlonewolf - added anims for slings A, B, C, D
' 0.13 - mcarter78 - add additional rubber positions for slings A, B, C, D, separate leaf switches for those slings, hoist high score init
' 0.14 - mcarter78 - fix sling rubbers render, fix desktop POV
' 0.15 - bthlonewolf - tweaked sling animations; implemented hit threshold hit events
' 0.16 - mcarter78 - Updated flipper physics per feedback from Cliffy
' 0.17 - mcarter78 - Fix looping ballrelease sound, add buzz to game start, increase flipper angles
' 0.18 - Ext2k - Add VR cabinet & minimal Room
' 0.19 - mcarter78 - Add VR mega room and vr room options.
' 0.20 - mcarter78 - Add EM reels for VR, add images for VR backglass lights, remove reflections from Mega room prims
' 0.21 - mcarter78 - Add VR backglass lights
' 0.22 - mcarter78 - Fixed VR cab backbox (thanks tomate!), new VR backglass mesh with inset reels
' 0.23 - mcarter78 - Fix scoring & knocker routines, New backglass image on backdrop, add top lanes spotting option, fix cabinet mode option
' 0.24 - mcarter78 - Change game name to match DOF config. Add rules card to F12 menu & fill out table info. Add hit events to Wall004/Wall005/Wall011/Wall012.
' 0.25 - mcarter78 - Fix flipper shadows DB, fix plunger material & strength.  Reduce drop targets' mass.
' 0.26 - bthlonewolf - modified cgamename for DOF, modified some DOF codes (new codes @ 150+, tried to keep compat w/ prev)
' 0.27 - mcarter78 - Refactor drop target handling, implement suggested hit thresholds, rework gate sounds, change rubber wheel material
' RC1  - mcarter78 - Fix VR env DB issues, remove FS image & make background black, add stray VR objects to correct collections.  Add credits to Table Info.
' RC2  - mcarter78 - Remove textbox objects
' RC3  - mcarter78 - Change "Replay" to "Add-a-ball" in options, Don't award add-a-ball points in novelty mode.
' RC4  - mcarter78 - Update credits in table info
' RC5  - mcarter78 - Animate VR Cabinet buttons
' RC6  - mcarter78 - Fix VR Plunger animation
' RC7  - mcarter78 - Stop flipper buzz on every Flipper KeyUp
' RC8  - mcarter78 - Disable calls to Controller in drop target code
' RC9  - mcarter78 - Add standup target blocker primitive & move scoring logic to standup code
' RC10 - mcarter78 - Fix flipper shadows alignment
' Release 1.0
' 1.0.1 - FrankEnstein - Added reflection options for quality and scope.
' 1.0.2 - mcarter78 - Fix material for rollover, add VR_BGLit to VR_Mega collection, fix bumper lights
' 1.0.3 - mcarter78 - Increase apron elasticity
' 1.0.4 - mcarter78 - Incorporate Sixtoe's RC8 changes
' Release 1.1

option explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "TeamOne_77VPX"

Const ShadowConfigFile = false

Const Ballsize = 50
Const BallMass = 1
Const tnob = 1
Const lob = 0

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim gilvl

Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed

Const HSFileName="TeamOne_77VPX.txt"
Const B2STableName="TeamOne_1977"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"

Dim B2SOn   'True/False if want backglass


'* this value adjusts score motor behavior - 0 allows you to continue scoring while the score motor is running - 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment=1

'* this is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment=1

Dim HSScore(5)        ' High Scores read in from config file
Dim HSName(5)       ' High Score Initials read in from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 75000
HSScore(2) = 70000
HSScore(3) = 60000
HSScore(4) = 55000
HSScore(5) = 50000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"


dim ScoreChecker
dim CheckAllScores
dim sortscores(4)
dim sortplayers(4)
Dim TextStr,TextStr2
Dim i,xx, tKickerTCount
Dim obj
Dim bgpos
Dim kgpos
Dim dooralreadyopen
Dim kgdooralreadyopen
Dim TargetSpecialLit
Dim Points210counter
Dim Points500counter
Dim Points1000counter
Dim Points2000counter
Dim InProgress
Dim BallInPlay
Dim CreditsPerCoin
Dim Score100K(4)
Dim Score(4)
Dim ScoreDisplay(4)
Dim HighScorePaid(4)
Dim HighScore
Dim HighScoreReward
Dim BonusMultiplier
Dim Credits
Dim Match
Dim Replay1
Dim Replay2
Dim Replay3
Dim Replay1Paid(4)
Dim Replay2Paid(4)
Dim Replay3Paid(4)
Dim TableTilted
Dim TiltCount
Dim BallsPerGame

Dim BonusBooster
Dim BonusBoosterCounter
Dim BonusCounter
Dim HoleCounter

Dim Ones
Dim Tens
Dim Hundreds
Dim Thousands

Dim Player
Dim Players

Dim rst
Dim bonuscountdown
Dim TempMultiCounter
dim TempPlayerup
dim RotatorTemp

Dim bump1
Dim bump2
Dim bump3

Dim LastChime10
Dim LastChime100
Dim LastChime1000

Dim Score10
Dim Score100

Dim LeftTargetCounter
Dim RightTargetCounter

Dim MotorRunning
Dim Replay1Table(15)
Dim Replay2Table(15)
Dim Replay3Table(15)
Dim ReplayTableSet
Dim ReplayLevel
Dim ReplayTableMax
Dim ReflectionsScope
Dim ReflectionsQuality
Dim BonusSpecialThreshold
Dim AlternatingRelayCounter
Dim ABCCompleted
Dim NoveltyMode,TiltEndsGame,WOWCounter,LStep,RStep


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper1_Ring_001: BP_Bumper1_Ring_001=Array(BM_Bumper1_Ring_001, LM_L_Bumper1Light_Bumper1_Ring_, LM_L_LeftInlaneLight5_Bumper1_R, LM_GI_Split_gi17_Bumper1_Ring_0, LM_GI_Split_gi18_Bumper1_Ring_0, LM_GI_Split_gi19_Bumper1_Ring_0, LM_GI_Split_gi28_Bumper1_Ring_0, LM_GI_Bumper1_Ring_001)
Dim BP_Bumper1_Skirt_001: BP_Bumper1_Skirt_001=Array(BM_Bumper1_Skirt_001, LM_L_Bumper1Light_Bumper1_Skirt, LM_GI_Split_gi28_Bumper1_Skirt_, LM_GI_Bumper1_Skirt_001)
Dim BP_Bumper2_Ring_001: BP_Bumper2_Ring_001=Array(BM_Bumper2_Ring_001, LM_L_Bumper2Light_Bumper2_Ring_, LM_L_RightInlaneLight6_Bumper2_, LM_GI_Split_gi19_Bumper2_Ring_0, LM_GI_Split_gi20_Bumper2_Ring_0, LM_GI_Split_gi8_Bumper2_Ring_00, LM_GI_Bumper2_Ring_001)
Dim BP_Bumper2_Skirt_001: BP_Bumper2_Skirt_001=Array(BM_Bumper2_Skirt_001, LM_L_Bumper2Light_Bumper2_Skirt, LM_GI_Split_gi8_Bumper2_Skirt_0, LM_GI_Bumper2_Skirt_001)
Dim BP_DT1: BP_DT1=Array(BM_DT1, LM_L_LeftTargetLight1_DT1, LM_L_LeftTargetLight2_DT1, LM_GI_DT1, LM_L_l25_DT1, LM_L_l25a_DT1)
Dim BP_DT10: BP_DT10=Array(BM_DT10, LM_L_RightTargetLight4_DT10, LM_L_RightTargetLight5_DT10, LM_GI_DT10, LM_L_l25_DT10, LM_L_l25a_DT10)
Dim BP_DT2: BP_DT2=Array(BM_DT2, LM_L_LeftTargetLight1_DT2, LM_L_LeftTargetLight2_DT2, LM_L_LeftTargetLight3_DT2, LM_GI_DT2, LM_L_l25_DT2, LM_L_l25a_DT2)
Dim BP_DT3: BP_DT3=Array(BM_DT3, LM_L_LFive_DT3, LM_L_LeftTargetLight1_DT3, LM_L_LeftTargetLight2_DT3, LM_L_LeftTargetLight3_DT3, LM_GI_DT3, LM_L_l25_DT3, LM_L_l25a_DT3)
Dim BP_DT4: BP_DT4=Array(BM_DT4, LM_L_LFive_DT4, LM_L_LeftTargetLight2_DT4, LM_L_LeftTargetLight3_DT4, LM_L_LeftTargetLight4_DT4, LM_GI_DT4, LM_L_l25_DT4, LM_L_l25a_DT4)
Dim BP_DT5: BP_DT5=Array(BM_DT5, LM_L_LFive_DT5, LM_L_LeftTargetLight3_DT5, LM_L_LeftTargetLight4_DT5, LM_L_LeftTargetLight5_DT5, LM_GI_DT5, LM_L_l25_DT5, LM_L_l25a_DT5)
Dim BP_DT6: BP_DT6=Array(BM_DT6, LM_L_LSix_DT6, LM_L_RightTargetLight1_DT6, LM_L_RightTargetLight2_DT6, LM_GI_DT6, LM_L_l25_DT6, LM_L_l25a_DT6)
Dim BP_DT7: BP_DT7=Array(BM_DT7, LM_L_LSix_DT7, LM_L_RightTargetLight2_DT7, LM_L_RightTargetLight3_DT7, LM_L_RightTargetLight4_DT7, LM_GI_DT7, LM_L_l25_DT7, LM_L_l25a_DT7)
Dim BP_DT8: BP_DT8=Array(BM_DT8, LM_L_LSix_DT8, LM_L_RightTargetLight3_DT8, LM_L_RightTargetLight4_DT8, LM_GI_DT8, LM_L_l25_DT8, LM_L_l25a_DT8)
Dim BP_DT9: BP_DT9=Array(BM_DT9, LM_L_RightTargetLight3_DT9, LM_L_RightTargetLight4_DT9, LM_L_RightTargetLight5_DT9, LM_GI_DT9, LM_L_l25_DT9, LM_L_l25a_DT9)
Dim BP_LF: BP_LF=Array(BM_LF, LM_GI_Split_gi19_LF)
Dim BP_LFU: BP_LFU=Array(BM_LFU, LM_GI_Split_gi18_LFU, LM_GI_Split_gi19_LFU)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_L_LFive_Layer1, LM_L_LSix_Layer1, LM_L_LeftTargetLight1_Layer1, LM_L_LeftTargetLight2_Layer1, LM_L_LeftTargetLight3_Layer1, LM_L_LeftTargetLight4_Layer1, LM_L_LeftTargetLight5_Layer1, LM_L_UpperLight2_Layer1, LM_L_UpperLight3_Layer1, LM_L_RightTargetLight1_Layer1, LM_L_RightTargetLight2_Layer1, LM_L_RightTargetLight3_Layer1, LM_L_RightTargetLight4_Layer1, LM_L_RightTargetLight5_Layer1, LM_GI_Split_gi17_Layer1, LM_GI_Split_gi18_Layer1, LM_GI_Split_gi19_Layer1, LM_GI_Split_gi20_Layer1, LM_GI_Split_gi28_Layer1, LM_GI_Split_gi8_Layer1, LM_GI_Layer1, LM_L_l25_Layer1, LM_L_l25a_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_GI_Layer2, LM_L_l25_Layer2, LM_L_l25a_Layer2)
Dim BP_LeafA: BP_LeafA=Array(BM_LeafA, LM_GI_LeafA)
Dim BP_LeafB: BP_LeafB=Array(BM_LeafB, LM_GI_LeafB)
Dim BP_LeafC: BP_LeafC=Array(BM_LeafC, LM_GI_LeafC)
Dim BP_LeafD: BP_LeafD=Array(BM_LeafD, LM_GI_LeafD)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_L_Bumper1Light_Parts, LM_L_Bumper2Light_Parts, LM_L_LeftInlaneLight5000_Parts, LM_L_LeftInlaneLight5_Parts, LM_L_RightInlaneLight6_Parts, LM_L_RightInlaneLight5000_Parts, LM_L_LCredit_Parts, LM_L_LFive_Parts, LM_L_LeftSpecialLight_Parts, LM_L_CenterSpecialLight_Parts, LM_L_RightSpecialLight_Parts, LM_L_LSix_Parts, LM_L_UpperLight1_Parts, LM_L_LeftTargetLight1_Parts, LM_L_LeftTargetLight2_Parts, LM_L_LeftTargetLight3_Parts, LM_L_LeftTargetLight4_Parts, LM_L_LeftTargetLight5_Parts, LM_L_UpperLight2_Parts, LM_L_UpperLight3_Parts, LM_L_UpperLight4_Parts, LM_L_RightTargetLight1_Parts, LM_L_RightTargetLight2_Parts, LM_L_RightTargetLight3_Parts, LM_L_RightTargetLight4_Parts, LM_L_RightTargetLight5_Parts, LM_GI_Split_gi17_Parts, LM_GI_Split_gi18_Parts, LM_GI_Split_gi19_Parts, LM_GI_Split_gi20_Parts, LM_GI_Split_gi28_Parts, LM_GI_Split_gi8_Parts, LM_GI_Parts, LM_L_l25_Parts, LM_L_l25a_Parts)
Dim BP_Pgate: BP_Pgate=Array(BM_Pgate, LM_GI_Pgate)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_L_Bumper1Light_Playfield, LM_L_Bumper2Light_Playfield, LM_L_LFive_Playfield, LM_L_LSix_Playfield, LM_L_LeftTargetLight1_Playfield, LM_L_LeftTargetLight2_Playfield, LM_L_LeftTargetLight3_Playfield, LM_L_LeftTargetLight4_Playfield, LM_L_LeftTargetLight5_Playfield, LM_L_RightTargetLight1_Playfiel, LM_L_RightTargetLight2_Playfiel, LM_L_RightTargetLight3_Playfiel, LM_L_RightTargetLight4_Playfiel, LM_L_RightTargetLight5_Playfiel, LM_GI_Split_gi17_Playfield, LM_GI_Split_gi18_Playfield, LM_GI_Split_gi19_Playfield, LM_GI_Split_gi20_Playfield, LM_GI_Split_gi28_Playfield, LM_GI_Split_gi8_Playfield, LM_GI_Playfield, LM_L_l25_Playfield, LM_L_l25a_Playfield)
Dim BP_RF: BP_RF=Array(BM_RF, LM_GI_Split_gi18_RF)
Dim BP_RFU: BP_RFU=Array(BM_RFU, LM_GI_Split_gi18_RFU, LM_GI_Split_gi19_RFU)
Dim BP_RailL: BP_RailL=Array(BM_RailL)
Dim BP_RailR: BP_RailR=Array(BM_RailR)
Dim BP_ST11: BP_ST11=Array(BM_ST11, LM_GI_ST11, LM_L_l25_ST11, LM_L_l25a_ST11)
Dim BP_SlingA_000: BP_SlingA_000=Array(BM_SlingA_000, LM_L_LFive_SlingA_000, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_GI_SlingA_000, LM_L_l25_SlingA_000, LM_L_l25a_SlingA_000)
Dim BP_SlingA_002: BP_SlingA_002=Array(BM_SlingA_002, LM_L_LFive_SlingA_002, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_GI_SlingA_002, LM_L_l25_SlingA_002, LM_L_l25a_SlingA_002)
Dim BP_SlingA_003: BP_SlingA_003=Array(BM_SlingA_003, LM_L_LFive_SlingA_003, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_GI_SlingA_003, LM_L_l25_SlingA_003, LM_L_l25a_SlingA_003)
Dim BP_SlingA_004: BP_SlingA_004=Array(BM_SlingA_004, LM_L_LFive_SlingA_004, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_GI_SlingA_004, LM_L_l25_SlingA_004, LM_L_l25a_SlingA_004)
Dim BP_SlingA_005: BP_SlingA_005=Array(BM_SlingA_005, LM_L_LFive_SlingA_005, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_GI_SlingA_005, LM_L_l25_SlingA_005, LM_L_l25a_SlingA_005)
Dim BP_SlingB_000: BP_SlingB_000=Array(BM_SlingB_000, LM_L_LSix_SlingB_000, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_GI_SlingB_000, LM_L_l25_SlingB_000, LM_L_l25a_SlingB_000)
Dim BP_SlingB_002: BP_SlingB_002=Array(BM_SlingB_002, LM_L_LSix_SlingB_002, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_GI_SlingB_002, LM_L_l25_SlingB_002, LM_L_l25a_SlingB_002)
Dim BP_SlingB_003: BP_SlingB_003=Array(BM_SlingB_003, LM_L_LSix_SlingB_003, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_GI_SlingB_003, LM_L_l25_SlingB_003, LM_L_l25a_SlingB_003)
Dim BP_SlingB_004: BP_SlingB_004=Array(BM_SlingB_004, LM_L_LSix_SlingB_004, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_GI_SlingB_004, LM_L_l25_SlingB_004, LM_L_l25a_SlingB_004)
Dim BP_SlingB_005: BP_SlingB_005=Array(BM_SlingB_005, LM_L_LSix_SlingB_005, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_GI_SlingB_005, LM_L_l25_SlingB_005, LM_L_l25a_SlingB_005)
Dim BP_SlingC_000: BP_SlingC_000=Array(BM_SlingC_000, LM_L_Bumper1Light_SlingC_000, LM_L_LeftTargetLight2_SlingC_00, LM_L_LeftTargetLight3_SlingC_00, LM_GI_Split_gi28_SlingC_000, LM_GI_SlingC_000)
Dim BP_SlingC_002: BP_SlingC_002=Array(BM_SlingC_002, LM_L_Bumper1Light_SlingC_002, LM_GI_Split_gi28_SlingC_002, LM_GI_SlingC_002)
Dim BP_SlingC_003: BP_SlingC_003=Array(BM_SlingC_003, LM_L_Bumper1Light_SlingC_003, LM_GI_Split_gi28_SlingC_003, LM_GI_SlingC_003)
Dim BP_SlingC_004: BP_SlingC_004=Array(BM_SlingC_004, LM_L_Bumper1Light_SlingC_004, LM_GI_Split_gi28_SlingC_004, LM_GI_SlingC_004)
Dim BP_SlingC_005: BP_SlingC_005=Array(BM_SlingC_005, LM_L_Bumper1Light_SlingC_005, LM_GI_Split_gi28_SlingC_005, LM_GI_SlingC_005)
Dim BP_SlingD_000: BP_SlingD_000=Array(BM_SlingD_000, LM_L_Bumper2Light_SlingD_000, LM_L_RightTargetLight5_SlingD_0, LM_GI_Split_gi8_SlingD_000, LM_GI_SlingD_000)
Dim BP_SlingD_002: BP_SlingD_002=Array(BM_SlingD_002, LM_L_Bumper2Light_SlingD_002, LM_GI_Split_gi8_SlingD_002, LM_GI_SlingD_002)
Dim BP_SlingD_003: BP_SlingD_003=Array(BM_SlingD_003, LM_L_Bumper2Light_SlingD_003, LM_GI_Split_gi8_SlingD_003, LM_GI_SlingD_003)
Dim BP_SlingD_004: BP_SlingD_004=Array(BM_SlingD_004, LM_L_Bumper2Light_SlingD_004, LM_GI_Split_gi8_SlingD_004, LM_GI_SlingD_004)
Dim BP_SlingD_005: BP_SlingD_005=Array(BM_SlingD_005, LM_L_Bumper2Light_SlingD_005, LM_GI_Split_gi8_SlingD_005, LM_GI_SlingD_005)
Dim BP_TriggerLeftInlane: BP_TriggerLeftInlane=Array(BM_TriggerLeftInlane, LM_GI_Split_gi17_TriggerLeftInl, LM_GI_Split_gi28_TriggerLeftInl)
Dim BP_TriggerLeftInlane2: BP_TriggerLeftInlane2=Array(BM_TriggerLeftInlane2, LM_GI_Split_gi17_TriggerLeftInl, LM_GI_Split_gi18_TriggerLeftInl)
Dim BP_TriggerLeftOutlane: BP_TriggerLeftOutlane=Array(BM_TriggerLeftOutlane, LM_GI_Split_gi17_TriggerLeftOut)
Dim BP_TriggerRightInlane: BP_TriggerRightInlane=Array(BM_TriggerRightInlane, LM_GI_Split_gi20_TriggerRightIn, LM_GI_Split_gi8_TriggerRightInl)
Dim BP_TriggerRightInlane2: BP_TriggerRightInlane2=Array(BM_TriggerRightInlane2, LM_GI_Split_gi19_TriggerRightIn, LM_GI_Split_gi20_TriggerRightIn)
Dim BP_TriggerRightOutlane: BP_TriggerRightOutlane=Array(BM_TriggerRightOutlane, LM_GI_Split_gi20_TriggerRightOu)
Dim BP_TriggerTop1: BP_TriggerTop1=Array(BM_TriggerTop1, LM_GI_TriggerTop1, LM_L_l25_TriggerTop1)
Dim BP_TriggerTop2: BP_TriggerTop2=Array(BM_TriggerTop2, LM_GI_TriggerTop2, LM_L_l25_TriggerTop2, LM_L_l25a_TriggerTop2)
Dim BP_TriggerTop3: BP_TriggerTop3=Array(BM_TriggerTop3, LM_GI_TriggerTop3, LM_L_l25_TriggerTop3, LM_L_l25a_TriggerTop3)
Dim BP_TriggerTop4: BP_TriggerTop4=Array(BM_TriggerTop4, LM_GI_TriggerTop4, LM_L_l25a_TriggerTop4)
Dim BP_lockdownbar: BP_lockdownbar=Array(BM_lockdownbar)
Dim BP_lsling_000: BP_lsling_000=Array(BM_lsling_000, LM_GI_lsling_000)
Dim BP_lsling_001: BP_lsling_001=Array(BM_lsling_001, LM_GI_lsling_001)
Dim BP_lsling_002: BP_lsling_002=Array(BM_lsling_002, LM_GI_lsling_002, LM_L_l25_lsling_002)
Dim BP_rsling_000: BP_rsling_000=Array(BM_rsling_000, LM_GI_rsling_000)
Dim BP_rsling_001: BP_rsling_001=Array(BM_rsling_001, LM_GI_Split_gi8_rsling_001, LM_GI_rsling_001)
Dim BP_rsling_002: BP_rsling_002=Array(BM_rsling_002, LM_GI_rsling_002, LM_L_l25a_rsling_002)
Dim BP_slingL: BP_slingL=Array(BM_slingL, LM_GI_slingL)
Dim BP_slingR: BP_slingR=Array(BM_slingR, LM_GI_slingR)
' Arrays per lighting scenario
Dim BL_GI: BL_GI=Array(LM_GI_Bumper1_Ring_001, LM_GI_Bumper1_Skirt_001, LM_GI_Bumper2_Ring_001, LM_GI_Bumper2_Skirt_001, LM_GI_DT1, LM_GI_DT10, LM_GI_DT2, LM_GI_DT3, LM_GI_DT4, LM_GI_DT5, LM_GI_DT6, LM_GI_DT7, LM_GI_DT8, LM_GI_DT9, LM_GI_Layer1, LM_GI_Layer2, LM_GI_LeafA, LM_GI_LeafB, LM_GI_LeafC, LM_GI_LeafD, LM_GI_Parts, LM_GI_Pgate, LM_GI_Playfield, LM_GI_ST11, LM_GI_SlingA_000, LM_GI_SlingA_002, LM_GI_SlingA_003, LM_GI_SlingA_004, LM_GI_SlingA_005, LM_GI_SlingB_000, LM_GI_SlingB_002, LM_GI_SlingB_003, LM_GI_SlingB_004, LM_GI_SlingB_005, LM_GI_SlingC_000, LM_GI_SlingC_002, LM_GI_SlingC_003, LM_GI_SlingC_004, LM_GI_SlingC_005, LM_GI_SlingD_000, LM_GI_SlingD_002, LM_GI_SlingD_003, LM_GI_SlingD_004, LM_GI_SlingD_005, LM_GI_TriggerTop1, LM_GI_TriggerTop2, LM_GI_TriggerTop3, LM_GI_TriggerTop4, LM_GI_lsling_000, LM_GI_lsling_001, LM_GI_lsling_002, LM_GI_rsling_000, LM_GI_rsling_001, LM_GI_rsling_002, LM_GI_slingL, LM_GI_slingR)
Dim BL_GI_Split_gi17: BL_GI_Split_gi17=Array(LM_GI_Split_gi17_Bumper1_Ring_0, LM_GI_Split_gi17_Layer1, LM_GI_Split_gi17_Parts, LM_GI_Split_gi17_Playfield, LM_GI_Split_gi17_TriggerLeftInl, LM_GI_Split_gi17_TriggerLeftInl, LM_GI_Split_gi17_TriggerLeftOut)
Dim BL_GI_Split_gi18: BL_GI_Split_gi18=Array(LM_GI_Split_gi18_Bumper1_Ring_0, LM_GI_Split_gi18_LFU, LM_GI_Split_gi18_Layer1, LM_GI_Split_gi18_Parts, LM_GI_Split_gi18_Playfield, LM_GI_Split_gi18_RF, LM_GI_Split_gi18_RFU, LM_GI_Split_gi18_TriggerLeftInl)
Dim BL_GI_Split_gi19: BL_GI_Split_gi19=Array(LM_GI_Split_gi19_Bumper1_Ring_0, LM_GI_Split_gi19_Bumper2_Ring_0, LM_GI_Split_gi19_LF, LM_GI_Split_gi19_LFU, LM_GI_Split_gi19_Layer1, LM_GI_Split_gi19_Parts, LM_GI_Split_gi19_Playfield, LM_GI_Split_gi19_RFU, LM_GI_Split_gi19_TriggerRightIn)
Dim BL_GI_Split_gi20: BL_GI_Split_gi20=Array(LM_GI_Split_gi20_Bumper2_Ring_0, LM_GI_Split_gi20_Layer1, LM_GI_Split_gi20_Parts, LM_GI_Split_gi20_Playfield, LM_GI_Split_gi20_TriggerRightIn, LM_GI_Split_gi20_TriggerRightIn, LM_GI_Split_gi20_TriggerRightOu)
Dim BL_GI_Split_gi28: BL_GI_Split_gi28=Array(LM_GI_Split_gi28_Bumper1_Ring_0, LM_GI_Split_gi28_Bumper1_Skirt_, LM_GI_Split_gi28_Layer1, LM_GI_Split_gi28_Parts, LM_GI_Split_gi28_Playfield, LM_GI_Split_gi28_SlingC_000, LM_GI_Split_gi28_SlingC_002, LM_GI_Split_gi28_SlingC_003, LM_GI_Split_gi28_SlingC_004, LM_GI_Split_gi28_SlingC_005, LM_GI_Split_gi28_TriggerLeftInl)
Dim BL_GI_Split_gi8: BL_GI_Split_gi8=Array(LM_GI_Split_gi8_Bumper2_Ring_00, LM_GI_Split_gi8_Bumper2_Skirt_0, LM_GI_Split_gi8_Layer1, LM_GI_Split_gi8_Parts, LM_GI_Split_gi8_Playfield, LM_GI_Split_gi8_SlingD_000, LM_GI_Split_gi8_SlingD_002, LM_GI_Split_gi8_SlingD_003, LM_GI_Split_gi8_SlingD_004, LM_GI_Split_gi8_SlingD_005, LM_GI_Split_gi8_TriggerRightInl, LM_GI_Split_gi8_rsling_001)
Dim BL_L_Bumper1Light: BL_L_Bumper1Light=Array(LM_L_Bumper1Light_Bumper1_Ring_, LM_L_Bumper1Light_Bumper1_Skirt, LM_L_Bumper1Light_Parts, LM_L_Bumper1Light_Playfield, LM_L_Bumper1Light_SlingC_000, LM_L_Bumper1Light_SlingC_002, LM_L_Bumper1Light_SlingC_003, LM_L_Bumper1Light_SlingC_004, LM_L_Bumper1Light_SlingC_005)
Dim BL_L_Bumper2Light: BL_L_Bumper2Light=Array(LM_L_Bumper2Light_Bumper2_Ring_, LM_L_Bumper2Light_Bumper2_Skirt, LM_L_Bumper2Light_Parts, LM_L_Bumper2Light_Playfield, LM_L_Bumper2Light_SlingD_000, LM_L_Bumper2Light_SlingD_002, LM_L_Bumper2Light_SlingD_003, LM_L_Bumper2Light_SlingD_004, LM_L_Bumper2Light_SlingD_005)
Dim BL_L_CenterSpecialLight: BL_L_CenterSpecialLight=Array(LM_L_CenterSpecialLight_Parts)
Dim BL_L_LCredit: BL_L_LCredit=Array(LM_L_LCredit_Parts)
Dim BL_L_LFive: BL_L_LFive=Array(LM_L_LFive_DT3, LM_L_LFive_DT4, LM_L_LFive_DT5, LM_L_LFive_Layer1, LM_L_LFive_Parts, LM_L_LFive_Playfield, LM_L_LFive_SlingA_000, LM_L_LFive_SlingA_002, LM_L_LFive_SlingA_003, LM_L_LFive_SlingA_004, LM_L_LFive_SlingA_005)
Dim BL_L_LSix: BL_L_LSix=Array(LM_L_LSix_DT6, LM_L_LSix_DT7, LM_L_LSix_DT8, LM_L_LSix_Layer1, LM_L_LSix_Parts, LM_L_LSix_Playfield, LM_L_LSix_SlingB_000, LM_L_LSix_SlingB_002, LM_L_LSix_SlingB_003, LM_L_LSix_SlingB_004, LM_L_LSix_SlingB_005)
Dim BL_L_LeftInlaneLight5: BL_L_LeftInlaneLight5=Array(LM_L_LeftInlaneLight5_Bumper1_R, LM_L_LeftInlaneLight5_Parts)
Dim BL_L_LeftInlaneLight5000: BL_L_LeftInlaneLight5000=Array(LM_L_LeftInlaneLight5000_Parts)
Dim BL_L_LeftSpecialLight: BL_L_LeftSpecialLight=Array(LM_L_LeftSpecialLight_Parts)
Dim BL_L_LeftTargetLight1: BL_L_LeftTargetLight1=Array(LM_L_LeftTargetLight1_DT1, LM_L_LeftTargetLight1_DT2, LM_L_LeftTargetLight1_DT3, LM_L_LeftTargetLight1_Layer1, LM_L_LeftTargetLight1_Parts, LM_L_LeftTargetLight1_Playfield, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00)
Dim BL_L_LeftTargetLight2: BL_L_LeftTargetLight2=Array(LM_L_LeftTargetLight2_DT1, LM_L_LeftTargetLight2_DT2, LM_L_LeftTargetLight2_DT3, LM_L_LeftTargetLight2_DT4, LM_L_LeftTargetLight2_Layer1, LM_L_LeftTargetLight2_Parts, LM_L_LeftTargetLight2_Playfield, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingC_00)
Dim BL_L_LeftTargetLight3: BL_L_LeftTargetLight3=Array(LM_L_LeftTargetLight3_DT2, LM_L_LeftTargetLight3_DT3, LM_L_LeftTargetLight3_DT4, LM_L_LeftTargetLight3_DT5, LM_L_LeftTargetLight3_Layer1, LM_L_LeftTargetLight3_Parts, LM_L_LeftTargetLight3_Playfield, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingC_00)
Dim BL_L_LeftTargetLight4: BL_L_LeftTargetLight4=Array(LM_L_LeftTargetLight4_DT4, LM_L_LeftTargetLight4_DT5, LM_L_LeftTargetLight4_Layer1, LM_L_LeftTargetLight4_Parts, LM_L_LeftTargetLight4_Playfield, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00)
Dim BL_L_LeftTargetLight5: BL_L_LeftTargetLight5=Array(LM_L_LeftTargetLight5_DT5, LM_L_LeftTargetLight5_Layer1, LM_L_LeftTargetLight5_Parts, LM_L_LeftTargetLight5_Playfield)
Dim BL_L_RightInlaneLight5000: BL_L_RightInlaneLight5000=Array(LM_L_RightInlaneLight5000_Parts)
Dim BL_L_RightInlaneLight6: BL_L_RightInlaneLight6=Array(LM_L_RightInlaneLight6_Bumper2_, LM_L_RightInlaneLight6_Parts)
Dim BL_L_RightSpecialLight: BL_L_RightSpecialLight=Array(LM_L_RightSpecialLight_Parts)
Dim BL_L_RightTargetLight1: BL_L_RightTargetLight1=Array(LM_L_RightTargetLight1_DT6, LM_L_RightTargetLight1_Layer1, LM_L_RightTargetLight1_Parts, LM_L_RightTargetLight1_Playfiel)
Dim BL_L_RightTargetLight2: BL_L_RightTargetLight2=Array(LM_L_RightTargetLight2_DT6, LM_L_RightTargetLight2_DT7, LM_L_RightTargetLight2_Layer1, LM_L_RightTargetLight2_Parts, LM_L_RightTargetLight2_Playfiel, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0)
Dim BL_L_RightTargetLight3: BL_L_RightTargetLight3=Array(LM_L_RightTargetLight3_DT7, LM_L_RightTargetLight3_DT8, LM_L_RightTargetLight3_DT9, LM_L_RightTargetLight3_Layer1, LM_L_RightTargetLight3_Parts, LM_L_RightTargetLight3_Playfiel, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0)
Dim BL_L_RightTargetLight4: BL_L_RightTargetLight4=Array(LM_L_RightTargetLight4_DT10, LM_L_RightTargetLight4_DT7, LM_L_RightTargetLight4_DT8, LM_L_RightTargetLight4_DT9, LM_L_RightTargetLight4_Layer1, LM_L_RightTargetLight4_Parts, LM_L_RightTargetLight4_Playfiel, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0)
Dim BL_L_RightTargetLight5: BL_L_RightTargetLight5=Array(LM_L_RightTargetLight5_DT10, LM_L_RightTargetLight5_DT9, LM_L_RightTargetLight5_Layer1, LM_L_RightTargetLight5_Parts, LM_L_RightTargetLight5_Playfiel, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingD_0)
Dim BL_L_UpperLight1: BL_L_UpperLight1=Array(LM_L_UpperLight1_Parts)
Dim BL_L_UpperLight2: BL_L_UpperLight2=Array(LM_L_UpperLight2_Layer1, LM_L_UpperLight2_Parts)
Dim BL_L_UpperLight3: BL_L_UpperLight3=Array(LM_L_UpperLight3_Layer1, LM_L_UpperLight3_Parts)
Dim BL_L_UpperLight4: BL_L_UpperLight4=Array(LM_L_UpperLight4_Parts)
Dim BL_L_l25: BL_L_l25=Array(LM_L_l25_DT1, LM_L_l25_DT10, LM_L_l25_DT2, LM_L_l25_DT3, LM_L_l25_DT4, LM_L_l25_DT5, LM_L_l25_DT6, LM_L_l25_DT7, LM_L_l25_DT8, LM_L_l25_DT9, LM_L_l25_Layer1, LM_L_l25_Layer2, LM_L_l25_Parts, LM_L_l25_Playfield, LM_L_l25_ST11, LM_L_l25_SlingA_000, LM_L_l25_SlingA_002, LM_L_l25_SlingA_003, LM_L_l25_SlingA_004, LM_L_l25_SlingA_005, LM_L_l25_SlingB_000, LM_L_l25_SlingB_002, LM_L_l25_SlingB_003, LM_L_l25_SlingB_004, LM_L_l25_SlingB_005, LM_L_l25_TriggerTop1, LM_L_l25_TriggerTop2, LM_L_l25_TriggerTop3, LM_L_l25_lsling_002)
Dim BL_L_l25a: BL_L_l25a=Array(LM_L_l25a_DT1, LM_L_l25a_DT10, LM_L_l25a_DT2, LM_L_l25a_DT3, LM_L_l25a_DT4, LM_L_l25a_DT5, LM_L_l25a_DT6, LM_L_l25a_DT7, LM_L_l25a_DT8, LM_L_l25a_DT9, LM_L_l25a_Layer1, LM_L_l25a_Layer2, LM_L_l25a_Parts, LM_L_l25a_Playfield, LM_L_l25a_ST11, LM_L_l25a_SlingA_000, LM_L_l25a_SlingA_002, LM_L_l25a_SlingA_003, LM_L_l25a_SlingA_004, LM_L_l25a_SlingA_005, LM_L_l25a_SlingB_000, LM_L_l25a_SlingB_002, LM_L_l25a_SlingB_003, LM_L_l25a_SlingB_004, LM_L_l25a_SlingB_005, LM_L_l25a_TriggerTop2, LM_L_l25a_TriggerTop3, LM_L_l25a_TriggerTop4, LM_L_l25a_rsling_002)
Dim BL_World: BL_World=Array(BM_Bumper1_Ring_001, BM_Bumper1_Skirt_001, BM_Bumper2_Ring_001, BM_Bumper2_Skirt_001, BM_DT1, BM_DT10, BM_DT2, BM_DT3, BM_DT4, BM_DT5, BM_DT6, BM_DT7, BM_DT8, BM_DT9, BM_LF, BM_LFU, BM_Layer1, BM_Layer2, BM_LeafA, BM_LeafB, BM_LeafC, BM_LeafD, BM_Parts, BM_Pgate, BM_Playfield, BM_RF, BM_RFU, BM_RailL, BM_RailR, BM_ST11, BM_SlingA_000, BM_SlingA_002, BM_SlingA_003, BM_SlingA_004, BM_SlingA_005, BM_SlingB_000, BM_SlingB_002, BM_SlingB_003, BM_SlingB_004, BM_SlingB_005, BM_SlingC_000, BM_SlingC_002, BM_SlingC_003, BM_SlingC_004, BM_SlingC_005, BM_SlingD_000, BM_SlingD_002, BM_SlingD_003, BM_SlingD_004, BM_SlingD_005, BM_TriggerLeftInlane, BM_TriggerLeftInlane2, BM_TriggerLeftOutlane, BM_TriggerRightInlane, BM_TriggerRightInlane2, BM_TriggerRightOutlane, BM_TriggerTop1, BM_TriggerTop2, BM_TriggerTop3, BM_TriggerTop4, BM_lockdownbar, BM_lsling_000, BM_lsling_001, BM_lsling_002, BM_rsling_000, BM_rsling_001, BM_rsling_002, BM_slingL, BM_slingR)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper1_Ring_001, BM_Bumper1_Skirt_001, BM_Bumper2_Ring_001, BM_Bumper2_Skirt_001, BM_DT1, BM_DT10, BM_DT2, BM_DT3, BM_DT4, BM_DT5, BM_DT6, BM_DT7, BM_DT8, BM_DT9, BM_LF, BM_LFU, BM_Layer1, BM_Layer2, BM_LeafA, BM_LeafB, BM_LeafC, BM_LeafD, BM_Parts, BM_Pgate, BM_Playfield, BM_RF, BM_RFU, BM_RailL, BM_RailR, BM_ST11, BM_SlingA_000, BM_SlingA_002, BM_SlingA_003, BM_SlingA_004, BM_SlingA_005, BM_SlingB_000, BM_SlingB_002, BM_SlingB_003, BM_SlingB_004, BM_SlingB_005, BM_SlingC_000, BM_SlingC_002, BM_SlingC_003, BM_SlingC_004, BM_SlingC_005, BM_SlingD_000, BM_SlingD_002, BM_SlingD_003, BM_SlingD_004, BM_SlingD_005, BM_TriggerLeftInlane, BM_TriggerLeftInlane2, BM_TriggerLeftOutlane, BM_TriggerRightInlane, BM_TriggerRightInlane2, BM_TriggerRightOutlane, BM_TriggerTop1, BM_TriggerTop2, BM_TriggerTop3, BM_TriggerTop4, BM_lockdownbar, BM_lsling_000, BM_lsling_001, BM_lsling_002, BM_rsling_000, BM_rsling_001, BM_rsling_002, BM_slingL, BM_slingR)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_Bumper1_Ring_001, LM_GI_Bumper1_Skirt_001, LM_GI_Bumper2_Ring_001, LM_GI_Bumper2_Skirt_001, LM_GI_DT1, LM_GI_DT10, LM_GI_DT2, LM_GI_DT3, LM_GI_DT4, LM_GI_DT5, LM_GI_DT6, LM_GI_DT7, LM_GI_DT8, LM_GI_DT9, LM_GI_Layer1, LM_GI_Layer2, LM_GI_LeafA, LM_GI_LeafB, LM_GI_LeafC, LM_GI_LeafD, LM_GI_Parts, LM_GI_Pgate, LM_GI_Playfield, LM_GI_ST11, LM_GI_SlingA_000, LM_GI_SlingA_002, LM_GI_SlingA_003, LM_GI_SlingA_004, LM_GI_SlingA_005, LM_GI_SlingB_000, LM_GI_SlingB_002, LM_GI_SlingB_003, LM_GI_SlingB_004, LM_GI_SlingB_005, LM_GI_SlingC_000, LM_GI_SlingC_002, LM_GI_SlingC_003, LM_GI_SlingC_004, LM_GI_SlingC_005, LM_GI_SlingD_000, LM_GI_SlingD_002, LM_GI_SlingD_003, LM_GI_SlingD_004, LM_GI_SlingD_005, LM_GI_TriggerTop1, LM_GI_TriggerTop2, LM_GI_TriggerTop3, LM_GI_TriggerTop4, LM_GI_lsling_000, LM_GI_lsling_001, LM_GI_lsling_002, LM_GI_rsling_000, LM_GI_rsling_001, LM_GI_rsling_002, LM_GI_slingL, LM_GI_slingR, LM_GI_Split_gi17_Bumper1_Ring_0, LM_GI_Split_gi17_Layer1, _
  LM_GI_Split_gi17_Parts, LM_GI_Split_gi17_Playfield, LM_GI_Split_gi17_TriggerLeftInl, LM_GI_Split_gi17_TriggerLeftInl, LM_GI_Split_gi17_TriggerLeftOut, LM_GI_Split_gi18_Bumper1_Ring_0, LM_GI_Split_gi18_LFU, LM_GI_Split_gi18_Layer1, LM_GI_Split_gi18_Parts, LM_GI_Split_gi18_Playfield, LM_GI_Split_gi18_RF, LM_GI_Split_gi18_RFU, LM_GI_Split_gi18_TriggerLeftInl, LM_GI_Split_gi19_Bumper1_Ring_0, LM_GI_Split_gi19_Bumper2_Ring_0, LM_GI_Split_gi19_LF, LM_GI_Split_gi19_LFU, LM_GI_Split_gi19_Layer1, LM_GI_Split_gi19_Parts, LM_GI_Split_gi19_Playfield, LM_GI_Split_gi19_RFU, LM_GI_Split_gi19_TriggerRightIn, LM_GI_Split_gi20_Bumper2_Ring_0, LM_GI_Split_gi20_Layer1, LM_GI_Split_gi20_Parts, LM_GI_Split_gi20_Playfield, LM_GI_Split_gi20_TriggerRightIn, LM_GI_Split_gi20_TriggerRightIn, LM_GI_Split_gi20_TriggerRightOu, LM_GI_Split_gi28_Bumper1_Ring_0, LM_GI_Split_gi28_Bumper1_Skirt_, LM_GI_Split_gi28_Layer1, LM_GI_Split_gi28_Parts, LM_GI_Split_gi28_Playfield, LM_GI_Split_gi28_SlingC_000, LM_GI_Split_gi28_SlingC_002, _
  LM_GI_Split_gi28_SlingC_003, LM_GI_Split_gi28_SlingC_004, LM_GI_Split_gi28_SlingC_005, LM_GI_Split_gi28_TriggerLeftInl, LM_GI_Split_gi8_Bumper2_Ring_00, LM_GI_Split_gi8_Bumper2_Skirt_0, LM_GI_Split_gi8_Layer1, LM_GI_Split_gi8_Parts, LM_GI_Split_gi8_Playfield, LM_GI_Split_gi8_SlingD_000, LM_GI_Split_gi8_SlingD_002, LM_GI_Split_gi8_SlingD_003, LM_GI_Split_gi8_SlingD_004, LM_GI_Split_gi8_SlingD_005, LM_GI_Split_gi8_TriggerRightInl, LM_GI_Split_gi8_rsling_001, LM_L_Bumper1Light_Bumper1_Ring_, LM_L_Bumper1Light_Bumper1_Skirt, LM_L_Bumper1Light_Parts, LM_L_Bumper1Light_Playfield, LM_L_Bumper1Light_SlingC_000, LM_L_Bumper1Light_SlingC_002, LM_L_Bumper1Light_SlingC_003, LM_L_Bumper1Light_SlingC_004, LM_L_Bumper1Light_SlingC_005, LM_L_Bumper2Light_Bumper2_Ring_, LM_L_Bumper2Light_Bumper2_Skirt, LM_L_Bumper2Light_Parts, LM_L_Bumper2Light_Playfield, LM_L_Bumper2Light_SlingD_000, LM_L_Bumper2Light_SlingD_002, LM_L_Bumper2Light_SlingD_003, LM_L_Bumper2Light_SlingD_004, LM_L_Bumper2Light_SlingD_005, _
  LM_L_CenterSpecialLight_Parts, LM_L_LCredit_Parts, LM_L_LFive_DT3, LM_L_LFive_DT4, LM_L_LFive_DT5, LM_L_LFive_Layer1, LM_L_LFive_Parts, LM_L_LFive_Playfield, LM_L_LFive_SlingA_000, LM_L_LFive_SlingA_002, LM_L_LFive_SlingA_003, LM_L_LFive_SlingA_004, LM_L_LFive_SlingA_005, LM_L_LSix_DT6, LM_L_LSix_DT7, LM_L_LSix_DT8, LM_L_LSix_Layer1, LM_L_LSix_Parts, LM_L_LSix_Playfield, LM_L_LSix_SlingB_000, LM_L_LSix_SlingB_002, LM_L_LSix_SlingB_003, LM_L_LSix_SlingB_004, LM_L_LSix_SlingB_005, LM_L_LeftInlaneLight5_Bumper1_R, LM_L_LeftInlaneLight5_Parts, LM_L_LeftInlaneLight5000_Parts, LM_L_LeftSpecialLight_Parts, LM_L_LeftTargetLight1_DT1, LM_L_LeftTargetLight1_DT2, LM_L_LeftTargetLight1_DT3, LM_L_LeftTargetLight1_Layer1, LM_L_LeftTargetLight1_Parts, LM_L_LeftTargetLight1_Playfield, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight2_DT1, LM_L_LeftTargetLight2_DT2, _
  LM_L_LeftTargetLight2_DT3, LM_L_LeftTargetLight2_DT4, LM_L_LeftTargetLight2_Layer1, LM_L_LeftTargetLight2_Parts, LM_L_LeftTargetLight2_Playfield, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingC_00, LM_L_LeftTargetLight3_DT2, LM_L_LeftTargetLight3_DT3, LM_L_LeftTargetLight3_DT4, LM_L_LeftTargetLight3_DT5, LM_L_LeftTargetLight3_Layer1, LM_L_LeftTargetLight3_Parts, LM_L_LeftTargetLight3_Playfield, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingC_00, LM_L_LeftTargetLight4_DT4, LM_L_LeftTargetLight4_DT5, LM_L_LeftTargetLight4_Layer1, LM_L_LeftTargetLight4_Parts, LM_L_LeftTargetLight4_Playfield, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, _
  LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight5_DT5, LM_L_LeftTargetLight5_Layer1, LM_L_LeftTargetLight5_Parts, LM_L_LeftTargetLight5_Playfield, LM_L_RightInlaneLight5000_Parts, LM_L_RightInlaneLight6_Bumper2_, LM_L_RightInlaneLight6_Parts, LM_L_RightSpecialLight_Parts, LM_L_RightTargetLight1_DT6, LM_L_RightTargetLight1_Layer1, LM_L_RightTargetLight1_Parts, LM_L_RightTargetLight1_Playfiel, LM_L_RightTargetLight2_DT6, LM_L_RightTargetLight2_DT7, LM_L_RightTargetLight2_Layer1, LM_L_RightTargetLight2_Parts, LM_L_RightTargetLight2_Playfiel, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight3_DT7, LM_L_RightTargetLight3_DT8, LM_L_RightTargetLight3_DT9, LM_L_RightTargetLight3_Layer1, LM_L_RightTargetLight3_Parts, LM_L_RightTargetLight3_Playfiel, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, _
  LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight4_DT10, LM_L_RightTargetLight4_DT7, LM_L_RightTargetLight4_DT8, LM_L_RightTargetLight4_DT9, LM_L_RightTargetLight4_Layer1, LM_L_RightTargetLight4_Parts, LM_L_RightTargetLight4_Playfiel, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight5_DT10, LM_L_RightTargetLight5_DT9, LM_L_RightTargetLight5_Layer1, LM_L_RightTargetLight5_Parts, LM_L_RightTargetLight5_Playfiel, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingD_0, LM_L_UpperLight1_Parts, LM_L_UpperLight2_Layer1, LM_L_UpperLight2_Parts, LM_L_UpperLight3_Layer1, LM_L_UpperLight3_Parts, LM_L_UpperLight4_Parts, LM_L_l25_DT1, LM_L_l25_DT10, LM_L_l25_DT2, LM_L_l25_DT3, LM_L_l25_DT4, LM_L_l25_DT5, LM_L_l25_DT6, LM_L_l25_DT7, _
  LM_L_l25_DT8, LM_L_l25_DT9, LM_L_l25_Layer1, LM_L_l25_Layer2, LM_L_l25_Parts, LM_L_l25_Playfield, LM_L_l25_ST11, LM_L_l25_SlingA_000, LM_L_l25_SlingA_002, LM_L_l25_SlingA_003, LM_L_l25_SlingA_004, LM_L_l25_SlingA_005, LM_L_l25_SlingB_000, LM_L_l25_SlingB_002, LM_L_l25_SlingB_003, LM_L_l25_SlingB_004, LM_L_l25_SlingB_005, LM_L_l25_TriggerTop1, LM_L_l25_TriggerTop2, LM_L_l25_TriggerTop3, LM_L_l25_lsling_002, LM_L_l25a_DT1, LM_L_l25a_DT10, LM_L_l25a_DT2, LM_L_l25a_DT3, LM_L_l25a_DT4, LM_L_l25a_DT5, LM_L_l25a_DT6, LM_L_l25a_DT7, LM_L_l25a_DT8, LM_L_l25a_DT9, LM_L_l25a_Layer1, LM_L_l25a_Layer2, LM_L_l25a_Parts, LM_L_l25a_Playfield, LM_L_l25a_ST11, LM_L_l25a_SlingA_000, LM_L_l25a_SlingA_002, LM_L_l25a_SlingA_003, LM_L_l25a_SlingA_004, LM_L_l25a_SlingA_005, LM_L_l25a_SlingB_000, LM_L_l25a_SlingB_002, LM_L_l25a_SlingB_003, LM_L_l25a_SlingB_004, LM_L_l25a_SlingB_005, LM_L_l25a_TriggerTop2, LM_L_l25a_TriggerTop3, LM_L_l25a_TriggerTop4, LM_L_l25a_rsling_002)
Dim BG_All: BG_All=Array(BM_Bumper1_Ring_001, BM_Bumper1_Skirt_001, BM_Bumper2_Ring_001, BM_Bumper2_Skirt_001, BM_DT1, BM_DT10, BM_DT2, BM_DT3, BM_DT4, BM_DT5, BM_DT6, BM_DT7, BM_DT8, BM_DT9, BM_LF, BM_LFU, BM_Layer1, BM_Layer2, BM_LeafA, BM_LeafB, BM_LeafC, BM_LeafD, BM_Parts, BM_Pgate, BM_Playfield, BM_RF, BM_RFU, BM_RailL, BM_RailR, BM_ST11, BM_SlingA_000, BM_SlingA_002, BM_SlingA_003, BM_SlingA_004, BM_SlingA_005, BM_SlingB_000, BM_SlingB_002, BM_SlingB_003, BM_SlingB_004, BM_SlingB_005, BM_SlingC_000, BM_SlingC_002, BM_SlingC_003, BM_SlingC_004, BM_SlingC_005, BM_SlingD_000, BM_SlingD_002, BM_SlingD_003, BM_SlingD_004, BM_SlingD_005, BM_TriggerLeftInlane, BM_TriggerLeftInlane2, BM_TriggerLeftOutlane, BM_TriggerRightInlane, BM_TriggerRightInlane2, BM_TriggerRightOutlane, BM_TriggerTop1, BM_TriggerTop2, BM_TriggerTop3, BM_TriggerTop4, BM_lockdownbar, BM_lsling_000, BM_lsling_001, BM_lsling_002, BM_rsling_000, BM_rsling_001, BM_rsling_002, BM_slingL, BM_slingR, LM_GI_Bumper1_Ring_001, _
  LM_GI_Bumper1_Skirt_001, LM_GI_Bumper2_Ring_001, LM_GI_Bumper2_Skirt_001, LM_GI_DT1, LM_GI_DT10, LM_GI_DT2, LM_GI_DT3, LM_GI_DT4, LM_GI_DT5, LM_GI_DT6, LM_GI_DT7, LM_GI_DT8, LM_GI_DT9, LM_GI_Layer1, LM_GI_Layer2, LM_GI_LeafA, LM_GI_LeafB, LM_GI_LeafC, LM_GI_LeafD, LM_GI_Parts, LM_GI_Pgate, LM_GI_Playfield, LM_GI_ST11, LM_GI_SlingA_000, LM_GI_SlingA_002, LM_GI_SlingA_003, LM_GI_SlingA_004, LM_GI_SlingA_005, LM_GI_SlingB_000, LM_GI_SlingB_002, LM_GI_SlingB_003, LM_GI_SlingB_004, LM_GI_SlingB_005, LM_GI_SlingC_000, LM_GI_SlingC_002, LM_GI_SlingC_003, LM_GI_SlingC_004, LM_GI_SlingC_005, LM_GI_SlingD_000, LM_GI_SlingD_002, LM_GI_SlingD_003, LM_GI_SlingD_004, LM_GI_SlingD_005, LM_GI_TriggerTop1, LM_GI_TriggerTop2, LM_GI_TriggerTop3, LM_GI_TriggerTop4, LM_GI_lsling_000, LM_GI_lsling_001, LM_GI_lsling_002, LM_GI_rsling_000, LM_GI_rsling_001, LM_GI_rsling_002, LM_GI_slingL, LM_GI_slingR, LM_GI_Split_gi17_Bumper1_Ring_0, LM_GI_Split_gi17_Layer1, LM_GI_Split_gi17_Parts, LM_GI_Split_gi17_Playfield, _
  LM_GI_Split_gi17_TriggerLeftInl, LM_GI_Split_gi17_TriggerLeftInl, LM_GI_Split_gi17_TriggerLeftOut, LM_GI_Split_gi18_Bumper1_Ring_0, LM_GI_Split_gi18_LFU, LM_GI_Split_gi18_Layer1, LM_GI_Split_gi18_Parts, LM_GI_Split_gi18_Playfield, LM_GI_Split_gi18_RF, LM_GI_Split_gi18_RFU, LM_GI_Split_gi18_TriggerLeftInl, LM_GI_Split_gi19_Bumper1_Ring_0, LM_GI_Split_gi19_Bumper2_Ring_0, LM_GI_Split_gi19_LF, LM_GI_Split_gi19_LFU, LM_GI_Split_gi19_Layer1, LM_GI_Split_gi19_Parts, LM_GI_Split_gi19_Playfield, LM_GI_Split_gi19_RFU, LM_GI_Split_gi19_TriggerRightIn, LM_GI_Split_gi20_Bumper2_Ring_0, LM_GI_Split_gi20_Layer1, LM_GI_Split_gi20_Parts, LM_GI_Split_gi20_Playfield, LM_GI_Split_gi20_TriggerRightIn, LM_GI_Split_gi20_TriggerRightIn, LM_GI_Split_gi20_TriggerRightOu, LM_GI_Split_gi28_Bumper1_Ring_0, LM_GI_Split_gi28_Bumper1_Skirt_, LM_GI_Split_gi28_Layer1, LM_GI_Split_gi28_Parts, LM_GI_Split_gi28_Playfield, LM_GI_Split_gi28_SlingC_000, LM_GI_Split_gi28_SlingC_002, LM_GI_Split_gi28_SlingC_003, LM_GI_Split_gi28_SlingC_004, _
  LM_GI_Split_gi28_SlingC_005, LM_GI_Split_gi28_TriggerLeftInl, LM_GI_Split_gi8_Bumper2_Ring_00, LM_GI_Split_gi8_Bumper2_Skirt_0, LM_GI_Split_gi8_Layer1, LM_GI_Split_gi8_Parts, LM_GI_Split_gi8_Playfield, LM_GI_Split_gi8_SlingD_000, LM_GI_Split_gi8_SlingD_002, LM_GI_Split_gi8_SlingD_003, LM_GI_Split_gi8_SlingD_004, LM_GI_Split_gi8_SlingD_005, LM_GI_Split_gi8_TriggerRightInl, LM_GI_Split_gi8_rsling_001, LM_L_Bumper1Light_Bumper1_Ring_, LM_L_Bumper1Light_Bumper1_Skirt, LM_L_Bumper1Light_Parts, LM_L_Bumper1Light_Playfield, LM_L_Bumper1Light_SlingC_000, LM_L_Bumper1Light_SlingC_002, LM_L_Bumper1Light_SlingC_003, LM_L_Bumper1Light_SlingC_004, LM_L_Bumper1Light_SlingC_005, LM_L_Bumper2Light_Bumper2_Ring_, LM_L_Bumper2Light_Bumper2_Skirt, LM_L_Bumper2Light_Parts, LM_L_Bumper2Light_Playfield, LM_L_Bumper2Light_SlingD_000, LM_L_Bumper2Light_SlingD_002, LM_L_Bumper2Light_SlingD_003, LM_L_Bumper2Light_SlingD_004, LM_L_Bumper2Light_SlingD_005, LM_L_CenterSpecialLight_Parts, LM_L_LCredit_Parts, LM_L_LFive_DT3, _
  LM_L_LFive_DT4, LM_L_LFive_DT5, LM_L_LFive_Layer1, LM_L_LFive_Parts, LM_L_LFive_Playfield, LM_L_LFive_SlingA_000, LM_L_LFive_SlingA_002, LM_L_LFive_SlingA_003, LM_L_LFive_SlingA_004, LM_L_LFive_SlingA_005, LM_L_LSix_DT6, LM_L_LSix_DT7, LM_L_LSix_DT8, LM_L_LSix_Layer1, LM_L_LSix_Parts, LM_L_LSix_Playfield, LM_L_LSix_SlingB_000, LM_L_LSix_SlingB_002, LM_L_LSix_SlingB_003, LM_L_LSix_SlingB_004, LM_L_LSix_SlingB_005, LM_L_LeftInlaneLight5_Bumper1_R, LM_L_LeftInlaneLight5_Parts, LM_L_LeftInlaneLight5000_Parts, LM_L_LeftSpecialLight_Parts, LM_L_LeftTargetLight1_DT1, LM_L_LeftTargetLight1_DT2, LM_L_LeftTargetLight1_DT3, LM_L_LeftTargetLight1_Layer1, LM_L_LeftTargetLight1_Parts, LM_L_LeftTargetLight1_Playfield, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight1_SlingA_00, LM_L_LeftTargetLight2_DT1, LM_L_LeftTargetLight2_DT2, LM_L_LeftTargetLight2_DT3, LM_L_LeftTargetLight2_DT4, LM_L_LeftTargetLight2_Layer1, _
  LM_L_LeftTargetLight2_Parts, LM_L_LeftTargetLight2_Playfield, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingA_00, LM_L_LeftTargetLight2_SlingC_00, LM_L_LeftTargetLight3_DT2, LM_L_LeftTargetLight3_DT3, LM_L_LeftTargetLight3_DT4, LM_L_LeftTargetLight3_DT5, LM_L_LeftTargetLight3_Layer1, LM_L_LeftTargetLight3_Parts, LM_L_LeftTargetLight3_Playfield, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingA_00, LM_L_LeftTargetLight3_SlingC_00, LM_L_LeftTargetLight4_DT4, LM_L_LeftTargetLight4_DT5, LM_L_LeftTargetLight4_Layer1, LM_L_LeftTargetLight4_Parts, LM_L_LeftTargetLight4_Playfield, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight4_SlingA_00, LM_L_LeftTargetLight5_DT5, _
  LM_L_LeftTargetLight5_Layer1, LM_L_LeftTargetLight5_Parts, LM_L_LeftTargetLight5_Playfield, LM_L_RightInlaneLight5000_Parts, LM_L_RightInlaneLight6_Bumper2_, LM_L_RightInlaneLight6_Parts, LM_L_RightSpecialLight_Parts, LM_L_RightTargetLight1_DT6, LM_L_RightTargetLight1_Layer1, LM_L_RightTargetLight1_Parts, LM_L_RightTargetLight1_Playfiel, LM_L_RightTargetLight2_DT6, LM_L_RightTargetLight2_DT7, LM_L_RightTargetLight2_Layer1, LM_L_RightTargetLight2_Parts, LM_L_RightTargetLight2_Playfiel, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight2_SlingB_0, LM_L_RightTargetLight3_DT7, LM_L_RightTargetLight3_DT8, LM_L_RightTargetLight3_DT9, LM_L_RightTargetLight3_Layer1, LM_L_RightTargetLight3_Parts, LM_L_RightTargetLight3_Playfiel, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, LM_L_RightTargetLight3_SlingB_0, _
  LM_L_RightTargetLight4_DT10, LM_L_RightTargetLight4_DT7, LM_L_RightTargetLight4_DT8, LM_L_RightTargetLight4_DT9, LM_L_RightTargetLight4_Layer1, LM_L_RightTargetLight4_Parts, LM_L_RightTargetLight4_Playfiel, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight4_SlingB_0, LM_L_RightTargetLight5_DT10, LM_L_RightTargetLight5_DT9, LM_L_RightTargetLight5_Layer1, LM_L_RightTargetLight5_Parts, LM_L_RightTargetLight5_Playfiel, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingB_0, LM_L_RightTargetLight5_SlingD_0, LM_L_UpperLight1_Parts, LM_L_UpperLight2_Layer1, LM_L_UpperLight2_Parts, LM_L_UpperLight3_Layer1, LM_L_UpperLight3_Parts, LM_L_UpperLight4_Parts, LM_L_l25_DT1, LM_L_l25_DT10, LM_L_l25_DT2, LM_L_l25_DT3, LM_L_l25_DT4, LM_L_l25_DT5, LM_L_l25_DT6, LM_L_l25_DT7, LM_L_l25_DT8, LM_L_l25_DT9, LM_L_l25_Layer1, _
  LM_L_l25_Layer2, LM_L_l25_Parts, LM_L_l25_Playfield, LM_L_l25_ST11, LM_L_l25_SlingA_000, LM_L_l25_SlingA_002, LM_L_l25_SlingA_003, LM_L_l25_SlingA_004, LM_L_l25_SlingA_005, LM_L_l25_SlingB_000, LM_L_l25_SlingB_002, LM_L_l25_SlingB_003, LM_L_l25_SlingB_004, LM_L_l25_SlingB_005, LM_L_l25_TriggerTop1, LM_L_l25_TriggerTop2, LM_L_l25_TriggerTop3, LM_L_l25_lsling_002, LM_L_l25a_DT1, LM_L_l25a_DT10, LM_L_l25a_DT2, LM_L_l25a_DT3, LM_L_l25a_DT4, LM_L_l25a_DT5, LM_L_l25a_DT6, LM_L_l25a_DT7, LM_L_l25a_DT8, LM_L_l25a_DT9, LM_L_l25a_Layer1, LM_L_l25a_Layer2, LM_L_l25a_Parts, LM_L_l25a_Playfield, LM_L_l25a_ST11, LM_L_l25a_SlingA_000, LM_L_l25a_SlingA_002, LM_L_l25a_SlingA_003, LM_L_l25a_SlingA_004, LM_L_l25a_SlingA_005, LM_L_l25a_SlingB_000, LM_L_l25a_SlingB_002, LM_L_l25a_SlingB_003, LM_L_l25a_SlingB_004, LM_L_l25a_SlingB_005, LM_L_l25a_TriggerTop2, LM_L_l25a_TriggerTop3, LM_L_l25a_TriggerTop4, LM_L_l25a_rsling_002)
' VLM  Arrays - End


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim BallsPerGameOption : BallsPerGameOption = 0 ' 0: 5 balls, 1: 3 balls
Dim CabinetMode : CabinetMode = 0 ' 0: Off, 1: On
Dim ChimesOn : ChimesOn = 1
Dim VRRoomChoice : VRRoomChoice = 0
Dim VRRoom
Dim Lane2Or3SpotsBoth : Lane2Or3SpotsBoth = 0
TiltEndsGame = 0 ' 0: Ball in play only: 1: Ball in play + 1 ball
BallsPerGame = 5
NoveltyMode = 0 ' 0: Off, 1: On
ReplayLevel = 1

Dim VRPreview : VRPreview = False
Dim VRMode : VRMode = False

If RenderingMode = 2 OR VRPreview Then
  VRMode = True
End If


' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 0, 11, 1, 0, 0, _
    Array("Normal", "Vibrant", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 0 Then Table1.ColorGradeImage = "colorgradelut256x16_1to1"
  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat100"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat90"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat80"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat70"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat60"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat40"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat30"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat20"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat10"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat00"

  BallsPerGameOption = Table1.Option("Balls Per Game", 0, 1, 1, 0, 0, Array("5 Balls", "3 Balls"))
  If BallsPerGameOption = 0 Then BallsPerGame = 5
  If BallsPerGameOption = 1 Then BallsPerGame = 3

  TiltEndsGame = Table1.Option("Tilt Penalty", 0, 1, 1, 0, 0, Array("Ball in play only", "Ball in play + 1 ball"))

  Lane2or3SpotsBoth = Table1.Option("Lane 2/3 Setting", 0, 1, 1, 0, 0, Array("Normal", "Spot Both"))

  NoveltyMode = Table1.Option("Novelty Mode", 0, 1, 1, 1, 0, Array("Off (Add-a-ball)", "On"))

  ReplayLevel = Table1.Option("Add-a-ball Level", 1, 8, 1, 1, 0)
  RefreshReplayCard

  ChimesOn = Table1.Option("Bells or Chimes", 0, 1, 1, 1, 0, Array("Bells", "Chimes"))

  CabinetMode = Table1.Option("Cabinet Mode", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetCabinetMode

  ' Playfield Reflections
  ReflectionsQuality = Table1.Option("Reflections Quality", 0, 1, 1, 0, 0, Array("Clean", "Rough"))
  Select Case ReflectionsQuality
    Case 0: playfield_mesh.ReflectionProbe = "Playfield Reflections": BM_Playfield.ReflectionProbe = "Playfield Reflections"
    Case 1: playfield_mesh.ReflectionProbe = "Playfield Reflections Rough": BM_Playfield.ReflectionProbe = "Playfield Reflections Rough"
  End Select

    ' Toggle Reflections Scope
  ReflectionsScope = Table1.Option("Reflections Scope", 0, 2, 1, 1, 0, Array("None", "Partial", "Full"))
  ReflectionToggle(ReflectionsScope)

  'VR Room
  VRRoomChoice = Table1.Option("VR Room", 0, 1, 1, 0, 0, Array("MEGA", "Minimal"))
  If VRMode Then VRRoom = VRRoomChoice Else VRRoom = 0
  SetupRoom

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  InstructCard1.imagea="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(TiltEndsGame,0)+FormatNumber(NoveltyMode,0)

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Function ReflectionToggle(state)
  ' Full reflections object subset
  Dim ReflObjectsArray: ReflObjectsArray = Array(BP_Parts, BP_LF, BP_LFU, BP_RF, BP_RFU, BP_Bumper1_Ring_001, BP_Bumper1_Skirt_001, BP_Bumper2_Ring_001, BP_Bumper2_Skirt_001, BP_DT1, BP_DT2, BP_DT3, BP_DT4, BP_DT5, BP_DT6, BP_DT7, BP_DT8, BP_DT9, BP_DT10, BP_Layer1, BP_Layer2, BP_ST11)
  ' Partial reflections object subset
  Dim ReflPrtObjsArray: ReflPrtObjsArray = Array(BM_Parts, BM_LF, BM_LFU, BM_RF, BM_RFU, BM_Bumper1_Ring_001, BM_Bumper1_Skirt_001, BM_Bumper2_Ring_001, BM_Bumper2_Skirt_001, BM_DT1, BM_DT2, BM_DT3, BM_DT4, BM_DT5, BM_DT6, BM_DT7, BM_DT8, BM_DT9, BM_DT10, BM_Layer1, BM_Layer2, BM_ST11)
  Dim IBP, BP, v1, v2
  Select Case state
    Case 0: v1 = false: v2 = false  'No reflections
    Case 1: v1 = false: v2 = true   'Only reflect static prims
    Case 2: v1 = true:  v2 = true   'Reflect everything (that matters as defined in ReflObjectsArray)
  End Select

  For Each IBP in ReflObjectsArray
    For Each BP in IBP
      BP.ReflectionEnabled = v1
    Next
  Next

  For Each BP in ReflPrtObjsArray
    BP.ReflectionEnabled = v2
  Next

  BM_Parts.ReflectionEnabled = v2
  BM_Playfield.ReflectionEnabled = False
End Function

Sub SetCabinetMode()
  Dim BP,VRThings
  If VRMode Then
    For Each BP in BP_lockdownbar : BP.visible = 0 : Next
    For Each BP in BP_RailL : BP.visible = 0 : Next
    For Each BP in BP_RailR : BP.visible = 0 : Next
  Else
    If CabinetMode = 1 Then
      For Each BP in BP_lockdownbar : BP.visible = 0 : Next
      For Each BP in BP_RailL : BP.visible = 0 : Next
      For Each BP in BP_RailR : BP.visible = 0 : Next
    Else
      For Each BP in BP_lockdownbar : BP.visible = 1 : Next
      For Each BP in BP_RailL : BP.visible = 1 : Next
      For Each BP in BP_RailR : BP.visible = 1 : Next
    End If
  End If
End Sub

Sub SetupRoom
  Dim VRThings
  If VRMode Then
    'For each VRThings in VRBackglass:VRThings.visible = 1:Next
    If VRRoom = 0 Then
      for each VRThings in VR_Minimal:VRThings.visible = 0:Next
      for each VRThings in VR_Mega:VRThings.visible = 1:Next
    End If
    If VRRoom = 1 Then
      for each VRThings in VR_Minimal:VRThings.visible = 1:Next
      for each VRThings in VR_Mega:VRThings.visible = 0:Next
    End If
  Else
    For each VRThings in VR_Minimal:VRThings.visible = 0:Next
    For each VRThings in VR_Mega:VRThings.visible = 0:Next
    'For each VRThings in VRBackglass:VRThings.visible = 0:Next
  End if
End Sub



'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 1000       'X position of punger lane left
Const PLRight = 1060      'X position of punger lane right
Const PLTop = 1225        'Y position of punger lane top
Const PLBottom = 1900       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,PLLeft,PLBottom,PLLeft,PLTop,PLRight,PLTop,PLRight,PLBottom) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-PLBottom))
    End If
    ' Assign color
    b_r = Int(d_w)
    b_g = Int(d_w)
    b_b = Int(d_w)
    If b_r > 255 Then b_r = 255
    If b_g > 255 Then b_g = 255
    If b_b > 255 Then b_b = 255
    gBOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
    'debug.print "--- ball.color level="&b_r
  Next
End Sub


'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 245 + 10)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next
End Sub

Dim SavedMtlColorArray
SaveMtlColors
Sub SaveMtlColors
  ReDim SavedMtlColorArray(UBound(RoomBrightnessMtlArray))
  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    SaveMaterialBaseColor RoomBrightnessMtlArray(i), i
  Next
End Sub

Sub SaveMaterialBaseColor(name, idx)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  SavedMtlColorArray(idx) = round(base,0)
End Sub


Sub ModulateMaterialBaseColor(name, idx, val)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  Dim red, green, blue, saved_base, new_base

  'First get the existing material properties
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  'Get saved color
  saved_base = SavedMtlColorArray(idx)

  'Next extract the r,g,b values from the base color
  red = saved_base And &HFF
  green = (saved_base \ &H100) And &HFF
  blue = (saved_base \ &H10000) And &HFF
  'msgbox red & " " & green & " " & blue

  'Create new color scaled down by 'val', and update the material
  new_base = RGB(red*val, green*val, blue*val)
  UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, new_base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub



'*******************************************
' ZTIM: Timers
'*******************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  'Add animation stuff here
  RollingUpdate       'update rolling sounds
  DoDTAnim        'handle drop target animations
  DoSTAnim        'handle stand up target animations
  BSUpdate
  CheckWOW
  UpdateStandupTargets
  UpdateDropTargets

End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

LoadCoreFiles
Sub LoadCoreFiles
  On Error Resume Next
  ExecuteGlobal GetTextFile("core.vbs") 'TODO: drop-in replacement for vpmTimer (maybe vpwQueueManager) and cvpmDictionary (Scripting.Dictionary) to remove core.vbs dependency
  If Err Then MsgBox "Can't open core.vbs"
  On Error GoTo 0
End Sub

Sub CheckWOW()
  If VRMode Then
    Select Case WOWCounter:
      Case 0:
        VR_L51.visible = 0
        VR_L52.visible = 0
        VR_L53.visible = 0
        VR_L54.visible = 0
        VR_L55.visible = 0
      Case 1:
        VR_L51.visible = 1
        VR_L52.visible = 0
        VR_L53.visible = 0
        VR_L54.visible = 0
        VR_L55.visible = 0
      Case 2:
        VR_L51.visible = 1
        VR_L52.visible = 1
        VR_L53.visible = 0
        VR_L54.visible = 0
        VR_L55.visible = 0
      Case 3:
        VR_L51.visible = 1
        VR_L52.visible = 1
        VR_L53.visible = 1
        VR_L54.visible = 0
        VR_L55.visible = 0
      Case 4:
        VR_L51.visible = 1
        VR_L52.visible = 1
        VR_L53.visible = 1
        VR_L54.visible = 1
        VR_L55.visible = 0
      Case 5:
        VR_L51.visible = 1
        VR_L52.visible = 1
        VR_L53.visible = 1
        VR_L54.visible = 1
        VR_L55.visible = 1
    End Select
  End If
End Sub

Dim TOBall, gBOT

Sub Table1_Init()
  If Table1.ShowDT = false then
    For each obj in DesktopCrap
      obj.visible=False
    next
  End If


  LoadEM
  LoadLMEMConfig2

  SetupReplayTables
  PlasticsOff
  BumpersOff
  HighScore=0
  MotorRunning=0
  HighScoreReward=1
  Credits=0
  BonusSpecialThreshold=1
  loadhs
  if HighScore=0 then HighScore=50000


  TableTilted=false


    Set TOBall = Drain.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBot = Array(TOBall)


  Match=int(Rnd*10)*10
  MatchReel.SetValue((Match/10)+1)

  CanPlayReel.SetValue(0)
  BallInPlayReel.SetValue(0)

  For each obj in PlayerHuds
    obj.SetValue(0)
  next
  For each obj in PlayerScoresOn
    obj.ResetToZero
  next


  For each obj in PlayerScores
    obj.ResetToZero
  next

  GameOverReel.SetValue(1)
  TiltReel.SetValue(1)

  If VRMode Then
    VR_GO.visible = 1
    VR_Tilt.visible = 1
  End If

  for each obj in Bonus
    obj.state=0
  next
  for each obj in NonNumberLights
    obj.state=0
  next
  for each obj in NumberLights
    obj.state=0
  next
  for each obj in LeftTargetLights
    obj.state=0
  next
  for each obj in RightTargetLights
    obj.state=0
  next


  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)

  BonusCounter=0
  HoleCounter=0
  AlternatingRelayCounter=1
  ABCCompleted=false
    bgpos=6
  kgpos=0


  dooralreadyopen=0
  kgdooralreadyopen=0


  Bumper1Light.state=0
  Bumper2Light.state=0

  TargetSpecialLit = 0
  Points210counter=0
  Points500counter=0
  Points1000counter=0
  Points2000counter=0

  BonusBooster=3
  BonusBoosterCounter=0
  Players=0
  RotatorTemp=1
  InProgress=false

  ScoreText.text=HighScore

  If B2SOn Then
    Controller.B2SSetMatch Match
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0

    Controller.B2SSetScorePlayer 2,HSScore(1)
    Controller.B2SSetTilt 0
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
  End If
  HSReel.SetValue(HSScore(1))
  If Credits > 0 then
    LCredit.state = 1
  Else
    LCredit.state = 0
  End If
  for i=1 to 4
    player=i
  If B2SOn Then
    Controller.B2SSetScorePlayer 1, 0
  End If
  next
  bump1=1
  bump2=1
  bump3=1
  InitPauser5.enabled=true
  If Credits > 0 Then DOF 126, 1

  SetupRoom
End Sub

Sub Table1_exit()
  savehs
  SaveLMEMConfig
  SaveLMEMConfig2
  If B2SOn Then Controller.Stop
end sub

const ReflipAngle = 20


Sub Table1_KeyDown(ByVal keycode)

  ' GNMOD
  if EnteringInitials then
    CollectInitials(keycode)
    exit sub
  end if

  If keycode = PlungerKey Then
    Plunger.PullBack
        SoundPlungerPull
    PlungerPulled = 1
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    FlipperActivate LeftFlipper, LFPress
        LF.Fire
        If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
            RandomSoundReflipUpLeft LeftFlipper
        Else
            SoundFlipperUpAttackLeft LeftFlipper
            RandomSoundFlipperUpLeft LeftFlipper
        End If
        PlaySound "BuzzL", -1, .01, AudioPan(LeftFlipper), .05,0, 0, 1, AudioFade(LeftFlipper)
    DOF 101, 1
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    RF.Fire
        If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
            RandomSoundReflipUpRight RightFlipper
        Else
            SoundFlipperUpAttackRight RightFlipper
            RandomSoundFlipperUpRight RightFlipper
        End If
        PlaySound "Buzz", -1, .01, AudioPan(RightFlipper), .05,0, 0, 1, AudioFade(RightFlipper)
    DOF 102, 1
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
        SoundNudgeLeft
    TiltIt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
        SoundNudgeRight
    TiltIt
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
        SoundNudgeCenter
    TiltIt
  End If

  If keycode = MechanicalTilt Then
    TiltCount=2
    TiltIt
  End If

  If keycode = AddCreditKey or keycode = 4 then
    If B2SOn Then
      Controller.B2SSetScorePlayer 2,HSScore(1)
    End If
    HSReel.SetValue(HSScore(1))
    LCredit.state = 1
    Select Case Int(rnd*3)
            Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
            Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
            Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
        End Select
    AddSpecial2
  end if

    if keycode = 5 then
        Select Case Int(rnd*3)
            Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
            Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
            Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
        End Select
    AddSpecial2
    LCredit.state = 1
    keycode= StartGameKey
  end if

  if keycode = StartGameKey Then
    SoundStartButton
    VR_Cab_StartButton.y = VR_Cab_StartButton.y - 4
  End If

  if keycode=StartGameKey and Credits > 0 and InProgress=false and Players=0 then
    Credits=Credits-1
    If Credits < 1 Then DOF 126, 0
    If Credits > 0 then
      LCredit.state = 1
    Else
      LCredit.state = 0
    End If
    CreditsReel.SetValue(Credits)
    Players=1
    CanPlayReel.SetValue(Players)
    MatchReel.SetValue(0)
    GameOverReel.SetValue(0)
    RolloverReel.SetValue(0)
    If VRMode Then
      VR_GO.visible = 0
      VR_Tilt.visible = 0
      VR_100k.visible = 0
      VR_200k.visible = 0
      VR_300k.visible = 0
      VR_400k.visible = 0
      VR_500k.visible = 0
      VR_600k.visible = 0
      VR_700k.visible = 0
      VR_800k.visible = 0
      VR_900k.visible = 0
      VR_1M.visible = 0
    End If
    Player=1
    playsound "StartUpSequence"
    PlaySound "crunchy_bass", 1, .01, AudioPan(Drain), .05,0, 0, 1, AudioFade(Drain)

    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
    rst=0
    BallInPlay=BallsPerGame
    InProgress=true
    resettimer.enabled=true
    BonusMultiplier=1
    ResetDrops
    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      if Credits=0 then
        Controller.B2SSetData 29,10
      else
        Controller.B2SSetData 29,Credits
      end if
      Controller.B2SSetScorePlayer 2,HSScore(1)
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      'Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0
    End If
    HSReel.SetValue(HSScore(1))
    For each obj in PlayerScores
'     obj.ResetToZero
    next
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    For each obj in PlayerHUDScores
      obj.state=0
    next
    If Table1.ShowDT = True then
      For each obj in PlayerScores
'       obj.ResetToZero
        obj.Visible=true
      next
      For each obj in PlayerScoresOn
'       obj.ResetToZero
        obj.Visible=false
      next

      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      PlayerHuds(Player-1).SetValue(1)
      PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).Visible=0
      PlayerScoresOn(Player-1).Visible=1
    end If

  end if

  If keycode = LeftFlipperKey Then
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X + 10
  End If

  If keycode = RightFlipperKey Then
    VR_CabFlipperRight.X = VR_CabFlipperRight.X - 10
  End If

End Sub

Sub Table1_KeyUp(ByVal keycode)

  ' GNMOD
  if EnteringInitials then
    exit sub
  end if

  If keycode = PlungerKey Then
    if PlungerPulled = 0 then
      exit sub
    end if
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    VR_Primary_plunger.Y = 2061
    SoundPlungerReleaseBall
    Plunger.Fire
  End If

  If keycode = LeftFlipperKey Then
    If InProgress=true and TableTilted=false Then
      LeftFlipper.RotateToStart
      If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
        RandomSoundFlipperDownLeft LeftFlipper
      End If
      FlipperLeftHitParm = FlipperUpSoundLevel
      DOF 101, 0
    End If
    StopSound "buzzL"
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X - 10
  End If

  If keycode = RightFlipperKey Then
    If InProgress=true and TableTilted=false Then
      RightFlipper.RotateToStart
      If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
        RandomSoundFlipperDownRight RightFlipper
      End If
      FlipperRightHitParm = FlipperUpSoundLevel
      DOF 102, 0
    End If
    StopSound "buzz"
    VR_CabFlipperRight.X = VR_CabFlipperRight.X + 10
  End If

  if keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y + 4
  End If

End Sub

Sub TimerPlunger_Timer
  If VR_Primary_plunger.Y < 2182 then
         VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  VR_Primary_plunger.Y = 2061 + (5* Plunger.Position) -20
End Sub

Dim max_angle_RF, min_angle_RF, mid_angle_RF            ' min and max angles from the right flipper
max_angle_RF = RightFlipper.StartAngle              ' right flipper down angle
min_angle_RF = RightFlipper.EndAngle              ' right flipper up angle
mid_angle_RF = (max_angle_RF-min_angle_RF)/2 + min_angle_RF   ' right flipper bake map switch point angle

Dim max_angle_LF, min_angle_LF, mid_angle_LF            ' min and max angles from the left flipper
max_angle_LF = LeftFlipper.StartAngle             ' left flipper down angle
min_angle_LF = LeftFlipper.EndAngle               ' left flipper up angle
mid_angle_LF = (max_angle_LF-min_angle_LF)/2 + min_angle_LF   ' left flipper bake map switch point angle


Sub LeftFlipper_Animate
  Dim BP
  Dim a: a = LeftFlipper.CurrentAngle       ' store flipper angle in a
  FlipperLSh.RotZ = a
  For Each BP in BP_LF
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a > mid_angle_LF
  Next
  For Each BP in BP_LFU
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle_LF
  Next
End Sub

Sub RightFlipper_Animate
  Dim BP
  Dim a: a = RightFlipper.CurrentAngle        ' store flipper angle in a
  FlipperRSh.RotZ = a
  For Each BP in BP_RF
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle_RF
  Next
  For Each BP in BP_RFU
    BP.visible = a > mid_angle_RF
    BP.RotZ = a                 ' rotate the maps
  Next
End Sub


'Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub

Sub Drain_Hit()
    DOF 121, 2
    RandomSoundDrain Drain
    Pause4Bonustimer.enabled=1
End Sub

Sub Trigger1_Unhit()
  DOF 124, 2
End Sub

Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  AddBonus
End Sub

Sub Gate_Animate()
  Dim a : a = Gate.currentangle
  Dim BP
  For Each BP in BP_Pgate : BP.rotx = -a : Next
End Sub


'***********************************

Dim AStep, BStep, CStep, DStep

Const UpperRubberHitThreshold = 4 ' tweak this for how hard the A/B rubber needs to be hit
Const LowerRubberHitThreshold = 2 ' tweak this for how hard the C/D rubber needs to be hit
Const BaseTimerInterval = 20

Sub DingWallA_Hit()
  ' rubber positions: 2 - 3 - (0) - 4 - 5

  Dim hitstrength, finalspeed, timerinterval

  'determine hit speed -- if too low, just vibrate the rubber but skip everything else
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  debug.print "A final speed: " & finalspeed

  if finalspeed < UpperRubberHitThreshold then
    'soft hit, not enough to trigger switch
    debug.print "A soft hit: " & finalspeed
    AStep = 8

    For Each BP in BP_SlingA_000 : BP.visible = 0 : Next
    For Each BP in BP_SlingA_003 : BP.visible = 1 : Next

    SlingATimer.UserValue = 2 ' used to tweak effect
    SlingATimer.Enabled = 1

    exit sub
  end if

  If TableTilted=false then
    AddScore(10)
  End if

  AStep = 0

  'hitstrength determines how many times to vibrate the rubber.
  '1 or 2 seems about right but tweak as needed
  hitstrength = 1
  if finalspeed > 10 then
    hitstrength = 2
  end if

  Dim BP
  'push the rubber in
  For Each BP in BP_SlingA_000 : BP.visible = 0 : Next
  For Each BP in BP_SlingA_002 : BP.visible = 1 : Next
  'bend the leaf
  For Each BP in BP_LeafA : BP.ObjRoty = -2: Next

  SlingATimer.UserValue = hitstrength ' used to tweak effect
  SlingATimer.Enabled = 1
End Sub

Sub SlingATimer_Timer
  ' sling starts in the "inner" position (002)
  ' rubber positions: 2 - 3 - (0) - 4 - 5
  Dim BP

  For Each BP in BP_SlingA_002 : BP.visible = 0 : Next
  For Each BP in BP_SlingA_003 : BP.visible = 0 : Next
  For Each BP in BP_SlingA_000 : BP.visible = 0 : Next
  For Each BP in BP_SlingA_004 : BP.visible = 0 : Next
  For Each BP in BP_SlingA_005 : BP.visible = 0 : Next

    Select Case AStep
        Case 0: For Each BP in BP_SlingA_003 : BP.visible = 1 : Next
        For Each BP in BP_LeafA : BP.ObjRoty = 0: Next
        Case 1: For Each BP in BP_SlingA_000 : BP.visible = 1 : Next
    Case 2: For Each BP in BP_SlingA_004 : BP.visible = 1 : Next
    Case 3: For Each BP in BP_SlingA_005 : BP.visible = 1 : Next
    Case 4: For Each BP in BP_SlingA_004 : BP.visible = 1 : Next
    Case 5: For Each BP in BP_SlingA_000 : BP.visible = 1 : Next
    Case 6: For Each BP in BP_SlingA_003 : BP.visible = 1 : Next
    Case 7: For Each BP in BP_SlingA_000 : BP.visible = 1 : Next
    End Select

  if AStep > 7 then
    'vibrate/settle the rubber
    Select Case AStep mod 4
      Case 0: For Each BP in BP_SlingA_004 : BP.visible = 1 : Next
      Case 1: For Each BP in BP_SlingA_000 : BP.visible = 1 : Next
      Case 2: For Each BP in BP_SlingA_003 : BP.visible = 1 : Next
      Case 3: For Each BP in BP_SlingA_000 : BP.visible = 1 : Next
    End Select
  end if

    AStep = AStep + 1
  SlingATimer.Interval = SlingATimer.Interval + 1 ' arbitrary slowing effect

  'vibrate rubber based on how hard the rubber was hit: 1,2,3...
  if AStep > (8 * SlingATimer.UserValue) then
    debug.print "Ending A anim with hitstrength of " & SlingATimer.UserValue & " at step " & AStep
    SlingATimer.Enabled = 0
    SlingATimer.Interval = BaseTimerInterval
    AStep = 0
    For Each BP in BP_SlingA_002 : BP.visible = 0 : Next
    For Each BP in BP_SlingA_003 : BP.visible = 0 : Next
    For Each BP in BP_SlingA_000 : BP.visible = 1 : Next
    For Each BP in BP_SlingA_004 : BP.visible = 0 : Next
    For Each BP in BP_SlingA_005 : BP.visible = 0 : Next
    For Each BP in BP_LeafA : BP.ObjRoty = 0: Next
  end if
End Sub

Sub DingWallB_Hit()
  ' rubber positions: 2 - 3 - (0) - 4 - 5

  Dim hitstrength, finalspeed, timerinterval

  'determine hit speed -- if too low, just vibrate the rubber but skip everything else
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  debug.print "B final speed: " & finalspeed

  if finalspeed < UpperRubberHitThreshold then
    'soft hit, not enough to trigger switch
    debug.print "B soft hit: " & finalspeed
    BStep = 8

    For Each BP in BP_SlingB_000 : BP.visible = 0 : Next
    For Each BP in BP_SlingB_003 : BP.visible = 1 : Next

    SlingBTimer.UserValue = 2 ' used to tweak effect
    SlingBTimer.Enabled = 1

    exit sub
  end if

  If TableTilted=false then
    AddScore(10)
  End if

  BStep = 0

  'hitstrength determines how many times to vibrate the rubber.
  '1 or 2 seems about right but tweak as needed
  hitstrength = 1
  if finalspeed > 10 then
    hitstrength = 2
  end if

  Dim BP
  'push the rubber in
  For Each BP in BP_SlingB_000 : BP.visible = 0 : Next
  For Each BP in BP_SlingB_002 : BP.visible = 1 : Next
  'bend the leaf
  For Each BP in BP_LeafB : BP.ObjRoty = 2 : Next

  SlingBTimer.UserValue = hitstrength ' used to tweak effect
  SlingBTimer.Enabled = 1
End Sub

Sub SlingBTimer_Timer
  ' sling starts in the "inner" position (002)
  ' rubber positions: 2 - 3 - (0) - 4 - 5
  Dim BP

  For Each BP in BP_SlingB_002 : BP.visible = 0 : Next
  For Each BP in BP_SlingB_003 : BP.visible = 0 : Next
  For Each BP in BP_SlingB_000 : BP.visible = 0 : Next
  For Each BP in BP_SlingB_004 : BP.visible = 0 : Next
  For Each BP in BP_SlingB_005 : BP.visible = 0 : Next

    Select Case BStep
        Case 0: For Each BP in BP_SlingB_003 : BP.visible = 1 : Next
        For Each BP in BP_LeafB : BP.ObjRoty = 0: Next
        Case 1: For Each BP in BP_SlingB_000 : BP.visible = 1 : Next
    Case 2: For Each BP in BP_SlingB_004 : BP.visible = 1 : Next
    Case 3: For Each BP in BP_SlingB_005 : BP.visible = 1 : Next
    Case 4: For Each BP in BP_SlingB_004 : BP.visible = 1 : Next
    Case 5: For Each BP in BP_SlingB_000 : BP.visible = 1 : Next
    Case 6: For Each BP in BP_SlingB_003 : BP.visible = 1 : Next
    Case 7: For Each BP in BP_SlingB_000 : BP.visible = 1 : Next
    End Select

  if BStep > 7 then
    Select Case BStep mod 4
      Case 0: For Each BP in BP_SlingB_004 : BP.visible = 1 : Next
      Case 1: For Each BP in BP_SlingB_000 : BP.visible = 1 : Next
      Case 2: For Each BP in BP_SlingB_003 : BP.visible = 1 : Next
      Case 3: For Each BP in BP_SlingB_000 : BP.visible = 1 : Next
    End Select
  end if

    BStep = BStep + 1
  SlingBTimer.Interval = SlingBTimer.Interval + 1 ' arbitrary slowing effect

  'vibrate rubber based on how hard the rubber was hit: 1,2,3...
  if BStep > (8 * SlingBTimer.UserValue) then
    debug.print "Ending B anim with hitstrength of " & SlingBTimer.UserValue & " at step " & BStep
    SlingBTimer.Enabled = 0
    SlingBTimer.Interval = BaseTimerInterval
    BStep = 0
    For Each BP in BP_SlingB_002 : BP.visible = 0 : Next
    For Each BP in BP_SlingB_003 : BP.visible = 0 : Next
    For Each BP in BP_SlingB_000 : BP.visible = 1 : Next
    For Each BP in BP_SlingB_004 : BP.visible = 0 : Next
    For Each BP in BP_SlingB_005 : BP.visible = 0 : Next
    For Each BP in BP_LeafB : BP.ObjRoty = 0: Next
  end if
End Sub

Sub DingWallC_Hit()
  ' rubber positions: 3 - 4 - (0) - 5 - 2

  Dim hitstrength, finalspeed, timerinterval

  'determine hit speed -- if too low, just vibrate the rubber but skip everything else
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  debug.print "C final speed: " & finalspeed

  if finalspeed < LowerRubberHitThreshold then
    'soft hit, not enough to trigger switch
    debug.print "C soft hit: " & finalspeed
    CStep = 8

    For Each BP in BP_SlingC_000 : BP.visible = 0 : Next
    For Each BP in BP_SlingC_004 : BP.visible = 1 : Next

    SlingCTimer.UserValue = 2 ' used to tweak effect
    SlingCTimer.Enabled = 1

    exit sub
  end if

  If TableTilted=false then
    AddScore(10)
  End if

  CStep = 0

  'hitstrength determines how many times to vibrate the rubber.
  '1 or 2 seems about right but tweak as needed
  hitstrength = 1
  if finalspeed > 10 then
    hitstrength = 2
  end if

  Dim BP
  'push the rubber in
  For Each BP in BP_SlingC_000 : BP.visible = 0 : Next
  For Each BP in BP_SlingC_003 : BP.visible = 1 : Next
  'bend the leaf
  For Each BP in BP_LeafC : BP.ObjRoty = -1: Next

  SlingCTimer.UserValue = hitstrength ' used to tweak effect
  SlingCTimer.Enabled = 1
End Sub

Sub SlingCTimer_Timer
  ' sling starts in the "inner" most position
  ' rubber positions to array index, 0 is resting: 3 - 4 - (0) - 5 - 2
  Dim BP

  For Each BP in BP_SlingC_000 : BP.visible = 0 : Next
  For Each BP in BP_SlingC_002 : BP.visible = 0 : Next
  For Each BP in BP_SlingC_003 : BP.visible = 0 : Next
  For Each BP in BP_SlingC_004 : BP.visible = 0 : Next
  For Each BP in BP_SlingC_005 : BP.visible = 0 : Next

    Select Case CStep
        Case 0: For Each BP in BP_SlingC_004 : BP.visible = 1 : Next
        For Each BP in BP_LeafC : BP.ObjRoty = 0: Next
        Case 1: For Each BP in BP_SlingC_000 : BP.visible = 1 : Next
    Case 2: For Each BP in BP_SlingC_005 : BP.visible = 1 : Next
    Case 3: For Each BP in BP_SlingC_002 : BP.visible = 1 : Next
    Case 4: For Each BP in BP_SlingC_005 : BP.visible = 1 : Next
    Case 5: For Each BP in BP_SlingC_000 : BP.visible = 1 : Next
    Case 6: For Each BP in BP_SlingC_004 : BP.visible = 1 : Next
    Case 7: For Each BP in BP_SlingC_000 : BP.visible = 1 : Next
    End Select

  if CStep > 7 then
    Select Case CStep mod 4
      Case 0: For Each BP in BP_SlingC_005 : BP.visible = 1 : Next
      Case 1: For Each BP in BP_SlingC_000 : BP.visible = 1 : Next
      Case 2: For Each BP in BP_SlingC_004 : BP.visible = 1 : Next
      Case 3: For Each BP in BP_SlingC_000 : BP.visible = 1 : Next
    End Select
  end if

    CStep = CStep + 1
  SlingCTimer.Interval = SlingCTimer.Interval + 1 ' arbitrary slowing effect

  'vibrate rubber based on how hard the rubber was hit: 1,2,3...
  if CStep > (8 * SlingCTimer.UserValue) then
    debug.print "Ending C anim with hitstrength of " & SlingCTimer.UserValue & " at step " & CStep
    SlingCTimer.Enabled = 0
    SlingCTimer.Interval = BaseTimerInterval
    CStep = 0
    For Each BP in BP_SlingC_000 : BP.visible = 1 : Next
    For Each BP in BP_SlingC_002 : BP.visible = 0 : Next
    For Each BP in BP_SlingC_003 : BP.visible = 0 : Next
    For Each BP in BP_SlingC_004 : BP.visible = 0 : Next
    For Each BP in BP_SlingC_005 : BP.visible = 0 : Next
    For Each BP in BP_LeafC : BP.ObjRoty = 0: Next
  end if
End Sub

Sub DingWallD_Hit()
  ' rubber positions: 3 - 4 - (0) - 5 - 2

  Dim hitstrength, finalspeed, timerinterval

  'determine hit speed -- if too low, just vibrate the rubber but skip everything else
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  debug.print "D final speed: " & finalspeed

  if finalspeed < LowerRubberHitThreshold then
    'soft hit, not enough to trigger switch
    debug.print "D soft hit: " & finalspeed
    DStep = 8

    For Each BP in BP_SlingD_000 : BP.visible = 0 : Next
    For Each BP in BP_SlingD_004 : BP.visible = 1 : Next

    SlingDTimer.UserValue = 2 ' used to tweak effect
    SlingDTimer.Enabled = 1

    exit sub
  end if

  If TableTilted=false then
    AddScore(10)
  End if

  DStep = 0

  'hitstrength determines how many times to vibrate the rubber.
  '1 or 2 seems about right but tweak as needed
  hitstrength = 1
  if finalspeed > 10 then
    hitstrength = 2
  end if

  Dim BP
  'push the rubber in
  For Each BP in BP_SlingD_000 : BP.visible = 0 : Next
  For Each BP in BP_SlingD_003 : BP.visible = 1 : Next
  'bend the leaf
  For Each BP in BP_LeafD : BP.ObjRoty = -1: Next

  SlingDTimer.UserValue = hitstrength ' used to tweak effect
  SlingDTimer.Enabled = 1
End Sub

Sub SlingDTimer_Timer
  ' sling starts in the "inner" most position
  ' rubber positions to array index, 0 is resting: 3 - 4 - (0) - 5 - 2
  Dim BP

  For Each BP in BP_SlingD_000 : BP.visible = 0 : Next
  For Each BP in BP_SlingD_002 : BP.visible = 0 : Next
  For Each BP in BP_SlingD_003 : BP.visible = 0 : Next
  For Each BP in BP_SlingD_004 : BP.visible = 0 : Next
  For Each BP in BP_SlingD_005 : BP.visible = 0 : Next

    Select Case DStep
        Case 0: For Each BP in BP_SlingD_004 : BP.visible = 1 : Next
        For Each BP in BP_LeafD : BP.ObjRotx = 0: Next
        Case 1: For Each BP in BP_SlingD_000 : BP.visible = 1 : Next
    Case 2: For Each BP in BP_SlingD_005 : BP.visible = 1 : Next
    Case 3: For Each BP in BP_SlingD_002 : BP.visible = 1 : Next
    Case 4: For Each BP in BP_SlingD_005 : BP.visible = 1 : Next
    Case 5: For Each BP in BP_SlingD_000 : BP.visible = 1 : Next
    Case 6: For Each BP in BP_SlingD_004 : BP.visible = 1 : Next
    Case 7: For Each BP in BP_SlingD_000 : BP.visible = 1 : Next
    End Select

  if DStep > 7 then
    Select Case DStep mod 4
      Case 0: For Each BP in BP_SlingD_005 : BP.visible = 1 : Next
      Case 1: For Each BP in BP_SlingD_000 : BP.visible = 1 : Next
      Case 2: For Each BP in BP_SlingD_004 : BP.visible = 1 : Next
      Case 3: For Each BP in BP_SlingD_000 : BP.visible = 1 : Next
    End Select
  end if

    DStep = DStep + 1
  SlingDTimer.Interval = SlingDTimer.Interval + 1 ' arbitrary slowing effect

  'vibrate rubber based on how hard the rubber was hit: 1,2,3...
  if DStep > (8 * SlingDTimer.UserValue) then
    debug.print "Ending D anim with hitstrength of " & SlingDTimer.UserValue & " at step " & DStep
    SlingDTimer.Enabled = 0
    SlingDTimer.Interval = BaseTimerInterval
    DStep = 0
    For Each BP in BP_SlingD_000 : BP.visible = 1 : Next
    For Each BP in BP_SlingD_002 : BP.visible = 0 : Next
    For Each BP in BP_SlingD_003 : BP.visible = 0 : Next
    For Each BP in BP_SlingD_004 : BP.visible = 0 : Next
    For Each BP in BP_SlingD_005 : BP.visible = 0 : Next
    For Each BP in BP_LeafD : BP.ObjRotx = 0: Next
  end if
End Sub


'***********************
' slingshots

InitSlings

Sub InitSlings()
  Dim BP
  For Each BP in BP_rsling_002 : BP.visible = 0 : Next
  For Each BP in BP_rsling_001 : BP.visible = 0 : Next
  For Each BP in BP_lsling_002 : BP.visible = 0 : Next
  For Each BP in BP_lsling_001 : BP.visible = 0 : Next

  For Each BP in BP_SlingA_002 : BP.visible = 0 : Next
  For Each BP in BP_SlingA_003 : BP.visible = 0 : Next
  For Each BP in BP_SlingB_002 : BP.visible = 0 : Next
  For Each BP in BP_SlingB_003 : BP.visible = 0 : Next
End Sub
'

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotRight zCol_Rubber_Post022
  DOF 154, DOFPulse

  Dim BP
  For Each BP in BP_rsling_000 : BP.visible = 0 : Next
  For Each BP in BP_rsling_002 : BP.visible = 1 : Next
  For Each BP in BP_slingR : BP.transy = -15 : Next
    RStep = 0
    RightSlingShot.TimerEnabled = 1

  If TableTilted = False Then
    AddScore 10
  End If
End Sub

Sub RightSlingShot_Timer
  Dim BP
    Select Case RStep
        Case 3:
      For Each BP in BP_rsling_002 : BP.visible = 0 : Next
      For Each BP in BP_rsling_001 : BP.visible = 1 : Next
      For Each BP in BP_slingR : BP.transy = -10 : Next
        Case 4:
      For Each BP in BP_rsling_001 : BP.visible = 0 : Next
      For Each BP in BP_rsling_000 : BP.visible = 1 : Next
      For Each BP in BP_slingR : BP.transy = 0 : Next
      RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft zCol_Rubber_Post019
  DOF 153, DOFPulse

  Dim BP
  For Each BP in BP_lsling_000 : BP.visible = 0 : Next
  For Each BP in BP_lsling_002 : BP.visible = 1 : Next
  For Each BP in BP_slingL : BP.transy = -15 : Next
    LStep = 0
  LeftSlingShot.TimerEnabled = 1

  If TableTilted = False Then
    AddScore 10
  End If
End Sub

Sub LeftSlingShot_Timer
  Dim BP
    Select Case LStep
        Case 3:
      For Each BP in BP_lsling_002 : BP.visible = 0 : Next
      For Each BP in BP_lsling_001 : BP.visible = 1 : Next
      For Each BP in BP_slingL : BP.transy = -10 : Next
        Case 4:
      For Each BP in BP_lsling_001 : BP.visible = 0 : Next
      For Each BP in BP_lsling_000 : BP.visible = 1 : Next
      For Each BP in BP_slingL : BP.transy = 0 : Next
      LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'***********************************

Sub Bumper1_Hit
  If TableTilted=false then
    RandomSoundBumperBottom Bumper1
    DOF 155, DOFPulse
    bump1 = 1
    If NoveltyMode=0 then
      AddScore(100)
    Else
      AddScore(1000)
    End If
    ToggleAlternatingRelay
    end if

  Dim BP
  For Each BP in BP_Bumper1_Skirt_001
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper1_timer
  Dim BP
  For Each BP in BP_Bumper1_Skirt_001
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper1_Animate
  Dim z: z = Bumper1.CurrentRingOffset

  Dim BP
  For Each BP in BP_Bumper1_Ring_001 : BP.transz = z : Next
End Sub

Sub Bumper2_Hit
  If TableTilted=false then
    RandomSoundBumperBottom Bumper2
    DOF 156, DOFPulse
    bump3 = 1
    If NoveltyMode=0 then
      AddScore(100)
    Else
      AddScore(1000)
    End If
    ToggleAlternatingRelay
    end if

  Dim BP
  For Each BP in BP_Bumper2_Skirt_001
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper2_timer
  Dim BP
  For Each BP in BP_Bumper2_Skirt_001
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper2_Animate
  Dim z: z = Bumper2.CurrentRingOffset

  Dim BP
  For Each BP in BP_Bumper2_Ring_001 : BP.transz = z : Next
End Sub



'******************************************************
'     SKIRT ANIMATION FUNCTIONS
'******************************************************
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animaation, adjust to your liking

'Const PI = 3.1415926
Const SkirtTilt=5   'angle of skirt tilting in degrees

Function SkirtAX(bumper, bumperball)
  skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)    'x component of angle
  if (bumper.y<bumperball.y) then skirtAX=skirtAX*-1    'adjust for ball hit bottom half
End Function

Function SkirtAY(bumper, bumperball)
  skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)    'y component of angle
  if (bumper.x>bumperball.x) then skirtAY=skirtAY*-1    'adjust for ball hit left half
End Function

Function SkirtA(bumper, bumperball)
  dim hitx, hity, dx, dy
  hitx=bumperball.x
  hity=bumperball.y

  dy=Abs(hity-bumper.y)         'y offset ball at hit to center of bumper
  if dy=0 then dy=0.0000001
  dx=Abs(hitx-bumper.x)         'x offset ball at hit to center of bumper
  skirtA=(atn(dx/dy)) '/(PI/180)      'angle in radians to ball from center of Bumper1
End Function


'***********************************

Sub sw11_Hit()
    STHit 11
  DOF 106, 2
end Sub


'***********************************

Sub TriggerTop1_Hit()
  If TableTilted=false then
    DOF 108, 2
    SetMotor(500)
    UpperLight1.state=0
    CheckABC
  end if
  AnimateWire BP_TriggerTop1, 1
End Sub

Sub TriggerTop1_UnHit()
  AnimateWire BP_TriggerTop1, 0
End Sub

Sub TriggerTop2_Hit()
  If TableTilted=false then
    DOF 109, 2
    SetMotor(500)
    UpperLight2.state=0
    If Lane2Or3SpotsBoth = 1 Then
      UpperLight3.state=0
    End If
    CheckABC
  end if
  AnimateWire BP_TriggerTop2, 1
End Sub

Sub TriggerTop2_UnHit()
  AnimateWire BP_TriggerTop2, 0
End Sub

Sub TriggerTop3_Hit()
  If TableTilted=false then
    DOF 110, 2
    SetMotor(500)
    UpperLight3.state=0
    If Lane2Or3SpotsBoth = 1 Then
      UpperLight2.state=0
    End If
    CheckABC
  end if
  AnimateWire BP_TriggerTop3, 1
End Sub

Sub TriggerTop3_UnHit()
  AnimateWire BP_TriggerTop3, 0
End Sub

Sub TriggerTop4_Hit()
  If TableTilted=false then
    DOF 150, 2
    SetMotor(500)
    UpperLight4.state=0
    CheckABC
  end if
  AnimateWire BP_TriggerTop4, 1
End Sub

Sub TriggerTop4_UnHit()
  AnimateWire BP_TriggerTop4, 0
End Sub

Sub TriggerLeftInlane_Hit()
  If TableTilted=false then
    DOF 112, 2
    If LeftInlaneLight5000.state=1 then
      SetMotor(5000)
      ResetDropsTimer.enabled=true
    else
      SetMotor(500)
    end if
  End If
  AnimateWire BP_TriggerLeftInlane, 1
End Sub

Sub TriggerLeftInlane_UnHit()
  AnimateWire BP_TriggerLeftInlane, 0
End Sub

Sub TriggerLeftInlane2_Hit()
  If TableTilted=false then
    DOF 113, 2
    SetMotor(500)
    LeftInlaneLight5.state=0
    LFive.state = 0
    CheckABC
  End If
  AnimateWire BP_TriggerLeftInlane2, 1
End Sub

Sub TriggerLeftInlane2_UnHit()
  AnimateWire BP_TriggerLeftInlane2, 0
End Sub

Sub TriggerRightInlane_Hit()
  If TableTilted=false then
    DOF 115, 2
    If RightInlaneLight5000.state=1 then
      SetMotor(5000)
      ResetDropsTimer.enabled=true
    else
      SetMotor(500)
    end if
  End If
  AnimateWire BP_TriggerRightInlane, 1
End Sub

Sub TriggerRightInlane_UnHit()
  AnimateWire BP_TriggerRightInlane, 0
End Sub

Sub TriggerRightInlane2_Hit()
  If TableTilted=false then
    DOF 116, 2
    SetMotor(500)
    RightInlaneLight6.state=0
    LSix.state = 0
    CheckABC
  End If
  AnimateWire BP_TriggerRightInlane2, 1
End Sub

Sub TriggerRightInlane2_UnHit()
  AnimateWire BP_TriggerRightInlane2, 0
End Sub

Sub TriggerLeftOutlane_Hit()
  If TableTilted=false then
    DOF 111, 2
    SetMotor(5000)
  End If
  AnimateWire BP_TriggerLeftOutlane, 1
End Sub

Sub TriggerLeftOutlane_UnHit()
  AnimateWire BP_TriggerLeftOutlane, 0
End Sub

Sub TriggerRightOutlane_Hit()
  If TableTilted=false then
    DOF 114, 2
    SetMotor(5000)
  End If
  AnimateWire BP_TriggerRightOutlane, 1
End Sub

Sub TriggerRightOutlane_UnHit()
  AnimateWire BP_TriggerRightOutlane, 0
End Sub

Sub AnimateWire(group, action) ' Action = 1 - to drop, 0 to raise)
  Dim BP
  If action = 1 Then
    For Each BP in group : BP.transz = -13 : Next
  Else
    For Each BP in group : BP.transz = 0 : Next
  End If
End Sub


Sub CheckABC
  if UpperLight1.state=0 AND UpperLight2.state=0 AND UpperLight3.state=0 AND UpperLight4.state=0 AND LeftInlaneLight5.state=0 AND RightInlaneLight6.state=0 then
    If ABCCompleted = False Then
      ' Award points for completing 1-6 sequence (first time only)
      If NoveltyMode = 1 Then
        SetMotor(50000)
      Else
        SetMotor(5000)
      End If
    End If
    ABCCompleted=true
    LeftSpecialLight.state=1
    RightSpecialLight.state=1
  else
    ABCCompleted=false
    LeftSpecialLight.state=0
    RightSpecialLight.state=0
  end if
end sub

'***********************************

Sub sw1_Hit()
  if TableTilted=false then
        DTHit 1
  end if
end sub

Sub sw2_Hit()
  if TableTilted=false then
        DTHit 2
  end if
end sub

Sub sw3_Hit()
  if TableTilted=false then
        DTHit 3
  end if
end sub

Sub sw4_Hit()
  if TableTilted=false then
        DTHit 4
  end if
end sub

Sub sw5_Hit()
  if TableTilted=false then
        DTHit 5
  end if
end sub

Sub sw12_hit()
  if TableTilted=false then
    If LeftTargetLight5.state=1 then
      If LeftSpecialLight.state=1 then
        AddSpecial
      else
        SetMotor(5000)
      end if
    else
      SetMotor(500)
    end if
    LeftInlaneLight5.state=0
    LFive.state = 0
    CheckABC
  end if

end sub

Sub sw6_Hit()
  if TableTilted=false then
        DTHit 6
  end if
end sub

Sub sw7_Hit()
  if TableTilted=false then
        DTHit 7
  end if
end sub

Sub sw8_Hit()
  if TableTilted=false then
        DTHit 8
  end if
end sub

Sub sw9_Hit()
  if TableTilted=false then
        DTHit 9
  end if
end sub

Sub sw10_Hit()
  if TableTilted=false then
        DTHit 10
  end if
end sub

Sub sw13_hit()
  if TableTilted=false then
    If RightTargetLight1.state=1 then
      If RightSpecialLight.state=1 then
        AddSpecial
      else
        SetMotor(5000)
      end if
    else
      SetMotor(500)
    end if
    RightInlaneLight6.state=0
    LSix.state = 0
    CheckABC
  end if
end sub

Sub ScoreDrop(sw)
  Select Case sw
    Case 1:
      if LeftTargetLight1.state=1 then
        if LeftSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 119, 2
    Case 2:
      if LeftTargetLight2.state=1 then
        if LeftSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 119, 2
    Case 3:
      if LeftTargetLight3.state=1 then
        if LeftSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 119, 2
    Case 4:
      if LeftTargetLight4.state=1 then
        if LeftSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 119, 2
    Case 5:
      if LeftTargetLight5.state=1 then
        if LeftSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 119, 2
    Case 6:
      if RightTargetLight1.state=1 then
        if RightSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 120, 2
    Case 7:
          if RightTargetLight2.state=1 then
        if RightSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 120, 2
    Case 8:
      if RightTargetLight3.state=1 then
        if RightSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 120, 2
    Case 9:
      if RightTargetLight4.state=1 then
        if RightSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 120, 2
    Case 10:
      if RightTargetLight5.state=1 then
        if RightSpecialLight.state=1 then
          AddSpecial
        else
          SetMotor(5000)
        end if
      else
        SetMotor(500)
      end if
      DOF 120, 2
  End Select
End Sub


Sub ResetCenterDrops
    RandomSoundDropTargetReset sw11
  LeftTargetCounter=4
  RightTargetCounter=4
    DTRaise 3
    DTRaise 8
end sub

Sub ResetLeftDrops
  RandomSoundDropTargetReset BM_DT5
  LeftTargetCounter=0
  LeftSpecialLight.state=0
    DTRaise 1
    DTRaise 2
    DTRaise 3
    DTRaise 4
    DTRaise 5

  DOF 151, DOFPulse
end Sub

Sub ResetRightDrops
    RandomSoundDropTargetReset BM_DT6
  RightTargetCounter=0
  RightSpecialLight.state=0
    DTRaise 6
    DTRaise 7
    DTRaise 8
    DTRaise 9
    DTRaise 10

  DOF 152, DOFPulse
end Sub


Sub ResetDrops
  '* prev implementation
  'RandomSoundDropTargetReset BM_DT5
  'RandomSoundDropTargetReset BM_DT6
    'DTRaise 1
    'DTRaise 2
    'DTRaise 3
    'DTRaise 4
    'DTRaise 5
    'DTRaise 6
    'DTRaise 7
    'DTRaise 8
    'DTRaise 9
    'DTRaise 10

  'CenterSpecialLight.state=0
  'LeftInlaneLight5000.state=0
  'RightInlaneLight5000.state=0
  'LeftTargetCounter=0
  'RightTargetCounter=0
  '* end prev implementation

  ResetLeftDrops
  ResetRightDrops

  CenterSpecialLight.state=0
  LeftInlaneLight5000.state=0
  RightInlaneLight5000.state=0
end Sub

Sub CheckAllDrops
  If RightDropsComplete then
    RightInlaneLight5000.state=1
  else
    RightInlaneLight5000.state=0
  end if

  If LeftDropsComplete then
    LeftInlaneLight5000.state=1
  else
    LeftInlaneLight5000.state=0
  end if

  If RightDropsComplete AND LeftDropsComplete then
    CenterSpecialLight.state=1
  else
    CenterSpecialLight.state=0
  end if
end Sub

Function RightDropsComplete()
  Dim DropsComplete : DropsComplete = True
  Dim RightDropsArr : RightDropsArr = Array(DT6, DT7, DT8, DT9, DT10)
  Dim drop
  For Each drop in RightDropsArr
    If Not drop.isDropped Then
      DropsComplete = False
      Exit For
    End If
  Next
  RightDropsComplete = DropsComplete
End Function

Function LeftDropsComplete()
  Dim DropsComplete : DropsComplete = True
  Dim LeftDropsArr : LeftDropsArr = Array(DT1, DT2, DT3, DT4, DT5)
  Dim drop
  For Each drop in LeftDropsArr
    If Not drop.isDropped Then
      DropsComplete = False
      Exit For
    End If
  Next
  LeftDropsComplete = DropsComplete
End Function

sub ResetDropsTimer_timer
    ResetDropsTimer.enabled=0
    ResetDrops
end sub

sub ResetDropsTimer2_timer
    ResetDropsTimer2.enabled=0
    LeftTarget5000.state=1
    RightTarget5000.state=1
    LeftTarget500.state=0
    RightTarget500.state=0
    ResetCenterDrops
end sub

Sub AddSpecial()
  DOF 125, 2

  If InProgress=false then
    Exit Sub
  end if

  if NoveltyMode = 1 then
    SetMotor(50000)
    exit sub
  end if

  WOWCounter=WOWCounter+1
  If WOWCounter>5 then
    WOWCounter=5
  end if
  WOWReel.SetValue(WOWCounter)
' BallInPlayReel.SetValue(BallInPlay)
  if B2SOn then
'   Controller.B2SSetBallInPlay BallInPlay
    For i=51 to 50+WOWCounter
      Controller.B2SSetData i,1
    next
  end if
End Sub

Sub AddSpecial2()
  Credits=Credits+1
  DOF 126, 1
  if Credits>9 then Credits=9
  If B2SOn Then
    if Credits=0 then
      Controller.B2SSetData 29,10
    else
      Controller.B2SSetData 29,Credits
    end if
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddBonus()
  bonuscountdown=bonuscounter
  If BonusMultiplier=1 then
    ScoreBonus.interval=150
  Else
    ScoreBonus.interval=400
  end if
  ScoreBonus.enabled=true
End Sub

Sub ToggleAlternatingRelay
  AlternatingRelayCounter=AlternatingRelayCounter+1
  if AlternatingRelayCounter>4 then
    AlternatingRelayCounter=0
  end if
  For each obj in LeftTargetLights
    obj.state=0
  next
  For each obj in RightTargetLights
    obj.state=0
  next
  RightTargetLights(AlternatingRelayCounter).state=1
  LeftTargetLights(AlternatingRelayCounter).state=1

end sub

Sub ResetBallDrops()
  if ABCCompleted=true then
    For each obj in NumberLights
      obj.state=1
    next
    ABCCompleted=false
    LeftSpecialLight.state=0
    RightSpecialLight.state=0
  end if

  'ResetDrops

  'dof 128, 2 'removed for new dof
  BonusCounter=0
  HoleCounter=0
End Sub

Sub LightsOut
  for each obj in Bonus
    obj.state=0
  next
  for each obj in StarLights
    obj.state=0
  next
  BonusCounter=0
  HoleCounter=0
  Bumper1Light.state=0
  Bumper2Light.state=0

  if dooralreadyopen=1 then
    closeg.enabled=true
  end if
  if kgdooralreadyopen=1 then
    closekg.enabled=true
  end if
end sub

Sub ResetBalls()

  TempMultiCounter=BallsPerGame-BallInPlay

  ResetBallDrops
  BonusMultiplier=1

  TableTilted=false
  TiltReel.SetValue(0)
  If B2Son then
    Controller.B2SSetTilt 0
  end if

    Drain.Kick 57,50
  PlaySound "ballrelease", 1, .01, AudioPan(Drain), .05,0, 0, 1, AudioFade(Drain)
  dof 122, 2
  if dooralreadyopen=1 then
    closeg.enabled=true
  end if
  if kgdooralreadyopen=1 then
    closekg.enabled=true
  end if
  BallInPlayReel.SetValue(BallInPlay)
  If VRMode Then
    VR_Tilt.visible = 0
    VR_Balls1.visible = 0
    VR_Balls2.visible = 0
    VR_Balls3.visible = 0
    VR_Balls4.visible = 0
    VR_Balls5.visible = 0
    Dim vr_bip : Set vr_bip = eval("VR_Balls" & BallInPlay)
    vr_bip.visible = 1
  End If
End Sub

sub delaykgclose_timer
  delaykgclose.enabled=false
  closekg.enabled=true
end sub

sub openg_timer
    bottgate(bgpos).isdropped=true
    bgpos=bgpos-1
    bottgate(bgpos).isdropped=false

  if bgpos=0 then
    playsound "postup"
    GateOpenLight.state=1
    openg.enabled=false
    dooralreadyopen=1
  end if

end sub

sub closeg_timer
    closeg.interval=10
    bottgate(bgpos).isdropped=true
    bgpos=bgpos+1
    bottgate(bgpos).isdropped=false

  if bgpos=6 then
    GateOpenLight.state=0
    closeg.enabled=false
    dooralreadyopen=0
  end if
end sub

 sub openkg_timer
    kickgate(kgpos).isdropped=true
    kgpos=kgpos+1
    kickgate(kgpos).isdropped=false

     if kgpos=6 then
    playsound "postup"
    KickGateOpenLight.state=1
    openkg.enabled=false
    kgdooralreadyopen=1
  end if

 end sub

sub closekg_timer
    closekg.interval=10
    kickgate(kgpos).isdropped=true
    kgpos=kgpos-1
    kickgate(kgpos).isdropped=false

  if kgpos=0 then
    KickGateOpenLight.state=0
    closekg.enabled=false
    kgdooralreadyopen=0
  end if
end sub

sub resettimer_timer
    rst=rst+1

  if rst<=BallsPerGame then
    BallInPlayReel.SetValue(rst)
    If VRMode Then
      VR_Balls1.visible = 0
      VR_Balls2.visible = 0
      VR_Balls3.visible = 0
      VR_Balls4.visible = 0
      VR_Balls5.visible = 0
      Dim vr_bip : Set vr_bip = eval("VR_Balls" & rst)
      vr_bip.visible = 1
    End If

    If B2SOn Then
      Controller.B2SSetBallInPlay rst
    End If
  end if
  if rst>5 and rst<16 then
    ResetReelsToZero(1)
  end if
    if rst=18 then
    playsound "StartBall1"
    end if
    if rst=20 then
    newgame
    resettimer.enabled=false
    end if
end sub

Sub ResetReelsToZero(reelzeroflag)
  dim d1(5)
  dim d2(5)
  dim scorestring1, scorestring2

  If reelzeroflag=1 then
    scorestring1=CStr(Score(1))
    scorestring2=CStr(Score(2))
    scorestring1=right("00000" & scorestring1,5)
    scorestring2=right("00000" & scorestring2,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
      d2(i)=CInt(mid(scorestring2,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if
      if d2(i)>0 then
        d2(i)=d2(i)+1
        if d2(i)>9 then d2(i)=0
      end if

    next
    Score(1)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    Score(2)=(d2(0)*10000) + (d2(1)*1000) + (d2(2)*100) + (d2(3)*10) + d2(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 1, Score(1)
      'Controller.B2SSetScorePlayer 2, Score(2)
    End If
    PlayerScores(0).SetValue(Score(1))
    PlayerScoresOn(0).SetValue(Score(1))
    PlayerScores(1).SetValue(Score(2))
    PlayerScoresOn(1).SetValue(Score(2))

    EMMODE = 1
    UpdateReels 0,0 ,score(0), 0, -1,-1,-1,-1,-1
    UpdateReels 1,1 ,score(1), 0, -1,-1,-1,-1,-1
    UpdateReels 2,2 ,score(2), 0, -1,-1,-1,-1,-1
    UpdateReels 3,3 ,score(3), 0, -1,-1,-1,-1,-1
    UpdateReels 4,4 ,score(4), 0, -1,-1,-1,-1,-1
    EMMODE = 0

  end if
  If reelzeroflag=2 then
    scorestring1=CStr(Score(3))
    scorestring2=CStr(Score(4))
    scorestring1=right("00000" & scorestring1,5)
    scorestring2=right("00000" & scorestring2,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
      d2(i)=CInt(mid(scorestring2,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if
      if d2(i)>0 then
        d2(i)=d2(i)+1
        if d2(i)>9 then d2(i)=0
      end if

    next
    Score(3)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    Score(4)=(d2(0)*10000) + (d2(1)*1000) + (d2(2)*100) + (d2(3)*10) + d2(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 3, Score(3)
      Controller.B2SSetScorePlayer 4, Score(4)
    End If
    PlayerScores(2).SetValue(Score(3))
    PlayerScoresOn(2).SetValue(Score(3))
    PlayerScores(3).SetValue(Score(4))
    PlayerScoresOn(3).SetValue(Score(4))

  end if

end sub



sub ScoreBonus_timer

  if bonuscountdown<=0 then
    ScoreBonus.enabled=false
    ScoreBonus.interval=50
    NextBallDelay.enabled=true
    exit sub
  end if
  if BonusCountdown>10 then
    SetMotor2(1000*BonusMultiplier)
    Bonus(bonuscountdown-10).state=0
  else
    SetMotor2(1000*BonusMultiplier)
    Bonus(bonuscountdown).state=0
  end if
  If BonusMultiplier=2 then
    Bonus(0).state=1
  end if
' If BonusMultiplier=1 then
'   ScoreBonus.interval=50
' Else
'   ScoreBonus.interval=150
' end if
  bonuscountdown=bonuscountdown-1

  if bonuscountdown>10 then
    Bonus(bonuscountdown-10).state=1
    exit sub
  end if
  if bonuscountdown>0 then
    Bonus(bonuscountdown).state=1
  end if

end sub

sub NextBallDelay_timer()
  NextBallDelay.enabled=false
  nextball
end sub

sub newgame
  InProgress=true
  queuedscore=0
  for i = 1 to 4
    Score(i)=0
    Score100K(1)=0
    HighScorePaid(i)=false
    Replay1Paid(i)=false
    Replay2Paid(i)=false
    Replay3Paid(i)=false
  next
  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0
    Controller.B2SSetMatch 0
'   Controller.B2SSetScorePlayer1 0
'   Controller.B2SSetScorePlayer2 0
'   Controller.B2SSetScorePlayer3 0
'   Controller.B2SSetScorePlayer4 0
    Controller.B2SSetBallInPlay BallInPlay
  End if

  for each obj in LeftTargetLights
    obj.state=0
  next
  for each obj in RightTargetLights
    obj.state=0
  next
  WOWCounter=0
  If B2SOn then
    For i = 1 to 5
      Controller.B2SSetData 50+i,0
    next
  end if
  WOWReel.SetValue(0)

  UpdateReels  0 ,0 ,Score(0), 0, 0,0,0,0,0
  UpdateReels  1 ,1 ,Score(0), 0, 0,0,0,0,0
  UpdateReels  2 ,2 ,Score(0), 0, 0,0,0,0,0
  UpdateReels  3 ,3 ,Score(0), 0, 0,0,0,0,0
  UpdateReels  4 ,4 ,Score(0), 0, 0,0,0,0,0


  AlternatingRelayCounter=0
  For each obj in LeftTargetLights
    obj.state=0
  next
  For each obj in RightTargetLights
    obj.state=0
  next
  RightTargetLights(AlternatingRelayCounter).state=1
  LeftTargetLights(AlternatingRelayCounter).state=1
  ABCCompleted=true
  Bumper1Light.state=1
  Bumper2Light.state=1

  BumpersOn
  ResetBalls
End sub

sub nextball
  If B2SOn Then
    Controller.B2SSetTilt 0
  End If
  Player=Player+1
  If Player>Players Then
    If WOWCounter>0 then
      If B2SOn then
        Controller.B2SSetData 50+WOWCounter,0
      end if
      WOWCounter=WOWCounter-1
      WOWReel.SetValue(WOWCounter)
    else
      BallInPlay=BallInPlay-1
    end if
    If BallInPlay<1 then
      PlaySound("MotorLeer")
      InProgress=false

      If B2SOn Then
        Controller.B2SSetGameOver 1
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetCanPlay 0
      End If
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
      end If
      BallInPlayReel.SetValue(0)
      CanPlayReel.SetValue(0)
      GameOverReel.SetValue(1)
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      If VRMode Then
        VR_GO.visible = 1
        VR_Balls1.visible = 0
        VR_Balls2.visible = 0
        VR_Balls3.visible = 0
        VR_Balls4.visible = 0
        VR_Balls5.visible = 0
      End If
      LightsOut
      BumpersOff
      PlasticsOff
      CheckHighScore
      Players=0
      HighScoreTimer.interval=100
      HighScoreTimer.enabled=True
    Else
      Player=1
      If B2SOn Then
        Controller.B2SSetPlayerUp Player
        Controller.B2SSetBallInPlay BallInPlay
      End If
      'PlaySound("RotateThruPlayers")
      TempPlayerUp=Player
      'PlayerUpRotator.enabled=true
      PlayStartBall.enabled=true
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
        PlayerHuds(Player-1).SetValue(1)
        PlayerHUDScores(Player-1).state=1
        PlayerScores(Player-1).visible=0
        PlayerScoresOn(Player-1).visible=1
      end If

      ResetBalls

    End If
  Else
    If B2SOn Then
      Controller.B2SSetPlayerUp Player
      Controller.B2SSetBallInPlay BallInPlay
    End If
    'PlaySound("RotateThruPlayers")
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
'   PlayStartBall.enabled=true
    if B2SOn then
      Controller.B2SSetPlayerUp Player
      Controller.B2SSetData 81,0
      Controller.B2SSetData 82,0
      Controller.B2SSetData 83,0
      Controller.B2SSetData 84,0
      Controller.B2SSetData 80+Player,1
    end if
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    For each obj in PlayerHUDScores
        obj.state=0
      next
    If Table1.ShowDT = True then
      For each obj in PlayerScores
        obj.visible=1
      Next
      For each obj in PlayerScoresOn
        obj.visible=0
      Next
      PlayerHuds(Player-1).SetValue(1)
      PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).visible=0
      PlayerScoresOn(Player-1).visible=1
    end If
    ResetBalls

  End If

End sub


sub CheckHighScore
  Dim playertops
    dim si
  dim sj
  dim stemp
  dim stempplayers
  for i=1 to 4
    sortscores(i)=0
    sortplayers(i)=0
  next
  playertops=0
  for i = 1 to Players
    sortscores(i)=Score(i)
    sortplayers(i)=i
  next

  for si = 1 to Players
    for sj = 1 to Players-1
      if sortscores(sj)>sortscores(sj+1) then
        stemp=sortscores(sj+1)
        stempplayers=sortplayers(sj+1)
        sortscores(sj+1)=sortscores(sj)
        sortplayers(sj+1)=sortplayers(sj)
        sortscores(sj)=stemp
        sortplayers(sj)=stempplayers
      end if
    next
  next
  ScoreChecker=4
  CheckAllScores=1
  NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)
  savehs
end sub


sub checkmatch
  Dim tempmatch
  Exit Sub
  tempmatch=Int(Rnd*10)
  Match=tempmatch*10
  MatchReel.SetValue(tempmatch+1)

  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 100
    Else
      Controller.B2SSetMatch Match
    End If
  End if
  for i = 1 to Players
    if Match=(Score(i) mod 100) then
      AddSpecial
    end if
  next
end sub

Sub TiltTimer_Timer()
  if TiltCount > 0 then TiltCount = TiltCount - 1
  if TiltCount = 0 then
    TiltTimer.Enabled = False
  end if
end sub

Sub TiltIt()
    If TableTilted=true then
        exit sub
    end if
    TiltCount = TiltCount + 1
    if TiltCount = 3 then
        TableTilted=True
        TiltReel.SetValue(1)
    If VRMode Then
      VR_Tilt.visible = 1
    End If
        If TiltEndsGame = 1 then
            If WOWCounter > 1 then
        If B2SOn Then
          Controller.B2SSetData WOWCounter+50,0
        End If
                WOWCounter=WOWCounter-1
            ElseIf BallInPlay > 1 then
                BallInPlay = BallInPlay - 1
                BallInPlayReel.SetValue(BallInPlay)
                if B2SOn then
                    Controller.B2SSetBallInPlay BallInPlay
                end if
        If VRMode Then
          VR_Balls1.visible = 0
          VR_Balls2.visible = 0
          VR_Balls3.visible = 0
          VR_Balls4.visible = 0
          VR_Balls5.visible = 0
          Dim vr_bip : vr_bip = eval("VR_Balls" & BallInPlay)
          vr_bip.visible = 1
        End If
            end if
        end if
        PlasticsOff
        BumpersOff

        LeftFlipper.RotateToStart

        StopSound "buzzL"

        RightFlipper.RotateToStart

        StopSound "buzz"

        If B2Son then
            Controller.B2SSetTilt 1
        end if
    else
        TiltTimer.Interval = 500
        TiltTimer.Enabled = True
    end if

end sub

Sub IncreaseBonus()


  If BonusCounter=15 then
    exit sub
  end if
  If BonusCounter<10 then
    Bonus(BonusCounter).state=0
  end if
  if BonusCounter>10 then
    Bonus(BonusCounter-10).state=0
  end if
  BonusCounter=BonusCounter+1

  if BonusCounter>10 then
    Bonus(10).State=1
    Bonus(BonusCounter-10).state=1
  else
    Bonus(BonusCounter).State=1
  end if


  if BonusMultiplier=2 then
    DoubleBonus.state=1
  end if


End Sub


Sub BonusBoost_Timer()
  IncreaseBonus
  BonusBoosterCounter=BonusBoosterCounter-1
  If BonusBoosterCounter=0 then
    BonusBoost.enabled=false
  end if

end sub

Sub CheckForLightSpecial()

  if (TopLightA.state=0) and (TopLightB.state=0) and (TopLightC.state=0) then
    TopRightTargetLight.State=1
    TopLeftTargetLight.State=1
  end if
end sub

Sub PlayStartBall_timer()

  PlayStartBall.enabled=false
  PlaySound("StartBall2-5")
end sub

Sub PlayerUpRotator_timer()
  If RotatorTemp<5 then
    TempPlayerUp=TempPlayerUp+1
    If TempPlayerUp>4 then
      TempPlayerUp=1
    end if
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    For each obj in PlayerHUDScores
      obj.state=0
    next
    If Table1.ShowDT = True then
      For each obj in PlayerScores
        obj.visible=1
      Next
      For each obj in PlayerScoresOn
        obj.visible=0
      Next
      PlayerHuds(TempPlayerUp-1).SetValue(1)
      PlayerHUDScores(TempPlayerUp-1).state=1
      PlayerScores(TempPlayerUp-1).visible=0
      PlayerScoresOn(TempPlayerUp-1).visible=1
    end If
    If B2SOn Then
      Controller.B2SSetPlayerUp TempPlayerUp
      Controller.B2SSetData 81,0
      Controller.B2SSetData 82,0
      Controller.B2SSetData 83,0
      Controller.B2SSetData 84,0
      Controller.B2SSetData 80+TempPlayerUp,1
    End If

  else
    if B2SOn then
      Controller.B2SSetPlayerUp Player
      Controller.B2SSetData 81,0
      Controller.B2SSetData 82,0
      Controller.B2SSetData 83,0
      Controller.B2SSetData 84,0
      Controller.B2SSetData 80+Player,1
    end if
    PlayerUpRotator.enabled=false
    RotatorTemp=1
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    For each obj in PlayerHUDScores
      obj.state=0
    next
    If Table1.ShowDT = True then
      For each obj in PlayerScores
        obj.visible=1
      Next
      For each obj in PlayerScoresOn
        obj.visible=0
      Next
      PlayerHuds(Player-1).SetValue(1)
      PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).visible=0
      PlayerScoresOn(Player-1).visible=1
    end If
  end if
  RotatorTemp=RotatorTemp+1


end sub

sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine Credits
    scorefile.writeline BallsPerGame
    scorefile.writeline ReplayLevel
    scorefile.writeline TiltEndsGame
    scorefile.writeline NoveltyMode
    for xx=1 to 5
      scorefile.writeline HSScore(xx)
    next
    for xx=1 to 5
      scorefile.writeline HSName(xx)
    next
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub loadhs
    ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
    dim temp1
    dim temp2
  dim temp3
  dim temp4
  dim temp5

  dim temp6
  dim temp7
  dim temp8
  dim temp9
  dim temp10
  dim temp11
  dim temp12
  dim temp13
  dim temp14
  dim temp15
  dim temp16
  dim temp17

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HSFileName) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    temp1=TextStr.ReadLine
    temp2=textstr.readline
    temp3=textstr.readline
    temp4=textstr.readline
    temp5=textstr.readline
    temp16=textstr.readline
    HighScore=cdbl(temp1)
    if HighScore<1 then
      temp6=textstr.readline
      temp7=textstr.readline
      temp8=textstr.readline
      temp9=textstr.readline
      temp10=textstr.readline
      temp11=textstr.readline
      temp12=textstr.readline
      temp13=textstr.readline
      temp14=textstr.readline
      temp15=textstr.readline
    end if
    TextStr.Close

      Credits=cdbl(temp2)
    BallsPerGame=cdbl(temp3)
    ReplayLevel=cdbl(temp4)
    TiltEndsGame=cdbl(temp5)
    NoveltyMode=cdbl(temp16)

    if HighScore<1 then
      HSScore(1) = int(temp6)
      HSScore(2) = int(temp7)
      HSScore(3) = int(temp8)
      HSScore(4) = int(temp9)
      HSScore(5) = int(temp10)

      HSName(1) = temp11
      HSName(2) = temp12
      HSName(3) = temp13
      HSName(4) = temp14
      HSName(5) = temp15
    end if

    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

sub SaveLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim temp1
  dim tempb2s
  tempb2s=0
  if B2SOn=true then
    tempb2s=1
  else
    tempb2s=0
  end if
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set LMConfig=FileObj.CreateTextFile(UserDirectory & LMEMTableConfig,True)
  LMConfig.WriteLine tempb2s
  LMConfig.Close
  Set LMConfig=Nothing
  Set FileObj=Nothing

end Sub

sub LoadLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim tempC
  dim tempb2s

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & LMEMTableConfig) then
    Exit Sub
  End if
  Set LMConfig=FileObj.GetFile(UserDirectory & LMEMTableConfig)
  Set TextStr2=LMConfig.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) then
    Exit Sub
  End if
  tempC=TextStr2.ReadLine
  TextStr2.Close
  tempb2s=cdbl(tempC)
  if tempb2s=0 then
    B2SOn=false
  else
    B2SOn=true
  end if
  Set LMConfig=Nothing
  Set FileObj=Nothing
end sub

sub SaveLMEMConfig2
  If ShadowConfigFile=false then exit sub
  Dim FileObj
  Dim LMConfig2
  dim temp1
  dim temp2
  dim tempBS
  dim tempFS

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set LMConfig2=FileObj.CreateTextFile(UserDirectory & LMEMShadowConfig,True)
  LMConfig2.WriteLine tempBS
  LMConfig2.WriteLine tempFS
  LMConfig2.Close
  Set LMConfig2=Nothing
  Set FileObj=Nothing

end Sub

sub LoadLMEMConfig2

  Dim FileObj
  Dim LMConfig2
  dim tempC
  dim tempD
  dim tempFS
  dim tempBS

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & LMEMShadowConfig) then
    Exit Sub
  End if
  Set LMConfig2=FileObj.GetFile(UserDirectory & LMEMShadowConfig)
  Set TextStr2=LMConfig2.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) then
    Exit Sub
  End if
  tempC=TextStr2.ReadLine
  tempD=TextStr2.Readline
  TextStr2.Close
  tempBS=cdbl(tempC)
  tempFS=cdbl(tempD)

  Set LMConfig2=Nothing
  Set FileObj=Nothing

end sub


Sub DisplayHighScore

end sub

sub InitPauser5_timer
  If B2SOn Then
    Controller.B2SSetScorePlayer 2,HSScore(1)
  End If
  HSReel.SetValue(HSScore(1))
  DisplayHighScore
  CreditsReel.SetValue(Credits)
  InitPauser5.enabled=false
end sub



sub BumpersOff
  Bumper1Light.state=0
  Bumper2Light.state=0


end sub

sub BumpersOn
  Bumper1Light.state=1
  Bumper2Light.state=1
end sub

Sub PlasticsOn

  For each obj in Flashers
    obj.Visible=1
  next
  Light1.state=1
  Light2.state=1

end sub

Sub PlasticsOff
  For each obj in Flashers
    obj.Visible=0
  next
  Light1.state=0
  Light2.state=0
  StopSound "buzzL"
    StopSound "buzz"
end sub
Sub SetupReplayTables

  Replay1Table(1)=30000
  Replay1Table(2)=30000
  Replay1Table(3)=40000
  Replay1Table(4)=50000
  Replay1Table(5)=50000
  Replay1Table(6)=60000
  Replay1Table(7)=70000
  Replay1Table(8)=80000
  Replay1Table(9)=61000
  Replay1Table(10)=65000
  Replay1Table(11)=69000
  Replay1Table(12)=74000
  Replay1Table(13)=77000
  Replay1Table(14)=83000
  Replay1Table(15)=999000

  Replay2Table(1)=70000
  Replay2Table(2)=80000
  Replay2Table(3)=90000
  Replay2Table(4)=100000
  Replay2Table(5)=100000
  Replay2Table(6)=120000
  Replay2Table(7)=120000
  Replay2Table(8)=130000
  Replay2Table(9)=79000
  Replay2Table(10)=83000
  Replay2Table(11)=87000
  Replay2Table(12)=92000
  Replay2Table(13)=98000
  Replay2Table(14)=99000
  Replay2Table(15)=999000

  Replay3Table(1)=100000
  Replay3Table(2)=120000
  Replay3Table(3)=130000
  Replay3Table(4)=140000
  Replay3Table(5)=160000
  Replay3Table(6)=170000
  Replay3Table(7)=160000
  Replay3Table(8)=170000
  Replay3Table(9)=999000
  Replay3Table(10)=999000
  Replay3Table(11)=999000
  Replay3Table(12)=999000
  Replay3Table(13)=999000
  Replay3Table(14)=999000
  Replay3Table(15)=999000

  ReplayTableMax=8

end sub

Sub RefreshReplayCard
  Dim tempst1
  Dim tempst2

  tempst1=FormatNumber(BallsPerGame,0)
  tempst2=FormatNumber(ReplayLevel,0)

  ReplayCard1.imagea = "SC" + tempst2
  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
end sub


'****************************************
'  SCORE MOTOR
'****************************************

ScoreMotorTimer.Enabled = 1
ScoreMotorTimer.Interval = 145
AddScoreTimer.Enabled = 1
AddScoreTimer.Interval = 145

Dim queuedscore
Dim MotorMode
Dim MotorPosition

Const BIG_AWARD_50K = 50000
Dim FiftyKBurstActive : FiftyKBurstActive = False

Sub Clear50KBurst()
    FiftyKBurstActive = False
End Sub

Sub SetMotor(y)
  Select Case ScoreMotorAdjustment
    Case 0:
      queuedscore=queuedscore+y
    Case 1:
      If MotorRunning<>1 And InProgress=true then
        queuedscore=queuedscore+y
      end if
    end Select
end sub

Sub SetMotor2(x)
  If MotorRunning<>1 And InProgress=true then
    MotorRunning=1

    Select Case x
      Case 10000:
        AddScore(10000)
        MotorMode = 10000
        MotorPosition = 0
        BumpersOff
      Case 5000:
        MotorMode = 1000
        MotorPosition = 5
        BumpersOff
      Case 4000:
        MotorMode = 1000
        MotorPosition = 4
        BumpersOff
      Case 3000:
        MotorMode = 1000
        MotorPosition = 3
        BumpersOff
      Case 2000:
        MotorMode = 1000
        MotorPosition = 2
        BumpersOff
      Case 1000:
        AddScore(1000)
        MotorRunning = 0
        BumpersOn
      Case 500:
        MotorMode = 100
        MotorPosition = 5
        BumpersOff
      Case 400:
        MotorMode = 100
        MotorPosition = 4
        BumpersOff
      Case 300:
        MotorMode = 100
        MotorPosition = 3
        BumpersOff
      Case 200:
        MotorMode = 100
        MotorPosition = 2
        BumpersOff
      Case 100:
        AddScore(100)
        MotorRunning = 0
        BumpersOn
      Case 50:
        MotorMode = 10
        MotorPosition = 5
        BumpersOff
      Case 40:
        MotorMode = 10
        MotorPosition = 4
        BumpersOff
      Case 30:
        MotorMode = 10
        MotorPosition = 3
        BumpersOff
      Case 20:
        MotorMode = 10
        MotorPosition = 2
        BumpersOff
      Case 10:
        AddScore(10)
        MotorRunning = 0
        BumpersOn
    End Select
  End If
End Sub

Sub AddScoreTimer_Timer
  Dim tempscore


  If MotorRunning<>1 And InProgress=true then
    'debug.Print("queuedscore: "&queuedscore)

    If queuedscore >= 10000 Then
      tempscore = 10000
      queuedscore = queuedscore - 10000
      SetMotor2(10000)

      KnockerSolenoid

      MotorPosition=0:MotorRunning=0:BumpersOn
      dim newqueuedscore : newqueuedscore = queuedscore
      queuedscore = 0
      SetMotor(newqueuedscore)
      Exit Sub
    End If

    if queuedscore>=5000 then
      tempscore=5000
      queuedscore=queuedscore-5000
      SetMotor2(5000)
      exit sub
    end if
    if queuedscore>=4000 then
      tempscore=4000
      queuedscore=queuedscore-4000
      SetMotor2(4000)
      exit sub
    end if

    if queuedscore>=3000 then
      tempscore=3000
      queuedscore=queuedscore-3000
      SetMotor2(3000)
      exit sub
    end if

    if queuedscore>=2000 then
      tempscore=2000
      queuedscore=queuedscore-2000
      SetMotor2(2000)
      exit sub
    end if

    if queuedscore>=1000 then
      tempscore=1000
      queuedscore=queuedscore-1000
      SetMotor2(1000)
      exit sub
    end if

    if queuedscore>=500 then
      tempscore=500
      queuedscore=queuedscore-500
      SetMotor2(500)
      exit sub
    end if
    if queuedscore>=400 then
      tempscore=400
      queuedscore=queuedscore-400
      SetMotor2(400)
      exit sub
    end if
    if queuedscore>=300 then
      tempscore=300
      queuedscore=queuedscore-300
      SetMotor2(300)
      exit sub
    end if
    if queuedscore>=200 then
      tempscore=200
      queuedscore=queuedscore-200
      SetMotor2(200)
      exit sub
    end if
    if queuedscore>=100 then
      tempscore=100
      queuedscore=queuedscore-100
      SetMotor2(100)
      exit sub
    end if

    if queuedscore>=50 then
      tempscore=50
      queuedscore=queuedscore-50
      SetMotor2(50)
      exit sub
    end if
    if queuedscore>=40 then
      tempscore=40
      queuedscore=queuedscore-40
      SetMotor2(40)
      exit sub
    end if
    if queuedscore>=30 then
      tempscore=30
      queuedscore=queuedscore-30
      SetMotor2(30)
      exit sub
    end if
    if queuedscore>=20 then
      tempscore=20
      queuedscore=queuedscore-20
      SetMotor2(20)
      exit sub
    end if
    if queuedscore>=10 then
      tempscore=10
      queuedscore=queuedscore-10
      SetMotor2(10)
      exit sub
    end if


  End If
end Sub

Sub ScoreMotorTimer_Timer
  If MotorPosition > 0 Then
    Select Case MotorPosition
      Case 5,4,3,2:
        If MotorMode=1000 Then
          AddScore(1000)
        end if
        if MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)
        End if
        MotorPosition=MotorPosition-1
      Case 1:
        If MotorMode=1000 Then
          AddScore(1000)
        end if
        If MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)
        End if
        MotorPosition=0:MotorRunning=0:BumpersOn
    End Select
  End If

End Sub

Sub AddScore(x)
  If TableTilted=true then exit sub
  Select Case ScoreAdditionAdjustment
    Case 0:
      AddScore1(x)
    Case 1:
      AddScore2(x)
  end Select

end sub


Sub AddScore1(x)
' debugtext.text=score
  Select Case x
    Case 1:
      PlayChime(10)
      Score(Player)=Score(Player)+1

    Case 10:
      PlayChime(10)
      Score(Player)=Score(Player)+10
'     debugscore=debugscore+10
      'ToggleAlternatingRelay
    Case 100:
      PlayChime(100)
      Score(Player)=Score(Player)+100
'     debugscore=debugscore+100

    Case 1000:
      PlayChime(1000)
      Score(Player)=Score(Player)+1000
'     debugscore=debugscore+1000
    Case 10000:
      PlayChime(1000)
      KnockerSolenoid
      Score(Player)=Score(Player)+10000
  End Select
  PlayerScores(Player-1).AddValue(x)
  PlayerScoresOn(Player-1).AddValue(x)
  If ScoreDisplay(Player)<100000 then
    ScoreDisplay(Player)=Score(Player)
  Else
    Score100K(Player)=Int(Score(Player)/100000)
    ScoreDisplay(Player)=Score(Player)-100000
  End If
  if Score(Player)=>100000 then
    If B2SOn Then
      If Player=1 Then
        Controller.B2SSetScoreRolloverPlayer1 Score100K(Player)
      End If
      If Player=2 Then
        Controller.B2SSetScoreRolloverPlayer2 Score100K(Player)
      End If

      If Player=3 Then
        Controller.B2SSetScoreRolloverPlayer3 Score100K(Player)
      End If

      If Player=4 Then
        Controller.B2SSetScoreRolloverPlayer4 Score100K(Player)
      End If
    End If
  End If
  If B2SOn Then
    Controller.B2SSetScorePlayer Player, ScoreDisplay(Player)
  End If
  If Score(Player)>Replay1 and Replay1Paid(Player)=false And NoveltyMode=0 then
    Replay1Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay2 and Replay2Paid(Player)=false And NoveltyMode=0 then
    Replay2Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay3 and Replay3Paid(Player)=false And NoveltyMode=0 then
    Replay3Paid(Player)=True
    AddSpecial
  End If
' ScoreText.text=debugscore
End Sub

Sub AddScore2(x)
  Dim OldScore, NewScore, OldTestScore, NewTestScore
    OldScore = Score(Player)

  Select Case x
        Case 1:
            Score(Player)=Score(Player)+1
    Case 10:
      Score(Player)=Score(Player)+10
    Case 100:
      Score(Player)=Score(Player)+100
    Case 1000:
      Score(Player)=Score(Player)+1000
    Case 10000:
      Score(Player)=Score(Player)+10000
  End Select
  NewScore = Score(Player)
  if Score(Player)=>100000 then
    Dim ReelValue : ReelValue = NewScore/100000
    RolloverReel.SetValue(Int(ReelValue))
    UpdateVRRollover ReelValue
    If B2SOn Then
      If Player=1 Then
        Controller.B2SSetScoreRolloverPlayer1 (Int(Score(Player)/100000))
      End If
      If Player=2 Then
        Controller.B2SSetScoreRolloverPlayer2 1
      End If

      If Player=3 Then
        Controller.B2SSetScoreRolloverPlayer3 1
      End If

      If Player=4 Then
        Controller.B2SSetScoreRolloverPlayer4 1
      End If
    End If
  End If

  OldTestScore = OldScore
  NewTestScore = NewScore
  Do
    if OldTestScore < Replay1 and NewTestScore >= Replay1 And NoveltyMode = 0 then
      AddSpecial()
      NewTestScore = 0
    Elseif OldTestScore < Replay2 and NewTestScore >= Replay2 And NoveltyMode = 0 then
      AddSpecial()
      NewTestScore = 0
    Elseif OldTestScore < Replay3 and NewTestScore >= Replay3 And NoveltyMode = 0 then
      AddSpecial()
      NewTestScore = 0
    End if
    NewTestScore = NewTestScore - 100000
    OldTestScore = OldTestScore - 100000
  Loop While NewTestScore > 0

    OldScore = int(OldScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
    NewScore = int(NewScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore&", OldScore Mod 10="&OldScore Mod 10 & ", NewScore % 10="&NewScore Mod 10)

  Dim oldTenThousands, newTenThousands, oldThousands, newThousands
  Dim oldHundreds, newHundreds, oldTens, newTens
  oldTenThousands = (OldScore \ 1000) Mod 10
  newTenThousands = (NewScore \ 1000) Mod 10
  oldThousands = (OldScore \ 100) Mod 10
  newThousands = (NewScore \ 100) Mod 10
  oldHundreds = (OldScore \ 10) Mod 10
  newHundreds = (NewScore \ 10) Mod 10
  oldTens = (OldScore) Mod 10
  newTens = (NewScore) Mod 10

  ' Ten-Thousands (knocker)
  If oldTenThousands <> newTenThousands Then
    KnockerSolenoid
  End If

  ' Thousands
  If oldThousands <> newThousands Then
    PlayChime 1000
  End If

  ' Hundreds
  If oldHundreds <> newHundreds Then
    PlayChime 100
  End If

  ' Tens
  If oldTens <> newTens Then
    PlayChime 10
  End If

  If B2SOn Then
    Controller.B2SSetScorePlayer Player, Score(Player)
  End If
' EMReel1.SetValue Score(Player)
  PlayerScores(Player-1).AddValue(x)
  PlayerScoresOn(Player-1).AddValue(x)
End Sub

Sub UpdateVRRollover(rv)
  VR_100k.visible = 0
  VR_200k.visible = 0
  VR_300k.visible = 0
  VR_400k.visible = 0
  VR_500k.visible = 0
  VR_600k.visible = 0
  VR_700k.visible = 0
  VR_800k.visible = 0
  VR_900k.visible = 0
  VR_1M.visible = 0
  If rv >= 1 And rv < 2 Then
    VR_100k.visible = 1
  ElseIf rv >= 2 And rv < 3 Then
    VR_200k.visible = 1
  ElseIf rv >= 3 And rv < 4 Then
    VR_300k.visible = 1
  ElseIf rv >= 4 And rv < 5 Then
    VR_400k.visible = 1
  ElseIf rv >= 5 And rv < 6 Then
    VR_500k.visible = 1
  ElseIf rv >= 6 And rv < 7 Then
    VR_600k.visible = 1
  ElseIf rv >= 7 And rv < 8 Then
    VR_700k.visible = 1
  ElseIf rv >= 8 And rv < 9 Then
    VR_800k.visible = 1
  ElseIf rv >= 9 And rv < 10 Then
    VR_900k.visible = 1
  ElseIf rv >= 10 Then
    VR_1M.visible = 1
  End If
End Sub


Sub PlayChime(x)
  if ChimesOn=0 then
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",129,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",129,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",130,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",130,DOFPulse,DOFChimes)
          LastChime100=1
        End If

    End Select
  else
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SJ_Chime_10a",129,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_10b",129,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SJ_Chime_100a",130,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_100b",130,DOFPulse,DOFChimes)
          LastChime100=1
        End If
      Case 1000
        If LastChime1000=1 Then
          PlaySound SoundFXDOF("SJ_Chime_1000a",131,DOFPulse,DOFChimes)
          LastChime1000=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_1000b",131,DOFPulse,DOFChimes)
          LastChime1000=1
        End If
    End Select
  end if
  EMMODE = 1
  UpdateReels 0,0 ,score(0), 0, -1,-1,-1,-1,-1
  UpdateReels 1,1 ,score(1), 0, -1,-1,-1,-1,-1
  UpdateReels 2,2 ,score(2), 0, -1,-1,-1,-1,-1
  UpdateReels 3,3 ,score(3), 0, -1,-1,-1,-1,-1
  UpdateReels 4,4 ,score(4), 0, -1,-1,-1,-1,-1
  EMMODE = 0 ' restore EM mode
End Sub


' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' ============================================================================================
Dim EnteringInitials    ' Normally zero, set to non-zero to enter initials
EnteringInitials = 0

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar      ' character under the "cursor" when entering initials

Dim HSTimerCount      ' Pass counter for HS timer, scores are cycled by the timer
HSTimerCount = 5      ' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString     ' the string holding the player's initials as they're entered

Dim AlphaString       ' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos      ' pointer to AlphaString, move forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh       ' The new score to be recorded





Sub HighScoreTimer_Timer

  if EnteringInitials then
    if HSTimerCount = 1 then
      SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
      HSTimerCount = 2
    else
      SetHSLine 3, InitialString
      HSTimerCount = 1
    end if
  elseif InProgress then
    SetHSLine 1, "HIGH SCORE1"
    SetHSLine 2, HSScore(1)
    SetHSLine 3, HSName(1)
    HSTimerCount = 5  ' set so the highest score will show after the game is over
    HighScoreTimer.enabled=false
  elseif CheckAllScores then
    NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)

  else
    ' cycle through high scores
    HighScoreTimer.interval=2000
    HSTimerCount = HSTimerCount + 1
    if HsTimerCount > 5 then
      HSTimerCount = 1
    End If
    SetHSLine 1, "HIGH SCORE"+FormatNumber(HSTimerCount,0)
    SetHSLine 2, HSScore(HSTimerCount)
    SetHSLine 3, HSName(HSTimerCount)
    If B2SOn Then
      Controller.B2SSetScorePlayer 2,HSScore(1)
    End If
    HSReel.SetValue(HSScore(1))
  end if
End Sub

Function GetHSChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  else
    FileName = FileName & ThisChar
  End If
  GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim StrLen
  dim LetterLine
  dim Index
  dim StartHSArray
  dim EndHSArray
  dim LetterName
  dim xfor
  StartHSArray=array(0,1,12,22)
  EndHSArray=array(0,11,21,31)
  StrLen = len(string)
  Index = 1

  for xfor = StartHSArray(LineNo) to EndHSArray(LineNo)
    Eval("HS"&xfor).image = GetHSChar(String, Index)
    Index = Index + 1
  next

End Sub

Sub NewHighScore(NewScore, PlayNum)
  if NewScore > HSScore(5) then
    HighScoreTimer.interval = 500
    HSTimerCount = 1
    AlphaStringPos = 1    ' start with first character "A"
    EnteringInitials = 1  ' intercept the control keys while entering initials
    InitialString = ""    ' initials entered so far, initialize to empty
    SetHSLine 1, "PLAYER "+FormatNumber(PlayNum,0)
    SetHSLine 2, "ENTER NAME"
    SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
    HSNewHigh = NewScore
    For xx=1 to HighScoreReward
'     AddSpecial
    next
  End if
  ScoreChecker=ScoreChecker-1
  if ScoreChecker=0 then
    CheckAllScores=0
  end if
End Sub

Sub CollectInitials(keycode)
  If keycode = LeftFlipperKey Then
    ' back up to previous character
    AlphaStringPos = AlphaStringPos - 1
    if AlphaStringPos < 1 then
      AlphaStringPos = len(AlphaString)   ' handle wrap from beginning to end
      if InitialString = "" then
        ' Skip the backspace if there are no characters to backspace over
        AlphaStringPos = AlphaStringPos - 1
      End if
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = RightFlipperKey Then
    ' advance to next character
    AlphaStringPos = AlphaStringPos + 1
    if AlphaStringPos > len(AlphaString) or (AlphaStringPos = len(AlphaString) and InitialString = "") then
      ' Skip the backspace if there are no characters to backspace over
      AlphaStringPos = 1
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = StartGameKey or keycode = PlungerKey Then
    SelectedChar = MID(AlphaString, AlphaStringPos, 1)
    if SelectedChar = "_" then
      InitialString = InitialString & " "
      PlaySound("Ding10")
    elseif SelectedChar = "<" then
      InitialString = MID(InitialString, 1, len(InitialString) - 1)
      if len(InitialString) = 0 then
        ' If there are no more characters to back over, don't leave the < displayed
        AlphaStringPos = 1
      end if
      PlaySound("Ding100")
    else
      InitialString = InitialString & SelectedChar
      PlaySound("Ding10")
    end if
    if len(InitialString) < 3 then
      SetHSLine 3, InitialString & SelectedChar
    End If
  End If
  if len(InitialString) = 3 then
    ' save the score
    for i = 5 to 1 step -1
      if i = 1 or (HSNewHigh > HSScore(i) and HSNewHigh <= HSScore(i - 1)) then
        ' Replace the score at this location
        if i < 5 then
' MsgBox("Moving " & i & " to " & (i + 1))
          HSScore(i + 1) = HSScore(i)
          HSName(i + 1) = HSName(i)
        end if
' MsgBox("Saving initials " & InitialString & " to position " & i)
        EnteringInitials = 0
        HSScore(i) = HSNewHigh
        HSName(i) = InitialString
        HSTimerCount = 5
        HighScoreTimer_Timer
        HighScoreTimer.interval = 2000
        PlaySound("Ding1000")
        exit sub
      elseif i < 5 then
        ' move the score in this slot down by 1, it's been exceeded by the new score
' MsgBox("Moving " & i & " to " & (i + 1))
        HSScore(i + 1) = HSScore(i)
        HSName(i + 1) = HSName(i)
      end if
    next
  End If

End Sub
' END GNMOD


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

'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light.
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(5)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 3 + iii / 1000
    objBallShadow(iii).visible = 0
  Next
End Sub


Sub BSUpdate
  Dim s: For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If gBOT(s).Z > 20 And gBOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity


'*******************************************
' Late 70's to early 80's

Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger


        x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, - 2.7
        x.AddPt "Polarity", 2, 0.16, - 2.7
        x.AddPt "Polarity", 3, 0.22, - 0
        x.AddPt "Polarity", 4, 0.25, - 0
        x.AddPt "Polarity", 5, 0.3, - 1
        x.AddPt "Polarity", 6, 0.4, - 2
        x.AddPt "Polarity", 7, 0.5, - 2.7
        x.AddPt "Polarity", 8, 0.65, - 1.8
        x.AddPt "Polarity", 9, 0.75, - 0.5
        x.AddPt "Polarity", 10, 0.81, - 0.5
        x.AddPt "Polarity", 11, 0.88, 0
        x.AddPt "Polarity", 12, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.15, 0.85
    x.AddPt "Velocity", 2, 0.2, 0.9
    x.AddPt "Velocity", 3, 0.23, 0.95
    x.AddPt "Velocity", 4, 0.41, 0.95
    x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
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
  '   Dim BOT
  '   BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
          gBOT(b).velx = gBOT(b).velx / 1.3
          gBOT(b).vely = gBOT(b).vely - 0.5
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

Const FlipperCoilRampupMode = 2 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
'Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
 Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

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
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b', BOT
    '   BOT = GetBalls

    For b = 0 To UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
      End If
    Next
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


'******************************************************
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' The Stand Up and Drop Target solutions improve the physics for targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full target animation, switch handling and deflection on hit. For drop targets there is also a slight lift when
' the drop targets raise, bricking, and popping the ball up if it's over the drop target when it raises.
'
' Add a Timers named DTAnim and STAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
' DTAnim.interval = 10
' DTAnim.enabled = True

' Sub DTAnim_Timer
'   DoDTAnim
' DoSTAnim
' End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.
'
' For each stand up target we'll use a vp target, a laid back collidable primitive, and one primitive for visuals and animation.
' The visual primitive should should have it's pivot point centered on the x and y axis and the z should be at or just below the playfield.
' The target should animate backwards using transy.
'
' To create visual target primitives that work with the stand up and drop target code, follow the below instructions:
' (Other methods will work as well, but this is easy for even non-blender users to do)
' 1) Open a new blank table. Delete everything off the table in editor.
' 2) Copy and paste the VP target from your table into this blank table.
' 3) Place the target at x = 0, y = 0  (upper left hand corner) with an orientation of 0 (target facing the front of the table)
' 4) Under the file menu, select Export "OBJ Mesh"
' 5) Go to "https://threejs.org/editor/". Here you can modify the exported obj file. When you export, it exports your target and also
'    the playfield mesh. You need to delete the playfield mesh here. Under the file menu, chose import, and select the obj you exported
'    from VPX. In the right hand panel, find the Playfield object and click on it and delete. Then use the file menu to Export OBJ.
' 6) In VPX, you can add a primitive and use "Import Mesh" to import the exported obj from the previous step. X,Y,Z scale should be 1.
'    The primitive will use the same target texture as the VP target object.
'
' * Note, each target must have a unique switch number. If they share a same number, add 100 to additional target with that number.
' For example, three targets with switch 32 would use 32, 132, 232 for their switch numbers.
' The 100 and 200 will be removed when setting the switch value for the target.

'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

Class DropTarget
  Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class

'Define a variable for each drop target
Dim DT1, DT2, DT3, DT4, DT5, DT6, DT7, DT8, DT9, DT10

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:      primary target wall to determine drop
' secondary:      wall used to simulate the ball striking a bent or offset target after the initial Hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             rotz must be used for orientation
'             rotx to bend the target back
'             transz to move it up and down
'             the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
' switch:       ROM switch number
' animate:      Array slot for handling the animation instrucitons, set to 0
'           Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:      Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.

Set DT1 = (new DropTarget)(sw1, sw1a, BM_DT1, 1, 0, false)
Set DT2 = (new DropTarget)(sw2, sw2a, BM_DT2, 2, 0, false)
Set DT3 = (new DropTarget)(sw3, sw3a, BM_DT3, 3, 0, false)
Set DT4 = (new DropTarget)(sw4, sw4a, BM_DT4, 4, 0, false)
Set DT5 = (new DropTarget)(sw5, sw5a, BM_DT5, 5, 0, false)
Set DT6 = (new DropTarget)(sw6, sw6a, BM_DT6, 6, 0, false)
Set DT7 = (new DropTarget)(sw7, sw7a, BM_DT7, 7, 0, false)
Set DT8 = (new DropTarget)(sw8, sw8a, BM_DT8, 8, 0, false)
Set DT9 = (new DropTarget)(sw9, sw9a, BM_DT9, 9, 0, false)
Set DT10 = (new DropTarget)(sw10, sw10a, BM_DT10, 10, 0, false)

Dim DTArray
DTArray = Array(DT1, DT2, DT3, DT4, DT5, DT6, DT7, DT8, DT9, DT10)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 120 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 48 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 10 'max degrees primitive rotates when hit
Const DTDropDelay = 40 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTMinBallVel = 5
Const DTEnableBrick = 1 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "target" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************
Sub DTHit(switch)

  Dim i
  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
  If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate =  - 1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 To UBound(DTArray)
    If DTArray(i).sw = switch Then
      DTArrayID = i
      Exit Function
    End If
  Next
End Function

Sub DTBallPhysics(aBall, angle, mass)
  Dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
  calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)

  aBall.velx = calc1 * Cos(rangle) + calc2
  aBall.vely = calc1 * Sin(rangle) + calc3
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
  Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > DTMinBallVel And  perpvelafter <= 0 Then
    If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > DTMinBallVel And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function

Sub DoDTAnim()
  Dim i
  For i = 0 To UBound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
  Dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  Dim ind
  ind = DTArrayID(switchid)

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  If animate = 0 Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    DTAnimate = animate
    Exit Function
  ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 1 'If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0 'updated by rothbauerw to account for edge case
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  If animate = 2 Then
    transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
    If prim.transz >  - DTDropUnits  Then
      prim.transz = transz
    End If

    prim.rotx = DTMaxBend * Cos(rangle) / 2
    prim.roty = DTMaxBend * Sin(rangle) / 2

    If prim.transz <= - DTDropUnits Then
      prim.transz =  - DTDropUnits
      secondary.collidable = 0
      DTArray(ind).isDropped = True 'Mark target as dropped
      ' controller.Switch(Switchid mod 100) = 1
      ScoreDrop Switchid
      primary.uservalue = 0
      DTAnimate = 0
      CheckAllDrops
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    End If
  End If

  If animate = 3 And animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
  ElseIf animate = 3 And animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  If animate =  - 1 Then
    transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1

    If prim.transz =  - DTDropUnits Then
      Dim b
      'Dim gBOT
      'gBOT = GetBalls

      For b = 0 To UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
          gBOT(b).velz = 20
        End If
      Next
    End If

    If prim.transz < 0 Then
      prim.transz = transz
    ElseIf transz > 0 Then
      prim.transz = transz
    End If

    If prim.transz > DTDropUpUnits Then
      DTAnimate =  - 2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = GameTime
    End If
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = False 'Mark target as not dropped
    ' controller.Switch(Switchid mod 100) = 0
  End If

  If animate =  - 2 And animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
    If prim.transz < 0 Then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    End If
  End If
End Function

Function DTDropped(switchid)
  Dim ind
  ind = DTArrayID(switchid)

  DTDropped = DTArray(ind).isDropped
End Function

Sub UpdateDropTargets
  dim BP, tz, rx, ry

    tz = BM_DT1.transz
  rx = BM_DT1.rotx
  ry = BM_DT1.roty
  For each BP in BP_DT1 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_DT2.transz
  rx = BM_DT2.rotx
  ry = BM_DT2.roty
  For each BP in BP_DT2 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_DT3.transz
  rx = BM_DT3.rotx
  ry = BM_DT3.roty
  For each BP in BP_DT3 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT4.transz
  rx = BM_DT4.rotx
  ry = BM_DT4.roty
  For each BP in BP_DT4: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT5.transz
  rx = BM_DT5.rotx
  ry = BM_DT5.roty
  For each BP in BP_DT5: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT6.transz
  rx = BM_DT6.rotx
  ry = BM_DT6.roty
  For each BP in BP_DT6: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT7.transz
  rx = BM_DT7.rotx
  ry = BM_DT7.roty
  For each BP in BP_DT7: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT8.transz
  rx = BM_DT8.rotx
  ry = BM_DT8.roty
  For each BP in BP_DT8: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT9.transz
  rx = BM_DT9.rotx
  ry = BM_DT9.roty
  For each BP in BP_DT9: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT10.transz
  rx = BM_DT10.rotx
  ry = BM_DT10.roty
  For each BP in BP_DT10: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub


'******************************************************
'****  END DROP TARGETS
'******************************************************




'******************************************************
' ZRST: STAND-UP TARGET INITIALIZATION
'******************************************************

Class StandupTarget
  Private m_primary, m_prim, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class

'Define a variable for each stand-up target
Dim ST11

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:      vp target to determine target hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             transy must be used to offset the target animation
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST11 = (new StandupTarget)(sw11, BM_ST11, 11, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST11)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  PlayTargetSound
  STArray(i).animate = STCheckHit(ActiveBall,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics ActiveBall, STArray(i).primary.orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 To UBound(STArray)
    If STArray(i).sw = switch Then
      STArrayID = i
      Exit Function
    End If
  Next
End Function

Function STCheckHit(aBall, target) 'Check if target is hit on it's face
  Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    STCheckHit = 1
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    STCheckHit = 1
  Else
    STCheckHit = 0
  End If
End Function

Sub DoSTAnim()
  Dim i
  For i = 0 To UBound(STArray)
    STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
  Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
  Dim animtime

  STAnimate = animate

  If animate = 0  Then
    primary.uservalue = 0
    STAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transy =  - STMaxOffset
    'vpmTimer.PulseSw switch mod 100
    If TableTilted=false then
      If CenterSpecialLight.state=1 then
        AddSpecial
        ResetDropsTimer.enabled=true
      else
        SetMotor(500)
      end if
    end if
    STAnimate = 2
    Exit Function
  ElseIf animate = 2 Then
    prim.transy = prim.transy + STAnimStep
    If prim.transy >= 0 Then
      prim.transy = 0
      primary.collidable = 1
      STAnimate = 0
      Exit Function
    Else
      STAnimate = 2
    End If
  End If
End Function

Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_ST11.transy
  For each BP in BP_ST11 : BP.transy = ty: Next

End Sub



'******************************************************
'***  END STAND-UP TARGETS
'******************************************************




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

'****************************************************************
'****  END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'****************************************************************

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
    RollingSoundFactor = 1.1 / 5

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
        PlaySoundAtLevelStatic SoundFX("em_knocker",DOFKnocker), KnockerSoundLevel, Plunger
    End Sub

    Sub KnockerClick()
        PlaySoundAtLevelStatic SoundFX("Click",DOFKnocker), KnockerSoundLevel, Plunger
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
    If TableTilted=false then
      PlaySound "buzz"
    End If
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

    Sub Apron1_Hit (idx)
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
        PlaySoundAtLevelStatic SoundFX("dropsup",DOFContactors), 1, obj
    End Sub

    Sub SoundDropTargetDrop(obj)
        PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
    End Sub

    '/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

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


'**********************************************
'*********************************
' ***************************************************************************
'          BASIC FSS(EM) 1-4 player 5x drums, 1 credit drum CORE CODE
' ****************************************************************************
' ********************* POSITION EM REEL DRUMS ON BACKGLASS *************************

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen
Dim inx
xoff =475
yoff = 0
zoff =735
xrot = -90

Const USEEMS = 1 ' 1-4 set between 1 to 4 based on number of players

const idx_emp1r1 =0 'player 1
const idx_emp4r6 =20 'credits


Dim BGObjEM(1)
if USEEMS = 1 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
end If

Sub center_objects_em()
Dim cnt,ii, xx, yy, yfact, xfact, objs
'exit sub
yoff = -150
zscale = 0.0000001
xcen =(960 /2) - (17 / 2)
ycen = (1065 /2 ) + (313 /2)

yfact = -25
xfact = -5

cnt =0
  For Each objs In BGObjEM(0)
  If Not IsEmpty(objs) then
    if objs.name = emp4r6.name then
    yoff = 50 ' credit drum is 60% smaller
    Else
    yoff = -110
    end if

  xx =objs.x

  objs.x = (xoff - xcen) + xx + xfact
  yy = objs.y
  objs.y =yoff

    If yy < 0 then
    yy = yy * -1
    end if

  objs.z = (zoff - ycen) + yy - (yy * zscale) + yfact

  'objs.rotx = xrot
  end if
  cnt = cnt + 1
  Next

end sub


' ********************* UPDATE EM REEL DRUMS CORE LIB *************************

Dim cred,ix, np,npp, reels(5, 7), scores(6,2)

'reset scores to defaults
for np =0 to 5
scores(np,0 ) = 0
scores(np,1 ) = 0
Next

'reset EM drums to defaults
For np =0 to 3
  For  npp =0 to 6
  reels(np, npp) =0 ' default to zero
  Next
Next

Sub SetScore(player, ndx , val)

  Dim ncnt

  if player = 5 or player = 6 then
    if val > 0 then
      If(ndx = 0)Then ncnt = val * 10
      If(ndx = 1)Then ncnt = val

      scores(player, 0) = scores(player, 0) + ncnt
    end if
  else
    if val > 0 then

    If(ndx = 0)then ncnt = val * 10000
    If(ndx = 1)then ncnt = val * 1000
    If(ndx = 2)Then ncnt = val * 100
    If(ndx = 3)Then ncnt = val * 10
    If(ndx = 4)Then ncnt = val

    scores(player, 0) = scores(player, 0) + ncnt
    'scores(player, 0) + ncnt

    end if
  end if
End Sub


Sub SetDrum(player, drum , val)
' debug.Print("setting drum " & drum & " to val " & val)
  Dim cnt
  Dim objs : objs =BGObjEM(0)

  If val = 0 then
    Select case player
    case -1: ' the credit drum
    If Not IsEmpty(objs(idx_emp4r6)) then
    objs(idx_emp4r6).rotX = 0 ' 285
    'cnt =objs(idx_emp4r6).ObjrotX
    end if
    Case 0:
      Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).rotX = 0: end if' 271
        Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).rotX = 0: end if
        Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).rotX = 0: end if
        Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).rotX = 0: end if
        Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).rotX = 0: end if
      End Select
  End Select

  else
  Select case player

    Case -1: ' the credit drum
      'emp4r6.ObjrotX = emp4r6.ObjrotX + val
      If Not IsEmpty(objs(idx_emp4r6)) then
      objs(idx_emp4r6).rotX = objs(idx_emp4r6).rotX + val
    end if

    Case 0:
    Select Case drum
      Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).rotX = objs(idx_emp1r1).rotX + val: end if
      Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).rotX= objs(idx_emp1r1+1).rotX + val: end if
      Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).rotX= objs(idx_emp1r1+2).rotX + val: end if
      Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).rotX= objs(idx_emp1r1+3).rotX + val: end if
      Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).rotX= objs(idx_emp1r1+4).rotX + val: end if
    End Select

  End Select
  end if
End Sub


Sub SetReel(player, drum, val)

  Dim  inc , cur, dif, fix, fval

  inc = 33.5
  fval = -5 ' graphic seam between 5 & 6 fix value, easier to fix here than photoshop

  If  (player <= 3) or (drum = -1) then

    If drum = -1 then drum = 0

    cur =reels(player, drum)

    If val <> cur then ' something has changed
    Select Case drum

      Case 0: ' credits drum

        if val > cur then
          dif =val - cur
          fix =0
            If cur < 5 and cur+dif > 5 then
            fix = fix- fval
            end if
          dif = dif * inc

          dif = dif-fix

          SetDrum -1,0,  -dif
        Else
          if val = 0 Then
          SetDrum -1,0,  0' reset the drum to abs. zero
          Else
          dif = 11 - cur
          dif = dif + val

          dif = dif * inc
          dif = dif-fval

          SetDrum -1,0,   -dif
          end if
        end if
      Case 1:
      'TB1.text = val
      if val > cur then
        dif =val - cur
        fix =0
          If cur < 5 and cur+dif > 5 then
          fix = fix- fval
          end if
        dif = dif * inc

        dif = dif-fix

        SetDrum player,drum,  -dif
      Else
        if val = 0 Then
        SetDrum player,drum,  0' reset the drum to abs. zero
        Else
        dif = 11 - cur
        dif = dif + val

        dif = dif * inc
        dif = dif-fval

        SetDrum player,drum,   -dif
        end if

      end if
      reels(player, drum) = val

      Case 2:
      'TB2.text = val

      if val > cur then
        dif =val - cur
        fix =0
          If cur < 5 and cur+dif > 5 then
          fix = fix- fval
          end if
        dif = dif * inc
        dif = dif-fix
        SetDrum player,drum,  -dif
      Else
        if val = 0 Then
        SetDrum player,drum,  0 ' reset the drum to abs. zero
        Else
        dif = 11 - cur
        dif = dif + val
        dif = dif * inc
        dif = dif-fval
        SetDrum player,drum,  -dif
        end if
      end if
      reels(player, drum) = val

      Case 3:
      'TB3.text = val

      if val > cur then
        dif =val - cur
        fix =0
          If cur < 5 and cur+dif > 5 then
          fix = fix- fval
          end if
        dif = dif * inc
        dif = dif-fix

        SetDrum player,drum,  -dif
      Else
        if val = 0 Then
        SetDrum player,drum,  0 ' reset the drum to abs. zero
        Else
        dif = 11 - cur
        dif = dif + val
        dif = dif * inc
        dif = dif-fval
        SetDrum player,drum,  -dif
        end if

      end if
      reels(player, drum) = val

      Case 4:
      'TB4.text = val

      if val > cur then
        dif =val - cur
        fix =0
          If cur < 5 and cur+dif > 5 then
          fix = fix- fval
          end if
        dif = dif * inc
        dif = dif-fix
        SetDrum player,drum,  -dif
      Else
        if val = 0 Then
        SetDrum player,drum,  0 ' reset the drum to abs. zero
        Else
        dif = 11 - cur
        dif = dif + val
        dif = dif * inc
        dif = dif-fval
        SetDrum player,drum,  -dif
        end if

      end if
      reels(player, drum) = val

      Case 5:
      'TB5.text = val

      if val > cur then
        dif =val - cur
        fix =0
          If cur < 5 and cur+dif > 5 then
          fix = fix- fval
          end if
        dif = dif * inc
        dif = dif-fix
        SetDrum player,drum,  -dif
      Else
        if val = 0 Then
        SetDrum player,drum,  0 ' reset the drum to abs. zero
        Else
        dif = 11 - cur
        dif = dif + val
        dif = dif * inc
        dif = dif-fval
        SetDrum player,drum,  -dif
        end if

      end if
      reels(player, drum) = val
     End Select

    end if
  end if
End Sub

Dim EMMODE: EMMODE = 0
Dim Score1000,Score10000,Score100000, ActivePLayer
Dim nplayer,playr,value,curscr,curplayr


Sub UpdateReels (Player,nReels ,nScore, n100K, Score10000 ,Score1000,Score100,Score10,Score1)

  ' to-do find out if player is one or zero based, if 1 based subtract 1.
  value =nScore'Score(Player)
  nplayer = Player -1

  curscr = value
  curplayr = nplayer


  scores(0,1) = scores(0,0)
  scores(0,0) = 0
  scores(1,1) = scores(1,0)
  scores(1,0) = 0
  scores(2,1) = scores(2,0)
  scores(2,0) = 0
  scores(3,1) = scores(3,0)
  scores(3,0) = 0

  For  ix =0 to 6
    reels(0, ix) =0
    reels(1, ix) =0
    reels(2, ix) =0
    reels(3, ix) =0
  Next

  For  ix =0 to 4
  SetDrum ix, 1 , 0
  SetDrum ix, 2 , 0
  SetDrum ix, 3 , 0
  SetDrum ix, 4 , 0
  SetDrum ix, 5 , 0
  Next

  For playr =0 to nReels

    if EMMODE = 0 then
    If (ActivePLayer) = playr Then
    nplayer = playr

    SetReel nplayer, 1 , Score10000 : SetScore nplayer,0,Score10000
    SetReel nplayer, 2 , Score1000 : SetScore nplayer,1,Score1000
    SetReel nplayer, 3 , Score100 : SetScore nplayer,2,Score100
    SetReel nplayer, 4 , Score10 : SetScore nplayer,3,Score10
    SetReel nplayer, 5 , 0 : SetScore nplayer,4,0 ' assumes ones position is always zero

    else
    nplayer = playr
    value =scores(nplayer, 1)


  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 1: end if

    end if
    Else
      If curplayr = playr Then
      nplayer = curplayr
      value = curscr
      else
      value =scores(playr, 1) ' store score
      nplayer = playr
      end if

    scores(playr, 0)  = 0 ' reset score
    if(value >= 100000) then
    value = value - 100000
    end if


  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 1: end if

    end if
  Next
End Sub


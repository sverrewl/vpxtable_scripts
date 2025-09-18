'********************************************************************************************
'  ____  _      ____   ____  _____    __  __          _____ _    _ _____ _   _ ______  _____
' |  _ \| |    / __ \ / __ \|  __ \  |  \/  |   /\   / ____| |  | |_   _| \ | |  ____|/ ____|
' | |_) | |   | |  | | |  | | |  | | | \  / |  /  \ | |    | |__| | | | |  \| | |__  | (___
' |  _ <| |   | |  | | |  | | |  | | | |\/| | / /\ \| |    |  __  | | | | . ` |  __|  \___ \
' | |_) | |___| |__| | |__| | |__| | | |  | |/ ____ \ |____| |  | |_| |_| |\  | |____ ____) |
' |____/|______\____/ \____/|_____/  |_|  |_/_/    \_\_____|_|  |_|_____|_| \_|______|_____/
'
'********************************************************************************************
'
'Blood Machines
'VPinWorkshop Original, 2022
'
'VPW Blood Machines
'==================
'Project Lead - iaakki
'Table Layout Design - iaakki, Dapheni
'Art Lead & 2D Graphics - Astronasty
'3D Objects & Textures - Tomate, Flupper, iaakki, Sixtoe, D.J.
'Scripting & Coding - iaakki, Apophis, oqqsan, Wylte, Lumigado, fluffhead35
'DMD - Oqsan, eMBee, Astronasty, Lumigado
'VR - Sixtoe, leojreimroc, iaakki
'Table Rules - Apophis, iaakki, VPW Team
'Shot Tester - gtxjoe
'Fixes & Additions - Entire VPW Team
'Testing - Rik, PinStratsDan, VPW Team
'
'A Film by Seth Ickerman
'Music by Carpenter Brut
'
' *IMPORTANT*
' In order to fully enjoy the table you will need to purchase the soundtrack MP3s by Carpenter Brut
' and place it in your vpinball/music directory. Usually: C:\VisualPinball\Music\Blood Machines OST/
' https://carpenterbrut.bandcamp.com/album/blood-machines-ost
'
' DMD will tell you, if this path is not found.
' You don't need to rename the songs or edit the script if you optain the songs from BandCamp page.
' If you are not hearing the opening tune properly in your setup, you probably need to raise the default MusicVol.
'
' Scorbit needs 3 exe files to be but into tables folder. sToken.exe and sQRCode.exe are the same as previous Scorbit tables.
' QRView.exe sources are at: https://github.com/iaakki/VPXQRView
' Remember to unblock these files
'
' This table requires FlexDMD v1.8.0 at least
' https://github.com/vbousquet/flexdmd/releases/
'
' This table requires VPX 10.7.3 at least
' https://github.com/vpinball/vpinball/releases/tag/10.7.3
'
' Watch the film on Shudder
' https://www.shudder.com/series/watch/blood-machines/033e95260abc2604?season=1
'

'1.22 various pup related edits by HeartbeatMD and Frank
'1.23 iaakki - Pupless Scorbit integration
'1.24 iaakki - merged some old stuff
'1.25 iaakki - EOBbonus counting can be skipped with flips, sling bottom corners removed from bouncer, PF mesh physics values cleaned
'        reverted the old EOBbonus speed
'1.26 iaakki - Cabinetmode scorbit flashers repositioned, Flips gets disabled when drained, game starting is not possible anymore until Scorbit has loaded properly
'1.27 iaakki - Fixed paths with spaces issue, renamed all pupchecks to HasPup variable, some minor improvements, possible fix for 4 player bug
'1.28 iaakki - HighScore input without pup crash fix, qrview binary moved to tables folder, csv log moved to BMQR folder, BMQR folder is now created if not existing.
'        QR image casting time raised from 500 to 700ms as some slower systems may fail to load the image that fast
'1.32 iaakki - all efx sounds uses calloutvolume now, SSF redone using Fleep sounds from TNA. Custom bumpersounds mixes from TNA and BPT samples, various fixes
'1.33 heartbeatDM - hasPuP checks to all related subs
'1.34 Frank  - MX Dof additions
'1.35 iaakki - Error messages shown if Scorbit binaries are not found in correct folders, flipkey and other flip ssf handling tweaks. EOB skip redone,
'1.36-37 - various pup related edits by HeartbeatMD and Frank
'1.38 iaakki - Scorbit player numbering fixed and tested to work with 2 players, Additional DOFs for flips removed, SlingSpin defaulted off, minor bugfix to wizard
'        SSF shaker added with DOF options, all rubber sounds redone
'1.40 iaakki - New baked metalwall
'1.41 Frank - MX Dof changes for  the MODES and MB
'1.42 iaakki -  UnderPF DTs redone with Rothbauer style DTs
'1.44 iaakki -  Hatch flasher tuned for VR. VR PUP DMD support created and reworked to play nice with flex still. It is all done a bit incorrectly
'       as we should be modifying VRBGSpeaker image only; not the backbox. This must be reworked one more time.
'1.45 iaakki - Half of the SSF sounds redone, underpf bug fixed
'1.46 iaakki - bug fixes, ball rebound swapped, some sounds removed
'1.47 iaakki - SSF sounds redone
'1.48 iaakki - Plunger material fix one more time, VR PUP DMD speaker grill redone, TargetCockpit sound change, underpf DT sounds fixed, song playback stutter
'        tested with delayed start. The problem comes from pupplayer side. If I disable audios from VLC player, the stutter disappears.
'1.49 iaakki - Wizard ending fixed, MimaTargetBig threshold increased slightly, ramp reflection fix, wizardphase4 ball rattle reduced
'1.50 iaakki - ssf and other sound balancing
'1.51 iaakki - Possible fix to game stuttering on song change when puppack is in use.
'1.53 iaakki - VR tested  on 10.8 #814. Room lighting baked again and added with GI colors and levels
'1.54 heartbeatDM - Attract and Game over attract Pup, Right flipper Skip Pup attract sequence, Lock Font Size
'RC1 iaakki - some fixes with positions. VRRoom platform added and floor lighting redone.
'RC2 leojreimroc - Fixes to VR Backglass and BG flashers 4-6 redone
'RC3 iaakki - Fix floor effects to not appear without vrroom, scorbit fixed one more time, qrview binary updated
'RC4 iaakki - Music player fixed, platform added to vr collection, floor height adjusted
'RC5 iaakki - Fixed EB handling in wizard, InsertBlooms balanced and added as option, slight tweak to vpx 10.8 hatch, wizard tester ON
'RC6 iaakki - Pup Score update sub redone, DTAnim arrays reworked, MusicDir fetch changed for 10.8.
'RC7 iaakki - pUpdateScores reworked one more time, fix shittalk, GunMotorSoundVolume bug fixed, forceFlexOn option added, optimizations for galdor battle, Apron buttonlight changed
'2.0 iaakki - default values

'SW versions to use:
' Desktop and cabinet:  64bit 10.7.3
' VR:           64bit 10.8 rev814
' FlexDMD 1.8.0


Option Explicit
Randomize

'**************************
'  FlexDmd or PupDmd
'**************************

Const HasPuP = false      'False=FlexDMD (default) , True=PUPDMD
Const forceFlexOn = false 'This is used for pup option 6 that enables Flex and Pup at the same time

'**************************
'   PinUp Player USER Config
'**************************

dim PuPDMDDriverType: PuPDMDDriverType=2   ' 0=LCD DMD, 1=RealDMD 2=FULLDMD (large/High LCD)
dim useRealDMDScale : useRealDMDScale=1    ' 0 or 1 for RealDMD scaling.  Choose which one you prefer.
dim useDMDVideos    : useDMDVideos=true   ' true or false to use DMD splash videos.


'*****************************************************************************************************************************************
'  Player Options
'*****************************************************************************************************************************************

'///////////////////////-----VR Room-----///////////////////////
Const VRRoom = 0          '0 - VR Room off, 1 - 360 Room, 2 - Minimal Room, 3 - Ultra Minimal

dim pGameName
if Not HasPuP then  'flex
  pGameName = "BloodMachines"  'pupvideos foldername, probably set to cGameName in realworld
else
  pGameName = "bloodmach"  'pupvideos foldername, probably set to cGameName in realworld
end if


'/////////////////////-----Scorbit Options-----////////////////////

          ScorbitActive       = 0   ' Is Scorbit Active
Const     ScorbitShowClaimQR  = 1   ' If Scorbit is active this will show a QR Code in the bottom left on ball 1 that allows player to claim the active player from the app
'Const     ScorbitClaimSmall  = 0   ' Make Claim QR Code smaller for high res backglass
Const     ScorbitUploadLog    = 1   ' Store local log and upload after the game is over
Const     ScorbitAlternateUUID  = 0   ' Force Alternate UUID from Windows Machine and saves it in VPX Users directory (C:\Visual Pinball\User\ScorbitUUID.dat)

'/////////////////////-----Other Options-----////////////////////


Const PlungerLaneStyle = 2      '1 - debris, 2 - grill
Const FlipperStyle = 1        '1 - purple logo, 2 - white, 3 - red logo
Const PinCabRailsColor = 1      '1 - Black, 2 - Red
Const DMDColor = 2          '1 - Orange, 2 - Red, 3 - White

Const VolumeDial = 0.8        'Recommended values should be no greater than 1.

Const MusicVol = 0.25       'Separate setting that only affects music volume. Range from 0 to 1.
'--> 'Default play mode plays 3 different really mellow tunes. Missions and MB's have more uplifting tunes.
Const StartupTune = true      'Play some opening tune after loading the game, I prefer you to disable this after you are sure you hear the music correctly
Const CalloutVol = 1        'Separate setting that only affects verbal callout volume. Range from 0 to 1

Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow, 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                  '2 = flasher image shadow, but it moves like ninuzzu's

Const FlasherBloomsEnabled = true 'true/false: Enabled and disables FlasherBlooms. Disable for performance
Const PupilLatency = True     'true/false: Enabled and disables Pupil Latency lighting feature.
Const bInsertBlooms = True      'true/false: Enabled and disables insert blooms lighting feature.
Dim Insert_Bloom_x_GIon : Insert_Bloom_x_GIon = 0.05  'Higher value, higher bloom

Const ballbrightnessMax = 170   'ball max brightness 0-255
Const ballbrightnessMin = 100   'ball min brightness 0-255

Const OptionWarmGi = False          'Set true for a bit more warm default GI
const bSlingSpin = false      'Sling corner spin feature. This is a random thing that happens with pinball.

Const ResetHighscores = False 'Set to True, start game , and close it down and put this back on False !!!

Const StagedFlipperMod = 0      '0 = not staged, 1 - staged (dual leaf switches)
Const KeyUpperLeft = 30       'A-key is the default bind for top flipper

'/////////////////////-----BallRoll Sound Amplification -----/////////////////////
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

Const AmbientAudioVolume = 0.00004  'Ambient volume. When no music available, the table will play ambient audio.

Const KeepLogs = false    'set True to save debug log file
Const KeepDOFLogs = False   'set True to save DOF logs. KeepLogs must also be True


'//  ShakerMotor:
'//  This setting will determine how will the shaker motor to be used during game play. Options:
'//  0 = Shaker motor completely disabled (DOF+Sound/SSF)
'//  1 = Shaker motor enabled for Sound/SSF only
'//  2 = Shaker motor enabled for DOF only
'//  3 = Shaker motor enabled for DOF + Sound/SSF
Const ShakerMotor = 3
'
'
'//  ShakerIntensity:
'//  This setting allows to set the overall intensity of the shaker motor (applicable for sound/SSF). Options:
'//  "Low"
'//  "Normal"
'//  "High
Const ShakerIntensity = "Low"

Const GunMotorSoundLevel = 0.7

dim ScorbitActive
'This should point to the path where all the original sound track music is placed
'dim MusicDirectory : MusicDirectory = "C:\Visual Pinball\Music\Blood Machines OST\"
dim MusicDirectory
'if VersionMinor => 8 Then
' MusicDirectory = musicdirectory("Blood Machines OST") 'this not working right now
'Else
  MusicDirectory = GetMusicFolder & "\Blood Machines OST\"
'end if

'The songs placed in the directory must bedefined here. The filenames must be correct and in this order
dim Songs(13)
Songs(1) = "Carpenter Brut - BLOOD MACHINES OST - 01 Intro.mp3"
Songs(2) = "Carpenter Brut - BLOOD MACHINES OST - 02 Blood Machines Theme.mp3"
Songs(3) = "Carpenter Brut - BLOOD MACHINES OST - 03 Attack Of The Amazons.mp3"
Songs(4) = "Carpenter Brut - BLOOD MACHINES OST - 04 The Ceremony.mp3"
Songs(5) = "Carpenter Brut - BLOOD MACHINES OST - 05 Mima.mp3"
Songs(6) = "Carpenter Brut - BLOOD MACHINES OST - 06 Souls Wreck.mp3"
Songs(7) = "Carpenter Brut - BLOOD MACHINES OST - 07 Touchdown.mp3"
Songs(8) = "Carpenter Brut - BLOOD MACHINES OST - 08 Heart Ship.mp3"
Songs(9) = "Carpenter Brut - BLOOD MACHINES OST - 09 The Last Ceremony.mp3"
Songs(10) = "Carpenter Brut - BLOOD MACHINES OST - 10 Bloody Kisses - The Swift.mp3"
Songs(11) = "Carpenter Brut - BLOOD MACHINES OST - 11 Lago's Sleep.mp3"
Songs(12) = "Carpenter Brut - BLOOD MACHINES OST - 12 Grand Final.mp3"
Songs(13) = "Carpenter Brut - BLOOD MACHINES OST - 13 Gone Now (Feat. Pencey Sloe).mp3"


Function GetMusicFolder()
    Dim GMF
    Set GMF = CreateObject("Scripting.FileSystemObject")
    GetMusicFolder= GMF.GetParentFolderName(userdirectory) & "\Music"
    set GMF = nothing
End Function



'*****************************************************************************************************************************************
' Script Options
'*****************************************************************************************************************************************


Dim CabinetMode, DesktopMode: DesktopMode = Table1.ShowDT
If Not DesktopMode and VRRoom=0 Then CabinetMode=1 Else CabinetMode=0


if PlungerLaneStyle = 1 Then
  pPlungerlaneMetal1.visible = true
  pPlungerlaneMetalRefl1.visible = true
  pPlungerlaneMetal2.visible = false
  pPlungerlaneMetalRefl2.visible = false
  plFlash1.TransmissionScale = 0.2
  plFlash2.TransmissionScale = 0.2
  plFlash3.TransmissionScale = 0.2
Else
  pPlungerlaneMetal1.visible = false
  pPlungerlaneMetalRefl1.visible = False
  pPlungerlaneMetal2.visible = true
  pPlungerlaneMetalRefl2.visible = true
  plFlash1.TransmissionScale = 0
  plFlash2.TransmissionScale = 0
  plFlash3.TransmissionScale = 0
end if


lflip.image = "flippers0" & FlipperStyle
LFlip1.image = "flippers0" & FlipperStyle
rflip.image = "flippers0" & FlipperStyle

if PinCabRailsColor = 1 Then
  Pincab_Rails.material = "Metal_Black_Powdercoat"
Else
  Pincab_Rails.material = "Metal_Red_Powdercoat"
End if




'*****************************************************************************************************************************************
' DOF IDs
'*****************************************************************************************************************************************

'E101 0/1 LeftFlipper
'E102 0/1 RightFlipper
'E103 2 Leftslingshot
'E104 2 Rightslingshot
'E105 2 top bumper left
'E106 2 top bumper center / EB Slot
'E107 2 top bumper right
'E108 2 cockpit target
'E109 2 mima target
'E110 2 scavenge targets
'E111 2 kickback
'E112 2 vascan targets
'E113 2 VUK1
'E114 2 VUK2, scavenge kicker -> right bumber or middle right
'E115 2 Flasher 1
'E116 2 Flasher 2
'E117 2 Flasher 3
'E118 2 Flasher 4
'E119 2 Flasher 5
'E120 2 Flasher 6
'E121 2 Flasher 7
'E122 0/1 Start game button ready
'E123 2 start game
'E124 0/1 Ball in plunger lane
'E125 2 Ball exits plunger lane
'E126 2 drain hit (ball not saved)
'E127 2 drain hit (ball saved)
'E128 2 spinner
'E128 2 spinner
'E129 2 skill shot 1
'E130 2 skill shot 2
'E131 2 skill shot 3
'E132 2 award extra ball
'E133 2 add credits
'E134 2 lite kickback
'E135 2 Flasher Undership
'E136 2 left orbit
'E137 2 right orbit
'E138 2 main ramp
'E138 2 sky ramp
'E140 2 jackpot
'E141 2 super jackpot
'E142 2 MB wizard mode jackpot
'E143 2 mission wizard mode completed
'E144 2 fire lazer cannon
'E145 2 ball released / autoplunger
'E146 0/1 red undercab
'E147 0/1 orange undercab
'E148 0/1 green undercab
'E149 0/1 blue undercab
'E150 0/1 purple undercab
'E151 2 magnet pulling ball
'E152 0/1 shaker motor (5 sec max time limit)
'E153 0/1 fire button
'E154 knocker ' ssftodo: add to configtool

'*****************************************************************************************************************************************
'  SCORE VALUES
'*****************************************************************************************************************************************

'AddScore() values

Const score_SlingShot = 1003          'Slingshot hit
Const score_Bumper = 10000            'Bumper hit
Const score_SuperJetBumper = 25000        'Super Jet Bumper hit
Const score_Inlanes = 1000            'Inlane hits
Const score_Outlanes = 25000          'Outlane hits
Const score_LagoCompleted = 250000        'LAGO lanes completed
Const score_KBMaxed = 1000000         'Kickback maxed out
Const score_spinnerspin = 10000         'spinner1 spin, multiplied by 5 when super spinner is active
Const score_swRightOrbTrigger1_hit = 2500   'hight right orbit trigger
Const score_TopLanes = 5000           'Top rollover lane hit
Const score_IncreasedEOBmulti = 7500      'Increased end of ball multiplier
Const score_MaxedEOBmulti = 2500000       'Exceeded maximum EOB multiplier
Const score_RampEntrance = 2500         'Ramp entrance
Const score_RampShot = 250000         'Successful ramp shot
Const score_SkyLoop = 500000          'Sky loop
Const score_GunTarget = 500000          'Gun Target
Const score_GunTarget_increment = 250000    'Gun Target increment per shot
Const score_ShipHit = 25000           'Tracy ship hit (vortex closed)
Const score_LoopEntrance = 2500         'Loop entrance switches
Const score_LoopComplete = 100000       'Loop shot completed
Const score_VUKSpinner = 250000         'Spinner (right) VUK hit
Const score_VUKGun = 250000           'Gun (left) VUK hit
Const score_DropsHit = 100000         'Drops hit
Const score_DropsComplete = 500000        'Drops completed
Const score_ScavengeScoop = 100000        'Right scoop hit
Const score_HarpoonHit = 500000         'Harpoon kicker hit
Const score_RingLoop = 1000           'Ring loop hit
Const score_ExtraBall = 7500000         'Extra ball points, awarded if disabled or max EB reached

Const score_MimaTargetBig_Hit = 50000     'Hit the Mima target
Const score_MimaTargetBig_Maxed = 500000    'Exceeded 3 Mima target hits
Const score_MimaLock = 2500000          'locked a ball for Mima MB
Const score_StartMimaMB = 1000000       'start Mima MB
Const score_MimaMBJackpot = 500000        'Mima MB Jackpot
Const score_MimaMBSuperJackpot = 3500000    'Mima MB Super Jackpot
Const score_StartCoreyMB = 1000000        'start Corey MB
Const score_CoreyMBJackpot = 500000       'Corey MB Jackpot
Const score_CoreyMBSuperJackpot = 1000000   'Corey MB Super Jackpot
Const score_StartTracyMB = 1000000        'start Tracy MB
Const score_TracyMBJackpot = 1000000      'Tracy MB Jackpot
Const score_TracyMBSuperJackpot = 5000000   'Tracy MB Super Jackpot


Const score_WizardPhase2OrbShot = 100000    'wizard orb shot points
Const score_WizardPhase2Passed  = 20000000    'orb shot stage passed
Const score_WizardPhase3Passed  = 30000000    'under pf stage passed
Const score_WizardOrbShot     = 100000    'Wizard Phase2 orb shot
Const score_WizardGaldorShot  = 100000    'Wizard Galdor shot * souls on table
Const score_WizardCompleted   = 2000000   'Wizard Galdor defeat * souls on table

Const score_Skillshot1_1st = 500000       'Skillshot 1, first hit
Const score_Skillshot1_2nd = 600000       'Skillshot 1, second hit
Const score_Skillshot1_3rd = 700000       'Skillshot 1, third hit
Const score_Skillshot1_4th = 800000       'Skillshot 1, fourth hit
Const score_Skillshot2_1st = 2000000      'Skillshot 2, first hit
Const score_Skillshot2_2nd = 2200000      'Skillshot 2, second hit
Const score_Skillshot2_3rd = 2400000      'Skillshot 2, third hit
Const score_Skillshot2_4th = 2600000      'Skillshot 2, fourth hit
Const score_Skillshot3_1st = 4000000      'Skillshot 3, first hit
Const score_Skillshot3_2nd = 4400000      'Skillshot 3, second hit
Const score_Skillshot3_3rd = 4800000      'Skillshot 3, third hit
Const score_Skillshot3_4th = 5200000      'Skillshot 3, fourth hit

'AddMissionScore() values. These can be reawarded in the Mission Wizard Mode

Const score_Mission1_ShotHit  = 1000000     'Mission 1, shot hit
Const score_Mission1_AllShots = 3000000     'Mission 1, all shots done
' 60M

Const score_Mission2_10Loops = 3000000      'Mission 2, 25 loops made
Const score_Mission2_EachLoop = 100000      'Mission 2, each loop scores this
' 55M

Const score_Mission3_1Shot  =  500000     'Mission 3, 1 successful shot
Const score_Mission3_2Shots = 1000000     'Mission 3, 2 successful shots
Const score_Mission3_3Shots = 1500000     'Mission 3, 3 successful shots
Const score_Mission3_4Shots = 2500000     'Mission 3, 3 successful shots
' 55M

Const score_Mission4_SkyShot  = 3000000     'Mission 4, sky shot successful
Const score_Mission4_RampShot = 1500000     'Mission 4, ramp shot successful
' 45 - 75m

Const score_Mission5_VascanShot = 1000000   'Mission 5, vascan shot successful
' 60m

Const score_Mission6_Bumper = 100000      'Mission 6, bumper hit
Const score_Mission6_25Bumpers  = 3000000   'Mission 6, got to 25 bumper hits
' 55m

Const score_Mission7_1stShot = 500000     'Mission 7, first shot in combo hit
Const score_Mission7_2ndShot = 2000000      'Mission 7, combo hit
Const score_Mission7_2combos = 3000000      'Mission 7, two combos hit
'60+ ( can do shot 1 many times)

'AddBonus() values. These are multiplied by the Bonus Multiplier and added to the score at end of ball

Const bonus_EOB_Ramps = 10000   ' *EOB*
Const bonus_EOB_Missions = 1000000
Const bonus_EOB_MimaLoops = 2500

Const bonus_SuperSpinners = 0         'Super spinner bonus
Const bonus_MimaMBSuperJackpot = 0        'Mima MB Super Jackpot
Const bonus_CoreyMBSuperJackpot = 0       'Corey MB Super Jackpot
Const bonus_TracyMBSuperJackpot = 0       'Tracy MB Super Jackpot
Const bonus_MBWizardJackpot = 0         'MB Wizard MB Jackpot
Const bonus_MissionWizardJackpot = 0      'Mission Wizard MB Jackpot

Const bonus_Mission1_Completed = 0        'Mission 1 completed
Const bonus_Mission2_Completed = 0        'Mission 2 completed
Const bonus_Mission3_Completed = 0        'Mission 3 completed
Const bonus_Mission4_SkyShot = 0        'Mission 4 sky shot successful
Const bonus_Mission4_Completed = 0        'Mission 4 completed
Const bonus_Mission5_Completed = 0        'Mission 5 completed
Const bonus_Mission6_Completed = 0        'Mission 6 completed
Const bonus_Mission7_Completed = 0        'Mission 7 completed

Const bonus_Skillshot1_1st = 0          'Skillshot 1, first hit
Const bonus_Skillshot1_2nd = 0          'Skillshot 1, second hit
Const bonus_Skillshot1_3rd = 0          'Skillshot 1, third hit
Const bonus_Skillshot1_4th = 0          'Skillshot 1, fourth hit
Const bonus_Skillshot2_1st = 0          'Skillshot 2, first hit
Const bonus_Skillshot2_2nd = 0          'Skillshot 2, second hit
Const bonus_Skillshot2_3rd = 0          'Skillshot 2, third hit
Const bonus_Skillshot2_4th = 0          'Skillshot 2, fourth hit
Const bonus_Skillshot3_1st = 0          'Skillshot 3, first hit
Const bonus_Skillshot3_2nd = 0          'Skillshot 3, second hit
Const bonus_Skillshot3_3rd = 0          'Skillshot 3, third hit
Const bonus_Skillshot3_4th = 0          'Skillshot 3, fourth hit

'*****************************************************************************************************************************************
'  VARIABLES
'*****************************************************************************************************************************************


ballrolleron = 1 ' set to 0 to turn off the ball roller if you use the "c" key in your cabinet

Const cGameName = "bloodmach"
Const TableName = "BloodMachines"
Const myVersion = "1.0.0"

'Constants
Const BallSize = 50
Const BallMass = 1
Const MaxPlayers = 4
Const BallSaverTime = 15
Const MaxMultiballs = 8
Const MaxExtraBallsPerGame = 2
Const bpgcurrent = 3
Const ReflipAngle = 20

dim ballbrightness : ballbrightness = -1


' Define Global Variables
Dim MultiBallActive
Dim SuperJets

Dim PlayerLights(4,30)


Dim ballrolleron
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits

' add some (4) for mystery/ss reward  free locked mima ball
Dim MysteryEB(4)
Dim CompletedSkillshots(4)
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim TotalExtraBallsAwarded(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim GunTargetScore(4)
Dim BumperHits(4)
Dim BumperHitsPup(4)
Dim BumperThousandPup(4)
Dim BumperTwentyThousandPup(4)
Dim BumperscoringPup(4)
Dim Mission1Points(4)
Dim Mission2Points(4)
Dim Mission3Points(4)
Dim Mission4Points(4)
Dim Mission5Points(4)
Dim Mission6Points(4)
Dim Mission7Points(4)
Dim Jackpot
Dim SuperJackpot
Dim Tilt
Dim TiltSensitivity
Dim Tilted
dim Slammed
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim bAutoPlunger
Dim bromconfig
Dim bAttractMode
Dim bEndCredits
Dim BallsOnPlayfield
Dim BallsInHole
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bMultiBallMode
Dim PupBumperTracy
Dim PupBumpertenthousand
Dim PupBumpersuperjets
Dim bMultiballReady
Dim bCoreyMBOngoing
Dim bTracyMBOngoing
Dim bMissionMode
Dim bWizardMode
dim NumLocksMima
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bExtraBallWonThisBall
Dim bMusicOn
Dim bJustStarted
Dim bFlippersEnabled
Dim bPortalFlippersEnabled
Dim PointMultiplier : PointMultiplier=1
Dim SpinnerMultiplier
Dim AddBallAvailable
dim MissionStart1LightState: MissionStart1LightState = 0
dim LightScavengeScoopState: LightScavengeScoopState = 0
dim CounterBallSearch : CounterBallSearch = 0


GunTargetScore(1) = score_GunTarget
GunTargetScore(2) = score_GunTarget
GunTargetScore(3) = score_GunTarget
GunTargetScore(4) = score_GunTarget


Dim plungerIM 'used mostly as an autofire plunger
Dim plungerKB 'kickback plunger
Dim VorMag, MiMag, GrabMag

LoadCoreFiles
Sub LoadCoreFiles
  On Error Resume Next
  ExecuteGlobal GetTextFile("core.vbs")
  If Err Then MsgBox "Can't open core.vbs"
  On Error Goto 0
End Sub
'
'Dim EnableBallControl
'EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys
'

'*****************************************************************************************************************************************
'  GI colorchange
'*****************************************************************************************************************************************
Dim GI_NewColor

Const GI_Brightness_multiplier = 1.5
Const GI_Transmit_multiplier = 0.5

Sub GI_Color ( aCol,aIntesity )
  If aCol="white" then Gioff
  GIcolortimer.enabled=True
  GI_NewColor=aCol
  GI_Brightness(0)=aIntesity * GI_Brightness_multiplier
  GI_Transmit(0)= GI_Transmit_multiplier
End Sub

dim RampGiColor

Sub Gicolortimer_Timer
  GIcolortimer.enabled=False
  Select Case GI_NewColor
    Case "white"
            if OptionWarmGi then
              SetGiColor rgb(255,200,160) , rgb(255,197,143)
            else
                SetGiColor rgb(250,220,200) , rgb(255,210,163)
            end if
      FlReinitBumper 1, "white"
      FlReinitBumper 2, "white"
      FlReinitBumper 3, "white"
      'ShipGiColor = RGB(242,123,36)
      ShipGiColor = RGB(202,123,76)
      RampGiColor = RGB(255,255,255)
      FlasherPL_BS.color = rgb(255,150,68)
      if cabinetmode = 0 then
        FlasherPL_LS.color = rgb(255,150,68)
        FlasherPL_RS.color = rgb(255,150,68)
      Else
        FlasherPL_LS_CM.color = rgb(255,150,68)
        FlasherPL_RS_CM.color = rgb(255,150,68)
      end if
      FlasherGI.color = rgb(255,150,68)
      Reflect_pf_left.color = rgb(255,255,255)
      Reflect_pf_right.color = rgb(255,255,255)
      Reflect2b.color = rgb(255,255,255)
      Reflect2b001.color = rgb(255,255,255)
      gion                    'todo: is this really needed here
    Case "green"
      SetGiColor rgb(0,255,0) , rgb(150,255,150)
        FlReinitBumper 1, "green"
      FlReinitBumper 2, "green"
      FlReinitBumper 3, "green"
      ShipGiColor = RGB(0,242,0)
      RampGiColor = RGB(200,255,200)
      FlasherPL_BS.color = rgb(64,222,24)
      if cabinetmode = 0 then
        FlasherPL_LS.color = rgb(64,222,24)
        FlasherPL_RS.color = rgb(64,222,24)
      Else
        FlasherPL_LS_CM.color = rgb(64,222,24)
        FlasherPL_RS_CM.color = rgb(64,222,24)
      end if
      FlasherGI.color = rgb(64,222,24)
      Reflect_pf_left.color = rgb(127,255,96)
      Reflect_pf_right.color = rgb(127,255,96)
      Reflect2b.color = rgb(127,255,96)
      Reflect2b001.color = rgb(127,255,96)
    Case "red"
      SetGiColor rgb(255,0,0) , rgb(255,150,150)
        FlReinitBumper 1, "red"
      FlReinitBumper 2, "red"
      FlReinitBumper 3, "red"
      ShipGiColor = RGB(242,0,0)
      RampGiColor = RGB(255,200,200)
      FlasherPL_BS.color = rgb(228,28,24)
      if cabinetmode = 0 then
        FlasherPL_LS.color = rgb(228,28,24)
        FlasherPL_RS.color = rgb(228,28,24)
      else
        FlasherPL_LS_CM.color = rgb(228,28,24)
        FlasherPL_RS_CM.color = rgb(228,28,24)
      end if
      FlasherGI.color = rgb(228,28,24)
      Reflect_pf_left.color = rgb(255,96,96)
      Reflect_pf_right.color = rgb(255,96,96)
      Reflect2b.color = rgb(255,96,96)
      Reflect2b001.color = rgb(255,96,96)
    Case "blue"
      SetGiColor rgb(0,0,255) , rgb(150,150,255)
        FlReinitBumper 1, "blue"
      FlReinitBumper 2, "blue"
      FlReinitBumper 3, "blue"
      ShipGiColor = RGB(0,0,242)
      RampGiColor = RGB(200,200,255)
      FlasherPL_BS.color = rgb(28,28,254)
      if cabinetmode = 0 then
        FlasherPL_LS.color = rgb(28,28,254)
        FlasherPL_RS.color = rgb(28,28,254)
      else
        FlasherPL_LS_CM.color = rgb(28,28,254)
        FlasherPL_RS_CM.color = rgb(28,28,254)
      end if
      FlasherGI.color = rgb(28,28,254)
      Reflect_pf_left.color = rgb(127,127,255)
      Reflect_pf_right.color = rgb(127,127,255)
      Reflect2b.color = rgb(127,127,255)
      Reflect2b001.color = rgb(127,127,255)
    Case "purple"
      SetGiColor rgb(200,1,250) , rgb(255,150,245)
        FlReinitBumper 1, "purple"
      FlReinitBumper 2, "purple"
      FlReinitBumper 3, "purple"
      ShipGiColor = RGB(200,0,242)
      RampGiColor = RGB(255,200,255)
      FlasherPL_BS.color = rgb(180,44,200)
      if cabinetmode = 0 then
        FlasherPL_LS.color = rgb(180,44,200)
        FlasherPL_RS.color = rgb(180,44,200)
      else
        FlasherPL_LS_CM.color = rgb(180,44,200)
        FlasherPL_RS_CM.color = rgb(180,44,200)
      end if
      FlasherGI.color = rgb(180,44,200)
      Reflect_pf_left.color = rgb(222,96,232)
      Reflect_pf_right.color = rgb(222,96,232)
      Reflect2b.color = rgb(222,96,232)
      Reflect2b001.color = rgb(222,96,232)
  End Select
  UpdateMaterial "Ship_gi_color",0,0,0,0,0,0,ShipGiIntensityMultiplier * gilvl,ShipGiColor,0,0,False,True,0,0,0,0
'UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity, OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive, float elasticity, float elasticityFalloff, float friction, float scatterAngle)
  UpdateMaterial "plasticRamps",0.1,0.8,0.7,1,1,0,0.999999,RampGiColor,rgb(250,200,180),rgb(250,200,180),False,True,0,0,0,0

End Sub

Sub SetGiColor ( newcol,newcolfull )
  Dim xx,i
  i=0
  For Each xx in GI
    i=i+1
    xx.intensity = GI_Brightness(i) * GI_Brightness(0)
    xx.TransmissionScale = GI_Transmit(i) * GI_Transmit(0)
    xx.color = newcol
    xx.colorfull = newcolfull
  Next
  pTunnel_light.color = newcol
End Sub

Dim GI_Brightness(100)
Dim GI_Transmit(100)

Sub GI_intensity
  Dim xx,i
  i=0
  For Each xx in GI
    i=i+1
    GI_Brightness(i)=xx.intensity
    GI_Transmit(i)=xx.TransmissionScale
  Next

End Sub


'E146 0/1 red undercab
'E147 0/1 orange undercab
'E148 0/1 green undercab
'E149 0/1 blue undercab
'E150 0/1 purple undercab
Sub UndercabGIOff
  DOF 146,0
  DOF 147,0
  DOF 148,0
  DOF 149,0
  DOF 150,0
End Sub

Sub UndercabGIOn
  Select Case GI_NewColor
    Case "white"
      DOF 146,1
    Case "orange"
      DOF 147,1
    Case "green"
      DOF 148,1
    Case "red"
      DOF 146,1
    Case "blue"
      DOF 149,1
    Case "purple"
      DOF 150,1
  End Select
End Sub


'*****************************************************************************************************************************************
'  Insert Blooms
'*****************************************************************************************************************************************
Dim Insert_Blooms(100)


Sub Init_Insert_Blooms
  dim x, i
  i=0

  for Each x in InsertBlooms
    i=i+1
    Insert_Blooms(i)=x.intensity
  next
end sub

'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "


Class DebugLogFile

    Private Filename
    Private TxtFileStream

    Private Function LZ(ByVal Number, ByVal Places)
        Dim Zeros
        Zeros = String(CInt(Places), "0")
        LZ = Right(Zeros & CStr(Number), Places)
    End Function

    Private Function GetTimeStamp
        Dim CurrTime, Elapsed, MilliSecs
        CurrTime = Now()
        Elapsed = Timer()
        MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
        GetTimeStamp = _
            LZ(Year(CurrTime),   4) & "-" _
            & LZ(Month(CurrTime),  2) & "-" _
            & LZ(Day(CurrTime),    2) & " " _
            & LZ(Hour(CurrTime),   2) & ":" _
            & LZ(Minute(CurrTime), 2) & ":" _
            & LZ(Second(CurrTime), 2) & ":" _
            & LZ(MilliSecs, 4)
    End Function

' *** Debug.Print the time with milliseconds, and a message of your choice
    Public Sub WriteToLog(label, message, code)
        Dim FormattedMsg, Timestamp
        'Filename = UserDirectory + "\" + cGameName + "_debug_log.txt"
    Filename = cGameName + "_debug_log.txt"

        Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, code, True)
        Timestamp  = GetTimeStamp
        FormattedMsg = GetTimeStamp + " : " + label + " : " + message
        TxtFileStream.WriteLine FormattedMsg
        TxtFileStream.Close
    debug.print label & " : " & message
  End Sub

End Class

Sub WriteToLog(label, message)
  if KeepLogs Then
    Dim LogFileObj
    Set LogFileObj = New DebugLogFile
    LogFileObj.WriteToLog label, message, 8
  end if
End Sub

Sub NewLog()
  if KeepLogs Then
    Dim LogFileObj
    Set LogFileObj = New DebugLogFile
    LogFileObj.WriteToLog "NEW LOG", " ", 2
  end if
End Sub


'*****************************************************************************************************************************************
'  PLAYER VARIABLES
'*****************************************************************************************************************************************

Sub PlayerLights_SaveLights
  WriteToLog "PlayerLights_SaveLights", ""
  PlayerLights(CurrentPlayer,0)  = 0 ' mystery off
  PlayerLights(CurrentPlayer,1)  = NumLocksMima
  PlayerLights(CurrentPlayer,2)  = MimaCount

  PlayerLights(CurrentPlayer,3)  = cLightCorey1.State
  PlayerLights(CurrentPlayer,4)  = cLightCorey2.State
  PlayerLights(CurrentPlayer,5)  = cLightCorey3.State
  PlayerLights(CurrentPlayer,6)  = cLightCorey4.State
  PlayerLights(CurrentPlayer,7)  = cLightCorey5.State

  PlayerLights(CurrentPlayer,8)  = cLightTracy1.State
  PlayerLights(CurrentPlayer,9)  = cLightTracy2.State
  PlayerLights(CurrentPlayer,10) = cLightTracy3.State
  PlayerLights(CurrentPlayer,11) = cLightTracy4.State
  PlayerLights(CurrentPlayer,12) = cLightTracy5.State

  PlayerLights(CurrentPlayer,13) = cLightMission1.State
  PlayerLights(CurrentPlayer,14) = cLightMission2.State
  PlayerLights(CurrentPlayer,15) = cLightMission3.State
  PlayerLights(CurrentPlayer,16) = cLightMission4.State
  PlayerLights(CurrentPlayer,17) = cLightMission5.State
  PlayerLights(CurrentPlayer,18) = cLightMission6.State
  PlayerLights(CurrentPlayer,19) = cLightMission7.State


  PlayerLights(CurrentPlayer,20) = cLightMimaMB.state
  PlayerLights(CurrentPlayer,21) = cLightCoreyMB.state
  PlayerLights(CurrentPlayer,22) = cLightTracyMB.state

  PlayerLights(CurrentPlayer,23) = MissionSelect
  PlayerLights(CurrentPlayer,24) = cLightCollectExtraBall.state

  PlayerLights(CurrentPlayer,25) = cLightMystery1.state
  PlayerLights(CurrentPlayer,26) = cLightMystery2.state
  PlayerLights(CurrentPlayer,27) = cLightMystery3.state

End Sub



Sub PlayerLights_ResetAll
  WriteToLog "PlayerLights_ResetAll", ""
  ' for new gamestarted
  dim xx,i
  for xx = 1 to 4

    for i = 0 to 30
      PlayerLights(xx,i) = 0
    next

    PlayerLights(xx,2) = 2          '  start with Mimacount = 2

    PlayerLights(xx,6) = 2          ' Corey light on Ramp start blinking for the first wizard
    PlayerLights(xx,int(rnd(1)*5)+8) = 2  ' Random Tracy light on Ramp start blinking for the first wizard

    PlayerLights(xx,23) = int(rnd(1)*7)+1   ' mission select start at random Light

    DTRaise 1
    DTRaise 2
    DTRaise 3
    RandomSoundDropTargetLeft(TargetScavenge2p)

  next
End Sub




Sub PlayerLights_NewPlayer
  WriteToLog "PlayerLights_NewPlayer", ""
  cLightScavengeScoop.state=PlayerLights(CurrentPlayer,0)
  NumLocksMima = PlayerLights(CurrentPlayer,1)
  If NumLocksMima>0 Then
    cMimalock1.state=1
  End If
  If NumLocksMima>1 Then
    cMimalock2.state=1
  End If
  MimaCount = PlayerLights(CurrentPlayer,2)

  cLightCorey1.State = PlayerLights(CurrentPlayer,3)
  cLightCorey2.State = PlayerLights(CurrentPlayer,4)
  cLightCorey3.State = PlayerLights(CurrentPlayer,5)
  cLightCorey4.State = PlayerLights(CurrentPlayer,6)
  cLightCorey5.State = PlayerLights(CurrentPlayer,7)

  cLightTracy1.State = PlayerLights(CurrentPlayer,8)
  cLightTracy2.State = PlayerLights(CurrentPlayer,9)
  cLightTracy3.State = PlayerLights(CurrentPlayer,10)
  cLightTracy4.State = PlayerLights(CurrentPlayer,11)
  cLightTracy5.State = PlayerLights(CurrentPlayer,12)

  If PlayerLights(CurrentPlayer,13)=1 Then cLightMission1.State=1
  If PlayerLights(CurrentPlayer,14)=1 Then cLightMission2.State=1
  If PlayerLights(CurrentPlayer,15)=1 Then cLightMission3.State=1
  If PlayerLights(CurrentPlayer,16)=1 Then cLightMission4.State=1
  If PlayerLights(CurrentPlayer,17)=1 Then cLightMission5.State=1
  If PlayerLights(CurrentPlayer,18)=1 Then cLightMission6.State=1
  If PlayerLights(CurrentPlayer,19)=1 Then cLightMission7.State=1

  cLightMimaMB.state = PlayerLights(CurrentPlayer,20)
  cLightCoreyMB.state = PlayerLights(CurrentPlayer,21)
  cLightTracyMB.state = PlayerLights(CurrentPlayer,22)

  bMultiballReady = false
  CheckMBReady

  bStartShitTalk = false : cShitTalkCounter = 0
  checkWizardStart

  If MimaCount > 0 Then cTargetMima1.state = 1 ' PTargetMima1.blenddisablelighting=12
  If MimaCount > 1 Then cTargetMima2.state = 1 ' PTargetMima2.blenddisablelighting=12
  If MimaCount > 2 Then
    cTargetMima3.state = 1 ' PTargetMima3.blenddisablelighting=12
    If PlayerLights(CurrentPlayer,20) <> 1 and Not bMultiballReady Then 'MimaMB not completed
      Rise_RampMima.enabled=True
      Rise_RampMima_Timer ' just to be sure, the timer reset the other one
    End If
  End If

  MissionSelect = PlayerLights(CurrentPlayer,23)
  If MissionSelect=1 Then : If cLightMission1.State=1 Then MissionSelect=8 Else cLightMission1.State=2 : End If
  If MissionSelect=2 Then : If cLightMission2.State=1 Then MissionSelect=8 Else cLightMission2.State=2 : End If
  If MissionSelect=3 Then : If cLightMission3.State=1 Then MissionSelect=8 Else cLightMission3.State=2 : End If
  If MissionSelect=4 Then : If cLightMission4.State=1 Then MissionSelect=8 Else cLightMission4.State=2 : End If
  If MissionSelect=5 Then : If cLightMission5.State=1 Then MissionSelect=8 Else cLightMission5.State=2 : End If
  If MissionSelect=6 Then : If cLightMission6.State=1 Then MissionSelect=8 Else cLightMission6.State=2 : End If
  If MissionSelect=7 Then : If cLightMission7.State=1 Then MissionSelect=8 Else cLightMission7.State=2 : End If
  If MissionSelect=8 Then Next_MissionSelect

  If MissionSelect=0 Then
    ' all done ... must finish wizard to reset It
  Else
    cLightMissionStart1.state=2
    cLightShot1.state = 0
    cLightShot5.state = 0
  End If


  cLightCollectExtraBall.state = PlayerLights(CurrentPlayer,24)

  if PlayerLights(CurrentPlayer,25) = 1 then cLightMystery1.state = 1 else cLightMystery1.state = 0
  if PlayerLights(CurrentPlayer,26) = 1 then cLightMystery2.state = 1 else cLightMystery2.state = 0
  if PlayerLights(CurrentPlayer,27) = 1 then cLightMystery3.state = 1 else cLightMystery3.state = 0

  WriteToLog "PlayerLights_NewPlayer","MimaCount=" & MimaCount
End Sub

Sub PlayerLights_ResetInserts ' turn them all off
  WriteToLog "PlayerLights_ResetInserts", ""
  Dim xx
  For Each xx in AllLights
    xx.state=0
  Next
  ResetMimaTargets ' and lowers ramp here
  cMimalock1.state=0
  cMimalock2.state=0
  cTargetMima1.state = 0
  cTargetMima2.state = 0
  cTargetMima3.state = 0

  DTRaise 1
  DTRaise 2
  DTRaise 3
  RandomSoundDropTargetLeft(TargetScavenge2p)

End Sub

dim OrbUnlockShot:OrbUnlockShot=0
sub OrbUnlockMission
  if cLightMissionStart1.state=2 then
    cLightShot1.state = 0
    cLightShot5.state = 0
    OrbUnlockShot = 0
    exit Sub
  end if

  'all missions played
  if cLightMission1.State = 1 and cLightMission2.State = 1 and cLightMission3.State = 1 and cLightMission4.State = 1 and cLightMission5.State = 1 and cLightMission6.State = 1 and cLightMission7.State = 1 Then
    exit Sub
  end if

  if CurrentMission = 0 And Not bMultiBallMode And Not bWizardMode then
    if RndInt(0,1) = 0 Then
      'msgbox "0"
      'right orb -> Gate001.collidable=False -> unless skillshot
      OrbUnlockShot = 1
      cLightShot1.state = 2
    Else
      'msgbox "1"
      'left orb shot -> Gate004.collidable=False -> unless skillshot
      OrbUnlockShot = 5
      cLightShot5.state = 2
    end if
    cLightMissionStart1.state=0
  end if
end sub

'*****************************************************************************************************************************************
'  Controller VBS stuff, but with b2s not started
'*****************************************************************************************************************************************
'***Controller.vbs version 1.2***'
'

Const directory = "HKEY_CURRENT_USER\SOFTWARE\Visual Pinball\Controller\"
Dim objShell
Dim PopupMessage
Dim B2SController
Dim Controller
Const DOFContactors = 1
Const DOFKnocker = 2
Const DOFChimes = 3
Const DOFBell = 4
Const DOFGear = 5
Const DOFShaker = 6
Const DOFFlippers = 7
Const DOFTargets = 8
Const DOFDropTargets = 9
Const DOFOff = 0
Const DOFOn = 1
Const DOFPulse = 2

Dim DOFeffects(9)
Dim B2SOn
Dim B2SOnALT

Sub LoadEM
  LoadController("EM")
End Sub

Sub LoadPROC(VPMver, VBSfile, VBSver)
  LoadVBSFiles VPMver, VBSfile, VBSver
  LoadController("PROC")
End Sub

Sub LoadVPM(VPMver, VBSfile, VBSver)
  LoadVBSFiles VPMver, VBSfile, VBSver
  LoadController("VPM")
End Sub

Sub LoadVPMALT(VPMver, VBSfile, VBSver)
  LoadVBSFiles VPMver, VBSfile, VBSver
  LoadController("VPMALT")
End Sub

Sub LoadVBSFiles(VPMver, VBSfile, VBSver)
  On Error Resume Next
  If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
  ExecuteGlobal GetTextFile(VBSfile)
  If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
  InitializeOptions
End Sub

Sub LoadVPinMAME
  Set Controller = CreateObject("VPinMAME.Controller")
  If Err Then MsgBox "Can't load VPinMAME." & vbNewLine & Err.Description
  If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
  If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
  On Error Goto 0
End Sub

Sub LoadController(TableType)
  Dim FileObj
  Dim DOFConfig
  Dim TextStr2
  Dim tempC
  Dim count
  Dim ISDOF
  Dim Answer

  B2SOn = False
  B2SOnALT = False
  tempC = 0
  on error resume next
  Set objShell = CreateObject("WScript.Shell")
  objShell.RegRead(directory & "ForceDisableB2S")
  If Err.number <> 0 Then
    PopupMessage = "This latest version of Controller.vbs stores its settings in the registry. To adjust the values, you must use VP 10.2 (or newer) and setup your configuration in the DOF section of the -Keys, Nudge and DOF- dialog of Visual Pinball."
    objShell.RegWrite directory & "ForceDisableB2S",0, "REG_DWORD"
    objShell.RegWrite directory & "DOFContactors",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFKnocker",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFChimes",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFBell",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFGear",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFShaker",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFFlippers",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFTargets",2, "REG_DWORD"
    objShell.RegWrite directory & "DOFDropTargets",2, "REG_DWORD"
    MsgBox PopupMessage
  End If
  tempC = objShell.RegRead(directory & "ForceDisableB2S")
  DOFeffects(1)=objShell.RegRead(directory & "DOFContactors")
  DOFeffects(2)=objShell.RegRead(directory & "DOFKnocker")
  DOFeffects(3)=objShell.RegRead(directory & "DOFChimes")
  DOFeffects(4)=objShell.RegRead(directory & "DOFBell")
  DOFeffects(5)=objShell.RegRead(directory & "DOFGear")
  DOFeffects(6)=objShell.RegRead(directory & "DOFShaker")
  DOFeffects(7)=objShell.RegRead(directory & "DOFFlippers")
  DOFeffects(8)=objShell.RegRead(directory & "DOFTargets")
  DOFeffects(9)=objShell.RegRead(directory & "DOFDropTargets")
  Set objShell = nothing

  If TableType = "PROC" or TableType = "VPMALT" Then
    If TableType = "PROC" Then
      Set Controller = CreateObject("VPROC.Controller")
      If Err Then MsgBox "Can't load PROC"
    Else
      LoadVPinMAME
    End If
    If tempC = 0 Then
      On Error Resume Next
      If Controller is Nothing Then
        Err.Clear
      Else
        Set B2SController = CreateObject("B2S.Server")
        If B2SController is Nothing Then
          Err.Clear
        Else
          B2SController.B2SName = B2ScGameName
          B2SController.Run()
          On Error Goto 0
          B2SOn = True
          B2SOnALT = True
        End If
      End If
    End If
  Else
    If tempC = 0 Then
      On Error Resume Next
      Set Controller = CreateObject("B2S.Server")
      If Controller is Nothing Then
        Err.Clear
        If TableType = "VPM" Then
          LoadVPinMAME
        End If
      Else
        Controller.B2SName = cGameName
        If TableType = "EM" Then
          Controller.Run()
        End If
        On Error Goto 0
        B2SOn = True
      End If
    Else
      If TableType = "VPM" Then
        LoadVPinMAME
      End If
    End If
    Set DOFConfig=Nothing
    Set FileObj=Nothing
  End If
End sub

Function SoundFX (Sound, Effect)
  If ((Effect = 0 And B2SOn) Or DOFeffects(Effect)=1) Then
    SoundFX = ""
  Else
    SoundFX = Sound
  End If
End Function

Function SoundFXDOF (Sound, DOFevent, State, Effect)
  If DOFeffects(Effect)=1 Then
    SoundFXDOF = ""
    DOF DOFevent, State
  ElseIf DOFeffects(Effect)=2 Then
    SoundFXDOF = Sound
    DOF DOFevent, State
  Else
    SoundFXDOF = Sound
  End If
End Function

Function SoundFXDOFALT (Sound, DOFevent, State, Effect)
  If DOFeffects(Effect)=1 Then
    SoundFXDOFALT = ""
    DOFALT DOFevent, State
  ElseIf DOFeffects(Effect)=2 Then
    SoundFXDOFALT = Sound
    DOFALT DOFevent, State
  Else
    SoundFXDOFALT = Sound
  End If
End Function


Sub DOF(DOFevent, State)
  If B2SOn Then
    If KeepDOFLogs then WriteToLog "DOF", "Event=" & DOFevent & " State=" & State
    If State = 2 Then
      Controller.B2SSetData DOFevent, 1:Controller.B2SSetData DOFevent, 0
    Else
      Controller.B2SSetData DOFevent, State
    End If
  End If
End Sub

Sub DOFALT(DOFevent, State)
  If B2SOnALT Then
    If State = 2 Then
      B2SController.B2SSetData DOFevent, 1:B2SController.B2SSetData DOFevent, 0
    Else
      B2SController.B2SSetData DOFevent, State
    End If
  End If
End Sub




'*****************************************************************************************************************************************
'   TABLE INITS & MATHS
'*****************************************************************************************************************************************

Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

Sub Table1_Init
  dim xx
  for each xx in underPFthings :
    if typename(xx) = "Wall" Then xx.SideVisible = false
    xx.visible=0
  Next
  For each xx in underPFclights : xx.state=False : Next

  NewLog
  WriteToLog "-------------", "TABLE INIT"

  GI_intensity
  Init_Insert_Blooms

  LoadEM
  Dim i
  Randomize

  FlexDMD_init

  'Impulse Plunger as autoplunger
  Const IMPowerSetting = 45 ' Plunger Power
  Const IMTime = 1.1        ' Time in seconds for Full Plunge
  Set plungerIM = New cvpmImpulseP
  With plungerIM
    .InitImpulseP swplunger, IMPowerSetting, IMTime
    .Random 1.5
    .InitExitSnd SoundFX("fx-kickout", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
    .CreateEvents "plungerIM"
  End With

  Set plungerKB = New cvpmImpulseP
  With plungerKB
    .InitImpulseP swKB, IMPowerSetting, IMTime
    .Random 1.5
    .CreateEvents "plungerKB"
  End With

  Set VorMag = New cvpmMagnet
  With VorMag
    .InitMagnet VortexMagnet,16
    .GrabCenter=False
    .strength = 15
    .CreateEvents "VorMag"
  end With

  'MiMag
  Set MiMag = New cvpmMagnet
  With MiMag
    .InitMagnet MimaMagnet,16
    .GrabCenter=False
    .strength = 15
    .CreateEvents "MiMag"
  end With

    Set GrabMag= New cvpmMagnet
    With GrabMag
        .InitMagnet GrabMagnet, 16
        .GrabCenter = True
    .strength = 12
        .CreateEvents "GrabMag"
    End With

  Loadhs
  bAttractMode = False
  bEndCredits = False
  bOnTheFirstBall = False
  bBallInPlungerLane = False
  bBallSaverActive = False
  bBallSaverReady = False
  bMultiBallMode = False
  PupBumperTracy = False
  PupBumpertenthousand = False
  PupBumpersuperjets = False
  bMultiballReady = False
  bCoreyMBOngoing = False
  bTracyMBOngoing = False
  bMissionMode = False
  bWizardMode = False
  bGameInPlay = False
  bAutoPlunger = False
  bTracyJPReady = False
  bTracySuperJPReady = False
  bCoreyJPReady = False
  bCoreySuperJPReady = False
  bMimaJPReady = False
  bMimaSuperJPReady = False
  bMusicOn = True
  BallsOnPlayfield = 0
  BallsInHole = 0
  NumLocksMima = 0
  Tilt = 0
  TiltSensitivity = 6
  Tilted = False
  bBonusHeld = False
  bJustStarted = True
  bFlippersEnabled = True
  bPortalFlippersEnabled = False
  PointMultiplier = 1
  SpinnerMultiplier = 1

  if ScorbitActive = 1 then vpmtimer.addtimer 2000, "showQRPairImage '"

  if Credits > 0 then DOF 122,1 'start button light

  ResetMimaTargets

  if not Slammed then
    StartAttractMode   ' moved to end of intro
    Slammed = false
  end if

  if StartupTune then vpmtimer.addtimer 2000, "PlaySong 5'"

  vpmtimer.addtimer 3000, "startb2s 4'"
PUPInit
End Sub

Sub Table1_Exit
  WriteToLog "Table1_Exit", ""
  Scorbit.StopSession2 Score(1), Score(2), Score(3), Score(4), PlayersPlayingGame, True   ' In case you stop mid game
  StopSong
  If B2SOn Then
    Controller.Pause = False
    Controller.Stop
  End If

' If Not FlexDMD is Nothing Then
'   FlexDMD.Show = False
'   FlexDMD.Run = False
'   FlexDMD = NULL
' End If
End Sub


Sub Table1_Paused
  WriteToLog "Table1_Paused", ""
  'pause the song
  If oPlayer1.playState = 3 Then oPlayer1.controls.pause
  If oPlayer2.playState = 3 Then oPlayer2.controls.pause
End Sub


Sub Table1_UnPaused
  WriteToLog "Table1_UnPaused", ""
  'play the taused song
  If oPlayer1.playState = 2 Then oPlayer1.controls.play
  If oPlayer2.playState = 2 Then oPlayer2.controls.play
End Sub



'*************************************************
'******** FLEXDMD ********************************
'*************************************************
' FlexDMD constants
Const   FlexDMD_RenderMode_DMD_GRAY = 0, _
    FlexDMD_RenderMode_DMD_GRAY_4 = 1, _
    FlexDMD_RenderMode_DMD_RGB = 2, _
    FlexDMD_RenderMode_SEG_2x16Alpha = 3, _
    FlexDMD_RenderMode_SEG_2x20Alpha = 4, _
    FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num = 5, _
    FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num_4x1Num = 6, _
    FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num = 7, _
    FlexDMD_RenderMode_SEG_2x7Num_2x7Num_10x1Num = 8, _
    FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num_gen7 = 9, _
    FlexDMD_RenderMode_SEG_2x7Num10_2x7Num10_4x1Num = 10, _
    FlexDMD_RenderMode_SEG_2x6Num_2x6Num_4x1Num = 11, _
    FlexDMD_RenderMode_SEG_2x6Num10_2x6Num10_4x1Num = 12, _
    FlexDMD_RenderMode_SEG_4x7Num10 = 13, _
    FlexDMD_RenderMode_SEG_6x4Num_4x1Num = 14, _
    FlexDMD_RenderMode_SEG_2x7Num_4x1Num_1x16Alpha = 15, _
    FlexDMD_RenderMode_SEG_1x16Alpha_1x16Num_1x7Num = 16

Const   FlexDMD_Align_TopLeft = 0, _
    FlexDMD_Align_Top = 1, _
    FlexDMD_Align_TopRight = 2, _
    FlexDMD_Align_Left = 3, _
    FlexDMD_Align_Center = 4, _
    FlexDMD_Align_Right = 5, _
    FlexDMD_Align_BottomLeft = 6, _
    FlexDMD_Align_Bottom = 7, _
    FlexDMD_Align_BottomRight = 8


Dim FlexDMD ',FlexDMD2
Dim Frame
Dim DMDMode

Sub FlexDMD_init
  WriteToLog "FlexDMD_init", ""
  Dim fso,curdir
  Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
    If FlexDMD is Nothing Then
        MsgBox "No FlexDMD found. This table will NOT run without it."
        Exit Sub
    End If
  SetLocale(1033)

  Select Case DMDColor
    Case 1: FlexDMD.Color = RGB(255, 88, 32)
    Case 2: FlexDMD.Color = RGB(200, 10, 10)
    Case 3: FlexDMD.Color = RGB(250, 250, 250)
  End Select

  if Not HasPuP then
    With FlexDMD
      .GameName = cGameName
      .TableFile = Table1.Filename & ".vpx"
      .RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
      .Width = 128
      .Height = 32
      .Clear = True
      .ProjectFolder = "./BloodMachinesDMD/"
      .Run = true
    End With
  elseif forceFlexOn Then
    With FlexDMD
      .GameName = cGameName
      .TableFile = Table1.Filename & ".vpx"
      .RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
      .Width = 128
      .Height = 32
      .Clear = True
      .ProjectFolder = "./BloodMachinesDMD/"
      .Run = true
    End With
  Else
    With FlexDMD
      .GameName = cGameName
      .TableFile = Table1.Filename & ".vpx"
      .RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
      .Width = 128
      .Height = 32
      .Clear = True
      .ProjectFolder = "./BloodMachinesDMD/"
      .Run = false
    End With
  end if


'
' Set FlexDMD2 = CreateObject("FlexDMD.FlexDMD")
'    If FlexDMD2 is Nothing Then
'        MsgBox "No FlexDMD found. This table will NOT run without it."
'        Exit Sub
'    End If
' With FlexDMD2
'   .GameName = cGameName
'   .TableFile = Table1.Filename & "2.vpx"
'   .Color = RGB(77, 255, 32)
'   .RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
'   .Width = 120
'   .Height = 110
'   .Clear = True
'   .Run = True
'   .ProjectFolder = "./BloodMachinesDMD/"
'   .show = false
' End With
'

' CreateApronDMD
'
' CreateDMD1


  CreateDMD_intro
  If VRRoom > 0 Then flexDMDonPF.enabled=True

  MissionTimer.Enabled = True

  DMDTimer.interval=17
  DMDTimer.enabled=true

End Sub


Sub flexDMDonPF_Timer    ' fixing delete
  Dim DMDp
  DMDp = FlexDMD.DmdPixels
  If Not IsEmpty(DMDp) Then
    DMDWidth = FlexDMD.Width
    DMDHeight = FlexDMD.Height
    DMDColoredPixels = DMDp
  End If
End Sub

Dim FontScoreActive
Dim FontScoreInactiv1
Dim FontScoreInactiv2

Dim IntroOver
Introover=False

Dim G_Image, GaldorQ1, GaldorQ2, GaldorFlash , GaldorFlashRed
G_Image=30 : G_ImageEnd=33

Dim GaldorReset : GaldorReset=false
Sub galdor_Timer
  If GaldorReset Then Next_Galdor : GaldorReset=False
  'debug.print "galdor" & G_Image
  If G_Image=30 And galdorq1<>"" Then Next_Galdor
  If ( Frame Mod 4 ) = 1 And G_Image>50 Then
    If GaldorQ2>16 Then
      G_Image=G_Image+1
      If G_Image > G_ImageEnd Then Next_Galdor
    End If
    G_Image=G_Image+1
    If G_Image > G_ImageEnd Then Next_Galdor
  End If

  If ( Frame Mod 2 ) = 1 and GaldorFlash>0 Then
    galdorflash=galdorFlash-1
    If GaldorFlashRed=0 Then apron_instruments001.image = "ZZ_00" Else apron_instruments001.image = "ZZ_0"
  Else
    apron_instruments001.image = "ZZ_" & G_Image
  End If

End Sub

Dim ShortVersion
Dim G_ImageEnd

Sub Next_Galdor
  If GaldorQ1 =""  Then
    GaldorQ2=0 : G_image=30
  Else
      Dim repeatcheck
      repeatcheck=Left(GaldorQ1,1)
      If repeatcheck="X" Then
        Shortversion=False
        repeatcheck=Mid(GaldorQ1,2,1)
      ElseIf repeatcheck="Y" Then
        repeatcheck=Mid(GaldorQ1,2,1)
        Shortversion=True
      ElseIf repeatcheck="Z" Then
        GaldorQ1=Right(GaldorQ1,len(GaldorQ1)-1)
        repeatcheck=Left(GaldorQ1,1)
        Shortversion=True
      Else
        Shortversion=False
      End If

      Select Case repeatcheck
        Case "a" : GaldorQ2=1  : G_image=93   :  G_ImageEnd=158  'Look and eyes wink
        Case "b" : GaldorQ2=2  : G_image=159  :  G_ImageEnd=225  'Looking pissed off
        Case "c" : GaldorQ2=3  : G_image=226  :  G_ImageEnd=303  'Eyes look down and up
        Case "d" : GaldorQ2=4  : G_image=304  :  G_ImageEnd=340  'Looking straight
        Case "e" : GaldorQ2=5  : G_image=340  :  G_ImageEnd=378  'Looks Shocked (laser beam reflection)
        Case "f" : GaldorQ2=6  : G_image=379  :  G_ImageEnd=407  'Galdor hurting/defeat
        Case "g" : GaldorQ2=7  : G_image=408  :  G_ImageEnd=518  'Listening and head move
        Case "h" : GaldorQ2=8  : G_image=519  :  G_ImageEnd=567 'Obey!
        Case "i" : GaldorQ2=9  : G_image=568  :  G_ImageEnd=703 'Looking left and right (Deepfake)
        Case "j" : GaldorQ2=10 : G_image=704  :  G_ImageEnd=895  'Looking around (Deepfake)
        Case "k" : GaldorQ2=11 : G_image=896  :  G_ImageEnd=925  'Yell/Angry (edited from Obey clip)
        Case "l" : GaldorQ2=12 : G_image=926  :  G_ImageEnd=1051  'opening eyes, looking around and closing eye
        Case "m" : GaldorQ2=13 : G_image=1052 :  G_ImageEnd=1169  'Straight stare and eye blink.

        Case "A" : GaldorQ2=14  : G_image=93  :  G_ImageEnd=158  'Look and eyes wink
        Case "B" : GaldorQ2=15  : G_image=159 :  G_ImageEnd=225  'Looking pissed off
        Case "C" : GaldorQ2=16  : G_image=226 :  G_ImageEnd=303  'Eyes look down and up
        Case "D" : GaldorQ2=17  : G_image=304 :  G_ImageEnd=340  'Looking straight
        Case "E" : GaldorQ2=18  : G_image=340 :  G_ImageEnd=378  'Looks Shocked (laser beam reflection)
        Case "F" : GaldorQ2=19  : G_image=379 :  G_ImageEnd=407  'Galdor hurting/defeat
        Case "G" : GaldorQ2=20  : G_image=408 :  G_ImageEnd=518  'Listening and head move
        Case "H" : GaldorQ2=21 : G_image=519  :  G_ImageEnd=567 'Obey!
        Case "I" : GaldorQ2=22 : G_image=568  :  G_ImageEnd=703 'Looking left and right (Deepfake)
        Case "J" : GaldorQ2=23 : G_image=704  :  G_ImageEnd=895  'Looking around (Deepfake)
        Case "K" : GaldorQ2=24 : G_image=896  :  G_ImageEnd=925  'Yell/Angry (edited from Obey clip)
        Case "L" : GaldorQ2=25 : G_image=926  :  G_ImageEnd=1051  'opening eyes, looking around and closing eye
        Case "M" : GaldorQ2=26 : G_image=1052 :  G_ImageEnd=1169  'Straight stare and eye blink.       ' DOUBLE SPEED = CAPS
      End Select
    If ShortVersion Then G_image=G_image+5 : G_ImageEnd=G_ImageEnd-5
    If Not Left(GaldorQ1,1)="X" And Not Left(GaldorQ1,1)="Y" Then GaldorQ1=Right(GaldorQ1,len(GaldorQ1)-1)
  End If
End Sub



'0 zz_01 to zz_45 = Globe Closed (loop)
'1 zz_46 to zz_92 = Globe Opening
'2 zz_93 to zz_158 = Look and eyes wink
'3 zz_159 to zz_225 = Looking pissed off
'4 zz_226 to zz_303 = Eyes look down and up
'5 zz_304 to zz_339 = Looking straight
'6 zz_340 to zz_377 = Looks Shocked (laser beam reflection)
'7 zz_379 to zz_407 = Galdor hurting/defeat
'8 zz_408 to zz_469 = Listening and head move
'9 zz_470 to zz_518 = Globe Closing
'zz_519 to zz_567 = 'Obey!
'zz_568 to zz_703 = 'Looking left and right (Deepfake)
'zz_704 to zz_895 = 'Looking around (Deepfake)
'zz_896 to zz_925 = 'Yell/Angry (edited from Obey clip)
'926 to 1051 => 'opening eyes, looking around and closing eyes
'1052 to 1169 => 'Straight stare and eye blink.


Sub CreateApronDMD   ' fixing delete


  Set FontScoreActive = FlexDMD2.NewFont("FONTS/udmd-f42by78.fnt", vbWhite, RGB(0,0,0), 1)
  Dim scene : Set scene = FlexDMD2.NewGroup("Score2")

  scene.AddActor FlexDMD2.NewImage("counterBG1" ,  "Misc/counter1.png" )  : scene.GetImage("counterBG1").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG2" ,  "Misc/counter2.png" )  : scene.GetImage("counterBG2").Visible = True
  scene.AddActor FlexDMD2.NewImage("counterBG3" ,  "Misc/counter3.png" )  : scene.GetImage("counterBG3").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG4" ,  "Misc/counter4.png" )  : scene.GetImage("counterBG4").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG5" ,  "Misc/counter5.png" )  : scene.GetImage("counterBG5").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG6" ,  "Misc/counter6.png" )  : scene.GetImage("counterBG6").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG7" ,  "Misc/counter7.png" )  : scene.GetImage("counterBG7").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG8" ,  "Misc/counter8.png" )  : scene.GetImage("counterBG8").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG9" ,  "Misc/counter9.png" )  : scene.GetImage("counterBG9").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG10" , "Misc/counter10.png" ) : scene.GetImage("counterBG10").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG11" , "Misc/counter11.png" ) : scene.GetImage("counterBG11").Visible = False
  scene.AddActor FlexDMD2.NewImage("counterBG12" , "Misc/counter12.png" ) : scene.GetImage("counterBG12").Visible = False

  scene.AddActor FlexDMD2.NewLabel("apron1", FontScoreActive, "000")
  scene.GetLabel("apron1").SetAlignedPosition 60,55, FlexDMD_Align_Center
  scene.GetLabel("apron1").visible=False

  FlexDMD2.LockRenderThread
  FlexDMD2.RenderMode = FlexDMD_RenderMode_DMD_RGB
  FlexDMD2.Stage.RemoveAll
  FlexDMD2.Stage.AddActor scene
  FlexDMD2.Show = False
  FlexDMD2.UnlockRenderThread

' MissionTimer.Enabled = True


End Sub


Sub CreateDMD_intro
  Frame = 0
' AudioCallout "boot"
  Dim i
  Set FontScoreActive = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, RGB(0,0,0), 0)
  Set FontScoreInactiv1 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(128, 128, 128), vbWhite, 0)
  Set FontScoreInactiv2 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(255, 255, 255), vbWhite, 0)
  Dim scene : Set scene = FlexDMD.NewGroup("Score")

  dim vidvid

' set vidvid=FlexDMD.Newvideo ("test","Videos//Bloodmachines logo.gif")
' scene.AddActor vidvid
' scene.Getvideo("test").SetBounds 0, 0, 128, 32
' scene.Getvideo("test").Visible = True


  scene.AddActor FlexDMD.NewImage("Border2" , "Misc/bg_border2.png" )
  scene.GetImage("Border2").visible=False
  scene.AddActor FlexDMD.NewImage("Border3" , "Misc/bg_border3.png" )
  scene.GetImage("Border3").visible=False

  FlexDMD.LockRenderThread
  FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
  FlexDMD.Stage.RemoveAll
  FlexDMD.Stage.AddActor scene
  If VRRoom = 0 Then  FlexDMD.Show = True Else FlexDMD.Show = False

  DMDmode=2

  FlexDMD.UnlockRenderThread

End Sub

'cybertopiaf7by13
'cybertopiaf7by13-outline

'Dim FontScoreActive3, FontScoreActive4,FontScoreActive5, FontScoreActive6,
Dim FontScoreActiveG
Dim FontWhite3, FontBlack3, Font6by10W, Font6by10B, Font6by10G, Font5by7W, Font5by7G, Font7by13W



Sub CreateDMD1
  Frame = 0

  Dim i
  Set FontWhite3 = FlexDMD.NewFont("FONTS/compacta_black_10x9.fnt", vbWhite, RGB(0,0,0), 0)
  Set FontBlack3 = FlexDMD.NewFont("FONTS/compacta_black_10x9.fnt", vbBlack, RGB(0,0,0), 0)
  Set Font6by10W = FlexDMD.NewFont("FONTS/udmd-f6by10.fnt", vbWhite, RGB(0,0,0), 0)
  Set Font6by10G = FlexDMD.NewFont("FONTS/udmd-f6by10.fnt", RGB(60,60,60), RGB(0,0,0), 0)
  Set Font6by10B = FlexDMD.NewFont("FONTS/udmd-f6by10.fnt", vbBlack, RGB(0,0,0), 0)

  Set Font5by7W = FlexDMD.NewFont("FONTS/udmd-f5by7.fnt", vbWhite, RGB(0,0,0), 0)
  Set Font5by7G = FlexDMD.NewFont("FONTS/udmd-f5by7.fnt", RGB(60,60,60), RGB(0,0,0), 0)

  Set FontScoreActive =  FlexDMD.NewFont("FONTS/compacta_black_10x18.fnt", vbWhite, RGB(0,0,0), 0)
  Set FontScoreActiveG = FlexDMD.NewFont("FONTS/compacta_black_10x18.fnt", RGB(60,60,60),RGB(0,0,0), 0)
' Set FontScoreActive3 = FlexDMD.NewFont("FONTS/BM-Font-2.fnt", vbWhite, RGB(0,0,0), 0)
' Set FontScoreActive4 = FlexDMD.NewFont("FONTS/udmd-f5by7.fnt", vbWhite, RGB(90,90,90), 1)
' Set FontScoreActive5 = FlexDMD.NewFont("FONTS/udmd-f5by7bot.fnt", vbWhite, RGB(90,90,90), 1)
' Set FontScoreActive6 = FlexDMD.NewFont("FONTS/udmd-f5by7top.fnt", vbWhite, RGB(90,90,90), 1)

  Set Font7by13W = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(255, 255, 255), vbWhite, 0)


' Set FontScoreInactiv1 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(128, 128, 128), vbWhite, 0)
' Set FontScoreInactiv2 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(255, 255, 255), vbWhite, 0)
  Set FontScoreInactiv1 = FlexDMD.NewFont("FONTS/teeny_tiny_pixls-5.fnt", RGB(128, 128, 128), vbWhite, 0)
  Set FontScoreInactiv2 = FlexDMD.NewFont("FONTS/teeny_tiny_pixls-5.fnt", RGB(255, 255, 255), vbWhite, 0)
  Dim scene : Set scene = FlexDMD.NewGroup("Score")


  ' score
  scene.AddActor FlexDMD.NewLabel("Score_1", FontScoreActive, "0")

  ' blackoverlay for score !
  scene.AddActor FlexDMD.NewImage("ScOverlay1" , "ScoreOverlay/sc_a13.png" )
  scene.GetImage("ScOverlay1").SetPosition 0,-32
  scene.GetImage("ScOverlay1").visible=True


  scene.AddActor FlexDMD.NewImage("highlight1" , "misc/3lines1highlight.png" ) ' highlight1 Underlay for mission screens  12px high move it around -1 10 21 height
  scene.GetImage("highlight1").visible=False


  scene.AddActor FlexDMD.NewLabel("Splash", FontScoreInactiv1, ".")
  scene.GetLabel("Splash").Visible=False
  scene.AddActor FlexDMD.NewLabel("Splash2", FontScoreInactiv1, ".")
  scene.GetLabel("Splash2").Visible=False
  scene.AddActor FlexDMD.NewLabel("Splash3", FontScoreInactiv1, ".")
  scene.GetLabel("Splash3").Visible=False



  Scene.AddActor FlexDMD.NewLabel("overlaytext4", FontScoreActive, " ") : Scene.GetLabel("overlaytext4").visible = False
  'ball/player
' scene.AddActor FlexDMD.NewLabel("Ball", FontScoreInactiv1, "BALL 0")
' scene.GetLabel("Ball").SetAlignedPosition 64,30, FlexDMD_Align_Center '116
' scene.GetLabel("Ball").visible = false

'
' scene.AddActor FlexDMD.NewLabel("Player1", FontScoreInactiv1, "P1")
' scene.GetLabel("Player1").SetAlignedPosition 5,30, FlexDMD_Align_Center
' scene.AddActor FlexDMD.NewLabel("Player2", FontScoreInactiv1, "P2")
' scene.GetLabel("Player2").SetAlignedPosition 14,30, FlexDMD_Align_Center
' scene.AddActor FlexDMD.NewLabel("Player3", FontScoreInactiv1, "P3")
' scene.GetLabel("Player3").SetAlignedPosition 23,30, FlexDMD_Align_Center
' scene.AddActor FlexDMD.NewLabel("Player4", FontScoreInactiv1, "P4")
' scene.GetLabel("Player4").SetAlignedPosition 32,30, FlexDMD_Align_Center

  scene.AddActor FlexDMD.NewLabel("Player1", FontScoreInactiv1, "00")
  scene.GetLabel("Player1").SetAlignedPosition 0,3, FlexDMD_Align_Left

  scene.AddActor FlexDMD.NewLabel("Player2", FontScoreInactiv1, "00")
  scene.GetLabel("Player2").SetAlignedPosition 128,3, FlexDMD_Align_Right

  scene.AddActor FlexDMD.NewLabel("Player3", FontScoreInactiv1, "00")
  scene.GetLabel("Player3").SetAlignedPosition 0,30, FlexDMD_Align_Left

  scene.AddActor FlexDMD.NewLabel("Player4", FontScoreInactiv1, "00")
  scene.GetLabel("Player4").SetAlignedPosition 128,30, FlexDMD_Align_Right



' these moved infron of mission background image as a test
  scene.AddActor FlexDMD.NewImage("Background2" , "Misc/bg_bumper2.png" )
  scene.GetImage("Background2").visible=False
  scene.AddActor FlexDMD.NewImage("Background3" , "Misc/bg_bumper4.png" )
  scene.GetImage("Background3").visible=False
  scene.AddActor FlexDMD.NewImage("Background4" , "Misc/bg_penta1.png" )
  scene.GetImage("Background4").visible=False
  scene.AddActor FlexDMD.NewImage("Background5" , "Misc/bg_penta2.png" )
  scene.GetImage("Background5").visible=False
  scene.AddActor FlexDMD.NewImage("Background6" , "Misc/bg_barLeft.png" )
  scene.GetImage("Background6").visible=False
  scene.AddActor FlexDMD.NewImage("Background7" , "Misc/bg_barRight.png" )
  scene.GetImage("Background7").visible=False





  scene.AddActor FlexDMD.NewImage("skillshot" , "Misc/coreydmd3.png" )   ' different pictures for splashscreens
  scene.GetImage("skillshot").visible=False

  scene.AddActor FlexDMD.NewLabel("SplashScreen1", FontWhite3, " ")  '  3 lines on top of the splash screens
  scene.GetLabel("SplashScreen1").visible=False
  scene.AddActor FlexDMD.NewLabel("SplashScreen2", FontWhite3, " ")
  scene.GetLabel("SplashScreen2").visible=False
  scene.AddActor FlexDMD.NewLabel("SplashScreen3", FontWhite3, " ")
  scene.GetLabel("SplashScreen3").visible=False


'ontop of everything execpt newly videos
  scene.AddActor FlexDMD.NewImage("ballsave1" , "Misc/ballsave1.png" )
  scene.GetImage("ballsave1").visible=False
  scene.AddActor FlexDMD.NewImage("ballsave2" , "Misc/ballsave2.png" )
  scene.GetImage("ballsave2").visible=False
  scene.AddActor FlexDMD.NewImage("balllost1" , "Misc/balllost1.png" )
  scene.GetImage("balllost1").visible=False
  scene.AddActor FlexDMD.NewImage("balllost2" , "Misc/balllost2.png" )
  scene.GetImage("balllost2").visible=False


  scene.AddActor FlexDMD.NewImage("Border2" , "Misc/bg_border2.png" )
  scene.GetImage("Border2").visible=False
  scene.AddActor FlexDMD.NewImage("Border3" , "Misc/bg_border3.png" )
  scene.GetImage("Border3").visible=False

' fixing  add images for attract highscore EOB bonus here  ... need 1-2-3-4-5etc after name used in commands
  Scene.AddActor FlexDMD.NewImage("vpwlogo1","Misc/VPWLogo32.png") : Scene.Getimage("vpwlogo1").visible = False
  Scene.AddActor FlexDMD.NewImage("attract1","Misc/Mission1BG.png") : Scene.Getimage("attract1").visible = False

  Scene.AddActor FlexDMD.NewImage("extraball1","Misc/extraball1.png") : Scene.Getimage("extraball1").visible = False
  Scene.AddActor FlexDMD.NewImage("extraball2","Misc/extraball2.png") : Scene.Getimage("extraball2").visible = False
'DMD_ShowImages "extraball",2,50,2000,0 ' eb fast blink
'DMD_ShowImages "extraballb",2,50,1,0 ' eb vanish ( addtimer that is less than 2000 )
  Scene.AddActor FlexDMD.NewImage("extraballb1","Misc/extraball3.png") : Scene.Getimage("extraballb1").visible = False
  Scene.AddActor FlexDMD.NewImage("extraballb2","Misc/extraball4.png") : Scene.Getimage("extraballb2").visible = False
  Scene.AddActor FlexDMD.NewImage("extraballb3","Misc/extraball5.png") : Scene.Getimage("extraballb3").visible = False
  Scene.AddActor FlexDMD.NewImage("extraballb4","Misc/extraball6.png") : Scene.Getimage("extraballb4").visible = False
  Scene.AddActor FlexDMD.NewImage("extraballb5","Misc/extraball7.png") : Scene.Getimage("extraballb5").visible = False


  Scene.AddActor FlexDMD.NewImage("jackpot1","Misc/jackpot1.png") : Scene.Getimage("jackpot1").visible = False
  Scene.AddActor FlexDMD.NewImage("jackpot2","Misc/jackpot2.png") : Scene.Getimage("jackpot2").visible = False
  Scene.AddActor FlexDMD.NewImage("jackpot3","Misc/jackpot3.png") : Scene.Getimage("jackpot3").visible = False
'DMD_ShowImages "jackpot",3,50,2000,0 ' jp fast blink

  Scene.AddActor FlexDMD.NewImage("jolt1","Misc/jolt1.png") : Scene.Getimage("jolt1").visible = False
  Scene.AddActor FlexDMD.NewImage("jolt2","Misc/jolt2.png") : Scene.Getimage("jolt2").visible = False
  Scene.AddActor FlexDMD.NewImage("jolt3","Misc/jolt3.png") : Scene.Getimage("jolt3").visible = False
'DMD_ShowImages "jolt",3,50,2000,0 ' jolt fast blink
  Scene.AddActor FlexDMD.NewImage("highscore1","Misc/highscore1.png") : Scene.Getimage("highscore1").visible = False
  Scene.AddActor FlexDMD.NewImage("highscore2","Misc/highscore2.png") : Scene.Getimage("highscore2").visible = False

  Scene.AddActor FlexDMD.NewLabel("overlaytext1", FontScoreActive, " ") : Scene.GetLabel("overlaytext1").visible = False
  Scene.AddActor FlexDMD.NewLabel("overlaytext2", FontScoreActive, " ") : Scene.GetLabel("overlaytext2").visible = False
  Scene.AddActor FlexDMD.NewLabel("overlaytext3", FontScoreActive, " ") : Scene.GetLabel("overlaytext3").visible = False



' insert coin on top !

  scene.AddActor FlexDMD.NewImage("insertcoin1" , "Misc/InsertCoin_1.png" )
  scene.GetImage("insertcoin1").visible=False
  scene.AddActor FlexDMD.NewImage("insertcoin2" , "Misc/InsertCoin_2.png" )
  scene.GetImage("insertcoin2").visible=False




' under = video overlay with Text
  scene.AddActor FlexDMD.Newvideo ("videoplaying","Videos/Vascan_Cory Walking trough Portal.gif")
  scene.Getvideo("videoplaying").Visible = False
  scene.AddActor FlexDMD.NewLabel("videotext", FontScoreInactiv1, VidList(1) )
  scene.Getlabel("videotext").Visible = False

  scene.AddActor FlexDMD.NewLabel("Score_2", FontScoreActive, "0")
  scene.GetLabel("Score_2").Visible = False
  scene.AddActor FlexDMD.NewLabel("Score_3", FontScoreActive, "0")
  scene.GetLabel("Score_3").Visible = False


  FlexDMD.LockRenderThread
  FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
  FlexDMD.Stage.RemoveAll
  FlexDMD.Stage.AddActor scene
  If VRRoom = 0 Then  FlexDMD.Show = True Else FlexDMD.Show = False

  DMDmode=1

  FlexDMD.UnlockRenderThread

End Sub




Dim playerblink, Ballblink, CreditBlink, Borderblink
Dim BG_Bump1, BG_Bump2, bumpersplash
Dim DMDExtraBall
Dim BG_Stars
Dim BG_BarLeft, BG_BarRight
Dim ShowMissionFailed, ShowMissionComplete
Dim Show_TracyHits, Tracy_bumpers
Dim showMissionVideo,showMissionVideoCounter









Dim fontblink2
Dim OverlayPos, OverlayDir
Dim Overlay1 : Overlay1=True


'Dim Videotest : Videotest=False
Dim VideoPlaying : VideoPlaying=1


Sub DMDShowVideos

  Select Case showMissionVideo
' MISSION 1 ********
    Case  1 : CurrentVideo = 82 : Start_Video : showMissionVideoCounter=Frame+266 ' mission1intro
    Case  2 : CurrentVideo = 30 : Start_Video : showMissionVideoCounter=Frame+133 ' loading gun
    Case  3 : CurrentVideo =  1 : Start_Video : showMissionVideoCounter=Frame+216 ' failed
    Case  4 : CurrentVideo = 70 : Start_Video : showMissionVideoCounter=Frame+70    ' shooting
    Case  5 : CurrentVideo = 86 : Start_Video : showMissionVideoCounter=Frame+207 ' Completed looking orb
' MISSION 2
    Case  6 : CurrentVideo = 27 : Start_Video : showMissionVideoCounter=Frame+220 ' start2  corey rising arms
    Case  7 : CurrentVideo = 51 : Start_Video : showMissionVideoCounter=Frame+200 ' Mima_pass camera to distance
    Case  8 : CurrentVideo = 52 : Start_Video : showMissionVideoCounter=Frame+200 ' Mima_Rising in space
    Case  9 : CurrentVideo = 29 : Start_Video : showMissionVideoCounter=Frame+200 ' Corey_tearing ship open
' MISSION 3
    Case 10 : CurrentVideo = 76 : Start_Video : showMissionVideoCounter=Frame+301 ' Spaceship Warpspeed
    Case 11 : CurrentVideo = 50 : Start_Video : showMissionVideoCounter=Frame+88  ' Mima_Floating over camera
    Case 12 : CurrentVideo = 47 : Start_Video : showMissionVideoCounter=Frame+200 ' Mima_Beauty shot
    Case 13 : CurrentVideo = 49 : Start_Video : showMissionVideoCounter=Frame+200 ' Mima_Floating in space
' MISSION 4
    Case 14 : CurrentVideo = 59 : Start_Video : showMissionVideoCounter=Frame+220 ' Ship Flying trough Dumpe
    Case 15 : CurrentVideo = 60 : Start_Video : showMissionVideoCounter=Frame+50  ' Ship Flying trough Dump_2
    Case 16 : CurrentVideo = 58 : Start_Video : showMissionVideoCounter=Frame+200 ' Ship Flyby
    Case 17 : CurrentVideo = 56 : Start_Video : showMissionVideoCounter=Frame+200 ' Scavengers pulling wires out of ship failed
' MISSION 5
    Case 18 : CurrentVideo = 90 : Start_Video : showMissionVideoCounter=Frame+220 ' Walking in cockpit.gif")   ' Portal
    Case 19 : CurrentVideo = 87 : Start_Video : showMissionVideoCounter=Frame+50  ' Vascan_Pulls off circuit board.gif")   ' Portal Target
    Case 20 : CurrentVideo = 89 : Start_Video : showMissionVideoCounter=Frame+50  ' Vascan_Throws board to Lago.gif")    ' Portal Target2
    Case 21 : CurrentVideo = 83 : Start_Video : showMissionVideoCounter=Frame+200 ' Vascan_Dies and falls on floor.gif")    ' Portal sucess
    Case 22 : CurrentVideo = 84 : Start_Video : showMissionVideoCounter=Frame+200 ' Vascan_Gets shot_Closeup"
' MISSION 6
    Case 23 : CurrentVideo = 41 : Start_Video : showMissionVideoCounter=Frame+220 ' Lago Hail Mary.gif
    Case 24 : CurrentVideo = 44 : Start_Video : showMissionVideoCounter=Frame+50  ' Lago_Hart of steel_thumbs up.gif")   ' Heartofsteel bumper
    Case 25 : CurrentVideo = 40 : Start_Video : showMissionVideoCounter=Frame+200 ' Lago walks to camera.gif")    ' Heartofsteel sucess
    Case 26 : CurrentVideo = 46 : Start_Video : showMissionVideoCounter=Frame+200 ' Lago_with hatchet.gif")   ' Heartofsteel failed
' MISSION 7
    Case 27 : CurrentVideo = 78 : Start_Video : showMissionVideoCounter=Frame+220 ' Tracy 360 shot.gif")   ' tracy fury
    Case 28 : CurrentVideo = 80 : Start_Video : showMissionVideoCounter=Frame+50  'racy_Head Shaking.gif")   ' tracy fury bumper
    Case 29 : CurrentVideo = 77 : Start_Video : showMissionVideoCounter=Frame+200 'Supernova appears.gif")    ' tracy fury sucess
    Case 30 : CurrentVideo = 69 : Start_Video : showMissionVideoCounter=Frame+200 '/Shooting circuit board.gif")   ' tracy fury failed
' Corey MB
    Case 31 : CurrentVideo = 7  : Start_Video : showMissionVideoCounter=Frame+250 '"Corey putting on mask"  'MB ready
    Case 32 : CurrentVideo = 9  : Start_Video : showMissionVideoCounter=Frame+100 '"Corey attack end"     ' MB end
    Case 33 : CurrentVideo = 10  : Start_Video : showMissionVideoCounter=Frame+100  '"Corey_Attack_hit1"    'drain while Corey MB
    Case 34 : CurrentVideo = 11  : Start_Video : showMissionVideoCounter=Frame+160  '"Corey_Attack_hit2"    ' Corey MB JP
    Case 35 : CurrentVideo = 12  : Start_Video : showMissionVideoCounter=Frame+200  ''VidList(12)="Corey_Attack_hit3" 'hit ship
    Case 36 : CurrentVideo = 13  : Start_Video : showMissionVideoCounter=Frame+50 ''VidList(13)="Corey_Attack_hit4" 'hit ship
    Case 37 : CurrentVideo = 14  : Start_Video : showMissionVideoCounter=Frame+120  ''VidList(14)="Corey_Attack_hit5" 'hit ship
    Case 38 : CurrentVideo = 15  : Start_Video : showMissionVideoCounter=Frame+120  ''VidList(15)="Corey_Attack_hit6" 'hit ship
    Case 39 : CurrentVideo = 16  : Start_Video : showMissionVideoCounter=Frame+50 ''VidList(16)="Corey_Attack_move1" 'MB right orb
    Case 40 : CurrentVideo = 17  : Start_Video : showMissionVideoCounter=Frame+58 ''VidList(17)="Corey_Attack_move2"  'MB left orb
    Case 41 : CurrentVideo = 18  : Start_Video : showMissionVideoCounter=Frame+140  ''VidList(18)="Corey_Attack_move3"  'fist up
    Case 42 : CurrentVideo = 19  : Start_Video : showMissionVideoCounter=Frame+70 ''VidList(19)="Corey_Attack_move4"
    Case 43 : CurrentVideo = 20  : Start_Video : showMissionVideoCounter=Frame+100  ''VidList(20)="Corey_Attack_move5"
    Case 44 : CurrentVideo = 21  : Start_Video : showMissionVideoCounter=Frame+80 ''VidList(21)="Corey_Attack_move6"
    Case 45 : CurrentVideo = 22  : Start_Video : showMissionVideoCounter=Frame+90 ''VidList(22)="Corey_Attack_move7"  MB start

'Tracy MB
    Case 46 : CurrentVideo = 45  : Start_Video : showMissionVideoCounter=Frame+235  'VidList(45)="Lago_pulls wires out off Tracy" 'MB ready
    Case 47 : CurrentVideo = 78  : Start_Video : showMissionVideoCounter=Frame+500  'VidList(78)="Tracy 360 shot" 'MB start
    Case 48 : CurrentVideo = 79  : Start_Video : showMissionVideoCounter=Frame+115  'VidList(79)="Tracy_arms out body" 'super JP
    Case 49 : CurrentVideo = 80  : Start_Video : showMissionVideoCounter=Frame+90 'VidList(80)="Tracy_Head Shaking"
    Case 50 : CurrentVideo = 81  : Start_Video : showMissionVideoCounter=Frame+90 'VidList(81)="Tracy_Robot Dies" 'MB end


' EOB BONUS X
    Case 51 : CurrentVideo = 91 : Start_Video : showMissionVideoCounter=Frame+133 ' 2x
    Case 52 : CurrentVideo = 92 : Start_Video : showMissionVideoCounter=Frame+133 ' 3x
    Case 53 : CurrentVideo = 93 : Start_Video : showMissionVideoCounter=Frame+133 ' 4x
    Case 54 : CurrentVideo = 94 : Start_Video : showMissionVideoCounter=Frame+133 ' 5x
    Case 55 : CurrentVideo = 95 : Start_Video : showMissionVideoCounter=Frame+133 ' 6x
    Case 56 : CurrentVideo = 96 : Start_Video : showMissionVideoCounter=Frame+133 ' 7x
    Case 57 : CurrentVideo = 97 : Start_Video : showMissionVideoCounter=Frame+133 ' 8x

    Case 60 : CurrentVideo = 37 : Start_Video : showMissionVideoCounter=Frame+300 'game over 2

    Case 61 : CurrentVideo = 71 : Start_Video : showMissionVideoCounter=Frame+180 ' gun target
'wizard videos
    Case 62 : CurrentVideo = 81  : Start_Video : showMissionVideoCounter=Frame+90 'VidList(81)="Tracy_Robot Dies" 'wizard start
    Case 63 : CurrentVideo = 79  : Start_Video : showMissionVideoCounter=Frame+115  'VidList(79)="Tracy_arms out body" ' tracy repaired
    Case 64 : CurrentVideo = 45  : Start_Video : showMissionVideoCounter=Frame+235  'VidList(45)="Lago_pulls wires out off Tracy" 'Wizard failed
    Case 65 : CurrentVideo = 76  : Start_Video : showMissionVideoCounter=Frame+310  'VidList(76)="Spaceship Warpspeed" 'Wizard phase 3 start
    Case 66 : CurrentVideo = 7  : Start_Video : showMissionVideoCounter=Frame+250 '"Corey putting on mask"  'PREPARE FOR GALDOR BATTLE
    Case 67 : CurrentVideo = 6  : Start_Video : showMissionVideoCounter=Frame+250 'VidList( 6)="Commander ship out hyperspace" 'START PHASE4
    Case 68 : CurrentVideo = 3  : Start_Video : showMissionVideoCounter=Frame+100 'VidList( 3)="Bodies fall on floor" Grabber magnet release
    Case 69 : CurrentVideo = 77 : Start_Video : showMissionVideoCounter=Frame+200 'Supernova appears.gif")        'wizard finish
    Case 70 : CurrentVideo = 53 : Start_Video : showMissionVideoCounter=Frame+200 'VidList(53)="Pyramid of Souls"     'more than 10 balls in grabber magnet
    Case 71 : CurrentVideo = 38 : Start_Video : showMissionVideoCounter=Frame+300 'VidList(38)="Game_Over_1"  'Wizard Completed while yell
'mima mb
        Case 72 : CurrentVideo = 50 : Start_Video : showMissionVideoCounter=Frame+188 ' Mima_Floating over camera super jp

'attract
    Case 73 : CurrentVideo = 98  : Start_Video : showMissionVideoCounter=Frame+1000 'VidList(98)="Attract_games_specs"
    Case 74 : CurrentVideo = 99  : Start_Video : showMissionVideoCounter=Frame+250  'VidList(99)="Attract_Push_Start"
    Case 75 : CurrentVideo = 100 : Start_Video : showMissionVideoCounter=Frame+700  'VidList(100)="Attract_flyer_text"
    Case 76 : CurrentVideo = 101 : Start_Video : showMissionVideoCounter=Frame+150  'VidList(101)="Attract_Bring_back_the_Mima"
    Case 77 : CurrentVideo = 102 : Start_Video : showMissionVideoCounter=Frame+110  'VidList(102)="DMD_Attract_EaShDa"
    Case 78 : CurrentVideo = 103 : Start_Video : showMissionVideoCounter=Frame+800  'VidList(103)="Attract_Intro_Credtis"
    Case 79 : CurrentVideo = 104 : Start_Video : showMissionVideoCounter=Frame+280  'VidList(104)="Attract_still_reading"
    Case 80 : CurrentVideo = 105 : Start_Video : showMissionVideoCounter=Frame+490  'VidList(105)="Bloodmachines logo"

  End Select

  showMissionVideo=0
  StartVideoFrame = 0

  FlexDMD.UnLockRenderThread

End Sub


'Sub DMD_FlipperInfo
' if PlayersPlayingGame = 1 Then Flipper_info = 0
' If Flipper_Info = 180 Then  'P1
'   DMD_ShowText "P1-" & Score(1) ,4,FontScoreInactiv2,28,False,40,1800
' End If
' If Flipper_Info = 300 Then  'P2
'   DMD_ShowText "P2-" & Score(2) ,4,FontScoreInactiv2,28,False,40,1800
' End If
' If Flipper_Info = 420 Then  'P3
'   If PlayersPlayingGame = 2 Then Flipper_info = 179
'   DMD_ShowText "P3-" & Score(3) ,4,FontScoreInactiv2,28,False,40,1800
' End If
' If Flipper_Info = 540 Then  'P4
'   If PlayersPlayingGame = 3 Then Flipper_info = 179
'   DMD_ShowText "P4-" & Score(4) ,4,FontScoreInactiv2,28,False,40,1800
' End If
' ' add more stuff here
' If Flipper_Info > 660 Then Flipper_Info = 179
'
'
'End Sub



Dim StartVideoFrame
Dim ShowBallsaver
Dim ShowBallLost
Dim TiltWarning
Dim DMD_StopOverlays : DMD_StopOverlays = False
Dim EOB_Bonus : EOB_Bonus = False       ' *EOB*

'Dim Flipper_Info

const BorderBlinkEnabled = 2

Sub MPCornerLabelsVisible(Enabled)
  if PlayersPlayingGame > 1 then
    FlexDMD.stage.GetLabel("Player1").Visible = Enabled
    FlexDMD.stage.GetLabel("Player2").Visible = Enabled
    FlexDMD.stage.GetLabel("Player3").Visible = Enabled
    FlexDMD.stage.GetLabel("Player4").Visible = Enabled
  end if
end sub

Sub DMDTimer_Timer

' If LeftF = 1 Or RightF=1 Then Flipper_Info = Flipper_Info + 1 Else Flipper_Info = 0

  Dim i, n, x, y, label, xx
  Frame = Frame + 1

  FlexDMD.LockRenderThread

  Select Case DmdMode
    Case 1
      If InsertCoins>0 Then DMD_Insert_Coins : Exit Sub
      If EOB_Bonus Then DMD_EOB_Bonus       ' *EOB*
      If hsbModeActive Then DMD_EnterHigh
      If bAttractMode Then DMD_Attract
      If bEndCredits Then DMD_ending_credits
      'If Flipper_info > 179 Then DMD_FlipperInfo

      DMD_OverlayText



      If FlexOverlay>0 Then
        DMD_Overlay

      Else ' do everything Else
        If Borderblink>0 Then
          if BorderBlinkEnabled = 1 then
            DMDBorderBlink1 ' can always run ?!
          Else
            DMDScoreBlink
          end if
        end if


        If ShowSkillshot>0 Then SkillshotScreen : Exit Sub
'       If videoTest Then DMDShowVideos2 : Exit Sub
        If showMissionVideo>0 And CurrentVideo=0 Then DMDShowVideos : Exit Sub
        If showMissionVideo=60 Then DMDShowVideos : Exit Sub
        If showMissionVideo=62 Then DMDShowVideos : Exit Sub
        If CurrentVideo>0 And showMissionVideoCounter<Frame Then StopVideosPlaying
        If ShowBallLost>0 Then DMD_Ball_Lost : Exit Sub
        If TiltWarning>0 Then DMD_Tilt_Warning : Exit Sub
        If Tilted And GaldorDefeated = 0 And Not Slammed Then DMD_Tilt : Exit Sub
        If Slammed Then DMD_Slam : Exit Sub


        Select Case CurrentMission

          Case 0: ' NO MISSION ****************************

            if Not bwizardmode then
              FlexDMD.Stage.GetImage("highlight1").Visible = False
              FlexDMD.Stage.GetLabel("Splash").Visible = False
              FlexDMD.Stage.GetLabel("Splash2").Visible = False
            end if

            If ShowBallsaver>0 Then DMD_BALL_SAVER : Exit Sub

            If PlayersPlayingGame = 1 Then
              FlexDMD.stage.GetLabel("Score_1").Visible = True
              FlexDMD.stage.GetLabel("Player1").Visible = False
            end if

            if Not bwizardmode then
              Select Case CurrentVideo
                Case  1 : DMD_Mission1_Failed
                Case 18 : DMD_CoreyMB_Super_Jackpot 'msgbox "super jackpot!"
                Case 29 : DMD_Mission2_Failed
                Case 40 : DMD_Mission6_Complete
                Case 46 : DMD_Mission6_Failed
                Case 47 : DMD_Mission3_Complete
                Case 49 : DMD_Mission3_Failed
                case 50 : if cLightMimaMB.state = 2 then DMD_MimaMB_Super_Jackpot
                Case 52 : DMD_Mission2_Complete
                Case 56 : DMD_Mission4_Failed
                Case 58 : DMD_Mission4_Complete
                Case 69 : DMD_Mission7_Failed
                Case 71 : DMD_GunTarget_Complete
                Case 77 : DMD_Mission7_Complete
                Case 79 : if cLightTracyMB.state = 2 then DMD_TracyMB_Super_Jackpot
                Case 83 : DMD_Mission5_Complete
                Case 84 : DMD_Mission5_Failed
                Case 86 : DMD_Mission1_Complete

                Case Else :

                  'by priority ?
                  If BG_Bump1>0 Or BG_Bump2>0 Or BumpCounter>0 Then
                    DMDStarflash   ' bumpers
                  Else
                    '* score Black overlay   **************************************************
                    OverlayBigScore ' every Frame
                    If (Frame Mod 3) = 0 Then  ' update only ever 3 frames *17ms
                      FlexDMD.stage.GetLabel("Score_1").Visible = True          'make score_1 visible in MP ????
                      If DMDMystery>0 And DMDMystery<52 Then ' by priority
                        FlexDMD.Stage.GetLabel("Score_1").Text = "MYSTERY"
                      Elseif DMDExtraBall>0 Then
                        DMDExtraBall=DMDExtraBall+1
                        If DMDExtraBall < 22 Then
                          FlexDMD.Stage.GetLabel("Score_1").Text = Left("EXTRA BALL",(DMDExtraBall-1)/2)
                        Else
                          If RndInt(1,2)=1 Then
                            FlexDMD.Stage.GetLabel("Score_1").Text = " "
                          Else
                            FlexDMD.Stage.GetLabel("Score_1").Text ="EXTRA BALL"
                          End If
                        End If
                        If DMDExtraBall>50 Then DMDExtraBall=0
                      Else ' normal big score
                        If PlayersPlayingGame > 1 Then
  '                       FlexDMD.Stage.GetLabel("Score_1").Text = " "
                        Else
                          FlexDMD.Stage.GetLabel("Score_1").Text = FormatScore(Score(CurrentPlayer))
                        end if
                      End If
                      if PlayersPlayingGame = 1 Then
                        FlexDMD.Stage.GetLabel("Score_1").Font = FontScoreActive
                      Else
                        FlexDMD.Stage.GetLabel("Score_1").Font = Font7by13W
                      end if
                      FlexDMD.Stage.GetLabel("Score_1").SetAlignedPosition 65,18, FlexDMD_Align_Center '13

                    End If
                  End If
              End Select
            End If
            '************* Wizard DMD *************
            if bWizardMode Then
              FlexDMD.stage.GetLabel("Score_1").Visible = False
              MPCornerLabelsVisible false

              Select Case CurrentVideo
                Case  79 : DMD_Wizard_Phase2_Completed
                Case  45 : if cLightTracyMB.state = 0 then DMD_Wizard_Failed  'hack to fix dmd when last played mode has been tracy mb
                Case  7 : DMD_Wizard_Phase3_Completed
                Case 77 : if Not WizardPhase < 2 then DMD_Wizard_Phase4_Completed
                Case 53 : 'debug.print "pyramid"
'             End Select
                Case else:
                select Case WizardPhase
                  Case 1:
'                   If CurrentVideo = 62 then
'                     DMD_Wizard_Phase1
'                   end If
                    FlexDMD.Stage.GetLabel("Splash2").Text = "Final Chapter"
                    FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
                    FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 66,17, FlexDMD_Align_Center
                    FlexDMD.Stage.GetLabel("Splash2").Visible = True
                    FlexDMD.Stage.GetImage("highlight1").Visible = False
                  Case 2:
                    if WizardOrbShotCount <> 0 then DMD_Wizard_Phase2
                  Case 3:
                    If CurrentVideo = 76 then
                      DMD_Wizard_Phase3_Startup
                    ElseIF DMDMissionStarted=0 Then
                      DMD_Wizard_Phase3_Setup
                    Else
                      If Not cLightUnderPFShot2.state = 1 then DMD_Wizard_Phase3
                      'to do WizardPhase3DrainCount info
                    end if
                  Case 4:
                    If CurrentVideo = 6 then 'opening video vid 6
                      DMD_Wizard_Phase4_Startup
                    ElseIF DMDMissionStarted=0 Then'setup labels
                      DMD_Wizard_Phase4_Setup
                    Elseif CurrentVideo = 3 And WizardSoulLost then
                      DMD_Wizard_Phase4_LostSoulOnVideo
                    Elseif WizardSoulLost then
                      DMD_Wizard_Phase4_LostSoul
                    Else  'before finish updates
                      DMD_Wizard_Phase4
                    end if
                  Case 5:
                    if CurrentVideo = 38 Then
                      DMD_Wizard_Phase5_Startup
                    else
                      'DMD_ending_credits
                    end if
                  Case Else:
                    FlexDMD.Stage.GetImage("highlight1").Visible = False
                    FlexDMD.Stage.GetLabel("Splash").Visible = False
                    FlexDMD.Stage.GetLabel("Splash2").Visible = False
                    FlexDMD.Stage.GetLabel("Splash").text = ""
                    FlexDMD.Stage.GetLabel("Splash2").text = ""
                end Select
              end Select
            end If


          Case 1 : ' MISSION 1 ****************************
            FlexDMD.stage.GetLabel("Score_1").Visible = False
            MPCornerLabelsVisible false

            If CurrentVideo = 82 then
              DMD_Mission1_Startup ' intro vid mission1
            Elseif CurrentVideo = 30 then
              DMD_Mission1_SHOOT
            Elseif CurrentVideo = 70 then
              DMD_Mission1_Shooting

            ElseIf DMDMissionStarted=0 Then
              DMD_Mission1_Setup 'normal mission1 setup
            Else
            ' normal scoring .. no numbers
              If IsEmpty(BallInGun) Then
                'FlexDMD.Stage.GetLabel("Splash2").Text = StartMissionName2
'             Else
                'FlexDMD.Stage.GetLabel("Splash2").Text = "GET " & 3 - TotalBladerunnerHits & " AI SHIPS"
                if TotalBladerunnerHits = 0 then
                  FlexDMD.Stage.GetLabel("Splash2").Text = 3 - TotalBladerunnerHits & " SHIPS TO GO"
                ElseIF TotalBladerunnerHits = 1 then
                  FlexDMD.Stage.GetLabel("Splash2").Text = 3 - TotalBladerunnerHits & " SHIPS MORE"
                ElseIF TotalBladerunnerHits = 2 then
                  FlexDMD.Stage.GetLabel("Splash2").Text = 3 - TotalBladerunnerHits & " SHIP LEFT"
                end if
              Else
                FlexDMD.Stage.GetLabel("Splash2").Text = "HIT AI SHIP"
              end if
              FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
              FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
            End If



          Case 2 : ' MISSION 2 ****************************
            FlexDMD.stage.GetLabel("Score_1").Visible = False
            MPCornerLabelsVisible false

            If CurrentVideo = 27 then
              DMD_Mission2_Startup ' intro vid mission2
            ElseIf DMDMissionStarted=0 Then
              DMD_Mission2_Setup 'normal
            Else
              If TotalMimaLoops<25 Then
                FlexDMD.Stage.GetLabel("Splash2").Text = "GET " & 25 - TotalMimaLoops & " LOOPS"
                FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
                FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
                FlexDMD.Stage.GetLabel("Splash2").Visible = True
              End If
            End If


          Case 3 :' MISSION 3 ****************************
            FlexDMD.stage.GetLabel("Score_1").Visible = False
            If CurrentVideo = 76 then
              DMD_Mission3_Startup ' intro vid mission2
            ElseIf DMDMissionStarted=0 Then
              DMD_Mission3_Setup 'normal
            Else
              FlexDMD.Stage.GetLabel("Splash2").Text = "NEED " & 4-TotalChaseHits & " SHOTS"
              FlexDMD.stage.GetLabel("Splash2").Font= FontWhite3
              FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
              FlexDMD.Stage.GetLabel("Splash2").Visible = True
            End If




          Case 4 :' MISSION 4 ****************************
            FlexDMD.stage.GetLabel("Score_1").Visible = False
            MPCornerLabelsVisible false
            If CurrentVideo = 59 then
              DMD_Mission4_Startup ' intro vid mission2
            ElseIf DMDMissionStarted=0 Then
              DMD_Mission4_Setup 'normal
            Else

              If TotalCemeteryHits=1 Then FlexDMD.Stage.GetLabel("Splash2").Text = "3 SHOTS LEFT"
              If TotalCemeteryHits=2 Then FlexDMD.Stage.GetLabel("Splash2").Text = "2 SHOTS LEFT"
              If TotalCemeteryHits=3 Then FlexDMD.Stage.GetLabel("Splash2").Text = "ONE MORE HIT"
              If TotalCemeteryHits=4 Then FlexDMD.Stage.GetLabel("Splash2").Text = "MISSION DONE"
              FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
              FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
              FlexDMD.Stage.GetLabel("Splash2").Visible = True

            End If


          Case 5 :' MISSION 5 ****************************
            FlexDMD.stage.GetLabel("Score_1").Visible = False
            MPCornerLabelsVisible false
            If CurrentVideo = 90 then
              DMD_Mission5_Startup ' intro vid mission2
            ElseIf DMDMissionStarted=0 Then
              DMD_Mission5_Setup 'normal
            End If


          Case 6 :' MISSION 6 ****************************
            FlexDMD.stage.GetLabel("Score_1").Visible = False
            MPCornerLabelsVisible false
            If CurrentVideo = 41 then
              DMD_Mission6_Startup ' intro vid mission2
            ElseIf DMDMissionStarted=0 Then
              DMD_Mission6_Setup 'normal
            Else
              FlexDMD.Stage.GetLabel("Splash2").Text = "GET " & 25-TotalHeartOfSteelHits & " BUMPERS"
              FlexDMD.stage.GetLabel("Splash2").Font=Font6by10W
              FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
              FlexDMD.Stage.GetLabel("Splash2").Visible = True
            End If



          Case 7 :' MISSION 7 ****************************
            FlexDMD.stage.GetLabel("Score_1").Visible = False
            MPCornerLabelsVisible false
            If CurrentVideo = 78 then
              DMD_Mission7_Startup ' intro vid mission2
            ElseIf DMDMissionStarted=0 Then
              DMD_Mission7_Setup 'normal
            End If


        End Select
  ' commons !
        DMD_BALL_IN_PLAY

        If CreditBlink>0 Then
          If playerblink=0 Then
            CreditBlink=CreditBlink-1
            If CreditBlink=0 Then
              FlexDMD.Stage.GetLabel("Splash3").Font = FontScoreInactiv1
              FlexDMD.Stage.GetLabel("Splash3").Text = " "
              ShowAddScoreText=" "
              ShowAddScore=0
            Else
              If ShowAddScore>0 Then
                FlexDMD.Stage.GetLabel("Splash3").Text = ShowAddScoreText
                FlexDMD.Stage.GetLabel("Splash3").SetAlignedPosition 64,30, FlexDMD_Align_Center
                FlexDMD.Stage.GetLabel("Splash3").Font = FontScoreInactiv2
                FlexDMD.Stage.GetLabel("Splash3").Visible = True
              Else
                FlexDMD.Stage.GetLabel("Splash3").Text = "CREDITS " & Credits
                FlexDMD.Stage.GetLabel("Splash3").SetAlignedPosition 64,30, FlexDMD_Align_Center
                FlexDMD.Stage.GetLabel("Splash3").Font = FontScoreInactiv2
                FlexDMD.Stage.GetLabel("Splash3").Visible = True
              End If
            End If
          Else
            FlexDMD.Stage.GetLabel("Splash3").Font = FontScoreInactiv1
          End If
        End if

        if (Frame Mod 3) = 1 Then
          If Not BG_Stars=0 Then DMDaddstars
          If BG_Star1>0 Then DMDstar1 ' starfalls
          If BG_Star2>0 Then DMDstar2

          If BG_BarLeft>0 Then DMDmovebar1
          If BG_BarRight<0 Then DMDmovebar2

          ' splashline goes under here, priority from top, next one will come when previous are done
          If DMDmystery>0 Then
            DMDMysteryDisplay
          End If
        End If
        'Elseif (Frame Mod 3) = 2 Then
        '    divide everything up into 3 parts ( every 3 part is only updated every 3rd frame

      End If


    Case 2: ' intromovie

      If frame = 30 Then
        CreateDMD1
        if ScorbitActive = 0 then Introover=True  ' This is here to enable the start immediately. Delayed with scorbit
      end if

      If frame = 330 Then FlexDMD.Stage.GetImage("Border2").visible=True
      If frame = 345 Then FlexDMD.Stage.GetImage("Border3").visible=True
      If frame = 350 Then FlexDMD.Stage.GetImage("Border3").visible=False
      If frame = 355 Then FlexDMD.Stage.GetImage("Border3").visible=True
      If frame = 360 Then FlexDMD.Stage.GetImage("Border3").visible=False
      If frame = 365 Then FlexDMD.Stage.GetImage("Border3").visible=True
      If frame = 375 Then FlexDMD.Stage.GetImage("Border3").visible=False
      If frame = 380 Then FlexDMD.Stage.GetImage("Border2").visible=False
      If frame = 488 Then CreateDMD1 : Introover=True

  End Select



  FlexDMD.UnlockRenderThread

End Sub


Dim DMDMissionStarted
'*************************************************
'*****   Wizard DMD
'*************************************************

sub DMD_Wizard_Phase1
' msgbox "phase1 dmd"
end sub


sub DMD_Wizard_Phase2
  'msgbox "phase2 dmd"
  If DMDMissionStarted=0 Then
    DMD_Wizard_Phase2_Setup 'normal mission1 setup
  Else
  ' normal scoring .. no numbers
    if WizardOrbShotCount > 1 then
      FlexDMD.Stage.GetLabel("Splash2").Text = "HIT " & WizardOrbShotCount & " ORBS"
    ElseIF WizardOrbShotCount = 1 then
      FlexDMD.Stage.GetLabel("Splash2").Text = "HIT " & WizardOrbShotCount & " ORB"
    end if
    FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
    FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  End If
end sub

sub DMD_Wizard_Phase2_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "TRACY BROKEN"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 65,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  FlexDMD.Stage.GetLabel("Splash2").Text = "SHOOT 7 ORBS"
  FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True
end sub

sub DMD_Wizard_Phase2_Completed
' DMDMissionStarted = 0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=3 Then
    'debug.print "phase2 videotext: " & StartVideoFrame
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "TRACY ALIVE"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,17, FlexDMD_Align_Center
    FlexDMD.Stage.GetLabel("Splash").Text = ""
    FlexDMD.Stage.GetLabel("Splash2").Text = ""
  End If
end Sub


sub DMD_Wizard_Phase3
' msgbox "phase3 dmd"

  ' normal scoring .. no numbers
' debug.print "normal: " & WizardOrbShotCount
  if Not cLightUnderPFShot1.state = 1 then
    FlexDMD.Stage.GetLabel("Splash2").Text = "SHOOT LEFT"
  Else
    FlexDMD.Stage.GetLabel("Splash2").Text = "SHOOT RIGHT"
  end if
  FlexDMD.stage.GetLabel("Splash2").Font=Font6by10W
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
end sub

sub DMD_Wizard_Phase3_Startup
  DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=1 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = "FREE THE SOULS"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,17, FlexDMD_Align_Center
    FlexDMD.Stage.getlabel("videotext").Visible = True
  End If
end sub

sub DMD_Wizard_Phase3_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "RELEASE SOULS"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 65,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  FlexDMD.Stage.GetLabel("Splash2").Text = "FLASHING SHOTS"
  FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True
end sub

sub DMD_Wizard_Phase3_Completed
' DMDMissionStarted = 0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = "PREPARE FOR GALDOR"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,22, FlexDMD_Align_Center
    FlexDMD.Stage.GetLabel("Splash").Text = ""
    FlexDMD.Stage.GetLabel("Splash2").Text = ""
  End If
end Sub




sub DMD_Wizard_Phase4_Startup
  DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=1 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "BATTLE GALDOR"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,17, FlexDMD_Align_Center
    FlexDMD.Stage.getlabel("videotext").Visible = True
  End If
end sub


sub DMD_Wizard_Phase4_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "BATTLE GALDOR"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 65,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  FlexDMD.Stage.GetLabel("Splash2").Text = "GET " & 5 & " HITS"
  FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True
end sub


sub DMD_Wizard_Phase4
  If GaldorShotCount > 1 then
    FlexDMD.Stage.GetLabel("Splash2").Text = "GET " & GaldorShotCount & " HITS"
    FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
  Elseif GaldorShotCount = 1 then
    FlexDMD.Stage.GetLabel("Splash").Text = "FINAL SHOT"
    FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
    FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 65,7, FlexDMD_Align_Center
    FlexDMD.Stage.GetLabel("Splash2").Text = "HIT THE SKYLOOP"
    FlexDMD.stage.GetLabel("Splash2").Font=Font6by10W
  end If
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
end sub

dim LostSoulFrames : LostSoulFrames = 0
sub DMD_Wizard_Phase4_LostSoul
  LostSoulFrames = LostSoulFrames + 1
  FlexDMD.Stage.GetLabel("Splash2").Text = BallsOnPlayfield & " SOULS LEFT"
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.stage.GetLabel("Splash2").Font=Font6by10W
  if LostSoulFrames > 50 Then
    LostSoulFrames = 0
    WizardSoulLost = False
  end if
end sub

sub DMD_Wizard_Phase4_LostSoulOnVideo
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=1 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = BallsOnPlayfield & " SOULS LEFT"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,17, FlexDMD_Align_Center
    FlexDMD.Stage.getlabel("videotext").Visible = True
  End If

  if StartVideoFrame > 50 Then
    WizardSoulLost = False
  end if
end sub


sub DMD_Wizard_Failed
' DMDMissionStarted = 0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "WIZARD FAILED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,22, FlexDMD_Align_Center
  End If
end sub


sub DMD_Wizard_Phase4_Completed
' DMDMissionStarted = 0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = "GALDOR DEFEATED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,22, FlexDMD_Align_Center
    FlexDMD.Stage.GetLabel("Splash").Text = ""
    FlexDMD.Stage.GetLabel("Splash2").Text = ""
    FlexDMD.Stage.GetImage("highlight1").Visible = False
  End If
end Sub

sub DMD_Wizard_Phase5_Startup
' DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = "GALDOR DEFEATED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,22, FlexDMD_Align_Center
    FlexDMD.Stage.GetLabel("Splash").Text = ""
    FlexDMD.Stage.GetLabel("Splash2").Text = ""
    FlexDMD.Stage.GetImage("highlight1").Visible = False
  End If
end Sub

'*************************************************
'*****   MISSION 7
'*************************************************

Sub DMD_Mission7_Failed ' after mission so :....:
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "FAILED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission7_Complete
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = ShowAddScoreText ' "COMPLETED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission7_Startup

  DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=10 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "BETRAYAL"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If

End Sub


Sub DMD_Mission7_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "BETRAYAL"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  FlexDMD.Stage.GetLabel("Splash2").Text = "SHOT COMBO"
  FlexDMD.stage.GetLabel("Splash2").Font= FontWhite3
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True

End Sub






'*************************************************
'*****   MISSION 6
'*************************************************

Sub DMD_Mission6_Failed ' after mission so :....:
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "FAILED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission6_Complete
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = ShowAddScoreText ' "COMPLETED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission6_Startup

  DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=10 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "HEART OF STEEL"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If

End Sub


Sub DMD_Mission6_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "HEARTOFSTEEL"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  FlexDMD.Stage.GetLabel("Splash2").Text = "GET " & 25-TotalHeartOfSteelHits & " BUMPERS"
  FlexDMD.stage.GetLabel("Splash2").Font= Font6by10W
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True

End Sub




'*************************************************
'*****   MISSION 5
'*************************************************


Sub DMD_Mission5_Failed ' after mission so :....:
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "FAILED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission5_Complete
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = ShowAddScoreText ' "COMPLETED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission5_Startup

  DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=10 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "PORTAL ROOM"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If

End Sub


Sub DMD_Mission5_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "PORTAL ROOM"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  FlexDMD.Stage.GetLabel("Splash2").Text = "VASCAN TARGETS"
  FlexDMD.stage.GetLabel("Splash2").Font= Font6by10W
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True

End Sub



'*************************************************
'*****   MISSION 4
'*************************************************

Sub DMD_Mission4_Failed ' after mission so :....:
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "FAILED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission4_Complete
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = ShowAddScoreText ' "COMPLETED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission4_Startup

  DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=10 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "CEMETERY"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If

End Sub


Sub DMD_Mission4_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "CEMETERY"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  FlexDMD.Stage.GetLabel("Splash2").Text = "NEED 4 SHOTS"
  FlexDMD.stage.GetLabel("Splash2").Font= FontWhite3
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True

End Sub




'*************************************************
'*****   MISSION 3
'*************************************************

Sub DMD_Mission3_Failed ' after mission so :....:
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "FAILED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission3_Complete
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = ShowAddScoreText ' "COMPLETED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission3_Startup

  DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=10 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "C H A S E"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,22, FlexDMD_Align_Center
  End If

End Sub


Sub DMD_Mission3_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "C H A S E"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  FlexDMD.Stage.GetLabel("Splash2").Text = "NEED 4 SHOTS"
  FlexDMD.stage.GetLabel("Splash2").Font= FontWhite3
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True

End Sub




'*************************************************
'*****   MISSION 2
'*************************************************

Sub DMD_Mission2_Failed ' after mission so :....:
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "FAILED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission2_Complete
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = ShowAddScoreText ' "COMPLETED"
    FlexDMD.Stage.getlabel("videotext").Visible = True
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,12, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission2_Startup

  DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=10 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "RESURRECTION"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,22, FlexDMD_Align_Center
  End If

End Sub


Sub DMD_Mission2_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "RESURRECTION"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  FlexDMD.Stage.GetLabel("Splash2").Text = "GET 25 LOOPS"
  FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True

End Sub




'*************************************************
'*****   MISSION 1
'*************************************************
Sub DMD_Mission1_Failed ' after mission so :....:
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "FAILED"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,7, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission1_Complete
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = ShowAddScoreText ' "COMPLETED"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,7, FlexDMD_Align_Center
  End If
End Sub


Sub DMD_Mission1_Shooting
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=10 Then
    StartMissionName2="LOAD THE GUN"
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = StartMissionName2
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,7, FlexDMD_Align_Center

  End If
End Sub


Sub DMD_Mission1_SHOOT
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=10 Then
    StartMissionName2="SHOOT AI SHIP"
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = StartMissionName2
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,7, FlexDMD_Align_Center
  End If

End Sub

Sub DMD_Mission1_Setup
  DMDMissionStarted=1
  'FontWhite3
  FlexDMD.Stage.GetLabel("Splash").Text = "THE HUNT"
  FlexDMD.stage.GetLabel("Splash").Font=FontBlack3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,7, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True

  StartMissionName2="SHOOT AI SHIP"
  FlexDMD.Stage.GetLabel("Splash2").Text = StartMissionName2
  FlexDMD.stage.GetLabel("Splash2").Font=FontWhite3
  FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,18, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash2").Visible = True
  FlexDMD.Stage.GetImage("highlight1").SetPosition 0,-1
  FlexDMD.Stage.GetImage("highlight1").Visible = True

End Sub

Sub DMD_Mission1_Startup

  DMDMissionStarted=0
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=10 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = "THE HUNT"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,7, FlexDMD_Align_Center
  End If

  If StartVideoFrame=100 Then
    FlexDMD.stage.GetLabel("videotext").font=FontBlack3
    FlexDMD.Stage.getlabel("videotext").Text = "THE HUNT"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,7, FlexDMD_Align_Center
  End If
End Sub


'*************************************************
'*****   COREY MB
'*************************************************

sub DMD_CoreyMB_Super_Jackpot
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = "SUPER JACKPOT"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If
  If StartVideoFrame=60 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = FormatScore(score_TracyMBSuperJackpot)
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If
end sub

'*************************************************
'*****   Mima MB
'*************************************************

sub DMD_MimaMB_Super_Jackpot
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = "SUPER JACKPOT"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If
  If StartVideoFrame=60 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = FormatScore(score_MimaMBSuperJackpot)
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If
end sub

'*************************************************
'*****   Tracy MB
'*************************************************

sub DMD_TracyMB_Super_Jackpot
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = "SUPER JACKPOT"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If
  If StartVideoFrame=60 Then
    FlexDMD.stage.GetLabel("videotext").font=Font6by10W
    FlexDMD.Stage.getlabel("videotext").Text = FormatScore(score_TracyMBSuperJackpot)
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,27, FlexDMD_Align_Center
  End If
end sub


'*************************************************
'*****   Gun Target
'*************************************************

Sub DMD_GunTarget_Complete
  StartVideoFrame=StartVideoFrame+1
  If StartVideoFrame=2 Then
    FlexDMD.stage.GetLabel("videotext").font=FontBlack3
    FlexDMD.Stage.getlabel("videotext").Text = "NICE SHOT" ' "COMPLETED"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,17, FlexDMD_Align_Center
  End If
  If StartVideoFrame=60 Then
    FlexDMD.stage.GetLabel("videotext").font=FontWhite3
    FlexDMD.Stage.getlabel("videotext").Text = ShowAddScoreText ' "COMPLETED"
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,17, FlexDMD_Align_Center
  End If
End Sub

'*************************************************
'*****   MISSION ENDS
'*************************************************



Sub DMD_BALL_IN_PLAY
  Dim i,label
  '* ballinplay   **************************************************


  For i = 1 to 4 ' player P1 P2 P3 P4 blinkers
    Set label = FlexDMD.Stage.GetLabel("Player" & i)
'   label.Text = "P"&i
    label.Text = FormatScore(Score(i))

    If i > PlayersPlayingGame Then
      label.visible = False
    Else
      if PlayersPlayingGame > 1 then
        if CurrentMission = 0 And Not bwizardmode then
          label.visible = True
        end if
        'move labels to align properly to current score
        If i = CurrentPlayer Then
          'label.Font = Font6by10W                    'todo: this font change needs to be moved
          Select Case i
            case 2:
              label.SetAlignedPosition 128,6, FlexDMD_Align_Right
            case 3:
              label.SetAlignedPosition 0,28, FlexDMD_Align_Left
            case 4:
              label.SetAlignedPosition 128,28, FlexDMD_Align_Right
          end select
        Else
          'label.Font = FontScoreInactiv1
          Select Case i
            case 2:
              label.SetAlignedPosition 128,3, FlexDMD_Align_Right
            case 3:
              label.SetAlignedPosition 0,30, FlexDMD_Align_Left
            case 4:
              label.SetAlignedPosition 128,30, FlexDMD_Align_Right
          end select
        End If
      end if
    end if

  Next

End Sub

sub DMD_Tilt_Warning
  If TiltWarning=1 Then StopVideosPlaying
  FlexDMD.stage.GetLabel("Score_1").Visible = false
  FlexDMD.stage.GetLabel("Splash2").Visible = false
  FlexDMD.Stage.GetImage("highlight1").Visible = false
  MPCornerLabelsVisible false

  FlexDMD.Stage.GetLabel("Splash").Text = "WARNING"
  FlexDMD.stage.GetLabel("Splash").Font=FontWhite3
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,15, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True
  FlexDMD.unLockRenderThread
end sub

sub DMD_Tilt
  If TiltWarning=1 Then StopVideosPlaying
  FlexDMD.stage.GetLabel("Score_1").Visible = false
  FlexDMD.stage.GetLabel("Splash2").Visible = false
  FlexDMD.Stage.GetImage("highlight1").Visible = false
  MPCornerLabelsVisible false


  FlexDMD.Stage.GetLabel("Splash").Text = "TILT"
  FlexDMD.stage.GetLabel("Splash").Font=FontScoreActive
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,15, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True
  FlexDMD.unLockRenderThread
end sub

sub DMD_Slam
  If TiltWarning=1 Then StopVideosPlaying
  FlexDMD.stage.GetLabel("Score_1").Visible = false
  FlexDMD.stage.GetLabel("Splash2").Visible = false
  FlexDMD.Stage.GetImage("highlight1").Visible = false
  MPCornerLabelsVisible false


  FlexDMD.Stage.GetLabel("Splash").Text = "SLAM TILT"
  FlexDMD.stage.GetLabel("Splash").Font=FontScoreActive
  FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,15, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash").Visible = True
  FlexDMD.unLockRenderThread
end sub

Sub DMD_BALL_LOST
  If ShowBallLost=1 Then StopVideosPlaying

  ShowBallLost=ShowBallLost+1

  Select Case ShowBallLost
    Case 2,42,82
      FlexDMD.Stage.GetImage("balllost1").Visible=True  : FlexDMD.Stage.GetImage("balllost2").Visible=False
    Case 22,62,102
      FlexDMD.Stage.GetImage("balllost1").Visible=False : FlexDMD.Stage.GetImage("balllost2").Visible=True
    Case 120 : ' end it at 2 seconds
      FlexDMD.Stage.GetImage("balllost2").Visible=False
      ShowBallLost=0
  End Select
  FlexDMD.unLockRenderThread
End Sub


Sub DMD_BALL_SAVER
  If ShowBallsaver=1 Then StopVideosPlaying

  ShowBallsaver=ShowBallsaver+1

  Select Case ShowBallsaver
    Case 2,42,82
      FlexDMD.Stage.GetImage("ballsave1").Visible=True  : FlexDMD.Stage.GetImage("ballsave2").Visible=False
    Case 22,62,102
      FlexDMD.Stage.GetImage("ballsave1").Visible=False : FlexDMD.Stage.GetImage("ballsave2").Visible=True
    Case 120 : ' end it at 2 seconds
      FlexDMD.Stage.GetImage("ballsave2").Visible=False
      ShowBallsaver=0
  End Select
  FlexDMD.unLockRenderThread
End Sub




Sub DMD_Insert_Coins
  InsertCoins=InsertCoins+1
  If InsertCoins > 119 Then InsertCoins = 0 : FlexDMD.Stage.GetImage("insertcoin2").Visible=False
  Select Case InsertCoins
    Case 2,42,82
      FlexDMD.Stage.GetImage("insertcoin1").Visible=True  : FlexDMD.Stage.GetImage("insertcoin2").Visible=False
    Case 22,62,102
      FlexDMD.Stage.GetImage("insertcoin1").Visible=False : FlexDMD.Stage.GetImage("insertcoin2").Visible=True
  End Select
  FlexDMD.unLockRenderThread
End Sub




Sub SkillshotScreen

  If ( frame mod 2 ) = 1 Then
    ShowSkillshot=ShowSkillshot-1
    Select Case ShowSkillshot
      Case 99:
        FlexDMD.Stage.GetImage("skillshot").Visible = True
        FlexDMD.Stage.GetLabel("SplashScreen1").Text = "SKILLSHOT"
        FlexDMD.Stage.GetLabel("SplashScreen1").Font = Font6by10W
        FlexDMD.Stage.GetLabel("SplashScreen1").SetAlignedPosition 7,11, FlexDMD_Align_Left
        FlexDMD.Stage.GetLabel("SplashScreen1").Visible = True


        FlexDMD.Stage.GetLabel("SplashScreen2").Text = ShowSkillshot2
        FlexDMD.Stage.GetLabel("SplashScreen2").Font = Font6by10W
        FlexDMD.Stage.GetLabel("SplashScreen2").SetAlignedPosition 7,24, FlexDMD_Align_Left
        FlexDMD.Stage.GetLabel("SplashScreen2").Visible = True

      case 90,80,70,60,50 :
        FlexDMD.Stage.GetLabel("SplashScreen1").Font = Font6by10W
      case 95,85,75,65,55 :
        FlexDMD.Stage.GetLabel("SplashScreen1").Font = Font6by10B
      case 33 :
        Borderblink=20
      Case 3:
        FlexDMD.Stage.GetImage("skillshot").Visible = False
        FlexDMD.Stage.GetLabel("SplashScreen1").Visible = False
        FlexDMD.Stage.GetLabel("SplashScreen2").Visible = False
    End Select
  End If

  FlexDMD.UnlockRenderThread

End Sub

Dim BumpCounter : BumpCounter=0
Sub DMDStarflash

' debug.print "CV"&currentvideo & "  BC" & BumpCounter
  If ( Frame mod 4 ) = 1 Then
      FlexDMD.stage.GetLabel("videotext").font=FontWhite3
  Elseif ( Frame mod 8 ) = 5 Then
      FlexDMD.stage.GetLabel("videotext").font=FontBlack3
  End If

  If CurrentVideo = 0 Then
    BumpCounter = 110 : BG_Bump1=0 : BG_Bump2 = 0
    CurrentVideo = 80
    showMissionVideoCounter=Frame+10000
    Start_Video

    If SuperJets>50 Then
      FlexDMD.Stage.getlabel("videotext").Text = FormatScore(score_Bumper)
    Else
      FlexDMD.Stage.getlabel("videotext").Text = FormatScore(score_SuperJetBumper)
    End If
    FlexDMD.Stage.getlabel("videotext").Font = FontWhite3
    FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 4+RndInt(45,1),5+RndInt(22,1), FlexDMD_Align_Left


  ElseIf currentvideo = 80 Then

    BumpCounter=BumpCounter-1
    If BG_Bump1>0 Or BG_Bump2>0 Then
      BumpCounter = 110 : BG_Bump1=0 : BG_Bump2 = 0
      If SuperJets < 50 Then
        FlexDMD.Stage.getlabel("videotext").Text = FormatScore(score_Bumper)
      Else
        FlexDMD.Stage.getlabel("videotext").Text = FormatScore(score_SuperJetBumper)
      End If
      FlexDMD.Stage.getlabel("videotext").Font = FontWhite3
      FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 4+RndInt(45,1),5+RndInt(22,1), FlexDMD_Align_Left
    End If

    If BumpCounter = 50 Then
      Dim tmptext : tmptext=""
      If ShowSuperJets=1 Then
        tmptext="ACTIVATED"
      ElseIf Show_TracyHits > 0 And bumpersplash > 0 Then
        If (15-Tracy_bumpers) < (50-SuperJets ) And cLightTracyMB.state = 0 Then
          tmptext= 15 - Tracy_bumpers & " FOR TRACY"
        Else
          tmptext= 50 - SuperJets & " SUPERJETS"
        End If
      Elseif BumperSplash > 0 And cLightTracyMB.state = 0 Then
        tmptext= 50 - SuperJets & " SUPERJETS"
      Else
        tmptext= 15 - Tracy_bumpers & " FOR TRACY"
      End If
      Show_TracyHits=0
      bumpersplash=0

      FlexDMD.Stage.getlabel("videotext").Text = tmptext
      FlexDMD.Stage.getlabel("videotext").Font = FontWhite3
      FlexDMD.Stage.getlabel("videotext").SetAlignedPosition 64,20, FlexDMD_Align_Center
    Elseif BumpCounter<2 Then
      BumpCounter = 0
      StopVideosPlaying
    End If
  Else
    BumpCounter = 0 : BG_Bump1=0 : BG_Bump2 = 0 : ShowSuperJets=0
  End If

End Sub


'Sub DMDTracylights
' if currentmission=0 Then
'
'   Show_TracyHits=Show_TracyHits-1
'   If Show_TracyHits = 0 Then
'     FlexDMD.Stage.GetLabel("Splash2").Visible = False
'   Else
'     FlexDMD.Stage.GetLabel("Splash2").Text = "TRACY LIT IN " & 15-Tracy_bumpers  & " HITS"
'     FlexDMD.Stage.GetLabel("Splash2").SetAlignedPosition 64,31, FlexDMD_Align_Center
'     FlexDMD.Stage.GetLabel("Splash2").Visible = True
'   End If
'
' End If
'End Sub
'
'Sub DMDSuperJets
' if currentmission=0 Then
'
'   BumperSplash=BumperSplash-1
'   If (bumpersplash mod 3)=1 Then
'     FlexDMD.Stage.GetLabel("Splash").Font = FontScoreInactiv1
'   Else
'     FlexDMD.Stage.GetLabel("Splash").Font = FontScoreInactiv2
'   End If
'
'   If bumpersplash = 0 Then
'     FlexDMD.Stage.GetLabel("Splash").Visible = False
'   Elseif SuperJets<50 Then
'     FlexDMD.Stage.GetLabel("Splash").Text = "SUPER JETS IN " & 50-SuperJets & " HITS"
'     FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,25, FlexDMD_Align_Center
'     FlexDMD.Stage.GetLabel("Splash").Visible = True
'   Else
'     FlexDMD.Stage.GetLabel("Splash").Text = "SUPER JETS ACTIVATED"
'     FlexDMD.Stage.GetLabel("Splash").SetAlignedPosition 64,25, FlexDMD_Align_Center
'     FlexDMD.Stage.GetLabel("Splash").Visible = True
'   End If
'
' End If
'End Sub
'

Sub Start_Video
    FlexDMD.Stage.GetVideo("videoplaying").remove : FlexDMD.stage.GetLabel("videotext").Remove
    FlexDMD.Stage.GetGroup("Score").Addactor FlexDMD.Newvideo("videoplaying","Videos/" & VidList(CurrentVideo) & ".gif")
    FlexDMD.Stage.GetGroup("Score").Addactor FlexDMD.NewLabel("videotext", FontScoreInactiv1, " " )
End Sub

Sub OverlayBigScore
  If Overlay1 Then ' black overlays on top of score only
    Select Case OverlayDir
      Case 0 :
      OverlayPos=OverlayPos+2
      If OverlayPos>200 Then OverlayDir=1
      Case 1 :
      OverlayPos=OverlayPos-2
      If OverlayPos<1 Then OverlayDir=0
    End Select
    If OverlayPos < 67 then FlexDMD.Stage.GetImage("ScOverlay1").SetPosition 0,OverlayPos-32
  End If
End Sub

Const MaxVideos = 105
Dim VidList(105)
VidList( 1)="Aiming Gun"
VidList( 2)="Bloodmachines logo"
VidList( 3)="Bodies fall on floor"
VidList( 4)="Close up_Hoodie off"
VidList( 5)="Close up_mask on"
VidList( 6)="Commander ship out hyperspace"
VidList( 7)="Corey putting on mask"
VidList( 8)="Corey shoots Vascan"
VidList( 9)="Corey_Attack_End"
VidList(10)="Corey_Attack_hit1"
VidList(11)="Corey_Attack_hit2"
VidList(12)="Corey_Attack_hit3"
VidList(13)="Corey_Attack_hit4"
VidList(14)="Corey_Attack_hit5"
VidList(15)="Corey_Attack_hit6"
VidList(16)="Corey_Attack_move1"
VidList(17)="Corey_Attack_move2"
VidList(18)="Corey_Attack_move3"
VidList(19)="Corey_Attack_move4"
VidList(20)="Corey_Attack_move5"
VidList(21)="Corey_Attack_move6"
VidList(22)="Corey_Attack_move7"
VidList(23)="Corey_Look over shoulder"
VidList(24)="Corey_platform rising"
VidList(25)="Corey_platform standing"
VidList(26)="Corey_Pulls out Vascans guts"
VidList(27)="Corey_raising arms"
VidList(28)="Corey_Slams hands in Vascan chest"
VidList(29)="Corey_tearing ship open"
VidList(30)="Draw Gun"
VidList(31)="Draw Gun_1"
VidList(32)="Draw Gun_2"
VidList(33)="Draw Gun_3"
VidList(34)="Draw Gun_4"
VidList(35)="Endshot movie"
VidList(36)="Game Over_1"
VidList(37)="Game Over_2"
VidList(38)="Game_Over_1"
VidList(39)="Ground glows"
VidList(40)="Lady walking to camera"
VidList(41)="Lago Hail Mary"
VidList(42)="Lago walks to camera"
VidList(43)="Lago walks to Ships Window"
VidList(44)="Lago_Hart of steel_thumbs up"
VidList(45)="Lago_pulls wires out off Tracy"
VidList(46)="Lago_with hatchet"
VidList(47)="Mima_Beauty shot"
VidList(48)="Mima_Body pass"
VidList(49)="Mima_Floating in space"
VidList(50)="Mima_Floating over camera"
VidList(51)="Mima_pass camera to distance"
VidList(52)="Mima_Rising in space"
VidList(53)="Pyramid of Souls"
VidList(54)="Scavengers breaks ships gun"
VidList(55)="Scavengers hitting ship"
VidList(56)="Scavengers pulling wires out of ship"
VidList(57)="Ship closing"
VidList(58)="Ship Flyby"
VidList(59)="Ship Flying trough Dump"
VidList(60)="Ship Flying trough Dump_2"
VidList(61)="Ship Flying trough Dump_3"
VidList(62)="Ship Follows Mima_side shot"
VidList(63)="Ship frontal shot"
VidList(64)="Ship Guns Aiming"
VidList(65)="Ship Guns Aiming_2"
VidList(66)="Ship landing"
VidList(67)="Ship opening hatch"
VidList(68)="Ship_side shot_closeup"
VidList(69)="Shooting circuit board"
VidList(70)="Shooting Gun_1"
VidList(71)="Shooting Gun_2"
VidList(72)="Shoots himself"
VidList(73)="Shoots himself_closeup"
VidList(74)="Spaceship cockpit  view"
VidList(75)="Spaceship in persuit"
VidList(76)="Spaceship Warpspeed"
VidList(77)="Supernova appears"
VidList(78)="Tracy 360 shot"
VidList(79)="Tracy_arms out body"
VidList(80)="Tracy_Head Shaking"
VidList(81)="Tracy_Robot Dies"
VidList(82)="Vascan_Cory Walking trough Portal"
VidList(83)="Vascan_Dies and falls on floor"
VidList(84)="Vascan_Gets shot_Closeup"
VidList(85)="Vascan_hits Corey with his gun"
VidList(86)="Vascan_Looking in Orb"
VidList(87)="Vascan_Pulls off circuit board"
VidList(88)="Vascan_Smacks Corey to ground"
VidList(89)="Vascan_Throws board to Lago"
VidList(90)="Walking in cockpit"
VidList(91)="x2"
VidList(92)="x3"
VidList(93)="x4"
VidList(94)="x5"
VidList(95)="x6"
VidList(96)="x7"
VidList(97)="x8"
VidList(98)="Attract_games_specs"
VidList(99)="Attract_Push_Start"
VidList(100)="Attract_flyer_text"
VidList(101)="Attract_Bring_back_the_Mima"
VidList(102)="DMD_Attract_EaShDa"
VidList(103)="Attract_Intro_Credtis"
VidList(104)="Attract_still_reading"
VidList(105)="Bloodmachines logo"


Dim CurrentVideo : CurrentVideo=0
Sub DMDShowVideos2
  'debug.print "curret" & CurrentVideo & "    playing" & VideoPlaying
    If ( Frame mod 8 ) = 1 Then
        FlexDMD.stage.GetLabel("videotext").font=FontScoreInactiv2
    Elseif ( Frame mod 8 ) = 5 Then
        FlexDMD.stage.GetLabel("videotext").font=FontScoreInactiv1
    End If

    If currentvideo<>VideoPlaying Then
      FlexDMD.stage.GetVideo("videoplaying").remove : FlexDMD.stage.GetLabel("videotext").Remove
      currentvideo=VideoPlaying
      FlexDMD.Stage.Addactor FlexDMD.Newvideo ("videoplaying","Videos/" & VidList(VideoPlaying) & ".gif")
'     FlexDMD.Stage.GetVideo("videoplaying").visible=True
      FlexDMD.Stage.Addactor FlexDMD.NewLabel("videotext", FontScoreInactiv1, VideoPlaying & " - " & VidList(VideoPlaying) )
      FlexDMD.Stage.GetLabel("videotext").SetAlignedPosition 4,27, FlexDMD_Align_Left
'     FlexDMD.Stage.GetLabel("videotext").visible=True
    End If
  FlexDMD.UnlockRenderThread
End Sub

Sub StopVideosPlaying
  If currentvideo > 0 Then
    BG_Bump1=0 : BG_Bump2=0 : BumpCounter=0
    currentvideo=0
    FlexDMD.stage.GetVideo("videoplaying").visible=False
    FlexDMD.stage.GetLabel("videotext").visible=False
  End If
End Sub


Sub DMDmovebar1
  BG_BarLeft=BG_BarLeft+20
  If BG_BarLeft>170 Then
    BG_BarLeft=0
    FlexDMD.Stage.GetImage("Background6").Visible=False
    Exit Sub
  End If

  FlexDMD.Stage.GetImage("Background6").SetPosition BG_BarLeft-50, 0
  FlexDMD.Stage.GetImage("Background6").Visible=True
End Sub


Sub DMDmovebar2
  BG_BarRight=BG_BarRight-20
  If BG_BarRight<-170 Then
    BG_BarRight=0
    FlexDMD.Stage.GetImage("Background6").Visible=False
    Exit Sub
  End If

  FlexDMD.Stage.GetImage("Background6").SetPosition BG_BarRight+80, 0
  FlexDMD.Stage.GetImage("Background6").Visible=True



End Sub

Sub DMDaddstars
  If BG_stars=-2 Then BG_Stars =0 : BG_Star1=10 : Exit Sub
  If BG_stars=-1 Then BG_Stars =0 : BG_Star2=10 : Exit Sub
  If BG_starSwap=0 And BG_Star1=0 Then BG_Star1=10 : BG_stars=BG_stars-1 : BG_starSwap=1 : Exit Sub
  If BG_starSwap=1 And BG_Star2=0 Then BG_Star2=10 : BG_stars=BG_stars-1 : BG_starSwap=0 : Exit Sub
End Sub


Dim BG_Star1, BG_Star2, BG_starSwap
Sub DMDstar2
  BG_Star2=BG_Star2-1
  Select Case BG_Star2
    Case  9 :   FlexDMD.Stage.GetImage("Background5").Visible=False
    Case  7 :   FlexDMD.Stage.GetImage("Background5").Visible=True
    Case  5 :   FlexDMD.Stage.GetImage("Background5").Visible=False
    Case  3 :   FlexDMD.Stage.GetImage("Background5").Visible=True
    Case  1 :   FlexDMD.Stage.GetImage("Background5").Visible=False
  End Select
End Sub
Sub DMDstar1
  BG_Star1=BG_Star1-1
  Select Case BG_Star1
    Case  9 :   FlexDMD.Stage.GetImage("Background4").Visible=False
    Case  7 :   FlexDMD.Stage.GetImage("Background4").Visible=True
    Case  5 :   FlexDMD.Stage.GetImage("Background4").Visible=False
    Case  3 :   FlexDMD.Stage.GetImage("Background4").Visible=True
    Case  1 :   FlexDMD.Stage.GetImage("Background4").Visible=False
  End Select
End Sub


Sub DMDBorderBlink1
    Borderblink=Borderblink-1
    Select Case Borderblink
      Case  4 :   FlexDMD.Stage.GetImage("Border2").Visible=True  : FlexDMD.Stage.GetImage("Border3").Visible=False
      Case  3 :   FlexDMD.Stage.GetImage("Border3").Visible=True  : FlexDMD.Stage.GetImage("Border2").Visible=False
      Case  2 :   FlexDMD.Stage.GetImage("Border2").Visible=True  : FlexDMD.Stage.GetImage("Border3").Visible=False
      Case  1 :   FlexDMD.Stage.GetImage("Border3").Visible=True  : FlexDMD.Stage.GetImage("Border2").Visible=False
      Case  0 :   FlexDMD.Stage.GetImage("Border3").Visible=False
    End Select
End Sub


Sub DMDScoreBlink
  dim label
'   Set label = FlexDMD.Stage.GetLabel("Player" & i)
  if PlayersPlayingGame = 0 then exit sub
  if PlayersPlayingGame = 1 then
    'debug.print "single player"
    Set label = FlexDMD.Stage.GetLabel("Score_1")
    Borderblink=Borderblink-1
    Select Case Borderblink
      Case  4 :   label.font = FontScoreActive
      Case  3 :   label.font = FontScoreActiveG
      Case  2 :   label.font = FontScoreActive
      Case  1 :   label.font = FontScoreActiveG
      Case  0 :   label.font = FontScoreActive
    End Select
  Else
    Set label = FlexDMD.Stage.GetLabel("Player" & CurrentPlayer)
    'debug.print "multi: " & Borderblink
    Borderblink=Borderblink-1
    Select Case Borderblink
      Case  4 :   label.font = Font5by7W
      Case  3 :   label.font = Font5by7G
      Case  2 :   label.font = Font5by7W
      Case  1 :   label.font = Font5by7G
      Case  0 :   label.font = Font5by7W
    End Select
  end if
End Sub


Sub DMDMysteryDisplay
' If DMDmystery>1 Then StopVideosPlaying
' FlexDMD.stage.GetLabel("Score_1").Visible = false

  DMDmystery=DMDmystery+1
  If (DMDmystery mod 3)=1 Then
    FlexDMD.Stage.GetLabel("Splash3").Font = FontScoreInactiv1
  Else
    FlexDMD.Stage.GetLabel("Splash3").Font = FontScoreInactiv2
  End If

  Select Case DMDMystery
    Case 1,5,9,13,17,21,25,29 : DMD_DisplayMystery int(rnd(1)*10)+1
    Case 33 : DMD_DisplayMystery DMDMysteryReward
        RewardMystery
        ScavengeKicker.timerenabled = true
        Scavengeball=1
    Case 38,44,50 : DMD_DisplayMystery DMDMysteryReward
    Case 36,42,48,55 : FlexDMD.Stage.GetLabel("Splash3").Text = "  "
  End Select
End Sub


Sub DMD_DisplayMystery(Number)
  Select Case Number
    Case 1 : FlexDMD.Stage.GetLabel("Splash3").Text = "10.000.000"
    Case 2 : FlexDMD.Stage.GetLabel("Splash3").Text = "800.000"
    Case 3 : FlexDMD.Stage.GetLabel("Splash3").Text = "500.000"
    Case 4 : FlexDMD.Stage.GetLabel("Splash3").Text = "15.000.000"
    Case 5 : FlexDMD.Stage.GetLabel("Splash3").Text = "17.000.000"
    Case 6 : FlexDMD.Stage.GetLabel("Splash3").Text = "20.000.000"
    Case 7 : FlexDMD.Stage.GetLabel("Splash3").Text = "20.000.000"
    Case 8 : FlexDMD.Stage.GetLabel("Splash3").Text = "20.000.000"
    Case 9 : FlexDMD.Stage.GetLabel("Splash3").Text = "EXTRABALL LIT"
  End Select

  FlexDMD.Stage.GetLabel("Splash3").SetAlignedPosition 64,30, FlexDMD_Align_Center
  FlexDMD.Stage.GetLabel("Splash3").Visible = True
End Sub



'*************************************************
'*****   FLEX DMD ENDS
'*************************************************





'********************
' MATHS
'********************

Const Pi=3.1415926535

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function


Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

Function min(a,b)
  if a > b then
    min = b
  Else
    min = a
  end if
end Function


'*****************************************************************************************************************************************
'   KEYS
'*****************************************************************************************************************************************
'Dim LeftF,RightF
dim bOpenTopGate:bOpenTopGate = False
dim bPlungerPulled : bPlungerPulled = False
dim SKBackup : SKBackup = 0

sub SkiptoEndTotal
  if HasPuP and (EobBonusCounter > 2 and EobBonusCounter < 809) then
    EobBonusCounter = 810 'skip EOBbonus counting PUP
  elseif Not HasPuP and (EobBonusCounter > 2 and EobBonusCounter < 499) then
    EobBonusCounter = 500 'skip EOBbonus counting Flex
  end if
'debug.print "skip to: " & EobBonusCounter

end sub

Sub Table1_KeyDown(ByVal Keycode)

  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Credits=Credits+1 : If Credits>30 Then Credits=30 Else Creditblink=20
    DOF 133,2 'add credits
    If pDMDmode="attract" and hasPUP then
        pDMDSetPage(9):puPlayer.LabelSet pDMD,"Credits2"," Credits: " & Credits,1,""
    End If
    if Credits > 0 then DOF 122,1 'start button light
    SaveCredits
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If


  If keycode = 19 then ScoreCard=1 : CardTimer.enabled=True

  If hsbModeActive Then EnterHighscoreName(keycode) : exit Sub


  if keycode = StartGameKey then
   soundStartButton
   CoinErrorText
  end if

  if bBallInPlungerLane And keycode = LeftFlipperKey Then
    'debug.print "OpenTopGate left flip"
    bOpenTopGate = true
  end if

  If keycode = RightFlipperKey Then
    ScrollAttract
  End if

  If keycode = PlungerKey Then
    'debug.print "plunger" & lfpress & " " & bOpenTopGate
    bPlungerPulled = true
    SoundPlungerPull()
    Plunger.Pullback
    If VRRoom > 0 Then
      TimerVRPlunger.Enabled = True
      TimerVRPlunger2.Enabled = False
    End If
  End If

  CheckSkillshotReady

  if bGameInPlay Then
    If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter:CheckTilt
    If keycode = MechanicalTilt Then SoundNudgeCenter:CheckTilt
    If keycode = PlungerKey Then ShootGun true
    If keycode = RightFlipperKey  Then SkiptoEndTotal
    If keycode = LeftFlipperKey  Then SkiptoEndTotal
    If NOT Tilted Then

      if StagedFlipperMod = 1 Then
        If keycode = KeyUpperLeft and bFlippersEnabled Then
          Flipper1.RotateToEnd
          If Flipper1.currentangle < Flipper1.endangle + ReflipAngle Then
            'Play partial flip sound and stop any flip down sound
            StopAnyFlipperUpperLeftDown()
'           RandomSoundFlipperUpperLeftReflipDOF Flipper1, 105, DOFOn
            RandomSoundFlipperUpperLeftReflip Flipper1
vr            'Play full flip sound
'           RandomSoundFlipperUpperLeftUpFullStrokeDOF Flipper1, 105, DOFOn
            RandomSoundFlipperUpperLeftUpFullStroke Flipper1
          End If
        end if
      end if

      If keycode = LeftFlipperKey and bFlippersEnabled Then
'       LeftF = 1 : Flipper_Info = 0

'       DOF 101,1
'       flash2 1
        If skillshot=0 Then
          if Not CurrentMission = 5 And Not WizardPhase = 3 then
            RotateLaneLightsLeft
          end if
        End If

        FlipperActivate LeftFlipper, LFPress
        LF.Fire
'       FlipperHoldCoilLeft SoundOn, LeftFlipper
        'over ramp flipper
        if StagedFlipperMod = 0 then
          Flipper1.RotateToEnd
          If Flipper1.currentangle < Flipper1.endangle + ReflipAngle Then
            'Play partial flip sound and stop any flip down sound
            StopAnyFlipperUpperLeftDown()
'           RandomSoundFlipperUpperLeftReflipDOF Flipper1, 105, DOFOn
            RandomSoundFlipperUpperLeftReflip Flipper1
          Else
            'Play full flip sound
'           RandomSoundFlipperUpperLeftUpFullStrokeDOF Flipper1, 105, DOFOn
            RandomSoundFlipperUpperLeftUpFullStroke Flipper1
          End If
        end if

        'underpf left flipper
        If bPortalFlippersEnabled Then
          LeftFlipperMini.RotateToEnd
          If LeftFlipperMini.currentangle < LeftFlipperMini.endangle + ReflipAngle Then
            'Play partial flip sound and stop any flip down sound
            StopAnyFlipperUpperLeftDown()
'           RandomSoundFlipperUpperLeftReflipDOF LeftFlipperMini, 103, DOFOn
            RandomSoundFlipperUpperLeftReflip LeftFlipperMini
          Else
            'Play full flip sound
'           RandomSoundFlipperUpperLeftUpFullStrokeDOF LeftFlipperMini, 103, DOFOn
            RandomSoundFlipperUpperLeftUpFullStroke LeftFlipperMini
          End If
        end if

        If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
          'Play partial flip sound and stop any flip down sound
          StopAnyFlipperLowerLeftDown()
          RandomSoundFlipperLowerLeftReflipDOF LeftFlipper, 101, DOFOn
        Else
          'Play full flip sound
          If BallNearLF = 0 Then
            RandomSoundFlipperLowerLeftUpFullStrokeDOF LeftFlipper, 101, DOFOn
          End IF
          If BallNearLF = 1 Then
            Select Case Int(Rnd*2)+1
              Case 1 : RandomSoundFlipperLowerLeftUpDampenedStrokeDOF LeftFlipper, 101, DOFOn
              Case 2 : RandomSoundFlipperLowerLeftUpFullStrokeDOF LeftFlipper, 101, DOFOn
            End Select
          End If
        End If
        If VRroom > 0 Then
          VRFlipperButtonLeft.X = VRFlipperButtonLeft.X + 8
        End if

        CheckSkillShotReady
      end if

      If keycode = RightFlipperKey and bFlippersEnabled Then
'       flash6 1
        If skillshot=0 Then
          if Not CurrentMission = 5 And Not WizardPhase = 3 then
            RotateLaneLightsRight
          end if
        End If

        FlipperActivate RightFlipper, RFPress
        RF.Fire
'       FlipperHoldCoilRight SoundOn, RightFlipper
        'underpf right flipper
        If bPortalFlippersEnabled Then
          RightFlipperMini.RotateToEnd

          If RightFlipperMini.currentangle < RightFlipperMini.endangle + ReflipAngle Then
            'Play partial flip sound and stop any flip down sound
            StopAnyFlipperUpperLeftDown()
'           RandomSoundFlipperUpperLeftReflipDOF RightFlipperMini, 104, DOFOn
            RandomSoundFlipperUpperLeftReflip RightFlipperMini
          Else
            'Play full flip sound
'           RandomSoundFlipperUpperLeftUpFullStrokeDOF RightFlipperMini, 104, DOFOn
            RandomSoundFlipperUpperLeftUpFullStroke RightFlipperMini
          End If

        end if

        If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
          StopAnyFlipperLowerRightDown()
          RandomSoundFlipperLowerRightReflipDOF RightFlipper, 102, DOFOn
        Else
          'Play full flip sound
          If BallNearRF = 0 Then
            RandomSoundFlipperLowerRightUpFullStrokeDOF RightFlipper, 102, DOFOn
          End IF
          If BallNearRF = 1 Then
            Select Case Int(Rnd*2)+1
              Case 1 : RandomSoundFlipperLowerRightUpDampenedStrokeDOF RightFlipper, 102, DOFOn
              Case 2 : RandomSoundFlipperLowerRightUpFullStrokeDOF RightFlipper, 102, DOFOn
            End Select
          End If
        End If

        If VRRoom > 0 Then
          VRFlipperButtonRight.X = VRFlipperButtonRight.X - 8
        End If
      end if

      If KeyCode=LockBarKey Then
        ShootGun true
      End If

      If (KeyCode=RightMagnaSave or KeyCode = LockBarKey) and bFlippersEnabled Then
        DropLoopWall true
        ShootGun true
        RightMagnaTimer.Enabled = True
      end if

      if Keycode=LeftMagnaSave and bFlippersEnabled Then
        LeftMagnaTimer.Enabled = True
      end if

      If (Keycode=LeftFlipperKey Or KeyCode=RightFlipperKey) and bFlippersEnabled Then
        if WizardPhase <> 1 then
          DropLoopWall true
        end if
      end if

'     'disabled as this has issues with dmd
'     If keycode = StartGameKey And LFPress And Not Slammed Then
'       'msgbox "SLAM!!"
'       Tilted = True
'       Slammed = True
'       DisableTable True
'       tilttableclear.enabled = true
'       TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
'       gioff
'       'ResetForNewGame
'       FlexDMD.stage.GetLabel("Score_1").visible = False
'       MPCornerLabelsVisible false
'       playsound "Knocker_1"
'       bBallSaverActive = False
'
'       vpmtimer.addtimer 3000, "Table1_Init '"
'       vpmtimer.addtimer 7000, "Slammed = False:StartAttractMode:FlexDMD.Stage.GetLabel(""Splash"").Text = """" '"
'
'     end if


      If keycode = StartGameKey  And Credits>0 Then
        If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then
          PlayersPlayingGame = PlayersPlayingGame + 1

          If PlayersPlayingGame = 2 Then
          GameplayerNotup = "2"
          SmallScores
          End If

          If PlayersPlayingGame = 3 Then
          GameplayerNotup = "3"
          SmallScores
          End If

          If PlayersPlayingGame = 4 Then
          GameplayerNotup = "4"
          SmallScores
          End If

          If PlayersPlayingGame > 1 Then
            'Hide center score
            FlexDMD.stage.GetLabel("Score_1").Visible = False
          End If

          Credits=Credits-1
          SaveCredits

          if Credits <= 0 then DOF 122,0  'Turn off start button light
          DOF 123,2 'start game

          Creditblink=20

          TotalGamesPlayed = TotalGamesPlayed + 1
          savegp
        End If
      Elseif keycode = StartGameKey  And Credits=0 Then
        StopVideosPlaying
        InsertCoins=1

      End If
    End If
  Else
    If NOT Tilted Then

'     If keycode = LeftFlipperKey Then helptime.enabled = true':DMDintroloop:introtime = 0
'     If keycode = RightFlipperKey Then helptime.enabled = true':DMDintroloop:introtime = 0
      If keycode = StartGameKey And Credits>0 And IntroOver Then
        If(BallsOnPlayfield = 0) Then
          ResetForNewGame()
          Credits=Credits-1
          SaveCredits
          CreditBlink=20
        End If
      Elseif keycode = StartGameKey And Credits=0 And IntroOver Then
        'insertcoin1
        StopVideosPlaying
        InsertCoins=1
      End If
    End If
  End If


End Sub
Dim InsertCoins : InsertCoins=0

Sub Table1_KeyUp(ByVal keycode)

  if keycode = 19 then ScoreCard=0

  If keycode = PlungerKey Then
    Plunger.Fire
    If (bBallInPlungerLane) Then
      SoundPlungerPullStop()
      SoundPlungerReleaseBall()
    Else
      SoundPlungerPullStop()
      SoundPlungerReleaseNoBall()
    End If
    PuPEvent(599)
'   bOpenTopGate = false
    bPlungerPulled = false
    If VRRoom > 0 Then
      TimerVRPlunger.Enabled = False
      TimerVRPlunger2.Enabled = True
    End If
  End If

  if keycode = LeftFlipperKey And bBallInPlungerLane Then
    bOpenTopGate = false
    LS_JP3all.stopPlay
    SkillshotStart
    'Skillshot = SKBackup
  end if


  ' Table specific

  If StagedFlipperMod = 1 Then
    If keycode = KeyUpperLeft Then
      Flipper1.RotateToStart
'     RandomSoundFlipperDownLeft Flipper1
      If Flipper1.currentangle < Flipper1.startAngle - 5 Then
'       RandomSoundFlipperUpperLeftDownDOF Flipper1, 105, DOFOff
        RandomSoundFlipperUpperLeftDown Flipper1
      End If
      FlipperLeftUpperHitParm = FlipperUpSoundLevel
    end if
  End If

  If keycode = LeftFlipperKey Then
'   DOF 101,0
    FlipperDeActivate LeftFlipper, LFPress
'   FlipperHoldCoilLeft SoundOff, LeftFlipper

    if StagedFlipperMod = 0 then
      Flipper1.RotateToStart
'     RandomSoundFlipperDownLeft Flipper1
      If Flipper1.currentangle < Flipper1.startAngle - 5 Then
'       RandomSoundFlipperUpperLeftDownDOF Flipper1, 105, DOFOff
        RandomSoundFlipperUpperLeftDown Flipper1
      End If
      FlipperLeftUpperHitParm = FlipperUpSoundLevel
    end if


    LeftFlipperMini.RotateToStart
    If bPortalFlippersEnabled Then
'     RandomSoundFlipperDownLeft LeftFlipperMini
      If LeftFlipperMini.currentangle < LeftFlipperMini.startAngle - 5 Then
'       RandomSoundFlipperUpperLeftDownDOF LeftFlipperMini, 103, DOFOff
        RandomSoundFlipperUpperLeftDown LeftFlipperMini
      End If
'     FlipperLeftUpperHitParm = FlipperUpSoundLevel
    end if


    LeftFlipper.RotateToStart
'   If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
'     RandomSoundFlipperDownLeft LeftFlipper
'   End If
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperLowerLeftDownDOF LeftFlipper, 101, DOFOff
    End If
    FlipperLeftLowerHitParm = FlipperUpSoundLevel




'   FlipperLeftHitParm = FlipperUpSoundLevel
    If VRRoom > 0 Then
      VRFlipperButtonLeft.X = 2102.962
    End If
  End If

  If keycode = RightFlipperKey Then
'   DOF 102,0
    FlipperDeActivate RightFlipper, RFPress
    RightFlipperMini.RotateToStart

    If bPortalFlippersEnabled Then
      If RightFlipperMini.currentangle < RightFlipperMini.startAngle - 5 Then
'       RandomSoundFlipperUpperLeftDownDOF RightFlipperMini, 104, DOFOff
        RandomSoundFlipperUpperLeftDown RightFlipperMini
      End If
    end if

'   RightFlipper.RotateToStart
'   If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
'     RandomSoundFlipperDownRight RightFlipper
'   End If
'   FlipperRightHitParm = FlipperUpSoundLevel

    RightFlipper.RotateToStart
'   FlipperHoldCoilRight SoundOff, RightFlipper
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperLowerRightDownDOF RightFlipper, 102, DOFOff
    End If
    FlipperRightLowerHitParm = FlipperUpSoundLevel


    If VRRoom > 0 Then
      VRFlipperButtonRight.X = 2093.387
    End If
  End If

  If bGameInPLay and hsbModeActive <> True Then
    if Keycode=LeftMagnaSave Then
      LeftMagnaTimer.Enabled = False
    end if

    if Keycode=RightMagnaSave or KeyCode = LockBarKey Then
      RightMagnaTimer.Enabled = False
    end if

  Else
'   If keycode = LeftFlipperKey Then helptime.enabled = false
'   If keycode = RightFlipperKey Then helptime.enabled = false
  End If

End Sub



Sub LeftMagnaTimer_Timer
  LeftMagnaTimer.Enabled = False
End Sub


Sub RightMagnaTimer_Timer
  RightMagnaTimer.Enabled = False
End Sub


'**********************************************************************************************************
'* InstructionCard *
'**********************************************************************************************************

Dim CardCounter, ScoreCard
Sub CardTimer_Timer
  If scorecard=1 Then
    CardCounter=CardCounter+2
    If CardCounter>50 Then CardCounter=50
  Else
    CardCounter=CardCounter-4
    If CardCounter<0 Then CardCounter=0
  End If
  InstructionCard.transX = CardCounter*5.2
  InstructionCard.transY = CardCounter*7
  InstructionCard.transZ = -cardcounter*2

    InstructionCard.objRotX = -cardcounter/3.7
  InstructionCard.size_x = 1+CardCounter/25
  InstructionCard.size_y = 1+CardCounter/25
  If CardCounter=0 Then
      CardTimer.Enabled=False
      InstructionCard.visible=0
  Else
      InstructionCard.visible=1
  End If
End Sub


'*****************************************************************************************************************************************
'  RAMP CONTROL
'*****************************************************************************************************************************************

'Shot2 - Ramp down
sub CaptKick_hit
  WriteToLog "CaptKick_hit", ""
' SoundSaucerLock
  RandomSoundScoopLeftEnter CaptKick
  'msgbox "Catch! " & activeball.ID
  CockpitLight true
  set CaptiveBallID = activeball
  if Not TargetCockpit.isdropped Then ' if captive ball should be down still, let it drop the ball back. This may happen with really straight and powerful hits.
    CaptKick.kick 0, 10
'   SoundSaucerKick 1, CaptKick
    RandomSoundScoopLeftEjectNormalVelocity CaptKick
    CockpitLight false
  end if
end sub


dim CaptiveBallID, CaptiveBallOrigX, CaptiveBallOrigY, CaptiveBallOrigZ
CaptiveBallOrigZ = -1


Const Capttargetx = 273
Const Capttargety = 240
Const Capttargetz = 95

Dim Counter_RampMima : Counter_RampMima = 5
Sub Rise_RampMima_Timer

  if CaptiveBallOrigZ = -1 then
    CaptiveBallOrigX = CaptiveBallID.x
    CaptiveBallOrigY = CaptiveBallID.y
    CaptiveBallOrigZ = CaptiveBallID.z

    'debug.print CaptiveBallOrigX & " - " & CaptiveBallOrigY & " - " & CaptiveBallOrigZ
  end if



  Rise_RampMima.interval=17
  Lower_RampMima.interval=300
  Lower_RampMima.Enabled=False
  Counter_RampMima = Counter_RampMima - 1
  CockpitRamp.collidable=False
  CockpitRamp001.collidable=False
  CockpitRamp002.collidable=False
  RampTopBlock.collidable=False
  TargetCockpit.isdropped = true
  'debug.print "Counter_Ramp=" & Counter_RampMima

  If Counter_RampMima > 4 Then
    ShipJaw.RotX = -Counter_RampMima
    'CaptiveBallID.x = -0.1
    if CaptiveBallID.x > Capttargetx then
      CaptiveBallID.x = CaptiveBallID.x - ((CaptiveBallOrigX - Capttargetx)/5)
    end if
    if CaptiveBallID.y > Capttargety then
      CaptiveBallID.y = CaptiveBallID.y - ((CaptiveBallOrigY - Capttargety)/5)
    end if
    if CaptiveBallID.z < Capttargetz then
      CaptiveBallID.z = CaptiveBallID.z + ((Capttargetz - CaptiveBallOrigZ)/20)
    end if
    'debug.print CaptiveBallID.x & " - " & CaptiveBallID.y & " - " & CaptiveBallID.z
  Else
    CaptiveBallID.x = Capttargetx
    CaptiveBallID.y = Capttargety
    CaptiveBallID.z = Capttargetz
    wVortexBlocker.isdropped = true
    VorMag.MagnetOn = 1
    ' fixing sound
    Rise_RampMima.interval=300
    Rise_RampMima.enabled=False
    if NumLocksMima >= 3 and cLightMimaMB.state = 0 then
      playsound "SPC_mima_lock_ready"
    end if
    WriteToLog "Rise_RampMima_Timer", "Vortex Open"
  End If

End Sub

Sub Lower_RampMima_Timer
  if not bWizardMode then
    Rise_RampMima.interval=300
    Rise_RampMima.Enabled=False


    Lower_RampMima.interval=17
    Counter_RampMima = Counter_RampMima + 1
    CockpitRamp.collidable=True
    CockpitRamp001.collidable=True
    CockpitRamp002.collidable=True
    RampTopBlock.collidable=True
    TargetCockpit.isdropped = false
    'debug.print "Counter_Ramp=" & Counter_RampMima

    If Counter_RampMima < 21 Then
      ShipJaw.RotX = -Counter_RampMima
    Else
      ' fixing sound
      wVortexBlocker.isdropped = false
      'VortexMagnet.Enabled = false
      VorMag.MagnetOn = 0
      Lower_RampMima.interval=300
      Lower_RampMima.enabled=False

      If IsEmpty(CaptiveBallID) Then
        CaptKick.createball
      end if
      CaptKick.kick 0, 10
      RandomSoundScoopLeftEjectNormalVelocity CaptKick
      WriteToLog "Lower_RampMima_Timer", "Vortex Closed"
    End If
  End If
End Sub

dim CaptiveTriggerCooldown : CaptiveTriggerCooldown = false

sub CaptiveTrigger_hit
' WriteToLog "CaptiveTrigger_hit", ""
    If Tilted Then Exit Sub
  set CaptiveBallID = activeball
  CaptiveBallID.frontdecal = "ballshadow"
  if activeball.vely < -2 And Not CaptiveTriggerCooldown then 'only accepting hits that come from below
    'msgbox "Captive ball hit: --> " & activeball.vely
    WriteToLog "CaptiveTrigger_hit", "activeball.vely: " & activeball.vely
    DOF 108,2
    Playsound "EFX_Ship_sound0" & rndint(1,4),0,CalloutVol,0,0,1,1,1
    CaptiveTriggerCooldown = true
    CheckBladerunnerShot 2
    CheckChaseShot 2
    AdvanceToMimaMB
    AddScore score_ShipHit

    If cLightCorey2.state = 2 and not bMultiballMode Then
      cLightCorey2.state = 1
      CheckCoreyMB
    End If
    If cLightTracy2.state = 2 and not bMultiballMode Then
      cLightTracy2.state = 1
      CheckTracyMB
    End If

    shake_tracy_shake.enabled = true
'   shake_ship.enabled = true
    CockpitLight true
  Else
    WriteToLog "CaptiveTrigger_hit", "too slow ball activeball.vely: " & activeball.vely
    CockpitLight false
  end if
end sub

sub CaptiveTrigger_unhit
  CockpitLight false 'make sure that it really goes off
end sub
'TargetCockpit is now the metal wire that holds the captive ball
Sub TargetCockpit_hit
' MetalWireBallGuideHit
  '4.5.17
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_4"), Vol(ActiveBall) * MetalImpactSoundFactor * 2 'Just to piss of Dan
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_5"), Vol(ActiveBall) * MetalImpactSoundFactor * 2
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_17"), Vol(ActiveBall) * MetalImpactSoundFactor * 2
  end Select
End Sub

dim CockpitLightLvl, CockpitLightLvlTarget
CockpitLightLvl = 0

sub CockpitLight(Enabled)
  if Enabled Then
    CockpitLightLvlTarget = 1
    CockpitLightTimer_timer
  else
    CockpitLightLvlTarget = 0
    CockpitLightTimer_timer
  end if
end sub

sub CockpitLightTimer_timer
  if Not CockpitLightTimer.enabled Then
    ship_2.visible = true
    CockpitLightTimer.enabled = true
  end If

  if CockpitLightLvl < CockpitLightLvlTarget Then
    CockpitLightLvl = Min(1, CockpitLightLvl * 1.5 + 0.3)
  Else
    CockpitLightLvl = Max(0,CockpitLightLvl * 0.9 - 0.01)
  end if

  'UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity, OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive, float elasticity, float elasticityFalloff, float friction, float scatterAngle)
  UpdateMaterial "Ship_cock_on",0,0,0,0,0,0,CockpitLightLvl,RGB(255,255,255),0,0,False,True,0,0,0,0

  if CockpitLightLvl = CockpitLightLvlTarget Then
    if CockpitLightLvl = 0 Then
      ship_2.visible = false
    end If
    CaptiveTriggerCooldown = False
    CockpitLightTimer.enabled = False
  end if

end sub

'dim shipshakecount : shipshakecount = 5
'dim pShipOrigX : pShipOrigX = ship_1.x
'dim pShipOrigY : pShipOrigY = ship_1.Y
'
'sub shake_ship_timer
' Select case shipshakecount
'   case 5: ship_1.y = pShipOrigY - 5:ship_1.x = pShipOrigX - 2
'   case 4: ship_1.y = pShipOrigY + 4:ship_1.x = pShipOrigX + 1.5
'   case 3: ship_1.y = pShipOrigY - 3:ship_1.x = pShipOrigX - 1
'   case 2: ship_1.y = pShipOrigY + 1.5:ship_1.x = pShipOrigX + 0.3
'   case 1: ship_1.y = pShipOrigY - 0.5:ship_1.x = pShipOrigX - 0.1
'   case 0: ship_1.y = pShipOrigY :ship_1.x = pShipOrigX
' End Select
'
' ship_2.y = ship_1.y
' ship_2.x = ship_1.x
' ship_3.y = ship_1.y
' ship_3.x = ship_1.x
'
' shipshakecount = shipshakecount - 1
' if shipshakecount < 0 then
'   shipshakecount = 5
'   me.enabled = false
' end if
'end sub



'*****************************************************************************************************************************************
'  GAME CODE
'*****************************************************************************************************************************************




'***********  tracy head **************
dim tracyEyesBrightness : tracyEyesBrightness = 60

if vrroom <> 0 then tracyEyesBrightness = 40

Primitive019.blenddisablelighting = tracyEyesBrightness

dim tracyposx, tracyposy, tracyTargetAngle, tracyCurrentAngle, tracySpeed, TracyTargetLocked, TracySoundCounter, tracyDir, TracyTrackingTime
tracyCurrentAngle = 0
TracySoundCounter = 30

'this is being increased when tracy is tracking the ball and TracyTargetLocked will increase faster
TracyTrackingTime = 0

'value to increase as mission proceeds. Maybe 1-10.
tracySpeed = 1
'value that increases on lock, and decreases if tracy looses tracking
TracyTargetLocked = 0

tracyposx = Primitive020.x
tracyposy = Primitive020.y

sub TracyTimer_timer
  dim BOT
  BOT = GetBalls

  'exit sub if no balls other than locked ball
  If UBound(BOT) = lob - 1 Then : TracyAngle 0 : Exit Sub : end If

  'follow only one ball on multiball
  'TracyAngle AnglePP(tracyposx, tracyposy, BOT(0).x, BOT(0).y) + 160 + 90
  tracyTargetAngle = AnglePP(tracyposx, tracyposy, BOT(1).x, BOT(1).y) + 160 - 90

  if tracyTargetAngle > 360 then tracyTargetAngle = tracyTargetAngle - 360
  if tracyTargetAngle < 0 then tracyTargetAngle = tracyTargetAngle + 360

  if tracyCurrentAngle > 360 then tracyCurrentAngle = tracyCurrentAngle - 360
  if tracyCurrentAngle < 0 then tracyCurrentAngle = tracyCurrentAngle + 360

  dim delta

  delta = Round(tracyTargetAngle - tracyCurrentAngle,1)
  if delta < 0 then
    delta = delta + 360
  end if
  if delta > 180 Then
    tracyDir = +1
  elseif delta < 180 Then
    tracyDir = -1
  elseif delta > 180 + 360 Then
    tracyDir = -1
  Else
    tracyDir = 0
  end if

  tracyCurrentAngle = tracyCurrentAngle + (tracyDir*tracySpeed*min(abs(delta-180)/10,1))

  'debug.print tracyTargetAngle & " <--- " & tracyCurrentAngle & "  DELTA: " & delta-180


  if abs(delta - 180) < 1 then
    if TracyTargetLocked < 100 then TracyTargetLocked = TracyTargetLocked + (0.55+(TracyTrackingTime*0.003))
    'debug.print "LOCK   " & TracyTargetLocked
  else
    if TracyTargetLocked > 0 then TracyTargetLocked = TracyTargetLocked - 1.1
    'debug.print "....   " & TracyTargetLocked
  end if


  dim red, green, blue
  red = TracyTargetLocked * 2.38 + 17
  green = (217 - TracyTargetLocked * 2.17)
  blue = (27 - TracyTargetLocked * 0.27)
  tracySpeed = tracySpeed + 0.0004
  if tracySpeed >= 6 then tracySpeed = 6  '4

  UpdateMaterial "tracyeyes"    ,0,0,0,0,0,0,1,RGB(red,green,blue),0,0,False,True,0,0,0,0
  Primitive019.blenddisablelighting = tracyEyesBrightness + 2 * TracyTargetLocked


  if TracySoundCounter <= 0 Then
    'play sound plastichit
    'PlaySoundAt "plastichit", Primitive019
    'PlaySound "EFX_alert3"
    'PlaySoundAt "plastichit", Primitive019
    PlaySoundAtLevelStatic "EFX_alert2", 0.01, Primitive019
    TracySoundCounter = 30

    ptracyCross.visible = true
    cTracyCross.state = 1

    cLightBumperHit.state=1
    cLightBumperHit.Timerenabled=True
  end if

  if TracyTargetLocked > 10 And TracyTargetLocked < 30 Then
    TracySoundCounter = TracySoundCounter - 0.3
'   DOF 152,0
    SoundShakerDOF soundoff, dofoff
  elseif TracyTargetLocked > 30 And TracyTargetLocked < 60 Then
    TracySoundCounter = TracySoundCounter - 0.8
'   DOF 152,0
    SoundShakerDOF soundoff, dofoff
  elseif TracyTargetLocked > 60 And TracyTargetLocked < 90 Then
    TracySoundCounter = TracySoundCounter - 1.6
'   DOF 152,1
    SoundShakerDOF soundon, dofon
    startB2S 6
  elseif TracyTargetLocked > 90 And TracyTargetLocked < 99 Then
    TracySoundCounter = TracySoundCounter - 2.4
'   DOF 152,1
    SoundShakerDOF soundon, dofon
    startB2S 6
  end if


  if TracyTargetLocked > 99 Then
    'msgbox "Game over!"
    Playsound "EFX_tracy_sound01",0,CalloutVol,0,0,1,1,1
    vpmtimer.addtimer 500, "PlaySound ""SPC_tracy_anything_broken"",0,CalloutVol,0,0,1,1,1 '"

    EndDoubleScoring
'   DOF 152,0
    SoundShakerDOF soundoff, dofoff
  end if


  TracyAngle tracyCurrentAngle

end sub

dim shakecount : shakecount = 0
sub shake_tracy_shake_timer
  If TracySpins.Enabled=True Then
    TracyAngle spinAngle+RndInt(-25, 25)
  Else
    TracyAngle TracyCurrentAngle+RndInt(-25, 25)
  End If
  'playsound("EFX_tracy_shake" & rndint(1,3))

  shakecount = shakecount + 1
  if shakecount > 10 then
    shakecount = 0
    shake_tracy_shake.enabled = false
    TracyAngle 0
  end if
end sub

sub TracyAngle(aAng)
  Primitive019.roty = aAng
  Primitive020.roty = aAng
  ptracyCross.roty = aAng
end sub


Dim GazeTime:GazeTime = 0

Sub TracyAttention_timer    'Want her to briefly follow the ball
  Dim BOT
  BOT = GetBalls
  If UBound(BOT) = lob - 1 Then : TracyAngle 0 : Exit Sub : end If
  If TracySpins.Enabled = True Or TracyTimer.Enabled = True Then : Me.enabled=False : Exit Sub : End If

  If GazeTime < 50 Then   '10ms * 50 = 0.5s
    tracyTargetAngle = AnglePP(tracyposx, tracyposy, BOT(1).x, BOT(1).y) + 160 - 90

    if tracyTargetAngle > 360 then tracyTargetAngle = tracyTargetAngle - 360
    if tracyTargetAngle < 0 then tracyTargetAngle = tracyTargetAngle + 360

    if tracyCurrentAngle > 360 then tracyCurrentAngle = tracyCurrentAngle - 360
    if tracyCurrentAngle < 0 then tracyCurrentAngle = tracyCurrentAngle + 360

    dim delta
    delta = Round(tracyTargetAngle - tracyCurrentAngle,1)
    if delta < 0 then
      delta = delta + 360
    end if

    if delta > 180 Then
      tracyDir = +1
    elseif delta < 180 Then
      tracyDir = -1
    elseif delta > 180 + 360 Then
      tracyDir = -1
    Else
      tracyDir = 0
    end if

    tracyCurrentAngle = tracyCurrentAngle + (tracyDir*min(abs(delta-180)/10,1))
  Else
    TracyReset.Enabled=True
    GazeTime = 0
    me.Enabled = False
  End If

  GazeTime = GazeTime + 1
  TracyAngle tracyCurrentAngle

End Sub

Sub TracyReset_timer      'Smooth but quick reset to default
  If tracyCurrentAngle > 361 Then
    tracyCurrentAngle = tracyCurrentAngle - 1
  ElseIf tracyCurrentAngle > 180 And tracyCurrentAngle < 359 Then
    tracyCurrentAngle = tracyCurrentAngle + 1
  ElseIf tracyCurrentAngle <= 180 And tracyCurrentAngle > 1 Then
    tracyCurrentAngle = tracyCurrentAngle - 1
  Else
    tracyCurrentAngle = 0
    TracyAngle tracyCurrentAngle
    me.Enabled=False
  End If
  TracyAngle tracyCurrentAngle
End Sub

Dim spinTime : spinTime = 0

Sub PoltergeistTracy(numSpin) 'Spinning Tracy, for Apophis ;) .. Thanks Wylte -apophis
  If TracyTimer.Enabled = True Then Exit Sub
  spinTime = spinTime + numSpin * 360
  TracySpins.Enabled = True
End Sub

Dim spinAngle : spinAngle = 0

Sub TracySpins_Timer
  If spinAngle > spinTime Then
    If  SpinnerMultiplier < 2 Then spinTime = 7200
    spinTime = spinTime - 3600
    spinAngle = 0
  End If
  spinAngle = spinAngle + 10
  TracyAngle spinAngle
  If spinTime <= 0 Then StopTracySpins
End Sub

Sub StopTracySpins
  spinTime = 0
  spinAngle = 0
  TracySpins.enabled = False
  TracyAngle 0
End Sub

'*********** Under playfield & teleports ***********

'tpTrigger1.Enabled = false
'tpTrigger2.Enabled = false
'tp1.isdropped = true
'tp2.isdropped = true

'sub swEnableTeleports_hit
''  tpTrigger1.Enabled = true
''  tpTrigger2.Enabled = true
''  tp1.isdropped = false
''  tp2.isdropped = false
'end sub

dim tpvelocity, Angleorig, Exitangle
dim origBall

Const teleport_out_level = 0.7
sub tpTrigger3_Hit()
  PlaySoundAt "EFX_teleport", tpTrigger3
  tpvelocity= BallVel(activeball) '( origBall.vely + origBall.velx ) / 2
  Angleorig = Atn2(activeball.vely,activeball.velx) * (180 / Pi) 'calc hit angle
  Exitangle = Angleorig + 90  'calc angle the ball should come from other side
' debug.print "TP3 arrival angle: " & Angleorig & " --> exit angle should be -> " & Exitangle
  activeball.destroyball
  vpmTimer.Addtimer 800,"PlaySoundAtLevelStatic (""EFX_teleport_out""), teleport_out_level, tpTrigger4:Kicker004.CreateBall:Kicker004.Kick exitangle, tpvelocity '"
' PlaySoundAtLevelStatic (""EFX_teleport_out""), 1, tpTrigger4
  'velocity parameter cannot be the same. It felt way too fast
  CounterBallSearch = -10
end sub

sub tpTrigger4_Hit()
  PlaySoundAt "EFX_teleport", tpTrigger4
  tpvelocity= BallVel(activeball) '( origBall.vely + origBall.velx ) / 2
  if tpvelocity < 10 then
    tpvelocity = 10 'give a little boost if it's too slow
  end if
  Angleorig = Atn2(activeball.vely,activeball.velx) * (180 / Pi) 'calc hit angle
  Exitangle = Angleorig + 90  'calc angle the ball should come from other side
' debug.print "TP4 arrival angle: " & Angleorig & " --> exit angle should be -> " & Exitangle
  activeball.destroyball
  'Kicker003.CreateBall:Kicker003.Kick exitangle, tpvelocity
  vpmTimer.Addtimer 800,"PlaySoundAtLevelStatic (""EFX_teleport_out""), teleport_out_level, tpTrigger3:Kicker003.CreateBall:Kicker003.Kick exitangle, tpvelocity '"
  'velocity parameter cannot be the same. It felt way too fast
  CounterBallSearch = -10
end sub

sub tpTrigger5_Hit()
  PlaySoundAt "EFX_teleport", tpTrigger5
  tpvelocity= BallVel(activeball) '( origBall.vely + origBall.velx ) / 2
  Angleorig = Atn2(activeball.vely,activeball.velx) * (180 / Pi) 'calc hit angle
  Exitangle = Angleorig + 90  'calc angle the ball should come from other side
' debug.print "TP5 arrival angle: " & Angleorig & " --> exit angle should be -> " & Exitangle
  activeball.destroyball
  vpmTimer.Addtimer 800,"PlaySoundAtLevelStatic (""EFX_teleport_out""), teleport_out_level, tpTrigger6:Kicker006.CreateBall:Kicker006.Kick exitangle, tpvelocity '"
  'Kicker006.CreateBall
  'Kicker006.Kick exitangle, tpvelocity 'velocity parameter cannot be the same. It felt way too fast
  CounterBallSearch = -10
end sub

sub tpTrigger6_Hit()
  PlaySoundAt "EFX_teleport", tpTrigger6
  tpvelocity= BallVel(activeball) '( origBall.vely + origBall.velx ) / 2
  if tpvelocity < 10 then
    tpvelocity = 10 'give a little boost if it's too slow
  end if
  Angleorig = Atn2(activeball.vely,activeball.velx) * (180 / Pi) 'calc hit angle
  Exitangle = Angleorig + 90  'calc angle the ball should come from other side
' debug.print "TP6 arrival angle: " & Angleorig & " --> exit angle should be -> " & Exitangle
  activeball.destroyball
  vpmTimer.Addtimer 800,"PlaySoundAtLevelStatic (""EFX_teleport_out""), teleport_out_level, tpTrigger5:Kicker005.CreateBall:Kicker005.Kick exitangle, tpvelocity '"
  'Kicker005.CreateBall
  'Kicker005.Kick exitangle, tpvelocity 'velocity parameter cannot be the same. It felt way too fast
  CounterBallSearch = -10
end sub


Dim UnderKicker,WizardPhase3DrainCount
WizardPhase3DrainCount = 3
Sub KickerDrain_Hit
  WriteToLog "KickerDrain_Hit", " "
    SwitchWasHit("UnderKicker")
  RandomSoundOutholeHit KickerDrain

  CounterBallSearch = -10
  'debug.print "kicker drain hit"
  If WizardPhase = 3 Then
    If cLightUnderPFShot1.state=1 and cLightUnderPFShot2.state=1 Then
      'mb wizard hatch close
      hatch false
'     solMagnetGrab 255           'enable grabber
'     SpaceWPlatform.timerenabled = true    'start multiball
      StopSong
      ScavengeKicker.timerinterval = 1590   'quessing the timing for the kick
      GaldorReset=true:GaldorQ1 = "eeeke"
      vpmtimer.addtimer 1500, "GaldorShade false '"
      AudioCallout "obey"
    Else
      WizardPhase3DrainCount = WizardPhase3DrainCount - 1
      if WizardPhase3DrainCount <= 0 Then
        'fail wizard
        hatch false
        RestoreLaneInsertStates
        WizardPhase3DrainCount = 3
        ScavengeKicker.timerinterval = 300
        EndWizard
      Elseif WizardPhase3DrainCount = 1 Then
        AudioCallout "final chance"
      end if
    End If
  Else
    hatch false
  End If
  If CurrentMission=5 Then StopMission
  KickerDrain.Destroyball
  underkicker=underkicker+1
  ScavengeKicker.timerenabled = true
  bFlippersEnabled = true
End Sub





'*********** Mystery Subway Kicker ************

Dim DMDMystery ' to display instead of score for 1 second and pick random mystery
Dim DMDMysteryReward  ' last displayd and blinked ?

Dim Scavengeball
Sub ScavengeKicker_Hit
  WriteToLog "ScavengeKicker_Hit", ""
  startB2S 5
  CounterBallSearch = 0
' SoundSaucerLock
  RandomSoundScoopRightEnter ScavengeKicker
  If Skillshot=1 or skillshot=2 Then skillshotoff
  If skillshot=3 Then
    Skillshotoff
    LS_skillshot2.Play SeqBlinking ,,20,5
    AwardSkillshot3
  End If

  gioff

  If cLightScavengeScoop.state>0 Then ' scavange
    LS_Scavenge.Play SeqBlinking,,5,13
    cLightScavengeScoop.state=0
    If cLightMystery1.state = 2 Then
      AudioCallout "mystery"
      PuPEvent(565)
      DOF 565, DOFPulse
      WriteToLog "ScavangeKicker","All Targets Pop Up"

      DTRaise 1
      DTRaise 2
      DTRaise 3
      RandomSoundDropTargetLeft(TargetScavenge2p)
      WriteToLog "ScavangeKicker","Light 1 Solid"
      cLightMystery1.state = 1
      DMDMystery=1
      SetDMDMysteryReward

    Elseif cLightMystery2.state = 2 Then
      AudioCallout "double scoring"
      PuPEvent(567)
      DOF 567, DOFPulse
      cLightMystery2.blinkinterval = 60
      StartDoubleScoring
      ScavengeKicker.timerenabled = true
      Scavengeball=1

    Elseif cLightMystery3.state = 2 Then
      AudioCallout "super spinners"
      PuPEvent(566)
      DOF 566, DOFPulse
      cLightMystery3.blinkinterval = 60
      StartSuperSpinners
      ScavengeKicker.timerenabled = true
      Scavengeball=1
    Else
      WriteToLog "ScavangeKicker","Why am I here. Kick the ball anyway"
      ScavengeKicker.timerenabled = true
      Scavengeball=1
    End If
    If cLightMystery1.state = 1 and cLightMystery2.state = 1 and cLightMystery3.state = 1 Then
      cLightMystery1.state = 0
      cLightMystery2.state = 0
      cLightMystery3.state = 0
      WriteToLog "ScavangeKicker","Lights Off"
    End If
  Else
    ScavengeKicker.timerenabled = true
    Scavengeball=1
  End If
    If Tilted Then Exit Sub
  AddScore score_ScavengeScoop
End Sub


Sub RotateScavangeLights
  if bWizardMode Then
    cLightMystery1.state = 0
    cLightMystery2.state = 0
    cLightMystery3.state = 0
  else
    If SpinnerMultiplier=1 And PointMultiplier=1 Then 'if not in the midst of a scavenge mode
      dim a, b : a = 0
      if cLightMystery1.state = 2 then a = 1
      if cLightMystery2.state = 2 then a = 2
      if cLightMystery3.state = 2 then a = 3
      if a > 0 Then
        ScavangeLights(a-1).state = 0
        for b = 1 to 3
          a = a + 1
          if a > 3 then a = 1
          if ScavangeLights(a-1).state = 0 Then
            ScavangeLights(a-1).state = 2
            WriteToLog "RotateScavangeLights ScavangeLights","Light " & a & " Blinking"
            exit sub
          end if
        next
      end if
    End If
  end if
End Sub


Sub SetDMDMysteryReward
  dim i

  i=9
  If MysteryEB(CurrentPlayer)=1 Then i = 8
  DMDMysteryReward=int(rnd(1)*i)+1

End Sub


Sub ScavengeKicker_timer
  Dim i
  i=ScavengeKicker.uservalue+1
  If i > 20 then i = 20
  ScavengeKicker.uservalue=i
  Select Case i
    Case 2 : LS_ScavengeSeq.Play SeqCircleOutOn,5
    Case 10 : LS_ScavengeSeq.Play SeqCircleOutOn,5
    Case 20 :
            WriteToLog "ScavengeKicker_timer","Scavengeball: " & Scavengeball & " UnderKicker: " & UnderKicker
      if WizardPhase <> 3 then gion
      If Scavengeball=1 Then
        Scavengeball=0
        'ScavengeKicker.createball:ScavengeKicker.Kick 230,55,1.5
        ScavengeKicker.Kick 230,55,1.5
        'SoundSaucerKick 1, ScavengeKicker
        RandomSoundScoopRightEjectHighVelocityDOF ScavengeKicker, 114, DOFPulse
        ScavengeKicker.uservalue=0
        ScavengeKicker.timerinterval = 150
        If underkicker=0 Then
          ScavengeKicker.timerenabled = false
        End If
      Elseif UnderKicker>0 Then
        UnderKicker=UnderKicker-1
        ScavengeKicker.Createball
        ScavengeKicker.Kick 230,55,1.5
'       SoundSaucerKick 1, ScavengeKicker
        RandomSoundScoopRightEjectHighVelocityDOF ScavengeKicker, 114, DOFPulse
        if bWizardModeUnderGame then        'under wizard game success
          'asdasdasd
          StartPhase4
        end if
        ScavengeKicker.uservalue=0
        ScavengeKicker.timerinterval = 150
        If underkicker=0 Then
          ScavengeKicker.timerenabled = false
        End If
      Else 'case when scavenge collected already?
        ScavengeKicker.Kick 230,55,1.5
'       SoundSaucerKick 1, ScavengeKicker
        RandomSoundScoopRightEjectHighVelocityDOF ScavengeKicker, 114, DOFPulse
        ScavengeKicker.uservalue=0
        ScavengeKicker.timerinterval = 150
        If underkicker=0 Then
          ScavengeKicker.timerenabled = false
        End If
      End If
    case Else :
      Flash2 true
      ScavengeKicker.timerinterval = ScavengeKicker.timerinterval * 0.85 + 5
      'debug.print ScavengeKicker.timerinterval
  End Select
End Sub


Sub RewardMystery
  WriteToLog "RewardMystery", ""
    If Tilted Then Exit Sub
  Select Case DMDMysteryReward   ' Fixing  EB special Litelock ?
    Case 1 : Addscore 10000000
    PuPEvent(570)
    DOF 570, DOFPulse
    Case 2 : Addscore  8000000
    PuPEvent(569)
    DOF 569, DOFPulse
    Case 3 : Addscore  5000000
    PuPEvent(568)
    DOF 568, DOFPulse
    Case 4 : Addscore 15000000
    PuPEvent(571)
    DOF 571, DOFPulse
    Case 5 : Addscore 17000000
    PuPEvent(572)
    DOF 572, DOFPulse
    Case 6 : Addscore 20000000
    PuPEvent(573)
    DOF 573, DOFPulse
    Case 7 : Addscore 20000000
    PuPEvent(573)
    DOF 573, DOFPulse
    Case 8 : Addscore 20000000
    PuPEvent(573)
    DOF 573, DOFPulse
    Case 9: cLightCollectExtraBall.state=2   ' EB
        AudioCallout "extra ball lit"
        PuPEvent(574)
        DOF 574, DOFPulse
        ' fixing sound blinks ..

  End Select
End Sub



'************ ship cannon kicker **************

sub ShipCannonKicker_Hit
  WriteToLog "ShipCannonKicker_Hit", " "
  AddScore score_HarpoonHit
  ShipCannonKicker.timerenabled = true

  'Add a ball if the coniditions are right
  'If AddBallAvailable = 1 and not (cLightMimaMB.state = 1 and cLightCoreyMB.state = 1 and cLightTracyMB.state = 1) Then
  If AddBallAvailable = 1 and bMultiBallMode and Not bWizardMode Then
    WriteToLog "ShipCannonKicker_Hit", "addmultiball 1"
    AddBallAvailable = 0
    AudioCallout "ball added"
    AddMultiball 1
  End If

  If cLightCollectExtraBall.state=2 Then
    WriteToLog "ShipCannonKicker_Hit", "AwardExtraBall"
    cLightCollectExtraBall.blinkinterval=40 ' superspoeed blink until kicked out

    MysteryEB(CurrentPlayer)=1
    ' fixing add blinking and some sound
    AudioCallout "extra ball"
    AwardExtraBall
'   cLightShootAgain1.state=1 ' fixing add more blinking lightsequencer
'   cLightShootAgain2.state=1
    ShipCannonKicker.timerinterval=1500
  End If

  EnableBallSaver 10            'ball saver added always
  CounterBallSearch = 0

end Sub

Sub ShipCannonKicker_timer
  cLightCollectExtraBall.blinkinterval=125

  ' more eb ?
  ' if then else here for other extrtaballs  to not turn off if they are available   fixing
  cLightCollectExtraBall.state=0

  ShipCannonKicker.timerinterval=10
  'ShipCannonKicker.Kick 187,60,25
  'ShipCannonKicker.Kick 184,50,10  'new value since minor wall change broke it
  ShipCannonKicker.Kick 185,60,10    'new value because previous one drained too easy
' DOF 106,2
    playsound "EFX_shiploop_end",0,CalloutVol,0,0,1,1,1
' SoundSaucerKick 1, ShipCannonKicker
  RandomSoundScoopRightEjectHighVelocityDOF ShipCannonKicker, 106 , DOFPulse
  ShipCannonKicker.timerenabled = false
end sub


'**************  gun diverter  **********

DiverterOpen.isdropped = true
Diverter.isdropped = false

Sub swOpenGunTrigger_hit
  WriteToLog "swOpenGunTrigger_hit", " "
  SetLazerReady
  TracyAttention.Enabled=True
  SwitchWasHit("swOpenGunTrigger")
end sub

Sub swCloseGunTrigger_hit
  WriteToLog "swCloseGunTrigger_hit", " "
  CloseGun
  If CurrentMission = 1 Then
    LightBladerunnerShot
  Else
'   if Not bWizardMode then
      LightGunTarget
'   end if
  end if
  WireRampOff
  WireRampOn False 'Wire Ramp
  PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"), RightPlasticRampHitsSoundLevel, swCloseGunTrigger' : Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"
  SwitchWasHit("swCloseGunTrigger")
End Sub

dim DiverterTarget
Sub CloseGun
  WriteToLog "CloseGun", " "
  DiverterOpen.isdropped = True
  Diverter.isdropped = False
  'prim to -33
  DiverterTarget = -31
  DiverterOpen.timerenabled = True
  PlaySoundAt "Diverter_UP", Screw002
  cLightLoadLazer.state = 0
End Sub



Sub SetLazerReady
  WriteToLog "SetLazerReady", ""
  DiverterOpen.isdropped = false
  If CurrentMission = 1 Then
    cLightLoadLazer.state = 2
    LS_Flash5.UpdateInterval = 700
    LS_Flash5.Play SeqBlinking ,,10000000,10
  Else
    cLightLoadLazer.state = 1

  End If
  Diverter.isdropped = true
  'prim to -7
  DiverterTarget = -7
  DiverterOpen.timerenabled = True
  PlaySoundAt "Diverter_DOWN", Screw002
End Sub

dim DivWobble : DivWobble = 0
sub DiverterOpen_timer
  if abs(divp.roty-DiverterTarget) < 2 Or DivWobble > 6 then
    DivP.roty = DiverterTarget
    DivWobble = 0
    DiverterOpen.timerenabled = False
  elseif DivP.roty < DiverterTarget Then
    '-33 to -7
    DivP.roty = 0.9*DivP.roty + 9
  elseif DivP.roty > DiverterTarget Then
    '-7 to -33
    DivP.roty = 0.9*DivP.roty - 9
  end if
  DivWobble = DivWobble + 1
end sub



'************ VUK 1 **************

Sub VUK1_Hit
  WriteToLog "VUK1_Hit", ""
  dim KickDelay
  KickDelay = 1300
  If CurrentMission=1 Then showMissionVideo=2

' SoundSaucerLock
  RandomSoundScoopLeftEnter VUK1
  AddScore score_VUKGun
  LS_Flash5.StopPlay: LS_Flash5.UpdateInterval = 200: LS_Flash5.Play SeqBlinking ,,3,10

  'Start a mission if the conitions are right
  if SwitchRecentlyHit("swOpenGunTrigger") then
    'We game here via hidden route, not starting a mission
  else
    If cLightMissionStart1.state=2 Then
      gioff
      LS_AllLightsAndFlashers.Play SeqBlinking ,,5,2
      LS_AllLightsAndFlashers.Play SeqUpOn, 10, 1

      cLightMissionStart1.state=1
      if Not HasPuP then        'flex
        Select Case MissionSelect
          Case 1: KickDelay = 3000
          Case 2: KickDelay = 2500
          Case 3: KickDelay = 2500
          Case 4: KickDelay = 2500
          Case 5: KickDelay = 2500
          Case 6: KickDelay = 2500
          Case 7: KickDelay = 2500
        End Select
      Else                'pup
        Select Case MissionSelect
          Case 1: KickDelay = 6500
          Case 2: KickDelay = 6500
          Case 3: KickDelay = 6500
          Case 4: KickDelay = 6500
          Case 5: KickDelay = 6500
          Case 6: KickDelay = 6500
          Case 7: KickDelay = 6500
        End Select
      end if
      Mission_Begin
    Else
      gioff
    End If
  end if

  VUK1.TimerInterval = KickDelay
  VUK1.TimerEnabled = True
  CounterBallSearch = 0

End Sub

Sub VUK1_Timer
  if (WizardPhase < 1 Or WizardPhase > 2) And CurrentMission <> 5 then
    'debug.print "---> VUK 1 GI on commanded"
    gion
  Else
    'debug.print "---> VUK 1 Mission 5 or wizard ongoing no GION"
  end if
  KickVUK1
  VUK1.TimerEnabled = False
end sub


Sub KickVUK1
  WriteToLog "KickVUK1", "Ball kicked"
' DOF 113,2
  startB2S 1
  VUK1.kickz 180,20,12,180 '190
' SoundSaucerKick 1, VUK1
  RandomSoundScoopLeftEjectNormalVelocityDOF VUK1, 113, DOFPulse
  wVUK1Hole.timerenabled = True
  'If CurrentMission <> 0 Then gion
End Sub

dim VUK1step:VUK1step=0
Sub wVUK1Hole_timer
  Select Case VUK1step
    Case 0: pVUK1.z = 12
    Case 1: pVUK1.z = 8
        Case 2: pVUK1.z = 4
        Case 3: pVUK1.z = 0
        Case 4: pVUK1.z = -5
        wVUK1Hole.TimerEnabled = False
        VUK1step = 0
    End Select
  VUK1step = VUK1step + 1
end sub



'*************** VUK 2 *****************

'Shot3
Sub VUK2_Hit()
  dim KickDelay
  PuPEvent(583)
  DOF 583, DOFPulse
  KickDelay = 200
  WriteToLog "VKU2_Hit", "skillshot " & Skillshot
  RandomSoundScoopRightEnter VUK2
' SoundSaucerLock

  If Skillshot=1 or skillshot=2 Then skillshotoff
  If skillshot=3 Then
    Skillshotoff
    LS_skillshot2.Play SeqBlinking ,,20,5
    AwardSkillshot2
  End If

  checkWizardStart  'additional check for wizard start if it is missed

  gioff

  CheckBladerunnerShot 3
  CheckChaseShot 3
  CheckBetrayalMissionShot 3
  CheckTracyMBSuperJackpot
  CheckWizardGaldorShot
  If Not Tilted Then AddScore score_VUKSpinner

  If cLightCorey3.state = 2 and not bMultiballMode Then
    cLightCorey3.state = 1
    CheckCoreyMB
  End If
  If cLightTracy3.state = 2 and not bMultiballMode Then
    cLightTracy3.state = 1
    CheckTracyMB
  End If

  If bMultiballReady and not bMultiballMode Then
        KickDelay = 5000
        LS_Flash4.StopPlay
        PlaySound "EFX_landing",0,CalloutVol,0,0,1,1,1
    PuPEvent(585)
    DOF 585, DOFPulse
        VUK2.uservalue = 113            '~4s
        'vpmTimer.Addtimer KickDelay,"StartMBs '"
    Else
        Vuk2.uservalue = 138
    end If
  CounterBallSearch = 0
    WriteToLog "VKU2_Hit", "KickDelay= " & KickDelay
    VUK2.TimerInterval = 200
    VUK2.TimerEnabled = True
End Sub

Sub VUK2_Timer()
    Dim i
    i=VUK2.uservalue + 1
    If i < 1 Then i = 1
    VUK2.uservalue = i

    Select Case i
        Case 1:    Flash3 True
        Case 4:    Flash1 True
        Case 8:    Flash7 True
        Case 14: Flash6 True
        Case 20,21 :
            WriteToLog "VUK2_Timer", "Ball kicked"
            DOF 114,2
            startB2S 1
            'VUK2.kickz 180,40,12,145
            VUK2.kickz 180,70,1,30
            gion
      If bMultiballReady and not bMultiballMode Then StartMBs   'moved here as vpmtimer may not be reliable
'            SoundSaucerKick 1, VUK2
      RandomSoundScoopLeftEjectNormalVelocityDOF VUK2, 114, DOFPulse
      wVUK2Hole.TimerEnabled = True
            VUK2.uservalue=0
            VUK2.timerinterval = 200
            VUK2.timerenabled = false
        case 2,3,5,6,7,9,10,11,12,13,15,16,17,18,19 :
            Flash4 true
            VUK2.timerinterval = VUK2.timerinterval * 0.85 + 5
        case 140 : VUK2.TimerInterval = 200 : VUK2.uservalue=0
    End Select
End Sub

dim VUK2step:VUK2step=0
Sub wVUK2Hole_timer
  Select Case VUK2step
    Case 0: pVUK2.z = 12
    Case 1: pVUK2.z = 8
        Case 2: pVUK2.z = 4
        Case 3: pVUK2.z = 0
        Case 4: pVUK2.z = -5
        wVUK2Hole.TimerEnabled = False
        VUK2step = 0
    End Select
  VUK2step = VUK2step + 1
end sub

'*************** Grabber Magnet *****************

dim magRestartRequested : magRestartRequested = 0
Sub solMagnetGrab(value)
  if WizardPhase = 4 then
    if Value > 0 then     'limited to only work when wizard is on.
      GrabMag.Strength = 25 * value / 255
      GrabMag.MagnetOn = 1
      MiMag.Strength = 0
      magRestartRequested = 0
      GrabMagnetRelease.enabled = true
    Else
      GrabMag.MagnetOn = 0
      GrabMag.Strength = 0
      magRestartRequested = 0
      vpmtimer.addtimer 200, "GrabMag.MagnetOn = 1 '"
      vpmtimer.addtimer 500, "GrabMag.MagnetOn = 0 '"
    end If
  Else
    GrabMag.MagnetOn = 0    'any call that comes here will disable the grabmag
    GrabMag.Strength = 0
    magRestartRequested = 0
  end if
end sub

dim GrabberVideo, gballcount : GrabberVideo = 0
sub GrabMagnetRelease_timer
  dim GrabbedBalls, BOTnow, BOTnowCount

  BOTnow = getballs
  BOTnowCount = Ubound(BOTnow)
'
' If wVortexBlocker.isdropped Then
'   BOTnowCount = BOTnowCount - 1 'one captive balls
''  Else
''    BOTnowCount = BOTnowCount - 2 'two captive balls
' end if

  GrabbedBalls = GrabMag.Balls
  gballcount = ubound(GrabbedBalls) + 1

  if Not WizardPhase = 4 Then     'terminate grabber if wizard over
    solMagnetGrab 0
    MiMag.Strength = 15
    GrabMagnetRelease.enabled = false
    exit sub
  end if

  if gballcount > 9 And GrabberVideo = 0 then 'show pyramid
    GrabberVideo = 1
    StopVideosPlaying
    showMissionVideo=70
  end if

  if gballcount + 0 >= BOTnowCount    Then  '0 - one ball in play, 1 - 2 balls in play
    solMagnetGrab 0
    StopVideosPlaying
    showMissionVideo=68
    if magRestartRequested = 0 then
      'debug.print "***** grabmag balls released: " & gballcount & " / " & BOTnowCount
      vpmtimer.addtimer 1500, "solMagnetGrab 255 : GrabberVideo = 0 '"
      magRestartRequested = 1
    end if
  Else
    'debug.print "grabmag balls: " & gballcount
    if WizardPhase = 4 And GrabMag.MagnetOn = 0 Then
      if magRestartRequested = 0 then
        'debug.print "*** restart grabmag"
        vpmtimer.addtimer 1000, "solMagnetGrab 255 : GrabberVideo = 0 '"
        magRestartRequested = 1
      end if
    end If
  end if

end sub

'*************** Spinner *****************
dim spinEFXcount : spinEFXcount = 0

Sub Spinner1_Spin
  DOF 128,2
  SoundSpinner()

    If Tilted Then Exit Sub

  if spinEFXcount = 0 Then
    spinEFXcount = 1
  else
    Spinner1.timerenabled = 1 'timer for resetting sounds to start from first
  end if

  if spinEFXcount = 1 Then
    stopsound("EFX_spin04")
  Else
    stopsound("EFX_spin0" & spinEFXcount - 1)
  end if

  playsound "EFX_spin0" & spinEFXcount,0,CalloutVol,0,0,1,1,1
  spinEFXcount = spinEFXcount + 1
  if spinEFXcount > 4 then spinEFXcount = 1

  Addscore score_spinnerspin*SpinnerMultiplier
  if SpinnerMultiplier > 1 Then
    PoltergeistTracy 1
    AddBonus bonus_SuperSpinners
  end if

End Sub

sub Spinner1_timer
  spinEFXcount = 0
  Spinner1.timerenabled = 0
end sub


'***************  Hatch Cover Opacity *****************

dim hatch_coverop
hatch_coverop = 1

sub HatchCoverOpacity(aLvl)
  'Playfield_cover
  'UpdateMaterial(string,       float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity, OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive, float elasticity, float elasticityFalloff, float friction, float scatterAngle)
  UpdateMaterial "Playfield_cover"  ,0,0,0,0,0,0,aLvl,RGB(181,181,181),0,0,False,True,0,0,0,0
  if aLvl = 1 Then
    pHatchCover.material = "Playfield"
  else
    pHatchCover.material = "Playfield_cover"
  end if
end sub


'***************  Hatch Open *****************

dim hatch_tangle, hatch_angle
dim hatch_tmove, hatch_move
dim hatch_tz, hatch_z


triang1.blenddisablelighting = 20
triang2.blenddisablelighting = 20
triang3.blenddisablelighting = 20
triang4.blenddisablelighting = 20
triang5.blenddisablelighting = 20

Dim triang1x, triang1y, triang2x ,triang2y, triang3x, triang3y, triang4x, triang4y, triang5x, triang5y

dim hatch_sink_step, hatch_move_step, hatch_rotate_step
hatch_sink_step = 1
hatch_move_step = 2
hatch_rotate_step = 9

hatch_z = 0 'triang1.z
hatch_move = 0

triang1x = triang1.x:triang1y = triang1.y:triang2x = triang2.x:triang2y = triang2.y:triang3x = triang3.x:triang3y = triang3.y
triang4x = triang4.x:triang4y = triang4.y:triang5x = triang5.x:triang5y = triang5.y

dim hatch_phase, hatch_open
hatch_phase = 1

if VersionMinor => 8 Then     'hatch is too bright if version 10.8 in use
' msgbox "10.8"
  Hatchflasher001.AddBlend = false
  Hatchflasher001.opacity = 90
  gi003.intensity = 0
end if

sub hatch(enabled)
  if enabled then
    triang1.blenddisablelighting = 100
    triang2.blenddisablelighting = 100
    triang3.blenddisablelighting = 100
    triang4.blenddisablelighting = 100
    triang5.blenddisablelighting = 100
    hatch_tz = -10
    hatch_tangle = 54
    hatch_tmove = 10
    hatch_open = true
    bPortalFlippersEnabled = True
    playsound "EFX_ship_cannons_turn",0,1*CalloutVol,0,0,11000,0,0,0

    dim xx
    for each xx in underPFthings
      if typename(xx) = "Wall" Then xx.SideVisible = True
      xx.visible=True
    Next
    Hatchflasher001.visible=0
  else
    triang1.blenddisablelighting = 10
    triang2.blenddisablelighting = 10
    triang3.blenddisablelighting = 10
    triang4.blenddisablelighting = 10
    triang5.blenddisablelighting = 10
    pUnderPF.blenddisablelighting = 0
    hatch_tz = 0
    hatch_tangle = 0
    hatch_tmove = 0
    hatch_open = false
    bPortalFlippersEnabled = False
    'playsound "EFX_ship_cannons_turn"
  end if
  hatchtimer.Interval = 30
  hatchtimer.enabled = true
end sub

Sub Hatchtrigger_Hit
  If hatch_open Then
    activeball.z = -31
    activeball.velz=-10
  End If
End Sub


sub hatchtimer_timer

  'fade in/out cover
  if hatch_phase = 1 then
    If hatch_open Then
      hatch_coverop = hatch_coverop - 0.01
      If hatch_coverop <= 0 Then
        hatch_coverop=0
        hatch_phase = 2
      End If
    Else
      hatch_coverop = hatch_coverop + 0.01
      If hatch_coverop >=  1 Then
        hatch_coverop=1
        hatchtimer.enabled = false
        dim xx
        for each xx in underPFthings
          if typename(xx) = "Wall" Then xx.SideVisible = False
          xx.visible=False
        Next
        For each xx in underPFclights : xx.state=False : Next

        Hatchflasher001.visible=True

      End If
    End If
    HatchCoverOpacity hatch_coverop
  end if

' 'sink
  if hatch_phase = 2 then
'   debug.print "sink: " & hatch_z & " target: " & hatch_tz
    if hatch_z < hatch_tz Then
      hatch_z = hatch_z + hatch_sink_step
    elseif hatch_z > hatch_tz Then
      hatch_z = hatch_z - hatch_sink_step
    end if
    if round(abs(hatch_z),2) = round(abs(hatch_tz),2) then
      hatch_z = hatch_tz
      if hatch_open Then
'       debug.print "sunk"
        hatch_phase = 3
      else
'       debug.print "succesfully closed"
      ' UnderDropper.collidable = true
        hatch_phase = 1
      end if
    end If
    triang_sink(hatch_z)
  end if

  if hatch_phase = 3 then
    'move
'   debug.print "move: " & hatch_move & " target: " & hatch_tmove
    if hatch_move < hatch_tmove Then
      hatch_move = hatch_move + hatch_move_step
    elseif hatch_move > hatch_tmove Then
      hatch_move = hatch_move - hatch_move_step
    end if
    if round(abs(hatch_move),2) = round(abs(hatch_tmove),2) then
      hatch_move = hatch_tmove
      if hatch_open Then
'       debug.print "moved open"
        hatch_phase = 4
      Else
'       debug.print "moved close"
        hatch_phase = 2
      end if
    end If
    triang_move(hatch_move)
  end if

  if hatch_phase = 4 then
    'rotate
    if hatch_angle < hatch_tangle Then
      hatch_angle = hatch_angle + hatch_rotate_step
    Elseif hatch_angle > hatch_tangle Then
      hatch_angle = hatch_angle - hatch_rotate_step
    end if
    if round(abs(hatch_angle),2) = round(abs(hatch_tangle),2) then
      hatch_angle = hatch_tangle
      if hatch_open Then
'       debug.print "succesfully opened"
        triang1.blenddisablelighting = 10
        triang2.blenddisablelighting = 10
        triang3.blenddisablelighting = 10
        triang4.blenddisablelighting = 10
        triang5.blenddisablelighting = 10
        pUnderPF.blenddisablelighting = 5
        'UnderDropper.isdropped = true
        'UnderDropper.collidable = false
        if CurrentMission = 5 Then    ''adding these here for better timing
          LightPortalRoomShots
        elseif wizardphase = 3 Then
          LightPortalRoomShotsLeft
        end if
        hatchtimer.enabled = false
      else
'       debug.print "rotated close"
        hatch_phase = 3
      end if
    end if
    triang_rotate(hatch_angle)
  end if
end sub


sub triang_move(avalue)

  'debug.print avalue

  const triang1angle = 0
  const triang2angle = 72
  const triang3angle = 144
  const triang4angle = 216
  const triang5angle = 288

  triang1.x = Cos(Radians(triang1angle+90)) * avalue + triang1x
  triang1.y = Sin(Radians(triang1angle+90)) * avalue + triang1y

  triang2.x = Cos(Radians(triang2angle+90)) * avalue + triang2x
  triang2.y = Sin(Radians(triang2angle+90)) * avalue + triang2y

  triang3.x = Cos(Radians(triang3angle+90)) * avalue + triang3x
  triang3.y = Sin(Radians(triang3angle+90)) * avalue + triang3y

  triang4.x = Cos(Radians(triang4angle+90)) * avalue + triang4x
  triang4.y = Sin(Radians(triang4angle+90)) * avalue + triang4y

  triang5.x = Cos(Radians(triang5angle+90)) * avalue + triang5x
  triang5.y = Sin(Radians(triang5angle+90)) * avalue + triang5y

end sub

sub triang_sink(avalue)
  triang1.z = avalue-0.1:triang2.z = avalue-0.1:triang3.z = avalue-0.1:triang4.z = avalue-0.1:triang5.z = avalue-0.1
end sub

sub triang_rotate(avalue)
  triang1.objrotz = avalue:triang2.objrotz = avalue:triang3.objrotz = avalue:triang4.objrotz = avalue:triang5.objrotz = avalue
end sub





'*********** Mima Loop Standup Ring **************

Sub LoopStandUp_Init
  LoopStandUp.isdropped = true
  PlaySoundAt "Diverter_DOWN", Screw015
End Sub

dim swLoopUpTriggerCounter: swLoopUpTriggerCounter = 0
dim KeepBallInLoop : KeepBallInLoop = 0
LoopStandUp.blenddisablelighting = 0.05

Sub swLoopUpTrigger_Hit
  if swLoopUpTriggerCounter < 3 then
    if WizardPhase <> 1 then
      LoopStandUp.TimerInterval = 390 + RndInt(1,20)  '800'initial timeout for wall drop
      LoopStandUp.TimerEnabled = 1
    end if
    if abs(activeball.velx) > 8 And abs(activeball.velx) < 40 then
      activeball.velz = 0
      activeball.velx = activeball.velx * (1.9 + (RndInt(2,200)/1000)) '1.8
'     Debug.print "--> speedup to: " & activeball.velx
    Else
'     Debug.print "--> no speedup"

    end if
    swLoopUpTriggerCounter = swLoopUpTriggerCounter + 1
    PuPEvent(579)
    DOF 579, DOFPulse
    if WizardPhase = 1 Then swLoopUpTriggerCounter = 4        'let it do only one speedup
  end if

  if KeepBallInLoop = 0 Then  'don't let wall raise, if has just dropped
'   if abs(activeball.velx) > 25 and swLoopUpTriggerCounter < 2 then
    if abs(activeball.velx) > 25 and LoopStandUp.isdropped then
'     Debug.print "sound: " & activeball.velx & " counter: " & swLoopUpTriggerCounter
      playsound "EFX_shiploop_start",0,CalloutVol,0,0,1,1,1
    end if
    LoopStandUp.isdropped = false
    PlaySoundAt "Diverter_UP", Screw015
    TriggerShipLight001.enabled = true
    TriggerShipLight002.enabled = true
    TriggerShipLight004.enabled = true
    TriggerShipLight005.enabled = true
    KeepBallInLoop = 1
  End If


  if abs(activeball.velx) < 25 and WizardPhase <> 1 then ' this will give some time to aim the ball
'   Debug.print "--> speed fading"
    LoopStandUp.TimerInterval = 590 + RndInt(1,20) '600
    LoopStandUp.TimerEnabled = 1
  end if

  if WizardPhase = 1 Then     'lock the ball for wizard start
    StartTracyChange      'this gets many calls
    LoopStandUp.TimerInterval = 6000 '600
    LoopStandUp.TimerEnabled = 1
    CounterBallSearch = -10
  end if

  'adding mission time for loop shots. unless in mission 2 resurrection
  if CurrentMission<>0 And MissionTimeCurrent > 2 And CurrentMission<>2 Then
    'debug.print "adding time: " & MissionTimeCurrent
    MissionTimeCurrent = MissionTimeCurrent - 2
    playsound "EFX_effect3",0,CalloutVol,0,0,1,1,1
  end if

  CheckMimaMBJackpot
  EOB_MimaLoops = EOB_MimaLoops + 1
  LightShip003.state = 1
  if Not bWizardMode then
    cLightMimaTarget.state = 2
  end if

    If Tilted Then Exit Sub
  AddScore score_RingLoop
  SwitchWasHit("swLoopUpTrigger")
End Sub

Sub swLoopUpTrigger_unhit
  LightShip003.state = 0
end Sub

sub TriggerShipLight001_hit:LightShip001.state = 1:end Sub
sub TriggerShipLight001_unhit:LightShip001.state = 0:end Sub
sub TriggerShipLight002_hit
  LightShip002.state = 1
end Sub
sub TriggerShipLight002_unhit:LightShip002.state = 0:end Sub
sub TriggerShipLight004_hit
  LightShip004.state = 1
end Sub
sub TriggerShipLight004_unhit:LightShip004.state = 0:end Sub
sub TriggerShipLight005_hit:LightShip005.state = 1:end Sub
sub TriggerShipLight005_unhit:LightShip005.state = 0:end Sub

Sub LoopStandUp_Timer
  LoopStandUp.TimerEnabled = false
  DropLoopWall true
End Sub

Sub DropLoopWall(Enabled)
  If Enabled Then
    if not LoopStandUp.isdropped Then
      if LoopStandUp.TimerEnabled then
'       debug.print "wall dropped by flip"
        stopsound "EFX_shiploop_start"
        playsound "EFX_shiploop_end",0,CalloutVol,0,0,1,1,1
        LoopStandUp.TimerEnabled = false
      end if
    end if
    cLightMimaTarget.state = 0
    LoopStandUp.isdropped = true
    PlaySoundAt "Diverter_DOWN", Screw015
    TriggerShipLight001.enabled = false
    TriggerShipLight002.enabled = false
    TriggerShipLight004.enabled = false
    TriggerShipLight005.enabled = false
    LightShip001.state = 0
    LightShip002.state = 0
    LightShip003.state = 0
    LightShip004.state = 0
    LightShip005.state = 0
'   debug.print "wall dropped"
    if WizardPhase = 1 And cTracyChange.state = 1 Then
      StartPhase2
    end if
    vpmtimer.addtimer 1000, "KeepBallInLoop = 0 '"  'delayed so that wall won't raise up too fast after previous drop
    vpmtimer.addtimer 1025, "swLoopUpTriggerCounter = 0 '"  'delayed so that speedup won't happen too often
  end if
End Sub

LightShip001.state = 0
LightShip002.state = 0
LightShip003.state = 0
LightShip004.state = 0
LightShip005.state = 0

'*********** Mima Loop Flasher **************

dim flasherUndershipGain
FlasherUndership.opacity = 0
FlasherUndership.visible = false
FlasherUndership.y = 510
FlasherUndership.x = 270

Sub FlashUndership(enabled)
  if enabled then
    flasherUndershipGain = 1
    FlasherUndership.visible = true
    FlasherUndership.timerenabled = True
    cLightFlashUndership.state = 0
    DOF 135,2
  end if
End Sub

Sub FlasherUndership_Timer
  flasherUndershipGain = flasherUndershipGain * 0.80 - 0.01
  If flasherUndershipGain <= 0 Then
    flasherUndershipGain = 0
    FlasherUndership.visible = false
    FlasherUndership.timerenabled = False
  End If
  FlasherUndership.opacity = 100 * flasherUndershipGain^1
End Sub



'**********************************************************************************************************
'  Orbit and Ramp Switches
'**********************************************************************************************************

Dim SwitchHitTimes(50)
Const SwitchHitRecentlyTime = 1500  'this threshold determines if a swtich was recently hit

Sub SwitchWasHit(swName)
  Select Case swName
    Case "swLeftOrbTrigger1"  : SwitchHitTimes(1) = gametime
    Case "swLeftOrbTrigger2"  : SwitchHitTimes(2) = gametime
    Case "swRightOrbTrigger1"   : SwitchHitTimes(3) = gametime
    Case "swRightOrbTrigger2"   : SwitchHitTimes(4) = gametime
    Case "swSkyRamp"      : SwitchHitTimes(5) = gametime
    Case "swMainRamp"       : SwitchHitTimes(6) = gametime
    Case "swRampRoll1"      : SwitchHitTimes(7) = gametime
    Case "swRampRoll2"      : SwitchHitTimes(8) = gametime
    Case "swRampRoll3"      : SwitchHitTimes(9) = gametime
    Case "swRampRoll4"      : SwitchHitTimes(10) = gametime
    Case "swRampRoll5"      : SwitchHitTimes(11) = gametime
    Case "swRampRoll6"      : SwitchHitTimes(12) = gametime
    Case "swRampRoll7"      : SwitchHitTimes(13) = gametime
    Case "swRampRoll9"      : SwitchHitTimes(14) = gametime
    Case "swRampRoll10"     : SwitchHitTimes(15) = gametime
    Case "swTopRightLane"     : SwitchHitTimes(16) = gametime
    Case "swTopLeftLane"    : SwitchHitTimes(17) = gametime
    Case "swPlungerRest"    : SwitchHitTimes(18) = gametime
    Case "swLeftOutLane"    : SwitchHitTimes(19) = gametime
    Case "swLeftInLane"     : SwitchHitTimes(20) = gametime
    Case "swRightInLane"    : SwitchHitTimes(21) = gametime
    Case "swRightOutLane"     : SwitchHitTimes(22) = gametime
    Case "swOpenGunTrigger"   : SwitchHitTimes(23) = gametime
    Case "swCloseGunTrigger"  : SwitchHitTimes(24) = gametime
    Case "swLoopUpTrigger"    : SwitchHitTimes(25) = gametime
    Case "RightSling"       : SwitchHitTimes(26) = gametime
    Case "LeftSling"      : SwitchHitTimes(27) = gametime
        Case "TargetVascan"     : SwitchHitTimes(28) = gametime
    Case "UnderPF"        : SwitchHitTimes(29) = gametime
    Case "Drain"        : SwitchHitTimes(30) = gametime
    Case "UnderKicker"      : SwitchHitTimes(31) = gametime
  End Select
' debug.print swName
  CounterBallSearch = 0
End Sub

Function SwitchRecentlyHit(swName)
  dim delta : delta = SwitchHitRecentlyTime + 1
  Select Case swName
    Case "swLeftOrbTrigger1"  : delta = gametime - SwitchHitTimes(1)
    Case "swLeftOrbTrigger2"  : delta = gametime - SwitchHitTimes(2)
    Case "swRightOrbTrigger1"   : delta = gametime - SwitchHitTimes(3)
    Case "swRightOrbTrigger2"   : delta = gametime - SwitchHitTimes(4)
    Case "swSkyRamp"      : delta = gametime - SwitchHitTimes(5)
    Case "swMainRamp"       : delta = gametime - SwitchHitTimes(6)
    Case "swRampRoll1"      : delta = gametime - SwitchHitTimes(7)
    Case "swRampRoll2"      : delta = gametime - SwitchHitTimes(8)
    Case "swRampRoll3"      : delta = gametime - SwitchHitTimes(9)
    Case "swRampRoll4"      : delta = gametime - SwitchHitTimes(10)
    Case "swRampRoll5"      : delta = gametime - SwitchHitTimes(11)
    Case "swRampRoll6"      : delta = gametime - SwitchHitTimes(12)
    Case "swRampRoll7"      : delta = gametime - SwitchHitTimes(13)
    Case "swRampRoll9"      : delta = gametime - SwitchHitTimes(14)
    Case "swRampRoll10"     : delta = gametime - SwitchHitTimes(15)
    Case "swTopRightLane"     : delta = gametime - SwitchHitTimes(16)
    Case "swTopLeftLane"    : delta = gametime - SwitchHitTimes(17)
    Case "swPlungerRest"    : delta = gametime - SwitchHitTimes(18)
    Case "swLeftOutLane"    : delta = gametime - SwitchHitTimes(19)
    Case "swLeftInLane"     : delta = gametime - SwitchHitTimes(20)
    Case "swRightInLane"    : delta = gametime - SwitchHitTimes(21)
    Case "swRightOutLane"     : delta = gametime - SwitchHitTimes(22)
    Case "swOpenGunTrigger"   : delta = gametime - SwitchHitTimes(23)
    Case "swCloseGunTrigger"  : delta = gametime - SwitchHitTimes(24)
    Case "swLoopUpTrigger"    : delta = gametime - SwitchHitTimes(25)
    Case "RightSling"       : delta = gametime - SwitchHitTimes(26)
    Case "LeftSling"      : delta = gametime - SwitchHitTimes(27)
    Case "TargetVascan"     : delta = gametime - SwitchHitTimes(28)
    Case "UnderPF"        : delta = gametime - SwitchHitTimes(29)
    Case "Drain"        : delta = gametime - SwitchHitTimes(30)
    Case "UnderKicker"      : delta = gametime - SwitchHitTimes(31)
  End Select
  If delta < SwitchHitRecentlyTime Then
    SwitchRecentlyHit = True
  Else
    SwitchRecentlyHit = False
  End If
End Function


Sub swLeftOrbTrigger1_hit
  WriteToLog "swLeftOrbTrigger1_hit", ""
  SwitchWasHit("swLeftOrbTrigger1")
' TracyAttention.Enabled=True
  If Tilted = false then AddScore score_LoopEntrance
End Sub

'Shot1
dim gate1status
dim gate4status

Sub swLeftOrbTrigger2_hit
    If Tilted Then Exit Sub
  WriteToLog "swLeftOrbTrigger2_hit", ""

  'fleep
' BallGuideOrbitLeft_Roll_Down()

  If SwitchRecentlyHit("swLeftOrbTrigger1") Then
    DOF 136,2 'left orbit
    startB2S 1
    BG_BarLeft=1

    if OrbUnlockShot = 1 And CurrentMission = 0 And Not bWizardmode And Not bMultiBallMode Then
      gate1status = Gate001.collidable
      Gate001.collidable=False    'let ball pass when orb unlock mission
      OrbUnlockShot = 10
    end if

    CheckBladerunnerShot 1
    CheckChaseShot 1
    CheckBetrayalMissionShot 1
    CheckTracyMBJackpot
    CheckWizardOrbShot
    AddScore score_LoopComplete

    If cLightCorey1.state = 2 and not bMultiballMode Then
      cLightCorey1.state = 1
      CheckCoreyMB
    End If
    If cLightTracy1.state = 2 and not bMultiballMode Then
      cLightTracy1.state = 1
      CheckTracyMB
    End If
    if bCoreyMBOngoing Then
      showMissionVideo = 40
    end if
  End If

  If SwitchRecentlyHit("swRightOrbTrigger2") Then
    if OrbUnlockShot = 50 And CurrentMission = 0 And Not bWizardmode And Not bMultiBallMode Then
      'debug.print "state returned"
      cLightShot5.state = 0
      cLightMissionStart1.state = 2
      Gate004.collidable=gate4status    'return previous status
      OrbUnlockShot = 0
    end if
  end if


  SwitchWasHit("swLeftOrbTrigger2")

  swLeftOrbTrigger2.TimerInterval = 150: swLeftOrbTrigger2.TimerEnabled = 1
End Sub

Sub swLeftOrbTrigger2_timer
  Dim i
  i=swLeftOrbTrigger2.uservalue+1
  If i > 9 then i = 9
  swLeftOrbTrigger2.uservalue=i
  Select Case i
    case 1,3,5,7 :
      Flash3 true : gioff
    case 2,4,6,8 :
      Flash1 true : gion
    Case 9 :
      gion
      swLeftOrbTrigger2.uservalue=0
      swLeftOrbTrigger2.timerenabled = false
  End Select
End Sub


Sub swRightOrbTrigger1_hit
    If Tilted Then Exit Sub
  WriteToLog "swRightOrbTrigger1_hit", ""
  if skillshot <> 0 then Skillshotoff : SKBackup = 0
  Addscore score_swRightOrbTrigger1_hit
' TracyAttention.Enabled=True
  If Tilted = false then AddScore score_LoopEntrance
  SwitchWasHit("swRightOrbTrigger1")
End Sub


'Shot5
Sub swRightOrbTrigger2_hit
    If Tilted Then Exit Sub
  WriteToLog "swRightOrbTrigger2_hit", ""

  'fleep sounds
' BallGuideOrbitRight_Roll_Up()
' BallGuideOrbitRight_Roll_Down()
' StopMetalRollWhenBallRollsSlowFromShooter = 1

  If SwitchRecentlyHit("swRightOrbTrigger1") Then
    DOF 137,2 'right orbit
    startB2S 1
    BG_BarRight=-1

    if OrbUnlockShot = 5 And CurrentMission = 0 And Not bWizardmode And Not bMultiBallMode Then
      gate4status = Gate004.collidable
      Gate004.collidable=False    'let ball pass when orb unlock mission
      OrbUnlockShot = 50
    end if

    CheckBladerunnerShot 5
    CheckChaseShot 5
    CheckBetrayalMissionShot 5
    CheckTracyMBJackpot
    CheckWizardGaldorShot
    CheckWizardOrbShot
    AddScore score_LoopComplete

    If cLightCorey5.state = 2 and not bMultiballMode Then
      cLightCorey5.state = 1
      CheckCoreyMB
    End If
    If cLightTracy5.state = 2 and not bMultiballMode Then
      cLightTracy5.state = 1
      CheckTracyMB
    End If
    if bCoreyMBOngoing Then
      showMissionVideo = 39
    end if
  End If
  If SwitchRecentlyHit("swLeftOrbTrigger2") Then
    if OrbUnlockShot = 10  And CurrentMission = 0 And Not bWizardmode And Not bMultiBallMode Then
      'debug.print "state returned"
      Gate001.collidable=gate1status    'return previous status
      cLightShot1.state = 0
      cLightMissionStart1.state = 2
      OrbUnlockShot = 0
    end if
  end if
  SKBackup = 0
  SwitchWasHit("swRightOrbTrigger2")
  swRightOrbTrigger2.TimerInterval = 150: swRightOrbTrigger2.TimerEnabled = 1
End Sub

Sub swRightOrbTrigger2_timer
  Dim i
  i=swRightOrbTrigger2.uservalue+1
  If i > 9 then i = 9
  swRightOrbTrigger2.uservalue=i
  Select Case i
    case 1,3,5,7 :
      Flash6 true
    case 2,4,6,8 :
      Flash7 true
    Case 9 :
      gion
      swRightOrbTrigger2.uservalue=0
      swRightOrbTrigger2.timerenabled = false
  End Select
End Sub



'Shot6
sub swSkyRamp_hit
    If Tilted Then Exit Sub
  WriteToLog "swSkyRamp_hit", ""
  DOF 139,2 'sky ramp
  PuPEvent(577)
  DOF 577, DOFPulse
  AddCoreyLight
  AddScore score_SkyLoop
  CheckCemeteryShot 6
  CheckCoreyMBSuperJackpot
  CheckGaldorFinalHit
  WireRampOff
  WireRampOn True 'Pastic Ramp
  PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_2"), RightPlasticRampHitsSoundLevel, swSkyRamp' : Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_2"
  SwitchWasHit("swSkyRamp")
end Sub

sub SkyloopTrigger_hit
  'just to start sound effect correct
  if activeball.velx > 10 Then
    PlaySound "EFX_Ship_sound04",0,CalloutVol,0,0,1,1,1
  end if
  'debug.print activeball.velx
end sub

dim bDoTopFlipTip : bDoTopFlipTip = True      'this is reset back to true after some time in attract (or new player in MP), so new player understands to use top flip
'Shot4
Sub swMainRamp_hit
    If Tilted Then Exit Sub
  WriteToLog "swMainRamp_hit", ""
  EOB_Ramps = EOB_Ramps +1

  if bDoTopFlipTip then FlickTopFlip

  if activeball.velx < -10 then activeball.velx = -10

  WireRampOff
  WireRampOn False 'Wire Ramp

  If not SwitchRecentlyHit("swSkyRamp") Then
    DOF 138,2 'main ramp
    PuPEvent(576)
    DOF 576, DOFPulse
    CheckBladerunnerShot 4
    CheckChaseShot 4
    CheckCemeteryShot 4
    CheckCoreyMBJackpot
    CheckWizardGaldorShot
    AddScore score_RampShot
    If cLightCorey4.state = 2 and not bMultiballMode Then
      cLightCorey4.state = 1
      CheckCoreyMB
    End If
    If cLightTracy4.state = 2 and not bMultiballMode Then
      cLightTracy4.state = 1
      CheckTracyMB
    End If
  End If
  SwitchWasHit("swMainRamp")
End Sub

sub FlickTopFlip
  flipper1.rotatetoend
  RandomSoundFlipperUpperLeftUpFullStroke flipper1
  vpmtimer.addtimer 80, "flipper1.rotatetostart:RandomSoundFlipperUpperLeftDown Flipper1 '"
  vpmtimer.addtimer 160, "flipper1.rotatetoend:RandomSoundFlipperUpperLeftUpFullStroke flipper1 '"
  vpmtimer.addtimer 240, "flipper1.rotatetostart:RandomSoundFlipperUpperLeftDown Flipper1 '"
  bDoTopFlipTip = False
end sub


sub tSpeedLimit_hit
  WriteToLog "tSpeedLimit_hit", ""
  WireRampOff
  WireRampOn True 'Pastic Ramp
  'debug.print "tSpeedLimit_hit"
  cLightPortalWall001.state = 1
  cLightPortalWall003.state = 1
  cLightPortalWall001.Timerenabled=False : cLightPortalWall001.Timerenabled=True
  cLightPortalWall003.Timerenabled=False : cLightPortalWall003.Timerenabled=True

  'debug.print "Speed limit: " & activeball.vely
  'debug.print "Speed limit x: " & activeball.velx
  if activeball.vely > 2.5 then activeball.vely = activeball.vely / 7 '14
  if activeball.velx > 6 then activeball.velx = activeball.vely / 2 '4
  'debug.print "y --> " & activeball.vely
  'debug.print "x --> " & activeball.velx
end sub


Sub swRampRoll1_Hit
  WriteToLog "swRampRoll1_Hit", ""
  WireRampOff
  PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_5"), RightPlasticRampHitsSoundLevel, swRampRoll1' : Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_5"
  if SwitchRecentlyHit("swRampRoll1") Then
    WireRampOn True 'Pastic Ramp
  Else
    WireRampOn False 'Wire Ramp
  End If
  SwitchWasHit("swRampRoll1")
End Sub


Sub swRampRoll2_Hit
  WriteToLog "swRampRoll2_Hit", ""
  WireRampOff
  PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"), RightPlasticRampHitsSoundLevel, swRampRoll2' : Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"
  WireRampOn False 'Wire Ramp
  SwitchWasHit("swRampRoll2")
End Sub


Sub swRampRoll3_Hit
  WriteToLog "swRampRoll3_Hit", ""
  cLightPortalWall001.state = 0
  cLightPortalWall002.state = 0
  cLightPortalWall003.state = 0
  cLightPortalWall001.Timerenabled=False
  cLightPortalWall002.Timerenabled=False
  cLightPortalWall003.Timerenabled=False
  WireRampOff
' RandomSoundMetal
  MetalGuideSideHit
  SwitchWasHit("swRampRoll3")
End Sub


Sub swRampRoll4_Hit
  WriteToLog "swRampRoll4_Hit", ""
  cLightPortalWall001.state = 0
  cLightPortalWall002.state = 0
  cLightPortalWall003.state = 0
  cLightPortalWall001.Timerenabled=False
  cLightPortalWall002.Timerenabled=False
  cLightPortalWall003.Timerenabled=False
  WireRampOff
  SwitchWasHit("swRampRoll4")
End Sub


Sub swRampRoll5_Hit
  WriteToLog "swRampRoll5_Hit", ""
  'LS_PortalWall.Play SeqRandom,1,,177
  cLightPortalWall003.state = 1
  cLightPortalWall002.state = 1
  If Tilted = false then AddScore score_RampEntrance
  cLightPortalWall003.Timerenabled=False : cLightPortalWall003.Timerenabled=True
  cLightPortalWall002.Timerenabled=False : cLightPortalWall002.Timerenabled=True
  WireRampOff
  WireRampOn True 'Plastic Ramp
  SwitchWasHit("swRampRoll5")
End Sub


Sub swRampRoll6_Hit
  WriteToLog "swRampRoll6_Hit", ""
  WireRampOff
  WireRampOn True 'Plastic Ramp
  SwitchWasHit("swRampRoll6")
End Sub


Sub swRampRoll7_Hit
  WriteToLog "swRampRoll7_Hit", ""
  cLightPortalWall003.state = 0
  cLightPortalWall002.state = 0
  cLightPortalWall003.timerenabled=False
  cLightPortalWall002.timerenabled=False
  cLightPortalWall001.timerenabled=False : cLightPortalWall001.timerenabled=True
  WireRampOff
  SwitchWasHit("swRampRoll7")
End Sub


Sub swRampRoll9_Hit
  WriteToLog "swRampRoll9_Hit", ""
  WireRampOff
  WireRampOn False 'Wire Ramp
  SwitchWasHit("swRampRoll9")
End Sub


Sub swRampRoll10_Hit
  WriteToLog "swRampRoll10_Hit", ""
  WireRampOff
' RandomSoundMetal
  MetalGuideSideHit
  SwitchWasHit("swRampRoll10")
End Sub



'timers to turn off portal lights after some period of time
Sub cLightPortalWall001_Timer
  cLightPortalWall001.state = 0
  cLightPortalWall001.timerenabled=False
End Sub

Sub cLightPortalWall002_Timer
  cLightPortalWall002.state = 0
  cLightPortalWall002.timerenabled=False
End Sub

Sub cLightPortalWall003_Timer
  cLightPortalWall003.state = 0
  cLightPortalWall003.timerenabled=False
End Sub



'**********************************************************************************************************
'  Multiballs
'**********************************************************************************************************

Dim MBInsertStates(10)

Sub SetMBReady
  WriteToLog "SetMBReady", " "
  bMultiballReady = True
  LS_Flash4.UpdateInterval = 300
  LS_Flash4.Play SeqBlinking ,,10000000,10
End Sub

Sub StartMBs
  WriteToLog "StartMBs", " "
  StartMimaMB
  if Not bMultiBallMode Then
    StartCoreyMB
    if Not bMultiBallMode Then
      StartTracyMB
    end if
  end if

  if bMultiBallMode Then
    'some multiball started
    bMultiballReady = False
    cLightShot1.state = 0
    cLightShot5.state = 0
    OrbUnlockShot = 0
    solBloodTrail True
    SaveMBInsertStates
    ClearMBInsertStates
    ScorbitBuildGameModes()
  End If
End Sub


Sub SaveMBInsertStates
  MBInsertStates(1)  = cLightCorey1.State
  MBInsertStates(2)  = cLightCorey2.State
  MBInsertStates(3)  = cLightCorey3.State
  MBInsertStates(4)  = cLightCorey4.State
  MBInsertStates(5)  = cLightCorey5.State

  MBInsertStates(6)  = cLightTracy1.State
  MBInsertStates(7)  = cLightTracy2.State
  MBInsertStates(8)  = cLightTracy3.State
  MBInsertStates(9)  = cLightTracy4.State
  MBInsertStates(10) = cLightTracy5.State

  MissionStart1LightState = cLightMissionStart1.state 'backup mission start light state for mb
  LightScavengeScoopState = cLightScavengeScoop.state
End Sub

Sub ClearMBInsertStates
  cLightCorey1.State = 0
  cLightCorey2.State = 0
  cLightCorey3.State = 0
  cLightCorey4.State = 0
  cLightCorey5.State = 0

  cLightTracy1.State = 0
  cLightTracy2.State = 0
  cLightTracy3.State = 0
  cLightTracy4.State = 0
  cLightTracy5.State = 0

  'Disable missions from starting
  cLightMissionStart1.State=0
  cLightScavengeScoop.state=0
End Sub

Sub RestoreMBInsertStates
  cLightCorey1.State = MBInsertStates(1)
  cLightCorey2.State = MBInsertStates(2)
  cLightCorey3.State = MBInsertStates(3)
  cLightCorey4.State = MBInsertStates(4)
  cLightCorey5.State = MBInsertStates(5)

  cLightTracy1.State = MBInsertStates(6)
  cLightTracy2.State = MBInsertStates(7)
  cLightTracy3.State = MBInsertStates(8)
  cLightTracy4.State = MBInsertStates(9)
  cLightTracy5.State = MBInsertStates(10)

  'restore mission start light state
  cLightMissionStart1.state = MissionStart1LightState
  cLightScavengeScoop.state = LightScavengeScoopState
End Sub

Sub StopMBs
  WriteToLog "StopMBs", " "
  if bMultiBallMode Then
    RestoreMBInsertStates
  End If
  bMultiballReady = False
  bMultiBallMode = False
  EndMimaMB
  EndCoreyMB
  EndTracyMB
  EndWizard
  solBloodTrail false

  If Not bMissionMode Then
    PlayDefaultSong
    if Not MissionStart1LightState = 2 then
      vpmtimer.addtimer MBGraceTime + 500, "OrbUnlockMission '"
      MissionStart1LightState = 0
    end if
    CheckMimaMB
    CheckCoreyMB
    CheckTracyMB
  End If

  checkWizardStart

  ScorbitBuildGameModes()

End Sub

dim bStartShitTalk : bStartShitTalk = false


sub checkWizardStart
  If cLightMimaMB.state = 1 and cLightCoreyMB.state = 1 and cLightTracyMB.state = 1 Then
    If cLightMission1.state=1 and cLightMission2.state=1 and cLightMission3.state=1 and cLightMission4.state=1 and cLightMission5.state=1 and cLightMission6.state=1 and cLightMission7.state=1 Then
      cLightMissionWizard.state = 2
      StartWizard
    Else
      bStartShitTalk = true
    end if
    else
      If cLightMission1.state=1 and cLightMission2.state=1 and cLightMission3.state=1 and cLightMission4.state=1 and cLightMission5.state=1 and cLightMission6.state=1 and cLightMission7.state=1 Then
      bStartShitTalk = true
    end if
  end if
end sub

Sub CheckMBReady
  WriteToLog "CheckMBReady", ""
  If NumLocksMima = 3 and not bWizardMode and cLightMimaMB.state<1 Then SetMBReady
  If (not bWizardMode) and cLightCoreyMB.state<1 and cLightCorey1.state = 1 and cLightCorey2.state = 1 and cLightCorey3.state = 1 and cLightCorey4.state = 1 and cLightCorey5.state = 1 Then SetMBReady
  If (not bWizardMode) and cLightTracyMB.state<1 and cLightTracy1.state = 1 and cLightTracy2.state = 1 and cLightTracy3.state = 1 and cLightTracy4.state = 1 and cLightTracy5.state = 1 Then SetMBReady
End Sub



'***********  Mima Multiball  ***************

Dim MimaLoopHits : MimaLoopHits = 0
Dim bMimaJPReady
Dim bMimaSuperJPReady
Dim bMimaMBOngoing
'Shot2 - Ramp Up
Sub MimaLock_Hit
' SoundSaucerLock
  RandomSoundScoopLeftEnter MimaLock
  PlaySound "EFX_sinkhole_shutdown",0,CalloutVol,0,0,1,1,1
  MimaLock.DestroyBall
  BallsOnPlayfield = BallsOnPlayfield - 1
  WriteToLog "MimaLock_Hit", "BallsOnPlayfield = " & BallsOnPlayfield
  CheckBladerunnerShot 2
  CheckChaseShot 2

  If cLightCorey2.state = 2 and not bMultiballMode Then
    cLightCorey2.state = 1
    CheckCoreyMB
  End If
  If cLightTracy2.state = 2 and not bMultiballMode Then
    cLightTracy2.state = 1
    CheckTracyMB
  End If


  if not bWizardMode then NumLocksMima = NumLocksMima + 1

  If NumLocksMima = 1 Then
    cMimalock1.state=1
    playsound "SPC_locked_ball_1",0,CalloutVol,0,0,1,1,1
    PuPEvent(547)
    DOF 547, DOFPulse
  End If

  If NumLocksMima = 2 Then
    cMimalock2.state=1
    playsound "SPC_locked_ball_2",0,CalloutVol,0,0,1,1,1
    PuPEvent(548)
    DOF 548, DOFPulse
  End If

  if NumLocksMima >= 3 then

    cMimalock1.state=1
    cMimalock2.state=1
    PuPEvent(549)
    DOF 549, DOFPulse

    'Mima MB Ready
    if cLightMimaMB.state=0 then AddMultiball 1
    CheckMimaMB
    bAutoPlunger = True
  else
    AddMultiball 1
    bAutoPlunger = True
    ResetMimaTargets
    ' locked sop sounds display  fixing
    if not tilted then AddScore score_MimaLock
  end If

  CheckMimaMBSuperJackpot
  Lower_RampMima.Enabled=True
End Sub

Sub CheckMimaMB
  WriteToLog "CheckMimaMB", ""
  If NumLocksMima >= 3 and not bWizardMode and not bMultiballMode and cLightMimaMB.state = 0 Then
    SetMBReady
    if cLightMimaMB.state=0 then AudioCallout "mima ready"
  End If
End Sub


Sub StartMimaMB
    If Tilted Then Exit Sub
  WriteToLog "StartMimaMB", ""
  If NumLocksMima = 3 and not bWizardMode and cLightMimaMB.state=0 Then
    WriteToLog "Mima MB", "START"
    bMultiBallMode = True
    PlaySong 1
    'PlaySongLoop 12, 240,284.62
    AddMultiball 2
    EnableBallSaver 15
    bAutoPlunger = True
    AddScore score_StartMimaMB
    PuPEvent(550)
    DOF 550, DOFOn
    MBPUP = "3"
    MimaMBloop
    bMimaMBOngoing = True
    cLightMimaMB.state = 2
    cLightMimaOrb1.state = 2
    cLightMimaOrb2.state = 2
    LS_FlashUnderShip.UpdateInterval = 700
    LS_FlashUnderShip.Play SeqBlinking ,,10000000,10
    MimaLoopHits = 0
    Jackpot = score_MimaMBJackpot
    SuperJackpot = score_MimaMBSuperJackpot
    bMimaJPReady = True
    bMimaSuperJPReady = False
    'If CurrentMission=5 Then StopMission : hatch false 'removed this


  End If
End Sub

Sub ResetMimaTargets
  MimaCount = 0
  cTargetMima1.state = 0
  cTargetMima2.state = 0
  cTargetMima3.state = 0
  Lower_RampMima.Enabled=True
  WriteToLog "ResetMimaTargets","MimaCount=" & MimaCount
End Sub


Sub CheckMimaMBJackpot
    If Tilted Then Exit Sub
  WriteToLog "CheckMimaMBJackpot", ""
  CheckResurrection

  If cLightMimaMB.state = 2 And bMimaJPReady Then
    MimaLoopHits = MimaLoopHits + 1
    if MimaLoopHits >= 7 Then
      MimaLoopHits = 0
      'Award Jackpot and setup for Super Jackpot
      WriteToLog "Mima MB", "JACKPOT!"
      DOF 140,2 'jackpot
      AudioCallout "jackpot"
      PuPEvent(557)
      DOF 557, DOFPulse
      LS_AllLightsAndFlashers.Play SeqCircleOutOn,10
      AddScore Jackpot

      DMD_ShowImages "jackpot",3,50,2000,0 ' jp fast blink


      If Not CurrentMission=2 Then ' dont turn off if mission 2 is active
        cLightMimaOrb1.state = 0
        cLightMimaOrb2.state = 0
      End If
      SetLightShot 2,1,"mimajp"
      Rise_RampMima.Enabled=True
      bMimaJPReady = False
      bMimaSuperJPReady = True
    End If
  End If
End Sub


Sub CheckMimaMBSuperJackpot
    If Tilted Then Exit Sub
  WriteToLog "CheckMimaMBSuperJackpot", ""
  If cLightMimaMB.state = 2 And bMimaSuperJPReady Then
    'Award Jackpot and setup for Jackpot
    WriteToLog "Mima MB", "SUPER JACKPOT!"
    DOF 141,2 'super jackpot
    showMissionVideo = 72 '-> 50
    AudioCallout "super jackpot"
    PuPEvent(561)
    DOF 561, DOFPulse
    LS_AllLightsAndFlashers.Play SeqDownOn, 10, 1
    AddScore SuperJackpot
    AddBonus bonus_MimaMBSuperJackpot
    SetLightShot 2,0,"mimajp"
    cLightMimaOrb1.state = 2
    cLightMimaOrb2.state = 2
    Jackpot = Jackpot + 0
    bMimaJPReady = True
    bMimaSuperJPReady = False
    AddMultiball 1
  End If
End Sub

Const MBGraceTime = 2000

Sub EndMimaMB
  MBPUP = "4"
  StopMimaMBPup
  WriteToLog "EndMimaMB", ""
  DOF 550, DOFOff
  If cLightMimaMB.state = 2 then
    tEndMimaMB.interval = MBGraceTime
    tEndMimaMB.enabled = True
  end If
End Sub

sub tEndMimaMB_timer
  WriteToLog "Mima MB", "END"
  cLightMimaMB.state = 1
  SetLightShot 2,0,"mimajp"
  cLightMimaOrb1.state = 0
  cLightMimaOrb2.state = 0
  LS_FlashUnderShip.StopPlay
  NumLocksMima = 0
  MimaLoopHits = 0
  Jackpot = 0
  SuperJackpot = 0
  bMimaJPReady = False
  bMimaSuperJPReady = False
  bMimaMBOngoing = False
  Lower_RampMima.Enabled=True
  checkWizardStart
  me.enabled = False
end sub

'********* Corey Multiball **************

Dim bCoreyJPReady
Dim bCoreySuperJPReady

Sub AddCoreyLight
  if not bMultiBallMode And cLightCoreyMB.state = 0 then
    WriteToLog "AddCoreyLight", ""
    startB2S 5
    dim xx,i
    i= int(rnd(1)*5)
    If i=0 And cLightCorey1.state=0 Then  cLightCorey1.state=2 : Exit Sub
    If i=1 And cLightCorey2.state=0 Then  cLightCorey2.state=2 : Exit Sub
    If i=2 And cLightCorey3.state=0 Then  cLightCorey3.state=2 : Exit Sub
    If i=3 And cLightCorey4.state=0 Then  cLightCorey4.state=2 : Exit Sub
    If i=4 And cLightCorey5.state=0 Then  cLightCorey5.state=2 : Exit Sub
    i= int(rnd(1)*5)
    If i=0 And cLightCorey1.state=0 Then  cLightCorey1.state=2 : Exit Sub
    If i=1 And cLightCorey2.state=0 Then  cLightCorey2.state=2 : Exit Sub
    If i=2 And cLightCorey3.state=0 Then  cLightCorey3.state=2 : Exit Sub
    If i=3 And cLightCorey4.state=0 Then  cLightCorey4.state=2 : Exit Sub
    If i=4 And cLightCorey5.state=0 Then  cLightCorey5.state=2 : Exit Sub

    for each xx in CoreyLights
      if xx.state = 0 then
        xx.state = 2
        playsound "EFX_spin01",0,CalloutVol,0,0,1,1,1
        exit Sub
      end if
    next
  end if
End Sub


Sub CheckCoreyMB
  WriteToLog "CheckCoreyMB", ""
  playsound "EFX_note01",0,CalloutVol,0,0,1,1,1
  If (not bWizardMode) and (not bMultiballMode) and cLightCoreyMB.state = 0 and cLightCorey1.state = 1 and cLightCorey2.state = 1 and cLightCorey3.state = 1 and cLightCorey4.state = 1 and cLightCorey5.state = 1 Then
    SetMBReady
    if cLightCoreyMB.state = 0 then AudioCallout "corey ready"
    PuPEvent(551)
    DOF 551, DOFPulse
    showMissionVideo = 31
  End If
End Sub

Sub StartCoreyMB
    If Tilted Then Exit Sub
  WriteToLog "StartCoreyMB", ""
  If (not bWizardMode) and cLightCoreyMB.state = 0 and cLightCorey1.state = 1 and cLightCorey2.state = 1 and cLightCorey3.state = 1 and cLightCorey4.state = 1 and cLightCorey5.state = 1 Then
    'Start Corey MB
    WriteToLog "Corey MB", "START"
    PuPEvent(552)
    DOF 552, DOFOn
    showMissionVideo = 45
    bMultiBallMode = True
    bCoreyMBOngoing = True
    MBPUP = "5"
    CoreyMBloop
    'PlaySong 1
    PlaySongLoop 2, 116.63, 170.6
'   If CurrentMission=5 Then StopMission 'mission 5 does not work with MBs
    cLightCoreyMB.state = 2
    SetLightShot 6,1,"coreyjp"
    SetLightShot 4,1,"coreyjp"
    AddMultiball 2
    EnableBallSaver 15
    bAutoPlunger = True
    Lower_RampMima.Enabled = True
    AddScore score_StartCoreyMB
    Jackpot = score_CoreyMBJackpot
    SuperJackpot = score_CoreyMBSuperJackpot
    bCoreyJPReady = True
    bCoreySuperJPReady = False : cLightShot6.state = 0
  End If
End Sub


Sub CheckCoreyMBJackpot
    If Tilted Then Exit Sub
  WriteToLog "CheckCoreyMBJackpot", ""
  If cLightCoreyMB.state = 2 And bCoreyJPReady Then
    'Award Jackpot and setup for Super Jackpot
    WriteToLog "Corey MB", "JACKPOT!"

    if Not bCoreySuperJPReady then
      vpmtimer.addtimer 4000, "bCoreySuperJPReady = True : cLightShot6.state = 2 : DMD_ShowText ""Super Jackpot Ready"",4,FontScoreInactiv2,29,True,40,3000 '"
    end if

    DOF 140,2 'jackpot
    AudioCallout "jackpot"
    DOF 140,2 'jackpot
    PuPEvent(558)
    DOF 558, DOFPulse
    LS_AllLightsAndFlashers.Play SeqCircleOutOn,10
    AddScore Jackpot
    DMD_ShowImages "jackpot",3,50,2000,0 ' jp fast blink
    LS_JP4.Play SeqBlinking,,5,20
  End If
End Sub


Sub CheckCoreyMBSuperJackpot
    If Tilted Then Exit Sub
  WriteToLog "CheckCoreyMBSuperJackpot", ""
  If cLightCoreyMB.state = 2 And bCoreySuperJPReady Then
    'Award Jackpot and setup for Jackpot
    WriteToLog "Corey MB", "SUPER JACKPOT!"
    bCoreySuperJPReady = False : cLightShot6.state = 0  'you need to get normal JP to enable super JP
    DOF 141,2 'super jackpot
    AudioCallout "super jackpot"
    PuPEvent(562)
    DOF 562, DOFPulse
    showMissionVideo = 41 '-> video nro 18
    LS_AllLightsAndFlashers.Play SeqDownOn, 10, 1
    AddScore SuperJackpot
    AddBonus bonus_CoreyMBSuperJackpot
    Jackpot = Jackpot + 0
    LS_JP6.Play SeqBlinking,,5,20
  End If
End Sub


Sub EndCoreyMB
  MBPUP = "6"
  StopCoreyMBPup
  WriteToLog "EndCoreyMB", ""
  DOF 552, DOFOff
  If cLightCoreyMB.state = 2 then
    tEndCoreyMB.interval = MBGraceTime
    tEndCoreyMB.enabled = True
  End If
End Sub

sub tEndCoreyMB_timer
  WriteToLog "Corey MB", "END"
  SetLightShot 4,0,"coreyjp"
  SetLightShot 6,0,"coreyjp"
  cLightShot6.state = 0
  cLightCoreyMB.state = 1
  If cLightCoreyMB.state = 1 and MimaCount = 3 and cLightMimaMB.state <> 1 Then
    Rise_RampMima.Enabled = True
  End If
  ResetCoreyLights      'Corey MB done, no need for Corey lights
  Jackpot = 0
  SuperJackpot = 0
  bCoreyJPReady = False
  bCoreySuperJPReady = False
  bCoreyMBOngoing = False
  showMissionVideo = 32
  checkWizardStart
  me.enabled = False
end sub


Sub ResetCoreyLights
  WriteToLog "ResetCoreyLights", ""
  cLightCorey1.state=0
  cLightCorey2.state=0
  cLightCorey3.state=0
  cLightCorey4.state=0
  cLightCorey5.state=0
End Sub



'********* Tracy Multiball **************

Dim TracyJPCount : TracyJPCount = 0
Dim bTracyJPReady
Dim bTracySuperJPReady

Sub AddTracyLight
  if not bMultiBallMode And Not cLightTracyMB.state = 1 And Not bWizardMode then
    WriteToLog "AddTracyLight", ""
    startB2S 6
    dim xx,i
    i= int(rnd(1)*5)
    If i=0 And cLightTracy1.state=0 Then  cLightTracy1.state=2 : Exit Sub
    If i=1 And cLightTracy2.state=0 Then  cLightTracy2.state=2 : Exit Sub
    If i=2 And cLightTracy3.state=0 Then  cLightTracy3.state=2 : Exit Sub
    If i=3 And cLightTracy4.state=0 Then  cLightTracy4.state=2 : Exit Sub
    If i=4 And cLightTracy5.state=0 Then  cLightTracy5.state=2 : Exit Sub
    i= int(rnd(1)*5)
    If i=0 And cLightTracy1.state=0 Then  cLightTracy1.state=2 : Exit Sub
    If i=1 And cLightTracy2.state=0 Then  cLightTracy2.state=2 : Exit Sub
    If i=2 And cLightTracy3.state=0 Then  cLightTracy3.state=2 : Exit Sub
    If i=3 And cLightTracy4.state=0 Then  cLightTracy4.state=2 : Exit Sub
    If i=4 And cLightTracy5.state=0 Then  cLightTracy5.state=2 : Exit Sub

    for each xx in TracyLights
      if xx.state = 0 then
        playsound "EFX_spin02",0,CalloutVol,0,0,1,1,1
        xx.state = 2
        exit Sub
      end if
    next
  end if
End Sub


Sub CheckTracyMB
  WriteToLog "CheckTracyMB", ""
  playsound "EFX_note01",0,CalloutVol,0,0,1,1,1
  If (not bWizardMode) and (not bMultiballMode) and cLightTracyMB.state = 0 and cLightTracy1.state = 1 and cLightTracy2.state = 1 and cLightTracy3.state = 1 and cLightTracy4.state = 1 and cLightTracy5.state = 1 Then
    SetMBReady
    showMissionVideo = 46
    PuPEvent(553)
    DOF 553, DOFPulse
    if cLightTracyMB.state = 0 then AudioCallout "tracy ready"
  End If
End Sub


Sub StartTracyMB
    If Tilted Then Exit Sub
  WriteToLog "StartTracyMB", ""
  If (not bWizardMode) and cLightTracyMB.state = 0 and cLightTracy1.state = 1 and cLightTracy2.state = 1 and cLightTracy3.state = 1 and cLightTracy4.state = 1 and cLightTracy5.state = 1 Then
    'Start Tracy MB
    WriteToLog "Tracy MB", "START"
    PuPEvent(554)
    DOF 554, DOFOn
    MBPUP = "1"
    TracyMBloop
    showMissionVideo = 47
    bMultiBallMode = True
    bTracyMBOngoing = True
    'PlaySong 10
    'PlaySongLoop 12, 178.357, 208.1
    PlaySongLoop 12, 178.357, 284
'   If CurrentMission=5 Then StopMission 'mission 5 does not work with MBs
    cLightTracyMB.state = 2
    SetLightShot 1,1,"tracyjp"
    SetLightShot 5,1,"tracyjp"
    AddMultiball 2
    EnableBallSaver 15
    bAutoPlunger = True
    Lower_RampMima.Enabled = True
    TracyJPCount = 0
    AddScore score_StartTracyMB
    Jackpot = score_TracyMBJackpot
    SuperJackpot = score_TracyMBSuperJackpot
    bTracyJPReady = True
    bTracySuperJPReady = False
  End If
End Sub


Sub CheckTracyMBJackpot
    If Tilted Then Exit Sub
  WriteToLog "CheckTracyMBJackpot", ""
  If cLightTracyMB.state = 2 And bTracyJPReady Then
    'Award Jackpot and setup for Super Jackpot
    WriteToLog "Tracy MB", "JACKPOT!"
    DOF 140,2 'jackpot
    AudioCallout "jackpot"
    PuPEvent(559)
    DOF 559, DOFPulse
    LS_AllLightsAndFlashers.Play SeqCircleOutOn,10
    AddScore Jackpot

    DMD_ShowImages "jackpot",3,50,2000,0 ' jp fast blink

    TracyJPCount = TracyJPCount + 1

    if TracyJPCount >= 2 Then
      TracyJPCount = 0
      SetLightShot 1,0,"tracyjp"
      SetLightShot 5,0,"tracyjp"
      SetLightShot 3,1,"tracyjp"

      LS_JP3.Stopplay : LS_JP3.Play SeqBlinking,,10,20

      bTracyJPReady = False
      bTracySuperJPReady = True
    End If
  End If
End Sub


Sub CheckTracyMBSuperJackpot
    If Tilted Then Exit Sub
  WriteToLog "CheckTracyMBSuperJackpot", ""
  If cLightTracyMB.state = 2 And bTracySuperJPReady Then
    'Award Jackpot and setup for Jackpot
    WriteToLog "Tracy MB", "SUPER JACKPOT!"
    DOF 141,2 'super jackpot
    AudioCallout "super jackpot"
    PuPEvent(563)
    DOF 563, DOFPulse
    showMissionVideo = 48 '-> 79
    LS_AllLightsAndFlashers.Play SeqDownOn, 10, 1
    AddScore SuperJackpot
    AddBonus bonus_TracyMBSuperJackpot
    SetLightShot 1,1,"tracyjp"
    SetLightShot 5,1,"tracyjp"
    SetLightShot 3,0,"tracyjp"
    LS_JP3.Stopplay : LS_JP3.Play SeqBlinking,,10,20

    Jackpot = Jackpot + 0
    bTracyJPReady = True
    bTracySuperJPReady = False
  End If
End Sub


Sub EndTracyMB
  MBPUP = "2"
  WriteToLog "EndTracyMB", ""
  DOF 554, DOFOff
  If cLightTracyMB.state = 2 then
    tEndTracyMB.interval = MBGraceTime
    tEndTracyMB.enabled = True
  End If
End Sub

sub tEndTracyMB_timer
  WriteToLog "Tracy MB", "END"
  SetLightShot 1,0,"tracyjp"
  SetLightShot 5,0,"tracyjp"
  SetLightShot 3,0,"tracyjp"

  cLightTracyMB.state = 1
  If cLightTracyMB.state = 1 and MimaCount = 3 and cLightMimaMB.state <> 1 Then
    Rise_RampMima.Enabled = True
  End If
  ResetTracyLights    'MB passed, reset Tracy Lights
  TracyJPCount = 0
  Jackpot = 0
  SuperJackpot = 0
  StopTracyMBPup
  bTracyJPReady = False
  bTracySuperJPReady = False
  bTracyMBOngoing = False
  checkWizardStart
  me.enabled = False
end sub


Sub ResetTracyLights
  WriteToLog "ResetTracyLights", ""
  cLightTracy1.state=0
  cLightTracy2.state=0
  cLightTracy3.state=0
  cLightTracy4.state=0
  cLightTracy5.state=0
End Sub


'***********  Final Chapter ************
dim WizardPhase:WizardPhase = 0

sub StartWizard
  WriteToLog "StartWizard", ""
  If cLightMissionWizard.state = 2 Then
    bWizardMode = True
    StartPhase1
  end if
end sub

dim WizardOrbShotCount : WizardOrbShotCount = 7

sub StartPhase1
  'initial phase
  'gioff
  WriteToLog "Wizard", "phase 1"

  WizardPhase = 1
  bWizardMode = true
  PuPevent(555)
  DOF 555, DOFPulse
  cLightMissionStart1.state = 0
  gioff
  bStartShitTalk = False
  if vrroom=1 then hexfloor.opacity = 1200 : hexfloor.visible = true
  cLightKB1.uservalue = cLightKB1.state:cLightKB2.uservalue = cLightKB2.state
  ScorbitBuildGameModes()

  vpmtimer.addtimer 1500, "GaldorShade false : GaldorQ1 = ""Ya"" '"

  PlayerLights_ResetInserts         'reset all lights off

  cLightKB1.state = cLightKB1.uservalue:cLightKB2.state = cLightKB2.uservalue

' 'skip past all main MB to get to Multiball Wizard
' cLightMimaMB.state = 1
' cLightCoreyMB.state = 1
' cLightTracyMB.state = 1
' 'skip past all missions to get to Mission Wizard
' cLightMission1.state = 1
' cLightMission2.state = 1
' cLightMission3.state = 1
' cLightMission4.state = 1
' cLightMission5.state = 1
' cLightMission6.state = 1
' cLightMission7.state = 1

  cLightMissionWizard.state = 2       'leave Wizard blinking
  'ship inserts blinking
  cLightMimaOrb1.state=2
  cLightMimaOrb2.state=2
  LS_FlashUnderShip.UpdateInterval = 700
  LS_FlashUnderShip.Play SeqBlinking ,,10000000,10
  'Show DMD that Wizard Mode Ready
  'check shot to swLoopUpTrigger        -> swloopuptrigger calls StartTracyChange if WizardPhase = 1
  'hold ball                  -> delay set at swloopuptrigger
  'tracy purple light and shake and sound   -> handled in StartTracyChange
  'once sequence done, dropwall       -> handled in swLoopUpTrigger timer
  'open gates on top to both directions   -> handled in StartTracyChange -> wizardPhase 2
  'shoot orb shots 7 times          -> WizardOrbShotCount = 7
end sub

dim TracyChangeHitCount : TracyChangeHitCount = 0
sub StartTracyChange
  WriteToLog "Wizard", "starttracychange"
  if cTracyChange.state = 0 then  'initial call to sub
    PuPEvent(556)
    DOF 556, DOFPulse
    WriteToLog "Wizard", "phase 1 tracy change"
    cTracyChange.state = 1
    ptracyCross.visible = true
    cTracyCross.BlinkInterval = 100
    cTracyCross.state = 2
    LS_FlashUnderShip.stopplay
    cLightMimaOrb1.state=0
    cLightMimaOrb2.state=0
    StopVideosPlaying
    showMissionVideo=62
    StopSong
'   GaldorReset=true:GaldorQ1 = "Yb"



    vpmtimer.addtimer 200, "PlaySound ""EFX_Tracy_dies"",0,CalloutVol,0,0,1,1,1 '"
    vpmtimer.addtimer 2200, "AudioCallout ""final chapter"" : ptracyCross.visible = true'"

  end if

  If ( TracyChangeHitCount Mod 2 ) = 1 Then
    ptracyCross.visible = true
  Else
    ptracyCross.visible = false
  end if

  if TracyChangeHitCount < 4 then     'limit the amount of shake calls
    shake_tracy_shake.enabled = true  'shakes tracy on every swloopuptrigger hit
    playsound "EFX_tracy_shake" & rndint(1,3),0,CalloutVol,0,0,1,1,1
  end if
  TracyChangeHitCount = TracyChangeHitCount + 1
  GaldorReset=true:GaldorQ1 = "hYb"
' GaldorReset=true:GaldorQ1 = "kYb"

  'add another timer that makes tracy do random stuff
end sub


sub StartPhase2
  WriteToLog "Wizard", "phase 2"
  cLightMissionStart1.state = 0
  'Orb shots
  WizardPhase = 2
  bWizardMode = true
  ScorbitBuildGameModes()
  'gion request from RIK
  GI_Color "red",0.5
  gion
  PlaySongLoop 10,121.54,131.65
  GaldorReset=true
  GaldorShade true
  PuPEvent(460)
  DOF 460, DOFPulse

  cTracyCross.BlinkInterval = 700
  cTracyCross.state = 2
  ptracyCross.visible = true
  Gate001.collidable=False
  Gate004.collidable=False
  LS_JP1all.StopPlay : LS_WizardOrbShotsLeft.Play SeqMiddleOutVertOn, 80, 100000
  LS_JP5all.StopPlay : LS_WizardOrbShotsRight.Play SeqMiddleOutVertOn, 80, 100000
  EnableBallSaver 10
end sub

Sub CheckWizardOrbShot
    If Tilted Then Exit Sub
  If WizardPhase = 2 Then
    WriteToLog "Wizard Orb shot", "Nice Shot!"
    DOF 142,2 'MB wizard orb shot
    AddScore score_WizardOrbShot

    'DMD_ShowImages "jackpot",3,50,2000,0 ' jp fast blink


    'AddBonus bonus_MBWizardJackpot

    shake_tracy_shake.enabled = true
    'show Lago_with hatchet and Vascan_Pulls off circuit board, Lago Hail Mary, Corey_tearing ship open

    WizardOrbShotCount = WizardOrbShotCount - 1
    LS_AllLightsAndFlashers.Play SeqUpOn, 10, 1
    AddScore score_WizardPhase2OrbShot


    if WizardOrbShotCount <=6 Then
      PuPEvent(461)
      DOF 461, DOFPulse
    end if
    If WizardOrbShotCount <=5 Then
      PuPEvent(462)
      DOF 462, DOFPulse
    end if
    If WizardOrbShotCount <=4 Then
      PuPEvent(463)
      DOF 463, DOFPulse
    end if
    If WizardOrbShotCount <=3 Then
      PuPEvent(464)
      DOF 464, DOFPulse
    end if
    If WizardOrbShotCount <=2 Then
      PuPEvent(465)
      DOF 465, DOFPulse
    end if
    If WizardOrbShotCount <=1 Then
      PuPEvent(466)
      DOF 466, DOFPulse
    end if
    If WizardOrbShotCount <=0 Then
      PuPEvent(450)
      DOF 450, DOFPulse
    end if

    if WizardOrbShotCount <= 0 Then
      endPhase2
      'show Tracy_arms out body video
      StopVideosPlaying
      showMissionVideo=63

      vpmtimer.addtimer 2500, "StartPhase3 '"
      'orb shots done, next to Phase3
    end if

  End If
End Sub


sub endPhase2
    If Tilted Then Exit Sub
  WriteToLog "Wizard", "end phase 2"
  AddScore score_WizardPhase2Passed
  'return default gates
  Gate001.collidable=True
  Gate004.collidable=False
' LS_JP1all.StopPlay
' LS_JP5all.StopPlay
  LS_WizardOrbShotsLeft.StopPlay
  LS_WizardOrbShotsRight.StopPlay
end Sub

dim bWizardModeUnderGame : bWizardModeUnderGame = false

sub StartPhase3
  WizardPhase = 3
  bWizardMode = true
  ScorbitBuildGameModes()
  cLightMissionStart1.state = 0
  'do NOT show supernova in DMD! It is already in playfield
  'Spaceship Warpspeed ? Spaceship in persuit ? Vascan_Looking in Orb Ground glows?
  StopVideosPlaying
  showMissionVideo=65

  SaveLaneInsertStates
  ClearLaneInsertStates

  '
  PlaySongLoop 10,148.88,186.4
  pUnderPF.image = "under_pf2"
  PlaySound "SPC_try_3_times",0,CalloutVol*2,0,0,1,1,1
  hatch true
  'PlaySong 12

  cLightUnderPFShot1.blinkinterval = 200
  cLightUnderPFShot2.blinkinterval = 200
  cLightUnderPFShot1.state = 0
  cLightUnderPFShot2.state = 0
  cLightMissionStart1.State=0
  PuPEvent(451)
  Dof 451, DOFPulse

  'LightPortalRoomShotsLeft 'moved to hatch routine

end sub

reflPortalRight.blenddisablelighting = 20


Sub CheckWizardUPFShot(nr)
    if tilted then exit sub
  If WizardPhase = 3 Then
    WriteToLog "Wizard", "Under PF Nice Shot!"
    If nr = 1 And cLightUnderPFShot1.state = 2 Then
      cLightUnderPFShot1.state = 1
      LightPortalRoomShotsRight
      PuPEvent(452)
      DOF 452, DOFPulse
    end if
    If nr = 2 And cLightUnderPFShot2.state = 2 Then
      cLightUnderPFShot2.state = 1
    end if

    If cLightUnderPFShot1.state = 1 and cLightUnderPFShot2.state = 1 Then
      bFlippersEnabled = False
      Flipper1.RotateToStart
      LeftFlipperMini.RotateToStart
      LeftFlipper.RotateToStart
      RightFlipperMini.RotateToStart
      RightFlipper.RotateToStart

      'darken underpf so that player notes, the minigame ended
      PlaySound "EFX_landing",0,CalloutVol,0,0,1,1,1
      pUnderPF.blenddisablelighting = 0

      'gion again on kick

      StopVideosPlaying
      showMissionVideo=66

      bWizardModeUnderGame = true   'So that VUK2 will start spaceplatform
      PuPEvent(453)
      DOF 453, DOFPulse

'     SetLightShot 1,1,"womanjp"
'     SetLightShot 2,1,"womanjp"
'     SetLightShot 3,1,"womanjp"
'     SetLightShot 4,1,"womanjp"
'     SetLightShot 5,1,"womanjp"
    End If
  End If
End Sub


Sub StartPhase4
    If Tilted Then Exit Sub
  'battle galdor <- called from underpf drain
'    - Once the minigame is over, that hatch closes, and Battle Galdor is started.
'    - GI state normal ON
'    - The Grapper maget is turned on, and 15 balls will be ejected from the back wall.
'        - Grapper will hold BOT-1 balls, so it is possible to have one ball in active play.
'        - So if the player shots his last ball to magnet, the magnet will release all the balls back in play.
'    - Galdor globe will show some impressions.
'    - There is no ball saver for this wizard mode
'    - Shots 3 thru 5 are lit for banishing Galdor, shots 1 and 2 skipped as grabber magnet is too close to them
'    - 5 shots needed to defeat Galdor. Final 5th shot should be the sky loop shot.
'        - Grand Wizard is passed succesfully if this skyloop shot succeeds.
'        - Flips, table and GI turns off and the last ball drains
'        - Final tune and light show plays off.
'        - After end credits, player gets his ball back and can continue the play.
'    - The Grand Wizard fails and ends when only one ball is left on the PF. The Mission lights get reset to off.
  AddScore score_WizardPhase3Passed

    WizardPhase = 4
  bWizardMode = true
  ScorbitBuildGameModes()
  gioff_underpf_boost false
  cLightMissionStart1.state = 0
  RestoreLaneInsertStates
  'Commander ship out hyperspace in DMD
  StopVideosPlaying
  showMissionVideo=67
  gion
  PlaySongLoop 12, 178.357, 284
  MimaMagnetOn True
  ClearMBInsertStates
  solMagnetGrab 255           'enable grabber
  SpaceWPlatform.timerenabled = true    'start multiball

  SetLightShot 3,1,"galdorshot"       'galdor shots
  SetLightShot 4,1,"galdorshot"
  SetLightShot 5,1,"galdorshot"
  GaldorQ1 = "Ye"
end sub

dim GaldorShotCount : GaldorShotCount = 5
dim finalGaldorHit : finalGaldorHit = False

Sub CheckWizardGaldorShot
    If Tilted Then Exit Sub
' WriteToLog "CheckMBWizardMBJackpot", ""
  If WizardPhase = 4 Then
    WriteToLog "CheckWizardGaldorShot", "Nice Shot!"
    DOF 142,2 'MB wizard jackpot



'   DMD_ShowImages "jackpot",3,50,2000,0 ' jp fast blink
'

    'AddBonus bonus_MBWizardJackpot
    LS_AllLightsAndFlashers.Play SeqUpOn, 10, 1

    if Not finalGaldorHit then
      GaldorShotCount = GaldorShotCount - 1
      GaldorReset=true:GaldorQ1 = "fYe"
    end If
    if GaldorShotCount = 4 Then
      PuPEvent(454)
      DOF 454, DOFPulse
    end if
    if GaldorShotCount = 3 Then
      PuPEvent(467)
      DOF 467, DOFPulse
    end if
    if GaldorShotCount = 2 Then
      PuPEvent(468)
      DOF 468, DOFPulse
    end if
    if GaldorShotCount = 1 Then
      PuPEvent(469)
      DOF 469, DOFPulse
      SetLightShot 3,0,"galdorshot"       'galdor shots
      SetLightShot 4,1,"galdorshot"       'leave this on to assist player to shoot towards skyloop
      SetLightShot 5,0,"galdorshot"
      'enable skyloop light for final shot
      'LS_JP6.StopPlay : LS_JP6.Play SeqBlinking,,5,20      'when hit
      finalGaldorHit = True
      LS_JP6.StopPlay : LS_JP6.Play SeqBlinking,,100000,50
    Else
      AddScore score_WizardGaldorShot * BallsOnPlayfield
    end if
  End If
End Sub

'**************  GaldorLockDiverter  **********
GaldorLock.isdropped = true

sub GaldorLockDiverter(Enabled)
  if Enabled Then
    GaldorLock.isdropped = false
    PlaySoundAt "Diverter_UP", Flasherlit1
  Else
    GaldorLock.isdropped = true
    PlaySoundAt "Diverter_DOWN", Flasherlit1
  end if
end sub

dim GaldorDefeated : GaldorDefeated = 0
sub CheckGaldorFinalHit
    If Tilted Then Exit Sub
  If WizardPhase = 4 and finalGaldorHit Then
    'msgbox "Galdor defeated!"
    GaldorReset=true:GaldorQ1 = "Yf"
    AddScore score_WizardCompleted * BallsOnPlayfield
    GaldorDefeated = 1
    LS_JP6.StopPlay : LS_JP6.Play SeqBlinking,,5,20     'when hit
    SetLightShot 4,0,"galdorshot"
    'enable hold behind backwall
    GaldorLockDiverter True


    StartPhase5
  end if
end sub

sub StartPhase5 'galdor defeated
  WizardPhase = 5
  bWizardMode = true
  ScorbitBuildGameModes()
  cLightMissionStart1.state = 0
  GaldorDefeated = 1
  PuPEvent(455)
  DOF 455, DOFPulse
  StopSong
  'release grabbed balls and let them drain
  solMagnetGrab 0 : MimaMagnetOn False
  bMultiBallMode = false
  'Play defeat yell
  AudioCallout "scream"
  DOF 143,2 'wizard completed
  'DMD yell
  StopVideosPlaying
  showMissionVideo = 71
  vpmtimer.addtimer 6000, "PlaySongPart 12, 302.44, 351.1 : bEndCredits = true : GaldorReset=true:GaldorQ1 = ""Yi"" '"
  'disable flips
  Tilted = true
  DisableTable True
  'GaldorLockRelease.enabled = true 'end wizard
  gioff
end sub

sub GaldorLockRelease
  'reset all the lights
  EndWizard
  'bWizardMode = False : WizardPhase = 0
  bEndCreditsCounter=0
  'GI_Color "white",1
  bEndCredits = false
  Tilted = False
  DisableTable false
  GaldorLockDiverter false
' GaldorLockRelease.enabled = false
  DMD_Stopoverlays = True

end sub




sub EndWizard
  WriteToLog "EndWizard", ""
  dim xx
  If bWizardMode Then
    WriteToLog "Wizard ", "END"
    if GaldorDefeated = 0 then showMissionVideo=64
    cLightUnderPFShot1.state = 0
    cLightUnderPFShot2.state = 0
    cLightMimaOrb1.state = 0
    cLightMimaOrb2.state = 0


'   to do later. This was more complex I originally though.
'   if wizardphase > 2 then   'wizard should continue if you are in 0,1,2 phase. so this is where it really ends
      bWizardMode = False
      WizardPhase = 0
      ScorbitBuildGameModes()
      bWizardModeUnderGame = false
      if GaldorDefeated = 0 then
        PlayDefaultSong
      end if
      WizardPhase3DrainCount = 3

      SetLightShot 1,0,"womanjp"
      SetLightShot 2,0,"womanjp"
      SetLightShot 3,0,"womanjp"
      SetLightShot 4,0,"womanjp"
      SetLightShot 5,0,"womanjp"
      MimaMagnetOn False
      ResetMimaTargets
      ResetCoreyLights
      ResetTracyLights
      WizardOrbShotCount = 7
      NumLocksMima = 0
      cMimalock1.state=0
      cMimalock2.state=0
      GI_Color "white",1
      if vrroom=1 then hexfloor.opacity = 0 : hexfloor.visible = false
      ShowDust False
      SetLightShot 2,0,"missionwizjp"
      RestoreMBInsertStates
      for each xx in MissionLights
        xx.blinkinterval=300
        xx.state = 0
      next
      cLightMissionStart1.state = 0
      GaldorDefeated = 0
      GaldorShotCount = 5
      finalGaldorHit = False
      TracyChangeHitCount = 0
      cTracyCross.state = 0
      ptracyCross.visible = false
      cTracyChange.state = 0
      SetLightShot 3,0,"galdorshot"
      SetLightShot 4,0,"galdorshot"
      SetLightShot 5,0,"galdorshot"
  '   LS_JP1all.StopPlay
  '   LS_JP5all.StopPlay
      LS_WizardOrbShotsLeft.StopPlay
      LS_WizardOrbShotsRight.StopPlay
      'restore MBs
      cLightMimaMB.state = 0
      cLightCoreyMB.state = 0
      cLightTracyMB.state = 0
      'restore Missions
      cLightMission1.state = 0
      cLightMission2.state = 0
      cLightMission3.state = 0
      cLightMission4.state = 0
      cLightMission5.state = 0
      cLightMission6.state = 0
      cLightMission7.state = 0

      cLightShot1.state = 0   'not mandatory, but helps testing
      cLightShot5.state = 0

      GaldorReset=true
      GaldorShade true
      swLeftOutLane.enabled=True    'make sure that left outlane is enabled after wizard
      OrbUnlockMission        'moved here to make sure all variables are reset
      hatch false
'   Else 'let wizard continue on next ball
'     'I guess we don't need anything here. Game should continue wizard normally
'   end if
  End If

end Sub


'dim tr,tg,tb,tmatcount,ttargetR,ttargetG,ttargetB
'tr = 121
'tg = 108
'tb = 57
'tmatcount = 0
'
'ttargetR = 224
'ttargetG = 20
'ttargetB = 168
'
'
'sub tTracyChange_timer
' tmatcount = tmatcount + 1
' 'MaterialColor "tracyskin", RGB(224,20,168)
''  (ttargetR/tr)*tmatcount
' UpdateMaterial pri.material,0,0,0,0,0,0,1,RGB(255*aLvl,80*aLvl,40*aLvl),0,0,False,True,0,0,0,0
'
'
'end sub




'*****todo delete these below*********

'***********  Multiball Wizard - 15 ball multiball ************
'
'sub StartMBWizardMB
' WriteToLog "StartMBWizardMB", ""
' If cLightMimaMB.state = 1 and cLightCoreyMB.state = 1 and cLightTracyMB.state = 1 Then
'   WriteToLog "Multiball Wizard MB", "START"
'   bMBWizardMode = True
'
'
'
'   pUnderPF.image = "under_pf2"
'   PlaySound "SPC_lets_see_if_you_can_pilot_me",0,CalloutVol*2,0,0,1,1,1
'   hatch true
'   PlaySong 12
'
'   cLightUnderPFShot1.state = 2
'   cLightUnderPFShot2.state = 2
'   cLightMissionStart1.State=0
' End If
'end Sub



'if bMBWizardModeUnderGame then
' solMagnetGrab 255           'enable grabber
' SpaceWPlatform.timerenabled = true    'start multiball
'end if


'Sub CheckMBWizardUPFShot(nr)
' WriteToLog "CheckMBWizardUPFShot", ""
' If bMBWizardMode Then
'   WriteToLog "Multiball Wizard MB", "Under PF Nice Shot!"
'   If nr = 1 Then cLightUnderPFShot1.state = 1
'   If nr = 2 Then cLightUnderPFShot2.state = 1
'
'   If cLightUnderPFShot1.state = 1 and cLightUnderPFShot2.state = 1 Then
'     bFlippersEnabled = False
'     Flipper1.RotateToStart
'     LeftFlipperMini.RotateToStart
'     LeftFlipper.RotateToStart
'     RightFlipperMini.RotateToStart
'     RightFlipper.RotateToStart
'
'     'darken underpf so that player notes, the minigame ended
'     PlaySound("EFX_landing")          'todo: maybe other sound
'     pUnderPF.blenddisablelighting = 0
'
'     'gion again on kick
'
'
'     bWizardModeUnderGame = true   'So that VUK2 will start spaceplatform
'
'     MimaMagnetOn True
'     ClearMBInsertStates
'     SetLightShot 1,1,"womanjp"
'     SetLightShot 2,1,"womanjp"
'     SetLightShot 3,1,"womanjp"
'     SetLightShot 4,1,"womanjp"
'     SetLightShot 5,1,"womanjp"
'   End If
' End If
'End Sub


'Sub CheckMBWizardMBJackpot
' WriteToLog "CheckMBWizardMBJackpot", ""
' If bMBWizardMode Then
'   WriteToLog "Multiball Wizard MB", "Nice Shot!"
'   DOF 142,2 'MB wizard jackpot
'   AddScore score_MBWizardJackpot
'
'   DMD_ShowImages "jackpot",3,50,2000,0 ' jp fast blink
'
'
'   AddBonus bonus_MBWizardJackpot
'   LS_AllLightsAndFlashers.Play SeqUpOn, 10, 1
' End If
'End Sub


sub SpaceWPlatform_timer
  dim BOTnow, BOTnowCount
  bMultiBallMode = True
  BOTnow = getballs
  BOTnowCount = Ubound(BOTnow)

  if BOTnowCount < 11 then
    swKicker001.createball
    swKicker001.kick 2,22,2
    BOTnowCount = BOTnowCount + 1
    BallsOnPlayfield = BallsOnPlayfield + 1
    WriteToLog "SpaceWPlatform_timer", "BallsOnPlayfield = "  & BallsOnPlayfield
  end if
  if BOTnowCount < 12 then
    swKicker002.createball
    swKicker002.kick 2,22,2
    BOTnowCount = BOTnowCount + 1
    BallsOnPlayfield = BallsOnPlayfield + 1
    WriteToLog "SpaceWPlatform_timer", "BallsOnPlayfield = "  & BallsOnPlayfield
  end if
  if BOTnowCount < 13 then
    swKicker003.createball
    swKicker003.kick 2,22,2
    BOTnowCount = BOTnowCount + 1
    BallsOnPlayfield = BallsOnPlayfield + 1
    WriteToLog "SpaceWPlatform_timer", "BallsOnPlayfield = "  & BallsOnPlayfield
  end if
  if BOTnowCount < 14 then
    swKicker004.createball
    swKicker004.kick 2,22,2
    BOTnowCount = BOTnowCount + 1
    BallsOnPlayfield = BallsOnPlayfield + 1
    WriteToLog "SpaceWPlatform_timer", "BallsOnPlayfield = "  & BallsOnPlayfield
  end if
  if BOTnowCount < 15 then
    swKicker005.createball
    swKicker005.kick 2,22,2
    BOTnowCount = BOTnowCount + 1
    BallsOnPlayfield = BallsOnPlayfield + 1
    WriteToLog "SpaceWPlatform_timer", "BallsOnPlayfield = "  & BallsOnPlayfield
  end if
  if BOTnowCount = 15 Then
    WriteToLog "SpaceWPlatform_timer", "15 on table -> exit"
    SpaceWPlatform.timerenabled = false
  end if
end sub




'***********  Mission Wizard - Progressive Multiball ************

'Dim MissionWizardCount
'
'sub StartMissionWizardMB
' WriteToLog "StartMissionWizardMB", ""
' dim xx
' If cLightMission1.state=1 and cLightMission2.state=1 and cLightMission3.state=1 and cLightMission4.state=1 and cLightMission5.state=1 and cLightMission6.state=1 and cLightMission7.state=1 Then
'   WriteToLog "Mission Wizard", "START"
'   bMultiBallMode = True
'   AudioCallout "final chapter"
'   PlaySong 12
'   Rise_RampMima.enabled = True
'   bMissionWizardMode = True
'   EnableBallSaver 10
'   AddMultiball 2
'   SaveMBInsertStates
'   ClearMBInsertStates
'   SetLightShot 2,1,"missionwizjp"
'   'LS_MissionLights.Play SeqBlinking,,1000000,20
'   for each xx in MissionLights
'     xx.blinkinterval=120
'     xx.state = 2
'   next
'   cLightMissionStart1.state = 0
'   MissionWizardCount = 0
' End If
'end Sub
'
'
'Sub CheckMissionWizardShot
' WriteToLog "CheckMissionWizardShot", ""
' If bMissionWizardMode Then
'   WriteToLog "Mission Wizard", "Nice Shot!"
'   MissionWizardCount = MissionWizardCount + 1
'   If MissionWizardCount <= 7 Then
'     MissionLights(MissionWizardCount-1).state = 1
'     EnableBallSaver 5
'     AddMultiball 1
'     LS_AllLightsAndFlashers.Play SeqDownOn, 10, 1
'     Select Case MissionWizardCount
'       Case 1 : GI_Color "green" ,1 :
'       Case 2 : GI_Color "red",1
'       Case 3 : GI_Color "blue",2
'       Case 4 : GI_Color "purple",1.2 : ShowDust True
'       Case 5 : GI_Color "red",1 : ShowDust False
'       Case 6 : GI_Color "purple",1.2
'       Case 7 : GI_Color "blue",2
'     End Select
'     'change GI lighting, do DMD stuff, etc
'   End If
'   AddScore score_MissionWizardJackpot
'   DMD_ShowImages "jackpot",3,50,2000,0 ' jp fast blink
'
'
'   AddBonus bonus_MissionWizardJackpot
'   If MissionWizardCount >= 8 Then
'     WriteToLog "Mission Wizard", "COMPLETED"
'     DOF 143,2 'MB wizard completed
'     AddScore score_MissionWizardCompleted
'     'do some crazy shit
'     EndMissionWizardMB
'   End If
' End If
'End Sub
'
'
'sub EndMissionWizardMB
' WriteToLog "EndMissionWizardMB", ""
' dim xx
' If bMissionWizardMode Then
'   WriteToLog "Mission Wizard", "END"
'   PlayDefaultSong
'   GI_Color "white",1
'   ShowDust False
'   SetLightShot 2,0,"missionwizjp"
'   RestoreMBInsertStates
'   for each xx in MissionLights
'     xx.blinkinterval=300
'     xx.state = 0
'   next
'   cLightMissionStart1.state = 2
'   bMissionWizardMode = False
'   If MimaCount>=3 and cLightMimaMB.state <> 1 Then Rise_RampMima.enabled=True Else Lower_RampMima.enabled=True
'   MissionWizardCount = 0
'   'LS_MissionLights.StopPlay
'   Next_MissionSelect
' End If
' hatch false
'end Sub
'




'**********************************************************************************************************
'  Missions
'**********************************************************************************************************

Dim MissionSelect
Dim MissionCompleted
Dim CurrentMission
Dim MissionTimeMax
Dim MissionTimeCurrent : MissionTimeCurrent=1000
Dim StartMissionName, StartMissionName2, StartMissionName3

Dim JPandMission(6,6)



Sub Mission_Begin
  StopVideosPlaying 'to eliminate some videos playing when mission starts
  WriteToLog "Mission_Begin", ""
  If MissionSelect=1 Then CurrentMission=1 : StartMissionName2 = "LOAD THE GUN"       : StartMissionName = "THE HUNT"   : StartMission1 : showMissionVideo=1
  If MissionSelect=2 Then CurrentMission=2 : StartMissionName2 = "MAKE 10 LOOPS"        : StartMissionName = "RESURRECTION" : StartMission2 : showMissionVideo=6
  If MissionSelect=3 Then CurrentMission=3 : StartMissionName2 = "GET 4 SUCSESSFUL SHOTS"   : StartMissionName = "CHASE"    : StartMission3 : showMissionVideo=10
  If MissionSelect=4 Then CurrentMission=4 : StartMissionName2 = "HIT 3 SHOTS TO COMPLETE"  : StartMissionName = "CEMETERY"   : StartMission4 : showMissionVideo=14
  If MissionSelect=5 Then CurrentMission=5 : StartMissionName2 = "HIT ALL VASCAN TARGETS"   : StartMissionName = "PORTAL ROOM"  : StartMission5 : showMissionVideo=18
  If MissionSelect=6 Then CurrentMission=6 : StartMissionName2 = "SAVE LAGO GET 25 BUMPERS" : StartMissionName = "HEARTOFSTEEL" : StartMission6 : showMissionVideo=23
  If MissionSelect=7 Then CurrentMission=7 : StartMissionName2 = "MAKE THE SHOT COMBO"    : StartMissionName = "BETRAYAL"   : StartMission7 : showMissionVideo=27
  If CurrentMission<>0 Then
    LS_MissionLights.Play SeqBlinking,,10,20
    ScorbitBuildGameModes()
  end if
End Sub




Sub Next_MissionSelect
  Dim xx
  'all missions played
  if cLightMission1.State = 1 and cLightMission2.State = 1 and cLightMission3.State = 1 and cLightMission4.State = 1 and cLightMission5.State = 1 and cLightMission6.State = 1 and cLightMission7.State = 1 Then
    exit Sub
  end if

  If CurrentMission=0 And Not bWizardMode Then
    If MissionSelect=1 and cLightMission1.state=2 Then cLightMission1.state=0
    If MissionSelect=2 and cLightMission2.state=2 Then cLightMission2.state=0
    If MissionSelect=3 and cLightMission3.state=2 Then cLightMission3.state=0
    If MissionSelect=4 and cLightMission4.state=2 Then cLightMission4.state=0
    If MissionSelect=5 and cLightMission5.state=2 Then cLightMission5.state=0
    If MissionSelect=6 and cLightMission6.state=2 Then cLightMission6.state=0
    If MissionSelect=7 and cLightMission7.state=2 Then cLightMission7.state=0

    For xx=1 to 7
      MissionSelect=MissionSelect+1
      If MissionSelect>7 Then MissionSelect=1

      If MissionSelect=1 And cLightMission1.state=0 Then cLightMission1.state=2 : Exit Sub
      If MissionSelect=2 And cLightMission2.state=0 Then cLightMission2.state=2 : Exit Sub
      If MissionSelect=3 And cLightMission3.state=0 Then cLightMission3.state=2 : Exit Sub
      If MissionSelect=4 And cLightMission4.state=0 Then cLightMission4.state=2 : Exit Sub
      If MissionSelect=5 And cLightMission5.state=0 Then cLightMission5.state=2 : Exit Sub
      If MissionSelect=6 And cLightMission6.state=0 Then cLightMission6.state=2 : Exit Sub
      If MissionSelect=7 And cLightMission7.state=0 Then cLightMission7.state=2 : Exit Sub
    Next
  End If
End Sub



Dim CounterBackground : CounterBackground=1
Dim CounterOldBackground : CounterOldBackground = 2
Dim CounterBG_Timer : CounterBG_Timer = 5

apronnumber1.blenddisablelighting = 2
apronnumber2.blenddisablelighting = 2

Sub MissionTimer_Timer
  'this will run all the time with 1 sec interval. We can add other checks here too

  If Not CurrentMission = 0 Then
    CounterBackground = 12
  Else
    CounterBG_Timer = CounterBG_Timer + 1
    If CounterBG_Timer > 6 Then
      CounterBG_Timer = 0
      CounterBackground = RndInt(11,1)
    End If
  End If

  MissionTimeCurrent = MissionTimeCurrent + 1

' FlexDMD2.LockRenderThread
  screen001.image = "counter" & CounterBackground

' FlexDMD2.Stage.GetImage("counterBG" & CounterOldBackground ).Visible = False
' FlexDMD2.Stage.GetImage("counterBG" & CounterBackground ).Visible = True
  CounterOldBackground = CounterBackground

  If MissionTimeCurrent < 1000 Then
    'WriteToLog "Mission Time", "t=" & MissionTimeCurrent
    If MissionTimeCurrent >= MissionTimeMax + 3 Then
      StopMission
'     FlexDMD2.Stage.GetLabel("apron1").Visible=False
      apronnumber1.visible=False
      apronnumber2.visible=False
    elseIf MissionTimeCurrent >= MissionTimeMax And MissionTimeCurrent < MissionTimeMax + 3 Then
      'grace period 3s
      'debug.print "mission grace"
      apronnumber1.blenddisablelighting = 10 : apronnumber2.blenddisablelighting = 10
      vpmtimer.addtimer 150, "apronnumber1.blenddisablelighting = 2 : apronnumber2.blenddisablelighting = 2 '"
      apronnumber1.image = 0
      apronnumber2.image = 0
      apronnumber1.visible=True
      apronnumber2.visible=True
    Else
      'debug.print "missiontime" & (MissiontimeMax-MissionTimeCurrent)
      dim x
      x = MissiontimeMax-MissionTimeCurrent
      If x > 99 then x = 99
      if x <= 10 And x > 0 then
        apronnumber1.blenddisablelighting = 10 : apronnumber2.blenddisablelighting = 10
        vpmtimer.addtimer 150, "apronnumber1.blenddisablelighting = 2 : apronnumber2.blenddisablelighting = 2 '"
        if x = 1 Then
          'end mission sounds etc
          StopCoutdownPup
          PlaySound "EFX_note01",0,CalloutVol,0,0,1,1,1
        elseif x = 10 Then
          PlaySound "SPC_10_left",0,CalloutVol,0,0,1,1,1
        elseif x = 5 Then
          PlaySound "SPC_5",0,CalloutVol,0,0,1,1,1
        else
          PlaySound "EFX_effect3",0,CalloutVol,0,0,1,1,1
        end if



      end if

      apronnumber1.image = Int( x / 10 ) : x = x - Int(x/10)*10
      apronnumber2.image = x' : debug.print x
      apronnumber1.visible=True
      apronnumber2.visible=True
    End If
  Else
      apronnumber1.visible=False
      apronnumber2.visible=False
  End If

  'Periodic CheckScript


  if GaldorDefeated = 0 then CheckBallSearch

  if PointMultiplier > 1 then TracyTrackingTime = TracyTrackingTime + 1

  if cOutholeKicker > 0 then RandomSoundOutholeKicker() : cOutholeKicker = 0

'Disabling this, as it can bug out and stop grabbing any balls
' if WizardPhase = 4 then GrabMagPeriodicRelease              'Release grabbed balls occasionally

  if bAttractMode then
    tAttractTime = tAttractTime + 1
    select case tAttractTime
      Case 30:  GI_Flicker_timer:bDoTopFlipTip = true
      Case 60:  PlaySound "SPC_set_me_free",0,CalloutVol,0,0,1,1,1
      Case 90:  PlaySongpart 6,56.2,96.2
      Case 120: GI_Flicker_timer
      Case 240: PlaySound "SPC_tracy_barely_felt",0,CalloutVol,0,0,1,1,1
      Case 400: PlaySongPart 12, 351.6, 402
      Case 480: GI_Flicker_timer
      Case 960: PlaySound "SPC_never_seen_wrecks",0,CalloutVol,0,0,1,1,1
      Case 1920:  GI_Flicker_timer
      Case 4000:  PlaySongpart 6,56.2,96.2
      Case 4096:  GI_Flicker_timer : tAttractTime = 121
    end select
  end if

  'return captive ball back if it has dropped out. This could be in some much slower timer.
  if IsObject(CaptiveBallID) then
    if Not IsEmpty(CaptiveBallID) then
      if Not InRect(CaptiveBallID.x,CaptiveBallID.y,  176,200,  317,200,  360,390,  190,400) then
        'debug.print "***** captive ball lost! Returning it back"
        CaptiveBallID.x = 282
        CaptiveBallID.y = 351
        CaptiveBallID.z = 41
      end if
    end if
  end if

  'all missions or MB's played
  if bStartShitTalk Then
    GaldorReset=true:GaldorQ1 = "YM":GaldorShade false
    CheckShitTalk
  Elseif cShitTalkCounter > 0 Then
    GaldorReset=true:GaldorShade True
    cShitTalkCounter = 0
  end if
End Sub

dim cShitTalkCounter : cShitTalkCounter = 0

sub CheckShitTalk
  if cShitTalkCounter > RndInt(10,50) Then
    select case RndInt(1,10)
      case 1,6,7: vpmtimer.addtimer 300, "PlaySound ""SPC_vascan_game_playing"",0,CalloutVol,0,0,1,1,1 '"
      case 2,9,10: vpmtimer.addtimer 300, "PlaySound ""SPC_vascan"",0,CalloutVol,0,0,1,1,1 '"
      case 3,8: vpmtimer.addtimer 300, "PlaySound ""SPC_signal_stronger"",0,CalloutVol,0,0,1,1,1 '"
      case 4: vpmtimer.addtimer 300, "PlaySound ""SPC_reinitiate_scan"",0,CalloutVol,0,0,1,1,1 '"
      case 5: vpmtimer.addtimer 300, "PlaySound ""SPC_my_god"",0,CalloutVol,0,0,1,1,1 '"
    end Select
    cShitTalkCounter = 0
  end if

  cShitTalkCounter = cShitTalkCounter + 1
end sub



dim CounterGrabMagPeriodicRelease : CounterGrabMagPeriodicRelease = 0
sub GrabMagPeriodicRelease
  if gballcount > 0 Then
'   debug.print "release"
    if CounterGrabMagPeriodicRelease = 20 Then
      solMagnetGrab 0
    elseif CounterGrabMagPeriodicRelease > 21 then
      solMagnetGrab 255
      CounterGrabMagPeriodicRelease = 0
    end if
    CounterGrabMagPeriodicRelease = CounterGrabMagPeriodicRelease + 1
  Else
'   debug.print "do nothing"
    CounterGrabMagPeriodicRelease = 0
  end if
end sub


dim bBallSearch : bBallSearch = false
sub CheckBallSearch
  if BallsOnPlayfield > 0 And Not bBallInPlungerLane then
    if CounterBallSearch > 15 And Scavengeball=0 And UnderKicker=0 Then
      WriteToLog "BallSearchStarted", ""
      CounterBallSearch = 2
      bBallSearch = true
      vpmtimer.addtimer 100, "VUK1.kickz 180,20,12,180 : RandomSoundScoopLeftEjectNormalVelocityDOF VUK1, 110, DOFPulse  '"
      vpmtimer.addtimer 2000, "VUK2.kickz 180,70,1,30 : RandomSoundScoopLeftEjectNormalVelocityDOF VUK2, 114, DOFPulse : VUK2.uservalue=0 : VUK2.timerinterval = 200 : VUK2.timerenabled = false '"
      vpmtimer.addtimer 1000, "ScavengeKicker.Kick 230,55,1.5 : RandomSoundScoopRightEjectHighVelocityDOF ScavengeKicker, 114, DOFPulse : ScavengeKicker.uservalue=0 : ScavengeKicker.timerinterval = 150 : ScavengeKicker.timerenabled = false '"
      vpmtimer.addtimer 500, "TracyAttention.Enabled = True '"
      if currentMission = 4  then
        vpmtimer.addtimer 200, "MimaMagnetOn False '"
        vpmtimer.addtimer 400, "MimaMagnetOn true '"
      Elseif MagnetMode Then
        vpmtimer.addtimer 200, "MimaMagnetOn False '"
      end if


      if WizardPhase = 4  then
        vpmtimer.addtimer 200, "solMagnetGrab 0 '"
        vpmtimer.addtimer 400, "solMagnetGrab 255 '"
      end if

      vpmtimer.addtimer 1700, "DropLoopWall True '"
'     vpmtimer.addtimer 3200, "Lower_RampMima.enabled=True '"
'     vpmtimer.addtimer 4300, "Rise_RampMima.enabled=True '"
    end if

    'Narnia ball check every 5 secs
    if CounterBallSearch = 2 Or CounterBallSearch = 7 then
'     debug.print "Narnia balls"
      Dim b, nBOT
      nBOT = GetBalls
      For b = 0 to UBound(nBOT)
        if nBOT(b).z < -200 Then
          'msgbox "Ball " &b& " in Narnia X: " & gBOT(b).x &" Y: "&gBOT(b).y & " Z: "&gBOT(b).z
          debug.print "Move narnia ball (X: "& nBOT(b).x &" Y: "&nBOT(b).y & " Z: "&nBOT(b).z&") to right scoop"
          nBOT(b).x = 786
          nBOT(b).y = 1278
          nBOT(b).z = -40
        end if
      next
    end if

    if CounterBallSearch = 0 And bBallSearch Then
'     debug.print "end ball search as some switch was hit by the ball"
      bBallSearch = false
      exit sub
    end if


    CounterBallSearch = CounterBallSearch + 1
  Else
    CounterBallSearch = 0
  end if
end sub

Sub StopMission
  WriteToLog "StopMission", ""
  GI_Color "white",1

  If CurrentMission=1 Then EndMission1
  If CurrentMission=2 Then EndMission2
  If CurrentMission=3 Then EndMission3
  If CurrentMission=4 Then EndMission4
  If CurrentMission=5 Then EndMission5
  If CurrentMission=6 Then EndMission6
  If CurrentMission=7 Then EndMission7

  CurrentMission = 0

  MissionTimeCurrent=1000

  If MissionCompleted Then
    LS_AllFlashers.Play SeqDownOn, 12, 1
    ShowMissionComplete=50
    ShowMissionFailed=0
    AudioCallout "mission completed"
  Else
    ShowMissionComplete=0
    ShowMissionFailed=50
    AudioCallout "mission failed"
  End If

  'If MissionCompleted Then MissionSelect=8 : Next_MissionSelect  'not needed anymore, as now mission is swapped after failure
  MissionCompleted = False

  'additional cleanups
    'GaldorShade true
    solMagnetGrab 0
    MimaMagnetOn false


  LS_MissionLights.Play SeqBlinking,,10,20

  If cLightMission1.state=1 and cLightMission2.state=1 and cLightMission3.state=1 and cLightMission4.state=1 and cLightMission5.state=1 and cLightMission6.state=1 and cLightMission7.state=1 Then
    If cLightMimaMB.state = 1 and cLightCoreyMB.state = 1 and cLightTracyMB.state = 1 Then
      cLightMissionWizard.state = 2
      StartWizard
    Else
      bStartShitTalk = true
    end if
  Else
  ' If Not bMultiBallMode Then    'we probably can call this eventhough mb can be ongoing at the same time
      cLightMissionStart1.state = 0
      OrbUnlockMission
  ' End If
    Next_MissionSelect        'force mission swap when previous ends
  End If
  ScorbitBuildGameModes()
End Sub


'Award EB light if 3 or 6 missions completed
Sub CheckMissionCompleteAward
  WriteToLog "CheckMissionCompleteAward", ""
  dim xx, numcomplete
  numcomplete = 0
  For each xx in MissionLights
    If xx.state=1 Then numcomplete = numcomplete + 1
  Next
  If (numcomplete = 3 or numcomplete = 6) and TotalExtraBallsAwarded(CurrentPlayer) < MaxExtraBallsPerGame Then
    cLightCollectExtraBall.state=2
    AudioCallout "extra ball lit"
  End If
End Sub




Sub SetLightShot(nr,state,jp_mission)
  WriteToLog "SetLightShot", "nr=" & nr & " state=" & state & " " & jp_mission

  Dim JPShotOn, MissionShotOn
  JPShotOn = False
  MissionShotOn = False


  If jp_mission = "mimajp" Then
    if state > 0 Then JPandMission(nr-1,0) = 1
    if state = 0 Then JPandMission(nr-1,0) = 0
  Elseif jp_mission = "coreyjp" Then
    if state > 0 Then JPandMission(nr-1,1) = 1
    if state = 0 Then JPandMission(nr-1,1) = 0
  Elseif jp_mission = "tracyjp" Then
    if state > 0 Then JPandMission(nr-1,2) = 1
    if state = 0 Then JPandMission(nr-1,2) = 0
  Elseif jp_mission = "galdorshot" Then
    if state > 0 Then JPandMission(nr-1,3) = 1
    if state = 0 Then JPandMission(nr-1,3) = 0
  Elseif jp_mission = "mission" Then
    if state > 0 Then JPandMission(nr-1,4) = 1
    if state = 0 Then JPandMission(nr-1,4) = 0
  Elseif jp_mission = "missionwizjp" Then
    if state > 0 Then JPandMission(nr-1,5) = 1
    if state = 0 Then JPandMission(nr-1,5) = 0
  Elseif jp_mission = "guntarget" Then
    if state > 0 Then JPandMission(nr-1,6) = 1
    if state = 0 Then JPandMission(nr-1,6) = 0
  End If

  If JPandMission(nr-1,0)=1 OR JPandMission(nr-1,1)=1 OR JPandMission(nr-1,2)=1 OR JPandMission(nr-1,3)=1 OR JPandMission(nr-1,5)=1 OR JPandMission(nr-1,6)=1 Then JPShotOn=True
  If JPandMission(nr-1,4)=1 Then MissionShotOn=True

  If JPShotOn OR MissionShotOn Then ShotLights(nr-1).state=2 Else ShotLights(nr-1).state=0 End If

  If JPShotOn Then
    If nr=1 Then LS_JP1all.StopPlay : LS_JP1all.Play SeqMiddleOutVertOn, 80, 100000
    If nr=2 Then LS_JP2all.StopPlay : LS_JP2all.Play SeqMiddleOutVertOn, 80, 100000
    If nr=3 Then LS_JP3all.StopPlay : LS_JP3all.Play SeqMiddleOutVertOn, 80, 100000
    If nr=4 Then LS_JP4all.StopPlay : LS_JP4all.Play SeqMiddleOutVertOn, 80, 100000
    If nr=5 Then LS_JP5all.StopPlay : LS_JP5all.Play SeqMiddleOutVertOn, 80, 100000
    If nr=6 Then LS_JP6.StopPlay : LS_JP6.Play SeqBlinking,,100000,50
  Else
    If nr=1 Then LS_JP1all.StopPlay
    If nr=2 Then LS_JP2all.StopPlay
    If nr=3 Then LS_JP3all.StopPlay
    If nr=4 Then LS_JP4all.StopPlay
    If nr=5 Then LS_JP5all.StopPlay
    If nr=6 Then LS_JP6.StopPlay
  End If

  If MissionShotOn Then
    If nr=1 Then LS_JP1.Play SeqBlinking,,100000,50
    If nr=2 Then LS_JP2.Play SeqBlinking,,100000,50
    If nr=3 Then LS_JP3.Play SeqBlinking,,100000,50
    If nr=4 Then LS_JP4.Play SeqBlinking,,100000,50
    If nr=5 Then LS_JP5.Play SeqBlinking,,100000,50
    If nr=6 Then LS_JP6.Play SeqBlinking,,100000,50
  Else
    If nr=1 Then LS_JP1.StopPlay
    If nr=2 Then LS_JP2.StopPlay
    If nr=3 Then LS_JP3.StopPlay
    If nr=4 Then LS_JP4.StopPlay
    If nr=5 Then LS_JP5.StopPlay
    If nr=6 Then LS_JP6.StopPlay
  End If

End Sub


Function GetLightShot(nr,jp_mission)
  If jp_mission = "mimajp" Then
    GetLightShot = JPandMission(nr-1,0)
  Elseif jp_mission = "coreyjp" Then
    GetLightShot = JPandMission(nr-1,1)
  Elseif jp_mission = "tracyjp" Then
    GetLightShot = JPandMission(nr-1,2)
  Elseif jp_mission = "galdorshot" Then
    GetLightShot = JPandMission(nr-1,3)
  Elseif jp_mission = "mission" Then
    GetLightShot = JPandMission(nr-1,4)
  Elseif jp_mission = "missionwizjp" Then
    GetLightShot = JPandMission(nr-1,5)
  Elseif jp_mission = "guntarget" Then
    GetLightShot = JPandMission(nr-1,6)
  Else
    GetLightShot = -1 'error
  End If
End Function


Sub AddMissionScore(MissionNumber, Points)
    If Tilted Then Exit Sub
  Select Case MissionNumber
    Case 1: Mission1Points(CurrentPlayer) = Mission1Points(CurrentPlayer) + Points
    Case 2: Mission2Points(CurrentPlayer) = Mission2Points(CurrentPlayer) + Points
    Case 3: Mission3Points(CurrentPlayer) = Mission3Points(CurrentPlayer) + Points
    Case 4: Mission4Points(CurrentPlayer) = Mission4Points(CurrentPlayer) + Points
    Case 5: Mission5Points(CurrentPlayer) = Mission5Points(CurrentPlayer) + Points
    Case 6: Mission6Points(CurrentPlayer) = Mission6Points(CurrentPlayer) + Points
    Case 7: Mission7Points(CurrentPlayer) = Mission7Points(CurrentPlayer) + Points
  End Select
  AddScore Points
End Sub




'************* The Hunt - Mission 1 ***************

Dim TotalBladerunnerHits
Dim Mission1ShotEnabled
Dim Mission1TargetPupCount


Sub StartMission1
  WriteToLog "Mission 1", "START"

  GI_Color "green" ,1

  Mission1Points(CurrentPlayer) = 0

  cLightMission1.blinkinterval=60
  bMissionMode = True
  MissionCompleted = False
  MissionTimeCurrent = 0
  MissionTimeMax = 100 '75
  Mission1ShotEnabled = False
  'AddMultiball 1         '2 ball mode removed
  Mission1TargetPupCount=3
  TotalBladerunnerHits = 0
  SetLazerReady
  vpmtimer.addtimer 1600, "PlaySong 2 '"
  AudioCallout "start mission 1"
  PuPEvent(513)
  PuPEvent(524)
  DOF 513, DOFOn
  ModeTextDelay1.Enabled=True

End Sub



Sub LightBladerunnerShot
  WriteToLog "LightBladerunnerShot", ""
  'Pick an available shot at semi-random
  dim xx,i
  i= int(rnd(1)*5)+1
  If GetLightShot(i,"mission")=0 Then SetLightShot i,1,"mission"  : Exit Sub
  i= int(rnd(1)*5)+1
  If GetLightShot(i,"mission")=0 Then SetLightShot i,1,"mission" : Exit Sub
  For xx = 1 to 5
    If GetLightShot(xx,"mission")=0 Then SetLightShot xx,1,"mission" : Exit Sub
  Next
End Sub

Sub LightGunTarget
  WriteToLog "LightGunTarget", ""
  'Pick an available shot at semi-random
  dim xx,i
  i= int(rnd(1)*5)+1
  If GetLightShot(i,"guntarget")=0 Then SetLightShot i,1,"guntarget"  : Exit Sub
  i= int(rnd(1)*5)+1
  If GetLightShot(i,"guntarget")=0 Then SetLightShot i,1,"guntarget" : Exit Sub
  For xx = 1 to 5
    If GetLightShot(xx,"guntarget")=0 Then SetLightShot xx,1,"guntarget" : Exit Sub
  Next
End Sub

Sub BladerunnerShotTimer_Timer
  dim xx
  'Turn off lit shot if it was not hit
  For xx=1 to 5
    if Mission1ShotEnabled Then
      SetLightShot xx,0,"mission"
    Else
      SetLightShot xx,0,"guntarget"
    end if
  Next
  if OrbUnlockShot <> 0 then OrbUnlockMission       'reset unlock shot
  Mission1ShotEnabled = False
  BladerunnerShotTimer.Enabled = False
End Sub

dim GunTargetScorePerPlayer : GunTargetScorePerPlayer = score_GunTarget
Sub CheckBladerunnerShot(nr)
    If Tilted Then Exit Sub
  If CurrentMission = 1 Then
    If GetLightShot(nr,"mission") = 1 and Mission1ShotEnabled Then
      WriteToLog "Mission 1", "Nice Shot!"
      AudioCallout "nice shot"
      Playsound "EFX_hit_to_target_with_gun",0,CalloutVol,0,0,1,1,1
      LS_ShotLights.Play SeqBlinking,,3,20
      SetLightShot nr,0,"mission"

      TotalBladerunnerHits = TotalBladerunnerHits + 1
      Mission1TargetPupCount = Mission1TargetPupCount - 1

      WriteToLog "Mission 1", "Mission1 Total Hits = " & TotalBladerunnerHits

      fontblink2=55
      If TotalBladerunnerHits = 1 Then
        'MissionCompleted = True
        AddMissionScore 1,score_Mission1_ShotHit
        ShowAddScoreText = FormatScore(score_Mission1_ShotHit*PointMultiplier)
        ShowAddScore = 1
        creditblink=30
        ScoreMissionPoints = "1"
        MissionPointsPup.Enabled=True
      End If
      If TotalBladerunnerHits = 2 Then
        AddMissionScore 1,score_Mission1_ShotHit
        ShowAddScoreText = FormatScore(score_Mission1_ShotHit*PointMultiplier)
        ShowAddScore = 1
        creditblink=30
        ScoreMissionPoints = "1"
        MissionPointsPup.Enabled=True
      End If

      If TotalBladerunnerHits >= 3 Then
        AddMissionScore 1,score_Mission1_ShotHit
        ShowAddScoreText = FormatScore(score_Mission1_AllShots*PointMultiplier)
        ShowAddScore = 1
        creditblink=30
        ScoreMissionPointsC = "1C"
        MissionPointsPupC.Enabled=True
        MissionCompleted = True
        AddBonus bonus_Mission1_Completed
        StopMission
      End If
    End If
  Else
    If GetLightShot(nr,"guntarget") = 1 Then
      WriteToLog "Gun Target", "Nice Shot!"
      AudioCallout "nice shot"
      GunshotPUP=1
      Playsound "EFX_hit_to_target_with_gun",0,CalloutVol,0,0,1,1,1
      LS_ShotLights.Play SeqBlinking,,3,20
      SetLightShot nr,0,"guntarget"
      if OrbUnlockShot <> 0 then OrbUnlockMission
      LS_AllLightsAndFlashers.Play SeqDownOn, 10, 1
      'todo add score + dmd thing
      fontblink2=55
      PuPEvent(877)
      GS.Enabled=True
      AddScore GunTargetScorePerPlayer
      ShowAddScoreText = FormatScore(GunTargetScorePerPlayer)
      GunTargetScorePerPlayer = GunTargetScorePerPlayer + score_GunTarget_increment
      GunTargetScore(CurrentPlayer) = GunTargetScorePerPlayer
      showMissionVideo=61
    end if
  End If
End Sub
Dim ShowAddScore : ShowAddScore=0
Dim ShowAddScoreText : ShowAddScoreText=" "



Sub EndMission1
  PlayDefaultSong
  cLightMission1.blinkinterval=300
  bMissionMode = False
  CloseGun
  SetLightShot 1,0,"mission"
  SetLightShot 2,0,"mission"
  SetLightShot 3,0,"mission"
  SetLightShot 4,0,"mission"
  SetLightShot 5,0,"mission"
  LS_Flash5.StopPlay
  cLightLoadLazer.state = 0

  If ShowAddScore=0 Then : ShowAddScoreText = "COMPLETED"

  If MissionCompleted Then
    EOB_Missions = EOB_Missions + 1

    cLightMission1.state=1
    CheckMissionCompleteAward
    WriteToLog "Mission 1", "Completed"
    PuPEvent(526)
    DOF 513, DOFOff
    PupMission= "Default"
    Mtimercountdown= "Default"
    StopVideosPlaying
    showMissionVideo=5
    MBPupCheck.Enabled=true
  Else
    cLightMission1.state=2
    WriteToLog "Mission 1", "Not Completed"
    PuPEvent(525)
    DOF 513, DOFOff
    PupMission= "Default"
    Mtimercountdown= "Default"
    StopVideosPlaying
    showMissionVideo=3
    MBPupCheck.Enabled=true
  End If
End Sub




'************* Resurrection - Mission 2 ***************


Dim TotalMimaLoops
Dim Mission1TargetPupCount2
Dim LoopPupScore: LoopPupScore=0

Sub StartMission2
  WriteToLog "Mission 2", "START"
  GI_Color "red",1

  Mission2Points(CurrentPlayer) = 0

  cLightMission2.blinkinterval=60
  bMissionMode = True
  MissionCompleted = False
  MissionTimeCurrent = 0
  MissionTimeMax = 75   '100
  Mission1TargetPupCount2=25
  TotalMimaLoops=0
  LoopPupScore= 0
  LS_FlashUnderShip.UpdateInterval = 700
  LS_FlashUnderShip.Play SeqBlinking ,,10000000,10

  cLightMimaOrb1.state=2
  cLightMimaOrb2.state=2

  'PlaySong 6
  PlaySongLoop 6,24.5,88.3
  AudioCallout "start mission 2"
  PuPEvent(512)
  PuPEvent(523)
  DOF 512, DOFOn
  ModeTextDelay2.Enabled=True

End Sub


Sub CheckResurrection
    If Tilted Then Exit Sub
  If CurrentMission = 2 Then
    TotalMimaLoops = TotalMimaLoops + 1
    Mission1TargetPupCount2 = Mission1TargetPupCount2 -1
    LoopPupScore =  100000
    WriteToLog "Mission 2", "Mima Loops = " & TotalMimaLoops
    If TotalMimaLoops = 25 Then
      MissionCompleted=True
      AddMissionScore 2,score_Mission2_10Loops

    Elseif TotalMimaLoops=1 then
      showMissionVideo=7
    End If


    If TotalMimaLoops >= 25 Then
      AddBonus bonus_Mission2_Completed
      ShowAddScoreText = FormatScore(score_Mission2_10Loops*PointMultiplier)
      ScoreMissionPointsC = "2C"
      MissionPointsPupC.Enabled=True
      ShowAddScore = 1
      creditblink=130
      StopMission ' completed
    Else
      AddMissionScore 2,score_Mission2_EachLoop
      ShowAddScoreText = FormatScore(score_Mission2_EachLoop*PointMultiplier)
      ScoreMissionPoints = "2"
      MissionPointsPup.Interval=700
      MissionPointsPup.Enabled=True
      ShowAddScore = 1
      creditblink=20
    End If
  End If
End Sub


Sub EndMission2
  PlayDefaultSong
  cLightMission2.blinkinterval=300
  LS_FlashUnderShip.StopPlay
  bMissionMode = False
  If Not cLightMimaMB.state=2 Then
    cLightMimaOrb1.state=0
    cLightMimaOrb2.state=0
  End If

  If ShowAddScore=0 Then: ShowAddScoreText = "COMPLETED"

  If MissionCompleted Then
    EOB_Missions = EOB_Missions + 1

    cLightMission2.state=1
    CheckMissionCompleteAward
    WriteToLog "Mission 2", "Completed"
    AudioCallout "end mission 2"
    StopVideosPlaying
    showMissionVideo=8
    PuPEvent(526)
    DOF 512, DOFOff
    PupMission= "Default"
    Mtimercountdown= "Default"
  Else
    cLightMission2.state=2
    WriteToLog "Mission 2", "Not Completed"
    StopVideosPlaying
    showMissionVideo=9
    PuPEvent(525)
    DOF 512, DOFOff
    PupMission= "Default"
    Mtimercountdown= "Default"
  End If
End Sub





'************* Chase - Mission 3 ***************


Dim TotalChaseHits
Dim Mission1TargetPupCount3
Dim ChasePupScore: ChasePupScore=0

Sub StartMission3
  WriteToLog "Mission 3", "START"
  GI_Color "blue",2

  Mission3Points(CurrentPlayer) = 0

  cLightMission3.blinkinterval=60
  bMissionMode = True
  MissionCompleted = False
  MissionTimeCurrent = 0
  MissionTimeMax = 75
  TotalChaseHits = 0
  ChasePupScore =0
  Mission1TargetPupCount3 = 4
  Gate001.collidable=False
  LightChaseShots
  PlaySong 7
  PuPEvent(509)
  PuPEvent(520)
  DOF 509, DOFOn
  ModeTextDelay3.Enabled=True

End Sub


Sub LightChaseShots
  WriteToLog "LightChaseShots", ""
  'Pick an available shot at semi-random
  dim xx,i
  'Set first shot
  i= int(rnd(1)*5)+1
  SetLightShot i,1,"mission"
  'Set second shot
  i= int(rnd(1)*5)+1
  If GetLightShot(i,"mission")=0 Then SetLightShot i,1,"mission"  : Exit Sub
  i= int(rnd(1)*5)+1
  If GetLightShot(i,"mission")=0 Then SetLightShot i,1,"mission" : Exit Sub
  For xx = 1 to 5
    If GetLightShot(xx,"mission")=0 Then SetLightShot xx,1,"mission" : Exit Sub
  Next
End Sub


Sub CheckChaseShot(nr)
    If Tilted Then Exit Sub
  If CurrentMission = 3 Then
    If GetLightShot(nr,"mission") = 1 Then
      WriteToLog "Mission 3", "Nice Shot!"
      Playsound "EFX_Ship_sound0" & rndint(1,4),0,CalloutVol,0,0,1,1,1
      AudioCallout "nice shot"
      LS_ShotLights.Play SeqBlinking,,3,20
      'Clear all shots
      SetLightShot 1,0,"mission"
      SetLightShot 2,0,"mission"
      SetLightShot 3,0,"mission"
      SetLightShot 4,0,"mission"
      SetLightShot 5,0,"mission"
      'Light some new chase shots
      LightChaseShots

      TotalChaseHits = TotalChaseHits + 1
      Mission1TargetPupCount3 = Mission1TargetPupCount3 -1
      ChasePupScore = ChasePupScore + 500000
      WriteToLog "Mission 3", "Total Hits = " & TotalChaseHits
      select case TotalChaseHits

        case 1 :
        AddMissionScore 3,score_Mission3_1Shot
        ShowAddScoreText = FormatScore(score_Mission3_1Shot * PointMultiplier)
        ScoreMissionPoints = "3"
        MissionPointsPup.Enabled=True
        ShowAddScore = 1
        creditblink=30
        showMissionVideo=11


        case 2 :
        AddMissionScore 3,score_Mission3_2Shots
        ShowAddScoreText = FormatScore(score_Mission3_2Shots * PointMultiplier)
        ScoreMissionPoints = "3"
        MissionPointsPup.Enabled=True
        ShowAddScore = 1
        creditblink=30
        showMissionVideo=11


        Case 3 :

        AddMissionScore 3,score_Mission3_3Shots
        ShowAddScoreText =FormatScore(score_Mission3_3Shots * PointMultiplier)
        ScoreMissionPoints = "3"
        MissionPointsPup.Enabled=True
        ShowAddScore = 1
        creditblink=30
        showMissionVideo=11

      Case 4,5 :
        AddMissionScore 3,score_Mission3_4Shots
        ShowAddScoreText =FormatScore( score_Mission3_4Shots * PointMultiplier)
        ScoreMissionPointsC = "3C"
        MissionPointsPupC.Enabled=True
        ShowAddScore = 1
        creditblink=130
        AddBonus bonus_Mission3_Completed
        MissionCompleted = True
        StopMission
      End Select
    End If
  End If
End Sub


Sub EndMission3
  PlayDefaultSong
  cLightMission3.blinkinterval=300
  bMissionMode = False
  SetLightShot 1,0,"mission"
  SetLightShot 2,0,"mission"
  SetLightShot 3,0,"mission"
  SetLightShot 4,0,"mission"
  SetLightShot 5,0,"mission"
  Gate001.collidable=True

  If ShowAddScore=0 Then: ShowAddScoreText = "COMPLETED"

  If MissionCompleted Then
    EOB_Missions = EOB_Missions + 1

    cLightMission3.state=1
    CheckMissionCompleteAward
    WriteToLog "Mission 3", "Completed"
    PuPEvent(526)
    DOF 509, DOFOff
    Mtimercountdown= "Default"
    PupMission= "Default"
    vpmtimer.addtimer 555, "StopVideosPlaying : showMissionVideo=12 '"
  Else
    cLightMission3.state=2
    WriteToLog "Mission 3", "Not Completed"
    StopVideosPlaying
    showMissionVideo=13
    PuPEvent(525)
    DOF 509, DOFOff
    Mtimercountdown= "Default"
    PupMission= "Default"
  End If
End Sub





'************* Cemetery - Mission 4 ***************



Dim TotalCemeteryHits
Dim Mission1TargetPupCount4

Sub StartMission4
  WriteToLog "Mission 4", "START"
  GI_Color "purple",1.2

  Mission4Points(CurrentPlayer) = 0
  Gate001.collidable=True
  Gate004.collidable=True

  ShowDust True
  MimaMagnetOn True

  cLightMission4.blinkinterval=60
  bMissionMode = True
  MissionCompleted = False
  MissionTimeCurrent = 0
  MissionTimeMax = 75
  TotalCemeteryHits = 0
  Mission1TargetPupCount4 = 4
  LightCemeteryShots
  PlaySong 13
  PuPEvent(508)
  PuPEvent(519)
  DOF 508, DOFOn
  ModeTextDelay4.Enabled=True

End Sub


Sub LightCemeteryShots
  WriteToLog "LightCemeteryShots", ""
  SetLightShot 4,1,"mission"
  SetLightShot 6,1,"mission"
End Sub


Sub CheckCemeteryShot(nr)
    If Tilted Then Exit Sub
  If CurrentMission = 4 Then
    If GetLightShot(nr,"mission") = 1 Then
      WriteToLog "Mission 4", "Nice Shot!"
      AudioCallout "nice shot"
      LS_ShotLights.Play SeqBlinking,,3,20
      If nr=4 Then LS_JP4.Play SeqBlinking,,10,20
      If nr=6 Then LS_JP6.Play SeqBlinking,,10,20

      TotalCemeteryHits = TotalCemeteryHits + 1
      Mission1TargetPupCount4 = Mission1TargetPupCount4 -1
      showMissionVideo=15

      WriteToLog "Mission 4", "Total Hits = " & TotalCemeteryHits

      If Not nr=6 Then
        AddMissionScore 4,score_Mission4_SkyShot
        AddBonus bonus_Mission4_SkyShot
        ShowAddScoreText = FormatScore(score_Mission4_SkyShot* PointMultiplier)
        Pupwithoutlightplay
        MissionPointsPup.Enabled=True
        ShowAddScore = 1
        creditblink=60

        WriteToLog "Mission 4", "Sky Hit"
      Else
        AddMissionScore 4,score_Mission4_RampShot
        ShowAddScoreText =FormatScore( score_Mission4_RampShot* PointMultiplier)
        Pupwithlightplay
        ShowAddScore = 1
        creditblink=60
        WriteToLog "Mission 4", "Ramp Hit"
      End If

      If TotalCemeteryHits = 3 Then
        AddBonus bonus_Mission4_Completed

      End If
      If TotalCemeteryHits >= 4 Then
        MissionCompleted = True
        StopMission
      End If
    End If
  End If
End Sub

Sub Pupwithoutlightplay
If HasPuP=False then exit Sub
  If TotalCemeteryHits < 4 Then
    ScoreMissionPoints = "4a"
    MissionPointsPup.Enabled=True
  End if
End Sub

Sub Pupwithoutlightend
If HasPuP=False then exit Sub
  If TotalCemeteryHits >= 4 Then
    ModeTextDelay4c1.Enabled=True
  End if
End Sub

Sub Pupwithlightplay
If HasPuP=False then exit Sub
  If TotalCemeteryHits < 4 Then
    ScoreMissionPoints = "4b"
    MissionPointsPup.Enabled=True
  End if
End Sub

Sub Pupwithlightend
If HasPuP=False then exit Sub
  If TotalCemeteryHits >= 4 Then
  ModeTextDelay4c2.Enabled=True
  End if
End Sub


Sub EndMission4

  ShowDust False
  MimaMagnetOn False

  PlayDefaultSong
  cLightMission4.blinkinterval=300
  bMissionMode = False
  Gate001.collidable=True
  Gate004.collidable=False
  SetLightShot 4,0,"mission"
  SetLightShot 6,0,"mission"

  If MissionCompleted Then
    EOB_Missions = EOB_Missions + 1
    Pupwithlightend
    Pupwithoutlightend
    cLightMission4.state=1
    CheckMissionCompleteAward
    WriteToLog "Mission 4", "Completed"
    PuPEvent(526)
    DOF 508, DOFOff
    Mtimercountdown= "Default"
    PupMission= "Default"
    vpmtimer.addtimer 555, "StopVideosPlaying : showMissionVideo=16 '"
  Else
    cLightMission4.state=2
    WriteToLog "Mission 4", "Not Completed"
    StopVideosPlaying
    showMissionVideo=17
    PuPEvent(525)
    DOF 508, DOFOff
    Mtimercountdown= "Default"
    PupMission= "Default"
  End If
End Sub





'************* Portal Room - Mission 5 ***************

Dim VascanShotCount
Dim PortalPupScore: PortalPupScore=0
Sub StartMission5
  CurrentMission = 5
  WriteToLog "Mission 5", "START"
  GI_Color "red",1

  vpmtimer.addtimer 1500, "GaldorShade false : GaldorQ1 = ""YM"" '"

  Mission5Points(CurrentPlayer) = 0
  PortalPupScore=0
  cLightMission5.blinkinterval=60
  SaveLaneInsertStates
  ClearLaneInsertStates

  bMissionMode = True
  MissionCompleted = False
  'LightPortalRoomShots
  pUnderPF.image = "under_pf1"
  hatch true
  PlaySong 9
  if vrroom=1 then hexfloor.opacity = 1200 : hexfloor.visible = true
  PuPEvent(511)
  PuPEvent(522)
  DOF 511, DOFOn
End Sub

'TargetVascan1.IsDropped = true
'DTDrop 4
'TargetVascan2.IsDropped = true
'TargetVascan3.IsDropped = true
'TargetVascan4.IsDropped = true
'TargetVascan5.IsDropped = true
'TargetVascan6.IsDropped = true

'SoundDropTarget_Release

Sub LightPortalRoomShots
  WriteToLog "LightPortalRoomShots", ""
  VascanShotCount = 0

'             PlaySoundAt "DropTarget_Up", TargetVascan1p : DTRaise 4 : cLightVascan1.state = 2
' vpmtimer.addtimer 600, "PlaySoundAt ""DropTarget_Up"", TargetVascan2p : DTRaise 5 : cLightVascan2.state = 2 '"
' vpmtimer.addtimer 1200, "PlaySoundAt ""DropTarget_Up"", TargetVascan3p : DTRaise 6 : cLightVascan3.state = 2 '"
'             PlaySoundAt "DropTarget_Up", TargetVascan4p : DTRaise 7 : cLightVascan4.state = 2
' vpmtimer.addtimer 600, "PlaySoundAt ""DropTarget_Up"", TargetVascan5p : DTRaise 8 : cLightVascan5.state = 2 '"
' vpmtimer.addtimer 1200, "PlaySoundAt ""DropTarget_Up"", TargetVascan6p : DTRaise 9 : cLightVascan6.state = 2 '"

              SoundDropTarget_Release(TargetVascan1p) : DTRaise 4 : cLightVascan1.state = 2
  vpmtimer.addtimer 600, "SoundDropTarget_Release(TargetVascan2p) : DTRaise 5 : cLightVascan2.state = 2 '"
  vpmtimer.addtimer 1200, "SoundDropTarget_Release(TargetVascan3p) : DTRaise 6 : cLightVascan3.state = 2 '"
              SoundDropTarget_Release(TargetVascan4p) : DTRaise 7 : cLightVascan4.state = 2
  vpmtimer.addtimer 600, "SoundDropTarget_Release(TargetVascan5p) : DTRaise 8 : cLightVascan5.state = 2 '"
  vpmtimer.addtimer 1200, "SoundDropTarget_Release(TargetVascan6p) : DTRaise 9 : cLightVascan6.state = 2 '"


End Sub

Sub LightPortalRoomShotsLeft
  WriteToLog "LightPortalRoomShotsLeft", ""
  VascanShotCount = 0

'             PlaySoundAt "DropTarget_Up", TargetVascan1p : DTRaise 4 : cLightVascan1.state = 2
' vpmtimer.addtimer 500, "PlaySoundAt ""DropTarget_Up"", TargetVascan2p : DTRaise 5 : cLightVascan2.state = 2 '"
' vpmtimer.addtimer 1000, "PlaySoundAt ""DropTarget_Up"", TargetVascan3p : DTRaise 6 : cLightVascan3.state = 2 '"
              SoundDropTarget_Release(TargetVascan1p) : DTRaise 4 : cLightVascan1.state = 2
  vpmtimer.addtimer 500, "SoundDropTarget_Release(TargetVascan2p) : DTRaise 5 : cLightVascan2.state = 2 '"
  vpmtimer.addtimer 1000, "SoundDropTarget_Release(TargetVascan3p) : DTRaise 6 : cLightVascan3.state = 2 '"


End Sub

Sub LightPortalRoomShotsRight
  WriteToLog "LightPortalRoomShotsRight", ""
  VascanShotCount = 0

'             PlaySoundAt "DropTarget_Up", TargetVascan4p : DTRaise 7 : cLightVascan4.state = 2
' vpmtimer.addtimer 500, "PlaySoundAt ""DropTarget_Up"", TargetVascan5p : DTRaise 8 : cLightVascan5.state = 2 '"
' vpmtimer.addtimer 1000, "PlaySoundAt ""DropTarget_Up"", TargetVascan6p : DTRaise 9 : cLightVascan6.state = 2 '"

              SoundDropTarget_Release(TargetVascan4p) : DTRaise 7 : cLightVascan4.state = 2
  vpmtimer.addtimer 500, "SoundDropTarget_Release(TargetVascan5p) : DTRaise 8 : cLightVascan5.state = 2 '"
  vpmtimer.addtimer 1000, "SoundDropTarget_Release(TargetVascan6p) : DTRaise 9 : cLightVascan6.state = 2 '"


End Sub

Sub CheckPortalRoomShot(nr)
    If Tilted Then Exit Sub
  If CurrentMission = 5 Then
    if VascanLights(nr-1).state = 2 Then
      VascanLights(nr-1).state = 0
'     PlaySoundAt "DropTarget_Down", Eval("TargetVascan" & nr & "p")
      SoundDropTarget(Eval("TargetVascan" & nr & "p"))
      WriteToLog "Mission 5", "Nice Shot!"
      AudioCallout "nice shot"
      playsound "EFX_click_bleep",0,CalloutVol,0,0,1,1,1
      VascanShotCount = VascanShotCount + 1
      PortalPupScore = 1000000
      PupPortalFinish
      PupPortalScore
      showMissionVideo=19+ int(Rnd(1)*2)
      ShowAddScoreText = FormatScore(score_Mission5_VascanShot* PointMultiplier)
      ShowAddScore = 1
      creditblink=30


      AddMissionScore 5,score_Mission5_VascanShot

      If VascanShotCount = 6 Then
        PlaySound "EFX_landing",0,CalloutVol,0,0,1,1,1
        VascanShotCount = 0
        pUnderPF.blenddisablelighting = 0
        GaldorReset=true:GaldorQ1 = "Yk"
        MissionCompleted = True
        AddBonus bonus_Mission5_Completed
        ShowAddScoreText = "COMPLETED"
        bFlippersEnabled = False
        Flipper1.RotateToStart
        LeftFlipperMini.RotateToStart
        LeftFlipper.RotateToStart
        RightFlipperMini.RotateToStart
        RightFlipper.RotateToStart
      End If

    End If
  elseif WizardPhase = 3 Then
    if VascanLights(nr-1).state = 2 Then
      VascanLights(nr-1).state = 0
'     PlaySoundAt "DropTarget_Down", Eval("TargetVascan" & nr & "p")
      SoundDropTarget(Eval("TargetVascan" & nr & "p"))
      WriteToLog "Wizardphase3", "Nice Shot!"
      'AudioCallout "nice shot"
      playsound "EFX_click_bleep",0,CalloutVol,0,0,1,1,1
      VascanShotCount = VascanShotCount + 1
      PortalPupScore = 1000000
      PupPortalScore
      AddMissionScore 5,score_Mission5_VascanShot

      If VascanShotCount = 3 Then
        'PlaySound("EFX_landing")
        VascanShotCount = 0
        if cLightUnderPFShot1.state = 0 then
          cLightUnderPFShot1.state = 2
        elseif cLightUnderPFShot2.state = 0 Then
          cLightUnderPFShot2.state = 2
        end if
      End If
    end if
  End If
End Sub

Sub PupPortalScore
If HasPuP=False then exit Sub
  If VascanShotCount < 6 Then
  ScoreMissionPoints = "5"
  MissionPointsPup.Enabled=True
  End If
End Sub

Sub PupPortalFinish
If HasPuP=False then exit Sub
  If VascanShotCount = 6 Then
    ScoreMissionPoints = "5c"
    MissionPointsPup.Interval= 4500
    MissionPointsPup.Enabled=True
  End if
End Sub


hexfloor.visible = false
Sub EndMission5
  PlayDefaultSong
  cLightMission5.blinkinterval=300
  bMissionMode = False
  gioff_underpf_boost false
  RestoreLaneInsertStates
  if vrroom=1 then hexfloor.opacity = 0 : hexfloor.visible = false

' TargetVascan1.IsDropped = true
  DTDrop 4 : DTDrop 5 : DTDrop 6
  DTDrop 7 : DTDrop 8 : DTDrop 9
' TargetVascan2.IsDropped = true
' TargetVascan3.IsDropped = true
' TargetVascan4.IsDropped = true
' TargetVascan5.IsDropped = true
' TargetVascan6.IsDropped = true

  dim xx
  for each xx in VascanLights
    xx.state = 0
  next

  vpmtimer.addtimer 2000, "GaldorShade true '"

  If MissionCompleted Then
    EOB_Missions = EOB_Missions + 1

    cLightMission5.state=1
    CheckMissionCompleteAward
    WriteToLog "Mission 5", "Completed"
    PuPEvent(526)
    DOF 511, DOFOff
    vpmtimer.addtimer 555, "StopVideosPlaying : showMissionVideo=21 '"
  Else
    cLightMission5.state=2
    WriteToLog "Mission 5", "Not Completed"
    StopVideosPlaying
    showMissionVideo=22
    PuPEvent(525)
    DOF 511, DOFOff
  End If
End Sub



'************* Heart Of Steel - Mission 6 ***************


Dim TotalHeartOfSteelHits
Dim Mission1TargetPupCount6
Dim HeartPupScore: HeartPupScore=0
Dim HeartBeatIndex
Dim HeartBeat(6) : HeartBeat(0)=1 : HeartBeat(1)=0 : HeartBeat(2)=1 : HeartBeat(3)=1 : HeartBeat(4)=0 : HeartBeat(5)=0 : HeartBeat(6)=0

Sub StartMission6
  WriteToLog "Mission 6", "START"
  GI_Color "purple",1.2

  Mission6Points(CurrentPlayer) = 0

  cLightMission6.blinkinterval=60
  bMissionMode = True
  MissionCompleted = False
  MissionTimeCurrent = 0
  MissionTimeMax = 75
  TotalHeartOfSteelHits = 0
  Mission1TargetPupCount6 = 25
  HeartPupScore=0
  Gate001.collidable=True
  Gate004.collidable=True

  HeartBeatIndex = 0
  Mission6BumperTimer.Interval = 200
  Mission6BumperTimer.Enabled = True

  Mission6LagoTimer.Interval = RndInt(2500,6000)
  Mission6LagoTimer.Enabled = True

  PlaySong 8
  PuPEvent(510)
  PuPEvent(521)
  DOF 510, DOFOn
  ModeTextDelay6.Enabled=True
End Sub


dim prevBeatDown : prevBeatDown = 1
Sub Mission6BumperTimer_Timer
  FlBumperFadeTarget(1) = HeartBeat(HeartBeatIndex)
  FlBumperFadeTarget(2) = HeartBeat(HeartBeatIndex)
  FlBumperFadeTarget(3) = HeartBeat(HeartBeatIndex)
  if HeartBeat(HeartBeatIndex) = 0 Then
    if prevBeatDown = 0 then
      PlaySound "EFX_heartbeat_down" & RndInt(1,4),0,CalloutVol,0,0,1,1,1
      prevBeatDown = 1
    end if
  Else
    if prevBeatDown = 1 then
      PlaySound "EFX_heartbeat_up" & RndInt(1,4),0,CalloutVol,0,0,1,1,1
      prevBeatDown = 0
    end if
  end if

  HeartBeatIndex = HeartBeatIndex + 1
  If HeartBeatIndex > 6 Then HeartBeatIndex = 0
End Sub

Sub Mission6LagoTimer_Timer
  playsound "EFX_lago_in_pain_0" & RndInt(1,5),0,CalloutVol,0,0,1,1,1
  Mission6LagoTimer.Interval = RndInt(2500,6000)
  startB2S 2
  If VRRoom > 0 Then
    VRBGFL2
  End If
End Sub

Sub CheckHeartOfSteelShot
    If Tilted Then Exit Sub
  If CurrentMission = 6 Then
    TotalHeartOfSteelHits = TotalHeartOfSteelHits + 1
    Mission1TargetPupCount6 = Mission1TargetPupCount6 -1
'   showMissionVideo=24
'     ShowAddScoreText = FormatScore(score_Mission6_Bumper* PointMultiplier)
'     ShowAddScore = 1
'     creditblink=30

    HeartPupScoring
    AddMissionScore 6,score_Mission6_Bumper
    WriteToLog "Mission 6", "Total Hits = " & TotalHeartOfSteelHits
    If TotalHeartOfSteelHits = 7 Then Mission6BumperTimer.Interval = 150
    If TotalHeartOfSteelHits = 14 Then Mission6BumperTimer.Interval = 100
    If TotalHeartOfSteelHits = 21 Then Mission6BumperTimer.Interval = 50
    If TotalHeartOfSteelHits = 20 Then Mission6BumperTimer.Interval = 30
    If TotalHeartOfSteelHits = 25 Then
      MissionCompleted = True
      PupHeartFinish
      addbonus bonus_Mission6_Completed
      AddMissionScore 6,score_Mission6_25Bumpers
      ShowAddScoreText = FormatScore(score_Mission6_25Bumpers* PointMultiplier)
      ShowAddScore = 1
      creditblink=130

      StopMission
    End If
  End If
End Sub

Sub HeartPupScoring
If HasPuP=False then exit Sub
  If TotalHeartOfSteelHits < 25 Then
  ScoreMissionPoints = "6"
  MissionPointsPup.Enabled=True
  End If
End Sub


Sub PupHeartFinish
If HasPuP=False then exit Sub
  ScoreMissionPoints = "6c"
  MissionPointsPup.Interval=3000
  MissionPointsPup.Enabled=True
End Sub





Sub EndMission6
  PlayDefaultSong
  cLightMission6.blinkinterval=300
  bMissionMode = False
  Gate001.collidable=True
  Gate004.collidable=False
  Mission6BumperTimer.Enabled = False
  Mission6LagoTimer.Enabled = False
  FlBumperFadeTarget(1) = 1
  FlBumperFadeTarget(2) = 1
  FlBumperFadeTarget(3) = 1
  prevBeatDown = 1

  If MissionCompleted Then
    EOB_Missions = EOB_Missions + 1

    cLightMission6.state=1
    CheckMissionCompleteAward
    WriteToLog "Mission 6", "Completed"
    PuPEvent(526)
    DOF 510, DOFOff
    Mtimercountdown= "Default"
    PupMission= "Default"
    MBPupCheck.Enabled=true
    vpmtimer.addtimer 555, "StopVideosPlaying : showMissionVideo=25 '"
  Else
    cLightMission6.state=2
    WriteToLog "Mission 6", "Not Completed"
    StopVideosPlaying
    showMissionVideo=26
    PuPEvent(525)
    DOF 510, DOFOff
    Mtimercountdown= "Default"
    PupMission= "Default"
    SkillPupFlash="Default"
  End If
End Sub




'************* Betrayal (Combo Mission) - Mission 7 ***************


Dim TotalBetrayalMissionHits



Sub StartMission7
  WriteToLog "Mission 7", "START"
  GI_Color "blue",2

  Mission7Points(CurrentPlayer) = 0

  cLightMission7.blinkinterval=60
  bMissionMode = True
  MissionCompleted = False
  MissionTimeCurrent = 0
  MissionTimeMax = 75
  Mission7ShotTimer.Enabled = False
  TotalBetrayalMissionHits = 0
  Gate001.collidable=False
  LightBetrayalMissionSetup
  ''PlaySong 1
  PlaySongLoop 12, 351.6, 402
  PuPEvent(507)
  PuPEvent(584)
  DOF 507, DOFOn
  ModeTextDelay7.Enabled=true
End Sub


Sub LightBetrayalMissionSetup
  SetLightShot 3,1,"mission"
End Sub


Sub LightBetrayalMissionShot1
  WriteToLog "LightBetrayalMissionShot1", ""
  SetLightShot 3,0,"mission"
  SetLightShot 1,1,"mission"
  Mission7ShotTimer.Enabled = True
End Sub


Sub Mission7ShotTimer_Timer
  Mission7ShotTimer.Enabled = False
  SetLightShot 1,0,"mission"
  LightBetrayalMissionSetup
End Sub


Sub CheckBetrayalMissionShot(nr)
    If Tilted Then Exit Sub
  If CurrentMission = 7 Then
    If GetLightShot(nr,"mission") = 1 Then

      showMissionVideo=28

      WriteToLog "Mission 7", "Nice Shot!"
      AudioCallout "nice shot"
      If nr=1 Then LS_ShotLights.Play SeqBlinking,,3,20 : TotalBetrayalMissionHits = TotalBetrayalMissionHits + 1
      If nr=3 Then LightBetrayalMissionShot1

      AddMissionScore 7,score_Mission7_1stShot
      BetrayPupScoring
      ShowAddScoreText =FormatScore(score_Mission7_1stShot* PointMultiplier)
      ShowAddScore = 1
      creditblink=30

      WriteToLog "Mission 7", "Total Hits = " & TotalBetrayalMissionHits
'     If TotalBetrayalMissionHits = 1 Then
'       AddMissionScore 7,score_Mission7_2ndShot
'       ShowAddScoreText = FormatScore(score_Mission7_1stShot* PointMultiplier)
'       ShowAddScore = 1
'       creditblink=30
'     End If
      If TotalBetrayalMissionHits >= 1 Then
        AddMissionScore 7,score_Mission7_2combos
        ShowAddScoreText = FormatScore(score_Mission7_2combos* PointMultiplier)
        PupBetrayFinish
        ShowAddScore = 1
        creditblink=130
        AddBonus bonus_Mission7_Completed
        MissionCompleted = True
        StopMission
      End If
    End If
  End If
End Sub

Sub BetrayPupScoring
If HasPuP=False then exit Sub
  ScoreMissionPoints = "7"
  MissionPointsPup.Enabled=True
End Sub


Sub PupBetrayFinish
If HasPuP=False then exit Sub
  ScoreMissionPoints = "7c"
  MissionPointsPup.Interval=3000
  MissionPointsPup.Enabled=True
End Sub


Sub EndMission7
  PlayDefaultSong
  Mission7ShotTimer.Enabled = False
  cLightMission7.blinkinterval=300
  bMissionMode = False
  SetLightShot 1,0,"mission"
  SetLightShot 2,0,"mission"
  SetLightShot 3,0,"mission"
  SetLightShot 4,0,"mission"
  SetLightShot 5,0,"mission"
  Gate001.collidable=True

  If MissionCompleted Then
    EOB_Missions = EOB_Missions + 1

    cLightMission7.state=1
    CheckMissionCompleteAward
    WriteToLog "Mission 7", "Completed"
    vpmtimer.addtimer 555, "StopVideosPlaying : showMissionVideo=29 '"
    Mtimercountdown= "Default"
    PuPEvent(526)
    DOF 507, DOFOff
  Else
    cLightMission7.state=2
    WriteToLog "Mission 7", "Not Completed"
    StopVideosPlaying
    showMissionVideo=30
    Mtimercountdown= "Default"
    PuPEvent(525)
    DOF 525, DOFPulse
    DOF 507, DOFOff

  End If
End Sub





'************* Tracy Glare - Double Scoring Mode ***************



Sub StartDoubleScoring
  WriteToLog "DoubleScoring","Started"
  'DTDrop 1 'these start DT timers and break the sync
  'DTDrop 2
  'DTDrop 3
  'WriteToLog "ScavangeTargets","Dropped"

  cpMimaTarget.blinkinterval=250
  cpMimaTarget.state = 2
  PointMultiplier = 2
  tracySpeed = 1      'value to increase as mission proceeds. Maybe 1-10.
  TracyTrackingTime = 0
  TracyTimer.enabled = true
  TracyTargetLocked = 0   'value that increases on lock, and decreases if tracy looses tracking
  solBloodTrail true
End Sub


Sub EndDoubleScoring
  WriteToLog "DoubleScoring","Ended"
  PointMultiplier = 1
  cpMimaTarget.blinkinterval=250
  cpMimaTarget.state = 0

  TracyAngle 0
  Primitive019.blenddisablelighting = tracyEyesBrightness
  UpdateMaterial "tracyeyes"    ,0,0,0,0,0,0,1,RGB(17,217,27),0,0,False,True,0,0,0,0
  ptracyCross.visible = false
  cTracyCross.state = 0

  DTRaise 1
  DTRaise 2
  DTRaise 3
  RandomSoundDropTargetLeft(TargetScavenge2p)
  WriteToLog "ScavangeTargets","Popped Up"
  cLightMystery2.state = 1
  WriteToLog "ScavangeTargets","Light 2 Solid"
  If cLightMystery1.state = 1 and cLightMystery2.state = 1 and cLightMystery3.state = 1 Then
    cLightMystery1.state = 0
    cLightMystery2.state = 0
    cLightMystery3.state = 0
    WriteToLog "ScavangeTargets","Lights Out"
  End If
  tracySpeed = 1
  TracyTrackingTime = 0
  TracyTimer.enabled = false
  solBloodTrail false
End Sub





'************* Super Spinners Mode ***************



Sub StartSuperSpinners
  WriteToLog "SuperSpinners","Started"
  'DTDrop 1
  'DTDrop 2
  'DTDrop 3
  'WriteToLog "ScavangeTargets","Dropped"

  LS_SuperSpinner.UpdateInterval = 40
  LS_SuperSpinner.Play SeqBlinking,,1000000,40
  SpinnerMultiplier = 5
  SuperSpinnerTimer.interval = 30000
  SuperSpinnerTimer.enabled = true
  PuPEvent(470)
  DOF 470, DOFPulse
End Sub

sub SuperSpinnerTimer_timer
  EndSuperSpinners
  SuperSpinnerTimer.enabled = false
end sub

Sub EndSuperSpinners
  WriteToLog "SuperSpinners","Ended"
  SpinnerMultiplier = 1
  StopTracySpins
  DTRaise 1
  DTRaise 2
  DTRaise 3
  RandomSoundDropTargetLeft(TargetScavenge2p)
  LS_SuperSpinner.stopplay
  WriteToLog "ScavangeTargets","Popped Up"
  cLightMystery3.state = 1
  WriteToLog "ScavangeTargets","Light 3 Solid"
  If cLightMystery1.state = 1 and cLightMystery2.state = 1 and cLightMystery3.state = 1 Then
    cLightMystery1.state = 0
    cLightMystery2.state = 0
    cLightMystery3.state = 0
    WriteToLog "ScavangeTargets","Lights Out"
  End If
End Sub







'**********************************************************************************************************
'  Lazers
'**********************************************************************************************************

dim lazerBloomGain
dim lazerPos : lazerPos = 0
dim lazerSpeed : lazerSpeed = 20
dim halfLazers : halfLazers = 0

LazersInit

sub LazersInit
  dim L, Lcount
  Lcount = 0
  Const Lsep = 9

  For each L in Lazers
    L.blenddisablelighting = 25
    L.z = (Lsep * Lcount) + 55
    L.size_y = 10
    L.TransY = -200
    L.RotZ = -61
    L.visible = false
    Lcount = Lcount + 1
    UpdateMaterial "Lazer"    ,0,0,0,0,0,0,0.3,RGB(11,249,22),0,0,False,True,0,0,0,0
  next
  FlasherbloomLazer.visible = false
end sub

sub PulseLaser(Enabled)
  If Enabled Then
    lazer009.size_z = 0.0001
    lazer010.size_z = 0.01
    lazerBloomGain = 1
    LazerTimer.enabled = true
    If CurrentMission = 1 Then
      SetLazerReady
      Mission1ShotEnabled = True
      BladerunnerShotTimer.Enabled = True
    else
      BladerunnerShotTimer.Enabled = True
    End If
  end if
end sub



sub LazerTimer_timer
  dim L
  dim lazerOpacity
  lazerOpacity = 0.3

  FlasherbloomLazer.opacity = 30000 * lazerBloomGain
  lazerBloomGain = max(0,lazerBloomGain * 0.8 - 0.01)
  UpdateMaterial "Dust2"    ,0,0,0,0,0,0,lazerBloomGain^0.3,RGB(82,252,10),0,0,False,True,0,0,0,0
  'Lazer009.image = "grain" & RndInt(1,3)

  lazerPos = lazerPos + lazerSpeed

  if lazerPos < 160 Then
    DustLayer1.material = "Dust2"
    DustLayer1.visible = true
    For each L in Lazers
      L.visible = true
      L.blenddisablelighting = 375 '275
      L.size_y = lazerPos
      L.TransY = (-lazerPos*10) - 270
    next
    Lazer009.image = "grain" & RndInt(1,3)
    Lazer009.blenddisablelighting = 0
    Lazer010.blenddisablelighting = 0
    FlasherbloomLazer.visible = true

  else
    For each L in Lazers
      L.blenddisablelighting = lazerPos * RndInt(2,4) '4
      'L.TransY = (-lazerPos*10) - 100
      L.TransY = (-lazerPos/3) - 1610
      lazerOpacity = (0.3/(lazerPos/10))-0.003
      if lazerOpacity < 0 then lazerOpacity = 0
      UpdateMaterial "Lazer"    ,0,0,0,0,0,0,lazerOpacity,RGB(11,249,22),0,0,False,True,0,0,0,0


      'debug.print L.name & "lazerpos: " & lazerPos &" --> "& lazerOpacity
    next
    Lazer009.blenddisablelighting = lazerPos / 10 * RndInt(0,30)
    Lazer010.blenddisablelighting = lazerPos * 2
    'Lazer009.TransY = (-lazerPos) - 100
  end if



  if lazerPos > 570 Then
    'Debug.print "eka: " & Lazer008.TransY
    'Debug.print "toka: " & Lazer007.TransY
    For each L in Lazers
      if halfLazers = 0 then
        L.TransY = (-lazerPos/6) - 1690
        halfLazers = 1
      Else
        L.TransY = (-lazerPos/3) - 1630
        halfLazers = 0
      end if

    next
    'Debug.print lazerPos
  end if

  if lazerPos > 1000 Or lazerOpacity = 0 Then
    FlasherbloomLazer.visible = false
    LazerTimer.enabled = false
    if Not DustLayer3.visible then  'if mission with dust is ongoing
      DustLayer1.visible = false
    end if
    DustLayer1.material = "Dust1"
    lazerPos = 0
    LazersInit
  end If
end sub

'UpdateLaserGun 120:pulselaser true

Dim BallInGun, GPos, BallInGunRadius, CannonStatus, GunSightRadius
BallInGunRadius = SQR((PCannon.X - Sw45.X)^2 + (PCannon.Y - Sw45.Y)^2)
GunSightRadius = SQR((PCannon.X - LightAutofireWarning.X)^2 + (PCannon.Y - LightAutofireWarning.Y)^2)


Sub ShootGun(Enabled)
  If Enabled Then
    If NOT IsEmpty(BallInGun) Then
      WriteToLog "ShootGun", " "
      DOF 144,2
      DOF 153,0
      startB2S 3
      If VRRoom > 0 Then
        VRBGFL3
      End If
      If CurrentMission=1 Then showMissionVideo=4
      PlaySoundAt SoundFX("LaserBlast1", DOFContactors), sw45
      PuPEvent(582)
      DOF 877, DOFPulse
      BlastBall.Interval = 30 : BlastBall_Timer()

      PulseLaser true
      'vpmtimer.addtimer 500, "Sw45.kick 300 + GPos, 50 '"    'moving this to own timer
      'this delay happens after the gun movement has stopped. Not affecting targeting, just a visual thing.
      ShootGunDelay.enabled = true
      'gunpark = 1
      BallInGun = Empty
      EnableBallSaver 1
      cLightAutofireWarning.state = 0
      LazerAutofireTimer1.Enabled = False
    End If
  End If
End Sub

sub ShootGunDelay_timer
  Sw45.kick 300 + GPos, 50
  ShootGunDelay.enabled = false
end Sub

dim blastballPos : blastballPos = 2100
const blastballspeed = 50
Sub BlastBall_Timer()
  if Not BlastBall.Enabled Then BlastBall.Enabled = true
  blastballPos = blastballPos - blastballspeed

  PlaySound "EFX_lbass3", 0, BlastSoundLevel * VolumeDial * CalloutVol, AudioPanPosX(470), 0, 0, 1, 0, AudioFadePosY(blastballPos)

  if blastballPos < -200 Then
    blastballPos = 2100
    BlastBall.Enabled = False
  end if

End Sub


dim GunLastPos, GunMotorStatus
Sub UpdateLaserGun(acurrpos)
  GPos = acurrpos / 2.5
  PCannon.RotY = GPos
  PClapet.RotY = GPos
  pGunSight.RotZ = GPos

  LightAutofireWarning.X = PCannon.X - GunSightRadius * Cos((GPos+30)*Pi/180)
  LightAutofireWarning.Y = PCannon.Y - GunSightRadius * Sin((GPos+30)*Pi/180)

  dim L : For each L in Lazers : L.RotZ = GPos - 60 : next
  'PCannonshadow.RotY  = GPos
  if GPos > 5 Then
    CannonStatus = True
  Else
    CannonStatus = False
  end If
  If Not IsEmpty(BallInGun) Then
    BallInGun.X = PCannon.X - BallInGunRadius * Cos((GPos+30)*Pi/180)
    BallInGun.Y = PCannon.Y - BallInGunRadius * Sin((GPos+30)*Pi/180)
  End If

  if round(GunLastPos,1) = round(acurrpos,1) Then
    if GunMotorStatus Then
      StopSound "EFX_GunMotor"
      GunMotorStatus = False
'     debug.print "same pos, motor off"
    end if
  Else
    'play once
    if Not GunMotorStatus Then
      vpmtimer.addtimer RndInt(70,170), "PlaySoundAtLevelStatic SoundFX(""EFX_GunMotor"",DOFGear), GunMotorSoundLevel, PCannon '"
      GunMotorStatus = true
'     debug.print "diff pos, motor on"
    end if
  end if


  GunLastPos = acurrpos
End Sub

Sub Sw45_hit()
  WriteToLog "Sw45_hit", "Ball in gun"
  PlaySound "EFX_gun_loading2",0,CalloutVol,0,0,1,1,1
  Set BallInGun = ActiveBall
  DOF 153,1

  LightAutofireWarning.color = rgb(19,255,0)
  LightAutofireWarning.colorfull = rgb(142,255,142)
  LightBlastButton.color = rgb(19,255,0)
  LightBlastButton.colorfull = rgb(142,255,142)

  cLightAutofireWarning.BlinkInterval = 250
  cLightAutofireWarning.state = 2

  LaserUpdate.Enabled = True
  LazerAutofireTimer1.Interval = 6000 + RndInt(1,1000)
  LazerAutofireTimer1.Enabled = True
End Sub

dim gunposi, gundir, gunpark
gunposi = 0
gundir = 0
gunpark = 0
Const gunright = 200
Const gunleft = 80


'sub testtimer_timer
' PCannon.roty = PCannon.roty + 1
'end Sub

sub LaserUpdate_timer()

  If NOT IsEmpty(BallInGun) Then
    if gundir = 0 Then
      gunposi = gunposi + 2
      if gunposi > gunright Then
        StopSound "EFX_GunMotor" : GunMotorStatus = False
'       debug.print "dir change, motor off"
        gundir = 1
      end if
    Else
      gunposi = gunposi - 2
      if gunposi < gunleft Then
        StopSound "EFX_GunMotor" : GunMotorStatus = False
'       debug.print "dir change, motor off"
        gundir = 0
      end if
    end if
  Else
    if LazerTimer.enabled Then
      'lazer shooting
    else
      'park gun after shot
      gunposi = gunposi - 2
      if gunposi = 0 then
        UpdateLaserGun 0
        me.Enabled = False
      end if
    end if
  end If


  UpdateLaserGun gunposi
end Sub


Sub LazerAutofireTimer1_Timer
  LazerAutofireTimer1.Enabled = False
  LazerAutofireTimer2.Enabled = True
  LightAutofireWarning.color = rgb(255,19,0)
  LightAutofireWarning.colorfull = rgb(255,142,142)
  LightBlastButton.color = rgb(255,19,0)
  LightBlastButton.colorfull = rgb(255,142,142)
  cLightAutofireWarning.BlinkInterval = 50
  cLightAutofireWarning.state = 2
End Sub

Sub LazerAutofireTimer2_Timer
  LazerAutofireTimer1.Enabled = False
  LazerAutofireTimer2.Enabled = False
  cLightAutofireWarning.state = 0
  ShootGun True
End Sub

'****************************************
' Galdor Shade
'****************************************

dim GaldorShadeTargetAngle, PrevGaldorShadeAngle, GaldorShadeAtEndPoint
GaldorShadeAtEndPoint = true
GaldorShadeTargetAngle = 0
PrevGaldorShadeAngle = 0

sub GaldorShade(Enabled)
  if Enabled Then
    GaldorShadeTargetAngle = 0
    GaldorQ1=""
    if PrevGaldorShadeAngle <> GaldorShadeTargetAngle And GaldorShadeAtEndPoint then
      PlaySoundAtLevelStatic ("EFX_shade_close"), 1, pGlobeShade
    end if
  Else
    GaldorShadeTargetAngle = 140
    if PrevGaldorShadeAngle <> GaldorShadeTargetAngle And GaldorShadeAtEndPoint then
      PlaySoundAtLevelStatic ("EFX_shade_open"), 1, pGlobeShade
    end if
  end If
end sub

sub UpdateGaldorShade()

  if PrevGaldorShadeAngle <> GaldorShadeTargetAngle Then
    GaldorShadeAtEndPoint = false
    PrevGaldorShadeAngle = pGlobeShade.rotX
    'debug.print GaldorShadeTargetAngle & " <--- " & PrevGaldorShadeAngle

    if round(PrevGaldorShadeAngle,0) < GaldorShadeTargetAngle Then
      pGlobeShade.rotX = PrevGaldorShadeAngle + ((GaldorShadeTargetAngle - PrevGaldorShadeAngle)/10) + 0.1
    Elseif round(PrevGaldorShadeAngle,0) > GaldorShadeTargetAngle Then
      pGlobeShade.rotX = PrevGaldorShadeAngle - ((PrevGaldorShadeAngle - GaldorShadeTargetAngle)/10) - 0.1
    Elseif round(PrevGaldorShadeAngle,0) = round(GaldorShadeTargetAngle,0) then
      PrevGaldorShadeAngle = GaldorShadeTargetAngle
      pGlobeShade.rotX = GaldorShadeTargetAngle
      GaldorShadeAtEndPoint = true
    Else
      PrevGaldorShadeAngle = GaldorShadeTargetAngle
      pGlobeShade.rotX = GaldorShadeTargetAngle
      GaldorShadeAtEndPoint = true
    end if
    ' stop glador movie if shade is on
    if PrevGaldorShadeAngle > 1 then
      galdor.enabled = true
    Else
      galdor.enabled = false
    end if
  end if
end sub


'****************************************
' Real Time updates
'****************************************

dim EOBGaugeTargetAngle, CurrEOBGaugeAng
EOBGaugeTargetAngle = 60

Sub UpdateEOBGauge

  if CurrEOBGaugeAng <> EOBGaugeTargetAngle Then
    CurrEOBGaugeAng = gauge_rollinside.rotx

    'debug.print gauge_rollinside.rotx

    if round(CurrEOBGaugeAng,0) < EOBGaugeTargetAngle Then
      gauge_rollinside.rotx = CurrEOBGaugeAng + (EOBGaugeTargetAngle-CurrEOBGaugeAng)/50 + 0.1
    Elseif round(CurrEOBGaugeAng,0) > EOBGaugeTargetAngle Then
      gauge_rollinside.rotx = CurrEOBGaugeAng - (CurrEOBGaugeAng-EOBGaugeTargetAngle)/50 - 0.1
    Elseif round(CurrEOBGaugeAng,0) = round(EOBGaugeTargetAngle) Then
      CurrEOBGaugeAng = EOBGaugeTargetAngle
      gauge_rollinside.rotx = EOBGaugeTargetAngle
    else
      CurrEOBGaugeAng = EOBGaugeTargetAngle
      gauge_rollinside.rotx = EOBGaugeTargetAngle
    end if

  end if

end sub

'****************************************
' Real Time updates
'****************************************

Sub FrameTimer_Timer()
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
  UpdateEOBGauge
  UpdateGaldorShade
  UpdateFlashersByDTs

  LFlip.roty = LeftFlipper.currentangle
  RFlip.roty = RightFlipper.currentangle
  LFlip1.roty = Flipper1.currentangle

  if bPortalFlippersEnabled then
    LFliprUnder.rotz = LeftFlipperMini.CurrentAngle
    LFlipbUnder.rotz = LeftFlipperMini.CurrentAngle
    RFliprUnder.rotz = RightFlipperMini.CurrentAngle
    RFlipbUnder.rotz = RightFlipperMini.CurrentAngle
  end if
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle

  pSpinner.RotX = spinner1.currentangle

  pSpinnerRod.TransX = sin( (spinner1.CurrentAngle+180) * (2*PI/360)) * 7
  pSpinnerRod.TransY = sin( (spinner1.CurrentAngle- 90) * (2*PI/360)) * 7

  ' VR_Primary_plunger.Y = 15.34 + (5* Plunger.Position) -20
  ' LampTimer
  if VRRoom = 1 then
    updateVRDust
    UpdateDisorient
    if MimaMagnetDisorient.timerenabled then
      vr360_room.blenddisablelighting = lampz.lvl(69) * (17 + RndInt(1,2))  + 1
    end if
  end if

End Sub

dim VRDustCntr : VRDustCntr = 0
sub updateVRDust
  VRDustCntr = VRDustCntr + 0.0005
  vr360_RoomHaze.rotx = VRDustCntr
  vr360_RoomHaze.roty = VRDustCntr * 2
  vr360_RoomHaze.rotz = VRDustCntr * -1.5
  vr360_RoomHaze2.rotx = VRDustCntr * 1.1
  vr360_RoomHaze2.roty = VRDustCntr * -1.2
  vr360_RoomHaze2.rotz = VRDustCntr * 1.5
  if VRDustCntr > 100000 then VRDustCntr = 0
end sub
'
'dim vrtunnellightcounter : vrtunnellightcounter = 0
'dim vrtunnellightdir : vrtunnellightdir = 1

Sub GameTimer_Timer()
  Cor.Update
  RollingUpdate
  DoDTAnim
  SongTimerTimer
  'pupil latency counters

  if pupilDcounter > 0 then
    pupilDcounter = pupilDcounter - 1
    if pupilDcounter = 0 Then
      cFlasherPupilLatencyD.state = 0
      SetInsertBlooms gilvl
    end if
  end if

  if pupilBcounter > 0 then
    pupilBcounter = pupilBcounter - 1
    if pupilBcounter = 0 then
      cFlasherPupilLatencyB.state = 0
      SetInsertBlooms gilvl
    end if
  end if
  'vr tunnel light flicker
' if vrroom = 1 Then
'   pTunnel_light.opacity = 100 + (100 * gilvl) + (vrtunnellightcounter * 2 * gilvl) + rndint(0,30)
'   vrtunnellightcounter = vrtunnellightcounter + vrtunnellightdir
'   if vrtunnellightcounter => 100 then vrtunnellightdir = -1
'   if vrtunnellightcounter =< 0 then vrtunnellightdir = 1
' end if

End Sub

sub SetInsertBlooms(ByVal levelGI)
  if bInsertBlooms then
    dim x,i
    for Each x in InsertBlooms 'restore blooms
      i=i+1
      x.Intensity = (Insert_Blooms(i) * (1-levelGI))*(1-Insert_Bloom_x_GIon) + (Insert_Blooms(i) * Insert_Bloom_x_GIon)
'     if i = 1 then debug.print round((Insert_Blooms(i) * (1-levelGI))*(1-Insert_Bloom_x_GIon) + (Insert_Blooms(i) * Insert_Bloom_x_GIon),2)
    next
  end if
end sub

sub UpdateFlashersByDTs
  if DTIsDropped(1) then
    LightFlash6b1d.state = 1
  Else
    LightFlash6b1d.state = 0
  end if

  if DTIsDropped(2) then
    LightFlash6b2d.state = 1
  Else
    LightFlash6b2d.state = 0
  end if

  if DTIsDropped(3) then
    LightFlash6b3d.state = 1
  Else
    LightFlash6b3d.state = 0
  end if
end sub

'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx7" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
' * with at least as many objects each as there can be balls, including locked balls
' Ensure you have a timer with a -1 interval that is always running

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
'***These must be organized in order, so that lights that intersect on the table are adjacent in the collection***
'***If there are more than 3 lights that overlap in a playable area, exclude the less important lights***
' This is because the code will only project two shadows if they are coming from lights that are consecutive in the collection, and more than 3 will cause "jumping" between which shadows are drawn
' The easiest way to keep track of this is to start with the group on the right slingshot and move anticlockwise around the table
' For example, if you use 6 lights: A & B on the left slingshot and C & D on the right, with E near A&B and F next to C&D, your collection would look like EBACDF
'
'G        H                     ^ E
'                             ^ B
' A    C                        ^ A
'  B    D     your collection should look like  ^ G   because E&B, B&A, etc. intersect; but B&D or E&F do not
'  E      F                       ^ H
'                             ^ C
'                             ^ D
'                             ^ F
'   When selecting them, you'd shift+click in this order^^^^^

'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
'Sub FrameTimer_Timer()
' If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
'End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
'Const lob = 0  'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 '2 = flasher image shadow, but it moves like ninuzzu's

Const fovY          = 0   'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor     = 0.99  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source
Const lob         = 1

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


'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******
Dim sourcenames, currentShadowCount, DSSources(30), numberofsources, numberofsources_hold
sourcenames = Array ("","","","","","","","","","","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(20), objrtx2(20)
dim objBallShadow(20)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9,BallShadowA10,BallShadowA11,BallShadowA12,BallShadowA13,BallShadowA14,BallShadowA15,BallShadowA16,BallShadowA17,BallShadowA18)

DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 0.11
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 0.12
    objrtx2(iii).visible = 0

    currentShadowCount(iii) = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = iii/1000 + 0.14
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

prightballrefl.blenddisablelighting = 7
pleftballrefl.blenddisablelighting = 7

dim rightballrefl, leftballrefl,rightBallReflStatus,leftBallReflStatus
rightBallReflStatus = 0 : leftBallReflStatus = 0




Sub DynamicBSUpdate

  if WizardPhase = 4 Then exit sub    'exit if galdor battle

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

  'ball brightness
  if ballbrightness <> -1 then
    For s=0 to UBound(BOT)
        BOT(s).color = ballbrightness + (ballbrightness * 256) + (ballbrightness * 256 * 256)
        if s = UBound(BOT) then 'until last ball brightness is set, then reset to -1
          if ballbrightness = ballbrightnessMax Or ballbrightness = ballbrightnessMin then ballbrightness = -1
        end if
    Next
  end if

'The Magic happens now
  For s = lob to UBound(BOT)

    '###### underpf detection. If ball on underpf and mission 5 -> gioff and gioff_underpf_boost
    if BOT(s).z < 0 and gioff_underpf_boost_status = 0 then
      'debug.print "ball under pf"
      if CurrentMission = 5 Or bWizardMode Then
        'debug.print "mission 5 is ongoing"
        if inRect(BOT(s).x, BOT(s).y, 76,1050,  788, 1050, 679, 1827, 157, 1824) then
          'debug.print "ball in perimeter"
          'setting gioff and under glow on. Mission 5 ending should bug gi back on.
          gioff
          gioff_underpf_boost true
        end if
      end if
    end if

    'Draw shadow to max of 6 balls
    If s > 6 Then Exit For    'No balls in play, exit

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

' *** Sideblade reflections
    if rightBallReflStatus = 0 then
      if InRect(BOT(s).x, BOT(s).y, 775,0,  952, 0, 952, 2100, 775, 2100) And BOT(s).z > 0 then
        set rightballrefl = BOT(s)
        rightBallReflStatus = 1
        pRightBallRefl.visible=true
      end if
    else
      if InRect(rightballrefl.x, rightballrefl.y, 775,0,  952, 0, 952, 2100, 775, 2100) And rightballrefl.z > 0 then
        pRightBallRefl.x = (tablewidth - rightballrefl.x) + tablewidth
        pRightBallRefl.y = rightballrefl.y
        pRightBallRefl.z = rightballrefl.z
      Else
        'debug.print "Ball released"
        rightballrefl = ""
        rightBallReflStatus = 0
        pRightBallRefl.visible=false
      end if
    end if

    if leftBallReflStatus = 0 Then
      if InRect(BOT(s).x, BOT(s).y, 0,0,  175, 0, 175, 2100, 0, 2100) And BOT(s).z > 0 then
        set leftballrefl = BOT(s)
        leftBallReflStatus = 1
        pLeftBallRefl.visible=true
      end If
    Else
      if InRect(leftballrefl.x, leftballrefl.y, 0,0,  175, 0, 175, 2100, 0, 2100) And leftballrefl.z > 0 then
        pLeftBallRefl.x = -leftballrefl.x
        pLeftBallRefl.y = leftballrefl.y
        pLeftBallRefl.z = leftballrefl.z
      Else
        'debug.print "Ball released"
        leftballrefl = ""
        leftBallReflStatus = 0
        pLeftBallRefl.visible=false
      end if
    end if
  Next
End Sub
'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************

'************************************************************
'   TILT
'************************************************************


'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                    'Called when table is nudged
  if BallsOnPlayfield <> 0 Then         'disable tilting when ball has drained
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity) AND(Tilt <15) Then 'show a warning
      'DMD Warning
      TiltWarning = 1
      TiltWarningTimer.enabled = True
      AudioCallout "tilt warning"
    PuPEvent(514)
    DOF 514, DOFPulse
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
      Tilted = True
      DisableTable True
      tilttableclear.enabled = true
      TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
      gioff
      hatch false
      bumper1.threshold = 666     'disable bumpers
      bumper2.threshold = 666
      bumper3.threshold = 666
      AudioCallout "tilt"
    PuPEvent(515)
    DOF 515, DOFPulse
    End If
  end if
End Sub

sub TiltWarningTimer_Timer
  TiltWarning = 0
  FlexDMD.Stage.GetLabel("Splash").Text = ""  'to hide TILT text after tilt recovery
  TiltWarningTimer.enabled = False
end sub

Dim tilttime:tilttime = 0

sub tilttableclear_timer
  tilttime = tilttime + 1
  Select Case tilttime
    Case 10
      tableclearing
  End Select
End Sub

Sub tableclearing

End Sub

Sub posttiltreset

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

    LeftFlipper.RotateToStart
    Flipper1.RotateToStart
    RightFlipper.RotateToStart
    LeftSlingshot.Disabled = 1
    RightSlingshot.Disabled = 1

  Else
    LeftSlingshot.Disabled = 0
    RightSlingshot.Disabled = 0
  End If
End Sub

Sub TiltRecoveryTimer_Timer()
  If(BallsOnPlayfield = 0) Then
    if not Slammed then EndOfBall()
    TiltRecoveryTimer.Enabled = False
    FlexDMD.Stage.GetLabel("Splash").Text = ""  'to hide TILT text after tilt recovery
    bumper1.threshold = 1.6     'enable bumpers
    bumper2.threshold = 1.6
    bumper3.threshold = 1.6
  End If
End Sub





'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 16 ' total number of balls
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

dim tracyMotorState : tracyMotorState = 0

Sub Rollingupdate()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
'   StopSound("BallRoll_" & b & "_amp9")
    'StopSound("fx_ballrolling" & b)
    StopSound(Cartridge_Ball_Roll & "_Ball_Roll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  If TracySpins.Enabled = True Or TracyTimer.Enabled = True Or TracyAttention.Enabled = True Then
    if tracyDir <> 0 then
      if tracyDir = -1 And tracyMotorState = 0 Then
        PlaySound "FX_tracy_motor", 0, 0.1, AudioPan(Primitive021), 0, 6000, 0, 0, AudioFade(Primitive021)
        tracyMotorState = -1
'       debug.print "minus"
      elseif tracyDir = 1 And tracyMotorState = 0 then
        PlaySound "FX_tracy_motor", 0, 0.1, AudioPan(Primitive021), 0, 0, 0, 0, AudioFade(Primitive021)
        tracyMotorState = 1
'       debug.print "pos"
      Elseif tracyMotorState > tracyDir then
'       debug.print "dir change"
        StopSound "FX_tracy_motor"
        PlaySound "fx_Tracy_Motor_Stop", 0, 0.2, AudioPan(Primitive021), 0, 0, 0, 0, AudioFade(Primitive021)
        tracyMotorState = 0
      Elseif tracyMotorState < tracyDir then
'       debug.print "dir change"
        StopSound "FX_tracy_motor"
        PlaySound "fx_Tracy_Motor_Stop", 0, 0.2, AudioPan(Primitive021), 6000, 0, 0, 0, AudioFade(Primitive021)
        tracyMotorState = 0
      end if
    Else
'     debug.print "tracydir 0"
      StopSound "FX_tracy_motor"
      'tracyMotorState = 0
    end if
  Else
    if tracyMotorState <> 0 Then
'     debug.print "timers ended, stopping"
      StopSound "FX_tracy_motor"
      tracyMotorState = 0
    end if
  end if



  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)
    if b > 6 then exit Sub 'Don't play rolling sound if more that 6 balls in play

    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
'     PlaySound ("BallRoll_" & b & "_amp9"), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
      PlaySound (Cartridge_Ball_Roll & "_Ball_Roll_" & b), -1, VolPlayfieldRoll(BOT(b)) * PlayfieldRollVolumeDial * VolumeDial * 1.25, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
      'PlaySound ("fx_ballrolling" & b), -1, VolPlayfieldRoll(BOT(b)) * 0.4 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
'       StopSound("BallRoll_" & b & "_amp9")
        StopSound(Cartridge_Ball_Roll & "_Ball_Roll_" & b)
        'StopSound("fx_ballrolling" & b)
        rolling(b) = False
      End If
    End If

    '***Ball Drop Sounds***
    If BOT(b).VelZ < -1 and ((BOT(b).z < 55 and BOT(b).z > 27) or (BOT(b).z < 55-100 and BOT(b).z > 27-100)) Then 'height adjust for ball drop sounds
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

'   'return captive ball back if it has dropped out. This could be in some much slower timer.
'   if IsObject(CaptiveBallID) then
'     if Not IsEmpty(CaptiveBallID) then
'       if Not InRect(CaptiveBallID.x,CaptiveBallID.y,  176,200,  317,200,  360,390,  190,400) then
'         'debug.print "***** captive ball lost! Returning it back"
'         CaptiveBallID.x = 282
'         CaptiveBallID.y = 351
'         CaptiveBallID.z = 41
'       end if
'     end if
'   end if
  Next
End Sub




'*****************************************************************************************************************************************
'   START GAME, END GANE
'*****************************************************************************************************************************************


'******** Skill Shots *************


Sub CheckSkillshotReady
  if bPlungerPulled And bBallInPlungerLane And LFPress=1 Then
    'debug.print "leftflip event while pulled"
    If skillshot=1 or Skillshot=2 Then
      SKBackup = Skillshot
      'debug.print "sk3"
      SkillshotOff
      Skillshot=3 ' activate skillshot 3

      LS_JP3all.Play SeqMiddleOutVertOn, 80, 100000
    End If
  end if
End Sub

Dim Skillshot
Sub SkillshotStart
  if bMultiBallMode then exit Sub
  Gate004.collidable=True
  cLightTopLeftLane.state=0:cLightTopRightLane.state=0
  if SKBackup = 0 then
    Skillshot=int(rnd(1)*2)+1
    SKBackup = Skillshot
  Else
    Skillshot = SKBackup
    'debug.print "setting sk from backup; " & Skillshot
  end if
  If Skillshot=1 Then cLightTopLeftLane.timerenabled=True Else cLightTopRightLane.timerenabled=True
End Sub

Sub Skillshot2_Timer : Skillshotoff : Skillshot2.enabled=False : End Sub


Sub cLightTopLeftLane_Timer : If me.state=1 Then : me.state=0 : Else : me.state=1 : End If : End Sub
Sub cLightTopRightLane_Timer : If me.state=1 Then : me.state=0 : Else : me.state=1 : End If : End Sub


Dim ShowSkillshot,ShowSkillshot2
Sub AwardSkillshot
    If Tilted Then Exit Sub
  if bMultiBallMode then exit Sub
  dim xx

  DOF 129,2 'skill shot 1
  CompletedSkillshots(CurrentPlayer)=CompletedSkillshots(CurrentPlayer)+1
  AudioCallout "skillshot"
  PuPEvent(505)
  DOF 505, DOFPulse

  Select Case CompletedSkillshots(CurrentPlayer)
    case 1 :
      if cLightCoreyMB.state <> 1 then
        for xx = 0 to 4
          if CoreyLights(xx).state = 0 then CoreyLights(xx).state = 2 : exit for
        Next
      end If
      Addscore score_Skillshot1_1st : addBonus bonus_Skillshot1_1st

      ShowSkillshot2="500k"
      ShowSkillshot=100
      SkillPupFlash= "1"
      SkillShotPup.Enabled=True

    case 2 :
      If MimaCount<3 and cLightMimaMB.state <> 1 Then
        WriteToLog "AwardSkillshot", "Mima Lock Ready"
        MimaCount = 3
        cTargetMima1.state = 1
        cTargetMima2.state = 1
        cTargetMima3.state = 1
        Rise_RampMima.enabled=True
        ShowSkillshot=100
        ShowSkillshot2="MIMA LOCK READY"
        WriteToLog "AwardSkillshot","MimaCount=" & MimaCount
      Else
        ShowSkillshot2="600k"
        ShowSkillshot=100
      End If
      Addscore score_Skillshot1_2nd : addBonus bonus_Skillshot1_2nd
      SkillPupFlash= "2"
      SkillShotPup2.Enabled=True
    case 3 :
      if cLightCoreyMB.state <> 1 then
        for xx = 0 to 4
          if CoreyLights(xx).state = 0 then CoreyLights(xx).state = 2 : exit for
        Next
      end If
      Addscore score_Skillshot1_3rd : addBonus bonus_Skillshot1_3rd
      ShowSkillshot2="700k"
      ShowSkillshot=100
      SkillPupFlash= "3"
      SkillShotPup3.Enabled=True
    case 4,5,6 :
      if cLightCoreyMB.state <> 1 then
        for xx = 0 to 4
          if CoreyLights(xx).state = 0 then CoreyLights(xx).state = 2 : exit for
        Next
      end If
      Addscore score_Skillshot1_4th : addBonus bonus_Skillshot1_4th
      ShowSkillshot2="800k"
      ShowSkillshot=100
      SkillPupFlash= "4"
      SkillShotPup4.Enabled=True
  End Select


End Sub


Sub AwardSkillshot2
    If Tilted Then Exit Sub
  Dim xx

  DOF 130,2 'skill shot 2
  playsound "zing"
  AudioCallout "skillshot"
  PuPEvent(472)
  DOF 472,DOFPulse

  CompletedSkillshots(CurrentPlayer)=CompletedSkillshots(CurrentPlayer)+1

  Select Case CompletedSkillshots(CurrentPlayer)
    case 1 :
      if cLightCoreyMB.state <> 1 then
        for xx = 0 to 4
          if CoreyLights(xx).state = 0 then CoreyLights(xx).state = 2 : exit for
        Next
      end If
      Addscore score_Skillshot2_1st : addBonus bonus_Skillshot2_1st
      ShowSkillshot2="2 MILLION"
      ShowSkillshot=100
      SkillPupFlash= "5"
      SkillShotPup5.Enabled=True
      AddMultiball 1
'   case 2 :
'     If Not cLightCollectExtraBall.state=2 and TotalExtraBallsAwarded(CurrentPlayer) < MaxExtraBallsPerGame Then
'       WriteToLog "AwardSkillshot2", "Extra ball is lit"
'       cLightCollectExtraBall.state=2
'       AudioCallout "extra ball lit"
'       ShowSkillshot=100
'       ShowSkillshot2="EXTRABALL IS LIT"
'     Else
'       ShowSkillshot2="2.2 MILLION"
'       ShowSkillshot=100
'     End If
'     Addscore score_Skillshot2_2nd : addBonus bonus_Skillshot2_2nd
    case 2 :
      Addscore score_Skillshot2_2nd : addBonus bonus_Skillshot2_2nd
      ShowSkillshot2="2.2 MILLION"
      ShowSkillshot=100
      SkillPupFlash= "6"
      SkillShotPup6.Enabled=True
      AddMultiball 1
    case 3 :
      Addscore score_Skillshot2_3rd : addBonus bonus_Skillshot2_3rd
      ShowSkillshot2="2.4 MILLION"
      ShowSkillshot=100
      SkillPupFlash= "7"
      SkillShotPup7.Enabled=True
      AddMultiball 1
    case 4,5,6 :
      Addscore score_Skillshot2_4th : addBonus bonus_Skillshot2_4th
      ShowSkillshot2="2.6 MILLION"
      ShowSkillshot=100
      SkillPupFlash= "8"
      SkillShotPup8.Enabled=True
      AddMultiball 1
  End Select


End Sub



Sub AwardSkillshot3 ' secret skillshot doubles ss2 points
    If Tilted Then Exit Sub
  Dim xx

  DOF 131,2 'skill shot 3
  playsound "zing"
  AudioCallout "skillshot"
  PuPEvent(472)
  DOF 472, DOFPulse

  CompletedSkillshots(CurrentPlayer)=CompletedSkillshots(CurrentPlayer)+1

  Select Case CompletedSkillshots(CurrentPlayer)
    case 1 :
      If Not cLightCollectExtraBall.state=2 and TotalExtraBallsAwarded(CurrentPlayer) < MaxExtraBallsPerGame Then
        WriteToLog "AwardSkillshot3", "Extra ball is lit"
        AudioCallout "extra ball lit"
        cLightCollectExtraBall.state=2
        ShowSkillshot=100
        ShowSkillshot2="EXTRABALL IS LIT"
        SkillPupFlash= "9"
        SkillShotPup9.Enabled=True
      Else
        ShowSkillshot2="4 MILLION"
        ShowSkillshot=100
      End If
      Addscore score_Skillshot3_1st : addBonus bonus_Skillshot3_1st

    case 2 :
      for xx = 0 to 4
        if CoreyLights(xx).state = 0 then CoreyLights(xx).state = 2 : exit for
      Next
      Addscore score_Skillshot3_2nd : addBonus bonus_Skillshot3_2nd
      ShowSkillshot2="4.4 MILLION"
      ShowSkillshot=100
      SkillPupFlash= "10"
      SkillShotPup10.Enabled=True
      AddMultiball 1
    case 3 :
      Addscore score_Skillshot3_3rd : addBonus bonus_Skillshot3_3rd
      ShowSkillshot2="4.8 MILLION"
      ShowSkillshot=100
      SkillPupFlash= "11"
      SkillShotPup11.Enabled=True
      AddMultiball 1
    case 4,5,6 :
      Addscore score_Skillshot3_4th : addBonus bonus_Skillshot3_4th
      ShowSkillshot2="5.2 MILLION"
      ShowSkillshot=100
      SkillPupFlash= "12"
      SkillShotPup12.Enabled=True
      AddMultiball 1
  End Select
End Sub




Sub Skillshotoff
  If CurrentMission<>6 Then Gate004.collidable=False
  Skillshot=0
  cLightTopLeftLane.timerenabled=False
  cLightTopRightLane.timerenabled=False
  cLightTopLeftLane.state=0
  cLightTopRightLane.state=0
  LS_JP3all.StopPlay
  LS_Skillshot2.StopPlay
  Skillshot2.Enabled=False
End Sub



Sub Gate002_Hit
  If Tilted Then Exit Sub
  If skillshot=3 Then Skillshot2.Enabled=True
End Sub



'********* End of Ball Multiplier *********

Sub swTopRightLane_unhit
  If Tilted Then Exit Sub
' playsound("EFX_effect3")
  AddScore score_TopLanes
  If skillshot=2 Then AwardSkillshot
  If skillshot>0 Then Skillshotoff
  cLightTopRightLane.state=1
  CheckEOBmulti
  SwitchWasHit("swTopRightLane")
End Sub

Sub swTopLeftLane_unhit
  If Tilted Then Exit Sub
' playsound("EFX_effect3")
  AddScore score_TopLanes
  If skillshot=1 Then AwardSkillshot
  If skillshot>0 Then Skillshotoff
  cLightTopLeftLane.state=1
  CheckEOBmulti
  SwitchWasHit("swTopLeftLane")
End Sub


Sub CheckEOBmulti
  If cLightTopLeftLane.state=1 And cLightTopRightLane.state=1 Then

    LS_Bonuslights.Play SeqBlinking,,10,20

    AddScore score_IncreasedEOBmulti
    cLightTopLeftLane.state=0
    cLightTopRightLane.state=0
    ' eob MP = up 1 level
'   If BonusMultiplier(CurrentPlayer)=1 Then
'     BonusMultiplier(CurrentPlayer)=2
'     AudioCallout "2x bonus"
'   Else
      BonusMultiplier(CurrentPlayer)=BonusMultiplier(CurrentPlayer)+1
      select case BonusMultiplier(CurrentPlayer)
        case 2 : AudioCallout "2x bonus":EOBGaugeTargetAngle = -10
          showMissionVideo=51
          If Mission6BumperTimer.Enabled = True Then
          PuPEvent(637)
          End If
          If Mission6BumperTimer.Enabled = False Then
          PuPEvent(532)
          DOF 532, DOFPulse
          End If
        case 3 : AudioCallout "3x bonus":EOBGaugeTargetAngle = -40
          showMissionVideo=52
          If Mission6BumperTimer.Enabled = True Then
          PuPEvent(637)
          End If
          If Mission6BumperTimer.Enabled = False Then
          PuPEvent(533)
          DOF 533, DOFPulse
          End If
        case 4 : AudioCallout "4x bonus":EOBGaugeTargetAngle = -72
          showMissionVideo=53
          If Mission6BumperTimer.Enabled = True Then
          PuPEvent(637)
          End If
          If Mission6BumperTimer.Enabled = False Then
          PuPEvent(534)
          DOF 534, DOFPulse
          End If
        case 5 : AudioCallout "5x bonus":EOBGaugeTargetAngle = -105
          showMissionVideo=54
          If Mission6BumperTimer.Enabled = True Then
          PuPEvent(637)
          End If
          If Mission6BumperTimer.Enabled = False Then
          PuPEvent(535)
          DOF 535, DOFPulse
          End If
        case 6 : AudioCallout "6x bonus":EOBGaugeTargetAngle = -135
          showMissionVideo=55
          If Mission6BumperTimer.Enabled = True Then
          PuPEvent(637)
          End If
          If Mission6BumperTimer.Enabled = False Then
          PuPEvent(536)
          DOF 536, DOFPulse
          End If
        case 7 : AudioCallout "7x bonus":EOBGaugeTargetAngle = -167
          showMissionVideo=56
          If Mission6BumperTimer.Enabled = True Then
          PuPEvent(637)
          End If
          If Mission6BumperTimer.Enabled = False Then
          PuPEvent(537)
          DOF 537, DOFPulse
          End If
        case 8 : AudioCallout "8x bonus":EOBGaugeTargetAngle = -199
          showMissionVideo=57
          If Mission6BumperTimer.Enabled = True Then
          PuPEvent(637)
          End If
          If Mission6BumperTimer.Enabled = False Then
          PuPEvent(538)
          DOF 538, DOFPulse
          End If
        case 9 : AudioCallout "9x bonus":EOBGaugeTargetAngle = -230
          If Mission6BumperTimer.Enabled = True Then
          PuPEvent(637)
          End If
          If Mission6BumperTimer.Enabled = False Then
          PuPEvent(539)
          DOF 539, DOFPulse
          End If
        case 10 : AudioCallout "10x bonus":EOBGaugeTargetAngle = -260
          If Mission6BumperTimer.Enabled = True Then
          PuPEvent(637)
          End If
          If Mission6BumperTimer.Enabled = False Then
          PuPEvent(540)
          DOF 540,DOFPulse
          End If
      end select

      If BonusMultiplier(CurrentPlayer)>10 Then BonusMultiplier(CurrentPlayer)=10 : Addscore score_MaxedEOBmulti


    End If

    'ResetEOBmultilights
'
' End If
End Sub







Sub ResetForNewGame()

  WriteTOLog "-------------","NEW GAME"
  SuperJets=0

  Dim i

  bGameInPLay = True
  bExtraBallWonThisBall = False
  bMultiballReady = False
  bMultiBallMode = False
  PupBumperTracy = True
  PupBumpertenthousand =False
  PupBumpersuperjets = False

  StopAttractMode
  TotalGamesPlayed = TotalGamesPlayed + 1
  savegp
  CurrentPlayer = 1
  PlayersPlayingGame = 1
  GameplayerNotup = "1"
  bOnTheFirstBall = True
  pDMDStartGame
  For i = 1 To MaxPlayers
    Score(i) = 0
    BonusPoints(i) = 0
    BonusHeldPoints(i) = 0
    BonusMultiplier(i) = 1
    BallsRemaining(i) = 3
    Ballblink=20
    ExtraBallsAwards(i) = 0
    TotalExtraBallsAwarded(i) = 0
    BumperHits(i) = 0
    BumperHitsPup(i) = 0
    BumperThousandPup(i) = 0
    BumperTwentyThousandPup(i) = 0
    BumperscoringPup(i) =0
    MysteryEB(i)=0
    Mission1Points(i)=0
    Mission2Points(i)=0
    Mission3Points(i)=0
    Mission4Points(i)=0
    Mission5Points(i)=0
    Mission6Points(i)=0
    Mission7Points(i)=0
    CompletedSkillshots(i)=0
    EOB_Ramps = 0
    EOB_Missions = 0
    EOB_MimaLoops = 0
  Next

  'ResetEOBmultilights
  EOBGaugeTargetAngle = 25

  PlayerLights_ResetInserts ' turn them all off
  PlayerLights_ResetAll
  PlayerLights_NewPlayer

  Tilt = 0
  Game_Init()
  if Not Slammed then
    vpmtimer.addtimer 1500, "FirstBall '"
    Slammed = False
  end if

  Scorbit.StartSession()

  'audios
  PlayDefaultSong
  AudioCallout "boot"
End Sub



Sub EndOfGame()
  StopVideosPlaying
  showMissionVideo=60
  PuPEvent(409)
  Scorbit.StopSession Score(1), Score(2), Score(3), Score(4), PlayersPlayingGame
  AttractScoreNumbers
  'Playsound "Game_Over_1_Yell"
  vpmtimer.addtimer 4500, "bGameInPLay = False : bJustStarted = False : StartAttractMode : StopVideosPlaying :pDMDGameOver  '"
End Sub


'*****************************************************************************************************************************************
'   ATTRACT MODE
'*****************************************************************************************************************************************

dim tAttractTime : tAttractTime = 0

Sub StartAttractMode()
  'PlaySong 3
  'PlaySongLoop 9,39.82,47.72
  bAttractMode = True
  DMD_Stopoverlays = False
  bAttractModeCounter = 0


  if AmbientAudio = 2 Then
    StopSound "nebuleuse"
  end if

  GI_Color "white",1
  gion
  StartAttractLightSeq

End Sub


Sub StopAttractMode()


  'these will not remove the backfround from dmd
  StopVideosPlaying

  DMD_ShowText_Reset
  DMD_ShowImages_Reset

  'these will not remove the backfround from dmd
  DMD_Stopoverlays = False ' turn off attracktoverlay ( and all the others too )

  FlexDMD.lockRenderThread
  FlexDMD.Stage.GetImage("attract1").Visible=False
  FlexDMD.UnlockRenderThread

  bAttractMode = False
  tAttractTime = 0

  FlexOverlay=0
  Flex_Overlay_CurrentImage=0

  gioff
  LS_Attract.StopPlay

End Sub




'*****************************************************************************************************************************************
'   LIGHT SEQUENCES
'*****************************************************************************************************************************************
dim GI_flickerCounter : GI_flickerCounter = 0
dim GI_flickerState : GI_flickerState = 1
dim GI_sounds : GI_sounds = true

sub GI_Flicker_timer
  if Not GI_Flicker.enabled then
    GI_sounds = false
    LS_Attract.StopPlay
    Dim xx
    For Each xx in AllLights
      xx.state=0
    Next
    GI_Flicker.enabled = true
  end If

  GI_flickerCounter = GI_flickerCounter + 1

  if GI_flickerCounter < 6 Then
    GI_Flicker.interval = 280 + RndInt(0,2)
  elseif GI_flickerCounter > 5 And GI_flickerCounter < 13 Then
    GI_Flicker.interval = 120 + RndInt(0,2)
  elseif GI_flickerCounter > 12 And GI_flickerCounter < 20 Then
    GI_Flicker.interval = 60 + RndInt(0,2)
  elseif GI_flickerCounter > 19 And GI_flickerCounter < 44 Then
    GI_Flicker.interval = 30 + RndInt(0,2)
  end If

  if GI_flickerState = 1 Then
    gioff
    GI_flickerState = 0
  Else
    gion
    GI_flickerState = 1
  end if

  if GI_flickerCounter > 44 then
    if not bGameInPLay then
      gion : StartAttractLightSeq
    else
      LS_Attract.StopPlay
    end if

    GI_flickerCounter = 0
    GI_flickerState = 1
    GI_sounds = true
    GI_Flicker.enabled = False
    GI_Flicker.interval = 280
  end if

end sub

'vpmtimer.addtimer 3000, "GI_Flicker_timer '"

'set all lights on.. Just for giggles..
Dim a
For each a in AllLights
  a.State = 1
Next

Sub StartAttractLightSeq()

  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 50, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqCircleOutOn, 15, 2
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 25, 1
  LS_Attract.UpdateInterval = 10
  LS_Attract.Play SeqCircleOutOn, 15, 3
  LS_Attract.UpdateInterval = 5
  LS_Attract.Play SeqRightOn, 50, 1
  LS_Attract.UpdateInterval = 5
  LS_Attract.Play SeqLeftOn, 50, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqRightOn, 50, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 50, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqRightOn, 40, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 40, 1
  LS_Attract.UpdateInterval = 10
  LS_Attract.Play SeqRightOn, 30, 1
  LS_Attract.UpdateInterval = 10
  LS_Attract.Play SeqLeftOn, 30, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqRightOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqRightOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 15, 1
  LS_Attract.UpdateInterval = 10
  LS_Attract.Play SeqCircleOutOn, 15, 3
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqRightOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 25, 1
  LS_Attract.UpdateInterval = 5
  LS_Attract.Play SeqStripe1VertOn, 50, 2
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqCircleOutOn, 15, 2
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqStripe1VertOn, 50, 3
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqRightOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqCircleOutOn, 15, 2
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqStripe2VertOn, 50, 3
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqRightOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 25, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqStripe1VertOn, 25, 3
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqStripe2VertOn, 25, 3
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqUpOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqDownOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqRightOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqRightOn, 15, 1
  LS_Attract.UpdateInterval = 8
  LS_Attract.Play SeqLeftOn, 15, 1
End Sub

Sub LS_Attract_PlayDone()
  StartAttractLightSeq()
End Sub



'*****************************************************************************************************************************************
'   Songs
'*****************************************************************************************************************************************

Dim oPlayer1, oPlayer2, fso, AmbientAudio

AmbientAudio = 0

Set oPlayer1 = CreateObject("WMPlayer.OCX")
Set oPlayer2 = CreateObject("WMPlayer.OCX")
Set fso = CreateObject("Scripting.FileSystemObject")

initPlayer

sub initPlayer
  if Not fso.FolderExists(MusicDirectory) Then
    'InfoBox.text = "Music path: " & MusicDirectory & " not found. Checking default path."
    MusicDirectory = "C:\vPinball\VisualPinball\Music\Blood Machines OST\"
    if Not fso.FolderExists(MusicDirectory) Then
      MusicDirectory = "D:\Visual Pinball\Music\Blood Machines OST\"
      if Not fso.FolderExists(MusicDirectory) Then
        'InfoBox.text = "Music path: " & MusicDirectory & " not found. Go get a Music pack."
        'msgbox "No music detected, will play ambient audio"
        AmbientAudio = 1
      else
        'InfoBox.text = ""
      end If
    else
      'InfoBox.text = ""
    end If
  end if
end sub

dim SongLoopStart : SongLoopStart = 0
dim SongLoopEnd : SongLoopEnd = 0
dim CurrentSong : CurrentSong = 0

sub PlayDefaultSong
  if Not bMultiBallMode then  'If MB ongoing, it won't play default song
    Select Case RndInt(1,3)
      Case 1 : PlaySong 3
      Case 2 : PlaySong 4
      Case 3 : PlaySong 11
    end select
  end if
  if AmbientAudio = 1 Then
    'PlaySound playsoundparams, -1, aVol * VolumeDial
    PlaySound "nebuleuse",-1,AmbientAudioVolume
    AmbientAudio = 2
  end if
end sub


sub PlaySong(aNro)
  StopSong
  CurrentSong = aNro
  oPlayer1.URL = MusicDirectory & Songs(CurrentSong)
  oPlayer1.settings.volume = MusicVol*100
  oPlayer1.controls.play
end sub


'dim songNro
'sub PlaySong(aNro)
' songNro = aNro
' StopSong
' CurrentSong = songNro
' oPlayer1.settings.volume = MusicVol*100
' SongStartDelay.interval = 500
' SongStartDelay.enabled = true
'end sub
'
'sub SongStartDelay_timer
''  debug.print "delayed"
' oPlayer1.URL = MusicDirectory & "\" & Songs(CurrentSong)
' oPlayer1.controls.play
' SongStartDelay.enabled = false
'end sub


dim playloop:playloop = 1
sub PlaySongLoop(aNro,starting,ending)
  StopSong
  CurrentSong = aNro
  SongLoopStart = max(starting,0)
  SongLoopEnd = max(ending,starting)
  playloop = 1
  PlaySongStartingAt 1, CurrentSong, SongLoopStart
end sub


sub PlaySongPart(aNro,starting,ending)
  StopSong
  CurrentSong = aNro
  SongLoopStart = max(starting,0)
  SongLoopEnd = max(ending,starting)
  playloop = 0
  PlaySongStartingAt 1, CurrentSong, SongLoopStart
end sub


sub PlaySongStartingAt(chan,aNro,seconds)
  CurrentSong = aNro
  If chan = 1 Then
    oPlayer1.URL = MusicDirectory & Songs(CurrentSong)
    oPlayer1.controls.currentPosition = seconds
    oPlayer1.settings.volume = MusicVol*100
    oPlayer1.controls.play
  Elseif chan = 2 Then
    oPlayer2.URL = MusicDirectory & Songs(CurrentSong)
    oPlayer2.controls.currentPosition = seconds
    oPlayer2.settings.volume = MusicVol*100
    oPlayer2.controls.play
  End If
end sub


'sub StopSongOld
''  If oPlayer1.playState <> 1 Then oPlayer1.close
''  If oPlayer2.playState <> 1 Then oPlayer2.close
' SongLoopStart = 0
' SongLoopEnd = 0
' CurrentSong = 0
'end sub

sub StopSong
    If oPlayer1.playState <> 1 Then oPlayer1.controls.stop
    If oPlayer2.playState <> 1 Then oPlayer2.controls.stop
    SongLoopStart = 0
    SongLoopEnd = 0
    CurrentSong = 0
end sub

Function SongTime
  Dim t1, t2
  t1 = oPlayer1.controls.currentPosition
  t2 = oPlayer2.controls.currentPosition
  SongTime = max(t1,t2)
End Function


Const songdelta = 0.08

sub SongTimerTimer
  'Automatically play new song if the last one has ended
  If CurrentSong > 0 and oPlayer1.playState = 1 And Not bAttractMode Then PlayDefaultSong

  'InfoBox.text = oPlayer1.controls.currentPosition & " - " & oPlayer2.controls.currentPosition

  'Play a song loop if that has been enabled
  If SongLoopEnd > 0 and oPlayer1.controls.currentPosition > SongLoopEnd Then
    oPlayer1.close
    if playloop = 0 And Not bAttractMode Then
      PlayDefaultSong
    end if
    WriteToLog "Song Timer", "oPlayer1.close"
  ElseIf SongLoopEnd > 0 and oPlayer1.controls.currentPosition > SongLoopEnd-songdelta Then
    If (oPlayer2.playState=0 or oPlayer2.playState=1 or oPlayer2.playState=10) And playloop = 1 Then PlaySongStartingAt 2, CurrentSong, SongLoopStart  : WriteToLog "Song Timer", "oPlayer2 startat"
  End If

  If SongLoopEnd > 0 and oPlayer2.controls.currentPosition > SongLoopEnd Then
    oPlayer2.close
    WriteToLog "Song Timer", "oPlayer2.close"
  ElseIf SongLoopEnd > 0 and oPlayer2.controls.currentPosition > SongLoopEnd-songdelta Then
    If (oPlayer1.playState=0 or oPlayer1.playState=1 or oPlayer1.playState=10) And playloop = 1 Then PlaySongStartingAt 1, CurrentSong, SongLoopStart : WriteToLog "Song Timer", "oPlayer1 startat"
  End If

end sub

'******************************************************
'   Speech callout handling
'******************************************************

Sub AudioCallout(scenario)
  StopAllCallouts
  startB2S 7
  If VRroom > 0 Then
    VRBGflash7.enabled = true
  End If
  select case scenario
    case "boot" :
      select case Int(Rnd * 6) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_never_seen_wrecks"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_welcome_to_blood_machines"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_set_me_free"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_prepare_yourself"",0,CalloutVol,0,0,1,1,1 '"
        case 5 : vpmtimer.addtimer 300, "PlaySound ""SPC_try_not_to_be_too_rough_on_me"",0,CalloutVol,0,0,1,1,1 '"
        case 6 : vpmtimer.addtimer 300, "PlaySound ""SPC_do_you_have_what_it_takes_to_win"",0,CalloutVol,0,0,1,1,1 '"
        'case 7 : vpmtimer.addtimer 300, "PlaySound ""SPC_lets_see_if_you_can_pilot_me"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "tilt"
      select case Int(Rnd * 3) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_respect_ship"",0,CalloutVol,0,0,1,1,1 '"
        'case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_missed_that"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_tracy_barely_felt"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_dont_do_that_again"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "tilt warning"
      select case Int(Rnd * 6) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_vascan"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_vascan_game_playing"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_slow_down_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_bad_idea"",0,CalloutVol,0,0,1,1,1 '"
        case 5 : vpmtimer.addtimer 300, "PlaySound ""SPC_be_gentle_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 6 : vpmtimer.addtimer 300, "PlaySound ""SPC_stop"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "outlane drain"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_tracy_anything_broken"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_my_god"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "start mission 1" : vpmtimer.addtimer 300, "PlaySound ""SPC_bring_mima_quickly"",0,CalloutVol,0,0,1,1,1 '"

    case "start mission 2" : vpmtimer.addtimer 300, "PlaySound ""SPC_death_brutal_ceremony"",0,CalloutVol,0,0,1,1,1 '"
    case "end mission 2" : vpmtimer.addtimer 300, "PlaySound ""SPC_witnessed_miracle"",0,CalloutVol,0,0,1,1,1 '"
    case "LAGO miracle" : vpmtimer.addtimer 300, "PlaySound ""SPC_witnessed_miracle"",0,CalloutVol,0,0,1,1,1 '"

    case "2x bonus" : vpmtimer.addtimer 300, "PlaySound ""SPC_a_two_x_bonus"",0,CalloutVol,0,0,1,1,1 '"
    case "3x bonus" : vpmtimer.addtimer 300, "PlaySound ""SPC_a_three_x_bonus"",0,CalloutVol,0,0,1,1,1 '"
    case "4x bonus" : vpmtimer.addtimer 300, "PlaySound ""SPC_a_four_x_bonus"",0,CalloutVol,0,0,1,1,1 '"
    case "5x bonus" : vpmtimer.addtimer 300, "PlaySound ""SPC_a_five_x_bonus"",0,CalloutVol,0,0,1,1,1 '"
    case "6x bonus" : vpmtimer.addtimer 300, "PlaySound ""SPC_a_six_x_bonus"",0,CalloutVol,0,0,1,1,1 '"

    case "mima ready"
      select case Int(Rnd * 5) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_mima_multiball_ready"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_mima_multiball_ready_2"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_start_mima_multiball"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_start_mima_multiball_2"",0,CalloutVol,0,0,1,1,1 '"
        case 5 : vpmtimer.addtimer 300, "PlaySound ""SPC_initiate_mima_multiball"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "corey ready"
      select case Int(Rnd * 4) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_corey_multiball_ready"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_start_corey_multiball"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_start_corey_multiball_2"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_initiate_corey_multiball"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "tracy ready"
      select case Int(Rnd * 4) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_intiate_tracy_multiball"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_start_tracy_multiball"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_start_tracy_multiball_2"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_tracy_multiball_ready"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "mission completed"
      select case Int(Rnd * 3) + 1
        case 1 : vpmtimer.addtimer 2000, "PlaySound ""SPC_mission_complete"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 2000, "PlaySound ""SPC_mission_completed_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 2000, "PlaySound ""SPC_mission_successful"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "mission failed"
      select case Int(Rnd * 3) + 1
        case 1 : vpmtimer.addtimer 2000, "PlaySound ""SPC_mission_failed"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 2000, "PlaySound ""SPC_mission_failed_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 2000, "PlaySound ""SPC_mission_failure"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "nice shot"
      select case Int(Rnd * 11) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_nice_shot"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_nice_shot_2"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_nice_shot_sir"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_well_done_sir"",0,CalloutVol,0,0,1,1,1 '"
        case 5 : vpmtimer.addtimer 300, "PlaySound ""SPC_well_done_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 6 : vpmtimer.addtimer 300, "PlaySound ""SPC_great_shot"",0,CalloutVol,0,0,1,1,1 '"
        case 7 : vpmtimer.addtimer 300, "PlaySound ""SPC_great_shot_sir"",0,CalloutVol,0,0,1,1,1 '"
        case 8 : vpmtimer.addtimer 300, "PlaySound ""SPC_good_aim_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 9 : vpmtimer.addtimer 300, "PlaySound ""SPC_do_it_again_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 10 : vpmtimer.addtimer 300, "PlaySound ""SPC_do_that_again"",0,CalloutVol,0,0,1,1,1 '"
        case 11 : vpmtimer.addtimer 300, "PlaySound ""SPC_destroy_it"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "extra ball lit" :
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_extra_ball_is_lit"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_get_the_extra_ball"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "extra ball"
      select case Int(Rnd * 4) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_you_got_the_extra_ball"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_successful_extra_ball"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_extra_ball"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_extra_ball_captain"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "ball added"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_a_ball_has_been_added_sir"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_a_ball_has_been_added"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "mystery lit"
      select case Int(Rnd * 3) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_get_the_mystery"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_mystery_is_lit"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_mystery_is_lit_captain"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "mystery"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_find_your_mystery"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_life_is_a_mystery"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "double scoring lit"
      select case Int(Rnd * 3) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_double_scoring_is_lit"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_double_scoring_is_lit_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_get_double_scoring_started"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "double scoring"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_double_scoring_activated"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_double_scoring_is_activated"",0,CalloutVol,0,0,1,1,1 '"
      end select


    case "super spinners lit"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_get_super_spinners"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_super_spinners_is_lit"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "super spinners"
      select case Int(Rnd * 4) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_super_spinner_activated"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_super_spinners"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_super_spinners_activated"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_the_super_spinner_is_activated"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "super jets" : vpmtimer.addtimer 300, "PlaySound ""SPC_super_jets_started"",0,CalloutVol,0,0,1,1,1 '"

    case "jackpot"
      select case Int(Rnd * 3) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_jackpot"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_jackpot_2"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_jackpot_3"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "super jackpot"
      select case Int(Rnd * 3) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_super_jackpot"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_super_jackpot_2"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_super_jackpot_3"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "drain" :
      StopAllCallouts
      select case Int(Rnd * 6) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_be_careful"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_dont_do_that_again"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_try_not_to_miss"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_she_got_the_better_of_you"",0,CalloutVol,0,0,1,1,1 '"
        case 5 : vpmtimer.addtimer 300, "PlaySound ""SPC_is_that_the_best_youve_got"",0,CalloutVol,0,0,1,1,1 '"
        case 6 : vpmtimer.addtimer 300, "PlaySound ""SPC_oh_no_captain"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "final chapter" : vpmtimer.addtimer 300, "PlaySound ""SPC_welcome_to_the_final_chapter"",0,CalloutVol,0,0,1,1,1 '"

    case "obey" : vpmtimer.addtimer 300, "PlaySound ""SPC_obey"",0,CalloutVol,0,0,1,1,1 '"

    case "scream" : vpmtimer.addtimer 300, "PlaySound ""SPC_scream"",0,CalloutVol,0,0,1,1,1 '"

    case "final chance" : vpmtimer.addtimer 300, "PlaySound ""SPC_final_chance"",0,CalloutVol,0,0,1,1,1 '"

    case "yes"
      select case Int(Rnd * 4) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_oh_yes"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_yes"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_yes_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_yes_captain_2"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "no"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_no_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_no_captain_2"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "touch me"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_touch_me_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_touch_me_captain_2"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "high score"
      select case Int(Rnd * 3) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_you_got_a_high_score"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_very_good_score_captain"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_very_good_game_sir"",0,CalloutVol,0,0,1,1,1 '"
      end select

    case "player1"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_player1"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_player.1"",0,CalloutVol,0,0,1,1,1 '"
      end select
    case "player2"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_player2"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_player.2"",0,CalloutVol,0,0,1,1,1 '"
      end select
    case "player3"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_player3"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_player.3"",0,CalloutVol,0,0,1,1,1 '"
      end select
    case "player4"
      select case Int(Rnd * 2) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_player4"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_player.4"",0,CalloutVol,0,0,1,1,1 '"
      end select
    case "skillshot"
      select case Int(Rnd * 4) + 1
        case 1 : vpmtimer.addtimer 300, "PlaySound ""SPC_skills1"",0,CalloutVol,0,0,1,1,1 '"
        case 2 : vpmtimer.addtimer 300, "PlaySound ""SPC_skills2"",0,CalloutVol,0,0,1,1,1 '"
        case 3 : vpmtimer.addtimer 300, "PlaySound ""SPC_skills3"",0,CalloutVol,0,0,1,1,1 '"
        case 4 : vpmtimer.addtimer 300, "PlaySound ""SPC_skills4"",0,CalloutVol,0,0,1,1,1 '"
      end select

  end select
end sub



sub StopAllCallouts()
  StopSound "SPC_never_seen_wrecks"
  StopSound "SPC_respect_ship"
  StopSound "SPC_missed_that"
  StopSound "SPC_tracy_barely_felt"
  StopSound "SPC_vascan"
  StopSound "SPC_vascan_game_playing"
  StopSound "SPC_slow_down_captain"
  StopSound "SPC_bad_idea"
  StopSound "SPC_tracy_anything_broken"
  StopSound "SPC_bring_mima_quickly"
  StopSound "SPC_death_brutal_ceremony"
  'StopSound "SPC_witnessed_miracle"  'rude hack to let Lago witness the miracle while Astro is crushing Jackpots
  StopSound "SPC_a_five_x_bonus"
  StopSound "SPC_a_four_x_bonus"
  StopSound "SPC_a_six_x_bonus"
  StopSound "SPC_a_three_x_bonus"
  StopSound "SPC_a_two_x_bonus"
  StopSound "SPC_be_careful"
  StopSound "SPC_be_gentle_captain"
  StopSound "SPC_corey_multiball_ready"
  StopSound "SPC_destroy_it"
  StopSound "SPC_do_it_again_captain"
  StopSound "SPC_do_that_again"
  StopSound "SPC_do_you_have_what_it_takes_to_win"
  StopSound "SPC_dont_do_that_again"
  StopSound "SPC_dont_fuck_up_again"
  StopSound "SPC_double_scoring_activated"
  StopSound "SPC_double_scoring_is_activated"
  StopSound "SPC_double_scoring_is_lit"
  StopSound "SPC_double_scoring_is_lit_captain"
  StopSound "SPC_double_scoring_is_lit_captain_2"
  StopSound "SPC_extra_ball_is_lit"
  StopSound "SPC_find_your_mystery"
  StopSound "SPC_get_double_scoring"
  StopSound "SPC_get_double_scoring_started"
  StopSound "SPC_get_super_spinners"
  StopSound "SPC_get_the_extra_ball"
  StopSound "SPC_get_the_mystery"
  StopSound "SPC_good_aim_captain"
  StopSound "SPC_great_shot"
  StopSound "SPC_great_shot_sir"
  StopSound "SPC_initiate_corey_multiball"
  StopSound "SPC_initiate_mima_multiball"
  StopSound "SPC_intiate_tracy_multiball"
  StopSound "SPC_is_that_the_best_youve_got"
  StopSound "SPC_jackpot"
  StopSound "SPC_jackpot_2"
  StopSound "SPC_jackpot_3"
  StopSound "SPC_keep_your_fucking_hands_to_yourself"
  StopSound "SPC_lets_see_if_you_can_pilot_me"
  StopSound "SPC_life_is_a_mystery"
  StopSound "SPC_mima_multiball_ready"
  StopSound "SPC_mima_multiball_ready_2"
  StopSound "SPC_mission_complete"
  StopSound "SPC_mission_completed"
  StopSound "SPC_mission_completed_captain"
  StopSound "SPC_mission_failed"
  StopSound "SPC_mission_failed_captain"
  StopSound "SPC_mission_failure"
  StopSound "SPC_mission_successful"
  StopSound "SPC_mystery_is_lit"
  StopSound "SPC_mystery_is_lit_captain"
  StopSound "SPC_nice_shot"
  StopSound "SPC_nice_shot_2"
  StopSound "SPC_nice_shot_sir"
  StopSound "SPC_no_captain"
  StopSound "SPC_no_captain_2"
  StopSound "SPC_oh_no_captain"
  StopSound "SPC_oh_yes"
  StopSound "SPC_prepare_yourself"
  StopSound "SPC_set_me_free"
  StopSound "SPC_she_got_the_better_of_you"
  StopSound "SPC_start_corey_multiball"
  StopSound "SPC_start_corey_multiball_2"
  StopSound "SPC_start_mima_multiball"
  StopSound "SPC_start_mima_multiball_2"
  StopSound "SPC_start_tracy_multiball 2"
  StopSound "SPC_start_tracy_multiball"
  StopSound "SPC_stop"
  StopSound "SPC_super_jackpot"
  StopSound "SPC_super_jackpot_2"
  StopSound "SPC_super_jackpot_3"
  StopSound "SPC_super_jets_started"
  StopSound "SPC_super_spinner_activated"
  StopSound "SPC_super_spinners"
  StopSound "SPC_super_spinners_activated"
  StopSound "SPC_super_spinners_is_lit"
  StopSound "SPC_the_double_scoring_is_lit"
  StopSound "SPC_the_double_scoring_is_lit_captain"
  StopSound "SPC_the_super_spinner_is_activated"
  StopSound "SPC_touch_me_captain"
  StopSound "SPC_touch_me_captain_2"
  StopSound "SPC_tracy_multiball_ready"
  StopSound "SPC_try_not_to_be_too_rough_on_me"
  StopSound "SPC_try_not_to_miss"
  StopSound "SPC_very_good_game_sir"
  StopSound "SPC_very_good_score_captain"
  StopSound "SPC_welcome_to_blood_machines"
  StopSound "SPC_welcome_to_the_final_chapter"
  StopSound "SPC_well_done_captain"
  StopSound "SPC_well_done_sir"
  StopSound "SPC_what_the_fuck_was_that"
  StopSound "SPC_yes"
  StopSound "SPC_yes_captain"
  StopSound "SPC_yes_captain_2"
  StopSound "SPC_you_got_a_high_score"
  StopSound "SPC_you_got_the_extra_ball"
  StopSound "SPC_successful_extra_ball"
  StopSound "SPC_extra_ball"
  StopSound "SPC_extra_ball_captain"
end sub

'*****************************************************************************************************************************************
'  High Scores
'*****************************************************************************************************************************************

' load em up


Dim hschecker:hschecker = 0

Sub Loadhs

  If ResetHighscores Then
    HighScore(0) = 50000000
    HighScore(1) = 30000000
    HighScore(2) = 10000000

    HighScoreName(0) = "VPX"
    HighScoreName(1) = "VPX"
    HighScoreName(2) = "VPX"
    Savehs

  Else
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 50000000 End If

    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "VPX" End If

    x = LoadValue(TableName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 30000000 End If

    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "VPX" End If

    x = LoadValue(TableName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 10000000 End If

    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "VPX" End If

    x = LoadValue(TableName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 10000000 End If

    x = LoadValue(TableName, "Credits")
    If(x <> "") then Credits = CInt(x) Else Credits = 0 End If

    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If
  End If

' If hschecker = 0 Then
'   checkorder
' End If


End Sub

Dim hs3,hs2,hs1,hs0,hsn3,hsn2,hsn1,hsn0


'Sub checkorder
' hschecker = 1
' hs3 = HighScore(3)
' hs2 = HighScore(2)
' hs1 = HighScore(1)
' hs0 = HighScore(0)
' hsn3 = HighScoreName(3)
' hsn2 = HighScoreName(2)
' hsn1 = HighScoreName(1)
' hsn0 = HighScoreName(0)
' If hs3 > hs0 Then
'   HighScore(0) = hs3
'   HighScoreName(0) = hsn3
'   HighScore(1) = hs0
'   HighScoreName(1) = hsn0
'   HighScore(2) = hs1
'   HighScoreName(2) = hsn1
'   HighScore(3) = hs2
'   HighScoreName(3) = hsn2
'
' ElseIf hs3 > hs1 Then
'   HighScore(0) = hs0
'   HighScoreName(0) = hsn0
'   HighScore(1) = hs3
'   HighScoreName(1) = hsn3
'   HighScore(2) = hs1
'   HighScoreName(2) = hsn1
'   HighScore(3) = hs2
'   HighScoreName(3) = hsn2
' ElseIf hs3 > hs2 Then
'   HighScore(0) = hs0
'   HighScoreName(0) = hsn0
'   HighScore(1) = hs1
'   HighScoreName(1) = hsn1
'   HighScore(2) = hs3
'   HighScoreName(2) = hsn3
'   HighScore(3) = hs2
'   HighScoreName(3) = hsn2
' ElseIf hs3 < hs2 Then
'   HighScore(0) = hs0
'   HighScoreName(0) = hsn0
'   HighScore(1) = hs1
'   HighScoreName(1) = hsn1
'   HighScore(2) = hs2
'   HighScoreName(2) = hsn2
'   HighScore(3) = hs3
'   HighScoreName(3) = hsn3
' End If
'
' savehs
'End Sub


Sub Savehs
  SaveValue TableName, "HighScore1", HighScore(0)
  SaveValue TableName, "HighScore1Name", HighScoreName(0)
  SaveValue TableName, "HighScore2", HighScore(1)
  SaveValue TableName, "HighScore2Name", HighScoreName(1)
  SaveValue TableName, "HighScore3", HighScore(2)
  SaveValue TableName, "HighScore3Name", HighScoreName(2)
  SaveValue TableName, "HighScore4", HighScore(3)
  SaveValue TableName, "HighScore4Name", HighScoreName(3)
End Sub

Sub SaveCredits
  SaveValue TableName, "Credits", Credits
End Sub

Sub Savegp

  SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
  vpmtimer.addtimer 1000, "Loadhs'"
End Sub


Dim hsbModeActive:hsbModeActive = False

'*****************************************************************************************************************************************
'   SCORING & AWARD FUNCTIONS
'*****************************************************************************************************************************************

Sub AddScore(points)
  If Not Tilted Then
    Borderblink=5

    Score(CurrentPlayer) = Score(CurrentPlayer) + points*PointMultiplier
    Scorbit.SendUpdate Score(1), Score(2), Score(3), Score(4), Balls, CurrentPlayer, PlayersPlayingGame
  End if
End Sub


Sub AddBonus(points)
  If Not Tilted Then
    Borderblink=5

    BonusPoints(CurrentPlayer)  = BonusPoints(CurrentPlayer)  + points*PointMultiplier
  End if
End Sub

Sub AwardSpecial
  Credits = Credits + 1
  SaveCredits
  KnockerSolenoidDOF 154, DOFPulse 'ssftodo: check number
End Sub


Sub AwardExtraBall
  TotalExtraBallsAwarded(CurrentPlayer) = TotalExtraBallsAwarded(CurrentPlayer) + 1
  if TotalExtraBallsAwarded(CurrentPlayer) <= MaxExtraBallsPerGame Then
    'DMDExtraBall=1
    DMD_ShowImages "extraball",2,50,2000,0 ' eb fast blink
    DOF 132,2 'award extra ball
    DOF 504, DOFPulse
    PuPevent (504)
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    cLightShootAgain1.state=1
    cLightShootAgain2.state=1
  End If
End Sub




'*****************************************************************************************************************************************
'   BALL FUNCTIONS & DRAINS
'*****************************************************************************************************************************************


Function Balls
  Dim tmp
  tmp = bpgcurrent - BallsRemaining(CurrentPlayer) + 1
  If tmp> bpgcurrent Then
    Balls = bpgcurrent
  Else
    Balls = tmp
  End If
End Function


Sub FirstBall
  WriteToLog "FirstBall", ""
  ResetForNewPlayerBall()
  CreateNewBall()
  PupBumperTracy = True

End Sub

Sub ResetForNewPlayerBall()
  WriteToLog "ResetForNewPlayerBall", ""
' cLightScavengeScoop.state=2           ' free mystery every ball ? =??
  SuperJets=0                 ' reset superjets for all balls
  cLightKB2.state=1 : cLightKB1.state=0     ' reset kickback
  ShowBallCount true
  PlayersPUP
  Superjetsdrainreset
  If PlayersPlayingGame > 1 Then
    FlexDMD.stage.GetLabel("Score_1").text = "Player " & CurrentPlayer
    AudioCallout "player" & CurrentPlayer
    PlayerUpFull
'   If CurrentPlayer = 1 Then
'     AudioCallout "player1"
'   Elseif currentplayer = 2 Then
'     AudioCallout "player2"
'   Elseif currentplayer = 3 Then
'     AudioCallout "player3"
'   Elseif currentplayer = 4 Then
'     AudioCallout "player4"
'   End If
  Else

  End If

  AddScore 0

  'do bonus points additions here
  Score(CurrentPlayer) = Score(CurrentPlayer) + BonusPoints(CurrentPlayer)*BonusMultiplier(CurrentPlayer)

  'then reset bonus
  BonusPoints(CurrentPlayer) = 0
  BonusMultiplier(CurrentPlayer) = 1
  bBonusHeld = False
  'ResetEOBmultilights
  EOBGaugeTargetAngle = 25
  AddBallAvailable = 1      'only one ball can be added via ShipCannonKicker

  bFlippersEnabled = True
  bExtraBallWonThisBall = False
  ResetNewBallLights()
  ResetNewBallVariables
  bBallSaverReady = True
  SkillshotStart  '**'
  MimaMagnetOn False
  solBloodTrail false

  GunTargetScorePerPlayer = GunTargetScore(CurrentPlayer)

  ScorbitBuildGameModes()

  gion
End Sub


Sub CreateNewBall()
  BallRelease.CreateSizedball BallSize / 2
  BallsOnPlayfield = BallsOnPlayfield + 1
  WriteToLog "CreateNewBall", "BallsOnPlayfield = "  & BallsOnPlayfield
' RandomSoundTroughKickoutDOF BallRelease, 104, DOFPulse  'ssftodo: check dof --> sling right
  RandomSoundShooterFeeder()
  DOF 145,2 'ball released
  bOpenTopGate = false
  BallRelease.Kick 90, 4
  If BallsOnPlayfield > 1 Then
'   bMultiBallMode = True
    bAutoPlunger = True
  End If

End Sub


Sub AddMultiball(nballs)
  mBalls2Eject = mBalls2Eject + nballs
  WriteToLog "AddMultiball", "mBalls2Eject = "  & mBalls2Eject
  CreateMultiballTimer.Enabled = True
End Sub

Sub CreateMultiballTimer_Timer()
  If bBallInPlungerLane Then
    WriteToLog "CreateMultiballTimer", "bBallInPlungerLane = True"
    Exit Sub
  Else
    If BallsOnPlayfield <MaxMultiballs Then
      CreateNewBall()
      mBalls2Eject = mBalls2Eject -1
      WriteToLog "CreateMultiballTimer", "mBalls2Eject = "  & mBalls2Eject
      If mBalls2Eject = 0 Then
        Me.Enabled = False
      End If
    Else
      mBalls2Eject = 0
      Me.Enabled = False
    End If
  End If
End Sub



Sub EndOfBall()

  WriteToLog "EndOfBall", ""
  bMultiBallMode = False
  bOnTheFirstBall = False
  Gameplayerup = "0"
  GameplayerNotup = "Default"
  PlayersAreInGame
  'DrainMissionPup.Enabled=True
  LS_Flash4.StopPlay
  If CurrentMission > 0 Then StopMission
  If PointMultiplier > 1 Then EndDoubleScoring
  pDMDSetPage(pDMDBlank)
  cLightMissionStart1.state = 0
  cLightShot1.state = 0
  cLightShot5.state = 0
  OrbUnlockShot = 0
  PFMultiplier.enabled = 0
  If SpinnerMultiplier > 1 Then EndSuperSpinners
  'EndWizard  'old end of wizard position
  LS_AllFlashers.play SeqStripe2VertOn, 50, 1     'drain effect. may need some variations?
  AudioCallout "drain"
  If NOT Tilted and NOT Slammed Then
    bFlippersEnabled = false
    EOB_Bonus = True '*EOB*
'   vpmtimer.addtimer 9860, "EndOfBall2 '"    '*EOB*
  Else

    vpmtimer.addtimer 900, "EndOfBall2 '"
  End If
End Sub



Sub EndOfBall2()
  WriteToLog "EndOfBall2", ""
  Tilted = False
  'Slammed = False
  Tilt = 0
  DisableTable False
  tilttableclear.enabled = False
  tilttime = 0

  GaldorReset=true
  GaldorShade true

  If(ExtraBallsAwards(CurrentPlayer) <> 0) Then
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1
    ' fix add LS on this one
    If(ExtraBallsAwards(CurrentPlayer) = 0) Then
      cLightShootAgain1.State = 0
      cLightShootAgain2.State = 0
    End If
    'make extraball have mission available without unlock, unless wizard ongoing
    if Not bWizardMode then
      OrbUnlockShot = 0
      cLightMissionStart1.state = 2
    end if
    vpmtimer.addtimer 300, "PlaySound ""SPC_extra_ball"",0,CalloutVol,0,0,1,1,1 '"
    CreateNewBall()
    ResetForNewPlayerBall
  Else
    BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1
    Ballblink=20
    EndWizard   'moved to here that extraball could work in wizard too
    If(BallsRemaining(CurrentPlayer) <= 0) Then
      CheckHighScore()
    Else
      EndOfBallComplete()
    End If
  End If
End Sub




Sub EndOfBallComplete()
  WriteToLog "EndOfBallComplete", ""
  ResetNewBallVariables
  Dim NextPlayer
  If(PlayersPlayingGame> 1) Then
    NextPlayer = CurrentPlayer + 1
    If(NextPlayer> PlayersPlayingGame) Then
      NextPlayer = 1
    End If
  Else
    NextPlayer = CurrentPlayer

  End If
  If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
    PuPEvent(503)
    Dof 503, DOFPulse
    EndOfGame()
  Else
    PlayerLights_SaveLights
    'set small font to previous player
    FlexDMD.Stage.GetLabel("Player" & currentplayer).font = FontScoreInactiv1
    CurrentPlayer = NextPlayer
    debug.print "next player " & CurrentPlayer
    PlayerLights_ResetInserts
    PlayerLights_NewPlayer

    AddScore 0
    ResetForNewPlayerBall()
    CreateNewBall()
    If PlayersPlayingGame> 1 Then
'     PlaySound "vo_player" &CurrentPlayer
      ' update new player score label
      FlexDMD.Stage.GetLabel("Player" & currentplayer).font = Font5by7W
    End If
    ScorbitBuildGameModes()
  End If
End Sub


dim WizardSoulLost : WizardSoulLost = false
Sub Drain_Hit()
  Drain.DestroyBall
  SwitchWasHit("Drain")
  BallsOnPlayfield = BallsOnPlayfield - 1
  WriteToLog "Drain_Hit", "BallsOnPlayfield = "  & BallsOnPlayfield

  if bCoreyMBOngoing Then
    showMissionVideo = 33
  elseif  bTracyMBOngoing Then
    showMissionVideo = 46
  elseif WizardPhase = 4 Then
    WizardSoulLost = true
  end if

  RandomSoundOutholeHit Drain
  If(bGameInPLay = True) AND(Tilted = False) AND(Slammed = False) Then
    If(bBallSaverActive = True) Then
      PupBallSaveMBControl
      DOF 127,2 'ball saved
      If BallsOnPlayfield=0 Then ShowBallsaver=1
      AddMultiball 1
      bAutoPlunger = True
      If bMultiBallMode = False Then
      End If
    Else
      DOF 126,2 'drain hit
      If BallsOnPlayfield=1 and mBalls2Eject=0 Then
        StopMBs
      End If
      If BallsOnPlayfield=0 and mBalls2Eject=0 Then
        DrainMissionPup.Enabled=True
        hatch false
        gioff
        PlaySound "EFX_sinkhole_shutdown",0,CalloutVol,0,0,1,1,1
        ShowBallLost=1
        bMultiBallMode = False
        vpmtimer.addtimer 2600, "EndOfBall '"
        Primitive019.blenddisablelighting = tracyEyesBrightness
        UpdateMaterial "tracyeyes"    ,0,0,0,0,0,0,1,RGB(17,217,27),0,0,False,True,0,0,0,0

      End If
    End If
  End If
End Sub


'*****************************************************************************************************************************************
'  BALL SAVE & LAUNCH
'*****************************************************************************************************************************************
'

Sub swBallSaveStart_hit
  If(bBallSaverReady = True) And (bBallSaverActive = False) Then
    EnableBallSaver 10
  End If
End Sub


Sub swPlungerRest_Hit()
  bBallInPlungerLane = True
  DOF 124,1 'ball in plunger lane

' SoundTroughToShooterLane()

  If bAutoPlunger Then
    autoplungerdelay.interval = 300
    autoplungerdelay.enabled = True
  Else
    ScorbitClaimQR(True)
  End If
' ShowBallCount true
  SwitchWasHit("swPlungerRest")
  PlungerLight.blenddisablelighting = 20
  'add correct brightness to new ball
  activeball.color = ballbrightness + (ballbrightness * 256) + (ballbrightness * 256 * 256)
  Plunger.timerenabled = 1
End Sub

sub autoplungerdelay_timer
  WriteToLog "autoplunger", "Autofire the ball"
  DOF 145,2 'ball autoplunged
  PlungerIM.Strength = 45
  PlungerIM.AutoFire
  PlungerIM.Strength = Plunger.MechStrength

  bAutoPlunger = False
  autoplungerdelay.enabled = False
end sub


sub plunger_timer
  'debug.print "Plunger position value. Not working with plunger key if analog pad connected: " & Plunger.Position
  PlungerLight.transy = 5 * Plunger.Position - 4.2
  If Plunger.Position > 10 Then
    bPlungerPulled = True
    CheckSkillshotReady
  Else
    bPlungerPulled = False
  End If
end sub

Sub swPlungerRest_UnHit()
  bBallInPlungerLane = False
  DOF 124,0 'ball not in plunger lane
  DOF 125,2 'ball exits plunger lane
  swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
  If Not bMultiBallMode Then EnableBallSaver 5
  ScorbitClaimQR(False)
  PlaySound "EFX_mima_launch",0,CalloutVol,0,0,1,1,1
  'PlaySoundAtLevelStaticRandomPitch SoundFX("LaserBlast1",DOFContactors), BlastSoundLevel, 0.055, activeball
  'Set BallInGun = activeball
  'BlastBall.Interval = 10
  'BlastBall.Enabled = 1


  Plunger.timerenabled = 0
  PlungerLight.transy = 0

  PlungerLaneFlashers true
  ShowBallCount false
  PlungerLight.blenddisablelighting = (1.5 * -activeball.vely) + 20
  vpmtimer.addtimer 100, "PlungerLight.blenddisablelighting = 0 '"
End Sub

sub ShowBallCount(Enabled)
  dim i
  if Enabled Then
    i = 4 - BallsRemaining(CurrentPlayer)
    If i > 0 And i < 4 Then
      FlexDMD.Stage.GetLabel("Splash3").Text = "BALL " & i
      FlexDMD.Stage.GetLabel("Splash3").SetAlignedPosition 64,30, FlexDMD_Align_Center
    Else
      FlexDMD.Stage.GetLabel("Splash3").Text = " "
    End If
    FlexDMD.Stage.GetLabel("Splash3").visible = true
    FlexDMD.stage.GetLabel("Score_1").text = "Player " & CurrentPlayer
  Else
    if PlayersPlayingGame > 1 then FlexDMD.stage.GetLabel("Score_1").text = " "
    FlexDMD.Stage.GetLabel("Splash3").text = " "
  end if
end sub

sub PlungerLaneFlashers(enabled)
  if Enabled Then
    plFlash1.timerenabled = True
  end If
end sub

dim PlungerLaneFlasherCounter : PlungerLaneFlasherCounter = 0
sub plFlash1_timer
  select case PlungerLaneFlasherCounter
    case 0:
      plFlash1.state = 1
      plFlash2.state = 0
      plFlash3.state = 0
    case 1:
      plFlash1.state = 0
      plFlash2.state = 1
      plFlash3.state = 0
    case 3:
      plFlash1.state = 0
      plFlash2.state = 0
      plFlash3.state = 1
    case 4:
      plFlash1.state = 1
      plFlash2.state = 0
      plFlash3.state = 0
    case 5:
      plFlash1.state = 0
      plFlash2.state = 1
      plFlash3.state = 0
    case 6:
      plFlash1.state = 0
      plFlash2.state = 0
      plFlash3.state = 1
    case 7:
      plFlash1.state = 1
      plFlash2.state = 0
      plFlash3.state = 0
    case 8:
      plFlash1.state = 0
      plFlash2.state = 1
      plFlash3.state = 0
    case 9:
      plFlash1.state = 0
      plFlash2.state = 0
      plFlash3.state = 1
  end Select

  PlungerLaneFlasherCounter = PlungerLaneFlasherCounter + 1

  if PlungerLaneFlasherCounter >= 9 Then
    plFlash1.timerenabled = false
    PlungerLaneFlasherCounter = 0
    plFlash1.state = 0
    plFlash2.state = 0
    plFlash3.state = 0
  end if

end sub

dim swPlungerRest_Timer_counter

Sub swPlungerRest_Timer
  swPlungerRest.TimerEnabled = 0
End Sub

Sub EnableBallSaver(seconds)
  if Not bBallSaverActive And WizardPhase<>4 Then
    bBallSaverActive = True
    bBallSaverReady = False
    BallSaverTimerExpired.Interval = 1000 * seconds
    BallSaverTimerExpired.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    cLightShootAgain1.BlinkInterval = 160
    cLightShootAgain2.BlinkInterval = 160
    cLightShootAgain1.State = 2
    cLightShootAgain2.State = 2
  End If
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimerExpired_Timer()
  WriteToLog "BallSaverTimerExpired", "Ballsaver ended"
  BallSaverTimerExpired.Enabled = False
  ' clear the flag
  Dim waittime
  waittime = 3000
  vpmtimer.addtimer waittime, "ballsavegrace'"
  ' if you have a ball saver light then turn it off at this point
  If ExtraBallsAwards(CurrentPlayer) > 0 Then
    cLightShootAgain1.State = 1
    cLightShootAgain2.State = 1
  Else
    cLightShootAgain1.State = 0
    cLightShootAgain2.State = 0
  End If
End Sub

Sub ballsavegrace
  bBallSaverActive = False
  WriteToLog "BallsaveGrace", "bBallSaverActive=False"
End Sub

Sub BallSaverSpeedUpTimer_Timer()
  'debug.print "Ballsaver Speed Up Light"
  BallSaverSpeedUpTimer.Enabled = False
  ' Speed up the blinking
  cLightShootAgain1.BlinkInterval = 80
  cLightShootAgain2.BlinkInterval = 80
  cLightShootAgain1.State = 2
  cLightShootAgain2.State = 2
End Sub



'*****************************************************************************************************************************************
'  TABLE VARIABLES
'*****************************************************************************************************************************************
'


'*****************************************************************************************************************************************
'  GAME STARTING & RESETS
'*****************************************************************************************************************************************
'

Sub Game_Init() 'called at the start of a new game

End Sub

Sub StopEndOfBallMode()              'this sub is called after the last ball is drained

End Sub

Sub ResetNewBallVariables()

End Sub

Sub TurnOffPlayfieldLights()
  Dim a
  For each a in alights
    a.State = 0
  Next
End Sub

Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released
  TurnOffPlayfieldLights()
End Sub





'*****************************************************************************************************************************************
'  TARGETS
'*****************************************************************************************************************************************

sub MimaTargetBig_Hit
    If Tilted Then Exit Sub
  DOF 109,2
  startB2S 4
  AddScore score_MimaTargetBig_Hit
  MimaTargetWasHit
  TargetMimaStep=0:TargetMimaStep=0:MimaTargetBig.TimerEnabled = true
  'LS_PortalWall.stopPlay
  'LS_PortalWall.Play SeqRandom,1,,1000

end sub

dim TargetMimaStep, MimaCount
MimaCount = 0

Sub MimaTargetBig_timer
  dim x
  Select Case TargetMimaStep
    Case 1:for each x in cMimaTarget:x.transx = 5:x.transy = 5:next: if PointMultiplier<2 then cpMimaTarget.state = 1
        Case 2:for each x in cMimaTarget:x.transx = 3:x.transy = 3:next
        Case 3:for each x in cMimaTarget:x.transx = 1:x.transy = 1:next
        Case 4:for each x in cMimaTarget:x.transx = 0:x.transy = 0:next: MimaTargetBig.TimerEnabled = False:TargetMimaStep = 0 : if PointMultiplier<2 then cpMimaTarget.state = 0
    End Select
  TargetMimaStep = TargetMimaStep + 1
end sub

dim PFmultiHitCount : PFmultiHitCount = 0
Sub MimaTargetWasHit
  if Not bWizardMode then
    'after 3 hits, start 30s PF multiplier, if 3 more within that 30s, increase multiplier until 5x
    PFmultiHitCount = PFmultiHitCount + 1
    if PFmultiHitCount => 3 Then
      PFmultiHitCount = 0
      'Playsound("EFX_tracy_bash")
      shake_tracy_shake.enabled = true
      select case PointMultiplier
        Case 1 : PointMultiplier = 2 : cpMimaTarget.blinkinterval=250 : AudioCallout "double scoring"       : PFMultiplierCounter = 30 : DMD_ShowText "2X FOR 30 SECONDS",4,FontScoreInactiv2,28,True,40,3000
        PuPEvent(580)
        DOF 580, DOFPulse
        Case 2 : PointMultiplier = 3 : cpMimaTarget.blinkinterval=150 : AudioCallout "3x bonus"       : PFMultiplierCounter = 30 : DMD_ShowText "3X FOR 30 SECONDS",4,FontScoreInactiv2,28,True,40,3000
        PuPEvent(533)
        DOF 533, DOFPulse
        Case 3 : PointMultiplier = 4 : cpMimaTarget.blinkinterval=100  : AudioCallout "4x bonus"      : PFMultiplierCounter = 30 : DMD_ShowText "4X FOR 30 SECONDS",4,FontScoreInactiv2,28,True,40,3000
        PuPEvent(534)
        DOF 534, DOFPulse
        Case 4 : PointMultiplier = 5 : cpMimaTarget.blinkinterval=50  : AudioCallout "5x bonus"       : PFMultiplierCounter = 30 : DMD_ShowText "5X FOR 30 SECONDS",4,FontScoreInactiv2,28,True,40,3000
        PuPEvent(535)
        DOF 535, DOFPulse
      end select
      WriteToLog "PF Multiplier","Level=" & PointMultiplier
      solBloodTrail true
      PFMultiplier.enabled = 1
    Else
      Playsound "EFX_switch_bleep0" & RndInt(1,3),0,CalloutVol,0,0,1,1,1
    end If
  end if
end sub

dim PFMultiplierCounter:PFMultiplierCounter = 30
sub PFMultiplier_timer
  PFMultiplierCounter = PFMultiplierCounter - 1
  WriteToLog "PF Multiplier ongoing","Time=" & PFMultiplierCounter
  if PFMultiplierCounter <= 0 then
    PointMultiplier = 1
    PFMultiplierCounter = 30
    WriteToLog "PF Multiplier end","Level=" & PointMultiplier
    cpMimaTarget.blinkinterval=250
    cpMimaTarget.state = 0
    solBloodTrail false
    PFMultiplier.enabled = 0
  else
    cpMimaTarget.state = 2
  end if
end sub


Sub AdvanceToMimaMB
    If Tilted Then Exit Sub
  If not bMultiBallMode And Not bWizardMode And Not cLightMimaMB.state = 1 And Not CurrentMission = 1 then 'advance disabled when in Hunt mission
    If MimaCount = 0 Then
      MimaCount = 1
'     Playsound("EFX_switch_bleep01")
      cTargetMima1.state = 1
    ElseIf MimaCount = 1 Then
      MimaCount = 2
      cTargetMima2.state = 1
'     Playsound("EFX_switch_bleep02")
    ElseIf MimaCount= 2 Then
      MimaCount = 3
      cTargetMima3.state = 1
'     Playsound("EFX_switch_bleep03")
      if cLightMimaMB.state <> 1 Then Rise_RampMima.enabled=True
    Else
      AddScore score_MimaTargetBig_Maxed
'     Playsound("EFX_switch_bleep04")
    End If
    WriteToLog "AdvanceToMimaMB","MimaCount=" & MimaCount
  End if
End Sub


Sub TargetScavenge1_Hit : DTHit 1 : TargetBouncer Activeball, 1.3 : End Sub
Sub TargetScavenge2_Hit : DTHit 2 : TargetBouncer Activeball, 1.3 : End Sub
Sub TargetScavenge3_Hit : DTHit 3 : TargetBouncer Activeball, 1.3 : End Sub



Sub TargetScavenge1_Timer
  TargetScavenge1.timerenabled=False
  If DTIsDropped(1) And DTIsDropped(2) And DTIsDropped(3) And Not bWizardMode Then
    addscore score_DropsComplete
    if Not bMultiBallMode then
      cLightScavengeScoop.state = 2
    Else                    'lit scavenge only after the MB
      if LightScavengeScoopState = 0 then
        LightScavengeScoopState = 2
      end if
    end if

    If cLightMystery1.state = 0 Then
      cLightMystery1.state = 2
      WriteToLog "TargetScavenge1 ScavangeTargets","Light 1 Blinking"
    ElseIf cLightMystery2.state = 0 Then
      cLightMystery2.state=2
      WriteToLog "TargetScavenge1 ScavangeTargets","Light 2 Blinking"
    ElseIf cLightMystery3.state = 0 Then
      cLightMystery3.state=2
      WriteToLog "TargetScavenge1 ScavangeTargets","Light 3 Blinking"
    Else
    End If
    LS_Scavenge.Play SeqBlinking,,5,25
  End If
End Sub



Sub TargetVascan1_Hit
' DOF 112,2
  DTHit 4
  TargetBouncer Activeball, 1.3
'    SwitchWasHit("TargetVascan")
' If VRRoom > 0 Then
'   VRBGFL3
' End If
' playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
' GaldorReset=true:GaldorQ1 = "fYa"
' CheckPortalRoomShot 1
End Sub

Sub TargetVascan2_Hit
  DTHit 5
  TargetBouncer Activeball, 1.3
' DOF 112,2
'    SwitchWasHit("TargetVascan")
' startB2S 3
' If VRRoom > 0 Then
'   VRBGFL3
' End If
' playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
' GaldorReset=true:GaldorQ1 = "fYb"
' CheckPortalRoomShot 2
End Sub

Sub TargetVascan3_Hit
  DTHit 6
  TargetBouncer Activeball, 1.3
' DOF 112,2
'    SwitchWasHit("TargetVascan")
' startB2S 3
' If VRRoom > 0 Then
'   VRBGFL3
' End If
' playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
' GaldorReset=true:GaldorQ1 = "fYa"
' CheckPortalRoomShot 3
End Sub

Sub TargetVascan4_Hit
  DTHit 7
  TargetBouncer Activeball, 1.3
' DOF 112,2
'    SwitchWasHit("TargetVascan")
' startB2S 3
' If VRRoom > 0 Then
'   VRBGFL3
' End If
' playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
' GaldorReset=true:GaldorQ1 = "fYb"
' CheckPortalRoomShot 4
End Sub

Sub TargetVascan5_Hit
  DTHit 8
  TargetBouncer Activeball, 1.3
' DOF 112,2
'    SwitchWasHit("TargetVascan")
' startB2S 3
' If VRRoom > 0 Then
'   VRBGFL3
' End If
' playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
' GaldorReset=true:GaldorQ1 = "fYa"
' CheckPortalRoomShot 5
End Sub

Sub TargetVascan6_Hit
  DTHit 9
  TargetBouncer Activeball, 1.3
' DOF 112,2
'    SwitchWasHit("TargetVascan")
' startB2S 3
' If VRRoom > 0 Then
'   VRBGFL3
' End If
' playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
' GaldorReset=true:GaldorQ1 = "fYb"
' CheckPortalRoomShot 6
End Sub



Sub swUnderPF1_Hit
    If Tilted Then Exit Sub
    SwitchWasHit("UnderPF")
' msgbox "hit left"
  CheckWizardUPFShot 1
End Sub

Sub swUnderPF2_Hit
    If Tilted Then Exit Sub
    SwitchWasHit("UnderPF")
' msgbox "hit right"
  CheckWizardUPFShot 2
End Sub


'*****************************************************************************************************************************************
'  SLINGSHOTS
'*****************************************************************************************************************************************



Dim RStep, Lstep


Sub RightSlingShot_Slingshot
    if tilted then exit sub

  startB2S 3
  If VRRoom > 0 Then
    VRBGFL3
  End If
  BG_stars=-1
  SwitchWasHit("RightSling")

  Addscore score_SlingShot
  RotateScavangeLights
  PlaySound "EFX_sling",0,CalloutVol,0,0,1,1,1
' RandomSoundSlingshotRight()
  RandomSoundSlingshotMainRightDOF swRightInlane, 104, DOFPulse
  RSling.Visible = 0
  RSling1.Visible = 1
  sling1.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1

End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
    Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    if tilted then exit sub
  DOF 103,2
  startB2S 2
  If VRRoom > 0 Then
    VRBGFL2
  End If
  BG_stars=-2

  SwitchWasHit("LeftSling")
  Addscore score_SlingShot
  RotateScavangeLights
  PlaySound "EFX_sling",0,CalloutVol,0,0,1,1,1
' RandomSoundSlingshotLeft swLeftInlane
' DOF 103, DOFPulse
' RandomSoundSlingshotLeft()
  RandomSoundSlingshotMainLeftDOF swLeftInlane, 103, DOFPulse
  LSling.Visible = 0
  LSling1.Visible = 1
  sling2.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1

End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
    Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub


'*****************************************************************************************************************************************
'   LaneLights
'*****************************************************************************************************************************************
Dim LaneInsertStates(7)

Sub SaveLaneInsertStates
  LaneInsertStates(1)  = cLightTopLeftLane.State
  LaneInsertStates(2)  = cLightTopRightLane.State
  LaneInsertStates(3)  = cLightLeftOutlane.State
  LaneInsertStates(4)  = cLightLeftInlane.State
  LaneInsertStates(5)  = cLightRightInlane.State
  LaneInsertStates(6)  = cLightRightOutlane.State
end sub

Sub ClearLaneInsertStates
  cLightTopLeftLane.State   = 0
  cLightTopRightLane.State  = 0
  cLightLeftOutlane.State   = 0
  cLightLeftInlane.State    = 0
  cLightRightInlane.State   = 0
  cLightRightOutlane.State  = 0
end sub

sub RestoreLaneInsertStates
  cLightTopLeftLane.State   = LaneInsertStates(1)
  cLightTopRightLane.State  = LaneInsertStates(2)
  cLightLeftOutlane.State   = LaneInsertStates(3)
  cLightLeftInlane.State    = LaneInsertStates(4)
  cLightRightInlane.State   = LaneInsertStates(5)
  cLightRightOutlane.State  = LaneInsertStates(6)
end sub

Sub RotateLaneLightsLeft
  Dim i
  i=cLightTopLeftLane.State
  cLightTopLeftLane.state=cLightTopRightLane.State
  cLightTopRightLane.state=i
  i=cLightLeftOutlane.State
  cLightLeftOutlane.state=cLightLeftInlane.State
  cLightLeftInlane.state=cLightRightInlane.State
  cLightRightInlane.state=cLightRightOutlane.State
  cLightRightOutlane.state=i
End Sub

Sub RotateLaneLightsRight
  Dim i
  i=cLightTopRightLane.State
  cLightTopRightLane.state=cLightTopLeftLane.State
  cLightTopLeftLane.state=i
  i=cLightRightOutlane.State
  cLightRightOutlane.state=cLightRightInlane.State
  cLightRightInlane.state=cLightLeftInlane.State
  cLightLeftInlane.state=cLightLeftOutlane.State
  cLightLeftOutlane.state=i
End Sub


'*****************************************************************************************************************************************
'   BUMPERS
'*****************************************************************************************************************************************

Sub Bumper1_Hit
  If cLightTopLeftLane.state=2 or cLightTopRightLane.state=2 Then skillshotoff
  If NOT Tilted Then
'   DOF 105,2
    startB2S 6
    DMD_bumper  ' Multiballactive=0 only
    AddBumperHit
    TracyPupCount
    If PupBumpertenthousand= True Then
    BumperTenthousand
    End if
    BumperTwentythousand
    CheckHeartOfSteelShot
'   RandomSoundBumperTop Bumper1
'   RandomSoundPopBumperDOF Bumper1, 105, DOFPulse
    RandomSoundJetBumperLeftDOF Bumper1, 105, DOFPulse
'   RandomSoundBallBounceAfterPopBumper()
    PlaySound "EFX_jolt",0,CalloutVol/2,0,0,1,1,1
    If CurrentMission=0 Then Next_MissionSelect
    CounterBallSearch = 0
  End If

End Sub



Sub Bumper2_Hit
  If cLightTopLeftLane.state=2 or cLightTopRightLane.state=2 Then skillshotoff
  If NOT Tilted Then
'   DOF 106,2
    DMD_bumper  ' Multiballactive=0 only
    startB2S 6
    AddBumperHit
    TracyPupCount
    If PupBumpertenthousand= True Then
    BumperTenthousand
    End if
    BumperTwentythousand
    CheckHeartOfSteelShot
    PlaySound "EFX_jolt",0,CalloutVol/2,0,0,1,1,1
'   RandomSoundBumperMiddle Bumper2
'   RandomSoundPopBumperDOF Bumper2, 106, DOFPulse
    RandomSoundJetBumperLowDOF Bumper2, 106, DOFPulse
'   RandomSoundBallBounceAfterPopBumper()
    If CurrentMission=0 Then Next_MissionSelect
'   shake_tracy_shake.enabled = true
    CounterBallSearch = 0
  End If

End Sub



Sub Bumper3_Hit
  If cLightTopLeftLane.state=2 or cLightTopRightLane.state=2 Then skillshotoff
  If NOT Tilted Then
'   DOF 107,2
    DMD_bumper  ' Multiballactive=0 only
    startB2S 6
    AddBumperHit
    TracyPupCount
    If PupBumpertenthousand= True Then
    BumperTenthousand
    End if
    BumperTwentythousand
    CheckHeartOfSteelShot
    PlaySound "EFX_jolt",0,CalloutVol/2,0,0,1,1,1
'   RandomSoundBumperBottom Bumper3
'   RandomSoundPopBumperDOF Bumper3, 107, DOFPulse
    RandomSoundJetBumperUpDOF Bumper3, 107, DOFPulse
'   RandomSoundBallBounceAfterPopBumper()
    If CurrentMission=0 Then Next_MissionSelect
    CounterBallSearch = 0
  End If

End Sub


Sub DMD_bumper
  if Not CurrentMission = 6 then
    If Not bMultiBallMode Then
      If BG_Bump1=0 Then
        BG_Bump1=10
      Elseif BG_Bump2=0 Then
        BG_Bump2=10
      End If
    End If
  Else
    DMD_ShowImages "jolt",3,50,2000,0 ' jolt fast blink for heartof steel = need to fix the image so it centered ?
  end if
End Sub


Sub DMD_bumper_old
  If MultiBallActive=0 Then
    If BG_Bump1=0 Then
      BG_Bump1=10
    Elseif BG_Bump2=0 Then
      BG_Bump2=10
    End If
  End If
End Sub


Dim ShowSuperJets
Sub AddBumperHit
    if tilted then exit sub
  'First check if we need to lite a Tracy MB light

  BumperHits(CurrentPlayer) = BumperHits(CurrentPlayer) + 1
  Tracy_bumpers = BumperHits(CurrentPlayer) Mod 10

  If Tracy_bumpers = 0 Then
    AddTracyLight
  Elseif cLightTracy1.state=0 Or cLightTracy2.state=0 Or cLightTracy3.state=0 Or cLightTracy4.state=0 Or cLightTracy5.state=0 Then
    if cLightTracyMB.state = 0 then 'don't show this if tracy mb is ready or played
      Show_TracyHits = 31
    end if
  End If



  'Next accumulate count for Super Jets
  SuperJets = SuperJets + 1
  If SuperJets = 50 Then AudioCallout "super jets":Superjetspop: ShowSuperJets=1


  If SuperJets < 50 Then
    AddScore score_Bumper
    bumpersplash = 31
  Else
    AddScore score_SuperJetBumper
  End If

  'Flash insert
  cLightBumperHit.state=1
  cLightBumperHit.Timerenabled=True

End Sub

Sub Superjetspop
  If SuperJets = 50 Then
  PupBumpertenthousand = False
  PupBumpersuperjets = True
  BumperStages
  PuPEvent(633)
  DOF 471 , DOFPulse
 End if
End Sub


Sub cLightBumperHit_Timer
  cLightBumperHit.state=0
  if PointMultiplier > 1 Then
    ptracyCross.visible = false
    cTracyCross.state = 0
  end if
  cLightBumperHit.Timerenabled=False
End Sub


'**********************************************************************************************************
'    Left outlane kicker
'**********************************************************************************************************

plunger1.PullBack

Sub swLeftOutlane_hit
  if CurrentMission = 5 Or WizardPhase = 3 then exit sub
  WriteToLog "swLeftOutlane_hit",""
  cLightLeftOutlane.state=1
  playsound "EFX_effect3",0,CalloutVol,0,0,1,1,1
  CheckLiteKB
  TracyAttention.Enabled=True
  SwitchWasHit("swLeftOutLane")
End Sub


Sub swLeftOutlane_unhit
  RandomSoundLanes
  If Tilted Then Exit Sub
  If cLightKB2.state=1 Then
    DOF 111,2
    DOF 575, DOFPulse
    PuPEvent(575)
    WriteToLog "plungerKB","AutoFire"
    LS_Flipperlanes.Play SeqBlinking,,5,13
    swLeftOutLane.enabled=False
    vpmTimer.addTimer 250,"swLeftOutLane.enabled=True '"
    plungerKB.Strength = 45
    plungerKB.AutoFire
    plunger1.Fire 'This is just for show (momentum trasfer is Zero)
'   playsound SoundFX("fx_kicker",DOFContactors)
    RandomSoundLockingKickerSolenoid
    plunger1.PullBack
    'EnableBallSaver 1
    If cLightKB1.state=1 Then
      cLightKB1.state=0
    Else
      cLightKB2.state=0
    End If
    AddScore score_Inlanes
  Else
    AddScore score_Outlanes
  End If
End Sub


Sub swLeftInLane_Hit
  RandomSoundInlanes
  If Tilted Then Exit Sub
  if CurrentMission = 5 Or WizardPhase = 3 then exit sub
  AddScore score_Inlanes
  cLightLeftInlane.state=1
  playsound "EFX_effect3",0,CalloutVol,0,0,1,1,1
  CheckLiteKB
' TracyAttention.Enabled=True
  SwitchWasHit("swLeftInLane")
End Sub


Sub swRightInLane_Hit
  RandomSoundInlanes
  If Tilted Then Exit Sub
  if CurrentMission = 5 Or WizardPhase = 3 then exit sub
  AddScore score_Inlanes
  cLightRightInlane.state=1
  playsound "EFX_effect3",0,CalloutVol,0,0,1,1,1
  CheckLiteKB
' TracyAttention.Enabled=True
  SwitchWasHit("swRightInLane")
End Sub


Sub swRightOutLane_Hit
  RandomSoundLanes
  If Tilted Then Exit Sub
  if CurrentMission = 5 Or WizardPhase = 3 then exit sub
  AddScore score_Outlanes
  cLightRightOutlane.state=1
  playsound "EFX_effect3",0,CalloutVol,0,0,1,1,1
  CheckLiteKB
  TracyAttention.Enabled=True
  'AudioCallout "outlane drain"
  SwitchWasHit("swRightOutLane")
End Sub


Sub CheckLiteKB
  If Tilted Then Exit Sub
  If cLightLeftOutlane.state=1 And cLightLeftInlane.state=1 And cLightRightInlane.state=1 And cLightRightOutlane.state=1 Then
    DOF 134,2 'lite kickback
    LS_Flipperlanes.Play SeqBlinking,,10,20
    cLightLeftOutlane.state=0
    cLightLeftInlane.state=0
    cLightRightInlane.state=0
    cLightRightOutlane.state=0
    stopsound("EFX_effect3")
    playsound "EFX_mima_launch",0,CalloutVol,0,0,1,1,1
    If cLightKB2.state=0 Then
      cLightKB2.state=1
      AddScore score_LagoCompleted
      PuPEvent(587)
      DOF 587, DOFPulse
      DMD_ShowText "LAGO collected " & score_LagoCompleted,4,FontScoreInactiv2,30,True,40,3000
    Elseif cLightKB1.state=0 Then
      cLightKB1.state=1
      AddScore score_LagoCompleted
      PuPEvent(587)
      DOF 587, DOFPulse
      DMD_ShowText "LAGO collected " & score_LagoCompleted,4,FontScoreInactiv2,30,True,40,3000
    Else
      AddScore score_KBMaxed  'extra points instead of 3rd stacked KB
      AudioCallout "LAGO miracle"
      PuPEvent(588)
      DOF 588, DOFPulse
      DMD_ShowText "LAGO collected " & score_KBMaxed,4,FontScoreInactiv2,30,True,30,3000
    End If
  End If
End Sub







'******************************************************
'         FLIPPERS
'******************************************************


Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

'' iaakki Rubberizer
'sub Rubberizer(parm)
' if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
'   'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
'   activeball.angmomz = activeball.angmomz * 1.2
'   activeball.vely = activeball.vely * 1.2
'   'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
' Elseif parm <= 2 and parm > 0.2 Then
'   'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
'   activeball.angmomz = activeball.angmomz * -1.1
'   activeball.vely = activeball.vely * 1.4
'   'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
' end if
'end sub
'
'' apophis Rubberizer
'sub Rubberizer2(parm)
' if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
'   'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
'   activeball.angmomz = -activeball.angmomz * 2
'   activeball.vely = activeball.vely * 1.2
'   'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
' Elseif parm <= 2 and parm > 0.2 Then
'   'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
'   activeball.angmomz = -activeball.angmomz * 0.5
'   activeball.vely = activeball.vely * (1.2 + rnd(1)/3 )
'   'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
' end if
'end sub

'******************************************************
'  SLINGSHOT CORNER SPIN
'******************************************************

if bSlingSpin Then
  LSlingSpin.enabled = true
  RSlingSpin.enabled = true
end if

const MaxSlingSpin = 350
const MinCollVelX = 7
const MinCorVelocity = 10

dim lspinball
sub LSlingSpin_hit
' debug.print "lhit: " & activeball.angmomz & " velx: " & activeball.velx & " vely: " & activeball.vely & " corvel: " & cor.ballvel(activeball.id)
  if activeball.velx < -MinCollVelX And cor.ballvel(activeball.id) > MinCorVelocity then
'   debug.print "lspinball set, vel: " & activeball.angmomz
    lspinball = activeball.id
  end if

end sub

sub LSlingSpin_unhit
  if lspinball <> -1 then
    activeball.angmomz = (cor.ballvel(activeball.id)/2) * activeball.angmomz
    if activeball.angmomz  > MaxSlingSpin then activeball.angmomz  = MaxSlingSpin
    if activeball.angmomz  < -MaxSlingSpin then activeball.angmomz  = -MaxSlingSpin
'   debug.print "lunhit: " & activeball.angmomz & " velx: " & activeball.velx & " vely: " & activeball.vely & ", vel: " & cor.ballvel(activeball.id)
    lspinball = -1
  end if
end sub


dim rspinball
sub RSlingSpin_hit
' debug.print "rhit: " & activeball.angmomz & " velx: " & activeball.velx & " vely: " & activeball.vely & " corvel: " & cor.ballvel(activeball.id)
  if activeball.velx > MinCollVelX And cor.ballvel(activeball.id) > MinCorVelocity then
'   debug.print "rspinball set, vel: " & activeball.angmomz
    rspinball = activeball.id
  end if

end sub

sub RSlingSpin_unhit
  if rspinball <> -1 then
    activeball.angmomz = (cor.ballvel(activeball.id)/2) * activeball.angmomz
    if activeball.angmomz  > MaxSlingSpin then activeball.angmomz  = MaxSlingSpin
    if activeball.angmomz  < -MaxSlingSpin then activeball.angmomz  = -MaxSlingSpin
'   debug.print "runhit: " & activeball.angmomz & " velx: " & activeball.velx & " vely: " & activeball.vely & ", vel: " & cor.ballvel(activeball.id)
    rspinball = -1
  end if
end sub




'******************************************************
'       FLIPPER AND RUBBER CORRECTION
'******************************************************


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
                                        BOT(b).velx = BOT(b).velx / 1.5
                                        BOT(b).vely = BOT(b).vely - 0.5
'                   debug.print "nudge"
                                end If
                        Next
                End If
        Else
                If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 15 then
                        EOSNudge1 = 0
                end if
        End If
End Sub

'*****************
' Maths
'*****************
'Const PI = 3.1415927

'Function dSin(degrees)
'        dsin = sin(degrees * Pi/180)
'End Function
'
'Function dCos(degrees)
'        dcos = cos(degrees * Pi/180)
'End Function

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
'        dim pi
'        pi = 4*Atn(1)

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
Const EOSReturn = 0.018

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

sub zCol_Rubber_Corner_008_hit
  RubbersD.dampen Activeball
end sub

sub zCol_Rubber_Corner_006_hit
  RubbersD.dampen Activeball
end sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub

Sub dPlastics_Hit(idx)
  PlasticsD.Dampen Activeball
End Sub

dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

dim PlasticsD : Set PlasticsD = new Dampener  'this is just rubber but cut down to 95%...
PlasticsD.name = "Plastics"
PlasticsD.debugOn = False 'shows info in textbox "TBPout"
PlasticsD.Print = False 'debug, reports in debugger (in vel, out cor)
PlasticsD.CopyCoef RubbersD, 0.95

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
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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
'     debug.print "dampen f"
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub

End Class


'******************************************************
'   TRACK ALL BALL VELOCITIES
'     FOR RUBBER DAMPENER AND DROP TARGETS
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



'******************************************************
'****  DROP TARGETS by Rothbauerw
'******************************************************


'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

'Define a variable for each drop target
Dim DT1, DT2, DT3, DT4, DT5, DT6,DT7, DT8, DT9

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
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
' Values for annimate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target

DT1 = Array(TargetScavenge1, TargetScavenge1a, TargetScavenge1p, 1, 0)
DT2 = Array(TargetScavenge2, TargetScavenge2a, TargetScavenge2p, 2, 0)
DT3 = Array(TargetScavenge3, TargetScavenge3a, TargetScavenge3p, 3, 0)
DT4 = Array(TargetVascan1, TargetVascan1a, TargetVascan1p, 4, 0)
DT5 = Array(TargetVascan2, TargetVascan2a, TargetVascan2p, 5, 0)
DT6 = Array(TargetVascan3, TargetVascan3a, TargetVascan3p, 6, 0)
DT7 = Array(TargetVascan4, TargetVascan4a, TargetVascan4p, 7, 0)
DT8 = Array(TargetVascan5, TargetVascan5a, TargetVascan5p, 8, 0)
DT9 = Array(TargetVascan6, TargetVascan6a, TargetVascan6p, 9, 0)

Dim DTArray
DTArray = Array(DT1, DT2, DT3, DT4, DT5, DT6, DT7, DT8, DT9)
Dim DTIsDropped(9)
Dim DTWasDropped(9)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
'Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
'Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


Dim DTArray0, DTArray1, DTArray2, DTArray3, DTArray4
Redim DTArray0(UBound(DTArray)), DTArray1(UBound(DTArray)), DTArray2(UBound(DTArray)), DTArray3(UBound(DTArray)), DTArray4(UBound(DTArray))

Dim DTIdx
For DTIdx = 0 to UBound(DTArray)
   Set DTArray0(DTIdx) = DTArray(DTIdx)(0)
   Set DTArray1(DTIdx) = DTArray(DTIdx)(1)
   Set DTArray2(DTIdx) = DTArray(DTIdx)(2)
   DTArray3(DTIdx) = DTArray(DTIdx)(3)
   DTArray4(DTIdx) = DTArray(DTIdx)(4)
Next

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

' PlayTargetSound
  DTArray4(i) =  DTCheckBrick(Activeball,DTArray2(i))
  If DTArray4(i) = 1 or DTArray4(i) = 3 or DTArray4(i) = 4 Then
    DTBallPhysics Activeball, DTArray2(i).rotz, DTMass
  End If
  DoDTAnim
  Select Case Switch
    case 1: LightMystery1C.visible = false : SoundDropTarget3Bank TargetScavenge1p
    case 2: LightMystery2C.visible = false : SoundDropTarget3Bank TargetScavenge2p
    case 3: LightMystery3C.visible = false : SoundDropTarget3Bank TargetScavenge3p
    case 4: SoundDropTarget TargetVascan1p
    case 5: SoundDropTarget TargetVascan2p
    case 6: SoundDropTarget TargetVascan3p
    case 7: SoundDropTarget TargetVascan4p
    case 8: SoundDropTarget TargetVascan5p
    case 9: SoundDropTarget TargetVascan6p
  end Select
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)
  DTWasDropped(Switch) = DTIsDropped(switch)
  DTIsDropped(switch) = False

  DTArray4(i) = -1
  DoDTAnim
  Select Case Switch
    case 1: LightMystery1C.visible = true
    case 2: LightMystery2C.visible = true
    case 3: LightMystery3C.visible = true
  end Select
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray4(i) = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 to uBound(DTArray)
    If DTArray3(i) = switch Then DTArrayID = i:Exit Function
  Next
End Function


sub DTBallPhysics(aBall, angle, mass)
  dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
  calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

  aBall.velx = calc1 * cos(rangle) + calc2
  aBall.vely = calc1 * sin(rangle) + calc3
End Sub


'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    If DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function


Sub DoDTAnim()
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray4(i) = DTAnimate(DTArray0(i),DTArray1(i),DTArray2(i),DTArray3(i),DTArray4(i))
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch,  animate)
  dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If (animate = 1 or animate = 4) and animtime < DTDropDelay Then
    primary.collidable = 0
  If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    DTAnimate = animate
    Exit Function
    elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    animate = 2
'   SoundDropTargetDrop prim
'   debug.print "drop " & Switch
    Select Case Switch
      case 1: 'SoundDropTarget_Release TargetScavenge1p
      case 2: 'SoundDropTarget_Release TargetScavenge2p
      case 3: 'SoundDropTarget_Release TargetScavenge3p
      case 4: SoundDropTarget_Release TargetVascan1p
      case 5: SoundDropTarget_Release TargetVascan2p
      case 6: SoundDropTarget_Release TargetVascan3p
      case 7: SoundDropTarget_Release TargetVascan4p
      case 8: SoundDropTarget_Release TargetVascan5p
      case 9: SoundDropTarget_Release TargetVascan6p
    end Select
  End If

  if animate = 2 Then
    transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
    if prim.transz > -DTDropUnits  Then
      prim.transz = transz
    end if

    prim.rotx = DTMaxBend * cos(rangle)/2
    prim.roty = DTMaxBend * sin(rangle)/2

    if prim.transz <= -DTDropUnits Then
      prim.transz = -DTDropUnits
      secondary.collidable = 0
      DTAction switchid
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    end If
  End If

  If animate = 3 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
  elseif animate = 3 and animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  if animate = -1 Then
    transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

    If prim.transz = -DTDropUnits Then
      Dim BOT, b
      BOT = GetBalls

      For b = 0 to UBound(BOT)
        If InRotRect(BOT(b).x,BOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and BOT(b).z < prim.z+DTDropUnits+25 Then
          BOT(b).velz = 20
        End If
      Next
    End If

    if prim.transz < 0 Then
      prim.transz = transz
    elseif transz > 0 then
      prim.transz = transz
    end if

    if prim.transz > DTDropUpUnits then
      DTAnimate = -2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1

  End If

  if animate = -2 and animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
    if prim.transz < 0 then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
'debug.print "raise " & Switch
'     Select Case Switch
'       case 1: DropTarget_Reset_Coil 1, TargetScavenge1p
'       case 2: DropTarget_Reset_Coil 2, TargetScavenge2p
'       case 3: DropTarget_Reset_Coil 3, TargetScavenge3p
'       case 4: DropTarget_Reset_Coil 4, TargetVascan1p
'       case 5: DropTarget_Reset_Coil 5, TargetVascan2p
'       case 6: DropTarget_Reset_Coil 6, TargetVascan3p
'       case 7: DropTarget_Reset_Coil 7, TargetVascan4p
'       case 8: DropTarget_Reset_Coil 8, TargetVascan5p
'       case 9: DropTarget_Reset_Coil 9, TargetVascan6p
'     end Select
    end If
  End If
End Function

Sub DTAction(switchid)
    if tilted then exit sub
  DTIsDropped(switchid) = True
  Select Case switchid
    Case 1:
'     DOF 110,2
      startB2S 4
      playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
      AddScore score_DropsHit
      LightFlash6b1d.state = 1
      TargetScavenge1.timerenabled=True
    Case 2:
'     DOF 110,2
      startB2S 4
      playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
      AddScore score_DropsHit
      LightFlash6b2d.state = 1
      TargetScavenge1.timerenabled=True
    Case 3:
'     DOF 110,2
      startB2S 4
      playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
      AddScore score_DropsHit
      LightFlash6b3d.state = 1
      TargetScavenge1.timerenabled=True
        Case 4:
            DOF 112,2
            SwitchWasHit("TargetVascan")
            If VRRoom > 0 Then
                VRBGFL3
            End If
            playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
            GaldorReset=true:GaldorQ1 = "fYa"
            CheckPortalRoomShot 1
        Case 5:
          DOF 112,2
            SwitchWasHit("TargetVascan")
            startB2S 3
            If VRRoom > 0 Then
                VRBGFL3
            End If
            playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
            GaldorReset=true:GaldorQ1 = "fYb"
            CheckPortalRoomShot 2
        Case 6:
          DOF 112,2
            SwitchWasHit("TargetVascan")
            startB2S 3
            If VRRoom > 0 Then
                VRBGFL3
            End If
            playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
            GaldorReset=true:GaldorQ1 = "fYa"
            CheckPortalRoomShot 3
        Case 7:
          DOF 112,2
            SwitchWasHit("TargetVascan")
            startB2S 3
            If VRRoom > 0 Then
                VRBGFL3
            End If
            playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
            GaldorReset=true:GaldorQ1 = "fYb"
            CheckPortalRoomShot 4
        Case 8:
          DOF 112,2
            SwitchWasHit("TargetVascan")
            startB2S 3
            If VRRoom > 0 Then
                VRBGFL3
            End If
            playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
            GaldorReset=true:GaldorQ1 = "fYa"
            CheckPortalRoomShot 5
        Case 9:
            DOF 112,2
            SwitchWasHit("TargetVascan")
            startB2S 3
            If VRRoom > 0 Then
                VRBGFL3
            End If
            playsound "EFX_ui_sound",0,CalloutVol,0,0,1,1,1
            GaldorReset=true:GaldorQ1 = "fYb"
            CheckPortalRoomShot 6
  End Select
End Sub



'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************


' Used for drop targets
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

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
    Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
    Dim rotxy
    rotxy = RotPoint(ax,ay,angle)
    rax = rotxy(0)+px : ray = rotxy(1)+py
    rotxy = RotPoint(bx,by,angle)
    rbx = rotxy(0)+px : rby = rotxy(1)+py
    rotxy = RotPoint(cx,cy,angle)
    rcx = rotxy(0)+px : rcy = rotxy(1)+py
    rotxy = RotPoint(dx,dy,angle)
    rdx = rotxy(0)+px : rdy = rotxy(1)+py

    InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************




'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

'sub TargetBouncer(aBall,defvalue)
'    dim zMultiplier, vel, vratio
'    if TargetBouncerEnabled <> 0 and aball.z < 30 then
'        debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
'        vel = BallSpeed(aBall)
'        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
'        Select Case Int(Rnd * 6) + 1
'            Case 1: zMultiplier = 0.1*defvalue
'      Case 2: zMultiplier = 0.2*defvalue
'            Case 3: zMultiplier = 0.3*defvalue
'      Case 4: zMultiplier = 0.4*defvalue
'            Case 5: zMultiplier = 0.5*defvalue
'            Case 6: zMultiplier = 0.6*defvalue
'        End Select
'        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
'        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
'        aBall.vely = aBall.velx * vratio
'        debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
'        debug.print "conservation check: " & BallSpeed(aBall)/vel
'    end if
'end sub

sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel
  vel = BallSpeed(aBall)
' debug.print "bounce"
  if aball.z < 30 then
    'debug.print "velz: " & activeball.velz
    Select Case Int(Rnd * 4) + 1
      Case 1: zMultiplier = defvalue+1.1
      Case 2: zMultiplier = defvalue+1.05
      Case 3: zMultiplier = defvalue+0.7
      Case 4: zMultiplier = defvalue+0.3
    End Select
    aBall.velz = aBall.velz * zMultiplier * 1.1
    'debug.print "----> velz: " & activeball.velz
    'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub




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
dim RampBalls(7,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(7)

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

  'Exit adding balls to Rampballs, if more than 6 balls in play
  'dim BOT: BOT = GetBalls
  'if uBound(BOT) > 6 then Exit Sub


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


'//  Calculates the roll volume of the sound based on the ball speed
Dim TempBallVelPlastic
Function VolPlasticMetalRampRoll(ball)
  'VolPlasticMetalRampRoll = RollingOnDiscSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
  TempBallVelPlastic = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/RollingSoundFactor)
  If TempBallVelPlastic = 1 Then TempBallVelPlastic = 0.999
  If TempBallVelPlastic = 0 Then TempBallVelPlastic = 0.001
  'debug.print TempBallVel
  TempBallVelPlastic = Csng(1/(1+(0.275*(((0.75*TempBallVelPlastic)/(1-TempBallVelPlastic))^(-2)))))
  VolPlasticMetalRampRoll = TempBallVelPlastic
End Function

''//  Calculates the roll pitch of the sound based on the ball speed
'Dim TempPitchBallVel
'Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
' 'PitchPlayfieldRoll = BallVel(ball) ^2 * 15
' 'PitchPlayfieldRoll = Csng(BallVel(ball))/50 * 10000
' 'PitchPlayfieldRoll = (1-((Csng(BallVel(ball))/50)^0.2)) * 20000
'
' 'PitchPlayfieldRoll = (2*((Csng(BallVel(ball)))^0.7))/(2+(Csng(BallVel(ball)))) * 16000
' TempPitchBallVel = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/50)
' If TempPitchBallVel = 1 Then TempPitchBallVel = 0.999
' If TempPitchBallVel = 0 Then TempPitchBallVel = 0.001
' TempPitchBallVel = Csng(1/(1+(0.275*(((0.75*TempPitchBallVel)/(1-TempPitchBallVel))^(-2))))) * 10000
' PitchPlayfieldRoll = TempPitchBallVel
'End Function

'//  Calculates the pitch of the sound based on the ball speed.
''//  Used for plastic ramps roll sound
'Function PitchPlasticRamp(ball)
'    PitchPlasticRamp = BallVel(ball) * 20
'End Function



Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
'         PlaySound("RampLoop" & x & "_amp9"), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
'         StopSound("wireloop" & x & "_amp9")
          StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Metal_Wire_BallRoll_" & x)
          PlaySound (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_BallRoll_" & x), 0, (VolPlasticMetalRampRoll(RampBalls(x,0))) * PlasticRampRollSoundFactor * RampRollVolume * VolumeDial * 1.25, AudioPan(RampBalls(x,0)), 0, 0, 1, 0, AudioFade(RampBalls(x,0))

        Else
'         StopSound("RampLoop" & x & "_amp9")
'         PlaySound("wireloop" & x & "_amp9"), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Plastic_BallRoll_" & x)
          PlaySound (Cartridge_Metal_Ramps & "_Ramp_Left_Metal_Wire_BallRoll_" & x), 0, (VolPlasticMetalRampRoll(RampBalls(x,0))) * MetalWireRampRollSoundFactor * RampRollVolume * VolumeDial * 1.25, AudioPan(RampBalls(x,0)), 0, 0, 1, 0, AudioFade(RampBalls(x,0))
'         debug.print "wireramp"
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
'       StopSound("RampLoop" & x & "_amp9")
'       StopSound("wireloop" & x & "_amp9")
        StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Metal_Wire_BallRoll_" & x)
        StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Plastic_BallRoll_" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
'       StopSound("RampLoop" & x & "_amp9")
'       StopSound("wireloop" & x & "_amp9")
        StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Metal_Wire_BallRoll_" & x)
        StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Plastic_BallRoll_" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
'     StopSound("RampLoop" & x & "_amp9")
'     StopSound("wireloop" & x & "_amp9")
      StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Metal_Wire_BallRoll_" & x)
      StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Plastic_BallRoll_" & x)
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


'START of Fleep Audio
'/////////////////////////////  MECHANICAL SOUNDS  '/////////////////////////////

'////////////////////////////////////////////////////////////////////////////////
'////          Mechanical Sounds, by Fleep                                   ////
'////                     Last Updated: January, 2022                        ////
'////////////////////////////////////////////////////////////////////////////////
'
'/////////////////////////////////  CARTRIDGES  /////////////////////////////////
'
'//  Specify which mechanical sound cartridge to use for each group of elements.
'//  Mechanical sounds naming convention: <CARTRIDGE>_<Soundset_Name>
'//
'//  Cartridge name is composed using the following convention:
'//  <TABLE MANUFACTURER ABBREVIATION>_<TABLE NAME ABBREVIATION>_<SOUNDSET REVISION NUMBER>
'//
'//  General Mechanical Sounds Cartridges:
Const Cartridge_Bumpers         = "WS_PBT_REV01" 'Williams Pinbot REV01
Const Cartridge_Slingshots        = "WS_PBT_REV01" 'Williams Pinbot REV01
Const Cartridge_Flippers        = "WS_PBT_REV01" 'Williams Pinbot REV01
Const Cartridge_Kickers         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Kickers2        = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Knocker         = "WS_WHD_REV02" 'Williams Whirlwind Cartridge REV02
Const Cartridge_Relays          = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Trough          = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Rollovers       = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Targets         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Gates         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Spinner         = "SY_TNA_REV01" 'Spooky Total Nuclear Annihilation Cartridge REV01
Const Cartridge_Rubber_Hits       = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Metal_Hits        = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Cabinet_Sounds      = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Apron         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Ball_Roll       = "BY_TOM_REV01" 'Bally Theatre of Magic Cartridge REV01
Const Cartridge_BallBallCollision   = "BY_WDT_REV01" 'Bally WHO Dunnit Cartridge REV01
Const Cartridge_Ball_Drop_Bump      = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Plastic_Ramps     = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Metal_Ramps       = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Table_Specifics     = "SY_TNA_REV01" 'Spooky Total Nuclear Annihilation Cartridge REV01


'////////////////////////////  SOUND SOURCE CREDITS  ////////////////////////////
'//  Special thanks go to the following contributors who have provided audio
'//  footage recordings:
'//
'//  Spooky Total Nuclear Annihilation - WildDogArcade, Ed and Gary
'//  Bally Theatre of Magic - CalleV, nickbuol
'//  Bally WHO Dunnit - Amazaley1
'//  Williams Whirlwind - Blackmoor, wrd1972
'//  Williams Pinbot - major_drain_pinball

'///////////////////////////////  USER PARAMETERS  //////////////////////////////
'
'//  Sounds Parameter with suffix "SoundLevel" can have any value in range [0..1]
'//  Sounds Parameter with suffix "SoundMultiplier" can have any value

'///////////////////////////  SOLENOIDS (COILS) CONFIG  /////////////////////////
'//  FLIPPER COILS:
'//  Flippers in this table: Lower Left Flipper, Lower Right Flipper, Upper Left Fliiper
'//  Flipper Hit Param initialize with FlipperUpSoundLevel
'//  and dynamically modified calculated by ball hits

Dim FlipperUpSoundLevel, FlipperDownSoundLevel
Dim FlipperLeftLowerHitParm, FlipperLeftUpperHitParm, FlipperRightLowerHitParm
Dim FlipperHoldSoundLevel

FlipperUpSoundLevel = 1.0                                   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.05                                  'volume level; range [0, 1]
FlipperLeftLowerHitParm = FlipperUpSoundLevel             'sound helper; not configurable
FlipperLeftUpperHitParm = FlipperUpSoundLevel             'sound helper; not configurable
FlipperRightLowerHitParm = FlipperUpSoundLevel              'sound helper; not configurable
FlipperHoldSoundLevel = 0.3                       'volume level; range [0, 1]

'//  CONTROLLED / SWITCHED COILS:
'//  Slingshots in this table: Main Left, Main Right, Core Left, Core Top, Core Right
'//  Bumpers in this table: Only one
Dim SlingshotSoundFactor, BumperSoundFactor, KnockerSoundLevel
Dim AutoPlungerSoundLevel, ScoopKickerSoundLevel, TroughKickoutSoundLevel, GateCoilHoldSoundLevel
Dim Solenoid_OutholeKicker_SoundLevel, Solenoid_ShooterFeeder_SoundLevel, KickouBallDropSoundLevel, ScoopEnterSoundLevel

AutoPlungerSoundLevel = 1                       'volume level; range [0, 1]
Solenoid_OutholeKicker_SoundLevel = 1
Solenoid_ShooterFeeder_SoundLevel = 1
SlingshotSoundFactor = 0.95                       'volume multiplier
BumperSoundFactor = 0.95                        'volume multiplier
KnockerSoundLevel = 1                           'volume level; range [0, 1]
AutoPlungerSoundLevel = 1                       'volume level; range [0, 1]
ScoopKickerSoundLevel = 1                       'volume level; range [0, 1]
TroughKickoutSoundLevel = 1                       'volume level; range [0, 1]
GateCoilHoldSoundLevel = 0.0025                     'volume level; range [0, 1]
KickouBallDropSoundLevel = 1
ScoopEnterSoundLevel = 0.85

'//  CONTROLLED / SWITCHED COILS:
'//  DropTargets in this table: Drop Target1 (Top), Drop Target2 (Middle), Drop Target3 (Bottom)
Dim DropTargetResetCoilWhenDownSoundLevel, DropTargetResetCoilWhenUpSoundLevel, DropTargetResetCoilHoldSoundLevel
Dim DropTargetKnockDownCoilSoundLevel
Dim DropTargetReleaseSoundLevel, DropTargetHitSoundLevel

DropTargetResetCoilWhenDownSoundLevel = 1               'volume level; range [0, 1]
DropTargetResetCoilWhenUpSoundLevel = 0.5               'volume level; range [0, 1]
DropTargetResetCoilHoldSoundLevel = 0.0025                'volume level; range [0, 1]
DropTargetKnockDownCoilSoundLevel = 0.5                 'volume level; range [0, 1]
DropTargetReleaseSoundLevel = 1                     'volume level; range [0, 1]
DropTargetHitSoundLevel = 1                       'volume level; range [0, 1]


'//  MOTORS:
Dim ShakerSoundLevel, BeaconSoundLevel

ShakerSoundLevel = 1                          'volume level; range [0, 1]
BeaconSoundLevel = 1                          'volume level; range [0, 1]


'// RAMPS:
dim PlasticRampRollSoundFactor, MetalWireRampRollSoundFactor

PlasticRampRollSoundFactor = 0.2
MetalWireRampRollSoundFactor = 0.4

'////////////////////////////  SWITCHES SOUNDS CONFIG  //////////////////////////
Dim StandupTargetSoundFactor, SpinnerSpinSoundLevel, RolloverSoundLevel, OutLaneRolloverSoundLevel

StandupTargetSoundFactor = 0.0005                   'volume multiplier
SpinnerSpinSoundLevel = 0.005                     'volume level; range [0, 1]
RolloverSoundLevel = 0.15                                       'volume level; range [0, 1]
OutLaneRolloverSoundLevel = 0.8


'////////////////////  BALL HITS, BUMPS, DROPS SOUND CONFIG  ////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor
Dim BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftSoundFactor, BallBouncePlayfieldHardSoundFactor
Dim BallBounceAfterPopBumperSoundLevel, BallBounceAfterSlingshotSoundLevel
Dim BallBouncePlayfieldFromMetalShooterLaneSoundLevel
Dim ScoopEntrySoundLevel, ScoopHitSoundLevel
Dim SpinnerHitSoundLevel
Dim Switch_Gate_SoundLevel

RubberStrongSoundFactor = 0.10025                   'volume multiplier
RubberWeakSoundFactor = 0.10025                     'volume multiplier
RubberFlipperSoundFactor = 0.08                     'volume multiplier
BallWithBallCollisionSoundFactor = 3.2                  'volume multiplier
BallBouncePlayfieldSoftSoundFactor = 0.005                'volume multiplier
BallBouncePlayfieldHardSoundFactor = 0.005                'volume multiplier
BallBounceAfterPopBumperSoundLevel = 1                  'volume level; range [0, 1]
BallBounceAfterSlingshotSoundLevel = 1                  'volume level; range [0, 1]
BallBouncePlayfieldFromMetalShooterLaneSoundLevel = 1         'volume level; range [0, 1]
ScoopEntrySoundLevel = 1                        'volume level; range [0, 1]
ScoopHitSoundLevel = 1                          'volume level; range [0, 1]
SpinnerHitSoundLevel = 1                        'volume level; range [0, 1]
Switch_Gate_SoundLevel = 1

'///////////////////////  OTHER PLAYFIELD ELEMENTS CONFIG  //////////////////////
Dim RollingSoundFactor, TroughDrainSoundLevel
Dim BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor, RightPlasticRampHitsSoundLevel
Dim MetalGuideHitSoundLevel, MetalShooterLaneSoundFactor, TroughToShooterLaneOnPlungerSoundLevel, WallImpactSoundFactor
Dim MetalBallGuideOrbitEntranceSoundFactor, MetalBallGuideOrbitRollSoundFactor, MetalImpactSoundFactor, RightPlasticRampEnteranceSoundLevel
dim BlastSoundLevel

RollingSoundFactor = 50                         'volume multiplier
BlastSoundLevel = 0.8
TroughDrainSoundLevel = 0.8                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier
MetalGuideHitSoundLevel = 1                       'volume level; range [0, 1]
MetalShooterLaneSoundFactor = 1                     'volume multiplier
TroughToShooterLaneOnPlungerSoundLevel = 1                'volume level; range [0, 1]
WallImpactSoundFactor = 0.02                      'volume multiplier
MetalImpactSoundFactor = 0.05
MetalBallGuideOrbitEntranceSoundFactor = 1                'volume multiplier
MetalBallGuideOrbitRollSoundFactor = 1                  'volume multiplier
RightPlasticRampEnteranceSoundLevel = 0.1
RightPlasticRampHitsSoundLevel = 1

'///////////////////////////  CABINET SOUND PARAMETERS  /////////////////////////
Dim CoinSoundLevel, StartButtonSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel
Dim NudgeLeftSoundLevel, NudgeRightSoundLevel, NudgeCenterSoundLevel

CoinSoundLevel = 1                            'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                       'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                 'volume level; range [0, 1]
PlungerPullSoundLevel = 1                       'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                         'volume level; range [0, 1]
NudgeRightSoundLevel = 1                        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                       'volume level; range [0, 1]


'//  RELAYS:
'//  Solenoid 16 = Lower Playfield Relay GI Relay (P/N 5580-12145-00) / Backbox GI Relay (P/N 5580-09555-01)
'//  Solenoid 11 = Upper Playfield Relay GI Relay (P/N 5580-12145-00)
'//  Solenoid 12 = Solenoid A/C Select Relay (5580-09555-01)
'//  Fake Solenoid = Flahser Relay

Dim RelayLowerGISoundLevel, RelayUpperGISoundLevel, RelaySolenoidACSelectSoundLevel, RelayFlasherSoundLevel
RelayLowerGISoundLevel = 0.25
RelayUpperGISoundLevel = 0.25
RelaySolenoidACSelectSoundLevel = 0.3
RelayFlasherSoundLevel = 0.005


'////////////////////////////////  SOUND HELPERS  ///////////////////////////////
'//  Helper for top metal ball guide
Dim ballVariableForMetalBallGuideTop
'//  Helpers for metal shooter lane
Dim ballVariableMetalShooterLane
Dim isBallDroppedFromMetalShooterLane

'//  Orbit entrance helpers
Dim OrbitLeft_EnterFromPlayfieldRightEntrance
Dim OrbitRight_EnterFromPlayfieldRightEntrance
Dim OrbitRight_EnterFromShooterLane
Dim StopMetalRollWhenBallRollsSlowFromShooter

'//  Helper for sound toggle when using timers
Dim SoundOn : SoundOn = 1
Dim SoundOff : SoundOff = 0

'//  Helper for Main (Lower) flippers dampened stroke
Dim BallNearLF : BallNearLF = 0
Dim BallNearRF : BallNearRF = 0

Sub TriggerBallNearLF_Hit()
  'Debug.Print "BallNearLF = 1"
  BallNearLF = 1
End Sub

Sub TriggerBallNearLF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearLF = 0
End Sub

Sub TriggerBallNearRF_Hit()
  'Debug.Print "BallNearRF = 1"
  BallNearRF = 1
End Sub

Sub TriggerBallNearRF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearRF = 0
End Sub


'//  Helpers for drop targets
Dim TargetUp : TargetUp = 1
Dim TargetDown : TargetDown = -1



'///////////////////////  SOUND PLAYBACK SUBS / FUNCTIONS  //////////////////////
'//////////////////////  POSITIONAL SOUND PLAYBACK METHODS  /////////////////////
'//  Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
'//  These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
'//  For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
'//  For stereo setup - positional sound playback functions will only pan between left and right channels
'//  For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels
'//
'//  PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
'//  Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 1, AudioFade(tableobj)
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


'//////////////////////  SUPPORTING BALL & SOUND FUNCTIONS  /////////////////////
'Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

'//  Fades between front and back of the table
'//  (for surround systems or 2x2 speakers, etc), depending on the Y position
'//  on the table.
Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  Select Case PositionalSoundPlaybackConfiguration
    Case 1
      AudioFade = 0
    Case 2
      AudioFade = 0
    Case 3
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
  End Select
End Function


Function AudioFadePosY(positiony) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = positiony * 2 / tableheight-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
    AudioFadePosY = Csng(tmp ^10)
    Else
        AudioFadePosY = Csng(-((- tmp) ^10) )
    End If
End Function

'//  Calculates the pan for a tableobj based on the X position on the table.
Function AudioPan(tableobj)
    Dim tmp
  Select Case PositionalSoundPlaybackConfiguration
    Case 1
      AudioPan = 0
    Case 2
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
    Case 3
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
  End Select
End Function

Function AudioPanPosX(positionx) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = positionx * 2 / tablewidth-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
        AudioPanPosX = Csng(tmp ^10)
    Else
        AudioPanPosX = Csng(-((- tmp) ^10) )
    End If
End Function

'//  Calculates the volume of the sound based on the ball speed
Function Vol(ball)
  Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = Csng((ball.velz) ^2)
End Function

'//  Calculates the pitch of the sound based on the ball speed
Function Pitch(ball)
    Pitch = BallVel(ball) * 20
End Function

'//  Calculates the ball speed
Function BallVel(ball)
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Dim TempBallVel
Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  TempBallVel = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/RollingSoundFactor)
  If TempBallVel = 1 Then TempBallVel = 0.999
  If TempBallVel = 0 Then TempBallVel = 0.001
  'debug.print TempBallVel
  TempBallVel = Csng(1/(1+(0.275*(((0.75*TempBallVel)/(1-TempBallVel))^(-2)))))
  VolPlayfieldRoll = TempBallVel
End Function

'//  Calculates the roll pitch of the sound based on the ball speed
Dim TempPitchBallVel
Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  'PitchPlayfieldRoll = BallVel(ball) ^2 * 15
  'PitchPlayfieldRoll = Csng(BallVel(ball))/50 * 10000
  'PitchPlayfieldRoll = (1-((Csng(BallVel(ball))/50)^0.2)) * 20000

  'PitchPlayfieldRoll = (2*((Csng(BallVel(ball)))^0.7))/(2+(Csng(BallVel(ball)))) * 16000
  TempPitchBallVel = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/50)
  If TempPitchBallVel = 1 Then TempPitchBallVel = 0.999
  If TempPitchBallVel = 0 Then TempPitchBallVel = 0.001
  TempPitchBallVel = Csng(1/(1+(0.275*(((0.75*TempPitchBallVel)/(1-TempPitchBallVel))^(-2))))) * 10000
  PitchPlayfieldRoll = TempPitchBallVel
End Function

'//  Sets a random number integer between min and max
Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)
End Function

'//  Sets a random number between min and max
Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min
End Function


'///////////////////////////  PLAY SOUNDS SUBROUTINES  //////////////////////////
'//
'//  These Subroutines implement all mechanical playsounds including timers
'//
'//////////////////////////  GENERAL SOUND SUBROUTINES  /////////////////////////
Sub SoundStartButton()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Start_Button"), StartButtonSoundLevel, StartButtonPosition
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerPullStop()
  StopSound Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Ball_" & Int(Rnd*3)+1), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Empty"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySoundAtLevelStatic ("Nudge_" & Int(Rnd*3)+1), NudgeCenterSoundLevel * VolumeDial, CenterNudgePosition
End Sub


'/////////////////////////  GAME MECH SOUND SUBROUTINES  ////////////////////////
'///////////////////////////////  KNOCKER SOLENOID  /////////////////////////////
Sub KnockerSolenoid()
  PlaySound (Cartridge_Knocker & "_Knocker_Coil"), KnockerSoundLevel
End Sub

Sub KnockerSolenoidDOF(DOFevent, DOFstate)
  PlaySound SoundFXDOF(Cartridge_Knocker & "_Knocker_Coil", DOFevent, DOFstate, DOFKnocker), KnockerSoundLevel
End Sub




'////////////////////////////////////  DRAIN  ///////////////////////////////////
'///////////////////////////////  OUTHOLE SOUNDS  ///////////////////////////////
dim cOutholeKicker : cOutholeKicker = 0
Sub RandomSoundOutholeHit(drainswitch)
  PlaySoundAtLevelStatic (Cartridge_Trough & "_Outhole_Drain_Hit_" & Int(Rnd*4)+1), Solenoid_OutholeKicker_SoundLevel, drainswitch
  cOutholeKicker = 1
End Sub

Sub RandomSoundOutholeKicker()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Outhole_Kicker_" & Int(Rnd*4)+1,DOFContactors), Solenoid_OutholeKicker_SoundLevel, drain
End Sub

'/////////////////////  BALL SHOOTER FEEDER SOLENOID SOUNDS  ////////////////////
Sub RandomSoundShooterFeeder()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Shooter_Feeder_" & Int(Rnd*6)+1,DOFContactors), Solenoid_ShooterFeeder_SoundLevel, ballrelease
End Sub



dim Solenoid_Kickback_SoundLevel : Solenoid_Kickback_SoundLevel = 1
'///////////////////////////  LOCKING KICKER SOLENOID  //////////////////////////
Sub RandomSoundLockingKickerSolenoid()
' debug.print "kick"
  PlaySoundAtLevelStatic (Cartridge_Kickers2 & "_Locking_Kickback_" & Int(Rnd*4)+1), Solenoid_Kickback_SoundLevel, KickbackPosition
End Sub



'/////////////////////////  SLINGSHOT SOLENOID SOUNDS  //////////////////////////
Sub RandomSoundSlingshotMainLeft(sling)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)

  'debug.print "Slingshot volume: " & BallVel(Activeball) / 50 * SlingshotSoundFactor
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Left_" & Int(Rnd*26)+1,DOFContactors), BallVel(Activeball) / 50 * SlingshotSoundFactor, Sling
End Sub

Sub RandomSoundSlingshotMainRight(sling)
  'debug.print "Slingshot volume: " & BallVel(ActiveBall) / 50 * SlingshotSoundFactor
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Right_" & Int(Rnd*25)+1,DOFContactors), BallVel(Activeball) / 50 * SlingshotSoundFactor, Sling
End Sub



Sub RandomSoundSlingshotMainLeftDOF(sling, DOFevent, DOFstate)
  'debug.print "Slingshot volume: " & BallVel(ActiveBall) / 50 * SlingshotSoundFactor
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Slingshots & "_Slingshot_Left_" & Int(Rnd*26)+1, DOFevent, DOFstate,DOFContactors), BallVel(Activeball) / 50 * SlingshotSoundFactor, Sling
End Sub

Sub RandomSoundSlingshotMainRightDOF(sling, DOFevent, DOFstate)
  'debug.print "Slingshot volume: " & BallVel(ActiveBall) / 50 * SlingshotSoundFactor
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Slingshots & "_Slingshot_Right_" & Int(Rnd*25)+1, DOFevent, DOFstate,DOFContactors), BallVel(Activeball) / 50 * SlingshotSoundFactor, Sling
End Sub




'///////////////////////////  BUMPER SOLENOID SOUNDS  ///////////////////////////
Sub RandomSoundJetBumperLeft(Bump)
  'debug.print "Bumper volume: " & BallVel(ActiveBall) / 50 * BumperSoundFactor
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Left_" & Int(Rnd*22)+1,DOFContactors), BallVel(ActiveBall) / 25 * BumperSoundFactor, Bump
End Sub

Sub RandomSoundJetBumperLeftDOF(Bump, DOFevent, DOFstate)
  'debug.print "Bumper volume: " & BallVel(ActiveBall) / 50 * BumperSoundFactor
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Bumpers & "_Jet_Bumper_Left_" & Int(Rnd*22)+1, DOFevent, DOFstate,DOFContactors), BallVel(ActiveBall) / 25 * BumperSoundFactor, Bump
End Sub

Sub RandomSoundJetBumperLow(Bump)
  'debug.print "Bumper volume: " & BallVel(ActiveBall) / 50 * BumperSoundFactor
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Low_" & Int(Rnd*28)+1,DOFContactors), BallVel(ActiveBall) / 25 * BumperSoundFactor, Bump
End Sub

Sub RandomSoundJetBumperLowDOF(Bump, DOFevent, DOFstate)
  'debug.print "Bumper volume: " & BallVel(ActiveBall) / 50 * BumperSoundFactor
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Bumpers & "_Jet_Bumper_Low_" & Int(Rnd*28)+1, DOFevent, DOFstate,DOFContactors), BallVel(ActiveBall) / 25 * BumperSoundFactor, Bump
End Sub

Sub RandomSoundJetBumperUp(Bump)
  'debug.print "Bumper volume: " & BallVel(ActiveBall) / 50 * BumperSoundFactor
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Up_" & Int(Rnd*25)+1,DOFContactors), BallVel(ActiveBall) / 25 * BumperSoundFactor, Bump
End Sub

Sub RandomSoundJetBumperUpDOF(Bump, DOFevent, DOFstate)
  'debug.print "Bumper volume: " & BallVel(ActiveBall) / 50 * BumperSoundFactor
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Bumpers & "_Jet_Bumper_Up_" & Int(Rnd*25)+1, DOFevent, DOFstate,DOFContactors), BallVel(ActiveBall) / 25 * BumperSoundFactor, Bump
End Sub


'///////////////////////  FLIPPER BATS SOUND SUBROUTINES  ///////////////////////
'//////////////////////  FLIPPER BATS SOLENOID CORE SOUND  //////////////////////


Sub RandomSoundFlipperLowerLeftUpFullStrokeDOF(flipper, DOFevent, DOFstate)
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & RndInt(1,10), DOFevent, DOFstate, DOFFlippers), FlipperLeftLowerHitParm, Flipper
End Sub

  'we have no upper flip sounds so we use lever left sounds
    Sub RandomSoundFlipperUpperLeftUpFullStroke(flipper)
      PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & RndInt(1,10),DOFFlippers), FlipperLeftUpperHitParm, Flipper
    End Sub


Sub RandomSoundFlipperLowerLeftUpDampenedStrokeDOF(flipper, DOFevent, DOFstate)
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & RndInt(1,23), DOFevent, DOFstate, DOFFlippers), FlipperLeftLowerHitParm * 1.2, Flipper
End Sub




Sub RandomSoundFlipperLowerRightUpFullStrokeDOF(flipper, DOFevent, DOFstate)
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & RndInt(1,11), DOFevent, DOFstate, DOFFlippers), FlipperRightLowerHitParm, Flipper
End Sub

Sub RandomSoundFlipperLowerRightUpDampenedStrokeDOF(flipper, DOFevent, DOFstate)
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & RndInt(1,23), DOFevent, DOFstate, DOFFlippers), FlipperRightLowerHitParm * 1.2, Flipper
End Sub


Sub RandomSoundFlipperLowerLeftReflipDOF(flipper, DOFevent, DOFstate)
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Flippers & "_Flipper_Lower_Left_Reflip_" & RndInt(1,3), DOFevent, DOFstate, DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

    Sub RandomSoundFlipperUpperLeftReflip(flipper)
      PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Reflip_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    End Sub



Sub RandomSoundFlipperLowerRightReflipDOF(flipper, DOFevent, DOFstate)
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Flippers & "_Flipper_Lower_Right_Reflip_" & RndInt(1,3), DOFevent, DOFstate, DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub




Sub RandomSoundFlipperLowerLeftDownDOF(flipper, DOFevent, DOFstate)
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & RndInt(1,10), DOFevent, DOFstate, DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'
    Sub RandomSoundFlipperUpperLeftDown(flipper)
      PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & RndInt(1,10),DOFFlippers), FlipperDownSoundLevel, Flipper
    End Sub


Sub RandomSoundFlipperLowerRightDownDOF(flipper, DOFevent, DOFstate)
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & RndInt(1,11), DOFevent, DOFstate, DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundLowerLeftQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, LeftFlipper
End Sub

Sub RandomSoundLowerRightQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, RightFlipper
End Sub

Sub StopAnyFlipperLowerLeftUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperUpperLeftDown()
  StopAnyFlipperLowerLeftUp()
End Sub

Sub StopAnyFlipperLowerRightUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 11
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperLowerLeftDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & anyFullDownSound)
  Next
End Sub

Sub StopAnyFlipperLowerRightDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & anyFullDownSound)
  Next
End Sub



'///////////////////////  FLIPPER BATS BALL COLLIDE SOUND  //////////////////////
Sub LeftFlipperCollide(parm)
  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperLeftLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperLeftLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

'    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  RandomSoundRubberFlipper(parm)
End Sub

Sub LeftFlipper1_Collide(parm)  'Note: no live catch added here for upper flipper
    If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperLeftUpperHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperLeftUpperHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperLeftUpperHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If
  RandomSoundRubberFlipper(parm)
End Sub

Sub LeftFlipperMini_Collide(parm) 'Note: no live catch added here for upper flipper

  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperMini_Collide(parm)  'Note: no live catch added here for upper flipper

  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperRightLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperRightLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If
'    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
' debug.print "flipper coll " & parm
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Flipper_Hit_" & Int(Rnd*7)+1), parm / 2 * RubberFlipperSoundFactor
End Sub


'////////////////////////////  SCOOP KICKER SOUNDS  /////////////////////////////
Sub RandomSoundScoopLeftEjectHighVelocity()
' PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Scoop_Left_Eject_High_Velocity_" & Int(Rnd*3)+1), ScoopKickerSoundLevel, LeftScoop
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_Kickout_Cellar_" & Int(Rnd*3)+1), ScoopKickerSoundLevel, LeftScoop
  SoundCellarKickoutBallDrop(LeftScoop)
End Sub

Sub RandomSoundScoopLeftEjectNormalVelocity(prim)
' PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Scoop_Left_Eject_Normal_Velocity_" & Int(Rnd*5)+1), ScoopKickerSoundLevel, prim
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_Kickout_Cellar_" & Int(Rnd*3)+1), ScoopKickerSoundLevel, prim
End Sub

Sub RandomSoundScoopRightEjectHighVelocity()
' PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Scoop_Right_Eject_High_Velocity_" & Int(Rnd*3)+1), ScoopKickerSoundLevel, RightScoop
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_Kickout_Cellar_" & Int(Rnd*3)+1), ScoopKickerSoundLevel, RightScoop
  SoundCellarKickoutBallDrop(RightScoop)
End Sub

Sub RandomSoundScoopRightEjectNormalVelocity()
' PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Scoop_Right_Eject_Normal_Velocity_" & Int(Rnd*5)+1), ScoopKickerSoundLevel, RightScoop
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_Kickout_Cellar_" & Int(Rnd*3)+1), ScoopKickerSoundLevel, RightScoop
End Sub


Sub RandomSoundScoopLeftEjectHighVelocityDOF(ScoopKicker, DOFevent, DOFstate)
' PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Table_Specifics & "_Scoop_Left_Eject_High_Velocity_" & Int(Rnd*3)+1, DOFevent, DOFstate, DOFPulse), ScoopKickerSoundLevel, ScoopKicker
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Kickers & "_Scoop_Kickout_Cellar_" & Int(Rnd*3)+1, DOFevent, DOFstate, DOFPulse), ScoopKickerSoundLevel, ScoopKicker
  SoundCellarKickoutBallDrop(ScoopKicker)
End Sub

Sub RandomSoundScoopLeftEjectNormalVelocityDOF(ScoopKicker, DOFevent, DOFstate)
' PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Table_Specifics & "_Scoop_Left_Eject_Normal_Velocity_" & Int(Rnd*5)+1, DOFevent, DOFstate, DOFPulse), ScoopKickerSoundLevel, ScoopKicker
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Kickers & "_Scoop_Kickout_Cellar_" & Int(Rnd*3)+1, DOFevent, DOFstate, DOFPulse), ScoopKickerSoundLevel, ScoopKicker
End Sub

Sub RandomSoundScoopRightEjectHighVelocityDOF(ScoopKicker, DOFevent, DOFstate)
' PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Table_Specifics & "_Scoop_Right_Eject_High_Velocity_" & Int(Rnd*3)+1, DOFevent, DOFstate, DOFPulse), ScoopKickerSounwdLevel, ScoopKicker
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Kickers & "_Scoop_Kickout_Cellar_" & Int(Rnd*3)+1, DOFevent, DOFstate, DOFPulse), ScoopKickerSoundLevel, ScoopKicker
  SoundCellarKickoutBallDrop(ScoopKicker)
End Sub

Sub RandomSoundScoopRightEjectNormalVelocityDOF(ScoopKicker, DOFevent, DOFstate)
' PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Table_Specifics & "_Scoop_Right_Eject_Normal_Velocity_" & Int(Rnd*5)+1, DOFevent, DOFstate, DOFPulse), ScoopKickerSoundLevel, ScoopKicker
  PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Kickers & "_Scoop_Kickout_Cellar_" & Int(Rnd*3)+1, DOFevent, DOFstate, DOFPulse), ScoopKickerSoundLevel, ScoopKicker
End Sub


'//////////////////////////  SCOOP ENTER / HIT SOUNDS  //////////////////////////
Sub RandomSoundScoopLeftEnter(obj)

  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_Left_Cellar_Enter_" & Int(Rnd*4)+1), ScoopEnterSoundLevel * finalspeed/40, obj
End Sub

Sub RandomSoundScoopRightEnter(obj)

  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_Right_Cellar_Enter_" & Int(Rnd*5)+1), ScoopEnterSoundLevel * finalspeed/40, obj
End Sub


Sub SoundCellarKickoutBallDrop(prim)
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_BallDrop_After_Kickout_" & Int(Rnd*2)+1), KickouBallDropSoundLevel, prim
End Sub




'/////////////////////  DROP TARGET RESET SOLENOID SOUNDS  //////////////////////
Sub SoundDropTarget_Release(prim)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Targets & "_Drop_Target_1Bank_Reset_Up_" & Int(Rnd*8)+1,DOFContactors), DropTargetReleaseSoundLevel, prim
End Sub

Sub RandomSoundDropTargetLeft(prim)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Targets & "_Drop_Target_3Bank_Reset_Up_" & Int(Rnd*6)+1,DOFContactors), DropTargetReleaseSoundLevel, prim
End Sub



'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////
Sub SoundDropTarget(prim)
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_1Bank_Release_Down_" & Int(Rnd*2)+1), DropTargetHitSoundLevel, prim
End Sub

Sub SoundDropTarget3Bank(prim)
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_3Bank_Release_Down_" & Int(Rnd*6)+1), DropTargetHitSoundLevel, prim
End Sub



'///////////////////////////////////  GATES  ////////////////////////////////////
'Sub GateL_Hit()
' PlaySoundAtLevelStatic (Cartridge_Gates & "_Gate_Left_Through"), GateFlapSoundLevel, GateL
'End Sub

Sub GateLTrigger_Hit()
  if activeball.velx > 0 Then
'debug.print "passing left"
'   PlaySoundAtLevelStatic (Cartridge_Gates & "_Gate_Left_Through"), GateFlapSoundLevel, Gate004
    PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_1"), Switch_Gate_SoundLevel * 0.3, Gate004
  else
    if Gate004.collidable Then
'debug.print "Left closed"
'     PlaySoundAtLevelStatic (Cartridge_Gates & "_Gate_Closed_Hit_" & Int(Rnd*8)+1), GateHitSoundLevel, Gate004
      Stopsound Cartridge_Gates & "_Bracket_Gate_1"
      PlaySoundAtLevelStatic (Cartridge_Gates & "_Oneway_Ball_Gate_" & Int(Rnd*3)+1), Switch_Gate_SoundLevel * 0.3, Gate004
    Else
'debug.print "Left open"
'     PlaySoundAtLevelStatic (Cartridge_Gates & "_Gate_Left_Through"), GateFlapSoundLevel, Gate004
      PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_1"), Switch_Gate_SoundLevel * 0.3, Gate004
    end if
  end if
End Sub



Sub GateRTrigger_Hit()
  if activeball.velx < 0 Then
'debug.print "passing right"
'   PlaySoundAtLevelStatic (Cartridge_Gates & "_Gate_Right_Through"), GateFlapSoundLevel, Gate001
    PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_2"), Switch_Gate_SoundLevel * 0.3, Gate001
  else
    if Gate001.collidable Then
'debug.print "Right closed"
'     PlaySoundAtLevelStatic (Cartridge_Gates & "_Gate_Closed_Hit_" & Int(Rnd*8)+1), GateHitSoundLevel, Gate001
      Stopsound Cartridge_Gates & "_Bracket_Gate_2"
      PlaySoundAtLevelStatic (Cartridge_Gates & "_Oneway_Ball_Gate_" & Int(Rnd*3)+1), Switch_Gate_SoundLevel * 0.3, Gate001
    Else
'debug.print "Right open"
'     PlaySoundAtLevelStatic (Cartridge_Gates & "_Gate_Right_Through"), GateFlapSoundLevel, Gate001
      PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_2"), Switch_Gate_SoundLevel * 0.3, Gate001
    end if
  end if
End Sub





'//////////////////////////////////  SHAKER  ////////////////////////////////////
'DOF numbers for shaker intensity
'152 - Shaker Intensity 1
'152 - Shaker Intensity 2
'152 - Shaker Intensity 3

'//  Possible Shaker Configurations (determined by const ShakerMotor in sound options above)
'//  0 = Shaker motor completely disabled (DOF+Sound/SSF)
'//  1 = Shaker motor enabled for Sound/SSF only
'//  2 = Shaker motor enabled for DOF only
'//  3 = Shaker motor enabled for DOF + Sound/SSF

'SoundShakerDOF soundon, dofon
'SoundShakerDOF soundoff, dofoff
'ShakerIntensity = "" : SoundShakerDOF soundon, dofon
'SoundShakerDOF soundoff, dofoff



Sub SoundShakerDOF(toggle, DOFstate)
' debug.print "shaker: " & toggle
  Select Case ShakerMotor
    Case 0
    Case 1
      Select Case toggle
        Case SoundOn
          Select Case ShakerIntensity
            Case "Low"
              PlaySoundAtLevelExistingStaticLoop (Cartridge_Table_Specifics & "_Shaker_Level_01_RampUp_and_Loop"), ShakerSoundLevel, ShakerPosition
            Case "Normal"
              PlaySoundAtLevelExistingStaticLoop (Cartridge_Table_Specifics & "_Shaker_Level_02_RampUp_and_Loop"), ShakerSoundLevel, ShakerPosition
            Case "High"
              PlaySoundAtLevelExistingStaticLoop (Cartridge_Table_Specifics & "_Shaker_Level_03_RampUp_and_Loop"), ShakerSoundLevel, ShakerPosition
          End Select
        Case SoundOff
          Select Case shakerintensity
            Case "Low"
              PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Shaker_Level_01_RampDown_Only"), ShakerSoundLevel, ShakerPosition
              StopSound Cartridge_Table_Specifics & "_Shaker_Level_01_RampUp_and_Loop"
            Case "Normal"
              PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Shaker_Level_02_RampDown_Only"), ShakerSoundLevel, ShakerPosition
              StopSound Cartridge_Table_Specifics & "_Shaker_Level_02_RampUp_and_Loop"
            Case "High"
              PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Shaker_Level_03_RampDown_Only"), ShakerSoundLevel, ShakerPosition
              StopSound Cartridge_Table_Specifics & "_Shaker_Level_03_RampUp_and_Loop"
          End Select
      End Select
    Case 2
'     Select Case toggle
'       Case SoundOn
          Select Case ShakerIntensity
            Case "Low"
              PlaySoundAtLevelExistingStaticLoop SoundFXDOF("", 152, DOFstate, DOFShaker), ShakerSoundLevel, ShakerPosition
            Case "Normal"
              PlaySoundAtLevelExistingStaticLoop SoundFXDOF("", 152, DOFstate, DOFShaker), ShakerSoundLevel, ShakerPosition
            Case "High"
              PlaySoundAtLevelExistingStaticLoop SoundFXDOF("", 152, DOFstate, DOFShaker), ShakerSoundLevel, ShakerPosition
          End Select
'       Case SoundOff
'     End Select
    Case 3
      Select Case toggle
        Case SoundOn
          Select Case ShakerIntensity
            Case "Low"
              PlaySoundAtLevelExistingStaticLoop SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_01_RampUp_and_Loop", 152, DOFstate, DOFShaker), ShakerSoundLevel, ShakerPosition
            Case "Normal"
              PlaySoundAtLevelExistingStaticLoop SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_02_RampUp_and_Loop", 152, DOFstate, DOFShaker), ShakerSoundLevel, ShakerPosition
            Case "High"
              PlaySoundAtLevelExistingStaticLoop SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_03_RampUp_and_Loop", 152, DOFstate, DOFShaker), ShakerSoundLevel, ShakerPosition
          End Select
        Case SoundOff
          Select Case shakerintensity
            Case "Low"
'             PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Shaker_Level_01_RampDown_Only"), ShakerSoundLevel, Drain
              PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_01_RampDown_Only", 152, DOFstate, DOFShaker), ShakerSoundLevel, ShakerPosition
              StopSound Cartridge_Table_Specifics & "_Shaker_Level_01_RampUp_and_Loop"
            Case "Normal"
'             PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Shaker_Level_02_RampDown_Only"), ShakerSoundLevel, Drain
              PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_02_RampDown_Only", 152, DOFstate, DOFShaker), ShakerSoundLevel, ShakerPosition
              StopSound Cartridge_Table_Specifics & "_Shaker_Level_02_RampUp_and_Loop"
            Case "High"
'             PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Shaker_Level_03_RampDown_Only"), ShakerSoundLevel, Drain
              PlaySoundAtLevelStatic SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_03_RampDown_Only", 152, DOFstate, DOFShaker), ShakerSoundLevel, ShakerPosition
              StopSound Cartridge_Table_Specifics & "_Shaker_Level_03_RampUp_and_Loop"
          End Select
      End Select
  End Select
End Sub




'///////////////////////////////////  SPINNER  //////////////////////////////////
Sub SoundSpinner()
  PlaySoundAtLevelStatic (Cartridge_Spinner & "_Spinner_Spin_Loop"), SpinnerSpinSoundLevel, Spinner1
End Sub


'//////////////////////////////  SPINNER EVENTS  ////////////////////////////////
Sub SpinnerTrigger_Hit()
  If activeball.vely < 0 Then
    'Ball rolls up
    PlaySoundAtLevelStatic (Cartridge_Spinner & "_Spinner_Hit_Front_" & Int(Rnd*6)+1), SpinnerHitSoundLevel, Spinner1
  End If
  If activeball.vely > 0 Then
    'Ball rolls down
    PlaySoundAtLevelStatic (Cartridge_Spinner & "_Spinner_Hit_Back_" & Int(Rnd*3)+1), SpinnerHitSoundLevel, Spinner1
  End If
End Sub


'//////////////////////////////  ROLLOVER SOUNDS  ///////////////////////////////
Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_All_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub


Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub


Sub RandomSoundInlanes()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_Outlane_" & Int(Rnd*4)+1), OutLaneRolloverSoundLevel
End Sub

Sub RandomSoundLanes()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_Outlane_" & Int(Rnd*4)+1), OutLaneRolloverSoundLevel
End Sub


Sub Sound_GI_Relay(toggle)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayUpperGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayUpperGISoundLevel, GILowerPosition
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayUpperGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayUpperGISoundLevel, GILowerPosition
  End Select
End Sub


'///////////////////////////////  FLASHERS RELAY  ///////////////////////////////
Sub Sound_Flasher_Relay(toggle, tableobj)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlasherSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlasherSoundLevel, tableobj
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlasherSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlasherSoundLevel, tableobj
  End Select
End Sub


'////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  /////////////////////
'/////////////////////////////  RUBBERS AND POSTS  //////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ///////////////////////////////
Sub HitsRubbers_Hit(idx)
' debug.print "rubber hit"
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 10 then
    RandomSoundRubberWeak()
  End If
End Sub


'/////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  /////////////////////
Sub RandomSoundRubberStrong()
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_10"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub


'///////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  /////////////////////
Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_" & Int(Rnd*8)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub



'////////////////////////////  RAMP ENTRANCE EVENTS  ////////////////////////////
'/////////////////////////  RIGHT RAMP ENTRANCE SOUNDS  /////////////////////////

Sub RRAMPUP_Hit()
  ' Play the right plastic ramp entrance lifter/down sound for each ball
  If BallVel(ActiveBall) > 1 and ActiveBall.VelY < 0 Then
'   debug.print "Ball rolls upwards"
    PlaySoundAtLevelTimerExistingActiveBall (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Ramp_Enter_" & Int(Rnd*4)+1), RightPlasticRampEnteranceSoundLevel, ActiveBall
  End If
End Sub

Sub RRAMPDOWN_Hit()
  ' Play the right plastic ramp entrance lifter/down sound for each ball
  If BallVel(ActiveBall) > 1 and ActiveBall.VelY > 0 Then
'   debug.print "Ball rolls downwards"
    PlaySoundAtLevelTimerExistingActiveBall (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Ramp_RollBack_" & Int(Rnd*2)+1), RightPlasticRampEnteranceSoundLevel, ActiveBall
  End If
End Sub

Sub RRHelper1_Hit()
  If ActiveBall.VelX < 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_5"), RightPlasticRampHitsSoundLevel, RRHelper1 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"
End Sub


'///////////////////////////////  WALL IMPACTS  /////////////////////////////////
Sub RandomSoundWall()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
' debug.print finalspeed
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor * 0.05
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  /////////////////////////////
Sub RandomSoundMetal()
' debug.print "randr metal: " & Vol(ActiveBall)
  Select Case Int(Rnd*20)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_1"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_2"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_3"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_4"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_5"), Vol(ActiveBall) * 0.02 * MetalImpactSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_6"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_7"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_8"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_9"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_10"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 11 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_11"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 12 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_12"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 13 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_13"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 14 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_14"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 15 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_15"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 16 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_16"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 17 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_17"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 18 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_18"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 19 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_19"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 20 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_20"), Vol(ActiveBall) * MetalImpactSoundFactor
  End Select
End Sub

'///////////////////////  METAL WIRE BALL GUIDE HIT SOUNDS  /////////////////////
Sub MetalWireBallGuideHit()
' PlaySoundAtLevelActiveBall (Cartridge_Ball_Guides & "_Wire_Ball_Guide_Rail_" & Int(Rnd*6)+1), Vol(ActiveBall) * MetalGuideHitSoundLevel
  RandomSoundWall
End Sub


'/////////////////////////  METAL GUIDE SIDE HIT SOUNDS  ////////////////////////
Sub MetalGuideSideHit()
' PlaySoundAtLevelActiveBall (Cartridge_Ball_Guides & "_Metal_Ball_Guide_Side_Hit_" & Int(Rnd*14)+1), Vol(ActiveBall) * MetalGuideHitSoundLevel
  RandomSoundWall
End Sub


'/////////////////////////  METAL GUIDE TOP HIT SOUNDS  /////////////////////////
Sub MetalGuideTopHit()
' PlaySoundAtLevelActiveBall (Cartridge_Ball_Guides & "_Metal_Ball_Guide_Top_Hit_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalGuideHitSoundLevel
  RandomSoundMetal
End Sub


'///////////////////////  METAL BALL GUIDE COLLECTIONS  /////////////////////////
Sub Metal_Wire_Ball_Guides_Hit(idx)
'debug.print "Metal_Wire_Ball_Guides_Hit"
' MetalWireBallGuideHit()
  RandomSoundMetal
End Sub

'Sub Metal_Ball_Guides_Hit(idx)
Sub Metals_Hit(idx)

  RandomSoundMetal
End Sub




'/////////////////////////////  METAL SHOOTER LANE  /////////////////////////////
Sub MetalLaneTrigger_Hit()
'debug.print "MetalLaneTrigger_Hit"
  'Ball rolls up entering the metal lane
  Set ballVariableMetalShooterLane = ActiveBall
  Call SoundMetalLane(SoundOn, ballVariableMetalShooterLane)
End Sub

Sub MetalLaneTrigger_unhit()
'debug.print "MetalLaneTrigger_UnHit"
  'Ball exiting the metal lane - upwards or downwards
  Call SoundMetalLane(SoundOff, ballVariableMetalShooterLane)
  PlaySoundAtLevelActiveBall (Cartridge_Table_Specifics & "_Shooter_Lane_Metal_BallStop"), Vol(ActiveBall) * MetalShooterLaneSoundFactor
  If Activeball.vely < 0 Then
    'Ball rolls upwards
'debug.print "ball rolls upwards"
    OrbitRight_EnterFromShooterLane = 1
  End If
End Sub

Sub SoundMetalLane(toggle, ballVariableMetalShooterLane)
  Select Case toggle
    Case SoundOn
      MetalLaneTimer.Interval = 10
      MetalLaneTimer.Enabled = 1
    Case SoundOff
      StopSound Cartridge_Table_Specifics & "_Shooter_Lane_Metal_BallRoll"
  End Select
End Sub

Sub MetalLaneTimer_Timer()
  '4 point polygon that contains a complete metal shooter lane
  'This timer is enabled from SoundMetalLane sub
  If InRect(ballVariableMetalShooterLane.x, ballVariableMetalShooterLane.y, 880, 880, 950, 640, 950, 1320, 880, 1320) Then
    PlaySoundAtLevelTimerExistingActiveBall (Cartridge_Table_Specifics & "_Shooter_Lane_Metal_BallRoll"), MetalShooterLaneSoundFactor * Csng(BallVel(ballVariableMetalShooterLane) / 25 * VolumeDial), ballVariableMetalShooterLane
  Else
    Me.Enabled = 0
  End If
End Sub




'///////////////////////////  BOTTOM ARCH BALL GUIDE  ///////////////////////////
'//////////////////////  BOTTOM ARCH BALL GUIDE - BOUNCES  //////////////////////
Sub Apronwall_Hit (idx)
  RandomSoundApron activeball
End Sub


Sub RandomSoundApron(aball)
  dim finalspeed
    finalspeed=SQR(aball.velx * aball.velx + aball.vely * aball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_1"),  Vol(aball) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_2"),  Vol(aball) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Medium_" & Int(Rnd*3)+1),  Vol(aball) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Soft_" & Int(Rnd*7)+1),  Vol(aball) * FlipperBallGuideSoundFactor
  End if
End Sub



'////////////////////////   STANDUP TARGET HIT SOUNDS  //////////////////////////
'DOF events for the following standup targets are already triggered in the corresponding, specific target hit subs
'108 - RAD Left Standup Target bank
'109 - Grid Targetx,y,z standup target
'114 - Reactor Stand Up targets


Sub Targets_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft Activeball
  Else
    RandomSoundTargetHitWeak()
  End If
  TargetBouncer activeball, 1.4
End Sub

'//////////////////////////  STADNING TARGET HIT SOUNDS  ////////////////////////
Sub RandomSoundTargetHitStrong()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_5",DOFTargets), Vol(ActiveBall) * StandupTargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_6",DOFTargets), Vol(ActiveBall) * StandupTargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_7",DOFTargets), Vol(ActiveBall) * StandupTargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_8",DOFTargets), Vol(ActiveBall) * StandupTargetSoundFactor
  End Select
End Sub

Sub RandomSoundTargetHitWeak()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_1",DOFTargets), Vol(ActiveBall) * StandupTargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_2",DOFTargets), Vol(ActiveBall) * StandupTargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_3",DOFTargets), Vol(ActiveBall) * StandupTargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_4",DOFTargets), Vol(ActiveBall) * StandupTargetSoundFactor
  End Select
End Sub



'/////////////////////////////  BALL BOUNCE SOUNDS  /////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aball)
' debug.print "soft"
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_2"), Volz(aball) * BallBouncePlayfieldSoftSoundFactor, aBall
    Case 2 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_12"), Volz(aball) * BallBouncePlayfieldSoftSoundFactor, aBall
    Case 3 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_14"), Volz(aball) * BallBouncePlayfieldSoftSoundFactor, aBall
    Case 4 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_18"), Volz(aball) * BallBouncePlayfieldSoftSoundFactor, aBall
    Case 5 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_19"), Volz(aball) * BallBouncePlayfieldSoftSoundFactor, aBall
    Case 6 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_20"), Volz(aball) * BallBouncePlayfieldSoftSoundFactor, aBall
    Case 7 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_21"), Volz(aball) * BallBouncePlayfieldSoftSoundFactor, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aball)
' debug.print "hard"
  Select Case Int(Rnd*12)+1
    Case 1 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_1"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 2 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_3"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 3 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_7"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 4 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_8"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 5 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_9"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 6 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_11"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 7 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_13"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 8 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_15"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 9 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_16"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 10 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_17"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 11 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_22"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
    Case 12 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_23"), Volz(aball) * BallBouncePlayfieldHardSoundFactor, aBall
  End Select
End Sub




'////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  'debug.print "Ball Collide Volume: " & Csng(velocity) / 500 * BallWithBallCollisionSoundFactor * VolumeDial
' if abs(ball1.vely) < 12 And abs(ball2.vely) < 12 And InRect(ball1.x, ball1.y, 246,830,446,830,446,1030,246,1030) then
''    debug.print "exit " & velocity
'   exit sub  'don't rattle the locked balls
' end if
  if WizardPhase = 4 then exit sub  'removing collision sounds for galdor battle
  PlaySound (Cartridge_BallBallCollision & "_BallBall_Collide_" & Int(Rnd*7)+1), 0, Csng(velocity) / 500 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub



'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////



'////////////////////////////////////////////////////////////////////////////////
'////          Mechanical Sounds Options, by Fleep                           ////
'////////////////////////////////////////////////////////////////////////////////
'
'//////////////////////////  MECHANICAL SOUNDS OPTIONS  /////////////////////////
'//  This section allows to set various general sound options for the mechanical sounds.
'//  For the entire sound system scripts see mechanical sounds block down below in the project.
'
'////////////////////////////  GENERAL SOUND OPTIONS  ///////////////////////////
'
'//  PositionalSoundPlaybackConfiguration:
'//  Specifies the sound playback configuration. Options:
'//  1 = Mono
'//  2 = Stereo (Only L+R channels)
'//  3 = Surround (Full surround sound support for SSF - Front L + Front R + Rear L + Rear R channels)
Const PositionalSoundPlaybackConfiguration = 3
'
'
'//  VolumeDial:
'//  VolumeDial is the actual global volume multiplier for the mechanical sounds.
'//  Values smaller than 1 will decrease mechanical sounds volume.
'//  Recommended values should be no greater than 1.
'//  Default value is 0.8 and is optimal.
'Const VolumeDial = 1.0
'
'
'//  PlayfieldRollVolumeDial:
'//  PlayfieldRollVolumeDial is a constant volume multiplier for the playfield rolling sound.
'//  Default value should be 1 which will guarantee a proper carefully calculated dynamic volume changes profile.
'//  Any values different than the default will impact the volume level and the dynamic voume changes profile.
Const PlayfieldRollVolumeDial = 1.0
'
'//  RelaysPosition:
'//  1 = Relays positioned with power board (Provides sound spread from the left and right back channels)
'//  2 = Relays positioned with GI strings (Provides sound spread from left/right/front/rear surround channels)
Const RelaysPosition = 1
'
''//  ShakerMotor:
''//  This setting will determine how will the shaker motor to be used during game play. Options:
''//  0 = Shaker motor completely disabled (DOF+Sound/SSF)
''//  1 = Shaker motor enabled for Sound/SSF only
''//  2 = Shaker motor enabled for DOF only
''//  3 = Shaker motor enabled for DOF + Sound/SSF
'Const ShakerMotor = 3
''
''
''//  ShakerIntensity:
''//  This setting allows to set the overall intensity of the shaker motor (applicable for sound/SSF). Options:
''//  "Low"
''//  "Normal"
''//  "High
'Const ShakerIntensity = "Low"
'
'
'////////////////////////////////////////////////////////////////////////////////
'////          End of Fleep Mechanical Sounds Options                        ////
'////////////////////////////////////////////////////////////////////////////////




'
''///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
'Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor,
dim WireformAntiRebountRailSoundFactor
'dim BlastSoundLevel
'
'BlastSoundLevel = 0.8
'DrainSoundLevel = 0.8                            'volume level; range [0, 1]
'BallReleaseSoundLevel = 1                        'volume level; range [0, 1]
'BottomArchBallGuideSoundFactor = 0.2                 'volume multiplier; must not be zero
'FlipperBallGuideSoundFactor = 0.015                    'volume multiplier; must not be zero
WireformAntiRebountRailSoundFactor = 0.04               'volume multiplier; must not be zero?
'

'End Sub
'
''/////////////////////////////  METAL - EVENTS  ////////////////////////////
'


Sub ShooterDiverter_collide(idx)
' RandomSoundMetal
  MetalGuideSideHit
End Sub
'
'
'/////////////////////////////  WIREFORM ANTI-REBOUNT RAILS  ////////////////////////////
Sub RandomSoundWireformAntiRebountRail()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 2 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelActiveBall ("TOM_Wireform_Anti_Rebound_Rail_3"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("TOM_Wireform_Anti_Rebound_Rail_4"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall ("TOM_Wireform_Anti_Rebound_Rail_5"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 4 : PlaySoundAtLevelActiveBall ("TOM_Wireform_Anti_Rebound_Rail_6"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 5 : PlaySoundAtLevelActiveBall ("TOM_Wireform_Anti_Rebound_Rail_7"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
    End Select
  End if
  If finalspeed < 2 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelActiveBall ("TOM_Wireform_Anti_Rebound_Rail_1"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("TOM_Wireform_Anti_Rebound_Rail_2"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall ("TOM_Wireform_Anti_Rebound_Rail_8"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
    End Select
  End if
End Sub

Sub Wall17_hit() : RandomSoundWireformAntiRebountRail() : End Sub
Sub Wall22_hit() : RandomSoundWireformAntiRebountRail() : End Sub



'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************






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

FlInitBumper 1, "white"
FlInitBumper 2, "white"
FlInitBumper 3, "white"


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


Sub FlReinitBumper(nr, col)
  FlBumperActive(nr) = True
  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1 : FlBumperFadeTarget(nr) = 1.1: FlBumperColor(nr) = col
  ' set the color for the two VPX lights
  select case col
    Case "red"
      FlBumperSmallLight(nr).color = RGB(255,4,0) : FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
      FlBumperBigLight(nr).color = RGB(255,32,0) : FlBumperBigLight(nr).colorfull = RGB(255,32,0)
      FlBumperHighlight(nr).color = RGB(64,255,0)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
      FlBumperSmallLight(nr).TransmissionScale = 0
      MaterialColor "bumpertopmat" & nr, RGB(255,120,50)
    Case "blue"
      FlBumperBigLight(nr).color = RGB(32,80,255) : FlBumperBigLight(nr).colorfull = RGB(32,80,255)
      FlBumperSmallLight(nr).color = RGB(0,80,255) : FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperHighlight(nr).color = RGB(255,16,8)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      MaterialColor "bumpertopmat" & nr, RGB(110,170,255)
    Case "green"
      FlBumperSmallLight(nr).color = RGB(8,255,8) : FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
      FlBumperBigLight(nr).color = RGB(32,255,32) : FlBumperBigLight(nr).colorfull = RGB(32,255,32)
      FlBumperHighlight(nr).color = RGB(255,32,255)
      FlBumperSmallLight(nr).TransmissionScale = 0.005
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      MaterialColor "bumpertopmat" & nr, RGB(140,255,90)
    Case "orange"
      FlBumperHighlight(nr).color = RGB(255,130,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).color = RGB(255,130,0) : FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
      FlBumperBigLight(nr).color = RGB(255,190,8) : FlBumperBigLight(nr).colorfull = RGB(255,190,8)
      MaterialColor "bumpertopmat" & nr, RGB(255,160, 50)
    Case "white"
      FlBumperBigLight(nr).color = RGB(255,230,190) : FlBumperBigLight(nr).colorfull = RGB(255,230,190)
      FlBumperHighlight(nr).color = RGB(255,180,100) :
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
      MaterialColor "bumpertopmat" & nr, RGB(255,235,220)
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
      MaterialColor "bumpertopmat" & nr, RGB(255,235,100)
    Case "purple"
      FlBumperBigLight(nr).color = RGB(80,32,255) : FlBumperBigLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).color = RGB(80,32,255) : FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).TransmissionScale = 0 :
      FlBumperHighlight(nr).color = RGB(32,64,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      MaterialColor "bumpertopmat" & nr, RGB(180,100,255)
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
      FlBumperTop(nr).BlendDisableLighting = 7 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z +0.97 * Z^3)
      Flbumperbiglight(nr).intensity = 45 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 10000 * (Z^3) / (0.5 + DNA90)
            'MaterialColor "bumpertopmat" & nr, RGB(110,170,255)

    Case "green"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(16 + 16 * sin(Z*3.14),255,16 + 16 * sin(Z*3.14)), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
      FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 6000 * (Z^3) / (1 + DNA90)
            'MaterialColor "bumpertopmat" & nr, RGB(140,255,90)

    Case "red"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 16 - 11*Z + 16 * sin(Z*3.14),0), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
      'MaterialColor "bumpertopmat" & nr, RGB(255,120,50)

    Case "orange"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 100 - 22*z  + 16 * sin(Z*3.14),Z*32), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (1 + DNA90)
      'MaterialColor "bumpertopmat" & nr, RGB(255,160 + Z*50, 50)

    Case "white"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
      FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 14 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255,255 - 20*Z,255-65*Z) : FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20*Z,255-65*Z)
      'MaterialColor "bumpertopmat" & nr, RGB(255,235 - z*36,220 - Z*90)

    Case "blacklight"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 1, RGB(30-27*Z^0.03,30-28*Z^0.01, 255), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z^3
      Flbumperbiglight(nr).intensity = 40 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255-240*(Z^0.1),255 - 240*(Z^0.1),255) : FlBumperSmallLight(nr).colorfull = RGB(255-200*z,255 - 200*Z,255)
      MaterialColor "bumpertopmat" & nr, RGB(255-190*Z,235 - z*180,220 + 35*Z)

    Case "yellow"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 180 + 40*z, 48* Z), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      'MaterialColor "bumpertopmat" & nr, RGB(255,235, (24 - 24 * z)+100)

    Case "purple" :
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(128-118*Z - 32 * sin(Z*3.14), 32-26*Z ,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 15  + 200 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z +0.97 * Z^3)
      Flbumperbiglight(nr).intensity = 50 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (0.5 + DNA90)
      'MaterialColor "bumpertopmat" & nr, RGB(180,100,255)
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



' #####################################
' ###### FLUPPER DOMES2.2 #####
' #####################################

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherOffBrightness, FlasherBloomIntensity, FlasherBakesIntensity

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.2   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.4   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
FlasherBloomIntensity = 0.2   ' *** lower this, if the blooms / hazes are too bright (i.e. 0.1) ***
FlasherBakesIntensity = 0.3   ' *** lower this, if the bakes are too bright (i.e. 0.1)      ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objlight(20)
Dim FlashOffDL : FlashOffDL = FlasherOffBrightness

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "purple"
InitFlasher 2, "white"
InitFlasher 3, "purple"
InitFlasher 4, "red"
InitFlasher 5, "green"
InitFlasher 6, "purple"
InitFlasher 7, "purple"

LightFlash1a.IntensityScale = 0
LightFlash1b.IntensityScale = 0
LightFlash3a.IntensityScale = 0
LightFlash3b.IntensityScale = 0
LightFlash4a.IntensityScale = 0
LightFlash4b.IntensityScale = 0
LightFlash5b.IntensityScale = 0
LightFlash6a.IntensityScale = 0
LightFlash6b.IntensityScale = 0
LightFlash6b1d.IntensityScale = 0
LightFlash6b2d.IntensityScale = 0
LightFlash6b3d.IntensityScale = 0


LightFlash7a.IntensityScale = 0
LightFlash7b.IntensityScale = 0



'InitFlasher 7, "green" : InitFlasher 8, "red"
'InitFlasher 9, "green" : InitFlasher 10, "red" : InitFlasher 11, "white"
' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 4,17 : RotateFlasher 5,0 : RotateFlasher 6,90
'RotateFlasher 7,0 : RotateFlasher 8,0
'RotateFlasher 9,-45 : RotateFlasher 10,90 : RotateFlasher 11,90
'RotateFlasher 9,-45 : RotateFlasher 10,90 : RotateFlasher 11,90

Sub InitFlasher(nr, col)
  Dim RedDomeColor : RedDomeColor = RGB(250,15,0)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
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
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0

  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If

  FlasherPL_BS_RedFL.color = RedDomeColor
  if CabinetMode = 0 then
    FlasherPL_RS_redFL.color = RedDomeColor
    FlasherPL_LS_redFL.color = RedDomeColor
  else
    FlasherPL_RS_redFL_CM.color = RedDomeColor
    FlasherPL_LS_redFL_CM.color = RedDomeColor
  end if


  FlasherPFRedFL.color = RedDomeColor

End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

dim flash1off : flash1off = 0
dim flash2off : flash2off = 0
dim flash3off : flash3off = 0
dim flash4off : flash4off = 0
dim flash5off : flash5off = 0
dim flash6off : flash6off = 0
dim flash7off : flash7off = 0

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then
    objflasher(nr).TimerEnabled = True
    objflasher(nr).visible = 1
    objlit(nr).visible = 1
    select case nr
      Case 1:
        FlasherPL_BS_PurpleLFL.visible = 1
        if CabinetMode = 0 then
          FlasherPL_RS_PurpleLFL.visible = 1
          FlasherPL_LS_PurpleLFL.visible = 1
        Else
          FlasherPL_RS_PurpleLFL_CM.visible = 1
          FlasherPL_LS_PurpleLFL_CM.visible = 1
        end if
        FlasherPFPurpleLFL.visible = 1
        if FlasherBloomsEnabled then FlasherBloomPL.visible = 1
        If VRRoom > 0 Then
          VRBGFL4_1.visible = 1
          VRBGFL4_2.visible = 1
          VRBGFL4_3.visible = 1
          VRBGFL4_4.visible = 1
          VRBGFL4_5.visible = 1
        End If
      Case 2:
        FlasherPL_BS_WhiteFL.visible = 1
        if CabinetMode = 0 then
          FlasherPL_LS_WhiteFL.visible = 1
          FlasherPL_RS_WhiteFL.visible = 1
        else
          FlasherPL_LS_WhiteFL_CM.visible = 1
          FlasherPL_RS_WhiteFL_CM.visible = 1
        end if
        FlasherPFWhiteFL.visible = 1
        if FlasherBloomsEnabled then FlasherBloom2.visible = 1
        If VRRoom > 0 Then
          VRBGFL6_1.visible = 1
          VRBGFL6_2.visible = 1
          VRBGFL6_3.visible = 1
          VRBGFL6_4.visible = 1
          VRBGFL6_5.visible = 1
        End If
      Case 3:
        FlasherPL_BS_PurpleLFL.visible = 1
        if CabinetMode = 0 then
          FlasherPL_RS_PurpleLFL.visible = 1
          FlasherPL_LS_PurpleLFL.visible = 1
        else
          FlasherPL_RS_PurpleLFL_CM.visible = 1
          FlasherPL_LS_PurpleLFL_CM.visible = 1
        end if
        FlasherPFPurpleLFL.visible = 1
        if FlasherBloomsEnabled then FlasherBloomPL.visible = 1
        If VRRoom > 0 Then
          VRBGFL4_1.visible = 1
          VRBGFL4_2.visible = 1
          VRBGFL4_3.visible = 1
          VRBGFL4_4.visible = 1
          VRBGFL4_5.visible = 1
        End If
      Case 4:
        FlasherPL_BS_RedFL.visible = 1
        if CabinetMode = 0 then
          FlasherPL_RS_redFL.visible = 1
          FlasherPL_LS_redFL.visible = 1
        else
          FlasherPL_RS_redFL_CM.visible = 1
          FlasherPL_LS_redFL_CM.visible = 1
        end if
        FlasherPFRedFL.visible = 1
        if FlasherBloomsEnabled then FlasherBloom4.visible = 1
        If VRRoom > 0 Then
          VRBGFL8_1.visible = 1
          VRBGFL8_2.visible = 1
          VRBGFL8_3.visible = 1
          VRBGFL8_4.visible = 1
          VRBGFL8_5.visible = 1
          VRBGFL8_7.visible = 1
          VRBGFL8_8.visible = 1
        End If
      Case 5:
        FlasherPL_BS_GreenFL.visible = 1
        if CabinetMode = 0 then
          FlasherPL_LS_GreenFL.visible = 1
          FlasherPL_RS_GreenFL.visible = 1
        else
          FlasherPL_LS_GreenFL_CM.visible = 1
          FlasherPL_RS_GreenFL_CM.visible = 1
        end if
        FlasherPFGreenFL.visible = 1
        if FlasherBloomsEnabled then FlasherBloom5.visible = 1
      Case 6:
        FlasherPL_BS_PurpleRFL.visible = 1
        if CabinetMode = 0 then
          FlasherPL_RS_PurpleRFL.visible = 1
          FlasherPL_LS_PurpleRFL.visible = 1
        else
          FlasherPL_RS_PurpleRFL_CM.visible = 1
          FlasherPL_LS_PurpleRFL_CM.visible = 1
        end if
        FlasherPFPurpleRFL.visible = 1
        if FlasherBloomsEnabled then FlasherBloomPR.visible = 1
        If VRRoom > 0 Then
          VRBGFL5_1.visible = 1
          VRBGFL5_2.visible = 1
          VRBGFL5_3.visible = 1
          VRBGFL5_4.visible = 1
          VRBGFL5_5.visible = 1
        End If
      Case 7:
        FlasherPL_BS_PurpleRFL.visible = 1
        if CabinetMode = 0 then
          FlasherPL_RS_PurpleRFL.visible = 1
          FlasherPL_LS_PurpleRFL.visible = 1
        else
          FlasherPL_RS_PurpleRFL_CM.visible = 1
          FlasherPL_LS_PurpleRFL_CM.visible = 1
        end if
        FlasherPFPurpleRFL.visible = 1
        if FlasherBloomsEnabled then FlasherBloomPR.visible = 1
        If VRRoom > 0 Then
          VRBGFL5_1.visible = 1
          VRBGFL5_2.visible = 1
          VRBGFL5_3.visible = 1
          VRBGFL5_4.visible = 1
          VRBGFL5_5.visible = 1
        End If
    end Select

  End If

  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlashOffDL + 1 * ObjLevel(nr)^3
  'objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^0.6
  'objlit(nr).BlendDisableLighting = 100 * ObjLevel(nr)^2
  objlit(nr).BlendDisableLighting = 30 * ObjLevel(nr)^0.6
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  'ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  ObjLevel(nr) = ObjLevel(nr) * 0.8 - 0.01


  select case nr
    Case 1:
      LightFlash1a.IntensityScale = 2 * ObjLevel(nr)^1
      LightFlash1b.IntensityScale = 2 * ObjLevel(nr)^2
      FlasherPL_BS_PurpleLFL.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2) 'average of 1 and 3
      if CabinetMode = 0 then
        FlasherPL_RS_PurpleLFL.opacity = 150 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2)^2 'average of 1 and 3
        FlasherPL_LS_PurpleLFL.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2) 'average of 1 and 3
      Else
        FlasherPL_RS_PurpleLFL_CM.opacity = 150 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2)^2 'average of 1 and 3
        FlasherPL_LS_PurpleLFL_CM.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2) 'average of 1 and 3
      end if
      FlasherPFPurpleLFL.opacity = 70 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2) 'average of 1 and 3
      FlasherBloomPL.opacity = 100 * FlasherBloomIntensity * ((ObjLevel(1)+ObjLevel(3))/2)^2 'average of 1 and 3
      If VRRoom > 0 Then
        VRBGFL4_1.opacity = 100 * ObjLevel(nr)^2
        VRBGFL4_2.opacity = 100 * ObjLevel(nr)^2
        VRBGFL4_3.opacity = 100 * ObjLevel(nr)^2
        VRBGFL4_4.opacity = 100 * ObjLevel(nr)^2
        VRBGFL4_5.opacity = 100 * ObjLevel(nr)^3
      End If

      if ObjLevel(nr) < 0.6 And flash1off = 0 then Call Sound_Flasher_Relay(0, LightRelayPosition) : flash1off = 1
    Case 2:
      FlasherPL_BS_WhiteFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)
      if CabinetMode = 0 then
        FlasherPL_LS_WhiteFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)^2
        FlasherPL_RS_WhiteFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)
      else
        FlasherPL_LS_WhiteFL_CM.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)^2
        FlasherPL_RS_WhiteFL_CM.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)
      end if
      FlasherPFWhiteFL.opacity = 90 * FlasherBakesIntensity * ObjLevel(nr)
      FlasherBloom2.opacity = 50  * FlasherBloomIntensity * ObjLevel(nr)^2
      if ObjLevel(nr) < 0.6 And flash2off = 0 then Call Sound_Flasher_Relay(0, LightRelayPosition) : flash2off = 1
      If VRRoom > 0 Then
        VRBGFL6_1.opacity = 100 * ObjLevel(nr)^2
        VRBGFL6_2.opacity = 100 * ObjLevel(nr)^2
        VRBGFL6_3.opacity = 100 * ObjLevel(nr)^2
        VRBGFL6_4.opacity = 100 * ObjLevel(nr)^2
        VRBGFL6_5.opacity = 100 * ObjLevel(nr)^2
      End If
    Case 3:
      LightFlash3a.IntensityScale = 2 * ObjLevel(nr)^1
      LightFlash3b.IntensityScale = 2 * ObjLevel(nr)^2
      FlasherPL_BS_PurpleLFL.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2) 'average of 1 and 3
      if CabinetMode = 0 then
        FlasherPL_RS_PurpleLFL.opacity = 150 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2)^2 'average of 1 and 3
        FlasherPL_LS_PurpleLFL.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2) 'average of 1 and 3
      else
        FlasherPL_RS_PurpleLFL_CM.opacity = 150 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2)^2 'average of 1 and 3
        FlasherPL_LS_PurpleLFL_CM.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2) 'average of 1 and 3
      end if
      FlasherPFPurpleLFL.opacity = 70 * FlasherBakesIntensity * ((ObjLevel(1)+ObjLevel(3))/2) 'average of 1 and 3
      FlasherBloomPL.opacity = 100 * FlasherBloomIntensity * ((ObjLevel(1)+ObjLevel(3))/2)^2 'average of 1 and 3
      If VRRoom > 0 Then
        VRBGFL4_1.opacity = 100 * ObjLevel(nr)^2
        VRBGFL4_2.opacity = 100 * ObjLevel(nr)^2
        VRBGFL4_3.opacity = 100 * ObjLevel(nr)^2
        VRBGFL4_4.opacity = 100 * ObjLevel(nr)^2
        VRBGFL4_5.opacity = 100 * ObjLevel(nr)^3
      End If
      if ObjLevel(nr) < 0.6 And flash3off = 0 then Call Sound_Flasher_Relay(0, LightRelayPosition) : flash3off = 1
    Case 4:
      LightFlash4a.IntensityScale = 2 * ObjLevel(nr)
      LightFlash4b.IntensityScale = 2 * ObjLevel(nr)^2
      FlasherPL_BS_RedFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)^2
      if CabinetMode = 0 then
        FlasherPL_RS_redFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)
        FlasherPL_LS_redFL.opacity = 180 * FlasherBakesIntensity * ObjLevel(nr)^2
      else
        FlasherPL_RS_redFL_CM.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)
        FlasherPL_LS_redFL_CM.opacity = 180 * FlasherBakesIntensity * ObjLevel(nr)^2
      end if
      FlasherPFRedFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)
      FlasherBloom4.opacity = 100 * FlasherBloomIntensity * ObjLevel(nr)^2
      If VRRoom > 0 Then
        VRBGFL8_1.opacity = 100 * ObjLevel(nr)^2
        VRBGFL8_2.opacity = 100 * ObjLevel(nr)^2
        VRBGFL8_3.opacity = 100 * ObjLevel(nr)^2
        VRBGFL8_4.opacity = 100 * ObjLevel(nr)^2
        VRBGFL8_5.opacity = 100 * ObjLevel(nr)^2
        VRBGFL8_7.opacity = 100 * ObjLevel(nr)^2
        VRBGFL8_8.opacity = 100 * ObjLevel(nr)^2
      End If
      if ObjLevel(nr) < 0.6 And flash4off = 0 then flash4off = 1 ': Call Sound_Flasher_Relay(0, LightRelayPosition) 'ssftodo
    Case 5:
      LightFlash5b.IntensityScale = 2 * ObjLevel(nr)
      FlasherPL_BS_GreenFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)^3
      if CabinetMode = 0 then
        FlasherPL_LS_GreenFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)
        FlasherPL_RS_GreenFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)^2
      else
        FlasherPL_LS_GreenFL_CM.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)
        FlasherPL_RS_GreenFL_CM.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)^2
      end if
      FlasherPFGreenFL.opacity = 100 * FlasherBakesIntensity * ObjLevel(nr)
      FlasherBloom5.opacity = 100 * FlasherBloomIntensity * ObjLevel(nr)^2
      if ObjLevel(nr) < 0.6 And flash5off = 0 then Call Sound_Flasher_Relay(0, LightRelayPosition) : flash5off = 1
    Case 6:
      FlasherPL_BS_PurpleRFL.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2) 'average of 1 and 3
      if CabinetMode = 0 then
        FlasherPL_RS_PurpleRFL.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2) 'average of 1 and 3
        FlasherPL_LS_PurpleRFL.opacity = 150 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2)^2 'average of 1 and 3
      else
        FlasherPL_RS_PurpleRFL_CM.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2) 'average of 1 and 3
        FlasherPL_LS_PurpleRFL_CM.opacity = 150 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2)^2 'average of 1 and 3
      end if

      FlasherPFPurpleRFL.opacity = 70 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2) 'average of 1 and 3
      FlasherBloomPR.opacity = 100 * FlasherBloomIntensity * ((ObjLevel(6)+ObjLevel(7))/2)^2 'average of 1 and 3
      If VRRoom > 0 Then
        VRBGFL5_1.opacity = 100 * ObjLevel(nr)^2
        VRBGFL5_2.opacity = 100 * ObjLevel(nr)^2
        VRBGFL5_3.opacity = 100 * ObjLevel(nr)^2
        VRBGFL5_4.opacity = 100 * ObjLevel(nr)^2
        VRBGFL5_5.opacity = 100 * ObjLevel(nr)^3
      End If
      LightFlash6a.IntensityScale = 2 * ObjLevel(nr)^1
      LightFlash6b.IntensityScale = 2 * ObjLevel(nr)^2
      LightFlash6b1d.IntensityScale = 2 * ObjLevel(nr)^2
      LightFlash6b2d.IntensityScale = 2 * ObjLevel(nr)^2
      LightFlash6b3d.IntensityScale = 2 * ObjLevel(nr)^2
      if ObjLevel(nr) < 0.6 And flash6off = 0 then Call Sound_Flasher_Relay(0, LightRelayPosition) : flash6off = 1
    Case 7:
      FlasherPL_BS_PurpleRFL.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2) 'average of 1 and 3
      if CabinetMode = 0 then
        FlasherPL_RS_PurpleRFL.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2) 'average of 1 and 3
        FlasherPL_LS_PurpleRFL.opacity = 150 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2)^3 'average of 1 and 3
      else
        FlasherPL_RS_PurpleRFL.opacity = 100 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2) 'average of 1 and 3
        FlasherPL_LS_PurpleRFL.opacity = 150 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2)^2 'average of 1 and 3
      end if
      FlasherPFPurpleRFL.opacity = 70 * FlasherBakesIntensity * ((ObjLevel(6)+ObjLevel(7))/2) 'average of 1 and 3
      FlasherBloomPR.opacity = 100 * FlasherBloomIntensity * ((ObjLevel(6)+ObjLevel(7))/2)^2 'average of 1 and 3
      If VRRoom > 0 Then
        VRBGFL5_1.opacity = 100 * ObjLevel(nr)^2
        VRBGFL5_2.opacity = 100 * ObjLevel(nr)^2
        VRBGFL5_3.opacity = 100 * ObjLevel(nr)^2
        VRBGFL5_4.opacity = 100 * ObjLevel(nr)^2
        VRBGFL5_5.opacity = 100 * ObjLevel(nr)^3
      End If
      LightFlash7a.IntensityScale = 2 * ObjLevel(nr)^1
      LightFlash7b.IntensityScale = 2 * ObjLevel(nr)^2
      if ObjLevel(nr) < 0.6 And flash7off = 0 then Call Sound_Flasher_Relay(0, LightRelayPosition) : flash7off = 1
  end Select

  If ObjLevel(nr) < 0 Then
    objflasher(nr).TimerEnabled = False
    objflasher(nr).visible = 0
    objlit(nr).visible = 0
    select case nr
      Case 1:
        if ObjLevel(3) < 0 then
          FlasherPL_BS_PurpleLFL.visible = 0
          if CabinetMode = 0 then
            FlasherPL_RS_PurpleLFL.visible = 0
            FlasherPL_LS_PurpleLFL.visible = 0
          else
            FlasherPL_RS_PurpleLFL_CM.visible = 0
            FlasherPL_LS_PurpleLFL_CM.visible = 0
          end if
          FlasherPFPurpleLFL.visible = 0
          FlasherBloomPL.visible = 0
          If VRRoom > 0 Then
            VRBGFL4_1.visible = 0
            VRBGFL4_2.visible = 0
            VRBGFL4_3.visible = 0
            VRBGFL4_4.visible = 0
            VRBGFL4_5.visible = 0
          End If

        end if
      Case 2:
        FlasherPL_BS_WhiteFL.visible = 0
        if CabinetMode = 0 then
          FlasherPL_LS_WhiteFL.visible = 0
          FlasherPL_RS_WhiteFL.visible = 0
        else
          FlasherPL_LS_WhiteFL_CM.visible = 0
          FlasherPL_RS_WhiteFL_CM.visible = 0
        end if
        FlasherPFWhiteFL.visible = 0
        FlasherBloom2.visible = 0
        If VRRoom > 0 Then
          VRBGFL6_1.visible = 0
          VRBGFL6_2.visible = 0
          VRBGFL6_3.visible = 0
          VRBGFL6_4.visible = 0
          VRBGFL6_5.visible = 0
        End If
      Case 3:
        if ObjLevel(1) < 0 then
          FlasherPL_BS_PurpleLFL.visible = 0
          if CabinetMode = 0 then
            FlasherPL_RS_PurpleLFL.visible = 0
            FlasherPL_LS_PurpleLFL.visible = 0
          else
            FlasherPL_RS_PurpleLFL_CM.visible = 0
            FlasherPL_LS_PurpleLFL_CM.visible = 0
          end if
          FlasherPFPurpleLFL.visible = 0
          FlasherBloomPL.visible = 0
          If VRRoom > 0 Then
            VRBGFL4_1.visible = 0
            VRBGFL4_2.visible = 0
            VRBGFL4_3.visible = 0
            VRBGFL4_4.visible = 0
            VRBGFL4_5.visible = 0
          End If
        end if
      Case 4:
        FlasherPL_BS_RedFL.visible = 0
        if CabinetMode = 0 then
          FlasherPL_RS_redFL.visible = 0
          FlasherPL_LS_redFL.visible = 0
        else
          FlasherPL_RS_redFL_CM.visible = 0
          FlasherPL_LS_redFL_CM.visible = 0
        end if
        FlasherPFRedFL.visible = 0
        FlasherBloom4.visible = 0
        If VRRoom > 0 Then
          VRBGFL8_1.visible = 0
          VRBGFL8_2.visible = 0
          VRBGFL8_3.visible = 0
          VRBGFL8_4.visible = 0
          VRBGFL8_5.visible = 0
          VRBGFL8_7.visible = 0
          VRBGFL8_8.visible = 0
        End If
      Case 5:
        FlasherPFGreenFL.visible = 0
        FlasherBloom5.visible = 0
        FlasherPL_BS_GreenFL.visible = 0
        if CabinetMode = 0 then
          FlasherPL_LS_GreenFL.visible = 0
          FlasherPL_RS_GreenFL.visible = 0
        else
          FlasherPL_LS_GreenFL_CM.visible = 0
          FlasherPL_RS_GreenFL_CM.visible = 0
        end if
      Case 6:
        if ObjLevel(7) < 0 then
          FlasherPL_BS_PurpleRFL.visible = 0
          if CabinetMode = 0 then
            FlasherPL_RS_PurpleRFL.visible = 0
            FlasherPL_LS_PurpleRFL.visible = 0
          else
            FlasherPL_RS_PurpleRFL_CM.visible = 0
            FlasherPL_LS_PurpleRFL_CM.visible = 0
          end if
          FlasherPFPurpleRFL.visible = 0
          FlasherBloomPR.visible = 0
          If VRRoom > 0 Then
            VRBGFL5_1.visible = 0
            VRBGFL5_2.visible = 0
            VRBGFL5_3.visible = 0
            VRBGFL5_4.visible = 0
            VRBGFL5_5.visible = 0
          End If
        end if
      Case 7:
        if ObjLevel(6) < 0 then
          FlasherPL_BS_PurpleRFL.visible = 0
          if CabinetMode = 0 then
            FlasherPL_RS_PurpleRFL.visible = 0
            FlasherPL_LS_PurpleRFL.visible = 0
          else
            FlasherPL_RS_PurpleRFL_CM.visible = 0
            FlasherPL_LS_PurpleRFL_CM.visible = 0
          end if
          FlasherPFPurpleRFL.visible = 0
          FlasherBloomPR.visible = 0
          If VRRoom > 0 Then
            VRBGFL5_1.visible = 0
            VRBGFL5_2.visible = 0
            VRBGFL5_3.visible = 0
            VRBGFL5_4.visible = 0
            VRBGFL5_5.visible = 0
          End If
        end if
    end Select
  End If
End Sub

'sub Sound_Flasher_Relay(aaa,vvv)
'end sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub
Sub FlasherFlash7_Timer() : FlashFlasher(7) : End Sub


 Sub Flash1(Enabled)
  If Enabled Then
    Objlevel(1) = 1 : FlasherFlash1_Timer
    if flash1off = 1 then Call Sound_Flasher_Relay(1, LightRelayPosition)
    flash1off = 0
    cLightFlash1.state = 0
    DOF 115,2
  End If
 End Sub

 Sub Flash2(Enabled)
  If Enabled Then
    Objlevel(2) = 1 : FlasherFlash2_Timer
    if flash2off = 1 then Call Sound_Flasher_Relay(1, LightRelayPosition)
    flash2off = 0
    cLightFlash2.state = 0
    DOF 116,2
  End If
 End Sub

 Sub Flash3(Enabled)
  If Enabled Then
    Objlevel(3) = 1 : FlasherFlash3_Timer
    'if flash3off = 1 then Call Sound_Flasher_Relay(1, LightRelayPosition)
    if flash3off = 1 then Call Sound_Flasher_Relay(1, LightRelayPosition)
    flash3off = 0
    cLightFlash3.state = 0
    DOF 117,2
  End If
 End Sub


 Sub Flash4(Enabled)
  If Enabled Then
'   debug.print "FL4 value: " & Enabled
    if enabled = True Then
'     debug.print "boolean"
      Objlevel(4) = 1
    Else
'     debug.print "modulated"
      Objlevel(4) = Enabled
    end if
    FlasherFlash4_Timer

    if Enabled = 1 or Enabled = True then
      if flash4off = 1 then Call Sound_Flasher_Relay(1, LightRelayPosition)
      flash4off = 0
    end if
    cLightFlash4.state = 0
    DOF 118,2
  Else
'   debug.print "FL4 OFF"
  End If
 End Sub

 Sub Flash5(Enabled)
  If Enabled Then
    'debug.print "5 ON"
    Objlevel(5) = 1 : FlasherFlash5_Timer
    if flash5off = 1 then Call Sound_Flasher_Relay(1, LightRelayPosition)
    flash5off = 0
    cLightFlash5.state = 0
    DOF 119,2
  End If
 End Sub

 Sub Flash6(Enabled)
  If Enabled Then
    Objlevel(6) = 1 : FlasherFlash6_Timer
    if flash6off = 1 then Call Sound_Flasher_Relay(1, LightRelayPosition)
    flash6off = 0
    cLightFlash6.state = 0
    DOF 120,2
  End If
 End Sub

 Sub Flash7(Enabled)
  If Enabled Then
    Objlevel(7) = 1 : FlasherFlash7_Timer
    if flash7off = 1 then Call Sound_Flasher_Relay(1, LightRelayPosition)
    flash7off = 0
    cLightFlash7.state = 0
    DOF 121,2
  End If
 End Sub




'***************** Dust ***********************

Const MaxDustOpacity = 0.85  'make this a 1.0 for a lot of dust!

DustLayer1.blenddisablelighting = 22
DustLayer3.blenddisablelighting = 22

DustLayer1.visible = 0
DustLayer3.visible = 0

'DustLayer1.visible=1

'Function PI()
' PI = 4*Atn(1)
'End Function

dim langle : langle = 0
dim dlangle : dlangle = 0
dim dradius, xstart, ystart
dim dustopacity : dustopacity = 0
dim dustfadedir : dustfadedir = 0

dradius= 50

'xstart = 476
'ystart = 1132
xstart = 0
ystart = 0

sub ShowDust(enabled)
  If enabled Then
    DustLayer1.visible = True
    DustLayer3.visible = True
    dustopacity = 0
    dustfadedir = 1
    DustTimer.enabled = True
  Else
    dustfadedir = -1
  End If
end sub



sub DustTimer_timer()
  ' Fade in or out the dust
  If dustfadedir <> 0 Then
    dustopacity = dustopacity + 0.01*dustfadedir
    If dustopacity >= MaxDustOpacity Then
      dustfadedir = 0
      dustopacity = MaxDustOpacity
    ElseIf dustopacity <= 0 Then
      dustfadedir = 0
      dustopacity = 0
      DustTimer.enabled = False
      DustLayer1.visible = False
      DustLayer3.visible = False
    End If
    UpdateMaterial "Dust1",0,0,0,0,0,0,dustopacity,RGB(221,71,225),0,0,False,True,0,0,0,0
    UpdateMaterial "Dust3",0,0,0,0,0,0,dustopacity,RGB(192,103,103),0,0,False,True,0,0,0,0
  End If

  langle = langle + 0.1
  dlangle = dlangle + 1

  DustLayer1.x = xstart - (dradius*cos(radians(langle)))
  DustLayer1.y = ystart - (dradius*sin(radians(langle)))

  DustLayer3.x = xstart - (dradius*cos(radians(langle+190)))
  DustLayer3.y = ystart - (dradius*sin(radians(langle+190)))

  DustLayer1.blenddisablelighting = 7 + 6 * sin(radians(dlangle+90))
  DustLayer3.blenddisablelighting = 7 + 6 * sin(radians(dlangle+290))

  if langle > 360 Then
    langle = 0
  end if
  if dlangle > 360 Then
    dlangle = 0
  end if
end sub




'******************************************************
'****  LAMPZ by nFozzy
'******************************************************

Const UsingROM = False      ' Change this to true if you are using Lampz with on ROM based table

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
InitLampsNF              ' Setup lamp assignments
LampTimer.Interval = 20
LampTimer.Enabled = 1


Sub LampTimer_Timer()
  If UsingROM Then
    dim x, chglamp
    If b2son then chglamp = Controller.ChangedLamps
    If Not IsEmpty(chglamp) Then
      For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
        Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
      next
    End If
  Else  'apophis - Use the InPlayState of the 1st light in the Lampz.obj array to set the Lampz.state
    dim idx : for idx = 0 to uBound(Lampz.Obj)
      if Lampz.IsLight(idx) then
        if IsArray(Lampz.obj(idx)) then
          dim tmp : tmp = Lampz.obj(idx)
          Lampz.state(idx) = tmp(0).GetInPlayStateBool
          'debug.print tmp(0).name & " " &  tmp(0).GetInPlayStateBool & " " & tmp(0).IntensityScale  & vbnewline
        Else
          Lampz.state(idx) = Lampz.obj(idx).GetInPlayStateBool
          'debug.print Lampz.obj(idx).name & " " &  Lampz.obj(idx).GetInPlayStateBool & " " & Lampz.obj(idx).IntensityScale  & vbnewline
        end if
      end if
    Next
  End If
  Lampz.Update1 'update (fading logic only)
End Sub

dim FrameTime, InitFrameTime : InitFrameTime = 0
LampTimer2.Interval = -1
LampTimer2.Enabled = True
Sub LampTimer2_Timer()
  FrameTime = gametime - InitFrameTime : InitFrameTime = gametime 'Count frametime. Unused atm?
  Lampz.Update 'updates on frametime (Object updates only)
End Sub

Function FlashLevelToIndex(Input, MaxSize)
  FlashLevelToIndex = cInt(MaxSize * Input)
End Function

'***Material Swap***
'Fade material for green, red, yellow colored Bulb prims
Sub FadeMaterialColoredBulb(pri, group, ByVal aLvl) 'cp's script
  ' if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  Select case FlashLevelToIndex(aLvl, 3)
    Case 0:pri.Material = group(0) 'Off
    Case 1:pri.Material = group(1) 'Fading...
    Case 2:pri.Material = group(2) 'Fading...
    Case 3:pri.Material = group(3) 'Full
  End Select
  'if tb.text <> pri.image then tb.text = pri.image : 'debug.print pri.image end If 'debug
  pri.blenddisablelighting = aLvl * 1 'Intensity Adjustment
End Sub


'Fade material for red, yellow colored bulb Filiment prims
Sub FadeMaterialColoredFiliment(pri, group, ByVal aLvl) 'cp's script
  ' if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  Select case FlashLevelToIndex(aLvl, 3)
    Case 0:pri.Material = group(0) 'Off
    Case 1:pri.Material = group(1) 'Fading...
    Case 2:pri.Material = group(2) 'Fading...
    Case 3:pri.Material = group(3) 'Full
  End Select
  'if tb.text <> pri.image then tb.text = pri.image : 'debug.print pri.image end If 'debug
  pri.blenddisablelighting = aLvl * 50  'Intensity Adjustment
End Sub

dim insertDLMult: insertDLMult = 1

Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity * insertDLMult * (10 * Lampz.lvl(84) + 1) 'pupillatency increases insert brighness
End Sub

Sub DisableLightingNormal(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity  'no pupil latency
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity * insertDLMult
End Sub

Sub DisableLightingMinMix(pri, DLMin, DLMax, ByVal aLvl)  'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = (aLvl * (DLMax - DLMin)) + DLMin
End Sub

Sub DisableLightingPortals(pri, DLintensity, ByVal aLvl)
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)
  pri.blenddisablelighting = aLvl * DLintensity
  UpdateMaterial pri.material,0,0,0,0,0,0,aLvl,RGB(255,80,40),0,0,False,True,0,0,0,0
End Sub

Sub DisableLightingPortalsBack(pri, DLintensity, ByVal aLvl)
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)
  pri.blenddisablelighting = aLvl * DLintensity
  UpdateMaterial pri.material,0,0,0,0,0,0,1,RGB(255*aLvl,80*aLvl,40*aLvl),0,0,False,True,0,0,0,0
  'UpdateMaterial pri.material,0,0,0,0,0,0,aLvl,RGB(255,80,40),0,0,False,True,0,0,0,0
End Sub

sub FadeTracySkin(ByVal aLvl)
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)
  'tracyskin 121,108,57
  'clear coat 109,102,61
  'new 224,20,168
  'UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  'OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive, float elasticity, float elasticityFalloff, float friction, float scatterAngle)
  UpdateMaterial "tracyskin",0.25,0.8,0,0,0,0,1,RGB(16*aLvl+208,131-111*aLvl,132*aLvl+36),RGB(103*aLvl+121,108-88*aLvl,111*aLvl+57),RGB(103*aLvl+121,108-88*aLvl,111*aLvl+57),True,False,0,0,0,0
end sub


Sub InitLampsNF()

  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

  'Adjust fading speeds (1 / full MS fading time)
  dim x

  for x = 0 to 110 : Lampz.FadeSpeedUp(x) = 1/2 : Lampz.FadeSpeedDown(x) = 1/10 : next
  for x = 111 to 140 : Lampz.FadeSpeedUp(x) = 1/3: Lampz.FadeSpeedDown(x) = 1/12 : next 'slow lamps for gi testing

  Lampz.FadeSpeedUp(81) = 1/20
  Lampz.FadeSpeedDown(81) = 1/60

  Lampz.FadeSpeedUp(84) = 1/20
  Lampz.FadeSpeedDown(84) = 1/200

  Lampz.FadeSpeedUp(85) = 1/10
  Lampz.FadeSpeedDown(85) = 1/100

  'Lampz Assignments
  '  In a ROM based table, the lamp ID is used to set the state of the Lampz objects

  'MassAssign is the way to do assignments. It'll create arrays automatically / append objects to existing arrays
  'If not using a ROM, then the first light in the object array should be an invisible control light (in this example they
    'are named starting with "lc"). Then, lights should always be controlled via these control lights, including
    'when using Light Sequencers.


  Lampz.MassAssign(6)= cLightMission1
  Lampz.Callback(6) = "DisableLightingNormal L4, 5,"

  Lampz.MassAssign(7)= cLightMission2
  Lampz.Callback(7) = "DisableLightingNormal L3, 5,"

  Lampz.MassAssign(8)= cLightMission3
  Lampz.Callback(8) = "DisableLightingNormal L2, 5,"

  Lampz.MassAssign(9)= cLightMission4
  Lampz.Callback(9) = "DisableLightingNormal L1, 5,"

  Lampz.MassAssign(10)= cLightMission5
  Lampz.Callback(10) = "DisableLightingNormal R4, 5,"

  Lampz.MassAssign(11)= cLightMission6
  Lampz.Callback(11) = "DisableLightingNormal R3, 5,"

  Lampz.MassAssign(12)= cLightMission7
  Lampz.Callback(12) = "DisableLightingNormal R2, 5,"

  Lampz.MassAssign(13)= cLightShootAgain1
  Lampz.MassAssign(13)= LightShootAgain1
  Lampz.MassAssign(13)= LightShootAgain1b
  Lampz.Callback(13) = "DisableLighting pLightShootAgain1, 200,"

  Lampz.MassAssign(14)= cLightShootAgain2
  Lampz.MassAssign(14)= LightShootAgain2
  Lampz.MassAssign(14)= LightShootAgain2b
  Lampz.Callback(14) = "DisableLighting pLightShootAgain2, 200,"

  Lampz.MassAssign(15)= cLightMimaMB
  Lampz.MassAssign(15)= LightMimaMB
  Lampz.MassAssign(15)= LightMimaMBb
  Lampz.Callback(15) = "DisableLighting pLightMimaMB, 200,"

  Lampz.MassAssign(16)= cLightCoreyMB
  Lampz.MassAssign(16)= LightCoreyMB
  Lampz.MassAssign(16)= LightCoreyMBb
  Lampz.Callback(16) = "DisableLighting pLightCoreyMB, 200,"

  Lampz.MassAssign(17)= cLightTracyMB
  Lampz.MassAssign(17)= LightTracyMB
  Lampz.MassAssign(17)= LightTracyMBb
  Lampz.Callback(17) = "DisableLighting pLightTracyMB, 200,"

  Lampz.MassAssign(18)= cLightLeftOutlane
  Lampz.MassAssign(18)= LightLeftOutlane
  Lampz.MassAssign(18)= LightLeftOutlaneB
  Lampz.MassAssign(18)= LightLeftOutlaneC
  Lampz.MassAssign(18)= LightLeftOutlaneR
  Lampz.Callback(18) = "DisableLighting pLightLeftOutlane, 100,"

  Lampz.MassAssign(19)= cLightLeftInlane
  Lampz.MassAssign(19)= LightLeftInlane
  Lampz.MassAssign(19)= LightLeftInlaneB
  Lampz.MassAssign(19)= LightLeftInlaneC
  Lampz.MassAssign(19)= LightLeftInlaneR
  Lampz.Callback(19) = "DisableLighting pLightLeftInlane, 100,"

  Lampz.MassAssign(20)= cLightRightInlane
  Lampz.MassAssign(20)= LightRightInlane
  Lampz.MassAssign(20)= LightRightInlaneb
  Lampz.MassAssign(20)= LightRightInlaneC
  Lampz.Callback(20) = "DisableLighting pLightRightInlane, 100,"

  Lampz.MassAssign(21)= cLightRightOutlane
  Lampz.MassAssign(21)= LightRightOutlane
  Lampz.MassAssign(21)= LightRightOutlaneb
  Lampz.MassAssign(21)= LightRightOutlaneC
  Lampz.Callback(21) = "DisableLighting pLightRightOutlane, 100,"

  Lampz.MassAssign(22)= cLightMystery1
  Lampz.MassAssign(22)= LightMystery1
  Lampz.MassAssign(22)= LightMystery1B
  Lampz.MassAssign(22)= LightMystery1C
  Lampz.Callback(22) = "DisableLighting pLightMystery1, 100,"

  Lampz.MassAssign(23)= cLightMystery2
  Lampz.MassAssign(23)= LightMystery2
  Lampz.MassAssign(23)= LightMystery2B
  Lampz.MassAssign(23)= LightMystery2C
  Lampz.Callback(23) = "DisableLighting pLightMystery2, 100,"

  Lampz.MassAssign(24)= cLightMystery3
  Lampz.MassAssign(24)= LightMystery3
  Lampz.MassAssign(24)= LightMystery3B
  Lampz.MassAssign(24)= LightMystery3C
  Lampz.Callback(24) = "DisableLighting pLightMystery3, 100,"

  Lampz.MassAssign(25)= cLightScavengeScoop
  Lampz.MassAssign(25)= LightScavengeScoop
  Lampz.MassAssign(25)= LightScavengeScoopB
  Lampz.Callback(25) = "DisableLighting pLightScavengeScoop, 150,"

  Lampz.MassAssign(26)= cLightMissionStart1
  Lampz.MassAssign(26)= LightMissionStart1
  Lampz.MassAssign(26)= LightMissionStart1B
  Lampz.Callback(26) = "DisableLighting pLightMissionStart1, 200,"

  'Lampz.MassAssign(27)= cLightMissionStart2
  'Lampz.MassAssign(27)= LightMissionStart2
  'Lampz.Callback(27) = "DisableLighting pLightMissionStart2, 200,"

  Lampz.MassAssign(28)= cLightLoadLazer
  Lampz.MassAssign(28)= LightLoadLazer
  Lampz.MassAssign(28)= LightLoadLazerB
  Lampz.Callback(28) = "DisableLighting pLightLoadLazer, 200,"

  Lampz.MassAssign(29)= cLightTracy1
  Lampz.MassAssign(29)= LightTracy1
  Lampz.MassAssign(29)= LightTracy1B
  Lampz.Callback(29) = "DisableLighting pLightTracy1, 200,"

  Lampz.MassAssign(30)= cLightTracy2
  Lampz.MassAssign(30)= LightTracy2
  Lampz.MassAssign(30)= LightTracy2B
  Lampz.Callback(30) = "DisableLighting pLightTracy2, 200,"

  Lampz.MassAssign(31)= cLightTracy3
  Lampz.MassAssign(31)= LightTracy3
  Lampz.MassAssign(31)= LightTracy3B
  Lampz.Callback(31) = "DisableLighting pLightTracy3, 200,"

  Lampz.MassAssign(32)= cLightTracy4
  Lampz.MassAssign(32)= LightTracy4
  Lampz.MassAssign(32)= LightTracy4B
  Lampz.Callback(32) = "DisableLighting pLightTracy4, 200,"

  Lampz.MassAssign(33)= cLightTracy5
  Lampz.MassAssign(33)= LightTracy5
  Lampz.MassAssign(33)= LightTracy5B
  Lampz.MassAssign(33)= LightTracy5C
  Lampz.Callback(33) = "DisableLighting pLightTracy5, 200,"

  Lampz.MassAssign(34)= cLightCorey1
  Lampz.MassAssign(34)= LightCorey1
  Lampz.MassAssign(34)= LightCorey1B

  Lampz.MassAssign(35)= cLightCorey2
  Lampz.MassAssign(35)= LightCorey2
  Lampz.MassAssign(35)= LightCorey2B

  Lampz.MassAssign(36)= cLightCorey3
  Lampz.MassAssign(36)= LightCorey3
  Lampz.MassAssign(36)= LightCorey3B

  Lampz.MassAssign(37)= cLightCorey4
  Lampz.MassAssign(37)= LightCorey4
  Lampz.MassAssign(37)= LightCorey4B

  Lampz.MassAssign(38)= cLightCorey5
  Lampz.MassAssign(38)= LightCorey5
  Lampz.MassAssign(38)= LightCorey5B
  Lampz.MassAssign(38)= LightCorey5C

  Lampz.MassAssign(39)= cLightShot1
  Lampz.MassAssign(39)= LightShot1
  Lampz.MassAssign(39)= LightShot1B
  Lampz.Callback(39) = "DisableLighting pLightShot1, 200,"

  Lampz.MassAssign(40)= cLightShot2
  Lampz.MassAssign(40)= LightShot2
  Lampz.MassAssign(40)= LightShot2B
  Lampz.Callback(40) = "DisableLighting pLightShot2, 200,"

  Lampz.MassAssign(41)= cLightShot3
  Lampz.MassAssign(41)= LightShot3
  Lampz.MassAssign(41)= LightShot3B
  Lampz.Callback(41) = "DisableLighting pLightShot3, 200,"

  Lampz.MassAssign(42)= cLightShot4
  Lampz.MassAssign(42)= LightShot4
  Lampz.MassAssign(42)= LightShot4B
  Lampz.Callback(42) = "DisableLighting pLightShot4, 200,"

  Lampz.MassAssign(43)= cLightShot5
  Lampz.MassAssign(43)= LightShot5
  Lampz.MassAssign(43)= LightShot5B
  Lampz.MassAssign(43)= LightShot5C
  Lampz.Callback(43) = "DisableLighting pLightShot5, 200,"

  Lampz.MassAssign(44)= cLightShot6
  Lampz.Callback(44) = "DisableLightingPortalsBack pLightPortalWall001, 50,"
  Lampz.Callback(44) = "DisableLightingPortalsBack reflPortalRight, 10,"

  Lampz.MassAssign(45)= cLightMimaOrb1
  Lampz.MassAssign(45)= LightMimaOrb1
  Lampz.MassAssign(45)= LightMimaOrb1B
  Lampz.Callback(45) = "DisableLighting pLightMimaOrb1, 50,"

  Lampz.MassAssign(46)= cLightMimaOrb2
  Lampz.MassAssign(46)= LightMimaOrb2
  Lampz.MassAssign(46)= LightMimaOrb2B
  Lampz.Callback(46) = "DisableLighting pLightMimaOrb2, 50,"

  Lampz.MassAssign(47)= cLightTopLeftLane
  Lampz.MassAssign(47)= LightTopLeftLane
  Lampz.MassAssign(47)= LightTopLeftLaneB
  Lampz.MassAssign(47)= LightTopLeftLaneC
  Lampz.Callback(47) = "DisableLighting pLightTopLeftLane, 100,"

  Lampz.MassAssign(48)= cLightTopRightLane
  Lampz.MassAssign(48)= LightTopRightLane
  Lampz.MassAssign(48)= LightTopRightLaneB
  Lampz.MassAssign(48)= LightTopRightLaneC
  Lampz.Callback(48) = "DisableLighting pLightTopRightLane, 100,"

  Lampz.MassAssign(49)= cLightVascan1
  Lampz.MassAssign(49)= LightVascan1
  Lampz.Callback(49) = "DisableLighting pLightVascan1On, 200,"

  Lampz.MassAssign(50)= cLightVascan2
  Lampz.MassAssign(50)= LightVascan2
  Lampz.Callback(50) = "DisableLighting pLightVascan2On, 200,"

  Lampz.MassAssign(51)= cLightVascan3
  Lampz.MassAssign(51)= LightVascan3
  Lampz.Callback(51) = "DisableLighting pLightVascan3On, 200,"

  Lampz.MassAssign(52)= cLightVascan4
  Lampz.MassAssign(52)= LightVascan4
  Lampz.Callback(52) = "DisableLighting pLightVascan4On, 200,"

  Lampz.MassAssign(53)= cLightVascan5
  Lampz.MassAssign(53)= LightVascan5
  Lampz.Callback(53) = "DisableLighting pLightVascan5On, 200,"

  Lampz.MassAssign(54)= cLightVascan6
  Lampz.MassAssign(54)= LightVascan6
  Lampz.Callback(54) = "DisableLighting pLightVascan6On, 200,"

  Lampz.MassAssign(55)= cLightUnderPFShot1
  Lampz.MassAssign(55)= LightUnderPFShot1
  Lampz.Callback(55) = "DisableLighting pLightUnderPFShot1, 200,"

  Lampz.MassAssign(56)= cLightUnderPFShot2
  Lampz.MassAssign(56)= LightUnderPFShot2
  Lampz.Callback(56) = "DisableLighting pLightUnderPFShot2, 200,"

  Lampz.MassAssign(57)= cLightCollectExtraBall
  Lampz.MassAssign(57)= LightCollectExtraBall
  Lampz.MassAssign(57)= LightCollectExtraBallB
  Lampz.Callback(57) = "DisableLighting pLightCollectExtraBall, 150,"

  Lampz.MassAssign(58)= cLightKB1
  Lampz.MassAssign(58)= LightKB1
  Lampz.MassAssign(58)= LightKB1B
  Lampz.MassAssign(58)= LightKB1r
  Lampz.Callback(58) = "DisableLighting pLightKB1, 150,"

  Lampz.MassAssign(59)= cLightKB2
  Lampz.MassAssign(59)= LightKB2
  Lampz.MassAssign(59)= LightKB2B
  Lampz.MassAssign(59)= LightKB2r
  Lampz.Callback(59) = "DisableLighting pLightKB2, 150,"

  Lampz.MassAssign(60)= cLightBumperHit
  Lampz.MassAssign(60)= LightBumperHit
  Lampz.MassAssign(60)= LightBumperHitB
  Lampz.Callback(60) = "DisableLighting pLightBumperHit, 200,"

  Lampz.MassAssign(61)= cLightFlash1
  Lampz.Callback(61) = "Flash1 "

  Lampz.MassAssign(62)= cLightFlash2
  Lampz.Callback(62) = "Flash2 "

  Lampz.MassAssign(63)= cLightFlash3
  Lampz.Callback(63) = "Flash3 "

  Lampz.MassAssign(64)= cLightFlash4
  Lampz.Callback(64) = "Flash4 "

  Lampz.MassAssign(65)= cLightFlash5
  Lampz.Callback(65) = "Flash5 "

  Lampz.MassAssign(66)= cLightFlash6
  Lampz.Callback(66) = "Flash6 "

  Lampz.MassAssign(67)= cLightFlash7
  Lampz.Callback(67) = "Flash7 "

  Lampz.MassAssign(68)= cLightFlashUndership
  Lampz.Callback(68) = "FlashUndership "

  Lampz.MassAssign(69)= cLightTurbo
  Lampz.MassAssign(69)= LightTurbo
  Lampz.MassAssign(69)= LightTurboB
  Lampz.MassAssign(69)= LightTurboGlow
  Lampz.Callback(69) = "DisableLighting pLightTurbo, 300,"

  Lampz.MassAssign(70)= cLightMimaTarget
  Lampz.MassAssign(70)= LightMimaTarget
  Lampz.MassAssign(70)= LightMimaTargetB
  Lampz.Callback(70) = "DisableLighting pLightMimaTarget, 300,"

  Lampz.MassAssign(71)= cLightPortalWall001
  Lampz.Callback(71) = "DisableLightingPortalsBack pLightPortalWall001, 50,"
  Lampz.Callback(71) = "DisableLightingPortalsBack reflPortalRight, 10,"

  Lampz.MassAssign(72)= cLightPortalWall002
  Lampz.Callback(72) = "DisableLightingPortalsBack pLightPortalWall002, 50,"

  Lampz.MassAssign(73)= cLightPortalWall003
  Lampz.Callback(73) = "DisableLightingPortalsBack pLightPortalWall003, 50,"
  Lampz.Callback(73) = "DisableLightingPortalsBack reflPortalLeft, 10,"

  Lampz.MassAssign(74)= cLightMissionWizard
  Lampz.Callback(74) = "DisableLightingNormal R1, 12,"

  Lampz.MassAssign(75)= cTargetMima1
  Lampz.MassAssign(75)= LightTargetMima1
  Lampz.Callback(75) = "DisableLighting PTargetMima1, 5,"

  Lampz.MassAssign(76)= cTargetMima2
  Lampz.MassAssign(76)= LightTargetMima2
  Lampz.Callback(76) = "DisableLighting PTargetMima2, 5,"

  Lampz.MassAssign(77)= cTargetMima3
  Lampz.MassAssign(77)= LightTargetMima3
  Lampz.Callback(77) = "DisableLighting PTargetMima3, 5,"


  Lampz.MassAssign(78)= cMimalock1
  Lampz.MassAssign(78)= Mimalock1
  Lampz.Callback(78) = "DisableLighting Mimalock1p, 5,"
'
'
  Lampz.MassAssign(79)= cMimalock2
  Lampz.MassAssign(79)= Mimalock2
  Lampz.Callback(79) = "DisableLighting Mimalock2p, 5,"


  Lampz.MassAssign(80)= cTracyChange
  Lampz.MassAssign(80)= lTracyChange
  Lampz.Callback(80) = "FadeTracySkin "


  Lampz.MassAssign(81)= cTracyCross
  Lampz.Callback(81) = "DisableLightingMinMix ptracyCross, 10, 30,"

  Lampz.MassAssign(82)= cpMimaTarget
  Lampz.Callback(82) = "DisableLightingMinMix pMimaTarget, gilvl, 40,"

  Lampz.MassAssign(83)= cLightAutofireWarning
  Lampz.MassAssign(83)= LightAutofireWarning
  Lampz.MassAssign(83)= LightBlastButton
  Lampz.Callback(83) = "DisableLighting pGunSight, 0.2,"
  Lampz.Callback(83) = "DisableLighting pButton, 0.2,"


  Lampz.MassAssign(84)= cFlasherPupilLatencyD
  Lampz.MassAssign(84)= FlasherPupilLatencyD

  Lampz.MassAssign(85)= cFlasherPupilLatencyB
  Lampz.MassAssign(85)= FlasherPupilLatencyB

  Lampz.MassAssign(111) = cGI
  'Lampz.obj(111) = ColtoArray(GI)
  Lampz.Callback(111) = "GIUpdates"


  'Turn off all lamps on startup
  Lampz.Init  'This just turns state of any lamps to 1

  'Immediate update to turn on GI, turn off lamps
  Lampz.Update

  for Each x in GI : x.state=1 : next

  cFlasherPupilLatencyD.state = 0
  cFlasherPupilLatencyB.state = 0


  cLightPortalWall001.state = 0
  cLightPortalWall002.state = 0
  cLightPortalWall003.state = 0
  cLightPortalWall002.timerenabled=False
  cLightPortalWall002.timerenabled=False
  cTracyChange.state = 0

End Sub

dim ShipGiColor, gilvl, ShipGiIntensityMultiplier
ShipGiIntensityMultiplier = 0.4
ShipGiColor = RGB(242,123,36)


dim pupilDcounter, pupilBcounter
sub GIUpdates(ByVal aLvl)
  'debug.print "GI lvl: " & aLvl

  dim x, i
  i = 0

  if Not gilvl = aLvl then

    if vrroom = 1 Then
      pTunnel_light.opacity = 100 + (300 * gilvl)
    end if

    SetInsertBlooms aLvl

    if cGI.state = 1 Then
      LightCorey1.intensity = 33
      LightCorey2.intensity = 33
      LightCorey3.intensity = 33
      LightCorey4.intensity = 33
      LightCorey5.intensity = 15
    Else
      LightCorey1.intensity = 100
      LightCorey2.intensity = 100
      LightCorey3.intensity = 100
      LightCorey4.intensity = 100
      LightCorey5.intensity = 100
    end if

    'pupil latency
    if PupilLatency then
      if gilvl = 1 then
        FlasherPupilLatencyB.visible = false
        FlasherPupilLatencyD.visible = true
        cFlasherPupilLatencyD.state = 1
        pupilDcounter = 150
      elseif gilvl = 0 then
        FlasherPupilLatencyD.visible = false
        FlasherPupilLatencyB.visible = true
        lampz.lvl(84) = 0
        cFlasherPupilLatencyB.state = 1
        pupilBcounter = 80
      Else
        if cGI.state = 1 and cFlasherPupilLatencyD.state = 1 Then
          lampz.lvl(84) = 0
          FlasherPupilLatencyD.visible = false
        end if
        if cGI.state = 0 and cFlasherPupilLatencyB.state = 1 Then
          lampz.lvl(85) = 0
          FlasherPupilLatencyB.visible = false
        end if
      end if
    end if

    for Each x in GI : x.intensityscale=aLvl : next

    FlashOffDL = FlasherOffBrightness*(2/3*aLvl + 1/3)
    for Each x in FlashBases : x.blenddisablelighting = FlashOffDL : next

    If VRRoom > 0 Then
      for Each x in GI_VRBG : x.opacity=aLvl^3 * 10 : next
    End If

'   pmetalwalls.blenddisablelighting = 0.1 * aLvl^2 + 0.1
    pMetalsOn.opacity = 120 * aLvl^2
    pMetalsOn.color = ShipGiColor
    Plastics.blenddisablelighting = 1.2 * aLvl
    PlasticRamp.blenddisablelighting = 0.1*aLvl
    LeftPlastic.blenddisablelighting = 0.05*aLvl - 0.05
    pWireRamps.blenddisablelighting = 0.1 * aLvl + 0.03
    pMimaTarget.blenddisablelighting = aLvl


    if gioff_underpf_boost_status = 1 then
      plywoodholes.blenddisablelighting = 280
    Else
      plywoodholes.blenddisablelighting = 0.2 * aLvl
    end if

    FlBumperFadeTarget(1) = aLvl
    FlBumperFadeTarget(2) = aLvl
    FlBumperFadeTarget(3) = aLvl

    gauge_rollinside.blenddisablelighting = 13*aLvl + 2

    '3,921569E-02
    Primitive020.blenddisablelighting = 0.03 * aLvl - 0.015
    Primitive021.blenddisablelighting = 0.03 * aLvl - 0.015

    for each x in GI_DL : x.blenddisablelighting = 0.2*aLvl : next

    FlasherPL_BS.opacity = aLvl * 100

    if CabinetMode = 0 then
      FlasherPL_RS.opacity = aLvl * 100
      FlasherPL_LS.opacity = aLvl * 100
    else
      FlasherPL_RS_CM.opacity = aLvl * 100
      FlasherPL_LS_CM.opacity = aLvl * 100
    end if

    FlasherGI.opacity = aLvl * 20

    'AO_shadow.opacity = 80 - (aLvl * 10)

    'reflections
'   Reflect1.opacity = 70 * aLvl + 80
'   Reflect2.opacity = 50 * aLvl + 40
    Reflect_pf_left.opacity = 60 * aLvl + 30
    Reflect_pf_right.opacity = 60 * aLvl + 30
    Reflect2b.opacity = 100 * aLvl + 100
    Reflect2b001.opacity = 100 * aLvl + 100
    PlasticsRefl.blenddisablelighting = aLvl + 1

    apron_new.blenddisablelighting = 0.2 * aLvl

    UpdateMaterial "Ship_gi_color",0,0,0,0,0,0,ShipGiIntensityMultiplier * aLvl,ShipGiColor,0,0,False,True,0,0,0,0

'   const ballbrightnessMax = 190   'ball max brightness
'   const ballbrightnessMin = 110   'ball max brightness

    ballbrightness = INT(alvl * (ballbrightnessMax - ballbrightnessMin) + ballbrightnessMin)

  else
    ballbrightness = -1
  end if

' debug.print alvl
  gilvl = alvl
end sub

sub gion
  if tilted then exit sub
' dim x
' for Each x in GI : x.state=1 : next
' for each x in pPlywoodHoles : x.blenddisablelighting = 0.5 : next
' for each x in Metals : x.blenddisablelighting = 0.3 : next
' Plastics.blenddisablelighting = 0.5
' debug.print "gion"

  'if WizardPhase < 1 Or WizardPhase > 2 then
  if WizardPhase <> 1 then
    cGI.state = 1
    UndercabGIOn
'   PlaySoundAt "Relay_On",LightRelayPosition
    Sound_GI_Relay(SoundOn)
    If B2son then Controller.B2SSetData 8, 1 'Backglass logo ON
  end if

' HatchFlasher001.Opacity = 200
' gauge_rollinside.blenddisablelighting = 15

end Sub

sub gioff
' HatchFlasher001.Opacity = 70
' gauge_rollinside.blenddisablelighting = 2

' dim x
' for Each x in GI : x.state=0 : next
' for each x in pPlywoodHoles : x.blenddisablelighting = 0 : next
' for each x in Metals : x.blenddisablelighting = 0.02 : next
' Plastics.blenddisablelighting = 0
  cGI.state = 0
  UndercabGIOff
' PlaySoundAt "Relay_Off",LightRelayPosition
  Sound_GI_Relay(SoundOff)
  if GI_sounds then PlaySound "EFX_GIOFF" & RndInt(1,4),0,CalloutVol,0,0,1,1,1
  If B2son then Controller.B2SSetData 8, 0 'Backglass logo OFF
  startB2S 1
end Sub

dim gioff_underpf_boost_status : gioff_underpf_boost_status = 0
sub gioff_underpf_boost(enabled)
  dim x
  if enabled then
    plywoodholes.blenddisablelighting = 280
    pUnderPF.blenddisablelighting = 15
    gioff_underpf_boost_status = 1
  else
    plywoodholes.blenddisablelighting = 0.2
    pUnderPF.blenddisablelighting = 5
    gioff_underpf_boost_status = 0
  end if
end sub

'set gi on on start
'gion


'Helper functions

Function ColtoArray(aDict)  'converts a collection to an indexed array. Indexes will come out random probably.
  redim a(999)
  dim count : count = 0
  dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
  redim preserve a(count-1) : ColtoArray = a
End Function


'====================
'Class jungle nf
'====================

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?

Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
'Version 0.14 - apophis - added IsLight property to the class
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks

Class LampFader
  Public IsLight(140)         'apophis
  Public FadeSpeedDown(140), FadeSpeedUp(140)
  Private Lock(140), Loaded(140), OnOff(140)
  Public UseFunction
  Private cFilter
  Public UseCallback(140), cCallback(140)
  Public Lvl(140), Obj(140)
  Private Mult(140)
  Public FrameTime
  Private InitFrame
  Public Name

  Sub Class_Initialize()
    InitFrame = 0
    dim x : for x = 0 to uBound(OnOff)  'Set up fade speeds
      FadeSpeedDown(x) = 1/100  'fade speed down
      FadeSpeedUp(x) = 1/80   'Fade speed up
      UseFunction = False
      lvl(x) = 0
      OnOff(x) = False
      Lock(x) = True : Loaded(x) = False
      Mult(x) = 1
      IsLight(x) = False      'apophis
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   ''debug.print Lampz.Locked(100) 'debug
  Public Property Get state(idx) : state = OnOff(idx) : end Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function
  'Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property
  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    'cCallback(idx) = String 'old execute method
    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    dim tmp : tmp = Split(cCallback(idx), "___")

    dim str, x : for x = 0 to uBound(tmp) 'build proc contents
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x)="" then str = str & tmp(x) & " aLVL:"
    Next
    'msgbox "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    'if idx = 132 then msgbox out 'debug
    ExecuteGlobal Out

  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    if Input <> OnOff(idx) then  'discard redundant updates
      OnOff(idx) = input
      Lock(idx) = False
      Loaded(idx) = False
    End If
  End Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
        if typename(aInput) = "Light" then IsLight(aIdx) = True   'apophis - If first object in array is a light, this will be set true+
      end if
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    end if
  end Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub  'Solenoid Handler

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        'debugstr = debugstr & "array found at " & idx & "..."
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
          tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
        obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
      end if
    Next
    ''debug.print debugstr
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
  Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

  Public Sub Update1()   'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
  End Sub

  Public Sub Update2()   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)*Mult(x)) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = Lvl(x)*Mult(x) : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(Lvl(x)*Mult(x))
          Else
            obj(x).Intensityscale = Lvl(x)
          End If
        end if
        if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,Lvl(x)*mult(x) 'Proc
        If Lock(x) Then
          if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True 'finished fading
        end if
      end if
    Next
  End Sub
End Class


'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function


'Helper functions
Sub Proc(string, Callback)  'proc using a string and one argument
  'On Error Resume Next
  dim p : Set P = GetRef(String)
  P Callback
  If err.number = 13 then  msgbox "Proc error! No such procedure: " & vbnewline & string
  if err.number = 424 then msgbox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the end of a 1 dimensional array
  if IsArray(aInput) then 'Input is an array...
    dim tmp : tmp = aArray
    If not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else            'Append existing array with aInput array
      Redim Preserve tmp(uBound(aArray) + uBound(aInput)+1) 'If existing array, increase bounds by uBound of incoming array
      dim x : for x = 0 to uBound(aInput)
        if isObject(aInput(x)) then
          Set tmp(x+uBound(aArray)+1 ) = aInput(x)
        Else
          tmp(x+uBound(aArray)+1 ) = aInput(x)
        End If
      Next
      AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      Redim Preserve aArray(uBound(aArray)+1) 'If array, increase bounds by 1
      if isObject(aInput) then
        Set aArray(uBound(aArray)) = aInput
      Else
        aArray(uBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function

'******************************************************
'****  END LAMPZ
'******************************************************



'************* Mima Magnet ****************

'MiMag.MagnetOn = 1

'
'sub MiMag_hit
' msgbox "hit"
'end sub

Dim MagnetMode : MagnetMode=False

Sub MimaMagnetOn(Enabled)
    If Enabled Then
        MiMag.MagnetOn = 1
        MimaMagnetRelease.enabled = true
        MimaMagnetFlicker.enabled = true
        cLightTurbo.blinkinterval=20
        cLightTurbo.state = 1
        MagnetMode = True
    Else
        MiMag.MagnetOn = 0
        MimaMagnetFlicker.enabled = false
        cLightTurbo.state = 0
        MagnetMode = False
    End If
End Sub

sub MimaMagnetRelease_unhit
    if MagnetMode then
        'debug.print "unhit -> turn on again"
        vpmtimer.addtimer 1000, "MiMag.MagnetOn = 1'"
        vpmtimer.addtimer 1000, "cLightTurbo.state = 1'"
        vpmtimer.addtimer 1500, "MiMag.MagnetOn = 0'"
        vpmtimer.addtimer 2000, "MiMag.MagnetOn = 1'"
    end if
end sub

sub MimaMagnetRelease_hit
    'failsafe to disable magnet
    'debug.print "Magnet released"
    MiMag.MagnetOn = 0
    cLightTurbo.state = 0
end sub


Sub MimaMagnetFlicker_hit
  MimaMagnetFlicker.timerinterval = 800
  MimaMagnetFlicker.timerenabled = true
  cLightTurbo.state = 2
  DOF 151,2 'magnet pulling ball
End Sub

Sub MimaMagnetFlicker_Timer
  cLightTurbo.state = 1
  MimaMagnetFlicker.timerenabled = false
End Sub

dim DisorientBall, DisorientTargetX, DisorientTargetY
sub MimaMagnetDisorient_hit
' debug.print "mimamagnet hit"
  if MiMag.MagnetOn = 1 And VRRoom <> 0 And Not bWizardMode Then
    Set DisorientBall = activeball
    MimaMagnetDisorient.timerenabled = true
  end if
end sub

sub MimaMagnetDisorient_unhit
  MimaMagnetDisorient.timerenabled = false
  vr360_room.blenddisablelighting = 1
end sub

sub MimaMagnetDisorient_timer
' debug.print DisorientBall.x & " - " & DisorientBall.y
  DisorientTargetX = vr360_Room.rotx + DisorientBall.velx
  DisorientTargetY = vr360_Room.roty + DisorientBall.vely
end sub

Const DisorientSpeed = 0.3
sub UpdateDisorient
  if MimaMagnetDisorient.timerenabled Then
    if vr360_Room.rotx < DisorientTargetX Then
      vr360_Room.rotx = vr360_Room.rotx + DisorientSpeed
    Elseif vr360_Room.rotx > DisorientTargetX Then
      vr360_Room.rotx = vr360_Room.rotx - DisorientSpeed
    Else
      vr360_Room.rotx = DisorientTargetX
    end if
    if vr360_Room.roty < DisorientTargety Then
      vr360_Room.roty = vr360_Room.roty + DisorientSpeed
    Elseif vr360_Room.roty > DisorientTargety Then
      vr360_Room.roty = vr360_Room.roty - DisorientSpeed
    Else
      vr360_Room.roty = DisorientTargety
    end if
  end if
end sub


Function FormatScore(ByVal Num)
    dim NumString
    NumString = CStr(abs(Num) )
'    If len(NumString)<10 then
    If len(NumString)>9 then NumString = left(NumString, Len(NumString)-9) & "," & right(NumString,9)
    If len(NumString)>6 then NumString = left(NumString, Len(NumString)-6) & "," & right(NumString,6)
    If len(NumString)>3 then NumString = left(NumString, Len(NumString)-3) & "," & right(NumString,3)
' End If

  if NumString = 0 then NumString = "00"
  FormatScore = NumString
End function




Dim DMDTextOnScore
Dim DMDTextDisplayTime
Dim DMDTextEffect
Sub DMDBigText ( tex,time,effect)
  If UseFlexDMD = 0 Then Exit Sub
  DMDTextOnScore=tex
  DMDTextDisplayTime = FLEXframe + time

  DMDTextEffect=effect
End Sub


Dim EOB_Ramps    ' *EOB*
Dim EOB_Missions
Dim EOB_MimaLoops
Dim EobBonusCounter
Dim TOTALEOBBONUS
Dim TOTALEOBBONUSvisible
dim CurrentMissionCount
'Dim SkipEnd
Sub DMD_EOB_Bonus

  EobBonusCounter = EobBonusCounter + 1
  CurrentMissionCount = 0


  if cLightMission1.state = 1 then CurrentMissionCount = CurrentMissionCount + 1
  if cLightMission2.state = 1 then CurrentMissionCount = CurrentMissionCount + 1
  if cLightMission3.state = 1 then CurrentMissionCount = CurrentMissionCount + 1
  if cLightMission4.state = 1 then CurrentMissionCount = CurrentMissionCount + 1
  if cLightMission5.state = 1 then CurrentMissionCount = CurrentMissionCount + 1
  if cLightMission6.state = 1 then CurrentMissionCount = CurrentMissionCount + 1
  if cLightMission7.state = 1 then CurrentMissionCount = CurrentMissionCount + 1


  if HasPuP then
    Select Case EobBonusCounter
      Case    2 : StopVideosPlaying
'           SkipEnd=True
            DMD_ShowImages "attract",1,500,-1,0
            TOTALEOBBONUS = 0
            TOTALEOBBONUS = TOTALEOBBONUS + (CurrentMissionCount * bonus_EOB_Missions )
            TOTALEOBBONUS = TOTALEOBBONUS +  ( EOB_Ramps * bonus_EOB_Ramps )
            TOTALEOBBONUS = TOTALEOBBONUS + ( EOB_MimaLoops * bonus_EOB_MimaLoops )
            TOTALEOBBONUSvisible = 0
            PuPEvent(401)
            PuPEvent(541)
            DOF 541, DOFPulse


      Case    5 : DMD_ShowText "BONUS",1,FontWhite3,16,True,40,3000
            playsound "EFX_gun_load",1,0.05*CalloutVol


      Case   65 : DMD_ShowText "MISSIONS " & CurrentMissionCount ,1,FontWhite3,11,False,40,3000
            LS_Flash5.StopPlay: LS_Flash5.UpdateInterval = 200: LS_Flash5.Play SeqBlinking ,,EOB_Missions,10
            PuPEvent(542)
            DOF 542, DOFPulse
            PuPlayer.LabelShowPage pDMD,9,2,""
            puPlayer.LabelSet pDMD, "BonusMissions1", ""& FormatScore( CurrentMissionCount  * bonus_EOB_Missions ),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            puPlayer.LabelSet pDMD, "BonusMissions", ""& CurrentMissionCount,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"

      Case   165 : playsound "EFX_gun_load",1,0.02*CalloutVol
            DMD_ShowText FormatScore( CurrentMissionCount * bonus_EOB_Missions ),2,FontWhite3,22,TRUE,40,3000
      Case   180 : DMD_ShowText FormatScore( CurrentMissionCount * bonus_EOB_Missions ),2,FontWhite3,22,False,40,1500
            TOTALEOBBONUSvisible = TOTALEOBBONUSvisible + (CurrentMissionCount * bonus_EOB_Missions )
            playsound "EFX_gun_load",1,0.02*CalloutVol
            pDMDSetPage(pDMDBlank)
      Case   210 : DMD_ShowText "RAMPS " & EOB_Ramps,1,FontWhite3,11,False,40,3000
            LS_Flash5.StopPlay: LS_Flash5.UpdateInterval = 200: LS_Flash5.Play SeqBlinking ,,EOB_Ramps,10
            PuPEvent(543)
            DOF 543, DOFPulse
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD, "Ramps1", ""& FormatScore( EOB_Ramps * bonus_EOB_Ramps ),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            puPlayer.LabelSet pDMD, "Ramptotal", ""& EOB_Ramps,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"

      Case  310 : DMD_ShowText FormatScore( EOB_Ramps * bonus_EOB_Ramps ),2,FontWhite3,22,TRUE,40,3000
      Case  325 : DMD_ShowText FormatScore( EOB_Ramps * bonus_EOB_Ramps ),2,FontWhite3,22,False,40,1500
            TOTALEOBBONUSvisible = TOTALEOBBONUSvisible +  ( EOB_Ramps * bonus_EOB_Ramps )
            playsound "EFX_gun_load",1,0.02*CalloutVol
            pDMDSetPage(pDMDBlank)
      Case  355 : DMD_ShowText "MIMALOOP " & EOB_MimaLoops ,1,FontWhite3,11,False,40,3000
            LS_Flash5.StopPlay: LS_Flash5.UpdateInterval = 200: LS_Flash5.Play SeqBlinking ,,EOB_MimaLoops,10
            PuPEvent(544)
            DOF 544, DOFPulse
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD, "MimaLoopTotal", ""& FormatScore( EOB_MimaLoops * bonus_EOB_MimaLoops ),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            puPlayer.LabelSet pDMD, "MimaLoop3", ""& EOB_MimaLoops ,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case  455 : DMD_ShowText FormatScore( EOB_MimaLoops * bonus_EOB_MimaLoops ),2,FontWhite3,22,TRUE,40,3000
      Case  470 : DMD_ShowText FormatScore( EOB_MimaLoops * bonus_EOB_MimaLoops ),2,FontWhite3,22,False,40,1500
            TOTALEOBBONUSvisible = TOTALEOBBONUSvisible + ( EOB_MimaLoops * bonus_EOB_MimaLoops )
            playsound "EFX_gun_load",1,0.022*CalloutVol
            pDMDSetPage(pDMDBlank)
      Case  500 : DMD_ShowText "BONUS",1,FontWhite3,11,False,40,3000
            PuPEvent(545)
            DOF 545, DOFPulse
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD, "BonusDrainTotal", ""&FormatScore( TOTALEOBBONUSvisible) ,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            DMD_ShowText FormatScore( TOTALEOBBONUSvisible ),2,FontWhite3,22,False,40,3000
            playsound "EFX_gun_load",1,0.023*CalloutVol
      Case  610 : pDMDSetPage(pDMDBlank)
      Case  650 : If BonusMultiplier(CurrentPlayer) >1 Then
              DMD_ShowText "2 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.025*CalloutVol
            PuPEvent(410)
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD,"Multi", "2X",1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"

              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 2 ),2,FontWhite3,22,False,40,3000
            Else
              EobBonusCounter = 810
            End If
      Case  670 : If BonusMultiplier(CurrentPlayer) >2 Then
              DMD_ShowText "3 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.028*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 3 ),2,FontWhite3,22,False,40,3000
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD,"Multi", "3X",1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            Else
              EobBonusCounter = 810
            End If
      Case  690 : If BonusMultiplier(CurrentPlayer) >3 Then
              DMD_ShowText "4 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.03*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 4 ),2,FontWhite3,22,False,40,3000
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD,"Multi", "4X",1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            Else
              EobBonusCounter = 810
            End If
      Case  710 : If BonusMultiplier(CurrentPlayer) >4 Then
              DMD_ShowText "5 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.035*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 5 ),2,FontWhite3,22,False,40,3000
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD,"Multi", "5X",1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            Else
              EobBonusCounter = 810
            End If
      Case  730 : If BonusMultiplier(CurrentPlayer) >5 Then
              DMD_ShowText "6 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.04*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 6 ),2,FontWhite3,22,False,40,3000
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD,"Multi", "6X",1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            Else
              EobBonusCounter = 810
            End If
      Case  750 : If BonusMultiplier(CurrentPlayer) >6 Then
              DMD_ShowText "7 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.04*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 7 ),2,FontWhite3,22,False,40,3000
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD,"Multi", "7X",1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            Else
              EobBonusCounter = 810
            End If
      Case  770 : If BonusMultiplier(CurrentPlayer) >7 Then
              DMD_ShowText "8 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.05*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 8 ),2,FontWhite3,22,False,40,3000
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD,"Multi", "8X",1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            Else
              EobBonusCounter = 810
            End If
      Case  790 : If BonusMultiplier(CurrentPlayer) >8 Then
              DMD_ShowText "9 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.06*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 9 ),2,FontWhite3,22,False,40,3000
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD,"Multi", "9X",1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            Else
              EobBonusCounter = 810
            End If
      Case  810 : If BonusMultiplier(CurrentPlayer) >9 Then
              DMD_ShowText "10 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.07 *CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 10 ),2,FontWhite3,22,False,40,3000
            PuPlayer.LabelShowPage pDMD,9,1.5,""
            puPlayer.LabelSet pDMD,"Multi", "10X",1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            End If
      Case  830 : pDMDSetPage(pDMDBlank)
'           SkipEnd = False
      Case  840: Score(currentplayer) = Score(currentplayer) + ( TOTALEOBBONUS * BonusMultiplier(CurrentPlayer) )
            DMD_ShowText "TOTAL SCORE" ,1,FontWhite3,11,False,40,3000
            PuPEvent(546)
            DOF 546, DOFPulse
            PuPlayer.LabelShowPage pDMD,9,2,""
            puPlayer.LabelSet pDMD, "ScoreTotal", ""&FormatScore( Score(currentplayer)) ,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
            DMD_ShowText FormatScore( Score(currentplayer)),2,FontWhite3,22,true,40,3000
            playsound "EFX_gun_load",1,0.05*CalloutVol
      Case  960 : pDMDSetPage(pDMDBlank)
      Case  961 : DMD_ShowText_Reset
            EobBonusCounter = 0
            EOB_Bonus = False
            EOBGaugeTargetAngle = 60  'should EOB reset to 0X at this point or when bonus figure is shown in DMD?
            EOB_Ramps = 0
            EOB_Missions = 0
            EOB_MimaLoops = 0
            DMD_Stopoverlays = True
            LS_Flash5.StopPlay
            'PuPEvent(402)
            PuPEvent(999)
            pDMDSetPage(pScores)
            EndOfBall2
    End Select
  Else
    Select Case EobBonusCounter
      Case    2 : StopVideosPlaying
            DMD_ShowImages "attract",1,500,-1,0
            TOTALEOBBONUS = 0
            TOTALEOBBONUS = TOTALEOBBONUS + (CurrentMissionCount * bonus_EOB_Missions )
            TOTALEOBBONUS = TOTALEOBBONUS +  ( EOB_Ramps * bonus_EOB_Ramps )
            TOTALEOBBONUS = TOTALEOBBONUS + ( EOB_MimaLoops * bonus_EOB_MimaLoops )
            TOTALEOBBONUSvisible = 0

      Case    5 : DMD_ShowText "BONUS",1,FontWhite3,16,True,40,3000
            playsound "EFX_gun_load",1,0.05*CalloutVol


      Case   50 : DMD_ShowText "MISSIONS " & CurrentMissionCount ,1,FontWhite3,11,False,40,3000
            LS_Flash5.StopPlay: LS_Flash5.UpdateInterval = 200: LS_Flash5.Play SeqBlinking ,,EOB_Missions,10
      Case   75 : playsound "EFX_gun_load",1,0.02*CalloutVol
            DMD_ShowText FormatScore( CurrentMissionCount * bonus_EOB_Missions ),2,FontWhite3,22,TRUE,40,1000
      Case   80 : DMD_ShowText FormatScore( CurrentMissionCount * bonus_EOB_Missions ),2,FontWhite3,22,False,40,500
            TOTALEOBBONUSvisible = TOTALEOBBONUSvisible + (CurrentMissionCount * bonus_EOB_Missions )
            playsound "EFX_gun_load",1,0.02*CalloutVol

      Case  110 : DMD_ShowText "RAMPS " & EOB_Ramps,1,FontWhite3,11,False,40,3000
            LS_Flash5.StopPlay: LS_Flash5.UpdateInterval = 200: LS_Flash5.Play SeqBlinking ,,EOB_Ramps,10
      Case  135 : DMD_ShowText FormatScore( EOB_Ramps * bonus_EOB_Ramps ),2,FontWhite3,22,TRUE,40,1000
      Case  150 : DMD_ShowText FormatScore( EOB_Ramps * bonus_EOB_Ramps ),2,FontWhite3,22,False,40,500
            TOTALEOBBONUSvisible = TOTALEOBBONUSvisible +  ( EOB_Ramps * bonus_EOB_Ramps )
            playsound "EFX_gun_load",1,0.02*CalloutVol

      Case  170 : DMD_ShowText "MIMALOOP " & EOB_MimaLoops ,1,FontWhite3,11,False,40,3000
            LS_Flash5.StopPlay: LS_Flash5.UpdateInterval = 200: LS_Flash5.Play SeqBlinking ,,EOB_MimaLoops,10
      Case  195 : DMD_ShowText FormatScore( EOB_MimaLoops * bonus_EOB_MimaLoops ),2,FontWhite3,22,TRUE,40,1000
      Case  210 : DMD_ShowText FormatScore( EOB_MimaLoops * bonus_EOB_MimaLoops ),2,FontWhite3,22,False,40,500
            TOTALEOBBONUSvisible = TOTALEOBBONUSvisible + ( EOB_MimaLoops * bonus_EOB_MimaLoops )
            playsound "EFX_gun_load",1,0.022*CalloutVol

      Case  230 : DMD_ShowText "BONUS",1,FontWhite3,11,False,40,3000
            DMD_ShowText FormatScore( TOTALEOBBONUSvisible ),2,FontWhite3,22,False,40,3000
            playsound "EFX_gun_load",1,0.023*CalloutVol

      Case  260 : If BonusMultiplier(CurrentPlayer) >1 Then
              DMD_ShowText "2 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.025*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 2 ),2,FontWhite3,22,False,40,3000
            Else
              EobBonusCounter = 500
            End If
      Case  290 : If BonusMultiplier(CurrentPlayer) >2 Then
              DMD_ShowText "3 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.028*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 3 ),2,FontWhite3,22,False,40,3000
            Else
              EobBonusCounter = 500
            End If
      Case  320 : If BonusMultiplier(CurrentPlayer) >3 Then
              DMD_ShowText "4 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.03*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 4 ),2,FontWhite3,22,False,40,3000
            Else
              EobBonusCounter = 500
            End If
      Case  350 : If BonusMultiplier(CurrentPlayer) >4 Then
              DMD_ShowText "5 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.035*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 5 ),2,FontWhite3,22,False,40,3000
            Else
              EobBonusCounter = 500
            End If
      Case  380 : If BonusMultiplier(CurrentPlayer) >5 Then
              DMD_ShowText "6 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.04*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 6 ),2,FontWhite3,22,False,40,3000
            Else
              EobBonusCounter = 500
            End If
      Case  410 : If BonusMultiplier(CurrentPlayer) >6 Then
              DMD_ShowText "7 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.04*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 7 ),2,FontWhite3,22,False,40,3000
            Else
              EobBonusCounter = 500
            End If
      Case  440 : If BonusMultiplier(CurrentPlayer) >7 Then
              DMD_ShowText "8 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.05*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 8 ),2,FontWhite3,22,False,40,3000
            Else
              EobBonusCounter = 500
            End If
      Case  470 : If BonusMultiplier(CurrentPlayer) >8 Then
              DMD_ShowText "9 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.06*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 9 ),2,FontWhite3,22,False,40,3000
            Else
              EobBonusCounter = 500
            End If
      Case  500 : If BonusMultiplier(CurrentPlayer) >9 Then
              DMD_ShowText "10 X",1,FontWhite3,11,True,40,3000
              playsound "EFX_gun_load",1,0.07*CalloutVol
              DMD_ShowText FormatScore( TOTALEOBBONUSvisible * 10 ),2,FontWhite3,22,False,40,3000
            End If
      Case  530 : Score(currentplayer) = Score(currentplayer) + ( TOTALEOBBONUS * BonusMultiplier(CurrentPlayer) )
            DMD_ShowText "TOTAL SCORE" ,1,FontWhite3,11,False,40,3000
            DMD_ShowText FormatScore( Score(currentplayer)),2,FontWhite3,22,true,40,3000
            playsound "EFX_gun_load",1,0.05*CalloutVol
      Case  620 : DMD_ShowText_Reset
            EobBonusCounter = 0
            EOB_Bonus = False
            EOBGaugeTargetAngle = 60  'should EOB reset to 0X at this point or when bonus figure is shown in DMD?
            EOB_Ramps = 0
            EOB_Missions = 0
            EOB_MimaLoops = 0
            DMD_Stopoverlays = True
            LS_Flash5.StopPlay
            EndOfBall2

    End Select
  end if

End Sub

Dim bEndCreditsCounter
sub DMD_ending_credits
  bEndCreditsCounter = bEndCreditsCounter +1
  Select Case bEndCreditsCounter
    Case   2 : gioff : DMD_ShowImages "attract",1,500,-1,0
    Case   30 : DMD_ShowText "CONGRATULATIONS"  ,2,Font6by10W,17,True,50,1500
          LS_AllLightsAndFlashers.Play SeqCircleOutOn,10

    Case  150 : DMD_ShowText "YOU DID IT"     ,1,FontWhite3,11,True,400,2000
          DMD_ShowText ""   ,2,FontWhite3,25,True,400,2000
          LS_AllLightsAndFlashers.Play SeqBlinking ,,5,2:LS_AllLightsAndFlashers.Play SeqUpOn, 10, 1


    Case  300 : DMD_ShowText "VPX TABLE BY"   ,1,FontWhite3,19,False,100,1300
          LS_AllLightsAndFlashers.Play SeqDownOn, 10, 1

    Case  400 : DMD_ShowText_Reset
          DMD_ShowImages "vpwlogo",1,500,-1,0
          LS_AllLightsAndFlashers.Play SeqBlinking ,,5,2:LS_AllLightsAndFlashers.Play SeqUpOn, 10, 1

    Case  500 : DMD_ShowImages "attract",1,500,-1,0
          DMD_ShowText "LEAD DEV"     ,1,FontWhite3,11,False,100,2000
          DMD_ShowText "iaakki"     ,2,Font6by10W,25,True,100,2000
          LS_AllLightsAndFlashers.Play SeqFanLeftUpOn, 40

    Case  700 : DMD_ShowText "OTHER DEVS"     ,1,FontWhite3,11,False,100,6000
          DMD_ShowText "Oqqsan + Daphene"   ,2,Font6by10W,25,False,100,1000
          LS_AllLightsAndFlashers.Play SeqFanRightDownOn, 40
    Case  800 : DMD_ShowText "apophis + Tomate"   ,2,Font6by10W,25,False,100,1000
          LS_AllLightsAndFlashers.Play SeqFanLeftUpOn, 40
    Case  900 : DMD_ShowText "Flupper + Sixtoe"   ,2,Font6by10W,25,False,100,1000
          LS_AllLightsAndFlashers.Play SeqFanRightDownOn, 40
    Case 1000 : DMD_ShowText "Lumi + Embee + VPW" ,2,Font6by10W,25,False,100,1000
          LS_AllLightsAndFlashers.Play SeqFanLeftUpOn, 40

    Case 1100 : DMD_ShowText "GRAPHICS"     ,1,FontWhite3,11,False,100,2000
          DMD_ShowText "AstroNasty"   ,2,Font6by10W,25,True,100,2000
          LS_AllLightsAndFlashers.Play SeqFanRightDownOn, 40

    Case 1300 : DMD_ShowText "TESTING"      ,1,FontWhite3,11,False,100,4500
          DMD_ShowText "RIK"        ,2,Font6by10W,25,True,100,2000
          LS_AllLightsAndFlashers.Play SeqFanLeftUpOn, 40

    Case 1450 : DMD_ShowText "PinStratsDan"   ,2,Font6by10W,25,True,100,2000
          LS_AllLightsAndFlashers.Play SeqFanRightDownOn, 40

    Case 1600 : DMD_ShowText "MOVIE BY"     ,1,FontWhite3,11,False,100,2000
          DMD_ShowText "SETH ICKERMAN"    ,2,Font6by10W,25,True,100,2000
          LS_AllLightsAndFlashers.Play SeqBlinking ,,5,2:LS_AllLightsAndFlashers.Play SeqUpOn, 10, 1

    Case 1750 : DMD_ShowText "MUSIC BY"     ,1,FontWhite3,11,False,100,2000
          DMD_ShowText "CARPENTER BRUT"   ,2,Font6by10W,25,True,100,2000
          LS_AllLightsAndFlashers.Play SeqBlinking ,,5,2:LS_AllLightsAndFlashers.Play SeqUpOn, 10, 1

'   Case  430 : DMD_ShowText_Reset
'         DMD_ShowImages "attract",1,500,-1,0
    Case 2000 : GaldorLockRelease
  end select
end sub



Dim bAttractModeCounter, AttractVideoCounter

Sub DMD_Attract
  bAttractModeCounter = bAttractModeCounter +1

  Select Case bAttractModeCounter

    'Case   0 :   DMD_ShowText_Reset:DMD_ShowImages_Reset
    Case   1 :  showMissionVideo = 80

    'Case  100 : DMD_ShowText "PINBALL BY VPW 2022",3,FontScoreInactiv2,28,True,100,3000

    Case  490 : DMD_ShowImages "attract",1,500,-1,0:StopVideosPlaying 'DMD_ShowText_Reset
    'Case  280 : DMD_ShowImages "attract",1,500,-1,0
    Case  500 : If credits = 0 Then
            DMD_ShowImages "insertcoin",2,500,-1,0
          Else
            DMD_ShowText "CREDITS  " & credits ,1,FontWhite3,16,False,400,2000
            if AmbientAudio = 1 Then
              DMD_ShowText "NO SONGS FOUND",2,Font6by10W,28,True,200,2000
            Else
              DMD_ShowText "PRESS START TO PLAY",2,FontScoreInactiv2,28,True,200,2000
            end if
          End If
    Case  630 : DMD_ShowText_Reset
          DMD_ShowImages "attract",1,500,-1,0

    Case  650 : DMD_ShowText "HIGHSCORE 1",1,FontWhite3,7,True,50,6600
    Case  670 : DMD_ShowText FormatScore(HighScore(0)),2,FontWhite3,17,False,100,1800
          DMD_ShowText HighScoreName(0),3,FontWhite3,27,False,100,1800
    Case  710 : DMD_ShowText "HIGHSCORE 1",1,FontWhite3,7,False,100,6600
    Case  800 :
          DMD_ShowText "HIGHSCORE 2",1,FontWhite3,7,True,100,6600
          DMD_ShowText FormatScore(HighScore(1)),2,FontWhite3,17,False,100,1800
          DMD_ShowText HighScoreName(1),3,FontWhite3,27,False,100,1800
    Case  880 : DMD_ShowText "HIGHSCORE 2",1,FontWhite3,7,False,100,6600

    Case  940 : DMD_ShowText "HIGHSCORE 3",1,FontWhite3,7,True,200,6600
          DMD_ShowText FormatScore(HighScore(2)),2,FontWhite3,17,False,100,1800
          DMD_ShowText HighScoreName(2),3,FontWhite3,27,False,100,1800
    Case  1030 :  DMD_ShowText "HIGHSCORE 3",1,FontWhite3,7,False,100,6600

    Case  1050 : DMD_ShowText_Reset
    Case  1120 : DMD_ShowText "LAST PLAYER",1,FontWhite3,8,False,100,2300
    Case  1150 : DMD_ShowText FormatScore(Score(currentplayer)),2,FontWhite3,20,True,40,1800
    Case  1190 : DMD_ShowText FormatScore(Score(currentplayer)),2,FontWhite3,20,False,1800,1800
    Case 1290 : DMD_ShowText_Reset
    Case 1340 : DMD_ShowText "GAME     ",1,FontWhite3,13,True,500,4500
    Case 1355 : DMD_ShowText "     OVER",2,FontWhite3,13,True,500,4500
    Case 1342 : if AmbientAudio = 1 then DMD_ShowText "NO SONGS FOUND",3,Font6by10W,27,False,200,3000

    Case 1540 : DMD_ShowText_Reset
'   Case 1490 : DMD_ShowText "PINBALL BY VPW 2022",3,FontScoreInactiv2,28,True,100,3000
    Case 1590 : DMD_ShowText "PINBALL BY",1,FontWhite3,10,True,40,3000
    'Case 1491 :  DMD_ShowText "VPin Workshop 2022",2,Font6by10W,25,True,50,3000
    Case 1640 : DMD_ShowText "PINBALL BY",1,FontWhite3,10,False,100,3000
          DMD_ShowText "VPin Workshop 2022",2,Font6by10W,25,False,100,3000
    Case 1785 : DMD_ShowText_Reset
    Case 1790 : DMD_ShowImages_Reset
    Case 1850 : PlayAttractVideo
    Case 4400 : bAttractModeCounter=0
  End Select
End Sub

Sub DMD_Attract2
  bAttractModeCounter = bAttractModeCounter +1

  Select Case bAttractModeCounter

    Case   10 : DMD_ShowImages "attract",1,500,-1,0
    Case   50 : DMD_ShowText "BLOOD",1,FontWhite3,8,True,40,3000
          DMD_ShowText "MACHINES",2,FontWhite3,19,True,50,3000
          DMD_ShowText "PINBALL BY VPW 2022",3,FontScoreInactiv2,28,False,100,3000
    Case  100 : DMD_ShowText "BLOOD",1,FontWhite3,8,False,100,3000
          DMD_ShowText "MACHINES",2,FontWhite3,19,False,100,3000
          DMD_ShowText "PINBALL BY VPW 2022",3,FontScoreInactiv2,28,True,100,3000
    Case  270 : DMD_ShowText_Reset
    Case  300 : If credits = 0 Then
            DMD_ShowImages "insertcoin",2,500,-1,0
          Else
            DMD_ShowText "CREDITS  " & credits ,1,FontWhite3,16,False,400,2000
            DMD_ShowText "PRESS START TO PLAY",2,FontScoreInactiv2,28,True,200,2000
          End If
    Case  430 : DMD_ShowText_Reset
          DMD_ShowImages "attract",1,500,-1,0

    Case  450 : DMD_ShowText "HIGHSCORE 1",1,FontWhite3,7,True,50,6600
    Case  470 : DMD_ShowText FormatScore(HighScore(0)),2,FontWhite3,17,False,100,1800
          DMD_ShowText HighScoreName(0),3,FontWhite3,27,False,100,1800
    Case  510 : DMD_ShowText "HIGHSCORE 1",1,FontWhite3,7,False,100,6600
    Case  600 :
          DMD_ShowText "HIGHSCORE 2",1,FontWhite3,7,True,100,6600
          DMD_ShowText FormatScore(HighScore(1)),2,FontWhite3,17,False,100,1800
          DMD_ShowText HighScoreName(1),3,FontWhite3,27,False,100,1800
    Case  680 : DMD_ShowText "HIGHSCORE 2",1,FontWhite3,7,False,100,6600

    Case  740 : DMD_ShowText "HIGHSCORE 3",1,FontWhite3,7,True,200,6600
          DMD_ShowText FormatScore(HighScore(2)),2,FontWhite3,17,False,100,1800
          DMD_ShowText HighScoreName(2),3,FontWhite3,27,False,100,1800
    Case  830 : DMD_ShowText "HIGHSCORE 3",1,FontWhite3,7,False,100,6600

    Case  850 : DMD_ShowText_Reset
    Case  920 : DMD_ShowText "LAST PLAYER",1,FontWhite3,8,False,100,2300
    Case  950 : DMD_ShowText FormatScore(Score(currentplayer)),2,FontWhite3,20,True,40,1800
    Case  990 : DMD_ShowText FormatScore(Score(currentplayer)),2,FontWhite3,20,False,1800,1800
    Case 1090 : DMD_ShowText_Reset
    Case 1140 : DMD_ShowText "GAME     ",1,FontWhite3,16,True,500,4500
    Case 1155 : DMD_ShowText "     OVER",2,FontWhite3,16,True,500,4500
    Case 1385 : DMD_ShowText_Reset
    Case 1390 : DMD_ShowImages_Reset
    Case 1450 : PlayAttractVideo
    Case 4000 : bAttractModeCounter=0
  End Select
End Sub

sub PlayAttractVideo
  select case AttractVideoCounter
    Case 0 : showMissionVideo = 73 : vpmTimer.addtimer 11500, " StopVideosPlaying: bAttractModeCounter=0 '"
    Case 1 : showMissionVideo = 75 : vpmTimer.addtimer 11500, " StopVideosPlaying: bAttractModeCounter=0 '"
    Case 2 : showMissionVideo = 76 : vpmTimer.addtimer 4500, " StopVideosPlaying: bAttractModeCounter=0 '"
    Case 3 : showMissionVideo = 78 : vpmTimer.addtimer 15500, " StopVideosPlaying: bAttractModeCounter=0 '"
    Case 4 : showMissionVideo = 79 : vpmTimer.addtimer 6000, " StopVideosPlaying: bAttractModeCounter=0 '"
    Case 5 : showMissionVideo = 77 : vpmTimer.addtimer 2500, " StopVideosPlaying: bAttractModeCounter=0 '"
  end select
  AttractVideoCounter = AttractVideoCounter + 1
  if AttractVideoCounter > 5 then AttractVideoCounter = 0
end sub


'************************************************************************************************************
'Example commands to use in your script
'************************************************************************************************************
' DMD_ShowImages "gameover",4,100,5000,0  ' 100ms each of the 4 frames for max 5000ms
' DMD_ShowImages "gameover",4,100,5000,1000 '(last image= 1000ms)
' DMD_ShowImages "gameover",4,100,10,0    ' run sequence 10 times only
' DMD_ShowImages "gameover",4,100,-1,0    ' run sequence forever until new command is issued

' DMD_ShowImages "balllost",1,100,2000,0  ' these two does the same since its just 1 image
' DMD_ShowImages "balllost",1,2000,2000,0

' DMD_ShowText "THIS IS TESTING1",1,FontScoreActive2,5,True,100,5000
' DMD_ShowText "THIS IS two",2,FontScoreActive,12,false,100,5000
' DMD_ShowText "last one",3,FontScoreinActive2,22,True,100,5000
' DMD_ShowText_Reset   ' if it needs to stop

Sub CheckHighscore()
  If Score(CurrentPlayer) > HighScore(2) Then


    AwardSpecial
    vpmtimer.addtimer 1000, "PlaySound ""vo_contratulationsgreatscore"" '"
    'enter player's name
    StartHighscoreName
    StopVideosPlaying

    DMD_Stopoverlays = False
    DOF 403, DOFPulse   'DOF MX - Hi Score
  Else
    EndOfBallComplete
  End If
End Sub


Dim EnterName
Dim EnterNamePos
Dim EnterNameCurrenLetter
Sub StartHighscoreName
  hsbModeActive = True
  PuPEvent(1000)
  PuPEvent(401)
  Playsound "bellhs"   ' fixing need sound =!=
  EnterName = ""
  EnterNamePos = 1
  EnterNameCurrenLetter = 1
  bEnterHighCounter=0
  Enterblinking=0
End Sub


Dim Enterblinking
Dim bEnterHighCounter
Sub DMD_EnterHigh
  PupDMDHSInput
  Dim Displaytext
  Enterblinking = Enterblinking +1
  If ( Enterblinking mod 50 ) >20 Then
    If Len(EnterName) = 0 Then Displaytext = "<" & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ",EnterNamePos,1) & ">"
    If Len(EnterName) = 1 Then Displaytext = Entername & "<" & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ",EnterNamePos,1) & ">"
    If Len(EnterName) = 2 Then Displaytext = Entername & "<" & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ",EnterNamePos,1) & ">"
    If Len(EnterName) = 3 Then Displaytext = "- " & Entername & " -"
  Else
    If Len(EnterName) = 0 Then Displaytext = "- >"
    If Len(EnterName) = 1 Then Displaytext = Entername & "- -"
    If Len(EnterName) = 2 Then Displaytext = Entername & "- -"
    If Len(EnterName) = 3 Then Displaytext = "-    -"
  End If

' If ( Enterblinking mod 50 ) >20 Then
'   If Len(EnterName) = 0 Then Displaytext = "< " & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",EnterNamePos,1) & " >"
'   If Len(EnterName) = 1 Then Displaytext = Entername & "< " & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",EnterNamePos,1) & " >"
'   If Len(EnterName) = 2 Then Displaytext = Entername & "< " & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",EnterNamePos,1) & " >"
'   If Len(EnterName) = 3 Then Displaytext = "- " & Entername & " -"
' Else
'   If Len(EnterName) = 0 Then Displaytext = "-   >"
'   If Len(EnterName) = 1 Then Displaytext = Entername & "-   -"
'   If Len(EnterName) = 2 Then Displaytext = Entername & "-   -"
'   If Len(EnterName) = 3 Then Displaytext = "-        -"
' End If

  If bEnterHighCounter < 190 Then DMD_ShowText Displaytext,2,FontWhite3,19,False,100,-1

  If bEnterHighCounter > 9 Then bEnterHighCounter=bEnterHighCounter+1

  Select Case bEnterHighCounter

    Case   0 :
      DMD_ShowImages "attract",1,1500,-1,0
      DMD_ShowText "ENTER NAME",1,FontWhite3,7,False,100,-1
      DMD_ShowText Formatscore(Score(CurrentPlayer)),3,FontScoreInactiv2,28,False,100,-1


    Case  200,201 :
      DMD_showtext_reset

    Case 205 :

      If Score(CurrentPlayer) > HighScore(0) Then
        HighScore(2)=HighScore(1)
        HighScoreName(2)=HighScoreName(1)
        HighScore(1)=HighScore(0)
        HighScoreName(1)=HighScoreName(0)
        HighScore(0)=Score(CurrentPlayer)
        HighScoreName(0)=EnterName

      ElseIf Score(CurrentPlayer) > HighScore(1) Then
        HighScore(2)=HighScore(1)
        HighScoreName(2)=HighScoreName(1)
        HighScore(1)=Score(CurrentPlayer)
        HighScoreName(1)=EnterName
      ElseIf Score(CurrentPlayer) > HighScore(2) Then
        HighScore(2)=Score(CurrentPlayer)
        HighScoreName(2)=EnterName
      End If

      savehs

    Case 210 :
      PuPEvent(1001)
      PuPEvent(402)
      pDMDSetPage(pScores)
      DMD_Stopoverlays = True
      bEnterHighCounter=0
      hsbModeActive = False
      EndOfBallComplete

  End Select

End Sub


Sub EnterHighscoreName(keycode)
  If Len(EnterName) < 3 Then
    If keycode = LeftFlipperKey Then
      EnterNamePos = EnterNamePos - 1
      If EnterNamePos < 1 Then EnterNamePos = 37
'     Playsound "fx_Previous"             ' add a sound for this
    End If

    If keycode = RightFlipperKey Then
      EnterNamePos = EnterNamePos + 1
      If EnterNamePos >37 Then EnterNamePos = 1
'     Playsound "fx_Next"               ' add a sound for this
    End If

    If keycode = StartGameKey or keycode = PlungerKey Then
      soundStartButton()
      EnterName=EnterName + mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ",EnterNamePos,1)
      If Len(EnterName) =3 Then
        bEnterHighCounter=10 ' ending started
      Else
        EnterNamePos = 1 'tweak to start always on letter A
      end if

    End if
  End If
End Sub


Dim Flex_Overlay_TextBlink(4)
Dim Flex_Overlay_TextBlinkOn(4)
Dim Flex_Overlay_TextBlinkDelay(4)
Dim Flex_Overlay_TextBlinkDelayCounter(4)
Dim Flex_Overlay_TextTimer(4)

Sub DMD_ShowText_Reset
  Flex_Overlay_TextTimer(1)=0
  Flex_Overlay_TextTimer(2)=0
  Flex_Overlay_TextTimer(3)=0
End Sub


Sub DMD_ShowText ( txt , nr , font , posy , blinking , blinkdelay , timer )
  FlexDMD.Stage.GetLabel("overlaytext"&nr).font = font
  FlexDMD.Stage.GetLabel("overlaytext"&nr).Text = txt
  FlexDMD.Stage.GetLabel("overlaytext"&nr).SetAlignedPosition 64, posy , FlexDMD_Align_Center

  Flex_Overlay_TextBlink(nr) = blinking
  Flex_Overlay_TextBlinkOn(nr) = False
  Flex_Overlay_TextBlinkDelay(nr) = blinkdelay
  Flex_Overlay_TextBlinkDelayCounter(nr) = 0
  Flex_Overlay_TextTimer(nr) = timer
End Sub



Sub DMD_OverlayText

  Dim nr
  For nr = 1 to 4
    If Flex_Overlay_TextTimer(nr) = -1 Then
      FlexDMD.Stage.GetLabel("overlaytext"&nr).visible = True
    Elseif Flex_Overlay_TextTimer(nr) > 0 Then
      Flex_Overlay_TextTimer(nr) = Flex_Overlay_TextTimer(nr) - Flextimer
      If Flex_Overlay_TextTimer(nr) < 0 Then Flex_Overlay_TextTimer(nr) = 0
      If Flex_Overlay_TextBlink(nr) Then

        If Flex_Overlay_TextBlinkDelayCounter(nr) > 0 Then
          Flex_Overlay_TextBlinkDelayCounter(nr) = Flex_Overlay_TextBlinkDelayCounter(nr) - Flextimer
        Else
          Flex_Overlay_TextBlinkDelayCounter(nr) = Flex_Overlay_TextBlinkDelay(nr)
          If Flex_Overlay_TextBlinkOn(nr) Then
            Flex_Overlay_TextBlinkOn(nr)= false
            FlexDMD.Stage.GetLabel("overlaytext"&nr).visible = False
          Else
            Flex_Overlay_TextBlinkOn(nr)=True
            FlexDMD.Stage.GetLabel("overlaytext"&nr).visible = True
          End If
        End If
      Else
        FlexDMD.Stage.GetLabel("overlaytext"&nr).visible = True
      End If
    Else
      FlexDMD.Stage.GetLabel("overlaytext"&nr).visible = False
    End If
  Next
End Sub


Dim FlexOverlay
Const FlexTimer=17 ' dmd update timer ...find the interval on the timer : 17ms is normal for 60hz dont need to update dmd faster than that
Dim Flex_Overlay_FrameCounter
Dim Flex_Overlay_DelayCounter
Dim Flex_Overlay_TotalTimeCounter
Dim Flex_Overlay_CurrentImage
Sub DMD_Overlay

  If DMD_Stopoverlays Then
    DMD_Stopoverlays = False
    Flex_Overlay_TextTimer(1) =0
    Flex_Overlay_TextTimer(2) =0
    Flex_Overlay_TextTimer(3) =0
    FlexOverlay = 0
    If Flex_Overlay_CurrentImage > 0 Then FlexDMD.Stage.GetImage(Flex_Overlay_image & Flex_Overlay_CurrentImage ).Visible=False
    Exit Sub
  End If

  If Flex_Overlay_TotalTime > 100 Then
    Flex_Overlay_TotalTimeCounter = Flex_Overlay_TotalTimeCounter - FlexTimer
    If Flex_Overlay_TotalTimeCounter < 0 Then
      FlexOverlay = 0
      If Flex_Overlay_CurrentImage > 0 Then FlexDMD.Stage.GetImage(Flex_Overlay_image & Flex_Overlay_CurrentImage ).Visible=False
      Exit Sub
    End If
  End If

  If Flex_Overlay_DelayCounter<1 Then

    Flex_Overlay_FrameCounter = Flex_Overlay_FrameCounter + 1
    Flex_Overlay_DelayCounter = Flex_Overlay_Delay
    If Flex_Overlay_FrameCounter = Flex_Overlay_Frames Then Flex_Overlay_DelayCounter = Flex_Overlay_LastFrame
    If Flex_Overlay_FrameCounter > Flex_Overlay_Frames Then Flex_Overlay_FrameCounter = 1

    If Flex_Overlay_TotalTime > 0 And Flex_Overlay_TotalTime < 100 Then
      Flex_Overlay_TotalTimeCounter = Flex_Overlay_TotalTimeCounter -1
      If Flex_Overlay_TotalTimeCounter < 1 Then
        FlexOverlay = 0
        If Flex_Overlay_CurrentImage > 0 Then FlexDMD.Stage.GetImage(Flex_Overlay_image & Flex_Overlay_CurrentImage ).Visible=False
        Exit Sub
      End If
    End If

  Else
    Flex_Overlay_DelayCounter = Flex_Overlay_DelayCounter - FlexTimer
  End If

  If Flex_Overlay_CurrentImage <> Flex_Overlay_FrameCounter Then
    If Flex_Overlay_CurrentImage > 0 Then FlexDMD.Stage.GetImage(Flex_Overlay_image & Flex_Overlay_CurrentImage ).Visible=False
    Flex_Overlay_CurrentImage = Flex_Overlay_FrameCounter
  End If
' debug.print "img=" & Flex_Overlay_image & Flex_Overlay_FrameCounter
  FlexDMD.Stage.GetImage(Flex_Overlay_image & Flex_Overlay_FrameCounter ).Visible=True



End Sub

Dim Flex_Overlay_Image
Dim Flex_Overlay_Frames
Dim Flex_Overlay_Delay
Dim Flex_Overlay_TotalTime
Dim Flex_Overlay_LastFrame
Sub DMD_ShowImages( img, nrframes, delay , totaltime , lastframe ) ' TOTALTIME <100 = play sequence only nr of timer : totaltime can be -1 = until turned off with a new command

  If nrframes<1 Then nrframes = 1
  If delay<1 Then delay = 1
  If img="" Then Exit Sub

  If FlexOverlay=1 And Flex_Overlay_CurrentImage > 0 Then FlexDMD.Stage.GetImage(Flex_Overlay_image & Flex_Overlay_CurrentImage ).Visible=False

  Flex_Overlay_CurrentImage=0

  FlexOverlay=1
  Flex_Overlay_image = img
  Flex_Overlay_Frames = nrframes
  Flex_Overlay_Delay = delay
  Flex_Overlay_TotalTime = totaltime
  Flex_Overlay_LastFrame = lastframe
  If Flex_Overlay_LastFrame < Flex_Overlay_Delay Then Flex_Overlay_LastFrame = Flex_Overlay_Delay

  Flex_Overlay_FrameCounter = 1
  Flex_Overlay_DelayCounter = Flex_Overlay_Delay
  Flex_Overlay_TotalTimeCounter = Flex_Overlay_TotalTime

End Sub

Sub DMD_ShowImages_Reset
  FlexOverlay=0
End Sub



'***
' B2S Light Show
' cause i mean everyone loves a good light show
'1= the background
'2= left dude
'3= right dude
'4= left chick
'5= right chick
'6= robot
'7= robot eyes
'8= the logo

' /////////////////////
' example B2S call
' startB2S(#)   <---- this is the trigger code (if you want to add it to something like a bumper)

Dim b2sstep
b2sstep = 0
b2sflash.enabled = 0
Dim b2satm

Sub startB2S(aB2S)
    b2sflash.enabled = 1
    b2satm = ab2s
End Sub

Sub b2sflash_timer
  dim i
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
        for i = 1 to 7
          Controller.B2SSetData i, 0
        next
    End Select
    End If
End Sub

'*********************
'VR Mode
'*********************
DIM VRThings
If VRRoom > 0 Then
' PinCab_Backbox.image="Pincab_Backbox_DM"
  if HasPuP then
    fDMD.visible = 0 : pupDMD.visible = 1
    VRBGSpeaker.visible = 1
    VRBGSpeaker.imageA = "Pincab_Speaker_pup"
  Else
    fDMD.visible = 1 : pupDMD.visible = 0
    VRBGSpeaker.visible = 1
    VRBGSpeaker.imageA = "Pincab_Speaker2"
  end if

  PinCab_Backglass.blenddisablelighting = 5
  PinCab_Rails.visible = 1
  If VRRoom = 1 Then
    for each VRThings in VR_360:VRThings.visible = 1:Next
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
  End If
  If VRRoom = 2 Then
    for each VRThings in VR_360:VRThings.visible = 0:Next
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 1:Next
  End If
  If VRRoom = 3 Then
    for each VRThings in VR_360:VRThings.visible = 0:Next
    for each VRThings in VR_Cab:VRThings.visible = 0:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
    DMD.visible=1
    PinCab_Backbox.visible = 1
    PinCab_Backglass.visible = 1
  End If
' for each VRThings in InsertBlooms : VRThings.TransmissionScale = 0 : next
  for each VRThings in InsertLamps : VRThings.TransmissionScale = 0 : next  'added this as insert lamps casted some issues on side blades
  LightShip001.TransmissionScale = 0
  LightShip002.TransmissionScale = 0
  LightShip003.TransmissionScale = 0
  LightShip004.TransmissionScale = 0
  LightShip005.TransmissionScale = 0
' pWireRamps.image = ""
' pWireRamps.material = "VRMetalWires"
  SetBackglass
  PinCab_Backglass.visible = 0
  BGDark.visible = 1
' VRBGSpeaker.visible = 1
  for Each VRThings in GI_VRBG : VRThings.visible = 1 : next
Else
  for each VRThings in VR_360:VRThings.visible = 0:Next
  for each VRThings in VR_Cab:VRThings.visible = 0:Next
  for each VRThings in VR_Min:VRThings.visible = 0:Next
  for Each VRThings in GI_VRBG : VRThings.visible = 0 : next
  fDMD.visible = 0 : pupDMD.visible = 0
  VRBGSpeaker.visible = 0
End If

If CabinetMode=1 then
' FlasherPL_LS_CM.amount = 2000
' FlasherPL_LS_CM.opacity = 200
' PinCab_Blades.blenddisablelighting = 20

  PinCab_Blades.size_y = 2000
  hide.rotx = -9
  hide.z = 600
  hide.depthbias = 900
  hide001.rotx = -9
  hide001.z = 600
  hide001.depthbias = 900
  Pincab_Rails.visible=0
else
  PinCab_Blades.size_y = 1000
  hide.rotx = -2.5
  hide.z = 137
  hide.depthbias = 450
  hide001.rotx = -2.5
  hide001.z = 137
  hide001.depthbias = 450
  Pincab_Rails.visible=1
end if

'******************* VR Plunger **********************

Sub TimerVRPlunger_Timer
  If Pincab_Shooter.Y < 45 then
       Pincab_Shooter.Y = Pincab_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger2_Timer
  Pincab_Shooter.Y = -70 + (5* Plunger.Position) -20
  timervrplunger2.enabled = 0
End Sub

'****************************************************************


'******************************************************
'*******  Set VR Backglass Flashers *******
'******************************************************

Sub SetBackglass()
  Dim obj

  For Each obj In VRBackglass
    obj.x = obj.x
    obj.height = - obj.y + 300
    obj.y = -40 'adjusts the distance from the backglass towards the user
  Next


  For Each obj In VRBackglassSpeaker
    obj.x = obj.x
    obj.height = - obj.y + 287
    obj.y = 27.2 'adjusts the distance from the backglass towards the user
    obj.rotx = -86
  Next

End Sub

'*****************************************************************************************************
' VR Backglass Flasher Code
'*****************************************************************************************************

dim VRBGFL2lvl
dim VRBGFL3lvl
dim VRBGFL7lvl

sub VRBGFL2
  If VRRoom > 0 Then
    VRBGFL2lvl = 1
    VRBGFL2timer_timer
  End If
end sub

sub VRBGFL2timer_timer
  if Not VRBGFL2timer.enabled then
    VRBGFL2_1.visible = true
    VRBGFL2_2.visible = true
    VRBGFL2_3.visible = true
    VRBGFL2_4.visible = true
    VRBGFL2timer.enabled = true
  end if
  VRBGFL2lvl = 0.8 * VRBGFL2lvl - 0.01
  if VRBGFL2lvl < 0 then VRBGFL2lvl = 0
  VRBGFL2_1.opacity = 100 * VRBGFL2lvl^1.5
  VRBGFL2_2.opacity = 100 * VRBGFL2lvl^1.5
  VRBGFL2_3.opacity = 100 * VRBGFL2lvl^1.5
  VRBGFL2_4.opacity = 100 * VRBGFL2lvl^2
  if VRBGFL2lvl =< 0 Then
    VRBGFL2_1.visible = false
    VRBGFL2_2.visible = false
    VRBGFL2_3.visible = false
    VRBGFL2_4.visible = false
    VRBGFL2timer.enabled = false
  end if
end sub

sub VRBGFL3
  If VRRoom > 0 Then
    VRBGFL3lvl = 1
    VRBGFL3timer_timer
  End If
end sub

sub VRBGFL3timer_timer
  if Not VRBGFL3timer.enabled then
    VRBGFL3_1.visible = true
    VRBGFL3_2.visible = true
    VRBGFL3_3.visible = true
    VRBGFL3_4.visible = true
    VRBGFL3timer.enabled = true
  end if
  VRBGFL3lvl = 0.8 * VRBGFL3lvl - 0.01
  if VRBGFL3lvl < 0 then VRBGFL3lvl = 0
  VRBGFL3_1.opacity = 100 * VRBGFL3lvl^1.5
  VRBGFL3_2.opacity = 100 * VRBGFL3lvl^1.5
  VRBGFL3_3.opacity = 100 * VRBGFL3lvl^1.5
  VRBGFL3_4.opacity = 100 * VRBGFL3lvl^2
  if VRBGFL3lvl =< 0 Then
    VRBGFL3_1.visible = false
    VRBGFL3_2.visible = false
    VRBGFL3_3.visible = false
    VRBGFL3_4.visible = false
    VRBGFL3timer.enabled = false
  end if
end sub

sub VRBGFL7
  If VRRoom > 0 Then
    VRBGFL7lvl = 1
    VRBGFL7timer_timer
  End If
end sub

sub VRBGFL7timer_timer
  if Not VRBGFL7timer.enabled then
    VRBGFLEyeL.visible = true
    VRBGFLEyeR.visible = true
    VRBGBulb008.visible = true
    VRBGBulb009.visible = true
    VRBGFL7timer.enabled = true
  end if
  VRBGFL7lvl = 0.80 * VRBGFL7lvl - 0.01
  if VRBGFL7lvl < 0 then VRBGFL7lvl = 0
  VRBGFLEyeL.opacity = 100 * VRBGFL7lvl^0.8
  VRBGFLEyeR.opacity = 100 * VRBGFL7lvl^0.8
  VRBGBulb008.opacity = 100 * VRBGFL7lvl^1
  VRBGBulb009.opacity = 100 * VRBGFL7lvl^1
  if VRBGFL7lvl =< 0 Then
    VRBGFLEyeL.visible = false
    VRBGFLEyeR.visible = false
    VRBGBulb008.visible = false
    VRBGBulb009.visible = false
    VRBGFL7timer.enabled = false
  end if
end sub

Dim VRBGstep7
VRBGstep7 = 0
VRBGFlash7.enabled = false

Sub VRBGFlash7_timer
  If VRRoom > 0 Then
    VRBGstep7 = VRBGstep7 + 1
    Select case VRBGstep7
      case 0
        VRBGFL7
      case 1
        VRBGFL7
      Case 2
        VRBGFL7
      case 3
        VRBGFL7
      case 4
        VRBGFL7
      Case 6
        VRBGFL7
      Case 7
        VRBGFL7
      Case 8
        VRBGFL7
      Case 9
        VRBGFL7
        VRBGstep7 = 0
        VRBGFlash7.enabled = false
    End Select
  End If
End Sub

'reposition dmd for VR
if VRRoom <> 0 then
  fdmd.y = 10
  fdmd.height = 647
end if




' ***************************************************************************************************
' QRView flashers by iaakki

dim TablesDir : TablesDir = GetTablesFolder


dim bQRPairLoaded : bQRPairLoaded = false
dim preloadClaimStarted : preloadClaimStarted = False

'angle the QR code if in desktop
if DesktopMode then
  FScorbitQRIcon.height = 220
  FScorbitQR.height = 238
  FScorbitQRclaim.height = 238
  FScorbitQRIcon.rotx = -35
  FScorbitQR.rotx = -35
  FScorbitQRclaim.rotx = -35
end if

sub showQRPairImage
  if Scorbit.bNeedsPairing then
    FScorbitQR.visible = 0
    Dim sImageName
    sImageName = TablesDir & "\BMQR\QRcode.png"
    dim oWS, oExec
    set oWS = createobject("wscript.shell")
    dim command : command = """" & TablesDir & "\QRView.exe"" -i """ & sImageName & """"
'   msgbox command

    oExec = oWS.Run(command,4,false)

    FScorbitQR.timerenabled = true

    vpmtimer.addtimer 3000, "revealQR '"
  Else
    if not bQRPairloaded then
      'msgbox "too early?"
      vpmtimer.addtimer 200, "showQRPairImage '"
    Else
      if Not preloadClaimStarted then
'msgbox "Paired already, preload claim immediately"
        preloadQRClaim
  '     vpmtimer.addtimer 3000, "preloadQRClaim '"
      end if
      Introover=True  'enable game to start
    end if
  end if
end sub



sub preloadQRClaim
  preloadClaimStarted = True
' if preloadClaimStarted then exit sub
  Dim sImageName, fso
  sImageName = TablesDir & "\BMQR\QRclaim.png"

  Set fso = CreateObject("Scripting.FileSystemObject")

  If fso.FileExists(sImageName) then
    dim oWS2, oExec2
    set oWS2 = createobject("wscript.shell")
'msgbox "claim image"
    oExec2 = oWS2.Run("""" & TablesDir & "\QRView.exe"" -i """ & sImageName & """",4,false)

    FScorbitQRclaim.timerenabled = true
  Else
'msgbox "too early claim image?"
    vpmtimer.addtimer 1000, "preloadQRClaim '"
  end if
  Introover=True  'enable game to start
end sub

sub hideQR
  QRfadeCase = 0
  FScorbitQRIcon.timerenabled = true
' debug.print "preload claim2"

end sub


sub revealQR
  FScorbitQRIcon.imageA = "QRCodeS" 'scan to activate
  QRfadeCase = 1
  screen.visible = 1    'blocker
  screen001.visible = 0 'display
  FScorbitQR.opacity = 0
  FScorbitQR.visible = 1
  FScorbitQRIcon.opacity = 0
  FScorbitQRIcon.visible = 1
  FScorbitQRIcon.timerenabled = true
end sub


sub hideQRClaim
  QRfadeCase = 2
  FScorbitQRIcon.timerenabled = true
end sub

sub showQRClaimImage
  FScorbitQRIcon.imageA = "QRCodeB" 'scan to claim
  QRfadeCase = 3
  screen.visible = 1    'blocker
  screen001.visible = 0 'display
  FScorbitQRclaim.opacity = 0
  FScorbitQRclaim.visible = 1
  FScorbitQRIcon.opacity = 0
  FScorbitQRIcon.visible = 1
  FScorbitQRIcon.timerenabled = true
end sub

sub FScorbitQRIcon_timer  'used to fade the qr code in
' debug.print QRfadeCase

  select case QRfadeCase
    case 0: 'hideQR
        FScorbitQRIcon.opacity = 0.8 * FScorbitQRIcon.opacity - 1
        FScorbitQR.opacity = 10 * FScorbitQRIcon.opacity
    case 1: 'revealQR
        FScorbitQRIcon.opacity = 1.1 * FScorbitQRIcon.opacity + 5
        FScorbitQR.opacity = 10 * FScorbitQRIcon.opacity
    case 2: 'hideQRClaim
        FScorbitQRIcon.opacity = 0.8 * FScorbitQRIcon.opacity - 1
        FScorbitQRclaim.opacity = 10 * FScorbitQRIcon.opacity
    case 3: 'showQRClaimImage
        FScorbitQRIcon.opacity = 1.1 * FScorbitQRIcon.opacity + 5
        FScorbitQRclaim.opacity = 10 * FScorbitQRIcon.opacity
  end Select

  if FScorbitQRIcon.opacity <= 0 then
    FScorbitQR.opacity = 0
    FScorbitQRIcon.opacity = 0

    FScorbitQRIcon.visible = 0
    FScorbitQR.visible = 0

    screen.visible = 0    'blocker
    screen001.visible = 1 'display

    FScorbitQRIcon.timerenabled = false

    if QRfadeCase = 0 and Not preloadClaimStarted then
'     msgbox "claim load after pairing? Delay it a bit as new claim image may be in the works"
      preloadClaimStarted = true
      vpmtimer.addtimer 3000, "preloadQRClaim '"  'make it here as otherwise it can break the fading..
    end if

  Elseif FScorbitQRIcon.opacity >= 100 Then
    if QRfadeCase = 1 then
      FScorbitQR.opacity = 1000
    elseif QRfadeCase = 3 then
      FScorbitQRclaim.opacity = 1000
    end if
    FScorbitQRIcon.opacity = 100

    FScorbitQRIcon.timerenabled = false
  end if

end sub

FScorbitQR.VideoCapWidth=400
FScorbitQR.VideoCapHeight=400
FScorbitQRclaim.VideoCapWidth=400
FScorbitQRclaim.VideoCapHeight=400
dim QRfadeCase


Sub FScorbitQR_Timer()    'updates the QR image into a flasher
  FScorbitQR.VideoCapUpdate = "QRView"
  dim oWS3
  set oWS3 = createobject("wscript.shell")
  oWS3.Run "taskkill /im QRView.exe", , True

  FScorbitQR.timerenabled = False
  bQRPairloaded = true
End Sub

Sub FScorbitQRclaim_Timer()   'updates the QR claim image into a flasher
  FScorbitQRclaim.VideoCapUpdate = "QRView"
  dim oWS4
  set oWS4 = createobject("wscript.shell")
  oWS4.Run "taskkill /im QRView.exe", , True

  FScorbitQRclaim.timerenabled = False
  bQRPairloaded = true
  if not Scorbit.bNeedsPairing then PlaySound "scorbit_login",0,CalloutVol,0,0,1,1,1
End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  SCORBIT Interface
' To Use:
' 1) Define a timer tmrScorbit
' 2) Call DoInit at the end of PupInit or in Table Init if you are nto using pup with the appropriate parameters
'   if Scorbit.DoInit(389, "PupOverlays", "1.0.0", "GRWvz-MP37P") then
'     tmrScorbit.Interval=2000
'     tmrScorbit.UserValue = 0
'     tmrScorbit.Enabled=True
'   End if
' 3) Customize helper functions below for different events if you want or make your own
' 4) Call
'   StartSession - When a game starts
'   StopSession - When the game is over
'   SendUpdate - When Score Changes
'   SetGameMode - When different game events happen like starting a mode, MB etc.  (ScorbitBuildGameModes helper function shows you how)
' 5) Drop the binaries sQRCode.exe and sToken.exe in your Pup Root so we can create session kets and QRCodes.
' 6) Callbacks
'   Scorbit_Paired      - Called when machine is successfully paired.  Hide QRCode and play a sound
'   Scorbit_PlayerClaimed - Called when player is claimed.  Hide QRCode, play a sound and display name
'
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' TABLE CUSTOMIZATION START HERE

Sub Scorbit_Paired()                ' Scorbit callback when new machine is paired
'debug.print "Scorbit PAIRED"
' PlaySoundVol "scorbit_login", VolSfx
' PuPlayer.LabelSet pOverVid, "ScorbitQR", "PuPOverlays\\clear.png",0,""
' PuPlayer.LabelSet pOverVid, "ScorbitQRIcon", "PuPOverlays\\clear.png",0,""
  'make scorbit disappear here
  hideQR
  bQRPairLoaded = true
End Sub

Sub Scorbit_PlayerClaimed(PlayerNum, PlayerName)  ' Scorbit callback when QR Is Claimed
'debug.print "Scorbit LOGIN: " & PlayerNum & " - " & PlayerName
  PlaySound "scorbit_login",0,CalloutVol,0,0,1,1,1
  ScorbitClaimQR(False)
  Scorepop.enabled=True
  Scorepop.interval=3000
  puPlayer.LabelSet pDMD,"Scorlogo","Gif\\sbt.gif",1,"{'mt':2,'color':111111, 'width': 35, 'height': 30., 'anigif': 130 ,}"
  puPlayer.LabelSet pDMD,"Player","" & PlayerName ,1,"{'mt':2, 'shadowcolor':10646039, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }"
  'PUPtodo
  ShowBallCount False
  DMD_ShowText PlayerName & " claimed player " & PlayerNum,4,FontScoreInactiv1,30,True,40,2500
  vpmtimer.addtimer 3000, "ShowBallCount True '"
' msgbox PlayerName
End Sub


Sub ScorbitClaimQR(bShow)           '  Show QRCode on first ball for users to claim this position
  if Scorbit.bSessionActive=False then Exit Sub
  if ScorbitShowClaimQR=False then Exit Sub
  if Scorbit.bNeedsPairing then exit sub

' msgbox "Currentplayer and name: " & CurrentPlayer &" - "& Scorbit.GetName(CurrentPlayer)
  if bShow and balls=1 and bGameInPlay and Scorbit.GetName(CurrentPlayer)="" then
'   if PupOption=0 or ScorbitClaimSmall=0 then ' Desktop Make it Larger
'     PuPlayer.LabelSet pDMDText, "ScorbitQR", "PuPOverlays\\QRclaim.png",1,"{'mt':2,'width':20, 'height':40,'xalign':0,'yalign':0,'ypos':40,'xpos':5}"
'     PuPlayer.LabelSet pDMDText, "ScorbitQRIcon", "PuPOverlays\\QRcodeB.png",1,"{'mt':2,'width':23, 'height':52,'xalign':0,'yalign':0,'ypos':38,'xpos':3.5,'zback':1}"
'   else
'     PuPlayer.LabelSet pDMDText, "ScorbitQR", "PuPOverlays\\QRclaim.png",1,"{'mt':2,'width':12, 'height':24,'xalign':0,'yalign':0,'ypos':60,'xpos':5}"
'     PuPlayer.LabelSet pDMDText, "ScorbitQRIcon", "PuPOverlays\\QRcodeB.png",1,"{'mt':2,'width':14, 'height':32.5,'xalign':0,'yalign':0,'ypos':58,'xpos':4,'zback':1}"
'   End if
    'Show generated QRclaim and qrcodeB here
    showQRClaimImage
  Else
'   PuPlayer.LabelSet pDMDText, "ScorbitQR", "PuPOverlays\\clear.png",0,""
'   PuPlayer.LabelSet pDMDText, "ScorbitQRIcon", "PuPOverlays\\clear.png",0,""
'   debug.print "make scorbit disappear here"
    hideQRClaim
  End if
End Sub

Sub ScorbitBuildGameModes()   ' Custom function to build the game modes for better stats
  dim GameModeStr
  if Scorbit.bSessionActive=False then Exit Sub
  GameModeStr="NA:"

  Select Case CurrentMission
    case 0:' No Mode Selected
    Case 1: ' THE HUNT
      GameModeStr="NA{green}:The Hunt"
        Case 2: ' RESURRECTION
      GameModeStr="NA{red}:Resurrection"
        Case 3: ' CHASE
      GameModeStr="NA{blue}:Chase"
        Case 4: ' CEMETERY
      GameModeStr="NA{purple}:Cemetery"
        Case 5: ' PORTAL ROOM
      GameModeStr="NA{red}:PortalRoom"
        Case 6: ' HEARTOFSTEEL
      GameModeStr="NA{purple}:Heart of Steel"
        Case 7: ' BETRAYAL
      GameModeStr="NA{blue}:Betrayal"
  End Select

  'MB's
  if bCoreyMBOngoing then
    if GameModeStr<>"" then GameModeStr=GameModeStr & ";"
    GameModeStr=GameModeStr&"MB{green}:Corey"
  End if
  if bTracyMBOngoing then
    if GameModeStr<>"" then GameModeStr=GameModeStr & ";"
    GameModeStr=GameModeStr&"MB{purple}:Tracy"
  End if
  if bMimaMBOngoing then
    if GameModeStr<>"" then GameModeStr=GameModeStr & ";"
    GameModeStr=GameModeStr&"MB{purple}:Mima"
  End if

  'Wizard
  if WizardPhase = 1 then
    if GameModeStr<>"" then GameModeStr=GameModeStr & ";"
    GameModeStr=GameModeStr&"WM{black}:Lockdown"
  End if

  if WizardPhase = 2 then
    if GameModeStr<>"" then GameModeStr=GameModeStr & ";"
    GameModeStr=GameModeStr&"WM{red}:Rewind Tracy"
  End if

  if WizardPhase = 3 then
    if GameModeStr<>"" then GameModeStr=GameModeStr & ";"
    GameModeStr=GameModeStr&"WM{white}:Supernova"
  End if

  if WizardPhase = 4 then
    if GameModeStr<>"" then GameModeStr=GameModeStr & ";"
    GameModeStr=GameModeStr&"WM{red}:Battle Galdor"
  End if

  if WizardPhase = 5 then
    if GameModeStr<>"" then GameModeStr=GameModeStr & ";"
    GameModeStr=GameModeStr&"WM{orange}:Galdor Defeated"
  End if

  Scorbit.SetGameMode(GameModeStr)

End Sub

Sub Scorbit_LOGUpload(state)  ' Callback during the log creation process.  0=Creating Log, 1=Uploading Log, 2=Done
  Select Case state
    case 0:
      Debug.print "CREATING LOG"
    case 1:
      Debug.print "Uploading LOG"
    case 2:
      Debug.print "LOG Complete"
  End Select
End Sub
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' TABLE CUSTOMIZATION END HERE - NO NEED TO EDIT BELOW THIS LINE


dim Scorbit : Set Scorbit = New ScorbitIF
' Workaround - Call get a reference to Member Function
Sub tmrScorbit_Timer()                ' Timer to send heartbeat
  Scorbit.DoTimer(tmrScorbit.UserValue)
  tmrScorbit.UserValue=tmrScorbit.UserValue+1
  if tmrScorbit.UserValue>5 then tmrScorbit.UserValue=0
End Sub
Function ScorbitIF_Callback()
  Scorbit.Callback()
End Function
Class ScorbitIF

  Public bSessionActive
  Public bNeedsPairing
  Private bUploadLog
  Private bActive
  Private LOGFILE(10000000)
  Private LogIdx

  Private bProduction

  Private TypeLib
  Private MyMac
  Private Serial
  Private MyUUID
  Private TableVersion

  Private SessionUUID
  Private SessionSeq
  Private SessionTimeStart
  Private bRunAsynch
  Private bWaitResp
  Private GameMode
  Private GameModeOrig    ' Non escaped version for log
  Private VenueMachineID
  Private CachedPlayerNames(4)
  Private SaveCurrentPlayer

  Public bEnabled
  Private sToken
  Private machineID
  Private dirQRCode
  Private opdbID
  Private wsh

  Private objXmlHttpMain
  Private objXmlHttpMainAsync
  Private fso
  Private Domain

  Public Sub Class_Initialize()
    bActive="false"
    bSessionActive=False
    bEnabled=False
  End Sub

  Property Let UploadLog(bValue)
    bUploadLog = bValue
  End Property

  Sub DoTimer(bInterval)  ' 2 second interval
    dim holdScores(4)
    dim i

    if bInterval=0 then
      SendHeartbeat()
    elseif bRunAsynch then ' Game in play
      Scorbit.SendUpdate Score(1), Score(2), Score(3), Score(4), Balls, CurrentPlayer, PlayersPlayingGame
    End if
  End Sub

  Function GetName(PlayerNum) ' Return Parsed Players name
    if PlayerNum<1 or PlayerNum>4 then
      GetName=""
    else
      GetName=CachedPlayerNames(PlayerNum)
    End if
  End Function

  Function DoInit(MyMachineID, Directory_PupQRCode, Version, opdb)
    dim Nad
    Dim EndPoint
    Dim resultStr
    Dim UUIDParts
    Dim UUIDFile

    bProduction=1
'   bProduction=0
    SaveCurrentPlayer=0
    VenueMachineID=""
    bWaitResp=False
    bRunAsynch=False
    DoInit=False
    opdbID=opdb
    dirQrCode=Directory_PupQRCode
    MachineID=MyMachineID
    TableVersion=version
    bNeedsPairing=False
    if bProduction then
      domain = "api.scorbit.io"
    else
      domain = "staging.scorbit.io"
      domain = "scorbit-api-staging.herokuapp.com"
    End if
'   msgbox "production: " & bProduction
    Set fso = CreateObject("Scripting.FileSystemObject")
    dim objLocator:Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Dim objService:Set objService = objLocator.ConnectServer(".", "root\cimv2")
    Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
    Set objXmlHttpMainAsync = CreateObject("Microsoft.XMLHTTP")
    objXmlHttpMain.onreadystatechange = GetRef("ScorbitIF_Callback")
    Set wsh = CreateObject("WScript.Shell")

    ' Get Mac for Serial Number
    dim Nads: set Nads = objService.ExecQuery("Select * from Win32_NetworkAdapter where physicaladapter=true")
    for each Nad in Nads
      if not isnull(Nad.MACAddress) then
'       msgbox "Using MAC Addresses:" & Nad.MACAddress & " From Adapter:" & Nad.description
        MyMac=replace(Nad.MACAddress, ":", "")
        Exit For
      End if
    Next
    Serial=eval("&H" & mid(MyMac, 5))
'   Serial=123456
'   debug.print "Serial: " & Serial
    serial = serial + MachineID
'   debug.print "New Serial with machine id: " & Serial

    ' Get System UUID
    set Nads = objService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
    for each Nad in Nads
'     msgbox "Using UUID:" & Nad.UUID
      MyUUID=Nad.UUID
      Exit For
    Next

    if MyUUID="" then
      MsgBox "SCORBIT - Can get UUID, Disabling."
      Exit Function
    elseif MyUUID="03000200-0400-0500-0006-000700080009" or ScorbitAlternateUUID then
      If fso.FolderExists(UserDirectory) then
        If fso.FileExists(UserDirectory & "ScorbitUUID.dat") then
          Set UUIDFile = fso.OpenTextFile(UserDirectory & "ScorbitUUID.dat",1)
          MyUUID = UUIDFile.ReadLine()
          UUIDFile.Close
          Set UUIDFile = Nothing
        Else
          MyUUID=GUID()
          Set UUIDFile=fso.CreateTextFile(UserDirectory & "ScorbitUUID.dat",True)
          UUIDFile.WriteLine MyUUID
          UUIDFile.Close
          Set UUIDFile=Nothing
    End if
      End if
    End if

    ' Clean UUID
    UUIDParts=split(MyUUID, "-")
'   msgbox UUIDParts(0)
    MyUUID=LCASE(Hex(eval("&h" & UUIDParts(0))+MyMachineID)) & UUIDParts(1) &  UUIDParts(2) &  UUIDParts(3) & UUIDParts(4)     ' Add MachineID to UUID
'   msgbox UUIDParts(0)
    MyUUID=LPad(MyUUID, 32, "0")
'   MyUUID=Replace(MyUUID, "-",  "")
'   Debug.print "MyUUID:" & MyUUID


' Debug
'   myUUID="adc12b19a3504453a7414e722f58737f"
'   Serial="123456778"

    'create own folder for table QRCodes TablesDir & "\" & dirQrCode
    If Not fso.FolderExists(TablesDir & "\" & dirQrCode) then
      fso.CreateFolder(TablesDir & "\" & dirQrCode)
    end if

    ' Authenticate and get our token
    if getStoken() then
      bEnabled=True
'     SendHeartbeat
      DoInit=True
    End if
  End Function



  Sub Callback()
    Dim ResponseStr
    Dim i
    Dim Parts
    Dim Parts2
    Dim Parts3
    if bEnabled=False then Exit Sub

    if bWaitResp and objXmlHttpMain.readystate=4 then
'     Debug.print "CALLBACK: " & objXmlHttpMain.Status & " " & objXmlHttpMain.readystate
      if objXmlHttpMain.Status=200 and objXmlHttpMain.readystate = 4 then
        ResponseStr=objXmlHttpMain.responseText
'       Debug.print "RESPONSE: " & ResponseStr

        ' Parse Name
        if CachedPlayerNames(SaveCurrentPlayer-1)="" then  ' Player doesnt have a name
          if instr(1, ResponseStr, "cached_display_name") <> 0 Then ' There are names in the result
            Parts=Split(ResponseStr,",{")             ' split it
            if ubound(Parts)>=SaveCurrentPlayer-1 then        ' Make sure they are enough avail
              if instr(1, Parts(SaveCurrentPlayer-1), "cached_display_name")<>0 then  ' See if mine has a name
                CachedPlayerNames(SaveCurrentPlayer-1)=GetJSONValue(Parts(SaveCurrentPlayer-1), "cached_display_name")    ' Get my name
                CachedPlayerNames(SaveCurrentPlayer-1)=Replace(CachedPlayerNames(SaveCurrentPlayer-1), """", "")
                Scorbit_PlayerClaimed SaveCurrentPlayer, CachedPlayerNames(SaveCurrentPlayer-1)
'               Debug.print "Player Claim:" & SaveCurrentPlayer & " " & CachedPlayerNames(SaveCurrentPlayer-1)
              End if
            End if
          End if
        else                            ' Check for unclaim
          if instr(1, ResponseStr, """player"":null")<>0 Then ' Someone doesnt have a name
            Parts=Split(ResponseStr,"[")            ' split it
'Debug.print "Parts:" & Parts(1)
            Parts2=Split(Parts(1),"}")              ' split it
            for i = 0 to Ubound(Parts2)
'Debug.print "Parts2:" & Parts2(i)
              if instr(1, Parts2(i), """player"":null")<>0 Then
                CachedPlayerNames(i)=""
              End if
            Next
          End if
        End if
      End if
      bWaitResp=False
    End if
  End Sub



  Public Sub StartSession()
    if bEnabled=False then Exit Sub
'msgbox  "Scorbit Start Session"
    CachedPlayerNames(0)=""
    CachedPlayerNames(1)=""
    CachedPlayerNames(2)=""
    CachedPlayerNames(3)=""
    bRunAsynch=True
    bActive="true"
    bSessionActive=True
    SessionSeq=0
    SessionUUID=GUID()
    SessionTimeStart=GameTime
    LogIdx=0
    SendUpdate 0, 0, 0, 0, 1, 1, 1
  End Sub

  Public Sub StopSession(P1Score, P2Score, P3Score, P4Score, NumberPlayers)
    StopSession2 P1Score, P2Score, P3Score, P4Score, NumberPlayers, False
  End Sub

  Public Sub StopSession2(P1Score, P2Score, P3Score, P4Score, NumberPlayers, bCancel)
    Dim i
    dim objFile
    if bEnabled=False then Exit Sub
    if bSessionActive=False then Exit Sub
Debug.print "Scorbit Stop Session"

    bRunAsynch=False
    bActive="false"
    bSessionActive=False
    SendUpdate P1Score, P2Score, P3Score, P4Score, -1, -1, NumberPlayers
'   SendHeartbeat

    if bUploadLog and LogIdx<>0 and bCancel=False then
      Debug.print "Creating Scorbit Log: Size" & LogIdx
      Scorbit_LOGUpload(0)
'     Set objFile = fso.CreateTextFile(puplayer.getroot&"\" & cGameName & "\sGameLog.csv")
      Set objFile = fso.CreateTextFile(TablesDir & "\BMQR\sGameLog.csv")
      For i = 0 to LogIdx-1
        objFile.Writeline LOGFILE(i)
      Next
      objFile.Close
      LogIdx=0
      Scorbit_LOGUpload(1)
'     pvPostFile "https://" & domain & "/api/session_log/", puplayer.getroot&"\" & cGameName & "\sGameLog.csv", False
      pvPostFile "https://" & domain & "/api/session_log/", TablesDir & "\BMQR\sGameLog.csv", False
      Scorbit_LOGUpload(2)
      on error resume next
'     fso.DeleteFile(puplayer.getroot&"\" & cGameName & "\sGameLog.csv")
      fso.DeleteFile(TablesDir & "\BMQR\sGameLog.csv")
      on error goto 0
    End if

  End Sub

  Public Sub SetGameMode(GameModeStr)
    GameModeOrig=GameModeStr
    GameMode=GameModeStr
    GameMode=Replace(GameMode, ":", "%3a")
    GameMode=Replace(GameMode, ";", "%3b")
    GameMode=Replace(GameMode, " ", "%20")
    GameMode=Replace(GameMode, "{", "%7B")
    GameMode=Replace(GameMode, "}", "%7D")
  End sub

  Public Sub SendUpdate(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers)
    SendUpdateAsynch P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bRunAsynch
  End Sub

  Public Sub SendUpdateAsynch(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bAsynch)
    dim i
    Dim PostData
    Dim resultStr
    dim LogScores(4)

    if bUploadLog then
      if NumberPlayers>=1 then LogScores(0)=P1Score
      if NumberPlayers>=2 then LogScores(1)=P2Score
      if NumberPlayers>=3 then LogScores(2)=P3Score
      if NumberPlayers>=4 then LogScores(3)=P4Score
      LOGFILE(LogIdx)=DateDiff("S", "1/1/1970", Now()) & "," & LogScores(0) & "," & LogScores(1) & "," & LogScores(2) & "," & LogScores(3) & ",,," &  CurrentPlayer & "," & CurrentBall & ",""" & GameModeOrig & """"
      LogIdx=LogIdx+1
    End if

    if bEnabled=False then Exit Sub
    if bWaitResp then exit sub ' Drop message until we get our next response
'   debug.print "Current players: " & CurrentPlayer
    SaveCurrentPlayer=CurrentPlayer
'   PostData = "session_uuid=" & SessionUUID & "&session_time=" & DateDiff("S", "1/1/1970", Now()) & _
'         "&session_sequence=" & SessionSeq & "&active=" & bActive
    PostData = "session_uuid=" & SessionUUID & "&session_time=" & GameTime-SessionTimeStart+1 & _
          "&session_sequence=" & SessionSeq & "&active=" & bActive

    SessionSeq=SessionSeq+1
    if NumberPlayers > 0 then
      for i = 0 to NumberPlayers-1
        PostData = PostData & "&current_p" & i+1 & "_score="
        if i <= NumberPlayers-1 then
                    if i = 0 then PostData = PostData & P1Score
                    if i = 1 then PostData = PostData & P2Score
                    if i = 2 then PostData = PostData & P3Score
                    if i = 3 then PostData = PostData & P4Score
        else
          PostData = PostData & "-1"
        End if
      Next

      PostData = PostData & "&current_ball=" & CurrentBall & "&current_player=" & CurrentPlayer
      if GameMode<>"" then PostData=PostData & "&game_modes=" & GameMode
    End if
    resultStr = PostMsg("https://" & domain, "/api/entry/", PostData, bAsynch)
    if resultStr<>"" then Debug.print "SendUpdate Resp:" & resultStr
  End Sub

' PRIVATE BELOW
  Private Function LPad(StringToPad, Length, CharacterToPad)
    Dim x : x = 0
    If Length > Len(StringToPad) Then x = Length - len(StringToPad)
    LPad = String(x, CharacterToPad) & StringToPad
  End Function

  Private Function GUID()
    Dim TypeLib
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    GUID = Mid(TypeLib.Guid, 2, 36)

'   Set wsh = CreateObject("WScript.Shell")
'   Set fso = CreateObject("Scripting.FileSystemObject")
'
'   dim rc
'   dim result
'   dim objFileToRead
'   Dim sessionID:sessionID=puplayer.getroot&"\" & cGameName & "\sessionID.txt"
'
'   on error resume next
'   fso.DeleteFile(sessionID)
'   On error goto 0
'
'   rc = wsh.Run("powershell -Command ""(New-Guid).Guid"" | out-file -encoding ascii " & sessionID, 0, True)
'   if FileExists(sessionID) and rc=0 then
'     Set objFileToRead = fso.OpenTextFile(sessionID,1)
'     result = objFileToRead.ReadLine()
'     objFileToRead.Close
'     GUID=result
'   else
'     MsgBox "Cant Create SessionUUID through powershell. Disabling Scorbit"
'     bEnabled=False
'   End if

  End Function

  Private Function GetJSONValue(JSONStr, key)
    dim i
    Dim tmpStrs,tmpStrs2
    if Instr(1, JSONStr, key)<>0 then
      tmpStrs=split(JSONStr,",")
      for i = 0 to ubound(tmpStrs)
        if instr(1, tmpStrs(i), key)<>0 then
          tmpStrs2=split(tmpStrs(i),":")
          GetJSONValue=tmpStrs2(1)
          exit for
        End if
      Next
    End if
  End Function

  Private Sub SendHeartbeat()
    Dim resultStr
    dim TmpStr
    Dim Command
    Dim rc
'   Dim QRFile:QRFile=puplayer.getroot&"\" & cGameName & "\" & dirQrCode
    Dim QRFile:QRFile=TablesDir & "\" & dirQrCode
    if bEnabled=False then Exit Sub
    resultStr = GetMsgHdr("https://" & domain, "/api/heartbeat/", "Authorization", "SToken " & sToken)
'Debug.print "Heartbeat Resp:" & resultStr
    If VenueMachineID="" then

      if resultStr<>"" and Instr(resultStr, """unpaired"":true")=0 then   ' We Paired
        bNeedsPairing=False
        debug.print "Paired"
        Scorbit_Paired()
      else
        debug.print "Needs Pairing"
        bNeedsPairing=True
'       if not FScorbitQRIcon.visible then showQRPairImage
      End if

      TmpStr=GetJSONValue(resultStr, "venuemachine_id")
      if TmpStr<>"" then
        VenueMachineID=TmpStr
'Debug.print "VenueMachineID=" & VenueMachineID
'       Command = puplayer.getroot&"\" & cGameName & "\sQRCode.exe " & VenueMachineID & " " & opdbID & " " & QRFile
        debug.print "RUN sqrcode"
        Command = """" & TablesDir & "\sQRCode.exe"" " & VenueMachineID & " " & opdbID & " """ & QRFile & """"
'       msgbox Command
        rc = wsh.Run(Command, 0, False)
      End if
    End if
  End Sub

  Private Function getStoken()
    Dim result
    Dim results
'   dim wsh
    Dim tmpUUID:tmpUUID="adc12b19a3504453a7414e722f58736b"
    Dim tmpVendor:tmpVendor="vscorbitron"
    Dim tmpSerial:tmpSerial="999990104"
'   Dim QRFile:QRFile=puplayer.getroot&"\" & cGameName & "\" & dirQrCode
    Dim QRFile:QRFile=TablesDir & "\" & dirQrCode
'   Dim sTokenFile:sTokenFile=puplayer.getroot&"\" & cGameName & "\sToken.dat"
    Dim sTokenFile:sTokenFile=TablesDir & "\sToken.dat"

    ' Set everything up
    tmpUUID=MyUUID
    tmpVendor="vpin"
    tmpSerial=Serial

    on error resume next
    fso.DeleteFile("""" & sTokenFile & """")
    On error goto 0

    ' get sToken and generate QRCode
'   Set wsh = CreateObject("WScript.Shell")
    Dim waitOnReturn: waitOnReturn = True
    Dim windowStyle: windowStyle = 0
    Dim Command
    Dim rc
    Dim objFileToRead

'   msgbox """" & " 55"

'   Command = puplayer.getroot&"\" & cGameName & "\sToken.exe " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " " & QRFile & " " & sTokenFile & " " & domain
    debug.print "RUN sToken"
    Command = """" & TablesDir & "\sToken.exe"" " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " """ & QRFile & """ """ & sTokenFile & """ " & domain
'msgbox "RUNNING:" & Command
    rc = wsh.Run(Command, windowStyle, waitOnReturn)
'msgbox "Return:" & rc
'   if FileExists(puplayer.getroot&"\" & cGameName & "\sToken.dat") and rc=0 then
'   msgbox """" & TablesDir & "\sToken.dat"""
    if FileExists(TablesDir & "\sToken.dat") and rc=0 then
'     Set objFileToRead = fso.OpenTextFile(puplayer.getroot&"\" & cGameName & "\sToken.dat",1)
'     msgbox """" & TablesDir & "\sToken.dat"""
      Set objFileToRead = fso.OpenTextFile(TablesDir & "\sToken.dat",1)
      result = objFileToRead.ReadLine()
      objFileToRead.Close
      Set objFileToRead = Nothing
'msgbox result

      if Instr(1, result, "Invalid timestamp")<> 0 then
        MsgBox "Scorbit Timestamp Error: Please make sure the time on your system is exact"
        getStoken=False
      elseif Instr(1, result, "Internal Server error")<> 0 then
        MsgBox "Scorbit Internal Server error ??"
        getStoken=False
      elseif Instr(1, result, ":")<>0 then
        results=split(result, ":")
        sToken=results(1)
        sToken=mid(sToken, 3, len(sToken)-4)
Debug.print "Got TOKEN:" & sToken
        getStoken=True
      Else
Debug.print "ERROR:" & result
        getStoken=False
      End if
    else
'msgbox "ERROR No File:" & rc
    End if

  End Function

  private Function FileExists(FilePath)
    If fso.FileExists(FilePath) Then
      FileExists=CBool(1)
    Else
      FileExists=CBool(0)
    End If
  End Function

  Private Function GetMsg(URLBase, endpoint)
    GetMsg = GetMsgHdr(URLBase, endpoint, "", "")
  End Function

  Private Function GetMsgHdr(URLBase, endpoint, Hdr1, Hdr1Val)
    Dim Url
    Url = URLBase + endpoint & "?session_active=" & bActive
'Debug.print "Url:" & Url  & "  Async=" & bRunAsynch
    objXmlHttpMain.open "GET", Url, bRunAsynch
'   objXmlHttpMain.setRequestHeader "Content-Type", "text/xml"
    objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
    if Hdr1<> "" then objXmlHttpMain.setRequestHeader Hdr1, Hdr1Val

'   on error resume next
      err.clear
      objXmlHttpMain.send ""
      if err.number=-2147012867 then
        MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
        bEnabled=False
      elseif err.number <> 0 then
        debug.print "Server error: (" & err.number & ") " & Err.Description
      End if
      if bRunAsynch=False then
          Debug.print "Status: " & objXmlHttpMain.status
          If objXmlHttpMain.status = 200 Then
            GetMsgHdr = objXmlHttpMain.responseText
        Else
            GetMsgHdr=""
          End if
      Else
        bWaitResp=True
        GetMsgHdr=""
      End if
'   On error goto 0

  End Function

  Private Function PostMsg(URLBase, endpoint, PostData, bAsynch)
    Dim Url

    Url = URLBase + endpoint
'debug.print "PostMSg:" & Url & " " & PostData

    objXmlHttpMain.open "POST",Url, bAsynch
    objXmlHttpMain.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXmlHttpMain.setRequestHeader "Content-Length", Len(PostData)
    objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
    objXmlHttpMain.setRequestHeader "Authorization", "SToken " & sToken
    if bAsynch then bWaitResp=True

    on error resume next
      objXmlHttpMain.send PostData
      if err.number=-2147012867 then
        MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
        bEnabled=False
      elseif err.number <> 0 then
        'debug.print "Multiplayer Server error (" & err.number & ") " & Err.Description
      End if
      If objXmlHttpMain.status = 200 Then
        PostMsg = objXmlHttpMain.responseText
      else
        PostMsg="ERROR: " & objXmlHttpMain.status & " >" & objXmlHttpMain.responseText & "<"
      End if
    On error goto 0
  End Function

  Private Function pvPostFile(sUrl, sFileName, bAsync)
Debug.print "Posting File " & sUrl & " " & sFileName & " " & bAsync & " File:" & Mid(sFileName, InStrRev(sFileName, "\") + 1)
    Dim STR_BOUNDARY:STR_BOUNDARY  = GUID()
    Dim nFile
    Dim baBuffer()
    Dim sPostData
    Dim Response

    '--- read file
    Set nFile = fso.GetFile(sFileName)
    With nFile.OpenAsTextStream()
      sPostData = .Read(nFile.Size)
      .Close
    End With
'   fso.Open sFileName For Binary Access Read As nFile
'   If LOF(nFile) > 0 Then
'     ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
'     Get nFile, , baBuffer
'     sPostData = StrConv(baBuffer, vbUnicode)
'   End If
'   Close nFile

    '--- prepare body
    sPostData = "--" & STR_BOUNDARY & vbCrLf & _
      "Content-Disposition: form-data; name=""uuid""" & vbCrLf & vbCrLf & _
      SessionUUID & vbcrlf & _
      "--" & STR_BOUNDARY & vbCrLf & _
      "Content-Disposition: form-data; name=""log_file""; filename=""" & SessionUUID & ".csv""" & vbCrLf & _
      "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
      sPostData & vbCrLf & _
      "--" & STR_BOUNDARY & "--"

'Debug.print "POSTDATA: " & sPostData & vbcrlf

    '--- post
    With objXmlHttpMain
      .Open "POST", sUrl, bAsync
      .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
      .SetRequestHeader "Authorization", "SToken " & sToken
      .Send sPostData ' pvToByteArray(sPostData)
      If Not bAsync Then
        Response= .ResponseText
        pvPostFile = Response
Debug.print "Upload Response: " & Response
      End If
    End With

  End Function

  Private Function pvToByteArray(sText)
    pvToByteArray = StrConv(sText, 128)   ' vbFromUnicode
  End Function

End Class
'  END SCORBIT
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


''QRView support by iaakki

Function GetTablesFolder()
    Dim GTF
    Set GTF = CreateObject("Scripting.FileSystemObject")
    GetTablesFolder= GTF.GetParentFolderName(userdirectory) & "\Tables"
    set GTF = nothing
End Function



if Scorbitactive = 1 then
  ScorbitExesCheck
' msgbox "doinit"
' if Scorbit.DoInit(2082, "PupOverlays", myVersion, "GRWvz-MP37P") then   ' Staging
  if Scorbit.DoInit(2115, "BMQR", myVersion, "ESDis-GREAT") then  ' Prod
    tmrScorbit.Interval=2000
    tmrScorbit.UserValue = 0
    tmrScorbit.Enabled=True
    Scorbit.UploadLog = ScorbitUploadLog
  End if

End if

' Checks that all needed binaries are available in correct place..
Sub ScorbitExesCheck
  dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")

  If fso.FileExists(TablesDir & "\sToken.exe") then
'   msgbox "Stoken.exe found at: " & TablesDir & "\sToken.exe"
  else
    msgbox "Stoken.exe NOT found at: " & TablesDir & "\sToken.exe Disabling Scorbit for now."
    Scorbitactive = 0
  end if

  If fso.FileExists(TablesDir & "\sQRCode.exe") then
'   msgbox "sQRCode.exe found at: " & TablesDir & "\sQRCode.exe"
  else
    msgbox "sQRCode.exe NOT found at: " & TablesDir & "\sQRCode.exe Disabling Scorbit for now."
    Scorbitactive = 0
  end if

  If fso.FileExists(TablesDir & "\qrview.exe") then
'   msgbox "qrview.exe found at: " & TablesDir & "\qrview.exe"
  else
    msgbox "qrview.exe NOT found at: " & TablesDir & "\qrview.exe Disabling Scorbit for now."
    Scorbitactive = 0
  end if

end sub

'*****************************************
' iaakki's  Blood Trails
'*****************************************
Dim BloodTrails, xxx

dim objBlood(19)

for xxx = 0 to 19
' msgbox xxx
' Eval("BallShadow" & xxx+1).image = "arrow"
  Set objBlood(xxx) = Eval("BloodTrail0" & xxx)
  objBlood(xxx).image = "blood" & RndInt(1,6)
' objBlood(xxx).image = "arrow"
  objBlood(xxx).z = 0 + xxx/100
  objBlood(xxx).depthbias = -10 - xxx + 100 'hack to make it go under the GI flasher
  objBlood(xxx).uservalue = 0
next

dim Bcnt : Bcnt = 0
dim prevBcnt, rotateCnt

'dim TableSlope
'TableSlope = table1.SlopeMin + (table1.SlopeMax - table1.SlopeMin) * table1.GlobalDifficulty

Sub solBloodTrail(Enabled)  'make sure that no debris is left on PF after the mode.
  if Enabled then
    for xxx = 0 to 19
      objBlood(xxx).x = 685: objBlood(xxx).y = 2234
      objBlood(xxx).visible = true
      objBlood(xxx).z = 0 + xxx/100
      objBlood(xxx).depthbias = -10 - xxx + 100
      objBlood(xxx).uservalue = 0
    next
    BloodTrailUpdate.enabled = true
  Else
    BloodTrailUpdate.enabled = false
    for xxx = 0 to 19
      objBlood(xxx).x = 685: objBlood(xxx).y = 2234
      objBlood(xxx).visible = false
      objBlood(xxx).z = 0 + xxx/100
      objBlood(xxx).depthbias = -10 - xxx + 100
      objBlood(xxx).uservalue = 0
    next
  end if

end sub

Sub BloodTrailUpdate_timer()
    Dim BOT, b, currentBallVelx, currentBallVely, currentBallVel, currentBallZ
  if CurrentMission = 5 Or bWizardMode then exit Sub 'disable for PR and Wizard, but enable for MB

    BOT = GetBalls

  For b = 0 to UBound(BOT)
    currentBallVel = BallVel(BOT(b))
    if currentBallVel > 1 then
      currentBallZ = BOT(b).z
      'move
      objBlood(Bcnt).X = BOT(b).X
      objBlood(Bcnt).Y = BOT(b).Y

      'make it thinner by time and make them fall unless on ramp. Use uservalue for ramp status
      dim CurrentFallValue
      for xxx = 0 to 19
        if Not xxx = Bcnt then
          CurrentFallValue = objBlood(xxx).uservalue
          objBlood(xxx).size_x = objBlood(xxx).size_x * 0.95
          if CurrentFallValue > 0 Then
            CurrentFallValue = CurrentFallValue - 30
            if CurrentFallValue < 0 then CurrentFallValue = 0

            objBlood(xxx).rotx = objBlood(xxx).rotx / 2
'           objBlood(xxx).rotx = 0
            objBlood(xxx).z = CurrentFallValue + xxx/100

            objBlood(xxx).uservalue = CurrentFallValue
            if CurrentFallValue = 0 then
              objBlood(xxx).rotx = 0
            end if
          Elseif currentBallZ < 0 then
            if InRect(objBlood(xxx).x, objBlood(xxx).y, 212,300,311,274,351,367,247,406) then
              'vortex sink
              CurrentFallValue = CurrentFallValue - 10
              objBlood(xxx).z = CurrentFallValue + xxx/100
              objBlood(xxx).uservalue = CurrentFallValue
            end if
          end if
        end if
      next


      'resize by speed
      if currentBallVel > 50 then currentBallVel = 50
      objBlood(Bcnt).size_x = 5 - 2*(currentBallVel/50)
      objBlood(Bcnt).size_y = 5 + (currentBallVel/8)

      'ball direction
      objBlood(Bcnt).objrotz = Atn2(BOT(b).vely,BOT(b).velx)/(PI/180) - 90

      'ramp
      if currentBallZ > 27 then
        objBlood(Bcnt).z = currentBallZ - 24 + Bcnt/100
        objBlood(Bcnt).depthbias = -10 - Bcnt - currentBallZ + 100 'dynamic depth bias

        'ramp angle
        objBlood(Bcnt).rotx = (Atn(BOT(b).velz / SQR(BOT(b).velx*BOT(b).velx+BOT(b).vely*BOT(b).vely)))/(PI/180) - 3


        if InRect(BOT(b).x, BOT(b).y, 0,744,126,709,303,1145,0,1123) Or (InRect(BOT(b).x, BOT(b).y,614,319,723,178,725,1098,566,1065) And currentBallZ < 140) or InRect(BOT(b).x, BOT(b).y, 80,16,500,0,723,180,613,316) or InRect(BOT(b).x, BOT(b).y, 875,842,951,843,962,1392,870,1389) Then 'solid ramps here
'         debug.print "ramp!"
          objBlood(Bcnt).uservalue = 0
        Else  'not on a solid ramp
'         debug.print "not solid ramp"
          objBlood(Bcnt).uservalue = objBlood(Bcnt).z 'fall amount
        end if
      else
        objBlood(Bcnt).z = Bcnt/100
        objBlood(Bcnt).rotx = 0
        objBlood(Bcnt).uservalue = 0
        objBlood(Bcnt).depthbias = -10 - Bcnt + 100
      end if

      prevBcnt = Bcnt
      Bcnt = Bcnt + 1
      if Bcnt > 19 Then
        Bcnt = 0
      end if
    end if
  Next
End Sub


'025 iaakki - Teleports pop up when you shoot through small ramp under main Ramp, Gun diverter opens when you shoot loop ramp once, gun diverter closes when ball is directed towards gun, speedup added to loopramp
'028 iaakki - right arch reworked, rubbers added, some walls aligned to return better
'029 iaakki - merge of two 028 versions
'030 iaakki - autoplunger functionality added. It won't shoot automatically after lock, but will shoot once multiball start. Not sure should ballsave launch automatically??
'031 iaakki - ship added
'032 iaakki - gun draft added
'033 iaakki - gun bracket added and minor tweaks
'034 iaakki - tracy head samples added
'035 iaakki - make tracy flip
'036 iaakki - pf draft
'038 iaakki - pf updates, horizontal ramp reworked, POV readjusted and some issue fixed
'039 gtxjoe - add debug outlane/drain blocker (press 2) and shot tester.  Press and hold keys to capture and test shots (W E R Y U I P A S)
'040 dapheni - rework to the playfield, left VUK using Hole_1. EnableTeleports trigger is disabled for testing, just enable it if you want the lower PF
'043 iaakki - removed another tracy, set up some materials
'044 iaakki - left ramp heights adjusted, making more room for upper flipper, added flupper flip logos
'045 apophis - added Fleep sounds and set up nFozzy posts and sleeves on table. Double checked physics material assignments. Moved cor.update to gametimer (10ms)
'047 apophis - lengthened upper flipper and tweaked the sky orbit
'048 iaakki - made left outlane suck more. Added kicker on left outlane
'049 apophis - reformatted script with auto-indent routine
'050 iaakki - Multiball Wizard mb tied to left magna<
'054 iaakki - merged in some old stuff, cleaned up pf, repositioned inserts
'055 iaakki - textured the hatch and made a wall that would drop when it opens. Hatch mechanism is now tied to magna buttons
'056 iaakki - pf tuned, pf mesh added, lower playfield tuned, right scoop fixed
'057 oqqsan - Started on skillshot bonus and 4 flipperlanelights , not done but "working"
'058 apophis - Added Flupper bumpers
'059 iaakki - reworked ramps higher, plastic samples contoured
'060 oqqsan - flippers now move lights on top and bottom. added 2 smal lightsequencers for bonus lights and flipperlanes
'061 apophis - merged 60 and 59_lazer. Added lazer and spinner sound effects. Made most plastics walls collidable.
'062c oqqsan - KB should be reset and working, gate2 twoway disabled on skillshot, leftflipper before plunger is pressed opens it and start skillshot 2
'063 oqqsan - removed ultra dmd started on FlexDMD only, scores work and thats it :P
'064 oqqsan - new dmd layout, added some light blinking to everything on topline on default scoreboard, a few new images for the dmd
'065 iaakki - reworked lazer and gun code. lazertimer set from 5 -> 10
'066 oqqsan - added some flupperdomes 2.2 ( no bloomflashers ) made 3 kickers set off flashers before KickOut
'066b oqq - deleted the onscreenbackglass pupplayers .. mby not a00% clean yet
'066c oqq - added mystery ( some text missleading, only gives points atm )  DT on left reset It
'     - broken the next ball thing
'068 iaakki - shipkicker added, lazer tuned, inserts moved around
'071 apophis - Based on version 68. Fixed next ball thing. Fixed end of ball. Updated Mima target logic. Lots of script clean up. Saved file in build 305 of 10.7
'072 oqqsan - new topramp prim, added Ramp control to lower and rise it up . stresstest is enabled on flipper L+R when game is On
'073 apophis - Renamed lights, targets, and switches. Updated script accordingly. Organized the game code a bit. Update bumper material properties. Programmed Corey MB lock & start logic. Jackpots not programmed in yet.
'074b oqqsan  - mima should work, added some random to corey lights. fixed skillshot
'075 iaakki - reworked the mima ramp area a lot. Reworked lots of metals. Right outlane has a thicker wood wall added.
'076 oqqsan - fixed some more on mima and corey mBalls
'077 apophis - Mima targets and ramp reset properly. Changed name "Ramp017" to "RampMima". Added Tracy MB locks and logic. Added Mima MB Jackpots.
'078 iaakki - Mima target redone, various plastics reworked and some material tune
'079 apophis - All main MBs are complete now, with all JPs and Super JPs working. Multiball Wizard MB still not made. Load Lazer insert working.
'080 oqqsan - save/swap lights on multiplayer/newball.
'081 apophis - Made Mima locks lite after only 3 target hits instead of 4. And Mimacount starts at 2 for each player. Made sky loop a little easier (added friction to platform to slow down ball). Added ball shot delay and 3 sec ballsaver on lazer.
'082 oqqsan - Mission select .. probably good to go :::.. we will find out when 1st mission is done  . saving credits to memory fixed, no longer resets
'083 apophis - Made MBs a little easier to achieve. Added Mission 1 (Bladerunner). Mission shot lights are conflicting with MB JP shot lights. Need to find a solution to that.
'084 oqqsan - fixed some joint forces on the JP/mission lights, credits fix again
'085 apophis - created SetLightShot function and GetLightShot subroutine
'086 oqqsan - mysteryEB up and running ! trying to start mission 2, cleaned out some unnused timers
'087 apophis - added Mission 3. The mission lights are not resetting properly when mission stops... needs fixing.
'088 iaakki - Mission lights moved to apron, multipliers under octagon, mission names splashed to dmd when started.
'089 iaakki - new laneguides in
'090 iaakki - backpanel flashers reworked and some triggers tied for now. Some prim positions adjusted. Lower GI started.
'091 oqqsan - added introvideo fast and furious.. changed dmd folder name !! so update this
'091 oqqsan - added introvideo fast and furious.. changed dmd folder name !! so update this
'092 apophis - added Mission 4. Also, now mission lights resetting properly when mission stops.
'093 iaakki - GI continued; some wall edits missing in top area. Not sure why right most laneguide bulb looks dimmer. Probably some depth bias thing.
'094 apophis - added Missions 5 and 6. Made all MBs stackable. Rubberizer was only on right flipper, now on both. Adjusted subway eject
'095 iaakki - tracy fury testing...
'096 apophis - tied Tracy Fury to mission 7. Made Tracy head speed up during mission. Adjusted Tracy head movement equations to make it more stable.
'097 apophis - added ballsaver at plunge and KB. Moved swOpenGunTrigger to secret lane. Moved multiplier inserts under apron for now. Deleted wizard insert. Deleted Trigger2 (not used). Made ballsave inserts align with eyes. Tweaked Tracy Fury parameters
'098 apophis - Added Multiball Wizard MB. For testing: hold down left magna for 3 seconds to start Multiball Wizard. Made Lazer autofire after roughly 10 seconds. Now Tracey head only moves during Tracy Fury mission.
'099 oqqsan - update some DMD stuff
'100 apophis - added Mission Wizard MB. For testing: hold down right magna for 3 seconds to start Mission Wizard. Fixed some lazer autofire logic. Changed "Heart Attack" to "Heart of Steel".
'101 oqqsan - added extra splash line on sb..... fixed lvl 3+4 ( apophis found it first )
'102 iaakki - under pf cover and under pf mesh done. underpf Image swap done for mission 5 start. ball shadows fixed
'103 oqqsan - imported some insert prims and changed up the scoreboard alittle took some digging to get it right this time, wrong order is a biatch sometimes
'104 apophis - fixed DTs not resetting on new game. Slowed down hatch opening animation a bit and made cover fade in and out. Made current mission light blink rapidly when mission active. Fixed mission lights to reset properly between multi-players. Fixed mission 5 to work with MBs (mostly)...still has some issues.
'104bc oqq  - new intro / and made some with the gif lumi made :)
'105 apophis - added JP shots to Multiball Wizard MB. Made SS2 use LightSeq for blinking. Made super jets active in 50 bumpers instead of 25. Added EB light lit after 3 and 6 completed missions. Max EBs per game set to 3. Made Mission 5 stop if MB starts... it just doesnt work with MBs.
'106 iaakki - added music with PlaySong and ReturnDefaultSong subs. Place BloodMachinesOST folder into \Visual Pinball\Music\
'106a apophis - testing out using windows media player activex control for playing songs. Allows for starting a song at specific time in seconds
'107 iaakki - refining inserts positions, types and sizes to match real life insert parts
'108 oqqsan - ss1 ss2 randomrewards almost done probably need to fix something.
'109 iaakki - DT's updated. Added some fancy to Cemetery mission.
'110 oqqsan - added DMD displaying extraball +mission failed/completed, tried to fix mission 5 alittle, was not resetting or closing on balldrain
'110a oqqsan- moved the kickback lights ( might wanna resotre those or ?  showing tracy bumpers left on DMD
'111 apophis - fixed DTs not resetting on new ball. Fixed Mission6 to stack with other MBs. Fixed SS1 and SS2 Start Mission award to skip to next award if mission cannot start. Made autoplunger activate after Mima Lock. Fixed DMD to show MISSION COMPLETE or FAILED correctly. Fortified back wall a bit to prevent stuck ballz
'112 DJ - Added plastic Flupper Ramp
'113 iaakki - under apron GI added, flip length tuned, some depth bias fixes
'114 iaakki - moved things around. Added gioff_underpf_boost sub. Shake_tracy_shake added to bumper3
'115 apophis - added Lampz and control lights. Set up a few 3D inserts, but they are not tuned.
'116 apophis - added more 3D inserts
'117 iaakki - ramp heights reworked, inserts moved around, the whole upper ramp redone *one more time*, now it should drop slow balls.
'118 iaakki - removed ship ramp and tested having plain wall in there.. Now it works, but doesn't look too good. Fixed ramps behind back panel a bit. Moved insers under ramp.
'119 apophis - added PlaySongLoop sub. Added AstroNaty's updated PF art and insert overlays. Fixed a wall preventing top loop shot. Fixed swSkyRamp position. Now only 1 Corey light is awarded per sky ramp. Fixed underPFshot control lights.
'120 iaakki - updated env image, messed up with lights alot. One can use gioff and gion commands in editor to see how it lights up. DN slider at zero just to make it really dark when GI is off
'121 apophis - Update insert colors. Added Astro's latest PF. Removed insert overlay image. Added attract mode light sequences. Updated Flupper flashers to work with Lampz. Made Flash5 blink using light seq on Mission 1. Updated Mission Wizard mode.
'122 apophis - added Astro's lastest PF art. Fixed PlaySongLoop to have no delay between loops.
'123 apophis - Updated Mission 7 to Betrayal Mission. Changed the way mystery scoop works.. it is now the scavenge scoop. Added large insert near bumpers. Adjusted upper wire rail near flipper so ball not catchable. Now song pauses when table is paused.
'124 apophis - Added dynamic ball shadows. Added TargetBouncer. Added ramp rolling sounds. Possibly fixed ball release bug when mission ends in game over.
'125 apophis - Added flasher bloom to Lazer blast. Made pentagon fade in/out 10x slower. Updated PF image. Added baldgeek's logging functionality. Fixed wireloop ramp rolling issue. Fix ball drop sounds to work on under playfield. Added several env image HDRs to pick from.. need to decide and tune env lighting.
'126 Wylte - No more goddamn "bloop" on inlanes
'127 apophis - Made Dust fade in and out. Made ship flash underneath with ball loops. Fixed Mission 2 WriteToLog error. Fixed Mission Wizard error. Fixed Mima MB releasing too many balls.
'129 iaakki - redid all my 128 changes as file size got bloated by 20mb for some reason
'130 oqqsan - changing Gi colors on missions  GI_Color "green" ( or the other colors  red,orange,blue,purple  all tem colors )
'131 iaakki - adjusting GI and adding first prim plastic
'132 oqqsan - flexdmd2 on apron test1
'133 iaakki - new pf mesh. Vortex added under ship. Missions and modes broken, as logic needs some fixing.
'134 Wylte - Tracy Attention test: currently follows ball for 0.5s on bottom lanes and gun trigger, but spins 360 to reset
'135 iaakki - New ship model from Tomate, GI fixed and finalized, changed default gi color to be more white, reworked the dropping wall in the vortex. This needs ship elevator animation.
'136 Wylte - Tracy Attention fixed & added to orbits, head shake added on bash toy hit, loop wall set to only drop with flipper
'137 iaakki - ship made smaller and raised a bit, pf mesh vortex made deeper, bash toy resized and some audio added, vortex bottom and tube created, vortex gi light added, ship_shake on bash added, wVortexBlocker added to prevent ball stuck just in front of the bash wall, loop speedup multiplier reduced from 2 to 1.1.
'138 apophis - changed Mission 1 to be quick MB, and Mission 6 is now a timed mission. Also made ShipCannonKicker (the EB Kicker) an add-a-ball feature. Added some light sequences for Jackpots. Deleted some unused env images
'139 Lumigado - added AudioCallout, StopAllCallouts subroutines; began implementing movie quotes into the sound
'140 apophis - Added MusicVol script option. Set light height to 3000. Updated Astro's newest PF image. Made Double Scoring insert blink fast when that mode is active. Added Super Spinner mode. KeepLogs True by default
'141 iaakki - pre-skyloop speedlimitter tuned, vortex redone *one more time*, There is a helper magnet called "VorMag". It's strenght is 4 now. Test it out is it too easy now.
'142 apophis - fixed light super JP seq bug on Corey MB. Fixed EndMission2 bug. Enabled loopramp timer when vormag is on. Made all mission stop early if player meets some max success criteria. Converted kickback to a cvpmImpulseP object for better reliability.
'143 Wylte - Fixed Tracy, added 2 more sources for dynamic shadows, adjusted shadows +0.1 z to prevent clipping with pentagon, added Poltergeist spinning head effect (somewhere...), added Apophis' flipper fix, lowered LiveCatch to 16
'144 iaakki - MimaMagnet added with release mechanism. Is is now tied to Cemetery mission. Strength is now at 10, tested on 15 and it was a bit too much. Dust materials split to 3 materials to give some more depth. Some old code cleanup from teleports.
'144 apophis - Changed env image back to normal one, changed light height to 700. Fixed Mima lock bug (wasnt allowing lock sometimes). Fixed scavage light bug (hopefully). Added a lot more debug logging outputs. Made Poltergeist Tracy effect only occur during Super Spinners.
'145 apophis - merged two previous 144 versions. Added wire guide behind upper flipper to prevent ball falling off ramp.
'145a apophis - adding extra logging outputs
'146 oqqsan - added 5 gifs on mission1 the easy way... a little preview/test
'147 Wylte - Merged 145a+146, made portal room work with multiballs, Tracy Subs all have compatibility/priorities, Tracy shake at gaze angle, end poltergeist early when spinner multiplier lost, set Tracy eyes to reset on ball lost
'148 oqqsan - did some more on mission1 DMD( will get new fonts ), added the new intro.gif
'149 apophis - Updated PF image to Astro's final version. Removed "start mission" insert below ship. Made load cannon insert brighter. Added light for Turbo Insert. Updated Mima Magnet subs for handling Turbo insert blinking. Added MimaMagnetFlicker switch. Tweaked helper wall on sky loop to make it easier. Made Tracy lights lit after 10 bumpers instead of 15. Made most missions a little easier to complete.
'150 iaakki - back panel holes cut and pincab_sides image updated. That can be used as a template. Flashers needs some repositioning.
'151 oqqsan - scoop fixed on, underpf multiball fix ( subway will probably be alittle better ) lets see
'152 Dapheni - The magnet is now reset at the start of a new ball if the mission was abruptly stopped. (just added MimaMagnetOn False in ResetForNewPlayerBall)
'153 iaakki - changed underdropped from drop to non-collidable, fixed underpf image swap for spacewoman start. Reduced vortex magnet
'154 iaakki - all plastics redone and uv mapped, plastics texture file added, gi redone to use lampz fading. Some collection tied to GI events to fade along. Wall001 at least is somehow broken, as DL is not affecting it.
'155 apophis - Fixed issue with Ending Multiball Wizard MB prematurely. Fixed collections indexing error in AwardSkillshot2. Fixed issue with Tracy head spinning forever if ball drains during a spin. Fixed RotateScavangeLights. Added more glowiness to Turbo insert. Added some delay before VUK1 kicks prior to mission start (allowing time for intro animation). Added variables to save mission points per player (This info will be used during the Mission Wizard). Made AddBonus sub and sprinkled bonus points around the game (point schema needs review). Bonus points * Bonus multiplier now added to score at end of ball.
'156 oqqsan - Added videos on mission2  ( counter goes to apron =? so not on dmd )
'157 apophis - Added newest PF image. Removed unused Time Square hdr image. Fixed Mission 4 canceling on sky orbit hit. Fixed KB from using both KB lights on a single kick. Removed ballsaver after KB kick.
'158 apophis - Updated to latest PF image. Added Mima Target insert. Added portal glow prim and logic to ramp switches. Added insert prims for the lightning bolt inserts. Changed the blink behavior of the Turbo insert.
'159 oqqsan - mission 3 gifs added, might need to change some of them and set proper timers same as mission 2 : New images plastic pf backwall. magnetscriptupdate
'160 iaakki - Adjusted laneguide elasticity and friction, shooterlane ramp, flasher parameters, dampener collections, flippernudge values
'161 apophis - Added some basic DOF scripting (more needed). Switched TargetBouncer back to original code. Added target bounce to cockpit target. Tweaked drop loop to avoid insta-drains. Tweaked portal material color.
'162 oqqsan - just a start on mission 4567  + a few smal fixes
'163 oqqsan - mission 1-7 all with gifs playing on dmd . could use background images for all missions ( only 1 so far is used on all of them ) added a invisible wall at bumpers to stop ball get stuck
'164d oqqsan - fixed 3 portal lights on backwall. cleaned up the dmd for obsolete code : NEW skillshots, added ss3 at scavenge kicker ( lets see if its possible to get there in time ) : removed flexpath, lumi had a better solution
'165 oqqsan - experiment  DMD FontScoreActive
'166 iaakki - VUK2 and its ramp redone
'167 oqqsan - more fonts, adjusted alittle and added a custom 5by7 font flikkering on mission 1 ( only one added on so far )
'168 tomate - new apron prims added, temporary apron textures added
'169 iaakki - Cannon lane ramp redone, laneguides to red plastic coat material, wire ramp inlane ends done.
'170 apophis - Added a bunch of audio callouts.
'171 iaakki - small angle tune to ScavengeKicker, so it lands on flip correctly, Mission lights moved. Fine tunings here and there
'172 tomate - apron objects separated as independent prims
'173b oqqsan - added new galdor sphere images.  added more background images for missions : ..  @embee sure makes alot for the dmd :)  hooked up the missionlight primitives
'174 apophis - updated plastics image. Added callout volume script option. Made vortex open during mission wizard mode. Turned off shadows when ball is in lower PF. Reworked portal light logic. Removed sky ramp insert and replaced with portal light blink. Improved backwall flasher feng shui. Adjusted flasher parameters. Made top left ramp a 2-wire ramp.
'175 oqqsan - all callouts added vpmtimer 300 ... fixed on missionlights added timer on apron ( only showing in mission)
'176 oqqsan - Mission buttons new color and 2 letter mission text .. mission 1 redone to timer 100s and no failed on first ball drained ( mby too much ?
'177b oqqsan - added some prims for 2-10xbonus, fixed on mission 1 again and again
'178 apophis - made MBs only start from a VUK2 hit. Delay before VUK kick is adjustable.
'179 iaakki - primitive walls done and added to GI events. Also some other prims added too. Plunger light panel added, but cannot read plunger position with a key only.
'180 iaakki - prim walls redone one more time and UV mapped briefly. Plunger Light improved
'181 apophis - added a bunch of DOF commands. Updated plastics image to Astro's final. Made Tracy lights brighter. Fixed bug where bMultiballReady never set back to false after MB.
'181a apophis - deletes old debug log file at table1_init
'182 oqqsan - Galdor      command:  GaldorQ1="1234567" and or "ABCDEFG" for double speed ... string can be as long you want, open close on auto
'183 apophis - Fixed some gameplay bugs around MB starts and stops. Added DOF E152 command. Disabled reflections for PinCab_Blades. Added KeepDOFLogs flag.
'184 iaakki - insert colors changed
'185 iaakki - flips redone and repositioned, both slings and laneguides adjusted, laser more narrow and bright, few bugs fixed
'186 apophis - Moved all the score value definitions to top of the script. Renamed Mystery targets to Scavenge targets. Renamed Space Woman to MB Wizard
'187 Lumigado - First pass at scoring values, split inlane and outlane constants
'188 Lumigado - Finished initial scoring values, added lots of score constants
'189 iaakki - AO shadow layer done, PF updated
'191 iaakki - plastics and bevels redone
'192 iaakki - Ship bash target captive ball done, triggers may need tuning still.
'193 iaakki - Plywood Holes, nuts, bolts, tuning sling primitives, materials for posts...
'194 iaakki - VortexMagnet 3 -> 6, ShipCannonKicker new values, RubberPost003 moved left a bit, Plunger Scatter Velocity 0 -> 5, testing darker PF (thanks skitso for the tip),
'             captive ball trigger now used for bash events, vpmtimer "EndOfBall '" 500 -> 2500, hidden gun entrance wont start missions anymore
'195 iaakki - cross by the launcher pulse added, added consts to adjust GI brightness and transmit, fixed some hatch related shadow issues. Not sure how to make VortexHoleEdge visible. It shows fine in F6 mode.
'196 iaakki - added prim wireramps
'197 oqqsan/embee - Added the rest of the Galdor animations...  GaldorQ1="abcdefghijklmABCDEFGHIJKLM" : any order/length, open/closes on auto can reset this anytime by giving new order
'198 iaakki - added new bumber colors and materials that are tied to GI color. Made some MimaLoop Wall improvements.
'199 oqqsan - made a simple gauge with spinning wheel..1st attempt :P
'200 iaakki - Gauge textured and tied to multipliers
'201 Wylte  - Tracy movement smoothed and unified across subs, flashers for ramps added, Tracy and shadow code updated to account for locked ball, shadow code updated/optimized, latest comments/instructions added
'202 iaakki - Proper GaldorShade created (GaldorShade true/false); now we can remove all the galdor animations related to it closing. Also improved performance for EOB gauge.
'203 iaakki - UnderPF rework, Galdor animations added to mission5.
'204 iaakki - Music path detection improvements. Test table for song loops.
'205 Dapheni - Small rework of angles in the portals on the lower pf
'206 iaakki - plywood hole prim, left platform prim, metalwalls prim modified, tonnes of scrip changes, new underpf images, GI adjustments, plastics made darker
'207 iaakki - MB end now returns default songs, new plastic ramp prim, adjusted materials, prims tied to GI events.
'208 Wylte  - Ball following reduced to gun and outlanes when not in double scoring mode, Wall011 added to prevent stuck ball
'209 Oqqsan - Started on "new" dmd, organized the dmd folder abit too, so this need be UpdateDebugBox
'211 iaakki - 5 pf level flashers added and tied to other domes, DT detection has own update in frametimer, but that needs to be reworked out later, gauge glass added
'212 oqqsan - more added to new dmd setup :)  normal skillshot ( mby the others work fine too )   bumper hits with video ( not completed ending =)
'213a oqqsan - Mission 1 added along with a few fixes.... insertcoins added
'214 iaakki - fixed collision sounds and reworked the loop, fixed some walls
'215c oqqsan - dmd update: mission2 and 3 + lostball/ballsave ( ballsave outside missions for now ) some smal fixes
'216 iaakki - baked sideblades and tied to colored gi. meaning that flashed color is changed on fly to match GI color
'217a oqqsan - mission 1-7 added to new dmd ... one playthrough almost both wizards :P
'218a iaakki - LoopWall lights and sounds reworked
'219 oqqsan - new plastic image from astro ! Added Attractmode + enterhighscorename mode , fixed gameover alittle
'220 oqqsan - bottom right display upgrade ' attract and highscore mode fixed on, added some for jackpot, hooked up apron stuff at gion/off
'221 oqqsan - EOB bonus !!!  mimaloops, ramps, missions... const added in the list for those  25k 100k 1M
'222 iaakki - Added sounds here and there. It is starting to breath.
'223 iaakki - wireramp sound fixes, wireramp bump sounds, laser optimizations, gun rotation optimizations
'224 apophis - Removed ability to make progress toward MB during a MB (fixed some MB logic issues). Fixed crash when Mission1 ends with no balls on PF. Made both top gates collidable on Mission 4. Removed AdultCallouts. Added "ball added" callout.
'224a apophis - Fixed CheckMBReady issue inside PlayerLights_NewPlayer. Made it so now not even Mima, Corey, or Tracy lights are lit (blinking) during MB.
'225 iaakki - Lago in pain sounds added to Heat of Steel mission
'226 apophis - Inserts behave differently during MB (they now use light sequencers). Created SwitchRecentlyHit and SwitchWasHit to improve the switch checking functionality during multiballs.
'227b iaakki - New EPIC ship primites from Flupper. Using 3 primitives to achieve all the lighting steps using fading and runtime material updates, so ship reflection matches GI color.
'228 iaakki - sideblade image was wrong and now there is some issue that they are not fully visible. Sideblade and backpanel flashers redone, red dome flasher images baked and tied to flupper dome code
'229a iaakki - GI replaced with RGB flasher. GI ball reflections reworked, hatch z fighting solved, Captive ball cannot escape anymore
'230 iaakki - GI, Red And green dome reflections rebaked and tied to code. One can test them using magna buttons for now. Purple dome reflections are still missing
'231 iaakki - Purple dome reflections tied in.
'232 iaakki - reorganized and combined huge amount of timers, probably broke something too. Set tens of prim non-collidable and disabled reflections.
'233 iaakki - some more perf improvements in. Ships 1 and 2 are now set to render backfacing transparent.
'234 iaakki - white done reflections added, jolt sound removed from lanehguides. Only effect when kickback gets enabled
'235 iaakki - underpf improved, flashers tuned, under ship flasher adjusted etc
'236 iaakki - tracy shakes reduced to bumper2 only, EB max 3->2. Tracy figure pf reflections disabled, Started adding insert blooms. They are at 0 when GI is on. So only visible once gi goes off.
'237 fluffhead35 - Updated ramproll sound logic, added ramproll and ballroll amplification sounds.
'238 iaakki - Ballroll, ramproll and shadows limited to max 6 balls. Maybe works.. Inserts fiddled
'239 fluffhead35 - Adjusted RampRoll logic for 6 max balls as well as removed RampRoll Sounds that would not be used now.
'240 apophis - Divided all scores by factor of 10. Limited EBs to two per player per game. Reduced SSR to 0.1. Fixed Turbo insert.
'241 apophis - Added some B2S commands. Allow for commas in the high score.
'242 iaakki - Spinner EFX added, ship GI levels tuned
'243 iaakki - MimaBulbs on plastic redone, sounds for skyloop, drain, sinkhole, gun loading, lane switches, ship bash toy, plastic image fixed
'244 iaakki - plungerlane shoot sound, heart beat sounds for mission 6, some missing fleep sounds added, tracy&corey light add and collect sounds
'245 iaakki - plungerlane metal redone
'247 iaakki - remerge, New darkened PF, all general lighting settings readjusted. Tracy_shake disabled from bumpers totally. Top lane EFX sounds removed
'248 iaakki - some cleanup, flasher light intensity tweaks and improving the flasher effects
'249 Sixtoe - VR cab and minimal room, new playfield
'251 iaakki - 360 vrroom, dust reworked. (Cabinet art missing, VR dmd not working, Flasher audios buggy still)
'252 oqqsan - fixed minimalvr to show dmd, redone apron dmd, vr room didnt want to share that flex feature :)
'253 iaakki - Adjusted DL for apronnumbers, modified ramp corner rubbers, modified ship loop, added diverter sounds
'254 iaakki - laser updated. Still needs new textures. nice tool https://ezgif.com/png-to-webp
'255 iaakki - laser updated more, LoopStandUp style changed
'256 iaakki - HOS mission fixed and font size change, Ship lights tuned, VUK2 sound added when mb is about to start
'258 iaakki - back portal lights VR issue fixed, underpf vr issue fixed, VUK1 style changed
'260 iaakki - some fonts fixed from missions. Won't return to default song, guntarget shot started
'261 iaakki - making sure Sixtoe will puke his spacesuit
'262 iaakki - increased vuk2 delay to 4s to fit EFX better, Tuned VRRoom, learning how to add dmd videos for MB's
'263 iaakki - Tracy and Corey MB DMD videos cutted and tied to MB modes.
'264 iaakki - Gun target shot finished, captive ball hits count as opening mimalock, Large Target hits enable PF multiplier for 30 seconds, Bulbs still on the plastic, but those can be removed.
'265 iaakki - PF image updated and Insert PF overlay ramp added and adjusted. Do not touch them anymore pls.
'266 oqqsan - added some flipper info to show sooring for all players when holding flipper 3 seconds
'267 iaakki - corner scores styling changed
'268 iaakki - dmd reworked and some callouts added
'269 iaakki - borderblink default mode changed, score display logic rework
'270 iaakki - Fonts reworked, DMD code made even more complex
'271 iaakki - Insert blinking speeds updated; easier to distinct JP shots from normal shot lights, HeartofSteel randomized better, Sling sound effects, GI sparks when going off
'272 iaakki - teeny font modified to be more compact, DMD font and position fiddle, added few AudioCallouts
'273 iaakki - flip physics redone, flips repositioned and angles adjusted, post pass adjusted a bit, ball saver times reduces, add-a-ball feature limited, main ramp return tuned, scavenge scoop scatter added, ramp had wrong physics material
'274 iaakki - Flip shapes and physics tuned, Rubberizer 1 now default, flip logos borrowed from Bord's Genie, bumpers and some prims tied to gi fading, skyloop shot made harder by reducing speedlimitter effect, AllFlashers sequencer added, mission change after mission end
'275 iaakki - MB Wizard mode debugging, new grabber magnet added
'276 iaakki - another plungerlane from Tomate, underpf flip logos updated, grab magnet fixes. Now leaves one ball in play while holding others. VUK2 kicking might fail occasionally. Not sure why..
'277 iaakki - side reflections flashers added with hide prim. Backpanel flashers not yet tied to RGB gi.
'278 iaakki - reworking the wizard. Most of the phases added, but needs a lot of work still
'279 iaakki - VUK1&2 reworked not to use vpmtimers, grapmag release improved not to spam vpmtimer, Loop speed up limited for wizard phase 1,
'281 iaakki - apron rework with Astros edits
'282 iaakki - tracy perf fix and motor sounds
'283 iaakki - wizard ending logic quite done, trace model optimized, tracy motor filtering adjusted, wireramp reflections done
'284 iaakki - Hides tuned to work in cabinet mode, plastics and plungerlane prim reflects added, PF side reflection added, wizards ending improved, added some GI to wizards(color needs variation), UnderPF blocker added on top of flips, one underpf ball stuck fixed
'285 iaakki - New VPW Flip bats by Tomate. Wizard phase1 DMD and sound work started, various fixes to light logic
'286 iaakki - sideblade ball mirroring implemented. Only one is reflected per side.
'287 iaakki - mirroring bug fixed, RubberPost005 moved, top flip bat swapped, set default visual values, fiddled underpf & DMD a bit
'288 tomate - new wireRamps and plasticRamps textures added
'289 iaakki - plastics tied to RGB GI
'290 iaakki - Wizard musics done and some callouts added and timed. Wizard ending teardown drafted. PlaySongPart method implemented
'291 iaakki - tracy fiddle for wizard, mission start needs to be enabled with an orb shot, cannon-kicker now gives ballsaver and only adds-a-ball in MB mode, Scavenge scoop AllInserts effect added; maybe need similar for every VUK.
'292 iaakki - Mission start now available straight for a new ball; unlock only after playing a mission, fixed orbUnlock mechanism with guntarget, fixed a bug with mission lights, pMimaTarget blinking tied to double scoring
'293 iaakki - Betrayal song change, rubberizer deleted, tilt and warning DMD infos added, Mystery DMD fixed, mission maxtimes checked, Mission end infos added, adding mission time for loop shots. unless in mission 2 resurrection, Chase, Betrayal and Cemetery need all the shots now to complete.
'294 iaakki - fixes to eliminate OrbUnlock to happen while in mission and endless ball saver for wizard
'295 iaakki - orbunlock one_more_time, laserblast sound tied to ssf and ball trajectory, some other SSF fixes, 2s grace added to multiballs, AI ship DMD counts added,
'296 iaakki - Betrayal 1 combo to finish, ballsave fix, wizard logic fixes, Wizard phase 2 and 3 DMD work started, phase3 takes now 3 shots to both targets, KB status retained in wizard
'297 iaakki - Wizard phase 4 dmd draft in. Still a lot to do.
'298 iaakki - bumpertop collidable walls, mimag disorient taken out from wizard and non vrroom, wizardphase3 redone, portal room DT's tuned, underpf image swapped, ballblast sound redone totally with new sound, wizardphase 3 fails the wizard on 3rd underpf drain.
'299 iaakki - Wizard ending with final show added
'300 iaakki - Flex DMD font edits, audio callouts added, optimizing table 60megs, apron wire color change, AO shadow combined to PF image (affects perf)
'301 iaakki - Spinner rework
'302 apophis - Updated scavenger drop targets to Roth style. Added Fleep DT sound effects. Added teleport sound effect (EFX_teleport). Fixed logo backglass call.
'303 Wylte  - Updated shadow code, inrect/mode underpf check removed, faster Tracy head reset
'304 iaakki - merged some fixes, fixes tearing down wizard properly, autoplunger DOF and delay fiddled, orbunlock reworked partially, mission start light state retained over MB (really not too common to happen)
'305 iaakki - flipper info removed and skillshots tested, KB maxed callout and DMD info added, Wizardmode MP scoring labels fixed, galdor globe impressions added to wizard, typo fix
'306 iaakki - SuperSpinner timer 20s, disable scavenging while in MB modes, The Hunt mission DMD altered, slamtilt with start+left_flip (dmd breaks), galdor shade sounds, underpf teleport delayed 800ms
'307 iaakki - DMD fixes, various mission logic fixes, galdorshade illegal call fixed, shade timed better, GetMusicFolder code from NailBuster, some options cleared, back panel portal reflections, right outlane tuned
'308 leojreimroc - VR Backglass added.  Flashers and GI added to backglass and speaker panel.  Animated VR Plunger and flipper buttons.
'309 iaakki - various VR playfield and lighting fixes, cabinet hinges probably needs to look different as they block reflections
'310 iaakki - ball min/max brightness options added, pf flashers dimmed, orbunlock tweaks one more time, wizard scores added
'311 iaakki - CabinetMode implemented
'312 iaakki - added more orbunlock checks, gun load insert PF cut and insert layer redo, captive ball trigger adjusted, high score submit layout tweak, made it possible to have several MB's ready
'313 iaakki - turbo magned hold fail-safe increased, mb restart check after another mb, captive ball trigger filtering improved, orbunlock after failed mission -> MB case
'314 iaakki - flasher blooms with script option, bMultiballReady flag status checks,
'315 iaakki - skillshots fixed, reflections fixed for some inclinations, POV import lighting break workaround, ball search (keep ball gradled for 10s and it will start)
'             Removed underDustLayer and some PF mayonaise, brought back the under pf boost, disabled inlane insert rotation for mission 5 and wizard phase 3
'316 iaakki - Corey SuperJP needs to be lit by getting a jackpot first, audiocallout hack, PinCabRailsColor option added, railblocker prims added, subs to maintain lane insert statuses for missions with hatch,
'     - fixed dmd logic after tracyMB, Extraball and Jolt DMD modified, default POVs adjusted, SuperJP DMD's added, cleanups
'317 Sixtoe - Fixed under_pf drop targets, added new vr360 portall floor, hooked it up to portal mission at the moment.
'318 iaakki - GunSight added; another globe that Rik can eat.
'319 iaakki - GunMotor sounds and motor shaft added. GunSight color changes when autofire warning start
'320 iaakki - Scavenge Light Rotation and sync fixed
'321 iaakki - GaldorShade material updated, extra ball start call out added.
'322 iaakki - Mission Select disabled, Left Magna hold enables wizard for now, improvements to ballsearch, 2nd ball removed from The Hunt Mission, MimaMB Super Jackpot scoring fixed, disabled advancing MimaMB when in mission 1,
'       Disable bottom lane rollovers for missions with Hatch, annoying tracy alert sound added to double scoring and double scoring made to end faster, placeholders done to start improving attract mode.
'323 apophis - Made cabinet mode automatic. Updated ball and ramp rolling sounds and scripts. Deleted a bunch of unused sounds.
'324 iaakki - Added more switches to SwitchWasHit, scavengescoop not lit while in MB, red dome color adjust, SuperSpinner time increased to 30s, Attract DMD rework, songplayer tweak, other DMD improvements, ball search tweaks,
'325 iaakki - Mechanical Tilt binding and Rules Card added for Thalamus. Letters "R" and default is "T".
'326 iaakki - one more scavengescoop tweak, slingshot force reduced, DT elasticity reduced, some variable resets,
'327 iaakki&oqqsan - DMD fixes, ball search changes, disable orbunlocks when MB starts, StopVideosPlaying added to mission_begin; just to eliminate bumper video continuing
'328 iaakki - ball image swapped, gun targeting shot score increment added per player, tracy gaze adjusted, blast button added, limited plate added,
'     ambient sounds added for players who have not obtained the soundtrack, Slingshots and right outlane adjusted
'329 iaakki - blast plate adjusted, savehs fix, coreylight issue with skillshot fixed, tilt disabled when no balls on table, tilt warning dmd fix,
'330 iaakki - Hide z issue in VR, DMD sunken, Flashers balancing with new parameters, own LS's for wizard orb shots, some sequences added to mission complete,
'RC1.1 iaakki - DMD to use coloredPixels, TILT handling improved, flasher solenoid sounds fixed, shottester removed, cleanups, gioff_underboost redone, plywood hole fading with GI, some insert sideblade reflections, vuk1 hidden
'RC1.2 iaakki - Ballsaver grace increased from 2 to 3 seconds, ball visuals adjusted, minor apron collidable change, lightflash5b adjust, LS_flash5 after draining, slight randomness added to mimaloop timers, debugs and logging disabled, OptionWarmGi
'RC1.3 iaakki - Side hide and reflection flashers tuned, VUK1 set invisible again, GI flood light shapes fixed near the cab edge,
'RC1.4 iaakki - vuk's and diverter redone, default music volume check, increased audiocallout vol, vrdmd and vrapron blockers from Rawd added
'RC1.5 Sixtoe - visible rubbers tweaked, scoreboard initials fix (thanks cupii), b2s hack for mima (B2SSetData 4) , script tidied.
'v1.0 Release
'1.01 apophis - Added DMDColor option. Updated Table1_Exit to stop controller if B2SOn. Updated the hack to make Mima light on backglass at startup. Fixed Skillshot for analog plunger.
'1.02 apophis - Fixed flipper stuck issue when game ends. Added fix to swLoopUpTrigger_Hit to prevent ball catapult. Tied flasher dome off DL to GI level. Added cannon fire button (LockBarKey) and DOF functionality (E153).
'1.03 iaakki - Fixed skillshot and mystery inconsistency, added startup music, orbunlock after wizard, wizard dmd fixes, testing EOB bonuses, higher non-visible wall added over apron, fixed a rare case when attract stays on when game started in certain point on attract, minor tweak to galdor shade closing, changed the mission EOB bonus scheme
'1.04 iaakki - Staged flips support, typo fixes, instructions update, Normalized audiocallouts added (thanks idigstuff), scoring change to slings, top flipper hint animation, slight change to default cabinet POV
'1.05 iaakki - New VR legs and fittings, plunger animation fix
'1.1 iaakki - Pincab_Shooter anim value change, grabber releases balls after 6 seconds of hold, leftoutlane switch fix after wizard, mima lock callout bugfix
'1.11 iaakki - Disable skillshot while MB is ongoing, improve the state when all missions are played, but mb's available, start wizard phase improved, some additional resets to mission ending, some vpmtimers removed, tracylight fix for wizard
'1.12 oqqsan - Ball reflections on hatch .... missing touchup on the image "playfield" ( can cut all but the hatch area away to save MB's
'                     need a smal border to overlap the edge in pf_text image
'1.13_3 oqqsan - improved hatch cover one more time =)  It is getting close, will it be good enough ? the image can be improved alittle more on some edges if we go for it :)
'1.14 iaakki - added insert outlines back and removed some rollover reflections, tracyskin color change, adding some insert reflections, top lane insert bug with right orb fixed
'1.15 iaakki - MultiBallReady issue in multiplayer mode fixed (thanks Circusfreak). ShipCannonKicker DOF 106 added. CaptiveBall tracker improved and moved to missiontimer (Thanks SeaberJenkins), Narnia balls checked after 2 and 7 secs if no switch touched. ShitTalk algorithm added
'1.17 iaakki - PupilLatency redone, SlingSpin feature added, songtimer compined to gametimer
'1.18 iaakki - Skitso style pentagon inserts and some other reflection tuning
'1.19 iaakki - minor cleanup and bloom adjust
'1.20 iaakki - Insert bloom level sync issue fix as sometimes inserts appeared too hot, MagnetMode checks added to ball search, Thanks Circusfreak for finding these.
'1.21 iaakki - More lighting fixes..








'********************* START OF PUPDMD FRAMEWORK v1.0 *************************
'******************** DO NOT MODIFY STUFF BELOW   THIS LINE!!!! ***************
'******************************************************************************
'*****   Create a PUPPack within PUPPackEditor for layout config!!!  **********
'******************************************************************************
'
'
'  Quick Steps:
'      1>  create a folder in PUPVideos with Starter_PuPPack.zip and call the folder "yourgame"
'      2>  above set global variable pGameName="yourgame"
'      3>  copy paste the settings section above to top of table script for user changes.
'      4>  on Table you need to create ONE timer only called pupDMDUpdate and set it to 250 ms enabled on startup.
'      5>  go to your table1_init or table first startup function and call PUPINIT function
'      6>  Go to bottom on framework here and setup game to call the appropriate events like pStartGame (call that in your game code where needed)...etc
'      7>  attractmodenext at bottom is setup for you already,  just go to each case and add/remove as many as you want and setup the messages to show.
'      8>  Have fun and use pDMDDisplay(xxxx)  sub all over where needed.  remember its best to make a bunch of mp4 with text animations... looks the best for sure!
'
'
'Note:  for *Future Pinball* "pupDMDupdate_Timer()" timer needs to be renamed to "pupDMDupdate_expired()"  and then all is good.
'       and for future pinball you need to add the follow lines near top
'Need to use BAM and have com idll enabled.
'       Dim icom : Set icom = xBAM.Get("icom") ' "icom" is name of "icom.dll" in BAM\Plugins dir
'       if icom is Nothing then MSGBOX "Error cannot run without icom.dll plugin"
'       Function CreateObject(className)
'             Set CreateObject = icom.CreateObject(className)
'       End Function




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
Const pPopUP2=11


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
  END IF

  pSetPageLayouts

  pDMDSetPage(pDMDBlank)   'set blank text overlay page.
  pDMDStartUP        ' firsttime running for like an startup video..
End Sub 'end PUPINIT



'PinUP Player DMD Helper Functions

Sub pDMDLabelHide(labName)
  If HasPuP=False then exit Sub
  PuPlayer.LabelSet pDMD,labName,"",0,""
end sub




Sub pDMDScrollBig(msgText,timeSec,mColor)
  If HasPuP=False then exit Sub
  PuPlayer.LabelShowPage pDMD,2,timeSec,""
  PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':" & (timeSec*1000000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
end sub

Sub pDMDScrollBigV(msgText,timeSec,mColor)
  If HasPuP=False then exit Sub
  PuPlayer.LabelShowPage pDMD,2,timeSec,""
  PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':2,'yps':1,'ype':-1,'len':" & (timeSec*1000000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
end sub


Sub pDMDSplashScore(msgText,timeSec,mColor)
  If HasPuP=False then exit Sub
  PuPlayer.LabelSet pDMD,"MsgScore",msgText,0,"{'mt':1,'at':1,'fq':250,'len':"& (timeSec*1000) &",'fc':" & mColor & "}"
end Sub

Sub pDMDSplashScoreScroll(msgText,timeSec,mColor)
  If HasPuP=False then exit Sub
  PuPlayer.LabelSet pDMD,"MsgScore",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':"& (timeSec*1000) &", 'mlen':"& (timeSec*1000) &",'tt':0, 'fc':" & mColor & "}"
end Sub

Sub pDMDZoomBig(msgText,timeSec,mColor)  'new Zoom
  If HasPuP=False then exit Sub
  PuPlayer.LabelShowPage pDMD,2,timeSec,""
  PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':3,'hstart':5,'hend':80,'len':" & (timeSec*1000) & ",'mlen':" & (timeSec*500) & ",'tt':5,'fc':" & mColor & "}"
end sub

Sub pDMDTargetLettersInfo(msgText,msgInfo, timeSec)  'msgInfo = '0211'  0= layer 1, 1=layer 2, 2=top layer3.
'this function is when you want to hilite spelled words.  Like B O N U S but have O S hilited as already hit markers... see example.
  If HasPuP=False then exit Sub
  PuPlayer.LabelShowPage pDMD,5,timeSec,""  'show page 5
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

  PuPlayer.LabelSet pDMD,"Back5"  ,backText  ,1,""
  PuPlayer.LabelSet pDMD,"Middle5",middleText,1,""
  PuPlayer.LabelSet pDMD,"Flash5" ,flashText ,0,"{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) & "}"
end Sub


Sub pDMDSetPage(pagenum)
  If HasPuP=False then exit Sub
    PuPlayer.LabelShowPage pDMD,pagenum,0,""   'set page to blank 0 page if want off
    PDMDCurPage=pagenum
end Sub

Sub pHideOverlayText(pDisp)
  If HasPuP=False then exit Sub
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "& pDisp &", ""FN"": 34 }"             'hideoverlay text during next videoplay on DMD auto return
end Sub



Sub pDMDShowLines3(msgText,msgText2,msgText3,timeSec)
  If HasPuP=False then exit Sub
  Dim vis:vis=1
  if pLine1Ani<>"" Then vis=0
  PuPlayer.LabelShowPage pDMD,3,timeSec,""
  PuPlayer.LabelSet pDMD,"Splash3a",msgText,vis,pLine1Ani
  PuPlayer.LabelSet pDMD,"Splash3b",msgText2,vis,pLine2Ani
  PuPlayer.LabelSet pDMD,"Splash3c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDShowLines2(msgText,msgText2,timeSec)
  If HasPuP=False then exit Sub
  Dim vis:vis=1
  if pLine1Ani<>"" Then vis=0
  PuPlayer.LabelShowPage pDMD,4,timeSec,""
  PuPlayer.LabelSet pDMD,"Splash4a",msgText,vis,pLine1Ani
  PuPlayer.LabelSet pDMD,"Splash4b",msgText2,vis,pLine2Ani
end Sub

Sub pDMDShowCounter(msgText,msgText2,msgText3,timeSec)
  If HasPuP=False then exit Sub
  Dim vis:vis=1
  if pLine1Ani<>"" Then vis=0
  PuPlayer.LabelShowPage pDMD,6,timeSec,""
  PuPlayer.LabelSet pDMD,"Splash6a",msgText,vis, pLine1Ani
  PuPlayer.LabelSet pDMD,"Splash6b",msgText2,vis,pLine2Ani
  PuPlayer.LabelSet pDMD,"Splash6c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDShowBig(msgText,timeSec, mColor)
  If HasPuP=False then exit Sub
  Dim vis:vis=1
  if pLine1Ani<>"" Then vis=0
  PuPlayer.LabelShowPage pDMD,2,timeSec,""
  PuPlayer.LabelSet pDMD,"Splash",msgText,vis,pLine1Ani
end sub


Sub pDMDShowHS(msgText,msgText2,msgText3,timeSec) 'High Score
  If HasPuP=False then exit Sub
  Dim vis:vis=1
  if pLine1Ani<>"" Then vis=0
  PuPlayer.LabelShowPage pDMD,7,timeSec,""
  PuPlayer.LabelSet pDMD,"Splash7a",msgText,vis,pLine1Ani
  PuPlayer.LabelSet pDMD,"Splash7b",msgText2,vis,pLine2Ani
  PuPlayer.LabelSet pDMD,"Splash7c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDSetBackFrame(fname)
  If HasPuP=False then exit Sub
  PuPlayer.playlistplayex pDMD,"PUPFrames",fname,0,1
end Sub

Sub pDMDStartBackLoop(fPlayList,fname)
  If HasPuP=False then exit Sub
  PuPlayer.playlistplayex pDMD,fPlayList,fname,0,1
  PuPlayer.SetBackGround pDMD,1
end Sub

Sub pDMDStopBackLoop
  If HasPuP=False then exit Sub
  PuPlayer.SetBackGround pDMD,0
  PuPlayer.playstop pDMD
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
DIM pDMDVideoPlaying: pDMDVideoPlaying=false


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
  If HasPuP=False then exit Sub
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

  if pDMDVideoPlaying Then
    PuPlayer.playstop pDMD
    pDMDVideoPlaying=False
  End if


  if (VideoName<>"") and (useDMDVideos) Then  'we are showing a splash video instead of the text.

    PuPlayer.playlistplayex pDMD,"DMDSplash",VideoName,pDMDDefVolume,pPriority  'should be an attract background (no text is displayed)
    pDMDVideoPlaying=true
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

'dim cUpdateCntr : cUpdateCntr = 0
Sub pupDMDupdate_Timer()
  If HasPuP=False then exit Sub
  pUpdateScores

' msgbox gametime - cUpdateCntr
' cUpdateCntr = gametime

    if PriorityReset>0 Then  'for splashes we need to reset current prioirty on timer
       PriorityReset=PriorityReset-pupDMDUpdate.interval
       if PriorityReset<=0 Then
            pDMDCurPriority=-1
            if pInAttract then pAttractReset=pAttractBetween ' pAttractNext  call attract next after 1 second
      pDMDVideoPlaying=false
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
  If HasPuP=False then exit Sub
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
  If HasPuP=False then exit Sub

  DIM dmddef
  DIM dmdalt
  DIM dmdscr
  DIM dmdfixed
  DIM pDMD_FontScale

  'labelNew <screen#>, <Labelname>, <fontName>,<size%>,<colour>,<rotation>,<xalign>,<yalign>,<xpos>,<ypos>,<PageNum>,<visible>
  '***********************************************************************'
  '<screen#>, in standard we’d set this to pDMD ( or 1)
  '<Labelname>, your name of the label. keep it short no spaces (like 8 chars) although you can call it anything really. When setting the label you will use this labelname to access the label.
  '<fontName> Windows font name, this must be exact match of OS front name. if you are using custom TTF fonts then double check the name of font names.
  '<size%>, Height as a percent of display height. 20=20% of screen height.
  '<colour>, integer value of windows color.
  '<rotation>, degrees in tenths   (900=90 degrees)
  '<xAlign>, 0= horizontal left align, 1 = center horizontal, 2= right horizontal
  '<yAlign>, 0 = top, 1 = center, 2=bottom vertical alignment
  '<xpos>, this should be 0, but if you want to ‘force’ a position you can set this. it is a % of horizontal width. 20=20% of screen width.
  '<ypos> same as xpos.
  '<PageNum> IMPORTANT… this will assign this label to this ‘page’ or group.
  '<visible> initial state of label. visible=1 show, 0 = off.


  if PuPDMDDriverType=pDMDTypeFULL THEN  'Using FULL BIG LCD PuPDMD  ************ lcd **************

    'dmddef="Impact"
    dmdalt="Rogue Hero"
    dmdfixed="Rogue Hero"
    dmdscr="Rogue Hero"  'main score font
    dmddef="Compacta Black"
    pDMD_FontScale=1

    'Page 1 (default score display)
    PuPlayer.LabelNew pDMD,"Play1"   ,dmdscr,(6*pDMD_FontScale),16777215  ,0,1,1,14,87,1,0
    PuPlayer.LabelNew pDMD,"Ball"    ,dmdscr,(5*pDMD_FontScale),16777215  ,0,1,1,89,5,1,0
    PuPlayer.LabelNew pDMD,"MsgScore",dmddef,(45*pDMD_FontScale),33023   ,0,1,0, 0,40,1,0
    PuPlayer.LabelNew pDMD,"CurScore1",dmdscr,(7*pDMD_FontScale), 65535 ,0,1,1, 17,93,1,0
    PuPlayer.LabelNew pDMD,"CurScore2",dmdscr,(7*pDMD_FontScale), 65535 ,0,1,1, 39,93,1,0
    PuPlayer.LabelNew pDMD,"CurScore3",dmdscr,(7*pDMD_FontScale), 65535 ,0,1,1, 61,93,1,0
    PuPlayer.LabelNew pDMD,"CurScore4",dmdscr,(7*pDMD_FontScale), 65535 ,0,1,1, 83,93,1,0
    PuPlayer.LabelNew pDMD,"Smallscore1",dmdscr,(4*pDMD_FontScale), 11842740 ,0,1,1, 12,90,1,0
    PuPlayer.LabelNew pDMD,"Smallscore2",dmdscr,(4*pDMD_FontScale), 11842740 ,0,1,1, 44,90,1,0
    PuPlayer.LabelNew pDMD,"Smallscore2l",dmdscr,(4*pDMD_FontScale), 11842740 ,0,1,1, 34,90,1,0
    PuPlayer.LabelNew pDMD,"Smallscore3",dmdscr,(4*pDMD_FontScale), 11842740 ,0,1,1, 66,90,1,0
    PuPlayer.LabelNew pDMD,"Smallscore3l",dmdscr,(4*pDMD_FontScale), 11842740 ,0,1,1, 56,90,1,0
    PuPlayer.LabelNew pDMD,"Smallscore4",dmdscr,(4*pDMD_FontScale), 11842740 ,0,1,1, 88,90,1,0
    PuPlayer.LabelNew pDMD,"Credits",dmdscr,(5*pDMD_FontScale), 16777215 ,0,1,1, 11,5,1,0
    PuPlayer.LabelNew pDMD,"M1",dmdscr,(10*pDMD_FontScale), 263340 ,0,1,1,86,22,1,0
    PuPlayer.LabelNew pDMD,"M1P",dmddef,(10*pDMD_FontScale),  263340,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M1PC",dmddef,(10*pDMD_FontScale), 16753873 ,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M2",dmdscr,(10*pDMD_FontScale), 263340,0,1,1,18,35,1,0
    PuPlayer.LabelNew pDMD,"M2P",dmddef,(10*pDMD_FontScale),  263340,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M2PC",dmddef,(10*pDMD_FontScale), 16753873,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M3",dmdscr,(10*pDMD_FontScale), 263340 ,0,1,1,86,44,1,0
    PuPlayer.LabelNew pDMD,"M3P",dmddef,(10*pDMD_FontScale),  263340,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M3PC",dmddef,(10*pDMD_FontScale), 16753873,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M4",dmdscr,(10*pDMD_FontScale), 263340 ,0,1,1,63,28,1,0
    PuPlayer.LabelNew pDMD,"M4P1",dmddef,(10*pDMD_FontScale),  263340,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M4P2",dmddef,(10*pDMD_FontScale),  263340,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M4PC",dmddef,(10*pDMD_FontScale), 16753873,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M5P",dmddef,(10*pDMD_FontScale),  16777215,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M5PC",dmddef,(10*pDMD_FontScale), 16753873,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M6",dmdscr,(9*pDMD_FontScale), 263340,0,1,1,18,37,1,0
    PuPlayer.LabelNew pDMD,"M6P",dmddef,(10*pDMD_FontScale),  263340,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M6PC",dmddef,(10*pDMD_FontScale), 16753873,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M7P",dmddef,(10*pDMD_FontScale),  263340,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"M7PC",dmddef,(10*pDMD_FontScale), 16753873,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"PS1",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"PS2",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"PS2B",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,17,1,1
    PuPlayer.LabelNew pDMD,"PS3",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"PS4",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"PS5",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,70,1,1
    PuPlayer.LabelNew pDMD,"PS6",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,70,1,1
    PuPlayer.LabelNew pDMD,"PS7",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,70,1,1
    PuPlayer.LabelNew pDMD,"PS8",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,70,1,1
    PuPlayer.LabelNew pDMD,"PS9",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,70,1,1
    PuPlayer.LabelNew pDMD,"PS10",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,70,1,1
    PuPlayer.LabelNew pDMD,"PS11",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,70,1,1
    PuPlayer.LabelNew pDMD,"PS12",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,70,1,1
    PuPlayer.LabelNew pDMD,"PS9B",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,17,1,1
    PuPlayer.LabelNew pDMD,"GS",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,17,1,1
    PuPlayer.LabelNew pDMD,"Player",dmddef,(10*pDMD_FontScale), 16777215 ,0,1,1, 50,67,1,1
    PuPlayer.LabelNew pDMD,"Scorlogo",dmddef,(10*pDMD_FontScale), 263340 ,0,1,1, 50,53,1,1
    PuPlayer.LabelNew pDMD,"Mcountdown" ,dmdscr,(5*pDMD_FontScale), 59624 ,0,1,1,83,76,1,1


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
    PuPlayer.LabelNew pDMD,"Splash6b",dmddef,50,33023,0,1,0,60,0,6,0
    PuPlayer.LabelNew pDMD,"Splash6c",dmddef,40,33023,0,1,0,60,50,6,0

    'Page 7 (Show High Scores Fixed Fonts)
    PuPlayer.LabelNew pDMD,"Attract 1 Score",dmddef,(20*pDMD_FontScale),8454143,0,1,0,0,2,7,0
    PuPlayer.LabelNew pDMD,"Splash7b",dmdfixed,(40*pDMD_FontScale),33023,0,1,0,0,20,7,0
    PuPlayer.LabelNew pDMD,"Splash7c",dmdfixed,(40*pDMD_FontScale),33023,0,1,0,0,50,7,0

    'Page 8 (Sequence Text 3 Lines)
    PuPlayer.LabelNew pDMD,"Bonus",dmddef,(20*pDMD_FontScale),16777215,0,1,1, 50,30,8,0
    PuPlayer.LabelNew pDMD,"Btotal",dmddef,(20*pDMD_FontScale),16777215,0,1,1, 50,60,8,0
    PuPlayer.LabelNew pDMD,"Total",dmddef,(20*pDMD_FontScale),16777215,0,1,1, 50,30,8,0

    'Page 9 (Attract Text 3 Lines)
    PuPlayer.LabelNew pDMD,"AttractA",dmdalt,(18*pDMD_FontScale), 16777215 ,0,1,1,50,50,9,0
    PuPlayer.LabelNew pDMD,"Player1",dmddef,(15*pDMD_FontScale),183,0,1,1,50,35,9,0
    PuPlayer.LabelNew pDMD,"HSNUMB",dmddef,(15*pDMD_FontScale),16777215,0,1,1,50,45,9,0
    PuPlayer.LabelNew pDMD,"GRANDCHAMP",dmddef,(15*pDMD_FontScale),16777215,0,1,1,50,30,9,0
    PuPlayer.LabelNew pDMD,"HSINIT",dmddef,(20*pDMD_FontScale),16777215,0,1,1,50,62,9,0
    PuPlayer.LabelNew pDMD,"CreditsATT",dmdscr,(6*pDMD_FontScale), 16777215 ,0,1,1,90,96,9,0
    PuPlayer.LabelNew pDMD,"BonusDrain",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,30,9,0
    PuPlayer.LabelNew pDMD,"BonusMissions",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,50,9,0
    PuPlayer.LabelNew pDMD,"BonusMissions1",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,70,9,0
    PuPlayer.LabelNew pDMD,"Ramptotal",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,58,9,0
    PuPlayer.LabelNew pDMD,"Ramps1",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,78,9,0
    PuPlayer.LabelNew pDMD,"ScoreTotal",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,50,9,0
    PuPlayer.LabelNew pDMD,"MimaLoop3",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,55,9,0
    PuPlayer.LabelNew pDMD,"MimaLoopTotal",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,75,9,0
    PuPlayer.LabelNew pDMD,"BonusDrainTotal",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,60,9,0
    PuPlayer.LabelNew pDMD,"Multi",dmdscr,(25*pDMD_FontScale), 16777215 ,0,1,1,50,55,9,0
    PuPlayer.LabelNew pDMD,"Bonus Multiplier",dmdscr,(25*pDMD_FontScale), 16777215 ,0,1,1,50,55,9,0
    PuPlayer.LabelNew pDMD,"Bonus Multiplier Total",dmdscr,(15*pDMD_FontScale), 16777215 ,0,1,1,50,50,9,0
    PuPlayer.LabelNew pDMD,"Credits2",dmdscr,(6*pDMD_FontScale), 16777215 ,0,1,1, 49,5,9,0
    PuPlayer.LabelNew pDMD,"GameOver",dmdscr,(10*pDMD_FontScale), 16777215 ,0,1,1, 50,87,9,0
    PuPlayer.LabelNew pDMD,"Error1"   ,dmdscr,(6*pDMD_FontScale),459262 ,0,1,1,51,87,9,0
    PuPlayer.LabelNew pDMD,"InsertCoin1",dmdscr,(6*pDMD_FontScale), 459262 ,0,1,1,86,87,9,0
    PuPlayer.LabelNew pDMD,"HIGHSCORE1P",dmdscr,(10*pDMD_FontScale), 16777215 ,0,1,1,50,55,9,0
    PuPlayer.LabelNew pDMD,"HIGHSCORE2P",dmdscr,(10*pDMD_FontScale), 16777215 ,0,1,1,50,75,9,0
    PuPlayer.LabelNew pDMD,"HSSCORE",dmdscr,(20*pDMD_FontScale), 65021 ,0,1,1,50,30,9,0

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
  If HasPuP=False then exit Sub
  pInAttract=false
  pDMDSetPage(pScores)
  pDMDmode="default"
  pDMD_CurSequencePos=0
  pDMD_Sequence.Enabled=false
    PuPEvent(500)
  Dof 518, DOFOff
end Sub


Sub pDMDGameOver
  If HasPuP=False then exit Sub
  pDMDSetPage(9)
  pDMD_CurSequencePos=0
  pDMD_Sequence.Interval = 500
  pDMD_Sequence.Enabled=true
end Sub

Sub pAttractStart
  If HasPuP=False then exit Sub
  pDMDSetPage(0)   'set blank text overlay page.
  pDMDmode="attract"
  pDMD_CurSequencePos=0
  pDMD_Sequence.Interval = 500
  pDMD_Sequence.Enabled=true
end Sub

Sub pDMDStartUP
  If HasPuP=False then exit Sub
  'pDMDSetPage(9):puPlayer.LabelSet pDMD,"Credits2"," Credits: " & Credits,1,""
  'puPlayer.LabelSet pDMD,"GameOver","Game Over" ,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  'pDMDSetPage(9)
  pDMDmode="attract"
  pDMD_CurSequencePos=0
  pDMD_Sequence.Interval = 50
  pDMD_Sequence.Enabled=true
end Sub



'********************** DMD *************************************************************

Dim pDMD_CurSequencePos:pDMD_CurSequencePos=0

Dim pDMDmode: pDMDmode="default"

Sub pDMD_Sequence_Timer
  If HasPuP=False then exit Sub
  PuPlayer.playlistplayex pMusic2,"PuPOverlays","pupdmd.png",100,1
  pDMDSetPage(9):puPlayer.LabelSet pDMD,"Credits2"," Credits: " & Credits,1,""
  pDMDSetPage(9):puPlayer.LabelSet pDMD,"GameOver","Game Over" ,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  pDMD_CurSequencePos=pDMD_CurSequencePos+1
  if pDMDmode="attract" then
    Select Case pDMD_CurSequencePos
    Case 1 pDMD_Sequence.Interval =2
    DOF 518, DOFOn
    Case 2 pDMD_Sequence.Interval = 27000
    Case 3 PuPEvent(406):pDMD_Sequence.Interval = 1000
    Case 4 pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(0)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","GRAND CHAMPION",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(0),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 5 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(1)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 1",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(1),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 6 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 7 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval =1000

    Case Else
    PuPEvent(518)
    pDMDSetPage(9)
    pDMD_CurSequencePos=0
    end Select
  end if
   if pDMDmode="go1" then
    Select Case pDMD_CurSequencePos
    Case 1 PuPEvent(405) :pDMD_Sequence.Interval = 500
    Case 2 pDMDSetPage(9):puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 1",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 3 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(1)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 4",0,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 20
    Case 4 PuPEvent(406):pDMD_Sequence.Interval = 1000
    Case 5 pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(0)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","GRAND CHAMPION",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(0),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 6 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(1)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 1",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(1),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 7 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 8 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval =1000
    Case 9 PuPEvent (777):pDMD_Sequence.Interval = 34000
    Case 10 pDMD_Sequence.Interval = 1
    Case Else
    pDMDSetPage(9)
    pDMD_CurSequencePos=0
    end Select
  end if

   if pDMDmode="go2" then
    Select Case pDMD_CurSequencePos
    Case 1 PuPEvent(405) :pDMD_Sequence.Interval = 500
    Case 2 pDMDSetPage(9):puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 1",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 3 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 2",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }"
    Case 4 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(2)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 4",0,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 20
    Case 5 PuPEvent(406):pDMD_Sequence.Interval = 1000
    Case 6 pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(0)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","GRAND CHAMPION",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(0),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 7 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(1)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 1",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(1),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 8 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 9 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval =1000
    Case 10 PuPEvent (777):pDMD_Sequence.Interval = 34000
    Case 11 pDMD_Sequence.Interval = 1

    Case Else
    pDMDSetPage(9)
    pDMD_CurSequencePos=0
    end Select
  end if

   if pDMDmode="go3" then
    Select Case pDMD_CurSequencePos
    Case 1 PuPEvent(405) :pDMD_Sequence.Interval = 500
    Case 2 pDMDSetPage(9):puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 1",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 3 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 2",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }"
    Case 4 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(3)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 3",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }"
    Case 5 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(3)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 4",0,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 20
    Case 6 PuPEvent(406):pDMD_Sequence.Interval = 1000
    Case 7 pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(0)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","GRAND CHAMPION",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(0),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 8 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(1)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 1",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(1),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 9 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
    Case 10 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval =1000
    Case 11 PuPEvent (777):pDMD_Sequence.Interval = 34000
    Case 12 pDMD_Sequence.Interval = 1

    Case Else
    pDMDSetPage(9)
    pDMD_CurSequencePos=0
    end Select
  end if

   if pDMDmode="go4" then
    Select Case pDMD_CurSequencePos
      Case 1 PuPEvent(405) :pDMD_Sequence.Interval = 500
      Case 2 pDMDSetPage(9):puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 1",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
      Case 3 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 2",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }"
      Case 4 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(3)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 3",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }"
      Case 5 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(4)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 4",1,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }"
      Case 6 puPlayer.LabelSet pDMD,"AttractA",""& FormatScore(Score(4)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':1,'xoffset':2, 'yoffset':6, 'bold':1, 'outline':2 }":puPlayer.LabelSet pDMD,"Player1","PLAYER 4",0,"{'mt':2, 'shadowcolor':7929917, 'shadowstate':1,'xoffset':2, 'yoffset':4, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 20
      Case 7 PuPEvent(406):pDMD_Sequence.Interval = 1000
      Case 8 pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(0)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","GRAND CHAMPION",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(0),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
      Case 9 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(1)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 1",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(1),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
      Case 10 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),1,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval = 3000
      Case 11 puPlayer.LabelSet pDMD,"HSNUMB",""& FormatScore(HighScore(2)),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"GRANDCHAMP","HIGH SCORE: 2",0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMDSetPage(9):puPlayer.LabelSet pDMD,"HSINIT",""& HighScoreName(2),0,"{'mt':2,'shadowcolor':1184274, 'shadowstate':2,'xoffset':2, 'yoffset':3, 'bold':1, 'outline':2 }":pDMD_Sequence.Interval =1000
      Case 12 PuPEvent (777):pDMD_Sequence.Interval = 34000
      Case 13 pDMD_Sequence.Interval = 1

      Case Else
      pDMDSetPage(9)
      pDMD_CurSequencePos=0
    end Select
  end if


End Sub




'************************ called during gameplay to update Scores ***************************

Dim PupMission : PupMission = "Default"
Dim SkillPupFlash: SkillPupFlash = "Default"
Dim ScoreMissionPoints: ScoreMissionPoints = "Default"
Dim ScoreMissionPointsC: ScoreMissionPointsC = "Default"
Dim ScoreMissionPointsC1:ScoreMissionPointsC1 = "Default"
Dim MBPUP : MBPUP = "Default"
Dim GunshotPUP: GunshotPUP = "Default"
Dim Gameplayerup: Gameplayerup = "Default"
Dim GameplayerNotup: GameplayerNotup = "Default"
Dim Smalldisplayscore:Smalldisplayscore = "Default"
Dim Bstage:Bstage = "Default"
Dim Mtimercountdown: Mtimercountdown = "Default"

dim PupMissionDefault : PupMissionDefault = 0
Dim ScoreMissionPointsDefault: ScoreMissionPointsDefault = 0
Dim SkillPupFlashDefault: SkillPupFlashDefault = 0

Sub pUpdateScores  'call this ONLY on timer 300ms is good enough
  If HasPuP=False then exit Sub
  if pDMDCurPage <> pScores then Exit Sub

  puPlayer.LabelSet pDMD,"Ball","Ball: " & (bpgcurrent - BallsRemaining(CurrentPlayer) + 1),1,""
  puPlayer.LabelSet pDMD,"Credits","Credits: " & Credits,1,""



  'missions
  if PupMission <> "Default" then
'   debug.print "PupMission: " & PupMission
    Select Case PupMission
      Case "PMission1" : puPlayer.LabelSet pDMD,"M1","" & Mission1TargetPupCount,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "PMission2" : puPlayer.LabelSet pDMD,"M2","" & Mission1TargetPupCount2,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':0, 'yoffset':3, 'bold':1, 'outline':2 }"
      Case "PMission3" : puPlayer.LabelSet pDMD,"M3","" & Mission1TargetPupCount3,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "PMission4" : puPlayer.LabelSet pDMD,"M4","" & Mission1TargetPupCount4,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "PMission6" : puPlayer.LabelSet pDMD,"M6","" & Mission1TargetPupCount6,1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    end select
    PupMissionDefault = 0
  Else
    if PupMissionDefault = 0 then
      'default only once
      puPlayer.LabelSet pDMD,"M1","" & Mission1TargetPupCount,0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M2","" & Mission1TargetPupCount,0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M3","" & Mission1TargetPupCount,0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M4","" & Mission1TargetPupCount,0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M6","" & Mission1TargetPupCount,0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      PupMissionDefault = 1
    end if
  end if


  'mission points
  if ScoreMissionPoints <> "Default" Then
'   debug.print "ScoreMissionPoints: " & ScoreMissionPoints
    Select Case ScoreMissionPoints
      Case "1" : puPlayer.LabelSet pDMD,"M1P","1,000,000",1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "2" : puPlayer.LabelSet pDMD,"M2P",""& FormatNumber((LoopPupScore),0),1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "3" : puPlayer.LabelSet pDMD,"M3P",""& FormatNumber((ChasePupScore),0),1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "4a" : puPlayer.LabelSet pDMD,"M4P1","3,000,000",1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "4b" : puPlayer.LabelSet pDMD,"M4P2","1,500,000",1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "4c" : puPlayer.LabelSet pDMD,"M4PC","3,000,000",1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "4c2" : puPlayer.LabelSet pDMD,"M4PC","1,500,000",1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "5" : puPlayer.LabelSet pDMD,"M5P",""& FormatNumber((PortalPupScore),0),1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "5c" : puPlayer.LabelSet pDMD,"M5PC",""& FormatNumber((PortalPupScore),0) ,1,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "6" : puPlayer.LabelSet pDMD,"M6P","100,000",1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "6c" : puPlayer.LabelSet pDMD,"M6PC","3,000,000",1,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "7" : puPlayer.LabelSet pDMD,"M7P","500,000",1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      Case "7c" : puPlayer.LabelSet pDMD,"M7PC","3,000,000",1,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
    end select
    ScoreMissionPointsDefault = 0
  Else
    if ScoreMissionPointsDefault = 0 Then
      puPlayer.LabelSet pDMD,"M1P","",0,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M2P","",0,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M3P","",0,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M4P1","",0,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M4P2","",0,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M4PC","",0,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M5P","",0,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M5PC","" ,0,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M6P","",0,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M6PC","" ,0,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M7P","",0,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
      puPlayer.LabelSet pDMD,"M7PC","" ,0,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
      ScoreMissionPointsDefault = 1
    end if
  end if

  If ScoreMissionPointsC = "1C" Then
    puPlayer.LabelSet pDMD,"M1PC","3,000,000",1,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf ScoreMissionPointsC = "2C" Then
    puPlayer.LabelSet pDMD,"M2PC","3,000,000",1,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
  Elseif ScoreMissionPointsC = "3C" Then
    puPlayer.LabelSet pDMD,"M3PC","2,500,000",1,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf ScoreMissionPointsC = "Default" Then
    puPlayer.LabelSet pDMD,"M2PC","" ,0,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"M1PC","" ,0,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"M3PC","" ,0,"{'mt':2, 'shadowcolor':7667770, 'shadowstate':2,'xoffset':1, 'yoffset':8, 'bold':1, 'outline':2 }"
  End If

  'SkillShot
  if SkillPupFlash <> "Default" Then
'   debug.print "SkillPupFlash: " & SkillPupFlash
    Select Case SkillPupFlash
      Case "1" : puPlayer.Labelset pDMD,"PS1","500,000",0,"{'mt':1, 'at':1,'fq':110, 'len':999, 'fc':53760}"
      Case "2"
        puPlayer.Labelset pDMD,"PS2","600,000",0,"{'mt':1, 'at':1,'fq':110, 'len':999, 'fc':53760}"
        puPlayer.Labelset pDMD,"PS2B","MIMA LOCK READY",0,"{'mt':1, 'at':1,'fq':110, 'len':999, 'fc':53760}"
      Case "3" : puPlayer.Labelset pDMD,"PS3","700,000",0,"{'mt':1, 'at':1,'fq':110, 'len':999, 'fc':53760}"
      Case "4" : puPlayer.Labelset pDMD,"PS4","800,000",0,"{'mt':1, 'at':1,'fq':110, 'len':999, 'fc':53760}"
      Case "5" : puPlayer.Labelset pDMD,"PS5","2,000,000",0,"{'mt':1, 'at':1,'fq':80, 'len':999, 'fc':13473944}"
      Case "6" : puPlayer.Labelset pDMD,"PS6","2,200,000",0,"{'mt':1, 'at':1,'fq':80, 'len':999, 'fc':13473944}"
      Case "7" : puPlayer.Labelset pDMD,"PS7","2,400,000",0,"{'mt':1, 'at':1,'fq':80, 'len':999, 'fc':13473944}"
      Case "8" : puPlayer.Labelset pDMD,"PS8","2,600,000",0,"{'mt':1, 'at':1,'fq':80, 'len':999, 'fc':13473944}"
      Case "9"
        puPlayer.Labelset pDMD,"PS9","4,000,000",0,"{'mt':1, 'at':1,'fq':80, 'len':999, 'fc':13473944}"
        puPlayer.Labelset pDMD,"PS9B","EXTRA BALL",0,"{'mt':1, 'at':1,'fq':80, 'len':999, 'fc':13473944}"
      Case "10" : puPlayer.Labelset pDMD,"PS10","4,400,000",0,"{'mt':1, 'at':1,'fq':80, 'len':999, 'fc':13473944}"
      Case "11" : puPlayer.Labelset pDMD,"PS11","4,800,000",0,"{'mt':1, 'at':1,'fq':80, 'len':999, 'fc':13473944}"
      Case "12" : puPlayer.Labelset pDMD,"PS12","5,200,000",0,"{'mt':1, 'at':1,'fq':80, 'len':999, 'fc':13473944}"
    end select
    SkillPupFlashDefault = 0
  Else
    if SkillPupFlashDefault = 0 Then
      puPlayer.Labelset pDMD,"PS1","",0,""
      puPlayer.Labelset pDMD,"PS2","",0,""
      puPlayer.Labelset pDMD,"PS2B","",0,""
      puPlayer.Labelset pDMD,"PS3","",0,""
      puPlayer.Labelset pDMD,"PS4","",0,""
      puPlayer.Labelset pDMD,"PS5","",0,""
      puPlayer.Labelset pDMD,"PS6","",0,""
      puPlayer.Labelset pDMD,"PS7","",0,""
      puPlayer.Labelset pDMD,"PS8","",0,""
      puPlayer.Labelset pDMD,"PS9","",0,""
      puPlayer.Labelset pDMD,"PS9B","",0,""
      puPlayer.Labelset pDMD,"PS10","",0,""
      puPlayer.Labelset pDMD,"PS11","",0,""
      puPlayer.Labelset pDMD,"PS12","",0,""
      SkillPupFlashDefault = 1
    end if
  end if


  If Mtimercountdown= "1" Then
    puPlayer.LabelSet pDMD,"Mcountdown","MISSION TIME: "& ((MissiontimeMax-MissionTimeCurrent)),1,"{'mt':2, 'shadowcolor':66, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf Mtimercountdown= "Default"  Then
    puPlayer.LabelSet pDMD,"Mcountdown","",0,""
  End If

  '*************************************************************************
  'Gamescore "Stern Style

  '*************************************************************************

  If Gameplayerup = "1" Then
    LockPlayer2
    LockPlayer3L
    LockPlayer4l
    puPlayer.LabelSet pDMD,"CurScore1","" &  FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore1","" & FormatScore(Score(1)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"

  ElseIf Gameplayerup = "2" Then
    LockPlayer1
    LockPlayer3
    LockPlayer4M
    puPlayer.LabelSet pDMD,"CurScore2","" &  FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore2","" & FormatScore(Score(2)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf Gameplayerup = "3" Then
    LockPlayer3R
    LockPlayer4R
    puPlayer.LabelSet pDMD,"CurScore3","" &  FormatScore(Score(3)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore3","" & FormatScore(Score(3)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf Gameplayerup = "4" Then
    LockPlayer4
    puPlayer.LabelSet pDMD,"CurScore4","" & FormatScore(Score(4)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore4","" & FormatScore(Score(4)) ,0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf Gameplayerup = "0" Then
    puPlayer.LabelSet pDMD,"CurScore4","" & FormatScore(Score(4)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"CurScore3","" & FormatScore(Score(3)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"CurScore2","" &  FormatScore(Score(2)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"CurScore1","" &  FormatScore(Score(1)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore2l","" & FormatScore(Score(2)),0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End If

  'Waiting to play
  If  GameplayerNotup= "1" Then
    puPlayer.LabelSet pDMD,"Smallscore2","PRESS START",1,"{'mt':2,'color':37632, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore3","PRESS START",1,"{'mt':2,'color':37632, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore4","PRESS START",1,"{'mt':2,'color':37632, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf  GameplayerNotup= "2" Then
    puPlayer.LabelSet pDMD,"Smallscore3","PRESS START",1,"{'mt':2,'color':37632, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore4","PRESS START",1,"{'mt':2,'color':37632, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf  GameplayerNotup= "3" Then
    puPlayer.LabelSet pDMD,"Smallscore4","PRESS START",1,"{'mt':2,'color':37632, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf  GameplayerNotup = "Default" Then
    puPlayer.LabelSet pDMD,"Smallscore2","PRESS START",0,"{'mt':2,'color':37632, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore3","PRESS START",0,"{'mt':2,'color':37632, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore4","PRESS START",0,"{'mt':2,'color':37632, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf  GameplayerNotup= "6"  Then
    puPlayer.LabelSet pDMD,"Smallscore2","OFFLINE",1,"{'mt':2,'color':151, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore3","OFFLINE",1,"{'mt':2,'color':151, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore4","OFFLINE",1,"{'mt':2,'color':151, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf  GameplayerNotup= "7"  Then
    puPlayer.LabelSet pDMD,"Smallscore3","OFFLINE",1,"{'mt':2,'color':151, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    puPlayer.LabelSet pDMD,"Smallscore4","OFFLINE",1,"{'mt':2,'color':151,'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  ElseIf  GameplayerNotup= "8"  Then
    puPlayer.LabelSet pDMD,"Smallscore4","OFFLINE",1,"{'mt':2,'color':151, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if

End Sub


Sub PlayersPUP 'NextPlayer
  If HasPuP=False then exit Sub
  If CurrentPlayer =1 Then
  Gameplayerup ="1"
  PuPEvent(700)
  End If
  If CurrentPlayer =2 Then
  Gameplayerup ="2"
  PuPEvent(701)
   End If
  If CurrentPlayer =3 Then
  Gameplayerup ="3"
  PuPEvent(702)
   End If
  If CurrentPlayer =4 Then
  Gameplayerup ="4"
  PuPEvent(703)
   End If
End Sub

Sub PlayersAreInGame
  If HasPuP=False then exit Sub
  If PlayersPlayingGame =1 Then
    GameplayerNotup= "6"
  End If
  If PlayersPlayingGame =2 Then
    GameplayerNotup= "7"
  End If
  If PlayersPlayingGame =3 Then
    GameplayerNotup= "8"
  End If
  If PlayersPlayingGame =4 Then
    GameplayerNotup= "9"
  End If
End Sub


Sub AttractScoreNumbers
  If PlayersPlayingGame =1 Then
    pDMDmode="go1"
  End if
  If PlayersPlayingGame =2 Then
    pDMDmode="go2"
  End if
  If PlayersPlayingGame =3 Then
    pDMDmode="go3"
  End if
  If PlayersPlayingGame =4 Then
    pDMDmode="go4"
  End if
End Sub


Sub SmallScores  'checkStartGameKey
  If HasPuP=False then exit Sub
  If GameplayerNotup= "2" Then
    puPlayer.LabelSet pDMD,"Smallscore2","PRESS START",0,"{'mt':2,'color':11842740 , 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    Player2Signin.Enabled= True
  End If

  If GameplayerNotup= "3" Then
    puPlayer.LabelSet pDMD,"Smallscore3","PRESS START",0,"{'mt':2,'color':11842740 , 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    Player3Signin.Enabled=True
  End If

  If GameplayerNotup= "4" Then
    puPlayer.LabelSet pDMD,"Smallscore4","PRESS START",0,"{'mt':2,'color':11842740 , 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    Player4Signin.Enabled=True
  End If
End Sub

'*************************************************************************

' Locking Player text

Sub LockPlayer1
If HasPuP=False then exit Sub
  If GameplayerNotup= "7" Then
  puPlayer.LabelSet pDMD,"Smallscore1","" & FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if
End Sub

Sub LockPlayer2
If HasPuP=False then exit Sub
  If GameplayerNotup= "7" Then
  puPlayer.LabelSet pDMD,"Smallscore2","" & FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if
End Sub

Sub LockPlayer3
If HasPuP=False then exit Sub
  If GameplayerNotup= "8" Then
  puPlayer.LabelSet pDMD,"Smallscore3","" & FormatScore(Score(3)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore1","" & FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if
End Sub

Sub LockPlayer3L
If HasPuP=False then exit Sub
  If GameplayerNotup= "8" Then
  puPlayer.LabelSet pDMD,"Smallscore3","" & FormatScore(Score(3)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore2","" & FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if
End Sub

Sub LockPlayer3R
If HasPuP=False then exit Sub
  If GameplayerNotup= "8" Then
  puPlayer.LabelSet pDMD,"Smallscore1","" & FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore2l","" & FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if
End Sub

Sub LockPlayer4
If HasPuP=False then exit Sub
  If GameplayerNotup= "9" Then
  puPlayer.LabelSet pDMD,"Smallscore1","" & FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore2l","" & FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore3l","" & FormatScore(Score(3)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if
End Sub

Sub LockPlayer4l
If HasPuP=False then exit Sub
  If GameplayerNotup= "9" Then
  puPlayer.LabelSet pDMD,"Smallscore2","" & FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore3","" & FormatScore(Score(3)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore4","" & FormatScore(Score(4)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if
End Sub

Sub LockPlayer4M
If HasPuP=False then exit Sub
  If GameplayerNotup= "9" Then
  puPlayer.LabelSet pDMD,"Smallscore4","" & FormatScore(Score(4)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore1","" & FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore3","" & FormatScore(Score(3)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if
End Sub

Sub LockPlayer4R
If HasPuP=False then exit Sub
  If GameplayerNotup= "9" Then
  puPlayer.LabelSet pDMD,"Smallscore2l","" & FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore1","" & FormatScore(Score(1)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  puPlayer.LabelSet pDMD,"Smallscore4","" & FormatScore(Score(4)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End if
End Sub


'TracyBumper

Sub TracyPupCount
  If HasPuP=False then exit Sub
  BumperHitsPup(CurrentPlayer)= BumperHitsPup(CurrentPlayer) + 1
  If PupBumperTracy = True Then
    Select Case BumperHitsPup(CurrentPlayer)
      Case 1: PuPEvent(617)
      Case 2: PuPEvent(618)
      Case 3: PuPEvent(619)
      Case 4: PuPEvent(620)
      Case 5: PuPEvent(621)
      Case 6: PuPEvent(622)
      Case 7: PuPEvent(623)
      Case 8: PuPEvent(624)
      Case 9: PuPEvent(625)
      Case 10: PuPEvent(626)
      Case 11: PuPEvent(627)
      Case 12: PuPEvent(628)
      Case 13: PuPEvent(629)
      Case 14: PuPEvent(630)
      Case 15: PuPEvent(631):DOF 632, DOFPulse:PupBumpertenthousand=True:BumperStages
    end Select
  End if
End Sub


Sub BumperTenthousand
  If HasPuP=False then exit Sub
   BumperThousandPup(CurrentPlayer)= BumperThousandPup(CurrentPlayer) + 1
  If PupBumpertenthousand= True Then
    Select Case BumperThousandPup(CurrentPlayer)
      Case 1: PuPEvent(638)
      Case 2: PuPEvent(639)
      Case 3: PuPEvent(640)
      Case 4: PuPEvent(641)
      Case Else
      BumperThousandPup(CurrentPlayer)=0
    end Select
  end if
End Sub

Sub BumperTwentythousand
   If HasPuP=False then exit Sub
  BumperTwentyThousandPup(CurrentPlayer)= BumperTwentyThousandPup(CurrentPlayer) + 1
  If PupBumpersuperjets = True Then
    Select Case BumperTwentyThousandPup(CurrentPlayer)
      Case 1: PuPEvent(634)
      Case 2: PuPEvent(635)
      Case 3: PuPEvent(636)
      Case 4: PuPEvent(637)
      Case Else
      BumperTwentyThousandPup(CurrentPlayer)=0
    end Select
   end if
End Sub



Sub BumperStages
  If HasPuP=False then exit Sub
  BumperscoringPup(CurrentPlayer)= BumperscoringPup(CurrentPlayer) + 1
  Select Case BumperscoringPup(CurrentPlayer)
    Case 1: Bstage= "1"
    Case 2: Bstage= "2"
  end Select
End Sub

Sub Superjetsdrainreset
 If HasPuP=False then exit Sub
  If BumperscoringPup(CurrentPlayer)=0 then
    PupBumperTracy = True
    PupBumpertenthousand= False
    PupBumpersuperjets = False
  End if
  If BumperscoringPup(CurrentPlayer)=1 then
    PupBumpertenthousand= True
    PupBumperTracy = False
    PupBumpersuperjets = False
  End if
  If BumperscoringPup(CurrentPlayer)>=2 Then
     PupBumpersuperjets = False
     PupBumperTracy = False
     PupBumpertenthousand= True
 End If
End Sub

'Multiball

Sub TracyMBloop
If HasPuP=False then exit Sub
  If MBPUP = "1" Then
  PuPEvent 810
  End if
End Sub

Sub StopTracyMBPup
If HasPuP=False then exit Sub
  If MBPUP = "2" & bMultiBallMode = False then
  PuPEvent 811
  End If
End Sub

Sub MimaMBloop
If HasPuP=False then exit Sub
  If MBPUP = "3" Then
  PuPEvent 812
  End If
End Sub

Sub StopMimaMBPup
If HasPuP=False then exit Sub
  If MBPUP = "4" & bMultiBallMode = False Then
  PuPEvent 813
  End If
End Sub

Sub CoreyMBloop
If HasPuP=False then exit Sub
  If MBPUP = "5" Then
  PuPEvent 814
  End If
End Sub

Sub StopCoreyMBPup
If HasPuP=False then exit Sub
  If MBPUP = "6" & bMultiBallMode = False Then
  PuPEvent 815
  End If
End Sub



'Ball Save

Sub PupBallSaveMBControl
If HasPuP=False then exit Sub
  If bMultiBallMode = False Then
  PuPEvent(502)
  Dof 502, DOFPulse
  End If
  If bMultiBallMode = True then
  PuPEvent(636)
  End If
End Sub

Sub MultiStopHeart
If HasPuP=False then exit Sub
  If bMissionMode = True Then
  PuPEvent(637)
  End If
End Sub

'Players
Sub PlayerUpFull
If HasPuP=False then exit Sub
  If CurrentPlayer=1 Then
  PuPEvent 888
  DOF 888, DOFPulse
  End If
  If CurrentPlayer=2 Then
  PuPEvent 889
  DOF 889, DOFPulse
  End If
  If CurrentPlayer=3 Then
  PuPEvent 890
  DOF 890, DOFPulse
  End If
  If CurrentPlayer=4 Then
  PuPEvent 891
  DOF 891, DOFPulse
  End If
End Sub

'skip to total
'Sub SkiptoEndTotal
'If HasPuP=False then exit Sub
' If  SkipEnd = False then exit Sub
' If SkipEnd  = True Then
'  EobBonusCounter = 810
'  SkipEnd = False
'End if
'End Sub

'coin

Sub CoinErrorText
 If HasPuP= False Then Exit Sub
  If pDMDmode="attract" And Credits=0 Then
    puPlayer.LabelShowPage pDMD,9,1,""
    puPlayer.LabelSet pDMD, "Error1", "Error                                                                  Insert Coin",0,"{'mt':1, 'at':1,'fq':150, 'len':3000, 'fc':459262}"
    End if
    If pDMDmode="go1" And Credits=0 Then
    puPlayer.LabelShowPage pDMD,9,1,""
    puPlayer.LabelSet pDMD, "Error1", "Error                                                                  Insert Coin",0,"{'mt':1, 'at':1,'fq':150, 'len':3000, 'fc':459262}"
  end if
    If pDMDmode="go2" And Credits=0 Then
    puPlayer.LabelShowPage pDMD,9,1,""
    puPlayer.LabelSet pDMD, "Error1", "Error                                                                  Insert Coin",0,"{'mt':1, 'at':1,'fq':150, 'len':3000, 'fc':459262}"
  end if
    If pDMDmode="go3" And Credits=0 Then
    puPlayer.LabelShowPage pDMD,9,1,""
    puPlayer.LabelSet pDMD, "Error1", "Error                                                                  Insert Coin",0,"{'mt':1, 'at':1,'fq':150, 'len':3000, 'fc':459262}"
  end if
    If pDMDmode="go4" And Credits=0 Then
    puPlayer.LabelShowPage pDMD,9,1,""
    puPlayer.LabelSet pDMD, "Error1", "Error                                                                  Insert Coin",0,"{'mt':1, 'at':1,'fq':150, 'len':3000, 'fc':459262}"
  end if
End Sub

'Time Stop

Sub StopCoutdownPup
  Mtimercountdown= "Default"
End Sub


'High Score

Sub PupDMDHSInput
If HasPuP=False then exit Sub
  pDMDSetPage(9):puPlayer.LabelSet pDMD, "HIGHSCORE1P", "ENTER INITIALS:", 1,""
  pDMDSetPage(9):puPlayer.LabelSet pDMD,  "HSSCORE","" & FormatNumber(Score(CurrentPlayer),0),1,""
  If ( Enterblinking mod 50 ) >20 Then
    If Len(EnterName) = 0 Then pDMDSetPage(9):puPlayer.LabelSet pDMD, "HIGHSCORE2P","<" & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ",EnterNamePos,1) & ">", 1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    If Len(EnterName) = 1 Then pDMDSetPage(9):puPlayer.LabelSet pDMD, "HIGHSCORE2P",Entername & "<" & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ",EnterNamePos,1) & ">", 1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    If Len(EnterName) = 2 Then pDMDSetPage(9):puPlayer.LabelSet pDMD, "HIGHSCORE2P",Entername & "<" & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ",EnterNamePos,1) & ">", 1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    If Len(EnterName) = 3 Then pDMDSetPage(9):puPlayer.LabelSet pDMD, "HIGHSCORE2P","- " & Entername & " -", 1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  Else
    If Len(EnterName) = 0 Then pDMDSetPage(9):puPlayer.LabelSet pDMD, "HIGHSCORE2P", "- >", 1,""
    If Len(EnterName) = 1 Then pDMDSetPage(9):puPlayer.LabelSet pDMD, "HIGHSCORE2P", Entername & "- -", 1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    If Len(EnterName) = 2 Then pDMDSetPage(9):puPlayer.LabelSet pDMD, "HIGHSCORE2P",Entername & "- -", 1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
    If Len(EnterName) = 3 Then pDMDSetPage(9):puPlayer.LabelSet pDMD, "HIGHSCORE2P", "-    -", 1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
  End If
End Sub



'Pupdmd Gameplay Timers
Sub Player2Signin_Timer
If HasPuP=False then exit Sub
  Player2Signin.Enabled=False
  puPlayer.LabelSet pDMD,"Smallscore2","" & FormatScore(Score(2)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
End Sub

Sub Player3Signin_Timer
If HasPuP=False then exit Sub
  Player3Signin.Enabled=False
  puPlayer.LabelSet pDMD,"Smallscore3",""& FormatScore(Score(3)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
End Sub

Sub Player4Signin_Timer
If HasPuP=False then exit Sub
  Player4Signin.Enabled=False
  puPlayer.LabelSet pDMD,"Smallscore4",""& FormatScore(Score(4)),1,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
End Sub



Sub DrainMissionPup_Timer
If HasPuP=False then exit Sub
  DrainMissionPup.Enabled=False
  PupMission= "Default"
  DrainPupvideo.Enabled=True
End Sub

Sub DrainPupvideo_Timer
If HasPuP=False then exit Sub
  DrainPupvideo.Enabled=False
  PuPEvent(501)
  Mtimercountdown= "Default"
  Dof 501, DOFPulse
  DOF 512, DOFOff
  DOF 508, DOFOff
  DOF 509, DOFOff
  DOF 513, DOFOff
  DOF 511, DOFOff
  DOF 510, DOFOff
  DOF 507, DOFOff
  DOF 552, DOFOff
  DOF 554, DOFOff
  DOF 550, DOFOff

End Sub

Sub MissionPointsPup_Timer
If HasPuP=False then exit Sub
  MissionPointsPup.Enabled=False
  ScoreMissionPoints = "Default"
End Sub

Sub MissionPointsPupC_Timer
  MissionPointsPupC.Enabled=False
  ScoreMissionPointsC = "Default"
End Sub

Sub MBPupCheck_Timer
If HasPuP=False then exit Sub
  MBPupCheck.Enabled=False
  If bTracyMBOngoing = True Then
  PuPEvent 810
  End if
End Sub

Sub ModeTextDelay1_Timer
If HasPuP=False then exit Sub
  ModeTextDelay1.Enabled=False
  PupMission= "PMission1"
  Mtimercountdown= "1"
End Sub

Sub ModeTextDelay2_Timer
  ModeTextDelay2.Enabled=False
  PupMission= "PMission2"
  Mtimercountdown= "1"
End Sub

Sub ModeTextDelay3_Timer
If HasPuP=False then exit Sub
  ModeTextDelay3.Enabled=False
  PupMission= "PMission3"
  Mtimercountdown= "1"
End Sub

Sub ModeTextDelay4_Timer
If HasPuP=False then exit Sub
  ModeTextDelay4.Enabled=False
  PupMission= "PMission4"
  Mtimercountdown= "1"
End Sub

Sub ModeTextDelay4c2_Timer
If HasPuP=False then exit Sub
  ModeTextDelay4c2.Enabled=False
  PupMission= "Default"
  ScoreMissionPoints = "4c2"
  MissionPointsPup.Enabled=True
End Sub

Sub ModeTextDelay4c1_Timer
If HasPuP=False then exit Sub
  ModeTextDelay4c1.Enabled=False
  PupMission= "Default"
  ScoreMissionPoints = "4c"
  MissionPointsPup.Enabled=True
End Sub


Sub ModeTextDelay6_Timer
If HasPuP=False then exit Sub
  ModeTextDelay6.Enabled=False
  PupMission= "PMission6"
  Mtimercountdown= "1"
End Sub

Sub ModeTextDelay7_Timer
If HasPuP=False then exit Sub
  ModeTextDelay7.Enabled=False
  Mtimercountdown= "1"
End Sub

Sub SkillShotPup_Timer
If HasPuP=False then exit Sub
  SkillShotPup.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup2_Timer
If HasPuP=False then exit Sub
  SkillShotPup2.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup3_Timer
If HasPuP=False then exit Sub
  SkillShotPup3.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup4_Timer
If HasPuP=False then exit Sub
  SkillShotPup4.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup5_Timer
If HasPuP=False then exit Sub
  SkillShotPup5.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup6_Timer
If HasPuP=False then exit Sub
  SkillShotPup6.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup7_Timer
If HasPuP=False then exit Sub
  SkillShotPup7.Enabled=False
  SkillPupFlash="Default"
End Sub
Sub SkillShotPup8_Timer
If HasPuP=False then exit Sub
  SkillShotPup8.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup9_Timer
If HasPuP=False then exit Sub
  SkillShotPup9.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup10_Timer
If HasPuP=False then exit Sub
  SkillShotPup10.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup11_Timer
If HasPuP=False then exit Sub
  SkillShotPup11.Enabled=False
  SkillPupFlash="Default"
End Sub

Sub SkillShotPup12_Timer
If HasPuP=False then exit Sub
  SkillShotPup12.Enabled=False
  SkillPupFlash="Default"
End Sub


Sub Scorepop_Timer
 Scorepop.enabled=False
 puPlayer.LabelSet pDMD,"Player","" ,0,"{'mt':2, 'shadowcolor':2949120, 'shadowstate':2,'xoffset':2, 'yoffset':8, 'bold':1, 'outline':2 }"
puPlayer.LabelSet pDMD,"Scorlogo","Gif\\blank.gif",1,"{'mt':2,'color':111111, 'width': 35, 'height': 30., 'anigif': 130 ,}"
End Sub

' Scroll Attract
Sub ScrollAttract
  If pDMDmode="attract" then
   pDMD_Sequence.Interval = 1
  End If
  If pDMDmode="go1" then
   pDMD_Sequence.Interval = 1
  End If
  If pDMDmode="go2" then
   pDMD_Sequence.Interval = 1
  End If
  If pDMDmode="go3" then
   pDMD_Sequence.Interval = 1
  End If
  If pDMDmode="go4" then
   pDMD_Sequence.Interval = 1
  End If
End Sub

'   KKM   KKK III MMM       MMM
'   KKK  KKK    III MMMM     MMMM
'   KKK KKK     III MMM MM MM MMM
'   KKKKK       III MMM  MMM  MMM
'   KKK KKK   III MMM     MMM
'   KKK  KKK  III MMM       MMM
'   KKK   KKK III MMM       MMM
'
'
'   WWW       WWW III LLL       DDDDDD    EEEEEEEEE
'   WWW       WWW III LLL       DDDDDDDD  EEEEEEEEE
'   WWW       WWW III LLL       DDD   DDD EEE
'   WWW  WWW  WWW III LLL       DDD    DD EEEEE
'   WWW WW WW WWW III LLL       DDD   DDD EEE
'   WWWW     WWWW III LLLLLLLLL DDDDDDDD  EEEEEEEEE
'   WWW       WWW III LLLLLLLLL DDDDDD    EEEEEEEEE

' Build 4.00
'
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  INSTALLATION
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'
' This table comes without the music tracks, so music is disabled by default
' download the tracks at https://drive.google.com/drive/folders/130FA3R9u9GXIt6lqcsqgqJw9DGmmwVqy?usp=sharing
' Put the tracks in the music folder of vpx
' Then, in the options page, enable the music

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  Credits
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Design & Code by Ed Vink (Supered)
' Code snippets by JP Salas, Flupper, DJRobX, NailBuster, ninuzzu, HauntFreaks rothbauerw, Team VPW & many more.
' Based on the OrbitalPin Framework from Scott Wickberg
' Denis (Arngrim) for his DOF support

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZZXXX)
'
' ZZVOC : Vocal Queue System
'   ZZMAT : General Math Functions
'   ZZNFF : FLIPPER CORRECTIONS
'   ZZFLX : FLEX DMD
' ZZFLP : Flippers
'   ZZMBA : Manual Ball Controll
'   ZZSHA : Shadows
' ZZFLE : MECHANICAL SOUND SYSTEM
' ZZKEY : Key Press Handling
'   ZZTIL : Tilt
'   ZZ3DI : DYNAMIC INSERT LIGHTS USING PRIMITIVES
'   ZZDMP : RUBBERS AND DAMPNERS
' ZZSSC : SLINGSHOT CORRECTION FUNCTIONS
'   ZZSCR : HIGH SCORES
'   ZZATR : ATTRACT MODE
'   ZZRAI : LIGHTING / RAINBOW LIGHTS
'   ZZFLD : FLUPPER DOMES
'   ZZSCO : SCORING FUNCTIONS
'   ZZBFU : BALL FUNCTIONS & DRAINS
'   ZZBSL : BALL SAVE & LAUNCH
'   ZZINT : DMD INTRO ATTRACT MODE
'   ZZVAR : TABLE VARIABLES
'   ZZGAM : GAME STARTING & RESETS & SKILLSHOT
'   ZZSEC : SECONDARY HIT EVENTS
'   ZZLAN : LANE SHIFTING
'   ZZFLB : FLUPPER BUMPERS
'   ZZBUM : BUMPERS
'   ZZFLA : FLIPPER LANES
'   ZZKWL : KIM WILDE LAMPS & START MAGIC
'   ZZKMB : KIM MULTIBALL
'   ZZKJP : KIM JACKPOT
'   ZZKMA : KIM MAGIC
'   ZZKFS : KIM FLASHER SEQENCES
'   ZZKKI : KIM KICKER UNDER RECORDPLAYER
'   ZZKKA : KIM KICKBACK ACTIVATOR
'   ZZKUL : KIM UPPER LANES
'   ZZKSP : KIM SPINNERS
'   ZZKBS : KIM BONUS SCORE
'   ZZKBU : KIM BUNKER
'   ZZKAN : KIM SPINNING DISCS
'   ZZGLO : KIM GLOWING BALLS
'   ZZKDB : KIM DISCO BALL ANIMATION
'   ZZKSS : KIMS SONG SCROLLER
'
'
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'
'
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [CLOG] Changelog
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

' Build 1.0.0
' Initial Release
'
' Build 1.0.1
' Better check on B2S for systems without B2S Server
' changed embedded controller.vbs to linked controller.VBS
' Added glowing balls for multiball and 5x playfield
' Changed behaviour of the KIDS light, audio was looped
' Added some extra GI at the left and right area
' Fixed issue that the WILDE lights stayed on after collecting Magic Bonus

' Build 1.0.2
' Fixes on MultiBall seqence. It was not an error but lack of timers
' Changed DOF 105 and added DOF 106. This makes a different for autoplunger (strobe) or not. (Thx Arngrim)
' Changed DOF 127 (beacon) to DOFPuls, and not DOFOn (Thx Arngrim)
' Removed 8 bumper back, seems that it is not supported in the future.
' DOF Config added MX Unercab ligting.
' Removed the double DOF's at the bumpers (Thx Arngrim)
' Removed strobe at every ball launch

' Build 1.0.3b
' KickBack activater is now adding, so when earned 2x, it will be activated 2x
' Added DiscoMode :-) using right lane every 120sec
' Added and volume up the mine motor.
' Changed the code for the KIDS bunker
'
' Build 1.0.3.c
' Removed the audio tracks from the .vpx and put them in the music folder. Size of the .vpx was > 500mb
' Added "GARDEN" sequence
'
' Build 1.0.3.d
' Added mechanical tilt (bob)
' Changed code so tilt is actually tilt (disable bumpers etc)
' Changed the 'disabled table' code so everything is disabled
' Changed script so that not more then 1 special mode at the same time is possible (So no multiball in MultiBall)
' Moved 2 pins a little. Ball got stuck under right bumper and above right-right lane.
' Changed the height of the pins from 29 -> 26
' Added Beacon and RGB undercab during disco (Why not :-))
' Changed the size of the Disco Ball and set it some lower so its under the top light.
' Changed some minor code snippets.
' Changed/added some sound samples.
' Added Jacpot lights when multiball is started using Magic Gate.
' Added some extra lighteffects when a mode is ready (set focus on where to Hit)

' Build 3.0 - 3.25
' Playfield change, light changes, colors, sounds, FlexDMD....... to many changes, new rebuild

' Build 3.26
' Fixed issue playsound at line 4760. Thank JoGrady7 for reporting it.

' Build 4
' New playfield graphics
' New lights (VPW Inserts)
'
' New table sounds (VPW)
' Added volume configs in F12 Table Options
' Fixed some sounds not using volume settings.
' Changed ramp lights to white and more lights
' Added side reflections for VPX8 (Render probes)
' Flipper Corrections (VPW)
' Bumpers (VPW)
' Flashers (VPW)
' Table Physics (VPW)

' Changed the voice-overs to an AI generated voice cloning using alltalk_tts and an interview as reference audio.
' Added 5 samples for every voice to get some variation
' Added check for music files existing when music is enabled
' Added Vocal Queue system so voices don't overlap
' Added option to disable Backglass sound Effects
' Added MX led DOF.
' New Backglass with art from Micael Gibs (thanks!)
' Removed UltraDMD, B2S scoring and JP's DMD. Only FlexDMD now.
' Code clean-Up
' Added transparant plastics
' removed 2 bumpers (now 3, not 5)
' Changed GlowBall intensity to 5 (not 10)
' Re-mastered the audio tracks
' Added FlexDMD gifs .\KimWildeDMD directory
' Changed Kickback procedure / grace timer
' Added script option to disable DMD background (EnableDMDBG = 0 at line 350) for real DMD (=faster rendering)
' Added more tracks
' Added option for alternate disco/multiball track (F12)
' Added option to remove siderails (F12)
' Changed scores

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [DODE] DOF DEFINITIONS REMINDER
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'-E101-Flipper Left-Solenoid
'-E102-Flipper Right-Solenoid
'-E103-Slingshot Left-Solenoid
'-E104-Slingshot right-Solenoid
'-E105-Ball Release-Solenoid-Strobe 300
'-E106-Ball Release-Solenoid
'-E107-10 Bumper Middle Left-Bumper-5 Flasher Left
'-E108-10 Bumper Middle Center-Bumper-5 Flasher Center
'-E109-10 Bumper Middle Right-Bumper-5 Flasher Right
'-E110-10 Bumper Back Center-Bumper-5 Flasher Center
'-E111-10 Bumper Back Left-Bumper-5 Flasher Outside Left
'-E112-10 Bumper Back Right-Bumper-5 Flasher Outside Right
'-E113
'-E114-Strobe 250 5
'-E115
'-E116-Strobe 500 10
'-E117-Strobe 1000 20
'-E118-Strobe 2000 40
'-E119
'-E120-5 Flasher Outside Left-Lane Out Left
'-E121-5 Flasher Outside Right-Lane Out Right'
'-E122-5 Flasher Left-Lane Left
'-E123-5 Flasher Right-Lane Right
'-E124-10 Bumper Middle Left-Kickback
'-E125
'-E126-Beacon
'-E127-Beacon M3000 MAX10000-Multiball
'-E128-5 Flasher Center F800-Drain
'-E129-Launch Button
'-E130
'-E131
'-E132
'-E133
'-E134
'-E135
'-E136-Knocker-Strobe 300
'-E137-10 Bumper Middle Left-8 Bumper Left
'-E138-10 Bumper Middle Center-8 Bumper Center
'-E139-10 Bumper Middle Right-8 Bumper Right
'-E140-10 Bumper Back Center-8 Bumper Back
'-E141-10 Bumper Back Left-8 Bumper Left
'-E142-10 Bumper Back Right-8 Bumper Right
'-E143
'-E144
'-E145
'-E146
'-E147-10 Bumper Middle Left-8 Bumper Left
'-E148-10 Bumper Middle Center-8 Bumper Center
'-E149-10 Bumper Middle Right-8 Bumper Right
'-E150-10 Bumper Back Center-8 Bumper Back
'-E151-10 Bumper Back Left-8 Bumper Left
'-E152-10 Bumper Back Right-8 Bumper Right
'-E153
'-E154
'-E155
'-E156
'-E157-10 Bumper Middle Left-8 Bumper Left
'-E158-10 Bumper Middle Center-8 Bumper Center
'-E159-10 Bumper Middle Right-8 Bumper Right
'-E160-10 Bumper Back Center-8 Bumper Back
'-E161-10 Bumper Back Left-8 Bumper Left
'-E162-10 Bumper Back Right-8 Bumper Right
'-E163
'-E164
'-E165
'-E166
'-E167
'-E168
'-E169
'-E170-5 Flasher Left-Purple Spinner
'-E171-5 Flasher Right-Purple Spinner
'-E172
'-E173
'-E174
'-E175
'-E176
'-E177
'-E178
'-E179
'-E180-Flasher Center -> Outer-Blue-Bonus
'-E181-Flasher Left->Right-Yellow
'-E182-Flasher Right->Left-Yellow
'-E183-Red Targets
'-E184-Yellow Targets
'-E185-Green Targets

'
'MX LED EXTRA DOF
'-E200-Ball In PlungerLane (MX)
'-E201-Kickback (MX)
'-E202-Drained: Left Side
'-E203-Drained: Right Side
'-E204-Left OuterLane Rollover / Drain
'-E205-Right OuterLane Rollover / Drain
'-E206-Left InnerLane Rollover
'-E207-Right InnerLane Rollover
'-E208-Tilt
'-E209-Tilt Warning
'-E210-W letter on background
'-E211-I letter on background
'-E212-L letter on background
'-E213-D letter on background
'-E214-E letter on background
'-E215 Disco Letters Flash
'-E216 "MAGIC" letters flash
'-E217 "Jackpot" Flashing
'-E218 Extra Ball Flashing
'-E219 Ball Locked Flashing
'-E220 "SKILL" letters flash
'-E230 Spinner
'-E231 Spinner
'-E232 Positioned where Flasher is in the game
'-E233 Positioned where Flasher is in the game
'-E234 Positioned where Flasher is in the game
'-E235 Positioned where Flasher is in the game
'-E236 Positioned where Flasher is in the game
'-E237 Positioned where Flasher is in the game
'-E238 Stars (during Disco and Super Magic)
'-E239 Stars (during Multiball)
'-E240 Explosion (target Bunker)
'-E241 Comet left to right (Ramp)
'-E242 Comet Right to left (Ramp)
'-E245 Laser Blast Top2Bottom Left
'-E246 Laser Blast Top2Bottom Right

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [SAMP] CUSTOM SAMPLES DEFINITIONS REMINDER
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'SFX_Cloverfield = Drain
'SFX_RagMachine = bumper
'SFX_Meanfall = bumper
'SFX_Polarized = bumper
'SFX_Kiwi = Value Upgrade
'SFX_FlyAway = lane lost
'SFX_Sample1 = lane
'SFX_SpaceRacer = WIlde letter
'SFX_SpaceTime = Exit Bumpers
'SFX_Sample15 = Wilde Letter
'SFX_Nuck = Spell Wilde
'SFX_WaveLength = MultiBall
'SFX_Sample3 = KIM
'SFX_AnotherWorld = MultiBall
'SFX_Sample16 = Ramp Loop Award
'SFX_Uprise3 = Garden
'SFX_Sample9 = Secret
'SFX_TheEnd = End Super Magix
'SFX_Sample14 = Last 10 sec
'SFX_Sample18 = Check GAte Laights
'SFX_Sample11 = Multiplier
'SFX_Stolen = Bonus Score
'SFX_Xarax = Bonus Held
'SFX_70s = Bunker Entered
'SFX_faser6 = KI release
'SFX_Faulty = Closed Biunker


'
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  CODE START
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'
Option Explicit
Randomize

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZOPT] USER SETTINGS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Dim UseFlexDMD : UseFlexDMD = 1             '0 = no FlexDMD, 1 = enable FlexDMD
Dim FlexONPlayfield : FlexONPlayfield= 1          '0 = off, 1=DMD on playfield ( vrroom overrides this )
Dim EnableDMDBG : EnableDMDBG = 1           '0 = off, 1=Enable background image on DMD. Disable it for better performance
Const ballrolleron = 1                  ' set to 0 to turn off the ball roller if you use the "c" key in your cabinet
Dim bShowBallFlash : bShowBallFlash = 0         'Show Ball Flash Animation at plunger

Dim LightLevel : LightLevel = 0.25            ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1               ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8                 ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5         ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5         ' Level of ramp rolling volume. Value between 0 and 1
Dim StagedFlippers : StagedFlippers = 0
Dim bMusicOn : bMusicOn = 0                   ' Background Music. 0 = Disabled, 1 = Enabled
Dim BackgroundMusicVolume : BackGroundMusicVolume = 0.6 ' BackGround Music Volume
Dim bBgSFXSounds : bBgSFXSounds = 0           ' Enable or Disable special SFX BackGlass Sounds
Dim nSFXVolume : nSFXVolume = 0.02            ' Special SFX on Backglass Volume
Dim nVoiceVolume : nVoiceVolume = 0.8         ' Spoken voice volume
Dim PlayWelcomeVoice : PlayWelcomeVoice = 1       ' Plays the welcome voice when tables is loaded
Dim SideRails : SideRails = 1
Dim bpgcurrent : bpgcurrent = 3
Dim v
Dim KimTrackSpecial : KimTrackSpecial = 0       ' Alternate Track for Multiball etc

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
    ColorLUT = Table1.Option("Color Saturation", 1, 11, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 1 Then Table1.ColorGradeImage = ""
  if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100"
  'SideRails
  SideRails = Table1.Option("SideRails", 0, 1, 1, 1, 0, Array("Hidden", "Visible"))
  If SideRails = 1 Then
    ramp15.visible = true
    ramp16.visible = true
  Else
    ramp15.visible = false
    ramp16.visible = false
  End If
  'Number of Balls
  bpgcurrent = Table1.Option("Number of Balls", 3, 5, 1, 3, 0)
    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)
    ' Background Music
    bMusicOn = Table1.Option("Background Music", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))
    BackgroundMusicVolume = Table1.Option("Back Music Volume", 0, 1, 0.01, 0.6, 1)
  'Sound effects Backglass
    bBgSFXSounds = Table1.Option("Enable Sound Effects BackGlass", 0, 1, 1, 1, 0, Array("Disabled", "Enabled"))
    nSFXVolume = Table1.Option("B2S Effects Volume" , 0, 1, 0.01, 0.02, 1)
  'Voice volume
    nVoiceVolume = Table1.Option("Spoken Voice Volume", 0, 1, 0.01, 0.8, 1)
  'Welcome voice
  PlayWelcomeVoice = Table1.Option("Play Welcome Sound", 0, 1, 1, 1, 0, Array("Disabled", "Enabled"))
  'Alternate Track Special
  KimTrackSpecial = Table1.Option("Alternate Soundtrack Special modes", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))
  ' Room brightness

  LightLevel = NightDay/100

  bShowBallFlash = Table1.Option("Show animation at Plunger",0,1,1,1,0, Array( "Disabled", "Enabled"))
  'DMD Stuff
  UseFlexDMD = Table1.Option("Use FlexDMD",0,1,1,1,0, Array( "Disabled", "Enabled"))
  FlexONPlayfield = Table1.Option("FlexDMD on Playfield",0,1,1,1,0, Array( "Disabled", "Enabled"))
' If eventId = 3 or eventid = 0 then
    If FlexONPlayfield = 1 Then
      dotmatrix.visible = true
      TimerDMD.enabled = true
    else
      dotmatrix.visible = false
      TimerDMD.enabled = false
    end if
' end if

  If PlayWelcomeVoice = 1 and eventid = 0 then
    AddSpeechToQueue "KIM-Welcome" , 7000 , 5
  end if

  If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [CORE] Core Options
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Const cGameName = "KIMWILDE"
Const TableName = "Kim Wilde"
Const cAssetsFolder = "Kim Wilde" ' UtraDMD Folder
Const myVersion = "4.00"
Const DebugGeneral = True

Sub LoadCoreFiles
  On Error Resume Next
  ExecuteGlobal GetTextFile("core.vbs")
  If Err Then MsgBox "Can't open core.vbs"
  On Error Goto 0
End Sub
Sub LoadControllerVBS
  On Error Resume Next
  ExecuteGlobal GetTextFile("Controller.vbs")
  If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the Scripts folder of Visual Pinball."
  On Error Goto 0
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZCON] VARIABLES AND CONSTANTS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Const BallSize = 50
Const MaxPlayers = 4
Const BallSaverTime = 15
Const MaxMultiplier = 6
Const MaxMultiballs = 5
'Const bpgcurrent = 3

Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim WildeBonus(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot
Dim SuperJackpot
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue(4)
Dim bAutoPlunger
Dim bInstantInfo
Dim bromconfig
Dim bAttractMode
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim b5xPlayfield
Dim c5xPlayfieldSecondsLeft
Dim bSuperMagic
Dim cSuperMagicSecondsLeft
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim Song
Dim bDiscoMode
Dim bBallFlashDone
Dim dLine(4)
Dim bMusicPlungerStart
Dim EnableKeyPress

EnableKeyPress = False


Dim tablewidth : tablewidth = Table1.width
Dim tableheight : tableheight = Table1.height

LoadCoreFiles
LoadControllerVBS

Dim EnableBallControl : EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZ3DI] DYNAMIC INSERT LIGHTS USING PRIMITIVES
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub LI1_animate: p1.BlendDisableLighting = 200 * (LI1.GetInPlayIntensity / LI1.Intensity): End Sub
Sub LI2_animate: p2.BlendDisableLighting = 200 * (LI2.GetInPlayIntensity / LI2.Intensity): End Sub
Sub LI3_animate: p3.BlendDisableLighting = 200 * (LI3.GetInPlayIntensity / LI3.Intensity): End Sub

Sub LI4_animate: p4.BlendDisableLighting = 200 * (LI4.GetInPlayIntensity / LI4.Intensity): End Sub
Sub LI5_animate: p5.BlendDisableLighting = 200 * (LI5.GetInPlayIntensity / LI5.Intensity): End Sub
Sub LI6_animate: p6.BlendDisableLighting = 200 * (LI6.GetInPlayIntensity / LI6.Intensity): End Sub

Sub LI8_animate: p8.BlendDisableLighting = 200 * (LI8.GetInPlayIntensity / LI8.Intensity): End Sub
Sub LI9_animate: p9.BlendDisableLighting = 200 * (LI9.GetInPlayIntensity / LI9.Intensity): End Sub

Sub LI10_animate: p10.BlendDisableLighting = 200 * (LI10.GetInPlayIntensity / LI10.Intensity): End Sub
Sub LI10B_animate: p10b.BlendDisableLighting = 200 * (LI10B.GetInPlayIntensity / LI10B.Intensity): End Sub
Sub LI11_animate: p11.BlendDisableLighting = 200 * (LI11.GetInPlayIntensity / LI11.Intensity): End Sub

Sub LI12_animate: p12.BlendDisableLighting = 200 * (LI12.GetInPlayIntensity / LI12.Intensity): End Sub
Sub LI13_animate: p13.BlendDisableLighting = 200 * (LI13.GetInPlayIntensity / LI13.Intensity): End Sub
Sub LI14_animate: p14.BlendDisableLighting = 200 * (LI14.GetInPlayIntensity / LI14.Intensity): End Sub
Sub LI15_animate: p15.BlendDisableLighting = 200 * (LI15.GetInPlayIntensity / LI15.Intensity): End Sub

Sub LI16_animate: p16.BlendDisableLighting = 200 * (LI16.GetInPlayIntensity / LI16.Intensity): End Sub
Sub LI17_animate: p17.BlendDisableLighting = 200 * (LI17.GetInPlayIntensity / LI17.Intensity): End Sub
Sub LI18_animate: p18.BlendDisableLighting = 200 * (LI18.GetInPlayIntensity / LI18.Intensity): End Sub
Sub LI19_animate: p19.BlendDisableLighting = 200 * (LI19.GetInPlayIntensity / LI19.Intensity): End Sub
Sub LI20_animate: p20.BlendDisableLighting = 200 * (LI20.GetInPlayIntensity / LI20.Intensity): End Sub
Sub LI21_animate: p21.BlendDisableLighting = 200 * (LI21.GetInPlayIntensity / LI21.Intensity): End Sub

Sub LI23_animate: p23.BlendDisableLighting = 200 * (LI23.GetInPlayIntensity / LI23.Intensity): End Sub
Sub LI24_animate: p24.BlendDisableLighting = 200 * (LI24.GetInPlayIntensity / LI24.Intensity): End Sub
Sub LI25_animate: p25.BlendDisableLighting = 200 * (LI25.GetInPlayIntensity / LI25.Intensity): End Sub

Sub LI26_animate: p26.BlendDisableLighting = 200 * (LI26.GetInPlayIntensity / LI26.Intensity): End Sub

Sub LI27_animate: p27.BlendDisableLighting = 200 * (LI27.GetInPlayIntensity / LI27.Intensity): End Sub
Sub LI28_animate: p28.BlendDisableLighting = 200 * (LI28.GetInPlayIntensity / LI28.Intensity): End Sub
Sub LI29_animate: p29.BlendDisableLighting = 200 * (LI29.GetInPlayIntensity / LI29.Intensity): End Sub

Sub LI30_animate: p30.BlendDisableLighting = 200 * (LI30.GetInPlayIntensity / LI30.Intensity): End Sub
Sub LI31_animate: p31.BlendDisableLighting = 200 * (LI31.GetInPlayIntensity / LI31.Intensity): End Sub
Sub LI32_animate: p32.BlendDisableLighting = 200 * (LI32.GetInPlayIntensity / LI32.Intensity): End Sub
Sub LI33_animate: p33.BlendDisableLighting = 200 * (LI33.GetInPlayIntensity / LI33.Intensity): End Sub
Sub LI34_animate: p34.BlendDisableLighting = 200 * (LI34.GetInPlayIntensity / LI34.Intensity): End Sub
Sub LI35_animate: p35.BlendDisableLighting = 200 * (LI35.GetInPlayIntensity / LI35.Intensity): End Sub

Sub LI36_animate: p36.BlendDisableLighting = 200 * (LI36.GetInPlayIntensity / LI36.Intensity): End Sub
Sub LI37_animate: p37.BlendDisableLighting = 200 * (LI37.GetInPlayIntensity / LI37.Intensity): End Sub
Sub LI38_animate: p38.BlendDisableLighting = 200 * (LI38.GetInPlayIntensity / LI38.Intensity): End Sub

Sub LI39_animate: p39.BlendDisableLighting = 200 * (LI39.GetInPlayIntensity / LI39.Intensity): End Sub
Sub LI40_animate: p40.BlendDisableLighting = 200 * (LI40.GetInPlayIntensity / LI40.Intensity): End Sub
Sub LI41_animate: p41.BlendDisableLighting = 200 * (LI41.GetInPlayIntensity / LI41.Intensity): End Sub

Sub LI43_animate: p43.BlendDisableLighting = 200 * (LI43.GetInPlayIntensity / LI43.Intensity): End Sub
Sub LI44_animate: p44.BlendDisableLighting = 200 * (LI44.GetInPlayIntensity / LI44.Intensity): End Sub
Sub LI45_animate: p45.BlendDisableLighting = 200 * (LI45.GetInPlayIntensity / LI45.Intensity): End Sub
Sub LI46_animate: p46.BlendDisableLighting = 200 * (LI46.GetInPlayIntensity / LI46.Intensity): End Sub

Sub LI47_animate: p47.BlendDisableLighting = 200 * (LI47.GetInPlayIntensity / LI47.Intensity): End Sub
Sub LI48_animate: p48.BlendDisableLighting = 200 * (LI48.GetInPlayIntensity / LI48.Intensity): End Sub
Sub LI49_animate: p49.BlendDisableLighting = 200 * (LI49.GetInPlayIntensity / LI49.Intensity): End Sub

Sub LIScore10_animate: pScore10.BlendDisableLighting = 200 * (LIscore10.GetInPlayIntensity / LIscore10.Intensity): End Sub
Sub LIScore20_animate: pScore20.BlendDisableLighting = 200 * (LIscore20.GetInPlayIntensity / LIscore20.Intensity): End Sub
Sub LIScore30_animate: pScore30.BlendDisableLighting = 200 * (LIscore30.GetInPlayIntensity / LIscore30.Intensity): End Sub

Sub LIscore1_animate: pScore1.BlendDisableLighting = 200 * (LIscore1.GetInPlayIntensity / LIscore1.Intensity): End Sub
Sub LIscore2_animate: pScore2.BlendDisableLighting = 200 * (LIscore2.GetInPlayIntensity / LIscore2.Intensity): End Sub
Sub LIscore3_animate: pScore3.BlendDisableLighting = 200 * (LIscore3.GetInPlayIntensity / LIscore3.Intensity): End Sub
Sub LIscore4_animate: pScore4.BlendDisableLighting = 200 * (LIscore4.GetInPlayIntensity / LIscore4.Intensity): End Sub
Sub LIscore5_animate: pScore5.BlendDisableLighting = 200 * (LIscore5.GetInPlayIntensity / LIscore5.Intensity): End Sub
Sub LIscore6_animate: pScore6.BlendDisableLighting = 200 * (LIscore6.GetInPlayIntensity / LIscore6.Intensity): End Sub
Sub LIscore7_animate: pScore7.BlendDisableLighting = 200 * (LIscore7.GetInPlayIntensity / LIscore7.Intensity): End Sub
Sub LIscore8_animate: pScore8.BlendDisableLighting = 200 * (LIscore8.GetInPlayIntensity / LIscore8.Intensity): End Sub
Sub LIscore9_animate: pScore9.BlendDisableLighting = 200 * (LIscore9.GetInPlayIntensity / LIscore9.Intensity): End Sub

Sub LI2x_animate: p2x.BlendDisableLighting = 200 * (li2X.GetInPlayIntensity / LI2x.Intensity): End Sub
Sub LI3x_animate: p3x.BlendDisableLighting = 200 * (LI3X.GetInPlayIntensity / LI3x.Intensity): End Sub
Sub LI4x_animate: p4x.BlendDisableLighting = 200 * (LI4x.GetInPlayIntensity / LI4x.Intensity): End Sub
Sub LI5x_animate: p5x.BlendDisableLighting = 200 * (LI5X.GetInPlayIntensity / LI5x.Intensity): End Sub
Sub LI6x_animate: p6x.BlendDisableLighting = 200 * (LI6X.GetInPlayIntensity / LI6x.Intensity): End Sub

Sub LightShootAgain_animate: pLightShootAgain.BlendDisableLighting = 200 * (LightShootAgain.GetInPlayIntensity / LightShootAgain.Intensity): End Sub

Sub LIDisco_animate: pDisco.BlendDisableLighting = 200 * (LIDisco.GetInPlayIntensity / LIDisco.Intensity): End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZINT] TABLE INITS AND EXIT
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Sub Table1_Init()
  Call InitFlexDMD
  JPDMD_Init 'Jp's DMD Song Scroller
  LoadEM 'Sub from external script
  Randomize
  Loadhs 'Load High Scores
  bAttractMode = False
  bOnTheFirstBall = False
  bBallInPlungerLane = False
  bBallSaverActive = False
  bBallSaverReady = False
  bMultiBallMode = False
  bGameInPlay = False
  bAutoPlunger = False
  BallsOnPlayfield = 0
  LastSwitchHit = ""
  Tilt = 0
  TiltSensitivity = 6
  Tilted = False
  ' Turn off the bumper lights
  FlBumperFadeTarget(1) = 0
  FlBumperFadeTarget(2) = 0
  FlBumperFadeTarget(4) = 0
  GiOff
  StartAttractMode
  EnableKeyPress = true
End Sub



Sub Table1_Exit()
  If Not FlexDMD is Nothing Then
    FlexDMD.Show = False
    FlexDMD.Run = False
    FlexDMD = NULL
  End If
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   RAMP TRIGGERS TO START RAMP SOUND
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub ramptrigger02_hit()
  WireRampOn False   'Turn on the Ramp Sound
  DOF 241, DofPulse
End Sub


Sub ramptrigger03_hit()
  WireRampOff  'Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger03_unhit()
  RandomSoundRampStop ramptrigger03
End Sub

Sub ramptrigger04_hit()
  WireRampOn True  'Play Plastic Ramp Sound
  DOF 242, DofPulse
End Sub

Sub ramptrigger05_hit()
  WireRampOff  'Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger05_unhit()
  WireRampOn False  'On Wire Ramp, Play Wire Ramp Sound
End Sub

Sub ramptrigger06_hit()
  WireRampOff  'Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger06_unhit()
  RandomSoundRampStop ramptrigger03
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   MUSIC FUNCTIONS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Dim bEndSong : bEndSong = false

Sub Table1_MusicDone()
  if tilted then exit sub
  If bMusicOn = 1 Then
    if bendsong = false then
      Playmusic Song, BackgroundMusicVolume
    else
      bEndSong = False
    end if
  end if
End Sub

Sub PlaySong(name)
  if tilted then exit sub
  name = "KimWilde\" & name
  If debugGeneral Then debug.print "Sub PlaySong name = " & name & " " & Song
  If bMusicOn = 1 Then

    ' exit if file is not found in music folder
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    If not fso.FileExists(GetMusicFolder & "\" & name) then
      dLine(2) = "Audiofiles not found in music folder"
      SongDigitsUpdate()
      exit Sub
    end If

    bEndSong = false
    If (Song <> name) Then
      endmusic
      Song = name
      PlayMusic Song , BackgroundMusicVolume
      TimerLpRotate.enabled = true
    end if
  End If
End Sub

Sub PlayEndSong(name)
  if tilted then exit sub
  name = "KimWilde\" & name
  If debugGeneral Then debug.print "Sub PlaySongEnd name = " & name & " " & Song
  If bMusicOn = 1 Then

    ' exit if file is not found in music folder
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    If not fso.FileExists(GetMusicFolder & "\" & name) then
      dLine(2) = "Audiofiles not found in music folder"
      SongDigitsUpdate()
      exit Sub
    end If

    If (Song <> name) Then
      bEndSong = true
      endmusic
      Song = name
      PlayMusic Song, BackgroundMusicVolume
      TimerLpRotate.enabled = true
    end if
  End If
End Sub

Sub StopSong()
  If debugGeneral Then debug.print "Sub StopSong name = " & Song
    dLine(2) = ""
    SongDigitsUpdate()
    If bMusicOn = 1 Then
      bendsong = false
      endmusic
      song = ""
      TimerLpRotate.enabled = false
    end if
End Sub

Sub TR_Start_Music_Hit()
  if tilted then exit sub
  bBallFlashDone = true
  If bShowBallFlash = 1 then FlasherBall.visible = false
  Dim a
  if BallsOnPlayfield = 1 then
    TimerHitDisco.enabled = 1

    a = int(RndNum(1,6))

    select case a
      case 1 : DOF 900, DOFOn 'White RGB Undercab
      case 2 : DOF 901, DOFOn 'Red Undercab
      case 3 : DOF 902, DOFOn 'Green Undercab
      case 4 : DOF 903, DOFOn 'Blue Undercab
      case 5 : DOF 904, DOFOn 'Yellow Undercab
      case 6 : DOF 905, DOFOn 'Cyan Undercab
    end Select
'   If bMusicPlungerStart = false then
      Call PlayRandomSong
'   Else
'     bMusicPlungerStart = false
'   end if
  end if
End Sub

Sub PlayRandomSong()
  If bMusicOn = 0 then exit sub

  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  Dim sName
  Dim b : b = int(RndNum(1,44))
  sName = "KIMTRACK" & cstr(b) & ".MP3"
  If DebugGeneral then Debug.Print "Sub PlayRandomSong name = " & sName
  ' exit if file is not found in music folder
  If not fso.FileExists(GetMusicFolder & "\KimWilde\" & sname) then
    dLine(2) = "Audiofiles not found in music folder"
    SongDigitsUpdate()
    exit Sub
  end If

  playsong sName
  Select case b
    Case 1 : dLine(2) = "TRACK 1 - LIGHTS DOWN LOW" :SongDigitsUpdate()
    Case 2 : dLine(2) = "TRACK 2 - I WANT WHAT I WANT" :  SongDigitsUpdate()
    Case 3 : dLine(2) = "TRACK 3 - HEY YOU" :SongDigitsUpdate()
    Case 4 : dLine(2) = "TRACK 4 - LOVING YOU MORE" :SongDigitsUpdate()
    Case 5 : dLine(2) = "TRACK 5 - GET OUT" :SongDigitsUpdate()
    Case 6 : dLine(2) = "TRACK 6 - CANT GET ENOUGH OF YOUR LOVE" :SongDigitsUpdate()
    Case 7 : dLine(2) = "TRACK 7 - STONE" :SongDigitsUpdate()
    Case 8 : dLine(2) = "TRACK 8 - CAMBODIA" :SongDigitsUpdate()
    Case 9 : dLine(2) = "TRACK 9 - NEVER TRUST A STRANGER" :SongDigitsUpdate()
    Case 10 :dLine(2) = "TRACK 10 - YOU CAME" :SongDigitsUpdate()
    Case 11 :dLine(2) = "TRACK 11 - KIDS IN AMERICA" :SongDigitsUpdate()
    Case 12 :dLine(2) = "TRACK 12 - Chequered Love" :SongDigitsUpdate()
    Case 13 :dLine(2) = "TRACK 13 - Water On Glass" :SongDigitsUpdate()
    Case 14 :dLine(2) = "TRACK 14 - 26580" :SongDigitsUpdate()
    Case 15 :dLine(2) = "TRACK 15 - Boys" :SongDigitsUpdate()
    Case 16 :dLine(2) = "TRACK 16 - Our Town" :SongDigitsUpdate()
    Case 17 :dLine(2) = "TRACK 17 - Everything we know" :SongDigitsUpdate()
    Case 18 :dLine(2) = "TRACK 18 - You will Never be so wrong" :SongDigitsUpdate()
    Case 19 :dLine(2) = "TRACK 19 - view from a bridge" :SongDigitsUpdate()
    Case 20 :dLine(2) = "TRACK 20 - love blonde" :SongDigitsUpdate()
    Case 21 :dLine(2) = "TRACK 21 - house of salome" :SongDigitsUpdate()
    Case 22 :dLine(2) = "TRACK 22 - dancing in the dark" :SongDigitsUpdate()
    Case 23 :dLine(2) = "TRACK 23 - child come away" :SongDigitsUpdate()
    Case 24 :dLine(2) = "TRACK 24 - take me tonight" :SongDigitsUpdate()
    Case 25 :dLine(2) = "TRACK 25 - stay awhile" :SongDigitsUpdate()
    Case 26 :dLine(2) = "TRACK 26 - you keep me hanging on" :SongDigitsUpdate()
    Case 27 :dLine(2) = "TRACK 27 - the touch" :SongDigitsUpdate()
    Case 28 :dLine(2) = "TRACK 28 - is it over" :SongDigitsUpdate()
    Case 29 :dLine(2) = "TRACK 29 - the second time" :SongDigitsUpdate()
    Case 30 :dLine(2) = "TRACK 30 - ego" :SongDigitsUpdate()
    Case 31 :dLine(2) = "TRACK 31 - words fell down" :SongDigitsUpdate()
    Case 32 :dLine(2) = "TRACK 32 - action city" :SongDigitsUpdate()
    Case 33 :dLine(2) = "TRACK 33 - chaos at the airport" :SongDigitsUpdate()
    Case 34 :dLine(2) = "TRACK 34 - Anyplace, Anywhere, Anytime" :SongDigitsUpdate()
    Case 35 :dLine(2) = "TRACK 35 - To France" :SongDigitsUpdate()
    Case 36 :dLine(2) = "TRACK 36 - Pop dont stop" :SongDigitsUpdate()
    Case 37 :dLine(2) = "TRACK 37 - Kandy Krush" :SongDigitsUpdate()
    Case 38 :dLine(2) = "TRACK 38 - Birthday" :SongDigitsUpdate()
    Case 39 :dLine(2) = "TRACK 39 - Yours Til The End" :SongDigitsUpdate()
    Case 40 :dLine(2) = "TRACK 40 - Cyber Nation War" :SongDigitsUpdate()
    Case 41 :dLine(2) = "TRACK 41 - Trail Of Destruction" :SongDigitsUpdate()
    Case 42 :dLine(2) = "TRACK 42 - Love Is Love" :SongDigitsUpdate()
    Case 43 :dLine(2) = "TRACK 43 - Rocket To The Moon" :SongDigitsUpdate()
    Case 44 :dLine(2) = "TRACK 44 - Stones And Bones" :SongDigitsUpdate()
  end select

End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   START GAME, END GANE
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'
Sub ResetForNewGame()
  Dim i
  bGameInPLay = True
  StopAttractMode
  GiOn
  TotalGamesPlayed = TotalGamesPlayed + 1
  SaveGamesPlayed
  CurrentPlayer = 1
  PlayersPlayingGame = 1
  bOnTheFirstBall = True
  For i = 1 To MaxPlayers
    Score(i) = 0
    BonusPoints(i) = 0
    BonusHeldPoints(i) = 0
    BonusMultiplier(i) = 1
    BallsRemaining(i) = bpgcurrent
    ExtraBallsAwards(i) = 0
  Next
  Tilt = 0
  tilted = 0
  Game_Init()
  vpmtimer.addtimer 1500, "FirstBall '"
End Sub


Sub EndOfGame()
  ShowFlexScene1 "GAME OVER" , "WANT TO PLAY AGAIN?" , 2 , True , False
  vpmtimer.addtimer 2000, "PlayEndSong ""KIMTRACKGAMEOVER2.MP3"" '"
  AddSpeechToQueue "KIM-I am sorry, but your game is over" , 3600 , 5
  AddSpeechToQueue "KIM-Do you want to play with me again" , 6600 , 5
  introposition = 0
  bGameInPLay = False
  GiOff
  EnableKeyPress = false
  vpmtimer.addtimer 3000, "StartAttractMode '"
End Sub

Sub B2SLight(nLightNumber,nPulseTime)
  If B2SOn = true Then
    If nPulseTime = 0 then npulsetime = 1000
    select case nLightNumber
      case 1:Controller.B2SSetData 1, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 1, 0 '"
      case 2:Controller.B2SSetData 2, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 2, 0 '"
      case 3:Controller.B2SSetData 3, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 3, 0 '"
      case 4:Controller.B2SSetData 4, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 4, 0 '"
      case 5:Controller.B2SSetData 5, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 5, 0 '"
      case 6:Controller.B2SSetData 6, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 6, 0 '"
      case 7:Controller.B2SSetData 7, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 7, 0 '"
      case 8:Controller.B2SSetData 8, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 8, 0 '"
      case 9:Controller.B2SSetData 9, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 9, 0 '"
      case 10:Controller.B2SSetData 10, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 10, 0 '"
      case 11:Controller.B2SSetData 11, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 11, 0 '"
      case 12:Controller.B2SSetData 12, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 12, 0 '"
      case 13:Controller.B2SSetData 13, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 13, 0 '"
      case 14:Controller.B2SSetData 14, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 14, 0 '"
      case 15:Controller.B2SSetData 15, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 15, 0 '"
      case 21:Controller.B2SSetData 21, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 21, 0 '"
      case 22:Controller.B2SSetData 22, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 22, 0 '"
      case 23:Controller.B2SSetData 23, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 23, 0 '"
      case 24:Controller.B2SSetData 24, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 24, 0 '"
      case 25:Controller.B2SSetData 25, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 25, 0 '"
      case 26:Controller.B2SSetData 26, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 26, 0 '"
      case 27:Controller.B2SSetData 27, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 27, 0 '"
      case 31:Controller.B2SSetData 31, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 31, 0 '"
      case 32:Controller.B2SSetData 32, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 32, 0 '"
      case 33:Controller.B2SSetData 33, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 33, 0 '"
      case 34:Controller.B2SSetData 34, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 34, 0 '"
      case 35:Controller.B2SSetData 35, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 35, 0 '"
      case 36:Controller.B2SSetData 36, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 36, 0 '"
      case 37:Controller.B2SSetData 37, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 37, 0 '"
      case 38:Controller.B2SSetData 38, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 38, 0 '"
      case 39:Controller.B2SSetData 39, 1:vpmtimer.addtimer nPulseTime, "Controller.B2SSetData 39, 0 '"
    end Select
  end if
end Sub

Sub B2SLightOnOff(nLightNumber,nMode)
  If B2SOn = true Then
    Controller.B2SSetData nLightNumber, nMode
  end if
end sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZSCR] HIGH SCORES
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub Loadhs
  Dim x
  x = LoadValue(TableName, "HighScore1") : If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 0 End If
  x = LoadValue(TableName, "HighScore1Name") : If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "EVI" End If
  x = LoadValue(TableName, "HighScore2") : If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 0 End If
  x = LoadValue(TableName, "HighScore2Name") : If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "DVI" End If
  x = LoadValue(TableName, "HighScore3") : If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 0 End If
  x = LoadValue(TableName, "HighScore3Name") : If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "JVI" End If
  x = LoadValue(TableName, "HighScore4") : If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 0 End If
  x = LoadValue(TableName, "HighScore4Name") : If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "LVI" End If
  x = LoadValue(TableName, "Credits") : If(x <> "") then Credits = CInt(x) Else Credits = 0 End If
  x = LoadValue(TableName, "TotalGamesPlayed") : If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If
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
End Sub

Sub ResetHS()
    SaveValue TableName, "HighScore1", 0
    SaveValue TableName, "HighScore1Name", "EVI"
    SaveValue TableName, "HighScore2", 0
    SaveValue TableName, "HighScore2Name", "EVI"
    SaveValue TableName, "HighScore3", 0
    SaveValue TableName, "HighScore3Name", "EVI"
    SaveValue TableName, "HighScore4", 0
    SaveValue TableName, "HighScore4Name", "EVI"
  Loadhs
  ShowFlexScene2 "HIGH SCORES" , "ERASED FROM SYSTEM" , 4 , 0.5 , 0.5 , true , false
End Sub

Sub SaveGamesPlayed
  SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
  vpmtimer.addtimer 1000, "Loadhs'"
End Sub

Dim hsbModeActive:hsbModeActive = False
Dim hsEnteredName
Dim hsEnteredDigits(3)
Dim hsCurrentDigit
Dim hsValidLetters
Dim hsCurrentLetter
Dim hsLetterFlash

Sub CheckHighscore()
    Dim tmp,a
    tmp = Score(1):a=1

    If Score(2)> tmp Then tmp = Score(2):a=2
    If Score(3)> tmp Then tmp = Score(3):a=3
    If Score(4)> tmp Then tmp = Score(4):a=4

    If tmp> HighScore(1) Then 'add 1 credit for beating the highscore
        AwardSpecial
  End If

    If tmp> HighScore(3) Then
    AddSpeechToQueue "KIM-Congratulations! You got a High Score!" , 3300 , 5
        HighScore(3) = tmp
        'enter player's name
        vpmtimer.addtimer 4000, "HighScoreEntryInit() '"
    Else
        vpmtimer.addtimer 2000, "EndOfBallComplete '"
    End If
End Sub

Sub HighScoreEntryInit()
  hsbModeActive = True
  AddSpeechToQueue "KIM-Enter your initials!" , 2000 , 5

    hsEnteredDigits(0) = "A"
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ<+-0123456789" ' < is used to delete the last letter
    hsCurrentLetter = 1
    HighScoreDisplayName()
End Sub

' flipper moving around the letters

Sub EnterHighScoreKey(keycode)
  If keycode = LeftFlipperKey Then
    'Playsound "sfx_Faulty" , 1 , nSFXVolume
    PlayBgEffect "sfx_Faulty"
      DOF 182, DOFPulse 'right to left
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayName()
  End If

  If keycode = RightFlipperKey Then
    'Playsound "sfx_Faulty" , 1 , nSFXVolume
    PlayBgEffect "sfx_Faulty"
      DOF 181, DOFPulse 'left to right
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayName()
  End If

    If keycode = StartGameKey OR keycode = PlungerKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<") then
            PlayBgEffect "SFX_Commit"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) then
                HighScoreCommitName()
            else
                HighScoreDisplayName()
            end if
        else
            PlayBgEffect "SFX_Commit"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit> 0) then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayName()
        end if
    end if
End Sub

Dim hsletter : hsletter = 1
dim hsdigit:hsdigit = 1


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
    ShowFlexScene1 "ENTER YOUR NAME", Mid(TempStr, 2, 5), 100 , False , False
End Sub

  ' post the high score letters
Sub HighScoreCommitName()
  hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
  HighScoreName(3) = hsEnteredName
    SortHighscore
  savehs
  EndOfBallComplete()
  hsbModeActive = False
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


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZATR] ATTRACT MODE
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub StartAttractMode()
  bAttractMode = True
  EnableKeyPress = true
  if b2son then
    controller.B2SSetGameOver 1
  end if
  UltraDMDTimer.Enabled = 1
  StartLightSeq
  DMDintroloop
  StartRainbow2 GI
  Call StartIntro
End Sub

Sub StopAttractMode()
  bAttractMode = False
  if b2son then
    controller.B2SSetGameOver 0
  end if
  ShowFlexScene0
  LightSeqAttract.StopPlay
  LightSeqAttract2.StopPlay
  StopRainbow2 GI
  Call StopIntro
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZRAI] LIGHTING / RAINBOW LIGHTS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'colors
Dim red : red = 10
Dim orange : orange = 9
Dim amber : amber = 8
Dim yellow : yellow = 7
Dim darkgreen : darkgreen = 6
Dim green : green = 5
Dim blue : blue = 4
Dim darkblue : darkblue = 3
Dim purple : purple = 2
Dim white : white = 1
Dim base : base = 11

Sub SetLightColor(n, col, stat)
  Select Case col
    Case red : n.color = RGB(18, 0, 0) :  n.colorfull = RGB(255, 0, 0)
    Case orange : n.color = RGB(18, 3, 0) : n.colorfull = RGB(255, 64, 0)
    Case amber : n.color = RGB(193, 49, 0) : n.colorfull = RGB(255, 153, 0)
    Case yellow : n.color = RGB(18, 18, 0) : n.colorfull = RGB(255, 255, 0)
    Case darkgreen : n.color = RGB(0, 8, 0) : n.colorfull = RGB(0, 64, 0)
    Case green : n.color = RGB(0, 18, 0) : n.colorfull = RGB(0, 255, 0)
    Case blue : n.color = RGB(0, 18, 18) : n.colorfull = RGB(0, 255, 255)
    Case darkblue : n.color = RGB(0, 8, 8) : n.colorfull = RGB(0, 64, 64)
    Case purple : n.color = RGB(128, 0, 128) : n.colorfull = RGB(255, 0, 255)
    Case white : n.color = RGB(255, 252, 224) : n.colorfull = RGB(193, 91, 0)
    Case base : n.color = RGB(255, 197, 143) : n.colorfull = RGB(255, 255, 236)
  End Select
  If stat <> -1 Then : n.State = 0 : n.State = stat : End If
End Sub

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

Sub StopRainbow(n)
  Dim obj
  RainbowTimer.Enabled = 0
  RainbowTimer.Enabled = 0
    For each obj in RainbowLights
      SetLightColor obj, white, 0
    Next
End Sub

Sub StopRainbow2(n)
  Dim obj
  RainbowTimer1.Enabled = 0
    For each obj in RainbowLights2
      SetLightColor obj, white, 0
      obj.state = 1
      obj.Intensity = 15
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


Sub StartLightSeq()
  Dim a
  For each a in alights
    a.State = 1
  Next

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
  For each a in GI
    a.State = 1
    a.Intensity = 80
  Next
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 50, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqCircleOutOn, 15, 2
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 25, 1
  LightSeqAttract2.UpdateInterval = 10
  LightSeqAttract2.Play SeqCircleOutOn, 15, 3
  LightSeqAttract2.UpdateInterval = 5
  LightSeqAttract2.Play SeqRightOn, 50, 1
  LightSeqAttract2.UpdateInterval = 5
  LightSeqAttract2.Play SeqLeftOn, 50, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqRightOn, 50, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 50, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqRightOn, 40, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 40, 1
  LightSeqAttract2.UpdateInterval = 10
  LightSeqAttract2.Play SeqRightOn, 30, 1
  LightSeqAttract2.UpdateInterval = 10
  LightSeqAttract2.Play SeqLeftOn, 30, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqRightOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqRightOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 15, 1
  LightSeqAttract2.UpdateInterval = 10
  LightSeqAttract2.Play SeqCircleOutOn, 15, 3
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqRightOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 25, 1
  LightSeqAttract2.UpdateInterval = 5
  LightSeqAttract2.Play SeqStripe1VertOn, 50, 2
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqCircleOutOn, 15, 2
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqStripe1VertOn, 50, 3
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqRightOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqCircleOutOn, 15, 2
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqStripe2VertOn, 50, 3
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqRightOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 25, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqStripe1VertOn, 25, 3
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqStripe2VertOn, 25, 3
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqUpOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqDownOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqRightOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqRightOn, 15, 1
  LightSeqAttract2.UpdateInterval = 8
  LightSeqAttract2.Play SeqLeftOn, 15, 1
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

Dim OldGiState
OldGiState = -1   'start witht the Gi off

Sub ChangeGi(col) 'changes the gi color
  Dim bulb
  For each bulb in GI
    SetLightColor bulb, col, -1
  Next
End Sub

Sub GiOn
  DOF 900, DOFOn 'White RGB Undercab
  Dim bulb
  For each bulb in GI
    SetLightColor bulb, base, -1
    bulb.State = 1
  Next
  GIOverhead.state = 1
End Sub

Sub GiOff
  DOF 900, DOFOff 'RGB Undercab
  DOF 901, DOFOff 'RGB Undercab
  DOF 902, DOFOff 'RGB Undercab
  DOF 903, DOFOff 'RGB Undercab
  DOF 904, DOFOff 'RGB Undercab
  DOF 905, DOFOff 'RGB Undercab

  Dim bulb
  For each bulb in GI
    bulb.State = 0
  Next
  GIOverhead.state = 0
End Sub

' GI & light sequence effects

Sub GiEffect(n)
  Select Case n
    Case 0 'all off
      LightSeqGi.Play SeqAlloff
    Case 1 'all blink
      LightSeqGi.UpdateInterval = 4
      LightSeqGi.Play SeqBlinking, , 5, 100
    Case Else
      if DebugGeneral = true then debug.print "Sub GiEffect has unknown case number " & cstr(n)
  End Select
End Sub

Const LE_AllOff = 0
Const LE_AllBlink = 1
Const LE_Random = 2
Const LE_SeqUp = 3
Const LE_LeftRight = 4
Const LE_RampLights = 6
Const LE_CircleOut = 8
Const LE_CircleIn = 9
Const LE_LaneLights = 13
Const LE_WildeLights = 14
Const LE_Score = 15
Const LE_RandomShort = 16
Const LE_RandomLong = 17

Sub LightEffect(n)
  Select Case n
    Case LE_AllOff ' all off
      LightSeqInserts.Play SeqAlloff
    Case LE_AllBlink 'all blink
      LightSeqInserts.UpdateInterval = 4
      LightSeqInserts.Play SeqBlinking, , 5, 100
    Case LE_Random 'random
      LightSeqInserts.UpdateInterval = 10
      LightSeqInserts.Play SeqRandom, 5, , 1000
    Case LE_SeqUp 'upon
      LightSeqInserts.UpdateInterval = 4
      LightSeqInserts.Play SeqUpOn, 10, 1
    Case LE_LeftRight ' left-right
      if bSuperMagic = false then
        LightSeqInserts.UpdateInterval = 5
        LightSeqInserts.Play SeqLeftOn, 10, 1
        LightSeqInserts.UpdateInterval = 5
        LightSeqInserts.Play SeqRightOn, 10, 1
      end if
    Case LE_RampLights
      LightSeqRampLights.UpdateInterval = 0.5
      LightSeqRampLights.play SeqClockRightOn, 90,2
    Case LE_CircleOut
      'Out to in circle
      LightSeqInserts.UpdateInterval = 4
      LightSeqInserts.Play SeqCircleOutOn,10
      LightSeqInserts.Play SeqCircleOutOn,10
      LightSeqInserts.Play SeqCircleOutOn,10
    Case LE_CircleIn
      'Out to in circle
      LightSeqInserts.UpdateInterval = 4
      LightSeqInserts.Play SeqCircleInOn,10
      LightSeqInserts.Play SeqCircleInOn,10
      LightSeqInserts.Play SeqCircleInOn,10
    Case LE_LaneLights ' Lane Lights
      LightSeqLaneLights.UpdateInterval = 10
      LightSeqLaneLights.Play SeqBlinking, , 5, 25
    Case LE_WildeLights ' Wilde Lights
      LightSeqWildeLights.UpdateInterval = 10
      LightSeqWildeLights.Play SeqBlinking, , 5, 100
      LightSeqLaneLights.Play SeqRandom, 5, , 1000
    Case LE_Score ' Bonus Score
      LightSeqScore.UpdateInterval = 0.1
      LightSeqScore.play SeqClockRightOff, 10 , 1
      LightSeqScore.play SeqClockRightOn, 10 , 1
    Case LE_RandomShort
      LightSeqInserts.UpdateInterval = 10
      LightSeqInserts.Play SeqRandom, 5, , 100
    Case LE_RandomLong 'random SuperMagic
      LightSeqInserts.UpdateInterval = 10
      LightSeqInserts.Play SeqRandom, 5, , 10000
  End Select
End Sub

' Flasher Effects using lights

Dim FEStep, FEffect
FEStep = 0
FEffect = 0

Const FE_AllBlink = 1
Const FE_Random = 2

Sub FlashEffect(n)
  dim a
  Select case n
'   Case 0 ' all off
'     LightSeqFlasher.Play SeqAlloff
    Case FE_AllBlink 'all blink
      StartFlash1Sequence
      StartFlash2Sequence
      StartFlash3Sequence
    Case FE_Random
      for a = 1 to 5
        select case int(RndNum(1,3))
          case 1
            vpmtimer.addtimer a*500, "StartFlash1Sequence'"
          case 2
            vpmtimer.addtimer a*500, "StartFlash2Sequence'"
          case 3
            vpmtimer.addtimer a*500, "StartFlash3Sequence'"
        end Select
      next
  End Select
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZSCO] SCORING FUNCTIONS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub AddScore(points)
  If Tilted Then exit sub

  ' add the points to the current players score variable
  if b5xplayfield = true then
    Score(CurrentPlayer) = Score(CurrentPlayer) + (points*5)
  Else
    If bSuperMagic = true and Points > 200 Then
      PlayBgEffect "SFX_Cloverfield"
      Score(CurrentPlayer) = Score(CurrentPlayer) + 10000
    Else
      if bDiscoMode = false then
        Score(CurrentPlayer) = Score(CurrentPlayer) + (points)
      Else
        Score(CurrentPlayer) = Score(CurrentPlayer) + (points*2)
      end if
    end If
  end If
End Sub

Sub AwardExtraBall()
  If Tilted Then exit sub

  If NOT bExtraBallWonThisBall Then
    ShowFlexScene2 "EXTRA" , "BALL" , 2 , 0.2 , 0.2 , True , False
    AddSpeechToQueue "KIM-Congratulations! You got an Extra Ball!" , 4100 , 5
    LightEffect LE_CircleIn
    FlashEffect FE_AllBlink
    DOF 127, DOFPulse 'Beacon
    DOF 218, DOFPulse 'MX Extra Ball
    LightShootAgain.State = 1
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 10000
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    bExtraBallWonThisBall = True
    AddScore (10000*ExtraBallsAwards(CurrentPlayer))
  Else
    AddScore (10000*ExtraBallsAwards(CurrentPlayer))
  END If
End Sub

Sub AwardSpecial()
  DOF 127, DOFPulse
  Credits = Credits + 1
  PlaySound SoundFXDOF("fx_knocker",136,DOFPulse,DOFKnocker)
  GiEffect 1
  LightEffect LE_AllBlink
End Sub


Sub AwardSkillshot()
  If Tilted Then exit sub

  AddSpeechToQueue "KIM-Skillshot!" , 1200 , 5
  ShowFlexScene2 "SKILLSHOT" , "SKILLSHOT" , 2 , 0.2 , 0.1 , True , False
  DOF 116, DOFPulse 'Strobe 5x
  DOF 127, DOFPulse 'Beacon
  DOF 220, DOFPulse 'MX SKILLSHOT
  ResetSkillShotTimer_Timer
  AddScore SkillshotValue(CurrentPLayer)
  GiEffect 1
  LightEffect LE_Random
  LightEffect LE_CircleIn
  FlashEffect FE_AllBlink
  SkillShotValue(CurrentPLayer) = SkillShotValue(CurrentPLayer) + 10000
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZBFU] BALL FUNCTIONS & DRAINS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

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
  ResetForNewPlayerBall()
  CreateNewBall()
End Sub

Sub ResetForNewPlayerBall()

  If PlayersPlayingGame > 1 Then
    If CurrentPlayer = 1 Then
      AddSpeechToQueue "KIM-Player One!" , 1200 , 5
    Elseif currentplayer = 2 Then
      AddSpeechToQueue "KIM-Player Two!" , 1200 , 5
    Elseif currentplayer = 3 Then
      AddSpeechToQueue "KIM-Player Three!" , 1200 , 5
    Elseif currentplayer = 4 Then
      AddSpeechToQueue "KIM-Player Four!" , 1200 , 5
    End If
  End If

  ShowFlexScene2 "PLAYER "+ cstr(CurrentPlayer) , "LAUNCH THE BALL" , 2 , 0.5 , 0.5 , True , True
  AddScore 0
  BonusPoints(CurrentPlayer) = 0
  bExtraBallWonThisBall = False
  ResetNewBallLights()
  ResetNewBallVariables
  bBallSaverReady = True
  bSkillShotReady = True
  UpdateSkillShot
End Sub


Sub CreateNewBall()
  BallRelease.CreateSizedball BallSize / 2
  BallsOnPlayfield = BallsOnPlayfield + 1
  AddSpeechToQueue "sfx_kim_sample2" , 1500 , 1
  RandomSoundBallRelease BallRelease
  DOF 106, DOFPulse
  BallRelease.Kick 90, 4
  If BallsOnPlayfield > 1 Then
    bMultiBallMode = True
      DOF 127, DOFPulse    'Beacon ON
    bAutoPlunger = True
  End If
End Sub


Sub AddMultiball(nballs)
  mBalls2Eject = mBalls2Eject + nballs
  CreateMultiballTimer.Enabled = True
End Sub

Sub CreateMultiballTimer_Timer()
  If bBallInPlungerLane Then
    Exit Sub
  Else
    If BallsOnPlayfield <MaxMultiballs Then
      CreateNewBall()
      mBalls2Eject = mBalls2Eject -1
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
  if b5xPlayfield = true Then
    Timer5xPlayfield.enabled = 0
    Timer5xPlayfieldStop.enabled = 0
    b5xPlayfield = false
  end If
  bMultiBallMode = False
  bOnTheFirstBall = False
  If NOT Tilted Then
    collectbonusscore
    'Timer started after collecting score
    'vpmtimer.addtimer 3500, "EndOfBall2 '"
  Else
    vpmtimer.addtimer 500, "EndOfBall2 '"
  End If
End Sub


Sub EndOfBall2()
  Tilted = False
  If B2SOn then
    Controller.B2SSetTilt 0
  end if
  StopSong
  AddSpeechToQueue "KIM-Get Ready" , 1100 , 5
  Tilt = 0
  DisableTable False
  tilttableclear.enabled = False
  tilttime = 0
  If(ExtraBallsAwards(CurrentPlayer) <> 0) Then
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1
    If(ExtraBallsAwards(CurrentPlayer) = 0) Then
      LightShootAgain.State = 0
    End If
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 2000
    CreateNewBall()
    ResetForNewPlayerBall
    ShowFlexScene2 "EXTRA BALL", "SHOOT AGAIN" , 2 , 0.5 , 0.2 , True , False
    AddSpeechToQueue "KIM-The Same Player shoots Again!" , 2100 , 5
    DOF 218, DOFPulse 'MX Extra Ball
  Else
    BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1
    If(BallsRemaining(CurrentPlayer) <= 0) Then
      CheckHighScore()
    Else
      EndOfBallComplete()
    End If
  End If
End Sub

Sub EndOfBallComplete()
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

    EndOfGame()
  Else
    CurrentPlayer = NextPlayer
    AddScore 0
    ResetForNewPlayerBall()
    CreateNewBall()
  End If
End Sub

Sub Balldrained
  DOF 128, DOFPulse 'Flasher Left Red
  PlayBgEffect "SFX_powerdown"
  TurnOffPlayfieldLights
  GiOff
End Sub

Sub Drain_Hit()
  Drain.DestroyBall
  BallsOnPlayfield = BallsOnPlayfield - 1
  RandomSoundDrain Drain
  PlayBgEffect "SFX_Cloverfield"
  If Tilted Then
    StopEndOfBallMode
  End If
  If(bGameInPLay = True) AND(Tilted = False) Then
    If(bBallSaverActive = True) Then
      AddMultiball 1
      bAutoPlunger = True
      If bMultiBallMode = False Then
        'Ballsaved
        ShowFlexScene1 "BALL" , "SAVED" , 2 , True , False
        AddSpeechToQueue "KIM-Your ball is saved, you will get it back!" , 3600 , 5
      End If
    Else
      If(BallsOnPlayfield = 1) Then
        If(bMultiBallMode = True) then
          ' Stop Jackpot Mode
          StopJackpot
          call PlayRandomSong
          bMultiBallMode = False
          DOF 127, DOFOff   'DOF - Beacon - OFF
          DOF 239, DOFOff 'Stars OFF
          ChangeGi white
          ChangeBall(0)
        End If
        If(bDiscoMode = True) Then
          disco false
        end If

        bMultiBallMode = False
        ChangeGi white
        ChangeBall(0)
    End If
      If(BallsOnPlayfield = 0) Then
        ShowFlexScene2 "BALL" , "LOST" , 2 , 0.1 , 0.2 , True , False
        AddSpeechToQueue "KIM-You lost your ball!" , 2100 , 5
        bMultiBallMode = False
        ChangeGi white
        ChangeBall(0)
        TimerHitDisco.enabled = 0
        LI48.state = 0
        StopSong
        if bSuperMagic = true Then
          StopSuperMagic
        end If
        If(bDiscoMode = True) Then
          disco false
        end If
        vpmtimer.addtimer 1000, "Balldrained '"
        vpmtimer.addtimer 3000, "EndOfBall '"
        StopEndOfBallMode
      End If
    End If
  End If
End Sub

Sub TriggerDrainLeft
  DOF 202, DOFPulse 'MX Effect
end Sub
Sub TriggerDrainRight
  DOF 203, DOFPulse 'MX Effect
End sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZBSL] BALL SAVE & LAUNCH
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'

Sub ballsavestarttrigger_hit
  If(bBallSaverReady = True) AND(15 <> 0) And(bBallSaverActive = False) Then
  EnableBallSaver 15
  End If
End Sub

Sub swPlungerRest_Hit()
  bBallFlashDone = false
  If bShowBallFlash = 1 then FlasherBall.visible = true
  DOF 200, DOFOn 'MX Effect
  PlaySoundAtVol "fx_sensor", ActiveBall, VolumeDial
  bBallInPlungerLane = True
  If bAutoPlunger Then
    DOF 106, DOFPulse 'Right Sling
    swPlungerRest.TimerEnabled = 1
    bAutoPlunger = False
  End If
  DOF 141, DOFOn 'Launch Button
  vpmtimer.addtimer 2000, "swPlungerRestPlaySound'"
  If bSkillShotReady Then
    swPlungerRest.TimerEnabled = 0
  End If
  LastSwitchHit = "swPlungerRest"
End Sub

Sub swPlungerRestPlaySound()
  If bBallInPlungerLane = true then
    AddSpeechToQueue "KIM-Shoot the ball!" , 2000 , 5
'   if bMusicPlungerStart = false Then
'     bMusicPlungerStart = True
'     PlayRandomSong
'   end if
  end if
End sub

Sub swPlungerRest_UnHit()
  DOF 200, DOFOff 'MX Effect
  DOF 141, DOFOff 'Launch Button
  bBallInPlungerLane = False
  swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
  If bSkillShotReady Then
    ResetSkillShotTimer.Enabled = 1
  End If
End Sub

Sub swPlungerRest_Timer
  swPlungerRest.TimerEnabled = 0
  plunger.autoplunger = True
  plunger.fire
  DOF 105, DOFPulse 'Right Sling strobe
  plunger.autoplunger = false
End Sub

Sub EnableBallSaver(seconds)
  bBallSaverActive = True
  bBallSaverReady = False
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
  BallSaverTimerExpired.Enabled = False
  ' clear the flag
  Dim waittime
  waittime = 4000
  vpmtimer.addtimer waittime, "ballsavegrace'"
  ' if you have a ball saver light then turn it off at this point
  If bExtraBallWonThisBall = True Then
    LightShootAgain.State = 1
  Else
    LightShootAgain.State = 0
  End If
End Sub

Sub ballsavegrace
  bBallSaverActive = False
End Sub

Sub BallSaverSpeedUpTimer_Timer()
  'debug.print "Ballsaver Speed Up Light"
  BallSaverSpeedUpTimer.Enabled = False
  ' Speed up the blinking
  LightShootAgain.BlinkInterval = 80
  LightShootAgain.State = 2
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZINT] DMD INTRO ATTRACT MODE
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Dim introposition , introtime
'introposition = 0
'introtime = 0

Sub StartIntro()
  introtime = 0
  introposition = 1
  intromover.enabled = True
End Sub

Sub StopIntro
  introposition = 0
  introtime = 0
  intromover.enabled = False
  Dof 238,DofOff
  ShowFlexScene0
End Sub

Sub DMDintroloop
  introtime = 0 'Reset Intro Time
  introposition = introposition + 1 'Move one position
  Debug.print "IntroPosition = " & cstr(introposition)
  Select Case introposition
    Case 1
      ShowFlexScene4 "** SUPERED **" , "PRESENTS" , 1 , 2 , False , False
      DOF 238,DofOff 'Stars
    Case 2 : ShowFlexScene4 "** KIM WILDE **" , "PINBALL MACHINE" , 1 , 2, False , False : AddSpeechToQueue "sfx_kim_sample1" , 1500 , 1
    Case 3 : ShowFlexScene1 "ALL TIME" , "HIGH SCORES" , 3 , False , False
    Case 4 : ShowFlexScene7 "1 "+ HighScoreName(0), cstr(HighScore(0)) , 0.5 , 2 ,  False , False
    Case 5 : ShowFlexScene7 "2 "+ HighScoreName(1), cstr(HighScore(1)) , 0.5 , 2 ,  False , False
    Case 6 : ShowFlexScene7 "3 "+ HighScoreName(2), cstr(HighScore(2)) , 0.5 , 2 ,  False , False
    Case 7 : ShowFlexScene7 "4 "+ HighScoreName(3), cstr(HighScore(3)) , 0.5 , 2 ,  False , False
    Case 8 : ShowFlexScene1 "GAMES PLAYED" , cstr(TotalGamesPlayed) , 3 , False , False
    Case 9 : ShowFlexScene1 "BUILD VERSION" , "4.00 RC3" , 3 , False , False
    Case 10: ShowFlexScene1 "44 audio tracks" , "AI Generated Voices" , 3 , False , False
    Case 11: ShowFlexScene7 "Artwork by" , "Michael Gibs" , 0.5 , 2 , False , False
    Case 12: ShowFlexScene1 "Hit the F7 key" , "to erase High Scores" , 3 , False , False
    Case 13
      ShowFlexScene2 "Lets go back . . ." , ". . . To the 80s" , 7 , 0.5 , 0.5 , False , False
      DOF 238,DofOn 'Stars
      introposition = 0
  End Select
End Sub

Sub intromover_timer
  'Ticks every 1 second
  introtime = introtime + 1
  Select Case introposition
    Case 0 : If IntroTime = 4 then DMDintroloop
    Case 1 : If IntroTime = 4 then DMDintroloop
    Case 2 : If IntroTime = 4 then DMDintroloop
    Case 3 : If IntroTime = 4 then DMDintroloop
    Case 4 : If IntroTime = 4 then DMDintroloop
    Case 5 : If IntroTime = 4 then DMDintroloop
    Case 6 : If IntroTime = 4 then DMDintroloop
    Case 7 : If IntroTime = 4 then DMDintroloop
    Case 8 : If IntroTime = 4 then DMDintroloop
    Case 9 : If IntroTime = 4 then DMDintroloop
    Case 10 : If IntroTime = 4 then DMDintroloop
    Case 11 : If IntroTime = 4 then DMDintroloop
    Case 12 : If IntroTime = 4 then DMDintroloop
    Case 13 : If IntroTime = 8 then DMDintroloop
  End Select
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZVAR] TABLE VARIABLES
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

' Game Variables
Dim LaneBonus
Dim TargetBonus
Dim RampBonus
Dim OrbitBonus
Dim spinvalue
Dim finalflips
Dim Saves
Dim Drains
Dim inmode

' Player Variables
Dim PlayerBall1Locked(4)
Dim PlayerBall2Locked(4)
Dim PlayerMultiBallGateOpen(4)
Dim PlayerWildeWLight(4)
Dim PlayerWildeILight(4)
Dim PlayerWildeLLight(4)
Dim PlayerWildeDLight(4)
Dim PlayerWildeELight(4)
Dim PlayerMagicReady(4)
Dim PlayerKimMagicA(4)
Dim PlayerKimMagicB(4)
Dim PlayerKimMagicC(4)
Dim PlayerKimMagicD(4)
Dim PlayerKimMagicE(4)

Dim PlayerKimMultiKLight(4)
Dim PlayerKimMultiILight(4)
Dim PlayerKimMultiMLight(4)

Dim PlayerKimKLight(4)
Dim PlayerKimILight(4)
Dim PlayerKimMLight(4)
Dim PlayerKickBackActivated(4)
Dim PlayerPlayField5x(4)
Dim PlayerBonusMultiplierHeld(4)
Dim PlayerBonusMultiplier(4)
Dim PlayerLI2xLight(4)
Dim PlayerLI3xLight(4)
Dim PlayerLI4xLight(4)
Dim PlayerLI5xLight(4)
Dim PlayerLI6xLight(4)
Dim PlayerKidsKLight(4)
Dim PlayerKidsILight(4)
DIm PlayerKidsDLight(4)
Dim PlayerKidsSLight(4)
Dim PlayerBumperValue(4)
Dim PlayerBumperMaxCounter(4)
Dim PlayerBumper(4)
Dim PlayerGardenCounter(4)
Dim PlayerGardenMaxCounter(4)
Dim PlayerDiscoReady(4)


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZGAM] GAME STARTING & RESETS & SKILLSHOT
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'

Sub Game_Init() 'called at the start of a new game
  Dim i
  For i = 0 to 4
    SkillshotValue(i) = 10000 ' increases by 10000 each time it is collected
    PlayerBall1Locked(i) = 0
    PlayerBall2Locked(i) = 0
    PlayerMultiBallGateOpen(i) = 0
    PlayerWildeWLight(i) = 0
    PlayerWildeILight(i) = 0
    PlayerWildeLLight(i) = 0
    PlayerWildeDLight(i) = 0
    PlayerWildeELight(i) = 0
    PlayerMagicReady(i) = 0
    PlayerKimMagicA(i)=0
    PlayerKimMagicB(i)=0
    PlayerKimMagicC(i)=0
    PlayerKimMagicD(i)=0
    PlayerKimMagicE(i)=0
    PlayerKimKLight(i)=0
    PlayerKimILight(i)=0
    PlayerKimMLight(i)=0
    PlayerKimMultiKLight(i)=0
    PlayerKimMultiILight(i)=0
    PlayerKimMultiMLight(i)=0
    PlayerKickBackActivated(i)=0
    PlayerPlayField5x(i) = 0
    PlayerBonusMultiplierHeld(i) = 0
    PlayerBonusMultiplier(i) = 1
    PlayerLI2xLight(i)= 0
    PlayerLI3xLight(i)= 0
    PlayerLI4xLight(i)= 0
    PlayerLI5xLight(i)= 0
    PlayerLI6xLight(i)= 0
    PlayerKidsKLight(i) = 0
    PlayerKidsILight(i) = 0
    PlayerKidsDLight(i) = 0
    PlayerKidsSLight(i) = 0
    PlayerBumperValue(i) = 10
    PlayerBumperMaxCounter(i) = 100
    PlayerBumper(i) = 0
    PlayerGardenCounter(i) = 0
    PlayerGardenMaxCounter(i) = 20
  Next

  MultiballLockOffTimer.enabled = 1
  sw1.isdropped = 0
  sw2.isdropped = 0
  sw3.isdropped = 0

  kickbackgate.open = false

  lrflashtime.Enabled = False

  bExtraBallWonThisBall = False

End Sub

Sub StopEndOfBallMode()              'this sub is called after the last ball is drained
  ResetSkillShotTimer_Timer
End Sub

Sub ResetNewBallVariables()
  PlayerBumper(CurrentPlayer) = 0
End Sub

Sub TurnOffPlayfieldLights()
  Dim a
  For each a in alights
    a.State = 0
  Next
  B2SLightOnOff 4,0
  B2SLightOnOff 5,0
  B2SLightOnOff 6,0
  B2SLightOnOff 7,0
  B2SLightOnOff 8,0

  TimerFlasherSun5.enabled = false
End Sub

Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released
  TurnOffPlayfieldLights()
  ChangeBall(0)
  LI4.State = PlayerBall1Locked(CurrentPlayer)
  LI5.State = PlayerBall2Locked(CurrentPlayer)

  'Init Wilde Lights
  If PlayerWildeWLight(Currentplayer) = 1 Then
    LI16.state = 1
    B2SLightOnOff 4,1
  end if
  If PlayerWildeILight(Currentplayer) = 1 Then
    LI17.state = 1
    B2SLightOnOff 5,1
  end if
  If PlayerWildeLLight(Currentplayer) = 1 Then
    LI18.state = 1
    B2SLightOnOff 6,1
  end if
  If PlayerWildeDLight(Currentplayer) = 1 Then
    LI19.state = 1
    B2SLightOnOff 7,1
  end if
  If PlayerWildeELight(Currentplayer) = 1 Then
    LI20.state = 1
    B2SLightOnOff 8,1
  end if
  If PlayerMagicReady(CurrentPlayer) = 1 Then
    LI21.state = 2
  end if

  ' Init Kim MultiBall Lights
  li1.state = PlayerKimMultiKLight(CurrentPlayer)
  LI2.state = PlayerKimMultiILight(CurrentPlayer)
  LI3.state = PlayerKimMultiMLight(CurrentPlayer)

  If PlayerMultiBallGateOpen(CurrentPlayer) = 1 then
    MultiballLockOnTimer.enabled = 1
    TimerFlasherSun5.enabled = true
  Else
    MultiballLockOffTimer.enabled = 1
    TimerFlasherSun5.enabled = false
  end If

  ' Init Kim light (Right)
  If PlayerKimKLight(CurrentPlayer) = 1 Then
    LI23.State = 1
  end if
  If PlayerKimILight(CurrentPlayer) = 1 Then
    LI24.State = 1
  end if
  If PlayerKimMLight(CurrentPlayer) = 1 Then
    LI25.State = 1
  end if

  'Init kickbackgate
  If PlayerKickBackActivated(CurrentPlayer) > 0 Then
    kickbackgate.open = 1
    li26.state = 1
  Else
    kickbackgate.open = 0
    li26.state = 0
  end if
  KI06.Timerenabled = False

  'Init Multiplier
  LI6x.state = 0
  li5x.state = 0
  li4x.state = 0
  li3x.state = 0
  li2x.state = 0

  If PlayerBonusMultiplierHeld(CurrentPlayer) = 1 Then
    select case PlayerBonusMultiplier(currentplayer)
      case 2
        li2x.state = 1
      case 3
        li3x.state = 1
        li2x.state = 1
      case 4
        li4x.state = 1
        li3x.state = 1
        li2x.state = 1
      case 5
        li5x.state = 1
        li4x.state = 1
        li3x.state = 1
        li2x.state = 1
      case 6
        LI6x.state = 1
        li5x.state = 1
        li4x.state = 1
        li3x.state = 1
        li2x.state = 1
    end Select
    If PlayerBonusMultiplier(CurrentPlayer) > 6 Then
        LI6x.state = 2
        li5x.state = 1
        li4x.state = 1
        li3x.state = 1
        li2x.state = 1
    end if
  end if

  ' KIDS BONUS Reset
  'KIDS Light
  LI35.State = PlayerKidsKLight(CurrentPlayer)
  LI36.State = PlayerKidsILight(CurrentPlayer)
  LI37.State = PlayerKidsDLight(CurrentPlayer)
  LI38.State = PlayerKidsSLight(CurrentPlayer)

  If LI35.state = 1 and LI36.State = 1 and LI37.State = 1 and LI38.state = 1 then
    If T04.Isdropped = 0 then
      T04.IsDropped = 1
      playsoundatvol "fx_mine_motor" , T04, MechSoundLevel
    end if
    'LI22.state = 2
    TimerFlasherSun6.enabled = true
    LI47.state = 1

  Else
    If T04.Isdropped = 1 then
      T04.IsDropped = 0
      playsoundatvol "fx_mine_motor" , T04, MechSoundLevel
    end if
    'LI22.state = 0
    TimerFlasherSun6.enabled = false
    LI47.state = 0
  end if

  'Bumper Value Lights
  Select case PlayerBumperValue(CurrentPlayer)
    case 10
      LI43.state = 1
    case 20
      LI43.state = 1
      LI44.state = 1
    case 40
      LI43.state = 1
      LI44.state = 1
      LI45.state = 1
    case 80
      LI43.state = 1
      LI44.state = 1
      LI45.state = 1
      LI46.state = 1
    case Else
      LI43.state = 1
      LI44.state = 1
      LI45.state = 1
      LI46.state = 2
  end select
  ' Magic Lights
  LI30.state = PlayerKimMagicA(CurrentPlayer)
  LI31.state = PlayerKimMagicB(CurrentPlayer)
  LI32.state = PlayerKimMagicC(CurrentPlayer)
  LI33.state = PlayerKimMagicD(CurrentPlayer)
  LI34.state = PlayerKimMagicE(CurrentPlayer)

  SelectRandomSecret

  ' Super Magic Reset
  if bSuperMagic = true then
    StopSuperMagic
  end if

  ' Disco Mode Reset
  if bDiscoMode = true Then
    disco False
  end if

  ' Disable ready For Disco
  LI48.state = 0
  LI49.state = 0
  PlayerGardenCounter(CurrentPlayer) = 0

End Sub

Sub UpdateSkillShot() 'Updates the skillshot light
  LI39.state = 0
  LI40.state = 0
  LI41.state = 0

  dim a
  a = int(RndNum(1,3))
  select case a
    case 1
      LI39.state = 2
    case 2
      LI40.state = 2
    case 3
      LI41.state = 2
  end select
End Sub

Sub SkillshotOff_Hit 'trigger to stop the skillshot due to a weak plunger shot
  If bSkillShotReady Then
    ResetSkillShotTimer_Timer
  End If
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
  ResetSkillShotTimer.Enabled = 0
  bSkillShotReady = False
    LI39.state = 0
    LI40.state = 0
    LI41.state = 0
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZLAN] LANE SHIFTING
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub RotateLaneLightsLeft
  Dim TempState
  Dim TempState2
  TempState = LI12.State
  LI12.State = LI13.State
  LI13.State = LI14.State
  LI14.State = LI15.State
  LI15.State = TempState

  TempState2 = LI27.State
  LI27.State = LI28.State
  LI28.State = LI29.State
  LI29.state = TempState2

End Sub

Sub RotateLaneLightsRight
  Dim TempState
  Dim TempState2
  TempState = LI15.State
  LI15.State = LI14.State
  LI14.State = LI13.State
  LI13.State = LI12.State
  LI12.State = TempState

  TempState2 = LI29.State
  LI29.State = LI28.State
  LI28.State = LI27.State
  LI27.state = TempState2
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZBUM] BUMPERS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'

Sub Bumper1_Hit
  If NOT Tilted Then
    PlayBgEffect "SFX_Predator1"
    RandomSoundBumperMiddle Bumper1
    FlBumperFadeTarget(1) = 1
    Bumper1.timerenabled = True
    DOF 107, DOFPulse
    B2SLight 3, 100
    B2SLight RndInt(11,15) , 100
    AddScore PlayerBumperValue(CurrentPlayer)
    BumperCounterChange
  End If
End Sub


Sub Bumper1_Timer
  FlBumperFadeTarget(1) = 0
End Sub

Sub Bumper2_Hit
  If NOT Tilted Then
    PlayBgEffect "SFX_Predator2"
    RandomSoundBumperTop Bumper2
    FlBumperFadeTarget(2) = 1
    Bumper2.timerenabled = True
    DOF 110, DOFPulse
    B2SLight 3, 100
    B2SLight RndInt(11,15) , 100
    AddScore PlayerBumperValue(CurrentPlayer)
    BumperCounterChange
  End If
End Sub

Sub Bumper2_Timer
  FlBumperFadeTarget(2) = 0
End Sub

Sub Bumper4_Hit
  If NOT Tilted Then
    PlayBgEffect "SFX_Meanfall"
    RandomSoundBumperMiddle Bumper4
    FlBumperFadeTarget(4) = 1
    Bumper4.timerenabled = True
    DOF 108, DOFPulse
    B2SLight 3, 100
    B2SLight RndInt(11,15) , 100
    AddScore PlayerBumperValue(CurrentPlayer)
    BumperCounterChange
  End If
End Sub

Sub Bumper4_Timer
  FlBumperFadeTarget(4) = 0
End Sub

Sub TopSlingSHot1_SlingShot
  If NOT Tilted Then
    ObjLevel(2) =1
    FlasherFlash2_Timer
    PlayBgEffect "SFX_RagMachine"
    PlaySoundAt SoundFXDOF("fx_bumper", 107, DOFPulse, DOFContactors), ActiveBall
  End If
End Sub

Sub TopSlingSHot3_Slingshot
  If NOT Tilted Then
    ObjLevel(3) =1:FlasherFlash3_Timer
    PlayBgEffect "SFX_RagMachine"
    PlaySoundAt SoundFXDOF("fx_bumper", 108, DOFPulse, DOFContactors), ActiveBall
  End If
End Sub

Sub TopSlingSHot4_SlingShot
  If NOT Tilted Then
    ObjLevel(1) =1:FlasherFlash1_Timer
    ObjLevel(3) =1:FlasherFlash3_Timer
    PlayBgEffect "SFX_RagMachine"
    PlaySoundAt SoundFXDOF("fx_bumper", 109, DOFPulse, DOFContactors), ActiveBall
  End If
End Sub

Sub rsband004_Hit
  ObjLevel(1) =1:FlasherFlash1_Timer
End Sub

Sub BumperCounterChange()
  if tilted then exit Sub

  PlayerBumper(CurrentPlayer) = PlayerBumper(CurrentPlayer) + 1
  DMDDisplayTopTextOnScore cstr(PlayerBumperMaxCounter(CurrentPlayer)-PlayerBumper(CurrentPlayer))+" BUMPER HITS TO GO!" , 1
  If PlayerBumper(CurrentPlayer) > PlayerBumperMaxCounter(CurrentPlayer) Then
    'Counter limit reached, make it more dificult but multiply Score
    LightEffect LE_SeqUp

    PlayerBumper(CurrentPlayer) = 0
    PlayerBumperMaxCounter(CurrentPlayer) = PlayerBumperMaxCounter(CurrentPlayer) + 100
    PlayerBumperValue(CurrentPlayer) = PlayerBumperValue(CurrentPlayer) * 2
    ShowFlexScene1 "BUMPER VALUE" , "UPGRADE" , 2 , True , False
    PlayBgEffect "SFX_Kiwi"
    Select case PlayerBumperValue(CurrentPlayer)
      case 10
        LI43.state = 1
      case 20
        LI43.state = 1
        LI44.state = 1
      case 40
        LI44.state = 1
        LI45.state = 1
      case 80
        LI44.state = 1
        LI45.state = 1
        LI46.state = 1
      case Else
        LI44.state = 1
        LI45.state = 1
        LI46.state = 2
    end select
  end if
end sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZFLA] FLIPPER LANES
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'
Sub LeftInlane_Hit
  If Tilted Then Exit Sub
  leftInlaneSpeedLimit
  PlayBgEffect "SFX_Sample1"
  DOF 122, DOFPulse ' Left Flasher Cyan
  DOF 206, DOFPulse ' MX effect
  LaneBonus = LaneBonus + 1
  If LI13.state = 0 then LightEffect LE_LaneLights
  LI13.State = 1
  CheckAllLaneLights
  ' Do some sound or light effect
  LastSwitchHit = "lane"
End Sub

Sub LeftInlane1_Hit
  If Tilted Then Exit Sub
  PlayBgEffect "SFX_FlyAway"
  AddSpeechToQueue "KIM-Ooh Noo!" , 1700 , 5
  DOF 120, DOFPulse
  DOF 204, DOFPulse 'MX Effect
  LaneBonus = LaneBonus + 1
  If LI12.state = 0 then LightEffect LE_LaneLights
  LI12.State = 1
  CheckAllLaneLights
  ' Do some sound or light effect
  LastSwitchHit = "outerlane"
End Sub

Sub RightInlane_Hit
  If Tilted Then Exit Sub
  rightInlaneSpeedLimit
  PlayBgEffect "SFX_Sample1"
  DOF 123, DOFPulse
  DOF 207, DOFPulse 'MX Effect
  LaneBonus = LaneBonus + 1
  If LI14.state = 0 then LightEffect LE_LaneLights
  LI14.State = 1
  CheckAllLaneLights
  ' Do some sound or light effect
  LastSwitchHit = "lane"
End Sub

Sub RightInlane1_Hit
  If Tilted Then Exit Sub
  PlayBgEffect "SFX_FlyAway"
  AddSpeechToQueue "KIM-Ooh Noo!" , 1700 , 5
  DOF 121, DOFPulse
  DOF 205, DOFPulse 'MX Effect
  LaneBonus = LaneBonus + 1
  If LI15.state = 0 then LightEffect LE_LaneLights
  LI15.State = 1
  CheckAllLaneLights
  ' Do some sound or light effect
  LastSwitchHit = "outerlane"
End Sub

' Inlane switch speedlimit code

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


Sub CheckAllLaneLights()
  AddScore 500
  DOF 184, DOFPulse
  If LI12.State = 1 and LI13.State = 1 and LI14.State = 1 and LI15.State = 1 Then
  ' All Lanes are lit
    LI12.State = 0
    LI13.state = 0
    LI14.State = 0
    LI15.State = 0

    LightEffect LE_WildeLights
    FlashEffect FE_AllBlink

    ShowFlexScene1 "BONUS COLLECTED" , "WILDE LETTER" , 1 , True , False
    PlayBgEffect "SFX_SpaceRacer"

    AddBonusScore
  ' Find a "WILDE" lamp to lit
    If LI16.State = 0 Then
      LI16.State = 1
      B2SLightOnOff 4,1
      PlayerWildeWLight(currentplayer) = 1
    Else
      If LI17.State = 0 Then
        LI17.State = 1
        B2SLightOnOff 5,1
        PlayerWildeILight(currentplayer) = 1
      Else
        If LI18.State = 0 Then
          LI18.state = 1
          B2SLightOnOff 6,1
          PlayerWildeLLight(currentplayer) = 1
        Else
          If LI19.State = 0 Then
            LI19.State = 1
            B2SLightOnOff 7,1
            PlayerWildeLLight(currentplayer) = 1
          Else
            If LI20.State = 0 Then
              LI20.state = 1
              B2SLightOnOff 8,1
              PlayerWildeELight(currentplayer) = 1
            end If
          end If
        end If
      end If
    end If
    CheckWildeLights
  end If
end sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  VISUAL EFFECTS ONLY
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'

Sub TriggerMiddle_Hit()
  if tilted then exit sub
  LightEffect LE_CircleOut
  DOF 245, DofPulse 'Laser Left
  DOF 246, DofPulse 'Laer Right
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKWL] KIM WILDE LAMPS & START MAGIC
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub TargetLeftW_Hit()
  if tilted then exit sub
  DD_Spell_Wilde
  If LI16.state <> 1 and bSuperMagic = false Then
    LI16.state = 1
    B2SLightOnOff 4,1
    DOF 210, DOFPulse
    PlayerWildeWLight(currentplayer) = 1
    AddScore 700
    CheckWildeLights
  Else
    AddScore 200
  end if
End Sub

Sub TargetLeftI_Hit()
  if tilted then exit sub
  DD_Spell_Wilde
  If LI17.state <> 1 and bSuperMagic = false Then
    LI17.state = 1
    B2SLightOnOff 5,1
    DOF 211, DOFPulse
    PlayerWildeILight(currentplayer) = 1
    AddScore 700
    CheckWildeLights
  Else
    AddScore 200
  end if
End Sub

Sub TargetLeftL_Hit()
  if tilted then exit sub
  DD_Spell_Wilde
  If LI18.state <> 1 and bSuperMagic = false Then
    LI18.state = 1
    B2SLightOnOff 6,1
    DOF 212, DOFPulse
    PlayerWildeLLight(currentplayer) = 1
    AddScore 700
    CheckWildeLights
  Else
    AddScore 200
  end if
End Sub

Sub TargetLeftD_Hit()
  if tilted then exit sub
  DD_Spell_Wilde
  If LI19.state <> 1 and bSuperMagic = false Then
    LI19.state = 1
    B2SLightOnOff 7,1
    DOF 213, DOFPulse
    PlayerWildeDLight(currentplayer) = 1
    AddScore 700
    CheckWildeLights
  Else
    AddScore 200
  end if
End Sub

Sub TargetLeftE_Hit()
  if tilted then exit sub
  DD_Spell_Wilde
  If LI20.state <> 1 and bSuperMagic = false Then
    LI20.state = 1
    B2SLightOnOff 8,1
    DOF 214, DOFPulse
    PlayerWildeELight(currentplayer) = 1
    AddScore 700
    CheckWildeLights
  Else
    AddScore 200
  end if
End Sub

Sub CheckWildeLights()
  If LI16.state = 1 and LI17.state = 1 and LI18.state = 1 and LI19.state = 1 and LI20.state = 1 Then
    'All lights are lit, start Magic ready
    DOF 215, DOFPulse
    AddSpeechToQueue "KIM-The magic gate is now open!" , 3600 , 5
    LI21.state = 2
    PlayerMagicReady(currentplayer) = 1
    ShowFlexScene2 "MAGIC GATE" , "IS NOW OPEN" , 2 , 0.2 , 0.2 , True , False
    PlayBgEffect "id_fx_MultiHit"
    lightseqmagic.play SeqCircleInOn ,10, 1
  Else
    PlayBgEffect "SFX_Sample15"
    LightEffect LE_WildeLights
  end if
End Sub

Sub DD_Spell_Wilde()
    ShowFlexScene2 "SPELL WILDE TO" , "UNLOCK THE MAGIC" , 2 , 0.2 , 0.2 , True , False
    PlayBgEffect "SFX_Nuck"
    dof 184, DOFPulse
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKMB] : KIM MULTIBALL
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub EnableMultiballLockSequense()
  If LI1.state = 1 and LI2.state = 1 and LI3.state = 1 and bDiscoMode = false and bMultiBallMode = false and bSuperMagic = false Then
    AddBonusScore
    MultiballLockOnTimer.enabled = 1
    KI01.enabled = 1
    playsoundatvol "fx_mine_motor", MBTarget0 , MechSoundLevel
    ' Save State
    PlayerMultiBallGateOpen(CurrentPlayer) = 1
    LightEffect LE_AllBlink

    IF LI4.state = 0 and LI5.state = 0 and LI6.state = 0 Then
      ' this is the first lock sequence ball1
      DOF 116,DOFPulse
      LI4.state = 2
      PlayerBall1Locked(CurrentPlayer) = 2
      ShowFlexScene2 "MULTIBALL" , "LOCK ENABLED" , 2 , 0.5 , 0.1 , True , False
      PlayBgEffect "SFX_WaveLength"
      AddSpeechToQueue "KIM-The Multi Ball lock is now enabled!" , 2700 , 5
      TimerFlasherSun5.enabled = true
      AddBonusScore
      AddScore 800
    end If
    If LI4.state = 1 and LI5.state = 0 and LI6.state = 0 Then
      ' this is the second lock sequence ball2
      DOF 116,DOFPulse
      LI5.state = 2
      PlayerBall2Locked(CurrentPlayer) = 2
      ShowFlexScene2 "MULTIBALL" , "LOCK ENABLED" , 2 , 0.5 , 0.1 , True , False
      PlayBgEffect "SFX_WaveLength"
      AddSpeechToQueue "KIM-The Multi Ball lock is now enabled!" , 2700 , 5
      TimerFlasherSun5.enabled = true
      AddBonusScore
      AddScore 800
    end If
    If LI4.state = 1 and LI5.state = 1 and LI6.state = 0 Then
      ' this is the third lock sequence ball3
      DOF 116,DOFPulse
      LI6.state = 2
      ShowFlexScene2 "MULTIBALL" , "LOCK ENABLED" , 2 , 0.5 , 0.1 , True , False
      PlayBgEffect "SFX_WaveLength"
      AddSpeechToQueue "KIM-The Multi Ball lock is now enabled!" , 2700 , 5
      TimerFlasherSun5.enabled = true
      AddBonusScore
      AddScore 800
    end If
  Else
    PlayerMultiBallGateOpen(CurrentPlayer) = 0
  end if
end sub

Sub TimerFlasherSun5_Timer()
  ObjSunlevel(5) = 2
  FlasherSun5_Timer
  vpmtimer.addtimer 350, "ObjSunlevel(5) = 2 '"
  vpmtimer.addtimer 360, "FlasherSun5_Timer '"
end sub

Sub MultiballLockOnTimer_Timer()
  dim CurrentZ
  CurrentZ = mbtarget1.z
  if currentZ > -60 Then
    MBTarget0.z = currentz - 1
    mbtarget1.z = currentz - 1
    mbtarget2.z = currentz - 1
    mbtarget3.z = currentz - 1
  Else
    MultiballLockOnTimer.enabled = 0
    ki01.enabled = 1
  end if

  if currentz < -40 Then
    sw1.isdropped = 1
    sw2.isdropped = 1
    sw3.isdropped = 1
    li1.state = 0
    li2.state = 0
    li3.state = 0
    PlayerKimMultiKLight(CurrentPlayer) = 0
    PlayerKimMultiILight(CurrentPlayer) = 0
    PlayerKimMultiMLight(CurrentPlayer) = 0
  end if
End Sub

Sub MultiballLockOffTimer_Timer()
  MultiballLockOnTimer.enabled = 0
  dim CurrentZ
  CurrentZ = mbtarget1.z
  if currentZ < -10 Then
    MBTarget0.z = currentz + 1
    mbtarget1.z = currentz + 1
    mbtarget2.z = currentz + 1
    mbtarget3.z = currentz + 1
  Else
    MultiballLockOffTimer.enabled = 0
    ki01.enabled = 0
  end if

  if currentz > -20 Then
    sw1.isdropped = 0
    sw2.isdropped = 0
    sw3.isdropped = 0
  end if
End Sub

Sub sw1_Hit()
  if tilted then exit sub
  ShowFlexScene2 "HIT KIM TARGETS TO" , "UNLOCK MULTIBALL" , 2 , 0.2 , 0.2 , True , False
  PlayBgEffect "SFX_Sample3"
  addscore 200
  DOF 183, DOFPulse
  if bMultiBallMode = false then
    LI1.state = 1
    PlayerKimMultiKLight(CurrentPlayer) = 1
    EnableMultiballLockSequense
  end if
End Sub

Sub sw2_Hit()
  if tilted then exit sub
  ShowFlexScene2 "HIT KIM TARGETS TO" , "UNLOCK MULTIBALL" , 2 , 0.2 , 0.2 , True , False
  PlayBgEffect "SFX_Sample3"
  addscore 200
  DOF 183, DOFPulse
  if bMultiBallMode = false then
    LI2.state = 1
    PlayerKimMultiILight(CurrentPlayer) = 1
    EnableMultiballLockSequense
  end if
End Sub

Sub sw3_Hit()
  if tilted then exit sub
  ShowFlexScene2 "HIT KIM TARGETS TO" , "UNLOCK MULTIBALL" , 2 , 0.2 , 0.2 , True , False
  PlayBgEffect "SFX_Sample3"
  addscore 200
  DOF 183, DOFPulse
  if bMultiBallMode = false then
    LI3.state = 1
    PlayerKimMultiMLight(CurrentPlayer) = 1
    EnableMultiballLockSequense
  end if
End Sub

Sub Ki01_Hit()
  if tilted then
    ki01.DestroyBall
    KI02.CreateSizedball BallSize / 2
    KI02.Kick 90, 30
    exit sub
  end if
  ' Random
  LightEffect LE_Random
  AddScore 10000
  ki01.DestroyBall
  KI01.enabled = 0
  TimerFlasherSun5.enabled = false
  ' Save State
  PlayerMultiBallGateOpen(CurrentPlayer) = 0
  DOF 117,DOFPulse

  If li4.state = 2 Then
    ' This is the first ball locked
    ShowFlexScene7 "BALL 1 IS" , "LOCKED" , 0.2 , 2 , True , False
    PlayBgEffect "SFX_AnotherWorld"
    AddSpeechToQueue "KIM-Ball one is locked!" , 1600 , 5
    DOF 219, DOFPulse 'MX BALL LOCKED
    StartFlash3Sequence()
    ' Save the Ball Locked State for current player
    PlayerBall1Locked(CurrentPlayer) = 1
    li4.state = 1
    MultiballLockOffTimer.enabled = 1
    playsoundatvol "fx_mine_motor" , MBTarget0, MechSOundLevel
    bBallSaverReady = True
    vpmtimer.addtimer 2000, "MultiBallBallRelease'"
  end if

  If li5.state = 2 Then
    ' This is the second ball locked
    ShowFlexScene7 "BALL 2 IS" , "LOCKED" , 0.2 , 2 , True , False
    PlayBgEffect "SFX_AnotherWorld"
    AddSpeechToQueue "KIM-Ball Two is locked!" , 1600 , 5
    DOF 219, DOFPulse 'MX BALL LOCKED
    StartFlash3Sequence()
    ' Save the Ball Locked State for current player
    PlayerBall2Locked(CurrentPlayer) = 1
    li5.state = 1
    MultiballLockOffTimer.enabled = 1
    playsoundatvol "fx_mine_motor", MBTarget0 , MechSoundLevel
    bBallSaverReady = True
    vpmtimer.addtimer 2000, "MultiBallBallRelease'"
  end if

  If li6.state = 2 and bDiscoMode = false and bSuperMagic = false and bMultiBallMode = false  Then
    ' This is the third ball locked
    ' Start Multiball sequence
    lightseqMultiball.UpdateInterval = 10
    lightseqmultiball.play SeqAllOff
    ShowFlexScene7 "BALL 3 IS" , "LOCKED" , 0.2 , 2 , True , False
    PlayBgEffect "SFX_AnotherWorld"
    AddSpeechToQueue "KIM-Ball Three is locked!" , 1600 , 5
    DOF 219, DOFPulse 'MX BALL LOCKED
    changegi red
    ChangeBall(1)
    StartFlash3Sequence()
      DOF 119,DOFPulse

    li6.state = 0
    li5.state = 0
    li4.state = 0

    ' Save the Ball Locked State for current player
    PlayerBall1Locked(CurrentPlayer) = 0
    PlayerBall2Locked(CurrentPlayer) = 0
    PlayerMultiBallGateOpen(CurrentPlayer) = 0

    MultiballLockOffTimer.enabled = 1
    playsoundatvol "fx_mine_motor" , MBTarget0, mechsoundlevel

    KI02.TimerEnabled = 1
    StopSong
  end if

End Sub

Sub MultiBallBallRelease()
  BallRelease.CreateSizedball BallSize / 2
  DOF 106, DOFPulse
  SoundSaucerKick 1, BallRelease
  BallRelease.Kick 90, 4
End sub

Sub KI02_Timer()
  KI02.TimerEnabled = 0
  DOF 127,DOFPulse '(Beacon 3000ms minimal)
  DOF 239, DOFOn 'Stars
  lightseqmultiball.play SeqCircleInOn ,15, 3
  FlashEffect FE_Random
  ShowFlexScene2 "MULTIBALL" , "MULTIBALL" , 2 , 0.1 , 0.2 , True , False
  AddSpeechToQueue "KIM-Multi Ball!" , 2000 , 5
  vpmtimer.addtimer 2000, "addmultiball 2 '"
  vpmtimer.addtimer 2000, "ReleaseLockedMultiBall '"
  bBallSaverReady = True
  RandomJackpotLight
  StartFlash3Sequence()
End Sub

Sub ReleaseLockedMultiBall()
  If KimTrackSpecial = 0 then
    playsong "KIMTRACKSPECIAL.MP3"
    dLine(2) = "MULTIBALL - CAMBODIA & REPRISE"
  Else
    playsong "KIMTRACKSPECIAL2.MP3"
    dLine(2) = "MULTIBALL - LES NUITS SANS KIM WILDE"
  end if
  SongDigitsUpdate()
  StartFlash3Sequence()
  DOF 159, DOFPulse
  SoundSaucerKick 1 , KI02
  KI02.CreateSizedball BallSize / 2
  KI02.Kick 90, 30
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKJP] : KIM JACKPOT
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Sub RandomJackpotLight()
  AddBonusScore

  Select Case int(RndNum(1,5))
    Case 1 : LI8.State = 2
    Case 2 : LI9.State = 2
    Case 3 : LI10.State = 2
    Case 4 : LI11.State = 2
    Case 5 : LI10B.State = 2
  End Select
End Sub

Sub StopJackpot()
    LI8.State = 0
    LI9.State = 0
    LI10.state = 0
    LI10B.state = 0
    LI11.state = 0
end sub

Sub T8_Hit()
  if tilted then exit sub
  If LI8.state <> 0 Then
    DOF 114,DOFPulse '(Strobe)
    LI8.state = 0
    ShowFlexScene1 "JACKPOT" , "" , 1 , True , False
    AddSpeechToQueue "KIM-JACKPOT!!!" , 1600 , 5
    Addscore 8000
    DOF 183, DOFPulse
    DOF 217, DOFPulse 'MX LED JACKPOT
    FlashEffect FE_Random
    LightEffect LE_SeqUp
    RandomJackpotLight
  end if
End Sub

Sub T9_Hit()
  if tilted then exit sub
  If LI9.state <> 0 Then
    DOF 114,DOFPulse '(Strobe)
    LI9.state = 0
    ShowFlexScene1 "JACKPOT" , "" , 1 , True , False
    AddSpeechToQueue "KIM-JACKPOT!!!" , 1600 , 5
    Addscore 9000
    DOF 183, DOFPulse
    DOF 217, DOFPulse 'MX LED JACKPOT
    FlashEffect FE_Random
    LightEffect LE_SeqUp
    RandomJackpotLight
  else
    LightEffect LE_RampLights
    AddBonusScore
    ShowFlexScene1 "RAMP LOOP", "AWARD" , 1 , True , False
    addscore 1000
  end if
End Sub

Sub T10_Hit()
  if tilted then exit sub
  If LI10.state <> 0 Then
    DOF 114,DOFPulse '(Strobe)
    LI10.state = 0
    ShowFlexScene1 "JACKPOT" , "" , 1 , True , False
    AddSpeechToQueue "KIM-JACKPOT!!!" , 1600 , 5
    Addscore 8000
    FlashEffect FE_Random
    LightEffect LE_SeqUp
    DOF 183, DOFPulse
    DOF 217, DOFPulse 'MX LED JACKPOT
    RandomJackpotLight
  else
    AddBonusScore
    LightEffect LE_LeftRight
  end if
  If LI49.state = 2 Then
    ShowFlexScene2 "SECRET GARDEN" , "6000 POINTS AWARD" , 2 , 0.5 , 0.2 , True , False
    PlayBgEffect "SFX_Uprise3"
    addscore 6000
    LightEffect LE_Random
    LightEffect LE_SeqUp
    LightEffect LE_Random
    LI49.state = 0
    PlayerGardenCounter(CurrentPlayer) = 0
    PlayerGardenMaxCounter(CurrentPlayer) = PlayerGardenMaxCounter(CurrentPlayer) + 20
  else
    GardenCounterChange
  end if
End Sub

Sub GardenCounterChange()
  PlayerGardenCounter(CurrentPlayer) = PlayerGardenCounter(CurrentPlayer) + 1
  ShowFlexScene1 "GARDEN COUNTER", cstr(PlayerGardenMaxCounter(CurrentPlayer)-PlayerGardenCounter(CurrentPlayer)+1)+" HITS LEFT" , 2 , true , False
  If PlayerGardenCounter(CurrentPlayer) > PlayerGardenMaxCounter(CurrentPlayer) Then
    lightseqgarden.play SeqCircleInOn ,10, 1
    LightEffect LE_SeqUp
    LI49.state = 2
  end if
end sub


Sub T10B_Hit()
  if DebugGeneral then debug.print "T10B hit"
  if tilted then exit sub
  If LI10B.state <> 0 Then
    DOF 114,DOFPulse '(Strobe)
    LI10B.state = 0
    ShowFlexScene1 "JACKPOT" , "" , 1 , True , False
    AddSpeechToQueue "KIM-JACKPOT!!!" , 1600 , 5
    Addscore 8000
    FlashEffect FE_Random
    LightEffect LE_SeqUp
    DOF 183, DOFPulse
    DOF 217, DOFPulse 'MX LED JACKPOT
    RandomJackpotLight
  Else
    if DebugGeneral then debug.print "T10B hit"
    DOF 240, DofPulse 'Explosion
    lightseqbunker.play SeqCircleOutOn , 10 , 1
    PlayBgEffect "SFX_Faulty"
  end if
End Sub

Sub T11_Hit()
  if tilted then exit sub
  If LI11.state <> 0 Then
    DOF 114,DOFPulse '(Strobe)
    LI11.state = 0
    ShowFlexScene1 "JACKPOT" , "" , 1 , True , False
    AddSpeechToQueue "KIM-JACKPOT!!!" , 1600 , 5
    Addscore 8000
    FlashEffect FE_Random
    LightEffect LE_SeqUp
    DOF 183, DOFPulse
    DOF 217, DOFPulse 'MX LED JACKPOT
    RandomJackpotLight
  end if
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKMA] : KIM MAGIC
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub KI03_Hit()
  if tilted then KI03.Kick 0,35,0 : exit sub

  If PlayerMagicReady(currentplayer) <> 1 or bSuperMagic = true or bDiscoMode = true or bMultiBallMode = true Then
    DOF 114,DOFPulse '(Strobe)
    AddBonusScore
    addscore 1100
    KI03.TimerInterval = 1000
    KI03.TimerEnabled = 1
    StartFlash2Sequence()
    PlayBgEffect "SFX_faser2"
  else
    DOF 116,DOFPulse '(Strobe)
    AddBonusScore
    AwardSecret
    SelectRandomSecret
  end If
End Sub

Sub AwardSecret()

  ' First dim all Wilde lights and Magic lights
  LI16.state = 0
  LI17.state = 0
  LI18.state = 0
  LI19.state = 0
  LI20.state = 0
  LI21.state = 0

  B2SLightOnOff 4,0
  B2SLightOnOff 5,0
  B2SLightOnOff 6,0
  B2SLightOnOff 7,0
  B2SLightOnOff 8,0

  PlayerWildeWLight(currentplayer) = 0
  PlayerWildeILight(currentplayer) = 0
  PlayerWildeLLight(currentplayer) = 0
  PlayerWildeDLight(currentplayer) = 0
  PlayerWildeELight(currentplayer) = 0
  PlayerMagicReady(currentplayer) = 0

  If LI30.state = 1 and LI31.state = 1 and LI32.state = 1 and LI33.state = 1 and LI34.state = 1 Then
    ' Super Magix Bonus Check Comes Here
    PlayerKimMagicA(CurrentPlayer) = 0
    PlayerKimMagicB(CurrentPlayer) = 0
    PlayerKimMagicC(CurrentPlayer) = 0
    PlayerKimMagicD(CurrentPlayer) = 0
    PlayerKimMagicE(CurrentPlayer) = 0
    StopSong
    AddSpeechToQueue "KIM-You unlocked the SUPER MAGIC mode!" , 2500 , 5
    AddSpeechToQueue "KIM-Every target is now 10000 points!" , 3600 , 5
    AddSpeechToQueue "KIM-Hold on for Multi Ball!" , 2100 , 5
    AddSpeechToQueue "KIM-Let's play some magic!" , 2300 , 5
    LightEffect LE_RandomLong
    FlashEffect FE_AllBlink
    KI03.TimerInterval=10000
    ki03.TimerEnabled = 1
    vpmtimer.addtimer 10000, "StartSuperMagic '"
  else

    If LI30.state = 2 Then
      LI30.state = 1
      PlayerKimMagicA(CurrentPlayer) = LI30.state
      ShowFlexScene2 "MAGIC AWARD" , "EXTRA BALL" , 2 , 0.2 , 0.1 , True , False
      PlayBgEffect "SFX_Sample9"
      LightEffect LE_Random
      LightEffect LE_SeqUp
      LightEffect LE_Random
      LightEffect LE_Random

      AwardExtraBall
    end if

    If LI31.state = 2 Then
      LI31.state = 1
      PlayerKimMagicB(CurrentPlayer) = LI31.state
      ShowFlexScene2 "MAGIC AWARD" , "100000 POINTS" , 2 , 0.2 , 0.1 , True , False
      PlayBgEffect "SFX_Sample9"
      AddSpeechToQueue "KIM-You earned 100000 points!" , 2700 , 5
      LightEffect LE_Random
      LightEffect LE_SeqUp
      LightEffect LE_Random
      LightEffect LE_Random

      addscore 100000
    end if

    If LI32.state = 2 Then
      LI32.state = 1
      PlayerKimMagicC(CurrentPlayer) = LI32.state
      ShowFlexScene2 "MAGIC AWARD" , "MULTIBALL" , 2 , 0.2 , 0.1 , True , False
      PlayBgEffect "SFX_Sample9"
      AddSpeechToQueue "KIM-Multi Ball!" , 1300 , 5
      LightEffect LE_Random
      LightEffect LE_SeqUp
      LightEffect LE_Random
      LightEffect LE_Random

      changegi red
      ChangeBall(1)
      RandomJackpotLight

      If KimTrackSpecial = 0 then
        playsong "KIMTRACKSPECIAL.MP3"
        dLine(2) = "MULTIBALL - CAMBODIA & REPRISE"
      Else
        playsong "KIMTRACKSPECIAL2.MP3"
        dLine(2) = "MULTIBALL - LES NUITS SANS KIM WILDE"
      end if

      SongDigitsUpdate()
      StartFlash3Sequence()
      DOF 119,DOFPulse

      vpmtimer.addtimer 2000, "AddMultiball 2 '"
    end if

    If LI33.state = 2 Then
      LI33.state = 1
      PlayerKimMagicD(CurrentPlayer) = LI33.state
      ShowFlexScene2 "MAGIC AWARD" , "LOCK MULTIPLIER" , 2 , 0.2 , 0.1 , True , False
      PlayBgEffect "SFX_Sample9"
      AddSpeechToQueue "KIM-Your bonus multiplier is locked!" , 2700 , 5
      LightEffect LE_Random
      LightEffect LE_SeqUp
      LightEffect LE_Random
      LightEffect LE_Random

      PlayerBonusMultiplierHeld(CurrentPlayer) = 1
    end if

    If LI34.state = 2 Then
      LI34.state = 1
      PlayerKimMagicE(CurrentPlayer) = LI34.state
      ShowFlexScene2 "MAGIC AWARD" , "5X PLAYFIELD" , 2 , 0.2 , 0.1 , True , False
      PlayBgEffect "SFX_Sample9"
      AddSpeechToQueue "KIM-You have got 5 times the playfield!" , 3000 , 5
      LightEffect LE_Random
      LightEffect LE_SeqUp
      LightEffect LE_Random
      LightEffect LE_Random

      Start5xPlayfield
    end if

    SelectRandomSecret

    KI03.TimerInterval=5000
    ki03.TimerEnabled = 1
  end if
end sub

Sub StartSuperMagic()
  FlashEffect FE_AllBlink
  changeball(2)
  ChangeGi purple
  cSuperMagicSecondsLeft = 60
  LI30.state = 2
  LI31.state = 2
  LI32.state = 2
  LI33.state = 2
  LI34.state = 2
  'LIHead.state = 2
  bSuperMagic = true
  TimerSuperMagicMax.enabled = 1
  TimerSuperMagic.enabled = 1
  If KimTrackSpecial = 0 then
    playsong "KIMTRACKSPECIAL.MP3"
    dLine(2) = "SUPERMAGIC - CAMBODIA & REPRISE"
  Else
    playsong "KIMTRACKSPECIAL2.MP3"
    dLine(2) = "SUPERMAGIC - LES NUITS SANS KIM WILDE"
  end if
  SongDigitsUpdate()
  AddMultiball 3
  DOF 127, DOFPulse
  DOF 118, DOFPULSE
  DOF 238, DofOn

end sub

Sub StopSuperMagic()
  changeball(0)
  TimerSuperMagicMax.enabled = 0
  TimerSuperMagic.Enabled = 0
  cSuperMagicSecondsLeft = 0
  bSuperMagic = false
  LightEffect LE_RandomShort
  LI30.state = 0
  LI31.state = 0
  LI32.state = 0
  LI33.state = 0
  LI34.state = 0
  ChangeGi white
    SelectRandomSecret
  DOF 238, DofOff
end sub

Sub TimerSuperMagicMax_Timer()
  TimerSuperMagicMax.enabled = 0
  TimerSuperMagic.Enabled = 0
  bSuperMagic = false
  LI30.state = 0
  LI31.state = 0
  LI32.state = 0
  LI33.state = 0
  LI34.state = 0
  ChangeGi white
  StopSuperMagic
end Sub

Sub TimerSuperMagic_Timer()
  dof 116, DOFPulse
  cSuperMagicSecondsLeft = cSuperMagicSecondsLeft - 1
  if cSuperMagicSecondsLeft < 1 Then
    bSuperMagic = False
    ShowFlexScene2 "SUPER MAGIC" , "TIME IS UP" , 4 , 0.2 , 0.1 , True , False
    PlayBgEffect "SFX_TheEnd"
    cSuperMagicSecondsLeft = 0
    TimerSuperMagicMax.enabled = 0
    TimerSuperMagic.enabled = 0
    ChangeGi white
    StopSuperMagic
  end if

  If cSuperMagicSecondsLeft < 10 then
    DOF 127, DOFPulse
    PlayBgEffect "SFX_Sample14"
    FlashEffect FE_AllBlink
  end if
  LightEffect LE_RandomShort
end Sub

Sub Start5xPlayfield()
  ChangeBall(3)
  c5xPlayfieldSecondsLeft = 60
  Timer5xPlayfieldStop.enabled = 1
  Timer5xPlayfield.enabled = 1
  b5xplayfield = True
  changegi green
end sub

Sub Timer5xplayfield_timer()
  c5xPlayfieldSecondsLeft = c5xPlayfieldSecondsLeft - 1
  if c5xPlayfieldSecondsLeft < 1 Then
    b5xplayfield = False
    ShowFlexScene2 "5x PLAYFIELD" , "TIME IS UP" , 2 , 0.2 , 0.1 , True , False
    PlayBgEffect "SFX_TheEnd"
    c5xPlayfieldSecondsLeft = 0
    Timer5xPlayfieldStop.enabled = 0
    Timer5xPlayfield.enabled = 0
    ChangeBall(0)
    ChangeGi white
  end if

  If c5xPlayfieldSecondsLeft < 10 then
    PlayBgEffect "SFX_Sample14"
    FlashEffect FE_AllBlink
  end if
end Sub

Sub SelectRandomSecret()
  Dim a
  Dim b
  b = 0
  Dim aFound
  aFound = False
  If LI30.state = 2 or LI31.state = 2 or LI32.state = 2 or LI33.state = 2 or LI34.state = 2 Then
    exit Sub
  end If
  Do
    b = b + 1
    a = int(RndNum(1,5))
    Select case a
      case 1
        if LI30.state = 0 Then
          LI30.state = 2
          PlayerKimMagicA(CurrentPlayer) = 2
          aFound = true
        end If
      case 2
        if LI31.state = 0 Then
          LI31.state = 2
          PlayerKimMagicB(CurrentPlayer) = 2
          aFound = true
        end If
      case 3
        if LI32.state = 0 Then
          LI32.state = 2
          PlayerKimMagicC(CurrentPlayer) = 2
          aFound = true
        end If
      case 4
        if LI33.state = 0 Then
          LI33.state = 2
          PlayerKimMagicD(CurrentPlayer) = 2
          aFound = true
        end If
      case 5
        if LI34.state = 0 Then
          LI34.state = 2
          PlayerKimMagicE(CurrentPlayer) = 2
          aFound = true
        end If
    end Select
  Loop until aFound = true or b > 100
end Sub


Sub KI03_Timer()
  DOF 181, DOFPulse 'left to right
  KI03.TimerEnabled = 0
  KI03.Kick 0,35,0
  DOF 157, DOFPulse
  SoundSaucerKick 1, KI03
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKFS] KIM FLASHER SEQENCES
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub StartFlash1Sequence()
  ObjLevel(1) = 1
  TimerPulseFlasher1.interval = 300
  TimerPulseFlasher1.enabled = 1
  TimerFlash1ForXSeconds.enabled = 1
End Sub

Sub TimerPulseFlasher1_Timer()
  ObjLevel(1) = 1
  Flasherflash1_Timer
  Sound_Flash_Relay 1, Flasherbase1
  me.enabled = false
  TimerPulseFlasher1.interval = TimerPulseFlasher1.interval - 30
  If TimerPulseFlasher1.interval < 30 Then
    me.enabled = False
  Else
    me.enabled = true
  end if
End Sub

Sub TimerFlash1ForXSeconds_Timer()
  TimerPulseFlasher1.enabled = 0
  me.enabled = false
End Sub

Sub StartFlash2Sequence()
  ObjLevel(2) = 1
  TimerPulseFlasher2.interval = 300
  TimerPulseFlasher2.enabled = 1
  TimerFlash2ForXSeconds.enabled = 1
End Sub

Sub TimerPulseFlasher2_Timer()
  ObjLevel(2) = 1
  Flasherflash2_Timer
  Sound_Flash_Relay 1, Flasherbase2
  me.enabled = false
  TimerPulseFlasher2.interval = TimerPulseFlasher2.interval - 30
  If TimerPulseFlasher2.interval < 30 Then
    me.enabled = False
  Else
    me.enabled = true
  end if
End Sub

Sub TimerFlash2ForXSeconds_Timer()
  TimerPulseFlasher2.enabled = 0
  me.enabled = false
End Sub

Sub StartFlash3Sequence()
  ObjLevel(3) = 1
  TimerPulseFlasher3.interval = 300
  TimerPulseFlasher3.enabled = 1
  TimerFlash3ForXSeconds.enabled = 1
End Sub

Sub TimerPulseFlasher3_Timer()
  ObjLevel(3) = 1
  Flasherflash3_Timer
  Sound_Flash_Relay 1, Flasherbase3
  me.enabled = false
  TimerPulseFlasher3.interval = TimerPulseFlasher3.interval - 30
  If TimerPulseFlasher3.interval < 30 Then
    me.enabled = False
  Else
    me.enabled = true
  end if
End Sub

Sub TimerFlash3ForXSeconds_Timer()
  TimerPulseFlasher3.enabled = 0
  me.enabled = false
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKKI] KIM KICKER UNDER RECORDPLAYER
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub KI04_Hit()
  If tilted then KI04.Kick 0,60,0 : PlaySoundAt SoundFXDOF("fx_Kicker_Release", 159, DOFPulse, DOFContactors), KI04 : exit sub
  PlayBgEffect "SFX_faser2"
  StartFlash1Sequence
  If LI48.state <> 2 then
    KI04.timerenabled = 1
  Else
    if bDiscoMode = false and bMultiBallMode = false and bSuperMagic = false then
      disco True
    Else
      AddBonusScore
      KI04.timerenabled = 1
    end if
  end if
End Sub

Sub KI04_Timer()
  KI04.TimerEnabled = 0
  KI04releaseball
End Sub

Sub KI04ReleaseBall()
  KI04.Kick 0,60,0
  DOF 182, DOFPulse 'right to left
  DOF 159, DOFPulse
  SoundSaucerKick 1, KI04
end sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKKA] KIM KICKBACK ACTIVATOR
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub T01_Hit()
  if tilted then exit sub
  addscore 300
  DOF 184, DOFPulse
  If LI23.state <> 1 Then
    LI23.state = 1
    PlayerKimKLight(currentplayer)=1
    CheckGateOpenLights
  end If
End Sub
Sub T02_Hit()
  if tilted then exit sub
    addscore 300
  DOF 184, DOFPulse
  If LI24.state <> 1 Then
    LI24.state = 1
    PlayerKimILight(currentplayer)=1
    CheckGateOpenLights
  end If
End Sub
Sub T03_Hit()
  if tilted then exit sub
    addscore 300
  DOF 184, DOFPulse
  If LI25.state <> 1 Then
    LI25.state = 1
    PlayerKimMLight(currentplayer)=1
    CheckGateOpenLights
  end If
End Sub

Sub CheckGateOpenLights()
  PlayBgEffect "SFX_Sample18"
  DOF 184, DOFPulse
  If LI23.state = 1 and LI24.state = 1 and LI25.state = 1 Then
    AddBonusScore
    PlayerKickBackActivated(CurrentPlayer) = PlayerKickBackActivated(CurrentPlayer) + 1
    if DebugGeneral then debug.print "Sub CheckGateLights : PlayerKickBackActivated current player = " & PlayerKickBackActivated(CurrentPlayer)
    LI23.state = 0
    PlayerKimKLight(currentplayer)=0
    LI24.state = 0
    PlayerKimILight(currentplayer)=0
    LI25.state = 0
    PlayerKimMLight(currentplayer)=0
    kickbackgate.open = true
    LI26.state = 1
    ki06.timerenabled = 0
    LightEffect LE_LeftRight
    ShowFlexScene1 "KICKBACK" , "ACTIVATED" , 1 , True , False
    AddSpeechToQueue "KIM-Kickback is now activated!" , 2300 , 5
  end if
end sub

Sub KI06_Hit()
  if tilted then KI06.Kick 0,70,0 : exit sub
  'Kickback gate/hole
  If KI06.TimerEnabled = 0 then
    PlayerKickBackActivated(CurrentPlayer) = PlayerKickBackActivated(CurrentPlayer) - 1
    if DebugGeneral then debug.print "Sub KI06_Hit : PlayerKickBackActivated current player = " & PlayerKickBackActivated(CurrentPlayer)
    ' Start Grace Timer
    If PlayerKickBackActivated(CurrentPlayer) = 0 then
      KI06.timerenabled = 1
      LI26.state = 2
    end if
  Else
    ShowFlexScene1 "KICKBACK" , "GRACE TIMER" , 1 , True , False
    PlayBgEffect "SFX_Sample18"
  End if
  vpmtimer.addtimer 2000, "KI06LaunchBall '"
End Sub

Sub KI06_Timer()
  KI06.TimerEnabled = 0
  kickbackgate.open = false
  LI26.state = 0
  PlayBgEffect "SFX_Sample18"
End Sub

Sub KI06LaunchBall()
  DOF 201, DOFPulse
  KI06.Kick 0,70,0
  PlaySoundAt SoundFXDOF("fx_Kicker_Release", 124, DOFPulse, DOFContactors), KI06
End SUb


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKUL] KIM UPPER LANES
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub T05_Hit()
  if tilted then exit sub
  addscore 400
  If bSkillShotReady and LI39.state = 2 Then
    AwardSkillshot
  end if
  AddBonusScore
  If LI27.state <> 1 Then
    DOF 183, DOFPulse
    LI27.state = 1
    CalculateBonusMultiplier
  end If
End Sub
Sub T06_Hit()
  if tilted then exit sub
  addscore 400
  If bSkillShotReady and LI40.state = 2 Then
    AwardSkillshot
  end if
  AddBonusScore
  If LI28.state <> 1 Then
    DOF 183, DOFPulse
    LI28.state = 1
    CalculateBonusMultiplier
  end If
End Sub
Sub T07_Hit()
  if tilted then exit sub
  addscore 400
  If bSkillShotReady and LI41.state = 2 Then
    AwardSkillshot
  end if
  AddBonusScore
  If LI29.state <> 1 Then
    DOF 183, DOFPulse
    LI29.state = 1
    CalculateBonusMultiplier
  end If
End Sub

Sub CalculateBonusMultiplier()
  dim StrA
  If LI27.STATE = 1 and LI28.state = 1 and LI29.state = 1 Then
    AddSpeechToQueue "KIM-You got a bonus multiplier!" , 2700 , 5
    playerbonusmultiplier(currentplayer) = playerbonusmultiplier(currentplayer) + 1
    StrA = cstr(playerbonusmultiplier(currentplayer)) & "X"
    ShowFlexScene1 StrA & " BONUS" , "MULTIPLIER" , 2 , True , False
    PlayBgEffect "SFX_Sample11"
    LI27.state = 0
    LI28.state = 0
    LI29.state = 0
    LightEffect LE_Random
'   LI6x.state = li5x.State
'   li5x.state = li4x.State
'   li4x.state = li3x.State
'   li3x.state = li2x.State
'   li2x.state = 1
    select case PlayerBonusMultiplier(currentplayer)
      case 2
        li2x.state = 1
      case 3
        li3x.state = 1
        li2x.state = 1
      case 4
        li4x.state = 1
        li3x.state = 1
        li2x.state = 1
      case 5
        li5x.state = 1
        li4x.state = 1
        li3x.state = 1
        li2x.state = 1
      case 6
        LI6x.state = 1
        li5x.state = 1
        li4x.state = 1
        li3x.state = 1
        li2x.state = 1
    end Select
    If PlayerBonusMultiplier(CurrentPlayer) > 6 Then
        LI6x.state = 2
        li5x.state = 1
        li4x.state = 1
        li3x.state = 1
        li2x.state = 1
    end if
    DOF 117, DOFPulse
    DOF 127, DOFPulse

  end if
end Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKSP] KIM SPINNERS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub Spinner002_Spin()
  if tilted then exit sub
  DOF 171,DOFPulse
  DOF 231, DOFPulse 'MX Spinner
  PlayBgEffect "SFX_Sample10"
  SoundSpinner Spinner002
  ObjSunLevel(4) = 1:FlasherSun4_Timer
  addscore 100
End Sub

Sub Spinner001_Spin()
  if tilted then exit sub
  DOF 170,DOFPulse
  DOF 230, DOFPulse 'MX Spinner
  PlayBgEffect "SFX_Sample10"
  SoundSpinner Spinner001
  ObjSunLevel(3) = 1:FlasherSun3_Timer
  addscore 100
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKBS] KIM BONUS SCORE
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub AddBonusScore()
  if tilted then exit sub
  if bSuperMagic = false and bMultiBallMode = false and bDiscoMode = false Then
    LightEffect LE_Score
  end If
  PlayBgEffect "SFX_Stolen"
  BonusPoints(CurrentPlayer) = BonusPoints(Currentplayer) + 1
  DOF 180, DOFPulse 'center to out blue
  ActivateBonusPointsLight(BonusPoints(currentplayer))
End Sub

Sub CollectBonusScore()
  bSuperMagic = False
  If BonusPoints(CurrentPlayer) > 0 then
    PlaySong "KIMSCORE.mp3"
    'ShowFlexScene7 "COLLECTING BONUS" , cstr(PlayerBonusMultiplier(currentplayer)*1000*BonusPoints(CurrentPlayer))+" POINTS" , 0.2 , 2 , True , False
'   TimerBonusCollect.interval = 200
'   TimerBonusCollect.enabled = 1
    ShowBonusStep1
    vpmtimer.addtimer 2600, "ShowBonusStep2 '"
    vpmtimer.addtimer 5200, "ShowBonusStep3 '"
    vpmtimer.addtimer 7800, "ShowBonusStep4 '"

  else
    vpmtimer.addtimer 3500, "EndOfBall2 '"
  end if
end Sub

Sub ShowBonusStep1
  playsound "SFX_Sample3" , 0 ,  VolumeDial , 0 ,0,1,1,1
  ShowFlexScene7 "BONUS" , cstr(1000*BonusPoints(CurrentPlayer)) + " POINTS" , 0.2 , 2 , False , False
End Sub
Sub ShowBonusStep2
  playsound "SFX_Sample3" , 0 ,  VolumeDial , 0 ,0,1,1,1
  ShowFlexScene7 "BONUS MULTIPLIER" , cstr(PlayerBonusMultiplier(currentplayer)) + " X" , 0.2 , 2 , False , False
End Sub
Sub ShowBonusStep3
  playsound "SFX_Sample4" , 0 ,  VolumeDial , 0 ,0,1,1,1
  ShowFlexScene7 "TOTAL BONUS" , cstr(PlayerBonusMultiplier(currentplayer)*1000*BonusPoints(CurrentPlayer))+" POINTS" , 0.2 , 2 , True , False
End Sub
Sub ShowBonusStep4
  TimerBonusCollect.interval = 200
  TimerBonusCollect.enabled = 1
End Sub


Sub TimerBonusCollect_Timer()
  'playsound "SFX_Xarax"
  If BonusPoints(CurrentPlayer) > 0 then
    addscore (1000 * playerbonusmultiplier(currentplayer))
    'playsound "SFX_Spinner" , 0 ,  VolumeDial , 0 ,0,1,1,1
    playsound "SFX_Sample10" , 0 ,  VolumeDial , 0 ,0,1,1,1
    BonusPoints(CurrentPlayer) = BonusPoints(Currentplayer) - 1
    ActivateBonusPointsLight(BonusPoints(currentplayer))
    TimerBonusCollect.enabled = 0
    If TimerBonusCollect.Interval > 50 Then
      TimerBonusCollect.Interval = TimerBonusCollect.Interval - 5
    end if
    TimerBonusCollect.enabled = 1
  else
    TimerBonusCollect.enabled = 0
    If PlayerBonusMultiplierHeld(CurrentPlayer) = 0 Then
      PlayerBonusMultiplier(currentplayer) = 1
    else
      ShowFlexScene1 "PLAYER "+cstr(currentplayer) , cstr(PlayerBonusMultiplier(currentplayer))+"x BONUS HELD" , 2 , True , False
      PlayBgEffect "SFX_Xarax"
    end if
    'stopsong
    vpmtimer.addtimer 1500, "EndOfBall2 '"
  end if
end sub


Sub ActivateBonusPointsLight(lScore)
  dim a
  For each a in ScoreLights
    a.state = 0
  Next

  If lScore < 10 Then
  end If
  If lScore > 9 and lScore < 20 Then
    LIScore10.state = 1
    lScore = Lscore - 10
  end If
  If lScore > 19 and lScore < 30 Then
    LIScore10.state = 1
    liScore20.state = 1
    lScore = Lscore - 20
  end If
  If lScore > 29 and lScore < 40 Then
    LIScore10.state = 1
    liScore20.state = 1
    liScore30.state = 1
    lScore = lscore - 30
  end If
  If lScore > 39 Then
    LIScore10.state = 1
    liScore20.state = 1
    liscore30.state = 2
    lScore = 9
  end If
  If lScore = 1 Then
    LIScore1.state = 1
  end If
  If lScore = 2 Then
    LIScore1.state = 1
    LIScore2.state = 1
  end If
  If lScore = 3 Then
    LIScore1.state = 1
    LIScore2.state = 1
    LIScore3.state = 1
  end If
  If lScore = 4 Then
    LIScore1.state = 1
    LIScore2.state = 1
    LIScore3.state = 1
    LIScore4.state = 1
  end If
  If lScore = 5 Then
    LIScore1.state = 1
    LIScore2.state = 1
    LIScore3.state = 1
    LIScore4.state = 1
    LIScore5.state = 1
  end If
  If lScore = 6 Then
    LIScore1.state = 1
    LIScore2.state = 1
    LIScore3.state = 1
    LIScore4.state = 1
    LIScore5.state = 1
    LIScore6.state = 1
  end If
  If lScore = 7 Then
    LIScore1.state = 1
    LIScore2.state = 1
    LIScore3.state = 1
    LIScore4.state = 1
    LIScore5.state = 1
    LIScore6.state = 1
    LIScore7.state = 1
  end If
  If lScore = 8 Then
    LIScore1.state = 1
    LIScore2.state = 1
    LIScore3.state = 1
    LIScore4.state = 1
    LIScore5.state = 1
    LIScore6.state = 1
    LIScore7.state = 1
    LIScore8.state = 1
  end If
  If lScore = 9 Then
    LIScore1.state = 1
    LIScore2.state = 1
    LIScore3.state = 1
    LIScore4.state = 1
    LIScore5.state = 1
    LIScore6.state = 1
    LIScore7.state = 1
    LIScore8.state = 1
    LIScore9.state = 1
  end If
end sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKBU] KIM BUNKER
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Sub T04_Hit()
  ' Target Hit
End Sub

Sub TargetKidsK_Hit()
  if tilted then exit sub
  If LI35.state = 0 Then
    DD_Spell_Kids
    LI35.state = 2
  else
    PlayBgEffect "SFX_Nuck"
    addscore 900
    If LI35.state = 2 Then
      LI35.state = 1
    end if
  end if
  PlayerKidsKLight(CurrentPlayer) = LI35.state
  CheckKidsLight
End Sub
Sub TargetKidsI_Hit()
  if tilted then exit sub
  If LI36.state = 0 Then
    DD_Spell_Kids
    LI36.state = 2
  else
    PlayBgEffect "SFX_Nuck"
    addscore 900
    If LI36.state = 2 Then
      LI36.state = 1
    end if
  end if
  PlayerKidsILight(CurrentPlayer) = LI36.state
  CheckKidsLight
End Sub
Sub TargetKidsD_Hit()
  if tilted then exit sub
  If LI37.state = 0 Then
    DD_Spell_Kids
    LI37.state = 2
  else
    PlayBgEffect "SFX_Nuck"
    addscore 900
    If LI37.state = 2 Then
      LI37.state = 1
    end if
  end if
  PlayerKidsDLight(CurrentPlayer) = LI37.state
  CheckKidsLight
End Sub
Sub TargetKidsS_Hit()
  if tilted then exit sub
  If LI38.state = 0 Then
    DD_Spell_Kids
    LI38.state = 2
  else
    PlayBgEffect "SFX_Nuck"
    addscore 900
    If LI38.state = 2 Then
      LI38.state = 1
    end if
  end if
  PlayerKidsSLight(CurrentPlayer) = LI38.state
  CheckKidsLight
End Sub

Sub CheckKidsLight()
  DOF 185, DOFPulse
  If LI35.state = 1 and LI36.State = 1 and LI37.State = 1 and LI38.state = 1 then
    if T04.IsDropped = 0 then
      AddSpeechToQueue "KIM-Hit the bunker to collect your bonus!" , 2700 , 5
      T04.IsDropped = 1
      playsoundat "fx_mine_motor" , T04
      TimerFlasherSun6.enabled = true
      LI47.state = 1
      ShowFlexScene1 "KIDS UNLOCKED" , "HIT THE BUNKER ", 1 , True , False
      AddBonusScore
      lightseqbunker.play SeqCircleInOn ,10, 1
    end if
  end if
End Sub

Sub TimerFlasherSun6_Timer()
  ObjSunlevel(6) = 2
  FlasherSun6_Timer
  vpmtimer.addtimer 400, "ObjSunlevel(6) = 1 '"
  vpmtimer.addtimer 410, "FlasherSun6_Timer '"
end sub

Sub DD_Spell_Kids()
    ShowFlexScene1 "SPELL KIDS TO" , "OPEN THE BUNKER" , 1 , True , False
    addscore 500
End Sub

Sub KI05_Hit()
  if tilted then KI05.kick 0,10,0 : exit sub

  If bMultiBallMode = False Then
    AddSpeechToQueue "KIM-You entered the bunker!" , 1700 , 5
    LightEffect LE_Random
    DOF 118, dofpulse ' (Strobe 2000 40)
    DOF 127, dofpulse '(Beacon 3000ms minimal)
    DOF 185, DOFPulse
    ShowFlexScene2 "BUNKER ENTERED", "COLLECTING BONUS" , 2 , 0.5 , 0.2 , True , False
    PlayBgEffect "SFX_70s"
    KI05.DestroyBall
    KI03.CreateSizedball BallSize / 2
    KI03.Timerinterval = 3000
    KI03.TimerEnabled = 1
    PlayBgEffect "SFX_Centura"
    StartFlash2Sequence()
    T04.Isdropped = 0
    playsoundat "fx_mine_motor" , T04
    TimerFlasherSun6.enabled = false
    LI47.state = 0
    LI35.state = 0
    LI36.state = 0
    LI37.state = 0
    LI38.state = 0
    addscore 8000
    LightEffect LE_Random
  Else
    ShowFlexScene1 "BUNKER HIT" , "KIDS BONUS" , 1 , True , False
    PlayBgEffect "SFX_70s"
    addscore 4000
    LightEffect LE_Random
    PlayBgEffect "SFX_Centura"
    KI05.kick 0,10,0
  end if
End Sub

Sub KI07_Hit()
  KI07.DestroyBall
  If tilted then KI02.CreateSizedball BallSize / 2 :  KI02.Kick 90, 30 : exit sub
  LightEffect LE_SeqUp
  select case int(RndNum(1,2))
  case 1
    StartFlash2Sequence()
    PlayBgEffect "SFX_faser6"
    KI03.CreateSizedball BallSize / 2
    KI03.Timerinterval = 1000
    KI03.TimerEnabled = 1
  case 2
    StartFlash3Sequence()
    PlayBgEffect "SFX_faser6"
    KI02.CreateSizedball BallSize / 2
    Timerki02b.interval = 1000
    Timerki02b.enabled = 1
  end select
End Sub

Sub TimerKI02B_Timer()
  Timerki02b.enabled = 0
  KI02.Kick 90, 30
  PlaySoundAt SoundFXDOF("fx_Kicker_Release", 159, DOFPulse, DOFContactors), KI02
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKAN] KIM SPINNING DISCS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub TimerDiscRotate_Timer()
  Dim a

  a = PrimitiveRampRing.objrotx
  if a > 359 then
    a = 0
  else
    a = a + 0.2
  end If
  PrimitiveRampRing.objrotx = a

  a = PrDiscoBall.roty
  if a > 359 then
    a = 0
  else
    a = a + 2
  end If
  PrDiscoBall.roty = a

End Sub

Sub TimerLpRotate_Timer()
  Dim a
  a = KimDisc3.objroty
  if a > 354 then
    a = 0'
  else
    a = a + 2
  end If
  KimDisc3.objroty = a
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZGLO] KIM GLOWING BALLS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Dim GlowBall, ChooseBall, CustomBulbIntensity(10), red3(10), green3(10), Blue3(10)
Dim CustomBallImage(10), CustomBallLogoMode(10), CustomBallDecal(10), CustomBallGlow(10)
ChooseBall = 0

' *** prepare the variable with references to three lights for glow ball ***
Dim Glowing(4)
Set Glowing(0) = Glowball0
Set Glowing(1) = Glowball1
Set Glowing(2) = Glowball2
Set Glowing(3) = Glowball3
Set Glowing(4) = Glowball4


' default Ball
CustomBallGlow(0) =     False
CustomBallImage(0) =    "ball_hdr"
CustomBallLogoMode(0) =   False
CustomBallDecal(0) =    "scratches"
CustomBulbIntensity(0) =  0.01
Red3(0) = 0 : Green3(0) = 0 : Blue3(0) = 0

' Magma Red GlowBall
CustomBallGlow(1) =     True
CustomBallImage(1) =    "ball_hdr"
CustomBallLogoMode(1) =   False
CustomBallDecal(1) =    "scratches"
CustomBulbIntensity(1) =  0
red3(1) = 255 : Green3(1) = 128 : Blue3(1) = 128

' Magma Blue GlowBall
CustomBallGlow(2) =     True
CustomBallImage(2) =    "ball_hdr"
CustomBallLogoMode(2) =   False
CustomBallDecal(2) =    "scratches"
CustomBulbIntensity(2) =  0
red3(2) = 128 : Green3(2) = 128 : Blue3(2) = 255

' Magma Green GlowBall
CustomBallGlow(3) =     True
CustomBallImage(3) =    "ball_hdr"
CustomBallLogoMode(3) =   False
CustomBallDecal(3) =    "scratches"
CustomBulbIntensity(3) =  0
red3(3) = 128 : Green3(3) = 255 : Blue3(3) = 128

'*** change ball appearance ***

Sub ChangeBall(ballnr)
  If DebugGeneral then debug.print "BallNummer " & cstr(ballnr)
  Dim BOT, ii, col
  Table1.BallDecalMode = CustomBallLogoMode(ballnr)
  Table1.BallFrontDecal = CustomBallDecal(ballnr)
  Table1.DefaultBulbIntensityScale = CustomBulbIntensity(ballnr)
  Table1.BallImage = CustomBallImage(ballnr)
  GlowBall = CustomBallGlow(ballnr)
  For ii = 0 to 4
    col = RGB(red3(ballnr), green3(ballnr), Blue3(ballnr))
    If DebugGeneral then debug.print cstr(ii) & "--->>> " & cstr(col)
    Glowing(ii).color = col : Glowing(ii).colorfull = col
  Next
End Sub

Sub ChangeBallDiscoMode()
  Dim BOT, ii, col
  GlowBall = True
  For ii = 0 to 4
    col = RGB(red3(ii+1), green3(ii+1), Blue3(ii+1))
    Glowing(ii).color = col : Glowing(ii).colorfull = col
  Next
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZKDB] KIM DISCO BALL ANIMATION
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub TimerHitDisco_Timer()
  lightseqdisco.play SeqCircleInOn ,10, 1
  LI48.State = 2
  TimerStopHitDisco.enabled = 1
  AddSpeechToQueue "KIM-Hit the right lane to start the Disco Mode!" , 2900 , 5
End Sub

Sub TimerStopHitDisco_Timer()
  TimerStopHitDisco.enabled = 0
  LI48.State = 0
End Sub

Sub Disco(Enabled)
  If Enabled Then
    bDiscoMode = true
    ' Kill all the playfield Lights
    lighteffect LE_AllOff
    gieffect 0
    DOF 216, DOFPulse
    TimerHitDisco.enabled = 0
    LIDisco.state = 1
    stopsong
    AddSpeechToQueue "KIM-Initiating Disco Party, hold on!" , 2500 , 5
    vpmtimer.addtimer 1500, "ShowDiscoBall True '"
    vpmtimer.addtimer 4000, "lasermotor True '"
    vpmtimer.addtimer 5000, "discoparty True '"
  Else
    DOF 126, DofOff ' beacon
    DOF 238, DofOff 'Stars
    bDiscoMode = false
    discoparty False
    lasermotor False
    showdiscoball false
    LIDisco.state = 0
    AddSpeechToQueue "KIM-I am so sorry but the Disco PArty is over!" , 4600 , 5
    TimerHitDisco.enabled = 1
  end if
end sub

Sub ShowDiscoBall(Enabled)
  If enabled Then
    playsound "FX_mine_Motor",0,0.5
    DiscoBallDownDelay = 28
    TimerLaserDiscoBallDown.Enabled = 1
    PrDiscoBall.visible = 1
  Else
    playsound "FX_mine_Motor",0,0.5
    DiscoBallUpDelay = 2
    TimerLaserDiscoBallUp.Enabled = 1
  end If
end sub

Sub DiscoParty(Enabled)
  If enabled Then
    DOF 126, dofon 'Beacon
    LightEffect LE_Random
    gieffect 2
    ki04releaseball
    AddMultiball 1
    AddSpeechToQueue "KIM-YES! , Lets get the disco party starting!" , 3600 , 5
    If KimTrackSpecial = 0 then
      playsong "KIMTRACKSPECIAL.MP3"
      dLine(2) = "DISCOPARTY - CAMBODIA & REPRISE"
    Else
      playsong "KIMTRACKSPECIAL2.MP3"
      dLine(2) = "DISCOPARTY - LES NUITS SANS KIM WILDE"
    end if
    SongDigitsUpdate()
    ChangeBallDiscoMode
    DOF 238, DofOn 'Stars
  Else
  end If
end sub

Sub LaserMotor(enabled)
  if Enabled then
    LaserN.visible = 1
    LaserNt.Enabled = 1
    TimerLaserColorChange.Enabled = 1
    LaserColorStep = 1
  Else
    LaserN.visible = 0
    LaserNt.Enabled = 0
    TimerLaserColorChange.Enabled = 0
  End If
End Sub

Sub LaserNt_Timer()
  FlashSeq Lasern, "disco", 10
End Sub

dim LaserStep : LaserStep = 0

Sub FlashSeq(object, imgseq, steps) 'Primitive texture image sequence
  dim x, x2
  object.ImageA = imgseq & LaserStep
  Laserstep = LaserStep + 1
  if Laserstep > steps-1 then Laserstep = 0
End Sub

dim LaserColorStep

Sub TimerLaserColorChange_Timer()
  Select Case LaserColorStep
  Case 1
    lasern.color = RGB(255,128,128)
  Case 2
    lasern.color = RGB(128,255,128)
  Case 3
    lasern.color = RGB(128,128,255)
  Case 3
    lasern.color = RGB(255,255,255)
  end select
  LaserColorstep = LaserColorStep + 1
  if LaserColorStep > 4 then LaserColorStep = 1

  DOF 900, DOFOff 'RGB Undercab
  DOF 901, DOFOff 'RGB Undercab
  DOF 902, DOFOff 'RGB Undercab
  DOF 903, DOFOff 'RGB Undercab
  DOF 904, DOFOff 'RGB Undercab
  DOF 905, DOFOff 'RGB Undercab

  dim a
  a = int(RndNum(1,6))

  select case a
    case 1
      DOF 900, DOFOn 'White RGB Undercab
    case 2
      DOF 901, DOFOn 'Red Undercab
    case 3
      DOF 902, DOFOn 'Green Undercab
    case 4
      DOF 903, DOFOn 'Blue Undercab
    case 5
      DOF 904, DOFOn 'Yellow Undercab
    case 6
      DOF 905, DOFOn 'Cyan Undercab
  end Select
End Sub

Dim DiscoBallDownDelay : DiscoBallDownDelay = 28
Sub TimerLaserDiscoBallDown_Timer()
  Dim a
  a = PrDiscoBall.Z
  a = a - DiscoBallDownDelay
  if a < 200 then TimerLaserDiscoBallDown.enabled = 0
  PrDiscoBall.z = a
    If Discoballdowndelay > 2 then DiscoBallDownDelay = DiscoBallDownDelay - 1
End Sub

Dim DiscoBallUpDelay : DiscoBallUpDelay = 2
Sub TimerLaserDiscoBallUp_Timer()
  Dim a
  a = PrDiscoBall.Z
  a = a + DiscoBallUpDelay
  if a > 700 then
    TimerLaserDiscoBallUp.enabled = 0
    PrDiscoBall.visible = 0
  end if
  PrDiscoBall.z = a
    If DiscoballUpdelay < 28 then DiscoBallUpDelay = DiscoBallUpDelay + 1
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' [ZZKSS] KIMS SONG SCROLLER JP's DMD Using Flasher, Scrolling song playing text routine using flashers
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Dim jpDigits, SongDigits, jpChars(255)

Sub JPDMD_Init
    Dim i
  ' Digits for song playing text
  SongDigits = Array(digit001, digit002, digit003, digit004, digit005, digit006, digit007, digit008, digit009, digit010, _
        digit011, digit012, digit013, digit014, digit015, digit016, digit017, digit018, digit019, digit020, _
        digit021, digit022, digit023, digit024, digit025, digit026, digit027, digit028, digit029, digit030)

    For i = 0 to 255:jpChars(i) = "dempty":Next


    jpChars(32) = "dempty"  '(Space)
  jpChars(45) = "dmin"     '-
    jpChars(46) = "dot"      '.
    jpChars(48) = "d0"       '0
    jpChars(49) = "d1"       '1
    jpChars(50) = "d2"       '2
    jpChars(51) = "d3"       '3
    jpChars(52) = "d4"       '4
    jpChars(53) = "d5"       '5
    jpChars(54) = "d6"       '6
    jpChars(55) = "d7"       '7
    jpChars(56) = "d8"       '8
    jpChars(57) = "d9"       '9
    jpChars(60) = "dless"    '<
    jpChars(61) = "dequal"   '=
    jpChars(62) = "dgreater" '>
    jpChars(64) = "bkempty"  '@
    jpChars(65) = "da"       'A
    jpChars(66) = "db"       'B
    jpChars(67) = "dc"       'C
    jpChars(68) = "dd"       'D
    jpChars(69) = "de"       'E
    jpChars(70) = "df"       'F
    jpChars(71) = "dg"       'G
    jpChars(72) = "dh"       'H
    jpChars(73) = "di"       'I
    jpChars(74) = "dj"       'J
    jpChars(75) = "dk"       'K
    jpChars(76) = "dl"       'L
    jpChars(77) = "dm"       'M
    jpChars(78) = "dn"       'N
    jpChars(79) = "do"       'O
    jpChars(80) = "dp"       'P
    jpChars(81) = "dq"       'Q
    jpChars(82) = "dr"       'R
    jpChars(83) = "ds"       'S
    jpChars(84) = "dt"       'T
    jpChars(85) = "du"       'U
    jpChars(86) = "dv"       'V
    jpChars(87) = "dw"       'W
    jpChars(88) = "dx"       'X
    jpChars(89) = "dy"       'Y
    jpChars(90) = "dz"       'Z
    jpChars(94) = "dup"      '^

End Sub

Sub DMDUpdate(id)
End Sub

Sub DMDDisplayChar(achar, adigit)
End Sub

Sub SongDigitsUpdate()
  TimerSongText.enabled = False
  TimerSongText.interval = 30
  TimerSongText.enabled = true
  dline(3) = space(30) & dLine(2) & space(50)
  Dim digit, value
  for digit = 0 to 23
    DMDSongsDisplayChar mid(ucase(dline(3)), digit + 1, 1), digit
  next
  TimerSongText.enabled = true
End Sub

Sub TimerSongtext_Timer()
  if len(dLine(3))>30 Then
    if len(dline(3)) > (len(dLine(2))+51) then
      if TimerSongText.interval <> 30 Then
        TimerSongText.enabled = False
        TimerSongText.interval = 30
        TimerSongText.enabled = true
      end If
    Elseif len(dline(3)) = (len(dLine(2))+51) then
      if TimerSongText.interval <> 5000 Then
      TimerSongText.enabled = False
      TimerSongText.interval = 5000
      TimerSongText.enabled = true
      end If
    else
      if TimerSongText.interval <> 100 Then
      TimerSongText.enabled = False
      TimerSongText.interval = 100
      TimerSongText.enabled = true
      end If
    end if
  end if
  if len(dLine(3))>30 Then
    dLine(3) = Right(dLine(3),len(dLine(3))-1)
  Else
    dLine(3) = space(30) & dLine(2) & space(50)
  end If
  SongDigitsUpdate2
End Sub

Sub SongDigitsUpdate2()
  Dim digit, value
  for digit = 0 to 23
    DMDSongsDisplayChar mid(ucase(dline(3)), digit + 1, 1), digit
  next
End Sub

Sub DMDSOngsDisplayChar(achar, adigit)
    If achar = "" Then achar = " "
    achar = ASC(achar)
    SongDigits(adigit).ImageA = jpChars(achar)
End Sub

Function CL(NumString) 'center line
    Dim Temp, TempStr
  NumString = LEFT(NumString, 20)
    Temp = (20- Len(NumString)) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Sub TimerVPXText_Timer()
  TimerVPXText.enabled = False
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZVOC] VOCAL QUEUE SYSTEM
'  Used to add vocal callouts to a queue so no overlap
'  Depends on TimerVocalQueue
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Dim nVocalQueue(255,3)
Dim nVocalQueueLength

Sub AddSpeechToQueue(strSample, numTime, numVariations)
  if DebugGeneral then debug.print "Sub AddSpeechToQueue() : " & strSample & " " & numTime & " " & numVariations

  nVocalQueue(nVocalQueueLength,0) = strSample
  nVocalQueue(nVocalQueueLength,1) = numTime
  nVocalQueue(nVocalQueueLength,2) = numVariations

  nVocalQueueLength = nVocalQueueLength + 1

  if TimerVocalQueue.enabled = 0 then
    TimerVocalQueue.interval = 5
    TimerVocalQueue.enabled = 1
  end if
End Sub

Sub TimerVocalQueue_Timer()
  dim a

  TimerVocalQueue.enabled = 0
  TimerVocalQueue.interval = nVocalQueue(0,1) 'sample length

  If nVocalQueueLength > 0 then
    ' There is at least 1 in the queue
    If nVocalQueue(0,2) > 1 Then
      'There are variations
      PlaySound nVocalQueue(0,0) & RndInt(1,nVocalQueue(0,2)) , 0 , nVoiceVolume ,0,0,1,1,1
    Else
      PlaySound nVocalQueue(0,0) , 0 , nVoiceVolume ,0,0,1,1,1
    end if
    TimerVocalQueue.enabled = 1
  end if

  if nVocalQueueLength > 1 then
    ' There are more then 1 in the queue, so shift all 1 down
    for a = 0 to nVocalQueueLength
      nVocalQueue(a,0) = nVocalQueue(a+1,0)
      nVocalQueue(a,1) = nVocalQueue(a+1,1)
      nVocalQueue(a,2) = nVocalQueue(a+1,2)
    next
  end If
  if  nVocalQueueLength > 0 then
    nVocalQueueLength = nVocalQueueLength - 1
  end if
End Sub

Sub PlayBgEffect (name)
  If bBgSFXSounds = 1 Then
    PlaySound name , 0 ,nSFXVolume , 0 , 0 , 1 , 1 , 1
  end If
end sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZMAT] COMMON FUNCTIONS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Function GetMusicFolder()
    Dim GTF
    Set GTF = CreateObject("Scripting.FileSystemObject")
    GetMusicFolder= GTF.GetParentFolderName(userdirectory) & "\Music"
    set GTF = nothing
End Function

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

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZNFF] FLIPPER CORRECTIONS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
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
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'
'        x.AddPt "Polarity", 0, 0, 0
'        x.AddPt "Polarity", 1, 0.05, - 2.7
'        x.AddPt "Polarity", 2, 0.16, - 2.7
'        x.AddPt "Polarity", 3, 0.22, - 0
'        x.AddPt "Polarity", 4, 0.25, - 0
'        x.AddPt "Polarity", 5, 0.3, - 1
'        x.AddPt "Polarity", 6, 0.4, - 2
'        x.AddPt "Polarity", 7, 0.5, - 2.7
'        x.AddPt "Polarity", 8, 0.65, - 1.8
'        x.AddPt "Polarity", 9, 0.75, - 0.5
'        x.AddPt "Polarity", 10, 0.81, - 0.5
'        x.AddPt "Polarity", 11, 0.88, 0
'        x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 3.7
    x.AddPt "Polarity", 2, 0.16, - 3.7
    x.AddPt "Polarity", 3, 0.22, - 0
    x.AddPt "Polarity", 4, 0.25, - 0
    x.AddPt "Polarity", 5, 0.3, - 2
    x.AddPt "Polarity", 6, 0.4, - 3
    x.AddPt "Polarity", 7, 0.5, - 3.7
    x.AddPt "Polarity", 8, 0.65, - 2.3
    x.AddPt "Polarity", 9, 0.75, - 1.5
    x.AddPt "Polarity", 10, 0.81, - 1
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

''*******************************************
''  Late 80's early 90's
'
'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5
'   x.AddPt "Polarity", 2, 0.16, - 5
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 4.0
'   x.AddPt "Polarity", 8, 0.7, - 3.5
'   x.AddPt "Polarity", 9, 0.75, - 3.0
'   x.AddPt "Polarity", 10, 0.8, - 2.5
'   x.AddPt "Polarity", 11, 0.85, - 2.0
'   x.AddPt "Polarity", 12, 0.9, - 1.5
'   x.AddPt "Polarity", 13, 0.95, - 1.0
'   x.AddPt "Polarity", 14, 1, - 0.5
'   x.AddPt "Polarity", 15, 1.1, 0
'   x.AddPt "Polarity", 16, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945

' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

'*******************************************
' Early 90's and after

'Sub InitPolarity()
' Dim x, a
' a = Array(LF, RF)
' For Each x In a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5.5
'   x.AddPt "Polarity", 2, 0.16, - 5.5
'   x.AddPt "Polarity", 3, 0.20, - 0.75
'   x.AddPt "Polarity", 4, 0.25, - 1.25
'   x.AddPt "Polarity", 5, 0.3, - 1.75
'   x.AddPt "Polarity", 6, 0.4, - 3.5
'   x.AddPt "Polarity", 7, 0.5, - 5.25
'   x.AddPt "Polarity", 8, 0.7, - 4.0
'   x.AddPt "Polarity", 9, 0.75, - 3.5
'   x.AddPt "Polarity", 10, 0.8, - 3.0
'   x.AddPt "Polarity", 11, 0.85, - 2.5
'   x.AddPt "Polarity", 12, 0.9, - 2.0
'   x.AddPt "Polarity", 13, 0.95, - 1.5
'   x.AddPt "Polarity", 14, 1, - 1.0
'   x.AddPt "Polarity", 15, 1.05, -0.5
'   x.AddPt "Polarity", 16, 1.1, 0
'   x.AddPt "Polarity", 17, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.23, 0.85
'   x.AddPt "Velocity", 2, 0.27, 1
'   x.AddPt "Velocity", 3, 0.3, 1
'   x.AddPt "Velocity", 4, 0.35, 1
'   x.AddPt "Velocity", 5, 0.6, 1 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

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
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
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
'   Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.035  'mid 80's to early 90's
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
    Dim b,  BOT
    BOT = GetBalls

    For b = 0 To UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
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

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZFLP] FLIPPERS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire 'rightflipper.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZFLB] FLUPPER BUMPERS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

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
  DNA30 = 0
  DNA45 = (NightDay - 10) / 20
  DNA90 = 0
  DayNightAdjust = 0.4
Else
  DNA30 = (NightDay - 10) / 30
  DNA45 = (NightDay - 10) / 45
  DNA90 = (NightDay - 10) / 90
  DayNightAdjust = NightDay / 25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt
For cnt = 1 To 6
  FlBumperActive(cnt) = False
Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight
FlInitBumper 1, "red"
FlInitBumper 2, "red"
FlInitBumper 4, "red"


' ### uncomment the statement below to change the color for all bumpers ###
'   Dim ind
'   For ind = 1 To 5
'    FlInitBumper ind, "green"
'   Next

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True

  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1
  FlBumperFadeTarget(nr) = 1.1
  FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr)
  FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
  Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
  FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
  FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)

  ' set the color for the two VPX lights
  Select Case col
    Case "red"
      FlBumperSmallLight(nr).color = RGB(255,4,0)
      FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
      FlBumperBigLight(nr).color = RGB(255,32,0)
      FlBumperBigLight(nr).colorfull = RGB(255,32,0)
      FlBumperHighlight(nr).color = RGB(64,255,0)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
      FlBumperSmallLight(nr).TransmissionScale = 0

    Case "blue"
      FlBumperBigLight(nr).color = RGB(32,80,255)
      FlBumperBigLight(nr).colorfull = RGB(32,80,255)
      FlBumperSmallLight(nr).color = RGB(0,80,255)
      FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
      FlBumperHighlight(nr).color = RGB(255,16,8)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "green"
      FlBumperSmallLight(nr).color = RGB(8,255,8)
      FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
      FlBumperBigLight(nr).color = RGB(32,255,32)
      FlBumperBigLight(nr).colorfull = RGB(32,255,32)
      FlBumperHighlight(nr).color = RGB(255,32,255)
      MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
      FlBumperSmallLight(nr).TransmissionScale = 0.005
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "orange"
      FlBumperHighlight(nr).color = RGB(255,130,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).color = RGB(255,130,0)
      FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
      FlBumperBigLight(nr).color = RGB(255,190,8)
      FlBumperBigLight(nr).colorfull = RGB(255,190,8)

    Case "white"
      FlBumperBigLight(nr).color = RGB(255,230,190)
      FlBumperBigLight(nr).colorfull = RGB(255,230,190)
      FlBumperHighlight(nr).color = RGB(255,180,100)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99

    Case "blacklight"
      FlBumperBigLight(nr).color = RGB(32,32,255)
      FlBumperBigLight(nr).colorfull = RGB(32,32,255)
      FlBumperHighlight(nr).color = RGB(48,8,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "yellow"
      FlBumperSmallLight(nr).color = RGB(255,230,4)
      FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
      FlBumperBigLight(nr).color = RGB(255,240,50)
      FlBumperBigLight(nr).colorfull = RGB(255,240,50)
      FlBumperHighlight(nr).color = RGB(255,255,220)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0

    Case "purple"
      FlBumperBigLight(nr).color = RGB(80,32,255)
      FlBumperBigLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).color = RGB(80,32,255)
      FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperHighlight(nr).color = RGB(32,64,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
  End Select
End Sub

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
  '   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  '        OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
  '        float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 ) * DayNightAdjust

  Select Case FlBumperColor(nr)
    Case "blue"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(38 - 24 * Z,130 - 98 * Z,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 500 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z + 0.97 * Z ^ 3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 10000 * (Z ^ 3) / (0.5 + DNA90)

    Case "green"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(16 + 16 * Sin(Z * 3.14),255,16 + 16 * Sin(Z * 3.14)), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
      FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 6000 * (Z ^ 3) / (1 + DNA90)

    Case "red"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 16 - 11 * Z + 16 * Sin(Z * 3.14),0), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z * 4,8 - Z * 8)

    Case "orange"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 100 - 22 * z + 16 * Sin(Z * 3.14),Z * 32), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z * 50, 0)

    Case "white"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
      FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 8 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255,255 - 20 * Z,255 - 65 * Z)
      FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20 * Z,255 - 65 * Z)
      MaterialColor "bumpertopmat" & nr, RGB(255,235 - z * 36,220 - Z * 90)

    Case "blacklight"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 1, RGB(30 - 27 * Z ^ 0.03,30 - 28 * Z ^ 0.01, 255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z ^ 3
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255 - 240 * (Z ^ 0.1),255 - 240 * (Z ^ 0.1),255)
      FlBumperSmallLight(nr).colorfull = RGB(255 - 200 * z,255 - 200 * Z,255)
      MaterialColor "bumpertopmat" & nr, RGB(255 - 190 * Z,235 - z * 180,220 + 35 * Z)

    Case "yellow"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 180 + 40 * z, 48 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

    Case "purple"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(128 - 118 * Z - 32 * Sin(Z * 3.14), 32 - 26 * Z ,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 15 + 200 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z + 0.97 * Z ^ 3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (0.5 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(128 - 60 * Z,32,255)
  End Select
End Sub

Sub BumperTimer_Timer
  Dim nr
  For nr = 1 To 6
    If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
      If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
    If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
      If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
  Next
End Sub

'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZMBA] MANUAL BALLCONTROL
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
  '

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -0.01 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)

Sub StartControl_Hit()
   Set ControlBall = ActiveBall
   contballinplay = true
End Sub

Sub StopControl_Hit()
   contballinplay = false
End Sub

Sub BallControl_Timer()
   If Contball and ContBallInPlay then
      If bcright = 1 Then
         ControlBall.velx = bcvel*bcboost
      ElseIf bcleft = 1 Then
         ControlBall.velx = - bcvel*bcboost
      Else
         ControlBall.velx=0
      End If

     If bcup = 1 Then
        ControlBall.vely = -bcvel*bcboost
     ElseIf bcdown = 1 Then
        ControlBall.vely = bcvel*bcboost
     Else
        ControlBall.vely= bcyveloffset
     End If
   End If
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZSHA] SHADOWS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Const anglecompensate = 15

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls

  IF GlowBall Then
    For b = 0 to 4
      If GlowBall and Glowing(b).state = 1 Then Glowing(b).state = 0 End If
    Next
  End If

    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 20
        If BOT(b).Z > 120 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    If GlowBall and b <4 Then
      If Glowing(b).state = 0 Then Glowing(b).state = 1 end if
      Glowing(b).BulbHaloHeight = BOT(b).z + 52
      Glowing(b).x = BOT(b).x : Glowing(b).y = BOT(b).y + anglecompensate
    End If

    Next
end sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZFLE] SOUND
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

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


'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'     * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'     * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'     * Create a Timer called RampRoll, that is enabled, with a interval of 100
'     * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'     * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'     * To stop tracking ball
'        * call WireRampOff
'        * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

Sub WireRampOn(input)
  Waddball ActiveBall, input
  RampRollUpdate
End Sub

Sub WireRampOff()
  WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  Dim x
  For x = 1 To UBound(RampBalls)  'Check, don't add balls twice
    If RampBalls(x, 1) = input.id Then
      If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 To UBound(RampBalls)
    If IsEmpty(RampBalls(x, 1)) Then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      Exit Sub
    End If
    If x = UBound(RampBalls) Then  'debug
      Debug.print "WireRampOn error, ball queue Is full: " & vbNewLine & _
      RampBalls(0, 0) & vbNewLine & _
      TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
      TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
      TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
      TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
      TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
      " "
    End If
  Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
  '   Debug.Print "In WRemoveBall() + Remove ball from loop array"
  Dim ballcount
  ballcount = 0
  Dim x
  For x = 1 To UBound(RampBalls)
    If ID = RampBalls(x, 1) Then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
  Next
  If BallCount = 0 Then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
  RampRollUpdate
End Sub

Sub RampRollUpdate()  'Timer update
  Dim x
  For x = 1 To UBound(RampBalls)
    If Not IsEmpty(RampBalls(x,1) ) Then
      If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
        If RampType(x) Then
          PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      End If
      If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
  Next
  If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
  "1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
  "2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
  "3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
  "4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
  "5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
  "6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
  " "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
  BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

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
'///////////////////////-----Mech and Relais-----///////////////////////
Dim MechSoundLevel
MechSoundLevel = 0.8

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
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

Const tnob = 5 ' total number of balls

'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 To tnob
    rolling(i) = False
  Next
End Sub

Sub RollingTimer_Timer()
  Dim b
  Dim BOT
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 To tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(BOT)
    If BallVel(BOT(b)) > 1 And BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ <  - 1 And BOT(b).z < 55 And BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz >  - 7 Then
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

Sub Kickers_Hit (idx) :SoundSaucerLock : End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZDMP] RUBBERS AND DAMPNERS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

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
      aBall.velz = aBall.velz * coef
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

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
Sub RDampen_Timer
  Cor.Update
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZSSC] : SLINGSHOT CORRECTION FUNCTIONS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

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

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZSEC] SECONDARY HIT EVENTS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'

Dim RStep, Lstep

Sub RightSlingShot1_slingshot
  PlaySoundAt SoundFXDOF("fx_slingshot", 104, DOFPulse, DOFContactors), RightInlane
End Sub

Sub LeftSlingShot1_slingshot
  PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors), LeftInlane
End Sub

'****************************************************************
' ZSLG: Slingshots
'****************************************************************

' RStep and LStep are the variables that increment the animation
' Dim RStep, LStep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  PlaySoundAt SoundFXDOF("fx_slingshot", 104, DOFPulse, DOFContactors), RightInlane
  ObjSunLevel(2)=1:FlasherSun2_Timer
  RSling1.Visible = 1
  Sling1.TransY =  - 20   'Sling Metal Bracket
  RStep = 0
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 10
  '   vpmTimer.PulseSw 52 'Slingshot Rom Switch
  RandomSoundSlingshotRight Sling1
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 3
      RSLing1.Visible = 0
      RSLing2.Visible = 1
      Sling1.TransY =  - 10
    Case 4
      RSLing2.Visible = 0
      Sling1.TransY = 0
      RightSlingShot.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors), LeftInlane
  ObjSunLevel(1)=1:FlasherSun1_Timer
  LSling1.Visible = 1
  Sling2.TransY =  - 20   'Sling Metal Bracket
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
  '   vpmTimer.PulseSw 51 'Slingshot Rom Switch
  RandomSoundSlingshotLeft Sling2
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 3
      LSLing1.Visible = 0
      LSLing2.Visible = 1
      Sling2.TransY =  - 10
    Case 4
      LSLing2.Visible = 0
      Sling2.TransY = 0
      LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZKEY] KEYS
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Sub Table1_KeyDown(ByVal Keycode)
  If EnableKeyPress = false then exit Sub
  If ballrolleron = 1 then
    if keycode = 46 then ' C Key
       If contball = 1 Then
          contball = 0
       Else
          contball = 1
       End If
    End If
  End If
  if keycode = 48 then 'B Key
     If bcboost = 1 Then
        bcboost = bcboostmulti
     Else
        bcboost = 1
     End If
  End If
  if keycode = 203 then bcleft = 1  ' Left Arrow
  if keycode = 200 then bcup = 1    ' Up Arrow
  if keycode = 208 then bcdown = 1  ' Down Arrow
  if keycode = 205 then bcright = 1   ' Right Arrow

  If keycode = PlungerKey Then
    Call SoundPlungerPull
    Plunger.Pullback
  End If

  'If keycode = 64 and dmdmode = 0 and turnonultradmd = 1 Then ShowDMDOptions cGameName 'F6 Key
  if keycode = 65 then Call ResetHS 'F7 reset High Scores

  If hsbModeActive = True Then
    ' We are entering High Score
    EnterHighScoreKey(keycode)
    elseif bGameInPlay Then
      If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25:CheckTilt
      If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25:CheckTilt
      If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound SoundFX("fx_nudge",0), 0, 1, 1, 0.25:CheckTilt
      if keycode = MechanicalTilt then Call CheckMechanicalTilt
      If NOT Tilted Then

        If keycode = LeftFlipperKey Then
          SolLFlipper True
          DOF 101, DOFON
          B2SLight 1, 300
          If bSkillshotReady = False Then
            Call RotateLaneLightsLeft
          End If
        End If

        If keycode = RightFlipperKey Then
          SolRFlipper True
          DOF 102, DOFON
          B2SLight 2, 300
          If bSkillshotReady = False Then
            RotateLaneLightsRight
          End If
        End If

        If keycode = StartGameKey Then
          If((PlayersPlayingGame < MaxPlayers) AND (bOnTheFirstBall = True) ) Then
            PlayersPlayingGame = PlayersPlayingGame + 1

            If PlayersPlayingGame = 2 Then
              If B2SOn Then
                Controller.B2SSetScorePlayer 2, 0
              End If
              AddSpeechToQueue "KIM-Two Players" , 1300 , 5
            End If

            If PlayersPlayingGame = 3 Then
              If B2SOn Then
                Controller.B2SSetScorePlayer 3, 0
              End If
              AddSpeechToQueue "KIM-Three Players" , 1700 , 5
            End If

            If PlayersPlayingGame = 4 Then
              If B2SOn Then
                Controller.B2SSetScorePlayer 4, 0
              End If
              AddSpeechToQueue "KIM-Four Players" , 1000 , 5
            End If

            TotalGamesPlayed = TotalGamesPlayed + 1
            call SaveGamesPlayed
          End If
        End If
      End If
    Else
    If NOT Tilted Then
      If keycode = StartGameKey Then
        If(BallsOnPlayfield = 0) Then
          ResetForNewGame()
        End If
      End If
    End If
  End If

  if keycode = "3" then ' Hit the 2 on your keyboard for testing thins during gameplay
    'KI02_Timer 'MultiBallBall
    disco true
    'DMDBigText2 "EXTRA" , "BALL" , 144 , 1
    'CheckHighscore
    'MultiBallBallRelease
    'StartSuperMagic
  end if

End Sub


Sub Table1_KeyUp(ByVal keycode)

  'Manual Ball Control

  if keycode = 203 then bcleft = 0 ' Left Arrow
  if keycode = 200 then bcup = 0 ' Up Arrow
  if keycode = 208 then bcdown = 0 ' Down Arrow
  if keycode = 205 then bcright = 0 ' Right Arrow

  If keycode = PlungerKey Then
    'PlaySoundAt "fx_plunger", Plunger
    SoundPlungerReleaseBall
    Plunger.Fire
  End If

  ' Table specific

  If bGameInPLay and hsbModeActive <> True Then

    If keycode = LeftFlipperKey Then
      SolLFlipper False
      DOF 101, DofOff
    End If
    If keycode = RightFlipperKey Then
      SolRFlipper False
      DOF 102, DofOff
    End If
  End If

End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZTIL] TILT
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub CheckMechanicalTilt()
  Call CheckTilt ' Mechanical Tilt does the same as nudge
End Sub

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Dim tilttime:tilttime = 0

Sub CheckTilt                                    'Called when table is nudged
  Tilt = Tilt + TiltSensitivity                'Add to tilt count
  TiltDecreaseTimer.Enabled = True
  If(Tilt> TiltSensitivity) AND(Tilt <15) Then 'show a warning
    AddSpeechToQueue "KIM-Tilt sensor activated!" , 2300 , 5
    'DD CenterTop("*** WARNING ***"), CenterBottom("* TILT SENSOR *"),"",eblinkfast,eblinkfast,0,500,True,"SFX_buzz"
    DOF 131, DOFPulse
    DOF 209, DOFPulse 'MX Effect
  End if
  If Tilt> 15 Then 'If more that 15 then TILT the table
    DOF 208, DOFPulse 'MX Effect
    Tilted = True
    dmdflush
    If B2SOn then
      Controller.B2SSetTilt 1
    end if
    AddSpeechToQueue "KIM-Tilt" , 7000 , 5
    'DD CenterTop("**** TILT ****"), CenterBottom("**** TILT ****"),"",eblink,eblink,0,1000,True,"SFX_powerdown"
    DisableTable True
    tilttableclear.enabled = true
    TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
  End If
End Sub

sub tilttableclear_timer
  tilttime = tilttime + 1
  Select Case tilttime
    Case 10
      tableclearing
  End Select
End Sub

Sub TiltDecreaseTimer_Timer
  ' DecreaseTilt
  If Tilt> 0 Then
    Tilt = Tilt - 0.1
  Else
    TiltDecreaseTimer.Enabled = False
  End If
End Sub

Sub TiltRecoveryTimer_Timer()
  If(BallsOnPlayfield = 0) Then
    EndOfBall()
    TiltRecoveryTimer.Enabled = False
  End If
End Sub

Sub DisableTable(Enabled)
  If Enabled Then
    GiOff
    LightSeqTilt.Play SeqAllOff
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
    LeftSlingshot.Disabled = 1
    RightSlingshot.Disabled = 1
    bumper1.hashitevent = false
    bumper2.hashitevent = false
    bumper4.hashitevent = false
    TopSlingShot1.hashitevent = False
    TopSlingShot3.hashitevent = False
    TopSlingShot4.hashitevent = false
    stopsong
  Else
    GiOn
    LightSeqTilt.StopPlay
    LeftSlingshot.Disabled = 0
    RightSlingshot.Disabled = 0
    bumper1.hashitevent = true
    bumper2.hashitevent = true
    bumper4.hashitevent = true
    TopSlingShot1.hashitevent = true
    TopSlingShot3.hashitevent = true
    TopSlingShot4.hashitevent = true
  End If
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'  [ZZFLD] FLUPPER DOMES
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
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
' Example: RotateFlasher 1, 180    'where 1 is the flasher number and 180 is the angle of Z rotation

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
'   ObjTargetLevel(1) = 1
' Else
'   ObjTargetLevel(1) = 0
' End If
'   FlasherFlash1_Timer
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
' ObjTargetLevel(1) = level/255 : FlasherFlash1_Timer
' End Sub

Sub Flash1(Enabled)
  If Enabled Then
    ObjTargetLevel(1) = 1
  Else
    ObjTargetLevel(1) = 0
  End If
  FlasherFlash1_Timer
  Sound_Flash_Relay enabled, Flasherbase1
End Sub

Sub Flash2(Enabled)
  If Enabled Then
    ObjTargetLevel(2) = 1
  Else
    ObjTargetLevel(2) = 0
  End If
  FlasherFlash2_Timer
  Sound_Flash_Relay enabled, Flasherbase2
End Sub

Sub Flash3(Enabled)
  If Enabled Then
    ObjTargetLevel(3) = 1
  Else
    ObjTargetLevel(3) = 0
  End If
  FlasherFlash3_Timer
  Sound_Flash_Relay enabled, Flasherbase3
End Sub


Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness
Dim FlasherFlareSunIntensity
' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object      ***
Set TableRef = Table1      ' *** change this, if your table has another name           ***
FlasherLightIntensity = 0.1  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)     ***
FlasherFlareIntensity = 0.3  ' *** lower this, if the flares are too bright (i.e. 0.1)       ***
FlasherBloomIntensity = 0.2  ' *** lower this, if the blooms are too bright (i.e. 0.1)       ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************
FlasherFlareSunIntensity = 0.3

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20), objSun(20), objSunLevel(20)

InitFlasher 1, "green"
InitFlasher 2, "yellow"
InitFlasher 3, "red"
'
InitSunFlasher 1 , "blue"
InitSunFlasher 2 , "blue"
InitSunFlasher 3 , "blue"
InitSunFlasher 4 , "blue"
InitSunFlasher 5 , "red"
InitSunFlasher 6 , "green"

InitSunFlasher 11 , "yellow"
InitSunFlasher 12 , "yellow"
InitSunFlasher 13 , "yellow"
InitSunFlasher 14 , "yellow"
InitSunFlasher 15 , "yellow"


Sub InitSunFlasher(nr,col)
  Set objSun(nr)=Eval("FlasherSun" & nr)
  select case col
    Case "blue" :    objSun(nr).color = RGB(4,32,255) :
    Case "green" :   objSun(nr).color = RGB(12,255,4)
    Case "red" :     objSun(nr).color = RGB(255,32,4)
    Case "purple" :  objSun(nr).color = RGB(255,64,255)
    Case "yellow" :  objSun(nr).color = RGB(255,200,50)
    Case "white" :   objSun(nr).color = RGB(100,86,59)
  end select
End Sub

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateFlasher 1,17
'   RotateFlasher 2,0
'   RotateFlasher 3,90
'   RotateFlasher 4,90

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr)
  Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr)
  Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)

  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbase(nr).x) / (objbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ
    objflasher(nr).height = objbase(nr).z + 40
  End If

  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0
  objlit(nr).visible = 0
  objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX
  objlit(nr).RotY = objbase(nr).RotY
  objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX
  objlit(nr).ObjRotY = objbase(nr).ObjRotY
  objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x
  objlit(nr).y = objbase(nr).y
  objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 Then
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
  objlight(nr).bulbhaloheight = objbase(nr).z - 10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  Dim xthird, ythird
  xthird = tablewidth / 3
  ythird = tableheight / 3
  If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  End If

  ' set the texture and color of all objects
  Select Case objbase(nr).image
    Case "dome2basewhite"
      objbase(nr).image = "dome2base" & col
      objlit(nr).image = "dome2lit" & col

    Case "ronddomebasewhite"
      objbase(nr).image = "ronddomebase" & col
      objlit(nr).image = "ronddomelit" & col

    Case "domeearbasewhite"
      objbase(nr).image = "domeearbase" & col
      objlit(nr).image = "domeearlit" & col
  End Select
  If TestFlashers = 0 Then
    objflasher(nr).imageA = "domeflashwhite"
    objflasher(nr).visible = 0
  End If
  Select Case col
    Case "blue"
      objlight(nr).color = RGB(4,120,255)
      objflasher(nr).color = RGB(200,255,255)
      objbloom(nr).color = RGB(4,120,255)
      objlight(nr).intensity = 5000

    Case "green"
      objlight(nr).color = RGB(12,255,4)
      objflasher(nr).color = RGB(12,255,4)
      objbloom(nr).color = RGB(12,255,4)

    Case "red"
      objlight(nr).color = RGB(255,32,4)
      objflasher(nr).color = RGB(255,32,4)
      objbloom(nr).color = RGB(255,32,4)

    Case "purple"
      objlight(nr).color = RGB(230,49,255)
      objflasher(nr).color = RGB(255,64,255)
      objbloom(nr).color = RGB(230,49,255)

    Case "yellow"
      objlight(nr).color = RGB(200,173,25)
      objflasher(nr).color = RGB(255,200,50)
      objbloom(nr).color = RGB(200,173,25)

    Case "white"
      objlight(nr).color = RGB(255,240,150)
      objflasher(nr).color = RGB(100,86,59)
      objbloom(nr).color = RGB(255,240,150)

    Case "orange"
      objlight(nr).color = RGB(255,70,0)
      objflasher(nr).color = RGB(255,70,0)
      objbloom(nr).color = RGB(255,70,0)
  End Select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle)
  angle = ((angle + 360 - objbase(nr).ObjRotZ) Mod 180) / 30
  objbase(nr).showframe(angle)
  objlit(nr).showframe(angle)
End Sub

Sub FlashFlasher(nr)
  If Not objflasher(nr).TimerEnabled Then
    objflasher(nr).TimerEnabled = True
    objflasher(nr).visible = 1
    objbloom(nr).visible = 1
    objlit(nr).visible = 1
  End If
  Sound_Flash_Relay 1, objflasher(nr)
  objflasher(nr).opacity = 1000 * FlasherFlareIntensity * ObjLevel(nr) ^ 2.5
  objbloom(nr).opacity = 100 * FlasherBloomIntensity * ObjLevel(nr) ^ 2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr) ^ 3
  objbase(nr).BlendDisableLighting = FlasherOffBrightness + 10 * ObjLevel(nr) ^ 3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr) ^ 2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  If Round(ObjTargetLevel(nr),1) > Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) + 0.3
    If ObjLevel(nr) > 1 Then ObjLevel(nr) = 1
  ElseIf Round(ObjTargetLevel(nr),1) < Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
    If ObjLevel(nr) < 0 Then ObjLevel(nr) = 0
  Else
    ObjLevel(nr) = Round(ObjTargetLevel(nr),1)
    objflasher(nr).TimerEnabled = False
  End If
  '   ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then
    objflasher(nr).TimerEnabled = False
    objflasher(nr).visible = 0
    objbloom(nr).visible = 0
    objlit(nr).visible = 0
    Sound_Flash_Relay 2, objflasher(nr)
  End If
End Sub


Sub SunFlasher(nr)
  On error resume next
  If not objSun(nr).TimerEnabled Then objSun(nr).TimerEnabled = True : objSun(nr).visible = 1 : End If
  objSun(nr).opacity = 1000 *  FlasherFlareSunIntensity * ObjSunLevel(nr)^2.5
  ObjSunLevel(nr) = ObjSunLevel(nr) * 0.9 - 0.01
  If ObjSunLevel(nr) < 0 Then objSun(nr).TimerEnabled = False : objSun(nr).visible = 0  : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : DOF 233, DOFPulse : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : DOF 232, DOFPulse : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : DOF 235, DOFPulse : End Sub
'
Sub FlasherSun1_Timer() : SunFlasher(1) : End Sub
Sub FlasherSun2_Timer() : SunFlasher(2) : End Sub
Sub FlasherSun3_Timer() : SunFlasher(3) : End Sub
Sub FlasherSun4_Timer() : SunFlasher(4) : End Sub
Sub FlasherSun5_Timer() : SunFlasher(5) : End Sub
Sub FlasherSun6_Timer() : SunFlasher(6) : End Sub

Sub FlasherSun11_Timer() : SunFlasher(11) : End Sub
Sub FlasherSun12_Timer() : SunFlasher(12) : End Sub
Sub FlasherSun13_Timer() : SunFlasher(13) : End Sub
Sub FlasherSun14_Timer() : SunFlasher(14) : End Sub
Sub FlasherSun15_Timer() : SunFlasher(15) : End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   [ZZFLX] FLEX DMD
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

' FlexDMD constants
Const   FlexDMD_RenderMode_DMD_GRAY_2 = 0, _
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

Dim FlexScoreMode : FlexScoreMode = True
Dim FlexDMD
Dim Frame : Frame = 0

Sub InitFlexDMD
  If UseFlexDMD = 0 Then Exit Sub
  Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
    If FlexDMD is Nothing Then
        MsgBox "No FlexDMD found. This table will NOT run without it."
        Exit Sub
    End If
  SetLocale(1033)
  With FlexDMD
    .GameName = cGameName
    .TableFile = Table1.Filename & ".vpx"
    '.Color = RGB(255, 88, 32)
    .Color = RGB(32, 88, 255)
    .RenderMode = FlexDMD_RenderMode_DMD_RGB
    .Width = 128
    .Height = 32
    .ProjectFolder = "./KimWildeDMD/"
    .Clear = True
    .Run = True
  End With

  call CreateScenes()

  ShowFlexScene1 "Starting Table" , "Please wait......." , 10 , False , False
End Sub

Dim FlexScenes(100)        'Array of FlexDMD scenes
Dim FontScoreInactive, FontScoreActive  'Two fonts used globally by timer score script

Sub CreateScenes()
  DotMatrix.color = RGB(255,255,255)
  Dim bigFont : Set bigFont = FlexDMD.NewFont("sys80.fnt", vbWhite , vbBlack, 0)
  Dim bigFontShadow : Set bigFontShadow = FlexDMD.NewFont("sys80.fnt", RGB ( 10,10,10) ,vbBlack, 0)
  Set FontScoreActive = FlexDMD.NewFont("TeenyTinyPixls5.fnt", vbWhite, vbWhite, 0)
  Set FontScoreInactive = FlexDMD.NewFont("TeenyTinyPixls5.fnt", RGB(100, 100, 100), vbWhite, 0)

  ' FlexScene SCORE
  Set FlexScenes(0) = FlexDMD.NewGroup("Score")
  With FlexScenes(0)
    If DebugGeneral then debug.print "EnableDMDBG =" & EnableDMDBG
    If EnableDMDBG = 1 then
      .AddActor FlexDMD.NewImage("bg","DMD-BG-SCORE.png")
      .Getimage("bg").visible = True ' False
    End If

    Dim i
    For i = 1 to 4
      .AddActor FlexDMD.NewLabel("Score_" & i, FontScoreInactive, "0")
    Next
    .AddActor FlexDMD.NewFrame("VSeparator")
    .GetFrame("VSeparator").Thickness = 1
    .GetFrame("VSeparator").SetBounds 45, 0, 1, 32
    .AddActor FlexDMD.NewGroup("Content")
    .GetGroup("Content").Clip = True
    .GetGroup("Content").SetBounds 47, 0, 81, 32
    .GetGroup("Content").AddActor FlexDMD.NewLabel("Top", FontScoreActive, ">>> KIM WILDE <<<")
    .GetGroup("Content").AddActor FlexDMD.NewLabel("Line1s",  bigFontShadow, " SCORE ")
    .GetGroup("Content").AddActor FlexDMD.NewLabel("Line1",  bigFont, " SCORE ")
    .GetGroup("Content").AddActor FlexDMD.NewLabel("Ball", FontScoreActive, " ")
    .GetGroup("Content").AddActor FlexDMD.NewLabel("Credit", FontScoreActive, "Free Play")
  End With

  Set FlexScenes(1) = FlexDMD.NewGroup("ModeType1")
  With FlexScenes(1)
    If EnableDMDBG = 1 then
      .AddActor FlexDMD.NewImage("bg","DMD-BG.png")
    End if
    .AddActor FlexDMD.NewLabel("Line1s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line1", bigFont, " ")
    .AddActor FlexDMD.NewLabel("Line2s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line2", bigFont, " ")
  End With

  Set FlexScenes(2) = FlexDMD.NewGroup("ModeType2")
  With FlexScenes(2)
    If EnableDMDBG = 1 then
      .AddActor FlexDMD.NewImage("bg","DMD-BG.png")
    End if
    .AddActor FlexDMD.NewLabel("Line1s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line1", bigFont, " ")
    .AddActor FlexDMD.NewLabel("Line2s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line2", bigFont, " ")
  End With

  Set FlexScenes(3) = FlexDMD.NewGroup("ModeType3")
  With FlexScenes(3)
    If EnableDMDBG = 1 then
      .AddActor FlexDMD.NewImage("bg","DMD-BG.png")
    End if
    .AddActor FlexDMD.NewLabel("Line1s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line1", bigFont, " ")
    .AddActor FlexDMD.NewLabel("Line2s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line2", bigFont, " ")
  End With

  Set FlexScenes(4) = FlexDMD.NewGroup("ModeType4")
  With FlexScenes(4)
    If EnableDMDBG = 1 then
      .AddActor FlexDMD.NewImage("bg","DMD-BG.png")
    End if
    .AddActor FlexDMD.NewLabel("Line1s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line1", bigFont, " ")
    .AddActor FlexDMD.NewLabel("Line2s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line2", bigFont, " ")
  End With

  Set FlexScenes(5) = FlexDMD.NewGroup("ModeType5")
  With FlexScenes(5)
    If EnableDMDBG = 1 then
      .AddActor FlexDMD.NewImage("bg","DMD-BG.png")
    End if
    .AddActor FlexDMD.NewLabel("Line1s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line1", bigFont, " ")
    .AddActor FlexDMD.NewLabel("Line2s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line2", bigFont, " ")
  End With

  Set FlexScenes(6) = FlexDMD.NewGroup("ModeType6")
  With FlexScenes(6)
    If EnableDMDBG = 1 then
      .AddActor FlexDMD.NewImage("bg","DMD-BG.png")
    End if
    .AddActor FlexDMD.NewLabel("Line1s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line1", bigFont, " ")
    .AddActor FlexDMD.NewLabel("Line2s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line2", bigFont, " ")
  End With

  Set FlexScenes(7) = FlexDMD.NewGroup("ModeType7")
  With FlexScenes(7)
    If EnableDMDBG = 1 then
      .AddActor FlexDMD.NewImage("bg","DMD-BG.png")
    End if
    .AddActor FlexDMD.NewLabel("Line1s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line1", bigFont, " ")
    .AddActor FlexDMD.NewLabel("Line2s", bigFontShadow, " ")
    .AddActor FlexDMD.NewLabel("Line2", bigFont, " ")
  End With
End Sub

Sub ShowFlexScene0
  debug.print "Scene0"
  FlexScoreMode = True 'Enable Score DMD Timer
  With FlexDMD
    .LockRenderThread
    .RenderMode = FlexDMD_RenderMode_DMD_RGB
    .Stage.RemoveAll
    .Stage.AddActor FlexScenes(0)
    .UnlockRenderThread
  End With
End Sub

Sub ShowFlexScene1(TextLine1 , TextLine2 , ShowTime , ShowScoreAfterwards , PlayNotification)
' Displays 2 lines of text
  FlexScoreMode = False 'Disable Score DMD
  FlexDMD.LockRenderThread
  FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_RGB
  FlexDMD.Stage.RemoveAll

  With FlexScenes(1)
    .GetLabel("Line1").Text = TextLine1
    .GetLabel("Line1s").Text = TextLine1
    .GetLabel("Line2").Text = TextLine2
    .GetLabel("Line2s").Text = TextLine2
    .GetLabel("Line1").SetAlignedPosition 64, 8, FlexDMD_Align_Center
    .GetLabel("Line1s").SetAlignedPosition 65, 9, FlexDMD_Align_Center
    .GetLabel("Line2").SetAlignedPosition 64, 23, FlexDMD_Align_Center
    .GetLabel("Line2s").SetAlignedPosition 65, 24, FlexDMD_Align_Center
  End With

  FlexDMD.Stage.AddActor FlexScenes(1)
  FlexDMD.UnlockRenderThread

  If ShowScoreAfterwards = True then
    vpmtimer.addtimer ShowTime * 1000, "ShowFlexScene0 '"
  End If

End Sub

Sub ShowFlexScene2(TextLine1 , TextLine2, ShowTime , FlashIntervalLine1 , FlashIntervalLine2 , ShowScoreAfterwards , PlayNotification )
' Displays 2 lines of text flashing
  FlexScoreMode = False 'Disable Score DMD
  FlexDMD.LockRenderThread
  FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_RGB
  FlexDMD.Stage.RemoveAll

  With FlexScenes(2)
    .GetLabel("Line1").Text = TextLine1
    .GetLabel("Line1s").Text = TextLine1
    .GetLabel("Line2").Text = TextLine2
    .GetLabel("Line2s").Text = TextLine2
    .GetLabel("Line1").SetAlignedPosition 64, 8, FlexDMD_Align_Center
    ' Shadow Line
    .GetLabel("Line1s").SetAlignedPosition 65, 9, FlexDMD_Align_Center
    .GetLabel("Line2").SetAlignedPosition 64, 23, FlexDMD_Align_Center
    ' Shadow Line
    .GetLabel("Line2s").SetAlignedPosition 65, 24, FlexDMD_Align_Center

    Dim af

    Set af = .GetLabel("Line1s").ActionFactory
    dim Line1sSeq : Set Line1sSeq = af.Sequence()

    Line1sSeq.Add af.Show(False)
    Line1sSeq.Add af.Wait(FlashIntervalLine1)
    Line1sSeq.Add af.Show(True)
    Line1sSeq.Add af.Wait(FlashIntervalLine1)
    .GetLabel("Line1s").ClearActions
    .GetLabel("Line1s").AddAction af.Repeat(Line1sSeq, -1)

    Set af = .GetLabel("Line1").ActionFactory
    dim Line1Seq : Set Line1Seq = af.Sequence()

    Line1Seq.Add af.Show(False)
    Line1Seq.Add af.Wait(FlashIntervalLine1)
    Line1Seq.Add af.Show(True)
    Line1Seq.Add af.Wait(FlashIntervalLine1)
    .GetLabel("Line1").ClearActions
    .GetLabel("Line1").AddAction af.Repeat(Line1Seq, -1)

    Set af = .GetLabel("Line2s").ActionFactory
    dim Line2sSeq : Set Line2sSeq = af.Sequence()

    Line2sSeq.Add af.Show(False)
    Line2sSeq.Add af.Wait(FlashIntervalLine2)
    Line2sSeq.Add af.Show(True)
    Line2sSeq.Add af.Wait(FlashIntervalLine2)
    .GetLabel("Line2s").ClearActions
    .GetLabel("Line2s").AddAction af.Repeat(Line2sSeq, -1)

    Set af = .GetLabel("Line2").ActionFactory
    dim Line2Seq : Set Line2Seq = af.Sequence()

    Line2Seq.Add af.Show(False)
    Line2Seq.Add af.Wait(FlashIntervalLine2)
    Line2Seq.Add af.Show(True)
    Line2Seq.Add af.Wait(FlashIntervalLine2)
    .GetLabel("Line2").ClearActions
    .GetLabel("Line2").AddAction af.Repeat(Line2Seq, -1)
  End With

  FlexDMD.Stage.AddActor FlexScenes(2)
  FlexDMD.UnlockRenderThread

  If ShowScoreAfterwards = True then
    vpmtimer.addtimer ShowTime * 1000, "ShowFlexScene0 '"
  End If

End Sub

Sub ShowFlexScene3(TextLine1 , TextLine2 , ScrollTime , HoldTime, ShowScoreAfterwards , PlayNotification )
' Display 2 lines, scrolling from up to down
  FlexScoreMode = False 'Disable Score DMD
  FlexDMD.LockRenderThread
  FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_RGB
  FlexDMD.Stage.RemoveAll

  With FlexScenes(3)
    .GetLabel("Line1").Text = TextLine1
    .GetLabel("Line1s").Text = TextLine1
    .GetLabel("Line2").Text = TextLine2
    .GetLabel("Line2s").Text = TextLine2
    .GetLabel("Line1s").ClearActions
    .GetLabel("Line1").ClearActions
    .GetLabel("Line2s").ClearActions
    .GetLabel("Line2").ClearActions

    Dim FontWidth : FontWidth = 6
    Dim LineOffset : LineOffset = (128 - (FontWidth * len(TextLine1))) / 2
    Dim Effect : Effect = 1.25

    Dim af,afs
    Set afs = .GetLabel("Line1s").ActionFactory
    dim Line1sSeq : Set Line1sSeq = afs.Sequence()
    Set af = .GetLabel("Line1").ActionFactory
    dim Line1Seq : Set Line1Seq = af.Sequence()

    If HoldTime = 0 then
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, -32 + 1, 0)
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, 32 + 1, ScrollTime)
      Line1Seq.Add af.MoveTo(LineOffset, -32, 0)
      Line1Seq.Add af.MoveTo(LineOffset, 32, ScrollTime)
    Else
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, -32 + 1, 0)
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, 2 + 1, ScrollTime)
      Line1sSeq.Add afs.Wait(HoldTime)
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, 32 + 1, ScrollTime)
      Line1Seq.Add af.MoveTo(LineOffset, -32, 0)
      Line1Seq.Add af.MoveTo(LineOffset, 2, ScrollTime)
      Line1Seq.Add af.Wait(HoldTime)
      Line1Seq.Add af.MoveTo(LineOffset, 32, ScrollTime)
    End if
    .GetLabel("Line1s").AddAction afs.Repeat(Line1sSeq, 1)
    .GetLabel("Line1").AddAction af.Repeat(Line1Seq, 1)

    LineOffset = (128 - (FontWidth * len(TextLine2))) / 2

    Set afs = .GetLabel("Line2s").ActionFactory
    dim Line2sSeq : Set Line2sSeq = afs.Sequence()
    Set af = .GetLabel("Line2").ActionFactory
    dim Line2Seq : Set Line2Seq = af.Sequence()

    If HoldTime = 0 then
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, -16 + 1, 0)
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, 48 + 1, ScrollTime)
      Line2Seq.Add af.MoveTo(LineOffset, -16, 0)
      Line2Seq.Add af.MoveTo(LineOffset, 48, ScrollTime)
    Else
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, -16 + 1, 0)
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, 16 + 1, ScrollTime / Effect)
      Line2sSeq.Add afs.Wait(HoldTime )
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, 48 + 1, ScrollTime / Effect)
      Line2Seq.Add af.MoveTo(LineOffset, -16, 0)
      Line2Seq.Add af.MoveTo(LineOffset, 16, ScrollTime / Effect)
      Line2Seq.Add af.Wait(HoldTime  )
      Line2Seq.Add af.MoveTo(LineOffset, 48, ScrollTime / Effect)
    End If
    .GetLabel("Line2s").AddAction afs.Repeat(Line2sSeq, 1)
    .GetLabel("Line2").AddAction af.Repeat(Line2Seq, 1)
  End With

  FlexDMD.Stage.AddActor FlexScenes(3)
  FlexDMD.UnlockRenderThread

  If ShowScoreAfterwards = True then
    vpmtimer.addtimer ((ScrollTime * 1000) * 2) + (holdtime * 1000), "ShowFlexScene0 '"
  End If
End Sub

Sub ShowFlexScene4(TextLine1 , TextLine2 , ScrollTime , HoldTime , ShowScoreAfterwards , PlayNotification  )
' Display 2 lines, scrolling from down to up
  FlexScoreMode = False 'Disable Score DMD
  FlexDMD.LockRenderThread
  FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_RGB
  FlexDMD.Stage.RemoveAll

  With FlexScenes(4)
    .GetLabel("Line1").Text = TextLine1
    .GetLabel("Line1s").Text = TextLine1
    .GetLabel("Line2").Text = TextLine2
    .GetLabel("Line2s").Text = TextLine2
    .GetLabel("Line1s").ClearActions
    .GetLabel("Line1").ClearActions
    .GetLabel("Line2s").ClearActions
    .GetLabel("Line2").ClearActions

    Dim FontWidth : FontWidth = 6
    Dim LineOffset : LineOffset = (128 - (FontWidth * len(TextLine1))) / 2

    Dim Effect : Effect = 1.25

    Dim af,afs
    Set afs = .GetLabel("Line1s").ActionFactory
    dim Line1sSeq : Set Line1sSeq = afs.Sequence()
    Set af = .GetLabel("Line1").ActionFactory
    dim Line1Seq : Set Line1Seq = af.Sequence()

    If HoldTime = 0 then
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, 32 + 1, 0)
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, -32 + 1, ScrollTime)
      Line1Seq.Add af.MoveTo(LineOffset, 32, 0)
      Line1Seq.Add af.MoveTo(LineOffset, -32, ScrollTime)
    Else
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, 32 + 1, 0)
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, 2 + 1, ScrollTime / Effect)
      Line1sSeq.Add afs.Wait(HoldTime)
      Line1sSeq.Add afs.MoveTo(LineOffset + 1, -32 + 1, ScrollTime / Effect)
      Line1Seq.Add af.MoveTo(LineOffset, 32, 0)
      Line1Seq.Add af.MoveTo(LineOffset, 2, ScrollTime / Effect)
      Line1Seq.Add af.Wait(HoldTime)
      Line1Seq.Add af.MoveTo(LineOffset, -32, ScrollTime / Effect)
    End if
    .GetLabel("Line1s").AddAction afs.Repeat(Line1sSeq, 1)
    .GetLabel("Line1").AddAction af.Repeat(Line1Seq, 1)

    LineOffset = (128 - (FontWidth * len(TextLine2))) / 2

    Set afs = .GetLabel("Line2s").ActionFactory
    dim Line2sSeq : Set Line2sSeq = afs.Sequence()
    Set af = .GetLabel("Line2").ActionFactory
    dim Line2Seq : Set Line2Seq = af.Sequence()

    If HoldTime = 0 then
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, 48 + 1, 0)
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, -16 + 1, ScrollTime)
      Line2Seq.Add af.MoveTo(LineOffset, 48, 0)
      Line2Seq.Add af.MoveTo(LineOffset, -16, ScrollTime)
    Else
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, 48 + 1, 0)
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, 16 + 1, ScrollTime )
      Line2sSeq.Add afs.Wait(HoldTime )
      Line2sSeq.Add afs.MoveTo(LineOffset + 1, -16 + 1, ScrollTime)
      Line2Seq.Add af.MoveTo(LineOffset, 48, 0)
      Line2Seq.Add af.MoveTo(LineOffset, 16, ScrollTime)
      Line2Seq.Add af.Wait(HoldTime  )
      Line2Seq.Add af.MoveTo(LineOffset, -16, ScrollTime)
    End If
    .GetLabel("Line2s").AddAction afs.Repeat(Line2sSeq, 1)
    .GetLabel("Line2").AddAction af.Repeat(Line2Seq, 1)

  End With

  FlexDMD.Stage.AddActor FlexScenes(4)
  FlexDMD.UnlockRenderThread

  If ShowScoreAfterwards = True then
    vpmtimer.addtimer ((ScrollTime * 1000) * 2) + (holdtime * 1000), "ShowFlexScene0 '"
  End If
End Sub

Sub ShowFlexScene5(TextLine1 , TextLine2 , ScrollTime , HoldTime , ShowScoreAfterwards , PlayNotification )
  ' Display 2 lines, both scrolling right to left
  FlexScoreMode = False 'Disable Score DMD
  FlexDMD.LockRenderThread
  FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_RGB
  FlexDMD.Stage.RemoveAll

  With FlexScenes(5)
    .GetLabel("Line1").Text = TextLine1
    .GetLabel("Line1s").Text = TextLine1
    .GetLabel("Line2").Text = TextLine2
    .GetLabel("Line2s").Text = TextLine2

    ' Calculate to center lines
    Dim FontWidth : FontWidth = 6
    Dim Offset1, Offset2
    Dim TextLine1Len , TextLine2Len

    TextLine1Len = (FontWidth * len(TextLine1))
    TextLine2Len = (FontWidth * len(TextLine2))
    If TextLine1Len > TextLine2Len then
      Offset2 = (TextLine1Len - TextLine2Len) / 2
    Else
      Offset1 = (TextLine2Len - TextLine1Len) / 2
    end if

    Dim af,afs
    Set af = .GetLabel("Line1").ActionFactory
    dim Line1Seq : Set Line1Seq = af.Sequence()
    Set afs = .GetLabel("Line1s").ActionFactory
    dim Line1sSeq : Set Line1sSeq = afs.Sequence()


    If HoldTime = 0 then
      Line1sSeq.Add afs.MoveTo(128 + Offset1 + 1, 2 + 1, 0)
      Line1sSeq.Add afs.MoveTo(-128 + Offset1 + 1, 2 + 1, ScrollTime)
      Line1Seq.Add af.MoveTo(128 + Offset1, 2, 0)
      Line1Seq.Add af.MoveTo(-128 + Offset1, 2, ScrollTime)
    Else
      Line1sSeq.Add afs.MoveTo(128 + 1, 2 + 1, 0)
      Line1sSeq.Add afs.MoveTo(((128 - TextLine1Len) / 2) + 1, 2 + 1, ScrollTime )
      Line1sSeq.Add afs.Wait(HoldTime)
      Line1sSeq.Add afs.MoveTo(- TextLine1Len + 1, 2 + 1, ScrollTime )
      Line1Seq.Add af.MoveTo(128, 2, 0)
      Line1Seq.Add af.MoveTo(((128 - TextLine1Len) / 2) , 2, scrollTime )
      Line1Seq.Add af.Wait(HoldTime)
      Line1Seq.Add af.MoveTo(- TextLine1Len, 2, ScrollTime )
    End If
    .GetLabel("Line1s").ClearActions
    .GetLabel("Line1s").AddAction afs.Repeat(Line1sSeq, 1)
    .GetLabel("Line1").ClearActions
    .GetLabel("Line1").AddAction af.Repeat(Line1Seq, 1)

    Set af = .GetLabel("Line2").ActionFactory
    dim Line2Seq : Set Line2Seq = af.Sequence()
    Set afs = .GetLabel("Line2s").ActionFactory
    dim Line2sSeq : Set Line2sSeq = afs.Sequence()

    If HoldTime = 0 then
      Line2sSeq.Add afs.MoveTo(128 + Offset1 + 1, 17 + 1, 0)
      Line2sSeq.Add afs.MoveTo(-128 + Offset1 + 1, 17 + 1, ScrollTime)
      Line2Seq.Add af.MoveTo(128 + Offset1, 17, 0)
      Line2Seq.Add af.MoveTo(-128 + Offset1, 17, ScrollTime)
    Else
      Line2sSeq.Add afs.MoveTo(128 + 1, 17 + 1, 0)
      Line2sSeq.Add afs.MoveTo(((128 - TextLine2Len) / 2) + 1 , 17 + 1, scrollTime )
      Line2sSeq.Add afs.Wait(HoldTime)
      Line2sSeq.Add afs.MoveTo(- TextLine2Len + 1, 17 + 1, ScrollTime )
      Line2Seq.Add af.MoveTo(128, 17, 0)
      Line2Seq.Add af.MoveTo(((128 - TextLine2Len) / 2) , 17, scrollTime )
      Line2Seq.Add af.Wait(HoldTime)
      Line2Seq.Add af.MoveTo(- TextLine2Len, 17, ScrollTime )
    End If
    .GetLabel("Line2s").ClearActions
    .GetLabel("Line2s").AddAction afs.Repeat(Line2sSeq, 1)
    .GetLabel("Line2").ClearActions
    .GetLabel("Line2").AddAction af.Repeat(Line2Seq, 1)
  End With

  FlexDMD.Stage.AddActor FlexScenes(5)
  FlexDMD.UnlockRenderThread

  If ShowScoreAfterwards = True then
    vpmtimer.addtimer ((ScrollTime * 1000) * 2) + (HoldTime * 1000), "ShowFlexScene0 '"
  End If
End Sub

Sub ShowFlexScene6(TextLine1 , TextLine2 , ScrollTime , HoldTime , ShowScoreAfterwards , PlayNotification )
  ' Display 2 lines, both scrolling left to rights
  FlexScoreMode = False 'Disable Score DMD
  FlexDMD.LockRenderThread
  FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_RGB
  FlexDMD.Stage.RemoveAll

  With FlexScenes(6)
    .GetLabel("Line1").Text = TextLine1
    .GetLabel("Line1s").Text = TextLine1
    .GetLabel("Line2").Text = TextLine2
    .GetLabel("Line2s").Text = TextLine2

    ' Calculate to center lines
    Dim FontWidth : FontWidth = 6
    Dim Offset1, Offset2
    Dim TextLine1Len , TextLine2Len

    TextLine1Len = (FontWidth * len(TextLine1))
    TextLine2Len = (FontWidth * len(TextLine2))
    If TextLine1Len > TextLine2Len then
      Offset2 = (TextLine1Len - TextLine2Len) / 2
    Else
      Offset1 = (TextLine2Len - TextLine1Len) / 2
    end if

    Dim af,afs
    Set af = .GetLabel("Line1").ActionFactory
    dim Line1Seq : Set Line1Seq = af.Sequence()
    Set afs = .GetLabel("Line1s").ActionFactory
    dim Line1sSeq : Set Line1sSeq = afs.Sequence()

    If HoldTime = 0 then
      Line1sSeq.Add afs.MoveTo(-128 + Offset1 + 1, 2 + 1, 0)
      Line1sSeq.Add afs.MoveTo(128 + Offset1 + 1, 2 + 1, ScrollTime)
      Line1Seq.Add af.MoveTo(-128 + Offset1, 2, 0)
      Line1Seq.Add af.MoveTo(128 + Offset1, 2, ScrollTime)
    Else
      Line1sSeq.Add afs.MoveTo(- TextLine1Len + 1, 2 + 1, 0)
      Line1sSeq.Add afs.MoveTo(((128 - TextLine1Len) / 2) + 1 , 2 + 1, scrollTime )
      Line1sSeq.Add afs.Wait(HoldTime)
      Line1sSeq.Add afs.MoveTo(128 + 1, 2 + 1, ScrollTime )
      Line1Seq.Add af.MoveTo(- TextLine1Len, 2, 0)
      Line1Seq.Add af.MoveTo(((128 - TextLine1Len) / 2) , 2, scrollTime )
      Line1Seq.Add af.Wait(HoldTime)
      Line1Seq.Add af.MoveTo(128, 2, ScrollTime )
    End If
    .GetLabel("Line1s").ClearActions
    .GetLabel("Line1s").AddAction afs.Repeat(Line1sSeq, 1)
    .GetLabel("Line1").ClearActions
    .GetLabel("Line1").AddAction af.Repeat(Line1Seq, 1)

    Set af = .GetLabel("Line2").ActionFactory
    dim Line2Seq : Set Line2Seq = af.Sequence()
    Set afs = .GetLabel("Line2s").ActionFactory
    dim Line2sSeq : Set Line2sSeq = afs.Sequence()
    If HoldTime = 0 then
      Line2sSeq.Add afs.MoveTo(128 + Offset2 + 1, 17 + 1, 0)
      Line2sSeq.Add afs.MoveTo(-128 + Offset2 + 1, 17 + 1, ScrollTime)
      Line2Seq.Add af.MoveTo(128 + Offset2, 17, 0)
      Line2Seq.Add af.MoveTo(-128 + Offset2, 17, ScrollTime)
    Else
      Line2sSeq.Add afs.MoveTo(- TextLine2Len + 1, 17 + 1, 0)
      Line2sSeq.Add afs.MoveTo(((128 - TextLine2Len) / 2) + 1 , 17 + 1, scrollTime )
      Line2sSeq.Add afs.Wait(HoldTime)
      Line2sSeq.Add afs.MoveTo(128 + 1, 17 + 1, ScrollTime )
      Line2Seq.Add af.MoveTo(- TextLine2Len, 17, 0)
      Line2Seq.Add af.MoveTo(((128 - TextLine2Len) / 2) , 17, scrollTime )
      Line2Seq.Add af.Wait(HoldTime)
      Line2Seq.Add af.MoveTo(128, 17, ScrollTime )
    End If
    .GetLabel("Line2s").ClearActions
    .GetLabel("Line2s").AddAction afs.Repeat(Line2sSeq, 1)
    .GetLabel("Line2").ClearActions
    .GetLabel("Line2").AddAction af.Repeat(Line2Seq, 1)
  End With

  FlexDMD.Stage.AddActor FlexScenes(6)
  FlexDMD.UnlockRenderThread

  If ShowScoreAfterwards = True then
    vpmtimer.addtimer ((ScrollTime * 1000) * 2) + (HoldTime * 1000), "ShowFlexScene0 '"
  End If
End Sub

Sub ShowFlexScene7(TextLine1 , TextLine2 , ScrollTime , HoldTime , ShowScoreAfterwards , PlayNotification )
  ' Display 2 lines, both scrolling different direction
  FlexScoreMode = False 'Disable Score DMD
  FlexDMD.LockRenderThread
  FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_RGB
  FlexDMD.Stage.RemoveAll

  With FlexScenes(7)
    .GetLabel("Line1").Text = TextLine1
    .GetLabel("Line1s").Text = TextLine1
    .GetLabel("Line2").Text = TextLine2
    .GetLabel("Line2s").Text = TextLine2

    ' Calculate to center lines
    Dim FontWidth : FontWidth = 6
    Dim Offset1, Offset2
    Dim TextLine1Len , TextLine2Len

    TextLine1Len = (FontWidth * len(TextLine1))
    TextLine2Len = (FontWidth * len(TextLine2))
    If TextLine1Len > TextLine2Len then
      Offset2 = (TextLine1Len - TextLine2Len) / 2
    Else
      Offset1 = (TextLine2Len - TextLine1Len) / 2
    end if

    Dim af,afs
    Set afs = .GetLabel("Line1s").ActionFactory
    dim Line1sSeq : Set Line1sSeq = afs.Sequence()
    Set af = .GetLabel("Line1").ActionFactory
    dim Line1Seq : Set Line1Seq = af.Sequence()

    If HoldTime = 0 then
      Line1sSeq.Add afs.MoveTo(128 + Offset1 + 1, 2 + 1, 0)
      Line1sSeq.Add afs.MoveTo(-128 + Offset1 + 1, 2 + 1, ScrollTime)
      Line1Seq.Add af.MoveTo(128 + Offset1, 2, 0)
      Line1Seq.Add af.MoveTo(-128 + Offset1, 2, ScrollTime)
    Else
      Line1sSeq.Add afs.MoveTo(128 + 1, 2 + 1, 0)
      Line1sSeq.Add afs.MoveTo(((128 - TextLine1Len) / 2) + 1 , 2 + 1, scrollTime )
      Line1sSeq.Add afs.Wait(HoldTime)
      Line1sSeq.Add afs.MoveTo(- TextLine1Len + 1, 2 + 1, ScrollTime )
      Line1Seq.Add af.MoveTo(128, 2, 0)
      Line1Seq.Add af.MoveTo(((128 - TextLine1Len) / 2) , 2, scrollTime )
      Line1Seq.Add af.Wait(HoldTime)
      Line1Seq.Add af.MoveTo(- TextLine1Len, 2, ScrollTime )
    End If
    .GetLabel("Line1s").ClearActions
    .GetLabel("Line1s").AddAction afs.Repeat(Line1sSeq, 1)
    .GetLabel("Line1").ClearActions
    .GetLabel("Line1").AddAction af.Repeat(Line1Seq, 1)

    Set afs = .GetLabel("Line2s").ActionFactory
    dim Line2sSeq : Set Line2sSeq = afs.Sequence()
    Set af = .GetLabel("Line2").ActionFactory
    dim Line2Seq : Set Line2Seq = af.Sequence()
    If HoldTime = 0 then
      Line2sSeq.Add afs.MoveTo(128 + Offset2 + 1, 17 + 1, 0)
      Line2sSeq.Add afs.MoveTo(-128 + Offset2 + 1, 17 + 1, ScrollTime)
      Line2Seq.Add af.MoveTo(128 + Offset2, 17, 0)
      Line2Seq.Add af.MoveTo(-128 + Offset2, 17, ScrollTime)
    Else
      Line2sSeq.Add afs.MoveTo(- TextLine2Len + 1, 17 + 1, 0)
      Line2sSeq.Add afs.MoveTo(((128 - TextLine2Len) / 2) + 1 , 17 + 1, scrollTime )
      Line2sSeq.Add afs.Wait(HoldTime)
      Line2sSeq.Add afs.MoveTo(128 + 1, 17 + 1, ScrollTime )
      Line2Seq.Add af.MoveTo(- TextLine2Len, 17, 0)
      Line2Seq.Add af.MoveTo(((128 - TextLine2Len) / 2) , 17, scrollTime )
      Line2Seq.Add af.Wait(HoldTime)
      Line2Seq.Add af.MoveTo(128, 17, ScrollTime )
    End If
    .GetLabel("Line2s").ClearActions
    .GetLabel("Line2s").AddAction afs.Repeat(Line2sSeq, 1)
    .GetLabel("Line2").ClearActions
    .GetLabel("Line2").AddAction af.Repeat(Line2Seq, 1)
  End With

  FlexDMD.Stage.AddActor FlexScenes(7)
  FlexDMD.UnlockRenderThread

  If ShowScoreAfterwards = True then
    vpmtimer.addtimer ((ScrollTime * 1000) * 2) + (HoldTime * 1000), "ShowFlexScene0 '"
  End If
End Sub

Dim DMDTextOnScore
Dim DMDTextOnScoreDisplayTime
Dim DMDTopTextOnScore
Dim DMDTopTextOnScoreDisplayTime
Dim DMDTextOnScoreEffect

Sub DMDDisplayTextOnScore( Text2Display , DisplayTime , Flash)
  DMDTextOnScore = Text2Display
  DMDTextOnScoreDisplayTime = frame + ((DisplayTime * 1000) / DMDTimer.Interval)
  DMDTextOnScoreEffect = Flash
End Sub

Sub DMDDisplayTopTextOnScore( Text2Display , DisplayTime )
  DMDTopTextOnScore = Text2Display
  DMDTopTextOnScoreDisplayTime = frame + + ((DisplayTime * 1000) / DMDTimer.Interval)
end sub

Sub DMDTimer_Timer
  Dim i, n, x, y, label
  Frame = Frame + 1

  If FlexScoreMode = True then
    FlexDMD.LockRenderThread

    If (Frame Mod 16) = 0 Then
      For i = 1 to 4
        Set label = FlexDMD.Stage.GetLabel("Score_" & i)
        If i = CurrentPlayer Then
          label.Font = FontScoreActive
        Else
          label.Font = FontScoreInactive
        End If
        label.Text = FormatNumber(Score(i), 0)
        label.SetAlignedPosition 45, 1 + (i - 1) * 6, FlexDMD_Align_TopRight
      Next
    End If

    If DMDTopTextOnScoreDisplayTime > Frame then
      FlexDMD.Stage.GetLabel("Top").Text = DMDTopTextOnScore
    Else
      FlexDMD.Stage.GetLabel("Top").Text = ">>> KIM WILDE <<<"
      DMDTopTextOnScore = " "
    End if

    If DMDTextOnScoreDisplayTime > frame Then
      'There is temporary a text to display
      If DMDTextOnScoreEffect = 1 And (frame Mod 20) > 10 Then 'Flashing
        FlexDMD.Stage.GetLabel("Line1").Text = " "
        FlexDMD.Stage.GetLabel("Line1s").Text = " "
      Else
        FlexDMD.Stage.GetLabel("Line1").Text = DMDTextOnScore
        FlexDMD.Stage.GetLabel("Line1s").Text = DMDTextOnScore
      End If
    Else
        FlexDMD.Stage.GetLabel("Line1").Text = FormatNumber(Score(CurrentPlayer), 0)
        FlexDMD.Stage.GetLabel("Line1s").Text = FormatNumber(Score(CurrentPlayer), 0)
    End If

    FlexDMD.Stage.GetLabel("Ball").Text = "Ball " & cstr(balls)
    FlexDMD.Stage.GetLabel("Ball").SetAlignedPosition 0, 33, FlexDMD_Align_BottomLeft
    FlexDMD.Stage.GetLabel("Credit").SetAlignedPosition 81, 33, FlexDMD_Align_BottomRight
    FlexDMD.Stage.GetLabel("Line1s").SetAlignedPosition 43, 17, FlexDMD_Align_Center
    FlexDMD.Stage.GetLabel("Line1").SetAlignedPosition 42, 16, FlexDMD_Align_Center
    FlexDMD.Stage.GetLabel("Top").SetAlignedPosition 42, 5, FlexDMD_Align_Center
    FlexDMD.UnlockRenderThread
  End If
End Sub

Sub DotMatrix_Timer
  Dim DMDp
  If FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_RGB Then
    DMDp = FlexDMD.DmdColoredPixels
    If Not IsEmpty(DMDp) Then
      DMDWidth = FlexDMD.Width
      DMDHeight = FlexDMD.Height
      DMDColoredPixels = DMDp
    End If
  Else
    DMDp = FlexDMD.DmdPixels
    If Not IsEmpty(DMDp) Then
      DMDWidth = FlexDMD.Width
      DMDHeight = FlexDMD.Height
      DMDPixels = DMDp
    End If
  End If
End Sub





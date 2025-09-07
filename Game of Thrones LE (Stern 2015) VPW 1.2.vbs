'        ,@@@********@@@@
'     .@@%             (&         %            @              &    %%%%%%%%%%,
'    @@@                         @@#           @@.          .@@      @@*    /
'   (@@                         @% @/          @@@%        #@@@      @@*           .@@&@(   (@@@@@
'   @@@                        #@  /@*        #@ &@@      @@ &@.     @@/....@    &&  %*   @  @
'   (@@,              @@@     /@@@@@@@*       @%  /@@    @&   @&     @@/    @    @   %*   &# @@@..
'    @@@,             @@@    ,@      @@/      @    .@@  @/    @@     @@/          @  %*   @  @
'     &@@@            @@@   .@        @@@    ,@      @@@      @@.    @@/     (       #@(    . .
'       #@@@@         @@@             @@@@@(%@@@#     ,       @@@@% %@@@@@@@@,
'           *@@@@@@@@@/
'                           @@&##################################################################################################%@*
'                          ,@       .@@@                                                                                          @.
'                                   .@@@  ...  .      ...  . .&&&&&&%*             &@@@@@(        &          .%###(  &&&&&&&&&@      .&@@@@,
'                                   .@@@    @@*         #@@    (@@    &@@      @@   , (@   @@     @@@           @.     @@     (    @@     /
'                                   .@@@    @@*         #@@    (@@     @@,   *@@    , (@    (@%   @ #@@/        @.     @@          @@(
'                                   .@@@    @@&%%%%%%%%%@@@    (@@     @*    @@     , (@     @@,  @   ,@@@      @.     @@,,,,#@      @@@
'                                   .@@@    @@,         #@@    (@@@@@@@      @@.    , (@     &@(  @      @@@    @.     @@     .        %@@&
'                                   .@@@    @@,         #@@    (@@    #@@    .@@    , (@     @@   @.       @@@, @.     @@                (@@
'                                   .@@@    @@,         #@@    (@@      @@@    /@&  , (@   /@/    @,         *@@@      @@      *          @,
'                                   .@@@  .&@@@&,      &@@@@* @@@@@*      /@@@(.   &@@@@@@.     @@&&@/          @    *@@@@@@@@@    &@@@@(
'                                 #%#((#%@#
'
' Version Log
' 100 - Initial beta release
' 101 - Fix playfield friction, flipper geometry, move rubbers to separate layer,
'       fix FlexDMD and StartTGHurryUp script bugs, add cabinet-mode option to hide side-rails
'       Add quick-restart (FreePlay mode only, hold left flipper and hit Start to restart game)
'       Add flipper fast-forward through battle intro scenes
' 102 - replace rubber post walls with primitives, add nfozzy flipper physics to upper PF flippers, import VPW physics materials,
'       adjust Bumper and Sling strengths to match VPW ref, fix battering ram impact math, add ball saver grace period,
'       add random house option, fix speech in Choose House, fix minor Iron Throne mode bugs, add high score saving to HotK, WHC, IT modes,
'       fix Award Replay to run after Extra Ball or Bonus awarded, write Mystery SuperJackpot support
'       write Bonus Hold support
'       Switched to VPW LUTs. Hold left magnasave and press right magnasave to toggle through
' 103 - "Choose Random" moved to between Martell and Targaryen, fix UFP scoring bug, fix stacked battle not ending properly on drain,
'       fix Targaryen HurryUp bug. Award held bonus immediatley if on last ball. Pause battle mode timers when ball is in the bumpers.
'       Add Hound callout for timer countdown during BW Super Jackpot. Add flipper speeches during Attract Mode
'       Add Midnight Madness multiball. Make sure your system clock is set
'   Add Spinner and Bumper Flashers
' 104 - In Iron Throne mode, don't add extra balls after the 7 main castles are done
'       Use flippers to advance through Attract mode
' 105 - oqqsan LAMPZ / inserts and extra lights / working rgb .. some colors need "help"
' 106 - Add Daphishbowl's Service Menu, refactored for FlexDMD. Integration Service Menu options into code
' 109 - Tomate Blender work, Niwak VLM Toolkit integration work started
' 110 - fix up missing/misnamed primitives, add code to animate all versions of 3D prims
' 111 - new batch added, inserts and flashers splitted, some material corrections and more fixes in blender side
' 112 - renamed flashers and moved RMEK LMEK to movable collection
' 113 - Wired all the new Blender images into Lampz
' 114 - new batch added, upper pf inserts, gates, dragon mouth light, lateral castle plastics fixed in blender
' 115 - Code new Blender lightmaps
' 116 - Sixtoe - Fix ghost ball issue, significant refactoring of physics level objects, removal of old non-used assets, filled in lots of ball trap holes, removed secondary GI bits and peices, removed lots of old primitives, removed duplicates, hopefully haven't broken anything...
' 117 - Fixed UpperPF Inserts (Thanks Apophis!), Add timer to wildfire targets after 2 Blackwaters. Refactor UpperPF logic to handle stacked Blackwater+Battle mode;
'       Fix GI variable Intensity durng WiC HurryUps; Add a few more ballsave callouts; fix a Service Menu bug; Move switch wires on hit
' 118 - Sixtoe - Fixed upper playfield shield, cut lamps so nothing sticks outside the cabinet.
' 119 - Sixtoe - Fixed uppoer playfield that I broke
' 120 - apophis - Physics review and update (materials, scripts, settings). Added slingshot corrections. New droptarget sound effects. Added dynamic shadows and flipper shadows.
' 121 - apophis - Small tweaks to make post pass work well and to make right ramp backhand shot more reliable.
' 121a - apophis - Added primitive ramp from tomate. Ramp triggers and gates  referenced to correct surface heights. PF friction set to 0.22
' 122 - Fix Castle loop not resetting UPF lights; Possible fix for 'out of stack space' error; misc bug fixes
' 123 - tomate - new Blender render
' 124 - Merge in Blender changes; Refactor high score sequence; Fix Ball save bug; use timer to trigger Baratheon battle hit after target hit, to hopefully fix OOM errors; Bump up time for battle instructions
' 125/126 - tomate - New blender render: fix PF inserts, fix target orientation, add lockbar light
' 127 - Merge in blender changes
' 128 - Sixtoe - Fixed ramp to UPF, Fixed drop targets not dropping (vlm.solid to vlm.active), fixed insert primitives and lighting and rough pass on light tuning, redid light materials, removed old text layers, removed vpx bloom lighting, removed rundundant images from file, added Apophis Backdrop,
' 129 - apophis - Fixed target rotations. Hooked up standup target animations. Updating target BM and LM animations at same rate. Added dynamic ball brightness code.
' 130 - Sixtoe - Add VR components
' 131 - Add Options menu; minor bug fixes
' 132 - Rawd - Configured, placed, tweaked and coded all the VR room components for Hybrid use
' 133 - Finish lockbar button lamp logic; bug fixes
' 134 - tomate - new Blender render
' 135 - Merge blender changes into script; wire up apron/lockbar button lights; bug fixes
' 136 - tomate - new Blender render
' 137 - Merge blender changes in; bug fixes
' 138 - Rawd - New VR topper with lights and dragon
' 139 - tomate - new Blender render
' 140 - Move Options DMD to backglass to fix VR bug; Attempt at Ball shadows on upper PF; minor bug fixes
' 141 - tomate - 4K blender render, VR fixes
' 142 - Merge blender changes; actually fix HS entry; actually fix UPF ball shadows; fix Stern Service menu; add kick to balls behind sword
' 145 - leojreimroc - Redid VR Backglass Lighting
' 146 - apophis - Added ramp rolling sounds. Updated ball rolling sounds. Adjusted ball brightness. Adjusted desktop POV. Plunger tweaks.
' 147 - Sixtoe - Image and file optimisation, removed old unused stuff, blocker wall added under upper playfield
' 148 - minor bug fixes
' 149 - Fixed a few crashes, boost intensity of shield lamps, tweak flipper strength
' 150 - apophis - Moved upper PF inserts to z=169. Added desktop action button. Digital nudge strength increased a tad.
' 151 - minor bug fixes; Add warning if FlexDMD dir not found

' 1.0.0 - Release
' 1.0.1 - Fix DOF fire hazard, Freeplay toggle, score formatting regression; Save ROM volume; give kick to balls behind sword
' 1.0.2 - Fix race condition that could end multiball prematurely; Add Iron Bank Cashout for Tyrell; fix occasional hangs in Mystery and Match; Allow "EB is lit" to stack;
'         fix coin sound after first game; fix fast-reset; implement some missing Service Menu options; Fix bug that would start HOTK prematurely; Fix bug that would use Mystery to complete an already completed house
' Unreleased - DaRdog - created Iron Throne VR Room
' 1.0.3 - bthlonewolf - Fix missing FlexDMD images. Tweaked/softened Iron Throne kicker. Adding DOF to items like battering ram, castle black, multiball (see DOF list).
'     Replaced sounds fx_kicker (missing) with fx_AutoPlunger, DTDrop (missing) with Drop_Target_Down_n.
' 1.0.3.1 - bthlonewolf - Tweaked Throne ramps, added FlexDMD locking test
' 1.0.3.2 - bthlonewolf - Tweaked locking in NewSceneWithImage, added some in NewSceneWithVideo when loading images. Replaced fx_flipperdown (missing) with Flipper_Left_Down_1/Flipper_Right_Down_1
' 1.0.3.3 - bthlonewolf - Tweaked DOF for MB modes, DOF is now on throughout entire MB sequence; use delays/max in DOF to fine tune intro/etc.
' 1.0.3.4 - bthlonewolf - Added DOF for: IT Kicker, Extra Ball, Dragon Roar, Dragon Wings; Folded in DaRdog's IT VR Room
' 1.0.3.5 - dgrimmreaper - Fix VR button animations and added animation for fire button
' 1.0.3.6 - bthlonewolf - Added HauntFreak's ball options in Tweak menu. Reset DMD when score rolls over 1bil to adjust font size
' 1.0.3.7 - mcarter78 - Updated flipper triggers & physics code, updated code for roth standups/drops, removed 10ms timers & moved animation calls to FrameTimer
' 1.0.3.8 - bthlonewolf - Migrated ball choice to DMD menu, added DMD settings reminder to F12 menu; Implemented Apophis' recommendation of altering elasticity falloff of zCol_Targets 0.3 -> 0.05 and zCol_DropTarget 0.08 -> 0.02
' 1.0.3.9 - DaRdog - add 2 new mega rooms, now with 3: Iron Throne, Dragon Ice Mega, Tower Throne Mega
' 1.1 - bthlonewolf - tweaked flexdmd shutdown. Fix for adding player during active scene. Add VR to settings. Minor service menu tweak when closing.
' 1.1.1 - DaRdog - improved VR poly count
' 1.1.2 - bthlonewolf - fixed VR settings changes
' 1.1.3 - apophis - Added missing standup target physical walls, and adjusted the target masses and material and to improve ball hit response.
' 1.1.4 - DGrimmReaper - Updated VR BG's with Hauntfreaks updated assets
' 1.1.5 - bthlonewolf - added reflections/settings, migrated settings to F12 Tweak.
' 1.1.6 - mcarter78 - fix flippers array in InitPolarity
' 1.1.7 - bthlonewolf - fixed settings syncing between dmd/f12, moved brightness to previous stored value, separate from day/night
' 1.1.8 - bthlonewolf - tweaked upperpf reflections
' 1.1.9 - iaakki - tweaked left sling, right orbit control points
' 1.1.10 - bthlonewolf - removed savehs call from addcredits to table exit for lag.
' 1.1.11 - bthlonewolf - disable Action button during PlayerMode transitions; added constraint to lannisterhit to ensure 1..5 range; added direwolf flag to BattleState
'    - apophis - adjust STMass to 0.1, increase dragon kickback power with some variability
' 1.1.12 - bthlonewolf - adjusted left ramp height/geo to fix ball too low (thank you Gravy on VPU); disabled bulb reflections on balls on upper pf.
' 1.1.13 - bthlonewolf - added invis wall above left drop targets; fixed playermode bug when draining
' 1.1.14 - iaakki - Fixed the insert ball reflections, tuned GI ball reflections, tuned insert colors, added dummy lights on upper PF and plungerlane to block ball reflections
' 1.1.15 - Lumigado - tweaked light colours further to closer match the real game, added a yellowgreen colour constant for Blackwater effects to match the real game
' 1.1.16 - Lumigado - tweaked match animation, main screen with the combo multipliers, added a display font for the combo multipliers
' 1.1.17 - bthlonewolf - mystery 'house lit' fix; update DMD when add players with BIP; update dmd with Targaryen Freeze ability; return bwmb house sigils to normal at end of intro
' 1.1.18 - bthlonewolf - VR overrides cabmode; added AwardMysteryHouse method; mystery house randomly selected; direwolf improved for multihouse battles; slight bwmb intro tweaks (ref: https://youtu.be/XOfVLODSMUA?si=6ZcVcNcptSUWqXyJ&t=277)
' 1.1.19 - bthlonewolf - Fixed BWMB sequence bug introduced in 1.1.11
' 1.1.20 - bthlonewolf - fixed VR / cabmode. Again.
' 1.1.21 - bthlonewolf - fixed quickrestart, added dof for action buttons (161-167), added castle drop target (119) to DOF config
' 1.1.22 - bthlonewolf - new ram animation; fixed ball save/timers on quick restart
' 1.1.23 - bthlonewolf - tweaked ram slightly; fixed CastlesCollected not resetting between games
' 1.1.24 - bthlonewolf - tweaked ram physics; fixed wildfiremini sound in DoBatteringRamScene; added prim to sling/flipper corner; possible fix ebislit between games
' 1.1.25 - bthlonewolf - OCD ram physics tweak; added apophis' narnia ball return;
'
'On the DOF website put Exxx
'101 Left Flipper
'102 Right Flipper
'103 left slingshot
'104 right slingshot
'105 ball release
'106
'107 Center Bumper
'108 RIGHT Bumper
'109 Left bumper
'110 Rear Elevator VUK
'111 Iron Throne VUK
'112 Dragon Kicker
'119 Reset drop Targets
'120 AutoFire
'122 knocker
'123 sword
'124 battering ram
'125 Ball in Iron Throne Kicker
'126 WIC hurry-up
'127 BWMB active
'128 Castle MB active
'129 WHC MB active
'130 HOTK MB active
'131 IT MB active
'132 Wall MB active
'133 Extra Ball award
'134 Dragon Roar
'135 Dragon Wings (triggered each 'flap')
'136 Timers Inactive - primarily for debugging
'137 Ball Save Active - primarily for debugging

' Lamps
' 140 Start button
' 141 Action button white
' 142 Action button yellow
' 143  " " red
' 144  " " purple
' 145  " " orange
' 146  " " green
' 147  " " blue
'
' 151 action button blink white
' 152 action button blink yellow
'  <etc>
'
' house abilities
' 160 n/a
' 161 Stark
' 162 Baratheon
' 163 Lannister
' 164 Tyrell
' 165 Greyjoy - not implemented (no ability)
' 166 Martell
' 167 Targaryen

Option Explicit
Randomize

Const TableVersion = "1.1.25"

'***********TABLE VOLUME LEVELS *********
' [Value is from 0 to 1 where 1 is full volume.
' NOTE: you can go past 1 to amplify sounds]
Const cVolBGMusic = 0.15  ' Volume for table background music
Const cVolCallout = 0.9   ' Volume of voice callouts
Const cVolDef = 0.5    ' Volume for GoT sound effects
Const cVolSfx = 0.2    ' Volume for table "physical" Sound effects (non Fleep sounds)
Const cVolTable = 0.2   ' Used by Daphishbowl's Service menu
'*** Fleep ****
'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Dim VolumeDial ' Now set in the Options menu = 0.8
Dim BallRollVolume ' Now set in the Options menu = 0.5

'////// Cabinet Options
Dim bHaveLockbarButton ' Set in Options menu now   -Set to true if you have a lockdown bar button. Will disable the flasher on the apron
Const DMDMode = 1                   ' Use FlexDMD (currently the only option supported)
Const bUsePlungerForSternKey = False ' If true, use the plunger key for the Action/Select button.

'----- Shadow Options -----
Dim DynamicBallShadowsOn    ' Set in the Options menu
'DynamicBallShadowsOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                  '2 = flasher image shadow, but it moves like ninuzzu's

'---- Reflection Options ----
dim ReflectionStrength : ReflectionStrength = 20
dim ReflectionScope: ReflectionScope = 2

'***** VR Options *********
'Note:These options are the default and overriden by settings menu
Dim VRRoom
Dim VRRoomChoice: VRRoomChoice = 5   '0 - Desktop/FS   1 - Minimal Room    2 - Ultra Minimal Room   3 - Iron Throne, 4 - Dragon Ice Mega, 5 - Tower Throne Mega
Dim Scratches: Scratches = 1         '0 - VR Glass Scratches Off  1 - VR Glass Scratches On
Dim Topper: Topper = 1           '0 - VR Topper Off  1 - VR Topper On
Const VRTest = 0           'Test VR in Live Editor
'**************************

'/////// No User Configurable options beyond this point ////////

Dim DesktopMode: DesktopMode = table1.ShowDT

Dim Stuff ' for Hybrid Toggle

Sub ResetVRRoom
  'reset the visibility of everything to account for live changes

  StartButtonTimer.enabled = False
  for each Stuff in VR_Cab: Stuff.visible = False: next
  for each Stuff in VR_Room: Stuff.visible = False: next
  for each Stuff in VR_IronThrone: Stuff.visible = False: next
    for each Stuff in VR_DragonIce: Stuff.visible = False: next
    for each Stuff in VR_TowerThrone: Stuff.visible = False: next
  for each Stuff in VR_Topper: Stuff.visible = False : next
  GlassImpurities.visible = false
  StopVRDragon

  'disable all topper related items
  if (Topper = 0 OR VRRoom = 2) then
    VRTopperTimer.enabled = False
    VRTopperLights.enabled = False
    VRtmrDragonWings.Enabled = False
    VRTopperLightTimer.enabled = False
    SmokeTimer.enabled = False
  Else
    'VRTopperLights - used in game, not attract mode
    if bAttractMode Then
      'if in attract mode, enable the topper attract light sequence
      VRTopperTimer.enabled = true
      VRTopperLights.enabled = false
    Else
      VRTopperTimer.enabled = false
      VRTopperLights.enabled = true
    end if

    if (VRRoom = 1 or VRRoom = 3 or VRRoom = 4 or VRRoom = 5) then
      StartVRDragon
      SmokeTimer.enabled = true
    end if
  end If

  'enable dragon smoke if in room 3
  if VRRoom = 3 then
    SmokeTimerDragon.enabled = True
    Smoke3.Visible = True
    Smoke4.Visible = True
  else
    SmokeTimerDragon.enabled = False
    Smoke3.Visible = False
    Smoke4.Visible = False
  end if

end Sub

Sub VRRoomInit
  debug.print "Init VR Room"

  If RenderingMode = 2 or VRTest = 1 Then VRRoom = VRRoomChoice Else VRRoom = 0

  'VRRoomInit can be called from tweak menu, so exit if not in VR room
  if VRRoom = 0 then
    debug.print "Exiting VR setup, not in VR mode"
    exit Sub
  end if

  debug.print "Topper enabled: " & Topper

  ResetVRRoom

    If VRRoom = 1 Then
        StartButtonTimer.enabled = true
        VRPlunger2.enabled = true ' starts the timer needed for VR Plunger animation
        for each Stuff in VR_Cab: Stuff.visible = true: next
        for each Stuff in VR_Room: Stuff.visible = true: next
        for each Stuff in VR_IronThrone: Stuff.visible = False: next
        for each Stuff in VR_DragonIce: Stuff.visible = False: next
        for each Stuff in VR_TowerThrone: Stuff.visible = False: next
        for each Stuff in VR_Topper: Stuff.visible = (Topper=1): next
        If Scratches = 1 then GlassImpurities.visible = true
    End if

    If VRRoom = 2 Then
        for each Stuff in VR_UltraMin: Stuff.visible = true: next
        'PinCab_Rails.visible = true
    SideRailMapLeft.visible = true
    SideRailMapRight.visible = true
        for each Stuff in VR_DragonIce: Stuff.visible = False: next
        for each Stuff in VR_TowerThrone: Stuff.visible = False: next
        for each Stuff in VR_Room: Stuff.visible = false: next
        for each Stuff in VR_IronThrone: Stuff.visible = false: next
    for each Stuff in VR_Topper: Stuff.visible = false : next
    End if

    If VRRoom = 3 Then
        StartButtonTimer.enabled = true
        VRPlunger2.enabled = true ' starts the timer needed for VR Plunger animation
        for each Stuff in VR_Cab: Stuff.visible = true: next
        for each Stuff in VR_Room: Stuff.visible = false: next
        for each Stuff in VR_IronThrone: Stuff.visible = true: next
        for each Stuff in VR_Topper: Stuff.visible = (Topper=1): next
        If Scratches = 1 then GlassImpurities.visible = true
        VR_Outlet.visible = False
        VR_CordPlug.visible = False
        VR_Cord.visible = False
    End if

    If VRRoom = 4 Then
        StartButtonTimer.enabled = true
        VRPlunger2.enabled = true ' starts the timer needed for VR Plunger animation
        for each Stuff in VR_Cab: Stuff.visible = true: next
        for each Stuff in VR_Room: Stuff.visible = false: next
        for each Stuff in VR_IronThrone: Stuff.visible = False: next
        for each Stuff in VR_DragonIce: Stuff.visible = True: next
        for each Stuff in VR_TowerThrone: Stuff.visible = False: next
        for each Stuff in VR_Topper: Stuff.visible = (Topper=1): next
        If Scratches = 1 then GlassImpurities.visible = true
        VR_Outlet.visible = False
        VR_CordPlug.visible = False
        VR_Cord.visible = False
    End if

    If VRRoom = 5 Then
        StartButtonTimer.enabled = true
        VRPlunger2.enabled = true ' starts the timer needed for VR Plunger animation
        for each Stuff in VR_Cab: Stuff.visible = true: next
        for each Stuff in VR_Room: Stuff.visible = false: next
        for each Stuff in VR_IronThrone: Stuff.visible = False: next
        for each Stuff in VR_DragonIce: Stuff.visible = False: next
        for each Stuff in VR_TowerThrone: Stuff.visible = True: next
        for each Stuff in VR_Topper: Stuff.visible = (Topper=1): next
        If Scratches = 1 then GlassImpurities.visible = true
        VR_Outlet.visible = False
        VR_CordPlug.visible = False
        VR_Cord.visible = False
    End if

    ' End VR Room Init....
End Sub


Const BallSize = 50     ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1      ' standard ball mass

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
Const cGameName = "gameofthrones"  ' Used by B2S for DOF
Const myGameName = "gameofthrones" ' Used for FlexDMD, High score storage, etc
Const myVersion = "1.1.25"
Const MaxPlayers = 4          ' from 1 to 4 - don't change
Const MaxMultiplier = 5       ' limit playfield multiplier
Const MaxBonusMultiplier = 20 'limit Bonus multiplier
Dim BallsPerGame
Const MaxMultiballs = 6       ' max number of balls during multiballs
Const UsingROM = False

'*****************************************************************************************************
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

Const   UltraDMD_Animation_None = 14
Const   FDsep = ","     ' The latest, not yet published, version has switched to "|"
Dim FlexDMD
'********* End FlexDMD **************


' Define Global Variables that aren't game-specific
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim PuPlayer                ' Not currently used
Dim plungerIM 'used mostly as an autofire plunger
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim Score(4)
Dim HighScore(16)
Dim HighScoreName(16)
Dim ReplayScore
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)     ' I don't think we need an array here - EBs can't be carried over
Dim ReplayAwarded(4)
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim BallsOnPlayfield        ' Active balls on playfield, including real locked ones
Dim RealBallsInLock     ' These are actually locked on the table
Dim BallsInLock             ' Number of balls the current player has locked
Dim BallSearchCnt
Dim LastSwitchHit
Dim LightSaveState(100,4)   ' Array for saving state of lights. We save state, colour, fade to restore after Sequences (sequences only save state)
Dim LightNames(300)         ' Saves the mapping of light names (li<xxx>) to index into LightSaveState
Dim TotalPlayfieldLights

' flags
Dim bMultiBallMode
Dim bAutoPlunger
Dim bAutoPlunged
Dim bJustPlunged
Dim bBallSaved
Dim bInstantInfo
Dim bAttractMode
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bTableReady
Dim bUseFlexDMD
Dim bUsePUPDMD
Dim bPupStarted
Dim bBonusHeld
Dim bJustStarted            ' Not sure what this is used for
Dim bShowMatch
Dim bSkillShotReady
Dim bMysteryAwardActive     ' Table specific - used by Mystery Award for flipper control
Dim WHCFlipperFrozen        ' table-specific - used to "freeze" the left flipper during Winter-Has-Come


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

  MovableHelper
  TargetMovableHelper

  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
'    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
'    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
    LeftUFlipperTop.RotZ = LeftUFlipper.CurrentAngle
    RightUFlipperTop.RotZ = RightUFlipper.CurrentAngle
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

' Hack to return Narnia ball back in play
Sub Narnia_Timer
    Dim b, gBOT
  gBOT=getballs
  For b = 0 to UBound(gBOT)
    if gBOT(b).z < -500 Then
      'msgbox "Ball " &b& " in Narnia X: " & gBOT(b).x &" Y: "&gBOT(b).y & " Z: "&gBOT(b).z
      debug.print "Move narnia ball ("& gBOT(b).x &" Y: "&gBOT(b).y & " Z: "&gBOT(b).z&") to left orbit"
      gBOT(b).VelX = -4 : gBOT(b).VelY = 4 : gBOT(b).VelZ = 0
      gBOT(b).x = 400
      gBOT(b).y = 160
      gBOT(b).z = 26
    end if
  next
end sub

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    LoadEM
    Dim i
    Randomize

  vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)

  bTableReady=False
  bUseFlexDMD=False
  bUsePUPDMD=False
  bPupStarted=False
  if DMDMode = 1 then
    bUseFlexDMD= True
    set PuPlayer = New PinupNULL
  elseif DMDMode = 2 Then
    bUsePUPDMD = True
  Else
    set PuPlayer = New PinupNULL
  End if

  if b2son then
    'Controller.PupHide = 1
    Controller.B2SSetData 1,1
    Controller.B2SSetData 2,1
    Controller.B2SSetData 3,1
    Controller.B2SSetData 4,1
    Controller.B2SSetData 5,1
    Controller.B2SSetData 6,1
    Controller.B2SSetData 7,1
    Controller.B2SSetData 8,1
  End if

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFX("fx_autoplunger", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' Creat the array of light name indexes
    SavePlayfieldLightNames

    ' initalise the DMD display
    DMD_Init

    ' Load table color
    LoadLut

    'load saved values, highscore, names, jackpot
    Options_Load
  'VRRoomInit
    DMDSettingsInit() ' Init the Service Menu
    ReplayScore = DMDStd(kDMDStd_AutoReplayStart)
    Loadhs

    ' initialse any other flags
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bGameInPlay = False
  'bGameInPlayHidden = False
  bShowMatch = False
  'bCreatedBall = False
    bAutoPlunger = False
  'bAutoPlunged = False
    BallsOnPlayfield = 0
  RealBallsInLock=0
    BallsInLock = 0
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bInstantInfo = False
    bMysteryAwardActive = False


  bTableReady = True
  ' set any lights for the attract mode
    GiOff
    TurnOffPlayfieldLights
    SetLockbarLight
    StartAttractMode

    ' Start the RealTime timer
    RealTime.Enabled = 1

  'Desktop mode Stuff
  EMReel1.visible = DesktopMode
  EMReel2.visible = DesktopMode
  liDTaction.visible = DesktopMode

End Sub

Private Function BigMod(Value1, Value2)
    BigMod = Value1 - (Int(Value1 / Value2) * Value2)
End Function

Sub Table1_Exit()
  On Error Resume Next

  'try to save all HS/stats
  savehs

  'changed flexdmd shutdown to work in GL version better
  'If TypeName(FlexDMD) = "Object" And Not IsEmpty(FlexDMD) Then
  'If Not FlexDMD is Nothing Then
  If IsObject(FlexDMD) Then
    FlexDMD.Show = False
    FlexDMD.Run = False
    FlexDMD = NULL
  End If

    If B2SOn Then Controller.Stop
  On Error Goto 0
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    'VR Specific...
    If keycode = LeftMagnaSave and VRRoom > 0 then VRCabArt =  VRCabArt + 1: ArtSwap
    If keycode = LeftFlipperKey and VRRoom > 0 then VR_FlipperLeft.x = VR_FlipperLeft.x + 5
    If keycode = RightFlipperKey and VRRoom > 0 then VR_FlipRight.x = VR_FlipRight.x - 5
    If keycode = StartGameKey and VRRoom > 0 then
        VR_StartButtOuter.y = VR_StartButtOuter.y - 4
        StartButtonTimer.enabled = false
        VR_StartButtOuter.blenddisablelighting = 0
        VR_TourneyButt.blenddisablelighting = 0
    end if
  If keycode = RightMagnaSave or keycode = LockBarKey or _
    (keycode = PlungerKey and bUsePlungerForSternKey) Then

    for each Stuff in LockdownButton_BL: Stuff.z = Stuff.z - 5: next
    for each Stuff in LockdownButton_BM: Stuff.z = Stuff.z - 5: next

  end if
    ' end VR Specific

    If keycode = LeftTiltKey Then Nudge 90, 2:SoundNudgeLeft()
    If keycode = RightTiltKey Then Nudge 270, 2:SoundNudgeRight()
    If keycode = CenterTiltKey Then Nudge 0, 2:SoundNudgeCenter()

    If bInOptions Then
    Options_KeyDown keycode
    Exit Sub
  End If

    If keycode = LeftMagnaSave Then
    If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
    ElseIf keycode = RightMagnaSave Then
    If bOptionsMagna Then
            Options_Open()
            Exit Sub
        Else
            If Not bGameInPlay Then bOptionsMagna = True
        End if
    End If

    If Keycode = AddCreditKey or Keycode = AddCreditKey2 Then
        Credits = Credits + 1
        if DMDStd(kDMDStd_FreePlay) = False Then DOF 140, DOFOn
        GameAddCredit
        'savehs  ' removing this call due to lag
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
        SoundPlungerPull()
        If VRroom > 0 then  'VR Specific
            VRPlunger.Enabled = True
            VRPlunger2.Enabled = False
        End If
    End If
    If hsbModeActive > 0 Then
        if hsbModeActive=1 Then EnterHighScoreKey(keycode)
        Exit Sub
    End If

    If bMysteryAwardActive Then
        UpdateMysteryAward(keycode)
        Exit Sub
    End If

    StartServiceMenu keycode

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

        If bMadnessMB = 1 Then Exit Sub

        If keycode = LeftTiltKey Then CheckTilt 'only check the tilt during game
        If keycode = RightTiltKey Then CheckTilt
        If keycode = CenterTiltKey Then CheckTilt
        If keycode = MechanicalTilt Then SoundNudgeCenter:CheckTilt

        If keycode = LeftFlipperKey Then
            If WHCFlipperFrozen = False Then FlipperActivate LeftFlipper, LFPress : FlipperActivate LeftUFlipper, ULFPress : SolLFlipper 1
            InstantInfoTimer.Enabled = True
            RotateLaneLights 1
            If InstantInfoTimer.UserValue = 0 Then
                InstantInfoTimer.UserValue = keycode ' Record which key triggered the InstantInfo
            ElseIf InstantInfoTimer.UserValue <> keycode And bInstantInfo Then
                InfoPage = InfoPage + 1:InstantInfo
            End If
        ElseIf keycode = RightFlipperKey Then
            FlipperActivate RightFlipper, RFPress :  FlipperActivate RightUFlipper, URFPress : SolRFlipper 1
            InstantInfoTimer.Enabled = True
            RotateLaneLights 0
            If InstantInfoTimer.UserValue = 0 Then
                InstantInfoTimer.UserValue = keycode ' Record which key triggered the InstantInfo
            ElseIf InstantInfoTimer.UserValue <> keycode And bInstantInfo Then
                InfoPage = InfoPage + 1:InstantInfo
            End if
        End If
        '  Action Button, Start Mode, fire ball
    If keycode = RightMagnaSave or keycode = LockBarKey or _
      (keycode = PlungerKey and bUsePlungerForSternKey) Then

      if bAutoPlunger=False and bBallInPlungerLane = True then  ' Auto fire ball with stern key
        plungerIM.Strength = 60
        'plungerIM.InitImpulseP swplunger, 60, 0    ' Change impulse power while we are here
        PlungerIM.AutoFire
        DOF 112, DOFPulse
        plungerIM.Strength = 45
        'plungerIM.InitImpulseP swplunger, 45, 1.1
      Else
        CheckActionButton
      end if
    End if

        If CheckLocalKeydown(keycode) Then Exit Sub

        If keycode = StartGameKey Then
            If LFPress=1 and (DMDStd(kDMDStd_LeftStartReset)=1 Or (DMDStd(kDMDStd_LeftStartReset)=2 And DMDStd(kDMDStd_FreePlay))) Then  ' Quick Restart
                DoQuickRestart
                Exit Sub
            End If
            If((PlayersPlayingGame <MaxPlayers)AND(bOnTheFirstBall = True))Then

                If(DMDStd(kDMDStd_FreePlay))Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
          AddNewPlayerWhileBIP ' 1.1.17
                    DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, False, ""
                Else
                    If(Credits> 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
            AddNewPlayerWhileBIP ' 1.1.17
                        DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, False, ""
                        If Credits <1 And DMDStd(kDMDStd_FreePlay) = False Then DOF 140, DOFOff
                    Else
                        ' Not Enough Credits to start a game.
                        DisplayDMDText "CREDITS 0", "INSERT COINS", 1000
                    End If
                End If
            End If
        End If
    Else ' If (GameInPlay)

        If keycode = StartGameKey or keycode = LockBarKey Then
            If(DMDStd(kDMDStd_FreePlay))Then
                If(BallsOnPlayfield = 0)Then
                    ResetForNewGame()
                End If
            Else
                If(Credits> 0)Then
                    If(BallsOnPlayfield = 0)Then
                        Credits = Credits - 1
                        If Credits <1 And DMDStd(kDMDStd_FreePlay) = False Then DOF 140, DOFOff
                        ResetForNewGame()
                    End If
                Else
                    ' Not Enough Credits to start a game.
                    ' This is Table-specific:
                    PlaySoundVol "gotfx-nocredits",VolDef
                    If bAttractMode Then
                        tmrAttractModeScene.UserValue = 5
                        tmrAttractModeScene_Timer
                    Else
                        DisplayDMDText "CREDITS 0","INSERT COINS",1000
                    End If
                End If
            End If
        ElseIf bAttractMode And (keycode = LeftFlipperKey Or keycode = RightFlipperKey)  Then
            doFlipperSpeech keycode
        End If
    End If ' If (GameInPlay)
End Sub

Sub Table1_KeyUp(ByVal keycode)

  'VR Specific....
  If keycode = StartGameKey and VRRoom > 0 then VR_StartButtOuter.y = VR_StartButtOuter.y + 4
  If keycode = LeftFlipperKey and VRRoom > 0 then VR_FlipperLeft.x = VR_FlipperLeft.x - 5
  If keycode = RightFlipperKey and VRRoom > 0 then VR_FlipRight.x = VR_FlipRight.x + 5
  If keycode = RightMagnaSave or keycode = LockBarKey or _
    (keycode = PlungerKey and bUsePlungerForSternKey) Then

    for each Stuff in LockdownButton_BL: Stuff.z = Stuff.z + 5: next
    for each Stuff in LockdownButton_BM: Stuff.z = Stuff.z + 5: next

  end if
  ' end VR Specific

    If keycode = LeftMagnaSave And Not bInOptions Then bOptionsMagna = False
    If keycode = RightMagnaSave And Not bInOptions Then bOptionsMagna = False

    If keycode = PlungerKey Then
        Plunger.Fire
    If VRroom > 0 then  'VR Specific
            VRPlunger.Enabled = False
            VRPlunger2.Enabled = True
            VR_Plunger.Y = -117.5515
    End if
        If bBallInPlungerLane Then SoundPlungerReleaseBall() Else SoundPlungerReleaseNoBall()
    End If

    If hsbModeActive Then
        Exit Sub
    End If

    ' Table specific

    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then
            If WHCFlipperFrozen = False Then FlipperDeactivate LeftFlipper, LFPress : FlipperDeactivate LeftUFlipper, ULFPress : SolLFlipper 0
            If InstantInfoTimer.UserValue = keycode Then
                InstantInfoTimer.UserValue = 0
                InstantInfoTimer.Enabled = False
                InstantInfoTimer.Interval = 10000
                If bInstantInfo Then
                    tmrDMDUpdate.Enabled = true
                    DMDFlush : AddScore 0
                    bInstantInfo = False
                End If
            End If
        ElseIf keycode = RightFlipperKey Then
            FlipperDeactivate RightFlipper, RFPress : FlipperDeactivate RightUFlipper, URFPress : SolRFlipper 0
            If InstantInfoTimer.UserValue = keycode Then
                InstantInfoTimer.UserValue = 0
                InstantInfoTimer.Enabled = False
                InstantInfoTimer.Interval = 10000
                If bInstantInfo Then
                    tmrDMDUpdate.Enabled = true
                    DMDFlush : AddScore 0
                    bInstantInfo = False
                End If
            End If
        End If
    End If
End Sub

Dim InfoPage
Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    If bGameInPlay And Tilted=False And hsbModeActive = False And bMultiBallMode = False Then
        bInstantInfo = True
        tmrDMDUpdate.Enabled = False
        DMDFlush
        InfoPage = 0
        InstantInfo
    End If
End Sub

'********************
'     Flippers
'********************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
        LF.Fire
        ULF.Fire
        DOF 101, DOFOn
        If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
            RandomSoundReflipUpLeft LeftFlipper
            RandomSoundReflipUpLeft LeftUFlipper
        Else
            SoundFlipperUpAttackLeft LeftFlipper
            RandomSoundFlipperUpLeft LeftFlipper
            SoundFlipperUpAttackLeft LeftUFlipper
            RandomSoundFlipperUpLeft LeftUFlipper
        End If
    Else
        LeftFlipper.RotateToStart
        LeftUFlipper.RotateToStart
        DOF 101, DOFOff
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
            RandomSoundFlipperDownLeft LeftUFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        RF.Fire
        URF.Fire
        DOF 102, DOFOn
        If rightflipper.currentangle > leftflipper.endangle - ReflipAngle Then
            RandomSoundReflipUpRight RightFlipper
            RandomSoundReflipUpRight RightUFlipper
        Else
            SoundFlipperUpAttackRight RightFlipper
            RandomSoundFlipperUpRight RightFlipper
            SoundFlipperUpAttackRight RightUFlipper
            RandomSoundFlipperUpRight RightUFlipper
        End If
    Else
        RightFlipper.RotateToStart
        RightUFlipper.RotateToStart
        DOF 102, DOFOff
        If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
            RandomSoundFlipperDownRight RightUFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
  LF.ReProcessBalls ActiveBall
    LeftFlipperCollide parm
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
End Sub

Sub RightFlipper_Collide(parm)
  RF.ReProcessBalls ActiveBall
    RightFlipperCollide parm
    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
End Sub

Sub LeftUFlipper_Collide(parm)
  ULF.ReProcessBalls ActiveBall
    LeftFlipperCollide parm
End Sub

Sub RightUFlipper_Collide(parm)
  URF.ReProcessBalls ActiveBall
    RightFlipperCollide parm
End Sub

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                  'Called when table is nudged
    Tilt = Tilt + TiltSensitivity              'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity)AND(Tilt <15)Then 'show a warning
        DMD "_", CL(1, "CAREFUL"), "_", eNone, eBlinkFast, eNone, 1500, True, "fx-tiltwarning"
        ' Table specific:
        PlaySoundVol "say-tiltwarning"&RndNbr(12),VolCallout
    End if
    If Tilt> 15 Then 'If more than 15 then TILT the table
        Tilted = True
        'display Tilt
        DMDFlush
        DMD "", "TILT", "", eNone, eNone, eNone, 5000, True, "fx-tilt"
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
        If Song <> "" Then StopSound Song : Song=""
        'turn off GI and turn off all the lights
        GiOff
        LightSeqTilt.Play SeqAllOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        LeftUFlipper.RotateToStart
        RightUFlipper.RotateToStart
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
    If(BallsOnPlayfield-RealBallsInLock = 0)Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        vpmtimer.Addtimer 2000, "EndOfBall() '"
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

Dim bQuickRestartRequested : bQuickRestartRequested = False
Dim bQuickRestartRunning : bQuickRestartRunning = False
Sub DoQuickRestart

  'if ball is currently draining, do not process quick restart
  if bBallIsDraining Then
    debug.print "Skipping DoQuickRestart while ball is draining"
    bQuickRestartRequested = True
    exit sub
  end If

  if bQuickRestartRunning Then
    debug.print "Skipping DoQuickRestart; already running"
    bQuickRestartRequested = False ' shouldn't be needed
    exit sub
  end if

  bQuickRestartRequested = False
  bQuickRestartRunning = True

  debug.print "QuickResart Beginning"

  'try to solve ball saver issue where ball saver is not reset
  TimerFlags(tmrBallSave) = TimerFlags(tmrBallSave) AND 254
  TimerFlags(tmrBallSaveSpeedUp) = TimerFlags(tmrBallSaveSpeedUp) AND 254
  TimerFlags(tmrBallSaverGrace) = TimerFlags(tmrBallSaverGrace) AND 254

  bBallSaverActive = False
  bBallSaverReady = False

    DisableTable true
    Tilted=true
    StopGameTimers
    if bBallInPlungerLane Then PlungerIM.AutoFire
    bGameInPLay = False
    bShowMatch = False
    tmrBallSearch.Enabled = False
    RealBallsInLock=0
    BallsInLock = 0
    EndOfGame
    SetLockbarLight
    LFPress=0 : RFPress=0

    vpmTimer.AddTimer 1000,"CheckQuickRestartFinished '"
End Sub



Sub CheckQuickRestartFinished
    if BallsOnPlayfield = 0 Then
    DisableTable False
    GiOff
    bQuickRestartRunning = False
    debug.print "CheckQuickRestartFinished completed"
    Exit Sub
  end If
    vpmTimer.AddTimer 1000,"CheckQuickRestartFinished '"
End Sub

Sub PlaySoundVol(soundname, Volume)
  PlaySound soundname, 1, Volume
End Sub

' Play an already playing sound (starts anew if not playing). Restart=whether to restart the sound. Presumably 0 = just let it keep playing
Sub PlayExistingSoundVol(soundname, Volume, Restart)
  PlaySound soundname, 1, Volume, 0, 0, 0, 1, Restart
End Sub

Sub PlaySoundLoopVol(soundname, Volume)
  PlaySound soundname, -1, Volume
End Sub

'********************
' Music as wav sounds
'********************

Dim Song
Song = ""

Sub ThemeSong
  PlaySong("Song-1")
  SongNum=1
End Sub

Sub RotateSong()
'debug.print "Rotate " & SongNum
  PlaySong "Song-" & SongNum
  SongNum=SongNum+1
  if (SongNum>=4) then SongNum=1
End Sub


dim bPlayPaused
bPlayPaused = False
Sub PlaySong(name)
'debug.print "PlaySong " & name & " " & song
  dim PlayLength
  if bUsePUPDMD then      ' Use Pup if we have it so we can pause the music
'   PlaySongPup(name)
    exit sub
  End If
  StopSound "m_wait"
  StopSound Song  ' Stop the old song
  if name <> "" then Song = name
  PlayLength = -1
  If Song = "m_end" Then PlayLength = 0
  bPlayPaused=False
  PlaySound Song, PlayLength, VolBGMusic 'this last number is the volume, from 0 to 1
End Sub


Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function


'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx7" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
' * Count from 0 up, with at least as many objects each as there can be balls, including locked balls.  You'll get an "eval" warning if tnob is higher
' * Warning:  If merging with another system (JP's ballrolling), you may need to check tnob math and add an extra BallShadowA# flasher (out of range error)
' Ensure you have a timer with a -1 interval that is always running
' Set plastic ramps DB to *less* than the ambient shadows (-10000) if you want to see the pf shadow through the ramp

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
' It's recommended that you be selective in which lights go in this collection, as there are limitations:
' 1. The shadows can "pass through" solid objects and other light sources, so be mindful of where the lights would actually able to cast shadows
' 2. If there are more than two equidistant sources, the shadows can suddenly switch on and off, so places like top and bottom lanes need attention
' 3. At this time the shadows get the light on/off from tracking gilvl, so if you have lights you want shadows for that are on at different times you will need to either:
' a) remove this restriction (shadows think lights are always On)
' b) come up with a custom solution (see TZ example in script)
' After confirming the shadows work in general, use ball control to move around and look for any weird behavior

'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
' Example timer sub:

'Sub FrameTimer_Timer()
' If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
'End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
'Const lob = 0  'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top for easy access by players
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 '2 = flasher image shadow, but it moves like ninuzzu's

' *** This segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' *** Change gBOT to BOT if using existing getballs code
' *** Includes lines commonly found there, for reference:
' ' stop the sound of deleted balls
' For b = UBound(gBOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
'   ...rolling(b) = False
'   ...StopSound("BallRoll_" & b)
' Next
'
' ...rolling and drop sounds...

'   If DropCount(b) < 5 Then
'     DropCount(b) = DropCount(b) + 1
'   End If
'
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If gBOT(b).Z > 30 Then
'       BallShadowA(b).height=gBOT(b).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
'     Else
'       BallShadowA(b).height=gBOT(b).z - BallSize/2 + 5
'     End If
'     BallShadowA(b).Y = gBOT(b).Y + Ballsize/5 + offsetY
'     BallShadowA(b).X = gBOT(b).X + offsetX
'     BallShadowA(b).visible = 1
'   End If

' *** Required Functions, enable these if they are not already present elswhere in your table
' Function max(a,b)
'   if a > b then
'     max = a
'   Else
'     max = b
'   end if
' end Function

'Function Distance(ax,ay,bx,by)
' Distance = SQR((ax - bx)^2 + (ay - by)^2)
'End Function

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

'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******

' *** These define the appearance of shadows in your table  ***

'Ambient (Room light source)
Const AmbientBSFactor     = 0.9 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const offsetX       = 0   'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY       = 0   'Offset y position under ball  (for example 5,5 if the light is in the back left corner)
'Dynamic (Table light sources)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source

' ***                           ***

Const falloff         = 150   'Max distance to light sources, can be changed if you have a reason
Dim GIfalloff : GIfalloff   = 250
Const PLOffset = 0.05
Dim PLGain: PLGain = (1-PLOffset)/(1260-2100)

' *** Trim or extend these to *match* the number of balls/primitives/flashers on the table!
dim objrtx1(9), objrtx2(9)
dim objBallShadow(9)
Dim OnPF(9)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9)
Dim DSSources(30), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

Dim ClearSurface:ClearSurface = True    'Variable for hiding flasher shadow on wire and clear plastic ramps
                  'Intention is to set this either globally or in a similar manner to RampRolling sounds

'Initialization
DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob - 1               'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = 1 + iii/1000 + 0.21      'Separate z for layering without clipping
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = 1 + iii/1000 + 0.22
    objrtx2(iii).visible = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii/1000 + 0.24
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
'   If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1  'Adapted for TZ with GI left / GI right
    iii = iii + 1
  Next
  numberofsources = iii
end sub


Sub BallOnPlayfieldNow(yeh, num)    'Only update certain things once, save some cycles
  If yeh Then
    OnPF(num) = True
'   debug.print "Back on PF"
    UpdateMaterial objBallShadow(num).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(num).size_x = 5
    objBallShadow(num).size_y = 4.5
    objBallShadow(num).visible = 1
    BallShadowA(num).visible = 0
  Else
    OnPF(num) = False
'   debug.print "Leaving PF"
    If Not ClearSurface Then
      BallShadowA(num).visible = 1
      objBallShadow(num).visible = 0
    Else
      objBallShadow(num).visible = 1
    End If
  End If
End Sub

Sub DynamicBSUpdate
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
  Dim gBOT: gBOT=getballs 'Uncomment if you're deleting balls - Don't do it! #SaveTheBalls

  'Hide shadow of deleted balls
  For s = UBound(gBOT) + 1 to tnob - 1
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(gBOT) < lob Then Exit Sub   'No balls in play, exit

'The Magic happens now
  Dim l
  Dim d_w
  Dim b_base, b_r, b_g, b_b
  b_base = 70 * 0.01 * Playfield_LM_GI1.Opacity + 55 * (LightLevel / 100) + 10
  For s = lob to UBound(gBOT)

' *** Compute ball lighting from GI and ambient lighting
    d_w = GIfalloff
    For Each l in aGiLights
      LSd = Distance(gBOT(s).x, gBOT(s).y, l.x, l.y) 'Calculating the Linear distance to the Source
      If LSd < d_w then d_w = LSd
    Next
    d_w = b_base + 70 * (1 - d_w / GIfalloff) * 0.01 * Playfield_LM_GI1.Opacity
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,870,2100,870,1260,930,1260,930,2100) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-2100))
    End If
    ' Assign color
    b_r = Int(d_w)
    b_g = Int(d_w)
    b_b = Int(d_w)
    If b_r > 255 Then b_r = 255
    If b_g > 255 Then b_g = 255
    If b_b > 255 Then b_b = 255
    gBot(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
    'debug.print "--- ball.color level="&b_r

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
            If gBOT(s).Z > 190 and gBOT(s).Z < 200 and InRect(gBOT(s).x,gBOT(s).y,32,90,32,785,500,785,500,90 ) Then ' On upper PF
                objBallShadow(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX + 5
                objBallShadow(s).Y = gBOT(s).Y + offsetY
                objBallShadow(s).Z = gBOT(s).Z + s/1000 - 24.9
                objBallShadow(s).size_x = 5
                objBallShadow(s).size_y = 5
                UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
      ElseIf gBOT(s).Z > 30 And Not InRect(gBOT(s).x,gBOT(s).y,32,90,32,785,500,785,500,90 ) Then               'The flasher follows the ball up ramps while the primitive is on the pf
        If OnPF(s) Then BallOnPlayfieldNow False, s   'One-time update

        If Not ClearSurface Then              'Don't show this shadow on plastic or wire ramps (table-wide variable, for now)
          BallShadowA(s).X = gBOT(s).X + offsetX
          BallShadowA(s).Y = gBOT(s).Y + BallSize/5
          BallShadowA(s).height=gBOT(s).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
        Else
          If gBOT(s).X < tablewidth/2 Then
            objBallShadow(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX + 5
          Else
            objBallShadow(s).X = ((gBOT(s).X) + (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX - 5
          End If
          objBallShadow(s).Y = gBOT(s).Y + BallSize/10 + offsetY
          objBallShadow(s).size_x = 5 * ((gBOT(s).Z+BallSize)/80)     'Shadow gets larger and more diffuse as it moves up
          objBallShadow(s).size_y = 4.5 * ((gBOT(s).Z+BallSize)/80)
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor*(30/(gBOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
        End If

      Elseif gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then  'On pf, primitive only
        If Not OnPF(s) Then BallOnPlayfieldNow True, s

        If gBOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX + 5
        Else
          objBallShadow(s).X = ((gBOT(s).X) + (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX - 5
        End If
        objBallShadow(s).Y = gBOT(s).Y + offsetY
        objBallShadow(s).Z = s/1000 + 0.04    'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

      Else                        'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If gBOT(s).Z > 30 Then              'In a ramp
        If Not ClearSurface Then              'Don't show this shadow on plastic or wire ramps (table-wide variable, for now)
          BallShadowA(s).X = gBOT(s).X + offsetX
          BallShadowA(s).Y = gBOT(s).Y + BallSize/5
          BallShadowA(s).height=gBOT(s).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
        Else
          BallShadowA(s).X = gBOT(s).X + offsetX
          BallShadowA(s).Y = gBOT(s).Y + offsetY
        End If
      Elseif gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If gBOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX + 5
        Else
          BallShadowA(s).X = ((gBOT(s).X) + (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX - 5
        End If
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height=0.1
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If gBOT(s).Z < 30 And gBOT(s).X < 850 Then  'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
        dist1 = falloff:
        dist2 = falloff
        For iii = 0 to numberofsources - 1 ' Search the 2 nearest influencing lights
          LSd = Distance(gBOT(s).x, gBOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
          If LSd < falloff then 'And gilvl > 0 Then
'         If LSd < dist2 And ((DSGISide(iii) = 0 And Lampz.State(100)>0) Or (DSGISide(iii) = 1 And Lampz.State(104)>0)) Then  'Adapted for TZ with GI left / GI right
            dist2 = dist1
            dist1 = LSd
            src2 = src1
            src1 = iii
          End If
        Next
        ShadowOpacity1 = 0
        If dist1 < falloff Then
          objrtx1(s).visible = 1 : objrtx1(s).X = gBOT(s).X : objrtx1(s).Y = gBOT(s).Y
          'objrtx1(s).Z = gBOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity1 = 1 - dist1 / falloff
          objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
          UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx1(s).visible = 0
        End If
        ShadowOpacity2 = 0
        If dist2 < falloff Then
          objrtx2(s).visible = 1 : objrtx2(s).X = gBOT(s).X : objrtx2(s).Y = gBOT(s).Y + offsetY
          'objrtx2(s).Z = gBOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity2 = 1 - dist2 / falloff
          objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
          UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx2(s).visible = 0
        End If
        If AmbientBallShadowOn = 1 Then
          'Fades the ambient shadow (primitive only) when it's close to a light
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor*(1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
        End If
      Else 'Hide dynamic shadows everywhere else, just in case
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub
'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************



'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 9 ' total number of balls
Const lob = 0

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
  Dim b, BOT
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
        rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND (BOT(b).z < 30 Or (BOT(b).Z > 190 and BOT(b).Z < 200 and InRect(BOT(b).x,BOT(b).y,32,90,32,785,500,785,500,90 ) ) )Then
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
        BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height=0.1
      End If
      BallShadowA(b).Y = BOT(b).Y + offsetY
      BallShadowA(b).X = BOT(b).X + offsetX
      BallShadowA(b).visible = 1
    End If
  Next
End Sub



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
dim RampBalls(10,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(10)

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



Sub trigRamp1_unhit
  If Activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub trigRamp2_unhit
  WireRampOff
End Sub

Sub trigRamp3_hit
  WireRampOn False
End Sub

Sub trigRamp4_hit
  WireRampOff
End Sub

Sub trigRamp5_hit
  WireRampOn False
End Sub

Sub trigRamp6_hit
  WireRampOff

  'disable then enable IT Kicker after ~<1 sec, toggle IT ramp walls
  KickerIT.enabled = False
  dim r
  for each r in ITRamps
    r.collidable = True
  Next

  vpmTimer.AddTimer 400,"EnableITKicker '"
End Sub

Sub trigRamp7_unhit
  If Activeball.vely < 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub trigRamp8_hit
  WireRampOff
End Sub

sub EnableITKicker
  KickerIT.enabled = True
  dim r
  for each r in ITRamps
    r.collidable = False
  Next
end Sub

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************


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


'********************
'     FlippersPol
'********************


dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity
dim ULF : Set ULF = New FlipperPolarity
dim URF : Set URF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF, ULF, URF)
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
  ULF.SetObjects "ULF", LeftUFlipper, TriggerULF
  URF.SetObjects "URF", RightUFlipper, TriggerURF
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
  Dim y     'Y output
  Dim L : L=1 'Line
  'find active line
  Dim ii
  For ii = 1 To UBound(xKeyFrame)
    If xInput <= xKeyFrame(ii) Then
      L = ii
      Exit For
    End If
  Next
  If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)    'catch line overrun
  'Debug.Print("L in LinearEnvelope: " & L)
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )     'Clamp lower
  If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )    'Clamp upper

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

    FlipperTricks LeftUFlipper, ULFPress, ULFCount, ULFEndAngle, ULFState
  FlipperTricks RightUFlipper, URFPress, URFCount, URFEndAngle, URFState
  FlipperNudge RightUFlipper, URFEndAngle, URFEOSNudge, LeftUFlipper, ULFEndAngle
  FlipperNudge LeftUFlipper, ULFEndAngle, ULFEOSNudge,  RightUFlipper, URFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge
Dim ULFEOSNudge, URFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b, gBOT
  '   Dim BOT
  gBOT = GetBalls

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

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

dim ULFPress, URFPress, ULFCount, URFCount
dim ULFState, URFState
dim URFEndAngle, ULFEndAngle

Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
ULFState = 1
URFState = 1

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8
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
Const EOSReturn = 0.025

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

ULFEndAngle = LeftUflipper.endangle
URFEndAngle = RightUFlipper.endangle

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

'****************************************************************************
'PHYSICS DAMPENERS
'****************************************************************************

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub

Sub zCol_Rubber_Post002  'not applying target bouncer to bottom sling posts
  RubbersD.dampen Activeball
End Sub

Sub zCol_Rubber_Post003  'not applying target bouncer to bottom sling posts
  RubbersD.dampen Activeball
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

'Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate4":End Sub
'Sub aRollovers_Hit(idx):PlaySoundAtBall "fx_sensor":End Sub
'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7

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
Sub TargetBounce_Hit(idx)
  TargetBouncer activeball, 1
End Sub




' Slingshots has been hit

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
  LS.VelocityCorrect(ActiveBall)
    'PlaySoundAt SoundFXDOF("CrispySlingLeft", 103, DOFPulse, DOFContactors), Lemk
    RandomSoundSlingshotLeft Lemk
    DOF 103, DOFPulse
    Dim sl : For each sl in LSling2_BL : sl.visible = 0 : Next
    For each sl in LSling1_BL : sl.visible = 1 : Next
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerInterval = 10
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 170
    ' add some effect to the table?
    'FlashForMs l20, 1000, 50, 0:FlashForMs l20f, 1000, 50, 0
    'FlashForMs l21, 1000, 50, 0:FlashForMs l21f, 1000, 50, 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
    Dim i : i = RndNbr(3)
    PlaySoundVol "gotfx-sling"&i,VolDef
End Sub

Sub LeftSlingShot_Timer
    Dim sl,v1,v2
    Select Case LStep
        Case 1: v1=true :  v2=false
        Case 2: v1=false : v2=true  : Lemk.RotX = 2
        Case 3: v1=false : v2=false : Lemk.RotX = -10:LeftSlingShot.TimerEnabled = False
    End Select
    For each sl in LSling2_BL : sl.visible = v2 : Next
    For each sl in LSling1_BL : sl.visible = v1 : Next
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
  RS.VelocityCorrect(ActiveBall)
    'PlaySoundAt SoundFXDOF("CrispySlingRight", 104, DOFPulse, DOFContactors),Remk
    RandomSoundSlingshotRight Remk
    DOF 104,DOFPulse
    Dim sl : For each sl in RSling2_BL : sl.visible = 0 : Next
    For each sl in RSling1_BL : sl.visible = 1 : Next
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerInterval = 10
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 170
    ' add some effect to the table?
    'FlashForMs l22, 1000, 50, 0:FlashForMs l22f, 1000, 50, 0
    'FlashForMs l23, 1000, 50, 0:FlashForMs l23f, 1000, 50, 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
    Dim i : i = RndNbr(3)
    PlaySoundVol "gotfx-sling"&i,VolDef
End Sub

Sub RightSlingShot_Timer
    Dim sl,v1,v2
    Select Case RStep
        Case 1: v1=true :  v2=false
        Case 2: v1=false : v2=true  : Remk.RotX = 2
        Case 3: v1=false : v2=false : Remk.RotX = -10:RightSlingShot.TimerEnabled = False
    End Select
    For each sl in RSling2_BL : sl.visible = v2 : Next
    For each sl in RSling1_BL : sl.visible = v1 : Next
    RStep = RStep + 1
End Sub


'************************************
' High Score support
'************************************
'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim i,x
    For i = 1 to 16
        x = LoadValue(myGameName, "HighScore"&i)
        If(x <> "")Then
            HighScore(i-1) = CDbl(x)
        Else
            if i < 6 Then HighScore(i-1) = DMDStd(kDMDStd_HighScoreGC + i - 1) Else HighScore(i-1) = (17-i)*1000000
        End If
        x = LoadValue(myGameName, "HighScore"&i&"Name")
        If(x <> "")Then HighScoreName(i-1) = x Else HighScoreName(i-1) = Chr(64+i)&Chr(64+i)&Chr(64+i) End If
    Next
    x = LoadValue(myGameName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0
    If Credits = 0 And DMDStd(kDMDStd_FreePlay) = False Then DOF 140, DOFOff Else DOF 140, DOFOn
    x = LoadValue(myGameName, "TotalGamesPlayed")
    If(x <> "")then TotalGamesPlayed = CInt(x)Else TotalGamesPlayed = 0 End If
End Sub

Sub reseths
    Dim i
    For i = 0 to 15
        HighScore(i) = (16-i)*10000000
        HighScoreName(i) = Chr(64+i)&Chr(64+i)&Chr(64+i)
    Next
    HighScore(0) = DMDStd(kDMDStd_HighScoreGC)
    HighScore(1) = DMDStd(kDMDStd_HighScore1)
    HighScore(2) = DMDStd(kDMDStd_HighScore2)
    HighScore(3) = DMDStd(kDMDStd_HighScore3)
    HighScore(4) = DMDStd(kDMDStd_HighScore4)
    savehs
End Sub

Sub Clearhs : reseths : End Sub
Sub ClearAll
    reseths
    Dim i
    For i = 0 to kDMDStd_LastStdSetting
        SaveValue myGameName, "DMDStd_"&i,""
    Next
    For i =  kDMDStd_LastStdSetting+1 to 200
        SaveValue myGameName, "DMDFet_"&i,""
    Next
End Sub

Sub Savehs
    Dim i
    For i = 1 to 16
        SaveValue myGameName, "HighScore"&i, HighScore(i-1)
        SaveValue myGameName, "HighScore"&i&"Name", HighScoreName(i-1)
    Next
    SaveValue myGameName, "Credits", Credits
    SaveValue myGameName, "TotalGamesPlayed", TotalGamesPlayed
End Sub

' ***********************************************************
'  High Score Initals Entry Functions - based on Black's code
' ***********************************************************

Dim hsbModeActive
Dim hsEnteredName
Dim hsEnteredDigits(10)
Dim hsCurrentDigit
Dim hsMaxDigits
Dim hsX
Dim hsStart, hsEnd
Dim hsValidLetters
Dim hsCurrentLetter
Dim hsLetterFlash
Dim HSscene

Sub CheckHighscore()
    Dim tmp
    tmp = GameCheckHighScore(Score(CurrentPlayer))

    Select Case tmp
        Case 0: EndOfBallComplete()
        Case 1: HighScoreEntryInit()
        Case 2,3,4,5,6
            Credits = Credits + tmp-1
            DOF 140, DOFOn
            KnockerSolenoid
            DOF 122, DOFPulse
            If tmp > 2 Then
                vpmTimer.AddTimer 500+100*(tmp-3), "HighScoreEntryInit() '"
                Do While tmp > 2
                    vpmTimer.AddTimer 200*(tmp-2),"KnockerSolenoid : DOF 122, DOFPulse '"
                    tmp = tmp - 1
                Loop
            Else
                vpmTimer.AddTimer 300, "HighScoreEntryInit() '"
            End If
    End Select
End Sub

Sub HighScoreEntryInit()
    Dim i
    hsbModeActive = 1
    ' Table-specific
    i = RndNbr(4)
    PlaySoundVol "say-enterinitials"&i,VolCallout
    PlaySoundVol "got-track-hstd",VolDef/8
    hsLetterFlash = 0

    hsMaxDigits = DMDStd(kDMDStd_Initials)
    For i = 0 to hsMaxDigits-1 : hsEnteredDigits(i) = "" : Next
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<" ' < is back arrow
    if hsMaxDigits > 3 Then
        hsValidLetters = hsValidLetters + "~" ' Add the ~ key to the end to exit long initials entry
        hsX = 10
    Else
        hsX = 49
    End If
    hsCurrentLetter = 1
    If bUseFlexDMD Then
        Set HSscene = FlexDMD.NewGroup("highscore")
        tmrDMDUpdate.Enabled = False
        ' Note, these fonts aren't included with FlexDMD. Change to stock fonts for other tables
        HSscene.AddActor FlexDMD.NewLabel("name",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0),"YOUR NAME:")
        HSScene.GetLabel("name").SetAlignedPosition 2,2,FlexDMD_Align_TopLeft
        HSscene.AddActor FlexDMD.NewLabel("initials",FlexDMD.NewFont("skinny10x12.fnt", vbWhite, vbWhite, 0),"")
        HSScene.GetLabel("initials").SetAlignedPosition hsX,16,FlexDMD_Align_TopLeft
        DMDFlush()
        DMDDisplayScene HSscene
    End If
    HighScoreDisplayNameNow()

    HighScoreFlashTimer.Interval = 250
    HighScoreFlashTimer.Enabled = True
End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        'playsound "fx_Previous"
        'Table-specific
        PlaySoundVol "gotfx-hstd-left",VolDef
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0)then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        'playsound "fx_Next"
        'Table-specific
        PlaySoundVol "gotfx-hstd-right",VolDef
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters))then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Or keycode = RightMagnaSave Or keycode = LockBarKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<")then
            'playsound "fx_Enter"
            'Table-specific
            PlaySoundVol "gotfx-hstd-enter",VolDef
            if hsMaxDigits = 10 And (mid(hsValidLetters, hsCurrentLetter, 1) = "~") Then HighScoreCommitName() : Exit Sub
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3 And hsMaxDigits = 3) Or (hsCurrentDigit = 10 And hsMaxDigits = 10)then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = ""
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
    TempBotStr = hsStart
    For i = 0 to hsMaxDigits-1
        if(hsCurrentDigit> i)then TempBotStr = TempBotStr & hsEnteredDigits(i)
    Next

    if(hsCurrentDigit <> hsMaxDigits)then
        if(hsLetterFlash <> 0)then
            TempBotStr = TempBotStr & "_"
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    For i = 1 to hsMaxDigits-1
        if(hsCurrentDigit <i)then TempBotStr = TempBotStr & hsEnteredDigits(i)
    Next

    TempBotStr = TempBotStr & hsEnd

    If bUseFlexDMD Then
        FlexDMD.LockRenderThread
        With HSscene.GetLabel("initials")
            .Text = TempBotStr
            .SetAlignedPosition hsX,16,FlexDMD_Align_TopLeft
        End With
        FlexDMD.UnlockRenderThread
    Else
        dLine(0) = ExpandLine(TempTopStr, 0)
        dLine(1) = ExpandLine(TempBotStr, 1)
        'DMDUpdate 0
        'DMDUpdate 1
    End If
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2)then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    Dim i,bBlank,bm,scores
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = 2

    hsEnteredName="":bBlank=true
    For i = 0 to hsMaxDigits-1
        hsEnteredName = hsEnteredName & hsEnteredDigits(i)
        if hsEnteredDigits(i) <> " " Then bBlank=False
    Next
    if bBlank then hsEnteredName = "YOU"

    bm = GameSaveHighScore(Score(CurrentPlayer),hsEnteredName)
    SortHighscore
    Savehs
    For i = 0 to 4
        If Score(CurrentPlayer) = HighScore(i) then scores = 2^i
    Next
    scores = scores Or bm
    GameDoDMDHighScoreScene scores
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 0 to 4
        For j = 0 to 3
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


'********************************************************
' DMD Support. Updated to FlexDMD-specific API calls
' and eventually PuP
'********************************************************

Const eNone = 0        ' Instantly displayed
Const eBlink = 1
Const eBlinkFast = 2

'Dim FlexPath
Dim UltraDMD
Sub LoadFlexDMD
    Dim curDir
  Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
    If FlexDMD is Nothing Then
        MsgBox "No FlexDMD found. This table will not run without it."
        bUseFlexDMD = False
        Exit Sub
    End If
  SetLocale(1033)
  With FlexDMD
    .GameName = myGameName
    .TableFile = Table1.Filename & ".vpx"
        .ProjectFolder = ".\"&myGameName&".FlexDMD\"
    .Color = RGB(255, 44, 16)
    .RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
    .Width = 128
    .Height = 32
    .Clear = True
    .Run = True
        If VRRoom = 0 Then .Show = True Else .Show = False
  End With

    'Dim fso
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'curDir = fso.GetAbsolutePathName(".")
    'FlexPath = curDir & "\"&myGameName &".FlexDMD\"

    ' Check to make sure FlexDMD media is present
    Dim actor
    Set actor = FlexDMD.NewVideo("testvid","got-intro.gif")
    if actor is Nothing Then
        msgbox "FlexDMD directory gameofthrones.FlexDMD appears to be missing. Game will not run. Check to make sure it's in your Tables directory or wherever this table's VPX file is"
    End If
End Sub

Sub DMD_Clearforhighscore()
  DMDClearQueue
End Sub

Sub DMDClearQueue
  if bUseFlexDMD Then
    DMDqHead=0:DMDqTail=0
        FlexDMD.LockRenderThread
        FlexDMD.Stage.RemoveAll
        FlexDMD.UnlockRenderThread
        bDefaultScene = False
        DisplayingScene = Empty
  End If
End Sub

Sub PlayDMDScene(video, timeMs)
  debug.print "PlayDMDScene called"
End Sub

Sub DisplayDMDText(Line1, Line2, duration)
  debug.print "OldDMDText " & Line1 & " " & Line2
  if bUseFlexDMD Then
    'UltraDMD.DisplayScene00 "", Line1, 15, Line2, 15, 14, duration, 14
  Elseif bUsePUPDMD Then
    If bPupStarted then
      if Line1 = "" or Line1 = "_" then
        pupDMDDisplay "-", Line2, "" ,Duration/1000, 0, 10
      else
        pupDMDDisplay "-", Line1 & "^" & Line2, "" ,Duration/1000, 0, 10
      End If
    End If
  End If
End Sub

Sub DisplayDMDText2(Line1, Line2, duration, pri, blink)
  if bUseFlexDMD Then
    'UltraDMD.DisplayScene00 "", Line1, 15, Line2, 15, 14, duration, 14
  Elseif bUsePUPDMD Then
    If bPupStarted then
      if Line1 = "" or Line1 = "_" then
        pupDMDDisplay "-", Line2, "" ,Duration/1000, blink, pri
      else
        pupDMDDisplay "-", Line1 & "^" & Line2, "" ,Duration/1000, blink, pri
      End If
    End If
  End If
End Sub



Sub DMDId(id, toptext, bottomtext, duration) 'used in the highscore entry routine
  if bUseFlexDMD then
    'UltraDMD.DisplayScene00ExwithID id, false, "", toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
  Elseif bUsePUPDMD Then
    If bPupStarted then pupDMDDisplay "default", toptext & "^" & bottomtext, "" ,Duration/1000, 0, 10
  End If
End Sub

Sub DMDMod(id, toptext, bottomtext, duration) 'used in the highscore entry routine
  if bUseFlexDMD then
    'UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
  Elseif bUsePUPDMD Then
    If bPupStarted then pupDMDDisplay "default", toptext & "^" & bottomtext, "" ,Duration/1000, 0, 10
  End If
End Sub


Dim dCharsPerLine(2)
Dim dLine(2)

Sub DMD_Init() 'default/startup values
    Dim i, j

  if bUseFlexDMD then LoadFlexDMD()

    DMDFlush()
    dCharsPerLine(0) = 12
    dCharsPerLine(1) = 12
    dCharsPerLine(2) = 12
    For i = 0 to 2
        dLine(i) = Space(dCharsPerLine(i))
    Next
    DMD dLine(0), dLine(1), dLine(2), eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDFlush()
  DMDClearQueue
End Sub

Sub DMDBlank()
    Dim scene
    DMDClearQueue
    Set DefaultScene = FlexDMD.NewGroup("blank")
End Sub

Sub DMDScore()
    Dim tmp, tmp1, tmp2
  Dim TimeStr

  if bUseFlexDMD Then
        debug.log "Call to old DMDScore routine"
    ' DisplayDMDText RL(0,FormatScore(Score(CurrentPlayer))), "", 1000
  End If
End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
    Dim scene,line0
    if bUseFlexDMD Then
        Set scene = FlexDMD.NewGroup("dmd")
        If Text0 <> "" Then
            scene.AddActor FlexDMD.NewLabel("line0",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0), Text0)
            Set line0 = scene.GetLabel("line0")
            line0.SetAlignedPosition 64,5,FlexDMD_Align_Center
            if Effect0 = eBlink Then
                BlinkActor line0,0.1,int(TimeOn/200)
            Elseif Effect0 = eBlinkFast Then
                BlinkActor line0,0.05,int(TimeOn/100)
            End If
        End If
        If Text1 <> "" Then
            scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbWhite, 0), Text1)
            Set line0 = scene.GetLabel("line1")
            line0.SetAlignedPosition 64,16,FlexDMD_Align_Center
            if Effect1 = eBlink Then
                BlinkActor line0,0.1,int(TimeOn/200)
            Elseif Effect1 = eBlinkFast Then
                BlinkActor line0,0.05,int(TimeOn/100)
            End if
        End If
        If Text2 <> "" Then
            scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0), Text2)
            Set line0 = scene.GetLabel("line2")
            line0.SetAlignedPosition 64,26,FlexDMD_Align_Center
            if Effect2 = eBlink Then
                BlinkActor line0,0.1,int(TimeOn/200)
            Elseif Effect2 = eBlinkFast Then
                BlinkActor line0,0.05,int(TimeOn/100)
            End if
        End If
        if bFlush Then DMDClearQueue
        DMDEnqueueScene scene,0,TimeOn,TimeOn,1000,Sound
    Else
        DisplayDMDText Text0, Text1, TimeOn
        'if bUsePUPDMD and bPupStarted Then pupDMDDisplay "default", Text0 & "^" & Text1, "" ,2, 0, 10
        'if (bUsePUPDMD) Then pupDMDDisplay "attract", Text1 & "^" Text2, "@vidIntro.mp4" ,9, 1,    10
        PlaySoundVol Sound, VolDef
    End If
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

    if num <> 0 then
    NumString = CStr(abs(Num))
  else
    FormatScore = "00"
    exit function
  end if
    i = InStr(NumString,"+")
    If i > 0 Then
        Dim exp
        ' We got a scientific notation number, convert to a string
        exp = right(Numstring,Len(NumString)-i)
        Numstring = left(NumString,i-1)
        'Get rid of the period between the first and second digit
        NumString = Replace(NumString,".","")
        NumString = Replace(NumString,"E","")
        ' And add 0s to the right length
        For i = Len(NumString) to exp : NumString = NumString & "0" : Next
    End If

    For i = Len(NumString)-3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1))then
            NumString = left(NumString, i) & "," & right(NumString, Len(NumString)-i)
        end if
    Next
    FormatScore = NumString
End function

Function CL(id, NumString) 'center line
    Dim Temp, TempStr
  NumString = LEFT(NumString, dCharsPerLine(id))
    Temp = (dCharsPerLine(id)- Len(NumString)) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(id, NumString) 'right line
    Dim Temp, TempStr
  NumString = LEFT(NumString, dCharsPerLine(id))
    Temp = dCharsPerLine(id)- Len(NumString)
    TempStr = Space(Temp) & NumString
    RL = TempStr
End Function

Function FL(id, aString, bString) 'fill line
    Dim tmp, tmpStr
  aString = LEFT(aString, dCharsPerLine(id))
  bString = LEFT(bString, dCharsPerLine(id))
    tmp = dCharsPerLine(id)- Len(aString)- Len(bString)
  If tmp <0 Then tmp = 0
    tmpStr = aString & Space(tmp) & bString
    FL = tmpStr
End Function


'*************************
' FlexDMD Queue Management
'*************************
'
' FlexDMD supports queued scenes using its built-in Sequence class. However, there's no way to set priorities
' to allow new scenes to override playing scenes. In addition, there's no support for 'minimum play time' vs
' 'total play time', or for playing a sound with a scene. We want the ability to let a scene of a given priority play for at least 'minimum play time'
' as long as no scene of higher priority gets queued. If another scene of equal priority is queued, the playing scene
' will be replaced once it has played for 'minimum play time' ms.
' Queued higher priority scenes immediately replace playing lower priority scenes
' When no scenes are queued, show default scene (Score, "Choose..." or GameOver)
'
' If a scene gets queued that would take too long before it can be played due to items ahead of it, it gets dropped

Dim DMDSceneQueue(64,6)     ' Queue of scenes. Each entry has 7 fields: 0=Scene, 1=priority, 2=mintime, 3=maxtime, 4=waittime, 5=optionalSound, 6=timestamp
Dim DMDqHead,DMDqTail
Dim DMDtimestamp

'Queue up a FlexDMD scene in a virtual queue.

Sub DMDEnqueueScene(scene,pri,mint,maxt,waitt,sound)
    Dim i
    If bDefaultScene = False And Not IsEmpty(DisplayingScene) Then
        If DisplayingScene is scene Then
            ' Already playing. Update it
            DMDSceneQueue(DMDqHead,1) = pri
            DMDSceneQueue(DMDqHead,2) = mint
            DMDSceneQueue(DMDqHead,3) = maxt
            DMDSceneQueue(DMDqHead,4) = DMDtimestamp + waitt
            DMDSceneQueue(DMDqHead,5) = sound
            DMDSceneQueue(DMDqHead,6) = DMDtimestamp
            Exit Sub
        End If
    End If

    ' Check to see whether the scene is worth queuing
    If Not DMDCheckQueue(pri,waitt) Then
        ' debug.print "Discarding scene request with priority " & pri & " and waitt "&waitt
        Exit Sub
    End If

    ' Check to see if this is an update to an existing queued scene (e.g pictopops)
    Dim found:found=False
    If DMDqTail <> 0 Then
        For i = DMDqHead to DMDqTail-1
            If DMDSceneQueue(i,0) Is scene Then
                Found=True
                Exit For
            End If
        Next
    End If
    'Otherwise add to end of queue
    If Found = False Then i = DMDqTail:DMDqTail = DMDqTail + 1
    Set DMDSceneQueue(i,0) = scene
    DMDSceneQueue(i,1) = pri
    DMDSceneQueue(i,2) = mint
    DMDSceneQueue(i,3) = maxt
    DMDSceneQueue(i,4) = DMDtimestamp + waitt
    DMDSceneQueue(i,5) = sound
    DMDSceneQueue(i,6) = 0
    If DMDqTail > 64 Then       ' Ran past the end of the queue!
        DMDqTail = 64
    End if
    debug.print "Enqueued scene at "&i& " name: "&scene.Name
End Sub

' Check the queue to see whether a scene willing to wait 'waitt' time would play
Function DMDCheckQueue(pri,waitt)
    Dim i,wait:wait=0
    If DMDqTail=0 Then DMDCheckQueue = True: Exit Function
    DMDCheckQueue = False
    For i = DMDqHead to DMDqTail
        If DMDSceneQueue(i,4) > DMDtimestamp Then
            If DMDSceneQueue(i,1) = pri Then        'equal priority queued scene
                wait = wait + DMDSceneQueue(i,2)    ' so use mintime
            ElseIf DMDSceneQueue(i,1) < pri Then    'higher priority queued scene
                wait = wait + DMDSceneQueue(i,3)
            End If
            If wait > waitt Then Exit Function
        End If

    Next
    DMDCheckQueue = True
End Function

' Return the total length of the display queue, in ms
' Mainly used by end-of-ball processing to delay Bonus until all scenes have shown
' This isn't 100% accurate, as the last scene at a given priority level will play for maxtime
' before allowing a scene at a lower priority level play. We just add up all the mintimes
Function DMDGetQueueLength
    Dim i,j,len
    DMDGetQueueLength = 0
    j=0:len=0
    If DMDqTail = 0 Then Exit Function
    For j = 0 to 3  ' We don't care about really low priority scenes in this context
        For i = DMDqHead to DMDqTail
            If DMDSceneQueue(i,4) > DMDtimestamp Then
                If DMDSceneQueue(i,1) = j And DMDSceneQueue(i,4) > DMDtimestamp+len Then        'equal priority queued scene
                    len = len + DMDSceneQueue(i,2)    ' so use mintime
                End If
            End If
        Next
    Next
    DMDGetQueueLength = len
End Function

' Update DMD Scene. Called every 100ms
' Most of the work is done here. If scene queue is empty, display default scene (score, Game Over, etc)
' If scene queue isn't empty, check to see whether current scene has been on long enough or overwridden by a higher priority scene
' If it has, move to next spot in queue and search all of the queue for scene with highest priority, skipping any scenes that have timed out while waiting
Dim bDefaultScene,DefaultScene
Sub tmrDMDUpdate_Timer
    Dim i,j,bHigher,bEqual
    tmrDMDUpdate.Enabled = False
    DMDtimestamp = DMDtimestamp + 100   ' Set this to whatever frequency the timer uses
    If DMDqTail <> 0 Then
        ' Process queue
        ' Check to see if queue is idle (default scene on). If so, immediately play first item
        If bDefaultScene or (IsEmpty(DisplayingScene) And DMDqHead = 0) Then
            bDefaultScene = False
            debug.print "Idle: Displaying scene at " & DMDqHead & " Tail: "&DMDqTail
            DMDDisplayScene DMDSceneQueue(DMDqHead,0)
            DMDSceneQueue(DMDqHead,6) = DMDtimestamp
            If DMDSceneQueue(DMDqHead,5) <> ""  Then
                PlaySoundVol DMDSceneQueue(DMDqHead,5),VolDef
            ' Note, code below is game-specific - triggers the SuperJackpot scene update timer as soon as the scene plays
            ElseIf DMDSceneQueue(DMDqHead,0).Name = "bwsjp" Then
                tmrSJPScene.UserValue = 0
                tmrSJPScene.Interval = 300
                tmrSJPScene.Enabled = True
            End If
        Else
            ' Check to see whether there are any queued scenes with equal or higher priority than currently playing one
            bEqual = False: bHigher = False
            If DMDqTail > DMDqHead+1 Then
                For i = DMDqHead+1 to DMDqTail-1
                    If DMDSceneQueue(i,1) < DMDSceneQueue(DMDqHead,1) Then bHigher=True:Exit For
                    If DMDSceneQueue(i,1) = DMDSceneQueue(DMDqHead,1) Then bEqual = True:Exit For
                Next
            End If
            If bHigher Or (bEqual And (DMDSceneQueue(DMDqHead,6)+DMDSceneQueue(DMDqHead,2) <= DMDtimestamp)) Or _
                    (DMDSceneQueue(DMDqHead,6)+DMDSceneQueue(DMDqHead,3) <= DMDtimestamp) Then 'Current scene has played for long enough

                ' Skip over any queued scenes whose wait times have expired
                Do
                    DMDqHead = DMDqHead+1
                Loop While DMDSceneQueue(DMDqHead,4) < DMDtimestamp And DMDqHead < DMDqTail

                If DMDqHead > 64 Then       ' Ran past the end of the queue!
                    debug.print "DMDSceneQueue too big! Resetting"
                    DMDqHead = 0:DMDqTail = 0
                    tmrDMDUpdate.Enabled = True
                ElseIf DMDqHead = DMDqTail Then ' queue is empty
                    DMDqHead = 0:DMDqTail = 0
                    tmrDMDUpdate.Enabled = True
                Else
                    ' Find the next scene with the highest priority
                    j = DMDqHead
                    For i = DMDqHead to DMDqTail-1
                        If DMDSceneQueue(i,1) < DMDSceneQueue(j,1) Then j=i:DMDqHead=i
                    Next

                    ' Play the scene, and a sound if there's one to accompany it
                    bDefaultScene = False
                    debug.print "Displaying scene at " &j & " name: "&DMDSceneQueue(j,0).Name & " Head: "&DMDqHead & " Tail: "&DMDqTail
                    DMDSceneQueue(j,6) = DMDtimestamp
                    DMDDisplayScene DMDSceneQueue(j,0)
                    If DMDSceneQueue(j,5) <> ""  Then
                        PlaySoundVol DMDSceneQueue(j,5),VolDef
                    ' Note, code below is game-specific - triggers the SuperJackpot scene update timer as soon as the scene plays
                    ElseIf DMDSceneQueue(j,0).Name = "bwsjp" Then
                        tmrSJPScene.UserValue = 0
                        tmrSJPScene.Interval = 600
                        tmrSJPScene.Enabled = True
                    End If
                End If
            End If
        End If
    End If
    If DMDqTail = 0 Then ' Queue is empty
        ' Exit fast if defaultscene is already showing
        if bDefaultScene or IsEmpty(DefaultScene) then tmrDMDUpdate.Enabled = True : Exit Sub
        bDefaultScene = True
        If TypeName(DefaultScene) = "Object" Then
            DMDDisplayScene DefaultScene
        Else
            debug.print "DefaultScene is not an object!"
        End If
    End If
    tmrDMDUpdate.Enabled = True
End Sub

Dim DisplayingScene     ' Currently displaying scene
Sub DMDDisplayScene(scene)
    If TypeName(scene) <> "Object" then
    debug.print "DMDDisplayScene: scene is not an object! Type=" & TypeName(scene)
    exit sub
  ElseIf scene Is Nothing or IsEmpty(scene) Then
    debug.print "DMDDisplayScene: scene is empty!"
    exit Sub
  End If
    If Not IsEmpty(DisplayingScene) Then If DisplayingScene Is scene Then Exit Sub
    FlexDMD.LockRenderThread
    FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
    FlexDMD.Stage.RemoveAll
    FlexDMD.Stage.AddActor scene
    FlexDMD.Show = True
    FlexDMD.UnlockRenderThread
    Set DisplayingScene = scene
End Sub

' Create a new scene with a video file. If the video file
' is not found, look for an image file. If that's not found,
' create a new blank scene
Function NewSceneWithVideo(name,videofile)
  'Debug.Print "NewSceneWithVideo|Group:"&name&",Video:"&videofile
    Dim actor
    Set NewSceneWithVideo = FlexDMD.NewGroup(name)
    Set actor = FlexDMD.NewVideo(name&"vid",videofile & ".gif")
    If actor is Nothing Then
    Debug.Print "NewSceneWithVideo - Falling back to image: "&name&",File:"&videofile
    FlexDMD.LockRenderThread
        Set actor = FlexDMD.NewImage(name&"img",videofile&".png")
    FlexDMD.UnLockRenderThread
        if actor is Nothing Then
            debug.print "Warning: "&videofile&" image not found"
            Exit Function
        End if
    End If
    NewSceneWithVideo.AddActor actor
End Function

' Create a new scene with an image file. If that's not found,
' create a new blank scene
Function NewSceneWithImage(name,imagefile)
    Dim actor
  'Debug.Print "NewSceneWithImage|Group:"&name&",Image:"&imagefile
  FlexDMD.LockRenderThread
    Set NewSceneWithImage = FlexDMD.NewGroup(name)
    Set actor = FlexDMD.NewImage(name&"img",imagefile&".png")
    if actor is Nothing Then
    FlexDMD.UnLockRenderThread
    Exit Function
  end if
    NewSceneWithImage.AddActor actor
  FlexDMD.UnLockRenderThread
End Function

' Create a scene from a series of images. The only reason to use this
' function is if you need to use transparent images. If you don't, use
' an animated GIF - much easier. However, this does have one other advantage over
' an animated GIF: FlexDMD will loop animated GIFs, regardless of what the loop attribute is set to in the GIF
'  name   - name of the scene object returned
'  imgdir - directory inside the FlexDMD project folder where the images are stored
'  start  - number of first image
'  num    - number of images, numbered from image1..image<num>
'  fps    - frames per second - a delay of 1/fps is used between frames
'  hold   - if non-zero, how long to hold the last frame visible. If 0, the last scene will end with the last frame visible
'  repeat - Number of times to repeat. 0 or 1 means don't repeat
Function NewSceneFromImageSequence(name,imgdir,num,fps,hold,repeat)
    Set NewSceneFromImageSequence = NewSceneFromImageSequenceRange(name,imgdir,1,num,fps,hold,repeat)
End Function

Function NewSceneFromImageSequenceRange(name,imgdir,start,num,fps,hold,repeat)
    Dim scene,i,actor,af,blink,total,delay,e
    total = num/fps + hold
    delay = 1/fps
    e = start+num-1
    Set scene = FlexDMD.NewGroup(name)
    For i = start to e
        Set actor = FlexDMD.NewImage(name&i,imgdir&"\image"&i&".png")
        actor.Visible = 0
        Set af = actor.ActionFactory
        Set blink = af.Sequence()
        blink.Add af.Wait((i-start)*delay)
        blink.Add af.Show(True)
        blink.Add af.Wait(delay*1.2)    ' Slightly longer than one frame length to ensure no flicker
        if i=e And hold > 0 Then blink.Add af.Wait(hold)
        if repeat > 1 or i<e Then
            blink.Add af.Show(False)
            blink.Add af.Wait((e-i)*delay)
        End If
        If repeat > 1 Then actor.AddAction af.Repeat(blink,repeat) Else actor.AddAction blink
        scene.AddActor actor
    Next
    Set NewSceneFromImageSequenceRange = scene
End Function


' Add a blink action to an Actor in a FlexDMD scene.
' Usage: BlinkActor scene.GetActor("name"),blink-interval-in-seconds,repetitions
' Blink action is only natively supported in FlexDMD 1.9+
' poplabel.AddAction af.Blink(0.1, 0.1, 5)
Sub BlinkActor(actor,interval,times)
    Dim af,blink
    Set af = actor.ActionFactory
    Set blink = af.Sequence()
    blink.Add af.Show(True)
    blink.Add af.Wait(interval)
    blink.Add af.Show(False)
    blink.Add af.Wait(interval)
    actor.AddAction af.Repeat(blink,times)
End Sub

' Add action to delay toggling the state of an actor
' If on=true then delay then show, otherwise, delay then hide
Sub DelayActor(actor,delay,bOn)
    Dim af,blink
    Set af = actor.ActionFactory
    Set blink = af.Sequence()
    blink.Add af.Wait(delay)
    blink.Add af.Show(bOn)
    actor.AddAction blink
End Sub

' Add action to an actor to delay turning it on, wait, then turn it off
Sub FlashActor(actor,delayon,delayoff)
    Dim af,blink
    Set af = actor.ActionFactory
    Set blink = af.Sequence()
    blink.Add af.Wait(delayon)
    blink.Add af.Show(true)
    blink.Add af.Wait(delayoff)
    blink.Add af.Show(false)
    actor.AddAction blink
End Sub

' Add action to delay turning on an actor, then blink it
Sub DelayBlinkActor(actor,delay,blinkinterval,times)
    Dim af,seq,blink
    Set af = actor.ActionFactory
    Set seq = af.Sequence()
    Set blink = af.Sequence()
    seq.Add af.Wait(delay)
    blink.Add af.Show(True)
    blink.Add af.Wait(blinkinterval)
    blink.Add af.Show(False)
    blink.Add af.Wait(blinkinterval)
    seq.Add af.Repeat(blink,times)
    actor.AddAction seq
End Sub


Dim TableName : TableName = myGameName
' ***********************************************************************
' DAPHISHBOWL - STERN Service Menu
' - Modified by Skillman for SPIKE1/FlexDMD
'   - Refactored to remove distinction between Std and Game-specific features in feature array
'   - Added variables to indicate where game-specific features start in the array
'   - Moved all display logic outside the ServiceMenu sub
'   - Added Defaults to Settings array
'   - Optimized initialization of settings
'   - Added support for Boolean options in the Service Menu Adjustment
' ***********************************************************************

' SERVICE Key Definitions
Const ServiceCancelKey  = 8     ' 7 key
Const ServiceUpKey    = 9     ' 8 key
Const ServiceDownKey  = 10    ' 9 key
Const ServiceEnterKey   = 11    ' 0 key

Dim serviceSaveAttract
dim serviceIdx
Dim serviceLevel
dim bServiceMenu
dim bServiceVol
dim bServiceEdit
dim serviceOrigValue
Dim bInService:bInService=False
Dim MasterVol:MasterVol=100
'Dim VolBGVideo:VolBGVideo = cVolBGVideo
Dim VolBGMusic :VolBGMusic = cVolBGMusic
Dim VolCallout : VolCallout = cVolCallout
Dim VolDef:VolDef = cVolDef
Dim VolSfx:VolSfx = cVolSfx
Dim VolTable:VolTable = cVolTable
Dim TopArray:TopArray = Array("GO TO DIAGNOSTIC MENU","GO TO AUDITS MENU","GO TO ADJUSTMENTS MENU","GO TO UTILITIES MENU","GO TO TOURNAMENT MENU","GO TO REDEMPTION MENU","EXIT SERVICE MENU")
Dim AdjArray:AdjArray = Array("STANDARD ADJUSTMENTS","FEATURE ADJUSTMENTS","PREVIOUS MENU","EXIT SRVICE MENU","HELP")
Const kMenuNone   = -1
Const kMenuTop    = 0
Const kMenuAdj    = 1
Const kMenuAdjStd = 2
Const kMenuAdjGame  = 3

Sub StartServiceMenu(keycode)
  Dim i
  Dim maxVal
  Dim minVal
  Dim dataType
  dim TmpStr
  Dim direction
  Dim values
  Dim valuesLst
  Dim NewVal
  Dim displayText
    Dim SvcFontSm
    Dim SvcFontBg
    Dim line3

  if keycode=ServiceEnterKey and bInService=False then
'debug.print "Start Service:" & bInService
    PlaySoundVol "stern-svc-enter", VolSfx
    serviceSaveAttract=bAttractMode
    bAttractMode=False
    bServiceMenu=False
    bServiceEdit=False
    bInService=True
    bServiceVol=False
    serviceIdx=-1
    serviceLevel=kMenuNone

        tmrDMDUpdate.Enabled = False
        tmrAttractModeScene.Enabled = False

        CreateServiceDMDScene serviceIdx

  elseif keycode=ServiceCancelKey then          ' 7 key cancels the service menu
    if bInService then
      Select case serviceLevel:
        case kMenuTop, kMenuNone:
          bInService=False
          bServiceVol=False
          bAttractMode=serviceSaveAttract

          'fix: if volume timer is running and service menu is exited, force attract mode
          'leaving this as a separate if/then to accomodate future tweaks to service menu
          if tmrService.Enabled then
            bAttractMode = True
          end if

          if bAttractMode then
            StartAttractMode
                    Else
                        bDefaultScene = False
                        tmrDMDUpdate.Enabled = True
          End if

        case kMenuAdj:
          serviceLevel=kMenuNone
          serviceIdx=0
          StartServiceMenu 11
                    Exit Sub
        case kMenuAdjStd,kMenuAdjGame:
          if bServiceEdit then
            bServiceEdit=False
                        DMDStd(DMDMenu(serviceIdx).StdIdx)=serviceOrigValue
                        UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
          else
                        serviceLevel=kMenuTop
            serviceIdx=2
            StartServiceMenu 11
                        Exit Sub
          End if
      End Select
            PlaySoundVol "stern-svc-cancel", VolSfx
    End if
  elseif bInService Then
    if keycode=ServiceEnterKey then   ' Select
            Dim svcsound : svcsound = "stern-svc-enter"
      select case serviceLevel
        case kMenuNone:
          serviceLevel=kMenuTop
          serviceIdx=0
                    CreateServiceDMDScene serviceLevel
        case kMenuTop:
          if serviceIdx=2 then
            serviceLevel=kMenuAdj : serviceIdx=0
            CreateServiceDMDScene serviceLevel
          elseif serviceIdx=6 then
            StartServiceMenu 8    ' Exit
            Exit sub
          else
            svcsound = "sfx-deny"
          End if
        case kMenuAdj:
                    Select Case serviceIdx
              Case 0,1:
                            serviceLevel=kMenuAdjStd+serviceIdx
                            If serviceLevel = kMenuAdjStd Then serviceIdx=0 Else serviceIdx = MinFetDMDSetting
                            CreateServiceDMDScene serviceLevel
                            StartServiceMenu 0
              Case 2:     ' Go Up
                            serviceLevel=kMenuNone : serviceIdx=0
                            StartServiceMenu 11
                            Exit sub
              Case 3:
                            StartServiceMenu 8    ' Exit
                            Exit sub
              Case else:
                svcsound = "sfx-deny"
          End Select
        case kMenuAdjStd,kMenuAdjGame:
          if DMDMenu(serviceIdx).ValType="FUN" then   ' Function
                        Select Case DMDMenu(serviceIdx).StdIdx
                            Case 0: reseths
                            Case 1: ClearAll
                        End Select
            UpdateServiceDMDScene "<DONE>"
                        svcsound = "stern-svc-set"
          else
            if bServiceEdit=False then  ' Start Editing
              bServiceEdit=True
                            serviceOrigValue=DMDStd(DMDMenu(serviceIdx).StdIdx)
              UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
            else              ' Save the change
              bServiceEdit=False
              UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
                            svcsound = "stern-svc-set"
            End if
          End if
      End Select
            If svcsound <> "" Then PlaySoundVol svcsound, VolSfx
    elseif serviceLevel<>kMenuNone then
      if bServiceEdit = False then
        if keycode=ServiceUpKey then  ' Left
          PlaySoundVol "stern-svc-minus", VolSfx
          serviceIdx=serviceIdx-1
          if serviceIdx<0 then serviceIdx=0
        elseif keycode=ServiceDownKey then  ' Right
          PlaySoundVol "stern-svc-plus", VolSfx
          serviceIdx=serviceIdx+1
        End if

        select case serviceLevel
          case kMenuTop:
            if serviceIdx>6 then serviceIdx=0
            UpdateServiceDMDScene ""
          case kMenuAdj:
            if serviceIdx>4 then serviceIdx=0
            UpdateServiceDMDScene ""
          case kMenuAdjStd
                        if serviceIdx > MaxStdDMDSetting Then serviceIdx = MaxStdDMDSetting
                        if DMDMenu(serviceIdx).ValType<>"FUN" then
              UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
            else
              UpdateServiceDMDScene "<EXECUTE>"
            End if
                    case kMenuAdjGame:
                        if serviceIdx < MinFetDMDSetting Then serviceIdx = MinFetDMDSetting
                        If serviceIdx > MaxFetDMDSetting Then serviceIdx = MaxFetDMDSetting
                        UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
        End Select
      else

            ' ******************************************************    HANDLE EDIT MODE
        dataType=mid(DMDMenu(serviceIdx).ValType, 1, 3)
        minVal=-1
        maxVal=-1
        if Instr(DMDMenu(serviceIdx).ValType, "[")<> 0 then
          TmpStr =mid(DMDMenu(serviceIdx).ValType, 5, Len(DMDMenu(serviceIdx).ValType)-5)
          values = Split(TmpStr, ":")
          if dataType="INT" or dataType="PCT" then
            minVal=CLNG(values(0))
            maxVal=CLNG(values(1))
          End if
        End if

        direction=0
        if keycode=ServiceUpKey then  ' Left
          PlaySoundVol "stern-svc-minus", VolSfx
          direction=-1
        elseif keycode=ServiceDownKey then  ' Right
          PlaySoundVol "stern-svc-plus", VolSfx
          direction=1
        End if
        if direction<>0 then
          if dataType="INT" or dataType="PCT" then
                        NewVal=DMDStd(DMDMenu(serviceIdx).StdIdx) + (direction*DMDMenu(serviceIdx).Increment)
                        if minVal=-1 or (NewVal <= maxVal and NewVal >= minVal) then
                            DMDStd(DMDMenu(serviceIdx).StdIdx)=DMDStd(DMDMenu(serviceIdx).StdIdx) + (direction*DMDMenu(serviceIdx).Increment)
                            SaveValue TableName, "DMDStd_"&DMDMenu(serviceIdx).StdIdx, DMDStd(DMDMenu(serviceIdx).StdIdx) ' SAVE
                            if DMDMenu(serviceIdx).StdIdx<>kDMDStd_Initials and DMDMenu(serviceIdx).StdIdx<>kDMDStd_LeftStartReset then
                                SaveValue TableName, "dmdCriticalChanged", "True"   ' SAVE
                                dmdCriticalChanged=True
                            End if
                        end if
                        displayText=DMDStd(DMDMenu(serviceIdx).StdIdx)

          elseif dataType="LST" then
                        For i = 0 to ubound(values)
                            valuesLst=Split(values(i), ",")
'debug.print "EDIT LST:" & valuesLst(0) & " " & valuesLst(1) & " val: " & DMDStd(DMDMenu(serviceIdx).StdIdx) & " idx:" & i & " ubound:" & ubound(values)
                            if DMDStd(DMDMenu(serviceIdx).StdIdx)&"" = valuesLst(0) and direction=1 then
                                If i <  ubound(values) Then valuesLst=Split(values(i+1), ",") Else  valuesLst=Split(values(0), ",")
                                DMDStd(DMDMenu(serviceIdx).StdIdx)=valuesLst(0)
                                SaveValue TableName, "DMDStd_"&DMDMenu(serviceIdx).StdIdx, DMDStd(DMDMenu(serviceIdx).StdIdx) ' SAVE
                                displayText=valuesLst(1)
                                Exit For
                            elseif DMDStd(DMDMenu(serviceIdx).StdIdx)&"" = valuesLst(0) and direction=-1 then
                                If i > lbound(values) Then valuesLst=Split(values(i-1), ",") Else valuesLst=Split(values(ubound(values)), ",")
                                DMDStd(DMDMenu(serviceIdx).StdIdx)=valuesLst(0)
                                SaveValue TableName, "DMDStd_"&DMDMenu(serviceIdx).StdIdx, DMDStd(DMDMenu(serviceIdx).StdIdx) ' SAVE
                                displayText=valuesLst(1)
                                Exit For
                            End if
                        Next
                    elseif dataType="BOO" then 'Boolean
                        If DMDStd(DMDMenu(serviceIdx).StdIdx) = 1 Then
                            DMDStd(DMDMenu(serviceIdx).StdIdx) = 0 : displayText = "OFF"
                        Else
                            DMDStd(DMDMenu(serviceIdx).StdIdx) = 1 : displayText = "ON"
                        End if
                        SaveValue TableName, "DMDStd_"&DMDMenu(serviceIdx).StdIdx, DMDStd(DMDMenu(serviceIdx).StdIdx) ' SAVE
          End if

          if displayText<>"" then UpdateServiceDMDScene displayText
        End If
        if DMDMenu(serviceIdx).StdIdx = kDMDStd_BallsPerGame Then BallsPerGame = DMDStd(kDMDStd_BallsPerGame)
      End if
    End if
  elseif keycode=ServiceUpKey or keycode=ServiceDownKey then    ' If you press 8 & 9 without being in service you do volume
    if keycode=ServiceUpKey and MasterVol>0 then  ' Left
      MasterVol=MasterVol-1
      PlaySoundVol "stern-svc-minus", VolSfx
    elseif keycode=ServiceDownKey and MasterVol<100 then  ' Right
      PlaySoundVol "stern-svc-plus", VolSfx
      MasterVol=MasterVol+1
    End if

    SvcSetVolume

    if bServiceVol=False then
      bServiceVol=True
      serviceSaveAttract=bAttractMode
      bAttractMode=False
            tmrDMDUpdate.Enabled = False
            tmrAttractModeScene.Enabled = False
            CreateServiceDMDScene 0
    End if

    UpdateServiceDMDScene ""

    tmrService.Enabled = False
    tmrService.Interval = 5000
    tmrService.Enabled = True
  End if
End Sub

Sub SvcSetVolume
    'VolBGVideo = cVolBGVideo * (MasterVol/100.0)
    VolBGMusic = cVolBGMusic * (MasterVol/100.0)
    VolDef   = cVolDef * (MasterVol/100.0)
    VolSfx   = cVolSfx * (MasterVol/100.0)
    VolCallout = cVolCallout * (MasterVol/100.0)
End Sub

Sub tmrService_Timer()
  tmrService.Enabled = False

    SaveValue TableName,"MasterVol",MasterVol
  bServiceVol=False
  bAttractMode=serviceSaveAttract
  if bAttractMode then
    StartAttractMode
    Else
        bDefaultScene = False ' Force a refresh of the DMD
        tmrDMDUpdate.Enabled = true
  End if
End Sub

' Decode the ValType from the feature definition array to extract what kind of setting it is
' If it's INT or PCT, just pass back the current value as the text
' If it's LST, look up the current value in the LST to find the string equiv
' If it's BOOL, convert to "OFF" or "ON"
Function GetTextForAdjustment(idx)
    Dim txt,dataType,values,valuesLst,TmpStr,i
    txt = DMDStd(DMDMenu(idx).StdIdx) ' For INT and PCT types

    dataType=mid(DMDMenu(serviceIdx).ValType, 1, 3) ' First 3 chars of valType
    if Instr(DMDMenu(serviceIdx).ValType, "[")<> 0 then
        TmpStr =mid(DMDMenu(serviceIdx).ValType, 5, Len(DMDMenu(serviceIdx).ValType)-5)
        values = Split(TmpStr, ":")
    End if

    Select Case dataType
        Case "BOO": If DMDStd(DMDMenu(idx).StdIdx) <> 0 Then txt = "ON" Else txt = "OFF"
        Case "LST"
            For i = 0 to ubound(values)
        valuesLst=Split(values(i), ",")
                If DMDStd(DMDMenu(idx).StdIdx)&"" = valuesLst(0) Then txt = valuesLst(1) : Exit For
            Next
    End Select
    debug.print "DataType: "&dataType&"  Returning "&txt
    GetTextForAdjustment = txt
End Function

Class DMDSettings
  Public Name
  Public StdIdx
  Public Deflt
  Public ValType          ' bool, pct
  Public Increment      ' value to increment by
    Sub Class_Initialize
        me.Name="EMPTY"
    End Sub
  Public sub Setup(Name, StdIdx, Deflt, ValType, Increment)
    me.name=name
    me.StdIdx=StdIdx
        me.Deflt = Deflt
    me.ValType=ValType
    me.Increment=Increment
  End sub
End Class

' Set this according to the size of your settings list
Dim MaxStdDMDSetting,MinFetDMDSetting,MaxFetDMDSetting : MaxStdDMDSetting = 27
Dim DMDMenu(60)       ' MaxDMDSetting
Dim dmdChanged:dmdChanged=False             ' Did we change a value
Dim dmdCriticalChanged:dmdCriticalChanged=False     ' Did we change a critical value

' Indexes into the DMDStd array, where the actual value is stored. These indexes could just as easily
' count up from 0, but these values were taken from Pinball Browser and match the actual game
Const kDMDStd_ExtraBallLimit = &H3B
Const kDMDStd_ExtraBallPCT = &H3C
Const kDMDStd_MatchPCT = &H3D           '√
Const kDMDStd_TiltWarn = &H3F
Const kDMDStd_TiltDebounce = &H40
Const kDMDStd_LeftStartReset = &H46     '√
Const kDMDStd_BallSave = &H4C           '√
Const kDMDStd_BallSaveExtend = &H4D     '√
Const kDMDStd_ReplayType = &H2C         '√
Const kDMDStd_AutoReplayStart = &H30    '√
Const kDMDStd_BallsPerGame = &H3E       '√
Const kDMDStd_FreePlay = &H42           '√
Const kDMDStd_HighScoreGC = &H54        '√
Const kDMDStd_HighScore1 = &H55         '√
Const kDMDStd_HighScore2 = &H56         '√
Const kDMDStd_HighScore3 = &H57         '√
Const kDMDStd_HighScore4 = &H58         '√
Const kDMDStd_HighScoreAward = &H59     '√
Const kDMDStd_GCAwards = &H5A           '√
Const kDMDStd_HS1Awards = &H5B          '√
Const kDMDStd_HS2Awards = &H5C          '√
Const kDMDStd_HS3Awards = &H5D          '√
Const kDMDStd_HS4Awards = &H5E          '√
Const kDMDStd_Initials = &H5F           '√
Const kDMDStd_MusicAttenuation = &H60
Const kDMDStd_SpeechAttenuation = &H61

Const kDMDStd_LastStdSetting = &H67

'Table-specific indexes below here
Const kDMDFet_HouseCompleteExtraBall = &H68     '√
Const kDMDFet_CasualPlayerMode = &H6C
Const kDMDFet_PlayerHouseChoice = &H6D          '√
Const kDMDFet_DisableRampDropTarget = &H71      '√
Const kDMDFet_ActionButtonAction = &H73   ' Toggle between house actions or "Iron Bank" (pre v1.37 behaviour)
Const kDMDFet_LannisterButtonsPerGame = &H74    '√
Const kDMDFet_TargaryenFreezeTime = &H75        '√
Const kDMDFet_TargaryenHousePower = &H76        '√ toggle between 1, 2 or all dragons completed. 1 or 2 scores 30M per at start of game
Const kDMDFet_DireWolfFrequency = &H77 ' Start Action button: 1-5 per game, 1 per ball, or 0
Const kDMDFet_SwordsUnlockMultiplier = &H78     '√
Const kDMDFet_RamHitsForMult = &H79
Const kDMDFet_TwoBankDifficulty = &H7A          '√
Const kDMDFet_BWMB_SaveTimer = &H7D             '√
Const kDMDFet_WallMB_SaveTimer = &H7E           '√
Const kDMDFet_HOTKMB_SaveTimer = &H7F           '√
Const kDMDFet_ITMB_SaveTimer = &H80             '√
Const kDMDFet_MMMB_SaveTimer = &H81             '√
Const kDMDFet_WHCMB_SaveTimer = &H82            '√
Const kDMDFet_CMB_SaveTimer = &H83              '√
Const kDMDFet_CastleLoopCompletesHouse = &H85
Const kDMDFet_InactivityPauseTimer = &H96       '√
Const kDMDFet_ReduceWHCLamps = &H97             '√
Const kDMDFet_MidnightMadness = &H98


Sub DMDSettingsInit()
  Dim i
  Dim x

  For i = 0 to 60
    set DMDMenu(i)=new DMDSettings
    Next

        '  Name, Index, Default, Variable type
    DMDMenu(0).Setup "HSTD  INITIALS",            kDMDStd_Initials  , 3,    "LST[3,3 INITIALS:10,10 LETTER NAME]", 1
    DMDMenu(1).Setup "EXTRA  BALL  LIMIT",          kDMDStd_ExtraBallLimit, 5,    "INT[0:9]", 1
    DMDMenu(2).Setup "EXTRA  BALL  PERCENT",        kDMDStd_ExtraBallPCT, 25,   "PCT[0:50]", 1
    DMDMenu(3).Setup "MATCH  PERCENT",            kDMDStd_MatchPCT,   9,    "PCT[0:10]", 1
    DMDMenu(4).Setup "LEFT+START  RESETS",          kDMDStd_LeftStartReset, 2,    "LST[0,DISABLED:1,ALWAYS:2,FREE PLAY]", 1
    DMDMenu(5).Setup "BALL  SAVE  SECONDS",           kDMDStd_BallSave,   5,    "INT[0:15]", 1
    DMDMenu(6).Setup "BALL  SAVE  EXTEND  SEC",         kDMDStd_BallSaveExtend, 3,    "INT", 1
    DMDMenu(7).Setup "TILT  WARNINGS",            kDMDStd_TiltWarn  , 2,    "INT[0:3]", 1
    DMDMenu(8).Setup "TILT  DEBOUNCE",            kDMDStd_TiltDebounce, 1000, "INT[750:1500]", 1
    DMDMenu(9).Setup "REPLAY  TYPE",            kDMDStd_ReplayType,     1,    "LST[1,EXTRA GAME:2,EXTRA BALL]", 1
    DMDMenu(10).Setup "AUTO  REPLAY  START",          kDMDStd_AutoReplayStart,300000000,"INT[10000000:1000000000]", 5000000
    DMDMenu(11).Setup "BALL  PER  GAME",                      kDMDStd_BallsPerGame,   3,     "INT[3:5]", 1
    DMDMenu(12).Setup "FREE  PLAY",                          kDMDStd_FreePlay,       0,     "BOOL", 1
    DMDMenu(13).Setup "GRAND  CHAMPION  SCORE",               kDMDStd_HighScoreGC,    750000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(14).Setup "HIGH  SCORE  1",                       kDMDStd_HighScore1,     500000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(15).Setup "HIGH  SCORE  2",                       kDMDStd_HighScore2,     400000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(16).Setup "HIGH  SCORE  3",                       kDMDStd_HighScore3,     300000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(17).Setup "HIGH  SCORE  4",                       kDMDStd_HighScore4,     200000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(18).Setup "HIGH  SCORE  AWARDS",                  kDMDStd_HighScoreAward, 1,      "BOOL", 1
    DMDMenu(19).Setup "GRAND CHAMPION  AWARDS",              kDMDStd_GCAwards,       1,     "LST[0,0 CREDITS:1,1 CREDIT:2,2 CREDITS:3,3 CREDITS:4,4 CREDITS]", 1
    DMDMenu(20).Setup "HIGH  SCORE  1  AWARDS",                kDMDStd_HS1Awards,      1,     "LST[0,0 CREDITS:1,1 CREDIT:2,2 CREDITS:3,3 CREDITS]", 1
    DMDMenu(21).Setup "HIGH  SCORE  2  AWARDS",                kDMDStd_HS2Awards,      0,     "LST[0,0 CREDITS:1,1 CREDIT:2,2 CREDITS]", 1
    DMDMenu(22).Setup "HIGH  SCORE  3  AWARDS",                kDMDStd_HS3Awards,      0,     "LST[0,0 CREDITS:1,1 CREDIT]", 1
    DMDMenu(23).Setup "HIGH  SCORE  4  AWARDS",                kDMDStd_HS4Awards,      0,     "LST[0,0 CREDITS:1,1 CREDIT]", 1
    DMDMenu(24).Setup "MUSIC  ATTENUATION",                  kDMDStd_MusicAttenuation,0,    "INT[-60:60]", 5
    DMDMenu(25).Setup "SPEECH  ATTENUATION",                 kDMDStd_SpeechAttenuation,0,   "INT[-60:60]", 5
    DMDMenu(26).Setup "CLEAR  HIGHSCORE",           0,    0,            "FUN", 1
  DMDMenu(27).Setup "RESET  ALL  SETTINGS",         1,    1,                  "FUN", 1

    '***********************
    'Table-Specific Settings
    '***********************

    MinFetDMDSetting = 30
    DMDMenu(30).Setup "HOUSE  COMPLETE  EXTRA  BALL",         kDMDFet_HouseCompleteExtraBall,   8,      "LST[2,2:3,3:4,4:5,5:6,6:7,7:8,AUTO]", 1
    DMDMenu(31).Setup "CASUAL  PLAYER  MODE",                 kDMDFet_CasualPlayerMode,         0,      "BOOL", 1
    DMDMenu(32).Setup "PLAYER  HOUSE  CHOICE",                kDMDFet_PlayerHouseChoice,        0,      "LST[0,PLAYERS CHOICE:1,STARK:2,BARATHEON:3,LANNISTER:4,GREYJOY:5,TYRELL:6,MARTELL:7,TARGARYEN:8,RANDOM]", 1
    DMDMenu(33).Setup "DISABLE  RAMP  DROP  TARGET",          kDMDFet_DisableRampDropTarget,    0,      "BOOL", 1
    DMDMenu(34).Setup "ACTION  BUTTON  ACTION",               kDMDFet_ActionButtonAction,       1,      "LST[0,IRON BANK:1,HOUSE POWER]", 1
    DMDMenu(35).Setup "LANNISTER  BUTTONS  PER  GAME",        kDMDFet_LannisterButtonsPerGame,  8,      "INT[4:12]", 1
    DMDMenu(36).Setup "TARGARYEN  FREEZE  TIME",              kDMDFet_TargaryenFreezeTime,      12,     "INT[4:16]", 1
    DMDMenu(37).Setup "TARGARYEN  HOUSE  POWER",              kDMDFet_TargaryenHousePower,      3,      "LST[1,ONE DRAGON:2,TWO DRAGONS:3,COMPLETED]", 1
    DMDMenu(38).Setup "DIRE  WOLF  FREQUENCY",                kDMDFet_DireWolfFrequency,        6,      "LST[0,DISABLED:1,1 PER GAME:2,2 PER GAME:3,3 PER GAME:4,4 PER GAME:5,5 PER GAME:6,1 PER BALL", 1
    DMDMenu(39).Setup "SWORDS  UNLOCK  MULTIPLIER  X",        kDMDFet_SwordsUnlockMultiplier,   1,      "INT[0:2]", 1
    DMDMenu(40).Setup "BATTERING  RAM  HITS  FOR  MULTIPLIER",kDMDFet_RamHitsForMult,           1,      "INT[0:2]", 1
    DMDMenu(41).Setup "TWO  BANK  DIFFICULTY",                kDMDFet_TwoBankDifficulty,        1,      "LST[0,EASY:1,NORMAL:2,HARD]", 1
    DMDMenu(42).Setup "BLACKWATER  BALL  SAVE  TIMER",        kDMDFet_BWMB_SaveTimer,           10,     "INT[5:60]", 1
    DMDMenu(43).Setup "WALL  MULTIBALL  BALL  SAVE  TIMER",   kDMDFet_WallMB_SaveTimer,         20,     "INT[5:60]", 1
    DMDMenu(44).Setup "HAND  OF  THE  KING  BALL  SAVE  TIMER",kDMDFet_HOTKMB_SaveTimer,        20,     "INT[5:60]", 1
    DMDMenu(45).Setup "IRON  THRONE  BALL  SAVE  TIMER",      kDMDFet_ITMB_SaveTimer,           25,     "INT[5:60]", 1
    DMDMenu(46).Setup "MIDNIGHT  MADNESS  BALL  SAVE  TIMER", kDMDFet_MMMB_SaveTimer,           12,     "INT[5:60]", 1
    DMDMenu(47).Setup "WINTER  HAS  COME  BALL  SAVE  TIMER", kDMDFet_WHCMB_SaveTimer,          15,     "INT[5:60]", 1
    DMDMenu(48).Setup "CASTLE  MULTIBALL  BALL  SAVE  TIMER", kDMDFet_CMB_SaveTimer,            15,     "INT[5:60]", 1
    DMDMenu(49).Setup "CASTLE  LOOP  COMPLETES  HOUSE",       kDMDFet_CastleLoopCompletesHouse, 1,      "BOOL", 1
    DMDMenu(50).Setup "ALLOW  INACTIVITY  TO  PAUSE  TIMERS", kDMDFet_InactivityPauseTimer,     1,      "BOOL", 1
    DMDMenu(51).Setup "REDUCE  WINTER  IS  COMING  FLASHERS", kDMDFet_ReduceWHCLamps,           0,      "BOOL", 1
    DMDMenu(52).Setup "MIDNIGHT  MADNESS",                    kDMDFet_MidnightMadness,          1,      "BOOL", 1
    MaxFetDMDSetting = 52

    '** End Table-specific settings **

    ' Load Values from NVRAM. If not set, load default from array
    For i = 0 to MaxFetDMDSetting
        If DMDMenu(i).name <> "EMPTY" Then
            x = LoadValue(TableName, "DMDStd_"&DMDMenu(i).StdIdx)
            If (x <> "") Then
                DMDStd(DMDMenu(i).StdIdx) = x
            Else
                DMDStd(DMDMenu(i).StdIdx) = DMDMenu(i).Deflt
            End If
        End if
    Next
    x = LoadValue(TableName, "dmdCriticalChanged"): If(x<>"") Then dmdCriticalChanged=True
    x = LoadValue(TableName,"MasterVol") : If x<>"" Then MasterVol=CInt(x) : SvcSetVolume

    BallsPerGame = DMDStd(kDMDStd_BallsPerGame)


End Sub
Dim DMDStd(200)

' Create the FlexDMD scene based on what service menu level we're at

Dim SvcScene
Sub CreateServiceDMDScene(lvl)
    Dim scene,i,max,offset,SvcFontSm,SvcFontBg
    offset=0
    Set SvcFontSm = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
    Set SvcFontBg = FlexDMD.NewFont("udmd-f7by10.fnt",vbWhite,vbWhite,0)
    Set scene = FlexDMD.NewGroup("service")
    If bServiceVol Then ' Not in Service menu, just adjusting volume
        scene.AddActor FlexDMD.NewLabel("svcl1",SvcFontSm,"USE +/- TO ADJUST VOLUME")
        scene.GetLabel("svcl1").SetAlignedPosition 64,0,FlexDMD_Align_Top
        scene.AddActor FlexDMD.NewLabel("svcl2",SvcFontBg,"VOLUME "&MasterVol)
        scene.GetLabel("svcl2").SetAlignedPosition 64,14,FlexDMD_Align_Top
    Else
        Select Case lvl
            Case kMenuNone
                scene.AddActor FlexDMD.NewLabel("svcl1",SvcFontSm,"GAME OF THRONES PREMIUM"&vbLf&"V1.37.0    "&myVersion&"  SYS. 2.31.0")
                scene.GetLabel("svcl1").SetAlignedPosition 64,0,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl2",SvcFontBg,"SERVICE MENU")
                scene.GetLabel("svcl2").SetAlignedPosition 64,14,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl3",SvcFontSm,"PRESS SELECT TO CONTINUE")
                scene.GetLabel("svcl3").SetAlignedPosition 64,25,FlexDMD_Align_Top
            Case kMenuTop,kMenuAdj  ' one of the icon levels. Create row of icons
                If lvl=kMenuTop Then
                    max = 6
                    scene.AddActor FlexDMD.NewLabel("svcl4",SvcFontSm,TopArray(0))
                else
                    max = 4 : offset=17
                    scene.AddActor FlexDMD.NewLabel("svcl4",SvcFontSm,AdjArray(0))
                End if
                For i = 0 to max
                    scene.AddActor FlexDMD.NewImage("iconoff"&i,"got-svciconoff"&lvl&i&".png")
                    With scene.GetImage("iconoff"&i)
                        .SetAlignedPosition offset+i*17,0,FlexDMD_Align_TopLeft
                        if i = 0 Then .Visible = 0
                    End With
                    scene.AddActor FlexDMD.NewImage("iconon"&i,"got-svciconon"&lvl&i&".png")
                    With scene.GetImage("iconon"&i)
                        .SetAlignedPosition offset+i*17,0,FlexDMD_Align_TopLeft
                        If i > 0 Then .Visible = 0
                    End With
                Next
                scene.GetImage("iconoff0").Visible = 0
                scene.GetLabel("svcl4").SetAlignedPosition 64,25,FlexDMD_Align_Top
            Case kMenuAdjStd,kMenuAdjGame ' Adjustment level, create the base adjustment screen
                scene.AddActor FlexDMD.NewLabel("svcl1",SvcFontSm,"ADJUSTMENT")
                scene.GetLabel("svcl1").SetAlignedPosition 64,0,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl2",FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0),"SETTING NAME")
                scene.GetLabel("svcl2").SetAlignedPosition 64,7,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl3",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt",vbWhite,vbWhite,0),"VALUE")
                scene.GetLabel("svcl3").SetAlignedPosition 64,16,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl4",SvcFontSm,"INSTALLED/FACTORY DEFAULT")
                scene.GetLabel("svcl4").SetAlignedPosition 64,25,FlexDMD_Align_Top
        End Select
    End if
    Set SvcScene = scene
    DMDDisplayScene SvcScene
End Sub

Sub UpdateServiceDMDScene(line3)
    Dim max,i
    If serviceLevel = kMenuTop Then max = 6 else max = 4
    FlexDMD.LockRenderThread
    If bServiceVol Then ' Just update the Volume level
        With SvcScene.GetLabel("svcl2")
            .Text = "VOLUME "&MasterVol
            .SetAlignedPosition 64,14,FlexDMD_Align_Top
        End With
    Else
        If serviceLevel = kMenuTop Or serviceLevel = kMenuAdj Then
            For i = 0 to max
                With SvcScene.GetImage("iconon"&i)
                    If i=serviceIdx Then .Visible=1 Else .Visible=0
                End With
                With SvcScene.GetImage("iconoff"&i)
                    If i=serviceIdx Then .Visible=0 Else .Visible=1
                End With
            Next
        End If
        Select Case serviceLevel
            Case kMenuTop:
                SvcScene.GetLabel("svcl4").Text = TopArray(serviceIdx)
                SvcScene.GetLabel("svcl4").SetAlignedPosition 64,25,FlexDMD_Align_Top
            Case kMenuAdj:
                SvcScene.GetLabel("svcl4").Text = AdjArray(serviceIdx)
                SvcScene.GetLabel("svcl4").SetAlignedPosition 64,25,FlexDMD_Align_Top
            Case kMenuAdjStd,kMenuAdjGame
                With SvcScene.GetLabel("svcl1")
                    if serviceIdx < MinFetDMDSetting Then
                        .Text = "STANDARD ADJUSTMENT #" & serviceIdx & "(" & DMDMenu(serviceIdx).StdIdx & ")"
                    Elseif DMDMenu(serviceIdx).ValType<>"FUN" then
                        .Text = "GAME ADJUSTMENT #" & serviceIdx & "(" & DMDMenu(serviceIdx).StdIdx & ")"
                    Else
                        .Text = "GAME FUNCTION #" & serviceIdx
                    End If
                    .SetAlignedPosition 64,0,FlexDMD_Align_Top
                End With
                With SvcScene.GetLabel("svcl2")
                    .Text = DMDMenu(serviceIdx).Name
                    .SetAlignedPosition 64,7,FlexDMD_Align_Top
                End With
                With SvcScene.GetLabel("svcl3")
                    .Text = line3
                    .SetAlignedPosition 64,16,FlexDMD_Align_Top
                End with
                If bServiceEdit Then
                   With SvcScene.GetLabel("svcl2")
                        .ClearActions()
                        .Visible=1
                    End With
                    With SvcScene.GetLabel("svcl3")
                        .ClearActions()
                        .Visible=1
                    End With
                    DelayBlinkActor SvcScene.GetLabel("svcl3"),0.75,0.25,999
                Else
                    With SvcScene.GetLabel("svcl3")
                        .ClearActions()
                        .Visible=1
                    End With
                    With SvcScene.GetLabel("svcl2")
                        .ClearActions()
                        .Visible=1
                    End With
                    DelayBlinkActor SvcScene.GetLabel("svcl2"),0.75,0.25,999
                End If
                With SvcScene.GetLabel("svcl4")
                    If CLng(DMDStd(DMDMenu(serviceIdx).StdIdx)) = DMDMenu(serviceIdx).Deflt Then .Visible = 1 Else .Visible = 0
                End With
        End Select
    End if
    FlexDMD.UnlockRenderThread
End Sub


'  END DAPHISHBOWL - STERN DMD



'*********
'   LUT
'*********

Dim bLutActive, LUTImage
Sub LoadLUT
    Dim x
    bLutActive = False
    x = LoadValue(myGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    table1.ColorGradeImage = "LUT"&LUTImage
End Sub

Sub SaveLUT
    SaveValue myGameName, "LUTImage", LUTImage
End Sub

Sub NxtLUT:LUTImage = (LUTImage + 1)MOD 16: table1.ColorGradeImage = "LUT"&LUTImage : SaveLUT : End Sub

Sub UpdateLUT
    table1.ColorGradeImage = "LUT"&LUTImage
End Sub




'********************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first myVersion

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
Const yellowgreen = 14
Const cyan = 13
Const midblue = 12
Const ice = 11
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
'******************************************
' Change light color - simulate color leds
' changes the light color and state
' colors: red, orange, yellow, green, blue, white, purple, amber
' Note: Colors tweaked slightly to match GoT color scheme
'******************************************
' Modified to handle changing light colors while LightSeqPlayfield is running
' If it's running then
'  - If state is "-1", or sequence isn't running a coloured sequence, then change the colour right on the light
'  - Otherwise, save the colour change in the LightSaveState table and let it be restored once the sequence ends

Sub SetLightColor(n, col, stat) 'stat 0 = off, 1 = on, 2 = blink, -1= no change
    Dim color,colorfull
    Select Case col
        Case cyan
            color = RGB(2,18,18)
            colorfull = RGB(10, 224, 240)
        Case midblue
            color = RGB(2,2,18)
            colorfull = RGB(10,10,245)
        Case ice
            color = RGB(6, 12, 18)
            colorfull = RGB(130, 225, 245)
        Case red
            color = RGB(18, 2, 2)
            colorfull = RGB(245, 10, 10)
        Case orange
            color = RGB(18, 5, 2)
            colorfull = RGB(245, 64, 10)
        Case amber
            color = RGB(193, 49, 10)
            colorfull = RGB(245, 128, 10)
        Case yellow
            color = RGB(18, 18, 2)
            colorfull = RGB(245, 245, 10)
        Case darkgreen
            color = RGB(2, 8, 2)
            colorfull = RGB(2, 64, 2)
        Case green
            color = RGB(2, 16, 2)
            colorfull = RGB(10, 192, 10)
        Case blue
            color = RGB(2, 4, 18)
            colorfull = RGB(10, 40, 245)
        Case darkblue
            color = RGB(1, 8, 8)
            colorfull = RGB(4, 64, 64)
        Case purple
            color = RGB(96, 10, 96)
            colorfull = RGB(192, 30, 170)
        Case white
            color = RGB(240, 197, 143)
            colorfull = RGB(240, 230, 220)
        Case teal
            color = RGB(1, 64, 62)
            colorfull = RGB(2, 128, 126)
    Case yellowgreen
      color = RGB(6,18,2)
      colorfull = RGB(100,245,10)
    End Select
    If stat = -1 Or LightSeqPlayfield.UserValue = 0 Then n.color = color : n.colorfull = colorfull Else SavePlayfieldLightColor n,color,colorfull

    ' This is table-specific. Remove tmrFireSeq.Enabled for other tables
    If stat <> -1 And tmrFireSeq.Enabled = False Then
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

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub GiOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
    For each bulb in aBGGI
        bulb.Visible = 1
    Next
' playfield_mesh.blenddisablelighting = 1
'    ShadowGI.visible= 1
    'GameGiOn
End Sub

Sub GiOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    For each bulb in aBGGI
        bulb.Visible = 0
    Next
' playfield_mesh.blenddisablelighting = 0
'    ShadowGI.visible= 0
    'GameGiOff
End Sub

' GI, light & flashers sequence effects

Sub GiEffect(n)
    Dim ii
    LightSeqGi.StopPlay
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
        Case 6 ' blink once
            LightSeqGi.UpdateInterval = 20
            LightSeqGi.Play SeqBlinking,0,1,10
            'LightSeqGi.Play SeqAlloff
    End Select
End Sub

Sub GiIntensity(i)
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = i
    Next
    For each bulb in aFiLights
        bulb.IntensityScale = i
    Next
    ' Table-specific - adjusts intensity in Lampz
    Lampz.state(80) = i
    Lampz.state(81) = i
    Lampz.state(82) = i
' if i > 1 then
'   playfield_mesh.blenddisablelighting = 1
' else
'   playfield_mesh.blenddisablelighting = i
' end if
End Sub

' In order to be able to manipulate RGB lamp colour and fade rate during sequences,
' we have to save all of that information into arrays so it can be restored after the sequence
' (the built-in VPX light sequencer only restores lamp on/off state, not colour/fade, etc)
' Right now we only save colour and fade speed, but we could save anything

' This is called before a sequence that requires custom colours is run
Sub SavePlayfieldLightState
    Dim i,a
    i = 0
    For each a in aPlayfieldLights
        LightSaveState(i,0) = a.State
        LightSaveState(i,1) = a.Color
        LightSaveState(i,2) = a.Colorfull
        LightSaveState(i,3) = a.FadeSpeedUp
        LightSaveState(i,4) = a.FadeSpeedDown
        i = i + 1
    Next
End Sub

' This is called once at table start. It saves all lamp names into an array, whose index is equal to the number after "li"
' For this to work, all lamp objects that you want to be saved must be named "li<x>" where x is a not-zero-padded number
Sub SavePlayfieldLightNames
    Dim i,j,a
    i = 0
    On Error Resume Next
    For each a in aPlayfieldLights
        If Left(a.Name,2) = "li" Then
            j = CInt(Mid(a.Name,3))
            If j > 0 And Not Err Then LightNames(j) = i
        End If
        i = i + 1
    Next
    On Error Goto 0
End Sub

' If SetLightColor is called while a sequence is running, it calls this sub to save the desired colour into an array so it can be
' restored once the sequence ends
Sub SavePlayfieldLightColor(a,col,colf)
    On Error Resume Next
    Dim j
    If Left(a.Name,2) = "li" Then
        j = CInt(Mid(a.Name,3))
        If j > 0 And Not Err Then LightSaveState(LightNames(j),1) = col : LightSaveState(LightNames(j),2) = colf
    End If
    On Error Goto 0
End Sub

' This is called by the LightSeq_PlayDone handler to restore light states after the sequence ends
Sub RestorePlayfieldLightState(state)
    Dim i,a
    i = 0
    For each a in aPlayfieldLights
        If state Then a.State = LightSaveState(i,0)
        a.Color = LightSaveState(i,1)
        a.ColorFull = LightSaveState(i,2)
        a.FadeSpeedUp = LightSaveState(i,3)
        a.FadeSpeedDown = LightSaveState(i,4)
        i = i + 1
    Next
    For i = 0 to 69
        Lampz.FadeSpeedUp(i) = 1/2 ' Value copied from Lampz.init
        Lampz.FadeSpeedDown(i) = 1/4
    Next
End Sub

' Set Playfield lights to slow fade a color
Sub PlayfieldSlowFade(color,fadespeed)
    Dim a
    For each a in aPlayfieldLights
        If color >= 0 Then SetLightColor a,color,-1
    Next
    For a = 0 to 69 ' for x = 0 to 140 : Lampz.FadeSpeedUp(x) = 1/2 : Lampz.FadeSpeedDown(x) = 1/10 : next

        Lampz.FadeSpeedUp(a) = fadespeed
        Lampz.FadeSpeedDown(a) = fadespeed/4
    Next
End Sub


'*******************************
' Attract mode support
' (should be table-independent)
'*******************************

Sub StartAttractMode
    'StartLightSeq
    GameStartAttractMode
End Sub

Sub StopAttractMode
    GameStopAttractMode
    'LightSeqAttract.StopPlay
End Sub

Sub LightSeqAttract_PlayDone()
    RestorePlayfieldLightState False    ' Restore color and fade speed but not state. Sequencer takes care of that
End Sub

Sub LightSeqTilt_PlayDone()
    LightSeqTilt.Play SeqAllOff
End Sub

Sub LightSeqSkillshot_PlayDone()
    LightSeqSkillshot.Play SeqAllOff
End Sub

Sub LightSeqPlayfield_PlayDone()
    If LightSeqPlayfield.UserValue <> 0 Then RestorePlayfieldLightState False
    LightSeqPlayfield.UserValue = 0
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    li89.TimerEnabled = False ' Table-specific
    For each a in aLights
        a.State = 0
    Next
    UPFFlasher001.visible = 0 : UPFFlasher002.visible = 0
    fl242.state = 0 : fl237.state = 0
End Sub

'**************************************
' Non-game-specific Light sequences
'**************************************

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
        Case 5 'top down
            LightSeqPlayfield.UpdateInterval = 4
            LightSeqPlayfield.Play SeqDownOn, 15, 2
        Case 6 'down to top
            LightSeqPlayfield.UpdateInterval = 4
            LightSeqPlayfield.Play SeqUpOn, 15, 1
    End Select
End Sub


' Lamp Fader code. Right now this is only used by a couple of Flashers but could be used for any lamp
' Copied from Ghostbuster ESD table.

Dim bulb
Dim LampState(10), FadingLevel(10), chgLamp(10)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
Lamp2Timer.Interval = -1 'lamp fading speed
Lamp2Timer.Enabled = 1

Sub InitLamps()
    Dim x
    For x = 0 to 10
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
    Next
End Sub

Sub Lamp2Timer_Timer()
    Dim num, chg, ii, vid, prevlvl, targetlvl

    For ii = 0 To 10
        prevlvl = LampState(ii)   'previous frame value
        targetlvl = chgLamp(ii)         'latest frame value

        'Fadinglevel filtering
        ' 0 = static off -> skipped in subs
        ' 1 = static on -> skipped in subs
        ' 3 = not used
        ' 4 = fading down -> subs have equation that will fade the level down
        ' 5 = fading up -> This is not actually in use with vpinspa as it provides slow fade ups already

        if prevlvl <> targetlvl then
            if prevlvl < targetlvl Then
                FadingLevel(ii) = 5             'fading up...
                LampState(ii) = targetlvl           'as vpinspa output is modulated, we write the level here
            end if

            if prevlvl > targetlvl Then
                FadingLevel(ii) = 4               'fading down...
                'LampState(ii) = (prevlvl + targetlvl) / 2    'Skipping level set, let the subs to fade it slow
            end if
        Else              'no change in intensity. Vpinspa outputs these occasionally, but we should let them skip in fading subs
            if targetlvl = 255 Then
                FadingLevel(ii ) = 1 'on already
            elseif targetlvl = 0 Then
                FadingLevel(ii ) = 0 'off already
            Else
                'debug.print "same level outputted for consecutive frames -> don't care"
            end if
        end if
    Next

    UpdateLamps
End Sub


Sub SingleLamp(lampobj, id)
  dim intensity

  if FadingLevel(id) > 1 then

    Intensity = LampState(id)
'   debug.print "SingleLamp: " & intensity
    lampobj.IntensityScale = LampState(id) / 255

    if FadingLevel(id) = 4 Then
      Intensity = LampState(id) * 0.7 - 0.01
      if Intensity <= 0 then : Intensity = 0 : FadingLevel(id) = 0 : lampobj.state = 0
    end if

    if FadingLevel(id) = 5 Then
      if Intensity = 255 then FadingLevel(id) = 1
            lampobj.state = 1
    end if

    LampState(id) = Intensity
  end if
End Sub

Sub SingleLampM(lampobj, id)
    if FadingLevel(id) > 1 then
    lampobj.IntensityScale = LampState(id) / 255
        if LampState(id) > 0 Then lampobj.state=1 Else lampobj.state=0
  end If
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' lost a ball :-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one. This is not table-specific code - calls out to subs specific to this game
'
Dim bDrainHitWhileBallReleasing

Dim bBallIsDraining : bBallIsDraining = False ' True when ball is draining including follow up calls to EndOfBall and EndOfBall2

Sub Drain_Hit()
  'debug.print "**Drain_Hit"
  bBallIsDraining = True
    ' Destroy the ball
    Drain.DestroyBall
    BallsOnPlayfield = BallsOnPlayfield - 1 : If BallsOnPlayfield < 0 Then BallsOnPlayfield = 0
    ' Exit Sub ' only for debugging - this way you can add balls from the debug window

    ' pretend to knock the ball into the ball storage mech
    'PlaySoundAt "fx_drain", Drain
    RandomSoundDrain Drain

    ' Table-specific
    If bGameInPLay = False Or bMadnessMB = 1 Then
    bBallIsDraining = False
    Exit Sub 'don't do anything, just delete the ball
  end If

    'if Tilted the end Ball Mode
    If Tilted Then
        StopEndOfBallModes
        if bMultiBallMode Then StopMBmodes
    End If

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True)AND(Tilted = False)Then
        bDrainHitWhileBallReleasing = False
        ' is the ball saver active, (or ready to be activated but the ball sewered right away)
        If(bBallSaverActive = True and bEarlyEject = False) Or (bBallSaverReady = True AND DMDStd(kDMDStd_BallSave) <> 0 And bBallSaverActive = False) Then
      bBallIsDraining = False
            DoBallSaved 0
        Else
          bEarlyEject = False
            ' cancel any multiball if on last ball (ie. lost all other balls)
            ' NOTE: This is GoT-specific. Remove tmrBWMultiBallRelease check for other tables
            If BallsOnPlayfield-RealBallsInLock < 2 And bMultiBallMode Then
                If tmrBWmultiballRelease.Enabled Then
                    debug.print "Ball drained in MB mode but tmrBWMultiballRelease is enabled"
                    bDrainHitWhileBallReleasing = True
                ElseIf CreateMultiballTimer.Enabled = False Then
                    ' not in multiball mode any more
                    bMultiBallMode = False

                    ' turn off any multiball specific lights
                    ChangeGi white
                    'stop any multiball modes
                    StopMBmodes
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield-RealBallsInLock = 0 And tmrBWmultiballRelease.Enabled = False And CreateMultiballTimer.Enabled = False) Then
                ' End Mode and timers
                ChangeGi white
                ' Show the end of ball animation
                ' and continue with the end of ball
                ' DMD something?
                StopEndOfBallModes
                vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, if there is no animation then move to the end of ball
            Else
        bBallIsDraining = False
      End If
        End If
  else ' no game in progress
    bBallIsDraining = False
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    'PlaySoundAt "fx_sensor", swPlungerRest
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
    'LaunchLight.State = 2

    'be sure to update the Scoreboard after the animations, if any

    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        vpmtimer.addtimer 1000, "PlungerIM.AutoFire:DOF 120, DOFPulse:PlaySoundAt ""fx_AutoPlunger"", swPlungerRest:bAutoPlunger = False:bAutoPlunged = True '"
    End If

    SetLockbarLight

    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
        ChangeSong
        ResetSkillShotTimer.Enabled = 1
    End If
    bJustPlunged = True
    tmrJustPlunged.Interval = 2500
    tmrJustPlunged.Enabled = 1

    GameDoBallLaunched
    bAutoPlunged = False
    bBallSaved = False
    SetLockbarLight
End Sub

Sub tmrJustPlunged_Timer : bJustPlunged = False : End Sub

' Not used in this game. GoT has its own BallSaver logic
' Sub EnableBallSaver(seconds)
'     'debug.print "Ballsaver started"
'     ' set our game flag
'     bBallSaverActive = True
'     bBallSaverReady = False
'     ' start the timer
'     BallSaverTimerExpired.Interval = 1000 * seconds
'     BallSaverTimerExpired.Enabled = True
'     BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
'     BallSaverSpeedUpTimer.Enabled = True
'     ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
'     LightShootAgain.BlinkInterval = 160
'     LightShootAgain.State = 2
' End Sub

' ' The ball saver timer has expired.  Turn it off AND reset the game flag
' '
' Sub BallSaverTimerExpired_Timer()
'     'debug.print "Ballsaver ended"
'     BallSaverTimerExpired.Enabled = False
'     ' clear the flag
'     bBallSaverActive = False
'     ' if you have a ball saver light then turn it off at this point
'     LightShootAgain.State = 0
' End Sub

' Sub BallSaverSpeedUpTimer_Timer()
'     'debug.print "Ballsaver Speed Up Light"
'     BallSaverSpeedUpTimer.Enabled = False
'     ' Speed up the blinking
'     LightShootAgain.BlinkInterval = 80
'     LightShootAgain.State = 2
' End Sub






















'******************************************
' Local game code starts here
'
' Global Variables, arrays, constants come first
'
' PlayerState class: This stores the state of most global variables between balls
' House class: This is the main class. There is one instance per player. It handles all state related to your chosen house,
'              and processes hits to the 7 major shots, as well as all Upper Playfield shots. Since the 7 main shots are also
'              all jackpot shots, it handles all jackpot processing for all multiball modes
' BattleState class: Each player has 7 instances of the battlestate class - one for each house they do battle with. In battle mode
'              the battlestate class tracks all progress, including across repeated attempts.
'
' Table, game, and ball initialization, as well as end-of-ball and bonus processing comes next
'
' Then all of the switch_hit subs
'
' Next is timer management. The game has its own software-based set of timers, as well as VPX timer objects. The software-based timers
' are run from a master timer that updates 10x per second. The software timers are used anywhere that pausable timers are needed, which
' is almost everywhere (ball saver, playfield validation, battle timers, Hurry Ups, etc). The set of software timers and their associated
' subroutines are defined in the global variables at the beginning of the code
'
' Next up is lighting management, including various sequences for different events
'
' Next is Instant Info, and Attract mode. Both work on a similar principal, stepping through each state in large state machine. There's also
' a custom lighting attract mode, as this table uses RGB effects in attract mode, which the standard light sequencer can't handle
'
' Finally, all of the DMD-related code. There's also lots of DMD code sprinkled through other subs as needed, but where possible, it's
' concentrated in subs in this area. These are all subs for generating scenes, such as jackpot, shot hits, score scenes, etc. The actual code
' for putting scenes on the DMD is handled in the generic code at the top of this script
'
'
'
'
'
'
'
'
'
'******************************************

Const Stark = 1
Const Baratheon = 2
Const Lannister = 3
Const Greyjoy = 4
Const Tyrell = 5
Const Martell = 6
Const Targaryen = 7


' Constants we might need to tweak later
Const SpinnerAddValue = 500      ' Base amount that Spinner's value increases by for each target hit.

' Global table-specific variables
Dim HouseColor
Dim HouseSigil
Dim HouseShield
Dim HouseAbility
Dim LoLLights
Dim ComboLaneMap
Dim ComboLights
Dim GoldTargetLights
Dim pfmuxlights
Dim BattleObjectives
Dim BattleObjectivesShort
Dim QAnimateTimes       ' Length of each qualifying shot's animation time
Dim SwordNames
Dim bMadnessMB
Dim bGameJustPlayed : bGameJustPlayed = false

' Global variables with player data - saved across balls and between players
Dim PlayerMode          ' Current player's mode. 0=normal, -1 = select house, -2 = select battle, -3 = select mystery, 1 = in battle, 2 = WIC
Dim SelectedHouse       ' Current Player's selected house
Dim bTopLanes(2)        ' State of top lanes
Dim LoLTargetsCompleted ' Number of times the target bank has been completed
Dim WildfireTargetsCompleted ' Number of times wildfire target bank has been completed
Dim BWMultiballsCompleted
Dim bBlackwaterSJPMode
Dim WallMBCompleted
Dim WallMBLevel
Dim bWallMBReady
Dim WallJPValue
Dim bHotkMBReady
Dim bHotkMBDone
Dim bITMBReady
Dim bITMBDone
Dim bITMBActive
Dim ITScore
Dim bLockIsLit
Dim EBisLit
Dim bPictoEBAwarded
Dim bWildfireTargets(2) ' State of Wildfire targets
Dim bLoLLit             ' Whether Lord of Light Outlanes are lit
Dim bLoLUsed            ' Whether Lord of Light has been used this game
Dim bEarlyEject         ' Indicates whether outlanes caused an early eject of a new ball
Dim CompletedHouses     ' Number of completed houses - determines max spinner level and triggers HOTK and Iron Throne modes
Dim TotalGold           ' Total gold collected in the game
Dim CurrentGold         ' Current gold balance
Dim TotalWildfire
Dim CurrentWildfire
Dim SwordMask
Dim SwordsCollected
Dim CastlesCollected
Dim bGoldTargets(5)
Dim bTargaryenInProgress
Dim bMysteryLit
Dim bSwordLit

' Support for game timers
Dim GameTimeStamp       ' Game time in 1/10's of a second, since game start
Dim bGameTimersEnabled  ' Flag for whether any timers are enabled
Dim TimerFlags(30)      ' Flags for each timer's state
Dim TimerTimestamp(30)  ' Each timer's end timestamp
Dim TimerSubroutine     ' Names of subroutines to call when each timer's time expires

' Timers 1 - MaxUnfreezableTimerIndex (was originally 4) can't be frozen.
' if adding more unfreezable timers, put them next, and adjust MaxUnfreezableTimerIndex -- needed for tmrGame_Timer sub and targaryan timers
TimerSubroutine = Array("","UpdateChooseBattle","PreLaunchBattleMode","LaunchBattleMode","UpdateBattleMode","MysteryAwardTimer", "TargaryenFreezeUpdateDMD", _
                        "BattleModeTimer1","BattleModeTimer2", _
                        "MartellBattleTimer","HurryUpTimer","ResetComboMultipliers","ModePauseTimer","BlackwaterSJPTimer","WildfireModeTimer", _
                        "UPFMultiplierTimer","PFMStateTimer","PFMultiplierTimer","BallSaveTimer","BallSaverSpeedUpTimer","WHCTimer","LoLSaveTimer","TargaryenFreezeTimer", _
                        "BallSaverGraceTimer")
Const MaxUnfreezableTimerIndex = 6 ' the timers at or below this value are unfreezable

Const tmrUpdateChooseBattle = 1   ' Update DMD timers during Choose Your Battle
Const tmrChooseBattle   = 2         ' Countdown timer for choosing your Battle
Const tmrLaunchBattle   = 3         ' After battle is chosen, countdown to launch while scenes play. Can be aborted with flippers
Const tmrUpdateBattleMode = 4     ' Update DMD during battle mode
Const tmrMysteryAward   = 5
Const tmrTargaryenFreezeUpdateDMD = 6 ' Updates the DMD when using this ability (countdown)
' put unfreezeable timers above this, freezable beneath
Const tmrBattleMode1    = 7         ' "Top" house battle countdown timer
Const tmrBattleMode2    = 8         ' "Bottom" house battle countdown timer
Const tmrMartellBattle  = 9       ' 10-second timer for Martell orbits
Const tmrHurryUp        = 10        ' When HurryUp is active, update DMD 5 times/second
Const tmrComboMultplier = 11      ' Timeout timer for Combo multipliers
Const tmrModePause      = 12        ' "no activity" timer that will pause battle mode timers if it elapses
Const tmrBlackwaterSJP  = 13      ' SuperJackpot battering ram countdown timer
Const tmrWildfireMode   = 14        ' Wildfire Mini Mode timer
Const tmrUPFMultiplier  = 15
Const tmrPFMState       = 16
Const tmrPFMultiplier   = 17
Const tmrBallSave       = 18
Const tmrBallSaveSpeedUp= 19
Const tmrWinterHasCome  = 20      ' 1 second update timer for WinterHasCome mode
Const tmrLoLSave        = 21
Const tmrTargaryenFreeze= 22
Const tmrBallSaverGrace = 23

Const MaxTimers         = 23      ' Total number of defined timers. There MUST be a corresponding subroutine for each

'HurryUp Support
Dim HurryUpValue
Dim bHurryUpActive
Dim HurryUpCounter
Dim HurryUpGrace
Dim HurryUpScene
Dim HurryUpChange
Dim TGHurryUpValue
Dim bTGHurryUpActive
Dim TGHurryUpCounter
Dim TGHurryUpGrace
Dim TGHurryUpScene
Dim TGHurryUpChange

' Player state data
Dim House(4)  ' Current state of each house - some house modes aren't saved, while others are. May need a Class to save detailed state
Dim PlayerState(4) ' Structure to save global player-specific variables across balls

'Other
Dim bBattleInstructionsDone ' Whether Battle selection instructions have been shown this game

' Ball-specific variables (not saved across balls)
Dim PlayfieldMultiplierVal
Dim SpinnerValue
Dim AccumulatedSpinnerValue ' Amount that has accumulated in the last 2 seconds
Dim SpinnerLevel
Dim DroppedTargets      ' Number of targets dropped
Dim ComboMultiplier(5)
Dim bWildfireLit
Dim bPlayfieldValidated
Dim bElevatorShotUsed   ' Whether a shot to the upper playfield via the right orbit has been made this ball yet or not
Dim bCastleShotAvailable ' Whether the ball has just been plunged to the upper PF
Dim HouseBattle1        ' When in battle, the primary (top) House
Dim HouseBattle2        ' When in two-way battle, the second House
Dim PFMState
Dim bAddABallUsed
Dim bInlanes(2)
Dim bBallReleasing      ' Track whether a ball is in the process of being released from the ramp
Dim MultiballBallsLocked
Dim bBWMultiballActive
Dim bCastleMBActive
Dim bWallMBActive
Dim bHotkMBActive
Dim bWHCMBActive
Dim bHotkIntroRunning
Dim WHCMBScore
Dim BlackwaterScore
Dim CastleMBScore
Dim WallMBScore
Dim HotkScore
Dim MMTotalScore
Dim ITHighScore(4)
Dim HotkHighScore(4)
Dim WHCHighScore(4)
Dim bMysterySJPMode
Dim RandomHouse

'debug timer
sub tmrObserveGameTimers_timer
  debug.print "***GameTimeStamp:" & GameTimeStamp
  debug.print "***BallSaveTimer:" & TimerTimestamp(tmrBallSave) & " BallSaveSpeedup:" & TimerTimestamp(tmrBallSaveSpeedUp) & " BallSaveGrace:" & TimerTimestamp(tmrBallSaverGrace)

  'For i = 1 to MaxTimers
     '   TimerFlags(i) = TimerFlags(i) And 253
    'Next
end sub

' Handles transitions between PlayerModes
Dim bReadyForAction : bReadyForAction = True

Sub SetReadyForAction(value)
  'debug.print "Setting bReadyForAction from " & bReadyForAction & " to " & value
  bReadyForAction = value
end Sub

Sub SetPlayerMode(value)
  'handles all locations where PlayerMode is changed
  'modes: 0=normal, -1=select house, -2=select battle, -2.1=prelaunch, -3=select mystery, 1=in battle, 2=WIC
  debug.print "Adjusting PlayerMode from " & PlayerMode & " to " & value
  if PlayerMode=0 and value=-2 Then
    'changing from normal to house battle
    SetReadyForAction False
    'todo: add timer to set it back as failsafe? shouldn't be needed
    'vpmtimer.Addtimer 2000, "SetReadyForAction True '"
  end if
  PlayerMode = value
end Sub



HouseColor = Array(white,white,yellow,red,purple,green,amber,blue)
' Assignment of centre playfield shields
HouseSigil = Array(li38,li38,li41,li44,li47,li50,li53,li32)
' Assignment of "shot" shields. Last 3 are Upper PF target lights
HouseShield = Array(li141,li141,li26,li114,li86,li77,li156,li98,li189,li192,li195)
' House Ability strings, used during House Selection
HouseAbility = Array("","INCREASE WINTER IS COMING","ADVANCE WALL MULTIBALL","COLLECT MORE GOLD","PLUNDER RIVAL ABILITIES","INCREASE HAND OF THE KING","ACTION BUTTON=ADD A BALL","FREEZE TIMERS","THE GAME WILL CHOOSE")

BattleObjectives = Array("", _
            "ARYA  BECOMES  AN  ASSASSIN"&vbLf&"RAMPS  BUILD  VALUE"&vbLf&"ORBITS  COLLECT  VALUE", _
            "STANNIS  VS  THE  WILDLINGS"&vbLf&"SPINNER  BUILDS  VALUE"&vbLf&"COLLECT  AT  THE  3  TARGETS", _
            "BRING  BACK  MYRCELLA"&vbLf&"GOLD  TARGETS  LIGHT  RED  SHOTS"&vbLf&"5  RED  SHOTS  TO  FINISH", _
            "GREYJOY  TAKES  WINTERFELL"&vbLf&"5  SHOTS  TO  FINISH"&vbLf&"TIMER  RESETS  AFTER  EACH", _
            "LORD  LORAS  JOUSTING  THE  MOUNTAIN"&vbLf&"TWO  BANK  WILL  SCORE  HITS"&vbLf&"SCORE  3  HITS  TO  WIN", _
            "VIPER  VERSUS  THE  MOUNTAIN"&vbLf&"SHOOT  3  ORBITS  IN  A  ROW"&vbLf&"LEFT  RAMP  COLLECTS",_
            "DEFEAT  VISERION"&vbLf&"SHOOT  3  HURRY  UPS"&vbLf&"TO  DEFEAT  VISERION", _
            "DEFEAT  RHAEGAL"&vbLf&"SHOOT  5  HURRY  UPS"&vbLf&"TO  DEFEAT  RHAEGAL", _
            "DEFEAT  DROGON"&vbLf&"SHOOT  3  HURRY  UPS"&vbLf&"TO  DEFEAT  DROGON")

BattleObjectivesShort = Array("","ARYA'S TRAINING","AID FOR THE WALL","SAVE MYRCELLA","WINTERFELL BURNS",_
            "JOUSTING","TRIAL BY COMBAT","DEFEAT VISERION","DEFEAT RHAEGAL","DEFEAT DROGON")

' Length of each scene in the qualifying shot animations. Arranged as (House# x 3) + HitNumber
QAnimateTimes = Array(0,0,0,0,1,1,1.8,3,1,1,2,1.5,2,2.5,2.5,2.5,1.3,2.6,0,4.8,1.3,2.2,1.6,3.9,3.1)

SwordNames = Array("NEEDLE","ICE","OATHKEEPER","LONGCLAW","DARK SISTER","LIGHTBRINGER","HEARTEATER","WIDOW'S WAIL")

' Assignment of Lol Target lights
LoLLights = Array(li17,li20,li23)
'Assignment of Gold target lights
GoldTargetLights = Array(li92,li105,li120,li135,li147)
' Map of house name to combo lane (Greyjoy is combo lane1, Targaryen is Combo lane2, etc)
ComboLaneMap = Array(0,4,0,3,1,0,5,2)

ComboLights = Array(li89,li89,li101,li117,li144,li159)

pfmuxlights = Array(li56,li59,li62,li65)

' Upper PF lights
Dim UPFLights
UPFLights = Array(li186,li186,li189,li180,li192,li183,li195,li198,li210,li207,li204,li201,li213,li216)

' This class holds player state that is carried over across balls
Class cPState
    Dim bWFTargets(2)
    Dim WFTargetsCompleted
    Dim LTargetsCompleted
    Dim myLoLLit
    Dim myMysteryLit
    Dim mySwordsLit
    Dim myLoLUsed
    Dim myLockIsLit
    Dim myEBisLit
    Dim myPictoEBAwarded
    Dim myBWMultiballsCompleted
    Dim myWallMBCompleted
    Dim myWallMBLevel
    Dim myWallMBReady
    Dim myHotkMBDone
    Dim myITMBReady
    Dim myITMBDone
    Dim myITMBActive
    Dim myITScore
    Dim myWallJPValue
    Dim myBallsInLock
    Dim myGoldTargets(5)
    Dim myTotalGold
    Dim myCurrentGold
    Dim myTotalWildfire
    Dim mySwordMask         ' Bitmask of which swords have been collected
    Dim mySwordsCollected
    Dim myCastlesCollected
    Dim myCurrentWildfire
    Dim myTargaryenInProgress
    Dim myPFMState
    Dim myBonusHeld

    Public Sub Save
        Dim i
        If BWMultiballsCompleted < 2 Then
            bWFTargets(0) = bWildfireTargets(0):bWFTargets(1) = bWildfireTargets(1)
        Else
            bWFTargets(0) = 0 : bWFTargets(1) = 0
        End If
        WFTargetsCompleted = WildfireTargetsCompleted
        LTargetsCompleted = LoLTargetsCompleted
        myLoLLit = bLoLLit
        myLoLUsed = bLoLUsed
        myLockIsLit = bLockIsLit
        myEBisLit = EBisLit
        myPictoEBAwarded = bPictoEBAwarded
        myBWMultiballsCompleted = BWMultiballsCompleted
        myWallMBCompleted = WallMBCompleted
        myWallMBLevel = WallMBLevel
        myWallMBReady = bWallMBReady
        myHotkMBDone = bHotkMBDone
        myITMBReady = bITMBReady
        myITMBDone = bITMBDone
        myITMBActive = bITMBActive
        myITScore = ITScore
        myBallsInLock = BallsInLock
        myWallJPValue = WallJPValue
        myTotalGold = TotalGold
        myCurrentGold = CurrentGold
        myTotalWildfire = TotalWildfire
        myCurrentWildfire = CurrentWildfire
        mySwordsCollected = SwordsCollected
        myCastlesCollected = CastlesCollected
        myTargaryenInProgress = bTargaryenInProgress
        mySwordsLit = bSwordLit
        myMysteryLit = bMysteryLit
        mySwordMask = SwordMask
        myBonusHeld = bBonusHeld
        If PFMState = 2 Then myPFMState = 2 Else myPFMState = 0
        For i = 0 to 5:myGoldTargets(i) = bGoldTargets(i):Next
    End Sub

    Public Sub Restore
        Dim i
        bWildfireTargets(0) = bWFTargets(0):bWildfireTargets(1) = bWFTargets(1)
        WildfireTargetsCompleted = WFTargetsCompleted
        LoLTargetsCompleted = LTargetsCompleted
        bLoLLit = myLoLLit
        bLoLUsed = myLoLUsed
        bLockIsLit = myLockIsLit
        EBisLit = myEBisLit
        bPictoEBAwarded = myPictoEBAwarded
        BWMultiballsCompleted = myBWMultiballsCompleted
        WallMBCompleted = myWallMBCompleted
        WallMBLevel = myWallMBLevel
        bWallMBReady = myWallMBReady
        bHotkMBDone = myHotkMBDone
        bITMBReady = myITMBReady
        bITMBDone = myITMBDone
        bITMBActive = myITMBActive
        ITScore = myITScore
        WallJPValue = myWallJPValue
        BallsInLock = myBallsInLock
        SwordsCollected = mySwordsCollected
        TotalGold = myTotalGold
        TotalWildfire = myTotalWildfire
        CurrentWildfire = myCurrentWildfire
        bTargaryenInProgress = myTargaryenInProgress
        PFMState = myPFMState
        bMysteryLit = myMysteryLit
        bSwordLit = mySwordsLit
        bBonusHeld = myBonusHeld


        CurrentGold = myCurrentGold
        CastlesCollected = myCastlesCollected
        For i = 0 to 5:bGoldTargets(i) = myGoldTargets(i):Next

    End Sub
End Class

Dim BWExplosionTimes
BWExplosionTimes = Array(1,1,2.9,2.8,3.6,1)

' This class holds everything to do with House logic, which is basically almost all of the rules of game
' Only 7 shots on the main playfield are involved, but Upper PF shots are also handled in this class.
' Multiball start/end is handled outside, but the jackpot shots for multiball are also all tracked in this class,
' as well as all wizard modes
' There is one instance of this class per player.
Class cHouse
    Dim bSaid(7)             ' Whether the house's name has been said yet during ChooseHouse state
    Dim SayTimeStamp         ' Timestamp of last Choose House announcement
    Dim bQualified(7)        ' Whether the house has qualified for battle
    Dim bCompleted(7)        ' Whether battle has been completed
    Dim MyBattleState(7)      ' Placeholder for current battle state
    Dim QualifyCount(7)     ' Count of how many times the qualifying shot has been made for each house
    Dim HouseSelected
    Dim QualifyValue        ' Hold the current value for a qualifying target hit
    Dim bBattleReady
    Dim BWJackpotShots(7)     ' Track how many shots are needed for each jackpot shot
    Dim BWJackpotLevel
    Dim BWSJPLevel
    Dim BWSJPBaseValue
    Dim BWJackpotValue
    Dim BWState
    Dim UPFState            ' Upper PF State
    Dim UPFLevel            ' Progress towards Castle Multiball
    Dim UPFShotMask         ' Lit shots on the Upper PF
    Dim UPFCastleShotMask   ' Saved shot mask state for standard mode
    Dim UPFBattleShotMask   ' To track shots in Battle mode
    Dim UPFMultiplier
    Dim UPFSJP
    Dim CMBJackpot
    Dim CMBJPIncr
    Dim CMBSJPValue
    Dim WallMBShotCount
    Dim WallMBState
    Dim WallJPIncr
    Dim WHCMBState
    Dim WHCMBShots
    Dim WHCJPValue
    Dim WHCJPIncr
    Dim WHCHordeTimestamp   ' Timer for Horde mode
    Dim WHCShotTimestamp(7) ' Timer for each shot in Lieutenant mode
    Dim WiCValue            ' Current WiC HurryUp Value
    Dim WiCTotal            ' Total accumulated WiC HurryUp
    Dim WiCShots            ' Completed Iced Over shots
    Dim WiCs                ' Completed WiC HurryUps (countdown to Winter Has Come MB)
    Dim WiCMask             ' Mask of which shots have completed WiCs
    Dim CurrentWiCShot      ' If a WiC HurryUp is active, which shot is flashing. 0 if not active
    Dim CurrentWiCShotCombo ' Combo multiplier on shot that WiC HurryUp started on, when it started
    Dim ActionAbility       ' The house ability you currently have
    Dim ActionButtonUsed    ' Whether the ability has been used on this ball
    Dim HotkMask
    Dim HotkShotValue
    Dim HotkSJPValue
    Dim HotkSuperSJPValue
    Dim HotkLevel
    Dim HotkState
    Dim ITState             ' Current state of Iron Throne mode, for state machine
    Dim ITCastleMask        ' Bitmask of which castles are still to do
    Dim ITActiveCastle      ' Currently active castle
    Dim ITCastlesCollected  ' Count of castles finished
    Dim ITExtraCastlesCollected ' Count of extra castles finished, after the 7 main ones
    Dim ITLastShotTimeStamp ' Time stamp of last shot hit

    Private Sub Class_Initialize(  )
    dim i
    ResetHouse
        HouseSelected = 0
        bBattleReady = False
        LockWall.collidable = False : Lockwall.Uservalue = 0 : LockPost False : MoveSword True
        UPFState = 0
        UPFLevel = 1
        UPFMultiplier = 1
        UPFShotMask = 42    ' Shots 1, 3 and 5
        UPFCastleShotMask = 42
        WiCs = 0 ' For testing Winter Has Come
        ActionButtonUsed = 0
  End Sub

    ' This sub sets variables for the beginning of a game, but also after Iron Throne completes
    Private Sub ResetHouse
        Dim i
        For i = 0 to 7
      bQualified(i) = False
            bCompleted(i) = False
            bSaid(i) = False
            QualifyCount(i) = 0
            Set MyBattleState(i) = New cBattleState
            MyBattleState(i).SetHouse = i
    Next
        QualifyValue = 100000
        WicTotal = 0
        WicShots = 0
        WiCs = 0
        WiCMask = 0
        HotkMask = 0
    End Sub

    Public Property Let MyHouse(h)
        HouseSelected = h
        ActionAbility = h


        '
        ' Action buttons
        '  √ Stark: Dire Wolf: Only active during Battle - finishes current HouseBattle1 and scores 5M
        '  √ Barahteon: Red Woman - light Lord Of Light for a few seconds only. Can be used once per ball
        '  √ Lannister: Buy playfield X's for successively more gold
        '  √ Greyjoy: Plunders other houses' action button
        '  √ Tyrell: Iron Bank - sell all X's for gold
        '  √ Martell: Add-A-Ball
        '  √ Targaryen: Freeze timers for 15s (except ball save timer). Timers started during freeze don't start until after
        If h = Tyrell Then bInlanes(0) = True : SetInlaneLights

        if bITMBDone Then Exit Property

        bQualified(h) = True
        QualifyCount(h) = 3

        If h = Lannister Then AddGold 750 Else TotalGold = 100 : CurrentGold = 100
        If h = Baratheon Then AdvanceWallMultiball 2 : WallJPValue = 9250000 Else WallJPValue = 3400000
        If Tilted=False Then PlaySoundVol "say-"&HouseToString(h)&"motto",VolCallout

        If h = Greyjoy Then
            bCompleted(h) = True
        ElseIf h = Targaryen Then
            Select Case DMDStd(kDMDFet_TargaryenHousePower)
                Case 1: MyBattleState(Targaryen).State = 5 : BattleReady = True : Score(CurrentPlayer) = Score(CurrentPlayer) + 30000000
                Case 2: MyBattleState(Targaryen).State = 10 : BattleReady = True : Score(CurrentPlayer) = Score(CurrentPlayer) + 60000000
                Case 3: bCompleted(Targaryen) = True
            End Select
        Else
            BattleReady = True
        End If
    End Property
  Public Property Get MyHouse : MyHouse = HouseSelected : End Property

    Public Property Let BattleReady(e)
        if bMultiBallMode = False Then
            bBattleReady = e
            if (e) Then
                LockWall.collidable = True : Lockwall.Uservalue = 1 : LockPost True
                SetLightColor li108,white,2
            ElseIf RealBallsInLock = 0 And bLockIsLit = False And PlayerMode <> -2 Then
                LockWall.collidable = False : Lockwall.Uservalue = 0 : LockPost False
            End if
            If e=False And bITMBReady = False and bHotkMBReady = False then li108.State=0
        End If
    End Property

    Public Property Get Qualified(h) : Qualified = bQualified(h) : End Property
    Public Property Let Qualified(h,v) : bQualified(h) = v : End Property
    Public Property Get Completed(h) : Completed = bCompleted(h) : End Property
    Public Property Get BattleState(h) : Set BattleState = MyBattleState(h) : End Property
    Public Property Get BWJackpot : BWJackpot = BWJackpotValue : End Property
    Public Property Let BWJackpot(v) : BWJackpotValue = v : End Property

    Public Sub SetActionAbility(h) : ActionAbility = h : ActionButtonUsed = 0 : End Sub

    ' Reset any state that starts over on a new ball
    Public Sub ResetForNewBall
        Dim i,t
        t = bTargaryenInProgress
        For i = 1 to 7
            If bQualified(i) and bCompleted(i) = False then t = True
        Next
        If HasAbility(Tyrell) Then bInlanes(0) = True
        BattleReady = t
        SetUPFState False
        UPFMultiplier = 1
        If ActionAbility <> Lannister Then ActionButtonUsed = False
        If bITMBActive Then DMDCreateITMBScoreScene ITState,ITActiveCastle,7-ITCastlesCollected
    End Sub

    Public Sub ResetAfterIronThrone
        ResetHouse
        Me.MyHouse = SelectedHouse
        LoLTargetsCompleted = 0
    End Sub

    Public Function HasActionAbility
        Dim i
        HasActionAbility = ActionAbility
        Select Case ActionAbility
            Case Stark: If ActionButtonUsed Or PlayerMode <> 1 Then HasActionAbility = 0
            Case Baratheon: If ActionButtonUsed Or bLoLLit Then HasActionAbility = 0
            Case Lannister: If ActionButtonUsed >= DMDStd(kDMDFet_LannisterButtonsPerGame) Or CurrentGold < ((PlayfieldMultiplierVal+1) * 600) Then HasActionAbility = 0
            Case Tyrell
                HasActionAbility = 0
                If ActionButtonUsed Then Exit Function
                For i = 1 to 5
                    If ComboMultiplier(i) > 1 Then HasActionAbility = Tyrell : Exit Function
                Next
            Case Targaryen: If ActionButtonUsed Then HasActionAbility = 0
            Case Martell: If ActionButtonUsed Or (bMultiBallMode = False And bITMBActive = False) Or bMadnessMB  <> 0 Then HasActionAbility = 0
        End Select
    End Function

    ' Say the house name. Include "house " if not said before
    Public Sub Say(h)
        Dim tmp
        If h > 7 Or (GameTimeStamp < SayTimeStamp) then Exit Sub
        if (bSaid(h)=2) Then tmp="" : SayTimeStamp = GameTimeStamp+7 Else tmp="house-":bSaid(h)=2 : SayTimeStamp=GameTimeStamp+11
        PlaySoundVol "say-" & tmp & HouseToString(h) & "1", VolCallout
    End Sub

    Public Sub StopSay(h)
        Dim tmp
        If h > 7 then Exit Sub
        if bSaid(h)=0 Then Exit Sub
        if (bSaid(h)=2) Then tmp="" Else tmp="house-"
        StopSound "say-" & tmp & HouseToString(h) & "1"
        bSaid(h) = 2
    End Sub

    ' Set the shield/sigil lights to the houses' current state
    Public Sub ResetLights
        If HouseSelected = 0 Then Exit Sub       ' Do nothing if we're still in Choose House mode
        SetModeLights   ' If in regular mode, this disables the timers on the shield lights

        If PlayerMode = 2 Then SetLightColor HouseShield(CurrentWiCShot), ice, 2 : Exit Sub

        SetLightColor li108,white,0
        Dim i
        Dim j
        j=0
        ' Set the Sigil (center playfield) lights
        For i = Stark to Targaryen
            HouseSigil(i).BlinkInterval = 100
            If bITMBActive Then
                If (ITCastleMask And 2^i) = 0 Then SetLightColor HouseSigil(i),HouseColor(i),2 : HouseSigil(i).BlinkInterval = 66
            ElseIf PlayerMode = 1 Then
                If i = HouseBattle1 or i = HouseBattle2 Then SetLightColor HouseSigil(i),HouseColor(i),2 Else HouseSigil(i).State = 0
            Else
                If bCompleted(i) Then
                    SetLightColor HouseSigil(i),HouseColor(HouseSelected),1
                    j = j + 1
                ElseIf bCompleted(i) = False and (bQualified(i)) Then
                    SetLightColor HouseSigil(i),HouseColor(i),2
                Else
                    HouseSigil(i).State = 0
                End If
            End If
        Next
        If bITMBActive Then SetLightColor li35,cyan,2
        If bITMBActive Or bMultiBallMode Or PlayerMode = 1 Then Exit Sub

        CompletedHouses = j

        ' Set the Shield (shot) lights
        For i = Stark to Targaryen
            If (bQualified(i) or bCompleted(i)) Then
                If (WiCMask And 2^i) > 0 Then
                    SetLightColor HouseShield(i),HouseColor(HouseSelected),1
                Else
                    SetLightColor HouseShield(i), ice, 1
                End If
            Else
                SetLightColor HouseShield(i),HouseColor(i),1
            End If
        Next

        if bBattleReady Then SetLightColor li108,white,2
        If bHotkMBDone Then SetLightColor li29,white,1
    End Sub

    ' Set Shield (Shot) lights in multiball mode
    Public Sub SetShieldLights
        Dim i,j,clr
        If bBWMultiballActive Then
            clr = yellowgreen
        Elseif bCastleMBActive Or bITMBActive Then clr = cyan
        ElseIf bWallMBActive Then clr = purple
        ElseIf bWHCMBActive Then clr = ice
        ElseIf bHotkMBActive Then clr = white
        End If
        For i = 1 to 7
            ModeLightState(i,0) = 1
            ModeLightState(i,1) = 0
            if (bBWMultiballActive Or bCastleMBActive or bWallMBActive or bWHCMBActive or bHotkMBActive Or bITMBActive) And BWJackpotShots(i) > 0 Then
                If bITMBActive And i = ITActiveCastle And ITState = 1 Then ModeLightState(i,1) = HouseColor(i) Else ModeLightState(i,1) = clr
                ModeLightState(i,2) = 0
                ModeLightState(i,0) = 2
            End If
            If PlayerMode = -2.1 or PlayerMode = 1 Then
                If i = HouseBattle1 or i = HouseBattle2 Then
          SetLightColor HouseSigil(i),HouseColor(i),2
        Else
          HouseSigil(i).State = 0
        end if
            End if
        Next

        If HouseBattle1 > 0 Then MyBattleState(HouseBattle1).SetBattleLights
        If HouseBattle2 > 0 Then MyBattleState(HouseBattle2).SetBattleLights
        SetUPFLights
    End Sub


    Public Sub SetBWJackpots
        BWJackpotLevel = 1
        BlackwaterScore = 0
        BWState = 1
        SetUPFState True
        SetJackpots
    End Sub

    Public Sub SetJackpots
        Dim i,j
        For i = 1 to 7 : BWJackpotShots(i) = 0 : Next
        If bWallMBActive Then
            WallMBShotCount = 0
            Select Case (WallMBState MOD 4)
                Case 0: BWJackpotShots(lannister) = 3 : BWJackpotShots(Stark) = 3
                Case 2: BWJackpotShots(Greyjoy) = 3 : BWJackpotShots(Martell) = 3
                Case 1,3: BWJackpotShots(Targaryen) = 1
            End Select
        ElseIf bWHCMBActive Then
            If WHCMBState = 1 Then
                ' Horde mode - all shots lit
                For i = 1 to 7 : BWJackpotShots(i) = 1 : Next
            Else
                'Lieutenant mode - 3 random shots lit with timers set
                j = 0
                Do
                    i = RndNbr(7)
                    If i <> Baratheon and i <> Tyrell And BWJackpotShots(i) = 0 Then
                        BWJackpotShots(i) = 1 : WHCShotTimestamp(i) = GameTimeStamp + 150 : j=j+1
                    End If
                Loop While j < 3
            End If
        ElseIf bHotkMBActive Then
            If (HotkMask And 2^Martell) > 0 Then j=2 Else j=1
            For i = 1 to 7
                If HotkState = 2 Then
                    BWJackpotShots(i) = 99
                ElseIf HotkState = 0 And ((HotkMask And 2^i) > 0 Or (HotkMask And 2^Baratheon) > 0 ) Then BWJackpotShots(i) = j
                End If
            Next
        ElseIf bITMBActive Then
            If ITCastlesCollected < 7 Then
                For i = 1 to 7
                    If (ITState = 0 And (ITCastleMask And 2^i) > 0) Or ITState > 0 Then
                        BWJackpotShots(i) = 1
                        If i = Baratheon Then ResetDropTargets
                    End if
                Next
            Else
                j = 0
                Do
                    i = RndNbr(7)
                    If  BWJackpotShots(i) = 0 Then
                        BWJackpotShots(i) = 1 : j=j+1
                        If i = Baratheon Then ResetDropTargets
                    End If
                Loop While j < 3
            End If
        Else ' Blackwater MB
            For i = 1 to 7
                If i <> Baratheon and i <> Tyrell Then BWJackpotShots(i) = 1
            Next
        End If
        SetModeLights
    End Sub

    Public Sub ResetForWallMB
        WallMBState = 0
        WallJPIncr = 300000
        If HouseSelected = Baratheon Then WallJPIncr = 900000
        SetJackpots
    End Sub

    Public Sub ResetForWHCMB
        WHCMBState = 1
        WHCJPValue = 1825000 : WHCJPIncr = 525000
        WHCHordeTimestamp = GameTimeStamp + 300
        SetJackpots
    End Sub

    Public Sub AddWHCShot
        Dim i, j : j = True
        For i = 1 to 7
            If BWJackpotShots(i) = 0 Then j = False
        Next
        If j then Exit Sub  ' All shots lit, can't add another
        Do
            i = RndNbr(7)
            If (i <> Baratheon and i <> Tyrell And BWJackpotShots(i) = 0) Or (WHCMBState = 1 And BWJackpotShots(i) = 0) Then
                BWJackpotShots(i) = 1 : WHCShotTimestamp(i) = GameTimeStamp + 150 : j=True
            End If
        Loop While j = False
        SetModeLights
    End Sub

    Public Sub ResetForHotkMB
        If (HotkMask And 2^Tyrell) > 0 Then HotkSJPValue = 15000000 Else HotkSJPValue = 0
        If (HotkMask And 2^Targaryen) > 0 Then HotkShotValue = 3000000 Else HotkShotValue = 2500000
        If (HotkMask And 2^Greyjoy) > 0 Then HotkLevel = 2 Else HotkLevel = 1
        HotkState = 0
        UpdateHotkScoreScene
        If (HotkMask And 2^Lannister) > 0 Then HotkSuperSJPValue = 200000000 Else HotkSuperSJPValue = 100000000
        SetJackpots
    End Sub

    Public Sub ResetForITMB
        ITState = 0
        ITCastleMask = 254
        ITCastlesCollected = 0
        ITExtraCastlesCollected = 0
        SetJackpots
    End Sub


    'From 1-5 WF, ~100,000 per WF. From 6-25, ~50K per WF, 25-45 ~45K per WF, 45-55, 30K per, and 55+, 10k per. Max is 3M
    Public Sub AddWildfire(wf)
        Dim i,j,k
        k = 1000 ' k = random range, in 10s (1000=10K); j=base value, in 10s
        If BWJackpotValue >= 3000000 Then CurrentWildfire = CurrentWildfire + wf: Exit Sub
        For i = 1 to wf
            CurrentWildfire = CurrentWildfire + 1
            If CurrentWildfire < 6 Then
                j=9500
            Elseif CurrentWildfire < 25 Then
                j = 4500
            ElseIf CurrentWildfire < 45 Then
                j = 4000
            ElseIf CurrentWildfire < 55 Then
                j = 2500
            Else
                j = 750:k=500
            End If
            BWJackpotValue = BWJackpotValue + 10*int(j+RndNbr(k))
            If BWJackpotValue > 3000000 Then BWJackpotValue = 3000000
        Next
    End Sub

    Public Sub IncreaseBWJackpotLevel
        Dim i
        BWState = BWState + 1
        BWJackpotLevel = BWJackpotLevel + 1
        For i = 1 to 7
            If i <> Baratheon and i <> Tyrell Then BWJackpotShots(i) = BWJackpotLevel
        Next
        UpdateBWMBScene
        SetModeLights
    End Sub

    ' m = UPF multiplier. 0 if not a UPF shot
    Public Sub ScoreSJP(m)
        Dim mylevel,myvalue
        If bMysterySJPMode > 0 Then mylevel = bMysterySJPMode Else mylevel = BWJackpotLevel
        If bMysterySJPMode > 0 And BWJackpotValue < 1000000 Then myvalue = 1000000 Else myvalue = BWJackpotValue

        DMDBlackwaterSJPScene FormatScore((myvalue*mylevel*6 + 10000000)*(PlayfieldMultiplierVal+m)),mylevel
        AddScoreNoX ((myvalue*mylevel*6+10000000)*(PlayfieldMultiplierVal+m))
        If bBWMultiballActive Then BlackwaterScore = BlackwaterScore + ((myvalue*mylevel*6+10000000)*(PlayfieldMultiplierVal+m))
    End Sub

    ' Main function that handles processing the 7 main shots in the game.
    ' These shots have 3 major modes of operation
    '  - qualifying a house for battle
    '  - after qualifying, advancing Winter Is Coming
    '  - during House Battle mode, completing battles
    '
    ' In addition, functions are stacked on top of those modes at various times:
    '  - at any time, making a shot advances combo multiplier for all other shots
    '  - during multiball, shots award jackpots and lead to super jackpot
    Public Sub RegisterHit(h)
        Dim line0,line1,line2
        Dim i,t
        Dim delay
        Dim combo:combo=1
        Dim combotext: combotext=""
        Dim qsound:qsound = "gotfx-qualify-sword-hit1"

    'debug.print "RegisterHit PM:" & PlayerMode & ", h:" & h & ", bBattleReady:" & bBattleReady

        If PlayerMode = 2 Then ' In WiC HurryUp Mode
            If h <> CurrentWiCShot Then Exit Sub
            ' This was the HurryUp Shot - Finish WiC HurryUp mode
            StopHurryUp
            'PlayerMode = 0
      SetPlayerMode 0
            AddScore HurryUpValue*CurrentWiCShotCombo
            WiCTotal = WiCTotal + HurryUpValue*CurrentWiCShotCombo
            WiCMask = WiCMask Or 2^h
            WiCs = WiCs + 1
            SetLightColor HouseShield(h),HouseColor(SelectedHouse),1
            If WiCs >= 4 And bBWMultiballActive = False Then
                StartWHCMB h
                WiCMask = 255
            End If
            DMDPlayHitScene "got-wiccomplete","gotfx-wiccomplete",0,"WINTER   IS   COMING",FormatScore(HurryUpValue*CurrentWiCShotCombo*PlayfieldMultiplierVal),"",CurrentWiCShotCombo,7
            Exit Sub
        End if

        If bHotkMBReady Or bITMBReady Then
            If h = Lannister Then
                If bHotkMBReady Then StartHotkMBIntro : Else StartIronThroneMB
            End If
            Exit Sub
        End If

        If bSwordLit And PlayerMode <> 2 And h = Stark Then DoAwardSword

        If bMadnessMB = 2 Then
            If h = Stark or h = Martell Then
                PlaySoundVol "gotfx-rightramphit",VolDef
            Elseif h = Greyjoy or h = Lannister Then
                PlaySoundVol "gotfx-ramphit",VolDef
            ' TODO: Do we want a default sound for main shots during Midnight Madness?
            End If
            DoMadnessMBHit
            Exit Sub
        End If

        If ComboLaneMap(h) Then combo = ComboMultiplier(ComboLaneMap(h))
        If h = Martell Then ROrbitsw31.UserValue = combo

        If bITMBActive Then
            If BWJackpotShots(h) > 0 Then
                BWJackpotShots(h) = BWJackpotShots(h) - 1
                ScoreITJP h,combo,false
            End If
        End If

        If bMultiBallMode And PlayerMode <> 2 And ((BWState MOD 2) <> 0 Or bBWMultiballActive = False) Then
            'Handle Jackpot hits for Blackwater, Wall, Winter Has Come, and Castle Multiball
            If BWJackpotShots(h) > 0 Then
                BWJackpotShots(h) = BWJackpotShots(h) - 1

                'Check to see if all jackpot shots have been made the required number of times
                t = True
                For i = 1 to 7
                    If BWJackpotShots(i) > 0 Then t=False:Exit For
                Next

                If bBWMultiballActive Then
                    AddScore(BWJackpotValue*combo*BWJackpotLevel)
                    BlackwaterScore = BlackwaterScore + (BWJackpotValue*combo*BWJackpotLevel*PlayfieldMultiplierVal)
                    i = ComboLaneMap(h)
                    If i = 0 then i = 1
                    PlaySoundVol "gotfx-bwexplosion",VolDef
                    DMDPlayHitScene "got-bwexplosion"&i-1,"gotfx-bwaward",BWExplosionTimes(i-1), _
                                    BWJackpotLevel&"X  BLACKWATER  JACKPOT",FormatScore(BWJackpotValue*combo*BWJackpotLevel*PlayfieldMultiplierVal),"",combo,3
                    debug.print "jackpot hit. BWJackpotvalue: "&BWJackpotValue&"  New BWScore: "&BlackwaterScore
                    DoBWJackpotSeq
                    i = RndNbr(6) : i = i - 2
                    If i < 1 Then i = 1
                    PlaySoundVol "say-jackpot"&i,VolCallout
                    SetModeLights   ' Update shield light colors
                    If t Then
                        BWState = BWState + 1
                        SetGameTimer tmrBlackwaterSJP,200
                        bBlackwaterSJPMode = True
                        SetBatteringRamLights
                        UpdateBWMBScene
                    End If
                ElseIf bCastleMBActive Then
                    ScoreCJP False,combo
                    If t Then SetJackpots : Else SetModeLights
                ElseIf bWallMBActive Then
                    WallMBShotCount = WallMBShotCount + 1
                    If WallMBShotCount = 3 Or h = Targaryen Then WallMBState = WallMBState + 1 : SetJackpots
                    ScoreWallJP h,combo
                ElseIf bWHCMBActive Then
                    If WHCMBState = 3 Then
                        ' Flipper was frozen. No score, but thaw flipper and restart Lieutenant mode
                        WHCFlipperFrozen = False
                        SetJackpots
                        WHCMBState = 2
                        DisableBallSaver
                        UpdateWHCScene
                        PlaySoundVol "gotfx-whchit",VolDef
                    ElseIf WHCMBState = 4 Then
                        ' Last chance - hit the lit shot to get back into Lieutenant mode
                        WHCMBState = 2
                        SetJackpots
                        WHCMBShots = 0
                        UpdateWHCScene
                        StopSound "gotfx-ticktock"
                        PlaySoundVol "gotfx-whchit",VolDef
                    Else
                        ScoreWHCJP h,combo
                    End If
                ElseIf bHotkMBActive Then
                    ScoreHotkJackpot combo,false
                End If
            End If ' BWJackpotShots > 0
        End If 'bMultiballMode

        if PlayerMode = 0 And bMultiBallMode = False Then
            Dim cbtimer: cbtimer=1000
            Dim fmt

            if QualifyCount(h) < 3 Then
                QualifyCount(h) = QualifyCount(h) + 1
                If h <> Baratheon and h <> Tyrell Then AddBonus 100000

                If h = Martell Then ROrbitsw31.UserValue = combo + 10

                line0 = "HOUSE   " & HouseToUCString(h)
                if QualifyCount(h) = 3 Then
                    bQualified(h) = True
                    BattleReady = True
                    ResetLights
                    line2 = "HOUSE IS LIT"
                Else
                    line2 = (3 - QualifyCount(h)) & " MORE TO LIGHT"
                End If
                line1 = FormatScore(QualifyValue*combo*PlayfieldMultiplierVal)

                AddScore(QualifyValue*combo)

                ' Play the animation and sound(s)
                delay = QAnimateTimes(h*3+QualifyCount(h))
                i = RndNbr(3)
                If h = Martell or ((h = Tyrell or h = Greyjoy) and QualifyCount(h) = 2) _
                    Or (h = Stark and QualifyCount(h) < 3) Or (h = Targaryen and i = QualifyCount(h)) Then
                        PlaySoundVol "say-"&HouseToString(h)&"qualify"&QualifyCount(h),VolCallout
                End If
                if h = Targaryen then fmt = 6 else fmt = 0
                DMDPlayHitScene "got-"&HouseToString(h)&"qualify"&QualifyCount(h),qsound,delay,line0,line1,line2,combo,fmt
                ' Play a tickle sound
                If h = Lannister Or h = Stark or h = Martell Then
                    PlaySoundVol "gotfx-tickle",VolDef
                Else PlaySoundVol "gotfx-tickle"&HouseToString(h),VolDef
                End If

                 ' Increase Qualify value for next shot. Lots of randomness seems to factor in here
                if QualifyValue = 100000 Then
                    QualifyValue = 430000
                Else
                    i = RndNbr(5)
                    if i = 1 Then
                        ' Use a wide range for random increase
                        QualifyValue = QualifyValue + 100000 + (RndNbr(50)*10000)
                    Else
                        ' Use a narrower range for increase
                        QualifyValue = QualifyValue + 275000 + (RndNbr(30)*5000)
                    End If
                End If
            ElseIf (WiCMask And 2^h) = 0 Then
                AddScore 50000*combo
                AddBonus 25000

                If WiCValue = 0 Then
                    If HouseSelected = Stark Then WiCValue = 14075000 Else WiCValue = 4075000
                Else
                    WiCValue = WiCValue + RndNbr(700)*1000 + 300000 ' TODO Better way of calculating next WiC Value
                End If
                WiCShots = WiCShots + 1
                If (SelectedHouse = Greyjoy And WiCShots = 4) Or (SelectedHouse <> Greyjoy And WiCShots = 3) Then
                    CurrentWiCShot = h
                    CurrentWiCShotCombo = combo
                    StartWICHurryUp WiCValue,h
                    WiCShots = 0
                    Exit Sub
                    ' LockBall will be handled by the sw48_hit sub, as ChooseBattle will never fire below
                Else
                    DMDDoWiCScene WiCValue,WiCShots
                    cbtimer = cbtimer + 1000
                End If
            Else
                ' Shot is qualified and WIC is done - play a different sound. Orbits and ramps each get their own.
                ' Targaryen is presumably a dragon roar
                If h = Stark or h = Martell Then
                    PlaySoundVol "gotfx-rightramphit",VolDef
                Elseif h = Greyjoy or h = Lannister Then
                    PlaySoundVol "gotfx-ramphit",VolDef
                ElseIf h = Targaryen Then
                    DoDragonRoar 1
                End If
            End If

            if bBattleReady and h = Lannister And bITMBActive = False Then    ' Kick off House Battle selection
                'PlayerMode = -2
        SetPlayerMode -2
                bReadyForFlips=False
                FreezeAllGameTimers
                If BallsInLock < 2 And bLockIsLit Then ' Multiball not about to start, lock the ball first
                    vpmtimer.addtimer 400, "LockBall '"     ' Slight delay to give ball time to settle
                    cbtimer = cbtimer + 800
                End If
                vpmtimer.addtimer cbtimer, "StartChooseBattle '"
            End If
        ElseIf PlayerMode = 1 Then
            If HouseBattle2 > 0 Then MyBattleState(HouseBattle2).RegisterHit(h)
            If HouseBattle1 > 0 Then MyBattleState(HouseBattle1).RegisterHit(h)
        End If

        If h = Targaryen And bMultiBallMode = False And bITMBActive = False Then AdvanceWallMultiball 1

        IncreaseComboMultiplier(h)
    End Sub

    Public Sub StartWIC
        Dim h
        If WiCValue = 0 Then
            If HouseSelected = Stark Then WiCValue = 14075000 Else WiCValue = 4075000
        End If
        For h = 1 to 7
            If QualifyCount(h) >= 3 And (WiCMask And 2^h) = 0 Then
                CurrentWiCShot = h
                CurrentWiCShotCombo = 1
                StartWICHurryUp WiCValue,h
                WiCShots = 0
                Exit Sub
            End If
        Next
    End Sub

    Public Sub GoldHit(n)
        Dim i,j
        AddScore 30
        If PlayerMode = 1 And (HouseBattle1 = Lannister Or HouseBattle2 = Lannister) Then
            If HouseBattle1 = Lannister Then
                MyBattleState(HouseBattle1).RegisterGoldHit n
            Else
                MyBattleState(HouseBattle2).RegisterGoldHit n
            End If
        Else
            ' Any mode except Lannister Battle mode
            If bGoldTargets(n) = False Then
                PlaySoundVol "gotfx-goldcoin",VolDef
                If HasAbility(Lannister) Then AddGold 100 Else AddGold 15
                bGoldTargets(n) = True
                SetLightColor GoldTargetLights(n),yellow,1
                j = True
                For i = 0 to 4
                    If bGoldTargets(i) = False Then j=False
                Next
                If j Then
                    ' Target bank completed. Light mystery, turn off target lights
                    ' Probably need to play a sound here
                    If HasAbility(Lannister) Then AddGold 500 Else AddGold 250
                    For i = 0 to 4: bGoldTargets(i) = False: Next
                    bMysteryLit = True : SetMystery
                    ' tell the gold target lights to turn off in 1 second. There's a timer on the first light
                    GoldTargetLights(0).TimerInterval = 1000: GoldTargetLights(0).TimerEnabled = True
                End If
            Else
                ' Already lit
                If HasAbility(Lannister) Then AddGold 15 Else AddGold 5
                PlaySoundVol "gotfx-litgold",VolDef/2
            End If
        End If
    End Sub

    Sub HouseCompleted(h)
        Dim i
        bQualified(h) = False
        If bCompleted(h) = False Then
            bCompleted(h) = True
            CompletedHouses = CompletedHouses + 1
            DoSwordsAreLit
            If (DMDStd(kDMDFet_HouseCompleteExtraBall)=8 And CompletedHouses=3) Or _
                DMDStd(kDMDFet_HouseCompleteExtraBall) = CompletedHouses Then DoEBisLit
            If CompletedHouses = 4 Then
                For i = 1 to 7
                    If bCompleted(i) Then HotkMask = HotkMask Or 2^i
                Next
            End If
            If CompletedHouses >=4 Then CheckForHotkOrITReady
            If h = Lannister Then
                If MyHouse = Lannister Then AddGold 400 Else AddGold 250
            End If
            If MyHouse = Greyjoy Then ActionAbility = h : ActionButtonUsed = False
        End If
    End Sub

    ' Called once per second during Winter-Has-Come
    Sub CheckWHCTimer
        Dim i,j,k,l
        If WHCMBState = 1 Then
            'Horde mode. Check to see if we should start ticking, and whether to add back a shot
            If WHCHordeTimestamp < GameTimeStamp+200 Then
                If WHCHordeTimestamp > GameTimeStamp Then
                    ' Play the clock tick on repeat. Don't restart it if already playing
                    PlaySound "gotfx-ticktock", -1, VolDef, 0, 0, 0, 1, 0
                Else
                    ' Time's up. Add back a shot
                    StopSound "gotfx-ticktock"
                    PlaySoundVol "gotfx-chime",VolDef/2
                    AddWHCShot
                    WHCHordeTimestamp = GameTimeStamp + 250
                End If
            Else
                StopSound "gotfx-ticktock"
            End If
        ElseIf WHCMBState = 2 Then
            'Lieutenant mode
            ' First see if any shots have timed out
            j = False : k = False : l = False
            For i = 1 to 7
                If BWJackpotShots(i) > 0 Then
                    If WHCShotTimestamp(i) < GameTimeStamp Then
                        ' Time has run out on this shot
                        BWJackpotShots(i) = 0 : k = True
                    Else
                        j = True
                    End If
                    If WHCShotTimestamp(i) < GameTimeStamp+100 Then l = True
                End If
            Next
            If l Then PlaySound "gotfx-ticktock", -1, VolDef, 0, 0, 0, 1, 0 Else StopSound "gotfx-ticktock"
            If k Then SetModeLights ' Yes, re-set the lights
            If j = False Then
                ' All shots have timed out. Move to State 3 - frozen flipper, and light left ramp only
                StopSound "gotfx-ticktock"
                For i = 1 to 7 : BWJackpotShots(i) = 0 : WHCShotTimestamp(i) = 0 : Next
                BWJackpotShots(Lannister) = 1 : WHCShotTimestamp(Lannister) = GameTimeStamp + 150
                WHCMBState = 3
                SetModeLights
                EnableBallSaver 16
                WHCFlipperFrozen = True : SolLFlipper 0
                PlaySoundVol "gotfx-whcfreeze",VolDef
            Else
                PlaySound "gotfx-ticktock", -1, VolDef, 0, 0, 0, 1, 0
            End If
        ElseIf WHCMBState = 3 Then
            ' Frozen flipper state. See if the Center ramp shot has timed out
            If BWJackpotShots(Lannister) > 0 And WHCShotTimestamp(Lannister) < GameTimeStamp Then
                ' Yup. Now we need to shoot a random shot within 10 seconds to get back in the game
                PlaySoundVol "gotfx-whcchime",VolDef/2
                PlaySound "gotfx-ticktock", -1, VolDef, 0, 0, 0, 1, 0
                BWJackpotShots(Lannister) = 0
                WHCFlipperFrozen = False
                WHCMBState = 4
                i = RndNbr(7)
                BWJackpotShots(i) = 1 : WHCShotTimestamp(i) = GameTimeStamp + 100
                UpdateWHCScene
                SetModeLights
            End If
        ElseIf WHCMBState = 4 Then
            For i = 1 to 7
                If BWJackpotShots(i) > 0 And WHCShotTimestamp(i) < GameTimeStamp Then
                    PlaySoundVol "gotfx-chime",VolDef/2
                    ResetForWHCMB
                    DMDCreateWHCMBScoreScene
                    Exit Sub
                End If
            Next
        End If
    End Sub

    Public Sub UpdateWHCScene
        Dim line,obj2
        If bUseFlexDMD = False Then Exit Sub
        If ScoreScene.Name <> "whcscore2" Then Exit Sub ' This should never happen, but just in case
        If WHCMBShots < 4 Then line = 5-WHCMBShots & " SHOTS NEEDED" Else line = "1 SHOT NEEDED"
        Set obj2 = ScoreScene.GetLabel("obj2")
        FlexDMD.LockRenderThread
        ScoreScene.GetLabel("obj").Text = line
        If WHCMBState = 3 Then
            With obj2
                .Visible = 1
                .Text = "FLIPPER FROZEN"&vbLf&"SHOOT LEFT RAMP TO THAW"
                .SetAlignedPosition 64,14,FlexDMD_Align_Center
            End With
        ElseIf WHCMBState = 4 Then
            with obj2
                .Visible = 1
                .Text = "TIME IS RUNNING OUT"&vbLf&"SHOOT LIT SHOT"
                .SetAlignedPosition 64,14,FlexDMD_Align_Center
            End With
        Else
            obj2.Visible = 0
        End If
        FlexDMD.UnlockRenderThread
    End Sub

    Public Sub UpdateBWMBScene
        If ScoreScene.Name <> "bwmb" Then Exit Sub
        FlexDMD.LockRenderThread
        If (BWState MOD 2) = 1 Then
            ScoreScene.GetLabel("obj").Text = "SHOOT GREEN JACKPOTS"
            ScoreScene.GetLabel("Score").SetAlignedPosition 127,0,FlexDMD_Align_TopRight
            ScoreScene.GetLabel("tmr1").Visible = 0
        Else
            ScoreScene.GetLabel("obj").Text = "SHOOT BATTERING RAM"
            ScoreScene.GetLabel("Score").SetAlignedPosition 102,0,FlexDMD_Align_TopRight
            ScoreScene.GetLabel("tmr1").Visible = 1
        End If
        If BWState > 1 Then ScoreScene.GetLabel("line1").Text = "BLACKWATER"&vbLf&"PHASE 2"
        FlexDMD.UnlockRenderThread
    End Sub


    '***************************
    ' Upper Playfield Processing
    '***************************

    ' UPF lights
    ' 1 - li186 - Castle
    ' 2 - li189 - Left target
    ' 3 - li180 - Left outlane
    ' 4 - li192 - Center target
    ' 5 - li183 - Right outlane
    ' 6 - li195 - Right target
    ' 7 - li198 - shield
    ' 11 - li201 - Castle MB
    ' 10 - li204 - Breach
    ' 9 - li207 - Charge
    ' 8 - li210 - Arrows
    ' 12 - li213 - 2X
    ' 13 - li216 - 3X

    Public Sub SetUPFLights
        Dim clr,i
        If UPFState = 2 Then 'Blackwater MB
            clr = yellowgreen
        ElseIf UPFState = 1 Then
            clr = HouseColor(HouseBattle1)
        Else clr = cyan
        End If
        For i = 1 to 13 : UPFLights(i).State = 0 : UPFLights(i).BlinkInterval = 100 : Next
        SetLightColor UPFLights(1),amber,1  ' Castle is lit amber unless it's not

        debug.print "Setting UPF lights to color "&clr&" with mask "&UPFShotMask
        ' top lights
        If UPFState <> 1 Or HouseBattle2 = 0 Then ' normal single flashing colour
            For i = 1 to 8
                If (UPFShotMask And 2^(i-1)) > 0 Then
                    If UPFState = 3 And (i = 3 or i = 5) Then ' CastleMB Super Jackpot shots
                        SetLightColor UPFLights(i),amber,2
                    Else
                        SetLightColor UPFLights(i),clr,2
                    End If
                End If
            Next
            For i = 8 to 10 : HouseShield(i).TimerEnabled = False : Next
        Else ' Stacked battle mode, use timers to cycle through colours
            For i = 8 to 10
                If (i=8 and (UPFShotMask And 2) > 0) Or (i=9 And (UPFShotMask And 8) > 0 ) Or (i=10 And (UPFShotMask And 32) > 0) Then
                    ModeLightState(i,0) = 3
                    ModeLightState(i,1) = 0
                    ModeLightState(i,2) = clr
                    ModeLightState(i,3) = HouseColor(HouseBattle2)
                    HouseShield(i).TimerInterval = 100
                    HouseShield(i).TimerEnabled = True
                Else
                    HouseShield(i).TimerEnabled = False
                End If
            Next
        End If

        ' Shield
        If UPFState > 0 And UPFState < 3 Then SetLightColor li198,clr,2
        ' Levels
        If UPFState = 0 Or UPFState = 3 Then
            For i = 1 to 4
                If UPFState = 3 Then
                    SetLightColor UPFLights(i+7),clr,1
                ElseIf UPFLevel = i Then
                    SetLightColor UPFLights(i+7),clr,2
                ElseIf UPFLevel > i Then
                    SetLightColor UPFLights(i+7),clr,1
                End If
            Next
        End If

        ' UPF playfield mult
        If UPFMultiplier = 2 Then
            SetLightColor li213,amber,2
        ElseIf UPFMultiplier = 3 Then
            SetLightColor li213,amber,1
            SetLightColor li216,amber,2
        End If
    End Sub

    ' Set up the UpperPF state.
    ' If reset=true, reset to initial state for the mode its entering
    Public Sub SetUPFState(reset)
        Dim i
        If (PlayerMode = 1 Or PlayerMode = -2.1) And reset Then UPFBattleShotMask = 42
        If bWallMBActive Or bWHCMBActive Or bHotkMBActive Or bHotkMBReady Or bITMBActive Or PlayerMode=2 Then
            UPFState = 4 : UPFShotMask = 0
            SetUPFFlashers 1,cyan
        ElseIf bCastleMBActive Then
            UPFState = 3 : UPFShotMask = 42
            SetUPFFlashers 1,cyan
        ElseIf bBWMultiballActive Then
            UPFState = 2 : UPFShotMask = 42
            SetUPFFlashers 2,yellowgreen
        ElseIf PlayerMode = 1 Or PlayerMode = -2.1 Then
            UPFState = 1
            UPFShotMask = UPFBattleShotMask
            SetUPFFlashers 1,amber
        Else ' PlayerMode 0
            UPFState = 0
            SetUPFFlashers 1,amber
            If reset then UPFShotMask = 42 Else UPFShotMask = UPFCastleShotMask
        End If
        SetUPFLights
    End Sub

    Public Sub ResetUPFMultiplier : UPFMultiplier = 1 : li213.State=0 : li216.State=0 : End Sub


    ' Process a hit on a UPF switch. Numbering:
    '  1 - Castle loop
    '  2 - left target
    '  3 - left outlane
    '  4 - center target
    '  5 - right outlane
    '  6 - right target
    '  7 - left inlane
    '  8 - right inlane
    Public Sub RegisterUPFHit(sw)
        Dim i
        If PlayerMode = 2 Then Exit Sub     ' No UPF progress during WiC HurryUp
        i = RndNbr(10)
        PlaySoundVol "gotfx-upfhit"&i,VolDef
        If bMadnessMB = 2 Then DoMadnessMBHit : Exit Sub
        i = RndNbr(3)
        debug.print "register UPF hit. Sw: "&sw&" State: "&UPFState&" UPFShotMask: "&UPFShotMask
        Select Case sw
            Case 1 ' Castle loop shot
                If UPFMultiplier < 3 Then ' Increase UPF multiplier
                    UPFMultiplier = UPFMultiplier + 1
                    SetGameTimer tmrUPFMultiplier,300
                    SetUPFLights
                End If
                If ((UPFShotMask And 1) > 0) Or (PlayerMode = 1 And (UPFBattleShotMask And 1) > 0) Then     ' Castle was lit
                    If UPFState = 0 Then
                        IncreaseUPFLevel : UPFShotMask = 42 : UPFCastleShotMask = 42 : PlaySoundVol "gotfx-wildfiremini2",VolDef
                    ElseIf PlayerMode = 1 Then ' Battlemode hit
                        UPFBattleShotMask = 42
                        Dim hc
                        ' Award a castle for each battle mode that's active
                        ' 1st castle is 25M, 2nd is 50M, etc plus 7.5M Bonus per
                        ' If 2 castles are scored at once, only second is displayed
                        If HouseBattle2 > 0 Then
                            CastlesCollected = CastlesCollected + 1
                            AddScore 25000000*CastlesCollected*UPFMultiplier
                            hc = HouseBattle2
                            myBattleState(HouseBattle1).RegisterCastleComplete
                        Else
                            hc = HouseBattle1
                        End If
                        CastlesCollected = CastlesCollected + 1
                        AddScore 25000000*CastlesCollected*UPFMultiplier
                        myBattleState(hc).RegisterCastleComplete
                        DMDPlayHitScene "got-castlecollected","gotfx-castlecollected",0,"CASTLE COLLECTED","HOUSE "&HouseToUCString(hc), _
                                FormatScore(25000000*CastlesCollected*UPFMultiplier*PlayfieldMultiplierVal),UPFMultiplier,10
                        PlaySoundVol "say-castlecollected",VolCallout
                    End If
                    If UPFState = 1 Then UPFShotMask = UPFBattleShotMask
                    SetUPFLights
                ElseIf PlayerMode = 1 Then
                    ' Castle wasn't lit but we're in battle mode
                    House(CurrentPlayer).BattleState(HouseBattle1).RegisterCastleHit
                    If HouseBattle2 > 0 Then House(CurrentPlayer).BattleState(HouseBattle2).RegisterCastleHit
                End If
            Case 2,4,6  ' Standup target
                If ((UPFShotMask And (2^(sw-1))) > 0) Or (PlayerMode = 1 And (UPFBattleShotMask  And (2^(sw-1))) > 0 ) Then ' Target was lit
                    Select Case UPFState
                        Case 0 ' Standard Play
                            DoDragonRoar i
                            PlaySoundVol "gotfx-elevatorupf",VolDef
                            AddBonus 50000
                            UPFShotMask = UPFShotMask Xor (2^(sw-1))
                            If (UPFShotMask And 254) = 0 Then
                                UPFShotMask = UPFShotMask Or 20 ' Light outlanes
                                AddScore 3000000*UPFMultiplier
                                PlaySoundVol "gotfx-upfdone",VolDef
                                DMDPlayHitScene "got-upfbackground","",0,"CASTLE  AWARD",3000000*UPFMultiplier*PlayfieldMultiplierVal,"",UPFMultiplier,5
                            Else
                                AddScore 250000*UPFMultiplier
                            End If
                            UPFCastleShotMask = UPFShotMask
                        Case 2,3 ' BW or Castle Multiball mode
                            If UPFSJP Then
                                Dim jpscore
                                UPFSJP = False
                                UPFShotMask = 42
                                If UPFMultiplier > 1 Then ScoreSJP UPFMultiplier Else ScoreSJP 0
                            Else
                                UPFShotMask = UPFShotMask Xor (2^(sw-1))
                                If UPFState = 3 Then
                                    If UPFShotMask = 0 Then UPFShotMask = 20
                                    ScoreCJP False,0
                                Else
                                    If UPFMultiplier = 1 Then
                                        jpscore = BWJackpotValue*BWJackpotLevel*PlayfieldMultiplierVal
                                    Else
                                        jpscore = BWJackpotValue*BWJackpotLevel*(PlayfieldMultiplierVal+UPFMultiplier)
                                    End If
                                    ' Do Jackpot scene
                                    AddScoreNoX jpscore
                                    BlackwaterScore = BlackwaterScore + jpscore
                                    PlaySoundVol "gotfx-bwexplosion",VolDef
                                    DMDPlayHitScene "got-bwexplosion5","gotfx-bwaward",BWExplosionTimes(5), _
                                            BWJackpotLevel&"X  BLACKWATER  JACKPOT",FormatScore(BWJackpotValue*UPFMultiplier*BWJackpotLevel*PlayfieldMultiplierVal),"",UPFMultiplier,3
                                    LightEffect 3
                                    If UPFShotMask = 0 Then UPFShotMask = 8 : UPFSJP = True
                                End If
                            End If
                    End Select
                    If PlayerMode = 1 Then ' Battle mode
                        House(CurrentPlayer).BattleState(HouseBattle1).AddTime 5
                        If HouseBattle2 > 0 Then House(CurrentPlayer).BattleState(HouseBattle2).AddTime 5
                        ' In Battle mode, hit sequence goes:
                        '  - hit any target to light two outer targets
                        '  - hit either outside target to light center target
                        '  - hit center target to light Castle
                        Select Case UPFBattleShotMask
                            Case 42: UPFBattleShotMask = 34 ' light outside targets
                            Case 34: UPFBattleShotMask = 8  ' light center target
                            Case 8: UPFBattleShotMask = 1   ' light Castle loop
                        End Select
                    End if
                    If UPFState = 1 Then UPFShotMask = UPFBattleShotMask
                    SetUPFLights
                Else
                    ' Not lit. Just give a few points
                    AddScore 560*UPFMultiplier
                End If
            Case 3,5    ' Outlanes
                If ((UPFShotMask And (2^(sw-1))) > 0) Or (PlayerMode = 1 And (UPFBattleShotMask  And (2^(sw-1))) > 0 ) Then  ' outlane was lit
                    Select Case UPFState
                        Case 0: IncreaseUPFLevel : UPFShotMask = (UPFShotMask And 1) Or 42 : UPFCastleShotMask = UPFShotMask ' Castle MB mode
                        Case 1 ' Battle mode - are these used at all?
                            AddScore 560*UPFMultiplier
                            ' TODO: If the Castle doesn't finish the mode, then outlane shots do advance the mode.
                        Case 2 ' BWMB mode - are these used?
                            AddScore 560*UPFMultiplier
                        Case 3 ' Castle MB mode
                            UPFShotMask = 42
                            ScoreCJP True,0
                    End Select
                    SetUPFLights
                End If
            Case 7: AddScore 560*UPFMultiplier : PlaySoundVol "gotfx-ramphit1",VolDef/4 ' Left inlane
            Case 8      ' Right Inlane
                PlaySoundVol "gotfx-ramphit1",VolDef/4
                AddScore 560*UPFMultiplier
                If bCastleShotAvailable And UPFState = 0 Then
                    UPFShotMask = UPFShotMask Or 1
                    bCastleShotAvailable = False
                    SetUPFLights
                End If
        End Select

    End Sub

    ' Advance the UPF level, either thru Castle hit or outlane
    ' Play the animation, add points. If it's the 4th hit, start Castle Multiball
    Public Sub IncreaseUPFLevel
        Dim delay,line1,line2,score,combo,format,i
        ' Score a level towards Castle MB
        score = (2500000 + 2500000*UPFLevel)*UPFMultiplier
        AddScore score
        i = RndNbr(3)
        line2 = FormatScore(score*PlayfieldMultiplierVal)
        format=3:combo=UPFMultiplier
        Select Case UPFLevel
            Case 1: line1 = "ARCHERS!"
            Case 2: line1 = "CHARGE!"
            Case 3: line1 = "BREACH!"
            Case 4: line1 = "CASTLE":line2="MULTIBALL":format=4:combo=0
        End Select

        DoDragonRoar i
        PlaySoundVol "gotfx-elevatorupf",VolDef

        If UPFLevel = 4 Then delay = 2.5 Else delay = 1.5
        DMDPlayHitScene "got-castlemblevel"&UPFLevel,"gotfx-castlemblevel"&UPFLevel,delay,line1,line2,"",combo,format
        UPFLevel = UPFLevel + 1
        If UPFLevel = 5 Then
            StartCastleMultiball
            UPFLevel=1
        ElseIf (DMDStd(kDMDFet_DisableRampDropTarget) = 0) Then
            'PlaySoundAt SoundFXDOF("DTReset", 119, DOFPulse, DOFcontactors), Target90
            DOF 119, DOFPulse
      RandomSoundDropTargetReset Target90
            DTRaise 4 : TargetState(4) = 0
        End If
    End Sub

    '*******************
    ' Jackpot shot processing for various MB modes
    '*******************

    ' Score a Castle Jackpot during Castle Multiball
    ' - add score
    ' - increase jackpot and super JP
    ' - update super JP in alternate score scene
    ' - do scene and sound
    Public Sub ScoreCJP(super,combo)
        Dim jpname,score,fmt,sound,dly,line2
        If combo = 0 then combo = UPFMultiplier
        If super Then
            score = CMBSJPValue*combo
            jpname = "got-cmbsuperjp"
            sound = "gotfx-dragonroar2"
            dly = 2 : fmt = 8
            line2 = "SUPER JACKPOT"
        Else
            ' setup for scene
            score = CMBJackpot*combo
            jpname = "got-cmbjp"
            sound = "gotfx-dragonroar1"
            line2 = "AWARD"
            dly = 0.9 : fmt = 9
            combo = 0
            ' setup for next JP
            CMBJackpot = CMBJackpot + CMBJPIncr
            CMBJPIncr = CMBJPIncr + 25000
            CMBSJPValue = 4 * CMBJackpot
            If CMBSJPValue > 28000000 Then CMBSJPValue = 5 * CMBJackpot
            If CMBSJPValue > 50000000 Then CMBSJPValue = 50000000
            If ScoreScene.Name = "cmb" Then ScoreScene.GetLabel("sjpval").Text = FormatScore(CMBSJPValue)
        End If
        AddScore score
        CastleMBScore = CastleMBScore + (score*PlayfieldMultiplierVal)
        DMDPlayHitScene jpname,sound,dly,"CASTLE   MULTIBALL"&vbLf&line2,FormatScore(score*PlayfieldMultiplierVal),"",combo,fmt
        FlashPlayfieldLights cyan,10,15
    End Sub

    ' Battle For The Wall Multiball Jackpots
    Public Sub ScoreWallJP(h,combo)
        Dim vid,snd,delay,score,i
        If combo = 0 then combo = 1
        If h = Targaryen Then
            vid = "got-wallmbsjp"
            snd = "gotfx-wallmbsjp"
            delay = 8
            score = (10000000 + WallJPValue)*combo
            i = RndNbr(3)
            if i = 1 Then PlaySoundVol "say-super-jackpot",VolCallout
         Else
            i = RndNbr(7)
            If i < 5 Then PlaySoundVol "say-walljp" & i,VolCallout
            'i =  WallMBShotCount
            'If (WallMBState MOD 4) = 2 Then i = i + 3
            vid = "got-wallmbhit" & i
            snd = "gotfx-increasewalljp"
            If i = 2 or i = 5 then delay = 1 else delay = 2
            score = WallJPValue*combo
        End If
        DMDPlayHitScene vid,snd,delay,"BATTLE  FOR  THE  WALL",FormatScore(score*PlayfieldMultiplierVal),"AWARD",combo,0
        DoWallMBJackpotSeq
        AddScore score
        WallMBScore = WallMBScore + (score*PlayfieldMultiplierVal)
        WallJPValue = WallJPValue + WallJPIncr
        WallJPIncr = WallJPIncr + 5000 : If HouseSelected = Baratheon Then WallJPIncr = WallJPIncr + 45000
        If ScoreScene.Name = "wallscore" Then
            If WallMBShotCount < 2 then i = "S" Else i = ""
            FlexDMD.LockRenderThread
            Select Case (WallMBState MOD 4)
                Case 0: ScoreScene.GetLabel("obj").Text = 3-WallMBShotCount & " MORE RAMP"&i
                Case 2: ScoreScene.GetLabel("obj").Text = 3-WallMBShotCount & " MORE ORBIT"&i
                Case 1,3: ScoreScene.GetLabel("obj").Text = "SHOOT SUPER JACKPOT"
            End Select
            FlexDMD.UnlockRenderThread
        End if
    End Sub

    ' Score a Winter-Has-Come Jackpot
    Public Sub ScoreWHCJP(h,combo)
        Dim score,i,t
        If WHCMBState = 1 Then
            ' Horde mode
            AddScore WHCJPValue*combo
            WHCMBScore = WHCMBScore + WHCJPValue*combo*PlayfieldMultiplierVal
            DMDPlayHitScene "whchit","gotfx-whchit",0,"WINTER   HAS   COME","HORDE AWARD",FormatScore(WHCJPValue*combo*PlayfieldMultiplierVal),combo,12
            WHCJPValue = WHCJPValue + WHCJPIncr
            WHCJPIncr = WHCJPIncr + 75000
            t = True
            For i = 1 to 7
                If BWJackpotShots(i) > 0 Then t=False:Exit For
            Next
            If t Then
                ' All Horde shots have been made. Switch to Lieutenant mode
                WHCMBState = 2
                SetJackpots
                WHCMBShots = 0
                DMDCreateWHCMBScore2Scene
            Else
                ' Add time back
                If WHCHordeTimestamp < GameTimeStamp + 250 Then WHCHordeTimestamp = GameTimeStamp + 250
                SetModeLights
            End If
        Else
            ' Lieutenant Mode
            WHCMBShots = WHCMBShots + 1
            StopSound "gotfx-ticktock"
            If WHCMBShots = 5 Then
                'Beat the Boss! Super Jackpot!
                score = (WiCTotal + 30000000)*combo
                AddScore score
                DMDPlayHitScene "got-whcsjp","gotfx-whcsjp",1.7,"WINTER   HAS   COME","SUPER JACKPOT",FormatScore(score*PlayfieldMultiplierVal),combo,12
                ' Set back to Horde mode
                WHCMBState = 1
                WHCHordeTimestamp = GameTimeStamp + 300
                SetJackpots
                DMDCreateWHCMBScoreScene
            Else
                If WHCMBShots < 3 Then AddWHCShot : Else SetModeLights
                UpdateWHCScene
                ' Score is calculated using time left on the shot - 900k per second left - plus a 7M base and 0-1M random component
                score = (7000000 + (WHCShotTimestamp(h)-GameTimeStamp)*90000 + RndNbr(40)*25000)*combo
                AddScore score
                WHCMBScore = WHCMBScore + score*PlayfieldMultiplierVal
                DMDPlayHitScene "whchit","gotfx-whchit",0,"WINTER   HAS   COME","LIEUTENANT AWARD",FormatScore(score*PlayfieldMultiplierVal),combo,12
            End If
        End If
    End Sub

    ' Hand Of The King Multiball Jackpot shots
    Public Sub ScoreHotkJackpot(combo,sjp)
        Dim score,t,i
        If sjp Then
            score = HurryUpValue
            AddScore score
            StopHurryUp
            FlashPlayfieldLights white,10,15
            SetSpotFlashers 3
            If (HotkMask And 2^Stark) > 0 Then
                HotkState = 2
                SetGameTimer tmrBlackwaterSJP,200
                SetGameTimer tmrUpdateBattleMode,5
            Else
                HotkState = 0
                HotkLevel = HotkLevel + 1
                If (HotkMask And 2^Tyrell) > 0 Then HotkSJPValue = 15000000 Else HotkSJPValue = 0
                If HotkLevel = 4 Then ResetForHotkMB
            End If
            SetJackpots
        Else
            score = HotkShotValue*combo
            AddScore score
            FlashPlayfieldLights white,10,15
            SetUPFFlashers 3,white
            SetSpotFlashers 5
            If HotkLevel = 3 Then
                HotkSuperSJPValue = HotkSuperSJPValue + score*PlayfieldMultiplierVal
            Else
                HotkSJPValue = HotkSJPValue + score*PlayfieldMultiplierVal
            End If
            t = True
            For i = 1 to 7
                If BWJackpotShots(i) > 0 Then t = False : Exit For
            Next
            If t Then
                ' Move to SJP state
                HotkState = 1
                SetBatteringRamLights
                If HotkLevel = 3 Then
                    StartHurryUp HotkSuperSJPValue,ScoreScene,20
                Else
                    StartHurryUp HotkSJPValue,ScoreScene,20
                End If
            End If
            SetModeLights
        End If
        HotkScore = HotkScore + score*PlayfieldMultiplierVal
        DMDPlayHitScene "got-hotk-sigil","gotfx-hotkhit",0,"",FormatScore(score*PlayfieldMultiplierVal),"",combo,11
        UpdateHotkScoreScene
    End Sub

    ' Handle all processing of Iron Throne shots
    Public Sub ScoreITJP(h,combo,sjp)
        Dim i,t,elapsed,score
        If ITState < 2 And sjp then Exit Sub
        If ITState = 0 Then
            ' Choose the next castle
            ITCastleMask = ITCastleMask And (2^h Xor 254)
            score = 10000000*combo
            DMDScrollITMap ITActiveCastle,h,score*PlayfieldMultiplierVal
            AddScore score
            ITScore = ITScore + score*PlayfieldMultiplierVal
            ITActiveCastle = h
            ITState = 1
            ITLastShotTimeStamp = GameTimeStamp
            SetJackpots
            DMDCreateITMBScoreScene ITState,ITActiveCastle,0
        ElseIf ITState = 1 Then
            ' Main mode, hitting lit shots
            elapsed = GameTimeStamp - ITLastShotTimeStamp
            score = 50000000 - (elapsed*250000)     ' 250,000 pts knocked off per 100ms elapsed
            if score < 12500000 Then score = 12500000
            AddScore score*combo
            ITScore = ITScore + score*combo*PlayfieldMultiplierVal
            ITLastShotTimeStamp = GameTimeStamp
            DMDPlayITHitScene ITActiveCastle,score*combo*PlayfieldMultiplierVal
            DoITMBJackpotSeq

            ' See if all shots have been hit
            t = True
            For i = 1 to 7
                If BWJackpotShots(i) > 0 Then t = False: Exit For
            Next
            If t Then
                If ITCastlesCollected < 7 Then
                    ITState = 2
                    SetModeLights
                    SetBatteringRamLights
                Else
                    ITActiveCastle = ITActiveCastle + 1
                    SetJackpots
                    DMDCreateITMBScoreScene ITState,ITActiveCastle,0
                End If
            Else
                SetModeLights
            End If
        ElseIf ITState = 2 And sjp Then
            ' Hit to the battering ram - Completed this set
            If (BallsOnPlayfield-RealBallsInLock) < 5 And ITCastlesCollected < 7 Then AddMultiballFast 1 : EnableBallSaver 17
            elapsed = GameTimeStamp - ITLastShotTimeStamp
            score = 50000000 - (elapsed*250000)     ' 250,000 pts knocked off per 100ms elapsed
            if score < 12500000 Then score = 12500000
            AddScore score*combo
            ITScore = ITScore + score*combo*PlayfieldMultiplierVal
            ITLastShotTimeStamp = GameTimeStamp
            If ITCastlesCollected = 6 Then ITActiveCastle = 0
            DMDPlayITHitScene ITActiveCastle,score*combo*PlayfieldMultiplierVal
            DoITMBJackpotSeq
            ITCastlesCollected = ITCastlesCollected + 1
            If ITCastlesCollected < 7 Then
                ITState = 0
                SetJackpots
                DMDCreateITMBScoreScene ITState,0,7-ITCastlesCollected
            Else
                ' Done all main castles, now we pick other ones at random for more shots
                bITMBDone = True
                If ITActiveCastle < 8 Then ITActiveCastle = 8 Else ITActiveCastle = ITActiveCastle + 1
                ITState = 1
                SetJackpots
                DMDCreateITMBScoreScene ITState,ITActiveCastle,0
            End If
            SetBatteringRamLights
        End If
    End Sub

    Sub UpdateHotkScoreScene
        Dim a
        If bUseFlexDMD = False Then Exit Sub
        If ScoreScene.Name <> "hotkscore" Then Exit Sub ' Scene hasn't been set up yet
        FlexDMD.LockRenderThread
        Select Case HotkState
            Case 0
                ScoreScene.GetGroup("state0").Visible = 1
                ScoreScene.GetGroup("state2").Visible = 0
                ScoreScene.GetLabel("obj").Text = "COMPLETE ALL SHOTS"
                If HotkLevel = 3 Then
                    ScoreScene.GetLabel("HurryUp").Text = FormatScore(HotkSuperSJPValue)
                    ScoreScene.GetLabel("sets").Text = "1 SET"
                Else
                    ScoreScene.GetLabel("HurryUp").Text = FormatScore(HotkSuperSJPValue)
                    ScoreScene.GetLabel("sets").Text = 4-HotkLevel & " SETS"
                End If
            Case 1: ScoreScene.GetLabel("obj").Text = "HIT BATTERING RAM"
            Case 2
                ScoreScene.GetGroup("state0").Visible = 0
                ScoreScene.GetGroup("state2").Visible = 1
        End Select
        FlexDMD.UnlockRenderThread
    End Sub

    Public Sub HotkHurryUpExpired
        If (HotkMask And 2^Stark) > 0 Then
            HotkState = 2
            UpdateHotkScoreScene
            SetGameTimer tmrBlackwaterSJP,200
            SetGameTimer tmrUpdateBattleMode,5
            SetJackpots
        Else
            HotkTimerExpired
        End If
    End Sub

    Public Sub HotkTimerExpired
        HotkState = 0
        UpdateHotkScoreScene
        HotkLevel = HotkLevel + 1
        If (HotkMask And 2^Tyrell) > 0 Then HotkSJPValue = 15000000 Else HotkSJPValue = 0
        If HotkLevel = 4 Then ResetForHotkMB
        SetJackpots
    End Sub

    ' Called when someone hits the Action button
    Sub CheckActionBtn
        Dim scene
        Select Case ActionAbility
            Case Stark
                If PlayerMode <> 1 or ActionButtonUsed Then Exit Sub
                'If bUseFlexDMD Then
                '    Set scene = NewSceneWithVideo("dwolf","got-direwolf")
                '    scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),HouseToUCString(HouseBattle1)&vbLf&"COMPLETED"&vbLf&"5,000,000")
                '    scene.GetLabel("txt").SetAlignedPosition 88,16,FlexDMD_Align_Center
                '    DMDEnqueueScene scene,1,1500,2000,3000,"gotfx-direwolf"
                'End If
        'debug.print "House completed with Direwolf: " & BattleState(HouseBattle1).GetCompletedWithDireWolf
                Me.HouseCompleted HouseBattle1
        BattleState(HouseBattle1).SetCompletedWithDireWolf = True
                BattleState(HouseBattle1).EndBattleMode
                SetLightColor HouseSigil(MyHouse),HouseColor(SelectedHouse),1
                AddScore 5000000
                ActionButtonUsed = True
        'play direwolf here to account for multihouse battles
        vpmTimer.AddTimer 400,"QueueDireWolf '"
        DOF 161, DOFPulse
            Case Baratheon
                If ActionButtonUsed Or bLoLLit Then Exit Sub
                DoLordOfLight False
                ActionButtonUsed = True
                SetGameTimer tmrLoLSave,120
        DOF 162, DOFPulse
            Case Lannister
                If ActionButtonUsed >= DMDStd(kDMDFet_LannisterButtonsPerGame) Or PlayfieldMultiplierVal = 5 Or PlayfieldMultiplierVal >= SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier)+3 Or PlayerMode < 0 Or PlayerMode > 1 Then Exit Sub
                If CurrentGold < ((PlayfieldMultiplierVal+1) * 600) Then Exit Sub 'Not enough gold.
                PlayfieldMultiplierVal = PlayfieldMultiplierVal + 1
                SetGameTimer tmrPFMultiplier,800-(PlayfieldMultiplierVal*100)
                SetPFMLights
                ActionButtonUsed = ActionButtonUsed + 1
                CurrentGold = CurrentGold - ((PlayfieldMultiplierVal+1) * 600)
                DoBuyMultiplierScene  ((PlayfieldMultiplierVal+1) * 600),12-ActionButtonUsed
        DOF 163, DOFPulse
            ' Greyjoy has no Action Button functionality
            Case Tyrell
                If ActionButtonUsed Then Exit Sub
                Dim Cashout,x,i : x=0
                For i=1 to 5 : x=x+ComboMultiplier(i)*PlayfieldMultiplierVal : Next
                Cashout = (100000+RndNbr(10000)*10)*x
                AddScore Cashout
                if bUseFlexDMD Then
                    Set scene = FlexDMD.NewGroup("bank")
                    scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("skinny10x12.fnt",vbwhite,vbwhite,0),"IRON BANK"&vbLf&FormatScore(Cashout))
                    scene.GetLabel("txt").SetAlignedPosition 64,16,FlexDMD_Align_Center
                    DMDEnqueueScene scene,1,2000,3000,2000,"gotfx-mmchime"
                End if
                ActionButtonUsed = True
                TimerFlags(tmrPFMultiplier) = 0
                PlayfieldMultiplierVal = 1
                SetPFMLights
                ResetComboMultipliers
        DOF 165, DOFPulse
            Case Martell
                If ActionButtonUsed Or bMultiBallMode = False Then Exit Sub
                DoAddABall
                ActionButtonUsed = True
        DOF 166, DOFPulse
            Case Targaryen
                If ActionButtonUsed Then Exit Sub
                TargaryenFreezeAllTimers
                ActionButtonUsed = True
                If bUseFlexDMD Then
                    Set scene = FlexDMD.NewGroup("freeze")
                    scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("skinny10x12.fnt",vbwhite,vbwhite,0),"TARGARYEN"&vbLf&"FREEZE TIMERS")
                    scene.GetLabel("txt").SetAlignedPosition 64,16,FlexDMD_Align_Center
                    DMDEnqueueScene scene,1,1000,1500,2000,""
                End If
        DOF 167, DOFPulse
        End Select
        SetLockbarLight
    End Sub

    Function HasAbility(h)
        If HouseSelected = h Or (HouseSelected = Greyjoy and bCompleted(h)) Then HasAbility = True Else HasAbility = False
    End Function

    Sub CheckBattleReady
        ' Check to see whether there are any non-lit houses. If not, always light BattleReady at the end of a battle
        ' Also light BattleReady if two balls are locked and at least one house is qualified
        Dim i,br,br1:br=True:br1=False

        If CompletedHouses = 7 Then Me.BattleReady = False : Exit Sub

        if bTargaryenInProgress = False Then
            For i = 1 to 7
                If bQualified(i) = False And bCompleted(i) = False  Then br = False
                If BallsInLock = 2 and bLockIsLit And bQualified(i) And bCompleted(i) = False Then br1=True
            Next
        End If
        If br or br1 And PlayerMode = 0 And bMultiBallMode = False Then Me.BattleReady = True Else Me.BattleReady = False
    End Sub

End Class

Dim ModeLightPattern
Dim AryaKills
Dim CompletionAnimationTimes
CompletionAnimationTimes = Array(0,2,2,1.5,4,2.3,6,2)

'Each number is a bit mask of which shields (shots) light up for the given mode
ModeLightPattern = Array(0,10,16,0,218,138,80,36)

AryaKills = Array("","","joffrey","cercai","walder frey","tywin","the red woman","beric dondarrion","Thoros of Myr", _
                "meryn trant","the hound", "the mountain","ilyn payne","")


' The cBattleState class tracks the state of a single mode (battle). There is one instance of this
' class for each mode a player has to complete

Class cBattleState
    Dim CompletedShots          ' Total shots accumulated for this battle
    Dim ShotMask           ' bitmask of shots that have been lit up
    Dim LannisterGreyjoyMask    ' bitmask of shots completed
    Dim GreyjoyMask             ' Mask of shots completed
    Dim GoldShotMask            ' Mask of which gold targets are lit during Lannister battle
    Dim CompletedDragons
    Dim MyHouse                 ' The house associated with this BattleState instance
    Dim State                   ' Current state of this house's battle
    Dim bComplete               ' Battle is complete
    Dim TotalScore              ' Total score accumulated battling this house
    Dim HouseValue              ' Most houses build value as the battle progresses. Stored here
    Dim HouseValueIncrement     ' Amount house value builds by, per shot, if machine-generated
    Dim MyHurryUps(3)           ' Holds the index values of any running HurryUps. Only Targaryen has more than one concurrently
    Dim OtherHouseBattle        ' Store the other house battle, if two are stacked. Used for doing Battle Total at the end
    Dim DrogonHits              ' Track health of Drogon, for Targaryen battle
    Dim TGShotCount
  Dim HouseCompleteWithDireWolf ' Track if this house was completed with Stark/DireWolf ability


    Private Sub Class_Initialize(  )
        CompletedShots = 0
        LannisterGreyjoyMask = 0
        GreyjoyMask = 218
        CompletedDragons = 0
        ShotMask = 0
        GoldShotMask = 31
        State = 1
        bComplete = False
        TotalScore = 0
        HouseValueIncrement = 0
        DrogonHits = 0
    HouseCompleteWithDireWolf = False
    End Sub

  Public Property Let SetCompletedWithDireWolf(value)
    HouseCompleteWithDireWolf = value
  End Property
  Public Property Get GetCompletedWithDireWolf
    GetCompletedWithDireWolf = HouseCompleteWithDireWolf
  End Property

    Public Property Let SetHouse(h)
        MyHouse = h
        Select Case h
            Case Stark: GreyjoyMask = 90
            Case Baratheon: GreyjoyMask = 152
            Case Lannister: HouseValue = 1000000 : GreyjoyMask = 218
            Case Greyjoy: HouseValue = 600000 + RndNbr(4)*25000 : GreyjoyMask = 218
            Case Tyrell: HouseValue = 9000000 + RndNbr(7)*125000 : GreyjoyMask = 138
            Case Martell: GreyjoyMask = 10
            Case Targaryen: HouseValue = 8000000 : GreyjoyMask = 0
        End Select
    End Property
  Public Property Get GetHouse : GetHouse = MyHouse : End Property
    Public Property Let OtherHouse(h) : OtherHouseBattle = h : End Property

    Public Sub SetBattleLights
        Dim mask,i

        ' Add Greyjoy-specific shots
        If SelectedHouse = Greyjoy And GreyjoyMask > 0 And MyHouse <> Greyjoy Then
            For i = 1 to 7
                If (GreyjoyMask And 2^i) > 0 Then
                    ModeLightState(i,0) = ModeLightState(i,0) + 1
                    ModeLightState(i,(ModeLightState(i,0))) = purple
                End If
            Next
        End If
        ' Load the starting state mask
        mask = ModeLightPattern(MyHouse)
        ' Adjust based on house and state
        Select Case MyHouse
            Case Stark
                If State = 2 Then mask = mask or 80     ' Light the orbits for State 2
            Case Baratheon
                If State = 2 Then mask = mask or 128    ' Light dragon shot
                If State = 3 Then mask = mask or 4      ' Light LoL target bank
            Case Lannister,Targaryen
                mask = ShotMask
            Case Greyjoy
                mask = GreyjoyMask
            Case Tyrell
                Select Case State
                    Case 2,4,6: mask = 32
                    Case 3: mask = 10
                    Case 5: mask = 2
                End Select
            Case Martell
                If State = 2 Then mask = 10
        End Select

        For i = 1 to 7
            If (mask And 2^i) > 0 Then
                ModeLightState(i,0) = ModeLightState(i,0) + 1
                ModeLightState(i,(ModeLightState(i,0))) = HouseColor(MyHouse)
            End If
        Next

        If MyHouse = Lannister Then
            For i = 0 to 4
                If (GoldShotMask And 2^i) > 0 Then
                    GoldTargetLights(i).State = 2
                    GoldTargetLights(i).BlinkInterval = 30
                Else
                    GoldTargetLights(i).State = 0
                End If
            Next
        ElseIf MyHouse = HouseBattle1 And HouseBattle2 <> Lannister Then SetGoldTargetLights
        End If
    End Sub

    ' Called to initialize battle mode for this house.
    Public Sub StartBattleMode
        Dim tmr: tmr=400    ' 10ths of a second
        Dim i,j,snd : snd=""
        OtherHouseBattle = 0
        Select Case MyHouse
            Case Stark
                State = 1
                GreyjoyMask = 90
                HouseValue = 500000
                ' TODO Stark shot increase value may depend on how quickly you make the shots
                If HouseValueIncrement = 0 Then HouseValueIncrement = 3000000 + RndNbr(15) * 125000
                If CompletedShots > 0 Then CompletedShots = 2
            Case Baratheon
                State = 1
                GreyjoyMask = 152
                OpenTopGates
                HouseValue = 1250000
                HouseValueIncrement = 900000
                ResetDropTargets
            Case Lannister: State=1
            Case Greyjoy: OpenTopGates : tmr = 150 : snd="say-greyjoybattlestart"
            Case Martell
                HouseValue = 0: tmr = 300
                If State=1 Then snd="say-martellbattlestart"
                State = 1
                GreyjoyMask = 10
                CompletedShots = 0
                OpenTopGates
            Case Tyrell
                If (State MOD 2) = 0 Then
                    ' Reset Wildfire targets
                    bWildfireTargets(0) = False: bWildfireTargets(1) = False
                    li80.BlinkInterval = 50
                    li83.BlinkInterval = 50
                    li80.State = 2
                    li83.State = 2
                Else
                    SetWildfireLights
                End If
            Case Targaryen
                tmr = 0
                bTargaryenInProgress = True
                If State = 1 Then snd="say-targaryenbattlestart"
                If State = 1 and SelectedHouse <> Greyjoy then State = 2
                Select Case State
                    Case 1: ShotMask = 36
                    Case 2: ShotMask = 10
                    Case 10
                        ResetDropTargets
                        ' Generate 3 randomly lit shots
                        If TGShotCount = 0 Then GenerateRandomShots
                End Select
        End Select
        If snd <> "" Then
            If MyHouse = HouseBattle1 Then
                PlaySoundVol snd,VolCallout*1.5
            Else
                vpmTimer.AddTimer 3000,"PlaySoundVol """&snd&""","&VolCallout*1.5&" '"
            End If
        End If
        If tmr > 0 Then
            If MyHouse = HouseBattle2 Then SetGameTimer tmrBattleMode2,tmr Else SetGameTimer tmrBattleMode1,tmr
        End If

    End Sub

    ' Update the state machine based on the ball hitting a target
    Public Sub RegisterHit(h)
        Dim hit,done,hitscene,hitsound,ScoredValue,i,j
        ScoredValue = 0
        ThawAllGameTimers
        if bComplete Then Exit Sub
        hitscene="":hitsound=""
        Select Case MyHouse
            ' Stark Battle mode. State 1 is shooting the left ramp. Once 3 shots have been made, switch to State 2
            ' In State 2, left ramp shots can continue to be made, and an orbit shot finishes the mode
            Case Stark
                If (SelectedHouse <> Greyjoy And (h = Lannister or h = Stark)) Or (SelectedHouse=Greyjoy And (h=Lannister or h=Stark or (GreyjoyMask And 2^h) > 0)) Then
                    ' Process ramp shot
                    If SelectedHouse=Greyjoy And (GreyjoyMask And 2^h) > 0 Then GreyjoyMask = GreyjoyMask And (2^h Xor 255) : SetModeLights
                    HouseValue = HouseValue + HouseValueIncrement
                    If HouseValue > 75000000 Then HouseValue = 75000000
                    HouseValueIncrement = HouseValueIncrement + 750000
                    PlaySoundVol "gotfx-ramphit",VolDef/4
                    If h = Lannister or h = Stark Then
                        CompletedShots = CompletedShots + 1 : If CompletedShots > 13 Then CompletedShots = 13
                        If (CompletedShots = 3 And SelectedHouse <> Greyjoy) Or (SelectedHouse=Greyjoy and GreyjoyMask = 0 And CompletedShots >=3) Then
                            State = 2
                            SetModeLights
                        End If
                        debug.print "Stark hits: "&CompletedShots
                        If CompletedShots >= 3 Then
                            'TODO: On the real table, if you restart the mode, it doesn't reuse the same victims. It also randomizes the order
                            ' Show Arya's kill list scene.
                            ' Photos alternate between right and left side of scene so adjust text alignment
                            Dim just1, just2
                            just1 = FlexDMD_Align_TopRight:just2 = FlexDMD_Align_BottomLeft
                            Select Case CompletedShots
                                Case 5,6,8,10,12,13,14: just1=FlexDMD_Align_TopLeft:just2 = FlexDMD_Align_BottomRight
                            End Select
                            ' Render battle hit scene. 'House,Scene #, Score, Text1, Text2, Score+Text1 text justification, text2 justification,sound
                            DMDStarkBattleScene Stark,CompletedShots-2,HouseValue,"STARK VALUE GROWS",AryaKills(CompletedShots),just1,just2,"say-aryakill"&CompletedShots-1
                        End If
                    End If
                    UpdateBattleScene
                ElseIf State = 2 And (h = Greyjoy or h = Martell) Then
                    DoCompleteMode h
                End if

            ' Baratheon Battle mode. State 1 only involves the spinner. Once enough value is built, switch to State 2
            ' State 2: Shot to the Dragon, followed by a shot to the left bank.
            Case Baratheon
                If SelectedHouse = Greyjoy And (GreyjoyMask And 2^h) > 0 Then
                    GreyjoyMask = GreyjoyMask And (2^h Xor 255)
                    hitscene = "hit2"
                    HouseValue = HouseValue + 750000
                    If ShotMask = 0 And GreyjoyMask = 0 Then
                        State = 3
                        ResetDropTargets
                    End If
                    SetModeLights
                End If
                If State = 2 Then
                    If (ShotMask And 2^h) > 0 Then
                        ShotMask = ShotMask And (2^h Xor 255)
                        hitscene = "hit2"
                        HouseValue = HouseValue + 750000
                        If ShotMask = 0 And (SelectedHouse <> Greyjoy Or GreyjoyMask = 0) Then
                            State = 3
                            ResetDropTargets
                            SetModeLights
                        End If
                    End If
                End If

            ' Lannister battle mode. Only 1 state. Shoot gold targets to light ramps on either side, then shoot
            ' ramps. Accumulate 5 shots total. Progress is saved. Each gold target can only be hit once, but there is
            ' some overlap in shots that they light, so some shots can be made twice to avoid harder shots. Greyjoy must
            ' shoot all 5 unique shots
            Case Lannister
                If SelectedHouse = Greyjoy And (GreyjoyMask And 2^h) > 0 Then
                    GreyjoyMask = GreyjoyMask And (2^h Xor 255)
                    hitscene = "hit2"
                    SetModeLights
                    HouseValue = HouseValue + 400000 + RndNbr(12)*25000
                End If
                If (ShotMask And 2^h) > 0 Then
                    ShotMask = ShotMask And (2^h Xor 255)
                    LannisterGreyjoyMask = LannisterGreyjoyMask Or 2^h
                    CompletedShots = CompletedShots + 1
                    If CompletedShots > 5 Then CompletedShots = 5
                    If CompletedShots = 5 And ((SelectedHouse = GreyJoy And GreyjoyMask = 0) or (SelectedHouse <> Greyjoy)) Then
                        DoCompleteMode h
                        hitscene = ""
                    Else
            'todo: verify lannister/greyjoy rules
            ' need to ensure the lannister hit scene is 1..5
            dim hitnum : hitnum = CompletedShots
            if CompletedShots <=0 then hitnum = 1 else if CompletedShots > 5 then hitnum=5
                        hitscene = "hit"&(hitnum) 'changing to ensure scene is from 1-5
            'hitscene = "hit"&(CompletedShots+1) 'hit1 is for gold target hit
            'if CompletedShots <> hitnum then debug.print "******** c-h " & CompletedShots & " " & hitscene
                        SetModeLights
                        ScoredValue = HouseValue
                        ' TODO: Figure out if there's a better pattern to capture Lannister mode value increases
                        HouseValue = HouseValue + (CompletedShots * 200000)+RndNbr(12)*25000
                        FlashPlayfieldLights red,10,15
                    End If
                End If


            ' Greyjoy battle mode. Only 1 state: Shoot the 5 main shots. 15 second timer, but the timer resets after
            ' each shot
            Case Greyjoy
                If (GreyjoyMask And 2^h) > 0 Then
                    ' Completed shot
                    CompletedShots = CompletedShots+1
                    GreyjoyMask = GreyjoyMask And (2^h Xor 255)
                    If GreyjoyMask = 0 Then 'Completed req'd shots!
                        DoCompleteMode h
                    Else
                        hitscene = "hit"&CompletedShots
                        hitsound = "hit1"
                        ScoredValue = HouseValue
                        Select Case CompletedShots
                            Case 1,3: HouseValue = HouseValue + 2250000 + RndNbr(8)*25000
                            Case 2: HouseValue = HouseValue + 1500000 + RndNbr(12)*15000
                            Case 4: HouseValue = HouseValue + 5000000 + RndNbr(12)*250000
                        End Select
                        SetModeLights
                        ' Reset mode timer
                        If MyHouse = HouseBattle2 Then SetGameTimer tmrBattleMode2,150 Else SetGameTimer tmrBattleMode1,150
                        FlashPlayfieldLights purple,10,15
                    End If
                End If

            ' Tyrell battle mode. 6 States. States 1/3/5 involve choice of 3/2/1 main shot(s), States 2/4/6 require a right-bank target shot
            Case Tyrell

                ' Are we playing as Greyjoy? If so, add value when Greyjoy shots are lit.
                ' All Greyjoy shots must be completed before mode can be completed
                If SelectedHouse = Greyjoy And (GreyjoyMask And 2^h) > 0 Then
                    GreyjoyMask = GreyjoyMask And (2^h Xor 255)
                    hitscene = "hit2"
                    SetModeLights
                    HouseValue = HouseValue + 900000 + (325000*(State-1))
                    ScoredValue = HouseValue
                    If State = 7 And GreyjoyMask = 0 Then DoCompleteMode 0
                End If
                hit=False
                Select Case State
                    Case 1
                        If h = Targaryen or h = Stark or h = Lannister Then hit = true
                    Case 2,4,6
                        If h = Tyrell Then
                            hit = true
                            ' Target hit. Add 5 seconds to the mode timer
                            If MyHouse = HouseBattle2 Then
                                TimerTimestamp(tmrBattleMode2) = TimerTimestamp(tmrBattleMode2)+50
                            Else
                                TimerTimestamp(tmrBattleMode1) = TimerTimestamp(tmrBattleMode1)+50
                            End If
                        End if
                    Case 3
                        If h = Stark or h = Lannister Then hit = true
                    Case 5
                        If h = Stark then hit = true
                End Select
                If hit Then
                    State = State + 1
                    If (State MOD 2) = 0 Then
                        ' Reset Wildfire targets
                        bWildfireTargets(0) = False: bWildfireTargets(1) = False
                        li80.BlinkInterval = 50
                        li83.BlinkInterval = 50
                        li80.State = 2
                        li83.State = 2
                    Else
                        SetWildfireLights
                    End If
                    If State = 7 And (SelectedHouse <> Greyjoy Or GreyjoyMask = 0) Then
                        DoCompleteMode 0
                        hitscene = ""
                    Else
                        hitscene = "hit"& (State MOD 2) + 1
                        ScoredValue = HouseValue
                        HouseValue = HouseValue + 900000 + (325000*(State-1))
                        If HouseValue > 18000000 Then HouseValue = 18000000
                        SetModeLights
                        FlashPlayfieldLights green,10,15
                    End If
                End If

            ' Martell battle mode: State 1 requires 3 orbit shots within a 10 second period. After this, mode is
            ' complete, but State 2 is started, which is a HurryUp at the ramps. Shoot the ramps for bonus. If
            ' you miss the HurryUp, mode still completes
            Case Martell
                Dim huvalue

                ' Are we playing as Greyjoy? If so, Greyjoy shots add value but aren't mandatory
                If SelectedHouse = Greyjoy And (GreyjoyMask And 2^h) > 0 Then
                    GreyjoyMask = GreyjoyMask And (2^h Xor 255)
                    hitscene = "hit1"
                    SetModeLights
                    If HouseValue = 0 Then HouseValue = 19250000 Else HouseValue = HouseValue + 500000 + RndNbr(25)*125000
                End If

                If State = 1 And (h = Greyjoy or h = Martell) Then
                    CompletedShots = CompletedShots + 1
                    debug.print "Martell Compl shots: " & CompletedShots & " LastSwHit: " &LastSwitchHit
                    hitscene = "hit"&CompletedShots
                    If CompletedShots = 3 Then 'State 1 complete
                        HouseValue = HouseValue + 9000000 + RndNbr(7)*125000
                        ScoredValue = HouseValue
                        TimerFlags(tmrMartellBattle) = 0
                        State = 2
                        UpdateBattleScene
                        StartHurryUp ScoredValue,MyBattleScene,0
                        HouseValue = 0
                        SetModeLights
                    Else
                        If HouseValue = 0 Then HouseValue = 19250000 Else HouseValue = HouseValue + 500000 + RndNbr(25)*125000:
                        ' Start or reset a 10 second timer
                        SetGameTimer tmrMartellBattle,100
                        ' Reset mode timer if less than 10 seconds left
                        If (MyHouse = HouseBattle1 And TimerTimestamp(tmrBattleMode1)-GameTimeStamp < 100) Then
                            SetGameTimer tmrBattleMode1,300
                        ElseIf (MyHouse = HouseBattle2 And TimerTimestamp(tmrBattleMode2)-GameTimeStamp < 100) Then
                            SetGameTimer tmrBattleMode2,300
                        End If
                    End If
                End if
                If State = 2 And (h = Stark or h = Lannister) Then
                    huvalue = HurryUpValue
                    StopHurryUp
                    If huvalue > 0 Then
                        'Hurry-up hit in time
                        HouseValue = huvalue
                        DoCompleteMode h
                    End if

                End If

            ' Targaryen: 3 LEVELS, each with 4 states except last which has 3 that repeat 4 times - 11 states total
            ' Only Greyjoy sees States 1,5, and 9
            ' All shots involve HurryUps. Mode ends if HurryUp runs down
                ' Level 1 Start: light 2 ramps for HurryUp. 8M
                ' State 1: HurryUp at the target banks for Greyjoy
                ' State 2: Shoot 1 of the lit ramps to advance to State 2
                ' State 3: Light 2 loops. Shoot one to advance to dragon. 10M
                ' State 4: Start hurry-up on Dragon. 12M
                ' Level 2: Repeat Level 1, but require all 4 shots in State 2(6) & 3(7)
                ' Level 3: Light 3 random shots as hurry-ups (mode ends if hurry-ups timeout?)
                ' State 9: HurryUp at the target banks for Greyjoy
                ' State 10: Shoot all 3 hurry-ups. Shooting Dragon spots a hurry-up.
                ' State 11: Shoot Dragon hurry-up
                ' States 10 & 11 repeat until Drogon health is used up - 15 shots total, then clear the current Wave (State 10 & 11) to end the mode
                ' Timer on Level 3: If you take too long, you are attacked with “DRAGON FIRE”, and wave restarts with new randomly chosen shots (State 1, but same Wave)
            Case Targaryen
                hit = False:done=False
                Select Case State
                    Case 1,5
                        If h = Tyrell or h = Baratheon Then hit=true:done=True:ShotMask = 10
                    Case 2
                        If h = Stark or h = Lannister or h = Targaryen Then hit=true:done=True:ShotMask = 80
                    Case 3
                        If h = Greyjoy or h = Martell or h = Targaryen Then hit=true:done=True:ShotMask = 128
                    Case 4,8
                        If h = Targaryen Then
                            hit=true:done=true
                            bTargaryenInProgress = False
                            ShotMask = 10
                        End If
                    Case 6
                        If (ShotMask And 2^h) > 0 Then
                            hit=true
                            ShotMask = ShotMask And (2^h Xor 255)
                            If ShotMask = 0 Then done=true:ShotMask=80
                        End If
                        If h = Targaryen Then
                            hit = true
                            If ShotMask = 8 or Shotmask = 2 Then done = true:ShotMask=80 Else ShotMask = 8
                        End If
                    Case 7
                        If (ShotMask And 2^h) > 0 Then
                            hit=true
                            ShotMask = ShotMask And (2^h Xor 255)
                            If ShotMask = 0 Then done=true:ShotMask=128
                        ElseIf h = Targaryen Then
                            hit = true
                            If ShotMask = 16 or Shotmask = 64 Then done = true:ShotMask=128 Else ShotMask = 16
                        End If
                    Case 9
                        If h = Tyrell or h = Baratheon Then hit=true:done=True:TGShotCount=0 : GenerateRandomShots
                    Case 10
                        If (ShotMask And 2^h) > 0 Then
                            hit=true
                            ShotMask = ShotMask And (2^h Xor 255)
                            TGShotCount = TGShotCount + 1
                            DrogonHits = DrogonHits + 1
                            If ShotMask = 0 Then done=true:ShotMask=128
                        ElseIf h = Targaryen Then
                            hit = true
                            For i = 1 to 6
                                If (ShotMask And 2^i) > 0 Then
                                    ShotMask = ShotMask And (2^i Xor 255)
                                    Exit For
                                End If
                            Next
                            TGShotCount = TGShotCount + 1
                            DrogonHits = DrogonHits + 1
                            If ShotMask = 0 Then done=true:ShotMask=128
                        End If
                    Case 11
                        If h = Targaryen Then
                            hit=true:done=true
                            DrogonHits = DrogonHits + 1
                        End If
                End Select
                If hit Then
                    ScoredValue = TGHurryUpValue
                    StopTGHurryUp
                    If State < 5 Then
                        hitscene = "hit1"
                    ElseIf State < 9 Then
                        hitscene = "hit2"
                        hitsound = "hit1"
                    ElseIf State = 9 or State = 10 Then
                        hitscene = "hit" & 3 + TGShotCount MOD 3
                        hitsound = "hit1"
                    Else
                        hitscene = "hit6"
                        hitsound = "hit3"
                    End if
                    If (State=6 Or State=7 Or State=10) And done = False then TGStartHurryUp
                    If done = False then SetModeLights
                End If
                If done Then
                    State = State + 1
                    Select Case State
                        Case 5,9
                            If SelectedHouse = Greyjoy Then ShotMask = 36 Else State = State + 1
                            EndBattleMode
                        Case 12
                            If DrogonHits < 15 Then
                                TGShotCount = 0
                                State = 10
                                ShotMask = 0
                                For i = 1 to 3
                                    Do
                                        j = RndNbr(6)
                                    Loop While (ShotMask And 2^j) > 0
                                    ShotMask = ShotMask Or 2^j
                                Next
                                TGStartHurryUp
                            Else
                                HouseValue = 0 ' Score is taken care of by hitscene below
                                DoCompleteMode 0
                            End if
                        Case Else: TGStartHurryUp
                    End Select
                    SetModeLights
                End If

        End Select

        If hitscene <> "" Then
            Dim line2,line3,name,sound,combo
            AddBonus 100000
            line3 = "JACKPOT BUILDS"
            combo = 0
            'Play a scene and add score
            If ScoredValue <> 0 Then
                line3 = "AWARDED"
                combo = ComboMultiplier(ComboLaneMap(h))
                line2 = FormatScore(ScoredValue*combo*PlayfieldMultiplierVal)
                AddScore ScoredValue*combo
                TotalScore = TotalScore + (ScoredValue*combo*PlayfieldMultiplierVal)
            Else
                line2 = FormatScore(HouseValue)
            End If
            name = "got-"&HouseToString(MyHouse)&"battle"&hitscene
            if hitsound <> "" Then
                sound = "gotfx-"&HouseToString(MyHouse)&"battle"&hitsound
            Else
                sound = "gotfx-"&HouseToString(MyHouse)&"battle"&hitscene
            End if
            If MyHouse = Targaryen Then
                DMDPlayHitScene name,sound,1.5,"HOUSE TARGARYEN",line2,"",combo,3
            Else
                DMDPlayHitScene name,sound,1.5,BattleObjectivesShort(MyHouse),line2,line3,combo,1
            End If
            UpdateBattleScene
        End If

    End Sub

    ' Finish the mode. 'Shot' is the shot # that completed the mode, in case a combo multiplier is involved
    Public Sub DoCompleteMode(shot)
    'debug.print "DoCompleteMode " & GameTimeStamp
        Dim comboval,line2,name,sound,delay

        If MyHouse = Targaryen Then bTargaryenInProgress = False
        bComplete = True
        House(CurrentPlayer).HouseCompleted MyHouse

        EndBattleMode

        SetLightColor HouseSigil(MyHouse),HouseColor(SelectedHouse),1

        If shot > 0 And MyHouse <> Baratheon And shot <> Tyrell Then
            comboval = ComboMultiplier(ComboLaneMap(shot)) * PlayfieldMultiplierVal
        Else
            comboval = PlayfieldMultiplierVal
        End If
        TotalScore = TotalScore + HouseValue * comboval

        AddScore HouseValue * comboval
        line2 = FormatScore(HouseValue * comboval)
        If comboval = 1 Then comboval = 0  ' Don't bother printing Combo value for final shot if it's just 1x
        name = "got-"&HouseToString(MyHouse)&"battlecomplete"
        sound = "gotfx-"&HouseToString(MyHouse)&"battlecomplete"
        ' Baratheon gets two sounds
        If MyHouse = Baratheon Then
            PlaySoundVol sound,VolDef
            sound = "say-baratheonbattlecomplete"
        End If
        delay = CompletionAnimationTimes(MyHouse)
        DMDPlayHitScene name,sound,delay,BattleObjectivesShort(MyHouse),line2,"COMPLETE",comboval,1
        If HouseBattle1 = 0 And HouseBattle2 = 0 And bUseFlexDMD Then
      if HouseCompleteWithDireWolf = True then debug.print "**** House completed with DireWolf (DoCompleteMode)"
            tmrBattleCompleteScene.Interval=3000
            tmrBattleCompleteScene.Enabled = 1
            tmrBattleCompleteScene.UserValue = MyHouse
        End If
    End Sub

    ' Return to normal play after battle mode
    Public Sub EndBattleMode
    'debug.print "EndBattleMode " & GameTimeStamp
        Dim br,i
        ' Disable mode timer and HouseBattle
        If MyHouse = Targaryen Then
            StopDragonWings
            If bTGHurryUpActive Then StopTGHurryUp
        End If
        If MyHouse = HouseBattle1 Then
            TimerFlags(tmrBattleMode1) = 0
            HouseBattle1 = 0
            If HouseBattle2 <> 0 Then ' Move Stacked battle and its timer to the primary position
                House(CurrentPlayer).BattleState(HouseBattle2).OtherHouse = MyHouse
                HouseBattle1=HouseBattle2 : HouseBattle2=0
                timerTimestamp(tmrBattleMode1) = timerTimestamp(tmrBattleMode2)
                TimerFlags(tmrBattleMode1) = TimerFlags(tmrBattleMode2)
                TimerFlags(tmrBattleMode2) = TimerFlags(tmrBattleMode2) and 254
            End If
        Else
            House(CurrentPlayer).BattleState(HouseBattle1).OtherHouse = MyHouse
            TimerFlags(tmrBattleMode2) = 0
            HouseBattle2 = 0
        End If

        If HouseBattle1 = 0 And HouseBattle2 = 0 Then
            'PlayerMode = 0
            SetPlayerMode 0
            House(CurrentPlayer).CheckBattleReady
            CheckForHotkOrITReady

            If bMultiBallMode = False Then
                TimerFlags(tmrUpdateBattleMode) = 0     ' Disable the timer that updates the Battle Alternate Scene
                DMDResetScoreScene
                House(CurrentPlayer).SetUPFState False
            End If

            If bUseFlexDMD And bComplete = False Then
        if HouseCompleteWithDireWolf = True then debug.print "**** House completed with DireWolf (EndBattleMode)"
                tmrBattleCompleteScene.Interval=3000
                tmrBattleCompleteScene.Enabled = 1
                tmrBattleCompleteScene.UserValue = MyHouse
            End If
        ElseIf bMultiBallMode = False Then ' Another house battle is still active and no MB, so regenerate battle scene
            DMDCreateAlternateScoreScene HouseBattle1,0
        End If
        SetTopGates
        PlayModeSong
        SetPlayfieldLights
    End Sub

    ' Called by the timer when the mode timer has expired
    Public Sub BattleTimerExpired
        If MyHouse = Martell And State = 2 Then
            DoCompleteMode 0
        Else
            EndBattleMode
        End if
    End Sub

    Public Sub BEndHurryUp
        If MyHouse = Martell And State = 2 Then
            DoCompleteMode 0
        Else
            If MyHouse = Targaryen And (State = 10 Or State = 11) Then
                'Dragon Fire!!
                DMDPlayHitScene "got-targaryenbattlehit7","gotfx-targaryenbattlehit3",3.5,"DRAGON","FIRE","",0,4
                State = 10
                TGShotCount = 0
                DrogonHits = DrogonHits - 1
                GenerateRandomShots
                TGStartHurryUp
                SetModeLights
                UpdateBattleScene
                DoTargaryenFireLightSeq
            Else
                EndBattleMode
            End If
        End if
    End Sub

    ' Add time to a mode
    Public Sub AddTime(t)
        If MyHouse = Targaryen Or (MyHouse = Martell And State = 2) Then ' Add value to HurryUp
            HurryUpValue = HurryUpValue + (HurryUpChange * 5 * t)
        Else    ' timer
            If MyHouse = HouseBattle2 Then
                TimerTimestamp(tmrBattleMode2) = TimerTimestamp(tmrBattleMode2)+(t*10)
            Else
                TimerTimestamp(tmrBattleMode1) = TimerTimestamp(tmrBattleMode1)+(t*10)
            End If
        End If
        ' TODO - play the "Add Time" scene
    End Sub

    ' Upper Playfield Castle loop.
    ' When not lit, contributes to mode completion
    Public Sub RegisterCastleHit
        Dim i
        Select Case MyHouse
            Case Stark: Me.RegisterHit Stark : Me.AddTime 5
            Case Baratheon: RegisterSpinnerHit : Me.AddTime 5
            Case Lannister
                If ShotMask = 0 Then
                    ' Register hit against first lit gold target
                    For i = 0 to 4
                        If (GoldShotMask And 2^i) > 0 then RegisterGoldHit i : Exit For
                    Next
                Else
                    ' Register hit against first shot
                    For i = 1 to 7
                        If (ShotMask And 2^i) > 0 Then RegisterHit i : Exit For
                    Next
                End If
                Me.AddTime 5
            Case Greyjoy
                ' Register hit against first shot
                For i = 1 to 7
                    If (ShotMask And 2^i) > 0 Then RegisterHit i : Exit For
                Next
            Case Tyrell: If (State MOD 2) = 0 Then Me.RegisterHit Tyrell Else Me.RegisterHit Stark
            Case Martell: If (State = 1) Then Me.RegisterHit Greyjoy Else Me.AddTime 5
            Case Targaryen: Me.AddTime 5
        End Select
    End Sub

    ' Upper Playfield Castle Loop completes this Mode when lit
    ' But no score awarded, other than what they get from the castle itself
    ' TODO: Castle Complete rules are actually kinda complicated:
    '   - If neither house in stacked battle is Targaryen, complete both
    '   - If stacked battles are running and one is Targaryen:
    '       - If this house isn't Targaryen, complete it only if the Targaryen house hasn't been completed with a Castle yet
    '       - If this house is Targaryen, complete this Level, unless it's Level 3
    '   - If no stack is running, complete this house/Level as long as it's not Level 3
    Public Sub RegisterCastleComplete
        If MyHouse <> Targaryen Then
            bComplete = True
            House(CurrentPlayer).HouseCompleted MyHouse
            SetLightColor HouseSigil(MyHouse),HouseColor(SelectedHouse),1
        Else
            If TGLevel = 2 Then Exit Sub ' Can't use Castle to finish Targaryen
            If State < 5 Then
                State = 5 : ShotMask = 36
            ElseIf State < 9 Then State = 9
            End If
            If SelectedHouse <> GreyJoy Then State=State+1 : ShotMask = 10
        End If
        If TotalScore = 0 Then TotalScore = 15000000 : AddScore 15000000
        EndBattleMode
    End Sub

    ' Some battles involve the spinner
    Public Sub RegisterSpinnerHit
        If MyHouse <> Baratheon Then Exit Sub
        ' ShotMask = ShotMask And 239 ' turn off bit 4
        HouseValue = HouseValue + HouseValueIncrement
        If HouseValue > 150000000 Then HouseValue = 150000000
        DMDBaratheonSpinnerScene HouseValue
        UpdateBattleScene

        If State = 1 And HouseValue > 20000000 Then
            State = 2
            If SelectedHouse = Greyjoy Then ShotMask = 144 Else ShotMask = 128
            SetModeLights
        End If
    End Sub

    ' Some battles need to know about individual target hits
    ' Right now we don't care about individual targets, just which bank. 0 = Left, 1 = Right
    Public Sub RegisterTargetHit(tgt)
        If MyHouse = Baratheon And State = 3 And tgt = 0 Then
            ' Mode completed!
            DoCompleteMode Baratheon
        ElseIf MyHouse = Targaryen And (ShotMask And 4) > 0 And tgt = 0 Then
            Me.RegisterHit Baratheon
        ElseIf MyHouse = Targaryen And (ShotMask And 32) > 0 And tgt = 1 Then
            Me.RegisterHit Tyrell
        ElseIf tgt = 0 And MyHouse = Baratheon Then
            ResetDropTargets
        ElseIf tgt = 1 Then
            Me.RegisterHit Tyrell
        End if
    End Sub

    ' Lannister battle mode needs to know about gold target hits
    Public Sub RegisterGoldHit(tgt)
        Dim litmask,litshots,growval
        If MyHouse <> Lannister Then Exit Sub
        If (GoldShotMask And 2^tgt) = 0 Then
            PlaySoundVol "gotfx-litgold",VolDef
            If House(CurrentPlayer).HasAbility(Lannister) Then AddGold 15 Else AddGold 5
            Exit Sub   ' Gold target wasn't lit
        End If
        If House(CurrentPlayer).HasAbility(Lannister) Then AddGold 100 Else AddGold 15
        PlaySoundVol "gotfx-goldcoin",VolDef
        GoldShotMask = GoldShotMask Xor 2^tgt
        Select Case tgt
            Case 0: litmask = 144
            Case 1: litmask =  136
            Case 2,3: litmask = 10
            Case 4: litmask = 66
        End Select
        litshots = "+1 SHOT LIT"
        If (ShotMask And litmask) = litmask Then
            litshots = ""
        ElseIf (ShotMask And litmask) = 0 Then
            litshots = "+2 SHOTS LIT"
        End If
        ShotMask = ShotMask Or litmask
        HouseValue = HouseValue + 125000
        SetModeLights
        DMDPlayHitScene "got-lannisterbattlehit1","gotfx-lannisterbattlehit1",1.5,"HOUSE  LANNISTER  VALUE  GROWS",HouseValue,litshots,0,2
    End Sub

    ' Called when the 10 second timer runs down
    Public Sub MartellTimer: CompletedShots = 0: UpdateBattleScene: End Sub

    Public Sub TGStartHurryUp
        If MyHouse <> Targaryen Then Exit Sub
        Me.SetTGHurryUpValue
        MyBattleScene.GetLabel("TGHurryUp").Visible = True
        StartTGHurryUp HouseValue,MyBattleScene,5
    End Sub

    Public Sub SetTGHurryUpValue
        If MyHouse <> Targaryen Then Exit Sub
        Select Case State
            Case 1,2,3,4: HouseValue = 4000000 + 2000000*State ' verified
            Case 5,6,7,8: HouseValue = 3000000*State - 6000000 ' verified
            Case  9: HouseValue = 15000000
            Case 10: HouseValue = 9000000+3000000*TGShotCount ' real vals: 9M, 12M, 15M,
            Case 11: HouseValue = 12000000
        End Select
        ' There's a bit of a race condition here, since the battle scene gets set up
        ' before the HurryUp is started. So stuff the right value into the Targaryen HurryUpValue
        ' so that the scene renders correctly at the beginning
        TGHurryUpValue = HouseValue
    End Sub

    Public Function TGLevel
        If MyHouse <> Targaryen Then TGLevel = 0:Exit Function
        If State < 5 Then
            TGLevel = 0
        Elseif State < 9 Then
            TGLevel = 1
        Else
            TGLevel = 2
        End If
    End Function

    ' Generate 3 lit random shots for Targaryen battle mode
    Public Sub GenerateRandomShots
        Dim i,j
        ShotMask = 0
        For i = 1 to 3
            Do
                j = RndNbr(6)
            Loop While (ShotMask And 2^j) > 0
            ShotMask = ShotMask Or 2^j
            If j = Tyrell Then
                ' Reset Wildfire targets
                bWildfireTargets(0) = False: bWildfireTargets(1) = False
                li80.BlinkInterval = 50
                li83.BlinkInterval = 50
                li80.State = 2
                li83.State = 2
            End if
        Next
    End Sub

'  bit   data (Label name)
'   0    Score
'   1    Ball
'   2    Credits
'   3    combo1 thru 5
'   4    HurryUp
'   5    BattleTimer1 (tmr1)
'   6    BattleTimer2 (tmr2)
'   7    MartellBattleTimer (tmr3)
'  Create a Battle Score scene specific to this House Battle
    Dim MyBattleScene
    Dim bSmallBattleScene
    Public Sub CreateBattleProgressScene(ByRef BattleScene)
        Dim tinyfont,ScoreFont,line3,x3,x4,y4,i

        Set tinyfont = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)

        If MyHouse = Targaryen Then
            ' Everything about the Targaryen scene is different
            Dim Dragon
            Select Case Me.TGLevel
                Case 0: Dragon = "VISERION"
                Case 1: Dragon = "RHAEGAL"
                Case 2: Dragon = "DROGON"
            End Select

            Me.SetTGHurryUpValue

            BattleScene.AddActor FlexDMD.NewLabel("Score",FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0),FormatScore(Score(CurrentPlayer)))
            BattleScene.GetLabel("Score").SetAlignedPosition 0,0,FlexDMD_Align_TopLeft
            BattleScene.AddActor FlexDMD.NewLabel("TGHurryUp",tinyfont,HouseValue) 'Placeholder value to ensure text is centered
            BattleScene.GetLabel("TGHurryUp").SetAlignedPosition 0,19,FlexDMD_Align_TopLeft
            BattleScene.AddActor FlexDMD.NewLabel("dragon",tinyfont,Dragon)
            BattleScene.GetLabel("dragon").SetAlignedPosition 127,7,FlexDMD_Align_TopRight
            If Me.TGLevel = 2 Then
                BattleScene.AddActor FlexDMD.NewFrame("health")
                With BattleScene.GetFrame("health")
                    .Thickness = 1
                    .SetBounds 61, 0, 66, 4      ' Health box is 62W by 4H, and offset by 65 pixels
                End With
                For i = 0 to 15
                    BattleScene.AddActor FlexDMD.NewFrame("health"&i)
                    With BattleScene.GetFrame("health"&i)
                        '.Thickness = 1
                        .SetBounds 62+i*4, 1, 4,2      ' Each Health unit is 4W by 2H, and offset by 65+i*4 pixels
                        If i >= 15-Me.DrogonHits Then .Thickness = 0 Else .Thickness = 1
                    End With
                Next
            End If
        Else
            ' x3,y3 = line3 location
            y4 = 20 ' x4,y4 = timer location
            Select Case MyHouse
                Case Stark,Baratheon: line3 = "VALUE = "&FormatScore(HouseValue) : x3 = 40: x4 = 40: y4 = 22
                Case Lannister,Greyjoy: line3 = "SHOTS = " & 5-CompletedShots : x3 = 20: x4 = 56
                Case Tyrell: x4 = 40
                Case Martell: line3 = "SHOOT ORBITS":x3=25:x4 = 60
            End Select

            BattleScene.AddActor FlexDMD.NewLabel("obj",tinyfont,BattleObjectivesShort(MyHouse))
            BattleScene.AddActor FlexDMD.NewLabel("tmr1",tinyfont,Int((TimerTimestamp(tmrBattleMode1)-GameTimeStamp)/10))
            If MyHouse <> Tyrell Then
                BattleScene.AddActor FlexDMD.NewLabel("line3",tinyfont,line3)
                BattleScene.GetLabel("line3").SetAlignedPosition x3,16,FlexDMD_Align_Center
            End If
            BattleScene.GetLabel("tmr1").SetAlignedPosition x4,y4,FlexDMD_Align_Center
            BattleScene.GetLabel("obj").SetAlignedPosition 40,3,FlexDMD_Align_Center

            If MyHouse = Martell Then
                BattleScene.AddActor FlexDMD.NewLabel("tmr3",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0),"10")
                BattleScene.GetLabel("tmr3").SetAlignedPosition 40,13,FlexDMD_Align_Center
                If CompletedShots = 0 Or State > 1 Then BattleScene.GetLabel("tmr3").Visible = 0
                BattleScene.AddActor FlexDMD.NewLabel("HurryUp",tinyfont,"20000000") 'Placeholder value to ensure text is centered
                BattleScene.GetLabel("HurryUp").SetAlignedPosition 32,13,FlexDMD_Align_Center
                If State < 2 Then
                    BattleScene.GetLabel("HurryUp").Visible = 0
                Else
                    BattleScene.GetLabel("line3").SetAlignedPosition 30,21,FlexDMD_Align_Center
                End If
            Else
                ' Every other house has the score showing
                BattleScene.AddActor FlexDMD.NewLabel("Score",FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0),FormatScore(Score(CurrentPlayer)))
                BattleScene.GetLabel("Score").SetAlignedPosition 40,9,FlexDMD_Align_Center
            End if
        End If

        For i = 1 to 5
            BattleScene.AddActor FlexDMD.NewLabel("combo"&i, FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbBlack, 1), "0")
        Next
        Set MyBattleScene = BattleScene
        bSmallBattleScene = False
    End Sub

    Public Function CreateSmallBattleProgressScene(ByRef BattleScene, n)
        Dim tinyfont,ScoreFont,line2,x3,x4,y4,i
        'Set ScoreFont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
        Set tinyfont = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbBlack, 0)
        BattleScene.AddActor FlexDMD.NewLabel("obj",tinyfont,HouseToUCString(MyHouse))
        If MyHouse <> Targaryen Then BattleScene.AddActor FlexDMD.NewLabel("tmr"&n,tinyfont,Int((TimerTimestamp(tmrBattleMode1)-GameTimeStamp)/10))

        y4 = 16
        Select Case MyHouse
            Case Stark,Baratheon: line2 = "VALUE = "& vbLf & FormatScore(HouseValue)
            Case Lannister,Greyjoy: line2 = "SHOTS = " & 5-CompletedShots
            Case Tyrell: line2=""
            Case Martell: line2 = "SHOOT ORBITS"
        End Select

        If MyHouse = Martell Then
            BattleScene.AddActor FlexDMD.NewLabel("tmr3",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite,vbWhite, 0),"10")
            BattleScene.GetLabel("tmr3").SetAlignedPosition 32,13,FlexDMD_Align_Center
            If CompletedShots = 0 Or State > 1 Then BattleScene.GetLabel("tmr3").Visible = 0
        End If
        If MyHouse = Martell Then
            BattleScene.AddActor FlexDMD.NewLabel("HurryUp",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite,vbWhite, 0),"20000000") 'Placeholder value to ensure text is centered
            BattleScene.GetLabel("HurryUp").SetAlignedPosition 32,13,FlexDMD_Align_Center
            If State < 2 Then BattleScene.GetLabel("HurryUp").Visible = 0
        ElseIf MyHouse = Targaryen Then
            BattleScene.AddActor FlexDMD.NewLabel("TGHurryUp",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite,vbWhite, 0),"12000000") 'Placeholder value to ensure text is centered
            BattleScene.GetLabel("TGHurryUp").SetAlignedPosition 32,13,FlexDMD_Align_Center
        End If

        If MyHouse <> Targaryen Then
            BattleScene.AddActor FlexDMD.NewLabel("line3",tinyfont,line2)
            BattleScene.GetLabel("line3").SetAlignedPosition 32,16,FlexDMD_Align_Center
            BattleScene.GetLabel("tmr"&n).SetAlignedPosition 62,1,FlexDMD_Align_TopRight
        End If
        BattleScene.GetLabel("obj").SetAlignedPosition 1,1,FlexDMD_Align_TopLeft

        Set MyBattleScene = BattleScene
        bSmallBattleScene = True
    End Function

    ' Timers and score are updated by main score subroutine. We just take care of battle-specific values
    Public Sub UpdateBattleScene
        Dim line3
        if bUseFlexDMD = False Then Exit Sub
        If MyHouse = Martell Then
            FlexDMD.LockRenderThread
            If State = 2 Then
                MyBattleScene.GetLabel("tmr3").Visible = 0
                With MyBattleScene.GetLabel("line3")
                    .Text = "SHOOT RAMPS"
                    .SetAlignedPosition 30,21,FlexDMD_Align_Center
                End with
                MyBattleScene.GetLabel("HurryUp").Visible = 1
            Else
                If CompletedShots = 0 Then
                    MyBattleScene.GetLabel("tmr3").Visible = 0
                    MyBattleScene.GetLabel("line3").SetAlignedPosition 30,16,FlexDMD_Align_Center
                Else
                    MyBattleScene.GetLabel("tmr3").Visible = 1
                    MyBattleScene.GetLabel("line3").SetAlignedPosition 30,21,FlexDMD_Align_Center
                End if
            End if
            FlexDMD.UnlockRenderThread
            Exit Sub
        ElseIf MyHouse <> Targaryen Then
            Select Case MyHouse
                Case Stark,Baratheon
                    If bSmallBattleScene Then
                        line3 = "VALUE = "& vbLf & FormatScore(HouseValue)
                    Else
                        line3 = "VALUE = " & FormatScore(HouseValue)
                    End If
                Case Lannister,Greyjoy: line3 = "SHOTS = " & 5-CompletedShots
                Case Else: Exit Sub
            End Select
            FlexDMD.LockRenderThread
            MyBattleScene.GetLabel("line3").Text = line3
            FlexDMD.UnlockRenderThread
        ElseIf TGLevel = 2 And Not (MyBattleScene.GetFrame("health") Is Nothing) Then
            Dim i
            FlexDMD.LockRenderThread
                For i = 0 to 15
                    If i >= 15-DrogonHits Then MyBattleScene.GetFrame("health"&i).Thickness = 0 Else MyBattleScene.GetFrame("health"&i).Thickness = 1
                Next
            FlexDMD.UnlockRenderThread
        End If
    End Sub

    Public Sub DoBattleCompleteScene
        DoMyBattleCompleteScene
        If OtherHouseBattle <> 0 Then House(CurrentPlayer).BattleState(OtherHouseBattle).DoMyBattleCompleteScene
    End Sub

    Public Sub DoMyBattleCompleteScene
        If bUseFlexDMD = False Then Exit Sub
        Dim scene

        If TotalScore > 0 Then
            Set scene = NewSceneWithVideo("btotal","got-"&HouseToString(MyHouse)&"battlesigil")
            scene.GetVideo("btotalvid").SetAlignedPosition 127,0,FlexDMD_Align_TopRight
            scene.AddActor FlexDMD.NewLabel("ttl",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0),HouseToUCString(MyHouse)&" TOTAL")
            scene.AddActor FlexDMD.NewLabel("bscore",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),FormatScore(TotalScore))
            scene.GetLabel("ttl").SetAlignedPosition 0,10,FlexDMD_Align_Left
            scene.GetLabel("bscore").SetAlignedPosition 40,20,FlexDMD_Align_Center
            DMDEnqueueScene scene,0,2000,3000,4000,"gotfx-battletotal"
            If MyHouse <> Baratheon and bComplete Then PlaySoundVol "say-"&HouseToString(MyHouse)&"battlecomplete",VolCallout
    Else
            ' Fail!!
            If MyHouse = Stark and not HouseCompleteWithDireWolf Then PlaySoundVol "say-starkbattlefailed",VolCallout
        End If
    End Sub

End Class

Function HouseToString(h)
    Select Case h
        Case 0
            HouseToString = ""
        Case Stark
            HouseToString = "stark"
        Case Baratheon
            HouseToString = "baratheon"
        Case Lannister
            HouseToString = "lannister"
        Case Greyjoy
            HouseToString = "greyjoy"
        Case Tyrell
            HouseToString = "tyrell"
        Case Martell
            HouseToString = "martell"
        Case Targaryen
            HouseToString = "targaryen"
        Case 8: HouseToString = "random"
    End Select
End Function

Function HouseToUCString(h)
    Select Case h
        Case 0
            HouseToUCString = ""
        Case Stark
            HouseToUCString = "STARK"
        Case Baratheon
            HouseToUCString = "BARATHEON"
        Case Lannister
            HouseToUCString = "LANNISTER"
        Case Greyjoy
            HouseToUCString = "GREYJOY"
        Case Tyrell
            HouseToUCString = "TYRELL"
        Case Martell
            HouseToUCString = "MARTELL"
        Case Targaryen
            HouseToUCString = "TARGARYEN"
        Case 8: HouseToUCString = "RANDOM"
    End Select
End Function

'**************************************************
' Table, Game, and ball initialization code
'**************************************************

Sub VPObjects_Init
    Dim i,sl
    BumperWeightTotal = 0
    For i = 1 To BumperAwards:BumperWeightTotal = BumperWeightTotal + PictoPops(i)(2): Next
    For each sl in RSling1_BL : sl.visible = 0 : Next
    For each sl in LSling1_BL : sl.visible = 0 : Next
    For each sl in RSling2_BL : sl.visible = 0 : Next
    For each sl in LSling2_BL : sl.visible = 0 : Next
    SetDefaultPlayfieldLights   ' Sets all playfield lights to their default colours
    vpmTimer.AddTimer 1000,"WarmUpDone : UpdateMods : PlaySoundVol ""say-whathouse"",VolCallout '"
End Sub

Sub Game_Init()     'called at the start of a new game
    SetDefaultPlayfieldLights   ' Sets all playfield lights to their default colours
    TurnOffPlayfieldLights
    SetLockbarLight
    ResetBallSearch
    ResetComboMultipliers
End Sub

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
    Dim i

    bGameInPLay = True
    Tilted = False
    GameTimeStamp = 0
    'resets the score display, and turn off attract mode
    StopAttractMode
    bGameJustPlayed = True
    StopSound "got-track-gameover"
    GiOn
    LightSeqPlayfield.UserValue = 0

    bAlternateScoreScene=False
    ScoreScene = Empty

    ChooseHouseScene = Empty

    ' Just in case
    BallsOnPlayfield = 0
  RealBallsInLock=0
    BallsInLock = 0

  '1.1.24 - possible state issue on new game
  bLockIsLit = False
  EBisLit = 0
  bPictoEBAwarded = False
  bLoLLit = False
  bLoLUsed = False
  TotalGold = 0
  CurrentGold = 0
  TotalWildfire = 0
  CurrentWildfire = 0
  bMysteryLit = False
  bSwordLit = False

    bMadnessMB = 0

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    bBattleInstructionsDone = False
    bWildfireTargets(0) = False:bWildfireTargets(1) = False
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
        ReplayAwarded(i) = False
        ITHighScore(i) = 0
        HotkHighScore(i) = 0
        WHCHighScore(i) = 0
        Set House(i) = New cHouse
        Set PlayerState(i) = New cPState
    Next
    if DMDStd(kDMDFet_PlayerHouseChoice)=8 Then RandomHouse = RndNbr(7)

    ' initialise any other flags
    Tilt = 0

    ' initialise Game variables
    Game_Init()

    DMDBlank

    tmrGame.Interval = 100
    tmrGame.Enabled = 1

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

    'Reset any table specific
    ResetNewBallVariables

    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True

    'This table doesn't use a skill shot
    bSkillShotReady = False

    bHotkMBReady = False
    bHurryUpActive = False
    bBWMultiballActive = False
    bBlackwaterSJPMode = False
    bCastleMBActive = False
    bWallMBActive = False
    bWHCMBActive = False
    bAddABallUsed = False
    bInlanes(0) = False : bInlanes(1) = False
    WHCFlipperFrozen = False

    ' The below settings will be overwritten by PlayerState.Restore if this isn't Ball 1
    WildfireTargetsCompleted = 0
    LoLTargetsCompleted = 0
    CompletedHouses = 0
    TotalGold = 0
    CurrentGold = 0
    TotalWildfire = 0
    CurrentWildfire = 0
    SwordsCollected = 0
  CastlesCollected = 0 ' 1.1.23
    bLockIsLit = False
    EBisLit = 0
    bPictoEBAwarded = False
    BallsInLock = 0
    BWMultiballsCompleted = 0
    WallMBCompleted = 0
    WallMBLevel = 0
    bWallMBReady = False
    bLoLLit = False
    bLoLUsed = False
    bHotkMBDone = False
    bITMBReady =False
    bITMBDone = False
    bITMBActive = False
    bBonusHeld = False
    bWildfireTargets(0) = False:bWildfireTargets(1) = False
    ' End restore

    bMysteryLit = False
    bSwordLit = False
    bElevatorShotUsed = False
    bCastleShotAvailable = False
    SetTopGates
    MoveDiverter 1

    ' Drop the right ramp target
    PlaySoundAt "Drop_Target_Down_" & RndNbr(6), Target90
    DTDrop 4 : TargetState(4) = 1

    HouseBattle1 = 0 : HouseBattle2 = 0

    ' Reset Combo multipliers
    ResetComboMultipliers

    FreezeAllGameTimers ' First score will thaw them

    if (House(CurrentPlayer).MyHouse = 0) Then
        Select Case DMDStd(kDMDFet_PlayerHouseChoice)
            Case 0
                'PlayerMode = -1
        SetPlayerMode -1
                SelectedHouse = 1
                FlashShields SelectedHouse,1
                SetLockbarLight
                If CurrentPlayer = 1 Then
                    PlaySoundVol "say-choose-your-house2",VolCallout
                Else
                    PlaySoundVol "player"&CurrentPlayer,VolDef
                    vpmTimer.AddTimer 1500,"PlaySoundVol ""say-choose-your-house2"",VolCallout '"
                End If
                ChooseHouse 0
            Case 8
                House(CurrentPlayer).MyHouse = RandomHouse
                House(CurrentPlayer).ResetLights
                'PlayerMode = 0
        SetPlayerMode 0
                DMDFlush
                DMDLocalScore
                SetLockbarLight
            Case Else
                House(CurrentPlayer).MyHouse = DMDStd(kDMDFet_PlayerHouseChoice)
                House(CurrentPlayer).ResetLights
                'PlayerMode = 0
        SetPlayerMode 0
                DMDFlush
                ' TODO: Display additional text about house chosen on ball launch
                DMDLocalScore
                SetLockbarLight
        End Select
    Else
        PlayerState(CurrentPlayer).Restore
        'PlayerMode = 0
    SetPlayerMode 0
        SelectedHouse = House(CurrentPlayer).MyHouse
        House(CurrentPlayer).ResetForNewBall
    End If
    SetPlayfieldLights
    CheckForHotkOrITReady
    PlayModeSong
End Sub

Sub ResetNewBallVariables() 'reset variables for a new ball or player
    dim i
    'turn on or off the needed lights before a new ball is released
    ResetPictoPops
    ResetDropTargets
     ' Top lanes start out off on the Premium/LE
    For i = 0 to 1 : bTopLanes(i) = False : bInlanes(i) = False : Next
    'playfield multipiplier
    pfxtimer.Enabled = 0
    PlayfieldMultiplierVal = 1
    PFMState = 0
    SpinnerLevel = 1
    AccumulatedSpinnerValue = 0
    SpinnerValue = 500 + (1+BallsPerGame-BallsRemaining(CurrentPlayer))*2000 ' Appears to start at 2500 on ball 1 and 4500 on ball 2
    SetPFMLights
    ' Gold targets
    For i = 0 to 5 : bGoldTargets(i) = False : Next

    bBallReleasing = False
    bWildfireLit = False
    bPlayfieldValidated = False
    bBattleCreateBall = False
    bHotkIntroRunning = False
    bMysterySJPMode = 0
End Sub

Sub SetPlayfieldLights
    If bHotkMBReady Or bITMBReady Then CheckForHotkOrITReady : Exit Sub
    TurnOffPlayfieldLights
    If PlayerMode = 1 or PlayerMode = -2.1 or bITMBActive Then
        SetModeLights       ' Set Sigils and Shields for battle, as well as UpperPF lights
    ElseIf PlayerMode = 0 Or PlayerMode = 2 Then
        If House(CurrentPlayer).MyHouse > 0 Then House(CurrentPlayer).ResetLights    ' Set Sigils and Shields for normal play
        SetGoldTargetLights
    ElseIf PlayerMode = -1 Then 'ChooseHouse
        FlashShields SelectedHouse,True
    End If
    SetLockLight
    SetBumperLights
    SetOutlaneLights
    SetInlaneLights
    SetTopLaneLights
    SetWildfireLights
    SetPFMLights
    SetShootAgainLight
    SetLockbarLight
    If PlayerMode = 2 Then Exit Sub

    ' Lights below here stay off during a WiC HurryUp
    SetMysteryLight
    SetSwordLight
    SetComboLights
    SetTargetLights
    SetBatteringRamLights
    setEBLight
    SetCastleBlackLight
    SetBattleReadyLight
    House(CurrentPlayer).SetUPFState False  ' This sets the UPFFlasher state as well
End Sub

Sub QueueDireWolf
  dim scene
  If bUseFlexDMD Then
    Set scene = NewSceneWithVideo("dwolf","got-direwolf")
    scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),HouseToUCString(HouseBattle1)&vbLf&"COMPLETED"&vbLf&"5,000,000")
    scene.GetLabel("txt").SetAlignedPosition 88,16,FlexDMD_Align_Center
    DMDEnqueueScene scene,1,1500,2000,3000,"gotfx-direwolf"
  End If
End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    'PlaySoundAt SoundFXDOF("fx_Ballrel", 105, DOFPulse, DOFContactors), BallRelease
    RandomSoundBallRelease BallRelease
    DOF 105, DOFPulse
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multiball flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If (BallsOnPlayfield-RealBallsInLock = 2) And LastSwitchHit = "OutlaneSW" And (bBallSaverActive = True Or bLoLLit = True) Then
        ' Preemptive ball save
        bAutoPlunger = True
        If bBallSaverActive = False Then bLoLLit = False: SetOutlaneLights
    ElseIf (BallsOnPlayfield-RealBallsInLock > 1 Or bMadnessMB > 0) Then
        DOF 143, DOFPulse
        bMultiBallMode = True
        bAutoPlunger = True
    End If
End Sub


Sub DoBallSaved(l)
    ' create a new ball in the shooters lane
    ' we use the Addmultiball in case the multiballs are being ejected
    AddMultiball 1
    ' we kick the ball with the autoplunger
    bAutoPlunger = True
    bBallSaved = True
    If l Then
        PlaySoundVol "gotfx-lolsave",VolDef
        bLoLLit = False
    Else
        if bMultiBallMode = False Then
            DisableBallSaver
            DMDDoBallSaved
        End If
    End If
End Sub

Sub EnableBallSaver(seconds)
    'debug.print "EnableBallSaver called"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
    SetGameTimer tmrBallSave,seconds*10
    SetGameTimer tmrBallSaveSpeedUp,(seconds-5)*10
    TimerFlags(tmrBallSaverGrace) = 0
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.BlinkPattern = "10"
    SetShootAgainLight
  DOF 137, DOFOn
End Sub

Sub DisableBallSaver
    TimerFlags(tmrBallSave) = 0
    TimerFlags(tmrBallSaveSpeedUp) = 0
    TimerFlags(tmrBallSaverGrace) = 0
    BallSaveTimer
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaveTimer
    ' clear the flag
    SetGameTimer tmrBallSaverGrace, DMDStd(kDMDStd_BallSaveExtend)*10  ' 3 second grace for Ball Saver
    SetShootAgainLight
End Sub

Sub BallSaverGraceTimer
    bBallSaverActive = False
    SetShootAgainLight
  DOF 137, DOFOff
End Sub

Sub BallSaverSpeedUpTimer
  ' debug.print "BallSaverSpeedUpTimer called"
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
End Sub

Sub TargaryenFreezeUpdateDMD
  'used to update the DMD for targaryan freeze ability
  'this is called 1 sec after the score scene is last updated
  DMDLocalScore
End Sub

Sub AddScore(points)
    AddScoreNoX points*PlayfieldMultiplierVal
End Sub

Sub AddScoreNoX(points)
  dim prevScore : prevScore = Score(CurrentPlayer)
    tmrBallSearch.Enabled = False ' Reset Ball Search
    tmrBallSearch.Interval = 20000
  BallSearchCnt=0
  tmrBallSearch.Enabled = True
    'TODO: GoT allows you to hit certain targets without validating the playfield and starting timers
    ThawAllGameTimers
    If bPlayfieldValidated = False Then bPlayfieldValidated = True : PlayModeSong
    If (Tilted = False) Then
        ' if there is a need for a ball saver, then start off a timer
        ' only start if it is ready, and it is currently not running, else it will reset the time period
        If(bBallSaverReady = True)AND(DMDStd(kDMDStd_BallSave) <> 0)And(bBallSaverActive = False)Then
            EnableBallSaver DMDStd(kDMDStd_BallSave)
        End If

        If Score(CurrentPlayer) < ReplayScore And Score(CurrentPlayer) + points > ReplayScore Then
            ReplayAwarded(CurrentPlayer) = True
            ' If Bonus is being display, slightly delay Replay scene so it plays after bonus scene
            If bBonusSceneActive Then vpmTimer.AddTimer 1500,"DoAwardReplay '" : Else DoAwardReplay
        End If

    Score(CurrentPlayer) = Score(CurrentPlayer) + points

    'if score rolls over 1 billion, reset the score scene to allow smaller font
    if Score(CurrentPlayer) > 1000000000 and prevScore < 1000000000 Then
      debug.print "> 1 bil, resetting DMD score scene"
      DMDResetScoreScene
    end If

        ' update the score display
        DMDLocalScore
    End If
End Sub

sub ResetBallSearch()
  if BallSearchResetting then Exit Sub  ' We are resetting jsut exit for now
  'debug.print "Ball Search Reset"
  tmrBallSearch.Enabled = False ' Reset Ball Search
    tmrBallSearch.Interval = 20000
  BallSearchCnt=0
  tmrBallSearch.Enabled = True
End Sub

dim BallSearchResetting:BallSearchResetting=False

Sub tmrBallSearch_Timer() ' We timed out
  ' See if we are in mode select, a flipper is up that might be holding the ball or a ball is in the lane

  'debug.print "Ball Search"
  if bGameInPlay and _
        PlayerMode >= 0 and _
    bBallInPlungerLane = False and _
        bMysteryAwardActive = False And _
        hsbModeActive = 0 And _
        tmrEndOfBallBonus.Enabled = False And _
    LeftFlipper.CurrentAngle = LeftFlipper.StartAngle and _
    RightFlipper.CurrentAngle = RightFlipper.StartAngle Then

    debug.print "Ball Search - NO ACTIVITY " & BallSearchCnt

    BallSearchResetting=True
    BallSearchCnt = BallSearchCnt + 1
        tmrBallSearch.Interval = 10000
        DMD "_", "BALL SEARCH", "",eNone,eNone,eNone,3000,true,""
        If BallSearchCnt = 1 Then
            DTDrop 1 : DTDrop 2 : DTDrop 3
            vpmTimer.AddTimer 1000,"ResetDropTargets '"
            KickerIT.Kick 180,3
        ElseIf BallSearchCnt = 2 Then
            ' Not on the playfield - try the lock wall
            ReleaseLockedBall 0
        Elseif BallSearchCnt >= 3 Then
      dim Ball
      debug.print "--- listing balls ---"
      For each Ball in GetBalls
        Debug.print "Ball: (x=" & Ball.x & ",y=" & Ball.y & ",z=" & Ball.z & ")"
      Next
      debug.print "--- listing balls ---"

          ' Delete all balls on the playfield and add back a new one
      BallsOnPlayfield = 0 : RealBallsInLock = 0
      for each Ball in GetBalls
        Ball.DestroyBall
      Next
      BallSearchCnt = 0

      AddMultiball(1)
            DMD "_", "BALL SEARCH FAIL", "",eNone,eNone,eNone,3000,true,""
      Exit sub
    End if

    vpmtimer.addtimer 3000, "BallSearchResetting = False '"
  Else
    ResetBallSearch
  End if
End Sub

'************************************
' End of ball/game processing
'************************************

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded. This is a mini state machine. State is
' stored in timer.UserValue. States:
' "BONUS" 0.25s
' Base Bonus 0.45s
' X Houses complete 0.45s
' Swords 150k each 0.45s
' Castles 0.45s
' Gold (405=121500) 0.45s  e.g “405 GOLD” on line 1, “121,500” on line 2
' Wildfire (22=121000) 0.55s
' Then bonus X times all added together 1.2s
' Then Total bonus. 1.4s
'
' Note, there are 7 "notes" in the ROM, but only 6 bonuses. Perhaps note 7 is reserved.
' In any case, we skip over it below, in case we want it in the future.

Dim tmpBonusTotal
dim bonusCnt
Dim BonusScene
Dim bBonusSceneActive
Sub tmrEndOfBallBonus_Timer()
    Dim line1,line2,ol,skip,font,i,j,incr,line1held
    ol = False
    skip = False
    incr = 1 : i = 0
    tmrEndOfBallBonus.Enabled = False
    tmrEndOfBallBonus.Interval = 500
    j = Int(tmrEndOfBallBonus.UserValue)
    ' State machine based on GoTG line 2549 onwards
    Select Case j
        Case 0
            bonusCnt = 0
            StopSound Song : Song = ""
            line1 = "BONUS"
            ol = True
            bBonusSceneActive = True
            'tmrEndOfBallBonus.Interval = 250
        Case 1
            line1 = "BASE BONUS" : line2 = FormatScore(BonusPoints(CurrentPlayer))
            BonusCnt = BonusPoints(CurrentPlayer)
        Case 2
            If CompletedHouses > 0 Then
                line1 = CompletedHouses & " HOUSES COMPLETE":line2= FormatScore(175000*CompletedHouses)
                BonusCnt = BonusCnt + (175000*CompletedHouses)
            Else
                Skip = True
            End If
        Case 3
            If SwordsCollected > 0 Then
                line1 = SwordsCollected & " SWORD"
                If SwordsCollected > 1 Then line1 = line1&"S"
                line2 = FormatScore(150000*SwordsCollected)
                BonusCnt = BonusCnt + (150000*SwordsCollected)
            Else
                Skip = True
            End If
        Case 4
            If CastlesCollected > 0 Then
                line1 = CastlesCollected & " CASTLE"
                If CastlesCollected > 1 Then line1 = line1&"S"
                line2 = FormatScore(7500000*CastlesCollected)
                BonusCnt = BonusCnt + (7500000*CastlesCollected)
            Else
                Skip = True
            End If
        Case 5
            If TotalGold > 0 Then
                line1 = FormatScore(TotalGold) & " GOLD" : line2 = FormatScore(300*TotalGold)
                BonusCnt = BonusCnt + (300*TotalGold)
            Else
                Skip = True
            End If
        Case 6
            If TotalWildfire > 0 Then
                line1 = FormatScore(TotalWildfire) & " WILDFIRE" : line2 = FormatScore(5500*TotalWildfire)
                BonusCnt = BonusCnt + (5500*TotalWildfire)
                tmrEndOfBallBonus.Interval = 600
            Else
                Skip = True
            End If
        Case 8
            i = Int(tmrEndOfBallBonus.UserValue*100) - 799
            line1 = i&"X" : line2 = FormatScore(i*BonusCnt)
            If i = BonusMultiplier(CurrentPlayer) Then
                tmrEndOfBallBonus.Interval = 1200
                incr = 0
                tmrEndOfBallBonus.UserValue = 9
            Else
                if i = 1 then tmrEndOfBallBonus.Interval = 175 Else tmrEndOfBallBonus.Interval = 125
                incr = 0.01
            End If
        Case 9
            If bBonusHeld = 2 Then
                line1 = "HELD BONUS" : line2 = FormatScore(BonusHeldPoints(CurrentPlayer))
                bBonusHeld = 0
                incr = 0.1
            Else
                line1 = "TOTAL BONUS" : line2 = FormatScore(BonusCnt * BonusMultiplier(CurrentPlayer))
                PlayfieldMultiplierVal = 1
                Dim t : t = ReplayAwarded(CurrentPlayer)
                AddScore (BonusCnt * BonusMultiplier(CurrentPlayer))+BonusHeldPoints(CurrentPlayer)
                if bBonusHeld = 0 Then BonusHeldPoints(CurrentPlayer) = 0
                if bBonusHeld = 1 Then
                    BonusHeldPoints(CurrentPlayer) = BonusCnt * BonusMultiplier(CurrentPlayer)
                    line1held = "HELD BONUS" & vbLf & FormatScore(BonusHeldPoints(CurrentPlayer))
                    bBonusHeld = 2
                    If BallsRemaining(CurrentPlayer) = 1 And ExtraBallsAwards(CurrentPlayer) = 0 Then ' on last ball
                        line1held = "HELD BONUS" & vbLf & FormatScore(BonusHeldPoints(CurrentPlayer))
                        AddScore BonusHeldPoints(CurrentPlayer)
                        bBonusHeld = 3
                    End if
                End If
                If t <> ReplayAwarded(CurrentPlayer) Then tmrEndOfBallBonus.Interval = 5700 Else tmrEndOfBallBonus.Interval = 1700
                incr = 0
                tmrEndOfBallBonus.UserValue = 10
            End If
        Case 10
            bBonusSceneActive = False
            vpmtimer.addtimer 100, "EndOfBall2 '"
            Exit Sub
        Case Else
            Skip = True
    End Select

    If (LFPress And RFPress) And j > 0 And j < 9 Then Skip = True

    If Skip Then
        tmrEndOfBallBonus.Interval = 10
    Else
        ' Do Bonus Scene
        If bUseFlexDMD Then
            If j=8 And i > 1 Then
                FlexDMD.LockRenderThread
                With BonusScene.GetLabel("line1")
                    .Text = line1 & vbLf & line2
                    .SetAlignedPosition 64,16,FlexDMD_Align_CENTER
                End With
                FlexDMD.UnlockRenderThread
                PlaySoundVol "gotfx-spike-count8",VolDef
            Else
                Set BonusScene = FlexDMD.NewGroup("bonus")
                If ol Then
                    Set font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbWhite, 0)
                Else
                    line1 = line1 & vbLf & line2
                    Set font = FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite, vbWhite, 0)
                End if
                BonusScene.AddActor FlexDMD.NewFrame("bonusbox")
                With BonusScene.GetFrame("bonusbox")
                    .Thickness = 1
                    .SetBounds 0, 0, 128, 32
                End With
                BonusScene.AddActor FlexDMD.NewLabel("line1",font,line1)
                BonusScene.GetLabel("line1").SetAlignedPosition 64,16,FlexDMD_Align_CENTER
                If j=9 And bBonusHeld >= 2 Then
                    Select Case bBonusHeld
                        Case 2
                            ' Show the "held" image after a short delay
                            BonusScene.AddActor FlexDMD.NewImage("held","got-bonusheld.png")
                            Dim bh
                            Set bh = BonusScene.GetImage("held")
                            If Not (bh Is Nothing) Then
                                bh.SetAlignedPosition 96,16,FlexDMD_Align_CENTER
                                bh.visible = 0
                                DelayActor bh,0.450,True
                            End If
                        Case 3
                            ' On last ball, show Held Bonus instead
                            DelayActor BonusScene.GetLabel("line1"),0.450,False
                            BonusScene.AddActor FlexDMD.NewLabel("line1held",font,line1held)
                            With BonusScene.GetLabel("line1held")
                                .SetAlignedPosition 64,16,FlexDMD_Align_CENTER
                                .visible = 0
                            End With
                            DelayActor BonusScene.GetLabel("line1held"),0.450,true
                    End Select
                    vpmTimer.addTimer 450,"PlaySoundVol ""gotfx-spike-count9"",VolDef '"
                End If
                DMDClearQueue
                DMDEnqueueScene BonusScene,0,2000,3000,10,"gotfx-spike-count"&j
            End If
        Else
            If ol Then
                DisplayDMDText "",line1,tmrEndOfBallBonus.Interval
            Else
                DisplayDMDText line1,line2,tmrEndOfBallBonus.Interval
            End If
            PlaySoundVol "gotfx-spike-count"&j, VolDef
        End If
    End If

    tmrEndOfBallBonus.UserValue = tmrEndOfBallBonus.UserValue + incr
    tmrEndOfBallBonus.Enabled = True
End Sub

Sub EndOfBall()
  'debug.print "**End Of Ball"
    Dim AwardPoints, TotalBonus,delay
    AwardPoints = 0
    TotalBonus = 0
    ' the first ball has been lost. From this point on no new players can join in

  'see if player is trying to restart
  if bQuickRestartRequested Then
    debug.print "(EndOfBall) QuickRestartRequested, executing"
    bQuickRestartRequested = False
    bBallIsDraining = False
    DoQuickRestart
    exit sub
  end if

    If bHotkMBReady Then ResetHotkLights

    StopGameTimers
    If HouseBattle2 > 0 Then BattleModeTimer2
    If HouseBattle1 > 0 Then BattleModeTimer1
    EndHurryUp
    'PlayerMode = 0
  SetPlayerMode 0

    Target43.TimerEnabled = 0

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)

    If(Tilted = False)Then
        TurnOffPlayfieldLights
        DMDBlank

        delay = 0
        If tmrMultiballCompleteScene.Enabled Then
            delay=delay+2000
            tmrMultiballCompleteScene_Timer
        ElseIf bITMBActive And ITScore > 0 Then
            DMDIronThroneMBCompleteScene
            delay = delay + 2000
        End If
        If tmrBattleCompleteScene.Enabled Then
            delay=delay+3000
            tmrBattleCompleteScene_Timer
        End If

        ' Delay for a Battle Total screen to be shown
        tmrEndOfBallBonus.Interval = delay + 400
        ' Start the Bonus timer - this timer calls the Bonus Display code when it runs down
    tmrEndOfBallBonus.UserValue = 0
    tmrEndOfBallBonus.Enabled = true
        ' Bonus will start EndOfBall2 when it is done
    Else
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the currentplayer)
'
Sub EndOfBall2()
  'debug.print "**End Of Ball 2"

  dim i
  dim thisMode
    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilted = False
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

    StopDragonWings ' Just in case

  'see if player is trying to restart
  if bQuickRestartRequested Then
    debug.print "(EndOfBall2) QuickRestartRequested, executing"
    bQuickRestartRequested = False
    bBallIsDraining = False
    DoQuickRestart
    exit sub
  end if

    ' has the player won an extra-ball ? (might be multiple outstanding)
    If(ExtraBallsAwards(CurrentPlayer) <> 0) and bBallInPlungerLane=False Then  ' Save Extra ball for later if there is a ball in the plunger lane
        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            SetShootAgainLight
        End If

        If bUseFlexDMD Then
            Dim Scene
            Set Scene = NewSceneWithVideo("shootagain","got-shootagain")
            DMDEnqueueScene Scene,0,4000,8000,2000,"" : DoDragonRoar 1
            vpmTimer.addTimer 1500,"PlaySoundVol ""say-shoot-again"",VolCallout '"
        Else
            DMD "_", CL(1, ("EXTRA BALL")), "_", eNone, eBlink, eNone, 1000, True, "say-shoot-again"
        End if

         ' Save the current player's state - needed so that when it's restored in a moment, it won't screw everything up
        PlayerState(CurrentPlayer).Save
        ' reset the playfield for the new player (or new ball) (also restores player state)
        ResetForNewPlayerBall()
        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0) Then       ' GAME OVER

      ' Turn off DOF so we dont accidently leave it on
      PlaySoundAt SoundFXDOF("Flipper_Left_Down_1", 101, DOFOff, DOFFlippers), LeftFlipper
      LeftFlipper.RotateToStart
            LeftUFlipper.RotateToStart

      PlaySoundAt SoundFXDOF("Flipper_Right_Down_1", 102, DOFOff, DOFFlippers), RightFlipper
      RightFlipper.RotateToStart
      RightUFlipper.RotateToStart

            ' Submit the currentplayers score to the High Score system
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
'all of the same player)
'
Sub EndOfBallComplete()
  'debug.print "**EndOfBallComplete bBallIsDraining is " & bBallIsDraining & " and will be set to false"
  bBallIsDraining = False

  'see if player is trying to restart
  if bQuickRestartRequested Then
    debug.print "(EndOfBallComplete) QuickRestartRequested, executing"
    bQuickRestartRequested = False
    bBallIsDraining = False
    DoQuickRestart
    exit sub
  end if

    Dim NextPlayer
  dim Match
    Dim i

    ' Save the current player's state
    PlayerState(CurrentPlayer).Save

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame > 1)Then
        ' then move to the next player
        DMDResetScoreScene
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer > PlayersPlayingGame)Then
            NextPlayer = 1
            bOnTheFirstBall = False
        End If
    Else
        NextPlayer = CurrentPlayer
        bOnTheFirstBall = False
    End If

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then

        bGameInPLay = False                 ' EndOfGame sets this but need to set early to disable flippers
        DMDBlank

        ' Drop the lock walls to release any locked balls
        SwordWall.collidable = False : MoveSword True
        LockWall.collidable = False : Lockwall.Uservalue = 0 : LockPost False
        BallsInLock = 0 : RealBallsInLock = 0

    bShowMatch = True

        Dim t : t = False
        For i = 1 to 4
            If ReplayAwarded(i) Then ReplayScore = ReplayScore + DMDStd(kDMDStd_AutoReplayStart) : t = True : Exit For
        Next
        If t = False Then ReplayScore = DMDStd(kDMDStd_AutoReplayStart)

    ' Do Match end score code
    Match=10 * INT(RND * 9)
    t=false
        For i = 1 to PlayersPlayingGame
            if BigMod(Score(i), 100) = Match then
                ' This approximates the Match Percent setting, but if a previous player already
                ' matched with this number, we can't not award  it to another player that has the
                ' same match
                if t=true Or RndNbr(10) <= DMDStd(kDMDStd_MatchPCT) Then
                    Credits = Credits + 1
                    vpmtimer.addtimer 5500+i*200, "PlayYouMatched '"
                    t=true
                End If
            End If
        Next

        ' In case the Match Percent test resulted in a match not being awarded even though they matched,
        ' ensure that the numbers don't match if it wasn't awarded
        If t=false Then
            Do
                t=false
                For i = 1 to PlayersPlayingGame
                    if BigMod(Score(i), 100) = Match then t=true : Exit For
                Next
                if t=true Then Match=10 * INT(RND * 9)
            Loop While t
        End If

        DMDDoMatchScene Match

    vpmtimer.addtimer 5000, "if bShowMatch then DMDBlank : EndOfGame() '"

    Else
        ' set the next player
        CurrentPlayer = NextPlayer
        DMDResetScoreScene
    'UpdateNumberPlayers        ' Update the Score Sizes
        ' make sure the correct display is up to date
        AddScore 0

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()

        ' play a sound if more than 1 player
        If PlayersPlayingGame > 1 Then
            i = RndNbr(2)
            PlaySoundVol "say-player" &CurrentPlayer&"-youre-up"&i, VolCallout
            DMD "", CL(1, "PLAYER " &CurrentPlayer), "", eNone, eNone, eNone, 800, True, ""
        End If
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    Dim i
    If bGameInPLay = True then Exit Sub ' In case someone pressed 'Start' during Match sequence

    bGameInPLay = False
    Tilted = True
  bShowMatch = False
  tmrBallSearch.Enabled = False
    TurnOffPlayfieldLights

    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

  ' Drop the lock walls just in case the ball is behind it (just in Case)
  SwordWall.collidable = False : MoveSword True
    LockWall.collidable = False : Lockwall.Uservalue = 0 : LockPost False

    PlaySoundVol "got-track-gameover",VolDef/8
    i = RndNbr(12)
    PlaySoundVol "say-gameover"&i,VolDef
    DMD "_", "GAME OVER", "",eNone,eNone,eNone,6000,true,""

    GiOff

    vpmTimer.AddTimer 5000,"StartAttractMode '"

  if (VRRoom = 1 or VRRoom = 3 or VRRoom = 4 or VRRoom = 5) then VR_StartButtOuter.blenddisablelighting = 0.8: StartButtonTimer.enabled = true  ' start the Front button temp attract lighting

End Sub

Sub PlayYouMatched
  'PlaySoundVol "YouMatchedPlayAgain", VolDef
  DOF 140, DOFOn
  DMDFlush
  DMD "_", CL(1, "CREDITS: " & Credits), "", eNone, eNone, eNone, 4000, True, ""
    KnockerSolenoid : DOF 122, DOFPulse
End Sub

' Plays the right song for the current situation
Sub PlayModeSong
    Dim mysong,i
    mysong = ""
    If PlayerMode = 2 Then
        mysong = "got-track-wic"
    ElseIf bHotkMBReady Or bITMBReady Or bWallMBReady Then
        mysong = "got-track-hotk-ready"
    ElseIf bITMBActive Then
        mysong = "got-track-ironthrone"
    ElseIf bPlayfieldValidated = False Then
        mysong = "got-track-playfieldunvalidated"
    ElseIf PlayerMode = -2 Then
        mysong = "got-track-choosebattle"
    ElseIf bMysteryAwardActive Then
        mysong = "got-track-choosemystery"
    ElseIf bMultiBallMode Then
        If bBWMultiballActive Then
            mysong = "got-track-blackwater"
        ElseIf bCastleMBActive Then
            mysong = "got-track-castlemultiball"
        ElseIf bWallMBActive Then
            mysong = "got-track-wallmb"
        ElseIf bWHCMBActive Then
            mysong = "gotfx-long-wind-blowing"
        ElseIf bHotkMBActive Then
            mysong = "got-track-hotk"
        Elseif bMadnessMB = 2 Then
            mysong = "got-track-targaryen3"
        End If
    ElseIf Playermode = 1 or PlayerMode = -2.1 Then
        If HouseBattle2 <> 0 Then
            mysong = "got-track-stackedbattle"
        ElseIf HouseBattle1 = Targaryen Then mysong = "got-track-targaryen" & House(CurrentPlayer).BattleState(Targaryen).TGLevel+1
        Else  mysong = "got-track-"&HouseToString(HouseBattle1)
        End If
    ElseIf bWallMBReady Then
        mysong = "got-track-wallmb-ready"
    End If
    If mysong = "" Then
        If Song = "got-track1" Or Song = "got-track2" Or Song = "got-track3" Then Exit Sub
        i = BallsPerGame - BallsRemaining(CurrentPlayer) + 1 : if i > 3 then i = 3
        mysong = "got-track"&i
    End If
    If Song <> mysong Then PlaySong mysong
End Sub


' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
End Sub
Sub AddMultiballFast(nballs)
  if CreateMultiballTimer.Enabled = False then
    CreateMultiballTimer.Interval = 100   ' shortcut the first time through
  End If
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
    ' wait if there is a ball in the plunger lane
  CreateMultiballTimer.Interval = 2000
    If bBallInPlungerLane Then
        Exit Sub
    Else
        If (BallsOnPlayfield-RealBallsInLock) < MaxMultiballs Then
            CreateNewBall()
            mBalls2Eject = mBalls2Eject -1
            If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
                Me.Enabled = False
            End If
        Else 'the max number of multiballs is reached, so stop the timer
            mBalls2Eject = 0
            Me.Enabled = False
        End If
    End If
End Sub

Sub StopEndOfBallModes() 'this sub is called after the last ball is drained
    If (TimerFlags(tmrLoLSave) And 1) > 0 Then LoLSaveTimer    ' Baratheon used his button ability. Clear it
End Sub

' Called when the last ball of multiball is lost
Sub StopMBmodes
    tmrMultiballCompleteScene.Interval = 1000
    If Not Tilted Then tmrMultiballCompleteScene.Enabled = 1
    if bBWMultiballActive Then
        BWMultiballsCompleted = BWMultiballsCompleted + 1
        tmrMultiballCompleteScene.UserValue = "DMDBlackwaterCompleteScene"
        House(CurrentPlayer).BWJackpot = 0
        If bBlackwaterSJPMode Then SetGameTimer tmrBlackwaterSJP,30     ' 3 second grace period to hit the SJP
    DOF 127, DOFOff
    End If
    If bCastleMBActive Then
        tmrMultiballCompleteScene.UserValue = "DMDCastleMBCompleteScene"
        StopDragonWings
    DOF 128, DOFOff
    End If
    If bWallMBActive Then
        WallMBCompleted = WallMBCompleted + 1
        tmrMultiballCompleteScene.UserValue = "DMDWallMBCompleteScene"
        'TODO: Find out whether Wall JP Value resets after Wall Multiball
        If SelectedHouse = Baratheon Then WallJPValue = 9250000 Else WallJPValue = 3400000
    DOF 132, DOFOff
    End If
    If bWHCMBActive Then
        TimerFlags(tmrWinterHasCome) = 0
        StopSound "gotfx-ticktock"
        WHCFlipperFrozen = False
        tmrMultiballCompleteScene.UserValue = "DMDWHCMBCompleteScene"
        tmrWiCLightning.Enabled = False
        GiIntensity 1
        If WHCMBScore > WHCHighScore(CurrentPlayer) Then WHCHighScore(CurrentPlayer) = WHCMBScore
    DOF 129, DOFOff
    End If
    If bHotkMBActive Then
        tmrMultiballCompleteScene.UserValue = "DMDHotkMBCompleteScene"
        bInlanes(1) = 0
        If HotkScore > HotkHighScore(CurrentPlayer) Then HotkHighScore(CurrentPlayer) = HotkScore
        If House(CurrentPlayer).HasAbility(Tyrell) Then bInlanes(0) = True Else bInlanes(0) = False
    DOF 130, DOFOff
    End If
    If bITMBActive And bITMBDone Then
        ' Down to single ball and Iron Throne 7 houses have been completed. IT mode is over
        tmrMultiballCompleteScene.UserValue = "DMDIronThroneMBCompleteScene"
        If ITScore > ITHighScore(CurrentPlayer) Then ITHighScore(CurrentPlayer) = ITScore
        bITMBActive = False
        House(CurrentPlayer).ResetAfterIronThrone
    DOF 131, DOFOff
    End If
    If bMadnessMB = 2 Then
        bMadnessMB = -1
        tmrMultiballCompleteScene.UserValue = "DMDMadnessMBCompleteScene"
        tmrMadnessMBToggle.Enabled = 0
        GiIntensity 1
        GiOn
    End If
    bHotkMBActive = False
    bWallMBActive = False
    bWHCMBActive = False
    bBWMultiballActive = False
    bCastleMBActive = False
    bAddABallUsed = False
    SwordWall_Timer ' Ensure the lock walls are set correctly
    If PlayerMode = 0 Then TimerFlags(tmrUpdateBattleMode) = 0 : House(CurrentPlayer).CheckBattleReady
    PlayModeSong
    SetTopGates
    SetPlayfieldLights
    If PlayerMode = 1 Then
        DMDCreateAlternateScoreScene HouseBattle1,HouseBattle2
    ElseIf bITMBActive = False Then
        DMDResetScoreScene
        CheckForHotkOrITReady
    End If
    House(CurrentPlayer).SetUPFState False
End Sub

Sub RotateLaneLights(dir)
    If bTopLanes(0) or bTopLanes(1) Then
        bTopLanes(0) = Not bTopLanes(0)
        bTopLanes(1) = Not bTopLanes(1)
        SetTopLaneLights
    End If
    If bInlanes(0) <> bInlanes(1) Then
        bInlanes(0) = not bInlanes(0)
        bInlanes(1) = not bInlanes(1)
        SetInlaneLights
    End If
End Sub

Sub OpenTopGates: topgatel.open = True: topgater.open = True: End Sub
Sub CloseTopGates: SetTopGates : End Sub

Sub SetTopGates
    Dim lstate,rstate
    lstate=False:rstate=False
    If PlayerMode = 2 Or HouseBattle1 = Baratheon or HouseBattle2 = Baratheon or HouseBattle1 = Martell or HouseBattle2 = Martell or HouseBattle1=Greyjoy or HouseBattle2 = Greyjoy Then
        lstate=True:rstate=True
    End If
    If EBisLit > 0 or bElevatorShotUsed = False or bCastleMBActive or _
        (bWallMBReady And bMultiBallMode = False And bITMBActive = False) Or _
        (bMysteryLit And bMultiBallMode = False And PlayerMode = 0 And bITMBActive = False) Then
            lstate=True : MoveDiverter(1)
    Else MoveDiverter(0)
    End If
    If ComboMultiplier(1) > 1 Then rstate=True
    topgater.open = rstate
    topgatel.open = lstate
End Sub

Sub MoveDiverter(o)
    if o then
        Diverter.ObjRotZ = 22
        Diverter.collidable = 0
    Else
        Diverter.ObjRotZ = 0
        Diverter.collidable = 1
    End If
End Sub

Sub ResetComboMultipliers
    Dim i
    For i = 0 to 5: ComboMultiplier(i) = 1: Next
    SetComboLights
    SetLockbarLight
    SetTopGates
    DMDLocalScore
End Sub

Sub SetMystery
    If bMysteryLit = True And PlayerMode = 0 And bMultiBallMode = False And bITMBActive = False And bWallMBReady = False Then
        SetLightColor li153, white, 2  ' Turn on Mystery light
        MoveDiverter 1
        topgatel.open = True
    End If
End Sub

'Sub GameGiOn
'    Fi001.Visible = 1
'    Fi002.Visible = 1
'    Fi003.Visible = 1
'    Fi004.Visible = 1
'    Fi005.Visible = 1
'    Fi006.Visible = 1
'End Sub

'Sub GameGiOff
'    Fi001.Visible = 0
'    Fi002.Visible = 0
'    Fi003.Visible = 0
'    Fi004.Visible = 0
'    Fi005.Visible = 0
'    Fi006.Visible = 0
'End Sub

' This is called by LampTimer_Timer, which is in the generic code section
' This sub is table-specific and updates the lamps specific to this table
Sub UpdateLamps
    'SingleLampM fl237a, 0
  'SingleLamp fl237, 0
'    SingleLampM fl242a, 1
' SingleLamp fl242, 1
End Sub


Sub SetPFMLights
    ' Update the playfield multiplier lights
    Dim i
    For i = 2 to 5
        If PlayfieldMultiplierVal < i Then
            SetLightColor pfmuxlights(i-2),amber,0
        Elseif PlayfieldMultiplierVal = i Then
            pfmuxlights(i-2).BlinkInterval = 100
            SetLightColor pfmuxlights(i-2),amber,2
        Else
            SetLightColor pfmuxlights(i-2),amber,1
        End If
    Next
End Sub

' Set an effect on the Upper Playfield flashers
' fx : one of a number of flasher effects
'    0 = off, 1=on, 2=flash indefinitely, 3=flash 5 times fast, 4=flash 10 times fast
' col : colour const to use for the colour
Dim UPFFlasherState(2,2)
Sub SetUPFFlashers(fx,col)
    UPFFlasherState(2,0) = li222.state
    UPFFlasherState(2,1) = li222.colorfull
    SetFlasherWithColor li222,fx,col
    ' dim times,interval,i,fl
    ' Select Case fx
    '     Case 0: UPFFlasher001.visible = 0 : UPFFlasher002.visible = 0 : Exit Sub
    '     Case 1: times = 0
    '     Case 2: times = -1 : interval = 100 ' blink indefinitely
    '     Case 3: times = 5 : interval = 50 ' flash 5 times rapidly (half second flash)
    '     Case 4: times = 10 : interval= 50 ' flash 10 times rapidly (1 second)
    '     Case 5: times = 12 : interval= 88 ' flash 15 times rapidly (1.5 seconds) (Super Jackpot)
    '     Case 6: times = 1 : interval = 25 ' Flash once, used during Midnight Madness
    ' End Select
    ' i = 0
    ' For each fl in Array(UPFFlasher001,UPFFlasher002)
    '     ' Save current state
    '     If fl.TimerEnabled = 0 Then UPFFlasherState(i,0) = fl.visible : UPFFlasherState(i,1) = fl.color
    '     ' Turn flasher on and set colour
    '     SetFlashColor fl,col,1
    '     If times = 0 Then
    '         fl.TimerEnabled = 0
    '     Else
    '         ' Set up a timer on the flasher with a subroutine that we define. Sub
    '         fl.TimerInterval = interval
    '         fl.UserValue = times * 2 - 1
    '         fl.TimerEnabled = 0
    '         fl.TimerEnabled = 1
    '     End If
    '     i = i + 1
    ' Next
End Sub

Sub UPFFlasher001_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.Visible = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.TimerEnabled=0
        me.visible = UPFFlasherState(0,0)
        me.color=UPFFlasherState(0,1)
    End if
End Sub

Sub UPFFlasher002_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.Visible = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.Visible = 1
        Me.TimerEnabled=0
        me.visible = UPFFlasherState(1,0)
        me.color=UPFFlasherState(1,1)
    End if
End Sub

' Known lockbar light states:
' During regular play
'  - Lit solid for Baratheon, Targaryen if not yet used this ball
'  - Lit solid for Tyrell if there's at least one shot lit for Combos, or a playfield multiplier
'  - Lit solid for Lannister if sufficient gold
' During multiball
'  - Also lit solid for Martell if not yet used
' During Battle
'  - Also lit solid for Stark
' During Choose House
'  - Flashes colour of house
' During Mystery
'  - Flashes white or yellow?
' During Battle Selection
'  - Flashes red if one house is chosen; flashes between red and colour of second house if stacked
' After first ball, when ball is in the plunger lane
'  - Flashes white

Sub SetLockbarLight
    Dim st,col,ab,b
    st = 0 : col = 0 : ab=0 ' Defaults

    ' Calculate the color and state of the action button
    if bGameInPLay Then
        If PlayerMode = -1 Then
            If SelectedHouse = 8 Then 'Random
                li230States(0) = 7
                For b = 1 to 7 : li230States(b) = b : Next
                st = 3
            Else
                st = 2 : col = SelectedHouse
            End If
        ElseIf bBallInPlungerLane Then
            st = 2 : col = stark
        Elseif PlayerMode = 2 Then st = 0
        ElseIf bMysteryAwardActive Then
            st = 2 : col = baratheon
        Elseif PlayerMode = -2 Then
            st = 2 : col = lannister
            If HouseBattle2 > 0 Then
                st = 3
                li230States(0) = 4
                li230States(1) = 0
                li230States(2) = Lannister
                li230States(3) = 0
                li230States(4) = HouseBattle2
            End If
        Else
            col = House(CurrentPlayer).HasActionAbility
            If col <> 0 Then st = 1 Else st = 0
        End if
    Else
        If (DMDStd(kDMDStd_FreePlay) Or Credits > 0) Then st = 2 : col = Baratheon
    End if

    ' Set the light state
    If st < 2 Then
        For b = 1 to 7
            If b = col Then DOF 140+b,DOFon Else DOF 140+b,DOFoff
        Next
    End if
    if st < 2 Then
        li230.TimerEnabled = False
        'if Cabinetmode And Not bHaveLockbarButton Then
    if Cabinetmode And RenderingMode <> 2 And Not bHaveLockbarButton Then
            SetLightColor li231,HouseColor(col),st
        Else
            SetLightColor li230,HouseColor(col),st
        End If
    ElseIf st = 2 Then
        li230States(0) = 2 : li230States(1) = 0 : li230States(2) = col
    End If

    If st > 1 Then li230.TimerInterval = 100 : li230.TimerEnabled = True
End Sub

Dim li230States(10)
Sub li230_Timer
    Dim tmp,st,i
    me.TimerEnabled = False
    tmp = me.UserValue
    tmp = tmp+1
    if tmp > li230States(0) Then tmp = 1
    For i = 141 to 147 : DOF i,DOFOff : Next
    if li230States(tmp) > 0 then
        'if Cabinetmode And Not bHaveLockbarButton Then
    if Cabinetmode And RenderingMode <> 2 And Not bHaveLockbarButton Then
            SetLightColor li231,HouseColor(li230States(tmp)),1
        Else
            SetLightColor li230,HouseColor(li230States(tmp)),1
        End If
        DOF 140+li230States(tmp),DOFOn
    Else
        SetLightColor li230,white,0
        SetLightColor li231,white,0
    End if
    me.UserValue = tmp
    me.TimerInterval = 100
    me.TimerEnabled = True
End Sub


' Set an effect on the Playfield spotlight flashers
' fx : one of a number of flasher effects
'    0 = off, 1=on, 2=flash indefinitely, 3=flash 5 times fast, 4=flash 10 times fast
' col : colour const to use for the colour
Sub SetSpotFlashers(fx)
    SetFlasher fl244,fx
End Sub

' Set an effect on a specified flasher
' fx : one of a number of flasher effects
'    0 = off, 1=on, 2=flash indefinitely, 3=flash 5 times fast, 4=flash 10 times fast
' col : colour const to use for the colour
Sub SetFlasher(flasher,fx)
    SetFlasherWithColor flasher,fx,white
End Sub

Sub SetFlasherWithColor(flasher,fx,color)
    dim times,interval
    Select Case fx
        Case 0: flasher.TimerEnabled=False : flasher.state=0 : Exit Sub
        Case 1: times = 0
        Case 2: times = -1 : interval = 100 ' blink indefinitely
        Case 3: times = 5 : interval = 50 ' flash 5 times rapidly (half second flash)
        Case 4: times = 10 : interval= 50 ' flash 10 times rapidly (1 second)
        Case 5: times = 5 : interval = 25 ' really fast flash 5 times
        Case 6: times = 1 : interval = 25 ' Flash once, used during Midnight Madness
        Case 7: times = 2 : interval = 50 ' Flash twice, used during Attract Mode sequence
    End Select
    if (flasher.name = "li222") then debug.print "setting li222. Color: "&color
  if fx = 5 Then

  end if
    SetLightColor flasher,color,1
    If times = 0 Then
        flasher.TimerEnabled = 0
    Else
        flasher.TimerInterval = interval
        flasher.UserValue = times * 2 - 1
        flasher.TimerEnabled = 0
        flasher.TimerEnabled = 1
    End If
End Sub

' Timer subs for the flasher lamps. Flashers are now done in Blender, so the lamps
' represent their current state and (potentially) colour.
' We use the lamp's timer and UserValue to track the desired effect

' li222 - Upper PF floods
' fl235 - Sword flasher
' fl236 - right ramp flasher
' fl237 - bumper flasher
' fl242 - spinner flasher
' fl243 - Ram flasher
' fl244 - Sling spotlights
' fl245 - Throne flasher, backwall

' Upper PF flasher is special, as it restores state once done flashing
Sub li222_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.State = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.TimerEnabled=0
        me.state = UPFFlasherState(2,0)
        me.colorfull=UPFFlasherState(2,1)
    End if
End Sub

Sub fl235_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.State = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.State = 0
        Me.TimerEnabled=0
    End if
End Sub

Sub fl236_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.State = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.State = 0
        Me.TimerEnabled=0
    End if
End Sub

Sub fl237_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.State = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.State = 0
        Me.TimerEnabled=0
    End if
End Sub

Sub fl242_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.State = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.State = 0
        Me.TimerEnabled=0
    End if
End Sub

Sub fl243_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.State = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.State = 0
        Me.TimerEnabled=0
    End if
End Sub

Sub fl244_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.State = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.State = 0
        Me.TimerEnabled=0
        'GiOn
    End if
End Sub

Sub fl245_Timer
    Dim tmp
    tmp=me.UserValue
    tmp=tmp-1
    Me.State = tmp MOD 2
    me.UserValue = tmp
    If tmp = 0 Then
        Me.State = 0
        Me.TimerEnabled=0
    End if
End Sub

' During Battle mode, Shield lights may be in one of several states
' They may also alternate colour. To deal with this, create an array of
' light states and set a timer on each light to cycle through its states
' 1st element of array is number of states for this light
' This Sub sets the jackpot light. The SetModeLights in the BattleState class
' handles all of the battle-related colours
Dim ModeLightState(10,10)
Sub SetModeLights
    Dim i

    House(CurrentPlayer).SetShieldLights

    ' Set up timers on the Shield lights that have more than one state defined
    ' Turn off the lights of those that don't
    For i = 1 to 7
        If ModeLightState(i,0) < 2 Then
            HouseShield(i).TimerEnabled = False
            HouseShield(i).State = 0
        Else
            HouseShield(i).TimerEnabled = True
        End If
        HouseShield(i).TimerInterval = 100
        HouseShield(i).UserValue = 1
    Next
End Sub

' House Shield light timers. Used to cycle through color states during battle mode
'HouseShield = Array(li141,li141,li26,li114,li86,li77,li156,li98)
Sub li141_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(1,uv) > 0 Then SetLightColor Me,ModeLightState(1,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(1,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

Sub li26_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(2,uv) > 0 Then SetLightColor Me,ModeLightState(2,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(2,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

Sub li114_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(3,uv) > 0 Then SetLightColor Me,ModeLightState(3,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(3,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

Sub li86_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(4,uv) > 0 Then SetLightColor Me,ModeLightState(4,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(4,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

Sub li77_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(5,uv) > 0 Then SetLightColor Me,ModeLightState(5,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(5,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

Sub li156_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(6,uv) > 0 Then SetLightColor Me,ModeLightState(6,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(6,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

Sub li98_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(7,uv) > 0 Then SetLightColor Me,ModeLightState(7,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(7,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

Sub li189_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(8,uv) > 0 Then SetLightColor Me,ModeLightState(8,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(8,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

Sub li192_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(9,uv) > 0 Then SetLightColor Me,ModeLightState(9,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(9,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

Sub li195_Timer
    Dim uv
    uv = Me.UserValue
    Me.TimerEnabled = False
    If Tilted Then exit Sub
    If ModeLightState(10,uv) > 0 Then SetLightColor Me,ModeLightState(10,uv),1 Else Me.state=0
    uv = uv + 1
    If uv > ModeLightState(10,0) Then uv = 1
    Me.UserValue = uv
    Me.TimerEnabled = True
End Sub

'*******************************************************
' Combo multiplier light timer
' Used to speed up the flash rate as the timer runs down
'*******************************************************
'ComboLights = Array(li89,li89,li101,li117,li144,li159)
Sub li89_Timer
    Dim i,bi
    For i = 1 to 5
        If ComboMultiplier(i) > 1 And (i <> 5 Or bWallMBReady = False) And (i<>3 Or (bHotkMBReady = False And bITMBReady = False)) Then
            bi = TimerTimestamp(tmrComboMultplier) - GameTimeStamp
            if bi > 50 then bi = 150 else bi = bi*2 + 50
            ComboLights(i).State = 0
            ComboLights(i).BlinkInterval = bi
            ComboLights(i).State = 2 ' Hopefully this ensure they all blink in unison.
        End if
        If i = 3 And (bHotkMBReady Or bITMBReady) Then ComboLights(i).State = 2
    Next
End Sub

Sub CheckActionButton
    If PlayerMode = -2 Then
    if bReadyForAction then
      PreLaunchBattleMode
    Else
      Debug.print "Ignoring action button until ready"
    end If
    ElseIf PlayerMode = -2.1 Then
        ' Skip battle intro animation and get to it
        DMDClearQueue
        LaunchBattleMode
    ElseIf PlayerMode >= 0 Then
        House(CurrentPlayer).CheckActionBtn
    End If
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly
' There is no bonus multiplier lights in this table

Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level
End Sub

Sub IncreaseBonusMultiplier(bx)
    If BonusMultiplier(CurrentPlayer) = MaxBonusMultiplier then Exit Sub
    BonusMultiplier(CurrentPlayer) = BonusMultiplier(CurrentPlayer) + bx
    If BonusMultiplier(CurrentPlayer) > MaxBonusMultiplier Then BonusMultiplier(CurrentPlayer) = MaxBonusMultiplier
    If bMadnessMB = 2 Then Exit Sub
    Dim scene
    If bUseFlexDMD Then
        Set scene = FlexDMD.NewGroup("testscene")
        scene.AddActor FlexDMD.NewLabel("lbl1",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbBlack, 1),BonusMultiplier(CurrentPlayer)&"X")
        scene.GetLabel("lbl1").SetAlignedPosition 64,16,FlexDMD_Align_CENTER
        BlinkActor scene.GetLabel("lbl1"),50,30
        scene.AddActor NewSceneFromImageSequence("img1","bonusx",50,20,2,0)
        DMDEnqueueScene scene,1,3000,4500,4000,"gotfx-wind-blowing"
    Else
        DisplayDMDText "",BonusMultiplier(CurrentPlayer)&"X BONUS",2000
    End If
End Sub

' ComboMultiplier increaser
' The logic is rather complicated. In general, a hit on one shot will increase the
' multiplier on the other shots to this shot's multiplier + 1. Ramps affect all other
' shots. L orbit affects only the next 2 adjacent shots. R orbit affects 3 middle shots.
' Dragon affects ALL shots (including itself) but the timer is half the length.
' In addition, it looks like a hit on a shot that's at 1x but is *different* from the
' the last shot will increase the multipliers on the other shots rather than just giving you
' one more than 2x
Dim LastComboHit
Sub IncreaseComboMultiplier(h)
    Dim i,c,x,tmr,mask,max
    tmr = 150           ' Default combo rundown timer
    c = ComboLaneMap(h)
    If c = 0 And h <> 0 Then Exit Sub  ' Target bank was hit - they don't increase multipliers
    If bHotkMBActive Then max = 6 Else max = 5 ' Max is 6X during HOTK
    If max > 3+SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier) Then max = 3+SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier)
    If c = 0 Then
        For i = 1 to 5
            If ComboMultiplier(i) > x Then x = ComboMultiplier(i)
        Next
    Else
        x = ComboMultiplier(c)
    End If
    x = x + 1
    if x > max Then x = max
    Select Case c
        Case 0          ' Used for Inlane hit increases
            mask = 62   ' Increase x of all shots
        Case 1          ' Left orbit
            mask = 12   ' Turn on shots 2 & 3
        Case 2          ' Dragon shot
            mask = 62
            tmr = 80
        Case 3          ' L ramp
            mask = 54
        Case 4          ' R ramp
            mask = 46
        Case 5          ' Right orbit
            mask = 28
    End Select
    For i = 1 to 5
        If (mask And 2^i) > 0 Then
            If c = LastComboHit or x > ComboMultiplier(i) Then
                ComboMultiplier(i) = x
            ElseIf ComboMultiplier(i) < max Then
                ComboMultiplier(i) = ComboMultiplier(i) + 1
            End If
        Else
            ComboMultiplier(i) = 1
        End If
    Next
    If c <> 0 Then LastComboHit = c
    SetGameTimer tmrComboMultplier,tmr
    SetComboLights
    SetLockbarLight
    SetTopGates
    DMDLocalScore 'Update the DMD. TODO: On the real game, the DMD flashes the multipliers when they first change
End Sub

Sub AddGold(g)
    Dim scene
    TotalGold = TotalGold + g
    CurrentGold = CurrentGold + g
    If bUseFlexDMD And PlayerMode <> 2 And bMadnessMB < 1 Then
        Set scene = NewSceneWithImage("goldstack","goldstack")
        scene.AddActor FlexDMD.NewLabel("addgold",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbBlack, 0),"+"&g&" GOLD")
        scene.GetLabel("addgold").SetAlignedPosition 2,0,FlexDMD_Align_TopLeft
        BlinkActor scene.GetLabel("addgold"),150,4
        scene.AddActor FlexDMD.NewLabel("gold",FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbBlack, 0),"TOTAL GOLD: "&CurrentGold)
        scene.GetLabel("gold").SetAlignedPosition 2,27,FlexDMD_Align_BottomLeft
        DMDEnqueueScene scene,2,1200,2000,2000,""
    End If
    SetLockbarLight
End Sub

Sub AddBonus(b)
    BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + b
End Sub

' Check for key presses specific to this game.
' If PlayerMode is < 0, in a 'Select' state, so use flippers to toggle
Function CheckLocalKeydown(ByVal keycode)
    CheckLocalKeydown = False
    if (keycode = LeftFlipperKey or keycode = RightFlipperKey) Then
        If PlayerMode < 0 Then CheckLocalKeydown = True
        if PlayerMode = -1 Then
            ChooseHouse(keycode)
        ElseIf PlayerMode = -2 And bReadyForFlips Then
            ChooseBattle(keycode)
        End If
        ' Check for both flippers pushed
        If ((LFPress=1 and keycode = RightFlipperKey) Or (RFPress=1 And keycode = LeftFlipperKey)) Then
            If PlayerMode = -2.1 Then
                tmrStartBattleModeScene_Timer
                tmrStartBattleModeScene_Timer ' called twice, once for each potentially stacked battle. The second call launches battle mode
            ElseIf bWallMBActive And tmrStartWallMB.Enabled <> 0 Then
                ' Skip Wall MB Intro Scene
                tmrStartWallMB_Timer
            ElseIf bBWMultiballActive And tmrBWmultiballRelease.Interval=5000 And tmrBWmultiballRelease.Enabled <> 0 Then
                ' Skip BWMB Intro
                tmrBWmultiballRelease_Timer
            End If
        End If
    End If
End Function


Sub ChooseHouse(ByVal keycode)
    Dim t ' for debugging map mover
    If keycode = LeftFlipperKey Then
        FlashShields SelectedHouse,False
        'House(CurrentPlayer).StopSay(SelectedHouse)
        t=SelectedHouse
        If SelectedHouse = Stark Then
            SelectedHouse = Targaryen
        ElseIf SelectedHouse = targaryen Then
            SelectedHouse = 8
        ElseIf SelectedHouse = 8 Then
            SelectedHouse = Martell
        Else
            SelectedHouse = SelectedHouse - 1
        End If
        FlashShields SelectedHouse,True
        House(CurrentPlayer).Say(SelectedHouse)
        'DMDScrollITMap t,SelectedHouse,10000000
    ElseIf keycode = RightFlipperKey Then
        FlashShields SelectedHouse,False
        'House(CurrentPlayer).StopSay(SelectedHouse)
        If SelectedHouse = Martell Then
            SelectedHouse = 8
        ElseIf SelectedHouse = 8 Then
            SelectedHouse = Targaryen
        ElseIf SelectedHouse = Targaryen Then
            SelectedHouse = Stark
        Else
            SelectedHouse = SelectedHouse + 1
        End If
        FlashShields SelectedHouse,True
        House(CurrentPlayer).Say(SelectedHouse)
    End If
    SetLockbarLight
    DMDChooseScene1 "CHOOSE YOUR HOUSE",HouseToUCString(SelectedHouse), HouseAbility(SelectedHouse),"got-choose" & HouseToString(SelectedHouse)
End Sub

' Turn on the flashing shield sigils when choosing a house
Sub FlashShields(h,State)
    Dim i,j,k
    If h = 8 Then j=1:k=7 Else j=h:k=h
    For i = j to k
        if State Then
            SetLightColor HouseSigil(i),HouseColor(i),2
            SetLightColor HouseShield(i),HouseColor(i),2
        Else
            HouseSigil(i).State = 0
            HouseShield(i).State = 0
        End If
    Next
End Sub

' Handle game-specific processing when ball is launched
Sub GameDoBallLaunched
    If PlayerMode = -1 Then
        If SelectedHouse = 8 Then SelectedHouse = RndNbr(7)
        House(CurrentPlayer).MyHouse = SelectedHouse
        House(CurrentPlayer).ResetLights
        'PlayerMode = 0
    SetPlayerMode 0
        DMDFlush
        ' TODO: Display additional text about house chosen on ball launch
        DMDLocalScore
        SetLockbarLight
    End If
    If bMultiBallMode = False Then  PlaySoundVol "gotfx-balllaunch",VolDef
End Sub


Sub ChooseBattle(ByVal keycode)
    If keycode = LeftFlipperKey or keycode = RightFlipperKey Then
        If bBattleInstructionsDone = False Then tmrBattleInstructions_Timer : Exit Sub
        TimerTimestamp(tmrChooseBattle) = TimerTimestamp(tmrChooseBattle) + 80
        If TimerTimestamp(tmrChooseBattle) > GameTimeStamp+200 Then TimerTimestamp(tmrChooseBattle) =  GameTimeStamp+200
        If keycode = LeftFlipperKey Then
            PlaySoundVol "gotfx-choosebattle-left",VolDef
            CurrentBattleChoice = CurrentBattleChoice - 1
            if CurrentBattleChoice < 0 Then CurrentBattleChoice = TotalBattleChoices - 1
        Else
            PlaySoundVol "gotfx-choosebattle-right",VolDef
            CurrentBattleChoice = CurrentBattleChoice + 1
            if CurrentBattleChoice >= TotalBattleChoices Then CurrentBattleChoice = 0
        End If
        UpdateChooseBattle
        ResetBallSearch
    End If
End Sub

'*****************************
'  Handle target hits
'*****************************

' LoL Drop Targets
' Any target increases spinner value.
Sub Target9_Hit 'LoL target 1
    If Tilted Or BallSearchCnt > 0 Then Exit Sub
    DTHit 3
End Sub

Sub Target8_Hit 'LoL target 2
    If Tilted Or BallSearchCnt > 0 Then Exit Sub
    DTHit 2
End Sub

Sub Target7_Hit 'LoL target 3
    If Tilted Or BallSearchCnt > 0 Then Exit Sub
    DTHit 1
End Sub

Dim TargetState(4)
Sub DTAction(switchid)
    TargetState(switchid) = 1
  If BallSearchCnt = 0 And switchid <> 4 Then DoTargetsDropped
End Sub

Sub DoTargetsDropped
    Dim i,tmp
    Addscore 330
    tmp = False
    if bMadnessMB = 2 Then DoMadnessMBHit

    ' In case two targets were hit at once, stop the sound for the first target before playing the one for the second
    StopSound "gotfx-loltarget-hit" & DroppedTargets
    DroppedTargets = DroppedTargets + 1
    If PlayerMode = 1 Then
        If HouseBattle2 > 0 Then vpmTimer.addTimer 10,"House("&CurrentPlayer&").BattleState("&HouseBattle2&").RegisterTargetHit 0 '"
        If HouseBattle1 > 0 Then vpmTimer.addTimer 10,"House("&CurrentPlayer&").BattleState("&HouseBattle1&").RegisterTargetHit 0 '"
    ElseIf PlayerMode = 2 Or bWHCMBActive Or bHotkMBActive or bITMBActive Then House(CurrentPlayer).RegisterHit(Baratheon) : tmp = True
    ElseIf PlayerMode = 0 Then ' Only increase Spinner Value and play target dropped sound in regular play mode
        PlaySoundVol "gotfx-loltarget-hit" & DroppedTargets, VolDef
        ' According to Chukwurt, Max spinner value per level are roughly 30k,60k,150k,250k,400k,500k,650k,800k
        ' Also, Level only jumps up every 3 bank completions, so we divide the spinner level by 3. To get the above
        ' numbers, we can approximate by squaring the spinner level and multiplying by 10k
        SpinnerValue = SpinnerValue + (SpinnerAddValue * RndNbr(4) * Int((SpinnerLevel+2)/3) * Int((SpinnerLevel+2)/3))
        If SpinnerValue > (10000 * Int((SpinnerLevel+2)/3+1) * Int((SpinnerLevel+2)/3+1)) Then SpinnerValue = (10000 * Int((SpinnerLevel+2)/3+1) * Int((SpinnerLevel+2)/3+1))
    End If
    If DroppedTargets = 3 Or (TargetState(1)=1 AND  TargetState(2)=1 AND  TargetState(3)=1) Then
        ' Target bank completed
        AddBonus 100000
        If PlayerMode = 0 And bMultiBallMode = False Then LoLTargetsCompleted = LoLTargetsCompleted + 1
        vpmTimer.addTimer 100, "ResetDropTargets '"
        If bLoLLit = False and bLoLUsed = False Then DoLordOfLight False
        For i = 0 to 2
            'TODO: Revisit this to see whether LoL lights that are on solid still flash when bank is completed
            FlashForMs LoLLights(i),500,100,2
        Next
        If SpinnerLevel <= CompletedHouses Then SpinnerLevel = SpinnerLevel + 1 'TODO: Should SpinnerLevel still increase if in a Mode?
        If tmp = False Then House(CurrentPlayer).RegisterHit(Baratheon)
        If House(CurrentPlayer).HasAbility(Baratheon) And bMultiBallMode = False Then AdvanceWallMultiball 1
    End If
End Sub

Sub ResetDropTargets
    ' PlaySoundAt "fx_resetdrop", Target010
    Dim t,b
    If TargetState(1)=1 OR  TargetState(2)=1 OR  TargetState(3)=1 Then
        DTRaise 1 : TargetState(1) = 0
        DTRaise 2 : TargetState(2) = 0
        DTRaise 3 : TargetState(3) = 0
    End If
    DroppedTargets = 0
    SetTargetLights
    DOF 119,DOFPulse
End Sub

Sub Target7_Timer: Me.TimerEnabled = 0 : ResetDropTargets : End Sub

'*********************
' Wildfire target hits
'*********************
Sub Target43_Hit
  STHit 9
'    doWFTargetHit 0
End Sub

Sub Target44_Hit
  STHit 10
'    doWFTargetHit 1
End Sub

Sub doWFTargetHit(t)
    If Tilted then Exit Sub
    Dim t1:t1=1
    If t Then t1=0
    AddScore 230

    if bMadnessMB = 2 Then DoMadnessMBHit

    If bITMBReady or bHotkMBReady Then Exit Sub
    If bWildFireTargets(t) Then PlaySoundVol "gotfx-wftarget-lit",VolDef : Else PlaySoundVol "gotfx-wftarget-unlit",VolDef

    If PlayerMode = 2 Or bWHCMBActive Or bHotkMBActive or bITMBActive Then House(CurrentPlayer).RegisterHit(Tyrell) : Exit Sub
    If ((BWMultiballsCompleted+DMDStd(kDMDFet_TwoBankDifficulty)) < 2 or bWildfireTargets(t1)) And bMultiBallMode = False Then LightLock
    if bWildfireTargets(t) then Exit Sub
    bWildfireTargets(t) = True
    If bBWMultiballActive = False Then House(CurrentPlayer).AddWildfire 1
    If PlayerMode = 1 Then
        If HouseBattle2 > 0 Then House(CurrentPlayer).BattleState(HouseBattle2).RegisterTargetHit 1
        If HouseBattle1 > 0 Then House(CurrentPlayer).BattleState(HouseBattle1).RegisterTargetHit 1
    End If

    if bWildfireTargets(t1) Then
        'Target bank completed
        AddBonus 100000
        'Light both lights for 1 second, then shut them off
        FlashForMs li80,1000,1000,0
        FlashForMs li83,1000,1000,0
        ' Lights don't always seem to restore their state properly after flashing, so stick a timer on it
        li80.TimerInterval = 1100
        li80.TimerEnabled = True
        bWildfireTargets(0)=False:bWildfireTargets(1)=False
        If PlayerMode <> 1 Then House(CurrentPlayer).RegisterHit(Tyrell)
        DoWildfireLit
        Target43.TimerEnabled = 0
    Else
        AddBonus 10000
        Select Case t
            Case 0
                SetLightColor li80,green,1
                FlashForMs li80,1000,100,2
            Case 1
                SetLightColor li83,green,1
                FlashForMs li83,1000,100,2
        End Select
        If (BWMultiballsCompleted+DMDStd(kDMDFet_TwoBankDifficulty)) > 2 Then Target43.TimerInterval = 30000 : Target43.TimerEnabled = 1
    End If
    LastSwitchHit = "wftarget"&t
End Sub

' This sub fires if the Wildfire targets time-out before completion
Sub Target43_Timer
    If (BWMultiballsCompleted+DMDStd(kDMDFet_TwoBankDifficulty)) > 2 Then bWildFireTargets(0) = False : bWildfireTargets(1) = False : SetWildfireLights
    Me.TimerEnabled = 0
End Sub

'WF target light timer
Sub li80_Timer
    SetWildfireLights
    Me.TimerEnabled = False
End Sub

'Gold target light timer
Sub li92_Timer
    Me.TimerEnabled = False
    SetGoldTargetLights
End Sub

Sub LightLock
    Dim i

    if bLockIsLit or bMultiBallMode Or bITMBActive Then Exit Sub
    bLockIsLit = True
    ' Flash the lock light
    li111.BlinkInterval = 300
    SetLightColor li111,darkgreen,2

    ' Enable the lock wall
    LockWall.collidable = 1 : Lockwall.Uservalue = 1 : LockPost True
    If RealBallsInLock > 0 Then SwordWall.collidable = 1

    i = RndNbr(3)
    if i > 1 Then i = ""
    PlaySoundVol "say-lock-is-lit"&i, VolCallout

    ' Ensure Battle is enabled for the start of multiball, as long as at least one house is qualified
    If BallsInLock = 2 And PlayerMode <> 1 Then
        For i = Stark to Targaryen
            If House(CurrentPlayer).Qualified(i) and House(CurrentPlayer).Completed(i) = False Then
                House(CurrentPlayer).BattleReady = True
                SetLightColor li108,white,2
                Exit For
            End If
        Next
    End If
End Sub

'******************
' Top Lane switches
'******************

Sub sw53_Hit
    UpdateWires 53,true
    TopLane_Hit 0
End Sub

Sub sw54_Hit
    UpdateWires 54,true
    TopLane_Hit 1
End Sub

Sub TopLane_Hit(sw)
    If Tilted Then Exit Sub
    if bMadnessMB = 2 Then DoMadnessMBHit
    if bTopLanes(sw) = False Then
        AddScore 1000
        AddBonus 5000
        bTopLanes(sw) = True
        PlaySoundVol "gotfx-toplane-unlit",VolDef
        SetTopLaneLights    ' sub takes care of flashing them if they're both lit
        If bTopLanes(0) And bTopLanes(1) Then
            PlaySoundVol "gotfx-toplanes-complete",VolDef
            IncreaseBonusMultiplier 1
            If House(CurrentPlayer).HasAbility(Lannister) Then AddGold 225 Else AddGold 135
            bTopLanes(0) = False: bToplanes(1) = False
        End if
    Else
        PlaySoundVol "gotfx-drum7",VolDef/4
        AddScore 110
    End If
    LastSwitchHit = "toplane"
End Sub

'*******************
' Outlane Switches
'*******************

Sub sw3_Hit
    UpdateWires 3,true
    Outlane_Hit
End Sub

Sub sw4_Hit
    UpdateWires 4,true
    Outlane_Hit
End Sub

Sub Outlane_Hit
    If Tilted then Exit Sub
    PlaySoundVol "gotfx-outlanelost",VolDef
    if bMadnessMB = 2 Then DoMadnessMBHit
    AddScore 25000
    AddBonus 25000
    LastSwitchHit = "OutlaneSW"
    If bMultiBallMode Then
        If bBallSaverActive Then Exit Sub
        If bLoLLit and bMadnessMB < 1 Then
            bLolUsed = True
            EnableBallSaver 5
            DoLordOfLight True
            bLoLLit = False
            SetOutlaneLights
        End If
        Exit Sub
    End If
    If bBallSaverActive Then
        bEarlyEject = True
        CreateMultiballTimer.Interval = 100
        DoBallSaved 0
    Elseif bLoLLit Then
        bEarlyEject = True
        CreateMultiballTimer.Interval = 100
        DoBallSaved 1
        bLoLUsed = True
        SetOutlaneLights
    End If
End Sub

'*****************
' Ramp entrance switches
'*****************

' Left ramp
Sub sw38_Hit
    If Tilted then Exit Sub
    If bHotkMBReady Then
        PlaySoundVol "gotfx-ramp-hotk-ready",VolDef
    ElseIf PlayerMode = 1 Then PlaySoundVol "gotfx-ramphit1",VolDef/4
    Else PlaySoundVol "gotfx-swordswoosh",VolDef
    End If
    LastSwitchHit = "sw38"
End Sub

' Right Ramp
Sub sw41_Hit
    If Tilted then Exit Sub
    If PlayerMode = 1 Then PlaySoundVol "gotfx-ramphit2",VolDef/4 Else PlaySoundVol "gotfx-swordswoosh",VolDef
    LastSwitchHit = "sw41"
End Sub

'******************
' 5 main shot hits
'******************
Sub LOrbitSW30_Hit
    If Tilted then Exit Sub
    AddScore 1000
    If LastSwitchHit = "sw30a" Then House(CurrentPlayer).RegisterHit(Greyjoy)
    'House(CurrentPlayer).RegisterHit(Greyjoy)
    LastSwitchHit = "LOrbitSW30"
End Sub

Sub sw30a_Hit
    LastSwitchHit = "sw30a"
End Sub

Sub sw30b_Hit
    LastSwitchHit = "sw30b"
End Sub
' Left ramp switch
Sub sw39_Hit
    If Tilted then Exit Sub
    AddScore 1000
    House(CurrentPlayer).RegisterHit(Lannister)

    LastSwitchHit = "sw39"
    sw48.UserValue = "sw39"
End Sub

'Right ramp switch
Sub sw42_Hit
    If Tilted then Exit Sub
    AddScore 1000
    House(CurrentPlayer).RegisterHit(Stark)
    LastSwitchHit = "sw42"
End Sub

Sub ROrbitsw31_Hit
    If Tilted then Exit Sub
    Me.UserValue = 0
    If LastSwitchHit <> "swPlungerRest" Then
        AddScore 1000
        If LastSwitchHit <> "sw30b" And bJustPlunged = False Then House(CurrentPlayer).RegisterHit(Martell)
        'House(CurrentPlayer).RegisterHit(Martell)
    End If
    LastSwitchHit = "ROrbitsw31"
End Sub

' Right ramp drop target
Sub Target90_Hit
    PlaySoundAt "Drop_Target_Down_" & RndNbr(6), Target90
    If Tilted Then Exit Sub
    DTHit 4
End Sub

'******************
' CastleWall Kicker
'******************

Sub Kicker37_Hit
    If Tilted then Exit Sub
    AddScore 1000
    House(CurrentPlayer).RegisterHit(Targaryen)
    PlaySoundAt SoundFXDOF("fx_kickback", 112, DOFPulse, DOFContactors),kicker37
    Kicker37.Kick 190,RndNum(55,60)   'Angle,Power
End Sub

'*****************
' Elevator Kickers
'*****************
' kickerfloor - bottom of elevator
' kickerUPF - UpperPF level
' kickertopfloor - Iron Throne level
' kickerIT - kicker in Iron Throne

Sub KickerFloor_Hit
    Me.DestroyBall
    SoundSaucerLock KickerFloor
    MoveDiverter(0)
    Dim delay : delay = 500
    If ROrbitsw31.UserValue > 10 Then delay = 3000 ' A qualifying hit on the Right orbit was made, so give time for callout to finish
    If ((((bMysteryLit And PlayerMode = 0) Or (bWallMBReady And PlayerMode < 2 )) And bITMBActive = False And bMultiBallMode = False) or EBisLit > 0) And bJustPlunged = False Then
        vpmTimer.AddTimer delay,"ElevatorKick 2 '"
        If bMultiBallMode = False Then FreezeAllGameTimers
    Else
        If bElevatorShotUsed = False Then bElevatorShotUsed = True : bCastleShotAvailable = True
        vpmTimer.AddTimer 500,"ElevatorKick 1 '"
    End If
    SetTopGates
End Sub

Sub ElevatorKick(f)
    'PlaySoundAt SoundFXDOF("fx_kicker", 110, DOFPulse, DOFContactors), KickerFloor
  dim newball
    SoundSaucerKick 1, KickerFloor
    DOF 110,DOFPulse
    Select Case f
        Case 1
            set newball = KickerUPF.CreateBall
            KickerUPF.Kick 180,5
            PlaySoundVol "gotfx-elevatorupf",VolDef
            SetUPFFlashers 3,red
        Case 2
            KickerTopFloor.CreateBall
            KickerTopFloor.Kick 90,3
    End Select
End Sub

Sub KickerUPF_Hit
    KickerUPF.Kick 180,5
End Sub

Sub KickerIT_Hit
    Dim delay
    RandomSoundMetal
    SoundSaucerLock KickerIT
    if bMadnessMB > 0 or Tilted Then
        IronThroneKickout
        Exit Sub
    End if
  DOF 125, DOFOn
    delay = 500
    If EBisLit > 0 Then
        DoAwardExtraBall
        EBisLit = 0 : setEBLight
        delay = 6500
    End If
    If bWallMBReady And PlayerMode < 2 And bMultiBallMode = False And bITMBActive = False Then
        vpmTimer.AddTimer delay,"StartWallMB '"
    ElseIf bMysteryLit And PlayerMode = 0 And bMultiBallMode = False Then
        vpmTimer.AddTimer delay,"DoMysteryAward '"
    Else
        vpmTimer.AddTimer delay,"IronThroneKickout '"
    End If
End Sub

Sub IronThroneKickout
    'PlaySoundAt SoundFXDOF("fx_kicker", 111, DOFPulse, DOFContactors), KickerIT
    SoundSaucerKick 1, KickerIT
    DOF 111, DOFPulse
  DOF 125, DOFOff
    KickerIT.Kick 180,3
End Sub

'******************
' lock switch
'******************
Sub sw48_Hit
    If Tilted then Exit Sub
    ' Debounce - ignore if the ramp switch wasn't just hit
    If sw48.UserValue <> "sw39" Then Exit Sub
    sw48.UserValue = ""

    If bHotkIntroRunning Then Exit Sub ' StartHOTKMB will take care of releasing the ball

    if bMultiBallMode Then
        tmrBWmultiballRelease.Interval = 1000
        MultiballBallsLocked = MultiballBallsLocked + 1
        If bBWMultiballActive Then
            BallsInLock = BallsInLock + 1
            RealBallsInLock = RealBallsInLock + 1
        End If
        If RealBallsInLock > 1 Or MultiballBallsLocked > 1 Then tmrBWmultiballRelease.Interval = 300
        bBallReleasing = True
        tmrBWmultiballRelease.Enabled = True
        Exit Sub
    Else
        MultiballBallsLocked = 0
    End If

    If PlayerMode < 0 Then Exit Sub     ' ChooseBattle mode already started - let it take care of doing ball lock when done
    If bLockIsLit And bHotkMBReady = False And bITMBReady = False Then
        bBallReleasing = True
        FreezeAllGameTimers
        vpmtimer.addtimer 400, "LockBall '"     ' Slight delay to give ball time to settle
    ElseIf (RealBallsInLock > 0 or Lockwall.UserValue = 1) Then     ' Lock isn't lit but we have a ball locked
    debug.print "Lock isn't lit but we have a ball locked"
        bBallReleasing = True
        If RealBallsInLock = 0 Then vpmTimer.addTimer 750, "ReleaseLockedBall 0 '" Else ReleaseLockedBall 0
    End If
    LastSwitchHit = "sw48"
End Sub


'*************************
' Upper Playfield Switches
'*************************

'Castle loop
Sub sw79_Hit
    UpdateWires 79,true
    If Tilted then Exit Sub
    House(CurrentPlayer).RegisterUPFHit 1
End Sub

' Left target
Sub Target80_Hit
    If Tilted then Exit Sub
  STHit 13
'    House(CurrentPlayer).RegisterUPFHit 2
End Sub

'Left outlane
Sub sw77_Hit
    UpdateWires 77,true
    If Tilted then Exit Sub
    House(CurrentPlayer).RegisterUPFHit 3
End Sub

'Center target
Sub Target81_Hit
    If Tilted then Exit Sub
  STHit 12
'    House(CurrentPlayer).RegisterUPFHit 4
End Sub

' Right outlane
Sub sw78_Hit
    UpdateWires 78,true
    If Tilted then Exit Sub
    House(CurrentPlayer).RegisterUPFHit 5
End Sub

' Right target
Sub Target82_Hit
    If Tilted then Exit Sub
  STHit 11
'    House(CurrentPlayer).RegisterUPFHit 6
End Sub

' Left inlane
Sub sw83_Hit
  activeball.angmomz = 0: activeball.vely = activeball.vely*0.8
    UpdateWires 83,true
    If Tilted then Exit Sub
    House(CurrentPlayer).RegisterUPFHit 7
End Sub

' Right inlane
Sub sw84_Hit
  activeball.angmomz = 0: activeball.vely = activeball.vely*0.8
    UpdateWires 84,true
    If Tilted then Exit Sub
    House(CurrentPlayer).RegisterUPFHit 8
End Sub

' Left sling
' TODO: Add a sound?
Sub rlbandsw85_Hit
    If Tilted then Exit Sub
    AddScore 110
End Sub

' Right sling
Sub rlbandsw86_Hit
    If Tilted then Exit Sub
    AddScore 110
End Sub


'*****************
' Gold targets hit
'*****************
Sub Target32_Hit
  STHit 4
'    GoldHit 0
End Sub
Sub Target33_Hit
  STHit 5
'    GoldHit 1
End Sub
Sub Target34_Hit
  STHit 6
'    GoldHit 2
End Sub
Sub Target35_Hit
  STHit 7
'    GoldHit 3
End Sub
Sub Target36_Hit
  STHit 8
'    GoldHit 4
End Sub

Sub GoldHit(n)
    If Tilted then Exit Sub
    if bMadnessMB = 2 Then DoMadnessMBHit
    House(CurrentPlayer).GoldHit(n)
    LastSwitchHit = "gold"&n
End Sub

' Left Inlane
Sub sw1_Hit
    activeball.angmomz = 0: activeball.vely = activeball.vely*0.8
    UpdateWires 1,true
    If Tilted then Exit Sub
    AddScore 560
    DoInlaneHit 0
    LastSwitchHit = "sw1"
End Sub

' Right Inlane
Sub sw2_Hit
    activeball.angmomz = 0: activeball.vely = activeball.vely*0.8
    UpdateWires 2,true
    If Tilted or Not bGameInPlay then Exit Sub
    AddScore 560
    DoInlaneHit 1
    LastSwitchHit = "sw2"
End Sub

Sub DoInlaneHit(t)
    PlaySoundVol "gotfx-drum7",VolDef/4
    if bMadnessMB = 2 Then DoMadnessMBHit : Exit Sub
    If bInlanes(t) Then
        PlaySoundVol "gotfx-inlanelit",VolDef
        IncreaseComboMultiplier 0
    End If
End Sub

Sub sw1_UnHit : UpdateWires 1,false : End Sub
Sub sw2_UnHit : UpdateWires 2,false : End Sub
Sub sw3_UnHit : UpdateWires 3,false : End Sub
Sub sw4_UnHit : UpdateWires 4,false : End Sub
Sub sw53_UnHit : UpdateWires 53,false : End Sub
Sub sw54_UnHit : UpdateWires 54,false : End Sub
Sub sw77_UnHit : UpdateWires 77,false : End Sub
Sub sw78_UnHit : UpdateWires 78,false : End Sub
Sub sw79_UnHit : UpdateWires 79,false : End Sub
Sub sw83_UnHit : UpdateWires 83,false : End Sub
Sub sw84_UnHit : UpdateWires 84,false : End Sub

'******************
' Battering Ram
'******************
Sub BatteringRam_Hit
    If Tilted Then Exit Sub
    Dim hitstrength
  hitstrength = ABS(ActiveBall.VelY)

    ' Slow the ball down, to simulate hitting a big battering ram
  ActiveBall.AngMomX = ActiveBall.AngMomX / 1+(Rnd*3)

  DoRamDOF hitstrength
  DoRam hitstrength

  If hitstrength >= 10 then DoRamScoring
End Sub

Sub BatteringRam_Mid_Hit
    If Tilted Then Exit Sub
    Dim hitstrength
  hitstrength = ABS(ActiveBall.VelY)

    ' Slow the ball down, to simulate hitting a big battering ram
  ActiveBall.AngMomX = ActiveBall.AngMomX / 1+(Rnd*3)

  DoRamDOF hitstrength
  DoRam hitstrength

  If hitstrength >= 10 then DoRamScoring
End Sub

Sub BatteringRam_High_Hit
    If Tilted Then Exit Sub
    Dim hitstrength
  hitstrength = ABS(ActiveBall.VelY)

    ' Slow the ball down, to simulate hitting a big battering ram
  ActiveBall.AngMomX = ActiveBall.AngMomX / 1+(Rnd*3)

  DoRamDOF hitstrength
  DoRam hitstrength

  If hitstrength >= 10 then DoRamScoring
End Sub

sub DoRamDOF(hitstrength)
  if hitstrength < 5 then exit sub

  'trigger shaker DOF, turn off depending on velocity (<5-6 soft hit, 9-12 med, 13+ strong)
  DOF 124,DOFOn
  vpmTimer.AddTimer 125*hitstrength,"BatteringRamDOF_Off '"
end Sub

Sub BatteringRamDOF_Off
  DOF 124,DOFOff
  'Debug.Print "BatteringRamDOF Off"
End Sub

Sub DoRamScoring
  Dim scene,line1,i

  SetFlasher fl243,6 ' Flash the battering ram Flasher

    if bMadnessMB = 2 Then DoMadnessMBHit
    AddScore 1130
    AddBonus 5000
    PlaySoundVol "gotfx-battering-ram",Voldef
    If bWildfireLit = 2 Then ' Mini-mode hit
        House(CurrentPlayer).AddWildfire 10
        TotalWildfire = TotalWildfire + 10
        If bUseFlexDMD And bBlackwaterSJPMode = False Then
            ' Do Scene for Wildfire Mini Mode
            Set scene = NewSceneWithVideo("wfmm","got-wildfiremode")
            scene.AddActor FlexDMD.NewLabel("ttl",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbBlack, 0), _
                            CurrentWildfire&" TOTAL"&vbLf&"JACKPOT = "& House(CurrentPlayer).BWJackpot)
            scene.AddActor FlexDMD.NewLabel("obj",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),"+10 WILDFIRE")
            scene.GetLabel("obj").SetAlignedPosition 64,8,FlexDMD_Align_Center
            DelayActor scene.GetLabel("obj"),1,True
            scene.GetLabel("ttl").SetAlignedPosition 64,22,FlexDMD_Align_Center
            DelayActor scene.GetLabel("ttl"),1,True
            i = (Int(CurrentWildfire/10) MOD 2) + 1
            DMDEnqueueScene scene,1,1800,4000,1500,"gotfx-wildfiremini"&i
        End If
    ElseIf bWildfireLit = True And bMultiBallMode = False Then
        ' Start Wildfire Mini Mode
        SetGameTimer tmrWildfireMode,220    ' 22 second mode timer
        li126.BlinkInterval = 100
        SetLightColor li126, darkgreen, 2
        If bUseFlexDMD And bBlackwaterSJPMode = False Then
            ' Do Scene for Wildfire Mini Mode
            Set scene = NewSceneWithVideo("wfmm","got-wildfiremode")
            scene.AddActor FlexDMD.NewLabel("mode",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbBlack, 0),"WILDFIRE MINI-MODE")
            scene.AddActor FlexDMD.NewLabel("obj",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbBlack, 0),"HIT   BATTERING   RAM"&vbLf&"TO   COLLECT   WILDFIRE")
            scene.GetLabel("mode").SetAlignedPosition 64,7,FlexDMD_Align_Center
            scene.GetLabel("obj").SetAlignedPosition 64,20,FlexDMD_Align_Center
            DelayActor scene.GetLabel("obj"),1,True
            DelayActor scene.GetLabel("mode"),1,True
            DMDEnqueueScene scene,1,1800,4000,1500,"gotfx-wildfireministart"
        End If
        bWildfireLit = 2
    End If
    If bBlackwaterSJPMode Then House(CurrentPlayer).ScoreSJP 0    'Super Jackpot!!
    If bMysterySJPMode > 0 Then House(CurrentPlayer).ScoreSJP 0 : bMysterySJPMode = 0

    If bHotkMBActive And House(CurrentPlayer).HotkState = 1 Then House(CurrentPlayer).ScoreHotkJackpot 1,1
    If bITMBActive And House(CurrentPlayer).ITState = 2 Then House(CurrentPlayer).ScoreITJP 1,1,1

    'Playfield Multiplier support
    'PFMState:
    ' 0 - all lights off
    ' 1 - blue arrow flashing - timer active
    ' 2 - PFM light on solid
    ' 3 - PFM light flashing - timer active
    ' Hits advance state. A hit on state 3 awards PF multiplier. Timer resets state to 0
    Select Case PFMState
        Case 0
            If PlayfieldMultiplierVal < 3 Or (PlayfieldMultiplierVal < 3 + SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier) And PlayfieldMultiplierVal < 5) Then
                SetGameTimer tmrPFMState,200    ' 20 seconds
                line1 = ""
            End If
        Case 1
            TimerFlags(tmrPFMState) = 0
            line1 = "LIGHT  PLAYFIELD"&vbLf&"MULTIPLIER"
        Case 2
            SetGameTimer tmrPFMState,200    ' 20 seconds
            line1 = ""
        Case 3
            TimerFlags(tmrPFMState) = 0
            PlayfieldMultiplierVal = PlayfieldMultiplierVal + 1
            SetGameTimer tmrPFMultiplier,800-(PlayfieldMultiplierVal*100)
            SetPFMLights
            line1 = "+1  PLAYFIELD"&vbLf&"MULTIPLIER"
            SetLockbarLight
        Case Else
            ' Special case. When the PF Multiplier timer runs out, there's a short grace period where a single
            ' hit will restore it
            If PFMState > 3 Then
                PlayfieldMultiplierVal = PFMState - 2
                SetGameTimer tmrPFMultiplier,760-(PlayfieldMultiplierVal*80)
                SetPFMLights
                SetLockbarLight
                PFMState = 3
                line1 = ""
            End If
    End Select
    If PFMState <> 0 Or (PlayfieldMultiplierVal < 3 Or (PlayfieldMultiplierVal < 3 + SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier) And PlayfieldMultiplierVal < 5) ) Then
        PFMState = PFMState + 1 : If PFMState = 4 Then PFMState = 0
        If bWildfireLit = False Then DoBatteringRamScene line1
    End If
    SetBatteringRamLights
End Sub

' *** New Ram Physics Implementation

Const RamSpeedFactor = 10 ' higher val accelerates the ram speed, gives appearance of less mass
Const MaxRamMovement = 35 ' max number of units ram is allowed to move on the Y axis

Dim RamVelocity : RamVelocity = 0
Dim RamApex : RamApex = 0 ' highest point during a hit
Dim RamApexIndex : RamApexIndex = 0   ' stores the time index of the apex
Dim RamDirection : RamDirection = 0   ' 0=still, 1=moving up, 2=moving down
Dim RamOffset : RamOffset = 0       ' Y-offset when hit, typically 0 at rest
Dim RamBeginTime : RamBeginTime = 0   ' gametime when the ram starts
Dim RamTimerIntervalMs : RamTimerIntervalMs = 0 ' actual interval of the timer

'does all the setup for the ram being hit
' - scales the velocity to an appropriate range for min/max range of ram
' - figures out if the ram is hit while stationary or moving
sub DoRam(vel)

  'scale velocity so ram peaks at about ~35 units max
  'vel tends to max out around 15-16 for a solid direct hit but varies based on physics
  'minS and minT are currently 0 but leaving here for possible changes
  dim scaledVel, minT, maxT, minS, maxS, f
  minT = 0 : maxT = 27 ' scaled range to make ram move entire range; lower this to make ram move less
  minS = 0 : maxS = 16 ' typical range of most shots

  'make sure the RamSettleTimer is disabled
  if RamSettleTimer.enabled <> 0 then
    RamSettleTimer.enabled = 0
    For each f in ram_BL : f.transy = 0 : Next
  end if

  'very soft hit, jiggle the ram
  if vel < 5 and RamDirection = 0 and not RamTimer.Enabled then
    RamVelocity = vel + 16 'set an artifical ram velocity for the settle timer
    RamSettleTimer.UserValue = 0
    RamSettleTimer.enabled = true
    exit sub
  end if

  'take the maxT value in case an unusually large hit comes in beyond maxS
  'otherwise, scale the value
  scaledVel = min(maxT,((vel - minS) / (maxS - minS)) * (maxT - minT) + minT)
  'debug.print "Ramming vel: " & vel & " scaled to " & scaledVel & " gametime: " & gametime

  'is ram already moving?
  if RamTimer.Enabled and RamDirection <> 0 then
    'ram is moving
    if RamDirection = 1 then
      'moving up, continue moving up
      RamOffset = abs(ram_BM_Dark_Room.TransY)
      RamVelocity = scaledVel * 1.2 ' add in some vel gain
      'debug.print "**hit moving up, offset: " & RamOffset
    elseif RamDirection = 2 then
      'moving down, but need to reverse
      RamDirection = 1 ' switch to up
      RamOffset = abs(ram_BM_Dark_Room.TransY)
      RamVelocity = scaledVel * 0.8 ' add in some vel loss
      'debug.print "**hit moving down, offset: " & RamOffset
    end if
  else
    'ram is not moving or is at rest
    RamOffset = 0
    RamVelocity = scaledVel
  end if

  RamBeginTime = gametime
  RamApex = 0
  RamTimer.Enabled = 1
  RamTimer.UserValue = 0
end Sub

'calculates the height of the ram as a projectile based on time (ticks of the timer) and starting velocity
Function CalcRamHeight(v, time)
  Const g = 9.81 ' Acceleration m/s^2
  'tweak projectileAngle for visual effect (the 'arc' of the object)
  'higher values (max 90) produce a sharp parabola, like throwing a ball straight up
  Const projectileAngle = 75 ' good overall arc
  Dim rad ' convert to radians
  rad = projectileAngle * (Atn(1) * 4) / 180
  CalcRamHeight = (v * Sin(rad) * time) - (0.5 * g * (time^2))
end Function

Sub SetRamWallState(height) ' 0=default, 1=mid, 2=high
  'sets which wall in the ram cascade to use
  'debug.print "RamWallState: " & height
  Select case height
    case 1:
      BatteringRam_Mid.Collidable = True
      BatteringRam.Collidable = False
      BatteringRam_High.Collidable = False
    case 2:
      BatteringRam_High.Collidable = True
      BatteringRam.Collidable = False
      BatteringRam_Mid.Collidable = False
    case else:
      BatteringRam.Collidable = True
      BatteringRam_Mid.Collidable = False
      BatteringRam_High.Collidable = False
  end Select

  'BatteringRam.Collidable = False
  'BatteringRam.Visible = False
  'BatteringRam.SideVisible = False
  'BatteringRam_Mid.Collidable = False
  'BatteringRam_Mid.Visible = False
  'BatteringRam_Mid.SideVisible = False
  'BatteringRam_High.Collidable = False
  'BatteringRam_High.Visible = False
  'BatteringRam_High.SideVisible = False
end Sub

Sub RamTimer_Timer
  'updates the ram based on the velocity and time
  'timer is written so that the interval can change without affecting location
  'this requires knowing the interval between ticks
    Dim timeinterval, curHeight, i, f

  if RamTimer.UserValue = 0 then
    'first hit, reset interval - only needed if setting to a higher interval for performance
    RamTimer.Interval = -1 ' recommend < 50 if not using -1
    RamDirection = 1 ' moving up
  end if

  RamTimer.UserValue = RamTimer.UserValue + 1

  timeinterval = ((gametime - RamBeginTime)/1000) * RamSpeedFactor 'multiply to inc/dec anim time
  curHeight = CalcRamHeight(RamVelocity,timeinterval) + RamOffset

  'constrain movement between 0 and MaxYHeight units
  if curHeight < 0 then curHeight = 0
  if curHeight > MaxRamMovement then
    'should only get here if ram is hit while already high, causing it to peg the top.
    debug.print "Ram clipping height " & curHeight & " to " & MaxRamMovement

    'get the current actual interval in ms
    RamTimerIntervalMs = (gametime - RamBeginTime)/RamTimer.UserValue

    'figure out how long (timer ticks) the ram would be pegged
    for i = RamTimer.UserValue + 1 to RamTimer.UserValue + 200 ' magic number, varies based on interval + velocity
      timeinterval = (i * (RamTimerIntervalMs / 1000)) * RamSpeedFactor 'multiply to inc/dec anim time
      if CalcRamHeight(RamVelocity,timeinterval) + RamOffset <= MaxRamMovement then
        'found the point in time where the ram will no longer be pegged
        'debug.print "not clipping at index: " & i & " timeinterval: " & timeinterval
        exit for
      end if
    next

    'if the ram is going to hang more than about 200ms, accelerate so it drops faster
    if (i - RamTimer.UserValue)*RamTimerIntervalMs > 200 then
      'moves back the start time time make the ram drop faster
      RamBeginTime = RamBeginTime - ((i - RamTimer.UserValue)*RamTimerIntervalMs - 200)
    end if

    curHeight = MaxRamMovement
  end if

  'track highest point
  if curHeight > RamApex then
    RamApex = curHeight
    RamApexIndex =  RamTimer.UserValue
    RamDirection = 1 ' must be moving up
  else
    RamDirection = 2 ' must be moving down
  end if

  For each f in ram_BL : f.transy = -(curHeight) : Next ' negate the value since ram is moving up the pf

  'cascade the hit target walls
  if abs(ram_BM_Dark_Room.TransY) > 19 then
    SetRamWallState 2 'ram is high
  elseif abs(ram_BM_Dark_Room.TransY) > 8 then
    SetRamWallState 1 'ram is mid
  else
    SetRamWallState 0 'ram is at rest
  end if

  if curHeight = 0 then
    'track the interval count to tweak smoothness. should be able to change interval higher or lower
    debug.print "Ram timer ended, scaled vel: " & RamVelocity & " count: " & RamTimer.UserValue & " apex: " & RamApex
    RamTimer.Enabled = 0
    RamTimer.UserValue = 0
    RamDirection = 0
    RamApex = 0
    RamOffset = 0

    if RamVelocity > 10 then ' if a solid hit, bounce the ram to "settle" it
      RamSettleTimer.UserValue = 0
      RamSettleTimer.enabled = true
    end if
  end If
End Sub

Sub RamSettleTimer_Timer
  'bounce the ram when settling
  dim f
  RamSettleTimer.UserValue = RamSettleTimer.UserValue + 1

  if RamSettleTimer.UserValue < cInt(RamVelocity/4) then
    'move the ram up a bit to simulate a bounce effect
    For each f in ram_BL : f.transy = f.transy-1 : Next
  else
    For each f in ram_BL : f.transy = f.transy+1 : Next
  end if

  if ram_BM_Dark_Room.TransY >= 0 then
    debug.print "Ram settle timer ended, counter: " & RamSettleTimer.UserValue & ", height: " & cInt(RamVelocity/4)
    For each f in ram_BL : f.transy = 0 : Next
    RamSettleTimer.UserValue = 0
    RamSettleTimer.enabled = 0
  end if
End Sub

' *** End New Ram Implementation

Sub DoBatteringRamScene(line1)
    Dim scene
    If bUseFlexDMD Then
        If line1 = "" Then
            Set scene = NewSceneWithVideo("pfmq","got-batteringram")
            DMDEnqueueScene scene,1,500,500,1000,""
        Else
            Set scene = NewSceneWithVideo("wfmm","got-wildfiremode")
            scene.AddActor FlexDMD.NewLabel("obj",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbBlack, 0),line1)
            scene.GetLabel("obj").SetAlignedPosition 64,16,FlexDMD_Align_Center
            DelayActor scene.GetLabel("obj"),1,True
            DMDEnqueueScene scene,1,1800,4000,1500,"gotfx-wildfiremini"& max(1,PFMState/2)
        End If
    Else
        If line1 <> "" Then
            DisplayDMDText line1,"",2000
            PlaySoundVol "gotfx-wildfiremini"& max(1,PFMState/2),VolDef
        End If
    End If
End Sub

Sub WildfireModeTimer
    bWildfireLit = False
    SetLightColor li126, darkgreen, 0
    TimerFlags(tmrWildfireMode) = 0
End Sub

Sub PFMStateTimer
    PFMState = 0
    SetLightColor li132, darkblue, 0
    SetLightColor li123, amber, 0
End Sub

Sub PFMultiplierTimer
    PFMState = PlayfieldMultiplierVal + 2
    PlayfieldMultiplierVal = 1
    SetPFMLights
    SetLockbarLight
    vpmTimer.AddTimer 4000,"PFMState=0 '"
End Sub

Sub SetBatteringRamLights
    Select Case PFMState
        Case 2: SetLightColor li123, amber, 1
        Case 3: li123.BlinkInterval = 100 : SetLightColor li123, amber, 2
        Case Else: li123.State = 0
    End Select
    If (PFMState MOD 2) = 1 Or bWildfireLit <> 0 Or bBlackwaterSJPMode Then
        SetLightColor li132,midblue,2
    Else
        SetLightColor li132,midblue,0
    End If
    If bBlackwaterSJPMode Or bMysterySJPMode > 0 Then
        li132.BlinkInterval = 66
        li129.BlinkInterval = 66
        SetLightColor li129,red,2
    ElseIf bHotkMBActive And House(CurrentPlayer).HotkState = 1 Then
        li132.BlinkInterval = 66
        li129.BlinkInterval = 66
        SetLightColor li129,white,2
        SetLightColor li132,white,2
    ElseIf bITMBActive And House(CurrentPlayer).ITState = 2 Then
        li132.BlinkInterval = 66
        li129.BlinkInterval = 66
        SetLightColor li129,cyan,2
        SetLightColor li132,cyan,2
    Else
        li132.BlinkInterval = 100
        li129.State = 0
    End If
End Sub

'**********
' Spinner
'**********
Sub Spinner001_Spin
    If Tilted then Exit Sub
    Dim spinval
    Me.TimerEnabled = False
    Me.TimerInterval = 1000
    Me.TimerEnabled  = True
    If PlayerMode = 1 and (HouseBattle1 = Baratheon or HouseBattle2 = Baratheon) Then
        If AccumulatedSpinnerValue = 0 Then FlashPlayfieldLights yellow,10,15
        House(CurrentPlayer).BattleState(Baratheon).RegisterSpinnerHit
        AccumulatedSpinnerValue = 1 ' Ensure a new scene isn't created for each spin hit
        Exit Sub
    End If

    SetFlasher fl242,6    ' Flash the spinner Flasher

    DOF 330,DOFPulse

    If bMadnessMB = 2 Then DoMadnessMBHit
    ' According to Chukwurt, Max spinner value per level are roughly 30k,60k,150k,250k,400k,500k,650k,800k
    spinval = SpinnerValue
    AccumulatedSpinnerValue = AccumulatedSpinnerValue + spinval
    AddScore spinval
    If bMadnessMB = 2 Then DoMadnessMBHit : Else DMDSpinnerScene spinval
    PlaySoundVol "gotfx-drum"&SpinnerLevel,VolDef/4
End Sub

Sub Spinner001_Timer: AccumulatedSpinnerValue = 0: End Sub


'************************
' Dragon Control
'************************

Dim WingSteps
WingSteps = Array(0,0,1,3,9,15,21,27,29,30,30,29,27,21,15,9,3,1,0)
Sub StartDragonWings(lp)
    PlaySound "gotfx-dragonwings",-1,VolDef/4
    tmrDragonWings.UserValue = 0
    If lp = False Then tmrDragonWings.UserValue = -18
    tmrDragonWings.Interval = 80
    tmrDragonWings.Enabled = True
    If (VRRoom = 1 or VRRoom = 3 or VRRoom = 4 or VRRoom = 5)and Topper = 1 then StartVRDragon
End Sub

Sub StopDragonWings
    StopSound "gotfx-dragonwings"
    If tmrDragonWings.UserValue > 0 Then
        tmrDragonWings.Uservalue = tmrDragonWings.UserValue * -1
    ElseIf tmrDragonWings.UserValue = 0 Then
        tmrDragonWings.Enabled = False
    End If
  If (VRRoom = 1 or VRRoom = 3 or VRRoom = 4 or VRRoom = 5) and Topper = 1 then StopVRDragon
End Sub

Sub tmrDragonWings_Timer
    Dim i,wi
    i = tmrDragonWings.UserValue
    i = i + 1
    If i > 18 Then i = 1

  'trigger dof on when wings begin at either i=-17 or i=2
  'i: -17 is when triggered solo, 2 is when continuous
  If i=-17 or i=2 Then DOF 135, DOFPulse

    For each wi in Rwing_BL : wi.ObjRotY = WingSteps(Abs(i))    : Next
    For each wi in Lwing_BL : wi.ObjRotY = WingSteps(Abs(i))*-1 : Next
    tmrDragonWings.UserValue = i
    If i = 0 Then
        tmrDragonWings.Enabled = False
        StopSound "gotfx-dragonwings"
      If (VRRoom = 1 or VRRoom = 3 or VRRoom = 4 or VRRoom = 5) and Topper = 1 then StopVRDragon
    End If
End Sub

Sub DoDragonRoar(num)
    If num > 0 Then PlaySoundVol "gotfx-dragonroar"&num,VolDef
  DOF 134, DOFPulse
    StartDragonWings False
    If Options_Flames_Enabled Then
        DragonFire.UserValue = 1
        DragonFire.imageA = "flame1"
        DragonFire.Visible = 1
        DragonFire.TimerInterval = 100
        DragonFire.TimerEnabled = True
    End If
End Sub

Sub DragonFire_Timer
    Dim i
    i = DragonFire.UserValue
    i = i + 1
    If i = 3 Then DragonFire.TimerInterval = 300 Else DragonFire.TimerInterval = 100
    DragonFire.TimerEnabled = 0
    If i > 5 Then
        DragonFire.visible = 0
    Else
        DragonFire.ImageA = "flame"&i
        DragonFire.UserValue = i
        DragonFire.TimerEnabled = True
    End If
End Sub


'*************************
' Mystery Award
'******************

Dim MysteryAwards(29)
Dim MysteryVals(2)
Dim MysterySelected

MysteryAwards(0) = ARRAY("KEEP YOUR"&vbLf&"GOLD",0)
MysteryAwards(1) = Array(1,120)
MysteryAwards(2) = Array("LIGHT A"&vbLf&"HOUSE",140)
MysteryAwards(3) = Array("+1 BONUS X",160)
MysteryAwards(4) = Array("GROW"&vbLf&"WALL"&vbLf&"JACKPOT",190)
MysteryAwards(5) = Array("LIGHT"&vbLf&"WILDFIRE",220)
MysteryAwards(6) = Array("LIGHT"&vbLf&"LOCK",250)
MysteryAwards(7) = Array("+5"&vbLf&"WILDFIRE",270)
MysteryAwards(8) = Array("LIGHT"&vbLf&"SWORDS",320)
MysteryAwards(9) = Array("LIGHT PF"&vbLf&"MULTIPLY",420)
MysteryAwards(10) = Array("COLLECT"&vbLf&"SWORD",620)
MysteryAwards(11) = Array("ADVANCE"&vbLf&"CASTLE"&vbLf&"MULTIBALL",830)
MysteryAwards(12) = Array("ADVANCE"&vbLf&"WALL"&vbLf&"MULTIBALL",1130)
MysteryAwards(13) = Array("LIGHT 1X"&vbLf&"SUPER"&vbLf&"JACKPOT",1130)
MysteryAwards(14) = Array("+2 BONUS X",1430)
MysteryAwards(15) = Array("+10"&vbLf&"WILDFIRE",1880)
MysteryAwards(16) = Array("START"&vbLf&"WINTER"&vbLf&"IS COMING",2030)
MysteryAwards(17) = Array("+2 CASTLE"&vbLf&"MULTIBALL",2330)
MysteryAwards(18) = Array(5,2600)
MysteryAwards(19) = Array("+3 BONUS X",2970)
MysteryAwards(20) = Array("LIGHT"&vbLf&"LORD OF"&vbLf&"LIGHT",3530)
MysteryAwards(21) = Array("BARATHEON"&vbLf&"BUTTON"&vbLf&"ABILITY",3800)
MysteryAwards(22) = Array("MARTELL"&vbLf&"BUTTON"&vbLf&"ABILITY",3800)
MysteryAwards(23) = Array("LANNISTER"&vbLf&"BUTTON"&vbLf&"ABILITY",3800)
MysteryAwards(24) = Array("LIGHT 2X"&vbLf&"SUPER"&vbLf&"JACKPOT",3920)
MysteryAwards(25) = Array("HOLD"&vbLf&"BONUS",4200)
MysteryAwards(26) = Array("LIGHT"&vbLf&"EXTRA"&vbLf&"BALL",4500)
MysteryAwards(27) = Array("LIGHT 3X"&vbLf&"SUPER"&vbLf&"JACKPOT",4750)
MysteryAwards(28) = Array(25,6000)

' Choose 2 "random" mysteries from the array. One will use a moderate amount of their gold, the other will use as much as possible.
' The third (First) option is option 0 - keep gold
' Player uses flipper buttons to choose and action button to select
' Start 3 timers. One fires after a few seconds and says "Choose!", one fires with 5 seconds left and says "choose now!"
' Last timer fires after 20? seconds and chooses whatever is selected
dim TestMysteryHouse : TestMysteryHouse = False
Sub DoMysteryAward
    Dim i,j : j=0
    If Tilted Then exit Sub
    MysteryVals(0) = 0 : MysteryVals(1)=0 : MysteryVals(2)=0
    For i = 1 to 18+RndNbr(10)
        If MysteryAwards(i)(1) =< CurrentGold And MysteryAwardGoNoGo(i) Then MysteryVals(2) = i
    Next
    Do
        i = RndNbr(28)
        j = j + 1
    Loop Until (MysteryAwards(i)(1) =< CurrentGold And i <> MysteryVals(2) And MysteryAwardGoNoGo(i)) Or j = 100
    If j = 100 Then MysteryVals(1) = MysteryVals(2) Else MysteryVals(1) = i
  'test: force mystery lite house
  if TestMysteryHouse then MysteryVals(1) = 2
    bMysteryAwardActive = True
    MysterySelected = 0
    bMysteryLit = False : SetMysteryLight
    i = RndNbr(5)
    MATstep = 0
    PlaySoundVol "say-make-your-choice"&i,VolCallout
    SetGameTimer tmrMysteryAward,10
    PlayModeSong
    SetLockbarLight

    DMDMysteryAwardScene
End Sub

Function MysteryAwardGoNoGo(aw)
    Dim i
    MysteryAwardGoNoGo = True
    Select Case aw
        Case 2
            For i = 1 to 7
                If House(CurrentPlayer).Qualified(i) = False And House(CurrentPlayer).Completed(i) = False Then Exit Function
            Next
            MysteryAwardGoNoGo = False
        Case 3: If BonusMultiplier(CurrentPlayer) >= MaxBonusMultiplier Then MysteryAwardGoNoGo = False
        Case 5: If bWildfireLit Then MysteryAwardGoNoGo = False
        Case 6: If bLockIsLit Then MysteryAwardGoNoGo = False
        Case 8: If bSwordLit Then MysteryAwardGoNoGo = False
        Case 9: If PlayfieldMultiplierVal = 5 Or PlayfieldMultiplierVal >= 3+SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier) Or PFMState > 1 Then MysteryAwardGoNoGo = False
        Case 10: If SwordsCollected > 10 Then MysteryAwardGoNoGo = False
        Case 11: If House(CurrentPlayer).UPFLevel = 4 Then MysteryAwardGoNoGo = False
        Case 14: If BonusMultiplier(CurrentPlayer) >= 19 Then MysteryAwardGoNoGo = False
        Case 16: If House(CurrentPlayer).WiCMask >= 254 Then MysteryAwardGoNoGo = False
        Case 17: If House(CurrentPlayer).UPFLevel >= 3 Then MysteryAwardGoNoGo = False
        Case 19: If BonusMultiplier(CurrentPlayer) >= 18 Then MysteryAwardGoNoGo = False
        Case 20: If bLoLUsed = False Then MysteryAwardGoNoGo = False
        Case 21: If House(CurrentPlayer).ActionAbility = Baratheon Then MysteryAwardGoNoGo = False
        Case 22: If House(CurrentPlayer).ActionAbility = Martell Then MysteryAwardGoNoGo = False
        Case 23: If House(CurrentPlayer).ActionAbility = Lannister Then MysteryAwardGoNoGo = False
        Case 25: If bBonusHeld > 0 Then MysteryAwardGoNoGo = False
    End Select
End Function


Dim MATstep
Sub MysteryAwardTimer
    Dim i,j,snd
    snd = "":i=1
    MATstep = MATstep + 1
    If MATstep=6 And MysterySelected = 0 Then snd="say-make-your-choice-quickly":i=6
    If MATstep=10 And MysterySelected = 0 Then snd="say-gold-is-always-useful":i=2
    j = RndNbr(i)
    If snd <> "" Then PlaySoundVol snd&j,VolCallout
    If bUseFlexDMD Then
        FlexDMD.LockRenderThread
        MysteryScene.GetLabel("tmr").Text = CStr(10-MATstep)
        FlexDMD.UnlockRenderThread
    End If
    If MATstep = 10 Then
        SelectMysteryAward
    Else
        SetGameTimer tmrMysteryAward,10
    End If
End Sub

Sub UpdateMysteryAward(keycode)
    If keycode = RightMagnaSave or keycode = LockBarKey or _
      (keycode = PlungerKey and bUsePlungerForSternKey) Then SelectMysteryAward
    If keycode = LeftFlipperKey Then
        MysterySelected = MysterySelected - 1
        If MysterySelected < 0 Then MysterySelected = 2
        PlaySoundVol "gotfx-mysteryleft",VolDef/4
        DMDMysteryAwardScene
    ElseIf keycode = RightFlipperKey Then
        MysterySelected = MysterySelected + 1
        If MysterySelected > 2 Then MysterySelected = 0
        PlaySoundVol "gotfx-mysteryright",VolDef/4
        DMDMysteryAwardScene
    End If
End Sub

Sub SelectMysteryAward
    Dim i,j,combo
    bMysteryAwardActive = False
    bMysteryLit = False : SetMysteryLight : SetTopGates
    DMDMysteryAwardScene
    TimerFlags(tmrMysteryAward) = 0
    PlaySoundVol "gotfx-mysteryselect",VolDef
    combo = ROrbitsw31.UserValue : if combo > 10 then combo = combo - 10
    If combo = 0 then combo = 1
    Select Case MysteryVals(MysterySelected)
        Case 0: 'keep gold
        Case 1
            AddScore 1000000*combo
            DMD "BIG POINTS",FormatScore(1000000*PlayfieldMultiplierVal*combo),"",eNone,eNone,eNone,1000,True,""
        Case 2  ' light a house
            'For i = 1 to 7
            '    If House(CurrentPlayer).Qualified(i) = False And House(CurrentPlayer).Completed(i) = False Then Exit For
            'Next
      'AwardMysteryHouse(i)
      i = GetRandomUnlitHouse()
      AwardMysteryHouse i
            DMD "",HouseToUCString(i)&" IS LIT","",eNone,eNone,eNone,1000,True,""
        Case 3: IncreaseBonusMultiplier 1  ' +1X
        Case 4: IncreaseWallJackpot
        Case 5: DoWildfireLit
        Case 6: LightLock
        Case 7: TotalWildfire = TotalWildfire + 5: House(CurrentPlayer).AddWildfire 5  ' Increase wildfire. TODO: Should play a scene
        Case 8: DoSwordsAreLit
        Case 9: PFMState = 2 : SetBatteringRamLights : DoBatteringRamScene "LIGHT PLAYFIELD"&vbLf&"MULTIPLIER"
        Case 10: DoAwardSword
        Case 11: House(CurrentPlayer).IncreaseUPFLevel : House(CurrentPlayer).SetUPFLights
        Case 12: AdvanceWallMultiball 1
        Case 13,24,27
            Select Case MysteryVals(MysterySelected)
                Case 13: bMysterySJPMode = 1
                Case 24: bMysterySJPMode = 2
                Case 27: bMysterySJPMode = 3
            End Select
            PlaySoundVol "say-getsjp",VolCallout
            SetBatteringRamLights
        Case 14: IncreaseBonusMultiplier 2
        Case 15: TotalWildfire = TotalWildfire + 10: House(CurrentPlayer).AddWildfire 10 ' Increase wildfire. TODO: Should play a scene
        Case 16: House(CurrentPlayer).StartWIC
        Case 17: House(CurrentPlayer).IncreaseUPFLevel : House(CurrentPlayer).IncreaseUPFLevel : House(CurrentPlayer).SetUPFLights ' TODO: Allow method to take an argument
        Case 18
            AddScore 5000000*combo
            DMD "BIGGER POINTS",FormatScore(5000000*PlayfieldMultiplierVal*combo),"",eNone,eNone,eNone,1000,True,""
        Case 19: IncreaseBonusMultiplier 3
        Case 20: DoLordOfLight False
        Case 21: House(CurrentPlayer).SetActionAbility Baratheon
        Case 22: House(CurrentPlayer).SetActionAbility Martell
        Case 23: House(CurrentPlayer).SetActionAbility Lannister
        Case 25: bBonusHeld = 1 : DMD "","BONUS HELD","",eNone,eNone,eNone,1000,True,""
        Case 26: DoEBisLit
        Case 28
            AddScore 25000000*combo
            DMD "BIGGEST POINTS",FormatScore(25000000*PlayfieldMultiplierVal*combo),"",eNone,eNone,eNone,1000,True,""
        Case 29: AwardSpecial
    End Select
    CurrentGold = CurrentGold - MysteryAwards(MysteryVals(MysterySelected))(1)
    vpmTimer.AddTimer 500,"IronThroneKickout '"
    DMDLocalScore
    PlayModeSong
    SetLockbarLight
End Sub

'returns a random unlit house, or zero if no unlit houses
Function GetRandomUnlitHouse()
  dim i, count
  dim unlitHouses()
  count = 0
  For i = 1 to 7
    If House(CurrentPlayer).Qualified(i) = False And House(CurrentPlayer).Completed(i) = False Then
      ReDim Preserve unlitHouses(count)
      unlitHouses(count) = i
      count = count + 1
    end If
    Next
  if count > 0 then
    GetRandomUnlitHouse = unlitHouses(RndInt(0, count-1))
  Else
    GetRandomUnlitHouse = 0
  End If
end Function

'handles specifics of lighting a house from mystery
Sub AwardMysteryHouse(i)
  dim existingCount ' count of qualified shots before award

  'debug.print "Pre QualCount:" & House(CurrentPlayer).QualifyCount(i) & " WiCShots:" & House(CurrentPlayer).WiCShots & " WiCs:" & House(CurrentPlayer).WiCs

  existingCount = House(CurrentPlayer).QualifyCount(i)
  House(CurrentPlayer).Qualified(i) = True
  House(CurrentPlayer).QualifyCount(i) = 3
  House(CurrentPlayer).bBattleReady = true

  'number of existing qualified shots = WiCShots
  House(CurrentPlayer).WiCShots = existingCount

  'todo:track WICs/inserts etc on a per-house basis?
  'this should be modified with whatever should happen when lighting a specific house
  Select Case i
    case Stark
    case Baratheon
      'drop target lights
      LoLTargetsCompleted = LoLTargetsCompleted + 3
      SetTargetLights
    Case Lannister
    Case Greyjoy
    Case Tyrell
    Case Martell
    Case Targaryen
  End Select

  debug.print "Post QualCount:" & House(CurrentPlayer).QualifyCount(i) & " WiCShots:" & House(CurrentPlayer).WiCShots & " WiCs:" & House(CurrentPlayer).WiCs

  SetBattleReadyLight
  House(CurrentPlayer).ResetLights
End Sub

'**************
'Bumper Hits
'**************
Dim bring1, bring2, bring3
Sub Bumper1_Hit
    If Tilted then Exit Sub
    'PlaySoundAt SoundFXDOF("fx_bumper1", 107, DOFPulse, DOFContactors),Bumper1
    RandomSoundBumperTop Bumper1
    DOF 107, DOFPulse
    doPictoPops 0
    bring1=1
    tmrBumperRing1.enabled=1
End Sub

Sub Bumper2_Hit
    If Tilted then Exit Sub
    'PlaySoundAt SoundFXDOF("fx_bumper2", 108, DOFPulse, DOFContactors),Bumper2
    RandomSoundBumperMiddle Bumper2
    DOF 108, DOFPulse
    doPictoPops 1
    bring2=1
    tmrBumperRing2.enabled=1
End Sub

Sub Bumper3_Hit
    If Tilted then Exit Sub
    'PlaySoundAt SoundFXDOF("fx_bumper3", 109, DOFPulse, DOFContactors),Bumper3
    RandomSoundBumperBottom Bumper3
    DOF 109, DOFPulse
    doPictoPops 2
    bring3=1
    tmrBumperRing3.enabled=1
End Sub

' On Bumper hit, Move the locations of the Bumper ring lightmap/bakemap prims
Sub tmrBumperRing1_Timer
    Dim z,b
    Select Case bring1
        Case 1:z = -30:bring1 = 2
        Case 2:z = -20:bring1 = 3
        Case 3:z = -10:bring1 = 4
        Case 4:z = -0:Me.Enabled = 0
    End Select
    For Each b in Bumper_Ring_BL : b.TransZ = z : Next
End Sub

Sub tmrBumperRing2_Timer
    Dim z,b
    Select Case bring2
        Case 1:z = -30:bring2 = 2
        Case 2:z = -20:bring2 = 3
        Case 3:z = -10:bring2 = 4
        Case 4:z = 0:Me.Enabled = 0
    End Select
    For Each b in Bumper_Ring_001_BL : b.TransZ = z : Next
End Sub

Sub tmrBumperRing3_Timer
    Dim z,b
    Select Case bring3
        Case 1:z = -30:bring3 = 2
        Case 2:z = -20:bring3 = 3
        Case 3:z = -10:bring3 = 4
        Case 4:z = 0:Me.Enabled = 0
    End Select
    For Each b in Bumper_Ring_002_BL : b.TransZ = z : Next
End Sub

'***** Bumper Light Timers - Turn bumpers back on when they're done flashing ******
Sub Li168_Timer : Me.State = 1 : Me.TimerEnabled = 0 : End Sub
Sub Li171_Timer : Me.State = 1 : Me.TimerEnabled = 0 : End Sub
Sub Li174_Timer : Me.State = 1 : Me.TimerEnabled = 0 : End Sub

' Bumper1 timer is used to pause battle mode timers when ball hits a bumper
Sub Bumper1_Timer
    TimerFlags(tmrBattleMode1) = TimerFlags(tmrBattleMode1) And 253
    TimerFlags(tmrBattleMode2) = TimerFlags(tmrBattleMode2) And 253
    Bumper1.TimerEnabled = 0
End Sub

'********************************
' PictoPops support!
' (that's what Stern calls their
'  rotating award pop bumpers)
'********************************
' From the GoT ROM code, here's the awards seen:
' Add_A_Ball
' Add_Mode_Time
' Advance_Toward_Wall_Multiball
' Award_3_Bonus_Multipliers
' Award_5_Wild_Fire
' Award_Big_Points
' Award_Bonus_Multiplier
' Award_Gold
' Award_Special
' Bump_Wall_Jackpot_Value
' Bump_Winter_Is_Coming_Value
' LIGHT_SWORDS
' Light_Extra_ball
' Light_Lock
' Light_Lord_Of_Light
' Light_Mystery
' Light_Wild_Fire
'
' Here's how the algorithm works:
' Each award carries a weight indicating how preferred it is as a random selection
' Whenever a pop bumper is hit, a random number is generated between 0 and the sum of all weights
' Each weight is subtracted from the random number until the number is less than the weight of the next item.
' That item is the selected choice.
' If a chosen item can't be awarded at this time (e.g. Add a Ball during regular play), repeat the process
'
' Once an award has been given, generate 3 new random awards for the bumpers. At least one should be different
'
' The weights are predefined in an array and the sum can be calculated at table_init

Const BumperAwards = 17
Dim BumperWeightTotal
Dim BumperVals(2)
Dim PictoPops(17) 'Each element represents one pop award which is an array of 'long name','short name','weight', and 'mode'
                  'Mode determines when the award can be won. 0=anytime, 1=during multiball, 2=not multiball and not LockIsLit,
                  ' 3=during mode or hurry-up, 4=after LoLused, 5=mystery not lit, 6=swords not lit, 7=wild-fire not lit, 8=EB not yet awarded
PictoPops(1) = Array("+1 BONUS X","+1X",20,0)
PictoPops(2) = Array("+5 "&vbLf&"WILDFIRE","+WFIRE",20,0)
PictoPops(3) = Array("+150"&vbLf&"GOLD","+GOLD",20,0)
PictoPops(4) = Array("LIGHT "&vbLf&"SWORDS","L.SWORD",20,6)
PictoPops(5) = Array("INCREASE"&vbLf&"WINTER IS"&vbLf&"COMING","+WINTER",20,0) ' may need to change this if winter has come
PictoPops(6) = Array("INCREASE"&vbLf&"WALL"&vbLf&"JACKPOT","+POT",50,0) ''Battle for Wall Value Increases. Value=xxx'
PictoPops(7) = Array("LIGHT"&vbLf&"LOCK","L.LOCK",20,2)
PictoPops(8) = Array("BIG"&vbLf&"POINTS","+1M",20,0)
PictoPops(9) = Array("+3 BONUS X","+3X",12,0)
PictoPops(10) = Array("ADD TIME","+TIME",50,3) ' Higher weight, but only valid during Modes
PictoPops(11) = Array("ADD A BALL","+BALL",20,1)
PictoPops(12) = Array("ADVANCE"&vbLf&"WALL"&vbLf&"MULTIBALL","+WALL MB",20,0) '
PictoPops(13) = Array("LIGHT"&vbLf&"EXTRA"&vbLf&"BALL","EB LIT",12,8)
PictoPops(14) = Array("LORD"&vbLf&"OF"&vbLf&"LIGHT","LoL",10,4)
PictoPops(15) = Array("LIGHT"&vbLf&"MYSTERY","MYSTERY",20,5)
PictoPops(16) = Array("LIGHT"&vbLf&"WILDFIRE","WF LIT",20,7)
PictoPops(17) = Array("AWARD"&vbLf&"SPECIAL","SPECIAL",5,0)


Dim BumperLights
BumperLights = Array(li168,li171,li174)
Sub doPictoPops(b)
    Dim i,tmp
    ' Pick a random drum sound effect
    i = RndNbr(10)
    PlaySoundVol "gotfx-drum"&i,VolDef
    AddScore 1000

    Dim b1,b2:b1=1:b2=2
    Select Case b
        Case 1
            b1=0
        Case 2
            b2=0
    End Select

    ' Bumper Lights
    For i = 0 to 2
        If b = i Then BumperLights(i).BlinkInterval = 50 : BumperLights(i).State = 2 Else BumperLights(i).State = 0
        BumperLights(i).TimerEnabled = 0 : BumperLights(i).TimerInterval = 400 : BumperLights(i).TimerEnabled = 1
    Next

    SetFlasher fl237,6 ' Flash the Bumper Flasher

    If bMadnessMB = 2 Then DoMadnessMBHit : Exit Sub

    ' If Battle is running, pause battle mode timers and start a 1 second timer on Bumper1 to restart them
    If PlayerMode = 1 Then
        Bumper1.TimerEnabled = 0
        Bumper1.TimerInterval = 1000
        TimerFlags(tmrBattleMode1) = TimerFlags(tmrBattleMode1) Or 2
        TimerFlags(tmrBattleMode2) = TimerFlags(tmrBattleMode2) Or 2
        Bumper1.TimerEnabled = 1
    End If

    If (BumperVals(b) = BumperVals(b1) or BumperVals(b) = BumperVals(b2)) Then
        ' This bumper already matches one other. Check to make sure the value they're locked to is still valid
        If CheckPictoAward(BumperVals(b1)) = False Then GeneratePictoAward b1
        If CheckPictoAward(BumperVals(b2)) = False Then GeneratePictoAward b2
        DMDPictoScene
        Exit Sub
    End If
    GeneratePictoAward b
    DMDPictoScene

    ' Check to see whether all 3 match
    If (BumperVals(b) <> BumperVals(b1) or BumperVals(b) <> BumperVals(b2)) Then Exit Sub
    ' We have a winner!
    i = BumperVals(0)
    debug.print "PictoPops: Award " & PictoPops(i)(0)
    ResetPictoPops  ' Get em ready for the next round
    Select Case i
        Case 1: IncreaseBonusMultiplier 1
        Case 2      ' Increase wildfire. TODO: Should play a scene
            TotalWildfire = TotalWildfire + 5: House(CurrentPlayer).AddWildfire 5
        Case 3      ' Increase Gold
            If House(CurrentPlayer).HasAbility(Lannister) Then AddGold 250 Else AddGold 150
        Case 4      ' Light Swords.
            DoSwordsAreLit
        Case 5      ' Increase Winter Is Coming value
            IncreaseWinterIsComing
        Case 6      ''Battle for Wall Value Increases. Value=xxx'
            IncreaseWallJackpot
        Case 7: LightLock
        Case 8
            AddScore 1000000
            PlaySoundVol "gotfx-bigpoints",VolDef
            DMD "BIG POINTS",FormatScore(1000000*PlayfieldMultiplierVal),"",eNone,eNone,eNone,1000,True,""
        Case 9: IncreaseBonusMultiplier 3
        Case 10     ' Add Time (to mode or Hurry Up)
            If PlayerMode = 1 Then
                If HouseBattle1 > 0 Then House(CurrentPlayer).BattleState(HouseBattle1).AddTime 10
                If HouseBattle2 > 0 Then House(CurrentPlayer).BattleState(HouseBattle2).AddTime 10
            ElseIf bHurryUpActive Then
                HurryUpValue = HurryUpValue + (HurryUpChange * 50)
                'TODO: pLay an "Add Time" scene with two hourglasses
            End if
        Case 11: If bMultiBallMode Then DoAddABall : bAddABallUsed = True
        Case 12: AdvanceWallMultiball 1
        Case 13: DoEBisLit : bPictoEBAwarded = True
        Case 14: DoLordOfLight False
        Case 15: bMysteryLit = True : SetMystery
        Case 16: DoWildfireLit
        Case 17
            AwardSpecial
    End Select
End Sub

Sub GeneratePictoAward(b)
    Dim i,tmp,foundval: foundval = False
    Do While foundval = False
        tmp = RndNbr(BumperWeightTotal)-1
        For i = 1 to BumperAwards
            If tmp < PictoPops(i)(2) Then Exit For
            tmp = tmp - PictoPops(i)(2)
        Next
        ' Check to see if the new award is valid right now
        foundval = CheckPictoAward(i)
    Loop
    BumperVals(b) = i
End Sub

Function CheckPictoAward(val)
    CheckPictoAward = True
    Select Case PictoPops(val)(3)
        Case 1
            CheckPictoAward = (bMultiBallMode Or bITMBActive) And bAddABallUsed = False
        Case 2
            If bMultiBallMode Or bLockIsLit or bITMBActive Then CheckPictoAward = False
        Case 3
            If PlayerMode <> 1 And bHurryUpActive = False Then CheckPictoAward = False
        Case 4
            CheckPictoAward = bLoLUsed
        Case 5
            CheckPictoAward = Not bMysteryLit
        Case 6
            CheckPictoAward = Not bSwordLit
        Case 7
            If bWildfireLit or bBWMultiballActive Then CheckPictoAward = False
        Case 8
            If bPictoEBAwarded Then CheckPictoAward = False
    End Select
End Function

' Set new random values on the Pictopops. Ensure at least one is different
Sub ResetPictoPops
    Dim i
    Do
        For i = 0 to 2:GeneratePictoAward i: Next
    Loop While BumperVals(0) = BumperVals(1) And BumperVals(0) = BumperVals(2)
End Sub

Dim bBattleCreateBall
Sub LockBall
    Dim i,scene,g
    BallsInLock = BallsInLock + 1
    RealBallsInLock = RealBallsInLock + 1
    bLockIsLit = False
    SetLightColor li111,darkgreen,0     ' Turn off Lock light
    If House(CurrentPlayer).HasAbility(Lannister) Then g = 250 Else g = 75
    TotalGold = TotalGold + g
    CurrentGold = CurrentGold + g
    TotalWildfire = TotalWildfire + 5
    House(CurrentPlayer).AddWildfire 5
    i = RndNbr(3)
    if i > 1 Then i = ""
    PlaySoundVol "say-ball-" & BallsInLock & "-locked" & i, VolCallout
    If BallsInLock = 3 Then
        ' Was player unlucky enough to start a WiC? We have to clear BattleReady if so
        If PlayerMode = 2 Then House(CurrentPlayer).BattleReady = False
        'Start BW multiball in 1 second - gives a chance to say 'ball locked'
        vpmtimer.addtimer 1600, "StartBWMultiball '"
    ElseIf PlayerMode >= 0 Then  ' Regular mode - release new ball
        If RealBallsInLock > BallsInLock Then
            RealBallsInLock = RealBallsInLock - 1
            ReleaseLockedBall 0
        Else
            bAutoPlunger = True
            bPlayfieldValidated = False
            CreateNewBall
        End If
    Else ' battle mode selection and not all balls locked - let PreLaunchBattleMode take care of releasing a ball
        bBattleCreateBall = True
    End If
    If BallsInLock < 3 Then
        DoLockBallSeq
        If bUseFlexDMD Then
            Set scene = NewSceneWithImage("lock","got-balllock")
            scene.AddActor FlexDMD.NewLabel("lbl1",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite, vbWhite, 0) ,"BLACKWATER")
            scene.GetLabel("lbl1").SetAlignedPosition 2,1,FlexDMD_Align_TopLeft
            scene.AddActor FlexDMD.NewLabel("lbl2",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite, vbWhite, 0) ,"BALL "&BallsInLock&" LOCKED")
            scene.GetLabel("lbl2").SetAlignedPosition 2,10,FlexDMD_Align_TopLeft
            BlinkActor scene.GetLabel("lbl2"),100,10
            scene.AddActor FlexDMD.NewLabel("wf", FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0) ,"+5 WILDFIRE")
            scene.GetLabel("wf").SetAlignedPosition 2,25,FlexDMD_Align_BottomLeft
            scene.AddActor FlexDMD.NewLabel("gold", FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0) ,"+"&g&" GOLD")
            scene.GetLabel("gold").SetAlignedPosition 2,31,FlexDMD_Align_BottomLeft
            DMDEnqueueScene scene,0,1000,2000,500,""
        End If
    End If
End Sub

Sub StartBWMultiball
    bMultiBallMode = True
    bBWMultiballActive = True
    WildfireModeTimer   ' Stop Wildfire Minimode if it happened to be running
    bWildfireLit = False : SetWildfireLights
    AddScore 10000000
    CurrentWildfire = 0
    Dim scene
    If bUseFlexDMD Then
        Set scene = NewSceneWithVideo("bwmb","got-blackwatermb")
        scene.AddActor FlexDMD.NewLabel("balllock", FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0) ,"BALL 3"&vbLf&"LOCKED")
        scene.GetLabel("balllock").SetAlignedPosition 126,30,FlexDMD_Align_BottomRight
        BlinkActor scene.GetLabel("balllock"),0.1,12
        DMDEnqueueScene scene,0,2000,5000,500,"gotfx-blackwater-multiball-start"
    Else
        PlaySoundVol "gotfx-blackwater-multiball-start",VolDef
    End If
    PlayModeSong
    DoBWMultiballSeq
    tmrBWmultiballRelease.Interval = 5000 ' Long initial delay to give sequence time to complete
    tmrBWmultiballRelease.Enabled = True
    House(CurrentPlayer).SetBWJackpots
    BlackwaterScore = 10000000
    DMDCreateBWMBScoreScene
    EnableBallSaver DMDStd(kDMDFet_BWMB_SaveTimer)+5 ' feature setting plus the 5 second delay before MB starts
    SetLockbarLight
End Sub

Sub StartCastleMultiball
    bMultiBallMode = True
    bCastleMBActive = True
    WildfireModeTimer   ' Stop Wildfire Minimode if it happened to be running
    bWildfireLit = False : SetWildfireLights
    CastleMBScore = 0
    House(CurrentPlayer).CMBJackpot = 5000000
    House(CurrentPlayer).CMBJPIncr = 250000
    House(CurrentPlayer).CMBSJPValue = 20000000
    Dim scene
    If bUseFlexDMD Then
        Set scene = FlexDMD.NewGroup("castlemb")
        scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),"COMPLETE"&vbLf&"CASTLE DRAGON SHOTS"&vbLf&"TO LIGHT SUPER JACKPOT")
        scene.GetLabel("txt").SetAlignedPosition 64,16,FlexDMD_Align_Center
        DMDEnqueueScene scene,1,3000,3000,3000,""
    End If
  DOF 128, DOFOn
    SetPlayfieldLights
    House(CurrentPlayer).SetUPFState true
    House(CurrentPlayer).SetJackpots
    StartDragonWings true
    DoCastleMBSeq ' light seq
    DMDCreateCastleMBScoreScene ' alternate scoring scene
    EnableBallSaver DMDStd(kDMDFet_CMB_SaveTimer)
    SetTopGates ' all shots go to upper PF
    PlayModeSong
    SetLockbarLight
    StartNonBWMB 2
End Sub

Sub StartNonBWMB(balls)
    ' Changed for tmrBallLockCheck
    ' If RealBallsInLock > 0 Then
    '     tmrBWmultiballRelease.Interval = 1500
    '     tmrBWmultiballRelease.Enabled = True
    '     If balls > RealBallsInLock Then AddMultiball balls-RealBallsInLock
    ' Else
         AddMultiball balls ' add the balls
    ' End If
End Sub

Sub StartWallMB
    If Tilted Then exit Sub
    bMultiBallMode = True
    bWallMBActive = True
    bWallMBReady = False
    WallMBScore = 0
    tmrStartWallMB.Interval = 9000
    tmrStartWallMB.Enabled = 1
    If bUseFlexDMD Then
        Dim Scene
        Set scene = NewSceneWithVideo("wallmb","got-wallmbintro")
        DMDEnqueueScene scene,0,2000,8000,2000,"gotfx-wallmbintro"
    End If
    PlaySoundVol "say-wallmb-start"&RndNbr(3),VolCallout
    EnableBallSaver DMDStd(kDMDFet_WallMB_SaveTimer)
    SetPlayfieldLights
    House(CurrentPlayer).SetUPFState true
    House(CurrentPlayer).ResetForWallMB
    DMDCreateWallMBScoreScene
    DoWallMBSeq
    PlayModeSong
    SetLockbarLight
  DOF 132, DOFOn
End Sub

Sub tmrStartWallMB_Timer
    tmrStartWallMB.Enabled=0
    StartNonBWMB 2
    IronThroneKickout
End Sub

Sub StartWHCMB(shot)
    bMultiBallMode = True
    bWHCMBActive = True
    EnableBallSaver DMDStd(kDMDFet_WHCMB_SaveTimer)
    WHCMBScore = 0
    Dim balls : balls = 3
    'If shot = Lannister Then balls = 4   ' Changed for tmrBallLockCheck
    vpmTimer.AddTimer 1000, "StartNonBWMB "&balls&" '"
    If bUseFlexDMD Then
        Dim Scene
        Set Scene = NewSceneWithVideo("whcintro","got-whcintro")
        Scene.AddActor FlexDMD.NewImage("winter","got-whcwinter.png")
        Scene.GetImage("winter").SetAlignedPosition 86,16,FlexDMD_Align_Center
        Scene.AddActor FlexDMD.NewImage("has","got-whchas.png")
        With Scene.GetImage("has")
            .SetAlignedPosition 86,16,FlexDMD_Align_Center
            .Visible = 0
        End With
        Scene.AddActor FlexDMD.NewImage("come","got-whccome.png")
        With Scene.GetImage("come")
            .SetAlignedPosition 86,16,FlexDMD_Align_Center
            .Visible = 0
        End With
        DelayActor Scene.GetImage("winter"),1,false
        FlashActor Scene.GetImage("has"),1,1.5
        DelayActor Scene.GetImage("come"),2.5,true
        DMDEnqueueScene Scene,0,4000,4000,2000,"gotfx-whcintro"
    End If
  DOF 129, DOFOn
    vpmTimer.AddTimer 1500, "PlaySoundVol ""say-whcstart"&RndNbr(2)&""",VolCallout '"
    SetPlayfieldLights
    House(CurrentPlayer).SetUPFState true
    House(CurrentPlayer).ResetForWHCMB
    DMDCreateWHCMBScoreScene
    SetGameTimer tmrWinterHasCome,5
    tmrWiCLightning.Interval = 50
    tmrWiCLightning.Enabled = True
    PlayModeSong
    SetLockbarLight
End Sub

Sub CheckForHotkOrITReady
    Dim a,clr,readyscene
    bHotkMBReady = False : bITMBReady = False
    If CompletedHouses < 4 Or bMultiBallMode Or PlayerMode > 0 Then Exit Sub
    If bHotkMBActive = False And bHotkMBDone = False And bHotkMBReady = False Then
        bHotkMBReady = True
        clr = white
        readyscene = HotkMB
    Else
        If CompletedHouses < 7 Or bITMBActive Or bITMBReady Or bITMBDone Then Exit Sub
        bITMBReady = True
        readyscene = ITMB
        clr = cyan
    End If
    If Lockwall.Uservalue = 0 Then
        LockWall.Collidable = True : Lockwall.Uservalue = 1
        LockPost True
    End If
    DMDCreateReadyScene readyscene
    SavePlayfieldLightState
    House(CurrentPlayer).SetUPFState True
    ResetComboMultipliers
    SetModeLights
    TurnOffPlayfieldLights
    If bHotkMBReady Then SetLightColor li29,white,2 Else SetLightColor li35,cyan,2 : SetLightColor li29,cyan,1
    li108.BlinkPattern="1000"
    li111.BlinkPattern="0100"
    li114.BlinkPattern="0010"
    li117.BlinkPattern="0001"
    For each a in Array(li108,li111,li114,li117)
        SetLightColor a,clr,2
        a.blinkInterval = 100
    Next
    PlayModeSong
End Sub

Sub ResetHotkLights
    Dim a
    For each a in Array(li108,li111,li114,li117)
        a.BlinkPattern = "10"
    Next
    RestorePlayfieldLightState False
End Sub

Sub StartHotkMBIntro
    Dim scene,a
    bHotkIntroRunning = True    ' Tell sw48 to ignore the ball
    bBallReleasing = True
    ' Set up the HOTK intro lighting
    ResetHotkLights
    TurnOffPlayfieldLights
    GiEffect 3
    li29.BlinkInterval = 25
    SetLightColor li29,white,2
    For each a in Array(li17,li20,li23,li80,li83)
        a.BlinkInterval = 50
        SetLightColor a,white,2
    Next
    ' Do the intro scene.
    ' Start with "HAND OF THE KING" in 6x8 displayed for one second
    ' The rest is handled by tmrHotkIntro
    tmrHotkIntro.Interval = 1000
    tmrHotkIntro.Enabled = True
    tmrHotkIntro.UserValue = 0
    If bUseFlexDMD Then
        Set scene = FlexDMD.NewGroup("hotkintro")
        scene.AddActor FlexDMD.NewLabel("hotk",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),"HAND OF THE KING")
        scene.GetLabel("hotk").SetAlignedPosition 64,16,FlexDMD_Align_Center
        DMDFlush
        DMDEnqueueScene scene,0,1000,1500,1000,""
    Else
        DisplayDMDText "","HAND OF THE KING",1000
    End If
  DOF 130, DOFOn
End Sub

    ' Do the intro scene.
    ' Start with "HAND OF THE KING" in 6x8 displayed for one second. This is handled by StartHotkMBIntro
    ' Then switch to an offset 4 line display. "HAND OF THE KING" in top right, then house name 4 lines down on the left
    '  then what it brings to HOTK below on two lines. Each scene displays for ~2.4 seconds, with 0.25 seconds delay before
    '  adding line 2 and then lines 3/4.
    ' There are 4 scenes, one per house.
    ' Half way through the final scene, game says "You are now the hand of the king" in Kate's voice
    ' 3/4 way through final scene, held ball is released and other balls are added.
Dim HotkAbility
HotkAbility = Array("","BONUS  ROUND"&vbLf&"20  SECONDS  OF  FREE  SHOOTING","ALL  7  SHOTS"&vbLf&"INSTEAD  OF  4  SHOTS","100,000,000"&vbLf&"ADDED  TO  HURRY  UP", _
                "ONE  LESS  SET  NEEDED"&vbLf&"TO  START  HURRY  UP","+15,000,000"&vbLf&"PER  SUPER  JACKPOT","EACH  SHOT  NEEDS  TO"&vbLf&"BE  COMPLETED  TWICE", _
                "+500,000"&vbLf&"PER  SHOT  AWARD")
Sub tmrHotkIntro_Timer
    Dim scene,font,i,j
    tmrHotkIntro.Enabled = False
    If tmrHotkIntro.UserValue < 4 Then
        i = 0
        For j = 1 to 7
            If (House(CurrentPlayer).HotkMask And 2^j) > 0 Then
                If i = tmrHotkIntro.UserValue Then Exit For Else i = i + 1
            End If
        Next
        tmrHotkIntro.UserValue = tmrHotkIntro.UserValue + 1
        If tmrHotkIntro.UserValue = 4 Then tmrHotkIntro.Interval = 1200 Else tmrHotkIntro.Interval = 2400
        tmrHotkIntro.Enabled = True

        Set scene = FlexDMD.NewGroup("hotkintro")
        Set font = FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0)
        scene.AddActor FlexDMD.NewLabel("hotk1",font,"HAND  OF  THE  KING")
        scene.GetLabel("hotk1").SetAlignedPosition 127,0,FlexDMD_Align_TopRight

        scene.AddActor FlexDMD.NewLabel("hotk2",font,"HOUSE "&HouseToUCString(j))
        With scene.GetLabel("hotk2")
            .SetAlignedPosition 0,4,FlexDMD_Align_TopLeft
            .Visible = 0
        End With
        DelayActor scene.GetLabel("hotk2"),0.25,1

        if j > 7 Then msgbox "Oops. An error occurred. Please report details to Skillman. j="&j&"; CompletedHouses="&CompletedHouses&"; HotkMask="&House(CurrentPlayer).HotkMask : Exit Sub
        scene.AddActor FlexDMD.NewLabel("hotk3",font,HotkAbility(j))
        With scene.GetLabel("hotk3")
            .SetAlignedPosition 64,16,FlexDMD_Align_Top
            .Visible = 0
        End With
        DelayActor scene.GetLabel("hotk3"),0.5,1

        DMDFlush
        DMDEnqueueScene scene,0,2400,2500,500,"gotfx-hotkdrumintro"
    Else
        PlaySoundVol "say-hotk-start",VolCallout
        bHotkIntroRunning = False
        vpmTimer.AddTimer 500,"StartHotkMB '"
    End If
End Sub

Sub StartHotkMB
    bHotkMBActive = True
    bMultiBallMode = True
    bHotkMBDone = True
    bHotkMBReady = False
    HotkScore = 0
    EnableBallSaver DMDStd(kDMDFet_HOTKMB_SaveTimer)
    ReleaseLockedBall 1
    AddMultiball 3
    bInlanes(0) = True : bInlanes(1) = True
    House(CurrentPlayer).ResetForHotkMB
    SetPlayfieldLights
    PlayModeSong
    SetLockbarLight
    DMDCreateHotkMBScoreScene
End Sub

Sub StartIronThroneMB
    Dim scene
    bITMBActive = True
    bMultiBallMode = True
    bHotkIntroRunning = True ' Tell sw48 to ignore the ball for now
    bITMBReady = False
    bLockIsLit = False
    ResetDropTargets
    ResetHotkLights
    ITScore = 0
    EnableBallSaver DMDStd(kDMDFet_ITMB_SaveTimer)
    bBallReleasing = True
    vpmTimer.AddTimer 4000, "ReleaseLockedBall 1 : StartNonBWMB 1 : bHotkIntroRunning = False '"
    If bUseFlexDMD Then
        Set scene = NewSceneWithVideo("itintro","got-itintrobg")
        scene.AddActor NewSceneFromImageSequence("itfg","itintro",110,20,0,0)
        DMDEnqueueScene scene,0,5000,5000,2000,""
    End If
    House(CurrentPlayer).ResetForITMB
    SetPlayfieldLights
    PlayModeSong
    SetLockbarLight
    DMDCreateITMBScoreScene 0,0,7
  DOF 131, DOFOn
End Sub

' Midnight Madness: A 1 minute (minimum) 6-ball multiball that starts at midnight if game is in play
' - At midnight, turn off all lights, GI lights, flippers, and release all locked balls
' - Play 12 strikes, flashing "Midnight Madness" on each strike
' - After the 12th strike, launch all 6 balls
' - 45? second ball saver
' - every switch hit plays "hodor" or "stick em with the pointy end", as long as it's not already saying
' - every switch hit is worth 100k x playfieldX
' - battle timers frozen during MM

Sub StartMidnightMadnessMBIntro
    Dim scene
    If PlayerMode = 2 Then EndHurryUp
    StopGameTimers
    If PlayerMode < 0 Then
        ' In the middle of choosing a battle. Nip that in the bud
        'PlayerMode = 0
    SetPlayerMode 0
        HouseBattle1=0:HouseBattle2=0
        House(CurrentPlayer).BattleReady = True
        SetPlayfieldLights
        tmrStartBattleModeScene.Enabled = 0
    End if
    If bMysteryAwardActive Then
        IronThroneKickout
        bMysteryAwardActive = False
    End If
    tmrMidnightMadness.UserValue=0
    tmrMidnightMadness.Interval=3000
    tmrMidnightMadness.Enabled=1
    bMadnessMB = 1
    MMTotalScore = 0
    TurnOffPlayfieldLights
    GiOff
    GiIntensity 2
    ' Drop the lock walls to release any locked balls
    SwordWall.collidable = False : MoveSword True
    LockWall.collidable = False : Lockwall.Uservalue = 0 : LockPost False
    RealBallsInLock = 0

    debug.print "Timer: "&Timer()

    ' create "midnight madness" intro scene
    Set scene = FlexDMD.NewGroup("mmintro")
    scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("skinny10x12.fnt",vbWhite,vbWhite,0),"MIDNIGHT"&vbLf&"MADNESS")
    With scene.GetLabel("txt")
        .SetAlignedPosition 64,16,FlexDMD_Align_Center
        .Visible = 0
    End with
    DelayBlinkActor scene.GetLabel("txt"),3,0.5,12
    DMDBlank
    DMDEnqueueScene scene,0,15000,15000,1000,""
    If Song <> "" Then StopSound Song : Song = ""
    vpmTimer.AddTimer 15000,"StartMidnightMadnessMB '"
    FlipperDeactivate RightFlipper, RFPress : FlipperDeactivate RightUFlipper, URFPress : SolRFlipper 0
    FlipperDeactivate LeftFlipper, LFPress : FlipperDeactivate LeftUFlipper, ULFPress : SolLFlipper 0
End Sub

Sub StartMidnightMadnessMB
    bMadnessMB = 2
    bMultiBallMode = True
    AddMultiballFast 6
    EnableBallSaver DMDStd(kDMDFet_MMMB_SaveTimer)
    PlayModeSong
    DMDCreateMadnessMBScoreScene
    tmrMadnessMBToggle.Interval = 300
    tmrMadnessMBToggle.Enabled = 1
End Sub

Sub tmrMidnightMadness_Timer
    tmrMidnightMadness.UserValue = tmrMidnightMadness.UserValue + 1
    tmrMidnightMadness.Enabled=0
    tmrMidnightMadness.Interval=1000
    If tmrMidnightMadness.UserValue > 12 Then StartMidnightMadnessMB : Exit Sub
    tmrMidnightMadness.Enabled=1
    PlaySoundVol "gotfx-mmchime",VolDef
End Sub

Sub tmrMadnessMBToggle_Timer
    Dim e,o,l
    ' Timer that toggles playfield lights between two states
    ' after states are set, certain lights can be set to their real state: PF mul, battering ram, Shoot Again
    If tmrMadnessMBToggle.UserValue = 0 Then
        e=1:o=0:tmrMadnessMBToggle.UserValue = 1
    Else
        e=0:o=1:tmrMadnessMBToggle.UserValue = 0
    End if
    For each l in aPFLightsEven
        l.state = e
    Next
    For each l in aPFLightsOdd
        l.state = o
    Next
    If PlayfieldMultiplierVal > 1 Then SetPFMLights
    If PFMState > 0 Then SetBatteringRamLights
    if bBallSaverActive Then SetShootAgainLight
End Sub

Sub tmrMadnessSpeech_Timer
    tmrMadnessSpeech.Enabled = 0
    DMDUpdateMadnessScoreScene 1
End Sub

' Process hits to anything during Midnight Madness Multiball
Sub DoMadnessMBHit
    Dim i
    AddScore 100000
    MMTotalScore = MMTotalScore + (100000*PlayfieldMultiplierVal)
    GiEffect 6
    SetSpotFlashers 6
    SetUPFFlashers 6,white
    SetFlasher fl237,6 : SetFlasher fl242,3
    DMDUpdateMadnessScoreScene(100000*PlayfieldMultiplierVal)
End Sub


' Handle physically releasing a locked ball, as well as any sound effects needed
Sub ReleaseLockedBall(sword)
  debug.print "ReleaseLockedBall:" & sword
    SwordWall.TimerInterval = 500   ' how long the release solenoid stays down for
    SwordWall.TimerEnabled = True
    SwordWall.Collidable = True
    SwordWall.UserValue = sword
    LockWall.Collidable = False : Lockwall.Uservalue = 0
    'Move the actuator primitive down
    LockPost False
    'PlaySoundAt "fx_kicker",sw46
    SoundSaucerKick 0, sw46
    If sword Then
        'Rotate sword primitive to "chop off" the ball
        MoveSword False
        'TODO Play sword chopping sound
    End If
End Sub

Sub MoveSword(up)

    If up Then
    tmrMoveSword.UserValue = 5
  Else
    tmrMoveSword.UserValue = -1
    DOF 123,DOFPulse
  End If

    tmrMoveSword.Interval = 10
    tmrMoveSword.Enabled = 1

End Sub

Sub tmrMoveSword_Timer
    Dim X,Y,Z,s
    Dim tmp : tmp = me.UserValue
    if tmp >= 0 Then
        X = (tmp*2.4)
        Y = (tmp*5)
        Z = tmp
    Else
        X = tmp*10
        Y = tmp*-8.33
        Z = tmp*-1.66
        If Y > 20 And RealBallsInLock > 0 Then ' Sword has dropped far enough to hit remaining locked ball(s)
            ' Find the ball that's sitting against the sword wall
            Dim b,gBOT : gBOT = getballs
            For each b in gBOT
                If b.z > 90 Then
                    If InRect(b.x,b.y,840,1020,900,1038,914,985,853,967) Then b.velY = -12 : b.velX = 3.4 : debug.print "  Giving ball a push" : Exit For
                End If
            Next
        End if
    End If
    For each s in sword_003_BL : s.ObjRotY = Y : Next
    For each s in sword_002_BL : s.TransZ = X : s.TransX = Z : Next
    if tmp = 0 or tmp =< -3 then me.Enabled=0
    tmp = tmp-1
    me.UserValue=tmp
End Sub


Sub LockPost(up)
    Dim Z,lp
    If up then Z = 0 Else Z = -60
    For each lp in lockpost_001_BL : lp.TransZ = Z : Next
    DOF 123,DOFPulse
End Sub


' Called after the release solenoid has been down for ~500ms. Re-opens the invisible sword wall to let the
' next ball through
Sub SwordWall_Timer
    SwordWall.TimerEnabled = False
    If tmrBWmultiballRelease.Enabled = False Then bBallReleasing = False
    If RealBallsInLock > 0 or MultiballBallsLocked > 0 Or ((House(CurrentPlayer).bBattleReady Or bLockIsLit Or bHotkMBReady or bITMBReady) And Not bMultiBallMode) Then
        LockWall.Collidable = True : Lockwall.Uservalue = 1
        LockPost True
    End If
    SwordWall.Collidable = False
    If SwordWall.UserValue = 1 Then
        'Rotate sword primitive back to "open"
        MoveSword True
    End If
End Sub


Sub tmrBWmultiballRelease_Timer
    tmrBWmultiballRelease.Enabled = False
    tmrBWmultiballRelease.Interval = 1000
    ReleaseLockedBall 1
    MultiballBallsLocked = MultiballBallsLocked - 1
    If MultiballBallsLocked > 0 Then tmrBWmultiballRelease.Enabled = True
    If bBWMultiballActive = False Then Exit Sub ' Added for tmrBallLockCheck
    BallsInLock = BallsInLock - 1
    RealBallsInLock = RealBallsInLock - 1
    If RealBallsInLock > 0 Then tmrBWmultiballRelease.Enabled = True: Exit Sub
    If bBWMultiballActive Then
        If BallsInLock > 0 Then
            AddMultiball BallsInLock : BallsInLock = 0
        ElseIf bDrainHitWhileBallReleasing And BallsOnPlayfield-RealBallsInLock < 2 Then
            ' Multiball ended while the only remaining ball was being released - end multiball mode
            bMultiBallMode = False
            StopMBmodes
        End If
    End if
End Sub

Sub tmrGame_Timer
    Dim i
    GameTimeStamp = GameTimeStamp + 1
    If Timer() < 180 And bMadnessMB = 0 And bGameInPlay And bBallInPlungerLane = False And (PlayerMode=0 Or PlayerMode=1) And bMultiBallMode=False Then StartMidnightMadnessMBIntro
    if bGameTimersEnabled = False Then Exit Sub
    bGameTimersEnabled = False
    For i = 1 to MaxTimers
    'change for unfreezable timers 1.1.17
        'If ((TimerFlags(i) AND 2) = 2 Or (TimerFlags(i) And 4) = 4) And (i > 5 and i<>tmrTargaryenFreezeUpdateDMD) Then
        If ((TimerFlags(i) AND 2) = 2 Or (TimerFlags(i) And 4) = 4) And (i > MaxUnfreezableTimerIndex) Then
      TimerTimestamp(i) = TimerTimestamp(i) + 1 ' "Frozen" timer - increase its expiry by 1 step (Timers 1 - MaxUnfreezableTimerIndex (was 5) can't be frozen.)
    End If
        If (TimerFlags(i) AND 1) = 1 Then
            bGameTimersEnabled = True
            If TimerTimestamp(i) <= GameTimeStamp Then
                TimerFlags(i) = TimerFlags(i) AND 254   ' Set bit0 to 0: Disable timer
                Execute(TimerSubroutine(i))
            End If
        End if
    Next
End Sub

Sub SetGameTimer(tmr,val)
    TimerTimestamp(tmr) = GameTimeStamp + val
    TimerFlags(tmr) = TimerFlags(tmr) or 1
    bGameTimersEnabled = True
End Sub

Sub StopGameTimers
    Dim i
    For i = 1 to MaxTimers:TimerFlags(i) = TimerFlags(i) And 254: Next
End Sub

Sub ResetAllTimers
  debug.print "***Resetting all timers"
  dim i
  For i = 1 to MaxTimers
    if TimerTimestamp(i) > 0 then TimerTimestamp(i) = 0
    'TimerFlags(i) = TimerFlags(i) And 254
    Next
end sub

' Freeze all timers, including ball saver timers
' This should only be called when no balls are in play - i.e. the only "live" ball is held somewhere
Dim bAllGameTimersFrozen
Sub FreezeAllGameTimers
    Dim i
    For i = 1 to MaxTimers
        TimerFlags(i) = TimerFlags(i) Or 2
    Next
    bAllGameTimersFrozen = True
  DOF 136, DOFOn
End Sub

Sub ThawAllGameTimers
    Dim i
    ' If mode timers are running, set up a mode pause timer that will pause them 2 seconds from now
    ' This timer is reset every time this sub is called (i.e. every time there's a scoring event)
    If DMDStd(kDMDFet_InactivityPauseTimer) > 0 And ((TimerFlags(tmrBattleMode1) And 1) > 0 Or (TimerFlags(tmrBattleMode2) And 1) > 0) Then
        ' This just repeats the SetGameTimer sub, but avoids a sub call and hopefully an OutOfStackSpace error
        TimerTimestamp(tmrModePause) = GameTimeStamp + 20
        TimerFlags(tmrModePause) = TimerFlags(tmrModePause) or 1
        bGameTimersEnabled = True
    End if
  DOF 136, DOFOff
    If bAllGameTimersFrozen = False Then Exit Sub
    For i = 1 to MaxTimers
        TimerFlags(i) = TimerFlags(i) And 253
    Next
End Sub

Sub MartellBattleTimer
    Dim h
    If HouseBattle2 = Martell Then h = HouseBattle2 else h = HouseBattle1
    House(CurrentPlayer).BattleState(h).MartellTimer
End Sub

Sub BattleModeTimer1
    House(CurrentPlayer).BattleState(HouseBattle1).BattleTimerExpired
End Sub

Sub BattleModeTimer2
    House(CurrentPlayer).BattleState(HouseBattle2).BattleTimerExpired

End Sub

Sub ModePauseTimer
    If (TimerFlags(tmrBattleMode1) And 1) > 0 Then TimerFlags(tmrBattleMode1) = TimerFlags(tmrBattleMode1) Or 2
    If (TimerFlags(tmrBattleMode2) And 1) > 0 Then TimerFlags(tmrBattleMode2) = TimerFlags(tmrBattleMode2) Or 2
    If (TimerFlags(tmrMartellBattle) And 1) > 0 Then TimerFlags(tmrMartellBattle) = TimerFlags(tmrMartellBattle) Or 2
  DOF 136, DOFOn
End Sub

Sub BlackwaterSJPTimer
    TimerFlags(tmrBlackwaterSJP) = 0
    bBlackwaterSJPMode = False
    If bHotkMBActive Then House(CurrentPlayer).HotkTimerExpired
    SetBatteringRamLights
    If bBWMultiballActive Then House(CurrentPlayer).IncreaseBWJackpotLevel
End Sub

Sub UPFMultiplierTimer
    House(CurrentPlayer).ResetUPFMultiplier
End Sub

Sub tmrBattleCompleteScene_Timer
  'debug.print "tmrBattleCompleteScene interval: " & tmrBattleCompleteScene.interval
    tmrBattleCompleteScene.Enabled = 0
    If tmrBattleCompleteScene.UserValue > 0 Then House(CurrentPlayer).BattleState(tmrBattleCompleteScene.UserValue).DoBattleCompleteScene
    tmrBattleCompleteScene.UserValue = 0
End Sub

Sub tmrMultiballCompleteScene_Timer
    dim ex : ex = CStr(tmrMultiballCompleteScene.UserValue)
    tmrMultiballCompleteScene.Enabled = 0
    If ex <> "0" And ex <> "" Then Execute(ex)
    tmrMultiballCompleteScene.UserValue = ""
End Sub

Sub WHCTimer
    If bWHCMBActive Then SetGameTimer tmrWinterHasCome,5
    House(CurrentPlayer).CheckWHCTimer
End Sub

Sub LoLSaveTimer
    bLoLLit = False
    SetOutlaneLights
End Sub

Sub TargaryenFreezeTimer
    Dim i
    'change to MaxUnfreezableTimerIndex 1.1.17
    'For i = 6 to MaxTimers
  For i = (MaxUnfreezableTimerIndex+1) to MaxTimers
        TimerFlags(i) = TimerFlags(i) And 251
    Next
End Sub

Sub TargaryenFreezeAllTimers
    Dim i
  'change to MaxUnfreezableTimerIndex 1.1.17
    'For i = 6 to MaxTimers
  For i = (MaxUnfreezableTimerIndex+1) to MaxTimers
        If i <> tmrBallSave And i <> tmrBallSaveSpeedUp And i <> tmrTargaryenFreeze Then TimerFlags(i) = TimerFlags(i) Or 4
    Next
    SetGameTimer tmrTargaryenFreeze,DMDStd(kDMDFet_TargaryenFreezeTime)*10
End Sub

' Check the location of all balls on the playfield to see if any are stuck on the ramp
' that shouldn't be
Dim BallCheckFails
Sub tmrBallLockCheck_Timer
    Dim BallsOnRamp : BallsOnRamp = 0
    Dim BOT,b
    If PlayerMode < 0 Or bBallReleasing or tmrBWMultiBallRelease.Enabled Or bGameInPlay = False Then Exit Sub
    BOT = GetBalls
    For each b in BOT
        If b.z > 90 Then
            If b.x > 835 and b.y > 830 and b.y < 1090 Then BallsOnRamp = BallsOnRamp + 1
        End If
    Next
    If BallsOnRamp > RealBallsInLock Then
        If BallCheckFails >= 2 Then
            bBallReleasing = True
            ReleaseLockedBall 0
            BallCheckFails = 0
            debug.print "A BALL GOT STUCK ON RAMP. RELEASING"
        Else
            BallCheckFails = BallCheckFails + 1
        End If
    Else
        BallCheckFails = 0
    End If
End Sub


'*********************
' HurryUp Support
'*********************

' Called every 200ms by the GameTimer to update the HurryUp value
Sub HurryUpTimer
    Dim lbl
    If bHurryUpActive Then HurryUpCounter = HurryUpCounter + 1
    If bTGHurryUpActive Then TGHurryUpCounter = TGHurryUpCounter + 1
    If bUseFlexDMD Then
        If bHurryUpActive And Not IsEmpty(HurryUpScene) Then
            If Not (HurryUpScene is Nothing) Then
                Set lbl = HurryUpScene.GetLabel("HurryUp")
                If Not lbl is Nothing Then
                    FlexDMD.LockRenderThread
                    lbl.Text = FormatScore(HurryUpValue)
                    FlexDMD.UnlockRenderThread
                End If
            End if
        End If
        If bTGHurryUpActive And Not IsEmpty(TGHurryUpScene) Then
            If Not (TGHurryUpScene is Nothing) Then
                Set lbl = TGHurryUpScene.GetLabel("TGHurryUp")
                If Not lbl is Nothing Then
                    FlexDMD.LockRenderThread
                    lbl.Text = FormatScore(TGHurryUpValue)
                    FlexDMD.UnlockRenderThread
                End If
            End if
        End If
    End if
    if bHurryUpActive And HurryUpCounter > HurryUpGrace Then
        HurryUpValue = HurryUpValue - HurryUpChange
        If HurryUpValue <= 0 Then HurryUpValue = 0 : EndHurryUp
    End If
    if bTGHurryUpActive And TGHurryUpCounter > TGHurryUpGrace Then
        TGHurryUpValue = TGHurryUpValue - TGHurryUpChange
        If TGHurryUpValue <= 0 Then TGHurryUpValue = 0 : EndTGHurryUp
    End If
    If bTGHurryUpActive or bHurryUpActive Then SetGameTimer tmrHurryUp,2
End Sub

' Start a HurryUp
'  value: Starting value of HurryUp
'  scene: A FlexDMD scene containing a Label named "HurryUp". The text of the label will be
'         updated every HurryUp period (200ms)
'  grace: Grace period in 200ms ticks. Value will start declining after this many ticks have elapsed
'
' HurryUpChange value calculated by watching change value of numerous Hurry Ups on real GoT tables. Ratio of change to original value was always the same
' The real GoT table sometimes introduces variability to the change value (e.g  alternating between +10K and -10K from base value) but we're not
' going to bother
Sub StartHurryUp(value,scene,grace)
    if bHurryUpActive Then
        debug.print "HurryUp already active! Can't have two!"
        Exit Sub
    End If
    If bUseFlexDMD Then
        Set HurryUpScene = scene
    End If
    HurryUpGrace = grace
    HurryUpValue = value
    HurryUpCounter = 0
    HurryUpChange = Int(HurryUpValue / 1033.32) * 10
    bHurryUpActive = True
    SetGameTimer tmrHurryUp,2
End Sub

' Called when the HurryUp runs down. Ends battle mode if running
Sub EndHurryUp
    StopHurryUp
    If PlayerMode = 1 Then
        If HouseBattle2 > 0 Then House(CurrentPlayer).BattleState(HouseBattle2).BEndHurryUp
        If HouseBattle1 > 0 Then House(CurrentPlayer).BattleState(HouseBattle1).BEndHurryUp
    End if
    If bHotkMBActive Then House(CurrentPlayer).HotkHurryUpExpired
End Sub

' Called when the HurryUp has been scored. Ends the HurryUp but doesn't end the battle (if applicable)
Sub StopHurryUp
    Dim lbl
    bHurryUpActive = False
    if bTGHurryUpActive = False Then TimerFlags(tmrHurryUp) = TimerFlags(tmrHurryUp) And 254
    If PlayerMode = 2 Then
        'PlayerMode = 0
    SetPlayerMode 0
        House(CurrentPlayer).SetUPFState False
        StopSound "gotfx-long-wind-blowing"
        tmrWiCLightning.Enabled = False
    DOF 126, DOFOff
        GiIntensity 1
        SetPlayfieldLights
        SetTopGates
        PlayModeSong
        If bBWMultiballActive Then DMDCreateBWMBScoreScene : Else DMDResetScoreScene
        DMDFlush
        AddScore 0
    End If
    if bUseFlexDMD Then
        If Not IsEmpty(HurryUpScene) Then
            If Not (HurryUpScene is Nothing) Then
                Set lbl = HurryUpScene.GetLabel("HurryUp")
                If Not lbl is Nothing Then
                    FlexDMD.LockRenderThread
                    lbl.Visible = False
                    FlexDMD.UnlockRenderThread
                End if
            End If
        End if
    End if
End Sub

'Targaryen-specific HurryUp - allows it to co-exist with another HurryUp
Sub StartTGHurryUp(value,scene,grace)
    if bTGHurryUpActive Then
        debug.print "HurryUp already active! Can't have two!"
        Exit Sub
    End If
    If bUseFlexDMD Then
        Set TGHurryUpScene = scene
    End If
    TGHurryUpGrace = grace
    TGHurryUpValue = value
    TGHurryUpCounter = 0
    TGHurryUpChange = Int(TGHurryUpValue / 1033.32) * 10
    bTGHurryUpActive = True
    SetGameTimer tmrHurryUp,2
End Sub

' Called when the HurryUp runs down. Ends battle mode if running
Sub EndTGHurryUp
    StopTGHurryUp
    If PlayerMode = 1 Then
        If HouseBattle1 = Targaryen Then House(CurrentPlayer).BattleState(Targaryen).BEndHurryUp
        If HouseBattle2 = Targaryen Then House(CurrentPlayer).BattleState(Targaryen).BEndHurryUp
    End if
End Sub

' Called when the HurryUp has been scored. Ends the HurryUp but doesn't end the battle (if applicable)
Sub StopTGHurryUp
    Dim lbl
    bTGHurryUpActive = False
    if bHurryUpActive = False Then TimerFlags(tmrHurryUp) = TimerFlags(tmrHurryUp) And 254
    if bUseFlexDMD Then
        If Not IsEmpty(TGHurryUpScene) Then
            If Not (TGHurryUpScene is Nothing) Then
                Set lbl = TGHurryUpScene.GetLabel("TGHurryUp")
                If Not lbl is Nothing Then
                    FlexDMD.LockRenderThread
                    lbl.Visible = False
                    FlexDMD.UnlockRenderThread
                End if
            End If
        End if
    End if
End Sub

' Start the Winter Is Coming HurryUp
' All playfield lights turn off except flashing shot
' Hurry Up starts, wind blows, "lightning" flashes
Dim WICHurryUpScene
Sub StartWICHurryUp(value,shot)
    'PlayerMode = 2
  SetPlayerMode 2
    House(CurrentPlayer).SetUPFState True
    SetTopGates
    SetPlayfieldLights
    If bUseFlexDMD Then
        Set WICHurryUpScene = NewSceneWithVideo("wic","got-wichurryup")
        WICHurryUpScene.AddActor FlexDMD.NewLabel("HurryUp",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),"")
        WICHurryUpScene.GetLabel("HurryUp").SetAlignedPosition 4,21,FlexDMD_Align_TopLeft
    End if
    DMDFlush
    DMDEnqueueScene WICHurryUpScene,0,4000,6000,500,""
    DMDSetAlternateScoreScene WICHurryUpScene,16
    StartHurryUp value,WICHurryUpScene,10
    PlaySoundVol "gotfx-wicstart",VolDef
    PlaySoundVol "gotfx-long-wind-blowing",VolDef/2
    PlayModeSong
    tmrWiCLightning.Interval = 50
    tmrWiCLightning.Enabled = True
  DOF 126, DOFOn
End Sub

Sub tmrWiCLightning_Timer
    Dim step,i
    tmrWiCLightning.Enabled = False
    If (DMDStd(kDMDFet_ReduceWHCLamps)) Then Exit Sub
    step = tmrWiCLightning.UserValue
    If step > 7 then step = 0
    Select Case step
        Case 0,1,3: i = 0
        Case 2,5: i = 0.2
        Case 4: i = 1
        Case 6,7: i = 0.4
    End Select
    GiIntensity i
    step = step + 1
    tmrWiCLightning.UserValue = step
    tmrWiCLightning.Enabled = true
End Sub

Sub IncreaseWinterIsComing
    'TODO Handle Winter Is Coming increase (likely move inside House Class)
    ' Play sound & animation
End Sub

Sub IncreaseWallJackpot
    If WallJPValue < 1500000 Then WallJPValue = WallJPValue + 3000000 Else WallJPValue = WallJPValue + 500000 + RndNbr(100)*5000
    DMDPlayHitScene "got-increasewalljp","gotfx-increasewalljp",1,"BATTLE  FOR  WALL  INCREASES","VALUE="&FormatScore(WallJPValue),"",0,3
End Sub

Sub AdvanceWallMultiball(n)
    'Countdown to Wall Multiball
    ' Play rotating clock animation based on where we're at
    ' If we're at "Ready" then Set up for Wall Multiball
    Dim start,num,scene
    If bWallMBReady or bWallMBActive Then Exit Sub

    If WallMBCompleted > 0 Then Start=0 Else Start=220
    Start = Start + WallMBLevel*20
    If Start = 0 Then Start = 1
    num = n*20 + 1

    WallMBLevel = WallMBLevel + n

    If WallMBLevel >= 11 Or (WallMBCompleted = 0 And WallMBLevel >= 6 ) Then
        bWallMBReady = True
        DMDCreateReadyScene WallMB
        PlayModeSong
        setEBLight
        SetMysteryLight
        SetTopGates
        WallMBLevel = 0
    End If

    if bUseFlexDMD Then
        Set scene = NewSceneFromImageSequenceRange("clock","wallclock",start,num,25,0,1)
        DMDEnqueueScene scene,1,800*n+400,800*n+400,5000,"gotfx-wallcountdown"
    Else
        DisplayDMDText "WALL MULTIBALL","LEVEL "&WallMBLevel,1000
        PlaySoundVol,"gotfx-wallcountdown"
    End If
End Sub

Sub AwardSpecial
    DoAwardReplay
End Sub

Sub DoLordOfLight(saved)
    Dim Scene,line,pri
    pri = 2
    If bUseFlexDMD Then
        If saved then line = "KEEP   SHOOTING" : pri=0 : Else line = "LORD   OF   LIGHT"&vbLf&"OUTLANE   BALL-SAVE   LIT"
        Set Scene = NewSceneWithVideo("lol","got-keepshooting")
        Scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0),line)
        Scene.GetLabel("txt").SetAlignedPosition 40,8,FlexDMD_Align_Center
        If saved then BlinkActor Scene.GetLabel("txt"),100,15
        DMDEnqueueScene Scene,pri,1000,1500,1500,"gotfx-lolsave"
    End If
    bLolLit = Not Saved
    SetOutlaneLights
End Sub

Sub DoAwardSword
    Dim scene,font,i,j
    bSwordLit = False : SetSwordLight
    If SwordsCollected = 8 then Exit Sub    ' SHOULD never happen
    PlaySoundVol "gotfx-swordaward",VolDef
    SwordsCollected = SwordsCollected + 1
    Do
        i = RndNbr(8)-1
    Loop While (SwordMask And (2^i)) > 0
    SwordMask = SwordMask Or (2^i)
    j = RndNbr(1000)*10000+500000
    AddScore j
    If bUseFlexDMD Then
        Set scene = NewSceneWithVideo("sword","got-swordawarded")
        Set font = FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0)
        scene.AddActor FlexDMD.NewLabel("swname",font,SwordNames(i))
        scene.AddActor FlexDMD.NewLabel("score",font,FormatScore(j))
        scene.GetLabel("swname").SetAlignedPosition 120,2,FlexDMD_Align_TopRight
        scene.GetLabel("score").SetAlignedPosition 80,26,FlexDMD_Align_Center
        DMDEnqueueScene scene,0,1500,2500,1000,"say-sword"&i
    Else
        PlaySoundVol "say-sword"&i,VolDef
        DisplayDMDText SwordNames(i),FormatScore(j),1000
    End if
End Sub

Sub DoSwordsAreLit
    Dim scene
    bSwordLit = True : SetSwordLight
    If bUseFlexDMD Then
        Set scene = NewSceneWithVideo("swordlit","got-swordslit")
        scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f6by8.fnt",vbwhite,vbwhite,0),"SWORDS"&vbLf&"ARE LIT")
        scene.GetLabel("txt").SetAlignedPosition 96,9,FlexDMD_Align_Center
        DMDEnqueueScene scene,1,1500,2500,2000,"gotfx-swordsarelit"
    End If
End Sub

Sub DoWildfireLit
    Dim scene
    bWildfireLit = True : SetLightColor li126, darkgreen, 1
    If bUseFlexDMD Then
        Set scene = NewSceneWithImage("wflit","got-wildfirelit")
        scene.GetImage("wflitimg").SetAlignedPosition 0,-58,FlexDMD_Align_TopLeft
        scene.GetImage("wflitimg").AddAction scene.GetImage("wflitimg").ActionFactory().MoveTo(0,0,1)
        scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f6by8.fnt",vbwhite,vbwhite,0),"WILDFIRE"&vbLf&"IS LIT")
        scene.GetLabel("txt").SetAlignedPosition 102,9,FlexDMD_Align_Center
        DMDEnqueueScene scene,1,2000,3000,2000,""
    End If
End Sub

Sub DoEBisLit
    Dim scene,i
    EBisLit = EBisLit+1 : SetEBLight
    i = RndNbr(2)
    PlaySoundVol "say-extra-ball-is-lit"&i,VolCallout
    If bUseFlexDMD Then
        Set scene = NewSceneWithVideo("ebislit","got-ebislit")
        DMDEnqueueScene scene,1,1400,1400,4000,"gotfx-ebislit"
    Else
        DisplayDMDText "EXTRA BALL","IS LIT",1000
    End If
End Sub

Sub DoAwardExtraBall
    Dim scene
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + EBisLit
    EBisLit = 0
    SetShootAgainLight
    AddScore 15000000
    PlaySoundVol "gotfx-extra-ball1",VolDef
  DOF 133, DOFPulse
    If bUseFlexDMD Then
        Set scene = NewSceneWithVideo("eb","got-extraball")
        DMDEnqueueScene scene,0,5500,7500,3000,"gotfx-extra-ball"
    Else
        DisplayDMDText "","EXTRA BALL",1000
    End if
End Sub

Sub DoAwardReplay
    If DMDStd(kDMDStd_ReplayType) = 2 Then
        Dim oldEBisLit : oldEBisLit = EBisLit
        EBisLit=1
        DoAwardExtraBall
        EBisLit = oldEBisLit
        Exit Sub
    End if
    Dim scene
    'PlaySound SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
    vpmTimer.AddTimer 3000, "DOF 122,DOFPulse : KnockerSolenoid '"
    Credits = Credits + 1
    If bUseFlexDMD Then
        Set scene = NewSceneWithVideo("replay","got-replay")
        DMDEnqueueScene scene,0,2500,4000,4000,"gotfx-replay"
    End if
End Sub

Sub DMDDoBallSaved
    Dim scene,i
    If bUseFlexDMD Then
        Set scene = NewSceneWithVideo("ballsave","got-ballsaved")
        scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("skinny10x12.fnt",vbWhite,vbWhite,0),"BALL SAVED")
        scene.GetLabel("txt").SetAlignedPosition 64,16,FlexDMD_Align_Center
        DelayActor scene.GetLabel("txt"),1.5,true
        BlinkActor scene.GetLabel("txt"),0.1,10
        DMDEnqueueScene scene,0,2500,3500,2000,"gotfx-dragonroar"&RndNbr(3)
    Else
        DisplayDMDText "","BALL SAVED",2000
    End If
    i = RndNbr(9)
    vpmTimer.addTimer 1000,"PlaySoundVol ""say-ballsaved"&i&""",VolCallout '"
End Sub

Sub DMDDoWiCScene(value,shots)
    Dim scene,line3,i,left
    left = 3 - shots : If SelectedHouse = Greyjoy then left = left + 1
    If bUseFlexDMD Then
        Set scene = NewSceneWithVideo("wic","got-winterstorm")
        If left = 1 Then line3 = "1 MORE SHOT" Else line3 = left & " MORE SHOTS"
        scene.AddActor FlexDMD.NewLabel("txt1",FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0),"WINTER   IS   COMING"&vbLf&"VALUE   GROWS")
        scene.AddActor FlexDMD.NewLabel("txt2",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),FormatScore(value))
        scene.AddActor FlexDMD.NewLabel("txt3",FlexDMD.NewFont("tiny3by5.fnt",vbWhite,vbWhite,0),line3)
        scene.GetLabel("txt1").SetAlignedPosition 64,0,FlexDMD_Align_Top
        scene.GetLabel("txt2").SetAlignedPosition 64,16,FlexDMD_Align_Top
        scene.GetLabel("txt3").SetAlignedPosition 64,25,FlexDMD_Align_Top
        DMDEnqueueScene scene,1,1000,1200,2000,"gotfx-wind-blowing"
    Else
        DisplayDMDText "WIC VALUE GROWS",value,1000
        PlaySoundVol "gotfx-wind-blowing",VolDef
    End if
    If left = 1 then
        i = RndNbr(4)
        PlaySoundVol "say-winter-is-coming"&i,VolCallout
    End If
End Sub

Sub DoAddABall
    AddMultiballFast 1
    PlaySoundVol "say-add-a-ball",VolCallout
    DMD "","ADD A BALL","",eNone,eBlink,eNone,1000,True,""
    If bBallSaverActive And (TimerFlags(tmrBallSave) And 1) = 1 Then
        ' Main ball timer still active, just add time to it
        TimerTimestamp(tmrBallSave) = TimerTimestamp(tmrBallSave) + 70
        TimerTimestamp(tmrBallSaveSpeedUp) = TimerTimestamp(tmrBallSaveSpeedUp) + 70
    Else
        ' Already done or in the grace period: start a new ball server
        EnableBallSaver 7
    End If
End Sub

Sub DoBuyMultiplierScene(gold,remain)
    Dim Scene,font
    If bUseFlexDMD Then
        Set Scene = FlexDMD.NewGroup("xbuy")
        Set font = FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0)
        Scene.AddActor FlexDMD.NewLabel("txt",font,"BUY  MULTIPLIER  FOR  "&gold&"  GOLD")
        Scene.GetLabel("txt").SetAlignedPosition 64,1,FlexDMD_Align_Top
        Scene.AddActor FlexDMD.NewLabel("txt2",font,"NEW  GOLD  TOTAL: "&CurrentGold)
        Scene.GetLabel("txt2").SetAlignedPosition 0,31,FlexDMD_Align_BottomLeft
        Scene.AddActor FlexDMD.NewLabel("txt3",font,remain & "  MORE")
        Scene.GetLabel("txt3").SetAlignedPosition 127,31,FlexDMD_Align_BottomRight
        Scene.AddActor FlexDMD.NewLabel("mult",FlexDMD.NewFont("udmd-f7by10.fnt",vbWhite,vbWhite,0),"MULTIPLIER: "&PlayfieldMultiplierVal&"X")
        Scene.GetLabel("mult").SetAlignedPosition 64,16,FlexDMD_Align_Center
        DMDEnqueueScene Scene,0,1000,2000,1500,""
    Else
        DisplayDMDText remain&" MORE. NEW GOLD: " & CurrentGold,PlayfieldMultiplierVal & "X",1000
    End If
End Sub


'********************************
' Support for Battle (Mode) Start
'********************************

' StartChooseBattle: Called when BattleReady is lit and ball is shot up left ramp
'   - Create the array of battle choices
'   - Create the Instructions and Choose Battle scenes
'   - Play the Choose Battle song
'   - Set a timer to display the Choose Battle scene

Dim BattleChoices(49)   ' Max possible number of choices
Dim TotalBattleChoices,CurrentBattleChoice,bReadyForFlips
Sub StartChooseBattle
    Dim i,j,tmrval,done
    HouseBattle1 = Empty
    HouseBattle2 = Empty
    CBScene = Empty

    'PlayerMode = -2
  SetPlayerMode -2

    TurnOffPlayfieldLights
    SetLockbarLight

    DMDChooseBattleScene "","","",10

    ' Check to see if there are any unlit houses. If not, "Pass For Now" is not allowed
    done = True
    For i = 1 to 7
        If House(CurrentPlayer).Qualified(i) = False And House(CurrentPlayer).Completed(i) = False  Then done = False
    Next
    if done = False Then
        BattleChoices(0) = 0:TotalBattleChoices = 1 ' Pass For Now is allowed
    Else
        TotalBattleChoices = 0
    End if

    ' Create the array of choices
    For i = 0 to 7
        If (SelectedHouse <> Greyjoy And House(CurrentPlayer).Qualified(i) And House(CurrentPlayer).Completed(i) = False) or i = 0 then   ' Greyjoy can't stack house battles
            For j = 1 to 7
                If House(CurrentPlayer).Qualified(j) And House(CurrentPlayer).Completed(j) = False And j<>i Then
                    If i=0 Then  BattleChoices(TotalBattleChoices) = j*8 Else BattleChoices(TotalBattleChoices) = i*8+j
                    TotalBattleChoices = TotalBattleChoices + 1
                End If
            Next
        End If
    Next
    If TotalBattleChoices > 1 Then CurrentBattleChoice = 1 Else CurrentBattleChoice = 0

    ' Set up the launch timer
    SetGameTimer tmrChooseBattle,50

    ' Disable BattleReady for next time
    House(CurrentPlayer).BattleReady = False

    If bBattleInstructionsDone Then
        UpdateChooseBattle
    Else
        ' Set up the update timer to update after instructions have been displayed for 1.5 seconds
        tmrBattleInstructions.Interval = 1500
        tmrBattleInstructions.Enabled = 1
    End If
    PlayModeSong
    bReadyForFlips=true
  SetReadyForAction True
End Sub

Sub tmrBattleInstructions_Timer
    tmrBattleInstructions.Enabled = 0
    DMDFlush
    UpdateChooseBattle
    bBattleInstructionsDone = True
End Sub

' UpdateChooseBattle
' Set House string values based on the currently selected BattleChoice
' Update timers
' Update DMD

Sub UpdateChooseBattle
    Dim house1, house2, tmr, i

    ' Just in case we got called by accident
    If PlayerMode <> -2 Then
        TimerFlags(tmrUpdateChooseBattle) = 0
        Exit Sub
    End if

    ' Enable the game timer to call this sub again in 1 second
    SetGameTimer tmrUpdateChooseBattle,10

    If IsEmpty(CBScene) Then Exit Sub
    Set DefaultScene = CBScene

    HouseBattle2 = BattleChoices(CurrentBattleChoice) MOD 8
    HouseBattle1 = Int(BattleChoices(CurrentBattleChoice)/8)
    house2 = ""
    If HouseBattle1 = 0 Then
        house1 = "PASS FOR NOW"
    Else
        house1 = "HOUSE "&HouseToString(HouseBattle1)
        house2 = HouseToString(HouseBattle2)
        If house2 <> "" Then house2 = "HOUSE "&house2
    End If

    ' Flash the shield(s) of the currently selected house(s)
    For i = 1 to 7
        If i = HouseBattle1 or i = HouseBattle2 Then
            SetLightColor HouseSigil(i),HouseColor(i),2
        Else
            HouseSigil(i).State = 0
        End If
    Next
    tmr = Int( (TimerTimestamp(tmrChooseBattle)-GameTimeStamp) / 10) + 1
    DMDChooseBattleScene "CHOOSE YOUR BATTLE", house1, house2, tmr
    SetLockbarLight
End Sub

' PreLaunchBattleMode
' Shut off ChooseBattle timers
' Play animation with sound for housebattle1 & 2. Show objective for housebattle1 for length of sound
' Start battle:
'   Set mode timer(s)
' Check if we locked a ball and if so, do lock ball processing

Sub PreLaunchBattleMode
    Dim scene,tmr
    TimerFlags(tmrUpdateChooseBattle) = 0
    TimerFlags(tmrChooseBattle) = 0
    If BattleChoices(CurrentBattleChoice) = 0 Then ' Pass for now
        'PlayerMode = 0
    SetPlayerMode 0
        AddScore 0
        PlaySoundVol "gotfx-passfornow",VolDef
        PlayModeSong
        SetPlayfieldLights
        SetLockbarLight
        DMDFlush
        LaunchBattleMode
        Exit Sub
    End If

    HouseBattle2 = BattleChoices(CurrentBattleChoice) MOD 8
    HouseBattle1 = Int(BattleChoices(CurrentBattleChoice)/8)

    ' Start battle!
    tmrStartBattleModeScene.UserValue=0
    DMDHouseBattleScene HouseBattle1

    tmr = Int(SceneSoundLengths(HouseBattle1)/100)
    If HouseBattle2 > 0 Then tmr=tmr + Int(SceneSoundLengths(HouseBattle2)/100)

    'PlayerMode = -2.1
  SetPlayerMode -2.1

    House(CurrentPlayer).BattleState(HouseBattle1).StartBattleMode
    If HouseBattle2 > 0 Then House(CurrentPlayer).BattleState(HouseBattle2).StartBattleMode
    House(CurrentPlayer).SetUPFState True
    SetPlayfieldLights
    SetLockbarLight

    PlayModeSong

    DMDCreateAlternateScoreScene HouseBattle1,HouseBattle2

    SetGameTimer tmrLaunchBattle,tmr
End Sub

Sub tmrStartBattleModeScene_Timer
    tmrStartBattleModeScene.Enabled = 0
    If Tilted Then Exit Sub
    Select Case tmrStartBattleModeScene.UserValue
        Case 0
            DMDFlush
            StopSound "gotfx-"&HouseToString(HouseBattle1)&"battleintro"
            DMDHouseBattleScene HouseBattle2
        Case 1
            DMDFlush
            If HouseBattle2 > 0 Then StopSound "gotfx-"&HouseToString(HouseBattle2)&"battleintro"
            LaunchBattleMode
    End Select
    tmrStartBattleModeScene.UserValue = tmrStartBattleModeScene.UserValue + 1
End Sub

' Handle releasing or locking the ball after choosing battle
Sub LaunchBattleMode
    TimerFlags(tmrLaunchBattle) = 0
    tmrStartBattleModeScene.Enabled = 0

    PlaySoundAt "Drop_Target_Down_" & RndNbr(6), Target90
    DTDrop 4 : TargetState(4) = 1
    If PlayerMode = -2.1 Then
    'PlayerMode = 1
    SetPlayerMode 1
  end if

    ' Start the targaryen HurryUp if appropriate
    If HouseBattle1 = Targaryen Then
        House(CurrentPlayer).BattleState(HouseBattle1).TGStartHurryUp
        StartDragonWings True
    ElseIf HouseBattle2 = Targaryen Then
        House(CurrentPlayer).BattleState(HouseBattle2).TGStartHurryUp
        StartDragonWings True
    End If
    If bBattleCreateBall Then   'LockBall has already run. Create the new ball now
        bBattleCreateBall = False
        If RealBallsInLock > BallsInLock Then
            RealBallsInLock = RealBallsInLock - 1
            ReleaseLockedBall 0
        Else
            bAutoPlunger = True
            bPlayfieldValidated = False
            CreateNewBall
        End If
    ElseIf bLockIsLit Then      ' Should be 3rd ball locked. Tweaked timing to ensure he doesn't speak over top of song intro
        LockBall
    Else                            ' Lock isn't lit but we have a ball locked
        ReleaseLockedBall 0
    End If

  SetReadyForAction True
End Sub

' Called twice per second during battle mode
Sub UpdateBattleMode
    SetGameTimer tmrUpdateBattleMode,5
    DMDLocalScore
End Sub


'*************************
' Table lighting control
'*************************

Sub SetWildfireLights
    SetLightColor li80,green,ABS(bWildfireTargets(0))
    SetLightColor li83,green,ABS(bWildfireTargets(1))
End Sub

Sub SetTopLaneLights
    If bTopLanes(0) And bTopLanes(1) Then ' Flash the lights for a second and use a timer to shut them off
        li162.BlinkInterval=100:li165.blinkInterval=100
        li162.State=2:li165.State=2
        li162.TimerInterval=1000
        li162.TimerEnabled=1
    Else
        SetLightColor li162, white, ABS(bTopLanes(0))
        SetLightColor li165, white, ABS(bTopLanes(1))
    End if
End Sub

Sub SetTargetLights
    Dim i
    For i = 0 to 2
        if i >= LoLTargetsCompleted Then LoLLights(i).State = 0 Else SetLightColor LoLLights(i),yellow,1
    Next
End Sub

Sub SetOutlaneLights
    SetLightColor li11,white,ABS(bLoLLit)
    SetLightColor li74,white,ABS(bLoLLit)
End Sub

Sub SetInlaneLights
    If bInlanes(0) Then SetLightColor li14,yellow,2 Else SetLightColor li14,yellow,0
    If bInlanes(1) Then SetLightColor li71,yellow,2 Else SetLightColor li71,yellow,0
End Sub

Sub li162_Timer
    li162.State = 0
    li165.State = 0
    me.TimerEnabled=0
End Sub

Sub SetLockLight
    If bLockIsLit And bMultiBallMode = False Then
        ' Flash the lock light
        li111.BlinkInterval = 300
        SetLightColor li111,darkgreen,2
    Else
        SetLightColor li111,darkgreen,0
    End If
End Sub

Sub SetSwordLight
    If bSwordLit Then
        li138.BlinkPattern = 110
        li138.BlinkInterval = 150
        SetLightColor li138,yellow,2
    Else
        SetLightColor li138,yellow,0
    End If
End Sub

Sub SetMysteryLight
    If bWallMBReady And bITMBActive = False And bMultiBallMode = False Then
        SetLightColor li153,purple,2
        li153.blinkpattern="010"
        li153.blinkinterval = 100
    Elseif bMysteryLit And PlayerMode = 0 And bMultiBallMode = False And bITMBActive = False Then
        li153.BlinkInterval = 250
        SetLightColor li153,white,2
    Else
        SetLightColor li153,white,0
    End If
End Sub

Sub SetGoldTargetLights
    Dim i
    For i=0 to 4
        SetLightColor GoldTargetLights(i),yellow,ABS(bGoldTargets(i))
    Next
End Sub

' We'll use the timer on each lit combo light to gradually speed up the flashing
Sub SetComboLights
    Dim i,stat,x
    stat = 0:x=False
    For i = 1 to 5
        If bWallMBReady And bITMBActive = False And i=5 Then
            SetLightColor ComboLights(5),purple,2
        Else
            If ComboMultiplier(i) > 1 Then stat=2:x=True Else stat=0
            If i <> 3 Or (bHotkMBReady = False And bITMBReady = False) Then SetLightColor ComboLights(i),red,stat
        End If
    Next
    If x Then
        ComboLights(1).TimerEnabled = True
        ComboLights(1).TimerInterval = 1000
    Else
        ComboLights(1).TimerEnabled = False
    End if
End Sub

Sub setEBLight
    If bWallMBReady And bITMBActive = False And bMultiBallMode = False Then
        SetLightColor li150,purple,2
        li150.blinkpattern = "100"
        li150.blinkinterval = 100
        SetLightColor ComboLights(5),purple,2
        With ComboLights(5)
            .blinkpattern = "001"
            .blinkinterval = 100
        End With
    ElseIf EBisLit > 0 Then
        SetLightColor li150,amber,2
    Else
        SetLightColor li150,white,0
    End If
End Sub

Sub SetCastleBlackLight
    If PlayerMode >=0 And PlayerMode < 2 And bMultiBallMode = False And bITMBActive = False Then
        SetLightColor li95,purple,1
    Else
        li95.state = 0
    End If
End Sub

Sub SetBumperLights
    Dim a,clr
    If PlayerMode = 2 Then clr = ice Else clr = white
    For Each a in Array(li168,li171,li174)
        SetLightColor a,clr,1
    Next
End Sub

Sub SetBattleReadyLight
    If House(CurrentPlayer).bBattleReady And bMultiBallMode=False And PlayerMode=0 And bHotkMBReady=False And bITMBReady=False Then
        SetLightColor li108,white,2
    Else
        li108.state = 0
    End if
End Sub

Sub SetShootAgainLight
    If bBallSaverActive And (TimerFlags(tmrBallSave) And 1) = 1 Then
        SetLightColor LightShootAgain,white,2
    ElseIf ExtraBallsAwards(CurrentPlayer) > 0 Then
        SetLightColor LightShootAgain,white,1
    Else
        LightShootAgain.State = 0
    End If
End Sub

' Set up all of the default colours for playfield lights
Sub SetDefaultPlayfieldLights
    Dim i
    For i = 1 to 5: SetLightColor ComboLights(i),red,-1: Next
    For i = 0 to 4: SetLightColor GoldTargetLights(i),yellow,-1: Next
    For i = 0 to 2: SetLightColor LoLLights(i),yellow,-1: Next
    For i = 0 to 3: SetLightColor pfmuxlights(i),amber,-1: Next
    SetLightColor li95,purple,-1    ' Castle Black light (Dragon Shot)
    SetLightColor li80,green,-1     ' WF targets
    SetLightColor li83,green,-1     ' WF targets
    SetLightColor li150,amber,-1    ' EB light
    SetLightColor li153,white,-1    ' Mystery light
    SetLightColor li138,yellow,-1   ' Sword light
    SetLightColor li111,darkgreen,-1 ' Lock light
    SetLightColor li11,white,-1     ' Outlane lights
    SetLightColor li74,white,-1     ' Outlane lights
    SetLightColor li162,white,-1    ' Toplane lights
    SetLightColor li165,white,-1    ' Toplane lights
    SetLightColor li132,midblue,-1  ' Battering Ram
    SetLightColor li129,red,-1      ' SuperJackpot Light
    SetLightColor li126,darkgreen,-1    ' Wildfire light
    SetLightColor li123,white,-1    ' Playfield Mul light
    For i = 1 to 7
        SetLightColor HouseShield(i),HouseColor(i),-1
        SetLightColor HouseSigil(i),HouseColor(i),-1
    Next
    For i = 1 to 13 : SetLightColor UPFLights(i),cyan,-1 : Next

End Sub

' flash all lights on the playfield a single colour
Sub FlashPlayfieldLights(color,blinks,interval)
    If LightSeqPlayfield.UserValue > 0 Then Exit Sub ' Already running a sequence
    SavePlayfieldLightState
    PlayfieldSlowFade color,2
    LightSeqPlayfield.UserValue = 1
    LightSeqPlayfield.UpdateInterval = 2
    LightSeqPlayfield.Play SeqBlinking,,blinks,interval
    LightSeqPlayfield.Play SeqAllOff
End Sub

' ******************************
' ** Various custom light sequences
' ******************************

Sub DoBWJackpotSeq
    If LightSeqPlayfield.UserValue > 0 Then
        LightSeqPlayfield.StopPlay
        LightSeqPlayfield_PlayDone
    End If
    SavePlayfieldLightState
    PlayfieldSlowFade yellowgreen,2
    LightSeqPlayfield.UserValue = 1
    LightSeqPlayfield.UpdateInterval = 3
    LightSeqPlayfield.Play SeqDownOn, 25, 4
    LightSeqPlayfield.Play SeqAllOff
    SetSpotFlashers 4
    SetUPFFlashers 4,yellowgreen
    GiEffect 3
End Sub

Sub DoBWSJPSeq(fast)
    If LightSeqPlayfield.UserValue > 0 Then
        LightSeqPlayfield.StopPlay
        LightSeqPlayfield_PlayDone
    End If
    If fast Then FlashPlayfieldLights yellowgreen,25,16 : Else FlashPlayfieldLights yellowgreen,12,50
End Sub

Sub DoWallMBSeq
    If LightSeqPlayfield.UserValue > 0 Then Exit Sub ' Already running a sequence
    SavePlayfieldLightState
    PlayfieldSlowFade -1,0.02
    LightSeqPlayfield.UserValue = 1
    LightSeqPlayfield.UpdateInterval = 10
    LightSeqPlayfield.Play SeqBlinking,,4,250
    LightSeqPlayfield.Play SeqAllOff
End Sub

Sub DoWallMBJackpotSeq
    If LightSeqPlayfield.UserValue > 0 Then Exit Sub ' Already running a sequence
    SavePlayfieldLightState
    PlayfieldSlowFade white,5
    LightSeqPlayfield.UserValue = 1
    LightSeqPlayfield.UpdateInterval = 8
    If RndNbr(2) = 2 Then LightSeqPlayfield.Play SeqRightOn, 33, 1 : Else LightSeqPlayfield.Play SeqLeftOn, 33, 1
    LightSeqPlayfield.Play SeqAllOff
End Sub

Sub DoITMBJackpotSeq
    If LightSeqPlayfield.UserValue > 0 Then Exit Sub ' Already running a sequence
    SavePlayfieldLightState
    PlayfieldSlowFade white,5
    LightSeqPlayfield.UserValue = 1
    LightSeqPlayfield.UpdateInterval = 8
    LightSeqPlayfield.Play SeqLeftOn, 33, 2
    LightSeqPlayfield.Play SeqAllOff
    LightSeqPlayfield.UpdateInterval = 1000
    LightSeqPlayfield.Play SeqAllOn
    LightSeqPlayfield.UpdateInterval = 1
    LightSeqPlayfield.Play SeqAllOff
    SetSpotFlashers 4
End Sub

Sub DoLockBallSeq
    If LightSeqPlayfield.UserValue > 0 Then Exit Sub ' Already running a sequence
    SavePlayfieldLightState
    PlayfieldSlowFade green,0.1
    LightSeqPlayfield.UserValue = 1
    LightSeqPlayfield.UpdateInterval = 10
    LightSeqPlayfield.Play SeqUpOn, 25, 1
    LightSeqPlayfield.Play SeqBlinking,,2,250
    LightSeqPlayfield.Play SeqAllOff
End Sub

' Do the light sequence for the start of BW Multiball
Sub DoBWMultiballSeq
    If LightSeqPlayfield.UserValue > 0 Then Exit Sub ' Already running a sequence
    SavePlayfieldLightState
    PlayfieldSlowFade yellowgreen,0.1
    LightSeqPlayfield.UserValue = 1
    LightSeqPlayfield.UpdateInterval = 10
    LightSeqPlayfield.Play SeqBlinking,,2,250
    LightSeqPlayfield.Play SeqDownOn,75,2
    LightSeqPlayfield.Play SeqAllOff
    LightSeqPlayfield.UpdateInterval = 25
    LightSeqPlayfield.Play SeqBlinking,,60,25

    tmrMultiballSequencer.Interval = 2500  ' Tune this value to ensure it starts just after the wave sequence ends
    tmrMultiballSequencer.UserValue = 0
    tmrMultiballSequencer.Enabled = True

  DOF 127, DOFOn
End Sub

' Lights in fire sequence start from the UPF and move to the centre of the sigils
' sequence: li198,li201,li204,li207,li210,li101,li98,li95,li32,li41,(li38,li47,li50,li44),(li29,li62,li56,li53,li59,li65,li35)
Dim FireLights,FireLightPattern
FireLights = Array(li198,li201,li204,li207,li210,li101,li98,li95,li32,li41,li38,li47,li50,li44,li29,li62,li56,li53,li59,li65,li35)
FireLightPattern = Array("111100000000000000","011110000000000000","001111000000000000","000111100000000000","000011110000000000","000001111000000000","000000111100000000", _
                    "000000011110000000","000000001111111111","000000000111111100","000000000011111100","000000000011111100","000000000011111100","000000000011111100",_
                    "000000000001111111","000000000001111111","000000000001111111","000000000001111111","000000000001111111","000000000001111111","000000000001111111")
Sub DoTargaryenFireLightSeq
    Dim i
    If LightSeqPlayfield.UserValue > 0 Then
        LightSeqPlayfield.StopPlay
        LightSeqPlayfield_PlayDone
    End If
    LightSeqPlayfield.UserValue = 1
    SavePlayfieldLightState
    TurnOffPlayfieldLights
    PlayfieldSlowFade orange,0.5
    For i = 0 to UBound(FireLights)
        FireLights(i).blinkinterval = 50
        FireLights(i).blinkpattern = FireLightPattern(i)
        FireLights(i).State = 2
    Next
    tmrFireSeq.Interval = 900
    tmrFireSeq.Enabled = 1
End Sub

Sub tmrFireSeq_Timer
    Dim i
    tmrFireSeq.Enabled = 0
    For i = 0 to UBound(FireLights)
        FireLights(i).blinkinterval = 100
        FireLights(i).blinkpattern = "10"
    Next
    LightSeqPlayfield.UserValue = 0
    RestorePlayfieldLightState True
End Sub

' During multiball light sequence, randomly change the playfield colours
' between green and red during the fast random flash sequence
Sub tmrMultiballSequencer_Timer
    Dim a,c,i
    tmrMultiballSequencer.Enabled = False
    tmrMultiballSequencer.UserValue = tmrMultiballSequencer.UserValue + 1
    if tmrMultiballSequencer.UserValue > 50 Then    ' We've run for 2.5 seconds
        LightSeqPlayfield.StopPlay
        LightSeqPlayfield_PlayDone
        Exit Sub
    End If
    For each a in aPlayfieldLights
        c = yellowgreen
        i = RndNbr(2)
        if i = 2 then c = red
        SetLightColor a,c,-1
    Next
    tmrMultiballSequencer.Interval = 50
    tmrMultiballSequencer.Enabled = True
End Sub

' returns -1 if not a house sigil, or the house sigil index 0-6
' added in 1.1.17
Function GetLightHouseSigilIndex(lightname)
  dim SigilLight, i
  GetLightHouseSigilIndex = -1

  for i = 0 to UBound(HouseSigil)
    if lightname = HouseSigil(i).name Then
      GetLightHouseSigilIndex = i
      exit For
    end if
  next

end Function

Sub DoCastleMBSeq
    If LightSeqPlayfield.UserValue > 0 Then Exit Sub
    SavePlayfieldLightState
    PlayfieldSlowFade cyan,0.1
    LightSeqPlayfield.UserValue = 1
    LightSeqPlayfield.UpdateInterval = 10
    LightSeqPlayfield.Play SeqBlinking,,2,250
    LightSeqPlayfield.Play SeqDownOn,75,2
    LightSeqPlayfield.Play SeqAllOff
End Sub

'***************
' INSTANT INFO
'***************
' Format tells us how to format each scene (copied from Attract mode)
'  1: 1 line of text
'  2: 2 lines of text (same size)
'  3: 3 lines of text (small, big, medium)
'  4: image only, no text
'  5: video only, no text
'  6: video, scrolling text
'  7: 2 lines of text (big, small)
'  8: 3 lines of text (same size)
'  9: scrolling image
' 10: 1 line of text with outline
' 11: 3 lines of text (small, medium, medium)
Sub InstantInfo
    Dim scene,format,font,skipifnoflex,y,img
    Dim line1,line2,line3,font1
    InstantInfoTimer.Enabled = False
    If Tilted Then Exit Sub
    font = "udmd-f7by10.fnt" ' Most common font
    Select Case InfoPage
        Case 0: format=1:line1="INSTANT INFO"
        Case 1 ' current ball/credits/player
            If bGameInPlay = False Then InfoPage = 29 : InstantInfo : Exit Sub
            format=8:font="udmd-f6by8.fnt"
            line1="BALL "&BallsPerGame-BallsRemaining(CurrentPlayer)+1
            If DMDStd(kDMDStd_FreePlay) Then
                line2 = "FREE PLAY"
            Else
                Line2 = "CREDITS "&Credits
            End if
            line3 = "PLAYER "&CurrentPlayer&" IS UP"
        Case 2,3,4,5 ' current scores
            If PlayersPlayingGame < InfoPage-1 Then InfoPage = 6 : InstantInfo : Exit Sub
            format=2:line1="PLAYER "&InfoPage-1:line2=FormatScore(Score(InfoPage-1)):skipifnoflex=False
        Case 6 ' current player's gold
            format=1:line1=CurrentGold&" GOLD"
        Case 7 ' LoL status
            format=8:font="udmd-f3by7.fnt"
            line1="LORD   OF   LIGHT"
            line3=""
            If bLoLLit Then
                line2="LIT"
            ElseIf bLoLUsed Then
                line2="USED"
            Else
                line2="SHOOT   LEFT   TARGETS":line3="FOR   OUTLANE   BALL   SAVE"
            End If
        Case 8 ' Blackwater status
            format=11:font="FlexDMD.Resources.udmd-f4by5.fnt"
            line1="BLACKWATER MULTIBALL"
            If bLockIsLit Then line2="LOCK IS LIT" Else line2="SHOOT RIGHT TARGETS"
            If BallsInLock = 1 Then line3="1 BALL IN LOCK" Else line3 = BallsInLock & " BALLS IN LOCK"
        Case 9
            Dim i
            format=11:font="udmd-f6by8.fnt"
            line1 = "WALL MULTIBALL"
            If bWallMBReady Then
                line2 = "SHOOT RIGHT ORBIT"
            Else
                If WallMBCompleted = 0 Then i = 6-WallMBLevel Else i = 11-WallMBLevel
                line2 = i & " ADVANCEMENTS"
            End If
            line3 = "TO START MULTIBALL"
        Case 10
            If House(CurrentPlayer).WiCs >= 4 Then InfoPage=12:InstantInfo:Exit Sub
            format=3
            line1 = "WINTER IS COMING"
            line2 = 3-House(CurrentPlayer).WiCShots : If SelectedHouse = Greyjoy then line2 = line2 + 1
            If line2 = 1 Then line3 = "SHOT TO START" Else line3 = "SHOTS TO START"
        Case 11
            line1 = "CASTLE MULTIBALL"
            line2 = "CURRENT LEVEL"
            line3 = House(CurrentPlayer).UPFLevel - 1
            format=8:font="udmd-f6by8.fnt"
        Case 12
            If CompletedHouses >= 3 And EBisLit = 0 Then InfoPage=13:InstantInfo:Exit Sub
            format=2:font="udmd-f3by7.fnt"
            If EBisLit > 0 Then
                line1 = "EXTRA   BALL   IS   LIT"
                line2 = "SHOOT   RIGHT   ORBIT"
            Else
                line1 = 3-CompletedHouses & "  HOUSE  COMPLETIONS  NEEDED"
                line2 = "FOR   EXTRA   BALL"
            End If
        Case 13
            format=8:font="udmd-f3by7.fnt"
            line1 = "SELECTED  HOUSE:  "&HouseToUCString(SelectedHouse)
            line2 = "ACTION   BUTTON"
            line3 = HouseAbility(House(CurrentPlayer).ActionAbility)
        Case 14,15,16,17,18,19,20
            format=2:font="udmd-f6by8.fnt"
            line1 = "HOUSE "&HouseToUCString(InfoPage-13)
            If House(CurrentPlayer).Completed(InfoPage-13) Then
                line2="COMPLETED"
            ElseIf House(CurrentPlayer).Qualified(InfoPage-13) Then
                line2="IS LIT"
            Else line2 = (3 - House(CurrentPlayer).QualifyCount(InfoPage-13)) & " MORE TO LIGHT"
            End If
        Case 21
            format=12:font="udmd-f6by8.fnt"
            line1="SWORD COLLECTION"
            line2="SWORDS: "&SwordsCollected
            If SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier) < 3 Then line3="NEXT UNLOCKS "&SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier)+3&"X TIMES MULTIPLIER" Else line3=""
        Case 22
            format=11:font="udmd-f6by8.fnt"
            line1="SPINNER"
            line2="LEVEL "&Int((SpinnerLevel+2)/3)
            line3=SpinnerValue & " A SPIN"
        Case 23
            format=12:font="udmd-f6by8.fnt"
            line1="TOTAL BONUS"
            line2=FormatScore(BonusPoints(CurrentPlayer))
            line3="CURRENT MULTIPLIER "&BonusMultiplier(CurrentPlayer)&"X"
        'Could add more Info screens here. Real game doesn't have any more though
        Case 24: InfoPage=29:InstantInfo:Exit Sub

        Case 29:format=1:line1 = "REPLAY AT" & vbLf & FormatScore(ReplayScore)
        Case 30,31,32,33,34
            format = 8 : font="udmd-f6by8.fnt" : :skipifnoflex=False
            If InfoPage = 30 Then line1 = "GRAND CHAMPION" Else line1 = "HIGH SCORE #" & InfoPage-30
            line2 = HighScoreName(InfoPage-30) : line3 = FormatScore(HighScore(InfoPage-30))
        Case 35,36,37,38,39,40,41,42,43,44
            If bGameInPlay Then InfoPage=0:InstantInfo:Exit Sub
            format=3:skipifnoflex=False
            line1 = ChampionNames(i-35)&" CHAMPION"
            line2 = HighScoreName(i-30):line3 = HighScore(i-30)
    End Select
    If InfoPage >= 45 Then InfoPage = 0:InstantInfo:Exit Sub
    If bUseFlexDMD=False And skipifnoflex=True Then InfoPage=InfoPage+1:InstantInfo:Exit Sub

    ' Create the scene
    if bUseFlexDMD Then
        If format=4 or Format=5 or Format=6 or Format=9 Then
            Set scene = NewSceneWithVideo("attract"&InfoPage,img)
        Else
            Set scene = FlexDMD.NewGroup("attract"&InfoPage)
        End If

        ' Most of these modes aren't used for InstantInfo but we could probably combine the code with AttractMode to make it DRY
        Select Case format
            Case 1
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                scene.GetLabel("line1").SetAlignedPosition 64,16,FlexDMD_Align_Center
            Case 2
                Set font1 = FlexDMD.NewFont(font,vbWhite,vbWhite,0)
                scene.AddActor FlexDMD.NewLabel("line1",font1,line1)
                scene.AddActor FlexDMD.NewLabel("line2",font1,line2)
                scene.GetLabel("line1").SetAlignedPosition 64,9,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,22,FlexDMD_Align_Center
            Case 3
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont("skinny10x12.fnt",vbWhite,vbWhite,0),line2)
                scene.AddActor FlexDMD.NewLabel("line3",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),line3)
                scene.GetLabel("line1").SetAlignedPosition 64,3,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,15,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,27,FlexDMD_Align_Center
            Case 6
                scene.AddActor FlexDMD.NewGroup("scroller")
                scene.SetBounds 0,0-y,128,32+(2*y)  ' Create a large canvas for the text to scroll through
                With scene.GetGroup("scroller")
                    .SetBounds 0,y+32,128,y
                    .AddAction scene.GetGroup("scroller").ActionFactory().MoveTo(0,0,scrolltime)
                    .AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                End With
                scene.GetVideo("attract"&i&"vid").SetAlignedPosition 0,y,FlexDMD_Align_TopLeft ' move image to screen
                scene.GetLabel("line1").SetAlignedPosition 64,0,FlexDMD_Align_Top
            Case 7
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line2)
                scene.GetLabel("line1").SetAlignedPosition 64,14,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,27,FlexDMD_Align_Center
            Case 8
                Set font1 = FlexDMD.NewFont(font,vbWhite,vbWhite,0)
                scene.AddActor FlexDMD.NewLabel("line1",font1,line1)
                scene.AddActor FlexDMD.NewLabel("line2",font1,line2)
                scene.AddActor FlexDMD.NewLabel("line3",font1,line3)
                scene.GetLabel("line1").SetAlignedPosition 64,5,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,16,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,27,FlexDMD_Align_Center
            Case 9
                scene.SetBounds 0,0-y,128,32+(2*y)  ' Create a large canvas for the image to scroll through
                With scene.GetImage("attract"&i&"img")
                    .SetAlignedPosition 0,y+32,FlexDMD_Align_TopLeft
                    .AddAction scene.GetImage("attract"&i&"img").ActionFactory().MoveTo(0,0,scrolltime)
                End With
            Case 10
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,RGB(64, 64, 64),1),line1)
                scene.GetLabel("line1").SetAlignedPosition 64,16,FlexDMD_Align_Center
            Case 11
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line2)
                scene.AddActor FlexDMD.NewLabel("line3",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line3)
                scene.GetLabel("line1").SetAlignedPosition 64,3,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,13,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,25,FlexDMD_Align_Center
            Case 12
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line2)
                scene.AddActor FlexDMD.NewLabel("line3",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line3)
                scene.GetLabel("line1").SetAlignedPosition 64,5,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,16,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,27,FlexDMD_Align_Center
        End Select

        DMDDisplayScene scene
    End If
End Sub

'***************************
' Game-specific Attract Mode & High Score mgmt
'***************************

Function GameCheckHighScore(score)
    If DMDStd(kDMDStd_HighScoreAward) = 0 Then GameCheckHighScore=0 : Exit Function
    ' Did Player beat GC or high score 1-4? If so, award 0 or more credits based on Service Menu setting
    If score > HighScore(0) Then GameCheckHighScore= 1 + DMDStd(kDMDStd_GCAwards) : Exit Function
    If score > HighScore(1) Then GameCheckHighScore= 1 + DMDStd(kDMDStd_HS1Awards) : Exit Function
    If score > HighScore(2) Then GameCheckHighScore= 1 + DMDStd(kDMDStd_HS2Awards) : Exit Function
    If score > HighScore(3) Then GameCheckHighScore= 1 + DMDStd(kDMDStd_HS3Awards) : Exit Function
    If score > HighScore(4) Then GameCheckHighScore= 1 + DMDStd(kDMDStd_HS4Awards) : Exit Function
    GameCheckHighScore=1
    ' Did player beat their house's high score?
    If score > HighScore(4+SelectedHouse) Then Exit Function
    If  WHCHighScore(CurrentPlayer) > HighScore(12) Or _
        HotkHighScore(CurrentPlayer) > HighScore(13) Or _
        ITHighScore(CurrentPlayer) > HighScore(14) Then Exit Function
    GameCheckHighScore=0
End Function

Function GameSaveHighScore(score,name)
    Dim bm : bm=0
    If score > HighScore(4) Then HighScore(4) = score : HighScoreName(4) = name
    If score > HighScore(4+SelectedHouse) Then HighScore(4+SelectedHouse) = score : HighScoreName(4+SelectedHouse) = name : bm = bm Or (2^(4+SelectedHouse))
    If WHCHighScore(CurrentPlayer) > HighScore(12) Then HighScore(12) = WHCHighScore(CurrentPlayer) : HighScoreName(12) = name : bm = bm Or 4096
    If HotkHighScore(CurrentPlayer) > HighScore(13) Then HighScore(13) = HotkHighScore(CurrentPlayer) : HighScoreName(13) = name : bm = bm Or 8192
    If ITHighScore(CurrentPlayer) > HighScore(14) Then HighScore(14) = ITHighScore(CurrentPlayer) : HighScoreName(14) = name : bm = bm Or 16384
    GameSaveHighScore = bm
End Function

Sub GameAddCredit
    PlaySound ("Coin_In_"&RndNbr(3)), 0, CoinSoundLevel, 0, 0.25
    ' Jump to the Attract Mode Scene to the "Credits" screen
    if bAttractMode Then
        DMDFlush
        tmrAttractModeScene.UserValue = 5
        tmrAttractModeScene.Enabled = False
        tmrAttractModeScene.Interval = 10
        tmrAttractModeScene.Enabled = True
  Else
    DMDResetScoreScene
    End If
End Sub

Dim OathSeq
OathSeq = Array("NIGHT GATHERS"&vbLf&"AND NOW"&vbLf&"MY WATCH"&vbLf&"BEGINS"&vbLf&"IT SHALL"&vbLf&"NOT END"&vbLf&"UNTIL"&vbLf&"MY DEATH",_
            "I SHALL"&vbLf&"TAKE NO WIFE"&vbLf&"HOLD"&vbLf&"NO LANDS"&vbLf&"FATHER"&vbLf&"NO CHILDREN", _
            "I SHALL"&vbLf&"WEAR"&vbLf&"NO CROWNS"&vbLf&"AND WIN"&vbLf&"NO GLORY"&vbLf&"I SHALL"&vbLf&"LIVE AND DIE"&vbLf&"AT MY POST", _
            "I AM"&vbLf&"THE SWORD"&vbLf&"IN THE"&vbLf&"DARKNESS"&vbLf&"I AM"&vbLf&"THE WATCHER"&vbLf&"ON THE WALLS", _
            "I AM"&vbLf&"THE FIRE"&vbLf&"THAT BURNS"&vbLf&"AGAINST"&vbLf&"THE COLD"&vbLf&"THE LIGHT"&vbLf&"THAT BRINGS"&vbLf&"THE DAWN", _
            "THE HORN"&vbLf&"THAT WAKES"&vbLf&"THE SLEEPERS"&vbLf&"THE SHIELD"&vbLf&"THAT GUARDS"&vbLf&"THE REALMS"&vbLf&"OF MEN", _
            "I PLEDGE"&vbLf&"MY LIFE"&vbLf&"AND HONOR"&vbLf&"TO THE"&vbLf&"NIGHT'S WATCH"&vbLf&"FOR THIS NIGHT"&vbLf&"AND ALL"&vbLf&"THE NIGHTS"&vbLf&"TO COME")
Dim OathCnt
OathCnt = 0

Function GetNextOath()
    GetNextOath = OathSeq(OathCnt)
    OathCnt = OathCnt + 1
    If OathCnt >= 7 then OathCnt = 0
End Function

Dim ChampionNames
ChampionNames = Array("GRAND CHAMPION","HIGH SCORE #1","HIGH SCORE #2","HIGH SCORE #3","HIGH SCORE #4","STARK CHAMPION","BARATHEON CHAMPION", _
                    "LANNISTER CHAMPION","GREYJOY CHAMPION","TYRELL CHAMPION","MARTELL CHAMPION","TARGARYEN CHAMPION","WINTER HAS COME CHAMPION", _
                    "HAND OF THE KING CHAMPION","IRON THRONE CHAMPION")

Sub GameStartAttractMode
  debug.print "GameStartAttractMode"
    If bGameInPLay Then GameStopAttractMode : Exit Sub
    tmrDMDUpdate.Enabled = False
    SetUPFFlashers 0,cyan
    bAttractMode = True
    tmrAttractModeScene.UserValue = 0
    if bGameJustPlayed Then tmrAttractModeScene.UserValue = 24
    tmrAttractModeScene.Interval = 10
    tmrAttractModeScene.Enabled = True

    SavePlayfieldLightState
    tmrAttractModeLighting.UserValue = 10
    tmrAttractModeLighting.Interval = 10
    tmrAttractModeLighting.Enabled = True

    tmrFlipperSpeech.UserValue=0

    ' start the Front button attract lighting and topper attract for VR
  if (VRRoom = 1 or VRRoom = 3 or VRRoom = 4 or VRRoom = 5) then VR_StartButtOuter.blenddisablelighting = 0.8: StartButtonTimer.enabled = true
  if (VRRoom = 1 or VRRoom = 3 or VRRoom = 4 or VRRoom = 5) and Topper = 1 then
    StartVRDragon
    VRTopperTimer.enabled = true
    SmokeTimer.enabled = true
    VRTopperLights.enabled = false
  end if

  if VRRoom = 3 then SmokeTimerDragon.enabled = true else SmokeTimerDragon.enabled = false
End Sub

Sub GameStopAttractMode
    bAttractMode = False
    tmrAttractModeScene.Enabled = False
    tmrAttractModeLighting.Enabled = False
    LightSeqAttract.StopPlay
    RestorePlayfieldLightState True
    DMDClearQueue
    tmrDMDUpdate.Enabled = True
  if (VRRoom = 1 or VRRoom = 3 or VRRoom = 4 or VRRoom = 5) and Topper = 1 then
    'If Topper = 1 then
        StopVRDragon
        VRTopperTimer.enabled = false
        TopperLightsOff
        VRTopperLights.enabled = true
    End if
End Sub

Sub tmrFlipperSpeech_Timer
    Dim i
    Select Case tmrFlipperSpeech.UserValue
        Case 0: i=6
        Case 6: i=7
        Case 7: i=8
        Case 8: i=15
        Case Else: i=20     ' Default delay between speeches, in seconds
    End Select
    tmrFlipperSpeech.UserValue = i
    tmrFlipperSpeech.Enabled = 0
End Sub

Sub doFlipperSpeech(keycode)
    tmrAttractModeScene.Enabled = False
    If keycode = LeftFlipperKey Then
        tmrAttractModeScene.UserValue = tmrAttractModeScene.UserValue - 2
        If tmrAttractModeScene.UserValue < 0 Then tmrAttractModeScene.UserValue =  tmrAttractModeScene.UserValue + 24
    End If
    tmrAttractModeScene_Timer

    if tmrFlipperSpeech.Enabled <> 0 Then Exit Sub
    tmrFlipperSpeech.Interval = tmrFlipperSpeech.UserValue*1000
    tmrFlipperSpeech.Enabled=1
    PlaySoundVol "say-flipper"&RndNbr(43),VolCallout
End Sub

' To launch attract mode, disable DMDUpdateTimer and enble tmrAttractModeScene
' Attract mode is a big state machine running through various scenes in a loop. The timer is called
' after the scene has displayed for the set time, to move onto the next scene
' Scenes:
' 0: Stern logo - 3 seconds
' 1: PRESENTS - 2 seconds
' 2: GoT logo video - 21 seconds
' 3: part of Nights Watch oath on winter storm bg - 9 seconds
' 4: most recent score (just p1?) - 2 sec
' 5: credits - 2 sec
' 6: Replay at <x> - 2 sec
' 7: more oath - 9 sec
' 8: game logo - 3 sec
' 9-24: various high scores - 2 seconds each

' Format tells us how to format each scene
'  1: 1 line of text
'  2: 2 lines of text (same size)
'  3: 3 lines of text (small, big, medium)
'  4: image only, no text
'  5: video only, no text
'  6: video, scrolling text
'  7: 2 lines of text (big, small)
'  8: 3 lines of text (same size)
'  9: scrolling image
' 10: 1 line of text with outline
Sub tmrAttractModeScene_Timer
    Dim scene,scene2,img,line1,line2,line3
    Dim skip,font,format,scrolltime,y,delay,skipifnoflex,i
    Dim font1,font2,font3
    skip = False
    tmrAttractModeScene.Enabled = False
    If bGameInPLay Or bAttractMode = False Then GameStopAttractMode : Exit Sub
    delay = 2000
    skipifnoflex = True  ' Most scenes won't render without FlexDMD
    scrolltime = 0
    i = tmrAttractModeScene.UserValue
    Select Case tmrAttractModeScene.UserValue
        Case 0:img = "got-sternlogo":format=9:scrolltime=3:y=73:delay=3000
        Case 1:line1 = "PRESENTS":format=10:font="skinny10x12.fnt":delay=2000
        Case 2:img = "got-intro":format=5:delay=17200
        Case 3,7:img = "got-winterstorm":format=6:line1 = GetNextOath():delay=9000:font = "skinny10x12.fnt":scrolltime=9:y=100   ' Oath Text
        Case 4
            format=7:font="udmd-f11by18.fnt":line1=FormatScore(Score(1)):skipifnoflex=False  ' Last score
            If Score(1) > 999999999 Then font="udmd-f7by10.fnt"
            If DMDStd(kDMDStd_FreePlay) Then line2 = "FREE PLAY" Else Line2 = "CREDITS "&Credits
        Case 5
            format=1:font="udmd-f7by10.fnt":skipifnoflex=False
            If DMDStd(kDMDStd_FreePlay) Then
                line1 = "FREE PLAY"
            ElseIf Credits > 0 Then
                Line1 = "CREDITS "&Credits
            Else
                Line1 = "INSERT COINS"
            End if
        Case 6:format=1:font="udmd-f7by10.fnt":line1 = "REPLAY AT" & vbLf & FormatScore(ReplayScore)
        Case 8:img = "got-introframe":format=4:delay=3000
        Case 9,10,11,12,13,14,15,16,17,18,19,20,21,22,23
            If HighScore(i-9) = 0 Then delay = 5
            format=3:skipifnoflex=False
            line1 = ChampionNames(i-9)
            line2 = HighScoreName(i-9):line3 = FormatScore(HighScore(i-9))
        Case 24,25,26,27:
            If PlayersPlayingGame < i-23 Then
                i=27:delay=10
            Else
                delay=4000
                If DMDStd(kDMDStd_FreePlay) Then line2 = "FREE PLAY" Else Line2 = "CREDITS "&Credits
                If PlayersPlayingGame > 1 Then
                    format=8:line1="PLAYER "&i-24
                    line3 = line2
                    line2=FormatScore(Score(i-24))
                    font="udmd-f6by8.fnt"
                Else
                    format=7:font="udmd-f11by18.fnt":line1=FormatScore(Score(1)):skipifnoflex=False  ' Last score
                    If Score(1) > 999999999 Then font="udmd-f7by10.fnt"
                End If
            End If

    End Select
    If i = 23 or i = 27 Then tmrAttractModeScene.UserValue = 0 Else tmrAttractModeScene.UserValue = i + 1
    If bUseFlexDMD=False And skipifnoflex=True Then tmrAttractModeScene.Interval = 10 Else tmrAttractModeScene.Interval = delay
    tmrAttractModeScene.Enabled = True

    ' Create the scene
    if bUseFlexDMD Then
        If format=4 or Format=5 or Format=6 or Format=9 Then
            Set scene = NewSceneWithVideo("attract"&i,img)
        Else
            Set scene = FlexDMD.NewGroup("attract"&i)
        End If

        Select Case format
            Case 1
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                scene.GetLabel("line1").SetAlignedPosition 64,16,FlexDMD_Align_Center
            Case 3
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont("skinny10x12.fnt",vbWhite,vbWhite,0),line2)
                scene.AddActor FlexDMD.NewLabel("line3",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),line3)
                scene.GetLabel("line1").SetAlignedPosition 64,3,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,15,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,27,FlexDMD_Align_Center
            Case 6
                scene.AddActor FlexDMD.NewGroup("scroller")
                scene.SetBounds 0,0-y,128,32+(2*y)  ' Create a large canvas for the text to scroll through
                With scene.GetGroup("scroller")
                    .SetBounds 0,y+32,128,y
                    .AddAction scene.GetGroup("scroller").ActionFactory().MoveTo(0,0,scrolltime)
                    .AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                End With
                scene.GetVideo("attract"&i&"vid").SetAlignedPosition 0,y,FlexDMD_Align_TopLeft ' move image to screen
                scene.GetLabel("line1").SetAlignedPosition 64,0,FlexDMD_Align_Top
            Case 7
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line2)
                scene.GetLabel("line1").SetAlignedPosition 64,14,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,27,FlexDMD_Align_Center
            Case 8
                Set font1 = FlexDMD.NewFont(font,vbWhite,vbWhite,0)
                scene.AddActor FlexDMD.NewLabel("line1",font1,line1)
                scene.AddActor FlexDMD.NewLabel("line2",font1,line2)
                scene.AddActor FlexDMD.NewLabel("line3",font1,line3)
                scene.GetLabel("line1").SetAlignedPosition 64,5,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,16,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,27,FlexDMD_Align_Center
            Case 9
                scene.SetBounds 0,0-y,128,32+(2*y)  ' Create a large canvas for the image to scroll through
                With scene.GetImage("attract"&i&"img")
                    .SetAlignedPosition 0,y+32,FlexDMD_Align_TopLeft
                    .AddAction scene.GetImage("attract"&i&"img").ActionFactory().MoveTo(0,0,scrolltime)
                End With
            Case 10
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,RGB(64, 64, 64),1),line1)
                scene.GetLabel("line1").SetAlignedPosition 64,16,FlexDMD_Align_Center
        End Select

        DMDDisplayScene scene
    End If
End Sub

' Attract mode light sequence
' If a game has just ended, the first sequence is the Flashers - two strobes of each flasher, moving around the table for 20 seconds
' Then game moves into standard attract mode:
' For first part, light sequence is sigils and playfield X rotating, all others flashing quickly bottom to top.
' The lights that are on are all the same colour as whichever sigil is lit, if we care
' Then sequences:
'       - 4-color slow transition thru RGBP
'       - bottom-to-top fade sweeps in sequential colours 6 total) ROYGBP
'       - top-to-bottom fade sweeps
'       - right to left fade sweeps, faster
'       - left to right fade sweep
'       - 7 color faster transition

Dim AttractFlashers : AttractFlashers = Array(li222,fl235,fl242,fl245,fl243,fl244,fl237,fl236)
Dim AttractPFcolors
Dim PFCurrentColor
AttractPFcolors = Array(red,amber,yellow,green,blue,purple)
Dim ColorWheel
' Wheel of colours used to fade between 6 primary colours
' These values were calculated using a script. They are the intermediate values between the primary colours,
' calculated as "R*65536 + G*256 + B", equivalent to the hex number &H00RRGGBB
ColorWheel = Array(16711680,16711936,16712192,16712448,16712704,16712960,16713216,16713472,16713728,16713984,16714496,16714752,16715008,16715264,16715520,16715776,16716032,16716288,16716544,16716800, _
                    16717312,16717568,16717824,16718080,16718336,16718592,16718848,16719104,16719360,16719616,16720128,16720384,16720640,16720896,16721152,16721408,16721664,16721920,16722176,16722688, _
                    16722944,16723200,16723456,16723712,16723968,16724224,16724480,16724736,16724992,16725504,16725760,16726016,16726272,16726528,16726784,16727040,16727296,16727552, _
                    16728064,16728832,16729600,16730368,16731392,16732160,16732928,16733952,16734720,16735488,16736256,16737280,16738048,16738816,16739840,16740608,16741376,16742144,16743168,16743936, _
                    16744704,16745728,16746496,16747264,16748288,16749056,16749824,16750592,16751616,16752384,16753152,16754176,16754944,16755712,16756480,16757504,16758272,16759040,16760064,16760832, _
                    16761600,16762624,16763392,16764160,16764928,16765952,16766720,16767488,16768512,16769280,16770048,16770816,16771840,16772608,16773376,16774400,16775168,16775936, _
                    16776960,16449280,16187136,15859456,15597312,15335168,15007488,14745344,14417664,14155520,13893376,13565696,13303552,12975872,12713728,12451584,12123904,11861760,11534080,11271936, _
                    11009792,10682112,10419968,10092288,9830144,9568000,9240320,8978176,8650496,8388352,8126208,7798528,7536384,7208704,6946560,6684416,6356736,6094592,5766912,5504768, _
                    5242624,4914944,4652800,4325120,4062976,3800832,3473152,3211008,2883328,2621184,2359040,2031360,1769216,1441536,1179392,917248,589568,327424, _
                    65280,64004,62984,61709,60689,59669,58394,57374,56099,55079,54059,52784,51764,50489,49469,48449,47174,46154,44879,43859, _
                    42839,41564,40544,39269,38249,37229,35954,34934,33659,32639,31619,30344,29324,28049,27029,26009,24734,23714,22439,21419, _
                    20399,19124,18104,16829,15809,14789,13514,12494,11219,10199,9179,7904,6884,5609,4589,3569,2294,1274, _
                    255,131325,262396,393467,524538,721145,852216,983287,1114358,1245429,1442036,1573107,1704177,1835248,1966319,2162926,2293997,2425068,2556139,2687210, _
                    2883817,3014888,3145959,3277030,3408100,3604707,3735778,3866849,3997920,4128991,4325598,4456669,4587740,4718811,4915418,5046488,5177559,5308630,5439701,5636308, _
                    5767379,5898450,6029521,6160592,6357199,6488270,6619341,6750411,6881482,7078089,7209160,7340231,7471302,7602373,7798980,7930051,8061122,8192193,8388800     )

Sub tmrAttractModeLighting_Timer
    Dim i,seqtime,c,a
    tmrAttractModeLighting.Enabled = False
    i=1
    Select Case Int(tmrAttractModeLighting.UserValue)
        Case 0  ' Random
            LightSeqAttract.StopPlay
            RestorePlayfieldLightState True
            LightSeqAttract.UpdateInterval = 75
            LightSeqAttract.Play SeqRandom,25,,10000
            LightSeqAttract.Play SeqAllOff
            seqtime = 10000
        Case 1,9
            ' Step through the colorwheel values, once every 30ms (about 10 seconds to fade between the 6 main colours)
            seqtime = 30:i=0.001
            c = int((tmrAttractModeLighting.UserValue - Int(tmrAttractModeLighting.UserValue))*1000)
            if c = 0 Then
                LightSeqAttract.StopPlay
                PlayfieldSlowFade red,10
                For each a in aPlayfieldLights
                    a.State = 1
                Next
            End if
            If c > 290 then
                ' Jump to the next effect
                i = 0: tmrAttractModeLighting.UserValue = Int(tmrAttractModeLighting.UserValue)+1
                For each a in aPlayfieldLights
                    a.State = 0
                Next
            Else
                c = ColorWheel(c)
                For each a in aPlayfieldLights
                    a.colorfull = c
                Next
            End If
        Case 2
            LightSeqAttract.StopPlay
            c = int((tmrAttractModeLighting.UserValue - Int(tmrAttractModeLighting.UserValue))*10)
            if c = 5 then Me.UserValue = 3:i = 0 Else i = 0.1
            PlayfieldSlowFade AttractPFcolors(c),0.1
            LightSeqAttract.UpdateInterval = 10
            LightSeqAttract.Play SeqUpOn, 50, 1
            seqtime = 1700
        Case 3
            LightSeqAttract.StopPlay
            c = int((tmrAttractModeLighting.UserValue - Int(tmrAttractModeLighting.UserValue))*10)
            if c = 5 then Me.UserValue = 4:i = 0 Else i = 0.1
            PlayfieldSlowFade AttractPFcolors(c),0.1
            LightSeqAttract.UpdateInterval = 10
            LightSeqAttract.Play SeqDownOn, 50, 1
            seqtime = 1700
        Case 4
            LightSeqAttract.StopPlay
            c = int((tmrAttractModeLighting.UserValue - Int(tmrAttractModeLighting.UserValue))*10)
            if c = 5 then Me.UserValue = 5:i = 0 Else i = 0.1
            PlayfieldSlowFade AttractPFcolors(c),0.1
            LightSeqAttract.UpdateInterval = 10
            LightSeqAttract.Play SeqRightOn, 100, 1
            seqtime = 1200
        Case 5
            LightSeqAttract.StopPlay
            c = int((tmrAttractModeLighting.UserValue - Int(tmrAttractModeLighting.UserValue))*10)
            if c = 5 then Me.UserValue = 6:i = 0 Else i = 0.1
            PlayfieldSlowFade AttractPFcolors(c),0.1
            LightSeqAttract.UpdateInterval = 10
            LightSeqAttract.Play SeqLeftOn, 100, 1
            seqtime = 1200
        Case 6
            LightSeqAttract.StopPlay
            c = int((tmrAttractModeLighting.UserValue - Int(tmrAttractModeLighting.UserValue))*10)
            if c = 5 then Me.UserValue = 7:i = 0 Else i = 0.1
            PlayfieldSlowFade AttractPFcolors(c),0.1
            LightSeqAttract.UpdateInterval = 20
            LightSeqAttract.Play SeqCircleOutOn, 25, 1
            seqtime = 2500
        Case 7
            LightSeqAttract.StopPlay
            c = int((tmrAttractModeLighting.UserValue - Int(tmrAttractModeLighting.UserValue))*10)
            if c = 5 then Me.UserValue = 8:i = 0 Else i = 0.1
            PlayfieldSlowFade AttractPFcolors(c),0.1
            LightSeqAttract.UpdateInterval = 7
            LightSeqAttract.Play SeqClockLeftOn, 25, 1
            seqtime = 2500
        Case 8
            LightSeqAttract.StopPlay
            c = int((tmrAttractModeLighting.UserValue - Int(tmrAttractModeLighting.UserValue))*10)
            if c = 5 then Me.UserValue = 9:i = 0 Else i = 0.1
            PlayfieldSlowFade AttractPFcolors(c),0.1
            LightSeqAttract.UpdateInterval = 7
            LightSeqAttract.Play SeqRadarRightOn, 25, 1
            seqtime = 2500
        Case 10
            If bGameJustPlayed Then
                ' Run the Flasher sequence
                seqtime = 333:i=0.001
                c = int((tmrAttractModeLighting.UserValue - Int(tmrAttractModeLighting.UserValue))*1000)
                if c = 0 then LightSeqAttract.StopPlay : SetUPFFlashers 0,white  ' initial run thru
                if c < 60 Then
                    Set a = AttractFlashers(c MOD 8)
                    if (a.name = "li222") Then SetFlasherWithColor a,7,ice Else SetFlasher a,7
                Else ' Done
                    Me.UserValue = 11 : i = 0
                    bGameJustPlayed = False
                End If
            End If
        Case Else
            ' Loop back to first effect
      tmrAttractModeLighting.UserValue = 0:i=0:seqtime=10
    End Select
    tmrAttractModeLighting.UserValue = tmrAttractModeLighting.UserValue + i
    tmrAttractModeLighting.Interval = seqtime
    tmrAttractModeLighting.Enabled = True

End Sub

'**************************
' Game-specific DMD support
'**************************

' This just warms the filesystem cache by preloading all of the images used in
' animation sequences. Alternatively we could actually precreate the scenes and
' store them in array to be referenced in the game
Dim SequenceScene(30)
Sub PreloadImageSequenceScenes
    Dim scene
    Set scene = NewSceneFromImageSequenceRange("clock","wallclock",1,340,25,0,1)
    Set scene = NewSceneFromImageSequence("hitscenefire","cmbsjpfire",52,25,0,0)
    Set scene = NewSceneFromImageSequence("itfg","itintro",110,20,0,0)
    Set scene = NewSceneFromImageSequence("img1","bonusx",50,20,2,0)
    Set scene = NewSceneFromImageSequence("hitscenefire","cmbsjpfire",52,25,0,0)
    Set scene = NewSceneFromImageSequence("hitscenevid","got-targaryenqualify1",32,25,0,0)
    Set scene = NewSceneFromImageSequence("hitscenevid","got-targaryenqualify2",77,25,0,0)
    Set scene = NewSceneFromImageSequence("hitscenevid","got-targaryenqualify3",62,25,0,0)
    Set scene = NewSceneFromImageSequence("img1","match",280,30,0,0)
End Sub


' The SuperJackpot scene is pretty complicated, but a highlight of the real table, so
' we need to try and reproduce it.
' Sequence is:
'   - battering ram doors open on DMD leaving a space in the middle
'   - in the middle space, the letters S-U-P-E-R-J-A-C-K-P-O-T spell out sequential with a deep drum hit for each
'   - drum is timed with the letter showing, so need a timer that does both at once
'   - at the end of the letters, scene flashes rapidly between score and all white
'   - all lights on playfield flash.
Dim BWSJPScene
Sub DMDBlackwaterSJPScene(score,lvl)
    Dim i
    If bUseFlexDMD Then
        Set BWSJPScene = NewSceneWithVideo("bwsjp","got-blackwatersjp")
        For i = 1 to 12
            BWSJPScene.AddActor FlexDMD.NewImage("img"&i,"got-bwsjpletter"&i&".png")
            With BWSJPScene.GetImage("img"&i)
                .SetAlignedPosition 64,16,FlexDMD_Align_Center
                .Visible = 0
            End With
        Next
        ' Add the score and white background
        BWSJPScene.AddActor FlexDMD.NewLabel("score1",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),lvl&"X SUPER JACKPOT")
        BWSJPScene.AddActor FlexDMD.NewLabel("score2",FlexDMD.NewFont("udmd-f7by10.fnt", vbWhite, vbBlack, 0),score)
        BWSJPScene.AddActor FlexDMD.NewImage("blank","got-blankwhite.png")
        BWSJPScene.AddActor FlexDMD.NewLabel("scoreinv1",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),lvl&"X SUPER JACKPOT")
        BWSJPScene.AddActor FlexDMD.NewLabel("scoreinv2",FlexDMD.NewFont("udmd-f7by10.fnt", vbWhite, vbBlack, 0),score)
        With BWSJPScene.GetLabel("score1")
            .SetAlignedPosition 64,8,FlexDMD_Align_Center
            .Visible = 0
        End With
        With BWSJPScene.GetLabel("score2")
            .SetAlignedPosition 64,20,FlexDMD_Align_Center
            .Visible = 0
        End With
        With BWSJPScene.GetLabel("scoreinv1")
            .SetAlignedPosition 64,8,FlexDMD_Align_Center
            .Visible = 0
        End With
        With BWSJPScene.GetLabel("scoreinv2")
            .SetAlignedPosition 64,20,FlexDMD_Align_Center
            .Visible = 0
        End With
        BWSJPScene.GetImage("blank").Visible = 0

        DMDEnqueueScene BWSJPScene,0,2600,5000,2500,""
    Else
        DisplayDMDText "SUPER JACKPOT",score,2000
        PlaySoundVol "say-super-jackpot",VolCallout
    End If
End Sub

'This timer is triggered from DMDUpdate_Timer once the super jackpot scene starts playing
Sub tmrSJPScene_Timer
    Dim i,delay
    Me.Enabled = False
    i = Me.UserValue
    i = i + 1
    delay = 175
    FlexDMD.LockRenderThread
    If i = 1 Then   ' Turn on the first letter
        BWSJPScene.GetImage("img"&i).Visible = 1
        PlaySoundVol "gotfx-sjpdrum",VolDef
        DoBWSJPSeq False
        LightSeqGi.UpdateInterval = 5
        LightSeqGi.Play SeqBlinking, , 12, 50
        SetUPFFlashers 5,yellowgreen    ' flash 10 times rapidly
    ElseIf i < 13 Then  ' turn on the next letter
        BWSJPScene.GetImage("img"&(i-1)).Visible = 0
        BWSJPScene.GetImage("img"&i).Visible = 1
        PlaySoundVol "gotfx-sjpdrum",VolDef
    ElseIf i = 13 Then ' turn off the last letter and turn on the score
        BWSJPScene.GetImage("img"&(i-1)).Visible = 0
        BWSJPScene.GetVideo("bwsjpvid").Visible = 0
        BWSJPScene.GetLabel("score1").Visible = 1
        BWSJPScene.GetLabel("score2").Visible = 1
        PlaySoundVol "say-super-jackpot",VolCallout
        PlaySoundVol "gotfx-sjpahh",VolDef
        delay = 33
        DoBWSJPSeq True
    ElseIf i < 73 Then  ' toggle inverted score with white background on and off
        delay = 50
        If (i MOD 2) = 0 Then
            BWSJPScene.GetLabel("scoreinv1").Visible = 1
            BWSJPScene.GetLabel("scoreinv2").Visible = 1
            BWSJPScene.GetImage("blank").Visible = 1
        Else
            BWSJPScene.GetLabel("scoreinv1").Visible = 0
            BWSJPScene.GetLabel("scoreinv2").Visible = 0
            BWSJPScene.GetImage("blank").Visible = 0
            PlayExistingSoundVol "gotfx-sjpdrum",VolDef,1
        End if
    Else
        FlexDMD.UnlockRenderThread
        Exit Sub
    End If
    FlexDMD.UnlockRenderThread
    Me.UserValue = i
    Me.Interval = delay
    Me.Enabled = True
End Sub

' Played at the end of Blackwater Multiball
Sub DMDBlackwaterCompleteScene
    Dim scene
    If bUseFlexDMD Then
        If BlackwaterScore > 0 Then
            Set scene = NewSceneWithVideo("bwcomplete","got-blackwatertotal")
            scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbBlack, 0),"BLACKWATER TOTAL")
            scene.GetLabel("txt").SetAlignedPosition 64,10,FlexDMD_Align_Center
            scene.AddActor FlexDMD.NewLabel("score",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),FormatScore(BlackwaterScore))
            scene.GetLabel("score").SetAlignedPosition 64,20,FlexDMD_Align_Center
            DMDEnqueueScene scene,0,5000,5000,2000,"gotfx-blackwatertotal"
        End if
    Else
        DisplayDMDText "BLACKWATER TOTAL",FormatScore(BlackwaterScore),4000
        PlaySoundVol "gotfx-blackwatertotal",VolDef
    End If
    If BlackwaterScore > 50000000 Then vpmTimer.addTimer 2500,"PlaySoundVol ""say-blackwatertotal"&RndNbr(2)&""",VolCallout '"
End Sub

Sub DMDCastleMBCompleteScene
    Dim scene
    If bUseFlexDMD Then
        If CastleMBScore > 0 Then
            Set scene = FlexDMD.NewGroup("cmbcomplete")
            scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbBlack, 0),"CASTLE   MULTIBALL"&vbLf&"TOTAL")
            scene.GetLabel("txt").SetAlignedPosition 64,10,FlexDMD_Align_Center
            scene.AddActor FlexDMD.NewLabel("score",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),FormatScore(CastleMBScore))
            scene.GetLabel("score").SetAlignedPosition 64,25,FlexDMD_Align_Center
            DMDEnqueueScene scene,0,3000,3000,2000,"gotfx-battletotal"
        End If
    Else
        DisplayDMDText "MULTIBALL TOTAL",FormatScore(CastleMBScore),4000
        PlaySoundVol "gotfx-battletotal",VolDef
    End If
End Sub

Sub DMDWallMBCompleteScene
Dim scene
    If bUseFlexDMD Then
        If WallMBScore > 0 Then
            Set scene = FlexDMD.NewGroup("wmbcomplete")
            scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbBlack, 0),"BATTLE   FOR   THE   WALL"&vbLf&"TOTAL")
            scene.GetLabel("txt").SetAlignedPosition 64,10,FlexDMD_Align_Center
            scene.AddActor FlexDMD.NewLabel("score",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),FormatScore(WallMBScore))
            scene.GetLabel("score").SetAlignedPosition 64,25,FlexDMD_Align_Center
            DMDEnqueueScene scene,0,3000,3000,2000,"gotfx-wallmb-total"
        End If
    Else
        DisplayDMDText "MULTIBALL TOTAL",FormatScore(WallMBScore),4000
        PlaySoundVol "gotfx-wallmb-total",VolDef
    End If
End Sub

Sub DMDWHCMBCompleteScene
Dim scene
    If bUseFlexDMD Then
        If WHCMBScore > 0 Then
            Set scene = FlexDMD.NewGroup("whccomplete")
            scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbBlack, 0),"WINTER   HAS   COME"&vbLf&"TOTAL")
            scene.GetLabel("txt").SetAlignedPosition 64,10,FlexDMD_Align_Center
            scene.AddActor FlexDMD.NewLabel("score",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),FormatScore(WHCMBScore))
            scene.GetLabel("score").SetAlignedPosition 64,25,FlexDMD_Align_Center
            DMDEnqueueScene scene,0,3000,3000,2000,"gotfx-whctotal"
        End If
    Else
        DisplayDMDText "MULTIBALL TOTAL",FormatScore(WHCMBScore),4000
        PlaySoundVol "gotfx-whctotal",VolDef
    End If
End Sub

Sub DMDHotkMBCompleteScene
    Dim scene
    If bUseFlexDMD Then
        If HotkScore > 0 Then
            Set scene = NewSceneWithImage("hotkcomplete","got-hotk-sigil")
            scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbBlack, 0),"HAND  OF  THE  KING"&vbLf&"TOTAL")
            scene.GetLabel("txt").SetAlignedPosition 96,10,FlexDMD_Align_Center
            scene.AddActor FlexDMD.NewLabel("score",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),FormatScore(HotkScore))
            scene.GetLabel("score").SetAlignedPosition 96,25,FlexDMD_Align_Center
            DMDEnqueueScene scene,0,3000,3000,2000,"gotfx-hotktotal"
        End If
    Else
        DisplayDMDText "MULTIBALL TOTAL",FormatScore(WHCMBScore),4000
        PlaySoundVol "gotfx-whctotal",VolDef
    End If
    PlaySoundVol "say-hotk-total",VolCallout
End Sub

Sub DMDIronThroneMBCompleteScene
    Dim Scene
    If ITScore > ITHighScore(CurrentPlayer) Then ITHighScore(CurrentPlayer) = ITScore
    If ITScore > 0 Then
        Set Scene = FlexDMD.NewGroup("itcomplete")
        scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbBlack, 0),"IRON   THRONE"&vbLf&"TOTAL")
        scene.GetLabel("txt").SetAlignedPosition 64,10,FlexDMD_Align_Center
        scene.AddActor FlexDMD.NewLabel("score",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),FormatScore(ITScore))
        scene.GetLabel("score").SetAlignedPosition 64,25,FlexDMD_Align_Center
        DMDEnqueueScene scene,0,3000,3000,2000,"gotfx-battletotal"
    End If
End Sub

Sub DMDMadnessMBCompleteScene
    Dim Scene
    Set Scene = FlexDMD.NewGroup("mmcomplete")
    scene.AddActor FlexDMD.NewLabel("txt",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbBlack, 0),"MIDNIGHT   MADNESS"&vbLf&"TOTAL")
    scene.GetLabel("txt").SetAlignedPosition 64,10,FlexDMD_Align_Center
    scene.AddActor FlexDMD.NewLabel("score",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 0),FormatScore(MMTotalScore))
    scene.GetLabel("score").SetAlignedPosition 64,25,FlexDMD_Align_Center
    DMDEnqueueScene scene,0,3000,3000,2000,"gotfx-battletotal"
End Sub

Dim BaratheonSpinnerScene
Sub DMDBaratheonSpinnerScene(value)
    If bUseFlexDMD Then
        If AccumulatedSpinnerValue = 0 or IsEmpty(BaratheonSpinnerScene) Then ' Spinner has gone a little while without spinning. Start new scene
            Set BaratheonSpinnerScene = NewSceneWithVideo("barspin","got-baratheonbattlespinner")
            BaratheonSpinnerScene.AddActor FlexDMD.NewLabel("bartop",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbBlack, 0),"BARATHEON")
            BaratheonSpinnerScene.AddActor FlexDMD.NewLabel("barmid",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbBlack, 0),"VALUE"&vbLf&"INCREASES")
            BaratheonSpinnerScene.AddActor FlexDMD.NewLabel("barval",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbBlack, 0),FormatScore(value))
            BaratheonSpinnerScene.GetLabel("bartop").SetAlignedPosition 2,2,FlexDMD_Align_TopLeft
            BaratheonSpinnerScene.GetLabel("barmid").SetAlignedPosition 2,11,FlexDMD_Align_TopLeft
        Else
            FlexDMD.LockRenderThread
            BaratheonSpinnerScene.GetLabel("barval").Text = FormatScore(value)
            FlexDMD.UnlockRenderThread
        End If
        BaratheonSpinnerScene.GetLabel("barval").SetAlignedPosition 2,30,FlexDMD_Align_BottomLeft
        DMDEnqueueScene BaratheonSpinnerScene,0,1000,2000,1000,"gotfx-baratheonbattlespinner"
    Else
        DisplayDMDText "VALUE INCREASES",value,0
        PlaySoundVol "gotfx-baratheonbattlespinner",VolDef
    End If
End Sub

' Create a "Hit" scene which plays every time a qualifying or battle target is hit
'  vid   - the name of the video for the first part of the scene
'  sound - the sound to play with the video
'  delay - How long to wait before cutting to the second part of the scene, in seconds (float)
'  line1-3 - Up to 3 lines of text
'  combo - If 0, text is full width. Otherwise, Combo multiplier is on the right side
'  format - The following formats (layouts) are supported
'           0 - top line 3x7, middle line 7x12 skinny, bottom line 3x5. This is used for all Qualfying hits
'           1 - top line 5x7, middle line 7x12 skinny, bottom line 7x5. Used for most battle scenes
'           2 - 3 lines of 3x7 font, used for lannister battle gold hits
'           3 - 2 lines of text. Top line 3x7, main line 6x8 score. Used for jackpots, Castle MB levels, and Targaryen hurry-up hits
'           4 - 2 lines of Skinny font. Used for Castle Multiball start
'           5 - 2 lines of text, same as 3, but no video, just a framed text scene and combo text. Used for UPF awards
'           6 - same layout as 0, but video is made using image sequence, and runs at same time as text. used for Targaryen qualify hits
'           7 - same text layout as 3, but video runs in the background at the same time
'           8 - same as 3 but 3 lines. used for Castle MB Super JP
'           9 - same as 8, but text overlays video after delay, rather than replaces it. used for Castle MB Jackpots
'           10 -3 lines of text, line1 is 3x5, line2,3 are 3x7. Text is shifted right and background image is on the left. Used for UPF castle awards
'           11 - Same as 10, but 1 line of 6x8 text
'           12 - 3 lines of text. Line1,2 are 3x7, line 3 is 6x8
Sub DMDPlayHitScene(vid,sound,delay,line1,line2,line3,combo,format)
    Dim scene,scenevid,font1,font2,font3,x,y1,y2,y3,combotxt,pri,l3vis
    Set scenevid = Nothing
    If bUseFlexDMD Then
        If format = 6 Then
            Set scene = FlexDMD.NewGroup("hitscene")
        Else
            Set scene = NewSceneWithVideo("hitscene",vid)
            Set scenevid = scene.GetVideo("hitscenevid")
            If scenevid is Nothing Then Set scenevid = scene.getImage("hitsceneimg")
            If vid="got-cmbsuperjp" Then scene.AddActor NewSceneFromImageSequence("hitscenefire","cmbsjpfire",52,25,0,0)
        End If
        y1 = 4: y2 = 15: y3 = 26
        x = 64
        l3vis = False
        Select Case format
            Case 0,6
                Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font2 = FlexDMD.NewFont("skinny7x12.fnt", vbWhite, vbWhite, 0)
                Set font3 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
                l3vis = True
            Case 1
                if Len(line1) > 12 Then
                    Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Else
                    Set font1 = FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0)
                End if
                Set font2 = FlexDMD.NewFont("skinny7x12.fnt", vbWhite, vbWhite, 0)
                set font3 = FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0)
            Case 2
                Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font2 = font1
                Set font3 = font1
                l3vis = True
            Case 3,5,7,8,9
                Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font2 = FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0)
                Set font3 = font1
                If format = 8 or format = 9 then y1=10 : y2=23
            Case 4
                Set font1 = FlexDMD.NewFont("skinny10x12.fnt", vbWhite, vbWhite, 0)
                Set font2 = font1
                Set font3 = font1
                y1 = 8: y2 = 23
            Case 10
                Set font1 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
                Set font2 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font3 = font1
                x = 50
                l3vis = True
            Case 11
                Set font2 = FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0)
                Set font1 = font2
                Set font3 = font2
                x = 50
                l3vis = False
            Case 12
                Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font3 = FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0)
                Set font2 = font1
                l3vis = True
        End Select

        scene.AddActor FlexDMD.NewGroup("hitscenetext")
        With scene.GetGroup("hitscenetext")
            .AddActor FlexDMD.NewLabel("line1",font1,line1)
            .AddActor FlexDMD.NewLabel("line2",font2,line2)
            .AddActor FlexDMD.NewLabel("line3",font3,line3)
            .Visible = False
        End With
        ' If a combo multiplier was specified, set it up on the right side
        If combo > 0 Then
            If format < 10 then x = 46
            combotxt = ""
            Select Case True
                Case (combo > 1 and PlayfieldMultiplierVal = 1) : combotxt = "COMBO"
                Case (combo = 1 and PlayfieldMultiplierVal > 1) : combotxt = "PLAYFIELD"
                Case (combo > 1 and PlayfieldMultiplierVal > 1) : combotxt = "MIXED"
                Case ((format = 5 Or format=8) and combo > 1 and PlayfieldMultiplierVal = 1) : combotxt = "UPPER"
            End Select
            With scene.GetGroup("hitscenetext")
                .AddActor FlexDMD.NewLabel("combo",FlexDMD.NewFont("FlexDMD.Resources.udmd-f12by24.fnt", vbWhite, vbWhite, 0),(combo*PlayfieldMultiplierVal)&"X")
                .AddActor FlexDMD.NewLabel("combotxt",FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0),combotxt)
                .GetLabel("combo").SetAlignedPosition 104,19,FlexDMD_Align_Center
                .GetLabel("combotxt").SetAlignedPosition 104,3,FlexDMD_Align_Center
            End With
        End If
        With scene.GetGroup("hitscenetext")
            .GetLabel("line1").SetAlignedPosition x,y1,FlexDMD_Align_Center
            .GetLabel("line2").SetAlignedPosition x,y2,FlexDMD_Align_Center
            .GetLabel("line3").SetAlignedPosition x,y3,FlexDMD_Align_Center
        End With

        ' If line2 is a score, flash it
        If format <> 2 and format <> 4 and format < 6 Then BlinkActor scene.GetGroup("hitscenetext").GetLabel("line2"),100,10

        scene.GetGroup("hitscenetext").GetLabel("line3").Visible = l3vis

        ' After delay, disable video/image and enable text
        If format <> 6 And delay > 0 And Not (scenevid Is Nothing) Then
            If format <> 9 then DelayActor scenevid,delay,False
            If vid="got-cmbsuperjp" then DelayActor scene.GetGroup("hitscenefire"),delay,False
            DelayActor scene.GetGroup("hitscenetext"),delay,True
        Else
            scene.GetGroup("hitscenetext").Visible = True
            delay = 0
        End If
        If format = 6 Then
            Dim i
            Select Case vid
                Case "got-targaryenqualify1": i=32
                Case "got-targaryenqualify2": i=77
                Case "got-targaryenqualify3": i=62
            End Select
            scene.AddActor NewSceneFromImageSequence("hitscenevid",vid,i,25,0,0)
            delay = int(((i-30)/25)+1)
        End if

        'Special case - make Jackpot hit scenes priority 0
        If sound = "gotfx-bwexplosion" Then pri=0 Else pri=1

        ' Castle awards are a big deal, keep them on screen for longer
        If format = 10 then delay = 2

        DMDEnqueueScene scene,pri,delay*1000+1000,delay*1000+2000,3000,sound
    Else
        DisplayDMDText line1,line2,2000
        PlaySoundVol sound,VolDef
    End If

End Sub

' Combo Scene is used to display the results of combo shots. They are formatted as:
'
'     LINE 1 7x3 Charset
'      SCORE 12x7 digits     X multi (10x18 (or so) charset)
'     LINE 3 5x3 Charset
'

Sub DMDComboScene(line0,line1,line2,combox,combotext,duration,sound)
    Dim ComboScene,HouseFont,ScoreFont,ActionFont,ComboFont,CombotextFont
    if bUseFlexDMD Then
        Set ComboScene = FlexDMD.NewGroup("ComboScene")
        Set HouseFont  = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
        Set ActionFont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
        Set ScoreFont  = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbWhite, 0)
        Set ComboFont  = FlexDMD.NewFont("FlexDMD.Resources.udmd-f12by24.fnt", vbWhite, vbWhite, 0)

        ' Add Text labels
        ComboScene.AddActor FlexDMD.NewLabel("House", HouseFont, "0")
        ComboScene.AddActor FlexDMD.NewLabel("Score", ScoreFont, "0")
        ComboScene.AddActor FlexDMD.NewLabel("Action", ActionFont, "0")
        ComboScene.AddActor FlexDMD.NewLabel("ComboText", ActionFont, "0")
        ComboScene.AddActor FlexDMD.NewLabel("Combo", ComboFont, "0")
        ' Fill in text and align
        With ComboScene.GetLabel("House")
            .Text = line0
            .SetAlignedPosition 40,4,FlexDMD_Align_Center
        End With
        With ComboScene.GetLabel("Score")
            .Text = line1
            .SetAlignedPosition 40,16,FlexDMD_Align_Center
        End With
        With ComboScene.GetLabel("Action")
            .Text = line2
            .SetAlignedPosition 40,25,FlexDMD_Align_Center
        End With
        With ComboScene.GetLabel("ComboText")
            .Text = combotext
            .SetAlignedPosition 104,3,FlexDMD_Align_Center
        End With
        With ComboScene.GetLabel("Combo")
            .Text = combox & "X"
            .SetAlignedPosition 104,19,FlexDMD_Align_Center
        End With
        DMDEnqueueScene ComboScene,1,1000,2000,2500,sound
    Else
        DisplayDMDText line0 & "  " & combotext, line1 & "   " & combox & "X", duration
        PlaySoundVol sound,VolDef
    End If
End Sub

' Choose Scene is used for choosing your house at the beginning of game. Format is
'
'   house    CHOOSE YOUR HOUSE 8x5 charset
'   sigil    <|  House Name 9x5 charset  |>
'             house action button 5x3 charset
'
' Once house is selected, bottom row goes away and top two rows give more detail on house action
Dim ChooseHouseScene
Sub DMDChooseScene1(line0,line1,line2,sigil)    ' sigil is an image name
    Dim sigilimage
    If bUseFlexDMD Then
        If SelectedHouse = 8 Then
            Set ChooseHouseScene = NewSceneWithVideo("choosehouse",sigil)
        Else
            Set ChooseHouseScene = NewSceneWithImage("choosehouse",sigil)
        End if
        ChooseHouseScene.AddActor FlexDMD.NewLabel("action", FlexDMD.NewFont("tiny3by5.fnt", vbWhite, vbWhite, 0) ,line2)
        ChooseHouseScene.GetLabel("action").SetAlignedPosition 77,28,FlexDMD_Align_Center
        Set DefaultScene = ChooseHouseScene
        DMDFlush
    Else
        DisplayDMDText line0, line1, 0
    End if
End Sub

' DMDChooseBattleScene is used for choosing your House Battle(s). Format is
'
'    CHOOSE YOUR BATTLE (5x6 font)
'         HOUSE NAME (8x6 font)
'(optional)  and
' tmr     HOUSE NAME       tmr

Dim CBScene
Sub DMDChooseBattleScene(line0,line1,line2,tmr)
    Dim font,fatfont,smlfont,instscene
    If line0 = "" Then  ' Initial screen
        If bUseFlexDMD Then
            ' Create the instructions scene first
            DMDFlush
            If bBattleInstructionsDone = False Then
                Set instscene = FlexDMD.NewGroup("choosebattleinstr")
                instscene.AddActor FlexDMD.NewLabel("instructions",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0), _
                        "CHOOSE   YOUR   BATTLE" & vbLf & "USE   FLIPPERS   TO" & vbLf & "CHANGE   YOUR   CHOICE" )
                instscene.GetLabel("instructions").SetAlignedPosition 64,16,FlexDMD_Align_Center
                DMDEnqueueScene instscene,0,1500,1500,100,""
            End If

            ' Create ChooseBattle scene
            If IsEmpty(CBScene) Then
                Set CBScene = FlexDMD.NewGroup("choosebattle")
                Set font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0)
                Set fatfont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by5.fnt", vbWhite, vbWhite, 0)
                Set smlfont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
                CBScene.AddActor FlexDMD.NewLabel("choose",font,"CHOOSE YOUR BATTLE")
                CBScene.GetLabel("choose").SetAlignedPosition 64,4,FlexDMD_Align_Center
                CBScene.AddActor FlexDMD.NewLabel("house1",font,"")
                CBScene.GetLabel("house1").SetAlignedPosition 64,20,FlexDMD_Align_Center
                CBScene.AddActor FlexDMD.NewLabel("and",fatfont,"AND")
                With CBScene.GetLabel("and")
                    .SetAlignedPosition 64,20,FlexDMD_Align_Center
                    .Visible = False
                End With
                CBScene.AddActor FlexDMD.NewLabel("house2",font,"")
                With CBScene.GetLabel("house2")
                    .SetAlignedPosition 64,28,FlexDMD_Align_Center
                    .Visible = False
                End With
                CBScene.AddActor FlexDMD.NewLabel("tmrl",smlfont,"")
                CBScene.GetLabel("tmrl").SetAlignedPosition 3,28,FlexDMD_Align_BottomLeft
                CBScene.AddActor FlexDMD.NewLabel("tmrr",smlfont,"")
                CBScene.GetLabel("tmrr").SetAlignedPosition 123,28,FlexDMD_Align_BottomRight
            End If
            Set DefaultScene = CBScene
        Else ' No FlexDMD
            DisplayDMDText "USE FLIPPERS TO","CHOOSE YOUR BATTLE",1500
        End If
    Else ' Update existing screen
        If bUseFlexDMD Then
            FlexDMD.LockRenderThread
            CBScene.GetLabel("house1").Text = line1
            If line2 = "" Then
                If SelectedHouse = Greyjoy And line1 <> "PASS FOR NOW" Then
                    CBScene.GetLabel("house1").SetAlignedPosition 64,12,FlexDMD_Align_Center
                    With CBScene.GetLabel("and")
                        .Text = "IN ALLIANCE WITH"
                        .SetAlignedPosition 64,20,FlexDMD_Align_Center
                        .Visible = True
                    End With
                    With CBScene.GetLabel("house2")
                        .Visible = True
                        .Text = "GREYJOY"
                        .SetAlignedPosition 64,28,FlexDMD_Align_Center
                    End With
                Else
                    CBScene.GetLabel("house1").SetAlignedPosition 64,20,FlexDMD_Align_Center
                    CBScene.GetLabel("and").Visible = False
                    CBScene.GetLabel("house2").Visible = False
                End if
            Else
                CBScene.GetLabel("house1").SetAlignedPosition 64,12,FlexDMD_Align_Center
                CBScene.GetLabel("and").Visible = True
                With CBScene.GetLabel("house2")
                    .Visible = True
                    .Text = line2
                    .SetAlignedPosition 64,28,FlexDMD_Align_Center
                End With
            End If
            CBScene.GetLabel("tmrl").Text = CStr(abs(tmr))
            CBScene.GetLabel("tmrl").SetAlignedPosition 3,28,FlexDMD_Align_BottomLeft
            CBScene.GetLabel("tmrr").Text = CStr(abs(tmr))
            CBScene.GetLabel("tmrr").SetAlignedPosition 123,28,FlexDMD_Align_BottomRight
            FlexDMD.UnlockRenderThread
        Else
            If line2="" Then
                DisplayDMDText line0, line1, 0
            Else
                DisplayDMDText line1, line2, 0
            End if
        End If
    End If
End Sub

' Play the intro video, music, and goals for a House Battle Mode
' We create a single scene consisting of the animation followed by the objective.
' The total scene play length is the same as the music
Dim SceneSoundLengths: SceneSoundLengths = Array(0,5654,5855,7305,5368,5665,5702,5500,5500,5500)  ' Battle sound lengths in 1/1000th's of a second, minus half a second
Sub DMDHouseBattleScene(hse)
    Dim scene,vid,af,blink,hname,delay,vidname,lvl,h

    If hse = 0 Then Exit Sub
    h = hse
    If bUseFlexDMD Then
        hname = HouseToString(h)
        If h = Targaryen Then
            lvl = House(CurrentPlayer).BattleState(Targaryen).TGLevel
            h = h + lvl
            vidname = "got-targaryen"&lvl+1&"battleintro"
        Else
            vidname = "got-"&hname&"battleintro"
        End If
        Set scene = NewSceneWithVideo(hname&"battleintro",vidname)
        Set vid = scene.GetVideo(hname&"battleintrovid")
        If vid is Nothing Then Set vid = scene.getImage(hname&"battleintroimg")
        scene.AddActor FlexDMD.NewLabel("objective",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0),BattleObjectives(h))
        With scene.GetLabel("objective")
            .SetAlignedPosition 64,16, FlexDMD_Align_Center
            .Visible = False
        End With
        ' After x seconds, disable video/image and enable text objective
        Select Case h
            Case Stark: delay=1
            Case Martell: delay=1.9
            Case Lannister,Baratheon: delay=2.5
            Case Else: delay=3
        End Select
        If Not (vid Is Nothing) Then
            DelayActor vid,delay,False
            DelayActor scene.GetLabel("objective"),delay,True
        Else
            scene.GetLabel("objective").Visible = True
        End If
'        DMDEnqueueScene scene,1,SceneSoundLengths(h),SceneSoundLengths(h),10000,"gotfx-"&hname&"battleintro"
        DMDEnqueueScene scene,1,7000,7000,10000,"gotfx-"&hname&"battleintro"
        tmrStartBattleModeScene.Interval = SceneSoundLengths(h)
        tmrStartBattleModeScene.Enabled = 1
    Else
        DisplayDMDText BattleObjectives(h),"",SceneSoundLengths(h)
        PlaySoundVol "gotfx-"&hname&"battleintro",VolDef
    End If
End Sub

' PictoPops Scene is a 3-frame layout with award for each
' pop in each frame. If all 3 awards match, flash text for 1 second
' If no FlexDMD, use short text on one line
Dim PictoScene
Sub DMDPictoScene
    Dim matched:matched=False
    Dim i
    Dim Frame(2)
    Dim pri,mintime:mintime=250:pri=3
    Dim PopsFont
    If BumperVals(0) = BumperVals(1) And BumperVals(0) = BumperVals(2) Then matched=True:mintime=1000:pri=1 'And Flash too
    If bUseFlexDMD Then
        If IsEmpty(PictoScene) Then
            ' Create the scene
            Set PictoScene = FlexDMD.NewGroup("pops")

            ' Create 3 frames. In each frame, put the text of the corresponding bumper award
            Dim poplabel
            For i = 0 to 2
                PictoScene.AddActor FlexDMD.NewFrame("popbox" & i)
                With PictoScene.GetFrame("popbox" & i)
                    .Thickness = 2
                    .SetBounds i*42, 0, 43, 32      ' Each frame is 43W by 32H, and offset by 0, 42, or 84 pixels
                End With
                PictoScene.AddActor FlexDMD.NewLabel("pop"&i, FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0), PictoPops(BumperVals(i))(0))

                ' Place the text in the middle of the frame and let FlexDMD figure it out
                Set poplabel = PictoScene.GetLabel("pop"&i)
                poplabel.SetAlignedPosition i*42+21, 16, FlexDMD_Align_Center
                If matched Then BlinkActor poplabel,0.1,5
            Next
        Else
            ' Existing scene. Update the text
            FlexDMD.LockRenderThread
            For i = 0 to 2
                Set poplabel = PictoScene.GetLabel("pop"&i)
                With poplabel
                    .Text = PictoPops(BumperVals(i))(0)
                    .SetAlignedPosition i*42+21, 16, FlexDMD_Align_Center
                ' Remove any existing action
                    .ClearActions()
                    .Visible = True
                End With
                ' If the bumpers all match, flash the text and keep scene on screen for a second
                If matched Then BlinkActor poplabel,0.1,5:mintime=1000:pri=1
            Next
            FlexDMD.UnlockRenderThread
        End If

        DMDEnqueueScene PictoScene,pri,mintime,1000,300,""
    End If
End Sub

Dim MysteryScene
Sub DMDMysteryAwardScene
    Dim i,combo
    Dim Frame(2),font,line1
    If bUseFlexDMD Then
        If IsEmpty(MysteryScene) Then
            ' Create the scene
            Set font = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
            Set MysteryScene = FlexDMD.NewGroup("myst")

            ' Create 3 frames. In each frame, put the text of the corresponding mystery award
            Dim poplabel
            For i = 0 to 2
                MysteryScene.AddActor FlexDMD.NewFrame("popbox" & i)
                With MysteryScene.GetFrame("popbox" & i)
                    If i = MysterySelected Then .Thickness = 2 Else .Thickness = 1
                    .SetBounds i*42, 0, 43, 32      ' Each frame is 43W by 32H, and offset by 0, 42, or 84 pixels
                End With
                line1 = MysteryAwards(MysteryVals(i))(0) & vbLf & MysteryAwards(MysteryVals(i))(1) & " GOLD"
                Select Case MysteryVals(i)
                    Case 0: line1 = "KEEP MY" & vbLf & CurrentGold & " GOLD"
                    Case 1,18,28
                        combo = ROrbitsw31.UserValue : If combo > 9 Then combo = combo - 10
                        If combo = 0 then combo = 1
                        line1 = CStr(MysteryAwards(MysteryVals(i))(0)*PlayfieldMultiplierVal*combo) & vblf & _
                             "MILLION" & vbLf & "POINTS" & vbLf & MysteryAwards(MysteryVals(i))(1) & " GOLD"
                End Select
                MysteryScene.AddActor FlexDMD.NewLabel("pop"&i, font, line1)

                ' Place the text in the middle of the frame and let FlexDMD figure it out
                Set poplabel = MysteryScene.GetLabel("pop"&i)
                poplabel.SetAlignedPosition i*42+21, 16, FlexDMD_Align_Center
                ' Choice has been made, flash the selected option
                If i = MysterySelected And bMysteryAwardActive = False Then BlinkActor poplabel,0.1,5
            Next
            MysteryScene.AddActor FlexDMD.NewLabel("tmr",font,"10")
            MysteryScene.GetLabel("tmr").SetAlignedPosition 3,2,FlexDMD_Align_TopLeft
            If MATstep = 0 then DMDFlush
            Set DefaultScene = MysteryScene
        Else
            ' Existing scene. Update the text
            FlexDMD.LockRenderThread
            For i = 0 to 2
                With MysteryScene.GetFrame("popbox" & i)
                    If i = MysterySelected Then .Thickness = 2 Else .Thickness = 1
                    .SetBounds i*42, 0, 43, 32      ' Each frame is 43W by 32H, and offset by 0, 42, or 84 pixels
                End With
                line1 = MysteryAwards(MysteryVals(i))(0) & vbLf & MysteryAwards(MysteryVals(i))(1) & " GOLD"
                Select Case MysteryVals(i)
                    Case 0: line1 = "KEEP MY" & vbLf & CurrentGold & " GOLD"
                    Case 1,18,28
                        combo = ROrbitsw31.UserValue : If combo > 9 Then combo = combo - 10
                        If combo = 0 then combo = 1
                        line1 = CStr(MysteryAwards(MysteryVals(i))(0)*PlayfieldMultiplierVal*combo) & vblf & _
                             "MILLION" & vbLf & "POINTS" & vbLf & MysteryAwards(MysteryVals(i))(1) & " GOLD"
                End Select
                Set poplabel = MysteryScene.GetLabel("pop"&i)
                With poplabel
                    .Text = line1
                    .SetAlignedPosition i*42+21, 16, FlexDMD_Align_Center
                ' Remove any existing action
                    .ClearActions()
                    .Visible = True
                End With
                ' Choice has been made, flash the selected option
                If i = MysterySelected And bMysteryAwardActive = False Then BlinkActor poplabel,0.1,5
            Next
            MysteryScene.GetLabel("tmr").Text = CStr(10-MATstep)
            FlexDMD.UnlockRenderThread
            If MATstep = 0 Then DMDFlush : Set DefaultScene = MysteryScene
        End If
    Else
        DisplayDMDText MysteryAwards(MysteryVals(i))(0),MysteryAwards(MysteryVals(i))(1) & " GOLD",0
    End If
End Sub


' Summarize Battle. 2 scenes - animation and then summary. Format:
'
'   Battle Objective
'      SCORE             Combo X
'    "COMPLETED"
'
' Scenes: Stark: Arya stabbing guy on floor

' Battle Mode Target Hit Scene. Layout Might be unique to Stark
Sub DMDStarkBattleScene(house,num,score,line1,line2,just1,just2,sound)
    Dim scene,j3,x1,x2
    if bUseFlexDMD Then
        j3 = just1 + 6  ' Put Score on the same side as Line1 text, but at the bottom
        x1 = 2: x2 = 126
        If just1 = FlexDMD_Align_TopRight Then x1=126:x2=2
        Set scene = NewSceneWithVideo(HouseToString(house)&"hit","got-"&HouseToString(house)&"battlehit"&num)

        scene.AddActor FlexDMD.NewLabel("score",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbWhite, 0),score)
        scene.AddActor FlexDMD.NewLabel("line1", FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0),line1)
        scene.AddActor FlexDMD.NewLabel("line2", FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line2)
        scene.GetLabel("score").SetAlignedPosition x1,30,j3
        scene.GetLabel("line1").SetAlignedPosition x1,5,just1
        scene.GetLabel("line2").SetAlignedPosition x2,30,just2
        DMDEnqueueScene scene,1,2000,2000,1000,sound
    Else
        DisplayDMDText line1,score,2000
        PlaySoundVol sound,VolDef
    End If
End Sub

Dim Castlexy
' map coordinates
' - Kings landing: 170,310
' - Casterly Rock: 36,324
' - Storm's End: 240,370
' - Highgarden: 105,405
' - Pyke: 100,252
' - SunSpear: 228,466
' - Dragonston: 280,320
' - Winterfell: 170,110
Castlexy = Array(170,310,170,110,240,370,36,324,100,252,105,405,228,466,280,320)
Sub DMDScrollITMap(oldcastle,newcastle,score)
    Dim scene,ox,oy,nx,ny
    ox = Castlexy(oldcastle*2)*-1 : oy = Castlexy(oldcastle*2+1)*-1
    nx = Castlexy(newcastle*2)*-1 : ny = Castlexy(newcastle*2+1)*-1
    Set scene = NewSceneWithImage("itmap","got-map")
    scene.SetBounds 0,0,128+(399*2),32+(614*2)  ' Create a canvas twice the size of the image
    With scene.GetImage("itmapimg")
        .SetAlignedPosition ox,oy,FlexDMD_Align_TopLeft
        .AddAction scene.GetImage("itmapimg").ActionFactory().MoveTo(nx,ny,1.2)
    End With
    scene.AddActor FlexDMD.NewLabel("award",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),FormatScore(score))
    With scene.GetLabel("award")
        .SetAlignedPosition 64,16,FlexDMD_Align_Center
        .Visible = 0
    End With
    DelayBlinkActor scene.GetLabel("award"),1.3,0.07,7
    DMDEnqueueScene scene,0,1700,2000,1000,""
End Sub

Dim SpinScene
Dim SpinNum
Sub DMDSpinnerScene(spinval)
    Dim suffix,x,y,locked,tinyfont
    locked=False
    If bUseFlexDMD Then
        Set tinyfont = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
        If IsEmpty(SpinScene) Then
            SpinNum = 0
            Set SpinScene = FlexDMD.NewGroup("spinner")
            SpinScene.SetBounds 0,-8,128,40
        End If
        If Not IsEmpty(DisplayingScene) Then
            If DisplayingScene Is SpinScene Then FlexDMD.LockRenderThread:locked=true
        End If
        If spinval=AccumulatedSpinnerValue Or SpinScene.GetLabel("level") Is Nothing Then ' First spin this scene: clear the scene
            SpinScene.RemoveAll
            SpinScene.AddActor FlexDMD.NewLabel("level",tinyfont,"0")
            SpinScene.AddActor FlexDMD.NewLabel("spin", FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", RGB(167, 165, 165), vbWhite, 0),"SPINNER")
            SpinScene.AddActor FlexDMD.NewLabel("value",tinyfont,"0")
            SpinScene.GetLabel("spin").SetAlignedPosition 64,24, FlexDMD_Align_Center
        End If

        With SpinScene.GetLabel("level")
            .Text = "LEVEL "&SpinnerLevel
            .SetAlignedPosition 64,13, FlexDMD_Align_Center
        End With
        With SpinScene.GetLabel("value")
            .Text = FormatScore(AccumulatedSpinnerValue)
            .SetAlignedPosition 64,36, FlexDMD_Align_Center
        End With
        suffix="K":spinval = int(spinval/1000)
        If spinval >= 1000000 Then suffix="M":spinval = int(spinval/1000)
        SpinScene.AddActor FlexDMD.NewLabel("s"&SpinNum,FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbBlack, 1),spinval&suffix)
        x = RndNbr(100)+13
        y = RndNbr(20) + 16
        With SpinScene.GetLabel("s"&SpinNum)
            .SetAlignedPosition x,y, FlexDMD_Align_BottomLeft
            .AddAction SpinScene.GetLabel("s"&SpinNum).ActionFactory.MoveTo(x,0,0.4)
        End With
        SpinNum = SpinNum + 1
        If locked Then FlexDMD.UnlockRenderThread:locked=False
        DMDEnqueueScene SpinScene,2,500,2000,1500,""
    Else
        DisplayDMDText FormatScore(AccumulatedSpinnerValue),spinval,100
    End if
End Sub

Sub DMDCreateBWMBScoreScene
    If bUseFlexDMD = False Then Exit Sub
    Dim scene,i,ComboFont
    Set scene = FlexDMD.NewGroup("bwmb")
    scene.AddActor FlexDMD.NewLabel("line1", FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),"BLACKWATER"&vbLf&"PHASE 1")
    scene.AddActor FlexDMD.NewLabel("Score", FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),Score(CurrentPlayer))
    scene.GetLabel("Score").SetAlignedPosition 127,0,FlexDMD_Align_TopRight
    scene.AddActor FlexDMD.NewLabel("obj", FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0),"SHOOT GREEN JACKPOTS")
    scene.GetLabel("obj").SetAlignedPosition 64,18,FlexDMD_Align_Center
    BlinkActor scene.GetLabel("obj"),200,9999
    scene.AddActor FlexDMD.NewLabel("tmr1", FlexDMD.NewFont("udmd-f11by18.fnt", vbWhite, vbWhite, 0),"20")
    With scene.GetLabel("tmr1")
        .SetAlignedPosition 127,23,FlexDMD_Align_BottomRight
        .Visible = 0
    End With
    ' Combo Multipliers
    Set ComboFont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
    For i = 1 to 5
        scene.AddActor FlexDMD.NewLabel("combo"&i, ComboFont, "0")
    Next
    DMDSetAlternateScoreScene scene,265     ' Score, Combos, SJP timer
    SetGameTimer tmrUpdateBattleMode,5
End Sub

Sub DMDCreateCastleMBScoreScene
    If bUseFlexDMD = False Then Exit Sub
    Dim scene,i,font,combofont
    Set scene = FlexDMD.NewGroup("cmb")
    Set font = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
    scene.AddActor FlexDMD.NewLabel("line1", font,"CASTLE MULTIBALL")
    scene.AddActor FlexDMD.NewLabel("Score", font,Score(CurrentPlayer))
    scene.GetLabel("Score").SetAlignedPosition 80,0,FlexDMD_Align_TopLeft
    scene.AddActor FlexDMD.NewLabel("obj", font,"SHOOT CASTLE DRAGON SHOTS"&vbLf&"SUPER JACKPOT VALUE")
    scene.GetLabel("obj").SetAlignedPosition 64,13,FlexDMD_Align_Center
    scene.AddActor FlexDMD.NewLabel("sjpval", font,FormatScore(20000000))
    scene.GetLabel("sjpval").SetAlignedPosition 64,22,FlexDMD_Align_Center
    ' Combo Multipliers
    Set ComboFont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
    For i = 1 to 5
        scene.AddActor FlexDMD.NewLabel("combo"&i, ComboFont, "0")
    Next
    DMDSetAlternateScoreScene scene,9     ' Score, Combos
    SetGameTimer tmrUpdateBattleMode,5
End Sub

Sub DMDCreateWallMBScoreScene
    If bUseFlexDMD = False Then Exit Sub
    Dim scene,i,scorefont,combofont
    Set combofont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
    Set scene = NewSceneWithVideo("wallscore","got-wallmbbackground")
    scene.AddActor FlexDMD.NewLabel("line1", combofont,"BATTLE FOR THE WALL")
    scene.GetLabel("line1").SetAlignedPosition 64,2,FlexDMD_Align_Top

    scene.AddActor FlexDMD.NewLabel("Score",FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0),FormatScore(Score(CurrentPlayer)))
    scene.GetLabel("Score").SetAlignedPosition 64,9,FlexDMD_Align_Top

    scene.AddActor FlexDMD.NewLabel("obj", combofont,"3 MORE RAMPS")
    scene.GetLabel("obj").SetAlignedPosition 64,19,FlexDMD_Align_Top

    For i = 1 to 5
        scene.AddActor FlexDMD.NewLabel("combo"&i, ComboFont, "0")
    Next
    DMDSetAlternateScoreScene scene,9     ' Score, Combos
    SetGameTimer tmrUpdateBattleMode,5
End Sub

Const WallMB = 1
Const HotkMB = 2
Const ITMB = 3
Sub DMDCreateReadyScene(MB)
    If bUseFlexDMD = False Then Exit Sub
    Dim scene,i,scorefont,combofont,line1,obj
    Set combofont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
    Set scene = FlexDMD.NewGroup("ready")
    Select Case MB
        Case WallMB
            line1 = "WALL MULTIBALL READY"
            obj = "SHOOT THE RIGHT ORBIT"
        Case HotkMB
            line1 = "HAND OF THE KING READY"
            obj = "SHOOT THE LEFT RAMP"
        Case ITMB
            line1 = "IRON THRONE READY"
            obj = "SHOOT THE LEFT RAMP"
    End Select
    scene.AddActor FlexDMD.NewLabel("line1", combofont,line1)
    scene.GetLabel("line1").SetAlignedPosition 64,0,FlexDMD_Align_Top

    scene.AddActor FlexDMD.NewLabel("Score",FlexDMD.NewFont("tiny3by5.fnt", vbWhite, vbWhite, 0),FormatScore(Score(CurrentPlayer)))
    scene.GetLabel("Score").SetAlignedPosition 64,7,FlexDMD_Align_Top

    scene.AddActor FlexDMD.NewLabel("obj", FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0),obj)
    scene.GetLabel("obj").SetAlignedPosition 64,14,FlexDMD_Align_Top

    For i = 1 to 5
        scene.AddActor FlexDMD.NewLabel("combo"&i, ComboFont, "0")
    Next
    DMDSetAlternateScoreScene scene,9     ' Score, combos
End Sub

Sub DMDCreateWHCMBScoreScene
    If bUseFlexDMD = False Then Exit Sub
    Dim scene,i,scorefont
    'TODO: Change the font color to black once it has a transparent background
    Set scorefont = FlexDMD.NewFont("tiny3by5.fnt", vbWhite, vbBlack, 0)
    Set scene = NewSceneWithVideo("whcscore","got-whcbg0")

    scene.AddActor FlexDMD.NewLabel("Score",scorefont,FormatScore(Score(CurrentPlayer)))
    scene.GetLabel("Score").SetAlignedPosition 123,31,FlexDMD_Align_BottomRight

    DMDSetAlternateScoreScene scene,1     ' Just score
End Sub

Sub DMDCreateWHCMBScore2Scene
    If bUseFlexDMD = False Then Exit Sub
    Dim scene,i,scorefont,combofont
    'TODO: Change the font color to black once it has a transparent background
    Set scorefont = FlexDMD.NewFont("tiny3by5.fnt", vbWhite, vbBlack, 0)
    Set scene = NewSceneWithVideo("whcscore2","got-whcbg1")

    scene.AddActor FlexDMD.NewLabel("Score",scorefont,FormatScore(Score(CurrentPlayer)))
    scene.GetLabel("Score").SetAlignedPosition 4,31,FlexDMD_Align_BottomLeft

    scene.AddActor FlexDMD.NewLabel("obj", scorefont,"5 SHOTS NEEDED")
    scene.GetLabel("obj").SetAlignedPosition 1,1,FlexDMD_Align_TopLeft

    scene.AddActor FlexDMD.NewLabel("obj2", FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0),"")
    With scene.GetLabel("obj2")
        .SetAlignedPosition 64,14,FlexDMD_Align_Center
        .Visible = 0
    End With

    DMDSetAlternateScoreScene scene,1     ' Just score
End Sub

Sub DMDCreateHotkMBScoreScene
    If bUseFlexDMD = False Then Exit Sub
    Dim scene,scene2,i,font,line
    Set scene = FlexDMD.NewGroup("hotkscore")
    Set font = FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0)

    scene.AddActor FlexDMD.NewLabel("hotk1",font,"HAND OF THE KING")
    scene.GetLabel("hotk1").SetAlignedPosition 127,0,FlexDMD_Align_TopRight
    scene.AddActor FlexDMD.NewLabel("Score",font,FormatScore(Score(CurrentPlayer)))
    scene.GetLabel("Score").SetAlignedPosition 54,0,FlexDMD_Align_TopRight

    Set scene2 = FlexDMD.NewGroup("state2")
    ' Only used for Stark bonus round
    scene2.AddActor FlexDMD.NewLabel("bonus",font,"BONUS ROUND"&vbLf&"FREE SHOTS FOR 20 SECS")
    scene2.GetLabel("bonus").SetAlignedPosition 54,16,FlexDMD_Align_Center

    Set font = FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0)

    ' Only used for Stark bonus round
    scene2.AddActor FlexDMD.NewLabel("tmr1",font,"20")
    scene2.GetLabel("tmr1").SetAlignedPosition 127,13,FlexDMD_Align_TopRight

    scene.AddActor scene2
    scene.GetGroup("state2").visible = 0

    Set scene2 = FlexDMD.NewGroup("state0")
    scene2.AddActor FlexDMD.NewLabel("obj",font,"COMPLETE ALL SHOTS")
    scene2.GetLabel("obj").SetAlignedPosition 64,9,FlexDMD_Align_Top
    BlinkActor scene2.GetLabel("obj"),0.1,9999

    Set font = FlexDMD.NewFont("tiny3by5.fnt",vbWhite,vbWhite,0)
    scene2.AddActor FlexDMD.NewLabel("obj2",font,"HURRY UP: ")
    scene2.GetLabel("obj2").SetAlignedPosition 0,19,FlexDMD_Align_TopLeft
    scene2.AddActor FlexDMD.NewLabel("HurryUp",font,"00")
    scene2.GetLabel("HurryUp").SetAlignedPosition 36,19,FlexDMD_Align_TopLeft

    If (House(CurrentPlayer).HotkMask And 2^Greyjoy) > 0 Then line="2 SETS" Else line="3 SETS"
    scene2.AddActor FlexDMD.NewLabel("sets",font,line)
    scene2.GetLabel("sets").SetAlignedPosition 127,19,FlexDMD_Align_TopRight

    scene.AddActor scene2
    scene.GetGroup("state0").visible = 0

    For i = 1 to 5
        scene.AddActor FlexDMD.NewLabel("combo"&i, FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0), "0")
    Next
    DMDSetAlternateScoreScene scene,265     ' Score, combos, SJP timer
End Sub

Sub DMDPlayITHitScene(castle,score)
    Dim scene,x,y
    Set scene = DMDITScoreScene(castle,score)
    BlinkActor scene.GetLabel("Score"),0.1,5
    DMDEnqueueScene scene,0,500,1500,1000,""
End Sub

Sub DMDCreateITMBScoreScene(State,Castle,CastlesLeft)
    Dim Scene,font,i
    Set font = FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0)
    If State = 0 Then
        Set Scene = FlexDMD.NewGroup("itscore")
        Scene.AddActor FlexDMD.NewLabel("line1", FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),"CHOOSE A CASTLE")
        Scene.GetLabel("line1").SetAlignedPosition 64,0,FlexDMD_Align_Top
        scene.AddActor FlexDMD.NewLabel("Score",font,FormatScore(Score(CurrentPlayer)))
        scene.GetLabel("Score").SetAlignedPosition 0,9,FlexDMD_Align_TopLeft
        scene.AddActor FlexDMD.NewLabel("left",font,"CASTLES LEFT: "&CastlesLeft)
        scene.GetLabel("left").SetAlignedPosition 127,11,FlexDMD_Align_TopRight
        For i = 1 to 5
            scene.AddActor FlexDMD.NewLabel("combo"&i, FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0), "0")
        Next
        DMDSetAlternateScoreScene Scene,9
    Else
        Set Scene = DMDITScoreScene(Castle,Score(CurrentPlayer))
        DMDSetAlternateScoreScene Scene,1
    End If
End Sub

Function DMDITScoreScene(Castle,score)
    Dim Scene,x,y
    If Castle > 0 And Castle < 8  Then
        Set Scene = NewSceneWithImage("itscore","got-it"&HouseToString(Castle)&"castle")
    Else
        Set Scene = NewSceneWithImage("itscore","got-itcastle"&Castle)
    End If
    Select Case Castle
        Case Stark,Greyjoy,Martell,Targaryen: x = 5 : y=10
        Case Baratheon: x=5:y=22
        Case Lannister: x=64:y=22
        Case Tyrell: x=64:y=16
    End Select
    scene.AddActor FlexDMD.NewLabel("Score",FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0),FormatScore(score))
    scene.GetLabel("Score").SetAlignedPosition x,y,FlexDMD_Align_TopLeft
    Set DMDITScoreScene = Scene
End Function

Dim mmscores(8)
Dim mmfonts(4)
Dim MMScene
Sub DMDCreateMadnessMBScoreScene
    Dim Scene,i,font,x,y
    Set font = FlexDMD.NewFont("tiny3by5.fnt",vbWhite,vbWhite,0)
    Set mmfonts(0) = font
    Set mmfonts(1) = FlexDMD.NewFont("tiny3by5.fnt",&hCCCCCC,vbblack,0)
    Set mmfonts(2) = FlexDMD.NewFont("tiny3by5.fnt",&h999999,vbblack,0)
    Set mmfonts(3) = FlexDMD.NewFont("tiny3by5.fnt",&h666666,vbblack,0)
    Set mmfonts(4) = FlexDMD.NewFont("tiny3by5.fnt",&h333333,vbblack,0)
    Set Scene = NewSceneWithVideo("mmscore","got-winterstorm")
    x = 10
    For i = 0 to 7
        if i > 3 Then x = 96
        y = (i MOD 4)*8 + 2
        mmscores(i) = 0
        scene.AddActor FlexDMD.NewLabel("value"&i,font,FormatScore(100000))
        With scene.GetLabel("value"&i)
            .visible = 0
            .SetAlignedPosition x+(i MOD 2)*4, y, FlexDMD_Align_TopLeft
        End With
    Next
    scene.AddActor FlexDMD.NewLabel("mmtop",FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0),"MIDNIGHT MADNESS")
    scene.GetLabel("mmtop").SetAlignedPosition 64,4,FlexDMD_Align_Center
    scene.AddActor FlexDMD.NewLabel("Score",FlexDMD.NewFont("udmd-f7by10.fnt",vbWhite,vbWhite,0),FormatScore(Score(CurrentPlayer)))
    scene.GetLabel("Score").SetAlignedPosition 64,16,FlexDMD_Align_Center
    scene.AddActor FlexDMD.NewLabel("mmtotal",font,0)
    scene.GetLabel("mmtotal").SetAlignedPosition 64,26,FlexDMD_Align_Center
    scene.AddActor FlexDMD.NewLabel("stickem",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),"STICK EM WITH"&vbLf&"THE POINTY END")
    With scene.GetLabel("stickem")
        .SetAlignedPosition 64,10,FlexDMD_Align_Center
        .visible = 0
    End with
    scene.AddActor FlexDMD.NewLabel("stickemscore",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0),FormatScore(200000))
    With scene.GetLabel("stickemscore")
        .SetAlignedPosition 64,24,FlexDMD_Align_Center
        .visible = 0
    End with
    DMDSetAlternateScoreScene Scene,1
    Set MMScene = scene
End Sub


' Update the Midnight Madness DMD score display with a new hit
' There are 8 slots on the display for scores that fade out. Pick one at random that isn't in use
' Randomly, if not currently speaking, say either "HODOR" or "Stick em..". If saying Stick Em then
' flip the DMD to display the Stick Em message. Set the speech timer to prevent talking over
' previous speeches
Sub DMDUpdateMadnessScoreScene(value)
    Dim i,h,line,lbl,c,flip

    c=0 : flip = 0
    Do
        i = RndNbr(8)-1
        c=c+1
    Loop While (mmscores(i) <> 0 And c < 20)
    If c >=20 Then exit Sub ' Give up looking for a free slot
    if tmrMadnessSpeech.Enabled = 0 Then
        If value = 1 Then
            flip = -1
        Else
            h = RndNbr(8)
            if h = 1 Then
                PlaySoundVol "say-mmstickem"&RndNbr(2),VolCallout
                flip=1
                tmrMadnessSpeech.Interval=1500 : tmrMadnessSpeech.Enabled=1
            ElseIf h > 4 Then
                mmscores(i) = 1
                tmrMadnessSpeech.Interval=750 : tmrMadnessSpeech.Enabled=1
                line="HODOR"
                PlaySoundVol "say-mmhodor"&RndNbr(7),VolCallout
            End If
        End If
    Else
        mmscores(i) = 5
        line = FormatScore(value)
    End If
    FlexDMD.LockRenderThread
        if flip <> 0 Then
            Select Case flip
                Case 1
                    MMscene.GetLabel("mmtop").visible=0
                    MMscene.GetLabel("mmtotal").visible=0
                    MMscene.GetLabel("Score").visible=0
                    MMscene.GetLabel("stickem").visible=1
                    With MMscene.GetLabel("stickemscore")
                        .Visible = 1
                        .Text = FormatScore(value)
                    End With
                Case -1
                    MMscene.GetLabel("mmtop").visible=1
                    MMscene.GetLabel("mmtotal").visible=1
                    MMscene.GetLabel("Score").visible=1
                    MMscene.GetLabel("stickem").visible=0
                    MMscene.GetLabel("stickemscore").visible=0
            End Select
        Else
            Set lbl = MMScene.GetLabel("value"&i)
            If Not lbl is Nothing Then
                With lbl
                    .visible = 1
                    .text = line
                    .font = mmfonts(0)
                    debug.print "Updating mmscore"&i&" to "&line
                End With
            End If
            With MMScene.GetLabel("mmtotal")
                .Text = FormatScore(MMTotalScore)
                .SetAlignedPosition 64,26,FlexDMD_Align_Center
            End With
            MMScene.GetLabel("Score").SetAlignedPosition 64,16,FlexDMD_Align_Center
        End if
    FlexDMD.UnlockRenderThread
    if tmrMMDMDUpdate.Enabled = 0 Then
        tmrMMDMDUpdate.Interval = 100
        tmrMMDMDUpdate.Enabled = 1
    End If
End Sub

Sub tmrMMDMDUpdate_Timer
    Dim i,t
    t=False
    tmrMMDMDUpdate.Enabled=0
    FlexDMD.LockRenderThread
    For i = 0 to 7
        If mmscores(i) <> 0 Then
            If mmscores(i) = 4 Or mmscores(i) = 11 Then
                MMScene.GetLabel("value"&i).visible=0
                mmscores(i) = 0
            Else
                t=true
                mmscores(i) = mmscores(i)+1
                If mmscores(i) > 7 Then
                    debug.print "mmscore"&i& " text: "&MMScene.GetLabel("value"&i).Text
                    MMScene.GetLabel("value"&i).Font = mmfonts(mmscores(i)-7)
                End If
            End If
        End If
    Next
    FlexDMD.UnlockRenderThread
    If t Then tmrMMDMDUpdate.Enabled = 1
End Sub


' Set up an alternate score scene for battle mode. If two houses are stacked
' for battle, create a split screen scene
Sub DMDCreateAlternateScoreScene(h1,h2)
    Dim scene,scene1,scene2,vid,mask,i
    if bUseFlexDMD = False Then Exit Sub
    If h2 <> 0 Then
        Set scene = NewSceneWithVideo("battle","got-"&HouseToString(h2)&"battlesigil")
    ElseIf h1 = Baratheon or h1 = Greyjoy Then
        Set scene = NewSceneWithVideo("battle","got-"&HouseToString(h1)&"battlesigil")
    Else
        Set scene = NewSceneWithVideo("battle","got-"&HouseToString(h1)&"battleprogress")
    End If
    Set vid = scene.GetVideo("battlevid")
    If Not (vid Is Nothing) Then vid.SetAlignedPosition 127,0,FlexDMD_Align_TopRight
    If h2 <> 0 Then
        Set vid = FlexDMD.NewVideo("battlevid2","got-" & HouseToString(h1) & "battlesigil.gif")
        If Not (vid is Nothing) Then
            scene.AddActor vid
            Set vid = scene.GetVideo("battlevid2")
            vid.SetAlignedPosition 63,0,FlexDMD_Align_TopRight
        End If
        Set scene1 = FlexDMD.NewGroup(HouseToString(h1))
        Set scene2 = FlexDMD.NewGroup(HouseToString(h2))
        House(CurrentPlayer).BattleState(h1).CreateSmallBattleProgressScene scene1,1
        House(CurrentPlayer).BattleState(h2).CreateSmallBattleProgressScene scene2,2
        scene1.SetAlignedPosition 0,0,FlexDMD_Align_TopLeft
        scene2.SetAlignedPosition 64,0,FlexDMD_Align_TopLeft
        scene.AddActor scene1
        scene.AddActor scene2
        For i = 1 to 5
            scene.AddActor FlexDMD.NewLabel("combo"&i, FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbBlack, 1), "1X")
        Next
        mask = 104  ' combos, tmr1, tmr2
        If h1 = Targaryen or h2 = Targaryen Then
            mask = mask Or 512 ' TGHurryUp
            If h1=Targaryen Then mask = mask And 223
            If h2=Targaryen Then mask = mask And 191
        End If
        If h1 = Martell or h2 = Martell Then mask = mask Or 144 'HurryUp, tmr3
    Else
        Set scene1 = FlexDMD.NewGroup(HouseToString(h1))
        House(CurrentPlayer).BattleState(h1).CreateBattleProgressScene scene1
        scene1.SetAlignedPosition 0,0,FlexDMD_Align_TopLeft
        scene.AddActor scene1
        mask = 41   ' score, combos, tmr1
        If h1 = Martell Then mask = 184
        If h1 = Targaryen Then mask = 521
    End If
    SetGameTimer tmrUpdateBattleMode,5
    DMDSetAlternateScoreScene scene,mask
End Sub

' We support multiple score "scenes", depending on what mode the table is in. Not all modes
' support all fields, so define a SceneMask that decides which fields need to be updated
'  bit   data (Label name)
'   0    Score
'   1    Ball
'   2    Credits
'   3    combo1 thru 5
'   4    HurryUp
'   5    BattleTimer1 (tmr1)
'   6    BattleTimer2 (tmr2)
'   7    MartellBattleTimer (tmr3)
'   8    SJP Timer (tmr1)
'   9    Targaryen HurryUp
'
' "scene" is a pre-created scene with all of the proper text labels already created. There MUST be a label
' corresponding with every bit set in the scenemask
Sub DMDSetAlternateScoreScene(scene,mask)
    bAlternateScoreScene = True
    Set ScoreScene = scene
    AlternateScoreSceneMask = mask
    DMDLocalScore
End Sub

' Set Score scene back to default for regular play
Sub DMDResetScoreScene
    bAlternateScoreScene = False
    If Not IsEmpty(DisplayingScene) Then
        If DisplayingScene Is ScoreScene Then DMDClearQueue
    End if
    ScoreScene = Empty
    AlternateScoreSceneMask = 0
    DMDLocalScore
End Sub

Sub AddNewPlayerWhileBIP ' 1.1.17
  'if adding new players, update the score scene if it's currently displayed
    If Not IsEmpty(DisplayingScene) Then
        If DisplayingScene Is ScoreScene Then
      DMDResetScoreScene
    end if
    End if
End Sub

Dim ScoreScene,bAlternateScoreScene,AlternateScoreSceneMask,remaining,StatusFont,ComboFont,ComboFontVar,ScoreFont,CreditFont,tinyfont
Sub DMDLocalScore
    Dim i,j,x,y,font
    If bUseFlexDMD Then
        If IsEmpty(ScoreScene) And bAlternateScoreScene = False Then
            Set ScoreScene = FlexDMD.NewGroup("ScoreScene")

      Set CreditFont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
      Set StatusFont = FlexDMD.NewFont("stern-tinyfont-3x5.fnt",vbWhite,vbWhite,0)
      Set ComboFont = FlexDMD.NewFont("got-combotext.fnt", &h999999, &h999999, 0)
      Set ComboFontVar = FlexDMD.NewFont("got-combotext-var.fnt", &h999999, &h999999, 0)

            If Score(CurrentPlayer) < 1000000000 And PlayersPlayingGame < 3 Then
        font = "stern-scorefont-7x12.fnt"
      Elseif Score(CurrentPlayer) >= 1000000000 and PlayersPlayingGame < 3 Then
        font = "stern-scorefont-6x12.fnt"
      Elseif PlayersPlayingGame >= 3 then
        font = "stern-scorefont-6x8.fnt"
      End if
            Set ScoreFont = FlexDMD.NewFont(font, vbWhite, vbWhite, 0)
            Set tinyfont = FlexDMD.NewFont("stern-scorefont-4x7.fnt",vbWhite,vbWhite,0)
            ' Score text
            For i = 1 to PlayersPlayingGame
                If i = CurrentPlayer Then
                    j=""
                    ScoreScene.AddActor FlexDMD.NewLabel("Score", ScoreFont, FormatScore(Score(i)))
                Else
                    j=i
                    ScoreScene.AddActor FlexDMD.NewLabel("Score"&i, tinyfont, FormatScore(Score(i)))
                End If
                x = 80 : y = 0
                if i > 2 Then
                    y = 9
                End If
                'If (i MOD 2) = 0 Then
                '    x = 127
                'ElseIf (CurrentPlayer MOD 2) = 0 Then
                '    x = 46
                'End if
                'ScoreScene.GetLabel("Score"&j).SetAlignedPosition x,y, FlexDMD_Align_TopRight
        If i <> CurrentPlayer and PlayersPlayingGame > 2 then
          Select Case i
            Case 1: ScoreScene.GetLabel("Score"&j).SetAlignedPosition 0,y, FlexDMD_Align_TopLeft
            Case 2: ScoreScene.GetLabel("Score"&j).SetAlignedPosition 127,y, FlexDMD_Align_TopRight
            Case 3: ScoreScene.GetLabel("Score"&j).SetAlignedPosition 0,y, FlexDMD_Align_TopLeft
            Case 4: ScoreScene.GetLabel("Score"&j).SetAlignedPosition 127,y, FlexDMD_Align_TopRight
          end Select
        elseif i = CurrentPlayer and PlayersPlayingGame > 2 then
          Select Case i
            case 1: ScoreScene.GetLabel("Score").SetAlignedPosition 63,y, FlexDMD_Align_TopRight
            case 2: ScoreScene.GetLabel("Score").SetAlignedPosition 64,y, FlexDMD_Align_TopLeft
            case 3: ScoreScene.GetLabel("Score").SetAlignedPosition 63,y-1, FlexDMD_Align_TopRight
            case 4: ScoreScene.GetLabel("Score").SetAlignedPosition 64,y-1, FlexDMD_Align_TopLeft
          end select
        elseif PlayersPlayingGame < 3 and i <> CurrentPlayer then
          if i = 1 then
            ScoreScene.GetLabel("Score"&j).SetAlignedPosition 0,y, FlexDMD_Align_TopLeft
          elseif i = 2 then
            ScoreScene.GetLabel("Score"&j).SetAlignedPosition 127,y, FlexDMD_Align_TopRight
          end if
        elseif PlayersPlayingGame < 3 and i = CurrentPlayer then
          if i = 1 then
            ScoreScene.GetLabel("Score").SetAlignedPosition 76,y, FlexDMD_Align_TopRight
          elseif i = 2 then
            ScoreScene.GetLabel("Score").SetAlignedPosition 51,y, FlexDMD_Align_TopLeft
          end if
        end if
            Next
            ' Ball, credits
            ScoreScene.AddActor FlexDMD.NewLabel("Ball", CreditFont, "BALL 1")
            ScoreScene.AddActor FlexDMD.NewLabel("Credit", CreditFont, "CREDITS 0")
            If DMDStd(kDMDStd_FreePlay) Then ScoreScene.GetLabel("Credit").Text = "Free Play"
            ' Align them
            ScoreScene.GetLabel("Ball").SetAlignedPosition 32,25, FlexDMD_Align_Bottom
            ScoreScene.GetLabel("Credit").SetAlignedPosition 96,25, FlexDMD_Align_Bottom
            ' Divider
            ScoreScene.AddActor FlexDMD.NewFrame("HSeparator")
            ScoreScene.GetFrame("HSeparator").Thickness = 1
            ScoreScene.GetFrame("HSeparator").SetBounds 0, 24, 128,1
      ScoreScene.GetFrame("HSeparator").BorderColor = &h666666
            ' Combo Multipliers
            For i = 1 to 5
                if (ComboMultiplier(i)*PlayfieldMultiplierVal) <= 2 then
          ScoreScene.AddActor FlexDMD.NewLabel("combo"&i, ComboFont, "0")
        Else
          ScoreScene.AddActor FlexDMD.NewLabel("combo"&i, ComboFontVar, "0")
        end if
            Next
        End If

        FlexDMD.LockRenderThread
        ' Update fields
        If bAlternateScoreScene = False or (AlternateScoreSceneMask And 1) = 1 Then ScoreScene.GetLabel("Score").Text = FormatScore(Score(CurrentPlayer))
        If bAlternateScoreScene = False or (AlternateScoreSceneMask And 2) = 2 Then ScoreScene.GetLabel("Ball").Text = "BALL " & CStr(BallsPerGame - BallsRemaining(CurrentPlayer) + 1)
        If DMDStd(kDMDStd_FreePlay) = False And (bAlternateScoreScene = False or (AlternateScoreSceneMask And 4) = 4) Then
            With ScoreScene.GetLabel("Credit")
                .Text = "CREDITS " & CStr(Credits)
                .SetAlignedPosition 96,25, FlexDMD_Align_Bottom
            End With
        End If

        ' Update score position
        If bAlternateScoreScene = False Then
            x = 80 : y = 0
            If CurrentPlayer = 2 Or CurrentPlayer = 4 Then x = 127
            If CurrentPlayer > 2 Then y = 10
            ScoreScene.GetLabel("Score").SetAlignedPosition x,y, FlexDMD_Align_TopRight
        End If
        If bAlternateScoreScene = False or (AlternateScoreSceneMask And 8) = 8 Then
            ' Update combo x
            Dim t : t = False
            If (TimerFlags(tmrTargaryenFreeze) And 1) = 1 Then t = True
            For i = 1 to 5
                With ScoreScene.GetLabel("combo"&i)
                    If t And i=1 Then
            .font = StatusFont
                        .Text = "TIMERS FROZEN: " & Int((TimerTimestamp(tmrTargaryenFreeze)-GameTimeStamp)/10)
            .SetAlignedPosition 1,33,FlexDMD_Align_BottomLeft
            'add or reset the timer to ensure dmd updates in 1 second for timer
            SetGameTimer tmrTargaryenFreezeUpdateDMD,10
                    Else
            if (ComboMultiplier(i)*PlayfieldMultiplierVal) <= 2 then
              .font = ComboFont
            Else
              .font = ComboFontVar
            end if
            .SetAlignedPosition ((i-1)*25)+14,32,FlexDMD_Align_Bottom
                        .Text = (ComboMultiplier(i)*PlayfieldMultiplierVal)&"X"
                    End If

                    If t And i > 1 Then .Visible = 0 Else .Visible = 1
                End With
            Next
        End If

        ' Update special battlemode fields
        If bAlternateScoreScene Then
            If (AlternateScoreSceneMask And 16) = 16 Then ScoreScene.GetLabel("HurryUp").Text = FormatScore(HurryUpValue)
            If (AlternateScoreSceneMask And 512)=512 Then ScoreScene.GetLabel("TGHurryUp").Text = FormatScore(TGHurryUpValue)
            If (AlternateScoreSceneMask And 32) = 32 Then ScoreScene.GetLabel("tmr1").Text = Int((TimerTimestamp(tmrBattleMode1) - GameTimeStamp)/10)
            If (AlternateScoreSceneMask And 64) = 64 Then ScoreScene.GetLabel("tmr2").Text = Int((TimerTimestamp(tmrBattleMode2) - GameTimeStamp)/10)
            If (AlternateScoreSceneMask And 128)=128 Then ScoreScene.GetLabel("tmr3").Text = Int((TimerTimestamp(tmrMartellBattle) - GameTimeStamp)/10)
            If (AlternateScoreSceneMask And 256)=256 Then
                Dim newremaining
                newremaining = Int((TimerTimestamp(tmrBlackwaterSJP) - GameTimeStamp)/10)
                if newremaining >= 0 Then
                    If Not IsEmpty(DisplayingScene) Then
                        If newremaining <> remaining And newremaining < 11 And bBWMultiballActive = True And ScoreScene is DisplayingScene Then _
                            PlaySoundVol "say-hound"&newremaining,VolCallout
                    End if
                    ScoreScene.GetLabel("tmr1").Text = newremaining : remaining = newremaining
                End if
            End If
        End If
        FlexDMD.UnlockRenderThread
        Set DefaultScene = ScoreScene
    Else
        DisplayDMDText "",FormatScore(Score(CurrentPlayer)),0
    End If
End Sub

Sub DMDDoMatchScene(m)
    Dim scene

    If bUseFlexDMD Then
        Set scene = FlexDMD.NewGroup("match")
        scene.AddActor FlexDMD.NewImage("bkgr","got-blankgrey.png")
        scene.AddActor FlexDMD.NewLabel("match1",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbBlack, 1),"MATCH")
        scene.GetLabel("match1").SetAlignedPosition 84,16,FlexDMD_Align_CENTER
        If m = 0 Then m = "00"
        scene.AddActor FlexDMD.NewLabel("match2",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbBlack, 1),m)
        With scene.GetLabel("match2")
            .SetAlignedPosition 84,16,FlexDMD_Align_CENTER
            .Visible = 0
        End With
        DelayActor scene.GetLabel("match1"),2.1,False
        DelayActor scene.GetLabel("match2"),2.1,True
        scene.AddActor FlexDMD.NewLabel("Score1", FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),FormatScore(Score(1)))
        scene.GetLabel("Score1").SetAlignedPosition 1,1,FlexDMD_Align_TopLeft
        If PlayersPlayingGame > 1 Then
            scene.AddActor FlexDMD.NewLabel("Score2", FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),FormatScore(Score(2)))
            scene.GetLabel("Score2").SetAlignedPosition 1,30,FlexDMD_Align_BottomLeft
        End If
        If PlayersPlayingGame > 2 Then
            scene.AddActor FlexDMD.NewLabel("Score3", FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),FormatScore(Score(3)))
            scene.GetLabel("Score3").SetAlignedPosition 126,1,FlexDMD_Align_TopRight
        End If
        If PlayersPlayingGame > 3 Then
            scene.AddActor FlexDMD.NewLabel("Score4", FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),FormatScore(Score(4)))
            scene.GetLabel("Score4").SetAlignedPosition 126,30,FlexDMD_Align_BottomRight
        End If

        scene.AddActor NewSceneFromImageSequence("img1","match",280,60,0,0)
        DMDFlush
        DMDEnqueueScene scene,0,9000,9000,4000,"gotfx-match"
    Else
        DisplayDMDText "","MATCH "&m,4000
    End If
End Sub

Sub GameDoDMDHighScoreScene(scoremask)
    Dim scene,i,scenenum,actor,ondly,offdly
    Set scene = FlexDMD.NewGroup("hs")
    ondly = 0 : offdly = 4
    For i = 0 to 13 ' number of highscore slots
        if (scoremask And (2^i)) > 0 Then ' This is one of the player's high scores
            scenenum = scenenum+1
            scene.AddActor FlexDMD.NewLabel("scname"&scenenum, FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0), ChampionNames(i) )
            Set actor = scene.GetLabel("scname"&scenenum)
            With actor
                .SetAlignedPosition 64,3,FlexDMD_Align_Center
                .Visible = 0
            End With
            FlashActor actor,ondly,offdly
            scene.AddActor FlexDMD.NewLabel("scscore"&scenenum, FlexDMD.NewFont("udmd-f7by10.fnt",vbWhite,vbWhite,0),FormatScore(HighScore(i)))
            Set actor = scene.GetLabel("scscore"&scenenum)
            With actor
                .SetAlignedPosition 64,16,FlexDMD_Align_Center
                .Visible=0
            End With
            FlashActor actor,ondly,offdly
            ondly = ondly+offdly
            offdly = 2
        End If
    Next
    DMDFlush
    DMDDisplayScene scene
    vpmTimer.AddTimer 2000+2000*scenenum-100,"StopSound ""got-track-hstd"" : hsbModeActive=0 : DMDBlank : tmrDMDUpdate.Enabled = True : EndOfBallComplete() '"
End Sub


'*****************
' PinUP Support
'*****************

Class PinupNULL ' Dummy Pinup class so I dont have to keep adding if cases when people dont choose pinup
  Public Sub LabelShowPage(screen, pagenum, vis, Special)
  End Sub
  Public Sub LabelSet(screen, label, text, vis, Special)
  End Sub
  Public Sub playlistplayex(screen, dir, fname, volume, priority)
  End Sub
End Class




'******************************************************
'****  DROP TARGETS by Rothbauerw
'******************************************************
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.


' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.


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
Dim DT7, DT8, DT9, DT90

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

Set DT7 = (new DropTarget)(target7, target7a, T7_BM_Dark_Room, 1, 0, false)
Set DT8 = (new DropTarget)(target8, target8a, T8_BM_Dark_Room, 2, 0, false)
Set DT9 = (new DropTarget)(target9, target9a, T9_BM_Dark_Room, 3, 0, false)
Set DT90= (new DropTarget)(target90, target90a, T90_BM_Dark_Room,  4, 0, true)

Dim DTArray
DTArray = Array(DT7, DT8, DT9, DT90)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 60 'in milliseconds
Const DTDropUpSpeed = 30 'in milliseconds
Const DTDropUnits = 50 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 5 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 6 'max degrees primitive rotates when hit
Const DTDropDelay = 10 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 10 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
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

  If perpvel > 0 And  perpvelafter <= 0 Then
    If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
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
      If UsingROM Then
        controller.Switch(Switchid mod 100) = 1
      Else
        DTAction switchid
      End If
      primary.uservalue = 0
      DTAnimate = 0
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
      Dim b, gBOT
      gBOT = GetBalls

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
    If UsingROM Then controller.Switch(Switchid mod 100) = 0
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
'   STAND-UP TARGET INITIALIZATION
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
Dim ST04, ST05, ST06, ST07, ST08, ST09, ST10, ST11, ST12, ST13

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

Set ST04 = (new StandupTarget)(Target32, targets_004_BM_Dark_Room, 4,  0)
Set ST05 = (new StandupTarget)(Target33, targets_005_BM_Dark_Room, 5,  0)
Set ST06 = (new StandupTarget)(Target34, targets_006_BM_Dark_Room, 6,  0)
Set ST07 = (new StandupTarget)(Target35, targets_007_BM_Dark_Room, 7,  0)
Set ST08 = (new StandupTarget)(Target36, targets_008_BM_Dark_Room, 8,  0)
Set ST09 = (new StandupTarget)(Target43, targets_009_BM_Dark_Room, 9,  0)
Set ST10 = (new StandupTarget)(Target44, targets_010_BM_Dark_Room, 10, 0)
Set ST11 = (new StandupTarget)(Target82, targets_011_BM_Dark_Room, 11, 0)
Set ST12 = (new StandupTarget)(Target81, targets_012_BM_Dark_Room, 12, 0)
Set ST13 = (new StandupTarget)(Target80, targets_BM_Dark_Room    , 13, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST04, ST05, ST06, ST07, ST08, ST09, ST10, ST11, ST12, ST13)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5         'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.1        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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
    If UsingROM Then
      vpmTimer.PulseSw switch mod 100
    Else
      STAction switch
    End If
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

Sub STAction(Switch)
  Select Case Switch
    Case 4:  GoldHit 0
    Case 5:  GoldHit 1
    Case 6:  GoldHit 2
    Case 7:  GoldHit 3
    Case 8:  GoldHit 4
    Case 9:  doWFTargetHit 0
    Case 10: doWFTargetHit 1
    Case 11: House(CurrentPlayer).RegisterUPFHit 6
    Case 12: House(CurrentPlayer).RegisterUPFHit 4
    Case 13: House(CurrentPlayer).RegisterUPFHit 2
  End Select
End Sub




'******************************************************
'   END STAND-UP TARGETS
'******************************************************





'******************************************************
'****  LAMPZ by nFozzy
'******************************************************


Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
InitLampsNF              ' Setup lamp assignments
LampTimer.Interval = -1
LampTimer.Enabled = 1


Sub LampTimer_Timer()
    'apophis - Use the InPlayState of the 1st light in the Lampz.obj array to set the Lampz.state
    dim idx : for idx = 0 to 69 ' Table-specific - don't update the Flashers or GI here
        if Lampz.IsLight(idx) then
            if IsArray(Lampz.obj(idx)) then
                dim tmp : tmp = Lampz.obj(idx)
                Lampz.state(idx) = tmp(0).GetInPlayStateBool
                'debug.print tmp(0).name & " " &  tmp(0).GetInPlayStateBool & " " & tmp(0).IntensityScale  & vbnewline
                Lampz.SetColor(idx) = tmp(0).colorfull ' Also set color for RGB changeable lights
            Else
                Lampz.state(idx) = Lampz.obj(idx).GetInPlayStateBool
                Lampz.SetColor(idx) = Lampz.obj(idx).colorfull ' Also set color for RGB changeable lights
                'debug.print Lampz.obj(idx).name & " " &  Lampz.obj(idx).GetInPlayStateBool & " " & Lampz.obj(idx).IntensityScale  & vbnewline
            end if
        end if
    Next

  Lampz.Update1 'update (fading logic only)
End Sub

' Table-specific - this timer updates the Flashers in Lampz more frequently, to ensure quick flashes are reproduced
LampTimer3.Interval = 5
LampTimer.Enabled=1
Sub LampTimer3_Timer()
    ' Only do the Flashers
    dim idx : for idx = 70 to 79
        if lampz.IsLight(idx) then
            if IsArray(Lampz.obj(idx)) then
                dim tmp : tmp = Lampz.obj(idx)
                Lampz.state(idx) = tmp(0).GetInPlayStateBool
                if idx > 76 then Lampz.SetColor(idx) = tmp(0).colorfull
            Else
                Lampz.state(idx) = Lampz.obj(idx).GetInPlayStateBool
                if idx > 76 then Lampz.SetColor(idx) = Lampz.obj(idx).colorfull
            end if
        end if
    Next
    for idx = 80 to 82 ' GI
        if lampz.IsLight(idx) then
            if lampz.state(idx) = 0 Or lampz.state(idx) >= 1 Then ' GI isn't in dimmed state, ok to update from bulb
                if IsArray(Lampz.obj(idx)) then
                    tmp = Lampz.obj(idx)
                    Lampz.state(idx) = tmp(0).GetInPlayStateBool
                Else
                    Lampz.state(idx) = Lampz.obj(idx).GetInPlayStateBool
                end if
            End if
        end if
    Next
End Sub

LampTimer2.Interval = -1
LampTimer2.Enabled = True
Sub LampTimer2_Timer()

  ponLightShootAgain.color = LightShootAgain.colorfull
  p26on.color = li26.colorfull
  p32on.color = li32.colorfull
  p38on.color = li38.colorfull
  p41on.color = li41.colorfull
  p44on.color = li44.colorfull
  p47on.color = li47.colorfull
  p50on.color = li50.colorfull
  p53on.color = li53.colorfull
  p11on.color = li11.colorfull
  p14on.color = li14.colorfull
  p29on.color = li29.colorfull
  p35on.color = li35.colorfull
  p71on.color = li71.colorfull
  p74on.color = li74.colorfull

  p17on.color = li17.colorfull
  p20on.color = li20.colorfull
  p23on.color = li23.colorfull
  p80on.color = li80.colorfull
  p83on.color = li83.colorfull

  p86on.color = li86.colorfull
  p98on.color = li98.colorfull
  p114on.color = li114.colorfull
  p141on.color = li141.colorfull
  p156on.color = li156.colorfull
  p77on.color = li77.colorfull
  p95on.color = li95.colorfull
  p111on.color = li111.colorfull
  p108on.color = li108.colorfull
  p138on.color = li138.colorfull
  p153on.color = li153.colorfull
  p150on.color = li150.colorfull
  p92on.color = li92.colorfull
  p105on.color = li105.colorfull
  p120on.color = li120.colorfull
  p135on.color = li135.colorfull
  p147on.color = li147.colorfull

  p89on.color = li89.colorfull
  p101on.color = li101.colorfull
  p117on.color = li117.colorfull
  p132on.color = li132.colorfull
  p144on.color = li144.colorfull
  p159on.color = li159.colorfull
  p62on.color = li62.colorfull
  p65on.color = li65.colorfull
  p56on.color = li56.colorfull
  p59on.color = li59.colorfull

  p123on.color = li123.colorfull
  p126on.color = li126.colorfull
  p129on.color = li129.colorfull

    p242on.color = fl242.colorfull
  p162on.color = li162.colorfull
  p165on.color = li165.colorfull

    'UPF
    p180on.color = li180.colorfull
    p183on.color = li183.colorfull
    p186on.color = li186.colorfull
    p189on.color = li189.colorfull
    p192on.color = li192.colorfull
    p195on.color = li195.colorfull
    p198on.color = li198.colorfull
    p201on.color = li201.colorfull
    p204on.color = li204.colorfull
    p207on.color = li207.colorfull
    p210on.color = li210.colorfull
    p213on.color = li213.colorfull
    p216on.color = li216.colorfull

    'Bumpers
    li168a.colorfull=li168.colorfull : li168b.colorfull = li168.colorfull
    li171a.colorfull=li171.colorfull : li171b.colorfull = li171.colorfull
    li174a.colorfull=li174.colorfull : li174b.colorfull = li174.colorfull

  'Desktop action button
  liDTaction.colorfull = li230.colorfull


  Lampz.Update 'updates on frametime (Object updates only)
End Sub

Function FlashLevelToIndex(Input, MaxSize)
  FlashLevelToIndex = cInt(MaxSize * Input)
End Function

'Sub UpdateLightmap(lightmap, intensity, ByVal aLvl)  ' additiveblend prims
' if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)    'Callbacks don't get this filter automatically
'    lightmap.Opacity = aLvl * intensity
'End Sub

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
  UpdateMaterial "tracyskin",0.25,0.8,0,0,0,0,1,RGB(103*aLvl+121,108-88*aLvl,111*aLvl+57),RGB(103*aLvl+121,108-88*aLvl,111*aLvl+57),RGB(103*aLvl+121,108-88*aLvl,111*aLvl+57),True,False,0,0,0,0
end sub


Sub InitLampsNF()

  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

  'Adjust fading speeds (1 / full MS fading time)
  dim x

  for x = 0 to 140 : Lampz.FadeSpeedUp(x) = 1/2 : Lampz.FadeSpeedDown(x) = 1/4 : next
' for x = 111 to 140 : Lampz.FadeSpeedUp(x) = 1/5: Lampz.FadeSpeedDown(x) = 1/22 : next 'slow lamps for gi testing

  'Lampz.FadeSpeedUp(81) = 1/20 : Lampz.FadeSpeedDown(81) = 1/60

    ' Room lighting
  Lampz.FadeSpeedUp(150) = 100 : Lampz.FadeSpeedDown(150) = 100 : Lampz.Modulate(150) = 1/100

  'Lampz Assignments
  '  In a ROM based table, the lamp ID is used to set the state of the Lampz objects

  'MassAssign is the way to do assignments. It'll create arrays automatically / append objects to existing arrays
  'If not using a ROM, then the first light in the object array should be an invisible control light (in this example they
    'are named starting with "lc"). Then, lights should always be controlled via these control lights, including
    'when using Light Sequencers.

  Lampz.MassAssign(1)= LightShootAgain
' Lampz.MassAssign(1)= LightShootAgainbloom
  Lampz.Callback(1) = "DisableLighting ponLightShootAgain, 50,"

  Lampz.MassAssign(2)= li11
' Lampz.MassAssign(2)= li11bloom
  Lampz.Callback(2) = "DisableLighting  p11on, 50,"
  Lampz.MassAssign(3)= li14
' Lampz.MassAssign(3)= li14bloom
  Lampz.Callback(3) = "DisableLighting  p14on, 50,"

  Lampz.MassAssign(4)= li17
' Lampz.MassAssign(4)= li17bloom
  Lampz.Callback(4) = "DisableLighting  p17on, 50,"
  Lampz.MassAssign(5)= li20
' Lampz.MassAssign(5)= li20bloom
  Lampz.Callback(5) = "DisableLighting  p20on, 50,"
  Lampz.MassAssign(6)= li23
' Lampz.MassAssign(6)= li23bloom
  Lampz.Callback(6) = "DisableLighting  p23on, 50,"
  Lampz.MassAssign(7)= li26             ' control light -100h
' Lampz.MassAssign(7)= li26bloom            ' bloom light +1 height
  Lampz.Callback(7) = "DisableLighting  p26on, 200,"  ' on primitive

  Lampz.MassAssign(8)= li29
' Lampz.MassAssign(8)= li29bloom
  Lampz.Callback(8) = "DisableLighting  p29on, 50,"
  Lampz.MassAssign(9)= li32
' Lampz.MassAssign(9)= li32bloom
  Lampz.Callback(9) = "DisableLighting  p32on, 100,"
    Lampz.MassAssign(10)= li35
' Lampz.MassAssign(10)= li35bloom
  Lampz.Callback(10) = "DisableLighting  p35on, 50,"
  Lampz.MassAssign(11)= li38
' Lampz.MassAssign(11)= li38bloom
  Lampz.Callback(11) = "DisableLighting  p38on, 200,"
  Lampz.MassAssign(12)= li41
' Lampz.MassAssign(12)= li41bloom
  Lampz.Callback(12) = "DisableLighting  p41on, 200,"
  Lampz.MassAssign(13)= li44
' Lampz.MassAssign(13)= li44bloom
  Lampz.Callback(13) = "DisableLighting  p44on, 200,"
  Lampz.MassAssign(14)= li47
' Lampz.MassAssign(14)= li47bloom
  Lampz.Callback(14) = "DisableLighting  p47on, 200,"
  Lampz.MassAssign(15)= li50
' Lampz.MassAssign(15)= li50bloom
  Lampz.Callback(15) = "DisableLighting  p50on, 200,"
  Lampz.MassAssign(16)= li53
' Lampz.MassAssign(16)= li53bloom
  Lampz.Callback(16) = "DisableLighting  p53on, 200,"
  Lampz.MassAssign(17)= li56
' Lampz.MassAssign(17)= li56bloom
  Lampz.Callback(17) = "DisableLighting  p56on, 50,"
  Lampz.MassAssign(18)= li59
' Lampz.MassAssign(18)= li59bloom
  Lampz.Callback(18) = "DisableLighting  p59on, 50,"
  Lampz.MassAssign(19)= li62
' Lampz.MassAssign(19)= li62bloom
  Lampz.Callback(19) = "DisableLighting  p62on, 50,"
  Lampz.MassAssign(20)= li65
' Lampz.MassAssign(20)= li65bloom
  Lampz.Callback(20) = "DisableLighting  p65on, 50,"
  Lampz.MassAssign(21)= li71
' Lampz.MassAssign(21)= li71bloom
  Lampz.Callback(21) = "DisableLighting  p71on, 50,"
  Lampz.MassAssign(22)= li74
' Lampz.MassAssign(22)= li74bloom
  Lampz.Callback(22) = "DisableLighting  p74on, 50,"
  Lampz.MassAssign(23)= li77
' Lampz.MassAssign(23)= li77bloom
  Lampz.Callback(23) = "DisableLighting  p77on, 200,"
  Lampz.MassAssign(24)= li80
' Lampz.MassAssign(24)= li80bloom
  Lampz.Callback(24) = "DisableLighting  p80on, 50,"
  Lampz.MassAssign(25)= li83
' Lampz.MassAssign(25)= li83bloom
  Lampz.Callback(25) = "DisableLighting  p83on, 50,"
  Lampz.MassAssign(26)= li86
' Lampz.MassAssign(26)= li86bloom
  Lampz.Callback(26) = "DisableLighting  p86on, 200,"
  Lampz.MassAssign(27)= li89
' Lampz.MassAssign(27)= li89bloom
  Lampz.Callback(27) = "DisableLighting  p89on, 50,"

  Lampz.MassAssign(28)= li92
' Lampz.MassAssign(28)= li92bloom
  Lampz.Callback(28) = "DisableLighting  p92on, 50,"
  Lampz.MassAssign(29)= li95
' Lampz.MassAssign(29)= li95bloom
  Lampz.Callback(29) = "DisableLighting  p95on, 50,"
  Lampz.MassAssign(30)= li98
' Lampz.MassAssign(30)= li98bloom
  Lampz.Callback(30) = "DisableLighting  p98on, 200,"
  Lampz.MassAssign(31)= li101
' Lampz.MassAssign(31)= li101bloom
  Lampz.Callback(31) = "DisableLighting  p101on, 50,"

  Lampz.MassAssign(32)= li105
' Lampz.MassAssign(32)= li105bloom
  Lampz.Callback(32) = "DisableLighting  p105on, 50,"

  Lampz.MassAssign(33)= li108
' Lampz.MassAssign(33)= li108bloom
  Lampz.Callback(33) = "DisableLighting  p108on, 50,"

  Lampz.MassAssign(34)= li111
' Lampz.MassAssign(34)= li111bloom
  Lampz.Callback(34) = "DisableLighting  p111on, 50,"

  Lampz.MassAssign(35)= li114
' Lampz.MassAssign(35)= li114bloom
  Lampz.Callback(35) = "DisableLighting  p114on, 200,"
  Lampz.MassAssign(36)= li117
' Lampz.MassAssign(36)= li117bloom
  Lampz.Callback(36) = "DisableLighting  p117on, 50,"
  Lampz.MassAssign(37)= li120
' Lampz.MassAssign(37)= li120bloom
  Lampz.Callback(37) = "DisableLighting  p120on, 50,"
  Lampz.MassAssign(38)= li123
' Lampz.MassAssign(38)= li123bloom
  Lampz.Callback(38) = "DisableLighting  p123on, 50,"
  Lampz.MassAssign(39)= li126
' Lampz.MassAssign(39)= li126bloom
  Lampz.Callback(39) = "DisableLighting  p126on, 50,"
  Lampz.MassAssign(40)= li129
' Lampz.MassAssign(40)= li129bloom
  Lampz.Callback(40) = "DisableLighting  p129on, 50,"
  Lampz.MassAssign(41)= li132
' Lampz.MassAssign(41)= li132bloom
  Lampz.Callback(41) = "DisableLighting  p132on, 50,"
  Lampz.MassAssign(42)= li135
' Lampz.MassAssign(42)= li135bloom
  Lampz.Callback(42) = "DisableLighting  p135on, 50,"
    Lampz.MassAssign(43)= li138
' Lampz.MassAssign(43)= li138bloom
  Lampz.Callback(43) = "DisableLighting  p138on, 50,"
  Lampz.MassAssign(44)= li141
' Lampz.MassAssign(44)= li141bloom
  Lampz.Callback(44) = "DisableLighting  p141on, 200,"
    Lampz.MassAssign(45)= li144
' Lampz.MassAssign(45)= li144bloom
  Lampz.Callback(45) = "DisableLighting  p144on, 50,"
    Lampz.MassAssign(46)= li147
' Lampz.MassAssign(46)= li147bloom
  Lampz.Callback(46) = "DisableLighting  p147on, 50,"
  Lampz.MassAssign(47)= li150
' Lampz.MassAssign(47)= li150bloom
  Lampz.Callback(47) = "DisableLighting  p150on, 50,"
    Lampz.MassAssign(48)= li153
' Lampz.MassAssign(48)= li153bloom
  Lampz.Callback(48) = "DisableLighting  p153on, 50,"
  Lampz.MassAssign(49)= li156
' Lampz.MassAssign(49)= li156bloom
  Lampz.Callback(49) = "DisableLighting  p156on, 200,"
    Lampz.MassAssign(50)= li159
' Lampz.MassAssign(50)= li159bloom
  Lampz.Callback(50) = "DisableLighting  p159on, 50,"
  Lampz.MassAssign(51)= li162
' Lampz.MassAssign(51)= li162bloom
  Lampz.Callback(51) = "DisableLighting  p162on, 50,"
  Lampz.MassAssign(52)= li165
' Lampz.MassAssign(52)= li165bloom
  Lampz.Callback(52) = "DisableLighting  p165on, 50,"


    ' UPF
    Lampz.MassAssign(53)= li180
' Lampz.MassAssign(53)= li180bloom
  Lampz.Callback(53) = "DisableLighting  p180on, 200,"
    Lampz.MassAssign(54)= li183
' Lampz.MassAssign(54)= li183bloom
  Lampz.Callback(54) = "DisableLighting  p183on, 200,"
    Lampz.MassAssign(55)= li186
' Lampz.MassAssign(55)= li186bloom
  Lampz.Callback(55) = "DisableLighting  p186on, 200,"
    Lampz.MassAssign(56)= li189
' Lampz.MassAssign(56)= li189bloom
  Lampz.Callback(56) = "DisableLighting  p189on, 200,"
    Lampz.MassAssign(57)= li192
' Lampz.MassAssign(57)= li192bloom
  Lampz.Callback(57) = "DisableLighting  p192on, 200,"
    Lampz.MassAssign(58)= li195
' Lampz.MassAssign(58)= li195bloom
  Lampz.Callback(58) = "DisableLighting  p195on, 200,"
    Lampz.MassAssign(59)= li198
' Lampz.MassAssign(59)= li198bloom
  Lampz.Callback(59) = "DisableLighting  p198on, 200,"
    Lampz.MassAssign(60)= li201
' Lampz.MassAssign(60)= li201bloom
  Lampz.Callback(60) = "DisableLighting  p201on, 200,"
    Lampz.MassAssign(61)= li204
' Lampz.MassAssign(61)= li204bloom
  Lampz.Callback(61) = "DisableLighting  p204on, 200,"
    Lampz.MassAssign(62)= li207
' Lampz.MassAssign(62)= li207bloom
  Lampz.Callback(62) = "DisableLighting  p207on, 200,"
    Lampz.MassAssign(63)= li210
' Lampz.MassAssign(63)= li210bloom
  Lampz.Callback(63) = "DisableLighting  p210on, 200,"
    Lampz.MassAssign(64)= li213
' Lampz.MassAssign(64)= li213bloom
  Lampz.Callback(64) = "DisableLighting  p213on, 200,"
    Lampz.MassAssign(65)= li216
' Lampz.MassAssign(65)= li216bloom
  Lampz.Callback(65) = "DisableLighting  p216on, 200,"
    'Bumpers
    Lampz.MassAssign(66)= li168
  Lampz.MassAssign(66)= li168a
  Lampz.MassAssign(66)= li168b
 '   Lampz.MassAssign(66)= li168bloom

    Lampz.MassAssign(67)= li171
  Lampz.MassAssign(67)= li171a
  Lampz.MassAssign(67)= li171b
  '  Lampz.MassAssign(67)= li171bloom


    Lampz.MassAssign(68)= li174
  Lampz.MassAssign(68)= li174a
  Lampz.MassAssign(68)= li174b
   ' Lampz.MassAssign(68)= li174bloom

    ' Spinner Flasher
    For x = 70 to 76 : Lampz.FadeSpeedUp(x) = 1 : Lampz.FadeSpeedDown(x) = 1 : Next
    Lampz.MassAssign(73)= fl242
  'Lampz.MassAssign(73)= li242bloom
  Lampz.Callback(73) = "DisableLighting  p242on, 100,"

    ' Other Flashers - most not yet programmed
    ' Sword flasher
    Lampz.MassAssign(70)= fl235
    ' Right ramp flasher
    Lampz.MassAssign(71)= fl236
    ' Bumper flasher
    Lampz.MassAssign(72)= fl237
    ' Battering ram Flasher
    Lampz.MassAssign(74)= fl243
    ' Sling spotlight flashers
    Lampz.MassAssign(75)= fl244
    ' Backboard flasher
    Lampz.MassAssign(76)= fl245
    ' UPF Flashers
    Lampz.MassAssign(77)= li222

    ' Lockbar Button Light
    Lampz.MassAssign(78)= li230
  If DesktopMode Then Lampz.MassAssign(78)= liDTaction
    ' Apron button light
    Lampz.MassAssign(79)= li231

    ' GI lighting
    Lampz.MassAssign(80)= gi010
    Lampz.MassAssign(81)= gi026
    Lampz.MassAssign(82)= gi036

    ' Init Blender lightmaps
    LampzHelper

    ' Hide lamps
    HideLightHelper

  'Turn off all lamps on startup
  Lampz.Init  'This just turns state of any lamps to 1

  'Immediate update to turn on GI, turn off lamps
  Lampz.Update

End Sub

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
  Public IsLight(150)         'apophis
  Public FadeSpeedDown(150), FadeSpeedUp(150)
  Private Lock(150), Loaded(150), OnOff(150)
    Public ColorState(150)
  Public UseFunction
  Private cFilter
  Public UseCallback(150), cCallback(150)
  Public Lvl(150), Obj(150)
  Private Mult(150)
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
      OnOff(x) = 0
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
        if TypeName(input) <> "Double" and typename(input) <> "Integer"  and typename(input) <> "Long" then
      If input Then
        input = 1
      Else
        input = 0
      End If
    End If
    if Input <> OnOff(idx) then  'discard redundant updates
      OnOff(idx) = input
      Lock(idx) = False
      Loaded(idx) = False
    End If
  End Property

    ' Handle RGB light color changes without state change, on Original tables
    ' Note, all this does is detect color change and clear the lock/loaded flags to trigger an update.
    Public Property Let SetColor(ByVal idx, color)
        Dim c,c1 : c = CDbl(color) : c1 = CDbl(ColorState(idx)) ' VBScript can't compare unsigned ints, so convert to DBLs for comparison
        if c <> c1 then
            ColorState(idx) = color
            Lock(idx) = False
            Loaded(idx) = False
        End if
    End Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
        if typename(aInput) = "Light" then IsLight(aIdx) = True   'apophis - If first object in array is a light, this will be set true
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
        if OnOff(x) > Lvl(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          if Lvl(x) <= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        end if
      end if
    Next
  End Sub

  Public Sub Update2()   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) > Lvl(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= OnOff(x)  then Lvl(x) = OnOff(x) : Lock(x) = True
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
            obj(x).Intensityscale = Lvl(x)*Mult(x)
          End If
        end if
        if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,Lvl(x)*mult(x) 'Proc
        If Lock(x) Then
          if Lvl(x) = OnOff(x) or Lvl(x) = 0 then Loaded(x) = True  'finished fading
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


'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************

Dim Cabinetmode               '0 - Siderails On, 1 - Siderails Off
Dim OutPostMod
Dim LightLevel : LightLevel = 65
VolumeDial = 0.8              'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim RampRollVolume : RampRollVolume = 0.5   'Level of ramp rolling volume. Value between 0 and 1
DynamicBallShadowsOn = 1          '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Dim BallImageId : BallImageId = 0       'Default ball image / classic

' Base options
Const Opt_Light = 0
'Const Opt_Outpost = 1
Const Opt_Flames = 1
Const Opt_Volume = 2
Const Opt_Volume_Ramp = 3
Const Opt_Volume_Ball = 4
Const Opt_Ball_Image = 5
' Table mods & toys
Const Opt_Cabinet = 6
Const Opt_LockBar = 7
' Advanced options
Const Opt_DynBallShadow = 8
'VR
Const Opt_VR_RoomChoice = 9
Const Opt_VR_Scratches = 10
Const Opt_VR_Topper = 11
' Informations
Const Opt_LUT = 12
Const Opt_Info_1 = 13 ' note: selectable options value must be < Opt_Info_1
Const Opt_Info_2 = 14

Const NOptions = 15

Dim OptionDMD: Set OptionDMD = Nothing
Dim bOptionsMagna, bInOptions, optionsSaveAttract : bOptionsMagna = False
Dim OptPos, OptSelected, OptN, OptTop, OptBot, OptSel
Dim OptFontHi, OptFontLo
Dim Options_Flames_Enabled

Sub Options_Open
    optionsSaveAttract=bAttractMode
    bAttractMode=False
    tmrDMDUpdate.Enabled = False
    tmrAttractModeScene.Enabled = False

    bOptionsMagna = False

  Debug.Print "Option UI opened"
  bInOptions = True
  OptPos = 0
  OptSelected = False

  Dim a, scene, font
  Set scene = FlexDMD.NewGroup("Options")
  Set OptFontHi = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
  Set OptFontLo = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(100, 100, 100), RGB(100, 100, 100), 0)
  Set OptSel = FlexDMD.NewGroup("Sel")
  Set a = FlexDMD.NewLabel(">", OptFontLo, ">>>")
  a.SetAlignedPosition 1, 16, FlexDMD_Align_Left
  OptSel.AddActor a
  Set a = FlexDMD.NewLabel(">", OptFontLo, "<<<")
  a.SetAlignedPosition 127, 16, FlexDMD_Align_Right
  OptSel.AddActor a
  scene.AddActor OptSel
  OptSel.SetBounds 0, 0, 128, 32
  OptSel.Visible = False

  Set a = FlexDMD.NewLabel("Info1", OptFontLo, "MAGNA EXIT/ENTER")
  a.SetAlignedPosition 1, 32, FlexDMD_Align_BottomLeft
  scene.AddActor a
  Set a = FlexDMD.NewLabel("Info2", OptFontLo, "FLIPPER SELECT")
  a.SetAlignedPosition 127, 32, FlexDMD_Align_BottomRight
  scene.AddActor a
  Set OptN = FlexDMD.NewLabel("Pos", OptFontLo, "LINE 1")
  Set OptTop = FlexDMD.NewLabel("Top", OptFontLo, "LINE 1")
  Set OptBot = FlexDMD.NewLabel("Bottom", OptFontLo, "LINE 2")
  scene.AddActor OptN
  scene.AddActor OptTop
  scene.AddActor OptBot
  Options_OnOptChg
  DMDDisplayScene scene
End Sub

Sub Options_Close
  bInOptions = False
  bAttractMode=optionsSaveAttract
    if bAttractMode then
        StartAttractMode
    Else
        bDefaultScene = False
        tmrDMDUpdate.Enabled = True
    End if

  'write settings to keep f12/tweak in sync
  WriteSettings
End Sub

Function Options_OnOffText(opt)
  If opt Then
    Options_OnOffText = "ON"
  Else
    Options_OnOffText = "OFF"
  End If
End Function

Sub Options_OnOptChg
  FlexDMD.LockRenderThread
  OptN.Text = (OptPos+1) & "/" & NOptions
  If OptSelected Then
    OptTop.Font = OptFontLo
    OptBot.Font = OptFontHi
    OptSel.Visible = True
  Else
    OptTop.Font = OptFontHi
    OptBot.Font = OptFontLo
    OptSel.Visible = False
  End If
  If OptPos = Opt_Light Then
    OptTop.Text = "LIGHT LEVEL"
    OptBot.Text = "LEVEL " & LightLevel
    SaveValue cGameName, "LIGHT", LightLevel
  ' ElseIf OptPos = Opt_Outpost Then
  '   OptTop.Text = "OUT POST DIFFICULTY"
  '   If OutPostMod = 0 Then
  '     OptBot.Text = "EASY"
  '   ElseIf OutPostMod = 1 Then
  '     OptBot.Text = "MEDIUM"
  '   ElseIf OutPostMod = 2 Then
  '     OptBot.Text = "HARD"
  '   ElseIf OutPostMod = 3 Then
  '     OptBot.Text = "HARDEST"
  '   End If
  '   SaveValue cGameName, "OUTPOST", OutPostMod
    ElseIf OptPos = Opt_Flames Then
        OptTop.Text = "DRAGON FLAMES"
        OptBot.Text = Options_OnOffText(Options_Flames_Enabled)
        SaveValue cGameName, "DRAGONFLAMES", Options_Flames_Enabled
  ElseIf OptPos = Opt_Volume Then
    OptTop.Text = "MECH VOLUME"
    OptBot.Text = "LEVEL " & CInt(VolumeDial * 100)
    SaveValue cGameName, "VOLUME", VolumeDial
  ElseIf OptPos = Opt_Volume_Ramp Then
    OptTop.Text = "RAMP VOLUME"
    OptBot.Text = "LEVEL " & CInt(RampRollVolume * 100)
    SaveValue cGameName, "RAMPVOLUME", RampRollVolume
  ElseIf OptPos = Opt_Volume_Ball Then
    OptTop.Text = "BALL VOLUME"
    OptBot.Text = "LEVEL " & CInt(BallRollVolume * 100)
    SaveValue cGameName, "BALLVOLUME", BallRollVolume
  ElseIf OptPos = Opt_Ball_Image Then
    OptTop.Text = "BALL IMAGE"
    if BallImageId = 0 then
      OptBot.Text = "CLASSIC"
    elseif BallImageId = 1 then
      OptBot.Text = "OLD ARSE"
    else
      OptBot.Text = "OLD ARSE WARM"
    end if
    SaveValue cGameName, "BALLIMAGE", BallImageId
  ElseIf OptPos = Opt_Cabinet Then
    OptTop.Text = "CABINET MODE"
    OptBot.Text = Options_OnOffText(CabinetMode)
    SaveValue cGameName, "CABINET", CabinetMode
    ElseIf OptPos = Opt_LockBar Then
        OptTop.text = "PHYSICAL LOCKBAR BUTTON"
        If bHaveLockbarButton Then OptBot.Text = "YES" Else OptBot.Text = "NO"
        SaveValue cGameName, "LBBUT", bHaveLockbarButton
  ElseIf OptPos = Opt_DynBallShadow Then
    OptTop.Text = "DYN. BALL SHADOWS"
    OptBot.Text = Options_OnOffText(DynamicBallShadowsOn)
    SaveValue cGameName, "DYNBALLSH", DynamicBallShadowsOn
  'VR
  ElseIf OptPos = Opt_VR_RoomChoice Then
    OptTop.Text = "VR ROOM"
    'if VRRoomChoice = 0 then
    ' OptBot.Text = "DESKTOP"
    if VRRoomChoice = 1 then
      OptBot.Text = "MINIMAL"
    elseif VRRoomChoice = 2 then
      OptBot.Text = "ULTRA MINIMAL"
    elseif VRRoomChoice = 3 then
      OptBot.Text = "IRON THRONE MEGA"
    elseif VRRoomChoice = 4 then
      OptBot.Text = "DRAGON ICE MEGA"
    else
      OptBot.Text = "TOWER THRONE MEGA"
    end if
    SaveValue cGameName, "VRROOM", VRRoomChoice
    VRRoomInit
    ElseIf OptPos = Opt_VR_Scratches Then
        OptTop.Text = "VR GLASS SCRATCHES"
        OptBot.Text = Options_OnOffText(Scratches)
    SaveValue cGameName, "VRSCRATCHES", Scratches
    VRRoomInit
  ElseIf OptPos = Opt_VR_Topper Then
        OptTop.Text = "VR TOPPER"
        OptBot.Text = Options_OnOffText(Topper)
    SaveValue cGameName, "VRTOPPER", Topper
    VRRoomInit
  ElseIf OptPos = Opt_Info_1 Then
    OptTop.Text = "VPX " & VersionMajor & "." & VersionMinor & "." & VersionRevision
    OptBot.Text = "GAME OF THRONES " & TableVersion
  ElseIf OptPos = Opt_Info_2 Then
    OptTop.Text = "RENDER MODE"
    If RenderingMode = 0 Then OptBot.Text = "DEFAULT"
    If RenderingMode = 1 Then OptBot.Text = "STEREO 3D"
    If RenderingMode = 2 Then OptBot.Text = "VR"
    ElseIf OptPos = Opt_LUT Then
    debug.print "LUT Displayed"
        OptTop.Text = "LUT"
        OptBot.Text = LUTImage
  End If
  OptTop.SetAlignedPosition 127, 1, FlexDMD_Align_TopRight
  OptBot.SetAlignedPosition 64, 16, FlexDMD_Align_Center
  FlexDMD.UnlockRenderThread
  UpdateMods
End Sub

Sub Options_Toggle(amount)
  debug.print "Toggle:" & amount
  If OptPos = Opt_Light Then
    LightLevel = LightLevel + amount * 5
    If LightLevel < 0 Then LightLevel = 100
    If LightLevel > 100 Then LightLevel = 0
  ' ElseIf OptPos = Opt_Outpost Then
  '   OutPostMod = OutPostMod + amount
  '   If OutPostMod < 0 Then OutPostMod = 3
  '   If OutPostMod > 3 Then OutPostMod = 0
    ElseIf OptPos = Opt_Flames Then
        Options_Flames_Enabled = 1 - Options_Flames_Enabled
  ElseIf OptPos = Opt_Volume Then
    VolumeDial = VolumeDial + amount * 0.1
    If VolumeDial < 0 Then VolumeDial = 1
    If VolumeDial > 1 Then VolumeDial = 0
  ElseIf OptPos = Opt_Volume_Ramp Then
    RampRollVolume = RampRollVolume + amount * 0.1
    If RampRollVolume < 0 Then RampRollVolume = 1
    If RampRollVolume > 1 Then RampRollVolume = 0
  ElseIf OptPos = Opt_Volume_Ball Then
    BallRollVolume = BallRollVolume + amount * 0.1
    If BallRollVolume < 0 Then BallRollVolume = 1
    If BallRollVolume > 1 Then BallRollVolume = 0
  ElseIf OptPos = Opt_Ball_Image Then
    BallImageId = BallImageId + amount
    If BallImageId < 0 Then BallImageId = 2
    If BallImageId > 2 Then BallImageId = 0
  ElseIf OptPos = Opt_Cabinet Then
    CabinetMode = 1 - CabinetMode
    ElseIf OptPos = Opt_LockBar Then
        bHaveLockbarButton = 1 - bHaveLockbarButton
  ElseIf OptPos = Opt_DynBallShadow Then
    DynamicBallShadowsOn = 1 - DynamicBallShadowsOn
    Elseif OptPos = Opt_LUT Then
    debug.print "LUT MODIFIED"
        LUTImage = LUTImage + amount
        If LUTImage < 0 Then LUTImage = 15
        LUTImage = ABS(LUTImage) MOD 16
        table1.ColorGradeImage = "LUT"&LUTImage
        SaveLUT
  ElseIf OptPos = Opt_VR_RoomChoice Then
    VRRoomChoice = VRRoomChoice + amount
    If VRRoomChoice < 1 Then VRRoomChoice = 5
    If VRRoomChoice > 5 Then VRRoomChoice = 1
    ElseIf OptPos = Opt_VR_Scratches Then
    Scratches = 1 - Scratches
  ElseIf OptPos = Opt_VR_Topper Then
        Topper = 1 - Topper
  End If
End Sub

Sub Options_KeyDown(ByVal keycode)
  If OptSelected Then
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      OptSelected = False
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      OptSelected = False
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      Options_Toggle  -1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      Options_Toggle  1
    End If
  Else
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      Options_Close
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      If OptPos < Opt_Info_1 Then OptSelected = True
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      OptPos = OptPos - 1
      If OptPos < 0 Then OptPos = NOptions - 1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      OptPos = OptPos + 1
      If OptPos >= NOPtions Then OptPos = 0
    End If
  End If
  Options_OnOptChg
End Sub

Sub Options_Load
  Dim x
    x = LoadValue(cGameName, "LIGHT")
    If x <> "" Then LightLevel = CInt(x) Else LightLevel = 65
    'x = LoadValue(cGameName, "OUTPOST")
    'If x <> "" Then OutPostMod = CInt(x) Else OutPostMod = 1
    x = LoadValue(cGameName,"DRAGONFLAMES")
    If x <> "" Then Options_Flames_Enabled = CInt(x) Else Options_Flames_Enabled = 1
    x = LoadValue(cGameName, "VOLUME")
    If x <> "" Then VolumeDial = CNCDbl(x) Else VolumeDial = 0.8
    x = LoadValue(cGameName, "RAMPVOLUME")
    If x <> "" Then RampRollVolume = CNCDbl(x) Else RampRollVolume = 0.5
    x = LoadValue(cGameName, "BALLVOLUME")
    If x <> "" Then BallRollVolume = CNCDbl(x) Else BallRollVolume = 0.5
    x = LoadValue(cGameName, "CABINET")
    If x <> "" Then CabinetMode = CInt(x) Else CabinetMode = 0
    x = LoadValue(cGameName,"LBBUT")
    If x <> "" Then bHaveLockbarButton = CInt(x) Else bHaveLockbarButton = False
    x = LoadValue(cGameName, "DYNBALLSH")
    If x <> "" Then DynamicBallShadowsOn = CInt(x) Else DynamicBallShadowsOn = 1
    x = LoadValue(cGameName, "BALLIMAGE")
    If x <> "" Then BallImageId = CInt(x) Else BallImageId = 0
  'VR
  x = LoadValue(cGameName, "VRROOM")
    If x <> "" Then VRRoomChoice = CInt(x) Else VRRoomChoice = 5
  x = LoadValue(cGameName, "VRSCRATCHES")
    If x <> "" and (x = "1" or x = "0") Then Scratches = CInt(x) Else Scratches = 1
  x = LoadValue(cGameName, "VRTOPPER")
    If x <> "" and (x = "1" or x = "0") Then Topper = CInt(x) Else Topper = 1

  UpdateMods
End Sub

Sub UpdateMods
  Dim y, enabled,nenabled

  'if (CabinetMode <> 0  And bHaveLockbarButton = 0) then enabled=1 else enabled=0

  if RenderingMode = 2 Then
    enabled=0
  else
    if (CabinetMode <> 0  And bHaveLockbarButton = 0) then enabled=1 else enabled=0
    'if (CabinetMode <> 1 And RenderingMode <> 2 And bHaveLockbarButton = 0) then enabled=1 else enabled=0
  end if

    nenabled = 1-enabled
    For each y in ApronFButton_BL : y.visible=abs(enabled) : Next
    For each y in fb_base_BL : y.visible=abs(enabled) : Next
    For each y in LockdownButton_BL : y.visible = nenabled : Next
    For each y in fl_base_BL : y.visible = nenabled : Next

    'if CabinetMode=0 then enabled=1 else enabled=0
  if CabinetMode=0 or RenderingMode=2 then enabled=1 else enabled=0
    SideRailMapLeft.visible = enabled : SideRailMapRight.visible = enabled

    Select Case LUTImage
        Case 14: FlexDMD.Color = RGB(16,255,16)
        Case 15: FlexDMD.Color = RGB(255,255,255)
        Case Else: FlexDMD.Color = RGB(255,44,16)
    End Select

  SetBall(BallImageId)

    ' Not Implemented
  ' If OutPostMod = 0 Then ' Easy
  '   routpost_BM_Dark_Room.x = 817.4415
  '   routpost_BM_Dark_Room.y = 1469.287
  '   loutpost_BM_Dark_Room.x = 51.87117
  '   loutpost_BM_Dark_Room.y = 1469.451
  '   rout_easy.collidable = true
  '   rout_medium.collidable = false
  '   rout_hard.collidable = false
  '   rout_hardest.collidable = false
  '   lout_easy.collidable = true
  '   lout_medium.collidable = false
  '   lout_hard.collidable = false
  '   lout_hardest.collidable = false
  ' ElseIf OutPostMod = 1 Then ' Medium
  '   routpost_BM_Dark_Room.x = 823.0609
  '   routpost_BM_Dark_Room.y = 1459.596
  '   loutpost_BM_Dark_Room.x = 48.93166
  '   loutpost_BM_Dark_Room.y = 1457.558
  '   rout_easy.collidable = false
  '   rout_medium.collidable = true
  '   rout_hard.collidable = false
  '   rout_hardest.collidable = false
  '   lout_easy.collidable = false
  '   lout_medium.collidable = true
  '   lout_hard.collidable = false
  '   lout_hardest.collidable = false
  ' ElseIf OutPostMod = 2 Then ' Hard
  '   routpost_BM_Dark_Room.x = 829.0158
  '   routpost_BM_Dark_Room.y = 1449.152
  '   loutpost_BM_Dark_Room.x = 45.53271
  '   loutpost_BM_Dark_Room.y = 1446.529
  '   rout_easy.collidable = false
  '   rout_medium.collidable = false
  '   rout_hard.collidable = true
  '   rout_hardest.collidable = false
  '   lout_easy.collidable = false
  '   lout_medium.collidable = false
  '   lout_hard.collidable = true
  '   lout_hardest.collidable = false
  ' ElseIf OutPostMod = 3 Then ' Hardest
  '   routpost_BM_Dark_Room.x = 834.4323
  '   routpost_BM_Dark_Room.y = 1439.318
  '   loutpost_BM_Dark_Room.x = 42.86211
  '   loutpost_BM_Dark_Room.y = 1435.624
  '   rout_easy.collidable = false
  '   rout_medium.collidable = false
  '   rout_hard.collidable = false
  '   rout_hardest.collidable = true
  '   lout_easy.collidable = false
  '   lout_medium.collidable = false
  '   lout_hard.collidable = false
  '   lout_hardest.collidable = true
  ' End If

  ' ' FIXME this should be setup in Blender
  ' routpost_LM_GI.x = routpost_BM_Dark_Room.x
  ' routpost_LM_GI.y = routpost_BM_Dark_Room.y
  ' routpost_LM_Lit_Room.x = routpost_BM_Dark_Room.x
  ' routpost_LM_Lit_Room.y = routpost_BM_Dark_Room.y
  ' routpost_LM_flashers_l130.x = routpost_BM_Dark_Room.x
  ' routpost_LM_flashers_l130.y = routpost_BM_Dark_Room.y
  ' routpost_LM_flashers_l131.x = routpost_BM_Dark_Room.x
  ' routpost_LM_flashers_l131.y = routpost_BM_Dark_Room.y
  ' routpost_LM_flashers_l132.x = routpost_BM_Dark_Room.x
  ' routpost_LM_flashers_l132.y = routpost_BM_Dark_Room.y

  ' ' FIXME this should be setup in Blender
  ' loutpost_LM_GI.x = loutpost_BM_Dark_Room.x
  ' loutpost_LM_GI.y = loutpost_BM_Dark_Room.y
  ' loutpost_LM_Lit_Room.x = loutpost_BM_Dark_Room.x
  ' loutpost_LM_Lit_Room.y = loutpost_BM_Dark_Room.y
  ' loutpost_LM_flashers_l130.x = loutpost_BM_Dark_Room.x
  ' loutpost_LM_flashers_l130.y = loutpost_BM_Dark_Room.y
  ' loutpost_LM_flashers_l131.x = loutpost_BM_Dark_Room.x
  ' loutpost_LM_flashers_l131.y = loutpost_BM_Dark_Room.y

  Lampz.state(150) = LightLevel
End Sub

' Culture neutral string to double conversion (handles situation where you don't know how the string was written)
Function CNCDbl(str)
    Dim strt, Sep, i
    If IsNumeric(str) Then
        CNCDbl = CDbl(str)
    Else
        Sep = Mid(CStr(0.5), 2, 1)
        Select Case Sep
        Case "."
            i = InStr(1, str, ",")
        Case ","
            i = InStr(1, str, ".")
        End Select
        If i = 0 Then
            CNCDbl = Empty
        Else
            strt = Mid(str, 1, i - 1) & Sep & Mid(str, i + 1)
            If IsNumeric(strt) Then
                CNCDbl = CDbl(strt)
            Else
                CNCDbl = Empty
            End If
        End If
    End If
End Function

Sub SetBall(ballid)
  dim BOT
  dim b
  dim ballImageName

  Select Case ballid
    Case 0: ballImageName = "ball_HDR_brightest"
    Case 1: ballImageName = "old_ass_eyes_ball"
    Case 2: ballImageName = "old_ass_eyes_ball_warm"
    Case Else ballImageName = "ball_HDR_brightest"
  End Select

  Table1.BallImage = ballImageName

  BOT = GetBalls
  For b = 0 to UBound(BOT)
    BOT(b).Image = ballImageName
    Next
End Sub

' ===============================================================
' ZVLM       Virtual Pinball X Light Mapper generated code
'
' This file provide default implementation and template to add bakemap
' & lightmap synchronization for position and lighting.
'
' Lines ending with a comment starting with "' VLM." are meant to
' be copy/pasted ONLY ONCE, since the toolkit will take care of
' updating them directly in your table script, each time an
' export is made.

' ===============================================================
' The following code NEEDS to be copy/pasted to hide the elements
' that are placed to avoid stutters when VPX loads all the nestmaps
' to the GPU. It also NEEDS the following line to be added to the.
' table init function:

' vpmTimer.AddTimer 1000, "WarmUpDone '"

Sub WarmUpDone
  VLM_Warmup_Nestmap_0.Visible = False
  VLM_Warmup_Nestmap_1.Visible = False
  VLM_Warmup_Nestmap_2.Visible = False
  VLM_Warmup_Nestmap_3.Visible = False
  VLM_Warmup_Nestmap_4.Visible = False
  VLM_Warmup_Nestmap_5.Visible = False
  VLM_Warmup_Nestmap_6.Visible = False
  VLM_Warmup_Nestmap_7.Visible = False
  VLM_Warmup_Nestmap_8.Visible = False
  VLM_Warmup_Nestmap_9.Visible = False
  VLM_Warmup_Nestmap_10.Visible = False
  VLM_Warmup_Nestmap_11.Visible = False
  VLM_Warmup_Nestmap_12.Visible = False
  VLM_Warmup_Nestmap_13.Visible = False
  VLM_Warmup_Nestmap_14.Visible = False
  VLM_Warmup_Nestmap_15.Visible = False
End Sub


' ===============================================================
' The following code can be copy/pasted to have premade array for
' movable objects:
' - _LM suffixed arrays contains the lightmaps
' - _BM suffixed arrays contains the bakemap
' - _BL suffixed arrays contains both the bakemap & the lightmaps
Dim ApronFButton_LM: ApronFButton_LM=Array(FireButton_LM_Flashers_f242, FireButton_LM_GI1, FireButton_LM_Inserts_LightApro, FireButton_LM_Inserts_li59, FireButton_LM_Inserts_li65, FireButton_LM_Lit_Room) ' VLM.Array;LM;ApronFButton
Dim ApronFButton_BM: ApronFButton_BM=Array(FireButton_BM_Dark_Room) ' VLM.Array;BM;ApronFButton
Dim ApronFButton_BL: ApronFButton_BL=Array(FireButton_BM_Dark_Room, FireButton_LM_Flashers_f242, FireButton_LM_GI1, FireButton_LM_Inserts_LightApro, FireButton_LM_Inserts_li59, FireButton_LM_Inserts_li65, FireButton_LM_Lit_Room) ' VLM.Array;BL;ApronFButton
Dim Bumper_Ring_LM: Bumper_Ring_LM=Array(Bumper_Ring_LM_Flashers_f236, Bumper_Ring_LM_Flashers_f237, Bumper_Ring_LM_Flashers_f243, Bumper_Ring_LM_Flashers_f245, Bumper_Ring_LM_GI2, Bumper_Ring_LM_Inserts_li168, Bumper_Ring_LM_Inserts_li171, Bumper_Ring_LM_Inserts_li174, Bumper_Ring_LM_Lit_Room) ' VLM.Array;LM;Bumper_Ring
Dim Bumper_Ring_BM: Bumper_Ring_BM=Array(Bumper_Ring_BM_Dark_Room) ' VLM.Array;BM;Bumper_Ring
Dim Bumper_Ring_BL: Bumper_Ring_BL=Array(Bumper_Ring_BM_Dark_Room, Bumper_Ring_LM_Flashers_f236, Bumper_Ring_LM_Flashers_f237, Bumper_Ring_LM_Flashers_f243, Bumper_Ring_LM_Flashers_f245, Bumper_Ring_LM_GI2, Bumper_Ring_LM_Inserts_li168, Bumper_Ring_LM_Inserts_li171, Bumper_Ring_LM_Inserts_li174, Bumper_Ring_LM_Lit_Room) ' VLM.Array;BL;Bumper_Ring
Dim Bumper_Ring_001_LM: Bumper_Ring_001_LM=Array(Bumper_Ring_001_LM_Flashers_f23, Bumper_Ring_001_LM_Flashers_f23, Bumper_Ring_001_LM_Flashers_f23, Bumper_Ring_001_LM_Flashers_f24, Bumper_Ring_001_LM_Flashers_f24, Bumper_Ring_001_LM_GI2, Bumper_Ring_001_LM_Inserts_li16, Bumper_Ring_001_LM_Inserts_li17, Bumper_Ring_001_LM_Inserts_li17, Bumper_Ring_001_LM_Lit_Room) ' VLM.Array;LM;Bumper_Ring_001
Dim Bumper_Ring_001_BM: Bumper_Ring_001_BM=Array(Bumper_Ring_001_BM_Dark_Room) ' VLM.Array;BM;Bumper_Ring_001
Dim Bumper_Ring_001_BL: Bumper_Ring_001_BL=Array(Bumper_Ring_001_BM_Dark_Room, Bumper_Ring_001_LM_Flashers_f23, Bumper_Ring_001_LM_Flashers_f23, Bumper_Ring_001_LM_Flashers_f23, Bumper_Ring_001_LM_Flashers_f24, Bumper_Ring_001_LM_Flashers_f24, Bumper_Ring_001_LM_GI2, Bumper_Ring_001_LM_Inserts_li16, Bumper_Ring_001_LM_Inserts_li17, Bumper_Ring_001_LM_Inserts_li17, Bumper_Ring_001_LM_Lit_Room) ' VLM.Array;BL;Bumper_Ring_001
Dim Bumper_Ring_002_LM: Bumper_Ring_002_LM=Array(Bumper_Ring_002_LM_Flashers_f23, Bumper_Ring_002_LM_Flashers_f23, Bumper_Ring_002_LM_Flashers_f24, Bumper_Ring_002_LM_GI2, Bumper_Ring_002_LM_Inserts_li16, Bumper_Ring_002_LM_Inserts_li17, Bumper_Ring_002_LM_Inserts_li17, Bumper_Ring_002_LM_Lit_Room) ' VLM.Array;LM;Bumper_Ring_002
Dim Bumper_Ring_002_BM: Bumper_Ring_002_BM=Array(Bumper_Ring_002_BM_Dark_Room) ' VLM.Array;BM;Bumper_Ring_002
Dim Bumper_Ring_002_BL: Bumper_Ring_002_BL=Array(Bumper_Ring_002_BM_Dark_Room, Bumper_Ring_002_LM_Flashers_f23, Bumper_Ring_002_LM_Flashers_f23, Bumper_Ring_002_LM_Flashers_f24, Bumper_Ring_002_LM_GI2, Bumper_Ring_002_LM_Inserts_li16, Bumper_Ring_002_LM_Inserts_li17, Bumper_Ring_002_LM_Inserts_li17, Bumper_Ring_002_LM_Lit_Room) ' VLM.Array;BL;Bumper_Ring_002
Dim Gates_001_LM: Gates_001_LM=Array(Gates_001_LM_Flashers_f237, Gates_001_LM_Flashers_f243, Gates_001_LM_Flashers_f245, Gates_001_LM_GI2, Gates_001_LM_Lit_Room) ' VLM.Array;LM;Gates_001
Dim Gates_001_BM: Gates_001_BM=Array(Gates_001_BM_Dark_Room) ' VLM.Array;BM;Gates_001
Dim Gates_001_BL: Gates_001_BL=Array(Gates_001_BM_Dark_Room, Gates_001_LM_Flashers_f237, Gates_001_LM_Flashers_f243, Gates_001_LM_Flashers_f245, Gates_001_LM_GI2, Gates_001_LM_Lit_Room) ' VLM.Array;BL;Gates_001
Dim LEMK_LM: LEMK_LM=Array(LEMK_LM_GI1, LEMK_LM_Lit_Room) ' VLM.Array;LM;LEMK
Dim LEMK_BM: LEMK_BM=Array(LEMK_BM_Dark_Room) ' VLM.Array;BM;LEMK
Dim LEMK_BL: LEMK_BL=Array(LEMK_BM_Dark_Room, LEMK_LM_GI1, LEMK_LM_Lit_Room) ' VLM.Array;BL;LEMK
Dim LFlogo_001_LM: LFlogo_001_LM=Array(LFlogo_001_LM_Flashers_f242, LFlogo_001_LM_Inserts_li198, LFlogo_001_LM_Inserts_li201, LFlogo_001_LM_Inserts_li204, LFlogo_001_LM_Inserts_li207, LFlogo_001_LM_Inserts_li210, LFlogo_001_LM_Inserts_li213, LFlogo_001_LM_Inserts_li216, LFlogo_001_LM_Lit_Room, LFlogo_001_LM_Upper_GI) ' VLM.Array;LM;LFlogo_001
Dim LFlogo_001_BM: LFlogo_001_BM=Array(LFlogo_001_BM_Dark_Room) ' VLM.Array;BM;LFlogo_001
Dim LFlogo_001_BL: LFlogo_001_BL=Array(LFlogo_001_BM_Dark_Room, LFlogo_001_LM_Flashers_f242, LFlogo_001_LM_Inserts_li198, LFlogo_001_LM_Inserts_li201, LFlogo_001_LM_Inserts_li204, LFlogo_001_LM_Inserts_li207, LFlogo_001_LM_Inserts_li210, LFlogo_001_LM_Inserts_li213, LFlogo_001_LM_Inserts_li216, LFlogo_001_LM_Lit_Room, LFlogo_001_LM_Upper_GI) ' VLM.Array;BL;LFlogo_001
Dim LFlogo_003_LM: LFlogo_003_LM=Array(LFlogo_003_LM_GI1, LFlogo_003_LM_Inserts_LightApro, LFlogo_003_LM_Inserts_LightShoo, LFlogo_003_LM_Inserts_li38, LFlogo_003_LM_Inserts_li41, LFlogo_003_LM_Inserts_li44, LFlogo_003_LM_Inserts_li47, LFlogo_003_LM_Inserts_li50, LFlogo_003_LM_Inserts_li53, LFlogo_003_LM_Inserts_li56, LFlogo_003_LM_Inserts_li59, LFlogo_003_LM_Inserts_li62, LFlogo_003_LM_Inserts_li65, LFlogo_003_LM_Lit_Room) ' VLM.Array;LM;LFlogo_003
Dim LFlogo_003_BM: LFlogo_003_BM=Array(LFlogo_003_BM_Dark_Room) ' VLM.Array;BM;LFlogo_003
Dim LFlogo_003_BL: LFlogo_003_BL=Array(LFlogo_003_BM_Dark_Room, LFlogo_003_LM_GI1, LFlogo_003_LM_Inserts_LightApro, LFlogo_003_LM_Inserts_LightShoo, LFlogo_003_LM_Inserts_li38, LFlogo_003_LM_Inserts_li41, LFlogo_003_LM_Inserts_li44, LFlogo_003_LM_Inserts_li47, LFlogo_003_LM_Inserts_li50, LFlogo_003_LM_Inserts_li53, LFlogo_003_LM_Inserts_li56, LFlogo_003_LM_Inserts_li59, LFlogo_003_LM_Inserts_li62, LFlogo_003_LM_Inserts_li65, LFlogo_003_LM_Lit_Room) ' VLM.Array;BL;LFlogo_003
Dim LSling1_LM: LSling1_LM=Array(LSling1_LM_GI1, LSling1_LM_Lit_Room) ' VLM.Array;LM;LSling1
Dim LSling1_BM: LSling1_BM=Array(LSling1_BM_Dark_Room) ' VLM.Array;BM;LSling1
Dim LSling1_BL: LSling1_BL=Array(LSling1_BM_Dark_Room, LSling1_LM_GI1, LSling1_LM_Lit_Room) ' VLM.Array;BL;LSling1
Dim LSling2_LM: LSling2_LM=Array(LSling2_LM_GI1, LSling2_LM_Lit_Room) ' VLM.Array;LM;LSling2
Dim LSling2_BM: LSling2_BM=Array(LSling2_BM_Dark_Room) ' VLM.Array;BM;LSling2
Dim LSling2_BL: LSling2_BL=Array(LSling2_BM_Dark_Room, LSling2_LM_GI1, LSling2_LM_Lit_Room) ' VLM.Array;BL;LSling2
Dim LockdownBar_LM: LockdownBar_LM=Array(LockdownBar_LM_Inserts_LightFir, LockdownBar_LM_Lit_Room) ' VLM.Array;LM;LockdownBar
Dim LockdownBar_BM: LockdownBar_BM=Array(LockdownBar_BM_Dark_Room) ' VLM.Array;BM;LockdownBar
Dim LockdownBar_BL: LockdownBar_BL=Array(LockdownBar_BM_Dark_Room, LockdownBar_LM_Inserts_LightFir, LockdownBar_LM_Lit_Room) ' VLM.Array;BL;LockdownBar
Dim LockdownButton_LM: LockdownButton_LM=Array(lockdown_LM_Inserts_LightApron, lockdown_LM_Inserts_LightFire, lockdown_LM_Lit_Room) ' VLM.Array;LM;LockdownButton
Dim LockdownButton_BM: LockdownButton_BM=Array(lockdown_BM_Dark_Room) ' VLM.Array;BM;LockdownButton
Dim LockdownButton_BL: LockdownButton_BL=Array(lockdown_BM_Dark_Room, lockdown_LM_Inserts_LightApron, lockdown_LM_Inserts_LightFire, lockdown_LM_Lit_Room) ' VLM.Array;BL;LockdownButton
Dim LorbitSW30_LM: LorbitSW30_LM=Array(LorbitSW30_LM_GI3) ' VLM.Array;LM;LorbitSW30
Dim LorbitSW30_BM: LorbitSW30_BM=Array(LorbitSW30_BM_Dark_Room) ' VLM.Array;BM;LorbitSW30
Dim LorbitSW30_BL: LorbitSW30_BL=Array(LorbitSW30_BM_Dark_Room, LorbitSW30_LM_GI3) ' VLM.Array;BL;LorbitSW30
Dim Lwing_LM: Lwing_LM=Array(Lwing_LM_GI2, Lwing_LM_Lit_Room, Lwing_LM_Upper_GI) ' VLM.Array;LM;Lwing
Dim Lwing_BM: Lwing_BM=Array(Lwing_BM_Dark_Room) ' VLM.Array;BM;Lwing
Dim Lwing_BL: Lwing_BL=Array(Lwing_BM_Dark_Room, Lwing_LM_GI2, Lwing_LM_Lit_Room, Lwing_LM_Upper_GI) ' VLM.Array;BL;Lwing
Dim Micro_Switch_LM: Micro_Switch_LM=Array(Micro_Switch_LM_GI3) ' VLM.Array;LM;Micro_Switch
Dim Micro_Switch_BM: Micro_Switch_BM=Array(Micro_Switch_BM_Dark_Room) ' VLM.Array;BM;Micro_Switch
Dim Micro_Switch_BL: Micro_Switch_BL=Array(Micro_Switch_BM_Dark_Room, Micro_Switch_LM_GI3) ' VLM.Array;BL;Micro_Switch
Dim Plastic_door_LM: Plastic_door_LM=Array(Plastic_door_LM_Flashers_f236, Plastic_door_LM_Flashers_f243, Plastic_door_LM_GI2, Plastic_door_LM_Lit_Room) ' VLM.Array;LM;Plastic_door
Dim Plastic_door_BM: Plastic_door_BM=Array(Plastic_door_BM_Dark_Room) ' VLM.Array;BM;Plastic_door
Dim Plastic_door_BL: Plastic_door_BL=Array(Plastic_door_BM_Dark_Room, Plastic_door_LM_Flashers_f236, Plastic_door_LM_Flashers_f243, Plastic_door_LM_GI2, Plastic_door_LM_Lit_Room) ' VLM.Array;BL;Plastic_door
Dim REMK_LM: REMK_LM=Array(REMK_LM_GI1, REMK_LM_Lit_Room) ' VLM.Array;LM;REMK
Dim REMK_BM: REMK_BM=Array(REMK_BM_Dark_Room) ' VLM.Array;BM;REMK
Dim REMK_BL: REMK_BL=Array(REMK_BM_Dark_Room, REMK_LM_GI1, REMK_LM_Lit_Room) ' VLM.Array;BL;REMK
Dim RFlogo_001_LM: RFlogo_001_LM=Array(RFlogo_001_LM_GI2, RFlogo_001_LM_Inserts_li198, RFlogo_001_LM_Inserts_li201, RFlogo_001_LM_Inserts_li204, RFlogo_001_LM_Inserts_li207, RFlogo_001_LM_Inserts_li210, RFlogo_001_LM_Inserts_li213, RFlogo_001_LM_Inserts_li216, RFlogo_001_LM_Lit_Room, RFlogo_001_LM_Upper_GI) ' VLM.Array;LM;RFlogo_001
Dim RFlogo_001_BM: RFlogo_001_BM=Array(RFlogo_001_BM_Dark_Room) ' VLM.Array;BM;RFlogo_001
Dim RFlogo_001_BL: RFlogo_001_BL=Array(RFlogo_001_BM_Dark_Room, RFlogo_001_LM_GI2, RFlogo_001_LM_Inserts_li198, RFlogo_001_LM_Inserts_li201, RFlogo_001_LM_Inserts_li204, RFlogo_001_LM_Inserts_li207, RFlogo_001_LM_Inserts_li210, RFlogo_001_LM_Inserts_li213, RFlogo_001_LM_Inserts_li216, RFlogo_001_LM_Lit_Room, RFlogo_001_LM_Upper_GI) ' VLM.Array;BL;RFlogo_001
Dim RFlogo_003_LM: RFlogo_003_LM=Array(RFlogo_003_LM_GI1, RFlogo_003_LM_Inserts_LightApro, RFlogo_003_LM_Inserts_LightShoo, RFlogo_003_LM_Inserts_li38, RFlogo_003_LM_Inserts_li41, RFlogo_003_LM_Inserts_li44, RFlogo_003_LM_Inserts_li47, RFlogo_003_LM_Inserts_li50, RFlogo_003_LM_Inserts_li53, RFlogo_003_LM_Inserts_li56, RFlogo_003_LM_Inserts_li59, RFlogo_003_LM_Inserts_li62, RFlogo_003_LM_Inserts_li65, RFlogo_003_LM_Lit_Room, RFlogo_003_LM_Upper_GI) ' VLM.Array;LM;RFlogo_003
Dim RFlogo_003_BM: RFlogo_003_BM=Array(RFlogo_003_BM_Dark_Room) ' VLM.Array;BM;RFlogo_003
Dim RFlogo_003_BL: RFlogo_003_BL=Array(RFlogo_003_BM_Dark_Room, RFlogo_003_LM_GI1, RFlogo_003_LM_Inserts_LightApro, RFlogo_003_LM_Inserts_LightShoo, RFlogo_003_LM_Inserts_li38, RFlogo_003_LM_Inserts_li41, RFlogo_003_LM_Inserts_li44, RFlogo_003_LM_Inserts_li47, RFlogo_003_LM_Inserts_li50, RFlogo_003_LM_Inserts_li53, RFlogo_003_LM_Inserts_li56, RFlogo_003_LM_Inserts_li59, RFlogo_003_LM_Inserts_li62, RFlogo_003_LM_Inserts_li65, RFlogo_003_LM_Lit_Room, RFlogo_003_LM_Upper_GI) ' VLM.Array;BL;RFlogo_003
Dim RSling1_LM: RSling1_LM=Array(RSling1_LM_Flashers_f244B, RSling1_LM_GI1, RSling1_LM_Lit_Room) ' VLM.Array;LM;RSling1
Dim RSling1_BM: RSling1_BM=Array(RSling1_BM_Dark_Room) ' VLM.Array;BM;RSling1
Dim RSling1_BL: RSling1_BL=Array(RSling1_BM_Dark_Room, RSling1_LM_Flashers_f244B, RSling1_LM_GI1, RSling1_LM_Lit_Room) ' VLM.Array;BL;RSling1
Dim RSling2_LM: RSling2_LM=Array(RSling2_LM_GI1, RSling2_LM_Lit_Room) ' VLM.Array;LM;RSling2
Dim RSling2_BM: RSling2_BM=Array(RSling2_BM_Dark_Room) ' VLM.Array;BM;RSling2
Dim RSling2_BL: RSling2_BL=Array(RSling2_BM_Dark_Room, RSling2_LM_GI1, RSling2_LM_Lit_Room) ' VLM.Array;BL;RSling2
Dim Rivet_1_002_LM: Rivet_1_002_LM=Array(Rivet_1_002_LM_Inserts_LightApr, Rivet_1_002_LM_Inserts_LightFir, Rivet_1_002_LM_Lit_Room) ' VLM.Array;LM;Rivet_1_002
Dim Rivet_1_002_BM: Rivet_1_002_BM=Array(Rivet_1_002_BM_Dark_Room) ' VLM.Array;BM;Rivet_1_002
Dim Rivet_1_002_BL: Rivet_1_002_BL=Array(Rivet_1_002_BM_Dark_Room, Rivet_1_002_LM_Inserts_LightApr, Rivet_1_002_LM_Inserts_LightFir, Rivet_1_002_LM_Lit_Room) ' VLM.Array;BL;Rivet_1_002
Dim RorbitSW31_LM: RorbitSW31_LM=Array(RorbitSW31_LM_Flashers_f237, RorbitSW31_LM_Flashers_f245, RorbitSW31_LM_GI2) ' VLM.Array;LM;RorbitSW31
Dim RorbitSW31_BM: RorbitSW31_BM=Array(RorbitSW31_BM_Dark_Room) ' VLM.Array;BM;RorbitSW31
Dim RorbitSW31_BL: RorbitSW31_BL=Array(RorbitSW31_BM_Dark_Room, RorbitSW31_LM_Flashers_f237, RorbitSW31_LM_Flashers_f245, RorbitSW31_LM_GI2) ' VLM.Array;BL;RorbitSW31
Dim Rwing_LM: Rwing_LM=Array(Rwing_LM_GI2, Rwing_LM_Lit_Room, Rwing_LM_Upper_GI) ' VLM.Array;LM;Rwing
Dim Rwing_BM: Rwing_BM=Array(Rwing_BM_Dark_Room) ' VLM.Array;BM;Rwing
Dim Rwing_BL: Rwing_BL=Array(Rwing_BM_Dark_Room, Rwing_LM_GI2, Rwing_LM_Lit_Room, Rwing_LM_Upper_GI) ' VLM.Array;BL;Rwing
Dim Spinners_LM: Spinners_LM=Array(Spinners_LM_Flashers_f235, Spinners_LM_Flashers_f236, Spinners_LM_Flashers_f242, Spinners_LM_Flashers_f243, Spinners_LM_Flashers_f244A, Spinners_LM_GI3, Spinners_LM_Inserts_li126, Spinners_LM_Inserts_li129, Spinners_LM_Inserts_li132, Spinners_LM_Inserts_li141, Spinners_LM_Inserts_li23, Spinners_LM_Inserts_li86, Spinners_LM_Inserts_li89, Spinners_LM_Inserts_li92, Spinners_LM_Lit_Room) ' VLM.Array;LM;Spinners
Dim Spinners_BM: Spinners_BM=Array(Spinners_BM_Dark_Room) ' VLM.Array;BM;Spinners
Dim Spinners_BL: Spinners_BL=Array(Spinners_BM_Dark_Room, Spinners_LM_Flashers_f235, Spinners_LM_Flashers_f236, Spinners_LM_Flashers_f242, Spinners_LM_Flashers_f243, Spinners_LM_Flashers_f244A, Spinners_LM_GI3, Spinners_LM_Inserts_li126, Spinners_LM_Inserts_li129, Spinners_LM_Inserts_li132, Spinners_LM_Inserts_li141, Spinners_LM_Inserts_li23, Spinners_LM_Inserts_li86, Spinners_LM_Inserts_li89, Spinners_LM_Inserts_li92, Spinners_LM_Lit_Room) ' VLM.Array;BL;Spinners
Dim T7_LM: T7_LM=Array(T7_LM_Flashers_f244A, T7_LM_Flashers_f244B, T7_LM_GI1, T7_LM_GI3, T7_LM_Inserts_li17, T7_LM_Inserts_li20, T7_LM_Inserts_li23, T7_LM_Inserts_li26, T7_LM_Inserts_li86, T7_LM_Inserts_li89, T7_LM_Inserts_li92, T7_LM_Lit_Room) ' VLM.Array;LM;T7
Dim T7_BM: T7_BM=Array(T7_BM_Dark_Room) ' VLM.Array;BM;T7
Dim T7_BL: T7_BL=Array(T7_BM_Dark_Room, T7_LM_Flashers_f244A, T7_LM_Flashers_f244B, T7_LM_GI1, T7_LM_GI3, T7_LM_Inserts_li17, T7_LM_Inserts_li20, T7_LM_Inserts_li23, T7_LM_Inserts_li26, T7_LM_Inserts_li86, T7_LM_Inserts_li89, T7_LM_Inserts_li92, T7_LM_Lit_Room) ' VLM.Array;BL;T7
Dim T8_LM: T8_LM=Array(T8_LM_Flashers_f242, T8_LM_Flashers_f244A, T8_LM_Flashers_f244B, T8_LM_GI1, T8_LM_GI3, T8_LM_Inserts_li17, T8_LM_Inserts_li20, T8_LM_Inserts_li23, T8_LM_Inserts_li26, T8_LM_Inserts_li86, T8_LM_Inserts_li89, T8_LM_Lit_Room) ' VLM.Array;LM;T8
Dim T8_BM: T8_BM=Array(T8_BM_Dark_Room) ' VLM.Array;BM;T8
Dim T8_BL: T8_BL=Array(T8_BM_Dark_Room, T8_LM_Flashers_f242, T8_LM_Flashers_f244A, T8_LM_Flashers_f244B, T8_LM_GI1, T8_LM_GI3, T8_LM_Inserts_li17, T8_LM_Inserts_li20, T8_LM_Inserts_li23, T8_LM_Inserts_li26, T8_LM_Inserts_li86, T8_LM_Inserts_li89, T8_LM_Lit_Room) ' VLM.Array;BL;T8
Dim T9_LM: T9_LM=Array(T9_LM_Flashers_f244A, T9_LM_Flashers_f244B, T9_LM_GI1, T9_LM_GI3, T9_LM_Inserts_li14, T9_LM_Inserts_li17, T9_LM_Inserts_li20, T9_LM_Inserts_li23, T9_LM_Inserts_li26, T9_LM_Inserts_li29, T9_LM_Lit_Room) ' VLM.Array;LM;T9
Dim T9_BM: T9_BM=Array(T9_BM_Dark_Room) ' VLM.Array;BM;T9
Dim T9_BL: T9_BL=Array(T9_BM_Dark_Room, T9_LM_Flashers_f244A, T9_LM_Flashers_f244B, T9_LM_GI1, T9_LM_GI3, T9_LM_Inserts_li14, T9_LM_Inserts_li17, T9_LM_Inserts_li20, T9_LM_Inserts_li23, T9_LM_Inserts_li26, T9_LM_Inserts_li29, T9_LM_Lit_Room) ' VLM.Array;BL;T9
Dim T90_LM: T90_LM=Array(T90_LM_Flashers_f235, T90_LM_Flashers_f242, T90_LM_GI1, T90_LM_GI2, T90_LM_GI3, T90_LM_Inserts_li123, T90_LM_Inserts_li126, T90_LM_Inserts_li129, T90_LM_Inserts_li132, T90_LM_Inserts_li135, T90_LM_Inserts_li141, T90_LM_Inserts_li144, T90_LM_Inserts_li153, T90_LM_Inserts_li156, T90_LM_Inserts_li159, T90_LM_Lit_Room) ' VLM.Array;LM;T90
Dim T90_BM: T90_BM=Array(T90_BM_Dark_Room) ' VLM.Array;BM;T90
Dim T90_BL: T90_BL=Array(T90_BM_Dark_Room, T90_LM_Flashers_f235, T90_LM_Flashers_f242, T90_LM_GI1, T90_LM_GI2, T90_LM_GI3, T90_LM_Inserts_li123, T90_LM_Inserts_li126, T90_LM_Inserts_li129, T90_LM_Inserts_li132, T90_LM_Inserts_li135, T90_LM_Inserts_li141, T90_LM_Inserts_li144, T90_LM_Inserts_li153, T90_LM_Inserts_li156, T90_LM_Inserts_li159, T90_LM_Lit_Room) ' VLM.Array;BL;T90
Dim fb_base_LM: fb_base_LM=Array(fb_base_LM_Flashers_f242, fb_base_LM_GI1, fb_base_LM_Inserts_LightApron, fb_base_LM_Inserts_LightFire, fb_base_LM_Lit_Room) ' VLM.Array;LM;fb_base
Dim fb_base_BM: fb_base_BM=Array(fb_base_BM_Dark_Room) ' VLM.Array;BM;fb_base
Dim fb_base_BL: fb_base_BL=Array(fb_base_BM_Dark_Room, fb_base_LM_Flashers_f242, fb_base_LM_GI1, fb_base_LM_Inserts_LightApron, fb_base_LM_Inserts_LightFire, fb_base_LM_Lit_Room) ' VLM.Array;BL;fb_base
Dim fl_base_LM: fl_base_LM=Array(fl_base_LM_Inserts_LightApron, fl_base_LM_Inserts_LightFire, fl_base_LM_Lit_Room) ' VLM.Array;LM;fl_base
Dim fl_base_BM: fl_base_BM=Array(fl_base_BM_Dark_Room) ' VLM.Array;BM;fl_base
Dim fl_base_BL: fl_base_BL=Array(fl_base_BM_Dark_Room, fl_base_LM_Inserts_LightApron, fl_base_LM_Inserts_LightFire, fl_base_LM_Lit_Room) ' VLM.Array;BL;fl_base
Dim lockpost_001_LM: lockpost_001_LM=Array(lockpost_001_LM_Flashers_f235, lockpost_001_LM_Flashers_f244A, lockpost_001_LM_Flashers_f244B, lockpost_001_LM_GI1, lockpost_001_LM_Inserts_li80, lockpost_001_LM_Inserts_li83, lockpost_001_LM_Lit_Room) ' VLM.Array;LM;lockpost_001
Dim lockpost_001_BM: lockpost_001_BM=Array(lockpost_001_BM_Dark_Room) ' VLM.Array;BM;lockpost_001
Dim lockpost_001_BL: lockpost_001_BL=Array(lockpost_001_BM_Dark_Room, lockpost_001_LM_Flashers_f235, lockpost_001_LM_Flashers_f244A, lockpost_001_LM_Flashers_f244B, lockpost_001_LM_GI1, lockpost_001_LM_Inserts_li80, lockpost_001_LM_Inserts_li83, lockpost_001_LM_Lit_Room) ' VLM.Array;BL;lockpost_001
Dim metalPlates_002_LM: metalPlates_002_LM=Array(metalPlates_002_LM_Flashers_f23, metalPlates_002_LM_Flashers_f23, metalPlates_002_LM_Flashers_f24, metalPlates_002_LM_GI2, metalPlates_002_LM_GI3, metalPlates_002_LM_Lit_Room, metalPlates_002_LM_Upper_GI) ' VLM.Array;LM;metalPlates_002
Dim metalPlates_002_BM: metalPlates_002_BM=Array(metalPlates_002_BM_Dark_Room) ' VLM.Array;BM;metalPlates_002
Dim metalPlates_002_BL: metalPlates_002_BL=Array(metalPlates_002_BM_Dark_Room, metalPlates_002_LM_Flashers_f23, metalPlates_002_LM_Flashers_f23, metalPlates_002_LM_Flashers_f24, metalPlates_002_LM_GI2, metalPlates_002_LM_GI3, metalPlates_002_LM_Lit_Room, metalPlates_002_LM_Upper_GI) ' VLM.Array;BL;metalPlates_002
Dim metal_plate_LM: metal_plate_LM=Array(metal_plate_LM_Inserts_LightApr, metal_plate_LM_Inserts_LightFir, metal_plate_LM_Lit_Room) ' VLM.Array;LM;metal_plate
Dim metal_plate_BM: metal_plate_BM=Array(metal_plate_BM_Dark_Room) ' VLM.Array;BM;metal_plate
Dim metal_plate_BL: metal_plate_BL=Array(metal_plate_BM_Dark_Room, metal_plate_LM_Inserts_LightApr, metal_plate_LM_Inserts_LightFir, metal_plate_LM_Lit_Room) ' VLM.Array;BL;metal_plate
Dim outlane_left_LM: outlane_left_LM=Array(outlane_left_LM_Flashers_f235, outlane_left_LM_GI1, outlane_left_LM_Inserts_li11, outlane_left_LM_Lit_Room) ' VLM.Array;LM;outlane_left
Dim outlane_left_BM: outlane_left_BM=Array(outlane_left_BM_Dark_Room) ' VLM.Array;BM;outlane_left
Dim outlane_left_BL: outlane_left_BL=Array(outlane_left_BM_Dark_Room, outlane_left_LM_Flashers_f235, outlane_left_LM_GI1, outlane_left_LM_Inserts_li11, outlane_left_LM_Lit_Room) ' VLM.Array;BL;outlane_left
Dim outlane_right_LM: outlane_right_LM=Array(outlane_right_LM_GI1, outlane_right_LM_Lit_Room) ' VLM.Array;LM;outlane_right
Dim outlane_right_BM: outlane_right_BM=Array(outlane_right_BM_Dark_Room) ' VLM.Array;BM;outlane_right
Dim outlane_right_BL: outlane_right_BL=Array(outlane_right_BM_Dark_Room, outlane_right_LM_GI1, outlane_right_LM_Lit_Room) ' VLM.Array;BL;outlane_right
Dim ram_LM: ram_LM=Array(ram_LM_Flashers_f235, ram_LM_Flashers_f236, ram_LM_Flashers_f243, ram_LM_GI1, ram_LM_GI2, ram_LM_GI3, ram_LM_Inserts_li120, ram_LM_Inserts_li132, ram_LM_Inserts_li135, ram_LM_Inserts_li159, ram_LM_Lit_Room) ' VLM.Array;LM;ram
Dim ram_BM: ram_BM=Array(ram_BM_Dark_Room) ' VLM.Array;BM;ram
Dim ram_BL: ram_BL=Array(ram_BM_Dark_Room, ram_LM_Flashers_f235, ram_LM_Flashers_f236, ram_LM_Flashers_f243, ram_LM_GI1, ram_LM_GI2, ram_LM_GI3, ram_LM_Inserts_li120, ram_LM_Inserts_li132, ram_LM_Inserts_li135, ram_LM_Inserts_li159, ram_LM_Lit_Room) ' VLM.Array;BL;ram
Dim sw1_LM: sw1_LM=Array(sw1_LM_GI1, sw1_LM_GI3, sw1_LM_Lit_Room) ' VLM.Array;LM;sw1
Dim sw1_BM: sw1_BM=Array(sw1_BM_Dark_Room) ' VLM.Array;BM;sw1
Dim sw1_BL: sw1_BL=Array(sw1_BM_Dark_Room, sw1_LM_GI1, sw1_LM_GI3, sw1_LM_Lit_Room) ' VLM.Array;BL;sw1
Dim sw2_LM: sw2_LM=Array(sw2_LM_GI1, sw2_LM_Lit_Room) ' VLM.Array;LM;sw2
Dim sw2_BM: sw2_BM=Array(sw2_BM_Dark_Room) ' VLM.Array;BM;sw2
Dim sw2_BL: sw2_BL=Array(sw2_BM_Dark_Room, sw2_LM_GI1, sw2_LM_Lit_Room) ' VLM.Array;BL;sw2
Dim sw3_LM: sw3_LM=Array(sw3_LM_GI1, sw3_LM_GI3, sw3_LM_Lit_Room) ' VLM.Array;LM;sw3
Dim sw3_BM: sw3_BM=Array(sw3_BM_Dark_Room) ' VLM.Array;BM;sw3
Dim sw3_BL: sw3_BL=Array(sw3_BM_Dark_Room, sw3_LM_GI1, sw3_LM_GI3, sw3_LM_Lit_Room) ' VLM.Array;BL;sw3
Dim sw4_LM: sw4_LM=Array(sw4_LM_GI1, sw4_LM_Lit_Room) ' VLM.Array;LM;sw4
Dim sw4_BM: sw4_BM=Array(sw4_BM_Dark_Room) ' VLM.Array;BM;sw4
Dim sw4_BL: sw4_BL=Array(sw4_BM_Dark_Room, sw4_LM_GI1, sw4_LM_Lit_Room) ' VLM.Array;BL;sw4
Dim sw53_LM: sw53_LM=Array(sw53_LM_Flashers_f237, sw53_LM_Flashers_f243, sw53_LM_Flashers_f245, sw53_LM_GI2, sw53_LM_Inserts_li168) ' VLM.Array;LM;sw53
Dim sw53_BM: sw53_BM=Array(sw53_BM_Dark_Room) ' VLM.Array;BM;sw53
Dim sw53_BL: sw53_BL=Array(sw53_BM_Dark_Room, sw53_LM_Flashers_f237, sw53_LM_Flashers_f243, sw53_LM_Flashers_f245, sw53_LM_GI2, sw53_LM_Inserts_li168) ' VLM.Array;BL;sw53
Dim sw54_LM: sw54_LM=Array(sw54_LM_Flashers_f237, sw54_LM_Flashers_f245, sw54_LM_GI2) ' VLM.Array;LM;sw54
Dim sw54_BM: sw54_BM=Array(sw54_BM_Dark_Room) ' VLM.Array;BM;sw54
Dim sw54_BL: sw54_BL=Array(sw54_BM_Dark_Room, sw54_LM_Flashers_f237, sw54_LM_Flashers_f245, sw54_LM_GI2) ' VLM.Array;BL;sw54
Dim sw77_LM: sw77_LM=Array(sw77_LM_GI2, sw77_LM_Lit_Room, sw77_LM_Upper_GI) ' VLM.Array;LM;sw77
Dim sw77_BM: sw77_BM=Array(sw77_BM_Dark_Room) ' VLM.Array;BM;sw77
Dim sw77_BL: sw77_BL=Array(sw77_BM_Dark_Room, sw77_LM_GI2, sw77_LM_Lit_Room, sw77_LM_Upper_GI) ' VLM.Array;BL;sw77
Dim sw78_LM: sw78_LM=Array(sw78_LM_GI2, sw78_LM_Lit_Room, sw78_LM_Upper_GI) ' VLM.Array;LM;sw78
Dim sw78_BM: sw78_BM=Array(sw78_BM_Dark_Room) ' VLM.Array;BM;sw78
Dim sw78_BL: sw78_BL=Array(sw78_BM_Dark_Room, sw78_LM_GI2, sw78_LM_Lit_Room, sw78_LM_Upper_GI) ' VLM.Array;BL;sw78
Dim sw79_LM: sw79_LM=Array(sw79_LM_GI2, sw79_LM_Upper_GI) ' VLM.Array;LM;sw79
Dim sw79_BM: sw79_BM=Array(sw79_BM_Dark_Room) ' VLM.Array;BM;sw79
Dim sw79_BL: sw79_BL=Array(sw79_BM_Dark_Room, sw79_LM_GI2, sw79_LM_Upper_GI) ' VLM.Array;BL;sw79
Dim sw83_LM: sw83_LM=Array(sw83_LM_Lit_Room, sw83_LM_Upper_GI) ' VLM.Array;LM;sw83
Dim sw83_BM: sw83_BM=Array(sw83_BM_Dark_Room) ' VLM.Array;BM;sw83
Dim sw83_BL: sw83_BL=Array(sw83_BM_Dark_Room, sw83_LM_Lit_Room, sw83_LM_Upper_GI) ' VLM.Array;BL;sw83
Dim sw84_LM: sw84_LM=Array(sw84_LM_Lit_Room, sw84_LM_Upper_GI) ' VLM.Array;LM;sw84
Dim sw84_BM: sw84_BM=Array(sw84_BM_Dark_Room) ' VLM.Array;BM;sw84
Dim sw84_BL: sw84_BL=Array(sw84_BM_Dark_Room, sw84_LM_Lit_Room, sw84_LM_Upper_GI) ' VLM.Array;BL;sw84
Dim sword_002_LM: sword_002_LM=Array(sword_002_LM_Flashers_f235, sword_002_LM_GI1, sword_002_LM_GI2, sword_002_LM_Inserts_li80, sword_002_LM_Lit_Room) ' VLM.Array;LM;sword_002
Dim sword_002_BM: sword_002_BM=Array(sword_002_BM_Dark_Room) ' VLM.Array;BM;sword_002
Dim sword_002_BL: sword_002_BL=Array(sword_002_BM_Dark_Room, sword_002_LM_Flashers_f235, sword_002_LM_GI1, sword_002_LM_GI2, sword_002_LM_Inserts_li80, sword_002_LM_Lit_Room) ' VLM.Array;BL;sword_002
Dim sword_003_LM: sword_003_LM=Array(sword_003_LM_Flashers_f235, sword_003_LM_GI1, sword_003_LM_Inserts_li77, sword_003_LM_Inserts_li80, sword_003_LM_Inserts_li83, sword_003_LM_Lit_Room) ' VLM.Array;LM;sword_003
Dim sword_003_BM: sword_003_BM=Array(sword_003_BM_Dark_Room) ' VLM.Array;BM;sword_003
Dim sword_003_BL: sword_003_BL=Array(sword_003_BM_Dark_Room, sword_003_LM_Flashers_f235, sword_003_LM_GI1, sword_003_LM_Inserts_li77, sword_003_LM_Inserts_li80, sword_003_LM_Inserts_li83, sword_003_LM_Lit_Room) ' VLM.Array;BL;sword_003
Dim targets_LM: targets_LM=Array(targets_LM_Inserts_li104, targets_LM_Inserts_li189, targets_LM_Lit_Room, targets_LM_Upper_GI) ' VLM.Array;LM;targets
Dim targets_BM: targets_BM=Array(targets_BM_Dark_Room) ' VLM.Array;BM;targets
Dim targets_BL: targets_BL=Array(targets_BM_Dark_Room, targets_LM_Inserts_li104, targets_LM_Inserts_li189, targets_LM_Lit_Room, targets_LM_Upper_GI) ' VLM.Array;BL;targets
Dim targets_004_LM: targets_004_LM=Array(targets_004_LM_Flashers_f244A, targets_004_LM_Flashers_f244B, targets_004_LM_GI1, targets_004_LM_Inserts_li86, targets_004_LM_Inserts_li92, targets_004_LM_Inserts_li95, targets_004_LM_Lit_Room) ' VLM.Array;LM;targets_004
Dim targets_004_BM: targets_004_BM=Array(targets_004_BM_Dark_Room) ' VLM.Array;BM;targets_004
Dim targets_004_BL: targets_004_BL=Array(targets_004_BM_Dark_Room, targets_004_LM_Flashers_f244A, targets_004_LM_Flashers_f244B, targets_004_LM_GI1, targets_004_LM_Inserts_li86, targets_004_LM_Inserts_li92, targets_004_LM_Inserts_li95, targets_004_LM_Lit_Room) ' VLM.Array;BL;targets_004
Dim targets_005_LM: targets_005_LM=Array(targets_005_LM_Flashers_f235, targets_005_LM_Flashers_f242, targets_005_LM_Flashers_f244A, targets_005_LM_Flashers_f244B, targets_005_LM_Inserts_li105, targets_005_LM_Inserts_li117, targets_005_LM_Lit_Room) ' VLM.Array;LM;targets_005
Dim targets_005_BM: targets_005_BM=Array(targets_005_BM_Dark_Room) ' VLM.Array;BM;targets_005
Dim targets_005_BL: targets_005_BL=Array(targets_005_BM_Dark_Room, targets_005_LM_Flashers_f235, targets_005_LM_Flashers_f242, targets_005_LM_Flashers_f244A, targets_005_LM_Flashers_f244B, targets_005_LM_Inserts_li105, targets_005_LM_Inserts_li117, targets_005_LM_Lit_Room) ' VLM.Array;BL;targets_005
Dim targets_006_LM: targets_006_LM=Array(targets_006_LM_Flashers_f235, targets_006_LM_Flashers_f244B, targets_006_LM_Inserts_li117, targets_006_LM_Inserts_li120, targets_006_LM_Inserts_li132, targets_006_LM_Lit_Room) ' VLM.Array;LM;targets_006
Dim targets_006_BM: targets_006_BM=Array(targets_006_BM_Dark_Room) ' VLM.Array;BM;targets_006
Dim targets_006_BL: targets_006_BL=Array(targets_006_BM_Dark_Room, targets_006_LM_Flashers_f235, targets_006_LM_Flashers_f244B, targets_006_LM_Inserts_li117, targets_006_LM_Inserts_li120, targets_006_LM_Inserts_li132, targets_006_LM_Lit_Room) ' VLM.Array;BL;targets_006
Dim targets_007_LM: targets_007_LM=Array(targets_007_LM_Flashers_f235, targets_007_LM_Inserts_li135, targets_007_LM_Lit_Room) ' VLM.Array;LM;targets_007
Dim targets_007_BM: targets_007_BM=Array(targets_007_BM_Dark_Room) ' VLM.Array;BM;targets_007
Dim targets_007_BL: targets_007_BL=Array(targets_007_BM_Dark_Room, targets_007_LM_Flashers_f235, targets_007_LM_Inserts_li135, targets_007_LM_Lit_Room) ' VLM.Array;BL;targets_007
Dim targets_008_LM: targets_008_LM=Array(targets_008_LM_Flashers_f235, targets_008_LM_GI1, targets_008_LM_Inserts_li147, targets_008_LM_Lit_Room) ' VLM.Array;LM;targets_008
Dim targets_008_BM: targets_008_BM=Array(targets_008_BM_Dark_Room) ' VLM.Array;BM;targets_008
Dim targets_008_BL: targets_008_BL=Array(targets_008_BM_Dark_Room, targets_008_LM_Flashers_f235, targets_008_LM_GI1, targets_008_LM_Inserts_li147, targets_008_LM_Lit_Room) ' VLM.Array;BL;targets_008
Dim targets_009_LM: targets_009_LM=Array(targets_009_LM_Flashers_f235, targets_009_LM_Flashers_f244A, targets_009_LM_Flashers_f244B, targets_009_LM_GI1, targets_009_LM_Inserts_li153, targets_009_LM_Inserts_li80, targets_009_LM_Lit_Room) ' VLM.Array;LM;targets_009
Dim targets_009_BM: targets_009_BM=Array(targets_009_BM_Dark_Room) ' VLM.Array;BM;targets_009
Dim targets_009_BL: targets_009_BL=Array(targets_009_BM_Dark_Room, targets_009_LM_Flashers_f235, targets_009_LM_Flashers_f244A, targets_009_LM_Flashers_f244B, targets_009_LM_GI1, targets_009_LM_Inserts_li153, targets_009_LM_Inserts_li80, targets_009_LM_Lit_Room) ' VLM.Array;BL;targets_009
Dim targets_010_LM: targets_010_LM=Array(targets_010_LM_Flashers_f235, targets_010_LM_Flashers_f244A, targets_010_LM_GI1, targets_010_LM_Inserts_li80, targets_010_LM_Inserts_li83, targets_010_LM_Lit_Room) ' VLM.Array;LM;targets_010
Dim targets_010_BM: targets_010_BM=Array(targets_010_BM_Dark_Room) ' VLM.Array;BM;targets_010
Dim targets_010_BL: targets_010_BL=Array(targets_010_BM_Dark_Room, targets_010_LM_Flashers_f235, targets_010_LM_Flashers_f244A, targets_010_LM_GI1, targets_010_LM_Inserts_li80, targets_010_LM_Inserts_li83, targets_010_LM_Lit_Room) ' VLM.Array;BL;targets_010
Dim targets_011_LM: targets_011_LM=Array(targets_011_LM_Inserts_li104, targets_011_LM_Inserts_li183, targets_011_LM_Inserts_li192, targets_011_LM_Inserts_li195, targets_011_LM_Lit_Room, targets_011_LM_Upper_GI) ' VLM.Array;LM;targets_011
Dim targets_011_BM: targets_011_BM=Array(targets_011_BM_Dark_Room) ' VLM.Array;BM;targets_011
Dim targets_011_BL: targets_011_BL=Array(targets_011_BM_Dark_Room, targets_011_LM_Inserts_li104, targets_011_LM_Inserts_li183, targets_011_LM_Inserts_li192, targets_011_LM_Inserts_li195, targets_011_LM_Lit_Room, targets_011_LM_Upper_GI) ' VLM.Array;BL;targets_011
Dim targets_012_LM: targets_012_LM=Array(targets_012_LM_Inserts_li104, targets_012_LM_Inserts_li180, targets_012_LM_Inserts_li183, targets_012_LM_Inserts_li189, targets_012_LM_Inserts_li192, targets_012_LM_Inserts_li195, targets_012_LM_Inserts_li92, targets_012_LM_Lit_Room, targets_012_LM_Upper_GI) ' VLM.Array;LM;targets_012
Dim targets_012_BM: targets_012_BM=Array(targets_012_BM_Dark_Room) ' VLM.Array;BM;targets_012
Dim targets_012_BL: targets_012_BL=Array(targets_012_BM_Dark_Room, targets_012_LM_Inserts_li104, targets_012_LM_Inserts_li180, targets_012_LM_Inserts_li183, targets_012_LM_Inserts_li189, targets_012_LM_Inserts_li192, targets_012_LM_Inserts_li195, targets_012_LM_Inserts_li92, targets_012_LM_Lit_Room, targets_012_LM_Upper_GI) ' VLM.Array;BL;targets_012


' ===============================================================
' The following code can be copy/pasted if using Lampz fading system
' It links each Lampz lamp/flasher to the corresponding light and lightmap

Sub UpdateLightMap(lightmap, intensity, ByVal aLvl)
   if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl) 'Callbacks don't get this filter automatically
   lightmap.Opacity = aLvl * intensity
End Sub

Sub UpdateLightMapWithColor(light, lightmap, intensity, ByVal aLvl)
   if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl) 'Callbacks don't get this filter automatically
   lightmap.Opacity = aLvl * intensity
   lightmap.color = light.colorfull
End Sub

Sub LampzHelper
  ' Lampz.MassAssign(150) = L150 ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Playfield_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Layer1_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Layer2_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Layer3_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Layer4_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Layer5_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Bumper_Ring_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Bumper_Ring_001_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Bumper_Ring_002_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Gates_001_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap LEMK_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap LFlogo_001_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap LFlogo_003_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap LSling1_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap LSling2_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap LockdownBar_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Lwing_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Plastic_door_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap REMK_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap RFlogo_001_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap RFlogo_003_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap RSling1_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap RSling2_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Rivet_1_002_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Rwing_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Spinners_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap T7_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap T8_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap T9_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap T90_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap fb_base_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap fl_base_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap lockpost_001_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap metalPlates_002_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap metal_plate_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap outlane_left_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap outlane_right_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap ram_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sw1_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sw2_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sw3_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sw4_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sw77_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sw78_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sw83_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sw84_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sword_002_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap sword_003_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_004_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_005_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_006_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_007_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_008_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_009_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_010_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_011_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap targets_012_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap FireButton_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap lockdown_LM_Lit_Room, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Playfield_LM_Lit_Room_001, 100.0, " ' VLM.Lampz;Lit Room
  Lampz.Callback(150) = "UpdateLightMap Playfield_LM_Lit_Room_002, 100.0, " ' VLM.Lampz;Lit Room
  ' Sync on LightApron ' VLM.Lampz;Inserts-LightApron
  Lampz.Callback(79) = "UpdateLightMapWithColor li231,Playfield_LM_Inserts_LightApron, 100.0, " ' VLM.Lampz;Inserts-LightApron
  Lampz.Callback(79) = "UpdateLightMapWithColor li231,LFlogo_003_LM_Inserts_LightApro, 100.0, " ' VLM.Lampz;Inserts-LightApron
  Lampz.Callback(79) = "UpdateLightMapWithColor li231,RFlogo_003_LM_Inserts_LightApro, 100.0, " ' VLM.Lampz;Inserts-LightApron
  Lampz.Callback(79) = "UpdateLightMapWithColor li231,Rivet_1_002_LM_Inserts_LightApr, 100.0, " ' VLM.Lampz;Inserts-LightApron
  Lampz.Callback(79) = "UpdateLightMapWithColor li231,fb_base_LM_Inserts_LightApron, 100.0, " ' VLM.Lampz;Inserts-LightApron
  Lampz.Callback(79) = "UpdateLightMapWithColor li231,fl_base_LM_Inserts_LightApron, 100.0, " ' VLM.Lampz;Inserts-LightApron
  Lampz.Callback(79) = "UpdateLightMapWithColor li231,metal_plate_LM_Inserts_LightApr, 100.0, " ' VLM.Lampz;Inserts-LightApron
  Lampz.Callback(79) = "UpdateLightMapWithColor li231,FireButton_LM_Inserts_LightApro, 100.0, " ' VLM.Lampz;Inserts-LightApron
  Lampz.Callback(79) = "UpdateLightMapWithColor li231,lockdown_LM_Inserts_LightApron, 100.0, " ' VLM.Lampz;Inserts-LightApron
  ' Sync on LightFire ' VLM.Lampz;Inserts-LightFire
  Lampz.Callback(78) = "UpdateLightMapWithColor li230,Playfield_LM_Inserts_LightFire, 100.0, " ' VLM.Lampz;Inserts-LightFire
  Lampz.Callback(78) = "UpdateLightMapWithColor li230,LockdownBar_LM_Inserts_LightFir, 100.0, " ' VLM.Lampz;Inserts-LightFire
  Lampz.Callback(78) = "UpdateLightMapWithColor li230,Rivet_1_002_LM_Inserts_LightFir, 100.0, " ' VLM.Lampz;Inserts-LightFire
  Lampz.Callback(78) = "UpdateLightMapWithColor li230,fb_base_LM_Inserts_LightFire, 100.0, " ' VLM.Lampz;Inserts-LightFire
  Lampz.Callback(78) = "UpdateLightMapWithColor li230,fl_base_LM_Inserts_LightFire, 100.0, " ' VLM.Lampz;Inserts-LightFire
  Lampz.Callback(78) = "UpdateLightMapWithColor li230,metal_plate_LM_Inserts_LightFir, 100.0, " ' VLM.Lampz;Inserts-LightFire
  Lampz.Callback(78) = "UpdateLightMapWithColor li230,lockdown_LM_Inserts_LightFire, 100.0, " ' VLM.Lampz;Inserts-LightFire
  ' Sync on LightShootAgain ' VLM.Lampz;Inserts-LightShootAgain
  Lampz.Callback(1) = "UpdateLightMapWithColor LightShootAgain,Playfield_LM_Inserts_LightShoot, 100.0, " ' VLM.Lampz;Inserts-LightShootAgain
  Lampz.Callback(1) = "UpdateLightMapWithColor LightShootAgain,LFlogo_003_LM_Inserts_LightShoo, 100.0, " ' VLM.Lampz;Inserts-LightShootAgain
  Lampz.Callback(1) = "UpdateLightMapWithColor LightShootAgain,RFlogo_003_LM_Inserts_LightShoo, 100.0, " ' VLM.Lampz;Inserts-LightShootAgain
  ' Lampz.MassAssign(1235) = f235 ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap Playfield_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap Layer3_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap Bumper_Ring_001_LM_Flashers_f23, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap Spinners_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap T90_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap lockpost_001_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap outlane_left_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap ram_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap sword_002_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap sword_003_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap targets_005_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap targets_006_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap targets_007_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap targets_008_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap targets_009_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  Lampz.Callback(70) = "UpdateLightMap targets_010_LM_Flashers_f235, 100.0, " ' VLM.Lampz;Flashers-f235
  ' Lampz.MassAssign(1236) = f236 ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap Playfield_LM_Flashers_f236, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap Layer1_LM_Flashers_f236, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap Layer3_LM_Flashers_f236, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap Layer4_LM_Flashers_f236, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap Bumper_Ring_LM_Flashers_f236, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap Bumper_Ring_001_LM_Flashers_f23, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap Bumper_Ring_002_LM_Flashers_f23, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap Plastic_door_LM_Flashers_f236, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap Spinners_LM_Flashers_f236, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap metalPlates_002_LM_Flashers_f23, 100.0, " ' VLM.Lampz;Flashers-f236
  Lampz.Callback(71) = "UpdateLightMap ram_LM_Flashers_f236, 100.0, " ' VLM.Lampz;Flashers-f236
  ' Lampz.MassAssign(1237) = f237 ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap Playfield_LM_Flashers_f237, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap Layer1_LM_Flashers_f237, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap Layer4_LM_Flashers_f237, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap Bumper_Ring_LM_Flashers_f237, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap Bumper_Ring_001_LM_Flashers_f23, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap Bumper_Ring_002_LM_Flashers_f23, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap Gates_001_LM_Flashers_f237, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap RorbitSW31_LM_Flashers_f237, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap metalPlates_002_LM_Flashers_f23, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap sw53_LM_Flashers_f237, 100.0, " ' VLM.Lampz;Flashers-f237
  Lampz.Callback(72) = "UpdateLightMap sw54_LM_Flashers_f237, 100.0, " ' VLM.Lampz;Flashers-f237
  ' Lampz.MassAssign(1242) = f242 ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap Playfield_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap Layer1_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap Layer3_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap Layer5_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap LFlogo_001_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap Spinners_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap T8_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap T90_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap fb_base_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap targets_005_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  Lampz.Callback(73) = "UpdateLightMap FireButton_LM_Flashers_f242, 100.0, " ' VLM.Lampz;Flashers-f242
  ' Lampz.MassAssign(1243) = f243 ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Playfield_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Layer1_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Layer3_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Layer4_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Layer5_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Bumper_Ring_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Bumper_Ring_001_LM_Flashers_f24, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Bumper_Ring_002_LM_Flashers_f24, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Gates_001_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Plastic_door_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap Spinners_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap metalPlates_002_LM_Flashers_f24, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap ram_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  Lampz.Callback(74) = "UpdateLightMap sw53_LM_Flashers_f243, 100.0, " ' VLM.Lampz;Flashers-f243
  ' Lampz.MassAssign(1244) = f244A ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap Playfield_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap Spinners_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap T7_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap T8_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap T9_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap lockpost_001_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap targets_004_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap targets_005_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap targets_009_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap targets_010_LM_Flashers_f244A, 100.0, " ' VLM.Lampz;Flashers-f244A
  Lampz.Callback(75) = "UpdateLightMap Playfield_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap Layer3_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap RSling1_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap T7_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap T8_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap T9_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap lockpost_001_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap targets_004_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap targets_005_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap targets_006_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  Lampz.Callback(75) = "UpdateLightMap targets_009_LM_Flashers_f244B, 100.0, " ' VLM.Lampz;Flashers-f244B
  ' Lampz.MassAssign(1245) = f245 ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap Playfield_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap Layer1_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap Layer2_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap Layer3_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap Layer4_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap Layer5_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap Bumper_Ring_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap Bumper_Ring_001_LM_Flashers_f24, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap Gates_001_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap RorbitSW31_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap sw53_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  Lampz.Callback(76) = "UpdateLightMap sw54_LM_Flashers_f245, 100.0, " ' VLM.Lampz;Flashers-f245
  ' Sync on gi010 ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap Playfield_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap Layer3_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap LEMK_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap LFlogo_003_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap LSling1_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap LSling2_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap REMK_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap RFlogo_003_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap RSling1_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap RSling2_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap T7_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap T8_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap T9_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap T90_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap fb_base_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap lockpost_001_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap outlane_left_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap outlane_right_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap ram_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap sw1_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap sw2_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap sw3_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap sw4_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap sword_002_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap sword_003_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap targets_004_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap targets_008_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap targets_009_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap targets_010_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  Lampz.Callback(80) = "UpdateLightMap FireButton_LM_GI1, 100.0, " ' VLM.Lampz;GI1
  ' Sync on gi026 ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap Playfield_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap Layer1_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap Layer3_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap Layer5_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap LorbitSW30_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap Micro_Switch_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap Spinners_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap T7_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap T8_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap T9_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap T90_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap metalPlates_002_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap ram_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap sw1_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  Lampz.Callback(82) = "UpdateLightMap sw3_LM_GI3, 100.0, " ' VLM.Lampz;GI3
  ' Sync on gi035 ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Playfield_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Layer1_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Layer2_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Layer3_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Layer4_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Layer5_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Bumper_Ring_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Bumper_Ring_001_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Bumper_Ring_002_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Gates_001_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Lwing_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Plastic_door_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap RFlogo_001_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap RorbitSW31_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap Rwing_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap T90_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap metalPlates_002_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap ram_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap sw53_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap sw54_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap sw77_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap sw78_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap sw79_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  Lampz.Callback(81) = "UpdateLightMap sword_002_LM_GI2, 100.0, " ' VLM.Lampz;GI2
  ' Sync on li101 ' VLM.Lampz;Inserts-li101
  Lampz.Callback(31) = "UpdateLightMapWithColor li101,Playfield_LM_Inserts_li101, 100.0, " ' VLM.Lampz;Inserts-li101
  ' Sync on li104 ' VLM.Lampz;Inserts-li104
  Lampz.Callback(79) = "UpdateLightMapWithColor li104,Playfield_LM_Inserts_li104, 100.0, " ' VLM.Lampz;Inserts-li104
  Lampz.Callback(79) = "UpdateLightMapWithColor li104,targets_LM_Inserts_li104, 100.0, " ' VLM.Lampz;Inserts-li104
  Lampz.Callback(79) = "UpdateLightMapWithColor li104,targets_011_LM_Inserts_li104, 100.0, " ' VLM.Lampz;Inserts-li104
  Lampz.Callback(79) = "UpdateLightMapWithColor li104,targets_012_LM_Inserts_li104, 100.0, " ' VLM.Lampz;Inserts-li104
  ' Sync on li105 ' VLM.Lampz;Inserts-li105
  Lampz.Callback(32) = "UpdateLightMapWithColor li105,Playfield_LM_Inserts_li105, 100.0, " ' VLM.Lampz;Inserts-li105
  Lampz.Callback(32) = "UpdateLightMapWithColor li105,Layer3_LM_Inserts_li105, 100.0, " ' VLM.Lampz;Inserts-li105
  Lampz.Callback(32) = "UpdateLightMapWithColor li105,targets_005_LM_Inserts_li105, 100.0, " ' VLM.Lampz;Inserts-li105
  ' Sync on li108 ' VLM.Lampz;Inserts-li108
  Lampz.Callback(33) = "UpdateLightMapWithColor li108,Playfield_LM_Inserts_li108, 100.0, " ' VLM.Lampz;Inserts-li108
  Lampz.Callback(33) = "UpdateLightMapWithColor li108,Layer3_LM_Inserts_li108, 100.0, " ' VLM.Lampz;Inserts-li108
  ' Sync on li111 ' VLM.Lampz;Inserts-li111
  Lampz.Callback(34) = "UpdateLightMapWithColor li111,Playfield_LM_Inserts_li111, 100.0, " ' VLM.Lampz;Inserts-li111
  Lampz.Callback(34) = "UpdateLightMapWithColor li111,Layer3_LM_Inserts_li111, 100.0, " ' VLM.Lampz;Inserts-li111
  ' Sync on li114 ' VLM.Lampz;Inserts-li114
  Lampz.Callback(35) = "UpdateLightMapWithColor li114,Playfield_LM_Inserts_li114, 100.0, " ' VLM.Lampz;Inserts-li114
  Lampz.Callback(35) = "UpdateLightMapWithColor li114,Layer3_LM_Inserts_li114, 100.0, " ' VLM.Lampz;Inserts-li114
  ' Sync on li117 ' VLM.Lampz;Inserts-li117
  Lampz.Callback(36) = "UpdateLightMapWithColor li117,Playfield_LM_Inserts_li117, 100.0, " ' VLM.Lampz;Inserts-li117
  Lampz.Callback(36) = "UpdateLightMapWithColor li117,Layer1_LM_Inserts_li117, 100.0, " ' VLM.Lampz;Inserts-li117
  Lampz.Callback(36) = "UpdateLightMapWithColor li117,Layer3_LM_Inserts_li117, 100.0, " ' VLM.Lampz;Inserts-li117
  Lampz.Callback(36) = "UpdateLightMapWithColor li117,targets_005_LM_Inserts_li117, 100.0, " ' VLM.Lampz;Inserts-li117
  Lampz.Callback(36) = "UpdateLightMapWithColor li117,targets_006_LM_Inserts_li117, 100.0, " ' VLM.Lampz;Inserts-li117
  ' Sync on li11 ' VLM.Lampz;Inserts-li11
  Lampz.Callback(2) = "UpdateLightMapWithColor li11,Playfield_LM_Inserts_li11, 100.0, " ' VLM.Lampz;Inserts-li11
  Lampz.Callback(2) = "UpdateLightMapWithColor li11,outlane_left_LM_Inserts_li11, 100.0, " ' VLM.Lampz;Inserts-li11
  ' Sync on li120 ' VLM.Lampz;Inserts-li120
  Lampz.Callback(37) = "UpdateLightMapWithColor li120,Playfield_LM_Inserts_li120, 100.0, " ' VLM.Lampz;Inserts-li120
  Lampz.Callback(37) = "UpdateLightMapWithColor li120,ram_LM_Inserts_li120, 100.0, " ' VLM.Lampz;Inserts-li120
  Lampz.Callback(37) = "UpdateLightMapWithColor li120,targets_006_LM_Inserts_li120, 100.0, " ' VLM.Lampz;Inserts-li120
  ' Sync on li123 ' VLM.Lampz;Inserts-li123
  Lampz.Callback(38) = "UpdateLightMapWithColor li123,Playfield_LM_Inserts_li123, 100.0, " ' VLM.Lampz;Inserts-li123
  Lampz.Callback(38) = "UpdateLightMapWithColor li123,T90_LM_Inserts_li123, 100.0, " ' VLM.Lampz;Inserts-li123
  ' Sync on li126 ' VLM.Lampz;Inserts-li126
  Lampz.Callback(39) = "UpdateLightMapWithColor li126,Playfield_LM_Inserts_li126, 100.0, " ' VLM.Lampz;Inserts-li126
  Lampz.Callback(39) = "UpdateLightMapWithColor li126,Spinners_LM_Inserts_li126, 100.0, " ' VLM.Lampz;Inserts-li126
  Lampz.Callback(39) = "UpdateLightMapWithColor li126,T90_LM_Inserts_li126, 100.0, " ' VLM.Lampz;Inserts-li126
  ' Sync on li129 ' VLM.Lampz;Inserts-li129
  Lampz.Callback(40) = "UpdateLightMapWithColor li129,Playfield_LM_Inserts_li129, 100.0, " ' VLM.Lampz;Inserts-li129
  Lampz.Callback(40) = "UpdateLightMapWithColor li129,Spinners_LM_Inserts_li129, 100.0, " ' VLM.Lampz;Inserts-li129
  Lampz.Callback(40) = "UpdateLightMapWithColor li129,T90_LM_Inserts_li129, 100.0, " ' VLM.Lampz;Inserts-li129
  ' Sync on li132 ' VLM.Lampz;Inserts-li132
  Lampz.Callback(41) = "UpdateLightMapWithColor li132,Playfield_LM_Inserts_li132, 100.0, " ' VLM.Lampz;Inserts-li132
  Lampz.Callback(41) = "UpdateLightMapWithColor li132,Spinners_LM_Inserts_li132, 100.0, " ' VLM.Lampz;Inserts-li132
  Lampz.Callback(41) = "UpdateLightMapWithColor li132,T90_LM_Inserts_li132, 100.0, " ' VLM.Lampz;Inserts-li132
  Lampz.Callback(41) = "UpdateLightMapWithColor li132,ram_LM_Inserts_li132, 100.0, " ' VLM.Lampz;Inserts-li132
  Lampz.Callback(41) = "UpdateLightMapWithColor li132,targets_006_LM_Inserts_li132, 100.0, " ' VLM.Lampz;Inserts-li132
  ' Sync on li135 ' VLM.Lampz;Inserts-li135
  Lampz.Callback(42) = "UpdateLightMapWithColor li135,Playfield_LM_Inserts_li135, 100.0, " ' VLM.Lampz;Inserts-li135
  Lampz.Callback(42) = "UpdateLightMapWithColor li135,T90_LM_Inserts_li135, 100.0, " ' VLM.Lampz;Inserts-li135
  Lampz.Callback(42) = "UpdateLightMapWithColor li135,ram_LM_Inserts_li135, 100.0, " ' VLM.Lampz;Inserts-li135
  Lampz.Callback(42) = "UpdateLightMapWithColor li135,targets_007_LM_Inserts_li135, 100.0, " ' VLM.Lampz;Inserts-li135
  ' Sync on li138 ' VLM.Lampz;Inserts-li138
  Lampz.Callback(43) = "UpdateLightMapWithColor li138,Playfield_LM_Inserts_li138, 100.0, " ' VLM.Lampz;Inserts-li138
  ' Sync on li141 ' VLM.Lampz;Inserts-li141
  Lampz.Callback(44) = "UpdateLightMapWithColor li141,Playfield_LM_Inserts_li141, 100.0, " ' VLM.Lampz;Inserts-li141
  Lampz.Callback(44) = "UpdateLightMapWithColor li141,Spinners_LM_Inserts_li141, 100.0, " ' VLM.Lampz;Inserts-li141
  Lampz.Callback(44) = "UpdateLightMapWithColor li141,T90_LM_Inserts_li141, 100.0, " ' VLM.Lampz;Inserts-li141
  ' Sync on li144 ' VLM.Lampz;Inserts-li144
  Lampz.Callback(45) = "UpdateLightMapWithColor li144,Playfield_LM_Inserts_li144, 100.0, " ' VLM.Lampz;Inserts-li144
  Lampz.Callback(45) = "UpdateLightMapWithColor li144,T90_LM_Inserts_li144, 100.0, " ' VLM.Lampz;Inserts-li144
  ' Sync on li147 ' VLM.Lampz;Inserts-li147
  Lampz.Callback(46) = "UpdateLightMapWithColor li147,Playfield_LM_Inserts_li147, 100.0, " ' VLM.Lampz;Inserts-li147
  Lampz.Callback(46) = "UpdateLightMapWithColor li147,targets_008_LM_Inserts_li147, 100.0, " ' VLM.Lampz;Inserts-li147
  ' Sync on li14 ' VLM.Lampz;Inserts-li14
  Lampz.Callback(3) = "UpdateLightMapWithColor li14,Playfield_LM_Inserts_li14, 100.0, " ' VLM.Lampz;Inserts-li14
  Lampz.Callback(3) = "UpdateLightMapWithColor li14,T9_LM_Inserts_li14, 100.0, " ' VLM.Lampz;Inserts-li14
  ' Sync on li150 ' VLM.Lampz;Inserts-li150
  Lampz.Callback(47) = "UpdateLightMapWithColor li150,Playfield_LM_Inserts_li150, 100.0, " ' VLM.Lampz;Inserts-li150
  ' Sync on li153 ' VLM.Lampz;Inserts-li153
  Lampz.Callback(48) = "UpdateLightMapWithColor li153,Playfield_LM_Inserts_li153, 100.0, " ' VLM.Lampz;Inserts-li153
  Lampz.Callback(48) = "UpdateLightMapWithColor li153,T90_LM_Inserts_li153, 100.0, " ' VLM.Lampz;Inserts-li153
  Lampz.Callback(48) = "UpdateLightMapWithColor li153,targets_009_LM_Inserts_li153, 100.0, " ' VLM.Lampz;Inserts-li153
  ' Sync on li156 ' VLM.Lampz;Inserts-li156
  Lampz.Callback(49) = "UpdateLightMapWithColor li156,Playfield_LM_Inserts_li156, 100.0, " ' VLM.Lampz;Inserts-li156
  Lampz.Callback(49) = "UpdateLightMapWithColor li156,T90_LM_Inserts_li156, 100.0, " ' VLM.Lampz;Inserts-li156
  ' Sync on li159 ' VLM.Lampz;Inserts-li159
  Lampz.Callback(50) = "UpdateLightMapWithColor li159,Playfield_LM_Inserts_li159, 100.0, " ' VLM.Lampz;Inserts-li159
  Lampz.Callback(50) = "UpdateLightMapWithColor li159,T90_LM_Inserts_li159, 100.0, " ' VLM.Lampz;Inserts-li159
  Lampz.Callback(50) = "UpdateLightMapWithColor li159,ram_LM_Inserts_li159, 100.0, " ' VLM.Lampz;Inserts-li159
  ' Sync on li162 ' VLM.Lampz;Inserts-li162
  Lampz.Callback(51) = "UpdateLightMapWithColor li162,Playfield_LM_Inserts_li162, 100.0, " ' VLM.Lampz;Inserts-li162
  Lampz.Callback(51) = "UpdateLightMapWithColor li162,Layer4_LM_Inserts_li162, 100.0, " ' VLM.Lampz;Inserts-li162
  Lampz.Callback(51) = "UpdateLightMapWithColor li162,Layer5_LM_Inserts_li162, 100.0, " ' VLM.Lampz;Inserts-li162
  ' Sync on li165 ' VLM.Lampz;Inserts-li165
  Lampz.Callback(52) = "UpdateLightMapWithColor li165,Playfield_LM_Inserts_li165, 100.0, " ' VLM.Lampz;Inserts-li165
  Lampz.Callback(52) = "UpdateLightMapWithColor li165,Layer1_LM_Inserts_li165, 100.0, " ' VLM.Lampz;Inserts-li165
  ' Sync on li168 ' VLM.Lampz;Inserts-li168
  Lampz.Callback(66) = "UpdateLightMapWithColor li168,Playfield_LM_Inserts_li168, 100.0, " ' VLM.Lampz;Inserts-li168
  Lampz.Callback(66) = "UpdateLightMapWithColor li168,Layer1_LM_Inserts_li168, 100.0, " ' VLM.Lampz;Inserts-li168
  Lampz.Callback(66) = "UpdateLightMapWithColor li168,Layer3_LM_Inserts_li168, 100.0, " ' VLM.Lampz;Inserts-li168
  Lampz.Callback(66) = "UpdateLightMapWithColor li168,Layer4_LM_Inserts_li168, 100.0, " ' VLM.Lampz;Inserts-li168
  Lampz.Callback(66) = "UpdateLightMapWithColor li168,Layer5_LM_Inserts_li168, 100.0, " ' VLM.Lampz;Inserts-li168
  Lampz.Callback(66) = "UpdateLightMapWithColor li168,Bumper_Ring_LM_Inserts_li168, 100.0, " ' VLM.Lampz;Inserts-li168
  Lampz.Callback(66) = "UpdateLightMapWithColor li168,Bumper_Ring_001_LM_Inserts_li16, 100.0, " ' VLM.Lampz;Inserts-li168
  Lampz.Callback(66) = "UpdateLightMapWithColor li168,Bumper_Ring_002_LM_Inserts_li16, 100.0, " ' VLM.Lampz;Inserts-li168
  Lampz.Callback(66) = "UpdateLightMapWithColor li168,sw53_LM_Inserts_li168, 100.0, " ' VLM.Lampz;Inserts-li168
  ' Sync on li171 ' VLM.Lampz;Inserts-li171
  Lampz.Callback(67) = "UpdateLightMapWithColor li171,Playfield_LM_Inserts_li171, 100.0, " ' VLM.Lampz;Inserts-li171
  Lampz.Callback(67) = "UpdateLightMapWithColor li171,Layer1_LM_Inserts_li171, 100.0, " ' VLM.Lampz;Inserts-li171
  Lampz.Callback(67) = "UpdateLightMapWithColor li171,Layer3_LM_Inserts_li171, 100.0, " ' VLM.Lampz;Inserts-li171
  Lampz.Callback(67) = "UpdateLightMapWithColor li171,Bumper_Ring_LM_Inserts_li171, 100.0, " ' VLM.Lampz;Inserts-li171
  Lampz.Callback(67) = "UpdateLightMapWithColor li171,Bumper_Ring_001_LM_Inserts_li17, 100.0, " ' VLM.Lampz;Inserts-li171
  Lampz.Callback(67) = "UpdateLightMapWithColor li171,Bumper_Ring_002_LM_Inserts_li17, 100.0, " ' VLM.Lampz;Inserts-li171
  ' Sync on li174 ' VLM.Lampz;Inserts-li174
  Lampz.Callback(68) = "UpdateLightMapWithColor li174,Playfield_LM_Inserts_li174, 100.0, " ' VLM.Lampz;Inserts-li174
  Lampz.Callback(68) = "UpdateLightMapWithColor li174,Layer1_LM_Inserts_li174, 100.0, " ' VLM.Lampz;Inserts-li174
  Lampz.Callback(68) = "UpdateLightMapWithColor li174,Layer4_LM_Inserts_li174, 100.0, " ' VLM.Lampz;Inserts-li174
  Lampz.Callback(68) = "UpdateLightMapWithColor li174,Bumper_Ring_LM_Inserts_li174, 100.0, " ' VLM.Lampz;Inserts-li174
  Lampz.Callback(68) = "UpdateLightMapWithColor li174,Bumper_Ring_001_LM_Inserts_li17, 100.0, " ' VLM.Lampz;Inserts-li174
  Lampz.Callback(68) = "UpdateLightMapWithColor li174,Bumper_Ring_002_LM_Inserts_li17, 100.0, " ' VLM.Lampz;Inserts-li174
  ' Sync on li17 ' VLM.Lampz;Inserts-li17
  Lampz.Callback(4) = "UpdateLightMapWithColor li17,Playfield_LM_Inserts_li17, 100.0, " ' VLM.Lampz;Inserts-li17
  Lampz.Callback(4) = "UpdateLightMapWithColor li17,T7_LM_Inserts_li17, 100.0, " ' VLM.Lampz;Inserts-li17
  Lampz.Callback(4) = "UpdateLightMapWithColor li17,T8_LM_Inserts_li17, 100.0, " ' VLM.Lampz;Inserts-li17
  Lampz.Callback(4) = "UpdateLightMapWithColor li17,T9_LM_Inserts_li17, 100.0, " ' VLM.Lampz;Inserts-li17
  ' Sync on li180 ' VLM.Lampz;Inserts-li180
  Lampz.Callback(53) = "UpdateLightMapWithColor li180,Playfield_LM_Inserts_li180, 100.0, " ' VLM.Lampz;Inserts-li180
  Lampz.Callback(53) = "UpdateLightMapWithColor li180,Layer1_LM_Inserts_li180, 100.0, " ' VLM.Lampz;Inserts-li180
  Lampz.Callback(53) = "UpdateLightMapWithColor li180,Layer3_LM_Inserts_li180, 100.0, " ' VLM.Lampz;Inserts-li180
  Lampz.Callback(53) = "UpdateLightMapWithColor li180,targets_012_LM_Inserts_li180, 100.0, " ' VLM.Lampz;Inserts-li180
  ' Sync on li183 ' VLM.Lampz;Inserts-li183
  Lampz.Callback(54) = "UpdateLightMapWithColor li183,Playfield_LM_Inserts_li183, 100.0, " ' VLM.Lampz;Inserts-li183
  Lampz.Callback(54) = "UpdateLightMapWithColor li183,Layer1_LM_Inserts_li183, 100.0, " ' VLM.Lampz;Inserts-li183
  Lampz.Callback(54) = "UpdateLightMapWithColor li183,Layer3_LM_Inserts_li183, 100.0, " ' VLM.Lampz;Inserts-li183
  Lampz.Callback(54) = "UpdateLightMapWithColor li183,targets_011_LM_Inserts_li183, 100.0, " ' VLM.Lampz;Inserts-li183
  Lampz.Callback(54) = "UpdateLightMapWithColor li183,targets_012_LM_Inserts_li183, 100.0, " ' VLM.Lampz;Inserts-li183
  ' Sync on li186 ' VLM.Lampz;Inserts-li186
  Lampz.Callback(55) = "UpdateLightMapWithColor li186,Playfield_LM_Inserts_li186, 100.0, " ' VLM.Lampz;Inserts-li186
  Lampz.Callback(55) = "UpdateLightMapWithColor li186,Layer1_LM_Inserts_li186, 100.0, " ' VLM.Lampz;Inserts-li186
  ' Sync on li189 ' VLM.Lampz;Inserts-li189
  Lampz.Callback(56) = "UpdateLightMapWithColor li189,Playfield_LM_Inserts_li189, 100.0, " ' VLM.Lampz;Inserts-li189
  Lampz.Callback(56) = "UpdateLightMapWithColor li189,Layer3_LM_Inserts_li189, 100.0, " ' VLM.Lampz;Inserts-li189
  Lampz.Callback(56) = "UpdateLightMapWithColor li189,targets_LM_Inserts_li189, 100.0, " ' VLM.Lampz;Inserts-li189
  Lampz.Callback(56) = "UpdateLightMapWithColor li189,targets_012_LM_Inserts_li189, 100.0, " ' VLM.Lampz;Inserts-li189
  ' Sync on li192 ' VLM.Lampz;Inserts-li192
  Lampz.Callback(57) = "UpdateLightMapWithColor li192,Playfield_LM_Inserts_li192, 100.0, " ' VLM.Lampz;Inserts-li192
  Lampz.Callback(57) = "UpdateLightMapWithColor li192,Layer3_LM_Inserts_li192, 100.0, " ' VLM.Lampz;Inserts-li192
  Lampz.Callback(57) = "UpdateLightMapWithColor li192,targets_011_LM_Inserts_li192, 100.0, " ' VLM.Lampz;Inserts-li192
  Lampz.Callback(57) = "UpdateLightMapWithColor li192,targets_012_LM_Inserts_li192, 100.0, " ' VLM.Lampz;Inserts-li192
  ' Sync on li195 ' VLM.Lampz;Inserts-li195
  Lampz.Callback(58) = "UpdateLightMapWithColor li195,Playfield_LM_Inserts_li195, 100.0, " ' VLM.Lampz;Inserts-li195
  Lampz.Callback(58) = "UpdateLightMapWithColor li195,Layer1_LM_Inserts_li195, 100.0, " ' VLM.Lampz;Inserts-li195
  Lampz.Callback(58) = "UpdateLightMapWithColor li195,targets_011_LM_Inserts_li195, 100.0, " ' VLM.Lampz;Inserts-li195
  Lampz.Callback(58) = "UpdateLightMapWithColor li195,targets_012_LM_Inserts_li195, 100.0, " ' VLM.Lampz;Inserts-li195
  ' Sync on li198 ' VLM.Lampz;Inserts-li198
  Lampz.Callback(59) = "UpdateLightMapWithColor li198,Playfield_LM_Inserts_li198, 100.0, " ' VLM.Lampz;Inserts-li198
  Lampz.Callback(59) = "UpdateLightMapWithColor li198,Layer1_LM_Inserts_li198, 100.0, " ' VLM.Lampz;Inserts-li198
  Lampz.Callback(59) = "UpdateLightMapWithColor li198,Layer3_LM_Inserts_li198, 100.0, " ' VLM.Lampz;Inserts-li198
  Lampz.Callback(59) = "UpdateLightMapWithColor li198,LFlogo_001_LM_Inserts_li198, 100.0, " ' VLM.Lampz;Inserts-li198
  Lampz.Callback(59) = "UpdateLightMapWithColor li198,RFlogo_001_LM_Inserts_li198, 100.0, " ' VLM.Lampz;Inserts-li198
  ' Sync on li201 ' VLM.Lampz;Inserts-li201
  Lampz.Callback(60) = "UpdateLightMapWithColor li201,Playfield_LM_Inserts_li201, 100.0, " ' VLM.Lampz;Inserts-li201
  Lampz.Callback(60) = "UpdateLightMapWithColor li201,Layer3_LM_Inserts_li201, 100.0, " ' VLM.Lampz;Inserts-li201
  Lampz.Callback(60) = "UpdateLightMapWithColor li201,LFlogo_001_LM_Inserts_li201, 100.0, " ' VLM.Lampz;Inserts-li201
  Lampz.Callback(60) = "UpdateLightMapWithColor li201,RFlogo_001_LM_Inserts_li201, 100.0, " ' VLM.Lampz;Inserts-li201
  ' Sync on li204 ' VLM.Lampz;Inserts-li204
  Lampz.Callback(61) = "UpdateLightMapWithColor li204,Playfield_LM_Inserts_li204, 100.0, " ' VLM.Lampz;Inserts-li204
  Lampz.Callback(61) = "UpdateLightMapWithColor li204,Layer3_LM_Inserts_li204, 100.0, " ' VLM.Lampz;Inserts-li204
  Lampz.Callback(61) = "UpdateLightMapWithColor li204,LFlogo_001_LM_Inserts_li204, 100.0, " ' VLM.Lampz;Inserts-li204
  Lampz.Callback(61) = "UpdateLightMapWithColor li204,RFlogo_001_LM_Inserts_li204, 100.0, " ' VLM.Lampz;Inserts-li204
  ' Sync on li207 ' VLM.Lampz;Inserts-li207
  Lampz.Callback(62) = "UpdateLightMapWithColor li207,Playfield_LM_Inserts_li207, 100.0, " ' VLM.Lampz;Inserts-li207
  Lampz.Callback(62) = "UpdateLightMapWithColor li207,Layer3_LM_Inserts_li207, 100.0, " ' VLM.Lampz;Inserts-li207
  Lampz.Callback(62) = "UpdateLightMapWithColor li207,LFlogo_001_LM_Inserts_li207, 100.0, " ' VLM.Lampz;Inserts-li207
  Lampz.Callback(62) = "UpdateLightMapWithColor li207,RFlogo_001_LM_Inserts_li207, 100.0, " ' VLM.Lampz;Inserts-li207
  ' Sync on li20 ' VLM.Lampz;Inserts-li20
  Lampz.Callback(5) = "UpdateLightMapWithColor li20,Playfield_LM_Inserts_li20, 100.0, " ' VLM.Lampz;Inserts-li20
  Lampz.Callback(5) = "UpdateLightMapWithColor li20,T7_LM_Inserts_li20, 100.0, " ' VLM.Lampz;Inserts-li20
  Lampz.Callback(5) = "UpdateLightMapWithColor li20,T8_LM_Inserts_li20, 100.0, " ' VLM.Lampz;Inserts-li20
  Lampz.Callback(5) = "UpdateLightMapWithColor li20,T9_LM_Inserts_li20, 100.0, " ' VLM.Lampz;Inserts-li20
  ' Sync on li210 ' VLM.Lampz;Inserts-li210
  Lampz.Callback(63) = "UpdateLightMapWithColor li210,Playfield_LM_Inserts_li210, 100.0, " ' VLM.Lampz;Inserts-li210
  Lampz.Callback(63) = "UpdateLightMapWithColor li210,Layer3_LM_Inserts_li210, 100.0, " ' VLM.Lampz;Inserts-li210
  Lampz.Callback(63) = "UpdateLightMapWithColor li210,LFlogo_001_LM_Inserts_li210, 100.0, " ' VLM.Lampz;Inserts-li210
  Lampz.Callback(63) = "UpdateLightMapWithColor li210,RFlogo_001_LM_Inserts_li210, 100.0, " ' VLM.Lampz;Inserts-li210
  ' Sync on li213 ' VLM.Lampz;Inserts-li213
  Lampz.Callback(64) = "UpdateLightMapWithColor li213,Playfield_LM_Inserts_li213, 100.0, " ' VLM.Lampz;Inserts-li213
  Lampz.Callback(64) = "UpdateLightMapWithColor li213,Layer3_LM_Inserts_li213, 100.0, " ' VLM.Lampz;Inserts-li213
  Lampz.Callback(64) = "UpdateLightMapWithColor li213,LFlogo_001_LM_Inserts_li213, 100.0, " ' VLM.Lampz;Inserts-li213
  Lampz.Callback(64) = "UpdateLightMapWithColor li213,RFlogo_001_LM_Inserts_li213, 100.0, " ' VLM.Lampz;Inserts-li213
  ' Sync on li216 ' VLM.Lampz;Inserts-li216
  Lampz.Callback(65) = "UpdateLightMapWithColor li216,Playfield_LM_Inserts_li216, 100.0, " ' VLM.Lampz;Inserts-li216
  Lampz.Callback(65) = "UpdateLightMapWithColor li216,Layer3_LM_Inserts_li216, 100.0, " ' VLM.Lampz;Inserts-li216
  Lampz.Callback(65) = "UpdateLightMapWithColor li216,LFlogo_001_LM_Inserts_li216, 100.0, " ' VLM.Lampz;Inserts-li216
  Lampz.Callback(65) = "UpdateLightMapWithColor li216,RFlogo_001_LM_Inserts_li216, 100.0, " ' VLM.Lampz;Inserts-li216
  ' Sync on li222 ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,Playfield_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,Layer1_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,Layer4_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,LFlogo_001_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,Lwing_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,RFlogo_001_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,RFlogo_003_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,Rwing_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,metalPlates_002_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,sw77_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,sw78_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,sw79_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,sw83_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,sw84_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,targets_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,targets_011_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  Lampz.Callback(77) = "UpdateLightMapWithColor li222,targets_012_LM_Upper_GI, 100.0, " ' VLM.Lampz;Upper GI
  ' Sync on li23 ' VLM.Lampz;Inserts-li23
  Lampz.Callback(6) = "UpdateLightMapWithColor li23,Playfield_LM_Inserts_li23, 100.0, " ' VLM.Lampz;Inserts-li23
  Lampz.Callback(6) = "UpdateLightMapWithColor li23,Spinners_LM_Inserts_li23, 100.0, " ' VLM.Lampz;Inserts-li23
  Lampz.Callback(6) = "UpdateLightMapWithColor li23,T7_LM_Inserts_li23, 100.0, " ' VLM.Lampz;Inserts-li23
  Lampz.Callback(6) = "UpdateLightMapWithColor li23,T8_LM_Inserts_li23, 100.0, " ' VLM.Lampz;Inserts-li23
  Lampz.Callback(6) = "UpdateLightMapWithColor li23,T9_LM_Inserts_li23, 100.0, " ' VLM.Lampz;Inserts-li23
  ' Sync on li26 ' VLM.Lampz;Inserts-li26
  Lampz.Callback(7) = "UpdateLightMapWithColor li26,Playfield_LM_Inserts_li26, 100.0, " ' VLM.Lampz;Inserts-li26
  Lampz.Callback(7) = "UpdateLightMapWithColor li26,T7_LM_Inserts_li26, 100.0, " ' VLM.Lampz;Inserts-li26
  Lampz.Callback(7) = "UpdateLightMapWithColor li26,T8_LM_Inserts_li26, 100.0, " ' VLM.Lampz;Inserts-li26
  Lampz.Callback(7) = "UpdateLightMapWithColor li26,T9_LM_Inserts_li26, 100.0, " ' VLM.Lampz;Inserts-li26
  ' Sync on li29 ' VLM.Lampz;Inserts-li29
  Lampz.Callback(8) = "UpdateLightMapWithColor li29,Playfield_LM_Inserts_li29, 100.0, " ' VLM.Lampz;Inserts-li29
  Lampz.Callback(8) = "UpdateLightMapWithColor li29,T9_LM_Inserts_li29, 100.0, " ' VLM.Lampz;Inserts-li29
  ' Sync on li32 ' VLM.Lampz;Inserts-li32
  Lampz.Callback(9) = "UpdateLightMapWithColor li32,Playfield_LM_Inserts_li32, 100.0, " ' VLM.Lampz;Inserts-li32
  ' Sync on li35 ' VLM.Lampz;Inserts-li35
  Lampz.Callback(10) = "UpdateLightMapWithColor li35,Playfield_LM_Inserts_li35, 100.0, " ' VLM.Lampz;Inserts-li35
  ' Sync on li38 ' VLM.Lampz;Inserts-li38
  Lampz.Callback(11) = "UpdateLightMapWithColor li38,Playfield_LM_Inserts_li38, 100.0, " ' VLM.Lampz;Inserts-li38
  Lampz.Callback(11) = "UpdateLightMapWithColor li38,LFlogo_003_LM_Inserts_li38, 100.0, " ' VLM.Lampz;Inserts-li38
  Lampz.Callback(11) = "UpdateLightMapWithColor li38,RFlogo_003_LM_Inserts_li38, 100.0, " ' VLM.Lampz;Inserts-li38
  ' Sync on li41 ' VLM.Lampz;Inserts-li41
  Lampz.Callback(12) = "UpdateLightMapWithColor li41,Playfield_LM_Inserts_li41, 100.0, " ' VLM.Lampz;Inserts-li41
  Lampz.Callback(12) = "UpdateLightMapWithColor li41,LFlogo_003_LM_Inserts_li41, 100.0, " ' VLM.Lampz;Inserts-li41
  Lampz.Callback(12) = "UpdateLightMapWithColor li41,RFlogo_003_LM_Inserts_li41, 100.0, " ' VLM.Lampz;Inserts-li41
  ' Sync on li44 ' VLM.Lampz;Inserts-li44
  Lampz.Callback(13) = "UpdateLightMapWithColor li44,Playfield_LM_Inserts_li44, 100.0, " ' VLM.Lampz;Inserts-li44
  Lampz.Callback(13) = "UpdateLightMapWithColor li44,LFlogo_003_LM_Inserts_li44, 100.0, " ' VLM.Lampz;Inserts-li44
  Lampz.Callback(13) = "UpdateLightMapWithColor li44,RFlogo_003_LM_Inserts_li44, 100.0, " ' VLM.Lampz;Inserts-li44
  ' Sync on li47 ' VLM.Lampz;Inserts-li47
  Lampz.Callback(14) = "UpdateLightMapWithColor li47,Playfield_LM_Inserts_li47, 100.0, " ' VLM.Lampz;Inserts-li47
  Lampz.Callback(14) = "UpdateLightMapWithColor li47,LFlogo_003_LM_Inserts_li47, 100.0, " ' VLM.Lampz;Inserts-li47
  Lampz.Callback(14) = "UpdateLightMapWithColor li47,RFlogo_003_LM_Inserts_li47, 100.0, " ' VLM.Lampz;Inserts-li47
  ' Sync on li50 ' VLM.Lampz;Inserts-li50
  Lampz.Callback(15) = "UpdateLightMapWithColor li50,Playfield_LM_Inserts_li50, 100.0, " ' VLM.Lampz;Inserts-li50
  Lampz.Callback(15) = "UpdateLightMapWithColor li50,LFlogo_003_LM_Inserts_li50, 100.0, " ' VLM.Lampz;Inserts-li50
  Lampz.Callback(15) = "UpdateLightMapWithColor li50,RFlogo_003_LM_Inserts_li50, 100.0, " ' VLM.Lampz;Inserts-li50
  ' Sync on li53 ' VLM.Lampz;Inserts-li53
  Lampz.Callback(16) = "UpdateLightMapWithColor li53,Playfield_LM_Inserts_li53, 100.0, " ' VLM.Lampz;Inserts-li53
  Lampz.Callback(16) = "UpdateLightMapWithColor li53,LFlogo_003_LM_Inserts_li53, 100.0, " ' VLM.Lampz;Inserts-li53
  Lampz.Callback(16) = "UpdateLightMapWithColor li53,RFlogo_003_LM_Inserts_li53, 100.0, " ' VLM.Lampz;Inserts-li53
  ' Sync on li56 ' VLM.Lampz;Inserts-li56
  Lampz.Callback(17) = "UpdateLightMapWithColor li56,Playfield_LM_Inserts_li56, 100.0, " ' VLM.Lampz;Inserts-li56
  Lampz.Callback(17) = "UpdateLightMapWithColor li56,LFlogo_003_LM_Inserts_li56, 100.0, " ' VLM.Lampz;Inserts-li56
  Lampz.Callback(17) = "UpdateLightMapWithColor li56,RFlogo_003_LM_Inserts_li56, 100.0, " ' VLM.Lampz;Inserts-li56
  ' Sync on li59 ' VLM.Lampz;Inserts-li59
  Lampz.Callback(18) = "UpdateLightMapWithColor li59,Playfield_LM_Inserts_li59, 100.0, " ' VLM.Lampz;Inserts-li59
  Lampz.Callback(18) = "UpdateLightMapWithColor li59,LFlogo_003_LM_Inserts_li59, 100.0, " ' VLM.Lampz;Inserts-li59
  Lampz.Callback(18) = "UpdateLightMapWithColor li59,RFlogo_003_LM_Inserts_li59, 100.0, " ' VLM.Lampz;Inserts-li59
  Lampz.Callback(18) = "UpdateLightMapWithColor li59,FireButton_LM_Inserts_li59, 100.0, " ' VLM.Lampz;Inserts-li59
  ' Sync on li62 ' VLM.Lampz;Inserts-li62
  Lampz.Callback(19) = "UpdateLightMapWithColor li62,Playfield_LM_Inserts_li62, 100.0, " ' VLM.Lampz;Inserts-li62
  Lampz.Callback(19) = "UpdateLightMapWithColor li62,LFlogo_003_LM_Inserts_li62, 100.0, " ' VLM.Lampz;Inserts-li62
  Lampz.Callback(19) = "UpdateLightMapWithColor li62,RFlogo_003_LM_Inserts_li62, 100.0, " ' VLM.Lampz;Inserts-li62
  ' Sync on li65 ' VLM.Lampz;Inserts-li65
  Lampz.Callback(20) = "UpdateLightMapWithColor li65,Playfield_LM_Inserts_li65, 100.0, " ' VLM.Lampz;Inserts-li65
  Lampz.Callback(20) = "UpdateLightMapWithColor li65,LFlogo_003_LM_Inserts_li65, 100.0, " ' VLM.Lampz;Inserts-li65
  Lampz.Callback(20) = "UpdateLightMapWithColor li65,RFlogo_003_LM_Inserts_li65, 100.0, " ' VLM.Lampz;Inserts-li65
  Lampz.Callback(20) = "UpdateLightMapWithColor li65,FireButton_LM_Inserts_li65, 100.0, " ' VLM.Lampz;Inserts-li65
  ' Sync on li71 ' VLM.Lampz;Inserts-li71
  Lampz.Callback(21) = "UpdateLightMapWithColor li71,Playfield_LM_Inserts_li71, 100.0, " ' VLM.Lampz;Inserts-li71
  ' Sync on li74 ' VLM.Lampz;Inserts-li74
  Lampz.Callback(22) = "UpdateLightMapWithColor li74,Playfield_LM_Inserts_li74, 100.0, " ' VLM.Lampz;Inserts-li74
  ' Sync on li77 ' VLM.Lampz;Inserts-li77
  Lampz.Callback(23) = "UpdateLightMapWithColor li77,Playfield_LM_Inserts_li77, 100.0, " ' VLM.Lampz;Inserts-li77
  Lampz.Callback(23) = "UpdateLightMapWithColor li77,sword_003_LM_Inserts_li77, 100.0, " ' VLM.Lampz;Inserts-li77
  ' Sync on li80 ' VLM.Lampz;Inserts-li80
  Lampz.Callback(24) = "UpdateLightMapWithColor li80,Playfield_LM_Inserts_li80, 100.0, " ' VLM.Lampz;Inserts-li80
  Lampz.Callback(24) = "UpdateLightMapWithColor li80,lockpost_001_LM_Inserts_li80, 100.0, " ' VLM.Lampz;Inserts-li80
  Lampz.Callback(24) = "UpdateLightMapWithColor li80,sword_002_LM_Inserts_li80, 100.0, " ' VLM.Lampz;Inserts-li80
  Lampz.Callback(24) = "UpdateLightMapWithColor li80,sword_003_LM_Inserts_li80, 100.0, " ' VLM.Lampz;Inserts-li80
  Lampz.Callback(24) = "UpdateLightMapWithColor li80,targets_009_LM_Inserts_li80, 100.0, " ' VLM.Lampz;Inserts-li80
  Lampz.Callback(24) = "UpdateLightMapWithColor li80,targets_010_LM_Inserts_li80, 100.0, " ' VLM.Lampz;Inserts-li80
  ' Sync on li83 ' VLM.Lampz;Inserts-li83
  Lampz.Callback(25) = "UpdateLightMapWithColor li83,Playfield_LM_Inserts_li83, 100.0, " ' VLM.Lampz;Inserts-li83
  Lampz.Callback(25) = "UpdateLightMapWithColor li83,lockpost_001_LM_Inserts_li83, 100.0, " ' VLM.Lampz;Inserts-li83
  Lampz.Callback(25) = "UpdateLightMapWithColor li83,sword_003_LM_Inserts_li83, 100.0, " ' VLM.Lampz;Inserts-li83
  Lampz.Callback(25) = "UpdateLightMapWithColor li83,targets_010_LM_Inserts_li83, 100.0, " ' VLM.Lampz;Inserts-li83
  ' Sync on li86 ' VLM.Lampz;Inserts-li86
  Lampz.Callback(26) = "UpdateLightMapWithColor li86,Playfield_LM_Inserts_li86, 100.0, " ' VLM.Lampz;Inserts-li86
  Lampz.Callback(26) = "UpdateLightMapWithColor li86,Spinners_LM_Inserts_li86, 100.0, " ' VLM.Lampz;Inserts-li86
  Lampz.Callback(26) = "UpdateLightMapWithColor li86,T7_LM_Inserts_li86, 100.0, " ' VLM.Lampz;Inserts-li86
  Lampz.Callback(26) = "UpdateLightMapWithColor li86,T8_LM_Inserts_li86, 100.0, " ' VLM.Lampz;Inserts-li86
  Lampz.Callback(26) = "UpdateLightMapWithColor li86,targets_004_LM_Inserts_li86, 100.0, " ' VLM.Lampz;Inserts-li86
  ' Sync on li89 ' VLM.Lampz;Inserts-li89
  Lampz.Callback(27) = "UpdateLightMapWithColor li89,Playfield_LM_Inserts_li89, 100.0, " ' VLM.Lampz;Inserts-li89
  Lampz.Callback(27) = "UpdateLightMapWithColor li89,Spinners_LM_Inserts_li89, 100.0, " ' VLM.Lampz;Inserts-li89
  Lampz.Callback(27) = "UpdateLightMapWithColor li89,T7_LM_Inserts_li89, 100.0, " ' VLM.Lampz;Inserts-li89
  Lampz.Callback(27) = "UpdateLightMapWithColor li89,T8_LM_Inserts_li89, 100.0, " ' VLM.Lampz;Inserts-li89
  ' Sync on li92 ' VLM.Lampz;Inserts-li92
  Lampz.Callback(28) = "UpdateLightMapWithColor li92,Playfield_LM_Inserts_li92, 100.0, " ' VLM.Lampz;Inserts-li92
  Lampz.Callback(28) = "UpdateLightMapWithColor li92,Layer3_LM_Inserts_li92, 100.0, " ' VLM.Lampz;Inserts-li92
  Lampz.Callback(28) = "UpdateLightMapWithColor li92,Spinners_LM_Inserts_li92, 100.0, " ' VLM.Lampz;Inserts-li92
  Lampz.Callback(28) = "UpdateLightMapWithColor li92,T7_LM_Inserts_li92, 100.0, " ' VLM.Lampz;Inserts-li92
  Lampz.Callback(28) = "UpdateLightMapWithColor li92,targets_004_LM_Inserts_li92, 100.0, " ' VLM.Lampz;Inserts-li92
  Lampz.Callback(28) = "UpdateLightMapWithColor li92,targets_012_LM_Inserts_li92, 100.0, " ' VLM.Lampz;Inserts-li92
  ' Sync on li95 ' VLM.Lampz;Inserts-li95
  Lampz.Callback(29) = "UpdateLightMapWithColor li95,Playfield_LM_Inserts_li95, 100.0, " ' VLM.Lampz;Inserts-li95
  Lampz.Callback(29) = "UpdateLightMapWithColor li95,Layer3_LM_Inserts_li95, 100.0, " ' VLM.Lampz;Inserts-li95
  Lampz.Callback(29) = "UpdateLightMapWithColor li95,targets_004_LM_Inserts_li95, 100.0, " ' VLM.Lampz;Inserts-li95
  ' Sync on li98 ' VLM.Lampz;Inserts-li98
  Lampz.Callback(30) = "UpdateLightMapWithColor li98,Playfield_LM_Inserts_li98, 100.0, " ' VLM.Lampz;Inserts-li98
  Lampz.Callback(30) = "UpdateLightMapWithColor li98,Layer3_LM_Inserts_li98, 100.0, " ' VLM.Lampz;Inserts-li98
End Sub

' ===============================================================
' The following code can serve as a base for movable position synchronization.
' You will need to adapt the part of the transform you want to synchronize
' and the source on which you want it to be synchronized.

Sub MovableHelper
    Dim lfa,rfa,x1,x2,y,z,t,f
    lfa = LeftFlipper.CurrentAngle + 90
    rfa = RightFlipper.CurrentAngle - 30

    ' Upper PF flippers
    For each f in LFlogo_001_BL : f.RotZ = lfa : Next
  For each f in RFlogo_001_BL : f.RotZ = rfa : Next

    ' Main FP flippers
    For each f in LFlogo_003_BL : f.RotZ = lfa : Next
  For each f in RFlogo_003_BL : f.RotZ = rfa : Next

    ' Slings
    ' x1 = LeftSling4.visible
    ' x2 = LeftSling3.visible
  ' LSling1_BM_Dark_Room.Visible = x1 ' VLM.Props;BM;1;LSling1
  ' LSling1_LM_Lit_Room.Visible = x1 ' VLM.Props;LM;1;LSling1
  ' LSling1_LM_GI1.Visible = x1 ' VLM.Props;LM;1;LSling1
  ' LSling2_BM_Dark_Room.Visible = x2 ' VLM.Props;BM;1;LSling2
  ' LSling2_LM_GI1.Visible = x2 ' VLM.Props;LM;1;LSling2

    ' x1 = RightSling4.visible
    ' x2 = RightSling3.visible
  ' RSling1_BM_Dark_Room.Visible = x1 ' VLM.Props;BM;1;RSling1
  ' RSling1_LM_Lit_Room.Visible = x1 ' VLM.Props;LM;1;RSling1
  ' RSling1_LM_Flashers_f244B.Visible = x1 ' VLM.Props;LM;1;RSling1
  ' RSling1_LM_GI1.Visible = x1 ' VLM.Props;LM;1;RSling1
  ' RSling2_BM_Dark_Room.Visible = x2 ' VLM.Props;BM;1;RSling2
  ' RSling2_LM_Lit_Room.Visible = x2 ' VLM.Props;LM;1;RSling2
  ' RSling2_LM_GI1.Visible = x2 ' VLM.Props;LM;1;RSling2

    x1 = Lemk.RotX
    x2 = Remk.RotX
    For each f in LEMK_BL : f.RotY = x1 : Next
    For each f in REMK_BL : f.RotY = x2*-1 : Next

  'moving to ram timer in 1.1.25
    'Battering ram
  'y = ram_BM_Dark_Room.TransY
  'For each f in ram_BL : f.transy = y : Next

    ' Left Orbit Spinner
    Dim a:a=Spinner001.currentAngle
    For each f in Spinners_BL : f.RotX = a : Next

    ' Right Ramp top gate
    a = Gate001.currentAngle
    For each t in Gates_001_BL : t.RotZ=a : Next

End Sub

Sub TargetMovableHelper
  Dim t

    ' Castle Gate drop target
    For each t in T90_BL : t.transz = T90_BL(0).transz : t.rotx = T90_BL(0).rotx : t.roty = T90_BL(0).roty : Next

    ' LoL Drop Targets
    For each t in T7_BL : t.transz = T7_BL(0).transz : t.rotx = T7_BL(0).rotx : t.roty = T7_BL(0).roty : Next
    For each t in T8_BL : t.transz = T8_BL(0).transz : t.rotx = T8_BL(0).rotx : t.roty = T8_BL(0).roty : Next
    For each t in T9_BL : t.transz = T9_BL(0).transz : t.rotx = T9_BL(0).rotx : t.roty = T9_BL(0).roty : Next

    ' Stand Up Targets
    For each t in targets_004_BL : t.transz = targets_004_BL(0).transz : t.rotx = targets_004_BL(0).rotx : t.roty = targets_004_BL(0).roty : Next
    For each t in targets_005_BL : t.transz = targets_005_BL(0).transz : t.rotx = targets_005_BL(0).rotx : t.roty = targets_005_BL(0).roty : Next
    For each t in targets_006_BL : t.transz = targets_006_BL(0).transz : t.rotx = targets_006_BL(0).rotx : t.roty = targets_006_BL(0).roty : Next
    For each t in targets_007_BL : t.transz = targets_007_BL(0).transz : t.rotx = targets_007_BL(0).rotx : t.roty = targets_007_BL(0).roty : Next
    For each t in targets_008_BL : t.transz = targets_008_BL(0).transz : t.rotx = targets_008_BL(0).rotx : t.roty = targets_008_BL(0).roty : Next
    For each t in targets_009_BL : t.transz = targets_009_BL(0).transz : t.rotx = targets_009_BL(0).rotx : t.roty = targets_009_BL(0).roty : Next
    For each t in targets_010_BL : t.transz = targets_010_BL(0).transz : t.rotx = targets_010_BL(0).rotx : t.roty = targets_010_BL(0).roty : Next
    For each t in targets_011_BL : t.transz = targets_011_BL(0).transz : t.rotx = targets_011_BL(0).rotx : t.roty = targets_011_BL(0).roty : Next
    For each t in targets_012_BL : t.transz = targets_012_BL(0).transz : t.rotx = targets_012_BL(0).rotx : t.roty = targets_012_BL(0).roty : Next

End Sub

Sub UpdateWires(sw,pushed)
    Dim z : If pushed then z = -14 Else z = 0
    Dim o,obj
    Select Case sw
        Case 1 : obj = sw1_BL
        Case 2 : obj = sw2_BL
        Case 3 : obj = sw3_BL
        Case 4 : obj = sw4_BL
        Case 53 : obj = sw53_BL
        Case 54 : obj = sw54_BL
        Case 77 : obj = sw77_BL
        Case 78 : obj = sw78_BL
        Case 79 : obj = sw79_BL
        Case 83 : obj = sw83_BL
        Case 84 : obj = sw84_BL
    End Select
    For each o in obj : o.transx = z : Next
End Sub


' ===============================================================
' The following code can be copy/pasted to disable baked lights
' Lights are not removed on export since they are needed for ball
' reflections and may be used for lightmap synchronisation.

Sub HideLightHelper
  'L150.Visible = False
  LightShootAgain.Visible = False
  'f25.Visible = False
    'GI bottom
  gi002.Visible = False
  gi003.Visible = False
  gi004.Visible = False
  gi005.Visible = False
  gi010.Visible = False
  gi012.Visible = False
  gi016.Visible = False
    gi017.Visible = False
  gi018.Visible = False
  gi019.Visible = False
  gi020.Visible = False
  gi024.Visible = False
  gi026.Visible = False
  gi027.Visible = False
  gi028.Visible = False
  gi029.Visible = False
  gi030.Visible = False
  gi033.Visible = False
  gi035.Visible = False
  gi036.Visible = False
  gi037.Visible = False
  gi038.Visible = False
  gi045.Visible = False
  gi049.Visible = False
  gi050.Visible = False
  gi057.Visible = False
  gi058.Visible = False
  gi059.Visible = False
    ' GI top
    Dim g
    For each g in aGiLights : g.Visible = False : Next
  'l103.Visible = False
  li101.Visible = False
  'li104.Visible = False
  li105.Visible = False
  li108.Visible = False
  li11.Visible = False
  li111.Visible = False
  li114.Visible = False
  li117.Visible = False
  li120.Visible = False
  li123.Visible = False
  li126.Visible = False
  li129.Visible = False
  li132.Visible = False
  li135.Visible = False
  li138.Visible = False
  li14.Visible = False
  li141.Visible = False
  li144.Visible = False
  li147.Visible = False
  li150.Visible = False
  li153.Visible = False
  li156.Visible = False
  li159.Visible = False
  li162.Visible = False
  li165.Visible = False
  li168.Visible = False
  li17.Visible = False
  li171.Visible = False
  li174.Visible = False
  li20.Visible = False
  li23.Visible = False
  li26.Visible = False
  li29.Visible = False
  li32.Visible = False
  li35.Visible = False
  li38.Visible = False
  li41.Visible = False
  li44.Visible = False
  li47.Visible = False
  li50.Visible = False
  li53.Visible = False
  li56.Visible = False
  li59.Visible = False
  li62.Visible = False
  li65.Visible = False
  li71.Visible = False
  li74.Visible = False
  li77.Visible = False
  li80.Visible = False
  li83.Visible = False
  li86.Visible = False
  li89.Visible = False
  li92.Visible = False
  li95.Visible = False
  li98.Visible = False
    'UpperPF
    li180.Visible = False
    li183.Visible = False
    li186.Visible = False
    li189.Visible = False
    li192.Visible = False
    li195.Visible = False
    li198.Visible = False
    li201.Visible = False
    li204.Visible = False
    li207.Visible = False
    li210.Visible = False
    li213.Visible = False
    li216.Visible = False

  li222.Visible = False
    li230.Visible = False
    li231.Visible = False

    fl235.Visible = False
    fl236.Visible = False
    fl242.Visible = False
    fl243.Visible = False
    fl244.Visible = False
    fl245.Visible = False

End Sub

Function LightFade(light, is_on, percent)
  If is_on Then
    LightFade = percent*percent*(3 - 2*percent) ' Smoothstep
  Else
    LightFade = 1 - Sqr(1 - percent*percent) '
  End If
End Function

Sub UpdateLightMapFromFlasher(flasher, lightmap, intensity_scale, sync_color)
  If flasher.Visible Then
    If sync_color Then lightmap.Color = flasher.Color
    lightmap.Opacity = intensity_scale * flasher.IntensityScale * flasher.Opacity / 1000.0
  Else
    lightmap.Opacity = 0
  End If
End Sub

Sub UpdateLightMapFromLight(light, lightmap, intensity_scale, sync_color)
  light.FadeSpeedUp = light.Intensity / 50 '100
  light.FadeSpeedDown = light.Intensity / 200
  If sync_color Then lightmap.Color = light.Colorfull
  Dim t: t = LightFade(light, light.GetInPlayStateBool(), light.GetInPlayIntensity() / (light.Intensity * light.IntensityScale))
  lightmap.Opacity = intensity_scale * light.IntensityScale * t
End Sub


'******************* VR Stuff (Rawd) **********************

'VR Plunger...
Sub VRPlunger_Timer
  If VR_Plunger.Y < -17.5515 then
      VR_Plunger.Y = VR_Plunger.Y + 4
  End If
End Sub

Sub VRPlunger2_Timer
  VR_Plunger.Y = -117.5515 + (5* Plunger.Position) -20
End Sub

'VR Start/Tournament button Attract Lighting code...
VR_StartButtOuter.blenddisablelighting = 0.8
Sub StartButtonTimer_Timer  ' This is just here to test the start and tourney button lights for VR
    If VR_StartButtOuter.blenddisablelighting = 0.8 Then
        VR_StartButtOuter.blenddisablelighting = 0: VR_TourneyButt.blenddisablelighting = 0.8
    Else
        VR_StartButtOuter.blenddisablelighting = 0.8: VR_TourneyButt.blenddisablelighting = 0
    end if
End Sub

' Cab art swap stuff.....
Dim VRCabArt: VRCabArt = 1

Sub ArtSwap()
    If VRCabArt = 5 then VRCabArt = 1

    If VRCabArt = 1 then
        VRBackglassFlasher.imageA = "VRBG1"
        VRBackglassFlasher2.imageA = "VRBG1"
        VRBackglassFlasher2.imageB = "VRBG1"
        VRBackglassFlasher2.amount = 100
        VRBackglassFlasher2.opacity = 100
        VRBackglassFlasherFaded.imageA = "VRBG1faded"
        VRBackglassFlasherFaded.imageB = "VRBG1faded"
        VRBackglassFlasherFaded.amount = 100
        VRBackglassFlasherFaded.opacity = 100
    VRBackglassFlasherBulb.amount = 0
    VRBackglassFlasherBulb.opacity = 0
        PinCab_Cabinet.image = "Pincab_Cabinet_WinterNEW"
        PinCab_Backbox.image = "Pincab_Backbox_Winter"
        BackboxTrimLeft.visible = false
        BackboxTrimRight.visible = false
        BackboxTrimLeftsilver.visible = true
        BackboxTrimRightsilver.visible = true
        LELabel.visible = false
        VR_FlipRight.image = "FlipperBlue"
        VR_FlipperLeft.image = "FlipperBlue"
        PinCab_FlipperRing.image = "FlipperBlue"
    end If

    If VRCabArt = 2 then
        VRBackglassFlasher.imageA = "VRBG2"
        VRBackglassFlasher2.imageA = "VRBG2"
        VRBackglassFlasher2.amount = 100
        VRBackglassFlasher2.opacity = 100
        VRBackglassFlasher2.imageB = "VRBG2"
        VRBackglassFlasherFaded.imageA = "VRBG2faded"
        VRBackglassFlasherFaded.imageB = "VRBG2faded"
        VRBackglassFlasherFaded.amount = 100
        VRBackglassFlasherFaded.opacity = 100
    VRBackglassFlasherBulb.amount = 0
    VRBackglassFlasherBulb.opacity = 0
        PinCab_Cabinet.image = "Pincab_Cabinet_DragonNEW"
        PinCab_Backbox.image = "Pincab_Backbox_Dragon"
        BackboxTrimLeftsilver.visible = false
        BackboxTrimRightsilver.visible = false
        LELabel.visible = true
        BackboxTrimLeft.visible = true
        BackboxTrimRight.visible = true
        VR_FlipRight.image = "FlipperRed"
        VR_FlipperLeft.image = "FlipperRed"
        PinCab_FlipperRing.image = "FlipperRed"
    End if

    If VRCabArt = 3 then
        VRBackglassFlasher.imageA = "VRBG5"
        VRBackglassFlasher2.imageA = "VRBG5"
        VRBackglassFlasher2.imageB = "VRBG5"
        VRBackglassFlasher2.amount = 100
        VRBackglassFlasher2.opacity = 100
        VRBackglassFlasherFaded.imageA = "VRBG5faded"
        VRBackglassFlasherFaded.imageB = "VRBG5faded"
        VRBackglassFlasherFaded.amount = 100
        VRBackglassFlasherFaded.opacity = 100
    VRBackglassFlasherBulb.amount = 0
    VRBackglassFlasherBulb.opacity = 0
        PinCab_Cabinet.image = "Pincab_Cabinet_CharsNEW"
        PinCab_Backbox.image = "Pincab_Backbox_Chars"
        BackboxTrimRight.visible = false
        BackboxTrimLeft.visible = false
        LELabel.visible = false
    end If

    If VRCabArt = 4 then
        VRBackglassFlasher.imageA = "VRBG4"
        VRBackglassFlasher2.imageA = "VRBG4"
        VRBackglassFlasher2.imageB = "VRBG4"
        VRBackglassFlasher2.amount = 95
        VRBackglassFlasher2.opacity = 25
        VRBackglassFlasherFaded.imageA = "VRBG4faded"
        VRBackglassFlasherFaded.imageB = "VRBG4faded"
        VRBackglassFlasherFaded.amount = 75
        VRBackglassFlasherFaded.opacity = 100
    VRBackglassFlasherBulb.amount = 30
    VRBackglassFlasherBulb.opacity = 100
        PinCab_Cabinet.image = "Pincab_Cabinet_CharsNEW"
        PinCab_Backbox.image = "Pincab_Backbox_Chars"
        LELabel.visible = false
    end If
End Sub

' VR Topper....
Dim TopperLight
Topperlight = 0

' VR Topper Attract sequence
Sub VRTopperTimer_timer
    TopperLight = Topperlight + 1
    If TopperLight = 25 then TopperLight = 1
    If TopperLight = 1 then StartVRDragon
    If TopperLight = 13 then StopVRDragon
    If TopperLight = 1 or TopperLight = 13 Then TopperLightsOff:VRMartellON
    If TopperLight = 2 or TopperLight = 14 Then TopperLightsOff:VRGreyON
    If TopperLight = 3 or TopperLight = 15 Then TopperLightsOff:VRStarkON
    If TopperLight = 4 or TopperLight = 16 Then TopperLightsOff:VRTargON
    If TopperLight = 5 or TopperLight = 17 Then TopperLightsOff:VRLanON
    If TopperLight = 6 or TopperLight = 18 Then TopperLightsOff:VRBarON
    If TopperLight = 7 or TopperLight = 19 Then TopperLightsOff:VRTyrellON
    If TopperLight = 8 or TopperLight = 20 Then TopperLightsOff:VRBarON
    If TopperLight = 9 or TopperLight = 21 Then TopperLightsOff:VRLanON
    If TopperLight = 10 or TopperLight = 22 Then TopperLightsOff:VRTargON
    If TopperLight = 11 or TopperLight = 23 Then TopperLightsOff:VRStarkON
    If TopperLight = 12 or TopperLight = 24 Then TopperLightsOff:VRGreyON
End Sub


' All of the Topper VR light control subs below...
Sub VRMartellON
    VRMartell.imageA = "MartellOn"
    VRMartellFlash.visible = true
End Sub
Sub VRMartellOFF
    VRMartell.imageA = "MartellOff"
    VRMartellFlash.visible = false
End Sub

Sub VRGreyON
    VRGrey.imageA = "GreyJoyOn"
    VRGreyFlash.visible = true
End Sub
Sub VRGreyOFF
    VRGrey.imageA = "GreyJoyOff"
    VRGreyFlash.visible = false
End Sub

Sub VRStarkON
    VRStark.imageA = "StarkOn"
    VRStarkFlash.visible = true
End Sub
Sub VRStarkOFF
    VRStark.imageA = "StarkOff"
    VRStarkFlash.visible = false
End Sub

Sub VRTargON
    VRTarg.imageA = "TargOn"
    VRTargFlash.visible = true
End Sub
Sub VRTargOFF
    VRTarg.imageA = "TargOff"
    VRTargFlash.visible = false
End Sub

Sub VRLanON
    VRLan.imageA = "LanOn"
    VRLanFlash.visible = true
End Sub
Sub VRLanOFF
    VRLan.imageA = "LanOff"
    VRLanFlash.visible = false
End Sub

Sub VRBarON
    VRBar.imageA = "BarOn"
    VRBarFlash.visible = true
End Sub
Sub VRBarOFF
    VRBar.imageA = "BarOff"
    VRBarFlash.visible = false
End Sub

Sub VRTyrellON
    VRTyrell.imageA = "TyrellOn"
    VRTyrellFlash.visible = true
End Sub
Sub VRTyrellOFF
    VRTyrell.imageA = "TyrellOff"
    VRTyrellFlash.visible = false
End Sub

Sub TopperLightsOff
    VRMartell.imageA = "MartellOff"
    VRGrey.imageA = "GreyJoyOff"
    VRStark.imageA = "StarkOff"
    VRTarg.imageA = "TargOff"
    VRLan.imageA = "LanOff"
    VRBar.imageA = "BarOff"
    VRTyrell.imageA = "TyrellOff"

    VRMartellFlash.visible = false
    VRGreyFlash.visible = false
    VRStarkFlash.visible = false
    VRTargFlash.visible = false
    VRLanFlash.visible = false
    VRBarFlash.visible = false
    VRTyrellFlash.visible = false
End Sub

'************************
' VR Dragon Control
'************************

'Call these to start or stop the VR Dragon topper at any time.  It runs the wings and Strobe lights and smoke.
'StartVRDragon
'StopVRDragon

Dim VRWingSteps
VRWingSteps = Array(0,0,1,3,9,15,21,27,29,30,30,29,27,21,15,9,3,1,0)
Dim SmokeRunning
SmokeRunning = false

Sub StartVRDragon
  debug.print "Start VR Dragon"
    VRtmrDragonWings.UserValue = 0
    VRtmrDragonWings.Interval = 75
    VRtmrDragonWings.Enabled = True
  VRTopperLightTimer.enabled = true
  SmokeRunning = true
End Sub

Sub StopVRDragon
  debug.print "Stop VR Dragon"
    VRTopperLightTimer.enabled = false
    VRTopperLightRight.visible = False
    VRTopperLightLeft.visible = False
    VRTopperLightCenter.visible = true
    dragonbody.image = "Dragon_body_unlit_texture"
    dragonleftwing.image = "Dragon_leftwing_unlit_texture"
    dragonrightwing.image = "Dragon_rightwing_unlit_texture"
    SmokeRunning = false

    If VRtmrDragonWings.UserValue > 0 Then
       VRtmrDragonWings.Uservalue = VRtmrDragonWings.UserValue * -1
    Else If VRtmrDragonWings.UserValue = 0 Then
        VRtmrDragonWings.Enabled = False
    End If
end if
End Sub

Sub VRtmrDragonWings_Timer
    Dim i,wi
    i = VRtmrDragonWings.UserValue
    i = i + 1
    If i > 18 Then i = 1
    DragonRightWing.ObjRotY = WingSteps(Abs(i))
    DragonLeftWing.ObjRotY = WingSteps(Abs(i))*-1
    VRtmrDragonWings.UserValue = i
    If i = 0 Then VRtmrDragonWings.Enabled = False
End Sub


Sub VRTopperLightTimer_Timer
    if VRTopperLightRight.visible = False then
        VRTopperLightRight.visible = true
        VRTopperLightLeft.visible = false
        VRTopperLightCenter.visible = false
        dragonbody.image = "Dragon_body_unlit_texture"
        dragonleftwing.image = "Dragon_leftwing_unlit_texture"
        dragonrightwing.image = "Dragon_rightwing_unlit_texture"
    else
        VRTopperLightRight.visible = false
        VRTopperLightLeft.visible = true
        dragonbody.image = "Dragon_body_lit_texture"
        dragonleftwing.image = "Dragon_leftwing_lit_texture"
        dragonrightwing.image = "Dragon_rightwing_lit_texture"
        VRTopperLightCenter.visible = true
        exit sub
    end If
End Sub


dim SmokeBaseHeight : SmokeBaseHeight = 1780 ' default was 1780, use 2100 for ALT
dim SmokeRiseLimit : SmokeRiseLimit = 420
dim RiseRate1 : RiseRate1 = 7
dim RiseRate2 : RiseRate2 = 9

Sub SmokeTimer_timer()
    Smoke.height = Smoke.height + RiseRate1
    Smoke.opacity = Smoke.opacity  - 0.51 ' only 0.51 or higher works now?  so confused.  WHY????
    Smoke2.height = Smoke2.height + RiseRate2
    Smoke2.opacity = Smoke2.opacity  - 0.51 ' only 0.51 or higher works now?  so confused.  WHY????

    if Smoke.height > (SmokeBaseHeight+SmokeRiseLimit) then
        if SmokeRunning = true then
            Smoke.height = SmokeBaseHeight
            Smoke.opacity = 40
            if Smoke.Roty = 0 Then
                Smoke.Roty = 180
            else
                Smoke.Roty = 0
            end if
      RiseRate1 = RndInt(6,10)
        end if
    end If

    if Smoke2.height > (SmokeBaseHeight+SmokeRiseLimit) then
        if SmokeRunning = true then
            Smoke2.height = SmokeBaseHeight
            Smoke2.opacity = 40
            if Smoke2.Roty = 0 Then
                Smoke2.Roty = 180
            else
                Smoke2.Roty = 0
            end if
      RiseRate2 = RndInt(6, 10)
        end if
    end if

End sub

Sub SmokeTimerALT_timer()
    Smoke.height = Smoke.height + 4
    Smoke.opacity = Smoke.opacity  - 0.61 ' only 0.51 or higher works now?  so confused.  WHY????
    Smoke2.height = Smoke2.height + 4
    Smoke2.opacity = Smoke2.opacity  - 0.61 ' only 0.51 or higher works now?  so confused.  WHY????

    if Smoke.height > (SmokeBaseHeight+SmokeRiseLimit) then
        if SmokeRunning = true then
            Smoke.height = SmokeBaseHeight
            Smoke.opacity = 40
           ' if Smoke.Roty = 0 Then
                Smoke.Roty = 180
            'else
            '    Smoke.Roty = 0
            'end if
        end if
    end If

    if Smoke2.height > (SmokeBaseHeight+SmokeRiseLimit) then
        if SmokeRunning = true then
            Smoke2.height = SmokeBaseHeight
            Smoke2.opacity = 40
            'if Smoke2.Roty = 0 Then
            '    Smoke2.Roty = 180
            'else
            '    Smoke2.Roty = 0
            'end if
        end if
    end if

End sub

Dim VRSmokeSpeed : VRSmokeSpeed = 30 'Randomize the dragon speed

sub SmokeTimerDragon_timer()
    Smoke3.height = Smoke3.height + VRSmokeSpeed
    Smoke3.opacity = Smoke3.opacity  - 0.51 ' only 0.51 or higher works now?  so confused.  WHY????
    Smoke4.height = Smoke4.height + VRSmokeSpeed
    Smoke4.opacity = Smoke4.opacity  - 0.51 ' only 0.51 or higher works now?  so confused.  WHY????
    if Smoke3.height > 16940 or smoke4.height > 16940 then
        if SmokeRunning = true then
            Smoke3.height = 12000
            Smoke3.opacity = 50
      Smoke4.height = 12000
            Smoke4.opacity = 50
        VRSmokeSpeed = Int((91 - 30 + 1) * Rnd + 30)
        end if
    end If

end sub


Sub VRTopperLights_Timer
    if p32on.DisableLighting = 0 then
    VRTargOFF
    Else
    VRTargON
    end If

    if p38on.DisableLighting = 0 then
    VRStarkOFF
    Else
    VRStarkON
    end If

    if p41on.DisableLighting = 0 then
    VRBarOFF
    Else
    VRBarON
    end If

    if p44on.DisableLighting = 0 then
    VRLanOFF
    Else
    VRLanON
    end If

    if p47on.DisableLighting = 0 then
    VRGreyOFF
    Else
    VRGreyON
    end If

    if p50on.DisableLighting = 0 then
    VRTyrellOFF
    Else
    VRTyrellON
    end If

    if p53on.DisableLighting = 0 then
    VRMartellOFF
    Else
    VRMartellON
    end If
End Sub

'*******************************************
'  VR ROOM Lighting
'*******************************************
VR_Ironthrone001.blenddisablelighting = 0.2

' End VR stuff ***********************************************************************************

'*******************************************
' Settings
'*******************************************

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reset
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

Sub Table1_OptionEvent(ByVal eventId)

  dim val, daynight

  val = Table1.Option("Dragon Flames", 0, 1, 1, 1, 0, Array("Off","On"))
  Options_Flames_Enabled = val

  val = Table1.Option("Cabinet Mode", 0, 1, 1, 0, 0, Array("Off","On"))
  CabinetMode = val

  val = Table1.Option("Physical Lockbar Button", 0, 1, 1, 0, 0, Array("Off","On"))
  bHaveLockbarButton = val

  val = Table1.Option("Dynamic Ball Shadows", 0, 1, 1, 1, 0, Array("Off","On"))
  DynamicBallShadowsOn = val

  BallImageId = Table1.Option("Ball Image", 0, 2, 1, 2, 0, Array("Bright","Old Arse","Old Arse Warm"))

  'to keep compatability with DMD settings, store volume as a 0-1 value
  VolumeDial = Table1.Option("Mech Volume", .00, 1, .1, .80, 1)

  'to keep compatability with DMD settings, store volume as a 0-1 value
  RampRollVolume = Table1.Option("Ramp Volume", .00, 1, .1, .50, 1)

  'to keep compatability with DMD settings, store volume as a 0-1 value
  BallRollVolume = Table1.Option("Ball Volume", .00, 1, .1, .50, 1)

  'prev DMD settings use a brightness 1-100
  'LightLevel = cInt(NightDay)
  LightLevel = cInt(Table1.Option("Brightness", 0, 100, 5, 70, 0))

  'LUTImage -- array is only used to avoid the decimal point, also allows looping
  LUTImage = Table1.Option("LUT", 0, 15, 1, 0, 0, Array("0","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15"))

  ' Playfield Reflections Scope
  ReflectionScope = Table1.Option("Playfield Reflections Scope", 0, 2, 1, 1, 0, Array("Reflect Ball Only", "Reflect Ball & Parts", "Reflect Ball, Parts & Lighting"))
  ReflectionToggle ReflectionScope

  ReflectionStrength = Table1.Option("Reflection Strength", .01, 1, .01, .20, 1)
  ReflectionStrengthSetup

  'VR settings
  VRRoomChoice = Table1.Option("VR Room", 1, 5, 1, 5, 0, Array("Minimal","Ultra Minimal", "Iron Throne Mega", "Dragon Ice Mega", "Tower Throne Mega"))

  val = Table1.Option("VR Glass Scratches", 0, 1, 1, 0, 0, Array("Off","On"))
  Scratches = val

  val = Table1.Option("VR Topper", 0, 1, 1, 1, 0, Array("Off","On"))
  Topper = val

  If eventId < 3 Then
    debug.print "Options changed: " & eventId
    UpdateMods ' used primarily for old settings
    table1.ColorGradeImage = "LUT"&LUTImage
    VRRoomInit
  End If

  If eventId = 3 Then
    Debug.print "Closing tweak menu"
    SaveDMDSettings
  End If

  If eventId = 0 Then
    Debug.print "Loaded Settings"
    VRRoomInit
  End If

End Sub

sub WriteSettings
  'writes all the settings, used by other DMD settings to keep settings in sync if saved from DMD menu
  'dim val, daynight

  debug.print "Writing settings"

  Table1.Option("Dragon Flames", 0, 1, 1, 1, 0, Array("Off","On")) = Options_Flames_Enabled
  Table1.Option("Cabinet Mode", 0, 1, 1, 0, 0, Array("Off","On")) = CabinetMode
  Table1.Option("Physical Lockbar Button", 0, 1, 1, 0, 0, Array("Off","On")) = bHaveLockbarButton
  Table1.Option("Dynamic Ball Shadows", 0, 1, 1, 1, 0, Array("Off","On")) = DynamicBallShadowsOn
  Table1.Option("Ball Image", 0, 2, 1, 2, 0, Array("Bright","Old Arse","Old Arse Warm")) = BallImageId
  Table1.Option("Mech Volume", .00, 1, .1, .80, 1) = VolumeDial
  Table1.Option("Ramp Volume", .00, 1, .1, .50, 1) = RampRollVolume
  Table1.Option("Ball Volume", .00, 1, .1, .50, 1) = BallRollVolume

  Table1.Option("LUT", 0, 15, 1, 0, 0, Array("0","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15")) = LUTImage
  Table1.Option("Playfield Reflections Scope", 0, 2, 1, 1, 0, Array("Reflect Ball Only", "Reflect Ball & Parts", "Reflect Ball, Parts & Lighting")) = ReflectionScope
  Table1.Option("Reflection Strength", .01, 1, .01, .20, 1) = ReflectionStrength

  Table1.Option("VR Room", 1, 5, 1, 5, 0, Array("Minimal","Ultra Minimal", "Iron Throne Mega", "Dragon Ice Mega", "Tower Throne Mega")) = VRRoomChoice

  Table1.Option("VR Glass Scratches", 0, 1, 1, 0, 0, Array("Off","On")) = Scratches
  Table1.Option("VR Topper", 0, 1, 1, 1, 0, Array("Off","On")) = Topper

  'prev DMD settings use a brightness 1-100
  'NightDay = LightLevel
  Table1.Option("Brightness", 0, 100, 5, 65, 0) = LightLevel

end sub

sub ReflectionToggle(value)
  Select Case value
    Case 0: ReflectionsSetup false, false 'Only reflect the ball
    Case 1: ReflectionsSetup true, false  'Only reflect a selection of bakemaps
    Case 2: ReflectionsSetup true, true   'Reflect everything (that matters as defined in ReflObjectsArray)
  End Select
end sub

sub ReflectionStrengthSetup
  table1.PlayfieldReflectionStrength = CInt(ReflectionStrength * 100)
end Sub

sub ReflectionsSetup(ReflectObject, ReflectAllLight)

  'main playfield
  '--------------

  'spinners
  Spinners_BM_Dark_Room.ReflectionEnabled = ReflectObject
  Spinners_LM_Lit_Room.ReflectionEnabled = ReflectObject

  'left drop targets
  't9_BM_Dark_Room.ReflectionEnabled = ReflectObject
  t9_LM_Lit_Room.ReflectionEnabled = ReflectObject
  t9_LM_Inserts_li17.ReflectionEnabled = ReflectAllLight

  't8_BM_Dark_Room.ReflectionEnabled = ReflectObject
  t8_LM_Lit_Room.ReflectionEnabled = ReflectObject
  t8_LM_Inserts_li20.ReflectionEnabled = ReflectAllLight
  't8_LM_Inserts_li17.ReflectionEnabled = ReflectAllLight
  't8_LM_Inserts_li23.ReflectionEnabled = ReflectAllLight

  't7_BM_Dark_Room.ReflectionEnabled = ReflectObject
  t7_LM_Lit_Room.ReflectionEnabled = ReflectObject
  t7_LM_Inserts_li23.ReflectionEnabled = ReflectAllLight
  't7_LM_Inserts_li20.ReflectionEnabled = ReflectAllLight

  'gold targets left to right
  targets_004_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_004_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_004_LM_Inserts_li86.ReflectionEnabled = ReflectAllLight
  targets_004_LM_Inserts_li92.ReflectionEnabled = ReflectAllLight
  'targets_004_LM_Inserts_li95.ReflectionEnabled = ReflectAllLight

  targets_005_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_005_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_005_LM_Inserts_li105.ReflectionEnabled = ReflectAllLight
  'targets_005_LM_Inserts_li117.ReflectionEnabled = ReflectAllLight

  targets_006_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_006_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_006_LM_Inserts_li117.ReflectionEnabled = ReflectAllLight
  targets_006_LM_Inserts_li120.ReflectionEnabled = ReflectAllLight
  targets_006_LM_Inserts_li132.ReflectionEnabled = ReflectAllLight

  targets_007_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_007_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_007_LM_Inserts_li135.ReflectionEnabled = ReflectAllLight

  targets_008_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_008_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_008_LM_Inserts_li147.ReflectionEnabled = ReflectAllLight

  'castle drop
  t90_BM_Dark_Room.ReflectionEnabled = ReflectObject
  t90_LM_Lit_Room.ReflectionEnabled = ReflectObject
  t90_LM_Inserts_li144.ReflectionEnabled = ReflectAllLight
  't90_LM_Inserts_li141.ReflectionEnabled = ReflectAllLight
  't90_LM_Inserts_li132.ReflectionEnabled = ReflectAllLight

  'wildfire targets
  targets_009_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_009_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_009_LM_Inserts_li80.ReflectionEnabled = ReflectAllLight

  targets_010_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_010_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_010_LM_Inserts_li83.ReflectionEnabled = ReflectAllLight

  'sword
  'sword_003_BM_Dark_Room.ReflectionEnabled = ReflectObject
  sword_003_LM_Lit_Room.ReflectionEnabled = ReflectObject
  'sword_003_LM_Inserts_li77.ReflectionEnabled = ReflectAllLight
  'sword_003_LM_Inserts_li80.ReflectionEnabled = ReflectAllLight

  'sword_002_BM_Dark_Room.ReflectionEnabled = ReflectObject
  sword_002_LM_Lit_Room.ReflectionEnabled = ReflectObject
  'sword_002_LM_Inserts_li80.ReflectionEnabled = ReflectAllLight

  'most ramps, upper pf
  Parts_BM_Dark_Room.ReflectionEnabled = ReflectObject

  'upper playfield
  '---------------

  'targets_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_LM_Inserts_li189.ReflectionEnabled = ReflectAllLight

  'targets_011_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_011_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_011_LM_Inserts_li195.ReflectionEnabled = ReflectAllLight
  targets_011_LM_Inserts_li192.ReflectionEnabled = ReflectAllLight
  targets_011_LM_Inserts_li183.ReflectionEnabled = ReflectAllLight

  'targets_012_BM_Dark_Room.ReflectionEnabled = ReflectObject
  targets_012_LM_Lit_Room.ReflectionEnabled = ReflectObject
  targets_012_LM_Inserts_li192.ReflectionEnabled = ReflectAllLight
  targets_012_LM_Inserts_li195.ReflectionEnabled = ReflectAllLight
  targets_012_LM_Inserts_li189.ReflectionEnabled = ReflectAllLight
  targets_012_LM_Inserts_li183.ReflectionEnabled = ReflectAllLight
  targets_012_LM_Inserts_li180.ReflectionEnabled = ReflectAllLight
  targets_012_LM_Inserts_li92.ReflectionEnabled = ReflectAllLight

  Rwing_LM_Lit_Room.ReflectionEnabled = ReflectObject
  Lwing_LM_Lit_Room.ReflectionEnabled = ReflectObject
  Parts_BM_Dark_Room_004.ReflectionEnabled = ReflectObject 'dragon head,throne ramp

end Sub


sub SaveDMDSettings

  'saving DMD settings as a way to keep DMD settings used in original menu system and 10.8 in sync.

  SaveValue cGameName, "LIGHT", LightLevel
  SaveValue cGameName, "DRAGONFLAMES", Options_Flames_Enabled
  SaveValue cGameName, "VOLUME", VolumeDial
  SaveValue cGameName, "RAMPVOLUME", RampRollVolume
  SaveValue cGameName, "BALLVOLUME", BallRollVolume
  SaveValue cGameName, "BALLIMAGE", BallImageId
  SaveValue cGameName, "CABINET", CabinetMode
  SaveValue cGameName, "LBBUT", bHaveLockbarButton
  SaveValue cGameName, "DYNBALLSH", DynamicBallShadowsOn
  SaveValue cGameName, "VRROOM", VRRoomChoice
  SaveValue cGameName, "VRSCRATCHES", Scratches
  SaveValue cGameName, "VRTOPPER", Topper

  SaveLUT
  'VRRoomInit

end sub


'********* Fleep's mech sounds ***********

'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                                                                                                                'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                                                                                                        'volume level; range [0, 1]
NudgeRightSoundLevel = 1                                                                                                'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                                                                                                'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                                                                                                'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                                                                                        'volume level; range [0, 1]
PlungerPullSoundLevel = 1                                                                                                'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                                                           'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
SlingshotSoundLevel = 0.95                                                                                                'volume level; range [0, 1]
BumperSoundFactor = 4.25                                                                                                'volume multiplier; must not be zero
KnockerSoundLevel = 4 '1/VolumeDial                                                                                                         'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                                                                                        'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                                                                                'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                                                                        'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                                                                        'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                                                                                        'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                                                                                                        'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                                                      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                                                                                                                'volume level; range [0, 1]
BallReleaseSoundLevel = 1                                                                                                'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                                                                        'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                                                                                'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                                                                                                        'volume multiplier; must not be zero


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
    tmp = 1
    On Error Resume Next
  tmp = tableobj.y * 2 / tableheight-1
    On Error Goto 0

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
    tmp = 1
    On Error Resume Next
  tmp = tableobj.x * 2 / tablewidth-1
    On Error Goto 0

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
    'PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
    PlaySoundVol "stern-knocker",10
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
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
    FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
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
    If FlipperLeftHitParm > 1 Then FlipperLeftHitParm = 1
    FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
    RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
    FlipperRightHitParm = parm/10
    If FlipperRightHitParm > 1 Then FlipperRightHitParm = 1
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

Sub SoundSaucerLock(saucer)
    PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, saucer
End Sub

Sub SoundSaucerKick(scenario, saucer)
    Select Case scenario
        Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
        Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
    End Select
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd*6)+1), 200, obj
End Sub


'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  FlipperCradleCollision ball1, ball2, velocity
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
'                                        End Mechanical Sounds
'/////////////////////////////////////////////////////////////////




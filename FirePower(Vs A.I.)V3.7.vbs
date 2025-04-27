'                                _      ________   __   _______   __  _______  ______  ___  ___
'                               | | /| / /  _/ /  / /  /  _/ _ | /  |/  / __/ <  / _ \( _ )/ _ \
'                               | |/ |/ // // /__/ /___/ // __ |/ /|_/ /\ \   / /\_, / _  / // /
'                               |__/|__/___/____/____/___/_/ |_/_/  /_/___/  /_//___/\___/\___/
'
'                                    ______________  __________  ____ _       ____________
'                  **************   / ____/  _/ __ \/ ____/ __ \/ __ \ |     / / ____/ __ \   ***************
'                 **************   / /_   / // /_/ / __/ / /_/ / / / / | /| / / __/ / /_/ /  ***************
'                **************   / __/ _/ // _, _/ /___/ ____/ /_/ /| |/ |/ / /___/ _, _/  ***************
'               **************   /_/   /___/_/ |_/_____/_/    \____/ |__/|__/_____/_/ |_|  ***************
'
'                                               _   __        ___     ____
'                             ***********      | | / /__     / _ |   /  _/      ***********
'                            ***********       | |/ (_-<_   / __ |_ _/ /_      ***********
'                           ***********        |___/___(_) /_/ |_(_)___(_)    ***********
'                                                        '
'
'                A 3rdaxis, Rothbauerw, G5K & Slydog43 Collaboration V3.7
'
'
'(Hit ESC anytime during game to display instructions)
'      ///Table Instructions & Options Below///.

Option Explicit
Randomize

On Error Resume Next
  ExecuteGlobal GetTextFile("controller.vbs")
  If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="frpwr_b7"
'Const UseSolenoids=1
Const UseSolenoids=25 'FastFlips for system6
Const UseLamps = True 'VPM to control Lights
Const UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"
Const BallSize = 50
Const Ballmass = 1

LoadVPM "01560000", "S6.VBS", 3.26

'***********************************************************************************************************************************

Dim T1Step,T2Step,T3Step,T4Step,T5Step,T6Step,topPWRStep,midPWRStep,BtmPWRStep,CTStep,DTMod, trTrough, MyPrefs_Brightness, MyPrefs_DisableBrightness
Dim myObject, myPrefs_Drain, myPrefs_SoundEffects, myPrefs_Soundonoff, myPrefs_PlayMusic, myPrefs_ToggleSoundsOnOffKey, MyPrefs_DynamicCards, myPrefs_FStobesFrequency
Dim myPrefs_OutlaneDifficulty, myPrefs_InstructionCardType, myPrefs_Targets, myPrefs_DynamicBackground, myPrefs_PlayfieldLamps, AITaunt, myPrefs_AttractMode
Dim myPrefs_Player1Setup, myPrefs_Player2Setup, myPrefs_Player3Setup, myPrefs_Player4Setup, myPrefs_Difficulty, myPlayingEndGameSnd, myPrefs_FactorySetting
Dim myPrefs_FlipperBatType, BouncePassOnOff,myPrefs_AInudge, myPrefs_DisplayColor, myPrefs_Hal9000EyeOnOff, myPrefs_GlassHit, myPrefs_CustomPlayerSetup, myPrefs_FlipperStobes
Dim myPrefs_LogicPanelOnOffKey, myPrefs_ShowLogicPanel, myPanelDirection, myPrefs_Language, myPrefs_LanguageOnOffKey, myPrefs_ShowDirections, myPrefs_GameDataTracking

myPlayingEndGameSnd = False
myPrefs_Drain = 0
myPrefs_ShowLogicPanel = 0
myPanelDirection = -1

Dim myCreditCounter
myCreditCounter= 0

' _________   ___  __   ____  _____  _______________  __  _____________________  _  ______  ____      ____  ___  ______________  _  ______
'/_  __/ _ | / _ )/ /  / __/ /  _/ |/ / __/_  __/ _ \/ / / / ___/_  __/  _/ __ \/ |/ / __/ / __/___  / __ \/ _ \/_  __/  _/ __ \/ |/ / __/
' / / / __ |/ _  / /__/ _/  _/ //    /\ \  / / / , _/ /_/ / /__  / / _/ // /_/ /    /\ \   > _/_ _/ / /_/ / ___/ / / _/ // /_/ /    /\ \ 
'/_/ /_/ |_/____/____/___/ /___/_/|_/___/ /_/ /_/|_|\____/\___/ /_/ /___/\____/_/|_/___/  |_____/   \____/_/    /_/ /___/\____/_/|_/___/

'***********************************************************************************************************************************************************


'*Use the "L" key on keyboard to select language.
'*Use Flipper Buttons to setup human and or computer players before starting game. (Green Human, Red Computer).
'*While holding both flipper buttons use the Magna save buttons to select difficulty. (Easy, Medium, Hard, Expert and Legendary)
'*Use Magna Save buttons to select flipper bat type.
'*Once game is started no player or difficulty settings can be changed.
'*The A.I. can be "tilted out" during the A.I. turn by pressing and holding both Magna Save buttons if you wish to skip his turn.
'*Brightness can be change using the Magna-save buttons after the game is started.
'*myPrefs_CustomPlayerSetup must be set to 1 for changes to PlayerSetup 1-4 in the script below.
'*AI nudge when set to "0" will always be off.

'🚀🚀🚀🚀🚀🚀🚀🚀🚀🚀🚀🚀🚀🚀🚀 MAIN
myPrefs_FactorySetting = 0            '0=Off, 1=On This will automatically set options to a factory purest State. This will remove the A.I. completely.

'⌨⌨⌨⌨⌨⌨⌨⌨⌨⌨⌨⌨ LANGUAGE
myPrefs_Language = 0            '(NEW)     '0=English 1=German 2=Italian 3=French (Can be changed in game with the "L" key)
myPrefs_LanguageOnOffKey = 38                '("L" Key by Default) Use keycodes found in https://www.vpforums.org/Tutorials/KeyCodes.html to change for your set up.

'🛸🛸🛸🛸🛸🛸🛸🛸🛸🛸🛸🛸🛸🛸🛸 A.I.
myPrefs_CustomPlayerSetup = 0               '0=Default 1=Custom Setup. This must be set to 1 if changing the PlayerSetup below. This will lock-out ability to change in game. Player 1-4 will be set to "Default" if off.
myPrefs_Player1Setup = 0                         '0=Human 1=Computer Default(0)
myPrefs_Player2Setup = 1                         '0=Human 1=Computer Default(1)
myPrefs_Player3Setup = 1                         '0=Human 1=Computer Default(1)
myPrefs_Player4Setup = 1                         '0=Human 1=Computer Default(1)
myPrefs_Difficulty = 4                             '0=Easy 1=Normal 2=Hard 3=Expert 4=Legendary (Computer's Level Of Skill) While holding both flipper buttons use the Magna save buttons to select difficulty.
myPrefs_ShowDirections = 1                     '0=Off 1=On Shows Direction menu at the start of the first game.
myPrefs_Hal9000EyeOnOff = 1                   '0=Off 1=On Turns the visibility of the Hall9000's eye On/Off
myPrefs_SoundEffects = 1                         '0=off 1=On Extra Sound Effects And Lighting During Computer's Turn,
myPrefs_AInudge = 3                                   '0=Off, 1-5, 1=Minimum nudge, 5=Maximum nudge (Legendary Difficulty will default to "5" Maximum nudge)
                                                                       '*Note: If set to zero (0) Outlane Difficulty will move post to the easy possition for the A.I.
'📊📊📊📊📊📊📊📊📊📊📊📊📊📊 TABLE
BallShadowUpdate.Enabled = 1                 '0=Off 1=On Off=Performance On=Quality
myPrefs_OutlaneDifficulty = 0               '0=Factory 1=Easy (Right Outlane Post Position)
myPrefs_InstructionCardType = 1           '0=Factory 1=Modded Cards 2=THX Cards 3=U.I.Instruction cards (See THX setup below for directions to properly adjust your TV or Monitor)
MyPrefs_DynamicCards = 1                         '1=Instruction Cards will change to your prefered seletion when game is started and back to U.I.Instruction cards when game is over.
myPrefs_Targets = 0                                   '0=Factory 1=Modded
myPrefs_PlayfieldLamps  = 1                    '0=Factory 1=LED
myPrefs_DisplayColor = 0                         '0=Factory 1=Blue 2=Green 3=Red 4=White  Changes color of score displays.
myPrefs_GlassHit = 1                                 '0=Off, 1=Less often, 2=More often (This determines how often the ball will smack against the glass)
MyPrefs_Brightness = 4                             '0= 10% contrast/brightness, 1= 30%, 2= 50%, 3= 70%, 4= 100%
MyPrefs_DisableBrightness = 0               'Disables the ability to change Brightness option with magna saves in game when set to 1
myPrefs_PlayMusic = 1                               '0=Factory 1=On  Begining And End Game Extra Music.
myPrefs_DynamicBackground = 1               '0=Off 1=On
myPrefs_FlipperStobes = 3   '(New)    '0=Off, 1=Less often, 3=More often (Flippers will Strobe during certain parts of the game if a Glow-Bat type is selected)
myPrefs_AttractMode = 1       '(New)    '0=Off 1=On Will Turn on/off Flipper Flashers during attract mode.
myPrefs_FlipperBatType = 6 '(New)    'Flipper Bat Type. Option list below.

'*Use left/right Magna Save button to change in game before start.
'Can still be hard set as desired.

'0=  White Bat Red Rubber  (Factory)
'1=  White Bat Black Rubber
'2=  White Bat White Rubber
'3=  Yellow Bat Red Rubber
'4=  Yellow Bat Black Rubber
'5=  Red Glow Mod
'6=  Green Glow Mod
'7=  Green/Red Glow Mod
'8=  Red/Green Glow Mod
'9=  Blue Glow Mod
'10= White Glow Mod
'11= Yellow Glow Mod          (New)
'12= Red/Blue Glow Mod        (New)
'13= White Bat Yellow Rubber  (New)

'Recommendations:
'In-game AO should be left "Off"
'SCSP Reflect should be left "On"
'Make sure your "Sound Effects Volume" is set to maximum under "Preferences/Audio Options..."
'Go through the THX set-up (instructions below)

'THX Setup. (Select Instruction Card 2 in the script options)
'<Basic adjustment>
'While in game use the Magnasave buttons to adjust the brightness.
'During adjustment the THX Logo's shadow should just barely be visible on the right player card. This will usually get you close enough.
'<Advanced Adjustment>
'The right card is for adjusting yourTV's Brightness.
'While adjusting your brightness the THX Logo's shadow should just barely be visible and all 9 squares should be a different shade.
'The left card is for adjusting your TV's contrast.
'While adjusting your contrast all 8 squares should be a different shade.
'You may have to go back and forth between brightness and contrast to get the right ballence and calibration.

'//////////////////////////////////////////////////////////////////////
'// LAYER 01 = Timers, Flippers, Triggers, Rubbers, Spinner, Walls
'// LAYER 02 = Lamps B, Upper Walls
'// LAYER 03 = GI, VPX Pop Bumpers
'// LAYER 04 = Lamps
'// LAYER 05 = Flashers
'// LAYER 06 = AI Logic elements
'// LAYER 07 = Targets, Target Primitives, Target Laser's
'// LAYER 08 = Glass
'// LAYER 09 = G5K Primitives, RailsLockbar, Blades, U.I. Menu
'// LAYER 10 = G5K Primitives, GRP1
'// LAYER 11 = G5K Primitives, GRP2, Pop Bumper,Outlane Posts
'//////////////////////////////////////////////////////////////////////

'*HIDDEN MENU (AFTER GAME HAS STARTED)

'*Press and hold the Right Flipper Button for six seconds to display Logic panel, Player and Difficulty Status during game play. Logic Panel displays during DTmode only.
'*While still holding the Right Flipper Button use the Left Magna Save Button to change Flipper Bat Type.
'*While still holding the Right Flipper Button use the Right Magna Save Button to toggle on/off A.I. sound effects and laser sights.
'*While still holding the Right Flipper Button use the Left Flipper Button to change the Score Display colors.

 myPrefs_GameDataTracking = 0     '0=Off, 1= On. Adds extra game score tracking data but will use more CPU. (Dev Tools)
'**********************************************************************************************************************************
'        ______   _______  ___  _______  ____
'*****  / __/ /  /  _/ _ \/ _ \/ __/ _ \/ __/*****
'***** / _// /___/ // ___/ ___/ _// , _/\ \  *****
'*****/_/ /____/___/_/  /_/  /___/_/|_/___/  *****

Sub UITimer_Timer() 'Bad (Need to call instead and get all of this out of a Timer) D-

  if gamestarted=0 then
    if myPrefs_CustomPlayerSetup = 0 Then
      if Directionsoff = 0 then
    select case toggleplayers
      Case 0
        DP1.Visible = 1
        DP2.Visible = 0
        DP3.Visible = 0
        DP4.Visible = 0
        JoeyComp01.Visible = 0
        JoeyComp02.Visible = 1
        JoeyComp03.Visible = 1
        JoeyComp04.Visible = 1
      Case 1
        DP1.Visible = 1
        DP2.Visible = 1
        DP3.Visible = 0
        DP4.Visible = 0
        JoeyComp01.Visible = 0
        JoeyComp02.Visible = 0
        JoeyComp03.Visible = 1
        JoeyComp04.Visible = 1
      Case 2
        DP1.Visible = 1
        DP2.Visible = 1
        DP3.Visible = 1
        DP4.Visible = 0
        JoeyComp01.Visible = 0
        JoeyComp02.Visible = 0
        JoeyComp03.Visible = 0
        JoeyComp04.Visible = 1
      Case 3
        DP1.Visible = 1
        DP2.Visible = 1
        DP3.Visible = 1
        DP4.Visible = 1
        JoeyComp01.Visible = 0
        JoeyComp02.Visible = 0
        JoeyComp03.Visible = 0
        JoeyComp04.Visible = 0
      Case 4
        DP1.Visible = 1
        DP2.Visible = 0
        DP3.Visible = 1
        DP4.Visible = 0
        JoeyComp01.Visible = 0
        JoeyComp02.Visible = 1
        JoeyComp03.Visible = 0
        JoeyComp04.Visible = 1
      Case 5
        DP1.Visible = 0
        DP2.Visible = 1
        DP3.Visible = 0
        DP4.Visible = 1
        JoeyComp01.Visible = 1
        JoeyComp02.Visible = 0
        JoeyComp03.Visible = 1
        JoeyComp04.Visible = 0
      Case 6
        DP1.Visible = 0
        DP2.Visible = 0
        DP3.Visible = 0
        DP4.Visible = 0
        JoeyComp01.Visible = 1
        JoeyComp02.Visible = 1
        JoeyComp03.Visible = 1
        JoeyComp04.Visible = 1
    End Select
  end if
    end if
      end if

if gamestarted = 0 then
   Select Case myPrefs_Difficulty
        Case 0
    LightDifficulty0.Color = RGB(255, 0, 0)
    LightDifficulty0.ColorFull = RGB(255, 0, 0)
    LightDifficulty1.Color = RGB(0, 0, 0)
    LightDifficulty1.ColorFull = RGB(0, 0, 0)
    LightDifficulty2.Color = RGB(0, 0, 0)
    LightDifficulty2.ColorFull = RGB(0, 0, 0)
    LightDifficulty3.Color = RGB(0, 0, 0)
    LightDifficulty3.ColorFull = RGB(0, 0, 0)
    LightDifficulty4.Color = RGB(0, 0, 0)
    LightDifficulty4.ColorFull = RGB(0, 0, 0)
    LightD0.Color = RGB(255, 0, 0)
    LightD0.ColorFull = RGB(255, 0, 0)
    LightD1.Color = RGB(0, 0, 0)
    LightD1.ColorFull = RGB(0, 0, 0)
    LightD2.Color = RGB(0, 0, 0)
    LightD2.ColorFull = RGB(0, 0, 0)
    LightD3.Color = RGB(0, 0, 0)
    LightD3.ColorFull = RGB(0, 0, 0)
    LightD4.Color = RGB(0, 0, 0)
    LightD4.ColorFull = RGB(0, 0, 0)
        Case 1
    LightDifficulty0.Color = RGB(255, 0, 0)
    LightDifficulty0.ColorFull = RGB(255, 0, 0)
    LightDifficulty1.Color = RGB(255, 0, 0)
    LightDifficulty1.ColorFull = RGB(255, 0, 0)
    LightDifficulty2.Color = RGB(0, 0, 0)
    LightDifficulty2.ColorFull = RGB(0, 0, 0)
    LightDifficulty3.Color = RGB(0, 0, 0)
    LightDifficulty3.ColorFull = RGB(0, 0, 0)
    LightDifficulty4.Color = RGB(0, 0, 0)
    LightDifficulty4.ColorFull = RGB(0, 0, 0)
    LightD0.Color = RGB(255, 0, 0)
    LightD0.ColorFull = RGB(255, 0, 0)
    LightD1.Color = RGB(255, 0, 0)
    LightD1.ColorFull = RGB(255, 0, 0)
    LightD2.Color = RGB(0, 0, 0)
    LightD2.ColorFull = RGB(0, 0, 0)
    LightD3.Color = RGB(0, 0, 0)
    LightD3.ColorFull = RGB(0, 0, 0)
    LightD4.Color = RGB(0, 0, 0)
    LightD4.ColorFull = RGB(0, 0, 0)
        Case 2
    LightDifficulty0.Color = RGB(255, 0, 0)
    LightDifficulty0.ColorFull = RGB(255, 0, 0)
    LightDifficulty1.Color = RGB(255, 0, 0)
    LightDifficulty1.ColorFull = RGB(255, 0, 0)
    LightDifficulty2.Color = RGB(255, 0, 0)
    LightDifficulty2.ColorFull = RGB(255, 0, 0)
    LightDifficulty3.Color = RGB(0, 0, 0)
    LightDifficulty3.ColorFull = RGB(0, 0, 0)
    LightDifficulty4.Color = RGB(0, 0, 0)
    LightDifficulty4.ColorFull = RGB(0, 0, 0)
    LightD0.Color = RGB(255, 0, 0)
    LightD0.ColorFull = RGB(255, 0, 0)
    LightD1.Color = RGB(255, 0, 0)
    LightD1.ColorFull = RGB(255, 0, 0)
    LightD2.Color = RGB(255, 0, 0)
    LightD2.ColorFull = RGB(255, 0, 0)
    LightD3.Color = RGB(0, 0, 0)
    LightD3.ColorFull = RGB(0, 0, 0)
    LightD4.Color = RGB(0, 0, 0)
    LightD4.ColorFull = RGB(0, 0, 0)
        Case 3
    LightDifficulty0.Color = RGB(255, 0, 0)
    LightDifficulty0.ColorFull = RGB(255, 0, 0)
    LightDifficulty1.Color = RGB(255, 0, 0)
    LightDifficulty1.ColorFull = RGB(255, 0, 0)
    LightDifficulty2.Color = RGB(255, 0, 0)
    LightDifficulty2.ColorFull = RGB(255, 0, 0)
    LightDifficulty3.Color = RGB(255, 0, 0)
    LightDifficulty3.ColorFull = RGB(255, 0, 0)
    LightDifficulty4.Color = RGB(0, 0, 0)
    LightDifficulty4.ColorFull = RGB(0, 0, 0)
    LightD0.Color = RGB(255, 0, 0)
    LightD0.ColorFull = RGB(255, 0, 0)
    LightD1.Color = RGB(255, 0, 0)
    LightD1.ColorFull = RGB(255, 0, 0)
    LightD2.Color = RGB(255, 0, 0)
    LightD2.ColorFull = RGB(255, 0, 0)
    LightD3.Color = RGB(255, 0, 0)
    LightD3.ColorFull = RGB(255, 0, 0)
    LightD4.Color = RGB(0, 0, 0)
    LightD4.ColorFull = RGB(0, 0, 0)
        Case 4
    LightDifficulty0.Color = RGB(255, 0, 0)
    LightDifficulty0.ColorFull = RGB(255, 0, 0)
    LightDifficulty1.Color = RGB(255, 0, 0)
    LightDifficulty1.ColorFull = RGB(255, 0, 0)
    LightDifficulty2.Color = RGB(255, 0, 0)
    LightDifficulty2.ColorFull = RGB(255, 0, 0)
    LightDifficulty3.Color = RGB(255, 0, 0)
    LightDifficulty3.ColorFull = RGB(255, 0, 0)
    LightDifficulty4.Color = RGB(255, 128, 64)
    LightDifficulty4.ColorFull = RGB(255, 255, 0)
    LightD0.Color = RGB(255, 0, 0)
    LightD0.ColorFull = RGB(255, 0, 0)
    LightD1.Color = RGB(255, 0, 0)
    LightD1.ColorFull = RGB(255, 0, 0)
    LightD2.Color = RGB(255, 0, 0)
    LightD2.ColorFull = RGB(255, 0, 0)
    LightD3.Color = RGB(255, 0, 0)
    LightD3.ColorFull = RGB(255, 0, 0)
    LightD4.Color = RGB(255, 80, 0)
    LightD4.ColorFull = RGB(255, 255, 0)
        End Select
end if
End Sub

Sub GraphicsTimer_Timer()

    batleft.objrotz = FlipperLeft.CurrentAngle + 1
    glowbatleft.objrotz = FlipperLeft.CurrentAngle + 1
    batleftshadow.rotz = batleft.objrotz
    batleftshadowglowred.rotz = batleft.objrotz
    batleftshadowglowgreen.rotz = batleft.objrotz
    batleftshadowglowblue.rotz = batleft.objrotz
    batleftshadowglowwhite.rotz = batleft.objrotz
    batleftshadowglowyellow.rotz  = batleft.objrotz
    batright.objrotz = FlipperRight.CurrentAngle - 1
    glowbatright.objrotz = FlipperRight.CurrentAngle -1
    batrightshadowglowred.rotz  = batright.objrotz
    batrightshadowglowgreen.rotz  = batright.objrotz
    batrightshadowglowblue.rotz  = batright.objrotz
    batrightshadowglowwhite.rotz  = batright.objrotz
    batrightshadow.rotz  = batright.objrotz
    batrightshadowglowyellow.rotz  = batright.objrotz

'*************** A.I. Flipper color On  **********************

  If AIon = 1 then  'here
                batleft.Visible = 0
          batright.Visible = 0
                batleftshadowglowred.visible = true
                batrightshadowglowred.visible = true
                batleftshadowglowgreen.visible = False
                batrightshadowglowgreen.visible = False
    batleftshadowglowblue.visible = False
    batrightshadowglowblue.visible = False
    batleftshadowglowwhite.visible = False
    batrightshadowglowwhite.visible = False
    batleftshadowglowgreen.visible = False
    batrightshadowglowgreen.visible = False
    batleftshadowglowyellow.visible = False
    batrightshadowglowyellow.visible = False
                glowbatleft.Visible = 1
                glowbatright.Visible = 1
                glowbatleft.Image = "glowbat redblack"
                glowbatright.Image = "glowbat redblackRight"
                Light3.State=1
                Light4.State=1
          Light5.State=1
                Light6.State=1
    Light5.Intensity = 0
                Light6.Intensity = 0
  If B2SOn Then
    DOF 130, DOFON
    Else
  End If
    Else
  If B2SOn Then
      DOF 130, DOFOFF
    Else
  End If

'****************** User Flipper Change ********************  Bad (Need to call instead)

  Select Case myPrefs_FlipperBatType
    Case 0
      batleft.Visible = 1
      batright.Visible = 1
      batleft.image = "_flipper_white_red"
      batright.image = "_flipper_white_red"
      batleftshadow.visible = true
      batrightshadow.visible = true
      glowbatleft.Visible = 0
      glowbatright.Visible = 0
      batleftshadowglowred.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light6.State=0
    Case 1
      batleft.Visible = 1
      batright.Visible = 1
      batleft.image = "_flipper_white_black2"
      batright.image = "_flipper_white_black2"
      batleftshadow.visible = true
      batrightshadow.visible = true
      glowbatleft.Visible = 0
      glowbatright.Visible = 0
      batleftshadowglowred.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light6.State=0
    Case 2
      batleft.Visible = 1
      batright.Visible = 1
      batleft.image = "_flipper_white_white2"
      batright.image = "_flipper_white_white2"
      batleftshadow.visible = true
      batrightshadow.visible = true
      glowbatleft.Visible = 0
      glowbatright.Visible = 0
      batleftshadowglowred.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light6.State=0
    Case 3
      batleft.Visible = 1
      batright.Visible = 1
      batleft.image = "_flipper_yellow_red"
      batright.image = "_flipper_yellow_red"
      batleftshadow.visible = true
      batrightshadow.visible = true
      glowbatleft.Visible = 0
      glowbatright.Visible = 0
      batleftshadowglowred.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light6.State=0
    Case 4
      batleft.Visible = 1
      batright.Visible = 1
      batleft.image = "_flipper_yellow_black"
      batright.image = "_flipper_yellow_black"
      batleftshadow.visible = true
      batrightshadow.visible = true
      glowbatleft.Visible = 0
      glowbatright.Visible = 0
      batleftshadowglowred.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light6.State=0
    Case 5
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowred.visible = true
      batrightshadowglowred.visible = true
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbat redblack"
      glowbatright.Image = "glowbat redblackRight"
      Light3.State=1
      Light4.State=1
      Light5.State=1
      Light6.State=1
      Light5.Intensity = 0
      Light6.Intensity = 0
    Case 6
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowgreen.visible = true
      batrightshadowglowgreen.visible = true
      batleftshadowglowred.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbat greenleft"
      glowbatright.Image = "glowbat greenRight"
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    Case 7
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowgreen.visible = true
      batrightshadowglowred.visible = true
      batleftshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batrightshadowglowgreen.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbat greenLeft"
      glowbatright.Image = "glowbat redblackRight"
      Light3.State=0
      Light4.State=1
      Light5.State=1
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    Case 8
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowred.visible = true
      batrightshadowglowgreen.visible = true
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowgreen.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbat redblack"
      glowbatright.Image = "glowbat greenRight"
      Light3.State=1
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=1
      Light6.Intensity = 0
    Case 9
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowred.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowblue.visible = True
      batrightshadowglowblue.visible = True
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbat BlueLeft"
      glowbatright.Image = "glowbat BlueRight"
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    Case 10
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowred.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = True
      batrightshadowglowwhite.visible = True
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbatWhiteLeft"
      glowbatright.Image = "glowbatWhiteRight"
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    Case 11
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowred.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowyellow.visible = True
      batrightshadowglowyellow.visible = True
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbat yellowblackLeft"
      glowbatright.Image = "glowbat yellowblackRight"
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    Case 12
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowred.visible = true
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batrightshadowglowred.visible = false
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowblue.visible = false
      batrightshadowglowblue.visible = True
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbat redblack"
      glowbatright.Image = "glowbat BlueRight"
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    Case 13
      batleft.Visible = 1
      batright.Visible = 1
      batleft.image = "_flipper_white_yellow"
      batright.image = "_flipper_white_yellow"
      batleftshadow.visible = true
      batrightshadow.visible = true
      glowbatleft.Visible = 0
      glowbatright.Visible = 0
      batleftshadowglowred.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    Case 14
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowred.visible = true
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batrightshadowglowred.visible = false
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowblue.visible = false
      batrightshadowglowblue.visible = True
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbat redblack"
      glowbatright.Image = "glowbat BlueRight"
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    Case 15
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowred.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = False
      batrightshadowglowwhite.visible = False
      batleftshadowglowyellow.visible = True
      batrightshadowglowyellow.visible = True
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbat yellowblackLeft"
      glowbatright.Image = "glowbat yellowblackRight"
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    Case 16
      batleft.Visible = 0
      batright.Visible = 0
      batleftshadowglowred.visible = False
      batleftshadowglowgreen.visible = False
      batrightshadowglowgreen.visible = False
      batrightshadowglowred.visible = False
      batleftshadowglowblue.visible = False
      batrightshadowglowblue.visible = False
      batleftshadowglowwhite.visible = True
      batrightshadowglowwhite.visible = True
      batleftshadowglowyellow.visible = False
      batrightshadowglowyellow.visible = False
      glowbatleft.Visible = 1
      glowbatright.Visible = 1
      glowbatleft.Image = "glowbatWhiteLeft"
      glowbatright.Image = "glowbatWhiteRight"
      Light3.State=0
      Light4.State=0
      Light5.State=0
      Light5.Intensity = 0
      Light6.State=0
      Light6.Intensity = 0
    End Select
  end If

'********************************************************

    if myPrefs_FactorySetting = 0 then
    Select Case myPrefs_Player1Setup
      Case 1
        Player1Light.Color=RGB(255,0,0)
        Player1Light.Colorfull=RGB(255,0,0)
        LightDTP1.Color=RGB(255,0,0)
        LightDTP1.Colorfull=RGB(255,0,0)
        BGR1up.Image = "1up"
        P1Light.Color=RGB(255,0,0)
        P1Light.Colorfull=RGB(255,0,0)
      Case 0
        Player1Light.Color=RGB(0,255,0)
        Player1Light.Colorfull=RGB(0,255,0)
        LightDTP1.Color=RGB(0,255,0)
        LightDTP1.Colorfull=RGB(0,255,0)
        BGR1up.Image = "1upG"
        P1Light.Color=RGB(0,255,0)
        P1Light.Colorfull=RGB(0,255,0)
    End Select
    end if

    if myPrefs_FactorySetting = 0 then
    Select Case  myPrefs_Player2Setup
      Case 1
        Player2Light.Color=RGB(255,0,0)
        Player2Light.Colorfull=RGB(255,0,0)
        LightDTP2.Color=RGB(255,0,0)
        LightDTP2.Colorfull=RGB(255,0,0)
        BGR2up.Image = "2up"
        P2Light.Color=RGB(255,0,0)
        P2Light.Colorfull=RGB(255,0,0)
      Case 0
        Player2Light.Color=RGB(0,255,0)
        Player2Light.Colorfull=RGB(0,255,0)
        LightDTP2.Color=RGB(0,255,0)
        LightDTP2.Colorfull=RGB(0,255,0)
        BGR2up.Image = "2upG"
        P2Light.Color=RGB(0,255,0)
        P2Light.Colorfull=RGB(0,255,0)
    End Select
    end if

    if myPrefs_FactorySetting = 0 then
    Select Case myPrefs_Player3Setup
      Case 1
        Player3Light.Color=RGB(255,0,0)
        Player3Light.Colorfull=RGB(255,0,0)
        LightDTP3.Color=RGB(255,0,0)
        LightDTP3.Colorfull=RGB(255,0,0)
        BGR3up.Image = "3up"
        P3Light.Color=RGB(255,0,0)
        P3Light.Colorfull=RGB(255,0,0)
      Case 0
        Player3Light.Color=RGB(0,255,0)
        Player3Light.Colorfull=RGB(0,255,0)
        LightDTP3.Color=RGB(0,255,0)
        LightDTP3.Colorfull=RGB(0,255,0)
        BGR3up.Image = "3upG"
        P3Light.Color=RGB(0,255,0)
        P3Light.Colorfull=RGB(0,255,0)
    End Select
    end if

    if myPrefs_FactorySetting = 0 then
    Select Case myPrefs_Player4Setup
      Case 1
        Player4Light.Color=RGB(255,0,0)
        Player4Light.Colorfull=RGB(255,0,0)
        LightDTP4.Color=RGB(255,0,0)
        LightDTP4.Colorfull=RGB(255,0,0)
        BGR4up.Image = "4up"
        P4Light.Color=RGB(255,0,0)
        P4Light.Colorfull=RGB(255,0,0)
      Case 0
        Player4Light.Color=RGB(0,255,0)
        Player4Light.Colorfull=RGB(0,255,0)
        LightDTP4.Color=RGB(0,255,0)
        LightDTP4.Colorfull=RGB(0,255,0)
        BGR4up.Image = "4upG"
        P4Light.Color=RGB(0,255,0)
        P4Light.Colorfull=RGB(0,255,0)
    End Select
    end if

  if BGLtilt.State=1 And AIOn = 1 Then
    TiltTimer.Enabled = 1
  else
    TiltTimer.Enabled = 0
  end If

   If  AIOn = 1 And myPrefs_AInudge= 0 Then
    PostLightOnOff = 1
    TimerMovePost.Enabled = True '(Easy)
    REDPOST1.Collidable = True
    REDPOST.Collidable = False
    Else

    If myPrefs_AInudge= 0 Then
    PostLightOnOff = 1
    TimerMovePost2.Enabled = True'(Default)
    REDPOST1.Collidable = False
    REDPOST.Collidable = True
    End If
  End If

  if CBool(Controller.Lamp(56)) = True then
    if Solid = 1 Then
      CreditLight3.State = 2
      CreditLight3Timer.enabled = 1
    end if
  end if

  if CBool(Controller.Lamp(56)) = False then
    CreditLight3.State = 0
    Solid = 1
  end if

  if myPrefs_SoundEffects = 1 then:   Hal9000Eye.Image= "hal-9000":HalLight2.State = 2:HalLight2.Color=RGB(255,28,28):LF.ColorFull=RGB(255,28,28) end If
  if myPrefs_SoundEffects = 0 then: Hal9000Eye.Image= "hal-9000B":HalLight2.State = 2:HalLight2.Color=RGB(0,0,255):LF.ColorFull=RGB(0,0,255) end If

End Sub




'Sub AIFlipperOn 'here
' If AIon =1 or AIon = 0then
'                batleft.Visible = 0
'         batright.Visible = 0
'                batleftshadowglowred.visible = true
'                batrightshadowglowred.visible = true
'                batleftshadowglowgreen.visible = False
'                batrightshadowglowgreen.visible = False
'   batleftshadowglowblue.visible = False
'   batrightshadowglowblue.visible = False
'   batleftshadowglowwhite.visible = False
'   batrightshadowglowwhite.visible = False
'   batleftshadowglowgreen.visible = False
'   batrightshadowglowgreen.visible = False
'   batleftshadowglowyellow.visible = False
'   batrightshadowglowyellow.visible = False
'                glowbatleft.Visible = 1
'                glowbatright.Visible = 1
'                glowbatleft.Image = "glowbat redblack"
'                glowbatright.Image = "glowbat redblackRight"
'                Light3.State=1
'                Light4.State=1
'         Light5.State=1
'                Light6.State=1
'   Light5.Intensity = 0
'                Light6.Intensity = 0
' If B2SOn Then
'   DOF 130, DOFON
'   Else
' End If
'   Else
' If B2SOn Then
'     DOF 130, DOFOFF
'   Else
' End If
'end if
'End Sub


Sub CreditLight3Timer_Timer
  CreditLight3.State = 1
  CreditLight3Timer.enabled = 0
  Solid = 0
End Sub

Sub Halspeak2Timer_Timer
  playsound"_Concerned"
End Sub

Dim Solid
solid = 1

'******************** Flipper Flashers *********************
Dim GlowBat

Sub   myPowerFlash
  Select case myPrefs_FlipperBatType
    Case 5
      Call GlowBatOn
      GlowBat = 5
      myPrefs_FlipperBatType = 15
    Case  6
      Call GlowBatOn
      GlowBat = 6
      myPrefs_FlipperBatType = 16
    Case 7
      Call GlowBatOn
      GlowBat = 7
      myPrefs_FlipperBatType = 15
    Case 8
      Call GlowBatOn
      GlowBat = 8
      myPrefs_FlipperBatType = 15
    Case 9
      Call GlowBatOn
      GlowBat = 9
      myPrefs_FlipperBatType = 16
    Case 10
      Call GlowBatOn
      GlowBat = 10
      myPrefs_FlipperBatType = 16
    Case 11
      Call GlowBatOn
      GlowBat = 11
      myPrefs_FlipperBatType = 16
    Case 12
      Call GlowBatOn
      GlowBat = 12
      myPrefs_FlipperBatType = 16
  end select
End Sub

Sub GlowBatOn
  FlipperFlasher.enabled = 1
  LightRed.Duration 2, 80, 0
  LightBlue.Duration 2, 80, 0
  FlasherPower001.visible = True
  FlasherFire001.visible = True
  FlasherFlipper.visible = True
End Sub

Sub FlipperFlasher_Timer
  Call GlowBatOff
  Select Case Glowbat
    Case 5
      myPrefs_FlipperBatType=5
    Case 6
      myPrefs_FlipperBatType=6
    Case 7
      myPrefs_FlipperBatType=7
    Case 8
      myPrefs_FlipperBatType=8
    Case 9
      myPrefs_FlipperBatType=9
    Case 10
      myPrefs_FlipperBatType=10
    Case 11
      myPrefs_FlipperBatType=11
    Case 12
      myPrefs_FlipperBatType=12
  end select
End Sub

Sub GlowBatOff
  FlipperFlasher.enabled=false
  FlasherPower001.visible = False
  FlasherFire001.visible = False
  FlasherFlipper.visible = False
End Sub

'       ____  __  __________   ___   _  ______  ___  _______________________  ____ ________  __
'***** / __ \/ / / /_  __/ /  / _ | / |/ / __/ / _ \/  _/ __/ __/  _/ ___/ / / / //_  __/\ \/ /*****
'*****/ /_/ / /_/ / / / / /__/ __ |/    / _/  / // // // _// _/_/ // /__/ /_/ / /__/ /    \  / *****
'*****\____/\____/ /_/ /____/_/ |_/_/|_/___/ /____/___/_/ /_/ /___/\___/\____/____/_/     /_/  *****

Dim PostLightOnOff

  Sub TiltTimer_Timer
    myPrefs_AInudge= 0
    TiltTimer.Enabled = 0
  End Sub

  Sub TimerMovePost_Timer
      REDPOST1.TransX = REDPOST1.TransX-.1
    If  REDPOST1.TransX < 2  Then
      REDPOST1.TransX = 2
      TimerMovePost.Enabled = False
    End If
  End Sub

  Sub TimerMovePost2_Timer
      REDPOST1.TransX = REDPOST1.TransX+.1
    If  REDPOST1.TransX > 12  Then
      REDPOST1.TransX = 12
      TimerMovePost2.Enabled = False
    End If
  End Sub

If myPrefs_AInudge > 0 Then
  Select Case myPrefs_OutlaneDifficulty
    Case 0 'Default(0)
          REDPOST1.Visible = False
          REDPOST1.Collidable = False
          REDPOST.Visible = True
          REDPOST.Collidable = True
    Case 1 'Easy(1)
          REDPOST1.Visible = True
          REDPOST1.Collidable = True
          REDPOST.Visible = False
          REDPOST.Collidable = False
  End Select
End If

  Select Case myPrefs_Hal9000EyeOnOff
    Case 0 'Off
          Hal9000Eye.Visible = False
          Hal9000Ring.Visible = False
    HalLight2.Intensity = 0
    StrobeLight1.Intensity = 0
    Case 1 'On
          Hal9000Eye.Visible = True
          Hal9000Ring.Visible = True
    HalLight2.Intensity = 90
    StrobeLight1.Intensity = 10
  End Select

'      _________   ___  ___________________
'*****/_  __/ _ | / _ \/ ___/ __/_  __/ __/*****              Y+
'***** / / / __ |/ , _/ (_ / _/  / / _\ \  *****              ^
'*****/_/ /_/ |_/_/|_|\___/___/ /_/ /___/  *****                 >X+

' ************************************************
' HitTarget animation
' ************************************************


Sub Target1T_Hit
  Target1_10.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch 17, 0, "":    'Target1
  If Activeball.velY > 16 Then
      'Msgbox Activeball.velY
      Activeball.velZ = 17
  End If
End Sub

Sub ComboTrigger1_hit
  if myPrefs_FlipperStobes > 2 and AIOff then
    call myPowerFlash
  end if
  Target1_10.playanim 0,.5
  Target2_10.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch 17, 0, ""
  vpmTimer.PulseSwitch 18, 0, ""   'Target 1/2
  'MsgBox "Combo!!!"
End Sub

Sub Target2T_Hit
  Target2_10.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch 18, 0, "":     'Target2
  If Activeball.velY > 16 Then
    'Msgbox Activeball.velY
    Activeball.velZ = 17
        Select Case Int(Rnd()*5)'  Loose Bulb
            Case 0
                GI_9.State=2
                GIDimmer.Enabled=1
            Case 1
                GI_9.State=1
            Case 2
                GI_9.State=1
            Case 3
                GI_9.State=1
            Case 4
                GI_9.State=1
        End Select
  End If
End Sub

Sub ComboTrigger2_hit
  if myPrefs_FlipperStobes > 2  and AIOff then
    call myPowerFlash
  end if
  Target2_10.playanim 0,.5
  Target3_10.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch 18, 0, ""
  vpmTimer.PulseSwitch 19, 0, ""   'Target 2/3
  'MsgBox "Combo!!!"
End Sub

Sub Target3T_Hit
  Target3_10.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch 19, 0, "":     'Target3
  If Activeball.velY > 16 Then
    'Msgbox Activeball.velY
    Activeball.velZ = 17
  End If
End Sub

Sub Target4T_Hit
  Target4_10.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch 21, 0, "":     'Target4
  If Activeball.velY > 16 Then
    'Msgbox Activeball.velY
       Activeball.velZ = 17
  End If
End Sub

Sub ComboTrigger3_hit
  if myPrefs_FlipperStobes > 2  and AIOff then
    call myPowerFlash
  end if
  Target4_10.playanim 0,.5
  Target5_10.playanim 0,.5
  PlaysoundAtVol "target", Activeball, 1
  vpmTimer.PulseSwitch 21, 0, ""
  vpmTimer.PulseSwitch 22, 0, ""   'Target4/5
  'MsgBox "Combo!!!"
End Sub

Sub Target5T_Hit
  Target5_10.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch 22, 0, "":     'Target5
  If Activeball.velY > 16 Then
    'Msgbox Activeball.velY
       Activeball.velZ = 17
        Select Case Int(Rnd()*5)
            Case 0
                GI_9.State=2
                GIDimmer.Enabled=1
            Case 1
                GI_9.State=1
            Case 2
                GI_9.State=1
            Case 3
                GI_9.State=1
            Case 4
                GI_9.State=1
        End Select
    End If
End Sub

Sub ComboTrigger4_hit
  if myPrefs_FlipperStobes > 2 and AIOff then
    call myPowerFlash
  end if
  Target5_10.playanim 0,.5
  Target6_10.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch 22, 0, ""
  vpmTimer.PulseSwitch 23, 0, ""   'Target 5/6
  'MsgBox "Combo!!!"
End Sub

Sub Target6T_Hit
  Target6_10.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSwitch 23, 0, "":     'Target6
' Target6_10.playanim 0,.1
  If Activeball.velY > 16 Then
    'Msgbox Activeball.velY
    Activeball.velZ = 17
  End If
End Sub

Sub PTTarget_6_Hit
  PTTarget_6.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSw cTopPOWERTargetSW
End Sub

Sub ComboTriggerP12_Hit
  if myPrefs_FlipperStobes > 2 and AIOff then
    call myPowerFlash
  end if
  PTTarget_6.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSw cTopPOWERTargetSW
  PMTarget_6.playanim 0,.6
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSw cMiddlePOWERTargetSW
End Sub

Sub ComboTriggerP23_Hit
  if myPrefs_FlipperStobes > 2 and AIOff then
    call myPowerFlash
  end if
  PMTarget_6.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSw cMiddlePOWERTargetSW
  PBTarget_6.playanim 0,.6
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSw cBottomPOWERTargetSW
End Sub

Sub PMTarget_6_Hit
  if ShotonLight13.state = 1 and myPrefs_FlipperStobes > 2 and AIOff Then
    call myPowerFlash
  end if
  PMTarget_6.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSw cMiddlePOWERTargetSW
End Sub

Sub PBTarget_6_Hit
  PBTarget_6.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSw cBottomPOWERTargetSW
End Sub

Sub TopTarget_6_Hit
  TopTarget_6.playanim 0,.5
  PlaysoundAtVol "target", ActiveBall, 1
  vpmTimer.PulseSw  cTopCenterTargetSW
End Sub


'********************************
'***** 10 point switches ***** (Standup Switches) (8 total) (StandupTarget1-8)
'********************************

Sub StandupTarget1_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "", ActiveBall, 1 :End Sub
Sub StandupTarget1_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub

Sub StandupTarget2_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "", ActiveBall, 1 :End Sub
Sub StandupTarget2_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub

Sub StandupTarget3_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "", ActiveBall, 1:End Sub
Sub StandupTarget3_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub

Sub StandupTarget4_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "", ActiveBall, 1 :End Sub
Sub StandupTarget4_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub

Sub StandupTarget5_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "", ActiveBall, 1 :End Sub
Sub StandupTarget5_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub

Sub StandupTarget6_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "", ActiveBall, 1 :End Sub
Sub StandupTarget6_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub

Sub StandupTarget7_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "", ActiveBall, 1 :End Sub
Sub StandupTarget7_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub

Sub StandupTarget8_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "", ActiveBall, 1 :End Sub
Sub StandupTarget8_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub

'***********************
'***** Target Mods *****
'***********************

Select Case myPrefs_Targets
  Case 0
    Target1_10.Image = "targetmidleft1_red"
    Target2_10.Image = "targetmidleft2_red"
    Target3_10.Image = "targetmidleft3_red"
    Target4_10.Image = "targetmidright1_red"
    Target5_10.Image = "targetmidright2_red"
    Target6_10.Image = "targetmidright3_red"
  Case 1
    Target1_10.Image = "targetmidleft1_fp"
    Target2_10.Image = "targetmidleft2_fp"
    Target3_10.Image = "targetmidleft3_fp"
    Target4_10.Image = "targetmidright1_fp"
    Target5_10.Image = "targetmidright2_fp"
    Target6_10.Image = "targetmidright3_fp"
End Select

'         _____  _______________  __  _____________________  _  __  ________   ___  ___  ____
'*****   /  _/ |/ / __/_  __/ _ \/ / / / ___/_  __/  _/ __ \/ |/ / / ___/ _ | / _ \/ _ \/ __/*****
'***** _/ //    /\ \  / / / , _/ /_/ / /__  / / _/ // /_/ /    / / /__/ __ |/ , _/ // /\ \   *****
'*****/___/_/|_/___/ /_/ /_/|_|\____/\___/ /_/ /___/\____/_/|_/  \___/_/ |_/_/|_/____/___/   *****

  Select Case myPrefs_InstructionCardType
    Case 0
      CardLFactory.visible = 1
      CardRFactory.visible = 1
    Case 1
      CardLMod.visible = 1
      CardRMod.visible = 1
    Case 2
      CardLTHX.visible = 1
      CardRTHX.visible = 1
   Case 3
      CardAIRight.Visible = 1
      CardAILeft.Visible = 1
    End Select


'**************** Logic Panel Lights *****************

Set Lights(14) = B1
Set Lights(15) = B2
Set Lights(16) = B3
Set Lights(17) = B4
Set Lights(18) = B5
Set Lights(19) = B6
Set Lights(20) = B7
Set Lights(21) = B8
Set Lights(22) = B9
Set Lights(2) = B10
Set Lights(24) = B11
Set Lights(25) = B12
Set Lights(36) = B13
Set Lights(37) = B14
Set Lights(38) = B15
Set Lights(39) = B16
Set Lights(12) = B17
Set Lights(13) = B18
Set Lights(1) = B19

'        ___  ____  ___    ___  __  ____  ______  _______  ____
'*****  / _ \/ __ \/ _ \  / _ )/ / / /  |/  / _ \/ __/ _ \/ __/*****
'***** / ___/ /_/ / ___/ / _  / /_/ / /|_/ / ___/ _// , _/\ \  *****
'*****/_/   \____/_/    /____/\____/_/  /_/_/  /___/_/|_/___/  *****

Dim xx

Set Lights(44)=LBumperTopLeft
Set Lights(45)=LBumperTopRight
Set Lights(46)=LBumperBottomRight
Set Lights(47)=LBumperBottomLeft

Sub B1T1_Hit:BumpSkirt1.ROTX=6:                                        End Sub
Sub B1T2_Hit:BumpSkirt1.ROTX=6:  BumpSkirt1.ROTY=6:  End Sub
Sub B1T3_Hit:BumpSkirt1.ROTY=6:                                        End Sub
Sub B1T4_Hit:BumpSkirt1.ROTX=-6:BumpSkirt1.ROTY=6:  End Sub
Sub B1T5_Hit:BumpSkirt1.ROTY=-6:                                      End Sub
Sub B1T6_Hit:BumpSkirt1.ROTX=-6:BumpSkirt1.ROTY=-6:End Sub
Sub B1T7_Hit:BumpSkirt1.ROTY=-6:                                      End Sub
Sub B1T8_Hit:BumpSkirt1.ROTX=6:  BumpSkirt1.ROTY=-6:End Sub

Sub B2T1_Hit:BumpSkirt2.ROTX=6:                                        End Sub
Sub B2T2_Hit:BumpSkirt2.ROTX=6:  BumpSkirt2.ROTY=6:  End Sub
Sub B2T3_Hit:BumpSkirt2.ROTY=6:                                        End Sub
Sub B2T4_Hit:BumpSkirt2.ROTX=-6:BumpSkirt2.ROTY=6:  End Sub
Sub B2T5_Hit:BumpSkirt2.ROTY=-6:                                      End Sub
Sub B2T6_Hit:BumpSkirt2.ROTX=-6:BumpSkirt2.ROTY=-6:End Sub
Sub B2T7_Hit:BumpSkirt2.ROTY=-6:                                      End Sub
Sub B2T8_Hit:BumpSkirt2.ROTX=6:  BumpSkirt2.ROTY=-6:End Sub

Sub B3T1_Hit:BumpSkirt3.ROTX=6:                                        End Sub
Sub B3T2_Hit:BumpSkirt3.ROTX=6:  BumpSkirt3.ROTY=6:  End Sub
Sub B3T3_Hit:BumpSkirt3.ROTY=6:                                        End Sub
Sub B3T4_Hit:BumpSkirt3.ROTX=-6:BumpSkirt3.ROTY=6:  End Sub
Sub B3T5_Hit:BumpSkirt3.ROTY=-6:                                      End Sub
Sub B3T6_Hit:BumpSkirt3.ROTX=-6:BumpSkirt3.ROTY=-6:End Sub
Sub B3T7_Hit:BumpSkirt3.ROTY=-6:                                      End Sub
Sub B3T8_Hit:BumpSkirt3.ROTX=6:  BumpSkirt3.ROTY=-6:End Sub

Sub B4T1_Hit:BumpSkirt4.ROTX=6:                                        End Sub
Sub B4T2_Hit:BumpSkirt4.ROTX=6:  BumpSkirt4.ROTY=6:  End Sub
Sub B4T3_Hit:BumpSkirt4.ROTY=6:                                        End Sub
Sub B4T4_Hit:BumpSkirt4.ROTX=-6:BumpSkirt4.ROTY=6:  End Sub
Sub B4T5_Hit:BumpSkirt4.ROTY=-6:                                      End Sub
Sub B4T6_Hit:BumpSkirt4.ROTX=-6:BumpSkirt4.ROTY=-6:End Sub
Sub B4T7_Hit:BumpSkirt4.ROTY=-6:                                      End Sub
Sub B4T8_Hit:BumpSkirt4.ROTX=6:  BumpSkirt4.ROTY=-6:End Sub

'*********************************************************************

Sub B1T1_UnHit:BumpSkirt1.ROTX=0:BumpSkirt1.ROTY=0:End Sub
Sub B1T2_UnHit:BumpSkirt1.ROTX=0:BumpSkirt1.ROTY=0:End Sub
Sub B1T3_UnHit:BumpSkirt1.ROTX=0:BumpSkirt1.ROTY=0:End Sub
Sub B1T4_UnHit:BumpSkirt1.ROTX=0:BumpSkirt1.ROTY=0:End Sub
Sub B1T5_UnHit:BumpSkirt1.ROTX=0:BumpSkirt1.ROTY=0:End Sub
Sub B1T6_UnHit:BumpSkirt1.ROTX=0:BumpSkirt1.ROTY=0:End Sub
Sub B1T7_UnHit:BumpSkirt1.ROTX=0:BumpSkirt1.ROTY=0:End Sub
Sub B1T8_UnHit:BumpSkirt1.ROTX=0:BumpSkirt1.ROTY=0:End Sub

Sub B2T1_UnHit:BumpSkirt2.ROTX=0:BumpSkirt2.ROTY=0:End Sub
Sub B2T2_UnHit:BumpSkirt2.ROTX=0:BumpSkirt2.ROTY=0:End Sub
Sub B2T3_UnHit:BumpSkirt2.ROTX=0:BumpSkirt2.ROTY=0:End Sub
Sub B2T4_UnHit:BumpSkirt2.ROTX=0:BumpSkirt2.ROTY=0:End Sub
Sub B2T5_UnHit:BumpSkirt2.ROTX=0:BumpSkirt2.ROTY=0:End Sub
Sub B2T6_UnHit:BumpSkirt2.ROTX=0:BumpSkirt2.ROTY=0:End Sub
Sub B2T7_UnHit:BumpSkirt2.ROTX=0:BumpSkirt2.ROTY=0:End Sub
Sub B2T8_UnHit:BumpSkirt2.ROTX=0:BumpSkirt2.ROTY=0:End Sub

Sub B3T1_UnHit:BumpSkirt3.ROTX=0:BumpSkirt3.ROTY=0:End Sub
Sub B3T2_UnHit:BumpSkirt3.ROTX=0:BumpSkirt3.ROTY=0:End Sub
Sub B3T3_UnHit:BumpSkirt3.ROTX=0:BumpSkirt3.ROTY=0:End Sub
Sub B3T4_UnHit:BumpSkirt3.ROTX=0:BumpSkirt3.ROTY=0:End Sub
Sub B3T5_UnHit:BumpSkirt3.ROTX=0:BumpSkirt3.ROTY=0:End Sub
Sub B3T6_UnHit:BumpSkirt3.ROTX=0:BumpSkirt3.ROTY=0:End Sub
Sub B3T7_UnHit:BumpSkirt3.ROTX=0:BumpSkirt3.ROTY=0:End Sub
Sub B3T8_UnHit:BumpSkirt3.ROTX=0:BumpSkirt3.ROTY=0:End Sub

Sub B4T1_UnHit:BumpSkirt4.ROTX=0:BumpSkirt4.ROTY=0:End Sub
Sub B4T2_UnHit:BumpSkirt4.ROTX=0:BumpSkirt4.ROTY=0:End Sub
Sub B4T3_UnHit:BumpSkirt4.ROTX=0:BumpSkirt4.ROTY=0:End Sub
Sub B4T4_UnHit:BumpSkirt4.ROTX=0:BumpSkirt4.ROTY=0:End Sub
Sub B4T5_UnHit:BumpSkirt4.ROTX=0:BumpSkirt4.ROTY=0:End Sub
Sub B4T6_UnHit:BumpSkirt4.ROTX=0:BumpSkirt4.ROTY=0:End Sub
Sub B4T7_UnHit:BumpSkirt4.ROTX=0:BumpSkirt4.ROTY=0:End Sub
Sub B4T8_UnHit:BumpSkirt4.ROTX=0:BumpSkirt4.ROTY=0:End Sub

Sub Bumper1_hit
vpmTimer.PulseSw cTopLeftBumperSW
PlaysoundAtVol "Fx_Bumper2", ActiveBall, 1
'BR1Animation
  'SkirtTimer.Enabled=1
  If BGLmatch.State = 0 Then
    BR1_7.playanim 0,.6
  End If
  If BGLtilt.State = 1 Then
    BR1_7.playanim 0,.0
  End If
' if myPrefs_FlipperStobes > 0 and isMultiball = 1 and AIOff then
'   Call myPowerFlash
' end if
  if myPrefs_FlipperStobes > 1 and  LBumperTopLeft.State=1 and AIOff then
    Call myPowerFlash
  end if
End Sub

Sub Bumper2_hit
vpmTimer.PulseSw cTopRightBumperSW
PlaysoundAtVol "Fx_Bumper3", ActiveBall, 1
  If BGLmatch.State = 0 Then
    BR2_7.playanim 0,.6
  End If
  If BGLtilt.State = 1 Then
    BR2_7.playanim 0,.0
  End If
' if myPrefs_FlipperStobes > 0 and isMultiball = 1 and AIOff then
'   Call myPowerFlash
' end if
  if myPrefs_FlipperStobes > 1 and  LBumperTopRight.State=1 and AIOff then
    Call myPowerFlash
  end if
End Sub

Sub Bumper3_hit
vpmTimer.PulseSw cBottomRightBumperSW
PlaysoundAtVol "Fx_Bumper4", ActiveBall, 1
  If BGLmatch.State = 0 Then
    BR3_7.playanim 0,.6
  End If
  If BGLtilt.State = 1 Then
    BR3_7.playanim 0,.0
  End If
' if myPrefs_FlipperStobes > 0 and isMultiball = 1 and AIOff then
'   Call myPowerFlash
' end if
  if myPrefs_FlipperStobes > 1 and  LBumperBottomRight.State=1 and AIOff then
    Call myPowerFlash
  end if
End Sub

Sub Bumper4_hit
vpmTimer.PulseSw cBottomLeftBumperSW
PlaysoundAtVol "Fx_Bumper1", ActiveBall, 1
  If BGLmatch.State = 0 Then
    BR4_7.playanim 0,.6
  End If
  If BGLtilt.State = 1 Then
    BR4_7.playanim 0,.0
  End If
' if myPrefs_FlipperStobes > 0 and isMultiball = 1 and AIOff then
'   Call myPowerFlash
' end if
  if myPrefs_FlipperStobes > 1 and  LBumperBottomLeft.State=1 and AIOff then
    Call myPowerFlash
  end if
End Sub

'        ______   ___   ______ _________  ____
'*****  / __/ /  / _ | / __/ // / __/ _ \/ __/*****
'***** / _// /__/ __ |_\ \/ _  / _// , _/\ \  *****
'*****/_/ /____/_/ |_/___/_//_/___/_/|_/___/  *****


Sub Flashers(enabled)
  If enabled Then
    FlasherLightFire.State = 1
    FlasherLightPower.State = 1
    FlasherFire.Visible = True
    FlasherPower.Visible = True
    FlasherBig.Visible = True
    LFire1.State = 1
    LFire2.State = 1
    LPower1.State = 1
    LPower2.State = 1
    if gamestarted = 0 and myPrefs_AttractMode = 1 and myPrefs_FlipperStobes > 0 Then
      Call myPowerFlash
    end if
    if gamestarted = 1 and myPrefs_FlipperStobes > 0 and AIOff then
      Call myPowerFlash
    end if
  Else
    FlasherLightFire.State = 0
    FlasherLightPower.State = 0
    FlasherFire.Visible = False
    FlasherPower.Visible = False
    FlasherBig.Visible = False
    LFire1.State = 0
    LFire2.State = 0
    LPower1.State =0
    LPower2.State = 0
  End If
End Sub

'***************************************************
'Target Lights
Set Lights(26)=LT1
Set Lights(27)=LT2
Set Lights(28)=LT3
Set Lights(29)=LT4
Set Lights(30)=LT5
Set Lights(31)=LT6

'---------------------------------------------------
' Set up consts with easy names for solenoid numbers
Const cBallRelease = 1
Const cLeftEjectHole = 4
Const cRightEjectHole = 5
Const cUpperEjectHole = 6
Const cBallSaveKick = 7
Const cBallRampThrower = 8
Const cCreditKnocker = 14
Const cFlashLamps = 15
Const cTopLeftBumper = 17
Const cBottomLeftBumper = 18
Const cTopRightBumper = 19
Const cBottomRightBumper = 20
Const cRightSlingShot = 21
Const cLeftSlingShot = 22
Const MaxLut = 4
'---------------------------------------------------

'---------------------------------------------------
' Setup Consts with easy names for switch numbers
Const cOutHoleSW = 9
Const cLeftOutsideRolloverSW = 10
Const cLeftInsideRolloverSW = 11
Const cLeftKickerSW = 12
Const cLeftEjectHoleSW = 13
Const cUpperMiddleLeftStandupSW = 14
Const cSpinnerSW = 15
Const cTopLeftStandupSW = 16
'Const cTarget1SW = 17
'Const cTarget2SW = 18
'Const cTarget3SW = 19
'Const cTarget4SW = 21
'Const cTarget5SW = 22
'Const cTarget6SW = 23
Const cBottomLeftBumperSW = 25
Const cTopLeftBumperSW = 26
Const cTopRightBumperSW = 27
Const cBottomRightBumperSW = 28
Const cTopCenterTargetSW = 29
Const cRightEjectHoleSW = 30
Const cUpperTopRightStandupSW = 31
Const cFRolloverSW = 32
Const cIRolloverSW = 33
Const cRRolloverSW = 34
Const cERolloverSW = 35
Const cUpperRightEjectHoleSW = 36
Const cLowerTopRightStnadupSW = 37
Const cMIddleRightStandupSW = 38
Const cTopPOWERTargetSW = 39
Const cMiddlePOWERTargetSW = 40
Const cBottomPOWERTargetSW = 41
Const cRightKickerSW = 42
Const cRightInsideRolloverSW = 43
Const cRightOutsideRolloverSW = 44
Const cFlipperRightSW = 45
Const cBallShooterSW = 46
Const cPlayfieldTiltSW = 47
Const cLowerRightStandupSW = 48
Const cCenterMiddleLeftStandupSW = 49
Const cLowerMiddleLeftStuandupSW = 50
Const cLeftBallRampSW = 51
Const cLeftEjectRolloverSW = 53
Const cRightEjectRolloverSW = 54
Const cRightBallRampSW = 57
Const cCenterBallRampSW = 58
'--------------------------------------------------

'Solenoids Setup
SolCallback(cBallRelease) = "trTrough.SolIn"
SolCallback(cLeftEjectHole) = "LeftEjectHole"
SolCallback(cRightEjectHole) = "RightEjectHole"
SolCallback(cUpperEjectHole) = "UpperEjectHole"
SolCallback(cBallSaveKick) = "BallSaveKick"
SolCallback(cBallRampThrower) = "trTrough.SolOut"
SolCallback(cFlashLamps) = "Flashers"
SolCallback(cCreditKnocker) = "CreditKnocker"
SolCallback(cRightSlingShot) = "sRightSlingShot"
SolCallback(cLeftSlingShot) = "sLeftSlingShot"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
If dtmod=1 then
  SolCallback(2) = "LBankReset"
  SolCallback(3) = "RBankReset"
End If

'------------------------------------------------------------------
' Setup Solenoid Subs
Sub LeftEjectHole(enabled)
  if enabled Then
    KLeftEjectHole.Kick 180, 15
    PlaysoundAtVol "Popper_ball", KLeftEjectHole, 1
    controller.Switch(cLeftEjectHoleSW)=0
  end If
End Sub

Sub RightEjectHole(enabled)
  if enabled Then
    KRightEjectHole.Kick 110,25 'AXS  (180.15)
    Controller.Switch(cRightEjectHoleSW)=0
    PlaysoundAtVol "Popper_ball", KRightEjectHole, 1
  end If
End Sub

Sub UpperEjectHole(enabled)
  if enabled Then
    KUpperEjectHole.Kick -90,15
    Controller.Switch(cUpperRightEjectHoleSW)=0
  PlaysoundAtVol "Popper_ball", KUpperEjectHole, 1
  end if
End Sub


Sub CreditKnocker(enabled)
  if enabled then
  if myPrefs_FlipperStobes > 1 and gamestarted = 1 and AIOff then
    Call myPowerFlash
  end if
  Playsound "knocker"
  'DOF 14,2
  end if
end Sub

'        ______   _____  _______  ______ ______  __________
'*****  / __/ /  /  _/ |/ / ___/ / __/ // / __ \/_  __/ __/*****
'***** _\ \/ /___/ //    / (_ / _\ \/ _  / /_/ / / / _\ \  *****
'*****/___/____/___/_/|_/\___/ /___/_//_/\____/ /_/ /___/  *****

'***************************************************************
'Rstep and Lstep  are the variables that increment the animation
'***************************************************************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot()
  vpmTimer.PulseSW cRightKickerSW
End Sub

Sub LeftSlingShot_Slingshot()
  vpmTimer.PulseSW cLeftKickerSW
End Sub

Sub sRightSlingShot(enabled)
  If enabled Then
  if myPrefs_FlipperStobes > 0 and AIOff then
    Call myPowerFlash
  end if
    PlaySoundAtVol "slingshotRight", ActiveBall, 1
    'RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  End If
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling1.Visible = 0:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub sLeftSlingShot(enabled)
  If enabled Then
  if myPrefs_FlipperStobes > 0 and AIOff then
    Call myPowerFlash
  end if
    PlaySoundAtVol "slingshotLeft", ActiveBall, 1
    'LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub SlingHopL_Hit
     'Msgbox "left Hit"
     If Activeball.velX > 9 Then
    'Msgbox Activeball.velx
    SlingHopLTimer.Enabled = 1
    Activeball.velZ = 12
  End If
  Select Case Int(Rnd()*5)
    Case 0
    Solid = 1
    Case 1
    Case 2
    Case 3
    Case 4
  End Select
End Sub

Sub SlingHopLTimer_Timer
     PlaysoundAtVol "_ball_bounce", SlingHopL, 1
     SlingHopLTimer.Enabled = 0
End Sub

Sub SlingHopR_Hit
     'Msgbox "Right hit"
  If Activeball.velX > 7 Then
    'Msgbox Activeball.velx
    SlingHopRTimer.Enabled = 1
    Activeball.velZ = 12
  End If
        Select Case Int(Rnd()*5)
            Case 0
                GI_1.State=2
                GIDimmer.Enabled=1
            Case 1
                GI_1.State=1
            Case 2
                GI_1.State=1
            Case 3
                GI_1.State=1
            Case 4
                GI_1.State=1
        End Select
End Sub

Sub SlingHopRTimer_Timer
     PlaysoundAtVol "_ball_bounce", SlingHopR, 1
     SlingHopRTimer.Enabled = 0
End Sub

Sub GIDimmer_timer
       GI_1.State=1
       GI_9.State=1
End sub

dim MBon
MBon=0

' Flipper Solenoid Handlers ********************************************************

sub SolRFlipper(enabled)
  if enabled Then
    PlaySoundAtVol "Fx_FlipperUp", FlipperRight, 1
    FlipperRight.RotateToEnd
                     Light4.Intensity = 1000
                     Light5.Intensity = 500
  Else
    PlaySoundAtVol "Fx_FlipperDown", FlipperRight, 1
    FlipperRight.RotateToStart
                     Light4.Intensity = 1000
                     Light5.Intensity = 0
  end If
end Sub

sub SolLFlipper(enabled)
  if enabled Then
    'PlaySound "Fx_FlipperUp", 0, 1, -0.02, 0.05
    FlipperLeft.RotateToEnd
    Light3.Intensity = 500
                Light6.Intensity = 500
    if light2.intensity=0 or AIOn then
      PlaySoundAtVol "Fx_FlipperUp", FlipperLeft, 1
    end if
  Select Case Int(Rnd()*4)
    Case 0
    Solid = 1
    Case 1
    Case 2
    Case 3
  End Select
  Else
    PlaySoundAtVol "Fx_FlipperDown", FlipperLeft, 1
    FlipperLeft.RotateToStart
    Light3.Intensity = 1000
                     Light6.Intensity = 0
  end If
end Sub

' Handle eject hole hit and unhit events and set corresponding switches *************

Sub KLeftEjectHole_Hit()
  if ShotonLight1.State = 1 and myPrefs_FlipperStobes > 0 and AIOff then
    Call myPowerFlash
  end if
  controller.Switch(cLeftEjectHoleSW)=1
  MBon=MBon-1
End Sub

Sub KLeftEjectHole_UnHit()
  controller.Switch(cLeftEjectHoleSW)=0
  MBon=MBon+1
End Sub

Sub KRightEjectHole_Hit()
  if ShotonLight9.State = 1 and myPrefs_FlipperStobes > 0 and AIOff then
    Call myPowerFlash
  end if
  MBon=MBon-1
  controller.Switch(cRightEjectHoleSW)=1
End Sub

Sub KRightEjectHole_UnHit()
  controller.Switch(cRightEjectHoleSW)=0
  MBon=MBon+1
End Sub

Sub KUpperEjectHole_Hit()
  if ShotonLight2.State = 1 and myPrefs_FlipperStobes > 0 and AIOff then
    Call myPowerFlash
  end if
  MBon=MBon-1
  controller.Switch(cUpperRightEjectHoleSW)=1
End Sub

Sub KUpperEjectHole_Unhit()
  controller.Switch(cUpperRightEjectHoleSW)=0
  MBon=MBon+1
End Sub

Sub BallSaveKick(enabled)
  if enabled Then
    Playsound "slingshotleft", 0, .67, -0.05, 0.05
  End If
End Sub


'***************************************************************************************
'Initialize Table

Sub Table1_Init()
  vpmInit Me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Firepower Williams 1980" & vbNewLine & "Created for VPX by 3rdaxis, Slydog43 & G5K"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowFrame = 0
    .HandleMechanics = 0
    .ShowDMDOnly = 1
    .Hidden = 0
    .dip(0)=&h00  'Set to usa

  if myPrefs_CustomPlayerSetup = 1  Then
    myPrefs_ShowDirections = 0
  end if

  if myPrefs_CustomPlayerSetup = 0 then      'different
    myPrefs_Player1Setup = 0
    myPrefs_Player2Setup = 1
    myPrefs_Player3Setup = 1
    myPrefs_Player4Setup = 1
  end if

  if myPrefs_FactorySetting = 1 Then
    myPrefs_CustomPlayerSetup = 1                             '0=Off 1=On This must be set to 1 if changing the PlayerSetup below. This will also lock-out the ability to change it in game.
    myPrefs_Player1Setup = 0                         '0=Human 1=Computer
    myPrefs_Player2Setup = 0                         '0=Human 1=Computer
    myPrefs_Player3Setup = 0                         '0=Human 1=Computer
    myPrefs_Player4Setup = 0                         '0=Human 1=Computer
    myPrefs_ShowDirections = 0                     '0=Off 1=On Shows Direction menu at the start of the game.
    myPrefs_SoundEffects = 0                         '0=off 1=On Extra Sound Effects And Lighting During Computer's Turn,
    myPrefs_Hal9000EyeOnOff = 0                   '0=Off 1=On Turns the visibility of the Hall9000's eye On/Off
    myPrefs_PlayMusic = 0                               '0=Factory 1=On  Begining And End Game Extra Music.
    MyPrefs_DynamicCards = 0                         '1=Instruction Cards will change to your prefered seletion when game is started and back to U.I.Instruction cards when game is over
    LightDifficulty4.ShowBulbMesh=0:LightDifficulty4.intensity=0
    LightDifficulty3.ShowBulbMesh=0:LightDifficulty3.intensity=0
    LightDifficulty2.ShowBulbMesh=0:LightDifficulty2.intensity=0
    LightDifficulty1.ShowBulbMesh=0:LightDifficulty1.intensity=0
    LightDifficulty0.ShowBulbMesh=0:LightDifficulty0.intensity=0
    Player1Light.ShowBulbMesh=0:Player1Light.intensity=0
    Player2Light.ShowBulbMesh=0:Player2Light.intensity=0
    Player3Light.ShowBulbMesh=0:Player3Light.intensity=0
    Player4Light.ShowBulbMesh=0:Player4Light.intensity=0
    Hal9000Eye.Visible = False
    Hal9000Ring.Visible = False
    HalLight2.Intensity = 0
    StrobeLight1.Intensity = 0
    LightDTP1.intensity=0
    BGR1up.Image = "1up"
    LightDTP2.intensity=0
    BGR2up.Image = "2up"
    LightDTP3.intensity=0
    BGR3up.Image = "3up"
    LightDTP4.intensity=0
    BGR4up.Image = "4up"
    if myPrefs_FlipperBatType=6 then
      myPrefs_FlipperBatType = 0
    end if
  end if

  If myPrefs_ShowDirections = 0 Then
    Directions.visible = False
    DirectionBD.visible = False
    DirectionBDtrim.visible = False
    PlayerNumber.visible = False
    Light007.State=0
    Light008.State=0
    DP1.transz=-200'visible = False
    DP2.transz=-200
    DP3.transz=-200
    DP4.transz=-200
    JoeyComp01.transz=-200
    JoeyComp02.transz=-200
    JoeyComp03.transz=-200
    JoeyComp04.transz=-200
    LightD0.intensity =0
    LightD1.intensity =0
    LightD2.intensity =0
    LightD3.intensity =0
    LightD4.intensity =0
    P1Light.intensity =0
    P2Light.intensity =0
    P3Light.intensity =0
    P4Light.intensity =0
    Light007.intensity=0
  Else
    Light007.State=2
    Light008.State=2
  End If

  If myPrefs_AInudge = 0 then
    PostLight.Duration 2, 1200, 0
    playsound"servo_small"
  End If

call myUpdateDMDColor
'call LoadFlipperColor

    On Error Resume Next
    ' .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With
  Controller.Run
  PinMameTimer.enabled = 1
  vpmMapLights AllLights

  vpmNudge.TiltSwitch=cPlayfieldTiltSW
    vpmNudge.Sensitivity=1

  Call myAdjustTableToPrefs()
  Call myAdjustPlayfieldLamps()
  Call MyShowInstructionCard

  if myPrefs_TestWallOn = 1 Then
    TestWallLeft.collidable = 1
    TestWallRight.collidable = 1
    TestWallLeft.visible = 1
    TestWallRight.visible = 1
  end if

  '*********************************************
  'Setup Ball Trough Object
  Set trTrough = new cvpmTrough
    trTrough.CreateEvents "trTrough", Array(Drain, BallRelease)
    trTrough.balls = 3
    trTrough.size = 3
    trTrough.EntrySw = cOutHoleSW
    trTrough.InitEntrySounds "DrainShort", "DrainShort", "DrainShort"
    trTrough.InitExitSounds "BallRelease", "BallRelease"
    trTrough.addsw 2,cLeftBallRampSW
    trTrough.addsw 1, cCenterBallRampSW
    trTrough.addsw 0, cRightBallRampSW
    trTrough.initexit BallRelease, 90, 10
    trTrough.StackExitBalls = 1
    trTrough.MaxBallsPerKick = 1
    trTrough.Reset

  '**********************************************

call myUpdateDisplayColors

  Dim DesktopMode: DesktopMode = Table1.ShowDT

  If DesktopMode = True Then 'Show Desktop components
    DisplayTimer7.enabled = True
Dim myItem

For each myItem in BackboxStuff
  myItem.Visible = True
Next
    BaD1.visible = 1
    BaD2.visible = 1
    CrD1.visible = 1
    CrD2.visible = 1
    RailsLockbar.visible=1
    BLADES.visible=0
    LockdownBar.visible=1 'AXS
    SideRails.visible=1     'AXS

  Else
    DisplayTimer7.enabled = False
For each myItem in BackboxStuff
  myItem.Visible = False
Next
    BGR1up.visible=0
    BGR2up.visible=0
    BGR3up.visible=0
    BGR4up.visible=0
    BGR1cp.visible=0
    BGR2cp.visible=0
    BGR3cp.visible=0
    BGR4cp.visible=0
    BGRtilt.visible=0
    BGRgo.visible=0
    BGRsa.visible=0
    BGRmatch.visible=0
    BGRbip.visible=0
    BGRhs.visible=0
    RailsLockbar.visible=0
    BLADES.visible=1
    LockdownBar.visible=0 'AXS
    SideRails.visible=0     'AXS
  end If

  LUTBox.Visible = 0
  SetLUT

  GameDataTracking_Init

  Call myLightSeq

end Sub

Dim gameplayed
gameplayed = 0

Sub myLightSeq 'Fancy U.I. light sweep
  if gamestarted = 0 then
'msgbox "on"
    LightSeqTimer.enabled=1
    LightDifficulty4.Duration 2, 5100, 0
    LightDifficulty3.Duration 1, 5050, 0
    LightDifficulty2.Duration 1, 5000, 0
    LightDifficulty1.Duration 1, 4950, 0
    LightDifficulty0.Duration 1, 4900, 0
    Player1Light.Duration 1, 4850, 0
    Player2Light.Duration 1, 4800, 0
    Player3Light.Duration 1, 4750, 0
    Player4Light.Duration 1, 4700, 0
  if myPrefs_FactorySetting = 0 and gamestarted = 0 then
    if gameplayed = 0 then
      LightD4.Duration 2, 5100, 0
      LightD3.Duration 1, 5050, 0
      LightD2.Duration 1, 5000, 0
      LightD1.Duration 1, 4950, 0
      LightD0.Duration 1, 4900, 0
      P1Light.Duration 1, 4850, 0
      P2Light.Duration 1, 4800, 0
      P3Light.Duration 1, 4750, 0
      P4Light.Duration 1, 4700, 0
    end if
  end if
  end If
End Sub

Sub LightSeqTimer_Timer
'msgbox "off"
    LightSeqTimer.enabled=0
    LightSeqTimer2.enabled=1
    LightDifficulty4.Duration 0, 410, 2
    LightDifficulty3.Duration 0, 360, 1
    LightDifficulty2.Duration 0, 310, 1
    LightDifficulty1.Duration 0, 260, 1
    LightDifficulty0.Duration 0, 210, 1
    Player1Light.Duration 0, 160, 1
    Player2Light.Duration 0, 110, 1
    Player3Light.Duration 0, 60, 1
    Player4Light.Duration 0, 10, 1
  if myPrefs_FactorySetting = 0 and gamestarted = 0 then
    if gameplayed = 0 then
      LightD4.Duration 0, 410, 2
      LightD3.Duration 0, 360, 1
      LightD2.Duration 0, 310, 1
      LightD1.Duration 0, 260, 1
      LightD0.Duration 0, 210, 1
      P1Light.Duration 0, 160, 1
      P2Light.Duration 0, 110, 1
      P3Light.Duration 0, 60, 1
      P4Light.Duration 0, 10, 1
    end if
  end if
End Sub

Sub LightSeqTimer2_Timer
    Call myLightSeq
    LightSeqTimer2.enabled=0
End Sub

Sub SetLUT
  if PanelOpen = 0 then
  Select Case MyPrefs_Brightness
    Case 4:table1.ColorGradeImage = 0
    Case 3:table1.ColorGradeImage = "AA_FS_Lut30perc"
    Case 2:table1.ColorGradeImage = "AA_FS_Lut50perc"
    Case 1:table1.ColorGradeImage = "AA_FS_Lut70perc"
    Case 0:table1.ColorGradeImage = "AA_FS_Lut100perc"
  end Select
  Playsound "_button-click"
  end if
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  LUTBox.text = "Brightness: " & CStr(MyPrefs_Brightness)
  LUTBox.TimerEnabled = 1
End Sub

'*************************************************************
' Setup Switch Subs

Sub LeftOutsideRollover_hit()
  vpmtimer.pulsesw cLeftOutsideRolloverSW
  if LShieldOn.state = 1 and isGameOver = 0 Then
    KBallSaveKicker.enabled = 1
  Else
    KballsaveKicker.enabled = 0
  end if
End Sub

Sub LeftInsideRollover_hit:vpmtimer.pulsesw cLeftInsideRolloverSW:End Sub
Sub LeftInsideRollover_unhit:Controller.Switch(cLeftInsideRolloverSW)=0:End Sub

Sub UpperMiddleLeftStandup_hit:vpmtimer.pulsesw cUpperMiddleLeftStandupSW:PlaysoundAtVol "rubber_hit_3", ActiveBall, 1 :End Sub
Sub UpperMiddleLeftStandup_unhit:Controller.Switch(cUpperMiddleLeftStandupSW)=0:End Sub

Sub Spinner_Spin:vpmTimer.PulseSw cSpinnerSW:PlaySoundAtVol "fx_spinner", Spinner, 1:End Sub

Sub TopLeftStandup_hit:vpmtimer.pulsesw cTopLeftStandupSW:PlaysoundAtVol "rubber_hit_3", ActiveBall, 1 :End Sub
Sub TopLeftStandup_unhit:Controller.Switch(cTopLeftStandupSW)=0:End Sub

Sub FRollover_hit:vpmtimer.pulsesw cFRolloverSW:End Sub
Sub FRollover_unhit:Controller.Switch(cFRolloverSW)=0:End Sub

Sub IRollover_hit:vpmtimer.pulsesw cIRolloverSW:End Sub
Sub IRollover_unhit:Controller.Switch(cIRolloverSW)=0:End Sub

Sub RRollover_hit:vpmtimer.pulsesw cRRolloverSW:End Sub
Sub RRollover_unhit:Controller.Switch(cRRolloverSW)=0:End Sub

Sub ERollover_hit:vpmtimer.pulsesw cERolloverSW:End Sub
Sub ERollover_unhit:Controller.Switch(cERolloverSW)=0:End Sub

Sub RightInsideRollover_hit:vpmtimer.pulsesw cRightInsideRolloverSw:End Sub
Sub RightInsideRollover_unhit:Controller.Switch(cRightInsideRolloverSW)=0:End Sub

Sub RightOutsideRollover_hit:vpmtimer.pulsesw cRightOutsideRolloverSW:End Sub
Sub RightOutsideRollover_unhit:Controller.Switch(cRightOutsideRolloverSW)=0:End Sub

Sub BallShooter_hit:Controller.Switch(cBallShooterSW)=1:End Sub
Sub BallShooter_unhit:Controller.Switch(cBallShooterSW)=0:End Sub

Sub LeftEjectRollover_hit:vpmtimer.pulsesw cLeftEjectRolloverSW:End Sub
Sub LeftEjectRollover_unhit:Controller.Switch(cLeftEjectRolloverSW)=0:End Sub

Sub RightEjectRollover_hit:vpmtimer.pulsesw cRightEjectRolloverSW:End Sub
Sub RightEjectRollover_unhit:Controller.Switch(cRightEjectRolloverSW)=0:End Sub

'        ___  __  ____________    ___  __   _____  __
'*****  / _ |/ / / /_  __/ __ \  / _ \/ /  / _ \ \/ /*****
'***** / __ / /_/ / / / / /_/ / / ___/ /__/ __ |\  / ***** ASC
'*****/_/ |_\____/ /_/  \____/ /_/  /____/_/ |_|/_/  *****

Dim toggleplayers, Halspeak, Halspeak1, ToggleDifficulty, ComboL, ComboR, Combo2L, Combo2R,ComboButtons, CheckStatus, PlayerSelectVoice, DifficultyVoice, Coins1, Coins2
Halspeak1 = 1
ToggleDifficulty = 0
CheckStatus = 0
PlayerSelectVoice = 1
DifficultyVoice = 1
Coins1 = 0
Coins2 = 0

Sub Table1_KeyDown(ByVal keycode)

  if keycode = LeftMagnaSave and myPrefs_GameDataTracking = 1 Then
    msgbox "Scores" & vbcrlf & GameDataTracking_PlayerScore(1) & vbcrlf & GameDataTracking_PlayerScore(2) & vbcrlf & GameDataTracking_PlayerScore(3) & vbcrlf & GameDataTracking_PlayerScore(4) & _
      vbcrlf & vbcrlf &  "Leader  " & GameDataTracking_LeadingPlayer & _
      vbcrlf & vbcrlf &  "Credits  " & GameDataTracking_Credits & _
      vbcrlf & vbcrlf &  "Ball / Match  " & GameDataTracking_BallMatch

  End if

  if keycode = Startgamekey then
    if CBool(Controller.Lamp(56)) = True then
    gameplayed = 1
    gamestarted = 1
    LightSeqTimer.Enabled=0
    LightSeqTimer2.Enabled=0
    DirectionsTimer.enabled=1
    Stopsound "__Ode To The Sun (Outro)"
    StopSound "__SpaceAmbient(music)"
    StopSound "__SpaceWaltz(music)"
    StopSound "__My God, its full of stars(music)"
    Stopsound "__RequiemIntro(music)"
    Stopsound "_Hal-9000Introduction"
    Stopsound "_MissionControl"
    P1Light.State=0'Duration 2, 2000, 0
    P2Light.State=0'Duration 2, 2000, 0
    P3Light.State=0'Duration 2, 2000, 0
    P4Light.State=0'Duration 2, 2000, 0
    Player1Light.Duration 2, 2000, 0
    Player2Light.Duration 2, 1900, 0
    Player3Light.Duration 2, 1800, 0
    Player4Light.Duration 2, 1700, 0
    LightDIfficulty0.Duration 2, 2100, 0
    LightDIfficulty1.Duration 2, 2200, 0
    LightDIfficulty2.Duration 2, 2300, 0
    LightDIfficulty3.Duration 2, 2400, 0
    LightDIfficulty4.Duration 2, 2500, 0
    Trigger001.enabled=0
    If Directionsoff = 0 then
      If myPrefs_ShowDirections = 1 Then
        Light007.Duration 2, 80, 0
        P1Light.Duration 2, 2000, 0
        P2Light.Duration 2, 1900, 0
        P3Light.Duration 2, 1800, 0
        P4Light.Duration 2, 1700, 0
        LightD0.Duration 2, 2100, 0
        LightD1.Duration 2, 2200, 0
        LightD2.Duration 2, 2300, 0
        LightD3.Duration 2, 2400, 0
        LightD4.Duration 2, 2500, 0
        Light008.Duration 2, 80, 0
        Directionsoff = 1
      End if
    End if
      If MyPrefs_DynamicCards = 1 Then
        Call MyShowInstructionCard
        Playsound "CardChange"
        Light004.Duration 2, 80, 0
        Light005.Duration 2, 80, 0
      End If
    End If
  If CreditLight3.state = 2 Then
    If gamestarted = 0 Then
      playsound"_Concerned"
    End If
  End if
  End If

  If keycode = PlungerKey and AIOn =0 Then
    Plunger.PullBack
    PlaySoundAtVol "plungerpull", ActiveBall, 1
  End If

  If keycode = RightMagnaSave and gamestarted = 1 and PanelOpen = 1 and myPrefs_FactorySetting = 1 Then
    'Playsound "_button-click"
    'if myPrefs_FactorySetting = 1 Then
      myPrefs_FlipperBatType = myPrefs_FlipperBatType - 1
      Light12.Duration 2, 80, 0
      Light13.Duration 2, 80, 0
      Flasher6.visible = True
      Flasher6Timer.Enabled = 1
      Playsound "Laserblast"
    'end if
    if myPrefs_FlipperBatType < 0 Then
      myPrefs_FlipperBatType = 13
    end if
  end if

  If keycode = LeftMagnaSave and gamestarted = 1 Then
    'Playsound "button-clickOff"
  end if

  If keycode = 4 or keycode = 5 or keycode = 6 Then
    PlaysoundAtVol "CoinIn3", Drain, 1
    myCreditCounter  = myCreditCounter  + 1
  End if

  if not (AIon = 1 and (keycode = LeftFlipperKey or keycode = RightFlipperKey or keycode = lefttiltkey or keycode = righttiltkey or keycode = centertiltkey)) Then
    vpmKeyDown(keycode)
  End If
  'Msgbox Keycode  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>(un-comment "Msgbox Keycode" to find keystroke number value of prefered button in game)

  if keycode = LeftMagnaSave and gamestarted = 1 and HalShutdownTimer.Enabled=0 and AIon = 1 then  'Turn Hal off 1of2
    Combo2L = 1
  end if
  if keycode = RightMagnaSave and gamestarted = 1 and HalShutdownTimer.Enabled=0 and AIon = 1 Then 'Turn Hal off 2of2
    Combo2R = 1
  end if
  if Combo2L = 1 and Combo2R = 1 and HalShutdownTimer.Enabled=0 and AIon = 1 Then 'Turn Hal off Start (Dim light Hal voice)
'doleftnudge =1
'dorightnudge = 1
'docenternudge = 1
  Combo2L = 0
  Combo2R = 0
  HalShutdownTimer.Enabled=1
  HalShutdown2Timer.Enabled=1
  Hal9000eye.material = "Hal9000eye8"
  playsound "_HalClitch"
  playsound "_HalClitch2"
  playsound "_HalClitch"
  playsound "_HalClitch2"
  playsound "_HalClitch"
  playsound "_HalClitch2"
  playsound "_button-click"
  playsound "_DeactivatingAutoPlay"
  playsound "_DeactivatingAutoPlay"
  playsound "_DeactivatingAutoPlay"
  stopsound "_AreYouSure"
  stopsound "_Concerned"
  stopsound "_Daisy"
  stopsound "_DisconnectMe"
  stopsound "_DaveStop"
  stopsound "_DeactivatingAutoPlay"
        stopsound "_FirstLesson2"
        stopsound "_VoiceTest"
        stopsound "_VoiceTest2"
        stopsound "_GoingWell"
        stopsound "_SorryDave"
        stopsound "_Hal-9000Introduction"
        stopsound "_iampushingmyself"
        stopsound "_OpenthedoorHal"
        stopsound "_WorkingWithPeople"
  stopsound "_I'maHall9000"
  FlipperPulseTimer.enabled=0
end if

 if keycode = LeftFlipperKey and gamestarted = 0 Then
  ComboL = 1
        DisablePSLTimer.Enabled = 1
ResetPlayerSelectVoice.Enabled = 0
   end if

 if keycode = LeftFlipperKey and gamestarted = 1 and PanelOpen = 1 Then 'sds
  Call myPowerFlash
  Select Case Int(Rnd()*3)
    Case 0
      Playsound "Statc1":Playsound "Static1":playsound "_button-click":playsound "_button-click"
    Case 1
      Playsound "Static2":Playsound "Static2":playsound "_button-click":playsound "_button-click"
    Case 2
      Playsound "Static3":Playsound "Static3":playsound "_button-click":playsound "_button-click"
  End Select
  myPrefs_DisplayColor = myPrefs_DisplayColor + 1
  if myPrefs_DisplayColor > 4 Then
    myPrefs_DisplayColor = 0
  end if
    call myUpdateDMDColor
  call myUpdateDisplayColors

 end if

 if keycode = LeftMagnaSave and gamestarted = 1 and PanelOpen = 1 and AIOff Then
  'MyPrefs_DisableBrightness = 1
  myPrefs_FlipperBatType = myPrefs_FlipperBatType + 1
  Light12.Duration 2, 80, 0
  Light13.Duration 2, 80, 0
  Flasher6.visible = True
  Flasher6Timer.Enabled = 1
  Playsound "Laserblast"
  if myPrefs_FlipperBatType > 13 Then
    myPrefs_FlipperBatType = 0
  end if
 end if

 if keycode = RightMagnaSave and gamestarted = 1 and PanelOpen = 1 and myPrefs_FactorySetting = 0 Then
  playsound "button-clickOff"
  myPrefs_SoundEffects = myPrefs_SoundEffects + 1
  if myPrefs_SoundEffects > 1 Then
    myPrefs_SoundEffects = 0
  end if
 end if

 if keycode = RightFlipperKey and gamestarted = 0 Then
  ComboR = 1
        DisablePSRTimer.Enabled = 1
  ResetPlayerSelectVoice.Enabled = 0
  end if

  if keycode = RightFlipperKey and gamestarted = 1 Then
        CheckStatusTimer.Enabled = 1
' if Light2.intensity>0 then 'logic panel open
'   MyPrefs_DisableBrightness = 1
' end if
  end if

  if ComboL = 1 And ComboR = 1 and myPrefs_FactorySetting= 0 Then
      ComboButtonsTimer.Enabled = 1
        DisablePSLTimer.Enabled = 0
        DisablePSRTimer.Enabled = 0
 if DifficultyVoice = 1 Then
  DifficultyVoice = 0
    if myPrefs_Language = 0 then
  Playsound "Difficulty":Playsound "Difficulty"
        Stopsound "PlayerSelect"
    end If
    if myPrefs_Language = 1 Then
  Playsound "Difficulty_German":Playsound "Difficulty_German"
        Stopsound "PlayerSelect_German"
    end if
    if myPrefs_Language = 2 Then
  Playsound "Difficulty_Italian":Playsound "Difficulty_Italian"
        Stopsound "PlayerSelect_Italian"
    end if
    if myPrefs_Language = 3 Then
  Playsound "Difficulty_French":Playsound "Difficulty_French"
        Stopsound "PlayerSelect_French"
    end if
 end If

      if gamestarted = 0  Then
           ToggleDifficulty = 1
     else
           ToggleDifficulty = 0
    end If
      end if

 if keycode = RightMagnaSave and ToggleDifficulty = 1 and ComboButtons = 1 Then 'Difficulty
       myPrefs_Difficulty = myPrefs_Difficulty + 1
 end if
    if  myPrefs_Difficulty >4 Then
    myPrefs_Difficulty = 0
          end if

 if keycode = LeftMagnaSave and ToggleDifficulty = 1 and ComboButtons = 1 Then
      myPrefs_Difficulty = myPrefs_Difficulty - 1
 end if
          if  myPrefs_Difficulty < 0 Then
    myPrefs_Difficulty = 4
    end if

 if keycode = myPrefs_LanguageOnOffKey and gamestarted = 0 Then 'Language slect in game "L"
      PlayerSelectVoice = 1
if myPrefs_Language = 3 then playsound "English" end If
if myPrefs_Language = 0 then playsound "German" end If
if myPrefs_Language = 1 then playsound "Italian" end If
if myPrefs_Language = 2 then playsound "French":playsound "French":playsound "French" end If
     myPrefs_Language = myPrefs_Language +1
 end If
 if myPrefs_Language > 3 Then
      myPrefs_Language = 0
end if


 if myPrefs_Difficulty = 0 and ComboButtons = 1 and gamestarted = 0 Then
  Select Case myPrefs_Language
    Case 0
      Playsound "_button-click"
      Playsound "Easy_English":Playsound "Easy_English"
      Stopsound "Medium_English"
      Stopsound "Hard_English"
      Stopsound "Expert_English"
      Stopsound "Legendary"
    Case 1
      Playsound "_button-click"
      Playsound "Easy_German":Playsound "Easy_German"
      Stopsound "Medium_German"
      Stopsound "Hard_German"
      Stopsound "Expert_German"
      Stopsound "Legendary"
    Case 2
      Playsound "_button-click"
      Playsound "Easy_Italian":Playsound "Easy_Italian"
      Stopsound "Medium_Italian"
      Stopsound "Hard_Italian"
      Stopsound "Expert_Italian"
      Stopsound "Legendary"
    Case 3
      Playsound "_button-click"
      Playsound "Easy_French":Playsound "Easy_French"
      Stopsound "Medium_French"
      Stopsound "Hard_French"
      Stopsound "Expert_French"
      Stopsound "Legendary"
  End Select
end if

 if myPrefs_Difficulty = 1  and ComboButtons = 1 and gamestarted = 0 Then
        Select Case myPrefs_Language
    Case 0
      Playsound "_button-click"
      Playsound "Medium_English":Playsound "Medium_English"
      Stopsound "Easy_English"
      Stopsound "Hard_English"
      Stopsound "Expert_English"
      Stopsound "Legendary"
    Case 1
      Playsound "_button-click"
      Playsound "Medium_German": Playsound "Medium_German"
      Stopsound "Easy_German"
      Stopsound "Hard_German"
      Stopsound "Expert_German"
      Stopsound "Legendary"
    Case 2
      Playsound "_button-click"
      Playsound "Medium_Italian": Playsound "Medium_Italian"
      Stopsound "Easy_Italian"
      Stopsound "Hard_Italian"
      Stopsound "Expert_Italian"
      Stopsound "Legendary"
    Case 3
      Playsound "_button-click"
      Playsound "Medium_French": Playsound "Medium_French"
      Stopsound "Easy_French"
      Stopsound "Hard_French"
      Stopsound "Expert_French"
      Stopsound "Legendary"
  End Select
end if

 if myPrefs_Difficulty = 2  and ComboButtons = 1 and gamestarted = 0 Then
  Select Case myPrefs_Language
    Case 0
      Playsound "_button-click"
      Playsound "Hard_English":Playsound "Hard_English"
      Stopsound "Medium_English"
      Stopsound "Easy_English"
      Stopsound "Expert_English"
      Stopsound "Legendary"
    Case 1
      Playsound "_button-click"
      Playsound "Hard_German":Playsound "Hard_German"
      Stopsound "Medium_German"
      Stopsound "Easy_German"
      Stopsound "Expert_German"
      Stopsound "Legendary"
    Case 2
      Playsound "_button-click"
      Playsound "Hard_Italian":Playsound "Hard_Italian"
      Stopsound "Medium_Italian"
      Stopsound "Easy_Italian"
      Stopsound "Expert_Italian"
      Stopsound "Legendary"
    Case 3
      Playsound "_button-click"
      Playsound "Hard_French": Playsound "Hard_French"
      Stopsound "Easy_French"
      Stopsound "Medium_French"
      Stopsound "Expert_French"
      Stopsound "Legendary"
  End Select
end If

 if myPrefs_Difficulty = 3  and ComboButtons = 1 and gamestarted = 0 Then
  Select Case myPrefs_Language
    Case 0
      Playsound "_button-click"
      Playsound "Expert_English":Playsound "Expert_English"
      Stopsound "Medium_English"
      Stopsound "Hard_English"
      Stopsound "Easy_English"
      Stopsound "Legendary"
    Case 1
      Playsound "_button-click"
      Playsound "Expert_German":Playsound "Expert_German"
      Stopsound "Medium_German"
      Stopsound "Hard_German"
      Stopsound "Easy_German"
      Stopsound "Legendary"
    Case 2
      Playsound "_button-click"
      Playsound "Expert_Italian":Playsound "Expert_Italian"
      Stopsound "Medium_Italian"
      Stopsound "Hard_Italian"
      Stopsound "Easy_Italian"
      Stopsound "Legendary"
    Case 3
      Playsound "_button-click"
      Playsound "Expert_French": Playsound "Expert_French"
      Stopsound "Easy_French"
      Stopsound "Hard_French"
      Stopsound "Medium_French"
      Stopsound "Legendary"
  End Select
end if

 if myPrefs_Difficulty = 4  and ComboButtons = 1 and gamestarted = 0 Then
  Select Case myPrefs_Language
    Case 0
      Playsound "_button-click"
      Playsound "Legendary":Playsound "Legendary"
      Stopsound "Medium_English"
      Stopsound "Hard_English"
      Stopsound "Easy_English"
    Case 1
      Playsound "_button-click"
      Playsound "Legendary":Playsound "Legendary"
      Stopsound "Medium_German"
      Stopsound "Hard_German"
      Stopsound "Easy_German"
    Case 2
      Playsound "_button-click"
      Playsound "Legendary":Playsound "Legendary"
      Stopsound "Medium_Italian"
      Stopsound "Hard_Italian"
      Stopsound "Easy_Italian"
    Case 3
      Playsound "_button-click"
      Playsound "Legendary": Playsound "Legendary"
      Stopsound "Easy_French"
      Stopsound "Hard_French"
      Stopsound "Expert_French"
      Stopsound "Medium_French"
  End Select
end if

 if keycode = RightMagnaSave and gamestarted = 0 and ComboButtons = 0 then  'Flipper Bat Select
  Light12.Duration 2, 80, 0
  Light13.Duration 2, 80, 0
  Flasher6.visible = True
  Flasher6Timer.Enabled = 1
      myPrefs_FlipperBatType =  myPrefs_FlipperBatType + 1
Playsound "Laserblast"
 end if

 if myPrefs_FlipperBatType > 13 then
      myPrefs_FlipperBatType = 0
 end if

 if keycode = LeftMagnaSave and gamestarted = 0  and ComboButtons = 0 then
  Light12.Duration 2, 80, 0
  Light13.Duration 2, 80, 0
  Flasher6.visible = True
  Flasher6Timer.Enabled = 1
      myPrefs_FlipperBatType =  myPrefs_FlipperBatType - 1
Playsound "Laserblast"
 end if

 if myPrefs_FlipperBatType < 0 then
      myPrefs_FlipperBatType = 13
 end if

  if keycode = RightFlipperKey and gamestarted = 0 Then 'Player A.I. Setup
  toggleplayers = toggleplayers + 1
   If Halspeak1 = 1 then
        Halspeak1 = 0
   end if
  end if
  if toggleplayers > 6 Then
  toggleplayers = 0
  end If

  if keycode = LeftFlipperKey and gamestarted = 0 Then
  toggleplayers = toggleplayers - 1
   If Halspeak1 = 1 then
      Halspeak1 = 0
   end if
  end if

  if toggleplayers < 0 Then
  toggleplayers = 6
  end if

if myPrefs_CustomPlayerSetup = 0 Then
  select case toggleplayers
  case 0
    myPrefs_Player1Setup = 0
    myPrefs_Player2Setup = 1
    myPrefs_Player3Setup = 1
    myPrefs_Player4Setup = 1
  case 1
    myPrefs_Player1Setup = 0
    myPrefs_Player2Setup = 0
    myPrefs_Player3Setup = 1
    myPrefs_Player4Setup = 1
  case 2
    myPrefs_Player1Setup = 0
    myPrefs_Player2Setup = 0
    myPrefs_Player3Setup = 0
    myPrefs_Player4Setup = 1
  case 3
    myPrefs_Player1Setup = 0
    myPrefs_Player2Setup = 0
    myPrefs_Player3Setup = 0
    myPrefs_Player4Setup = 0
  case 4
    myPrefs_Player1Setup = 0
    myPrefs_Player2Setup = 1
    myPrefs_Player3Setup = 0
    myPrefs_Player4Setup = 1
  case 5
    myPrefs_Player1Setup = 1
    myPrefs_Player2Setup = 0
    myPrefs_Player3Setup = 1
    myPrefs_Player4Setup = 0
  case 6
    myPrefs_Player1Setup = 1
    myPrefs_Player2Setup = 1
    myPrefs_Player3Setup = 1
    myPrefs_Player4Setup = 1
  end select
 end If

End Sub

Sub   ComboButtonsTimer_Timer
        ComboButtons = 1
  PlayerSelectVoice = 1
End Sub

Sub HalShutdownTimer_Timer 'Hal shut down
    FlipperLeft.RotateToStart
    FlipperRight.RotateToStart
    vpmkeyup(LeftFlipperKey)
    vpmkeyup(RightFlipperKey)
    doleftnudge =1
    dorightnudge = 1
    docenternudge = 1
    ToggleAI(0)
    HalShutdownTimer.Enabled=0
    HalShutdown2Timer.Enabled=0
    playsound "_spin-down"
    stopsound "_DaveStop"
    stopsound "_DaveStop"
    stopsound "_DaveStop"
End Sub

Sub HalShutdown2Timer_Timer
doleftnudge =1
dorightnudge = 1
docenternudge = 1
    playsound "_air-release"
    playsound "_DaveStop"
    playsound "_DaveStop"
    playsound "_DaveStop"
    HalShutdown2Timer.Enabled=0
End Sub

Sub DisablePSRTimer_Timer
  if myPrefs_CustomPlayerSetup = 0 then
   Playsound "_plug-in"
         Playsound "_air-release"
   DisablePSRTimer.Enabled = 0
    if PlayerSelectVoice = 1 Then
         PlayerSelectVoice = 0
  select case myPrefs_Language
    case 0
      Playsound "PlayerSelect":Playsound "PlayerSelect"
      Stopsound "Difficulty"
    case 1
      Playsound "PlayerSelect_German":Playsound "PlayerSelect_German"
      Stopsound "Difficulty_German"
    case 2
      Playsound "PlayerSelect_Italian":Playsound "PlayerSelect_Italian"
      Stopsound "Difficulty_Italian"
    case 3
      Playsound "PlayerSelect_French":Playsound "PlayerSelect_French"
      Stopsound "Difficulty_French"
    end select
   end if
 end if

 if PlayerSelectVoice = 0 Then
    ResetPlayerSelectVoice.Enabled = 1
 end If
End Sub

Sub  ResetPlayerSelectVoice_Timer
          PlayerSelectVoice = 1
  ResetPlayerSelectVoice.Enabled = 0
End Sub

Sub DisablePSLTimer_Timer
  if myPrefs_CustomPlayerSetup = 0 then
  Playsound "_plug-inb"
  Playsound "_air-releaseb"
  DisablePSLTimer.Enabled = 0
    if PlayerSelectVoice = 1 Then
         PlayerSelectVoice = 0
select case myPrefs_Language
case 0
         Playsound "PlayerSelect":Playsound "PlayerSelect"
case 1
         Playsound "PlayerSelect_German":Playsound "PlayerSelect_German"
case 2
         Playsound "PlayerSelect_Italian":Playsound "PlayerSelect_Italian"
case 3
         Playsound "PlayerSelect_French":Playsound "PlayerSelect_French"
end select
    end if
  end if
End Sub

 Sub   CheckStatusTimer_Timer
  If  gamestarted = 1 Then
CheckStatusTimer.Enabled = 0
     if NOT TimerMovePanel.Enabled then
       TimerMovePanel.Enabled = true
          LogicBanner.Visible=True
    call myTurnPanelOnOff(0)
    'call myLightSeq
          Playsound "_button-click"
          Playsound"_ServoMotor"
      Elseif (LogicBanner.TransX = 200) or (LogicBanner.TransX = -200) then 'Panel has moved
            TimerMovePanel.Enabled = false
        call  myTurnPanelOnOff(0)
            Playsound "button-clickOff"
            Playsound"_ServoMotor"

       End If
  End If

 End Sub

 Sub   CheckStatusTimer1_Timer
  If  gamestarted = 1 Then
CheckStatusTimer1.Enabled = 0
     if NOT TimerMovePanel.Enabled then
       TimerMovePanel.Enabled = true
          LogicBanner.Visible=True
    call myTurnPanelOnOff(0)
          Playsound "_button-click"
          Playsound"_ServoMotor"
      Elseif (LogicBanner.TransX = 200) or (LogicBanner.TransX = -200) then 'Panel has moved
            TimerMovePanel.Enabled = false
        call  myTurnPanelOnOff(0)
            Playsound "button-clickOff"
            Playsound"_ServoMotor"

       End If
  End If
          'CheckStatus = 1
'msgbox "status"

 End Sub

 Sub Flasher6Timer_Timer
  Flasher6.visible = False
 End Sub

Sub StrobeLightTimer_Timer
  StrobeLight.State=0
  StrobeLight1.State=0
  StrobeLight3.State=0
  StrobeLight4.State=0
  StrobeLight5.State=0
  StrobeLight6.State=0
        ItsFullOfStars.Material = "StarField"
  StrobeLightTimer.Enabled=0
End Sub

Dim myTriggerLightsOnOff, PanelOpen
PanelOpen = 0

Sub myTurnPanelOnOff(myOnOff)
    if myOnOff = 1 then
        'LogicBanner.Visible=True
myTriggerLightsOnOff=1
Playsound"_WarmUp"
PanelOpen = 1
       Light001.Intensity=10
        Light002.Intensity=10
         ShotonLight1.Intensity=10
    ShotonLight2.Intensity=10
     ShotonLight3.Intensity=10
      ShotonLight4.Intensity=10
       ShotonLight5.Intensity=10
        ShotonLight6.Intensity=10
             ShotonLight7.Intensity=10
          ShotonLight8.Intensity=10
           ShotonLight9.Intensity=10
            ShotonLight10.Intensity=10
             ShotonLight11.Intensity=10
              ShotonLight12.Intensity=10
                     ShotonLight13.Intensity=10
                FireFL.Intensity=10
                 FireIL.Intensity=10
                  FireRL.Intensity=10
                   FireEL.Intensity=10
   B1.Intensity=10
    B2.Intensity=10
     B3.Intensity=10
      B4.Intensity=10
       B5.Intensity=10
        B6.Intensity=10
         B7.Intensity=10
          B8.Intensity=10
           B9.Intensity=10
            B11.Intensity=10
             B12.Intensity=10
              B13.Intensity=10
               B14.Intensity=10
                B15.Intensity=10
                 B16.Intensity=10
B10.Intensity=10
 B17.Intensity=10
  B18.Intensity=10
   B19.Intensity=10
    Light1.Intensity=1
     Light2.Intensity=100
      Light2.FadeSpeedDown =.09
       Light7.Intensity=10
  Light8.Intensity=10
   Light9.Intensity=10
    Light10.Intensity=10
     Light11.Intensity=10
      LightDIfficulty0.State=1
       LightDIfficulty1.State=1
        LightDIfficulty2.State=1
               LightDIfficulty3.State=1
                LightDIfficulty4.State=2
            Player1Light.State=1
       Player2Light.State=1
        Player3Light.State=1
         Player4Light.State=1
elseif myOnOff= 0 Then
PanelOpen = 0
       Light001.Intensity=0
        Light002.Intensity=0
         ShotonLight1.Intensity=0
    ShotonLight2.Intensity=0
     ShotonLight3.Intensity=0
      ShotonLight4.Intensity=0
       ShotonLight5.Intensity=0
        ShotonLight6.Intensity=0
         ShotonLight7.Intensity=0
          ShotonLight8.Intensity=0
           ShotonLight9.Intensity=0
            ShotonLight10.Intensity=0
             ShotonLight11.Intensity=0
              ShotonLight12.Intensity=0
                     ShotonLight13.Intensity=0
                FireFL.Intensity=0
                 FireIL.Intensity=0
                  FireRL.Intensity=0
                   FireEL.Intensity=0
   B1.Intensity=0
    B2.Intensity=0
     B3.Intensity=0
      B4.Intensity=0
       B5.Intensity=0
        B6.Intensity=0
         B7.Intensity=0
          B8.Intensity=0
           B9.Intensity=0
            B11.Intensity=0
             B12.Intensity=0
              B13.Intensity=0
               B14.Intensity=0
                B15.Intensity=0
                 B16.Intensity=0
B10.Intensity=0
 B17.Intensity=0
  B18.Intensity=0
   B19.Intensity=0
    Light1.Intensity=0
     Light2.Intensity=0
      Light2.FadeSpeedDown = 10
       Light7.Intensity=0
        Light8.Intensity=0
         Light9.Intensity=0
    Light10.Intensity=0
     Light11.Intensity=0
LightDIfficulty0.State=0
 LightDIfficulty1.State=0
  LightDIfficulty2.State=0
   LightDIfficulty3.State=0
    LightDIfficulty4.State=0
  LightD0.State=0
   LightD1.State=0
    LightD2.State=0
     LightD3.State=0
      LightD4.State=0
P1Light.State=0
 P2Light.State=0
  P3Light.State=0
    P4Light.State=0
Player1Light.State=0
 Player2Light.State=0
  Player3Light.State=0
    Player4Light.State=0
myTriggerLightsOnOff=0
End If

End Sub

Dim gamestarted, HSL, HSR

'*******************************************************
Dim FadeOutStep:FadeOutStep = 9
Dim FadeOutDir:FadeOutDir = 1

Sub FadeOutTimer_Timer
    FadeOutStep = FadeOutStep - FadeOutDir
    DirectionBD.Material = "MetalFading" & (Fadeoutstep)
    Directions.Material = "MetalFading" & (Fadeoutstep)
    DirectionBDtrim.Material = "MetalFading" & (Fadeoutstep)
    PlayerNumber.Material = "MetalFading" & (Fadeoutstep)
    DP1.Material = "MetalFading" & (Fadeoutstep)
    DP2.Material = "MetalFading" & (Fadeoutstep)
    DP3.Material = "MetalFading" & (Fadeoutstep)
    DP4.Material = "MetalFading" & (Fadeoutstep)
    JoeyComp01.Material = "MetalFading" & (Fadeoutstep)
    JoeyComp02.Material = "MetalFading" & (Fadeoutstep)
    JoeyComp03.Material = "MetalFading" & (Fadeoutstep)
    JoeyComp04.Material = "MetalFading" & (Fadeoutstep)
  if FadeOutStep = 0 Then
    Playsound "short-zing"
    FadeOutTimer.Enabled=0
    Directions.Visible = False
    DirectionBD.Visible = False
    DirectionBDtrim.Visible = False
    DP1.visible = False
    DP2.Visible = False
    DP3.Visible = False
    DP4.Visible = False
    JoeyComp01.visible = False
    JoeyComp02.Visible = False
    JoeyComp03.Visible = False
    JoeyComp04.Visible = False
    PlayerNumber.visible=False
    Light006.Duration 2, 80, 0
  End If
End Sub

Sub DirectionsTimer_Timer
  If myPrefs_ShowDirections = 1 Then
    FadeOutTimer.Enabled=1
    DirectionsTimer.enabled=0
  End If
End Sub

Sub MyShowInstructionCard
  if myPrefs_FactorySetting = 0 Then
    Call MyHideInstructionCards()
  end if
  If GameStarted Then
    Select Case myPrefs_InstructionCardType
      Case 0
        CardLFactory.visible = 1
        CardRFactory.visible = 1
      Case 1
        CardLMod.visible = 1
        CardRMod.visible = 1
      Case 2
        CardLTHX.visible = 1
        CardRTHX.visible = 1
      Case 3
        CardAIRight.Visible = 1
        CardAILeft.Visible = 1
    End Select
  Else
    if myPrefs_FactorySetting = 0 Then
      CardAIRight.Visible = 1
      CardAILeft.Visible = 1
    end if
  End If
End Sub

Sub MyHideInstructionCards
  CardLFactory.visible = 0
  CardRFactory.visible = 0
  CardLMod.visible = 0
  CardRMod.visible = 0
  CardAIRight.Visible = 0
  CardAILeft.Visible = 0
  CardLTHX.visible = 0
  CardRTHX.visible = 0
End Sub

Dim Directionsoff

Sub Table1_KeyUp(ByVal keycode)

  if not (AIon = 1 and (keycode = LeftFlipperKey or keycode = RightFlipperKey or keycode = lefttiltkey or keycode = righttiltkey or keycode = centertiltkey)) Then
    vpmKeyUp(keycode)
  end if

  If keycode = PlungerKey and AIOn = 0 Then
    Plunger.Fire
    PlaySoundAtVol "plunger", Plunger, 1
  End If

  if keycode = LeftFlipperKey Then
    ComboL = 0
    DifficultyVoice = 1
    ComboButtonsTimer.Enabled = 0
    ResetPlayerSelectVoice.Enabled = 1
  end if

  if keycode = RightFlipperKey Then
    'MyPrefs_DisableBrightness = 0
    ComboR = 0
    DifficultyVoice = 1
      CheckStatus= 0
    CheckStatusTimer.Enabled = 0
    ComboButtonsTimer.Enabled = 0
    ResetPlayerSelectVoice.Enabled = 1
  end if

  if ComboL = 0 And ComboR = 0 Then
    ComboButtons = 0
  end if

  If keycode = RightFlipperKey and gamestarted = 1 and LogicBanner.TransX < -392  Then
    CheckStatusTimer1.Enabled = 1
    CheckStatusTimer.Enabled = 0
  end If

 if keycode = LeftMagnaSave or keycode = RightMagnaSave and gamestarted = 1 and AIon = 1  and HalShutdownTimer.Enabled=1 and combo2R = 1 Then  'Turn Hal off cancel  Hello2 ' and combo2R = 1
  Hal9000eye.material = "Hal9000eye6"
  HalShutdownTimer.Enabled=0
  HalShutdown2Timer.Enabled=0
  FlipperPulseTimer.enabled=1
        'playsound "_HalStop"
  stopsound "_DeactivatingAutoPlay"
  stopsound "_DeactivatingAutoPlay"
  stopsound "_DeactivatingAutoPlay"
  stopsound "_AreYouSure"
  stopsound "_Concerned"
  stopsound "_Daisy"
  stopsound "_DisconnectMe"
  stopsound "_DaveStop"
  stopsound "_DeactivatingAutoPlay"
        stopsound "_FirstLesson2"
        stopsound "_VoiceTest"
        stopsound "_VoiceTest2"
        stopsound "_GoingWell"
        stopsound "_SorryDave"
        stopsound "_Hal-9000Introduction"
        stopsound "_iampushingmyself"
        stopsound "_OpenthedoorHal"
        stopsound "_WorkingWithPeople"
  stopsound "_I'maHall9000"
  Combo2L = 0
  Combo2R = 0
   end if


  If keycode = RightMagnaSave and gamestarted = 1 Then
    if MyPrefs_DisableBrightness = 0 and PanelOpen = 0 then
      MyPrefs_Brightness = MyPrefs_Brightness + 1
      if MyPrefs_Brightness > MaxLut then MyPrefs_Brightness = 0
        SetLUT
        ShowLUT
        'Playsound "_button-click"
      end if
    end if

  If keycode = LeftMagnaSave and gamestarted = 1 Then
    if MyPrefs_DisableBrightness = 0 and PanelOpen = 0 then
      MyPrefs_Brightness= MyPrefs_Brightness - 1
      if MyPrefs_Brightness < 0 then MyPrefs_Brightness = MaxLut
        SetLUT
        ShowLUT
        'Playsound "button-clickOff"
      end if
    end if

   if keycode = LeftMagnaSave and gamestarted = 1 and AIon = 1 Then
    HSL = 1
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_HalClitch"
    stopsound "_HalClitch2"
    stopsound "_HalClitch"
    stopsound "_HalClitch2"
    stopsound "_HalClitch"
    stopsound "_HalClitch2"
   end If

   if keycode = RightMagnaSave and gamestarted = 1 and AIon = 1 Then
    HSR = 1
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_HalClitch"
    stopsound "_HalClitch2"
    stopsound "_HalClitch"
    stopsound "_HalClitch2"
    stopsound "_HalClitch"
    stopsound "_HalClitch2"
   end If

   if HSL = 1 and HSR = 1 and gamestarted = 1 and AIon = 1 Then
    playsound "_HalStop"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_DeactivatingAutoPlay"
    stopsound "_HalClitch"
    stopsound "_HalClitch2"
    stopsound "_HalClitch"
    stopsound "_HalClitch2"
    stopsound "_HalClitch"
    stopsound "_HalClitch2"
    HSL = 0
    HSR = 0
   end If

End Sub

dim FlipperPulseStep:FlipperPulseStep = 0
dim FlipperPulseDir:FlipperPulseDir = 1

sub FlipperPulseTimer_Timer                            'LightFadding
  if isMultiball = 1 and myPrefs_Difficulty = 4 Then 'AIon = 1 and
    FlipperPulseStep = FlipperPulseStep + FlipperPulseDir
      if FlipperPulseStep = 8 then FlipperPulseDir = -1
  if FlipperPulseStep = 0 then FlipperPulseDir = 1
  glowbatleft.material = "GIMATERIALShading" & (FlipperPulseStep)
  glowbatright.material = "GIMATERIALShading" & (FlipperPulseStep)
  Hal9000Eye.material = "Hal9000eye" & (FlipperPulseStep)
  CreditLight4.State = 2
'CheckStatusTimer.Enabled = 1 'Hello
  else
  glowbatleft.material = "GIMATERIALShading8"
  glowbatright.material = "GIMATERIALShading8"
      if Combo2L = 0 and Combo2R = 0 then
  glowbatleft.material = "GIMATERIALShading8"
  glowbatright.material = "GIMATERIALShading8"
  Hal9000Eye.material = "Hal9000eye0"
  CreditLight4.State = 0
'CheckStatusTimer.Enabled = 0 'Hello
    end if
 end if
End Sub


'****************************************************************

'Backglass
Dim Value

Sub Backglass_Timer()

  If BGL1up.State=1 then BGR1up.SetValue(1): LightDTP1.State = 2 End If
  If BGL1up.State=0 then BGR1up.SetValue(0): LightDTP1.State = 1 End If
  If BGL2up.State=1 then BGR2up.SetValue(1): LightDTP2.State = 2 End If
  If BGL2up.State=0 then BGR2up.SetValue(0): LightDTP2.State = 1 End If
  If BGL3up.State=1 then BGR3up.SetValue(1): LightDTP3.State = 2 End If
  If BGL3up.State=0 then BGR3up.SetValue(0): LightDTP3.State = 1 End If
  If BGL4up.State=1 then BGR4up.SetValue(1): LightDTP4.State = 2 End If
  If BGL4up.State=0 then BGR4up.SetValue(0): LightDTP4.State = 1 End If
  If BGL1cp.State=1 then BGR1cp.SetValue(1) End If
  If BGL1cp.State=0 then BGR1cp.SetValue(0) End If
  If BGL2cp.State=1 then BGR2cp.SetValue(1) End If
  If BGL2cp.State=0 then BGR2cp.SetValue(0) End If
  If BGL3cp.State=1 then BGR3cp.SetValue(1) End If
  If BGL3cp.State=0 then BGR3cp.SetValue(0) End If
  If BGL4cp.State=1 then BGR4cp.SetValue(1) End If
  If BGL4cp.State=0 then BGR4cp.SetValue(0) End If
  If BGLbip.State=1 then BGRbip.SetValue(1) End If
  If BGLbip.State=0 then BGRbip.SetValue(0) End If
  If BGLmatch.State=1 then BGRmatch.SetValue(1) End If
  If BGLmatch.State=0 then BGRmatch.SetValue(0) End If
  If BGLgo.State=1 then BGRgo.SetValue(1) End If
  If BGLgo.State=0 then BGRgo.SetValue(0) End If
  If BGLsa.State=1 then BGRsa.SetValue(1) End If
  If BGLsa.State=0 then BGRsa.SetValue(0) End If
  If BGLhs.State=1 then BGRhs.SetValue(1) End If
  If BGLhs.State=0 then BGRhs.SetValue(0) End If
  If BGLtilt.State=1 then BGRtilt.SetValue(1) End If
  If BGLtilt.State=0 then BGRtilt.SetValue(0) End If

if AIoff = 1  and AIPlaying = 1 and myPrefs_Difficulty = 4 and AITaunt = 1 Then
  if KUpperEjectHole.ballcntover  and KRightEjectHole.ballcntover then
    Select Case Int(Rnd()*2)
      Case 0
        Playsound "_DeactivatingAutoPlay":Playsound "_DeactivatingAutoPlay"
      Case 1
        Playsound "_DaveStop":Playsound "_DaveStop":Playsound "_DaveStop"
      End Select
  AITaunt = 0
  End If
    End If

End Sub


'--------------------------------------------------------------
'*********BALLKICKER************

Sub KBallSaveKicker_Hit()
  if myPrefs_FlipperStobes > 2 and AIOff then
    Call myPowerFlash
  end if
    Me.kick 0, 25
    Playsound "SlingshotLeft"
End Sub

'--------------------------------------------------------------
' Play sounds when gates are Hit
Sub Gate3_Hit()
  PlaysoundAtVol "GateWire", ActiveBall, Vol(ActiveBall)
End Sub

Sub Gate1_Hit()
  PlaysoundAtVol "GateWire", ActiveBall, Vol(ActiveBall)
End Sub

Sub Gate2_Hit()
  PlaysoundAtVol "GateWire", ActiveBall, Vol(ActiveBall)
End Sub

Sub BallReleaseGate_Hit()
  PlaysoundAtVol "GateWire", ActiveBall, Vol(ActiveBall)
End Sub

'=========================================================
'                    LED Handling
'=========================================================
'Modified version of Scapino's LED code for Fathom
'and borrowed from Uncle Willy's VP9 Firepower
'
Dim SixDigitOutput(32)
Dim SevenDigitOutput(32)
Dim DisplayPatterns(11)
Dim DigStorage(32)

'Binary/Hex Pattern Recognition Array
DisplayPatterns(0) = 0    '0000000 Blank
DisplayPatterns(1) = 63   '0111111 zero
DisplayPatterns(2) = 6    '0000110 one
DisplayPatterns(3) = 91   '1011011 two
DisplayPatterns(4) = 79   '1001111 three
DisplayPatterns(5) = 102  '1100110 four
DisplayPatterns(6) = 109  '1101101 five
DisplayPatterns(7) = 125  '1111101 six
DisplayPatterns(8) = 7    '0000111 seven
DisplayPatterns(9) = 127  '1111111 eight
DisplayPatterns(10)= 111  '1101111 nine

'Assign 7-digit output to reels
Set SevenDigitOutput(0)  = P1D7
Set SevenDigitOutput(1)  = P1D6
Set SevenDigitOutput(2)  = P1D5
Set SevenDigitOutput(3)  = P1D4
Set SevenDigitOutput(4)  = P1D3
Set SevenDigitOutput(5)  = P1D2
Set SevenDigitOutput(6)  = P1D1

Set SevenDigitOutput(7)  = P2D7
Set SevenDigitOutput(8)  = P2D6
Set SevenDigitOutput(9)  = P2D5
Set SevenDigitOutput(10) = P2D4
Set SevenDigitOutput(11) = P2D3
Set SevenDigitOutput(12) = P2D2
Set SevenDigitOutput(13) = P2D1

Set SevenDigitOutput(14) = P3D7
Set SevenDigitOutput(15) = P3D6
Set SevenDigitOutput(16) = P3D5
Set SevenDigitOutput(17) = P3D4
Set SevenDigitOutput(18) = P3D3
Set SevenDigitOutput(19) = P3D2
Set SevenDigitOutput(20) = P3D1

Set SevenDigitOutput(21) = P4D7
Set SevenDigitOutput(22) = P4D6
Set SevenDigitOutput(23) = P4D5
Set SevenDigitOutput(24) = P4D4
Set SevenDigitOutput(25) = P4D3
Set SevenDigitOutput(26) = P4D2
Set SevenDigitOutput(27) = P4D1

Set SevenDigitOutput(28) = CrD2
Set SevenDigitOutput(29) = CrD1
Set SevenDigitOutput(30) = BaD2
Set SevenDigitOutput(31) = BaD1



Sub DisplayTimer7_Timer ' 7-Digit output
  On Error Resume Next
  Dim ChgLED,ii,chg,stat,num,obj,TempCount,temptext,adj

  ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)

  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(ChgLED)
      num=chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      If myPrefs_GameDataTracking = 1 Then GameDataTracking_Update num, stat
      If DesktopMode = True Then
        For TempCount = 0 to 10
          If stat = DisplayPatterns(TempCount) then
            If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
            DigStorage(chgLED(ii, 0)) = TempCount
          End If
          If stat = (DisplayPatterns(TempCount) + 128) then
            If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
            DigStorage(chgLED(ii, 0)) = TempCount
          End If
        Next
      End If
    Next
  End IF
End Sub


'*DOF method for non rom controller tables by Arngrim****************
'*******Use DOF 1**, 1 to activate a ledwiz output*******************
'*******Use DOF 1**, 0 to deactivate a ledwiz output*****************
'*******Use DOF 1**, 2 to pulse a ledwiz output**********************
'Sub DOF(dofevent, dofstate)
' If B2SOn=True Then
'   If dofstate = 2 Then
'     Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
'   Else
'     Controller.B2SSetData dofevent, dofstate
'   End If
' End If
'End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 500, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioPan(ActiveBall)
End Sub

'*************SUPPORTING SOUNDS************************

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioPan(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioPan(ActiveBall)
End Sub

Sub Metal_Hit (idx)
  PlaySound "metalhit_thin", 0, 0.05, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Rubber001_Hit
  PlaySound "metalhit_thin", 0, 0.05, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Rubber002_Hit
  PlaySound "metalhit_thin", 0, 0.05, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub RubbersLower_Hit'(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
  End Select
End Sub

Sub FlipperLeft_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub FlipperRight_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
  End Select
End Sub

'        __   ___   __  ______  ____
'*****  / /  / _ | /  |/  / _ \/ __/*****
'***** / /__/ __ |/ /|_/ / ___/\ \  *****
'*****/____/_/ |_/_/  /_/_/  /___/  *****

Sub myAdjustPlayfieldLamps()

  Select Case myPrefs_PlayfieldLamps
    Case 0 'Factory
      LF.Intensity=10:LF.Color=RGB(0,128,0):LF.ColorFull=RGB(0,255,0):LF.Image= "g_g5kplayfield3"
      LI.Intensity=10:LI.Color=RGB(0,128,0):LI.ColorFull=RGB(0,255,0):LI.Image= "g_g5kplayfield3"
      LR.Intensity=10:LR.Color=RGB(0,128,0):LR.ColorFull=RGB(0,255,0):LR.Image= "g_g5kplayfield3"
      LE.Intensity=10:LE.Color=RGB(0,111,0):LE.ColorFull=RGB(0,111,0):LE.Image= "g_g5kplayfield3"
      L2X.Intensity=200:L2X.Color=RGB(0,128,0):L2X.ColorFull=RGB(0,255,0):L2X.Image= "g_g5kplayfield3"
      L3X.Intensity=200:L3X.Color=RGB(0,128,0):L3X.ColorFull=RGB(0,255,0):L3X.Image= "g_g5kplayfield3"
      L4X.Intensity=200:L4X.Color=RGB(0,128,0):L4X.ColorFull=RGB(0,255,0):L4X.Image= "g_g5kplayfield3"
      L5X.Intensity=200:L5X.Color=RGB(0,128,0):L5X.ColorFull=RGB(0,255,0):L5X.Image= "g_g5kplayfield3"
      L1.Intensity=20:L1.Color=RGB(177,170,82):L1.ColorFull=RGB(177,170,82):L1.Image= "g_g5kplayfield3"
      L2.Intensity=20:L2.Color=RGB(177,170,82):L2.ColorFull=RGB(177,170,82):L2.Image= "g_g5kplayfield3"
      L3.Intensity=20:L3.Color=RGB(177,170,82):L3.ColorFull=RGB(177,170,82):L3.Image= "g_g5kplayfield3"
      L4.Intensity=20:L4.Color=RGB(177,170,82):L4.ColorFull=RGB(177,170,82):L4.Image= "g_g5kplayfield3"
      L5.Intensity=20:L5.Color=RGB(177,170,82):L5.ColorFull=RGB(177,170,82):L5.Image= "g_g5kplayfield3"
      L6.Intensity=20:L6.Color=RGB(177,170,82):L6.ColorFull=RGB(177,170,82):L6.Image= "g_g5kplayfield3"
      L7.Intensity=20:L7.Color=RGB(177,170,82):L7.ColorFull=RGB(177,170,82):L7.Image= "g_g5kplayfield3"
      L8.Intensity=20:L8.Color=RGB(177,170,82):L8.ColorFull=RGB(177,170,82):L8.Image= "g_g5kplayfield3"
      L9.Intensity=20:L9.Color=RGB(177,170,82):L9.ColorFull=RGB(177,170,82):L9.Image= "g_g5kplayfield3"
      LT1.Intensity=20:LT1.Color=RGB(223,80,0):LT1.ColorFull=RGB(255,226,128):LT1.Image= "g_g5kplayfield3"
      LT2.Intensity=20:LT2.Color=RGB(223,80,0):LT2.ColorFull=RGB(255,226,128):LT2.Image= "g_g5kplayfield3"
      LT3.Intensity=20:LT3.Color=RGB(223,80,0):LT3.ColorFull=RGB(255,226,128):LT3.Image= "g_g5kplayfield3"
      LT4.Intensity=20:LT4.Color=RGB(223,80,0):LT4.ColorFull=RGB(255,226,128):LT4.Image= "g_g5kplayfield3"
      LT5.Intensity=20:LT5.Color=RGB(223,80,0):LT5.ColorFull=RGB(255,226,128):LT5.Image= "g_g5kplayfield3"
      LT6.Intensity=20:LT6.Color=RGB(223,80,0):LT6.ColorFull=RGB(255,226,128):LT6.Image= "g_g5kplayfield3"
      L10.Intensity=20:L10.Color=RGB(177,170,82):L10.ColorFull=RGB(177,170,82):L10.Image= "g_g5kplayfield3"
      L20.Intensity=20:L20.Color=RGB(177,170,82):L20.ColorFull=RGB(177,170,82):L20.Image= "g_g5kplayfield3"
      L10K.Intensity=300:L10K.Color=RGB(255,128,0):L10K.ColorFull=RGB(255,183,111):L10K.Image= "g_g5kplayfield3"
      L30K.Intensity=300:L30K.Color=RGB(255,128,0):L30K.ColorFull=RGB(255,183,111):L30K.Image= "g_g5kplayfield3"
      L50K.Intensity=200:L50K.Color=RGB(255,128,0):L50K.ColorFull=RGB(255,183,111):L50K.Image= "g_g5kplayfield3"
      LFire1.Intensity=10:LFire1.Color=RGB(210,0,0):LFire1.ColorFull=RGB(255,255,255):LFire1.Image= "g_g5kplayfield3"
      LFire2.Intensity=10:LFire2.Color=RGB(210,0,0):LFire2.ColorFull=RGB(255,255,255):LFire2.Image= "g_g5kplayfield3"
      LPower1.Intensity=10:LPower1.Color=RGB(0,0,255):LPower1.ColorFull=RGB(132,132,255):LPower1.Image= "g_g5kplayfield3"
      LPower2.Intensity=10:LPower2.Color=RGB(0,0,255):LPower2.ColorFull=RGB(132,132,255):LPower2.Image= "g_g5kplayfield3"
      LBlueTop.Intensity=5:LBlueTop.Color=RGB(0,0,255):LBlueTop.ColorFull=RGB(132,132,255):LBlueTop.Image="g_g5kplayfield3"
      LSpinner.Intensity=200:LSpinner.Color=RGB(149,30,0):LSpinner.ColorFull=RGB(255,43,43):LSpinner.Image= "g_g5kplayfield3"
      LLeftHole.Intensity=50:LLeftHole.Color=RGB(0,128,0):LLeftHole.ColorFull=RGB(0,255,0): LLeftHole.Image= "g_g5kplayfield3"
      LRightHole.Intensity=100:LRightHole.Color=RGB(0,128,0):LRightHole.ColorFull=RGB(0,255,0): LRightHole.Image= "g_g5kplayfield3"
      LBackHole.Intensity=100:LBackHole.Color=RGB(0,128,0):LBackHole.ColorFull=RGB(0,255,0): LBackHole.Image= "g_g5kplayfield3"
      LBlueMiddle.Intensity=5:LBlueMiddle.Color=RGB(0,0,255):LBlueMiddle.ColorFull=RGB(132,132,255):LBlueMiddle.Image= "g_g5kplayfield3"
      LBlueBottom.Intensity=5:LBlueBottom.Color=RGB(0,0,255):LBlueBottom.ColorFull=RGB(132,132,255):LBlueBottom.Image= "g_g5kplayfield3"
      LExtraBall.Intensity=200:LExtraBall.Color=RGB(255,55,55):LExtraBall.ColorFull=RGB(255,88,17):LExtraBall.Image= "g_g5kplayfield3"
      LLeftBlue1000.Intensity=50:LLeftBlue1000.Color=RGB(0,0,255):LLeftBlue1000.ColorFull=RGB(132,132,255):LLeftBlue1000.Image=" g_g5kplayfield3"
      LRightBlue1000.Intensity=50:LRightBlue1000.Color=RGB(0,0,255):LRightBlue1000.ColorFull=RGB(132,132,255):LRightBlue1000.Image= "g_g5kplayfield3"
      LShieldOn.Intensity=300:LShieldOn.Color=RGB(115,28,23):LShieldOn.ColorFull=RGB(255,156,128):LShieldOn.Image= "g_g5kplayfield3"
      LLeftSpecial.Intensity=5:LLeftSpecial.Color=RGB(255,26,0):LLeftSpecial.ColorFull=RGB(255,254,253):LLeftSpecial.Image= "g_g5kplayfield3"
      LRightSpecial.Intensity=5:LRightSpecial.Color=RGB(255,26,0):LRightSpecial.ColorFull=RGB(255,254,253):LRightSpecial.Image= "g_g5kplayfield3"
      LShootAgain.Intensity=5:LShootAgain.Color=RGB(55,3,0):LShootAgain.ColorFull=RGB(255,198,198):LShootAgain.Image= "g_g5kplayfield3"
      LShootAgainB.Intensity=50:LShootAgainB.Color=RGB(55,3,0):LShootAgainB.ColorFull=RGB(255,198,198)
  Case 1 'LED
      LF.Intensity=300:LF.Color=RGB(0,132,0):LF.ColorFull=RGB(94,255,94):LF.Image= "g_g5kplayfield3"
      LI.Intensity=300:LI.Color=RGB(0,132,0):LI.ColorFull=RGB(94,255,94):LI.Image= "g_g5kplayfield3"
      LR.Intensity=300:LR.Color=RGB(0,132,0):LR.ColorFull=RGB(94,255,94):LR.Image= "g_g5kplayfield3"
      LE.Intensity=300:LE.Color=RGB(0,132,0):LE.ColorFull=RGB(94,255,94):LE.Image= "g_g5kplayfield3"
      L2X.Intensity=600:L2X.Color=RGB(0,79,0):L2X.ColorFull=RGB(104,255,83):L2X.Image= "g_g5kplayfield3"
      L3X.Intensity=600:L3X.Color=RGB(0,79,0):L3X.ColorFull=RGB(104,255,83):L3X.Image= "g_g5kplayfield3"
      L4X.Intensity=600:L4X.Color=RGB(0,79,0):L4X.ColorFull=RGB(104,255,83):L4X.Image= "g_g5kplayfield3"
      L5X.Intensity=600:L5X.Color=RGB(0,79,0):L5X.ColorFull=RGB(104,255,83):L5X.Image= "g_g5kplayfield3"
      L1.Intensity=300:L1.Color=RGB(207,203,148):L1.ColorFull=RGB(255,255,128):L1.Image= "g_g5kplayfield3"
      L2.Intensity=300:L2.Color=RGB(207,203,148):L2.ColorFull=RGB(255,255,128):L2.Image= "g_g5kplayfield3"
      L3.Intensity=300:L3.Color=RGB(207,203,148):L3.ColorFull=RGB(255,255,128):L3.Image= "g_g5kplayfield3"
      L4.Intensity=300:L4.Color=RGB(207,203,148):L4.ColorFull=RGB(255,255,128):L4.Image= "g_g5kplayfield3"
      L5.Intensity=300:L5.Color=RGB(207,203,148):L5.ColorFull=RGB(255,255,128):L5.Image= "g_g5kplayfield3"
      L6.Intensity=300:L6.Color=RGB(207,203,148):L6.ColorFull=RGB(255,255,128):L6.Image= "g_g5kplayfield3"
      L7.Intensity=300:L7.Color=RGB(207,203,148):L7.ColorFull=RGB(255,255,128):L7.Image= "g_g5kplayfield3"
      L8.Intensity=300:L8.Color=RGB(207,203,148):L8.ColorFull=RGB(255,255,128):L8.Image= "g_g5kplayfield3"
      L9.Intensity=300:L9.Color=RGB(207,203,148):L9.ColorFull=RGB(255,255,128):L9.Image= "g_g5kplayfield3"
      LT1.Intensity=600:LT1.Color=RGB(223,84,0):LT1.ColorFull=RGB(255,255,128):LT1.Image= "g_g5kplayfield3"
      LT2.Intensity=600:LT2.Color=RGB(223,84,0):LT2.ColorFull=RGB(255,255,128):LT2.Image= "g_g5kplayfield3"
      LT3.Intensity=600:LT3.Color=RGB(223,84,0):LT3.ColorFull=RGB(255,255,128):LT3.Image= "g_g5kplayfield3"
      LT4.Intensity=600:LT4.Color=RGB(223,84,0):LT4.ColorFull=RGB(255,255,128):LT4.Image= "g_g5kplayfield3"
      LT5.Intensity=600:LT5.Color=RGB(223,84,0):LT5.ColorFull=RGB(255,255,128):LT5.Image= "g_g5kplayfield3"
      LT6.Intensity=600:LT6.Color=RGB(223,84,0):LT6.ColorFull=RGB(255,255,128):LT6.Image= "g_g5kplayfield3"
      L10.Intensity=600:L10.Color=RGB(207,203,148):L10.ColorFull=RGB(255,255,128):L10.Image= "g_g5kplayfield3"
      L20.Intensity=600:L20.Color=RGB(207,203,148):L20.ColorFull=RGB(255,255,128):L20.Image= "g_g5kplayfield3"
      L10K.Intensity=1000:L10K.Color=RGB(255,51,0):L10K.ColorFull=RGB(255,148,40):L10K.Image= "g_g5kplayfield3"
      L30K.Intensity=1000:L30K.Color=RGB(255,51,0):L30K.ColorFull=RGB(255,148,40):L30K.Image= "g_g5kplayfield3"
      L50K.Intensity=1000:L50K.Color=RGB(255,51,0):L50K.ColorFull=RGB(255,148,40):L50K.Image= "g_g5kplayfield3"
      LFire1.Intensity=1000:LFire1.Color=RGB(255,34,34):LFire1.ColorFull=RGB(255,164,72):LFire1.Image= "g_g5kplayfield3"
      LFire2.Intensity=3000:LFire2.Color=RGB(255,34,34):LFire2.ColorFull=RGB(255,164,72):LFire2.Image= "g_g5kplayfield3"
      LPower1.Intensity=1000:LPower1.Color=RGB(0,0,128):LPower1.ColorFull=RGB(202,228,255):LPower1.Image= "g_g5kplayfield3"
      LPower2.Intensity=300:LPower2.Color=RGB(0,0,128):LPower2.ColorFull=RGB(202,228,255):LPower2.Image= "g_g5kplayfield3"
      LBlueTop.Intensity=100:LBlueTop.Color=RGB(0,0,255):LBlueTop.ColorFull=RGB(30,30,255):LBlueTop.Image= "g_g5kplayfield3"
      LSpinner.Intensity=600:LSpinner.Color=RGB(166,0,0):LSpinner.ColorFull=RGB(249,37,0):LSpinner.Image= "g_g5kplayfield3"
      LLeftHole.Intensity=3000:LLeftHole.Color=RGB(0,128,0):LLeftHole.ColorFull=RGB(0,255,0):LLeftHole.Image= "g_g5kplayfield3"
      LRightHole.Intensity=800:LRightHole.Color=RGB(0,128,0):LRightHole.ColorFull=RGB(0,255,0):LRightHole.Image= "g_g5kplayfield3"
      LBackHole.Intensity=800:LBackHole.Color=RGB(0,128,0):LBackHole.ColorFull=RGB(0,255,0):LBackHole.Image= "g_g5kplayfield3"
      LBlueMiddle.Intensity=100:LBlueMiddle.Color=RGB(0,0,255):LBlueMiddle.ColorFull=RGB(30,30,255):LBlueMiddle.Image= "g_g5kplayfield3"
      LBlueBottom.Intensity=100:LBlueBottom.Color=RGB(0,0,255):LBlueBottom.ColorFull=RGB(30,30,255):LBlueBottom.Image= "g_g5kplayfield3"
      LExtraBall.Intensity=400:LExtraBall.Color=RGB(128,0,0):LExtraBall.ColorFull=RGB(251,32,0):LExtraBall.Image= "g_g5kplayfield3"
      LLeftBlue1000.Intensity=150:LLeftBlue1000.Color=RGB(0,0,128):LLeftBlue1000.ColorFull=RGB(0,128,255):LLeftBlue1000.Image= "g_g5kplayfield3"
      LRightBlue1000.Intensity=150:LRightBlue1000.Color=RGB(0,0,128):LRightBlue1000.ColorFull=RGB(0,128,255):LRightBlue1000.Image= "g_g5kplayfield3"
      LShieldOn.Intensity=3000:LShieldOn.Color=RGB(115,28,23):LShieldOn.ColorFull=RGB(255,156,128):LShieldOn.Image= "g_g5kplayfield3"
      LLeftSpecial.Intensity=400:LLeftSpecial.Color=RGB(128,64,0):LLeftSpecial.ColorFull=RGB(255,128,0):LLeftSpecial.Image= "g_g5kplayfield3"
      LRightSpecial.Intensity=150:LRightSpecial.Color=RGB(128,64,0):LRightSpecial.ColorFull=RGB(255,128,0):LRightSpecial.Image= "g_g5kplayfield3"
      LShootAgain.Intensity=50:LShootAgain.Color=RGB(113,0,0):LShootAgain.ColorFull=RGB(244,184,0):LShootAgain.Image= "g_g5kplayfield3"
      LShootAgainB.Intensity=100:LShootAgainB.Color=RGB(55,3,0):LShootAgainB.ColorFull=RGB(255,198,198)
  End Select
End Sub

dim tiltcount, matchcount

Sub LightCopy_TIMER

    LBackHole1.State = LBackHole.State
    LBlueBottomB.State = LBlueBottom.State
    LBlueMiddleB.State = LBlueMiddle.State
    LBlueTopB.State = LBlueTop.State
    LShootAgainB.State = LShootAgain.State
    LBumperTopLeftB.State = LBumperTopLeft.State
    LBumperTopRightB.State = LBumperTopRight.State
    LBumperBottomRightB.State = LBumperBottomRight.State
    LBumperBottomLeftB.State = LBumperBottomLeft.State

  If LBumperTopLeft.State=1 Then
             BumpCap1.image ="g_BUMPERCAP1_LIGHTON"
             BR1_7.image = "g_BUMPERRING1_LIGHTON"
             BumpSkirt1.image =  "g_BUMPERSKIRT1_LIGHTON"
       Else
             BumpCap1.image = "g_BUMPERCAP1_LIGHTOFF"
             BR1_7.image = "g_BUMPERRING1_LIGHTOFF"
             BumpSkirt1.image =  "g_BUMPERSKIRT1_LIGHTOFF"
        End If

  If LBumperTopRight.State=1 Then
             BumpCap2.image ="g_BUMPERCAP2_LIGHTON"
       BR2_7.image = "g_BUMPERRING2_LIGHTON"
             BumpSkirt2.image =  "g_BUMPERSKIRT2_LIGHTON"
       Else
       BumpCap2.image = "g_BUMPERCAP2_LIGHTOFF"
       BR2_7.image = "g_BUMPERRING2_LIGHTOFF"
       BumpSkirt2.image =  "g_BUMPERSKIRT2_LIGHTOFF"
    End If

  If LBumperBottomRight.State=1 Then
        BumpCap3.image ="g_BUMPERCAP3_LIGHTON"
        BR3_7.image = "g_BUMPERRING3_LIGHTON"
        BumpSkirt3.image =  "g_BUMPERSKIRT3_LIGHTON"
  Else
        BumpCap3.image = "g_BUMPERCAP3_LIGHTOFF"
        BR3_7.image = "g_BUMPERRING3_LIGHTOFF"
        BumpSkirt3.image =  "g_BUMPERSKIRT3_LIGHTOFF"
    End If

  If LBumperBottomLeft.State=1 Then
         BumpCap4.image ="g_BUMPERCAP4_LIGHTON"
         BR4_7.image = "g_BUMPERRING4_LIGHTON"
         BumpSkirt4.image =  "g_BUMPERSKIRT4_LIGHTON"
   Else
    BumpCap4.image = "g_BUMPERCAP4_LIGHTOFF"
    BR4_7.image = "g_BUMPERRING4_LIGHTOFF"
    BumpSkirt4.image =  "g_BUMPERSKIRT4_LIGHTOFF"
     End If

'        ________  __  ___  _____    ____      __  _______  ______________  _  __
'*****  / __/ __ \/ / / / |/ / _ \  / __/___  /  |/  / __ \/_  __/  _/ __ \/ |/ /*****
'***** _\ \/ /_/ / /_/ /    / // /  > _/_ _/ / /|_/ / /_/ / / / _/ // /_/ /    / *****
'*****/___/\____/\____/_/|_/____/  |_____/  /_/  /_/\____/ /_/ /___/\____/_/|_/  *****


   If BGLmatch.State = 0 Then
    isGameOver = 0
    myPlayingEndGameSnd = False
    GameOverTimer.Enabled = 0
    'KBallSaveKicker.Enabled=1
    matchcount = 0
   Elseif isGameOver = 0 Then
    matchcount = matchcount + me.interval
    if matchcount > 1000 then
      Trigger001.enabled=1
      isGameOver = 1
      'debug.print matchcount
      If AIon  Then   ToggleAI(0)
      KBallSaveKicker.Enabled=0
      If  Not myPlayingEndGameSnd and myPrefs_PlayMusic then
        myPlayingEndGameSnd = True
                                TWall1Count = 0':TWall2Count = 0
        GameOverTimer.enabled = 1
        TriggerLF.enabled=0
        TriggerRF.enabled=0
        TWall.enabled=0
      End If
    end if
  End If

  If BGLtilt.State=1 Then
    tiltcount = tiltcount + me.interval
    if tiltcount > 1000 and tiltcount < 2000 then
      tiltcount = 2000
      'debug.print tiltcount
      ToggleAI(0)
      isTilted = 1
    end if
  Else
    isTilted = 0
    tiltcount = 0
  End If

  If isGameOver or isTilted Then
    Bumper1.threshold = 100
    Bumper2.threshold = 100
    Bumper3.threshold = 100
    Bumper4.threshold = 100

    LeftSlingShot.SlingShotThreshold = 100
    RightSlingShot.SlingShotThreshold =100
  Else
    Bumper1.threshold = 0.5
    Bumper2.threshold = 0.5
    Bumper3.threshold = 0.5
    Bumper4.threshold = 0.5

    LeftSlingShot.SlingShotThreshold = 5
    RightSlingShot.SlingShotThreshold = 5
  End If

End Sub

Dim isGameOver, isTilted
isGameOver = 0
isTilted = 0

Sub GameOverTimer_Timer
        gamestarted = 0
call MyShowInstructionCard
call myLightSeq
playsound "_air-release"
Player1Light.Duration 2, 2000, 1
Player2Light.Duration 2, 1900, 1
Player3Light.Duration 2, 1800, 1
Player4Light.Duration 2, 1700, 1
LightDifficulty0.Duration 2, 2100, 1
LightDifficulty1.Duration 2, 2200, 1
LightDifficulty2.Duration 2, 2300, 1
LightDifficulty3.Duration 2, 2400, 1
LightDifficulty4.Duration 2, 2500, 2
  Select Case Int(Rnd()*4)
    Case 0
      PlaySound "__SpaceWaltz(music)"
    Case 1
      PlaySound "__2001(music)"
      stopsound "__SpaceAmbient(music)"
    Case 2
      Playsound "__My God, its full of stars(music)"
                Case 3
                        Playsound "__Ode To The Sun (Outro)"
  End Select
  BGLmatch.State = 0
  myPlayingEndGameSnd = False

If MyPrefs_DynamicCards = 1 Then
Playsound "CardChange"
Light004.Duration 2, 80, 0
Light005.Duration 2, 80, 0
  End If
GameOverTimer.Enabled = 0
End Sub

 Sub myAdjustTableToPrefs 'AXS
    If myPrefs_PlayMusic then
        PlaySound "__SpaceAmbient(music)"
  Select Case Int(Rnd()*16)
    Case 0
      PlaySound "_Something Wonderful"
      Halspeak1 = 0
    Case 1
      PlaySound "_MissionControl"
      Halspeak1 = 0
        Case 2
      Playsound "_Hal-9000Introduction"
                        Halspeak1 = 0
        Case 3
      Playsound "__RequiemIntro(music)"
                Case 4
      Playsound ""
                Case 5
      Playsound ""
                Case 6
      Playsound ""
                Case 7
      Playsound ""
                Case 8
      Playsound ""
                Case 9
      Playsound ""
                Case 10
      Playsound "Good evening Dave"
                Case 11
      Playsound "Good evening Dave"
                Case 12
      Playsound "Good evening Dave"
                Case 13
      Playsound "Good evening Dave"
                Case 14
      Playsound "Good evening Dave"
                Case 15
      Playsound "Good evening Dave"
  End Select
    End If
 End Sub

'********************************************************************

 Sub Glass_Hit
   If myPrefs_GlassHit = 2 Then
    Select Case Int(Rnd()*4)
     Case 0
       Playsound "AXSGlassHit6"
       Playsound "AXSGlassHit6"
      Case 1
       Playsound "AXSGlassHit3"
       Playsound "AXSGlassHit3"
      Case 2
       Playsound "AXSGlassHit3"
      Case 3
       Playsound "AXSGlassHit6"
      End Select
       BB1Timer.Enabled = 1
     End If
   If myPrefs_GlassHit = 1 Then
    Select Case Int(Rnd()*6)
      Case 0
      Case 1
      Case 2
       Playsound "AXSGlassHit3"
      Case 3
       Playsound "AXSGlassHit6"
      Case 4
      Case 5
      End Select
       BB1Timer.Enabled = 1
     End If
   If myPrefs_GlassHit = 0 Then
       BB1Timer.Enabled = 1
     End If
   End Sub

   Sub BB1Timer_Timer
     Playsound "_ball_bounce"
     BB1Timer.Enabled = 0
    End Sub

'*********************************************************************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    GateFlap.roty= 0 - Gate3.Currentangle / 1
    SLWGate.rotx= -30 - Gate2.Currentangle / -1
    ULWGate.rotx= -30 - Gate1.Currentangle / -1
  Select case myPrefs_DynamicBackground
     Case 0
     StarsTimer.enabled = 0
     Case 1
     StarsTimer.enabled = 1
  End Select
   End Sub

Sub StarsTimer_Timer()
  if myPrefs_FactorySetting = 0 and isMultiball = 0 then
    ItsFullOfStars.RotZ = ItsFullOfStars.RotZ + .003
    Jupiter.RotZ = Jupiter.RotZ + .0008
  end if
  if myPrefs_FactorySetting = 0 and isMultiball = 1 then
    ItsFullOfStars.RotZ = ItsFullOfStars.RotZ + .016
    Jupiter.RotZ = Jupiter.RotZ + .0096
  end if
End Sub

'************** Logic Panel Motor ******************


 Sub TimerMovePanel_Timer()
       LogicBanner.TransX = LogicBanner.TransX + (4* mypanelDirection)
  If  LogicBanner.TransX < -392  Then
                call myTurnPanelOnOff(1)
    TimerMovePanel.Enabled = False
    mypanelDirection = mypanelDirection * -1
        End If
  If  LogicBanner.TransX > 392  Then
    TimerMovePanel.Enabled = False
    mypanelDirection = mypanelDirection * -1
    LogicBanner.TransX = 0
    call myTurnPanelOnOff(0)
         End If

 End Sub



'**********************************************************************
'************************* G5K Ball Shadows ***************************
'**********************************************************************

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 5
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 5
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

Sub L20_Init()

End Sub

Sub L20_Timer()

End Sub


'****************************************
'*** Accompanying AI Sounds & Scripts ***
'****************************************


Sub TWall_hit
  If myPrefs_SoundEffects Then
    If AIOn = 1 Then
   Select Case Int(Rnd()*9)
    Case 0
      Playsound "_TOASTY"
       CreditLight1.State = 2:CreditLight.State = 2
    Case 1
      Playsound "_TOASTYE"
      CreditLight1.State = 2:CreditLight.State = 2
    Case 2
      Playsound "_TOASTYG":Playsound "_TOASTYG"
      CreditLight1.State = 2:CreditLight.State = 2
    Case 3
      Playsound "_Fatality":Playsound "_Fatality":Playsound "_Fatality"
      CreditLight1.State = 2:CreditLight.State = 2
    Case 4
      Playsound "_Robotron":Playsound "_Robotron"
      CreditLight1.State = 2:CreditLight.State = 2
    Case 5
      Playsound "_Fatality":Playsound "_Fatality":Playsound "_Fatality"
      CreditLight1.State = 2:CreditLight.State = 2
    Case 6
      Playsound "_Robotron":Playsound "_Robotron"
      CreditLight1.State = 2:CreditLight.State = 2
    Case 7
      Playsound ""
      CreditLight1.State = 2:CreditLight.State = 2
    Case 8
      Playsound "_Glass Breakin":Playsound "_Glass Breakin"
      CreditLight1.State = 2:CreditLight.State = 2
   End Select
AITaunt = 1
  End If
 End If
End Sub

Sub TWall001_hit
  If myPrefs_SoundEffects and myPrefs_Difficulty = 4 and AIPlaying = 1 and AIOn = 0 Then
    Select Case Int(Rnd()*9)
    Case 0
      Playsound "Laugh":Playsound "Laugh"
    Case 1
      Playsound "Laugh2":Playsound "Laugh2"
    Case 2
      Playsound "Laugh2":Playsound "Laugh2"
    Case 3
      Playsound "Laugh3":Playsound "Laugh3"
    Case 4
      Playsound "Laugh4":Playsound "Laugh4":Playsound "Laugh4"
    Case 5
      playsound ""
    Case 6
      Playsound "Laugh5":Playsound "Laugh5"
    Case 7
      Playsound "Laugh6":Playsound "Laugh6"
    Case 8
      Playsound "Laugh7":Playsound "Laugh7"
    End Select
  End If
End Sub

Dim TWall1Count, TWall2Count, AIPlaying
TWall1Count = 0
TWall2Count = 0
AIPlaying = 0

 Sub TWall1_hit 'Hal's turn
    If myPrefs_SoundEffects And AIOn = 1 Then
Light001.State = 1
Light002.State = 0
          Select Case TWall1Count
    Case 0
    If myPrefs_Difficulty = 4 Then
                        Playsound "YouWillLose2":Playsound "YouWillLose2":Playsound "YouWillLose2":Playsound "YouWillLose2":Playsound "YouWillLose2":Halspeak = 1
    else
                   Select Case Int(Rnd()*2)
                Case 0
                        Playsound"_FirstLesson2":Playsound"_FirstLesson2":Playsound"_FirstLesson2":Playsound"_FirstLesson2":Playsound"_FirstLesson2":Halspeak = 1
                Case 1
                        Playsound "Good evening Dave":Playsound "Good evening Dave":Playsound "Good evening Dave"
                End Select
    end if
    Case 1
      Playsound "_RobotronLevelUp":Playsound "_RobotronLevelUp"
    Case 2
      playsound "_GoingWell":playsound "_GoingWell":playsound "_GoingWell":playsound "_GoingWell":Halspeak = 1:
    Case 3
      Playsound "_WorkingWithPeople":Playsound "_WorkingWithPeople":Playsound "_WorkingWithPeople":Halspeak = 1
                Case 4
                        'Playsound "Everything is running smoothly":Playsound "Everything is running smoothly":Playsound "Everything is running smoothly":Playsound "Everything is running smoothly":Halspeak = 1
    Case 5
            Playsound "_iampushingmyself":Playsound "_iampushingmyself":Playsound "_iampushingmyself":Halspeak = 1
    Case 6
      playsound "Sorry about this":playsound "Sorry about this":playsound "Sorry about this":playsound "Sorry about this":Halspeak = 1
    Case 7
      playsound "_Concerned":playsound "_Concerned":playsound "_Concerned":playsound "_Concerned":Halspeak = 1
    Case 8
      'playsound"_OpenthedoorHal":playsound"_OpenthedoorHal":playsound"_OpenthedoorHal"
    Case 9
       playsound"_SorryDave":playsound"_SorryDave":playsound"_SorryDave":Halspeak = 1
    Case 10
       Playsound "_DisconnectMe":Playsound "_DisconnectMe":Playsound "_DisconnectMe":Halspeak = 1
    Case 11
       'Playsound "_TooImportant":Playsound "_TooImportant":Playsound "_TooImportant":Halspeak = 1
                Case 12
                         'playsound "_DeactivatingAutoPlay":playsound "_DeactivatingAutoPlay":playsound "_DeactivatingAutoPlay"
                Case 13
                         'playsound "_DaveStop":playsound "_DaveStop":playsound "_DaveStop":Halspeak = 1
                Case 14
                         Playsound "_I'maHall9000":Playsound "_I'maHall9000":Playsound "_I'maHall9000"
                Case 15
                         Playsound"_Singasong":Playsound"_Singasong":Playsound"_Singasong":Halspeak = 1
                Case 16
                         Playsound "_Daisy":Playsound "_Daisy":Playsound "_Daisy":Playsound "_Daisy":Halspeak = 1
                Case 17
                         Playsound"_VoiceTest2":Playsound"_VoiceTest2":Playsound"_VoiceTest2":Halspeak = 1
                Case 18
                         Playsound"_VoiceTest":Playsound"_VoiceTest":Playsound"_VoiceTest":Halspeak = 1
                Case 19
      Playsound "Shall we play a game":Playsound "Shall we play a game"':Playsound "Shall we play a game"
    Case 20
      Playsound "_RobotronLevelUp1":Playsound "_RobotronLevelUp1"
      End Select
      TWall1Count = TWall1Count + 1
    If TWall1Count > 20 Then TWall1Count = 0
      If HalLight2.State = 2 Then
         AIPlaying = 1
      End If
       End If
         'End if
 End Sub

 Sub TWall2_hit 'Player's turn
    CreditLight3.State = 2
    CreditLight3Timer.enabled = 1
    If myPrefs_SoundEffects  Then
       If  AIPlaying = 1 And HalLight2.State = 0 Then
Light001.State = 0
Light002.State = 1
    If myPrefs_Difficulty = 4 Then
          Select Case TWall2Count
    'Case 0
      'Playsound"_FirstLesson2":Playsound"_FirstLesson2":Playsound"_FirstLesson2":Playsound"_FirstLesson2":Playsound"_FirstLesson2":Halspeak = 1
    Case 1
      'Playsound "_RobotronLevelUp":Playsound "_RobotronLevelUp"
    Case 2
      'playsound "_GoingWell":playsound "_GoingWell":playsound "_GoingWell":Halspeak = 1:
    Case 3
      'Playsound "_WorkingWithPeople":Playsound "_WorkingWithPeople":Playsound "_WorkingWithPeople":Halspeak = 1
    Case 4
      Playsound "You have improved":Playsound "You have improved":Playsound "You have improved":Playsound "You have improved":Halspeak = 1
    Case 5
      playsound "Everything is running smoothly":playsound "Everything is running smoothly":playsound "Everything is running smoothly":playsound "Everything is running smoothly"
    Case 6
      playsound "Ask you a question":playsound "Ask you a question":playsound "Ask you a question":Halspeak = 1
    Case 7
      playsound"Forgive me for asking":playsound"Forgive me for asking":playsound"Forgive me for asking"
    Case 8
      playsound"Second thoughts":playsound"Second thoughts":playsound"Second thoughts":Halspeak = 1
    Case 9
      Playsound "My own concern":Playsound "My own concern":Playsound "My own concern":Playsound "My own concern":Halspeak = 1
    Case 10
      Playsound "_TooImportant":Playsound "_TooImportant":Playsound "_TooImportant":Halspeak = 1
    Case 11
      Playsound "_DeactivatingAutoPlay":playsound "_DeactivatingAutoPlay":playsound "_DeactivatingAutoPlay"
    Case 12
      Playsound "_DaveStop":playsound "_DaveStop":playsound "_DaveStop":Halspeak = 1
    Case 13
      'Playsound "_I'maHall9000":Playsound "_I'maHall9000":Playsound "_I'maHall9000"
    Case 14
      'Playsound"_Singasong":Playsound"_Singasong":Playsound"_Singasong":Halspeak = 1
    Case 15
      'Playsound "_Daisy":Playsound "_Daisy":Playsound "_Daisy":Playsound "_Daisy":Halspeak = 1
    Case 16
      'Playsound"_VoiceTest2":Playsound"_VoiceTest2":Playsound"_VoiceTest2":Halspeak = 1
    Case 17
      'Playsound"_VoiceTest":Playsound"_VoiceTest":Playsound"_VoiceTest":Halspeak = 1
    Case 18
      'Playsound "_RobotronLevelUp1":Playsound "_RobotronLevelUp1"
      End Select
      TWall2Count = TWall2Count + 1
    If TWall2Count > 18 Then TWall2Count = 0
       End If
        End If
    End If
End Sub


'********************************** TriggerLF

 Sub TriggerLF_Hit
                Light8.Color=RGB(0,255,0)
                Light8.Colorfull=RGB(0,255,0)
    'Light8.Intensity = 1000
    If myPrefs_SoundEffects Then
     Select Case Int(Rnd()*11)
       Case 0
        Playsound "_laser-blast-1"
       Case 1
        PlaySound "_laser-blast-2"
       Case 2
        Playsound "_laser-blast-3"
       Case 3
        playsound "_laser-blast-4"
       Case 4
        Playsound "_laser-blast-5"
       Case 5
        Playsound "_TIE fighter2"
       Case 6
        Playsound "_laser-blast-6"
        Playsound "_laser-blast-7"
       Case 7
        Playsound "_laser-blast-8"
       Case 8
        PlaySound "_laser-blast-9"
       Case 9
        Playsound "_laser-blast-10"
       Case 10
        Playsound "_TIE fighter"
      End Select
    Select Case Int(Rnd()*3)
       Case 0
        'ItsFullOfStars.Material = "StarFieldBright"
        StrobeLight.State=1
        StrobeLightTimer.Enabled=1
       Case 1
        'ItsFullOfStars.Material = "StarFieldBright"
        StrobeLight3.State=1
        StrobeLightTimer.Enabled=1
       Case 2
        'ItsFullOfStars.Material = "StarFieldBright"
        StrobeLight4.State=1
        StrobeLightTimer.Enabled=1
      End Select
     End If
    vpmkeydown(LeftFlipperKey)
        'TriggerLF.enabled=0
        TimerLF.Enabled=1

 End Sub

 Sub TimerLF_Timer
  vpmkeyup(LeftFlipperKey)
        'TriggerLF.enabled=1
  TimerLF.interval = 200
        TimerLF.enabled=0
        ZoneLeft.Enabled=1
 End Sub


 Sub TriggerLF_UnHit
                Light8.Color=RGB(255,0,0)
                Light8.Colorfull=RGB(255,0,0)
    'Light8.Intensity = 10
End Sub

'*************************************** 'TriggerRF

 Sub TriggerRF_Hit
                Light9.Color=RGB(0,255,0)
                Light9.Colorfull=RGB(0,255,0)
    'Light9.Intensity = 1000
  'vpmtimer.pulsesw(cFlipperRightSW)
   If myPrefs_SoundEffects Then
     Select Case Int(Rnd()*11)
       Case 0
        Playsound "_laser-blast-1"
       Case 1
        PlaySound "_laser-blast-2"
       Case 2
        Playsound "_laser-blast-3"
       Case 3
        playsound "_laser-blast-4"
       Case 4
        Playsound "_laser-blast-5"
       Case 5
        Playsound "_TIE fighter2"
       Case 6
        Playsound "_laser-blast-6"
        Playsound "_laser-blast-7"
       Case 7
        Playsound "_laser-blast-8"
       Case 8
        PlaySound "_laser-blast-9"
       Case 9
        Playsound "_laser-blast-10"
       Case 10
        Playsound "_TIE fighter"
      End Select
    Select Case Int(Rnd()*3)
       Case 0
       ' ItsFullOfStars.Material = "StarFieldBright"
        StrobeLight.State=1
        StrobeLightTimer.Enabled=1
       Case 1
        'ItsFullOfStars.Material = "StarFieldBright"
        StrobeLight3.State=1
        StrobeLightTimer.Enabled=1
       Case 2
        'ItsFullOfStars.Material = "StarFieldBright"
        StrobeLight4.State=1
        StrobeLightTimer.Enabled=1
      End Select
     End If

   vpmkeydown(RightFlipperKey)
        'TriggerRF.enabled=0
           TimerRF.Enabled=1
 End Sub

Sub TriggerRF_UnHit
                Light9.Color=RGB(255,0,0)
                Light9.Colorfull=RGB(255,0,0)
    'Light9.Intensity = 10
End Sub

 Sub  TimerRF_Timer 'here
  vpmkeyup(RightFlipperKey)
  'TriggerRF.enabled=1
  TimerRF.interval=200
  TimerRF.enabled=0
  ZoneRight.Enabled=1
 End Sub

 Sub TriggerAutoFlipPlunger_Hit()
        LowerZoneState=0
  CreditLight1.State = 2
  CreditLight.State = 0
  Plunger.Pullback()
  PlaysoundAtVol "PlungerPull", ActiveBall, 1
  TimerAutoFlipPlunger.Enabled = True
 End Sub

 Sub TimerAutoFlipPlunger_Timer()
  Plunger.Fire()
  PlaysoundAtVol "Plunger", Plunger, 1
  TimerAutoFlipPlunger.Enabled = False
  TimerBallStuck.Enabled = True
 End Sub

'
'*************************************** AI Logic *******************************************
'
Dim AIOn
Dim AIOff
Dim CurPlayer, PrevPlayer

Sub ToggleAI(onoff)
  If onoff = 1 Then
    AIOn = 1
    AIOff = 0
    Flasher6.visible = True
    Flasher6Timer.Enabled = 1
    Light12.Duration 2, 80, 0
    Light13.Duration 2, 80, 0
  Else
    AIOn = 0
    AIOff = 1
  End If

  If AIOn = 1  Then
    If myPrefs_SoundEffects and ismultiball = 0 Then
      playsound "_AutoPlayActivate":'playsound "_PlungerPull"
      stopsound "_DeactivatingAutoPlay":stopsound "_button-click":stopsound "_AreYouSure":stopsound "_Concerned"
                        stopsound "_GoingWell":stopsound "_WorkingWithPeople":stopsound "_TooImportant":stopsound "_DisconnectMe":stopsound "_Daisy"
    End If
      vpmkeyup(LeftFlipperKey)
      vpmkeyup(RightFlipperKey)
      'ItsFullOfStars.Material = "StarFieldBright"
      CreditLight.State = 2:CreditLight1.State = 2
      'Hal9000Eye.Image= "hal-9000":HalLight2.State = 2':Hal9000Ring.Visible = 1
      TriggerLF.enabled=1:TriggerRF.enabled=1      '(Flipper Triggers Main AI)
      ZoneLeft.enabled = 1:ZoneRight.enabled = 1   '(Flipper Triggers Cradle Logic)
      ZoneLeft1.enabled = 1:ZoneRight1.enabled = 1
      LowerZoneTrigger.enabled=1
    if ismultiball = 0 then FireF.Enabled=1:FireI.Enabled=1:FireR.Enabled=1:FireE.Enabled=1  '(Lane Change AI)
    if ismultiball = 0 then FireF1.Enabled=1:FireI1.Enabled=1:FireR1.Enabled=1:FireE1.Enabled=1  '(Lane Change AI)
      TWall.enabled = 1:TWall1.enabled = 1 :TWall2.enabled = 0             '(Shoot and Drain Sound Effects Triggers)
      TriggerAutoFlipPlunger.enabled = 1
      StrobeLight.State=1:StrobeLight1.State=1:StrobeLightTimer.Enabled=1  'StrobeLight.State=1:
      Plunger.PullBack:TimerAutoFlipPlunger.Enabled = True
      plunger.mechplunger = false:controller.Switch(84) = 0:controller.Switch(82) = 0
    if myPrefs_SoundEffects = 1 then:   Hal9000Eye.Image= "hal-9000":HalLight2.State = 2:HalLight2.Color=RGB(255,28,28):LF.ColorFull=RGB(255,28,28) end If
    if myPrefs_SoundEffects = 0 then: Hal9000Eye.Image= "hal-9000B":HalLight2.State = 2:HalLight2.Color=RGB(0,0,255):LF.ColorFull=RGB(0,0,255) end If
    If PostLightOnOff = 1 Then
      PostLight.Duration 2, 1200, 0
      playsound"servo_small"
  End If
  Else
    If  AIOff = 1 Then
      playsound "_button-click":stopsound "_AutoPlayActivate"
      'Wall11.Collidable = 0 :Wall12.Collidable = 0 :Wall13.Collidable = 0
      StrobeLight2.State=0:CreditLight.State = 0:CreditLight1.State = 0
      HalLight2.State = 0:Hal9000Eye.Image= "hal-9000off"':Hal9000Ring.Visible = 0
      ZoneLeft.enabled = 0:ZoneRight.enabled = 0
      ZoneLeft1.enabled = 0:ZoneRight1.enabled = 0
                        LowerZoneTrigger.enabled=0
      FireF.Enabled=0:FireI.Enabled=0:FireR.Enabled=0:FireE.Enabled=0
      FireF1.Enabled=0:FireI1.Enabled=0:FireR1.Enabled=0:FireE1.Enabled=0
      TWall.enabled = 0:TWall1.enabled = 0:TWall2.enabled = 1
      TriggerLF.enabled=0:TriggerRF.enabled=0
      CradleLeft.enabled=0:CradleRight.enabled=0
      TriggerAutoFlipPlunger.enabled = 0
      plunger.mechplunger = true
      isMultiball = 0
    If PostLightOnOff = 1 Then
      PostLight.Duration 2, 1200, 0
      playsound"servo_small"
    End If
      End If
  End If

End Sub

Sub CheckAI()
          If myPrefs_Player1Setup = 1 and BGL1up.state = 1 Then
    ToggleAI(1)
  ElseIf myPrefs_Player2Setup = 1 and BGL2up.state = 1 Then
      ToggleAI(1)
  ElseIf myPrefs_Player3Setup = 1 and BGL3up.state = 1 Then
    ToggleAI(1)
  ElseIf myPrefs_Player4Setup = 1 and BGL4up.state = 1 Then
    ToggleAI(1)
  Else
    ToggleAI(0)
  End If
End Sub

sub Trigger2_Hit
  stopsound "__SpaceWaltz(music)"
  stopsound "__My God, its full of stars(music)"
  stopsound "__2001(music)"
        stopsound "__RequiemIntro(music)"
  stopsound "__Ode To The Sun (Outro)"
  CheckAI
End Sub


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 4 ' total number of balls
ReDim rolling(tnob)
InitRolling

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
        StopSound("fx_ballrolling" & b)
    Next

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        '********************** Legendary *********************************
         if AIon and myPrefs_Difficulty = 4 Then
    'BouncepassOnOff = 1
          FlipperLeft.Scatter = 0
    FlipperRight.Scatter = 0
  else
    'BouncepassOnOff = 0
          FlipperLeft.Scatter = 5
    FlipperRight.Scatter = 5
  end if


  '********************** Test Flipper Magic *************************
  If AIon = 1  and myPrefs_Difficulty > 1 and isMultiball = 0 And FlipperMagic=1 Then

    if InRect(BOT(b).x,BOT(b).y,168,1455,200,1494,185,1544,128,1504) Then
      'debug.print "Left Flipper velx = " & BOT(b).velx & " vely = " & BOT(b).vely
      if BOT(b).velx < 3.2 and BOT(b).velx > 0 then
        LLane = 1
      Elseif BOT(b).velx < 7 and BOT(b).velx > 0 then
        LLane = 2
      Else
        LLane = 0
      end if
    Elseif InRect(BOT(b).x,BOT(b).y,726,1441,760,1495,702,1535,700,1481) Then
      'debug.print "Right Flipper velx = " & BOT(b).velx & " vely = " & BOT(b).vely
      if BOT(b).velx > -3.2  and BOT(b).velx < 0 then
        RLane = 1
      elseif BOT(b).velx > -7  and BOT(b).velx < 0 then
        RLane = 2
      Else
        RLane = 0
      end if
    End If

    If LLane = 1 and BOT(B).x > 200 and BOT(b).y > 1400 and BOT(b).y < 1650 Then
      LLane = 0
      vpmkeydown(LeftFlipperKey)':Msgbox "Try to grab the ball"
      CradleLeft.Enabled=1
      TriggerLF.enabled=0
      Playsound"_bleep-bloops"
    Elseif LLane = 2 and BOT(B).x > 200 and BOT(b).y > 1400 and BOT(b).y < 1650 Then
      TriggerLF.enabled=0
      If BOT(B).x > 353 Then
        LLane = 0
        vpmkeydown(LeftFlipperKey)
        TimerLF.enabled = 1
      End If
    End If


    If RLane = 1 and BOT(B).x < 690 and BOT(b).y > 1400 and BOT(b).y < 1650 Then
      RLane = 0
      vpmkeydown(RightFlipperKey)':Msgbox "Try to grab the ball"
      CradleRight.Enabled=1
      TriggerRF.enabled=0
      Playsound"_bleep-bloops"
    ElseIf RLane = 2 and BOT(B).x < 690 and BOT(b).y > 1400 and BOT(b).y < 1650 Then
      TriggerRF.enabled=0
      If BOT(B).x < 526 and AIOn Then
        RLane = 0
        vpmkeydown(RightFlipperKey)
        TimerRF.enabled = 1
      End If
    End If
  End If

  If AIon = 1 and myPrefs_Difficulty > 1 And FlipperMagic=1 Then
    if InRect(BOT(b).x,BOT(b).y,383,1604,433,1604,433,1654,383,1654) Then
      if BOT(b).vely > 0 and BOT(b).vely > 3* ABS(BOT(b).velx) and timerlf.enabled =0  Then
        vpmkeydown(LeftFlipperKey)
        TimerLF.interval = 100
        TimerLF.enabled = 1
        TriggerLF.enabled = 0
      End If
    Elseif InRect(BOT(b).x,BOT(b).y,383,1664,433,1664,433,1714,383,1714) Then
      if BOT(b).vely > 0 and timerlf.enabled = 0 Then
        vpmkeydown(LeftFlipperKey)
        TimerLF.interval = 100
        TimerLF.enabled = 1
        TriggerLF.enabled = 0
      End If
    Elseif InRect(BOT(b).x,BOT(b).y,437,1604,487,1604,487,1654,437,1654) Then
      if BOT(b).vely > 0 and BOT(b).vely > 3*ABS(BOT(b).velx) and timerrf.enabled =0 Then
        vpmkeydown(RightFlipperKey)
        TimerRF.interval = 100
        TimerRF.enabled = 1
        TriggerRF.enabled = 0
      End If
    Elseif InRect(BOT(b).x,BOT(b).y,437,1664,487,1664,487,1714,437,1714) Then
      if BOT(b).vely > 0 and timerrf.enabled =0 Then
        vpmkeydown(RightFlipperKey)
        TimerRF.interval = 100
        TimerRF.enabled = 1
        TriggerRF.enabled = 0
      End If
    End If
  End If

    Next
End Sub

Dim LLane, RLane
Dim FlipperMagic
FlipperMagic=1

'******************************************* AI Cradle Logic *************************************************
'********************************************* Flipper Left **************************************************

Sub  ZoneLeft_Hit
     if isMultiball = 0 Then
          if myPrefs_Difficulty>0 Then
         If Activeball.velX<4 and Activeball.velX>-4 Then
         If Activeball.velY<16.5 and Activeball.velY>-16.5 Then
       vpmkeydown(LeftFlipperKey)':Msgbox Activeball.velx':Msgbox "Try to grab the ball"
       CradleLeft.Enabled=1
       TriggerLF.enabled=0
       Playsound"_bleep-bloops"
                         Light10.Color=RGB(0,0,255)
                         Light10.ColorFull=RGB(0,0,255)
                         Light10.state=2
        End If
        End If
        End If
      End If
End Sub

Sub ZoneLeft_Unhit':Msgbox "Let go and reset"
  vpmkeyup(LeftFlipperKey)
  CradleLeft.Enabled=0:CradleLeftTimer.Enabled=0
  TriggerLF.enabled=1
  Stopsound"_bleep-bloops"
        Light10.Color=RGB(255,0,0)
        Light10.ColorFull=RGB(255,0,0)
        Light10.State=1
End Sub

Sub CradleLeft_Hit:CradleLeftTimer.Enabled=1:End Sub
Sub CradleLeft_UnHit:CradleLeftTimer.Enabled=0:End Sub

Sub CradleLeftTimer_Timer
        'ZoneLeft1.Enabled=0 'BouncePass
  CradleLeftTimer.Enabled=0
  vpmkeyup(LeftFlipperKey)
  TriggerLF.enabled=0
  CheckShotPriority 0
        Light10.Color=RGB(255,0,0)
        Light10.ColorFull=RGB(255,0,0)
        Light10.State=1
End Sub

Sub TimerLShot2_Timer '(Post Pass)
  TriggerRF.enabled=0
  vpmkeydown(LeftFlipperKey)
  CradleLogicLResetTimer.Enabled=1
  CradleRight.Enabled=1
  PostPassLDelayTimer.Enabled=1
  TimerLShot2.Enabled=0
End Sub

Sub TimerLShot3_Timer
  vpmkeydown(LeftFlipperKey)
  TimerLShot3.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
    Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot3.intensity=1500
  End If
End Sub

Sub TimerLShot4_Timer
  vpmkeydown(LeftFlipperKey)
  TimerLShot4.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot4.intensity=1500
  End If
End Sub

Sub TimerLShot5_Timer
  vpmkeydown(LeftFlipperKey)
  TimerLShot5.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot5.intensity=1500
  End If
End Sub

Sub TimerLShot6_Timer
  vpmkeydown(LeftFlipperKey)
  TimerLShot6.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot6.intensity=1500
  End If
End Sub

Sub TimerLShot7_Timer
  vpmkeydown(LeftFlipperKey)
  TimerLShot7.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot7.intensity=1500
  End If
End Sub

Sub TimerLShot8_Timer
  vpmkeydown(LeftFlipperKey)
  TimerLShot8.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot8.intensity=1500
  End If
End Sub

Sub TimerLShot9_Timer
  vpmkeydown(LeftFlipperKey)
  TimerLShot9.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot9.intensity=1500
  End If
End Sub

Sub TimerLShot10_Timer 'Top Power Target
  vpmkeydown(LeftFlipperKey)
  TimerLShot10.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot10.intensity=1500
  End If
End Sub

Sub TimerLShot11_Timer 'Middle Power Target/Extra Ball
  vpmkeydown(LeftFlipperKey)
  TimerLShot11.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot11.intensity=1500
  End If
End Sub

Sub TimerLShot12_Timer 'Bottom Power Target
  vpmkeydown(LeftFlipperKey)
  TimerLShot12.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot12.intensity=1500
  End If
End Sub

'Sub TimerLShot11_Timer 'Extra Ball
' vpmkeydown(LeftFlipperKey)
' TimerLShot11.Enabled=0:CradleLogicLResetTimer.Enabled=1
' Stopsound "_laserAiming"
' If myPrefs_SoundEffects Then
'   Playsound "_retro-shot-blaster":Playsound "_retro-shot-blaster":Playsound "Laserblast"
'   StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
'   Shot13.intensity=2000
' End If
'End Sub

'*********************** AI Combo Shots (Expert/Legendary) ************************

Sub TimerLShot14_Timer
  vpmkeydown(LeftFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerLShot14.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot14.intensity=1500
  End If
End Sub

Sub TimerLShot15_Timer
  vpmkeydown(LeftFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerLShot15.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot15.intensity=1500
  End If
End Sub

Sub TimerLShot16_Timer
  vpmkeydown(LeftFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerLShot16.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot16.intensity=1500
  End If
End Sub

Sub TimerLShot17_Timer
  vpmkeydown(LeftFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerLShot17.Enabled=0:CradleLogicLResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot17.intensity=1500
  End If
End Sub

'***********************************************************************

Sub CradleLogicLResetTimer_Timer
  ZoneLeft.Enabled=1:ZoneRight.Enabled=1
  vpmkeyup(LeftFlipperKey)
  CradleLogicLResetTimer.Enabled=0
  Shot1.intensity=0
  Shot2.intensity=0
  Shot3.intensity=0
  Shot4.intensity=0
  Shot5.intensity=0
  Shot6.intensity=0
  Shot7.intensity=0
  Shot8.intensity=0
  Shot9.intensity=0
  Shot10.intensity=0
  Shot11.intensity=0
  Shot12.intensity=0
  Shot13.intensity=0
  Shot14.intensity=0
  Shot15.intensity=0
  Shot16.intensity=0
  Shot17.intensity=0
End Sub

Sub PostPassLDelayTimer_Timer
  vpmkeydown(RightFlipperKey)
  PostPassLDelayTimer.Enabled=0
End Sub


'******************************************** AI Flipper Right ************************************************

Sub  ZoneRight_Hit
   if isMultiball = 0 Then
       if myPrefs_Difficulty>0 Then
      If Activeball.velX<3.5 and Activeball.velX>-4 Then
    If Activeball.velY<16.5 and Activeball.velY>-16.5 Then
      vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)':Msgbox "Try to grab the ball"
      CradleRight.Enabled=1
      TriggerRF.enabled=0
      Playsound"_bleep-bloops"
                         Light11.Color=RGB(0,0,255)
                         Light11.ColorFull=RGB(0,0,255)
                         Light11.State=2
    End If
      End If
         End If
      End If
End Sub

Sub ZoneRight_Unhit':Msgbox "Let go and reset"
  vpmkeyup(RightFlipperKey)
  CradleRight.Enabled=0:CradleRightTimer.Enabled=0
  TriggerRF.enabled=1
  Stopsound"_bleep-bloops"
        Light11.Color=RGB(255,0,0)
        Light11.ColorFull=RGB(255,0,0)
  Light11.State=1
End Sub

Sub CradleRight_Hit:CradleRightTimer.Enabled=1:End Sub
Sub CradleRight_UnHit:CradleRightTimer.Enabled=0:End Sub

Sub CradleRightTimer_Timer
        'ZoneRight1.Enabled=0 'BouncePass
  CradleRightTimer.Enabled=0
  vpmkeyup(RightFlipperKey)
  TriggerRF.enabled=0
  CheckShotPriority 1
        Light11.Color=RGB(255,0,0)
        Light11.ColorFull=RGB(255,0,0)
  Light11.State=1
End Sub

Sub TimerRShot1_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot1.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot1.intensity=1500
  End If
End Sub

Sub TimerRShot2_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot2.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot2.intensity=1500
  End If
End Sub

Sub TimerRShot3_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot3.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot3.intensity=1500
  End If
End Sub

Sub TimerRShot4_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot4.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot4.intensity=1500
  End If
End Sub

Sub TimerRShot5_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot5.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot5.intensity=1500
  End If
End Sub

Sub TimerRShot6_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot6.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot6.intensity=1500
  End If
End Sub

Sub TimerRShot7_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot7.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot7.intensity=1500
  End If
End Sub

Sub TimerRShot8_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot8.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot8.intensity=1500
  End If
End Sub

Sub TimerRShot9_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot9.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot9.intensity=1500
  End If
End Sub

Sub TimerRShot10_Timer '(Post Pass)
  TriggerLF.enabled=0
  vpmkeydown(RightFlipperKey)
  CradleLogicRResetTimer.Enabled=1
  CradleLeft.Enabled=1
  PostPassRDelayTimer.Enabled=1
  TimerRShot10.Enabled=0
End Sub

'***************** AI Bounce-Pass Logic ****************** FIX

Sub ZoneLeft1_Hit
        If AIOn and Activeball.velY>16.7 and Activeball.velY<18  Then
           If myPrefs_Difficulty = 4 and isMultiball = 0 Then
        If Activeball.velX>2 Then
                   FlipperMagic=0
                   'Msgbox "Bounce Pass"
      Playsound "Laugh"
      TriggerLF.Enabled=0
      TriggerRF.Enabled=1 'NEW
      ZoneLeft.Enabled=0
      ZoneRight.Enabled=1 'NEW
      'BouncePassLTimer.Enabled=1
      BouncePassRTimer.Enabled=1
          End If
       End If
          End If
End Sub
'
Sub ZoneRight1_Hit
  If Activeball.velY>16.7 and Activeball.velY<17.5 Then
     If myPrefs_Difficulty>3 and isMultiball = 0 Then
        If Activeball.velX<-2 Then
                   FlipperMagic=0
                   'Msgbox "Bounce Pass"
      Playsound "Laugh2"
        TriggerLF.Enabled=1 'NEW
        TriggerRF.Enabled=0
        ZoneLeft.Enabled=1 'NEW
        ZoneRight.Enabled=0
        'BouncePassRTimer.Enabled=1
        BouncePassLTimer.Enabled=1
          End If
       End If
          End If
End Sub

Sub BouncePassLTimer_Timer
        TriggerLF.enabled=1 '(Gets Enabled above in ZoneLeft_Unhit)
        TriggerRF.Enabled=1
        ZoneRight.Enabled=1
        ZoneLeft.Enabled=1
        FlipperMagic=1
  BouncePassLTimer.Enabled=0
End Sub

Sub BouncePassRTimer_Timer
    TriggerLF.Enabled=1
    TriggerRF.enabled=1 '(Gets Enabled above in ZoneRight_Unhit)
    ZoneRight.Enabled=1
    ZoneLeft.Enabled=1
    FlipperMagic=1
    BouncePassRTimer.Enabled=0
End Sub

'*********************** AI Combo Shots (Expert) ************************

Sub TimerRShot14_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot14.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot14.intensity=1500
  End If
End Sub

Sub TimerRShot15_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot15.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
    Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot15.intensity=1500
  End If
End Sub

Sub TimerRShot16_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot16.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
    Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot16.intensity=1500
  End If
End Sub

Sub TimerRShot17_Timer
  vpmkeydown(RightFlipperKey):'vpmtimer.pulsesw(cFlipperRightSW)
  TimerRShot17.Enabled=0:CradleLogicRResetTimer.Enabled=1
  Stopsound "_laserAiming"
  If myPrefs_SoundEffects Then
    Playsound "_retro-shot-blaster"
    Playsound "Laserblast"
    StrobeLight6.State=1:StrobeLight5.State=0:StrobeLightTimer.Enabled=1
    Shot17.intensity=1500
  End If
End Sub

'***********************************************************************

Sub PostPassRDelayTimer_Timer
  vpmkeydown(LeftFlipperKey)
  PostPassRDelayTimer.Enabled=0
        TriggerRF.enabled=1
End Sub

Sub CradleLogicRResetTimer_Timer
  ZoneLeft.Enabled=1:ZoneRight.Enabled=1
  vpmkeyup(RightFlipperKey)
  CradleLogicRResetTimer.Enabled=0
  Shot1.intensity=0
  Shot2.intensity=0
  Shot3.intensity=0
  Shot4.intensity=0
  Shot5.intensity=0
  Shot6.intensity=0
  Shot7.intensity=0
  Shot8.intensity=0
  Shot9.intensity=0
  Shot10.intensity=0
  Shot11.intensity=0
  Shot12.intensity=0
  Shot13.intensity=0
  Shot14.intensity=0
  Shot15.intensity=0
  Shot16.intensity=0
  Shot17.intensity=0
End Sub

'********************************** AI Shot Priority ****************************************

Sub CheckShotPriority(flipper) 'flipper: 0 = left flipper; 1 = right flipper
  Dim PowerCount, LockCount

  LockCount = 0

  PowerCount = 0
  If shot10.state = 2 Then: PowerCount = PowerCount + 1: End If
  If shot11.state = 2 Then: PowerCount = PowerCount + 1: End If
  If shot12.state = 2 Then: PowerCount = PowerCount + 1: End If


  If  KRightEjectHole.ballcntover>0 and shotonLight9.state = 1 Then
    LockCount = LockCount + 1
  End If

  If  KUpperEjectHole.ballcntover>0 and shotonLight2.state = 1 Then
    LockCount = LockCount + 1
  End If

  If  KLeftEjectHole.ballcntover>0 and shotonLight1.state = 1 Then
    LockCount = LockCount + 1
  End If

  If flipper = 0 Then
    ZoneLeft.Enabled=0
  Else
    ZoneRight.Enabled = 0
  End If

  If ismultiball = 0 Then

    '***************************** Shot Priority when without multiball


    '**********  Extra Ball
    If Shot2.state = 2 and  L5x.state = 1 and LExtraBall.state =0 and LShootAgain.state = 0 Then

      If flipper = 0 Then
        LeftPostPass
      Else
        TimerRShot2.Enabled=1
        SoundStrobe "Shot2"
      End If

    ElseIf Shot13.state = 2 Then 'Extra ball

      If flipper = 0 Then
        TimerLShot11.Enabled=1
        SoundStrobe "Shot13"
      Else  '(Post Pass)
        RightPostPass
      End If

    '**********  Lock
    Elseif Shot1.state = 2 or Shot2.state=2 or Shot9.state = 2 Then 'Lock

      If flipper = 0 and Shot9.state =2 Then
        TimerLShot9.Enabled=1
        SoundStrobe "Shot9"
      Elseif flipper = 0 then  '(Post Pass)
        LeftPostPass
      Else
        If Shot1.state=2 Then
          TimerRShot1.Enabled=1
          SoundStrobe "Shot1"
        Elseif Shot2.state = 2 Then
          TimerRShot2.Enabled=1
          SoundStrobe "Shot2"
        Else
                                    Select Case Int(Rnd()*2)
                                        Case 0
             TimerRShot9.Enabled=1
                                           SoundStrobe "Shot9"
                                        Case 1
                                           RightPostPass
                                     End Select
        End If
      End If

    '**********  Power
              'Elseif PowerCount > 0 and PowerCount < 3  and LockCount < 2 and LShieldOn.state = 1 and myPrefs_Difficulty < 3 Then '(Disable Shot Priority in Expert)
    Elseif PowerCount > 0 and PowerCount < 3  and LockCount < 2 and LShieldOn.state = 1 Then  '(Enabled Power Shot Priority in Expert)
      If flipper = 0 Then
        If Shot10.State=2 Then
          TimerLShot10.Enabled=1
          SoundStrobe "Shot10"
        ElseIf Shot11.State=2 Then
          TimerLShot11.Enabled=1
          SoundStrobe "Shot11"
        Else
          TimerLShot12.Enabled=1
          SoundStrobe "Shot12"
        End If
      Else
        RightPostPass
      End If

    '**********  Targets
    Else
      TargetPriority flipper
    End if

  Else

    '***************************** Shot Priority with multiball - No Post Passes or maybe just shoot left loop or right loop?


    '**********  Extra Ball
    If Shot13.state = 2 and flipper = 0 Then 'Extra ball
      TimerLShot11.Enabled=1
      SoundStrobe "Shot13"

    '**********  Lock
    Elseif flipper=0 and Shot9.state = 2 Then
      TimerLShot9.Enabled=1
      SoundStrobe "Shot9"
    Elseif flipper = 1 and (Shot1.state = 2 or Shot2.state = 2) Then
      If Shot1.state=2 Then
        TimerRShot1.Enabled=1
        SoundStrobe "Shot1"
      Else
        TimerRShot2.Enabled=1
        SoundStrobe "Shot2"
      End If

    '**********  Power
    Elseif PowerCount > 0 and flipper = 0 Then
      If Shot10.State=2 Then
        TimerLShot10.Enabled=1
        SoundStrobe "Shot10"
      ElseIf Shot11.State=2 Then
        TimerLShot11.Enabled=1
        SoundStrobe "Shot11"
      Else
        TimerLShot12.Enabled=1
        SoundStrobe "Shot12"
      End If

    '**********  Targets
    Else
      TargetPriority flipper
    End If
  End If

End Sub

Sub LeftPostPass()
  TimerLShot2.Enabled=1
  TriggerRF.enabled=0
  TriggerLF.enabled=0
End Sub

Sub RightPostPass()
  TimerRShot10.Enabled=1
  TriggerRF.enabled=0
  TriggerLF.enabled=0
End Sub

Sub SoundStrobe(Shot)   'Learn This AXS
  If myPrefs_SoundEffects Then
    StrobeLight5.State=2
    Playsound "_laserAiming"
                Playsound "_laserAiming"
                Playsound "_laserAiming"
                Playsound "_ringing-laser-low2"
    Eval(Shot).intensity=1501
  End If
End Sub

Sub TargetPriority(flipper)

  If flipper = 0  Then
    If Shot16.state = 2 and  myPrefs_Difficulty > 2 Then
      TimerLShot16.Enabled = 1
      SoundStrobe "Shot16"
    Elseif Shot17.state=2  and  myPrefs_Difficulty > 2 Then
      TimerLShot17.Enabled = 1
      SoundStrobe "Shot17"
    ElseIf Shot15.state = 2  and  myPrefs_Difficulty > 2 Then
      TimerLShot15.Enabled = 1
      SoundStrobe "Shot15"
    Elseif Shot14.state = 2  and  myPrefs_Difficulty > 2 Then
      TimerLShot14.Enabled = 1
      SoundStrobe "Shot14"
    ElseIf Shot6.State=2  and  myPrefs_Difficulty > 2 Then '  and  myPrefs_Difficulty<>3
      TimerLShot6.Enabled=1
      SoundStrobe "Shot6"
    Elseif Shot7.State=2 Then
      TimerLShot7.Enabled=1
      SoundStrobe "Shot7"
    Elseif Shot8.State=2 and  myPrefs_Difficulty<>3 Then '  and  myPrefs_Difficulty<>3
      TimerLShot8.Enabled=1
      SoundStrobe "Shot8"
    ElseIf Shot3.State=2 Then
      TimerLShot3.Enabled=1
      SoundStrobe "Shot3"
    Elseif Shot4.State=2 Then
      TimerLShot4.Enabled=1
      SoundStrobe "Shot4"
    Elseif Shot5.State = 2 and  myPrefs_Difficulty<>3 Then
      TimerLShot5.Enabled=1
      SoundStrobe "Shot5"
    Else
      LeftPostPass
    End If
  Else
    If Shot15.state = 2 and  myPrefs_Difficulty > 2 Then
      TimerRShot15.Enabled = 1
      SoundStrobe "Shot15"
    Elseif Shot14.state=2 and  myPrefs_Difficulty > 2 Then
      TimerRShot14.Enabled = 1
      SoundStrobe "Shot14"
    ElseIf Shot16.state = 2 and  myPrefs_Difficulty > 2 Then
      TimerRShot16.Enabled = 1
      SoundStrobe "Shot16"
    Elseif Shot17.state = 2 and  myPrefs_Difficulty > 2 Then
      TimerRShot17.Enabled = 1
      SoundStrobe "Shot17"
    ElseIf Shot3.State=2 Then
      TimerRShot3.Enabled=1
      SoundStrobe "Shot3"
    Elseif Shot4.State=2 and  myPrefs_Difficulty<>3 Then '  and  myPrefs_Difficulty<>3
      TimerRShot4.Enabled=1
      SoundStrobe "Shot4"
    Elseif Shot5.State = 2 Then
      TimerRShot5.Enabled=1
      SoundStrobe "Shot5"
    Elseif Shot6.State=2 Then
      'RightPostPass 'Check this
      TimerRShot6.Enabled=1
      SoundStrobe "Shot6"
    Elseif Shot7.State=2 Then
      TimerRShot7.Enabled=1
      SoundStrobe "Shot7"
    Elseif Shot8.state=2 Then
      TimerRShot8.Enabled=1
      SoundStrobe "Shot8"
    Else
      RightPostPass
    End If
  End If

End Sub


Set Lights(33) = ShotonLight1
Set Lights(34) = ShotonLight9
Set Lights(35) = ShotonLight2
Set Lights(26) = ShotonLight3
Set Lights(27) = ShotonLight4
Set Lights(28) = ShotonLight5
Set Lights(29) = ShotonLight6
Set Lights(30) = ShotonLight7
Set Lights(31) = ShotonLight8
Set Lights(9)   = ShotonLight10
Set Lights(10) = ShotonLight11
Set Lights(11) = ShotonLight12
Set Lights(40) = ShotonLight13

dim onStart1, onInterval1: onStart1 = 0
dim offStart1, offInterval1: offStart1 = 0
dim onStart2, onInterval2: onStart2 = 0
dim offStart2, offInterval2: offStart2 = 0
dim onStart3, onInterval3: onStart3 = 0
dim offStart3, offInterval3: offStart3 = 0
dim onStart4, onInterval4: onStart4 = 0
dim offStart4, offInterval4: offStart4 = 0
dim onStart5, onInterval5: onStart5 = 0
dim offStart5, offInterval5: offStart5 = 0
dim onStart6, onInterval6: onStart6 = 0
dim offStart6, offInterval6: offStart6 = 0
dim onStart7, onInterval7: onStart7 = 0
dim offStart7, offInterval7: offStart7 = 0
dim onStart8, onInterval8: onStart8 = 0
dim offStart8, offInterval8: offStart8 = 0
dim onStart9, onInterval9: onStart9 = 0
dim offStart9, offInterval9: offStart9 = 0

sub highTimer_Timer

  '************************************* Shot1 (Ball Lock)

  if ShotonLight1.state=1 then
    if onStart1 = 0 Then
      'debug.print "Off: " & offInterval1
      offStart1 = 0
      onStart1 = 1
      OnInterval1 = 0
    Else
      onInterval1 = onInterval1 + me.interval
    End If
  Elseif ShotonLight1.state=0 then
    if offStart1 = 0 Then
      'debug.print "On: " & onInterval1
      onStart1 = 0
      offStart1 = 1
      offInterval1 = 0
        Else
      offInterval1 = offInterval1 + me.interval
    End If
  End If

  If onInterval1 < 200 and OffInterval1 < 200 and KLeftEjectHole.BallCntOver < 1 Then
    Shot1.State = 2
  Else
    Shot1.State = 0
  End If

  '**************************************** Shot2 (Ball Lock)

  if ShotonLight2.state=1 then
    if onStart2 = 0 Then
      'debug.print "Off: " & offInterval2
      offStart2 = 0
      onStart2 = 1
      OnInterval2 = 0
    Else
      onInterval2 = onInterval2 + me.interval
    End If
  Elseif ShotonLight2.state=0 then
    if offStart2 = 0 Then
      'debug.print "On: " & onInterval2
      onStart2 = 0
      offStart2 = 1
      offInterval2 = 0
    Else
      offInterval2 = offInterval2 + me.interval
    End If
  End If

  If (onInterval2 < 200 and OffInterval2 < 200 and  KUpperEjectHole.BallCntOver < 1) or (L5X.state = 1 and LExtraBall.state=0 and LShootAgain.state = 0) and (ShotonLight1.state=0 and ShotonLight2.state=0 and ShotonLight9.state=0)Then
    Shot2.State = 2
  Else
    Shot2.State = 0
  End If


  '**************************************** Shot3 (Target 1)

  if ShotonLight3.state=1 then
    if onStart3 = 0 Then
      'debug.print "Off: " & offInterval3
      offStart3 = 0
      onStart3 = 1
      OnInterval3 = 0
    Else
      onInterval3 = onInterval3 + me.interval
    End If
  Elseif ShotonLight3.state=0 then
    if offStart3 = 0 Then
      'debug.print "On: " & onInterval3
      onStart3 = 0
      offStart3 = 1
      offInterval3 = 0
    Else
      offInterval3 = offInterval3 + me.interval
    End If
  End If

  If onInterval3 < 200 and OffInterval3 < 200 Then
    Shot3.State = 2
  Else
    Shot3.State = 0
  End If

  '**************************************** Shot4 (Target 2)

  if ShotonLight4.state=1 then
    if onStart4 = 0 Then
      'debug.print "Off: " & offInterval4
      offStart4 = 0
      onStart4 = 1
      OnInterval4 = 0
    Else
      onInterval4 = onInterval4 + me.interval
    End If
  Elseif ShotonLight4.state=0 then
    if offStart4 = 0 Then
      'debug.print "On: " & onInterval4
      onStart4 = 0
      offStart4 = 1
      offInterval4 = 0
    Else
      offInterval4 = offInterval4 + me.interval
    End If
  End If

  If onInterval4 < 200 and OffInterval4 < 200 Then
    Shot4.State = 2
  Else
    Shot4.State = 0
  End If

  '**************************************** Shot5 (Target 3)

  if ShotonLight5.state=1 then
    if onStart5 = 0 Then
      'debug.print "Off: " & offInterval5
      offStart5 = 0
      onStart5 = 1
      OnInterval5 = 0
    Else
      onInterval5 = onInterval5 + me.interval
    End If
  Elseif ShotonLight5.state=0 then
    if offStart5 = 0 Then
      'debug.print "On: " & onInterval5
      onStart5 = 0
      offStart5 = 1
      offInterval5 = 0
    Else
      offInterval5 = offInterval5 + me.interval
    End If
  End If

  If onInterval5 < 200 and OffInterval5 < 200 Then
    Shot5.State = 2
  Else
    Shot5.State = 0
  End If

  '**************************************** Shot6 (Target 4)

  if ShotonLight6.state=1 then
    if onStart6 = 0 Then
      'debug.print "Off: " & offInterval6
      offStart6 = 0
      onStart6 = 1
      OnInterval6 = 0
    Else
      onInterval6 = onInterval6 + me.interval
    End If
  Elseif ShotonLight6.state=0 then
    if offStart6 = 0 Then
      'debug.print "On: " & onInterval6
      onStart6 = 0
      offStart6 = 1
      offInterval6 = 0
    Else
      offInterval6 = offInterval6 + me.interval
    End If
  End If

  If onInterval6 < 200 and OffInterval6 < 200 Then
    Shot6.State = 2
  Else
    Shot6.State = 0
  End If

  '**************************************** Shot7 (Target 5)

  if ShotonLight7.state=1 then
    if onStart7 = 0 Then
      'debug.print "Off: " & offInterval7
      offStart7 = 0
      onStart7 = 1
      OnInterval7 = 0
    Else
      onInterval7 = onInterval7 + me.interval
    End If
  Elseif ShotonLight7.state=0 then
    if offStart7 = 0 Then
      'debug.print "On: " & onInterval7
      onStart7 = 0
      offStart7 = 1
      offInterval7 = 0
    Else
      offInterval7 = offInterval7 + me.interval
    End If
  End If

  If onInterval7 < 200 and OffInterval7 < 200 Then
    Shot7.State = 2
  Else
    Shot7.State = 0
  End If

  '**************************************** Shot8 (Target 6)

  if ShotonLight8.state=1 then
    if onStart8 = 0 Then
      'debug.print "Off: " & offInterval8
      offStart8 = 0
      onStart8 = 1
      OnInterval8 = 0
    Else
      onInterval8 = onInterval8 + me.interval
    End If
  Elseif ShotonLight8.state=0 then
    if offStart8 = 0 Then
      'debug.print "On: " & onInterval8
      onStart8 = 0
      offStart8 = 1
      offInterval8 = 0
    Else
      offInterval8 = offInterval8 + me.interval
    End If
  End If

  If onInterval8 < 200 and OffInterval8 < 200 Then
    Shot8.State = 2
  Else
    Shot8.State = 0
  End If

  '**************************************** Shot9 (Ball Lock)

  if ShotonLight9.state=1 then
    if onStart9 = 0 Then
      'debug.print "Off: " & offInterval9
      offStart9 = 0
      onStart9 = 1
      OnInterval9 = 0
    Else
      onInterval9 = onInterval9 + me.interval
    End If
  Elseif ShotonLight9.state=0 then
    if offStart9 = 0 Then
      'debug.print "On: " & onInterval9
      onStart9 = 0
      offStart9 = 1
      offInterval9 = 0
    Else
      offInterval9 = offInterval9 + me.interval
    End If
  End If

  If onInterval9 < 200 and OffInterval9 < 200 and KRightEjectHole.BallCntOver < 1 Then
    Shot9.State = 2
  Else
    Shot9.State = 0
  End If

  '**************************************** Shot10 (Top Power Target)

  if LFire1.state=1 and ShotonLight10.state = 0 Then
          Shot10.State = 2
  Else
    Shot10.State = 0
  End If

  '**************************************** Shot11 (Middle Power Target)

  if LFire1.state=1 and ShotonLight11.state = 0  Then
    Shot11.State = 2
  Else
    Shot11.State = 0
  End If


  '**************************************** Shot12 (Bottom Power Target)

  if LFire1.state=1 and ShotonLight12.state = 0  Then
    Shot12.State = 2
  Else
    Shot12.State = 0
  End If

  '**************************************** Shot13 (Extra Ball) Priority 1

  if ShotonLight13.state=1 then
    Shot13.State = 2
  Else
    Shot13.State = 0
  End If

  '**************************************** Combo Shots

  If Shot3.state = 2 and Shot4.state =2 Then
    Shot14.state = 2
  Else
    Shot14.state = 0
  End If

  If Shot4.state = 2 and Shot5.state =2 Then
    Shot15.state = 2
  Else
    Shot15.state = 0
  End If

  If Shot6.state = 2 and Shot7.state =2 Then
    Shot16.state = 2
  Else
    Shot16.state = 0
  End If

  If Shot7.state = 2 and Shot8.state =2 Then
    Shot17.state = 2
  Else
    Shot17.state = 0
  End If



  '**************************************** Check Multiball

  Dim BallsInKickers

  BallsInKickers = KUpperEjectHole.ballcntover +  KLeftEjectHole.ballcntover + KRightEjectHole.ballcntover

  if AIon  and ismultiball = 0 and ((BallsinKickers = 3 and shotonlight9.state = 1 and shotonlight1.state = 1 and shotonlight2.state = 1 and BGLtilt.state =1 and BGLmatch.state = 1) or (trTrough.balls = 0 and BallsinKickers < 2))then
    ismultiball = 1
    if (trTrough.balls = 0 and BallsinKickers < 2) Then
      isMultiballStarted = 1
    Else
      isMultiballStarted = 0
    End If
    'debug.print "isMultiball = 1"
    FireF.Enabled=0
    FireI.Enabled=0
    FireR.Enabled=0
    FireE.Enabled=0

    FireF1.Enabled=0
    FireI1.Enabled=0
    FireR1.Enabled=0
    FireE1.Enabled=0
  end if

  if isMultiball = 1 Then
    TestMultiball
  end if

  If AIon  Then
    If Controller.Switch(84) = true Then
      LeftFlipCount = LeftFlipCount + me.interval
      If LeftFlipCount > 3000 Then
        TimerLF.enabled = 1
        LeftFlipCount = 0
      End If
    Else
      LeftFlipCount = 0
    End If


    If Controller.Switch(82) = true Then
      RightFlipCount = RightFlipCount + me.interval
      If RightFlipCount > 3000 Then
        TimerRF.enabled = 1
        RightFlipCount = 0
      End If
    Else
      RightFlipCount = 0
    End If
  End If

   if Shot13.state = 2 and myPrefs_Difficulty >2 Then
        isMultiball = 0
    end if

 End sub

Dim LeftFlipCount, RightFlipCount
LeftFlipCount = 0
RightFlipCount = 0

'**************************************** Flip Pass
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

Dim FPLCount, FPRCount
FPLCount = 0: FPRCount = 0

Sub FlipPassL_Timer()
  If FPLCount = 0 and controller.Switch(84) = true Then
    FPLCount = 1
    vpmkeyup(LeftFlipperKey)
  Else
    FPLCount = 1
    vpmkeydown(LeftFlipperKey)
    me.enabled = false
  End If
End Sub

Sub FlipPassR_Timer()
  If FPRCount = 0 and controller.Switch(82) = true Then
    FPRCount = 1
    vpmkeyup(RightFlipperKey)
  Else
    FPRCount = 1
    vpmkeydown(RightFlipperKey)
    me.enabled = false
  End If
End Sub


''************************************* AI Upper Lane Change ******************************************
Dim isMultiball: isMultiball = 0
Dim isMultiballStarted: isMultiballStarted = 0
Dim LowerZoneState: LowerZoneState = 0

   Sub  LowerZoneTrigger_Hit
       If myPrefs_Difficulty > 2 then
           LowerZoneState = LowerZoneState + 1
    If LowerZoneState  > 0 Then
    FireF.Enabled=False
    FireI.Enabled=False
    FireR.Enabled=False
    FireE.Enabled=False

    FireF1.Enabled=False
    FireI1.Enabled=False
    FireR1.Enabled=False
    FireE1.Enabled=False

                Light7.Color=RGB(255,0,0)
                Light7.Colorfull=RGB(255,0,0)
  end If
     end If
   End Sub

   Sub LowerZoneTrigger_UnHit
      If myPrefs_Difficulty > 2 then
           LowerZoneState = LowerZoneState - 1
  If LowerZoneState  <  1 Then
    FireF.Enabled=True
    FireI.Enabled=True
    FireR.Enabled=True
    FireE.Enabled=True

    FireF1.Enabled=True
    FireI1.Enabled=True
    FireR1.Enabled=True
    FireE1.Enabled=True

    Light7.Color=RGB(0,225,0)
                Light7.Colorfull=RGB(0,225,0)
  end If
     end If
   End Sub

Sub TestMultiball()
  Dim mBOT

  if isMultiballStarted = 0 and KLeftEjectHole.ballcntover =0 then isMultiballStarted = 1

  mBOT = 3 - trTrough.Balls

  If  KRightEjectHole.ballcntover>0 and shotonLight9.state = 1 Then
    mBOT = mBOT - 1
  End If

  If  KUpperEjectHole.ballcntover>0 and shotonLight2.state = 1 Then
    mBOT = mBOT - 1
  End If

  If  KLeftEjectHole.ballcntover>0 and shotonLight1.state = 1 Then
    mBOT = mBOT - 1
  End If

  If mBOT < 2 and isMultiballStarted  then
    ismultiball = 0
'   debug.print "isMultiball = 0"
'   debug.print "trtroughballs = " & trtrough.balls
'   debug.print "shotonLight1 = " & shotonlight1.state
'   debug.print "shotonLight2 = " & shotonlight2.state
'   debug.print "shotonLight9 = " & shotonlight9.state
'   debug.print "Right Kicker = " & KRightEjectHole.ballcntover
'   debug.print "Upper Kicker = " & KUpperEjectHole.ballcntover
'   debug.print "Left Kicker = " & KLeftEjectHole.ballcntover

    FireF.Enabled=True
    FireI.Enabled=True
    FireR.Enabled=True
    FireE.Enabled=True

    FireF1.Enabled=True
    FireI1.Enabled=True
    FireR1.Enabled=True
    FireE1.Enabled=True
           end if
      'end if
End Sub


  Set Lights(5) = FireFL
  Set Lights(6) = FireIL
  Set Lights(7) = FireRL
  Set Lights(8) = FireEL

Sub FireF_Hit
  If myPrefs_Difficulty=2 Then
    If FireFL.State=1 Then
      vpmkeydown(RightFlipperKey)
      'vpmtimer.pulsesw(cFlipperRightSW)
      TimerRF.Enabled=1
    End If
  ElseIf myPrefs_Difficulty > 2 Then
    If FireFL.State=1 Then
      vpmkeydown(RightFlipperKey)
      TimerRF.interval = 75
      TimerRF.Enabled=1
      FireF.TimerInterval = 150
      FireF.timerenabled = 1
      FireI.timerenabled=0
      FireR.timerenabled=0
      FireE.timerenabled=0
    Else
      FireF.timerenabled = 0
    End If
  End If
End Sub

Sub FireF1_Hit
  If myPrefs_Difficulty > 2 Then FireF_Hit
End Sub


Sub FireF_UnHit
  FireF_Hit
End Sub

 Sub FireF_Timer()
  FireF_Hit
 End Sub

 Sub FireI_Hit
  If myPrefs_Difficulty=2 Then
    If FireIL.State=1 Then
      vpmkeydown(RightFlipperKey)
      'vpmtimer.pulsesw(cFlipperRightSW)
      TimerRF.Enabled=1
    End If
  ElseIf myPrefs_Difficulty> 2 Then
    If FireIL.State=1 Then
      vpmkeydown(RightFlipperKey)
      TimerRF.interval = 75
      TimerRF.Enabled=1
      FireI.timerinterval = 150
      FireI.timerenabled = 1
      FireF.timerenabled=0
      FireR.timerenabled=0
      FireE.timerenabled=0
    Else
      FireI.timerenabled = 0
    End If
  End If
   End Sub

Sub FireI1_Hit
  If myPrefs_Difficulty> 2 Then FireI_Hit
End Sub

  Sub FireI_Unhit()
  FireI_Hit
 End Sub

 Sub FireI_Timer()
  FireI_Hit
 End Sub


Sub FireR_Hit
  If myPrefs_Difficulty=2 Then
    If FireRL.State=1 Then
      vpmkeydown(RightFlipperKey)
      'vpmtimer.pulsesw(cFlipperRightSW)
      TimerRF.Enabled=1
    End If
   ElseIf myPrefs_Difficulty> 2 Then
    If FireRL.State=1 Then
      vpmkeydown(RightFlipperKey)
      TimerRF.interval = 75
      TimerRF.Enabled=1
      FireR.timerinterval = 150
      FireR.timerenabled = 1
      FireF.timerenabled=0
      FireI.timerenabled=0
      FireE.timerenabled=0
    Else
      FireR.timerenabled = 0
    End If
  End If
End Sub

Sub FireR1_Hit
  If myPrefs_Difficulty> 2 Then FireR_Hit
End Sub

 Sub FireR_Unhit()
  FireR_Hit
 End Sub

 Sub FireR_Timer()
  FireR_Hit
 End Sub


Sub FireE_Hit
  If myPrefs_Difficulty=2 Then
    If FireEL.State=1 Then
      vpmkeydown(RightFlipperKey)
      'vpmtimer.pulsesw(cFlipperRightSW)
      TimerRF.Enabled=1
    End If
    ElseIf myPrefs_Difficulty> 2 Then
    If FireEL.State=1 Then
      vpmkeydown(RightFlipperKey)
      TimerRF.interval = 75
      TimerRF.Enabled=1
      FireE.timerinterval = 150
      FireE.timerenabled = 1
      FireF.timerenabled=0
      FireI.timerenabled=0
      FireR.timerenabled=0
    Else
      FireE.timerenabled = 0
    End If
  End If
End Sub

Sub FireE1_Hit
  If myPrefs_Difficulty> 2 Then FireE_Hit
End Sub

 Sub FireE_Unhit()
  FireE_Hit
 End Sub

 Sub FireE_Timer()
  FireE_Hit
 End Sub

'********************************* AI Nudging ****************************************

Dim doleftnudge, dorightnudge, docenternudge, nudgecount
doleftnudge = 0
dorightnudge = 0
docenternudge = 0

NudgeTimer.interval = 15
NudgeTimer.enabled = True

Sub NudgeTimer_Timer()
  If myPrefs_Difficulty=2 or myPrefs_Difficulty=3 or myPrefs_Difficulty=4 Then

  nudgecount = nudgecount + me.interval
  if nudgecount > 1300 then nudgecount = 1300

  If doleftnudge = 1 then
    doleftnudge = 0
    vpmkeydown(LeftTiltKey)
    nudgecount = 0
  elseif dorightnudge = 1 then
    dorightnudge = 0
    vpmkeydown(RightTiltKey)
    nudgecount = 0
  elseif docenternudge = 1 then
    docenternudge = 0
    vpmkeydown(CenterTiltKey)
    nudgecount = 0
  end if
    end if
End Sub

Sub NudgeSlingRight_Hit
  If myPrefs_Difficulty=2 or myPrefs_Difficulty=3 or myPrefs_Difficulty=4 Then
      If myPrefs_AInudge > 0 Then
    If AIOn and nudgecount > 1290 Then
      Playsound "_Nudge":Playsound "_Nudge":Playsound "_Nudge":Playsound "_Nudge"
      docenternudge = 1
      nudgecount = 0
'msgbox "on"
                 If myPrefs_AInudge = 1 then
      Rubber0041.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 2 then
      Rubber0042.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 3 then
      Rubber0043.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 4 then
      Rubber0044.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 5then
      Rubber0045.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
  If myPrefs_Difficulty = 4 Then
    myPrefs_AInudge= 5
    End If
  End If
end if
end if
End Sub

Sub NudgeUpperLock_Hit
  If myPrefs_Difficulty=4 and  shotonLight2.state = 1 Then
           If KUpperEjectHole.ballcntover<1 Then
       If myPrefs_AInudge > 0 Then
    If AIOn and nudgecount > 1290 Then
      Playsound "_Nudge":Playsound "_Nudge":Playsound "_Nudge":Playsound "_Nudge"
      docenternudge = 1
      nudgecount = 0
                 If myPrefs_AInudge = 1 then
      Rubber0031.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 2 then
      Rubber0032.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 3 then
      Rubber0033.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 4 then
      Rubber0034.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 5then
      Rubber0035.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
  If myPrefs_Difficulty = 4 Then
    myPrefs_AInudge= 5
    End If
             End If
    End If
      End If
end if
End Sub


Sub NudgeOutlaneRight_Hit
  If myPrefs_Difficulty=2 or myPrefs_Difficulty=3 or myPrefs_Difficulty=4 Then
      If myPrefs_AInudge > 0 Then
    If AIOn and nudgecount > 1290 Then
      Playsound "_Nudge":Playsound "_Nudge":Playsound "_Nudge":Playsound "_Nudge"
                 Select Case Int(Rnd()*2)
                   Case 0
      docenternudge = 1
      nudgecount = 0
                   Case 1
      dorightnudge = 1
      nudgecount = 0
                   End Select
'msgbox "on"
                 If myPrefs_AInudge = 1 then
      Rubber0011.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 2 then
      Rubber0012.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 3 then
      Rubber0013.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 4 then
      Rubber0014.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 5then
      Rubber0015.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
  If myPrefs_Difficulty = 4 Then
    myPrefs_AInudge= 5
            End If
        End If
    End If
end if
End Sub

Sub NudgeOutlaneLeft_Hit
  If myPrefs_Difficulty=2 or myPrefs_Difficulty=3 or myPrefs_Difficulty=4 Then
      If myPrefs_AInudge >0 Then
    If AIOn and nudgecount > 1290 Then
      Playsound "_Nudge":Playsound "_Nudge":Playsound "_Nudge":Playsound "_Nudge"
      docenternudge = 1
      nudgecount = 0
'msgbox "on"
                 If myPrefs_AInudge = 1 then
      Rubber0021.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 2 then
      Rubber0022.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 3 then
      Rubber0023.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 4 then
      Rubber0024.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
                 If myPrefs_AInudge = 5then
      Rubber0025.Collidable = 1
      OultLaneTimer.Enabled = 1
                End If
  If myPrefs_Difficulty = 4 Then
    myPrefs_AInudge= 5
    End If
            End If
  End If
end if
End Sub

Sub OultLaneTimer_Timer
  Rubber0011.Collidable = 0
  Rubber0012.Collidable = 0
  Rubber0013.Collidable = 0
  Rubber0014.Collidable = 0
  Rubber0015.Collidable = 0

  Rubber0021.Collidable = 0
  Rubber0022.Collidable = 0
  Rubber0023.Collidable = 0
  Rubber0024.Collidable = 0
  Rubber0025.Collidable = 0

  Rubber0031.Collidable = 0
  Rubber0032.Collidable = 0
  Rubber0033.Collidable = 0
  Rubber0034.Collidable = 0
  Rubber0035.Collidable = 0

  Rubber0041.Collidable = 0
  Rubber0042.Collidable = 0
  Rubber0043.Collidable = 0
  Rubber0044.Collidable = 0
  Rubber0045.Collidable = 0

  OultLaneTimer.Enabled = 0
End Sub


Sub ZoneLeftLight_Init()

End Sub

'END OF LINE...

'********************************* Game Data Tracking & Functions ****************************************
Dim gdtPlayerOneScore, gdtPlayerTwoScore, gdtPlayerThreeScore, gdtPlayerFourScore
Dim gdtCredits, gdtBallMatch
DIm gdtDict

Sub GameDataTracking_Init()

  If myPrefs_GameDataTracking = 1 Then

    gdtPlayerOneScore = Array("0","0","0","0","0","0","0")
    gdtPlayerTwoScore = Array("0","0","0","0","0","0","0")
    gdtPlayerThreeScore = Array("0","0","0","0","0","0","0")
    gdtPlayerFourScore = Array("0","0","0","0","0","0","0")
    gdtCredits = Array("0","0")
    gdtBallMatch = Array("0","0")

    Set gdtDict = CreateObject("Scripting.Dictionary")
    gdtDict.Add 0, "0" '" "
    gdtDict.Add 63, "0"
    gdtDict.Add 6, "1"
    gdtDict.Add 91, "2"
    gdtDict.Add 79, "3"
    gdtDict.Add 102, "4"
    gdtDict.Add 109, "5"
    gdtDict.Add 125, "6"
    gdtDict.Add 7, "7"
    gdtDict.Add 127,"8"
    gdtDict.Add 111,"9"

    DisplayTimer7.Enabled = True

  End If

End Sub

Sub GameDataTracking_Update(id, value)

  If myPrefs_GameDataTracking = 1 Then
    Dim chr
    If gdtDict.Exists (value) then
      chr = gdtDict.Item (value)
    Else
      chr = "0"
    End if

    Select Case True' id
    Case id < 7
      gdtPlayerOneScore(id) = chr
    Case id < 14
      gdtPlayerTwoScore(id - 7) = chr
    Case id < 21
      gdtPlayerThreeScore(id - 14) = chr
    Case id < 28
      gdtPlayerFourScore(id - 21) = chr
    Case id < 30
      gdtCredits(id - 28) = chr
    Case id < 32
      gdtBallMatch(id - 30) = chr
    End Select
  End If

End Sub

Function GameDataTracking_PlayerScore (PlayerNumber)
  If myPrefs_GameDataTracking = 1 Then
    Select Case PlayerNumber
    Case 1
      GameDataTracking_PlayerScore = Join(gdtPlayerOneScore,"")
    Case 2
      GameDataTracking_PlayerScore = Join(gdtPlayerTwoScore,"")
    Case 3
      GameDataTracking_PlayerScore = Join(gdtPlayerThreeScore,"")
    Case 4
      GameDataTracking_PlayerScore = Join(gdtPlayerFourScore,"")
    Case Else
      GameDataTracking_PlayerScore = 0
    End Select
  Else
    GameDataTracking_PlayerScore = 0
  End If
End Function

Function GameDataTracking_LeadingPlayer
  If myPrefs_GameDataTracking = 1 Then

    Dim P1, P2, P3, P4

    P1 = GameDataTracking_PlayerScore(1)
    P2 = GameDataTracking_PlayerScore(2)
    P3 = GameDataTracking_PlayerScore(3)
    P4 = GameDataTracking_PlayerScore(4)

    If P1 > P2 And P1 > P3 and P1 > P4 Then
      GameDataTracking_LeadingPlayer = 1
    ElseIf P2 > P1 And P2 > P3 and P2 > P4 Then
      GameDataTracking_LeadingPlayer = 2
    ElseIf P3 > P1 And P3 > P2 and P3 > P4 Then
      GameDataTracking_LeadingPlayer = 3
    ElseIf P4 > P1 And P4 > P2 and P4 > P3 Then
      GameDataTracking_LeadingPlayer = 4
    Else
      GameDataTracking_LeadingPlayer = 0 'tie
    End If
  Else
    GameDataTracking_LeadingPlayer = 0
  End If
End Function

Function GameDataTracking_Credits
  If myPrefs_GameDataTracking = 1 Then
    GameDataTracking_Credits = Join(gdtCredits,"")
  Else
    GameDataTracking_Credits = 0
  End If
End Function

Function GameDataTracking_BallMatch
  If myPrefs_GameDataTracking = 1 Then
    GameDataTracking_BallMatch = Join(gdtBallMatch,"")
  Else
    GameDataTracking_BallMatch = 0
  End If
End Function

'======================================================================================================
Sub myUpdateDisplayColors

    Dim myPlayerCounter, myDigitCounter, myDigitColor

  Select Case myPrefs_DisplayColor
    Case 0 'Factory
      myDigitColor = "FPScore"
    Case 1
      myDigitColor = "FPScoreBlue"
    Case 2
      myDigitColor = "FPScoreGreen"
    Case 3
      myDigitColor = "FPScoreRed"
    Case 4
      myDigitColor = "FPScoreWhite"
  End Select

  For myPlayerCounter = 1 to 4
    For myDigitCounter = 1 to 7
      eval("P" & myPlayerCounter & "D" & myDigitCounter).Image = myDigitColor
    Next
  Next
  BaD1.Image = myDigitColor
  BaD2.Image = myDigitColor
  CrD1.Image = myDigitColor
  CrD2.Image = myDigitColor

End Sub

Sub myUpdateDMDColor

  With Controller
    Select Case myPrefs_DisplayColor  'sds
      Case 0 'Factory
        .Games(cGameName).Settings.Value("dmd_red") = 255
        .Games(cGameName).Settings.Value("dmd_green") = 60
        .Games(cGameName).Settings.Value("dmd_blue") = 0
      Case 1 'Blue
        .Games(cGameName).Settings.Value("dmd_red") = 0
        .Games(cGameName).Settings.Value("dmd_green") = 0
        .Games(cGameName).Settings.Value("dmd_blue") = 255
      Case 2 'Green
        .Games(cGameName).Settings.Value("dmd_red") = 0
        .Games(cGameName).Settings.Value("dmd_green") = 255
        .Games(cGameName).Settings.Value("dmd_blue") = 0
      Case 3'Red
        .Games(cGameName).Settings.Value("dmd_red") = 255
        .Games(cGameName).Settings.Value("dmd_green") = 0
        .Games(cGameName).Settings.Value("dmd_blue") = 0
      Case 4'White
        .Games(cGameName).Settings.Value("dmd_red") = 255
        .Games(cGameName).Settings.Value("dmd_green") = 255
        .Games(cGameName).Settings.Value("dmd_blue") = 255
    End Select
  End with
End Sub

Dim myPrefs_TestWallOn
myPrefs_TestWallOn = 0 'always leave off before saving!

'*********************************************************************************************

'Dim FlipperColor
'
'Sub LoadFlipperColor
'    Dim x
'    x = LoadValue(cGameName, "FlipperColor")
'    If(x <> "")Then FlipperColor = x Else FlipperColor = 2
'    UpdateFlipperColor
'End Sub
'
'Sub SaveFlipperColor
'    SaveValue cGameName, "RFlipperColor", FlipperColor
'End Sub
'
'Sub ChangeFlipperColor
'    FlipperColor = (FlipperColor + 1)MOD 4
'    UpdateFlipperColor
'End Sub
'
'Sub UpdateFlipperColor
'    Dim x
'    Select Case FlipperColor
'        Case 0:x = RGB(0, 64, 255) 'blue
'        Case 1:x = RGB(96, 96, 96) 'White
'        Case 2:x = RGB(0, 128, 32) 'Green
'        Case 3:x = RGB(128, 0, 0)  'Red
'    End Select
'    MaterialColor "Plastic Transp Ramps1", x
'    SaveFlipperColor
'End Sub
'
'myPrefs_FlipperBatType

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
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
'            Supporting Ball, Sound Functions and Math
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
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
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
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


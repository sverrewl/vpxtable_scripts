' ****************************************************************
'                       VISUAL PINBALL X
'                Game FrameWork JP Salas using Darks DMD Flashes
'Scripting by JP Sals & RusstyT and scripting tinspred by Scotty Wickburgh
'   plain VPX script using core.vbs for supporting functions
'                         Version 3.0.10
' ****************************************************************
'This concept for this game was formed from banter between RusstyT.& HP. The conversation started with what super power we would own.
'I said that I'd control Time and HP of course would be to portal through time and mess you up.From that came endless
'scenarios and many laughs....HP asked if I could make a pinball game. 2years later and 3rd generation
'of the game here we are. HP set the bar so high and with JP,Gedankekojote & Outhere Joining us we ended up with a nice game to play in TimeLord 1
' and since then the Aago team has grown.
'desktop,cabs and cabs with DOF
'Enjoy TimeLord. It was made for fun

'*******************************************************************************************************************************
'  Credits
'GameConcept RusstyT & HP
'JP Salas Your an absolute Legend, I used a lot of your ideas Scotty Wick, Thnakyou
'Rotating Bumper caps, Dragon Nebula Spiner and tracking Dragon Eyeball ideas from Scotty Wickburghs tables
'Ramp ideas from Totan and Opening Ramp idea from Bride Of pinbot
'Gedankekojote97 applied nFozzy/Roth physics,FleepSounds, Ramp Textures & luts
'Crackers & Mozzie-Game production and testing
'DOF setup by Outhere
'Redwandass wanted to see more clocks so here they are.
'Sixtoe ideas for lights and a simplified gameplay.Great ideas.
'Alpophis did the mechanical tilt and support in the VpinHelp group. Thanks mate.
'PinStratDan table testing and gameplay ideas
'Wylte for all the final tweaks
'DonkeyKlonk GamePlayideas
'Thank RIK for the quiter Dragon fire noises
' Lots of conjecture that the game was too hard, and then it was too easy so now there is a hard easy mode.
'Thankyou to Flupper and Bord for tutorials
'Thankyou to all who made this VPX framework, Freezy and FlexDMD
'And finally JP Salas just kept helping with all the tricky scripting. Thankyou JP :)
'********************************************************************************************************************************
'V3.0.6-  3.0.7Update333
' Double lights for inserts to give light offstate sparkle
'Adusted playfield lighting to get better seperation between plastics and playfield
' Changed Luts to standard variety
'New Plastics to blue instead of purple
'Timers to prevent TimeLord Catch and reversetime bug in KickerLairOut-(TimeLord Catch)
'Reduced time TimeLord catch is enabled when a catch is made. Was 3 seconds now 0.1 seconds. Multiple bals were stopping the game.
'Adjusted Rolling ball and rubber sounds. Thanks StudlyGoorite or pointing this Out
'Removed some script that was calling for sounds that no longer are used
'Changed table sound embellishments from table to backglass
'5 second Callout overlap prevention timer now incorporated into the game.
'All TimeLord catch releases are now via the invisible spiral ramp.
'Increased the size of the extraBall kickerand extended the timer to make this easier.
' 25 Second ballsaver timer implemented for meteor shower multiball.
' Glory table time increased to 80Seconds from 60 seconds
'3.0.8-3.0.9 Serrated wall to stop infinite bumper bounce
'3.0.10 stopped wire ramp roll om opening shot to stop crash from excess ramp rolling sounds during stacked multiballs
'3.1.1b All New DMDs inin theme with new apron , AttackDragon Skillshot, Dynamic Light for all awards and award activation, Jackpot,
  ''SuperJackpot values hold for increased shots in current ball. These reset at maximum count
'4.0 Super loops can be activated from wheel at superloops light of at OOC if award hasn't been achieved. This is a significant improvement to game play.
  ' Rubber bounce turned down, flipper passing improved,
  ' New table and plastics graphics, new, backglass. Dragon roar volume down,Rubbers removed from opening shot to see graphics.
'4.0.2 The major fix in this update is Awards being lost during gameplay
'Added flags for starting and completing major awards and flags for the 4 steps to slay the dragon:
  'After TIME is complete, further advance time awards and slaying dragon awards are super JackPots.Advance time reset fixed to allow multiple advance time AwardSkillshot'
  'If an Extra ball has been made then when lit at award kicker a superjackpot is awarded
  'Fixed 5Balleasy mode defaulting to 3 ball in the hard/easy Select
  'Fixed inlane/outlane/ lights-Indicating wrong
  'Fixed Raise Ramp at next ball if ReverseTimeBall lock was ready in the previous ball
  '6 second addTimer for High score entry included to allow for glory table to complete
  'Created Ball release if ball is stuck in the dragon spinner after the wall raises
  'Fixed Multiplayer Reverse lock not activating for player new ball
  'Adjusted playfield lights to give a better contrast between on and off. Changed lane and apron lights to look more funky
  'Fixed
'4.0.2a Adjusted Wire ramp off triggers to extend to drop off point.
'4.1 Interactive backglass- Modified Dragon Spinner Ball trap, Dragon, Bird and Spider Toys added to highlight modes are active, Fixed ball lock stall after first Reversemultiball
  'Put invisible ramp and wall in plunger lane to stop balls catching under ramp, TimeWarpJackPots now SuperJackpots to make spiral ramp more in Player
  'Turned on flipper ball release. Adjusted outlane kicker enter sounds.reduced ball reflections, New VR cabinet.
  'Adjusted Camera height from 5000 to 10000 . Adjusted table GI lighting.
  'Adjusted flipper physics and table physics to allow better control and post passing
  'Flipper Start anngle reduced by 2 degrees to make shots more direct to spiral Ramp.
  'Dragon wound value increased to super jackpot. Reverse ramp during wizard mode is a super Jackpot. Changed glory billion shot from light to GloryTableOn Flag
  'Extra ball now given if the Glory Table is completed
  'Fixed 3 major bugs..Ball trapped in the dragon spinner, ball trapped in reverse lock, Ball jumping on the apron.
'4.1a MajorBugfixes in multiplayer mode,
'********************************************************************************************************************************
'101 - Flipper Left
'102 - Flipper Right
'103 - Slingshot Left
'104 - Slingshot Right
'105 - Bumper Back Left
'106 - Bumper Back Right
'107 - Bumper Back Center
'108 - Small Bumber
'109 -
'110 - Put Ball in Plunger Line ()
'111 - Drop Target Reset
'112 - Auto Fire
'113 - Kicker6
'114 - Kicker12,kickbackrg
'115 - Kicker8
'116 - WormKick4
'117 - Kicker7
'118 - Kicker11,kickbacklg
'119 - Shaker (With Nebula)
'120 - Shaker (Right Spinner)
'121 - See script
'122 - Knocker
'123 - Shaker (Left Spinner)
'124 - Beacon (With Nebula)
'125 - See script
'126 - Fan (With Nebula)

'Wylte - Segregated user options, fixed and updated ball shadows, deleted BallShadowUpdate timer(s?!), plastics added to Walls for hit sounds and given plastic physics, fixed left sleeves with peg physics,
'     de-orphaned top-right post, increased pf friction 0.015 -> 0.1, elasticity 0.02 -> 0.2, lowered flipper power and raised return strength for flipper tricks, increased kicker2 power,
'     halved plunger pull speed for DT, added credit key 2 support, halved DT nudge strength, re-normalized rolling sounds,
' PoV: x/y scale below 1, x scale just barely "flattening" circles on pf (like viewing circles edge-on IRL), inclination matches rails, FoV&Layback matched and raised as high as possible while spheres at back still spheres


Option Explicit
Randomize
Dim VRRoom
Dim Stuff
Dim UseFlexDMD
Dim GlassScratchesOn
Const BallSize = 50    ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1
Const cGameName = "TimeLord"
Const TableName = "TimeLord"
Const myVersion = "TimeLord4.0.2a"

  On Error Resume Next
  ExecuteGlobal GetTextFile("core.vbs")
  If Err Then MsgBox "Can't open core.vbs"
  ExecuteGlobal GetTextFile("controller.vbs")
  If Err Then MsgBox "Can't open controller.vbs"
  On Error Goto 0

'*********************************************************************************************
'  Player Options
'************************************************************************************************

Const FlexDMDHighQuality =  True

Const AmbientBallShadowOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const DynamicBallShadowsOn= 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that behaves like a true shadow!
                  '2 = flasher image shadow, but it moves like ninuzzu's
Dim eyeFollowS: eyeFollowS= 8 'ms timer interval of eye follow speed try 5-20 range
Const BallBright = 1      '0 - Normal, 1 - Bright

Const SongVolume = 0.4      '1 is full volume. Value is from 0 to 1
Const VolumeDial = 0.8      'Overall Mechanical sound effect volume. Recommended values should be no greater than 1
Const BallRollVolume = 0.8    'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.8    'Level of ramp rolling volume. Value between 0 and 1

Const ForceBackglass = 1    ' 1 = on ( for dt if you want the backglass on it )


'VR OPTIONS....

VRRoom = 0     ' 0 - Desktop/FS   1 - VRRoom   2 - Ultra Minimal Room
GlassScratchesOn = 1     ' 0 - Scratches OFF  1 - Scratches ON

'End VR Options
'*********************************************************************************************
'  End Player Options
'************************************************************************************************




'//////////////---- LUT (Colour Look Up Table) ----//////////////
'0 = Fleep Natural Dark 1
'1 = Fleep Natural Dark 2
'2 = VPW Original 1 to 1
'3 = Fleep Warm Dark
'4 = Fleep Warm Bright
'5 = Fleep Warm Vivid Soft
'6 = Fleep Warm Vivid Hard
'7 = Skitso Natural and Balanced
'8 = Skitso Natural High Contrast
'9 = 3rdaxis Referenced THX Standard
'10 = CalleV Punchy Brightness and Contrast
'11 = HauntFreaks Desaturated
'12 = Tomate Washed Out
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book



Dim LUTset, DisableLUTSelector, LutToggleSound, bLutActive
LutToggleSound = True
LoadLUT
'LUTset = 0     ' Override saved LUT for debug
SetLUT
DisableLUTSelector = 0  ' Disables the ability to change LUT option with magna saves in game when set to 1


If Table1.ShowDT = True then

UseFlexDMD = True
  If ForceBackglass = 1 Then
    UseFlexDMD = True '
  Else
    UseFlexDMD = False  'Dont use Flex in desktop  ... unless force bglass onto desktop
  End If

  for each Stuff in JPDMDAll: Stuff.visible = true: next  'make JP's flashers visible for desktop users
  ramp002.visible=False
  Ramp35.visible=False
Else
  if VRRoom > 0 then
    UseFlexDMD = False
    for each Stuff in JPDMDAll: Stuff.visible = true: next  'If VRroom is on (Some users DO have FS set on their VR cans), we want to force JP's Flasher DMD
  Else
    UseFlexDMD = true ' Use FlexDMD in FS mode for cabinets
  end If
End If


' Load VRRoom
If VRroom >0 Then
TimerPlunger2.enabled = true
if GlassScratchesOn = 1 then GlassImpurities.visible = true

'move the DMD into place..
for each Stuff in JPDMDAll: Stuff.x = Stuff.x +1118: next
for each Stuff in JPDMDAll: Stuff.y = Stuff.y -748: next
for each Stuff in JPDMDAll: Stuff.height = Stuff.height +315: next
for each Stuff in JPDMDAll: Stuff.rotx = 274: next

for each Stuff in VRDMDTop: Stuff.y = Stuff.y +22: next
for each Stuff in VRDMDBottom: Stuff.y = Stuff.y -20: next
'DMD done..

for each Stuff in VRCab: Stuff.visible = true: next

'make desktop sideblades and rails invisible..  why so many?
Wall35.visible = false
Wall002.visible = false
Wall35.Sidevisible = false
Wall002.Sidevisible = false
Ramp007.visible = false
Ramp35.visible = false
Ramp002.visible = false
Ramp003.visible = false
rrail.visible = false
lrail.visible = false
'VR_Backbox_Backglass.blenddisablelighting = 4

If VRRoom = 1 then
for each Stuff in VRRoomCOL: Stuff.visible = true: next  'make room stuff visible
for each Stuff in VRClock: Stuff.visible = true: next 'make clock stuff visible
ClockTimer.enabled = true
NewClockTimer.enabled = true
BeerTimer.enabled = true
end if
end if

' End VR Init.....
'********************************************************************************



Const MaxPlayers = 4     ' from 1 to 4
Const BallSaverTime = 20 ' in seconds
Const MaxMultiplier = 5  ' limit to 5x in this game, both bonus multiplier and playfield multiplier
'added by jpsalas
Dim BallsPerGame: BallsPerGame = 3   ' usually 3 or 5
Const MaxMultiballs = 9  ' max number of balls during multiball
Const tnob = 9
Const lob = 2
Const RubberizerEnabled = 1
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

'Dim tableheight: tableheight = Table1.height

'**************************
'Variables
'**************************

Dim ballrolleron
' Dim turnonultradmd
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim NextPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim PlayfieldMultiplier(4)
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot
Dim Jackpot1
Dim Jackpot2
Dim SuperJackpot
Dim SuperJackpot1
Dim SuperJackpot2
Dim SuperJackpot3
Dim Tilt
Dim MechTilt
Dim TiltSensitivity
Dim Tilted
Dim bMechTiltJustHit
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue(4)
Dim bAutoPlunger
Dim bInstantInfo
Dim bAttractMode
Dim bFlippersEnabled
Dim BallsThree:BallsThree=1
Dim BallsFive
Dim EasyHardSelected
'define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInHole
Dim BallsInLock(4)

'Define Game Flags
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim PlayerFirstBallPlayed(4)
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim bMusicOn
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim EndOfGameCheck
Dim EasyMode
Dim HardMode
Dim Hard:Hard=0
Dim OOCCount(4)
Dim AlertCallout
Dim DragonWounded
Dim Billions(4)
Dim ChaosCriticalReady(4)
Dim ChaosCriticalComplete(4)
Dim SuperLoopsReady(4)
Dim SuperLoopsComplete(4)
Dim ReverseLock1Ready(4)
Dim ReverseLock2Ready(4)
Dim ReverseLock1Complete(4)
Dim ReverseLock2Complete(4)
Dim ReverseTimeComplete(4)
Dim WormLocksActivated(4)
Dim WormLocksComplete(4)
Dim DragonWounded1(4)
Dim DragonWounded2(4)
Dim DragonWounded3(4)
Dim DragonWounded4(4)
Dim DragonWallDropActivated
Dim AdvanceTimeActivated
Dim GloryComplete(4)
Dim GloryTableOn(4)
Dim GloryTableStarted(4)
Dim EndOfGloryTable
' core.vbs variables
Dim plungerIM
Dim cbRight
Dim cbLock1
Dim cbLock2
Dim cbLock3

Dim objShell


Dim Controller
Dim B2SOn

Sub startcontroller
' Set Controller = CreateObject("B2S.Server")
' Controller.B2SName = "TimeLord"
' Controller.Run
  B2SOn = True
  bgtimer.enabled = True

End Sub




'
' Controller.B2SSetData 11,0
'1-6 front   21-32 is 1-12oclock
'1 - left
'2 - Right
'3 - staff middle
'4 - sphereleft
'5 - title
'6 - cockpit
'20 - speakers  2screen


'7-15 on the bottom   16+17+18 on top of DMD
' 7 - chaos
' 8 - backw
' 9 - sword1
'10 - sword2
'11 - middle
'12 - sword3
'13 - sword4
'14 - spidy
'15 - bird

'
Dim BG_count  : BG_count = 0
Dim BG_option : BG_option = 0
Dim BG_Wait   : BG_Wait = 0
Dim BG_Select : BG_Select = -1
Dim BG_blinks : BG_blinks = 0
Dim BG_Spin   : BG_Spin = 0

Dim BG_Clock2 : BG_Clock2 = 1
Dim BG_Clock_Time: BG_Clock_Time = 0
Dim BG_Clock_Dir: BG_Clock_Dir = 1  ' 1 or -1
Dim BG_CLight : BG_CLight = 0
Dim BG_attractCount
Dim BG_GloryT_Count
Dim BG_ALL_OFF


Dim BG_Anim ( 50,5)
Dim BGArray(42,10)
dim x
dim y
for x = 1 to 42

  for y = 1 to 10
  BGArray(x,y) = 0
  Next
  BGArray(x,0) = -1
  If x < 21 Then BGArray(x,2) = 1

Next
BGArray(3,2) = 1
BGArray(4,2) = 0
BGArray(19,2) = 0

'0 - last state ( -1 startup )
'1 - new state ( 0 at startup )
'2 - auto state ( last pri )
'3 - auto waittime
'4 - pri 1
'5 - pri 2
'6 - pri 3
'7 - pri 4
'8 - pri 5

'285= 1sec

Dim BG_attractwiz : BG_attractwiz = 1
Dim BG_EscapeCounter : BG_EscapeCounter = 0
bgtimer.interval = 27
Sub bgtimer_timer
  dim x
  dim tmp,tmp2

  If BG_ALL_OFF > 0 Then BG_ALL_OFF = BG_ALL_OFF - 1

  For x = 1 to 32
    BGArray(x,3) = BGArray(x,3) + 1
    tmp = 0
    If Tilted or BG_ALL_OFF > 0 Then
      tmp = 0
    ElseIf bAttractMode = True Then
      If BGArray(x,7) = 1 Then tmp = 1
    ElseIf BGArray(x,4) > 0 Then
      If BGArray(x,4) = 1 Then tmp = 1
    Elseif BGArray(x,5) > 0 Then
      If BGArray(x,5) = 1 Then tmp = 1
    Elseif BGArray(x,6) > 0 And BGArray(x,3) < 21 Then
      If BGArray(x,6) = 1 Then tmp = 1
    Elseif BGArray(x,2) > 0 And BGArray(x,3) > 20 Then
      tmp = 1
    End If
    BGArray(x,1) = tmp

    If BGArray(x,1) <> BGArray(x,0) And x <> 3 and x <> 4 Then
      BGArray(x,0) = BGArray(x,1)
      Controller.B2SSetData x,BGArray(x,0)
    End If
  Next

  For x = 33 to 42
      If BGArray(x,1) <> BGArray(x,0) Then
        BGArray(x,0) = BGArray(x,1)
        Controller.B2SSetData x,BGArray(x,0)
      End If
  Next

  tmp = 0
  If BGArray(3,1) = 1 Then tmp = 1
  tmp2 = 0
  If BGArray(4,1) = 1 Then tmp2 = 1
  If Tilted or BG_ALL_OFF > 0 Then tmp = 0 : Tmp2 = 0

  For x = 33 to 42 : BGArray(x,1) = 0 : Next

  If Robot.Visible Then ' Robot
    If tmp = 0 Then
      BGArray(36,1) = 1
    Else
      BGArray(37,1) = 1
    End If

  ElseIf QMan.Visible Then ' blue
    If tmp = 0 Then
      BGArray(33,1) = 1
    Else
      BGArray(34,1) = 1
    End If
    If tmp2 = 0 Then
      BGArray(39,1) = 1
    Else
      BGArray(42,1) = 1
    End If

  Else ' ww middle
    If tmp = 0 Then
      BGArray(35,1) = 1
    Else
      BGArray(40,1) = 1
    End If
    If tmp2 = 0 Then
      BGArray(38,1) = 1
    Else
      BGArray(41,1) = 1
    End If

  End If


' Robot.Visible=False
' QMan.Visible=1


'fixing
' MULTIBALL BOTH WIZ BLINK

' show wizards !
' BGArray(3,1) = Blinking wizards
' BGArray(4,1) = Blinking smal ones if visible ( remove bg light and add smal ON images !
'
'Put flags to switch wizard characters in the QManVisible, QManInvisible, TimethiefVisible and Timethief invisible subroutines.
'Note use the GloryComplete=1 flag to make the Robots visible and Wizards invisible. This is already scripted in the game.




  If bAttractMode = True Then
    bgtimer.interval = 45

    BG_attractCount  = BG_attractCount + 1
    Select Case BG_attractCount
      case 20,40,60,80,260,280 : For x = 1 to 32 : BGArray ( x,7) = 1 : Next
      case 30,50,70,90,270,290 : For x = 1 to 32 : BGArray ( x,7) = 0 : Next
      case 110 : BGArray ( 5,7) = 1
      case 130 : BGArray ( 1,7) = 1 : BGArray ( 2,7) = 1 : BGArray ( 6,7) = 1
      case 150 : BGArray ( 3,7) = 1 : BGArray ( 4,7) = 1
      case 170,300,320,340 : For x = 21 to 32 : BGArray ( x,7) = 1 : Next
      case     310,330,350 : For x = 21 to 32 : BGArray ( x,7) = 0 : Next
      case 190 : BGArray ( 20,7) = 1
      case 210 : BGArray ( 16,7) = 1 : BGArray ( 17,7) = 1 : BGArray ( 18,7) = 1
      case 230 : For x = 7 to 15 : BGArray ( x,7) = 1 : Next

      case 300 : For x = 21 to 32 : BGArray ( x,7) = 1 : Next

      case 510 : BG_attractCount = -120
            For x = 1 to 32 : BGArray ( x,7) = 0 : Next
            BG_attractwiz = BG_attractwiz + 1
            Select Case BG_attractwiz
              case 1 : robot.visible = False : qman.visible = False : TimeThief.Visible = true
              case 2 : robot.visible = False : qman.visible = true : TimeThief.Visible = False
              case 3 : robot.visible = true : qman.visible = False : TimeThief.Visible = False
              case 4 : robot.visible = False : qman.visible = False : TimeThief.Visible = true : BG_attractwiz = 1
            End Select
    End Select
    If BG_attractCount > 450 Then
      For x = 1 to 12
        BGArray (int(rnd(1)*33),7) = int(rnd(1)*2)
      Next
    End If
    If BG_attractCount < -50 Then
        BGArray (5,7) = int(rnd(1)*2)
    End If

  Else
    bgtimer.interval = 21
    ' clock 1sec on !
    BG_Clock_Time = BG_Clock_Time + 1
    If BG_Clock_Time > 46 then
      BG_Clock_Time = 0
      BGArray( 20 + BG_Clock2 ,2) = 0
      BG_Clock2 = BG_Clock2 + BG_Clock_Dir
      If BG_Clock2 > 12 Then BG_Clock2 = 1
      If BG_Clock2 < 1 Then BG_Clock2 = 12
      BGArray( 20 + BG_Clock2 ,2) = 1
    End If

' dmd swords
    BGarray(9,2)  = -l42.getinplaystatebool
    BGarray(10,2) = -l43.getinplaystatebool
    BGarray(12,2) = -l44.getinplaystatebool
    BGarray(13,2) = -l45.getinplaystatebool

' dmd orbs

    If l18.State=2 Then
      BGarray(7,2)=-l18.getinplaystatebool
    Else
      BGarray(7,2)=-l38.getinplaystatebool
    End If

    If l21A.State=2 Then
      BGarray(8,2)= -l21A.getinplaystatebool
    Else
      BGarray(8,2)= -l39.getinplaystatebool
    End If

    If l75.state = 2 Then
      BGarray(11,2) = -l75.getinplaystatebool
      BGarray(20,2) = -l75.getinplaystatebool
    Else
      BGarray(11,2) = -l45.getinplaystatebool
      BGarray(20,2) = -l45.getinplaystatebool
    End If

    BGarray(18,2) = -l24.getinplaystatebool
    If l24.state = 2 Then
      BGarray(14,2) = -l24.getinplaystatebool
      BGarray(18,2) = -l24.getinplaystatebool
    Else
      BGarray(14,2) = -l40.getinplaystatebool  '21a blinks then use that and else if 40 is on then use that
      BGarray(18,2) = 0
    End If
    If l50.state = 2 Then
      BGarray(15,2) = -l50.getinplaystatebool
      BGarray(16,2) = -l50.getinplaystatebool
    Else
      BGarray(15,2) = -l41.getinplaystatebool 'l50 blinks when active and l41 is on when made
      BGarray(16,2) = -l41.getinplaystatebool
    End If

    If l18.state=2 or l75.state=2 or l24.state=2 or l50.state=2 or l28.state=2 or l49.state=2 or l36.state=2 or l37.state=2 Then
      BGarray(4,6) = 1 - CLightbg2.getinplaystatebool
    Else
      BGarray(4,6) = 0
    End If

    BGarray(4,4) = 0 : BGarray(3,4) = 0
    If bMultiBallMode Then
'     BGarray(3,4) = 1 - CLightbg.getinplaystatebool
      BGarray(4,4) = 1 - CLightbg2.getinplaystatebool   ' only qman blink
    ElseIf GloryTableOn(CurrentPlayer)=True Then
      BG_GloryT_Count = BG_GloryT_Count + 1
      If BG_GloryT_Count mod 20 > 18 Then
        BGarray(4,4) = 1
      Else
        BGarray(4,4) = 2
      End If
      If BG_GloryT_Count mod 10 > 6 Then
        BGarray(3,4) = 1
      Else
        BGarray(3,4) = 2
      End If
    End If

    If l61.state = 2 Then   ' backwards time
      BG_Clock_Dir = -1
    Else
      BG_Clock_Dir = 1
    End If

    If Light_Escape.State=2 Then
      BG_EscapeCounter = BG_EscapeCounter + 1
      BGArray ( 1,1) = 1 - CLightbg.getinplaystatebool
      BGArray ( 2,1) = 1 - CLightbg.getinplaystatebool
      BGArray ( 6,1) = 1 - CLightbg.getinplaystatebool
    Else
      BGArray ( 1,1) = 0
      BGArray ( 2,1) = 0
      BGArray ( 6,1) = 0
    End If
  End If







  If BG_Select <> -1 Then
    BG_option = BG_Select
    BG_Select = -1
    BG_count = 0
    BG_All 0 ' turn all off
    exit sub
  End If

  Select Case BG_option
    Case 0 ' auto startup goto 1 after wait
      BG_count = BG_count + 1
      If BG_count = 111 Then BG_option = 1 : BG_count = 0 : BG_All 0   ' cnage option to 1 fixing

      If BG_count = 33 Then   BG_All 1
    case 1: 'startup
      BG_count = BG_count + 1
      Select Case BG_count
        Case 10,30,50,70,90 : ' OnS
          BG_All 1
        Case 20,40,60,80,100 : ' off
          BG_All 2
        case 105
          BG_option = -1 : BG_count = 0
          BG_All 0
      End Select
  End Select




  If BG_anim(1,1) > 0 Then BG_1 ' pri 6   clock spinCW
  If BG_anim(2,1) > 0  Then BG_2 ' pri 6   clock spinCCW
  If BG_anim(3,1) > 0 Then BG_3 ' pri 6   clock blink XX times
  If BG_anim(4,1) > 0 Then BG_4 ' pri 6   clock blink XX times
  If BG_anim(5,1) > 0 Then BG_5 ' pri 6   All BG blink XX times
  If BG_anim(6,1) > 0 Then BG_6 ' pri 6   All BG blink XX times
  If BG_anim(7,1) > 0 Then BG_7 ' pri 6   All  off turn ccw
  If BG_anim(33,1) > 0 Then BG_33 ' pri 6   titles blink in timer
  If BG_anim(34,1) > 0 Then BG_34 ' pri 6  LEFTLANES
  If BG_anim(35,1) > 0 Then BG_35 ' pri 6  RIGHTLANES
  If BG_anim(36,1) > 0 Then BG_36 ' pri 6  COCKPIT
  If BG_anim(38,1) > 0 Then BG_38 ' pri 6  BLUEWIZ
  If BG_anim(39,1) > 0 Then BG_39 ' pri 6  wWIZ
  If BG_anim(37,1) > 0 Then BG_37 ' pri 6  COCKPIT fast



'   BG_anim(nr,1) = 1
'   BG_anim(nr,2) = blinks
'   BG_anim(nr,2) = time





End Sub
' Attractmode thing override ( not made )
'
'
'
'

' BG CLoCK during gameplay auto about every 1 sec
'BG_Clock_Dir = 1 ' BG clock   CW
'BG_Clock_Dir = -1  ' BG clock   CCW

' BG_Sequence ( nr , blinks , time )
'BG_Sequence 1,0,200    'pri 6    clock spinCW  200updates
'BG_Sequence 2,0,200    'pri 6    clock spinCCW   200updates
'BG_Sequence 3,10,0     'pri 6    clock blinks x times
'BG_Sequence 4,0,200    'pri 6    clock blinks following blink pattern on CLightbg
'BG_Sequence 5,10,0     'pri 6    All BG blinks x times
'BG_Sequence 6,0,200    'pri 6    All BG blinks following blink pattern on CLightbg
'BG_Sequence 7,0,200    'pri 4    All off  turn ccw   2 extended


'BG_Sequence 33,0,100   'pri 6.NOSTOP title blinks following blink pattern on CLightbg
'BG_Sequence 34,0,100   'pri 6.NOSTOP   LEFT blinks following blink pattern on CLightbg
'BG_Sequence 35,0,100   'pri 6.NOSTOP   rIGHT blinks following blink pattern on CLightbg
'BG_Sequence 36,0,100   'pri 6.NOSTOP   COCKPIT
'BG_Sequence 37,0,110   'pri 6.NOSTOP   COCKPIT fast blink  takethepiss 1+2
'BG_Sequence 38,0,200   'pri 6    blue wiz blinks  Fast
'BG_Sequence 39,0,200   'pri 6    w wiz blinks  Fast





'will add more here


Sub BG_Sequence ( nr , blinks , time )
  Dim x

' If BG_option < BG_Select and BG_option > -1 Then

'   For x = 1 to 32 : BGArray(x,4) = 0 : BGArray(x,5) = 0 : BGArray(x,6) = 0 : Next

    Select Case nr
      Case 1,2  ' PRI 6
        BG_Stop_Pri6

        BG_anim(nr,1) = 1
        BG_anim(nr,2) = blinks * 2
        BG_anim(nr,3) = time
      Case 3
        BG_Stop_Pri6

        BG_anim(nr,1) = 1
        BG_anim(nr,2) = blinks * 2
        BG_anim(nr,3) = time
        CLightbg.blinkpattern = "10"
        CLightbg.blinkinterval = 120
      Case 4
        BG_Stop_Pri6

        BG_anim(nr,1) = 1
        BG_anim(nr,2) = blinks * 2
        BG_anim(nr,3) = time
        CLightbg.blinkpattern = "1010010001001010"
        CLightbg.blinkinterval = 70
        BG_Stop_Pri6

        BG_anim(nr,1) = 1
        BG_anim(nr,2) = blinks * 2
        BG_anim(nr,3) = time
        CLightbg.blinkpattern = "1010010001001010"
        CLightbg.blinkinterval = 70
      Case 5
        BG_Stop_Pri6

        BG_anim(nr,1) = 1
        BG_anim(nr,2) = blinks * 2
        BG_anim(nr,3) = time
        CLightbg.blinkpattern = "10"
        CLightbg.blinkinterval = 120
      Case 6
        BG_Stop_Pri6

        BG_anim(nr,1) = 1
        BG_anim(nr,2) = blinks * 2
        BG_anim(nr,3) = time
        CLightbg.blinkpattern = "1010010001001010100100"
        CLightbg.blinkinterval = 70
      Case 7
        BG_anim(nr,1) = 1
        BG_anim(nr,2) = blinks * 2
        BG_anim(nr,3) = time
      Case 33,34,35,36,37,38,39,40,41,42,43,44,45,46.47
        BG_anim(nr,1) = 1
        BG_anim(nr,2) = blinks * 2
        BG_anim(nr,3) = time
    End Select
' End If

End Sub

Sub BG_Stop_Pri6 : For x = 1 to 32 : BGArray(x,6) = 0 : BG_anim(x,1) = 0 : Next : End Sub
Sub BG_Clock(nr) : For x = 21 to 32 : BGArray ( x,6) = nr : BGArray(x,3) = 0 : Next : End Sub
Sub BG_All(nr) : For x = 1 to 32 : BGArray ( x,6) = nr : BGArray(x,3) = 0 : Next : End Sub

' 4 = Pri 1


' 5 = Pri 2







' 6 = Pri 3


Sub BG_1
  BG_anim(1,1) = BG_anim(1,1) + 1
    If BG_anim(1,1) > BG_anim(1,3) then ' end SpinCW
      BG_anim(1,1) = 0 : BG_Clock 0 : Exit Sub
    ElseIf BG_anim(1,1) mod 2 = 1 Then
      BG_Spin = BG_Spin + 1
      If BG_Spin > 12 Then BG_Spin = 1
    End If
    For x = 1 to 12
      If x = BG_Spin Then
        BGArray ( x+20,6 ) = 1
        BGArray ( x+20,3 ) = 0 ' reset Auto
      Else
        BGArray ( x+20,6 ) = 2
        BGArray ( x+20,3 ) = 0 ' reset Auto
      End If
    Next
End Sub

Sub BG_2
    BG_anim(2,1) = BG_anim(2,1) + 1
    If BG_anim(2,1) > BG_anim(2,3) then ' end SpinCCW
      BG_anim(2,1) = 0 : BG_Clock 0 : Exit Sub
    ElseIf BG_anim(2,1) mod 2 = 1 Then
      BG_Spin = BG_Spin - 1
      If BG_Spin < 1 Then BG_Spin = 12
    End If
    For x = 1 to 12
      If x = BG_Spin Then
        BGArray ( x+20,6 ) = 1
        BGArray ( x+20,3 ) = 0 ' reset Auto
      Else
        BGArray ( x+20,6 ) = 2
        BGArray ( x+20,3 ) = 0 ' reset Auto
      End If
    Next
End Sub

Sub BG_7   ' pri 4

    BG_anim(7,1) = BG_anim(7,1) + 1
    If BG_anim(7,1) > BG_anim(7,3) then ' end SpinCCW
      BG_anim(7,1) = 0
      For x = 1 to 40 : BGArray ( x,4) = 0 : BGArray(x,3) = 0 : Next
      Exit Sub
    ElseIf BG_anim(7,1) mod 2 = 1 Then
      BG_Spin = BG_Spin - 1
      If BG_Spin < 1 Then BG_Spin = 12
    End If
    For x = 1 to 12
      If x = BG_Spin Then
        BGArray ( x+20,4 ) = 1
        BGArray ( x+20,3 ) = 0 ' reset Auto
      Else
        BGArray ( x+20,4 ) = 2
        BGArray ( x+20,3 ) = 0 ' reset Auto
      End If
    Next
End Sub

Sub BG_3
    BG_anim(3,1) = BG_anim(3,1) + 1
    If BG_anim(3,2) < 0 then
      BG_anim(3,1) = 0 : BG_Clock 0 : Exit Sub
    End If
    If ( BG_anim(3,1) mod 12 ) = 2 Then
      BG_anim(3,2) = BG_anim(3,2) - 1
      BG_Clock 2
    Elseif ( BG_anim(3,1) mod 12 ) = 8 Then
      BG_anim(3,2) = BG_anim(3,2) - 1
      BG_Clock 1
    End If
End Sub

Sub BG_4
    BG_anim(4,1) = BG_anim(4,1) + 1
    If BG_anim(4,1) > BG_anim(4,3) then

      BG_anim(4,1) = 0 : BG_Clock 0 : Exit Sub
    End If
    For x = 1 to 12
      BGArray ( x+20,6 ) = 1 - CLightbg.getinplaystatebool
    Next
End Sub
Dim BG_addshow : BG_addshow = 0
Sub BG_5
    BG_anim(5,1) = BG_anim(5,1) + 1
    If BG_anim(5,2) < 0 then
      BG_anim(5,1) = 0 : BG_All 0
      If BG_addshow = 1 Then BG_addshow = 0 : BG_Sequence 1,0,200
      Exit Sub
    End If
    If ( BG_anim(5,1) mod 12 ) = 2 Then
      BG_anim(5,2) = BG_anim(5,2) - 1
      BG_All 2
    Elseif ( BG_anim(5,1) mod 12 ) = 8 Then
      BG_anim(5,2) = BG_anim(5,2) - 1
      BG_All 1
    End If
End Sub

Sub BG_6
    BG_anim(6,1) = BG_anim(6,1) + 1
    If BG_anim(6,1) > BG_anim(6,3) then
      BG_anim(6,1) = 0 : BG_All 0 : Exit Sub
    End If
      BG_All 1 - CLightbg.getinplaystatebool
End Sub

Sub BG_33
  BG_anim(33,1) = BG_anim(33,1) + 1
  If BG_anim(33,1) > BG_anim(33,3) then
    BGArray ( 5,6) = 0
    BGArray (17,6) = 0
    BG_anim(33,1) = 0 : Exit Sub
  End If
  BGArray ( 5,6) = 1 - CLightbg.getinplaystatebool : BGArray( 5,3) = 0
  BGArray (17,6) = 1 - CLightbg.getinplaystatebool : BGArray(17,3) = 0
End Sub
Sub BG_34
  BG_anim(34,1) = BG_anim(34,1) + 1
  If BG_anim(34,1) > BG_anim(34,3) then
    BGArray ( 1,6) = 0
    BG_anim(34,1) = 0 : Exit Sub
  End If
  BGArray ( 1,6) = 1 - CLightbg.getinplaystatebool : BGArray( 1,3) = 0
End Sub
Sub BG_35
  BG_anim(35,1) = BG_anim(35,1) + 1
  If BG_anim(35,1) > BG_anim(35,3) then
    BGArray ( 2,6) = 0
    BG_anim(35,1) = 0 : Exit Sub
  End If
  BGArray ( 2,6) = 1 - CLightbg.getinplaystatebool : BGArray( 2,3) = 0
End Sub
Sub BG_36
  BG_anim(36,1) = BG_anim(36,1) + 1
  If BG_anim(36,1) > BG_anim(36,3) then
    BGArray ( 6,6) = 0
    BG_anim(36,1) = 0 : Exit Sub
  End If
  BGArray ( 6,6) = 1 - CLightbg.getinplaystatebool : BGArray( 6,3) = 0
End Sub

Sub BG_37 '
  BG_anim(37,1) = BG_anim(37,1) + 1
  If BG_anim(37,1) > BG_anim(37,3) then
    BGArray ( 6,2) = 0
    BG_anim(37,1) = 0 : Exit Sub
  End If

  If BG_anim(37,1) mod 9 = 3  then BGArray ( 6,2) = 1 ' 1-32 = need x,2
  If BG_anim(37,1) mod 9 = 7 then BGArray ( 6,2) = 0
End Sub

Sub BG_38 '
  BG_anim(38,1) = BG_anim(38,1) + 1
  If BG_anim(38,1) > BG_anim(38,3) then
    BGArray ( 4,2) = 0
    BG_anim(38,1) = 0 : Exit Sub
  End If

  If BG_anim(38,1) mod 9 = 3  then BGArray ( 4,2) = 1 ' 1-32 = need x,2
  If BG_anim(38,1) mod 9 = 7  then BGArray ( 4,2) = 0
End Sub

Sub BG_39 '
  BG_anim(39,1) = BG_anim(39,1) + 1
  If BG_anim(39,1) > BG_anim(39,3) then
    BGArray ( 3,2) = 0
    BG_anim(39,1) = 0 : Exit Sub
  End If

  If BG_anim(39,1) mod 9 = 3  then BGArray ( 3,2) = 1 ' 1-32 = need x,2
  If BG_anim(39,1) mod 9 = 7  then BGArray ( 3,2) = 0
End Sub




'
' Controller.B2SSetData 11,0
'1-6 front   21-32 is 1-12oclock
'1 - left
'2 - Right
'3 - staff middle
'4 - sphereleft
'5 - title
'6 - cockpit
'19 - bald wizard
'20 - speakers  2screen


'7-15 on the bottom   16+17+18 on top of DMD
' 7 - chaos
' 8 - backw
' 9 - sword1
'10 - sword2
'11 - middle
'12 - sword3
'13 - sword4
'14 - spidy
'15 - bird




' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()


If Table1.ShowDT = True then
  If ForceBackglass = 1 Then startcontroller
Else
  if VRRoom > 0 then

  Else
    startcontroller
  end If
End If


  LoadLUT
  '   resetbackglass
  LoadEM
  Dim i
  Randomize

  'AutoPlunger

  Const IMPowerSetting = 10 ' Plunger Power
  Const IMTime = 1.1        ' Time in seconds for Full Plunge
  Set plungerIM = New cvpmImpulseP
  With plungerIM
    .InitImpulseP swPlunger, IMPowerSetting, IMTime
    .Random 1.5
    .InitExitSnd SoundFX("Popper", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
    .CreateEvents "plungerIM"
  End With

  'Captive Ball

  Set cbRight = New cvpmCaptiveBall
  With cbRight
    .InitCaptive CapTrigger1, CapWall1, Array(CapKicker1, CapKicker1a), 0
    .NailedBalls = 1
    .ForceTrans = .9
    .MinForce = 3.5
    .CreateEvents "cbRight"
    .Start
  End With
  CapKicker1.CreateBall

  ' Misc. VP table objects Initialisation, droptargets, animations...
  VPObjects_Init

  ' load saved values, highscore, names, jackpot
  Loadhs

  ' Initalise the DMD display
  DMD_Init

  ' freeplay or coins
  bFreePlay = False 'we want coins
  if bFreePlay Then DOF 125, DOFOn

  '   Loadhs
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
  BallsInHole = 0
  BallsInLock(1) = 0
  BallsInLock(2) = 0
  LastSwitchHit = ""
  Tilt = 0
  MechTilt = 0
  TiltSensitivity = 6
  Tilted = False
  bBonusHeld = False
  bJustStarted = True
  FlasherEasySelect3.Visible=1
  EasyMode=True
  HardMode=False
  Hard=0

  WallJamStopper.isdropped=True
  GiOff
  DMDFlush
  StartAttractMode


End Sub

'**************************************
'Object Initiation   See Serious Sam
'**************************************
' droptargets, animations, etc
Sub VPObjects_Init
End Sub

'********************
' MATHS
'********************

Function RndNum(min,max)
  RndNum = Int(Rnd()*(max-min+1))+min     ' Sets a random number between min AND max
End Function


'   Dim BIP
'   BIP = 0

'*****************************************
'      Moving Eyes
'*****************************************

Sub Lighteyeball3_Timer ' DragonEye leftflash
  If Lighteyeball3.timerenabled=0 Then
    Lighteyeball3.timerenabled=1
    Lighteyeball3.uservalue=0

  Else
Dim i
    i = Lighteyeball3.uservalue
    Select Case i
      Case 0 : Lighteyeball3.intensity=12
      Case 1 : Lighteyeball3.intensity=11
      Case 2 : Lighteyeball3.intensity=10
      Case 3 : Lighteyeball3.intensity=8
      Case 4 : Lighteyeball3.intensity=4
      Case 5 : Lighteyeball3.intensity=2
      Case 6 : Lighteyeball3.intensity=0
           Lighteyeball3.timerenabled=0
    End Select
    Lighteyeball3.uservalue=i+1
  End If
End Sub

Sub Lighteyeball4_Timer ' DragonEye Rightflash
  If Lighteyeball4.timerenabled=0 Then
    Lighteyeball4.timerenabled=1
    Lighteyeball4.uservalue=0

  Else
Dim i
    i = Lighteyeball4.uservalue
    Select Case i
      Case 0 : Lighteyeball4.intensity=12
      Case 1 : Lighteyeball4.intensity=11
      Case 2 : Lighteyeball4.intensity=10
      Case 3 : Lighteyeball4.intensity=8
      Case 4 : Lighteyeball4.intensity=4
      Case 5 : Lighteyeball4.intensity=2
      Case 6 : Lighteyeball4.intensity=0
           Lighteyeball4.timerenabled=0
    End Select
    Lighteyeball4.uservalue=i+1
  End If
End Sub

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
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = lob - 1 Then Exit Sub

  ' eyes follow the ball
  eyeballmod.RotZ = (600 - BOT(2).X)\6 'BOT(2) is the first playing ball due to at there are 2 locked balls
  eye2ballmod.RotZ = (600 - BOT(2).X)\6

  '***Ball Drop Sounds***

  '   If BOT(b).VelZ < -1 and BOT(b).z < 50 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
  '     PlaySoundAtBOTBallZ "fx_ballrampdrop" & b, BOT(b)
  '   End If

  '   debug.print BOT(b).x & " " & BOT(b).y
  '   if SQR((BOT(b).VelX ^2) + (BOT(b).VelY ^2)) < 0.2 and InRect(BOT(b).x, BOT(b).y, 420,290,440,290, 440, 300, 420,300) Then
  '     BOT(b).vely = 5
  '   end if
End Sub

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

'//////////////////////////////////////////////////////////////////////
'// LUT
'//////////////////////////////////////////////////////////////////////

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
    Case 0: LUTBox.text = "VPW original 1on1"
    Case 1: LUTBox.text = "Fleep Natural Dark 1"
    Case 2: LUTBox.text = "Fleep Natural Dark 2"
    Case 3: LUTBox.text = "Fleep Warm Dark"
    Case 4: LUTBox.text = "Fleep Warm Bright"
    Case 5: LUTBox.text = "Fleep Warm Vivid Soft"
    Case 6: LUTBox.text = "Fleep Warm Vivid Hard"
    Case 7: LUTBox.text = "Skitso Natural and Balanced"
    Case 8: LUTBox.text = "Skitso Natural High Contrast"
    Case 9: LUTBox.text = "3rdaxis Referenced THX Standard"
    Case 10: LUTBox.text = "CalleV Punchy Brightness and Contrast"
    Case 11: LUTBox.text = "HauntFreaks Desaturated"
    Case 12: LUTBox.text = "Tomate washed out"
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

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "TimeLordLUT.txt",True)
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
  If Not FileObj.FileExists(UserDirectory & "TimeLordLUT.txt") then
    LUTset=0
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "TimeLordLUT.txt")
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

'**************************
'   KEYS
'**************************

Sub Table1_KeyDown(ByVal Keycode)

'added by jpsalas
If NOT bGameInPlay Then
    if keycode = RightFlipperKey Then BallSelection = (BallSelection + 1) MOD 4: UpdateBallSelection
    if keycode = LeftFlipperKey Then
       BallSelection = BallSelection - 1
       If BallSelection = -1 then BallSelection = 3
       UpdateBallSelection
    End If
End If

  if keycode = RightFlipperKey and VRRoom > 0 then VRFlipperRight.x = VRFlipperRight.x -5
  if keycode = LeftFlipperKey and VRRoom > 0 then VRFlipperLeft.x = VRFlipperLeft.x +5

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

  If Keycode = AddCreditKey Or Keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
    If Credits < 15 Then: Credits = Credits + 1
    if bFreePlay = False Then DOF 125, DOFOn
    If(Tilted = False) Then
      DMDFlush
      DMD "_", CL(1, "CREDITS: " & Credits), "", eNone, eNone, eNone, 500, True,""

      If NOT bGameInPlay Then ShowTableInfo
    End If
  End If


If keycode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull()
  If VRroom > 0 then
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End if
End If

  If bGameInPlay Then
      If keycode = LeftTiltKey Then Nudge 90, 2.5:SoundNudgeLeft():CheckTilt
      If keycode = RightTiltKey Then Nudge 270, 2.5:SoundNudgeRight():CheckTilt
      If keycode = CenterTiltKey Then Nudge 0, 1.5:SoundNudgeCenter():CheckTilt
      If keycode = MechanicalTilt Then SoundNudgeCenter:CheckMechTilt
      '********************************************************************************************************************
      If keycode = LeftFlipperKey and bFlippersEnabled Then FlipperActivate LeftFlipper, LFPress:InstantInfoTimer.Enabled = True:SolLFlipper 1
      If keycode = RightFlipperKey and bFlippersEnabled Then FlipperActivate RightFlipper, RFPress:InstantInfoTimer.Enabled = True:SolRFlipper 1
      '**********************************************************************************************************************
      If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
      End If

      If keycode = StartGameKey Then
        If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = True))Then

          If(bFreePlay = True)Then
            PlayersPlayingGame = PlayersPlayingGame + 1



            TotalGamesPlayed = TotalGamesPlayed + 1

            DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eNone, eNone, 500, True, "tl_fanfare1"

          Else
            If(Credits > 0)then
              PlayersPlayingGame = PlayersPlayingGame + 1
              TotalGamesPlayed = TotalGamesPlayed + 1
              Credits = Credits - 1
              DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eNone, eNone, 500, True, "tl_fanfare1"
debug.print "PlayersPlayingGame" & PlayersPlayingGame
            Else
              ' Not Enough Credits to start a game.
              DOF 140, DOFOff
              DMD CL(0, "CREDITS " & Credits), CL(1, "INSERT COIN"), "", eNone, eBlink, eNone, 500, True, ""
              PlaySound "CO_NoCoin"
            End If
          End If
        End If
      End If
  Else ' If (GameInPlay)

    If keycode = StartGameKey Then
      If(bFreePlay = True)Then
        If(BallsOnPlayfield = 0)Then
          ResetForNewGame()

        End If
      Else
        If(Credits > 0)Then
          If(BallsOnPlayfield = 0)Then
            Credits = Credits - 1
            ResetForNewGame()

          End If
        Else
          ' Not Enough Credits to start a game.
          DOF 140, DOFOff
          DMDFlush
          DMD CL(0, "CREDITS " & Credits), CL(1, "INSERT COIN"), "", eNone, eBlink, eNone, 500, True, ""
          PlaySound "CO_NoCoin"
        End If
      End If
    End If

  End If ' If (GameInPlay)
End Sub

Sub PlayersCall
  DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eNone, eNone, 500, True, "tl_fanfare1"
End Sub

'added by jpsalas
Dim BallSelection: BallSelection = 0

Sub UpdateBallSelection
DMDFlush
Dim Hard




Select Case BallSelection
    Case 0: 'Easy3
           DMD CL(0, "EASY MODE " ), CL(1, "3 BALLS"), "", eNone, eNone, eNone, 1500, True, ""
           SelectDifficutly 1
      Hard=0:BallsThree=True:BallsFive=False
           EasyMode=True:HardMode=False
       PlaySound "CO_Easy3"
           FlasherHardSelect5.Visible=0
           FlasherHardSelect3.Visible=0
           FlasherEasySelect5.Visible=0
           FlasherEasySelect3.Visible=1
           BallsPerGame = 3
    Case 1: 'Easy5
           DMD CL(0, "EASY MODE " ), CL(1, "5 BALLS"), "", eNone, eNone, eNone, 1500, True, ""
           SelectDifficutly 1
      Hard=0:BallsThree=False:BallsFive=True
           EasyMode=True:HardMode=False
       PlaySound "CO_Easy5"
           FlasherHardSelect5.Visible=0
           FlasherHardSelect3.Visible=0
           FlasherEasySelect5.Visible=1
           FlasherEasySelect3.Visible=0
           BallsPerGame = 5

    Case 2: 'Hard3
           DMD CL(0, "HARD MODE " ), CL(1, "3 BALLS"), "", eNone, eNone, eNone, 1500, True, ""
           SelectDifficutly 2
      Hard=1:BallsThree=True:BallsFive=False
           HardMode=True:EasyMode=False
       PlaySound "CO_Hard3"
           FlasherHardSelect5.Visible=0
           FlasherHardSelect3.Visible=1
           FlasherEasySelect5.Visible=0
           FlasherEasySelect3.Visible=0
           BallsPerGame = 3

    Case 3: 'Hard5
           DMD CL(0, "HARD MODE " ), CL(1, "5 BALLS"), "", eNone, eNone, eNone, 1500, True, ""
           SelectDifficutly 2
      Hard=1:BallsThree=False:BallsFive=True
          HardMode=True:EasyMode=False
       PlaySound "CO_Hard5"
           FlasherHardSelect5.Visible=1
           FlasherHardSelect3.Visible=0
           FlasherEasySelect5.Visible=0
           FlasherEasySelect3.Visible=0
           BallsPerGame = 5
End Select
ShowTableInfo
End Sub


'*****************************************************Not sure how to do this

Sub SelectDifficutly(easy_hard)
  Select Case easy_hard
    Case 1: Hard=0
    Case 2: Hard=1
  End Select
End Sub

'***********************************************************************************************************************

'********************************************************************************************************


Sub DMDEasyMode
  EasyMode=True
End Sub


Sub DMDHardMode
  HardMode=True
End Sub

Sub Table1_KeyUp(ByVal keycode)

if keycode = RightFlipperKey and VRRoom > 0 then VRFlipperRight.x = VRFlipperRight.x +5
if keycode = LeftFlipperKey and VRRoom > 0 then VRFlipperLeft.x = VRFlipperLeft.x -5

  If keycode = LeftMagnaSave Then bLutActive = False


If KeyCode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall()
      If VRRoom > 0 then
      TimerPlunger.Enabled = False
      TimerPlunger2.Enabled = True
      VR_Primary_plunger.Y = -151.488
      End if
End If


  If hsbModeActive Then
    InstantInfoTimer.Enabled = False
    bInstantInfo = False
    Exit Sub
  End If

  ' Table specific

  If bGameInPLay AND NOT Tilted Then
    If keycode = LeftFlipperKey Then
      FlipperDeActivate LeftFlipper, LFPress
      SolLFlipper 0
      InstantInfoTimer.Enabled = False
      If bInstantInfo Then
        bInstantInfo = False
'       DMDScoreNow
      End If
    End If
    If keycode = RightFlipperKey Then
      FlipperDeActivate RightFlipper, RFPress
      SolRFlipper 0
      InstantInfoTimer.Enabled = False
      If bInstantInfo Then
        bInstantInfo = False
'       DMDScoreNow
      End If
    End If
  End If

  If  Tilted Then
    If keycode = LeftFlipperKey Then
      FlipperDeActivate LeftFlipper, LFPress
      SolLFlipper 0
    End If
    If keycode = RightFlipperKey Then
      FlipperDeActivate RightFlipper, RFPress
      SolRFlipper 0
    End If
  End If

End Sub

Sub InstantInfo

  DMD CL(0, "INSTANT INFO"), "", "", eNone, eNone, eNone, 800, False, ""

End Sub

Sub EndFlipperStatus
  If bInstantInfo Then
    bInstantInfo = False
'   DMDScoreNow
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
  If UseFlexDMD Then FlexDMD.Run = False
  If B2SOn = true Then Controller.Stop
End Sub


'//////////////////////////////////////////////////////////////////////
'// FLIPPERS
'//////////////////////////////////////////////////////////////////////

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
    x.TimeDelay = 10
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


'************************************************************************
'   TILT
'************************************************************************


'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

'Sub CheckTilt                                    'Called when table is nudged
' If NOT bGameInPlay Then Exit Sub
' Tilt = Tilt + TiltSensitivity                'Add to tilt count
' TiltDecreaseTimer.Enabled = True
' If(Tilt > TiltSensitivity)AND(Tilt < 15)Then 'show a warning
'   DMD "_", CL(1, "CAREFUL!"), "", eNone, eBlinkFast, eNone, 500, True, ""
' End if
' If Tilt > 15 Then 'If more that 15 then TILT the table
'   Tilted = True
'   'display Tilt
'   DMDFlush
'   DMD "", CL(1, "TILT"), "", eNone, eBlink, eNone, 100, False, ""
'   DMD CL(0, "TILT"), "", "", eNone, eBlink, eNone, 100, False, ""
'   DisableTable True
'   TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
' End If
'End Sub

Sub CheckTilt                                    'Called when table is nudged
  If NOT bGameInPlay Then Exit Sub
  Tilt = Tilt + TiltSensitivity                'Add to tilt count
  TiltDecreaseTimer.Enabled = True
  If(Tilt > TiltSensitivity) AND (Tilt <= 15) Then ShowTiltWarning  'show a warning
  If Tilt > 15 Then TiltMachine                'If more that 15 then TILT the table
End Sub

Sub CheckMechTilt                                 'Called when mechanical tilt bob switch closed
  If Not bGameInPlay Then Exit Sub
  If Not bMechTiltJustHit Then
    MechTilt = MechTilt + 1                   'Add to tilt count
    If(MechTilt > 0) AND (MechTilt <= 2) Then ShowTiltWarning 'show a warning
    If MechTilt > 2 Then TiltMachine        'If more than 2 then TILT the table
    bMechTiltJustHit = True
    TiltDebounceTimer.Enabled = True
  End If
End Sub

Sub ShowTiltWarning
  BG_ALL_OFF = 222
  PlaySound "CO_TiltWarning"
  DMD "", "", "DMD_Blank", eNone, eNone, eNone, 200, False, "" 'Light
  DMD "", "", "DMD_TiltWarning", eNone, eNone, eNone, 2500, False, "" 'Light
  vpmtimer.addtimer 2800, "DMDBlank'"
  vpmtimer.addtimer 3005, "DMDScoreNow'"
End Sub

Sub TiltMachine
  Tilted = True
  'display Tilt
  PlaySound "CO_Tilt"
  DMD "", "", "DMD_Blank", eNone, eNone, eNone, 200, False, "" 'Light
  DMD "", "", "DMD_Tilt", eNone, eNone, eNone, 2500, False, "" 'Light
  DisableTable True
  TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
End Sub

Sub TiltDecreaseTimer_Timer
  ' DecreaseTilt
  If Tilt > 0 Then
    Tilt = Tilt - 0.1
  Else
    TiltDecreaseTimer.Enabled = False
  End If
End Sub

Sub TiltDebounceTimer_Timer
  bMechTiltJustHit = False
  TiltDebounceTimer.Enabled = False
End Sub


Sub DisableTable(Enabled)
  If Enabled Then
    'turn off GI and turn off all the lights
    GiOff
    PlaySoundAt "Relay_GI_Off" , GISound
    LightSeqTilt.Play SeqAllOff
    'Disable slings, bumpers etc
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
    Bumper1.Force = 0
    Bumper2.Force = 0
    Bumper3.Force = 0
    Bumper4.Force = 0
    LeftSlingshot.Disabled = 1
    RightSlingshot.Disabled = 1
    bFlippersEnabled = False
  Else
    PlaySoundAt "Relay_GI_On" , GISound
    'turn back on GI and the lights
    GiOn
    LightSeqTilt.StopPlay
    Bumper1.Force = 7
    Bumper2.Force = 7
    Bumper3.Force = 7
    Bumper4.Force = 7
    LeftSlingshot.Disabled = 0
    RightSlingshot.Disabled = 0
    bFlippersEnabled = True
    'clean up the buffer display
    DMDFlush
  End If
End Sub

Sub TiltRecoveryTimer_Timer()
  ' if all the balls have been drained then..
  If(BallsOnPlayfield = 0)Then
    ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
    EndOfBall()
    TiltRecoveryTimer.Enabled = False
  End If
  ' else retry (checks again in another second or so)
End Sub


Dim Song
Song = ""

Sub PlaySong(name)
  If bMusicOn Then
    If Song <> name Then
      StopSound Song
      Song = name
      If Song = "m_end" Then
        PlaySound Song, 0, SongVolume  'this last number is the volume, from 0 to 1
      Else
        PlaySound Song, -1, SongVolume 'this last number is the volume, from 0 to 1
      End If
    End If
  End If
End Sub
Dim NewSong
Sub ChangeSong
  If(BallsOnPlayfield = 0) Then
    PlaySong "m_Slitheen"
  Else

    NewSong=NewSong +1
    Select Case NewSong
      Case 0 PlaySong "m_The Carrionites Swarm"
      Case 1 PlaySong "m_Westminster Bridge"
      Case 2 PlaySong "m_Slitheen"
      Case 3 PlaySong "m_Westminster Bridge"
      Case 4 PlaySong "m_The Carrionites Swarm"
      Case 5 PlaySong "m_Westminster Bridge"
      Case 6 PlaySong "m_Slitheen"
      Case 7 PlaySong "m_Westminster Bridge":NewSong=0

    End Select
  End If
End Sub

'================================
'Helper Functions

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

'*******************************************************
'   START GAME, END GAME- User Defined Script Events
'********************************************************

Sub ResetForNewGame()
  robot.visible = False : qman.visible = False : TimeThief.Visible = true
  Dim i
  bGameInPLay = True
  'resets the score display, and turn off attract mode
  StopAttractMode
  GiOn
  TotalGamesPlayed = TotalGamesPlayed + 1
  CurrentPlayer = 1
  PlayersPlayingGame = 1
  bOnTheFirstBall = True
debug.print "PlayersPlayingGame" & PlayersPlayingGame
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
  PlaySoundAt "ball_trough", lane4
  ' you may wish to start some music, play a sound, do whatever at this point
    vpmtimer.addtimer 1500, "FirstBall '"
End Sub





'***********FIRSTBALL
'FirstBall
'*************************************************

' This is used to delay the start of a game to allow any attract sequence to
' complete.  When it expires it creates a ball for the player to start playing with
Dim FirstBallStarted ' For DMD Tracking
Sub FirstBall

  FirstBallStarted=True
  ' create a new ball in the shooters lane
  vpmtimer.addtimer 1500, " CreateNewBall'"

End Sub

' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))

Sub ResetForNewPlayerBall()

  ' set the current players bonus multiplier back down to 1X
  '    SetBonusMultiplier 1
  ' reset any drop targets, lights, game modes etc..


  If bExtraBallWonThisBall=True Then ResetNewPlayerBallCommon:vpmtimer.addtimer 500, "ResetNewBallVariables '":debug.print "ResetNPBExtraBall":Exit Sub

  If  PlayerFirstBallPlayed(CurrentPlayer)=0 Then
      ResetStartofGameVariables
      TurnOnStartOfGameLights
    debug.print "NewPlayerResetStartofGameVariables"

  End If

  If GloryTableStarted(CurrentPlayer)=True And bExtraBallWonThisBall=False Then
    ResetStartofGameVariables
'   StopPlanets
    TurnOnStartOfGameLights
    GloryTableStarted(CurrentPlayer)=False

  End If

  If  PlayerFirstBallPlayed(CurrentPlayer)=1 And GloryTableStarted(CurrentPlayer) =False Then
  vpmtimer.addtimer 500, "ResetNewBallVariables '"
    debug.print "ResetNewBallVariablesCalledinResetNewPlayerBall"
  End If
' bExtraBallWonThisBall = False  Note Deactivated in ResetNew ball variables

  ResetNewPlayerBallCommon

End Sub

Sub ResetNewPlayerBallCommon
  'Reset any table specific
  'This is a new ball, so activate the ballsaver
  bBallSaverReady = True
  'and the skillshot
  '    bSkillShotPlayedOnce = False
  StartResetNewBallFlashers
  bSkillShotReady = True
  MeteorShowerInPlay=False
  StopColouredBalls
End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()

  debug.print "CreateNewBall"
' PlaySoundAt "ball_trough", lane4
  vpmtimer.addtimer 100, "DMDScoreNow'"
  AddScore 0
  ' create a ball in the plunger lane kicker.
  BallRelease.CreateSizedball BallSize / 2
  dof 110 ,DOFPulse

  ' There is a (or another) ball on the playfield
  BallsOnPlayfield = BallsOnPlayfield + 1
  ' kick it out..
  RandomSoundBallRelease BallRelease
  BallRelease.Kick 90, 4

  ' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
  ' set the bAutoPlunger flag to kick the ball in play automatically
  If BallsOnPlayfield > 1 Then
    DOF 129, DOFPulse
    bMultiBallMode = True
    bAutoPlunger = True
    '        PlaySong "m_multiball"
  End If
End Sub

Sub BallReleaseSound
' RandomSoundBallRelease BallRelease
End Sub


' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
  If MeteorShowerInPlay=True Or TWJPFiring=True Or GloryTableOn(CurrentPlayer)=True Then StartColouredBalls: End If
' If MeteorShowerInPlay=False Or TWJPFiring=False Then StopColouredBalls: End If

  bAutoPlunger = True
  Wall_OrbitDropper.IsDropped=True

  debug.print "addmultiball"
  mBalls2Eject = mBalls2Eject + nballs
  CreateMultiballTimer.Enabled = True
  'and eject the first ball
' vpmtimer.addtimer 2000, "DMDScoreNow'"

  vpmtimer.addtimer 1200, "CreateMultiballTimer_Timer '"
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
  FirstBallStarted=False

  debug.print "EndOfBall"
  '
  ' the first ball has been lost. From this point on no new players can join in
  bOnTheFirstBall = False

  ' only process any of this if the table is not tilted.  (the tilt recovery
  ' mechanism will handle any extra balls or end of game)
' vpmtimer.addtimer 200, "DMDScoreNow'"
  If NOT Tilted Then
    vpmtimer.addtimer 200, "EndOfBall2 '"
    Else
    vpmtimer.addtimer 100, "EndOfBall2 '" 'If tilted add short delay and move to the second part of end of the ball
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
  Tilt = 0
  MechTilt = 0
  DisableTable False 'enable again bumpers and slingshots
'If Glory Table is on at ball drain then it must be turned off . The GloryStopatendofballcomplete subroutine. turns off th eglory table and sets up for new player ball
  If GloryTableOn(CurrentPlayer)=True And GloryTableStarted(CurrentPlayer)=True Then GloryStopAtEndOfBallComplete:debug.print "GloryStop-AtEndOfBall":End If
'

  ' has the player won an extra-ball ? (might be multiple outstanding)
  If(ExtraBallsAwards(CurrentPlayer) <> 0) Then
    debug.print "Extra Ball"
    PlaySound "fx_fanfare8"
    ' yep got to give it to them
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1


    ' if no more EB's then turn off any shoot again light
    If(ExtraBallsAwards(CurrentPlayer) = 0) Then
      LightShootAgain.State = 0
      EBActive=False
      Debug.Print "EBActive=False"

    End If

    ' You may wish to do a bit of a song AND dance at this point
    DMD CL(0,"EXTRA BALL"), CL(1,"SHOOT AGAIN"), "", eNone, eNone, eBlink, 3000, True, ""

    ' In this table an extra ball will have the skillshot and ball saver, so we reset the playfield for the new ball

    vpmtimer.addtimer 3000, "ResetNewBallVariables:ChangeSong()'"
    vpmtimer.addtimer 3200, "NewBallAlert1'"
' Create a new ball in the shooters lane and change of song so attracts players attention
    vpmtimer.addtimer 3500, "CreateNewBall'"
debug.print "CreateNewBall-ExtraBall"

'just reload the planets and previousLightStates not Modesand Awards



    Else ' no extra balls

      BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

      PlayerFirstBallPlayed(CurrentPlayer)=1
debug.print "PlayerFirstBallPlayed(CurrentPlayer)=True" & CurrentPlayer
debug.print "CurrentPlayer & BallsRemaining=" & CurrentPlayer & BallsRemaining(CurrentPlayer)
    ' was that the last ball ?
    If(BallsRemaining(CurrentPlayer) <= 0) Then
debug.print "No More Balls, High Score Entry"

      ' Submit the CurrentPlayers score to the High Score system
       '10second delay to allow Glory table to stop if ball drains before completing
      ' you may wish to play some music at this point

      'EndOfBallCompleteis called in theHighScoreEntryCheck
      CheckHighScore()
    Else

      ' not the last ball (for that player)
      ' if multiple players are playing then move onto the next one
     vpmtimer.addtimer 3000, "EndOfBallComplete()'"

      debug.print "EndOfBallComplete"

    End If
  End If
End Sub

Sub NewBallAlert1
debug.print "ExtraBallNewBallAlert"
  PlaySound "CO_ReturnToSender"
  PlaySound "TriggerThats it"

End Sub


' This function is called when the end of bonus display
' (or high score entry finished) AND it either end the game or
' move onto the next player (or the next ball of the same player)
'
Dim TimeWarpEndOfBallComplete

Sub EndOfBallComplete()
  ChangeSong()

debug.print "PlayerFirstBallPlayed"
  DragonTargetFlasherBlinkTimer.Enabled=True


  ' are there multiple players playing this game ?
  If(PlayersPlayingGame> 1) Then
    ' then move to the next player
    NextPlayer = CurrentPlayer + 1


debug.print "PlayersPlayingGame&NextPlayer & BallsRemaining=" & PlayersPlayingGame & CurrentPlayer & BallsRemaining(CurrentPlayer)

    ' are we going from the last player back to the first
    ' (ie say from player 4 back to player 1)
    If(NextPlayer> PlayersPlayingGame) Then
      NextPlayer = 1
  debug.print "PlayersPlayingGame&Next Player = " & PlayersPlayingGame & CurrentPlayer & BallsRemaining(CurrentPlayer)
    End If
  Else
    NextPlayer = CurrentPlayer

  End If

  debug.print "Next Player = " & NextPlayer & BallsRemaining(1) & BallsRemaining(2) & BallsRemaining(3)& BallsRemaining(4)

  ' is it the end of the game ? (all balls been lost for all players)
  If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
    ' you may wish to do some sort of Point Match free game award here
    ' generally only done when not in free play mode
  debug.print "EndofGame Called Balls Remaining CurrentPlayer=" & BallsRemaining(CurrentPlayer)
    ' set the machine into game over mode
        vpmtimer.addtimer 3000, "EndOfGame()'"


    ' you may wish to put a Game Over message on the desktop/backglass
  Else

    ' set the next player
    CurrentPlayer = NextPlayer
debug.print "CurrentPlayer" & CurrentPlayer & BallsRemaining(CurrentPlayer)
    ' make sure the correct display is up to date
    AddScore 0

debug.print "WhyIsThisCalledAfterEndofGame" & BallsRemaining(CurrentPlayer)
    ' reset the playfield for the new player (or new ball)
    vpmtimer.addtimer 100, "ResetForNewPlayerBall()'"
debug.print "ResetForNewPlayerBall-SubEndofBallComplete"

 'And create a new ball
    vpmtimer.addtimer 200, "CreateNewBall():NewBallAlert2'"
debug.print "CreateNewBall-EndOfBallComplete"
    ' play a sound if more than 1 player

    If PlayersPlayingGame> 1 Then
      '         PlaySound "vo_player" &CurrentPlayer
      DMD "_", CL(1, "PLAYER " &CurrentPlayer), "_", eNone, eNone, eNone, 800, True, ""
    End If
  End If
End Sub


Sub NewBallAlert2
  PlaySound "TriggerThats it"
End Sub




' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
  debug.print "End Of Game"
  bGameInPLay = False
  ' just ended your game then play the end of game tune
  If NOT bJustStarted Then
    PlaySong "m_Slitheen"
  End If

  bJustStarted = False
  '   ' ensure that the flippers are down
  SolLFlipper 0
  SolRFlipper 0

  ' terminate all Mode - eject locked balls
  ' most of the Mode/timers terminate at the end of the ball



  '    PlayQuote.Enabled = 0  refer to ghost buster slimer. requires timer and Sub



  ' set any lights for the attract mode
  GiOff
  VpmTimer.addtimer 8000, "EndOfGameCallout'"
  StartAttractMode
  ' you may wish to light any Game Over Light you may have
End Sub

Sub EndOfGameCallout
  PlaySound "CO_VintageSumpOil"
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



Sub   TurnOffStartOfGameLights

End Sub


' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' lost a ball ;-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one

'Sub TriggerDrainControl1_Hit
' GateDrainControl1.open=0
' vpmtimer.addtimer 500, "OpenDrainGate1'"
' GateDrainControl2.timerenabled=1
'End Sub

'Sub OpenDrainGate1
''  GateDrainControl1.open=1

'End Sub
Sub KickerDrainControl_Hit
  PlaySoundAt"fx_kicker_enter",KickerDrainControl
  vpmtimer.addtimer 500, "DrainControlKick'"
End Sub

Sub DrainControlKick
  KickerDrainControl.Kick 95, 4
  PlaySoundAt"fx_kicker",KickerDrainControl
End Sub



Sub Drain_Hit()
debug.print "Drain Hit"
  ' Destroy the ball
  Drain.DestroyBall
  If Not bGameInPlay Then Exit Sub
  ' Exit Sub ' only for debugging - this way you can add balls from the debug window

  BallsOnPlayfield = BallsOnPlayfield - 1

  ' pretend to knock the ball into the ball storage mech

  RandomSoundDrain Drain
  'if Tilted the end Ball Mode
  If Tilted Then
    StopEndOfBallMode
  End If

  ' if there is a game in progress AND it is not Tilted
  If(bGameInPLay = True) AND(Tilted = False) Then

    ' is the ball saver active,
    If(bBallSaverActive = True) Then

      ' yep, create a new ball in the shooters lane
      ' we use the Addmultiball in case the multiballs are being ejected
      AddMultiball 1
      ' we kick the ball with the autoplunger
      bAutoPlunger = True
      ' you may wish to put something on a display or play a sound at this point
      DMD "_", CL(1, "BALL SAVED"), "_", eNone, eNone, eNone, 800, True, ""
    Else
      ' cancel any multiball if on last ball (ie. lost all other balls)
      If(BallsOnPlayfield = 1) Then
        ' AND in a multi-ball??
        If(bMultiBallMode = True) then
          ' not in multiball mode any more
          bMultiBallMode = False
          StopColouredBalls

          ' you may wish to change any music over at this point and
          '                   ChangeSong
        End If
      End If

      ' was that the last ball on the playfield
      If(BallsOnPlayfield = 0) Then
        BG_ALL_OFF = 260 ' just to have blank dmd screens
        StopEndOfBallMode
        vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, since there is no animation then move to the end of ball
      End If
    End If
  End If
End Sub


' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.


Sub ballsavestarttrigger_hit
  ' if there is a need for a ball saver, then start off a timer
  ' only start if it is ready, and it is currently not running, else it will reset the time period
  If(bBallSaverReady = True) AND(25 <> 0) And(bBallSaverActive = False) Then
    If Hard=1 Then EnableBallSaver 20
    If Hard=0 Then EnableBallSaver 25
  End If
  RailFlasherTimer.Enabled=1
  PlaySound "liftoff2_3sec"
  LastSwitchHit= "ballsavestarttrigger"
End Sub

Dim RailFlasher:RailFlasher=0
Sub RailFlasherTimer_Timer
  RailFlasher=RailFlasher+1
  Select Case RailFlasher
    Case 1: RailFlasherOn1.Visible=1
    Case 2: RailFlasherOn2.Visible=1
    Case 3: RailFlasherOn3.Visible=1
    Case 4: RailFlasherOn4.Visible=1
    Case 5: RailFlasherOn5.Visible=1
    Case 6: RailFlasherOn6.Visible=1
    Case 7: RailFlasherOn7.Visible=1
    Case 8: RailFlasherOn8.Visible=1
    Case 9: RailFlasherOn9.Visible=1
    Case 10: RailFlasherOn10.Visible=1
    Case 11: RailFlasherOn11.Visible=1
    Case 12: RailFlasherOn12.Visible=1
    Case 13: RailFlasherOn13.Visible=1
    Case 14: RailFlasherOn14.Visible=1
    Case 15: RailFlasherOn15.Visible=1
    Case 16: RailFlasherOn16.Visible=1
    Case 17: RailFlasherOn17.Visible=1
    Case 18: RailFlasherOn18.Visible=1
    Case 19: RailFlasherOn19.Visible=1
    Case 20: RailFlasherOn20.Visible=1
    Case 21: RailFlasherOn21.Visible=1
    Case 22: RailFlasherOn22.Visible=1
    Case 23: RailFlasherOn23.Visible=1
    Case 24: RailFlasherOn24.Visible=1
    Case 25: RailFlasherOn25.Visible=1
    Case 26: RailFlasherOn26.Visible=1
    Case 27: RailFlasherOn27.Visible=1
    Case 28: RailFlasherOn28.Visible=1
    Case 29: RailFlasherOn29.Visible=1
    Case 30: RailFlasherOn30.Visible=1
    Case 31: RailFlasherOn31.Visible=1
    Case 32: RailFlasherOn32.Visible=1
    Case 33: RailFlasherOn33.Visible=1
    Case 34: RailFlasherOn34.Visible=1
    Case 35: RailFlasherOn35.Visible=1
    Case 36: RailFlasherOn36.Visible=1
    Case 37: RailFlashersOff:RailFlasher=0:RailFlasherTimer.Enabled=0
'   Case 38: RailFlasherOn1.Visible=1
'   Case 39: RailFlasherOn2.Visible=1
'   Case 40: RailFlasherOn3.Visible=1
'   Case 41: RailFlasherOn4.Visible=1
'   Case 42: RailFlasherOn5.Visible=1
'   Case 43: RailFlasherOn6.Visible=1
'   Case 44: RailFlasherOn7.Visible=1
'   Case 45: RailFlasherOn8.Visible=1
'   Case 46: RailFlasherOn9.Visible=1
'   Case 47: RailFlasherOn10.Visible=1
'   Case 48: RailFlasherOn11.Visible=1
'   Case 49: RailFlasherOn12.Visible=1
'   Case 50: RailFlasherOn13.Visible=1
'   Case 51: RailFlasherOn14.Visible=1
'   Case 52: RailFlasherOn15.Visible=1
'   Case 53: RailFlasherOn16.Visible=1
'   Case 54: RailFlasherOn17.Visible=1
'   Case 55: RailFlasherOn18.Visible=1
'   Case 56: RailFlasherOn19.Visible=1
'   Case 57: RailFlasherOn20.Visible=1
'   Case 58: RailFlasherOn21.Visible=1
'   Case 59: RailFlasherOn22.Visible=1
'   Case 60: RailFlasherOn23.Visible=1
'   Case 61: RailFlasherOn24.Visible=1
'   Case 62: RailFlasherOn25.Visible=1
'   Case 63: RailFlasherOn26.Visible=1
'   Case 64: RailFlasherOn27.Visible=1
'   Case 65: RailFlasherOn28.Visible=1
'   Case 66: RailFlasherOn29.Visible=1
'   Case 67: RailFlasherOn30.Visible=1
'   Case 68: RailFlasherOn31.Visible=1
'   Case 69: RailFlasherOn32.Visible=1
'   Case 70: RailFlasherOn33.Visible=1
'   Case 71: RailFlasherOn34.Visible=1
'   Case 72: RailFlasherOn35.Visible=1
'   Case 73: RailFlasherOn36.Visible=1
'   Case 74: RailFlashersOff:RailFlasher=0:RailFlasherTimer.Enabled=0
  End Select
End Sub

Sub RailFlashersOff
' RailFlasherTimer.Enabled=0
  RailFlasherOn1.Visible=0:RailFlasherOn2.Visible=0:RailFlasherOn3.Visible=0:RailFlasherOn4.Visible=0:RailFlasherOn5.Visible=0:RailFlasherOn6.Visible=0
  RailFlasherOn7.Visible=0:RailFlasherOn8.Visible=0:RailFlasherOn9.Visible=0:RailFlasherOn10.Visible=0:RailFlasherOn11.Visible=0:RailFlasherOn12.Visible=0
  RailFlasherOn13.Visible=0:RailFlasherOn14.Visible=0:RailFlasherOn15.Visible=0:RailFlasherOn16.Visible=0:RailFlasherOn17.Visible=0
  RailFlasherOn18.Visible=0:RailFlasherOn19.Visible=0:RailFlasherOn20.Visible=0:RailFlasherOn21.Visible=0:RailFlasherOn22.Visible=0
  RailFlasherOn23.Visible=0:RailFlasherOn24.Visible=0:RailFlasherOn25.Visible=0:RailFlasherOn26.Visible=0:RailFlasherOn27.Visible=0
  RailFlasherOn28.Visible=0:RailFlasherOn29.Visible=0:RailFlasherOn30.Visible=0:RailFlasherOn31.Visible=0:RailFlasherOn32.Visible=0
  RailFlasherOn33.Visible=0:RailFlasherOn34.Visible=0:RailFlasherOn35.Visible=0:RailFlasherOn36.Visible=0
End Sub


Sub swPlungerRest_Hit()
  debug.print "ball in plunger lane"
  bBallInPlungerLane = True
  bSkillshotReady = True

  ' turn on Launch light is there is one
  'LaunchLight.State = 2
  ' kick the ball in play if the bAutoPlunger flag is on
If bMultiBallMode=True Then:bAutoPlunger=True
  If bAutoPlunger=True Then
    vpmtimer.addtimer 2500, "AutoPlungerDelay '"

  End If

  LastSwitchHit = "swPlungerRest"
End Sub

Sub swPlungerRest_UnHit()
  bBallInPlungerLane = False
  swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
  If bSkillShotReady Then
    '     ResetSkillShotTimer.Enabled = 1
  End If
End Sub

Sub AutoPlungerDelay
  PlungerIM.Strength = 0.2
  'PlungerIM.AutoFire
  PlungerIM.Strength = Plunger.MechStrength
  Plunger.AutoPlunger = True
  Plunger.Pullback
  Plunger.Fire
  PlaySoundAt "Popper" , Plunger
  DOF 125, DOFPulse
  DOF 112 ,DOFPulse
  bAutoPlunger = False
  Plunger.AutoPlunger = False
End Sub


Sub swPlungerRest_Timer
  swPlungerRest.TimerEnabled = 0
End Sub



Sub EnableBallSaver(seconds)
  debug.print "Ballsaver started"
  ' set our game flag
  bBallSaverActive = True
  bBallSaverReady = False
  ' start the timer
  BallSaverTimerExpired.Interval = 1000 * seconds
  BallSaverTimerExpired.Enabled = True
  BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
  BallSaverSpeedUpTimer.Enabled = True
  ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
  Light_BallSaver.BlinkInterval = 160
  Light_BallSaver.State = 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag

Sub BallSaverTimerExpired_Timer()
  debug.print "Ballsaver ended"
' BallSaverTimerExpired.Enabled = False
  ' clear the flag
' bBallSaverActive = False
  ' if you have a ball saver light then turn it off at this point
  '    Light_BallSaver.BlinkInterval = 80

  Light_BallSaver.State = 0
  vpmtimer.addtimer 2000, "GraceTime '"
End Sub


Sub BallSaverSpeedUpTimer_Timer()
  'debug.print "Ballsaver Speed Up Light"
  BallSaverSpeedUpTimer.Enabled = False
  ' Speed up the blinking
  Light_BallSaver.BlinkInterval = 80
  Light_BallSaver.State = 2
End Sub

Sub GraceTime
  bBallSaverActive = False
  BallSaverTimerExpired.Enabled = False
End Sub




' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board
' In this table we use SecondRound variable to double the score points in the second round after killing Malthael
Sub AddScore(points)
  If(Tilted = False) Then
    ' add the points to the current players score variable
    Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer)
  End if
  ' you may wish to check to see if the player has gotten a replay
End Sub

' Add bonus to the bonuspoints AND update the score board

Sub AddBonus(points) 'not used in this table, since there are many different bonus items.
  If(Tilted = False) Then
    ' add the bonus to the current players bonus variable
    BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
  End if
End Sub

' Add some points to the current Jackpot.
'
Sub AddJackpot(points)
  ' Jackpots only generally increment in multiball mode AND not tilted
  ' but this doesn't have to be the case
  '    If(Tilted = False) Then

  ' If(bMultiBallMode = True) Then
  '       Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + points
  '       DMD "_", CL(1, "INCREASED JACKPOT"), "_", eNone, eNone, eNone, 800, True, ""
  ' you may wish to limit the jackpot to a upper limit, ie..
  ' If (Jackpot >= 6000) Then
  '   Jackpot = 6000
  '   End if
  'End if
  '   End if
End Sub

Sub AddSuperJackpot(points) 'not used in this table
  '    If(Tilted = False) Then
  '   End if
End Sub

Sub AddBonusMultiplier(n)
  '    Dim NewBonusLevel
  ' if not at the maximum bonus level
  '    if(BonusMultiplier(CurrentPlayer) + n <= MaxMultiplier) then
  ' then add and set the lights
  NewBonusLevel = BonusMultiplier(CurrentPlayer) + n
  '       SetBonusMultiplier(NewBonusLevel)
  '       DMD "_", CL(1, "BONUS X " &NewBonusLevel), "_", eNone, eNone, eNone, 2000, True, "fx_bonus"
  '    Else
  '        AddScore 50000
  '        DMD "_", CL(1, "50000"), "_", eNone, eNone, eNone, 800, True, ""
  '   End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly

Sub SetBonusMultiplier(Level)
  '   ' Set the multiplier to the specified level
  '   BonusMultiplier(CurrentPlayer) = Level
  '   UPdateBonusXLights(Level)
End Sub

Sub UpdateBonusXLights(Level)
  ' Update the lights
  '    Select Case Level
  '       Case 1:light54.State = 0:light55.State = 0:light56.State = 0:light57.State = 0
  '       Case 2:light54.State = 1:light55.State = 0:light56.State = 0:light57.State = 0
  '      Case 3:light54.State = 0:light55.State = 1:light56.State = 0:light57.State = 0
  '      Case 4:light54.State = 0:light55.State = 0:light56.State = 1:light57.State = 0
  ''      Case 5:light54.State = 0:light55.State = 0:light56.State = 0:light57.State = 1
  '   End Select
End Sub

Sub AddPlayfieldMultiplier(n)
  '    Dim NewPFLevel
  ' if not at the maximum level x
  '   if(PlayfieldMultiplier(CurrentPlayer) + n <= MaxMultiplier) then
  ' then add and set the lights
  '        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) + n
  '        SetPlayfieldMultiplier(NewPFLevel)
  '        DMD "_", CL(1, "PLAYFIELD X " &NewPFLevel), "_", eNone, eNone, eNone, 2000, True, "fx_bonus"
  '    Else 'if the 5x is already lit
  '        AddScore 50000
  '        DMD "_", CL(1, "50000"), "_", eNone, eNone, eNone, 2000, True, ""
  '    End if
  'Start the timer to reduce the playfield x every 30 seconds
  '   pfxtimer.Enabled = 0
  '    pfxtimer.Enabled = 1
  'End Sub

  ' Set the Playfield Multiplier to the specified level AND set any lights accordingly

  'Sub SetPlayfieldMultiplier(Level)
  ' Set the multiplier to the specified level
  '   PlayfieldMultiplier(CurrentPlayer) = Level
  '   UpdatePFXLights(Level)
  'End Sub

  'Sub UpdatePFXLights(Level)
  ' Update the lights
  '   Select Case Level
  '       Case 1:light3.State = 0:light2.State = 0:light1.State = 0:light4.State = 0
  '       Case 2:light3.State = 1:light2.State = 0:light1.State = 0:light4.State = 0
  '       Case 3:light3.State = 0:light2.State = 1:light1.State = 0:light4.State = 0
  '       Case 4:light3.State = 0:light2.State = 0:light1.State = 1:light4.State = 0
  '       Case 5:light3.State = 0:light2.State = 0:light1.State = 0:light4.State = 1
  '   End Select
  ' show the multiplier in the DMD
  ' in this table the multiplier is always shown in the score display sub
End Sub

Dim EBActive
Sub AwardExtraBall()
  BG_Sequence 5,10,0
  ExtraBallActivated=0
  If bExtraBallWonThisBall Then AwardSuperJackpot:End If
  If NOT bExtraBallWonThisBall Then
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  LS_MajorAward.Play SeqBlinking,, 8,100
    DOF 122, DOFPulse
    PlaySoundAt "fx_knocker",Kicker8
    Playsound "CO_ExtraBall"
    DMD "", "", "DMD_ExtraBall1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_ExtraBall2a", eNone, eNone, eNone, 250, True, ""

    vpmtimer.addtimer 4500, "ExtraBallScore'"

    EBActive=1
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    bExtraBallWonThisBall = True
    LightShootAgain.State = 1 'light the shoot again lamp
    GiOff
    vpmTimer.AddTimer 1200, "GiOn'"
    LS_MajorAward.Play SeqBlinking,, 8,85
    StopExtraBall
    Debug.print "EBActive=True"

  End If
End Sub

Sub ExtraBallScore
  DMDScoreNow
  AddScore 500000
End Sub

Sub AwardSpecial()
  DMD "_", CL(1, ("REPLAY")), "_", eNone, eBlink, eNone, 2200, True,""
  DOF 122, DOFPulse
  vpmTimer.AddTimer 2000, "DMDScoreNow'"
  Credits = Credits + 1
  If bFreePlay = False Then DOF 125, DOFOn
  LS_MajorAward.Play SeqBlinking,, 4,100
  '    FlashEffect 2
End Sub

Sub AwardJackpot 'award a normal jackpot,
  BG_Sequence 2,0,200
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  LS_MajorAward.Play SeqBlinking,, 8,100
  PlaySoundAt "fx_knocker",Kicker8
  DOF 122, DOFPulse
  DMD "", "", "DMD_Jackpot2a", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_Jackpot2b", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_Jackpot2a", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_Jackpot2b", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_Jackpot2c", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 4500, "JackPotScore'"
  PlaySound "CO_Jackpot"
  LS_MajorAward.Play SeqBlinking,, 4,100
  '   FlashEffect 2
End Sub

Sub JackPotScore
  DMDScoreNow
  AddScore 100000000

End Sub

Sub AwardJackpot2 'award a normal jackpot,
  BG_Sequence 2,0,200

  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  LS_MajorAward.Play SeqBlinking,, 8,100
  PlaySoundAt "fx_knocker",Kicker8
  DOF 122, DOFPulse
  DMD "", "", "DMD_Jackpot2a", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_Jackpot2b", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_Jackpot2a", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_Jackpot2b", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_Jackpot2c", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 4500, "JackPot2Score'"
  PlaySound "CO_Jackpot"
End Sub

Sub JackPot2Score
  DMDScoreNow
  AddScore 200000000
End Sub

Sub AwardSuperJackpot 'this is actually xxx
  BG_Sequence 5,10,0
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  LS_MajorAward.Play SeqBlinking,, 8,100
  PlaySoundAt "fx_knocker",Kicker8
  DOF 122, DOFPulse
  PlaySound "CO_SuperJackpot"
  DMD "", "", "DMD_SuperJackpot2b", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2a", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2b", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2a", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2b", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2a", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2b", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 5500, "SuperJackPotScore'"
  PlaySound "CO_SuperJackpot"
  LS_MajorAward.UpdateInterval = 10
  LS_MajorAward.Play SeqLeftOn, 35, 1
  LS_MajorAward.UpdateInterval = 10
  LS_MajorAward.Play SeqRightOn, 50, 1
  LS_MajorAward.UpdateInterval = 10
  LS_MajorAward.Play SeqLeftOn, 35, 1
  LS_MajorAward.UpdateInterval = 10
  LS_MajorAward.Play SeqRightOn, 50, 1
  LS_MajorAward.Play SeqUpOn, 50, 1
  '   LightEffect 2
  '   FlashEffect 2
End Sub

Sub SuperJackPotScore
  DMDScoreNow
  AddScore 400000000

End Sub

Sub AwardSuperJackpot2 'this is actually xxx
  BG_Sequence 5,10,0

  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  LS_MajorAward.Play SeqBlinking,, 8,100
  PlaySoundAt "fx_knocker",Kicker8
  PlaySound "CO_SuperJackpot"
  DOF 122, DOFPulse
  DMD "", "", "DMD_SuperJackpot2c", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2d", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2c", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2d", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2c", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2d", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_SuperJackpot2c", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 4500, "SuperJackPot2Score'"
  PlaySound "CO_SuperJackpot"
  LS_MajorAward.Play SeqBlinking,, 6,100
End Sub

Sub SuperJackPot2Score
  BG_Sequence 5,10,0
  DMDScoreNow
  AddScore 350000000

End Sub

Sub AwardSkillshot() 'Notused in this game awards at the kickers
  BG_Sequence 5,10,0
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  LS_MajorAward.Play SeqBlinking,, 8,100
  If DragonWounded=False Then DropDragonTargets: DragonWallDropActivated=True: End If
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_SkillShot":End If

  '   ResetSkillShotTimer_Timer
  'show dmd animation
  DMD "", "", "DMD_AttackTheDragon1a", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_AttackTheDragon1b", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_AttackTheDragon1a", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_AttackTheDragon1b", eNone, eNone, eNone, 250, True, ""
  DMD "", "", "DMD_AttackTheDragon1a", eNone, eNone, eNone, 2500, True, ""

  vpmtimer.addtimer 3800, "AddSkillshotScore'"
'     DOF 127, DOFPulse
  ' increment the skillshot value with 250.000
  'do some light show
  '    GiEffect 2
  '    LightEffect 2
End Sub

Sub AddSkillshotScore
     Addscore SkillShotValue(CurrentPlayer)
     SkillShotValue(CurrentPlayer) = SkillShotValue(CurrentPlayer) + 50000000

End Sub




'*****************************
'    Load / Save / Highscore
'*****************************
'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
  Dim x
  x = LoadValue(TableName, "HighScore1")
  If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 15000000000 :End If
  x = LoadValue(TableName, "HighScore1Name")
  If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" :End If
  x = LoadValue(TableName, "HighScore2")
  If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 10000000000 :End If
  x = LoadValue(TableName, "HighScore2Name")
  If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "BBB" :End If
  x = LoadValue(TableName, "HighScore3")
  If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 3000000000 :End If
  x = LoadValue(TableName, "HighScore3Name")
  If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "CCC" :End If
  x = LoadValue(TableName, "HighScore4")
  If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 800000000 :End If
  x = LoadValue(TableName, "HighScore4Name")
  If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "DDD" :End If
  x = LoadValue(TableName, "Credits")
  If(x <> "") then Credits = CInt(x) Else Credits = 0:If bFreePlay = False Then DOF 125, DOFOff : End If : End If
  x = LoadValue(TableName, "TotalGamesPlayed")
  If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 :End If
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

Sub Reseths
  HighScoreName(0) = "AAA"
  HighScoreName(1) = "BBB"
  HighScoreName(2) = "CCC"
  HighScoreName(3) = "DDD"
  HighScore(0) = 15000000000
  HighScore(1) = 10000000000
  HighScore(2) = 3000000000
  HighScore(3) = 800000000
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

  If tmp > HighScore(0)Then 'add 1 credit for beating the highscore
    Credits = Credits + 1
    DOF 125, DOFOn
  End If

  If tmp > HighScore(3)Then
    PlaySound SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
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
  PlaySound "CO_YouKIlledThat"
  PlaySong "m_The Carrionites Swarm"
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
    if(hsCurrentLetter > len(hsValidLetters))then
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
      if(hsCurrentDigit > 0)then
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

Dim dLine(2)
Sub HighScoreDisplayName()
  Dim i
  Dim TempTopStr
  Dim TempBotStr

  TempTopStr = "YOUR NAME:"
  dLine(0) = ExpandLine(TempTopStr,0)
  DMDUpdate 0

  TempBotStr = "    > "
  if(hsCurrentDigit > 0)then TempBotStr = TempBotStr & hsEnteredDigits(0)
  if(hsCurrentDigit > 1)then TempBotStr = TempBotStr & hsEnteredDigits(1)
  if(hsCurrentDigit > 2)then TempBotStr = TempBotStr & hsEnteredDigits(2)

  if(hsCurrentDigit <> 3)then
    if(hsLetterFlash <> 0)then
      TempBotStr = TempBotStr & "_"
    else
      TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
    end if
  end if

  if(hsCurrentDigit < 1)then TempBotStr = TempBotStr & hsEnteredDigits(1)
  if(hsCurrentDigit < 2)then TempBotStr = TempBotStr & hsEnteredDigits(2)

  TempBotStr = TempBotStr & " <    "
  dLine(1) = ExpandLine(TempBotStr,1)
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
      If HighScore(j) < HighScore(j + 1)Then
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
          DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.dempty&dmd=2")
          Digits(i).Visible = False

          Digits(i).Visible = True

        Next
        digitgrid.Visible = False
        For i = 0 to 19 ' Top
          DMDScene.GetImage("Dig" & i).SetBounds 8 + i * 12, 6, 14, 22
        Next
        For i = 20 to 39 ' Bottom
          DMDScene.GetImage("Dig" & i).SetBounds 8 + (i - 20) * 12, 6 + 24 + 4, 14, 22
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
          DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.dempty&dmd=2")
          Digits(i).Visible = False
        Next
        digitgrid.Visible = False
        For i = 0 to 19 ' Top
          DMDScene.GetImage("Dig" & i).SetBounds 4 + i * 6, 3, 7, 11
        Next
        For i = 20 to 39 ' Bottom
          DMDScene.GetImage("Dig" & i).SetBounds 4 + (i - 20) * 6, 3 + 12 + 2, 7, 11
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
  dCharsPerLine(0) = 16 'characters lower line
  dCharsPerLine(1) = 20 'characters top line
  dCharsPerLine(2) = 1  'characters back line
  For i = 0 to 2
    dLine(i) = Space(dCharsPerLine(i) )
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
  Dim tmp, tmp1, tmp2
  if(dqHead = dqTail) Then

    tmp = RL(0, FormatScore(Score(Currentplayer) ) )
    '       tmp = CL(0, FormatScore(Score(Currentplayer) ) )
    tmp1 = CL(1, "PLAYER " & CurrentPlayer & " BALL " & Balls)
    tmp2 = ""
    '        tmp2 = "bkborder"
  End If
  DMD tmp, tmp1, tmp2, eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDScoreNow
  DMDFlush
  DMDScore
End Sub

Sub DMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
  if(dqTail <dqSize) Then
    if(Text0 = "_") Then
      dqEffect(0, dqTail) = eNone
      dqText(0, dqTail) = "_"
    Else
      dqEffect(0, dqTail) = Effect0
      dqText(0, dqTail) = ExpandLine(Text0, 0)
    End If

    if(Text1 = "_") Then
      dqEffect(1, dqTail) = eNone
      dqText(1, dqTail) = "_"
    Else
      dqEffect(1, dqTail) = Effect1
      dqText(1, dqTail) = ExpandLine(Text1, 1)
    End If

    if(Text2 = "_") Then
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
    if(dqTail = 1) Then
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
      Case eScrollLeft:deCountEnd(i) = Len(dqText(i, dqHead) )
      Case eScrollRight:deCountEnd(i) = Len(dqText(i, dqHead) )
      Case eBlink:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
        deBlinkCycle(i) = 0
      Case eBlinkFast:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
        deBlinkCycle(i) = 0
    End Select
  Next
  if(dqSound(dqHead) <> "") Then
    PlaySound(dqSound(dqHead) )
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
  if(dqHead = dqTail) Then
    if(dqbFlush(Head) = True) Then
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
    if(deCount(i) <> deCountEnd(i) ) Then
      deCount(i) = deCount(i) + 1

      select case(dqEffect(i, dqHead) )
        case eNone:
          Temp = dqText(i, dqHead)
        case eScrollLeft:
          Temp = Right(dLine(i), dCharsPerLine(i) - 1)
          Temp = Temp & Mid(dqText(i, dqHead), deCount(i), 1)
        case eScrollRight:
          Temp = Mid(dqText(i, dqHead), (dCharsPerLine(i) + 1) - deCount(i), 1)
          Temp = Temp & Left(dLine(i), dCharsPerLine(i) - 1)
        case eBlink:
          BlinkEffect = True
          if((deCount(i) MOD deBlinkSlowRate) = 0) Then
            deBlinkCycle(i) = deBlinkCycle(i) xor 1
          End If

          if(deBlinkCycle(i) = 0) Then
            Temp = dqText(i, dqHead)
          Else
            Temp = Space(dCharsPerLine(i) )
          End If
        case eBlinkFast:
          BlinkEffect = True
          if((deCount(i) MOD deBlinkFastRate) = 0) Then
            deBlinkCycle(i) = deBlinkCycle(i) xor 1
          End If

          if(deBlinkCycle(i) = 0) Then
            Temp = dqText(i, dqHead)
          Else
            Temp = Space(dCharsPerLine(i) )
          End If
      End Select

      if(dqText(i, dqHead) <> "_") Then
        dLine(i) = Temp
        DMDUpdate i
      End If
    End If
  Next

  if(deCount(0) = deCountEnd(0) ) and(deCount(1) = deCountEnd(1) ) and(deCount(2) = deCountEnd(2) ) Then

    if(dqTimeOn(dqHead) = 0) Then
      DMDFlush()
    Else
      if(BlinkEffect = True) Then
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
    TempStr = Space(dCharsPerLine(id) )
  Else
    if(Len(TempStr)> Space(dCharsPerLine(id) ) ) Then
      TempStr = Left(TempStr, Space(dCharsPerLine(id) ) )
    Else
      if(Len(TempStr) <dCharsPerLine(id) ) Then
        TempStr = TempStr & Space(dCharsPerLine(id) - Len(TempStr) )
      End If
    End If
  End If
  ExpandLine = TempStr
End Function

Function FormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
  dim i
  dim NumString

  NumString = CStr(abs(Num) )

  For i = Len(NumString) -3 to 1 step -3
    if IsNumeric(mid(NumString, i, 1) ) then
      NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1) ) + 48) & right(NumString, Len(NumString) - i)
    end if
  Next
  FormatScore = NumString
End function

Function CL(id, NumString)
  Dim Temp, TempStr
  Temp = (dCharsPerLine(id) - Len(NumString) ) \ 2
  TempStr = Space(Temp) & NumString & Space(Temp)
  CL = TempStr
End Function

Function RL(id, NumString)
  Dim Temp, TempStr
  Temp = dCharsPerLine(id) - Len(NumString)
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

Dim Digits, Chars(255), BumpS(30)

DMDInit

Sub DMDInit
  Dim i
  Digits = Array(digit001, digit002, digit003, digit004, digit005, digit006, digit007, digit008, digit009, digit010, _
  digit011, digit012, digit013, digit014, digit015, digit016, digit017, digit018, digit019, digit020, _
  digit021, digit022, digit023, digit024, digit025, digit026, digit027, digit028, digit029, digit030, _
  digit031, digit032, digit033, digit034, digit035, digit036, digit037, digit038, digit039, digit040, _
  digit041)

  For i = 0 to 255:Chars(i) = "dempty":Next

  Chars(32) = "dempty"
  Chars(43) = "dplus"    '+
  Chars(46) = "ddot"     '.
  Chars(48) = "d0"       '0
  Chars(49) = "d1"       '1
  Chars(50) = "d2"       '2
  Chars(51) = "d3"       '3
  Chars(52) = "d4"       '4
  Chars(53) = "d5"       '5
  Chars(54) = "d6"       '6
  Chars(55) = "d7"       '7
  Chars(56) = "d8"       '8
  Chars(57) = "d9"       '9
  Chars(60) = "dless"    '<
  Chars(61) = "dequal"   '=
  Chars(62) = "dmore"    '>
  Chars(64) = "bkempty"  '@
  Chars(65) = "da"       'A
  Chars(66) = "db"       'B
  Chars(67) = "dc"       'C
  Chars(68) = "dd"       'D
  Chars(69) = "de"       'E
  Chars(70) = "df"       'F
  Chars(71) = "dg"       'G
  Chars(72) = "dh"       'H
  Chars(73) = "di"       'I
  Chars(74) = "dj"       'J
  Chars(75) = "dk"       'K
  Chars(76) = "dl"       'L
  Chars(77) = "dm"       'M
  Chars(78) = "dn"       'N
  Chars(79) = "do"       'O
  Chars(80) = "dp"       'P
  Chars(81) = "dq"       'Q
  Chars(82) = "dr"       'R
  Chars(83) = "ds"       'S
  Chars(84) = "dt"       'T
  Chars(85) = "du"       'U
  Chars(86) = "dv"       'V
  Chars(87) = "dw"       'W
  Chars(88) = "dx"       'X
  Chars(89) = "dy"       'Y
  Chars(90) = "dz"       'Z
  Chars(94) = "dup"      '^
  '    Chars(95) = '_
  Chars(96) = "d0a"  '0.
  Chars(97) = "d1a"  '1. 'a
  Chars(98) = "d2a"  '2. 'b
  Chars(99) = "d3a"  '3. 'c
  Chars(100) = "d4a" '4. 'd
  Chars(101) = "d5a" '5. 'e
  Chars(102) = "d6a" '6. 'f
  Chars(103) = "d7a" '7. 'g
  Chars(104) = "d8a" '8. 'h
  Chars(105) = "d9a" '9  'i
  Chars(112) = "dp2" 'p 'p dark
  Chars(113) = "dk2" 'q 'k dark
  Chars(114) = "de2" 'r 'e dark
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

Dim red, orange, amber, yellow, darkgreen, green, blue, darkblue, purple, white, base,lightgrey,midgrey,darkgrey

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
lightgrey=12
midgrey=13
darkgrey=14

Sub SetLightColor(n, col, stat)
  Select Case col
    Case red
      n.color = RGB(18, 0, 0)
      n.colorfull = RGB(255, 0, 0)
    Case orange
      n.color = RGB(18, 3, 0)
      n.colorfull = RGB(255, 64, 64)
    Case amber
      n.color = RGB(193, 49, 0)
      n.colorfull = RGB(255, 153, 0)
    Case yellow
      n.color = RGB(18, 18, 0)
      n.colorfull = RGB(240, 220, 0)
    Case darkgreen
      n.color = RGB(0, 8, 0)
      n.colorfull = RGB(0, 150, 0) '(0, 64, 0)
    Case green
      n.color = RGB(0, 18, 0)
      n.colorfull = RGB(0, 200, 100) '(0, 255, 0)
    Case blue
      n.color = RGB(0, 18, 18)
      n.colorfull = RGB(0, 255, 255) '0, 255, 255
    Case darkblue
      n.color = RGB(0, 8, 8)
      n.colorfull = RGB(0, 20, 230) '(0, 0, 255)
    Case purple
      n.color = RGB(128, 128, 255)  '128.0,128
      n.colorfull = RGB(3, 0, 240)
    Case white
      n.color = RGB(255, 252, 224)
      n.colorfull = RGB(193, 91, 0)
    Case white
      n.color = RGB(255, 252, 224)
      n.colorfull = RGB(193, 91, 0)
    Case base
      n.color = RGB(128, 128, 255)
      n.colorfull = RGB(255, 252, 224)
    Case lightgrey
      n.color = RGB(192, 192, 192)
    Case midgrey
      n.color = RGB(120, 120, 120)
    Case darkgrey
      n.color = RGB(50, 50, 50)

  End Select
  If stat <> -1 Then
    n.State = 0
    n.State = stat
  End If
End Sub

'*************************
' Rainbow Changing Lights
'*************************

Sub ResetAllLightsColor ' Called at a new game
' SetLightColor LightShootAgain, red, -1
' SetLightColor l1,   red, -1
' SetLightColor l2,   orange, -1
' SetLightColor l3,   green, -1
' SetLightColor l4,   yellow, -1
' SetLightColor l75,  yellow, -1
' SetLightColor LightSword,   white, -1
'
' SetLightColor lGlory1,  white, -1
' SetLightColor lGlory2,  white, -1
' SetLightColor lGlory3,  white, -1
' SetLightColor lGlory4,  white, -1
' SetLightColor lGlory5,  white, -1
'
' SetLightColor l5,   purple, -1
' SetLightColor l6, purple, -1
' SetLightColor l7,   purple, -1
' SetLightColor l8,   purple, -1
' SetLightColor l9, purple, -1
' SetLightColor l10,  purple, -1
'
' SetLightColor l11,  red, -1
' SetLightColor l12,  red, -1
' SetLightColor l13,  red, -1
'
' SetLightColor l16,  yellow, -1
' SetLightColor l16A,   yellow, -1
' SetLightColor l17,  orange, -1
' SetLightColor l17A,   orange, -1
' SetLightColor l17B,   blue, -1
' SetLightColor l18,  blue, -1
' SetLightColor l18A, red, -1
' SetLightColor l18B, yellow, -1
' SetLightColor l18C, purple, -1
' SetLightColor l18D, blue, -1
' SetLightColor l19,  orange, -1
' SetLightColor l19A,   orange, -1
' SetLightColor l20,  orange, -1
' SetLightColor l20A,   orange, -1
' SetLightColor l21,  green, -1
' SetLightColor l21B,   green, -1
'
' SetLightColor l23A, green, -1
' SetLightColor l23B, yellow, -1
' SetLightColor l23C, white, -1
' SetLightColor l23D, red, -1
' SetLightColor l23E, green, -1
' SetLightColor l23F, yellow, -1
' SetLightColor l23G, white, -1
' SetLightColor l23H, red, -1
'
' SetLightColor LightSword,   white, -1
' SetLightColor l28,  white, -1
' SetLightColor l28A1,  white, -1
' SetLightColor l28A2,  white, -1
' SetLightColor l28A3,  white, -1
' SetLightColor l28A4,  white, -1
' SetLightColor l28A5,  white, -1
' SetLightColor l28A6,  white, -1
'
' SetLightColor l28B, darkblue, -1
' SetLightColor l28D, darkblue, -1
' SetLightColor l28E, darkblue, -1
' SetLightColor l28G, darkblue, -1
' SetLightColor l28H, darkblue, -1
' SetLightColor l_wl1, blue, -1
' SetLightColor l_wl2, blue, -1
' SetLightColor l_wl3, blue, -1
' SetLightColor l_wl4, blue, -1
' SetLightColor l24, purple, -1
'
' SetLightColor l30,  blue, -1
' SetLightColor l31,  blue, -1
' SetLightColor l32,  purple, -1
' SetLightColor l33,  red, -1
' SetLightColor l34,  darkgreen, -1
' SetLightColor l35,  yellow, -1
'
' SetLightColor l36,  red, -1
' SetLightColor l37,  red, -1
' SetLightColor l37a, white, -1
' SetLightColor l37b, red, -1
'
' SetLightColor l38,  red, -1
' SetLightColor l39,  red, -1
' SetLightColor l40,  red, -1
' SetLightColor l41,  red, -1
'
'
' SetLightColor l42,  red, -1
' SetLightColor l43,  red, -1
' SetLightColor l44,  red, -1
' SetLightColor l45,  red, -1
'
' SetLightColor l48,  yellow, -1
' SetLightColor l49,  orange, -1
' SetLightColor l50,  purple, -1
'
'
' SetLightColor l61,  red, -1
' SetLightColor l62A, white, -1
' SetLightColor l62B, white, -1
' SetLightColor l62C, white, -1
' SetLightColor l62Glory, yellow, -1
' SetLightColor l61A, yellow, -1
' SetLightColor l61B, orange, -1
'
'
' SetLightColor l63,  yellow, -1
' SetLightColor l64,  yellow, -1
' SetLightColor l65,  yellow, -1
' SetLightColor l66,  yellow, -1
' SetLightColor l67,  yellow, -1
' SetLightColor l68,  yellow, -1
' SetLightColor l69,  yellow, -1
' SetLightColor l70,  yellow, -1
' SetLightColor l71,  yellow, -1
' SetLightColor l72,  yellow, -1
' SetLightColor l80,  yellow, -1
' SetLightColor l81,  yellow, -1
' SetLightColor l82,  yellow, -1
' SetLightColor l83,  yellow, -1
' SetLightColor l84,  yellow, -1
' SetLightColor l85,  yellow, -1
' SetLightColor l86,  yellow, -1
' SetLightColor l87,  yellow, -1
' SetLightColor l88,  yellow, -1
' SetLightColor l89,  yellow, -1
'
' SetLightColor LightLeftKickback, blue, -1
' SetLightColor LightRightKickback, purple, -1
' SetLightColor LightLeftInlane,  white, -1
' SetLightColor LightRightInlane, white, -1
' SetLightColor Lightrightescape, blue, -1
' SetLightColor LightLeftEscape,  blue, -1
' SetLightColor LightClockBroken, white, -1
' SetLightColor LightHollagram, white, -1
'
' SetLightColor l_HourGlass, blue, -1
' SetLightColor l_HourGlass2, blue, -1
' SetLightColor Light_Catch2, blue, -1
' SetLightColor LightShootAgain, red, -1
' SetLightColor LightBallLock1, green, -1
' SetLightColor Light_WheelAward,blue, -1
' SetLightColor Light_WheelAwardb, orange, -1
' SetLightColor LightBallLock2,red, -1
'
' SetLightColor Light_Escape,blue, -1
' SetLightColor Light_BallSaver,purple, -1
' SetLightColor LightBird,blue, -1
'
' SetLightColor l_PlanetChaosEnabled, orange, -1
' SetLightColor l_PlanetReversEnabled, orange, -1
' SetLightColor l_PlanetBirdEnabled, orange, -1
' SetLightColor l_PlanetSpiderEnabled, orange, -1
' SetLightColor l_PlanetDragonEnabled, orange, -1
'
' 'transporter light flashes
' SetLightColor f3a1, white, -1
' SetLightColor f3a2, white, -1
' SetLightColor f3a3, white, -1
' SetLightColor f3a4, white, -1
' SetLightColor f3b1, white, -1
' SetLightColor f3b2, white, -1
' SetLightColor f3b3, white, -1
' SetLightColor f3b4, white, -1
' SetLightColor f3c1, white, -1
' SetLightColor f3c2, white, -1
' SetLightColor f3c3, white, -1
' SetLightColor f3c4, white, -1
' SetLightColor f3d1, white, -1
' SetLightColor f3d2, white, -1
' SetLightColor f3d3, white, -1
' SetLightColor f3d4, white, -1
' SetLightColor ftl, white, -1
' SetLightColor ftl2, white, -1
' SetLightColor ftl3, white, -1
' SetLightColor ftl4, white, -1
'
'
'
'
' 'Set GI light clours
'
'
' SetLightColor Bumper_Light1,  white, -1
' SetLightColor Bumper_Light1b,   white, -1
' SetLightColor Bumper_Light2,  white, -1
' SetLightColor Bumper_Light2b,   white, -1
' SetLightColor Bumper_Light3,  white, -1
' SetLightColor Bumper_Light3b,   white, -1
' SetLightColor Bumper_Light4,  white, -1
' SetLightColor Bumper_Light4b,   white, -1

End Sub

Sub UpdateBonusColors
End Sub





'Sub RotateLaneLightsLeft
'   Dim TempState
'   'flipper lanes
'   TempState = ll1.State
'   ll1.State = ll2.State
'   ll2.State = ll3.State
'   ll3.State = ll4.State
'   ll4.State = ll5.State
'   ll5.State = TempState
'   ll8.state = ll1.state
'   ll7.state = ll2.state
'   ll6.state = ll3.state
'   ll9.state = ll4.state
'   ll20.state = ll5.state
' End Sub

' Sub RotateLaneLightsRight
'   Dim TempState
'   'flipperlanes
'   TempState = ll5.State
'   ll5.State = ll4.State
'   ll4.State = ll3.State
'   ll3.State = ll2.State
'   ll2.State = ll1.State
'   ll1.State = TempState
'   ll8.state = ll1.state
'   ll7.state = ll2.state
'   ll6.state = ll3.state
'   ll9.state = ll4.state
'   l20.state = ll5.state
' End Sub



'*************************
' Rainbow Changing Lights
'*************************

Dim RGBStep, RGBFactor, rRed, rGreen, rBlue, RainbowLights

Sub StartRainbow(n)
' set RainbowLights = n
' RGBStep = 0
' RGBFactor = 5
' rRed = 255
' rGreen = 0
' rBlue = 0
' RainbowTimer.Enabled = 1
End Sub


Sub StopRainbow(n)
' Dim obj
' RainbowTimer.Enabled = 0
' RainbowTimer.Enabled = 0
' For each obj in RainbowLights
'   SetLightColor obj, "white", 0
' Next
End Sub


Sub RainbowTimer_Timer 'rainbow led light color changing
' Dim obj
' Select Case RGBStep
'   Case 0 'Green
'     rGreen = rGreen + RGBFactor
'     If rGreen > 255 then
'       rGreen = 255
'       RGBStep = 1
'     End If
'   Case 1 'Red
'     rRed = rRed - RGBFactor
'     If rRed < 0 then
'       rRed = 0
'       RGBStep = 2
'     End If
'   Case 2 'Blue
'     rBlue = rBlue + RGBFactor
'     If rBlue > 255 then
'       rBlue = 255
'       RGBStep = 3
'     End If
'   Case 3 'Green
'     rGreen = rGreen - RGBFactor
'     If rGreen < 0 then
'       rGreen = 0
'       RGBStep = 4
'     End If
'   Case 4 'Red
'     rRed = rRed + RGBFactor
'     If rRed > 255 then
'       rRed = 255
'       RGBStep = 5
'     End If
'   Case 5 'Blue
'     rBlue = rBlue - RGBFactor
'     If rBlue < 0 then
'       rBlue = 0
'       RGBStep = 0
'     End If
' End Select
' For each obj in RainbowLights
'   obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
'   obj.colorfull = RGB(rRed, rGreen, rBlue)
' Next
End Sub

Sub RainbowTimer1_Timer 'rainbow led light color changing
' Dim obj
' Select Case RGBStep2
'   Case 0 'Green
'     rGreen2 = rGreen2 + RGBFactor2
'     If rGreen2 > 255 then
'       rGreen2 = 255
'       RGBStep2 = 1
'     End If
'   Case 1 'Red
'     rRed2 = rRed2 - RGBFactor2
'     If rRed2 < 0 then
'       rRed2 = 0
'       RGBStep2 = 2
'     End If
'   Case 2 'Blue
'     rBlue2 = rBlue2 + RGBFactor2
'     If rBlue2 > 255 then
'       rBlue2 = 255
'       RGBStep2 = 3
'     End If
'   Case 3 'Green
'     rGreen2 = rGreen2 - RGBFactor2
'     If rGreen2 < 0 then
'       rGreen2 = 0
'       RGBStep2 = 4
'     End If
'   Case 4 'Red
'     rRed2 = rRed2 + RGBFactor2
'     If rRed2 > 255 then
'       rRed2 = 255
'       RGBStep2 = 5
'     End If
'   Case 5 'Blue
'     rBlue2 = rBlue2 - RGBFactor2
'     If rBlue2 < 0 then
'       rBlue2 = 0
'       RGBStep2 = 0
'     End If
' End Select
' For each obj in RainbowLights2
'   obj.color = RGB(rRed2 \ 10, rGreen2 \ 10, rBlue2 \ 10)
'   obj.colorfull = RGB(rRed2, rGreen2, rBlue2)
' Next
End Sub

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo

  If bGameInPLay = True Then: vpmtimer.addtimer 2200,  "Player1Now'": Exit Sub


  Dim tmp
  'info goes in a loop only stopped by the credits and the startkey
  If Score(1)Then
    DMD CL(0, "LAST SCORE"), CL(1, "PLAYER1 " &FormatScore(Score(1))), "", eNone, eNone, eNone, 10000, False, ""
  End If
  If Score(2)Then
    DMD CL(0, "LAST SCORE"), CL(1, "PLAYER2 " &FormatScore(Score(2))), "", eNone, eNone, eNone, 10000, False, ""
  End If
  If Score(3)Then
    DMD CL(0, "LAST SCORE"), CL(1, "PLAYER3 " &FormatScore(Score(3))), "", eNone, eNone, eNone, 10000, False, ""
  End If
  If Score(4)Then
    DMD CL(0, "LAST SCORE"), CL(1, "PLAYER4 " &FormatScore(Score(4))), "", eNone, eNone, eNone, 10000, False, ""
  End If
  DMD "", CL(1, "GAME OVER"), "", eNone, eBlink, eNone, 700, False, ""
  If bFreePlay Then
    DMD CL(0, "FREE PLAY"), CL(1, "PRESS START"), "", eNone, eBlink, eNone, 2000, False, ""
  Else
    If Credits > 0 Then
      DMD CL(0, "CREDITS " & Credits), CL(1, "PRESS START"), "", eNone, eBlink, eNone, 400, False, ""
    Else
      DMD CL(0, "CREDITS " & Credits), CL(1, "INSERT COIN"), "", eNone, eBlink, eNone, 2000, False, ""
    End If
    If bGameInPlay=True Then:DMDFlush:DMD "", "", "DMD_Blank", eNone, eNone, eNone, 100, False, "":vpmtimer.addtimer 2200,"Player1Now'"
    End If





  DMD "", "", "DMD_Blank", eNone, eNone, eNone, 1000, False, "" 'blank
  DMD "", "", "DMD_FlipperSelect2", eNone, eBlink, eNone, 5000, False, "" 'Avago (Have a Go)
' DMD "", "", "DMD_PinkFlamingos", eNone, eNone, eNone, 3000, False, "" 'Avago Presents(Have a Go)
' DMD "", "", "DMD_CrackersBar", eNone, eNone, eNone, 3000, False, "" 'Avago Presents(Have a Go)
' DMD "", "", "DMD_TimeLord", eNone, eNone, eNone, 3000, False, "" 'Crew
' DMD "", "", "DMD_Coldy", eNone, eNone, eNone, 3000, False, "" 'Crew
' DMD "", "", "DMD_CheersA", eNone, eNone, eNone, 3000, False, "" 'Crew
  DMD "", "", "DMD_AvagoPresents2", eNone, eNone, eNone, 3000, False, "" 'Avago Presents(Have a Go)
  DMD "", "", "DMD_TimeLord3", eNone, eNone, eNone, 3000, False, "" 'Avago Presents(Have a Go)
  DMD "", "", "DMD_RusstyT2", eNone, eNone, eNone, 3000, False, "" 'Russty
' DMD "", "", "DMD_StoryByHP3", eNone, eNone, eNone, 3000, False, "" 'HP
  DMD "", "", "DMD_Crackers2", eNone, eNone, eNone, 3000, False, "" 'Crackers
  DMD "", "", "DMD_JPSalasFramework", eNone, eNone, eNone, 3000, False, "" 'JPSalas
  DMD "", "", "DMD_FrankyT2", eNone, eNone, eNone, 3000, False, "" 'HP
  DMD "", "", "DMD_Rawd2", eNone, eNone, eNone, 3000, False, "" 'Rawd
  DMD "", "", "DMD_Oqqsan2", eNone, eNone, eNone, 3000, False, "" 'Oqqsan
  DMD "", "", "DMD_Gedankeko2", eNone, eNone, eNone, 3000, False, "" 'GedanKekoJote
  DMD "", "", "DMD_Tomate2", eNone, eNone, eNone, 3000, False, "" 'Tomate
  DMD "", "", "DMD_Outhere2", eNone, eNone, eNone, 3000, False, "" 'Outhere
' DMD "", "", "DMD_Apophis3", eNone, eNone, eNone, 3000, False, "" 'Apophis
' DMD "", "", "DMD_Wylte2", eNone, eNone, eNone, 3000, False, "" 'Wylte
  DMD "", "", "DMD_Smaug3", eNone, eNone, eNone, 3000, False, "" 'Smaurg
' DMD "", "", "DMD_JoeSoap2", eNone, eNone, eNone, 3000, False, "" 'JoeSoap
  DMD "", "", "DMD_RajoJoey2", eNone, eNone, eNone, 3000, False, "" 'Rajo
  DMD "", "", "DMD_Avago1b", eNone, eNone, eNone, 3000, False, "" 'Crew
  DMD "", "", "DMD_Avago2b", eNone, eNone, eNone, 3000, False, "" 'Crew

  DMD CL(0, "HIGHSCORES"), Space(dCharsPerLine(1)), "", eScrollLeft, eScrollLeft, eNone, 20, False, ""
  DMD CL(0, "HIGHSCORES"), "", "", eBlinkFast, eNone, eNone, 1000, False, ""
  DMD CL(0, "HIGHSCORES"), "1> " &HighScoreName(0) & " " &FormatScore(HighScore(0)), "", eNone, eScrollLeft, eNone, 2000, False, ""
  DMD "_", "2> " &HighScoreName(1) & " " &FormatScore(HighScore(1)), "", eNone, eScrollLeft, eNone, 2000, False, ""
  DMD "_", "3> " &HighScoreName(2) & " " &FormatScore(HighScore(2)), "", eNone, eScrollLeft, eNone, 2000, False, ""
  DMD "_", "4> " &HighScoreName(3) & " " &FormatScore(HighScore(3)), "", eNone, eScrollLeft, eNone, 2000, False, ""
  DMD Space(dCharsPerLine(0)), Space(dCharsPerLine(1)), "", eScrollLeft, eScrollLeft, eNone, 500, False, ""

End Sub


'**************************************************************************************
'  ATTRACT MODE
'*************************************************************************************
'


Sub StartAttractMode()
  PlaySong "m_The Unicorn & The Wasp"
  vpmtimer.addtimer 500, "PleaseSelect'"
  bAttractMode = True
  Stopblights
  StartLightSeq
  StartFlasherDomesAttractMode
  ShowTableInfo
  StartRainbow alights
  StopClockHandsCounterClockwise
  StartClockHandsClockwise
  l_Staff2.State=0 'TurnOff TQMan Ball Lights
  l_Staff1.State=0
  StartPlanetChaos
  StartPlanetReverseTime
  StartPlanetSpider
  StartPlanetBird
  StartPlanetDragon1
  TimeThiefVisible
  QManInvisible
  Hologram1Start
  SetLightColor textoverlay, midgrey, -1
  SetLightColor TimeThief, darkgrey, -1
  DragonMoveUp
  ChangeToyTimer.Enabled=1
End Sub

Dim ChangeToy:ChangeToy=0
Sub ChangeToyTimer_Timer
    ChangeToy=ChangeToy+1
  Select Case ChangeToy
    Case 1:SpiderMoveUp:DragonMoveDown
    Case 2:SpiderMoveDown:BirdMoveUp
    Case 3:BirdMoveDown:DragonMoveUp:ChangeToy=0
  End Select
End Sub

Sub PleaseSelect
  PlaySound "CO_PleaseSelect"
End Sub

Sub StopAttractMode()
  DragonMoveDown:SpiderMoveDown:BirdMoveDown
  ChangeToyTimer.Enabled=0
  bAttractMode = False
  DragonMoveDown:SpiderMoveDown:BirdMoveDown
  If Hard=0 Then
    DMDFlush
    DMD "", "", "DMD_Blank", eNone, eNone, eNone, 100, False, "" 'blank
    DMD CL(0, "EASY MODE " ), CL(1, ""), "", eNone, eNone, eNone, 1500, True, ""

  End If
  If Hard=1 Then
    DMDFlush
    DMD "", "", "DMD_Blank", eNone, eNone, eNone, 100, False, "" 'blank
    DMD CL(0, "HARD MODE " ), CL(1, ""), "", eNone, eNone, eNone, 1500, True, ""

  End If
  LightSeqAttract.StopPlay
  LightSeqAttract2.StopPlay
  '-----------------------------------
  StopRainbow alights
  '--------------------------------
  ResetAllLightsColor

  '   TurnOnStartOfGameLights
  StopPlanetChaos
  StopPlanetReverseTime
  StopPlanetSpider
  StopPlanetBird
  StopPlanetDragon1

  SetLightColor textoverlay, midgrey, -1
  SetLightColor TimeThief, lightgrey, -1
End Sub

Sub StartLightSeq()

  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqUpOn, 30, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqDownOn, 25, 1

  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqStripe1VertOn, 15, 2
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqStripe2VertOn, 10, 2
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqCircleOutOn, 15, 2

  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqCircleOutOn, 15, 2
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqUpOn, 25, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqDownOn, 25, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqUpOn, 25, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqDownOn, 25, 1
  LightSeqAttract.UpdateInterval = 15
  LightSeqAttract.Play SeqCircleOutOn, 15, 3
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqRightOn, 50, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqLeftOn, 35, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqRightOn, 50, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqCircleOutOn, 15, 2
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqLeftOn, 40, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqRightOn, 40, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqLeftOn, 20, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqRightOn, 40, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqCircleOutOn, 15, 3
  LightSeqAttract.UpdateInterval = 100
  LightSeqAttract.Play SeqLeftOn, 25, 1
  LightSeqAttract.UpdateInterval = 100
  LightSeqAttract.Play SeqRightOn, 25, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqCircleOutOn, 15, 2
  LightSeqAttract.Play SeqStripe2VertOn, 50, 2
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqLeftOn, 25, 1
  LightSeqAttract.UpdateInterval = 10
  LightSeqAttract.Play SeqUpOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqDownOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqUpOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqDownOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqStripe1VertOn, 50, 2
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqCircleOutOn, 15, 2
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqStripe1VertOn, 50, 3
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqLeftOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqRightOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqLeftOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqUpOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqDownOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqCircleOutOn, 15, 2
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqStripe2VertOn, 50, 3
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqLeftOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqRightOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqLeftOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqUpOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqDownOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqUpOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqDownOn, 25, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqStripe1VertOn, 25, 3
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqStripe2VertOn, 25, 3
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqUpOn, 15, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqDownOn, 15, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqRightOn, 15, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqLeftOn, 15, 1
  LightSeqAttract.UpdateInterval = 20
  LightSeqAttract.Play SeqRightOn, 15, 1
  LightSeqAttract.UpdateInterval = 20
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

Sub Stopblights
  Light_BlackCube1.State=0
  Light_BlackBox2.State=0

  Light_WomanLL.State=0
  Light_WomanRL.State=0
  LightFireClock2.State=0
  LightFireClock3.State=0
  Light_ClockGreen1.State=0
  Light_ClockGreen2.State=0
  Light_Red1.State=0
  LightRed2.State=0
  Light_DragonHead.State=0
  Light_WheelRed4.State=0
  Light_LeftLaneRedHand.State=0
  Light_ReverseGreen.State=0
  Light_RedHand.State=0
  Light_RightOrbitLane.State=0
  Light_WormLane.State=0
End Sub

Sub Startblights
  Light_BlackCube1.State=1
  Light_BlackBox2.State=1

  Light_WomanLL.State=1
  Light_WomanRL.State=1
  LightFireClock2.State=1
  LightFireClock3.State=1
  Light_ClockGreen1.State=1
  Light_ClockGreen2.State=1
  Light_Red1.State=1
  LightRed2.State=1
  Light_DragonHead.State=1
  Light_WheelRed4.State=1
  Light_LeftLaneRedHand.State=1
  Light_ReverseGreen.State=1
  Light_RedHand.State=1
  Light_RightOrbitLane.State=1
  Light_WormLane.State=1
End Sub

'**********************
'FlasherSequences
'**********************

Dim FlasherStepCount:FlasherStepCount=0

Sub StartFlasherDomesAttractMode
    Objlevel(3) = 1:FlasherFlash3_Timer:Objlevel(4) = 1:FlasherFlash4_Timer
    Objlevel(1) = 1:FlasherFlash1_Timer:Objlevel(2) = 1:FlasherFlash2_Timer
End Sub

Sub StartResetNewBallFlashers
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqBlinking,, 4,100
End Sub

Sub FlasherDomesLeftLoop
  LightSeqLeftLoop.UpdateInterval = 10
  LightSeqLeftLoop.Play SeqBlinking,,3, 150
End Sub

Sub FlasherDomesRightLoop
  LightSeqRightLoop.UpdateInterval = 10
  LightSeqRightLoop.Play SeqBlinking,, 4, 150
End Sub

Sub StartFlasherSpiralAward
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqBlinking,, 8,100
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqLeftOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqRightOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqCircleOutOn, 50, 1
End Sub

Sub StartFlasherAward
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqBlinking,, 8,100
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqLeftOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqRightOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqCircleOutOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqLeftOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqRightOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqCircleOutOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqLeftOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqRightOn, 50, 1
  LightSeqFlasherDomes1.UpdateInterval = 10
  LightSeqFlasherDomes1.Play SeqCircleOutOn, 50, 1
End Sub

'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************




Sub Player1Now
  DMD "", CL(1, "PLAYER 1 BALL 1 " ), "", eNone, eNone, eNone, 3000, True, ""
End Sub

Sub Game_Init() 'called at the start of a new game
  Dim i, j
  bExtraBallWonThisBall = False
  PlaySong "m_LockedOn"
  lrflashtime.Enabled = False
  bExtraBallWonThisBall = False
  MechTilt = 0
  bMechTiltJustHit = False
  bFlippersEnabled = True
    For i = 1 to 4
        SkillshotValue(i) = 50000000
  Next
  StopAttractMode
  ResetStartofGameVariables()
  TurnOnStartOfGameLights
AlertCallout=0
End Sub



Sub StopEndOfBallMode() 'this sub is called after the last ball is drained
' If bExtraBallWonThisBall = True Then:Exit Sub
  SaveLightStates
  clearlights
End Sub


Sub ResetNewBallVariables()
'Do a reload of current play states if Extra ball is awarded
'Do a reset of modes and awards
  debug.print " ResetNewBallVariables Sub Routine"
  If bExtraBallWonThisBall=True Then bExtraBallWonThisBall=False:EnableBallSaver 30'Need Ball Saver if extra ball is achieved

  vpmtimer.addtimer 200, "ResetNewBallLights'"
  bFlippersEnabled = True
  If SuperLoopsComplete(CurrentPlayer)=True Then StartPlanetBird
  If WormLocksComplete(CurrentPlayer)=True Then StartPlanetSpider
  If ChaosCriticalComplete(CurrentPlayer)=True Then StartPlanetChaos
  If ReverseTimeComplete(CurrentPlayer)=True Then StartPlanetReverseTime
  If DragonWounded4(CurrentPlayer)=True Then StartPlanetDragon1
  If GloryTableOn(CurrentPlayer)=True And bExtraBallWonThisBall=True Then EBActive=False:ResetStartofGameVariables(): Exit Sub
  If GloryTableOn(CurrentPlayer)=True Then ResetStartofGameVariables():Exit Sub
  If bExtraBallWonThisBall=True Then bExtraBallWonThisBall=False:EBActive=False:PlaySound "CO_ReturnToSender":ChangeSong:debug.print "NoModeandAwardReset":Exit Sub

  vpmtimer.addtimer 1000, "ResetModesAndAwards'"'This must occur after loading players last lights. This turns off meteor shower light, worm lights, super loops etc
 debug.print "Modes&AwardsReset"

' vpmtimer.addtimer 500, "StopDragon'" This is stopped in Turn off Light States
End Sub

Sub ResetNewBallLights()                                 'turn on or off the needed lights before a new ball is released
  LoadLightStates'LoadLightStates 'ensure the multiplier is displayed right
debug.print "LoadLightStates"
  LightSword.State=2
  l4.State=2
  GiOn
  End Sub

Sub ResetStartofGameVariables()
debug.print " StopPlanets"
debug.print " ResetStartOfGameVariables"
    StopPlanets
    GloryComplete(CurrentPlayer)=False
debug.print "GloryTableOn=False-ResetSOGV"
    Wall_OrbitDropper. IsDropped=True
    GiOn
    kickbacklg.open=True
    kickbackrg.open=True
    GateClogBlocker.open = True
    GateWormLocks.open = True
    Hologram1Start
    RightStarsTimer.enabled=False
    LightWheelTimer.enabled=False
    LordComplete= False
    TimeComplete=False
    KickerLairOut.Enabled=False
    MeteorShowerStarted=False
    DudCount =0
    DudCount2=0
    WLRCount =0
    WoundDragon=0
    NewSong =0
    GravityCall=0
    Crazy=0
    LLCO=0
    bBallSaverReady = True
    bBallSaverActive=False
    EBActive=0
    LiftBarricadeTargets
    DragonFlashBlink=0
    DragonTargetFlasherBlinkTimer.Enabled=1
    OOCCount(CurrentPlayer)=1
    ChaosCriticalReady(CurrentPlayer)=False
    ChaosCriticalComplete(CurrentPlayer)=False
    SuperLoopsReady(CurrentPlayer)=False
    SuperLoopsComplete(CurrentPlayer)=False
    ReverseLock1Ready(CurrentPlayer)=False
    ReverseLock2Ready(CurrentPlayer)=False
    ReverseLock1Complete(CurrentPlayer)=False
    ReverseLock2Complete(CurrentPlayer)=False
    ReverseTimeComplete(CurrentPlayer)=False
    WormLocksActivated(CurrentPlayer)=False
    WormLocksComplete(CurrentPlayer)=False
    DragonWounded1(CurrentPlayer)=False
    DragonWounded2(CurrentPlayer)=False
    DragonWounded3(CurrentPlayer)=False
    DragonWounded4(CurrentPlayer)=False
    DragonWallDropActivated=False
    Billions(CurrentPlayer)=0
    AdvanceTimeActivated=0
    ExtraBallActivated=0
    SuperLoopsComplete(CurrentPlayer)=False
    If GloryTableOn(CurrentPlayer)=False Then vpmTimer.AddTimer 500, "ResetModesAndAwards'"
    GloryTableOn(CurrentPlayer)=False  'ResetGloryTableOn flag if true
    WallReverseOutProtect.IsDropped=True

End Sub

Sub ResetModesAndAwards   'This must occur after loading players last lights. This turns off meteor shower light, worm lights, super loops etc
debug.print " ResetModesAndAwards"
  GloryComplete(CurrentPlayer)=False
  GloryTableOn(CurrentPlayer)=False
debug.print "GloryTableOn=False-ResetModes&Awards"
  WoundDragon=0
  GateClogBlocker.open = True
  GateWormLocks.open = True
  StopSuperLoops
  StopSuperJets
  StopAdvanceTime
  StopExtraBall
  StopLostInSpace
  StopOutOfControl
  StopChainedLightning
  ResetLeftTargetBankLights
  StopMeteorShowerLights
  DisableWormLocks
  TWJPFiring=False
  ResetRightSpinLights
  ResetLeftSpinLights
  LiftWall
  ResetBarricadeTargets
  ResetDragonLairTargetLights'*************************************************************SET UP TO ATTACK THE DRAGON
  LightLeftInlane.State=0 :LightLeftEscape.State=2 :kickbackleftenabled
  LightRightInlane.State=0:Lightrightescape.State=2: kickbackrightenabled
  KickerLairOut.Enabled=False
  TimeLordCatchEnabled=False
  l_Staff2.State=0
  ResetTimeWarpJackPotsEndOfBall
  CheckLowerRampL
  If ReverseLock1Ready(CurrentPlayer)=True Then RaiseRampL:KickerReverseLock1.Enabled=True:KickerReverseLock2.Enabled=False
  If ReverseLock2Ready(CurrentPlayer)=True Then RaiseRampL:KickerReverseLock1.Enabled=false:KickerReverseLock2.Enabled=True
    ResetReverseTimeLights
  If DragonUp=True Then DragonMoveDown
  If SpiderUp=True Then SpiderMoveDown
  If BirdUp=True Then BirdMoveDown
End Sub



Sub ResetGloryOtherBits
debug.print " GloryOtherbits"
  Wall_OrbitDropper. IsDropped=True
  GiOn
  kickbacklg.open=True
  kickbackrg.open=True
  GateClogBlocker.open = True
  Hologram1Start
  RightStarsTimer.enabled=False
  LightWheelTimer.enabled=False
  LordComplete= False
  TimeComplete=False
  KickerLairOut.Enabled=False
  MeteorShowerStarted=False
  DudCount =0
  DudCount2=0
  WLRCount =0
  WoundDragon=0
  NewSong =0
  GravityCall=0
  Crazy=0
  LLCO=0
' bBallSaverReady = True
  EBActive=0
  LiftBarricadeTargets
  StartReverseLockLights
End Sub

Sub TurnOffPlayfieldLights()
  Dim a
  For each a in aLights
    a.State = 0
  Next
End Sub

Sub CheckLowerRampL
  If LightBallLock1.State=2 Then: Exit Sub
  If LightBallLock2.State=2 Then: Exit Sub
  LowerRampL
End Sub

Sub TurnOnStartOfGameLights
debug.print " TurnOnStartOfGameLights"
  LightLeftInlane.State=0:LightRightInlane.State=0:LightLeftEscape.State=2:Lightrightescape.State=2  'Turn on in lane lights to activate kickbacks
  LightRightKickback.State=1:LightLeftKickback.State=1  'Turn on Skull Lights for ambience
  LightClockBroken.State=0:LightHollagram.State=1 'turn on lights for ambience
  l_HourGlass.State=2
  l_HourGlass2.State=2
  LightBird.State=1         'Initial Pop Bumper Score
  l31.State=2         'Enable First Skill Shot light
  l48.State=2     '10000 jet Start value
  l61A.State=2
  ResetLightWheel
  SetLightColor l_Staff1, blue, -1 'not included in alghts. associated with Qman Flasher
  SetLightColor l_Staff2, blue, -1 'not included in alghts. associated with Qman Flasher
  l_ChaosKicker.State=1:Light_WheelAwardb.State=1' Give the kickers some glow
  Startblights
  LightClockBroken.State=1
  Light_easyhard.State=1
  vpmtimer.addtimer 2000, "DelayLightsTurnOn'"
  GiOn

  PortalBlue.Visible=1:PortalGreen.Visible=0: PortalOrange.Visible=0:PortalRed.Visible=0:PortalYellow.Visible=0
' If  PortalYellow.Visible=True Then: l35.State=2: l34.State=0:l33.State=0:l32.State=0:l31.State=0:PortalBlue.Visible=0:PortalGreen.Visible=0: PortalOrange.Visible=0:PortalRed.Visible=0:PortalYellow.Visible=1:Exit Sub
' If  PortalRed.Visible=True Then: l35.State=0: l34.State=2:l33.State=0:l32.State=0:l31.State=0:PortalBlue.Visible=0:PortalGreen.Visible=0: PortalOrange.Visible=0:PortalRed.Visible=1:PortalYellow.Visible=0:Exit Sub
' If  PortalOrange.Visible=True Then: l35.State=0: l34.State=0:l33.State=2:l32.State=0:l31.State=0:PortalBlue.Visible=0:PortalGreen.Visible=0: PortalOrange.Visible=1:PortalRed.Visible=0:PortalYellow.Visible=0:Exit Sub
' If  PortalGreen.Visible=True Then: l35.State=0: l34.State=0:l33.State=0:l32.State=2:l31.State=0:PortalBlue.Visible=0:PortalGreen.Visible=1: PortalOrange.Visible=0:PortalRed.Visible=0:PortalYellow.Visible=0:Exit Sub
' If  PortalBlue.Visible=True Then: l35.State=0: l34.State=0:l33.State=0:l32.State=0:l31.State=2:PortalBlue.Visible=1:PortalGreen.Visible=0: PortalOrange.Visible=0:PortalRed.Visible=0:PortalYellow.Visible=0:Exit Sub

End Sub



Sub DelayLightsTurnOn
debug.print " DelayStartOfGameLightsOn"
  If Hard = 0 Then l16a.State=0:l19.State=1:l20.State=1:l21.State=2:l19A.State=0:l20A.State=0:l21B.State=1:l17B.State=1:End If 'SetUpReverseTime & ChaosCritical for Easy Mode
  If Hard=1 Then l19.State=2:l20.State=0:l21.State=0:l19A.State=1:l20A.State=0:l21B.State=0:l16A.State=1:l17B.State=0:End If'SetUpReverseTime & ChaosCritical for Hard Mode
  l11.State=2:l12.State=2:l13.State=2 :l16.State=0:l17.State=0:l18.State=0'Left Target Bank Initiation
  l61a.State=2
End Sub

Sub ResetReverseTimeLights
  debug.print " ResetReverseTimeLights"
  If Hard = 0 Then l16a.State=0:l19.State=1:l20.State=1:l21.State=2:l19A.State=0:l20A.State=0:l21B.State=1:l17B.State=1:End If 'SetUpReverseTime & ChaosCritical for Easy Mode
  If Hard=1 Then l19.State=2:l20.State=0:l21.State=0:l19A.State=1:l20A.State=0:l21B.State=0:l16A.State=1:l17B.State=0:End If'SetUpReverseTime & ChaosCritical for Hard Mode
  l11.State=2:l12.State=2:l13.State=2 :l16.State=0:l17.State=0:l18.State=0'Left Target Bank Initiation
  l61a.State=2
End Sub

Sub UpdateSkillShot()
  debug.print " ChangePlungerPortalFlasher"
    If CreateMultiballTimer.Enabled = True Then: Exit Sub
  If  PortalYellow.Visible=True Then: l35.State=0: l34.State=0:l33.State=0:l32.State=0:l31.State=2:PortalBlue.Visible=1:PortalGreen.Visible=0: PortalOrange.Visible=0:PortalRed.Visible=0:PortalYellow.Visible=0:Exit Sub
  If  PortalRed.Visible=True Then: l35.State=2: l34.State=0:l33.State=0:l32.State=0:l31.State=0:PortalBlue.Visible=0:PortalGreen.Visible=0: PortalOrange.Visible=0:PortalRed.Visible=0:PortalYellow.Visible=1:Exit Sub
  If  PortalOrange.Visible=True Then: l35.State=0: l34.State=2:l33.State=0:l32.State=0:l31.State=0:PortalBlue.Visible=0:PortalGreen.Visible=0: PortalOrange.Visible=0:PortalRed.Visible=1:PortalYellow.Visible=0: Exit Sub
  If  PortalGreen.Visible=True Then: l35.State=0: l34.State=0:l33.State=2:l32.State=0:l31.State=0:PortalBlue.Visible=0:PortalGreen.Visible=0: PortalOrange.Visible=1:PortalRed.Visible=0:PortalYellow.Visible=0: Exit Sub
  If  PortalBlue.Visible=True Then: l35.State=0: l34.State=0:l33.State=0:l32.State=2:l31.State=0:PortalBlue.Visible=0:PortalGreen.Visible=1: PortalOrange.Visible=0:PortalRed.Visible=0:PortalYellow.Visible=0:Exit Sub

End Sub



'***********************************
'Light Save & restore- Thanks JP
'***********************************
Dim MyLightStates(4,200)

Sub SaveLightStates
debug.print " SaveLightStates"
  Dim i, tmp
  i = 0
  For each tmp in aLights
    MyLightStates(CurrentPlayer,i) = tmp.State
    i = i + 1
  Next
  StopPlanets
  'ChangeSong
End Sub

Sub LoadLightStates
debug.print " LoadLightStates"
  Dim i, tmp
  i = 0
  For each tmp in aLights
    tmp.State = MyLightStates(CurrentPlayer,i)
    i = i + 1
  Next
  If ReverseLock1Ready(CurrentPlayer)=True Then RaiseRampL:KickerReverseLock1.Enabled=True:KickerReverseLock2.Enabled=False:End If
  If ReverseLock2Ready(CurrentPlayer)=True Then RaiseRampL:KickerReverseLock1.Enabled=False:KickerReverseLock2.Enabled=True:End If
  If ChaosCriticalComplete(CurrentPlayer)=1 Then:StartPlanetChaos : End If
  If ReverseTimeComplete(CurrentPlayer)=1 Then:StartPlanetReverseTime : End If
  If SuperLoopsComplete(CurrentPlayer)=1 Then:StartPlanetBird : End If
  If WormLocksComplete(CurrentPlayer)=1 Then:StartPlanetSpider : End If
  If DragonWounded4(CurrentPlayer) Then:StartPlanetDragon1 : End If
  StopSuperLoops
  StopSuperJets
  StopWormLights
  StopDragon
  UpdateSkillShot
End Sub

Sub clearlights
  dim i
  for each i in aLights
    i.State = 0
  next
End Sub



Sub StopDragon
  ResetDragonLairTargetLights
  ResetBarricadeTargets
  LiftBarricadeTargets
  KickerLairOut.Enabled=False
  vpmTimer.AddTimer 500, "TimeThiefVisible'"
  vpmTimer.AddTimer 200, "QManInvisible'"
  vpmTimer.AddTimer 500, "Hologram1Start'"


End Sub









'**********************
'     GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'**********************


Sub GiOn
  Dim bulb
  For each bulb in GI
    bulb.State = 1
  Next

End Sub

Sub GiOff
  Dim bulb
  For each bulb in GI
    bulb.State = 0
  Next

End Sub



' GI & light sequence effects

Sub GiEffect(n)
  Select Case n
    Case 0 all off
      LightSeqGi.Play SeqAlloff
    Case 1 all blink
      LightSeqGi.UpdateInterval = 4
      LightSeqGi.Play SeqBlinking, , 5, 100
    Case 2 random
      LightSeqGi.UpdateInterval = 10
      LightSeqGi.Play SeqRandom, 5, , 1000
    Case 3 'upon
      LightSeqGi.UpdateInterval = 4
      LightSeqGi.Play SeqUpOn, 5, 1
    Case 4  left-right-left
      LightSeqGi.UpdateInterval = 5
      LightSeqGi.Play SeqLeftOn, 10, 1
      LightSeqGi.UpdateInterval = 5
      LightSeqGi.Play SeqRightOn, 10, 1
  End Select
End Sub

Sub LightEffect(n)
  Select Case n
    Case 0  all off
      LightSeqInserts.Play SeqAlloff
    Case 1 all blink
      LightSeqInserts.UpdateInterval = 4
      LightSeqInserts.Play SeqBlinking, , 5, 100
    Case 2 random
      LightSeqInserts.UpdateInterval = 10
      LightSeqInserts.Play SeqRandom, 5, , 1000
    Case 3 upon
      LightSeqInserts.UpdateInterval = 4
      LightSeqInserts.Play SeqUpOn, 10, 1
    Case 4  left-right-left
      LightSeqInserts.UpdateInterval = 5
      LightSeqInserts.Play SeqLeftOn, 10, 1
      LightSeqInserts.UpdateInterval = 5
      LightSeqInserts.Play SeqRightOn, 10, 1
    Case 5 random
      LightSeqbumper.UpdateInterval = 4
      LightSeqbumper.Play SeqBlinking, , 5, 10
    Case 6 random
      LightSeqRSling.UpdateInterval = 4
      LightSeqRSling.Play SeqBlinking, , 5, 6
    Case 7 random
      LightSeqLSling.UpdateInterval = 4
      LightSeqLSling.Play SeqBlinking, , 5, 6
    Case 8 random
      LightSeqBack.UpdateInterval = 4
      LightSeqBack.Play SeqBlinking, , 5, 6
    Case 12 random
      LightSeqlr.UpdateInterval = 4
      LightSeqlr.Play SeqBlinking, , 5, 10
  End Select
End Sub

' Flasher Effects using lights

' Dim FEStep, FEffect
' FEStep = 0
' FEffect = 0
'
' Sub FlashEffect(n)
'   Select case n
'     Case 0 ' all off
'       LightSeqFlasher.Play SeqAlloff
'     Case 1 'all blink
'       LightSeqFlasher.UpdateInterval = 4
'       LightSeqFlasher.Play SeqBlinking, , 5, 100
'     Case 2 'random
'       LightSeqFlasher.UpdateInterval = 10
'       LightSeqFlasher.Play SeqRandom, 5, , 1000
'     Case 3 'upon
'       LightSeqFlasher.UpdateInterval = 4
'       LightSeqFlasher.Play SeqUpOn, 10, 1
'     Case 4 ' left-right-left
'       LightSeqFlasher.UpdateInterval = 5
'       LightSeqFlasher.Play SeqLeftOn, 10, 1
'       LightSeqFlasher.UpdateInterval = 5
'       LightSeqFlasher.Play SeqRightOn, 10, 1
'     Case 5 ' top flashers blink fast
'   End Select
' End Sub







' '****************************
' ' Flashers - Thanks Flupper
' '****************************
'
' Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4, FlashLevel5, FlashLevel6
' Flasherlight4.IntensityScale = 0
'
' '*** right red flasher ***
' Sub Flasherflash3_Timer()
'   dim flashx3, matdim
'   If not Flasherflash3.TimerEnabled Then
'     Flasherflash3.TimerEnabled = True
'     Flasherflash3.visible = 1
'     Flasherlit3.visible = 1
'   End If
'   flashx3 = FlashLevel3 * FlashLevel3 * FlashLevel3
'   Flasherflash3.opacity = 1500 * flashx3
'   Flasherlit3.BlendDisableLighting = 10 * flashx3
'   Flasherbase3.BlendDisableLighting =  flashx3
'   Flasherlight4.IntensityScale = flashx3
'   matdim = Round(10 * FlashLevel3)
'   Flasherlit3.material = "domelit" & matdim
'   FlashLevel3 = FlashLevel3 * 0.9 - 0.01
'   If FlashLevel3 < 0.15 Then
'     Flasherlit3.visible = 0
'   Else
'     Flasherlit3.visible = 1
'   end If
'   If FlashLevel3 < 0 Then
'     Flasherflash3.TimerEnabled = False
'     Flasherflash3.visible = 0
'   End If
' End Sub


'****************************
' Backwall Lamps
'****************************
Sub TriggerBackLamp2_Hit()
  If LastSwitchHit ="Trigger4" Then: Exit Sub:End If
  FlasherDomesRightLoop
  RightWallLightsOn

End Sub

Sub RightWallLightsOn
  Objlevel(3) = 1:FlasherFlash3_Timer:Objlevel(4) = 1:FlasherFlash4_Timer
End Sub

Sub RightWallLightsOff
  Objlevel(3) = 1:FlasherFlash3_Timer:Objlevel(4) = 1:FlasherFlash4_Timer
End Sub

Sub TriggerBackLamp1_Hit()
' FlasherDomesLeftLoop
  LeftWallLightsOn
End Sub



Sub LeftWallLightsOn
Objlevel(3) = 1:FlasherFlash3_Timer:Objlevel(4) = 1:FlasherFlash4_Timer
End Sub

Sub LeftWallLightsOff
Objlevel(3) = 1:FlasherFlash3_Timer:Objlevel(4) = 1:FlasherFlash4_Timer

End Sub




Sub bwallon
  Flasherflash3.visible = 1:Flasherflash4.visible = 1
End Sub

Sub bwalloff
  Flasherflash3.visible = 0:Flasherflash4.visible = 0
End Sub


'****************************
'  SECONDARY HIT EVENTS
'******************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  BG_Sequence 36,0,50
  LS_RightSling.Play SeqBlinking,, 1,100
  RS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotRight SlingshotRightSound
  DOF 104, DOFPulse
  RSling.Visible = 0
  RSling1.Visible = 1
  sling1.rotx = 20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
  AddScore 0
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
    Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
' Objlevel(1) = 1:FlasherFlash1_Timer
  BG_Sequence 36,0,50
  LS_LeftSling.Play SeqBlinking,, 1,100
  LS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft SlingshotLeftSound
  DOF 103, DOFPulse
  LSling.Visible = 0
  LSling1.Visible = 1
  sling2.rotx = 20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  AddScore 0
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
    Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub


' Sub RotateLaneLightsLeft
'   Dim TempState
'   TempState = l6.State
'   l6.State = l7.State
'   l7.State = l8.State
'   l8.State = TempState
' End Sub

' Sub RotateLaneLightsRight
'   Dim TempState
'   TempState = l8.State
'   l8.State = l7.State
'   l7.State = l6.State
'   l6.State = TempState
' End Sub





'************************************
'  MAIN SHOTS - PRIMARY HIT EVENTS
'************************************

'*********************************
'  BUMPERS
'**********************************

'*************************
'Bumper1
'*************************



spiralcaps1.enabled = 0
Sub spiralspin1
  spiralcaps1.enabled = 1
End Sub

Sub spiralcaps1_Timer
  BumperCap1.ObjRotZ = (BumperCap1.ObjRotZ - 1) Mod 360
  If BumperCap1.ObjRotZ = -359 Then
    spiralcaps1.enabled = 0
  End If
End Sub


Sub Bumper1_Hit

  If NOT Tilted Then
    BG_Sequence 36,0,50
    DOF 105,DOFPulse
    RandomSoundBumperTop Bumper1
    LightSeqBumper1.UpdateInterval = 4
    LightSeqBumper1.Play SeqBlinking,, 5, 25
    BumperSequence.enabled = 1
    spiralspin1
    If l48.State=2 Then: AddScore 10000: End If
    If l49.State=2 Then: AddScore 10000000: End If
  End If
End Sub

'*************************
'Bumper2
'*************************


spiralcaps2.enabled = 0
Sub spiralspin2
  spiralcaps2.enabled = 1
End Sub

Sub spiralcaps2_Timer
  BumperCap2.ObjRotZ = (BumperCap2.ObjRotZ - 1) Mod 360
  If BumperCap2.ObjRotZ = -359 Then
    spiralcaps2.enabled = 0
  End If
End Sub

Sub Bumper2_Hit
  If NOT Tilted Then
    BG_Sequence 36,0,50
    DOF 106,DOFPulse
    RandomSoundBumperMiddle Bumper2
    LightSeqBumper2.UpdateInterval = 4
    LightSeqBumper2.Play SeqBlinking,, 5, 25
    BumperSequence.enabled = 1
    spiralspin2
    If l48.State=2 Then: AddScore 10000: End If
    If l49.State=2 Then: AddScore 10000000: End If

  End If
End Sub

'*************************
'Bumper3
'*************************

spiralcaps3.enabled = 0
Sub spiralspin3
  spiralcaps3.enabled = 1
End Sub

Sub spiralcaps3_Timer
  BumperCap3.ObjRotZ = (BumperCap3.ObjRotZ - 1) Mod 360
  If BumperCap3.ObjRotZ = -359 Then
    spiralcaps3.enabled = 0
  End If
End Sub

Sub Bumper3_Hit

  If NOT Tilted Then
    BG_Sequence 36,0,50
    DOF 107,DOFPulse
    RandomSoundBumperBottom Bumper3
    LightSeqBumper3.Play SeqBlinking,,  5, 25
    BumperSequence.enabled = 1
    spiralspin3
    If l48.State=2 Then: AddScore 10000: End If
    If l49.State=2 Then: AddScore 10000000: End If
  End If
End Sub

'*************************
'Bumper4
'*************************


spiralcaps4.enabled = 0
Sub spiralspin4
  spiralcaps4.enabled = 1
End Sub

Sub spiralcaps4_Timer
  BumperCap4.ObjRotZ = (BumperCap4.ObjRotZ - 1) Mod 360
  If BumperCap4.ObjRotZ = -359 Then
    spiralcaps4.enabled = 0
  End If
End Sub

Sub Bumper4_Hit
  If NOT Tilted Then
    BG_Sequence 36,0,50
    DOF 108 ,DOFPulse
    RandomSoundBumperTop Bumper4
    LightSeqbumper4.UpdateInterval = 4
    LightSeqbumper4.Play SeqBlinking,,  5, 25
    BumperSequence.enabled = 1
    spiralspin4
    If l48.State=2 Then: AddScore 10000: End If
    If l49.State=2 Then: AddScore 10000000: End If
  End If
End Sub

'*****************************************
'Escape for ball caught at back of kicker
'******************************************
'Sub KickerBumperEscape_hit()

''    KickerBumperEscape.DestroyBall
'   WormKick3.TimerEnabled=1
'   KickerBumperEscape2.Enabled=1
'   CreateNewBallKickerBumperEscape2
'End Sub


' Sub CreateNewBallKickerBumperEscape2
' PlaySoundAt "VUKEnter", KickerBumperEscape2
' KickerBumperEscape2.CreateSizedball BallSize / 2
' KickerBumperEscape2.Kick 120, 15
' KickerBumperEscape2.Enabled=0
' vpmtimer.addtimer 1500, "light.State=0"
' End Sub

'

'**********************************************************************************
'Skillshot Kicker destroys ball and creates ball at right outlane portal
'This is acombined Openingshot portal that awards a skillhot if v the light is lit

'***********************************************************************************
Sub SkillShotLights
    LS_MajorAward.Play SeqBlinking,, 4,100
End Sub


'********************
'SkillShot1
'********************





Sub Kicker32_hit()
  LightPortalBlue.State=2
  WireRampOff
  Light_Escape.State=2
  PlaySoundAt "Ball_Bounce_Playfield_Hard_1", Kicker32
  PlaySoundAt "VUKEnter", Kicker32
  If l31.State=0 And FirstBallStarted=True Then CheckDragonAttackMusic
  If l31.State=0 Then vpmTimer.AddTimer 3000, "UpdateNonSkillShotScore'"
  If l31.State=2 Then
    AwardSkillshot
    CheckDragonAttackMusic
    UpdateSkillShot
    l30.State=2
    SkillShotLights
'   DMD CL(0, " "), CL(1, "SKILL SHOT"), "", eNone, eBlink, eNone, 1500, True, ""
'   vpmTimer.AddTimer 3020, "UpdateSkillShotScore'"
  End If
    GatePortalProtect.open = False
    vpmtimer.addtimer 500, "Liftoff '"
    vpmtimer.addtimer 1500, "DestroyKickerBall1 '"
End Sub

Sub CheckDragonAttackMusic
  If TWJPFiring=True Then :Exit Sub
  If StartWormLocksTimer.Enabled=True Then: Exit Sub
    PlaySong "m_Unit"
End Sub

Sub UpdateSkillShotScore
    AddScore 50000000
    l30.State=0
End Sub

Sub UpdateNonSkillShotScore
    AddScore 5000
    vpmtimer.addtimer 400, "UpdateSkillShot '"
    l49.State=0
End Sub


Sub LiftOff
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_PortalOpen":End If
End Sub


Sub DestroyKickerBall1

  Kicker32.DestroyBall
  PlaySound "fx_woosh"
  RaiseOrbitWall
  CreateNewBallKicker2
End Sub

Sub CreateNewBallKicker2

  Light_Escape.State=2
  PlaySoundAt "VUKEnter", Kicker2
  Kicker2.TimerEnabled=1
  Kicker2.CreateSizedball BallSize / 2
  vpmtimer.addtimer 1000, "EscapePortalExit '"
  vpmtimer.addtimer 2000, "DropOrbitWall '"
End Sub



'********************
'SkillShot2
'********************
Sub Kicker33_hit()
  LightPortalGreen.State=2
  WireRampOff
  Light_Escape.State=2
  PlaySoundAt "Ball_Bounce_Playfield_Hard_1", Kicker33
  PlaySoundAt "VUKEnter", Kicker33
  If l32.State=0 Then:vpmTimer.AddTimer 3000, "UpdateNonSkillShotScore'":End If

  If l32.State=2 Then
    AwardSkillshot
    CheckDragonAttackMusic
    UpdateSkillShot
    l30.State=2
    SkillShotLights
'   DMD CL(0, " "), CL(1, "SKILL SHOT"), "", eNone, eBlink, eNone, 1500, True, ""
'   vpmTimer.AddTimer 3020, "UpdateSkillShotScore'"
  End If
  GatePortalProtect.open = False
  vpmtimer.addtimer 500, "Liftoff '"
  vpmtimer.addtimer 1500, "DestroyKickerBall2 '"
End Sub



Sub DestroyKickerBall2
  Kicker33.DestroyBall
  Kicker2.TimerEnabled=1
  RaiseOrbitWall
  EscapePortalExit
  CreateNewBallKicker2
End Sub


'***************
'SkillShot3
'***************
Sub Kicker34_hit()
  LightPortalOrange.State=2
  WireRampOff
  Light_Escape.State=2
  PlaySoundAt "Ball_Bounce_Playfield_Hard_1", Kicker34
  PlaySoundAt "VUKEnter", Kicker34
  If l33.State=0 Then:vpmTimer.AddTimer 3000, "UpdateNonSkillShotScore'":End If
  If l33.State=2 Then
    AwardSkillshot
    CheckDragonAttackMusic
    UpdateSkillShot
    l30.State=2
    SkillShotLights
'   DMD CL(0, " "), CL(1, "SKILL SHOT"), "", eNone, eBlink, eNone, 1500, True, ""
'   vpmTimer.AddTimer 3020, "UpdateSkillShotScore'"
  End If
  GatePortalProtect.open = False
  vpmtimer.addtimer 500, "Liftoff '"
  vpmtimer.addtimer 1500, "DestroyKickerBall3 '"
End Sub

Sub DestroyKickerBall3
  Kicker34.DestroyBall
  Kicker2.TimerEnabled=1
  RaiseOrbitWall
  EscapePortalExit
  CreateNewBallKicker2
End Sub

'*********************
'SkillShot4
'*********************
Sub Kicker35_hit()
  LightPortalRed.State=2
  WireRampOff
  Light_Escape.State=2
  PlaySoundAt "Ball_Bounce_Playfield_Hard_1", Kicker35
  PlaySoundAt "VUKEnter", Kicker35
  If l34.State=0 Then: vpmTimer.AddTimer 3000, "UpdateNonSkillShotScore'":End If

  If l34.State=2 Then
    AwardSkillshot
    CheckDragonAttackMusic
    UpdateSkillShot
    l30.State=2
    SkillShotLights
'   DMD CL(0, " "), CL(1, "SKILL SHOT"), "", eNone, eBlink, eNone, 1500, True, ""
'   vpmTimer.AddTimer 3020, "UpdateSkillShotScore'"
  End If
  GatePortalProtect.open = False
  vpmtimer.addtimer 500, "Liftoff '"
  vpmtimer.addtimer 1500, "DestroyKickerBall4 '"
End Sub

Sub DestroyKickerBall4
  Kicker35.DestroyBall
  PlaySoundAt "Ball_Drop_Playfield_5_Delayed",Kicker35 'BalldropSound
  Kicker2.TimerEnabled=1
  RaiseOrbitWall
  EscapePortalExit
  CreateNewBallKicker2
End Sub

'******************
'SkillShot5
'******************
Sub Kicker36_hit()
  LightPortalYellow.State=2
  WireRampOff
  Light_Escape.State=2
  PlaySoundAt "Ball_Bounce_Playfield_Hard_1", Kicker36
  PlaySoundAt "VUKEnter", Kicker36
  If l35.State=0 Then: vpmTimer.AddTimer 3000, "UpdateNonSkillShotScore'"
  If l35.State=2 Then
    If bMultiBallMode=False Then AwardSkillshot:End If
    CheckDragonAttackMusic
    UpdateSkillShot
    l30.State=2
    SkillShotLights
'   DMD CL(0, " "), CL(1, "SKILL SHOT"), "", eNone, eBlink, eNone, 1500, True, ""
'   vpmTimer.AddTimer 3020, "UpdateSkillShotScore'"
  End If
  GatePortalProtect.open = False
  vpmtimer.addtimer 500, "Liftoff '"
  vpmtimer.addtimer 1500, "DestroyKickerBall5 '"

End Sub

Sub DestroyKickerBall5
  Kicker36.DestroyBall
  Kicker2.TimerEnabled=1
  RaiseOrbitWall
  EscapePortalExit
  CreateNewBallKicker2
End Sub

Sub Wall_OrbitDropper_Dropped
  vpmtimer.addtimer 1000, "DropOrbitWall '"
End Sub

Sub RaiseOrbitWall
  PlaySoundAt "fx_shaker",trigger7
  Wall_OrbitDropper.IsDropped=False

End Sub

Sub DropOrbitWall
  '   PlaySoundAt "fx_rubber2",trigger4
  Wall_OrbitDropper.IsDropped=True
End Sub



'****************************
'EscapePortalFlasher
'****************************

Sub EscapePortalExit()
  LightPortalBlue.State=0:LightPortalGreen.State=0:LightPortalOrange.State=0:LightPortalRed.State=0:LightPortalYellow.State=0
  FlashForMs f3a1, 500, 50, 0
  FlashForMs f3a2, 500, 50, 0
  FlashForMs f3a3, 500, 50, 0
  FlashForMs f3a4, 500, 50, 0
  PlaySoundAt "Popper" , Kicker2
' PlaySound "Woosh3woosh"
  DOF 121, DOFPulse
  Kicker2.Kick 10, 45
  Kicker2.TimerEnabled=0
  Light_Escape.State=0
  GatePortalProtect.open = True
End Sub

Sub PortalTimer_Timer

  PortalBlue.rotz = (PortalBlue.rotz -4)mod 360
  PortalGreen.rotz = (PortalGreen.rotz -4)mod 360
  PortalOrange.rotz = (PortalOrange.rotz -4)mod 360
  PortalRed.rotz = (PortalRed.rotz -4)mod 360
  PortalYellow.rotz = (PortalYellow.rotz -4)mod 360
End Sub

'***********************************************
'RightSpinner for LightWheelChanging
'**************************************
Dim RightSpins

Sub SpinnerRight_Spin
  RightSpins=RightSpins+1
  DOF 120,DOFPulse

  SoundSpinner SpinnerRight
  If Not Tilted Then
    CountRightSpins
  End If
End Sub

Sub CountRightSpins
' If GloryTableTimer.Enabled=1 Then: Exit Sub
  If LastSwitchHit = "Trigger5" Then:Exit Sub
  RightSpins=RightSpins+1
  vpmtimer.addtimer 1500, "CheckRightStars '"
End Sub

Sub CheckRightStars
  If RightSpins>=90 Then
    l89.State=1
    vpmTimer.AddTimer 2000, "RightStarsBlink'"
    DMD CL(0, "1000000"), CL(1, "AAAAW SHINY STARS"), "", eNone, eBlink, eNone, 1500, True, ""
    vpmTimer.AddTimer 2200, "DMDScoreNow'"
    AddScore 1000000
  End If

  If RightSpins>=80 Then l88.State=1 'RightSpins is reset to 0 at CheckWheelLightActivation
  If RightSpins>=70 Then l87.State=1:AddScore 20000
  If RightSpins>=60 Then l86.State=1
  If RightSpins>=50 Then l85.State=1:AddScore 20000
  If RightSpins>=40 Then l84.State=1
  If RightSpins>=30 Then l83.State=1:AddScore 20000
  If RightSpins>=20 Then l82.State=1
  If RightSpins>=10 Then : l81.State=1 : l80.State=1 : AddScore 20000:End If
End Sub

Sub RightStarsBlink
  l80.State = 2:l81.State = 2: l82.State = 2: l83.State = 2: l84.State = 2: l85.State = 2: l86.State = 2: l87.State = 2: l88.State = 2: l89.State = 2
  vpmtimer.addtimer 2200, "ResetRightSpinLights'"
End Sub


Sub ResetRightSpinLights
  l80.State=0: l81.State=0: l82.State=0: l83.State=0: l84.State=0: l85.State=0: l86.State=0: l87.State=0: l88.State=0:l89.State=0

  RightSpins=0
  debug.print "ResetRightSpins"
End Sub

'**************************
'LightWheel
'***************************
Sub Trigger4_Hit()
  ' PlaySound "WooshHigh2"
  If LastSwitchHit="Trigger9" Then :WheelLightRotate
  LastSwitchHit = "Trigger4"
End Sub

Sub Trigger9_Hit()
  LastSwitchHit = "Trigger9"
End Sub


Dim WLRCount
Sub WheelLightRotate
  WLRCount = WLRCount + 1
  Select Case WLRCount
    Case 1: l23A.State=2:l23B.State=0:l23C.State=0:l23D.State=0:l23E.State=0:l23F.State=0:l23G.State=0:l23H.State=0:Light_WheelAward.State=1
    Case 2: l23A.State=0:l23B.State=2:l23C.State=0:l23D.State=0:l23E.State=0:l23F.State=0:l23G.State=0:l23H.State=0:Light_WheelAward.State=1
    Case 3: l23A.State=0:l23B.State=0:l23C.State=2:l23D.State=0:l23E.State=0:l23F.State=0:l23G.State=0:l23H.State=0:Light_WheelAward.State=1
    Case 4: l23A.State=0:l23B.State=0:l23C.State=0:l23D.State=2:l23E.State=0:l23F.State=0:l23G.State=0:l23H.State=0:Light_WheelAward.State=1
    Case 5: l23A.State=0:l23B.State=0:l23C.State=0:l23D.State=0:l23E.State=2:l23F.State=0:l23G.State=0:l23H.State=0:Light_WheelAward.State=1
    Case 6: l23A.State=0:l23B.State=0:l23C.State=0:l23D.State=0:l23E.State=0:l23F.State=2:l23G.State=0:l23H.State=0:Light_WheelAward.State=1
    Case 7: l23A.State=0:l23B.State=0:l23C.State=0:l23D.State=0:l23E.State=0:l23F.State=0:l23G.State=2:l23H.State=0:Light_WheelAward.State=1
      Case 8: If LightShootAgain.State=1 Then :l23A.State=2:WLRCount=1:End If
      l23A.State=0:l23B.State=0:l23C.State=0:l23D.State=0:l23E.State=0:l23F.State=0:l23G.State=0:l23H.State=2:Light_WheelAward.State=1:WLRCount =0
  End Select
End Sub



'****************************
'LightWheelKicker
'****************************

Sub Kicker6_hit()
  Objlevel(2) = 1:FlasherFlash2_Timer
  Light_WheelAward.State=2:Light_WheelAwardb.State=2
  PlaySoundAt "fx_kicker_enter" , Kicker6
  vpmtimer.addtimer 3000, "PlaySoundAt""BallRelease3"", Kicker6: Kicker6.kick 197, 18'"
  If lGlory1.State=2 Then: Exit Sub
  If l23A.State + l23B.State + l23C.State + l23D.State + l23E.State + l23F.State + l23G.State + l23H.State= 0 Then: TakeThePiss
  If l23H.state=2 Then:vpmtimer.addtimer 100, "ActivateExtraBall'":l23H.State=0:  Objlevel(2) = 1:FlasherFlash2_Timer
  If l23G.state=2 Then:vpmtimer.addtimer 100, "ChainedLightning'":l23G.State=0: Objlevel(2) = 1:FlasherFlash2_Timer
  If l23F.State=2 Then:vpmtimer.addtimer 100, "CheckWormHoleJackpots'":l23F.State=0:  Objlevel(2) = 1:FlasherFlash2_Timer
  If l23E.State=2 Then:vpmtimer.addtimer 100, "OutofControlAward'":l23E.State=0:  Objlevel(2) = 1:FlasherFlash2_Timer
  If l23D.State=2 Then:vpmtimer.addtimer 100, "ActivateAdvanceTime'":l23D.State=0:  Objlevel(2) = 1:FlasherFlash2_Timer
  If l23C.State=2 Then:vpmtimer.addtimer 100, "SuperJets'":l23C.State=0:  Objlevel(2) = 1:FlasherFlash2_Timer
  If l23B.State=2 Then:vpmtimer.addtimer 100, "Mystery'":l23B.State=0:  Objlevel(2) = 1:FlasherFlash2_Timer
  If l23A.State=2 Then:vpmtimer.addtimer 100, "SuperLoops'":l23A.State=0: Objlevel(2) = 1:FlasherFlash2_Timer
  If Light_WheelAward.State=1 Then:LS_WheelLights.Play SeqBlinking,, 10,100

End Sub

Sub ResetLightWheel
  l23A.State=2:l23B.State=0:l23C.State=0:l23D.State=0:l23E.State=0:l23F.State=0:l23G.State=0:l23H.State=0:Light_WheelAward.State=0
End Sub

Sub Kicker6_unhit()
  PlaySoundAt "fx_kicker" , Kicker6
  Light_WheelAward.State=0:Light_WheelAwardb.State=1
  debug.print "Kicker6 Kick"
  DOF 113,DOFPulse
End Sub

'*************************
'CalloutLightWheelNotLit
'*************************
Dim DudCount
Sub TakeThePiss
  BG_Sequence 37,0,110

  DudCount = DudCount + 1
  Select Case DudCount
    Case 1: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_DidYouMean":End If
    Case 2: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_Knock":End If
    Case 3: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_PlanLuck"  :End If
    Case 4: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_PleaseConcentrate":End If
        DudCount=0
  End Select
End Sub

'************************
'SuperLoops
'************************
Sub SuperLoops
  Objlevel(2) = 1:FlasherFlash2_Timer
  BirdMoveUp
  SuperLoopsReady(CurrentPlayer)=True
  GiOff
  vpmtimer.addtimer 2200, "GiOn'"
  LS_MajorAward.Play SeqLeftOn,, 1,100
  l50.State=2:LightBird.State=2
  Wall_OrbitDropper.IsDropped=False
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_SuperLoopsActivated":End If
  DMD "", "", "DMD_SLActivated", eNone, eNone, eNone, 2000, True, ""
  vpmTimer.AddTimer 2820, "DMDResetNoScoreAdded'"
  SuperLoopsTimer.Enabled =True
  debug.print "StartSuperLoopstimer"
  TriggerSuperLoops.Enabled=1
  debug.print "ActivateTriggerSuperLoops"
  LS_SuperLoopsActivated.Play SeqBlinking,, 10,100
  BG_Sequence 1,0,200

End Sub

Sub TriggerSuperLoops_Hit()
  'debug.print "TriggerSuperLoopsHit"
  If  LastSwitchHit = "Trigger1B" Then: Exit Sub
  l41.State=1
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_SuperLoops":End If                     '-------------------------------------------------------D in Lord
  PlaySound "blackbirds4sec"

  debug.print "StartAwardSuperLoopsSub"
  SuperLoopsAwardLights
  AwardSuperLoops

End Sub

Sub SuperLoopsTimer_Timer
  If Hard=0 Then:vpmtimer.addtimer 40000, "StopSuperLoops '" Else:vpmtimer.addtimer 30000, "StopSuperLoops '"
  debug.print "Disable SuperloopsTimer"
End Sub

'*************************
'SuperLoopsAward- DLight
'**************************

Sub AwardSuperLoops
  If SuperLoopsComplete(CurrentPlayer)=False Then  SuperLoopsComplete(CurrentPlayer)=True:StartPlanetBird:CheckTimeLordStatusTimer.Enabled=1:debug.print "SuperLoops(CurrentPlayer)=True":   End If
  GiOff
  vpmtimer.addtimer 1000, "GiOn'"
  LS_MajorAward.Play SeqLeftOn,, 1,100
  DMD "", "", "DMD_SuperLoops", eNone, eNone, eNone, 2200, True, ""
  vpmTimer.AddTimer 2820, "UpdateSuperLoopsScoreScore'"
  debug.print "AwardSuperLoops"
End Sub

Sub SuperLoopsAwardLights
  LS_MajorAward.UpdateInterval = 20
  LS_MajorAward.Play SeqStripe2VertOn, 35, 1

End Sub

Sub UpdateSuperLoopsScoreScore
  DMDScoreNow
  AddScore 200000000
End Sub

Sub DMDResetNoScoreAdded
  DMDScoreNow
  AddScore 0
End Sub


Sub DMDBlank
  DMD "", "", "DMD_Blank", eNone, eNone, eNone, 200, False, "" 'blank
End Sub

Sub StopSuperLoops
  BirdMoveDown
  SuperLoopsReady(CurrentPlayer)=False
  debug.print "SuperLoopsTriggerDisabled"
  TriggerSuperLoops.Enabled=0
  SuperLoopsTimer.Enabled =False
  l50.State=0:LightBird.State=1
  vpmtimer.addtimer 500, "DropOrbitWall'"
End Sub


'***********************
'Light Extra Ball
'***********************
Dim ExtraBallActivated
Sub ActivateExtraBall
  BG_Sequence 1,0,200

  If LightShootAgain.State=1 Then AwardSuperJackPot: Exit Sub
  ExtraBallActivated=1
  debug.print "ExtraBallActivated"
  GiOff
  LS_MajorAward.Play SeqBlinking,, 6,100
  vpmtimer.addtimer 2000, "GiOn'"
  l37a.State=2:l37b.State=2
  l37.State=2
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_ExtraBallLit":End If
  DMD "", "", "DMD_ExtraBallLit2", eNone, eNone, eNone, 3000,True, ""
  LS_AwardReady.Play SeqBlinking,, 10,100
  vpmtimer.addtimer 3200, "DMDScoreNow'"
' AddScore 0

If EasyMode=True Then:vpmtimer.addtimer 40000, "StopExtraBall'":End If
If HardMode=True Then:vpmtimer.addtimer 30000, "StopExtraBall'":End If
End Sub

Sub StopExtraBall
  l37.State=0:l37a.State=0:l37b.State=1: l36.State=0:ExtraBallActivated=0

End Sub

'**********************
'Chained Lightning
'**********************

Sub ChainedLightning
  BG_Sequence 1,0,200

  GiOff
  LS_MajorAward.Play SeqBlinking,, 8,100
  vpmtimer.addtimer 3200, "GiOn'"
  ChainedLightningTimer.Enabled=True
  l28.State=2:l28A1.State=2:l28A2.State=2:l28A3.State=2:l28A4.State=2:l28A5.State=2:l28A6.State=2
  PlaySound  "CO_ChainedLightningActivated"
  DMD "", "", "DMD_ChainedLightningActivated5", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 3200, "CLActScore'"
End Sub

Sub CLActScore
  DMDScoreNow
  AddScore 10000
End Sub


Sub TargetCaptiveBall_hit()
  PlaySoundAt "fx_popper",TargetCaptiveBall
  If l28.State=0 Then:AddScore 500: Exit Sub
  ChainedLightningFlash
  LightningBoltandBall
  PlaySound "Thunder"

  If l28A2.State=1 Then l28A1.State=1: l28.State=1:AwardChainedLightning: Exit Sub
  If l28A3.State=1 Then l28A2.State=1:ChainedLightningMinorAward:Exit Sub
  ChainedLightningMinorAward
  l28A3.state=1
End sub

Sub AwardChainedLightning
  AwardJackpot
End Sub

Sub ChainedLightningMinorAward
  DMD "", "", "DMD_ChainedLightning5", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 3200, "CLMinorAward'"
End Sub

Sub CLMinorAward
  DMDScoreNow
  AddScore 20000000
End Sub

Sub LightningBoltandBall
  GiOff
  LS_MajorAward.UpdateInterval = 20
  LS_MajorAward.Play SeqDownOn, 25, 1
  LightningBolt.Visible=True
  LightningBall.Visible=True
  vpmtimer.addtimer 400, "StopLightning'"

End Sub

Sub StopLightning
  LightningBolt.Visible=False
  LightningBall.Visible=False
  vpmtimer.addtimer 1500, "GiOn'"
End Sub


Sub ChainedLightningFlash()
  FlashForMs fcb1, 500, 50, 0
  FlashForMs fcb2, 500, 50, 0
  FlashForMs fcb3, 500, 50, 0
  FlashForMs fcb4, 500, 50, 0

  DOF 121, DOFPulse
End Sub

Sub ChainedLightningTimer_Timer
  If Hard=0 Then:vpmtimer.addtimer 40000, "StopChainedLightning '" Else:vpmtimer.addtimer 30000, "StopChainedLightning '"
  debug.print "ChainedLightning Timer"
End Sub

Sub StopChainedLightning
  l28.State=0:l28A1.State=0:l28A2.State=0:l28A3.State=0:l28A4.State=0:l28A5.State=0:l28A6.State=0
  ChainedLightningTimer.Enabled=False
End Sub

'************************
'OutofControl
'************************
Sub OutofControl
  OutofControlAward
End Sub

'Dim OOCCount
Sub OutofControlAward

' Select Case Int(Rnd*4)+1
  Select Case OOCCount(CurrentPlayer)
    Case 1:   GiOff
          vpmtimer.addtimer 3200, "GiOn'"
          LS_MajorAward.Play SeqBlinking,, 6,100
          OOCCount(CurrentPlayer)=OOCCount(CurrentPlayer)+1
        If SuperLoopsComplete(CurrentPlayer) =False Then
          SuperLoops
        End If
        If SuperLoopsComplete(CurrentPlayer)=True Then
        If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_SpaceShip":End If
          DMD "", "", "DMD_OOC", eNone, eNone, eNone, 3000, True, ""
          vpmTimer.AddTimer 3500, "MeanderingMoonPoodle'"
        End If


    Case 2:   GiOff
          vpmtimer.addtimer 3200, "GiOn'"
          LS_MajorAward.Play SeqBlinking,, 6,100
          OOCCount(CurrentPlayer)=OOCCount(CurrentPlayer)+1
        If SuperLoopsComplete(CurrentPlayer)=False Then
          SuperLoops
        End If
        If SuperLoopsComplete(CurrentPlayer)=True Then
        If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_PleaseConcentrate":End If
          DMD "", "", "DMD_OOC", eNone, eNone, eNone, 3000, True, ""
          vpmTimer.AddTimer 3500, "OOCScore2'"
        End If

    Case 3:   GiOff
          vpmtimer.addtimer 3200, "GiOn'"
          LS_MajorAward.Play SeqBlinking,, 6,100
          OOCCount(CurrentPlayer)=OOCCount(CurrentPlayer)+1
        If SuperLoopsComplete(CurrentPlayer)=False Then
          SuperLoops
        End If
        If SuperLoopsComplete(CurrentPlayer)=True Then
        If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_PowerOfStupidity":End If
          DMD "", "", "DMD_OOC", eNone, eNone, eNone, 3000, True, ""
          vpmTimer.AddTimer 3500, "OOCScore3'"
        End If

    Case 4:   GiOff
          vpmtimer.addtimer 3200, "GiOn'"
          LS_MajorAward.Play SeqBlinking,, 6,100
          OOCCount(CurrentPlayer)=1
        If SuperLoopsComplete(CurrentPlayer)=False Then
          SuperLoops
        End If
        If SuperLoopsComplete(CurrentPlayer)=True Then
        If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_Fart":End If
          DMD "", "", "DMD_OOC", eNone, eNone, eNone, 3000, True, ""
          vpmTimer.AddTimer 3500, "OOCScore4'"
        End If

  End Select
End Sub

Sub OOCScore1
    DMDScoreNow
    AddScore 10000000
End Sub

Sub OOCScore2
    DMDScoreNow
    AddScore 5000000
End Sub

Sub OOCScore3
    DMDScoreNow
    AddScore 3000000
End Sub

Sub OOCScore4
    DMDScoreNow
    AddScore 50000000
End Sub


Sub MeanderingMoonPoodle
      If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_MeanderingTenThous":End If
      vpmtimer.addtimer 3000, "DMDScoreNow'"
    AddScore 10000

End Sub

Sub StopOutOfControl
  l23E.State=0
End Sub

'*******************************
'Activate Time Warp Award Light
'*******************************
Sub ActivateGloryAwardLight
  GiOff
  vpmtimer.addtimer 1200, "GiOn'"
  LS_AwardReady.Play SeqBlinking,, 6,100
  DMD "", "", "DMD_CLAct", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 3000, "DMDScoreNow'"
  AddScore 0
  l36.State=2:l37a.State=2
  vpmTimer.AddTimer 40000, "StopGloryAwardLight'"
End Sub

Sub StopGloryAwardLight
  l36.State=0:l37a.State=0
End Sub



Sub ActivateAdvanceTime
  BG_Sequence 1,0,200
PlaySound "CO_AdvanceTimeActivated"
  AdvanceTimeActivated=1
  l36.State=2:l37a.State=2:l37b.State=2
debug.print "AdvanceTimeActivated"
    GiOff
  vpmtimer.addtimer 1200, "GiOn'"
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_AdvanceTimeActivated":End If
  DMD "", "", "DMD_ADvTimeAct", eNone, eNone, eNone, 3000, True, ""

  LS_AwardReady.Play SeqBlinking,, 8,100
  vpmtimer.addtimer 3500, "DMDScoreNow'"
  AddScore 0

If EasyMode=True Then:vpmtimer.addtimer 60000, "StopAdvanceTime'":End If
If HardMode=True Then:vpmtimer.addtimer 45000, "StopAdvanceTime'":End If
End Sub



Sub StopAdvanceTime
  l36.State=0:l37a.State=0:l37b.State=1:AdvanceTimeTimer.Enabled=False:AdvanceTimeActivated=False
End Sub
'***************
'Super Jets
'***************

Sub SuperJets                           '
  If l49.State=2 Then: Exit Sub
  BG_Sequence 1,0,200

  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  l49.State=2:l48.State=0
  DMD "", "", "DMD_SuperJets", eNone, eNone, eNone, 3000, True, ""
  LS_MajorAward.UpdateInterval = 10
  LS_MajorAward.Play SeqStripe2VertOn, 20, 2
  vpmtimer.addtimer 3200, "SuperJetsAct'"
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_SuperJetMillions":End If
  vpmtimer.addtimer 60000, "StopSuperJets'"
End Sub

Sub SuperJetsAct
  DMDScoreNow
  AddScore 0
End Sub

Sub StopSuperJets
  l49.State=0:l48.State=2
End Sub

'**********************
'Slay Dragon Initiate
'**********************

Sub SlayDragonInitiate
  If l4.State=2 Then: Exit Sub
  If l75.State=2 Then: Exit Sub
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  LS_MajorAward.Play SeqBlinking,, 8,100
  DMD "", "", "DMD_AttackTheDragon2", eNone, eNone, eNone, 3000, True, ""
  AddScore 10000000
  vpmtimer.addtimer 3200, "DMDScoreNow'"


  Hologram1Stop
  PlaySong "m_Unit"
  vpmTimer.AddTimer 3000, "AttackDragon'"
  vpmTimer.AddTimer 200, "TimeThiefInvisible'"
  vpmTimer.AddTimer 500, "QManVisible'"
  vpmtimer.addtimer 400, "Hologram2Start'"
End Sub

Sub AttackDragon
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_DefeatTheDragon":End If
  AddScore 1000000
  LightSword.State=2:l4.State=2: l1.State=2:l2.State=2:l3.State=2
  LiftBarricadeTargets

End Sub


Sub CalloutDragon

End Sub
'***********************
'Lost In Space Mystery
'***************************
Sub LostInSpace
  DMD "", CL(1, "LOST AGAIN QMAN"), "", eNone, eBlink, eNone, 3000, True, ""
' DMD "", "", "DMD_Blank", eNone, eNone, eNone, 200, False, "" 'Light

  vpmtimer.addtimer 3200, "DMDScoreNow'"
  AddScore 0
  TakeThePiss2
  l23A.State=0
  ' vpmTimer.AddTimer 5000, "TurnOffWheelLights'"                           'Turn off wheel light
End Sub


Dim MysteryCount
Sub Mystery
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  LS_MajorAward.Play SeqBlinking,, 8,100
  Select Case Int(Rnd*4)+1
    Case 1: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_MysteryLackOfSkillAward":End If
          DMD "", "", "DMD_Mystery", eNone, eNone, eNone, 3000, True, ""
    '     DMD "", "", "DMD_Blank", eNone, eNone, eNone, 200, False, ""
          vpmtimer.addtimer 3200, "DMDScoreNow'"
          AddScore 5000000

    Case 2: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_AnyIdea":End If
          DMD "", "", "DMD_Mystery", eNone, eNone, eNone, 3000, True, ""
'         vpmtimer.addtimer 2200, "DMDBlank'"
          vpmtimer.addtimer 3200, "DMDScoreNow'"
          AddScore 10000
          vpmTimer.AddTimer 3000, "CheckWormHoleJackpots'"

    Case 3: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_DownANotch":End If
          AddScore 1000000
          DMD "", "", "DMD_Mystery", eNone, eNone, eNone, 3000, True, ""
'         DMD "", "", "DMD_Blank", eNone, eNone, eNone, 200, False, ""
          vpmtimer.addtimer 3000, "DMDScoreNow'"

    Case 4: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_LuckPlan": End If
          AddScore 20000000
          MysteryCount = 0
          DMD "", "", "DMD_Mystery", eNone, eNone, eNone, 3000, True, ""
'         DMD "", "", "DMD_Blank", eNone, eNone, eNone, 200, False, ""
          vpmtimer.addtimer 3200, "DMDScoreNow'"
          AddScore 20000000
  End Select
End Sub


Sub StopLostInSpace
  l23F.State=0
End Sub

'**************************
'CheckWormHoleJackpots
'**************************
Sub CheckWormHoleJackpots
  If l24.State=2 Then:l23F.State=0:l23G.State=2
  If l24.State<2 Then
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
    LS_MajorAward.UpdateInterval = 10
    LS_MajorAward.Play SeqUpOn, 35, 2
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_WormholeJackPotsActivated":End If
    vpmTimer.AddTimer 2000, "WormLocksActivate'"
  End If
End Sub




'***************************************
'Special Awards ExtraBall,Advance Time
'***************************************

Sub Kicker8_hit()
  PlaySoundAt "fx_kicker_enter", Kicker8
  vpmtimer.addtimer 500, "CheckAwards'"
  vpmtimer.addtimer 1500, "Kicker8Kick'"
  l37b.State=2
End Sub

Sub Kicker8Kick
  PlaySoundAt "fx_kicker", Kicker8
  Kicker8.kick 215, 18
  DOF 117,DOFPulse
  l37a.State=0:l37b.State=1
End Sub


Sub CheckAwards
  If l37a.State=0 Then TakeThePiss2: End If
  If ExtraBallActivated=1 Then
    AwardExtraBall
    Objlevel(2) = 1
    FlasherFlash2_Timer
    GiOff
    vpmtimer.addtimer 3200, "GiOn'"
End If
  If AdvanceTimeActivated=1 Then
    AdvanceTime:
    Objlevel(2) = 1
    FlasherFlash2_Timer
    GiOff
    vpmtimer.addtimer 3200, "GiOn'"
End If

End Sub

Sub CheckIfGloryAvailable
' If GloryTableTimer.Enabled=1 Then:Exit Sub
  StartGloryBillions
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  TurnOffModes
  StopGloryAwardLight
End Sub



'******************************
'AdvanceTime Awards and Scenes
'******************************
Dim TimeComplete


Sub AdvanceTime
  BG_Sequence 5,10,0
  AdvanceTimeActivated=0
  l36.State=0
  debug.print "AdvanceTimeActive Timer Enabled"
  LS_MajorAward.Play SeqBlinking,, 8,100
  LightSeqGi.Play SeqBlinking,, 8,100
  StartFlasherAward

  If DragonWounded4(CurrentPlayer)=True And bExtraBallWonThisBall=False Then AwardExtraBall:bExtraBallWonThisBall=False:Exit Sub
  If DragonWounded3(CurrentPlayer)=True Then l45.State=1:vpmtimer.addtimer 1200,"CalloutTimeComplete'":vpmtimer.addtimer 1300, "Scene4'":StartPlanetDragon1:Exit Sub
  If DragonWounded2(CurrentPlayer)=True Then l44.State=1:vpmtimer.addtimer 1200,"CalloutTimeAward'":vpmtimer.addtimer 1300, "Scene3'":Exit Sub
  If DragonWounded1(CurrentPlayer)=True Then  l43.State=1:vpmtimer.addtimer 1200,"CalloutTimeAward'":vpmtimer.addtimer 1300, "Scene2'":Exit Sub
  If DragonWounded1(CurrentPlayer)=False Then  l42.State=1:vpmtimer.addtimer 1200,"CalloutTimeAward'":vpmtimer.addtimer 1300, "Scene1'":Exit Sub
End sub

Sub Scene1
  DragonWounded1(CurrentPlayer)=True
  debug.print "DragonWounded1=True"
  StartFlasherAward
  PlaySong "m_Westminster Bridge"
  DMD "", "", "DMD_DragonWounded", eNone, eNone, eNone, 3000, True, ""
' DMD "", "", "DMD_Blank", eNone, eNone, eNone, 200, False, ""
  vpmtimer.addtimer 3200, "DMDScoreNow'"
  AddScore 500000000
  TimeThief.Visible=1
  l_HourGlass.State=2
  l_HourGlass2.State=2
  Light_Catch2.State=2
  QMan.Visible=0
  l_Staff1.State=0
  l_Staff2.State=0
  Hologram1Start
End Sub

Sub Scene2
  DragonWounded2(CurrentPlayer)=True
  debug.print "DragonWounded2=True"
  StartFlasherAward
  PlaySong "m_Slitheen"
  DMD "", "", "DMD_DragonWounded", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 3200, "DMDScoreNow'"
  AddScore 500000000
  TimeThief.Visible=1
  l_HourGlass.State=2
  l_HourGlass2.State=2
  Light_Catch2.State=2
  QMan.Visible=0
  l_Staff1.State=0
  l_Staff2.State=0
  Hologram1Start
End Sub

Sub Scene3
  DragonWounded3(CurrentPlayer)=True
  debug.print "DragonWounded3=True"
  StartFlasherAward
  debug.print "AdvanceTime3"
  PlaySong "m_Westminster Bridge"
  DMD "", "", "DMD_DragonWounded", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 3200, "DMDScoreNow'"
  AddScore 500000000
  TimeThief.Visible=1
  l_HourGlass.State=2
  l_HourGlass2.State=2
  Light_Catch2.State=2
  QMan.Visible=0
  l_Staff1.State=0
  l_Staff2.State=0
  Hologram1Start
End Sub

Sub Scene4
  DragonWounded4(CurrentPlayer)=True
  debug.print "DragonWounded4(CurrentPlayer)=True"
  StartFlasherAward
  debug.print "AdvanceTime4"
  TimeComplete=true
  PlaySong "m_Slitheen"
  DMD "", "", "DMD_DragonDefeated", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 3200, "DMDScoreNow'"
  AddScore 500000000
  QMan.Visible=1
  l_Staff1.State=2
  l_Staff2.State=2
  TimeThief.Visible=0
  l_HourGlass.State=0
  l_HourGlass2.State=0
  Light_Catch2.State=0
  Hologram2Start
  CheckTimeLordStatusTimer.Enabled=1
End Sub


'***********************************
'DudCount2-Callouts for OutofControl
'************************************


Dim DudCount2
Sub TakeThePiss2
  BG_Sequence 37,0,110

  DudCount2 = DudCount2 + 1
  Select Case DudCount2
    Case 1: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_LadyFriends":End If
    Case 2: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_TractorOil":End If
    Case 3: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_LadyFriends" :End If
    Case 4: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_TractorOil":End If
           DudCount2=0
  End Select
End Sub

'**************************
'Activate Worm Locks
'**************************


Sub WormLocksActivate
  If WormLocksActivated(CurrentPlayer)=True Then:Exit Sub
  BG_Sequence 1,0,200
  SpiderMoveUp
  WormLocksActivated(CurrentPlayer)=True

  DMD "", "", "DMD_WormHolesLit", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 3200, "DMDScoreNow'"


  WormKick1.Enabled=1:WormKick2.Enabled=1:WormKick3.Enabled=1
  l_wl1.State=2:l_wl2.State=2:l_wl3.State=2:l24.State=2
  l_w1b.State=2:l_w2b.State=2:l_w3b.State=2
  WormCompleteTimer.Enabled=1
End Sub

Sub WormCompleteTimer_Timer
  vpmTimer.AddTimer 60000, "DisableWormLocks'"
End Sub

Sub DisableWormLocks
  WormKick1.Enabled=0:WormKick2.Enabled=0:WormKick3.Enabled=0
  l_wl1.State=0:l_wl2.State=0:l_wl3.State=0:l24.State=0
  l_w1b.State=0:l_w2b.State=0:l_w3b.State=0
  WormCompleteTimer.Enabled=0
  WormLocksActivated(CurrentPlayer)=False
  GateWormLocks.open=True
  vpmtimer.addtimer 5000, "SpiderMoveDown'"
End sub



'*******************************************
'WormLaneKickers-StartPlanetBird & R in LORD
'*******************************************


'WormHole KickerJump
Sub WormKick1_hit()
  GateWormLocks.open = False
  PlaySoundAt "VUKEnter", WormKick1
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_Worms":End If
  DMD "", "", "DMD_Worms", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 3200, "DMDScoreNow'"
  vpmtimer.addtimer 3800, "PlaySoundAt""Popper"", WormKick1: WormKick1.kick 315, 10'"

End Sub


Sub WormKick2_hit()
  PlaySoundAt "VUKEnter", WormKick2
  CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_OilMyLinkages"
  DMD "", "", "DMD_GreaseMeUp", eNone, eNone, eNone, 3000, True, ""
  vpmtimer.addtimer 3200, "DMDScoreNow'"
  vpmtimer.addtimer 3800, "PlaySoundAt""Popper"", WormKick2: WormKick2.kick 310, 10'"
End Sub

Sub WormKick3_hit()
  PlaySoundAt "VUKEnter", WormKick3
  '*****************************************
  l40.State=1                   '-----------------------------------------------------------------Light R in LORD & Check Timelord Complete
  CheckTimeLordStatusTimer.Enabled=1
  '******************************************
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  If WormLocksActivated(CurrentPlayer)=True Then
      If Hard=0 Then: AwardSuperJackpot:End If
      If Hard=1 Then: AwardSuperJackpot2:End If
  End If

  If WormLocksComplete(CurrentPlayer)=False Then
    vpmTimer.AddTimer 2000, "StartPlanetSpider'"
    debug.print "WormLocksComplete(CurrentPlayer)=True"
    WormLocksComplete(CurrentPlayer)=True
  End If

    vpmtimer.addtimer 3800, "PlaySoundAt""Popper"", WormKick3: WormKick3.kick 310, 10'"
    DisableWormLocks
End Sub

Sub WormKick4_hit()
  PlaySoundAt "VUKEnter", WormKick4
  vpmtimer.addtimer 500, "KickerWormKick4Kick'"
  GateWormLocks.open = True
  AddScore 50000
End Sub

Sub KickerWormKick4Kick
  PlaySoundAt"Popper",WormKick4
  WormKick4.kick   0, 60, 1.56
  DOF 116,DOFPulse
End Sub

Sub BlinkWormLights
  l28B.State=2:l28D.State=2:l28E.State=2:l28G.State=2:l28H.State=2
End Sub

Sub StopWormLights
  l28B.State=0:l28D.State=0:l28E.State=0:l28G.State=0:l28H.State=0
End Sub

Sub Trigger15_Hit()
  PlaySound "Crystal_intro_3sec"
End Sub


'*********************************************
'WormHoleTrigger 2Trigger- No longer used
'*******************************************
Sub Trigger2_Hit()
  ' If  LastSwitchHit = "Trigger7" Then:  Exit Sub

End Sub

'*************
'  KICKBACKS
'*************
'

'kickbackl(CurrentPlayer), kickbackr(CurrentPlayer)

Sub closekickbacks
  kickbacklg.open = False
  LightLeftescape.State = 0
  kickbackrg.open = False
  Lightrightescape.State = 0
  LightLeftInlane.State = 2
  LightRightInlane.State = 2
End Sub

Sub kickbackleftenabled
  kickbacklg.open = True
  LightLeftInlane.State = 0
  LightLeftescape.State = 2
  PlaySoundAt "Kickback2",kickbacklg
End Sub

Sub kickbackleftdisabled
  kickbacklg.open = False
  LightLeftescape.State = 0
  LightLeftInlane.State = 2
  LightLeftescape.State = 0
End Sub

Sub kickbackrightenabled
  kickbackrg.open = True
  Lightrightescape.State = 2
  LightRightInlane.State = 0
  PlaySoundAt "Kickback2",kickbackrg
End Sub

Sub kickbackrightdisabled
  kickbackrg.open = False
  Lightrightescape.State = 0
  LightRightInlane.State = 2
End Sub

Sub Kicker11_hit
  PlaySoundAt SoundFXDOF("VUKEnter", 118, DOFPulse, DOFContactors), Kicker11

  Ls_KickBack.Play SeqBlinking,, 6,100
  vpmtimer.addtimer 1500, "LeftKickBack'"
End Sub

Sub LeftKickBack
  PlaySoundAt SoundFXDOF("fx_kicker", 118, DOFPulse, DOFContactors), Kicker11
  Kicker11.Kick 0, 35
  RightOutLaneSaveCallout :AddScore 200000
  vpmtimer.addtimer 300, "kickbackleftdisabled'"
  LastSwitchHit = "Kicker11"
End Sub

Sub Kicker12_hit
  PlaySoundAt SoundFXDOF("VUKEnter", 114, DOFPulse, DOFContactors), Kicker12
  Ls_KickBack.Play SeqBlinking,, 6,100
  vpmtimer.addtimer 1500, "RightKickBack'"
End Sub

Sub RightKickBack
  PlaySoundAt SoundFXDOF("fx_kicker", 114, DOFPulse, DOFContactors), Kicker12
  Kicker12.Kick 0, 35
  RightOutLaneSaveCallout:AddScore 200000
  vpmtimer.addtimer 300, "kickbackrightdisabled'"
  LastSwitchHit = "Kicker12"
End Sub

'***********************************
'Lanes
'*************************************
Sub Lane1Trigger2_Hit()
  LastSwitchHit = "Lane1Trigger2"
End Sub

Sub Lane4Trigger2_Hit()
  LastSwitchHit = "Lane4Trigger2"
End Sub

Sub lane1_hit
  If  LastSwitchHit = "Kicker11" Then:Exit Sub
  If LightLeftEscape.State=0 Then: LeftOutLaneCallout : End If
  If Tilted Then Exit Sub
  BG_Sequence 34,0,200
  LastSwitchHit = "lane1"
' LightLeftEscape.State=0
' LightLeftInlane.State = 0
End Sub

Dim LLCO
Sub LeftOutLaneCallout
  LLCO=LLCO+1
  Select Case LLCO
    Case 1 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_LittleCry":End If
    Case 2 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_MethusalasMother":End If
    Case 3 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_WillaWollaBingBang"  :End If
    Case 4 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_ShitCreek":End If
        LLCO=0
  End Select
End Sub


Sub lane2_hit
  AddScore 50000
  If Tilted Then Exit Sub
  BG_Sequence 34,0,100    'pri 6.NOSTOP   LEFT blinks following blink pattern on CLightbg
  LastSwitchHit = "lane2"
  If  LightLeftInlane.State = 0 Then: Exit Sub
  Ls_LaneLights.Play SeqBlinking,, 4,100
  LightLeftInlane.State = 0
  LightLeftEscape.State=2
  kickbackleftenabled
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_OutlaneReady":End If

End Sub

Sub lane3_hit
  AddScore 50000
  If Tilted Then Exit Sub
  BG_Sequence 35,0,100    'pri 6.NOSTOP   rIGHT blinks following blink pattern on CLightbg
  LastSwitchHit = "lane3"
  If  LightRightInlane.State = 0 Then: Exit Sub
  LightRightInlane.State = 0
  Lightrightescape.State=2
  Ls_LaneLights.Play SeqBlinking,, 4,100
  kickbackrightenabled
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_OutlaneReady":End If
End Sub


Sub lane4_hit
  If  LastSwitchHit = "Kicker12" Then:Exit Sub
  If Lightrightescape.State=0 Then: LeftOutLaneCallout: AddScore 200000:  End If
  If Tilted Then Exit Sub
  BG_Sequence 35,0,200

  LastSwitchHit = "lane4"
  '   checkforkickbacks
' Lightrightescape.State=0
' LightRightInlane.State = 0
End Sub

Dim RLCO
Sub RightOutLaneCallout
  RLCO=RLCO+1
  Select Case RLCO
    Case 1 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_LittleCry":End If
    Case 2 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_WillaWollaBingBang":End If
    Case 3 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_MethusalasMother":End If
    Case 4 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_ShitCreek":End If
           RLCO=0
  End Select
End Sub

Dim OutLaneSave
Sub RightOutLaneSaveCallout
  OutLaneSave=OutLaneSave+1
  Select Case OutLaneSave
    Case 1 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_LuckyDay":End If
    Case 2 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_SpaceShip":End If
    Case 3 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_ReturnToSender":End If
    Case 4 :If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_Perfectly":End If
           OutLaneSave=0
  End Select
End Sub

'***********************'Triggers for Sounds
'***********************


Sub Trigger7_Hit()
  LastSwitchHit = "Trigger7"
End Sub




Sub Trigger11_Hit()
  LastSwitchHit = "Trigger11"
End Sub

'Sub Trigger12_Hit()
' LastSwitchHit = "Trigger12"
'End Sub

'Sub Trigger13_Hit()
' LastSwitchHit = "Trigger13"
'End Sub


Sub Trigger5_Hit()
  LastSwitchHit = "Trigger5"
End Sub


Sub TriggerWarpedHiddenRamp_Hit()
End Sub

'*******************************
'LeftTargets
'********************************

Sub sw3_Hit
  If l11.State=0 Then: Exit Sub
  If l11.State=2 Then: l11.State=1:Objlevel(1) = 1:FlasherFlash1_Timer
  CheckLeftTargetsComplete
  AddScore 1000000
End Sub

Sub sw4_Hit
  If l11.State=0 Then: Exit Sub
  If l12.State=2 Then: l12.state=1:Objlevel(1) = 1:FlasherFlash1_Timer
  CheckLeftTargetsComplete
  AddScore 1000000
End Sub

Sub sw5_Hit
  If l11.State=0 Then: Exit Sub
  If l13.State=2 Then: l13.State=1:Objlevel(1) = 1:FlasherFlash1_Timer
  CheckLeftTargetsComplete
  AddScore 1000000
End Sub

Sub CheckLeftTargetsComplete
  If l11.State + l12.State  + l13.State = 3 Then
    vpmtimer.addtimer 200, "LeftTargetsComplete'"
  End If
End Sub

'****************************
'LeftTargetsCompleteCheck
'****************************
Dim LTC_Count
Sub LeftTargetsComplete
Light_LeftLaneRedHand.State=2
  addScore 500000
  l11.State=0: l12.State=0 :l13.State=0
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_OuterGalaxyActivated":End If
  PlaySong "m_Westminster Bridge"
  If Hard=0 Then: StartMeteors:End If
  CheckMeteorSequence
End Sub

Sub CheckMeteorSequence
  If l17B.state=1 Then:l16.State=1:l17.State=1:l18.State=2:ChaosActivated:BG_Sequence 1,0,200:Exit Sub
  If l17A.state=1 Then:l16.State=1:l17.State=2:l18.State=0:Exit Sub
  If l16A.state=1 Then:l16.State=2:l17.State=0:l18.State=0:Exit Sub
End Sub
'******************************************************
'MeteorShowerTrigger- L3eft lane Multiball and Kicker
'*******************************************************

Sub Trigger1A_Hit()
  If LastSwitchHit = "Trigger1B" Then:LastSwitchHit = "Trigger1A":  Exit Sub
  PlaySound "TriggerWarp"
  LastSwitchHit = "Trigger1A"
  If lGlory1.State=2 Then: Exit Sub
  If l18.State=1 Then: Exit Sub
  If l18.State=2 Then:AddScore 350000000:DMDScoreNow:MeteorShowerCallout:vpmtimer.addtimer 3200, "StartMeteorShower'":l18.State=1: End If
  BG_Sequence 1,0,20
'Fixing  all lights off exept for clock ?


  If l17.State=2 Then
    l17.State=1:l18.State=2
    If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_SpeedOfLight":End If
    StartMeteors
    AddScore 50000000
    DMDScoreNow
  End If

  If l16.State=2 Then
    l16.State=1:l17.State=2
    If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_SpeedOfLIghtActivated":End If
    AddScore 50000000
    DMDScoreNow
  End If
End Sub

Sub StartMeteors
  l18A.State=2:l18B.State=2:l18C.State=2:l18D.State=2
End Sub

Sub MeteorShowerCallout
  DMD "", "", "DMD_ChaosCrit", eNone, eNone, eNone, 4000, False, "" 'Light
LightSeqGi.Play SeqBlinking,, 6,100
  vpmtimer.addtimer 4500, "DMDScoreNow'"
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_SituationCritical":End If

End Sub
Dim MeteorShowerStarted
Dim MeteorShowerInPlay

Sub StartMeteorShower

Light_LeftLaneRedHand.State=1
  BG_Sequence 5,10,0
  BG_addshow = 1 ' BG_Sequence 1,0,200
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"

  EnableBallSaver 25
  MeteorShowerInPlay=True
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_MeteorShower" :End If
  l18A.State=2:l18B.State=2:l18C.State=2:l18D.State=2
  bAutoPlunger=True

  If ChaosCriticalComplete(CurrentPlayer)=False Then
    ChaosCriticalComplete(CurrentPlayer)=True
    StartPlanetChaos
    l38.State=1   '----------------------------------------------------------L in LORD
    debug.print "ChaosCriticalComplete(CurrentPlayer)=True"
    CheckTimeLordStatusTimer.Enabled=1
  End If

  MeteorShowerLights
  bBallSaverReady = True  'Activate Ball Saver for all Balls
  If bMultiBallMode=1 Then EnableBallSaver 30:AddMultiball (2)
  If bMultiBallMode=0 Then EnableBallSaver 30:AddMultiball (2)
  vpmtimer.addtimer 2500, "WormLocksActivate'"
  MeteorShowerStarted=True
  vpmtimer.addtimer 6000, "EndMeteorShower'"
End Sub

Sub ChaosActivated
  LS_ChaosActivated.Play SeqBlinking,, 1,50
End Sub

Sub MeteorShowerLights
  LS_MajorAward.Play SeqBlinking,, 8,100
StartFlasherAward
End Sub

Sub Trigger1B_Hit()
  LastSwitchHit = "Trigger1B"
End Sub

'****************************
'L-In-LORD-Light& PlanetChaos
'***************************
Sub Trigger1C_Hit()
  LastSwitchHit = "Trigger1C"
End Sub

Sub EndMeteorShower
  l18A.State=0:l18B.State=0:l18C.State=0:l18D.State=0
  l11.State=2:l12.State=2:l13.State=2
  If l17B.State=1 Then:l17B.State=0:l17A.State=1:l16A.State=0:l16.State=0:l17.State=0:l18.State=0:Exit Sub
  If l17A.State=1 Then:l17A.State=0:l16A.State=1:l17B.State=0:l16.State=0:l17.State=0:l18.State=0:Exit Sub
  If l16A.State=1 Then:l16A.State=0:l17B.State=1:l17A.State=0:l16.State=0:l17.State=0:l18.State=0:Exit Sub
  MeteorShowerStarted=False
End Sub



Sub StopMeteorShowerLights
  l18A.State=0:l18B.State=0:l18C.State=0:l18D.State=0
  l16.State=0:l17.State=0:l18.State=0
End Sub

Sub ResetLeftTargetBankLights
  l11.State=2:l12.State=2:l13.State=2
End Sub
'****************************
'MeteorShowerTop LaneKicker
'***************************
Sub Kicker7_hit()
  debug.print "Kicker7Hit"
  WallJamStopper.isdropped=false
  PlaySoundAt "fx_kicker_enter", Kicker7
  vpmtimer.addtimer 100, "CloseClogBlockerGate'"
  debug.print "CloseClogBlockGate"
  vpmtimer.addtimer 2000, "Kicker7Kick'"
End Sub

Sub Kicker7_unhit()
  vpmtimer.addtimer 2200, "OpenClogBlockerGate'"
  PlaySoundAt "fx_kicker",Kicker7
  debug.print "KickerUnhit"
  DOF 115,DOFPulse
End Sub

Sub Kicker7Kick
  Kicker7.Kick 315, 25
  debug.print "Kicker7Kick"
End Sub

Sub CloseClogBlockerGate
  GateClogBlocker.open = False
  debug.print "GateClogBlockOpen"
End Sub

Sub OpenClogBlockerGate
  GateClogBlocker.open = True
  WallJamStopper.isdropped=True
  debug.print "GateClogBlockOpen"
End Sub


'***********
'ReverseTime
'***********
Sub Trigger3_Hit()
  Objlevel(3) = 1:FlasherFlash3_Timer
If LightBallLock2.State=2 Then:Exit Sub 'It is possible to hit another ball up the ramp while the ball has just been trapped
  If lGlory1.State=2 Then AwardSuperJackpot2
  If TWJPFiring=True Then: AwardJackpot:Exit Sub
  If l21B.State=1 Then: ReverseLock2Shot:Exit Sub
  If l20A.State=1 Then: ReverseLock3Shot: Exit Sub
  If l19A.State=1 Then: ReverseLock4Shot: Exit Sub
End Sub


Sub ReverseLock4Shot
  If lGlory1.State=2 Then: Exit Sub
  If l61.State=2 Then: Exit Sub
  If l21.State=2 Then l21.State=1:l21A.State=2:RaiseRampL:CheckBallLocks: AddScore 3000000: End if
  If l20.State=2 Then
      l20.State=1:l21.State=2
      PlaySound "CO_RaiseLeftRamp"
      AddScore 2000000
  End if

  If l19.State=2 Then
      l19.State=1:l20.State=2
      PlaySound "CO_6thSense"
      AddScore 1000000
  End If
  LastSwitchHit = "Trigger3"
End Sub

Sub ReverseLock3Shot
  If lGlory1.State=2 Then: Exit Sub
  If l61.State=2 Then: Exit Sub
  If l21.State=2 Then l21.State=1:l21A.State=2:RaiseRampL:CheckBallLocks:AddScore 3000000:End if
  If l20.State=2 Then
      l20.State=1:l21.State=2
      PlaySound "CO_RaiseLeftRamp"
      AddScore 1000000
  End if
  LastSwitchHit = "Trigger3"
End Sub


Sub ReverseLock2Shot
  If lGlory1.State=2 Then: Exit Sub
  If l61.State=2 Then: Exit Sub
  If l21.State=2 Then:l21.State=1:l21A.State=2:AddScore 1000000:RaiseRampL:CheckBallLocks:  End if
  LastSwitchHit = "Trigger3"
End Sub


Sub StopReverseTime
  l19.State=0:l20.State=0:l21.State=0:l21A.State=0:LightBallLock1.State=0:LightBallLock2.State=0
End Sub

Sub CheckBallLocks
  If LightBallLock1.State=1 Then
    LightBallLock2.State=2:DisableBallLock1:EnbleBallLock2
    DMD "", "", "DMD_Lock2", eNone, eNone, eNone, 2000, True, "CO_Lock2Ready"
    vpmtimer.addtimer 3500, "Lock1Score'"
  End If
  If LightBallLock1.State=0 Then
    LightBallLock1.State=2:DisableBallLock2:EnbleBallLock1
    DMD "", "", "DMD_Lock1", eNone, eNone, eNone, 3000, True, "CO_Lock1Ready"
    vpmtimer.addtimer 3500, "Lock1Score'"
    AddScore 0
    End If
End Sub

Sub Lock1Score
    DMDScoreNow
    AddScore 0
End Sub

Sub EnbleBallLock1
  KickerReverseLock1.Enabled=1
  ReverseLock1Ready(CurrentPlayer)=True
  debug.print "ReverseLock1 enabled"
End Sub

Sub DisableBallLock1
  KickerReverseLock1.Enabled=0
  ReverseLock1Ready(CurrentPlayer)=False
  debug.print "ReverseLock1 disabled"
End Sub

Sub EnbleBallLock2
  KickerReverseLock2.Enabled=1
  ReverseLock2Ready(CurrentPlayer)=True
  debug.print "ReverseLock2Enabled"
End Sub

Sub DisableBallLock2
  KickerReverseLock2.Enabled=0
  ReverseLock2Ready(CurrentPlayer)=False
  debug.print "ReverseLock2 disabled"
End Sub

Sub ResetReverseLockLights
  If  l21B.State=1 Then:l21.State=2:l20.State=1:l19.State=1:Light_ReverseClock.State=1:LightBallLock1.State=0:LightBallLock2.State=0:Exit Sub
  If  l20A.State=1 Then:l20.State=2:l21.State=0:l19.State=1:Light_ReverseClock.State=1:LightBallLock1.State=0:LightBallLock2.State=0:Exit Sub
  If  l19A.State=1 Then:l19.State=2:l20.State=0:l21.State=0:Light_ReverseClock.State=1:LightBallLock1.State=0:LightBallLock2.State=0:Exit Sub

  If l19A.State=0 and l20A.State=0 and l21B.State=0 and EndOfGloryTable=True Then
    l19A.State=1:l21.State=2:l21A.State=0:LightBallLock1.State=0:LightBallLock1.State=0:LightBallLock2.State=0:Light_ReverseClock.State=1
    TWJPFiring=False
    EndOfGloryTable=False
    ReverseLock1Ready(CurrentPlayer)=False
    ReverseLock2Ready(CurrentPlayer)=False
  End If

End Sub


Sub StartReverseLockLights
  If EasyMode=True Then:ReverseEasyLightChange: End If
  If HardMode=True Then:ReverseHardLightChange: End If
End Sub


Sub TurnOffReverseTimeLights
  l19.State=0:l20.State=0:l21.State=0:l21A.State=0
End Sub

Sub TurnOffReverseLockLights
  l21A.State=0: LightClockBroken.State=0
End Sub

Dim ReverseReleaseActive
Sub KickerReverseLock1_hit()
  BG_Sequence 38,0,200' nr 4 blue wizard binking
  BG_Sequence 2,0,200 'clockspin ccw 200updates
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  If BallCatchInPlay=True Then ReverseLock1Escape
  If BallCatchInPlay=False Then ReverseLock1Release
End Sub

Sub ReverseLock1Escape  ' Reverse Lock1 Escape if TimeLord Catch is in play

  debug.print "Reverselock1Escape"
  KickerReverseLock1.DestroyBall
  KickerReverseRelease.Enabled=1
  KickerReverseRelease.CreateSizedball BallSize / 2
  PlaySoundAt"fx_kicker", KickerReverseRelease
  KickerReverseRelease.kick 5,110
  KickerReverseRelease.Enabled=0
  debug.print "ReverseLockRelease Disabled"
  StartFlasherAward
  PlaySoundAt "VUKEnter", KickerReverseLock1
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_WillaWollaBingBang":End If
  LowerRampL
  ReverseLock1Ready(CurrentPlayer)=False
  LightClockBroken.State=1
  LightBallLock1.State=1
  Lock1Display
  LightBallLock1.State=1
  If l21B.State=1 Then :l21B.State=1: l20A.State=0:l19A.State=0:l19.State=1:l20.State=1:l21.State=2 :l21A.State=0:Exit Sub
  If l20A.State=1 Then :l20A.State=1: l19A.State=0:l21B.State=0:l20.State=2 :l21.State=0:l19.State=1:l21A.State=0:Exit Sub
  If l19A.State=1 Then :l19A.State=1:l20A.State=0: l21B.State=0:l19.State=2:l20.State=0:l21.State=0:l21A.State=0: :Exit Sub
End Sub



Sub ReverseLock1Release
  AddScore 400000000
  DMDScoreNow
  ReverseReleaseActive=True:  debug.print "ReverseReleaseActive=True"
  StartFlasherAward
  WireRampOff
  PlaySoundAt "VUKEnter", KickerReverseLock1
  PlaySound "CO_YouMustReverseTime"
  LowerRampL
  ReverseLock1Ready(CurrentPlayer)=False
  GiOff
  LS_MajorAward.Play SeqBlinking,, 8,85
  vpmtimer.addtimer 4500, "DestroyLockball1'"
''  vpmtimer.addtimer 2000, "Lock1Display'"
  Lock1Display
  LightBallLock1.State=1
  If l21B.State=1 Then :l21B.State=1: l20A.State=0:l19A.State=0:l19.State=1:l20.State=1:l21.State=2 :l21A.State=0:Exit Sub
  If l20A.State=1 Then :l20A.State=1: l19A.State=0:l21B.State=0:l20.State=2 :l21.State=0:l19.State=1:l21A.State=0:Exit Sub
  If l19A.State=1 Then :l19A.State=1:l20A.State=0: l21B.State=0:l19.State=2:l20.State=0:l21.State=0:l21A.State=0: :Exit Sub
End Sub


Sub DestroyLockball1
  KickerReverseLock1.DestroyBall
  PlaySound "fx_woosh"
  LightBallLock1.State=1
  EscapeReverseLock1Exit

  KickerReverseTunnel.Enabled=1
  ' l19.State=2
WallReverseOutProtect.isdropped=False
  CreateNewBallKickerReverseTunnel
  vpmtimer.addtimer 1500, "EscapePortalExit '"
End Sub

Sub Lock1Display
    DMD "", "", "DMD_Ball1Locked1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball1Locked1b", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball1Locked1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball1Locked1b", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball1Locked1a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball1Locked1b", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball1Locked1c", eNone, eNone, eNone, 2500, True, ""
  vpmtimer.addtimer 3500, "DMDScoreNow'"
  AddScore 5000000
End Sub


Sub EscapeReverseLock1Exit()
  FlashForMs f3b1, 500, 50, 0
  FlashForMs f3b2, 500, 50, 0
  FlashForMs f3b3, 500, 50, 0
  FlashForMs f3b4, 500, 50, 0
  FlashForMs f3d1, 1000, 50, 0
  FlashForMs f3d2, 1000, 50, 0
  FlashForMs f3d3, 1000, 50, 0
  FlashForMs f3d4, 1000, 50, 0
  DOF 121, DOFPulse
End Sub

'*************************
'Lock2
'*************************
'Dim Lock2Hit
Sub KickerReverseLock2_hit()
  BG_Sequence 2,0,200 'clock spin counterclockwise   200 updates  SPECIAL

  BG_Sequence 38,0,200' nr 4 blue wizard binking
  BG_Sequence 2,0,200 'clockspin ccw 200updates
  GiOff
  vpmtimer.addtimer 3200, "GiOn'"
  If BallCatchInPlay=True Then ReverseLock2Escape
  If BallCatchInPlay=False Then ReverseLock2Release
End Sub


Sub ReverseLock2Escape        'This release is if lock  2 is made and the ball catch release is already in play
  debug.print "Reverselock2Escape Enabled"
  KickerReverseLock2.DestroyBall
  KickerReverseRelease.Enabled=1
  KickerReverseRelease.CreateSizedball BallSize / 2
  PlaySoundAt"fx_kicker", KickerReverseRelease
  KickerReverseRelease.kick 5,110
  KickerReverseRelease.Enabled=0
  debug.print "Reverselock2Escape Disabled"
  l39.State=1

  StartFlasherAward
  PlaySoundAt "VUKEnter", KickerReverseLock2
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_BreakEverthing2":End If
  LowerRampL
  ReverseLock2Ready(CurrentPlayer)=False
  LightClockBroken.State=2
  LightBallLock2.State=1
  StopClockHandsClockwise
  vpmtimer.addtimer 500, "StartClockHandsCounterClockwise'"
  vpmtimer.addtimer 600, "Lock2Display'"
  StartPlanetReverseTime
  If EasyMode=True Then:ReverseEasyLightChange: End If
  If HardMode=True Then:ReverseHardLightChange: End If
  PlaySound "Thunder"
  Hologram1Stop
  If ReverseTimeComplete(CurrentPlayer)=False Then
    l39.State=1   '------------------------------------------------------Start 'O' light in LORD
    ReverseTimeComplete(CurrentPlayer)=True
    debug.print "ReverseTimeComplete(CurrentPlayer)=True"
    CheckTimeLordStatusTimer.Enabled=1
  End If
  vpmtimer.addtimer 3000, "KickerReverseTunnelKick'"
  vpmtimer.addtimer 500, "Hologram1Start'"
  vpmtimer.addtimer 500, "TimeThiefInvisible'"
  vpmtimer.addtimer 600, "QManVisible'"
  vpmtimer.addtimer 3200, "CheckEnableTimeWarpJackpots'"
End Sub



Sub ReverseLock2Release
  AddScore 400000000
  DMDScoreNow
  ReverseReleaseActive=True:  debug.print "ReverseReleaseActive=True"
  l39.State=1
  CheckTIMELORDStatus
  WireRampOff
  StartFlasherAward
  PlaySoundAt "VUKEnter", KickerReverseLock2
  PlaySound "CO_BreakEverthing2"
  LowerRampL
  ReverseLock2Ready(CurrentPlayer)=False
  vpmtimer.addtimer 3000, "DestroyLockball2'"
  ' l_HourGlass.State=2
  LightClockBroken.State=2
  LightBallLock2.State=1
  StopClockHandsClockwise
  vpmtimer.addtimer 500, "StartClockHandsCounterClockwise'"
  vpmtimer.addtimer 600, "Lock2Display'"
  StartPlanetReverseTime
  If EasyMode=True Then:ReverseEasyLightChange: End If
  If HardMode=True Then:ReverseHardLightChange: End If
  LightSeqGi.Play SeqBlinking,, 10,100
End Sub

Sub ReverseHardLightChange
  l19A.State=1: l21B.State=0:l21.State=0:l20A.State=0:l19.State=2:l20.State=0:l21A.State=0
End Sub

Sub ReverseEasyLightChange
debug.print "Reverse Easy Sub"
  If l21B.State=1 Then :l21B.State=0: l20A.State=1:l19A.State=0 :l19.State=1:l20.State=2:l21.State=0: l21A.State=0:Exit Sub
  If l20A.State=1 Then :l20A.State=0: l19A.State=1:l21B.State=0 :l19.State=2:l20.State=0:l21.State=0:l21A.State=0:Exit Sub
  If l19A.State=1 Then :l19A.State=0: l21B.State=1:l21.State=2:l20A.State=0:l19.State=1:l20.State=1:l21A.State=0 :Exit Sub
End Sub

Sub Lock2Display
    DMD "", "", "DMD_Ball2Locked2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball2Locked2b", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball2Locked2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball2Locked2b", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball2Locked2a", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball2Locked2b", eNone, eNone, eNone, 250, True, ""
    DMD "", "", "DMD_Ball2Locked2c", eNone, eNone, eNone, 2500, True, ""
    DMD "", "", "DMD_ReverseTime", eNone, eNone, eNone, 3000, True, ""


''  vpmtimer.addtimer 4200, "Lock2Score'"

End Sub

Sub Lock2Score
  DMDSCoreNow
  AddScore 5000000
End Sub

Sub DestroyLockball2
  KickerReverseLock2.DestroyBall
WallReverseOutProtect.IsDropped=False
  PlaySound "fx_woosh"
  ' LightBallLock2.State=1
  EscapeReverseLock2Exit
  KickerReverseTunnel.Enabled=1
  CreateNewBallKickerReverseTunnel2
End Sub

Sub EscapeReverseLock2Exit()
  FlashForMs f3c1, 500, 50, 0
  FlashForMs f3c2, 500, 50, 0
  FlashForMs f3c3, 500, 50, 0
  FlashForMs f3c4, 500, 50, 0
  FlashForMs f3d1, 1500, 50, 0
  FlashForMs f3d2, 1500, 50, 0
  FlashForMs f3d3, 1500, 50, 0
  FlashForMs f3d4, 1500, 50, 0
  DOF 121, DOFPulse
End Sub

Sub CreateNewBallKickerReverseTunnel

  KickerReverseTunnel.CreateSizedball BallSize / 2
  PlaySound "Thunder"
  vpmTimer.AddTimer 3300, "GiOn'"
  vpmtimer.addtimer 3000, "KickerReverseTunnelKick'"
  Hologram1Stop
  vpmtimer.addtimer 500, "Hologram2Start'"
  vpmtimer.addtimer 500, "TimeThiefInvisible'"
  vpmtimer.addtimer 600, "QManVisible'"
End Sub

Sub KickerReverseTunnelKick
  PlaySoundAt"fx_kicker",KickerReverseTunnel
  KickerReverseTunnel.kick 28, 45
WallReverseOutProtect.isdropped=True
End Sub


'*************************
'O-Light
'**************************

Sub CreateNewBallKickerReverseTunnel2
  KickerReverseTunnel.CreateSizedball BallSize / 2
  PlaySound "Thunder"
  Hologram1Stop

  If ReverseTimeComplete(CurrentPlayer)=False Then
    l39.State=1   '------------------------------------------------------Start 'O' light in LORD
    ReverseTimeComplete(CurrentPlayer)=True
    debug.print "ReverseTimeComplete(CurrentPlayer)=True"
    CheckTimeLordStatusTimer.Enabled=1
  End If
  vpmtimer.addtimer 3000, "KickerReverseTunnelKick'"
  vpmtimer.addtimer 500, "Hologram1Start'"
  vpmtimer.addtimer 500, "TimeThiefInvisible'"
  vpmtimer.addtimer 600, "QManVisible'"
  vpmtimer.addtimer 3200, "CheckEnableTimeWarpJackpots'"
End Sub

Sub KickerReverseTunnel_Unhit()
  KickerReverseTunnel.Enabled=0:KickerReverseLock1.Enabled=0:KickerReverseLock2.Enabled=0
  l_HourGlass.State=1
  ReverseReleaseActive=False
  debug.print "ReverseReleaseActive=False"
End Sub

Sub KickerRampTrap_Hit()
  WireRampOff
  KickerRampTrap.DestroyBall
  vpmtimer.addtimer 100, "EnableKickerRampTrapRelease'"
  vpmtimer.addtimer 200, "CreateNewBallKickerRampTrapRelease'"
End Sub


Sub CreateNewBallKickerRampTrapRelease
  KickerRampTrapRelease.CreateSizedball BallSize / 2
  PlaySoundAt "popper_ball", KickerRampTrapRelease: KickerRampTrapRelease.kick 325, 45
  vpmtimer.addtimer 200, "DisableKickerRampRelease'"
End Sub

Sub EnableKickerRampTrapRelease
  KickerRampTrapRelease.Enabled=1
End sub

Sub DisableKickerRampRelease
  KickerRampTrapRelease.Enabled=0
End Sub
'********************
'Time Warp Jackpots
'*******************
'Start TimeWarp Jackpots and disable Reverse time until 30 second timer runs out

Sub CheckEnableTimeWarpJackpots
  If  TWJPFiring=False Then: StartTimeWarpJackPots
End Sub

Dim TWJPFiring
Sub StartTimeWarpJackPots
  BG_Sequence 2,0,200 'clockspin ccw 200updates
  BG_Sequence 38,0,444'   nr 4 blue wizard binking
  BG_Sequence 5,10,0'

  bBallSaverReady = True  'Activate Ball Saver for all Balls
  TWJPFiring=True
  l61.State=2:l62A.State=2
If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_TimeWarpJPAcitvated":End If
  TimeWarpTimer.Enabled=True
  If bMultiBallMode=1 Then EnableBallSaver 30
  If bMultiBallMode=0 Then EnableBallSaver 30:AddMultiball (2)
  Light_ReverseClock.State=2
  vpmtimer.addtimer 500, "ResetReverseLockLights'"
  vpmtimer.addtimer 200, "TurnOffReverseTimeLights'"
  vpmtimer.addtimer 500, "TimeWarpMultiballDisplay'"
End Sub

Sub TimeWarpMultiballDisplay
  PlaySong "m_Slitheen"
  vpmtimer.addtimer 8000, "TimeWarpScore'"

End Sub

Sub TimeWarpScore
  DMDScoreNow
  AddScore 5000000
End Sub

Sub Trigger10a_Hit()
  LastSwitchHit = "Trigger10a"
debug.print "Trigger10aHit"
End Sub

Sub Trigger10b_Hit()
: Objlevel(4) = 1:FlasherFlash4_Timer
  CheckSpiralRampAward
  LastSwitchHit = "Trigger10b"
debug.print "Trigger10bHit"
End Sub

Sub Trigger10c_Hit()
  LastSwitchHit = "Trigger10c"
debug.print "Trigger10cHit"
End Sub

Sub Trigger10_Hit()
  If LastSwitchHit = "Trigger10a" Then:WrongTrajectory
  LastSwitchHit = "Trigger10"
End Sub



Sub WrongTrajectory
  GravityCallout
  DMD "", "", "DMD_WrongTraj", eNone, eNone, eNone, 2000, True, ""
  vpmTimer.AddTimer 3000, "DMDScoreNow'"
  AddScore 1000
End Sub


Sub CheckSpiralRampAward
  If  GloryTableOn(CurrentPlayer)=True Then:GloryBillions:StartFlasherSpiralAward:Exit Sub
  If l61.State=2 and GloryTableOn(CurrentPlayer)=False Then:TimeWarpJackPotCheck: StartFlasherSpiralAward
  If l61B.State=2 and GloryTableOn(CurrentPlayer)=False and TWJPFiring=False Then: AwardDefyGravity:End If
  If l61A.State=2 and GloryTableOn(CurrentPlayer)=False and TWJPFiring=False Then: AwardSpaceWalk
End Sub

Dim SpacewalkAward
Sub AwardSpaceWalk
  SpacewalkAward=SpacewalkAward+1
  Select Case SpacewalkAward
    Case 1:DMD "", "", "DMD_SpaceWalkLuigi", eNone, eNone, eNone, 2000, True, "CO_SpaceWalk": SpaceWalkAward1
    Case 2:DMD "", "", "DMD_SpaceWalkLuigi", eNone, eNone, eNone, 2000, True, "CO_SpaceWalk":SpaceWalkAward2
    Case 3:DMD "", "", "DMD_SpaceWalkLuigi", eNone, eNone, eNone, 2000, True, "CO_SpaceWalk":SpaceWalkAward3:l61B.State=2: l61A.State=1:SpacewalkAward=0
  End Select
End Sub

Sub SpaceWalkAward1
  vpmTimer.AddTimer 3000, "DMDScoreNow'"
  AddScore 500000
End Sub

Sub SpaceWalkAward2
  vpmTimer.AddTimer 3000, "DMDScoreNow'"
  AddScore 1000000
End Sub

Sub SpaceWalkAward3
  vpmTimer.AddTimer 3000, "DMDScoreNow'"
  AddScore 1300000
End Sub

Dim Crazy
Sub CrazyCallout
  Crazy=Crazy+1
  Select Case Crazy
    Case 1: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_CrazyEyes":End If
    Case 2: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_Linkages":End If
    Case 3: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_Trajectory":End If
        Crazy=0
  End Select
End Sub


Dim DefyAward
Sub AwardDefyGravity
  DefyAward=DefyAward+1
  Select Case DefyAward
    Case 1:DMD "", "", "DMD_SuperNova5", eNone, eNone, eNone, 2000, True, "CO_SuperNova":SuperNovaAward1
    Case 2:DMD "", "", "DMD_SuperNova5", eNone, eNone, eNone, 2000, True, "CO_SuperNova":SuperNovaAward2
    Case 3:DMD "", "", "DMD_SuperNova5", eNone, eNone, eNone, 2000, True, "CO_SuperNova":l61B.State=1: l61A.State=2:DefyAward=0:SuperNovaAward3
  End Select
End Sub

Sub SuperNovaAward1
  vpmTimer.AddTimer 3000, "DMDScoreNow'"
  AddScore 10000000
End Sub

Sub SuperNovaAward2
  vpmTimer.AddTimer 3000, "DMDScoreNow'"
  AddScore 15000000
End Sub

Sub SuperNovaAward3
  vpmTimer.AddTimer 3000, "DMDScoreNow'"
  AddScore 20000000
End Sub


Dim GravityCall
Sub GravityCallout
  GravityCall=GravityCall+1
  Select Case GravityCall
    Case 1: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_Cranium":End If
    Case 2: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_PleaseConcentrate":End If
    Case 3: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_Fart":End If
    Case 4: If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_DownANotch2":End If
        GravityCall=0
  End Select
End Sub

Sub CheckTimeWarpEasyHard
End Sub

Sub TimeWarpTimer_Timer
  debug.print "TimeWarpTimer activated"
  vpmTimer.AddTimer 5000, "ResetTimeWarpJackpots'"
End Sub


Sub AwardAlmostAJackpot
  AddScore 50000000
  DMD CL(0, "500000000"), CL(1, "ALMOST A JACKPOT"), "", eNone, eBlink, eNone, 2000, True, ""
  vpmTimer.AddTimer 2200, "DMDScoreNow'"
End Sub


Sub TimeWarpJackPotCheck
debug.print "Trigger10_Hit"
  PlaySoundAt "RampOut1_6",Trigger10
  If l62C.State=2 Then TimeWarpSuperJackPot:PlaySound "CO_SuperJackPot"
  If l62B.State=2 Then l62C.State=2:l62B.State=0:PlaySound "CO_JackPot":TimeWarpJackPot2
  If l62A.State=2 Then l62B.State=2:l62A.State=0:PlaySound "CO_JackPot":TimeWarpJackPot2
End Sub


Sub TimeWarpJackPot1
  DMD "", "", "DMD_JackPot100M", eNone, eNone, eNone, 4000, True, ""
  LS_MajorAward.Play SeqBlinking,, 6,100
  vpmTimer.AddTimer 4100, "JackPot1Score'"
End Sub

Sub JackPot1Score
  AddScore 100000000
  DMDScoreNow
End Sub

Sub TimeWarpJackPot2
  AwardSuperJackpot2
  LS_MajorAward.Play SeqBlinking,, 6,100
End Sub


Sub TimeWarpSuperJackPot
  AwardSuperJackpot2
  LS_MajorAward.Play SeqBlinking,, 6,100
End Sub

Sub JackPot3Score
  AddScore 100000000
  DMDScoreNow
End Sub

Sub StopTimeWarpJackpots
  l62A.State=0::l62B.State=0:l62C.State=0:l61.State=0
  TWJPFiring=False
  StopClockHandsCounterClockwise
  vpmtimer.addtimer 200, "TimeThiefVisible'"
  vpmtimer.addtimer 500, "QManInvisible'"
  vpmtimer.addtimer 500, "StartClockHandsClockwise'"
  TimeWarpTimer.Enabled=False
End Sub

Sub ResetTimeWarpJackpots '--------------------------------Called Reste Modes and Awards for new player variable
  l62A.State=0::l62B.State=0:l62C.State=0:l61.State=0
vpmtimer.addtimer 200, "TimeThiefVisible'"
  vpmtimer.addtimer 500, "QManInvisible'"
  StopClockHandsCounterClockwise
  vpmtimer.addtimer 500, "StartClockHandsClockwise'"
  TWJPFiring=False
  TimeWarpTimer.Enabled=False
  If  l21B.State=1 Then:l21.State=2:l20.State=1:l19.State=1
  If  l20A.State=1 Then:l20.State=2:l21.State=0:l19.State=1
  If  l19A.State=1 Then:l19.State=2:l20.State=0:l21.State=0
End Sub


Sub ResetTimeWarpJackPotsEndOfBall
  l62A.State=0::l62B.State=0:l62C.State=0:l61.State=0
  StopClockHandsCounterClockwise
  vpmtimer.addtimer 500, "StartClockHandsClockwise'"
  TWJPFiring=False
  TimeWarpTimer.Enabled=False
' If  l21B.State=1 Then:l21.State=2:l20.State=1:l19.State=1
' If  l20A.State=1 Then:l20.State=2:l21.State=0:l19.State=1
' If  l19A.State=1 Then:l19.State=2:l20.State=0:l21.State=0
End Sub



Sub GloryBillions
  BG_Sequence 2,0,200
  LightSeqGi.Play SeqBlinking,, 6,25
  GloryTimer_Timer
  Billions(CurrentPlayer) = Billions(CurrentPlayer) + 1
  Select Case Billions(CurrentPlayer)
    Case 1: Glory1Billion

    Case 2: Glory1Billion

    Case 3: Glory2Billion

    Case 4: Glory2Billion

    Case 5: Glory2Billion

    Case 6: Glory3Billion

    Case 7: Glory3Billion

    Case 8: Glory3Billion

    Case 9: Glory3Billion

    Case 10: Glory3Billion

    Case 11 Glory3Billion

    Case 12: Glory3Billion

    Case 13:Glory3Billion

    Case 14:Glory3Billion

    Case 15:Glory3Billion

    Case 16: Glory3Billion
      Billions(CurrentPlayer)=0

  End Select

End Sub

Sub Glory1Billion
    PlaySound "CO_Glory"
    DMD "", "", "DMD_Glory_1Bill", eNone, eNone, eNone, 2000, True, ""
    vpmtimer.addtimer 3000, "DMDScoreNow'"
    AddScore 1000000000
End Sub

Sub Glory2Billion
    PlaySound "CO_GloryDouble"
    DMD "", "", "DMD_Glory_2Bill", eNone, eNone, eNone, 2000, True, ""
    vpmtimer.addtimer 3000, "DMDScoreNow'"
    AddScore 2000000000
End Sub

Sub Glory3Billion
    PlaySound "CO_GloryTriple"
    DMD "", "", "DMD_Glory_3Bill", eNone, eNone, eNone, 2000, True, ""
    vpmtimer.addtimer 3000, "DMDScoreNow'"
    AddScore 3000000000
End Sub

'*******************
'Dragon Eyes Rotate
'*******************

Sub Trigger14_Hit()
  spinningeye.enabled = 1
  eyespin = 1
  spinningeye2.enabled = 1
  eye2spin = 1
  LastSwitchHit = "Trigger14"
End Sub


'********************
'Ramp Raise and Lower
'********************



Sub RaiseRampL
  BG_Sequence 1,0,200

  Ramp7.HeightBottom = 80
  Ramp7.Collidable = False
  LeftRampProtector.HeightTop=80
  LeftRampProtector.HeightBottom=80
  LeftRampProtector.Collidable = False
  PlaySoundAt "Wall_Hit_9", KickerRampTrap
  vpmtimer.addtimer 100, "DisableKickerRampTrap'"

End Sub


Sub LowerRampL
  Ramp7.HeightBottom = 0
  Ramp7.Collidable = True
  LeftRampProtector.HeightTop=12.5
  LeftRampProtector.HeightBottom=0
  LeftRampProtector.Collidable = True
  PlaySoundAt "Wall_Hit_2", KickerRampTrap
  KickerRampTrap.Enabled=1
End Sub

Sub KickerReverseLock1_UnHit
  DOF 122, DOFPulse
End Sub

Sub TimerLowerRampL_Timer
  LowerRampL
  TimerLowerRampL.enabled = False
  vpmtimer.addtimer 1000, "EnableKickerRampTrap'"
End Sub

Sub DisableKickerRampTrap
  KickerRampTrap.Enabled=0
End Sub

Sub EnableKickerRampTrap
  KickerRampTrap.Enabled=1
End Sub


'******************************
'DragonLairBarricadeTargets
'******************************

'Targetheat effect

Sub l1_Timer ' target1 light
Dim i
  If l1.timerenabled=0 Then
    l1.timerenabled=1
    l1.uservalue=0

  Else
    i = l1.uservalue
    Select Case i
      Case 0 : swTarget1.blenddisablelighting=1.0
      Case 1 : swTarget1.blenddisablelighting=1.5
      Case 2 : swTarget1.blenddisablelighting=1.5
      Case 3 : swTarget1.blenddisablelighting=1.5
      Case 4 : swTarget1.blenddisablelighting=1.5
      Case 5 : swTarget1.blenddisablelighting=1.5
      Case 6 : swTarget1.blenddisablelighting=2.0
      Case 7 : swTarget1.blenddisablelighting=2.0
      Case 8 : swTarget1.blenddisablelighting=2.0
      Case 9 : swTarget1.blenddisablelighting=2.0
      Case 10 : swTarget1.blenddisablelighting=2.0
      Case 11 : swTarget1.blenddisablelighting=2.4
      Case 12 : swTarget1.blenddisablelighting=2.4
      Case 13 : swTarget1.blenddisablelighting=2.4
      Case 14:  swTarget1.blenddisablelighting=2.4
      Case 15 : swTarget1.blenddisablelighting=2.4
      Case 16 : swTarget1.blenddisablelighting=2.4
      Case 17 : swTarget1.blenddisablelighting=2.4
      Case 18 : swTarget1.blenddisablelighting=2.4
      Case 19 : swTarget1.blenddisablelighting=2.4
      Case 20 : swTarget1.blenddisablelighting=2.4
      Case 21 : swTarget1.blenddisablelighting=2.0
      Case 22 : swTarget1.blenddisablelighting=2.0
      Case 23 : swTarget1.blenddisablelighting=2.0
      Case 24 : swTarget1.blenddisablelighting=2.0
      Case 25 : swTarget1.blenddisablelighting=2.0
      Case 26 : swTarget1.blenddisablelighting=2.0
      Case 27 : swTarget1.blenddisablelighting=1.8
        Case 28 : swTarget1.blenddisablelighting=1.8
      Case 29 : swTarget1.blenddisablelighting=1.7
      Case 30 : swTarget1.blenddisablelighting=1.7
      Case 31 : swTarget1.blenddisablelighting=1.6
      Case 32 : swTarget1.blenddisablelighting=1.6
      Case 33 : swTarget1.blenddisablelighting=1.3
      Case 34 : swTarget1.blenddisablelighting=1.2
      Case 35 : swTarget1.blenddisablelighting=1.0
      Case 36 : swTarget1.blenddisablelighting=1.0
      Case 37 : swTarget1.blenddisablelighting=1.0 : l1.uservalue=0

    l1.uservalue=i+1
    End Select

  End If
End Sub

Sub l2_Timer ' target2
Dim i
  If l2.timerenabled=0 Then
    l2.timerenabled=1
    l2.uservalue=0

  Else
    i = l2.uservalue
    Select Case i
      Case 0 : swTarget2.blenddisablelighting=1.0
      Case 1 : swTarget2.blenddisablelighting=1.5
      Case 2 : swTarget2.blenddisablelighting=1.5
      Case 3 : swTarget2.blenddisablelighting=1.5
      Case 4 : swTarget2.blenddisablelighting=1.5
      Case 5 : swTarget2.blenddisablelighting=1.5
      Case 6 : swTarget2.blenddisablelighting=2.0
      Case 7 : swTarget2.blenddisablelighting=2.0
      Case 8 : swTarget2.blenddisablelighting=2.0
      Case 9 : swTarget2.blenddisablelighting=2.0
      Case 10 : swTarget2.blenddisablelighting=2.0
      Case 11 : swTarget2.blenddisablelighting=2.4
      Case 12 : swTarget2.blenddisablelighting=2.4
      Case 13 : swTarget2.blenddisablelighting=2.4
      Case 14:  swTarget2.blenddisablelighting=2.4
      Case 15 : swTarget2.blenddisablelighting=2.4
      Case 16 : swTarget2.blenddisablelighting=2.4
      Case 17 : swTarget2.blenddisablelighting=2.4
      Case 18 : swTarget2.blenddisablelighting=2.4
      Case 19 : swTarget2.blenddisablelighting=2.4
      Case 20 : swTarget2.blenddisablelighting=2.4
      Case 21 : swTarget2.blenddisablelighting=2.0
      Case 22 : swTarget2.blenddisablelighting=2.0
      Case 23 : swTarget2.blenddisablelighting=2.0
      Case 24 : swTarget2.blenddisablelighting=2.0
      Case 25 : swTarget2.blenddisablelighting=2.0
      Case 26 : swTarget2.blenddisablelighting=2.0
      Case 27 : swTarget2.blenddisablelighting=1.8
        Case 28 : swTarget2.blenddisablelighting=1.8
      Case 29 : swTarget2.blenddisablelighting=1.7
      Case 30 : swTarget2.blenddisablelighting=1.7
      Case 31 : swTarget2.blenddisablelighting=1.6
      Case 32 : swTarget2.blenddisablelighting=1.6
      Case 33 : swTarget2.blenddisablelighting=1.3
      Case 34 : swTarget2.blenddisablelighting=1.2
      Case 35 : swTarget2.blenddisablelighting=1.0
      Case 36 : swTarget2.blenddisablelighting=1.0
      Case 37 : swTarget2.blenddisablelighting=1.0 : l2.uservalue=0

    l2.uservalue=i+1
    End Select
  End If
End Sub

Sub l3_Timer ' target3 light
Dim i
  If l3.timerenabled=0 Then
    l3.timerenabled=1
    l3.uservalue=0

  Else
    i = l3.uservalue
    Select Case i
      Case 0 : swTarget3.blenddisablelighting=1.0
      Case 1 : swTarget3.blenddisablelighting=1.5
      Case 2 : swTarget3.blenddisablelighting=1.5
      Case 3 : swTarget3.blenddisablelighting=1.5
      Case 4 : swTarget3.blenddisablelighting=1.5
      Case 5 : swTarget3.blenddisablelighting=1.5
      Case 6 : swTarget3.blenddisablelighting=2.0
      Case 7 : swTarget3.blenddisablelighting=2.0
      Case 8 : swTarget3.blenddisablelighting=2.0
      Case 9 : swTarget3.blenddisablelighting=2.0
      Case 10 : swTarget3.blenddisablelighting=2.0
      Case 11 : swTarget3.blenddisablelighting=2.4
      Case 12 : swTarget3.blenddisablelighting=2.4
      Case 13 : swTarget3.blenddisablelighting=2.4
      Case 14:  swTarget3.blenddisablelighting=2.4
      Case 15 : swTarget3.blenddisablelighting=2.4
      Case 16 : swTarget3.blenddisablelighting=2.4
      Case 17 : swTarget3.blenddisablelighting=2.4
      Case 18 : swTarget3.blenddisablelighting=2.4
      Case 19 : swTarget3.blenddisablelighting=2.4
      Case 20 : swTarget3.blenddisablelighting=2.4
      Case 21 : swTarget3.blenddisablelighting=2.0
      Case 22 : swTarget3.blenddisablelighting=2.0
      Case 23 : swTarget3.blenddisablelighting=2.0
      Case 24 : swTarget3.blenddisablelighting=2.0
      Case 25 : swTarget3.blenddisablelighting=2.0
      Case 26 : swTarget3.blenddisablelighting=2.0
      Case 27 : swTarget3.blenddisablelighting=1.8
        Case 28 : swTarget3.blenddisablelighting=1.8
      Case 29 : swTarget3.blenddisablelighting=1.7
      Case 30 : swTarget3.blenddisablelighting=1.7
      Case 31 : swTarget3.blenddisablelighting=1.6
      Case 32 : swTarget3.blenddisablelighting=1.6
      Case 33 : swTarget3.blenddisablelighting=1.3
      Case 34 : swTarget3.blenddisablelighting=1.2
      Case 35 : swTarget3.blenddisablelighting=1.0
      Case 36 : swTarget3.blenddisablelighting=1.0
      Case 37 : swTarget3.blenddisablelighting=1.0 : l3.uservalue=0

    l3.uservalue=i+1
    End Select
  End If
End Sub

Sub swTarget1_Dropped
  LS_Eyes.Play SeqBlinking,, 1,25
  DragonTargetsUp=False
  PlaySound "fx_dragonfire"
  AddScore 1000000
  l1.State =0
  swTarget1.IsDropped= True 'This drops the target
  PlaySoundAT "Drop_Target_Down_1" ,swtarget1
  CheckBarricadeTargetsComplete
  l1.timerenabled=0
End Sub

Sub swTarget2_Dropped
  LS_Eyes.Play SeqBlinking,, 1,25
  DragonTargetsUp=False
  PlaySoundAt "Drop_Target_Down_2" ,swtarget2
  PlaySound "fx_dragonfire"
  AddScore 1000000
  l2.State =0
  swTarget2.IsDropped= True 'This drops the target
  CheckBarricadeTargetsComplete
  l2.timerenabled=0

End Sub

Sub swTarget3_Dropped
  LS_Eyes.Play SeqBlinking,, 1,25
  DragonTargetsUp=False
  PlaySound "fx_dragonfire"
  AddScore 1000000
  l3.State =0
  swTarget3.IsDropped= True 'This drops the target
  PlaySoundAT "Drop_Target_Down_3" ,swtarget3
  CheckBarricadeTargetsComplete
  l3.timerenabled=0
End Sub

Sub DropDragonTargets
  swTarget1.IsDropped= True :swTarget2.IsDropped= True :swTarget3.IsDropped= True :DragonLairWall.IsDropped= True
  PlaySoundAT "Drop_Target_Down_3" ,swtarget3
  l1.State=0:l2.State=0:l3.State=0:LightSword.State=2:l75.State=2:l4.State=2
  l1.timerenabled=0 :l2.timerenabled=0 :l3.timerenabled=0
  TriggerTrapped.Enabled=0:KickerSpinnerBallTrapped.Enabled=0
End Sub

Sub CheckBarricadeTargetsComplete
  AddScore 200000
  PlaySound "fx_dragonfire"
  If l1.State + l2.State + l3.State = 0 Then :DragonDropTargetsComplete
End Sub

Sub DragonDropTargetsComplete
  Debug.Print "DragonDropTargetsComplete"
  LightSword.State=2:l4.State=1
  l5.State=2:l6.State=2:l7.State=2:l8.State=2:l9.State=2:l10.State=2
End Sub

Dim DragonTargetsUp
Sub LiftBarricadeTargets
  DragonLairWall.IsDropped= False
  swTarget1.IsDropped= False
  swTarget2.IsDropped= False
  swTarget3.IsDropped= False
  DragonTargetsUp=True
  l1.timerenabled=1 :l2.timerenabled=1 :l3.timerenabled=1
End Sub

Sub ResetBarricadeTargets
  l1_Timer
  DragonTargetsUp=True
  DOF 111,DOFPulse
  swTarget1.IsDropped= False
  PlaysoundAt SoundFX("Drop_Target_Reset_1",DOFContactors) , swTarget1
  swTarget2.IsDropped= False
  PlaysoundAt SoundFX("Drop_Target_Reset_2",DOFContactors) , swTarget2
  swTarget3.IsDropped= False
  PlaysoundAt SoundFX("Drop_Target_Reset_3",DOFContactors) , swTarget3
  l1.State= 2
  l2.State= 2
  l3.State= 2
  l4.State= 2
  l75.State=0
  l5.State=0:l6.State=0:l7.State=0:l8.State=0:l9.State=0:l10.State=0
  l1.timerenabled=1 :l2.timerenabled=1 :l3.timerenabled=1
End Sub

Sub DragonLairWall_Hit
  BG_Sequence 5,10,0
  If l4.State=1 Then
  PlaySong "m_Unit"
  dropwall
  l75.State=2
  End If
End Sub




'************************
'Dragon Targets
'************************

Sub Target1_Hit
  l5. State = 1
  CheckDragonLairTargetLights
End Sub

Sub Target2_Hit
  l6. State = 1
  CheckDragonLairTargetLights
End Sub


Sub Target3_Hit
  l7. State = 1
  CheckDragonLairTargetLights
End Sub

Sub Target4_Hit
  l8. State = 1
  CheckDragonLairTargetLights
End Sub

Sub Target5_Hit
  l9. State = 1
  CheckDragonLairTargetLights
End Sub


Sub Target6_Hit
  l10. State = 1
  CheckDragonLairTargetLights
End Sub


Sub DragonTargetFlashersOn
  DragonTargetFlasher1.Visible=1:DragonTargetFlasher2.Visible=1:DragonTargetFlasher3.Visible=1
  DragonTargetFlasher4.Visible=1:DragonTargetFlasher5.Visible=1:DragonTargetFlasher6.Visible=1
End Sub

Sub DragonTargetFlashersOff
  DragonTargetFlasher1.Visible=0:DragonTargetFlasher2.Visible=0:DragonTargetFlasher3.Visible=0
  DragonTargetFlasher4.Visible=0:DragonTargetFlasher5.Visible=0:DragonTargetFlasher6.Visible=0
End Sub

Sub bLightsOn
  Light_BlackCube1.State=1:Light_BlackBox2.State=1:Light_WomanLL.State=1:Light_WomanRL.State=1:LightFireClock2.State=1
  Light_ClockGreen1.State=1:Light_ClockGreen2.State=1:Light_Red1.State=1:LightRed2.State=1:Light_DragonHead.State=1:Light_WheelRed4.State=1
  Light_RightOrbitLane.State=1:Light_WormLane.State=1:Light_RedHand.State=1:Light_ReverseGreen.State=1:Light_LeftLaneRedHand.State=1
End Sub

Sub bLightsoff
  Light_BlackCube1.State=0:Light_BlackBox2.State=0:Light_WomanLL.State=0:Light_WomanRL.State=0:LightFireClock2.State=0
  Light_ClockGreen1.State=0:Light_ClockGreen2.State=0:Light_Red1.State=0:LightRed2.State=0:Light_DragonHead.State=0:Light_WheelRed4.State=0
  Light_RightOrbitLane.State=0:Light_WormLane.State=0:Light_RedHand.State=0:Light_ReverseGreen.State=0:Light_LeftLaneRedHand.State=0
End Sub


Dim DragonFlashBlink
Sub DragonTargetFlasherBlinkTimer_Timer

DragonFlashBlink=DragonFlashBlink+1
  Select Case  DragonFlashBlink
    Case 1: DragonTargetFlashersOff:bLightsoff
    Case 4:DragonTargetFlashersOn:bLightsOn
    Case 7: DragonTargetFlashersOff:bLightsoff
    Case 10: DragonTargetFlashersOn:bLightsOn
    Case 13: DragonTargetFlashersOff:bLightsoff
    Case 16: DragonTargetFlashersOn:bLightsOn
    Case 19: DragonTargetFlashersOff:bLightsoff
    Case 22:DragonTargetFlashersOn:bLightsOn:DragonTargetFlasherBlinkTimer.Enabled=0:DragonFlashBlink=0:bLightsOn
  End Select
End Sub



Sub CheckDragonLairTargetLights
  If DragonWounded=True Then: Exit Sub: End If
  If l5.State + l6.State + l7.State +l8.State + l9.State + l10.State = 6 Then
    DragonWounded=True
    raisewall
    DragonTargetFlasherBlinkTimer.Enabled=True
    LightsDragonWounded
    CheckDragonWound
    debug.print "CheckDragonWound"
    DragonMoveUp
  End If
End Sub


Sub CheckDragonWound
  If DragonWounded4(CurrentPlayer)=True  Then DragonWound5:LightsDragonDefeated:debug.print "DragonWound4": Exit Sub: End If
  If DragonWounded3(CurrentPlayer)=True  Then DragonWound4:DragonWounded4(CurrentPlayer)=True:LightsDragonDefeated:debug.print "DragonWounded4(CurrentPlayer)":Exit Sub:End If
  If DragonWounded2(CurrentPlayer)=True  Then DragonWound3:DragonWounded3(CurrentPlayer)=True:LightsDragonWounded: debug.print "DragonWounded3(CurrentPlayer)":Exit Sub:End If
  If DragonWounded1(CurrentPlayer)=True Then DragonWound2:DragonWounded2(CurrentPlayer)=True:LightsDragonWounded:debug.print "DragonWounded2(CurrentPlayer)":Exit Sub :End If
  If DragonWounded1(CurrentPlayer)=False Then DragonWound1:LightsDragonWounded:DragonWounded1(CurrentPlayer)=True:debug.print "DragonWounded1(CurrentPlayer)"
End Sub

Sub LightsDragonWounded
  LS_DragonWound.Play SeqBlinking,, 4,100
End Sub
Sub LightsDragonDefeated
  LS_MajorAward.Play SeqBlinking,, 8,100
End Sub


Sub DragonWound1
    BG_Sequence 39,0,200  ' nr 3 white wizard binking
    BG_Sequence 5,10,0    ' (All BG lights blinking)
    GiOff
    vpmtimer.addtimer 3200, "GiOn'"
    StartFlasherAward
    l42.State=1: AddScore 500000000:l4.state=0:l75.State=2 'Set for trigger lair check to enable Time lord to capture the ball
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_DragonWounded3More":End If
    DMD "", "", "DMD_DragonWounded", eNone, eNone, eNone, 3000, False, ""
    vpmtimer.addtimer 2500, "WatchClosely'"
    vpmtimer.addtimer 3500, "DMDScoreNow'"
End Sub

Sub DragonWound2
    BG_Sequence 39,0,200  ' nr 3 white wizard binking
    BG_Sequence 5,10,0    ' (All BG lights blinking)
    GiOff
    vpmtimer.addtimer 3200, "GiOn'"
    StartFlasherAward
    l43.State=1: AddScore 500000000:l4.state=0:l75.State=2
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_DragonWounded2More":End If
    DMD "", "", "DMD_DragonWounded", eNone, eNone, eNone, 3000, False, ""
    vpmtimer.addtimer 2500, "WatchClosely'"
    vpmtimer.addtimer 3500, "DMDScoreNow'"
End Sub

Sub DragonWound3
    BG_Sequence 39,0,200  ' nr 3 white wizard binking
    BG_Sequence 5,10,0    ' (All BG lights blinking)
    GiOff
    vpmtimer.addtimer 3200, "GiOn'"
    StartFlasherAward
    l44.State=1: AddScore 500000000:l4.state=0:l75.State=2
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_DragonWounded1More":End If
    DMD "", "", "DMD_DragonWounded", eNone, eNone, eNone, 3000, False, ""
    vpmtimer.addtimer 2500, "WatchClosely'"
    vpmtimer.addtimer 3500, "DMDScoreNow'"
End Sub


Sub DragonWound4
    BG_Sequence 38,0,200
    BG_Sequence 1,0,20
    BG_Sequence 5,4,0
    GiOff
    vpmtimer.addtimer 3200, "GiOn'"
    StartFlasherAward

    l45.State=1: AddScore 1000000000:CheckTimeLordStatusTimer.Enabled=1:l23B.State=0:l4.state=0:l75.State=2:StartPlanetDragon1
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_YouKIlledThat":End If
    DMD "", "", "DMD_DragonDefeated", eNone, eNone, eNone, 3000, False, ""
'   vpmtimer.addtimer 2800, "WatchClosely'"
    vpmtimer.addtimer 3500, "DMDScoreNow'"
End Sub

Sub WatchClosely
' If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_WatchClosely":End If
  PlaySound "CO_WatchClosely"
End Sub

Sub DragonWound5
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_WatchClosely":End If
  AwardSuperJackpot
End Sub

Sub ResetDragonLairTargetLights
  l5.State= 0
  l6.State= 0
  l7.State= 0
  l8.State= 0
  l9.State= 0
  l10.State= 0

  l4.State=2
  LightSword.State=2
End Sub

'***********************
'DragonLairTurntable
'***********************
Sub DragonImageSpinTimer_Timer
  spindiscimg.rotz = spindiscimg.rotz + 2
End Sub


Sub spinning_timer
  DragonImageSpinTimer.Enabled=False
  spindiscimg.rotz = spindiscimg.rotz + 10
  DOF 119,DOFPulse
  DOF 124,DOFPulse
  DOF 126,DOFPulse
End Sub

dim spinner
Set spinner = New cvpmTurntable
With spinner
  .InitTurntable spindisc, 80
  .SpinDown = 20
  .CreateEvents "spinner"
End With
spinner.MotorOn = false

'****************************************
'Raisewall Sub'
'****************************************
Sub raisewall
  TriggerTrapped.Enabled=0
  '------------------------------------------------
  DragonLairWall.IsDropped= False
  '------------------------------------------------
  spinner.MotorOn = true
  spinning.enabled = True
debug.print "motorspinneron"
  vpmtimer.addtimer 2800, "dropwall '"
  vpmtimer.addtimer 3300, "LiftWall '"
  vpmtimer.addtimer 3800, "CheckReverseReleaseActive '"
End Sub

Sub StopMotorTimer
  MotorSpinnerTimer.Enabled=False
  vpmtimer.addtimer 2000, "EnableSpinnerBallTrappedTrigger '"
debug.print "motorspinneroff"
debug.print "BallSpinnerTriggerEnabled in StopMotorSub"

End Sub


'****************************************
' Enable a release Kicker in case a ball gets stuck. Ball is destroyed in the spinner and created outside.


Sub EnableSpinnerBallTrappedTrigger
  TriggerTrapped.Enabled=1
  debug.print "TriggerBallTrappedEnabled"

' vpmtimer.addtimer 5000, "DeactivateKickerSpinnerBallTrapped'"

End Sub

'Sub DeactivateKickerSpinnerBallTrapped
' TriggerTrapped.Enabled=0
' debug.print "TriggerBallTrappedDisabled"
'End Sub


Sub TriggerTrapped_Hit()
  vpmtimer.addtimer 2000, "ActivateKickerSpinnerBallTrapped'"
  TriggerTrapped.Enabled=0
  debug.print " TriggerTrapped disabled"
End Sub

Sub ActivateKickerSpinnerBallTrapped
  KickerSpinnerBallTrapped.Enabled=1
  debug.print "KickerBallTrappedEnabled"
End Sub

Sub KickerSpinnerBallTrapped_Hit()
  KickerSpinnerBallTrapped.DestroyBall
  debug.print "KickerTrappedBallDestroyed"
  PlaySoundAt "Ball_Drop_Playfield_5_Delayed",Kicker35 'BalldropSound
  CreateNewBallKickerSpinnerRelease
  KickerSpinnerBallTrapped.Enabled=0
  debug.print "KickerTrappedBallDisabled"
End Sub

Sub CreateNewBallKickerSpinnerRelease
  debug.print "KickerSpinnerReleaseBallCreated"
  KickerSpinnerRelease.Enabled=1
  debug.print "ReleaseBallCreated"
  PlaySoundAt "VUKEnter", Kicker2
  KickerSpinnerRelease.CreateSizedball BallSize / 2
  vpmtimer.addtimer 200, "KickerSpinnerReleaseKick'"
End Sub

Sub KickerSpinnerReleaseKick
  TriggerTrapped.Enabled=True
  debug.print "TriggerTrappedEnabled"
  KickerSpinnerRelease.Kick 10, 120
  KickerSpinnerRelease.Enabled=0
  debug.print "ReleaseKickerDisabled"
End Sub


'****************************************

'*************************
'LiftWall
'*************************
Sub LiftWall
DragonWallDropActivated=False
  DragonLairWall.IsDropped= False
' If GloryResetTimer.Enabled=True Then:vpmtimer.addtimer 4000, "DisableTriggerTrappedGloryStop '":Exit Sub
  debug.print "LiftWall"
End Sub

Sub DisableTriggerTrappedGloryStop
' TriggerTrapped.Enabled=0
  GloryResetTimer.Enabled=False
End Sub


'*************************************
'dropwall
'************************************
Sub dropwall
  TriggerTrapped.Enabled=False
  KickerSpinnerBallTrapped.Enabled=False
  debug.print "Trigger and Kicker Trapped disabled"
  DragonLairWall.IsDropped= True
  spinner.MotorOn = false
  spinning.enabled = false
  vpmtimer.addtimer 200, "TurnOffNebulaLights'"

End Sub

Sub TurnOffNebulaLights
  l5.State= 0
  l6.State= 0
  l7.State= 0
  l8.State= 0
  l9.State= 0
  l10.State= 0
  l4.State=2
  LightSword.State=2
End Sub
'*************************************
'StopMotorSpinner
'**************************************
Sub   StopSpinnerMotor
  debug.print "StopMotor"
  spinner.MotorOn = false
  spinning.enabled = false
End Sub

'********************************
'TriggerLairCheck
'********************************
Sub   TriggerLairCheck_Hit()
  debug.print "TriggerLairSwitch Hit"
  PlaySound "fx_dragonfire"

  If l75.State=2 Then:LastSwitchHit = "TriggerLairSwitch":l_HourGlass2.State=2
  If DragonTargetsUp=True Then:ReleaseTrappedBallBehindDragonTargets
End Sub

Sub ReleaseTrappedBallBehindDragonTargets
  swTarget1.IsDropped= True
  swTarget2.IsDropped= True
  swTarget3.IsDropped= True
  vpmtimer.addtimer 1500, "LiftBarricadeTargets '"
End Sub

'**********************************
'ActivateTimeLordCatch
'**********************************
Sub CheckReverseReleaseActive
  If ReverseReleaseActive=True Then TimeLordCatchDelayTimer.Enabled=True
  If ReverseReleaseActive=False Then ActivateTimeLordCatch
End Sub


Dim TimeLordCatchEnabled
Sub ActivateTimeLordCatch
  l75.State=0
  TimeLordCatchEnabled=True
  debug.print "TimeLordCatchEnabled=True"
' PlaySound "CO_WatchClosely"
  DragonWounded=False
  vpmTimer.AddTimer 100, "EnableKickerLairOut'"
  DragonImageSpinTimer.Enabled=True
  EnableSpinnerBallTrappedTrigger
End Sub

Sub TimeLordCatchDelayTimer_Timer
  ActivateTimeLordCatch
  TimeLordCatchDelayTimer.Enabled=False
End Sub


'*********************************
'DestroyBallKickerLairOut
'*********************************

Sub DestroyKickerLairOutBall
  KickerLairOut.DestroyBall
  PlaySoundAt "popper_ball",KickerLairOut
  vpmtimer.addtimer 100, "CreateNewBallKickerWizardRelease'"

End Sub

'*********************************************************************************************************************************************
'Dragon Wound & Time Lights- The advance time Kicker 8 can also light Time Lights & runs in parrallel with this section with independant subs
'**********************************************************************************************************************************************

Dim WoundDragon
Sub KickerLairOut_Hit()
  BG_Sequence 5,10,0
  l_HourGlass2.State=2
  l75.State=0
  GiOff
  vpmtimer.addtimer 8000, "GiOn'"
  TimeLordFireworks
  ResetDragonLights
  ResetBarricadeTargets
  ResetDragonLairTargetLights

  debug.print "TimeLordCatchEnabled=False"
  WoundDragon = WoundDragon + 1
  Select Case WoundDragon
    Case 1
      LiftWall
      CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_HaHaGoingToBeGood"
        PlaySong "m_Westminster Bridge"
      DestroyKickerLairOutBall
      vpmTimer.AddTimer 500, "TimeThiefVisible'"
      vpmTimer.AddTimer 200, "QManInvisible'"
      vpmtimer.addtimer 2000, "TimeLordCatchFlash'"
      vpmtimer.addtimer 50, "DisableKickerLairOut'"
      vpmTimer.AddTimer 500, "Hologram1Start'"
    Case 2
      LiftWall
      CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_WillaWollaBingBang"
        PlaySong "m_Slitheen"
      DestroyKickerLairOutBall
      vpmTimer.AddTimer 500, "TimeThiefVisible'"
      vpmTimer.AddTimer 200, "QManInvisible'"
      vpmtimer.addtimer 2000, "TimeLordCatchFlash'"
      vpmtimer.addtimer 50, "DisableKickerLairOut'"
      vpmTimer.AddTimer 500, "Hologram1Start'"
    Case 3
      LiftWall
      CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_BumblingNinny"
      DestroyKickerLairOutBall
      vpmTimer.AddTimer 500, "TimeThiefVisible'"
      vpmTimer.AddTimer 200, "QManInvisible'"
      vpmtimer.addtimer 2000, "TimeLordCatchFlash'"
      vpmtimer.addtimer 50, "DisableKickerLairOut'"
      vpmTimer.AddTimer 500, "Hologram1Start'"
    Case 4
      LiftWall
      CalloutTimer.Enabled=True:CalloutActive=true:PlaySound"CO_TimeComplete"
      DestroyKickerLairOutBall
      vpmTimer.AddTimer 500, "TimeThiefVisible'"
      vpmTimer.AddTimer 200, "QManInvisible'"
      vpmtimer.addtimer 2000, "TimeLordCatchFlash'"
      vpmtimer.addtimer 50, "DisableKickerLairOut'"
      vpmTimer.AddTimer 500, "Hologram1Start'"
      WoundDragon=0
  End Select
End Sub

Sub KickerLairOutKick
  PlaySoundAt"fx_kicker", KickerLairOut
  KickerLairOut.kick 220, 60
End Sub

Dim BallCatchInPlay
Sub DisableKickerLairOut
  KickerLairOut.Enabled=False
  BallCatchInPlay=False:debug.print "Ball Catch in Play=False"
End Sub

Sub TimeLordFireworks
  LS_TimeLordCatch.UpdateInterval = 10
  LS_TimeLordCatch.Play SeqUpOn, 15,  1
  LS_TimeLordCatch.UpdateInterval = 10
  LS_TimeLordCatch.Play SeqDownOn, 15, 2
  LS_TimeLordCatch.UpdateInterval = 10
  LS_TimeLordCatch.Play SeqUpOn, 15,  1
  LS_TimeLordCatch.UpdateInterval = 10
  LS_TimeLordCatch.Play SeqDownOn, 15, 2
  LS_TimeLordCatch.UpdateInterval = 15
  LS_TimeLordCatch.Play SeqCircleOutOn, 15, 1
End Sub


'******************************************
'Check TimeLord Status- Modes Complete
'********************************************

Sub CheckTimeLordStatusTimer_Timer()
  CheckTIMELORDStatus
  vpmtimer.addtimer 100,"DisableCheckTimeLordStatusTimer'"
End Sub


Sub DisableCheckTimeLordStatusTimer
  CheckTimeLordStatusTimer.Enabled=0
  DMD "", "", "DMD_Blank", eNone, eNone, eNone, 200, False, ""
  vpmtimer.addtimer 500, "DMDScoreNow'"
  debug.print "DisableCheckTimeLordStatusTimer"
End Sub

Dim TimeLordComplete
Sub CheckTIMELORDStatus
  debug.print "CheckTIMELORDStatus"
  If  SuperLoopsComplete(CurrentPlayer)=True and ChaosCriticalComplete(CurrentPlayer)=True and WormLocksComplete(CurrentPlayer)=True and ReverseTimeComplete(CurrentPlayer)=True and DragonWounded4(CurrentPlayer)=True Then
    l62A.State=0:l62B.State=0:l62A.State=0:l61.State=0
'   TimeLordCompleteTimer.Enabled=True
    '********************
    TimeLordComplete=True
    debug.print "TimeLordComplete"
    '********************
    bBallSaverReady = True  'Activate Ball Saver for all Balls
    vpmTimer.addtimer 2000, "StartGloryBillions'"
  End If
End Sub

'Sub TimeLordCompleteTimer_Timer ()
' TimeLordCompleteTimer.Enabled=True
'End Sub


'**********************
'CheckTimeLights Are Complete
'**********************

Sub CheckTimeComplete
  If l45.State=1 Then :TimeComplete= True: CalloutTimeComplete: End If
End Sub

Sub CalloutTimeComplete
  ' AddScore 500000000
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_TimeComplete":End If
  '   DMD CL(0, "500000000"), CL(1, "YOU CONTROL TIME"), "", eNone, eBlink, eNone, 2200, True, ""
  ' vpmTimer.AddTimer 2000, "DMDScoreNow'"
  vpmTimer.AddTimer 2000, "Hologram2Start'"
End Sub

Sub CalloutTimeAward
  'AddScore 10000000
  If CalloutActive=False Then CalloutTimer.Enabled=True:CalloutActive=true:PlaySound "CO_AdvanceTimeAward":End If
  '   DMD CL(0, "10000000"), CL(1, "ADVANCE TIME"), "", eNone, eBlink, eNone, 2200, True, ""
' vpmTimer.AddTimer 2000, "DMDScoreNow'"
End Sub
'***************************
'CheckLORDLightsAreComplete
'***************************
Dim LordComplete

Sub CheckLordComplete
  If DragonWounded4=1 Then: LordComplete= True
End Sub


'*********************
'TimeLord/QMan Change
'********************

Sub QManVisible

  Robot.Visible=False
  QMan.Visible=1
  l_Staff1.State=2
  l_Staff2.State=2
  TimeLordFlash
End Sub

Sub QManInvisible
  QMan.Visible=0
  l_Staff1.State=0
  l_Staff2.State=0
  TimeLordFlash
End Sub


Sub TimeThiefVisible
  Robot.Visible=False
  TimeThief.Visible=1
  l_HourGlass.State=2
  l_HourGlass2.State=2
  Light_Catch2.State=2
  TimeLordCatchFlash
  TimeLordFlash
End Sub

Sub TimeThiefInvisible
  TimeThief.Visible=0
  l_HourGlass.State=0
  l_HourGlass2.State=0
  Light_Catch2.State=0
  TimeLordFlash
End Sub

'*************************************
'Time Lord Catch Kicker1-KickerLairOut
'*************************************
Sub EnableKickerLairOut
  debug.print "EnableKickerLairOut"
  TriggerLairOut.Enabled=False
  KickerLairOut.Enabled=True
  BallCatchInPlay=True:debug.print "BallCatchInPlay"
  l_HourGlass2.State=2
End Sub




'******************************************
'CreateBallWizardReleaseInvisibleSpiralRamp
'*******************************************

Sub CreateNewBallKickerWizardRelease
  KickerWizardRelease.CreateSizedball BallSize / 2
  KickerLairOut.Enabled=0
  TimeLordCatchEnabled=False
  debug.print "KickerLairOutEnabled=FALSE"
  debug.print "TimeLordCatchEnabled=False"
  vpmtimer.addtimer 5000, "PlaySoundAt""fx_woosh"", KickerWizardRelease: KickerWizardRelease.kick 290, 120'"
  vpmtimer.addtimer 7000, "DragonMoveDown'"
  PlaySong "m_The Carrionites Swarm"
End Sub

'**********************************
'TimeLordBallCatchFlashes
'**********************************
Sub TimeLordCatchFlash()
  FlashForMs f3e1, 500, 50, 0
  FlashForMs f3e2, 500, 50, 0
  FlashForMs f3e3, 500, 50, 0
  FlashForMs f3e4, 500, 50, 0
  DOF 121, DOFPulse
End Sub

Sub TimeLordFlash()
  FlashForMs ftl, 500, 50, 0
  FlashForMs ftl2, 500, 50, 0
  FlashForMs ftl3, 500, 50, 0
  FlashForMs ftl4, 500, 50, 0
End Sub



Sub TimeLightsBlinking
  l42.State=2:l43.State=2:l44.State=2:l45.State=2
  vpmtimer.addtimer 5000, "ResetTimeLights'"
End Sub

'****************************
'ResetTimeLights
'****************************

Sub ResetTimeLights
  l42.State=0:l43.State=0:l44.State=0:l45.State=0:ResetDragonLights
  debug.print "ResetTimeLights"
End Sub

Sub ResetDragonLights
  LightSword.State=2:l4.State=2
End Sub
'************************
'End Dragon Mode
'************************

'*********************************************
'Start Billions-GloryTableTimer
'*********************************************

Sub StartGloryBillions
  GloryTableStarted(CurrentPlayer)=True
  GloryTableOn(CurrentPlayer)=True
debug.print "GloryTableOn=True-StartGlorybillions"
  BG_Sequence 2,0,200 '(clockspin ccw 200updates)
  If bMultiBallMode=1 Then EnableBallSaver 50
  If bMultiBallMode=0 Then EnableBallSaver 60:AddMultiball (2)
  l61.State=0:l61a.State=0:l61B.State=0
  Trigger10.Enabled=True:Trigger10a.Enabled=True:Trigger10b.Enabled=True
  StopTimeWarpJackpots
' DMD "", "", "DMD_Glory_hitCR", eNone, eNone, eNone, 2000, True, ""
  vpmtimer.addtimer 2005, "DMDScoreNow'"
  AddScore 500000000
  GloryMultiBallTimer.Enabled=1
  PlaySong "m_Slitheen"
  vpmtimer.addtimer 200, "TurnOffPlayfieldLights'"
  FlasherGloryTable.Visible=1
  SetLightColor lGlory1, white, -1
  lGlory1.State=2
  SetLightColor lGlory2, white, -1
  lGlory2.State=2
  SetLightColor lGlory3, white, -1
  lGlory3.State=2
  SetLightColor lGlory4, white, -1
  lGlory4.State=2
  SetLightColor lGlory5, white, -1
  lGlory5.State=2
  ClockGloryTableDragon
  ClockGloryBig
  ClockGlory
  spindiscimg.visible=False
End Sub


Sub GloryMultiBallTimer_Timer
'   If GloryTableOn(CurrentPlayer)=False Then GloryComplete(CurrentPlayer)=False:GloryTableOn(CurrentPlayer)=False:
    If GloryTableOn(CurrentPlayer)=True Then GloryComplete(CurrentPlayer)=True:GloryTableBonus:GloryStop

    GloryMultiBallTimer.Enabled=False
debug.print "GloryMultiballTimer-TimedOut"
End Sub

Sub GloryTableBonus
    PlaySound "CO_Glory"
    DMD "", "", "DMD_Glory_1Bill", eNone, eNone, eNone, 2000, True, ""
    vpmtimer.addtimer 3000, "DMDScoreNow'"
    AddScore 1000000000
End Sub

Sub GloryCallout
  PlaySound "CO_YouAreTheTimeLord"
End Sub


Sub TurnOffModes
  StopSuperLoops
  StopSuperJets
  StopChainedLightning
  StopAdvanceTime
  StopExtraBall
  StopLostInSpace
  StopReverseTime
  StopTimeWarpJackpots
  DisableWormLocks
  EndMeteorShower
  ResetLeftSpinLights
  ResetRightSpinLights
  l4.State=0:l75.State=0:LightSword.State=0
  QManVisible
End Sub


Sub GloryTimer_Timer
  FlasherGloryFlash
  vpmtimer.addtimer 600, "FlasherGloryFlash'"
  vpmtimer.addtimer 1200, "FlasherGloryFlash'"
  vpmtimer.addtimer 1800, "FlasherGloryFlash'"
  vpmtimer.addtimer 2400, "FlasherGloryFlash'"
  vpmtimer.addtimer 3000, "FlasherGloryFlash'"
  vpmtimer.addtimer 3600, "FlasherGloryFlash'"
  vpmtimer.addtimer 4000, "GloryFlasherOff'"
End Sub


Sub FlasherGloryFlash
  FlasherGlory.Visible=1
  FlashForMs FlasherGlory, 500, 100, 0

End Sub


Sub GloryFlasherOff
  FlasherGlory.Visible=0
End Sub



Sub GloryStop     'Game in play
debug.print "GloryStop"
  GloryTableOn(CurrentPlayer)=False
debug.print "GloryTableOn=False-GloryStop"
  GloryTableStarted(CurrentPlayer)=False
  GloryComplete(CurrentPlayer)=True
  EndOfGloryTable=True
' GloryTableTimer.Enabled=0
  lGlory1.State=0:lGlory2.State=0:lGlory3.State=0:lGlory4.State=0:lGlory5.State=0
  FlasherGloryTable.Visible=0
  StopClockGlory
  StopClockGloryTableDragon
  StopClockGloryBig
  FlasherGlory.Visible=0
  GloryMultiBallTimer.Enabled=False
  spindiscimg.visible=True
  l38.State=0:l39.State=0:l40.State=0:l41.State=0:l42.State=0:l43.State=0:l44.State=0:l45.State=0
  vpmtimer.addtimer 2000, "GiOff'"
  vpmtimer.addtimer 2000, "ResetStartofGameVariables'"
  vpmtimer.addtimer 3000, "StopPlanets'"
  vpmtimer.addtimer 4000, "RobotVisible'"
  vpmtimer.addtimer 5000, "TurnOnStartOfGameLights'"
  Light_WheelAward.State=0:l37a.State=0
  ChaosCriticalReady(CurrentPlayer)=False
  ChaosCriticalComplete(CurrentPlayer)=False
  SuperLoopsReady(CurrentPlayer)=False
  SuperLoopsComplete(CurrentPlayer)=False
  ReverseLock1Ready(CurrentPlayer)=False
  ReverseLock2Ready(CurrentPlayer)=False
  ReverseLock1Complete(CurrentPlayer)=False
  ReverseLock2Complete(CurrentPlayer)=False
  ReverseTimeComplete(CurrentPlayer)=False
  WormLocksActivated(CurrentPlayer)=False
  WormLocksComplete(CurrentPlayer)=False
  DragonWounded1(CurrentPlayer)=False
  DragonWounded2(CurrentPlayer)=False
  DragonWounded3(CurrentPlayer)=False
  DragonWounded4(CurrentPlayer)=False
  PlaySong "m_Cassandra_s-Waltz"
End Sub

'This Sub is setting up the current player for next ball and stopping active modes from interfering for the reset for new player ball
'It is only called if the ball drains while playing the glory Table1

Sub GloryStopAtEndOfBallComplete
    'Ball Drains Game in playGame in play
  GloryMultiBallTimer.Enabled=False
debug.print "ForceStopGloryMultiballTimer"
  GloryComplete(CurrentPlayer)=False
  GloryTableOn(CurrentPlayer)=True
  GloryTableStarted(CurrentPlayer)=False
debug.print "GloryTableOn=True-GloryStopEOBC"

  lGlory1.State=0:lGlory2.State=0:lGlory3.State=0:lGlory4.State=0:lGlory5.State=0
  FlasherGloryTable.Visible=0
  StopClockGlory
  StopClockGloryTableDragon
  StopClockGloryBig
  FlasherGlory.Visible=0
  EndOfGloryTable=True
  spindiscimg.visible=True
  l38.State=0:l39.State=0:l40.State=0:l41.State=0:l42.State=0:l43.State=0:l44.State=0:l45.State=0
' GloryResetTimer.Enabled=True

  ResetStartofGameVariables
debug.print "ResetStartofGameVariablesandLights in GloryStopAEOB"
  vpmtimer.addtimer 200, "GiOff:TurnOnStartOfGameLights'"
  vpmtimer.addtimer 800, "SaveLightStates'"
  StopPlanets
' vpmtimer.addtimer 4000, "RobotVisible'"
  Light_WheelAward.State=0:l37a.State=0
' PlaySong "m_Cassandra_s-Waltz"
  ChaosCriticalReady(CurrentPlayer)=False
  ChaosCriticalComplete(CurrentPlayer)=False
  SuperLoopsReady(CurrentPlayer)=False
  SuperLoopsComplete(CurrentPlayer)=False
  ReverseLock1Ready(CurrentPlayer)=False
  ReverseLock2Ready(CurrentPlayer)=False
  ReverseLock1Complete(CurrentPlayer)=False
  ReverseLock2Complete(CurrentPlayer)=False
  ReverseTimeComplete(CurrentPlayer)=False
  WormLocksActivated(CurrentPlayer)=False
  WormLocksComplete(CurrentPlayer)=False
  DragonWounded1(CurrentPlayer)=False
  DragonWounded2(CurrentPlayer)=False
  DragonWounded3(CurrentPlayer)=False
  DragonWounded4(CurrentPlayer)=False
  ReverseLock1Ready(CurrentPlayer)=False
  StopSuperLoops
  StopSuperJets
  StopWormLights
  StopDragon
  LowerRampL
  KickerReverseLock1.Enabled=False
  KickerReverseLock2.Enabled=False
  ReverseLock2Ready(CurrentPlayer)=False
  KickerReverseLock1.Enabled=False
  KickerReverseLock2.Enabled=False
  ResetReverseTimeLights
  If DragonUp=True Then DragonMoveDown
  If SpiderUp=True Then SpiderMoveDown
  If BirdUp=True Then BirdMoveDown
End Sub

Sub RobotVisible

  BG_Sequence 34,0,100
  BG_Sequence 35,0,100
  BG_Sequence 36,0,100
  Robot.Visible=True
  TimeThief.Visible=False
  QMan.Visible=False
' PlaySound "CO_KeyToTime"
  RobotCallout
End Sub

Sub RobotInvisible
  Robot.Visible=True
End Sub

Sub RobotCallout
  PlaySound "CO_SavingTheUniverse"
End Sub

Sub GameReset
  PlaySong "m_Cassandra_s-Waltz"
  l38.State=0:l39.State=0:l40.State=0:l41.State=0:l42.State=0:l43.State=0:l44.State=0:l45.State=0
' GloryResetTimer.Enabled=True
  vpmtimer.addtimer 2000, "ResetModesAndAwards'"
  If  GloryComplete(CurrentPlayer)=True Then vpmtimer.addtimer 4000, "RobotVisible'": End If
  vpmtimer.addtimer 5000, "TurnOnStartOfGameLights'"
  vpmtimer.addtimer 5000, "StopPlanets'"
  vpmtimer.addtimer 2000, "GiOff'"
  Light_WheelAward.State=0:l37a.State=0
  '*********************
  TimeLordComplete=False
  '*********************
End Sub


Sub HoldExtraBall
  LightShootAgain.State=1
  bExtraBallWonThisBall=True
End Sub


'SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS_SPINNER2

Dim LeftSpins


'SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSPINNERLEFT
Sub SpinnerLeft_Spin
  DOF 123,DOFPulse
  SoundSpinner SpinnerLeft
  LeftSpins=LeftSpins+1

  If Not Tilted Then
    CountLeftSpins
  End If
End Sub





Sub CountLeftSpins
' If GloryTableTimer.Enabled=1 Then: Exit Sub
  LeftSpins=LeftSpins+1

  If LeftSpins>=180 Then
    l72.State = 1:
    vpmTimer.AddTimer 500, "LeftStarsBlink'"
    DMD CL(0, "1000000"), CL(1, "AAAAW SHINY STARS"), "", eNone, eBlink, eNone, 1000, True, ""


    vpmtimer.addtimer 1500, "DMDScoreNow'"
    AddScore 1000000
    LeftSpins=0
  End If
  If LeftSpins>=160 Then: l71.State = 1:AddScore 20000:End If
  If LeftSpins>=140 Then: l70.State = 1:End If
  If LeftSpins>=120 Then: l69.State = 1:AddScore 20000:End If
  If LeftSpins>=100 Then: l68.State = 1:End If
  If LeftSpins>=80 Then: l67.State = 1:AddScore 20000:End If
  If LeftSpins>=60 Then: l66.State = 1:End If
  If LeftSpins>=40 Then: l65.State = 1:AddScore 20000:End If
  If LeftSpins>=40 Then: l64.State = 1:End If
  If LeftSpins>=20 Then: l63.State = 1:AddScore 20000:End If

End Sub


Sub LeftStarsBlink
  l63.State = 2:l64.State = 2: l65.State = 2: l66.State = 2: l67.State = 2: l68.State = 2: l69.State = 2: l70.State = 2: l71.State = 2: l72.State = 2
  vpmtimer.addtimer 2800, "ResetLeftSpinLights'"
End Sub

Sub ResetLeftSpinLights
  l63.State = 0:l64.State = 0: l65.State = 0: l66.State = 0: l67.State = 0: l68.State = 0: l69.State = 0: l70.State = 0: l71.State = 0: l72.State = 0
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'DRAGON EYE Right
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'
Dim eye2wiggle:eye2wiggle=0
Dim eye2pos:eye2pos=6
Dim currenteye2pos:currenteye2pos=6


' smooth eyeball mod

DIM eye2MoveToX : eye2MoveToX=0
DIM eye2MoveToY : eye2MoveToY=0
Dim eye2CurX    : eye2CurX=0
Dim eye2CurY    : eye2CurY=0
Dim eye2Moving  : eye2Moving=false
Dim curBall2X   : curBall2X=0
Dim curBall2Y   : curBall2Y=0


Function GetEye2Angle()
  ' Given y and x coords returns atan2
  ' by Jim Deutch, Syracuse, New York
  Dim ys
  Dim xs
  Dim theta, pi

  xs=eye2ballmod.x-curBallx
  ys=eye2ballmod.y-curBally
  pi = 3.1415926535897932384626433832795
  If xs <> 0 Then
    theta = Atn(ys / xs)
    If xs < 0 Then
      theta = theta + pi
    End If
  Else
    If ys < 0 Then
      theta = 3 * pi / 2 '90
    Else
      theta = pi / 2 '270
    End If
  End If
  GetEye2Angle = theta * 180 / pi
  if GetEye2Angle<0 then GetEye2Angle=360+GetEye2Angle
End Function

Sub checkeye2ball
  DIM xmov
  DIM ymov
  DIM cAngle
  If (bMultiBallMode = true) then exit sub
  if (eyespin>0) then exit Sub
  cAngle=GetEye2Angle  * 3.141592 / 180

  xmov=sin(cAngle) * 50.00
  ymov=cos(cAngle) * -50.00

  if xmov>40 then xmov=40

  eye2MoveToX =  round(xmov)
  eye2MoveToY =  round(ymov)

  eye2CurX=eye2ballmod.objrotx
  eye2CurY=eye2ballmod.objroty

  if (eye2CurX<>eyeMoveToX) OR (eye2CurY<>eye2MoveToY) Then
    eye2Moving=true
    if spinningeye2.interval<>eyeFollowS  Then spinningeye2.interval=eyeFollowS
    spinningeye2.enabled=1
  end if
end Sub

spinningeye2.enabled = 0
dim eye2spin:eye2spin = 0


sub spinningeye2_timer
  Dim mAmount
  if (eye2Moving) and (eye2spin=0) Then
    if (eye2CurX>eye2MoveToX) Then
      eye2ballmod.objrotx=eye2ballmod.objrotx-1
      eye2Curx=eye2Curx-1
    elseif (eye2CurX<eye2MoveToX) Then
      eye2ballmod.objrotx=eye2ballmod.objrotx+1
      eye2CurX=eye2CurX+1
    end if
    if (eye2CurY>eye2MoveToY) Then
      eye2ballmod.objroty=eye2ballmod.objroty-1
      eye2CurY=eye2CurY-1
    elseif (eye2CurY<eye2MoveToY) Then
      eye2ballmod.objroty=eye2ballmod.objroty+1
      eye2CurY=eye2CurY+1
    end if
    if (eye2CurX=eye2MoveToX) AND (eye2CurY=eye2MoveToY) Then
      eye2Moving=false
      spinningeye2.enabled=0
      spinningeye2.interval= 1  'return to normal
    end if

  Else 'if not moving eye to position
    if spinningeye2.interval<>1  Then spinningeye2.interval=1
    eye2wiggle = 1
    eye2spin = eye2spin + 1
    eye2ballmod.objrotx = eye2ballmod.objrotx  + 1
    eye2ballmod.objroty = eye2ballmod.objroty - 1
    if eye2spin>1500 Then
      eye2wiggle = 0
      eye2spin = 0
      eye2Moving=false
      currenteye2pos=-1
      eye2pos=1
      eye2ballmod.objrotx=eyeCurX
      eye2ballmod.objroty=eyeCurY
      eye2MoveToX=eyeCurX
      eye2MoveToY=eyeCurY
      curBallx=0
      curBally=0
      spinningeye2.enabled = 0
      checkeye2ball

    End if
  End if
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'DRAGON EYE left
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'
Dim eyewiggle:eyewiggle=0
Dim eyepos:eyepos=6
Dim currenteyepos:currenteyepos=6


' smooth eyeball mod

DIM eyeMoveToX : eyeMoveToX=0
DIM eyeMoveToY : eyeMoveToY=0
Dim eyeCurX    : eyeCurX=0
Dim eyeCurY    : eyeCurY=0
Dim eyeMoving  : eyeMoving=false
Dim curBallX   : curBallX=0
Dim curBallY   : curBallY=0


Function GetEyeAngle()
  ' Given y and x coords returns atan2
  ' by Jim Deutch, Syracuse, New York
  Dim ys
  Dim xs
  Dim theta, pi

  xs=eyeballmod.x-curBallx
  ys=eyeballmod.y-curBally
  pi = 3.1415926535897932384626433832795
  If xs <> 0 Then
    theta = Atn(ys / xs)
    If xs < 0 Then
      theta = theta + pi
    End If
  Else
    If ys < 0 Then
      theta = 3 * pi / 2 '90
    Else
      theta = pi / 2 '270
    End If
  End If
  GetEyeAngle = theta * 180 / pi
  if GetEyeAngle<0 then GetEyeAngle=360+GetEyeAngle
End Function

Sub checkeyeball
  DIM xmov
  DIM ymov
  DIM cAngle
  If (bMultiBallMode = true) then exit sub
  if (eyespin>0) then exit Sub
  cAngle=GetEyeAngle  * 3.141592 / 180

  xmov=sin(cAngle) * 50.00
  ymov=cos(cAngle) * -50.00

  if xmov>40 then xmov=40

  eyeMoveToX =  round(xmov)
  eyeMoveToY =  round(ymov)

  eyeCurX=eyeballmod.objrotx
  eyeCurY=eyeballmod.objroty

  if (eyeCurX<>eyeMoveToX) OR (eyeCurY<>eyeMoveToY) Then
    eyeMoving=true
    if spinningeye.interval<>eyeFollowS  Then spinningeye.interval=eyeFollowS
    spinningeye.enabled=1
  end if
end Sub

spinningeye.enabled = 0
dim eyespin:eyespin = 0


sub spinningeye_timer
  Dim mAmount
  if (eyeMoving) and (eyespin=0) Then
    if (eyeCurX>eyeMoveToX) Then
      eyeballmod.objrotx=eyeballmod.objrotx-1
      eyeCurx=eyeCurx-1
    elseif (eyeCurX<eyeMoveToX) Then
      eyeballmod.objrotx=eyeballmod.objrotx+1
      eyeCurX=eyeCurX+1
    end if
    if (eyeCurY>eyeMoveToY) Then
      eyeballmod.objroty=eyeballmod.objroty-1
      eyeCurY=eyeCurY-1
    elseif (eyeCurY<eyeMoveToY) Then
      eyeballmod.objroty=eyeballmod.objroty+1
      eyeCurY=eyeCurY+1
    end if
    if (eyeCurX=eyeMoveToX) AND (eyeCurY=eyeMoveToY) Then
      eyeMoving=false
      spinningeye.enabled=0
      spinningeye.interval= 1  'return to normal
    end if

  Else 'if not moving eye to position
    if spinningeye.interval<>1  Then spinningeye.interval=1
    eyewiggle = 1
    eyespin = eyespin + 1
    eyeballmod.objrotx = eyeballmod.objrotx  + 1
    eyeballmod.objroty = eyeballmod.objroty - 1
    if eyeSpin>1500 Then
      eyewiggle = 0
      eyespin = 0
      eyemoving=false
      currenteyepos=-1
      eyepos=1
      eyeballmod.objrotx=eyeCurX
      eyeballmod.objroty=eyeCurY
      eyeMoveToX=eyeCurX
      eyeMoveToY=eyeCurY
      curBallx=0
      curBally=0
      spinningeye.enabled = 0
      checkeyeball

    End if
  End if
End Sub


'********************
' Hologram animation
'********************

Dim Holo1Pos, Holo1Dir, Holo1Step, Holo1Random, Holo1Angle

Sub Hologram1Start()
  Hologram1.visible = 1
  Holo1Pos = 0
  Holo1Dir = 1
  Holo1Step = 1
  Holo1Random = 0
  Holo1Angle = 0
  Hologram1.Image = "hollogramqman" '&INT(RND * 7 + 1)
  Hologram1Timer.Enabled = 1
  Hologram1TimerExpired.Enabled = 1
End Sub

Sub Hologram2Start()
  Hologram1.visible = 1
  Holo1Pos = 0
  Holo1Dir = 1
  Holo1Step = 1
  Holo1Random = 0
  Holo1Angle = 0
  Hologram1.Image = "hollogramtimelord" '&INT(RND * 7 + 1)
  Hologram1Timer.Enabled = 1
  Hologram1TimerExpired.Enabled = 1
End Sub


Sub Hologram1Stop
  Hologram1.visible = 1
  Hologram1Timer.Enabled = 0
  Hologram1TimerExpired.Enabled = 0
End Sub



Sub Hologram1Timer_Timer()
  Dim radianes
  Holo1Pos = Holo1Pos + Holo1Dir
  If Holo1Pos < -25 Then Holo1Dir = 1:Holo1Random = Holo1Random + 1
  If Holo1Pos > 25 Then Holo1Dir = -1
  If Holo1Pos = 0 AND Holo1Random = 5 Then
    Holo1Step = INT(4 * RND)
    Holo1Random = 0
    Hologram1.TransX = 0
    Hologram1.TransY = 0
    Hologram1.Rotz = 0
    Hologram1.Roty = 0
  End If
  Select Case Holo1Step
    Case 0 'zoom
      Hologram1.Size_X = HoloGram1.Size_X + Holo1Dir
      Hologram1.Size_Y = Hologram1.Size_Y + Holo1Dir
    Case 1 'Rotate Y axis
      Hologram1.Rotz = Holo1pos / 2
    Case 2 'rotation circle
      radianes = Holo1Angle * (3.1415 / 180)
      Hologram1.TransX = 10 * Cos(radianes)
      Hologram1.TransY = 10 * Sin(radianes)
      Holo1Angle = (Holo1Angle + 1)MOD 360
    Case 3 'Rotate Z axis
      Hologram1.RotZ = Holo1Pos
  End Select
End Sub

Sub Hologram1TimerExpired_Timer()
  Hologram1Stop
End Sub

'***********************
'StartPlanets
'***********************

'************************
'  StartPlanetDragon1
'************************
Dim PlanetDragonStarted
Sub StartPlanetDragon1
  PlanetDragonStarted=1
  DragonPlanet1.visible = 1
  DragonPlanetLight.visible = 1
  DragonPlanetTimer.Enabled = 1
  l_PlanetDragonEnabled.State=1
' l42.State=1:l43.State=1:l44.State=1:l45.State=1
End Sub

Sub StopPlanetDragon1
  PlanetDragonStarted=0
  DragonPlanet1.visible = 0
  DragonPlanetLight.visible = 0
  DragonPlanetTimer.Enabled = 0
  l_PlanetDragonEnabled.State=0
  l42.State=0:l43.State=0:l44.State=0:l45.State=0
End Sub

Sub DragonPlanetTimer_Timer
  DragonPlanet1.rotz = (DragonPlanet1.rotz - 1.5)mod 360
  DragonPlanetLight.rotz = (DragonPlanetLight.rotz - 1.5)mod 360
End Sub


'**********************
'StartPlanetChaos
'**********************
Dim PlanetChaosStarted
Sub StartPlanetChaos
  PlanetChaosStarted=1
  PlanetChaos.visible = 1
  PlanetChaosTimer.Enabled = 1
  l_PlanetChaosEnabled.State=1
  l38.State=1
End Sub

Sub StopPlanetChaos
  PlanetChaosStarted=0
  PlanetChaos.visible = 0
  PlanetChaosTimer.Enabled = 0
  l_PlanetChaosEnabled.State=0
  l38.State=0

End Sub

Sub PlanetChaosTimer_Timer
  PlanetChaos.rotz = (PlanetChaos.rotz -2)mod 360
End Sub


Sub PlanetChaosStopTimer_Timer
  StopPlanetChaos
End Sub

'**********************
'StartPlanetReverseTime
'**********************
Dim PlanetReverseTimeStarted
Sub StartPlanetReverseTime
  PlanetReverseTimeStarted=1
  PlanetReverseTime.visible = 1
  PlanetReverseTimeTimer.Enabled = 1
  l_PlanetReversEnabled.State=1
  l39.State=1
End Sub

Sub StopPlanetReverseTime
  PlanetReverseTimeStarted=0
  PlanetReverseTime.visible = 0
  PlanetReverseTimeTimer.Enabled = 0
  l_PlanetReversEnabled.State=0
  l39.State=0
End Sub

Sub PlanetReverseTimeTimer_Timer
  PlanetReverseTime.rotz = (PlanetReverseTime.rotz +1)mod 360
End Sub


Sub PlanetReverseTimeStopTimer_Timer
  StopPlanetReverseTime
End Sub

'**********************
'StartPlanetSpider
'**********************
Dim PlanetSpiderStarted
Sub StartPlanetSpider
  BG_Sequence 38,0,200
  BG_Sequence 1,0,20
  BG_Sequence 5,4,0

  PlanetSpiderStarted=1
  PlanetSpider.visible = 1
  PlanetSpiderTimer.Enabled = 1
  l_PlanetSpiderEnabled.State=1
  l40.State=1
End Sub

Sub StopPlanetSpider
  PlanetSpiderStarted=0
  PlanetSpider.visible = 0
  PlanetSpiderTimer.Enabled = 0
  l_PlanetSpiderEnabled.State=0
  l40.State=0
End Sub

Sub PlanetSpiderTimer_Timer
  PlanetSpider.rotz = (PlanetSpider.rotz -1.7)mod 360
End Sub


Sub PlanetSpiderStopTimer_Timer
  StopPlanetSpider
End Sub

'**********************
'StartPlanetBird
'**********************
Dim PlanetBirdStarted
Sub StartPlanetBird
  BG_Sequence 38,0,200
  BG_Sequence 1,0,20
  BG_Sequence 5,4,0

  PlanetBirdStarted=1
  PlanetBird.visible = 1
  PlanetBirdTimer.Enabled = 1
  l_PlanetBirdEnabled.State=1
  l41.State=1
End Sub

Sub StopPlanetBird
  PlanetBirdStarted=0
  PlanetBird.visible = 0
  PlanetBirdTimer.Enabled = 0
  l_PlanetBirdEnabled.State=0
  l41.State=0
End Sub

Sub PlanetBirdTimer_Timer
  PlanetBird.rotz = (PlanetBird.rotz -2)mod 360
End Sub


Sub PlanetBirdStopTimer_Timer
  StopPlanetBird
End Sub

Sub StopPlanets
  StopPlanetChaos
  StopPlanetReverseTime
  StopPlanetSpider
  StopPlanetBird
  StopPlanetDragon1
End Sub


Sub QManBallTimer_Timer
  PlanetQMan1.rotz = (PlanetQMan1.rotz -2)mod 360
  PlanetQMan2.rotz = (PlanetQMan2.rotz -2)mod 360
  PlanetQMan3.rotz = (PlanetQMan3.rotz -2)mod 360
  PlanetQMan4.rotz = (PlanetQMan4.rotz -2)mod 360
  PlanetQMan5.rotz = (PlanetQMan5.rotz -2)mod 360
End Sub









'**********************
'Clocks Hands Rotation
'**********************

Sub StartClockHandsClockwise
  StartClockHandsCW
  StartClock2HandsCW
  StartClockApronCW
End Sub

Sub StopClockHandsClockwise
  StopDragonPlanetCW
  StopReverseTimeCW
  StopClockApronCW

End Sub


Sub StartClockHandsCounterClockwise
  StartClockHandsCCW
  StartClock2HandsCCW
  StartClockApronCCW
End Sub

Sub StopClockHandsCounterClockwise
  StopDragonPlanetCCW
  StopReverseTimeCCW
  StopClockApronCCW
End Sub

'********************************************************
'BigClockHands
Sub StartClockHandsCW
  ClockHandTimer.Enabled = 1
' SmallHand1Timer.Enabled=1
End Sub

Sub StopDragonPlanetCW
  ClockHandTimer.Enabled = 0
  SmallHand1Timer.Enabled=0
End Sub


dim smalhand : smalhand = 1450
Sub ClockHandTimer_Timer
' ClockHands.rotz = (ClockHands.rotz + 1)MOD 360
  pLightClockBroken001.roty = (pLightClockBroken001.roty + 1)MOD 360
  smalhand = (smalhand + 1) mod 4320
  pLightClockBroken002.roty = smalhand / 12
End Sub

Sub SmallHand1Timer_Timer
' ClockHandSmall.rotz = (ClockHandSmall.rotz + 0.165)MOD 360
' pLightClockBroken002.roty = (pLightClockBroken002.roty + 0.165)MOD 360

End Sub
'_________________________________________
'BigClockHandsReverse
Sub StartClockHandsCCW
  ClockHandTimerR1.Enabled = 1
' SmallHand1TimerR1.Enabled=1
End Sub

Sub StopDragonPlanetCCW
  ClockHandTimerR1.Enabled = 0
  SmallHand1Timerr1.Enabled=0
End Sub

Sub ClockHandTimerR1_Timer
' ClockHands.rotz = (ClockHands.rotz - 4)MOD 360

  pLightClockBroken001.roty = (pLightClockBroken001.roty - 1)MOD 360
  smalhand = (smalhand - 1) mod 4320
  pLightClockBroken002.roty = smalhand / 12

End Sub

Sub SmallHand1TimerR1_Timer
' ClockHandSmall.rotz = (ClockHandSmall.rotz + 0.66)MOD 360
End Sub

'**********************************
Sub StartClock2HandsCW
  ClockHand2Timer.Enabled = 1
  SmallHand2Timer.Enabled = 1
  minipocketwatch = 1
End Sub

Sub StopReverseTimeCW
  ClockHand2Timer.Enabled = 0
  SmallHand2Timer.Enabled = 0
  minipocketwatch = 0
End Sub

Sub ClockHand2Timer_Timer
  ClockHands2.rotz = (ClockHands2.rotz + 1)MOD 360
End Sub

dim smalhand2
Sub SmallHand2Timer_Timer
  smalhand2 = (smalhand2 + 1) mod 4320
  ClockHand2Small.rotz = smalhand / 12
End Sub


'_____________________________________________________
dim minipocketwatch
Sub StartClock2HandsCCW
  ClockHand2TimerR1.Enabled = 1
  SmallHand2TimerR1.Enabled=1
  minipocketwatch = 2
End Sub

Sub StopReverseTimeCCW
  ClockHand2TimerR1.Enabled = 0
  SmallHand2TimerR1.Enabled=0
  minipocketwatch = 0
End Sub

Sub ClockHand2TimerR1_Timer
  ClockHands2.rotz = (ClockHands2.rotz - 2)MOD 360
End Sub

Sub SmallHand2TimerR1_Timer
  smalhand2 = (smalhand2 - 1) mod 4320
  ClockHand2Small.rotz = smalhand / 12
' ClockHand2Small.rotz = (ClockHand2Small.rotz - 0.32)MOD 360
End Sub





'**********************************


'**********************************
'**********************
'Clocks Hands Apron Rotation
'**********************
'Apron Hands CW
Sub StartClockApronCW
  ClockHandBigApronTimer.Enabled = 1
  ClockHandSmallApronTimer.Enabled=1
End Sub

Sub StopClockApronCW
  ClockHandBigApronTimer.Enabled = 0
  ClockHandSmallApronTimer.Enabled=0
End Sub

Sub ClockHandBigApronTimer_Timer
  ClockHandBigApron.rotz = (ClockHandBigApron.rotz +2)MOD 360
End Sub

Sub ClockHandSmallApronTimer_Timer
  ClockHandSmallApron.rotz = (ClockHandSmallApron.rotz +0.32)MOD 360
End Sub

'_______________________________________________________

'Apron Hands CCW
Sub StartClockApronCCW
  ClockHandBigApronRTimer.Enabled = 1
  ClockHandSmallApronRTimer.Enabled=1
End Sub

Sub StopClockApronCCW
  ClockHandBigApronRTimer.Enabled = 0
  ClockHandSmallApronRTimer.Enabled=0
End Sub

Sub ClockHandBigApronRTimer_Timer
  ClockHandBigApron.rotz = (ClockHandBigApron.rotz -2)MOD 360
End Sub

Sub ClockHandSmallApronRTimer_Timer
  ClockHandSmallApron.rotz = (ClockHandSmallApron.rotz -0.32)MOD 360
End Sub
'_________________________________________________________________
'ClockGloryTableDragon

Sub ClockGloryTableDragon
  ClockHandCenterDragonBHTimer.Enabled = 1
  ClockHandCenterDragonSHTimer.Enabled=1
  ClockGloryDragonBigHand1.Visible=1
  ClockGloryDragonSmallHand1.Visible=1
End Sub

Sub StopClockGloryTableDragon
  ClockHandCenterDragonBHTimer.Enabled = 0
  ClockHandCenterDragonSHTimer.Enabled=0
  ClockGloryDragonBigHand1.Visible=0
  ClockGloryDragonSmallHand1.Visible=0
End Sub

Sub ClockHandCenterDragonBHTimer_Timer
  ClockGloryDragonBigHand1.rotz = (ClockGloryDragonBigHand1.rotz +8)MOD 360
End Sub

Sub ClockHandCenterDragonSHTimer_Timer
  ClockGloryDragonSmallHand1.rotz = (ClockGloryDragonSmallHand1.rotz -1.28)MOD 360
End Sub

'-------------------------------------------------------------------
'ClockGlory

Sub ClockGlory
  ClockHandGloryBHTimer.Enabled = 1
  ClockHandGlorySHTimer.Enabled=1
  ClockCentreBigHand1.Visible=1
  ClockCentreSmallHand1.Visible=1
End Sub

Sub StopClockGlory
  ClockHandGloryBHTimer.Enabled = 0
  ClockHandGlorySHTimer.Enabled=0
  ClockCentreBigHand1.Visible=0
  ClockCentreSmallHand1.Visible=0
End Sub

Sub ClockHandGloryBHTimer_Timer
  ClockCentreBigHand1.rotz = (ClockCentreBigHand1.rotz +4)MOD 360
End Sub

Sub ClockHandGlorySHTimer_Timer
  ClockCentreSmallHand1.rotz = (ClockCentreSmallHand1.rotz +0.64)MOD 360
End Sub

'-------------------------------------------------------------------
'ClockGloryBig

Sub ClockGloryBig
  ClockHandGloryBigBHTimer.Enabled = 1
  ClockHandGloryBigSHTimer.Enabled=1
  ClockHandGloryBig.Visible=1
  ClockHandGlorySmall.Visible=1
End Sub

Sub StopClockGloryBig
  ClockHandGloryBigBHTimer.Enabled = 0
  ClockHandGloryBigSHTimer.Enabled=0
  ClockHandGloryBig.Visible=0
  ClockHandGlorySmall.Visible=0
End Sub

Sub ClockHandGloryBigBHTimer_Timer
  ClockHandGloryBig.rotz = (ClockCentreBigHand1.rotz +4)MOD 360
End Sub

Sub ClockHandGloryBigSHTimer_Timer
  ClockHandGlorySmall.rotz = (ClockCentreSmallHand1.rotz +0.64)MOD 360
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'End of scrypt
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


pLightLeftEscapeoff.blenddisablelighting = 1
pLightrightescapeoff.blenddisablelighting = 1

p21Boff.blenddisablelighting = 4
p21off.blenddisablelighting = 4
p50off.blenddisablelighting = 4
p61off.blenddisablelighting = 4
p18off.blenddisablelighting = 4
p20off.blenddisablelighting = 4
p49off.blenddisablelighting = 4
p20Aoff.blenddisablelighting = 4
p19off.blenddisablelighting = 4
p48off.blenddisablelighting = 4
p61aoff.blenddisablelighting = 4
p16off.blenddisablelighting = 4
p19aoff.blenddisablelighting = 4
p11off.blenddisablelighting = 4 ' left red triangles
p12off.blenddisablelighting = 4 ' left red triangles
p13off.blenddisablelighting = 4 ' left red triangles
p75off.blenddisablelighting = 4 '
p4off.blenddisablelighting = 4 ' yellow tri

p23aoff.blenddisablelighting = 4  ' Wheel
p23boff.blenddisablelighting = 4  ' Wheel
p23coff.blenddisablelighting = 4  ' Wheel
p23doff.blenddisablelighting = 4  ' Wheel
p23eoff.blenddisablelighting = 4  ' Wheel
p23foff.blenddisablelighting = 4  ' Wheel
p23goff.blenddisablelighting = 4  ' Wheel
p23hoff.blenddisablelighting = 4  ' Wheel




p17off.blenddisablelighting = 4  ' test
p61Boff.blenddisablelighting = 4  ' test
p16aoff.blenddisablelighting = 4
p17Aoff.blenddisablelighting = 4
p17Boff.blenddisablelighting = 4
p1off.blenddisablelighting = 4
p2off.blenddisablelighting = 4
p3off.blenddisablelighting = 4
p5off.blenddisablelighting = 7
p6off.blenddisablelighting = 7
p7off.blenddisablelighting = 7
p8off.blenddisablelighting = 7
p9off.blenddisablelighting = 7
p10off.blenddisablelighting = 7
p28a6.blenddisablelighting = 4
p28a5.blenddisablelighting = 4
p28a4.blenddisablelighting = 4



p72off.blenddisablelighting = 4 ' whitestar
p71off.blenddisablelighting = 4 ' whitestar
p70off.blenddisablelighting = 4 ' whitestar
p69off.blenddisablelighting = 4 ' whitestar
p68off.blenddisablelighting = 4 ' whitestar
p67off.blenddisablelighting = 4 ' whitestar
p66off.blenddisablelighting = 4 ' whitestar
p65off.blenddisablelighting = 4 ' whitestar
p64off.blenddisablelighting = 4 ' whitestar
p63off.blenddisablelighting = 4 ' whitestar
p89off.blenddisablelighting = 4 ' whitestar
p88off.blenddisablelighting = 4 ' whitestar
p87off.blenddisablelighting = 4 ' whitestar
p86off.blenddisablelighting = 4 ' whitestar
p85off.blenddisablelighting = 4 ' whitestar
p84off.blenddisablelighting = 4 ' whitestar
p83off.blenddisablelighting = 4 ' whitestar
p82off.blenddisablelighting = 4 ' whitestar
p81off.blenddisablelighting = 4 ' whitestar
p80off.blenddisablelighting = 4 ' whitestar
p62Aoff.blenddisablelighting = 4 ' whitestar top
p62boff.blenddisablelighting = 4 ' whitestar top
p62coff.blenddisablelighting = 4 ' whitestar top
p_wl1off.blenddisablelighting = 6 ' bluestar top
p_wl2off.blenddisablelighting = 6 ' bluestar top
p_wl3off.blenddisablelighting = 6 ' bluestar top
p_WormKicker1enabledoff.blenddisablelighting = 6 ' WormKicker1
p_WormKicker2enabledoff.blenddisablelighting = 6 ' WormKicker2
p_WormKicker3enabledoff.blenddisablelighting = 6 ' WormKicker3
p_WormKicker4enabledoff.blenddisablelighting = 6 ' WormKicker4
p_ChaosKickerenabledoff.blenddisablelighting = 6 ' WormKicker4

p36off.blenddisablelighting = 10  ' AdvanceTime
p37off.blenddisablelighting = 10  ' Extraball
p38off.blenddisablelighting = 10  ' Lord
p39off.blenddisablelighting = 10  ' Lord
p40off.blenddisablelighting = 10  ' Lord
p41off.blenddisablelighting = 10  ' Lord
p42off.blenddisablelighting = 10  ' Time
p43off.blenddisablelighting = 10  ' Time
p44off.blenddisablelighting = 10  ' Time
p45off.blenddisablelighting = 10  ' Time


p30off.blenddisablelighting = 1 ' plungerlane
p31off.blenddisablelighting = 1 ' plungerlane
p32off.blenddisablelighting = 1 ' plungerlane
p33off.blenddisablelighting = 1 ' plungerlane
p34off.blenddisablelighting = 1 ' plungerlane
p35off.blenddisablelighting = 1 ' plungerlane
p_PlanetDragonEnabledoff.blenddisablelighting = 1 ' planets middle
p_PlanetSpiderEnabledoff.blenddisablelighting = 1 ' planets middle
p_PlanetBirdEnabledoff.blenddisablelighting = 1 ' planets middle
p_PlanetReversEnabledoff.blenddisablelighting = 1 ' planets middle
p_PlanetChaosEnabledoff.blenddisablelighting = 1 ' planets middle
pLightBallLock1off.blenddisablelighting = 1 ' balllock 1
pLightBallLock2off.blenddisablelighting = 1 ' balllock 2
p28a3off.blenddisablelighting = 1 '
p28a2off.blenddisablelighting = 1 '
p28a1off.blenddisablelighting = 1 '


pLightLeftInlaneoff.blenddisablelighting = 2  ' test
pLightRightInlane.blenddisablelighting = 2

pLightShootAgainoff.blenddisablelighting = 10
pLightClockBrokenoff.blenddisablelighting = 4
pLight_BallSaveroff.blenddisablelighting = 1
pLight_easyhardoff.blenddisablelighting = 1
pLight_WheelAwardoff.blenddisablelighting = 4
p_WheelKickerEnabledoff.blenddisablelighting = 4
p37aoff.blenddisablelighting = 3
p28off.blenddisablelighting = 2
'//////////////////////////////////////////////////////////////////////
'// TIMERS
'//////////////////////////////////////////////////////////////////////


' The game timer interval is 10 ms
Dim zoomeye
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
  RollingUpdate

  'pocketwatch mini
  If minipocketwatch = 1 Then
    ClockHand2Small001.rotz = ( ClockHand2Small001.rotz +6) mod 360
  elseif minipocketwatch = 2 Then
    ClockHand2Small001.rotz = ( ClockHand2Small001.rotz -6 ) mod 360
  End If

  ' gi updates



  dim L,X
  L = gi_laneguide1.GetInPlayIntensity
  For each x in Pegs
    x.blenddisablelighting = 0.067 * L
  Next
  TopPlastics.blenddisablelighting = L * 1.8 ' TopPlastics
  ApronPlastic.blenddisablelighting = L * 0.6
  LeftSlingCover.blenddisablelighting = Gi4.GetInPlayIntensity * 0.5'LeftSlingCover
  Peg5.blenddisablelighting = 0.067 * Gi4.GetInPlayIntensity
  Peg6.blenddisablelighting = 0.067 * Gi4.GetInPlayIntensity
  Peg7.blenddisablelighting = 0.067 * Gi4.GetInPlayIntensity
  RightSlingCover.blenddisablelighting = Gi1.GetInPlayIntensity * 0.5'RightSlingCover
  Peg.blenddisablelighting = 0.067 * Gi1.GetInPlayIntensity
  Peg1.blenddisablelighting = 0.067 * Gi1.GetInPlayIntensity
  Peg2.blenddisablelighting = 0.067 * Gi1.GetInPlayIntensity



  Wall002.blenddisablelighting = L * 0.1 'LeftWall
  Wall35.blenddisablelighting = L * 0.1 'RightWall
  Wall34.blenddisablelighting = L * 0.25 'backWall
  mainRamps.blenddisablelighting = L * 1.75'TranslucentRamps
''  TableFlasher.opacity = ( 1.5 - L ) * 60
  swTarget1.blenddisablelighting = L * 0.18
  swTarget2.blenddisablelighting = L * 0.18
  swTarget3.blenddisablelighting = L * 0.18
  TargetCaptiveBall.blenddisablelighting = L * 0.18
  sw3.blenddisablelighting = L * 0.02
  sw4.blenddisablelighting = L * 0.02
  sw5.blenddisablelighting = L * 0.02
  DragonLairWall.blenddisablelighting = L * 0.1
  Wall_Orbit.blenddisablelighting = L * 0.1
  WallWormLeft.blenddisablelighting = L * 0.1
  Wall_WormRight.blenddisablelighting = L * 0.1
  Wall_DragonLairRight.blenddisablelighting = L * 0.1
  Wall_DragonLairLeft.blenddisablelighting = L * 0.1
  Wall350.blenddisablelighting = L * 0.1
  Wall011.blenddisablelighting = L * 0.1
  Wall99.blenddisablelighting = L * 0.1
  Wall8.blenddisablelighting = L * 0.1
  Target1.blenddisablelighting = L * 0.1
  Target2.blenddisablelighting = L * 0.1
  Target3.blenddisablelighting = L * 0.1
  Target4.blenddisablelighting = L * 0.1
  Target5.blenddisablelighting = L * 0.1
  Target6.blenddisablelighting = L * 0.1

' light prims Update
  pLightLeftEscape.blenddisablelighting = LightLeftEscape.GetInPlayIntensity / 10
  pLightrightescape.blenddisablelighting = Lightrightescape.GetInPlayIntensity / 10
  eye2ballmod.blenddisablelighting = Lighteyeball2.GetInPlayIntensity + 0.1
  eyeballmod.blenddisablelighting = Lighteyeball1.GetInPlayIntensity + 0.1


  p21B.blenddisablelighting = 10 * l21B.GetInPlayIntensity
  p20A.blenddisablelighting = 10 * l20a.GetInPlayIntensity
  p19A.blenddisablelighting = 10 * l19a.GetInPlayIntensity
  p61.blenddisablelighting = 10 * l61.GetInPlayIntensity
  p21.blenddisablelighting = 10 * l21.GetInPlayIntensity
  p50.blenddisablelighting = 10 * l50.GetInPlayIntensity
  p18.blenddisablelighting = 10 * l18.GetInPlayIntensity
  p20.blenddisablelighting = 10 * l20.GetInPlayIntensity
  p48.blenddisablelighting = 6 * l48.GetInPlayIntensity
  p49.blenddisablelighting = 10 * l49.GetInPlayIntensity
  p19.blenddisablelighting = 6 * l19.GetInPlayIntensity
  p16.blenddisablelighting = 6 * l16.GetInPlayIntensity
  p11.blenddisablelighting = 4 * l11.GetInPlayIntensity ' left red triangles
  p12.blenddisablelighting = 4 * l12.GetInPlayIntensity ' left red triangles
  p13.blenddisablelighting = 4 * l13.GetInPlayIntensity ' left red triangles
  p21a.blenddisablelighting = 4 * l21a.GetInPlayIntensity ' left red triangles
  p24.blenddisablelighting = 4 * l24.GetInPlayIntensity ' left red triangles

  p75.blenddisablelighting = 4 * l75.GetInPlayIntensity ' yellow Tri
  p4.blenddisablelighting = 4 * l4.GetInPlayIntensity ' yellow Tri
  p23a.blenddisablelighting = 4 * l23a.GetInPlayIntensity ' Wheel
  p23b.blenddisablelighting = 4 * l23b.GetInPlayIntensity  ' Wheel
  p23c.blenddisablelighting = 4 * l23c.GetInPlayIntensity  ' Wheel
  p23d.blenddisablelighting = 4 * l23d.GetInPlayIntensity  ' Wheel
  p23e.blenddisablelighting = 4 * l23e.GetInPlayIntensity  ' Wheel
  p23f.blenddisablelighting = 4 * l23f.GetInPlayIntensity  ' Wheel
  p23g.blenddisablelighting = 4 * l23g.GetInPlayIntensity  ' Wheel
  p23h.blenddisablelighting = 4 * l23h.GetInPlayIntensity  ' Wheel




  p17.blenddisablelighting = 8 * l17.GetInPlayIntensity
  p61B.blenddisablelighting = 12 * l61B.GetInPlayIntensity
  p61a.blenddisablelighting = 6 * l61a.GetInPlayIntensity
  p16a.blenddisablelighting = 12 * l16a.GetInPlayIntensity
  p17a.blenddisablelighting = 12 * l17a.GetInPlayIntensity
  p17b.blenddisablelighting = 12 * l17b.GetInPlayIntensity
  p1.blenddisablelighting = 10 * l1.GetInPlayIntensity
  p2.blenddisablelighting = 10 * l2.GetInPlayIntensity
  p3.blenddisablelighting = 10 * l3.GetInPlayIntensity
  p5.blenddisablelighting = 19 * l5.GetInPlayIntensity
  p6.blenddisablelighting = 19 * l6.GetInPlayIntensity
  p7.blenddisablelighting = 19 * l7.GetInPlayIntensity
  p8.blenddisablelighting = 19 * l8.GetInPlayIntensity
  p9.blenddisablelighting = 19 * l9.GetInPlayIntensity
  p10.blenddisablelighting = 19 * l10.GetInPlayIntensity
  p28a4.blenddisablelighting = 19 * l28a4.GetInPlayIntensity
  p28a5.blenddisablelighting = 19 * l28a5.GetInPlayIntensity
  p28a6.blenddisablelighting = 19 * l28a6.GetInPlayIntensity


  p72.blenddisablelighting = 10 * l72.GetInPlayIntensity ' whitestars
  p71.blenddisablelighting = 10 * l71.GetInPlayIntensity ' whitestars
  p70.blenddisablelighting = 10 * l70.GetInPlayIntensity ' whitestars
  p69.blenddisablelighting = 10 * l69.GetInPlayIntensity ' whitestars
  p68.blenddisablelighting = 10 * l68.GetInPlayIntensity ' whitestars
  p67.blenddisablelighting = 10 * l67.GetInPlayIntensity ' whitestars
  p66.blenddisablelighting = 10 * l66.GetInPlayIntensity ' whitestars
  p65.blenddisablelighting = 10 * l65.GetInPlayIntensity ' whitestars
  p64.blenddisablelighting = 10 * l64.GetInPlayIntensity ' whitestars
  p63.blenddisablelighting = 10 * l63.GetInPlayIntensity ' whitestars
  p89.blenddisablelighting = 10 * l89.GetInPlayIntensity ' whitestars
  p88.blenddisablelighting = 10 * l88.GetInPlayIntensity ' whitestars
  p87.blenddisablelighting = 10 * l87.GetInPlayIntensity ' whitestars
  p86.blenddisablelighting = 10 * l86.GetInPlayIntensity ' whitestars
  p85.blenddisablelighting = 10 * l85.GetInPlayIntensity ' whitestars
  p84.blenddisablelighting = 10 * l84.GetInPlayIntensity ' whitestars
  p83.blenddisablelighting = 10 * l83.GetInPlayIntensity ' whitestars
  p82.blenddisablelighting = 10 * l82.GetInPlayIntensity ' whitestars
  p81.blenddisablelighting = 10 * l81.GetInPlayIntensity ' whitestars
  p80.blenddisablelighting = 10 * l80.GetInPlayIntensity ' whitestars
  p62A.blenddisablelighting = 10 * l62a.GetInPlayIntensity ' whitestars top
  p62B.blenddisablelighting = 10 * l62b.GetInPlayIntensity ' whitestars top
  p62C.blenddisablelighting = 10 * l62c.GetInPlayIntensity ' whitestars top
  p_wl1.blenddisablelighting = 15 * l_wl1.GetInPlayIntensity ' bluestars top
  p_wl2.blenddisablelighting = 15 * l_wl2.GetInPlayIntensity ' bluestars top
  p_wl3.blenddisablelighting = 15 * l_wl3.GetInPlayIntensity ' bluestars top

  p_WormKicker1enabled.blenddisablelighting = 2 * l_w1b.GetInPlayIntensity ' WormKicker1light
  p_WormKicker2enabled.blenddisablelighting = 2 * l_w2b.GetInPlayIntensity ' WormKicker2light
  p_WormKicker3enabled.blenddisablelighting = 2 * l_w3b.GetInPlayIntensity ' WormKicker3light
  p_WormKicker4enabled.blenddisablelighting = 2 * l_w4b.GetInPlayIntensity ' WormKicker3light
  p_ChaosKickerenabled.blenddisablelighting = 3 * l_ChaosKicker.GetInPlayIntensity ' WormKicker3light

  pLightShootAgain.blenddisablelighting = 2 * LightShootAgain.GetInPlayIntensity

' debug.print abs ( 200 * LightLeftEscape.GetInPlayStateBool ) &" " & LightLeftEscape.GetInPlayStateBool

  p30.blenddisablelighting = l30.GetInPlayIntensity / 10 ' plungerlane
  p31.blenddisablelighting = l31.GetInPlayIntensity / 10 ' plungerlane
  p32.blenddisablelighting = l32.GetInPlayIntensity / 10 ' plungerlane
  p33.blenddisablelighting = l33.GetInPlayIntensity / 10 ' plungerlane
  p34.blenddisablelighting = l34.GetInPlayIntensity / 10 ' plungerlane
  p35.blenddisablelighting = l35.GetInPlayIntensity / 10 ' plungerlane

p_AwardKickerEnabled.blenddisablelighting = l37b.GetInPlayIntensity / 10 ' lightWheelaward Kicker
p_WheelKickerEnabled.blenddisablelighting = Light_WheelAwardb.GetInPlayIntensity / 10 ' lightWheelaward Kicker
p_PlanetDragonEnabled.blenddisablelighting = l_PlanetDragonEnabled.GetInPlayIntensity / 10 ' planets middle
p_PlanetSpiderEnabled.blenddisablelighting = l_PlanetSpiderEnabled.GetInPlayIntensity / 10 ' planets middle
p_PlanetBirdEnabled.blenddisablelighting = l_PlanetBirdEnabled.GetInPlayIntensity / 10 ' planets middle
p_PlanetReversEnabled.blenddisablelighting = l_PlanetReversEnabled.GetInPlayIntensity / 10 ' planets middle
p_PlanetChaosEnabled.blenddisablelighting = l_PlanetChaosEnabled.GetInPlayIntensity / 10 ' planets middle
pLightBallLock1.blenddisablelighting = LightBallLock1.GetInPlayIntensity / 10 ' lock1
pLightBallLock2.blenddisablelighting = LightBallLock2.GetInPlayIntensity / 10 ' lock2
p28a3.blenddisablelighting = l28a3.GetInPlayIntensity / 10
p28a2.blenddisablelighting = l28a2.GetInPlayIntensity / 10
p28a1.blenddisablelighting = l28a1.GetInPlayIntensity / 10



  p36.blenddisablelighting = 55 * l36.GetInPlayIntensity
  p37.blenddisablelighting = 55 * l37.GetInPlayIntensity
  p38.blenddisablelighting = 55 * l38.GetInPlayIntensity
  p39.blenddisablelighting = 55 * l39.GetInPlayIntensity
  p40.blenddisablelighting = 55 * l40.GetInPlayIntensity
  p41.blenddisablelighting = 55 * l41.GetInPlayIntensity
  p42.blenddisablelighting = 55 * l42.GetInPlayIntensity
  p43.blenddisablelighting = 55 * l43.GetInPlayIntensity
  p44.blenddisablelighting = 55 * l44.GetInPlayIntensity
  p45.blenddisablelighting = 55 * l45.GetInPlayIntensity




  pLightLeftInlane.blenddisablelighting = LightLeftInlane.GetInPlayIntensity  / 10
  pLightRightInlane.blenddisablelighting = LightRightInlane.GetInPlayIntensity  / 10



  pLightClockBroken.blenddisablelighting = 7 * LightClockBroken.GetInPlayIntensity / 10
  pLightClockBroken001.blenddisablelighting = LightClockBroken.GetInPlayIntensity / 5
  pLightClockBroken002.blenddisablelighting = LightClockBroken.GetInPlayIntensity / 5
  pLight_BallSaver.blenddisablelighting = Light_BallSaver.GetInPlayIntensity / 10
  pLight_easyhard.blenddisablelighting = Light_easyhard.GetInPlayIntensity * 4
pLight_WheelAward.blenddisablelighting = Light_WheelAward.GetInPlayIntensity * 4
p37a.blenddisablelighting = l37a.GetInPlayIntensity * 4
p28.blenddisablelighting = l28.GetInPlayIntensity * 10

'spinningdragon

L = Lighteyeball1.GetInPlayIntensity *4 + Lighteyeball2.GetInPlayIntensity *4 ' 100
'spindiscimg.opacity = 120+L
'spindiscimg.color = rgb (255 , 153- L *1.5 , 51- L /2)



  if abs(LightShootAgain.GetInPlayStateBool) = 1 Then
    zoomeye = zoomeye + 1 : if zoomeye > 10 Then zoomeye = 10
  Else
    zoomeye = zoomeye -1 : if zoomeye < 1 Then zoomeye = 0
  End If
  pLightShootAgain.size_y = 333 -  (zoomeye * 5)
  pLightShootAgainoff.size_y = 333 - (zoomeye * 5)


' pLightShootAgain.size_X = 200 + zoomeye*2
' pLightShootAgainoff.size_x = 200 + zoomeye*2
' pLightShootAgain.size_Z = 200 + zoomeye*2
' pLightShootAgainoff.size_Z = 200 + zoomeye*2
' pLightShootAgain.roty = pLightShootAgain.roty + 1
' pLightShootAgainoff.roty = pLightShootAgainoff.roty + 1

End Sub


' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows




End Sub

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
' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

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

' *** Trim or extend these to *match* the number of balls/primitives/flashers on the table!
dim objrtx1(9), objrtx2(9)
dim objBallShadow(9)
Dim OnPF(9)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9)
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
    objrtx1(iii).z = 1 + iii/1000 + 0.01      'Separate z for layering without clipping
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = 1 + iii/1000 + 0.02
    objrtx2(iii).visible = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii/1000 + 0.04
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
  Dim falloff: falloff = 150 'Max distance to light sources, can be changed dynamically if you have a reason
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
  For s = lob to UBound(gBOT)

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If gBOT(s).Z > 30 Then              'The flasher follows the ball up ramps while the primitive is on the pf
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
'       objBallShadow(s).Z = gBOT(s).Z + s/1000 + 0.04    'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

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
          If LSd < falloff Then
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

Sub ramptrigger001_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger02_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub


Sub Trigger12_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
  LastSwitchHit= "Trigger12"
End Sub

Sub Trigger13_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
  LastSwitchHit="Trigger13"
End Sub

Sub ramptrigger02_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

'Sub Wall017_hit()
' WireRampOff ' Exiting Wire Ramp Stop Playing Sound
'End Sub

Sub Wall017_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger03
End Sub

'//////////////////////////////////////////////////////////////////////
'// Ball Rolling
'//////////////////////////////////////////////////////////////////////

'ReDim rolling(tnob)
'InitRolling

Dim DropCount
ReDim DropCount(tnob)

'Sub InitRolling
' Dim i
''  For i = 0 to tnob
'   rolling(i) = False
' Next
'End Sub
Dim BallLights
BallLights = array(BallLight001,BallLight002,BallLight003,BallLight004,BallLight005,BallLight006,BallLight007,BallLight008,BallLight009,BallLight010,BallLight011,BallLight012,BallLight013,BallLight014,BallLight015,BallLight016)
Dim BallColors
BallColors = array( RGB(255,255,0),RGB(0,255,255),RGB(0,255,255),RGB(0,255,0),RGB(255,0,0),RGB(0,255,0),RGB(0,255,255),RGB(255,0,255) )
Dim ColorBalls
Dim B_Color
Dim B_image

Sub StartColouredBalls
  ColorBalls = True
  B_Color = rnd(Int(1)*8)
  B_image = "ballgreydragon2"
End Sub

Sub StopColouredBalls
  ColorBalls = False
  B_Color = RGB(255,255,255)
  B_image = "ballgreydragon2"
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
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

        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
'                ballpitch = Pitch(BOT(b))
'                ballvol = Vol(BOT(b))
''            Else
 '               ballpitch = Pitch(BOT(b)) + 35000 'increase the pitch on a ramp
  '              ballvol = Vol(BOT(b)) * 10
           End If
            rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial * 0.7, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
            If rolling(b) = True Then
                StopSound("BallRoll_" & b)
                rolling(b) = False
           End If
        End If
        if b > 1 Then
            If ColorBalls = True Then
                BOT(b).image = B_image
                If Bot(b).ID < 22222 Then BOT(b).id = BOT(b).id + 22222 : BOT(b).color = BallColors(B_Color) : B_Color = B_Color + 1 : If B_Color > 7 Then B_Color = 0
            Else
                BOT(b).color = RGB(255,255,255)
                BOT(b).image = "ball_HDR_brighter"
                If BOT(b).id > 22222 Then BOT(b).id = BOT(b).id - 22222
            End If

            BallLights(b-2).x = bot(b).x
            BallLights(b-2).y = bot(b).y
            BallLights(b+6).x = bot(b).x
            BallLights(b+6).y = bot(b).y
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

FlipperUpAttackMinimumSoundLevel = 0.510                      'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.935                'volume level; range [0, 1]
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

BallWithBallCollisionSoundFactor = 5                  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.5                     'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.1                   'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.2                    'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                 'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                 'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                     'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5                          'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                     'volume multiplier; must not be zero
DTSoundLevel = 0.25                           'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                       'volume level; range [0, 1]
SpinnerSoundLevel = 0.7                                       'volume level; range [0, 1]

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

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

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

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
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

'Function RndNum(min, max)
' RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
'End Function

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

'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************

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


'VRSTUFF (Rawd)
'******************* VR Plunger **********************
Sub TimerPlunger_Timer
  If VR_Primary_plunger.Y < -71 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 3
  End If
End Sub

Sub TimerPlunger2_Timer
  VR_Primary_plunger.Y = -151.488 + (5* Plunger.Position) -20
End Sub
 '*** End VR Plunger **********************************

' VRClock Model cogs...
Sub NewClockTimer_timer()
ClockPiece1.roty = ClockPiece1.roty + 0.06
ClockPiece2.roty = ClockPiece2.roty - 0.16
ClockCylinder1.roty = ClockCylinder1.roty - 0.16
ClockPiece3.roty = ClockPiece3.roty - 0.16
ClockPiece4.roty = ClockPiece4.roty + 0.07
ClockCylinder2.roty = ClockCylinder2.roty + 0.1
End Sub

' ***************** VR Clock code below ******************
Dim CurrentMinute ' for VR clock
Sub ClockTimer_Timer()
NewSecondsHand.RotAndTra2 = (Second(Now()))*6 - 270
NewMinutessHand.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6 -270
NewHoursHand.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2) - 270
CurrentMinute=Minute(Now())
End Sub
 ' ********************** END CLOCK CODE   *********************************

' ***** Beer Bubble Code - Rawd *****
Sub BeerTimer_Timer()

Randomize(21)
BeerBubble1.z = BeerBubble1.z + Rnd(1)*0.5
if BeerBubble1.z > -771 then BeerBubble1.z = -955
BeerBubble2.z = BeerBubble2.z + Rnd(1)*1
if BeerBubble2.z > -768 then BeerBubble2.z = -955
BeerBubble3.z = BeerBubble3.z + Rnd(1)*1
if BeerBubble3.z > -768 then BeerBubble3.z = -955
BeerBubble4.z = BeerBubble4.z + Rnd(1)*0.75
if BeerBubble4.z > -774 then BeerBubble4.z = -955
BeerBubble5.z = BeerBubble5.z + Rnd(1)*1
if BeerBubble5.z > -771 then BeerBubble5.z = -955
BeerBubble6.z = BeerBubble6.z + Rnd(1)*1
if BeerBubble6.z > -774 then BeerBubble6.z = -955
BeerBubble7.z = BeerBubble7.z + Rnd(1)*0.8
if BeerBubble7.z > -768 then BeerBubble7.z = -955
BeerBubble8.z = BeerBubble8.z + Rnd(1)*1
if BeerBubble8.z > -771 then BeerBubble8.z = -955
End Sub

' End VR Stuff........***********************************************************
'********************************************************************************


'****************************************************
'Callout overlap prevention for timer set to 5seconds
'****************************************************
Dim CalloutActive
Sub CalloutTimer_Timer()
  CalloutActive=False
  debug.print "Callout Timer Activated"
  CalloutTimer.Enabled=False
End Sub


'*************************
'PhoenixBird-UP/DOWN Animation
'*************************

Dim BirdPos, BirdDir, BirdShakePos, BirdShakeDir,BirdUp

BirdPos =-135
BirdShakePos = 120

Sub BirdAnimTimer_Timer()

    BirdShakeTimer.Enabled = 0
    BirdPos = BirdPos + BirdDir
    'Bird is moving up
    If BirdPos >= -148Then
        Me.Enabled = 0
        BirdPos = -150
        BirdShakeDir = 1
        BirdShakeTimer.Enabled = 1
    End If
    'Bird is moving down
    If BirdPos <= -152 Then
        Me.Enabled = 0
        BirdPos = -150
  End If
Dim i
For each i in PhoenixBird: i.Transz = BirdPos: Next
End Sub

Sub BirdShakeTimer_Timer
    BirdShakePos = BirdShakePos+ BirdShakeDir
    If BirdShakePos > 125 Then
        BirdShakeDir = -1
    End If
    'Bird is moving down
    If BirdShakePos< 115 Then
        BirdShakeDir = 1
    End If
Dim i
For each i in PhoenixBird :i.Transz = BirdShakePos:  Next
End Sub

Sub BirdMoveUp()
  BirdUp=True
    BirdDir = 2
    BirdAnimTimer.Enabled = 1

End Sub

Sub BirdMoveDown()
  BirdUp=False
    BirdDir = -2
    BirdAnimTimer.Enabled = 1
End Sub

'*************************
'Spider-UP/walk Animation
'*************************

Dim SpiderPos, SpiderDir,SpiderShakePos, SpiderShakeDir, SpiderMoveX,SpiderMoveY,SpiderWalkCount,SpiderUp
SpiderWalkCount=0

SpiderPos =-120
SpiderShakePos = 134
SpiderMoveX=0
SpiderMoveY=0

Sub SpiderAnimTimer_Timer()
    SpiderShakeTimer.Enabled = 0


    SpiderPos= SpiderPos + SpiderDir
debug.print "SpiderPosCcalculation"

    'Spideris moving up
    If SpiderPos >= -119 Then
        Me.Enabled = 0
        SpiderPos =-120
        SpiderShakeDir = 1
        SpiderShakeTimer.Enabled = 1
Dim i
  For each i in Spider1: i.Transz = -SpiderPos:Next
    End If

    'Spider is moving down
    If SpiderPos <= -121 Then
        Me.Enabled = 0
        SpiderPos = -121
  For each i in Spider1: i.Transz = SpiderPos:Next
  End If

End Sub


Sub SpiderShakeTimer_Timer
 'debug.print "SpiderWalkCount"
    SpiderWalkCount=SpiderWalkCount+1
    SpiderShakePos = SpiderShakePos+ SpiderShakeDir

    'Spider is Rotating CW
    If  SpiderShakePos > 136 Then
        SpiderShakeDir = -2
    End If

    'Spider is Rotating CCW
    If  SpiderShakePos <132  Then
        SpiderShakeDir = 2
    End If
    Spider.Rotz = SpiderShakePos
End Sub


Dim SpiderDown
Sub SpiderMoveUp()
  SpiderDown=False
    'Play a Sound?
    SpiderDir = 2
    SpiderAnimTimer.Enabled = 1
  SpiderUp=True
End Sub

Sub SpiderMoveDown()
'If bAttractMode=True Then  For each i in Spider1:i.Transx=SpiderWalkCount * 0.15 :Next
'If bAttractMode=True Then  For each i in Spider1:i.TransY=SpiderWalkCount * 1:Next
  SpiderWalkCount=0
    SpiderAnimTimer.Enabled = 1
  SpiderDown=True
    SpiderDir = -2
  SpiderPos =-120
  SpiderUp=False
Dim i
  For each i in Spider1: i.Transz = -120:Next
End Sub



'*************************
'SDragon-UP/walk Animation
'*************************

Dim DragonPos, DragonDir, DragonShakePos, DragonShakeDir,DragonUp


DragonPos =-80
DragonShakePos = 70

Sub DragonAnimTimer_Timer()

    DragonShakeTimer.Enabled = 0
    DragonPos = DragonPos + DragonDir
    'Dragonis moving up
    If DragonPos >= -81 Then
        Me.Enabled = 0
        DragonPos = -80
        DragonShakeDir = 1
       DragonShakeTimer.Enabled = 1

    End If
    'Dragon is moving down
    If DragonPos <= -81 Then
        Me.Enabled = 0
        DragonPos = -80
  DragonPit1.Visible=0: DragonPit2.Visible=0:DragonPit3.Visible=0:FireChange=0
    End If
Dim i
For each i in Dragon1: i.Transz = DragonPos: Next
End Sub

Sub DragonShakeTimer_Timer
  FirePulse
    DragonShakePos = DragonShakePos+ DragonShakeDir
    If DragonShakePos > 80 Then
        DragonShakeDir = -2
    End If
    'Dragonis moving down
    If DragonShakePos< 60 Then
        DragonShakeDir = 2
    End If
Dim i
For each i in Dragon1 :i.Transz = DragonShakePos: Next
End Sub

Dim FireChange
Sub FirePulse
  FireChange=FireChange+1
  Select Case FireChange
    Case 1: DragonPit1.Visible=1: DragonPit2.Visible=0: DragonPit3.Visible=0
    Case 2: DragonPit1.Visible=0: DragonPit2.Visible=1: DragonPit3.Visible=0
    Case 3: DragonPit1.Visible=0: DragonPit2.Visible=0: DragonPit3.Visible=1
    Case 4: DragonPit1.Visible=0: DragonPit2.Visible=1: DragonPit3.Visible=0:FireChange=0
  End Select
End Sub

Dim DragonDown
Sub DragonMoveUp()
  DragonDown=False
    'Play a Sound?
    DragonDir = 2
    DragonAnimTimer.Enabled = 1
  DragonUp=True

End Sub

Sub DragonMoveDown()
  DragonDown=True
    DragonDir = -2
    DragonAnimTimer.Enabled = 1
  DragonUp=False
End Sub


' #####################################
' ###### Fluppers Domes V2.2 #####
' #####################################



Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.25  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.9   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.9
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20)
Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "white"
InitFlasher 2, "blue"
InitFlasher 3, "white"
InitFlasher 4, "blue" :
'InitFlasher 5, "red" :
'InitFlasher 6, "blue"
'InitFlasher 7, "red" :
'InitFlasher 8, "blue"
'InitFlasher 9, "blue" :
'InitFlasher 10, "red"
'InitFlasher 11, "white"

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
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :  objflasher(nr).imageA = "domeblueflash" :  objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4)
    Case "red" : objflasher(nr).imageA = "domeredflash" :  objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4): objlight(nr).intensity = 3900
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59)
  end select
  objlight(nr).colorfull = objlight(nr).color
' If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
'   objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
'   ObjFlasher(nr).y = ObjFlasher(nr).y + 10
' End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)

  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 50 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 50 * ObjLevel(nr)^2

'*** Ramp flash with color flashers   blue=prim034   red=prim062
' If nr=3 or nr=8 or nr=9 Then primitive034.BlendDisableLighting=10 * ObjLevel(nr)^2
' If nr=1 or nr=2 or nr=7 Then primitive062.BlendDisableLighting=10 * ObjLevel(nr)^2

  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then
    objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0
'   primitive034.BlendDisableLighting=0
'   primitive062.BlendDisableLighting=0
  End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
'Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
''Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub
'Sub FlasherFlash7_Timer() : FlashFlasher(7) : End Sub
'Sub FlasherFlash8_Timer() : FlashFlasher(8) : End Sub
'Sub FlasherFlash9_Timer() : FlashFlasher(9) : End Sub
'Sub FlasherFlash10_Timer() : FlashFlasher(10) : End Sub
'Sub FlasherFlash11_Timer() : FlashFlasher(11) : End Sub


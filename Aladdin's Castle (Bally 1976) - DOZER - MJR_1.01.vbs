' Aladdin's Castle
' Bally, 1976
' Designed by Greg Kmiec
' Electromechanical, 2-player
'
' About this Visual Pinball re-creation
'   Version MJR-1
'   Optimized for cabinets:  Full Screen, B2S (with three screens), and DOF
'   Requires VP9, version 9.9.1 or higher
'
' Special Feature: Open the coin door (press the End key) to display the operator
' menu.  This lets you set up the feature adjustments that were available in the
' original machine via various switches and plugs.  I think all of the adjustments
' that were possible on the original are faithfully implemented.
'
' Permissions: Mods and any other re-use are hereby allowed without separate
' permission.  Anyone may make further modifications to this version or use it
' the basis for other tables.
'
' This version is MJR's revamp of August, 2015, based on Itchigo's version of August,
' 2014, which was in turn based on an earlier desktop version by Greywolf.  This new
' version was created with Itchigo's permission.
'
' Main updates in this version:
'
'   Added option settings for all of the operator adjustments described in the
'       original Bally manual (balls per game, replay score levels, replay award
'       as credit or extra ball, special award as credit or extra ball, and
'     Aladdin's Alley "liberal" or "conservative" scoring).  Also added a
'   Free Play option, which allows starting a game (or adding a player)
'   with 0 credits.
'   Added an "operator menu" to adjust the new option settings interactively.
'       Open the coin door (press the End key) to display the menu; navigate it
'       using the flipper buttons and Start button.
'   Added high score initials, entered using the standard flipper button UI seen
'       in alphanumeric and DMD machines.
'   Changed Aladdin's Alley scoring to match the pattern described in the original
'       Bally operator's manual (including the Liberal and Conservative variations).
'   Fixed Aladdin's Alley scoring when lit for Special.  The old script awarded 5000
'       points on Special in addition to a credit; the real game doesn't award any
'       points, just the credit or extra ball.
'   Reversed the top lane light states to match the authentic behavior.  (On the
'       original machine, each lane's light is initially ON, and turns OFF when the
'       lane is completed.  The corresponding upper and rollover buttons turn on
'       when a lane is completed; this part was correct in the original script.)
'   Added DOF event support: custom DOF events fire for all of the key game events,
'       to allow triggering flashers, solenoids, knockers, and other feedback devices
'       in response to game events.
'   Suppressed digitized sound effects that are redundant with the usual complement
'       of DOF effects (for bumpers, flippers, knocker, etc) when DOF is present.
'       DOF effects for scoring chimes can be enabled and disabled separately.
'   Added controller selection via Visual Pinball\User\cController.txt.  Since
'       Aladdin's Castle isn't a ROM-based table, it doesn't use VPinMAME, so the
'       controller setting only affects whether DOF is used for audio effects.
'   Cleaned up graphics for the plastics (higher resolution, added lamp effects)
'   Improved the bumper cap visuals using new primitives
'   Improved the graphics for the gates and rubber stopper along the top arch
'   Added a VP9.9.1 custom plunger with mech plunger input enabled (for cabinet use)
'   Fixed some broken sound effects (ball rolling, rubber bumps, etc)
'   Added more ball-rolling sounds (the greater variety creates a more natural effect)
' Fixed some problems in saving and restoring settings
'   Added support for cabinet tilt bob switches using the "T" key (this just pulses
'       the virtual tilt switch in the game, without adding any nudging physics -
'       we assume that there's an accelerometer providing the actual nudge input)
'   Substantially reworked the VB Script code for clarity and maintainability.
'   Cleaned up "bell" audio samples (noise/hiss reduction)
'   Marked appropriate audio effects as "backglass output" for dual audio systems
'   Changed to more authentic "Over The Top" handling, with the backglass light
'       displayed only temporarily after rolling over the score reels.  (An option
'       allows enabling the old behavior.)
'   Reworked the scoring code for several features to use a simulated score motor
'       that models the behavior of the real machine's mechanical scoring system.
'       This is used for Aladdin's Alley and all of the targets and rollovers with
'       300-point values.  This makes the Aladdin's Alley scoring behavior match
'       the real machine's behavior more closely in some unusual cases.
'
' The companion B2S backglass also has several improvements:
'
'   Cleaned up the main graphics
'   Improved the score reel and credit reel geometry
'   Changed to authentic style and placement of all of the backglass indicator lights
'       (ball in play, match numbers, "Same Player Shoots Again", "Tilt", etc)
'   Added decorative flashing backlight lamps that approximate the original machine's
'       backglass lighting (in my version they're random, whereas the original probably
'       had a fixed sequence, but the fixed sequence looked random enough that the
'       random sequencing here seems to look about right)
'   Added a Bally logo backdrop on the third monitor, in lieu of a DMD
'   Added a high score display to the third monitor, showing the highest score to date
'       and the player's initials, using a simulated sticky note
'
'
' B2S lamp IDs
'
' For the backglass, we use the standard B2S IDs for the standard complement of
' indicator lamps (ball in play, match numbers, shoot again, tilt, etc).  We also
' have a few custom lamps not in the standard set:
'
'   9           Single shared "Over The Top" light for all players, used when the
'               authentic OTT style is selected
'   10-11       Separate "Over The Top" lights for Players 1 and 2, used when the
'               non-authentic modernized OTT style is selected
'
'   200-211     decorative illumination backlighting (these are scattered around
'               the backglass to simulate the authentic accent lights; we flash
'               these at random on a timer)
'
'   228         high score number currently displayed (High Score 1..4)
'   229         first digit of month of current high score
'   230         second digit of month
'   231         first digit of day of month of current high score
'   232         second digit of day
'   233         first digit of four-digit year of current high score
'   234         second digit of year
'   235         third digit of year
'   236         fourth digit of year
'   237         first letter of high score player's initials
'   238         second leter " "
'   239         third letter " "
'   240         digit 0 (units) of high score indicator
'   241         digit 1 (tens) of high score indicator
'   242         digit 2 (hundreds) " "
'   243         digit 3 (thousands) " "
'   244         digit 4 (ten thousands) " "
'   245         digit 5 (hundred thousands) " "
'
' The high score indicator uses ten lamps per ID to make up the score, because
' there doesn't seem to be any good way in B2S to set a text display directly.
' Each ID has ten associated lamps, with B2SValue=1 for "1", 2 for "2", ...,
' and 10 for "0".  Value 0 turns the digit off entirely.  The player's initials
' work similarly, with 1="A", 2="B", 3="C", ..., 26="Z".
'
' DOF event assignments:
'
'   101-104     Top A-D rollovers
'   105-108     Top A-D lane switches
'   109-112     Lower A-D rollovers
'   113         Aladdin's Alley rollovers
'   114-115     Left Outlane/Inlane
'   116         Right Outlane
'   117         Left stand-up target
'   118         Right stand-up target
'   119         Top bumper
'   120         Left bumper
'   121         Bottom bumper
'   122         Left slingshot
'   123         Drain
'   124         Ball in shooter lane
'   125         Spinner - 10pts
'   126         Spinner - 100pts
'   127         Knocker
'   128         Left flipper
'   129         Right flipper
'   130         Drain kicker (serve new ball to plunger)
'   131         10 point bell
'   132         100 point bell
'   133         1000 point bell


'Itchigo's 4 Player EM Template for Visual Pinball With Direct B2s
'-------------------------------------------------------------------------------------------------------------------------
'                                                  DO NOT REMOVE!!!!!!!!!!!!!!!
'This template is for anyone to use to make a table out of.
'Do not make another template out of it, use it to build something!! Yes, I had to say that BECAUSE IT'S BEEN DONE BEFORE!
'If you use it please give me credit that you used it. Any parts or code can be borrowed for other tables.
'Although the code is in for extraball and special- you have to code the event to the light to turn on, I've taken care
'of the rest. The basic DB2s code has been added for basic commands. This will not support animations or complex things,
'but you can add them. Since they will vary from table to table it's best to leave that to the individual author. However,
'ball in play, match, shoootagain,gameover,tilt, playerup, and insert coin are already coded in as they are basic things on most tables.

'Building in Desktop Mode: No modifications neccessary, build as normal.

'Building in Full Screen Mode: Simply uncomment the DB2s code and make sure your DB2s lights id's match the id's in the script.
'The id's in the script are straight from the editor, so as long as you id them correctly they will work right out of the box.
'FullScreen: All Objects on the backglass can be dragged off the screen so they are not seen, rather than deleting the code for each item.
'Values: For ball or Playerup, you need a DB2s value for that light. For a DB2s light that shows ball 1, the value is 1.
'The same thing is true for balls 2,3,4,5, etc, the ball number is the value in the B2s editor when the backglass is made.
'Use the same idea for Playerup. Example: Controller.B2SSetBallInPlay Ball (The parameter is ball, the value will tell DB2s WHAT ball).
'If you just want a light that says "ball" then use Controller.B2SSetBallInPlay 32,1 (32 is the light Id, 1, is the lightstate).
'Extraball: Extraball is precoded, you just have to turn the "Shootagain" light on, it's already precoded to give you an Extraball.
'Just add this when you turn the light on: Controller.B2SSetShootAgain 36,1 This will turn on "Shoot Again" on the backglass.

'I have optimised the coding as my coding has gotten better. Everything should still work just the same.
'Any bugs to report or anything you'd like to see added, you can find me at Roguepinball.com

'************************* Credits ***************************************************************************************
'Author of the original version this is based on: Itchigo.
'Slings, rubbers, posts, plastics, lane triggers and ball lane guides by Faralos.

' Thalamus 2018-07-24
' No special SSF tweaks yet.
' Added mechanicaltilt

Option Explicit
Randomize                          'This generates a random number for the match.

Dim Ballsize,BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000

'**** GLOBAL GFX / FX OPTIONS *****
Const GI_Dim = 1
'**** DESKTOP MODE GFX/FX Options ****
Const Show_Glass = 1
Const Show_Hands = 1
Const Show_Cig = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "Aladdins_Castle"   ' this is the ID that B2S uses to save settings and DOF uses to load its configuration -
                                     ' Aladdins_Castle is the name for this table used in the online DOF config tool database


' Do you have an accelerometer?  This version is optimized for cabinets, so we'll
' assume that the answer is yes.  Set this to false if you don't.
'
' The only effect of this option is on how the script handles the "T" key.  If an
' accelerometer is present, the "T" key will simply pulse the virtual tilt switch.
' If no accelerometer is present, the "T" key will ALSO apply a fake physics nudge,
' in addition to pulsing the virtual tilt switch.  In cabinets with accelerometers,
' you want all of the nudging to come from the accelerometer, not from keyboard
' inputs.  The "T" key in an accelerometer-equipped cabinet is usually mapped to
' the physical plumb bob switch.  When that triggers, it means that the cabinet has
' been given a hard shove, which will already have registered on the accelerometer
' and will already have been fed into VP's physics model that way.  So you just
' want the "T" key to register a Tilt switch pulse for the purposes of letting the
' game rules decide if a Tilt penalty is in order.  You *don't* want the physics model
' to get another fake shove.  Setting this to true will take care of that.
Const CabHasAccelerometer = True


'************************** Game Setup Options *****************************************************************************
'
' Authentic setup options:  These options simulate the adjustable features of the
' original EM table.  In the real machine, these were set via switches and plugs.
' In the virtual version, these can be changed via the Operator Menu, which you can
' display by opening the coin door or pressing the End key.
'
'  - Aladdin's Alley scoring mode:  The real machine has two possible scoring
'    patterns for this feature, set by a jumper:
'
'    Liberal: 500-1000-2000-3000-4000-5000-SPECIAL, then stays on SPECIAL for the
'             remainder of the ball in play
'    Conservative: 500-1000-2000-3000-4000-5000-SPECIAL-5000-SPECIAL-5000, then
'             stays on 5000 for the remaindier of the ball in play.
'
'  - Special Value: this can be set to score a credit or an extra ball.  Specials are
'    awarded by the Aladdin's Alley feature.
'
'  - Replay Value: this can be set to score a credit or an extra ball.  A Reply is
'    awarded when the score reaches one of the high score values.
'
'  - Replay Score 1: this sets the first (lower) replay score.  Set to 0 for none.
'
'  - Replay Score 2: this sets the second (higher) replay score.  Set to 0 for none.
'
'  - Balls per game: set to 3 or 5.
'
'  - Free Play: on or off.  When Free Play is on, the Start Game button will work
'    even if there are 0 credits on the credit reel.  A credit is still deducted
'    on Start Game if any credits are on the reel.  (This option is only
'    semi-authentic.  There's no mention of an official free play option in
'    the original operator's manual, but you could still set up a real Aladdin's
'    Castle on free play with a little know-how.  Most Williams and Bally EM
'    games from this generation can be easily converted to free play with a small
'    tweak to the coin taker wiring.)
'

Dim OptAlleyScoring      ' 0=liberal, 1=conservative (default)
Dim OptSpecialVal        ' 0=credit (default), 1=extra ball
Dim OptReplayVal         ' 0=credit (default), 1=extra ball
Dim OptReplay1           ' replay score 1 (default=65000)
Dim OptReplay2           ' replay score 2 (default=99000)
Dim OptBallsPerGame      ' 3 (default) or 5
Dim OptFreePlay          ' 0=one coin per credit (default), 1=free play

Dim moptions 'Declortaion for desktop options menu.

' High Score Initials Entry UI style - modern or simpler.
'
' If SimplerInitUI is set to true, we use the "simpler" UI.  In this mode, the
' entry process ends immediately when the player enters the third letter.  Some
' older pinballs and video games used this sort of UI.
'
' If SimplerInitUI is set to false, we use a slightly more complex UI that's more
' like what most real alphanumeric and DMD pinballs use.  In this mode, after the
' player enters the third initial, the UI displays an END symbol.  Selecting the
' END symbol commits the entry.  But if the player wishes  to make changes, s/he
' can use the flipper buttons to switch to the Backspace symbol instead of the
' END symbol.  Selecting the Backspace symbol backs up to the third initial and
' lets the user change that entry or back up further.  So with the modern UI, the
' user gets one more chance to makes corrections after entering the third letter.
Const SimplerInitUI = false


' Aladdin's Alley points awarded on Special.  On the real machine, no points are
' awarded on Special, so this is set to 0 by default.  The old version of this script
' awarded 5000 points on Special, so I've kept that here as an option.  It's not in
' the operator menu since that was never an adjustable option on the real machine.
Const AlleyPointsOnSpecial = 0


' "Over The Top" style - authentic or modernized.

' If this is set to True, we use the authentic behavior.  The original feature was
' the same as in many other Bally machines of the time.  There was a single "Over
' The Top" light at the top of the backglass, shared for both players.  This light
' came on for a few seconds along with a buzzer right after either player's score
' flipped past 99,990.  The light turned off after a few moments; it didn't stay
' on through the end of the game or even through the current ball in play.
'
' If this is set to False, we use a modernized version that's not authentic but is
' nicer UI-wise.  In this version, we use a separate "Over The Top" light for each
' player, located above the corresponding score reel.  Once triggered, the light
' stays on until the start of the next game.  The light in this mode is effectively
' a limited sixth reel that roughly doubles the maximum displayable score from
' 99,990 to 199,990.
Const AuthenticOverTheTop = true


'*************************************************


Dim ScorePlayer(4)                  'Current score for player (N)  (Player 1 -> N=1, Player 2 -> N=2, etc)
Dim Replay1Awarded(4), Replay2Awarded(4)   'Have we awarded Replay 1 and Replay 2 to player (N)?
Dim OverTheTop(4)                  'Have we signaled "Over The Top" for the player (N)?
Dim InProgress: InProgress=False   'InProgress means a game is running.
Dim Highscore(3),HighscoreDate(3),HighscoreInits(3)  ' Top four high scores to date - score value, date (mm/dd/yyyy), player's initials
Dim Credits,Shootball,Bonusamt,Ball,Player,Players,Value,Score,Playernumber,B1,B2,B3
Dim Tilt,Tiltcount,Matchnumber,I,MMM
Dim BonusMotor                     'Bonus motor counter - the bonus scoring timer uses this to simulate the mechanical scoring motor
Dim AlleyLevel                     'Number of times through Aladdin's Alley so far this ball - for figuring the alley score
Dim SReels(2)
Dim Dtoptions_A

' Dummy controller class - for cases where the user doesn't want B2S to be loaded.
' This class provides dummy methods for the interface members that we use in the
' controller.  This lets us just substitute an instance of this as the Controller
' object when B2S is turned off, so that we don't have to make every Controller
' method call conditional throughout the script.  Calls to the Controller methods
' will simply go through these dummy methods, which will be silently ignored.
class DummyController
  Public B2SName
  Public Sub Run : End Sub
  Public Sub B2SSetData(id, val) : End Sub
  Public Sub B2SSetGameOver(id, val) : End Sub
  Public Sub B2SSetTilt(id, val) : End Sub
  Public Sub B2SSetMatch(id, val) : End Sub
  Public Sub B2SSetScore(player, val) : End Sub
  Public Sub B2SSetCredits(val) : End Sub
  Public Sub B2SSetCanPlay(id, val) : End Sub
  Public Sub B2SSetBallInPlay(id, val) : End Sub
  Public Sub B2SSetPlayerUp(id, val) : End Sub
  Public Sub B2SSetShootAgain(id, val) : End Sub
end class


Dim object

    If ShowDT = True Then
    For each object in DT_Stuff
    Object.visible = 1
    Next
        CabinetRailLeft.visible = 1:CabinetRailRight.visible = 1
        LDB.visible = 1
  End If

  If ShowDt = False Then
    For each object in DT_Stuff
    Object.visible = 0
    Next
        CabinetRailLeft.visible = 0:CabinetRailRight.visible = 0
        LDB.visible = 0
  End If

If Show_Glass = 1 AND ShowDT = True Then
Glass.visible = 1
Else
Glass.visible = 0
End If

If ShowDT = False Then
HLR.visible = 0
HRR.visible = 0
End If

If Show_Hands = 1 AND ShowDT = True Then
HRR.visible = 1
HLR.visible = 1
LFB.visible = 1
RFB.visible = 1
Else
HLR.visible = 0
HRR.visible = 0
LFB.visible = 0
RFB.visible = 0
End If

If Show_Cig = 1 AND ShowDT = True Then
Cigarette.visible = 1
Cig_Smoke.visible = 1
Cig_Smoke1.visible = 1
Light65.visible = 1
Light66.visible = 1
Smoke.enabled = 1
Else
Cigarette.visible = 0
Cig_Smoke.visible = 0
Cig_Smoke1.visible = 0
Light65.visible = 0
Light66.visible = 0
Smoke.enabled = 0
End If

'************************** Initialize Table *******************************************************************************

 Sub Table1_Init()                  'Sub - All executable code must reside in a Sub (Function. Method, or Property) block.
     LoadEM
     Ball=1
     ' TextBox.Text= "Insert Coin"    'When table is first rendered show game over/insert coin.
   For Each Obj In Bumperparts:Obj.Isdropped=True:Next
   For Each Obj In Sidelights: Obj.State=0:Next
   For Each Obj In Toprolloverlights: Obj.State=0:Next
   For Each Obj In Bottomrolloverlights: Obj.State=0:Next
   For Each Obj In Toplights:Obj.State=0:Next
   Players=0                      'Players=0 because a game hasn't started yet.
   for i = 1 to ubound(ScorePlayer) : ScorePlayer(i) = 0 : next   ' All scores to zero
     AlleyLevel = 0                 'Reset the Aladdin's Alley counter
     Matchnumber=10
     If B2SOn Then
     Controller.B2ssetgameover 35,1
     Controller.B2ssetTilt 33,0
     Controller.B2ssetmatch 34, Matchnumber
     End If
     LoadInfo                       'Load the info for when the game boots up.
   BGInitTimer.Enabled = true
     'Declare the desktop reels for desktop view.
     Set SReels(1) = ScoreReel1
     Set SReels(2) = ScoreReel2
     For i = 1 to 2
   'SReels(i).setvalue(scoreplayer(i)) < Uncomment to populate the score reels with high score 1 and 2 in desktop mode.
   Next
   Init_Desktop_Reels
     DOF 165,1
 End Sub

'Populate the various reels and text boxes on the desktop view with variables loaded from the table's save file.
 Sub Init_Desktop_Reels
    op1mval1.text = OptBallsPerGame
    hs1.text = HighscoreInits(0)
  hs1_val.text = Highscore(0)
  HS2.text = HighscoreInits(1)
  hs2_val.text = Highscore(1)
    hs3.text = HighscoreInits(2)
  hs3_val.text = Highscore(2)
  HS4.text = HighscoreInits(3)
  hs4_val.text = Highscore(3)
    op1mval6.text = optreplay1
    op1mval7.text = optreplay2
    Gotilt.setvalue(0)

  For each object in DToptions
     Object.visible = 0
     Next
     If OptAlleyScoring = 0 Then
     op1mval3.text = "Lib"
     Else
     op1mval3.text = "Con"
     End If
     If OptFreePlay = 1 Then
     op1mval2.text = "Yes"
     Else
     op1mval2.text = "No"
     End If
     If OptReplayVal = 1 Then
     op1mval4.text = "EB"
     Else
     op1mval4.text = "CR"
     End If
     If OptSpecialVal = 1 Then
     op1mval5.text = "EB"
     Else
     op1mval5.text = "CR"
     End If
End Sub

' Backglass initializer.  We do this on a timer, since B2S needs a few
' seconds to initialize itself before it will accept any commands from
' the VP side.  This runs once on a timer, then disables itself.
Sub BGInitTimer_Timer
  ' Set the backglass display for the player scores, credits, and high score to date.
  ' Note: I removed the score restoration, because it invokes the motor sound effects
  ' on the backglass.  That seems less authentic than just leaving it zeroed.  The
  ' real machine *would* in fact power up with whatever score was showing when it
  ' was last shut down, since the mechanical reels would just stay put across the
  ' power cycle, but therein lies the problem: the reels on the real machine would
  ' just *stay put* on power up.  So my way is inauthentic we always power up at
  ' zero, but the active animation and sound of the reels is more *noticeably*
  ' unrealistic in that it calls attention to itself.  Leaving the reels at zero
  ' isn't perfect either but is less noticeably imperfect.
  'Controller.B2ssetscore 1, Scoreplayer(0) mod 100000
  'Controller.B2ssetscore 2, Scoreplayer(1) mod 100000
    If B2SOn Then
  Controller.B2ssetCredits Credits
    End If
  CurHighScoreDisp = 0
  HighScoreDisplayTimer.Enabled = True

  ' we only need to run this once, so disable the timer
  BGInitTimer.Enabled = false
End Sub

'*********** Save/restore settings *************************************************************************************

' Variables to save.  For each variable to save/restore, we have three consecutive
' values in the array:
'
'   "Variable Name", "Conversion", Default Value
'
' "Variable Name" is a string giving the name of the script global variable.  We'll
' save from and restore to the value of this variable.  "Conversion" is the conversion
' function to apply to the string value loaded from the persistent store; if the
' target variable is a string, this can simply be empty, and for a number, it can
' be a function like "CDbl" or "CInt", as appropriate.  "Default Value" is the default
' to use on load when the variable isn't defined in the persistent store.

Dim SavedVars : SavedVars = Array( _
  "ScorePlayer(1)", "CDbl", 0, _
  "ScorePlayer(2)", "CDbl", 0, _
  "Credits", "CDbl", 0, _
  "Highscore(0)", "CDbl", 70000, _
  "Highscore(1)", "CDbl", 0, _
  "Highscore(2)", "CDbl", 0, _
  "Highscore(3)", "CDbl", 0, _
  "HighscoreDate(0)", "", "6/16/1976", _
  "HighscoreDate(1)", "", "", _
  "HighscoreDate(2)", "", "", _
  "HighscoreDate(3)", "", "", _
  "HighscoreInits(0)", "", "GK[", _
  "HighscoreInits(1)", "", "", _
  "HighscoreInits(2)", "", "", _
  "HighscoreInits(3)", "", "", _
  "OptAlleyScoring", "CInt", 1, _
  "OptSpecialVal", "CInt", 0, _
  "OptReplayVal", "CInt", 0, _
  "OptReplay1", "CDbl", 65000, _
  "OptReplay2", "CDbl", 99000, _
  "OptBallsPerGame", "CInt", 3, _
  "OptFreePlay", "CInt", 0 _
)


Sub SaveInfo
  dim i
  for i = 0 to ubound(SavedVars)-1 step 3
    SaveValue cGameName, SavedVars(i), Eval(SavedVars(i))
  next
End Sub

Sub LoadInfo
  dim i
  for i = 0 to ubound(SavedVars)-1 step 3
    ' load the saved value from the persistent store
    Value = LoadValue(cGameName, SavedVars(i))

    ' if it's empty, it wasn't defined, so apply the default
    if Value = "" then Value = SavedVars(i+2)

    ' if there's a conversion function, apply it
    if SavedVars(i+1) <> "" then Value = Eval(SavedVars(i+1) & "(Value)")

    ' store the result in the target variable
    Execute SavedVars(i) & " = Value"
  next
End Sub

' Get the default value for a saved variable
Function GetSavedVarDefault(name)
  dim i
  for i = 0 to ubound(SavedVars)-1 step 3
    if SavedVars(i) = name then
      GetSavedVarDefault = SavedVars(i+2)
      exit function
    end if
  next
End Function

' Apply the factory reset high score values
Sub HighScoreReset
  ' apply each HighscoreXXX default from the saved variable list
  for i = 0 to ubound(SavedVars)-1 step 3
    dim name : name = SavedVars(i)
    dim defval : defval = SavedVars(i+2)
    if left(name, 9) = "Highscore" then
      Execute name & " = defval"
    end if
  next
  CurHighScoreDisp = 0
End Sub


' **************************************************************************************************************************
'
' Backglass light timer
' Randomly turns backglass decorative lamps on and off.  These lights don't
' signify anything - they're simply backlighting scattered around the artwork
' area for effect.  The original machine probably has a cam that flashes the
' lights in a particular pattern, but we can get a qualitatively similar effect
' by simply blinking them at random.
'
' The backglass decoration lights are numbered 200+.  There are 11 of these
' lamps (200-210), but in videos of the real machine, it appears that only the
' lights behind the "Aladdins' Castle" title graphics flash - the rest appear
' to be solid on.  The title graphics lamps in our B2S backglass are the first
' 6 (200-205), so only include these in the flashing.
Dim bgFlashers(6), lastBgFlasher
Sub BgLights_Timer()
    do while true
    dim n : n = ubound(bgFlashers)
    dim lno : lno = int(rnd*n)
    if lno < n and lno <> lastBgFlasher then
      lastBgFlasher = lno
      if bgFlashers(lno) then bgFlashers(lno) = 0 else bgFlashers(lno) = 1
      dof 200+lno, bgFlashers(lno)
      exit do
    end if
  loop
End Sub

'******* Key Down *********************************************************************************************************
Dim gxx

Sub Table1_KeyDown(ByVal keycode)   'This is what happens when you push a key down.
  If  Keycode = PlungerKey Then       'JP's plunger stuff.
    Plunger.PullBack                'Pull back the plunger.
        If Show_Hands = 1 AND ShowDT = True Then
        pld = 1:Plunger_Hand.enabled = 1
        End If
  End If
  If  Keycode = LeftFlipperKey And InProgress = True And Tilt=False Then
        If Show_Hands = 1 AND ShowDT=True Then
        hld = 1:hand_left.enabled = 1
        End If

    LeftFlipper.RotateToEnd         'If the above conditions are present then flipper goes up.
        PlaySoundat SoundFXDOF("flipperup",128,DOFOn,DOFFlippers), leftflipper
    PlaySound "BuzzL", -1
        If Gi_Dim = 1 Then
        For Each gxx in GI_Lights
        gxx.IntensityScale = 0.9
        Next
        DOF 166,1:DOF 165,0
        gi_bright.enabled = 1
        End If
  End If
  If  Keycode = RightFlipperKey And InProgress = True And Tilt=False Then
        If Show_Hands = 1 AND ShowDT = True Then
    rld = 1:hand_right.enabled = 1
        End If
    RightFlipper.RotateToEnd        'If the above conditions are present then flipper goes up.
    RightFlipper1.RotateToEnd       'If the above conditions are present then flipper goes up.
    PlaySoundat SoundFXDOF("flipperup",129,DOFOn,DOFFlippers), rightflipper
    PlaySound "BuzzR", -1
        If Gi_Dim = 1 Then
        For Each gxx in GI_Lights
        gxx.IntensityScale = 0.9
        Next
        DOF 166,1:DOF 165,0
        gi_bright.enabled = 1
        End If
  End If
  If  Keycode = AddCreditKey Then
    Credits=credits + 1             'Add a credit.
    MaxCredits                      'Call the max credits sub to check for maximum credits.
    PlaySound "credit"              'Play the sound.
    PlaySound SoundFXDOF("knocker-echo",127,DOFPulse,DOFKnocker)        'Play the sound.
        If B2SOn Then
    Controller.B2ssetCredits Credits
        End If
    ' If  InProgress= False Then TextBox.Text = "Press Start"   'If the game is over then show Press Start.
    PlaySound "coinin"               'I wish I had a quarter for every time I've heard this sound. Amen, Bob!
  End If
  If  Keycode = StartGameKey Then
    StartGame                       'Call the start game sub.
  End If
  If  Keycode = LeftTiltKey Then      'Left shake.
    Nudge 90, 2                     'Degree of shake and strength.
    BumpIt                          'Check for tilt
  End If
  If  Keycode = RightTiltKey Then     'Right shake.
    Nudge 270, 2                    'Degree of shake and strength.
    BumpIt                          'Check for tilt
  End If
  If  Keycode = CenterTiltKey Then    'Center shake.
    Nudge 0, 2                      'Degree of shake and strength.
    BumpIt                          'Check for tilt
  End If
  If  Keycode = MechanicalTilt Then 'Analog tiltbob.
'   Nudge 0, 2                      'Degree of shake and strength.
    BumpIt                          'Check for tilt
  End If
  if Keycode = 207 and Not InProgress and Not OpMenuActive AND showdt = false then   ' End = Coin Door key = operator menu
    ShowOperatorMenu
  end if
  if Keycode = 20 then                'Key 20 = keyBangBack = "T"
    if CabHasAccelerometer then
      ' There's an accelerometer, so DON'T apply any physics nudge.  The "T"
      ' key is usually mapped in cabinets with accelerometers to the physical
      ' tilt bob, so we only want this to pulse the virtual tilt switch.
      '
      ' Older EM games like this typically tilted immediately when the tilt
      ' bob switch made contact.  There wasn't usually any sort of warning or
      ' grace period.  So signal a full tilt here.
      TiltIt:Tilt=True
           ' TiltCount = 3:TiltTimer.enabled = 1
    else
      ' No accelerometer is present, so DO apply a simulated physics nudge.
      ' In non-cabinet setups, the "T" key is meant to represent a hard nudge
      ' nudge straight ahead, so use a large nudge factor (6 is the convention
      ' as the strength factor for this type of nudge).  Since this is such
      ' a hard nudge, treat it as a full tilt.
      Nudge 0, 6
      TiltIt:Tilt=True
             'TiltCount = 3:TiltTimer.enabled = 1
    end if
  end if

  ' do high score initial entry if applicable
  if EnteringHS then HighScoreInitKey keycode

  ' process an operator menu keystroke if the menu is up
  if OpMenuActive then OpMenuKey keycode

  if keycode = 30 then BeginHighScoreInits 2, 0

    'This is the code for the game's options menu in desktop mode.  Rather than convert MJR's cabinet view to desktop, I have
    'just used a series of text boxes that allowing the same variables to be changed.  This menu is opened in desktop mode by
    'pressing the "2" key which I think is the vpinmamem extra ball default key.  Navigation is accomplished with the magnasave keys.

    If keycode = 3 AND ShowDT = True AND Dtoptions_A = 0 Then
    moptions = 1:op1.state = 1:op10.State = 0
    DTOptions_Show
    dtoptions_A = 1
    End If

  If keycode=leftmagnasave AND Dtoptions_A = 1 then
    moptions=moptions+1
    If moptions=11 then moptions=1
    playsound "target"
    Select Case (moptions)
      Case 1:
        Op1.state = 1
        Op10.state = 0
      Case 2:
        Op2.state = 1
        Op1.state = 0
      Case 3:
        Op3.state = 1
        Op2.state = 0
            Case 4:
                Op4.State = 1
                Op3.State = 0
      Case 5:
                Op5.State = 1
                Op4.State = 0
      Case 6:
                Op6.State = 1
                Op5.State = 0
      Case 7:
                Op7.State = 1
                Op6.State = 0
      Case 8:
                Op8.State = 1
                Op7.State = 0
      Case 9:
                Op9.State = 1
                Op8.State = 0
           Case 10:
                Op10.State = 1
                Op9.State = 0
    End Select
  end if

    If keycode=Rightmagnasave AND Dtoptions_A = 1 then
    PlaySound "metalhit"
    Select Case (moptions)
    Case 1:
      if OptBallsPerGame = 3 then
        OptBallsPerGame=5
        op1mval1.text = 5
      else
      if OptBallsPerGame = 5 Then
            OptBallsPerGame=3
      op1mval1.text = 3
      end if
            end if
    Case 2:
      If OptFreePlay = 0 Then
               OptFreePlay = 1
               op1mval2.text = "Yes"
               Else
               OptFreePlay = 0
               op1mval2.text = "No"
            End If
    Case 3:
            If OptAlleyScoring = 0 Then
               OptAlleyScoring = 1
               op1mval3.text = "Con"
               Else
               OptAlleyScoring = 0
               op1mval3.text = "Lib"
            End If
      Case 4:
            If OptReplayVal = 0 Then
               OptReplayVal = 1
               op1mval4.text = "EB"
               Else
               OptReplayVal = 0
               op1mval4.text = "CR"
            End If
      Case 5:
            If OptSpecialVal = 0 Then
               OptSpecialVal = 1
               op1mval5.text = "EB"
               Else
               OptSpecialVal = 0
               op1mval5.text = "CR"
            End If
        Case 6:
             If OptReplay1 => 99000 Then
             OptReplay1 = 49000
             End If
             OptReplay1 = OptReplay1 + 1000
             op1mval6.text = OptReplay1
    Case 7:
             If OptReplay2 => 99000 Then
             OptReplay2 = 49000
             End If
             OptReplay2 = OptReplay2 + 1000
             op1mval7.text = OptReplay2
    Case 8:
             If OptHSReset = 0 Then
               OptHSReset = 1
               op1mval8.text = "Yes"
               Else
               OptHSReset = 0
               op1mval8.text = "No"
            End If
    Case 9:
            Dtoptions_A = 0
            If OptHSReset then
            HighScoreReset
            OptHSReset  = 0
            op1mval8.text = "No"
            End If
            hs1.text = HighscoreInits(0)
      hs1_val.text = Highscore(0)
      HS2.text = HighscoreInits(1)
      hs2_val.text = Highscore(1)
      hs3.text = HighscoreInits(2)
      hs3_val.text = Highscore(2)
      HS4.text = HighscoreInits(3)
      hs4_val.text = Highscore(3)
            moptions = 1
            op9.State = 0
            op1.State = 1
            DTOptions_Hide
            SaveInfo
        Case 10:
             Dtoptions_A = 0
             DTOptions_Hide
    End Select
  End If

End Sub

'*********** Key Up *******************************************************************************************************

Sub Table1_KeyUp(ByVal keycode)     'This is what happens when you release a key.
  If  keycode = PlungerKey Then       'JP's Plunger stuff.
    Plunger.Fire                    'Fire the plunger.
    Playsoundat "Plungerrelease", Screw61
        If Show_Hands = 1 AND ShowDT = True Then
        pld = 11:Plunger_Hand.enabled = 1
        End If
  End If
  If  Keycode = LeftFlipperKey then
    Stopsound "BuzzL"
    if InProgress = True And Tilt=False Then
            If Show_Hands = 1 AND ShowDT = True Then
            hld = 6:hand_left.enabled = 1
            End If
      LeftFlipper.RotateToStart       'If the above conditions are true the flipper goes down.
             PlaySoundat SoundFXDOF("flipperdown",128,DOFOff,DOFFlippers), LeftFlipper
    end if
  End If
  If  Keycode = RightFlipperKey then
    Stopsound "BuzzR"
    if InProgress = True And Tilt=False Then
            If Show_Hands = 1 AND ShowDT = True Then
            rld = 6:hand_right.enabled = 1
            End If
      RightFlipper.RotateToStart      'If the above conditions are true the flipper goes down.
      RightFlipper1.RotateToStart     'If the above conditions are true the flipper goes down.
      DOF 129, 0
      PlaySoundat SoundFXDOF("flipperdown",129,DOFOff,DOFFlippers), RightFlipper
    end if
  End If

  ' process a high score initial entry key if applicable
  if EnteringHS then HighScoreInitKeyUp keycode

  ' process an operator menu keystroke if the menu is up
  if OpMenuActive then OpMenuKeyUp keycode

End Sub

Sub MaxCredits()
  If  Credits>9 Then                  'If credits are greater than 9 then you have 9 credits.
    Credits=9
        If B2SOn Then
    Controller.B2ssetCredits Credits
        End If
  End If
End Sub

Sub Addcredit()
  Credits=credits + 1             'Add a credit.
    MaxCredits                      'Call the max credits sub to check for maximum credits.
    PlaySound SoundFXDOF("knocker-echo",127,DOFPulse,DOFKnocker)
    If B2SOn Then
    Controller.B2ssetCredits Credits
    End If
  if InProgress=False then SaveInfo
End Sub

'*************************** Desktop Options Menu ************************************************************************

Sub DTOptions_Show
    For each object in DToptions
    Object.visible = 1
    Next
End Sub

Sub DTOptions_Hide
    For each object in DToptions
    Object.visible = 0
    Next
End Sub


'********If set in the script options - pressing flippers will dim the gi lights - this timer rebrightens them 500ms later.*

Sub GI_Bright_Timer()
For Each gxx in GI_Lights
    gxx.Intensityscale = 1.1
    Next
    DOF 166,0:DOF 165,1
    me.enabled = 0
End Sub

'*************************** New Game ************************************************************************************

Sub StartGame
    ' If we're entering high score initials, or the operator menu is running,
  ' the Start button doesn't start a game after all.
  if EnteringHS or OpMenuActive then exit sub

  ' the button only works if we have credits or the machine is set to free play
  If Credits > 0 or OptFreePlay = 1 Then
    ' if no game is in progress, start a new game
    If Not InProgress Then
      NewGame                         'Call the New Game sub.
      InProgress = True               'This means a game is in progress.
      Playsound "Start1"
      Playsound "Motor"
    End If

    ' now add a player, if we haven't already added all possible players
    If InProgress = True And Players < 2 And Ball = 1 Then
      if credits > 0 then Credits = Credits - 1  ' Subtract a credit; if 0 credits, leave it at 0 as we must be on free play
            If B2SOn Then
      Controller.B2ssetCredits Credits
            End If
      Players = Players + 1           'Add a player.
      Playsound "bally-addplayer"
            If B2SOn Then
      Controller.B2ssetCanplay 31,Players
            End If
            'Turn on the appropriate lights in desktop mode for a 1 or two player game. These lights are not seen in cabinet mode due to
            'the collection of desktop items being hidden.
            If Players = 1 Then
            Plights1.State = 1
            Plights2.State = 0
            Else
            Plights1.State = 1
            Plights2.State = 1
            End If
    End If
  End If
End Sub

Sub NewGame                          ' Start game, kickass I found a quarter!
    Ball=1                           ' Ball 1.
  For Each Obj In Toplights:Obj.State=1:Next    ' Top lanes are all lit at the start of each ball
  For Each Obj In Sidelights: Obj.State=0:Next  ' Side lanes are unlit
  For Each Obj In Toprolloverlights: Obj.State=0:Next      ' Top rollovers are unlit
  For Each Obj In Bottomrolloverlights: Obj.State=0:Next   ' Bottom rollovers are unlit
  'For Each Obj In TiltedObjects:Obj.disabled = False: Next ' Clear tilt flags from all relevant objects
    Tilt_Disable' Clear tilt flags from all relevant objects
  For Each Obj In Alleylights: Obj.State=0: Next           ' Turn off all Aladdin's Alley score level lights
    l_gi2.state = 1:l_gi1.state = 1
    Alley1.State=1                   ' Turn on the level 1 Aladdin's Alley score light
    Light1k.State=1                  ' Turn on the 1000 bonus light
    Shootagain.State=0               ' Turn off the Shoot Again playfield light
    Light2x.State=0                  ' Turn off the Double Bonus playfield light
    Light5.State=0
    Player=1                         ' Go to Player1.
    Tilt=0                           ' Reset the Tiltcount.
    Tilt=False                       ' Reset the tilt..
  for i = 1 to ubound(ScorePlayer)
    ScorePlayer(i) = 0       ' Reset our internal score counter to zero
    If B2SOn Then
        Controller.B2Ssetscore i, 0  ' Reset the backglass score display to zero
    End If
        Replay1Awarded(i) = false    ' We haven't awarded any replay at level 1 yet to any players
    Replay2Awarded(i) = false    ' Likewise replay level 2
    OverTheTop(i) = false        ' We haven't gone "Over the Top" yet for any player
  next
    DelayTimer1.Enabled=true
    If B2SOn Then
    Controller.B2ssetgameover 35,0   ' turn off the Game Over backglass light
    Controller.B2ssetTilt 33,0       ' turn off the Tilt backglass light
    Controller.B2ssetmatch 34, 0     ' turn off the match lights
    Controller.B2ssetballinplay 32,Ball    ' Ball 1 light on
    End If
    BallReel.setvalue(Ball)
    If B2SOn Then
    Controller.B2ssetplayerup 30,Player    ' Player Up light on
    Controller.B2ssetshootagain 36,0       ' turn off the Shoot Again backglass light
  Controller.B2SSetData 9,0        ' turn off the shared Over The Top light
    Controller.B2SSetData 10,0       ' turn off the individual player 1 OTT light
    Controller.B2SSetData 11,0       ' turn off the individual player 2 OTT light
    End If
    'Set both desktop score reels to 0.
    SReels(1).setvalue(0)
    SReels(2).setvalue(0)
  ' While a game is in progress, turn off the rotating high scores, and
  ' instead show the lowest high score above the current player's score
  ' at the start of each ball.  This lets the player see the next goal
  ' score at a glance, and eliminates any performance impact of updating
  ' the display repeatedly during play.  At the start of a new game, the
  ' player's score starts at zero.
  HighScoreDisplayTimer.Enabled = False
  DispNextHighScoreAbove 0
End Sub

Sub DelayTimer1_Timer()
    PlaySoundAt SoundFXDOF("Ballrelease",130,DOFPulse,DOFContactors), Ballrelease
    BonusCounter = 1000             'Bonus=1000 to start.
    Ballrelease.CreateBall          'Make a new ball.
    Ballrelease.Kick 45,5           'Kick ball out to direction,strength.
    DelayTimer1.Enabled=False
    RandomSoundMetal
End Sub

Sub Trigger1_Hit()
  DOF 124, 1
End Sub

Sub Trigger1_UnHit()
  Playsoundat "Plunger", Screw60
  DOF 124, 0
End Sub

'*********Sound Effects**************************************************************************************************
                                       'Use these for your sound effects like ball rolling, etc.

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub RandomSoundHole()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySound "fx_Hole1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_Hole2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_Hole3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 4 : PlaySound "fx_Hole4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub RandomSoundMetal()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_metal_hit_1"
    Case 2 : PlaySound "fx_metal_hit_2"
    Case 3 : PlaySound "fx_metal_hit_3"
  End Select
End Sub

Sub Metals_Hit(idx)
RandomSoundMetal
End Sub
'******************* Bumpers **********************************************************************************************

Sub bumper1_Hit()
if Tilt = False Then
AddScore100
PlaySoundAt SoundFXDOF("Bumper",119,DOFPulse,DOFContactors), Primitive1
end if
End Sub

Sub bumper2_Hit()
if Tilt = False Then
AddScore100
PlaySoundAt SoundFXDOF("Bumper",120,DOFPulse,DOFContactors), Primitive2
end if
End Sub

Sub bumper3_Hit()
if Tilt = False Then
AddScore100
PlaySoundAt SoundFXDOF("Bumper",121,DOFPulse,DOFContactors), Primitive3
end if
End Sub

'********************Tilt*************************************************************************************************

Sub BumpIt
  If  InProgress = True And Tiltcount < 3 Then
        ' The game is In Progress and not tilted yet.  Count each time you "shake" the game,
    ' and tilt if we reach 3 bumps in within too short a time span.
    Tiltcount = Tiltcount + 1       ' add 1 to the tiltcount while starting the timer.
    TiltTimer.Enabled = True
  End If
  If  Tiltcount > 2 Then
        ' The tilt count (times you've shaken the game) has reached 3 within too short a
    ' time window.  Tilt the game.
    Tilt = True
  End If
  If  Tilt= True Then                 ' If you've tilted the game do this.
    TiltIt
  End If
End Sub

Sub TiltIt
  if InProgress then
    LeftFlipper.RotateToStart       'If the game is tilted then reset the flippers.
        PlaySoundAt SoundFXDOF("FlipperDown",128,DOFOff,DOFFlippers), LeftFlipper
        If Show_Hands = 1 AND ShowDT = True Then
        hld = 6:hand_left.enabled = 1
        End If
        If Show_Hands = 1 AND ShowDT = True Then
        rld = 6:Hand_Right.enabled = 1
        End If
    RightFlipper.RotateToStart      'If the game is tilted then reset the flippers.
    RightFlipper1.RotateToStart     'If the above conditions are true the flipper goes down.
    PlaySoundAt SoundFXDOF("FlipperDown",129,DOFOff,DOFFlippers), RightFlipper
    StopSound "BuzzL" : StopSound "BuzzR"
    BonusCounter=0                  'If the game is tilted then reset the bonus to 0.
    Playsound "Reset3"
    For Each Obj In Toplights:Obj.State=0:Next
    For Each Obj In Sidelights: Obj.State=0:Next
    For Each Obj In Alleylights: Obj.State=0: Next
    For Each Obj In Bottomrolloverlights: Obj.State=0: Next
    For Each Obj In Toprolloverlights: Obj.State=0:Next
    For Each Obj In Bonuslights: Obj.State=0:Next
        l_gi2.state = 0:l_gi1.state = 0
    Alley1.State=0
    Light1k.State=0
    Shootagain.State=0
        If B2SOn Then
    Controller.B2ssetshootagain 36,0
    End If
        Light2x.State=0
        Tilt_Enable 'Activate the object array which prevents dynamic table objects from interacting with the ball.
    'For Each Obj In TiltedObjects:Obj.disabled = True: Next'If the game is tilted disable all these objects in this collection.
    If B2SOn Then
        Controller.B2ssetTilt 33,1
        End If
  End If
End Sub

Sub Tilt_Enable
Bumper1.collidable = 0
Bumper2.collidable = 0
Bumper3.collidable = 0
Bumper1_Wall.isdropped = 0
Bumper2_Wall.isdropped = 0
Bumper3_Wall.isdropped = 0
LeftSlingShot_Wall.isdropped = 0
Gotilt.setvalue(1)
End Sub

Sub Tilt_Disable
Bumper1.collidable = 1
Bumper2.collidable = 1
Bumper3.collidable = 1
Bumper1_Wall.isdropped = 1
Bumper2_Wall.isdropped = 1
Bumper3_Wall.isdropped = 1
LeftSlingShot_Wall.isdropped = 1
Gotilt.setvalue(2)
End Sub

Sub TiltTimer_Timer
    TiltTimer.Enabled = False       'Turn off/reset the timer.
  If  Tiltcount=3 Then                'If the tilt count reches 3 then the game is tilted,
        Tilt = True                     'We have tilted the game.
  Else
    Tiltcount = 0                   'Else tilt count is 0 and reset the tilt timer.
  End If
End Sub


'********************* Scores and scoring *******************************************************************************

' Score 10 points, playing the 10-point bell (or firing the DOF chime)
Sub AddScore10
    PlaySound SoundFXDOF("Bell10",131,DOFPulse,DOFChimes)
  AddScore 10
End Sub

' Score 100 points, playing the 100-point bell (or firing the DOF chime)
Sub AddScore100
  PlaySound SoundFXDOF("Bell100",132,DOFPulse,DOFChimes)  ' 100-point chime
  AddScore 100
End Sub

' Score 1000 points, playing the 1000-point bell (or firing the DOF chime)
Sub AddScore1000
  PlaySound SoundFXDOF("Bell1000",133,DOFPulse,DOFChimes)
  AddScore 1000
End Sub

' Add points to the current player's score.  Counts the points, updates
' the backglass reels, and checks for replays.
Sub AddScore(Points)
  ScorePlayer(Player) = ScorePlayer(Player) + Points
    If InProgress Then
    SReels(Player).addvalue(Points)
    End If
    If B2SOn Then
  Controller.B2Ssetscore Player, ScorePlayer(Player) mod 100000
  End If
    CheckReplay
End Sub

' Check all players' current scores for replays and "Over The Top".
Sub CheckReplay
  If ScorePlayer(Player) >= 100000 and Not OverTheTop(Player) then
    OverTheTop(Player) = true
    Playsound "overthetop"
    if AuthenticOverTheTop then
      ' Authentic mode - this is the way the original machine worked.
      ' In this mode, there's a single Over The Top light that's shared
      ' for all players, and it only comes on for about 5 seconds after
      ' the score rolls over.  The single light has B2S ID 9.
            If B2SOn Then
      Controller.B2SSetData 9, 1
            End If
            Gotilt.SetValue(3) 'Set OTT in the desktop reel.
      ' put 5 seconds on the light clock and start the timer
      OverTheTopTime = 5
      OverTheTopTimer.Enabled = true
    else
      ' Modernized mode.  Each player has a separate Over The Top light,
      ' which turns on and stays on until a new game is started.  Simply
      ' turn on the player's light.  The B2S lamps for these are numbered
      ' consecutively from 10 for player 1.  This mode isn't true to the
      ' original EM machine, but it's a nicer UI in that tells us which
      ' player's score is over 100k, and provides a permanent indicator,
      ' effectively serving as a limited sixth reel in the score.
            If B2SOn Then
      Controller.B2SSetData 9 + Player, 1
            End If
            Gotilt.SetValue(3) 'Set OTT in the desktop reel.
    end if
  End If
  If OptReplay1 <> 0 And ScorePlayer(Player) >= OptReplay1 And Not Replay1Awarded(Player) Then
    AwardReplay
    Replay1Awarded(Player) = True
  End If
  If OptReplay2 <> 0 And ScorePlayer(Player) >= OptReplay2 And Not Replay2Awarded(Player) Then
    AwardReplay
    Replay2Awarded(Player) = True
  End If
End Sub

' Over The Top timer.  In authentic mode, the Over The Top light goes out
' after about 5 seconds.  The timer fires once per second, so we take one
' second off of the remaining time on each call, and turn the light off
' when the clock reaches zero.
Dim OverTheTopTime
Sub OverTheTopTimer_Timer
  if OverTheTopTime > 0 then
    ' there's still time on the clock - take a second off and keep going
    OverTheTopTime = OverTheTopTime - 1
  else
    ' the clock has reached zero - turn off the light and kill the timer
        If B2SOn Then
    Controller.B2SSetData 9, 0
        End If
        Gotilt.SetValue(2) 'Turn off the OTT reel in desktop mode.
    OverTheTopTimer.Enabled = false
  end if
End Sub

' Award a Special.  This awards either a credit or an extra ball, depending
' on the option setting.
Sub AwardSpecial
  if OptSpecialVal = 0 then
    ' 0 -> Special awards credit
    Addcredit
  else
    ' 1 -> Special awards extra ball
    AwardExtraBall
  end if
End Sub

' Award a replay for reaching a replay score level.  Awards an extra ball or
' credit, depending on the Replay Value option setting (OptReplayVal).
Sub AwardReplay
  if OptReplayVal = 0 then
    ' 0 -> Replay awards credit
    AddCredit
  else
    ' 1 -> Replay awards extra ball
    AwardExtraBall
  end if
End Sub

' Award an extra ball
Sub AwardExtraBall
  ShootAgain.State=1
    If B2SOn Then
  Controller.B2ssetshootagain 36,1
    End If
End Sub



' ************************* Score Motor *********************************************************************
'
' This simulates the score motor mechanism used on the real machine to award
' multiples of 100 and 1000 points.  It's also used to award all scores in
' Aladdin's Alley.
'
' Score motors were commonly used on EM machines of this generation.  These
' were used to handle scoring of quantities that couldn't be represented with
' a single click of one reel.  For example, if a target scores 300 points,
' it's necessary to advance the 100's reel by three clicks.  This requires
' three separate pulses to the solenoid that drives the reel, spread out
' enough over time to give the solenoid time to spring back to its starting
' position between each pulse.  EM machines didn't have microprocessors for
' counting out the pulses - they had to do everything mechanically.  The
' mechanical solution to this problem was the score motor.
'
' A score motor is simply an electric motor connected to cam.  As the cam
' rotates, it triggers a series of switches in sequence.  Each switch can
' be wired to carry out some scoring function when tripped.  So for our
' 300-point target, the target's switch would be connected to a set of
' relays that would latch on, providing power to the score motor and
' connecting three of the motor cam switches to the 100's reel.  The
' power connection would set the motor spinning; as it went around, it
' would trip the cam switches, sending three pulses to the 100's reel
' through the latched relays.  At the end of the cycle, the last cam
' switch would interrupt power to the latching relays, which would reset
' the relays, cutting power to the motor.  That would end the cycle and
' leave the motor in a position to handle the next target that needs it.
'
' That's the general score motor mechanism from pinballs of this era.  With
' VP, we don't really need to simulate the score motor faithfully, since we
' have a modern electronic computer at our disposal which can easily handle
' the arithmetic that the old score motors carried out mechanically.  However,
' emulating the underlying score motor mechanism is still useful for simulating
' the timing of the scoring chimes and backglass reel animation.
'
' For Aladdin's Castle in particular, there's a special benefit to emulating
' the full score motor behavior.  In the Aladdin's Alley feature, we have an
' unusual arrangement with two rollover switches.  When the ball rolls through
' the alley, it will hit both switches - but we only want *one* score award
' per traversal.  On the real machine, this is accomplished by way of the
' score motor.  When the ball comes rolling through the alley and rolls over
' the first switch, that switch will start a score motor cycle to award the
' current alley score.  As the ball continues through the alley, when it hits
' the second rollover, that rollover will also trigger the same relays as the
' first switch - but all of that will have no effect because the score motor
' is already running.
'
' The original version of this virtual table used a single hidden trigger object
' to handle each alley score.  For the most part, that worked just fine, in that
' it reproduced the usual case above where we only want a single score award per
' alley traversal, despite the fact that the ball hits both rollovers.  However,
' the old approach *always* ensured that a single score value was awarded, whereas
' the original machine's score motor setup only ensured this *most* of the time.
' A very slow traversal of the alley could actually collect two separate score
' awards on the real machine.  If the ball was moving so slowly that it didn't
' trigger the second rollover until after the score motor cycle had fully
' completed for the first rollover, the second rollover would be able to trigger
' another score motor cycle and collect a second award.  This new more faithful
' simulation of the score motor will properly simulate that effect, in addition
' to correctly awarding only one score value for the typical case where the ball
' is moving more quickly.

' Score motor virtual relay circuits.  When set to true, these tell the score
' motor to award certain point values at corresponding positions in the motor
' cycle.  All of these are automatically unlatched (set to false) at the end
' of each motor cycle.
'
' scmHundred() and scmThousand() are arrays of 5 relays, virtually wired to the
' cam switches at the corresponding positions. If scmHundred(N) is true, the
' score motor will score 100 points when at position N in the cycle.  So to
' award 300 points, we'd set scmHundred(1), scmHundred(2), and scmHundred(3)
' to true and start the score motor.  This will award 100 points apiece at
' positions 1, 2, and 3, for a total of 300 points.
'
' scmSpecial, if set to true, awards a Special at motor position 1.
'
' scmAlley, if set to true, advances the Aladdin's Alley scoring level by
' one slot at motor position 5.
'
Dim scmHundred(5), scmThousand(5), scmSpecial, scmAlley

' Current score motor position.  The motor cycle starts at 1 and ends at 6.
' When the motor isn't running, this is at 1, the start of the cycle.
Dim scmPos : scmPos = 1

' Add points via the score motor.  This can add any combination of 100s and
' 1000s, up to 5000 plus 500.  This (virtually) latches the relays for the
' first N hundred and thousand cam switches, then starts the score motor cycle.
' If a cycle is already in progress, this latches the relays but doesn't have
' any effect on the motor cycle - it just leaves the motor running where it
' already is.
Sub ScoreMotorAdd(points)
  for i = 1 to 5
    if points >= 1000 then
      scmThousand(i) = true
      points = points - 1000
    end if
  next
  for i = 1 to 5
    if points >= 100 then
      scmHundred(i) = true
      points = points - 100
    end if
  next
  ScoreMotorStart
End Sub

' Award a special via the score motor.  This latches the Special virtual relay
' and starts a score motor cycle.
Sub ScoreMotorAddSpecial
  scmSpecial = true
  ScoreMotorStart
End Sub

' Start the score motor.  If the motor is already running, this has no effect.
' If the motor isn't running, it plays the "rollover buzz" effect that's meant
' to imitate the sound of the score motor starting, and starts the motor timer.
Sub ScoreMotorStart
  if not ScoreMotorTimer.Enabled then
    Playsound "RolloverBuzz"
    ScoreMotorTimer.Enabled = true
  end if
End Sub

' Score motor timer.  This is invoked every motor timer interval when the
' score motor is running.  We award the score values for all latched relays
' for the current motor cam position, and advance to the next position.  When
' we reach the end of the cycle, we unlatch all of the relays, reset the
' motor cam to the first position, and turn off the motor by disabling the
' timer.
Sub ScoreMotorTimer_Timer
  if scmPos < 6 then
    ' in position 1, award a special if activated
    if scmPos = 1 and scmSpecial then AwardSpecial

    ' in positions 1-5, award multiples of 100 and 1000 points if activated
    if scmHundred(scmPos) then AddScore100
    if scmThousand(scmPos) then AddScore1000

    ' in position 5, advance the Aladdin's Alley score level
    if scmPos = 5 and scmAlley then AdvanceAladdinsAlley

    ' move to the next position
    scmPos = scmPos + 1
  else
    ' end of cycle - turn off the motor and clear all latching relays
    ScoreMotorTimer.Enabled = false
    for i = 1 to 5 : scmHundred(i) = false : scmThousand(i) = false : next
    scmSpecial = false
    scmAlley = false
    scmPos = 1
  end if
End Sub


'******************** Table Objects ****************************************************************************************

Sub Toplanes_Hit(x)                 'Any trigger in the collection has been hit.
        PlaySoundAtBall "fx_Sensor"
  If  Tilt= False Then                'If the game isn't tilted.
    Toplights(x).State=0              'Turn OFF the lane light
    Toprolloverlights(x).State=1      'Turn ON the corresponding top rollover light
    Bottomrolloverlights(x).State=1   'Turn ON the corresponding bottom rollover light.
    ScoreMotorAdd 300                 'Score 300 on the score motor
    DOF 105+x, 1
    Checklights
    AdvanceBonus
  End If
End Sub

Sub Toplanes_Unhit(x)
  DOF 105+x, 0
End Sub

Sub Checklights()
  If  Tilt=False Then
        ' if A-B lane lights are OFF, turn on the spinner and outlane lights
    If  Light1.State=0 And Light2.State=0 Then
    For Each Obj In Sidelights: Obj.State=1:Next
      Light5.State=1
    End If

        ' if C-D lane lights are OFF, turn on the double bonus
    If  Light3.State=0 And Light4.State=0 Then
      Light2x.State=1
    End If

        ' if all A-B-C-D lane lights are OFF, turn on extra ball
    If  Light1.State=0 And Light2.State=0 And Light3.State=0 And Light4.State=0 Then
      AwardExtraBall
    End If
  End If
End Sub

Sub Toptriggers_Hit(x)
  If  Tilt=False Then
    DOF 101+x, 1         ' DOF events 101-104 for the A-D rollovers
    AddScore100
    If  Toprolloverlights(x).State=1 Then
      AdvanceBonus
    End If
  End If
End Sub

Sub Toptriggers_Unhit(x)
  DOF 101+x, 0
End Sub

Sub Bottomtriggers_Hit(x)
  If  Tilt=False Then
    DOF 109+x, 1
    AddScore100          ' DOF events 109-112 for lower A-D rollovers
    If  Bottomrolloverlights(x).State=1 Then
      AdvanceBonus
    End If
  End If
End Sub

Sub Bottomtriggers_Unhit(x)
   DOF 109+x, 0
End Sub


Dim RStep, Lstep, Tstep, Blstep, TLstep, BRstep, TRstep

Sub LeftSlingShot_Slingshot
  if Tilt = False then
    PlaySoundAt SoundFXDOF("Slingshot",122,DOFPulse,DOFContactors), Sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval  = 10
    Addscore 10
    End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Rani_W_Hit()
    Rani.Visible = 0
    Rani1.Visible = 1
    TStep = 0
    Rani_W.TimerEnabled = 1
  Rani_W.TimerInterval  = 10
End Sub

Sub Rani_W_Timer
    Select Case TStep
        Case 3:Rani1.Visible = 0:Rani2.Visible = 1
        Case 4:Rani2.Visible = 0:Rani.Visible = 1:Rani_W.TimerEnabled = 0
    End Select
    TStep = TStep + 1
End Sub

Sub BLani_W_Hit()
    BLani.Visible = 0
    BLani1.Visible = 1
    BLStep = 0
    BLani_W.TimerEnabled = 1
  BLani_W.TimerInterval  = 10
End Sub

Sub BLani_W_Timer
    Select Case BLStep
        Case 3:BLani1.Visible = 0:BLani2.Visible = 1
        Case 4:BLani2.Visible = 0:BLani.Visible = 1:BLani_W.TimerEnabled = 0
    End Select
    BLStep = BLStep + 1
End Sub

Sub TLani_W_Hit()
    BLani.Visible = 0
    BLani3.Visible = 1
    TLStep = 0
    TLani_W.TimerEnabled = 1
  TLani_W.TimerInterval  = 10
End Sub

Sub TLani_W_Timer
    Select Case TLStep
        Case 3:BLANI3.Visible = 0:BLani4.Visible = 1
        Case 4:BLANI4.Visible = 0:BLani.Visible = 1:TLani_W.TimerEnabled = 0
    End Select
    TLStep = TLStep + 1
End Sub


Sub BRani_W_Hit()
    BRani.Visible = 0
    BRani1.Visible = 1
    BRStep = 0
    BRani_W.TimerEnabled = 1
  BRani_W.TimerInterval  = 10
End Sub

Sub BRani_W_Timer
    Select Case BRStep
        Case 3:BRani1.Visible = 0:BRani2.Visible = 1
        Case 4:BRani2.Visible = 0:BRani.Visible = 1:BLani_W.TimerEnabled = 0
    End Select
    BRStep = BRStep + 1
End Sub

Sub TRani_W_Hit()
    BRani.Visible = 0
    Trani3.Visible = 1
    TRStep = 0
    TRani_W.TimerEnabled = 1
  TRani_W.TimerInterval  = 10
End Sub

Sub TRani_W_Timer
    Select Case TRStep
        Case 3:TRANI3.Visible = 0:TRani4.Visible = 1
        Case 4:TRANI4.Visible = 0:BRani.Visible = 1:TRani_W.TimerEnabled = 0
    End Select
    TRStep = TRStep + 1
End Sub


Sub Slingshots_Hit(IDX)       'If a Slingshot is hit...
  If  Tilt= False Then            'If the game isn't tilted.
    AddScore10                  'Add the score.
  End If
End Sub

Sub Spinner1_Spin()                 'If the Spinner is hit...
    PlaySoundat "Fx_Spinner", PegPlasticT36
  If Tilt= False Then                'If the game isn't tilted.
    If  Light5.State=1 Then
      DOF 126, 2
      AddScore100                    'Add the score.
    Else
      DOF 125, 2
      AddScore10
    End If
  End If
End Sub

Sub LeftInlane_Hit()                ' Inlane has been hit.
  PlaySoundAtBall "fx_Sensor"
  If  Tilt= False Then            ' If the game isn't tilted.
    ScoreMotorAdd 300           ' score 300 on the score motor
    DOF 115, 1                  ' fire the DOF effect
    AdvanceBonus
  End If
End Sub

Sub LeftInlane_Unhit()
  DOF 115, 0
End Sub

Sub Targa_Hit()
    PlaySoundAt SoundFX("target",DOFTargets), TargA
  If  Tilt=False Then
  ScoreMotorAdd 300            ' add 300 on the score motor
  AdvanceBonus
  End If
End Sub

Sub TargB_Hit()
    PlaySoundAt SoundFX("target",DOFTargets), TargB
  If  Tilt=False Then
  ScoreMotorAdd 300            ' add 300 on the score motor
  AdvanceBonus
  End If
End Sub

Sub Prim_Gates_Hit(IDX)                  'Outlane has been hit.
    PlaySoundAtBall "fx_gate"
End Sub

Sub Outlanes_Hit(IDX)                  'Outlane has been hit.
    PlaySoundAtBall "fx_Sensor"
  If  Tilt=False Then
    If  Sidelight1.State=1 Then
      ScoreAladdinsAlley
    Else
      AddScore1000
    End If
  End If
End Sub

Sub s_LeftOut_Hit
  if Tilt = false then DOF 114, 1
End Sub
Sub s_LeftOut_Unhit
  if Tilt = false then DOF 114, 0
End Sub

Sub s_RightOut_Hit
  if Tilt = false then DOF 116, 1
End Sub
Sub s_RightOut_Unhit
  DOF 116, 0
End Sub


' ******************* Aladdin's Alley **********************************************************

' Alley rollovers, left and right.  Note that we award the alley score on EITHER
' rollover, even though the ball usually hits BOTH of them each time through the
' alley.  The score will still only be awarded once in most cases, because the
' score is always awarded through the score motor.  The score motor takes about
' a second to cycle through one alley score, which is long enough that the second
' rollover is usually hit while the cycle from the first rollover is still in
' progress.  This means that the second rollover will usually be ignored.  When
' the ball is moving very slowly, it can trigger a new score when hitting the
' second rollover, though.  This is consistent with the real machine, which had
' exactly the same timing properties for the alley scoring.
Sub s_AALeft_Hit()
  ScoreAladdinsAlley
End Sub

Sub s_AARight_Hit()
  ScoreAladdinsAlley
End Sub


' Aladdin's Alley scoring.  This has two possible patterns, depending on
' the LIBERAL or CONSERVATIVE setting (OptAlleyScoring - 0=liberal, 1=conservative):
'
'  LIBERAL (0):       500 - 1000 - 2000 - 3000 - 4000 - 5000 - SPECIAL...
'  CONSERVATIVE (1):  500 - 1000 - 2000 - 3000 - 4000 - 5000 - SPECIAL - 5000 - SPECIAL - 5000...
'
' In either case, when reaching the last state, we stay on that state for the
' remainder of the ball in play.  Note that SPECIAL awards a credit or extra ball,
' according to the Special Value option.
'
' To handle the progressive scoring, we keep two arrays of score levels, one for each
' mode, and use the appropriate array for the current mode.  Each element in each array
' is an index into our AlleyScores array.
Dim LiberalScores : LiberalScores = Array(0, 1, 2, 3, 4, 5, 6)
Dim ConservScores : ConservScores = Array(0, 1, 2, 3, 4, 5, 6, 5, 6, 5)

' The AlleyScores gives us the scoring level and playfield indicator light object for each
' possible scoring level.  A score of 0 means Special; other values are the actual number
' of points to award.  Note how each score value is paired with the corresponding playfield
' light object - that tells us which light to turn on to indicate the matching score level.
Dim AlleyScores : AlleyScores = Array( _
  500, Alley1, _
  1000, Alley2, _
  2000, Alley3, _
  3000, Alley4, _
  4000, Alley5, _
  5000, Alley6, _
  0, Alley7 _
)

Sub ScoreAladdinsAlley()
  If Tilt=False Then
    ' fire the alley DOF effects
    DOF 113, 2

    ' get the appropriate scoring list - liberal or conservative
    Dim sl : if OptAlleyScoring then sl = ConservScores else sl = LiberalScores

    ' get the AlleyScores index for the current position in the current mode
    Dim idx : idx = sl(AlleyLevel)*2

    ' award the score for the current level
    Dim s : s = AlleyScores(idx)
    if s = 0 then
      ' Special - award the credit or extra ball (depending on option settings)
      ScoreMotorAddSpecial
      scmAlley = true

      ' Award any desired points on Special.  The real machine doesn't award any
      ' points - just the credit or extra ball.  However, the old version of this
      ' virtual table awarded 5000 points.  I think that was just a misunderstanding
      ' of the original rules, but I'm keeping the option in case anyone really wants
      ' the points.  Set the constant AlleyPointsOnSpecial to the desired score (in
      ' the options section near the top of the script).
      if AlleyPointsOnSpecial then ScoreMotorAdd AlleyPointsOnSpecial
    else
      ScoreMotorAdd s
      scmAlley = true
    end if
  End If
End Sub

Sub AdvanceAladdinsAlley()
  ' get the appropriate scoring list - liberal or conservative
  Dim sl : if OptAlleyScoring then sl = ConservScores else sl = LiberalScores

  ' Move to the next award level.  When we reach the end of the score list,
  ' we stay on the last entry from then on.
  if AlleyLevel < ubound(sl) then
    ' turn off the alley light for the old level
    AlleyScores(sl(AlleyLevel)*2 + 1).State = 0

    ' move to the new level
    AlleyLevel = AlleyLevel + 1

    ' turn on the alley light for the new level
    AlleyScores(sl(AlleyLevel)*2 + 1).State = 1
  End If
End Sub


'********* Drain *****************************************************************************************************

' Start a new ball
Sub AddBall()
  If Shootagain.State=1 Then         'If the Shoot Again light is on then subtract a player. The Bonus Ended sub will add
    Player=Player-1 'a player to go to the next player, subtracting a player counteracts this, keeping the same player.
        If B2SOn Then
    Controller.B2ssetshootagain 36,0
    Controller.B2ssetballinplay 32,Ball
        Controller.B2ssetplayerup 30,Player
        End If
    BallReel.setvalue(Ball) ' Set the ball number on the desktop reel.
  End If
End Sub

' When the ball drains, start the bonus
Sub Drain_Hit()                     'Another one bites the dust.
    Drain.DestroyBall               'Destroy the ball when it hits the drain.
    PlaySoundat "Drain5", drain             'Play the sound.
  DOF 123, 2

  ' Start the bonus timer to simulate the mechanical bonus scoring motor.
  ' For regular single bonus, award points in batches of 5.  For the
  ' double bonus, award the bonus in batches of 2.  Between batches,
  ' we'll introduce a suitable delay for the motor to cycle.
    BonusTimer.Enabled = 1
  If Light2x.State = 0 Then BonusMotor = 5 Else BonusMotor = 2
End Sub

' When the bonus ends, proceed to the next ball
Sub BonusEnded()
  Addball                         'Call the Addball sub to check for Extra Ball.
    Player=Player+1                 'Go to the next player.
    BonusCounter = 1000             'Reset the bonus to 1000 for the next ball.
    Light1k.State = 1               'Turn on the 1000 light for the next ball.
    AlleyLevel = 0                  'Reset the Aladdin's Alley progressive score to the first level
  'For Each Obj In TiltedObjects:Obj.disabled = False: Next 'If things were disabled by tilt turn them back on.
    Tilt_Disable
  If  Player > Players Then           'If the player number exceeds the number of players then default to the next ball.
    Player = 1                      'Return to player1.
    Ball = Ball + 1                 'Go to the next ball.
        If B2SOn Then
    Controller.B2ssetballinplay 32,Ball
        End If
    BallReel.setvalue(Ball) ' Set the ball number on the desktop reel.

  End If
  If  Ball <= OptBallsPerGame Then   'Is ball played less than total Balls or is the Game Over.
    DelayTimer2.Enabled=True
    Playsound "Motor"
    For Each Obj In Toplights:Obj.State=1:Next        ' top lanes start each ball ON
    For Each Obj In Sidelights: Obj.State=0:Next
    For Each Obj In Alleylights: Obj.State=0: Next
    For Each Obj In Bottomrolloverlights: Obj.State=0: Next
    For Each Obj In Toprolloverlights: Obj.State=0:Next
        l_gi2.state = 1:l_gi1.state = 1
    Alley1.State=1
    Light1k.State=1
    Shootagain.State=0
    Light2x.State=0
    Light5.State=0
        BallReel.setvalue(Ball) ' Set the ball number on the desktop reel
    If B2SOn Then
        Controller.B2ssetballinplay 32,Ball
        Controller.B2ssetshootagain 36,0
    Controller.B2ssetplayerup 30,Player
    Controller.B2ssetTilt 33,0
        End If
    Tilt=False                      'Reset the tilt.
    Tiltcount=0                     'Reset the tiltcount.

    ' display the next high score above the player's current score
    DispNextHighScoreAbove ScorePlayer(Player)

  Else                                'Sorry, but the game is now over.
    GameOver                        'Go to game over to finish up.
    Light1k.State=0
  End If
End Sub

' Ball release timer
Sub DelayTimer2_Timer()
    PlaySoundAt SoundFXDOF("Ballrelease",130,DOFPulse,DOFContactors), Ballrelease
    RandomSoundMetal
    Ballrelease.CreateBall          'Make a new ball.
    Ballrelease.Kick 23,6           'Kick ball out to direction,strength.
    DelayTimer2.Enabled=False
    Stopsound "Motor"
End Sub

'************ Bonus Routine ************************************************************************************************

Dim Obj,BonusCounter,MultiplierCounter,Tens,Ones

Sub AdvanceBonus()
    BonusCounter = BonusCounter + 1000
  If  Bonuscounter >15000 Then Bonuscounter=15000
  ShowBonusLights
End Sub

' Bonus 1000s and 10000s lights.  Note that Light0K is a hidden off-screen light
' that we use as a placeholder for zero values, so that nothing visible lights.
' The bonus on this machine is limited to 15000, so we only need one 10K light.
Dim BonusOnesLights : BonusOnesLights = Array( _
  Light0K, Light1K, Light2K, Light3K, Light4K, _
  Light5K, Light6K, Light7K, Light8K, Light9K)
Dim BonusTensLights : BonusTensLights = Array(Light0K, Light10K)

' Show the current bonus value on the playfield lights
Sub ShowBonusLights ()
  ' turn off all of the lights to start fresh
  For Each Obj In BonusLights:Obj.State = 0:Next

  ' Figure the multiple of 1000 and 10000 for the current bonus value.
  Tens = ((Bonuscounter + 500) \ 10000) MOD 10   ' Get the multiple of 10,000
  Ones = ((Bonuscounter + 500) \ 1000) MOD 10    ' Get the multiple if 1,000
  BonusOnesLights(Ones).State = 1                ' Turn on the suitable 1000s light
  BonusTensLights(Tens).State = 1                ' Turn on the suitable 10000s light
End Sub

Sub BonusTimer_Timer
  ' if the bonus has reached zero, we're done - signal the end of the bonus
  ' processing and turn off the bonus timer
  If Bonuscounter = 0 Then
    BonusEnded
    Bonustimer.Enabled=False
    Exit Sub
  End If

  ' Award 1000 per cycle
  AddScore1000

  ' show the current bonus lights
  ShowBonusLights

  ' Adjust the counter.  For double bonus, we want to award points twice per
  ' 1000 on the nominal bonus counter, so only subtract 500 from the counter.
  ' For single bonus, subtract the full 1000 from the counter.
  dim d : if Light2x.State = 1 then d = 500 else d = 1000
    BonusCounter = BonusCounter - d

  ' Count one step of the simulated bonus motor.  If we've reached zero, it
  ' means that we've reached the end of the current mechanical cycle - start
  ' the delay timer that simulates the motor resetting, and turn off the
  ' main bonus timer until the delay routine resets things.
  BonusMotor = BonusMotor - 1
  If BonusMotor = 0 Then
    DelayTimer3.Enabled = True
    BonusTimer.Enabled = False
  End If
End Sub

' Bonus motor cycle reset delay timer.
Sub DelayTimer3_Timer()
  If Light2x.State = 1 Then
    ' Double Bonus.  Enable the delay timer for the double bonus cycle reset.
    DoubleBonusDelay.Enabled=True
  else
    ' Single Bonus.  Start a new batch of 5 cycles through the regular bonus timer.
    BonusMotor = 5
    BonusTimer.Enabled = True
  end if

  ' this is a one-shot timer - disable it now that it has fired
    DelayTimer3.Enabled=False
End Sub

' Double bonus delay timer
Sub DoubleBonusDelay_Timer
  ' Start a new batch of 2 cycles through the regular bonus timer
  BonusMotor = 2
    BonusTimer.Enabled = True

  ' this is a one-shot timer - disable it
    DoubleBonusDelay.Enabled = False
End Sub

'******************** Game Over ********************************************************************************************

Sub GameOver
    StopSound "BuzzL" : StopSound "BuzzR"
  InProgress=False         'We always know when the game is active with this variable.
    If B2SOn Then
    Controller.B2ssetplayerup 30,0
    Controller.B2ssetgameover 35,1
    Controller.B2ssetballinplay 32,0
    End If
  Match                'Call the match sub.
  LeftFlipper.RotateToStart       'If the above conditions are true the flipper goes down.
    PlaySoundAt SoundFXDOF("FlipperDown",128,DOFOff,DOFFlippers), LeftFlipper
  RightFlipper.RotateToStart      'If the above conditions are true the flipper goes down.
  RightFlipper1.RotateToStart     'If the above conditions are true the flipper goes down.
    PlaySoundAt SoundFXDOF("FlipperDown",129,DOFOff,DOFFlippers), RightFlipper
    If Show_Hands = 1 AND ShowDT = True Then
    hld = 6:hand_left.enabled = 1
    End If
    If Show_Hands = 1 AND ShowDT = True Then
    rld = 6:Hand_Right.enabled = 1
    End If

    BallReel.setvalue(0) 'Set the ball count to 0 on the desktop reel.
    Gotilt.setvalue(0) 'Set the game over image on the desktop reel.

  ' mark all players in the game for high score checks
  for i = 1 to ubound(NeedHighScoreCheck)
    NeedHighScoreCheck(i) = (i <= Players)
  next

  ' begin the high score checks
  CheckHighScores

  ' game over - no players in
  Players=0            'No one is playing now
    l_gi2.state = 0:l_gi1.state = 0
End Sub

' Check the next player's high score.  This runs through the HighScoreCheck
' list to find the lowest scoring player whose score we still haven't checked
' against the high score list, and does the check for that player.  We always
' process the LOWEST remaining score first so that we'll detect a winning
' score before a higher-scoring player knocks the lower ranking scores out
' of the list.  This ensures that everyone who beat any of the existing high
' scores gets awarded free game credits, even if they'll end up getting knocked
' out of the list by higher scoring players in this same session.
Dim NeedHighScoreCheck(5)
Sub CheckHighScores
  ' Look for the lowest score still in the game
  dim mi : mi = -1   ' player # with lowest score among those still to be checked
  dim n : n = 0      ' number of players still to be checked
  for i = 1 to ubound(NeedHighScoreCheck)
    ' if this player is still to be checked...
    if NeedHighScoreCheck(i) then
      ' count it
      n = n + 1

      ' if this is the lowest (or only) score we've seen so far, remember this index
      if mi = -1 then mi = i else if ScorePlayer(i) < ScorePlayer(mi) then mi = i
    end if
  next

  ' if we found a winner, process it; otherwise end the high score checks
  if n > 0 then
    ' mark this player as checked
    NeedHighScoreCheck(mi) = False

    ' check the high score for this player
    CheckHighScore mi, n-1
  else
    ' done with all checks - finish up
    EndHighScoreCheck
  end if
End Sub

' End the high score check.  This is called after we've checked the last
' player's score against the high scores and completed entry of initials.
' This is the final step at the end of a game, so we're ready to start a
' new game after this returns.
Sub EndHighScoreCheck
  ' run the high score display timer between games, starting at the top score
  CurHighScoreDisp = 0
  HighScoreDisplayTimer.Enabled = True

  ' schedule a save on a timer - the save stalls things for a couple of
  ' seconds, so we want to let the UI update first, to make the stall
  ' less obvious visually
  SaveTimer.Enabled = true
End Sub

Sub SaveTimer_Timer
  SaveInfo
  SaveTimer.Enabled = false
End Sub

' Number of credits to add for a new high score in slot #n.  (Slot 0 is
' high score 1, slot 1 is high score 2, etc).
Dim HighScoreAwards : HighScoreAwards = Array(2, 1, 1, 1)

' Check one player's score against the high score list.  If the player's score
' is higher than (or ties) any high score in the list, we'll insert the player's
' score in the list ahead of the highest existing score lower than or equal to
' the player's score, pushing the lowest score off of the list.  We'll also award
' free game credits according to the position in the list.
'
' If the player will *remain* in the list when we're done processing the other
' players with higher scores, we'll prompt the player to enter their initials.
' We'll skip this if the player will be pushed out of the list by other players
' in the same game, since it would be pointless.
'
' 'player' is the player number (1, 2...).  'nHigher' is the number of players
' in this round with higher scores than this player.
Sub CheckHighScore(player, nHigher)
  ' get the player's score
  dim s : s = ScorePlayer(player)

  ' Check high scores from the highest down.  We'll insert the new score
  ' ahead of the first one we find that's lower or equal.
  dim i
  for i = 0 to ubound(Highscore)
    ' if the player's score is higher than or equal to high score #i, insert
    ' the new score at #i
    if s >= Highscore(i) then
      ' push all scores below this one down one slot
      dim j
      for j = ubound(Highscore) to i+1 step -1
        Highscore(j) = Highscore(j-1)
        HighscoreDate(j) = HighscoreDate(j-1)
        HighscoreInits(j) = HighscoreInits(j-1)
      next

      ' set the slot to the new high score
      Highscore(i) = s
      dim d : d = Date()
      HighscoreDate(i) = Month(d) & "/" & Day(d) & "/" & Year(d)

      ' award the credits
      AddCredits HighScoreAwards(i)

      ' Determine if this player's initials will remain in the list after
      ' we process the other players in the same game with higher scores.
      ' We always process the lowest score first, so every player remaining
      ' to be checked has a higher score than we do, so every player still
      ' to be checked will push our position down one slot.  This means
      ' that our final position is going to be i + nHigher.  If that's a
      ' valid slot, we'll remain in the list, so enter initials.  Otherwise
      ' just move on to the next player.
      if i + nHigher <= ubound(Highscore) then
        BeginHighScoreInits player, i
      else
        CheckHighScores
      end if

      ' no need to look any further
      Exit Sub
    end if
  next

  ' If we get this far, it means that this player didn't beat any of
  ' the high scores.  Go check the remaining players.
  CheckHighScores
End Sub


' Display a high score value.  n is the high score list index to
' display - 0 for the top score, 1 for the second highest score, etc.,
' up to ubound(Highscore).  Use -1 to display the top score with no
' rank number.
'
' Each digit of the high score is represented in B2S by 10 lamps,
' for value 0-9.  The lamps are numbered like this:
'   240   units
'   241   tens
'   242   hundreds
'   243   thousands
'   244   ten thousands
'   245   hundred thousands
'
' The player initials are displayed similarly in the three elements
' 237, 238, and 240, for the initials in order left to right.  For
' these, value 1=A, 2=B, ..., 26=Z, 27=space.
'
Sub DispHighScore(n)
  dim i, val, s, inits, d, dispRank
If B2Son Then

  ' Set the high score number we're currently showing.  The 0th
  ' entry is "High Score 1", etc.  If n is -1 (or otherwise out
  ' of bounds), it'll translate to a non-existent light number on
  ' the backglass, so we'll just display a blank rank number.
  Controller.B2SSetData 228, n+1

  ' if n is -1 (or otherwise out of bounds), display the top score
  if n < 0 or n > ubound(Highscore) then n = 0

  ' get the current score to display
  s = Highscore(n)
  inits = HighscoreInits(n)
  d = HighscoreDate(n)

  ' set the score digits
  val = s
  for i = 0 to 5
    if val <> 0 or i = 0 then
      dim dig
      dig = ZeroToTen(val mod 10)
      Controller.B2SSetData 240+i, dig
      val = int(val/10)
    else
      Controller.B2SSetData 240+i, 0
    end if
  next

  ' set the new initials
  for i = 1 to 3 : Controller.B2SSetData 236+i, 0 : next
  for i = 1 to len(inits)
    val = asc(mid(inits, i)) - 64
    Controller.B2SSetData 236+i, val
  next

  ' set the date
  for i = 1 to 8 : Controller.B2SSetData 228+i, 0 : next
  if d <> "" then
    d = CDate(d)
    dim mm, dd, yy : mm = Month(d) : dd = Day(d) : yy = Year(d)
    Controller.B2SSetData 229, mm \ 10
    Controller.B2SSetData 230, ZeroToTen(mm mod 10)
    Controller.B2SSetData 231, ZeroToTen(dd \ 10)
    Controller.B2SSetData 232, ZeroToTen(dd mod 10)
    Controller.B2SSetData 233, (yy \ 1000)
    Controller.B2SSetData 234, ZeroToTen((yy \ 100) mod 10)
    Controller.B2SSetData 235, ZeroToTen((yy \ 10) mod 10)
    Controller.B2SSetData 236, ZeroToTen(yy mod 10)
  end if
End If
End Sub

' Display the next high score above a given score.  This is used to show
' the next goal at the start of each ball.
Sub DispNextHighScoreAbove(s)
  ' start with the lowest high score and work up from there
  for i = ubound(Highscore) to 0 step -1
    if i = 0 or (Highscore(i) <> 0 And Highscore(i) > s) then
      DispHighScore i
      exit sub
    end if
  next

  ' If we didn't find a score to display, the given score must be
  ' higher than the highest score to date, so display the highest score.
  DispHighScore 0
End Sub

' A B2S lamp value of 0 always means "off", so when we want to light up
' a lamp to show a "0" digit, we have to use a non-zero value for the lamp.
' To deal with this, all of our backglass lamps that represent "0" digits
' have lamp value 10.  This routine takes care of converting a digit value
' on the VP side from 0 to 10 to match up with the B2S numbering.
Function ZeroToTen(n)
  if n = 0 then ZeroToTen = 10 else ZeroToTen = n
End Function

' High score display timer.  This advances the high score display to the
' next score in the list.  This only runs between games; during a game,
' we stop the rotation and show a single score.
Dim CurHighScoreDisp
Sub HighScoreDisplayTimer_Timer
  DispHighScore CurHighScoreDisp
  CurHighScoreDisp = CurHighScoreDisp + 1
  if CurHighScoreDisp > ubound(Highscore) then CurHighScoreDisp = 0
  if Highscore(CurHighScoreDisp) = 0 then CurHighScoreDisp = 0
End Sub


' Add one or more credits.  This adds n to the pending credit award counter,
' and starts the credits timer.  The timer will award the credits one at a time,
' spaced out so that each awarded credit gets a distinct knock effect.
Dim CreditsToAdd
Sub AddCredits(n)
  CreditsToAdd = CreditsToAdd + n
  CreditsTimer.Enabled = 1
End Sub

Sub CreditsTimer_Timer
  If CreditsToAdd > 0 then
    CreditsToAdd = CreditsToAdd - 1
    AddCredit
  else
    CreditsTimer.Enabled = 0
  End If
End Sub


' Generate a match number, display it on the backglass, and award credits
' to matching scores.
Sub Match()
  MatchNumber = Int(Rnd(1)*10)*10
    If B2SOn Then
    Controller.B2ssetmatch 34, Matchnumber
    End If
  If  MatchNumber = 0 Then
        If B2SOn Then
    Controller.B2ssetmatch 34, 100
        End If
      MatchReel.SetValue(0):ML.state = 1 ' Sdet the match number on the desktop match reel and turn on the reel light.
  Else
    If B2SOn Then
        Controller.B2ssetmatch 34, Matchnumber
        End If
     MatchReel.SetValue(Matchnumber):ML.State = 1 ' Sdet the match number on the desktop match reel and turn on the reel light.
  End If
  For i = 1 to Players
    If  MatchNumber = (ScorePlayer(i) Mod 100) Then
      Addcredit
    End If
  Next
End Sub


' ****************************************************************************************************************************
'
' High Score Initial entry
'
' We want to keep track of the high score to date, including the initials of the player who
' set the last high score.  More modern pinballs do this with their new-fangled alphanumeric
' displays, but Aladdin's Castle was from an older generation that just had mechanical
' score reels.  So to provide this feature, we have to invent something out of whole cloth.
'
' Back in the day, some people kept track of high score records with sticky notes stuck to
' the apron or backglass.  So to keep in character with the simulation, we'll take that as
' our model.  Our companion B2S backglass uses the third monitor (where the DMD would be
' displayed if we had one) to display a simulated sticky note with the high score to date.
' And when a new high score is set, we'll pop up another little sticky note on top of the
' playfield, showing a UI where the player can enter their initials.  The actual process of
' entering initials is exactly like on a modern machine, using the flipper buttons to cycle
' through the alphabet, so even though it's anachronistic for this particular machine, it
' will at least be familiar in a pinball context in general.

' In terms of the concrete implementation, we also have to improvise a bit, since VP
' doesn't have any direct way to solicit input like this.  We use an EMReel for the sticky
' note graphics, and an additional reel for each character position.  The reels are useful
' for this because they're drawn in the foreground and we can set arbitrary images.
' Furthermore, a reel can show different "digits", which are actually just image cells.
' For the sticky note background, we use two cells for the two different prompts
' ("Player One" and "Player Two"), and for each initial slot we use an alphabet reel.
'
' The initials are represented as we enter them by values 0-25 for A-Z, 26 for blank, and
' 27 for backspace.  We don't allow blanks for backspaces in the first position.  (It
' would be easy to extend this idea to a larger symbol set if desired, but it seems
' better in terms of the UI experience to keep the symbol set small so there aren't
' too many letters to cycle through.)
'
' Note that there are four positions for initials, even though we only take the standard
' three initials.  The fourth position doesn't count as an initial, but is just there
' for the final "commit" step.  At this position, the player can only select between a
' special "Done" symbol and the backspace.  Selecting the backspace backs up and allows
' changing the previous character, and selecting the "Done" symbol commits the entry and
' ends the entry process.  This matches the UI used on most real pinballs, although some
' machines have a simpler variation where entering the third initial commits the whole
' entry without an extra chance to make corrections.  For anyone re-using this code in
' another table, you can change to the simpler UI by setting the SimplerInitUI to true.
Dim EnteringHS, EnteringHSSlot, HSPos, HSCurInit, HSInitStr
Dim HSInits
If ShowDT = False Then
HSInits = Array(HSInit1, HSInit2, HSInit3, HSInit4)
Else
HSInits = Array(HSInit5, HSInit6, HSInit7, HSInit8)
End If
Sub BeginHighScoreInits(player, slot)
  ' show the UI for entering initials
    If ShowDT = False Then
  HighScoreDlg.SetValue player-1
  HighScoreDlg.Image = "High Score Prompt"
  HSInit1.Image = "Alphabet Reel"
    Else
    HighScoreDlgDT.SetValue player-1
  HighScoreDlgDT.Image = "High Score Prompt_DT"
  HSInit5.Image = "Alphabet Reel DT"
    End If
  EnteringHS = true
  EnteringHSSlot = slot

  ' start at the first character at "A"
  HSInit1.SetValue 0
  HSInitStr = ""
  HSPos = 0
  HSCurInit = 0

  ' flash the cursor
  HSBlink = 0
  InitEntryTimerTick = 0
  HSLeftFlipperDown = 0
  HSRightFlipperDown = 0
  InitEntryTimer.Enabled = true
End Sub

Sub EndHighScoreInits
  ' set the new initials
  HighscoreInits(EnteringHSSlot) = HSInitStr

  ' debugging - display the final string
  ' if HSInitStr <> "" then msgbox "[" & HSInitStr & "]"

  ' done entering initials - hide the UI
  HideHighScoreInitUI

  ' check the next player's score, if there are more players to be checked
  CheckHighScores
End Sub

Sub HideHighScoreInitUI
    If ShowDT = False Then
  HighScoreDlg.Image = "Transparent Backdrop"
    Else
    HighScoreDlgDT.Image = "Transparent Backdrop"
    End If
  for each Obj in HSInits : Obj.SetValue 26 : Next
  InitEntryTimerTick = 0
  InitEntryTimer.Enabled = false
  EnteringHS = false
    'Populate the high score table in desktop mode with updated scores when the HS entry field is closed.
    hs1.text = HighscoreInits(0)
  hs1_val.text = Highscore(0)
  HS2.text = HighscoreInits(1)
  hs2_val.text = Highscore(1)
    hs3.text = HighscoreInits(2)
  hs3_val.text = Highscore(2)
  HS4.text = HighscoreInits(3)
  hs4_val.text = Highscore(3)
End Sub

' Blink the current initial, and handle auto-repeat keys.
Dim InitEntryTimerTick, HSBlink
Sub InitEntryTimer_Timer
  ' Blink the cursor every 500ms.  The alphabet reel has 60 slots, 30 for
  ' normal and 30 for highlighted.  So simply switch back and forth
  ' between N and N+30.
  InitEntryTimerTick = (InitEntryTimerTick + 1) mod 10
  if InitEntryTimerTick = 0 then
    HSBlink = (HSBlink + 30) mod 60
    HSInits(HSPos).SetValue HSCurInit + HSBlink
  end if

  ' If a flipper button is being held down, auto-repeat on every tick
  ' after the first 500ms
  if HSLeftFlipperDown = 10 then
    HSLeftFlipperKey
  elseif HSLeftFlipperDown > 0 then
    HSLeftFlipperDown = HSLeftFlipperDown + 1
  end if

  if HSRightFlipperDown = 10 then
    HSRightFlipperKey
  elseif HSRightFlipperDown > 0 then
    HSRightFlipperDown = HSRightFlipperDown + 1
  end if
End Sub

' High Score Entry key handler.  The normal keydown handler calls this when
' initial entry is in progress, as indicated by 'EnteringHS'.  We just need
' to handle the flipper buttons and the Start button to implement the typical
' UI for initial entry that virtually all alphanumeric and DMD machines use.
Dim HSLeftFlipperDown, HSRightFlipperDown
Sub HighScoreInitKey(keycode)
  If Keycode = LeftFlipperKey Then
    HSLeftFlipperKey
    HSLeftFlipperDown = 1
  Elseif Keycode = RightFlipperKey Then
    HSRightFlipperKey
    HSRightFlipperDown = 1
  Elseif Keycode = StartGameKey Then
    if HSCurInit = 27 and HSPos > 0 then
      ' backspace symbol - delete current position and back up to previous
      HSInits(HSPos).SetValue 26   ' set value to space to hide initial
      HSPos = HSPos - 1
      HSBlink = 0
      HSCurInit = Asc(Right(HSInitStr, 1))-65
      HSInitStr = Left(HSInitStr, Len(HSInitStr)-1)
    else
      ' any other symbol - move to next position, or commit entry at position 4
      HSInits(HSPos).SetValue HSCurInit   ' set base value in case we were blinked on
      HSPos = HSPos + 1
      if HSPos <= 3 then HSInitStr = HSInitStr & chr(65+HSCurInit)  ' add the initial to the string
      if HSPos = 4 or (SimplerInitUI and HSPos = 3) then
        ' fourth position or third position in simpler UI - commit the entry
        EndHighScoreInits
      else
        if HSPos = 3 then HSCurInit = 28   ' start at the same letter for a new slot, except start the fourth at "OK"
        If ShowDT = False Then
                HSInits(HSPos).Image = "Alphabet Reel"
                Else
                HSInits(HSPos).Image = "Alphabet Reel DT"
                End If
        HSBlink = 0
        HSInits(HSPos).SetValue HSCurInit
      end if
    end if
  End If
End Sub

Sub HighScoreInitKeyUp(keycode)
  if keycode = LeftFlipperKey then
    HSLeftFlipperDown = 0
  elseif keycode = RightFlipperKey then
    HSRightFlipperDown = 0
  end if
End Sub

Sub HSLeftFlipperKey
  HSCurInit = HSCurInit - 1
  if HSCurInit < 0 and HSPos = 0 then
    HSCurInit = 25       ' first position - wrap A -> Z
  elseif HSCurInit < 0 and HSPos <> 3 then
    HSCurInit = 27     ' 2nd/3rd position - wrap A -> backspace
  elseif HSPos = 3 and HSCurInit < 27 then
    HSCurInit = 28    ' 4th position - wrap backspace -> End
  end if
  HSInits(HSPos).SetValue HSCurInit
  HSBlink = 0
End Sub

Sub HSRightFlipperKey
  HSCurInit = HSCurInit + 1
  if HSCurInit > 25 and HSPos = 0 then
    HSCurInit = 0       ' first position - wrap Z -> A
  elseif HSPos <> 3 and HSCurInit > 27 then
    HSCurInit = 0       ' 2nd-3rd position - wrap Backspace -> A
  elseif HSPos = 3 and HSCurInit > 28 then
    HSCurInit = 27        ' 4th position - wrap End -> Backspace
  end if
  HSInits(HSPos).SetValue HSCurInit
End Sub

HideHighScoreInitUI

'****************************************************************************************************************************
'
' Operator Menu
'
' Bring up the operator menu by opening the coin door (END key).  This presents a UI for setting
' the various game options described in the original Bally instruction manual.

'
' Menu item map
' Each map element is an array that describes the menu line.  The first element of that array
' is the item type, which specifies the meaning of the rest of the items:
'
'   "radio" -> radio button:
'        Variable Name, Reel Object, First Reel Index, Last Reel Index
'   "digit" -> digit value:
'    Variable Name, Reel Object, Min Value, Max Value
'   "score" -> integer score value:
'        Variable Name, Ten-Thousands Reel, Thousands Reel, 000 Reel, 000 image index
'   "save", "cancel", "defaults" -> action button
'        Reel Object, Reel Image Index
'

Dim OpMenuMap : OpMenuMap = Array( _
  Array("radio", "OptFreePlay", ReelFreePlay, 0, 1), _
  Array("digit", "OptBallsPerGame", ReelBalls, 1, 5), _
  Array("radio", "OptAlleyScoring", ReelAlley, 2, 3), _
  Array("radio", "OptSpecialVal", ReelSpecial, 4, 5), _
  Array("radio", "OptReplayVal", ReelReplay, 4, 5), _
  Array("score", "OptReplay1", ReelReplay1TT, ReelReplay1T, ReelReplay1ZZZ, 6), _
  Array("score", "OptReplay2", ReelReplay2TT, ReelReplay2T, ReelReplay2ZZZ, 6), _
  Array("radio", "OptHSReset", ReelHSReset, 9, 10), _
  Array("defaults", ReelFactoryReset, 11), _
  Array("save", ReelOptSave, 7), _
  Array("cancel", ReelOptCancel, 8) _
)
Dim OpMenuTmp(10)   ' temporary working values for menu items while menu is open
Dim OpMenuActive


' flag variable used in menu only: on saving menu, reset high scores
Dim OptHSReset

Sub ShowOperatorMenu
  OpMenuActive = true
  OpSavePending = false
  OpMenuTimer.Enabled = true
  OpBackdrop.Image = "Options Backdrop"
  OpMenuStartBtnDown = 0
  OptHSReset = false

  dim i, ele, val
  for i = 0 to ubound(OpMenuMap)
    ele = OpMenuMap(i)
    select case ele(0)
      case "radio"
        val = Eval(ele(1))
        OpMenuTmp(i) = val
        ele(2).SetValue (ele(3) + val)*2
        ele(2).Image = "Options Reel"

      case "digit"
        val = Eval(ele(1))
        OpMenuTmp(i) = val
        ele(2).SetValue val*2
        ele(2).Image = "Options Number Reel"

      case "score"
        val = Eval(ele(1))
        OpMenuSetScore i, val, false
        ele(2).Image = "Options Number Reel"
        ele(3).Image = "Options Number Reel"
        ele(4).Image = "Options Reel"

      case "save", "cancel", "defaults"
        ele(1).SetValue ele(2)*2
        ele(1).Image = "Options Reel"

    end select
  next

  ' start on the last button (the Cancel button) for easy dismissal
  OpMenuLine = ubound(OpMenuMap)
  OpMenuHilite true
  OpMenuBlink = 0
End Sub

Sub OpMenuSave
  ' go through the item list, and save each data item to its global variable
  dim i, ele, varname, val
  for i = 0 to ubound(OpMenuMap)
    ele = OpMenuMap(i)
    select case ele(0)
      case "radio", "digit", "score"
        varname = ele(1)
        ExecuteGlobal varname & "=" & OpMenuTmp(i)
    end select
  next

  ' if a high score reset was requested, apply it
  if OptHSReset then HighScoreReset

  ' update the persistent store with the new settings
  SaveInfo
End Sub

Sub OpMenuSetScore(line, val, hilite)
  dim ele : ele = OpMenuMap(line)
  OpMenuTmp(line) = val
  dim ofs : if hilite then ofs = 1 else ofs = 0
  ele(2).SetValue ((val \ 10000) mod 10)*2 + ofs
  ele(3).SetValue ((val \ 1000) mod 10)*2 + ofs
  ele(4).SetValue ele(5)*2 + ofs
End Sub

Dim OpMenuLine
Sub OpMenuSelect(line)
  OpMenuHilite false
  dim n : n = ubound(OpMenuMap)
  if line > n then line = 0 else if line < 0 then line = n
  OpMenuLine = line
  OpMenuHilite true
End Sub

Sub OpMenuHilite(hilite)
  dim ele : ele = OpMenuMap(OpMenuLine)
  dim val : val = OpMenuTmp(OpMenuLine)
  dim ofs : if hilite then ofs = 1 else ofs = 0
  select case ele(0)
    case "radio":
      ele(2).SetValue (ele(3) + val)*2 + ofs
    case "digit"
      ele(2).SetValue OpMenuTmp(OpMenuLine)*2 + ofs
    case "score"
      OpMenuSetScore OpMenuLine, val, hilite
    case "save", "cancel", "defaults"
      ele(1).SetValue ele(2)*2 + ofs
  end select
End Sub

Dim OpMenuBlink, OpMenuStartBtnDown, OpSavePending
Sub OpMenuTimer_Timer
  ' if a save is pending, do the save now and stop the timer
  if OpSavePending then
    OpMenuSave
    OpMenuShutdown
    exit sub
  end if

  ' auto-repeat the Start button, if held down more than 500ms
  if OpMenuStartBtnDown >= 10 then
    OpMenuStartBtn
  elseif OpMenuStartBtnDown > 0 then
    OpMenuStartBtnDown = OpMenuStartBtnDown + 1
  end if

  ' blink every 500 ms, after a 750 ms initial delay
  OpMenuBlink = OpMenuBlink + 1
  if OpMenuBlink = 15 then
    OpMenuHilite false
  elseif OpMenuBlink = 25 then
    OpMenuHilite true
    OpMenuBlink = 5
  end if
End Sub

Sub OpMenuKey(keycode)
  dim i
  if keycode = LeftFlipperKey then
    OpMenuMove(-1)
  elseif keycode = RightFlipperKey then
    OpMenuMove(1)
  elseif keycode = StartGameKey then
    OpMenuStartBtnDown = 1
    OpMenuStartBtn
  end if
End Sub

Sub OpMenuStartBtn
  dim ele : ele = OpMenuMap(OpMenuLine)
  dim val : val = OpMenuTmp(OpMenuLine)
  OpMenuBlink = false
  select case ele(0)
    case "radio"
      ' radio button - advance to next selection, wrap at last selection
      val = val + 1
      if val + ele(3) > ele(4) then val = 0
      ele(2).SetValue (ele(3) + val)*2 + 1

    case "digit"
      ' digit - advance to next value, wrap if we exceed the maximum
      val = val + 1
      if val > ele(4) then val = ele(3)
      ele(2).SetValue val*2 + 1

    case "score"
      ' score line - add 1000, wrap to 0 at 99000, and jump from 0 to 50000
      if val = 99000 then val = 0 else if val = 0 then val = 50000 else val = val + 1000
      OpMenuSetScore OpMenuLine, val, true

    case "save"
      HideOperatorMenu true

    case "cancel"
      HideOperatorMenu false

    case "defaults"
      ' run through the menu items and set each one to the corresponding
      ' default value from the saved variable list
      for i = 0 to ubound(OpMenuMap)
        dim o : o = OpMenuMap(i)
        dim t : t = o(0)    ' type of this entry
        if t = "radio" or t = "digit" or t = "score" then
          ' these types are associated with option variable values - apply
          ' the default value for the variable
          val = GetSavedVarDefault(o(1))

          ' special case: set High Score Reset to ON
          if o(1) = "OptHSReset" then val = 1

          ' update the displayed item
          OpMenuTmp(i) = val
          select case t
            case "radio"
              o(2).SetValue (o(3) + val)*2

            case "digit"
              o(2).SetValue val*2

            case "score"
              OpMenuSetScore i, val, false
          end select
        end if
      next

  end select
  OpMenuTmp(OpMenuLine) = val
End Sub

Sub OpMenuKeyUp(keycode)
  if keycode = StartGameKey then
    OpMenuStartBtnDown = false
  end if
End Sub

Sub OpMenuMove(dir)
  OpMenuHilite false
  OpMenuLine = OpMenuLine + dir
  dim n : n = ubound(OpMenuMap)
  if OpMenuLine < 0 then
    OpMenuLine = n
  elseif OpMenuLine > n then
    OpMenuLine = 0
  end if
  OpMenuHilite true
  OpMenuBlink = false
End Sub

Sub HideOperatorMenu(save)
  ' hide the UI elements
  for each Obj in OpMenu : Obj.Image = "Transparent Backdrop" : Next

  ' If they want to save the new settings, flag a pending save for the timer.
  ' We do the save in thetimer rather than right now so that the UI disappears
  ' immediately, which provides a smoother feedback experience.  The save takes
  ' a couple of seconds, so if we don't remove the UI first, the button push
  ' feels laggy.  If they don't want to save the new settings, simply shut
  ' down the UI immediately.
  if save then
    ' flag the pending save for the timer
    OpSavePending = true
  else
    ' no save - shut down the UI immediately
    OpMenuShutdown
  end if
End Sub

' Shut down
Sub OpMenuShutdown
  OpMenuTimer.Enabled = false
  OpMenuActive = false
End Sub

HideOperatorMenu false

'*****************************************************
'TEST MENU
'*****************************************************


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************
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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
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

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, vol)
    PlaySound soundname, 1, (vol), AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingSoundUpdate()
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
    Next
End Sub


'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*****************************************
' Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)


Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls

  ' render the shadow for each ball
    For b = 0 to UBound(BOT)
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
    End If

      BallShadow(b).Y = BOT(b).Y + 10
      BallShadow(b).Z = 1
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

Sub Update_Stuff_Timer()
'Most of this stuff could be put in the corresponding routines but I'm lazy so I monitor it in realtime here.
CreditReel.setvalue(Credits)
If Player = 1 Then
P1_Light.State = 1:P2_Light.State = 0:P1_LightA.State = 1:P2_LightA.State = 0
End If
If Player = 2 Then
P2_Light.State = 1:P1_Light.State = 0:P1_LightA.State = 0:P2_LightA.State = 1
End If
If credits > 0 Then
Light6.State = 1
Else
Light6.State = 0
End If
'End of lazy bit.
RollingSoundUpdate
BallShadowUpdate
FlipperLSh.RotZ = LeftFlipper.currentangle
FlipperRSh.RotZ = RightFlipper1.currentangle
FlipperRSh1.RotZ = RightFlipper.currentangle
End Sub
 '********************* Game Over Man!! **************************************************************************************
Sub Table1_Exit
If B2SOn Then
Controller.Stop
End If
End Sub

Dim hld, rld, pld

Sub Hand_Left_Timer()

If hld => 1 AND hld <=6 Then
LFB.transz = LFB.transz + 1.5
End If

If hld => 6 AND hld <=11 Then
LFB.transz = LFB.transz - 1.5
End If

Select Case hld
Case 1:
HLR.visible = 0:HLR1.visible = 1:hld = 2
Case 2:
HLR1.visible = 0:HLR2.visible = 1:hld = 3
Case 3:
HLR2.visible = 0:HLR3.visible = 1:hld = 4
Case 4:
HLR3.visible = 0:HLR4.visible = 1:hld = 5
Case 5:
HLR4.visible = 0:HLR5.visible = 1:LFB.transz = 7.5 : me.enabled = 0
Case 6:
HLR5.visible = 0:HLR4.visible = 1:hld = 7
Case 7:
HLR5.visible = 0:HLR4.visible = 1:hld = 8
Case 8:
HLR4.visible = 0:HLR3.visible = 1:hld = 9
Case 9:
HLR3.visible = 0:HLR2.visible = 1:hld = 10
Case 10:
HLR2.visible = 0:HLR1.visible = 1:hld = 11
Case 11:
HLR1.visible = 0:HLR.visible = 1:LFB.transz = 0 : me.enabled = 0
End Select
End Sub

Sub Hand_Right_Timer()

If rld => 1 AND rld <=6 Then
RFB.transz = RFB.transz - 1.5
End If

If rld => 6 AND rld <=11 Then
RFB.transz = RFB.transz + 1.5
End If

Select Case rld
Case 1:
HRR.visible = 0:HRR1.visible = 1:rld = 2
Case 2:
HRR1.visible = 0:HRR2.visible = 1:rld = 3
Case 3:
HRR2.visible = 0:HRR3.visible = 1:rld = 4
Case 4:
HRR3.visible = 0:HRR4.visible = 1:rld = 5
Case 5:
HRR4.visible = 0:HRR5.visible = 1:RFB.transz = -7.5 : me.enabled = 0
Case 6:
HRR5.visible = 0:HRR4.visible = 1:rld = 7
Case 7:
HRR5.visible = 0:HRR4.visible = 1:rld = 8
Case 8:
HRR4.visible = 0:HRR3.visible = 1:rld = 9
Case 9:
HRR3.visible = 0:HRR2.visible = 1:rld = 10
Case 10:
HRR2.visible = 0:HRR1.visible = 1:rld = 11
Case 11:
HRR1.visible = 0:HRR.visible = 1:RFB.transz = 0 : me.enabled = 0
End Select
End Sub

Sub Plunger_Hand_Timer()
Select case pld
Case 1:HRR.visible = 0:HRRP1.visible = 1:pld = 2
Case 2:HRRP1.visible = 0:HRRP2.visible = 1:pld = 3
Case 3:HRRP2.visible = 0:HRRP3.visible = 1:pld = 4
Case 4:HRRP3.visible = 0:HRRP4.visible = 1:pld = 5
Case 5:HRRP4.visible = 0:HRRP5.visible = 1:pld = 6
Case 6:HRRP5.visible = 0:HRRP6.visible = 1:pld = 7
Case 7:HRRP6.visible = 0:HRRP7.visible = 1:pld = 8
Case 8:HRRP7.visible = 0:HRRP8.visible = 1:pld = 9
Case 9:HRRP8.visible = 0:HRRP9.visible = 1:pld = 10
Case 10:HRRP9.visible = 0:HRRP10.visible = 1:me.enabled = 0

Case 11:HRRP10.visible = 0:HRRP9.visible = 1:pld = 12
Case 12:HRRP9.visible = 0:HRRP8.visible = 1:pld = 13
Case 13:HRRP8.visible = 0:HRRP7.visible = 1:pld = 14
Case 14:HRRP7.visible = 0:HRRP6.visible = 1:pld = 15
Case 15:HRRP6.visible = 0:HRRP5.visible = 1:pld = 16
Case 16:HRRP5.visible = 0:HRRP4.visible = 1:pld = 17
Case 17:HRRP4.visible = 0:HRRP3.visible = 1:pld = 18
Case 18:HRRP3.visible = 0:HRRP2.visible = 1:pld = 19
Case 19:HRRP2.visible = 0:HRRP1.visible = 1:pld = 20
Case 20:HRRP1.visible = 0:HRR.visible = 1:me.enabled = 0
End Select
End Sub


Sub Smoke_Timer()
Cig_Smoke.height = Cig_Smoke.height + 0.1
Cig_Smoke1.height = Cig_Smoke1.height + 0.15

If Cig_Smoke.height <= 100 Then
Cig_smoke.opacity = Cig_smoke.opacity + 10
End If

If Cig_Smoke.height => 100 AND Cig_Smoke.height <= 149 Then
Cig_smoke.opacity = Cig_smoke.opacity - 10
End If

If Cig_Smoke.height => 150 Then
Cig_smoke.opacity = 0:Cig_Smoke.height = 80
End If

If Cig_Smoke1.height <= 90 Then
Cig_Smoke1.opacity = Cig_Smoke1.opacity + 5
End If

If Cig_Smoke1.height => 90 AND Cig_Smoke1.height <= 99 Then
Cig_smoke1.opacity = Cig_smoke1.opacity - 5
End If

If Cig_Smoke1.height => 100 Then
Cig_Smoke1.opacity = 0:Cig_Smoke1.height = 80
End If

End Sub

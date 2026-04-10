' ********  Green Pastures   *************
'           Gottlieb 1954
'
'  Version 1.0 PLB May 2020
'
'  Thanks to too many to name, as I have learned from
'  those who went before.  Anyone may feel free to steal any
'  Ideas, Code, Sounds or Images from this game at will and
'  without notice.
'
' My Goal was to try for as realistic game play as I could manage
' especially sounds and timing.  Images for playfield and Backglass
' are all originally taken from Real Games photographed on IPDB and
' other internet sites manipulated in Photoshop to mesh together with
' the dirt Cleaned off in some places (Think of it as photoshop Novus 2).
'
'
'       A Note on the Internal Handling of scoring:
'       -------------------------------------------
'
'  This game's display adds four zeros to the standard 1 point scoring,
'  but only on the display.  So internally the following is used:
'
'       10,000 points =   1 point internally
'       50,000 points =   5 points internally
'      100,000 points =  10 points internally
'      500,000 points =  50 points internally
'    1,000,000 points = 100 points internally
'
'  These are the only five point scoring values possible, so the
'  extra zeros are not carried in the code.
'
'  As an example. 5,510,000 points is internally carried as 551 points.
'  The display adds the extra zeros which never change.
'
'
'
Option Explicit 'Force variables to be declared

Dim TableName     'For defining the table name to be used when saving credits/highscore
Dim InProgress      'TRUE while game is in progress, FALSE otherwise
Dim Tilted        'TRUE when tilted, FALSE otherwise
Dim TiltSensitivity   'Obvious
Dim State       'Used for playfield objects that can be disabled during TILT
Dim BallsLost     'Tracks the balls played
Dim BallsLeft     'Tracks the balls still to be lifted
Dim BallsLoaded         'Startup Trough Loading Counter
Dim BallAtPlunger       'True when a Ball is at the plunger, false when not
Dim BallsOnField        'Count of Balls Lifted and not yet drained
Dim BIT         'Balls in Virtual Lifter Tray When draining trough
Dim FirstOut            'Set when Ball First Exits Gate
Dim RollFlag            'Ball Roll Sound Semiphore
Dim RollVary            'Flag to Vary Roll Sound
Dim RollBack            'Flag to Detect Failure to leave gate on Plunger Release
Dim CreditsPerCoin    '1 to 5
Dim Score       'The score itself
Dim ScoreToAdd      'Used in the AddScore sub
Dim SReelUnits      'Scoring Reel Values (reels can't be read back)
Dim SReelTens
Dim SReelHuns
Dim MUnits        '"Scoring Motor" Units to Pules
Dim MTens       '"Scoring Motor" Tens to Pulse
Dim HighScore     'Holds previous high score value
Dim HighScorePaid   'Set to TRUE after beating high score
Dim HighScoreReward   'Set to the number of replays awarded when HIGH SCORE is beaten
Dim Credits       'Keeps track of the credits
Dim MaxCredits          'Maximum Credits Allowed
'
Dim GSPhase       'Cycle Counter Game Start reset sequence
Dim DelayedStart        'True if start pressed during Game Over Sequence
Dim GOVPhase      'Cycle Counter for Game Over Sequence
Dim DelayedGOV      'True Druing Interval From Tilt to Game Over Seq Started
Dim DelayedReplays    'Count Of Delayed Replays to Add with Knocker
Dim Count       'Cycle Counter for Credit_Timer
Dim Count1        'Used in TiltTimer for tilt sensitivity
Dim X         'Local Generic variable used in for-next loops
Dim Y         'Local Generic variable used in for-next loops
Dim Z         'Local Generic variable used in for-next loops
'
Dim BuzzCount     'Tracks Flipper Play/Stop Buzz sounds to allow corrections
'
Dim Replay1       'The scores needed for a replay are held in these 3 variables
Dim Replay2
Dim Replay3
Dim Replay4
Dim Replay5
Dim Replay1Paid     'Mark when each replay has been paid
Dim Replay2Paid
Dim Replay3Paid
Dim Replay4Paid
Dim Replay5Paid
'
Dim Ones        'These are used in the score motor simulator
Dim Tens
Dim Hundreds
Dim Thousands
'
Dim Value
Dim DisplayHighScore  'Don't display high score if there is no saved value
Dim ReelSpeed
Dim Bells
Dim Msg(21)       ' Used to display the rules
'
Dim M5Q         ' Queue 5 point motor runs
Dim M50Q        ' Queue 50 point motor runs
Dim MotorPhase      ' Scoring Motor Cam Phase
'
'
Dim B2State             ' B2 Pop Bumper Animation State
Dim B7State             ' B7 Pop Bumper Animation State
Dim B9State             ' B9 Pop Bumper Animation State
Dim PW          ' Plunger Animation State
Dim ShootBall           ' Is animation Pulling, or Releasing?
Dim PS                  ' Pull Speed
Dim DrainWallHit    ' Limits to one it sound per ball on Drain Wall
Dim TrayHit       ' Limits to one sound per ball on Tray Hit
Dim BallID        ' Ball ID number
Dim BallDrained(25)   ' List of IDs of drained balls (index 0 = # entries in list)
Dim TroughOpening   ' True = Opening, False = Closing
'
Dim aBall(3)      ' Hole Ball Pointers
Dim aXpos(3)
Dim aXadd(3)
Dim aYpos(3)
Dim aYadd(3)
Dim aZpos(3)      ' Hole Drop Vertical Position
Dim aTicks(3)     ' Hole Attract Timer Tick Counter
'
Dim LRLast              ' Units Digit after Previous AddScore()
'
Dim MadeSpecial     ' A-B-C-D Special Has Been Made (True or False)
Dim Points        ' Total Points Made
Dim Spotted(13)     ' Number 1-12 spotted array If Spotted(12)=1 Sequence has been made
 Dim Rotate       ' 0-3 for which letter is special if MadeSpecial=True
'
Dim Debug
'
 Dim Controller
 LoadController
 Sub LoadController()
 Set Controller = CreateObject("B2S.Server")
 Controller.B2SName = "Green Pastures"  ' *****!!!!This must match the name of your directb2s file!!!!
 Controller.Run()
 End Sub

' ---------------------------------------
' *** Game Init -- Script Starts here ***
' ---------------------------------------
'
Sub Table_Init()
'
  FirstOut=0
  BallAtPlunger=False
  DelayedStart=False   'No Delayed Anything yet
  DelayedGOV=False
  DelayedReplays=0
'
  TableName="GreenPastures" 'Place your table name here
  CreditsPerCoin=5  '1 to 5 credits per coin
  MaxCredits = 29     'Size of Credit Reel Graphic
  HighScoreReward=1 'The number of replays to award when HIGH SCORE is beaten (1-5)
'
  TiltTimer.Enabled=True
  TiltSensitivity=4 '0-15, 4=normal - a higher number is less sensitive, 0=1 nudge and you're out
  Bells=True      'TRUE to play bell sounds, FALSE for no bells
  ReelSpeed=3     'Defines the speed of the reels, 1=slow, 9=fast, 3=default
'
  Replay1=460     'Place the score needed for replays in these variables
  Replay2=550
  Replay3=570
  Replay4=600
  Replay5=650
'
  Initialize()    'Call the main Init sub
End Sub
'
'
'**************************************
'****  Handle Contact Switch Hits *****
'**************************************
'
' Idx   Switch
 ' ----  ---------------
 '  0    LSW1 - 100,000
 '  1    RSW1 - 100,000
 '  2    LSW2 - 10,000
 '  3    RSW2 - 10,000
 '  4    LSW3 - 10,000
 '  5    RSW3 - 10,000
 '  6    LSW4 - 10,000
 '  7    RSW4 - 10,000
'
Sub StandUps_Hit(Idx)
  If StandUps(Idx).UserValue=True Then
     Exit Sub             ' Debounce rolling Ball
  End If
  StandUps(Idx).UserValue=True    ' This is hit, set debounce interval
  StandUps(Idx).TimerEnabled=True   ' Start Debounce timer
  If (Idx < 2) Then
     AddScore 10, 0         ' Hit 100,000 Contact
     Else
        AddScore 1, 0         ' hit 10,000 Contact
  End If
    AddSound "RollMore"
End Sub
'
Sub StandUps_Timer(Idx)
  StandUps(Idx).UserValue=False   ' End Debounce Window
  StandUps(Idx).TimerEnabled=False
End Sub
'
'----------------------------------------
'  Selected Ball Hits - Sound Only
'----------------------------------------
'
Sub Rubbers_Hit(Idx)
  Dim I, J
    If Rubbers(0).UserValue <> 0 Then
     Exit Sub           ' debouncing
  End If
  Rubbers(0).UserValue=1
  Rubbers(0).TimerEnabled=True
  I=ActiveBall.VelY
  J=ActiveBall.VelX
  If (SQR((I*I)+(J*J)) > 3) Then
     AddSound "RubberRoll"
  Else
     AddSound "Rubber"
  End If
End Sub
'
Sub Rubbers_Timer(Idx)
  Rubbers(0).UserValue=0      ' end debouncing
  Rubbers(Idx).TimerEnabled=False
End Sub
 '
Sub Divider_Hit()
  AddSound "Wood"
End Sub
'
Sub LeftRail_Hit()
  AddSound "Wood"
End Sub
'
Sub LeftBallWall_Hit()
  AddSound "Click"
End Sub
'
Sub RtBallWall_Hit()
  AddSound "Click"
End Sub
'
Sub BallRetain_Hit()
  If TrayHit=False Then
     ActiveBall.VelY=ActiveBall.VelY/4    ' Slow it's vertical speed to 1/4
     ActiveBall.VelX=ActiveBall.VelX*.667   ' Slow it's Horizontal speed by 1/3
     AddSound "Wood"              ' Hit back of Return Trough
     TrayHit=True               ' One Hit for this ball
  End If
End Sub
'
Sub DrainWall_Hit()
  If DrainWallHit=True Then
     AddSound "Roll9"       ' Roll After first hit
  Else
     AddSound "Wood"        ' First Drain Wall Hit
     DrainWallHit=True      ' Say Roll After this
  End If
End Sub
'
'  -------------------------------------
'  Routine to Set High Score in HS Reels
'  -------------------------------------
'
Sub SetHS(HScore)
  X=HScore MOD 10
  HSUnits.SetValue X
  X = INT((HScore MOD 100) /10)
  HSTens.SetValue X
  X = INT((HScore MOD 1000) / 100)
  HSHundreds.SetValue X
End Sub
'
'  ---------------------------------
'  Enable/Disable Bumpers and Lights
'  ---------------------------------
'
Sub Disable(State)            'Disable or enable objects when TILTed
  ' Disable/Enable Bumpers
  For X=1 to 10
   '    EVAL("B"&X&"Top").Disabled=State
   '    EVAL("B"&X&"Base").Disabled=State
  Next
  '
  ' Kill Flippers
  If State=True Then
     DisableFlippers()
  End If
End Sub
'
'  Sub to Disable Flippers
'
Sub DisableFlippers()
  LeftFlipper.RotateToStart   'Move flippers to start position
  RightFlipper.RotateToStart

  While BuzzCount > 0
     StopSound "Buzz"
     BuzzCount=BuzzCount - 1
  Wend

End Sub
'
'
'  **********************************
'  ***** Main program Init code *****
'  **********************************
'
Sub Initialize()
  Randomize             'Seed random Number Generator
  Debug=0
'
  Count=0               'Set variables to a known state
  Count1=0
  GOVPhase=0
  GSPhase=0

  BallsLost=0             'Reset ball counter
  BallsLeft=0             'The number of balls not yet lifted
  BallsOnField=0            'No Balls Lifted to Playfield
  InProgress=False          'Game not started
    '
  DisplayHighScore=True       'Display the high score
    '
    DisplayGOV(1)           ' Turn On "Game Over"
  '
  '  Init Ball Return Trough
  '
  TroughOpen.IsDropped=True     ' Trough inits to Closed
  TroughOpenHalf.IsDropped=True   '  ... Fully Closed
  TroughOpening=False         ' Say Closing was last action
  DrainTrig.UserValue=0       ' Init debouncer
  GrowlTimer.UserValue=False      ' Allow Another AddSound
    '
    LtGI1.IsDropped=True        ' GI Lights Off
    LtGI2.IsDropped=True
    '
    ' Init targets to Not Hit & Unlit
    '
    For X=0 to (TargetsHit.Count-1)
       TargetsHit(X).IsDropped=True
        LtTarg(X).IsDropped=True
    Next
  '
  '  Set Initial Light States:
  '
    For X=0 to (LtNum.Count-1)
       LtNum(X).IsDropped=True      ' Number Lights Off
    Next
  '
    For X=0 to (LtLet.Count-1)
       LtLet(X).IsDropped=True      ' Letter Lights Off
    Next
     Rotate=INT(RND()*4)          ' Init Letter Rotate to 0-3
     '
  LtSpecial.IsDropped=True      ' Special Light Off
  '
  '   All 10 Bumpers init to Off
  '
  '  Note: Bumpers 2, 7, and 9 are Active (Pop Bumpers)
    '        All Other bumpers are passive bumpers
  '
  For X=1 to 10
       BumperOff(X)
     If X=2 OR X=7 OR X = 9 Then
      ' Drop animation on POP Bumper
      Eval("B"&X&"Ring1").IsDropped=True
      Eval("B"&X&"Ring2").IsDropped=True
      Eval("B"&X&"Ring3").IsDropped=True
      Eval("B"&X&"Ring3").TimerInterval=10  ' animation Speed
     Else
        ' Set Not Debouncing on Passive Bumper
      Eval("B"&X&"Top").UserValue=False
     End If
  Next
  '
     For X=0 to (StandUps.Count-1)
       StandUps(X).UserValue=False    ' Init Contact Switch Debouncers'
       StandUps(X).TimerInterval=250
     Next
     LSW1Hit.TimerInterval=500      ' Top Two need longer debounce times
     RSW1Hit.TimerInterval=500      ' for slow rolling ball
    '
  Rubbers(0).TimerInterval=250    ' Init Rubber Hit Sound Debouncer
  Rubbers(0).UserValue=0
    '
    For X=0 to (KickersLit.Count-1)
       KickersLit(X).IsDropped=True   ' Init Kickers UnLit
    Next
    '
  '
  '  Buttons Not Lit
  '
  For X=0 to (Buttons.Count-1)
     ButtonsDn(X).IsDropped=True    ' Lower unlit Button Down Image
     ButtonsDnLit(X).IsDropped=True ' Lower Lit Button Down Image
     ButtonsUpLit(X).IsDropped=True ' Lower Litit Button Up Image
       Buttons(X).UserValue=False   ' Mark button unLit
  Next
  '
  '   All Trigger Wires Showing
  '
    For X=0 to (Wires.Count-1)
     Wires(X).IsDropped=True
  Next
  '
  RollFlag=0                          'No Ball Rolling Yet
  RollVary=0                          'Init Roll Sound Variation flag
  ShootBall=0                       ' Next Move is Pull
  '
  '   Create Balls in Trough
  '
  BIT=0               ' No Balls in Lifter Tray
  BallID=1              ' reset Ball ID
  BallDrained(0)=0          ' None drained
  For X=1 to 5
     EVAL("Load"&X).Enabled=False   ' Disable Load Kicker
     Set Y=EVAL("Load"&X).CreateBall  ' Create Ball
     Y.UserValue=X          ' ID it
     EVAL("Load"&X).Kick 0, 0     ' Release Ball to Trough
     Y.Velx=0             ' Do our best to stabilize it
     Y.VelY=0
     Y.VelZ=0
  Next
  BallsLoaded=5           ' Say all balls loaded
  BallsLost=5             ' Say Rolled in from Playfield
  '
  '  5 Balls in the visible trough
  '
    FlippersTimer.Enabled=True      ' Start VP9 Primitive Timer
  '
  On Error Resume Next          'Needed if highscore/credits have never been saved in this table name
  Credits=Cdbl(LoadValue(TableName,"Credits"))
  If Credits="" Then Credits=1      'If there was no entry, then set to default value
  DisplayCredits()
     '
  HighScore=Cdbl(LoadValue(TableName,"HighScore"))
  If HighScore="" Then HighScore=246      ' If there was no entry, then set to default value
  If DisplayHighScore=True Then
     SetHS(HighScore)             ' Put On Display
  End If
     '
  Points=Cdbl(LoadValue(TableName,"Points"))  ' Load Points value from previous game
  If Points="" Then Points=0          ' If there was no entry, then use default value
  DisplayPoints()               ' Display on Back Glass
  '
  Score=Cdbl(LoadValue(TableName,"Score"))  ' Load score value from previous game
  If Score="" Then Score=455          ' If there was no entry, then use default value
  SReelHuns=  INT((Score MOD 1000) / 100)   ' Init Scoring Reel Digits
  SReelTens = INT((Score MOD 100) /10)
  SReelUnits = (Score MOD 10)
  MUnits=0                  ' Init Scoring Motor to stopped
  MTens=0
  DisplayScore()                ' Display Scoring Lights
    '
 End Sub
'
' *********************************
' ***** Keys Pressed Handling *****
' *********************************
Sub Table_KeyDown(ByVal Keycode)            'The usual keycode stuff

    If Tilted=False Then

    If Keycode=LeftFlipperKey Then
      LeftFlipper.RotateToEnd
      PlaySoundAtVol "FlipperUp", LeftFlipper, 1
            If BuzzCount < 2 Then
         PlayLoopSoundAtVol "Buzz", LeftFlipper, 1                 ' Start Flipper Solenoid Buzz
         BuzzCount=BuzzCount+1
             End If
    End If


    If Keycode=RightFlipperKey Then
      RightFlipper.RotateToEnd
      PlaySoundAtVol "FlipperUp", RightFlipper, 1
            If BuzzCount < 2 Then
                PlayLoopSoundAtVol "Buzz", RightFlipper, 1                 ' Start Flipper Solenoid Buzz
         BuzzCount=BuzzCount+1
             End If
    End If

    If Keycode=LeftTiltKey Then
     Nudge 270,0.95
      CheckTilt()
   End If

    If Keycode=RightTiltKey Then
      Nudge 90,0.95
     CheckTilt()
   End If

    If Keycode=CenterTiltKey Then
     Nudge 0,0.95
      CheckTilt()
   End If
   ' Thalamus - added mechanicaltilt
    If Keycode=MechanicalTilt Then
      CheckTilt()
   End If
  '
  ' End flipper and Nudge TILT qualification
  '
     End If

  If Keycode=PlungerKey Then
     ShootBall=False              ' Indicate Animating Pull
     Plunger.PullBack
  End If

  If Keycode=AddCreditKey Then
     PlaySound "Coin"
     CreditReel.TimerEnabled=True
  End If

  If Keycode=StartGameKey And Credits>0 Then
     If (GOVPhase) <> 0 or (DelayedGOV=True) or (BallsLoaded <> 5) Then
      DelayedStart=True         'let End of game Sequence Finsh First
     Else
     Plunger.TimerEnabled=True      'Start the "Game Start" Reset Sequence
     End if
  End If

  ' "A" Activates Ball Lifter
  If Keycode=RightMagnaSave And BallsLeft>0 AND Tilted=False Then
     PlaySound"BallOut"
     Lifter.TimerEnabled=True       'Activates the ball lifter Timer
  End If

  ' Either "R" or "I" displays rules
  If Keycode=23 or KeyCode = 19 Then
    Rules()
  End If

  ' "2" = Reset High Score to 246
  If KeyCode=3 Then
    HighScore = 246
    SaveValue TableName,"HighScore",HighScore   'Save score value for next time
    If DisplayHighScore=True Then
       SetHS(HighScore)
    End If
  End If

  ' "3" = Reset Credits
  If KeyCode=4 Then
     Credits=0
     DisplayCredits()
     SaveValue TableName,"Credits",Credits  'Save the credits
  End If


End Sub
'
'  -----------------------
'  Flipper Tracking Timers
'  -----------------------
'
'
'  VP9 Version
'
Sub UpdateFlipperLogos
    LeftFlipperP.Rotz = LeftFlipper.CurrentAngle
    RightFlipperP.Rotz = RightFlipper.CurrentAngle

End Sub

Sub FlippersTimer_Timer()
    UpdateFlipperLogos
End Sub
'
' **********************************
' ***** Keys Released Handling *****
' **********************************
'
Sub Table_KeyUp(ByVal Keycode)

    If Keycode=LeftFlipperKey Then
      LeftFlipper.RotateToStart
      PlaySoundAtVol "FlipperDown", LeftFlipper, 1
      StopSound "Buzz"                        ' End Flipper Solenoid Buzz
      BuzzCount=BuzzCount-1
    End If

    If Keycode=RightFlipperKey Then
      RightFlipper.RotateToStart
      PlaySoundAtVol "FlipperDown", RightFlipper, 1
      StopSound "Buzz"                        ' End Flipper Solenoid Buzz
      BuzzCount=BuzzCount-1
    End If


  If Keycode=PlungerKey Then
    ShootBall=True                              'Reverse plunger animation
    Plunger.Fire
    If BallAtPlunger = True Then
       PlaySoundAtVol "PlungerRoll", Plunger, 1
    Else
       PlaySoundAtVol "PlungerSpring", Plunger, 1
    End If
  End If
End Sub
'
' **********************************************************
' ***** Routines to Handle Ball Motion/Sounds on Table *****
' **********************************************************
'
'  Keep Track of whether Ball Is At Plunger or Not
'
Sub OnPlunger_Hit()
  BallAtPlunger = True
End Sub

Sub OnPlunger_UnHit()
  BallAtPlunger = False
  RollBack=0              'Init Look for Roll-back Sound
End Sub
'
' -------------------
' Handle Ball At Gate
' -------------------
'
' Add Gate Noise On Ball Entering Playfield
' Also add bounce sounds as ball bounces back and froth at top
'
' Gate hit as ball enters playfield
Sub Gate_Hit()                    'Just for the sound
  PlaySoundAtVol "BallRoll1", ActiveBall, 1
  RollBack=0                        ' raise Stopper at right end of ball trough
End Sub
'
' Hit gate going wrong way at top
Sub Rebound_Hit()
  If FirstOut > 0 Then
     PlaySoundAtVol "Gate5", ActiveBall, 1              ' Gate Bounce Sound
   Else
    FirstOut=1
  End If
End Sub
'
' -----------------
' Hit Rebound Wheel at top left on Gate Exit
' -----------------
'
Sub Wheel_Hit()
  PlaysoundAtVol "rubber", ActiveBall, 1
   If FirstOut < 2 Then
     FirstOut=2
     AddSound "RollMore"
   End If
End Sub
'
'
' *******************************************
' ***    Ball Rolling Sound Generator     ***
' ***    Learned Concept from PacDude's   ***
' ***    Excellent Centaur table          ***
' *******************************************
'
' Ball Roll Sensors
'
' Sensor for when ball rolls on Shooting Lane
'
Sub LaneSwitch_Hit()
  If RollBack=1 Then
     ' Failed to Clear Gate on Plunger Release
     PlaySoundAtVol "Roll9", ActiveBall, 1
  Else
     ' On way out of lane from release
     RollBack=1
  End If
End Sub
'
' Sensor for Upper Right Playfield
'
Sub RollSensor2_Hit()
  RollOnTo()
End Sub
'
Sub RollSensor2_UnHit()
  RollFlag = 0
End Sub
'
' Upper Left Playfield
Sub RollSensor3_Hit()
  RollOnTo()
End Sub
'
Sub RollSensor3_UnHit()
  RollFlag=0
End Sub
'
' Sensor for MidField Above Kickers
Sub RollSensor4_Hit()
  RollonTo()
End Sub
'
Sub RollSensor4_UnHit()
  RollFlag = 0
End Sub
'
' Sensor for Lower Playfield
Sub RollSensor5_Hit()
  RollonTo()
End Sub
'
Sub RollSensor5_UnHit()
  RollFlag = 0
End Sub
'
'  If Ball Rolling onto sensor at speed, make roll sound
'
' Minimum Speeds to Play the Ball Roll Sound
Const SpeedX1 = 5 : Const SpeedX2 = -5
Const SpeedY1 = 5 : Const SpeedY2 = -5
'
Sub RollOnTo()
  If RollFlag = 0 and (ActiveBall.VelX > SpeedX1 OR _
             ActiveBall.VelX < SpeedX2 OR _
             ActiveBall.VelY > SpeedY1 OR _
             ActiveBall.VelY < SpeedY2) Then
     If RollVary < 2 Then
      AddSound "Roll9"
      RollVary = RollVary + 1
     Else
      AddSound "RollMore"
      RollVary=0
     End if
     RollFlag = 1
  End If
End Sub
'
'*****************************************************
'****  Ball Trough, and Drain management routines ****
'*****************************************************
'
'----------------
'  Ball hit Drain
' (Only happens at game start to
'  empty visible ball trough)
'----------------
'
Sub Drain_Hit()
  Me.DestroyBall              ' Delete Ball From table
  If BallsLost > 0 Then
     BallsLost=0              ' Ball Stranded on Field,
     BallsOnField=0           ' Say none on PlayField
     BallDrained(0)=0           ' Clear Drained List
     BallsLeft=5              ' Say 5 Left to Lift
  End If
End Sub
'
'------------------------------------
' Ball Crossed Threshold into visible ball drain
'------------------------------------
'
Sub DrainTrig_Hit()
  Dim ID
  ID=ActiveBall.UserValue             ' Grab ID of Ball
  If InProgress=False Then
     Exit Sub                   ' Game Init
  End If
  ActiveBall.VelX=ActiveBall.VelX*.9
  ActiveBall.VelY=ActiveBall.VelY*.9
  If BallDrained(0) <> 0 Then
     For X=1 to BallDrained(0)
      If BallDrained(X) = ID Then
       Exit Sub               ' Don't Count Same Ball Twice
      End If
     Next
  End If
  TrayHit=False                 ' Allow One Hit Noise
  BallDrained(0)=BallDrained(0)+1         ' Add to list of Drained
  BallDrained(BallDrained(0))=ID          ' Say this one is drained
  DrainTrig.UserValue=1
  DrainTrig.TimerEnabled=True           ' Debounce Trough Return
  BallsLost=BallsLost+1             ' Say One More Ball Lost
  BallsOnField=BallsOnField-1           ' and not on PlayField
  If Tilted=True Then               ' Game over if tilted
     InProgress=False               ' Game has finished (due to tilt)
     DisableFlippers()              ' In case Holding Flippers
     BallsLeft=0                  ' No More Balls Left
     GOVPhase=1                 ' Game Over Processing Has Begun
     M1sReelA.TimerEnabled=True         ' Start Game Over Sequence
     Exit Sub
  End If
  If BallsLost >= 5 Then              ' Game is over after 5 Balls
     GOVPhase=1                 ' Game Over Processing Has Begun
     If DisplayHighScore=True Then
      SetHS(HighScore)              ' Display the current HIGH SCORE
     End If
     M1sReelA.TimerEnabled=True         ' Start Game Over Sequence
  End If
End Sub
'
Sub DrainTrig_Timer()
  DrainTrig.TimerEnabled=False
  DrainTrig.userValue=0             ' End Debounce Window
End Sub
'
'------------------------------
' Next Ball "Lift" Sequence
'------------------------------
'
' After Delay from Drain Hit, Lifter Timer Creates
' A Ball to "Lift" to the plunger.
'
Sub Lifter_Timer()
  DrainWallHit=False          ' allow DrainWall hit Sound
  Lifter.CreateBall.UserValue=BallID  ' Create a Ball on the Lifter and ID it
  BallID=BallID+1           ' Increment Ball ID for next time
  Lifter.Kick 270,3         ' Kick it up to the plunger lane
  BallsLeft=BallsLeft-1       ' One less Ball remains
  BallsOnField=BallsOnField+1     ' One more is on the Playfield
  FirstOut=0
  Lifter.TimerEnabled=False     ' Disable This Timer 'Till Next time
End Sub
'
' If Ball ends up back in trough, kick it out
 '
Sub Lifter_Hit()
  Lifter.Kick 270, 3
End Sub
'
'*****************************************************
'****  End of the trough/lifter management routines ****
'*****************************************************
'
' ******************************
' **** Scoring and credits *****
' ******************************
'
'  AddScore is main scoring routine
'
'  Note: This version assumes ScoreToAdd is only 1, 5, 10, 30, or 50
'
'  Motor = 0 For normal Scoring call
'          1 for call from Scoring Motor
'
Sub AddScore(ScoreToAdd, Motor)
  If InProgress=False Or Tilted=True Then
       Exit Sub                 ' Disable score when TILTed or not started
    End If
    '
    ' Display the score
    '
    ' Note: "Scoring Motor" uses SMotor Timer
    '
    If (ScoreToAdd = 5) OR (ScoreToAdd=30) OR (ScoreToAdd=50)  Then
       Motor=1                ' kill sound here, let motor do it
       If SMotor.Enabled=False Then
          If ScoreToAdd=5 Then
             MUnits =  5          ' 5 = Pulse Units 5 Times
          Else
             MTens =  ScoreToAdd/10     ' 30 or 50 = Pulse Tens 3 or 5 Times
          End If
          M5Q=0
          M50Q=0              ' no scoring runs queued
          '
          ' Scoring Motor Not Running, Start it up
          '
          MotorPhase=0            ' Cycle hasn't started yet
          SMotor.Interval=140       ' Cycle = 140ms per pulse
          SMotor.Enabled=True       ' Start Scoring Motor Running
       Else
          '
          ' Queue a max of 1 of each 5 and 50
          '
          If ScoreToAdd=5 Then
             M5Q=1              ' 5 = Pulse Units 5 Times
          Else
             M50Q=1             ' 50 = Pulse Tens 5 Times
          End If
       End If
    Else
       '
       ' Score is either 1 or 10, Score it Now
        '
       Score = Score+ScoreToAdd           'Total of current score
       SReelHuns=  INT((Score MOD 1000) / 100)    'Get New Scoring Reel Digits
       SReelTens = INT((Score MOD 100) /10)
       SReelUnits = (Score MOD 10)
       If ScoreToAdd=1 Then
           LightsRotate()
       End IF
       DisplayScore()                ' Put Score on backglass
    End If
    '
    '  Check for replays Earned on Scoring
    '
    If Score=>Replay1 And Replay1Paid=False Then
       AddReplay()                    'Made Replay 1 Score
       Replay1Paid=True
    End If
    If Score=>Replay2 And Replay2Paid=False Then
       AddReplay()                    'Made Replay 2 Score
       Replay2Paid=True
    End If
    If Score=>Replay3 And Replay3Paid=False Then
       AddReplay()                    'Made Replay 3 Score
       Replay3Paid=True
    End If
    If Score=>Replay4 And Replay4Paid=False Then
       AddReplay()                    'Made Replay 4 Score
       Replay4Paid=True
    End If
       If Score=>Replay5 And Replay5Paid=False Then
       AddReplay()                    'Made Replay 5 Score
       Replay5Paid=True
     End If
'
'  Play Appropriate Bell sounds to match scoring
'
    If (Bells=True) AND (Motor=0) Then
       PlaySound ScoreToAdd           ' Sound Bell on 1 and 10 Sounds
    End if
End Sub
'
'  "Scoring Motor" Timer Routine
'
Sub Smotor_Timer()
     MotorPhase=MotorPhase+1
     If MotorPhase = 6 Then
        Exit Sub
     End If
     If MotorPhase > 6 Then
        MotorPhase=0
        If M5Q+M50Q=0 Then
           SMotor.Enabled=False     ' Stop Scoring Motor
           Exit Sub
        Else
           If M5Q > 0 Then
              MUnits=5          ' do queued 50,000 run
              M5Q=0
           Else
              MTens=5         ' do queued 500,000 run
              M50Q=0
           End If
       End If
    End If
    If MUnits > 0 Then
       MUnits=Munits-1
       AddScore 1, 0          ' Score a units digit pulse
       Exit Sub
    End If
    If MTens > 0 Then
       MTens = MTens-1
       AddScore 10, 0         ' Score a Tens Digit pulse
    End If
End Sub
'
'  Player Beat the High Score, Award It
'
Sub PayHigh()                     'When HIGH SCORE is beaten, go here
  For X=1 to HighScoreReward
     AddReplay()                    'Always add at least 1 replay
  Next
End Sub
'
' Made a Replay, Award it and "Knock"
'
Sub AddReplay()
    DelayedReplays=DelayedReplays+1       ' Add Delayed Replay
    HSTens.TimerInterval=120
  HSTens.TimerEnabled=True          'Enable the timer for Overlapping Replays
End Sub
'
' *******************************************************
' ****  Multiple Replay Award With Knocker Sequence  ****
' ****                                               ****
' ****    HSTens.Timer_Enabled=True to start         ****
' *******************************************************
'
'  Delayed Replays are added at 120ms intervals
'  based on value DelayedReplays
'
Sub HSTens_Timer()
  If DelayedReplays > 0 Then
     PlaySound"Knocker"           'Play the sound
     If Credits < MaxCredits Then       ' Our credit Reel maxes at 29
      Credits=Credits+1           'Add a credit
      DisplayCredits()            'Update the credit display
     End If
     DelayedReplays=DelayedReplays-1      'One Less To Go
  Else
     DelayedReplays=0             'No More Delayed Replays
     HSTens.TimerEnabled=False
  End If
End Sub
'
' ******************************
' **** Tilt Bobber Simulator ***
' ******************************
'
'  This routine is called each time game is Nudged
'  to Check to see if Player has tilted the game.
Sub CheckTilt()                     'Called when table is nudged
    Dim I
  Count1=Count1+1                   'Add to tilt count (hit lasts 1 second)
  If Count1>TiltSensitivity And Tilted=False Then   'If more than Allowed Counts then TILTED
     Tilted=True
     '
     '  PlayField Lights Off
     '
     LtSpecial.IsDropped=True       ' Special Light Off
        LtGI1.IsDropped=True          ' GI Lights Off
        LtGI2.IsDropped=True
       '
       For I=0 to (KickersLit.Count-1)
          KickersLit(I).IsDropped=True    ' Assure Kickers are
          KickerCovers(I).IsDropped=False ' Unlit
          Kickers(I).UserValue=0      ' Say Kicker Not Lit
       Next
        '
        For I=0 to (LtNum.Count-1)
           LtNum(I).IsDropped=True      ' Number Lights Off
        Next
        '
        For I=0 to (LtLet.Count-1)
           LtLet(I).IsDropped=True      ' Letter Lights Off
        Next
        '
        For I=0 to (LtTarg.Count-1)
           LtTarg(I).IsDropped=True     ' Target Lights Off
        Next
       '
     BumpersOff()             ' Set All Bumpers Off
        '
       For I=1 to (Buttons.Count)
        ButtonOff(1)            ' Buttons Off
       Next
       '
       If (InProgress=True) AND (BallsOnField > 0) Then
           DelayedGOV=True                'Say game Is Effectively over
       End If
     DisplayTilt(1)                 'Put "TILT" On Backglss
     PlaySound "Reset1"
     Disable(True)                  'Disable slings, bumpers etc
  End If
End Sub
'
' Tilt Timer Cycles every 250ms (1/4 sec) when game is running
'
Sub TiltTimer_Timer()       'Used to simulate a tilt bob
  If Count1>0 Then
     Count1=Count1-0.25     'Subtract .25 every 250ms
  End If
End Sub
'
' *************************************************
' ****         Coin Deposited Sequence         ****
' ****                                         ****
' ****  CreditReel.Timer_Enabled=True to start ****
' *************************************************
'
' Timer started when a coin is entered.
' Timer comes here every 200ms until stopped in Cycle 10 (2 secs)
'
' Adds number of credits indicated by CreditsPerCoin
' Less the one that was added before the sequence was started
Sub CreditReel_Timer()
  Count=Count+1
  Select Case Count
    Case 5:PlaySound"Motor"
    Case 6:If CreditsPerCoin>1 Then AddCredit
    Case 7:If CreditsPerCoin>2 Then AddCredit
    Case 8:If CreditsPerCoin=5 Then AddCredit
    Case 9:AddCredit            'Always add at least 1 credit
    Case 10:
      If CreditsPerCoin>3 Then AddCredit
      Count=0
      Me.TimerEnabled=False
  End Select
End Sub
'
' Add a Credit on the Backglass counter
Sub AddCredit()                     'Called by Credit_Timer
  If Credits < MaxCredits Then
     Credits=Credits+1                'Add a credit to the counter
     DisplayCredits()                 'Update the credit display
  End If
  PlaySound"StepUp"                 'Play stepping sound
End Sub
'
' *************************************
' ****  Game Start Reset Sequence  ****
' *************************************
'
' This Routine inspired by Bendigo's 60's Gottlieb Tutorial table
' Template5a.VPT (Thanks!).
'
' When started, Plunger Timer gets here every 200ms until we are done.
' Entire startup process is staged to take 2.6 seconds to
' emulate an EM machine Game Start Reset sequence
'
Sub Plunger_Timer()               'Game Start Reset sequence
  GSPhase=GSPhase+1             'Add 1 every time the counter cycles
  Select Case GSPhase             ' Select Cycle to do this tick
'
'  On Cycle 1 (.2s after start pressed) play EM motor sound
    Case 1:
       PlaySound"GameStart"           'Start the motor sound
       Disable(True)              'Disable playfield objects
       BallID=1                 ' Reset Ball ID
       BallDrained(0)=0             ' None Drained Yet
       TroughOpening=True           ' Say Trough is Opening
       TroughOpenHalf.IsDropped=True      ' Animate Half Open Trough
       TroughOpenHalf.TimerEnabled=True     ' Start Timer to Complete Open
'
'  Cycle 2 (.4s after start) play a Reset sound and
'  Release the balls from the visible trough
    Case 2:
       Credits=Credits-1            'Subtract a credit
       DisplayCredits()             'Update the credit display
'      DisplayGOV(0)              'Clear GAME OVER message
       Score=0                  'Reset SCORE
       Points=0                 ' Including Points
       MadeSpecial=False                        ' Haven't Made Lane Special yet
       DisplayTilt(0)             ' Clear "TILT" On Backglass
       BallsLost=0                'Reset balls used
       BallsLeft=5                'And Balls Left to Lift/Play
       If BallsOnField > 0 Then
        BallsLeft = BallsLeft-BallsOnField  'Start Hit Mid-previous Game
       End If
 '
       Replay1Paid=False            'Reset High Scores Hit flags
       Replay2Paid=False
       Replay3Paid=False
       Replay4Paid=False
       Replay5Paid=False
       HighScorePaid=False
'
       InProgress=True              'Game has started
       Tilted=False               ' and isn't yet tilted
       If HighScore=0 Then            'Make sure we don't pay for 1st ever game
        HighScorePaid=True
        DisplayHighScore=False
       End If
'
'  Cycle 3 (.6s after start pressed) Set up lights
'
    Case 3:
       '
       '  Set Initial Game Start Light States:
       '
            LtGI1.IsDropped=False       ' GI Lights On
            LtGI2.IsDropped=False
            '
           For X=0 to (LtNum.Count-1)
              LtNum(X).IsDropped=False      ' Number Lights ON
           Next
         '
           For X=0 to (LtLet.Count-1)
              LtLet(X).IsDropped=False      ' Letter Lights ON
           Next
           '
           For X=0 to (LtTarg.Count-1)
              LtTarg(X).IsDropped=True      ' Target Lights OFF
           Next
       '
       LtSpecial.IsDropped=True       ' Special Light OFF
           '
           For X=0 to (KickersLit.Count-1)
              KickersLit(X).IsDropped=True    ' Init Kickers
              KickerCovers(X).IsDropped=False ' Unlit
              Kickers(X).UserValue=0      ' Say Kicker Not Lit
           Next
           '
           For X=1 to 12
              Spotted(X)=0            ' Reset 1-12 spotted array
           Next
       '
       '  Set game Start Bumpers and Buttons
       '
       BumpersOn()              ' Set All Bumpers On
           For X=1 to (Buttons.Count)
              ButtonOff(X)            ' also Buttons Off
           Next
           LightsRotate()           ' Set 100,000 When Lit Bumpers Correctly
'
'  Cycle 4 (.8s after start pressed) just generates delay
'
'  Cycle 5 (1 full second after start) We Reset the score counter
'
    Case 5:
       ResetLightsToZero()          'Reset score counters
       DisplayPoints()            ' Display Points
 '
       If HighScore>0 AND DisplayHighScore = True Then
        SetHS(HighScore)            'Display highscore
       End If
'
' Cycle 6 (1.2s after start pressed) is delay
'
' Cycle 7 (1.4s after start pressed) repeats sound from Cycle 2
'
    Case 7:
       TroughOpen.IsDropped=True      ' Show ball Trough Closed
       TroughOpenHalf.IsDropped=False   ' Animate Closing
       TroughOpening=False          ' Say Closing
       TroughOpenHalf.TimerEnabled=True   ' Start timer to complete animating close
       BallRetain.IsDropped=False     ' Raise Drain Wall
'
' Cycles 8 through 12 are more delay
'
' Cycle 13 (2.6 seconds after start pressed)
'       End the startup Reset sequence and allow the game to start
'
    Case 13:
       GSPhase=0              'Reset Phase for next time
       Disable(False)           'Enable any disabled objects
       InProgress=True
       Me.TimerEnabled=False        'Disable this timer
  End Select
End Sub
'
'  Routine to complete animation of Trough Opening/Closing
'
Sub TroughOpenHalf_Timer()
  If TroughOpening=True Then
     TroughOpenHalf.IsDropped=True      ' Drop half open frame
     TroughOpen.IsDropped=False       ' Show ball Trough Open
     BallRetain.IsDropped=True        ' Allow Balls to drain
  Else
     TroughOpenHalf.IsDropped=True      ' Show Trough Fully Closed
  End If
  TroughOpenHalf.TimerEnabled=False     ' Stop Timer
End Sub
'
'
' **********************************
' ****  Game Over Sequence  ********
' **********************************
'
' (Uses Timer associated with a scoring EM Reel
'  It runs at 100ms ticks when enabled to end Game)
'
Sub M1sReelA_Timer()
  GOVPhase=GOVPhase+1               'Increment the Cycle counter
  Select Case GOVPhase
    Case 2:
       DelayedGOV=False             'End Any Tilt Delay
'
'  Cases 1-4 are just 400ms of delay
'
'  0.5s after end of game
    Case 5:
'     PlaySound "Motor"           'Play the score motor sound
'
'  Cases 6-15 are 900ms more delay
'
'   1.6s after end of game
    Case 16:
'      DisplayGOV(1)                         'Display GAME OVER
'
' Cases 17-20 are 400ms more delay
'
' 2.1s after End Of game sequence began we have the final cycle
    Case 21:
       SaveValue TableName,"Score",Score    'Save score value for next time
       SaveValue TableName,"Points",Points
       SaveValue TableName,"Credits",Credits  'Save the credits
       If Score>HighScore Then
        SaveValue TableName,"HighScore",Score'Save the HighScore
        HighScore=Score
        If DisplayHighScore=True Then
         SetHS(HighScore)
        End If
       End If
       InProgress=False             'Game is finished
       GOVPhase=0               'Reset Phase for next time
       Me.TimerEnabled=False          'Disable this timer
       If DelayedStart=True Then
        DelayedStart=False          ' Start Pressed Before End Of game Finished
        Plunger.TimerEnabled=True       ' Start the "Game Start" Reset Sequence
       End if
  End Select
End Sub
'
' ***** Subroutine to Display the Game Rules *****
'
Sub Rules()
  Msg(0)=TableName&Chr(10)&Chr(10)
  Msg(1)=""
  Msg(2)="                            GREEN PASTURES"
  Msg(3)=""
  Msg(4)="  HITTING NUMBERS 1 TO 12 IN THAT ORDER AWARDS 1 REPLAY AND LITES HOLES FOR SPECIAL."
  Msg(5)=""
  Msg(6)="  HITTING A-B-C-D ROLLOVERS LITES TARGETS, ROLLOVER BUTTONS AND LITES ONE ROLLOVER FOR SPECIAL."
  Msg(7)=""
  Msg(8)="  SPECIAL WHEN LIT ROLLOVERS AWARD 1 REPLAY."
  Msg(9)=""
  Msg(10)=" REPLAYS for Scores 4,600,000, 5,500,000, 5,700,000, 6,000,000, and 6,500,000."
  Msg(11)=""
  Msg(12)="  REPLAYS for 50, 70, 80, 90 AND 98 POINTS."
  Msg(13)=""
  Msg(14)="  Press 5 to Deposit Coin, 1 to Start a game"
  Msg(15)="  Press A to lift next ball into play"
  Msg(16)=""
  Msg(17)="  Press 2 to Reset high Score to 2,460,000."
    Msg(18)=""
    Msg(19)="  Press 3 to Clear Credits to 0"
  For X=1 To 19
    Msg(0)=Msg(0)+Msg(X)&Chr(13)
  Next
  MsgBox Msg(0),,"         Instructions and Rule Card"
End Sub
'
' ******************************
' ****  Handle Bumper Hits *****
' ******************************
'
' -----------------------------
'    "1"/"12" Passive Bumper
' -----------------------------
'
Sub B1BASE_Hit
  If B1Top.UserValue=False Then
     PlaySoundAtVol "Rubber", ActiveBall, 1
       AddSound "RollMore"
        SpotNumber(12)          ' Add 12 or 1 If in Sequence
     SpotNumber(1)
     AddScore 1, 0          '10,000 Points
  End If
  '
  '  Debounce Ball "Laying" on Passive Bumper
  '
  B1Top.TimerEnabled=False      'Stop Timer If Running
  B1Top.UserValue=True        'Say Debounce Started
  B1Top.TimerInterval=100       'Debounce for 100ms
  B1Top.TimerEnabled=True       'Start Debounce Window
End Sub
'
Sub B1Top_Timer()
  B1Top.TimerEnabled=False      'Stop Timer
  B1Top.UserValue=False       'Say Hit Counts Again
End Sub
'
' ------------------------
'    "2"/"11" POP Bumper
' ------------------------
'
Sub B2BASE_Hit
  PlaySoundAtVol "Bumper", ActiveBall, 1
    AddSound "RollMore"
    SpotNumber(11)            ' Spot 2 or 11 if in sequence
  SpotNumber(2)
  AddScore 1, 0           ' 10,000 Points
  If B2State <> 0 Then
     Exit Sub             ' Don't re-enter animation
  End If
  B2Rings()             ' Kick off Ring Animation
  B2Ring3.TimerEnabled=True
End Sub
'
'  Bumper 2 Pop Ring Animation Cycle Routine
'
Sub B2Rings()
  B2State=B2State+1
  Select Case B2State
    Case 1:
     B2Ring3.IsDropped=False
    Case 2:
     B2Ring2.IsDropped=False
    Case 3:
     B2Ring1.IsDropped=False
    Case 4:
     B2Ring2.IsDropped=False
    Case 5:
     B2Ring3.IsDropped=False
  End Select
End Sub
'
' Bumper 2 Ring Animation Timer
'
Sub B2Ring3_Timer()
  Select Case B2State
    Case 1:
      B2Ring3.IsDropped=True
    Case 2:
      B2Ring2.IsDropped=True
    Case 3:
      B2Ring1.IsDropped=True
    Case 4:
      B2Ring2.IsDropped=True
    Case 5:
      B2Ring3.IsDropped=True
  End Select
  If B2State<5 Then
     B2Rings()
  Else
     B2State=0
     B2Ring3.TimerEnabled=False
  End If
End Sub
 '
' -------------------
' "3" Passive Bumper
' -------------------
'
Sub B3BASE_Hit
  If B3Top.UserValue=False Then
     PlaySoundAtVol "Rubber", ActiveBall, 1
       AddSound "RollMore"
     SpotNumber(3)          ' Add to sequence made
     AddScore 1, 0          ' 10,000 Points
  End If
  '
  '  Debounce Ball "Laying" on Passive Bumper
  '
  B3Top.TimerEnabled=False      'Stop Timer If Running
  B3Top.UserValue=True        'Say Debounce Started
  B3Top.TimerInterval=100       'Debounce for 100ms
  B3Top.TimerEnabled=True       'Start Debounce Window
End Sub
'
Sub B3Top_Timer()
  B3Top.TimerEnabled=False      'Stop Timer
  B3Top.UserValue=False       'Say Hit Counts Again
End Sub
'
' ----------------------------
'    "4"/"9" Passive Bumper
' ----------------------------
'
Sub B4BASE_Hit
  If B4Top.UserValue=False Then
     PlaySoundAtVol "Rubber", ActiveBall, 1
       AddSound "RollMore"
       SpotNumber(9)
     SpotNumber(4)          ' Add to sequence made
     AddScore 1, 0          ' 10,000 Points
  End If
  '
  '  Debounce Ball "Laying" on Passive Bumper
  '
  B4Top.TimerEnabled=False      'Stop Timer If Running
  B4Top.UserValue=True        'Say Debounce Started
  B4Top.TimerInterval=100       'Debounce for 100ms
  B4Top.TimerEnabled=True       'Start Debounce Window
End Sub
'
Sub B4Top_Timer()
  B4Top.TimerEnabled=False      'Stop Timer
  B4Top.UserValue=False       'Say Hit Counts Again
End Sub
 '
' -------------------
' "5" Passive Bumper
' -------------------
'
Sub B5BASE_Hit
  If B5Top.UserValue=False Then
     PlaySoundAtVol "Rubber", ActiveBall, 1
       AddSound "RollMore"
     SpotNumber(5)          ' Add to sequence made
     AddScore 1, 0          ' 10,000 Points
  End If
  '
  '  Debounce Ball "Laying" on Passive Bumper
  '
  B5Top.TimerEnabled=False      'Stop Timer If Running
  B5Top.UserValue=True        'Say Debounce Started
  B5Top.TimerInterval=100       'Debounce for 100ms
  B5Top.TimerEnabled=True       'Start Debounce Window
End Sub
'
Sub B5Top_Timer()
  B5Top.TimerEnabled=False      'Stop Timer
  B5Top.UserValue=False       'Say Hit Counts Again
End Sub
 '
' --------------------------
'   "6"/"7" Passive Bumper
' --------------------------
'
Sub B6BASE_Hit
  If B6Top.UserValue=False Then
     PlaySoundAtVol "Rubber", ActiveBall, 1
       AddSound "RollMore"
       SpotNumber(7)          ' Spot 7 or 6 if next in sequence
     SpotNumber(6)
     AddScore 1, 0          ' 10,000 Points
  End If
  '
  '  Debounce Ball "Laying" on Passive Bumper
  '
  B6Top.TimerEnabled=False      'Stop Timer If Running
  B6Top.UserValue=True        'Say Debounce Started
  B6Top.TimerInterval=100       'Debounce for 100ms
  B6Top.TimerEnabled=True       'Start Debounce Window
End Sub
'
Sub B6Top_Timer()
  B6Top.TimerEnabled=False      'Stop Timer
  B6Top.UserValue=False       'Say Hit Counts Again
End Sub
'
' -------------------------------------
'    Left 100,00 When Lit POP Bumper
' -------------------------------------
'
Sub B7BASE_Hit
  PlaySoundAtVol "Bumper", ActiveBall, 1
    AddSound "RollMore"
  If klight.state=1 Then   AddScore 10, 0         ' 100,000 WHen Lit

    If klight.state=0 Then   AddScore 1, 0          ' 10,000 Points

  If B7State <> 0 Then
     Exit Sub             ' Don't re-enter animation
  End If
  B7Rings()             ' Kick off Ring Animation
  B7Ring3.TimerEnabled=True
End Sub
'
'  Bumper 7 Pop Ring Animation Cycle Routine
'
Sub B7Rings()
  B7State=B7State+1
  Select Case B7State
    Case 1:
     B7Ring3.IsDropped=False
    Case 2:
     B7Ring2.IsDropped=False
    Case 3:
     B7Ring1.IsDropped=False
    Case 4:
     B7Ring2.IsDropped=False
    Case 5:
     B7Ring3.IsDropped=False
  End Select
End Sub
'
' Bumper 7 Ring Animation Timer
'
Sub B7Ring3_Timer()
  Select Case B7State
    Case 1:
      B7Ring3.IsDropped=True
    Case 2:
      B7Ring2.IsDropped=True
    Case 3:
      B7Ring1.IsDropped=True
    Case 4:
      B7Ring2.IsDropped=True
    Case 5:
      B7Ring3.IsDropped=True
  End Select
  If B7State<5 Then
     B7Rings()
  Else
     B7State=0
     B7Ring3.TimerEnabled=False
  End If
End Sub
'
' ------------------------
'    "8" Passive Bumper
' ------------------------
'
Sub B8BASE_Hit
  If B8Top.UserValue=False Then
     PlaySoundAtVol "Rubber", ActiveBall, 1
       AddSound "RollMore"
     SpotNumber(8)          ' Add to sequence made
     AddScore 1, 0          ' 10,000 Points
  End If
  '
  '  Debounce Ball "Laying" on Passive Bumper
  '
  B8Top.TimerEnabled=False      'Stop Timer If Running
  B8Top.UserValue=True        'Say Debounce Started
  B8Top.TimerInterval=100       'Debounce for 100ms
  B8Top.TimerEnabled=True       'Start Debounce Window
End Sub
'
Sub B8Top_Timer()
  B8Top.TimerEnabled=False      'Stop Timer
  B8Top.UserValue=False       'Say Hit Counts Again
End Sub
 '
'  -----------------------------------------
'     Right 100,000 When Lit "POP" Bumper
'  -----------------------------------------
Sub B9BASE_Hit
  PlaySoundAtVol "Bumper", ActiveBall, 1
    AddSound "RollMore"
  If klight.state=1 Then   AddScore 10, 0         ' 100,000 WHen Lit
  If klight.state=0 Then   AddScore 10, 0         ' 10,000 Points

    If B9State <> 0 Then
     Exit Sub             ' Don't re-enter animation
  End If
  B9Rings()             ' Kick off Ring Animation
  B9Ring3.TimerEnabled=True
End Sub
'
'  Bumper 9 Pop Ring Animation Cycle Routine
'
Sub B9Rings()
  B9State=B9State+1
  Select Case B9State
    Case 1:
     B9Ring3.IsDropped=False
    Case 2:
     B9Ring2.IsDropped=False
    Case 3:
     B9Ring1.IsDropped=False
    Case 4:
     B9Ring2.IsDropped=False
    Case 5:
     B9Ring3.IsDropped=False
  End Select
End Sub
'
' Bumper 9 Ring Animation Timer
'
Sub B9Ring3_Timer()
  Select Case B9State
    Case 1:
      B9Ring3.IsDropped=True
    Case 2:
      B9Ring2.IsDropped=True
    Case 3:
      B9Ring1.IsDropped=True
    Case 4:
      B9Ring2.IsDropped=True
    Case 5:
      B9Ring3.IsDropped=True
  End Select
  If B9State<5 Then
     B9Rings()
  Else
     B9State=0
     B9Ring3.TimerEnabled=False
  End If
End Sub
'
' -------------------
' "10" Passive Bumper
' -------------------
'
Sub B10BASE_Hit
  If B10Top.UserValue=False Then
     PlaySoundAtVol "Rubber", AtiveBall, 1
       AddSound "RollMore"
     SpotNumber(10)         ' Add to sequence made
     AddScore 1, 0          ' 10,000 Points
  End If
  '
  '  Debounce Ball "Laying" on Passive Bumper
  '
  B10Top.TimerEnabled=False     'Stop Timer If Running
  B10Top.UserValue=True       'Say Debounce Started
  B10Top.TimerInterval=100        'Debounce for 100ms
  B10Top.TimerEnabled=True        'Start Debounce Window
End Sub
'
Sub B10Top_Timer()
  B10Top.TimerEnabled=False     'Stop Timer
  B10Top.UserValue=False        'Say Hit Counts Again
End Sub
'
'******************************************************
'  Subroutine to Spot A Number in the Sequence 1-12
'*******************************************************
'
Sub SpotNumber(Num)
    Dim I
    If (Tilted=True) OR (Spotted(Num)=1) Then
       Exit Sub               ' Tilt or already spotted = get out now
    End If
    If (Num=1) OR (Spotted(Num-1)=1) Then
       Spotted(Num) =1            ' Spot if next unspotted number in sequence
    Else
       Exit Sub               ' Not next number in sequence, can't spot
    End If
    Select Case Num
        Case 1, 2, 4, 6, 7, 9, 11, 12
           EVAL("Lt"&Num).IsDropped=True    ' Unlight number made
           If Num=7 Then
              BumperOff(6)          ' UnLight 6/7 Bumper
           End If
           If Num=9 Then
              BumperOff(4)          ' UnLight 4/9 Bumper
           End If
           If Num=11 Then
              BumperOff(2)          ' UnLight 2/11 POP Bumper
           End If
           If Num=12 Then
              BumperOff(1)          ' UnLight 1/12 Bumper
           End If
        Case 3, 5, 8, 10
           BumperOff(Num)         ' UnLight Number Made
     End Select
     If Num = 12 Then
        AddReplay()             ' 1 Replay for making 1-12
        For I=0 to (KickersLit.Count-1)
           KickerCovers(I).IsDropped=True ' Light Kickers
          KickersLit(I).IsDropped=False   ' For Special
          Kickers(I).UserValue=1      ' Say Kickers Lit
       Next
     End If
 End Sub
'
'  ***************************
'        Handle Buttons
'  ***************************
'
' Trigger hit lowers buttons and does button rolled over function
'
Sub Buttons_Hit(Idx)
  If Buttons(Idx).UserValue=True Then
       ' Button Is Lit
     ButtonsUpLit(Idx).IsDropped=True   ' Drop button
     ButtonsDnLit(Idx).IsDropped=False  ' Raise Dropped Lit Button Image
       AddPoints(2)             ' 2 Points Lit
  Else
       ' Button is UnLit
     ButtonsDn(Idx).IsDropped=False   ' Raise Button Down unLit Image
       AddScore 1, 0            ' 10,000 UnLit
  End If
    AddSound "RollMore"
End Sub
'
'  Trigger Unhit starts delay to button up
'
Sub Buttons_UnHit(Idx)
    ButtonsDn(Idx).TimerInterval=75     ' button up after delay
  ButtonsDn(Idx).TimerEnabled=True    ' Enable animation timer
End Sub
'
' End of delay, raise button
'
Sub ButtonsDn_Timer(Idx)
  ButtonsDn(Idx).TimerEnabled=False   ' Stop Animation Timer
  ButtonsDn(Idx).IsDropped=True     ' Lower UnLit Button Down Image
    ButtonsDnLit(Idx).IsDropped=True    ' Lower Lit Button Down Image
    If Buttons(Idx).UserValue=True Then
       ButtonsUpLit(Idx).IsDropped=False  ' Raise Lit Button Up Image
    End If
End Sub
'
'  ----------------------------------
'  Subs to light and unlight buttons
'  ----------------------------------
'
Sub ButtonOn(Num)
     If (Buttons(Num-1).UserValue=True) Then
        Exit Sub                ' Button Already On
     End If
    If ButtonsDn(Num-1).IsDropped=True Then
       ButtonsUpLit(Num-1).IsDropped=False  ' Raise Lit Button Up Image
    Else
       ButtonsDn(Num-1).IsDropped=True
       ButtonsDnLit(Num-1).IsDropped=False
    End If
    Buttons(Num-1).UserValue=True     ' Mark Button Lit
End Sub
'
Sub ButtonOff(Num)
     If (Buttons(Num-1).UserValue=False) Then
        Exit Sub                ' Button Already Off
     End If
    If ButtonsUpLit(Num-1).IsDropped=False Then
       ButtonsUpLit(Num-1).IsDropped=True ' Lower Lit Button Up Image
    Else
       ButtonsDnLit(Num-1).IsDropped=True
       ButtonsDn(Num-1).IsDropped=False
    End If
    Buttons(Num-1).UserValue=False      ' Mark Button UnLit
End Sub
 '
'****************************
'**** Handle Targets hit ****
'****************************
'
' Idx    Target
' ----  ---------------
'  0     Lower Left Target
'  1     Upper Left Target
'  2     Upper Right Target
'  3     Lower Right Target
'
Sub Targets_Hit(Idx)
    playsoundAtVol "checking", ActiveBall, 1
    TargetsHit(Idx).Isdropped=FALSE   ' Raise Rearmost Image
    Targets(Idx).Isdropped=TRUE     ' Lower Frontmost Image
    Targets(Idx).TimerEnabled=TRUE    ' Start Animation Timer
     '
     If Tilted=True Then
       Exit Sub             ' Done if tilted
    End If
    '
    If (LtTarg(Idx).IsDropped=False) Then
       AddPoints(10)          ' 10 Points When Lit
    Else
       AddPoints(2)           ' 2 Points UnLit
    End If
End Sub

Sub Targets_Timer(Idx)
    Targets(Idx).TimerEnabled=FALSE   ' Stop Animation Timer
    Targets(Idx).Isdropped=FALSE    ' Raise Frontmost Target image
    TargetsHit(Idx).Isdropped=TRUE    ' Lower Rearmost Target Image
End Sub
'
'******************************************************
'    Handler the Kickers
'******************************************************
'
'  Idx     Target
'  ----  --------------------------------------------
'    0     Left Kicker
'    1     Right Kicker
'----------------------------------------------------
'
Const kForce = 0.65         ' Kicker attraction force
'
'----------------------------------------------------
'
'======================================
'  Center Kicker Trigger Zone Entered
'======================================
'
Sub KickerTrigs_Hit(Idx)
    Set aBall(Idx)=ActiveBall         ' Grab Pointer to this ball
    If (AttractBall(aBall(Idx), Kickers(Idx).X, Kickers(Idx).Y, 65, 3) > 20) Then
        KickerTrigs(Idx).TimerInterval=10   'start 10ms updater
        KickerTrigs(Idx).TimerEnabled=True
        aTicks(Idx)=0             ' reset tick counter
    End If
    AddSound "RollMore"
End Sub
'
'  Timer to do attract every 10ms while Ball in zone
'
Sub KickerTrigs_Timer(Idx)
    Dim D, S, I, J
    aTicks(Idx)=aTicks(Idx)+1         ' count 1 more tick
    D=AttractBall(aBall(Idx), Kickers(Idx).X, Kickers(Idx).Y, 65, kForce)
    I=aBall(Idx).VelY
    J=aBall(Idx).VelX
    S=SQR((I*I)+(J*J))
    '
    ' D= Distance of ball from kicker center (in VP units, 47.06 = 1 inch)
    ' S= Scalar speed of ball in VP Units/100ms  (4.706 = 1 inch/sec)
    '
'    If (S < 1) Then
'       Exit Sub                ' slower than 0.1 inch/sec
'    End If
    If (D < 14) Then
       '
       ' Inside hole, let kicker grab ball
       '
       Kickers(Idx).Enabled=True      ' Enable Kicker
       aBall(Idx).VelX=0
       aBall(Idx).VelY=1
       aBall(Idx).X=Kickers(Idx).X
       aBall(Idx).Y=Kickers(Idx).Y-30   ' Position to roll into Kicker
    Else
       aBall(Idx).VelX=aBall(Idx).VelX*.95
       aBall(Idx).VelY=aBall(Idx).VelY*.95
       If (aTicks(Idx) > 17) AND (aBall(Idx).Y > (Kickers(Idx).Y+15)) Then
          KickerTrigs_UnHit(Idx)      ' below center & too long, let ball go
       End If
    End If
End Sub
'
'  Kicker 70 Unit Diameter Zone exited without
'  being grabbed by Kicker. Stop Attraction Timer
'
Sub KickerTrigs_UnHit(Idx)
    KickerTrigs(Idx).TimerEnabled=False   ' Stop Timer
    Kickers(Idx).Enabled=False        ' Assure Kicker Disabled
    AddSound "Roll9"
End Sub
'
'----------------------------------------
'      Actual Kicker Hole Entered
'----------------------------------------
'
Sub Kickers_Hit(Idx)
    Dim I, OB
    KickerTrigs(Idx).TimerEnabled=False ' Stop Attraction Timer
    If KickersLit(Idx).IsDropped=False Then
       Kickers(Idx).UserValue=1       ' Say Kicker Is Lit
       KickersLit(Idx).IsDropped=True   ' Lower Lit Cover
        If (Tilted=False) Then
           AddReplay()            ' Special When Lit
        End If
    Else
       Kickers(Idx).UserValue=0       ' Say Kicker Not Lit
       KickerCovers(Idx).IsDropped=True   ' Lower Lit Cover
    End If
    Kickers(Idx).TimerInterval = 560    ' Set delay to kickout
    Kickers(Idx).TimerEnabled = True    ' Start Kickout Timer
    PlaySound "motor"
    If Tilted=False Then
        AddScore 30, 0            ' 300,000 Points
    End If
End Sub
'
'   Kicker Kickout Timer
'
Sub Kickers_Timer(Idx)
    Dim Ang, S
    PlaySound "Kicker"
    Kickers(Idx).TimerEnabled = False   ' Stop Timer
     If (Kickers(Idx).UserValue=0) Then
        KickerCovers(Idx).IsDropped=False ' Raise UnLit Cover
     Else
       KickersLit(Idx).IsDropped=False    ' Raise Lit Cover
     End If
    '
    '  Kicker kickout angles are:
    '
    '  Kicker 1 40 Degrees
    '  Kicker 2 320
    '
    '  Note: Idx = Kicker # - 1
    '
    S=RND()*1
    Select Case Idx
       Case 0
          Ang=39              ' Kicker 1 kicks at 40 Degrees
        Case 1
          Ang=320             ' Kicker 2 kicks at 320 Degrees
          S=0
    End Select
    Kickers(Idx).Kick (Ang+1)-S, 9.2    ' Kick out the Ball +/- 1 degrees
    Kickers(Idx).Enabled=False        ' Disable kicker
End Sub
'
'-----------------------------------------------------
'  Attraction Routine For  Hole Enlargement Zone
'
'  Math taken from core.vbs Magnet routine
'-----------------------------------------------------
'
' On Entry:
'
'        aBall = Pointer to Ball Object being attracted
'            X = Hole Center X
'            Y = Hole Center Y
'         Size = Diameter of Enlarged Area
'     Strength = Attraction Strength
'
Function AttractBall(aBall,X,Y,Size,Strength)
    Dim dX, dY, dist, force, ratio
    '
    dX = aBall.X - X
    dY = aBall.Y - Y
    dist = Sqr(dX*dX + dY*dY)
    AttractBall=dist        ' Return Distance to center
    If dist > Size Or dist < 1 Then
       Exit Function        'Just to be safe
    End If
    '
    ' Exert Centripital Force on Ball
    '
    ratio = dist / (1.5 * Size)
    force = Strength * EXP(-0.2/ratio)/(ratio*ratio*56) * 1.5
    aBall.VelX = (aBall.VelX - dX * force / dist) * 0.985
    aBall.VelY = (aBall.VelY - dY * force / dist) * 0.985
    AttractBall=dist
End Function
'
'******************************************************
'     Handlers for Rollover Lane Triggers Hit
'******************************************************
'
'  Idx     Lane
' -----   ------------------
'   0     Bottom "A" Lane (300,000)
'   1     Bottom "B" Lane (300,000)
'   2     Bottom "C" Lane (300,000)
'   3     Bottom "D" Lane (300,000)
'
Sub Lanes_Hit(Idx)
    Dim I
  Wires(Idx).IsDropped=False      ' Animate Rollover Wire
    If Tilted=True Then
       Exit Sub             ' Done if tilted
    End If
  AddScore 30, 0            ' score 300,000
    PlaySoundAtVol "BallRoll1", ActiveBall, 1       ' Outlane Sound
    Select Case Idx
       Case 0
        If (LtSpecial.IsDropped=False) AND (LtA.IsDropped=False) Then
           AddReplay()        ' Special When Lit
        End If
          If (LtSpecial.IsDropped=True) Then
             LtA.IsDropped=True     ' UnLight "A" Lane
          End If
          ButtonOn(1)         ' Light Rollover Button
          LtTarg1.IsDropped=False   ' Light Green Target  (Target1)
       Case 1
        If (LtSpecial.IsDropped=False) AND (LtB.IsDropped=False) Then
           AddReplay()        ' Special When Lit
        End If
          If (LtSpecial.IsDropped=True) Then
             LtB.IsDropped=True     ' UnLight "B" Lane
          End If
          ButtonOn(2)         ' Light Rollover Button
          LtTarg2.IsDropped=False   ' Light White Target  (Target 2)
       Case 2
        If (LtSpecial.IsDropped=False) AND (LtC.IsDropped=False) Then
           AddReplay()        ' Special When Lit
        End If
          If (LtSpecial.IsDropped=True) Then
             LtC.IsDropped=True     ' UnLight "C" Lane
          End If
          ButtonOn(3)         ' Light Rollover Button
          LtTarg3.IsDropped=False   ' Light Red Target   (Target 3)
       Case 3
        If (LtSpecial.IsDropped=False) AND (LtD.IsDropped=False) Then
           AddReplay()        ' Special When Lit
        End If
          If (LtSpecial.IsDropped=True) Then
             LtD.IsDropped=True     ' UnLight "D" Lane
          End If
          ButtonOn(4)         ' Light Rollover Button
          LtTarg4.IsDropped=False   ' Light Yellow Target (Target 4)
    End Select
     If MadeSpecial=False Then
       For I=0 to (LtLet.Count-1)
           If LtLet(I).IsDropped=False Then
              Exit Sub
           End If
        Next
        MadeSpecial=True          ' A-B-C-D made
        LtSpecial.IsDropped=False   ' Light One Letter Special
        LightsRotate()          ' Light Letter that is special now
     End If
 End Sub
'
'  Check A-B-C Specials
'
Sub CheckSpecials()
End Sub
'
'------------------------
'  Rollover exit, start animation timer
'-----------------------------------------
'
Sub Lanes_UnHit(Idx)
    Wires(Idx).TimerInterval=75     ' set slight delay...
  Wires(Idx).TimerEnabled=True    ' start animation timer
End Sub
'
'  ----------------------
'  Wire Animation Timers.
'  ----------------------
'
'  All they Do is end the animation.
'
Sub Wires_Timer(Idx)
  Wires(Idx).IsDropped=True     ' Drop animation Wall
  Wires(Idx).TimerEnabled=False   ' Stop Timer
End Sub
'
' ---------------------------------------
' Do Cycle of 1 Point Buttons/Bumpers
' ---------------------------------------
'
'  If lit number special, cycle it
'
Sub LightsRotate()
     If MadeSpecial=True Then
        LtLet(Rotate).IsDropped=True    ' UnLight Current Special Letter
     End If
     Rotate=Rotate+1            ' Bump Special Letter rotator
     If Rotate > 3 Then
        Rotate=0
     End If
     If MadeSpecial=True Then
        LtLet(Rotate).IsDropped=False ' Light Next Special Letter
     End If
     Select Case (Score MOD 10)
        Case 2, 5, 7, 9
           BumperOn(7)          ' Light 100,00 When Lit Bumpers
           BumperOn(9)
           klight.state=1
        Case Else
           BumperOff(7)         ' UnLight 100,000 When Lit Bumpers
           BumperOff(9)
           klight.state=0
     End Select
End Sub
'
' -------------------------------------------
'  Routines to Turn Bumper Lights On and Off
' -------------------------------------------
'
Sub BumpersOn()
  For X=1 to 10
       BumperOn(X)
  Next
End Sub
'
Sub BumpersOff()
  For X=1 to 10
     BumperOff(X)
  Next
End Sub
'
'  Turn Single Bumper Off
'
Sub BumperOff(Num)
' EVAL("B"&Num&"Base").State=LightStateOff
  EVAL("B"&Num&"SideOn").IsDropped=True
  EVAL("B"&Num&"SideOff").IsDropped=False
' EVAL("B"&Num&"Top").State=LightStateOff
  Select Case Num
       Case 1, 3, 4, 5, 6, 8, 10
      EVAL("B"&Num&"RimOn").IsDropped=True
      EVAL("B"&Num&"RimOn1").IsDropped=True
        EVAL("B"&Num&"RimOn2").IsDropped=True
       Case 2, 7, 9
          EVAL("B"&Num&"Top1Lit").IsDropped=True
          EVAL("B"&Num&"Top2Lit").IsDropped=True
          EVAL("B"&Num&"Top3Lit").IsDropped=True
  End Select
End Sub
'
' Turn Single Bumper On
'
Sub BumperOn(Num)
' EVAL("B"&Num&"Base").State=LightStateOn
  EVAL("B"&Num&"SideOn").IsDropped=False
  EVAL("B"&Num&"SideOff").IsDropped=True
' EVAL("B"&Num&"Top").State=LightStateOn
    Select Case Num
       Case 1, 3, 4, 5, 6, 8, 10
      EVAL("B"&Num&"RimOn").IsDropped=False
      EVAL("B"&Num&"RimOn1").IsDropped=False
      EVAL("B"&Num&"RimOn2").IsDropped=False
       Case 2, 7, 9
          EVAL("B"&Num&"Top1Lit").IsDropped=False
          EVAL("B"&Num&"Top2Lit").IsDropped=False
          EVAL("B"&Num&"Top3Lit").IsDropped=False
    End Select
End Sub
 '
'***************************************************
'   Subroutine to Add 2-10 Points (even count only)
'***************************************************
'
Sub AddPoints(Num)
    Add2Point()           ' Score 2 Points
    If Num > 2 Then
       PtTimer.UserValue = Num-2    ' Rest of points on motor timing
       PtTimer.Enabled=True       ' Start Point scoring motor
    End If
End Sub
'
'  Point Scoring Motor
'
Sub PtTimer_Timer()
    Add2Point()             ' Add 2 Points
    PtTimer.UserValue = PtTimer.UserValue-2
    If PtTimer.UserValue <= 0 Then
       PtTimer.Enabled=False      ' Done, Stop Motor
    End If
End Sub
'
'**************************************
'  Subroutine to Add 2 Points
'**************************************
'
'  Replay for 50 Points, 70 Points, 80 Points,
'             90 Points, and 98 Points.
'
Sub Add2Point()
  If Points >= 98 Then
     Exit Sub             ' 98 is maximum for Points
  End If
  Points=Points+2           ' Update Point Score
  PlaySound "1ptBell"
  Select Case Points
     Case 50, 70, 80, 90, 98
      AddReplay()         ' Award Point Replays When Made
  End Select
  DisplayPoints()           ' Display Points
End Sub
'--------------------------------------------
' Routine to avoid VP9 Audio "growl" from
' too many overlapping PlaySounds.  Call
' like it was Playsound.
'--------------------------------------------
'
Sub AddSound(Snd)
  If GrowlTimer.Uservalue=True Then
     Exit Sub             ' Too Soon for more
  End If
  PlaySound Snd
  GrowlTimer.UserValue=True
  GrowlTimer.Enabled=True
End Sub
'
Sub GrowlTimer_Timer()
  GrowlTimer.UserValue=False      ' Allow Another AddSound
  GrowlTimer.Enabled=False      ' Stop Timer
End Sub
 '
'---------------------------------
'      Reset Lights to Zero
'---------------------------------
'
Sub ResetLightsToZero()
  TH100sReel.TimerInterval=180    ' Step at 180ms/step
  TH100sReel.TimerEnabled=True    ' Start the Reset
End Sub
'
'  Timer does the reset 1 tick at a time
'
Sub TH100SReel_Timer()
  X=True                ' Check for Score at Zero
  If SReelUnits <> 0 Then
     SReelUnits=SreelUnits-1      ' Step 10,000 Lights Down 1
     X=False              ' Don't Stop This Cycle
  End If
  If SReelTens <> 0 Then        ' Step 100,000 Lights Down 1
     SReelTens=SReelTens-1
     X=False              ' Don't Stop This Cycle
  End If
  If SReelHuns <> 0 Then
     SReelHuns=SReelHuns-1      ' Step Millions Lights Down 1
     X=False              ' Don't Stop This Cycle
  End If
  If X=True Then
     TH100sReel.TimerEnabled=False  ' Stop Reset, We're done
  End If
  DisplayScore()            ' Update Lights w/new value
End Sub
'
'============================================================
'   B A C K G L A S S   D I S P L A Y   R O U T I N E S
'============================================================
'
'***********************************************
'  Subroutine to Display GAME OVER on Backglass
'***********************************************
'
' There is no game over on this game
'
Sub DisplayGOV(State)
End Sub
 '
'***********************************************
'  Subroutine to Display TILT on Backglass
'***********************************************
'
' State = 1 to display TILT, 0 to clear TILT display
'
Sub DisplayTilt(State)
  TiltReel.Setvalue State
    Controller.B2SSetTilt 33,State
End Sub
'
'
'***********************************************
'  Subroutine to Display Credits on Backglass
'***********************************************
'
Sub DisplayCredits()
  CreditReel.Setvalue Credits
  CreditReel2.Setvalue Credits
    Controller.B2SSetCredits Credits
End Sub
'
'***********************************************
'  Subroutine to Display Points on Backglass
'***********************************************
'
 '  Note: Points will always be even as the backglass only
 '        display 2, 4, 6, & 8 because points score 2 at a time.
 '
Sub DisplayPoints()
    Pointslights
    Dim I
     '
     '  Display Points 10 digit (OFF or 10, 20, etc...)
     '
  I = INT((Points MOD 100) /10)   ' Get Points Tens Value (0 - 9)
  Pts10ReelM.SetValue I       ' Display Proper 10s value on Mini Glass
    If (I=1) OR (I=2) Then
       Pts10Reel.SetValue I       ' 10 & 20 display on Large Glass
    Else
       Pts10Reel.SetValue 0       ' The Rest Not on Large Glass
    End If
    '
     '  Display Points Units Digit (OFF or 2, 4, 6, 8)
     '
  I = (Points MOD 10)         ' Get Points Units Value 0 to 9
     I = INT(I/2)           ' 0&1=0, 2&3=1, 4&5=2, 6&7=3, 8&9=4
  Pts1Reel.Setvalue I         ' Display on Large Glass
  Pts1ReelM.Setvalue I        ' Display on Mini Glass too
End Sub

Sub Pointslights()
          Controller.B2SSetData 2, 0
          Controller.B2SSetData 4, 0
          Controller.B2SSetData 6, 0
          Controller.B2SSetData 8, 0
          Controller.B2SSetData 10, 0
          Controller.B2SSetData 20, 0
          Controller.B2SSetData 30, 0
          Controller.B2SSetData 40, 0
          Controller.B2SSetData 50, 0
          Controller.B2SSetData 60, 0
          Controller.B2SSetData 70, 0
          Controller.B2SSetData 80, 0
          Controller.B2SSetData 90, 0
Select Case Points
Case 2:   Controller.B2SSetData 2, 1
Case 4:   Controller.B2SSetData 4, 1
Case 6:   Controller.B2SSetData 6, 1
Case 8:   Controller.B2SSetData 8, 1
Case 10:   Controller.B2SSetData 10, 1
Case 12:   Controller.B2SSetData 10, 1
           Controller.B2SSetData 2, 1
Case 14:   Controller.B2SSetData 10, 1
           Controller.B2SSetData 4, 1
Case 16:   Controller.B2SSetData 10, 1
           Controller.B2SSetData 6, 1
Case 18:   Controller.B2SSetData 10, 1
           Controller.B2SSetData 8, 1
Case 20:   Controller.B2SSetData 20, 1
Case 22:   Controller.B2SSetData 20, 1
           Controller.B2SSetData 2, 1
Case 24:   Controller.B2SSetData 20, 1
           Controller.B2SSetData 4, 1
Case 26:   Controller.B2SSetData 20, 1
           Controller.B2SSetData 6, 1
Case 28:   Controller.B2SSetData 20, 1
           Controller.B2SSetData 8, 1
Case 30:   Controller.B2SSetData 30, 1
Case 32:   Controller.B2SSetData 30, 1
           Controller.B2SSetData 2, 1
Case 34:   Controller.B2SSetData 30, 1
           Controller.B2SSetData 4, 1
Case 36:   Controller.B2SSetData 30, 1
           Controller.B2SSetData 6, 1
Case 38:   Controller.B2SSetData 30, 1
           Controller.B2SSetData 8, 1
Case 40:   Controller.B2SSetData 40, 1
Case 42:   Controller.B2SSetData 40, 1
           Controller.B2SSetData 2, 1
Case 44:   Controller.B2SSetData 40, 1
           Controller.B2SSetData 4, 1
Case 46:   Controller.B2SSetData 40, 1
           Controller.B2SSetData 6, 1
Case 48:   Controller.B2SSetData 40, 1
           Controller.B2SSetData 8, 1
Case 50:   Controller.B2SSetData 50, 1
Case 52:   Controller.B2SSetData 50, 1
           Controller.B2SSetData 2, 1
Case 54:   Controller.B2SSetData 50, 1
           Controller.B2SSetData 4, 1
Case 56:   Controller.B2SSetData 50, 1
           Controller.B2SSetData 6, 1
Case 58:   Controller.B2SSetData 50, 1
           Controller.B2SSetData 8, 1
Case 60:   Controller.B2SSetData 60, 1
Case 62:   Controller.B2SSetData 60, 1
           Controller.B2SSetData 2, 1
Case 64:   Controller.B2SSetData 60, 1
           Controller.B2SSetData 4, 1
Case 66:   Controller.B2SSetData 60, 1
           Controller.B2SSetData 6, 1
Case 68:   Controller.B2SSetData 60, 1
           Controller.B2SSetData 8, 1
Case 70:   Controller.B2SSetData 70, 1
Case 72:   Controller.B2SSetData 70, 1
           Controller.B2SSetData 2, 1
Case 74:   Controller.B2SSetData 70, 1
           Controller.B2SSetData 4, 1
Case 76:   Controller.B2SSetData 70, 1
           Controller.B2SSetData 6, 1
Case 78:   Controller.B2SSetData 70, 1
           Controller.B2SSetData 8, 1
Case 80:   Controller.B2SSetData 80, 1
Case 82:   Controller.B2SSetData 80, 1
           Controller.B2SSetData 2, 1
Case 84:   Controller.B2SSetData 80, 1
           Controller.B2SSetData 4, 1
Case 86:   Controller.B2SSetData 80, 1
           Controller.B2SSetData 6, 1
Case 88:   Controller.B2SSetData 80, 1
           Controller.B2SSetData 8, 1
Case 90:   Controller.B2SSetData 90, 1
Case 92:   Controller.B2SSetData 90, 1
           Controller.B2SSetData 2, 1
Case 94:   Controller.B2SSetData 90, 1
           Controller.B2SSetData 4, 1
Case 96:   Controller.B2SSetData 90, 1
           Controller.B2SSetData 6, 1
Case 98:   Controller.B2SSetData 90, 1
           Controller.B2SSetData 8, 1
 End Select
 End Sub

'
'***********************************************
'  Subroutine to Display Score on Backglass
'***********************************************
'
'   SReelHuns = Millions Digit
'   SReelTens = 100 Thousands Digit
'  SReelUnits = 10 Thousands Digit
'
Sub DisplayScore()
  Showscore
  ' Do Millions Display
  '
  Select Case SReelHuns
       Case 1, 2, 3, 4
        M1sReelA.SetValue SReelHuns ' Numbers 1-4 On Left Side
        M1sReelB.SetValue 0     ' Not Right Side
       Case 5, 6
        M1sReelA.SetValue SReelHuns ' Numbers 5 & 6 On Both Sides
        M1sReelB.SetValue SReelHuns-4
     Case 7
        M1sReelA.SetValue 0     ' On Left Side Now
        M1sReelB.SetValue SReelHuns-4 ' Light 7
       Case Else
          M1sReelA.SetValue 0
          M1sReelB.SetValue 0     ' 0 = off, > 7 = off
  End Select
  '
  ' Do Hundred Thousands Display
  '
  TH100sReel.SetValue SReelTens   ' 100-900k on Left Side
  '
  ' Do Ten Thousands Display
  '
    Select Case SReelUnits
       Case 1
          TH10sReelA.SetValue 1       ' Light 10,000
          TH10sReelB.SetValue 0
       Case 2
          TH10sReelA.SetValue 0
          TH10sReelB.Setvalue 1       ' Light 20,000
       Case 3
          Th10sReelA.SetValue 2       ' Light 30,000
          TH10sReelB.SetValue 0
       Case 4
          TH10sReelA.SetValue 0
          TH10sReelB.SetValue 2       ' Light 40,000
        Case 5
          TH10sReelA.SetValue 3       ' Light 50,000
          TH10sReelB.SetValue 0
        Case 6, 7, 8, 9
          TH10sReelA.SetValue 0
          TH10sReelB.SetValue SReelUnits-3  ' Light 60,000 - 90,000
       Case Else
          TH10sReelA.SetValue 0       ' 0
          TH10sReelB.SetValue 0
    End Select
     If SReelUnits=1 Then
        TH10sReelM.SetValue 1       ' 10,000 also on MiniGlass
     Else
        TH10sReelM.SetValue 0       ' No Others there
     End If
End Sub

 Sub Showscore()
 Select Case SReelHuns  'Millions

       Case 0:
          Controller.B2SSetData 71, 0
          Controller.B2SSetData 72, 0
          Controller.B2SSetData 73, 0
          Controller.B2SSetData 74, 0
          Controller.B2SSetData 75, 0
          Controller.B2SSetData 76, 0
          Controller.B2SSetData 77, 0
          Controller.B2SSetData 78, 0
          Controller.B2SSetData 79, 0
       Case 1:
          Controller.B2SSetData 71, 1
       Case 2:
          Controller.B2SSetData 71, 0
          Controller.B2SSetData 72, 1
       Case 3:
          Controller.B2SSetData 72, 0
          Controller.B2SSetData 73, 1
       Case 4:
          Controller.B2SSetData 73, 0
          Controller.B2SSetData 74, 1
       Case 5:
          Controller.B2SSetData 74, 0
          Controller.B2SSetData 75, 1
       Case 6:
          Controller.B2SSetData 75, 0
          Controller.B2SSetData 76, 1
       Case 7:
          Controller.B2SSetData 76, 0
          Controller.B2SSetData 77, 1
 End Select

 Select Case SReelTens  'Hundred Thousands
       Case 0:
          Controller.B2SSetData 61, 0
          Controller.B2SSetData 62, 0
          Controller.B2SSetData 63, 0
          Controller.B2SSetData 64, 0
          Controller.B2SSetData 65, 0
          Controller.B2SSetData 66, 0
          Controller.B2SSetData 67, 0
          Controller.B2SSetData 68, 0
          Controller.B2SSetData 69, 0

       Case 1:
          Controller.B2SSetData 61, 1
       Case 2:
          Controller.B2SSetData 61, 0
          Controller.B2SSetData 62, 1
       Case 3:
          Controller.B2SSetData 62, 0
          Controller.B2SSetData 63, 1
       Case 4:
          Controller.B2SSetData 63, 0
          Controller.B2SSetData 64, 1
       Case 5:
          Controller.B2SSetData 64, 0
          Controller.B2SSetData 65, 1
       Case 6:
          Controller.B2SSetData 65, 0
          Controller.B2SSetData 66, 1
       Case 7:
          Controller.B2SSetData 66, 0
          Controller.B2SSetData 67, 1
       Case 8:
          Controller.B2SSetData 67, 0
          Controller.B2SSetData 68, 1
       Case 9:
          Controller.B2SSetData 68, 0
          Controller.B2SSetData 69, 1
 End Select

 Select Case SReelUnits  'Ten Thousands
       Case 0:
          Controller.B2SSetData 51, 0
          Controller.B2SSetData 52, 0
          Controller.B2SSetData 53, 0
          Controller.B2SSetData 54, 0
          Controller.B2SSetData 55, 0
          Controller.B2SSetData 56, 0
          Controller.B2SSetData 57, 0
          Controller.B2SSetData 58, 0
          Controller.B2SSetData 59, 0

       Case 1:
          Controller.B2SSetData 51, 1
       Case 2:
          Controller.B2SSetData 51, 0
          Controller.B2SSetData 52, 1
       Case 3:
          Controller.B2SSetData 52, 0
          Controller.B2SSetData 53, 1
:      Case 4:
          Controller.B2SSetData 53, 0
          Controller.B2SSetData 54, 1
       Case 5:
          Controller.B2SSetData 54, 0
          Controller.B2SSetData 55, 1
       Case 6:
          Controller.B2SSetData 55, 0
          Controller.B2SSetData 56, 1
       Case 7:
          Controller.B2SSetData 56, 0
          Controller.B2SSetData 57, 1
       Case 8:
          Controller.B2SSetData 57, 0
          Controller.B2SSetData 58, 1
       Case 9:
          Controller.B2SSetData 58, 0
          Controller.B2SSetData 59, 1
 End Select

 End Sub

'-----------------------------------------------------------


' Thalamus : Exit in a clean and proper way
Sub Table_exit
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
' Thalamus, AudioFade - Patched
  If tmp > 0 Then
    AudioFade = Csng(tmp ^5) 'was 10
  Else
    AudioFade = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
' Thalamus, AudioPan - Patched
  If tmp > 0 Then
    AudioPan = Csng(tmp ^5) 'was 10
  Else
    AudioPan = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
' Thalamus, Pan - Patched
  If tmp > 0 Then
    Pan = Csng(tmp ^5) 'was 10
  Else
    Pan = Csng(-((- tmp) ^5) ) ' was 10
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


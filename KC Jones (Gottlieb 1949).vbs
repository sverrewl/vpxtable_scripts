' ********   K. C. Jones    *************
'          Gottlieb 1949
'
'  Version 1.0 PLB February 2014
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
Option Explicit     'Force variables to be declared

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
Dim Replay6
Dim Replay1Paid     'Mark when each replay has been paid
Dim Replay2Paid
Dim Replay3Paid
Dim Replay4Paid
Dim Replay5Paid
Dim Replay6Paid
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
Dim Bell5ToDo           ' Flag for 5 Bell de-overlap timer
Dim Bell50ToDo          ' Flag for 50 point Bell de-overlap Timer
Dim MQ(30)        ' De-Overlap up to 29 scoring motor events
            ' (We should never get more than 2 or 3 deep.)
            '
            ' MotorQueue Structure:
            ' MQ(0) = Index of Entry motor is now executing (0=none)
            ' MQ(1) = index (2-20) to queue next event
            '   Note: If MQ(0)=0 motor is stopped.
            '         Any other value, and it is running.
            '   Motor Stops if current event finishes and
            '   MQ(1)=MQ(0)+1.  It resets MQ(0)=0 and MQ(1)=2
            '   and stops running.
Dim MX          ' Scoring Motor Temp variable
Dim MSkip       ' > 0 = Empty Motor Cycles
'
'
Dim B8State             ' B8 Animation State
Dim B9State             ' B9 Animation State
Dim PW          ' Plunger Animation State
Dim ShootBall           ' Is animation Pulling, or Releasing?
Dim PS                  ' Pull Speed
Dim DrainWallHit    ' Limits to one it sound per ball on Drain Wall
Dim TrayHit       ' Limits to one sound per ball on Tray Hit
Dim BallID        ' Ball ID number
Dim BallDrained(25)   ' List of IDs of drained balls (index 0 = # entries in list)
Dim TroughOpening   ' True = Opening, False = Closing
Dim aBall       ' Gobble/Kicker Ball Pointer
Dim aXpos
Dim aXadd
Dim aYpos
Dim aYadd
Dim aZpos       ' Gobble Hole Drop Vertical Position
'
Dim LRLast              ' Units Digit after Previous AddScore()
'
Dim MadeSpecial     ' Special Has Been Made
Dim MadeCount     ' Count of Times 1-7 has been made
Dim Spotted(16)     ' Array of which numbers are spotted
Dim RHPointer     ' 1 to 5 for RoundHouse Kicker Score Value
Dim RHExtra       ' Cycle Counter for Kicker Scoring
Dim GateCount     ' # of Ball through the inlane Gate
'
'  Z state is a special "5th ball CatchUp" mode that is
'  armed when the score is between 800,000 and 2,900,000 and the
'  5th ball is on the playfield.
'
'  Z State lights all of the extra value lights, bumpers, and contacts
'  and the kicker.
'
'  Z state ends when the motor runs (500K bumper or lower center
'  50K rollover) and the score is now over 2,900,000.
'
'  If the ball enters the lit kicker, the score is advanced by the
'  current roundhouse value until the score is > 2,900,000 at
'  which time Z state ends.
'
Dim ZState        ' 0= Not in Z State, 1= In ZState
 Dim InKicker     ' False if Ball Not in Kicker, True If Ball In Kicker
Dim MotorValue      ' Last Score Value Scoring Motor Ran For
'
 Dim Controller
 LoadController
 Sub LoadController()
 Set Controller = CreateObject("B2S.Server")
 Controller.B2SName = "KC Jones"  ' *****!!!!This must match the name of your directb2s file!!!!
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
  TableName="KCJones" 'Place your table name here
  CreditsPerCoin=1  '1 to 5 credits per coin
    MaxCredits = 99     'Size of Credit Reel Graphic
  HighScoreReward=1 'The number of replays to award when HIGH SCORE is beaten (1-5)
'
    TiltTimer.Enabled=True
    TiltSensitivity=4 '0-15, 4=normal - a higher number is less sensitive, 0=1 nudge and you're out
  Bells=True      'TRUE to play bell sounds, FALSE for no bells
  ReelSpeed=3     'Defines the speed of the reels, 1=slow, 9=fast, 3=default
'
  Replay1=400     'Replays Start 4,000,000
  Replay2=430     ' 4,300,000
    Replay3=450     ' 4,500,000
    Replay4=480     ' 4,800,000 & Each 100k Until...
    Replay5=600     ' Replay/100k ends at 6,000,000
'
  Initialize()    'Call the main Init sub
End Sub
'
'----------------------------------------
'  Selected Ball Hits - Sound Only
'----------------------------------------
'
Sub Rubbers_Hit(Idx)
    AddSound "Rubber"
End Sub
'
Sub Guides_Hit(Idx)
    AddSound "Metal"
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
    '
     ' Disable/Enable Bumpers
    '
     For X=1 to 11
  '      EVAL("B"&X&"Top").Disabled=State
  '      EVAL("B"&X&"Base").Disabled=State
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
  Randomize                 'Seed random Number Generator
'
    BuzzCount=0
    Count=0                   'Set variables to a known state
  Count1=0
  GOVPhase=0
  GSPhase=0
    MQ(0)=0                   'Init Scoring Motor Queue Empty
    MQ(1)=2                   'Next Even Goes here
    MSkip=0                   'No Skip Cycles Yet
  BallsLost=0                 'Reset ball counter
  BallsLeft=0                 'The number of balls not yet lifted
    BallsOnField=0                'No Balls Lifted to Playfield
  InProgress=False              'Game not started
    GrowlTimer.userValue=False          ' Init AddSound to "Can Add"
    '
  DisplayHighScore=True           'Display the high score
    '
    '  Init Ball Return Trough
    '
    TroughOpen.IsDropped=True     ' Trough inits to Closed
    TroughOpenHalf.IsDropped=True   '  ... Fully Closed
    TroughOpening=False         ' Say Closing was last action
    DrainTrig.UserValue=0       ' Init debouncer
'
'  Set Initial Light States:
'
'
'   All 11 Bumpers and 2 Switch Covers init to Off
'
'  Note: Bumpers 1, 2, 3, 4, 5, 10 & 11 are passive,
 '        "Bumpers" 6 & 7 are the side Contact Switches
'        Bumpers 8 & 9 are POP Bumpers
'        "Bumpers" 12 & 13 are the lower Contact Switch Covers
'
    For X=1 to 11
       BumperOff(X)
       If (X < 8) OR (X > 9) Then
          ' Set Not Debouncing on Passive Bumper
          Eval("B"&X&"Top").UserValue=False
       Else
          '
          ' B8 & B9 are POP bumpers
          '
          Eval("B"&X&"Ring1").IsDropped=True
          Eval("B"&X&"Ring2").IsDropped=True
          Eval("B"&X&"Ring3").IsDropped=True
          Eval("B"&X&"Ring3").TimerInterval=10  ' animation Speed
       End If
    Next
    B8State=0           ' Init POP Ring animation States
    B9State=0
'
'  Init Two Lower Switch covers to Unlit
'
    BumperOff(12)
    BumperOff(13)
    SW12Base.UserValue=False      ' Init Debouncers
    SW13Base.userValue=False
'
'  Special Lights Off
'
    LtExtraSpecial.IsDropped=True
    LtSemiGrn.IsDropped=True
    LtSemiRed.IsDropped=True
    LtSpecial1.IsDropped=True
    LtSpecial2.IsDropped=True
    LtSpecial5.IsDropped=True
'
'  Score Multipliers Off
'
    LtHeadlight.IsDropped=True
'
'   Init Extra Score to All Of
'
    For X=1 to 5
       EVAL("LtES"&X).IsDropped=True
    Next
    RHPointer=1             ' Init Round House Score Value
'
'   Init Exit Lane Trigger Wire and kicker in "off" position
'
    LaneESWire.IsDropped=True     ' Kicker Stowed
    LaneESWire2.IsDropped=True      ' Sensor Wire not Pressed

'
'  Kicker Not Lit
'
    CtrKickerCoverLit.IsDropped=True
     CtrKickerCoverOpen.IsDropped=True
    CtrKicker.UserValue=0       ' Say Center unlit
     InKicker=False           ' Init Ball Not In Kicker
'
'   Init Plunger animation to Fully Extended
'   Note: Plunger Walls Work upside Down from normal
'
    For X=2 To 12
       EVAL("PW"&X).isDropped=False   ' Drop all but PW1
    Next
    PW1.IsDropped=True              ' Raise PW1
    PW=0                            ' Say at Phase 0
    RollFlag=0                          'No Ball Rolling Yet
    RollVary=0                          'Init Roll Sound Variation flag
    '
    '   Match Animated Plunger Speed to real Plunger
    '      Note: Following Calc is for VP 8 w/ Plunger Stroke=110
    '
    PS=Plunger.PullSpeed
    If PS > 4 Then
       PS = 4             ' Max Animation Speed
    Else
       PS = ((5-PS)*8)+1        ' Match animation To plunger
    End If
    PW1.TimerInterval=PS              ' Set Plunger Speed
    ShootBall=0                       ' Next Move is Pull
    '
    '   Create Balls in Trough
    '
    BIT=0               ' No Balls in Lifter Tray
    BallID=1              ' reset Ball ID
    BallDrained(0)=0          ' None drained
    For X=1 to 5
       EVAL("Load"&X).CreateBall.UserValue=X  ' Create Ball and ID it
     EVAL("Load"&X).Kick 0, 0         ' Release Ball to Trough
       EVAL("Load"&X).Enabled=False       ' Disable Load Kicker
    Next
    BallsLoaded=5           ' Say all balls loaded
    BallsLost=5             ' Say Rolled in from Playfield
    '
    '  5 Balls in the visible trough
    '
    FlippersTimer.Enabled=True        ' Start VP9 Primitive Timer
     '
  On Error Resume Next    'Needed if highscore/credits have never been saved in this table name
  Credits=Cdbl(LoadValue(TableName,"Credits"))
  If Credits="" Then Credits=1          'If there was no entry, then set to default value
    DisplayCredits()              ' Display Credits on Back Glass
  HighScore=Cdbl(LoadValue(TableName,"HighScore"))
  If HighScore="" Then HighScore=100      'If there was no entry, then set to default value
  Score=Cdbl(LoadValue(TableName,"Score"))  'Load score value from previous game
    If DisplayHighScore=True Then
       SetHS(HighScore)             ' Put On Display
    End If
  If Score="" Then Score=455          ' If there was no entry, then use default value
    SReelHuns=  INT((Score MOD 1000) / 100)   ' Init Scoring Reel Digits
    SReelTens = INT((Score MOD 100) /10)
    SReelUnits = (Score MOD 10)
    DisplayScore()                ' Display Scoring Lights
    MUnits=0                  ' Init Scoring Motor to stopped
    MTens=0
'
End Sub
'
' *********************************
' ***** Keys Pressed Handling *****
' *********************************
Sub Table_KeyDown(ByVal Keycode)            'The usual keycode stuff

    If Keycode=LeftFlipperKey Then
      LeftFlipper.RotateToEnd
      PlaySoundatvol "FlipperUp", LeftFlipper, 1
            If BuzzCount=0 Then
                PlayLoopSoundAtVol "Buzz", LeftFlipper, 1          ' Start Flipper Solenoid Buzz
            End If
            BuzzCount=BuzzCount+1
      End If


    If Keycode=RightFlipperKey Then
      RightFlipper.RotateToEnd
      PlaySoundatvol "FlipperUp", RightFlipper, 1
            If BuzzCount=0 Then
                PlayLoopSoundAtVol "Buzz", RightFlipper, 1          ' Start Flipper Solenoid Buzz
            End If
            BuzzCount=BuzzCount+1
      End If

    If Keycode=LeftTiltKey Then
      Nudge 270,0.75
      CheckTilt()
    End If

    If Keycode=RightTiltKey Then
      Nudge 90,0.75
      CheckTilt()
    End If

    If Keycode=CenterTiltKey Then
      Nudge 0,0.75
      CheckTilt()
    End If
' Thalamus - added mechanicaltilt
    If Keycode=MechanicalTilt Then
      Nudge 0,0.75
      CheckTilt()
    End If

    '
    ' End flipper and Nudge
    '

  If Keycode=PlungerKey Then
       ShootBall=False              ' Indicate Animating Pull
       MovePlunge()                 ' Start video image of plunger moving
       Plunger.PullBack
    End If

  If Keycode=AddCreditKey Then
       PlaySound"Coin"
     CreditReel.TimerEnabled=True
       Showcredits
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

    ' "2" = Reset High Score to 1,000,000
    If KeyCode=3 Then
        HighScore = 100
        SaveValue TableName,"HighScore",HighScore   'Save score value for next time
        If DisplayHighScore=True Then
           SetHS(HighScore)
        End If
    End If

    ' "3" = Reset Credits
    If KeyCode=4 Then
       Credits=0
       DisplayCredits()             ' Display Credits on Back Glass
     SaveValue TableName,"Credits",Credits  'Save the credits
    End If

     If KeyCode=5 Then
        If B6Top1Lit.IsDropped=False Then
           For X=1 to 13
              BumperOff(X)
           Next
        Else
           For X=1 to 13
              BumperOn(X)
           Next
        End If
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
    RF.RotAndTra2 = RightFlipper.CurrentAngle
    LF.RotAndTra2 = LeftFlipper.CurrentAngle
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
      PlaySoundatvol "FlipperDown", LeftFlipper, 1
            BuzzCount=BuzzCount-1
            If BuzzCount <= 0 Then
               StopSound "Buzz"           ' End Flipper Solenoid Buzz
            End If
      End If

    If Keycode=RightFlipperKey Then
      RightFlipper.RotateToStart
      PlaySoundatvol "FlipperDown", RightFlipper, 1
            BuzzCount=BuzzCount-1
            If BuzzCount <= 0 Then
               StopSound "Buzz"           ' End Flipper Solenoid Buzz
            End If
      End If


  If Keycode=PlungerKey Then
        ShootBall=True                              'Reverse plunger animation
    Plunger.Fire
        If BallAtPlunger = True Then
           PlaySoundatvol  "PlungerRoll", Plunger, 1
        Else
           PlaySoundatVol "PlungerSpring", Plunger, 1
        End If
  End If
End Sub
'
' ****************************************************
' ****  Code To Move The Plunger Image on Screen  ****
' ****     Speed is set by PW1 Timer Interval     ****
' ****************************************************
Sub MovePlunge()
    If ShootBall=False Then
       ' Animating A Pull
       PW=PW+1
       PW1.TimerEnabled=True
       If PW < 12 Then
          EVAL("PW"&PW).IsDropped=False   ' Lower Current Wall
          EVAL("PW"&PW+1).IsDropped=True  ' Raise Next Wall
       End If
       If PW < 3 Then
          PW1.TimerInterval=85        ' Handle "hitch" in real plunger
       Else
          PW1.TimerInterval=PS
       End If
    Else
       ' Animating Release
       ' First lower all walls to get to known place
       For X=1 to 12
          EVAL("PW"&X).IsDropped=False
       Next
       If PW > 6 Then
          ' More than half=way down, release in two steps
          PW6.IsDropped=True
          PW1.TimerEnabled=True
          PW=6                              ' Set Phase to half drawn
       Else
          ' Let it all the way out and we're done
          PW1.IsDropped=True
          PW1.TimerEnabled=False
          PW1.TimerInterval=PS
          PW=0                              ' Set Phase to Full Released
          ShootBall=False
       End If
    End If
End Sub
'
'  Timer ticks and calls here to move pullback to next image
'
Sub PW1_Timer()
    MovePlunge()                 ' Move to next image and restart timer
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
  PlaySoundatvol "BallRoll1", ActiveBall, 1
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
Const SpeedX1 = 4 : Const SpeedX2 = -4
Const SpeedY1 = 4 : Const SpeedY2 = -4
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
       M1sReel.TimerEnabled=True          ' Start Game Over Sequence
       Exit Sub
  End If
    '
    If BallsLost >= 5 Then              ' Game is over after 5 Balls
       GOVPhase=1                 ' Game Over Processing Has Begun
       If DisplayHighScore=True Then
      SetHS(HighScore)              ' Display the current HIGH SCORE
     End If
       M1sReel.TimerEnabled=True          ' Start Game Over Sequence
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
'  Count Ball Number On Gate Exit
'
Sub Gate2_Hit()
    GateCount=GateCount+1         ' Count Another Ball Onto Playfield
    StartZCheck()             ' See if Z State Time
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
'  Note: This version assumes ScoreToAdd is only 1, 5, 10, or 50
'
'  Motor = 0 For normal Scoring call
'          1 for call from Scoring Motor
'          2 for score + Advance/Spot
'
Sub AddScore(ScoreToAdd, Motor)
    Dim NoSound
    NoSound=False               ' Assume Sound from here
  If InProgress=False Or Tilted=True Then
       Exit Sub                 'Disable score when TILTed or not started
    End If
    '
    ' Display the score
    '
    ' Note: "Scoring Motor" uses Th10sReel.Timer
    '
    If (ScoreToAdd = 5) OR (ScoreToAdd=50)  Then
       If ScoreToAdd=5 Then
          MUnits = MUnits + 5       ' 5 = Pulse Units 5 Times
       Else
          MTens = MTens + 5         ' 50 = Pulse Tens 5 Times
       End If
       If Th10sReel.TimerEnabled=False Then
          '
          ' Scoring Motor Not Running, Start it up
          '
           MotorValue=ScoreToAdd        ' Say what it's running for
          Th10sReel.TimerInterval=135   ' 140ms per pulse
          Th10sReel_Timer()         ' Score 1st Pulse Right Now
          Th10sReel.TimerEnabled=True   ' Start Scoring Motor Running
       End If
    Else
       '
       ' Score is either 1 or 10,
       '
       If (Motor = 1) OR (Th10sReel.TimerEnabled=False) Then
          '
          ' Motor is not running or this is a call from the motor
          '
          Score = Score+ScoreToAdd            'Total of current score
          SReelHuns=  INT((Score MOD 1000) / 100)   'Get New Scoring Reel Digits
          SReelTens = INT((Score MOD 100) /10)
          SReelUnits = (Score MOD 10)
          DisplayScore()                ' Display Scoring Lights
          If Motor=0 Then
             '
             ' Normal call, Start motor to assure minimum time between scores
             '
              MotorValue=ScoreToAdd       ' Say what it's running for
             Th10sReel.TimerInterval=135    ' 140ms per pulse
             Th10sReel.TimerEnabled=True    ' Start Scoring Motor Running
          End If
       Else
          '
          ' Normal 1 or 10 call and Motor is running
          ' Put in queue for motor to do
          '
          NoSound=True          ' No Sound at this level
          If MQ(0)=0 Then
             MQ(0)=2          ' Motor on 1st Queue Entry
             MQ(1)=3          ' This is Next queue Entry
             MQ(2)=ScoreToAdd     ' So it knows how many are active
          Else
             X=MQ(1)          ' Get next Queue Entry
              If X < 20 Then
                 '
                 '  Safety Check - Ignore events over 20 deep
                 '
                MQ(1)=X+1       ' Say it's used
                MQ(X)=ScoreToAdd    ' Store Event for motor to find
              End If
          End If
       End If
    End If
    '
    '  Check Mystery Special and Z State Entered from below
    '
    CheckMystery()            ' See if Mystery Special ON or Off
    StartZCheck()           ' See if Z State Entered
    '
    '  Check for replays Earned on Scoring
    '
    If Score=>Replay1 And Replay1Paid=False Then    'If the first replay score is reached,
       AddReplay()                    'then give a replay and
       Replay1Paid=True                 'mark the replay as paid
  End If
    If Score=>Replay2 And Replay2Paid=False Then    'If the first replay score is reached,
       AddReplay()                    'then give a replay and
       Replay2Paid=True                 'mark the replay as paid
  End If
    If Score=>Replay3 And Replay3Paid=False Then    'If the first replay score is reached,
       AddReplay()                    'then give a replay and
       Replay3Paid=True                 'mark the replay as paid
  End If
  If Score=>Replay6 And Replay6Paid=False Then    'Replay every 100k beginning at Replay5
       AddReplay()                    'then give a replay and
       Replay6=Replay6+10               'Next Replay at next 100,000
       If Replay6 > Replay5 Then
          Replay6Paid=True                'end hit. mark the replay as paid
       End If
  End If
    '
'
'  Play Appropriate Bell sounds to match scoring
'
  If (Bells=True) AND (Motor=0) AND (NoSound=False) Then
       If ScoreToAdd = 5 Then
          Play5Bells()                  'Avoid Overrun on 5 Bell Sounds
       Else
         If ScoreToAdd = 50 Then
            Play50Bells()               ' Ditto on 50 Bell Sounds
         Else
            PlaySound ScoreToAdd            ' Single Bell on 1 and 10 Sounds
         End If
      End If
    End if
 '
    If (Bells=True) AND (Motor=2) AND (NoSound=False) Then
       PlaySound "10kSpotMotor"
    End If

End Sub
'
'  "Scoring Motor" Timer Routine
'
Sub Th10sReel_Timer()
    If MSkip <> 0 Then
       MSkip=MSkip-1        ' Count a Skip Cycle
       Exit Sub           ' and wait till next tick
    End If
    If MUnits > 0 Then
       MUnits=Munits-1
       AddScore 1, 1        ' Score a units digit pulse
       Exit Sub
    End If
    If MTens > 0 Then
       MTens = MTens-1
       AddScore 10, 1       ' Score a Tens Digit pulse
       Exit Sub
    End If
     '
     ' Scoring Motor Cycle ended
     '
     If MotorValue=5 OR MotorValue=50 Then
        EndZCheck(0)          ' Check End of ZState Reached
     End If
    '
    ' See if any delayed scores
    '
    If MQ(0) <> 0 Then
       '
       ' We have one to do
       '
       MX=MQ(0)           'Get our queue index
       MX=MQ(MX)          'Get 1 or 10 to Score Next
       MQ(0)=MQ(0)+1        ' point to next entry
       If MQ(0)=MQ(1) Then
          '
          ' Queue Emtpy, reset pointers
          MQ(1)=2
          MQ(0)=0
       End If
        MotorValue=MX       ' Say what motor is running for
       PlaySound MX         ' Play the sound
       AddScore MX, 1       ' Add the score
    Else
       Th10sReel.TimerEnabled=False ' Stop Scoring Motor
    End If
End Sub
'
'  This Routine handles the possible rapid backup of
'  5 (50,000 point) and/or 50 (500,000 point) point bells,
'  and queues them to a timer to avoid overlapping them
'  if that is needed
'
Sub Play5Bells()
    Bell5ToDo = Bell5ToDo + 1
    If Bell50ToDo = 0 AND Bell5ToDo = 1 Then
       Playsound "5"                   ' No Sound queued, play one here
    End If
    If BumpBase.TimerEnabled=False Then
       BumpBase.TimerInterval=550         ' Wait min of 550ms
       BumpBase.TimerEnabled=True         ' Start Timer
    End If
End Sub
'
'  Same for 50 point bells
'
Sub Play50Bells()
    Bell50ToDo = Bell50ToDo + 1
    If Bell5ToDo = 0 AND Bell50ToDo = 1 Then
       PlaySound "50"
    End If
    If BumpBase.TimerEnabled=False Then
       BumpBase.TimerInterval=550          ' Wait Min of 550ms
       BumpBase.TimerEnabled=True
    End If
End Sub
'
'  50/5 bell overlap Timer
'
Sub BumpBase_Timer()
    If Bell50ToDo > 0 Then
       Bell50ToDo = Bell50ToDo - 1
       If Bell50ToDo > 0 Then
          PlaySound "50"        ' Do next 50 Bell
          Exit Sub                      ' Done Until Next Tick
       End If
       Bell50ToDo = 0                   ' No More 50's Check 5's
    End If
    Bell5ToDo = Bell5ToDo - 1
    If Bell5ToDo <= 0 Then
       Bell5ToDo = 0                    ' safety Check
       BumpBase.TimerEnabled=False      'Stop Timer
    Else
       PlaySound "5"                    ' Play Next Delayed 5 Bell
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
    If DelayedReplays = 0 Then
       PlaySound"Knocker"           'Play the sound
       If Credits < MaxCredits Then       'Credit Reel Max Check
          Credits=Credits+1           'Add a credit
          DisplayCredits()            ' Display Credits on Back Glass
       End If
    End If
    DelayedReplays=DelayedReplays+1
    HSTens.TimerEnabled=True          'Enable the timer for Overlapping Replays
End Sub
'
' *******************************************************
' ****  Multiple Replay Award With Knocker Sequence  ****
' ****                                               ****
' ****    HSTens.Timer_Enabled=True to start         ****
' *******************************************************
'
'  High Score Timer cycles every 250ms (1/4 sec)
'
'  4 cycles means it runs for 1 second when activated
'
'  Assumes 1 Replay and Knock done before timer started.
'  Delayed Replays are added at 250ms intervals
'  based on value DelayedReplays
'
Sub HSTens_Timer()
  If DelayedReplays > 1 Then
       PlaySound"Knocker"           'Play the sound
       If Credits < MaxCredits Then       ' Our credit Reel maxes at 99
          Credits=Credits+1           'Add a credit
          DisplayCredits()            ' Display Credits on Back Glass
       End If
       DelayedReplays=DelayedReplays-1      'One Less To Go
    Else
       DelayedReplays=0             'No More Delayed Replays
       Me.TimerEnabled=False
  End If
End Sub
'
' ******************************
' **** Tilt Bobber Simulator ***
' ******************************
'
'  This routine is called each time game is Nudged
'  to Check to see if Player has tilted the game.
Sub CheckTilt()                  'Called when table is nudged
    Count1=Count1+1              'Add to tilt count (hit lasts 1 second)
    If Count1>TiltSensitivity And Tilted=False Then  'If more than Allowed Counts then TILTED
       Tilted=True
       If InProgress=True Then
          DelayedGOV=True        'Say game Is Effectively over
       End If
       TiltReel.SetValue 1       'Put "TILT" On Backglss
       PlaySound "Reset1"
       Controller.B2ssettilt 33,1
       Disable(True)             'Disable slings, bumpers etc
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
     Case 4:PlaySound"Motor"
     Case 5:AddCredit           'Always add at least 1 credit
     Case 6:If CreditsPerCoin>1 Then AddCredit
     Case 7:If CreditsPerCoin>2 Then AddCredit
     Case 8:If CreditsPerCoin=5 Then AddCredit
     Case 9:
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
       DisplayCredits()                 ' Display Credits on Back Glass display
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
           GateCount=0                ' And None out the inlane gate yet
           TroughOpening=True           ' Say Trough is Opening
           TroughOpenHalf.IsDropped=True      ' Animate Half Open Trough
           TroughOpenHalf.TimerEnabled=True     ' Start Timer to Complete Open
 '
'  Cycle 2 (.4s after start) play a Reset sound and
'  Release the balls from the visible trough
    Case 2:
       Credits=Credits-1            ' Subtract a credit
           DisplayCredits()             ' Display Credits on Back Glass
         Score=0                  ' Reset SCORE
            ZState=0                  ' Reset Z State
           For X= 1 to 15
              Spotted(X)=0              ' Reset Numbers Spotted on BackGlass
           Next
           MadeSpecial=False                        ' Haven't Made Lane Special yet
           TiltReel.SetValue 0                      ' Clear "TILT" On Backglass
           Controller.B2SSetTilt 33,0
         BallsLost=0                ' Reset balls used
           BallsLeft=5                ' And Balls Left to Lift/Play
           If BallsOnField > 0 Then
              BallsLeft = BallsLeft-BallsOnField  'Start Hit Mid-previous Game
           End If
'
       Replay6=Replay4              'Reset 100K start
           Replay1Paid=False            'Reset Score Replays Paid Flags
           Replay2Paid=False
           Replay3Paid=False
           Replay4Paid=False
       Replay5Paid=False
       Replay6Paid=False
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
  Dim xx
       '
       '  Set Initial Game Start Light States:
       '
           '
           '  All Passive Bumpers Init to Off Except #1 & POPs
           '
           BumpersOff()             ' Set All Bumpers Off
           BumperOn(1)              ' Except 1
           BumperOn(8)              ' & POP Bumpers
           For Each xx in Bumperlights: xx.state=0:Next
           LightB8.state=1
           LightB9.state=1
            BumperOn(9)
            '
           '  Special Lights Off
           '
           LtExtraSpecial.IsDropped=True
           LtSemiGrn.IsDropped=True
           LtSemiRed.IsDropped=True
           LtSpecial1.IsDropped=True
           LtSpecial2.IsDropped=True
           LtSpecial5.IsDropped=True
           '
           '  Score Multipliers Off
           '
           LtHeadlight.IsDropped=True
           '
           '  Kicker Hole Light Off
           '
           CtrKickerCoverLit.IsDropped=True
            CtrKickerCoverOpen.IsDropped=True
           CtrKickerCover.IsDropped=False
           CtrKicker.UserValue=0        ' Say Center Not Lit
           '
           ' reset Round House Score to 100,000
           '
           RHPointer=1
           LtES1.IsDropped=False        ' Light ES1
           LtES2.IsDropped=True         ' ES2-ES5 are UnLit
           LtES3.IsDropped=True
           LtES4.IsDropped=True
           LtES5.IsDropped=True
           '
           MadeCount=0              ' 1 to 15 not made yet
'
'  Cycle 4 (.8s after start pressed) just generates delay
'
'  Cycle 5 (1 full second after start) We Reset the score counter
'
      Case 5:
           ResetLightsToZero()          'Reset score counters
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
Sub M1sReel_Timer()
  GOVPhase=GOVPhase+1               'Increment the Cycle counter
  Select Case GOVPhase
    Case 2:
       DelayedGOV=False             'End Any Tilt Delay
'
'  Cases 1-4 are just 400ms of delay
'
'  0.5s after end of game
    Case 5:
            PlaySound "Motor"           'Play the score motor sound
            Disable(True)
 '
'  Cases 6-15 are 900ms more delay
'
'   1.6s after end of game
    Case 16:
'
' Cases 17-20 are 400ms more delay
'
' 2.1s after End Of game sequence began we have the final cycle
    Case 21:
           SaveValue TableName,"Score",Score    'Save score value for next time
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
  Msg(2)="                         K. C. JONES"
  Msg(3)=""
  Msg(4)="  HIT LIT NUMBERED BUMPERS TO ADVANCE HOLE VALUE"
    Msg(5)=""
    Msg(6)="  HOLE AWARDS REPLAY AS INDICATED."
    Msg(7)=""
    Msg(8)="  SPECIAL WHEN LIT ROLLOVER AWARDS 1 REPLAY."
    Msg(9)=""
    Msg(10)="  1 REPLAY FOR 4 Million."
    Msg(11)=""
    Msg(12)="  1 REPLAY FOR 4 Million 300 THDS."
    Msg(13)=""
    Msg(14)="  1 REPLAY FOR 4 Million 500 THDS."
    Msg(15)=""
    Msg(16)="  1 REPLAY FOR 4 Million 800 THDS AND EVERY 100 THDS. THEREAFTER TO 6 MILLION."
    Msg(17)=""
  Msg(18)="  Press 5 to Deposit Coin, 1 to Start a game"
  Msg(19)="  Press A to lift next ball into play"
  Msg(20)=""
  Msg(21)="  2 Resets high Score to 1,000,000 and 3 Clears Credits to 0"
  For X=1 To 21
    Msg(0)=Msg(0)+Msg(X)&Chr(13)
  Next
  MsgBox Msg(0),,"         Instructions and Rule Card"
End Sub
'
' *********************************************
' ****  Handle 10,000 pt Contact switches *****
' *********************************************
'
Sub LSW1Rubber_Slingshot()
    AddScore 1, 0       ' 10,00 points
End Sub
'
Sub LSW2_Hit(Idx)
     If (ActiveBall.X > 400) AND _
        (ActiveBall.Y > 920) AND (ActiveBall.Y < 1060) Then
       If LtHeadLight.IsDropped=False Then
          AddScore 10, 0      ' 100,000 pts Lit
       Else
          AddScore 1, 0     ' 10,00 points unLit
       End If
     Else
        AddSound "Rubber"
     End If
End Sub
'
Sub RSW1Rubber_Slingshot()
    AddScore 1, 0       ' 10,00 points
End Sub
'
Sub RSW2_Hit(Idx)
    If (ActiveBall.X < 585) AND _
        (ActiveBall.Y > 920) AND (ActiveBall.Y < 1060) Then
       If LtHeadLight.IsDropped=False Then
          AddScore 10, 0      ' 100,000 pts Lit
       Else
          AddScore 1, 0     ' 10,00 points unLit
       End If
     Else
        AddSound "Rubber"
     End If
End Sub
'
' ******************************
' ****  Handle Bumper Hits *****
' ******************************
'
' -------------------
' "1" Passive Bumper
' -------------------
'
Sub Bumper1_Hit
    PlaySoundAtVol "Rubber", ActiveBall, 1
 If LightB1.state=0 Then AddScore 1, 2      ' 10,000 Pts + Spot Sound
 If LightB1.state=1 Then AddScore 1, 0     ' 10,000 Points
    LightB1.state=1

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
' -------------------
' "2" Passive Bumper
' -------------------
'
Sub Bumper2_Hit
    If LightB1.state=1 Then LightB2.state=1
    PlaySoundAtVOl "Rubber", ActiveBall, 1
 If LightB2.state=0 Then AddScore 1, 2      ' 10,000 Pts + Spot Sound
 If LightB2.state=1 Then AddScore 1, 0     ' 10,000 Points

    '
    '  Debounce Ball "Laying" on Passive Bumper
    '
    B2Top.TimerEnabled=False      'Stop Timer If Running
    B2Top.UserValue=True        'Say Debounce Started
    B2Top.TimerInterval=100       'Debounce for 100ms
    B2Top.TimerEnabled=True       'Start Debounce Window
End Sub
'
Sub B2Top_Timer()
    B2Top.TimerEnabled=False      'Stop Timer
    B2Top.UserValue=False       'Say Hit Counts Again
End Sub
'
' -------------------
' "3" Passive Bumper
' -------------------
'
Sub Bumper3_Hit
    If LightB2.state=1 Then LightB3.state=1
    PlaySoundAtVofl "Rubber", ActiveBall, 1
 If LightB3.state=0 Then AddScore 1, 2      ' 10,000 Pts + Spot Sound
 If LightB3.state=1 Then AddScore 1, 0     ' 10,000 Points

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
' -------------------
' "4" Passive Bumper
' -------------------
'
Sub Bumper4_Hit
    If LightB3.state=1 Then LightB4.state=1
    PlaySoundAtVol "Rubber", ActiveBall, 1
 If LightB4.state=0 Then AddScore 1, 2      ' 10,000 Pts + Spot Sound
 If LightB4.state=1 Then AddScore 1, 0     ' 10,000 Points
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
Sub Bumper5_Hit
    If LightB4.state=1 Then LightB5.state=1
    PlaySoundAtVol "Rubber", ActiveBall, 1
 If LightB5.state=0 Then AddScore 1, 2      ' 10,000 Pts + Spot Sound
 If LightB5.state=1 Then AddScore 1, 0     ' 10,000 Points
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
' -------------------
' "6" Contact Switch
' -------------------
'
Sub B6Switch_Hit(Idx)
     If LightB5.state=1 Then LightB6.state=1
     If ActiveBall.Y < 1425 OR ActiveBall.Y > 1660 Then
        AddSound "Rubber"       ' Hit Outside Contact Switch Zone
     Else
          PlaySoundAtVol "Rubber", ActiveBall, 1
          If SpotNumber(6)=1 Then
             AddScore 1, 2        ' 10,000 Pts + Spot Sound
          Else
             AddScore 1, 0        ' 10,000 Points
          End If
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
' -------------------
' "7" Contact Switch
' -------------------
'
Sub B7Switch_Hit(Idx)
     If LightB6.state=1 Then LightB7.state=1:LightB10.State=1:LightB11.State=1
     If ActiveBall.Y < 1425 OR ActiveBall.Y > 1660 Then
        AddSound "Rubber"       ' Hit Outside Contact Switch Zone
     Else
       If B7Top.UserValue=False Then
          PlaySoundAtVol "Rubber", ActiveBall, 1
          If SpotNumber(7)=1 Then
             AddScore 1, 2          ' 10,000 Pts + Spot Sound
          Else
             AddScore 1, 0          ' 10,000 Points
          End If
       End If
     End If
    '
    '  Debounce Ball "Laying" on Passive Bumper
    '
    B7Top.TimerEnabled=False      'Stop Timer If Running
    B7Top.UserValue=True        'Say Debounce Started
    B7Top.TimerInterval=100       'Debounce for 100ms
    B7Top.TimerEnabled=True       'Start Debounce Window
End Sub
'
'
Sub B7Top_Timer()
    B7Top.TimerEnabled=False      'Stop Timer
    B7Top.UserValue=False       'Say Hit Counts Again
End Sub
'
' -------------------
' "8" POP Bumper
' -------------------
'
Sub Bumper8_Hit()
    LightB8.state=0
    PlaySoundAtVol "Bumper", ActiveBall, 1
    AddScore 1, 0           ' 10,000 Points
    B8Rings()             ' Kick off Ring Animation
    B8Ring3.TimerEnabled=True
     BumperOff(8)           ' Flash When Hit
     B8Top.TimerInterval=100        ' For POP time
     B8Top.TimerEnabled=True
End Sub
 '
 '  End Flash Off
 '
 Sub B8Top_Timer()
     B8Top.TimerEnabled=False     ' Stop Timer
     BumperOn(8)              ' End Flash Off
 End Sub
'
'  Bumper 8 Pop Ring Animation Cycle Routine
'
Sub B8Rings()
    LightB8.state=1
    B8State=B8State+1
    Select Case B8State
      Case 1:
         B8Ring3.IsDropped=False
      Case 2:
         B8Ring2.IsDropped=False
      Case 3:
         B8Ring1.IsDropped=False
      Case 4:
         B8Ring2.IsDropped=False
      Case 5:
         B8Ring3.IsDropped=False
    End Select
End Sub
'
' Bumper 8 Ring Animation Timer
'
Sub B8Ring3_Timer()
    Select Case B8State
        Case 1:
          B8Ring3.IsDropped=True
        Case 2:
          B8Ring2.IsDropped=True
        Case 3:
          B8Ring1.IsDropped=True
        Case 4:
          B8Ring2.IsDropped=True
        Case 5:
          B8Ring3.IsDropped=True
    End Select
    If B8State<5 Then
       B8Rings()
    Else
       B8State=0
       B8Ring3.TimerEnabled=False
    End If
End Sub
'
' -------------------
' "9" POP Bumper
' -------------------
'
Sub Bumper9_Hit()
    LightB9.state=0
    PlaySoundAtVol "Bumper", ActiveBall, 1
    AddScore 1, 0           '10,000 Points
    B9Rings()             ' Kick off Ring Animation
    B9Ring3.TimerEnabled=True
     BumperOff(9)           ' Flash When Hit
     B9Top.TimerInterval=100        ' For POP time
     B9Top.TimerEnabled=True
End Sub
 '
 '  End Flash Off
 '
 Sub B9Top_Timer()
     B9Top.TimerEnabled=False     ' Stop Timer
     BumperOn(9)              ' End Flash Off
 End Sub
'
'  Bumper 9 Pop Ring Animation Cycle Routine
'
Sub B9Rings()
    B9State=B9State+1
    LightB9.state=1
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
' -------------------------
' "10" Green Passive Bumper
' -------------------------
'
Sub Bumper10_Hit
       PlaySoundAtVol "Rubber", ActiveBall, 1
       If LightB10.State=1 Then AddScore 50, 0        ' 500,000 Pts When Lit
       If LightB10.State=0 Then AddScore 5, 0         ' 50,000 Points

    EndZCheck(0)            ' See if Z state is over
    '
    '  Debounce Ball "Laying" on Passive Bumper
    '
    B10Top.TimerEnabled=False     'Stop Timer If Running
    B10Top.UserValue=True       'Say Debounce Started
    B10Top.TimerInterval=100      'Debounce for 100ms
    B10Top.TimerEnabled=True      'Start Debounce Window
End Sub
'
'
Sub B10Top_Timer()
    B10Top.TimerEnabled=False     'Stop Timer
    B10Top.UserValue=False        'Say Hit Counts Again
End Sub
'
' -------------------------
' "11" Green Passive Bumper
' -------------------------
'
Sub Bumper11_Hit
       PlaySoundAtVol "Rubber", ActiveBall, 1
       If LightB11.State=1 Then AddScore 50, 0       ' 500,000 Pts When Lit
       If LightB11.State=0 Then AddScore 5, 0         ' 50,000 Points

    EndZCheck(0)            ' See if Z state is over
    '
    '  Debounce Ball "Laying" on Passive Bumper
    '
    B11Top.TimerEnabled=False     'Stop Timer If Running
    B11Top.UserValue=True       'Say Debounce Started
    B11Top.TimerInterval=100      'Debounce for 100ms
    B11Top.TimerEnabled=True      'Start Debounce Window
End Sub
'
'
Sub B11Top_Timer()
    B11Top.TimerEnabled=False     'Stop Timer
    B11Top.UserValue=False        'Say Hit Counts Again
End Sub
'
'  Left 10,000/100,000 point Contact Switch
'
Sub SW12_Hit(Idx)
     If SW12Base.UserValue=True Then
        Exit Sub
     End If
    X=ActiveBall.X
    If X < 100 OR X > 275 Then
       PlaySoundAtVol "Rubber", ActiveBall, 1
       Exit Sub
    End If
    If SW12T1Lit.IsDropped=True Then
       AddScore 1, 0          ' Unlit - 10,000 Points
    Else
       AddScore 10, 0         ' Lit - 100,000 Points
    End If
     '
    '  Debounce Ball "Laying" on Lower Rubber
    '
    SW12Base.TimerEnabled=False     'Stop Timer If Running
    SW12Base.UserValue=True       'Say Debounce Started
    SW12Base.TimerInterval=250      'Debounce for 100ms
    SW12Base.TimerEnabled=True      'Start Debounce Window
End Sub
'
'
Sub SW12Base_Timer()
    SW12Base.TimerEnabled=False     'Stop Timer
    SW12Base.UserValue=False        'Say Hit Counts Again
End Sub
'
'  Right 10,000/100,000 point Contact Switch
'
Sub SW13_Hit(Idx)
     If SW13Base.UserValue=True Then
        Exit Sub
     End If
    X=ActiveBall.X
    If X < 700 OR X > 875 Then
       PlaySoundAtVol "Rubber", ActiveBall, 1
       Exit Sub
    End If
    If SW13T1Lit.IsDropped=True Then
       AddScore 1, 0          ' Unlit - 10,000 Points
    Else
       AddScore 10, 0         ' Lit - 100,000 Points
    End If
     '
    '  Debounce Ball "Laying" on Lower Rubber
    '
    SW13Base.TimerEnabled=False     'Stop Timer If Running
    SW13Base.UserValue=True       'Say Debounce Started
    SW13Base.TimerInterval=250      'Debounce for 100ms
    SW13Base.TimerEnabled=True      'Start Debounce Window
End Sub
'
'
Sub SW13Base_Timer()
    SW13Base.TimerEnabled=False     'Stop Timer
    SW13Base.UserValue=False        'Say Hit Counts Again
End Sub
'
'******************************************************
'  Handlers for Rollover/Kicker Lane Trigger Hit
'******************************************************
'
'  Bottom Center Rollover w/Kicker
 '
Sub LaneESTrig_Hit()
    LaneESWire2.IsDropped=False     ' Animate Sensor Switch Wire
    LaneESWire2.TimerEnabled=True   ' and start animation timer
    If LtExtraSpecial.IsDropped=False Then
       AddReplay()            ' Award Special
    End If
    AddScore 5, 0           ' 50,000 pts
    EndZCheck(0)            ' See if Z state is over
    '
    '  Kickout occurs after 6 has been spotted
    '
    If Spotted(6) = 1 Then
       LaneESWire.IsDropped=False   ' Animate Kicker Kicking
       LaneESWire.TimerEnabled=True   ' and start animation timer
       ActiveBall.VelY=-12        ' Kick Upward Force
       X=ActiveBall.VelX
       ActiveBall.VelX=-X       ' With Reflective Rebound Drift
       PlaySoundAtVol "Bumper", ActiveBall, 1
       Exit Sub
    End If
    '
End Sub
'
'  ----------------------------------------------
'  Outlane Kicker & Sensor Wire Animation Timers.
'  ----------------------------------------------
'
'  All they do is end the animation.
'
Sub LaneESWire_Timer()
    LaneESWire.IsDropped=True     ' Drop Kicker Up animation Wall
    LaneESWire.TimerEnabled=False   ' Stop Timer
End Sub
 '
Sub LaneESWire2_Timer()
    LaneESWire2.IsDropped=True      ' Drop Sensor Wire animation Wall
    LaneESWire2.TimerEnabled=False    ' Stop Timer
End Sub
'
'******************************************************
'  Subroutine to Spot a Number and
'  Check if a Key Sequence has been made
'*******************************************************
'
'
'  Key Sequences Are:
'
' 1-7 Advances RoundHouse Value
'
 '  Returns 0 if number not spotted, 1 if it is spotted
 '
Function SpotNumber(Num)
     SpotNumber=0           ' Assume Not Spotted
     If Spotted(Num) <> 0 Then
       Exit Function          ' Number Already Spotted
    End If
    '
    '  Following Code requires numbers to be made in strict sequence
    '
    If Num > 1 Then
       For X=1 To (Num-1)
          If Spotted(X) = 0 Then    ' Only Spot if next in Sequence
             Exit Function
          End If
       Next
    End If
    '
    '  OK to Spot this number
    '
    SpotNumber=1            ' Say we spotted it
     Spotted(Num)=1           ' Mark it as Spotted
    BumperOff(Num)            ' UnLight Bumper
     If Num < 7 Then
        BumperOn(Num+1)         ' Light Next Bumper
     End If
    AdvanceRH()             ' Advance Round House Value
End Function
'
'  Handle Round House Advance
'
Sub AdvanceRH()
    If RHPointer < 5 Then
       EVAL("LtES"&RHPointer).IsDropped=True  ' Unlight current RH score
       RHPointer=RHPointer+1
       EVAL("LtES"&RHPointer).IsDropped=False ' Light New RH score
    Else
       '
       '  At least 5 advances previously, do "Special" advancing
       '
        If LTSpecial5.IsDropped=False Then
           Exit Sub               ' Max Advances Already
        End If
       If LtSpecial1.IsDropped=True AND LtSpecial2.IsDropped=True Then
          LtSpecial1.IsDropped=False      ' Light 1 special
       Else
          If LtSpecial1.IsDropped=False Then
             LtSpecial1.IsDropped=True
             LtSpecial2.IsDropped=False     ' Light 2 Special
          Else
             LtSpecial2.IsDropped=True
             LtSpecial5.IsDropped=False     ' Light 5 Special
          End If
       End If
    End If
End Sub
'
'**************************
'    Handle Kicker Hit
'**************************
'
Sub CtrKickerTrig_Hit()
    Set aBall=ActiveBall        ' Grab Pointer to this ball
    If (AttractBall(aBall, CtrKicker.X, CtrKicker.Y, 70, 3) > 20) Then
        CtrKickerTrig.TimerInterval=10        'start 10ms updater
        CtrKickerTrig.TimerEnabled=True
    End If
    AddSound "RollMore"
End Sub
'
'  Timer to do attract every 10ms while Ball in zone
'
Sub CtrKickerTrig_Timer()
    If (AttractBall(aBall, CtrKicker.X, CtrKicker.Y, 70, 3) < 5) Then
       '
       ' Inside hole, let kicker grab ball
       '
       CtrKickerTrig.TimerEnabled=False
    End If
End Sub
'
'  Center Kicker 70 Unit Diameter Zone exited
'  without being grabbed by Kicker. Stop Attraction Timer
'
Sub CtrKickerTrig_UnHit()
    CtrKickerTrig.TimerEnabled=False
End Sub
'
'----------------------------------------
'  Actual Center Kicker Hole Entered
'----------------------------------------
'
Sub CtrKicker_Hit()
    CtrKickerTrig.TimerEnabled=False  ' Stop Attraction Timer
     InKicker=True            ' Say Ball In Kicker
    PlaySoundAtVol "KickerLand", ActiveBall, 1
    CtrKicker.UserValue=0       ' Assume not lit
    CtrKickerCover.IsDropped=True   ' Drop UnLit Cover
    If CtrKickerCoverLit.IsDropped=False Then
       CtrKickerCoverLit.IsDropped=True ' Lit - Drop Lit Cover
        CtrKickerCoverOpen.IsDropped=False  ' Raise Lit Half.cover
       CtrKicker.UserValue=1      ' Indicate it is lit
    End If
    '
    '  Check for Specials Lit
    '
    X=0                 ' Assume No Specials
    If CtrKicker.UserValue=0 Then
       If LtSpecial1.IsDropped=False Then : X=1 : End If
       If LtSpecial2.IsDropped=False Then : X=2 : End If
       If LtSpecial5.IsDropped=False Then : X=5 : End If
    End If
     '
    If X <> 0 Then
       For Y = 1 to X
          AddReplay()         ' Award Kicker Specials
       Next
    End If
    '
    If ZState <> 1 Then
        '
        '  Landed in Kicker in normal (not ZState) mode
        '
       RHTimer.Interval=140
       RHTimer.Enabled=True       ' Add Extra Score after delay
       CtrKicker.TimerInterval = 700  ' Set 700ms delay (motor position 4) to kickout
       CtrKicker.TimerEnabled = True  ' Enable Timer
    Else
       '
       '  Landed in Kicker in ZState, run score up
       ' to over 2,900,000 by cycling extra score value
       '
       RHTimer.Interval=50
       RHTimer.Enabled=True       ' Add Extra Score after delay
       CtrKickerCover.TimerInterval=980 ' Set scoring motor cycle time + 1 140ms tick
       CtrKickerCover.TimerEnabled=True ' start Catch-up Cycling
    End If
End Sub
'
' Center Kicker Cycling Timer
'
Sub CtrKicker_Timer()
    PlaySound "Kicker"
    CtrKicker.TimerEnabled = False    ' Stop Timer
     CtrKickerCoverOpen.IsDropped=True  ' Assure Half Cover Down
    If CtrKicker.UserValue=0 Then
       CtrKickerCover.IsDropped=False ' Raise UnLit Cover
       CtrKickerCoverLit.IsDropped=True ' Assure Lit Cover Is Dropped
    Else
       CtrKickerCoverLit.IsDropped=False ' Raise Lit Cover
       CtrKickerCover.IsDropped=True  ' Assure UnLit Cover Is Dropped
    End If
     InKicker=False           ' Say Ball No Longer in kicker
    CtrKicker.Kick 213-(RND()*10), 4  ' Kick out at 208 +/- 5 degrees
End Sub
'
'-----------------------------------------------------
'  Attraction Routine For Kicker Hole Enlargement Zone
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
End Function
 '
' ---------------------------------------
' Timer to Do ZState Kicker Catchup
' ---------------------------------------
'
 Sub CtrKickerCover_Timer()
      Dim I, J
      '
     EndZCheck(1)             ' See if We Got to End of ZState
     If ZState = 1 Then
        AddRHScore()            ' Add Another Extra Score
     Else
        CtrKickerCover.TimerEnabled=False ' End This Timer
        CtrKicker.TimerInterval = 140   ' Set 140ms delay to kickout
        CtrKicker.TimerEnabled = True   ' Enable kickout Timer
         '
        '  Award Any Specials at end of Z State
        '
        I=0                 ' Assume No Specials
        If CtrKicker.UserValue=0 Then
           If LtSpecial1.IsDropped=False Then : I=1 : End If
           If LtSpecial2.IsDropped=False Then : I=2 : End If
           If LtSpecial5.IsDropped=False Then : I=5 : End If
        End If
        '
        If I <> 0 Then
           For J = 1 to I
              AddReplay()         ' Award Kicker Specials
           Next
        End If
     End If
End Sub
 '
'  Timer for delay to add Extra Score
'
Sub RHTimer_Timer()
    RHTimer.Enabled=False         ' End Delay Timer
    AddRHScore()              ' Initiate Kicker Scoring Motor
End Sub
'
' --------------------------------------
' Add Current RoundHouse Score value
' ---------------------------------------
'
 '  Scoring Moter for Kicker scores on the 5 cycles as follows:
 '
 '
 '            1         2        3        4        5  <--- RHPointer Value
 '  Cycle
 '  ---------------------------------------------------
 '    1       0         0        0     100,000  100,000
 '    2       0         0     100,000  100,000  100,000
 '    3       0      100,000  100,000  100,000  100,000
 '    4    100,000   100,000  100,000  100,000  100,000
 '    5       0         0        0        0     100,000
 '         -------   -------  -------  -------  -------
 ' Total   100,000   200,000  300,000  400,000  500,000
 '
 '  A Cycle Scoring 0 should output sound "1"
 '
Sub AddRHScore()
     PlaySound "Motor"
    RHExtra=2               ' This is cycle 1 so next cycle is 2
    Select Case RHPointer
       Case 1
          PlaySound "1"           ' 100,000
           LtES2.TimerEnabled=True      ' start cycle timer
       Case 2
          PlaySound "1"           ' 200,000
          LtES2.TimerEnabled=True     ' start cycle timer
       Case 3
          PlaySound "1"           ' 300,000
          LtES2.TimerEnabled=True     ' start cycle timer
       Case 4
          AddScore 10, 0          ' 400,000
          LtES2.TimerEnabled=True     ' start cycle timer
       Case 5
          AddScore 10, 0          ' 500,000
          LtES2.TimerEnabled=True     ' start cycle timer
    End Select
End Sub
'
'  Timer simulates scoring motor doing 100k/pulse
'
'  Cycles as many times as EXExtra indicates, then stops.
'
Sub LtES2_Timer()
     Select Case RHExtra            ' Select cycle we are on (2-5)
        Case 2
           If RHPointer > 2 Then
              AddScore 10, 0          ' 100,000 this cycle
           Else
              PlaySound "1"         '  0 this cycle
           End If
        Case 3
           If RHPointer > 1 Then
              AddScore 10, 0          ' 100,000 this cycle
           Else
              PlaySound "1"         '  0 this cycle
           End If
        Case 4
           AddScore 10,0            ' Always 100,000 on cycle 4
        Case 5
           If RHPointer = 5 Then
              AddScore 10, 0          ' 100,000 this cycle
           Else
              PlaySound "Last0"       '  0 this cycle
           End If
    End Select
    RHExtra=RHExtra+1           ' count one more cycle done
    If RHExtra > 5 Then
       LtES2.TimerEnabled=False       ' done after cycle 5
    End If
End Sub
'
 '
' -------------------------------------------
'      Check on Start And/OR End of "Z" state
' -------------------------------------------
 '
'  5th Ball Game Logic Initiate Check
 '
 '  Init Z state on 5th Ball if Score between 800,000 and 2,900,000
'
 Sub StartZCheck()
     If (ZState=0) AND (GateCount=5) AND (Score > 80) AND (Score < 290) Then
       '
       ' Turn Z State On
       '
       ZState=1                   ' Mark in Z state
       LtHeadLight.IsDropped=False          ' Light Headlight on Ball 5
       BumperOn(10)                 ' Top Bumpers Light on Ball 5
       BumperOn(11)
       BumperOn(12)                 ' Bottom 10/100k To 100k
       BumperOn(13)
       CtrKickerCover.IsDropped=True        ' Drop UnLit Cover
        If InKicker=True Then
           CtrKickerCoverOpen.IsDropped=False   ' Raise Lit Half Cover
        Else
          CtrKickerCoverLit.IsDropped=False     ' Raise Lit Cover
        End If
       CtrKicker.UserValue=1            ' Indicate Kicker is lit
    End If
End Sub
'
'  Z State Ends if Score is > 2,900,000
'
'
'  Kicking = 0 if not called from kicker,
 '          = 1 if called from kicker
'
Sub EndZCheck(Kicking)
    If (ZState <> 0) AND (Score > 290) Then
       ZState=0                   ' Mark Not in Z State
       LtHeadLight.IsDropped=True         ' UnLight 100K Headlight
       BumperOff(10)                ' 50/500K Bumpers to 50K
       BumperOff(11)
       BumperOff(12)                ' Bottom 10/100k To 10k
       BumperOff(13)
       If Kicking=0 Then
          CtrKickerCoverLit.IsDropped=True      ' Drop Lit Cover
          CtrKickerCover.IsDropped=False      ' Raise UnLit Cover
       End If
       CtrKicker.UserValue=0            ' Indicate Kicker is Unlit
    End If
End Sub
'
' ---------------------------------------
'       Check on Mystery Special
' ---------------------------------------
'
'  Mystery Special happens if
'
'  last two digits of score = 7, 11, 22, or 24
'
Sub CheckMystery()
    Select Case INT(Score MOD 100)
       Case 7, 11, 22, 24
          LTExtraSpecial.IsDropped=False
          LtSemiRed.IsDropped=False
          LtSemiGrn.IsDropped=False
       Case Else
          LTExtraSpecial.IsDropped=True
          LtSemiRed.IsDropped=True
          LtSemiGrn.IsDropped=True
    End Select
End Sub
'
' ---------------------------------------
' Do Cycle of Semiphore Signal Lights
' ---------------------------------------
'
'
Sub SignalLights()
    LtSemiRed.TimerInterval=600     ' RR Signal cycle time
    LtSemiGrn.TimerInterval=5000    ' for 5 seconds
    LtSemiRed.IsDropped=False     ' Start with Red On
    LtSemiGrn.IsDropped=True      ' Safety, in case re-rentered
    LtSemiRed.TimerEnabled=True     ' Start WigWag Timer
    LtSemiGrn.TimerEnabled=True     ' Start session length timer
End Sub
'
'  Swap Signal Lighs when this timer goes off
'
Sub LtSemiRed_Timer()
    If LtSemiRed.IsDropped=False Then
       LtSemiRed.IsDropped=True     ' Cycle to Green Light On
       LtSemiGrn.IsDropped=False
    Else
       LtSemiRed.IsDropped=False    ' Cycle to Red Light On
       LtSemiGrn.IsDropped=True
    End If
End Sub
'
'  End Signal Light Session when this timer goes off
'
Sub LtSemiGrn_Timer()
    LtSemiRed.Timerenabled=False    ' Stop WigWag Timer
    LtSemiGrn.TimerEnabled=False    ' Stop Session Timer
    LtSemiRed.IsDropped=True      ' Turn Off Signal Lights
    LtSemiGrn.IsDropped=True
End Sub
'
' -------------------------------------------
'  Routines to Turn Bumper Lights On and Off
' -------------------------------------------
'
Sub BumpersOn()
    For X=1 to 13
       BumperOn(X)
    Next
End Sub
'
Sub BumpersOff()
    For X=1 to 13
       BumperOff(X)
    Next
End Sub
'
'  Turn Single Bumper Off
'
Sub BumperOff(Num)
    Select Case Num
       Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
          '
          ' 1-7 & 10-11 are passive bumpers 8 & 9 POP
       '   EVAL("B"&Num&"Base").State=LightStateOff
       '   EVAL("B"&Num&"SideOn").IsDropped=True
       '   EVAL("B"&Num&"SideOff").IsDropped=False
      '    EVAL("B"&Num&"Top").State=LightStateOff
       '   EVAL("B"&Num&"RimOn").IsDropped=True
          If Num < 8 OR Num > 9 Then
       '      EVAL("B"&Num&"Top1").IsDropped=False
          End If
      '    EVAL("B"&Num&"Top1Lit").IsDropped=True

       Case 12, 13
          '
          '  12 & 13 are lower contact switch cover light
          '
     '     EVAL("SW"&Num&"Light").IsDropped=True
      '    EVAL("SW"&Num&"T1Lit").IsDropped=True
       '   EVAL("SW"&Num&"T2Lit").IsDropped=True
       '   EVAL("SW"&Num&"T3Lit").IsDropped=True
    End Select
End Sub
'
' Turn Single Bumper On
'
Sub BumperOn(Num)
    Select Case Num
       Case 1, 2, 3, 4, 5, 6, 7, 8 , 9, 10, 11
          '
          ' 1-7 & 10-11 are passive bumpers 8 & 9 Pop
          '
     '     EVAL("B"&Num&"Base").State=LightStateOn
      '    EVAL("B"&Num&"SideOn").IsDropped=False
      '    EVAL("B"&Num&"SideOff").IsDropped=True
    '      EVAL("B"&Num&"Top").State=LightStateOn
      '    EVAL("B"&Num&"RimOn").IsDropped=False
          If Num < 8 Or Num > 9 Then
      '       EVAL("B"&Num&"Top1").IsDropped=True
          End If
      '    EVAL("B"&Num&"Top1Lit").IsDropped=False

       Case 12, 13
          '
          '  12 & 13 are lower contact switch cover light
          '
       '   EVAL("SW"&Num&"Light").IsDropped=False
       '   EVAL("SW"&Num&"T1Lit").IsDropped=False
        '  EVAL("SW"&Num&"T2Lit").IsDropped=False
        '  EVAL("SW"&Num&"T3Lit").IsDropped=False
    End Select
End Sub
'
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
'  ******************************************
'   Subroutine to Display Credits on Backglass
'  ******************************************
'
 Sub DisplayCredits()
     CreditReel.SetValue INT(Credits MOD 10)        ' Display Units
     Credit10Reel.SetValue INT((Credits MOD 100)/10)    ' Display Tens
     Showcredits
 End Sub
'
'  ******************************************
'   Subroutine to Display Score on Backglass
'  ******************************************
'
'   SReelHuns = Millions Digit
'   SReelTens = 100 Thousands Digit
'  SReelUnits = 10 Thousands Digit
'
Sub DisplayScore()
    Showscore
    ' Do Millions Display
    '
    M1sReel.SetValue SReelHuns      ' Numbers 1-4 On left Side
    '
    ' Do Hundred Thousands Display
    '
    If SReelTens < 6 Then
       TH100sReelA.Setvalue SReelTens ' 100 - 500 on left Side
       TH100sReelB.SetValue 0
    Else
        TH100sReelA.SetValue 0
       TH100sReelB.SetValue SReelTens-5 ' 600 - 900 = 1-4 on Right Side
    End If
    '
    ' Do Ten Thousands Display
    '
    TH10sReel.SetValue SReelUnits   ' Light 0 to 90,000 on Mini BackGlass
     If SReelUnits < 5 Then
        TH10sReelA.SetValue SReelUnits  ' 0-40,000 on left of large BackGlass
        TH10sReelB.SetValue 0
      Else
        TH10sReelA.SetValue 0
         TH10sREELb.SetValue SReelUnits-4 ' 50,000 - 90,000 on Right of Large Back Glass
     End If
End Sub
'
'---------------------------------
'      Reset Lights to Zero
'---------------------------------
'
Sub ResetLightsToZero()
    ResetTimer.Interval=180   ' Step at 180ms/step
    ResetTimer.Enabled=True   ' Start the Reset
End Sub
'
'  Timer does the reset 1 tick at a time
'
Sub ResetTimer_Timer()
    X=True                ' Check for Score at Zero
    If SReelUnits <> 0 Then
       SReelUnits=SreelUnits-1      ' Step 10,000 Lights Down 1
       X=False              ' Don't Stop This Cycle
    End If
    If SReelTens <> 0 Then        ' Step 100,000 Lights Down 1
       SReelTens=SreelTens-1
       X=False              ' Don't Stop This Cycle
    End If
    If SReelHuns <> 0 Then
       SReelHuns=SreelHuns-1      ' Step Millions Lights Down 1
       X=False              ' Don't Stop This Cycle
    End If
    If X=True Then
       ResetTimer.Enabled=False   ' Stop Reset, We're done
    End If
    DisplayScore()            ' Update Lights w/new value
End Sub
'
'-----------------------------------------------------------
Sub Showscore()
    Dim N1         ' Define Local Temp Variable
    '
    For N1=71 to 77
          Controller.B2SSetData N1, 0   ' Set All Millions Lights Off
    Next
    If SReelHuns > 0 Then
          Controller.B2SSetData 70+SReelHuns, 1 ' Set Correct Millions Light On
    End If
    '
    For N1=61 to 69
          Controller.B2SSetData N1, 0    ' Set All Hundred Thousand Lights Off
    Next
    If SReelTens > 0 Then
          Controller.B2SSetData 60+SReelTens, 1 'Set Correct Hundred Thousand Light On
    End If
    '
    For N1=51 to 59
          Controller.B2SSetData N1, 0   ' Set All Ten Thousand Lights Off
    Next
    If SReelUnits > 0 Then
          Controller.B2SSetData 50+SReelUnits, 1 ' Set Correct Ten Thousand Light On
    End If
End Sub

Sub Showcredits()
 If Credits>9 Then Credits=9
    Dim M1
    For M1=10 to 19
          Controller.B2SSetData M1, 0
    Next
 Select Case Credits
 Case 0: Controller.B2sSetData 10,1
 Case 1: Controller.B2sSetData 11,1
 Case 2: Controller.B2sSetData 12,1
 Case 3: Controller.B2sSetData 13,1
 Case 4: Controller.B2sSetData 14,1
 Case 5: Controller.B2sSetData 15,1
 Case 6: Controller.B2sSetData 16,1
 Case 7: Controller.B2sSetData 17,1
 Case 8: Controller.B2sSetData 18,1
 Case 9: Controller.B2sSetData 19,1
 End Select
 End Sub


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


' *********   C.O.D.  *************
'          Williams 1953
'
'  Version 1.0 PLB October 2018
'
'  Thanks to too many to name, as I have learned from
'  those who went before.  Anyone may feel free to steal any
'  Ideas, Sounds or Images from this game at will and without
'  notice.
'
' My Goal was to try for as realistic game play as I could manage
' especially sounds and timing.  Images for playfield and Backglass
' are all originally taken from Real Games photographed on IPDB and
' other internet sites manipulated in Photoshop to mesh together with
' the dirt Cleaned off in some places (Think of it as photoshop Novus 2).
'
'       A Note on the internal handling of scoring:
'       -------------------------------------------
'
'  This game's display adds four zeros to the standard 1 point scoring,
'  but only on the display.  So internally the following is used:
'
'       10,000 points =  1 point internally
'       50,000 points =  5 points internally
'      100,000 points = 10 points internally
'      500,000 points = 50 points internally
'
'  These are the only four point scoring values possible, so the
'  extra zeros are not carried in the code.
'
'  As an example. 5,510,000 points is internally carried as 551 points.
'  The display adds the extra zeros which never change.
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
Dim DelayedGOV      'True Druing Interval From Tilt to Game Over Seq Started
Dim DelayedReplays    'Count Of Delayed Replays to Add with Knocker
Dim Count       'Cycle Counter for Credit_Timer
Dim Count1        'Used in TiltTimer for tilt sensitivity
Dim GOVPhase      'Cycle Counter for Game Over Sequence
Dim Match       'The match number
Dim MatchEnabled    'TRUE to enable number match, FALSE otherwise
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
Dim Bells
Dim Msg(30)       ' Used to display the rules
'
Dim B1State             ' B1 Animation State
Dim B2State             ' B2 Animation State
Dim PW          ' Plunger Animation State
Dim ShootBall           ' Is animation Pulling, or Releasing?
Dim PS                  ' Pull Speed
Dim LRLast              ' Units Digit after Previous AddScore()
Dim RFlipButton     ' True=Right Flipper Button Pressed, False = Not Pressed
Dim LFlipButton     ' True=Left Flipper Button Pressed, False = Not Pressed
Dim FlipperMode     ' 1= Impulse, 0 = Modern
    Const Impulse=1
    Const Modern=0
'
Dim aBall(11)     ' Trap Hole active Ball pointers
Dim aXpos(11)
Dim aXadd(11)
Dim aYpos(11)
Dim aYadd(11)
Dim aZpos(11)     ' Trap Hole Drop Vertical Position
'
Dim HForce(11)      ' Force per hole
Dim HoleCheat(11)   ' Single Hit Cheat per hole
 Dim HoleTicks(11)    ' # 10ms ticks spent in trigger zone
 Dim HoleTLow(11)   ' #10ms ticks ball below hole center
'
Dim Spotted(10)     ' Spotted Number 1-8 Array
Dim DrainWallHit    ' Limits to one it sound per ball on Drain Wall
Dim TrayHit       ' Limits to one sound per ball on Tray Hit
Dim BallID        ' Ball ID number
Dim BallDrained(25)   ' List of IDs of drained balls (index 0 = # entries in list)
Dim TroughOpening   ' True = Opening, False = Closing
Dim RaBall        ' Right AutoFlipper Ball Pointer
Dim MotorPhase      ' scoring motor phase
Dim HoleHit(10)     ' List of Trap Holes 0=No Ball, 1= Ball Trapped in Hole
Dim HolePtr(10)     ' List of pointers to Balls in Trap Holes
Dim Made1to8      ' True= Numbers 1 to 8 Made
Dim MadeCOD       ' True= C-O-D Holes Made
Dim MadeStar      ' True= Star Hole Made
Dim MadeShamrock    ' True= Shamrock Hole Made
 Dim RubberMode     ' 0 = Space above Autoflipper, 1 = Rubber between posts
'

 Dim Controller
 LoadController
 Sub LoadController()
 Set Controller = CreateObject("B2S.Server")
 Controller.B2SName = "C.O.D."  ' *****!!!!This must match the name of your directb2s file!!!!
 Controller.Run()
 End Sub


' *** Game Init starts here ***
'
Sub Table_Init()
    FirstOut=False
    BallAtPlunger=False
    DelayedStart=False  'No Delayed Anything yet
    DelayedGOV=False
    DelayedReplays=0
    GrowlTimer.UserValue=0
'
  TableName="COD"   'Place your table name here
  CreditsPerCoin=5  '1 to 5 credits per coin
    MaxCredits = 39     'Size of Credit Reel Graphic
  HighScoreReward=0 'The number of replays to award when HIGH SCORE is beaten (1-5)
'
    TiltTimer.Enabled=True
    TiltSensitivity=4 '0-15, 4=normal - a higher number is less sensitive, 0=1 nudge and you're out
  MatchEnabled=True 'Match feature enabled
  Bells=True      'TRUE to play bell sounds, FALSE for no bells
'
  Replay1=450     ' 1st Replay at 450 (4,500,000) Points
  Replay2=500     ' 2nd Replay at 500 (5,000,000) Points
  Replay3=600     ' 3rd Replay at 600 (6,000,000) Points
    Replay4=700     ' 4th Replay at 700 (7,000,000) Points
    Replay5=800     ' 5th Replay at 800 (8,000,000) Points
'
  Initialize()    'Call the main Init sub
End Sub
'
'----------------------------------------
'  Selected Ball Hits - Sound Only
'----------------------------------------
'
Sub Rubbers_Hit(Idx)
    Dim I, J
    I=ActiveBall.VelY
    J=ActiveBall.VelX
    If (SQR((I*I)+(J*J)) > 3) Then
       AddSound "RubberRoll"
    Else
       AddSound "Rubber"
    End If
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
       AddSound "Wood"        ' Hit back of Return Trough
        TrayHit=True        ' One Hit for this ball
    End If
End Sub
'
Sub DrainWall_Hit()
    If DrainWallHit=True Then
       AddSound "Roll9"       ' Roll After first hit
    Else
       AddSound "Wood"        ' First Drain Wall Hit
        DrainWallHit=True     ' Say Roll After this
    End If
End Sub
'
'---------------------------------------
'  Routine to Set High Score in HS Reels
'---------------------------------------
'
Sub SetHS(HScore)
    X=HScore MOD 10
    HS10Thousand.SetValue X
    X = INT((HScore MOD 100) /10)
    HS100Thousand.SetValue X
    X = INT((HScore MOD 1000) / 100)
    HSMillion.SetValue X
End Sub
'
'
'-------------------------------------------
'***** Object and light Disable/Enable *****
'-------------------------------------------
'
Sub Disable(State)            'Disable or enable objects when TILTed
'    SlingShot.Disabled=State
    ' Disable/Enable Bumpers
    For X=1 to 2
   '    EVAL("B"&X&"Top").Disabled=State
    '   EVAL("B"&X&"Base").Disabled=State
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
  LeftFlipper.RotateToStart
  RightFlipper.RotateToStart


    While BuzzCount > 0
       StopSound "Buzz"
       BuzzCount=BuzzCount - 1
    Wend

End Sub
'
'
'**********************************
'***** Main program Init code *****
'**********************************
'
Sub Initialize()
  Randomize             'Seed random Number Generator
  Match=INT(Rnd()*10)         'Generate a random number 0-9 for the match counter to start with
'
    RFlipButton=False
    LFlipButton=False
    FlipperMode=Impulse         ' Default to Impulse Flippers
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
    GrowlTimer.UserValue=False      ' Init AddSound to "Can Add"
'
  DisplayHighScore=True       'Display the high score
'
    DisplayGOV(1)           ' Show "Game Over"
    '
    '  Init Ball Return Trough
    '
    TroughOpen.IsDropped=True     ' Trough inits to Closed
    TroughOpenHalf.IsDropped=True   '  ... Fully Closed
    TroughOpening=False         ' Say Closing was last action
    TroughEdge.IsDropped=True     ' Lower Trough Front Edge
    DrainTrig.UserValue=0       ' Init debouncer
     '
    '  Init Holes as closed & Empty
    '
    For X= 0 to (Holes.Count-1)
       HolesOpen(X).IsDropped=True
       If X < 5 Then
          EVAL("Hole"&X+1&"OpenFull").IsDropped=True
          HoleHit(X+1)=0
       End If
       Holes(X).Enabled=False     ' Init holes to disabled
    Next
    '
    '  Init 1 to 8 to Not Made
    '
    For X=1 to 8
       Spotted(X)=0           ' UnSpot 1 to 8
        EVAL("Lt"&X).IsDropped=True   ' Set Number Light Off
    Next
    '
    ' Init Lane Wire Covers to Down
    '
    '  Wire covers contained in collection "Wires"
    '
    For each x in Wires
        x.IsDropped=True
    Next
     '
     GIOff()                ' UnLight General Illumination
    '
    '  Init game Special logic
    '
    Made1to8=False            ' Mark special things not made yet
    MadeCOD=False
    MadeStar=False
    MadeShamrock=False
    LtCOD1.IsDropped=True       ' Unlight C.O.D Lights
    LtCOD2.IsDropped=True
    LtCOD3.IsDropped=True
    LtStar.IsDropped=True       ' UnLight Star Lights
    LtOnStar.IsDropped=True
    LtShamrock.IsDropped=True     ' UnLight Shamrock Lights
    LtOnShamRock.IsDropped=True
    LtSpecialOut.IsDropped=True     ' Unlight Special Lights
    LtExSpecialOut.IsDropped=True
    '
    '   All Bumpers init to Off
    '
    For X=1 to 3
       BumperOff(X)
        If X < 3 Then
          ' Reset animation on POP Bumpers
          Eval("B"&X&"Ring1").IsDropped=True
          Eval("B"&X&"Ring2").IsDropped=True
          Eval("B"&X&"Ring3").IsDropped=True
          Eval("B"&X&"Ring3").TimerInterval=10
        Else
      ' Set Not Debouncing on Passive Bumper
      Eval("B"&X&"Top").UserValue=False
        End If
    Next
    B1State=0
    B2State=0
    '
    '   init Contact Switch Debouncers
    '
    RSW1Hit.UserValue=False
    RSW1Hit.TimerInterval=500
    RSW2Hit.UserValue=False
    RSW2Hit.TimerInterval=250
    RSW3Hit.UserValue=False
    RSW3Hit.TimerInterval=250
    RSW4Hit.UserValue=False
    RSW4Hit.TimerInterval=250
    RSW5Hit.UserValue=False
    RSW5Hit.TimerInterval=250
    RSW6Hit.UserValue=False
    RSW6Hit.TimerInterval=250
    RSW7Hit.UserValue=False
    RSW7Hit.TimerInterval=500
    LSW1Hit.UserValue=False
    LSW1Hit.TimerInterval=250
    LSW2Hit.UserValue=False
    LSW2Hit.TimerInterval=500
    LSW3Hit.UserValue=False
    LSW3Hit.TimerInterval=250
    '
    RollFlag=0                          'No Ball Rolling Yet
    RollVary=0                          'Init Roll Sound Variation flag
    ShootBall=0                       ' Next Move is Pull
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
    FlippersTimer.Enabled=True
    '
  On Error Resume Next        'If Values have never been saved in this table name don't error out
    '
  Credits=Cdbl(LoadValue(TableName,"Credits"))
  If Credits="" Then Credits=1      'If there was no entry, then set to default value
    DisplayCredits()
    '
  FlipperMode=Cdbl(LoadValue(TableName,"FlipperMode"))
  If FlipperMode="" Then FlipperMode=Impulse  'If there was no entry, then set to default value
    '
    RubberMode=Cdbl(LoadValue(TableName,"RubberMode"))
  If RubberMode="" Then RubberMode=0        ' If there was no entry, then set to default value
    SetRubberMode()               ' Set Rubber as selected
     '
  HighScore=Cdbl(LoadValue(TableName,"HighScore"))
  If HighScore="" Then HighScore=246      'If there was no entry, then set to default value
    '
  Score=Cdbl(LoadValue(TableName,"Score"))  'Load score value from previous game
    If DisplayHighScore=True Then
       SetHS(HighScore)             ' Put On Display
    End If
  If Score="" Then Score=155          'If there was no entry, then use default value
    SReelHuns=  INT(Score/100)          ' Set score reels to previous score value
    SReelTens = INT((Score MOD 100) /10)
    SReelUnits = (Score MOD 10)
    DisplayScore()
    MUnits=0                  'Init Scoring Motor to stopped
    MTens=0
End Sub
 '
'-----------------------------------
'  General Illumination On/Off
'-----------------------------------
'
Sub GIOn()
    Dim I
    For I=0 to (LtGI.Count-1)
       LtGI(I).IsDropped=False
    Next
    For I=0 to (LtGIPf.Count-1)
       LtGIPf(I).IsDropped=True
    Next
End Sub
'
Sub GIOff()
    Dim I
    For I=0 to (LtGI.Count-1)
       LtGI(I).IsDropped=True
    Next
    For I=0 to (LtGIPf.Count-1)
       LtGIPf(I).IsDropped=False
    Next
End Sub
'
'
'*********************************
'***** Keys Pressed Handling *****
'*********************************
Sub Table_KeyDown(ByVal Keycode)            'The usual keycode stuff
  If Tilted=False Then
    '
    ' Flipper and Nudge only when game in
    ' progress and not tilted
    '
      If Keycode=LeftFlipperKey Then
           LFlipButton=True
           If RFlipButton=False Then
              CloseFlippers()
           End If
      End If

    If Keycode=RightFlipperKey Then
           RFlipButton=True
           If LFlipButton=False Then
              CloseFlippers()
           End If
      End If

    If Keycode=LeftTiltKey Then
      Nudge 270,0.8
      CheckTilt()
    End If

    If Keycode=RightTiltKey Then
      Nudge 90,0.8
      CheckTilt()
    End If

    If Keycode=CenterTiltKey Then
      Nudge 0,0.95
      CheckTilt()
    End If
    If Keycode=MechanicalTilt Then
      CheckTilt()
    End If

    '
    ' End flipper and Nudge Qualifier
    '
  End If

  If Keycode=PlungerKey Then
       ShootBall=False              ' Indicate Animating Pull
       Plunger.PullBack
    End If

  If Keycode=AddCreditKey Then
       PlaySound"Coin"
     CreditReel.TimerEnabled=True
  End If

  If (Keycode=StartGameKey) And (Credits > 0) Then
       If GOVPhase <> 0 or DelayedGOV=True  Then
          DelayedStart=True         'let End of game Sequence Finsh First
       Else
          Plunger.TimerEnabled=True     'Start the "Game Start" Reset Sequence
       End if
    End If

     ' "A" or RightMagnaSave Activates Ball Lifter
  If Keycode=30 And BallsLeft>0 AND Tilted=False Or Keycode=RightMagnaSave And BallsLeft>0 AND Tilted=False Then
     PlaySound"BallOut"
     Lifter.TimerEnabled=True       'Activates the ball lifter Timer
  End If

    ' Either "R" or "I" displays rules
  If Keycode=23 or KeyCode = 19 Then
        Rules()
    End If

     ' "2" = Reset High Score to 246  (2,460,000)
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

    ' "4" = Toggle Flipper Mode
    If KeyCode=5 Then
       If FlipperMode=Impulse Then
          FlipperMode=Modern
          Msg(0)="Flipper Mode Set to Modern."
       Else
          FlipperMode=Impulse
          Msg(0)="Flipper Mode Set to Impulse."
       End If
       Msg(1)="         Flipper Mode Changed"
       LeftFlipper.TimerInterval=150
       LeftFlipper.TimerEnabled=True
       SaveValue TableName,"FlipperMode",FlipperMode
    End If

     If KeyCode=7 Then
         RubberMode=RubberMode+1
        If RubberMode > 2 Then
           RubberMode=0
        End If
       SaveValue TableName,"RubberMode",RubberMode
       SetRubberMode()
     End If

End Sub
 '
'  Routine to Delay Flipper Change Message 150ms
'
Sub LeftFlipper_Timer()
    LeftFlipper.TimerEnabled=False
    MsgBox Msg(0),,Msg(1)
End Sub
'
'
'  -----------------------
'  Ganged Flipper Movement
'  -----------------------
'
Sub CloseFlippers()
  LeftFlipper.RotateToEnd
    LeftFlipper2.RotateToEnd
    RightFlipper.RotateToEnd
  PlaySoundAtVol "FlipperUp", LeftFlipper, 1
    If FlipperMode=Impulse Then
        ImpulseTimer.Interval=130     ' Impulse Length
        ImpulseTimer.Enabled=True     ' start impulse timer
    Else
       PlayLoopSoundAtVol "Buzz", Leftflipper, -1
       BuzzCount=BuzzCount+1
    End If
End Sub
'
Sub OpenFlippers()
    If FlipperMode=Modern Then
       If LFlipButton=True OR RFlipButton=True Then
          Exit Sub              ' Both Buttons Not Released
       End If
    End If
  LeftFlipper.RotateToStart
    LeftFlipper2.RotateToStart
    RightFlipper.RotateToStart
  PlaySoundAtVol "FlipperDown", LeftFlipper, 1
    If FlipperMode=Modern Then
       StopSound "Buzz"           ' Stop Any Flipper Solenoid Sound
       BuzzCount=BuzzCount-1
    End If
End Sub
'
'  Impulse Flipper Timer End
'
Sub ImpulseTimer_Timer()
    ImpulseTimer.Enabled=False      ' Stop Timer
    OpenFlippers()
End Sub
'
'
'  VP9 Version Flipper Trcking Timer
'
Sub UpdateFlipperLogos
    LF1.RotAndTra2 = LeftFlipper.CurrentAngle
    LF21.RotAndTra2 = LeftFlipper2.CurrentAngle
    RF1.RotAndTra2 = RightFlipper.CurrentAngle
    RF21.RotAndTra2 = RightAutoFlipper.CurrentAngle
End Sub

Sub FlippersTimer_Timer()
    UpdateFlipperLogos
End Sub
'
'***** Keys Released Handling *****
'
Sub Table_KeyUp(ByVal Keycode)

  If Keycode=LeftFlipperKey Then
       LFlipButton=False
       If FlipperMode=Modern Then
          OpenFlippers()
       End If
  End If

  If Keycode=RightFlipperKey Then
       RFlipButton=False
       If FlipperMode=Modern Then
          OpenFlippers()
       End If
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
'**********************************************************
'***** Routines to Handle Ball Motion/Sounds on Table *****
'**********************************************************
'
'  Keep Track of whether Ball Is At Plunger or Not
'
Sub OnPlunger_Hit()
    BallAtPlunger = True
End Sub

Sub OnPlunger_UnHit()
    BallAtPlunger = False
    RollBack=0                        'Init Look for Roll-back Sound
End Sub

' Add Gate Noise On shot
' Also add bounce sounds as ball bounces back and froth at top
'
' Gate hit as ball enters playfield
Sub Gate_Hit()                    'Just for the sound
  PlaySoundAtVol "BallRoll1", Gate, 1
End Sub
'
' Hit gate going wrong way at top
Sub Rebound_Hit()
    If FirstOut=False Then
       PlaySoundAtVol "Gate5", ActiveBall, 1              ' Gate Bounce Sound
       AddSound "RollMore"
    End If
    FirstOut=False
    RollBack=0              ' Got out of gate
End Sub
'
' Hit Rebound Wheel at top left on Gate Exit
Sub Wheel_Hit()
    If FirstOut=True Then
       FirstOut=False
    End If
    AddSound "RubberRoll"
End Sub
'
'
'*******************************************
'***    Ball Rolling Sound Generator     ***
'***    Learned Concept from PacDude's   ***
'***    Excellent Centaur table          ***
'*******************************************
'
' Ball Roll Sensors
'
' Sensor for when ball rolls on Shooting Lane
'
Sub LaneSwitch_Hit()
    If RollBack=1 Then
       ' Failed to Clear Gate on Plunger Release
       AddSound "Roll9"
    Else
       ' On way out of lane from release
       RollBack = 1
    End If
End Sub
'
' Sensor for Upper Right Playfield
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
  Activeball.velx=Activeball.velx/4   'Stabilize in left lane
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
          RollVary = RollVary + 1
        Select Case RollVary
           Case 0
             AddSound "Roll9"
           Case 1
              AddSound "Roll2125"
          Case 2
             PlaySound "RollMore"
             RollVary=0
       End Select
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
    TroughEdge.IsDropped=True           ' Allow Ball Into Visible Trough
    TroughEdge.TimerEnabled=False         ' Stop timer in case bouncing
    TroughEdge.TimerEnabled=True          ' Timer to restore upper edge
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
    If BallsLost >= 5 Then              ' Game is over after 5 Balls
       GOVPhase=1                 ' Game Over Processing Has Begun
       If DisplayHighScore=True Then
      SetHS(HighScore)              ' Display the current HIGH SCORE
     End If
       M1sReel.TimerEnabled=True          ' Start Game Over Sequence
  End If
End Sub
'
Sub TroughEdge_Timer()
    If InProgress=False Then
       Exit Sub                   ' Ignore during init
    End If
    TroughEdge.TimerEnabled=False         ' Stop Timer
    TroughEdge.IsDropped=False            ' Restore Front Trough Edge
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
    Dim BallPtr
    DrainWallHit=False          ' allow DrainWall hit Sound
    Set BallPtr=Lifter.CreateBall     ' Create all And Add Poitner to List
    BallPtr.UserValue=BallID      ' ID that Ball
    BallID=BallID+1           ' Increment Ball ID for next time
    Lifter.Kick 40,5          ' Kick it up to the plunger lane
    BallsLeft=BallsLeft-1       ' One less Ball remains
    BallsOnField=BallsOnField+1     ' One more is on the Playfield
    FirstOut=True
    Lifter.TimerEnabled=False     ' Disable This Timer 'Till Next time
End Sub
'
' If Ball ends up back in trough, kick it out
Sub Lifter_Hit()
    Lifter.Kick 40, 5
End Sub
'
'*****************************************************
'****  End of the trough/lifter management routines ****
'*****************************************************
'
'******************************
'**** Scoring and credits *****
'******************************
'
'  AddScore is main scoring routine
'
'  Note: This version assumes ScoreToAdd is only 1, 5, 10, or 50
'
'  Motor = 0 For normal Scoring call
'          1 for call from Scoring Motor
'          2 = Call for 200ms Motor Cycle (normal cycle = 140ms)
'          3 = Call for 220ms Motor Cycle
'
Sub AddScore(ScoreToAdd, Motor)
  If InProgress=False Or Tilted=True Then
       Exit Sub                   'Disable score when TILTed or not started
    End If
    '
    ' Display the score
    '
    ' Note: "Scoring Motor" uses Th100sReelb.Timer
    '
    If (ScoreToAdd = 5) OR (ScoreToAdd=50)  Then
        Motor=1               ' No Sound at this level
       If ScoreToAdd=5 Then
          If MUnits < 5 Then
             MUnits = MUnits + 5      ' 5 = Pulse Units 5 Times
      Else
              Motor=1           ' only nest 1 deep on motor runs
          End If
       Else
          MTens = MTens + 5         ' 50 = Pulse Tens 5 Times
       End If
       If Th100sReelb.TimerEnabled=False Then
          '
          ' Scoring Motor Not Running, Start it up
          '
          Th100sReelb.TimerInterval=140   ' Normal cycle = 140ms per pulse
          Select Case Motor
             Case 2
                Th100sReelb.TimerInterval=200 ' 200ms per pulse
             Case 3
               Th100sReelb.TimerInterval=220  ' 220ms per pulse
          End Select
          MotorPhase=0
          Th100sReelb_Timer()         ' Score 1st Pulse Right Now
          Th100sReelb.TimerEnabled=True   ' Start Scoring Motor Running
       End If
    Else
       '
       ' Score is either 1 or 10, Score it Now
       Score = Score+ScoreToAdd           'Total of current score
       SReelHuns=  INT(Score/100)         'Get New Scoring Reel Digits
       SReelTens = INT((Score MOD 100) /10)
       SReelUnits = (Score MOD 10)
       DisplayScore()               ' Put Score on backglass
    End If
    '
    '  Check for replays Earned on Scoring
    '
  If Score=>Replay1 And Replay1Paid=False Then    'If the first replay score is reached,
       AddReplay()                    'then give a replay and
     Replay1Paid=True                 'mark the replay as paid
  End If
  If Score=>Replay2 And Replay2Paid=False Then    'Same as above,except for 2nd replay
     AddReplay()
     Replay2Paid=True
  End If
  If Score=>Replay3 And RePlay3Paid=False Then    'Same for 3rd replay
     AddReplay()
     Replay3Paid=True
    End If
  If Score=>Replay4 And RePlay4Paid=False Then    'Same for 4th replay
     AddReplay()
     Replay4Paid=True
  End If
    If Score=>Replay5 And RePlay5Paid=False Then    'Same for 5th replay
     AddReplay()
     Replay5Paid=True
  End If
'
'  Play Appropriate Bell sounds to match scoring
'
    If (Bells=True) AND (Motor=0) Then
       PlaySound ScoreToAdd               ' Play Bell Sound
    End if
    If ScoreToAdd=1 Then
       LightsRandomize()
    End If
End Sub
'
'  "Scoring Motor" Timer Routine
'
Sub Th100sReelb_Timer()
    MotorPhase=MotorPhase+1
    If MotorPhase=6 Then
       MotorPhase=0
       Exit Sub           ' skip phase 6
    End If
    If MUnits > 0 Then
       MUnits=Munits-1
       AddScore 1, 0        ' Score a units digit pulse
       Exit Sub
    End If
    If MTens > 0 Then
       MTens = MTens-1
       AddScore 10, 0       ' Score a Tens Digit pulse
       Exit Sub
    End If
    Th100sReelb.TimerEnabled=False  ' Stop Scoring Motor
End Sub
'
'--------------------------------------
'  Player Beat the High Score, Award It
'--------------------------------------
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
       If Credits < MaxCredits Then     'Credit Reel Max Check
          Credits=Credits+1         'Add a credit
          DisplayCredits()          'Update the credit display
       End If
    End If
    DelayedReplays=DelayedReplays+1
    HS10Thousand.TimerEnabled=True        'Enable the timer for Overlapping Replays
End Sub
'
'*******************************************************
'****  Multiple Replay Award With Knocker Sequence  ****
'****                                               ****
'****    HS10Thousand.TimerEnabled=True to start    ****
'*******************************************************
'
'  High Score Timer cycles every 250ms (1/4 sec)
'
'  4 cycles means it runs for 1 second when activated
'
'  Assumes 1 Replay and Knock done before timer started.
'  Delayed Replays are added at 250ms intervals
'  based on value DelayedReplays
'
Sub HS10Thousand_Timer()
  If DelayedReplays > 1 Then
       PlaySound"Knocker"           'Play the sound
       If Credits < MaxCredits Then       ' Our credit Reel maxes at 29
          Credits=Credits+1           'Add a credit
          DisplayCredits()            'Update the credit display
       End If
       DelayedReplays=DelayedReplays-1      'One Less To Go
    Else
       DelayedReplays=0           'No More Delayed Replays
     Me.TimerEnabled=False
    End If
End Sub
'
'******************************
'**** Tilt Bobber Simulator ***
'******************************
'
'  This routine is called each time game is Nudged
'  to Check to see if Player has tilted the game.
Sub CheckTilt()                     'Called when table is nudged
    Dim I
  Count1=Count1+1                   'Add to tilt count (hit lasts 1 second)
  If Count1>TiltSensitivity Then            'If more that Allowed Counts then TILTED
     Tilted=True
       If BallsOnField > 0 Then
          DelayedGOV=True               'Say game Is Effectively over
       End If
       DisplayTilt(1)                                'Put "TILT" On Backglss
       PlaySound "Reset1"
       GIOff()                      ' Gen Illum Off
       BumpersOff()                   ' Bumpers Off
       For I=0 to (LtTilt.Count-1)
          LtTilt(I).IsDropped=True            ' UnLight Playfield
       Next
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
'*************************************************
'****         Coin Deposited Sequence         ****
'****                                         ****
'****  CreditReel.Timer_Enabled=True to start ****
'*************************************************
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
     Case 9:AddCredit           'Always add at least 1 credit
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
'*************************************
'****  Game Start Reset Sequence  ****
'*************************************
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
           PlaySound "GameStart"          'Start the motor sound
           '
           Disable(True)              'Disable playfield objects
           BallID=1                 ' Reset Ball ID
           BallDrained(0)=0             ' None Drained Yet
           TroughOpening=True           ' Say Trough is Opening
           TroughOpenHalf.IsDropped=True      ' Animate Half Open Trough
           TroughOpenHalf.TimerEnabled=True     ' Start Timer to Complete Open
           For Each X in HoleTrigs
                X.Enabled=True            ' Enable Hole Triggers
           Next
 '
'  Cycle 2 (.4s after start) play a Reset sound and
'  Release the balls from the visible trough
    Case 2:
       Credits=Credits-1            'Subtract a credit
           DisplayCredits()             'Update the credit display
           DisplayTilt(0)                       ' Clear "TILT" On Backglass
           DisplayGOV(0)              'Clear "Game Over" on Back Glass
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
           ResetScoreToZero()       ' Reset Scoring Reels
           '
           '  Init game logic & Game Start Light States
           '
           GIOn()
           Made1to8=False         ' Mark special things not made yet
           MadeCOD=False
           MadeStar=False
           MadeShamrock=False
           LtCOD1.IsDropped=True      ' Unlight COD, Star, Shamrock & Special Lights
           LtCOD2.IsDropped=True
           LtCOD3.IsDropped=True
           LtStar.IsDropped=True
           LtOnStar.IsDropped=True
           LtShamrock.IsDropped=True
           LtOnShamRock.IsDropped=True
           LtSpecialOut.IsDropped=True
           LtExSpecialOut.IsDropped=True
           For X=1 to 8
              Spotted(X)=0          ' UnSpot 1 to 8
              EVAL("Lt"&X).IsDropped=True ' And UnLight on PlayField
           Next
           BumpersOn()            ' Bumpers All Init to On
           BumperOff(3)           ' Except 100K When Lit Bumper
           B3BaseL.state=0
           B3BaseL2.state=0
           '
           LightsRandomize()        ' Light next rotating item
'
'  Cycle 4 (.8s after start pressed) just generates delay
'
'  Cycle 5 (1 full second after start) We Reset the score counter
'
      Case 5:
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
           FirstOut=True
       Me.TimerEnabled=False        'Disable this timer
  End Select
End Sub
'
'  Routine to complete animation of Trough Opening/Closing
'
Sub TroughOpenHalf_Timer()
    TroughOpenHalf.IsDropped=True       ' Drop half open frame
    If TroughOpening=True Then
       BallRetain.IsDropped=True
       TroughOpen.IsDropped=False       ' Show ball Trough Open
       DrainPlayField()             ' Remove Any Balls on Playfield
    Else
       TroughOpen.IsDropped=True        ' Show ball Trough Closed
       For X=1 to 5
          EVAL("Hole"&X&"OpenFull").IsDropped=True ' Drop Empty Hole Open
          EVAL("Hole"&X&"Cover").IsDropped=False  ' Raise All Hole Covers
       Next
    End If
    TroughOpenHalf.TimerEnabled=False     ' Stop Timer
End Sub
'
'  Subroutine to Drain All Balls on the Playfield
'
Sub DrainPlayField()
'
'  Need to Animate Drop from holes
'  but for now just disappear them
'
    For X=1 to 5
       EVAL("Hole"&X&"Cover").IsDropped=True  ' Lower All Hole Covers
       If HoleHit(X) <> 0 Then
          LowerIntoHole(X-1)          ' Lower Ball into Hole to Drain
       Else
          EVAL("Hole"&X&"OpenFull").IsDropped=False ' Show Empty Hole Full Open
       End If
    Next
'
End Sub
'
'  Lower each ball into its'hole to drain
'
'  Num = Hole Number - 1 (Hole collection index)
'
Sub LowerIntoHole(Idx)
    Holes(Idx).TimerInterval=10       ' Set timer to 10ms
    Holes(Idx).UserValue=5+INT(RND()*20)  ' Initialize "Hang" Delay Counter (50-140ms)
  Holes(Idx).TimerEnabled=True      ' And Turn It On
End Sub
'
'  Timer Lowers Ball into Hole Until out of sight
'  Then does the Final Draining action
'
Sub Holes_Timer(Idx)
    Dim Z, Num                ' Define Local Variables
     '
    If Idx > 4 Then
       Kickers_Timer(Idx)         ' This is a kicker, not a trap hole
       Exit Sub               ' call that timer and get out
    End If
    '
    If Holes(Idx).UserValue > 0 Then
       '
       '  Still Doing "Hang Time" Delay, count one 10ms "hang" done
       '
       Holes(Idx).UserValue=Holes(Idx).UserValue-1
    End If
    Z=HolePtr(Idx).Z
  Z = Z - 3               'Subtract 3 from ball Z position and repeat the line above
  HolePtr(Idx).Z=Z            ' Move the ball Z position to the value of variable Z
  If Z > -35 Then             'If the ball Z position is above -25 cycle
       Exit Sub
    End If
'
'------------------------------------
' Hole Drain Finished, Close it out
'------------------------------------
'
    Holes(Idx).TimerEnabled=False       ' Stop this timer
    HoleClick(Idx).Enabled=False        ' Disable click detect in empty hole
    HoleTrigs(Idx).Enabled=True         ' Enable Hole Attraction
    Num=Idx+1                 ' Turn index into Hole Number 1-9
    EVAL("Hole"&Num&"Open").IsDropped=True    ' Lower Ball-in-Hole Cover
    EVAL("Hole"&Num).DestroyBall        ' Delete Ball From table
    PlaySound "GobbleFF"
    HoleHit(Num)=0                ' Mark Hole Empty
    EVAL("Hole"&Num&"OpenFull").IsDropped=False ' Show Empty Hole Full Open
End Sub
'
'
'**********************************
'****  Game Over Sequence  ********
'**********************************
'
' (Uses Timer associated with the "Millions" EM Reel
'  It runs at 100ms ticks when enabled to end Game)
'
Sub M1sReel_Timer()
    If Th100sReelb.TimerEnabled=True Then
       Exit Sub                   ' Wait for scoring motor to stop
    End If
    InProgress=False                'Game is finished
  GOVPhase=GOVPhase+1               'Increment the Cycle counter
  Select Case GOVPhase
    Case 2:
       DelayedGOV=False             'End Any Tilt Delay
'
'  Cases 2-4 are just 400ms of delay
'
'  0.5s after end of game
    Case 5:
'
'  Cases 6-15 are 900ms more delay
'
'   1.6s after end of game
    Case 16:
           DisplayGOV(1)              'Display GAME OVER
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
           DisableFlippers()            'In Case Holding Flippers Now
       GOVPhase=0               'Reset Phase for next time
       Me.TimerEnabled=False          'Disable this timer
           If DelayedStart=True Then
               DelayedStart=False         ' Start Pressed Before End Of game Finished
               Plunger.TimerEnabled=True      ' Start the "Game Start" Reset Sequence
           End if
  End Select
End Sub
'
'***** Subroutine to Display the Game Rules *****
'
Sub Rules()
  Msg(0)="                                       C.O.D."
  Msg(1)=""
    Msg(2)=""
  Msg(3)="  Balls In Holes C-O-D-* .................... 1 Replay."
  Msg(4)="  Balls In Holes C-O-D-& ................... 1 Replay."
  Msg(5)="  Balls In Holes C-O-D-*-& ................. 3 Replays"
    Msg(6)="  BALL THRU BOTTOM CENTER CHANNEL when"
  Msg(7)="  numbers 1 to 8 or Holes C-O-D are made ... 1 Replay"
    Msg(8)="  BALL THRU BOTTOM CENTER CHANNEL when"
    Msg(9)="  numbers 1 to 8 and Holes C-O-D are made .. 3 Replays"
  Msg(10)="  BALL THRU BOTTOM CENTER CHANNEL when"
  Msg(11)="  C-O-D and * or & are made ................ 3 Replays"
  Msg(12)=""
    Msg(13)="  Replays for 4,500,000, 5,000,000, 6,000,000, 7,000,000 and 8,000,000."
  Msg(14)=""
  Msg(15)="  Press 5 to Deposit Coin, 1 to Start a game."
    Msg(16)="  Press A to lift next ball into play."
    Msg(17)=""
  Msg(18)="  Press 2 to Reset high Score to 2,460,000."
  Msg(19)=""
  Msg(20)="  Press 3 to Reset Credits to 0."
    Msg(21)=""
    Msg(22)="  Press 4 to toggle Impulse or Modern Flippers."
    Msg(23)=""
    Msg(24)="  Press 6 to toggle standard or alternate post"
    Msg(25)="  rubbers at the auto-flipper."
  For X=1 To 25
    Msg(0)=Msg(0)+Msg(X)&Chr(13)
  Next
  MsgBox Msg(0),,"         Instructions and Rule Card"
End Sub
 '
'**************************************
'****  Handle Contact Switch Hits *****
'**************************************
'
' Contact Switch Hit
'
Sub StandUps_Hit(Idx)
    If StandUps(Idx).UserValue=True Then
       Exit Sub             ' Debounce rolling Ball
    End If
    StandUps(Idx).UserValue=True    ' Flag as debouncing this switch
     StandUps(Idx).TimerEnabled=True    ' start debounce timer
    AddScore 1, 0           ' Score 10,000
End Sub
'
Sub StandUps_Timer(Idx)
    StandUps(Idx).UserValue=False   ' End Debounce Window
    StandUps(Idx).TimerEnabled=False
End Sub
'
'******************************
'****  Handle Bumper Hits *****
'******************************
'
'  ---------------------------
'      Upper POP Bumper
'  ---------------------------
'
Sub B1Trig_Hit

  PlaySoundAtVol "Bumper", ActiveBall, 1
    AddScore 1, 0           ' 1 (10,000) points
    If B1State=0 Then
       B1Rings()                        ' Kick Off animation
       B1Ring3.TimerEnabled=True
    End If
End Sub
'
'
'  Bumper 1 Ring Animation Cycle Routine
'
Sub B1Rings()
    B1State=B1State+1
    Select Case B1State
      Case 1:
         B1Ring3.IsDropped=False
      Case 2:
         B1Ring2.IsDropped=False
      Case 3:
         B1Ring1.IsDropped=False
      Case 4:
         B1Ring2.IsDropped=False
      Case 5:
         B1Ring3.IsDropped=False
    End Select
End Sub
'
' Bumper 1 Ring Animation Timer
'
Sub B1Ring3_Timer()
    Select Case B1State
       Case 1:
          B1Ring3.IsDropped=True
       Case 2:
          B1Ring2.IsDropped=True
       Case 3:
          B1Ring1.IsDropped=True
       Case 4:
          B1Ring2.IsDropped=True
       Case 5:
          B1Ring3.IsDropped=True
    End Select
    If B1State<5 Then
       B1Rings()
    Else
       B1State=0
       B1Ring3.TimerEnabled=False
    End If
End Sub
'
'  ------------------------------
'     Lower POP Bumper
'  ------------------------------
'
Sub B2Trig_Hit

  PlaySoundAtVol "Bumper", ActiveBall, 1
    AddScore 1, 0           ' 1 (10,000) points
    If B2State=0 Then
       B2Rings()                        ' Kick off Ring Animation
       B2Ring3.TimerEnabled=True
    End If
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
' -------------------------------------
'   Top 10,000/100,000 Passive Bumper
' -------------------------------------
'
Sub B3Trig_Hit()
  'If B3Top.UserValue=False Then
     PlaySoundAtVol  "Rubber", ActiveBall, 1
       AddSound "RollMore"
       If B3BaseL.state=0 and B3BaseL2.state=0 Then
          AddScore 1, 0         ' 10,000 if Unlit
       Else
          AddScore 10, 0        ' 100,000 if Lit
       End If
  'End If
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
'
'**************************************************************
'    H A N D L E   T R A P   H O L E S  &  K I C K E R S
'**************************************************************
'
' Idx   Hole
 ' ---   ----------------
'  0    Hole "C"
'  1    Hole "O"
'  2    Hole "D"
'  3    Hole "Star"
'  4    Hole "Shamrock"
'  5    Kicker "4"
'  6    Kicker "5"
'  7    Kicker "6"
'  8    Kicker "7"
'  9    Kicker "8"
'
'---------------------------------------------------------
 '                             |
 '  Note: Both Types of holes are combined as above in   |
 '        the collections and are processed identically  |
 '        until the ball is grabbed.  After grabbing     |
 '        the ball, the code "forks" depending on    |
 '        hole type -- trap hole or kicker.        |
 '                             |
 '---------------------------------------------------------
'
Const AForce = 0.7          ' Hole attraction force
'
Sub HoleTrigs_Hit(Idx)
    Dim D, I, J, S
    If (Idx < 5) AND (HoleHit(Idx+1) = 1) Then
       Exit Sub           ' Ball already In This Trap hole
    End If
'
'-----------------------------------------
'  Empty Hole Trigger Zone Entered
'-----------------------------------------
'
    HoleCheat(Idx)=0
    HoleTicks(Idx)=0            ' Start counting timer ticks in trigger zone
    HoleTLow(Idx)=0
    Set aBall(Idx)=ActiveBall
    HoleTrigs(Idx).TimerInterval=10     ' start 10ms ticker
    HoleTrigs(Idx).TimerEnabled=True
    HForce(Idx)=AForce
    D=Distance(aBall(Idx), Holes(Idx))    ' distance from center of ball to center of hole
    I=aBall(Idx).VelY
    J=aBall(Idx).VelX
    S= SQR((I*I)+(J*J))           ' S = Scalar Speed of ball
    HoleTrigs(Idx).UserValue=S        ' Save initial ball speed
    If (D <= 35) Then
       D=AttractBall(aBall(Idx), Holes(Idx).X, Holes(Idx).Y, 85, 3)
    End If
End Sub
'
'  Hole Trigger Zone exited
'  without being grabbed by Kicker. Stop Attraction Timer
'
Sub HoleTrigs_UnHit(Idx)
    HoleTrigs(Idx).TimerEnabled=False
    Holes(Idx).Enabled=False
    AddSound "Roll9"
End Sub
'
'  ------------------------------------------------------
'     Timer to do attract every 10ms while Ball in zone
'  ------------------------------------------------------
'
Sub HoleTrigs_Timer(Idx)
    HoleTicks(Idx)=HoleTicks(Idx)+1       ' count another 10ms
    If (aBall(Idx).Y > Holes(Idx).Y) Then
       HoleTLow(Idx) = HoleTLow(Idx)+1      ' count # ticks below hole center
    Else
       HoleTLow(Idx)=0              ' Reset if ball above center
    End If
    Dim D, S, I, J
    D=Distance(aBall(Idx), Holes(Idx))      ' D = distance from center of ball to center of hole
    I=aBall(Idx).VelY
    J=aBall(Idx).VelX
    S= SQR((I*I)+(J*J))             ' S = Scalar Speed of ball
    If (S < HoleTrigs(Idx).UserValue) Then
       HoleTrigs(Idx).UserValue = S       ' Save minimum speed acheived
    End If
    If (D > 65) Then
       HoleTrigs_UnHit(Idx)           ' Handle fast ball...
       Exit Sub                 ' That misses UnHit
    End If
    If (HoleTicks(Idx) > 35) OR _
       ((HoleTicks(Idx) > 9) AND (aBall(Idx).VelY > 0)) Then
       '
       ' Ball below hole travelling down after 100ms, let it go
       '
       HoleTrigs_UnHit(Idx)           ' end ball attraction timer
       Exit Sub                 ' and get out
    End If
    If HoleCheat(Idx)=0 Then
       '
       ' HoleCheat Happens 1 Time on ball > 30 units from center
       ' of hole.  It's purpose is to simulate the beveled
       ' rim of the Siver Skates holes and deflect the ball
       ' around the edge of the hole.
       '
       If (D > 30) AND (aBall(Idx).VelY > 0) AND _
          (aBall(Idx).Y > Holes(Idx).Y) Then
          aBall(Idx).VelY=aBall(Idx).VelY/2   ' Ball Below Center moving down
          HoleCheat(Idx)=1
       End If
       If (D > 30) AND (aBall(Idx).VelY < 0) AND _
          (aBall(Idx).Y < Holes(Idx).Y) Then
          aBall(Idx).VelY= aBall(Idx).VelY/2  ' Ball Above Center moving up
          HoleCheat(Idx)=1
       End If
       If (D > 30) AND (aBall(Idx).VelX > 0) AND _
          (aBall(Idx).X > Holes(Idx).X) Then
          aBall(Idx).VelX= aBall(Idx).VelX/2.5  ' Ball Right of Center Moving Right
          HoleCheat(Idx)=1
       End If
       If (D > 30) AND (aBall(Idx).VelX < 0) AND _
          (aBall(Idx).X < Holes(Idx).X) Then
          aBall(Idx).VelX= aBall(Idx).VelX/2.5  ' Ball Left of Center Moving Left
          HoleCheat(Idx)=1
       End If
    End If
    If D > 35 Then
       Exit Sub
    End If
    D = AttractBall(aBall(Idx), Holes(Idx).X, Holes(Idx).Y, 85, HForce(Idx))
    If ((D < 14) AND (S < 12)) OR ((S < 1.2) AND (D < 20)) Then
       '
       ' Inside hole, let kicker grab ball
       '
       Holes(Idx).Enabled=True
       HoleTrigs(Idx).TimerEnabled=False    'Stop Centripital Force Timer
       aBall(Idx).VelX=0
       aBall(Idx).VelY=1
       aBall(Idx).X=Holes(Idx).X
       aBall(Idx).Y=Holes(Idx).Y-30       ' Position to roll into Kicker
    Else
       aBall(Idx).VelX=aBall(Idx).VelX*.95
       aBall(Idx).VelY=aBall(Idx).VelY*.95
    End If
End Sub
'
'  --------------------------
'    Hole Hit, Land the Ball
'  --------------------------
'
' Finish landing ball in hole, then decide if this is
 ' a Trap Hole or kicker and handle accordingly...
'
Sub Holes_Hit(Idx)
    HoleTrigs(Idx).TimerEnabled=False ' Stop Centripital Force Timer
    HoleCovers(Idx).IsDropped=True    ' Drop Hole Empty Cover
    HolesOpen(Idx).IsDropped=False    ' Raise "ball-in-Hole Cover
    If Idx > 4 Then
       DoKicker(Idx)          ' This hole is a kicker, process that
       Exit Sub             ' done here
    End If
    '
    '---------------------------------------------------
    ' Ball Was Trapped, Do Scoring & Drain
    '---------------------------------------------------
    '
    HoleClick(Idx).Enabled=True     ' Enable impact detection on Trap Holes
    Set HolePtr(Idx)=ActiveBall     ' Save pointer to ball trapped in this hole
    If (Tilted=True) Then
       Exit Sub
    End If
    SpotHole(Idx+1)           ' Say Ball in Hole
    AddScore 50,0           ' 500,000 Points
    '
    '  Do "Drain" for End Of  Game Action
    '
    BallsLost=BallsLost+1       ' Say One More Ball Lost
    BallsOnField=BallsOnField-1     ' and not on PlayField
    If Tilted=True Then         ' Game over if tilted
       InProgress=False         ' Game has finished (due to tilt)
       DisableFlippers()        ' In case Holding Flippers
       BallsLeft=0            ' No More Balls Left
       GOVPhase=1           ' Game Over Processing Has Begun
       M1sReel.TimerEnabled=True    ' Start Game Over Sequence
       Exit Sub
  End If
    If BallsLost >= 5 Then        ' Game is over after 5 Balls
       GOVPhase=1           ' Game Over Processing Has Begun
       If DisplayHighScore=True Then
      SetHS(HighScore)        ' Display the current HIGH SCORE
     End If
       M1sReel.TimerEnabled=True    ' Start Game Over Sequence
  End If
End Sub
'
 '-------------------------------------------------
'  Make click sound on trapped ball impact
 '-------------------------------------------------
'
Sub HoleClick_Hit(Idx)
    PlaySound "Click"
End Sub
'
'---------------------------------------------------
'   Do Ball Landed in Trap Hole Scoring/Processing
'---------------------------------------------------
'
Sub SpotHole(Num)
    HoleHit(Num)=1              ' Mark Hole Trapped a Ball
    If Num < 4 Then
       EVAL("LtCOD"&Num).IsDropped=False  ' Light C, O or D Light
       If (HoleHit(1)+HoleHit(2)+HoleHit(3) = 3) Then
          MadeCOD=True            ' Say C-O-D Made
          If MadeStar=True Then
             AddReplay()          ' C-O-D-* = 1 Replay
          End If
          If MadeShamrock=True Then
             AddReplay()          ' C-O-D-& = 1 Replay
          End If
          If (MadeStar=True) AND (MadeShamrock=True) Then
             AddReplay()          ' C-O-D-*-& = 3 Replays
          End If
          LtSpecialOut.IsDropped=False      ' Light Outlane Special
          If (Made1to8=True) OR (MadeStar=True) OR (MadeShamrock=True) Then
             LtExSpecialOut.IsDropped=False ' Light Outlane Extra Special
          End If
       End If
    End If
    If Num=4 Then
       MadeStar=True            ' Say Made Star
       If MadeShamrock=True Then
          BumperOn(3)
         B3BaseL.state=1       ' *-& = 100K When Lit Bumper On
         B3BaseL2.state=1        ' *-& = 100K When Lit Bumper On
       End If
       If MadeCOD=True Then
          AddReplay()           ' C-O-D-* = 1 Replay
          If MadeShamrock=True Then
             AddReplay()          ' C-O-D-*-& = 3 Replays
             AddReplay()
          End If
       End If
       LtStar.IsDropped=False       ' Light Star Lights
       LtOnStar.IsDropped=False
    End If
    If Num=5 Then
       MadeShamrock=True
       If MadeStar=True Then
          BumperOn(3)           ' *-& = 100K When Lit Bumper On
          B3BaseL.state=1
          B3BaseL2.state=1
       End If
       If MadeCOD=True Then
          AddReplay()           ' C-O-D-& = 1 Replay
          If MadeStar=True Then
             AddReplay()          ' C-O-D-*-& = 3 Replays
             AddReplay()
          End If
       End If
       LtShamRock.IsDropped=False     ' Light Shamrock Lights
       LtOnShamRock.IsDropped=False
    End If
End Sub
'
'============================================================
'  Do processing for a kicker (Ball has been landed above)
'============================================================
'
' Idx = 5 to 9 for kickers 1 to 5 (scoring numbers 4 to 8)
'
Sub DoKicker(Idx)
    Holes(Idx).TimerInterval = 560    ' Set delay to kickout
    Holes(Idx).TimerEnabled = True    ' Start Kickout Timer
    PlaySound "motor"
    If Tilted=True Then
       Exit Sub             ' out if tilted
    End If
    AddScore 5, 0           ' Score 50,000
    SpotNumber(Idx-1)         ' Spot This Number
End Sub
 '
'   Kicker Kickout Timer
'
'   Note: Kickers_Timer is call from Holes_Timer when
'         this is a kicker.  A bit convoluted, but
'         necessary since trap hole and kickers share
'         the same collections.  So when the interval
'         Holes(Idx).TimerInterval (set above) expires
'         we end up here like this was the timer
'         interrupt.
 '
' Idx = 5 to 9 for kickers 1 to 5
'
Sub Kickers_Timer(Idx)
    Dim Ang, S
    PlaySound "Kicker"
    Holes(Idx).TimerEnabled = False   ' Stop Timer
    HolesOpen(Idx).IsDropped=True   ' Lower "ball-in-Hole" Cover
    HoleCovers(Idx).IsDropped=False   ' Raise Hole Cover
    '
    '  Kicker kickout angles are:
    '
    '  Kicker 1 168 Degrees
    '  Kicker 2 170 Degrees
    '  Kicker 3 170 Degrees
    '  Kicker 4 170 Degrees
    '  Kicker 5 170 Degrees
    '
    '  Note: Idx = Kicker # - 1
    '
    S=RND()*2               ' Scatter Factor
    Select Case Idx
       Case 5
          Ang=168             ' Kicker 1 kicks at 168 Degrees
       Case 6
          Ang=167             ' Kicker 2 kicks at 167 Degrees
       Case 7
          Ang=175             ' Kicker 3 kicks at 170 Degrees
        Case 8
          Ang=162             ' Kicker 4 kicks at 162 Degrees
        Case 9
          Ang=170             ' Kicker 5 kicks at 170 Degrees
    End Select
    Holes(Idx).Kick (Ang+1)-S, 9      ' Kick out the Ball +/- 1 degrees
    Holes(Idx).Enabled=False        ' Disable kicker
End Sub
'
'-----------------------------------------------------
'  Function to find the distance from center of a ball
'  to the center of a hole.
'-----------------------------------------------------
'
Function Distance(aBall,Hole)
    Dim dX, dY
    '
    dX = aBall.X - Hole.X
    dY = aBall.Y - Hole.Y
    Distance = Sqr(dX*dX + dY*dY)
End Function
 '
'-----------------------------------------------------
'  Attraction Routine For Kicker Hole Enlargement Zone
'  and AutoFlipper Saucer
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
'*****************************************************************
'             E N D   H O L E   P R O C E S S I N G
'*****************************************************************
'
'**************************************
'        Handle AutoFlipper
'**************************************
'
'     ------------------------------------
'         Right AutoFlipper Routines
'     -------------------------------------
'
'  Right AutoFlipper 40 Unit Diameter Saucer Zone Entered
'
Sub RightAutoTrig_Hit()
    Dim D
    Set RaBall=ActiveBall         ' Grab Pointer to this ball
    RightAutoTrig.UserValue=0       ' Say Not Grabbed Yet
    D = AttractBall(RaBall, RightAutoTrig.X, RightAutoTrig.Y, 50, 3)
    If (D <= 60) Then
       RightAutoTrig.TimerInterval=10   'start 10ms updater
       RightAutoTrig.TimerEnabled=True
    End If
End Sub
'
'  Timer to do attract every 10ms while Ball in Saucer
'
Sub RightAutoTrig_Timer()
    Dim D
    If RightAutoTrig.UserValue=1 Then
       '
       ' Ball Landed in Saucer, just hold until Flipper activates
       '
       RaBall.X=RightAutoTrig.X
       RaBall.Y=RightAutoTrig.Y       ' Lock in Center of Saucer
       RaBall.VelX=0            ' and Stop Movement
       RaBall.VelY=0
       Exit Sub
    End If
    D = AttractBall(RaBall, RightAutoTrig.X, RightAutoTrig.Y, 50, 3)
    If D > 60 Then
       RightAutoTrig.TimerEnabled=False   ' outside hole
       Exit Sub
    End If
    If (D < 5) Then
       '
       ' Inside Saucer, Hold It There and Start Flipper Timer
       '
       RaBall.X=RightAutoTrig.X
       RaBall.Y=RightAutoTrig.Y       ' Lock in Center of Saucer
       RaBall.VelX=0            ' and Stop Movement
       RaBall.VelY=0
       RightAutoTrig.UserValue=1      ' Say Grabbed
'       RightAutoFlipper.TimerInterval=260+INT(RND()*50)  ' Flip after 260 -310ms delay
       RightAutoFlipper.TimerInterval=150
        RightAutoFlipper.TimerEnabled=True  ' Start Timer
    End If
End Sub
'
'  Right Auto Flipper 50 Unit Diameter Saucer Zone exited
'  without being grabbed (landing). Stop Attraction Timer
'
Sub RightAutoTrig_UnHit()
    RightAutoTrig.TimerEnabled=False
End Sub
'
'  Time to Activate the Right AutoFlipper
'
Sub RightAutoFlipper_Timer()
     RaBall.VelX=0.4+ (RND()*.3)
    RightAutoTrig.TimerEnabled=False    ' Stop the Saucer Hold Timer
    RightAutoFlipper.TimerEnabled=False   ' And this Flipper Timer
    RightAutoFlipper.RotateToEnd      ' Activate the Flipper
    PlaySound"FlipperUp"          ' Give it sound
    RAFTimer.Interval=220         ' End AutoFlip after 200ms
    RAFTimer.Enabled=True         ' Start timer
    AddScore 5, 0             ' 50,000 Pts
End Sub
'
'  Time to end AutoFlip
'
Sub RAFTimer_Timer()
    RAFTimer.Enabled=False          ' Stop This Timer
    RightAutoFlipper.RotateToStart      ' UnFlip the Flipper
    PlaySound"FlipperDown"          ' Give it sound
End Sub
 '
'**************************************
'        End AutoFlipper Code
'**************************************
'
 '
'**************************
'  Handle Rollover Lanes
'**************************
'
' Idx   Lane
'  0    Top "1" lane (100,000)
'  1    Top "2" Lane (100,000)
'  2    Top "3" Lane (100,000)
'  3    Lower Red 100,000 or 500,000 if Star Lit Lane
'  4    Lower Green 100,000 or 500,000 if Shamrock Lit Lane
'  5    Center "500,000, Special or Extra Special" Out Lane
'
Sub Lanes_Hit(Idx)
    Dim I
    Wires(Idx).IsDropped=False      ' Raise Wire Cover
    If Tilted=True Then
       Exit Sub
    End If
    PlaySound "KQSolenoid"
    Select Case Idx
       Case 0, 1, 2
           SpotNumber(Idx+1)        ' 100,000 & Spot Number
       Case 3
          If MadeStar=True Then
             ROV50Timer.Enabled=True  ' 500,000 When Lit
             Exit Sub         ' and not 100,000
          End If
       Case 4
           If MadeShamrock=True Then
             ROV50Timer.Enabled=True  ' 500,000 When Lit
             Exit Sub         ' and not 100,000
          End If
       Case 5
            If LtSpecialOut.IsDropped=False Then
               AddReplay()        ' Special When Lit
            End If
            If LtExSpecialOut.IsDropped=False Then
               AddReplay()        ' Extra Special When Lit
               AddReplay()
            End If
           ROV50Timer.Enabled=True
           Exit Sub
    End Select
    ROV10Timer.Enabled=True       ' Score 5 (50,000) Points after delay
End Sub
 '
'******************************************************
'  Subroutine to Spot A Number 1 to 8
'*******************************************************
'
Sub SpotNumber(Num)
    Dim I
    If (Tilted=True) OR (Spotted(Num)=1) Then
       Exit Sub               ' already spotted or Tilted = get out now
    End If
     Spotted(Num)=1             ' spot this number
    EVAL("Lt"&Num).IsDropped=False      ' Light this number
    For I=1 to 8
        If Spotted(I)=0 Then
           Exit Sub             ' 1 to 8 not made yet
        End If
     Next
     Made1to8=True              ' say made 1 to 8
     LtSpecialOut.IsDropped=False     ' Light Outlane Special
     If MadeCOD=True Then
        LtExSpecialOut.IsDropped=False    ' Light Extra Special
     End If
End Sub
'
'  ROV50Timer does 500,000 Point scoring after delay
'
Sub ROV50Timer_Timer()
    AddScore 50, 0                ' Rollovers Score 500,000 points
    ME.Enabled=False
End Sub
'
'  ROV5Timer does 50,000 scoring after delay
'
Sub ROV10Timer_Timer()
    AddScore 10, 0                ' Rollovers Score 100,000 points
    ME.Enabled=False
End Sub
'
'  Routine to animate Roll Over Wire
'
Sub Lanes_UnHit(Idx)
    Wires(Idx).TimerInterval=75         ' after alight delay...
    Wires(Idx).TimerEnabled=True        ' Show Rollover Wire again
End Sub
'
'  End of Roll Over Wire Animation Timer
'
Sub Wires_Timer(Idx)
    Wires(Idx).IsDropped=True       ' Drop Wire Cover
    Wires(Idx).TimerEnabled=False     ' End Animation
End Sub
'
 '--------------------------------------------
'  Routines to Turn Bumper Lights On and Off
 '--------------------------------------------
'
Sub BumpersOn()
    For X=1 to 3
       BumperOn(X)
    Next
End Sub
'
Sub BumpersOff()
    For X=1 to 3
       BumperOff(X)
    Next
End Sub
'
'  Turn Single Bumper Off
'
Sub BumperOff(Num)
    EVAL("B"&Num&"SideOn").IsDropped=True
    EVAL("B"&Num&"SideOff").IsDropped=False
 '   EVAL("B"&Num&"Base").State=LightStateOff
 '   EVAL("B"&Num&"Top").State=LightStateOff
    EVAL("B"&Num&"Top1Lit").IsDropped=True
    EVAL("B"&Num&"Top2Lit").IsDropped=True
    EVAL("B"&Num&"Top3Lit").IsDropped=True
End Sub
'
' Turn Single Bumper On
'
Sub BumperOn(Num)
    EVAL("B"&Num&"SideOn").IsDropped=False
    EVAL("B"&Num&"SideOff").IsDropped=True
  '  EVAL("B"&Num&"Base").State=LightStateOn
   ' EVAL("B"&Num&"Top").State=LightStateOn
    EVAL("B"&Num&"Top1Lit").IsDropped=False
    EVAL("B"&Num&"Top2Lit").IsDropped=False
    EVAL("B"&Num&"Top3Lit").IsDropped=False
End Sub
'
'  Routine to rotate lighted items
'
'  The items which rotate lit/unlit states are"
'
'
Sub LightsRandomize()
End Sub
'
'-----------------------------------------------------------
'  Routine to set the optional rubber mode by autoflipper
'-----------------------------------------------------------
'
Sub SetRubberMode()
    Dim I
     Select Case RubberMode
       Case 0
           '
           '  As shown in the release flyer
           '
           '  Rubber BETWEEN 5 & 6, ON 7
           For I=1 to 5
              EVAL("Post5R"&I).IsDropped=True   ' Drop ON 5
             EVAL("Post6R"&I).IsDropped=True    ' Drop ON 6
              EVAL("P67R"&I).IsDropped=True     ' Drop Between 6 & 7
             EVAL("Post7R"&I).IsDropped=False   ' Raise ON 7
              EVAL("P56R"&I).IsDropped=False      ' Raise Between 5 & 6
           Next
       Case 1
           '
           ' First alternate, Rubber ON 5, BETWEEN 6 & 7
           '
          For I=1 to 5
              EVAL("Post5R"&I).IsDropped=False    ' Raise ON 5
              EVAL("P56R"&I).IsDropped=True     ' Drop Between 5 & 6
             EVAL("Post6R"&I).IsDropped=True    ' Drop ON 6
             EVAL("Post7R"&I).IsDropped=True    ' Drop ON 7
              EVAL("P67R"&I).IsDropped=False      ' Raise Between 6 & 7
          Next
        Case 2
           '
           ' Second Alternate, Rubber ON 5, 6, & 7
           '
           For I=1 to 5
              EVAL("P56R"&I).IsDropped=True     ' Drop Between 5 & 6
              EVAL("P67R"&I).IsDropped=True     ' Drop Between 6 & 7
             EVAL("Post5R"&I).IsDropped=False   ' Raise ON 5
              EVAL("Post6R"&I).IsDropped=False    ' Raise ON 6
             EVAL("Post7R"&I).IsDropped=False   ' Raise ON 7
           Next
     End Select
End Sub
'
'-----------------------------------
'      Reset Scoring Lights to Zero
'-----------------------------------
'
'  Simulates Scoring Steppers and relays resetting
'
Sub ResetScoreToZero()
    Score=0
    If SReelTens=0 Then
       SReelTens=1            ' Assure 1st Relay Click "used"
    End If
    Arch.TimerInterval=120        ' Step Reset at 120ms/step
    Arch.TimerEnabled=True        ' Start the Reset
End Sub
'
'  ResetTimer does the reset 1 tick at a time
'
Sub Arch_Timer()
    X=True                ' Check for Score at Zero
    If SReelTens <> 0 Then        ' 100,000 Lights Just Reset
       SReelTens=0
       DisplayScore()         ' Update Lights w/new value
       Exit Sub             ' Done w/1st cycle if needed
    End If
    If SReelUnits <> 0 Then
       SReelUnits=SreelUnits+1      ' Step 10,000 Lights Up 1
       If SReelUnits >= 10 Then
          SReelUnits=0          ' Wrap back to 0
       End If
       X=False              ' Done for This Cycle
    End If
    If SReelHuns <> 0 Then
       SReelHuns=SreelHuns+1      ' Step Millions Lights Up 1
       If SReelHuns > 8 Then
          SReelHuns=0         ' Wrap to 0
       End If
       X=False              ' Don't Stop This Cycle
    End If
    If X=True Then
       Arch.TimerEnabled=False      ' Stop Reset, We're done
    End If
    DisplayScore()            ' Update Lights w/new value
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
 '
'=========================================================
'   B A C K G L A S S   D I S P L A Y   R O U T I N E S
'=========================================================
'
'
'********************************************************
'  Routine to Display Credits on Backglass
'********************************************************
'
Sub DisplayCredits()
    CreditReel.Setvalue Credits     ' Display Credits
     CreditReel2.Setvalue Credits   ' Mini Glass Too
     Controller.B2SSetCredits Credits
End Sub
'
'********************************************************
'  Routine to Display Game Over on Backglass
'********************************************************
'
' State=0 for Game Over OFF, 1 For Game Over ON
'
' No Game Over in this game
'
Sub DisplayGOV(State)
End Sub
'
'********************************************************
'  Routine to Display TILT on Backglass
'********************************************************
'
' State=0 for TILT OFF, 1 For TILT ON
'
Sub DisplayTilt(State)
    TiltReel.SetValue State
    Controller.B2SSetTilt 33,State
End Sub
'
'********************************************************
'  Routine to Display Match Number on Backglass
'********************************************************
'
' Num = -1 to turn match number off
'     = 0-9 to display that match number
'
' No match feature in this game
'
Sub DisplayMatch(Num)
End Sub
'
'****************************************
'  Routine to Display Score on Backglass
'****************************************
'
'  All timing etc. is handled elsewhere,
'  this just puts up the numbers on the Backglass
'
'  SReelHuns = Millions Digit  (0-9)
'  SReelTens = Hundred Thousands Digit (0-9)
'  SReelUnits = Ten Thousands Digit (0-9)
'
Sub DisplayScore()
    Showscore
    M1sReel.SetValue SReelHuns      ' Display Millions on large glass
     If SReelHuns = 1 Then
        M1sReelM.SetValue 1       ' 1 million shows on Mini Glass
     Else
        M1SReelM.SetValue 0       ' Off for other values
     End If
     '
    Th100sReel.SetValue SReelTens   ' Display Hundred Thousands on Left Side
    Select Case SReelTens       ' Display Hundred Thousands on Right
       Case 3
          Th100SReelb.SetValue 1    ' 300,000
        Case 5, 6
           Th100sReelb.SetValue SReelTens-3 ' 500,000 & 600,000
        Case 8, 9
           Th100sReelb.SetValue SReelTens-4 ' 800,000 & 900,000
        Case Else
           Th100sReelb.SetValue 0   ' other values don't display on Right Side
    End Select
     '
    If (SReelUnits = 1) OR (SReelUnits=0) Then
       Th10sReel.SetValue SReelUnits  ' 0 or 10,000 on Left Side Digit
       Th10sReel2.SetValue 0      ' Blank Right Side
    Else
       TH10sReel.Setvalue 0       ' Blank Left Side
       TH10sReel2.SetValue SReelUnits-1 ' Right Side Digit (20,000 - 90,000)
    End If
     If SReelUnits < 8 THen
        TH10sReelM.SetValue SReelUnits  ' Display 10,000 - 70,000 on Mini Glass
     Else
        TH10sReelM.SetValue 0     ' 80,000 & 90,000 don't display on Mini Glass
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
       Case 8:
          Controller.B2SSetData 77, 0
          Controller.B2SSetData 78, 1
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
'--------------------------------------------------



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


'
'         _    _ _ _ _ _                              
'        | |  | (_) | (_)                             
'        | |  | |_| | |_  __ _ _ __ ___  ___          
'        | |/\| | | | | |/ _` | '_ ` _ \/ __|         
'        \  /\  / | | | | (_| | | | | | \__ \         
'         \/  \/|_|_|_|_|\__,_|_| |_| |_|___/         
'                                                     
'                                                     
' _____             _               _           _     
'|_   _|           | |             (_)         | |    
'  | |_ __ __ _  __| | _____      ___ _ __   __| |___ 
'  | | '__/ _` |/ _` |/ _ \ \ /\ / / | '_ \ / _` / __|
'  | | | | (_| | (_| |  __/\ V  V /| | | | | (_| \__ \
'  \_/_|  \__,_|\__,_|\___| \_/\_/ |_|_| |_|\__,_|___/
'                                                     
'                                                     
'              __   _____  ____  _____                
'             /  | |  _  |/ ___|/ __  \               
'             `| | | |_| / /___ `' / /'               
'              | | \____ | ___ \  / /                 
'             _| |_.___/ / \_/ |./ /___               
'             \___/\____/\_____/\_____/               
'
'                                        
'Williams (1962 TradeWinds) version 1.4a for VP10
'VP10 table and complete playfield re-draw by wrd1972
'Original VP9 table and excellent foundation scripting by Pbecker
'VP10 conversion, pop-bumper cap primitives and "left flipper options menu" by Borgdog
'Flipper primitives and rubbers animations by Randr
'Lighting and shadow layer by Haunt Freaks
'Controller.vbs and DOF by Arngrim
'Screw prims by Zany
'High score support by "Borgdog"
'DB2S backglass by "Wildman"
'
'VP10.2 minimum required to play this table

'To access game options menu. Allow the table to start then hold down the "left flipper" button.
'
'************************************************************************************************************
'
Option Explicit			'Force variables to be declared


Dim DesktopMode: DesktopMode = Tradewinds.ShowDT
If DesktopMode = True Then 'Show Desktop components


Railleft.visible = 1
Railright.visible = 1
'Sidewalls_DT.visible = 1
'Sidewalls_FS.visible = 0


	Else

Railleft.visible = 0
Railright.visible = 0
'Sidewalls_DT.visible = 0
'Sidewalls_FS.visible = 1

End if

If DesktopMode = True Then 'Show Desktop components
			Sidewalls_DT.visible = 1
			Sidewalls_FS.visible = 0
		Else
			Sidewalls_FS.visible = 1
			Sidewalls_DT.visible = 0
		End If

 
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "trade_winds"

Dim operatormenu, options 'for controlling the options menu
Dim TableName			'For defining the table name to be used when saving credits/highscore
Dim InProgress			'TRUE while game is in progress, FALSE otherwise
Dim Tilted				'TRUE when tilted, FALSE otherwise
Dim TiltSensitivity		'Obvious
Dim State				'Used for playfield objects that can be disabled during TILT
Dim BallsLost			'Tracks the balls played
Dim BallsLeft			'Tracks the balls still to be lifted
Dim BallsLoaded         'Startup Trough Loading Counter
Dim BallAtPlunger       'True when a Ball is at the plunger, false when not
Dim BallsOnField        'Count of Balls Lifted and not yet drained
Dim BallsPerGame        'Count of Balls per game (3 or 5)
Dim DelayedBPG			'New Balls Per Game (changed while game in progress)
Dim BIT					'Ball in Virtual Lifter Tray when Draining Trough
Dim FirstOut            'Set when Ball First Exits Gate
Dim RollFlag            'Ball Roll Sound Semiphore
Dim RollVary            'Flag to Vary Roll Sound
Dim RollBack            'Flag to Detect Failure to leave gate on Plunger Release
Dim TC(6)               'Tracks Whether a Ball is in Trough Position 1-5 (0=No Ball)
Dim CreditsPerCoin		'1 to 5
Dim ScoreToAdd			'Used in the AddScore sub
Dim HighScore			'Holds previous high score value
Dim HighScorePaid		'Set to TRUE after beating high score
Dim HighScoreReward		'Set to the number of replays awarded when HIGH SCORE is beaten
Dim Credits				'Keeps track of the credits
Dim MaxCredits          'Maximum Credits Allowed
'
Dim GSPhase				'Cycle Counter Game Start reset sequence
Dim DelayedStart        'True if start pressed during Game Over Sequence
Dim DelayedGOV			'True Druing Interval From Tilt to Game Over Seq Started
Dim DelayedReplays		'Count Of Delayed Replays to Add with Knocker
Dim Count				'Cycle Counter for Credit_Timer
Dim Count1				'Used in TiltTimer for tilt sensitivity
Dim GOVPhase			'Cycle Counter for Game Over Sequence
Dim Match				'The match number
Dim MatchEnabled		'TRUE to enable number match, FALSE otherwise
Dim X					'Local Generic variable used in for-next loops
Dim Y					'Local Generic variable used in for-next loops
Dim Z					'Local Generic variable used in for-next loops
Dim objekt				'generic
Dim Lstep, Rstep		'counters for sling animation
Dim Kstep				'counter for kicker animation
'
Dim Score				' Current Score
Dim Flag100				' Flag used for the 100's Bell
'
Dim Replay1				'The scores needed for a replay are held in these 3 variables
Dim Replay2
Dim Replay3
Dim Replay4
Dim Replay1Paid			'Mark when each replay has been paid
Dim Replay2Paid
Dim Replay3Paid
Dim Replay4Paid
'
Dim Ones				'These are used in the score motor simulator
Dim Tens
Dim Hundreds
Dim Thousands
'
Dim Dinc
Dim DCycle				' Final Score Cycling Counter
Dim DisplayHighScore	'Don't display high score if there is no saved value
Dim Bells
Dim Msg(33)				' Used to display the rules
Dim MQ(30)				' De-Overlap up to 29 scoring motor events
 						' (We should never get more than 2 or 3 deep.)
 						'	
 						' MotorQueue Structure:
 						' MQ(0) = Index of Entry motor is now executing (0=none)
 						' MQ(1) = index (2-50) to queue next event
 						'   Note: If MQ(0)=0 motor is stopped.
 						'         Any other value, and it is running.
 						'   Motor Stops if current event finishes and
 						'   MQ(1)=MQ(0)+1.  It resets MQ(0)=0 and MQ(1)=2
 						'   and stops running.
Dim MotorPhase			' Which 120ms pulse running scoring motor is on
Dim MX					' Scoring Motor Temp variable
'
Dim BState(8)			' Bumper Ring Animation State Array
Dim ShootBall           ' Is animation Pulling, or Releasing?
Dim PS                  ' Pull Speed
Dim LRLast              ' Units Digit after Previous AddScore()
'
Dim aBall				' Kicker active Ball pointer
Dim Spotted(5)			' N-S-E-W Spotted Array
Dim MadeSpecial			' True if N-S-E-W has been made
Dim CurLit				' Advance Light Now Lit
Dim CKValue				' kicker Scoring Motor Parameter

Dim dw1step
Dim dw2step
Dim dw3step
Dim dw4step
Dim dw5step
Dim dw6step
'
Dim Debug



 ' ---------------------------------------
' *** Game Init -- Script Starts here ***
' ---------------------------------------
'
Sub TradeWinds_Init()
	LoadEM
	HSEnterMode=False
    FirstOut=False
    BallAtPlunger=False
    DelayedStart=False   	'No Delayed Anything yet
    DelayedGOV=False
    DelayedReplays=0
'
 	TableName=cGameName	'Place table name here
	CreditsPerCoin=1		'1 to 5 credits per coin
    MaxCredits = 29     	'Size of Credit Reel Graphic
	HighScoreReward=1		'The number of replays to award when HIGH SCORE is beaten (1-5)
'
    TiltTimer.Enabled=True
  	TiltSensitivity=4		'0-15, 4=normal - a higher number is less sensitive, 0=1 nudge and you're out
	MatchEnabled=True		'Match feature enabled
	Bells=True				'TRUE to play bell sounds, FALSE for no bells
'
'
    If BallsPerGame=5 Then
       Replay1=1300			'Place the 5 Ball score needed for replays in these variables
       Replay2=1500
	   Replay3=1700
       Replay4=1800
    Else
       Replay1=1300			'Place the 3 Ball score needed for replays in these variables
       Replay2=1500
	   Replay3=1700
       Replay4=1800
    End If
'
 	Initialize()			'Call the main Init sub
End Sub
'
'
'  -------------------------------------
'  Routine to Set High Score in HS Reels
'  -------------------------------------
'
Sub SetHS(HScore)
'    X=HScore MOD 10
'    HSUnits.SetValue X
'    X = INT((HScore MOD 100) /10)
'    HSTens.SetValue X
'    X = INT((HScore MOD 1000) / 100)
'    HSHundreds.SetValue X
'    X = INT((HScore MOD 10000)/1000)
'    HSThousands.SetValue X
'    X = INT(HScore/10000)
'    HS10Thou.SetValue X
End Sub
'
'  ---------------------------------
'  Enable/Disable Bumpers and Lights
'  ---------------------------------
'
    Sub Disable(State)						'Disable or enable objects when TILTed
         LeftSlingShot.Disabled=State
         RightSlingShot.Disabled=State
        '
        ' Disable/Enable Bumpers
        '
        For X=1 to 7
            EVAL("B"&X&"Top").HasHitEvent=NOT(State)
        Next
        '
        ' Kill Flippers
    	If State=True Then
               LeftFlipper.RotateToStart		'Move flippers to start position
    	   RightFlipper.RotateToStart
    	   DOF 101, DOFOff
    	   DOF 102, DOFOff
    	End If
    End Sub
'
'
'  **********************************
'  ***** Main program Init code *****
'  **********************************
'
Sub Initialize()
 	Randomize									' Seed random Number Generator
	Match=INT(RND()*10)							' Generate a random number 0-9 for the match counter to start with
    '
    Score=0
    BallsPerGame=5								' Assume 5 Balls per Game to start
    DelayedBPG=0								' No Change Pending
     GrowlTimer.UserValue=False					' init to Allow AddSound
    '
  	Count=0										'Set variables to a known state
	Count1=0
	GOVPhase=0									'Init Game Over Sequence Phase
	GSPhase=0									'Init Game Start Sequence Phase
    MQ(0)=0										'Init Scoring Motor Queue Empty
    MQ(1)=2										'Next Even Goes here
 	BallsLost=0									'Reset ball counter
	BallsLeft=0									'The number of balls not yet lifted
    BallsOnField=0								'No Balls Lifted to Playfield
 	InProgress=False							'Game not started
    HideOptions									'Do not show options menu

 	DisplayHighScore=True						'Display the high score
    '
	DisplayMatch(-1)							' Turn off Match Numbers
    GOVReel.SetValue 1                          ' Turn On "Game Over"
	for each objekt in gilights: objekt.state=1:next	'turn on GIlights
    If B2SOn then 
		Controller.B2ssetgameover 35, 1
		for each objekt in backdropstuff: objekt.visible = 0 : next
	End If
    '
    '  Set Initial Game Conditions:
    '
     '
    '   Bumpers init to Off
    '
    For X=3 to 7								' Init POP Bumpers
       BumperOff(X)
       ' Drop animation on Active Bumpers
'       Eval("B"&X&"Ring1").IsDropped=True
'       Eval("B"&X&"Ring2").IsDropped=True
'       Eval("B"&X&"Ring3").IsDropped=True
'       Eval("B"&X&"Ring3").TimerInterval=10		' animation Speed
       BState(X-1)=0							' Init Ring animation State
    Next
    For X=1 to 2								' Init Passive Bumpers
       BumperOff(X)
       Eval("B"&X&"Top").UserValue=False
    Next
     '
    For X=1 to 4
       Spotted(X)=0								' Reset N-E-S-W made
    Next
    DisplayNESW()								' On BackGlass
    '
    '
    '  Lights Off
    '
    KickerLit.state=0					' Kicker Not Lit
    '
    For X=1 to 26
       EVAL("Lt"&X).state=0			' All Voyage Lights Off
    Next
    LtSpecial1.state=0				' Including Specials
    LtSpecial2.state=0
    LtSpecial3.state=0
    LtSouth.state=0						' And South Light
    ButtonsOff()								' Buttons Off Too
    '
    '
    '   Create Balls in Trough and turn them loose to roll down
    '
    '   Init End Trough Stoppers
 	Stopper1.IsDropped=False					' Left End Stopper Up
    BIT=0										' No Ball In Trough yet
    For X=2 To 5
       Eval("Stopper"&X).IsDropped=True         ' The rest down
    Next
    BallsLoaded=0
    BallsLost=0									' Say Rolled in from Playfield
    PlaySound "Roll9"
    Gate1.TimerEnabled=True					' Timer will create 5 Balls
    '
    '  5 Balls in the visible trough
    '
        FlippersTimer.Enabled=True				' Start VP9 Primitive Timer
    '
    '  Init Saved Variables (Credits, HS, etc.)
    '
 	On Error Resume Next    'Needed if highscore/credits have never been saved in this table name
    BallsPerGame=CDbl(LoadValue(TableName,"BallsPer"))
    If BallsPerGame="" Then
        BallsPerGame=5							' no entry sets 5 Balls Per Game
    Else
 '       CreditsPerCoin=BallsPerGame				' Set correct Credits per 25 cents
    End If
    If BallsPergame <> 3 Then
       Balls3.visible=False					' Display 5 Balls Per game
    End If
	OptionBalls.image="OptionsBalls"&BallsPerGame
   	Credits=Cdbl(LoadValue(TableName,"Credits"))
	If Credits="" Then Credits=1
	If Credits > 0 Then	DOF 112, DOFOn    			'If there was no entry, then set to default value
    CreditReel.Setvalue Credits
    If B2SOn then Controller.B2ssetcredits Credits
	HighScore=Cdbl(LoadValue(TableName,"HighScore"))
	HSA1=Cdbl(LoadValue(TableName,"HSA1"))
	HSA2=Cdbl(LoadValue(TableName,"HSA2"))
	HSA3=Cdbl(LoadValue(TableName,"HSA3"))
	if HSA1="" then HSA1=0
	if HSA2="" then HSA2=0
	if HSA3="" then HSA3=0
	UpdatePostIt
	If HighScore="" Then HighScore=500			'If there was no entry, then set to default value
	Score=Cdbl(LoadValue(TableName,"Score"))	'Load score value from previous game
    If DisplayHighScore=True Then
'       SetHS(HighScore)							' Put On Display
    End If
 	If Score="" Then Score=455					'If there was no entry, then use default value
	DisplayScore								'Set score reels to previous score value
    '
End Sub
'
'*******************************************************
'****  Initial Ball Loading into Trough at Startup  ****
'*******************************************************
'  The timer is used to prevent balls being created
'  inside each other. This only happens at first time init
'
Sub Gate1_Timer()
    If BallsLoaded=BallsPerGame Then
       Gate1.TimerEnabled=False
       If DelayedStart=True Then
          DelayedStart=False					' Start Pressed Before Initial Loading
          Plunger.TimerEnabled=True				' Start the "Game Start" Reset Sequence
       End if 
       Exit Sub
    End If
'    If BallsLoaded < 4 Then
'       Stopper5.IsDropped=True					' Lower Safety Walls
'    End If
'    Stopper6.IsDropped=True
 	Drain.CreatesizedBall (23)						' Create Ball at Left End of return trough
	Drain.Kick 72, 9			' Kick it to " roll down the trough
    BallsLoaded=BallsLoaded+1
    BallsLost=BallsLost+1					'Say Rolled in from Playfield
    StopSound "Roll9"
    PlaySound "Roll9"
End Sub
'
' *********************************
' ***** Keys Pressed Handling *****
' *********************************
Sub TradeWinds_KeyDown(ByVal Keycode)						'The usual keycode stuff
	If InProgress=True And Tilted=False Then
    '
    ' Flipper and Nudge only when game in
    ' progress and not tilted
    '
		If Keycode=LeftFlipperKey Then
			LeftFlipper.RotateToEnd
 			PlaySound SoundFXDOF("FlipperUp",101,DOFOn,DOFContactors)
  		End If 
 
 		If Keycode=RightFlipperKey Then
			RightFlipper.RotateToEnd
 			PlaySound SoundFXDOF("FlipperUp",102,DOFOn,DOFContactors)
  		End If

 		If Keycode=LeftTiltKey Then
			Nudge 90, 2
			CheckTilt()
		End If

 		If Keycode=RightTiltKey Then
			Nudge 270, 2
			CheckTilt()
		End If

         If Keycode=MechanicalTilt then
           Tilted=True
           DelayedGOV=True		'Say game Is Effectively over
           TiltReel.SetValue 1                   'Put "TILT" On Backglss
           If B2SOn then Controller.B2ssettilt 33, 1
           PlaySound "Tilt"
           Disable(True)		'Disable slings, bumpers etc
         End If

 		If Keycode=CenterTiltKey Then
			Nudge 0, 2
			CheckTilt()
		End If

	  ' "StartGameKey" Activates Ball Lifter when InProgress
		If (Keycode=StartGameKey or Keycode=RightMagnaSave) And BallsLeft>0 and FirstOut=False And Not HSEnterMode=true Then
		   PlaySound SoundFXDOF("BallOut",113,DOFPulse,DOFContactors)
		   Lifter.TimerEnabled=True				'Activates the ball lifter Timer
		End If


    '
    ' End flipper and Nudge Qualifier
    '
 	End If

	If keycode=LeftFlipperKey and Inprogress<>True and OperatorMenu=0 then
		OperatorMenuTimer.Enabled = true
	end if

	If keycode=LeftFlipperKey and State = false and OperatorMenu=1 then
		Options=Options+1
		If Options=4 then Options=1
		playsound "reels"
		Select Case (Options)
			Case 1:
				Option1.visible=true
				Option3.visible=False
			Case 2:
				Option2.visible=true
				Option1.visible=False
			Case 3:
				Option3.visible=true
				Option2.visible=False
		End Select
	end if

	If keycode=RightFlipperKey and State = false and OperatorMenu=1 then
	  PlaySound "WOOD1"
	  Select Case (Options)
		Case 1:
'			if BallsPerGame=3 then
'				BallsPerGame=5
'				Balls3.visible=false
'			  else
'				BallsPerGame=3
'				Balls3.visible=true
'			end if
'			OptionBalls.image = "OptionsBalls"&BallsPerGame     
		Case 2:
			HighScore = 500
			SaveValue TableName,"HighScore",HighScore   'Save score value for next time
			If DisplayHighScore=True Then
'			   SetHS(HighScore)
			End If
			UpdatePostIt
		Case 3:
			OperatorMenu=0
			HideOptions
	  End Select
	End If

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

 	If Keycode=AddCreditKey Then
       PlaySound"Coin"
	   CreditReel.TimerEnabled=True
	End If

 	If Keycode=StartGameKey And Credits>0 and Inprogress<>True And Not HSEnterMode=true Then
        If (GOVPhase) <> 0 or (DelayedGOV=True) or (BallsLoaded <> BallsPerGame) Then
          DelayedStart=True 			'let End of game Sequence Finsh First
       Else
          Plunger.TimerEnabled=True		'Start the "Game Start" Reset Sequence
       End if
    End If

	If HSEnterMode Then HighScoreProcessKey(keycode)
 
    ' Either "R" or "I" displays rules
 	'If Keycode=23 or KeyCode = 19 Then
     '   Rules()
    'End If
'
End Sub

Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub DisplayOptions
	OptionsBack.visible = true
	Option1.visible = True
	OptionBalls.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
	operatormenu=0
End Sub

'
Sub SetBallsPerGame(Num)
    BallsPerGame=Num
    BallsLoaded=Num
'    CreditsPerCoin=Num
    SaveValue TableName,"BallsPer",BallsPerGame
    If BallsPerGame=5 Then
       Replay1=1300			'Place the 5 Ball score needed for replays in these variables
       Replay2=1500
       Replay3=1700
       Replay4=1800
    Else
       Replay1=1300			'Place the 3 Ball score needed for replays in these variables
       Replay2=1500
	   Replay3=1700
       Replay4=1800
    End If
End Sub
'
'  Routine to Delay Change Message 100ms
'
Sub Balls3_Timer()
    Me.TimerEnabled=False
  	MsgBox Msg(0),,"       Balls Per Game Changed"
End Sub
'
'  -----------------------
'  Flipper Tracking Timers
'  -----------------------
'
'
'  VP9 Version
'

Sub FlippersTimer_Timer()
    LF1.Roty = LeftFlipper.CurrentAngle-90

    RF1.Roty = RightFlipper.CurrentAngle+90
End Sub

Sub Flipperrubbers_Timer
    Lrubber.objrotz = LeftFlipper.CurrentAngle - 120
    Rrubber.objrotz = RightFlipper.CurrentAngle + 120 

'Sub Flipperrubbers_Timer()
   'Lrubber.Rotz = LeftFlipper.CurrentAngle-180
   'Rrubber.Rotz = RightFlipper.CurrentAngle+90
End Sub




'
' **********************************
' ***** Keys Released Handling *****
' **********************************
'
Sub TradeWinds_KeyUp(ByVal Keycode)

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

 	If Inprogress=True And Tilted=False Then

 		If Keycode=LeftFlipperKey Then
			LeftFlipper.RotateToStart
			PlaySound SoundFXDOF("FlipperDown",101,DOFOff,DOFContactors)
  		End If

 		If Keycode=RightFlipperKey Then
			RightFlipper.RotateToStart
			PlaySound SoundFXDOF("FlipperDown",102,DOFOff,DOFContactors)
  		End If

 	End If


	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plungerspring",0,1,0.25,0.25
	End If

' 	If Keycode=PlungerKey Then
'        ShootBall=True                              'Reverse plunger animation
'		Plunger.Fire
'        If BallAtPlunger = True Then
'           PlaySound "PlungerRoll"
'        Else
'           PlaySound"PlungerSpring"
'        End If
'	End If
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
	DOF 114, DOFPulse
    BallAtPlunger = False
     RollBack=0                        'Init Look for Roll-back Sound
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
Sub Gate_Hit()			              'Just for the sound
    RollBack=0                        ' raise Stopper at right end of ball trough
     DrainWall.UserValue=False		  ' Allow 1 DrainWall Hit/Ball
End Sub
'
' Hit gate going wrong way at top
Sub Rebound_Hit()
    If FirstOut=False Then
       PlaySound "Gate5"              ' Gate Bounce Sound
    End If
    FirstOut=False
End Sub'
'
'
'************************************************************
'****  Ball Trough, Lifter and Drain management routines ****
'************************************************************
'
'  Because in VP8 balls that stay together eventually become one ball
'  this whole annoying trough wall thing is needed. Sigh...
 '  In VP9, this action is disabled.
 '
'  The general idea works like this:
'
'  1. All Trough Stopper Walls drop at game start,
'     then go up from Right End as each ball returns to trough
'
'  2. Each Stopper position has a trigger under it so a ball
'     rolling over it when the wall is down can be detected.
'     The array TC(Ball#) keeps track of whether a ball is
'     currently in the trough at that position or not.
'
'  3. Because nudging during play can actually cause a ball to
'     penetrate a trough stopper wall, the following stopper routines
'     detect this and lower the walls long enough to let such
'     a ball roll back to it's proper position.  
'
'  4. Architecture of the trough is as follows:
'
'      Stoppers and triggers are numbered from
'      left to right as 6, 5, 4, 3, 2, 1 with 1 on the right.
'
'      Gate1 is the ball entry to the trough, and
'      Drain2 is ball exit from the trough to the
'      virtual "lifter tray"
'
'      Lift is the ball's exit from the "lifter tray" and
'      onto the Playfield.  It is activated by the player
'      pressing the "A" key. (Note: The player can lift
'      multiple balls onto the playfield at once if so
'      desired, to the limit of balls in the "lift tray").
'
'      Drain is the playfield exit, and its timer delays
'      the re-creation of the ball in Drain 2 to match the
'      sound of the ball exiting the playfield and transiting
'      the virtual path into the Trough.
'***********************************************************
'
'  Ball hit Drain (the real one out on the table)
'
Sub Drain_Hit()
	PlaySound"Out_Hole"
	DOF 115, DOFPulse
	BallsLost=BallsLost+1							' Say One More Ball Lost
    BallsOnField=BallsOnField-1						' and not on PlayField
    Drain.TimerInterval=250							' Set .25 sec delay
    Drain.TimerEnabled=True                         ' Start Delay to add Ball To Trough
	If Tilted=True Then								' Game over if tilted
'       GOVReel.SetValue 1							' Game Over if Tilted
 	   InProgress=False								' Game has finished (due to tilt)
       BallsLeft=0 									' No More Balls Left
       GOVPhase=1									' Game Over Has Begun
       Score1.TimerEnabled=True						' Start Game Over Sequence
  	   Exit Sub
	End If
	If BallsLost >= BallsPerGame Then								' Game is over after 5 Balls
       If DisplayHighScore=True Then
' 		  SetHS(HighScore)							' Display the current HIGH SCORE
 	   End If
       GOVPhase=1									' Game Over Has Begun
 	   Score1.TimerEnabled=True						' Start the Game Over Sequence
	End If
End Sub
'
' After Delay from Drain Hit, Drain Timer Creates
' A Ball to "roll into" the visible trough from
' the left end
'
Sub Drain_Timer()
'    If TC(5) = 0 Then
       ' Don't lower Wall 5 if four balls already in trough
'       Stopper5.IsDropped=True
'    End If
'    Stopper6.IsDropped=True    					' Always Lower wall by Kicker
	Drain.DestroyBall
	Drain.CreateSizedBall 23
	Drain.Kick 72, 9					' Kick it to " roll down the trough
    Drain.TimerEnabled=False                 	' Disable This Timer 'Till Next time
End Sub
'
'
' Drain2 is ball Rolling out of Right End of Trough to load
' virtual lifter tray at game start. Destroy here, as it
' will get re-created when it is lifted onto Playfield
'
Sub Drain2_Hit()
	Me.DestroyBall								'Another Ball in the virtual "Lifter Tray"
    BIT=BIT+1
    If Drain2.TimerEnabled=False Then
        Drain2.TimerInterval=3000				' wait for 3 secs without ball
        Drain2.Uservalue=1						' 1st ball from vis trough out
        Drain2.TimerEnabled=True
    Else
        Drain2.UserValue=1						' Another Ball Out
    End If
End Sub
'
Sub Drain2_Timer()
    If Drain2.UserValue = 0 Then
       Drain2.TimerEnabled=False				'Game Start Sequence Over
       Stopper1.IsDropped=False					'Block Trough Exit Again
       BIT=BallsPerGame-BallsOnField			'Make trough count correct
    Else
       Drain2.UserValue=0						'Wait for visible balls to drain
   End If
End Sub
'
'  ******* Ball Lifter Routine ***********
'
' Lift is the kicker to put a ball up to the plunger
' It's timer is set to 400ms (2/5 sec) after the "A" key is
' pressed to match the timing and sound of a lifter in action
'
Sub Lifter_Timer()
	Lifter.Createsizedball (25)							' Create a Ball on the kicker
	Lifter.Kick 250,4									' Kick it up to the plunger lane
	DOF 113, DOFPulse
	BallsLeft=BallsLeft-1								' One less Ball remains
    BallsOnField=BallsOnField+1							' One more is on the Playfield
    FirstOut=True                                       ' Next gate Hit is first for this ball
  	Me.TimerEnabled=False								' Stop the Lift Timer
End Sub
'
' Stopper Hit/Unhit routines to manage trough separation walls
' so it looks like they don't exist.
'
' The Trough has three overall operating modes as follows:
'
'  1. First Time startup Trough Loading. This is indicated by
'     BallsLoaded being less than BallsPergame.  Happens only at game init.
'     Ends with 3 or 5 Balls locked in Trough awaiting Game Start
'
'  2. Game start sequence where the balls are released to roll from
'     the right end into the (virtual) "lifter tray" under the playfield
'     (that tray only exists in program logic).  This state is indicated 
'     by Stopper1.IsDropped=True (i.e. Stopper 1 has  been dropped).
'     Ends when first Ball passes out of playfield entry gate
'     and the trough is locked to enter state 3.
'
'  3. A game is in progress, and as balls drain from field they
'     are being returned to the Trough from left end.  This is indicated
'     by Stopper1.IsDropped=False (i.e. Stopper1 has been raised)
'
	sub stopper1_slingshot
		stopper2.isdropped=False
	end sub

	sub stopper2_slingshot
		Stopper3.isdropped=False
		PlaySound("fx_collide"), 0, .0001
	end sub

	sub stopper3_slingshot
		Stopper4.isdropped=False
		PlaySound("fx_collide"), 0, .0001
	end sub

	sub stopper4_slingshot
		Stopper5.isdropped=False
		PlaySound("fx_collide"), 0, .0001
	end sub

	sub stopper5_slingshot
		PlaySound("fx_collide"), 0, .0001
	end sub


'**************************************************
'****************Animated Rubbers *****************
'**************************************************

sub dingwall1_hit
	Rdw1.visible=0
	RDW1a.visible=1
	dw1step=1
	Me.timerenabled=1
end sub

sub dingwall1_timer
	select case dw1step
		Case 1: RDW1a.visible=0: Rdw1.visible=1
		case 2:	Rdw1.visible=0: rdw1b.visible=1
		Case 3: rdw1b.visible=0: Rdw1.visible=1: me.timerenabled=0
	end Select
	dw1step=dw1step+1
end sub

sub dingwall2_hit
	Rdw2.visible=0
	RDW2a.visible=1
	dw2step=1
	Me.timerenabled=1
end sub

sub dingwall2_timer
	select case dw2step
		Case 1: RDW2a.visible=0: Rdw2.visible=1
		case 2:	Rdw2.visible=0: rdw2b.visible=1
		Case 3: rdw2b.visible=0: Rdw2.visible=1: me.timerenabled=0
	end Select
	dw2step=dw2step+1
end sub

sub dingwall3_hit
	Rdw3.visible=0
	RDW3a.visible=1
	dw3step=1
	Me.timerenabled=1
end sub

sub dingwall3_timer
	select case dw3step
		Case 1: RDW3a.visible=0: Rdw3.visible=1
		case 2:	Rdw3.visible=0: rdw3b.visible=1
		Case 3: rdw3b.visible=0: Rdw3.visible=1: me.timerenabled=0
	end Select
	dw3step=dw3step+1
end sub

sub dingwall4_hit
	Rdw4.visible=0
	RDW4a.visible=1
	dw4step=1
	Me.timerenabled=1
end sub

sub dingwall4_timer
	select case dw4step
		Case 1: RDW4a.visible=0: Rdw4.visible=1
		case 2:	Rdw4.visible=0: rdw4b.visible=1
		Case 3: rdw4b.visible=0: Rdw4.visible=1: me.timerenabled=0
	end Select
	dw4step=dw4step+1
end sub






sub dingwall5_hit
	Rdw5.visible=0
	RDW5a.visible=1
	dw5step=1
	Me.timerenabled=1
AddScore 1, 0	
end sub

sub dingwall5_timer
	select case dw5step
		Case 1: RDW5a.visible=0: Rdw5.visible=1
		case 2:	Rdw5.visible=0: rdw5b.visible=1
		Case 3: rdw5b.visible=0: Rdw5.visible=1: me.timerenabled=0
	end Select
	dw5step=dw5step+1
end sub








sub dingwall6_hit
	Rdw6.visible=0
	RDW6a.visible=1
	dw6step=1
	Me.timerenabled=1
AddScore 1, 0	
end sub

sub dingwall6_timer
	select case dw6step
		Case 1: RDW5a.visible=0: Rdw6.visible=1
		case 2:	Rdw6.visible=0: rdw6b.visible=1
		Case 3: rdw6b.visible=0: Rdw6.visible=1: me.timerenabled=0
	end Select
	dw6step=dw6step+1
end sub


'
'
'************************************************
'****  End of the trough management routines ****
'************************************************
' 
' ******************************
' **** Scoring and credits *****
' ******************************
'
'  AddScore is main scoring routine
'
'  ScoreToAdd May be:
'      1, 10, 100 or 300
'
'  All values other than 1, 10 and 100 initiate scoring motor
'  for proper timing.
'
'  Sound is:
'     0 = Normal Sound for points scored
'     1 = 10 point + Rollover Solenoid Sound
'     2 = No Sound, just score
'
Sub AddScore(ScoreToAdd, Sound)
 	If InProgress=False Or Tilted=True Then
        Exit Sub										'Disable score when TILTed or not started
    End If
 	Score = Score + ScoreToAdd							'Total of current score
 	DisplayScore()										'Display the score
'
'  Check for replays Earned on Scoring
'
 	If Score=>Replay1 And _
       Replay1Paid=False Then							'If the first replay score is reached,
       AddReplay()										'then give a replay and
 	   Replay1Paid=True									'mark the replay as paid
	End If
	If Score=>Replay2 And _
       Replay2Paid=False Then							'Same as above,except for 2nd replay
	   AddReplay()
	   Replay2Paid=True
	End If
	If Score=>Replay3 And _
       RePlay3Paid=False Then							'Same for 3rd replay
	   AddReplay()
	   Replay3Paid=True
    End If
 	If Score=>Replay4 And _
       RePlay4Paid=False Then							'Same for 4th replay
	   AddReplay()
	   Replay4Paid=True
	End If
'
' If 1 points, Toggle Lights as needed
'
    If ScoreToAdd=10 Then
       LightsRandomize(0)
    End If
'
'  Play Appropriate Bell sounds to match scoring
'
    If Bells=True Then
       If ScoreToAdd=1 OR ScoreToAdd=10 OR ScoreToAdd=100 Then
          Select Case Sound
             Case 0:
				 If ScoreToAdd = 100 Then
					PlaySound SoundFXDOF("100",143,DOFPulse,DOFChimes)
				  elseif ScoreToAdd = 10 Then
					PlaySound SoundFXDOF("10",142,DOFPulse,DOFChimes)
				  Else
					PlaySound SoundFXDOF("1",141,DOFPulse,DOFChimes)
				End If				
             Case 1:
                PlaySound "10ptSwitch"
             Case 2:
                '
                 ' No Sound to be Played here
                 '
          End Select
       Else
          '
          ' Scoring Motor Sound/Timing Event
          '
          If MQ(0)=0 Then
             ' Scoring Motor Not Running, start it
             MotorPhase=1				' 1st 120ms phase
			 If ScoreToAdd = 100 Then
				PlaySound SoundFXDOF("100",143,DOFPulse,DOFChimes)
			  elseif ScoreToAdd = 10 Then
				PlaySound SoundFXDOF("10",142,DOFPulse,DOFChimes)
			  Else
				PlaySound SoundFXDOF("1",141,DOFPulse,DOFChimes)
			End If	
             MQ(0)=2					' Motor on 1st Queue Entry
             MQ(1)=3					' This is Next queue Entry
             MQ(2)=ScoreToAdd			' So it knows how many are active
             ThScore1.TimerInterval=135	' 135ms/pulse
             ThScore1.TimerEnabled=True	' Start the motor
          Else
             X=MQ(1)					' Get next Queue Entry
             MQ(1)=X+1					' Say it's used
             MQ(X)=ScoreToAdd			' Store Event for motor to find
          End If
       End If
 
    End If
End Sub
'
'  This is the scoring motor
'
' A full cycle is 5 120ms pulses or 700ms
'
Sub ThScore1_Timer()
    MX=MQ(0)						'Get our queue index
    ' Decrement another active pulse (if any)
    If MQ(MX) > 0 Then
       If MQ(MX) < 10 Then
          MQ(MX)=MQ(MX)-1			' Doing 1's, say one more bell out
       Else
          If MQ(MX) < 50 Then
             MQ(MX)=MQ(MX)-10		' Doing 10's, say one more out
          Else
             MQ(MX)=MQ(MX)-100		' Doing 100's, say one more out
          End If
       End If
    End If
    MX=MQ(MX)						' Get # Bells Left This cycle
    MotorPhase=MotorPhase+1
    If MotorPhase>=6 Then
       '
       ' Completing a cycle, see what's next (if anything)
       '
       MQ(0)=MQ(0)+1				' Point to next queue entry
       If MQ(0)=MQ(1) Then
          '
          ' Queue empty, stop motor
          '
          MQ(1)=2					' Mark queue empty
          MQ(0)=0					' And Motor Stopped
          MotorPhase=0				' reset motor phase
          ThScore1.TimerEnabled=False ' And stop it
       Else
          MotorPhase=1				' 1st 120ms phase for this entry
          MX=MQ(0)					' get next entry index
          PlaySound MQ(MX)			' play the sound
       End If
   End If   
End Sub
'
'  Player Beat the High Score, Award It
'
Sub PayHigh()											'When HIGH SCORE is beaten, go here
    For X=1 to HighScoreReward
 	   AddReplay()										'Always add at least 1 replay
    Next
End Sub
'
' Made a Replay, Award it and "Knock"
'
Sub AddReplay()
    If DelayedReplays = 0 Then
       PlaySound SoundFXDOF("Knocker",111,DOFPulse,DOFKnocker)	'Play the sound
	   DOF 123, DOFPulse
       If Credits < MaxCredits Then				'Credit Reel Max Check
          Credits=Credits+1						'Add a credit
          CreditReel.Setvalue Credits			'Update the credit display
          If B2SOn then Controller.B2ssetcredits Credits
       End If
    End If
    DelayedReplays=DelayedReplays+1
    HSTens.TimerEnabled=True					'Enable the timer for Overlapping Replays
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
       PlaySound SoundFXDOF("Knocker",111,DOFPulse,DOFKnocker)
	   DOF 123, DOFPulse						'Play the sound
       If Credits < MaxCredits Then				' Our credit Reel maxes at 29
          Credits=Credits+1
		  DOF 112, DOFOn						'Add a credit
          CreditReel.Setvalue Credits			'Update the credit display
          If B2SOn then Controller.B2ssetcredits Credits
       End If
       DelayedReplays=DelayedReplays-1			'One Less To Go
    Else
       DelayedReplays=0							'No More Delayed Replays
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
Sub CheckTilt()											'Called when table is nudged
	Count1=Count1+1										'Add to tilt count (hit lasts 1 second)
	If Count1>TiltSensitivity And Tilted=False Then		'If more than Allowed Counts then TILTED
	   Tilted=True
       DelayedGOV=True									'Say game Is Effectively over
       TiltReel.SetValue 1                              'Put "TILT" On Backglss
       If B2SOn then Controller.B2ssettilt 33, 1
       PlaySound "Tilt"
  	   Disable(True)									'Disable slings, bumpers etc
	End If
End Sub
'
' Tilt Timer Cycles every 250ms (1/4 sec) when game is running
'
Sub TiltTimer_Timer()				'Used to simulate a tilt bob
	If Count1>0 Then
       Count1=Count1-0.25			'Subtract .25 every 250ms
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
		Case 9:AddCredit						'Always add at least 1 credit
		Case 10:
          If CreditsPerCoin>3 Then AddCredit
		  Count=0
		  Me.TimerEnabled=False
	End Select
End Sub
'
' Add a Credit on the Backglass counter
Sub AddCredit()											'Called by Credit_Timer
    If Credits < MaxCredits Then
       Credits=Credits+1
	   DOF 112, DOFOn								'Add a credit to the counter
       CreditReel.Setvalue Credits						'Update the credit display
       If B2SOn then Controller.B2ssetcredits Credits    
	End If
  	PlaySound"StepUp"									'Play stepping sound
End Sub
'
' *************************************
' ****  Game Start Reset Sequence  ****
' *************************************
'
' When started, Plunger Timer gets here every 200ms until we are done.
' Entire startup process is staged to take 2.6 seconds to
' emulate an EM machine Game Start Reset sequence
'
Sub Plunger_Timer()									'Game Start Reset sequence
 	GSPhase=GSPhase+1								'Add 1 every time the counter cycles
	Select Case GSPhase								' Select Cycle to do this tick
'
'  On Cycle 1 (.2s after start pressed) play EM motor sound 
 		Case 1:
           PlaySound"GameStart2"						'Start of Game sound
           Disable(True)							'Disable playfield objects
			playsound "rollmore"


		   Credits=Credits-1
           If Credits < 1 Then DOF 112, DOFOff						'Subtract a credit
           CreditReel.Setvalue Credits				'Update the credit display
 		   GOVReel.SetValue 0						'Clear GAME OVER message
			If B2SOn then 
				   Controller.B2ssetcredits Credits
				   Controller.B2ssetgameover 35, 0
				   Controller.B2ssetdata 1, 0
				   Controller.B2ssetdata 2, 0
				   Controller.B2ssetdata 3, 0
				   Controller.B2ssetdata 4, 0
				   Controller.B2ssetdata 5, 0 
				   Controller.B2ssettilt 33, 0
			End If

           Score=0									' Reset Score
           Flag100=0								' And 100 crossing flag
           TiltReel.SetValue 0                      ' Clear "TILT" On Backglass
           BIT= (BallsPerGame - Ballslost)						' Number of Balls *not* in Trough now
           BallsLost=0								' Reset balls used
           BallsLeft=BallsPerGame					' And Balls Left to Lift/Play
           If BallsOnField > 0 Then
              BallsLeft = BallsLeft-BallsOnField	' Start Hit Mid-previous Game
           End If
'
 		   Replay1Paid=False						' Reset High Scores Hit flags
		   Replay2Paid=False
		   Replay3Paid=False
           Replay4Paid=False
		   HighScorePaid=False
'
 		   InProgress=True							'Game has started
		   Tilted=False								' and isn't yet tilted
		   If HighScore=0 Then						'Make sure we don't pay for 1st ever game
			  HighScorePaid=True
			  DisplayHighScore=False
		   End If	
	       DisplayMatch(-1)							' Turn off Match Numbers
  	       Match=INT(RND()*10)						'Generate 0-9 for the match this Game


		Case 2:
           DisplayScore()						'Reset score counter
 		   If HighScore>0 AND DisplayHighScore = True Then
'		      SetHS(HighScore)						'Display highscore
		   End If


'  Cycle 2 (.4s after start) play a Reset sound and
'  Release the balls from the visible trough

		Case 14:
			PlaySound "Rollmore"				
	'		Stopper1.isdropped=True
 		   For X=1 To 5
			  Eval("Stopper"&X).IsDropped=True		'Release the balls
		   Next
           If BIT=BallsPergame Then
              Stopper1.IsDropped=False				'None To Release
           End If
 '
'  Cycle 3 (.6s after start pressed) Set up lights
'
        Case 15:
  		   '
 		   '  Set Initial Game Start Light States:
 		   '
							' Kicker Not Lit
           KickerLit.state=0
           LtSouth.state=1				' South Light On
           '
           '    Special Lights Off
  		   '
           LtSpecial1.state=0
           LtSpecial2.state=0
           LtSpecial3.state=0
           '
           '  Voyage Lights Off
           '
           For X=2 to 26
              EVAL("Lt"&X).state=0
           Next
           Lt1.state=1					' Light Start in San Francisco
           CurLit=1								' Reset Advance Pointer
           '
           ButtonOn(1)							' Light "On" buttons
           ButtonOn(2)
           ButtonOff(3)							' UnLight Off Button
           ButtonOff(4)
           ButtonOn(5)							' Light Advance Button
           BumpersOff()							' All Bumpers Off
           BumperOn(1)							' Except top Left Advance
           BumperOn(4)							' And Center POP Bumper
           '
           For X=1 to 4
              Spotted(X)=0						' Reset N-E-S-W made
           Next
           DisplayNESW()						' On BackGlass
           MadeSpecial=False					' Special Made too
           '
           LightsRandomize(1)					' Check for initial 1 pt lighting
'
'  Cycle 5 (1 full second after start) We Reset the score counter
'

'

' Cycle 13 (2.6 seconds after start pressed)
'       End the startup Reset sequence and allow the game to start
'
 		Case 27:
           GSPhase=0							'Reset Phase for next time
		   Disable(False)						'Enable any disabled objects
           InProgress=True
  		   Me.TimerEnabled=False				'Disable this timer
	End Select
End Sub
'
'
' **********************************
' ****  Game Over Sequence  ********
' **********************************
'
' (Uses Timer associated with the scoring EM Reel
'  It runs at 100ms ticks when enabled to end Game)
'
Sub Score1_Timer()
	GOVPhase=GOVPhase+1								'Increment the Cycle counter
	Select Case GOVPhase
 		Case 2:
 		   DelayedGOV=False							'End Any Tilt Delay
'
'  Cases 1-4 are just 400ms of delay
'
'  0.5s after end of game
		Case 3:
            PlaySound "GameOver"					'Play end of game sound
		Case 6:
'            BumpersOff()							' UnLight Bumpers
 '            ButtonOff(1)							' UnLight 1 Pt Bumpers
 '            ButtonOff(2)
'
'  Cases 4-10 are more delay
'
'   1.6s after end of game
 		Case 10:
           If MatchEnabled=True AND Tilted=False Then	'If the match feature is enabled 
 			  If Match=(Score Mod 10) Then
                 AddReplay							'If numbers match, give a replay
              End If
               DisplayMatch(Match)					' Put Match Number on Back Glass
  		   End If
           GOVReel.SetValue 1                       'Display GAME OVER
           If B2SOn then Controller.B2ssetgameover 35, 1
'
' Cases 11 is end of game
'
' 2.1s after End Of game sequence began we have the final cycle
 		Case 11:
           savehs
		   InProgress=False							'Game is finished
           If DelayedBPG <> 0 Then
               SetBallsPerGame(DelayedBPG)			'Change Balls Per game for Future
               DelayedBPG=0
            End If
            Disable(True)							' Kill Flippers & Bumpers
 		   GOVPhase=0								'Reset Phase for next time
		   Me.TimerEnabled=False					'Disable this timer
           If DelayedStart=True Then
              DelayedStart=False					' Start Pressed Before End Of game Finished
              Plunger.TimerEnabled=True				' Start the "Game Start" Reset Sequence
           End if 
 	End Select
End Sub

sub savehs
		SaveValue TableName,"Score",Score		'Save Player score value for next time
		SaveValue TableName,"Credits",Credits	'Save the credits
        If Score > HighScore Then
			SaveValue TableName,"HighScore",Score	'Save the HighScore
			HighScore=Score
 		    If DisplayHighScore=True Then
'			   SetHS(HighScore)
            End If
			HighScoreEntryInit()
        End If
		SaveValue TableName,"HSA1",HSA1	'Save the hi score initials
		SaveValue TableName,"HSA2",HSA3	'Save the hi score initials
		SaveValue TableName,"HSA3",HSA2	'Save the hi score initials
end Sub
'
' ***** Subroutine to Display the Game Rules *****
'
Sub Rules()
	Msg(0)=""
	Msg(1)="                               TRADE WINDS"
	Msg(2)=""
	Msg(3)="  Red eject hole scores special when Wake, Samoa, or Tahiti is lit."
	Msg(4)=""
	Msg(5)="  Top rollover lane scores special when Tahiti is Lit."
	Msg(6)=""
	Msg(7)="  Side Bottom Rollover lanes lite, alternately, for special when North, South,"
	Msg(8)="  East, and West are Lit."
	Msg(9)=""
	Msg(10)="  1 Replay for matching last number in point score to lited number that"
    Msg(11)="  appears on back glass when game is over."
    Msg(12)=""
    Msg(13)=" Replays for 1,300, 1,500, 1,700 and 1,800 Points."
    Msg(14)=""
    Msg(15)="  Press 5 To Deposit Coin."
 	Msg(16)="  Press A to lift next ball into play"
    Msg(17)=""
    Msg(18)="  Press 1 To Start a Game."
    Msg(19)=""
    Msg(20)="  Press 2 to Reset High Score to 500."
	Msg(21)=""
	Msg(22)="  Press 3 to Reset Credits to 0."&CHR(13)&CHR(13)&"  Press 4 to Toggle 3 or 5 Ball Game."
    Msg(23)=""
 	For X=1 To 23
		Msg(0)=Msg(0)+Msg(X)&Chr(13)
	Next
	MsgBox Msg(0),,"         Instructions and Rule Card"
End Sub
'
' *********************************************
' **** Handle Things That Just Make Sounds
' *********************************************
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
Sub LeftRail_Hit()
    AddSound "Wood1"
End Sub
'
Sub Divider_Hit()
    AddSound "Wood1"
End Sub
'
Sub DrainWall_Hit()
    AddSound "Wood1"
End Sub

'
' *********************************************
' **** Handle Slingshot and Contact Switch Hits
' *********************************************
'
'  -----------------------
'  Ball hit Left Slingshot
'  -----------------------
'
Sub LeftSlingShot_SlingShot()
	PlaySound SoundFXDOF("Bumper",103,DOFPulse,DOFContactors)
	DOF 116, DOFPulse
    AddScore 1, 0					' 10 points
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -6
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -18
		Case 2:LSLing2.Visible = 0:LSLing3.Visible = 1:sling2.TransZ = -28
		Case 3:LSLing3.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -18
		Case 4:LSLing2.Visible = 0:LSLing1.Visible = 1:sling2.TransZ = -18
        Case 5:sling2.TransZ = 0:LSLing1.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'
'  ------------------------
'  Ball hit Right Slingshot
'  ------------------------
'
Sub RightSlingShot_SlingShot()
	PlaySound SoundFXDOF("Bumper",104,DOFPulse,DOFContactors)
	DOF 117, DOFPulse
	AddScore 1, 0					' 10 points
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -6
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -18
		Case 2:RSLing2.Visible = 0:RSLing3.Visible = 1:sling2.TransZ = -28
		Case 3:RSLing3.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -18
		Case 4:RSLing2.Visible = 0:RSLing1.Visible = 1:sling2.TransZ = -18
        Case 5:sling2.TransZ = 0:RSLing1.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'
'  ----------------------------
'  Ball Hit LSW1 Contact Switch
'  ----------------------------
'
Sub LSW1_Hit
'    If LSW1(0).UserValue=True Then
'       Exit Sub
'    End If
'    LSW1(0).UserValue=True			' Start Debounce Window
'    LSW1(0).TimerEnabled=True
'    X=ActiveBall.X
'    Y=ActiveBall.Y
'    If Y < 540 AND X > 85 AND X < 185 Then
       AddScore 10, 0				' hit Contact Switch Area
'    Else
'       Rubbers_Hit(0)				' hit outside contact area
'   End If
End Sub
'
'Sub LSW1_Timer(Idx)
'    LSW1(0).UserValue=False			' End Debounce Window
'    LSW1(Idx).TimerEnabled=False	' stop this timer
'End Sub
 '
'  ----------------------------
'  Ball Hit LSW2 Contact Switch
'  ----------------------------


'
'Sub LSW2_Hit
'Sub LSW2_Hit():PlaySound "stepup":End Sub
'    If LSW2(0).UserValue=True Then
'       Exit Sub
'    End If
'    LSW2(0).UserValue=True			' Start Debounce Window
'    LSW2(0).TimerEnabled=True
'    Y=ActiveBall.Y
'    If Y > 945 AND Y < 1015 Then
       'AddScore 1, 0				' hit Contact Switch Area
'    Else
'       Rubbers_Hit(0)				' hit outside contact area
'   End If
'End Sub
'
'Sub LSW2_Timer(Idx)
'    LSW2(0).UserValue=False			' End Debounce Window
'    LSW2(Idx).TimerEnabled=False	' stop this timer
'End Sub
'
'  ----------------------------
'  Ball Hit RSW1 Contact Switch
'  ----------------------------
'
Sub RSW1_Hit
'    If RSW1(0).UserValue=True Then
'       Exit Sub
'    End If
'    RSW1(0).UserValue=True			' Start Debounce Window
'    RSW1(0).TimerEnabled=True
'    X=ActiveBall.X
'    Y=ActiveBall.Y
'    If Y < 540 AND X > 780 AND X < 880 Then
       AddScore 10, 0				' hit Contact Switch Area
'    Else
'       Rubbers_Hit(0)				' hit outside contact area
'    End If
End Sub
'
'Sub RSW1_Timer(Idx)
'    RSW1(0).UserValue=False			' End Debounce Window
'    RSW1(Idx).TimerEnabled=False	' stop this timer
'End Sub
 '
'  ----------------------------
'  Ball Hit RSW2 Contact Switch
'  ----------------------------
'
'Sub RSW2_Hit
'    If RSW2(0).UserValue=True Then
'       Exit Sub
'    End If
'    RSW2(0).UserValue=True			' Start Debounce Window
'    RSW2(0).TimerEnabled=True
'    Y=ActiveBall.Y
'    If Y > 945 AND Y < 1015 Then
       'AddScore 1, 0				' hit Contact Switch Area
'    Else
'       Rubbers_Hit(0)				' hit outside contact area
'   End If
'End Sub
'
'Sub RSW2_Timer(Idx)
'    RSW2(0).UserValue=False			' End Debounce Window
'    RSW2(Idx).TimerEnabled=False	' stop this timer
'End Sub
'
' ******************************
' ****  Handle Bumper Hits *****
' ******************************
 '
' ----------------------------
'    Left Top Passive Bumper
' ----------------------------
'
Sub B1Top_Hit
    If B1Top.UserValue=False Then
       PlaySound "Rubber"
       If BTopLight1.State=LightStateOn Then
          DoAdvance()					' Advance When Lit
        End If
        AddScore 10, 0				' 10 points UnLit
    End If
    '
    '  Debounce Ball "Laying" on Passive Bumper
    '
    B1Top.TimerEnabled=False			'Stop Timer If Running
    B1Top.UserValue=True				'Say Debounce Started
    B1Top.TimerInterval=100				'Debounce for 100ms
    B1Top.TimerEnabled=True				'Start Debounce Window
End Sub
'
Sub B1Top_Timer()
    B1Top.TimerEnabled=False			'Stop Timer
    B1Top.UserValue=False				'Say Hit Counts Again
End Sub
 '
' ----------------------------
'   Right Top Passive Bumper
' ----------------------------
'
Sub B2Top_Hit
    If B2Top.UserValue=False Then
       PlaySound "Rubber"
       If BTopLight2.State=LightStateOn Then
          DoAdvance()					' Advance When Lit
        End If
        AddScore 10, 0					' 10 points
    End If
    '
    '  Debounce Ball "Laying" on Passive Bumper
    '
    B2Top.TimerEnabled=False			'Stop Timer If Running
    B2Top.UserValue=True				'Say Debounce Started
    B2Top.TimerInterval=100				'Debounce for 100ms
    B2Top.TimerEnabled=True				'Start Debounce Window
End Sub
'
Sub B2Top_Timer()
    B2Top.TimerEnabled=False			'Stop Timer
    B2Top.UserValue=False				'Say Hit Counts Again
End Sub
'
'  -------------------------
'     Left Red Pop Bumper
'  -------------------------
'
Sub B3Top_Hit
	B7Top.PlayHit
 	PlaySound SoundFXDOF("Bumper",105,DOFPulse,DOFContactors)
	DOF 118, DOFPulse
    If BTopLight3.State=LightStateOn Then
       AddScore 10, 0				' Score 10 Points
    Else
       AddScore  1, 0				' Score 1 Point
    End If
'    BRings(3)						' Kick off Ring Animation
'    B3Ring3.TimerEnabled=True
End Sub

 '
'  ------------------------
'    Right Red Pop Bumper
'  ------------------------
'
Sub B7Top_Hit
	B3Top.PlayHit
 	PlaySound SoundFXDOF("Bumper",109,DOFPulse,DOFContactors)
	DOF 122, DOFPulse
    If BTopLight7.State=LightStateOn Then
       AddScore 10, 0				' Score 10 Points
    Else
       AddScore  1, 0				' Score 1 Point
    End If
'    BRings(7)						' Kick off Ring Animation
'    B7Ring3.TimerEnabled=True
End Sub
'
'  -------------------------
'    Center Blue Pop Bumper
'  -------------------------
'
Sub B4Top_Hit
 	PlaySound SoundFXDOF("Bumper",106,DOFPulse,DOFContactors)
	DOF 121, 2
    If BTopLight4.State=LightStateOn Then
       AddScore 10, 0				' Score 10 Points
    Else
       AddScore  1, 0				' Score 1 Point
    End If
'   BRings(4)						' Kick off Ring Animation
'    B4Ring3.TimerEnabled=True
End Sub
 '
'  ---------------------------
'    Right Yellow Pop Bumper
'  ---------------------------
'
Sub B5Top_Hit
	B6Top.PlayHit
 	PlaySound SoundFXDOF("Bumper",107,DOFPulse,DOFContactors)
	DOF 120, DOFPulse
    If BTopLight5.State=LightStateOn Then
       AddScore 10, 0				' Score 10 Points
    Else
       AddScore  1, 0				' Score 1 Point
    End If
'    BRings(5)						' Kick off Ring Animation
'    B5Ring3.TimerEnabled=True
End Sub
 '
'  --------------------------
'    Left Yello Pop Bumper
'  --------------------------
'
Sub B6Top_Hit
	B5Top.PlayHit
 	PlaySound SoundFXDOF("Bumper",108,DOFPulse,DOFContactors)
	DOF 119, DOFPulse
    If BTopLight6.State=LightStateOn Then
       AddScore 10, 0				' Score 10 Points
    Else
       AddScore  1, 0				' Score 1 Point
    End If
'    BRings(6)						' Kick off Ring Animation
'    B6Ring3.TimerEnabled=True
End Sub

'
'  ----------------------------------------
'  Bumper Pop Ring Animation Cycle Routine
'  ----------------------------------------
'
Sub BRings(Num)
    X=Num-1
    BState(X)=BState(X)+1
    Select Case BState(X)
       Case 1:
          BRing3(X).IsDropped=False
       Case 2: 
          BRing2(X).IsDropped=False
       Case 3:
          BRing1(X).IsDropped=False
       Case 4:
          BRing2(X).IsDropped=False
       Case 5:
          BRing3(X).IsDropped=False
    End Select
End Sub
'
' Bumper POP Ring Animation Timer
'
Sub BRing3_Timer(Idx)
    Select Case BState(Idx)
       Case 1:
          BRing3(Idx).IsDropped=True 
       Case 2:
          BRing2(Idx).IsDropped=True
       Case 3:
          BRing1(Idx).IsDropped=True
       Case 4:
          BRing2(Idx).IsDropped=True
       Case 5:
          BRing3(Idx).IsDropped=True
    End Select
    If BState(Idx) < 5 Then 
       BRings(Idx+1)
    Else
       BState(Idx)=0
       BRing3(Idx).TimerEnabled=False
    End If
End Sub
'
'******************************************************
'  Handlers for Rollover Lane Triggers Hit
'******************************************************
'
'  We have the following Lanes:
'
'  0 - Top Left "East" Lane
'  1 - Top Center "North" Lane
'  2 - Top Right "West" Lane
'  3 - Lower Left "Special" OutLane
'  4 - Lower Left "South" OutLane
'  5 - Lower Left "West" Outlane
'  6 - Lower Right "East" OutLane
'  7 - Lower Right "South" OutLane
'  8 - Lower Right "Special" OutLane
'
'  The Lower Right "5" OutLane
'
Sub Lanes_Hit(Idx)
	DOF 200 + idx, DOFPulse
    If Idx < 3 Then
       ROV30()							' Score 30 Points
    Else
       ROV100()							' Score 100 Points
    End If
    Select Case Idx
       Case 0, 6
          ' Spot EAST
          Spotted(2)=1
          If B2SOn then Controller.B2ssetdata 3, 1 
         CheckSpecial()
       Case 1
          ' Spot NORTH
          Spotted(1)=1
          If B2SOn then Controller.B2ssetdata 1, 1
          CheckSpecial()
          If LtSpecial2.state=1 Then
             AddReplay()				' Top Center lane Special When Lit
          End If
       Case 2, 5
          ' Spot WEST
          Spotted(4)=1
          If B2SOn then Controller.B2ssetdata 4, 1
          CheckSpecial()
       Case 3
          If LtSpecial1.state=1 Then
             AddReplay()				' Left outlane Special When Lit
          End If
       Case 4, 7
          ' Spot SOUTH
          Spotted(3)=1
          If B2SOn then Controller.B2ssetdata 2, 1
          CheckSpecial()
       Case 8
          If LtSpecial3.state=1 Then
             AddReplay()				' Left outlane Special When Lit
          End If
    End Select
End Sub
'
'  Sub to do Rollover sound, delay 140ms then score 30 points
'
Sub ROV30()
    PlaySound "KQSolenoid"				' Rollover relay sound
    Wheel1.TimerInterval=10				' 40 ms delay
    Wheel1.TimerEnabled=True
End Sub
'
Sub Wheel1_Timer()
    Wheel1.TimerEnabled=false			' Stop Delay Timer
    KickerScore(30)						' top lanes 30 points
End Sub
'
'  Sub to do Rollover sound, delay 140ms then score 100 points
'
Sub ROV100()
    PlaySound "KQSolenoid"				' Rollover relay sound
    Wheel2.TimerInterval=140			' 140 ms delay
    Wheel2.TimerEnabled=True
End Sub
'
Sub Wheel2_Timer()
    Wheel2.TimerEnabled=false			' Stop Delay Timer
    AddScore 100, 0						' Outlanes 100 points
End Sub
'
'  Check for N-E-S-W all made
'
Sub CheckSpecial()
    If MadeSpecial=True Then
       Exit Sub							' special already made
    End If
    DisplayNESW()						' Display NEWS on Back Glass
    For X=1 to 4
       If Spotted(X)=0 Then
          Exit Sub						' at least one not made yet
       End If
    Next
    MadeSpecial=True					' N-E-S-W all made
    If BTopLight2.State=LightStateOn Then
       LtSpecial1.state=1		' Light outlane special
    Else
       LtSpecial3.state=1		' Opposite Top Advance Bumpers
    End If
End Sub


'******************
'Switch animations
'*******************




'***switch 4 animation***

Sub Lane4Trig2_Hit:Lane4Trig2.timerenabled = true:End Sub
Sub Lane4Trig2_UnHit:End Sub

Const Lane4Trigmin = 0
Const Lane4Trigmax = -20
Dim Lane4Trigdir
Lane4Trigdir = -2

Sub Lane4Trig2_timer()
 pRollover4.ObjRotX = pRollover4.ObjRotX + Lane4Trigdir
	If pRollover4.ObjRotX >= Lane4Trigmin Then
		Lane4Trig2.timerenabled = False
		pRollover4.ObjRotX = Lane4Trigmin
		Lane4Trigdir = -2
	End If
	If pRollover4.ObjRotX <= Lane4Trigmax Then
		Lane4Trigdir = 4
	End If
End Sub


'***switch 5 animation***

Sub Lane5Trig2_Hit:Lane5Trig2.timerenabled = true:End Sub
Sub Lane5Trig2_UnHit:End Sub

Const Lane5Trigmin = 0
Const Lane5Trigmax = -20
Dim Lane5Trigdir
Lane5Trigdir = -2

Sub Lane5Trig2_timer()
 pRollover5.ObjRotX = pRollover5.ObjRotX + Lane5Trigdir
	If pRollover5.ObjRotX >= Lane5Trigmin Then
		Lane5Trig2.timerenabled = False
		pRollover5.ObjRotX = Lane5Trigmin
		Lane5Trigdir = -2
	End If
	If pRollover5.ObjRotX <= Lane5Trigmax Then
		Lane5Trigdir = 4
	End If
End Sub

'***switch 6 animation***

Sub Lane6Trig2_Hit:Lane6Trig2.timerenabled = true:End Sub
Sub Lane6Trig2_UnHit:End Sub

Const Lane6Trigmin = 0
Const Lane6Trigmax = -20
Dim Lane6Trigdir
Lane6Trigdir = -2

Sub Lane6Trig2_timer()
 pRollover6.ObjRotX = pRollover6.ObjRotX + Lane6Trigdir
	If pRollover6.ObjRotX >= Lane6Trigmin Then
		Lane6Trig2.timerenabled = False
		pRollover6.ObjRotX = Lane6Trigmin
		Lane6Trigdir = -2
	End If
	If pRollover6.ObjRotX <= Lane6Trigmax Then
		Lane6Trigdir = 4
	End If
End Sub

'***switch 7 animation***

Sub Lane7Trig2_Hit:Lane7Trig2.timerenabled = true:End Sub
Sub Lane7Trig2_UnHit:End Sub

Const Lane7Trigmin = 0
Const Lane7Trigmax = -20
Dim Lane7Trigdir
Lane7Trigdir = -2

Sub Lane7Trig2_timer()
 pRollover7.ObjRotX = pRollover7.ObjRotX + Lane7Trigdir
	If pRollover7.ObjRotX >= Lane7Trigmin Then
		Lane7Trig2.timerenabled = False
		pRollover7.ObjRotX = Lane7Trigmin
		Lane7Trigdir = -2
	End If
	If pRollover7.ObjRotX <= Lane7Trigmax Then
		Lane7Trigdir = 4
	End If
End Sub

'***switch 8 animation***

Sub Lane8Trig2_Hit:Lane8Trig2.timerenabled = true:End Sub
Sub Lane8Trig2_UnHit:End Sub

Const Lane8Trigmin = 0
Const Lane8Trigmax = -20
Dim Lane8Trigdir
Lane8Trigdir = -2

Sub Lane8Trig2_timer()
 pRollover8.ObjRotX = pRollover8.ObjRotX + Lane8Trigdir
	If pRollover8.ObjRotX >= Lane8Trigmin Then
		Lane8Trig2.timerenabled = False
		pRollover8.ObjRotX = Lane8Trigmin
		Lane8Trigdir = -2
	End If
	If pRollover8.ObjRotX <= Lane8Trigmax Then
		Lane8Trigdir = 4
	End If
End Sub

'***switch 9 animation***

Sub Lane9Trig2_Hit:Lane9Trig2.timerenabled = true:End Sub
Sub Lane9Trig2_UnHit:End Sub

Const Lane9Trigmin = 0
Const Lane9Trigmax = -20
Dim Lane9Trigdir
Lane9Trigdir = -2

Sub Lane9Trig2_timer()
 pRollover9.ObjRotX = pRollover9.ObjRotX + Lane9Trigdir
	If pRollover9.ObjRotX >= Lane9Trigmin Then
		Lane9Trig2.timerenabled = False
		pRollover9.ObjRotX = Lane9Trigmin
		Lane9Trigdir = -2
	End If
	If pRollover9.ObjRotX <= Lane9Trigmax Then
		Lane9Trigdir = 4
	End If
End Sub


'
'******************************************************
'    Handler the Center Kicker
'******************************************************
'
Const AForce = 1.2					' Center Kicker attraction force
'
'
'======================================
'  Center Kicker Trigger Zone Entered
'======================================
'
'Sub KickerTrig_Hit()
'    Set aBall=ActiveBall			' Grab Pointer to this ball
'    If (AttractBall(aBall, Kicker.X, Kicker.Y, 65, 3) > 20) Then
'        KickerTrig.TimerInterval=10	'start 10ms updater
'        KickerTrig.TimerEnabled=True
'    End If
'    AddSound "RollMore"
'End Sub
''
''  Timer to do attract every 10ms while Ball in zone
''
'Sub KickerTrig_Timer()
'    Dim J
'    J=AttractBall(aBall, Kicker.X, Kicker.Y, 65, AForce)
'    If (J < 17) Then
'       '
'       ' Inside hole, let kicker grab ball
'       '
'       Kicker.Enabled=True
'       aBall.VelX=0
'       aBall.VelY=1
'       aBall.X=Kicker.X
'       aBall.Y=Kicker.Y-30				' Posiiton to roll into Kicker
'    Else
'       aBall.VelX=aBall.VelX*.95
'       aBall.VelY=aBall.VelY*.95
'    End If
'End Sub
''
''  Center Kicker 70 Unit Diameter Zone exited
''  without being grabbed by Kicker. Stop Attraction Timer
''
'Sub KickerTrig_UnHit()
'    KickerTrig.TimerEnabled=False
'    Kicker.Enabled=False
'    AddSound "Roll9"
'End Sub
'
'----------------------------------------
'  Actual Center Kicker Hole Entered
'----------------------------------------
'  
Sub Kicker_Hit()
'    KickerTrig.TimerEnabled=False		' Stop Attraction Timer
    If KickerLit.state=1 Then
       Kicker.UserValue=1				' Say Kicker Is Lit
       KickerLit.state=0				' Lower Lit Cover
       AddReplay()						' Award Center Kicker Special When Lit
    Else
       Kicker.UserValue=0				' Say Kicker Not Lit
       KickerLit.state=1		' Lower Lit Cover
    End If
    If LtSouth.state=1 Then
       Spotted(3)=1         		 	' Spot SOUTH When Lit
       If B2SOn then Controller.B2ssetdata 2, 1
       CheckSpecial()
    End If
    Kicker.TimerInterval = 175			' Set delay to kickout
    KickerScore(60)						' 30 Pts & 3 advances
    KScore.Enabled=True					'Start Scoring Timer
	Kstep=1								'kicker animation reset
    Kicker.TimerEnabled = True			'Start Kickout Timer
End Sub
'
' Center Kicker Kickout Timer
'
Sub Kicker_Timer()
	Select Case Kstep
		Case 1:
			PlaySound SoundFXDOF("Kicker",110,DOFPulse,DOFContactors)
			DOF 123, 2

			If Kicker.UserValue=1 Then
			   KickerLit.state=1		' Raise Lit Cover
			Else
			   KickerLit.state=0		' Raise UnLit Cover
			End If
			Pkickarm.rotz=15
			Kicker.Kick 190, 6					' Kick out at 195 degrees
		Case 3:
			Pkickarm.rotz=0
			Kicker.TimerEnabled = False			' Stop Timer
	end Select
	Kstep=kstep+1
End Sub
'
'  Start Kicker Scoring Motor for specific Points
'
'  Pts may be 30 or 60 (30 points + 3 advances) Only
'
Sub KickerScore(Pts)
    Dim Sound
    Select case Pts
       Case 30
          CKValue=0						' 30 Pts
          KScore.UserValue=3
       Case 60
          CKValue=1						' 30 Pts + 3 advances
          KScore.UserValue=3
    End Select
    KScore.Enabled=True					'Start Kicker Scoring Motor
End Sub
'
' Kicker Scoring Motor Timer
'
Sub KScore_Timer()
    Select Case CKValue
       Case 0
          AddScore 10, 0			' Another 10 pts
       Case 1
          AddScore 10, 0			' Another 10 pts
          DoAdvance()				' plus advance
       Case Else
          AddScore 100, 0			' Another 100 pts
    End Select
    KScore.UserValue=KScore.UserValue-1
    If KScore.Uservalue <= 0 Then
       KScore.Enabled=False			' Done, stop timer
    End If
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
    AttractBall=dist				' Return Distance to center
    If dist > Size Or dist < 1 Then
       Exit Function 				'Just to be safe
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
'  ***************************
'          Handle Buttons
'  ***************************
'
'   Trigger Hits for Rollover Buttons
'


Sub But1Trig_Hit
	ButtonOff(1)						' Swap ON and OFF Button Lit
    ButtonOn(3)
    BumperOn(5)         				' On Yellow Bumpers
    BumperOn(6)
    PlaySound "Reels"
	But1Prim.TransY=-5
End Sub

Sub But1Trig_UnHit
	But1Prim.TransY=0
End Sub

Sub But2Trig_Hit
          ButtonOff(2)						' Swap ON and OFF Button Lit
          ButtonOn(4)
          BumperOn(3)         				' On Red Bumpers
          BumperOn(7)
    PlaySound "Reels"
	But2Prim.TransY=-5
End Sub

Sub But2Trig_UnHit
	But2Prim.TransY=0
End Sub

Sub But3Trig_Hit
          ButtonOn(1)						' Swap ON and OFF Button Lit
          ButtonOff(3)
          BumperOff(5)         				' Off Yellow Bumpers
          BumperOff(6)
    PlaySound "Reels"
	But3Prim.TransY=-5
End Sub

Sub But3Trig_UnHit
	But3Prim.TransY=0
End Sub

Sub But4Trig_Hit
          ButtonOn(2)						' Swap ON and OFF Button Lit
          ButtonOff(4)
          BumperOff(3)         				' Off Red Bumpers
          BumperOff(7)
    PlaySound "Reels"
	But4Prim.TransY=-5
End Sub

Sub But4Trig_UnHit
	But4Prim.TransY=0
End Sub

Sub But5Trig_Hit
          ' 10 points & Advance
          AddScore 10, 0
          DoAdvance()
	But5Prim.TransY=-5
End Sub

Sub But5Trig_UnHit
	But5Prim.TransY=0
End Sub

'
'Sub Buttons_Hit(Idx)
''    ButtonsUp(Idx).TimerEnabled=True		' Enable animation timer
''    If ButtonLts(Idx).state=0 'IsDropped=True Then
''       ButtonsUp(Idx).IsDropped=True		' Drop UnLit button
''    Else
''       ButtonsUpLit(Idx).IsDropped=True		' Drop Lit button
''    End If
'    If Idx < 4 Then
'       PlaySound "Reels"
'    End If
'    Select Case Idx
'       Case 0
'          ButtonOff(1)						' Swap ON and OFF Button Lit
'          ButtonOn(3)
'          BumperOn(5)         				' On Yellow Bumpers
'          BumperOn(6)
'       Case 1
'          ButtonOff(2)						' Swap ON and OFF Button Lit
'          ButtonOn(4)
'          BumperOn(3)         				' On Red Bumpers
'          BumperOn(7)
'       Case 2
'          ButtonOn(1)						' Swap ON and OFF Button Lit
'          ButtonOff(3)
'          BumperOff(5)         				' Off Yellow Bumpers
'          BumperOff(6)
'       Case 3
'          ButtonOn(2)						' Swap ON and OFF Button Lit
'          ButtonOff(4)
'          BumperOff(3)         				' Off Red Bumpers
'          BumperOff(7)
'       Case 4
'          ' 10 points & Advance
'          AddScore 10, 0
'          DoAdvance()
'    End Select
'End Sub
''
'Sub ButtonsUp_Timer(Idx)
'    ButtonsUp(Idx).TimerEnabled=False		' Stop Animation Timer
'    If ButtonLts(Idx).state=0 Then
'       ButtonsUp(Idx).IsDropped=False		' Raise UnLit Button Case
'    Else
'       ButtonsUpLit(Idx).IsDropped=False	' Raise Lit Button Case
'    End If
'End Sub
'
' ---------------------------------------
'    Handle Button On/Off
' ---------------------------------------
'
Sub ButtonsOn()
    For X=1 to 5
        ButtonOn(X)
    Next
End Sub
'
Sub ButtonsOff()
    For X=1 to 5
       ButtonOff(X)
    Next
End Sub
'
Sub ButtonOn(Num)
    EVAL("But"&Num&"Lt").state=1  
End Sub
'
Sub ButtonOff(Num)
    EVAL("But"&Num&"Lt").state=0  'IsDropped=True
End Sub
'
' ---------------------------------------
'    Do "Advance of Voyage Lights
' ---------------------------------------
'
Sub DoAdvance()
    If CurLit >= 26 Then
       Exit Sub							' Already to Tahiti
    End If
    PlaySound "Reels"
    EVAL("Lt"&CurLit).state=0	' Turn Off Current light
    CurLit=CurLit+1
    EVAL("Lt"&CurLit).state=1	' Turn On Next light
    If (CurLit=11) OR (CurLit=21) OR (CurLit=26) Then
       KickerLit.state=1		' Lit Kicker for Special
    Else
	   KickerLit.state=0					' UnLight Kicker
    End If
    If (CurLit=6) OR (CurLit=16) OR (BallsPerGame=3) Then
       LtSouth.state=1			' Light South Special
       BumperOn(4)						' And Blue POP Bumper
    Else
       LtSouth.state=0			' UnLight South Special
       BumperOff(4)						' And Blue POP Bumper
    End If
    If CurLit=26 Then
       LtSpecial2.state=1		' Light Top Special at Tahiti
    End If
End Sub
'
' ---------------------------------------
' Do "randomization" of "Special" Lights
' ---------------------------------------
'
'  1. Alternate Top Advance Bumpers
'
'  2. If Outlane specials lit, alternate as opposite Advance Bumpers
'
Sub LightsRandomize(Init)
    If Init <> 0 Then
       Exit Sub
    End If
    If BTopLight1.State=LightStateOn Then
       BumperOff(1)
       BumperOn(2)
       If LtSpecial3.state=1 Then
          LtSpecial3.state=0
          LtSpecial1.state=1
       End If
    Else
       BumperOn(1)
       BumperOff(2)
       If LtSpecial1.state=1 Then
          LtSpecial1.state=0
          LtSpecial3.state=1
       End If
    End If
End Sub



'
'
' -------------------------------------------
'  Routines to Turn Bumper Lights On and Off
' -------------------------------------------
'
Sub BumpersOn()
    For X=1 to 7
       BumperOn(X)
    Next
End Sub
'
Sub BumpersOff()
    For X=1 to 7
       BumperOff(X)
    Next
End Sub
'
'  Turn Single Bumper Off
'
Sub BumperOff(Num)
'    EVAL("B"&Num&"SideOn").IsDropped=True
'    EVAL("B"&Num&"SideOff").IsDropped=False
'    EVAL("BBaseLight"&Num).State=LightStateOff
    EVAL("BTopLight"&Num).State=LightStateOff
'    Select Case Num
'       Case 3, 4, 5, 6, 7
'          EVAL("B"&Num&"Top1Lit").IsDropped=True	' Lower Lit Top
'          EVAL("B"&Num&"Top2Lit").IsDropped=True
'          EVAL("B"&Num&"Top3Lit").IsDropped=True
'       Case 1, 2
'          EVAL("B"&Num&"RimOn").IsDropped=True
'    End Select
End Sub
'
' Turn Single Bumper On
'
Sub BumperOn(Num)
'    EVAL("B"&Num&"SideOn").IsDropped=False
'    EVAL("B"&Num&"SideOff").IsDropped=True
'    EVAL("BBaseLight"&Num).State=LightStateOn
    EVAL("BTopLight"&Num).State=LightStateOn
'    Select Case Num
'       Case 3, 4, 5, 6, 7
'          EVAL("B"&Num&"Top1Lit").IsDropped=False	' Raise Lit Top
'          EVAL("B"&Num&"Top2Lit").IsDropped=False
'          EVAL("B"&Num&"Top3Lit").IsDropped=False
'       Case 1, 2
'          EVAL("B"&Num&"RimOn").IsDropped=False
'    End Select
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
       Exit Sub							' Too Soon for more
    End If
    PlaySound Snd
    GrowlTimer.UserValue=True
    GrowlTimer.Enabled=True
End Sub
'
Sub GrowlTimer_Timer()
    GrowlTimer.UserValue=False			' Allow Another AddSound
    GrowlTimer.Enabled=False			' Stop Timer
End Sub
'
'-------------------------------------------------
'  Display North, South, East, West on Back Glass
'-------------------------------------------------
'
Sub DisplayNESW()
    North.SetValue Spotted(1)			' Set North Light 0=off, 1=on
    East.SetValue Spotted(2)			' Set East Light  0=off, 1=on
    South.SetValue Spotted(3)			' Set South Light 0=off, 1=on
    West.SetValue Spotted(4)			' Set West Light  0=off, 1=on
End Sub
'
'--------------------------------------
'  Display Match Number on Back Glass
'--------------------------------------
'
'  Num = -1 to Turn Off Match Numbers
'  Num = 0-9 To display Num as Match Number
'
'  Num = any Other Number gives unpredictable result.
'
Sub DisplayMatch(Num)
    Select Case Num
       Case -1
          MatchReelA.SetValue 0				' All Match Numbers Off
          MatchReelB.SetValue 0
           If B2SOn then Controller.B2ssetmatch 34, 0
       Case 0, 1, 2, 3, 4, 5
          MatchReelA.SetValue Num+1			' 0=1 etc. Displays on Left Side
          MatchReelB.SetValue 0
          If B2SOn then Controller.B2ssetmatch 34, Match
       Case 6, 7, 8, 9
          MatchReelA.SetValue 0
          MatchReelB.SetValue Num-5			' 6=1, etc. Displays On Rt Side
          If B2SOn then Controller.B2ssetmatch 34, Match
    End Select
    MatchReelM.SetValue Num+1				' Mini Glass Displays All Nums
End Sub
'
'**********************************************************
'**** Subroutine to Display Current score on BackGlass ****
'**********************************************************
'
'  All scoring motor and reel advance timing is handled outside
'  of this routine.
'
'  On Entry: Score = Score to display (0=reset)
'
Sub DisplayScore()
    If Score=0 Then
       Score1.ResetToZero			' reset units, tens and hundreds
       Score10.ResetToZero
       Score100.ResetToZero
		If B2SOn then 
		   Controller.B2ssetscore 1, Score
		   Controller.B2ssetscore 2, Score
		   Controller.B2ssetscore 3, Score 
		End If
      ThScore1.SetValue 0			' set thousands to 0
    Else
       Score1.SetValue INT(Score MOD 10)	' Set units, tens and hundreds
    		If B2SOn then 
    		   Controller.B2ssetscore 3, INT(Score MOD 10)
    		   Controller.B2ssetscore 2, INT((Score MOD 100) /10)
    		   Controller.B2ssetscore 1, INT((Score MOD 1000) / 100)
    		   If Score > 1000 Then Controller.B2ssetdata 5, 1
    		End If
       Score10.Setvalue INT((Score MOD 100) /10)
       Score100.SetValue INT((Score MOD 1000) / 100)
       ThScore1.SetValue INT(Score/1000)	' Set Thousands Digit
    End If
End Sub
'
'--------------------------------------------------------
 
' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TradeWinds.width-1
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

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 6 ' total number of balls
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
            PlaySound("fx_ballrolling" & b), -1, (Vol(BOT(b) )*4), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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



'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and 
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


Sub TradeWinds_Exit()
	Disable(True)
	savehs
	If B2SOn Then Controller.stop
End Sub



Sub Buzzstart_Hit():PlaySound "Buzz":End Sub
Sub Buzzstop_Hit():StopSound "Buzz":End Sub


Sub Onbuzz_Hit():PlaySound "Onbuzz":End Sub
Sub Offbuzz_Hit():StopSound "Onbuzz":End Sub

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD: And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
'	if HSA1="" then HSA1=25
'	if HSA2="" then HSA2=25
'	if HSA3="" then HSA3=25
'	UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES


Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray  
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex	'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4

' ***********************************************************
'  HiScore DISPLAY 
' ***********************************************************

Sub UpdatePostIt
	dim tempscore
	HSScorex = HighScore
	TempScore = HSScorex
	HSScore1 = 0
	HSScore10 = 0
	HSScore100 = 0
	HSScoreK = 0
	HSScore10k = 0
	HSScore100k = 0
	HSScoreM = 0
	if len(TempScore) > 0 Then
		HSScore1 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreK = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreM = cint(right(Tempscore,1))
	end If
	Pscore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
	if HSScorex<1000 then
		PComma.image = HSArray(10)
	else
		PComma.image = HSArray(11)
	end if
	if HSScorex<1000000 then
		PComma2.image = HSArray(10)
	else
		PComma2.image = HSArray(11)
	end if
	HSName1.image = ImgFromCode(HSA1, 1)
	HSName2.image = ImgFromCode(HSA2, 2)
	HSName3.image = ImgFromCode(HSA3, 3)
End Sub

Function ImgFromCode(code, digit)
	Dim Image
	if (HighScoreFlashTimer.Enabled = True and hsLetterFlash = 1 and digit = hsCurrentLetter) then
		Image = "postitBL"
	elseif (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") then
		Image = "postit" & chr(code + ASC("A") - 1)
	elseif code = 27 Then
		Image = "PostitLT"
    elseif code = 0 Then
		image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
	end if
	ImgFromCode = Image
End Function

Sub HighScoreEntryInit()
	HSA1=0:HSA2=0:HSA3=0
	HSEnterMode = True
	hsCurrentDigit = 0
	hsCurrentLetter = 1:HSA1=1
	HighScoreFlashTimer.Interval = 250
	HighScoreFlashTimer.Enabled = True
	hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
	hsLetterFlash = hsLetterFlash-1
	UpdatePostIt
	If hsLetterFlash=0 then 'switch back
		hsLetterFlash = hsFlashDelay
	end if
End Sub


' ***********************************************************
'  HiScore ENTER INITIALS 
' ***********************************************************

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
				UpdatePostIt
			Case 2:
				HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
				UpdatePostIt
			Case 3:
				HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
				UpdatePostIt
		 End Select
    End If

	If keycode = RightFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1+1:If HSA1>26 Then HSA1=0
				UpdatePostIt
			Case 2:
				HSA2=HSA2+1:If HSA2>27 Then HSA2=0
				UpdatePostIt
			Case 3:
				HSA3=HSA3+1:If HSA3>27 Then HSA3=0
				UpdatePostIt
		 End Select
	End If
	
    If keycode = StartGameKey Then
		Select Case hsCurrentLetter
			Case 1:
				hsCurrentLetter=2 'ok to advance
				HSA2=HSA1 'start at same alphabet spot
'				EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
			Case 2:
				If HSA2=27 Then 'bksp
					HSA2=0
					hsCurrentLetter=1
				Else
					hsCurrentLetter=3 'enter it
					HSA3=HSA2 'start at same alphabet spot
				End If
			Case 3:
				If HSA3=27 Then 'bksp
					HSA3=0
					hsCurrentLetter=2
				Else
					savehs 'enter it
					HighScoreFlashTimer.Enabled = False
					HSEnterMode = False
				End If
		End Select
		UpdatePostIt
    End If
End Sub







'***MISC global. sounds
Sub rubbers_Hit(idx):PlaySound "fx_rubber3", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub metals_Hit(idx):PlaySound "fx_metalhit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub woods_Hit(idx):PlaySound "fx_target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "fx_rubber1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "fx_rubber3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub
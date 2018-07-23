' ****************************************************************
'                       VISUAL PINBALL X
' Author: Allknowing2012 2105/2016
'
' Credits to
'   JPSalas - refactored my code/table with his Pokemon table
'   VANLION on vpinball.com for his Gene Head 3D model
'   UltraPeepi for Ultradmd
'   Dark for the Speaker Model
'   Aldiode for the PF artwork
' ****************************************************************
' 20161025 - bonus animation, dmd background border, no credit audio, bsp vid
' 20161030 - apron mod (RustyCardores) and code version check added
'          - reset lockballs after gameover and store for each player
'          - code for super bumpers, super targets, combo shot on super ramps
' 20161031 - track shots made by player by song as some songs require more/less shots to complete
'          - save demon lock state
' 20161101 - dont stop/reset music on a ball save
' 20161102 - New Match Gifs - be sure to download them and put in ultradmd directory.
' 20161113 - New UDMD Location code from Seraph74, strengthen demon kicker, ballsaver on for start of multiball, random first song
' 20161119 - frontrow ball save, turn off super ramps between games, extend pause for DMB
' 20161120 - updated the ultradmd code from Seraph74
' 20161122 - pf/light update from Aldiode
' 20161122 - Bugs: Reset BumperLights, Front Row Save loops endlessly + KISSRules.vbs

' Thalamus 2018-07-23
' This tables doesn't have any standard "Positional Sound Playback Functions" or "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

If Table1.ShowDT then
  PaulStanleyDesktop.visible=True
  PaulStanleyFS.visible=False
Else
  PaulStanleyDesktop.visible=False
  PaulStanleyFS.visible=True
End if

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

On Error Resume Next
ExecuteGlobal GetTextFile("KISSCore.vbs")
If Err Then MsgBox "You need the KISSCore.vbs that accompanies this table"
if KissCoreV> "" then if KissCoreV < 1.01 or Err then msgbox "KISSCore.vbs needs updated version"
On Error Goto 0

On Error Resume Next
ExecuteGlobal GetTextFile("KISSMusic.vbs")
If Err Then MsgBox "You need the KISSMusic.vbs that accompanies this table"
if KissMusicV> "" then if KissMusicV < 1.01 or Err then msgbox "KISSMusic.vbs needs updated version"
On Error Goto 0

On Error Resume Next
ExecuteGlobal GetTextFile("KISSRules.vbs")
If Err Then MsgBox "You need the KISSRules.vbs that accompanies this table"
if KissRulesV> "" then if KissRulesV < 1.05 or Err then msgbox "KISSRules.vbs needs updated version"
On Error Goto 0

On Error Resume Next
ExecuteGlobal GetTextFile("KISSCode.vbs")
If Err Then MsgBox "You need the KISSCode.vbs that accompanies this table"
if KissCodeV> "" then if KissCodeV < 1.02 or Err then msgbox "KISSCode.vbs needs updated version"

On Error Resume Next
ExecuteGlobal GetTextFile("KISSDMDV2.vbs")
If Err Then MsgBox "You need the KISSDMDV2.vbs that accompanies this table"
if KissDMDV > "" then if KissDMDV < 1.08 or Err then msgbox "KISSDMDV2.vbs needs updated version"
On Error Goto 0

Const BallSize = 50 ' 50 is the normal size
Const UseUDMD=True  ' FALSE is only for testing .. not fully functional as FALSE
Const bgi = "Black.bmp"  ' or just "" if you want the border around the dmd


' *************************************
' Modify UltraDMD setttings here
' *************************************

UseFullColor = "True"              '    "True" / "False"
DMDColorSelect = "OrangeRed"        '     Rightclick on UDMD window to get full list of colours

DMDPosition = False                 '     Use Manual DMD Position, True / False
DMDPosX = 2303                      '     Position in Decimal,  Set a value here if you want to have a table specific value eg. 2303
DMDPosY = 660                       '     Position in Decimal

' Rightclick on UDMD window to get full list of colours
GetDMDColor


' Load the core.vbs for supporting Subs and functions
On Error Resume Next
  ExecuteGlobal GetTextFile("core.vbs")
  If Err Then MsgBox "Can't open core.vbs"
On Error Goto 0

' Define any Constants
Const cGameName  = "kiss_original_2016"
Const myVersion = "1.0.0"
Const MaxPlayers = 4
Const BallSaverTime = 15   '15 'in seconds
Const MaxMultiplier = 99 'no limit in this game
Const BallsPerGame = 3   ' 3 or 5

'load up the backglass and DOF
LoadEM


' Define Global Variables
Dim PlayersPlayingGame
Dim CurPlayer
Dim Credits
Dim Bonus
Dim BonusPoints(4)
Dim BonusMultiplier(4)
Dim BallsRemaining(4),LockedBalls(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)

Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim bAutoPlunger

' Define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock
Dim BallsInHole

' Define Game Flags
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim bMusicOn
Dim bExtraBallWonThisBall
Dim PI: PI = Round(4 * Atn(1), 6)

Dim plungerIM 'used mostly as an autofire plunger
Dim ttable

Dim LoveGunMode, DemonMBMode,InDemonLock,InScoop,KISSTargetStack,ArmyTargetStack,FrontRowStack
Dim am, ModeScore, ModeInProgress, MBScore, ChooseSongMode
Dim Spincnt(4)    ' how many times the spinner has went around
Dim DemonMB(4)
Dim Shots(4,15)   ' Players/Songs
Dim CreditAwarded(4)
Dim BumperCnt(4),BumperColor(4), ComboCnt(4)
Dim RampCnt(4)
Dim LastShot(4)
Dim TargetCnt(4)
Dim Instruments(4)
Dim CurSong(4), CurCity(4)
Dim SaveSong   ' keep track of track number while in MB
Dim CurScene
Dim HighCombo, HighComboName, LastScoreP1, LastScoreP2, LastScoreP3, LastScoreP4
Dim KISSBonus,ArmyBonus, HSMode, dd, FrontRowSave




' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    Dim i
    dd=10 ' just for debugging
    Randomize
    ' initalise the DMD display
    If UseUDMD then LoadUltraDMD

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFXDOF("fx_kicker",141,DOFPulse,DOFContactors), SoundFXDOF("fx_solenoid",141,DOFPulse,DOFContactors)
        .CreateEvents "plungerIM"
    End With

    Set ttable = New cvpmTurnTable
    With ttable
        .InitTurnTable Magnet1, 90
        .spinCW = False
        .MotorOn = True
        .CreateEvents "ttable"
    End With

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    'load saved values, highscore, names
    Loadhs

    'Init main variables
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        LockedBalls(i)=0
        ExtraBallsAwards(i) = 0
        CurCity(i)=1
        spincnt(i)=0    ' how many times the spinner has went around
        DemonMB(i)=0
        for xx=1 to 15
          Shots(i,xx)=0      ' How many shots achieved
        next
        BumperCnt(i)=0:RampCnt(i)=0:TargetCnt(i)=0
        BumperColor(i)=0   ' 1 per bumper not per player
        CurSong(i)=1
        ComboCnt(i)=0
        CreditAwarded(i)=False
    Next
    ResetBumpers()

    if UseUDMD Then  ' wait for the intro video to finish
       do while UltraDMD.isRendering
       loop
       UltraDMD.clear
    End if


    CurScene=""

    ' freeplay or coins
    bFreePlay = False 'we want coins

    ' initialse any other flags
    ChooseSongMode = False
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    FrontRowSave = False
    bGameInPlay = False
	bAutoPlunger = False
    bMusicOn = True
    BallsOnPlayfield = 0
    BallsInLock = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    EndOfGame()
End Sub

Function LPad(s, l, c)
  Dim n : n = 0
'  debug.print "LPAD S="  & s & " l=" & l & " c=" & c
  If l > Len(s) Then n = l - Len(s)
  LPad = String(n, c) & s
End Function

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
    If Keycode = AddCreditKey Then
        Credits = Credits + 1
        DOF 125, DOFOn
        PlaySound "fx_coin",0,1,0.25,0.2
        If NOT bGameInPlay then
          If(Tilted = False) Then
            AttractMode.enabled=False
			'OnScoreboardChanged()
            If UseUDMD then
               UltraDMD.CancelRendering:UltraDMD.Clear
               UltraDMD.DisplayScene00 "scene02.gif", "PRESS START", 15, "CREDITS " & credits, -1, UltraDMD_Animation_None, 3500, UltraDMD_Animation_None
            End if
            AttractMode.interval=1500:AttractMode.enabled=True
          End If
        Else
			OnScoreboardChanged()
        End If
    End If

    If keycode = PlungerKey Then
'        PlungerIM.AutoFire
        Plunger.PullBack
		PlaySound "fx_plungerpull",0,1,0.25,0.2
    End If

    If bGameInPlay AND NOT Tilted Then

    If keycode = LeftTiltKey Then   Nudge 90,  6:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then  Nudge 270, 6:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0,   7:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

dim xxx,yyy
if keycode = LeftMagnaSave and 1=2 Then
   debug.print "Enable debugging stuff"
  'i39.state=LightStateOn   ' ExtraBalls
  'i41.state=LightStateOn ' BackStagePass
  ' i44.state=LightStateOn ' New Track
   'i43.state=LightStateOn  ' Rock City
   'i40.state=LightStateOn  ' Kiss ArmyBonus
   ' i101.state=LightStateOn  ' starchild bumper
   ' i102.state=LightStateOn  ' starchild bumper
   ' i103.state=LightStateOn  ' starchild bumper
   ' i104.state=LightStateOn  ' starchild bumper
  ' i61.state=LightStateOn ' Super Bumper
   ' i95.state=LightStateOn ' Super Targets
    'i118.state=LightStateOn ' Super Ramps
   ' i122.state=LightStateOn ' Super Targets
   dd=dd+1:if dd>74 then dd=10
   xxx="scene" & lpad(dd,2,"0") & ".gif"
   'debug.print dd & " " & xxx
   'DMDGif xxx,dd,"",slen(dd)
   'DMDGif "","DONE","DONE",1000
   'DisplayI(dd)
    'RScene(dd)
    'spinv("5K")
    'SpinV "Spinner Count" &  SpinCnt(CurPlayer)
  '  BV(1)
    'msgbox "HighCombo=" & HighCombo
    'ComboCnt(1)=HighCombo+1
    'DMDGif "frame508.jpg", "", "9" + "0             ", 1000
    'FlashForMs SmallFlasher2, 500, 50, 0
'    UltraDMD.DisplayScene00  BonusLights1, "", 10, "", 10, UltraDMD_Animation_None, 1100, UltraDMD_Animation_None
'    UltraDMD.DisplayScene00  "black.bmp", "", 10, "", 10, UltraDMD_Animation_None, 1000, UltraDMD_Animation_None
  ' i17.state = LightStateOn
  ' i21.state = LightStateOn
  ' i25.state = LightStateOn
  'i9.state = LightStateOn:i13.state = LightStateOn:I106.state=LightStateOn:i107.state=LightStateOn:I108.state=LightStateOn:I109.state=LightStateOn:CheckBackStage()
  ' for yyy=1 to 4
   '  for xxx=1 to 15
    '   debug.print yyy & " " & xxx & "=" & Shots(yyy,xxx)
    ' Next
   'next
end if
if keycode=RightMagnaSave and 1=2 Then
  i6.state=LightStateBlinking
end if

         if (keycode = RightMagnaSave or keycode=LeftMagnaSave) and ChooseSongMode then
           NewTrackTimer.enabled=False
  	       ChooseSongMode=False
           ScoopDelay.interval=200
           ScoopDelay.Enabled = True
           vpmtimer.addtimer 500, "FlashForMs FlasherExitHole, 1500, 30, 0 '"
        end if

        If keycode = LeftFlipperKey Then
           SolLFlipper 1
  		   'Leftflipper.RotateToStart
           Leftflipper.TimerEnabled = 1  ' nFozzy Flipper Code
           Leftflipper.TimerInterval = 16
           Leftflipper.return = returnspeed * 0.5
        End If
        If keycode = RightFlipperKey Then
           SolRFlipper 1
  		   'Rightflipper.RotateToStart
           Rightflipper.TimerEnabled = 1  ' nFozzy Flipper Code
           Rightflipper.TimerInterval = 16
           Rightflipper.return = returnspeed * 0.5
        End if

        If keycode = StartGameKey and NOT hsbModeActive Then
            If((PlayersPlayingGame < MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                Else
                    If(Credits > 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
                        If Credits < 1 Then DOF 125, DOFOff
                    Else
                        ' Not Enough Credits to start a game.
                        DMDFlush:debug.print "Flush 2"
                        PlaySound "audio749"   ' You got to pay to play
                        DMDTextPauseI "INSERT COIN", "CREDITS " & Credits, 1000, "scene02.gif"
                        PlaySound "audio749"
                    End If
                End If
                If UseUDMD Then UltraDMD.CancelRendering:UltraDMD.Clear
                OnScoreboardChanged()
            End if
        End If
        If keycode = LeftFlipperKey and bBallInPlungerLane and BallsOnPlayfield = 1 Then
            debug.print "Change City"
            DMDFlush():debug.print "Flush 3"
            NextCity
        End If
        If keycode = RightFlipperKey and ((bBallInPlungerLane and BallsOnPlayfield = 1) or ChooseSongMode) Then
            debug.print "Change Song"
            DMDFlush():debug.print "Flush 4"
            if AutoPlungeTimer.enabled Then
               AutoPlungeTimer.enabled=False:AutoPlungeTimer.enabled=True
            End If
            NextSong()
        End If
        Else ' If (GameInPlay)
            If keycode = StartGameKey and NOT hsbModeActive Then
                If(bFreePlay = True) Then
                    If(BallsOnPlayfield = 0) Then
                        ResetForNewGame()
                    End If
                Else
                    If(Credits > 0) Then
                        If(BallsOnPlayfield = 0) Then
                            Credits = Credits - 1
                            if Credits < 1 Then DOF 125, DOFOff
                            ResetForNewGame()
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        DMDFlush:debug.print "Flush 5"
                        DMDTextPauseI "INSERT COIN", "CREDITS " & Credits, 1000, "scene02.gif"
                        PlaySound "audio749"
                    End If
                End If
            End If

    End If ' If (GameInPlay)

    If hsbModeActive Then UDMDTimer.Enabled=False:EnterHighScoreKey(keycode)

' Table specific
End Sub

' nFozzy Flipper Code
dim returnspeed, lfstep, rfstep
returnspeed = leftflipper.return
lfstep = 1
rfstep = 1

sub leftflipper_timer()
	select case lfstep
		Case 1: Leftflipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
		Case 2: Leftflipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
		Case 3: Leftflipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
		Case 4: Leftflipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
		Case 5: Leftflipper.return = returnspeed * 1 :lfstep = lfstep + 1
		Case 6: Leftflipper.timerenabled = 0 : lfstep = 1
	end select
end sub

sub rightflipper_timer()
	select case rfstep
		Case 1: Rightflipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
		Case 2: Rightflipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
		Case 3: Rightflipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
		Case 4: Rightflipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
		Case 5: Rightflipper.return = returnspeed * 1 :rfstep = rfstep + 1
		Case 6: Rightflipper.timerenabled = 0 : rfstep = 1
	end select
end sub

Sub SongPause_Timer
  SongPause.enabled=False
  debug.print "Playing Song" & Song
  PlayMusic Song
End Sub

Sub Table1_MusicDone
  PlayMusic Song  ' continue on with the Song
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then SolLFlipper 0
        If keycode = RightFlipperKey Then SolRFlipper 0
    End If
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "fx_plunger",0,1,0.25,0.25
	End If
End Sub

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                      'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                  'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt > TiltSensitivity) AND(Tilt < 15) Then 'show a warning
        DisplayI(2)
    End if
    If Tilt > 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        DMDFlush:debug.print "Flush 6"
        DisplayI(1)
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt > 0 Then
        Tilt = Tilt - 0.1
    Else
        Me.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        'turn off GI and turn off all the lights
        GiOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        Bumper1.Force = 0
        Bumper2.Force = 0
        Bumper3.Force = 0
        Bumper4.Force = 0

        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
        LeftSlingshot1.Disabled = 1
        RightSlingshot1.Disabled = 1
    Else
        'turn back on GI and the lights
        'GiOn
        Bumper1.Force = 6
        Bumper2.Force = 6
        Bumper3.Force = 6
        Bumper4.Force = 6
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        LeftSlingshot1.Disabled = 0
        RightSlingshot1.Disabled = 0
        'clean up the buffer display
        DMDFlush:debug.print "Flush 7"
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0) Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        EndOfBall()
        Me.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'********************
' Music as wav sounds
'********************

Dim Song
Song = ""

Sub PlaySoundEffect()
End Sub

Sub PlayFanfare()
End Sub

' Ramp Sounds
Sub RHelp1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub RHelp2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub


' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
    Dim i

    bGameInPLay = True

    'resets the score display, and turn off attrack mode
    StopAttractMode
    GiOn

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
        CurCity(i)=1
        spincnt(i)=0    ' how many times the spinner has went around
        DemonMB(i)=0
        For xx=1 to 15
          Shots(i,xx)=0      ' How many shots achieved
        Next
        BumperCnt(i)=0:RampCnt(i)=0:TargetCnt(i)=0
        BumperColor(i)=0   ' 1 per bumper not per player
        CurSong(i)=Int(RND*7)+1
        ComboCnt(i)=0
        CreditAwarded(i)=False
        LockedBalls(i)=0
    Next
    If UseUDMD Then UltraDMD.SetScoreboardBackgroundImage CurSong(CurPlayer) & ".png",15,13:UltraDMD.Clear:OnScoreboardChanged()


    ' initialise any other flags
    bMultiBallMode = False
    Tilt = 0

    ' initialise Game variables
    Game_Init()
    OnScoreboardChanged()

    ' you may wish to start some music, play a sound, do whatever at this point
    PlaySound "audio451",0, 0.5  ' Are You Ready?

    ' set up the start delay to handle any Start of Game Attract Sequence
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
    Dim i

    DMDTextPauseI  "PLAYER " & CurPlayer, "", 1000, bgi
    UDMDTimer.interval=500:UDMDTimer.Enabled=True

    ' make sure the correct display is upto date
   ' AddScore 0

    ' reset any drop targets, lights, game modes etc..
    Bonus = 0
    bExtraBallWonThisBall = False
    ResetNewBallLights()

    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True

    ModeScore=0
  ' Raise the Targets
  ' Stop the timers
    ArmyCombo.enabled=False
    KissCombo.enabled=False
    KissHurryUp.enabled=False
    ArmyHurryUp.enabled=False
    LoveGunMode=False:DemonMBMode=False
    InDemonLock=False
    InScoop=False
    KissTargetStack=0
    ArmyTargetStack=0
    FrontRowStack=0
    MBScore=0   ' total score during DemonMB
    ResetTargets()
    sctarget.isdropped=False
    BallsInLock=0

    For i = 1 to MaxPlayers
       instruments(i)=0
       LastShot(i)=-1  ' Track what was hit last
       ComboCnt(i)=0
    next

    ResetBumpers()
    ModeInProgress=False
End Sub


' Create a new ball on the Playfield


Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedball BallSize / 2
    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    PlaySound SoundFXDOF("fx_Ballrel",123,DOFPulse,DOFContactors), 0, 1, 0.1, 0.1
    BallRelease.Kick 90, 4

    ' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
	' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield > 1 Then
        bMultiBallMode = True
		bAutoPlunger = True
    Else
      ' If ballsaver or front row then dont reset song and music
      If MusicFlag=False then
        EndMusic
        Song = track(CurSong(CurPlayer))
        InitMode(CurSong(CurPlayer))
        SongPause.interval=100
        SongPause.enabled=True
      End If
      MusicFlag=False
      AutoPlungeTimer.interval=10000 ' 60s then autoplunge
      AutoPlungeTimer.enabled=True
    End If
End Sub

Sub AutoPlungeTimer_Timer
  debug.print "About to autoplunge"
  If bBallInPlungerLane Then
    PlungerIM.AutoFire
    debug.print "PLUNGE!"
  Else
    debug.print "Already plunged"
  End If
End Sub

' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
    ' wait if there is a ball in the plunger lane
    If bBallInPlungerLane Then
        Exit Sub
    Else
        CreateNewBall()
        mBalls2Eject = mBalls2Eject -1
        If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
            Me.Enabled = False
        End If
    End If
End Sub


'   if BonusCnt=1 then PlaySound "audio297",0,0.3,0.25,0.25
'   if BonusCnt=2 then PlaySound "audio298",0,0.3,0.25,0.25
'   if BonusCnt=3 then PlaySound "audio299",0,0.3,0.25,0.25
'   if BonusCnt=4 then PlaySound "audio300",0,0.3,0.25,0.25

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded
'
Dim SoundLoops

Sub EndOfBall()
    Dim BonusDelayTime
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False
    KissHurryUp.enabled=False
    ArmyHurryUp.enabled=False

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)
    SaveShots()
    SaveState()
    If(Tilted = False) Then
        Dim AwardPoints, TotalBonus
        AwardPoints = 0:TotalBonus = 0

' add in any bonus points (multipled by the bonus multiplier)
      if MBScore <> 0 then
        DMDTextPauseI "Demon Total", MBScore, 800, bgi
        BonusDelayTime = BonusDelayTime+1000
      end If
      if ModeScore <> 0 then
        DMDTextPauseI "Song Total", ModeScore, 800, bgi
        BonusDelayTime = BonusDelayTime+1000
      end if

      SoundLoops=0
      AwardPoints = i9.state * 50000 + i10.state * 50000 + i11.state * 50000 + i12.state * 50000
      if AwardPoints <> 0 then
        TotalBonus = TotalBonus + AwardPoints
        SoundLoops=SoundLoops+1
        'DMDText "KISS Bonus", AwardPoints
        UltraDMD.DisplayScene00  BonusLights1, "", 10, "", 10, UltraDMD_Animation_None, 1100, UltraDMD_Animation_None
      '  UltraDMD.DisplayScene00  "black.bmp", "", 10, "", 10, UltraDMD_Animation_None, 0, UltraDMD_Animation_None
        BonusDelayTime = BonusDelayTime+1800
      end if

      AwardPoints = i13.state * 50000 + i14.state * 50000 + i15.state * 50000 + i16.state * 50000
      debug.print "army" & AwardPoints
      if AwardPoints <> 0 then
        TotalBonus = TotalBonus + AwardPoints
        'DMDText "Army Bonus", AwardPoints
        SoundLoops=SoundLoops+1
        UltraDMD.DisplayScene00  BonusLights2, "", 10, "", 10, UltraDMD_Animation_None, 1100, UltraDMD_Animation_None
        'UltraDMD.DisplayScene00  "black.bmp", "", 10, "", 10, UltraDMD_Animation_None, 0, UltraDMD_Animation_None
        BonusDelayTime = BonusDelayTime+1800
      end if

      AwardPoints = i17.state * 150000 + i18.state * 150000 + i19.state * 50000 + i20.state * 150000
      if AwardPoints <> 0 then
        TotalBonus = TotalBonus + AwardPoints
'        DMDText "Band Bonus", AwardPoints
        SoundLoops=SoundLoops+1
        UltraDMD.DisplayScene00  BonusLights3, "", 10, "", 10, UltraDMD_Animation_None, 1100, UltraDMD_Animation_None
       ' UltraDMD.DisplayScene00  "black.bmp", "", 10, "", 10, UltraDMD_Animation_None, 0, UltraDMD_Animation_None
        BonusDelayTime = BonusDelayTime+1800
      end if

      AwardPoints = Instruments(CurPlayer) * 100000
      if AwardPoints <> 0 then
        TotalBonus = TotalBonus + AwardPoints
'        DMDText "Instrument Bonus", AwardPoints
        SoundLoops=SoundLoops+1
        UltraDMD.DisplayScene00  BonusLights4, "", 10, "", 10, UltraDMD_Animation_None, 1100, UltraDMD_Animation_None
      '  UltraDMD.DisplayScene00  "black.bmp", "", 10, "", 10, UltraDMD_Animation_None, 0, UltraDMD_Animation_None
        BonusDelayTime = BonusDelayTime+1800
      end if

      if BonusMultiplier(CurPlayer) <> 1 then
        DMDTextI "Total Bonus X " & BonusMultiplier(CurPlayer), TotalBonus, bgi
      Else
        DMDTextPauseI "Total Bonus", TotalBonus, 3000, bgi
      end if
      TotalBonus = TotalBonus * BonusMultiplier(CurPlayer)

      if SoundLoops > 0 Then
        EOBSoundTimer.interval=1200:EOBSoundTimer.enabled=True
      End if

      debug.print "Sent the bonus stuff to the dmd queue"
        ' add a bit of a delay to allow for the bonus points to be shown & added up
      BonusDelayTime = BonusDelayTime+5500
      vpmtimer.addtimer BonusDelayTime, "Addscore TotalBonus '"
    Else
      'no bonus to count so move quickly to the next stage
      BonusDelayTime = 100
    End If
    ' start the end of ball timer which allows you to add a delay at this point
    EndMusic

    vpmtimer.addtimer BonusDelayTime, "EndOfBall2 '"
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the CurPlayer)
'
Sub EOBSoundTimer_Timer
  EOBSoundTimer.enabled=False
  PlaySound "audio297Loop",SoundLoops+1,0.3,0.25,0.25
End Sub

Sub EndOfBall2()
    StopSound "audio297Loop"
    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilted = False
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

    ' has the player won an extra-ball ? (might be multiple outstanding)
debug.print "Checking if We won extra ball"
    If(ExtraBallsAwards(CurPlayer) <> 0) Then
        debug.print "Extra Ball"
        DisplayI(15)

        ' yep got to give it to them
        ExtraBallsAwards(CurPlayer) = ExtraBallsAwards(CurPlayer) - 1
        ' Turn on BallSaver again
        bBallSaverReady = True

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurPlayer) = 0) Then
            LightRockAgain.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
         Select case int(RND*2)+1
            case 1: PlaySound "audio570"
            case 2: PlaySound "audio573"
         End Select

        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls
debug.print "No extra ball"
        BallsRemaining(CurPlayer) = BallsRemaining(CurPlayer) - 1

        ' was that the last ball ?
        If(BallsRemaining(CurPlayer) <= 0) Then
            debug.print "No More Balls, High Score Entry"
            UDMDTimer.Enabled=False
            ' Submit the CurPlayers score to the High Score system
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
'
Sub EndOfBallComplete()
    Dim NextPlayer

    debug.print "EndOfBall - Complete"

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame > 1) Then
        ' then move to the next player
        NextPlayer = CurPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer > PlayersPlayingGame) Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurPlayer
    End If

    debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode
        UDMDTimer.Enabled=False
        if UseUDMD then UltraDMD.CancelRendering
        CheckMatch()


    ' you may wish to put a Game Over message on the desktop/backglass

       LastScoreP1=Score(1)
       LastScoreP2=Score(2)
       LastScoreP3=Score(3)
       LastScoreP4=Score(4)

    Else
        ' set the next player
        CurPlayer = NextPlayer

        ' make sure the correct display is up to date
        AddScore 0

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()
        OnScoreboardChanged()
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    debug.print "End Of Game"
    bGameInPLay = False

    ' just ended your game then play the end of game tune

    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all modes - eject locked balls
    ' set any lights for the attract mode
    GiOff
    StartAttractMode 1
' you may wish to light any Game Over Light you may have
End Sub

Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' lost a ball ;-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one
'
Dim MusicFlag   ' Set this to true if you dont want the music to stop
MusicFlag=False
Sub Drain_Hit()
    ' Destroy the ball
    Drain.DestroyBall
    BallsOnPlayfield = BallsOnPlayfield - 1
    ' pretend to knock the ball into the ball storage mech
    PlaySound "fx_drain"

    debug.print "Show Saved Song (c) #" & SaveSong

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True) AND(Tilted = False) Then

        ' is the ball saver active,
        MusicFlag=False
        If(bBallSaverActive = True or FrontRowSave=True) Then
            debug.print "bBallSaverActive = True"
            if FrontRowSave=True then
              FrontRowSave=False
              i6.state=LightStateOff
            end if
            ' yep, create a new ball in the shooters lane
            ' we use the Addmultiball in case the multiballs are being ejected
            AddMultiball 1
            MusicFlag=True
            Select case int(RND*2)+1
               case 1: PlaySound "audio570"   'Lets do that again
               case 2: PlaySound "audio573"
            End Select
			' we kick the ball with the autoplunger
			bAutoPlunger = True
            ' you may wish to put something on a display or play a sound at this point
            if UseUDMD then UltraDMD.CancelRendering
            DisplayI(26)
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1) Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True) then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ChangeGi "white"
                    ' you may wish to change any music over at this point and
                    ' turn off any multiball specific lights
    debug.print "Show Saved Song (d) #" & SaveSong
                    ResetJackpotLights
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0) Then
                ' handle the end of ball (change player, high score entry etc..)
                if UseUDMD then UltraDMD.CancelRendering
                ' End Modes and timers
                EndOfBall()
                ChangeGi "white"
            End If
      End if
    End If
    FrontRowSave=False
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    debug.print "ball in plunger lane"
    debug.print "Show Saved Song (b) #" & SaveSong
    ' some sound according to the ball position
    PlaySound "fx_sensor", 0, 1, 0.15, 0.25
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
    LaunchLight.State = 2
    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        debug.print "autofire the ball"
        PlungerIM.AutoFire
		DOF 124, DOFPulse
		DOF 121, DOFPulse
		bAutoPlunger = False
    End If
    ' resync target lights in case we may have chose Hotter Than Hell
    debug.print "swPlungerRest - resetting lights"
    if cursong(CurPlayer)=4 then ' Hotter Than Hell means they need to flash on Reset
      if sw38.isdropped=1 then i35.state=LightStateOn else i35.state=2 end if
      if sw39.isdropped=1 then i36.state=LightStateOn else i36.state=2 end if
      if sw40.isdropped=1 then i37.state=LightStateOn else i37.state=2 end if
      if sw41.isdropped=1 then i38.state=LightStateOn else i38.state=2 end if
    else
      if sw38.isdropped=1 then i35.state=LightStateOn else i35.state=LightStateOff end if
      if sw39.isdropped=1 then i36.state=LightStateOn else i36.state=LightStateOff end if
      if sw40.isdropped=1 then i37.state=LightStateOn else i37.state=LightStateOff end if
      if sw41.isdropped=1 then i38.state=LightStateOn else i38.state=LightStateOff end if
    end if

    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
      ' Just start the flash and not the timer itself
      LightRockAgain.BlinkInterval = 160
      LightRockAgain.State = 2
    End If
    BallLooping.enabled=False ' dont trigger switches on the plunge
End Sub

' The ball is released from the plunger turn off some flags

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
	DOF 141, DOFPulse
	DOF 121, DOFPulse
    ' turn off LaunchLight
    LaunchLight.State = 0
    AutoPlungeTimer.enabled=False ' No longer autoplunge after 30s
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
    End If
    BallLooping.Interval=1400:BallLooping.enabled=True ' dont trigger switches on the loop
End Sub

Sub EnableBallSaver(seconds)
    debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
     ' start the timer
    BallSaverTimer.Interval = 1000 * seconds
    BallSaverTimer.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightRockAgain.BlinkInterval = 160
    LightRockAgain.State = 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimer_Timer()
    debug.print "Ballsaver ended"
    BallSaverTimer.Enabled = False
    ' clear the flag
    bBallSaverActive = False
    ' if you have a ball saver light then turn it off at this point or turn on if you have ExtraBall earned
    If(ExtraBallsAwards(CurPlayer) = 0) Then
       LightRockAgain.State = 0
       debug.print "dont relight the extra ball"
    End if
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    LightRockAgain.BlinkInterval = 80
    LightRockAgain.State = 2
End Sub

' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board
'
Sub AddScore(points)
    If(Tilted = False) Then
        ' add the points to the current players score variable
        Score(CurPlayer) = Score(CurPlayer) + points
        If ModeInProgress then
          ModeScore=ModeScore+points
        End if
        If DemonMBMode Then
          MBScore=MBScore+points
        End if
        ' update the score displays
        OnScoreboardChanged()
    End if

' you may wish to check to see if the player has gotten a replay
    if Score(CurPlayer) > 15000000 and creditawarded(CurPlayer) = False then
        credits = credits + 1
        creditawarded(CurPlayer)=True
        PlaySound SoundFXDOF("fx_Knocker",122,DOFPulse,DOFKnocker)
        DisplayI(16)
        OnScoreboardChanged
    end if
End Sub


'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************

' tables walls and animations
Sub VPObjects_Init
  sw38.isDropped = False ' KISS Drop Targets
  sw39.isDropped = False
  sw40.isDropped = False
  sw41.isDropped = False
End Sub

' tables variables and modes init

Sub Game_Init()
    bExtraBallWonThisBall = False
    TurnOffPlayfieldLights()
    'Play some Music
    'Init Variables
    BallInHole = 0
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub ResetNewBallLights()
    'LightArrow1.State = 2
    'LightArrow6.State = 2
    I106.state = 0 ' ARMY
    I107.state = 0
    I108.state = 0
    I109.state = 0
    i64.state=0:i65.state=0:i66.state=0:i67.state=0
    SetLightColor i115,"white", 0  'HurryUp
    SetLightColor i77, "white",0   'HurryUp

    SetLightColor i82, "orange", 2  ' Demon Lock Light
    SetLightColor i84, "orange", 2  ' Demon Lock Light
    SetLightColor i7, "white", 2   ' Light Bumper
    SetLightColor i8, "white", 2   ' Combo
    SetLightColor i33, "orange", 2   ' Combo
    SetLightColor i97, "orange", 2   ' Instrument

    'Grid
    SetLightColor i9,  "white",0
    SetLightColor i10, "white",0
    SetLightColor i11, "white",0
    SetLightColor i12, "white",0
    SetLightColor i13, "yellow",0
    SetLightColor i14, "yellow",0
    SetLightColor i15, "yellow",0
    SetLightColor i16, "yellow",0
    SetLightColor i17, "white",0
    SetLightColor i18, "white",0
    SetLightColor i19, "white",0
    SetLightColor i20, "white",0
    SetLightColor i21, "red",0
    SetLightColor i22, "red",0
    SetLightColor i23, "red",0
    SetLightColor i24, "red",0

    SetLightColor i25, "white",0
    SetLightColor i26, "white",0
    SetLightColor i27, "white",0
    SetLightColor i29, "red",0
    SetLightColor i30, "red",0
    SetLightColor i31, "white",0
End Sub


' *********************************************************************
'                        Table Object Hit Events

Sub sw47_hit ' star child Kicker1_Hits
   PlaySound "fx_hole-enter", 0, 1, -0.1
   If f116.state=1 then
     FlashForMs Flasher9, 1000, 50, 0
     FlashForMs Flasher10, 1000, 50, 0
     Start_LoveGun()
     f116.state=0    ' turn off StarChild Flasher
     'i100.state=0   ' right orbit  - might not want to turn this off
   'else
     DMDGif "scene51.gif","STARCHILD","",slen(51)
   End if

   Me.TimerInterval = 1500
   Me.TimerEnabled = 1
End Sub

Sub sw47_unhit
    PlaySound SoundFXDOF("fx_popper",120,DOFPulse,DOFContactors), 0, 1, -0.1, 0.25
    DOF 121, DOFPulse
End Sub


sub sw47_timer
   debug.print "Exit star hole"
   Me.TimerEnabled=False
   sw47.kick 225,6,1
end sub


Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_slingshot",103,DOFPulse,DOFcontactors), 0, 1, -0.05, 0.05
	DOF 105, DOFPulse
    LeftSling4.Visible = 1:LeftSling1.Visible = 0
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    RandomScene()
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1
        Case 3:LeftSLing2.Visible = 0:LeftSling1.Visible = 1:Gi2.State = 1:LeftSlingShot.TimerEnabled = False
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_slingshot",104,DOFPulse,DOFcontactors), 0, 1, 0.05, 0.05
	DOF 106, DOFPulse
    RightSling4.Visible = 1:RightSling1.Visible = 0
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 110
    RandomScene()
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1
        Case 3:RightSLing2.Visible = 0:RightSLing1.Visible = 1:Gi1.State = 1:RightSlingShot.TimerEnabled = False
    End Select
    RStep = RStep + 1
End Sub

Dim LStep1, RStep1

Sub LeftSlingShot1_Slingshot
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_slingshot",143,DOFPulse,DOFContactors), 0, 1, -0.05, 0.05
    Sling3a.Visible = 1:Sling3.Visible = 0
    Lemk1.RotX = 26
    LStep1 = 0
    LeftSlingShot1.TimerEnabled = True
    ' add some points
    AddScore 510
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot1"
' add some effect to the table?
End Sub

Sub LeftSlingShot1_Timer
    Select Case LStep1
        Case 1:Sling3a.Visible = 0:Sling3b.Visible = 1:Lemk1.RotX = 14
        Case 2:Sling3b.Visible = 0:Sling3c.Visible = 1:Lemk.RotX = 0
        Case 3:Sling3c.Visible = 0:Sling3.Visible = 1:Lemk1.RotX = -20:LeftSlingShot1.TimerEnabled = False
    End Select
    LStep1 = LStep1 + 1
End Sub

Sub RightSlingShot1_Slingshot
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_slingshot",144,DOFPulse,DOFContactors), 0, 1, 1, 0.05
    Sling2a.Visible = 1:Sling2.Visible = 0
    Remk1.RotX = 26
    RStep1 = 0
    RightSlingShot1.TimerEnabled = True
    ' add some points
    AddScore 510
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot1"
' add some effect to the table?
End Sub

Sub RightSlingShot1_Timer
    Select Case RStep1
        Case 1:Sling2a.Visible = 0:Sling2b.Visible = 1:Remk.RotX = 14
        Case 2:Sling2b.Visible = 0:Sling2c.Visible = 1:Remk.RotX = 0
        Case 3:Sling2c.Visible = 0:Sling2.Visible = 1:Remk1.RotX = -20:RightSlingShot1.TimerEnabled = False
    End Select
    RStep1 = RStep1 + 1
End Sub

'Sub UpdateFlipperLogo_Timer
    'LFLogo.RotY = LeftFlipper.CurrentAngle
    'RFlogo.RotY = RightFlipper.CurrentAngle
'End Sub


Dim BallInHole, dBall, dZpos

Sub sw43_Hit
   PlaySound "fx_hole-enter", 0, 1, -0.1, 0.25
   Set dBall = ActiveBall
   dZpos=35
   sw43.TimerInterval = 4
   sw43.TimerEnabled = 1
End Sub

Sub sw43_Timer
    dBall.Z = dZpos
    dZpos = dZpos-4
    if dZpos < -30 Then
      sw43.timerenabled = 0
      BallInHole = BallInHole + 1
      sw43.DestroyBall

      vpmtimer.addtimer 700, "sw43HoleExit '"
    end if
End Sub



' SpinnerRod code from Cyperpez and http://www.vpforums.org/index.php?showtopic=35497
Sub CheckSpinnerRod_timer()
	SpinnerRod.TransZ = sin( (spinner.CurrentAngle+180) * (2*PI/360)) * 5
	SpinnerRod.TransX = -1*(sin( (spinner.CurrentAngle- 90) * (2*PI/360)) * 5)
End Sub


'**********
' Spinner
'**********

Sub Spinner_Spin()
    PlaySound "fx_spinner", 0, 1, 0.1
	DOF 136, DOFPulse
    If Tilted Then Exit Sub
    FlashForMs SmallFlasher2, 500, 50, 0
    if NOT i122.state=LightStateOff then ' Super Spinner
      debug.print "Spinner Cnt is " & SpinCnt(CurPlayer)
      Spincnt(CurPlayer)=Spincnt(CurPlayer)+1
      AddScore 1000*SpinCnt(CurPlayer)
      if SpinCnt(CurPlayer) > 50 Then
        i122.state=LightStateOff
        DisplayI(29) ' Completed
        SpinCnt(CurPlayer) = 0
      else
        SpinV "Spinner Count" &  SpinCnt(CurPlayer)
      end if
    else
      SpinV "2K"
      AddScore 2000
    end if
End Sub

'********
' Bumper
'********

Sub Bumper1_Hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_bumper",109,DOFPulse,DOFContactors), 0, 1, pan(ActiveBall)
	DOF 138, DOFPulse
    RandomBD() ' If BD then move shot
    if NOT i61.state=LightStateOff then ' Super Bumpers score 200K for upto 50Hits
      BumperCnt(CurPlayer)=BumperCnt(CurPlayer)+1
      if BumperCnt(CurPlayer)=50 then
        DisplayI(28)
        AddScore(200000)
        i61.state=LightStateOff
      else
        if BumperCnt(CurPlayer) < 50 then
          AddScore(200000)
          if rnd*10 > 4 then
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("200k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else
            DisplayI(32)
          end if
        Else
          AddScore(50000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("50k",INT(RND*8)+3," "), 15, 14, 100, 14
        End If
      End if
    Else
      AddScore 5000 * (BumperColor(1) +1)
      bv(1)
    End If
End Sub

Sub Bumper2_Hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_bumper",110,DOFPulse,DOFContactors), 0, 1, pan(ActiveBall)
	DOF 140, DOFPulse
    RandomBD() ' If BD then move shot
    if NOT i61.state=LightStateOff then ' Super Bumpers score 200K for upto 50Hits
      BumperCnt(CurPlayer)=BumperCnt(CurPlayer)+1
      if BumperCnt(CurPlayer)=50 then
        DisplayI(28)
        AddScore(200000)
        i61.state=LightStateOff
      else
        if BumperCnt(CurPlayer) < 50 then
          AddScore(200000)
          if rnd*10 > 4 then
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("200k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else
            DisplayI(32)
          end if
        Else
          AddScore(50000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("50k",INT(RND*8)+3," "), 15, 14, 100, 14
        End If
      End if
    Else
      AddScore 5000 * (BumperColor(2) +1)
      bv(2)
    End If
End Sub

Sub Bumper3_Hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_bumper",107,DOFPulse,DOFContactors), 0, 1, pan(ActiveBall)
	DOF 137, DOFPulse
    RandomBD() ' If BD then move shot
    if NOT i61.state=LightStateOff then ' Super Bumpers score 200K for upto 50Hits
      BumperCnt(CurPlayer)=BumperCnt(CurPlayer)+1
      if BumperCnt(CurPlayer)=50 then
        DisplayI(28)
        AddScore(200000)
        i61.state=LightStateOff
      else
        if BumperCnt(CurPlayer) < 50 then
          AddScore(200000)
          if rnd*10 > 4 then
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("200k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else
            DisplayI(32)
          end if
        Else
          AddScore(50000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("50k",INT(RND*8)+3," "), 15, 14, 100, 14
        End If
      End if
    Else
      AddScore 5000 * (BumperColor(3) +1)
      bv(3)
    End if
End Sub

Sub Bumper4_Hit
    If Tilted Then Exit Sub
    PlaySound SoundFXDOF("fx_bumper",108,DOFPulse,DOFContactors), 0, 1, pan(ActiveBall)
	DOF 139, DOFPulse
    RandomBD() ' If BD then move shot
    if NOT i61.state=LightStateOff then ' Super Bumpers score 200K for upto 50Hits
      BumperCnt(CurPlayer)=BumperCnt(CurPlayer)+1
      if BumperCnt(CurPlayer)=50 then
        DisplayI(28)
        AddScore(200000)
        i61.state=LightStateOff
      else
        if BumperCnt(CurPlayer) < 50 then
          AddScore(200000)
          if rnd*10 > 4 then
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("200k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else
            DisplayI(32)
          end if
        Else
          AddScore(50000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("50k",INT(RND*8)+3," "), 15, 14, 100, 14
        End If
      End if
    Else
      AddScore 5000 * (BumperColor(4) +1)
      bv(4)
    End If
End Sub

Sub ResetBumpers()
Dim i
    BumperCnt(CurPlayer) = 100
    Flasher1.Visible = 0
    Flasher2.Visible = 0
    Flasher3.Visible = 0
    Flasher4.Visible = 0
    SetLightColor B1L, "white", 0
    SetLightColor B2L, "white", 0
    SetLightColor B3L, "white", 0
    SetLightColor B4L, "white", 0
    for i = 1 to 4
      BumperColor(i)=0
    next
End Sub

Sub AwardExtraBall()
  debug.print "AwardExtraBall..."
'    If NOT bExtraBallWonThisBall Then
	    DOF 121, DOFPulse
        ExtraBallsAwards(CurPlayer) = ExtraBallsAwards(CurPlayer) + 1
        bExtraBallWonThisBall = True
        DisplayI(15)
        GiEffect 1
  '      LightEffect 2
'    END If
     debug.print "XBalls=" & ExtraBallsAwards(CurPlayer)
End Sub


Sub EffectTrigger1_Hit()
    FlashForMs Flasher11, 1000, 50, 0
End Sub

Sub EffectTrigger2_Hit()
    FlashForMs Flasher12, 1000, 50, 0
End Sub
'============================

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

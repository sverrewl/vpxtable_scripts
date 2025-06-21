'        ****************************************************************
'
'
'          ########  ##     ##  ######  ##     ##    ##       ########
'          ##     ## ##     ## ##    ## ##     ##    ##       ##
'          ##     ## ##     ## ##       ##     ##    ##       ##
'          ########  ##     ##  ######  #########    ##       ######
'          ##   ##   ##     ##       ## ##     ##    ##       ##
'          ##    ##  ##     ## ##    ## ##     ##    ##       ##
'          ##     ##  #######   ######  ##     ##    ######## ########
'
'                                 Rush LE v1.0 - By DrHotWing
'        ****************************************************************
' 4 coins required to play
'
' Collect blue albums by hitting blue album targets (2), cash them in by hitting Time Machine to begin 'The Spirit of Radio' Mode.
' Spirit of Radio Mode - Shooting the solid blue arrow shots decreases the radio waves, re-tune the radio by hitting the spinner (re-lights all arrows), cash out the bonus by hitting the dead end.
' Cashing out the bonus 5 times ends the mode.
'
' Collect red albums by hitting red album targets (2), cash them in by hitting Time Machine to begin 'The Big Money' Mode.
' The Big Money Mode- Shoot the red lit arrow shots to advance through 3 rounds to complete the mode.
'
' Collect green albums by hitting green album targets (2), cash them in by hitting Time Machine to begin 'Limelight' Mode.
' Limelight Mode- All green arrow shots start to advance your fame percentage, with 2 recommended shots that will flash faster that are worth more points and more fame progress.
' Get to 100% fame to win the mode. 15 second timer resets with every shot. If time expires you  have 15 seconds to hit time machine or mode ends.
'
' Hitting both green lock targets (L/R sides of R Ramp) lights the lock light. Shoot the ramp to light the green top kicker light. One hit into the top kicker when light is lit begins
' the Far Cry Multiball.
'
' Hit the Time Machine 4x begins the SubDivisions multiball.
'
' Chase the blinking instrument drop targets to drop all three and reveal the hidden target. Shoot any white arrow to reset all three targets.

' Roll the Bones Mystery - Roll over all 3 "roll the bones" return lanes to earn a mystery award
'
' 1-2-3 combos - Hit the targets in order then hit the Time Machine for award, go through two rounds before hitting Time Machine to light the super jackpot.
'
' Drum POPs and Headlong Flight Multiball - Hitting the POP bumpers will advance through 4 different levels, each level will change the POP bumpers color; (green) multiplies each hit 2x, (yellow)
' is the drum solo mode. You have 15 seconds to advance to the next level but each POP bumper hit will increase the time. If time runs out before advancing the POP bumpers reset. If you make It
' you advance to the (red) level which is the Headlong Flight Multiball.
'
' R-U-S-H Targets - Complete two rounds (solid lights/flashing Lights) to activate the Bastille Day HurryUp where you have so much time to hit the targets to score bonus points before time runs out.
'
'
'******************** SPECIAL THANK YOU *****************************************************
' Merlin RTP - coding and support
' mason - Improved the playfield graphics
' Tombg - LUT options in the F12 menu, coding, and support
' muhahaha aka kenji - PUP and code support
' LoadedWeapon - coding and support
' Lovinity Hearts - coding and support
' supergibson - support

Option Explicit
Randomize

'*************************************************************************
' // User Settings in F12 menu //
'*************************************************************************
Dim LUTimage

Sub Table1_OptionEvent(ByVal eventId)
    ' -- pause processes not needed to run --
    If eventId = 1 Then DisableStaticPreRendering = True
    DMDTimer.Enabled = False ' stop FlexDMD timer
    Controller.Pause = True

    ' --- Table LUT ---
    LUTimage = Table1.Option("Table Color Setting (LUT)", 0, 15, 1, 0, 0, Array("New ColorLUT","Normal","Warm Dark","Warm Bright","Warm Vivid Soft","Warm Vivid Hard","Natural Balanced","Natural High Contrast","THX Standard","Punchy","Desaturated","Washed Out","Natural Dark","Bassgeige","Blacklight","B&W Comic Book"))
    Table1.ColorGradeImage = "LUT" & LUTimage

    ' -- start paused processes --
    DMDTimer.Enabled = True
    Controller.Pause = False
    If eventId = 3 Then DisableStaticPreRendering = False

Dim SidewallChoice: SidewallChoice = 0
Dim RailChoice: RailChoice = True
'//////////////F12 Menu//////////////
' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reset
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
  If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If
    DMDTimer.Enabled = False ' stop FlexDMD timer
    Controller.Pause = True

      RailChoice = Table1.Option("Rails Visible", 0, 1, 1, 1, 0, Array("False", "True (Default)"))
  SetRails RailChoice

  SidewallChoice = Table1.Option("Sidewall Art", 0, 1, 1, 1, 0, Array("Desktop", "PinCab"))
  SetSidewall SidewallChoice
    DMDTimer.Enabled = True
    Controller.Pause = False
        If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
  End Sub

Sub SetRails(Opt)
  Select Case Opt
    Case 0:
      LeftRail.Visible = 0
      RightRail.Visible = 0
      'Primitive13.visible= 0
            'Primitive130.visible=0
    Case 1:
      LeftRail.Visible = 1
      RightRail.Visible = 1
      'Primitive13.visible=1
            'Primitive130.visible=1
  End Select
End Sub

Sub SetSidewall(Opt)
  Select Case Opt
    Case 0:
      PinCab_Blades.visible = 0
      'Wall055.sidevisible = 1
      'Wall056.sidevisible = 1
      'Wall055.Image = "SidewallLeft"
      'Wall056.Image = "SidewallRight"
    Case 1:
      'Wall055.sidevisible = 0
      'Wall056.sidevisible = 0
      'PinCab_Blades.Image = "Side_OFF"
      PinCab_Blades.visible = 1
  End Select
End Sub


'**************************
'   PinUp Player USER Config
'**************************

dim PuPDMDDriverType: PuPDMDDriverType=0   ' 0=LCD DMD, 1=RealDMD 2=FULLDMD (large/High LCD)
dim useRealDMDScale : useRealDMDScale=1    ' 0 or 1 for RealDMD scaling.  Choose which one you prefer.
dim useDMDVideos    : useDMDVideos=true   ' true or false to use DMD splash videos.
Dim pGameName       : pGameName="rush_le"  'pupvideos foldername

Dim bSongSelection
Dim SongNR
Dim nClockPosition(4)
Dim nCoin(5)
Dim nVidChg(3)
Dim bBMmode
Dim bmModeL2
Dim bmModeL3
Dim bllMode
Dim bSpiritMode
Dim sCheck(6)
Dim rBones(5)
Dim SoloState
Dim TMcomboL1
Dim TMcomboL2
Dim BumperLevel

Const nMaxSongs = 6
Const fMusicVolume = 0.8

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim bEnablePup: bEnablePup = True

'-----------------------------------------------------------------------------------------------

Const BallSize = 50    ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1
Const SongVolume = 0.1 ' 1 is full volume. Value is from 0 to 1

' Load the core.vbs for supporting Subs and functions

On Error Resume Next
ExecuteGlobal GetTextFile("core.vbs")
If Err Then MsgBox "Can't open core.vbs"
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0


'----- Shadow Options -----
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow, 1 = enable dynamic ball shadow


' Define any Constants
Const cGameName = "rush_le"
Const TableName = "Rush LE"
Const myVersion = "1.0.0"
Const MaxPlayers = 4     ' from 1 to 4
Const BallSaverTime = 20 ' in seconds
Const BallsPerGame = 3   ' usually 3 or 5
Const MaxMultiballs = 4  ' max number of balls during multiballs

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusGuitarActive(4)
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim PlayerScore(4)
Dim HighScore(5)
Dim HighScoreName(5)
Dim BallsInLockFarCry(4)
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim bAttractMode
Dim mBalls2Eject
Dim bAutoPlunger

' Define Game Control Variables
Dim BallsOnPlayfield
Dim BallsInLock
Dim BallsInHole
Dim BOT
Dim ballrolleron
' Define Game Flags
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim MBtm1Count
Dim bMusicOn
Dim bJustStarted
Dim bJackpot
Dim plungerIM
Dim LastSwitchHit
Dim Mode(4,10)
Dim AlbumColor
Dim AlbumsCollected
Dim bTimeMachineReady
Dim cTimeMachineReady
Dim dTimeMachineReady
Dim ComboLightOne
Dim ComboLightTwo
Dim ComboLightThree
Dim LevelofCombo
Dim ThreetoOne



'****************************************
'   USER OPTIONS
'****************************************

Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim StagedFlippers : StagedFlippers = 0         ' Staged Flippers. 0 = Disabled, 1 = Enabled


' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    PuPInit
    LoadEM
  Dim i
    Const IMPowerSetting = 50 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFXDOF("fx_kicker", 141, DOFPulse, DOFContactors), SoundFXDOF("fx_solenoid", 141, DOFPulse, DOFContactors)
        .CreateEvents "plungerIM"
    End With
    Loadhs      ' load saved values, highscore, names, jackpot
'Init main variables
    For i = 1 To MaxPlayers
        PlayerScore(i) = 0
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    Next

    bFreePlay = False ' Do not change this setting
  FlBumperFadeTarget(1) = 0
  FlBumperFadeTarget(2) = 0
  FlBumperFadeTarget(3) = 0
    bAttractMode = False
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bGameInPlay = False
    bMusicOn = True
    BallsOnPlayfield = 0
  bMultiBallMode = False
  bAutoPlunger = False
    BallsInLock = 0
  SoloState = 0
  BallsInLockFarCry(CurrentPlayer) = 0
  Flasher001.visible = 0
  Flasher002.visible = 0
    BallsInHole = 0
  LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bJustStarted = True
    GiOff
    StartAttractMode
End Sub



'**********************************
'General Math Functions
'**********************************

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


'******************************************************
'MECHANICAL SOUNDS
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

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////

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

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
'Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / tableheight - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioFade patched
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 5) 'was 10
	Else
		AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
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

' Thalamus, AudioPan patched
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 5) 'was 10
	Else
		AudioPan = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = CSng((ball.velz) ^ 2)
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


'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
    If Keycode = AddCreditKey Then
      StopAttractMode
      AttractEarlyClear
      AddCoins
            If(Tilted = False) Then
            End If
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
        PlaySoundAt "fx_reload", plunger
    ChangeVideo
    StopSong

    if bSongSelection And Not hsbModeActive Then ExitSongSelection
    End If

    if keycode = LeftFlipperKey And bSongSelection And Not hsbModeActive Then
      UpdateSongLeft
    End If

    if keycode = RightFlipperKey And bSongSelection And Not hsbModeActive Then
      UpdateSongRight
    End If

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

    If keycode = LeftFlipperKey Then
      SolLFlipper True
    End If

    If keycode = RightFlipperKey Then
      SolRFlipper True
    End If
      End If

           If keycode = StartGameKey Then
        soundStartButton()
                If(nCoin(CurrentPlayer) = 4) Then
                    If(BallsOnPlayfield = 0) Then
                        ResetForNewGame()
            UpdateMusicNow
            pDMDLabelHide "c4"
            PuPEvent 520
            PuPEvent 301
            End If
                    End If
                Else
                    If(nCoin(CurrentPlayer) < 4) Then
                        If(BallsOnPlayfield = 0) Then
                       End If
        End If
      End If
        If hsbModeActive Then HighScoreProcessKey(keycode)
        If keycode = 35 Then ResetHS 'H key
 End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode = PlungerKey Then
        Plunger.Fire
        PlaySoundAt "fx_plunger", plunger
        If bBallInPlungerLane Then PlaySoundAt "fx_fire", plunger
    End If

    If hsbModeActive Then
        Exit Sub
    End If

' Table specific

    If bGameInPLay AND NOT Tilted Then
    If keycode = LeftFlipperKey Then
      SolLFlipper False
      'SolULFlipper False
    End If

    If keycode = RightFlipperKey Then
      SolRFlipper False
    End If
    End If


End Sub

Sub AddCoins
  nCoin(CurrentPlayer) = nCoin(CurrentPlayer) + 1
  if nCoin(CurrentPlayer) > 4 Then nCoin(CurrentPlayer) = 4

  Select Case nCoin(CurrentPlayer)
    Case 1
      pDMDLabelShow "c1"
      Playsound "coin_in_1"
    Case 2
      pDMDLabelHide "c1"
      pDMDLabelShow "c2"
      PlaySound "coin_in_2"
    Case 3
      pDMDLabelHide "c2"
      pDMDLabelShow "c3"
      PlaySound "coin_in_3"
    Case 4
      pDMDLabelHide "c3"
      pDMDLabelShow "c4"
      PlaySound "coin_in_1"
  End Select
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


' xxxxxxx Flipper Code xxxxxxxxxxx


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
        RightFlipper2.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    RightFlipper2.RotateToStart
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


'******************************************************
' Flippers Polarity
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF)
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
End Sub



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
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)
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
Const EOSTnew = 1.2
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
Const EOSReturn = 0.035

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
    Dim b, BOT
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

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************

'******************************************************
'   RUBBER  DAMPENERS
'******************************************************
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

Dim RubbersD
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

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
  Public Print, debugOn
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


Sub CorTimer_Timer()
  Cor.Update
End Sub

'******************************************************
'SLINGSHOT CORRECTION FUNCTIONS
'******************************************************

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

  AddSlingsPt 0, 0.00, - 4
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4
End Sub

Sub AddSlingsPt(idx, aX, aY)
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
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled Then aBall.Velx = RotVxVy(0)
      If Enabled Then aBall.Vely = RotVxVy(1)
    End If
  End Sub
End Class


'*********
' TILT
'*********

Sub CheckTilt
    Tilt = Tilt + TiltSensitivity
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity) AND(Tilt <15) Then
  PuPEvent 504
    End if
    If Tilt> 15 Then 'If more that 15 then TILT the table
        Tilted = True
  PuPEvent 505
        DisableTable True
        TiltRecoveryTimer.Enabled = True
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    If Tilt> 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        GiOff
        LightSeqTilt.Play SeqAllOff
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        GiOn
        LightSeqTilt.StopPlay
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0) Then
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
End Sub

' xxxxxx Music sub xxxxxxxx

Dim Song, UpdateMusic
Song = ""

Sub PlaySong(name)
    If bMusicOn Then
        If Song <> name Then
            StopSound Song
            Song = name
            PlaySound Song, -1, SongVolume
        End If
    End If
End Sub

Sub StopSong
    If bMusicOn Then
        StopSound Song
        Song = ""
    End If
End Sub

' xxxxxxx GI Lighting xxxxxxxxxx

Dim OldGiState
OldGiState = -1   'start witht the Gi off

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = 1 Then
            GiOff
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
    For each bulb in aBumperLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    DOF 127, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    For each bulb in aBumperLights
        bulb.State = 0
    Next
End Sub

' GI, light & flashers sequence effects

Sub GiEffect(n)
    Dim ii
    Select Case n
        Case 0 'all off
            LightSeqGi.Play SeqAlloff
        Case 1 'all blink
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 10, 10
        Case 4 'all blink once
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 4, 1
    End Select
End Sub

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
        Case 4 'up 1 time
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 8, 1
        Case 5 'up 2 times
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqUpOn, 8, 2
        Case 6 'down 1 time
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqDownOn, 8, 1
        Case 7 'down 2 times
            LightSeqInserts.UpdateInterval = 4
            LightSeqInserts.Play SeqDownOn, 8, 2
    End Select
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.06, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'********************************************
'   JP's VP10 Rolling Sounds
'********************************************

Const tnob = 11 ' total number of balls
Const lob = 0   'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub GameTimer_timer
RollingUpdate
End Sub



Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)

        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 10
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        End If
    Next
End Sub


Sub RHelp1_Hit()
    StopSound "fx_metalrolling"
    PlaySoundAtBall "fx_ballrampdrop"
End Sub

Sub RHelp2_Hit()
    StopSound "fx_metalrolling"
    PlaySoundAtBall"fx_ballrampdrop"
End Sub


' xxxxxxxxxxxx Initialise the Table for a new Game xxxxxxxxxxxxx

Sub ResetForNewGame()
    Dim i

    bGameInPlay = True
    StopAttractMode
  BallsInLockFarCry(CurrentPlayer) = 0
  Flasher001.visible = 0
  Flasher002.visible = 0
  gi003.state = 0
    GiOn
    CurrentPlayer = 1
    PlayersPlayingGame = 1
  SoloState = 0
    bOnTheFirstBall = True

    For i = 1 To MaxPlayers
        PlayerScore(i) = 0
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    Next
    Tilt = 0
    Game_Init()
    InitSongSelection
StopSong
PlaySound ""
    vpmtimer.addtimer 1500, "FirstBall '"
End Sub

Sub FirstBall
    ' reset the table for a new ball
    ResetForNewPlayerBall()
    ' create a new ball in the shooters lane
    CreateNewBall()
End Sub

Sub ResetForNewPlayerBall()
    bBallSaverReady = True
  ClockBonus = 0
  AlbumBonus = 0
  ComboBonus = 0
    ResetNewBallVariables
    ResetNewBallLights()
  CheckFarCryLock
End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()
  LightSeqAttract.StopPlay
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1
    ' kick it out..
    PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4
    If BallsOnPlayfield> 1 Then
        bMultiBallMode = True
        bAutoPlunger = True
    End If
End Sub


' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
    'and eject the first ball
    CreateMultiballTimer_Timer
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
' wait if there is a ball in the plunger lane
    If bBallInPlungerLane Then
        Exit Sub
    Else
        If BallsOnPlayfield < MaxMultiballs Then
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

    Dim AwardPoints, TotalBonus, TotalClock, TotalAlb, TotalCom, ii

Sub EndOfBall()
    bOnTheFirstBall = False
  BonesClear
  StopAllMusic
    AwardPoints = 0
    TotalBonus = 0
  pDMDLabelShow "Aeob1"
  If(Tilted = False) Then
    vpmtimer.addtimer 1000, "BonusAddClock '"
    Else
    EndOfBall2
    End If
End Sub

Sub BonusAddClock()
    pDMDLabelHide "Aeob1"
    TotalClock = ClockBonus * 10000
        TotalBonus = TotalBonus + TotalClock
    pDMDLabelShow "Beob2"
    ShowClockScore
    vpmtimer.addtimer 2000, "BonusAddAlbum '"
End Sub

Sub BonusAddAlbum
    pDMDLabelHide "Beob2"
        TotalAlb = AlbumBonus * 25000
        TotalBonus = TotalBonus + TotalAlb
    pDMDLabelShow "Ceob3"
    HideClockScore
    ShowAlbumScore
    vpmtimer.addtimer 2000, "BonusAddCombo '"
End Sub

Sub BonusAddCombo
    pDMDLabelHide "Ceob3"
        TotalCom = ComboBonus * 50000
        TotalBonus = TotalBonus + TotalCom
    pDMDLabelShow "Deob4"
    HideAlbumScore
    ShowComboScore
    vpmtimer.addtimer 2000, "BonusAddTotal '"
End Sub

Sub BonusAddTotal
    pDMDLabelHide "Deob4"
    pDMDLabelShow "Eeob5"
    HideComboScore
    ShowBonusRound
    AddScore TotalBonus
    vpmtimer.addtimer 3000, "EndOfBall2 '"
End Sub

Sub ShowBonusRound
PuPlayer.LabelSet pDMD,"Bon"," "& TotalBonus &" ",1,"{'mt':2,'xpos':50 , 'ypos':80}"
End Sub

Sub HideBonusRound
PuPlayer.LabelSet pDMD,"Bon"," "& TotalBonus &" ",0,"{'mt':2,'xpos':50 , 'ypos':80}"
End Sub

Sub ShowClockScore
PuPlayer.LabelSet pDMD,"Bon"," "& TotalClock &" ",1,"{'mt':2,'xpos':50 , 'ypos':80}"
End Sub

Sub HideClockScore
PuPlayer.LabelSet pDMD,"Bon"," "& TotalClock &" ",0,"{'mt':2,'xpos':50 , 'ypos':80}"
End Sub

Sub ShowAlbumScore
PuPlayer.LabelSet pDMD,"Bon"," "& TotalAlb &" ",1,"{'mt':2,'xpos':50 , 'ypos':80}"
End Sub

Sub HideAlbumScore
PuPlayer.LabelSet pDMD,"Bon"," "& TotalAlb &" ",0,"{'mt':2,'xpos':50 , 'ypos':80}"
End Sub

Sub ShowComboScore
PuPlayer.LabelSet pDMD,"Bon"," "& TotalCom &" ",1,"{'mt':2,'xpos':50 , 'ypos':80}"
End Sub

Sub HideComboScore
PuPlayer.LabelSet pDMD,"Bon"," "& TotalCom &" ",0,"{'mt':2,'xpos':50 , 'ypos':80}"
End Sub

Sub EndOfBall2()
  HideBonusRound
  pDMDLabelHide "Eeob5"
    Tilted = False
    Tilt = 0
    DisableTable False 'enable bumpers and slingshots
    If(ExtraBallsAwards(CurrentPlayer) <> 0) Then
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
            LightShootAgain.State = 0
            LightShootAgain2.State = 0
        End If
        ' reset the playfield for the new ball
        ResetForNewPlayerBall()
        CreateNewBall()
    Else ' no extra balls
        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1
        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0) Then
            CheckHighScore()
        Else
            EndOfBallComplete()
        End If
    End If
End Sub

Sub EndOfBallComplete()
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
    StopSong
    PlaySound ""
    PuPEvent 803
    PuPEvent 302
        vpmtimer.addtimer 3000, "EndOfGame() '"

    Else
        ' set the next player
        CurrentPlayer = NextPlayer
        ResetForNewPlayerBall()
        CreateNewBall()
        If PlayersPlayingGame> 1 Then
            PlaySound "vo_player" &CurrentPlayer
         End If
    End If
End Sub

Sub EndOfGame()
  HideAllLabels
    LightSeqAttract.StopPlay
    bGameInPLay = False
    bJustStarted = False
    SolLFlipper 0
    SolRFlipper 0
    GiOff
    StartAttractMode
End Sub


Function BallsFunc
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer)
    If tmp> BallsPerGame Then
        BallsFunc = BallsPerGame
    Else
        BallsFunc = tmp
    End If
End Function

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

Sub Drain_Hit()
    Drain.DestroyBall
    BallsOnPlayfield = BallsOnPlayfield - 1
    PlaySoundAt "fx_drain", Drain
     If Tilted Then
        StopEndOfBallMode
    End If
    If(bGameInPLay = True) AND(Tilted = False) Then
        If(bBallSaverActive = True) Then
  AddMultiball 1
  bAutoPlunger = True

        Else

      If(BallsOnPlayfield = 1)Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True)then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                End If
            End If
            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0) Then
        CheckMusicTimer.Enabled = 0
        StopAllMusic
        InitSongSelection
                 StopEndOfBallMode
                vpmtimer.addtimer 200, "EndOfBall '"
            End If
        End If
    End If
End Sub

Sub Trigger1_Hit()
If bAutoPlunger Then
        PlungerIM.AutoFire
        DOF 121, DOFPulse
        PlaySoundAt "fx_fire", Trigger1
        PuPEvent 802
        bAutoPlunger = False
    End If

    bBallInPlungerLane = True
    If(bBallSaverReady = True) AND(BallSaverTime <> 0) And(bBallSaverActive = False) Then
        EnableBallSaver BallSaverTime
        Else
        Trigger1.TimerEnabled = 1
    End If
End Sub

' The ball is released from the plunger

Sub Trigger1_UnHit()
    bBallInPlungerLane = False
  if CheckMusicTimer.Enabled = 0 Then CheckMusicTimer.Enabled = 1
End Sub


Sub Trigger1_Timer
    trigger1.TimerEnabled = 0
End Sub

Sub EnableBallSaver(seconds)
     bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
    BallSaverTimer.Interval = 1000 * seconds
    BallSaverTimer.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.State = 2
    LightShootAgain2.BlinkInterval = 160
    LightShootAgain2.State = 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimer_Timer()
    BallSaverTimer.Enabled = False
    bBallSaverActive = False
  LightShootAgain.State = 0
  LightShootAgain2.State = 0
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    BallSaverSpeedUpTimer.Enabled = False
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
    LightShootAgain2.BlinkInterval = 80
    LightShootAgain2.State = 2
End Sub




' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

Sub AddScore(points)
    If Tilted Then Exit Sub
    PlayerScore(CurrentPlayer) = PlayerScore(CurrentPlayer) + points
    if Not bEnablePuP Then Exit sub
    PuPlayer.LabelSet pDMD,"CurrScore",FormatScore(PlayerScore(CurrentPlayer)),1,""
End Sub


Sub AwardExtraBall()
    DOF 121, DOFPulse
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    LightShootAgain.State = 1
    LightShootAgain2.State = 1
    LightEffect 2
End Sub

'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 100000 End If
    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If
    x = LoadValue(TableName, "HighScore2")
    If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 100000 End If
    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If
    x = LoadValue(TableName, "HighScore3")
    If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 100000 End If
    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If
    x = LoadValue(TableName, "HighScore4")
    If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 100000 End If
    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "") then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If
    x = LoadValue(TableName, "HighScore5")
    If(x <> "") then HighScore(4) = CDbl(x) Else HighScore(4) = 100000 End If
    x = LoadValue(TableName, "HighScore5Name")
    If(x <> "") then HighScoreName(4) = x Else HighScoreName(4) = "EEE" End If
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
    SaveValue TableName, "HighScore5", HighScore(4)
    SaveValue TableName, "HighScore5Name", HighScoreName(4)
End Sub


' xxxxxxx High Score Initals Entry Functions xxxxxxxxx

Dim hsbModeActive
Dim hsEnteredName
Dim hsEnteredDigits(3)
Dim hsCurrentDigit
Dim hsValidLetters
Dim hsCurrentLetter
Dim hsLetterFlash

Sub CheckHighscore()
  Dim tmp
  tmp = PlayerScore(CurrentPlayer)

  If tmp > HighScore(3) Then
    vpmtimer.addtimer 2000, "PlaySound ""vo_contratulationsgreatscore"" '"
    HighScore(3) = tmp
    HideAllLabels
    'enter player's name
    HighScoreEntryInit()
    DOF 103, DOFPulse
  Else
    EndOfBallComplete
  End If
End Sub


Sub HighScoreEntryInit()
    hsbModeActive = True
  pDMDLabelShow "hs"
    hsLetterFlash = 0

    hsEnteredDigits(0) = " "
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ'<>*+-/=\^0123456789`" ' ` is back arrow
    hsCurrentLetter = 1
    HighScoreDisplayNameNow()

    HighScoreFlashTimer.Interval = 250
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
        playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0) then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters) ) then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "`") then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3) then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit> 0) then
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
    Dim i, TempStr

    TempStr = " >"
    if(hsCurrentDigit> 0)then TempStr = TempStr & hsEnteredDigits(0)
    if(hsCurrentDigit> 1)then TempStr = TempStr & hsEnteredDigits(1)
    if(hsCurrentDigit> 2)then TempStr = TempStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3)then
        if(hsLetterFlash <> 0)then
            TempStr = TempStr & "_"
        else
            TempStr = TempStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit <1)then TempStr = TempStr & hsEnteredDigits(1)
    if(hsCurrentDigit <2)then TempStr = TempStr & hsEnteredDigits(2)

    TempStr = TempStr & "< "
  pDMDHighScore0 "Player 1 - Enter Initials", Mid(TempStr, 2, 5), 9999,""
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2) then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = False
    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ") then
        hsEnteredName = "RUS"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
    GameOver_Part2
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

Sub ResetHS
    HighScore(0) = 500000
    HighScoreName(0) = "AAA"
    HighScore(1) = 300000
    HighScoreName(1) = "AAA"
    HighScore(2) = 200000
    HighScoreName(2) = "AAA"
    HighScore(3) = 100000
    HighScoreName(3) = "AAA"
    HighScore(4) = 50000
    HighScoreName(4) = "AAA"
    savehs
End Sub

Sub GameOver_Part2
    Savehs
    GiOff
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
  RightFlipper2.RotateToStart
    PlayersPlayingGame = 0
    CurrentPlayer = 0
  CheckMusicTimer.Enabled = 0
    StopAllMusic
  pupevent 302
  pDMDLabelHide "hs"
  endofgame
End Sub


Function FormatScore(ByVal Num)
    dim i
    dim NumString

  Dim tmpScore

  if Num=0 then
    FormatScore="00"
    Exit Function
  End if

    NumString = CStr(abs(Num))

    For i = Len(NumString)-3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1))then
            NumString = left(NumString, i) & "," & right(NumString, Len(NumString)-i)
        end if
    Next
  tmpScore = " " &NumString &" "
    FormatScore = tmpScore
End function


' xxxxxxxxx Table info & Attract Mode xxxxxxxxxxx


Sub StartAttractMode
  StopVideos
  HideAllLabels
  ClearHS
  nCoin(CurrentPlayer) = 0
    StartLightSeq
  StartScoreAttractMode
End Sub

Sub StopAttractMode
  ClearHS
    LightSeqAttract.StopPlay
  bAttractMode = False
  PupEvent 303
End Sub

Sub StartScoreAttractMode
  pDMDLabelHide "CurrScore"
  pDMDLabelHide "BallValue"
  PupEvent 300
  ClearPupDMDHighTimer1.enabled=True
End Sub

Sub ClearPupDMDHighTimer1_Timer()
  PuPEvent 303
  PuPEvent 304
  PuPlayer.LabelShowPage pDMD,1,0,""
  pDMDHighScore "" & HighScoreName(0), "" & HighScore(0), 9999,""
  pDMDHighScore2 "" & HighScoreName(1), "" & HighScore(1), 9999,""
  pDMDHighScore3 "" & HighScoreName(2), "" & HighScore(2), 9999,""
  pDMDHighScore4 "" & HighScoreName(3), "" & HighScore(3), 9999,""
  pDMDHighScore5 "" & HighScoreName(4), "" & HighScore(4), 9999,""
  ClearPupDMDHighTimer2.Enabled=True
  ClearPupDMDHighTimer1.Enabled=False
End Sub

Sub ClearPupDMDHighTimer2_Timer()
  pDMDLabelHide "c2"
  PuPEvent 305
  PuPEvent 306
  ClearHS
  ClearPupDMDHighTimer2.Enabled=False
End Sub

Sub AttractEarlyClear
  pDMDLabelHide "CurrScore"
  pDMDLabelHide "BallValue"
  ClearPupDMDHighTimer1.Enabled=False
  ClearPupDMDHighTimer2.Enabled=False
  ClearHS
  PuPEvent 303
  PuPEvent 305
  PuPEvent 307
End Sub

Sub ClearHS
pDMDLabelHide "HSMessage"
pDMDLabelHide "HSInitials"
pDMDLabelHide "HSMessage2"
pDMDLabelHide "HSInitials2"
pDMDLabelHide "HSMessage3"
pDMDLabelHide "HSInitials3"
pDMDLabelHide "HSMessage4"
pDMDLabelHide "HSInitials4"
pDMDLabelHide "HSMessage5"
pDMDLabelHide "HSInitials5"
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqAttract.UpdateInterval = 25
    LightSeqAttract.Play SeqBlinking, , 5, 150
    LightSeqAttract.Play SeqRandom, 40, , 4000
    LightSeqAttract.Play SeqAllOff
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
End Sub

  'colors
  Dim red, orange, amber, yellow, darkgreen, green, blue, darkblue, purple, white, base

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

  Sub SetLightColor(n, col, stat)
    Select Case col
      Case red
        n.color = RGB(18, 0, 0)
        n.colorfull = RGB(255, 0, 0)
      Case orange
        n.color = RGB(18, 3, 0)
        n.colorfull = RGB(255, 64, 0)
      Case amber
        n.color = RGB(193, 49, 0)
        n.colorfull = RGB(255, 153, 0)
      Case yellow
        n.color = RGB(18, 18, 0)
        n.colorfull = RGB(255, 255, 0)
      Case darkgreen
        n.color = RGB(0, 8, 0)
        n.colorfull = RGB(0, 64, 0)
      Case green
        n.color = RGB(0, 18, 0)
        n.colorfull = RGB(0, 255, 0)
      Case blue
        n.color = RGB(0, 18, 18)
        n.colorfull = RGB(0, 255, 255)
      Case darkblue
        n.color = RGB(0, 8, 8)
        n.colorfull = RGB(0, 64, 64)
      Case purple
        n.color = RGB(128, 0, 128)
        n.colorfull = RGB(255, 0, 255)
      Case white
        n.color = RGB(255, 252, 224)
        n.colorfull = RGB(193, 91, 0)
      Case white
        n.color = RGB(255, 252, 224)
        n.colorfull = RGB(193, 91, 0)
      Case base
        n.color = RGB(255, 197, 143)
        n.colorfull = RGB(255, 255, 236)
    End Select
    If stat <> -1 Then
      n.State = 0
      n.State = stat
    End If
  End Sub

' xxxxxxxxxxxx Rainbow Lights Code xxxxxxxxxxxxx

Dim RainbowStep, RainbowColor, aRainbowLights
Set aRainbowLights = Nothing
RainbowStep = 1
RainbowColor = RGB(&HE8, &H14, &H16)

Sub RainbowTimer_Timer()
    Select Case RainbowStep
        Case 1: RainbowColor = RGB(255, 0, 0)  ' Red
        Case 2: RainbowColor = RGB(255, 64, 0)  ' Orange
        Case 3: RainbowColor = RGB(255, 255, 0)  ' Yellow
        Case 4: RainbowColor = RGB(0, 255, 0)  ' Green
        Case 5: RainbowColor = RGB(0, 128, 255)  ' Blue
        Case 6: RainbowColor = RGB(0, 0, 255)  ' Indigo
        Case 7: RainbowColor = RGB(128, 0, 255)  ' Violet
    End Select

    ' Apply the color change to each light
    Dim a
    For Each a In aRainbowLights
        a.color = RainbowColor
    Next

    ' Increment step and loop back after 7
    RainbowStep = RainbowStep + 1
    If RainbowStep > 7 Then RainbowStep = 1
End Sub

Sub RainbowLightsOn(n)  'Make a collection of lights n perform a rainbow
    Set aRainbowLights = n
    RainbowTimer.Enabled = True
    Dim a
    For Each a In aRainbowLights
        a.state = 1
    Next
End Sub

Sub RainbowLightsOff()  'Stop the rainbow on aRainbowLights
    RainbowTimer.Enabled = False
    Dim a
    For Each a In aRainbowLights
        a.state = 0
    Next
    Set aRainbowLights = Nothing
End Sub



' xxxxxxxxx Table Specific Script Starts Here xxxxxxxxxxxx

Dim ClockBonus, AlbumBonus, ComboBonus

Sub Game_Init() 'called at the start of a new game
    Dim i, j
  ClearHS
  HideAllLabels
  ResetDropTargets
  ClockBonus = 0
  AlbumBonus = 0
  ComboBonus = 0

  For i = 1 to 4
    nClockPosition(i) = 12
  Next

  For i = 1 to 4
    For j = 0 to 9
      Mode(i,j) = 0
    Next
  Next

    TurnOffPlayfieldLights()
  pDMDSetPage pScores, "InitGame"
  InitSongSelection
    StartMode1
End Sub

Sub ChangeVideo
  nVidChg(CurrentPlayer) = nVidChg(CurrentPlayer) + 1
  if nVidChg(CurrentPlayer) > 3 Then nVidChg(CurrentPlayer) = 1

  Select Case nVidChg(CurrentPlayer)
    Case 1
      PuPEvent 700
    Case 2
      PuPEvent 701
    Case 3
      PuPEvent 702
  End Select
End Sub

Sub StopVideos
  PuPEvent 704
  PuPEvent 705
  PuPEvent 706
End Sub

Sub ResetDropTargets
  sw13.IsDropped = False
  sw13.visible = 1
  sw12.IsDropped = False
  sw12.visible = 1
  sw11.IsDropped = False
  sw11.visible = 1
  sw20.IsDropped = False
  sw20.visible = 1
  TargetGuitarTimer.Enabled=True
End Sub


' xxxxxxxxxxxxx Drum Code xxxxxxxxxxxxxx

Dim ClockTime

Sub AdvanceClock
  nClockPosition(CurrentPlayer) = nClockPosition(CurrentPlayer) + 1
  if nClockPosition(CurrentPlayer) > 12 Then nClockPosition(CurrentPlayer) = 1

  Select Case nClockPosition(CurrentPlayer)
    Case 1
      primitive027.image = "drum_12"
      SoloState = SoloState + 1
      ClockBonus = ClockBonus + 1
      pDMDLabelHide "Cclockpm"
      pDMDLabelShow "Cmid"
      HideCTime
      CheckDrumSolo
    Case 2
      primitive027.image = "drum_1"
      pDMDLabelHide "Cmid"
      pDMDLabelShow "Cclockpm"
      ShowCTime
      ClockTime = 1
    Case 3
      primitive027.image = "drum_2"
      ClockTime = 2
      ShowCTime
    Case 4
      primitive027.image = "drum_3"
      ClockTime = 3
      ShowCTime
    Case 5
      primitive027.image = "drum_4"
      ClockTime = 4
      ShowCTime
    Case 6
      primitive027.image = "drum_5"
      ClockTime = 5
      ShowCTime
    Case 7
      primitive027.image = "drum_6"
      ClockTime = 6
      ShowCTime
    Case 8
      primitive027.image = "drum_7"
      ClockTime = 7
      ShowCTime
    Case 9
      primitive027.image = "drum_8"
      ClockTime = 8
      ShowCTime
    Case 10
      primitive027.image = "drum_9"
      ClockTime = 9
    Case 11
      primitive027.image = "drum_10"
      ClockTime = 10
      ShowCTime
    Case 12
      primitive027.image = "drum_11"
      ClockTime = 11
      ShowCTime
  End Select
End Sub

Sub ShowCTime
PuPlayer.LabelSet pDMD,"Clot"," "& ClockTime &" ",1,"{'mt':2,'xpos':92 , 'ypos':41}"
End Sub

Sub HideCTime
PuPlayer.LabelSet pDMD,"Clot"," "& ClockTime &" ",0,"{'mt':2,'xpos':92 , 'ypos':41}"
End Sub

Sub CheckDrumSolo
      If SoloState > 12 Then SoloState = 0
    If SoloState = 0 Then
    Dsolo.enabled = False
    SoloState = 0
    HideSoloTime
    Scountdown = 15
    BumperLevel = 0
    SetLightColor BumperLight001,white, 1
    SetLightColor BumperLight002,white, 1
    SetLightColor BumperLight003,white, 1
    I035.state = 0
    I007.state = 0
    End If
  If SoloState = 4 Then
  BumperLevel = 1
  SetLightColor BumperLight001,green, 1
  SetLightColor BumperLight002,green, 1
  SetLightColor BumperLight003,green, 1
  End If
    If SoloState = 8 Then
  Dsolo.enabled = True
  BumperLevel = 2
  SetLightColor BumperLight001,yellow, 1
  SetLightColor BumperLight002,yellow, 1
  SetLightColor BumperLight003,yellow, 1
  I007.state = 2
  End If
    If SoloState = 10 Then
  BumperLevel = 3
  HideCTime
  PuPEvent 541
  SetLightColor BumperLight001,red, 2
  SetLightColor BumperLight002,red, 2
  SetLightColor BumperLight003,red, 2
  I007.state = 0
  I035.state = 2
  AddMultiball 3
  End If
End Sub

Sub ShowSoloTime
PuPlayer.LabelSet pDMD,"SoloT"," "& Scountdown &" ",1,"{'mt':2,'xpos':94 , 'ypos':46.5}"
End Sub

Sub HideSoloTime
PuPlayer.LabelSet pDMD,"SoloT"," "& Scountdown &" ",0,"{'mt':2,'xpos':94 , 'ypos':46.5}"
End Sub

Dim Scountdown
  Scountdown = 15
Sub Dsolo_timer
    ShowSoloTime
  if Dsolo.enabled= true Then
  Scountdown = Scountdown -1
    If Scountdown = 0 Then
    Dsolo.enabled = False
    SoloState = 0
    HideSoloTime
    Scountdown = 15
    BumperLevel = 0
    SetLightColor BumperLight001,white, 1
    SetLightColor BumperLight002,white, 1
    SetLightColor BumperLight003,white, 1
    I035.state = 0
    I007.state = 0
    End If
  End If
End Sub



' xxxxxxxxxxxxxxxxxxxxxxxxx

Sub InitSongSelection
  bSongSelection = True
  SongNR = 1
  PupEvent 401
End Sub

Sub StopEndOfBallMode()     'this sub is called after the last ball is drained
End Sub

Sub ResetNewBallVariables() 'reset variables for a new ball or player
  bBallSaverReady = True
  DMDUpdateBallNumber Balls
  bTimeMachineReady = False
  cTimeMachineReady = False
  dTimeMachineReady = False
End Sub

Sub ResetNewBallLights()    'turn on or off the needed lights before a new ball is released
  gi1.state = 1
  gi2.state = 1
  gi3.state = 1
  gi4.state = 1
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   Initial table setup  (Shoot for a mode)

Sub StartMode1()
  If bBMmode = False And bmModeL2 = False And bmModeL3 = False And bllMode = False Then
    pDMDLabelShow "Talb"
  End If
AlbumColor = 0
AlbumsCollected = 0
DropAlbumCount
ShowAlbumLabel
ComboLightOne = True
TMcomboL1 = False
TMcomboL2 = False
ThreetoOne = True
BumperLevel = 0
Dsolo.enabled = False
SetLightColor I016,yellow,2
SetLightColor I028,yellow,2
SetLightColor I026,yellow,2
SetLightColor I019,yellow,2
SetLightColor I65,red,2
SetLightColor I009,red,2
SetLightColor I012,green,2
SetLightColor I011,blue,2
SetLightColor I010,blue,2
SetLightColor I110,green,2
SetLightColor I005,orange,2
SetLightColor I006,orange,2
SetLightColor I034,green,1
SetLightColor I014,green,2
SetLightColor I013,green,2
SetLightColor BumperLight001,white, 1
SetLightColor BumperLight002,white, 1
SetLightColor BumperLight003,white, 1
BonusGuitarActive(CurrentPlayer) = 0
BumperLight001.State=1
BumperLight002.State=1
BumperLight003.State=1
Fl001.State = 1
Fl002.State = 1
Fl003.State = 1
Fl004.State = 1
Fl005.State = 1
End Sub

' xxxxxxx Rush Targets xxxxxxxxxxxxxx

Dim RUSHlevel
Dim RushCountMode
Dim RushTotal
RUSHlevel = False
RushCountMode = False
RushTotal = 0

Sub sw001_Hit() 'R hit
  If RushCountMode = False And I45.state = 0 Then
  I45.State = 1
    pDMDLabelShow "RLit"
  HitRUSH
  End If
    If RUSHlevel = True Then
    I45.state = 2
    pDMDLabelHide "RLit"
    pDMDLabelShow "yRLit"
    HitRUSH
    End If
      If RushCountMode = True And RushTime.enabled = True Then
      AddScore 500000
      RushTotal = RushTotal + 500000
      I45.state = 0
      pDMDLabelHide "yRLit"
      HitRUSH
      End If
End Sub

Sub sw002_Hit()  'U hit
  If RushCountMode = False And I44.state = 0 Then
  I44.State = 1
    pDMDLabelShow "ULit"
  HitRUSH
  End If
    If RUSHlevel = True Then
    I44.state = 2
    pDMDLabelHide "ULit"
    pDMDLabelShow "yULit"
    HitRUSH
    End If
      If RushCountMode = True And RushTime.enabled = True Then
      AddScore 500000
      RushTotal = RushTotal + 500000
      I44.state = 0
      pDMDLabelHide "yULit"
      HitRUSH
      End If
End Sub

Sub sw003_Hit()  'S hit
  If RushCountMode = False And I43.state = 0 Then
  I43.State = 1
    pDMDLabelShow "SLit"
  HitRUSH
  End If
    If RUSHlevel = True Then
    I43.state = 2
    pDMDLabelHide "SLit"
    pDMDLabelShow "ySLit"
    HitRUSH
    End If
      If RushCountMode = True And RushTime.enabled = True Then
      AddScore 500000
      RushTotal = RushTotal + 500000
      I43.state = 0
      pDMDLabelHide "ySLit"
      HitRUSH
      End If
End Sub

Sub sw004_Hit()  'H hit
  If RushCountMode = False And I42.state = 0 Then
  I42.State = 1
    pDMDLabelShow "HLit"
  HitRUSH
  End If
    If RUSHlevel = True Then
    I42.state = 2
    pDMDLabelHide "HLit"
    pDMDLabelShow "yHLit"
    HitRUSH
    End If
      If RushCountMode = True And RushTime.enabled = True Then
      AddScore 500000
      RushTotal = RushTotal + 500000
      I42.state = 0
      pDMDLabelHide "yHLit"
      HitRUSH
      End If
End Sub

Sub HitRUSH
  If I45.state = 1 And I44.state = 1 And I43.state = 1 And I42.state = 1 Then
  RUSHlevel = True
  End If
    If I45.state = 2 And I44.state = 2 And I43.state = 2 And I42.state = 2 Then
    I099.state = 2
    FlasherLight003.state = 2
    RushCountMode = True
    RUSHlevel = False
    RushTime.enabled = True
    PuPEvent 542
    End If
      If RushTime.enabled = True And I45.state = 0 And I44.state = 0 And I43.state = 0 And I42.state = 0 Then
      ResetRush
      End If
End Sub

Dim rushdrain
  rushdrain = 60
Sub RushTime_timer
  pDMDLabelShow "tbx"
  ShowRushTimer
  If RushTime.enabled = True Then
  rushdrain = rushdrain - 1
  Select Case rushdrain
  case 0:
    ResetRush
  case 60:
End Select
End If
End Sub

Sub ResetRush
  ShowRushTotal
  PuPEvent 543
  RushTime.enabled = False
' HideRushTotal
  HideRushTimer
  pDMDLabelHide "tbx"
  ClearRUSHLetters
  I45.state = 0
  I44.state = 0
  I43.state = 0
  I42.state = 0
  RushCountMode = False
  RUSHlevel = False
  rushdrain = 60
  I099.state = 0
  FlasherLight003.state = 0
  RushTotal = 0
  vpmtimer.addtimer 3500, "HideRushTotal '"
End Sub

Sub ClearRUSHLetters
    pDMDLabelHide "RLit"
    pDMDLabelHide "ULit"
    pDMDLabelHide "SLit"
    pDMDLabelHide "HLit"
    pDMDLabelHide "yRLit"
    pDMDLabelHide "yULit"
    pDMDLabelHide "ySLit"
    pDMDLabelHide "yHLit"
End Sub

Sub ShowRushTimer
PuPlayer.LabelSet pDMD,"Rtim"," "& rushdrain &" ",1,"{'mt':2,'xpos':14 , 'ypos':45}"
End Sub

Sub HideRushTimer
PuPlayer.LabelSet pDMD,"Rtim"," "& rushdrain &" ",0,"{'mt':2,'xpos':14 , 'ypos':45}"
End Sub

Sub ShowRushTotal
PuPlayer.LabelSet pDMD,"Rtot"," "& RushTotal &" ",1,"{'mt':2,'xpos':49.5 , 'ypos':75}"
End Sub

Sub HideRushTotal
PuPlayer.LabelSet pDMD,"Rtot"," "& RushTotal &" ",0,"{'mt':2,'xpos':49.5 , 'ypos':75}"
End Sub

' xxxxxxxxxxxxxxxxxxxxxxxxxx

Sub Lock1_Hit() 'Far Cry Multiball lock
  if ActiveBall.VelY < 0 Then
    AddScore 12000
    I014.State = 1
    CheckLockState
  End If
End Sub

Sub Lock2_Hit()  'Far Cry Multiball lock
  if ActiveBall.VelY < 0 Then
    AddScore 12000
    I013.State = 1
    CheckLockState
  End If
End Sub

Sub CheckLockState
  If I013.State = 1 And I014.State = 1 Then
    SetLightColor I017,green,2
    PuPEvent 502 'FarCry MB locked
  End If
End Sub

Sub ShootRamp ' Far Cry MB shoot ramp check
  If I017.State = 2 Then
    FlasherLight006.State = 1
    I017.State = 0
  End If
End Sub


Sub Tlane2_Hit()
  if ActiveBall.VelY < 0 Then
    If I019.state = 2 And ComboLightOne = True Then
    CheckLevel2
    End If
    If bBMmode = False And bmModeL2 = False And bmModeL3 = False And bllMode = False And bSpiritMode = False Then
      I65.state = 1 ' Red album
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
      AlbumCollect
      ModeCheck
    End If
    If bBMmode = True Or bmModeL2 = True Or bmModeL3 = True Or bllMode = True Or bSpiritMode = True Then
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
    End If
  End If
End Sub

Sub RightRampTrigger_Hit()
  if ActiveBall.VelY < 0 Then
  ShootRamp
    If bBMmode = False And bmModeL2 = False And bmModeL3 = False And bllMode = False And bSpiritMode = False Then
      I009.state = 1 ' Red album
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
      AlbumCollect
      ModeCheck
    End If
    If bBMmode = True Or bmModeL2 = True Or bmModeL3 = True Or bllMode = True Or bSpiritMode = True Then
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
    End If
  End If
End Sub


Sub Trigger5_Hit()
  if ActiveBall.VelY < 0 Then
    If bBMmode = False And bmModeL2 = False And bmModeL3 = False And bllMode = False And bSpiritMode = False Then
      I011.state = 1 ' Blue Album
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
      AlbumCollect
      ModeCheck
    End If
    If bBMmode = True Then
    I031.state = 1
    AlbumsCollected = AlbumsCollected + 1
    AlbumBonus = AlbumBonus + 1
    ShowAlbumLabel
    AddScore 20000
    BMmodeCheckL1
    End If
    If bmModeL2 = True Then
    I031.state = 1
    AlbumsCollected = AlbumsCollected + 1
    AlbumBonus = AlbumBonus + 1
    ShowAlbumLabel
    AddScore 20000
    BMmodeCheckL2
    End If
    If bmModeL3 = True Then
    I031.state = 1
    AlbumsCollected = AlbumsCollected + 1
    AlbumBonus = AlbumBonus + 1
    ShowAlbumLabel
    AddScore 20000
    BMmodeCheckL3
    End If
  End If
End Sub

Sub Trigger6_Hit()
  if ActiveBall.VelY < 0 Then
      If bBMmode = False And bmModeL2 = False And bmModeL3 = False And bllMode = False And bSpiritMode = False Then
      I010.state = 1 ' Blue Album
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
      AlbumCollect
      ModeCheck
      End If
        If bBMmode = True Then
        I030.state = 1
        AlbumsCollected = AlbumsCollected + 1
        AlbumBonus = AlbumBonus + 1
        ShowAlbumLabel
        AddScore 20000
        BMmodeCheckL1
        End If
          If bmModeL2 = True Then
          I030.state = 1
          AlbumsCollected = AlbumsCollected + 1
          AlbumBonus = AlbumBonus + 1
          ShowAlbumLabel
          AddScore 20000
          BMmodeCheckL2
          End If
            If bmModeL3 = True Then
            I030.state = 1
            AlbumsCollected = AlbumsCollected + 1
            AlbumBonus = AlbumBonus + 1
            ShowAlbumLabel
            AddScore 20000
            BMmodeCheckL3
            End If
              If I025.state = 2 And ComboLightTwo = True Then
              Level3Combo
              End If
  End If
End Sub


Sub Trigger8_Hit()
  if ActiveBall.VelY < 0 Then
    If bBMmode = False And bmModeL2 = False And bmModeL3 = False And bllMode = False And bSpiritMode = False Then
      I012.state = 1 ' Green album
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
      AlbumCollect
      ModeCheck
    End If
    If bBMmode = True Or bmModeL2 = True Or bmModeL3 = True Or bllMode = True Or bSpiritMode = True Then
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
    End If
  End If
End Sub

Sub Trigger7_Hit()
  if ActiveBall.VelY < 0 Then
    If bBMmode = False And bmModeL2 = False And bmModeL3 = False And bllMode = False And bSpiritMode = False Then
      I110.state = 1 ' Green Album
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
      AlbumCollect
      ModeCheck
    End If
    If bBMmode = True Or bmModeL2 = True Or bmModeL3 = True Or bllMode = True Or bSpiritMode = True Then
      AlbumsCollected = AlbumsCollected + 1
      AlbumBonus = AlbumBonus + 1
      ShowAlbumLabel
      AddScore 20000
    End If
  End If
End Sub


' xxxxxxxxxx Album Code xxxxxxxxxxxxxxxx

Sub AlbumCollect
  AlbumColor = AlbumColor + 1
    If AlbumColor > 3 Then AlbumColor = 3
    Select Case AlbumColor
    Case 0
      pDMDLabelShow "Talb"
    Case 1
      pDMDLabelHide "Talb"
      pDMDLabelShow "alb1"
      If bTimeMachineReady = False And cTimeMachineReady = False And dTimeMachineReady = False Then
      pDMDLabelShow "Oalb"
      End If
    Case 2
      pDMDLabelShow "alb2"
    Case 3
      pDMDLabelShow "alb3"
    End Select
End Sub

Sub ShowAlbumLabel
PuPlayer.LabelSet pDMD,"AlbCt"," "& AlbumsCollected &" ",1,"{'mt':2,'xpos':1.3 , 'ypos':61.7}"
End Sub

Sub HidetAlbumLabel
PuPlayer.LabelSet pDMD,"AlbCt"," "& AlbumsCollected &" ",0,"{'mt':2,'xpos':1.3 , 'ypos':61.7}"
End Sub

Sub DropAlbumCount
  pDMDLabelHide "alb1"
  pDMDLabelHide "alb2"
  pDMDLabelHide "alb3"
  pDMDLabelHide "albct"
End Sub

Sub ModeCheck
  If I009.state = 1 And I65.state = 1 And cTimeMachineReady = False And dTimeMachineReady = False Then
    bTimeMachineReady = True
    I002.State = 2
    I150.state = 2
    I21.State = 2 'Light Big Money light
    pDMDLabelShow "Big" ' small label
    pDMDLabelHide "Srad" ' small label
    pDMDLabelHide "Lime" ' small label
    pDMDLabelHide "Oalb" ' 1 album label
    pDMDLabelShow "bm1"
    PuPEvent 503
  End If
  If I012.state = 1 And I110.state = 1 And bTimeMachineReady = False And cTimeMachineReady = False Then
    dTimeMachineReady = True
    I002.State = 2 ' Light time machine red
    I20.State = 2  ' Light Limelight light
    I150.state = 2
    pDMDLabelHide "Oalb" ' 1 album label
    pDMDLabelHide "big" ' small label
    pDMDLabelHide "Srad" ' small label
    pDMDLabelShow "Lime" ' small label
    pDMDLabelShow "lim1"
    PuPEvent 503
  End If
  If I010.state = 1 And I011.state = 1 And bTimeMachineReady = False And dTimeMachineReady = False Then
    cTimeMachineReady = True
    I002.State = 2  ' Light time machine red
    I150.state = 2
    I19.State = 2  ' Light Spirit of Radio light
    pDMDLabelHide "big" ' small label
    pDMDLabelHide "Lime" ' small label
    pDMDLabelHide "Oalb" ' 1 album label
    pDMDLabelShow "Srad" ' small label
    pDMDLabelShow "sr1"
    PuPEvent 503
  End If
End Sub

' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Sub Trigger004_Hit()
If bmModeL2 = True Then
    I032.state = 1
    BMmodeCheckL2
End If
If bmModeL3 = True Then
    I032.state = 1
    BMmodeCheckL3
End If
If bSpiritMode = True And I032.state = 1 Then
    SpiritTotal = SpiritTotal + 100000
    ShowSORValue
    AddScore 100000
    I032.state = 0
    SpiritAllDark
End If
End Sub

' xxxxxxxxxxxxxxxxxxxxx Time Machine xxxxxxxxxxxxxxxxxxxxxxxx
Sub tm1_Hit()
   AddScore 15000
  MBtm1Count=MBtm1Count+1
  if MBtm1Count=1 Then
    pDMDLabelShow "hit3"
  End If
  if MBtm1Count=2 Then
    pDMDLabelShow "hit2"
  End If
  if MBtm1Count=3 Then
    pDMDLabelShow "hitr"
    SDivLit
    I034.State = 2
    I150.state = 2
  End If
  if MBtm1Count=4 then
  AddMultiball 1
  PuPEvent 506
  MBtm1Count=0
  ClearSubDivisions
  End If
End Sub

Sub ClearSubDivisions
    pDMDLabelHide "hit3"
    pDMDLabelHide "hit2"
    pDMDLabelHide "hitr"
  I034.State = 0
  I150.state = 0
End Sub

Sub SDivLit
    pdmdlabelshow "subdiv"
    TriggerScript 5000, "SdLi"
  End Sub

  Sub SdLi
    pdmdlabelhide "subdiv"
  End Sub

Sub tm2_Hit()
    if ActiveBall.VelY < 0 Then
  if bTimeMachineReady = True Then
  BigMoneyMode ' Start Big Money Mode
    End If
  if dTimeMachineReady = True Then
  LimelightMode ' Start Limelight Mode
    End If
  if cTimeMachineReady = True Then
  SpiritMode ' Start Spirit of Radio Mode
    End If
  If bllMode = True And nofame.enabled = True Then
  nofame.enabled = False
  GreenArrow.enabled = True
  famedrain = 15
  greentimecount = 15
  I150.state = 0
    End If
      If TMcomboL1 = True Then
      TMcomboL1 = False
      ThreetoOne = True
      PuPEvent 538
      AddScore 3000000
      LevelofCombo = 0
      I003.state = 0
      ComboLevel1
      End If
        If TMcomboL2 = True Then
        TMcomboL2 = False
        ThreetoOne = True
        PuPEvent 537
        AddScore 10000000
        I003.state = 0
        LevelofCombo = 0
        ComboLevel1
        End If
  End If
End Sub

' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

'*****************
'Instrument collection
'***********
Sub CheckTargetGuitar
  If I13.state = 1 And I12.state = 1 And I11.state = 1 Then
  sw13.IsDropped = True
  sw13.visible = 0
  PuPEvent 512
  AddScore 275000
  BonusGuitarActive(CurrentPlayer) = 1
  End If
  If I11.state = 1 Then
  PuPEvent 514
  End If
  If I12.state = 1 Then
  PuPEvent 513
  End If
End Sub
Sub sw11_Hit()
  If I11.state = 2 Then
  I11.state = 1
  pDMDLabelShow "drumlit"
  sw11.IsDropped = True
  sw11.visible = 0
  AddScore 10000
  End If
  CheckTargetGuitar
End Sub
Sub sw12_Hit()
  If I12.state = 2 Then
  I12.state = 1
  pDMDLabelShow "guitlit"
  sw12.IsDropped = True
  sw12.visible = 0
  AddScore 10000
  End If
  CheckTargetGuitar
End Sub
Sub sw13_Hit()
  If I13.state = 2 Then
  I13.state = 1
  pDMDLabelShow "basslit"
  End If
  CheckTargetGuitar
End Sub
Sub sw20_Hit()
  If BonusGuitarActive(CurrentPlayer) = 1 Then
  I006.state = 2
  sw20.IsDropped = True
  sw20.visible = 0
  AddScore 1000000
  BonusGuitarActive(CurrentPlayer) = 2
    SetLightColor I006,white,2
    SetLightColor I032,white,2
    SetLightColor I107,white,2
  End If
End Sub
Sub LeftRampTrigger001_Hit()
  If BonusGuitarActive(CurrentPlayer) = 2 Then
  I006.state = 0
  I032.state = 0
  I107.state = 0
  pDMDLabelHide "basslit"
  pDMDLabelHide "guitlit"
  pDMDLabelHide "drumlit"
  BonusGuitarActive(CurrentPlayer) = 0
  I11.state = 0
  I12.state=0
  I13.state=0
  ResetDropTargets
  End If
    If bmModeL3 = True Then
    I006.state = 1
    BMmodeCheckL3
    End If
      If Fl7.state = 2 And TMcomboL2 = True Then
      Fl7.state = 0
      PuPEvent 540
      AddScore 20000000
      End If
End Sub

Sub Trigger001_Hit()
  if ActiveBall.VelY < 0 Then
    If I018.state = 2 And ComboLightTwo = True Then
    Level3Combo
    End If
  End If
End Sub

Sub LL1_Hit()
  If bllMode = True And GreenArrow.enabled = True Then
  greentimecount = 15
  FameCount = FameCount + 10
  PuPEvent 529
  AddScore 500000
  CheckLimelight
  End If
  If bSpiritMode = True And I62.state = 1 Then
  SpiritTotal = SpiritTotal + 100000
  ShowSORValue
  AddScore 100000
  I62.state = 0
  SpiritAllDark
  End If
End Sub

Sub LL2_Hit()
    If I028.state = 2 And ComboLightOne = True And ThreetoOne = True Then
    CheckLevel2
    End If
  If bllMode = True And GreenArrow.enabled = True Then
  greentimecount = 15
  FameCount = FameCount + 5
  PuPEvent 528
  AddScore 150000
  CheckLimelight
  End If
End Sub

Sub Trigger005_Hit()
  if ActiveBall.VelY < 0 Then
    If I029.state = 2 And ComboLightThree = True Then
    ThreetoOne = False
    ComboTM
    End If
  End If
End Sub

Sub Trigger006_Hit()
  if ActiveBall.VelY < 0 Then
    If I015.state = 2 And ComboLightTwo = True Then
    Level3Combo
    End If
  End If
End Sub

Sub LL3_Hit()
  if ActiveBall.VelY < 0 Then
    If I028.state = 2 And ComboLightOne = True Then
    CheckLevel2
    End If
      If I027.state = 2 And ComboLightThree = True Then
      ComboTM
      End If
        If bSpiritMode = True Then
        SpiritResetCheck
        End If
          If bllMode = True And GreenArrow.enabled = True Then
          greentimecount = 15
          FameCount = FameCount + 5
          PuPEvent 528
          AddScore 150000
          CheckLimelight
          End If
  End If
End Sub

Sub LL4_Hit()
  if ActiveBall.VelY < 0 Then
    If I026.state = 2 And ComboLightOne = True Then
    CheckLevel2
    End If
      If bSpiritMode = True Then
      PuPEvent 531
      SpiritTune
      End If
        If bllMode = True And GreenArrow.enabled = True Then
        greentimecount = 15
        FameCount = FameCount + 10
        PuPEvent 529
        AddScore 500000
        CheckLimelight
        End If
  End If
End Sub

Sub LL5_Hit()
    if ActiveBall.VelY < 0 Then
      If I021.state = 2 And ComboLightTwo = True Then
      Level3Combo
      End If
    End If
  If bllMode = True And GreenArrow.enabled = True Then
  greentimecount = 15
  FameCount = FameCount + 5
  PuPEvent 528
  AddScore 150000
  CheckLimelight
  End If
  If bSpiritMode = True And I107.state = 1 Then
  SpiritTotal = SpiritTotal + 100000
  ShowSORValue
  AddScore 100000
  I107.state = 0
  SpiritAllDark
  End If
End Sub


Sub LL6_Hit()
    If I023.state = 2 And ComboLightThree = True Then
    ComboTM
    End If
  If bllMode = True And GreenArrow.enabled = True Then
  greentimecount = 15
  FameCount = FameCount + 5
  PuPEvent 528
  AddScore 150000
  CheckLimelight
  End If
  If bSpiritMode = True And I006.state = 1 Then
  SpiritTotal = SpiritTotal + 100000
  ShowSORValue
  AddScore 100000
  I006.state = 0
  SpiritAllDark
  End If
End Sub

Sub LL7_Hit()
    If I016.state = 2 And ComboLightOne = True Then
    CheckLevel2
    End If
  If bllMode = True And GreenArrow.enabled = True Then
  greentimecount = 15
  FameCount = FameCount + 5
  PuPEvent 528
  AddScore 150000
  CheckLimelight
  End If
  If bSpiritMode = True And I103.state = 1 Then
  SpiritTotal = SpiritTotal + 100000
  ShowSORValue
  AddScore 100000
  I103.state = 0
  SpiritAllDark
  End If
End Sub


Sub Trigger003_Hit()
  If BonusGuitarActive(CurrentPlayer) = 2 Then
  I006.state = 0
  I032.state = 0
  I107.state = 0
  pDMDLabelHide "basslit"
  pDMDLabelHide "guitlit"
  pDMDLabelHide "drumlit"
  BonusGuitarActive(CurrentPlayer) = 0
  I11.state = 0
  I12.state=0
  I13.state=0
  sw13.IsDropped = False
  sw13.visible = 1
  sw12.IsDropped = False
  sw12.visible = 1
  sw11.IsDropped = False
  sw11.visible = 1
  sw20.IsDropped = False
  sw20.visible = 1
  TargetGuitarTimer.Enabled=True
  End If
End Sub

Sub Trigger002_Hit()
  If BonusGuitarActive(CurrentPlayer) = 2 Then
  I006.state = 0
  I032.state = 0
  I107.state = 0
  pDMDLabelHide "basslit"
  pDMDLabelHide "guitlit"
  pDMDLabelHide "drumlit"
  BonusGuitarActive(CurrentPlayer) = 0
  I11.state = 0
  I12.state=0
  I13.state=0
  sw13.IsDropped = False
  sw13.visible = 1
  sw12.IsDropped = False
  sw12.visible = 1
  sw11.IsDropped = False
  sw11.visible = 1
  sw20.IsDropped = False
  sw20.visible = 1
  TargetGuitarTimer.Enabled=True
  End If
    If bmModeL2 = True Then
    I107.state = 1
    BMmodeCheckL2
    End If
    If bmModeL3 = True Then
    I107.state = 1
    BMmodeCheckL3
    End If
End Sub

Sub Trigger007_Hit()
  If I024.state = 2 And ComboLightTwo = True Then
  Level3Combo
  End If
End Sub

Sub Trigger008_Hit()
ThreetoOne = True
End Sub


Sub TargetGuitarTimer_Timer()
  If I11.state = 0 And I12.state=0 And I13.state=0 Then
  I11.state = 2 And I12.state=0 And I13.state=0
  Elseif I11.state = 2 And I12.state=0 And I13.state=0 Then
  I11.state = 0 : I12.state=2 : I13.state=0
  Elseif I11.state = 2 And I12.state=1 And I13.state=0 Then
  I11.state = 0 : I12.state=1 : I13.state=2
  Elseif I11.state = 2 And I12.state=0 And I13.state=1 Then
  I11.state = 0 : I12.state=2 : I13.state=1
  Elseif I11.state = 2 And I12.state=1 And I13.state=1 Then
  I11.state = 2 : I12.state=1 : I13.state=1
  Elseif I11.state = 1 And I12.state=0 And I13.state=0 Then
  I11.state = 1 : I12.state=2 : I13.state=0

  Elseif I11.state = 1 And I12.state=1 And I13.state=0 Then
  I11.state = 1 : I12.state=1 : I13.state=2
  Elseif I11.state = 1 And I12.state=0 And I13.state=1 Then
  I11.state = 1 : I12.state=2 : I13.state=1
  Elseif I11.state = 0 And I12.state=1 And I13.state=1 Then
  I11.state = 2 : I12.state=1 : I13.state=1

  Elseif I11.state = 0 And I12.state=2 And I13.state=0 Then
  I11.state = 0 : I12.state=0 : I13.state=2
  Elseif I11.state = 0 And I12.state=2 And I13.state=1 Then
  I11.state = 2 : I12.state=0 : I13.state=1
  Elseif I11.state = 1 And I12.state=2 And I13.state=0 Then
  I11.state = 1 : I12.state=0 : I13.state=2
  Elseif I11.state = 1 And I12.state=2 And I13.state=1 Then
  I11.state = 1 : I12.state=2 : I13.state=1
  Elseif I11.state = 0 And I12.state=1 And I13.state=0 Then
  I11.state = 0 : I12.state=1 : I13.state=2

  Elseif I11.state = 0 And I12.state=0 And I13.state=2 Then
  I11.state = 0 : I12.state=0 : I13.state=2
  Elseif I11.state = 0 And I12.state=1 And I13.state=2 Then
  I11.state = 2 : I12.state=1 : I13.state=0
  Elseif I11.state = 1 And I12.state=0 And I13.state=2 Then
  I11.state = 1 : I12.state=2 : I13.state=0
  Elseif I11.state = 1 And I12.state=1 And I13.state=2 Then
  I11.state = 1 : I12.state=1 : I13.state=2
  Elseif I11.state = 0 And I12.state=0 And I13.state=1 Then
  I11.state = 2 : I12.state=0 : I13.state=1

  Elseif I11.state = 1 And I12.state=1 And I13.state=1 Then
  BonusGuitarActive(CurrentPlayer) = 1
  TargetGuitarTimer.Enabled=False
  End If
End Sub

Sub FarCryLock1
  Kicker2.destroyball
  Flasher001.visible = 1
  Flasher002.visible = 0
  Kicker003.CreateSizedBallWithMass BallSize / 2, BallMass
  Kicker003.Kick 0, 55
    PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease
  BallsInLockFarCry(CurrentPlayer) = 1
  gi003.state = 2
End Sub
Sub FarCryLock2
  Kicker2.destroyball
  Flasher001.visible = 1
  Flasher002.visible = 1
  Kicker003.CreateSizedBallWithMass BallSize / 2, BallMass
    Kicker003.Kick 0, 55
    PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease
  BallsInLockFarCry(CurrentPlayer) = 2
  gi003.state = 1
End Sub
Sub FarCryLock3 'Mball
  Kicker2.TimerEnabled = 1
  Flasher001.visible = 0
  Flasher002.visible = 0
  Kicker1.CreateSizedBallWithMass BallSize / 2, BallMass
  Kicker1.TimerEnabled = 1
  Kicker003.CreateSizedBallWithMass BallSize / 2, BallMass
    Kicker003.Kick 0, 55
  BallsOnPlayfield = BallsOnPlayfield + 2
    PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease
  BallsInLockFarCry(CurrentPlayer) = 0
  gi003.state = 0
End Sub
Sub CheckFarCryLock 'Curentplayer
  If BallsInLockFarCry(CurrentPlayer) = 0 Then
  Flasher001.visible = 0
  Flasher002.visible = 0
  gi003.state = 0
  End If
  If BallsInLockFarCry(CurrentPlayer) = 1 Then
  Flasher001.visible = 1
  Flasher002.visible = 0
  gi003.state = 2
  End If
  If BallsInLockFarCry(CurrentPlayer) = 2 Then
  Flasher001.visible = 1
  Flasher002.visible = 1
  End If
End Sub

Sub FarCryMBCheck
    If BallsInLockFarCry(CurrentPlayer) = 0 Then
      FarCryLock1
      FlasherLight006.State = 0
      SetLightColor I014,green,2
      SetLightColor I013,green,2
      PuPEvent 510
    Elseif BallsInLockFarCry(CurrentPlayer) = 1 Then
      FarCryLock2
      FlasherLight006.State = 0
      SetLightColor I014,green,2
      SetLightColor I013,green,2
      PuPEvent 510
    Elseif BallsInLockFarCry(CurrentPlayer) = 2 Then
      FarCryLock3
      FlasherLight006.State = 0
      SetLightColor I014,green,2
      SetLightColor I013,green,2
      PuPEvent 510
    End If
End Sub

Sub TLeftInlane_Hit()
    I95.state = 1
    PuPEvent 525
    RollBonesCheck
End Sub

Sub TRightInlane_Hit()
    I96.state = 1
    PuPEvent 525
    RollBonesCheck
End Sub

Sub TRightInlane2_Hit()
    I97.state = 1
    PuPEvent 525
    RollBonesCheck
End Sub

Sub RollBonesCheck
  If I95.state = 1 And I96.state = 1 And I97.state = 1 Then
    BonesAward
  End If
End Sub

Sub BonesAward
  rBones(CurrentPlayer) = rBones(CurrentPlayer) + 1
  if rBones(CurrentPlayer) > 4 Then rBones(CurrentPlayer) = 4

  Select Case rBones(CurrentPlayer)
    Case 1
      EnableBallSaver BallSaverTime
      I95.state = 0
      I96.state = 0
      I97.state = 0
    Case 2
      I95.state = 0
      I96.state = 0
      I97.state = 0
      I013.state = 1
      I014.state = 1
      I017.state = 2
      CheckFarCryLock
    Case 3
      AddScore 2500000
      I95.state = 0
      I96.state = 0
      I97.state = 0
    Case 4
      AddScore 5000000
      I95.state = 0
      I96.state = 0
      I97.state = 0
  End Select
End Sub

Sub BonesClear
      I95.state = 0
      I96.state = 0
      I97.state = 0
End Sub

Sub drumleft  'Bumper1 drum image
    pdmdlabelshow "drumL"
    TriggerScript 2000, "hideDL"
  End Sub

  Sub hideDL
    pdmdlabelhide "drumL"
  End Sub

Sub drummiddle 'Bumper2 drum image
    pdmdlabelshow "drumM"
    TriggerScript 2000, "hideM"
  End Sub

  Sub hideM
    pdmdlabelhide "drumM"
  End Sub

Sub drumright 'Bumper3 drum image
    pdmdlabelshow "drumR"
    TriggerScript 2000, "hideR"
  End Sub

  Sub hideR
    pdmdlabelhide "drumR"
  End Sub

' 1-2-3 Combos xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Sub ComboLevel1 ' 1's lit
ComboLightOne = True
ComboLightTwo = False
ComboLightThree = False
I024.state = 0
I021.state = 0
I025.state = 0
I015.state = 0
I023.state = 0
I027.state = 0
I029.state = 0
I019.state = 2
I016.state = 2
I026.state = 2
I028.state = 2
End Sub

Sub CheckLevel2 ' 2's lit
ComboLightOne = False
ComboLightTwo = True
ComboLightThree = False
AddScore 250000
ComboOne
I019.state = 0
I016.state = 0
I026.state = 0
I028.state = 0
I018.state = 2
I024.state = 2
I021.state = 2
I025.state = 2
I015.state = 2
End Sub

Sub Level3Combo ' 3's lit
ComboLightOne = False
ComboLightTwo = False
ComboLightThree = True
ComboTwo
AddScore 500000
I018.state = 0
I024.state = 0
I021.state = 0
I025.state = 0
I015.state = 0
I023.state = 2
I027.state = 2
I029.state = 2
End Sub

Sub ComboTM
  LevelofCombo = LevelofCombo + 1
    If LevelofCombo > 2 Then LevelofCombo = 0
    Select Case LevelofCombo
      Case 0
      Case 1
        PuPEvent 535
        ComboBonus = ComboBonus + 1
        AddScore 750000
        I003.state = 2
        TMcomboL1 = True
        ComboLevel1
      Case 2
        PuPEvent 536
        ComboBonus = ComboBonus + 1
        AddScore 750000
        I003.state = 2
        Fl7.state = 2
        TMcomboL1 = False
        TMcomboL2 = True
        ComboLevel1
      Case 3
    End Select
End Sub

Sub ComboOne
    pdmdlabelshow "Como"
    TriggerScript 4000, "hidec1"
  End Sub

  Sub hidec1
    pdmdlabelhide "Como"
  End Sub

Sub ComboTwo
    pdmdlabelshow "ComTw"
    TriggerScript 4000, "hidec2"
  End Sub

  Sub hidec2
    pdmdlabelhide "ComTw"
  End Sub

Sub ComboThree
    pdmdlabelshow "ComThr"
    TriggerScript 4000, "hidec3"
  End Sub

  Sub hidec3
    pdmdlabelhide "ComThr"
  End Sub


' MODES *****************************************************************************************

Sub ModeLightChange
SetLightColor I012,white, 1
SetLightColor I110,white, 1
SetLightColor I011,white, 1
SetLightColor I010,white, 1
SetLightColor I009,white, 1
SetLightColor I65,white, 1
RainbowLightsOn aMulti
End Sub

Sub ModeLightReset
RainbowLightsOff
SetLightColor I012,green, 2
SetLightColor I110,green, 2
SetLightColor I011,blue, 2
SetLightColor I010,blue, 2
SetLightColor I009,red, 2
SetLightColor I65,red, 2
End Sub

'*****Big Money Mode*******
Sub BigMoneyMode
    ModeLightChange
    DropAlbumCount
    AlbumColor = 0
    pDMDLabelShow "tbm" ' big label (main)
    pDMDLabelHide "Big" ' small label
    pDMDLabelHide "bm1" ' big label (pre)
    I21.State = 1
    I19.state = 0
    I20.state = 0
    I002.State = 0
    I150.state = 0
    PuPEvent 521
    StopAllMusic
        PlaySong "song-thebigmoney"
    bBMmode = True
    bmModeL2 = False
    bmModeL3 = False
    PuPEvent 501
    SetLightColor I030,red,2
    SetLightColor I005,red,2
    SetLightColor I031,red,2
End Sub

Sub BMmodeCheckL1
  If I030.state = 1 And I005.state =1 And I031.state =1 Then
  I030.state = 0 And I005.state=0 And I031.state=0
  PuPEvent 522
  AddScore 500000
  BigMoneyCount = BigMoneyCount + 500000
  ShowBMcount
  BigMoneyModeL2
  End If
End Sub

Sub BigMoneymodeL2
    pDMDLabelHide "Tbm"
    pDMDLabelShow "Tbml2"
    I21.State = 1
    I002.State = 0
    I150.state = 0
    I005.state = 0
    PuPEvent 521
    StopAllMusic
        PlaySong "song-thebigmoney"
    bBMmode = False
    bmModeL2 = True
    bmModeL3 = False
    SetLightColor I030,red,2
    SetLightColor I032,red,2
    SetLightColor I031,red,2
    SetLightColor I107,red,2
End Sub

Sub BMmodeCheckL2
  If I030.state = 1 And I032.state =1 And I031.state =1 And I107.state = 1 Then
  I030.state = 0 And I032.state=0 And I031.state=0 And I107.state = 0
  PuPEvent 523
  AddScore 1000000
  BigMoneyCount = BigMoneyCount + 1000000
  ShowBMcount
  BigMoneyModeL3
  End If
End Sub

Sub BigMoneymodeL3
    pDMDLabelHide "Tbml2"
    pDMDLabelShow "Tbml3"
    I21.State = 1
    I002.State = 0
    I150.state = 0
    bBMmode = False
    bmModeL2 = False
    bmModeL3 = True
    StopAllMusic
    PuPEvent 521
    SetLightColor I030,red,2
    SetLightColor I032,red,2
    SetLightColor I031,red,2
    SetLightColor I107,red,2
    SetLightColor I006,red,2
End Sub

Sub BMmodeCheckL3
  If I030.state = 1 And I032.state =1 And I031.state =1 And I107.state = 1 Then
  I030.state = 0 And I032.state=0 And I031.state=0 And I107.state = 0
  PuPEvent 524
  AddScore 1500000
  BigMoneyCount = BigMoneyCount + 1500000
  ShowBMcount
  BigMoneyCompleted
  End If
End Sub

Dim BigMoneyCount
  BigMoneyCount = 0
Sub ShowBMcount
PuPlayer.LabelSet pDMD,"CurrBM"," "& BigMoneyCount &" ",1,"{'mt':2,'xpos':4.5 , 'ypos':13.3}"
End Sub

Sub HideBMcount
PuPlayer.LabelSet pDMD,"CurrBM"," "& BigMoneyCount &" ",0,"{'mt':2,'xpos':4.5 , 'ypos':13.3}"
End Sub

Sub BigMoneyCompleted
    ModeLightReset
    SetLightColor I030,red,0
    SetLightColor I032,red,0
    SetLightColor I031,red,0
    SetLightColor I107,red,0
    SetLightColor I006,red,0
    I016.state = 2
    I018.state = 0
    I019.state = 2
    I21.state = 0
    AlbumColor = 0
    bTimeMachineReady = False
    bBMmode = False
    bmModeL2 = False
    bmModeL3 = False
    CloseModeLabels ' all big labels
    HideBMcount
    pDMDLabelShow "Talb"
End Sub


'******Limelight Mode*****************

Sub LimelightMode
    ModeLightChange
    DropAlbumCount
    AlbumColor = 0
    I20.State = 1
    I19.state = 0
    I21.state = 0
    I002.State = 0
    I150.state = 0
    PuPEvent 526
    pDMDLabelHide "lim1" ' big label (pre)
    pDMDLabelShow "lli" ' big label (main)
    pDMDLabelHide "Lime" ' small label
    I62.BlinkInterval = 50
    SetLightColor I62,green,2
    SetLightColor I005,green,2
    SetLightColor I006,green,2
    SetLightColor I103,green,2
    SetLightColor I032,green,2
    SetLightColor I031,green,2
    I030.BlinkInterval = 50
    SetLightColor I030,green,2
    SetLightColor I107,green,2
    bllMode = True
    GreenArrow.enabled = True
    StopAllMusic
    PlaySong "song-limelight"
End Sub


Dim greentimecount
  greentimecount = 15
Sub GreenArrow_timer
    ShowtimerLabel
  if GreenArrow.enabled= true Then
  greentimecount= greentimecount -1
  select case greentimecount
  case 0:
    PuPEvent 527
    GreenArrow.enabled= False
    nofame.enabled= True
  case 14:
  case 15:
end Select
end if
end sub

Dim famedrain
  famedrain = 15
Sub nofame_timer
  HidetimerLabel
  if nofame.enabled= true Then
  I150.state = 2
  famedrain= famedrain -1
  select case famedrain
  case 0:
    EndLimelight
  case 15:
end Select
end if
end sub

Sub EndLimelight
    ModeLightReset
    I20.State = 0
    I002.State = 0
    I150.state = 0
    I62.BlinkInterval = 150
    I030.BlinkInterval = 150
    I62.state = 0
    I005.state = 0
    I006.state = 0
    I103.state = 0
    I032.state = 0
    I031.state = 0
    I030.state = 0
    I107.state = 0
    I021.state = 0
    I028.state = 2
    I015.state = 0
    I029.state = 0
    bllMode = False
    GreenArrow.enabled = False
    nofame.enabled = False
    AlbumColor = 0
    dTimeMachineReady = False
    CloseModeLabels ' all big labels
    HidetFameLabel
    HidetimerLabel
    pDMDLabelShow "Talb"
End Sub

Dim FameCount
  FameCount = 0
Sub CheckLimelight
  ShowFameLabel
  If FameCount > 99 Then
    PuPEvent 530
    AddScore 750000
    EndLimelight
  End If
End Sub

Sub ShowtimerLabel
PuPlayer.LabelSet pDMD,"Currgreentimecount"," "& greentimecount &" ",1,"{'mt':2,'xpos':5 , 'ypos':15}"
End Sub

Sub HidetimerLabel
PuPlayer.LabelSet pDMD,"Currgreentimecount"," "& greentimecount &" ",0,"{'mt':2,'xpos':5 , 'ypos':15}"
End Sub


Sub ShowFameLabel
PuPlayer.LabelSet pDMD,"CurrFame"," "& FameCount &" ",1,"{'mt':2,'xpos':3 , 'ypos':23}"
End Sub

Sub HidetFameLabel
PuPlayer.LabelSet pDMD,"CurrFame"," "& FameCount &" ",0,"{'mt':2,'xpos':3 , 'ypos':23}"
End Sub


'*****Spirit of Radio Mode***********

Sub SpiritMode
    ModeLightChange
    DropAlbumCount
    AlbumColor = 0
    bSpiritMode = True
    pDMDLabelHide "Srad" ' small label
    pDMDLabelHide "sr1" ' big label (pre)
        pDMDLabelShow "sora" ' big label (main)
    I19.State = 1
    I20.state = 0
    I21.state = 0
    I002.State = 0
    I150.state = 0
    I030.BlinkInterval = 50
    I031.BlinkInterval = 50
    SetLightColor I030,blue,2
    SetLightColor I031,blue,2
    SetLightColor I032,blue,1
    SetLightColor I103,blue,1
    SetLightColor I107,blue,1
    SetLightColor I006,blue,1
    SetLightColor I62,blue,1
    SetLightColor I005,blue,1
    StopAllMusic
        PlaySong "Spirit"
    PuPEvent 511
    ShowSpiritLevel
    ShowSORValue
    SpiritLevel = 0
    SpiritTotal = 0
    SpiritResetCheck
End Sub

Sub SpiritResetCheck
    ShowSpiritLevel
    ShowSORValue
    If sCheck(CurrentPlayer) = 5 Then
      SpiritModeEnd
    End If
  sCheck(CurrentPlayer) = sCheck(CurrentPlayer) + 1
    Select Case sCheck(CurrentPlayer)
    Case 1
      I032.state = 1
      I103.state = 1
      I107.state = 1
      I006.state = 1
      I62.state = 1
      I005.state = 1
      SpiritLevel = 1
      SpiritTotal = SpiritTotal + 100000
'     PuPEvent 532
      AddScore 100000
    Case 2
      I032.state = 1
      I103.state = 1
      I107.state = 1
      I006.state = 1
      I62.state = 1
      I005.state = 1
      SpiritLevel = 2
      SpiritTotal = SpiritTotal + 100000
      PuPEvent 532
      AddScore 100000
    Case 3
      I032.state = 1
      I103.state = 1
      I107.state = 1
      I006.state = 1
      I62.state = 1
      I005.state = 1
      SpiritLevel = 3
      SpiritTotal = SpiritTotal + 100000
      PuPEvent 532
      AddScore 100000
    Case 4
      I032.state = 1
      I103.state = 1
      I107.state = 1
      I006.state = 1
      I62.state = 1
      I005.state = 1
      SpiritLevel = 4
      SpiritTotal = SpiritTotal + 250000
      PuPEvent 533
      AddScore 250000
    Case 5
      SpiritTotal = SpiritTotal + 500000
      PuPEvent 534
      AddScore 500000
    End Select
End Sub

Sub SpiritTune
    I032.state = 1
    I103.state = 1
    I107.state = 1
    I006.state = 1
    I62.state = 1
    I005.state = 1
End Sub

Sub SpiritAllDark
  If bSpiritMode = True And I032.state = 0 And I103.state = 0 And I107.state = 0 And I006.state = 0 And I62.state = 0 And I005.state = 0 Then
  SpiritResetCheck
  End If
End Sub

Sub SpiritModeEnd
    ModeLightReset
    bSpiritMode = False
    CloseModeLabels ' all big labels
    I032.state = 0
    I103.state = 0
    I107.state = 0
    I006.state = 0
    I62.state = 0
    I005.state = 0
    I19.state = 0
    I030.BlinkInterval = 150
    I031.BlinkInterval = 150
    I030.state = 0
    I031.state = 0
    AlbumColor = 0
    cTimeMachineReady = False
    HideSpiritLevel
    HideSORValue
    I027.state = 0
    I025.state = 0
    I026.state = 2
    I028.state = 2
    pDMDLabelShow "Talb"
End Sub

Dim SpiritLevel
Dim SpiritTotal

Sub ShowSpiritLevel
PuPlayer.LabelSet pDMD,"CLevel"," "& SpiritLevel &" ",1,"{'mt':2,'xpos':3.4 , 'ypos':23.7}"
End Sub

Sub HideSpiritLevel
PuPlayer.LabelSet pDMD,"CLevel"," "& SpiritLevel &" ",0,"{'mt':2,'xpos':3.4 , 'ypos':23.7}"
End Sub

Sub ShowSORValue
PuPlayer.LabelSet pDMD,"CTotal"," "& SpiritTotal &" ",1,"{'mt':2,'xpos':4.5 , 'ypos':16}"
End Sub

Sub HideSORValue
PuPlayer.LabelSet pDMD,"CTotal"," "& SpiritTotal &" ",0,"{'mt':2,'xpos':4.5 , 'ypos':16}"
End Sub

Sub CloseModeLabels
    pDMDLabelHide "sora"
    pDMDLabelHide "lli"
    pDMDLabelHide "tbm"
    pDMDLabelHide "tbml2"
    pDMDLabelHide "tbml3"
End Sub


'Slingshots
'****************************************************************

' RStep and LStep are the variables that increment the animation
Dim RStep, LStep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  Addscore 2000
  RSling1.Visible = 1
  Sling1.TransY =  - 20   'Sling Metal Bracket
  RStep = 0
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 10
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
  Addscore 2000
  LSling1.Visible = 1
  Sling2.TransY =  - 20   'Sling Metal Bracket
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
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


'Spinners
'*****************

Sub Spinner001_Spin
    PlaySoundAt "fx_spinner", Spinner001
    AddScore 1000
End Sub

'Kickers
'*****************

Sub kicker1_Hit()
  PlaySoundAt "fx_kicker_enter", Kicker1
    kicker1.TimerEnabled = 1
    AddScore 5000
End Sub

Sub kicker1_Timer
    kicker1.Kick 190, 15
  PlaySoundAt "fx_kicker", Kicker1
    kicker1.TimerEnabled = 0
End Sub

Sub kicker2_Hit()
  PlaySoundAt "fx_kicker_enter", Kicker2
   AddScore 5000
  If Flasherlight006.State = 1 Then
  FarCryMBCheck
  Else
    Kicker2.TimerEnabled = 1
  End If
    If bBMmode = True Then
    SetLightColor I005,red,1
    I005.state = 1
    BMmodeCheckL1
    End If
      if bllMode = True And I005.state = 2 Then
      SetLightColor I005,blue,1
      End If
        If bSpiritMode = True And I005.state = 1 Then
        SpiritTotal = SpiritTotal + 100000
        ShowSORValue
        AddScore 100000
        I005.state = 0
        SpiritAllDark
        End If
End Sub

Sub kicker2_Timer
    Kicker2.Kick 270, 30
  PlaySoundAt "fx_kicker", Kicker2
    Kicker2.TimerEnabled = 0
End Sub


Sub kicker3_Hit()
  PlaySoundAt "fx_kicker_enter", Kicker3
    Kicker3.TimerEnabled = 1
    AddScore 5000
End Sub

Sub kicker3_Timer
    Kicker3.Kick 1,55, 1.4
  PlaySoundAt "fx_kicker", Kicker3
    Kicker3.TimerEnabled = 0
End Sub



'Bumpers
'**************************

Sub Bumper001_hit()
PlaySoundAt"KickDrum1", Bumper001
drumleft
FlBumperFadeTarget(1) = 1
Bumper001.timerenabled = True
AdvanceClock
  If BumperLevel = 0 Then
  AddScore 10000
  End If
    If BumperLevel = 1 Then
    AddScore 20000
    End If
      If BumperLevel = 2 Then
      AddScore 30000
      Scountdown = Scountdown + 1
      End If
        If BumperLevel = 3 Then
        AddScore 50000
        End If
End sub

Sub Bumper001_timer
  FlBumperFadeTarget(1) = 0
End Sub

Sub Bumper002_hit()
drummiddle
PlaySoundAt"KickDrum2", Bumper002
FlBumperFadeTarget(2) = 1
Bumper002.timerenabled = True
AdvanceClock
  If BumperLevel = 0 Then
  AddScore 10000
  End If
    If BumperLevel = 1 Then
    AddScore 20000
    End If
      If BumperLevel = 2 Then
      I007.state = 2
      AddScore 30000
      Scountdown = Scountdown + 1
      End If
        If BumperLevel = 3 Then
        AddScore 50000
        End If
End sub

Sub Bumper002_timer
  FlBumperFadeTarget(2) = 0
End Sub

Sub Bumper003_hit()
drumright
PlaySoundAt"KickDrum3", Bumper003
FlBumperFadeTarget(3) = 1
Bumper003.timerenabled = True
AdvanceClock
  If BumperLevel = 0 Then
  AddScore 10000
  End If
    If BumperLevel = 1 Then
    AddScore 20000
    End If
      If BumperLevel = 2 Then
      I007.state = 2
      AddScore 30000
      Scountdown = Scountdown + 1
      End If
        If BumperLevel = 3 Then
        AddScore 50000
        End If
End sub

Sub Bumper003_timer
  FlBumperFadeTarget(3) = 0
End Sub


' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0 : DNA45 = (NightDay-10)/20 : DNA90 = 0 : DayNightAdjust = 0.4
Else
  DNA30 = (NightDay-10)/30 : DNA45 = (NightDay-10)/45 : DNA90 = (NightDay-10)/90 : DayNightAdjust = NightDay/25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt : For cnt = 1 to 6 : FlBumperActive(cnt) = False : Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight

FlInitBumper 1, "red"
FlInitBumper 2, "red"

' uncomment the statement below to change the color for all bumpers ###
' Dim ind : For ind = 1 to 5 : FlInitBumper ind, "green" : next

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True
  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1 : FlBumperFadeTarget(nr) = 1.1: FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr) : FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr) : Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr) : Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr) : FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr): FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
  ' set the color for the two VPX lights
  select case col
    Case "red"
      FlBumperSmallLight(nr).color = RGB(255,4,0) : FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
      FlBumperBigLight(nr).color = RGB(255,32,0) : FlBumperBigLight(nr).colorfull = RGB(255,32,0)
      FlBumperHighlight(nr).color = RGB(64,255,0)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
      FlBumperSmallLight(nr).TransmissionScale = 0
    Case "blue"
      FlBumperBigLight(nr).color = RGB(32,80,255) : FlBumperBigLight(nr).colorfull = RGB(32,80,255)
      FlBumperSmallLight(nr).color = RGB(0,80,255) : FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
      FlBumperSmallLight(nr).TransmissionScale = 0 : MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
      FlBumperHighlight(nr).color = RGB(255,16,8)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    Case "green"
      FlBumperSmallLight(nr).color = RGB(8,255,8) : FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
      FlBumperBigLight(nr).color = RGB(32,255,32) : FlBumperBigLight(nr).colorfull = RGB(32,255,32)
      FlBumperHighlight(nr).color = RGB(255,32,255) : MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
      FlBumperSmallLight(nr).TransmissionScale = 0.005
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    Case "orange"
      FlBumperHighlight(nr).color = RGB(255,130,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).color = RGB(255,130,0) : FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
      FlBumperBigLight(nr).color = RGB(255,190,8) : FlBumperBigLight(nr).colorfull = RGB(255,190,8)
    Case "white"
      FlBumperBigLight(nr).color = RGB(255,230,190) : FlBumperBigLight(nr).colorfull = RGB(255,230,190)
      FlBumperHighlight(nr).color = RGB(255,180,100) :
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
    Case "blacklight"
      FlBumperBigLight(nr).color = RGB(32,32,255) : FlBumperBigLight(nr).colorfull = RGB(32,32,255)
      FlBumperHighlight(nr).color = RGB(48,8,255) :
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    Case "yellow"
      FlBumperSmallLight(nr).color = RGB(255,230,4) : FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
      FlBumperBigLight(nr).color = RGB(255,240,50) : FlBumperBigLight(nr).colorfull = RGB(255,240,50)
      FlBumperHighlight(nr).color = RGB(255,255,220)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
    Case "purple"
      FlBumperBigLight(nr).color = RGB(80,32,255) : FlBumperBigLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).color = RGB(80,32,255) : FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).TransmissionScale = 0 :
      FlBumperHighlight(nr).color = RGB(32,64,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
  end select
End Sub

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 )* DayNightAdjust

  select case FlBumperColor(nr)

    Case "blue" :
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(38-24*Z,130 - 98*Z,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20  + 500 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z +0.97 * Z^3)
      Flbumperbiglight(nr).intensity = 45 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 10000 * (Z^3) / (0.5 + DNA90)

    Case "green"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(16 + 16 * sin(Z*3.14),255,16 + 16 * sin(Z*3.14)), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
      FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 6000 * (Z^3) / (1 + DNA90)

    Case "red"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 16 - 11*Z + 16 * sin(Z*3.14),0), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z*4,8-Z*8)

    Case "orange"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 100 - 22*z  + 16 * sin(Z*3.14),Z*32), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z*50, 0)

    Case "white"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
      FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 14 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255,255 - 20*Z,255-65*Z) : FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20*Z,255-65*Z)
      MaterialColor "bumpertopmat" & nr, RGB(255,235 - z*36,220 - Z*90)

    Case "blacklight"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 1, RGB(30-27*Z^0.03,30-28*Z^0.01, 255), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z^3
      Flbumperbiglight(nr).intensity = 40 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255-240*(Z^0.1),255 - 240*(Z^0.1),255) : FlBumperSmallLight(nr).colorfull = RGB(255-200*z,255 - 200*Z,255)
      MaterialColor "bumpertopmat" & nr, RGB(255-190*Z,235 - z*180,220 + 35*Z)

    Case "yellow"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 180 + 40*z, 48* Z), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

    Case "purple" :
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(128-118*Z - 32 * sin(Z*3.14), 32-26*Z ,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 15  + 200 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z +0.97 * Z^3)
      Flbumperbiglight(nr).intensity = 50 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (0.5 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(128-60*Z,32,255)


  end select
End Sub

Sub BumperTimer_Timer
  dim nr
  For nr = 1 to 6
    If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
      If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
    If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
      If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
  next
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------
'********************* START OF PUPDMD FRAMEWORK v1.0 *************************
Const HasPuP = True   'dont set to false as it will break pup

'//////////////////////////////////////////////////////////////////////
dim ScorbitActive
ScorbitActive         = 0   ' Is Scorbit Active
Const     ScorbitShowClaimQR  = 1   ' If Scorbit is active this will show a QR Code  on ball 1 that allows player to claim the active player from the app
Const     ScorbitUploadLog    = 0   ' Store local log and upload after the game is over
Const     ScorbitAlternateUUID  = 0   ' Force Alternate UUID from Windows Machine and saves it in VPX Users directory (C:\Visual Pinball\User\ScorbitUUID.dat)
'/////////////////////////////////////////////////////////////////////

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' DO NOT CHANGE ANYTHING IN THIS SECTION
Const     ScorbitClaimSmall   = 1   ' Make Claim QR Code smaller for high res backglass

  Const cWhite =  16777215
  Const cRed =  397512
  Const cGold =   1604786
  Const cGold2 = 46079
  Const cGreen = 32768
  Const cGrey =   8421504
  Const cYellow = 65535
  Const cOrange = 33023
  Const cPurple = 16711808
  Const cBlue = 16711680
  Const cLightBlue = 16744448
  Const cBoltYellow = 2148582
  Const cLightGreen = 9747818
  Const cBlack = 0
  Const cPink = 12615935

Const pFontBold="Bronx Bystreets"  'main score font
Const pFontGodfather = "The Scarface Free Trial"
Const dmddef="Neue Alte Grotesk Bold"

Const pDMD = 5
Const pBackglass = 2
Const pMusic = 4

'pages
Const pDMDBlank = 0
Const pScores = 1
Const pAttract = 2
Const pPrevScores = 3
Const pCredits = 4
Const pSlotMachine = 5
Const pBonus = 6
Const pEvent = 7
Const pHighScore = 8
Const pSongSelection = 15

Dim PuPlayer
dim PUPDMDObject  'for realtime mirroring.
Dim pDMDlastchk: pDMDLastchk= -1    'performance of updates
Dim pDMDCurPage: pDMDCurPage= 0     'default page is empty.
Dim pInAttract : pInAttract=false   'pAttract mode


'*************  starts PUP system
Sub PuPInit
Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
PuPlayer.B2SInit "", pGameName
  PuPlayer.LabelInit pDMD
pSetPageLayouts
pDMDSetPage pDMDBlank, "PupInit"   'set blank text overlay page.
End Sub 'end PUPINIT

' xxxxxxxxx Active Game PUP Helpers xxxxxxxxx

Sub pupCreateLabelImage(lName, lFilename,xpos, ypos, Iwidth, Iheight, pagenum, lvis)
  PuPlayer.LabelNew pBackglass,lName ,"",50,RGB(100,100,100),0,1,1,1,1,pagenum,lvis
  PuPlayer.LabelSet pBackglass,lName,lFilename,lvis,"{'mt':2,'width':"&IWidth&",'height':"&Iheight&",'xpos':"&xpos&",'ypos':"&ypos&"}"
end Sub

Sub pupCreateLabelImageDMD(lName, lFilename,xpos, ypos, Iwidth, Iheight, pagenum, lvis)
  PuPlayer.LabelNew pDMD,lName ,"",50,RGB(100,100,100),0,1,1,1,1,pagenum,lvis
  PuPlayer.LabelSet pDMD,lName,lFilename,lvis,"{'mt':2,'width':"&IWidth&",'height':"&Iheight&",'xpos':"&xpos&",'ypos':"&ypos&"}"
end Sub


Sub pDMDLabelHide(labName)
  PuPlayer.LabelSet pDMD,labName," ",0,""
end sub

Sub pDMDLabelShow(labName)
  PuPlayer.LabelSet pDMD,labName," ",1,""
end sub

Sub pDMDAlwaysPAD  'will pad all text with a space before and after to help with possible text clipping.
     PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 5, ""FN"":46, ""PA"":1 }"    '
end Sub

Sub pDMDAlwaysPADBG  'will pad all text with a space before and after to help with possible text clipping.
     PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":46, ""PA"":1 }"    'slow pc mode
end Sub

Sub PuPEvent(EventNum)
  if Not bEnablePuP then Exit Sub
  PuPlayer.B2SData "E" & EventNum, 1 'send event to puppack driver
End Sub

Dim OldPage
Sub pDMDSetPage(pagenum,caller)
  if Not bEnablePuP Then Exit sub
  pDMDCurPage = pagenum
    PuPlayer.LabelShowPage pDMD,pagenum,0,""   'set page to blank 0 page if want off
end Sub

Sub pBackglassSetPage(pagenum)
    PuPlayer.LabelShowPage pBackglass,pagenum,0,""   'set page to blank 0 page if want off
    PBackglassCurPage=pagenum
end Sub

Sub pDMDHighScore0(msgText,msgText2,timeSec,mColor)'
  PuPlayer.LabelSet pDMD,"HSMessage",msgText,1,""
  PuPlayer.LabelSet pDMD,"HSInitials",msgText2,1,""
end Sub

Sub pDMDHighScore(msgText,msgText2,timeSec,mColor)'
  PuPlayer.LabelSet pDMD,"HSMessage",msgText,1,""
  PuPlayer.LabelSet pDMD,"HSInitials",msgText2,1,""
end Sub

Sub pDMDHighScore2(msgText,msgText2,timeSec,mColor)'
  PuPlayer.LabelSet pDMD,"HSMessage2",msgText,1,""
  PuPlayer.LabelSet pDMD,"HSInitials2",msgText2,1,""
end Sub

Sub pDMDHighScore3(msgText,msgText2,timeSec,mColor)'
  PuPlayer.LabelSet pDMD,"HSMessage3",msgText,1,""
  PuPlayer.LabelSet pDMD,"HSInitials3",msgText2,1,""
end Sub

Sub pDMDHighScore4(msgText,msgText2,timeSec,mColor)'
  PuPlayer.LabelSet pDMD,"HSMessage4",msgText,1,""
  PuPlayer.LabelSet pDMD,"HSInitials4",msgText2,1,""
end Sub

Sub pDMDHighScore5(msgText,msgText2,timeSec,mColor)'
  PuPlayer.LabelSet pDMD,"HSMessage5",msgText,1,""
  PuPlayer.LabelSet pDMD,"HSInitials5",msgText2,1,""
end Sub


'==============================================================

Sub pDMDTargetLettersInfo(msgText,msgInfo, timeSec)  'msgInfo = '0211'  0= layer 1, 1=layer 2, 2=top layer3.
'this function is when you want to hilite spelled words.  Like B O N U S but have O S hilited as already hit markers... see example.
PuPlayer.LabelShowPage pDMD,5,timeSec,""  'show page 5
Dim backText
Dim middleText
Dim flashText
Dim curChar
Dim i
Dim offchars:offchars=0
Dim spaces:spaces=" "  'set this to 1 or more depends on font space width.  only works with certain fonts
                          'if using a fixed font width then set spaces to just one space.

For i=1 To Len(msgInfo)
    curChar="" & Mid(msgInfo,i,1)
    if curChar="0" Then
            backText=backText & Mid(msgText,i,1)
            middleText=middleText & spaces
            flashText=flashText & spaces
            offchars=offchars+1
    End If
    if curChar="1" Then
            backText=backText & spaces
            middleText=middleText & Mid(msgText,i,1)
            flashText=flashText & spaces
    End If
    if curChar="2" Then
            backText=backText & spaces
            middleText=middleText & spaces
            flashText=flashText & Mid(msgText,i,1)
    End If
Next

if offchars=0 Then 'all litup!... flash entire string
   backText=""
   middleText=""
   FlashText=msgText
end if

PuPlayer.LabelSet pDMD,"Back5"  ,backText  ,1,""
PuPlayer.LabelSet pDMD,"Middle5",middleText,1,""
PuPlayer.LabelSet pDMD,"Flash5" ,flashText ,0,"{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) & "}"
end Sub


Sub pDMDShowLines3(msgText,msgText2,msgText3,timeSec)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,3,timeSec,""
PuPlayer.LabelSet pDMD,"Splash3a",msgText,vis,pLine1Ani
PuPlayer.LabelSet pDMD,"Splash3b",msgText2,vis,pLine2Ani
PuPlayer.LabelSet pDMD,"Splash3c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDShowLines2(msgText,msgText2,timeSec)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,4,timeSec,""
PuPlayer.LabelSet pDMD,"Splash4a",msgText,vis,pLine1Ani
PuPlayer.LabelSet pDMD,"Splash4b",msgText2,vis,pLine2Ani
end Sub

Sub pDMDShowCounter(msgText,msgText2,msgText3,timeSec)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,6,timeSec,""
PuPlayer.LabelSet pDMD,"Splash6a",msgText,vis, pLine1Ani
PuPlayer.LabelSet pDMD,"Splash6b",msgText2,vis,pLine2Ani
PuPlayer.LabelSet pDMD,"Splash6c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDShowBig(msgText,timeSec, mColor)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,vis,pLine1Ani
end sub


Sub pDMDShowHS(msgText,msgText2,msgText3,timeSec) 'High Score
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,7,timeSec,""
PuPlayer.LabelSet pDMD,"Splash7a",msgText,vis,pLine1Ani
PuPlayer.LabelSet pDMD,"Splash7b",msgText2,vis,pLine2Ani
PuPlayer.LabelSet pDMD,"Splash7c",msgText3,vis,pLine3Ani
end Sub


Dim pNumLines

'Theme Colors for Text (not used currenlty,  use the |<colornum> in text labels for colouring.
Dim SpecialInfo
Dim pLine1Color : pLine1Color=8454143
Dim pLine2Color : pLine2Color=8454143
Dim pLine3Color :  pLine3Color=8454143
Dim curLine1Color: curLine1Color=pLine1Color  'can change later
Dim curLine2Color: curLine2Color=pLine2Color  'can change later
Dim curLine3Color: curLine3Color=pLine3Color  'can change later


Dim pDMDCurPriority: pDMDCurPriority =-1
Dim pDMDDefVolume: pDMDDefVolume = 0   'default no audio on pDMD

Dim pLine1
Dim pLine2
Dim pLine3
Dim pLine1Ani
Dim pLine2Ani
Dim pLine3Ani

Dim PriorityReset:PriorityReset=-1
DIM pAttractReset:pAttractReset=-1
DIM pAttractBetween: pAttractBetween=2000 '1 second between calls to next attract page
DIM pDMDVideoPlaying: pDMDVideoPlaying=false


'****************************************   DONT TOUCH THIS CODE   ************************************************************

Sub pupDMDDisplay(pEventID, pText, VideoName,TimeSec, pAni,pPriority)
DIM curPos
if pDMDCurPriority>pPriority then Exit Sub  'if something is being displayed that we don't want interrupted.  same level will interrupt.
pDMDCurPriority=pPriority
if timeSec=0 then timeSec=1 'don't allow page default page by accident


pLine1=""
pLine2=""
pLine3=""
pLine1Ani=""
pLine2Ani=""
pLine3Ani=""


if pAni=1 Then  'we flashy now aren't we
pLine1Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"
pLine2Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"
pLine3Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"
end If

curPos=InStr(pText,"^")   'Lets break apart the string if needed
if curPos>0 Then
   pLine1=Left(pText,curPos-1)
   pText=Right(pText,Len(pText) - curPos)

   curPos=InStr(pText,"^")   'Lets break apart the string
   if curPOS>0 Then
      pLine2=Left(pText,curPos-1)
      pText=Right(pText,Len(pText) - curPos)

      curPos=InStr("^",pText)   'Lets break apart the string
      if curPos>0 Then
         pline3=Left(pText,curPos-1)
      Else
        if pText<>"" Then pline3=pText
      End if
   Else
      if pText<>"" Then pLine2=pText
   End if
Else
  pLine1=pText  'just one line with no break
End if


'lets see how many lines to Show
pNumLines=0
if pLine1<>"" then pNumLines=pNumlines+1
if pLine2<>"" then pNumLines=pNumlines+1
if pLine3<>"" then pNumLines=pNumlines+1

if pDMDVideoPlaying Then
      PuPlayer.playstop pDMD
      pDMDVideoPlaying=False
End if


if (VideoName<>"") and (useDMDVideos) Then  'we are showing a splash video instead of the text.

    PuPlayer.playlistplayex pDMD,"DMDSplash",VideoName,pDMDDefVolume,pPriority  'should be an attract background (no text is displayed)
    pDMDVideoPlaying=true
end if 'if showing a splash video with no text


if StrComp(pEventID,"shownum",1)=0 Then              'check eventIDs
    pDMDShowCounter pLine1,pLine2,pLine3,timeSec
Elseif StrComp(pEventID,"target",1)=0 Then              'check eventIDs
    pDMDTargetLettersInfo pLine1,pLine2,timeSec
Elseif StrComp(pEventID,"highscore",1)=0 Then              'check eventIDs
    pDMDShowHS pLine1,pLine2,pline3,timeSec
Elseif (pNumLines=3) Then                'depends on # of lines which one to use.  pAni=1 will flash.
    pDMDShowLines3 pLine1,pLine2,pLine3,TimeSec
Elseif (pNumLines=2) Then
    pDMDShowLines2 pLine1,pLine2,TimeSec
Elseif (pNumLines=1) Then
    pDMDShowBig pLine1,timeSec, curLine1Color
Else
    pDMDShowBig pLine1,timeSec, curLine1Color
End if

PriorityReset=TimeSec*1000
End Sub 'pupDMDDisplay message


'******************** DO NOT MODIFY STUFF ABOVE THIS LINE!!!! ***************

Sub pSetPageLayouts
  Dim i
  pDMDAlwaysPAD

'  xxxxxxxxxxxxxxxxxxxxxxx Label layout xxxxxxxxxxxxxxxxxx
' pupCreateLabelImageDMD "[Short common]","PuPOverlays\\R_lit.png",[x pos],[y pos],[width],[height],[page number],[off/on]


' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx Rush GAME DEFINITIONS  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
  pupCreateLabelImageDMD "RLit","PuPOverlays\\R_lit.png",1.65,45,3.25,12.7,1,0
  pupCreateLabelImageDMD "ULit","PuPOverlays\\U_lit.png",5,44.2,3.25,10.1,1,0
  pupCreateLabelImageDMD "SLit","PuPOverlays\\S_lit.png",7.4,44.5,3.25,10.1,1,0
  pupCreateLabelImageDMD "HLit","PuPOverlays\\H_lit.png",10.6,44.8,3.25,12,1,0
  pupCreateLabelImageDMD "Talb","PuPOverlays\\T_album.png",4.9,18.3,12,29.5,1,0
  pupCreateLabelImageDMD "Oalb","PuPOverlays\\O_album.png",4.9,18.3,12,29.5,1,0
  pupCreateLabelImageDMD "Big","PuPOverlays\\B_mon.png",95,10,9.45,6,1,0
  pupCreateLabelImageDMD "Srad","PuPOverlays\\S_rad.png",95,10,9.45,6,1,0
  pupCreateLabelImageDMD "Lime","PuPOverlays\\L_ime.png",95,10,9.45,6,1,0
  pupCreateLabelImageDMD "tbm","PuPOverlays\\tbm1_mode.png",5.5,18,10.5,28,1,0
  pupCreateLabelImageDMD "tbml2","PuPOverlays\\tbm2_mode.png",5.5,18,10.5,28,1,0
  pupCreateLabelImageDMD "tbml3","PuPOverlays\\tbm3_mode.png",5.5,18,10.5,28,1,0
  pupCreateLabelImageDMD "sora","PuPOverlays\\S_orad.png",5.5,18,10.5,28,1,0
  pupCreateLabelImageDMD "lli","PuPOverlays\\L_lig.png",5.5,18,10.5,28,1,0
  pupCreateLabelImageDMD "hit3","PuPOverlays\\m_b3.png",95,27,9.45,6,1,0
  pupCreateLabelImageDMD "hit2","PuPOverlays\\m_b2.png",95,27,9.45,6,1,0
  pupCreateLabelImageDMD "hitr","PuPOverlays\\m_br.png",95,27,9.45,6,1,0
  pupCreateLabelImageDMD "subdiv","PuPOverlays\\sd_lit.png",50,80,70,19,1,0
  pupCreateLabelImageDMD "basslit","PuPOverlays\\bl_lit.png",1.5,78.5,5,18,1,0
  pupCreateLabelImageDMD "guitlit","PuPOverlays\\gm_lit.png",5.6,78.7,4.3,17.8,1,0
  pupCreateLabelImageDMD "drumlit","PuPOverlays\\ds_lit.png",9.3,78.5,4.5,17,1,0
  pupCreateLabelImageDMD "drumL","PuPOverlays\\dbump10k.png",26,74,23,30,1,0
  pupCreateLabelImageDMD "drumM","PuPOverlays\\dbump10k.png",50,65,23,30,1,0
  pupCreateLabelImageDMD "drumR","PuPOverlays\\dbump10k.png",75,74,23,30,1,0
  pupCreateLabelImageDMD "bm1","PuPOverlays\\big1.png",5.5,18,10.5,28,1,0
  pupCreateLabelImageDMD "lim1","PuPOverlays\\lime1.png",5.5,18,10.5,28,1,0
  pupCreateLabelImageDMD "sr1","PuPOverlays\\spirit1.png",5.5,18,10.5,28,1,0
  pupCreateLabelImageDMD "alb1","PuPOverlays\\lit1.png",1,55,4,6,1,0
  pupCreateLabelImageDMD "alb2","PuPOverlays\\lit2.png",5.8,55,4,6,1,0
  pupCreateLabelImageDMD "alb3","PuPOverlays\\lit3.png",10.6,55,4,6,1,0
  pupCreateLabelImageDMD "Como","PuPOverlays\\combo1.png",26.5,62,28,13,1,0
  pupCreateLabelImageDMD "ComTw","PuPOverlays\\combo2.png",26.5,62,28,13,1,0
  pupCreateLabelImageDMD "ComThr","PuPOverlays\\combo3.png",26.5,62,28,13,1,0
  pupCreateLabelImageDMD "Cmid","PuPOverlays\\midnight.png",94.5,41.5,7,6,1,0
  pupCreateLabelImageDMD "Cclockpm","PuPOverlays\\pm.png",94.5,41.5,7,6,1,0
  pupCreateLabelImageDMD "hs","PuPOverlays\\hscore.png",50,50,100,100,1,0
  pupCreateLabelImageDMD "yRLit","PuPOverlays\\yR_lit.png",1.65,45,3.25,12.7,1,0
  pupCreateLabelImageDMD "yULit","PuPOverlays\\yU_lit.png",5,44.2,3.25,10.1,1,0
  pupCreateLabelImageDMD "ySLit","PuPOverlays\\yS_lit.png",7.4,44.5,3.25,10.1,1,0
  pupCreateLabelImageDMD "yHLit","PuPOverlays\\yH_lit.png",10.6,44.8,3.25,12,1,0
  pupCreateLabelImageDMD "tbx","PuPOverlays\\timebox.png",15.1,45.5,5,6,1,0
  pupCreateLabelImageDMD "Aeob1","PuPOverlays\\ballbonus1.png",50,50,100,100,1,0
  pupCreateLabelImageDMD "Beob2","PuPOverlays\\ballbonus2.png",50,50,100,100,1,0
  pupCreateLabelImageDMD "Ceob3","PuPOverlays\\ballbonus3.png",50,50,100,100,1,0
  pupCreateLabelImageDMD "Deob4","PuPOverlays\\ballbonus4.png",50,50,100,100,1,0
  pupCreateLabelImageDMD "Eeob5","PuPOverlays\\ballbonus5.png",50,50,100,100,1,0

  pupCreateLabelImageDMD "c1","PuPOverlays\\credit1.png",50,50,100,100,1,0
  pupCreateLabelImageDMD "c2","PuPOverlays\\credit2.png",50,50,100,100,1,0
  pupCreateLabelImageDMD "c3","PuPOverlays\\credit3.png",50,50,100,100,1,0
  pupCreateLabelImageDMD "c4","PuPOverlays\\credit4.png",50,50,100,100,1,0

  PuPlayer.LabelNew pDMD,"CurrScore",         dmddef,6,cWhite   ,0,1,1, 12,96,pScores,0
  PuPlayer.LabelNew pDMD,"BallValue",              dmddef,4,cWhite   ,0,1,1,51,3,pScores,0
  PuPlayer.LabelNew pDMD,"Currgreentimecount",         dmddef,5.5,cWhite   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"CurrFame",         dmddef,5,cWhite   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"CurrBM",         dmddef,3.7,cWhite   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"CLevel",         dmddef,5.5,cWhite   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"CTotal",         dmddef,4.4,cYellow   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"AlbCt",         dmddef,3.5,cYellow   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"SoloT",         dmddef,5.5,cWhite   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"Clot",         dmddef,3.5,cWhite   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"Rtim",         dmddef,4,cWhite   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"Rtot",         dmddef,13,cRed   ,0,1,1, 60,16.75,pScores,0
  PuPlayer.LabelNew pDMD,"Bon",         dmddef,20,cOrange   ,0,1,1, 60,16.75,pScores,0

  PuPlayer.LabelNew pDMD, "HSMessage", dmddef, 9, cOrange, 0, 1, 1, 49, 23, pScores, 0
  PuPlayer.LabelNew pDMD, "HSInitials", dmddef, 10, cWhite, 0, 1, 1, 49, 33, pScores, 0
  PuPlayer.LabelNew pDMD, "HSMessage2", dmddef, 8, cOrange, 0, 1, 1, 29, 52, pScores, 0
  PuPlayer.LabelNew pDMD, "HSInitials2", dmddef, 9, cWhite, 0, 1, 1, 29, 60, pScores, 0
  PuPlayer.LabelNew pDMD, "HSMessage3", dmddef, 8, cOrange, 0, 1, 1, 71, 52, pScores, 0
  PuPlayer.LabelNew pDMD, "HSInitials3", dmddef, 9, cWhite, 0, 1, 1, 71, 60, pScores, 0
  PuPlayer.LabelNew pDMD, "HSMessage4", dmddef, 8, cOrange, 0, 1, 1, 29, 77, pScores, 0
  PuPlayer.LabelNew pDMD, "HSInitials4", dmddef, 9, cWhite, 0, 1, 1, 29, 85, pScores, 0
  PuPlayer.LabelNew pDMD, "HSMessage5", dmddef, 8, cOrange, 0, 1, 1, 71, 77, pScores, 0
  PuPlayer.LabelNew pDMD, "HSInitials5", dmddef, 9, cWhite, 0, 1, 1, 71, 85, pScores, 0

'  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

End Sub

Sub HideAllLabels
  pDMDLabelHide "RLit"
  pDMDLabelHide "ULit"
  pDMDLabelHide "SLit"
  pDMDLabelHide "HLit"
  pDMDLabelHide "Talb"
  pDMDLabelHide "Oalb"
  pDMDLabelHide "Big"
  pDMDLabelHide "Srad"
  pDMDLabelHide "Lime"
  pDMDLabelHide "tbm"
  pDMDLabelHide "tbml2"
  pDMDLabelHide "tbml3"
  pDMDLabelHide "sora"
  pDMDLabelHide "lli"
  pDMDLabelHide "hit3"
  pDMDLabelHide "hit2"
  pDMDLabelHide "hitr"
  pDMDLabelHide "subdiv"
  pDMDLabelHide "basslit"
  pDMDLabelHide "guitlit"
  pDMDLabelHide "drumlit"
  pDMDLabelHide "drumL"
  pDMDLabelHide "drumM"
  pDMDLabelHide "drumR"
  pDMDLabelHide "bm1"
  pDMDLabelHide "lim1"
  pDMDLabelHide "sr1"
  pDMDLabelHide "alb1"
  pDMDLabelHide "alb2"
  pDMDLabelHide "alb3"
  pDMDLabelHide "Como"
  pDMDLabelHide "ComTw"
  pDMDLabelHide "ComThr"
  pDMDLabelHide "Cmid"
  pDMDLabelHide "Cclockpm"
  pDMDLabelHide "hs"
  pDMDLabelHide "yRLit"
  pDMDLabelHide "yULit"
  pDMDLabelHide "ySLit"
  pDMDLabelHide "yHLit"
  pDMDLabelHide "tbx"
  pDMDLabelHide "CurrScore"
  pDMDLabelHide "BallValue"
  pDMDLabelHide "Currgreentimecount"
  pDMDLabelHide "CurrFame"
  pDMDLabelHide "CurrBM"
  pDMDLabelHide "CLevel"
  pDMDLabelHide "CTotal"
  pDMDLabelHide "AlbCt"
  pDMDLabelHide "SoloT"
  pDMDLabelHide "Clot"
  pDMDLabelHide "Rtim"
  pDMDLabelHide "Rtot"
End Sub

'***********************************************************PinUP Player DMD Helper Functions
Sub DMDUpdateBallNumber(nBallNr)
  if bEnablePuP Then
    PuPlayer.LabelSet pDMD,"BallValue","Ball " &nBallNr,1,""
  End If
End Sub
'                                                                                                 --------------------------------work on
Sub DMDUpdatePlayerName
  'PuPlayer.LabelSet pDMD,"CurrName","Player " & nPlayer,1,""
End Sub

Sub DMDUpdateAll
  DisplayOtherScores
  DMDUpdateBallNumber Balls
  'DisplayBonusValue
  'DisplayPlayfieldValue
End Sub

Sub DisplayOtherScores()

  'PuPlayer.LabelSet pDMD,"CurrScore",FormatScore(anScore(nPlayer)),1,""

Exit Sub

  if Not bEnablePuP or nPlayersInGame < 2 Then Exit sub

  if CurrentPlayer = 1 Then
    Select case PlayersPlayingGame
      case 1:
      case 2:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(1),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':80, 'ypos':93}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 2",1,"{'mt':2, 'color':" & cGold &", 'xpos':80, 'ypos':87}"
      case 3:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(1),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':30 , 'ypos':92.5}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 2",1,"{'mt':2, 'color':" & cGold &", 'xpos':30 , 'ypos':86.5}"
        PuPlayer.LabelSet pDMD,"Position3Score",FormatNumber(anScore(2),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':70 , 'ypos':92.5}"
        PuPlayer.LabelSet pDMD,"Position3Name","Player 3",1,"{'mt':2, 'color':" & cGold &", 'xpos':70 , 'ypos':86.5}"
      case 4:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(1),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':16,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 2",1,"{'mt':2, 'color':" & cGold &", 'xpos':16,'ypos':89}"
        PuPlayer.LabelSet pDMD,"Position3Score",FormatNumber(anScore(2),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':50,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position3Name","Player 3",1,"{'mt':2, 'color':" & cGold &", 'xpos':50,'ypos':89}"
        PuPlayer.LabelSet pDMD,"Position4Score",FormatNumber(anScore(3),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':83,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position4Name","Player 4",1,"{'mt':2, 'color':" & cGold &", 'xpos':83,'ypos':89}"
    End Select
  Elseif CurrentPlayer = 2 Then
    Select case PlayersPlayingGame
      case 1:
      case 2:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(0),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':80, 'ypos':93}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 1",1,"{'mt':2, 'color':" & cGold &", 'xpos':80, 'ypos':87}"
      case 3:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(0),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':30 , 'ypos':92.5}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 1",1,"{'mt':2, 'color':" & cGold &", 'xpos':30 , 'ypos':86.5}"
        PuPlayer.LabelSet pDMD,"Position3Score",FormatNumber(anScore(2),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':70 , 'ypos':92.5}"
        PuPlayer.LabelSet pDMD,"Position3Name","Player 3",1,"{'mt':2, 'color':" & cGold &", 'xpos':70 , 'ypos':86.5}"
      case 4:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(0),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':16,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 1",1,"{'mt':2, 'color':" & cGold &", 'xpos':16,'ypos':89}"
        PuPlayer.LabelSet pDMD,"Position3Score",FormatNumber(anScore(2),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':50,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position3Name","Player 3",1,"{'mt':2, 'color':" & cGold &", 'xpos':50,'ypos':89}"
        PuPlayer.LabelSet pDMD,"Position4Score",FormatNumber(anScore(3),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':83,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position4Name","Player 4",1,"{'mt':2, 'color':" & cGold &", 'xpos':83,'ypos':89}"
    End Select
  Elseif CurrentPlayer = 3 Then
    Select case PlayersPlayingGame
      case 1:
      case 2:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(0),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':80, 'ypos':93}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 1",1,"{'mt':2, 'color':" & cGold &", 'xpos':80, 'ypos':87}"
      case 3:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(0),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':30 , 'ypos':92.5}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 1",1,"{'mt':2, 'color':" & cGold &", 'xpos':30 , 'ypos':86.5}"
        PuPlayer.LabelSet pDMD,"Position3Score",FormatNumber(anScore(1),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':70 , 'ypos':92.5}"
        PuPlayer.LabelSet pDMD,"Position3Name","Player 2",1,"{'mt':2, 'color':" & cGold &", 'xpos':70 , 'ypos':86.5}"
      case 4:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(0),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':16,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 1",1,"{'mt':2, 'color':" & cGold &", 'xpos':16,'ypos':89}"
        PuPlayer.LabelSet pDMD,"Position3Score",FormatNumber(anScore(1),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':50,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position3Name","Player 2",1,"{'mt':2, 'color':" & cGold &", 'xpos':50,'ypos':89}"
        PuPlayer.LabelSet pDMD,"Position4Score",FormatNumber(anScore(3),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':83,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position4Name","Player 4",1,"{'mt':2, 'color':" & cGold &", 'xpos':83,'ypos':89}"
    End Select
  Elseif CurrentPlayer = 4 Then
    Select case PlayersPlayingGame
      case 1:
      case 2:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(0),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':80, 'ypos':93}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 1",1,"{'mt':2, 'color':" & cGold &", 'xpos':80, 'ypos':87}"
      case 3:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(0),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':30 , 'ypos':92.5}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 1",1,"{'mt':2, 'color':" & cGold &", 'xpos':30 , 'ypos':86.5,'ypos':89}"
        PuPlayer.LabelSet pDMD,"Position3Score",FormatNumber(anScore(1),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':70 , 'ypos':92.5}"
        PuPlayer.LabelSet pDMD,"Position3Name","Player 2",1,"{'mt':2, 'color':" & cGold &", 'xpos':70 , 'ypos':86.5,'ypos':89}"
      case 4:
        PuPlayer.LabelSet pDMD,"Position2Score",FormatNumber(anScore(0),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':16 ,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position2Name","Player 1",1,"{'mt':2, 'color':" & cGold &", 'xpos':16 ,'ypos':89}"
        PuPlayer.LabelSet pDMD,"Position3Score",FormatNumber(anScore(1),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':50,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position3Name","Player 2",1,"{'mt':2, 'color':" & cGold &", 'xpos':50,'ypos':89}"
        PuPlayer.LabelSet pDMD,"Position4Score",FormatNumber(anScore(2),0),1,"{'mt':2, 'color':" & cGold &", 'xpos':83,'ypos':94}"
        PuPlayer.LabelSet pDMD,"Position4Name","Player 3",1,"{'mt':2, 'color':" & cGold &", 'xpos':83,'ypos':89}"
    End Select
  End If
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub ExitSongSelection
  bSongSelection = False
  SwitchMusic ArraySongs(SongNR)
End Sub

Sub UpdateSongLeft
  Dim TmpSong
  SongNR = SongNR - 1
  if SongNR < 1 Then SongNR = nMaxSongs
  TmpSong = SongNR + 400
  PupEvent TmpSong
End Sub

Sub UpdateSongRight
  Dim TmpSong
  SongNR = SongNR + 1
  if SongNR > nMaxSongs Then SongNR = 1
  TmpSong = SongNR + 400
  PupEvent TmpSong
End Sub

'*******************************************
'Music
'*******************************************
Dim fCurrentMusicVol : fCurrentMusicVol = 0
Dim sMusicTrack : sMusicTrack = ""
Dim fSongVolume : fSongVolume = 1
Dim NewSong
Dim ArraySongs
ArraySongs = array("","Song1-WorkingMan","Song2-FlybyNight","song3-red","song4-sub","song5-bastille","song6-farcry")


'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
Sub SwitchMusic(sTrack)
  If sTrack <> sMusicTrack Then
    StopSound sMusicTrack
    sMusicTrack = sTrack
    PlaySound sTrack, -1, fMusicVolume,0,0,0,1,0,0
    fCurrentMusicVol = fMusicVolume
  End If
End Sub


Sub UpdateMusicNow
  StopAllMusic
    Select Case newSong
        Case 1:SwitchMusic "Song1-WorkingMan"
        Case 2:SwitchMusic "Song2-FlybyNight"
        Case 3:SwitchMusic "song3-red"
        Case 4:SwitchMusic "song4-sub"
        Case 5:SwitchMusic "song5-bastille"
        Case 6:SwitchMusic "song6-farcry"
    End Select
end sub


Sub StopAllMusic
  sMusicTrack = ""
  StopSound "Song1-WorkingMan"
  StopSound "Song2-FlybyNight"
  StopSound "song3-red"
  StopSound "song4-sub"
  StopSound "song5-bastille"
  StopSound "song6-farcry"
End Sub

Function Balls
    Dim tmp
    tmp = (BallsPerGame - BallsRemaining(CurrentPlayer)) + 1
    If tmp> BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function


Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function



'************** Nailbuster TriggerScript Code v1.20
' create a timer in table named exactly pTriggerScript
' set interval on timer to 100ms (not enabled on startup)
'   simpletest:  TriggerScript 3500,"MsgBox 1234"   'will show a dialog 1234


Const TriggerScriptSize=10
Dim pReset(10)                 ' TriggerScriptSize
Dim pStatement(10)             ' TriggerScriptSize - holds future scripts
Dim FX

Sub TriggersStopAll()
for fx=0 to TriggerScriptSize
    pReset(FX)=0
    pStatement(FX)=""
next
pTriggerScript.Enabled=False  'YOU MUST HAVE A TIMER NAMED pTriggerScript interval 100 not active on startup.
end Sub

TriggersStopAll



DIM pTriggerCounter:pTriggerCounter=pTriggerScript.interval    'YOU MUST HAVE A TIMER NAMED pTriggerScript interval 100 not active on startup.

Sub pTriggerScript_Timer()
    dim bMoreToRun:bMoreToRun=False
    for fx=0 to TriggerScriptSize
        if pReset(fx)>0 Then
            pReset(fx)=pReset(fx)-pTriggerCounter
            if pReset(fx)<=0 Then
                pReset(fx)=0
                execute(pStatement(fx))
            end if
            bMoreToRun=True
        End if
    next
    if bMoreToRun = False then pTriggerScript.Enabled=False    ' Disable when we dont need it
End Sub


Sub TriggerScript(pTimeMS, pScript) ' This is used to Trigger script after the pTriggerScript Timer
    for fx=0 to TriggerScriptSize
        if pReset(fx)=0 Then
            pReset(fx)=pTimeMS
            pStatement(fx)=pScript
            if pTriggerScript.enabled=false Then pTriggerScript.enabled=true
            Exit Sub
        End If
    next
end Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


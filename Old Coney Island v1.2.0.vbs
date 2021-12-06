'*****************************
'      Old Coney Isand
'        para JOLO
' (vale también para Leo :) )
'*****************************
' Updated DOF commands by arngrim
Option Explicit
Randomize


' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "00990300", "GamePlan.vbs", 3.1

Dim cCredits
cCredits = "Coney Island, GamePlan, 1979"
Const cGameName = "coneyis"

Const UseSolenoids = 1
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 1

' Standard vpinmame Sounds
Const SCoin = "fx_coin"

' Solenoids

SolCallback(8) = "bsTrough.SolOut"
SolCallback(11) = "dtDrop.SolDropUp"
SolCallback(15) = "bsSaucer.SolOut"
SolCallback(16) = "vpmNudge.SolGameOn"

' If you want chimes uncomment these lines:
'SolCallback(18) = "vpmSolSound ""fx_Chime4"","
'SolCallback(19) = "vpmSolSound ""fx_Chime3"","
'SolCallback(20) = "vpmSolSound ""fx_Chime2"","
'SolCallback(21) = "vpmSolSound ""fx_Chime1"","

' not used in the script
'SolCallback(12) bumper 2
'SolCallback(13) bumper 3
'SolCallback(14) bumper 1
'SolCallback(9)  left slingshot
'SolCallback(10) right slingshot

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub Rightflipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

'**************
'  Table Init
'**************

Dim bsTrough, dtDrop, bsSaucer

Sub Table1_Init
    On Error Resume Next
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = cCredits
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .HandleMechanics = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'"                              'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

  ' Press F6 during the game to change the dip switches
  ' destruk dip switches - awards extra ball
    'Controller.Dip(0) = (0 * 1 + 1 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 0 * 64 + 0 * 128) '01-08
    'Controller.Dip(1) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 0 * 64 + 1 * 128) '09-16
    'Controller.Dip(2) = (0 * 1 + 1 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 0 * 64 + 0 * 128) '17-24
    'Controller.Dip(3) = (1 * 1 + 1 * 2 + 1 * 4 + 0 * 8 + 0 * 16 + 1 * 32 + 1 * 64 + 0 * 128) '25-32

  ' Jolo dip switches - awards extra game
    'Controller.Dip(0) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 0 * 64 + 0 * 128) '01-08
    'Controller.Dip(1) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 1 * 16 + 0 * 32 + 0 * 64 + 1 * 128) '09-16
    'Controller.Dip(2) = (0 * 1 + 1 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 0 * 64 + 0 * 128) '17-24
    'Controller.Dip(3) = (1 * 1 + 1 * 2 + 1 * 4 + 0 * 8 + 1 * 16 + 1 * 32 + 1 * 64 + 1 * 128) '25-32

    ' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 11, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Left Eject Hole
    Set bsSaucer = New cvpmBallStack
    With bsSaucer
        .InitSaucer sw24, 24, 136, 28
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_kicker", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 1
    End With

    ' Drop targets
    set dtDrop = new cvpmdroptarget
    With dtDrop
        .InitDrop Array(sw31, sw32, sw35, sw36, sw4, sw10, sw17), Array(31, 32, 35, 36, 4, 10, 17)
        .initsnd "", SoundFX("fx_resetdrop", DOFContactors)
    '.CreateEvents "dtDrop" 'done manually in the script because of the 3d mesh droptargets animation.
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Map lights into array
    vpmMapLights aLights

  ' Remove the cabinet rails if in FS mode
  If Table1.ShowDT = False then
    lrail.Visible = False
    rrail.Visible = False
  End If
End Sub

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull",Plunger,1:Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger",Plunger,1:Plunger.Fire
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), Lemk, 1
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 15
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), Remk, 1
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 34
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select

    RStep = RStep + 1
End Sub

'Switches, targets, triggers

Sub Drain_Hit:PlaysoundAtVol "fx_drain", drain, 1:bsTrough.AddBall Me:End Sub
Sub RubberBand5_Hit:vpmTimer.PulseSw 9:End Sub
Sub RubberBand6_Hit:vpmTimer.PulseSw 9:End Sub
Sub RubberBand7_Hit:vpmTimer.PulseSw 9:End Sub
Sub RubberBand8_Hit:vpmTimer.PulseSw 9:End Sub
Sub RubberBand9_Hit:vpmTimer.PulseSw 9:End Sub
Sub RubberBand10_Hit:vpmTimer.PulseSw 9:End Sub
Sub RubberBand14_Hit:vpmTimer.PulseSw 9:End Sub
Sub sw12_Hit:Controller.Switch(12) = 1:End Sub
Sub sw12_unHit:Controller.Switch(12) = 0:End Sub
Sub sw13_Hit:Controller.Switch(13) = 1:End Sub
Sub sw13_unHit:Controller.Switch(13) = 0:End Sub
Sub sw14_Hit:Controller.Switch(14) = 1:End Sub
Sub sw14_unHit:Controller.Switch(14) = 0:End Sub
Sub sw16_Hit:Controller.Switch(16) = 1:DOF 102, DOFOn:End Sub
Sub sw16_unHit:Controller.Switch(16) = 0:DOF 102, DOFOff:End Sub
Sub sw16a_Hit:Controller.Switch(16) = 1:DOF 101, DOFOn:End Sub
Sub sw16a_unHit:Controller.Switch(16) = 0:DOF 101, DOFOff:End Sub
Sub sw18_Hit:Controller.Switch(18) = 1:End Sub
Sub sw18_unHit:Controller.Switch(18) = 0:End Sub
Sub Target1_Hit:vpmTimer.PulseSw(19):PlaySoundAtVol SoundFX("fx_target", DOFContactors),ActiveBall,VolTarg:End Sub
Sub Target2_Hit:vpmTimer.PulseSw(20):PlaySoundAtVol SoundFX("fx_target", DOFContactors),ActiveBall,VolTarg:End Sub
Sub Bumper1_Hit:vpmTimer.PulseSw(21):PlaySoundAtVol SoundFX("fx_bumper", DOFContactors),Bumper1,VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw(22):PlaySoundAtVol SoundFX("fx_bumper", DOFContactors),Bumper2,VolBump:End Sub
Sub Spinner1_Spin:vpmTimer.PulseSw(23):End Sub
Sub sw24_Hit:PlaySoundAtVol "fx_kicker_enter",sw24,VolKick:bsSaucer.AddBall 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:End Sub
Sub sw25_unHit:Controller.Switch(25) = 0:End Sub
Sub sw27_Hit:Controller.Switch(27) = 1:End Sub
Sub sw27_unHit:Controller.Switch(27) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:End Sub
Sub sw28_unHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:End Sub
Sub sw29_unHit:Controller.Switch(29) = 0:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1:End Sub
Sub sw30_unHit:Controller.Switch(30) = 0:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw(33):PlaySoundAtVol SoundFX("fx_bumper", DOFContactors),Bumper3,VolBump:End Sub
Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_unHit:Controller.Switch(37) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:End Sub
Sub sw38_unHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:End Sub
Sub sw39_unHit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:End Sub
Sub sw40_unHit:Controller.Switch(40) = 0:End Sub

'droptargets
Sub sw31_Dropped():dtDrop.Hit 1:End Sub
Sub sw32_Dropped():dtDrop.Hit 2:End Sub
Sub sw35_Dropped():dtDrop.Hit 3:End Sub
Sub sw36_Dropped():dtDrop.Hit 4:End Sub
Sub sw4_Dropped():dtDrop.Hit 5:End Sub
Sub sw10_Dropped():dtDrop.Hit 6:End Sub
Sub sw17_Dropped():dtDrop.Hit 7:End Sub

Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_UnPaused:Controller.Pause = False:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
    'GIUpdate
End Sub

'General Illumination

Set LampCallback = GetRef("GIUpdate")

Sub GiON
    Dim x
    For each x in aGiLights
        x.State = 1
    Next
  l53b.State = 2
End Sub

Sub GiOFF
    Dim x
    For each x in aGiLights
        x.State = 0
    Next
  l53b.State = 0
End Sub

Dim OldGiState
OldGiState = 0 'start witht he Gi off

Sub GIUpdate
    If Controller.Lamp(53) <> OldGiState Then
        OldGiState = Controller.Lamp(53)
        If Controller.Lamp(53) = 0 Then
            GiOff
        Else
            GiOn
        End If
    End If
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_metalhit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub adroptargets_Hit(idx):PlaySound SoundFX("fx_droptarget", DOFContactors), 0, VolTarg, pan(ActiveBall), 0.15, AudioFade(ActiveBall):End Sub
Sub alanes_Hit(idx):PlaySound SoundFX("fx_sensor", DOFContactors), 0, 1, pan(ActiveBall), 0.15, AudioFade(ActiveBall):End Sub

'Gameplan
'added by Inkochnito
Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,400,"Gameplan - DIP switches"
    .AddFrame 2,5,190,"Maximum credits",&H07000000,Array("5 credits",0,"10 credits",&H01000000,"15 credits",&H02000000,"20 credits",&H03000000,"25 credits",&H04000000,"30 credits",&H05000000,"35 credits",&H06000000,"40 credits",&H07000000)'dip 25&26&27
    .AddFrame 210,97,190,"High game to date award",&HC0000000,Array("no award",0,"1 credit",&H40000000,"2 credits",&H80000000,"3 credits",&HC0000000)'dip 31&32
    .AddFrame 210,5,190,"Special award",&H10000000,Array("extra ball",0,"replay",&H10000000)'dip 29
    .AddFrame 210,51,190,"Balls per game",&H08000000,Array("3 balls",0,"5 balls",&H08000000)'dip 28
    .AddChk 2,140,150,Array("Play tunes",32768)'dip 16
    .AddChk 2,155,150,Array("Match feature",&H20000000)'dip 30
    .AddChk 2,170,150,Array("Free play",&H00000080)'dip 8
    .AddLabel 30,200,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")

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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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

Const tnob = 1 ' total number of balls in this table is 4, but always use a higher number here because of the timing
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


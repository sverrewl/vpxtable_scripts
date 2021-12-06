' Harlem Globettroters on Tour - Bally 1979
' VPX version by JPSalas 2009, version 1.0
' Uses 7 digits ROM bootleg
' You need both roms: hglbtrtr.zip and hglbtrtb.zip
' Script based on Gaston's script
' Dedicado a Jolo :)
' DOF extension ny arngrim

Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed useSolenois=1 to 2, thanks for reporting Brer Frog
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
Const VolWires  = 1    ' Wires volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "Bally.vbs", 3.26

Dim bsTrough, bsMSaucer, bsRSaucer, dtDrop, x, plungerIM, uHole
Dim bump1, bump2, bump3

'Const cGameName = "hglbtrtr" ' normal 6 digits rom
Const cGameName = "hglbtrtb" ' bootleg 7 digits rom

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    vpmMapLights AllLamps ' Map all lamps into lights array

    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Harlem Globettroters on Tour - Bally 1979" & vbNewLine & "VPX table by JPSalas v.1.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot, Sling1, Sling2)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 8, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .Balls = 1
    End With

    ' Drop targets
    set dtDrop = new cvpmdroptarget
    With dtDrop
        .initdrop array(d1, d2, d3, d4), array(1, 2, 3, 4)
        .initsnd SoundFX("fx_droptarget",DOFContactors), SoundFX("fx_resetdrop",DOFContactors)
    End With

    ' Middle Saucer
    Set bsMSaucer = New cvpmBallStack
    With bsMSaucer
        .InitSaucer sw24, 24, 192, 12
        .KickAngleVar = 2
        .KickForceVar = 1
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_kicker",DOFContactors)
    End With

    ' Right Saucer
    Set bsRSaucer = New cvpmBallStack
    With bsRSaucer
        .InitSaucer sw32, 32, 302, 20
        .KickAngleVar = 2
        .KickForceVar = 1
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_kicker",DOFContactors)
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

  ' Remove the cabinet rails if in FS mode
  If Table1.ShowDT = False then
    lrail.Visible = False
    rrail.Visible = False
    ramp1.Visible = False
    ramp2.Visible = False
  End If
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull", plunger, 1:Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger", plunger, 1:Plunger.Fire
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********
' Switches
'*********

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), lemk, 1
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 37
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
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), remk, 1
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 36
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

Sub Sling1_Hit:PlaySoundAtVol "fx_Rubber", ActiveBall, 1:vpmTimer.PulseSw 34:End Sub
Sub Sling2_Hit:PlaySoundAtVol "fx_Rubber", ActiveBall, 1:vpmTimer.PulseSw 34:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 40:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),Bumper1, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 38:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),Bumper2, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 39:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),Bumper3, VolBump:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAtVol "fx_drain", drain, 1:bsTrough.AddBall Me:End Sub
Sub sw24_Hit:PlaySoundAtVol "fx_kicker_enter",sw24,VolKick:bsMSaucer.AddBall 0:End Sub
Sub sw32_Hit:PlaySoundAtVol "fx_kicker_enter",sw32,VolKick:bsRSaucer.AddBall 0:End Sub

' Rollovers
Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw23_UnHit::Controller.Switch(23) = 0:End Sub

'Spinners
Sub sw17_Spin():vpmTimer.PulseSw 17:PlaySoundAtVol "fx_spinner",sw17,VolSpin:End Sub
Sub sw33_Spin():vpmTimer.PulseSw 33:PlaySoundAtVol "fx_spinner",sw32,VolSpin:End Sub
Sub sw25_Spin():vpmTimer.PulseSw 25:PlaySoundAtVol "fx_spinner",sw25,VolSpin:End Sub

' Droptargets
Sub d1_Hit:dtDrop.Hit 1:PlaySoundAtVol SoundFX("fx_droptarget",DOFContactors),ActiveBall, VolTarg:End Sub
Sub d2_Hit:dtDrop.Hit 2:PlaySoundAtVol SoundFX("fx_droptarget",DOFContactors),ActiveBall, VolTarg:End Sub
Sub d3_Hit:dtDrop.Hit 3:PlaySoundAtVol SoundFX("fx_droptarget",DOFContactors),ActiveBall, VolTarg:End Sub
Sub d4_Hit:dtDrop.Hit 4:PlaySoundAtVol SoundFX("fx_droptarget",DOFContactors),ActiveBall, VolTarg:End Sub

' Targets
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, VolTarg:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, VolTarg:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, VolTarg:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, VolTarg:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, VolTarg:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, VolTarg:End Sub

'*********
'Solenoids
'*********

SolCallback(7) = "bsTrough.SolOut"
SolCallback(6) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallBack(13) = "bsMSaucer.SolOut"
SolCallBack(14) = "bsRSaucer.SolOut"
SolCallback(15) = "dtDrop.SolDropUp"
SolCallback(19) = "vpmNudge.SolGameOn"
SolCallback(17) = "Soldiv"

Sub Soldiv(Enabled)
    vpmSolDiverter RightLaneGate, True, Not Enabled
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), LeftFlipper, VolFlip
        PlaySoundAtVol "fx_flipperup", LeftFlipper1, VolFlip
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), LeftFlipper, VolFlip
        PlaySoundAtVol "fx_flipperdown", LeftFlipper1, VolFlip
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
    Diverter.RotZ = RightLaneGate.CurrentAngle
  GIUpdate
End Sub

Sub GiON
  For each x in aGiLights
    x.State = 1
  Next
End Sub

Sub GiOFF
  For each x in aGiLights
    x.State = 0
  Next
End Sub

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            GiOff
        Else
            GiOn
        End If
    End If
End Sub
'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetal_Wires_Hit(idx):PlaySound "fx_metalhit", 0, Vol(ActiveBall)*VolWires, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RightLaneGate_Collide(parm)
    PlaySoundAtVol "fx_metalhit", RightLaneGate, VolMetal
End Sub

'*************************************
'Bally Harlem Globetrotters 7 digits
'added by Inkochnito
'*************************************
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Harlem GlobeTrotters 7 digits - DIP switches"
        .AddFrame 2, 0, 190, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "free play (40 credits)", &H03000000)                    'dip 25&26
        .AddFrame 2, 87, 190, "Sound features", &H30000000, Array("chime effects", 0, "noises and no background", &H10000000, "noise effects", &H20000000, "noises and background", &H30000000) 'dip 29&30
        .AddFrame 2, 170, 190, "Special adjustment", &H00000060, Array("points", 0, "extra ball", &H00000040, "replay/extra ball", &H00000060)                                                  'dip 6&7
        .AddFrame 2, 230, 190, "5 side target lite adjustment", &H00800000, Array("targets will reset", 0, "targets are held in memory", &H00800000)                                            'dip 24
        .AddFrame 2, 276, 190, "G-L-O-B-E saucer scanning adjustment", &H00002000, Array("Globe lites do not scan", 0, "Globe lites keep scanning", &H00002000)                                 'dip 14
        .AddFrame 2, 322, 190, "Dunk shot target special", &H00400000, Array("is reset after collecting", 0, "stays lit", &H00400000)                                                           'dip 23
        .AddFrame 205, 0, 190, "High game to date", &H00200000, Array("no award", 0, "3 credits", &H00200000)                                                                                   'dip 22
        .AddFrame 205, 46, 190, "Score version", &H00100000, Array("6 digit scoring", 0, "7 digit scoring", &H00100000)                                                                         'dip 21
        .AddFrame 205, 92, 190, "Balls per game", &H40000000, Array("3 balls", 0, "5 balls", &H40000000)                                                                                        'dip 31
        .AddFrame 205, 138, 190, "Saucer targets reset", &H00000080, Array("when ball enters target saucer", 0, "on next ball in play", &H00000080)                                             'dip 8
        .AddFrame 205, 184, 190, "Super bonus", &H00004000, Array("is reset every ball", 0, "is held in memory", &H00004000)                                                                    'dip 15
        .AddFrame 205, 230, 190, "Globe special lite adjustment", &H80000000, Array("left outlane 25K only lit", 0, "left outlane 25K and Globe special lit", &H80000000)                       'dip 32
        .AddFrame 205, 276, 190, "Left and right spinner adjust", 32768, Array("left spinner only which alternates", 0, "left and right spinner lite on", 32768)                                'dip 16
        .AddChk 205, 330, 180, Array("Match feature", &H08000000)                                                                                                                               'dip 28
        .AddChk 205, 350, 115, Array("Credits displayed", &H04000000)                                                                                                                           'dip 27
        .AddLabel 50, 370, 350, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

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
Sub table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


' Flash 1979 - Williams / 4 Players
' VPX by JPSalas May 2016
' Based on the script by desktruk & gaston
' VPX 10.4 Surround Sound MOD by RustyCardores
' Additional lighting and image updates by Hauntfreaks
' Updated with permission October 2017
' Thank you JPSalas for your continued support of the VP Community
' VPX 10.4 is required for this table

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-12-18 : Added FFv2
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01120100", "s4.vbs", 3.02

' Thalamus - for Fast Flip v2
' NoUpperRightFlipper
NoUpperLeftFlipper

Dim bsTrough, dtRBank, dtLBank, bsLHole, x, plungerIM

Const cGameName = "flash_l1"

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = "fx_FlipperUp"
Const SFlipperOff = "fx_FlipperDown"
Const SCoin = "fx_coin"

'************
' Table init.
'************

Sub Table1_Init
    vpmInit me
    vpmMapLights AllLamps ' Map all lamps into lights array

    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Flash, Williams 1979" & vbNewLine & "VPX table by JPSalas - Surround Sound MOD by RustyCardores"
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = 0
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

    ' Nudging
    vpmNudge.TiltSwitch = 47
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, leftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 48, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Left 5 Drop targets
    set dtLBank = new cvpmdroptarget
    With dtLBank
        .InitDrop Array(sw36, sw35, sw34, sw33, sw32), Array(36, 35, 34, 33, 32)
        .initsnd SoundFX("fx_droptargetL", DOFContactors), SoundFX("fx_resetdropL", DOFContactors)
        .AllDownSw = 37
        .CreateEvents "dtLBank"
    End With

    ' Right 3 Drop targets
    set dtRBank = new cvpmdroptarget
    With dtRBank
        .InitDrop Array(sw30, sw29, sw28), Array(30, 29, 28)
        .initsnd SoundFX("fx_droptargetC", DOFContactors), SoundFX("fx_resetdropC", DOFContactors)
        .AllDownSw = 31
        .CreateEvents "dtRBank"
    End With

    ' Right Eject Hole
    Set bsLHole = New cvpmBallStack
    With bsLHole
        .InitSaucer sw27, 27, 190, 8
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_kicker", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 1
    End With



    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull",Plunger:Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = KeyRules Then Rules
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger",Plunger:Plunger.Fire
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********
' Switches
'*********

' Slings & Rubbers
' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors),Lemk,2
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 41
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
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors),Remk,2
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 42
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

Sub sw24_Hit:vpmTimer.PulseSw 24:PlaySoundAtBall "fx_rubber":End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySoundAtBall "fx_rubber":End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySoundAtBall "fx_rubber":End Sub
Sub sw16_Hit:vpmTimer.PulseSw 16:PlaySoundAtBall"fx_rubber":End Sub
Sub sw11_Hit:vpmTimer.PulseSw 11:PlaySoundAtBall "fx_rubber":End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySoundAtBall "fx_rubber":End Sub
Sub sw38a_Hit:vpmTimer.PulseSw 38:PlaySoundAtBall "fx_rubber":End Sub


' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 18:PlaySoundAtBumperVol SoundFX("fx_bumper1", DOFContactors), Bumper1,2:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 19:PlaySoundAtBumperVol SoundFX("fx_bumper2", DOFContactors), Bumper2,2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 20:PlaySoundAtBumperVol SoundFX("fx_bumper3", DOFContactors), Bumper3,2:End Sub


' Spinners

Sub sw17_Spin:vpmTimer.PulseSw 17:PlaySoundAt "fx_spinner",sw17:End Sub

' Drain & holes
Sub Drain_Hit:PlaySoundAtVol "fx_drain",Drain,.8:bsTrough.AddBall Me:End Sub
Sub sw27_Hit:PlaySoundAt "fx_kicker_enter",sw27:bsLHole.AddBall Me:End Sub

' Rollovers
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub

' Targets
Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAt SoundFX("fx_target", DOFContactors),ActiveBall:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAt SoundFX("fx_target", DOFContactors),ActiveBall:End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "SolReset 1,"
SolCallback(3) = "SolReset 2,"
SolCallback(4) = "dtRBank.SolDropUp"
SolCallback(5) = "bsLHole.SolOut"
solcallback(6) = "SolFlash"
solcallback(14) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
solcallback(23) = "SolRun"

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors),LeftFlipper,2
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors),LeftFlipper
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors),RightFlipper,2
        RightFlipper.RotateToEnd
        RightFlipper2.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors),RightFlipper:PlaySoundAt "fx_flipperdown2",RightFlipper2
        RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper",2
End Sub

Sub Rightflipper2_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper",2
End Sub

Sub Rightflipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper",2
End Sub

'Solenoid subs

Sub SolReset(No, Enabled)
    If Enabled Then
        Controller.Switch(37) = False
        If no < 2 Then
            dtLbank.SolUnHit 1, True
            dtLbank.SolUnHit 2, True
            dtLbank.SolUnHit 3, True
        Else
            dtLbank.SolUnHit 4, True
            dtLbank.SolUnHit 5, True
        End If
    End If
End Sub

Sub SolRun(Enabled)
    vpmNudge.SolGameOn Enabled
End Sub

Sub SolFlash(enabled)
    If enabled Then
        f1.state = 1
        f1a.visible = 1
        f1b.visible = 1
    Else
        f1.State = 0
        f1a.visible = 0
        f1b.visible = 0
    End if
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1700)
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

function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
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
  If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
      Else
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
      End If
    Else
      If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
      End If
    End If
  Next
End Sub

'***********Ball Shadow Update
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
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

Sub aMetals_Hit(idx):PlaySound "fx_metalhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*1.8, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*1.8, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*1.8, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, 2, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0,AudioFade(ActiveBall):End Sub

'******
' Rules
'******
Dim Msg(17)
Sub Rules()
    Msg(0) = "Flash - by Williams" &Chr(10) &Chr(10)
    Msg(1) = ""
    Msg(2) = "Insert coin and wait for machine to reset before inserting coin for"
    Msg(3) = " next player."
    Msg(4) = ""
    Msg(5) = "Making 1 - 2 - 3  lights 2x. Making 1 - 2 - 3 - 4  lights 3x."
    Msg(6) = "Making 3 bank drop targets advances thru Thunder, Lightning, Tempest"
    Msg(7) = "and Super Flash."
    Msg(8) = ""
    Msg(9) = "Making 5 bank targets 1st time advances Hole Kicker value, 2nd"
    Msg(10) = "time lights Extra Ball, 3rd time lights out lane specials."
    Msg(11) = ""
    Msg(12) = "Tilt penalty - Ball in Play - does not disqualify player."
    Msg(13) = "Special scores 1 Credit."
    Msg(14) = "Beating highest score scores 3 credits."
    Msg(15) = "Matching last two numbers on score with numbers in match window"
    Msg(16) = "on backglass scores 1 credit"
    Msg(17) = ""

    For X = 1 To 17
        Msg(0) = Msg(0) + Msg(X) &Chr(13)
    Next
    MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub



'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
        PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

'**************************************************************************
'                 Positional Sound Playback Functions by DJRobX
' VPX version check only required for 10.3 backwards compatibility.
' Version check and Else statement may be removed if table is > 10.4 only.
'**************************************************************************

'Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(sound, tableobj)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
  Else
    PlaySound sound, 1, 1, Pan(tableobj)
  End If
End Sub


'Set all as per ball position & speed.

Sub PlaySoundAtBall(sound)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
  Else
    PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
  End If
End Sub


'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
  Else
    PlaySound sound, 1, Vol, Pan(tableobj)
  End If
End Sub


'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
  Else
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
  End If
End Sub


'Set position as bumper and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
  Else
    PlaySound sound, 1, Vol, Pan(tableobj)
  End If
End Sub



' Notes: To be left in script so others can learn & understand new VPX Surround Code
'
' PlaySoundAtBall "sound",ActiveBall
'   * Sets position as ball and Vol to 1

' PlaySoundAtBallVol "sound",x
'   * Same as PlaySounAtBall but sets x as a volume multiplier (1-10) or partial multiplier (.01-.99)
'
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
'   * May us used as shown, or with any manual setting, in place of any above Sound Playback Function.
'
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
'   * May us used as shown, or with any manual setting, to maintain 10.3 backwards compatability.

'**********************************************************************


Dim NextOrbitHit:NextOrbitHit = 0

Sub MetalWallBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 3, 20000
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .1 + (Rnd * .2)
  end if
End Sub

' Requires fx_wallbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
  dim BumpSnd:BumpSnd= "fx_wallbump" & CStr(Int(Rnd*7)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


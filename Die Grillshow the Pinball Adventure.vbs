' Die Grillshow the Pinball Adventure
' By Bambi Plattfuss
' thx to
' - Trochjochel
' - JPJ and team PP
' - Arngrim
' - my wife and kid
' - vpin-shop.com
' - die Grillshow
' - everyone i forgot

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2020 February : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Dim DesktopMode:DesktopMode = Table1.ShowDT

 Dim bsTrough, dtRBank, dtTBank, x

 Const cGameName = "eldorado"

 Const UseSolenoids = 1
 Const UseLamps = 0
 Const UseGI = 0
 Const UseSync = 0
 Const HandleMech = 0

 ' Standard Sounds
 Const SSolenoidOn = "Solenoid"
 Const SSolenoidOff = ""
 Const SFlipperOn = "FlipperUp"
 Const SFlipperOff = "FlipperDown"
 Const SCoin = "coin"

LoadVPM "01120100", "sys80.vbs", 3.02

 '************
 ' Table init.
 '************

 Sub Table1_Init
  Dim ii
    With Controller
       .GameName = cGameName
       .Games(cGameName).Settings.Value("sound") = 0 'turn off rom sound
       .SplashInfoLine = "El Bueno, el Feo y el Malo" & vbNewLine & "basado en varias tablas de Gottlieb"
       .HandleMechanics = 0
       .HandleKeyboard = 0
       .ShowDMDOnly = 1
       .ShowFrame = 0
       .ShowTitle = 0
       .Hidden = DesktopMode
       If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

    ' Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingshot)

  Set bsTrough = New cvpmTrough
  With bsTrough
  .size = 1
  '.entrySw = 18
  .initSwitches Array(67)
  .Initexit BallRelease, 80, 6
  .InitExitSounds SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_ballrel",DOFContactors)
  .Balls = 1
  End With

    ' Top Drop targets
    set dtTBank = new cvpmdroptarget
    With dtTBank
       .InitDrop Array(sw0, sw10, sw20, sw30, sw40, sw1, sw11, sw21, sw31, sw41), Array(0, 10, 20, 30, 40, 1, 11, 21, 31, 41)
       .Initsnd SoundFX("fx_droptarget",DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    End With

    ' Right Drop targets
    set dtRBank = new cvpmdroptarget
    With dtRBank
       .InitDrop Array(sw2, sw12, sw22, sw32, sw42), Array(2, 12, 22, 32, 42)
       .Initsnd SoundFX("fx_droptarget",DOFDropTargets), SoundFX("fx_resetdrop",DOFContactors)
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Start the music from Fast Draw table
    MusicOn

  ' Remove the cabinet rails if in FS mode
  If Table1.ShowDT = False then
    lrail.Visible = False
    rrail.Visible = False
    ramp4.Visible = False
    ramp5.Visible = False
    For each ii in aReels
      ii.Visible = 0
    Next
  End If
 End Sub

 Sub Table1_Paused:Controller.Pause = 1:End Sub
 Sub Table1_unPaused:Controller.Pause = 0:End Sub

 Sub MusicOn
    Dim x
    x = INT(5 * RND(1) )
    Select Case x
       Case 0:PlayMusic "GS_1.mp3"
       Case 1:PlayMusic "GS_2.mp3"
       Case 2:PlayMusic "GS_3.mp3"
       Case 3:PlayMusic "GS_4.mp3"
       Case 4:PlayMusic "GS_5.mp3"
    End Select
 End Sub

 Sub Table1_MusicDone()
    MusicOn
 End Sub

 '**********
 ' Keys
 '**********

 Sub table1_KeyDown(ByVal Keycode)
    If vpmKeyDown(keycode) Then Exit Sub
    If keyCode = LeftFlipperKey Then Controller.Switch(6) = 1
    If keyCode = RightFlipperKey Then Controller.Switch(16) = 1
    If keycode = PlungerKey Then Plunger.Pullback
    If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
 End Sub

 Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keyCode = LeftFlipperKey Then Controller.Switch(6) = 0
    If keyCode = RightFlipperKey Then Controller.Switch(16) = 0
    If keycode = PlungerKey Then Plunger.Fire
 End Sub

 '*********
 ' Switches
 '*********

 ' Slingsshots
 Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), Lemk, 1
  DOF 101, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 66
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
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), Remk, 1
    DOF 102, DOFPulse
  RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 66
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

 ' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 55:PlaySoundAtVol SoundFXDOF("fx_bumper",103,DOFPulse,DOFContactors), ActiveBall, 1:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 55:PlaySoundAtVol SoundFXDOF("fx_bumper",104,DOFPulse,DOFContactors), ActiveBall, 1:End Sub

 ' Drain & holes
 Sub Drain_Hit:PlaysoundAtVol "drain", Drain, 1:bsTrough.AddBall Me:Gate.RotateToStart:PlayBonus.Enabled = 1:End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
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

'Set position as bumperX and Vol manually.

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
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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


 ' Rollovers
 Sub sw53_Hit:Controller.Switch(53) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l33.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub

 Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

 Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l34.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

 Sub sw54_Hit:Controller.Switch(54) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l35.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

 Sub sw64_Hit:Controller.Switch(64) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l36.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub

 Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l31.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

 Sub sw62_Hit:Controller.Switch(62) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l32.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

 Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l30.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

  Sub sw61a_Hit:Controller.Switch(61) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l30b.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw61a_UnHit:Controller.Switch(61) = 0:End Sub

 Sub sw50_Hit:Controller.Switch(50) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l27.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub

 Sub sw50a_Hit:Controller.Switch(50) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l27b.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw50a_UnHit:Controller.Switch(50) = 0:End Sub

 Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l28.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

  Sub sw60a_Hit:Controller.Switch(60) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l28b.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw60a_UnHit:Controller.Switch(60) = 0:End Sub

 Sub sw51_Hit:Controller.Switch(51) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l29.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

  Sub sw51a_Hit:Controller.Switch(51) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1
If l29b.State Then
  PlaySound SoundFXDOF("fx_5000",121,DOFPulse,DOFChimes)
Else
  PlaySound SoundFXDOF("fx_500",120,DOFPulse,DOFChimes)
End If
End Sub
 Sub sw51a_UnHit:Controller.Switch(51) = 0:End Sub

 ' Targets
 Sub sw65_Hit:vpmTimer.PulseSw 65:sw65p.X=sw65p.X+3:sw65p.Y=sw65p.Y-3:Me.TimerEnabled = 1:PlaySoundAtVol "fx_target", ActiveBall, 1 :Gate.RotateToStart:End Sub
 Sub sw65_Timer:sw65p.X=sw65p.X-3:sw65p.Y=sw65p.Y+3:PlaySound "Gun2":Me.TimerEnabled = 0:End Sub
 Sub sw65b_Hit:vpmTimer.PulseSw 65:sw65bp.X=sw65bp.X+3:sw65bp.Y=sw65bp.Y-3:Me.TimerEnabled = 1:PlaySoundAtVol "fx_target", ActiveBall, 1:Gate.RotateToEnd:End Sub
 Sub sw65b_Timer:sw65bp.X=sw65bp.X-3:sw65bp.Y=sw65bp.Y+3:PlaySound "Gun1":Me.TimerEnabled = 0:End Sub
  Sub sw65d_Hit:vpmTimer.PulseSw 65:sw65dp.X=sw65dp.X+3:sw65dp.Y=sw65dp.Y-3:Me.TimerEnabled = 1:PlaySoundAtVol "fx_target", ActiveBall, 1 :Gate.RotateToStart:End Sub
 Sub sw65d_Timer:sw65dp.X=sw65dp.X-3:sw65dp.Y=sw65dp.Y+3:PlaySound "Gun2":Me.TimerEnabled = 0:End Sub

 ' Droptargets
 Sub sw0_Hit:dtTBank.hit 1:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l12.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw10_Hit:dtTBank.hit 2:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l13.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw20_Hit:dtTBank.hit 3:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l14.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw30_Hit:dtTBank.hit 4:PlaySoundAtVol "fx_droptarget", Activeball, 1
If l15.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw40_Hit:dtTBank.hit 5:PlaySoundAtVol "fx_droptarget", Activeball, 1
If l16.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw1_Hit:dtTBank.hit 6:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l17.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw11_Hit:dtTBank.hit 7:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l18.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw21_Hit:dtTBank.hit 8:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l19.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw31_Hit:dtTBank.hit 9:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l20.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw41_Hit:dtTBank.hit 10:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l21.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw2_Hit:dtRBank.hit 1:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l22.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw12_Hit:dtRBank.hit 2:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l23.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw22_Hit:dtRBank.hit 3:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l24.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw32_Hit:dtRBank.hit 4:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l25.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw42_Hit:dtRBank.hit 5:PlaySoundAtVol "fx_droptarget", ActiveBall, 1
If l26.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub


 '*********
 'Solenoids
 '*********
 SolCallback(5) = "dtRBank.SolDropUp"
 SolCallback(6) = "dtTbank.SolDropUp"
 solcallback(8) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
 solcallback(9) = "bsTrough.SolOut"

 '**************
 ' Flipper Subs
 '**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), LeftFlipper, 1
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), LeftFlipper, 1
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), RightFlipper, 1
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), RightFlipper, 1
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

  '**************
 ' Extra sounds
 '**************

Sub aRubbers_Hit(idx):vpmTimer.PulseSw 66: PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPostRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

  '************************
 '      Check Bonus
 '************************

 Dim BonusLamps, OldState
 BonusLamps = Array(37, 38, 39, 40, 41, 42, 43, 44, 45, 46) 'bonus lights
 PlayBonus.Enabled = 0

 Sub PlayBonus_Timer()
    Dim lamp, ii, state
    ii = 0
    state = 0
    For each lamp in BonusLamps
       If LampState(lamp) Then state = state + 2 ^ii

       ii = ii + 1
    Next
    If(OldState <> state) Then PlaySound "fx_Bonus"
    OldState = state
    If(state = 1 and LampState(47) = 0) Or LampState(11) = 1 Then PlayBonus.Enabled = 0
 End Sub

 '************************************
 '          LEDs Display
 '************************************

 Dim Digits(32)

 Set Digits(0) = a0
 Set Digits(1) = a1
 Set Digits(2) = a2
 Set Digits(3) = a3
 Set Digits(4) = a4
 Set Digits(5) = a5
 Set Digits(6) = a6

 Set Digits(7) = b0
 Set Digits(8) = b1
 Set Digits(9) = b2
 Set Digits(10) = b3
 Set Digits(11) = b4
 Set Digits(12) = b5
 Set Digits(13) = b6

 Set Digits(14) = c0
 Set Digits(15) = c1
 Set Digits(16) = c2
 Set Digits(17) = c3
 Set Digits(18) = c4
 Set Digits(19) = c5
 Set Digits(20) = c6

 Set Digits(21) = d0
 Set Digits(22) = d1
 Set Digits(23) = d2
 Set Digits(24) = d3
 Set Digits(25) = d4
 Set Digits(26) = d5
 Set Digits(27) = d6

 Set Digits(28) = e0
 Set Digits(29) = e1
 Set Digits(30) = e2
 Set Digits(31) = e3

 Sub LEDs_Timer
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
       For ii = 0 To UBound(ChgLED)
          chg = chgLED(ii, 1):stat = chgLED(ii, 2)

          Select Case stat
             Case 0:Digits(chgLED(ii, 0) ).SetValue 0    'empty
             Case 63:Digits(chgLED(ii, 0) ).SetValue 1   '0
             Case 6:Digits(chgLED(ii, 0) ).SetValue 2    '1
             Case 91:Digits(chgLED(ii, 0) ).SetValue 3   '2
             Case 79:Digits(chgLED(ii, 0) ).SetValue 4   '3
             Case 102:Digits(chgLED(ii, 0) ).SetValue 5  '4
             Case 109:Digits(chgLED(ii, 0) ).SetValue 6  '5
             Case 124:Digits(chgLED(ii, 0) ).SetValue 7  '6
             Case 125:Digits(chgLED(ii, 0) ).SetValue 7  '6
             Case 252:Digits(chgLED(ii, 0) ).SetValue 7  '6
             Case 7:Digits(chgLED(ii, 0) ).SetValue 8    '7
             Case 127:Digits(chgLED(ii, 0) ).SetValue 9  '8
             Case 103:Digits(chgLED(ii, 0) ).SetValue 10 '9
             Case 111:Digits(chgLED(ii, 0) ).SetValue 10 '9
             Case 231:Digits(chgLED(ii, 0) ).SetValue 10 '9
             Case 128:Digits(chgLED(ii, 0) ).SetValue 0  'empty
             Case 191:Digits(chgLED(ii, 0) ).SetValue 1  '0
             Case 832:Digits(chgLED(ii, 0) ).SetValue 2  '1
             Case 896:Digits(chgLED(ii, 0) ).SetValue 2  '1
             Case 768:Digits(chgLED(ii, 0) ).SetValue 2  '1
             Case 134:Digits(chgLED(ii, 0) ).SetValue 2  '1
             Case 219:Digits(chgLED(ii, 0) ).SetValue 3  '2
             Case 207:Digits(chgLED(ii, 0) ).SetValue 4  '3
             Case 230:Digits(chgLED(ii, 0) ).SetValue 5  '4
             Case 237:Digits(chgLED(ii, 0) ).SetValue 6  '5
             Case 253:Digits(chgLED(ii, 0) ).SetValue 7  '6
             Case 135:Digits(chgLED(ii, 0) ).SetValue 8  '7
             Case 255:Digits(chgLED(ii, 0) ).SetValue 9  '8
             Case 239:Digits(chgLED(ii, 0) ).SetValue 10 '9
          End Select
       Next
    End IF
 End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array

            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If

    UpdateLamps
End Sub

 Sub UpdateLamps()

     NFadeL 3, l3
     NFadeL 4, l4
     NFadeL 5, l5
     NFadeL 6, l6
     NFadeL 7, l7
     FadeR 8, l8
     FadeR 10, l10
     FadeR 11, l11
     NFadeL 12, l12
      NFadeL 13, l13
     NFadeL 14, l14
     NFadeL 15, l15
      NFadeL 16, l16
     NFadeL 17, l17
     NFadeL 18, l18
     NFadeL 19, l19
     NFadeL 20, l20
     NFadeL 21, l21
     NFadeL 22, l22
     NFadeL 23, l23
     NFadeL 24, l24
     NFadeL 25, l25
      NFadeL 26, l26
     NFadeLm 27, l27b
      NFadeL 27, l27
     NFadeLm 28, l28b
      NFadeL 28, l28
      NFadeLm 29, l29b
      NFadeL 29, l29
      NFadeLm 30, l30b
     NFadeL 30, l30
     NFadeL 31, l31
      NFadeL 32, l32
     NFadeL 33, l33
     NFadeL 34, l34
     NFadeL 35, l35
     NFadeL 36, l36
     NFadeL 37, l37
     NFadeL 38, l38
     NFadeL 39, l39
     NFadeL 40, l40
     NFadeL 41, l41
     NFadeL 42, l42
      NFadeL 43, l43
     NFadeL 44, l44
     NFadeL 45, l45
     NFadeL 46, l46
     NFadeL 47, l47
    NFadeLm 90,bumper1light 'bumper 1
    NFadeL 90,bumper1light1 'bumper 1
     NFadeLm 91, bumper2light 'bumper 22
     NFadeL 91, bumper2light1 'bumper 22
  End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
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

'Set position as bumperX and Vol manually.

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
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here bacuse of the timing
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
    End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'******************
'   GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'******************

Dim OldGiState, GiIsOn
OldGiState = -1 'start witht the Gi off
GiIsOn = False

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then '-1 is the value returned by Getballs when no balls are on the table, so turn off the gi
            GiOff
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    Dim bulb
    GiIsOn = True
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    GiIsOn = False
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
  Gatep.RotZ = Gate.CurrentAngle
    RollingUpdate
  GIUpdate
End Sub

 'Gottlieb El Dorado City of Gold
 'by Inkochnito
 Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
       .AddForm 700, 400, "El Bueno el Feo y el Malo - DIP switches"
       .AddFrame 2, 2, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "20 credits", 49152)                                                                                   'dip 15&16
       .AddFrame 2, 80, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                   'dip 14
       .AddFrame 2, 126, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                       'dip 22
       .AddFrame 2, 172, 190, "Game mode", &H10000000, Array("replay", 0, "extra ball", &H10000000)                                                                                                                               'dip 29
       .AddChk 2, 227, 190, Array("Match feature", &H02000000)                                                                                                                                                                    'dip 26
       .AddChk 2, 248, 190, Array("Background sound", &H40000000)                                                                                                                                                                 'dip 31
       .AddFrame 205, 2, 190, "High score to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 replays", &H00400000, "displayed and 3 replays", &H00C00000) 'dip 23&24
       .AddFrame 205, 80, 190, "Balls per game", &H01000000, Array("5 balls", 0, "3 balls", &H01000000)                                                                                                                           'dip 25
       .AddFrame 205, 126, 190, "Replay limit", &H04000000, Array("no limit", 0, "one per ball", &H04000000)                                                                                                                      'dip 27
       .AddFrame 205, 172, 190, "Novelty", &H08000000, Array("normal", 0, "points", &H0800000)                                                                                                                                    'dip 28
       .AddFrame 205, 218, 190, "3rd coin chute credits control", &H20000000, Array("no effect", 0, "add 9", &H20000000)                                                                                                          'dip 30
       .AddLabel 50, 270, 300, 20, "After hitting OK, press F3 to reset game with new settings."
       .ViewDips
    End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

Sub Table1_Exit()
  Controller.Stop
End Sub

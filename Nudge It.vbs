
option explicit
Randomize

' Thalamus 2018-08-14 : Improved directional sounds
' Improved SSF directional sounds
' You need to import fx_coin and add RollingTimer

' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 20

Const BallSize = 90
Const swStartButton = 3
Const cGameName="nudgeit"
Const UseSolenoids=1,UseLamps=True,UseGI=0,UseSyn=1,SSolenoidOn="SolOn",SSolenoidOff="Soloff"
Const SCoin="fx_coin",cCredits="Nudge It"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
LoadVPM "01320000","gts3.vbs",3.1

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 0
else
    VarHidden = 1
end if
'if B2SOn = true then VarHidden = 1

'***************************************************************
'*             Keyboard Handlers                   *
'***************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=LeftFlipperKey Then Controller.Switch(4)=1
  If KeyCode=RightFlipperKey Then Controller.Switch(5)=1
If KeyCode=PlungerKey Then Plunger.PullBack:Controller.Switch(10)=True
If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=LeftFlipperKey Then Controller.Switch(4)=0
  If KeyCode=RightFlipperKey Then Controller.Switch(5)=0
If KeyCode=PlungerKey Then Plunger.Fire:PlaySoundAt "Plunger", plunger:Controller.Switch(10)=False
If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'***************************************************************
'*                   Solenoids                       *
'***************************************************************

'SolCallback(17)="flasher17"
'SolCallback(18)="vpmFlasher Light672,"
'SolCallback(19)="vpmFlasher Light673,"
'SolCallback(20)="vpmFlasher Light674,"
'SolCallback(21)="vpmFlasher Light675,"
'SolCallback(22)="vpmFlasher Light676,"
'SolCallback(23)="vpmSolSound ""bell"","
'SolCallback(24)=""                 'coin meter
'SolCallback(26)="vpmFlasher BGFlasher,"
'SolCallback(27)=""                 'ticket dispenser
SolCallback(28)="solpin"
'SolCallback(30)="vpmSolSound ""knocker"","
'SolCallback(31)="vpmNudge.SolGameOn"       'Tilt Relay (T)


'***************************************************************
'*                  Ball Gate                      *
'***************************************************************
dim foon, fo
Sub Solpin(Enabled)
     If Enabled Then
         pin.transy = -45
         pin.Collidable = 0
     Else
         pin.transy = 0
         pin.Collidable = 1
     End If
End Sub

'***************************************************************
'*              Animate Top Gate                     *
'***************************************************************

Sub GateTimer_Timer()
   Gate1P.RotZ = ABS(Gate1.currentangle)
End Sub

'***************************************************************
'*                Table Init                       *
'***************************************************************

Sub Table1_Init
  On Error Resume Next
  With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    '.SplashInfoLine="Nudge It" & vbnewline & "Table by Destruk"
    .ShowFrame=0
    .ShowTitle=0
    .ShowDMDOnly=1
    .HandleMechanics=0
    .HandleKeyboard=0
        .Hidden = VarHidden
    .Games(cGameName).Settings.Value("dmd_pos_x")=128
    .Games(cGameName).Settings.Value("dmd_pos_y")=330
    .Games(cGameName).Settings.Value("dmd_width")=300
    .Games(cGameName).Settings.Value("dmd_height")=65

  End With
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  'vpmNudge.TiltSwitch=151
  'vpmNudge.TiltObj=Array()
  'vpmNudge.Sensitivity=3

  vpmMapLights ALights
  Kicker.CreateSizedBall(29)
  Kicker.Kick 90,1

End Sub

'***************************************************************
'*             Rollover Switches                   *
'***************************************************************

Sub sw1_Hit():PlaySoundAt "rollover", ActiveBall:Controller.Switch(11) = 1:End Sub
Sub sw1_UnHit():Controller.Switch(11) = 0:End Sub
Sub sw2_Hit():PlaySoundAt "rollover", ActiveBall:Controller.Switch(12) = 1:End Sub
Sub sw2_UnHit():Controller.Switch(12) = 0:End Sub
Sub sw3_Hit():PlaySoundAt "rollover", ActiveBall:Controller.Switch(13) = 1:End Sub
Sub sw3_UnHit():Controller.Switch(13) = 0:End Sub
Sub sw4_Hit():PlaySoundAt "rollover", ActiveBall:Controller.Switch(14) = 1:End Sub
Sub sw4_UnHit():Controller.Switch(14) = 0:End Sub
Sub sw5_Hit():PlaySoundAt "rollover", ActiveBall:Controller.Switch(15) = 1:End Sub
Sub sw5_UnHit():Controller.Switch(15) = 0:End Sub
Sub sw6_Hit():PlaySoundAt "rollover", ActiveBall:Controller.Switch(16) = 1:End Sub
Sub sw6_UnHit():Controller.Switch(16) = 0:End Sub
Sub sw7_Hit:Controller.Switch(7) = 1:End Sub
Sub sw17_Hit():PlaySoundAt "rollover", ActiveBall:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit():Controller.Switch(17) = 0:End Sub

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

Const tnob = 5 ' total number of balls
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


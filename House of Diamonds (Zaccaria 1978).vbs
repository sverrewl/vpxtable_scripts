Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0
LoadVPM "01560000","ZAC2.VBS",3.2
Const cGameName="hod",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FX_FlipperUp",SFlipperOff="FX_FlipperDown"
Const SCoin="coin3",cCredits="House Of Diamonds"

' Thalamus 2019 October : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Dim obj
Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

Sub Table1_KeyDown(ByVal KeyCode)
If keycode=startgamekey then
  controller.switch(3)=1
  TextBox8.text=""
  BIP.text="Balls to Play"
  Exit Sub
End If

' Manual Ball Control
  If keycode = 46 Then          ' C Key
    If EnableBallControl = 1 Then
      EnableBallControl = 0
    Else
      EnableBallControl = 1
    End If
  End If
    If EnableBallControl = 1 Then
    If keycode = 48 Then        ' B Key
      If BCboost = 1 Then
        BCboost = BCboostmulti
      Else
        BCboost = 1
      End If
    End If
    If keycode = 203 Then BCleft = 1  ' Left Arrow
    If keycode = 200 Then BCup = 1    ' Up Arrow
    If keycode = 208 Then BCdown = 1  ' Down Arrow
    If keycode = 205 Then BCright = 1 ' Right Arrow
  End If

  If KeyCode=PlungerKey Then PlaySoundAtVol "Plungerpull", Plunger, 1:Plunger.Pullback
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
 'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If
  If keycode=startgamekey then
  controller.switch(3)=0
  Exit Sub
  End If

  If KeyCode=PlungerKey Then PlaySoundAtVol "Plunger", Plunger, 1:Plunger.Fire
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

SolCallback(1)="vpmSolSound ""lsling"","
SolCallback(6)="vpmSolSound ""rsling"","
'SolCallback(8)="vpmSolSound ""bumper"","
'SolCallback(9)="vpmSolSound ""bumper"","
SolCallback(10)="bsSaucer.SolOut"
SolCallback(11)="bsTrough.SolOut"
SolCallback(12)="vpmSolSound ""knocker"","
SolCallback(17)="vpmSolSound ""lsling"","
SolCallback(18)="vpmSolSound ""rsling"","
SolCallback(20)="vpmNudge.SolGameOn"
SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Flipper2,"
SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Flipper1,"

Dim bsTrough,bsSaucer,dtR

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

  If Table1.ShowDT = False then
    For each obj in Backdropstuff
      obj.visible=False
    next
  End If
  With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine=cCredits
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden=1
    .Run
    If Err Then MsgBox Err.Description
  End With
  PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=1:vpmNudge.Sensitivity=5:vpmNudge.TiltObj=Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,16,0,0,0,0,0,0
  bsTrough.InitKick BallRelease,90,5
  bsTrough.InitExitSnd "popper","ballsave"
  bsTrough.Balls=1

  Set dtR=New cvpmDropTarget
  dtR.InitSnd "flapclos","flapopen"

  Set bsSaucer=New cvpmBallStack
    bsSaucer.InitSaucer Kicker1,41,160,22
    bsSaucer.InitExitSnd "popper_ball","solon"

  vpmMapLights AllLights

 End Sub

Sub Drain_Hit
  Timer1.Interval = 1500:Timer1.Enabled = true
  PlaySoundAtVol "Drain", Drain, 1
  bsTrough.AddBall Me         'switch33
End Sub

Sub Timer1_Timer()
  Timer1.Enabled = false
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
  vpmTimer.PulseSw(18)
  GiEffect
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    vpmTimer.PulseSw(17)
  GiEffect
  LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Kicker1_Hit:PlaysoundAtVol "kicker_enter_center", ActiveBall, 1:bsSaucer.AddBall 0:End Sub
Sub Bumper1_Hit:vpmTimer.PulseSw(29):RandomSoundBumber:GiEffect:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw(37):RandomSoundBumber:GiEffect:End Sub

Sub LeftOutlane_Hit():vpmTimer.PulseSw(22):End Sub
Sub RightOutlane_Hit():vpmTimer.PulseSw(20):End Sub

Sub LeftInlane_Hit():vpmTimer.PulseSw(21):End Sub
Sub RightInlane_Hit():vpmTimer.PulseSw(19):End Sub
Sub Spinner1_Spin():vpmTimer.PulseSw(23):PlaySoundAtVol "fx_spinner", Spinner1, 1 :End Sub

Sub Side1_Hit():vpmTimer.PulseSw(31):End Sub 'top left rollover
Sub Trigger1_Hit():vpmTimer.PulseSw(39):End Sub 'mid right lane
Sub Spot1A_Hit():VpmTimer.PulseSw(30):End Sub
Sub Spot2A_Hit():VpmTimer.PulseSw(38):End Sub

Sub Drop1a_Hit:vpmTimer.PulseSw(28):End Sub
Sub Drop2a_Hit:vpmTimer.PulseSw(27):End Sub '?
Sub Drop3a_Hit:vpmTimer.PulseSw(26):End Sub
Sub Drop4a_Hit:vpmTimer.PulseSw(25):End Sub
Sub Drop5a_Hit:vpmTimer.PulseSw(24):End Sub

Sub Drop6a_Hit:vpmTimer.PulseSw(36):End Sub '?
Sub Drop7a_Hit:vpmTimer.PulseSw(35):End Sub
Sub Drop8a_Hit:vpmTimer.PulseSw(34):End Sub
Sub Drop9a_Hit:vpmTimer.PulseSw(33):End Sub
Sub Drop10a_Hit:vpmTimer.PulseSw(32):End Sub

Sub Wall4_Hit:vpmTimer.PulseSw(23):End Sub
Sub Wall5_Hit:vpmTimer.PulseSw(23):End Sub

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


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

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
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub aRubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundBumber
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtVol "fx_bumper1", ActiveBall, 1
    Case 2 : PlaySoundAtVol "fx_bumper2", ActiveBall, 1
    Case 3 : PlaySoundAtVol "fx_bumper3", ActiveBall, 1
    Case 4 : PlaySoundAtVol "fx_bumper4", ActiveBall, 1
  End Select
End Sub


Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub GiEffect
      LightSeqGi.UpdateInterval = 10
      LightSeqGi.Play SeqBlinking, , 4, 1
End Sub

'************************************
'          LEDs Display DT (From Kiwi's Farfalla)
'************************************

Sub LampTimer_Timer()
  UpdateLeds
End Sub



Dim Digits(33)

Set Digits(0) = da1
Set Digits(1) = da2
Set Digits(2) = da3
Set Digits(3) = da4
Set Digits(4) = da5
Set Digits(5) = da6

Set Digits(6) = db1
Set Digits(7) = db2
Set Digits(8) = db3
Set Digits(9) = db4
Set Digits(10) = db5
Set Digits(11) = db6

Set Digits(12) = dc1
Set Digits(13) = dc2
Set Digits(14) = dc3
Set Digits(15) = dc4
Set Digits(16) = dc5
Set Digits(17) = dc6

Set Digits(18) = dd1
Set Digits(19) = dd2
Set Digits(20) = dd3
Set Digits(21) = dd4
Set Digits(22) = dd5
Set Digits(23) = dd6

Set Digits(24) = de1
Set Digits(25) = de2
Set Digits(26) = de3
Set Digits(27) = de4
Set Digits(28) = de5
Set Digits(29) = de6

Set Digits(30) = df1
Set Digits(31) = df2
Set Digits(32) = df3
Set Digits(33) = df4



Sub UPdateLEDs
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
        'Case 125:Digits(chgLED(ii, 0) ).SetValue 7  '6
        Case 252:Digits(chgLED(ii, 0) ).SetValue 18  '6,
        Case 7:Digits(chgLED(ii, 0) ).SetValue 8    '7
        Case 127:Digits(chgLED(ii, 0) ).SetValue 9  '8
        Case 103:Digits(chgLED(ii, 0) ).SetValue 10 '9
        'Case 111:Digits(chgLED(ii, 0) ).SetValue 10 '9
        Case 231:Digits(chgLED(ii, 0) ).SetValue 21 '9,
        Case 128:Digits(chgLED(ii, 0) ).SetValue 11  'Comma
        Case 191:Digits(chgLED(ii, 0) ).SetValue 12  '0,
        'Case 832:Digits(chgLED(ii, 0) ).SetValue 2  '1
        'Case 896:Digits(chgLED(ii, 0) ).SetValue 2  '1
        'Case 768:Digits(chgLED(ii, 0) ).SetValue 2  '1
        Case 134:Digits(chgLED(ii, 0) ).SetValue 13  '1,
        Case 219:Digits(chgLED(ii, 0) ).SetValue 14  '2,
        Case 207:Digits(chgLED(ii, 0) ).SetValue 15  '3,
        Case 230:Digits(chgLED(ii, 0) ).SetValue 16  '4,
        Case 237:Digits(chgLED(ii, 0) ).SetValue 17  '5,
        'Case 253:Digits(chgLED(ii, 0) ).SetValue 18  '6,
        Case 135:Digits(chgLED(ii, 0) ).SetValue 19  '7,
        Case 255:Digits(chgLED(ii, 0) ).SetValue 20  '8,
        'Case 239:Digits(chgLED(ii, 0) ).SetValue 21 '9,
      End Select
    Next
  End IF
End Sub

' Thalamus - 2021-04-30 : added proper exit

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


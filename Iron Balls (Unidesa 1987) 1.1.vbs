Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="ironball",UseSolenoids=1,UseLamps=1,UseGI=0, SCoin="coin3"


LoadVPM "02060000","ironballs.vbs",3.2

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------

Const FourCaptiveBalls = 0    'set to 0 to start with 2 captive Balls on the passage loop, 1 will add another 2 balls

Const cCredits="Irnon Balls - Unidesa 1986"

Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp001.visible=1
Ramp002.visible=1
Primitive007.visible=1
Else
Ramp001.visible=0
Ramp002.visible=0
Primitive007.visible=0
End if

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------
'SolCallback(1)= "Sol1" 'GameOn
SolCallback(3)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(4)= "bsTrough.SolOut"
'5 Bumper
'6 Bumper
'7 Bumper
SolCallback(8)= "DropTargetBank.SolDropUp"
SolCallback(9)= "bsRightHole.SolOut"
SolCallback(10)="bsLeftHole.SolOut"
'11 Left Slingshot
'12 Right Slingshot

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("flipperup",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("flipperdown",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("flipperup",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("flipperdown",DOFFlippers), RIghtFlipper, 1:RightFlipper.RotateToStart
     End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'--------------------------
'------  Table Init  ------
'--------------------------

Dim bsTrough, bsLeftHole, bsRightHole, DropTargetBank

Sub Table1_Init
  vpminit me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Laser Cue Williams 1984"&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 1
        .Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch=4
    vpmNudge.Sensitivity=5

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 0,6,0,0,0,0,0,0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=1

  set DropTargetBank = new cvpmDropTarget
    DropTargetBank.InitDrop Array(sw42,sw52,sw43,sw53,sw44,sw54), Array(42,52,43,53,44,54)
    DropTargetBank.InitSnd SoundFX("Targetdrop1",DOFDropTargets),SoundFX("TargetBankreset1",DOFContactors)
    DropTargetBank.CreateEvents "DropTargetBank"

  Set bsLeftHole = New cvpmBallStack
    bsLeftHole.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("Solenoid",DOFContactors)
    bsLeftHole.initsaucer sw32,32,20+rnd(5),8+rnd(3)
    bsLeftHole.KickForceVar = 3
    bsLeftHole.KickAngleVar = 3

  Set bsRightHole = New cvpmBallStack
    bsRightHole.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("Solenoid",DOFContactors)
    bsRightHole.initsaucer sw31,31,-20-rnd(5),8+rnd(3)
    bsRightHole.KickForceVar = 3
    bsRightHole.KickAngleVar = 3

    vpmMapLights AllLights

  Controller.Switch(5)=1    'Coin Door closed

'main captive balls
  CaptiveKicker.createball
  CaptiveKicker.kick 180,1
  CaptiveKicker001.createball
  CaptiveKicker001.kick 180,1
  CaptiveKicker002.createball
  CaptiveKicker002.kick 180,1

'Options Based Captive Balls
  if FourCaptiveBalls = 1 then
    RightCaptiveKicker2.createball
    RightCaptiveKicker2.kick 180,1
    LeftCaptiveKicker2.createball
    LeftCaptiveKicker2.kick 180,1
  end if

End Sub


'------------------------------
'------  Trough Handler  ------
'------------------------------

Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain5" , Drain, 1: End Sub
Sub sw32_Hit:bsLeftHole.addBall 0 : playsoundAtVol "popper_ball", sw32, 1:End Sub
Sub sw31_Hit:bsRightHole.addBall 0 : playsoundAtVol "popper_ball",sw32, 1:End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftMagnaSave Then bLutActive = True
  If keycode = RightMagnaSave Then
    If bLutActive Then NextLUT: End If
    End If
  If keycode = PlungerKey Then Plunger.PullBack:playsoundAtVol"plungerpull", Plunger, 1
  If keycode = LeftMagnaSave then Controller.Switch(9) = 1
  If keycode = RightMagnaSave then Controller.Switch(10) = 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode = LeftMagnaSave Then bLutActive = False
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
  If keycode = LeftMagnaSave then Controller.Switch(9) = 0
  If keycode = RightMagnaSave then Controller.Switch(10) = 0
  If KeyUpHandler(keycode) Then Exit Sub
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 48
    PlaySoundAtVol SoundFX("Slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
' gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 47
    PlaySoundAtVol SoundFX("Slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
' gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

'********** Bumpers **********

Sub LeftBumper_hit : vpmTimer.PulseSw(35) : PlaysoundAtVol SoundFX("Bumper",DOFContactors),ActiveBall,1: End Sub
Sub RightBumper_hit : vpmTimer.PulseSw(36) : PlaysoundAtVol SoundFX("Bumper",DOFContactors),ActiveBall, 1: End Sub
Sub MiddleBumper_hit : vpmTimer.PulseSw(37) : PlaysoundAtVol SoundFX("Bumper",DOFContactors),ActiveBall, 1: End Sub

'********** Rollover Trigger **********

sub sw56_hit:Controller.Switch(56)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw56_unhit:Controller.Switch(56)=0:End Sub

sub sw21_hit:Controller.Switch(21)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw21_unhit:Controller.Switch(21)=0 :End Sub
sub sw16_hit:Controller.Switch(16)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw16_unhit:Controller.Switch(16)=0:End Sub

sub sw22_hit:Controller.Switch(22)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw22_unhit:Controller.Switch(22)=0:End Sub
sub sw17_hit:Controller.Switch(17)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw17_unhit:Controller.Switch(17)=0:End Sub

sub sw11_hit:Controller.Switch(11)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw11_unhit:Controller.Switch(11)=0:End Sub
sub sw12_hit:Controller.Switch(12)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw12_unhit:Controller.Switch(12)=0:End Sub

sub sw14_hit:Controller.Switch(14)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw14_unhit:Controller.Switch(14)=0:End Sub
sub sw13_hit:Controller.Switch(13)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw13_unhit:Controller.Switch(13)=0:End Sub
sub sw15_hit:Controller.Switch(15)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub sw15_unhit:Controller.Switch(15)=0:End Sub

'********** Round Targets **********

sub sw23_Hit:vpmTimer.PulseSw 23 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
sub sw41_hit:vpmTimer.PulseSw 41 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
sub sw51_hit:vpmTimer.PulseSw 51 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
sub sw26_hit:vpmTimer.PulseSw 26 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub

sub sw24_hit:vpmTimer.PulseSw 24 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
sub sw45_hit:vpmTimer.PulseSw 45 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
sub sw55_hit:vpmTimer.PulseSw 55 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
sub sw27_hit:vpmTimer.PulseSw 27 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub

sub sw25_hit:vpmTimer.PulseSw 25 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
sub sw33_hit:vpmTimer.PulseSw 33 : PlaySoundAtVol "Target1", ActiveBall, 1:End Sub


'********** Rubber hits **********

sub Rubber1_hit:vpmTimer.PulseSw 18 : PlaySoundAtVol "Bump", ActiveBall, 1:End Sub   'Low
sub Rubber2_hit:vpmTimer.PulseSw 38 : PlaySoundAtVol "Bump", ActiveBall, 1:End Sub   'Mid
sub Rubber3_hit:vpmTimer.PulseSw 28 : PlaySoundAtVol "Bump", ActiveBall, 1:End Sub   'Top
sub Rubber4_hit:vpmTimer.PulseSw 28 : PlaySoundAtVol "Bump", ActiveBall, 1:End Sub      'Top
' Thalamus : This sub is used twice - this means ... this one WAS NOT USED
' Probably a typo - would need to check table
' sub Rubber5_hit:vpmTimer.PulseSw 38 : PlaySound "Bump",0,0.1,0.15,0.15:End Sub      'Mid
sub Rubber5_hit:vpmTimer.PulseSw 18 : PlaySoundAtVol "Bump", ActiveBall, 1:End Sub      'Low

'********** Gates **********

Sub Gate1_Hit:PlaySoundAtVol "Gate5",ActiveBall, 1:End Sub
Sub Gate2_Hit:PlaySoundAtVol "Gate5",ActiveBall, 1:End Sub


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(35)

Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)
Digits(32) = Array(LED88,LED79,LED97,LED98,LED89,LED78,LED87)
Digits(33) = Array(LED109,LED107,LED118,LED119,LED117,LED99,LED108)
Digits(34) = Array(LED137,LED128,LED139,LED147,LED138,LED127,LED129)


Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 35) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else

      end if
    next
    end if
end if
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
    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

    LFPrim.objRotz = LeftFlipper.CurrentAngle
  RFPrim.objRotz= RightFlipper.CurrentAngle
End Sub

'*************
'   JP'S LUT
'*************

Dim bLutActive, LUTImage
Sub LoadLUT
Dim x
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 10: UpdateLUT: SaveLUT: End Sub

Sub UpdateLUT
Select Case LutImage
Case 0: table1.ColorGradeImage = "LUT0"
Case 1: table1.ColorGradeImage = "LUT1"
Case 2: table1.ColorGradeImage = "LUT2"
Case 3: table1.ColorGradeImage = "LUT3"
Case 4: table1.ColorGradeImage = "LUT4"
Case 5: table1.ColorGradeImage = "LUT5"
Case 6: table1.ColorGradeImage = "LUT6"
Case 7: table1.ColorGradeImage = "LUT7"
Case 8: table1.ColorGradeImage = "LUT8"
Case 9: table1.ColorGradeImage = "LUT9"

End Select
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

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


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

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
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

Sub Table1_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub

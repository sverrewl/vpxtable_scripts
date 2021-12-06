Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01530000","sega.vbs",3.1


' Thalamus 2020 April : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Const VolFlip   = 1    ' Flipper volume.

'********************************************
'**     Game Specific Code Starts Here     **
'********************************************

Const UseSolenoids=2,UseSync=1
Const SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="Coin3"

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=KeyUpperLeft Then Controller.Switch(1)=1
  If KeyCode=KeyUpperRight Then Controller.Switch(8)=1
  If KeyCode=PlungerKey Then Controller.Switch(53)=1
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=KeyUpperLeft Then Controller.Switch(1)=0
  If KeyCode=KeyUpperRight Then Controller.Switch(8)=0
  If KeyCode=PlungerKey Then Controller.Switch(53)=0
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'**************************************
'**     Bind Events To Solenoids     **
'**************************************

SolCallback(1)="SolTrough"
SolCallback(2)="vpmSolAutoPlunger Plunger,0,"
SolCallback(2)="AutoPlunger"
SolCallback(3)="bsTRVUK.SolOut"
SolCallback(4)="bsBRVUK.SolOut"
SolCallback(6)="bsCVUK.SolOut"
' SolCallback(9)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
' SolCallback(10)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
' SolCallback(11)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
' SolCallback(12)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
' SolCallback(13)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(9)="Bumper1"
SolCallback(10)="Bumper1"
SolCallback(11)="Bumper1"
SolCallback(12)="Bumper1"
SolCallback(13)="Bumper1"
SolCallback(14)="mDMag.MagnetOn="
SolCallback(17)="mTurnTable.MotorOn="
SolCallback(18)="vpmSolSound SoundFX(""Sling"",DOFContactors),"
SolCallback(19)="vpmSolSound SoundFX(""Sling"",DOFContactors),"
SolCallback(20)="SolRobot"
'FLASHERS
SolCallback(25)="vpmFlasher Array(F1A,F1B,F1C,F1D),"'25'F1=Red*4
SolCallback(26)="vpmFlasher Array(F2A,F2B,F2C,F2D),"'26'F2=Yellow*4
SolCallback(27)="vpmFlasher Array(F3A,F3B,F3C,F3D),"'27'F3=Green*4
'28'F4=WARNing*4
'29'F5=warnING*4
'SolCallback(30)="vpmFlasher Array(Light6,Light7,Light8),"
'31'F7=Pops*2
'32'F8=Ramp*2

Sub Bumper1(enabled)
  If Enabled Then
  PlaySoundAtVol SoundFX("Jet3", DOFContactors), S40, 1
  End If
End Sub

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

  Sub SolLFlipper(Enabled)
    If Enabled Then
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), LeftFlipper, VolFlip
    LeftFlipper.RotateToEnd
  Else
     PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), LeftFlipper, VolFlip
     LeftFlipper.RotateToStart
  End If
  End Sub

  Sub SolRFlipper(Enabled)
     If Enabled Then
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), RightFlipper, VolFlip
     RightFlipper.RotateToEnd
     Else
     PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), RightFlipper, VolFlip
     RightFlipper.RotateToStart
     End If
  End Sub

Sub AutoPlunger(Enabled)
    If Enabled Then
       Plunger.Fire
       PlaySoundAtVol SoundFX("SolOn",DOFContactors), Plunger, 1
  End If
End Sub


'*********************Flashers
Sub Sol25(Enabled)
  If Enabled Then
F1A.State=LightStateOn
F1B.State=LightStateOn
F1C.State=LightStateOn
F1D.State=LightStateOn
Else
F1A.State=LightStateOff
F1B.State=LightStateOff
F1C.State=LightStateOff
F1D.State=LightStateOff
End If
End Sub

Sub Sol26(Enabled)
  If Enabled Then
F2A.State=LightStateOn
F2B.State=LightStateOn
F2C.State=LightStateOn
F2D.State=LightStateOn
Else
F2A.State=LightStateOff
F2B.State=LightStateOff
F2C.State=LightStateOff
F2D.State=LightStateOff
End If
End Sub

Sub Sol27(Enabled)
  If Enabled Then
F3A.State=LightStateOn
F3B.State=LightStateOn
F3C.State=LightStateOn
F3D.State=LightStateOn
Else
F3A.State=LightStateOff
F3B.State=LightStateOff
F3C.State=LightStateOff
F3D.State=LightStateOff
End If
End Sub
'*****************************

Sub SolRobot(Enabled)
  If Enabled Then
    Robot1.TimerEnabled=1
    Robot1_Timer
  Else
    Robot1.TimerEnabled=0
  End If
End Sub

Sub Robot1_Timer
  If Robot2.IsDropped Then
    Robot1.IsDropped=1
    Robot2.IsDropped=0
  Else
    Robot2.IsDropped=1
    Robot1.IsDropped=0
  End If
End Sub

Sub SolTrough(Enabled)
  If Enabled Then
    If bsTrough.Balls Then
      vpmTimer.PulseSw 15
      bsTrough.ExitSol_On
    End If
  End If
End Sub

'********************************************
'**     Init The Table, Start VPinMAME     **
'********************************************

Dim bsTrough,bsCVUK,bsTRVUK,bsBRVUK,mTLMag,mTRMag,mDMag,TurnTable,X,SpinCounter,mTurnTable
SpinCounter=Array(spin1,spin2,spin3,spin4,spin5,spin6,spin7,spin8,spin9,spin10,spin11,spin12)

Sub Table1_Init
    Robot2.IsDropped=1
'*********************FlasherF
F1A.State=LightStateOff
F1B.State=LightStateOff
F1C.State=LightStateOff
F1D.State=LightStateOff
F2A.State=LightStateOff
F2B.State=LightStateOff
F2C.State=LightStateOff
F2D.State=LightStateOff
F3A.State=LightStateOff
F3B.State=LightStateOff
F3C.State=LightStateOff
F3D.State=LightStateOff
'*****************************
  For X=0 To 11
    SpinCounter(X).IsDropped=1
  Next
  Plunger.PullBack
    On Error Resume Next
  With Controller
  .GameName="lostspc"
  .SplashInfoLine=""
  .HandleKeyboard=0
  .ShowTitle=0
  .ShowDMDOnly=1
    .HandleMechanics=0
  .ShowFrame=0
    If Table1.ShowDT = false then
      'Scoretext.Visible = false
      .Hidden = 1
    End If

    If Table1.ShowDT = true then
    'Scoretext.Visible = false
      .Hidden = 0
    End If
  End With
  Controller.Run
  If Err Then MsgBox Err.Description
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=56:vpmNudge.Sensitivity=5:vpmNudge.TiltObj=Array(S36,S37,S38,S39,S40,LeftSlingshot,RightSlingshot)

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,14,13,12,11,0,0,0
  bsTrough.InitKick BallRelease,95,5
  bsTrough.Balls=4
  bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("SolOn",DOFContactors)

  Set bsCVUK=New cvpmBallStack
  bsCVUK.InitSw 0,46,0,0,0,0,0,0
  bsCVUK.InitKick CenterHole,160,8
  bsCVUK.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("SolOn",DOFContactors)

  Set bsTRVUK=New cvpmBallStack
  bsTRVUK.InitSw 0,47,0,0,0,0,0,0
  bsTRVUK.InitKick TRVExit,280,20
  bsTRVUK.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("SolOn",DOFContactors)

  Set bsBRVUK=New cvpmBallStack
  bsBRVUK.InitSw 0,48,0,0,0,0,0,0
  bsBRVUK.InitKick BRVExit,300,20
  bsBRVUK.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("SolOn",DOFContactors)

  Set mTLMag=New cvpmMagnet
  mTLMag.InitMagnet Magnet1,200
  mTLMag.Solenoid=5
  mTLMag.GrabCenter=1
  mTLMag.CreateEvents"mTLMag"

  Set mTRMag=New cvpmMagnet
  mTRMag.InitMagnet Magnet2,50
  mTRMag.Solenoid=7
  mTRMag.GrabCenter=1
  mTRMag.CreateEvents"mTRMag"

  Set mDMag=New cvpmMagnet
  mDMag.InitMagnet Magnet3,19
  mDMag.GrabCenter=0

  Set TurnTable=New cvpmMech
  TurnTable.Sol1=17
  TurnTable.mType=vpmMechLinear+vpmMechCircle+vpmMechOneSol+vpmMechFast
  TurnTable.Length=120
  TurnTable.Steps=12
  TurnTable.Acc=360
  TurnTable.Ret=2
  TurnTable.Callback=GetRef("UpdateTurnTable")
  TurnTable.Start

  Set mTurnTable=New cvpmTurnTable
  mTurnTable.InitTurnTable TT,50

  vpmMapLights AllLights
End Sub

set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
  If Enabled Then
    dim xx
    For each xx in GI:xx.State = 1: Next
    DOF 101, DOFOn
        PlaySound "fx_relay"

  Else For each xx in GI:xx.State = 0: Next
    DOF 101, DOFOff
        PlaySound "fx_relay"
  End If
End Sub

'SWITCHES
'--------------------------------------------------------------------------
                                        '1=Left UK Button
                                        '2=Coin4
                                        '3=Coin6
                                        '4=Coin2
                                        '5=Coin3
                                        '6=Coin1
                                        '7=Coin5
                                        '8=Right UK Button
                                        '9=NOT USED
                                        '10=NOT USED
Sub Drain_Hit:bsTrough.AddBall Me:End Sub                   '11,12,13,14,15
Sub S16_Hit:Controller.Switch(16)=1:End Sub                   '16
Sub S16_unHit:Controller.Switch(16)=0:End Sub
Sub S17_Hit:vpmTimer.PulseSw 17:End Sub                     '17
Sub S18_Hit:vpmTimer.PulseSw 18:End Sub                     '18
Sub S19_Hit:vpmTimer.PulseSw 19:End Sub                     '19
Sub S20_Hit:Controller.Switch(20)=1:End Sub                   '20
Sub S20_unHit:Controller.Switch(20)=0:End Sub
Sub S21_Hit:Controller.Switch(21)=1:End Sub                   '21
Sub S21_unHit:Controller.Switch(21)=0:End Sub
                                        '22=NOT USED
                                        '23=NOT USED
                                        '24=NOT USED
Sub S25_Hit:vpmTimer.PulseSw 25:End Sub                     '25
Sub S26_Hit:vpmTimer.PulseSw 26:End Sub                     '26
Sub S27_Hit:vpmTimer.PulseSw 27:End Sub                     '27
                                        '28=NOT USED
                                        '29=NOT USED
Sub S30_Hit:Controller.Switch(30)=1:End Sub                   '30
Sub S30_unHit:Controller.Switch(30)=0:End Sub
Sub S31_Hit:Controller.Switch(31)=1:End Sub                   '31
Sub S31_unHit:Controller.Switch(31)=0:End Sub
Sub S32_Hit:Controller.Switch(32)=1:End Sub                   '32
Sub S32_unHit:Controller.Switch(32)=0:End Sub
Sub S33_Hit:vpmTimer.PulseSw 33:End Sub                     '33
Sub S34_Hit:vpmTimer.PulseSw 34:End Sub                     '34
Sub S35_Hit:vpmTimer.PulseSw 35:End Sub                     '35
Sub S36_Hit:vpmTimer.PulseSw 36:End Sub                     '36
Sub S37_Hit:vpmTimer.PulseSw 37:End Sub                     '37
Sub S38_Hit:vpmTimer.PulseSw 38:End Sub                     '38
Sub S39_Hit:vpmTimer.PulseSw 39:End Sub                     '39
Sub S40_Hit:vpmTimer.PulseSw 40:End Sub                     '40
Sub PopExit_Hit:Me.DestroyBall:vpmTimer.PulseSwitch 41,100,"AddMystery":End Sub '41
                                        '42=NOT USED
                                        '43=NOT USED
                                        '44=NOT USED
Sub RobotEnter_Hit:Me.DestroyBall:vpmTimer.PulseSwitch 45,100,"AddBRV":End Sub  '45=UTrough Robot
Sub AddBRV(swNo):bsBRVUK.AddBall 0:End Sub
Sub CenterHole_Hit:bsCVUK.AddBall Me:End Sub                    '46
Sub AddMystery(swNo):bsCVUK.AddBall 0:End Sub
Sub TRHole_Hit:bsTRVUK.AddBall Me:End Sub                   '47
                                        '48=Bottom Right VUK
Sub S49_Spin:vpmTimer.PulseSw 49:End Sub                    '49
Sub S50_Spin:vpmTimer.PulseSw 50:End Sub                    '50
Sub S51_Hit:Controller.Switch(51)=1:End Sub                   '51
Sub S51_unHit:Controller.Switch(51)=0:End Sub
Sub S52_Hit:Controller.Switch(52)=1:End Sub                   '52
Sub S52_unHit:Controller.Switch(52)=0:End Sub
                                        '53=Launch Button
                                        '54=Start Button
                                        '55=Slam Tilt
                                        '56=Plumb Bob Tilt
Sub S57_Hit:Controller.Switch(57)=1:End Sub                   '57
Sub S57_unHit:Controller.Switch(57)=0:End Sub
Sub S58_Hit:Controller.Switch(58)=1:End Sub                   '58
Sub S58_unHit:Controller.Switch(58)=0:End Sub
Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 59:End Sub             '59
Sub S60_Hit:Controller.Switch(60)=1:End Sub                   '60
Sub S60_unHit:Controller.Switch(60)=0:End Sub
Sub S61_Hit:Controller.Switch(61)=1:End Sub                   '61
Sub S61_unHit:Controller.Switch(61)=0:End Sub
Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 62:End Sub            '62

'SUPPORTING ROUTINES
'--------------------------------------------------------------------------
Sub Magnet3_Hit
  mDMag.AddBall ActiveBall
  mDMag.AttractBall ActiveBall
End Sub
Sub Magnet3_unHit
  mDMag.RemoveBall ActiveBall
End Sub

Sub TT_Hit
  mTurnTable.AddBall ActiveBall
  mTurnTable.AffectBall ActiveBall
End Sub
Sub TT_unHit
  mTurnTable.RemoveBall ActiveBall
End Sub

Sub UpdateTurnTable(aNewPos,aSpeed,aLastPos)
  If aLastPos>-1 And aLastPos<10 Then SpinCounter(ALastPos).IsDropped=1
  If aNewPos>-1 And aNewPos<10 Then SpinCounter(aNewPos).IsDropped=0
End Sub

Sub Trigger1_Hit:ActiveBall.VelY=.1:ActiveBall.VelX=0:End Sub
Sub Trigger2_Hit:ActiveBall.VelY=1:ActiveBall.VelX=0:End Sub
Sub Trigger3_Hit:ActiveBall.VelY=1:ActiveBall.VelX=0:End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
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


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

Const tnob = 7 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, bb
    BOT = GetBalls

    ' stop the sound of deleted balls
    For bb = UBound(BOT) + 1 to tnob
        rolling(bb) = False
        StopSound("fx_ballrolling" & bb)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

   For bb = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(bb) ) > 1 AND BOT(bb).z < 30 Then
            rolling(bb) = True
           PlaySound("fx_ballrolling" & bb), -1, Vol(BOT(bb)), AudioPan(BOT(bb)), 0, Pitch(BOT(bb)), 1, 0, AudioFade(BOT(bb))
        Else
           If rolling(bb) = True Then
               StopSound("fx_ballrolling" & bb)
                rolling(bb) = False
           End If
        End If

        ' play ball drop sounds
       If BOT(bb).VelZ < -1 and BOT(bb).z < 55 and BOT(bb).z > 27 Then 'height adjust for ball drop sounds
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

    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle

    sw28p.RotY = sw28.CurrentAngle

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
  Controller.Stop
End Sub

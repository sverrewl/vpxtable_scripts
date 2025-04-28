Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName="petaco2",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits="Petaco 2"

LoadVPM "01520000","juegos.vbs",3.1

 Sub Table1_KeyDown(ByVal KeyCode)
 if keycode=16 then s3_hit
 if keycode=17 then s4_hit
 if keycode=18 then s5_hit

if keycode=21 then vpmtimer.pulsesw 15
if keycode=22 then vpmtimer.pulsesw 16

 if keycode=30 then vpmtimer.pulsesw 1
 if keycode=31 then vpmtimer.pulsesw 2
 if keycode=32 then vpmtimer.pulsesw 9
 if keycode=33 then vpmtimer.pulsesw 10
 if keycode=34 then vpmtimer.pulsesw 11
 if keycode=35 then vpmtimer.pulsesw 12
 if keycode=36 then vpmtimer.pulsesw 17
 if keycode=37 then vpmtimer.pulsesw 32
 if keycode=38 then vpmtimer.pulsesw 33
 if keycode=39 then vpmtimer.pulsesw 34

   If KeyCode=LeftFlipperKey Then Controller.Switch(84)=1
  If KeyCode=RightFlipperKey Then Controller.Switch(82)=1
 If KeyCode=PlungerKey Then Plunger.Pullback
 If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=LeftFlipperKey Then Controller.Switch(84)=0
  If KeyCode=RightFlipperKey Then Controller.Switch(82)=0
 If KeyCode=PlungerKey Then PlaySoundAtVol"plunger", Plunger, 1:Plunger.Fire
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

' SolCallback(8)="vpmFlasher Flasher8,"
' SolCallback(17)="VariReset"
 SolCallback(9)="vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(10)="vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(11)="vpmSolSound SoundFX(""jet3"",DOFContactors),"
SolCallback(12)="vpmSolSound SoundFX(""jet3"",DOFContactors),"
SolCallback(13)="dtR.SolDropUp"
SolCallback(14)="dtL.SolDropUp"
SolCallback(15)="vpmSolSound SoundFX(""knocker"",DOFContactors),"
 SolCallback(16)="bsTrough.SolOut"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("FlipperUp", DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("FlipperDown", DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("FlipperUp", DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("FlipperDown", DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart
     End If
End Sub


Dim bsTrough,cbCaptive,dtL,dtR,VariTReset

Sub Table1_Paused:Controller.Pause=True:End Sub
Sub Table1_UnPaused:Controller.Pause=False:End Sub

Sub Table1_Init
 Vari2.IsDropped=1:Vari3.IsDropped=1:Vari4.IsDropped=1:Vari5.IsDropped=1:Vari6.IsDropped=1
  TriggerVT1.Enabled=1:TriggerVT2.Enabled=1:TriggerVT3.Enabled=1:TriggerVT4.Enabled=1:TriggerVT5.Enabled=1:TriggerVT6.Enabled=1
   GPush.IsDropped=1
 With Controller
   .GameName=cGameName
   If Err Then MsgBox"Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
   .SplashInfoLine=cCredits
    .HandleMechanics=0
    .HandleKeyboard=0
   .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Run GetPlayerHwnd
    If Err Then MsgBox Err.Description
  End With

  PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=31:vpmNudge.Sensitivity=5:vpmNudge.TiltObj=Array(S13,S14,S15,S16)

 'Coins 1 (4 on keyboard) = dip 0 switch 1/dip 0 switch 3
 'OFF OFF = 1:1   OFF ON =1:2
 'ON ON   = 2:1   ON OFF =1:3
 'Coins 2 (5 on keyboard) = dip 0 switch 5/dip 0 switch 7
 'OFF OFF = 1:1   OFF ON =1:2
 'ON ON   = 1:10  ON OFF =1:3

 'Extra Ball Award = dip 0 switch 4
 'Replay Level A = ON = 600.000 EB
 'Replay Level B = ON = 700.000 EB
 'Replay Level C = ON = 900.000 EB
 'Replay Level D = ON = 800.000 EB

 'Replay Levels = dip 1 switch 1/dip 1 switch 2
 'Replay Level A = OFF OFF=800.000/1.000.000/1.200.000
'Replay Level B = ON OFF=900.000/1.000.000/1.200.000
 'Replay Level C = ON ON =1.000.000/1.500.000/2.000.000
 'Replay Level D = OFF ON=1.000.000/1.000.000/1.400.000

 'BALLS = Controller.Dip 1 Switch 3 - off=3/on=5
'Ignore first replay level if under 1.000.000= dip 0 switch 6
 'ON=Ignore, OFF=Award Free Game
 'Only allow 1 score-based replay per game - to disable score based replays, enable 'ignore first replay level' feature
 'dip 0 switch 7
 'MATCH = Controller.Dip 1 Switch 8 - off=disabled/on=enabled

  'Coin Audits = Controller.Dip 2 Switch 3 - off=disabled/on=enabled
'Display Credit audit at end of each game = dip 2 switch 4
 'ON=Enabled, OFF=Disabled


'                    *           *     *     *      *      *      *
'Controller.Dip(0)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '01-08
'                    *     *     *                                *
'Controller.Dip(1)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '09-16
'                                *     *
'Controller.Dip(2)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) 'A-D


   Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,25,0,0,0,0,0,0
  bsTrough.InitKick BallRelease,107,7
 bsTrough.InitExitSnd SoundFX("ballrel",DOFcontactors),"solon"
 bsTrough.Balls=1

  Set cbCaptive=New cvpmCaptiveBall
 cbCaptive.InitCaptive CaptiveTrigger1,Wall16,Kicker1,328
  cbCaptive.Start
 cbCaptive.ForceTrans=1.5
  cbCaptive.MinForce=3.5
  cbCaptive.CreateEvents"cbCaptive"

 Set dtL=New cvpmDropTarget
  dtL.InitDrop Array(S3,S4,S5),Array(3,4,5)
 dtL.InitSnd SoundFX("FlapClos",DOFContactors),SoundFX("FlapOpen",DOFContactors)

 Set dtR=New cvpmDropTarget
  dtR.InitDrop Array(S6,S7,S8),Array(6,7,8)
 dtR.InitSnd SoundFX("FlapClos",DOFContactors),SoundFX("FlapOpen",DOFContactors)

vpmMapLights GoodLights
 vpmMapLights TestLights
End Sub


' Quick Ball Sound V1.1 by STAT, stefanaustria
' -----------------------
Dim aX, aY, sX, RSound, SBall

Sub TriggerS_Timer()
  aX = int(SBall.VelX): aY = int(SBall.VelY)
  sX = -1.0: If int(Sball.X)>500 Then sX = 1.0
  If (aX>5 OR aY>5) AND Rsound = 0 Then
    RSound=int(RND*4)+1
    PlaySound "Roll "&RSound,0,1.0,sX,0.2 'replace sX to 0 for FullSound
  Elseif (aX<6 AND aY<6) AND Rsound > 0 Then
    StopSound "Roll "&Rsound
    Rsound = 0
  End If
End Sub

Sub TriggerS_Hit()
  Set SBall = Activeball
  TriggerS.TimerInterval = 100
  TriggerS.TimerEnabled = True
  PlaySound "Roll 1"
End Sub
' -----------------------


Sub GOpen_Hit:If ActiveBall.VelY<0 Then:GRest.IsDropped=1:GPush.IsDropped=0:End If:End Sub
 Sub GOpen_unHit:If ActiveBall.VelY>0 Then:GPush.IsDropped=1:GRest.IsDropped=0:End If:End Sub
 Sub GClose_Hit:GPush.IsDropped=1:GRest.IsDropped=0:End Sub

 'switches
 '1 = 60,000
 '2 = 60,000
Sub S3_Hit:dtL.Hit 1:End Sub
Sub S4_Hit:dtL.Hit 2:End Sub
Sub S5_Hit:dtL.Hit 3:End Sub
Sub S6_Hit:dtR.Hit 1:End Sub
Sub S7_Hit:dtR.Hit 2:End Sub
Sub S8_Hit:dtR.Hit 3:End Sub
 '9 = 50,000
 '10 = 50,000
 '11 = 50,000
 '12 = 300,000
 Sub S13_Slingshot:vpmTimer.PulseSw 13:End Sub
 Sub S14_Slingshot:vpmTimer.PulseSw 14:End Sub
 Sub S15_Hit:vpmTimer.PulseSw 15:End Sub
 Sub S16_Hit:vpmTimer.PulseSw 16:End Sub
 '17 = 10000
 Sub S18A_hit:vpmTimer.PulseSw 18:End Sub '10 point switch
 Sub S18B_hit:vpmTimer.PulseSw 18:End Sub '10 point switch
 Sub S18C_hit:vpmTimer.PulseSw 18:End Sub '10 point switch
 Sub S18D_hit:vpmTimer.PulseSw 18:End Sub '10 point switch
 Sub S18E_hit:vpmTimer.PulseSw 18:End Sub '10 point switch
 Sub Q19_Hit:vpmTimer.PulseSw 16:End Sub '10 point switch
 Sub Q11A_Hit:vpmTimer.PulseSw 16:End Sub'10 point switch
 '19-24=Vari-Target handled below
Sub S25_Hit:bsTrough.AddBall Me: TriggerS.TimerEnabled = False: End Sub
'26=coin1
 '27=coin2
 '28=coin3
 '29=start button
 '30=slam tilt
 '31=tilt

 '32 = 10000
 '33 = 300000
 '34 = 300000



 'vari-target
Sub TriggerVT1_Hit
  DOF 112, DOFPulse
 If ActiveBall.VelY<0 Then
    ActiveBall.VelY=ActiveBall.VelY/1.5
   Vari1.IsDropped=1
   Vari2.IsDropped=0
    Controller.Switch(24)=1
    VariTReset=1
    TriggerVT1.Enabled=0
  End If
End Sub'position 1

Sub TriggerVT2_Hit
  DOF 112, DOFPulse
 If ActiveBall.VelY<0 Then
    ActiveBall.VelY=ActiveBall.VelY/1.5
   Vari2.IsDropped=1
   Vari3.IsDropped=0
    Controller.Switch(23)=1
    Controller.Switch(24)=0
    TriggerVT2.Enabled=0
  End If
End Sub'position 2

Sub TriggerVT3_Hit
  DOF 112, DOFPulse
 If ActiveBall.VelY<0 Then
    ActiveBall.VelY=ActiveBall.VelY/1.5
   Vari3.IsDropped=1
   Vari4.IsDropped=0
    Controller.Switch(22)=1
    Controller.Switch(23)=0
    TriggerVT3.Enabled=0
  End If
End Sub'position 3

Sub TriggerVT4_Hit
  DOF 112, DOFPulse
 If ActiveBall.VelY<0 Then
    ActiveBall.VelY=ActiveBall.VelY/1.5
   Vari4.IsDropped=1
   Vari5.IsDropped=0
    Controller.Switch(21)=1
    Controller.Switch(22)=0
    TriggerVT4.Enabled=0
  End If
End Sub'position 4

Sub TriggerVT5_Hit
  DOF 112, DOFPulse
 If ActiveBall.VelY<0 Then
    ActiveBall.VelY=ActiveBall.VelY/1.5
   Vari5.IsDropped=1
   Vari6.IsDropped=0
    Controller.Switch(20)=1
    Controller.Switch(21)=0
    TriggerVT5.Enabled=0
  End If
End Sub'position 5

Sub TriggerVT6_Hit
  DOF 112, DOFPulse
  If ActiveBall.VelY<0 Then
    ActiveBall.VelY=ActiveBall.VelY/1.5
    Controller.Switch(19)=1
    Controller.Switch(20)=0
    TriggerVT6.Enabled=0
  End If
End Sub'position 6

Sub Trigger1_Hit
 If VariTReset=1 Then
  VariReset
End If
End Sub

Dim VariTStep

Sub VariReset:VariTStep=0:VTTimer.Enabled=1:VariTReset=0:End Sub
Sub VTTimer_Timer
  DOF 112, DOFPulse
  Select Case VariTStep
    Case 0:If Vari6.IsDropped=0 Then
        Controller.Switch(19)=0
        Vari6.IsDropped=1
        Vari5.IsDropped=0
        Controller.Switch(20)=1
        TriggerVT6.Enabled=1
      End If
    Case 1:If Vari5.IsDropped=0 Then
        Controller.Switch(20)=0
        Vari5.IsDropped=1
        Vari4.IsDropped=0
        Controller.Switch(21)=1
        TriggerVT5.Enabled=1
      End If
    Case 2:If Vari4.IsDropped=0 Then
        Controller.Switch(21)=0
        Vari4.IsDropped=1
        Vari3.IsDropped=0
        Controller.Switch(22)=1
        TriggerVT4.Enabled=1
      End If
    Case 3:If Vari3.IsDropped=0 Then
        Controller.Switch(22)=0
        Vari3.IsDropped=1
        Vari2.isDropped=0
        Controller.Switch(23)=1
        TriggerVT3.Enabled=1
      End If
    Case 4:If Vari2.IsDropped=0 Then
        Controller.Switch(23)=0
        Vari2.IsDropped=1
        Vari1.isDropped=0
        Controller.Switch(24)=1
        TriggerVT2.Enabled=1
      End If
    Case 5:If Vari1.IsDropped=0 Then
        Controller.Switch(24)=0
        TriggerVT1.Enabled=1
        VTTimer.Enabled=0
      End If
  End Select
  VariTStep = VariTStep + 1
End Sub

Sub Q8_Hit
  DOF 111, DOFPulse
End Sub

Sub Q32_Hit
  DOF 101, DOFPulse
End Sub

Sub Q12_Hit
  DOF 102, DOFPulse
End Sub

Sub Q24_Hit
  DOF 103, DOFPulse
End Sub

Sub Q11A_Hit:vpmTimer.PulseSw 17:End Sub
Sub Q11B_Hit:vpmTimer.PulseSw 17:End Sub
Sub Q99_Hit:vpmTimer.PulseSw 17:End Sub
Sub Q19_Hit:vpmTimer.PulseSw 17:End Sub
Sub Q26A_Hit:vpmTimer.PulseSw 17:End Sub
Sub Q26B_Hit:vpmTimer.PulseSw 17:End Sub
Sub Q28_Hit:vpmTimer.PulseSw 17:End Sub
Sub Q18_Hit:vpmTimer.PulseSw 17:End Sub
Sub Q10_Spin:vpmTimer.PulseSw 17:End Sub
Sub Q21_Spin:vpmTimer.PulseSw 17:End Sub

Sub Table1_Exit
  If B2SOn Then Controller.Stop
End Sub

 'Petaco 2
'Sub editDips
'Dim vpmDips:Set vpmDips=New cvpmDips
'With vpmDips
'.AddForm 700,600,"Petaco 2 - DIP switches"
'.AddFrame 0,0,190,"25 psts coin slot - credits",5,Array("2:1",5,"1:1",0,"1:2",4,"1:3",1)'vpm dip 1 and 3     '80
'.AddFrame 220,0,190,"100 psts coin slot - credits",80,Array("1:1",0,"1:2",64,"1:3",16,"1:10",80)'vpm dip 5 and 7  '80
'.AddFrame 0,80,190,"Balls per game",1024,Array("3 balls",0,"5 balls",1024)'vpm dip 11               '46
'.AddFrame 220,80,150,"Match Feature",32768,Array("ON",32768,"OFF",0)'vpm dip 16                   '46
'.AddLabel 106,130,270,20,"Disable score replays by checking both of these"                      '15
'.AddChk 126,145,200,Array("Ignore first replay if under 1.000.000",32)'vpm dip 6                  '15
'.AddChk 126,160,200,Array("Only allow 1 score replay per game",64)'vpm dip 7                    '15
'.AddChk 21,185,190,Array("Enable Extra Ball via score award:",8)'vpm dip 4                      '15
'.AddLabel 160,200,280,20,"600.000"
'.AddLabel 160,214,280,20,"700.000"
'.AddLabel 160,228,280,20,"900.000"
'.AddLabel 160,242,280,20,"800.000"
'.AddFrame 220,186,190,"Score Replay Levels",768,Array("800.000 - 1.000,000 - 1.200.000",0,"900.000 - 1.000,000 - 1.200.000",256,"1.000.000 - 1.500.000 - 2.000.000",768,"1.000.000 - 1.000.000 - 1.400.000",512)'vpm dip 9 and 10 '80
'.AddFrame 160,280,100,"Bookkeeping",524288,Array("bookkeeping off",0,"coin audits",262144,"play audits",524288)'vpm dip 19 and 20   '75
'.AddLabel 80,350,280,20,"After hitting OK, press F3 to reset game with new settings."
' .ViewDips
'End With
'End Sub
'Set vpmShowDips=GetRef("editDips")

Sub L09_Init()

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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.8, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.2, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus - this tables should get a total sound rewrite

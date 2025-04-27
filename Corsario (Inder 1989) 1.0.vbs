Option Explicit
Randomize
' Thalamus - seems ok

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="corsario",UseSolenoids=1,UseLamps=1,UseGI=0, UseSync=1, SCoin="coin3"

LoadVPM "01560000","inder.vbs",3.26

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp673.visible=1
Ramp674.visible=1
Primitive124.visible=1
EB1.visible=1
EB2.visible=1
TextBox1.visible=1
Else
Ramp673.visible=0
Ramp674.visible=0
Primitive124.visible=0
EB1.visible=0
EB2.visible=0
TextBox1.visible=0
End if

'*************************************************************
'Solenoid Call backs
'*************************************************************

SolCallback(3)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(4)= "bsTrough.SolOut"
'SolCallback(5)="GameStart"
SolCallback(8)="dtL.SolDropUp"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart
     End If
End Sub


'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

Dim bsTrough,dtL,bsBT

Sub Table1_Init
  DivOpen.IsDropped=1
  DivOpen1.IsDropped=1
  Primitive102.z=35
  Primitive103.z=-35
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = ""&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 1
        .Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
       Controller.SolMask(0)=0
         vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=53
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3)

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,91,0,0,0,0,0,0
  bsTrough.InitKick BallRelease,93,5
    bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls=1

  Set dtL=New cvpmDropTarget
  dtL.InitDrop Array(S51,S52,S53,S54),Array(71,72,73,74)
  dtL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
  dtL.CreateEvents "dtL"


  Set bsBT=New cvpmBallStack
  bsBT.InitSw 0,70,0,0,0,0,0,0
  bsBT.InitKick Kicker4,97,5
  bsBT.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solon",DOFContactors)
  bsBT.KickForceVar = 3
  bsBT.KickAngleVar = 2

  vpmMapLights AllLights
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If keycode = LeftMagnaSave Then bLutActive = True
  If keycode = RightMagnaSave Then
    If bLutActive Then NextLUT: End If
    End If
  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAtVol"plungerpull", Plunger, 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If keycode = LeftMagnaSave Then bLutActive = False
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'*******************
'Inder Switches +20
'*******************

'Kickers

Sub S47_Hit:S47.DestroyBall:vpmTimer.PulseSw 67:Kicker2.CreateBall:Kicker2.Kick 180,3:End Sub
Sub S50_Hit:bsBT.AddBall Me:Kicker4.Kick 97,5:End Sub


'Targets

Sub S61_Hit:vpmTimer.PulseSw 81: PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
Sub S64_Hit:vpmTimer.PulseSw 84: PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
Sub S65_Hit:vpmTimer.PulseSw 85: PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
Sub S75_Hit:vpmTimer.PulseSw 95: PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub


'Triggers

Sub S40_Hit:Controller.Switch(60)=1:PlaySoundAtVol "sensor", ActiveBall, 1: End Sub
Sub S40_unHit:Controller.Switch(60)=0:End Sub

Sub S41_Hit:Controller.Switch(61)=1:S41p.z = 33 : PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub S41_unHit:Controller.Switch(61)=0:S41p.z = 55:End Sub

Sub S42_Hit:Controller.Switch(62)=1:S42p.z = 33 : PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub S42_unHit:Controller.Switch(62)=0:S42p.z = 55:End Sub

Sub S43_Hit:Controller.Switch(63)=1:PlaySoundAtVol "sensor", ActiveBall, 1: End Sub
Sub S43_unHit:Controller.Switch(63)=0:End Sub

Sub S57_Hit:Controller.Switch(77)=1:PlaySoundAtVol "sensor", ActiveBall, 1: End Sub
Sub S57_unHit:Controller.Switch(77)=0:End Sub

Sub S62_Hit:Controller.Switch(82)=1:S62p.z = -30 : PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub S62_unHit:Controller.Switch(82)=0:S62p.z = -8:End Sub

Sub S63_Hit:Controller.Switch(83)=1:S63p.z = -30 : PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub S63_unHit:Controller.Switch(83)=0:S63p.z = -8:End Sub

Sub S67_Hit:Controller.Switch(87)=1:PlaySoundAtVol "sensor", ActiveBall, 1: End Sub
Sub S67_unHit:Controller.Switch(87)=0:End Sub

Sub S70_Hit:Controller.Switch(90)=1:PlaySoundAtVol "sensor", ActiveBall, 1: End Sub
Sub S70_unHit:Controller.Switch(90)=0: End Sub

Sub S73A_Hit:Controller.Switch(93)=1:PlaySoundAtVol "sensor", ActiveBall, 1: End Sub
Sub S73A_unHit:Controller.Switch(93)=0: End Sub

Sub S73B_Hit:Controller.Switch(93)=1:PlaySoundAtVol "sensor", ActiveBall, 1: End Sub
Sub S73B_unHit:Controller.Switch(93)=0:End Sub

Sub S74_Hit:Controller.Switch(94)=1:PlaySoundAtVol "sensor", ActiveBall, 1: End Sub
Sub S74_unHit:Controller.Switch(94)=0: End Sub

Sub S77_Hit:Controller.Switch(97)=1:PlaySoundAtVol "sensor", ActiveBall, 1: End Sub
Sub S77_unHit:Controller.Switch(97)=0:End Sub


'Spinner

Sub S44_Spin:vpmTimer.PulseSw 64:PlaySoundAtVol "spinner", S44, 1:End Sub


'Slings & Rebounds

Sub S60_Hit:vpmTimer.PulseSw 80:PlaySoundAtVol"fx_rubber", S60, 1:End Sub

Sub S72_Hit:vpmTimer.PulseSw 92:PlaySoundAtVol"fx_rubber", S72, 1:End Sub

'Door

Sub Drain_Hit
  If DivClosed.IsDropped Then
    DivOpen.IsDropped=1:PlaySoundAtVol "metalhit", Drain, 1
    DivOpen1.IsDropped=1
    DivClosed.IsDropped=0:PlaySoundAtVol "metalhit", Drain, 1
    DivClosed1.IsDropped=0
  End If
  Primitive102.z=35
  Primitive103.z=-35
  bsTrough.AddBall Me:PlaySoundAtVol "Drain", Drain, 1
End Sub


'Bumpers

Sub Bumper1_Hit: vpmTimer.PulseSw 66 : PlayBumperSound
    Light17.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
  Light17.State = 0
  Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit: vpmTimer.PulseSw 86 : PlayBumperSound
    Light18.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer
  Light18.State = 0
  Me.Timerenabled = 0
End Sub

Sub Bumper3_Hit: vpmTimer.PulseSw 96 : PlayBumperSound
    Light19.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer
  Light19.State = 0
  Me.Timerenabled = 0
End Sub

Sub PlayBumperSound()
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1
      Case 2 : PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), ActiveBall, 1
      Case 3 : PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), ActiveBall, 1
    End Select
End Sub

'Gates

Sub Gate1_Hit():PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub
Sub Gate2_Hit():PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub
Sub Gate3_Hit():PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub

'*****************
'      Trigger
'*****************

Sub Trigger2_UnHit
ActiveBall.VelX = -2
End Sub

'*****************
'Lamps & Flashers
'*****************

Set LampCallback=GetRef("UpdateMultipleLamps")

Dim Old14,New14,Old9,New9,Old15,New15,Old24,New24,Old25,New25,ST1,ST2
Old14=1:Old9=1:Old15=1:Old24=1:Old25=1:ST1=0:ST2=0

Sub UpdateMultipleLamps
  New14=Light14.State
  If New14<>Old14 Then
    If New14=0 Then
      If bsBT.Balls Then bsBT.ExitSol_On
    End If
  End If
  Old14=New14

  New9=Light9.State
  If New9<>Old9 Then
    If New9=0 Then
      DivClosed.IsDropped=1
      DivClosed1.IsDropped=1
      DivOpen.IsDropped=0
      DivOpen1.IsDropped=0
      Primitive102.z=-35
      Primitive103.z=35
    End If
  End If
  Old9=New9

  New24=A24.State
  If New24<>Old24 Then
    If New24=0 Then
      EB1.State=1
      ST1=1
    Else
      EB1.State=0
      ST1=0
    End If
  End If
  Old24=New24

  New25=A25.State
  If New25<>Old25 Then
    If New25=0 Then
      EB2.State=1
      ST2=1
    Else
      EB2.State=0
      ST2=0
    End If
  End If

    If A13.State = 0 then S41p.image = "bprim1"
  If A13.State = 1 then S41p.image = "bprim_luz1"
  If A9.State = 0 then S42p.image = "bprim1"
  If A9.State = 1 then S42p.image = "bprim_luz1"
  Old25=New25

  If ST1=0 And ST2=0 Then Textbox1.Text="0"
  If ST1=1 And ST2=0 Then Textbox1.Text="1"
  If ST1=0 And ST2=1 Then Textbox1.Text="2"
  If ST1=1 And ST2=1 Then Textbox1.Text="3"

End Sub

Sub Trigger1_Hit
  If DivClosed.IsDropped Then
    DivOpen.IsDropped=1
    DivOpen1.IsDropped=1
    DivClosed.IsDropped=0
    DivClosed1.IsDropped=0
    Primitive102.z=35
    Primitive103.z=-35
  End If
End Sub

'******
'LED's
'******

Dim Digits(32)
Digits(0)=Array(B1,B2,B3,B4,B5,B6,B7)
Digits(1)=Array(B8,B9,B10,B11,B12,B13,B14)
Digits(2)=Array(B15,B16,B17,B18,B19,B20,B21)
Digits(3)=Array(B22,B23,Light1,B25,B26,B27,B28)
Digits(4)=Array(B29,B30,B31,B32,B33,B34,B35)
Digits(5)=Array(B36,B37,B38,B39,B40,B41,B42)
Digits(6)=Array(B43,B44,B45,B46,B47,B48,B49)
Digits(7)=Array(B50,B51,B52,B53,B54,B55,B56)
Digits(8)=Array(B57,B58,B59,B60,B61,B62,B63)
Digits(9)=Array(B64,B65,B66,B67,B68,B69,B70)
Digits(10)=Array(B71,B72,B73,B74,B75,B76,B77)
Digits(11)=Array(B78,B79,B80,B81,B82,B83,B84)
Digits(12)=Array(B85,B86,B87,B88,B89,B90,B91)
Digits(13)=Array(B92,B93,B94,B95,B96,B97,B98)
Digits(14)=Array(B99,B100,B101,B102,B103,B104,B105)
Digits(15)=Array(B106,B107,B108,B109,B110,B111,B112)
Digits(16)=Array(B113,B114,B115,B116,B117,B118,B119)
Digits(17)=Array(B120,B121,B122,B123,B124,B125,B126)
Digits(18)=Array(B127,B128,B129,B130,B131,B132,B133)
Digits(19)=Array(B134,B135,B136,B137,B138,B139,B140)
Digits(20)=Array(B141,B142,B143,B144,B145,B146,B147)
Digits(21)=Array(B148,B149,B150,B151,B152,B153,B154)
Digits(22)=Array(B155,B156,B157,B158,B159,B160,B161)
Digits(23)=Array(B162,B163,B164,B165,B166,B167,B168)
Digits(24)=Array(B169,B170,B171,B172,B173,B174,B175)
Digits(25)=Array(B176,B177,B178,B179,B180,B181,B182)
Digits(26)=Array(B183,B184,B185,B186,B187,B188,B189)
Digits(27)=Array(B190,B191,B192,B193,B194,B195,B196)
Digits(28)=Array(B197,B198,B199,B200,B201,B202,B203)
Digits(29)=Array(B204,B205,B206,B207,B208,B209,B210)
Digits(30)=Array(B211,B212,B213,B214,B215,B216,B217)
Digits(31)=Array(B218,B219,B220,B221,B222,B223,B224)


Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 32) then
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

'*********************
'   Inder "Corsario"
 'Added By Inkochnito
'*********************

 Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,400,"Corsario - DIP switches"
    .AddFrame 0,0,340,"Credits per coin",&H00000003,Array("1 coin - 1 credit && 1 coin - 4 credits",0,"2 coins - 1 credit (4 coins - 3 credits) && 1 coin - 3 credits",&H00000003)'SL1-3&4 (dip 1&2)
    .AddFrame 0,46,100,"Handicap value",&H00000300,Array("4,500,000 points",0,"5,000,000 points",&H00000100,"5,500,000 points",&H00000200,"6,000,000 points",&H00000300)'SL2-4&3(dip 9&10)
    .AddFrame 240,46,100,"Balls per game",&H00000008,Array("3 balls",0,"5 balls",&H00000008)'SL1-1 (dip 4)
    .AddFrame 120,46,100,"Replay threshold",&H000000C0,Array("3,000,000 points",0,"3,500,000 points",&H00000080,"4,000,000 points",&H00000040,"4,500,000 points",&H000000C0)'SL1-6&5 (dip 7&8)
    .AddLabel 50,130,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
 End Sub
 Set vpmShowDips=GetRef("editDips")


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
        StopSound "fx_ballrolling" & b
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
    Next

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
    If BallVel(BOT(b) ) > 1 Then
            rolling(b) = True
      If BOT(b).Z > 30 Then
        ' ball on plastic ramp
        StopSound "fx_ballrolling" & b
        StopSound "fx_metalrolling" & b
        PlaySound "fx_plasticrolling" & b, -1, Vol(BOT(b)) / 2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
      Else
        ' ball on playfield
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
        PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b)) / 2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
      End If
    Else
      If rolling(b) Then
                StopSound "fx_ballrolling" & b
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
                rolling(b) = False
            End If
    End If

    '   ball drop sounds matching the adjusted height params but not the way down the ramps
    If BOT(b).VelZ < -1 And BOT(b).Z < 55 And BOT(b).Z > 27 then 'And Not InRect(BOT(b).X, BOT(b).Y, 610,320, 740,320, 740,550, 610,550) And Not InRect(BOT(b).X, BOT(b).Y, 180,400, 230,400, 230, 550, 180,550) Then
      PlaySound "fx_ball_drop" & Int(Rnd()*3), 0, ABS(BOT(b).VelZ)/17*5, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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

  LFPrim.roty=leftFlipper.CurrentAngle
  RFPrim.roty=rightFlipper.CurrentAngle

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
  PlaySound "Gate5", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

' Ramps sounds

Sub RampSound1_Hit: PlaySoundAtVol "fx_metalrolling1", ActiveBall, 1: End Sub
Sub RampSound3_Hit: PlaySoundAtVol "fx_metalrolling1", ActiveBall, 1: End Sub

' Stop Ramps Sounds

Sub RampSound2_Hit: StopSound "fx_metalrolling1": End Sub
Sub RampSound4_Hit: StopSound "fx_metalrolling1": End Sub


Sub Table1_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub

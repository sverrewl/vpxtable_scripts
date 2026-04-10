Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="mundial",UseSolenoids=2,UseLamps=1,UseGI=0,SCoin="coin3"

LoadVPM "01560000","inder.vbs",3.26

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
'Ramp16.visible=1
'Ramp15.visible=1
'Primitive13.visible=1
Else
'Ramp16.visible=0
'Ramp15.visible=0
'Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
'1 Start Button
'2 COIN SLOTS Lights
SolCallback(3)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(4)="bsTrough.SolOut"
SolCallback(5)="GameStart"
'SolCallback(6)="vpmSolSound ""bumper"","
'SolCallback(7)="vpmSolSound ""bumper"","
SolCallback(8)="dtDrop.SolDropUp"
SolCallback(9)="SolKickback"
SolCallback(10)="bsSaucer.SolOut"
'SolCallback(11)="vpmSolSound ""Popper"","  'Kicking Target

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolKickback(Enabled)
     If Enabled Then
        PlaySoundAtVol SoundFX("Popper",DOFContactors), Kickback, 1:
    kickback.fire
     Else
    kickback.pullback
     End If
End Sub

'Playfield GI
'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

Sub GameStart(Enabled)
    vpmNudge.SolGameOn Enabled
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucer, dtDrop
Dim Captive

Sub Table1_Init
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
    .hidden = 0
        .Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         Controller.SolMask(0)=0
         vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=53
  vpmNudge.Sensitivity=2
  vpmNudge.TiltObj=Array(S46,Bumper1,Bumper2)

  Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,91,0,0,0,0,0,0
    bsTrough.InitKick BallRelease,90,5
    bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls=1

  Set bsSaucer=New cvpmBallStack
    bsSaucer.InitSaucer sw70,70,0,40
    bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsSaucer.KickZ=0.785398163 '45 degrees
    bsSaucer.KickForceVar = 3
    bsSaucer.KickAngleVar = 3

  Set dtDrop=New cvpmDropTarget
    dtDrop.InitDrop Array(S51,S52,S53,S54,S55),Array(71,72,73,74,75)
    dtDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  Set Captive=New cvpmCaptiveBall
    Captive.InitCaptive CapTrigger,CaptiveWall,CaptiveKicker,340
    Captive.Start
    Captive.ForceTrans=2
    Captive.MinForce=3.5
    Captive.CreateEvents"Captive"

  vpmMapLights AllLights

    'Mini PF Legs
    Dim obj
    For Each obj In ColLegs:obj.IsDropped = 1:Next
    For Each obj In ColLegs2:obj.IsDropped = 1:Next
    For Each obj In ColLegs3:obj.IsDropped = 1:Next

  'Kicking Target
  Sling.IsDropped = 1:Sling2.IsDropped = 1:Sling3.IsDropped = 1

  'Diverter
  Lam1.IsDropped=0:Lam2.IsDropped=1

 End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol "plungerpull", Plunger, 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)

  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger", Plunger, 1
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'**********************************************************************************************************

'Drain and Kickers
Sub Drain_Hit:bsTrough.AddBall Me:End Sub
Sub sw70_Hit:bsSaucer.AddBall 0:PlaySoundAtVol"drain", Drain, 1:End Sub

'Mini PF holes
Sub S40_Hit:S40.DestroyBall:vpmTimer.PulseSw 60:ToyExit: PlaySoundAtVol "hole_enter", ActiveBall, 1:End Sub
Sub S41_Hit:S41.DestroyBall:vpmTimer.PulseSw 61:ToyExit: PlaySoundAtVol "hole_enter", ActiveBall, 1:End Sub
Sub S42_Hit:S42.DestroyBall:vpmTimer.PulseSw 62:ToyExit: PlaySoundAtVol "hole_enter", ActiveBall, 1:End Sub
Sub S43_Hit:S43.DestroyBall:vpmTimer.PulseSw 63:ToyExit: PlaySoundAtVol "hole_enter", ActiveBall, 1:End Sub
Sub S47_Hit:S47.DestroyBall:vpmTimer.PulseSw 67:ToyExit: PlaySoundAtVol "hole_enter", ActiveBall, 1:End Sub
Sub ToyExit:Kicker1.CreateBall:Kicker1.Kick 125,1:End Sub

'Targets
Sub S60_Hit:vpmTimer.PulseSw 80 : PlaySoundAtVol SoundFX("target",DOFTargets) , ActiveBall, 1: End Sub
Sub S61_Hit:vpmTimer.PulseSw 81 : PlaySoundAtVol SoundFX("target",DOFTargets) , ActiveBall, 1: End Sub
Sub S62_Hit:vpmTimer.PulseSw 82 : PlaySoundAtVol SoundFX("target",DOFTargets) , ActiveBall, 1: End Sub
Sub S63_Hit:vpmTimer.PulseSw 83 : PlaySoundAtVol SoundFX("target",DOFTargets) , ActiveBall, 1: End Sub
Sub S64_Hit:vpmTimer.PulseSw 84 : PlaySoundAtVol SoundFX("target",DOFTargets) , ActiveBall, 1: End Sub
Sub S65_Hit:vpmTimer.PulseSw 85 : PlaySoundAtVol SoundFX("target",DOFTargets) , ActiveBall, 1: End Sub
Sub S75_Hit:vpmTimer.PulseSw 95 : PlaySoundAtVol SoundFX("target",DOFTargets) , ActiveBall, 1: End Sub

'DropTargets
Sub S51_Hit:dtDrop.Hit 1:End Sub
Sub S52_Hit:dtDrop.Hit 2:End Sub
Sub S53_Hit:dtDrop.Hit 3:End Sub
Sub S54_Hit:dtDrop.Hit 4:End Sub
Sub S55_Hit:dtDrop.Hit 5:End Sub

'Kicking Target
Dim S46Step
Sub S46_Slingshot:Sling.IsDropped = 0:
    PlaySoundAtVol SoundFX("Popper",DOFContactors), ActiveBall, 1
  vpmTimer.PulseSw 66:
  S46Step = 0:
  Me.TimerEnabled = 1:
End Sub

Sub S46_Timer
  Select Case S46Step
    Case 0:SLing.IsDropped = 0
    Case 1: 'pause
    Case 2:SLing.IsDropped = 1:SLing2.IsDropped = 0
    Case 3:SLing2.IsDropped = 1:SLing3.IsDropped = 0
    Case 4:SLing3.IsDropped = 1:Me.TimerEnabled = 0
  End Select
  S46Step = S46Step + 1
End Sub

'Scoring Rubbers
Sub S72A_Hit:vpmTimer.PulseSw 92 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub
Sub S72B_Hit:vpmTimer.PulseSw 92 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub

'Wire Triggers
Sub S32_Hit:Controller.Switch(52)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: End Sub
Sub S32_unHit:Controller.Switch(52)=0:End Sub

Sub S45_Hit:Controller.Switch(65)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: End Sub
Sub S45_unHit:Controller.Switch(65)=0:End Sub

Sub S57_Hit:Controller.Switch(77)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: End Sub
Sub S57_unHit:Controller.Switch(77)=0:End Sub

Sub S67_Hit:Controller.Switch(87)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: End Sub
Sub S67_unHit:Controller.Switch(87)=0: End Sub

Sub S77_Hit:Controller.Switch(97)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: End Sub
Sub S77_unHit:Controller.Switch(97)=0:End Sub

Sub S70_Hit:Controller.Switch(90)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: End Sub
Sub S70_unHit:Controller.Switch(90)=0:End Sub

Sub S73A_Hit:Controller.Switch(93)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: End Sub
Sub S73A_unHit:Controller.Switch(93)=0:End Sub

Sub S73B_Hit:Controller.Switch(93)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: End Sub
Sub S73B_unHit:Controller.Switch(93)=0:End Sub

Sub S74A_Hit:Controller.Switch(94)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: DOF 101, DOFOn : End Sub
Sub S74A_unHit:Controller.Switch(94)=0: DOF 101, DOFOff :End Sub

Sub S74B_Hit:Controller.Switch(94)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: DOF 102, DOFOn : End Sub
Sub S74B_unHit:Controller.Switch(94)=0: DOF 102, DOFOff :End Sub

Sub S74C_Hit:Controller.Switch(93)=1:PlaySoundAtVol "sensor" , ActiveBall, 1: DOF 103, DOFOn : End Sub
Sub S74C_unHit:Controller.Switch(93)=0: DOF 103, DOFOff : End Sub

'Spinner
Sub S44_Spin:vpmTimer.PulseSw 64 : playsoundAtVol"fx_spinner" , s44, 1: End Sub

'Bumpers
Sub Bumper1_Hit:
  vpmTimer.PulseSw 86
    PlayBumperSound
  BL1.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
  BL1.State = 0
  Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit:
  vpmTimer.PulseSw 96
    PlayBumperSound
  BL2.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer
  BL2.State = 0
  Me.Timerenabled = 0
End Sub

Sub PlayBumperSound()
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtVol SoundFX("bumper",DOFContactors), ActiveBall, 1
      Case 2 : PlaySoundAtVol SoundFX("bumper",DOFContactors), ActiveBall, 1
      Case 3 : PlaySoundAtVol SoundFX("bumper",DOFContactors), ActiveBall, 1
    End Select
End Sub

'****************
'Diverter
'****************

Sub TopeBola_Hit:MueveLam.enabled=1:End Sub

Sub MueveLam_Hit
  PlaySoundAtVol "metalhit", ActiveBall, 1
  Lam1.IsDropped=1:Lam2.IsDropped=0
End Sub

Sub MueveLam_unHit
    PlaySoundAtVol "metalhit", ActiveBall, 1
  Lam1.IsDropped=0:Lam2.IsDropped=1
  MueveLam.enabled=0
End Sub

'***************
'Lamp Simulator
'***************
Dim lampPosition, lampLastPos

Sub SpinTimer_Timer
    Dim curPos
    lampPosition = Primitive_Wheel.RotY
    Do While lampPosition < 0
        lampPosition = lampPosition + 360
    Loop
    Do While lampPosition > 360
        lampPosition = lampPosition - 360
    Loop
    curPos = Int((lampPosition * ColLegs.Count) / 360)
    If curPos <> lampLastPos Then
        If lampLastPos >= 0 Then ' not first time
            ColLegs(lampLastPos).IsDropped = True
            ColLegs2(lampLastPos).IsDropped = True
            ColLegs3(lampLastPos).IsDropped = True
        End If
        On Error Resume Next
        ColLegs(curPos).IsDropped = False
        If Err Then msgbox curPoles
        ColLegs2(curPos).IsDropped = False
        If Err Then msgbox curPoles
        ColLegs3(curPos).IsDropped = False

        lampLastPos = curPos
    End If
End Sub

'*****************
'Lamps & Flashers
'*****************

Set LampCallback=GetRef("UpdateMultipleLamps")
Sub UpdateMultipleLamps

    GiraIzquierda.Enabled = l26.State
    GiraDerecha.Enabled = l27.State
End Sub

Sub GiraIzquierda_Timer()
Primitive_Wheel.RotY = (Primitive_Wheel.RotY -5) MOD 360
End Sub

Sub GiraDerecha_Timer()
Primitive_Wheel.RotY = (Primitive_Wheel.RotY +5) MOD 360
End Sub

'Generic Sounds
Sub S76_Hit:PlaySoundAtVol "hole_enter", ActiveBall, 1:End Sub
Sub S78_Hit:PlaySoundAtVol "metalrolling", ActiveBall, 1:End Sub

Sub Gate3_Hit():PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate5_Hit():PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate6_Hit():PlaySoundAtVol "gate", ActiveBall, 1:End Sub


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(32)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

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


'**********************************************************************************************************
'**********************************************************************************************************




'********************
' Inder Mundial 90
'Added By Inkochnito
'********************

Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,400,"250CC - DIP switches"
    .AddFrame 2,0,392,"Credits per coin",&H00000003,Array("1 coin - 1 credit && 1 coin - 4 credits",0,"2 coins - 1 credit (4 coins - 3 credits) && 1 coin - 3 credits",&H00000003)'SL1-3&4 (dip 1&2)
    .AddFrame 2,92,190,"Handicap value",&H00000300,Array("5,000,000 points",0,"5,200,000 points",&H00000100,"5,400,000 points",&H00000200,"5,600,000 points",&H00000300)'SL2-4&3(dip 9&10)
    .AddFrame 205,46,190,"Balls per game",&H00000008,Array("3 balls",0,"5 balls",&H00000008)'SL1-1 (dip 4)
    .AddFrame 205,92,190,"Replay threshold",&H000000C0,Array("3,000,000 points",0,"3,300,000 points",&H00000080,"3,500,000 points",&H00000040,"3,800,000 points",&H000000C0)'SL1-6&5 (dip 7&8)
    .AddFrame 2,46,190,"Allow extra ball",&H00400000,Array("yes",0,"no",&H00400000)'SL3-6 (dip 23)
    .AddLabel 50,180,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")

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

    LFPrim.RotY = LeftFlipper.CurrentAngle
  RFPrim.RotY = RightFlipper.CurrentAngle

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
'    If UBound(BOT) = -1 Then Exit Sub
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
  PlaySound "gate", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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



'************
 'Lamps
'Manual/VPM
'************

'1 = solenoid 1 = LUZ Pulsador Partidas (En Puerta) - START BUTTON
'2 = solenoid 2 = LUZ Inhibicion Monedero           - COIN SLOTS
'3 = 1 = LUZ Especial Derecha
'4 = 2 = LUZ Especial Diana Centro
'5 = 3 = LUZ Bola Extra Diana Cautiva
'6 = 4 = LUZ Bola Extra Inferior
'7 = 6 = LUZ Loteria (En Cabeza)
'8 = 7 = LUZ Final Partida (En Cabeza)
'9 = 8 = LUZ Corner
'10 = 10 = LUZ Bola Extra Diana Posterior Bancada
'11 = 12 = LUZ Bola Extra Izquierda Plataforma
'12 = 17 = LUZ Bonus X3 Superior
'13 = 18 = LUZ Especial Superior
'14 = 19 = LUZ Bola Extra Superior
'15 = 20 = LUZ Bonus X5 Superior
'16 = 24 = LUZ 500,000 Plataforma
'17 = 22 = LUZ Bola Extra Derecha Plataforma
'18 = 23 = LUZ Especial Plataforma
'19 = 21 = LUZ Sube Bonus Plataforma
'20 = 49 = LUZ Bonus X3 Inferior
'21 = 50 = LUZ Bonus X5 Inferior
'22 = 29 = LUZ Veleta
'23 = 30 = LUZ Diana 100,000 Puntos
'24 = 31 = LUZ Bola En Juego (En Cabeza)
'25 = 32 = LUZ Handicap (En Cabeza)
'26 = 33 = LUZ Gira Motor
'27 = 34 = LUZ Gira Motor
'28 = 36 (used for extra ball display) LUZ Bola Extra 1 (En Cabeza)
'29 = 37 (used for extra ball display) LUZ Bola Extra 2 (En Cabeza)
'30 = 41 = LUZ Diana 1
'31 = 42 = LUZ Diana 2
'32 = 43 = LUZ Diana 3
'33 = 44 = LUZ Diana 4
'34 = 45 = LUZ Diana 5
'35 = 47 = LUZ Diana Corner Inferior
'36 = 48 = LUZ Diana Corner Superior
'37 = 51 = LUZ Bonus 1
'38 = 52 = LUZ Bonus 2
'39 = 53 = LUZ Bonus 3
'40 = 54 = LUZ Bonus 4
'41 = 55 = LUZ Bonus 5
'42 = 56 = LUZ Bonus 6
'43 = 25 = LUZ Bonus 7
'44 = 26 = LUZ Bonus 8
'45 = 27 = LUZ Bonus 9
'46 = 28 = LUZ Bonus 10

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

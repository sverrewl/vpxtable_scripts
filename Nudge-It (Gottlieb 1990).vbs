'*************************************************************************
'*************************Nudge-It  Gotlieb 1990**************************
'*************************************************************************
'
'Version 1.0
'Table by Tom
'Graphics and Modeling by Mlager8
'Scripting and SSF by Thalamus



option explicit
Randomize

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM  "01560000", "gts3.VBS", 3.26

'********************
'Standard definitions
'********************

 Const UseSolenoids = 2
 Const UseLamps = 0
 Const UseSync = 0
 Const swStartButton = 3
'Standard Sounds
 Const SSolenoidOn = "Solenoid"
 Const SSolenoidOff = ""
 Const SCoin = "fx_coin"




'***************************
'Table Init
'***************************

Const cGameName="nudgeit"

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

  With Controller
    .GameName = cGameName
    .SplashInfoLine = "Nudge-It (Gottlieb 1994)"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = 0 'DesktopMode
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With


  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  vpmNudge.TiltSwitch = 151
  vpmNudge.Sensitivity = 3


  Kicker.CreateBall
  Kicker.Kick 90,1

End Sub


'***************************
'Keys
'***************************
Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=LeftFlipperKey Then Controller.Switch(4)=1
  If KeyCode=RightFlipperKey Then Controller.Switch(5)=1
  If KeyCode=PlungerKey OR KeyCode = LockbarKey Then Controller.Switch(10)=1
  If KeyCode=PlungerKey Then Plunger.PullBack
  If keycode = LeftMagnaSave Then bLutActive = True
  If keycode = RightMagnaSave Then
    If bLutActive Then NextLUT: End If
  End If
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=LeftFlipperKey Then Controller.Switch(4)=0
  If KeyCode=RightFlipperKey Then Controller.Switch(5)=0
  If KeyCode=PlungerKey OR KeyCode = LockbarKey Then Controller.Switch(10)=0
  If KeyCode=PlungerKey Then Plunger.Fire:PlaySoundAtVol "Plunger", Plunger, 1
  If keycode = LeftMagnaSave Then bLutActive = False
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'****************************
'Solenoids
'****************************

'SolCallback(17)="flasher17"
'SolCallback(18)="vpmFlasher Light672,"
'SolCallback(19)="vpmFlasher Light673,"
'SolCallback(20)="vpmFlasher Light674,"
'SolCallback(21)="vpmFlasher Light675,"
'SolCallback(22)=
SolCallback(23)="vpmSolSound ""bell"","
'SolCallback(24)=""                 'coin meter
'SolCallback(26)="vpmFlasher BGFlasher,"  Mars Light Beacon  plus shaker effect
'SolCallback(27)=""   ball gate
SolCallback(28)="solpin"
SolCallback(30)="SolKnocker"
SolCallback(31)="vpmNudge.SolGameOn"        'Tilt Relay (T)

'****************************
'Switches
'****************************

Sub sw1_Hit():PlaySoundAtVol "rollover", sw1, .5:Controller.Switch(11) = 1:End Sub
Sub sw1_UnHit():Controller.Switch(11) = 0:End Sub
Sub sw2_Hit():PlaySoundAtVol "rollover", sw2, .5:Controller.Switch(12) = 1:End Sub
Sub sw2_UnHit():Controller.Switch(12) = 0:End Sub
Sub sw3_Hit():PlaySoundAtVol "rollover", sw3, .5:Controller.Switch(13) = 1:End Sub
Sub sw3_UnHit():Controller.Switch(13) = 0:End Sub
Sub sw4_Hit():PlaySoundAtVol "rollover", sw4, .5:Controller.Switch(14) = 1:End Sub
Sub sw4_UnHit():Controller.Switch(14) = 0:End Sub
Sub sw5_Hit():PlaySoundAtVol "rollover", sw5, .5:Controller.Switch(15) = 1:End Sub
Sub sw5_UnHit():Controller.Switch(15) = 0:End Sub
Sub sw6_Hit():PlaySoundAtVol "rollover", sw6, .5:Controller.Switch(16) = 1:End Sub
Sub sw6_UnHit():Controller.Switch(16) = 0:End Sub
Sub sw7_Hit:Controller.Switch(7) = 1:End Sub
Sub sw17_Hit():PlaySoundAtVol "rollover", sw17, .5:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit():Controller.Switch(17) = 0:End Sub

'******************************
'Gate
'******************************

dim foon, fo

Sub Solpin(Enabled)
  If Enabled Then
    pin.transy = -45
    pin.Collidable = 0
    PlaySoundAtVol "metalhit2", pin, 1
  Else
    pin.transy = 0
    pin.Collidable = 1
  End If
End Sub


Sub GateTimer_Timer()
  Gate1P.RotZ = ABS(Gate1.currentangle)
End Sub

''******************************
''KNOCKER
''******************************
'
Sub SolKnocker(enabled)
  If enabled Then
    PlaySound SoundFX("knocker",DOFKnocker)
  End If
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
If LampState(1) = 0 Then
    LockdownBarbutton.Image = "LBD_Baked_Off"
    lockdownbarbutton.blenddisablelighting=.2
  End If

  If LampState(1) = 1 Then
    LockdownBarbutton.Image = "LBD_Baked_On"
    lockdownbarbutton.blenddisablelighting=.75
  End If
  If Not IsEmpty(chgLamp) Then
    For ii = 0 To UBound(chgLamp)
      LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
      FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
    Next
  End If
  UpdateLamps
End Sub

Sub InitLamps()
  Dim x
  For x = 0 to 200
    LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
    FadingLevel(x) = 4       ' used to track the fading state
    FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
    FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
    FlashMax(x) = 1          ' the maximum value when on, usually 1
    FlashMin(x) = 0          ' the minimum value when off, usually 0
    FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
  Next
End Sub

Sub UpdateLamps
  NFadeL 1, L1
  NFadeL 10, L10
  NFadeL 11, L11
  NFadeL 12, L12
  NFadeL 13, L13
  NFadeL 14, L14
  FlupperFlash 16, Flasherflash1, Flasherlit1, Flasherbase1, Flasherlight1
  FlupperFlash 17, Flasherflash5, Flasherlit5, Flasherbase5, Flasherlight5
  NFadeL 20, L20
  NFadeL 21, L21
  NFadeL 22, L22
  NFadeL 23, L23
  NFadeL 24, L24
  FlupperFlash 26, Flasherflash2, Flasherlit2, Flasherbase2, Flasherlight2
  FlupperFlash 27, Flasherflash6, Flasherlit6, Flasherbase6, Flasherlight6
  NFadeL 30, L30
  NFadeL 31, L31
  NFadeL 32, L32
  NFadeL 33, L33
  NFadeL 34, L34
  FlupperFlash 36, Flasherflash3, Flasherlit3, Flasherbase3, Flasherlight3
  FlupperFlash 37, Flasherflash7, Flasherlit7, Flasherbase7, Flasherlight7
  NFadeL 40, L40
  NFadeL 41, L41
  NFadeL 42, L42
  NFadeL 43, L43
  NFadeL 44, L44
  FlupperFlash 46, Flasherflash4, Flasherlit4, Flasherbase4, Flasherlight4
  FlupperFlash 47, Flasherflash8, Flasherlit8, Flasherbase8, Flasherlight8
  NFadeL 50, L50
  NFadeL 51, L51
  NFadeL 52, L52
  NFadeL 53, L53
  NFadeL 54, L54
  'FlupperFlash 76, Flasherflash1, Flasherlit1, Flasherbase1, Flasherlight1

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


'If LampState(1) = 1 Then
'   LockdownBar.Image = "LBD_Baked_On"
' Else
'   LockdownBar.Image = "LBD_Baked_Off"
' End If


'********************************
'FLUPPER FLASH
'********************************

Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4, FlashLevel5, FlashLevel6, FlasherLevel7, FlasherLevel8, FlasherLevel9, FlasherLevel10
FlasherLight1.IntensityScale = 0
Flasherlight2.IntensityScale = 0
Flasherlight3.IntensityScale = 0
Flasherlight4.IntensityScale = 0
Flasherlight5.IntensityScale = 0
Flasherlight6.IntensityScale = 0
Flasherlight7.IntensityScale = 0
Flasherlight8.IntensityScale = 0


Sub FlupperFlash(nr, FlashObject, LitObject, BaseObject, LightObject)
  FadeEmpty nr
  FlupperFlashm nr, FlashObject, LitObject, BaseObject, LightObject
End Sub

Sub FlupperFlashm(nr, FlashObject, LitObject, BaseObject, LightObject)
  'exit sub
  dim flashx3
  Select Case FadingLevel(nr)
    Case 4, 5
      ' This section adapted from Flupper's script
      flashx3 = FlashLevel(nr) * FlashLevel(nr) * FlashLevel(nr)
      FlashObject.IntensityScale = flashx3
      LitObject.BlendDisableLighting = 10 * flashx3
      BaseObject.BlendDisableLighting = (flashx3 * .6) + .4
      LightObject.IntensityScale = flashx3
      LitObject.material = "domelit" & Round(9 * FlashLevel(nr))
      LitObject.visible = 1
      FlashObject.visible = 1
    case 3:
      LitObject.visible = 0
      FlashObject.visible = 0
  end select
End Sub

Sub FadeEmpty(nr) 'Fade a lamp number, no object updates
  Select Case FadingLevel(nr)
    Case 3
      FadingLevel(nr) = 0
    Case 4 'off
      FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
      If FlashLevel(nr) < FlashMin(nr) Then
        FlashLevel(nr) = FlashMin(nr)
        FadingLevel(nr) = 3 'completely off
      End if
      'Object.IntensityScale = FlashLevel(nr)
    Case 5 ' on
      FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
      If FlashLevel(nr) > FlashMax(nr) Then
        FlashLevel(nr) = FlashMax(nr)
        FadingLevel(nr) = 6 'completely on
      End if
      'Object.IntensityScale = FlashLevel(nr)
    Case 6
      FadingLevel(nr) = 1
  End Select
End Sub


'**************************************
' Flasher objects
'**************************************

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
  Vol = Csng(BallVel(ball) ^2 / 1000)
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
' Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3)

Sub BallShadowUpdate_Timer()
  Dim BOT, b, shadowZ
  BOT = GetBalls

  ' render the shadow for each ball
  For b = 0 to UBound(BOT)
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
    End If

    If BOT(b).Z < 10 and BOT(b).Z>90 then BallShadow(b).Z = 1 else BallShadow(b).Z=BOT(b).z-20

    BallShadow(b).Y = BOT(b).Y + 12
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
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

'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

Const tnob = 50 ' total number of balls
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
    ' If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
    '   PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    ' End If
  Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 1000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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


Sub GateTimer_Init()

End Sub

Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

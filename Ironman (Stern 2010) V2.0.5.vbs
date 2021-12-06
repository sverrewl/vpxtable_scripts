Option explicit
Randomize

'************************************************************************
'             TABLE OPTIONS
'************************************************************************


' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Added InitVpmFFlipsSAM
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const Ballsize = 51
Const BallMass = 1.5

Const UseVPMModSol = True
Dim UseVPMDMD,CustomDMD,DesktopMode
DesktopMode = Table.ShowDT : CustomDMD = False
If Right(cGamename,1)="c" Then CustomDMD=True
If CustomDMD OR (NOT DesktopMode AND NOT CustomDMD) Then UseVPMDMD = False    'hides the internal VPMDMD when using the color ROM or when table is in Full Screen and color ROM is not in use
If DesktopMode AND NOT CustomDMD Then UseVPMDMD = True              'shows the internal VPMDMD when in desktop mode and color ROM is not in use
Scoretext.visible = NOT CustomDMD                       'hides the textbox when using the color ROM

LoadVPM "02800000", "Sam.VBS", 3.54


'********************
'Standard definitions
'********************

  Const cGameName = "im_183ve"

     Const UseSolenoids = 1
     Const UseLamps = 0
     Const UseSync = 0
     Const HandleMech = 1
     Const SSolenoidOn = "Solenoid"
     Const SSolenoidOff = ""
     Const SCoin = "CoinIn"

 '************
' Table init.
'***********

    Dim xx
    Dim Bump1, Bump2, Bump3, Mech3bank,bsTrough,bsRHole,DTBank4,turntable,Mag1,Mag2
  Dim PlungerIM

  Sub Table_Init
  vpmInit Me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "IronMan (Stern 2010)"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    InitVpmFFlipsSAM
        If NOT CustomDMD Then .Hidden = DesktopMode       'hides the external DMD when in desktop mode and color ROM is not in use
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With

    On Error Goto 0

'Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd "ballrelease", "Solenoid"
    bsTrough.Balls = 4

  Set mag1= New cvpmMagnet
  With mag1
    .InitMagnet Magnet1, 22
    .GrabCenter = False
    .solenoid=3
    .CreateEvents "mag1"
  End With

  Set mag2= New cvpmMagnet
  With mag2
    .InitMagnet Magnet2, 22
    .GrabCenter = False
    .solenoid=4
    .CreateEvents "mag2"
  End With


'Nudging
      vpmNudge.TiltSwitch=-7
      vpmNudge.Sensitivity=1
      vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,LeftSlingshot,RightSlingshot)

'DropTargets

      '**Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  'StandUp Init
  ResetAll
  GIOff

  End Sub

 Sub Table_Paused:Controller.Pause = 1:End Sub
 Sub Table_unPaused:Controller.Pause = 0:End Sub



'*****Keys
 Sub Table_KeyDown(ByVal keycode)

  If Keycode = LeftFlipperKey then
  End If
  If Keycode = RightFlipperKey then
  End If
  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAtVol "plungerpull",Plunger,.4
'   If keycode = LeftTiltKey Then LeftNudge 80, 1, 20
'    If keycode = RightTiltKey Then RightNudge 280, 1, 20
'    If keycode = CenterTiltKey Then CenterNudge 0, 1, 25
    If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub Table_KeyUp(ByVal keycode)
  If vpmKeyUp(keycode) Then Exit Sub
  If Keycode = LeftFlipperKey then
    SolLFlipper false
  End If
  If Keycode = RightFlipperKey then
    SolRFlipper False
  End If
  If Keycode = StartGameKey Then Controller.Switch(16) = 0
    If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAtVol "Plunger",Plunger,.4
  End If
End Sub


   'Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
'SolCallback(3) = "MongerMagnet"
'SolCallback(4) = whiplash magnet
SolCallback(5) = "WMKick"
SolCallback(6) = "orbitpost"
SolCallback(12) = "ClanePost"
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
SolCallback(19) = "Solmonger"
SolModCallBack(20) = "SetLampMod 120,"
SolModCallback(21) = "SetLampMod 121,"
SolModCallback(22) = "SetLampMod 122,"
SolModCallback(23) = "SetLampMod 123,"
SolModCallback(24) = "SetLampMod 124,"
SolModCallback(25) = "SetLampMod 125,"
SolModCallback(26) = "SetLampMod 126,"
SolModCallback(27) = "SetLampMod 127,"
SolModCallback(29) = "SetLampMod 129,"
SolModCallback(30) = "SetLampMod 130,"
SolModCallback(31) = "SetLampMod 131,"
SolModCallback(32) = "SetLampMod 132,"

' Modulated Flasher and Lights objects

Sub SetLampMod(nr, value)
  Debug.Print nr &", "& value
    If value > 0 Then
    LampState(nr) = 1
  Else
    LampState(nr) = 0
  End If
  FadingLevel(nr) = value
End Sub

Sub LampMod(nr, object)
  If TypeName(object) = "Light" Then
    Object.IntensityScale = FadingLevel(nr)/128
    Object.State = LampState(nr)
  End If
  If TypeName(object) = "Flasher" Then
    Object.IntensityScale = FadingLevel(nr)/128
    Object.visible = LampState(nr)
  End If
  If TypeName(object) = "Primitive" Then
    Object.DisableLighting = LampState(nr)
  End If
End Sub


Sub solTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  End If
 End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
 End Sub

Sub ClaneUpPost_Hit() PlaySoundAt "metalhit_medium", ActiveBall: End Sub

Sub ClanePost(Enabled)
  If Enabled Then
    ClaneUpPost.Isdropped=false
  Else
    ClaneUpPost.Isdropped=true
  End If
 End Sub

Sub UpPost_Hit() PlaySoundAt "metalhit_medium", ActiveBall: End Sub

Sub orbitpost(Enabled)
  If Enabled Then
    UpPost.Isdropped=false
  Else
    UpPost.Isdropped=true
  End If
 End Sub

Sub ClanePost(Enabled)
  If Enabled Then
    ClaneUpPost.Isdropped=false
  Else
    ClaneUpPost.Isdropped=true
  End If
 End Sub

Sub WMKick(enabled)
    If enabled Then
     PlaySoundAt "ballhit", sw10
       sw10.Kick 180, 30
       controller.switch(10) = false
    End If
End Sub

'***********************************************
   'Flipper Subs



Sub SolLFlipper(Enabled)
     If Enabled Then
     PlaySoundAt SoundFX("FlipperUpLeft", DOFFlippers),GILight3
     LeftFlipper.RotateToEnd
     Else
     PlaySoundAt SoundFX("FlipperDown", DOFFlippers),GILight3
     LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
     PlaySoundAt SoundFX("FlipperUpRight", DOFFlippers), GILight4
     RightFlipper.RotateToEnd
     Else
     PlaySoundAt SoundFX("FlipperDown",DOFFlippers), GILight4
     RightFlipper.RotateToStart
     End If
 End Sub

 'Drains and Kickers
Dim BallCount:BallCount = 0
   Sub Drain_Hit():PlaySoundAt "Drain", Drain
  'ClearBallID
  BallCount = BallCount - 1
  bsTrough.AddBall Me

   End Sub
   Sub BallRelease_UnHit()
  'NewBallID
    BallCount = BallCount + 1

  End Sub



'Switches

'1 monger down
'3 monger up
'4 monger left shoulder
'5 monger legs
Sub sw5_Hit:vpmTimer.PulseSw 5:PlaySound "target":End Sub
'6 monger rt shoulder

Sub sw7_Hit:Controller.Switch(7) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw7_UnHit:Controller.Switch(7) = 0:End Sub
Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub
Sub sw10_Hit():controller.switch(10)=true:End Sub
Sub sw11_Spin:vpmTimer.PulseSw 11::playsoundat"fx_spinner", sw11:End Sub
Sub sw12_Hit:Controller.Switch(12) = 1:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
Sub sw13_Spin:vpmTimer.PulseSw 13::playsoundat "fx_spinner", sw13:End Sub
Sub sw14_Spin:vpmTimer.PulseSw 14::playsoundat "fx_spinner", sw14:End Sub
Sub sw23_Hit:PlaySoundAt "rollover", ActiveBall:Controller.Switch(23)=1:End Sub
Sub sw23_UnHit:PlaySoundat "rollover", ActiveBall:Controller.Switch(23)=0:End Sub
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundat "rollover", ActiveBall:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundat "rollover", ActiveBall:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundat "rollover", ActiveBall:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundat "rollover", ActiveBall:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw33_Hit:vpmTimer.PulseSw 33:PlaySoundat "target", ActiveBall:End Sub
'Sub sw33_Timer:sw33.IsDropped = 0:sw33a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw34_Hit:vpmTimer.PulseSw 34:PlaySoundat "target", ActiveBall:End Sub
'Sub sw34_Timer:sw34.IsDropped = 0:sw34a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundat "target", ActiveBall:End Sub
'Sub sw35_Timer:sw35.IsDropped = 0:sw35a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySoundat "target", ActiveBall:End Sub
'Sub sw36_Timer:sw36.IsDropped = 0:sw36a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundat "gate", ActiveBall:RightCount = RightCount + 1:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundat "rollover", ActiveBall:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundat "rollover", ActiveBall:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundat "target", ActiveBall:End Sub
'Sub sw40_Timer:sw40.IsDropped = 0:sw40a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundat "target", ActiveBall:End Sub
'Sub sw41_Timer:sw41.IsDropped = 0:sw41a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundat "target", ActiveBall:End Sub
'Sub sw42_Timer:sw42.IsDropped = 0:sw42a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub TargetSw44_Hit:vpmTimer.PulseSw 44:PlaySoundat "target", ActiveBall:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
'Sub sw44_Timer:sw44.IsDropped = 0:sw44a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub TargetSw45_Hit:vpmTimer.PulseSw 45:PlaySoundat "target", ActiveBall:End Sub
'Sub sw45_Timer:sw45.IsDropped = 0:sw45a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub TargetSw46_Hit:vpmTimer.PulseSw 46:PlaySoundat "target", ActiveBall:End Sub
'Sub sw46_Timer:sw46.IsDropped = 0:sw46a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:PlaySoundat "target", ActiveBall:End Sub
'Sub sw47_Timer:sw47.IsDropped = 0:sw47a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:PlaySoundat "target", ActiveBall:End Sub
'Sub sw48_Timer:sw48.IsDropped = 0:sw48a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw49_Hit:Controller.Switch(49) = 1:PlaySoundat "gate", ActiveBall:LeftCount = LeftCount + 1:End Sub
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub
Sub TargetSw50_Hit:vpmTimer.PulseSw 50:PlaySoundat "target", ActiveBall:End Sub
'Sub sw50_Timer:sw50.IsDropped = 0:sw50a.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub ResetAll()
ClaneUpPost.Isdropped=true:UpPost.Isdropped=true
mongerhitpoint.isdropped=1:sw5.isdropped=1
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 27
  PlaySoundAt SoundFX("right_slingshot", DOFContactors), rightslinglight
  Me.TimerEnabled = 1
  RubberRightSS1.Visible = 0
    RubberRightSS2.Visible = 1
    RStep = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RubberRightSS2.Visible = 0:RubberRightSS3.Visible = 1
        Case 4:RubberRightSS3.Visible = 0:RubberRightSS1.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 26
  PlaySoundAt SoundFX("left_slingshot", DOFContactors), leftslinglight
  Me.TimerEnabled = 1
  RubberLeftSS1.Visible = 0
    RubberLeftSS2.Visible = 1
    LStep = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:RubberLeftSS2.Visible = 0:RubberLeftSS3.Visible = 1
        Case 4:RubberLeftSS3.Visible = 0:RubberLeftSS1.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


     ' Impulse Plunger
    Const IMPowerSetting = 50
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .InitExitSnd "plunger2", "plunger"
        .CreateEvents "plungerIM"
    End With

   'Bumpers
      Sub Bumper1b_Hit
      vpmTimer.PulseSw 31
      PlaySound SoundFX("fx_bumper1", DOFContactors)
      End Sub


      Sub Bumper2b_Hit
      vpmTimer.PulseSw 30
      PlaySound SoundFX("fx_bumper1", DOFContactors)
       End Sub

      Sub Bumper3b_Hit
      vpmTimer.PulseSw 32
      PlaySound SoundFX("fx_bumper1", DOFContactors)
       End Sub

Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim x

AllLampsOff()
LampTimer.Enabled = 1


'' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
        Next
    End If

    UpdateLamps
End Sub


 Sub UpdateLamps
NFadeL 3, l3
NFadeL 4, l4
NFadeL 5, l5
NFadeL 6, l6
NFadeL 7, l7
NFadeL 8, l8
NFadeL 9, l9
NFadeL 10, l10
NFadeL 11, l11
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
NFadeL 27, l27
NFadeL 28, l28
NFadeL 29, l29
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
NFadeL 48, l48
NFadeL 49, l49
NFadeL 50, l50
NFadeL 51, l51
NFadeL 52, l52
NFadeL 53, l53
NFadeL 54, l54
NFadeL 55, l55
NFadeL 56, l56
NFadeL 57, l57
NFadeL 58, l58
NFadeL 59, l59
NFadeL 63, l63
NFadeL 89, l89

LampMod 120, f120
LampMod 121, f121c
LampMod 121, f121a
LampMod 121, f121b
LampMod 122, f122
LampMod 123, f123
LampMod 125, f125a
if MongerPos < 10 then
  LampMod 125, f125b
Else
  f125b.Visible = False
end if

LampMod 126, f126a
LampMod 126, f126b
LampMod 126, f126c
LampMod 127, f127a
LampMod 127, f127b
LampMod 129, f129a
LampMod 129, f129b
LampMod 130, f130
LampMod 132, f132a
LampMod 132, f132b
LampMod 131, f131a
LampMod 131, f131b
LampMod 131, f131c
LampMod 131, f131d
NFadeL 60, Bumper2f160
LampMod 60, Bumper2bulbPr2
NFadeL 61, Bumper1f161
LampMod 61, Bumper1bulbPr3
NFadeL 62, Bumper3f162
LampMod 62, Bumper3bulbPr1


   End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 1
        Case 4:pri.image = b:FadingLevel(nr) = 2
        Case 5:pri.image = a:FadingLevel(nr) = 3
    End Select
End Sub

''Lights

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

' Flasher objects
' Uses own faster timer

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 1000 Then
                FlashLevel(nr) = 1000
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 0:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub


Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

'MONGER
'sw1 monger down
'sw3 monger up
'sw4 monger left shoulder
'sw5 monger legs
'sw6 monger rt shoulder


 '************************
 '    monger Animation
 '************************

Dim MongerMax:MongerMax = 240
Dim MongerSpeed:MongerSpeed = MongerMax / 1000
Dim monger, mongerPos, mongerDir, mongerFlash
mongerDir = 0:mongerPos = MongerMax:mongerFlash = 0

sw5.IsDropped = 1
switchframe.IsDropped = 1
 'For x = 1 to 9:monger(x).IsDropped = 1:Next

 Sub Solmonger(Enabled)
     If Enabled Then
         If mongerDir = 0 Then
             Controller.Switch(3) = 0
             Controller.Switch(1) = 1
             mongerOpen.Enabled = 0
             mongerClose.Enabled = 1
             mongerClose_Timer
             mongerDir = 1
         Else
             Controller.Switch(1) = 0
             Controller.Switch(3) = 1
             mongerClose.Enabled = 0
             mongerOpen.Enabled = 1
             mongerOpen_Timer
             mongerDir = 0
         End If
     End If
 End Sub




 Sub mongerClose_Timer()
     mongerPos = mongerPos + MongerSpeed
     If mongerPos> MongerMax Then
         mongerPos = MongerMax
         mongerOpen.Enabled = 0
     sw3.IsDropped = 1:switchframe.IsDropped = 1:sw5.IsDropped = 1
     End If
   Updatemonger
 End Sub

 Sub mongerOpen_Timer()
     mongerPos = mongerPos - MongerSpeed
     If mongerPos <0 Then
         mongerPos = 0
     sw3.IsDropped = 0:switchframe.IsDropped = 0:sw5.IsDropped = 0
         mongerClose.Enabled = 0
     End If
     Updatemonger
 End Sub

 Sub Updatemonger
   Primitive31.Z = MongerMax-MongerPos
   Primitive2.Z = Primitive31.Z - 122
   Primitive32.Z = Primitive31.Z
   f125b.Height = Primitive2.Z+170
 End Sub

Sub Trigger1_hit
  PlaySound "DROP_LEFT"
 End Sub

 Sub Trigger2_hit
  PlaySound "DROP_RIGHT"
 End Sub

Sub Table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
  'LFLogoUP.RotY = LeftFlipper1.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

set GICallback = GetRef("UpdateGI")

Sub UpdateGI(nr,enabled)
  Select Case nr
    Case 0
    If Enabled Then
      GIOn
    Else
      GIOff
    End If
  End Select
End Sub

Sub GIOn
  Table.ColorGradeImage= "ColorGrade_8"
  dim bulb
  for each bulb in GILights
    If TypeName(bulb) = "Flasher" Then
    bulb.visible = 1
  Else
    bulb.state = 1
  end if
  next
End Sub

Sub GIOff
  Table.ColorGradeImage= "ColorGrade_1"
  dim bulb
  for each bulb in GILights
  If TypeName(bulb) = "Flasher" Then
    bulb.visible = 0
  Else
    bulb.state = 0
  end if
  next
End Sub

 'Sub RightSlingShot_Timer:Me.TimerEnabled = 0:End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub RampEnters_Hit (idx)
  PlaySoundAtBall "pinhit_low"
End Sub

Sub Targets_Hit (idx)
  PlaySoundAtBall "target"
End Sub

Sub TargetBankWalls_Hit (idx)
  PlaySoundAtBall "target"
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySoundAtBall "metalhit_thin"
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySoundAtBall "metalhit_medium"
End Sub

Sub Metals2_Hit (idx)
  PlaySoundAtBall "metalhit2"
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4"
End Sub

Sub sw11_Spin
  PlaySoundAt "fx_spinner",sw11
End Sub

'************************
' RAMP HELPER SOUNDS
'************************
'
'Sub LRampKnock1_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub
'Sub LRampKnock2_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub
'Sub LRampKnock3_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub
'Sub LRampKnock4_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub
'Sub LRampKnock5_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub
'
'Sub RRampKnock1_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub
'Sub RRampKnock2_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub
'Sub RRampKnock3_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub
'Sub RRampKnock4_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub
'Sub RRampKnock5_Hit: PlaySoundAtBallVol "fx_lr7",.3:End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  debug.print finalspeed
  If finalspeed > 20 then
    PlaySoundAt "fx_rubber2", activeball
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
  debug.print finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySoundAt "fx_rubber2", ActiveBall
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBall "rubber_hit_1"
    Case 2 : PlaySoundAtBall "rubber_hit_2"
    Case 3 : PlaySoundAtBall "rubber_hit_3"
  End Select
End Sub

Sub RampBumps_Hit(idx)
  If Table.VersionMinor > 3 OR Table.VersionMajor > 10 Then
    RandomSoundRampBump()
  Else
    PlaySound "fx_lr7", 0, Vol(ActiveBall)*.3, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
    End If
End Sub

Sub RandomSoundRampBump()
' Select Case RndNum(1,3)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBallVol "fx_lr5",.3
    Case 2 : PlaySoundAtBallVol "fx_lr6",.3
    Case 3 : PlaySoundAtBallVol "fx_lr7",.3
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
    Case 1 : PlaySoundAtBall "flip_hit_1"
    Case 2 : PlaySoundAtBall "flip_hit_2"
    Case 3 : PlaySoundAtBall "flip_hit_3"
  End Select
End Sub


'Sub LRRail_Hit:PlaySound "fx_metalrolling", 0, 150, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub
'
'Sub RLRail_Hit:PlaySound "fx_metalrolling", 0, 150, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub

' *** ROLLING SOUND
Const tnob = 5
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
      if BOT(b).z < 30 then
        If Table.VersionMinor > 3 OR Table.VersionMajor > 10 Then
          PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b) ) * .8 , Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
        Else
          PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b) ) * .8 , Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        End If
      Else ' ball in air, probably on plastic.
        If Table.VersionMinor > 3 OR Table.VersionMajor > 10 Then
          PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b) ) * .2  , Pan(BOT(b) ), 0, Pitch(BOT(b) ) + 40000, 1, 0, AudioFade(BOT(b))
        Else
          PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b) ) * .2 , Pan(BOT(b) ), 0, Pitch(BOT(b) ) + 40000, 1, 0
        End If
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
  If Table.VersionMinor > 3 OR Table.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 80, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 80, Pan(ball1), 0, Pitch(ball1), 0, 0
  end if
End Sub


''*****************************************
''      BALL SHADOW
''*****************************************
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
        If BOT(b).X < Table.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


Dim LeftCount:LeftCount = 0
Sub leftdrop_hit
  If LeftCount = 1 then
    playsound "BallDrop"
  End If
  LeftCount = 0
End Sub

Dim RightCount:RightCount = 0
Sub rightdrop_hit
  If RightCount = 1 then
    playsound "BallDrop"
  End If
  RightCount = 0
End Sub

Sub RLS_Timer()
'              RampGate1.RotZ = -(Spinner4.currentangle)
'              RampGate2.RotZ = -(Spinner1.currentangle)
              RampGate1.RotZ = -(leftrampgate.currentangle)
              RampGate3.RotZ = -(rightrampgate.currentangle)
              SpinnerT3.RotZ = -(sw11.currentangle)
              SpinnerT1.RotZ = -(sw13.currentangle)
              SpinnerT2.RotZ = -(sw14.currentangle)
End Sub


Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 100)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Table" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / Table.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Sub PlaySoundAtBallVol(sound, VolMult)
  If Table.VersionMinor > 3 OR Table.VersionMajor > 10 Then
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
  Else
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
  End If
End Sub

Sub PlaySoundAt(sound, tableobj)
  If Table.VersionMinor > 3 OR Table.VersionMajor > 10 Then
    PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
  Else
    PlaySound sound, 1, 1, Pan(tableobj)
  End If
End Sub

Sub PlaySoundAtBall(sound)
  If Table.VersionMinor > 3 OR Table.VersionMajor > 10 Then
    PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
  Else
    PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
  End If
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
  If Table.VersionMinor > 3 OR Table.VersionMajor > 10 Then
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
  Else
    PlaySound sound, 1, Vol, Pan(tableobj)
  End If
End Sub

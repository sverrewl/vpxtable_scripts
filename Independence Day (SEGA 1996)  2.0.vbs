Option Explicit
Randomize

' Thalamus - added vpminit me to _init for cvpmflips / Fastflip

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01530000","SEGA.vbs",3.1
'********************************************
'**     Game Specific Code Starts Here     **
'********************************************

' Thalamus - ffv2 should not be used for this table according to nFozzy
Const cGameName="id4",UseSolenoids=1,UseLamps=0,UseGI=0, SCoin="coin"
' Const cSingleLFlip = 0
' Const cSingleRFlip = 0

'**************************************
'**     Bind Events To Solenoids     **
'**************************************

SolCallback(1)="bsTroughKick.SolOut"
SolCallback(2)="AutoPlunger"
SolCallback(3)="bsVUK.SolOut"

SolCallback(6) ="SetLamp 106," 'Flash Bottom
SolCallback(7) ="SetLamp 107," 'Flash Right Ramp
SolCallback(8) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'SolCallback(9) = "vpmSolSound ""Bumper1"","
'SolCallback(10) = "vpmSolSound ""Bumper2"","
'SolCallback(11) = "vpmSolSound ""Bumper3"","
'SolCallback(12) = "vpmSolSound ""Slingshot_Left"","
'SolCallback(13) = "vpmSolSound ""Slingshot_Right"","

SolCallback(17)="SolLockout"

SolCallback(19) ="LGate"  'Left Control Gate
SolCallback(20) ="RGate"  'Right Control Gate

SolCallback(23) ="SetLamp 123," 'Flash Bottom

SolCallback(25) ="SetLamp 125," 'F1
SolCallback(26) ="SetLamp 126," 'F2
SolCallback(27) ="SetLamp 127," 'F3
SolCallback(28) ="SetLamp 128," 'F4
SolCallback(29) ="SetLamp 129," 'F5
SolCallback(30) ="SetLamp 130," 'F6
SolCallback(31) ="SetLamp 131," 'F7
SolCallback(32) ="SetLamp 132," 'F8

SolCallback(34)="OpenAlien"

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
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToEnd:UpperRightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToStart:UpperRightFlipper.RotateToStart
     End If
End Sub


Sub AutoPlunger(Enabled)
  if enabled then
  Plunger.Kick 0,55
    PlaySoundAtVol SoundFX("Popper",DOFContactors), Plunger, 1
  end if
End Sub

Sub LGate(Enabled)
  if enabled then
  LeftGate.Open = 1
    PlaySoundAtVol SoundFX("Popper",DOFContactors), LeftGate, 1
  Else
  LeftGate.Open = 0
  end if
End Sub


Sub RGate(Enabled)
  if enabled then
  RightGate.Open = 1
    PlaySoundAtVol SoundFX("Popper",DOFContactors), RightGate, 1
  Else
  RightGate.Open = 0
  end if
End Sub

'Stern-Sega GI
set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
'   SetLamp 134, Enabled  'Backwall bulbs
  If Enabled Then
    Dim xx
    For each xx in GI:xx.State = 1:Next
    PlaySound "fx_relay"
    DOF 101, DOFOn
  Else
    For each xx in GI:xx.State = 0:Next
    PlaySound "fx_relay"
    DOF 101, DOFOff
  End If
End Sub


'*******************************
'**     Keyboard Handlers     **
'*******************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=PlungerKey or KeyCode=LockBarKey Then Controller.Switch(53)=1
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=PlungerKey or KeyCode=LockBarKey Then Controller.Switch(53)=0
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'********************************************
'**     Init The Table, Start VPinMAME     **
'********************************************

Dim bsTroughKick,bstrough,bsvuk

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "ID4 Sega"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 0
        .Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval:
  PinMAMETimer.Enabled=1:
  vpmNudge.TiltSwitch=1:
  vpmNudge.Sensitivity=2
  vpmNudge.TiltObj=Array(LeftSlingShot, RightSlingShot, LeftBumper, BotBumper, RightBumper)

    set bsTrough = new cvpmBallStack
      bsTrough.InitSw 0,14,13,12,11,0,0,0
      bsTrough.Balls = 4

  Set bsTroughKick = new cvpmBallStack
    bsTroughKick.Initsaucer BallRelease,15,85,7
    bsTroughKick.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)

  set bsvuk =new cvpmballstack
    bsvuk.initsw 0,44,0,0,0,0,0,0
    bsvuk.initkick Super_Vuk,180,10
    bsvuk.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)


    AlienHeadWall.IsDropped=False

End Sub

'*****************************
'**        LockBall         **
'*****************************

Sub SolLockout(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    BallRelease.CreateBall
    bsTroughKick.AddBall 0
  End If
End Sub

'*****************************
'**        FrontVuk         **
'*****************************

Sub  FrontVuk_Hit
     vpmTimer.PulseSw 45
     FrontVuk.DestroyBall
   playsoundAtVol "popper_ball", FrontVuk, 1
     bsvuk.AddBall 0
End Sub

'*****************************
'**        SideVuk          **
'*****************************

Sub  SideVuk_Hit
     vpmTimer.PulseSw 46
     vpmTimer.PulseSw 41
     SideVuk.DestroyBall
   playsoundAtVol "popper_ball", SideVuk, 1
     bsvuk.AddBall 0
End Sub

'*****************************
'**        RearVuk          **
'*****************************

Sub  RearVuk_Hit
     vpmTimer.PulseSw 43
     vpmTimer.PulseSw 41
     RearVuk.DestroyBall
   playsoundAtVol "popper_ball", RearVuk, 1
     bsvuk.AddBall 0
End Sub

'*****************************
'**          HitFX          **
'*****************************
Sub AlienHeadWall_Hit:PlaysoundAtVol "fx_rubber2", ActiveBall, 1:End Sub
Sub Gate001_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate002_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate003_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate004_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate005_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate006_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate007_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate008_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate009_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate010_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
Sub Gate011_Hit:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
'*****************************
'**        AlienHead        **
'*****************************

Sub TriggerAlienHeadOpto_Hit:vpmTimer.PulseSw 42:End Sub

Sub OpenAlien(Enabled)
    If Enabled Then
    Controller.Switch(52)=True
    AlienHeadWall.IsDropped=True
    LeftHeadPrim.TransX = +25
    LeftHeadPrim.RotZ = +25
    RightHeadPrim.TransX = -25
    RightHeadPrim.RotZ = -25
  PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftHeadPrim, 1
    Else
    Controller.Switch(52)=False
    AlienHeadWall.IsDropped=False
    LeftHeadPrim.TransX = 0
    LeftHeadPrim.RotZ = 0
    RightHeadPrim.TransX = 0
    RightHeadPrim.RotZ = 0
    End If
End Sub

'*****************************
'**        Targets          **
'*****************************

Sub MiniLoopSU_Hit:vpmTimer.PulseSw 25:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub CenterRampSU_Hit:vpmTimer.PulseSw 27:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub RightRampLeftSU_Hit:vpmTimer.PulseSw 28:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub RightRampRightSU_Hit:vpmTimer.PulseSw 29:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub superjackpotSU_Hit:vpmTimer.PulseSw 24:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub Area51Bot_Hit:vpmTimer.PulseSw 47:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub Area51Mid_Hit:vpmTimer.PulseSw 34:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub Area51Top_Hit:vpmTimer.PulseSw 33:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub HurryUpBot_hit:vpmTimer.PulseSw 48:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub HurryUpMid_Hit:vpmTimer.PulseSw 36:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub HurryUpTop_Hit:vpmTimer.PulseSw 35:PlaysoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub

'*****************************
'**     TurboBumpers        **
'*****************************

Sub LeftBumper_Hit:vpmTimer.PulseSw 49:PlaysoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1:End Sub
Sub BotBumper_Hit:vpmTimer.PulseSw 50:PlaysoundAtVol SoundFX("fx_bumper3",DOFContactors), ActiveBall, 1:End Sub
Sub RightBumper_Hit:vpmTimer.PulseSw 51:PlaysoundAtVol SoundFX("fx_bumper2",DOFContactors), ActiveBall, 1:End Sub



'*****************************
'**     Switch Handling     **
'*****************************

Sub Drain_Hit:bsTrough.AddBall Me:PlaysoundAtVol "drain", Drain, 1:End Sub
Sub Trigger16_Hit:Controller.switch(16) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger16_Unhit:Controller.switch(16) = False:end sub
Sub Trigger18_Hit:Controller.switch(18) = True : playsoundAtVol"RampUp" , ActiveBall, 1: End Sub
Sub Trigger18_Unhit:Controller.switch(18) = False:end sub
Sub Trigger19_Hit:Controller.switch(19) = True : playsoundAtVol"fx_metalrolling" , ActiveBall, 1: End Sub
Sub Trigger19_Unhit:Controller.switch(19) = False:end sub
Sub Trigger20_Hit:Controller.switch(20) = True : playsoundAtVol"RampUp" , ActiveBall, 1: End Sub
Sub Trigger20_Unhit:Controller.switch(20) = False:end sub
Sub Trigger21_Hit:Controller.switch(21) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger21_Unhit:Controller.switch(21) = False:end sub
Sub Trigger22_Hit:Controller.switch(22) = True : playsoundAtVol"RampUp" , ActiveBall, 1: End Sub
Sub Trigger22_Unhit:Controller.switch(22) = False:end sub
Sub Trigger23_Hit:Controller.switch(23) = True : playsoundAtVol"fx_metalrolling" , ActiveBall, 1: End Sub
Sub Trigger23_Unhit:Controller.switch(23) = False:end sub
Sub Trigger57_Hit:Controller.switch(57) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger57_Unhit:Controller.switch(57) = False:end sub
Sub Trigger58_Hit:Controller.switch(58) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger58_Unhit:Controller.switch(58) = False:end sub
Sub Trigger60_Hit:Controller.switch(60) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger60_Unhit:Controller.switch(60) = False:end sub
Sub Trigger61_Hit:Controller.switch(61) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger61_Unhit:Controller.switch(61) = False:end sub
Sub Trigger37_Hit:Controller.switch(37) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger37_Unhit:Controller.switch(37) = False:end sub
Sub Trigger38_Hit:Controller.switch(38) = True : playsoundAtVol"RampMidway" , ActiveBall, 1: End Sub
Sub Trigger38_Unhit:Controller.switch(38) = False:end sub
Sub Trigger39_Hit:Controller.switch(39) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger39_Unhit:Controller.switch(39) = False:end sub
Sub Trigger40_Hit:Controller.switch(40) = True : playsoundAtVol"RampEnd" , ActiveBall, 1: End Sub
Sub Trigger40_Unhit:Controller.switch(40) = False:end sub
Sub Trigger30_Hit:Controller.switch(30) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger30_Unhit:Controller.switch(30) = False:end sub
Sub Trigger31_Hit:Controller.switch(31) = True : playsoundAtVol"RampMidway" , ActiveBall, 1: End Sub
Sub Trigger31_Unhit:Controller.switch(31) = False:end sub
Sub Trigger32_Hit:Controller.switch(32) = True : playsoundAtVol"RampEnd" , ActiveBall, 1: End Sub
Sub Trigger32_Unhit:Controller.switch(32) = False:end sub
Sub Trigger9_Hit:Controller.switch(9) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger9_Unhit:Controller.switch(9) = False:end sub
'*****************************
'**      Support Code       **
'*****************************

Sub RampHelper1a_Hit
  Activeball.velz = 0
End Sub

Sub RampHelper1b_Hit
  Activeball.velz = 0
End Sub

Sub RampHelper1c_Hit
  Activeball.velz = 0
End Sub

Sub RampHelper2a_Hit
  Activeball.velz = 0
End Sub

Sub RampHelper2b_Hit
  Activeball.velz = 0
End Sub

Sub RampHelper2c_Hit
  Activeball.velz = 0
End Sub

Sub RampHelper3_Hit
  Activeball.velz = 0
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
LampTimer.Interval = 5 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step

     'Special Handling
     'If chgLamp(ii,0) = 2 Then solTrough chgLamp(ii,1)
     'If chgLamp(ii,0) = 4 Then PFGI chgLamp(ii,1)

        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()
NFadeL 1, Light1
NFadeL 2, Light2
NFadeL 3, Light3
NFadeL 4, Light4
NFadeL 5, Light5
NFadeL 6, Light6
NFadeL 7, Light7
NFadeL 8, Light8
NFadeL 9, Light9
NFadeL 10, Light10
NFadeL 11, Light11
NFadeL 12, Light12
NFadeL 13, Light13
NFadeL 14, Light14
NFadeL 15, Light15
NFadeL 16, Light16
NFadeL 17, Light17
NFadeL 18, Light18
NFadeL 19, Light19
NFadeL 20, Light20
NFadeL 21, Light21
NFadeL 22, Light22
NFadeL 23, Light23
NFadeL 24, Light24
NFadeL 25, Light25
NFadeL 26, Light26
NFadeL 27, Light27
NFadeL 28, Light28
NFadeL 29, Light29
NFadeL 30, Light30
NFadeL 31, Light31
NFadeL 32, Light32
NFadeL 33, Light33
NFadeL 34, Light34
NFadeL 35, Light35
NFadeL 36, Light36
NFadeL 37, Light37
NFadeL 38, Light38 'Ramp LED
NFadeL 39, Light39
'NFadeL 40, Light40 'Not used
NFadeL 41, Light41
NFadeL 42, Light42
NFadeL 43, Light43
NFadeL 44, Light44
NFadeL 45, Light45
NFadeL 46, Light46
NFadeL 47, Light47
NFadeL 48, Light48
NFadeL 49, Light49 'Bumper
NFadeL 50, Light50 'Bumper
NFadeL 51, Light51 'Bumper
NFadeL 52, Light52
NFadeL 53, Light53
NFadeL 54, Light54
NFadeL 55, Light55
NFadeL 56, Light56
'NFadeL 57, Light57 'Alien Head
NFadeL 58, Light58
NFadeL 59, Light59
NFadeL 60, Light60
NFadeL 61, Light61
NFadeL 62, Light62
NFadeL 63, Light63
'NFadeL 64, Light64 'Launch Button
NFadeL 65, Light65

'Solenoid Controlled Lamps

 NFadeL 106, S106
 NFadeL 107, S107
 NFadeL 123, S123
 NFadeL 125, S125
 NFadeL 126, S126
 NFadeL 127, S127
 NFadeL 128, S128
 NFadeL 129, S129
 NFadeLm 130, S130a
 NFadeL 130, S130
 NFadeL 131, S131
 NFadeL 132, S132

End Sub




' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 62
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1

End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:

    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 59
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1

End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
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

' Wire ramp sounds

Sub RampSound1_Hit: PlaySoundAtVol"fx_metalrolling", ActiveBall, 1: End Sub
Sub RampSound2_Hit: PlaySoundAtVol"fx_metalrolling", ActiveBall, 1: End Sub
Sub RampSound3_Hit: PlaySoundAtVol"fx_metalrolling", ActiveBall, 1: End Sub



'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
  UpdateFlipperLogo
End Sub

Sub UpdateFlipperLogo
  LogoSx.RotZ = LeftFlipper.CurrentAngle - 90
  LogoDx1.RotZ = UpperRightFlipper.CurrentAngle - 90
  LogoDx.RotZ = RightFlipper.CurrentAngle + 90
End Sub

' Thalamus - 2021-04-30 : added proper exit

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


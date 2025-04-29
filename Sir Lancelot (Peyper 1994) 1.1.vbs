'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Sir Lancelot                                                       ########
'#######          (Peyper 1994)                                                      ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.0 mfuegemann 2016
'
' Thanks to:
' Akiles for the images of the playfield and the plastics
' Fuzzel for providing the images for the playfield lights
' JPSalas for the fading Flasher code, the Ball Gate and some textures
'

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


Const cGameName="lancelot",UseSolenoids=1,UseLamps=0,UseGI=0, SCoin="coin3"



LoadVPM "02060000","Lancelot.VBS",3.2

Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp80.visible=1
Ramp81.visible=1
Primitive005.visible=1
Else
Ramp80.visible=0
Ramp81.visible=0
Primitive005.visible=0
End if

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------

Const FreePlay = True     'set to True for a faked FreePlay (every GameStart inserts a coin)
Const ForceFlippers = False   'set to True if You have the flippers not responding to every button push (delete nvram file can help too)

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------

'Sol1 CBumper
'Sol2 RBumper
'Sol3 LBumper
'Sol4 LSling
SolCallback(5)="Sol5"
SolCallback(6)="Sol6"
SolCallback(8)="SolResetComplete"     'Helper Function
SolCallback(9)="SolPlunger"

SolCallback(15)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"

SolCallback(17)="VpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd':LaunchFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart':LaunchFlipper.RotateToStart
     End If
End Sub


Set LampCallback = GetRef("UpdateLamps")

'Drop Target Reset Animation
Sub Sol5(enabled)
  if enabled then
    DropTargetBank2.DropSol_On
  end if
end sub

Sub Sol6(enabled)
  if enabled then
    DropTargetBank.DropSol_On
  end if
end sub

Sub SolPlunger(enabled)
  if enabled then
    LaunchFlipper.rotatetoend
    LaunchFlipper.timerenabled = True
  end if
End Sub
Sub LaunchFlipper_Timer
  LaunchFlipper.timerenabled = False
  LaunchFlipper.rotatetostart
End Sub


Dim InitComplete
InitComplete = False
Sub SolResetComplete(enabled)
  if enabled then
    InitComplete = True
  end if
End Sub
Sub InitBackupTimer_Timer
  InitBackupTimer.enabled = False
  InitComplete = True
End Sub

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough,obj,DropTargetBank,DropTargetBank2,cLeftCaptive,cRightCaptive,Flipperactive
Const TiltSwitch = 36

Sub Table1_Init


  vpminit me

  BallRelease.createBall
  BallRelease.kick 180,1

    Controller.GameName=cGameName
    Controller.SplashInfoLine="Sir Lancelot" & vbNewLine & "created by mfuegemann"
    Controller.HandleKeyboard=False
    Controller.ShowTitle=0
    Controller.ShowFrame=0
    Controller.ShowDMDOnly=1
' Controller.Hidden = 1     'enable to hide DMD if You use a B2S backglass


    Controller.HandleMechanics=0

  Controller.SolMask(0)=0
    vpmTimer.AddTimer 4000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds

    Controller.Run
    If Err Then MsgBox Err.Description
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch = TiltSwitch
    vpmNudge.Sensitivity = 3
  vpmNudge.TiltObj = Array(LeftFlipper,RightFlipper,LeftSlingshot,Bumper1,Bumper2,Bumper3)

'    vpmMapLights AllLights

  set DropTargetBank = new cvpmDropTarget
    DropTargetBank.InitDrop Array(DT1,DT2,DT3), Array(17,18,19)
    DropTargetBank.InitSnd SoundFX("Targetdrop1",DOFContactors),SoundFX("TargetBankreset1",DOFContactors)
    DropTargetBank.CreateEvents "DropTargetBank"

  set DropTargetBank2 = new cvpmDropTarget
    DropTargetBank2.InitDrop Array(DT4,DT5,DT6), Array(20,21,22)
    DropTargetBank2.InitSnd SoundFX("Targetdrop1",DOFContactors),SoundFX("TargetBankreset1",DOFContactors)
    DropTargetBank2.CreateEvents "DropTargetBank2"

End Sub

Sub Table1_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub


'------------------------------
'------  Trough Handler  ------
'------------------------------
Sub Drain_Hit:PlaySoundAtVol "Drain7", Drain, 1 : End Sub

'-------------------------------
'------  Keybord Handler  ------
'-------------------------------

Sub Table1_KeyDown(ByVal keycode)
  if InitComplete then
    if keycode = startgamekey then
      if FreePlay then
        vpmTimer.PulseSw 38
      end if
      vpmTimer.PulseSw 37
    end if
    If keycode = PlungerKey Then
      vpmTimer.PulseSw 34
    End If

    vpmKeyDown(KeyCode)

    If keycode = LeftFlipperKey Then
      controller.switch(35) = 1
      if Flipperactive then
        if ForceFlippers then
          LeftFlipper.rotatetoend
        end if
        PlaySoundAtVol SoundFX("FlipperUp_Akiles",DOFContactors), LeftFlipper, 1
      end if
    End If
    If keycode = RightFlipperKey Then
      controller.switch(34) = 1
      if Flipperactive then
        if ForceFlippers then
          RightFlipper.rotatetoend
        end if
        PlaySoundAtVol SoundFX("FlipperUp_Akiles",DOFContactors), RightFlipper, 1
      end if
    End If

    If keycode = LeftMagnaSave Then Nudge 90, 3:PlaySound "fx_nudge"
    If keycode = RightMagnaSave Then Nudge 270, 3:PlaySound "fx_nudge"

    'If vpmKeyDown(KeyCode) Then Exit Sub
  end if
End Sub

Sub Table1_KeyUp(ByVal keycode)
  vpmKeyUp(KeyCode)

  If keycode = LeftFlipperKey Then
    controller.switch(35) = 0
    if Flipperactive then
      if ForceFlippers then
        LeftFlipper.rotatetostart
      end if
      PlaySoundAtVol SoundFX("FlipperDown_Akiles",DOFContactors), LeftFlipper, 1
    end if
  End If
  If keycode = RightFlipperKey Then
    controller.switch(34) = 0
    if Flipperactive then
      if ForceFlippers then
        RightFlipper.rotatetostart
      end if
      PlaySoundAtVol SoundFX("FlipperDown_Akiles",DOFContactors), RightFlipper, 1
    end if
  End If

  'If vpmKeyUp(KeyCode) Then Exit Sub
End Sub


'------------------------------
'------  Switch Handler  ------
'------------------------------

Sub DT1_Hit:DropTargetBank.Hit 1:End Sub
Sub DT2_Hit:DropTargetBank.Hit 2:End Sub
Sub DT3_Hit:DropTargetBank.Hit 3:End Sub
Sub DT4_Hit:DropTargetBank2.Hit 1:End Sub
Sub DT5_Hit:DropTargetBank2.Hit 2:End Sub
Sub DT6_Hit:DropTargetBank2.Hit 3:End Sub

sub Trigger1_hit:Controller.Switch(1)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
sub Trigger1_unhit:Controller.Switch(1)=0:End Sub

Sub TopA_Hit:Controller.Switch(15)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub TopA_Unhit:Controller.Switch(15)=0:End Sub
Sub TopB_Hit:Controller.Switch(14)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub TopB_Unhit:Controller.Switch(14)=0:End Sub
Sub TopC_Hit:Controller.Switch(13)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub TopC_Unhit:Controller.Switch(13)=0:End Sub
Sub LeftLane_Hit:Controller.Switch(12)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub LeftLane_Unhit:Controller.Switch(12)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(11)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub LeftInlane_Unhit:Controller.Switch(11)=0:End Sub
Sub RightInlane1_Hit:Controller.Switch(10)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub RightInlane1_Unhit:Controller.Switch(10)=0:End Sub
Sub RightInlane2_Hit:Controller.Switch(9)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub RightInlane2_Unhit:Controller.Switch(9)=0:End Sub

'Bumpers
Sub Bumper1_Hit: vpmTimer.PulseSw 3: PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1:End Sub
Sub Bumper2_Hit: vpmTimer.PulseSw 4: PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1:End Sub
Sub Bumper3_Hit: vpmTimer.PulseSw 5: PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1:End Sub



Sub RightRampEnd_Hit
  vpmTimer.PulseSw 26
  PlaySoundAtVol "Gate2", ActiveBall, 1:
  P_RightRampEnd.ObjRotY = 30
End Sub

Sub RightRampEnd_Unhit:RightRampEnd.timerenabled = True:End Sub

Sub LeftRampEnd_Hit
  vpmTimer.PulseSw 25
  PlaySoundAtVol "Gate2", ActiveBall, 1:
  P_LeftRampEnd.ObjRotX = -30
End Sub

Sub LeftRampEnd_Unhit:LeftRampEnd.timerenabled = True:End Sub

Sub RightLane_Hit:Controller.Switch(32)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub RightLane_Unhit:Controller.Switch(32)=0::End Sub


Sub LeftTarget_Hit:vpmTimer.PulseSw 23:PlaySoundAtVol SoundFX("Target1",DOFContactors), ActiveBall, 1:End Sub

Sub LeftOutlane_Hit:Controller.Switch(30)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub LeftOutlane_Unhit:Controller.Switch(30)=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(31)=1:PlaySoundAtVol "Gate2", ActiveBall, 1:End Sub
Sub RightOutlane_Unhit:Controller.Switch(31)=0:End Sub

Sub Rubber1_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Rubber2_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Rubber3_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Rubber4_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub

Sub BehindTopDT_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol SoundFX("Target1", DOFContactors), ActiveBall, 1:End Sub
Sub BehindRightDT_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol SoundFX("Target1", DOFContactors), ActiveBall, 1:End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim Lstep


Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 2
    PlaySoundAtVol SoundFX("Slingshot_Akiles",DOFContactors), ActiveBall, 1
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

'--------------------------------
'------  Helper Functions  ------
'--------------------------------
Sub Gate_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub
Sub Gate6_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub

Sub Wall59_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall40_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall45_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall70_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall21_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall17_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall29_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall69_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall20_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall16_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub PostWall5_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall28_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall30_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall47_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wall57_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub
Sub Wallx_Hit:PlaySoundAtVol "Bump", ActiveBall, 1:End Sub

Sub MovimientoTimer_Timer
  MovimientoTimer.enabled = False
End Sub




'Ramp End Triggers
Sub RightRampEnd_timer
  P_RightRampEnd.ObjRotY = P_RightRampEnd.ObjRotY - 1.5
  if P_RightRampEnd.ObjRotY <= 0 then
    RightRampEnd.timerenabled = False
    P_RightRampEnd.ObjRotY = 0
  end if
End Sub

Sub LeftRampEnd_timer
  P_LeftRampEnd.ObjRotX = P_LeftRampEnd.ObjRotX + 1.5
  if P_LeftRampEnd.ObjRotX >= 0 then
    LeftRampEnd.timerenabled = False
    P_LeftRampEnd.ObjRotX = 0
  end if
End Sub


'Flipper and DT Primitives
Const DTUpperLimit = -1
Const DTBottomLimit = -47
Const DTTransZ = 4
Const DownSpeed = 15
Const UpSpeed = 10



'Right Wire Gate
Dim RightGateDir,RightGateSpeed,RightSwingBack
Sub RightGateUpTrigger_Hit
  if not RightGateTimer.enabled then
    vpmTimer.PulseSw 27
    PlaySoundAtVol "Gate5", ActiveBall, 1
    RightGateSpeed = 4.5
    RightGateDir = -1
    RightSwingBack = False
    RightGateTimer.enabled = True
  end if
End Sub

Sub RightGateDownTrigger_Hit
  if not RightGateTimer.enabled then
    PlaySoundAtVol "Gate5", ActiveBall, 1
    RightGateSpeed = 4.5
    RightGateDir = 1
    RightSwingBack = False
    RightGateTimer.enabled = True
  end if
End Sub


'Left Wire Gate
Dim LeftGateDir,LeftGateSpeed,LeftSwingBack
Sub LeftGateUpTrigger_Hit
  if not LeftGateTimer.enabled then
    vpmTimer.PulseSw 27
    PlaySoundAtVol "Gate5", ActiveBall, 1
    LeftGateSpeed = 4.5
    LeftGateDir = -1
    LeftSwingBack = False
    LeftGateTimer.enabled = True
  end if
End Sub

Sub LeftGateDownTrigger_Hit
  if not LeftGateTimer.enabled then
    PlaySoundAtVol "Gate5", ActiveBall, 1
    LeftGateSpeed = 4.5
    LeftGateDir = 1
    LeftSwingBack = False
    LeftGateTimer.enabled = True
  end if
End Sub

Sub UpdateLamps
  Flipperactive = Controller.solenoid(17)
  if Controller.L19 then
    L19Timer.enabled = False
  else
    L19Timer.enabled = True
  end if
End sub


Sub L19Timer_Timer
  L19Timer.enabled = False
  P_Bumper3.image = "Bumper_Lancelot"
  setflash 6, False
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
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps()

    NFadeL 1, lamp1
    NFadeL 2, lamp2
    NFadeL 3, lamp3
    NFadeL 4, lamp4
    NFadeL 5, lamp5
    NFadeL 6, lamp6
    NFadeL 7, lamp7
    NFadeL 9, lamp9

    NFadeL 10, lamp10
    NFadeL 11, lamp11
    NFadeL 12, lamp12
  NFadeL 13, lamp13
  NFadeL 14, lamp14
  NFadeL 15, lamp15

' NFadeL 17, L17
' NFadeL 18, L18

  NFadeL 19, L19 'Bumper3

  NFadeL 20, Lamp20 'Flasher
  NFadeL 21, lamp21' Flasher
  NFadeL 22, lamp22
  NFadeL 23, lamp23
  NFadeL 24, Lamp24

  NFadeL 26, lamp26
  NFadeL 27, lamp27
  NFadeL 28, lamp28
  NFadeL 29, lamp29

  NFadeL 30, lamp30
' NFadeL 31, lamp31

    NFadeL 33, lamp33
    NFadeL 34, lamp34
    NFadeL 35, lamp35
    NFadeL 36, lamp36
    NFadeL 37, lamp37
    NFadeL 38, lamp38
    NFadeL 39, lamp39

  NFadeL 40, lamp40
  NFadeL 41, lamp41
  NFadeL 42, lamp42
  NFadeL 43, lamp43
  NFadeL 44, lamp44
    NFadeL 45, lamp45

'    NFadeL 46, lamp46

    NFadeL 47, lamp47
    NFadeL 48, lamp48
  NFadeL 49, lamp49

    NFadeL 50, lamp50
    NFadeL 51, lamp51
    NFadeL 52, lamp52
    NFadeL 53, lamp53

'    NFadeL 54, lamp54
'    NFadeL 55, lamp55

    NFadeL 57, lamp57
    NFadeL 58, lamp58
    NFadeL 59, lamp59

    NFadeL 60, lamp60
    NFadeL 61, lamp61
    NFadeL 62, lamp62
    NFadeL 63, lamp63
    NFadeL 64, lamp64

    NFadeL 76, lamp76 'Flasher
    NFadeL 77, lamp77 'Flasher
    NFadeL 78, lamp78 'Flasher
    NFadeL 79, lamp79 'Flasher

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

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(61)

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
Digits(35) = Array(LED290,LED291,LED292,LED293,LED294,LED295,LED296)
Digits(36) = Array(LED297,LED298,LED299,LED300,LED301,LED302,LED303)
Digits(37) = Array(LED304,LED305,LED306,LED307,LED308,LED309,LED310)
Digits(38) = Array(LED311,LED312,LED313,LED314,LED315,LED316,LED317)
Digits(39) = Array(LED318,LED319,LED320,LED321,LED322,LED323,LED324)
Digits(40) = Array(LED325,LED326,LED327,LED328,LED329,LED330,LED331)
Digits(41) = Array(LED332,LED333,LED334,LED335,LED336,LED337,LED338)
Digits(42) = Array(LED339,LED340,LED341,LED342,LED343,LED344,LED345)
Digits(43) = Array(LED346,LED347,LED348,LED349,LED350,LED351,LED352)
Digits(44) = Array(LED353,LED354,LED355,LED356,LED357,LED358,LED359)
Digits(45) = Array(LED360,LED361,LED362,LED363,LED364,LED365,LED366)
Digits(46) = Array(LED367,LED368,LED369,LED370,LED371,LED372,LED373)
Digits(47) = Array(LED374,LED375,LED376,LED377,LED378,LED379,LED380)
Digits(48) = Array(LED381,LED382,LED383,LED384,LED385,LED386,LED387)
Digits(49) = Array(LED388,LED389,LED390,LED391,LED392,LED393,LED394)
Digits(50) = Array(LED395,LED396,LED397,LED398,LED399,LED400,LED401)
Digits(51) = Array(LED402,LED403,LED404,LED405,LED406,LED407,LED408)
Digits(52) = Array(LED409,LED410,LED411,LED412,LED413,LED414,LED415)

Digits(53) = Array(LED416,LED417,LED418,LED419,LED420,LED421,LED422)
Digits(54) = Array(LED423,LED424,LED425,LED426,LED427,LED428,LED429)
Digits(55) = Array(LED430,LED431,LED432,LED433,LED434,LED435,LED436)
Digits(56) = Array(LED437,LED438,LED439,LED440,LED441,LED442,LED443)
Digits(57) = Array(LED444,LED445,LED446,LED447,LED448,LED449,LED450)
Digits(58) = Array(LED451,LED452,LED453,LED454,LED455,LED456,LED457)
Digits(59) = Array(LED458,LED459,LED460,LED461,LED462,LED463,LED464)
Digits(60) = Array(LED465,LED466,LED467,LED468,LED469,LED470,LED471)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 61) then
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

'-----------------------------------
'------  DIP Switch Settings  ------
'-----------------------------------
Dim vpmDips

Sub editDips
  Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 150,240,"Sir Lancelot - DIP switches"
    .AddFrame 0,0,120,"Balls per game",&H00000004,Array("3 Balls",0,"5 Balls",&H00000004)'dip A-3
    .AddFrame 0,50,120,"Replay Score",&H00000003,Array("30/24/20M points",0,"35/28/24M points",&H00000001,"42/32/28M points",&H00000002,"50/36/32M points",&H00000003)'dip A-1&A-2
    .AddFrame 150,0,120,"Match Award",&H00000040,Array("Credit",0,"Double Score",&H00000040)'dip A-7
    .AddChk 150,55,120,Array("Buy-In Feature",&H00000008)'dip A-4
    .AddChk 150,73,120,Array("Enable Music",&H00000020)'dip A-6
    .AddChk 150,91,120,Array("Adaptive Scoring",&H00000010)'dip A-5
    .AddChk 150,109,120,Array("Save Replay Values",&H00000080)'dip A-8
    .AddLabel 0,130,280,20,"After hitting OK, press F3 to reset game with new settings."
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
  LFPrim.RotZ = LeftFlipper.CurrentAngle
  RFPrim.RotZ = RightFlipper.CurrentAngle

End Sub


'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

'Sub BallShadowUpdate_timer()
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
'End Sub



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

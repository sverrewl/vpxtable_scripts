Option Explicit
Randomize
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="nbamac",UseSolenoids=1,UseLamps=0,UseGI=0, SCoin="coin3"

Const FreePlay = True     'add a coin on StartGame

LoadVPM "01560000","mac.VBS",3.2

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp100.visible=1
Ramp101.visible=1
Primitive52.visible=1
Else
Ramp100.visible=0
Ramp101.visible=0
Primitive52.visible=0
End if

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------
SolCallback(1)="bsTrough.SolOut"        'OK
SolCallback(2)="Sol2"             'Drop Target Bank 'OK
SolCallback(3)="Sol3"             'Drop Target Bank 'OK
SolCallback(10)="Sol10"             'GameOver
SolCallback(17)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"    'OK


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled and Flipperactive Then
     PlaySoundAtVol SoundFX("FlipperUp_Akiles",DOFFlippers), LeftFlipper, 1
     LeftFlipper.RotateToEnd:
     Else
     PlaySoundAtVol SoundFX("FlipperDown_Akiles",DOFFlippers), LeftFlipper, 1
     LeftFlipper.RotateToStart:
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled and Flipperactive Then
     PlaySoundAtVol SoundFX("FlipperUp_Akiles",DOFFlippers), RightFlipper, 1
     RightFlipper.RotateToEnd
     Else
     PlaySoundAtVol SoundFX("FlipperDown_Akiles",DOFFlippers), RightFlipper, 1
     RightFlipper.RotateToStart
    End If
 End Sub

Sub Sol2(enabled)
  if enabled then
    LeftDropTargetBank.DropSol_On
    P_DT1.Transz = DTUpperLimit + 5
    P_DT1.Transy = DTTransy + 5
    P_DT2.Transz = DTUpperLimit + 5
    P_DT2.Transy = DTTransy + 5
    P_DT3.Transz = DTUpperLimit + 5
    P_DT3.Transy = DTTransy + 5
  end if
End Sub

Sub Sol3(enabled)
  if enabled then
    RightDropTargetBank.DropSol_On
    P_DT4.Transz = DTUpperLimit + 5
    P_DT4.Transy = DTTransy + 5
    P_DT5.Transz = DTUpperLimit + 5
    P_DT5.Transy = DTTransy + 5
    P_DT6.Transz = DTUpperLimit + 5
    P_DT6.Transy = DTTransy + 5
  end if
End Sub

'GameOn reverted
Sub Sol10(enabled)
  Flipperactive = not enabled
  VpmNudge.SolGameOn(not enabled)
  if not Flipperactive then
    LeftFlipper.Rotatetostart
    RightFlipper.Rotatetostart
  end if
End Sub


'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough,obj,LeftDropTargetBank,RightDropTargetBank,cCaptive,Flipperactive

Sub Table1_Init
  vpminit me

  Flipperactive = False

    Controller.GameName=cGameName
    Controller.SplashInfoLine="NBA Mac" & vbNewLine & "created by mfuegemann"
    Controller.HandleKeyboard=False
    Controller.ShowTitle=0
    Controller.ShowFrame=0
    Controller.ShowDMDOnly=1
  'Controller.Hidden = 1      'enable to hide DMD if You use a B2S backglass

    Controller.HandleMechanics=0

    Controller.Run
    If Err Then MsgBox Err.Description
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch=4
    vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(LeftFlipper,RightFlipper,Bumper1,Bumper2,Bumper3)


    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 0,44,43,0,0,0,0,0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=1

  set LeftDropTargetBank = new cvpmDropTarget
    LeftDropTargetBank.InitDrop Array(DT1,DT2,DT3), Array(26,27,28)
    LeftDropTargetBank.InitSnd SoundFX("Targetdrop1",DOFContactors),SoundFX("TargetBankreset1",DOFContactors)
    LeftDropTargetBank.CreateEvents "LeftDropTargetBank"

  set RightDropTargetBank = new cvpmDropTarget
    RightDropTargetBank.InitDrop Array(DT4,DT5,DT6), Array(34,35,36)
    RightDropTargetBank.InitSnd SoundFX("Targetdrop1",DOFContactors),SoundFX("TargetBankreset1",DOFContactors)
    RightDropTargetBank.CreateEvents "RightDropTargetBank"

  Set cCaptive=New cvpmCaptiveBall
    cCaptive.InitCaptive CaptiveTrigger,CaptiveWall,CaptiveKicker,15
    cCaptive.Start
    cCaptive.ForceTrans = 0.9
    cCaptive.MinForce = 2.5
    cCaptive.CreateEvents "cCaptive"

End Sub

'------------------------------
'------  Trough Handler  ------
'------------------------------

 Sub Drain_Hit:bsTrough.addball me : PlaySoundAtVol "Drain5", Drain, 1: End Sub

'-------------------------------
'------  Keybord Handler  ------
'-------------------------------

Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
  End If

  if FreePlay and keycode = Startgamekey then
    vpmtimer.pulsesw 2
  end if

  if (keycode = AddCreditKey2) or (Keycode = keyInsertCoin3) or (Keycode = keyInsertCoin4) then
    vpmtimer.pulsesw 2
    Exit Sub
  End if

  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
  End If

  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub


'------------------------------
'------  Switch Handler  ------
'------------------------------
Sub SW46_Hit:vpmTimer.PulseSw 46:End Sub    'Leave Ballrelease

Dim Bumper1_Dir,Bumper2_Dir,Bumper3_Dir

Sub Bumper1_Hit
    vpmTimer.PulseSw 21
    PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1
    L16.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
  L16.State = 0
  Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit
    vpmTimer.PulseSw 22
    PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1
    L1.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer
  L1.State = 0
  Me.Timerenabled = 0
End Sub

Sub Bumper3_Hit
    vpmTimer.PulseSw 23
    PlaySoundAtVol SoundFX("Bumper",DOFContactors), ActiveBall, 1
    L36.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer
  L36.State = 0
  Me.Timerenabled = 0
End Sub



sub LOutlane_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(38)=1:End Sub
sub LOutlane_unhit:Controller.Switch(38)=0:End Sub
sub LInlane_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(41)=1:End Sub
sub LInlane_unhit:Controller.Switch(41)=0:End Sub
sub LInlane2_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(41)=1:End Sub
sub LInlane2_unhit:Controller.Switch(41)=0:End Sub

sub ROutlane_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(39)=1:End Sub
sub ROutlane_unhit:Controller.Switch(39)=0:End Sub
sub RInlane_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(42)=1:End Sub
sub RInlane_unhit:Controller.Switch(42)=0:End Sub
sub RInlane2_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(42)=1:End Sub
sub RInlane2_unhit:Controller.Switch(42)=0:End Sub

sub TopLeft_1_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(9)=1:End Sub
sub TopLeft_1_unhit:Controller.Switch(9)=0:End Sub
sub TopLeft_2_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(10)=1:End Sub
sub TopLeft_2_unhit:Controller.Switch(10)=0:End Sub
sub TopLeft_3_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(11)=1:End Sub
sub TopLeft_3_unhit:Controller.Switch(11)=0:End Sub

sub TopRight_1_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(12)=1:End Sub
sub TopRight_1_unhit:Controller.Switch(12)=0:End Sub
sub TopRight_2_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(13)=1:End Sub
sub TopRight_2_unhit:Controller.Switch(13)=0:End Sub
sub TopRight_3_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(14)=1:End Sub
sub TopRight_3_unhit:Controller.Switch(14)=0:End Sub
sub TopRight_4_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(15)=1:End Sub
sub TopRight_4_unhit:Controller.Switch(15)=0:End Sub

sub Toplane_hit:PlaySoundAtVol "Gate2", ActiveBall, 1:Controller.Switch(20)=1:End Sub
sub Toplane_unhit:Controller.Switch(20)=0:End Sub

Sub Spinner_Spin:vpmTimer.PulseSw 30:PlaySoundAtVol "soloff", ActiveBall, 1:End Sub

'10 Point Rubbers
Sub Rubber1_Hit:vpmTimer.PulseSw 19:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Rubber2_Hit:vpmTimer.PulseSw 19:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Rubber3_Hit:vpmTimer.PulseSw 19:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Rubber4_Hit:vpmTimer.PulseSw 19:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Rubber7_Hit:vpmTimer.PulseSw 19:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
'100 Point Rubbers
Sub Rubber5_Hit:vpmTimer.PulseSw 18:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub
Sub Rubber6_Hit:vpmTimer.PulseSw 18:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub

'Targets
Sub LeftTarget_Hit:MoveLeftTarget:vpmTimer.PulseSw 33:PlaySoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub
Sub RightTarget_Hit:MoveRightTarget:vpmTimer.PulseSw 33:PlaySoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub

Sub BehindLeftDT_Hit:MoveBehindLeftDT:vpmTimer.PulseSw 29:PlaySoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub
Sub BehindRightDT_Hit:MoveBehindRightDT:vpmTimer.PulseSw 37:PlaySoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub

Sub Target17_hit:MoveTarget17:vpmTimer.PulseSw 17:PlaySoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub

Sub Captive_Hit:MoveCaptiveTarget:vpmTimer.PulseSw 31:PlaySoundAtVol SoundFX("Target1",DOFTargets), ActiveBall, 1:End Sub


'--------------------------------
'------  Helper Functions  ------
'--------------------------------

'Sub Gate3_Hit:PlaySound "Gate5",0,1,0.13,0.1:End Sub



'Targets
Sub MoveLeftTarget
  P_LeftTarget.Transy = 5
  LeftTarget.Timerenabled = False
  LeftTarget.Timerenabled = True
End Sub
Sub LeftTarget_Timer
  LeftTarget.Timerenabled = False
  P_LeftTarget.Transy = 0
End Sub

Sub MoveBehindLeftDT
  P_BehindLeftDT.Transy = 5
  BehindLeftDT.Timerenabled = False
  BehindLeftDT.Timerenabled = True
End Sub
Sub BehindLeftDT_Timer
  BehindLeftDT.Timerenabled = False
  P_BehindLeftDT.Transy = 0
End Sub

Sub MoveRightTarget
  P_RightTarget.Transy = 5
  RightTarget.Timerenabled = False
  RightTarget.Timerenabled = True
End Sub
Sub RightTarget_Timer
  RightTarget.Timerenabled = False
  P_RightTarget.Transy = 0
End Sub

Sub MoveBehindRightDT
  P_BehindRightDT.Transy = 5
  BehindRightDT.Timerenabled = False
  BehindRightDT.Timerenabled = True
End Sub
Sub BehindRightDT_Timer
  BehindRightDT.Timerenabled = False
  P_BehindRightDT.Transy = 0
End Sub

Sub MoveTarget17
  P_Target17.Transy = 5
  Target17.Timerenabled = False
  Target17.Timerenabled = True
End Sub
Sub Target17_Timer
  Target17.Timerenabled = False
  P_Target17.Transy = 0
End Sub

Sub MoveCaptiveTarget
  P_Captive.Transy = 5
  Captive.Timerenabled = False
  Captive.Timerenabled = True
End Sub
Sub Captive_Timer
  Captive.Timerenabled = False
  P_Captive.Transy = 0
End Sub



'Release Gate
Dim GateSpeed
Sub ReleaseGateOpen_Hit
  if (not ReleaseGateOpen.Timerenabled) and (ActiveBall.vely < 0) then
    GateSpeed = 1.2
    ReleaseGateOpen.Timerenabled = True
  end if
End Sub

Sub ReleaseGateOpen_Timer
  P_ReleaseGate.rotz = P_ReleaseGate.rotz + GateSpeed
  if P_ReleaseGate.rotz > 19 then
    GateSpeed = -1.2
  end if
  if P_ReleaseGate.rotz <= 0 then
    P_ReleaseGate.rotz = 0
    ReleaseGateOpen.Timerenabled = False
  end if
end Sub


'Flipper and DT Primitives
Const DTUpperLimit = -1
Const DTBottomLimit = -47
Const DTTransY = -4
Const DownSpeed = 15
Const UpSpeed = 10

sub FlipperTimer_Timer()
    pleftFlipper001.objrotz=leftFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
    prightFlipper1.objrotz=rightFlipper.CurrentAngle
  pleftFlipper.objrotz=leftFlipper.CurrentAngle
  prightFlipper.objrotz=rightFlipper.CurrentAngle

  if DT1.isdropped then
    P_DT1.Transy = DTTransy
    if P_DT1.Transz > DTBottomLimit then
      P_DT1.Transz = P_DT1.Transz - DownSpeed
      if P_DT1.Transz < DTBottomLimit then
        P_DT1.Transz = DTBottomLimit
      end if
    end if
  else
    if P_DT1.Transz < DTUpperLimit then
      P_DT1.Transz = P_DT1.Transz + UpSpeed
    else
      P_DT1.Transz = DTUpperLimit
      P_DT1.Transy = 0
    end if
  end if

  if DT2.isdropped then
    P_DT2.Transy = DTTransy
    if P_DT2.Transz > DTBottomLimit then
      P_DT2.Transz = P_DT2.Transz - DownSpeed
      if P_DT2.Transz < DTBottomLimit then
        P_DT2.Transz = DTBottomLimit
      end if
    end if
  else
    if P_DT2.Transz < DTUpperLimit then
      P_DT2.Transz = P_DT2.Transz + UpSpeed
    else
      P_DT2.Transz = DTUpperLimit
      P_DT2.Transy = 0
    end if
  end if

  if DT3.isdropped then
    DT3_Cover.isdropped = True
    P_DT3.Transy = DTTransy
    if P_DT3.Transz > DTBottomLimit then
      P_DT3.Transz = P_DT3.Transz - DownSpeed
      if P_DT3.Transz < DTBottomLimit then
        P_DT3.Transz = DTBottomLimit
      end if
    end if
  else
    DT3_Cover.isdropped = False
    if P_DT3.Transz < DTUpperLimit then
      P_DT3.Transz = P_DT3.Transz + UpSpeed
    else
      P_DT3.Transz = DTUpperLimit
      P_DT3.Transy = 0
    end if
  end if

  if DT4.isdropped then
    DT4_Cover.isdropped = True
    P_DT4.Transy = DTTransy
    if P_DT4.Transz > DTBottomLimit then
      P_DT4.Transz = P_DT4.Transz - DownSpeed
      if P_DT4.Transz < DTBottomLimit then
        P_DT4.Transz = DTBottomLimit
      end if
    end if
  else
    DT4_Cover.isdropped = False
    if P_DT4.Transz < DTUpperLimit then
      P_DT4.Transz = P_DT4.Transz + UpSpeed
    else
      P_DT4.Transz = DTUpperLimit
      P_DT4.Transy = 0
    end if
  end if

  if DT5.isdropped then
    P_DT5.Transy = DTTransy
    if P_DT5.Transz > DTBottomLimit then
      P_DT5.Transz = P_DT5.Transz - DownSpeed
      if P_DT5.Transz < DTBottomLimit then
        P_DT5.Transz = DTBottomLimit
      end if
    end if
  else
    if P_DT5.Transz < DTUpperLimit then
      P_DT5.Transz = P_DT5.Transz + UpSpeed
    else
      P_DT5.Transz = DTUpperLimit
      P_DT5.Transy = 0
    end if
  end if

  if DT6.isdropped then
    P_DT6.Transy = DTTransy
    if P_DT6.Transz > DTBottomLimit then
      P_DT6.Transz = P_DT6.Transz - DownSpeed
      if P_DT6.Transz < DTBottomLimit then
        P_DT6.Transz = DTBottomLimit
      end if
    end if
  else
    if P_DT6.Transz < DTUpperLimit then
      P_DT6.Transz = P_DT6.Transz + UpSpeed
    else
      P_DT6.Transz = DTUpperLimit
      P_DT6.Transy = 0
    end if
  end if

end sub

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

    NFadeL 1, L1   'Bumper2
    NFadeL 16, L16 'Bumper 1
    NFadeL 36, L36 'Bumper 3
    NFadeL 6, L6
    NFadeL 7, L7
    NFadeL 8, L8
    NFadeL 9, L9
    NFadeL 10, L10
    NFadeL 11, L11
    NFadeL 12, l12
    NFadeL 13, L13
    NFadeL 14, L14
    NFadeL 15, L15
    NFadeL 17, L17
    NFadeL 18, L18
    NFadeL 19, L19
    NFadeL 20, L20
    NFadeL 22, L22
    NFadeL 23, L23
    NFadeL 24, L24
    NFadeL 25, L25
    NFadeL 26, L26
    NFadeL 27, L27
    NFadeL 28, L28
    NFadeL 29, L29
    NFadeL 30, L30   'Using Extra Ball
    NFadeL 32, L32
    NFadeL 33, L33
    NFadeL 34, L34
    NFadeL 35, L35
    NFadeL 37, L37
    NFadeL 38, L38
    NFadeL 39, L39
    NFadeL 40, L40
    NFadeL 41, L41
    NFadeLm 42, L42
    NFadeLm 42, L42a
    NFadeL 42, L42b
    NFadeL 43, L43
    NFadeL 44, L44
    NFadeL 45, L45
    NFadeL 46, L46
    NFadeL 47, L47
    NFadeL 49, L49
    NFadeL 50, L50
    NFadeL 51, L51
    NFadeL 54, L54
    NFadeL 80, L80  'Extra ball



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
  PlaySound "Gate5", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

'Sub Spinner_Spin
'  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
'End Sub

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

'-----------------------------
'------ DIP switch menu ------
'-----------------------------
Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 0,200,"NBA DIP switches"
    .AddFrame 0,0,200,"Balls per game",&H00000001,Array("5 balls",0,"3 balls",&H00000001)'dip 1
    .AddFrame 0,50,200,"Coins per game",&H00000006,Array("1 coin   = 1 credit",&H00000004,"2 coins = 2 credits",&H00000002,"2 coins = 1 credit",&H00000006,"4 coins = 1 credit",0)'dip 2&3
    .AddLabel 0,130,200,30,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")

Sub Table1_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'DOF Mappings by Arngrim
'it is needed to separate the switches to have individual ids for catching effects for DOF
'E101 Left Slingshot
'E102 Right Slingshot
'E103 Left Outlane
'E104 Right Inlane
'E106 sw1TrigF Trigger Hit

'Ball size: Values above 50 will likely result in stuck balls.
'Const BallSize = 50
'Const BallMass = 1.1

'Const cGameName="granslam",UseSolenoids=2,UseLamps=0,UseGI=0, SCoin="coin"
Const cGameName="gransla4",UseSolenoids=2,UseLamps=0,UseGI=0, SCoin="coin"  'If myou want to use the 4 Players ROM

LoadVPM "01560000", "BALLY.VBS", 3.10

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
'SolCallback(1)  = "vpmSolSound ""GSBumper"","     'sLBumper
'SolCallback(2)  = "vpmSolSound ""GSBumper"","     'sRBumper
SolCallback(3)  = "bsSaucer.SolOut"                'sSaucer
SolCallback(4)  = "FAT_Reset"                   'Fly Away Target Reset
SolCallback(6)  = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(5)  = "bsTrough.SolOut"              'sBallRelease
'SolCallback(19) = "vpmNudge.SolGameOn"          'sEnable

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart
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

Dim bstrough, bsSaucer, dtbank

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
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch=15
    vpmNudge.Sensitivity=1
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)

  Set bsTrough = New cvpmBallStack ' Trough handler
    bsTrough.InitSw 0,8,0,0,0,0,0,0
    bsTrough.InitKick BallRelease, 90, 5
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 1

  Set bsSaucer = New cvpmBallStack
    bsSaucer.InitSaucer sw32,32,225,10
    bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    set dtbank=new cvpmdroptarget
    dtbank.initdrop array(sw1dt, sw2dt, sw3dt, sw4dt, sw5dt), array(1, 2, 3, 4, 5)
    dtbank.initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

  If keycode = RightFlipperKey Then Controller.Switch(7) = 1
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)

  If keycode = RightFlipperKey Then Controller.Switch(7) = 0
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'**********************************************************************************************************

'Drains and Kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"Drain" , Drain, 1: End Sub
Sub sw32_Hit: bsSaucer.AddBall 0 : End Sub

'Rollovers Upper
 Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
 Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
 Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
 Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

'Rollovers Base Pads
 Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw37_unHit:Controller.Switch(37)=0:End Sub
 Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw38_unHit:Controller.Switch(38)=0:End Sub
 Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw39_unHit:Controller.Switch(39)=0:End Sub
 Sub sw40_Hit:Controller.Switch(40) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw40_unHit:Controller.Switch(40)=0:End Sub

'Rollovers Lower
 Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:DOF 103, DOFOn:End Sub
 Sub sw31_UnHit:Controller.Switch(31) = 0:DOF 103, DOFOff:End Sub
 Sub sw31B_Hit:Controller.Switch(31) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:DOF 104, DOFOn:End Sub
 Sub sw31B_UnHit:Controller.Switch(31) = 0:DOF 104, DOFOff:End Sub
 Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
 Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

'Stand Ups
 Sub sw18_Hit:vpmTimer.PulseSw 18:PlaySoundAtVol "target", ActiveBall, 1:End Sub
 Sub sw19_Hit:vpmTimer.PulseSw 19:PlaySoundAtVol "target", ActiveBall, 1:End Sub
 Sub sw20_Hit:vpmTimer.PulseSw 20:PlaySoundAtVol "target", ActiveBall, 1:End Sub
 Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySoundAtVol "target", ActiveBall, 1:End Sub
 Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySoundAtVol "target", ActiveBall, 1:End Sub
 Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtVol "target", ActiveBall, 1:End Sub
 Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol "target", ActiveBall, 1:End Sub
 Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol "target", ActiveBall, 1:End Sub
 Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtVol "target", ActiveBall, 1:End Sub

'Spinners
Sub sw35_Spin:vpmTimer.PulseSw 35 : playsoundAtVol"fx_spinner" , sw35, 1: End Sub

'Gate Trigger
Sub sw36_Hit:vpmTimer.PulseSw 36 : playsoundAtVol"Gate" , sw36, 1: End Sub

'Scoring Rubbers
'Slings (Only Rubbers with switches in this one - no slingshots)
 Sub LeftSlingShot_Hit:vpmTimer.PulseSw 34:DOF 101, DOFPulse:End Sub
 Sub RightSlingShot_Hit:vpmTimer.PulseSw 34:DOF 102, DOFPulse:End Sub
 Sub sw34c_Hit:vpmTimer.PulseSw 34 :End Sub


'***Bumpers
Sub Bumper1_Hit: vpmTimer.PulseSw 13 : PlaySoundAtVol SoundFX("bumper1",DOFContactors), ActiveBall, 1
    L4.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
  L4.State = 0
  Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit: vpmTimer.PulseSw 14 : PlaySoundAtVol SoundFX("bumper2",DOFContactors), ActiveBall, 1
    L16.State = 1
  Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer
  L16.State = 0
  Me.Timerenabled = 0
End Sub


'Sub PlayBumperSound()
'   Select Case Int(Rnd*3)+1
'     Case 1 : PlaySound SoundFX("bumper1",DOFContactors)
'     Case 2 : PlaySound SoundFX("bumper2",DOFContactors)
'   End Select
'End Sub



'***Fly Away Targets

Dim sw1Pos, sw1Dir, sw1Speed, sw1State, sw1Rot
Dim sw2Pos, sw2Dir, sw2Speed, sw2State, sw2Rot
Dim sw3Pos, sw3Dir, sw3Speed, sw3State, sw3Rot
Dim sw4Pos, sw4Dir, sw4Speed, sw4State, sw4Rot
Dim sw5Pos, sw5Dir, sw5Speed, sw5State, sw5Rot

sw1Pos = 0  ' here the values will be from 90 to -90 - degrees
sw1Dir = 2  ' here the values will be 2 or -2, this is the speed and direction
sw1Speed = 6  ' this is the speed
sw1State = 0  'state of target 0 - down 1 - up
sw1Rot = 0  'here is the Rotation angle destination for the target
sw2Pos = 0  ' here the values will be from 90 to -90 - degrees
sw2Dir = 2 ' here the values will be 2 or -2, this is the speed and direction
sw2Speed = 6  ' this is the speed
sw2State = 0  'state of target 0 - down 1 - up
sw2Rot = 0  'here is the Rotation angle destination for the target
sw3Pos = 0  ' here the values will be from 90 to -90 - degrees
sw3Dir = 2  ' here the values will be 2 or -2, this is the speed and direction
sw3Speed = 6 ' this is the speed
sw3State = 0  'state of target 0 - down 1 - up
sw3Rot = 0  'here is the Rotation angle destination for the target
sw4Pos = 0  ' here the values will be from 90 to -90 - degrees
sw4Dir = 2  ' here the values will be 2 or -2, this is the speed and direction
sw4Speed = 6  ' this is the speed
sw4State = 0  'state of target 0 - down 1 - up
sw4Rot = 0  'here is the Rotation angle destination for the target
sw5Pos = 0  ' here the values will be from 90 to -90 - degrees
sw5Dir = 2  ' here the values will be 2 or -2, this is the speed and direction
sw5Speed = 6  ' this is the speed
sw5State = 0  'state of target 0 - down 1 - up
sw5Rot = 0  'here is the Rotation angle destination for the target

Sub sw1TrigF_Hit()
    If activeball.vely < 0 then
       If sw1State = 0 then
          sw1Dir = -2
          sw1Rot = -90
          sw1Time.enabled = 1
          sw1State = 1
          dtbank.hit 1
          PlaySoundAtVol "FATClick", sw1, 1
      DOF 106, DOFPulse
        End If
    End If
End Sub

Sub sw1TrigB_Hit()
    If activeball.vely > 0 then
       If sw1State = 0 then
          sw1Dir = 2
          sw1Rot = 90
          sw1Time.enabled = 1
          PlaySoundAtVol "Target", sw1, 1
        End If
    End If
End Sub

Sub sw2TrigF_Hit()
    If activeball.vely < 0 then
       If sw2State = 0 then
          sw2Dir = -2
          sw2Rot = -90
          sw2Time.enabled = 1
          sw2State = 1
          dtbank.hit 2
          PlaySoundAtVol "FATClick", sw2, 1
        End If
    End If
End Sub

Sub sw2TrigB_Hit()
    If activeball.vely > 0 then
       If sw2State = 0 then
          sw2Dir = 2
          sw2Rot = 90
          sw2Time.enabled = 1
          PlaySoundAtVol "Target", sw2, 1
        End If
    End If
End Sub

Sub sw3TrigF_Hit()
    If activeball.vely < 0 then
       If sw3State = 0 then
          sw3Dir = -2
          sw3Rot = -90
          sw3Time.enabled = 1
          sw3State = 1
          dtbank.hit 3
          PlaySoundAtVol "FATClick", sw3, 1
        End If
    End If
End Sub

Sub sw3TrigB_Hit()
    If activeball.vely > 0 then
       If sw3State = 0 then
          sw3Dir = 2
          sw3Rot = 90
          sw3Time.enabled = 1
          PlaySoundAtVol "Target", sw3, 1
        End If
    End If
End Sub

Sub sw4TrigF_Hit()
    If activeball.vely < 0 then
       If sw4State = 0 then
          sw4Dir = -2
          sw4Rot = -90
          sw4Time.enabled = 1
          sw4State = 1
          dtbank.hit 4
          PlaySoundAtVol "FATClick", sw4, 1
        End If
    End If
End Sub

Sub sw4TrigB_Hit()
    If activeball.vely > 0 then
       If sw4State = 0 then
          sw4Dir = 2
          sw4Rot = 90
          sw4Time.enabled = 1
          PlaySoundAtVol "Target", sw4, 1
        End If
    End If
End Sub

Sub sw5TrigF_Hit()
    If activeball.vely < 0 then
       If sw5State = 0 then
          sw5Dir = -2
          sw5Rot = -90
          sw5Time.enabled = 1
          sw5State = 1
          dtbank.hit 5
          PlaySoundAtVol "FATClick", sw5, 1
        End If
    End If
End Sub

Sub sw5TrigB_Hit()
    If activeball.vely > 0 then
       If sw5State = 0 then
          sw5Dir = 2
          sw5Rot = 90
          sw5Time.enabled = 1
          PlaySoundAtVol "Target", sw5, 1
        End If
    End If
End Sub

Sub sw1Time_Timer()
  If sw1Dir = 2 Then
      sw1Pos =sw1Pos + sw1Speed
    If sw1Pos > sw1Rot Then
      sw1Pos = sw1Rot
      sw1Dir = -2
            sw1Rot = 0
    End If
  End If
  If sw1Dir = -2 Then
      sw1Pos =sw1Pos - sw1Speed
    If sw1Pos < sw1Rot Then
      sw1Pos = sw1Rot
      sw1Time.Enabled = 0
    End If
  End If
  If sw1Pos =0 Then
    sw1Time.Enabled = 0
  End If
  sw1.Rotx = sw1Pos
End Sub

Sub sw2Time_Timer()
  If sw2Dir = 2 Then
      sw2Pos =sw2Pos + sw2Speed
    If sw2Pos > sw2Rot Then
      sw2Pos = sw2Rot
      sw2Dir = -2
            sw2Rot = 0
    End If
  End If
  If sw2Dir = -2 Then
      sw2Pos =sw2Pos - sw2Speed
    If sw2Pos < sw2Rot Then
      sw2Pos = sw2Rot
      sw2Time.Enabled = 0
    End If
  End If
  If sw2Pos =0 Then
    sw2Time.Enabled = 0
  End If
  sw2.Rotx = sw2Pos
End Sub

Sub sw3Time_Timer()
  If sw3Dir = 2 Then
      sw3Pos =sw3Pos + sw3Speed
    If sw3Pos > sw3Rot Then
      sw3Pos = sw3Rot
      sw3Dir = -2
            sw3Rot = 0
    End If
  End If
  If sw3Dir = -2 Then
      sw3Pos =sw3Pos - sw3Speed
    If sw3Pos < sw3Rot Then
      sw3Pos = sw3Rot
      sw3Time.Enabled = 0
    End If
  End If
  If sw3Pos =0 Then
    sw3Time.Enabled = 0
  End If
  sw3.Rotx = sw3Pos
End Sub

Sub sw4Time_Timer()
  If sw4Dir = 2 Then
      sw4Pos =sw4Pos + sw4Speed
    If sw4Pos > sw4Rot Then
      sw4Pos = sw4Rot
      sw4Dir = -2
            sw4Rot = 0
    End If
  End If
  If sw4Dir = -2 Then
      sw4Pos =sw4Pos - sw4Speed
    If sw4Pos < sw4Rot Then
      sw4Pos = sw4Rot
      sw4Time.Enabled = 0
    End If
  End If
  If sw4Pos =0 Then
    sw4Time.Enabled = 0
  End If
  sw4.Rotx = sw4Pos
End Sub

Sub sw5Time_Timer()
  If sw5Dir = 2 Then
      sw5Pos =sw5Pos + sw5Speed
    If sw5Pos > sw5Rot Then
      sw5Pos = sw5Rot
      sw5Dir = -2
            sw5Rot = 0
    End If
  End If
  If sw5Dir = -2 Then
      sw5Pos =sw5Pos - sw5Speed
    If sw5Pos < sw5Rot Then
      sw5Pos = sw5Rot
      sw5Time.Enabled = 0
    End If
  End If
  If sw5Pos =0 Then
    sw5Time.Enabled = 0
  End If
  sw5.Rotx = sw5Pos
End Sub

Sub FAT_Reset(Enabled)
    dtbank.DropSol_On
    if sw5State = 1 then
        sw5Dir = 2
        sw5State = 0
        sw5Time.Enabled = 1
        PlaySoundAtVol "FATClick", sw5, 1
    End If
    if sw4State = 1 then
        sw4Dir = 2
        sw4State = 0
        sw4Time.Enabled = 1
        PlaySoundAtVol "FATClick", sw4, 1
    End If
    if sw3State = 1 then
        sw3Dir = 2
        sw3State = 0
        sw3Time.Enabled = 1
        PlaySoundAtVol "FATClick", sw3, 1
    End If
    if sw2State = 1 then
        sw2Dir = 2
        sw2State = 0
        sw2Time.Enabled = 1
        PlaySoundAtVol "FATClick", sw2, 1
    End If
    if sw1State = 1 then
        sw1Dir = 2
        sw1State = 0
        sw1Time.Enabled = 1
        PlaySoundAtVol "FATClick", sw1, 1
    End If
End Sub
'****End of Fly Away Targets


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
     NFadeL 2, l2
     NFadeL 3, l3

     NFadeL 5, l5

'     Flash 1, F1
'     Flashm 5, l5u, "gib", "gib", "wf_b", 0
'     Flash 5, l5o, "gib", "gib", "gib", 0
'     Flashm 6, l6u, "gib", "gib", "wf_b", 0
'     Flash 6, l6o, "gib", "gib", "gib", 0

     NFadeL 6, l6

    NFadeL 7, l7
     NFadeL 8, l8
     NFadeL 9, l9
     NFadeL 10, l10
     NFadeL 11, l11
     NFadeL 12, l12
     NFadeL 13, l13
     NFadeL 14, l14
     NFadeLm 15, l15B
     NFadeL 15, l15
     NFadeL 18, l18
     NFadeL 19, l19

     NFadeL 21, l21
     NFadeL 22, l22
'     Flashm 21, l21u, "gib", "gib", "wf_b", 0
'     Flash 21, l21o, "gib", "gib", "gib", 0
'     Flashm 22, l22u, "gib", "gib", "wf_b", 0
'     Flash 22, l22o, "gib", "gib", "gib", 0
     NFadeL 23, l23
     NFadeL 24, l24
     NFadeL 25, l25
     NFadeL 26, l26
     NFadeL 27, l27
     NFadeL 29, l29
     NFadeL 30, l30
     NFadeL 31, l31
     NFadeL 34, l34
     NFadeL 35, l35

     NFadeL 37, l37

'     Flashm 37, l37u, "gib", "gib", "wf_b", 0
'     Flash 37, l37o, "gib", "gib", "gib", 0
     NFadeL 38, l38
     NFadeL 39, l39
     NFadeL 40, l40
     NFadeL 41, l41
     NFadeL 42, l42
     NFadeL 43, l43
     Flash 44, F44   'Credit Indicator
    'NFadeL 44, l44   'Credit Indicator
     NFadeL 45, l45
     NFadeL 46, l46
     NFadeL 47, l47
     NFadeL 50, l50
     NFadeL 51, l51

     NFadeL 53, l53

 '    Flashm 53, l53u, "gib", "gib", "wf_b", 0
 '    FlashA 53, l53o, "gib", "gib", "gib", 0
     NFadeL 54, l54
     NFadeL 55, l55
     NFadeL 56, l56
     NFadeL 57, l57
     NFadeL 58, l58
     NFadeL 59, l59
     NFadeL 61, l61
     NFadeL 62, l62
     NFadeL 63, l63
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

'*********************************************************************
'Bally Grand Slam
'Thanks Mike da Spike
'*********************************************************************
Sub editDips
      Dim vpmDips : Set vpmDips = New cvpmDips
      With vpmDips
         .AddForm 315, 450,"Grand Slam - DIP Switches"
         .AddFrame   2,   0, 160,"Coin Chute 1 (Coins/Credit)",&H0000001F,Array("2/1",&H0000000B,"1/1",&H00000000,"1/2",&H00000001)                                                              'dip 1-5
         .AddFrame   2,  62, 160,"Coin Chute 2 (Coins/Credit)",&H00001F00,Array("2/1",&H00000B00,"1/1",&H00000000,"1/2",&H00000100)                                                              'dip 9-13
         .AddFrame   2, 122, 160,"Coin Chute 3 (Coins/Credit)",&H000F0000,Array("Same as Chute#1",&H00000000,"1/1",&H00010000,"1/2",&H00020000)                                                  'dip 17-20
         .AddFrame   2, 184, 160,"Balls per game",&HC0000000,Array("2",&HC0000000,"3",0,"4",&H80000000,"5",&H40000000)                                                                           'dip 31&32
         .AddFrame   2, 260, 160,"Extra balls",&H00004000,Array("more then 1 extra ball per game",0,"1 extra ball per game",&H00004000)                                                          'dip 15
         .AddFrame 187,   0, 160,"Flyaway value lites on next ball",32768,Array("lites will not come on",0,"lites will come on",32768)                                                           'dip 16
         .AddFrame 187,  46, 160,"S&L return lane lite",&H00800000,Array("S&L lite not be tied",0,"S&L lite be tied together",&H00800000)                                                        'dip 24
         .AddFrame 187,  92, 160,"Replays per game",&H10000000,Array("one per game",0,"unlimited",&H10000000)                                                                                    'dip 29
         .AddFrame 187, 138, 160,"Outhole bonus score",&H20000000,Array("10K for each run",0,"5K for each run",&H20000000)                                                                       'dip 30
         .AddFrame 187, 184, 160,"Flyaway 25K lite",&H00000040,Array("off at start",0,"on at start",&H00000040)                                                                                  'dip 7
         .AddFrame 187, 230, 160,"Single & Double lite",&H00000020,Array("alternate",0,"always on",&H00000020)                                                                                   'dip 6
         .AddFrame 372,   0, 160,"Runs to beat advances",&H00000080,Array("always on",0, "step off by 10",&H00000080)                                                                            'dip 8
         .AddFrame 372,  46, 160,"Runs to beat",&H00700000,Array("20",&H00700000, "25",&H00600000, "30",&H00500000, "35",&H004000000, "40",&H00300000, "45",&H00200000, "50",&H00100000, "55",0) 'dip 21-23
         .AddFrame 372, 176, 160,"Flyaway special",&H00002000,Array("alternate",0,"always on",&H00002000)                                                                                        'dip 14
         .AddFrame 372, 222, 160,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"40 credits",&H03000000)                                      'dip 25&26
         .AddChk 187, 279, 160,Array("Credits displayed",&H04000000)                                                                                                                             'dip 27
         .AddChk 187, 291, 160,Array("Match",&H08000000)                                                                                                                                         'dip 28
         .AddLabel 89,342,385,40,"Set selftest position 16 & 17 for highscore feature, 19 for High score to date"
         .AddLabel 132,362,300,40,"After hitting OK, press F3 to reset game with new settings."
         .ViewDips
      End With
End Sub

Set vpmShowDips = GetRef("editDips")
'High Score feature :
'Self test position: 16 - 17
'Reply               3    3
'Extra ball          2    2
'Novelty             1    1
'No award            0    0

'High Score to date :
'Self test position:  19
'No Award             0
'One Credit           1
'Two credits          2
'Three Credits        3
'*********************************************************************

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
  FlipperLSh1.RotZ = LeftFlipper1.currentangle
  Primitive85.roty = LeftFlipper1.currentangle +155
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


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


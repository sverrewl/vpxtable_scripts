option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const UseSolenoids=2,UseLamps=1,UseGI=0,SCoin="coin3"

BallApronCard = 3     'Set to Either 3 or 5 depending on how many balls you are using

LoadVPM "01560000","juegos.vbs",3.2

Dim BallApronCard

Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp160.visible=1
Ramp150.visible=1
Primitive130.visible=1
Else
Ramp160.visible=0
Ramp150.visible=0
Primitive130.visible=0
End if

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------
Const cGameName="america"   'name of ROM file to be used
Const ApronMatch = True     'show Match animation on Apron


'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(7)= "Sol7"              'SolGameOn
SolCallback(8)= "Sol8"              '"B" insert light
SolCallback(16)= "bsTrough.SolOut"
SolCallback(13)= "DropTargetBank.SolDropUp" 'Drop Targets
SolCallback(14)= "SolVariTargetReset"
SolCallback(15)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'Dim Flipperactive

Sub SolLFlipper(Enabled)
     If Enabled Then
       PlaySoundAtVol SoundFX("Flipperup",DOFFlippers), LeftFlipper, 1: LeftFlipper.RotateToEnd
     Else
       PlaySoundAtVol SoundFX("Flipperdown",DOFFlippers), LeftFlipper, 1: LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
       PlaySoundAtVol SoundFX("Flipperup",DOFFlippers), RightFlipper, 1: RightFlipper.RotateToEnd:URightFlipper.RotateToEnd
     Else
       PlaySoundAtVol SoundFX("Flipperdown",DOFFlippers), RightFlipper, 1: RightFlipper.RotateToStart:URightFlipper.RotateToStart
     End If
End Sub


'**********************************************************************************************************
'Solenoid Controlled toys
'**********************************************************************************************************

Sub Sol7(enabled)
  VpmNudge.SolGameOn enabled
' Flipperactive = 1
end sub

Sub Sol8(enabled)
  LampSol8.state=enabled
end sub

'Playfield GI
'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough, DropTargetBank

Sub Table1_Init
    vpminit me
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
    .hidden = 1    'enable to hide DMD if You use a B2S backglass
        .Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch=30
    vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 0,25,0,0,0,0,0,0    '0.1 = Switch 0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=1

  set DropTargetBank = new cvpmDropTarget
    DropTargetBank.InitDrop Array(DT1,DT2,DT3), Array(6,7,8)
    DropTargetBank.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
    DropTargetBank.CreateEvents "DropTargetBank"

    vpmMapLights AllLights

  VariPos = 2
  SolVariTargetReset True

  Rubber1Wall2.isdropped = True
  Rubber2Wall2.isdropped = True
  Rubber4Wall2.isdropped = True

'Create Captive Balls
  CaptiveKicker.CreateBall
  CaptiveKicker.Kick 0,0

Controller.Games(cGameName).Settings.Value("volume") = 0

End Sub


'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************


Sub Table1_KeyDown(ByVal Keycode)
  If keycode = LeftMagnaSave Then bLutActive = True
  If keycode = RightMagnaSave Then
    If bLutActive Then NextLUT: End If
    End If
    If KeyCode=LeftFlipperKey Then Controller.Switch(84)=1
  If KeyCode=RightFlipperKey Then Controller.Switch(82)=1
  If KeyCode=PlungerKey Then Plunger.Pullback:playsound"plungerpull"
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
    If keycode = LeftMagnaSave Then bLutActive = False
  If KeyCode=LeftFlipperKey Then Controller.Switch(84)=0
  If KeyCode=RightFlipperKey Then Controller.Switch(82)=0
  If KeyCode=PlungerKey Then PlaySound"plunger":Plunger.Fire
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 13
    PlaySoundAtVol SoundFX("Slingshot_Akiles",DOFContactors), ActiveBall, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
' gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 14
    PlaySoundAtVol SoundFX("Slingshot_Akiles",DOFContactors),ActiveBall, 1
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



'------------------------------
'------  Trough Handler  ------
'------------------------------

Sub Drain_Hit:bsTrough.addball me : playsound"drain5" : End Sub

'---------------------------
'------  Vari Target  ------
'---------------------------

Dim Resistance,VariPos,VariPos2
Resistance=5

Sub VT_1_Hit
  if not VT_1_Wall.isdropped then
    If ActiveBall.VelY<0 Then
      PlaySoundAtVol "soloff", ActiveBall, 1
      VariTargetResetTimer.enabled = False
      VariPos = 2
      VT_1_Wall.IsDropped = True
      VT_2_Wall.IsDropped = False
      Controller.Switch(24) = 1
      ActiveBall.VelY = ActiveBall.VelY + Resistance
    End if
  end if
End Sub

Sub VT_2_Hit
  if not VT_2_Wall.isdropped then
    If ActiveBall.VelY<0 Then
      PlaySoundAtVol "soloff", ActiveBall, 1
      VariTargetResetTimer.enabled = False
      VariPos = 3
      VT_2_Wall.IsDropped = True
      VT_3_Wall.IsDropped = False
      Controller.Switch(24) = 0
      Controller.Switch(23) = 1
      ActiveBall.VelY = ActiveBall.VelY + Resistance
    End If
  End If
End Sub

Sub VT_3_Hit
  if not VT_3_Wall.isdropped then
    If ActiveBall.VelY<0 Then
      PlaySoundAtVol "soloff", ActiveBall, 1
      VariTargetResetTimer.enabled = False
      VariPos = 4
      VT_3_Wall.IsDropped = True
      VT_4_Wall.IsDropped = False
      Controller.Switch(23) = 0
      Controller.Switch(22) = 1
      ActiveBall.VelY = ActiveBall.VelY + Resistance
    End If
  End If
End Sub

Sub VT_4_Hit
  if not VT_4_Wall.isdropped then
    If ActiveBall.VelY<0 Then
      PlaySoundAtVol "soloff", ActiveBall, 1
      VariTargetResetTimer.enabled = False
      VariPos = 5
      VT_4_Wall.IsDropped = True
      VT_5_Wall.IsDropped = False
      Controller.Switch(22) = 0
      Controller.Switch(21) = 1
      ActiveBall.VelY = ActiveBall.VelY + Resistance
    End if
  End If
End Sub

Sub VT_5_Hit
  if not VT_5_Wall.isdropped then
    If ActiveBall.VelY<0 Then
      PlaySoundAtVol "soloff", ActiveBall, 1
      VariTargetResetTimer.enabled = False
      VariPos = 6
      VT_5_Wall.IsDropped = True
      VT_6_Wall.IsDropped = False
      Controller.Switch(21) = 0
      Controller.Switch(20) = 1
      ActiveBall.VelY = ActiveBall.VelY + Resistance
    End If
  End If
End Sub

Sub VT_6_Hit
  if not VT_6_Wall.isdropped then
    If ActiveBall.VelY<0 Then
      PlaySoundAtVol "soloff", ActiveBall, 1
      VariTargetResetTimer.enabled = False
      VariPos = 7
      VT_6_Wall.IsDropped = True
      Controller.Switch(20) = 0
      Controller.Switch(19) = 1
      ActiveBall.VelY = ActiveBall.VelY + Resistance
    End If
  End If
End Sub

Sub SolVariTargetReset(Enabled)
  If Enabled and VariPos > 1 Then
    PlaySoundAtVol "solon", Drain, 1
    select case VariPos
      case 7: VT_6_Wall.IsDropped = False
          Controller.Switch(19)=0
          Controller.Switch(20)=1
      case 6: VT_6_Wall.IsDropped = True
          VT_5_Wall.IsDropped = False
          Controller.Switch(20)=0
          Controller.Switch(21)=1
      case 5: VT_5_Wall.IsDropped = True
          VT_4_Wall.IsDropped = False
          Controller.Switch(21)=0
          Controller.Switch(22)=1
      case 4: VT_4_Wall.IsDropped = True
          VT_3_Wall.IsDropped = False
          Controller.Switch(22)=0
          Controller.Switch(23)=1
      case 3: VT_3_Wall.IsDropped = True
          VT_2_Wall.IsDropped = False
          Controller.Switch(23)=0
          Controller.Switch(24)=1
      case 2: VT_1_Wall.IsDropped = False
          VT_2_Wall.IsDropped = True
          VT_3_Wall.IsDropped = True
          VT_4_Wall.IsDropped = True
          VT_5_Wall.IsDropped = True
          VT_6_Wall.IsDropped = True
          Controller.Switch(24)=0
    end select
    VariPos = VariPos - 1
  end if
End Sub

Dim TargetRotX, TargetTransY
Sub MoveVariTarget_Timer
  select case VariPos
    case 7: TargetRotX = -5
        TargetTransY = 95
    case 6: TargetRotX = -3
        TargetTransY = 75
    case 5: TargetRotX = -1
        TargetTransY = 40
    case 4: TargetRotX = 0
        TargetTransY = 10
    case 3: TargetRotX = 1
        TargetTransY = -25
    case 2: TargetRotX = 3
        TargetTransY = -55
    case 1: TargetRotX = 5
        TargetTransY = -85
  end select

  if P_VariTarget.Rotx < TargetRotx then        'reset Target
    P_VariTarget.Rotx = P_VariTarget.Rotx + 0.2
    if P_VariTarget.Rotx > TargetRotx then
      P_VariTarget.Rotx = TargetRotx
    end if
  end if
  if P_VariTarget.Rotx > TargetRotx then        'move Target
    P_VariTarget.Rotx = P_VariTarget.Rotx - 0.75
    if P_VariTarget.Rotx < TargetRotx then
      P_VariTarget.Rotx = TargetRotx
    end if
  end if
  if P_VariTarget.TransY < TargetTransY then      'move Target
    P_VariTarget.TransY = P_VariTarget.TransY + 15
    if P_VariTarget.TransY > TargetTransY then
      P_VariTarget.TransY = TargetTransY
    end if
  end if
  if P_VariTarget.TransY > TargetTransY then      'reset Target
    P_VariTarget.TransY = P_VariTarget.TransY - 5
    if P_VariTarget.TransY < TargetTransY then
      P_VariTarget.TransY = TargetTransY
    end if
  end if


End Sub


' Bumpers


Sub Bumper1_Hit:vpmTimer.PulseSw 15:PlaySoundAtVol SoundFX("Bumper",DOFContactors),ActiveBall,1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 16:PlaySoundAtVol SoundFX("Bumper",DOFContactors),ActiveBall,1:End Sub

' Rollover

sub LOutlane_hit:Controller.Switch(33)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
sub LOutlane_unhit:Controller.Switch(33)=0:End Sub
sub LInlane_hit:Controller.Switch(34)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
sub LInlane_unhit:Controller.Switch(34)=0:End Sub

sub ROutlane_hit:Controller.Switch(35)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
sub ROutlane_unhit:Controller.Switch(35)=0:End Sub
sub RInlane_hit:Controller.Switch(36)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
sub RInlane_unhit:Controller.Switch(36)=0:End Sub


sub TopRollover_Left_hit:Controller.Switch(10)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
sub TopRollover_Left_unhit:Controller.Switch(10)=0:End Sub
sub TopRollover_Right_hit:Controller.Switch(9)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
sub TopRollover_Right_unhit:Controller.Switch(9)=0:End Sub

Sub sw17_Spin:vpmTimer.PulseSw 17:PlaySoundAtVol "soloff", sw17, 1:End Sub

' Targets

sub sw1_hit:vpmTimer.PulseSw 1:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
sub sw2_hit:vpmTimer.PulseSw 2:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
sub sw3_hit:vpmTimer.PulseSw 3:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
sub sw4_hit:vpmTimer.PulseSw 4:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
sub sw5_hit:vpmTimer.PulseSw 5:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub

Sub sw37_hit:vpmTimer.PulseSw 37:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
Sub sw38_hit:vpmTimer.PulseSw 38:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
Sub sw39_hit:vpmTimer.PulseSw 39:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub

Sub sw12_hit:vpmTimer.PulseSw 12:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub
Sub sw11_hit:vpmTimer.PulseSw 11:PlaySoundAtVol SoundFX("Target",DOFTargets), ActiveBall, 1:End Sub

' Rubber Walls

Sub Rubber1_hit:vpmTimer.PulseSw 18:MoveRubber1:End Sub
Sub Rubber2_hit:vpmTimer.PulseSw 18:End Sub
Sub Rubber3_hit:vpmTimer.PulseSw 18:End Sub
Sub Rubber4_hit:vpmTimer.PulseSw 18:MoveRubber4:End Sub

' Gates

Sub Gate2_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub
Sub Gate3_Hit:PlaySoundAtVol "Gate5",ActiveBall, 1:End Sub
Sub Gate6_Hit:PlaySoundAtVol "Gate1",ActiveBall, 1:End Sub

'--------------------------------
'------  Helper Functions  ------
'--------------------------------

'Plunger Lane Gate
Dim GateSpeed
Sub ReleaseGateOpen_Hit
  if not ReleaseGateOpen.Timerenabled and (Activeball.vely < 0) then
    if Activeball.vely > -4 then
      Activeball.vely = -4
    end if
    GateSpeed = 1.2
    ReleaseGateOpen.Timerenabled = True
  end if
End Sub

Sub ReleaseGateOpen_Timer
  P_ReleaseGate.rotz = P_ReleaseGate.rotz + GateSpeed
  if P_ReleaseGate.rotz > 17 then
    GateSpeed = -1.2
  end if
  if P_ReleaseGate.rotz <= 0 then
    P_ReleaseGate.rotz = 0
    ReleaseGateOpen.Timerenabled = False
  end if
end Sub


'Rubber Switches
Sub MoveRubber1
  Rubber1Wall1.isdropped = True
  Rubber1Wall2.isdropped = False
  Rubber1.Timerenabled = False
  Rubber1.Timerenabled = True
End Sub

Sub Rubber1_Timer
  Rubber1.Timerenabled = False
  Rubber1Wall1.isdropped = False
  Rubber1Wall2.isdropped = True
End Sub

Sub MoveRubber4
  Rubber4Wall1.isdropped = True
  Rubber4Wall2.isdropped = False
  Rubber4.Timerenabled = False
  Rubber4.Timerenabled = True
End Sub

Sub Rubber4_Timer
  Rubber4.Timerenabled = False
  Rubber4Wall1.isdropped = False
  Rubber4Wall2.isdropped = True
End Sub



'additional Lamp Callback
Sub LampTimer_Timer

  if controller.lamp(97) then
    gate3.open = True
  else
    gate3.open = False
  end if

End Sub

'-------------------------------
'------  DIP Switch Menu  ------
'-------------------------------

Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 200,280,"America 1492 DIP switches"
    .AddFrame 0, 0,270,"Balls per game",&H00000400,Array("3 balls",0,"5 balls",&H00000400)
    .AddChk   5,50,270,Array("Match feature",32768)
    .AddFrame 0,75,270,"Replay threshold (Extra ball/credit/credit)",&H00000300,Array("3.0M/3.7M/4.5M points",&H00000300,"2.5M/3.0M/3.7M points",&H00000100,"1.8M/2.5M/3.0M points",&H00000200,"1.2M/1.8M/2.5M points",0)
    .AddChk   5,155,270,Array("Enable replay extra ball",&H00000008)

    .AddLabel 0,180,280,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(35)

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


Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 35) then
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

'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

    pleftFlipper001.objrotz=LeftFlipper.CurrentAngle
  pleftFlipper.objrotz=LeftFlipper.CurrentAngle
  prightFlipper.objrotz=RightFlipper.CurrentAngle
    pleftFlipper002.objrotz=RightFlipper.CurrentAngle
  purightFlipper.objrotz=urightFlipper.CurrentAngle

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
Sub RampSound2_Hit: PlaySoundAtVol "fx_metalrolling1", ActiveBall, 1: End Sub
Sub RampSound3_Hit: PlaySoundAtVol "fx_metalrolling1", ActiveBall, 1: End Sub

' Stop Ramps Sounds
Sub RampSound5_Hit: StopSound "fx_metalrolling1": End Sub
Sub RampSound6_Hit: StopSound "fx_metalrolling1": End Sub



' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

changeballcard
Sub changeballcard
  If BallApronCard = 3 Then
    Primitive51.Image="3-Balls"
    End If
  If BallApronCard = 5 Then
      Primitive51.Image="5-Balls"
    End If
End Sub

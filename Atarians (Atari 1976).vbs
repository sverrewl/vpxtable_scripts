' *********************************************************************************************************************************
' *********************************************************************************************************************************
' ***************************                  Atarians -- Atari, 1976           **************************************************
' ***************************        Original table by Ash and George Opperman   **************************************************
' ***************************         Original Release RC-1 (May 19, 2002) VP9   **************************************************
' ***************************           This table is loosely based from It      **************************************************
' ***************************                Rebuild by Wiesshund                **************************************************
' ***************************       NEW Playfield and plastics by Patrick2610    **************************************************
' *********************************************************************************************************************************
' *********************************************************************************************************************************




Option Explicit

Const BallMass = 1.7

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


LoadVPM "01200000", "ATARI1.VBS", 3.1

Const cGameName     = "atarians"
Const cCredits      = "Atarians Table by Wiesshund, Original by Ash"
Const UseSolenoids  = 2
Const UseLamps      = True
Const UseGI         = False
Const SSolenoidOn   = "solon"
Const SSolenoidOff  = "soloff"
Const SFlipperOn    = "FlipperUp"
Const SFlipperOff   = "FlipperDown"
Const SCoin         = "coin3"

'Solenoid Callbacks
'SolCallback(19)    = "LLFELIPPER"
'SolCallback(20)    = "LRFLIPPER"
SolCallback(18)   = "SolLeftGate"
SolCallback(17)   = "SolRightGate"
SolCallback(14) = "SolAtari"
SolCallback(8)  = "KickerEBOut"
SolCallback(3)  = "KickerBonusOut"
'SolCallback(16)    = "vpmSolSound SoundFX(""Jet3"",DOFContactors),"
'SolCallback(6)   = "vpmSolSound SoundFX(""Jet3"",DOFContactors),"
'SolCallback(4)   = "vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(9)    = "vpmSolSound SoundFX(""sling"",DOFContactors),"
'SolCallback(13)    = "vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(10)   = "bsTrough.SolOut"
SolCallback(1)    = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"




' ************************************** SOLENOID INTERACTIONS *************************************
' *** OUTLANE GATES ***

Sub SolLeftGate(Enabled)
If Enabled Then
LeftGate.RotateToEnd:PlaySoundAt SoundFX("fx_diverter", DOFFlippers), LeftGate
Else
LeftGate.RotateToStart:PlaySoundAt SoundFX("fx_diverter", DOFFlippers), LeftGate
End If
End Sub


Sub SolRightGate(Enabled)
If Enabled Then
RightGate.RotateToEnd: PlaySoundAt SoundFX("fx_diverter", DOFFlippers), RightGate
Else
RightGate.RotateToStart: PlaySoundAt SoundFX("fx_diverter", DOFFlippers), RightGate
End If
End Sub

' *** FLIPPERS ***

Sub LLFELIPPER
 If gameon=1 then
      If MacTilt=0 and lfon = 1 Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipperTop
        LeftFlipperTop.RotateToEnd
        LeftFlipperTopOn = 1
      End If

      If MacTilt=0 and lfon = 0 Then
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipperTop
        LeftFlipperTop.RotateToStart
        LeftFlipperTopOn = 0
      End If
end if
End Sub



Sub LRFLIPPER
if gameon=1 then
      If MacTilt=0 and rfon = 1 Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipperTop
        RightFlipperTop.RotateToEnd
        RightFlipperTopOn = 1
      End If

      If MacTilt=0 and Rfon = 0 Then
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipperTop
        RightFlipperTop.RotateToStart
        RightFlipperTopOn = 0
      End If
end If
End Sub

' **** TOP LANE OUTHOLES ****

Sub SolAtari(Enabled)
  If Enabled Then

If Controller.Switch(28) then Controller.Switch(28)=0: KickerA.kick 180,4: PlaySoundAt SoundFX("FX_Kickout", DOFContactors), KickerA: End If
If Controller.Switch(29) then Controller.Switch(29)=0: KickerT.kick 180,4: PlaySoundAt SoundFX("FX_Kickout", DOFContactors), KickerT: End If
If Controller.Switch(30) then Controller.Switch(30)=0: KickerE.kick 180,4: PlaySoundAt SoundFX("FX_Kickout", DOFContactors), KickerE: End If
If Controller.Switch(31) then Controller.Switch(31)=0: KickerR.kick 180,4: PlaySoundAt SoundFX("FX_Kickout", DOFContactors), KickerR: End If
If Controller.Switch(32) then Controller.Switch(32)=0: KickerI.kick 180,4: PlaySoundAt SoundFX("FX_Kickout", DOFContactors), KickerI: End If

  End If
End Sub

Sub KickerBonusOut(Enabled)
  If Enabled Then
  Controller.Switch(33)=0: KickerBonus.kick 170,4: PlaySoundAt SoundFX("FX_Kickout", DOFContactors), KickerBonus
  End If
End Sub

Sub KickerEBOut(Enabled)
  If Enabled Then
  Controller.Switch(34)=0: KickerEB.kick 190,4: PlaySoundAt SoundFX("FX_Kickout", DOFContactors), KickerEB
  End If
End Sub

' ************************************** END SOLENOID INTERACTIONS *************************************



' ************************************************************************************************
' ************************************** SLINGSHOTS **********************************************
' ************************************************************************************************
Dim SLPos,SRPos
Dim LSlung
Dim RSlung


Sub LeftSlingshotBounce_Slingshot()
    If MacTilt=0 Then
  If LSlung = 0 Then
  Lslung = 1
  LSling1.Visible = 0:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 1: LSling.TransZ = -27
  vpmTimer.PulseSw 26
  PlaySoundAt SoundFX ("fx_slingshot",DOFContactors),LSling
  SLPos = 0: Me.TimerEnabled = 1
  End If
    End If
End Sub

Sub LeftSlingshotBounceTop_Slingshot()
    If MacTilt=0 Then
  If LSlung = 0 Then
  Lslung = 1
  LSling1.Visible = 0:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 1: LSling.TransZ = -27
  vpmTimer.PulseSw 26
  PlaySoundAt SoundFX ("fx_slingshot",DOFContactors),LSling
  SLPos = 0: LeftSlingshotBounce.TimerEnabled = 1
  End If
    End If
End Sub

Sub LeftSlingshotBounceBot_Slingshot()
    If MacTilt=0 Then
  If LSlung = 0 Then
  Lslung = 1
  LSling1.Visible = 0:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 1: LSling.TransZ = -27
  vpmTimer.PulseSw 26
  PlaySoundAt SoundFX ("fx_slingshot",DOFContactors),LSling
  SLPos = 0: LeftSlingshotBounce.TimerEnabled = 1
  End If
    End If
End Sub


Sub LeftSlingshotBounce_Timer
    Select Case SLPos
        Case 2: LSling1.Visible = 0:LSling2.Visible = 0:LSling3.Visible = 1: LSling4.Visible = 0: LSling.TransZ = -17
    Case 3: LSling1.Visible = 0:LSling2.Visible = 1:LSling3.Visible = 0: LSling4.Visible = 0: LSling.TransZ = -8
    Case 4: LSling1.Visible = 1:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 0: LSling.TransZ = 0 :LSlung = 0:Me.TimerEnabled = 0
    End Select
    SLPos = SLPos + 1
End Sub


Sub RightSlingshotBounce_Slingshot()
  If RSlung = 0 Then
  RSlung = 1
  RSling1.Visible = 0:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 1: RSling.TransZ = -27
  vpmTimer.PulseSw 27
  PlaySoundAt SoundFX ("fx_slingshot",DOFContactors), RSling
  SRPos = 0: Me.TimerEnabled = 1
  End If
End Sub

Sub RightSlingshotBounceTop_Slingshot()
  If RSlung = 0 Then
  RSlung = 1
  RSling1.Visible = 0:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 1: RSling.TransZ = -27
  vpmTimer.PulseSw 27
  PlaySoundAt SoundFX ("fx_slingshot",DOFContactors), RSling
  SRPos = 0: RightSlingshotBounce.TimerEnabled = 1
  End If
End Sub

Sub RightSlingshotBounceBot_Slingshot()
  If RSlung = 0 Then
  RSlung = 1
  RSling1.Visible = 0:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 1: RSling.TransZ = -27
  vpmTimer.PulseSw 27
  PlaySoundAt SoundFX ("fx_slingshot",DOFContactors), RSling
  SRPos = 0: RightSlingshotBounce.TimerEnabled = 1
  End If
End Sub


Sub RightSlingshotBounce_Timer
    Select Case SRPos
        Case 2: RSling1.Visible = 0:RSling2.Visible = 0:RSling3.Visible = 1: RSling4.Visible = 0: RSling.TransZ = -17
    Case 3: RSling1.Visible = 0:RSling2.Visible = 1:RSling3.Visible = 0: RSling4.Visible = 0: RSling.TransZ = -8
    Case 4: RSling1.Visible = 1:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 0: RSling.TransZ = 0 :RSlung = 0:Me.TimerEnabled = 0
    End Select
    SRPos = SRPos + 1
End Sub

Sub Slingshots_Init
  LSling1.Visible = 1:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 0: LSling.TransZ = 0
  RSling1.Visible = 1:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 0: RSling.TransZ = 0
End Sub

' ************************************************************************************************
' ************************************ END SLINGSHOTS ********************************************
' ************************************************************************************************


' ********************** HIT TARGETS *********************************
Sub Target2_Hit:vpmTimer.PulseSw 10
PlaySoundAt SoundFX("FX_HitTarget", DOFTargets), Target2
End Sub         '10 OK

Sub Target3_Hit:vpmTimer.PulseSw 11 '11 MAY Be Reversed with Switch 16
PlaySoundAt SoundFX("FX_HitTarget", DOFTargets), Target3
End Sub

Sub Target9_Hit:vpmTimer.PulseSw 17
PlaySoundAt SoundFX("FX_HitTarget", DOFTargets), Target9
End Sub         '17 OK


Sub Target8_Hit:vpmTimer.PulseSw 16
PlaySoundAt SoundFX("FX_HitTarget", DOFTargets), Target8
End Sub
' ********************************************************************


' ****************************** ROLLOVER TARGETS ********************
'1 Coin1
'2 Coin2
'3 Start Button
'4 Tilt lamp Light125
Sub Trigger1_Hit:vpmTimer.PulseSw 9:End Sub       '9 OK



Sub Trigger4_Hit:vpmTimer.PulseSw 12:End Sub      '12 OK

Sub Trigger5_Hit:vpmTimer.PulseSw 13:End Sub      '13 OK

Sub Trigger6_Hit:vpmTimer.PulseSw 14:End Sub      '14 OK

Sub Trigger7_Hit:vpmTimer.PulseSw 15:End Sub      '15 OK

Sub TriggerAdvanceCenter_Hit:vpmTimer.PulseSw 11:End Sub'16 MAY Be Reversed with Switch 11

Sub TriggerTopLeft_Hit:vpmTimer.PulseSw 18:End Sub    '18 OK

Sub TriggerTopRight_Hit:vpmTimer.PulseSw 18:End Sub

Sub TriggerOpenLeft_Hit:vpmTimer.PulseSw 20:End Sub   '20 OK

Sub TriggerOpenRight_Hit:vpmTimer.PulseSw 21:End Sub  '21 OK

Sub TriggerCloseLeft_Hit:vpmTimer.PulseSw 22:End Sub  '22 OK

Sub TriggerCloseRight_Hit:vpmTimer.PulseSw 22:End Sub

' ***********************************************************************

' ***************** RUBBERBAND TARGETS **********************************
Sub TopLeftIslandSlingshot_Hit            '19 OK
  If MacTilt=0 Then vpmTimer.PulseSw 19
End Sub
Sub TopRightIslandSlingshot_Hit
  If MacTilt=0 Then vpmTimer.PulseSw 19
End Sub
Sub UpperLeftIslandSlingshot_Hit
  If MacTilt=0 Then vpmTimer.PulseSw 19
End Sub
Sub LowerLeftIslandSlingshot_Hit
  If MacTilt=0 Then vpmTimer.PulseSw 19
End Sub
Sub UpperRightIslandSlingshot_Hit
  If MacTilt=0 Then vpmTimer.PulseSw 19
End Sub
Sub LowerRightIslandSlingshot_Hit
  If MacTilt=0 Then vpmTimer.PulseSw 19
End Sub
Sub CenterIslandSlingshot_Hit
  If MacTilt=0 Then vpmTimer.PulseSw 19
End Sub
' ***********************************************************************

' **************************** BUMPERS **********************************
Sub BumperLeft_Hit
Playsound SoundFX("fx_bumpers1", DOFContactors), 0, 0.033, Pan(BumperLeft), AudioFade(BumperLeft)
vpmTimer.PulseSw 23
End Sub
Sub BumperRight_Hit
Playsound SoundFX("fx_bumpers1", DOFContactors), 0, 0.033, Pan(BumperRight), AudioFade(BumperRight)
vpmTimer.PulseSw 24
End Sub
Sub BumperCenter_Hit
Playsound SoundFX("fx_bumpers1", DOFContactors), 0, 0.033, Pan(BumperCenter), AudioFade(BumperCenter)
vpmTimer.PulseSw 25
End Sub
' ***********************************************************************

' *********************** KICKERS TOP LANES *****************************
Sub KickerA_Hit
Controller.Switch(28)=1
PlaySoundAt SoundFX("fx_kicker_enter", DOFContactors), KickerA
End Sub           '28 OK

Sub KickerT_Hit
Controller.Switch(29)=1
PlaySoundAt SoundFX("fx_kicker_enter", DOFContactors), KickerT
End Sub           '29 OK

Sub KickerE_Hit
Controller.Switch(30)=1
PlaySoundAt SoundFX("fx_kicker_enter", DOFContactors), KickerE
End Sub           '30 OK

Sub KickerR_Hit
Controller.Switch(31)=1
PlaySoundAt SoundFX("fx_kicker_enter", DOFContactors), KickerR
End Sub           '31 OK

Sub KickerI_Hit
Controller.Switch(32)=1
PlaySoundAt SoundFX("fx_kicker_enter", DOFContactors), KickerI
End Sub           '32 OK

Sub KickerBonus_Hit
Controller.Switch(33)=1
PlaySoundAt SoundFX("fx_kicker_enter", DOFContactors), KickerBonus
End Sub     '33 OK

Sub KickerEB_Hit
Controller.Switch(34)=1
PlaySoundAt SoundFX("fx_kicker_enter", DOFContactors), KickerEB
End Sub     '34 OK
' ***********************************************************************

Sub Drain_Hit                       '35 OK
  bsTrough.AddBall Me
PlaySoundAt SoundFX("fx_Drain", DOFContactors), Drain
End Sub

Sub BallRelease_Hit:Me.Kick 80,5:End Sub

' ****** mechanicals timers ******
dim gameon
Sub GateTimer_Timer()
   Gate2Flap.RotZ = ABS(Gate2.currentangle)
   Gate3Flap.RotZ = ABS(Gate3.currentangle)
   rgateprim.RotY= RightGate.currentangle + 90
   lgateprim.RotY= leftGate.currentangle + 90
   RollingUpdate

gameon = LightGamer1.state
If gameon=0 Then

        If LeftFlipper.CurrentAngle < LeftFlipper.StartAngle Then
          PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
          LeftFlipper.RotateToStart
          LeftFlipperOn = 0
          PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipperTop
          LeftFlipperTop.RotateToStart
          LeftFlipperTopOn = 0
        End If

        If RightFlipper.CurrentAngle > RightFlipper.StartAngle Then
          PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
          RightFlipper.RotateToStart
          RightFlipperOn = 0
          PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipperTop
          RightFlipperTop.RotateToStart
          RightFlipperTopOn = 0
        End If
End If


End Sub

'tilt Function
Dim MacTilt,LeftF,RightF
MacTilt=0:LeftF=0:RightF=0

Sub Tilted_Timer
  MacTilt=Tilt.State
    If MacTilt=1 Then

        If LeftFlipper.CurrentAngle < LeftFlipper.StartAngle Then
          PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
          LeftFlipper.RotateToStart
          LeftFlipperOn = 0
          PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipperTop
          LeftFlipperTop.RotateToStart
          LeftFlipperTopOn = 0
        End If

        If RightFlipper.CurrentAngle > RightFlipper.StartAngle Then
          PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
          RightFlipper.RotateToStart
          RightFlipperOn = 0
          PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipperTop
          RightFlipperTop.RotateToStart
          RightFlipperTopOn = 0
        End If

      BumperLeft.Threshold=100
      BumperCenter.Threshold=100
      BumperRight.Threshold=100
      LeftSlingshotBounce.SlingshotThreshold=100
      RightSlingshotBounce.SlingshotThreshold=100
    Else
      BumperLeft.Threshold=1.5
      BumperCenter.Threshold=1.5
      BumperRight.Threshold=1.5
      LeftSlingshotBounce.SlingshotThreshold=1
      RightSlingshotBounce.SlingshotThreshold=1
    End If
End Sub

' ************************************************************************


' *************************************************************************
' ***************** LAMP MAPPING, will reassign to VPM later **************
' *************************************************************************
Set Lights(1)=LightTrigger4
Set Lights(2)=LightTrigger1
Set Lights(3)=LightTarget2
Set Lights(4)=LightTarget3
Set Lights(5)=LightTarget8
Set Lights(6)=LightTrigger7
Set Lights(7)=LightTrigger6
Set Lights(8)=LightTrigger5
Set Lights(9)=LightSpecialLeftLow'Light9
Set Lights(10)=LightSpecialLeftMiddle'Light10
Set Lights(11)=Light11
Set Lights(12)=Light12
Set Lights(13)=LightSpecialRightLow'Light13
Set Lights(14)=LightSpecialRightHigh'Light14
Set Lights(15)=LightSpecialCenter'Light15
Set Lights(16)=Light16
Set Lights(17)=LightBonus04 'J1-P1/B
Set Lights(18)=LightBonus03 'J1-P1/M
Set Lights(19)=LightBonus02 'J1-P1/X
Set Lights(20)=LightBonus01 'J2-P2/1
Set Lights(21)=LightBonus07 'J1-Pl/F
Set Lights(22)=LightBonus06 'J1-P1/8
Set Lights(23)=LightBonus05 'J1-P1/17
Set Lights(24)=Light24
Set Lights(25)=LightBonus10 'J1-P1/3
Set Lights(26)=LightBonus09 'J1-P1/12
Set Lights(27)=LightBonus08 'J1-P1/21
Set Lights(28)=Light28
Set Lights(29)=LightGamer1 'J1-P1/7
Set Lights(30)=LightGamer2 'J1-P1/T
Set Lights(31)=LightGamer3 'J2-P2/D
Set Lights(32)=Light32
Set Lights(33)=LightExtra04
Set Lights(34)=LightExtra03
Set Lights(35)=LightExtra02
Set Lights(36)=LightExtra01
Set Lights(37)=LightExtra07'Light37
Set Lights(38)=LightExtra06
Set Lights(39)=LightExtra05
Set Lights(40)=Light40
Set Lights(41)=LightExtraBall
Set Lights(42)=LightExtra09
Set Lights(43)=LightExtra08
Set Lights(44)=Light44
Set Lights(45)=Light45
Set Lights(46)=Light46    'MATCH
Set Lights(47)=LightGamer4 'J2-P2/E
Set Lights(48)=Light48
Set Lights(49)=Light49
Set Lights(50)=LightTarget9
Set Lights(51)=LightT 'J1-P1/15
Set Lights(52)=LightA 'J2-P2/3
Set Lights(53)=BumperLeftL 'J1-P1/10
Set Lights(54)=BumperRightL 'J1-P1/1
Set Lights(55)=BumperCenterL
Set Lights(56)=LightBonusX2 'J2-P2/5
Set Lights(57)=LightI 'J1-P1/5
Set Lights(58)=LightR 'J1-P1/14
Set Lights(59)=LightE 'J1-P2/B
Set Lights(60)=Light60
Set Lights(61)=Light61 'J1-P1/K EXTRA BALL
Set Lights(62)=Tilt 'J1-P1/V TILT
Set Lights(63)=GameOver 'GAME OVER
Set Lights(64)=Light64
Set Lights(65)=Light65
Set Lights(66)=Light66
Set Lights(67)=Light67
Set Lights(68)=Light68
Set Lights(69)=Light69
Set Lights(70)=Light70
Set Lights(71)=Light71
Set Lights(72)=Light72
Set Lights(73)=Light73
Set Lights(74)=Light74
Set Lights(75)=Light75
Set Lights(76)=Light76
Set Lights(77)=Light77
Set Lights(78)=Light78
Set Lights(79)=Light79
Set Lights(80)=Light80
Set Lights(81)=Light81
Set Lights(82)=Light82
Set Lights(83)=Light83
Set Lights(84)=Light84
Set Lights(85)=Light85
Set Lights(86)=Light86
Set Lights(87)=Light87
Set Lights(88)=Light88
Set Lights(89)=Light89
Set Lights(90)=Light90
Set Lights(91)=Light91
Set Lights(92)=Light92
Set Lights(93)=Light93
Set Lights(94)=Light94
Set Lights(95)=Light95
Set Lights(96)=Light96
Set Lights(97)=Light97
Set Lights(98)=Light98
Set Lights(99)=Light99
Set Lights(100)=Light90
Set Lights(101)=Light101
Set Lights(102)=Light102
Set Lights(103)=Light103
Set Lights(104)=Light104
Set Lights(105)=Light105
Set Lights(106)=Light106
Set Lights(107)=Light107
Set Lights(108)=Light108
Set Lights(109)=Light109
Set Lights(110)=Light110
Set Lights(111)=Light111
Set Lights(112)=Light112
Set Lights(113)=Light113
Set Lights(114)=Light114
Set Lights(115)=Light115
Set Lights(116)=Light116
Set Lights(117)=Light117
Set Lights(118)=LightSpecialRightLow'Light118
Set Lights(119)=Light119
Set Lights(120)=Light120
Set Lights(121)=Light121
Set Lights(122)=Light122
Set Lights(124)=Light124

Set Lights(125)=Light125 'tilt
Set Lights(126)=Light126
Set Lights(127)=Light127
Set Lights(128)=Light128

Set Lights(129)=LightS1
Set Lights(130)=LightS2
Set Lights(131)=LightS3
Set Lights(132)=LightS4
' *************************************************************************
' ************************* LAMP MAPPING END ******************************
' *************************************************************************


' ******************************************************* ATARIANS TABLE SETUP ******************************************************************
' input
Dim rfon, lfon
Sub Atarians_KeyDown(ByVal KeyCode)

  If vpmKeyDown(KeyCode) Then Exit Sub
  If KeyCode=PlungerKey Then PlaySoundAt"FX_Plungerpull", Plunger:Plunger.PullBack
  If keycode=LeftFlipperKey Then  lfon=1: LLFELIPPER: end if
  If keycode=RightFlipperKey Then rfon=1: LRFLIPPER: end if

End Sub


Sub Atarians_KeyUp(ByVal KeyCode)

  If vpmKeyUp(KeyCode) Then Exit Sub
  If KeyCode=PlungerKey Then PlaySoundAt"FX_Plunger", Plunger:Plunger.Fire
  If keycode=LeftFlipperKey Then lfon=0: LLFELIPPER: end if
  If keycode=RightFlipperKey Then rfon=0: LRFLIPPER: end if

End Sub

' environment
Dim bsTrough

' ******************************** TABLE INIT ***********************************
Sub Atarians_Init
    On Error Resume Next
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = cCredits
    .HandleMechanics = 0
    .ShowDMDOnly = 1
    .ShowFrame = False
    .ShowTitle = False
    If Err Then MsgBox Err.Description
    .Run
  End With
  On Error Goto 0
  Controller.Hidden=0

' Thalamus : Was missing 'vpminit me'
vpminit me

'  PinMAMETimer.Interval=PinMAMEInterval
'    PinMAMETimer.Enabled=True
    ' Nudging
  vpmNudge.TiltSwitch = 4
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj=Array(BumperLeft,BumperCenter,BumperRight,TopLeftIslandSlingshot,TopRightIslandSlingshot,UpperRightIslandSlingshot,LowerRightIslandSlingshot,CenterIslandSlingshot,UpperLeftIslandSlingshot,LowerLeftIslandSlingshot,LeftSlingshotBounce,RightSlingshotBounce)

  Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,35,0,0,0,0,0,0
    bsTrough.InitKick BallRelease,80,5
    bsTrough.InitExitSnd SoundFX("Ballrel",DOFContactors),SoundFX("solon",DOFContactors)
    bsTrough.Balls=1

End Sub

Sub Atarians_Paused:Controller.Pause=True:End Sub
Sub Atarians_UnPaused:Controller.Pause=False:End Sub

 'Atari The Atarians
'added by Inkochnito
Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,400,"The Atarians - DIP switches"
    .AddFrame 0,5,190,"Coins per credit",&H000000C3,Array("1 coin 1 credit",0,"1 coin 2 credits",&H00000002,"1 coin 3 credits",&H00000001,"1 coin 4 credits",&H00000003,"2 coins 1 credits",&H00000080)'SW2-3&SW2-4&SW2-5&SW2-6 (dip 2&1&8&7)
    .AddFrame 0,95,190,"Maximum credits",49152,Array("8 credits",0,"12 credits",32768,"15 credits",&H00004000,"20 credits",49152)'SW1-5&SW1-6 (dip 16&15)
    .AddFrame 0,172,190,"Game option",&H00000100,Array("off",0,"on",&H00000100)'SW1-4 (dip 9)
    .AddFrame 0,220,190,"Balls per game",&H00000008,Array("5 balls",0,"3 balls",&H00000008)'SW2-1 (dip 4)
    .AddFrame 0,268,190,"Game option",&H00001000,Array("off",0,"on",&H00001000)'SW1-8 (dip 13)
    .AddFrame 210,5,190,"Score threshold level",&H000F0000,Array("20K-30K-40K",0,"60K-90K-120K",&H00040000,"90K-130K-170K",&H00070000,"130K-190K-250K",&H000B0000,"170K-250K-330K",&H000F0000)'Rotary (dip 17&18&19&20)
    .AddFrame 210,95,190,"Special award",&H00000030,Array("200,000 points",0,"10,000 points",&H00000010,"extra ball",&H00000020,"replay",&H00000030)'SW2-7&SW2-8 (dip 6&5)
    .AddFrame 210,172,190,"Game option",&H00000400,Array("off",0,"on",&H00000400)'SW1-2 (dip 11)
    .AddFrame 210,220,190,"Game option",&H00000200,Array("off",0,"on",&H00000200)'SW1-3 (dip 10)
    .AddFrame 210,268,190,"Game option",&H00002000,Array("off",0,"on",&H00002000)'SW1-7 (dip 14)
    .AddChk 0,320,190,Array("Match feature",&H00000004)'SW2-2 (dip 3)
    .AddChk 210,320,190,Array("Diagnostic mode (must be off)",&H00000800)'SW1-1 (dip 12)
    .AddLabel 50,350,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")

' *********************************************************************************************************************************
' *********************************************************************************************************************************
' *********************************************************************************************************************************

' *********************************************************************************************************************************
' ******************************************* SEGMENTED LED DISPLAYS***************************************************************
' *********************************************************************************************************************************

 Dim Digits(27)
Digits(0)=Array(LX1,LX2,LX3,LX4,LX5,LX6,LX7)
Digits(1)=Array(LX8,LX9,LX10,LX11,LX12,LX13,LX14)
Digits(2)=Array(LX15,LX16,LX17,LX18,LX19,LX20,LX21)
Digits(3)=Array(LX22,LX23,LX24,LX25,LX26,LX27,LX28)
Digits(4)=Array(LX29,LX30,LX31,LX32,LX33,LX34,LX35)
Digits(5)=Array(LX36,LX37,LX38,LX39,LX40,LX41,LX42)
Digits(6)=Array(LX43,LX44,LX45,LX46,LX47,LX48,LX49)
Digits(7)=Array(LX50,LX51,LX52,LX53,LX54,LX55,LX56)
Digits(8)=Array(LX57,LX58,LX59,LX60,LX61,LX62,LX63)
Digits(9)=Array(LX64,LX65,LX66,LX67,LX68,LX69,LX70)
Digits(10)=Array(LX71,LX72,LX73,LX74,LX75,LX76,LX77)
Digits(11)=Array(LX78,LX79,LX80,LX81,LX82,LX83,LX84)
Digits(12)=Array(LX85,LX86,LX87,LX88,LX89,LX90,LX91)
Digits(13)=Array(LX92,LX93,LX94,LX95,LX96,LX97,LX98)
Digits(14)=Array(LX99,LX100,LX101,LX102,LX103,LX104,LX105)
Digits(15)=Array(LX106,LX107,LX108,LX109,LX110,LX111,LX112)
Digits(16)=Array(LX113,LX114,LX115,LX116,LX117,LX118,LX119)
Digits(17)=Array(LX120,LX121,LX122,LX123,LX124,LX125,LX126)
Digits(18)=Array(LX127,LX128,LX129,LX130,LX131,LX132,LX133)
Digits(19)=Array(LX134,LX135,LX136,LX137,LX138,LX139,LX140)
Digits(20)=Array(LX141,LX142,LX143,LX144,LX145,LX146,LX147)
Digits(21)=Array(LX148,LX149,LX150,LX151,LX152,LX153,LX154)
Digits(22)=Array(LX155,LX156,LX157,LX158,LX159,LX160,LX161)
Digits(23)=Array(LX162,LX163,LX164,LX165,LX166,LX167,LX168)
Digits(24)=Array(LX169,LX170,LX171,LX172,LX173,LX174,LX175)
Digits(25)=Array(LX176,LX177,LX178,LX179,LX180,LX181,LX182)
Digits(26)=Array(LX183,LX184,LX185,LX186,LX187,LX188,LX189) 'CREDITS X10
Digits(27)=Array(LX190,LX191,LX192,LX193,LX194,LX195,LX196) 'CREDITS X1

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii=0 To UBound(chgLED)
      num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
      For Each obj In Digits(num)
        If chg And 1 Then obj.State=stat And 1
        chg=chg\2:stat=stat\2
      Next
    Next
  End If
End Sub

' *********************************************************************************************************************************
' *********************************************************************************************************************************
' *********************************************************************************************************************************



' *********************************************************************************************************************************
' *********************************************************************************************************************************
' ***************************                  JP SALAS PRESENTS                 **************************************************
' ***************************                    TABLE PHYSICS                   **************************************************
' ***************************                  AND SSF SOUND FX                  **************************************************
' *********************************************************************************************************************************
' *********************************************************************************************************************************
' *********************************************************************************************************************************
' *********************************************************************************************************************************




' *************************** JP's FLIPPER PHYSICS *********************************
'*********************************************************
' Real Time Flipper adjustments - by JLouLouLou & JPSalas
'        (to enable flipper tricks)
'*********************************************************

Dim FlipperPower
Dim FlipperElasticity
Dim SOSTorque, SOSAngle
Dim FullStrokeEOS_Torque, LiveStrokeEOS_Torque
Dim LeftFlipperOn
Dim RightFlipperOn
Dim LeftFlipperTopOn
Dim RightFlipperTopOn

Dim LLiveCatchTimer
Dim RLiveCatchTimer
Dim LTLiveCatchTimer
Dim RTLiveCatchTimer
Dim LiveCatchSensivity

FlipperPower = 5000
FlipperElasticity = 0.85
FullStrokeEOS_Torque = 0.3  ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2  ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10
LeftFlipperTop.EOSTorqueAngle = 10
RightFlipperTop.EOSTorqueAngle = 10

SOSTorque = 0.1
SOSAngle = 6

LiveCatchSensivity = 10

LLiveCatchTimer = 0
RLiveCatchTimer = 0
LTLiveCatchTimer = 0
RTLiveCatchTimer = 0

LeftFlipper.TimerInterval = 1
LeftFlipper.TimerEnabled = 1

Sub LeftFlipper_Timer 'flipper's tricks timer
'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If LeftFlipperOn = 1 Then
    If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle then
      LeftFlipper.EOSTorque = FullStrokeEOS_Torque
      LLiveCatchTimer = LLiveCatchTimer + 1
      If LLiveCatchTimer < LiveCatchSensivity Then
        LeftFlipper.Elasticity = 0
      Else
        LeftFlipper.Elasticity = FlipperElasticity
        LLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    LeftFlipper.Elasticity = FlipperElasticity
    LeftFlipper.EOSTorque = LiveStrokeEOS_Torque
    LLiveCatchTimer = 0
  End If


'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
   If LeftFlipperTop.CurrentAngle >= LeftFlipperTop.StartAngle - SOSAngle Then LeftFlipperTop.Strength = FlipperPower * SOSTorque else LeftFlipperTop.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If LeftFlipperTopOn = 1 Then
    If LeftFlipperTop.CurrentAngle = LeftFlipperTop.EndAngle then
      LeftFlipperTop.EOSTorque = FullStrokeEOS_Torque
      LTLiveCatchTimer = LTLiveCatchTimer + 1
      If LTLiveCatchTimer < LiveCatchSensivity Then
        LeftFlipperTop.Elasticity = 0
      Else
        LeftFlipperTop.Elasticity = FlipperElasticity
        LTLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    LeftFlipperTop.Elasticity = FlipperElasticity
    LeftFlipperTop.EOSTorque = LiveStrokeEOS_Torque
    LTLiveCatchTimer = 0
  End If


'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If RightFlipperOn = 1 Then
    If RightFlipper.CurrentAngle = RightFlipper.EndAngle Then
      RightFlipper.EOSTorque = FullStrokeEOS_Torque
      RLiveCatchTimer = RLiveCatchTimer + 1
      If RLiveCatchTimer < LiveCatchSensivity Then
        RightFlipper.Elasticity = 0
      Else
        RightFlipper.Elasticity = FlipperElasticity
        RLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    RightFlipper.Elasticity = FlipperElasticity
    RightFlipper.EOSTorque = LiveStrokeEOS_Torque
    RLiveCatchTimer = 0
  End If


'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If RightFlipperTop.CurrentAngle <= RightFlipperTop.StartAngle + SOSAngle Then RightFlipperTop.Strength = FlipperPower * SOSTorque else RightFlipperTop.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If RightFlipperTopOn = 1 Then
    If RightFlipperTop.CurrentAngle = RightFlipperTop.EndAngle Then
      RightFlipperTop.EOSTorque = FullStrokeEOS_Torque
      RTLiveCatchTimer = RTLiveCatchTimer + 1
      If RTLiveCatchTimer < LiveCatchSensivity Then
        RightFlipperTop.Elasticity = 0
      Else
        RightFlipperTop.Elasticity = FlipperElasticity
        RTLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    RightFlipperTop.Elasticity = FlipperElasticity
    RightFlipperTop.EOSTorque = LiveStrokeEOS_Torque
    RTLiveCatchTimer = 0
  End If
End Sub

' *******************************************************************************
' ****************** END FLIPPER PHYSICS ****************************************
' *******************************************************************************

'**********************
' Mechanical Sounds
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

sub metals_hit(idx)
PlaySoundAtBall "fx_metalhit"
End Sub

Sub wires_hit(idx)
PlaySoundAtBall "fx_MetalWire"
End Sub

Sub bands_hit(idx)
PlaySoundAtBall "fx_rubber_band"
End Sub


Sub longbands_hit(idx)
PlaySoundAtBall "fx_rubber_longband"
End Sub

Sub pegs_hit(idx)
PlaySoundAtBall "fx_rubber_peg"
End Sub

Sub pins_hit(idx)
PlaySoundAtBall "fx_rubber_pin"
End Sub

Sub woods_hit(idx)
PlaySoundAtBall "fx_woodhit"
End Sub

Sub Gate1_hit
PlaySoundAt "FX_WireGate", Gate1
End Sub

Sub Gate2rebound_hit
PlaySoundAt "FX_Plate_Gate", Gate2
End Sub

Sub Gate3rebound_hit
PlaySoundAt "FX_Plate_Gate", Gate3
End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Atarians.width
TableHeight = Atarians.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Atarians" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function



'***********************************
'   JP's VP10 Rolling Sounds v4.0
'   JP's Ball Shadows
'   JP's Ball Speed Control
'   Rothbauer's dropping sounds
'***********************************

Const tnob = 2   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 35 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate() 'call this routine from any realtime timer you may have, running at an interval of 10 is good.

    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls and hide the shadow
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)

    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' draw the ball shadow
    For b = lob to UBound(BOT)


    'play the rolling sound for each ball
        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))*0.5
            Else
                ballpitch = Pitch(BOT(b)) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 5
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed control
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub

Sub Atarians_Exit
  Controller.Pause = False
  Controller.Stop
End Sub

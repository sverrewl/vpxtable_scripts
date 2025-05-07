'very special thanks to
'El_Condor,Kingzap,TheGuru,Destruk,Aussie33,Nedo,Simeon,wpcmame for all there help
'information,pictures and efforts on this table.
' Thalamus - seems ok

Option Explicit
On Error Resume Next

ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const UseVPMColoredDMD = true
Const UseVPMDMD = 0

LoadVPM "01120200", "GTS3.VBS", 3.10

Sub InitVPM()
  BigHurtOptions = CInt("0" & LoadValue("BigHurt","Options")) : BigHurtSetOptions
  On Error Resume Next
End Sub

if table1.showDT = false then rampL.visible = 0:rampR.visible = 0

Const UseSolenoids  = 2
Const UseLamps      = True
Const UseSync   = True
Const UseGI     = False        'true 'Opto
Const bladeArt  = 1 '1=On (Art), 2=Sideblades Off.


'Set LampCallback    = GetRef("UpdateMultipleLamps")

Const SSolenoidOn   = "SolOn"
Const SSolenoidOff  = "SolOn"
Const SFlipperOn    = "left_flipper_up"
Const SFlipperOff   = "left_flipper_down"
Const SCoin     = "Quarter"

Const swLCoin = 0
Const swRCoin = 1
Const swCCoin = 2
Const swCoinShuttle = 3
Const swStartButton = 4
Const swTournament = 5
Const swFrontDoor = 6
Const swDrop1 = 7

Const swLBumper = 10
Const swRBumper = 11
Const swLRubber = 12
Const swRRubber = 13
Const swTrough = 14
'Const sw = 15
'Const sw = 16
Const swDrop2 = 17

Const swShooterLane = 20
Const swRightUpKicker = 21
Const swHoleKicker = 22
Const swGlove= 23
Const swOuthole = 24
'Const sw = 25
'Const sw = 26
Const swDrop3 = 27

'40 to 77 not used
Const swLeftTop = 80
Const swRightTop = 81
Const swLFlip = 82
Const swRFlip=83
Const swLeftOutLane = 85
Const swRigtOutLane = 86
Const swLeftScoreBoard = 87

Const swLeftRamp = 90
Const swRightRamp = 91
Const swLeftInlane = 95
Const swRightInlane = 96
Const swRightScoreBoard = 97

Const swHomeRun = 100
Const swScoop   = 101
Const swLTarget = 105
Const swRightTopRollover = 106
Const swCaptiveBallRollover = 107

Const swLeftUpKicker = 110
Const swLeftBottomTarget = 115
Const swRightSideRollOver = 116
Const swGloveAllTheWay = 117

Dim VR_Room
If RenderingMode = 2 Then VR_Room=True Else VR_Room=False      'VRRoom set based on RenderingMode in version 10.72

Set Lights(0) = Light0
'Set Lights(1) = Light1
Set Lights(2) = Light2
Set Lights(3) = Light3
Set Lights(4) = Light4
Set Lights(5) = Light5
Set Lights(6) = Light6
Set Lights(7) = Light7

Set Lights(10) = Light10
Set Lights(11) = Nothing
Set Lights(12) = Nothing
Set Lights(13) = Nothing
Set Lights(14) = Nothing
Set Lights(15) = Nothing
Set Lights(16) = Nothing
Set Lights(17) = Nothing

Set Lights(20) = Light20
Set Lights(21) = Light21
Set Lights(22) = Light22
Set Lights(23) = Light23
Set Lights(24) = Light24
Set Lights(25) = Light25
Set Lights(26) = Light26
Set Lights(27) = Light27

Set Lights(30) = Light30
Set Lights(31) = Light31
Set Lights(32) = Light32
Set Lights(33) = Light33
Set Lights(34) = Light34
Set Lights(35) = Light35
Set Lights(36) = Light36
Set Lights(37) = Light37

Set Lights(40) = Light40
Set Lights(41) = Light41
Set Lights(42) = Light42
Set Lights(43) = Light43
Set Lights(44) = Light44
Set Lights(45) = Light45
Set Lights(46) = Light46
Set Lights(47) = Light47

Set Lights(50) = Light50
Set Lights(51) = Light51
Set Lights(52) = Light52
Set Lights(53) = Light53
Set Lights(54) = Light54
Lights(55) = array(Light55,Light55a)
Set Lights(56) = Light56
Lights(57) = array(Light57,Light57a)

Set Lights(60) = Light60
Set Lights(61) = Light61
Set Lights(62) = Nothing
Set Lights(63) = Nothing
Set Lights(64) = Nothing
Set Lights(65) = Light65
Set Lights(66) = Light66
Set Lights(67) = Light67

Set Lights(70) = Light70
Set Lights(71) = Light71
Set Lights(72) = Nothing
Set Lights(73) = Nothing
Set Lights(74) = Nothing
Set Lights(75) = Light75
Set Lights(76) = Light76
Set Lights(77) = Light77

Set Lights(80) = Light80
Set Lights(81) = Light81
Set Lights(82) = Light82
Set Lights(83) = Light83
Set Lights(84) = Light84
Set Lights(85) = Light85
Set Lights(86) = Light86
Set Lights(87) = Light87

Set Lights(90) = Nothing
Set Lights(91) = Light91
Set Lights(92) = Light92
Set Lights(93) = Light93
'Set Lights(94) = Light94     removed 94/ 97
'Set Lights(95) = Light95
'Set Lights(96) = Light96
'Set Lights(97) = Light97


'AuxBoard   140/147 sofar all aux board lights workin disabled them because they did nothing special
'Set Lights(146) = Bumper14
'Set Lights(147) = Bumper15




'Flippers/Slings/Jets/Trough/Knocker/Plunger
SolCallback(sLLFlipper)   = "SolFlipper LeftFlipper,Flipper3,"
SolCallback(sLRFlipper)   = "SolFlipper RightFlipper,Nothing,"

SolCallback(1)  = "vpmSolSound ""Jet1"","
SolCallback(2)  = "vpmSolSound ""Jet2"","
SolCallback(3)  = "vpmSolSound ""lSling"","
SolCallback(4)  = "vpmSolSound ""lSling"","
SolCallback(5)  = "SolShooter"
SolCallback(6)  = "SolVUK"'"bsRLowKicker.SolOut"
SolCallback(7)  = "bsLKicker.SolOut"
'SolCallback(8)  = "DropReset"
'SolCallback(9)  = "dtDrop1.SolDropUp"
'SolCallback(10) = "dtDrop2.SolDropUp"
'SolCallback(11) = "dtDrop3.SolDropUp"
SolCallback(12) = "CenterRampLift"
SolCallback(13) = "bsUpKicker.SolOut"
SolCallback(14) = "SolShooterlane"
SolCallback(16) = "Gate1.open = "             'TopRightPlungerGate
SolCallback(17) = "Gate2.open = "             'TopBallGate
SolCallback(18) = "vpmFlasher array(Bumper3,Bumper3a,Bumper3b),"         'LeftDome
SolCallback(19) = "vpmFlasher array(Bumper4,Bumper4a,Bumper4b),"         'RightDome
SolCallback(20) = "vpmFlasher Bumper7,"         'HomeRun
SolCallback(21) = "vpmFlasher array(Bumper6,Bumper6a),"         'GrandSlam
SolCallback(22) = "vpmFlasher Bumper5,"         'PopBumpersArea
SolCallback(26) = "PFGI"   'SolLightBoxRelay
SolCallback(27) = "vpmSolSound ""Knocker"","
SolCallback(28) = "bsTrough.SolOut"
SolCallback(29) = "bsTrough.SolIn"
SolCallback(30) = "vpmSolSound ""Knocker"","
SolCallback(31) = "SolNugdeRelay"
SolCallback(32) = "GameOverRelay"

'Playfield GI
Sub PFGI(Enabled)
 If Enabled Then
   dim xx
    For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
Table1.ColorGradeImage = "ColorGrade_4"
 Else
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
Table1.ColorGradeImage = "ColorGrade_8"
 End If
End Sub

Sub SolNugdeRelay(enabled)
  if enabled then
    playsound "SolOn"
'   TBumper.disabled = 1
'   LBumper.disabled = 1
    LSling.disabled = 1
    RSling.disabled = 1
    'lighttilt.state = 0
  else
'   TBumper.disabled = 0
'   LBumper.disabled = 0
    LSling.disabled = 0
    RSling.disabled = 0
    'lighttilt.state = 1
  end if
end sub


Sub GameOverRelay(enabled)
  if enabled then
    playsound "SolOn"
'   TBumper.disabled = 0
'   LBumper.disabled = 0
    LSling.disabled = 0
    RSling.disabled = 0
    'lighttilt.state = 1
  else
'   TBumper.disabled = 1
'   LBumper.disabled = 1
    LSling.disabled = 1
    RSling.disabled = 1
    'lighttilt.state = 0
  end if
end sub

'init table
Dim bsTrough,bsUpKicker,bsLKicker
Dim dtDrop1,dtDrop2,dtDrop3,dtDropR
Dim mGlove
Dim cGameName
cGameName = "cc_12"

Sub Table1_init()
  InitVPM
  With Controller
   .GameName = "bighurt"          '"cc_12"    'ArcadeRom
   If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine   = "Big Hurt Gottlieb 1995" & vbNewLine & "by bigus1"
    .HandleKeyboard   = False
   .ShowTitle      = False
   .ShowDMDOnly    = True
    .HandleMechanics  = False
   .ShowFrame      = False
Controller.Games(cGameName).Settings.Value("volume") = -12.0
    .DIP(0) = &H00
    On Error Resume Next
    .Run
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
 ' Main Timer init
 PinMAMETimer.Interval = PinMAMEInterval
 PinMAMETimer.Enabled = True
'  vpmBallImage = "ball"
  table1.yieldtime = 1
  Gate5.open = 0
  vpmCreateEvents Gloves

  ' Nudging
  vpmNudge.TiltSwitch = 151
  vpmNudge.Sensitivity = 6

  'Place 4 th ball in the outhole
  Drain.CreateBall
  Drain.Kick 0,3

  'Captive Ball
  With Captiveresting.CreateBall:end with
    Captiveresting.kick 180, 0.1
  Captiveresting.enabled = 0

  ' BallStack
  Set bsTrough = New cvpmBallStack
  bsTrough.InitSw 24,0,0,14,0,0,0,0
  bsTrough.InitKick BallRelease, 82, 4
  bsTrough.InitEntrySnd "SolOn", "SolOn"
  bsTrough.InitExitSnd "BallRel", "SolOn"
    bsTrough.Balls = 4
' bsTrough.BallImage="ball"

    'set bsRLowKicker = new cvpmBallStack
    'bsRLowKicker.InitSw 0,21,0,0,0,0,0,0
  'bsRLowKicker.InitSaucer Kicker2,21,220,2
  'bsRLowKicker.InitEntrySnd "Solenoid", "SolOn"
  'bsRLowKicker.InitExitSnd "Popper", "SolOn"
  'bsRLowKicker.Balls = 0
  'bsRLowKicker.BallImage = "ball"

  set bsLKicker = new cvpmBallStack
  bsLKicker.InitSw 0,110,0,0,0,0,0,0
  bsLKicker.InitKick Kicker6,235,5
  bsLKicker.InitEntrySnd "Popper", "SolOn"
  bsLKicker.InitExitSnd "BallRel", "SolOn"
  bsLKicker.Balls = 0
'    bsLKicker.BallImage = "ball"

    set bsUpKicker = new cvpmBallStack
    bsUpKicker.InitSw 0,22,0,0,0,0,0,0
  bsUpKicker.InitKick Kicker11,180,3
  bsUpKicker.InitEntrySnd "Popper", "SolOn"
  bsUpKicker.InitExitSnd "Popper", "SolOn"
  bsUpKicker.Balls = 0
'    bsUpKicker.BallImage = "ball"

    'glove walls
  Glove1.isdropped=false
  Glove2.isdropped=true
  Glove3.isdropped=true
  Glove4.isdropped=true
  Glove5.isdropped=true
  Glove6.isdropped=true
  Glove7.isdropped=true
  Glove8.isdropped=true
  Glove9.isdropped=true
  Glove10.isdropped=true

    Wall178.isdropped=False
    Wall179.isdropped=true
    Wall180.isdropped=true
    Wall181.isdropped=true
    Wall182.isdropped=true
    Wall183.isdropped=true
    Wall184.isdropped=true
    Wall185.isdropped=true
    Wall186.isdropped=true
    Wall187.isdropped=true

    'DropTargets
  Set dtDrop1 = new cvpmDropTarget
    dtDrop1.InitDrop Wall17,7
    dtDrop1.InitSnd "DropTarget","TargetReset"
    SolCallback(9)="SolDropT1"

    Set dtDrop2 = New cvpmDropTarget
    dtDrop2.InitDrop Wall18,17
    dtDrop2.InitSnd "DropTarget","TargetReset"
    SolCallback(10)="SolDropT2"

    Set dtDrop3 = new cvpmDropTarget
    dtDrop3.InitDrop Wall19,27
    dtDrop3.InitSnd "DropTarget","TargetReset"
    SolCallback(11)="SolDropT3"

    'DropTargets Reset
    Set dtDropR = new cvpmDropTarget
    dtDropR.InitDrop Array(Wall17,Wall18,Wall19),Array(7,17,27)
    dtDropR.InitSnd "DropTarget","TargetReset"
    SolCallback(8)="SolDropTarget"

    'GloveMech
    Set mGlove = New cvpmMech
  With mGlove
    .Sol1 = 25
    .Length = 70
    .Steps = 19
    .MType =vpmMechOneSol + vpmMechReverse + vpmMechLinear
        .Acc = 20
        .Ret =1
        .CallBack = GetRef ("Glovemotion")
    .Start
  End With

  dim VRobject
  if VR_Room Then
   frontlockbar.visible = False
    RampL.visible = False
   RampR.visible = False
   Ramp007.Image = "Black"
   Ramp012.Image = "Black"
   Ramp008.visible = False
 Else
    for each vrobject in colVR_Stuff : vrobject.visible = false : next
    Ramp012.visible = False
 end If

End Sub

Dim BigHurtOptions
Private Sub BigHurtShowDips
  If Not IsObject(vpmDips) Then ' First time
    Set vpmDips = New cvpmDips
    With vpmDips
      .AddForm 220, 240, "Game options"
      .AddFrame 0, 0, 80, "Country", &Hf0,_
        Array("USA", &H00, "USA", &Hf0, "European", &Hd0,_
              "Export", &Ha0, "Export Alt", &H80, "France", &Hb0,_
              "France 1", &H10, "France 2", &H30, "France 3", &H90,_
              "Germany", &H20, "Spain",     &He0, "UK", &Hc0)
      .AddFrameExtra 90, 0, 140, "Not used", 0, array("Not used1",0,"Not used2",0)
'     .AddFrameExtra 90,120,140,"conservative liberal",0,_
'         Array("Rubber on leftpost", &H01, "Rubber on right post", &H05)
    End With
  End If
  BigHurtOptions = vpmDips.ViewDipsExtra(BigHurtOptions)
  BigHurtSetOptions : SaveValue "BigHurt","Options",BigHurtOptions
End Sub
Set vpmShowDips = GetRef("BigHurtShowDips")
Sub BigHurtSetOptions
  'If BigHurtOptions And &H02 Then GICallback = Empty Else Set GICallback = GetRef("UpdateGI")
  'If BigHurtOptions And &H10 Then Controller.switch(5) = 1 Else Set Controller.switch(5) = 0
  'If BigHurtOptions And &H20 Then Controller.switch(6) = 1 Else Set Controller.switch(6) = 0
' Wall25.IsDropped = (BigHurtOptions And &H01) = 0
' Wall164.isdropped = (BigHurtOptions And &H05) = 0
End Sub


' Choose Side Blades
 if bladeArt = 1 then
    PinCab_Blades1.Image = "Sidewalls BH"
   PinCab_Blades1.visible = 1
  elseif bladeArt = 2 then
    PinCab_Blades1.visible = 0
  End if


'Const StartButton = 2
'ExtraKeyHelp = KeyName(StartButton) & vbTab & "StartButton"

'Keyboard handlers
Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftFlipperKey then
   Controller.Switch(82) = 1
   VR_Flipperbuttonleft.X = VR_Flipperbuttonleft.X +8
    end if
  If keycode = RightFlipperkey then
    Controller.Switch(83) = 1
   VR_FlipperbuttonRight.X = VR_FlipperbuttonRight.X -8
    End if
    If keycode = LeftFlipperKey and RightFlipperkey then
'    TBumper.Disabled = true
'    LBumper.Disabled = true
    end if
 If Keycode = StartGameKey Then  VR_Cab_StartButton.y = VR_Cab_StartButton.y -8
    If vpmKeyDown(keycode) Then Exit Sub
  If keycode = PlungerKey Then
    Plunger.Pullback
    TimerVRPlunger.enabled = true
    TimerVRPlunger2.enabled = False
 End if
  If keycode = 16 Then Controller.Switch(152) = 1
 If keycode = 17 Then Controller.Switch(6) = 0
 If keycode = 18 Then Controller.Switch(5) = 0
 if keycode = 19 then Show_Information

End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode = LeftFlipperKey then
    Controller.Switch(82) = 0
   VR_Flipperbuttonleft.X = VR_Flipperbuttonleft.X -8
  end if
  If keycode = RightFlipperkey then
   Controller.Switch(83) = 0
   VR_FlipperbuttonRight.X = VR_FlipperbuttonRight.X +8
    End if
  If keycode = LeftFlipperKey and RightFlipperkey then
'    TBumper.Disabled = FALSE
'    LBumper.Disabled = FALSE
    end if

 If Keycode = StartGameKey Then  VR_Cab_StartButton.y = VR_Cab_StartButton.y +8

  if vpmKeyUp(keycode) Then Exit Sub
  if keycode = PlungerKey Then
    Plunger.Fire
    TimerVRPlunger.enabled = False
   TimerVRPlunger2.enabled = True
  End if
  If keycode = 17 Then Controller.Switch(6) = 1
 If keycode = 16 Then Controller.Switch(152) = 0
 If keycode = 18 Then Controller.Switch(5) = 1

End Sub
Sub Show_Information()
  Dim info, rboxstyle
  rboxstyle = vbInformation
  info = "GAME INFORMATION"   & vbNewLine
  info = info & "1) PRESS Q FOR SLAMTILT" & vbNewLine
  info = info & "2) PRESS W TO OPEN THE COINDOOR" & vbNewLine
  info = info & "" & vbNewLine
  info = info & "INGAME OPTIONS" & vbNewLine
  info = info & "PRESS F6 FOR PLAYFIELD ADJUSTMENTS" & vbNewLine
  info = info & "" & vbNewLine
  info = info & "FACTORY RESET" & vbNewLine
  info = info & "1) PRESS 7 TO ENTER GAME MENU" & vbNewLine
  info = info & "2) SELECT GAME ADJUSTMENTS WITH LEFT FLIPPER" & vbNewLine
  info = info & "3) PRESS 7" & vbNewLine
  info = info & "4) PRESS 1 TO SELECT THE FACTORY RESET"& vbNewLine
  info = info & "5) PRESS 1 AGAIN TO ACTIVED" & vbNewLine
  info = info & "6) EXIT GAME AND LOAD AGAIN"
  MsgBox  info, rboxstyle, ""
end sub



'DropTargets
Sub Wall17_Hit  ()dtDropR.Hit 1 : End Sub
Sub Wall18_Hit  ()dtDropR.Hit 2 :  end Sub
Sub Wall19_Hit  ()dtDropR.Hit 3 : End Sub

Sub SolDropTarget(Enabled)
  If Enabled Then
  dtDropR.DropSol_On
  Wall17.IsDropped=0
  Wall18.IsDropped=0
  Wall19.IsDropped=0
  End If
End Sub


Sub SolDropT3(Enabled)
  If Enabled Then
  dtDrop3.DropSol_On
  Wall19.IsDropped=0
  'Controller.Switch(35)=1
  else
  Wall19.IsDropped=1

  End If
End Sub

Sub SolDropT2(Enabled)
  If Enabled Then
  dtDrop2.DropSol_On
  Wall18.IsDropped=0
  else
  Wall18.IsDropped=1
  'Controller.Switch(25)=1
  End If
End Sub

Sub SolDropT1(Enabled)
  If Enabled Then
  dtDrop1.DropSol_On
  Wall17.IsDropped=0
  else
  Wall17.IsDropped=1
  'Controller.Switch(15)=0
  End If
End Sub

Sub Trigger1_Hit(): Controller.Switch(20) = 1 : End Sub
Sub Trigger1_UnHit(): Controller.Switch(20) = 0 : End Sub

'addball to stacks
Sub Drain_Hit()
  PlaySoundAtVol "popper", Drain, 1
    bsTrough.AddBall Me
End Sub

'left upper Saucer
Sub Kicker11_Hit()
  PlaySoundAtVol "SolOn", ActiveBall, 1
    bsUpKicker.AddBall me
    Controller.Switch(22) = True
End Sub

'BottomRight Hole
'Sub Kicker1_Hit()
' bsRLowKicker.AddBall 0
'End Sub
Sub Kicker1_Hit()
  PlaySoundAtVol "SolOn", ActiveBall, 1
  Me.Enabled = true
  Controller.Switch(21) = True
Gate5.open = 0
End Sub

Sub TriggerGate_Hit():Gate5.open = 0:End Sub

Sub SolVUK(Enabled)
  If Enabled Then
    If Controller.Switch(21) = True Then
      PlaySoundAtVol "BallRel", Kicker1, 1
      Kicker1.DestroyBall
      Kicker1.Enabled = True
      DoVUK
            Controller.Switch(21) = 0
    End If
  End If
End Sub

Sub DoVUK ()
  With Kicker2.CreateBall :End With
  Timer11.Enabled=True
End Sub

Sub Timer11_Timer()
  Kicker2.Kick 250,2
Gate5.open = 1
vpmtimer.addtimer 500, "Gate5.open = 0'"
Timer11.Enabled=False
End Sub

'Balpen
Sub Kicker26_Hit()
  PlaySoundAtVol "Solenoid", ActiveBall, 1
  Kicker26.DestroyBall
  Kicker25.CreateBall
    Kicker25.Kick 180,2
    vpmtimer.addtimer 1000, "Gate5.open = 1'"
    YellowFlasher.duration 2, 3500, 0
    YellowFlasher1.duration 2, 3500, 0
End Sub

'Left up kicker               '
Sub Kicker5_Hit()
  bsLKicker.AddBall Me
End Sub

'homerun top playfield
Sub TriggerHR_Hit()
  PlaySoundAtVol "SolOn", ActiveBall, 1
    vpmTimer.PulseSwitch(100),0,""
  bsLKicker.AddBall Me
End Sub

'-------------------------
Sub LSling_SlingShot() : vpmTimer.PulseSwitch 12,0,0 : End Sub
Sub RSling_SlingShot() : vpmTimer.PulseSwitch 13,0,0 : End Sub

Sub TBumper_Hit() : vpmTimer.PulseSwitch 11,0,0 : End Sub
Sub LBumper_Hit() : vpmTimer.PulseSwitch 10,0,0 : End Sub


Sub RightOutlane_Hit()  : Controller.Switch(86) = 1 : End Sub
Sub RightOutlane_UnHit(): Controller.Switch(86) = 0 : End Sub
Sub RightInlane_Hit() : Controller.Switch(96) = 1 : End Sub
Sub RightInlane_UnHit() : Controller.Switch(96) = 0 : End Sub
Sub LeftOutlane_Hit() : Controller.Switch(85) = 1 : End Sub
Sub LeftOutlane_UnHit() : Controller.Switch(85) = 0 : End Sub
Sub LeftInlane_Hit()  : Controller.Switch(95) = 1 : End Sub
Sub LeftInlane_UnHit()  : Controller.Switch(95) = 0 : End Sub

'CaptiveBall rollover
Sub Trigger2_Hit():Controller.Switch(107) = 1:CapFlash.duration 1, 300, 0: End Sub
Sub Trigger2_UnHit()  : Controller.Switch(107) = 0 : End Sub

'rollover bumper section
Sub Trigger4_Hit()    : Controller.Switch(106) = 1 : End Sub
Sub Trigger4_UnHit()  : Controller.Switch(106) = 0 : End Sub

'Orbit's
Sub Trigger6_Hit()    : Controller.Switch(81) = 1 : End Sub
Sub Trigger6_UnHit()  : Controller.Switch(81) = 0 : End Sub
Sub Trigger5_Hit()    : Controller.Switch(80) = 1 : End Sub
Sub Trigger5_UnHit()  : Controller.Switch(80) = 0 : End Sub

'scoop ramp
Sub Trigger7_Hit()    : Controller.Switch(101) = 1 : End Sub
Sub Trigger7_UnHit()  : Controller.Switch(101) = 0 : End Sub

'bullpen
Sub Trigger18_Hit()   : Controller.Switch(116) = 1 : End Sub
Sub Trigger18_UnHit() : Controller.Switch(116) = 0 : End Sub

'left & right ramp
Sub Trigger9_Hit()    : Controller.Switch(90) = 1 : End Sub
Sub Trigger9_UnHit()  : Controller.Switch(90) = 0 : End Sub
Sub Trigger10_Hit()   : Controller.Switch(91) = 1 : End Sub
Sub Trigger10_UnHit() : Controller.Switch(91) = 0 : End Sub

'standup's
Sub Target11_Hit(): vpmTimer.PulseSwitch(115),0, "" : End Sub
Sub Target12_Hit(): vpmTimer.PulseSwitch(105),0, "" : End Sub

'backboards
Sub Wall1_Hit(): vpmTimer.PulseSwitch(87),0, "" :End Sub
'Sub Wall188_Hit(): vpmTimer.PulseSwitch(87),0, "" :End Sub
Sub Wall15_Hit(): vpmTimer.PulseSwitch(97),0, "" :End Sub


Dim BallSpeed

Sub glove1_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub glove2_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub glove3_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub glove4_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub glove5_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub glove6_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub glove7_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub glove8_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub glove9_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub glove10_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall178_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall179_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall180_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall181_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall182_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall183_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall184_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall185_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall186_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub Wall187_Hit
  With ActiveBall
  BallSpeed = Sqr(.VelX * .VelX + .VelY * .VelY)
  End With
  vpmTimer.PulseSwitch(23),0,"HighCatch"
End Sub

Sub HighCatch(enabled)
  If enabled And BallSpeed > 15 Then
    vpmTimer.PulseSwitch(117),0,""
  End If
End Sub



'HelpKickers (will be removed later)

'scoop
Sub Kicker13_Hit()
    Kicker13.DestroyBall
    With Kicker14.CreateBall :end with
    Kicker14.Kick 180,5
End Sub

 'leftramp
Sub Kicker7_Hit()
    Kicker7.Kick 178,1
End Sub

'scoopexit
Sub Kicker12_Hit()
    Kicker12.DestroyBall
     With Kicker18.CreateBall :end with
    Kicker18.Kick 345, 1
End Sub


'speed kicker / left upflipper
Sub Kicker19_Hit()
  Kicker19.Kick 250,8
End Sub

'ShooterSolenoid
Sub solShooter(enabled)
  if enabled then
    Plunger.Pullback
  else
      Plunger.Fire
  end if
end sub

'shooterlane ramp up/down
Sub SolShooterlane(enabled)
    if enabled then
    Ramp1.HeightBottom = 60
    Ramp1a.HeightBottom = 60.1
    Ramp1.collidable = 0
    PlaySoundAtVol "flapopen", Ramp1, 1
  else
    Ramp1.HeightBottom = 0
    Ramp1a.HeightBottom = 0.1
    Ramp1.collidable = 1
    PlaySoundAtVol "flapclos", Ramp1, 1
  end if
end sub

'Glove Mech & hit events
Dim Gloveposition
Dim Glovedirection

Sub glovemotion (aNewPos,aSpeed, alastPos)
Gloveposition=aNewPos

Select Case Gloveposition

Case 0
Glove1.isdropped=0
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 0
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 1
Glove1.isdropped=1
Glove1a.isdropped=0
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 0
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 2
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=0
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 0
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 3
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=0
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 0
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 4
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=0
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 0
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 5
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=0
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 0
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 6
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=0
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 0
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 7
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=0
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1


Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 0
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 8
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=0
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1


Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 0
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 9
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=0
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 0
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1
playsound "motor_long"
Case 10
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=0
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 0
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1
Stopsound "motor_long"
Case 11
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=0
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 0
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1
playsound "motor_long"
Case 12
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=0
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 0
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 13
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=0
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 0
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 14
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=0
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 0
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 15
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=0
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 0
Wall186.isdropped = 1
Wall187.isdropped = 1

Case 16
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=0
Glove9a.isdropped=1
Glove10.isdropped=1

Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 0
Wall187.isdropped = 1

Case 17
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=0
Glove10.isdropped=1


Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 0
Wall187.isdropped = 1

Case 18
Glove1.isdropped=1
Glove1a.isdropped=1
Glove2.isdropped=1
Glove2a.isdropped=1
Glove3.isdropped=1
Glove3a.isdropped=1
Glove4.isdropped=1
Glove4a.isdropped=1
Glove5.isdropped=1
Glove5a.isdropped=1
Glove6.isdropped=1
Glove6a.isdropped=1
Glove7.isdropped=1
Glove7a.isdropped=1
Glove8.isdropped=1
Glove8a.isdropped=1
Glove9.isdropped=1
Glove9a.isdropped=1
Glove10.isdropped=0


Wall178.isdropped = 1
Wall179.isdropped = 1
Wall180.isdropped = 1
Wall181.isdropped = 1
Wall182.isdropped = 1
Wall183.isdropped = 1
Wall184.isdropped = 1
Wall185.isdropped = 1
Wall186.isdropped = 1
Wall187.isdropped = 0

End Select

If aNewPos>aLastpos Then'we're going right
Glovedirection=1
End IF
If aNewPos<aLastpos Then'we're going left
Glovedirection=2
End If
End Sub


'======================================

Sub CenterRampLift(Enabled)
If  Enabled Then
Ramp5.HeightTop = 90
PlaySoundAtVol "flapopen", Ramp5, 1
Else
Ramp5.HeightTop = 75
PlaySoundAtVol "flapclos", Ramp5, 1
End If
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

Const tnob = 6 ' total number of balls
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )/5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )/7, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
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

 '**********************
'Flipper Shadows
'***********************
Sub RealTime_Timer
  lfs.RotZ = LeftFlipper.CurrentAngle
  rfs.RotZ = RightFlipper.CurrentAngle
LeftFlipperP.RotY = LeftFlipper.CurrentAngle - 90
RightFlipperP.RotY = RightFlipper.CurrentAngle - 90
ULFlipperP.RotY = Flipper3.CurrentAngle - 90
BallShadowUpdate
End Sub


Sub BallShadowUpdate()
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6)
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
    BallShadow(b).X = BOT(b).X
    ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 and BOT(b).Z < 200 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
if BOT(b).z > 30 Then
ballShadow(b).height = BOT(b).Z - 20
ballShadow(b).opacity = 110
Else
ballShadow(b).height = BOT(b).Z - 24
ballShadow(b).opacity = 90
End If
    Next
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.Stop:End Sub

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
 PlaySound "target", 0, Vol(ActiveBall)*8, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*8, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*8, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*8, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*8, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

'***************************************************************************
'Beer Bubble Code - Rawd
'***************************************************************************
Sub BeerTimer_Timer()

  Randomize(21)
 BeerBubble1.z = BeerBubble1.z + Rnd(1)*0.5
  if BeerBubble1.z > -771 then BeerBubble1.z = -955
 BeerBubble2.z = BeerBubble2.z + Rnd(1)*1
  if BeerBubble2.z > -768 then BeerBubble2.z = -955
 BeerBubble3.z = BeerBubble3.z + Rnd(1)*1
  if BeerBubble3.z > -768 then BeerBubble3.z = -955
 BeerBubble4.z = BeerBubble4.z + Rnd(1)*0.75
 if BeerBubble4.z > -774 then BeerBubble4.z = -955
 BeerBubble5.z = BeerBubble5.z + Rnd(1)*1
  if BeerBubble5.z > -771 then BeerBubble5.z = -955
 BeerBubble6.z = BeerBubble6.z + Rnd(1)*1
  if BeerBubble6.z > -774 then BeerBubble6.z = -955
 BeerBubble7.z = BeerBubble7.z + Rnd(1)*0.8
  if BeerBubble7.z > -768 then BeerBubble7.z = -955
 BeerBubble8.z = BeerBubble8.z + Rnd(1)*1
  if BeerBubble8.z > -771 then BeerBubble8.z = -955
End Sub


'***************************************************************************
'VR Clock code below - THANKS RASCAL
'***************************************************************************


' VR Clock code below....
Sub ClockTimer_Timer()
 dim CurrentMinute: Currentminute = Minute(Now())
  VRClockMinutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
 VRClockhours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
    VRClockseconds.RotAndTra2 = (Second(Now()))*6
 CurrentMinute=Minute(Now())
End Sub

'***************************************************************************
' VR Plunger Animataion Code
'***************************************************************************

Sub TimerVRPlunger_Timer
  if VR_PLungerrod.Y < (Plunger.y -1780) then VR_PLungerrod.Y = VR_PLungerrod.y +12
 if VR_PLungerKnob.Y < (Plunger.y - 1780) then VR_PLungerKnob.Y = VR_PLungerKnob.y +12
End Sub

Sub TimerVRPlunger2_Timer
 VR_PLungerrod.Y = Plunger.y -1870 + (5* Plunger.Position)
  VR_PLungerKnob.Y = Plunger.y -1870 + (5* Plunger.Position)
end sub


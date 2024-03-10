'************************************************************
'************************************************************
'
'  Hot Shots / IPD No. 1248 / April, 1989 / 4 Players
'
'   Credits:
'
'   VPX by randr, shannon
'
'
'************************************************************
'************************************************************

Option Explicit
Randomize

'******************************************************
'           STANDARD DEFINITIONS
'******************************************************


Const UseSolenoids=1
Const UseLamps=1
Const UseSync=1
Const UseGI=0

Const SSolenoidOn="SolOn"
Const SSolenoidOff="SolOff"
Const SFlipperOn="FlipperUp"
Const SFlipperOff="FlipperDown"
Const SCoin="coin3"

'******************************************************
'           TABLE INIT
'******************************************************

On Error Resume Next
  ExecuteGlobal GetTextFile("controller.vbs")
  If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01000100", "sys80.vbs", 2.31

Const cGameName="hotshots"

Dim VarRol, VarHidden
VarHidden=0

If Table1.ShowDT = true then VarRol=0 Else VarRol=1
If B2SOn = true Then VarHidden=1

Dim bsTrough,dtBlue1,dtRed1,dtYellow1,dtGreen1,dtTop,dtBottom

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

  On Error Resume Next
  With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine="Hot Shots"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=0
    .ShowFrame=0
    .ShowTitle=0
    .Hidden = VarHidden
    .Games(cGameName).Settings.Value("rol")=VarRol
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0

  Controller.SolMask(0)=0
  vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
  Controller.Run

  KickBack.PullBack
  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=57
  vpmNudge.Sensitivity=3
  vpmNudge.TiltObj=Array(Bumper1,LeftSlingshot,RightSlingshot)

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 66,0,76,0,0,0,0,0
  bsTrough.InitKick BallRelease,90,5
  bsTrough.InitExitSnd "BallRel", "SolOn"
  bsTrough.Balls=2

  Set dtBlue1=New cvpmDropTarget
  dtBlue1.InitDrop Array(SW00,SW10,SW20,SW30), Array(0,10,20,30)
  dtBlue1.InitSnd "","droptargetreset"
  dtBlue1.CreateEvents "dtBlue1"

  Set dtRed1=New cvpmDropTarget
  dtRed1.InitDrop Array(SW40,SW50,SW60,SW70),Array(40,50,60,70)
  dtRed1.InitSnd "","droptargetreset"
  dtRed1.CreateEvents "dtRed1"

  Set dtYellow1=New cvpmDropTarget
  dtYellow1.InitDrop Array(SW1,SW11,SW21,SW31),Array(1,11,21,31)
  dtYellow1.InitSnd "","droptargetreset"
  dtYellow1.CreateEvents "dtYellow1"

  Set dtGreen1=New cvpmDropTarget
  dtGreen1.InitDrop Array(SW41,SW51,SW61,SW71),Array(41,51,61,71)
  dtGreen1.InitSnd "","droptargetreset"
  dtGreen1.CreateEvents "dtGreen1"

  Set dtTop=New cvpmDropTarget
  dtTop.InitDrop SW2,Array(2)
  dtTop.InitSnd "","droptargetreset"
  dtTop.CreateEvents "dtTop"

  Set dtBottom=New cvpmDropTarget
  dtBottom.InitDrop SW3,Array(3)
  dtBottom.InitSnd "","droptargetreset"
  dtBottom.CreateEvents "dtBottom"

  'vpmCreateEvents AllSwitches
  vpmMapLights ALights
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'******************************************************
'             KEYS
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
  If KeyCode=PlungerKey Then Plunger.Pullback:PlaySoundat "plungerpull", Plunger
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If KeyCode=PlungerKey Then Plunger.Fire:PlaySoundat "plunger", Plunger
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'******************************************************
'             SOLENOIDS
'******************************************************

Dim xx, objekt

SolCallback(1) = "Solone" 'Blue drop targets
SolCallback(2) = "SolTwo" 'Red drop targets
SolCallback(3) = "SolThree" 'Kicker
Solcallback(4) = "SolFour" 'Center flippers drop target
SolCallback(5) = "SolFive" 'Yellow drop targets
SolCallback(6) = "SolSix" 'Green Drop targets
SolCallback(7) = "SolSeven" 'Top Drop
SolCallback(8) = "vpmSolSound""knocker"","
SolCallback(9) = "bsTrough.SolIn"
SolCallback(10) = "SU"

'******************************************************
'             BLUE DROPS
'******************************************************

Sub SolOne(enabled)
  if enabled then
    If OL12=0 Then
      dtBlue1.DropSol_On'Breset.enabled=True
    Else
      vpmFlasher Array(Light50,Light51),True
    End If
  Else
    If OL12=1 Then vpmFlasher Array(Light50,Light51),False
  End If

End Sub


Sub Breset_timer
  PlaySoundAtVol "dropbankreset", SW00, 1
  For each objekt in DTBlue: objekt.isdropped=0: next
       Breset.enabled=false
End Sub


'******************************************************
'             RED DROPS
'******************************************************

Sub SolTwo(enabled)'Red Drop Targets
  if enabled then
    If OL12=0 Then
    dtRed1.DropSol_On'Rreset.enabled=True
  Else
      vpmFlasher Array(Light38,Light36),True
    End If
  Else
    If OL12=1 Then vpmFlasher Array(Light37,Light36),False
  End If

End Sub

Sub Rreset_timer
  PlaySoundAtVol "dropbankreset", SW40, 1
  For each objekt in DTRed: objekt.isdropped=0: next
  Rreset.enabled=false
End Sub

'******************************************************
'             KICKER
'******************************************************

Sub SolThree(Enabled)
  controller.Switch(56) = 0
  If Kicker1.BallCntOver = 0 Then
    PlaySoundAt SoundFX("SolOn",DOFContactors), Kicker1
  Else
    PlaySoundAt SoundFX("fx_scooopExit",DOFContactors), Kicker1
  End If
  Kicker1.kickz 180,10,0,65
End Sub

Sub Kicker1_Hit
  PlaySoundAt "kicker_enter", kicker1
  Controller.Switch(56) = 1
End Sub

'******************************************************
'           CENTER DROP
'******************************************************

Sub SolFour(Enabled)
  If Enabled Then
    If OL12=0 Then
      dtBottom.DropSol_On
    Else
      vpmFlasher Array(Bumper6,Bumper7),True
    End If
  Else
    If OL12=1 Then vpmFlasher Array(Bumper6,Bumper7),False
  End If
End Sub

'******************************************************
'           YELLOW DROPS
'******************************************************

Sub SolFive(enabled)'Yellow Drop Targets
  if enabled then 'Yreset.enabled=True
    FlashLevel1 = 1:FlashLevel2 = 1:Flasherflash1.visible = 1:FlasherFlash1.timerenabled=true:Flasherflash2.visible = 1:FlasherFlash2.timerenabled=true
    If OL12=0 Then
      dtYellow1.DropSol_On
    End If
  Else
    'If OL12=1 Then FlashLevel1 = 0:FlashLevel2 = 0:FlasherFlash1.timerenabled=false:FlasherFlash2.timerenabled=false
  End If
End Sub

Sub Yreset_timer
  PlaySoundAtVol "dropbankreset", SW1, 1
  For each objekt in DTYellow: objekt.isdropped=0: next
  yreset.enabled=false
End Sub

Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4
FlasherLight1.IntensityScale = 0
Flasherlight2.IntensityScale = 0
Flasherlight3.IntensityScale = 0
Flasherlight4.IntensityScale = 0

'*** white flashers ***
Sub FlasherFlash1_timer()
  dim flashx3, matdim

  FlasherFlash1.visible = 1

  flashx3 = FlashLevel1 * FlashLevel1 * FlashLevel1
  Flasherflash1.opacity = 1000 * flashx3^0.8
  FlasherLit1.BlendDisableLighting = 10 * flashx3
  Flasherbase1.BlendDisableLighting =  flashx3 + 0.2
  FlasherLight1.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel1)
  if matdim = 10 then matdim = 9
  FlasherLit1.material = "domelit" & matdim
  FlashLevel1 = FlashLevel1 * 0.8 - 0.01
  If FlashLevel1 < 0.15 Then
    FlasherLit1.visible = 0
  Else
    FlasherLit1.visible = 1
  end If
  If FlashLevel1 < 0 Then
    FlashLevel1 = 0
    FlasherFlash1.TimerEnabled = False
    FlasherFlash1.visible = 0
    Flasherbase1.BlendDisableLighting = 0.2
  End If
End Sub


Sub FlasherFlash2_timer()
  dim flashx3, matdim

  FlasherFlash2.visible = 1

  flashx3 = FlashLevel2 * FlashLevel2 * FlashLevel2
  Flasherflash2.opacity = 1000 * flashx3^0.8
  FlasherLit2.BlendDisableLighting = 10 * flashx3
  Flasherbase2.BlendDisableLighting =  flashx3 + 0.2
  FlasherLight2.IntensityScale = flashx3
  if matdim = 10 then matdim = 9
  matdim = Round(10 * FlashLevel2)
  FlasherLit2.material = "domelit" & matdim
  FlashLevel2 = FlashLevel2 * 0.8 - 0.01
  If FlashLevel2 < 0.15 Then
    FlasherLit2.visible = 0
  Else
    FlasherLit2.visible = 1
  end If
  If FlashLevel2 < 0 Then
    FlashLevel2 = 0
    FlasherFlash2.TimerEnabled = False
    FlasherFlash2.visible = 0
    Flasherbase2.BlendDisableLighting = 0.2
  End If
End Sub

'******************************************************
'           GREEN DROPS
'******************************************************

Sub SolSix(enabled)'green Drop Targets
  if enabled then 'Greset.enabled=True
    FlashLevel3 = 1:FlashLevel4 = 1:Flasherflash3.visible = 1:FlasherFlash3.timerenabled=true:Flasherflash4.visible = 1:FlasherFlash4.timerenabled=true
    If OL12=0 Then
      dtGreen1.DropSol_On
    End If
  Else
    'If OL12=1 Then FlashLevel3 = 0:FlashLevel4 = 0:FlasherFlash3.timerenabled=False:FlasherFlash4.timerenabled=False
  End If
End Sub

Sub Greset_timer
PlaySoundAtVol "dropbankreset", SW41, 1
     For each objekt in DTGreen: objekt.isdropped=0: next
    Greset.enabled=false
End Sub


Sub FlasherFlash3_timer()
  dim flashx3, matdim

  FlasherFlash3.visible = 1

  flashx3 = FlashLevel3 * FlashLevel3 * FlashLevel3
  Flasherflash3.opacity = 1000 * flashx3^0.8
  FlasherLit3.BlendDisableLighting = 10 * flashx3
  Flasherbase3.BlendDisableLighting =  flashx3 + 0.2
  FlasherLight3.IntensityScale = flashx3
  if matdim = 10 then matdim = 9
  matdim = Round(10 * FlashLevel3)
  FlasherLit3.material = "domelit" & matdim
  FlashLevel3 = FlashLevel3 * 0.8 - 0.01
  If FlashLevel3 < 0.15 Then
    FlasherLit3.visible = 0
  Else
    FlasherLit3.visible = 1
  end If
  If FlashLevel3 < 0 Then
    FlashLevel3 = 0
    FlasherFlash3.TimerEnabled = False
    FlasherFlash3.visible = 0
    Flasherbase3.BlendDisableLighting = 0.2
  End If
End Sub

Sub FlasherFlash4_timer()
  dim flashx3, matdim

  FlasherFlash4.visible = 1

  flashx3 = FlashLevel4 * FlashLevel4 * FlashLevel4
  Flasherflash4.opacity = 1000 * flashx3^0.8
  FlasherLit4.BlendDisableLighting = 10 * flashx3
  Flasherbase4.BlendDisableLighting =  flashx3 + 0.2
  FlasherLight4.IntensityScale = flashx3
  if matdim = 10 then matdim = 9
  matdim = Round(10 * FlashLevel4)
  FlasherLit4.material = "domelit" & matdim
  FlashLevel4 = FlashLevel4 * 0.8 - 0.01
  If FlashLevel4 < 0.15 Then
    FlasherLit4.visible = 0
  Else
    FlasherLit4.visible = 1
  end If
  If FlashLevel4 < 0 Then
    FlashLevel4 = 0
    FlasherFlash4.TimerEnabled = False
    FlasherFlash4.visible = 0
    Flasherbase4.BlendDisableLighting = 0.2
  End If
End Sub

'******************************************************
'           TOP DROP
'******************************************************

Sub SolSeven(Enabled)
  If Enabled Then
    If OL12=0 Then
      dtTop.DropSol_On
    Else
      vpmFlasher Array(Bumper8,Bumper9),True
    End If
  Else
    If OL12=1 Then vpmFlasher Array(Bumper8,Bumper9),False
  End If
End Sub

'******************************************************
'           ACCELERATOR
'******************************************************

Dim SolA
SolA=0

Sub SU(Enabled)
  If Enabled then
    SolA=1:PlaysoundAt "Electric_Motor", trigger7
  Else
    SolA=0
  End If
End Sub

Dim MyBall,MyBallX,MyBallY

Sub Trigger7_Hit
  Set MyBall=ActiveBall
  MyBallX=MyBall.VelX
  MyBallY=MyBall.VelY
  If SolA=1 Then

    if Int(Rnd*4)+1 <> 1 then
      MyBall.VelX=MyBall.VelX*1.9
      MyBall.VelY=MyBall.VelY*1.9
    end if
  Else
    MyBall.VelX=MyBall.VelX*1
    MyBall.VelY=MyBall.VelY*1
  End If
End Sub

Sub Trigger11_Hit
' debug.print activeball.velx
  If activeball.velx > -6 Then
    activeball.velx = 0
    activeball.vely = 5
  end if
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub Drain_Hit:PlaySoundat "drain", Drain:bsTrough.AddBall Me:End Sub

'******************************************************
'         FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("fx_FlipperUp",DOFContactors), LeftFlipper
    LeftFlipper.RotateToEnd
  Else
    if leftflipper.currentangle < leftflipper.startangle - 5 then
      PlaySoundAt SoundFX("fx_FlipperDown",DOFContactors), LeftFlipper
    end if
    LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("fx_FlipperUp",DOFContactors), RightFlipper
    RightFlipper.RotateToEnd
  Else
    if RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      PlaySoundAt SoundFX("fx_FlipperDown",DOFContactors), RightFlipper
    End If
    RightFlipper.RotateToStart

  End If
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


'******************************************************
'           LIGHTS
'******************************************************

set lights(3)=Light3
set lights(4)=light4
set lights(5)=light5
set lights(6)=light6
set lights(7)=light7
set lights(8)=light8
set lights(9)=light9
set lights(10)=light10
set lights(11)=light11
set lights(16)=light16
set lights(17)=light17
set lights(17)=light17a
set lights(18)=light18
set lights(19)=light19
set lights(20)=light20
set lights(21)=light21
set lights(22)=light22
set lights(22)=light22a
set lights(23)=light23
set lights(23)=light23a
set lights(24)=light24
set lights(25)=light25
set lights(26)=light26
set lights(27)=light27
set lights(28)=light28
set lights(29)=light29
set lights(30)=light30
set lights(31)=light31
set lights(32)=light32
set lights(33)=light33
set lights(34)=light34
set lights(35)=light35
set lights(36)=light36
set lights(37)=light37
set lights(38)=light38
set lights(39)=light39
set lights(40)=light40
set lights(41)=light41
set lights(42)=light42
set lights(43)=light43
set lights(44)=light44
set lights(45)=light45
set lights(46)=light46
set lights(47)=light47


'****Switches ****

Sub LeftInlane_Hit:Controller.Switch(52)=1:End Sub
Sub LeftInlane_unHit:Controller.switch(52)=0:End Sub
Sub LeftOutlane_Hit:Controller.Switch(42)=1:End Sub
Sub LeftOutlane_unHit:Controller.Switch(42)=0:KickBack.PullBack:End Sub
Sub RightInlane_Hit:Controller.Switch(53)=1:End Sub
Sub RightInlane_unHit:Controller.Switch(53)=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(43)=1:End Sub
Sub RightOutlane_unHit:Controller.Switch(43)=0:End Sub
Sub Trigger1_Hit:Controller.Switch(46)=1:End Sub
Sub Trigger1_unHit:Controller.Switch(46)=0:End Sub
Sub Trigger3_Hit:Controller.Switch(63)=1:End Sub
Sub Trigger3_unHit:Controller.Switch(63)=0:End Sub
Sub Trigger2_Hit:Controller.Switch(62)=1:End Sub
Sub Trigger2_unHit:Controller.Switch(62)=0:End Sub
Sub Trigger4_Hit:Controller.Switch(5)=1:End Sub
Sub Trigger4_unHit:Controller.Switch(5)=0:End Sub
Sub Trigger5_Hit:Controller.Switch(25)=1:End Sub
Sub Trigger5_unHit:Controller.Switch(25)=0:End Sub

'Sub Trigger12_Hit:Controller.Switch(15)=1:End Sub
'Sub Trigger12_unHit:Controller.Switch(15)=0:End Sub
Sub Trigger8_Hit:Controller.Switch(15)=1:End Sub
Sub Trigger8_unHit:Controller.Switch(15)=0:End Sub
' NEED TO ADD MORE SWITCHES YET

'****Targets***

Sub Target1_hit:vpmTimer.PulseSw 35 End Sub
Sub Target2_hit:vpmTimer.PulseSw 13 End Sub
Sub Target3_hit:vpmTimer.PulseSw 23 End Sub
Sub Target4_hit:vpmTimer.PulseSw 33 End Sub
Sub Target5_hit:vpmTimer.PulseSw 32 End Sub
Sub Target6_hit:vpmTimer.PulseSw 22 End Sub
Sub Target7_hit:vpmTimer.PulseSw 12 End Sub
Sub Target8_hit:vpmTimer.PulseSw 45 End Sub
Sub Target9_hit:vpmTimer.PulseSw 55 End Sub

'*****************************************
' Drop Targets
'*****************************************

Sub SW3_dropped:controller.switch(3)=1:PlaySoundAtBall "droptarget":End Sub       'center Flippers

Sub SW2_dropped:controller.switch(2)=1:PlaySoundAtBall "droptarget":End Sub       'upper target

Sub SW70_dropped:controller.switch(70)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW60_dropped:controller.switch(60)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW50_dropped:controller.switch(50)=1:PlaySoundAtBall "droptarget":End Sub       'left side
Sub SW40_dropped:controller.switch(40)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW30_dropped:controller.switch(30)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW20_dropped:controller.switch(20)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW10_dropped:controller.switch(10)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW00_dropped:controller.switch(0)=1:PlaySoundAtBall "droptarget":End Sub

Sub SW1_dropped:controller.switch(1)=1:PlaySoundAtBall "droptarget":End Sub       'right side
Sub SW11_dropped:controller.switch(11)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW21_dropped:controller.switch(21)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW31_dropped:controller.switch(31)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW41_dropped:controller.switch(41)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW51_dropped:controller.switch(51)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW61_dropped:controller.switch(61)=1:PlaySoundAtBall "droptarget":End Sub
Sub SW71_dropped:controller.switch(71)=1:PlaySoundAtBall "droptarget":End Sub


'**** Primitive Flipper Code ****

Sub FlipperTimer_Timer()
  'dim PI:PI=3.1415926
  LFlip.rotz = LeftFlipper.currentangle
  RFlip.rotz = RightFlipper.currentangle
  LflipRubber.rotz = LeftFlipper.currentangle
  RFlipRubber.rotz = RightFlipper.currentangle
  FlipperLSh.rotz = LeftFlipper.currentangle
  FlipperRSh.rotz = RightFlipper.currentangle
End Sub

'*****************************************
' Pop Bumpers
'*****************************************

Sub Bumper1_Hit
    vpmTimer.PulseSw 65
    PlaySound"fx_bumper1"
End Sub

'****************
' Sling Shots
'****************

Sub LeftSlingShot_Slingshot
  VpmTimer.PulseSw 75
  Playsound "SlingL"
  LSling.Visible = 0
    LSling1.Visible = 1
  slingL.rotx = 20
    me.uservalue = 1
    Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case me.uservalue
       Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.rotx = 10
       Case 4:slingL.rotx = 0:LSLing2.Visible = 0:LSLing.Visible = 1:me.TimerEnabled = 0
    End Select
    me.uservalue = me.uservalue + 1
End Sub

Sub RightSlingShot_Slingshot
  VpmTimer.PulseSw 75
  Playsound "SlingR"
  RSling.Visible = 0
    RSling1.Visible = 1
  slingR.rotx = 20
    me.uservalue = 1
    Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case me.uservalue
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.rotx = 10
        Case 4:slingR.rotx = 0:RSLing2.Visible = 0:RSLing.Visible = 1:me.TimerEnabled = 0
    End Select
    me.uservalue = me.uservalue + 1
End Sub





Set LampCallback=GetRef("UpdateMultipleLamps")
Dim L2,OL,L12,OL12,L13,L14,OL13,OL14,L15,OL15
OL=0:OL12=1:OL13=1:OL14=1:OL15=1

Sub UpdateMultipleLamps
  L12=Light1.State
  If L12<>OL12 Then OL12=L12
  L2=Light2.State 'BALL RELEASE
  If L2<>OL Then
    If L2=0 Then
      If bsTrough.Balls Then bsTrough.ExitSol_On
    End If
  End If
  OL=L2
  L13=Light333.State 'Top Rollover Trip
  If L13<>OL13 Then
    If L13=0 Then
      dtTop.Hit 1
    End If
  End If
  OL13=L13
  L14=Light4.State 'Bottom Rollover Trip
  If L14<>OL14 Then
    If L14=0 Then
      dtBottom.Hit 1
    End If
  End If
  OL14=L14
  L15=Light50.State 'Bottom Left Shooter
  If L15<>OL15 Then
    If L15=0 Then
      KickBack.PullBack
      Else
      KickBack.Fire:Playsound "ballsaver"
    End If
  End If
  OL15=L15

End Sub


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Deleted obsoleted subs

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / table1.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / table1.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2)'*1.2
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
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

Const tnob = 2 ' total number of balls
ReDim rolling(tnob)
InitRolling

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3)

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
    StopSound("fx_Rolling_Plastic" & b)
    BallShadow(b).visible = 0
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)
    If BallVel(BOT(b) ) > 1 Then

      ' ***Ball on playfield***
      if BOT(b).z < 27 Then
        PlaySound("fx_ballrolling"& b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        rolling(b) = True
        StopSound("fx_Rolling_Plastic" & b)
      ' ***Ball on ramp***
      Elseif  BOT(b).z > 50  Then
        PlaySound("fx_Rolling_Plastic" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        rolling(b) = True
        StopSound("fx_ballrolling"& b)
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling")
          StopSound("fx_Rolling_Plastic" & b)
          rolling(b) = False
        end If
      End If
    Else
      If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          StopSound("fx_Rolling_Plastic" & b)
          rolling(b) = False
      End If
    End If

    '***Ball Shadows***
    If BOT(B).z < 30 Then
      BallShadow(b).X = BOT(b).X + (BOT(b).X - Table1.Width/2)/12
      BallShadow(b).Y = BOT(b).Y + 10
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If

    '***Ball Drop Sounds***

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
' Object Hit Sounds
'**********************
Sub a_Woods_Hit(idx)
  PlaySound "fx_Woodhit", 0, Vol(ActiveBall), Audiopan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Triggers_Hit(idx)
  PlaySound "fx_sensor", 0, Vol(ActiveBall), Audiopan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
  PlaySound SoundFX("fx_target",DOFTargets), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_DropTargets_Hit (idx)
  PlaySound "droptarget", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub


Sub a_Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub a_Posts_Hit(idx)
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



'*****************************************
'           BALL SHADOW by ninnuzu
'*****************************************
'Dim BallShadow
'BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)
'
'Sub BallShadowUpdate_timer()
'    Dim BOT, b
' Dim maxXoffset
' maxXoffset=15
'    BOT = GetBalls
'
' ' render the shadow for each ball
'    For b = 0 to UBound(BOT)
'   BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(b).X)/(Eclipse.Width/2))
'   BallShadow(b).Y = BOT(b).Y + 10
'   If BOT(b).Z > 0 and BOT(b).Z < 30 Then
'     BallShadow(b).visible = 1
'   Else
'     BallShadow(b).visible = 0
'   End If
' Next
'End Sub

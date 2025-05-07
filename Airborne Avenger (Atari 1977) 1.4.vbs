Option Explicit
Randomize

Const BallSize = 48
Const BallMass = 1.1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01200000", "ATARI1.VBS", 3.1

' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


Dim BallSound, Decal, AltColor, PlayNums, GatePos

'Dim F1Xon, f1xoff, f1on, f1off, F2Xon, f2xoff, f2on, f2off, F3Xon, f3xoff, f3on, f3off

'***************************REMOVE DECAL IN CABINET MODE*****************
'****************Change value below to Decal=0 to remove Decal***********
Decal=1
'************************************************************************
'************************************************************************

'***************************ALTERNATE COLOR PLAYER LIGHTS****************
'*****Change value below to AltColor=1 to use alternate light colors*****
AltColor=0
'************************************************************************
'************************************************************************

'***************************NUMBERS ON PLAYER LIGHTS*********************
'*************Change value below to PlayNums=1 to show numbers***********
PlayNums=0
'************************************************************************
'************************************************************************


' Thalamus - this sub was called twice - I'm disabling this one, since that
' is what is going to happen anyway.

' Sub Table1_init()
'
' GatePos=0
'
' FlashTimer
'   BallSound=1
'   FlashSync.Enabled=True
' BumpTimer.Enabled=False
' End Sub

' Thalamus : Was missing 'vpminit me'
  vpminit me

Const cGameName     = "aavenger"   ' PinMAME short name
Const cCredits      = "VPX table by DarthMarino"
Const UseSolenoids  = True
Const UseLamps      = True
Const UseGI         = False

' Standard Sounds
Const SSolenoidOn   = ""
Const SSolenoidOff  = ""
Const SFlipperOn    = "FlipperUp"
Const SFlipperOff   = "FlipperDown"
Const SCoin         = "coin3"

'Solenoid Definitions
Const sGate     = 3'1
'LeftFlipper=11'2
'RightFlipper=1'3
Const sOutHole    = 10'4
Const sLSling   = 9'5
Const sRSling   = 13'6
Const sEject    = 5'7
Const sSaucer1    = 8'8
Const sSaucer2    = 14'9
Const sREject   = 12'10
Const sTLJet    = 4'11
Const sTRJet    = 16'12
Const sBJet     = 2'13
Const sKnocker    = 17
'Const sEnable    = 18
'Const sCLO     = 19



'Solenoid Callbacks
SolCallback(sGate)    = "SolGate"
SolCallback(sOutHole) = "bsTrough.SolOut"
SolCallback(sLSling)  = "vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(sRSling)  = "vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(sEject)   = "bsEject.SolOut"
SolCallback(sSaucer1) = "bsSaucer1.SolOut"
SolCallback(sSaucer2) = "bsSaucer2.SolOut"
SolCallback(sREject)  = "bsREject.SolOut"
SolCallback(sTLJet)   = "vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(sTRJet)   = "vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(sBJet)    = "vpmSolSound SoundFX(""Jet3"",DOFContactors),"
'SolCallback(sEnable) = "vpmNudge.SolGameOn"
SolCallback(sKnocker) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"

'SolCallback(11)      = "vpmSolFlipper LeftFlipper ,nothing,"
'SolCallback(1)     = "vpmSolFlipper RightFlipper,nothing,"

'Smoothed Solenoid routines for gates

' Sub SolGate(Enabled)
'   If Not LeftDelay.Enabled Then vpmSolDiverter Gate,True,True:PlaySound "savegate"
'   LeftDelay.Enabled=0
'   LeftDelay.Enabled=1
' End Sub
'
' Sub LeftDelay_Timer:LeftDelay.Enabled=0:vpmSolDiverter Gate,True,False:PlaySound "savegate":End Sub
' 'End of solenoid smoothing for gates

Sub SolGate(Enabled)
  If GatePos=0 then
    vpmSolDiverter Gate,True,True:GatePos=1:playsound "savegate"
  Else
    vpmSolDiverter Gate,True,False:GatePos=0:playsound "savegate"
  End If
End Sub

Dim bsTrough,bsSaucer1,bsSaucer2,bsEject,bsREject,cbCaptive, wallsoundrandom

Sub Table1_Init
    On Error Resume Next
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = cCredits
    .HandleMechanics = 0
    .ShowDMDOnly = 1
    .ShowFrame = False
    .ShowTitle = False
    .Run
    .Hidden=1
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=True
    ' Nudging
  vpmNudge.TiltSwitch = 18
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj=Array(RightSlingshot,LeftSlingshot,topsling)

Captured.createball
Captured.kick 180,1

  Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,52,0,0,0,0,0,0
    bsTrough.InitKick BallRelease,90,5
    bsTrough.InitExitSnd SoundFX("Ballrel",DOFContactors),SoundFX("solon",DOFContactors)
    bsTrough.Balls=1
    GatePos=0

  Set bsSaucer1=New cvpmBallStack
    bsSaucer1.InitSaucer lmkicker,41,65,10
    bsSaucer1.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("popper",DOFContactors)

  Set bsSaucer2=New cvpmBallStack
    bsSaucer2.InitSaucer rmkicker,48,208,7
    bsSaucer2.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("popper",DOFContactors)

  Set bsEject=New cvpmBallStack
    bsEject.InitSaucer lkicker,47,-10,26.5
    bsEject.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("popper",DOFContactors)
    bsEject.KickForceVar=2

  Set bsREject=New cvpmBallStack
    bsREject.InitSaucer rkicker,46,-20,24
    bsREject.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("popper",DOFContactors)
    bsREject.KickForceVar=2

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then
  FSbar.visible=False
Else
  If Decal=1 Then
    FSbar.visible=True
  Else
    FSbar.visible=False
  End If
end If

End Sub


Sub CredLight
If light176.state=1 and light169.state=1 or light176.state=0 and light177.state=1 and light178.state=1 and light179.state=1 and light180.state=1 and light181.state=1 and light182.state=1 and light169.state=0 and light170.state=1 and light171.state=1 and light172.state=1 and light173.state=1 and light174.state=1 and light175.state=1 then
creditlight.state=0
else
creditlight.state=1
End if
End Sub

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=LeftFlipperKey And MacTilt=0 Then
    Controller.Switch(19)=1
    LeftFlipper.RotateToEnd
    PlaySoundAtVol"flipperup", LeftFlipper, VolFlip
    LeftF=1
  End If
  If KeyCode=RightFlipperKey And MacTilt=0 Then
    Controller.Switch(20)=1
    RightFlipper.RotateToEnd
    PlaySoundAtVol"flipperup", RightFlipper, VolFlip
    RightF=1
  End If
  If vpmKeyDown(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then Plunger.PullBack
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=LeftFlipperKey And MacTilt=0 Then
    Controller.Switch(19)=0
    LeftFlipper.RotateToStart
    PlaySoundAtVol"flipperdown", LeftFlipper, VolFlip
    LeftF=0
  End If
  If KeyCode=RightFlipperKey And MacTilt=0 Then
    Controller.Switch(20)=0
    RightFlipper.RotateToStart
    PlaySoundAtVol"flipperdown", RightFlipper, VolFlip
    RightF=0
  End If
  If vpmKeyUp(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then PlaySoundAtVol"plunger",plunger,1:Plunger.Fire
End Sub


Sub MSunhit_hit
Controller.Switch(58)=0
End Sub

Sub midtarget_Hit:playsoundAtVol "target",ActiveBall,VolTarg:vpmTimer.PulseSw 21:End Sub    'switch 21
Sub Target3_Hit:playsoundAtVol "target", ActiveBall, VolTarg:vpmTimer.PulseSw 22:End Sub          'switch 22
Sub Target2_Hit:playsoundAtVol "target", ActiveBall, VolTarg::vpmTimer.PulseSw 23:End Sub         'switch 23
Sub Target1_Hit:playsoundAtVol "target", ActiveBall, VolTarg::vpmTimer.PulseSw 24:End Sub
Sub Trigger3_hit:playsoundAtVol "wallhit2", ActiveBall, 1:End Sub

Sub Gate1_hit:playsoundAtVol "gate", gate1, VolGates: End Sub
Sub Gate2_hit:playsoundAtVol "gate", gate2, VolGates: End Sub 'switch 24
Sub out3_Hit:Controller.Switch(34)=1::End Sub       'switch 34
Sub out3_unHit:Controller.Switch(34)=0:End Sub
Sub out2_Hit:Controller.Switch(35)=1::End Sub       'switch 35
Sub out2_unHit:Controller.Switch(35)=0:End Sub
Sub out1_Hit:Controller.Switch(36)=1::End Sub       'switch 36
Sub out1_unHit:Controller.Switch(36)=0:End Sub
Sub in2_Hit:Controller.Switch(37)=1::End Sub          'switch 37
Sub in2_unHit:Controller.Switch(37)=0:End Sub
Sub in1_Hit:Controller.Switch(38)=1::End Sub          'switch 38
Sub in1_unHit:Controller.Switch(38)=0:End Sub
Sub out4_Hit:Controller.Switch(39)=1::End Sub       'switch 39
Sub out4_unHit:Controller.Switch(39)=0:End Sub
Sub out6_Hit:Controller.Switch(39)=1::End Sub
Sub out6_unHit:Controller.Switch(39)=0:End Sub
Sub out5_Hit:Controller.Switch(40)=1::End Sub
Sub out5_unHit:Controller.Switch(40)=0:End Sub
Sub lmkicker_Hit:playsound "kicker":bsSaucer1.AddBall 0:End Sub       'switch 41
Sub Bumper1_Hit : vpmTimer.PulseSw(42) :  playsoundAtVol "fx_bumper1",bumper1,VolBump:End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(43) :  playsoundAtVol "fx_bumper1",bumper2,VolBump:End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(44) :  playsoundAtVol "fx_bumper1",bumper3,VolBump:End Sub

Sub Spinner_Hit
  PlaysoundAtVol "droptarget2", Spinnner, VolSpin
End Sub

Sub T3000unhit1_hit
controller.switch(60)=0
End Sub
Sub T3000unhit2_hit
controller.switch(60)=0
End Sub
Sub topunhit1_hit
controller.switch(60)=0
End Sub
Sub topunhit2_hit
controller.switch(60)=0
End Sub
Sub MSBunhit1_hit
controller.switch(59)=0
End Sub
Sub MSBunhit2_hit
controller.switch(59)=0
End Sub
Sub MSBunhit3_hit
controller.switch(59)=0
End Sub


Sub Spinner_Spin:vpmTimer.PulseSw 45:End Sub        'switch 45
Sub rkicker_Hit:playsoundAtVol "kicker",rkicker,VolKick:bsREject.AddBall 0:End Sub          'switch 46
Sub lkicker_Hit:playsoundAtVol "kicker",lkicker,VolKick:bsEject.AddBall 0:End Sub         'switch 47
Sub rmkicker_Hit:playsoundAtVol "kicker",rmkicker,VolKick:bsSaucer2.AddBall 0:End Sub       'switch 48
Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSwitch (49), 0, ""
  LSFlash.state=1
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LSFlash.state=0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub
Sub topsling_slingshot                    'switch 50
  playsoundAtVol"sling", ActiveBall, 1
  vpmTimer.PulseSw 51
End Sub
Sub bottomsling_slingshot
  playsoundAtVol"sling", ActiveBall, 1
  vpmTimer.PulseSw 57
End Sub
Sub rmsling_slingshot
  playsoundAtVol"sling", ActiveBall, 1
  vpmTimer.PulseSw 50
End Sub
Sub lmsling_slingshot
  playsoundAtVol"sling", ActiveBall, 1
  vpmTimer.PulseSw 50
End Sub
Sub ULsling_slingshot
  playsoundAtVol"sling", ActiveBall, 1
  vpmTimer.PulseSw 50
End Sub
Sub Wall73_slingshot
  If MacTilt=0 Then
  playsoundAtVol"sling", ActiveBall, 1
  vpmTimer.PulseSw 50
  End If
End Sub
Sub Wall74_slingshot
  If MacTilt=0 Then
  playsoundAtVol"sling", ActiveBall, 1
  vpmTimer.PulseSw 50
  End If
End Sub
Sub TRsling_slingshot
  playsoundAtVol"sling", ActiveBall, 1
  vpmTimer.PulseSw 51
End Sub
Sub TLsling_slingshot
  playsoundAtVol"sling", Activeball, 1
  vpmTimer.PulseSw 51
End Sub
Sub Trigger1_Hit:Controller.Switch(51)=1::End Sub     'switch 51
Sub Trigger1_unHit:Controller.Switch(51)=0:End Sub
Sub Drain_Hit:playsoundAtVol "drain",drain,1:bsTrough.AddBall Me:End Sub          'switch 52
Sub RightSlingShot_Slingshot
  vpmTimer.PulseSwitch (56), 0, ""
  RSFlash.state=1
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RSFlash.state=0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub midswitch_Hit:Controller.Switch(58)=1::End Sub      'switch 58
Sub midswitch_unHit:Controller.Switch(58)=0:End Sub
Sub midswitchB_Hit:Controller.Switch(59)=1::End Sub     'switch 59
Sub midswitchB_unHit:Controller.Switch(59)=0:End Sub
Sub trigger3000_Hit:Controller.Switch(60)=1::End Sub      'switch 60
Sub trigger3000_unHit:Controller.Switch(60)=0:End Sub


Set Lights(1)=alight
Set Lights(2)=blight
Set Lights(3)=clight
Set Lights(4)=doublelight
Set Lights(5)=Light5o
Set Lights(6)=special1
Set Lights(7)=Light7o
Set Lights(8)=Light8o
Set Lights(9)=Light9o
Set Lights(10)=aside
Set Lights(11)=bside
Set Lights(12)=cside
Set Lights(13)=spinlight
Set Lights(14)=Light198
Set Lights(15)=Light199
Set Lights(16)=l1
Set Lights(17)=spell1
Set Lights(18)=spell2
Set Lights(19)=spell3
Set Lights(20)=spell4
Set Lights(21)=spell5
Set Lights(22)=spell6
Set Lights(23)=spell7
Set Lights(24)=spell8
Set Lights(25)=spell9
Set Lights(26)=spell10
Set Lights(27)=spell11
Set Lights(28)=spell12
Set Lights(29)=spell13
Set Lights(30)=spell14
Set Lights(31)=Light31o
Set Lights(32)=Light32o
Set Lights(33)=Light33o
Set Lights(34)=Light34o
Set Lights(35)=Light35o
Set Lights(36)=Light36o
Set Lights(37)=TILT
Set Lights(38)=BallInPlay
Set Lights(39)=Light39o
Set Lights(40)=p1light
Set Lights(41)=p2light
Set Lights(42)=p3light
Set Lights(43)=p4light
Set Lights(44)=GameOver
Set Lights(45)=shootagain
Set Lights(46)=Light46o
Set Lights(47)=special2
Set Lights(48)=Light48o
Set Lights(49)=bonus1
Set Lights(50)=bonus2
Set Lights(51)=bonus3
Set Lights(52)=bonus4
Set Lights(53)=bonus5
Set Lights(54)=bonus6
Set Lights(55)=bonus7
Set Lights(56)=bonus8
Set Lights(57)=bonus9
Set Lights(58)=bonus10
Set Lights(59)=bonus20
Set Lights(60)=Light60o
Set Lights(61)=Light61o
Set Lights(62)=Light62o
Set Lights(63)=special3
Set Lights(64)=Light197
Set Lights(65)=l3

Set Lights(81)=l4
Set Lights(113)=l2
Set Lights(129)=LightS1
Set Lights(130)=LightS2
Set Lights(131)=LightS3
Set Lights(132)=LightS4










'NfadeL 61, l61
'Flash 61,Flash61



Set LampCallback=GetRef("UpdateMultipleLamps")

Dim MacTilt,OldTilt,LeftF,RightF
MacTilt=0:OldTilt=0:LeftF=0:RightF=0

Sub FlashTimer1_Timer
If Light61o.state=1 Then
FlashL1.state=1
FlashL2.state=1
FlashL3.state=1
FlashL4.state=1
Else
FlashL1.state=0
FlashL2.state=0
FlashL3.state=0
FlashL4.state=0
End If

If Light60o.state=1 Then
FlashL5.state=1
FlashL6.state=1
FlashL7.state=1
FlashL8.state=1
FlashL15.state=1
FlashL16.state=1
Else
FlashL5.state=0
FlashL6.state=0
FlashL7.state=0
FlashL8.state=0
FlashL15.state=0
FlashL16.state=0
End If

If Light62o.state=1 Then
FlashL9.state=1
FlashL10.state=1
FlashL11.state=1
FlashL12.state=1
FlashL13.state=1
FlashL14.state=1
Else
FlashL9.state=0
FlashL10.state=0
FlashL11.state=0
FlashL12.state=0
FlashL13.state=0
FlashL14.state=0
End If

If GameOver.state= 1 then FlasherGO.Visible=True Else  FlasherGo.Visible=False
If BallInPlay.state= 1 then FlasherBall.Visible=True Else  FlasherBall.Visible=False
If CreditLight.state= 1 then FlasherCredit.Visible=True Else  FlasherCredit.Visible=False
If Light39o.state= 1 then FlasherMatch.Visible=True Else  FlasherMatch.Visible=False
If TILT.state= 1 then FlasherTilt.Visible=True Else  FlasherTilt.Visible=False
CredLight
End sub


Sub UpdateMultipleLamps

If AltColor=0 then
p1light1.visible=0
p3light1.visible=0
p4light1.visible=0
  Else
p1light.visible=0
p3light.visible=0
p4light.visible=0
End If

If PlayNums=0 Then
FlasherP1.visible=0
FlasherP2.visible=0
FlasherP3.visible=0
FlasherP4.visible=0
End If

If p1light.state=1 then p1light1.state=1 else p1light1.state=0
If p3light.state=1 then p3light1.state=1 else p3light1.state=0
If p4light.state=1 then p4light1.state=1 else p4light1.state=0

  MacTilt=TILT.State
  If P1Light.State=0 And P2Light.State=0 And P3Light.State=0 And P4Light.State=0 Then MacTilt=1
  If GameOver.State=1 Then MacTilt=1
  If MacTilt<>OldTilt Then
    If MacTilt=1 And LeftF=1 Then
      LeftFlipper.RotateToStart
      PlaySoundAtVol"flipperdown", LeftFlipper, VolFlip
      Controller.Switch(19)=0
      Controller.Switch(83)=0
      LeftF=0
    End If
    If MacTilt=1 And RightF=1 Then
      RightFlipper.RotateToStart
      PlaySoundAtVol"flipperdown", RightFlipper, VolFlip
      Controller.Switch(20)=0
      Controller.Switch(81)=0
      RightF=0
    End If
    If MacTilt=1 Then
      RightSlingshot.SlingshotForce=0
      LeftSlingshot.SlingshotForce=0
      topsling.SlingshotForce=0
      bottomsling.SlingshotForce=0


      Wall62.SlingshotForce=0
      Wall61.SlingshotForce=0
    Else
      RightSlingshot.SlingshotForce=6
      LeftSlingshot.SlingshotForce=6
      topsling.SlingshotForce=3
      bottomsling.SlingshotForce=8
      Wall62.SlingshotForce=3
      Wall61.SlingshotForce=3
    End If
  OldTilt=MacTilt
  End If
End Sub

sub FlipperTimer_Timer()
  Primitive16.RotY = gate.currentangle
End Sub

Dim msg(38)
Dim char(9999)
char(0) = " "
char(63) = "0"
char(6) = "1"
char(91) = "2"
char(79) = "3"
char(102) = "4"
char(109) = "5"
char(124) = "6"
char(7) = "7"
char(127) = "8"
char(103) = "9"

'1 on the center line (for Gottlieb) - $CA (202)
'1 on the center with comma - $CB (203)
'6 without top line (Atari, Gottlieb) - $CC (204)
'6 without top line with comma - $CD (205)
'9 without bottom line (Atari, Gottlieb) - $CE (206)
'9 without bottom line with comma - $CF (207)
'char(768) = chr(&HCA) ' 1 on Gottlieb tables
'char(896) = chr(&HCB) ' 1, on Gottlieb tables
char(124) = chr(&HCC) ' 6 without top line
char(252) = chr(&HCD) ' 6, without top line
char(103) = chr(&HCE) ' 9 without bottom line
char(231) = chr(&HCF) ' 9, without bottom line

' Thalamus : This sub is used twice - this means ... this one IS NOT USED

' Sub DisplayTimer_Timer
'   Dim ChgLED,ii,num,chg,stat,obj
' ' ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)
'   ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
'   If Not IsEmpty(ChgLED) Then
'     For ii = 0 To UBound(chgLED)
'       num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
'       if (num < 39) then
' '       MsgBox "Status Code is:" & vbNewLine & stat
'         if char(stat) > "" then msg(num) = char(stat)
'       end if
'     next
'   end if
'   ScoreText1.Text = msg(0) & msg(1) & msg(2) & msg(3) & msg(4) & msg(5)
'   ScoreText2.Text = msg(6) & msg(7) & msg(8) & msg(9) & msg(10) & msg(11)
'   ScoreText3.Text = msg(12) & msg(13) & msg(14) & msg(15) & msg(16)& msg(17)
'   ScoreText4.Text = msg(18) & msg(19) & msg(20) & msg(21) & msg(22) & msg(23)
'   BallText.Text = msg(24) & msg(25)
'   CreditText.Text=msg(26) & msg(27)
' End Sub

'Atari Airborne Avenger
'added by Inkochnito
Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,440,"Airborne Avenger - DIP switches"
    .AddFrame 0,5,190,"Coins per credit",&H000000C3,Array("1 coin 1 credit",0,"1 coin 2 credits",&H00000002,"1 coin 3 credits",&H00000001,"1 coin 4 credits",&H00000003,"2 coins 1 credits",&H00000080)'SW2-3&SW2-4&SW2-5&SW2-6 (dip 2&1&8&7)
    .AddFrame 0,95,190,"Maximum credits",49152,Array("8 credits",0,"12 credits",32768,"15 credits",&H00004000,"20 credits",49152)'SW1-5&SW1-6 (dip 16&15)
    .AddFrame 0,177,190,"Special reward",&H00000030,Array("20,000 points",&H00000030,"30,000 points",&H00000010,"extra ball",0,"replay",&H00000020)'SW2-7&SW2-8 (dip 6&5)
    .AddFrame 0,253,190,"Balls per game",&H00000008,Array("5 balls",0,"3 balls",&H00000008)'SW2-1 (dip 4)
    .AddFrame 210,5,200,"Score threshold level",&H000F0000,Array("50K-70K-90K or 210K-310K-410K",0,"90K-130K-170K or 250K-370K-490K",&H00040000,"130K-190K-250K or 290K-430K-570K",&H00080000,"170K-250K-330K or 330K-490K-650K",&H000C0000,"200K-300K-400K or 360K-540K-720K",&H000F0000)'rotary (dip 17&18&19&20)
    .AddFrame 210,95,200,"High or low threshold scores",&H00000100,Array("low scores",0,"high scores",&H00000100)'SW1-4 (dip 9)
    .AddFrame 210,143,200,"Threshold score reward",&H00003000,Array("nothing",0,"extra ball",&H00002000,"replay",&H00001000)'SW1-7&SW1-8 (dip 14&13)
    .AddFrame 210,205,200,"Spellout reward",&H00000400,Array("extra ball",0,"20,000 points",&H00000400)'SW1-2 (dip 11)
    .AddFrame 210,253,200,"Last ball bonus setting",&H00000200,Array("single bonus",0,"double bonus",&H00000200)'SW1-3 (dip 10)
    .AddChk 210,310,190,Array("Diagnostic mode (must be off)",&H00000800)'SW1-1 (dip 12)
    .AddChk 0,310,190,Array("Match feature",&H00000004)'SW2-2 (dip 3)
    .AddLabel 50,330,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")

 Dim Digits(27)
Digits(0)=Array(Light1,Light2,Light3,Light4,Light5,Light6,Light7)
Digits(1)=Array(Light8,Light9,Light10,Light11,Light12,Light13,Light14)
Digits(2)=Array(Light15,Light16,Light17,Light18,Light19,Light20,Light21)
Digits(3)=Array(Light22,Light23,Light24,Light25,Light26,Light27,Light28)
Digits(4)=Array(Light29,Light30,Light31,Light32,Light33,Light34,Light35)
Digits(5)=Array(Light36,Light37,Light38,Light39,Light40,Light41,Light42)
Digits(6)=Array(Light43,Light44,Light45,Light46,Light47,Light48,Light49)
Digits(7)=Array(Light50,Light51,Light52,Light53,Light54,Light55,Light56)
Digits(8)=Array(Light57,Light58,Light59,Light60,Light61,Light62,Light63)
Digits(9)=Array(Light64,Light65,Light66,Light67,Light68,Light69,Light70)
Digits(10)=Array(Light71,Light72,Light73,Light74,Light75,Light76,Light77)
Digits(11)=Array(Light78,Light79,Light80,Light81,Light82,Light83,Light84)
Digits(12)=Array(Light85,Light86,Light87,Light88,Light89,Light90,Light91)
Digits(13)=Array(Light92,Light93,Light94,Light95,Light96,Light97,Light98)
Digits(14)=Array(Light99,Light100,Light101,Light102,Light103,Light104,Light105)
Digits(15)=Array(Light106,Light107,Light108,Light109,Light110,Light111,Light112)
Digits(16)=Array(Light113,Light114,Light115,Light116,Light117,Light118,Light119)
Digits(17)=Array(Light120,Light121,Light122,Light123,Light124,Light125,Light126)
Digits(18)=Array(Light168,Light167,Light166,Light165,Light164,Light163,Light162)
Digits(19)=Array(Light161,Light160,Light159,Light158,Light157,Light156,Light155)
Digits(20)=Array(Light154,Light153,Light152,Light151,Light150,Light149,Light148)
Digits(21)=Array(Light147,Light146,Light145,Light144,Light143,Light142,Light141)
Digits(22)=Array(Light140,Light139,Light138,Light137,Light136,Light135,Light134)
Digits(23)=Array(Light133,Light132,Light131,Light130,Light129,Light128,Light127)
Digits(24)=Array(Light196,Light195,Light194,Light193,Light192,Light191,Light190)
Digits(25)=Array(Light189,Light188,Light187,Light186,Light185,Light184,Light183)
Digits(26)=Array(Light182,Light181,Light180,Light179,Light178,Light177,Light176) 'CREDITS X10
Digits(27)=Array(Light175,Light174,Light173,Light172,Light171,Light170,Light169) 'CREDITS X1

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

Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 2 ' total number of balls
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

    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub Rubbers_Hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Metals_Hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "metal_hit1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "metal_hit2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "metal_hit3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub Gate_Collide(parm)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "metal_hit1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "metal_hit2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "metal_hit3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*100, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


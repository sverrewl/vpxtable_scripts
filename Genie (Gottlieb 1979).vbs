'************************************************************
'************************************************************
'
'  Genie / IPD No. 997 / October, 1979 / 4 Players
'
'   Credits:
'
'   VPX by bord, rothbauerw
'
'
'************************************************************
'************************************************************

Option Explicit
Randomize

'******************************************************
'             OPTIONS
'******************************************************

Const LibOrCons = 2       '1 for Liberal, 2 for Conservative (liberal also needs pf image set to pf_lib)
Const VolumeDial = 10       'Change volume of hit events
Const RollingSoundFactor = 1    'Change volume of rolling sounds
Const ChimesOn = 0        '1 to use Chimes, 0 to turn them off
Const VROn = 0          '1 to for VR or Cabinet, 0 for desktop

'******************************************************
'           STANDARD DEFINITIONS
'******************************************************

Dim BallMass ,BallSize
Ballsize = 50
Ballmass = 1.0

Const UseSolenoids=2
Const UseLamps=1
Const UseSync=1
Const UseGI=0

' Standard Sounds

Const SSolenoidOn = "SolOn"       'Solenoid activates
Const SSolenoidOff  = "SolOff"      'Solenoid deactivates
Const SKnocker    = "Knocker"
Const SFlipperOn  = "fx_Flipperup"
Const SFlipperOff = "fx_Flipperdown"
Const SCoin     = "coin"
Const cCredits    = ""


'******************************************************
'           TABLE INIT
'******************************************************

On Error Resume Next
  ExecuteGlobal GetTextFile("controller.vbs")
  If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM"01120100","GTS1.VBS",3.02

Const cGameName="genie"

Dim DesktopMode: DesktopMode = Table1.ShowDT

'*********** Desktop/Cabinet settings ************************

Dim GenieBall, xx, dtDrop1, dtDrop2

Sub Table1_Init
  vpmInit Me

  With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine="Genie (Gottlieb 1979)"
    .HandleKeyboard=0
    .ShowTitle=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .Hidden = 0
    On Error Resume Next
    .SolMask(0) = 0
    vpmTimer.AddTimer 1000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 1 seconds
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With

  ' Nudging
  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = 1

  vpmNudge.TiltSwitch = 4
  vpmNudge.Sensitivity = 3
  vpmNudge.TiltObj = Array(LeftSlingshot, Bumper1, Bumper2, Bumper3)

  Set GenieBall = Drain.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Controller.Switch(66) = 1

  Set dtDrop1=New cvpmDropTarget
  dtDrop1.InitDrop Array(sw20,sw21,sw23,sw24),Array(20,21,23,24)
  'dtDrop1.InitSnd"DTDrop","DTReset"

  Set dtDrop2=New cvpmDropTarget
  dtDrop2.InitDrop Array(sw30,sw70,sw31,sw71,sw60,sw74,sw61),Array(30,70,31,71,60,74,61)
  'dtDrop2.InitSnd"DTDrop","DTReset"

  If LibOrCons = 1 Then
    For each xx in physics_lib
      xx.collidable = True
    Next
    For each xx in visible_lib
      xx.visible = True
    Next

    For each xx in physics_cons
      xx.collidable = false
    Next
    For each xx in visible_cons
      xx.visible = False
    Next
  Else
    For each xx in physics_lib
      xx.collidable = False
    Next
    For each xx in visible_lib
      xx.visible = False
    Next

    For each xx in physics_cons
      xx.collidable = True
    Next
    For each xx in visible_cons
      xx.visible = True
    Next
  End If

  setup_backglass()

End Sub


Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'******************************************************
'             KEYS
'******************************************************

 Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftFlipperKey Then LFPress = 1
  If keycode = RightFlipperKey Then RFPress = 1

  If vpmKeyDown(KeyCode) Then Exit Sub
  If keycode = AddCreditKey Then
    vpmTimer.AddTimer 750, "vpmTimer.PulseSw swCoin1'":Playsoundat SCoin, Drain
  End If
  If keycode=PlungerKey Then Plunger.Pullback:playsoundat "plungerpull", Plunger
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then
    lfpress = 0
    leftflipper.eostorqueangle = EOSA
    leftflipper.eostorque = EOST
    leftflipper1.eostorqueangle = EOSA
    leftflipper1.eostorque = EOST
  End If
  If keycode = RightFlipperKey Then
    rfpress = 0
    rightflipper.eostorqueangle = EOSA
    rightflipper.eostorque = EOST
    rightflipper1.eostorqueangle = EOSA
    rightflipper1.eostorque = EOST
    rightflipper2.eostorqueangle = EOSA
    rightflipper2.eostorque = EOST
  End If

  If vpmKeyUp(KeyCode) Then Exit Sub
  If keycode=PlungerKey Then Plunger.Fire:playsoundat "plunger", Plunger
End Sub

'******************************************************
'             SOLENOIDS
'******************************************************

SolCallback(1)  = "SolRelease"
SolCallback(2)  = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(3)  = "SolChime1"
SolCallback(4)  = "SolChime2"
SolCallback(5)  = "SolChime3"
SolCallback(6)  = "SollSaucer"
SolCallback(7)  = "SolRightDrop"
SolCallback(8)  = "SolLeftDrop"
SolCallback(17) = "vpmNudge.SolGameOn"

'******************************************************
'             Chimes (Optional)
'******************************************************

Sub solchime1(Enable)
    If Enable then
    If ChimesOn=1 Then
      PlaySound SoundFX("sj_chime_10a",DOFChimes), 1, 1
      DOF 125, 2
    End If
  End If
End Sub

Sub solchime2(Enable)
    If Enable then
    If ChimesOn=1 Then
      PlaySound SoundFX("sj_chime_100a",DOFChimes), 1, 1
      DOF 126, 2
    End If
  End If
End Sub

Sub solchime3(Enable)
    If Enable then
    If ChimesOn=1 Then
      PlaySound SoundFX("sj_chime_1000a",DOFChimes), 1, 1
      DOF 127, 2
    End If
  End If
End Sub


'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub Drain_Hit()
  Controller.Switch(66) = 1
  PlaySoundAT "drain", Drain
End Sub

Sub Drain_UnHit()  'Drain
  Controller.Switch(66) = 0
End Sub

Sub SolRelease(enabled)
  If enabled Then
    If Drain.BallCntOver = 0 Then
      PlaySoundAt SoundFX("solon",DOFContactors), Drain
    Else
      PlaySoundAt SoundFX("ballrelease",DOFContactors), Drain
    End If
    Drain.kick 60, 20
  End If
End Sub

'******************************************************
'           KICKER
'******************************************************

'*** PI returns the value for PI
Function PI()
  PI = 4*Atn(1)
End Function

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

Sub TopHole_hit
  PlaySoundAtBall "fx_kicker-enter"
  set KickerBall = activeball
  Controller.switch(41)=1
End Sub

Sub TopHole_unHit
  Controller.Switch(41) = 0
End Sub

Dim KickerBall, kickstep1

Sub SolLSaucer(Enable)
    If Enable then
    PlaySoundAt "fx2_kicker_enter", tophole
'   bsSaucer.ExitSol_On
    kickstep1 = 0
    pkickarm.rotz=12
    If tophole.ballcntover > 0 then
      KickBall KickerBall, 208, 9, 5, 30
    End If
    tophole.timerenabled = 1
    End If
End Sub

Sub tophole_timer
    Select Case kickstep1
        Case 3:pkickarm.rotz=12
        Case 4:pkickarm.rotz=12
        Case 5:pkickarm.rotz=12
        Case 6:pkickarm.rotz=12
        Case 7:pkickarm.rotz=4
        Case 8:pkickarm.rotz=2
    Case 9:pkickarm.rotz=0:tophole.TimerEnabled = 0: if tophole.ballcntover > 0 then SolLSaucer -1
    End Select
   kickstep1 = kickstep1 + 1
End Sub

'******************************************************
'         RIGHT DROP TARGETS
'******************************************************

dim DTInterval, DTStep1, DTStep2, DTStep3, DTStep4
DTInterval = 17

DTStep1 = -12
DTStep2 = -24
DTStep3 = -36
DTStep4 = -43

sw20.timerinterval = DTInterval
sw21.timerinterval = DTInterval
sw23.timerinterval = DTInterval
sw24.timerinterval = DTInterval

sw30.timerinterval = DTInterval
sw31.timerinterval = DTInterval
sw60.timerinterval = DTInterval
sw61.timerinterval = DTInterval
sw70.timerinterval = DTInterval
sw71.timerinterval = DTInterval
sw74.timerinterval = DTInterval

Sub HitDT (sw, prim, step)
  playsoundat SoundFX("DTDROP",DOFContactors),sw
  prim.transz= DTStep1
  step = 0
  sw.timerenabled=1
  sw.collidable=0
End Sub

Sub AnimDT (sw, prim, step, flasher)
  Select Case step
    Case 1:prim.transz= DTStep2
    Case 2:prim.transz= DTStep3:flasher.visible=0
    Case 3:prim.transz= DTStep4:sw.TimerEnabled = 0
  End Select
  Step = Step + 1
End Sub

Sub RaiseDT (Prims, Flashers, step)
  dim x, y, z
  If Step = 1 Then
    x = DTStep3
  elseif Step = 2 Then
    x = DTStep2
    For Each z In Flashers
      z.visible = True
    Next
  elseif Step = 3 Then
    x = DTStep1
  Else
    x = 0
  End If

  For Each y In Prims
    If y.transz < x then y.transz = x
  Next
End Sub

dim d20step, d21step, d23step, d24step, d30step, d31step, d60step, d61step, d70step, d71step, d74step
dim DTRPrims, DTLPrims, DTRFlashers, DTLFlashers

DTRPrims = Array(layer3sw20, layer3sw21, layer3sw23, layer3sw24)
DTRFlashers = Array(Flasher20, Flasher21, Flasher23, Flasher24)
DTLPrims = Array(layer3sw30, layer3sw31, layer3sw60, layer3sw61, layer3sw70, layer3sw71, layer3sw74)
DTLFlashers = Array(Flasher30, Flasher31, Flasher60, Flasher61, Flasher70, Flasher71, Flasher74)

Sub SolRightDrop(Enabled)
  If Enabled Then
    If GenieBall.x > 640 and GenieBall.x < 860 and GenieBall.y < 835 and GenieBall.y > 765 Then
      vpmTimer.Addtimer 200, "SolRightDrop -1'"
    Else
      dtDrop1.SolDropUp Enabled
      playsoundat SoundFX("DTReset",DOFContactors),sw21
      sw20.collidable=1:sw21.collidable=1:sw23.collidable=1:sw24.collidable=1
      RaiseDT DTRPrims, DTRFlashers,  1
      vpmTimer.Addtimer DTInterval, "RaiseDT DTRPrims, DTRFlashers, 2'"
      vpmTimer.Addtimer DTInterval*2, "RaiseDT DTRPrims, DTRFlashers, 3'"
      vpmTimer.Addtimer DTInterval*3, "RaiseDT DTRPrims, DTRFlashers, 4'"
    End If
  End If
End Sub

Sub sw20_Hit
  dtdrop1.Hit 1
  HitDT sw20, Layer3sw20, d20step
End Sub

Sub sw20_timer
  AnimDT sw20, Layer3sw20, d20step, flasher20
End Sub

Sub sw21_Hit
  dtdrop1.Hit 2
  HitDT sw21, Layer3sw21, d21step
End Sub

Sub sw21_timer
  AnimDT sw21, Layer3sw21, d21step, flasher21
End Sub

Sub sw23_Hit
  dtdrop1.Hit 3
  HitDT sw23, Layer3sw23, d23step
End Sub

Sub sw23_timer
  AnimDT sw23, Layer3sw23, d23step, flasher23
End Sub

Sub sw24_Hit
  dtdrop1.Hit 4
  HitDT sw24, Layer3sw24, d24step
End Sub

Sub sw24_timer
  AnimDT sw24, Layer3sw24, d24step, flasher24
End Sub


'******************************************************
'         LEFT DROP TARGETS
'******************************************************

Sub SolLeftDrop(Enabled)
  If Enabled Then
    If GenieBall.x > 115 and GenieBall.x < 505 and GenieBall.y < 190 Then
      vpmTimer.Addtimer 200, "SolLeftDrop -1'"
    Else
      dtDrop2.SolDropUp Enabled
      playsoundat SoundFX("DTReset",DOFContactors),sw60
      sw30.collidable=1:sw31.collidable=1:sw60.collidable=1:sw61.collidable=1:sw70.collidable=1:sw71.collidable=1:sw74.collidable=1
      RaiseDT DTLPrims,  DTLFlashers, 1
      vpmTimer.Addtimer DTInterval, "RaiseDT DTLPrims, DTLFlashers,  2'"
      vpmTimer.Addtimer DTInterval*2, "RaiseDT DTLPrims, DTLFlashers,  3'"
      vpmTimer.Addtimer DTInterval*3, "RaiseDT DTLPrims, DTLFlashers,  4'"
    End If
  End If
End Sub

Sub sw30_Hit:
  dtdrop2.Hit 1
  HitDT sw30, Layer3sw30, d30step
End Sub

Sub sw30_timer
  AnimDT sw30, Layer3sw30, d30step, flasher30
End Sub

Sub sw31_Hit
  dtdrop2.Hit 3
  HitDT sw31, Layer3sw31, d31step
End Sub

Sub sw31_timer
  AnimDT sw31, Layer3sw31, d31step, flasher31
End Sub

Sub sw60_Hit
  dtdrop2.Hit 5
  HitDT sw60, Layer3sw60, d60step
End Sub

Sub sw60_timer
  AnimDT sw60, Layer3sw60, d60step, flasher60
End Sub

Sub sw61_Hit
  dtdrop2.Hit 7
  HitDT sw61, Layer3sw61, d61step
End Sub

Sub sw61_timer
  AnimDT sw61, Layer3sw61, d61step, flasher61
End Sub

Sub sw70_Hit
  dtdrop2.Hit 2
  HitDT sw70, Layer3sw70, d70step
End Sub

Sub sw70_timer
  AnimDT sw70, Layer3sw70, d70step, flasher70
End Sub

Sub sw71_Hit
  dtdrop2.Hit 4
  HitDT sw71, Layer3sw71, d71step
End Sub

Sub sw71_timer
  AnimDT sw71, Layer3sw71, d71step, flasher71
End Sub

Sub sw74_Hit
  dtdrop2.Hit 6
  HitDT sw74, Layer3sw74, d74step
End Sub

Sub sw74_timer
  AnimDT sw74, Layer3sw74, d74step, flasher74
End Sub

'******************************************************
'         FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("fx_Flipperup",DOFContactors), LeftFlipper
    LeftFlipper1.RotateToEnd
    lf.fire
  Else
    if leftflipper.currentangle < leftflipper.startangle - 5 then
      PlaySoundAt SoundFX("fx_flipperdown",DOFContactors), LeftFlipper
    end if
    LeftFlipper1.RotateToStart
    LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("fx_Flipperup",DOFContactors), RightFlipper
    RightFlipper2.RotateToEnd
    RF.fire
    RF1.fire
  Else
    if RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      PlaySoundAt SoundFX("fx_Flipperdown",DOFContactors), RightFlipper
    End If
    RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
    RightFlipper2.RotateToStart
  End If
End Sub

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, "Left"
  FlipperTricks LeftFlipper1, LFPress, LFCount1, LFEndAngle1, "Left"
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, "Right"
  FlipperTricks RightFlipper1, RFPress, RFCount1, RFEndAngle1, "Right"
  FlipperTricks RightFlipper2, RFPress, RFCount2, RFEndAngle2, "Right"
end sub

dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSAnew
dim Frampup, FElasticity, EOSRampup, SOSRampup
dim RFEndAngle, RFEndAngle1, RFEndAngle2, LFEndAngle, LFEndAngle1, LFCount, LFCount1, RFCount, RFCount1, RFCount2, LiveCatch

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
EOSTnew = 1.5 'FEOST
EOSAnew = 0.2
EOSRampup = 1.5
SOSRampup = 8.5
LiveCatch = 8

LFEndAngle = Leftflipper.endangle
LFEndAngle1 = Leftflipper1.endangle
RFEndAngle = RightFlipper.endangle
RFEndAngle1 = RightFlipper1.endangle
RFEndAngle2 = RightFlipper2.endangle

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, LoR)
  If Flipper.currentangle = Flipper.endangle and FlipperPress = 1 then
    Flipper.eostorqueangle = EOSAnew
    Flipper.eostorque = EOSTnew
    Flipper.rampup = EOSRampup
    if FCount = 0 Then FCount = GameTime
    if GameTime - FCount < LiveCatch Then
      Flipper.Elasticity = 0.1
      If Flipper.endangle <> FEndAngle Then Flipper.endangle = FEndAngle
    Else
      Flipper.Elasticity = FElasticity
    end if
  elseif (LoR = "Right" and Flipper.currentangle < Flipper.startangle + 0.05) or (LoR = "Left" and Flipper.currentangle > Flipper.startangle - 0.05) Then
    Flipper.rampup = SOSRampup
    Flipper.endangle = FEndAngle + 3
    Flipper.Elasticity = FElasticity
    FCount = 0
  elseif (LoR = "Right" and Flipper.currentangle < Flipper.endangle - 0.01) or  (LoR = "Left" and Flipper.currentangle > Flipper.endangle + 0.01) Then
    Flipper.eostorque = EOST
    Flipper.eostorqueangle = EOSA
    Flipper.rampup = Frampup
    Flipper.Elasticity = FElasticity
  end if
End Sub




Sub LeftFlipper_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub LeftFlipper1_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub RightFlipper_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub RightFlipper1_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub RightFlipper2_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub


'******************************************************
'         TRIGGERS & TARGETS
'******************************************************

'***********************************************************
'****     Rollover Wire Trigger CODE        ****
'***********************************************************
Sub SW10A_Hit:Controller.Switch(10)=1 : playsoundat"rollover", SW10A :DOF 103, DOFOn: End Sub
Sub SW10A_UnHit:Controller.Switch(10)=0:DOF 103, DOFOff:End Sub
Sub SW11A_Hit:Controller.Switch(11)=1 :  playsoundat"rollover", SW11A :DOF 105, DOFOn : End Sub
Sub SW11A_UnHit:Controller.Switch(11)=0:DOF 105, DOFOff :End Sub
Sub SW13A_Hit:Controller.Switch(13)=1 :  playsoundat"rollover", SW13A :DOF 107, DOFOn : End Sub
Sub SW13A_UnHit:Controller.Switch(13)=0:DOF 107, DOFOff :End Sub
Sub SW14A_Hit:Controller.Switch(14)=1 :  playsoundat"rollover", SW14A:DOF 109, DOFOn : End Sub
Sub SW14A_UnHit:Controller.Switch(14)=0:DOF 109, DOFOff :End Sub

Sub SW10B_Hit:Controller.Switch(10)=1 :  playsoundat"rollover", SW10B: DOF 104, DOFOn :End Sub
Sub SW10B_UnHit:Controller.Switch(10)=0:DOF 104, DOFOff :End Sub
Sub SW11B_Hit:Controller.Switch(11)=1 : playsoundat"rollover", SW11B : DOF 106, DOFOn : End Sub
Sub SW11B_UnHit:Controller.Switch(11)=0:DOF 106, DOFOff :End Sub
Sub SW13B_Hit:Controller.Switch(13)=1 : playsoundat"rollover", SW13B : DOF 108, DOFOn :End Sub
Sub SW13B_UnHit:Controller.Switch(13)=0:DOF 108, DOFOff :End Sub
Sub SW14B_Hit:Controller.Switch(14)=1 : playsoundat"rollover", SW14B:  DOF 110, DOFOn :End Sub
Sub SW14B_UnHit:Controller.Switch(14)=0:DOF 110, DOFOff :End Sub

Sub SW54A_Hit:Controller.Switch(54)=1 : playsoundat"rollover", SW54A :DOF 115, DOFOn : End Sub
Sub SW54A_UnHit:Controller.Switch(54)=0:DOF 115, DOFOff :End Sub
Sub SW54B_Hit:Controller.Switch(54)=1 : playsoundat"rollover", SW54B :DOF 116, DOFOn : End Sub
Sub SW54B_UnHit:Controller.Switch(54)=0:DOF 116, DOFOff :End Sub

Sub SW63A_Hit:Controller.Switch(63)=1 : playsoundat"rollover", SW63A :DOF 112, DOFOn : End Sub
Sub SW63A_UnHit:Controller.Switch(63)=0:DOF 112, DOFOff :End Sub
Sub SW63B_Hit:Controller.Switch(63)=1 : playsoundat"rollover", SW63B :DOF 113, DOFOn : End Sub
Sub SW63B_UnHit:Controller.Switch(63)=0:DOF 113, DOFOff :End Sub
Sub SW63C_Hit:Controller.Switch(63)=1 : playsoundat"rollover", SW63C :DOF 114, DOFOn : End Sub
Sub SW63C_UnHit:Controller.Switch(63)=0:DOF 114, DOFOff :End Sub

Sub SW73A_Hit:Controller.Switch(73)=1 : playsoundat"rollover", SW73A:DOF 117, DOFOn : End Sub
Sub SW73A_UnHit:Controller.Switch(73)=0:DOF 117, DOFOff :End Sub
Sub SW73B_Hit:Controller.Switch(73)=1 : playsoundat"rollover", SW73B :DOF 118, DOFOn : End Sub
Sub SW73B_UnHit:Controller.Switch(73)=0:DOF 118, DOFOff :End Sub

'***********************************************************
'****       Star SWITCH CODE        ****
'***********************************************************
Sub SW53A_Hit:Controller.Switch(53)=1 : playsoundat"rollover", SW53A : DOF 122, DOFOn :star53a.transz=-4 : End Sub
Sub SW53A_UnHit:Controller.Switch(53)=0:star53a.transz=0 :DOF 122, DOFOff : End Sub
Sub SW53B_Hit:Controller.Switch(53)=1 : playsoundat"rollover", SW53B : DOF 123, DOFOn :star53b.transz=-4 : End Sub
Sub SW53B_UnHit:Controller.Switch(53)=0:star53b.transz=0 :DOF 123, DOFOff : End Sub
Sub SW53C_Hit:Controller.Switch(53)=1 : playsoundat"rollover", SW53C :DOF 124, DOFOn : star53c.transz=-4 : End Sub
Sub SW53C_UnHit:Controller.Switch(53)=0:star53c.transz=0 :DOF 124, DOFOff : End Sub
Sub SW64A_Hit:Controller.Switch(64)=1 : playsoundat"rollover", SW64A : DOF 119, DOFOn :star64a.transz=-4 : End Sub
Sub SW64A_UnHit:Controller.Switch(64)=0:star64a.transz=0 :DOF 119, DOFOff : End Sub
Sub SW64B_Hit:Controller.Switch(64)=1 : playsoundat"rollover", SW64B : DOF 120, DOFOn :star64b.transz=-4 : End Sub
Sub SW64B_UnHit:Controller.Switch(64)=0:star64b.transz=0 :DOF 120, DOFOff : End Sub

'***********************************************************
'****         SPINNER CODE          ****
'***********************************************************
Sub SW33_Spin:vpmTimer.PulseSw 33 : playsoundat"fx_spinner", sw33 : End Sub

'***********Rotate Spinner
'Dim Angle

Sub SpinnerTimer_Timer
  SpinnerP.Rotx = sw33.CurrentAngle
  SpinnerRod.TransX = sin( (sw33.CurrentAngle+180) * (2*PI/360)) * 3.5
  SpinnerRod.TransZ = sin( (sw33.CurrentAngle- 90) * (2*PI/360)) * 3.5
End Sub

'***********************************************************
'****       STAND UP TARGET CODE        ****
'***********************************************************
dim step40, step44, step51, step63

Sub sw40_Hit
  vpmTimer.PulseSw (40)
  Layer3sw40.transy=-3
  Timer40.enabled = 1
  step40 = 0
end sub

Sub Timer40_timer
  Select Case step40
    Case 3: Layer3sw40.transy=3
    Case 4: Layer3sw40.transy=-3
    Case 5: Layer3sw40.transy=2
    Case 6: Layer3sw40.transy=-2
    Case 7: Layer3sw40.transy=1
    Case 8: Layer3sw40.transy=-1
    Case 9: Layer3sw40.transy=0
    Case 10: Layer3sw40.transy=-1: Timer40.enabled = 0
  End Select
  step40=step40 + 1
End Sub

Sub sw44_Hit
  vpmTimer.PulseSw (44)
  Layer3sw44.transy=-3
  Timer44.enabled = 1
  step44 = 0
end sub

Sub Timer44_timer
  Select Case step44
    Case 3: Layer3sw44.transy=3
    Case 4: Layer3sw44.transy=-3
    Case 5: Layer3sw44.transy=2
    Case 6: Layer3sw44.transy=-2
    Case 7: Layer3sw44.transy=1
    Case 8: Layer3sw44.transy=-1
    Case 9: Layer3sw44.transy=0
    Case 10: Layer3sw44.transy=-1: Timer44.enabled = 0
  End Select
  step44=step44 + 1
End Sub

Sub sw51_Hit
  vpmTimer.PulseSw (51)
  Layer3sw51.transy=-3
  Timer51.enabled = 1
  step51 = 0
end sub

Sub Timer51_timer
  Select Case step51
    Case 3: Layer3sw51.transy=3
    Case 4: Layer3sw51.transy=-3
    Case 5: Layer3sw51.transy=2
    Case 6: Layer3sw51.transy=-2
    Case 7: Layer3sw51.transy=1
    Case 8: Layer3sw51.transy=-1
    Case 9: Layer3sw51.transy=0
    Case 10: Layer3sw51.transy=0: Timer51.enabled = 0
  End Select
  step51=step51 + 1
End Sub

Sub sw63_Hit
  DOF 111, DOFPulse
  vpmTimer.PulseSw (63)
  Layer3sw63.transy=-3
  Timer63.enabled = 1
  step63 = 0
end sub

Sub Timer63_timer
  Select Case step63
    Case 3: Layer3sw63.transy=3
    Case 4: Layer3sw63.transy=-3
    Case 5: Layer3sw63.transy=2
    Case 6: Layer3sw63.transy=-2
    Case 7: Layer3sw63.transy=1
    Case 8: Layer3sw63.transy=-1
    Case 9: Layer3sw63.transy=0
    Case 10: Layer3sw63.transy=-1: Timer63.enabled = 0
  End Select
  step63=step63 + 1
End Sub

'***********************************************************
'****         SLING CODE            ****
'***********************************************************

Dim LStep

Sub LeftSlingShot_Slingshot
  PlaySoundAt SoundFx("Slingshot",DOFContactors), SlingL
  DOF 121, DOFPulse
  LSling.Visible = 0
  LSling001.Visible = 1
  SlingL.Rotx = 22
  LStep = 0
  vpmTimer.PulseSw 50
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 1: LSLing001.Visible = 0
        LSLing002.Visible = 1
        Slingl.Rotx = 9
    Case 2: LSLing002.Visible = 0
        LSLing003.Visible = 1
        SlingL.Rotx = 0
    Case 3: LSLing003.Visible = 0
        LSLing.Visible = 1
        SlingL.Rotx = 0
        LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub

Sub sw50c_Hit(idx):vpmTimer.PulseSw 50 : End Sub

'***********************************************************
'****         BUMPER CODE           ****
'***********************************************************
dim bump1step, bump2step, bump3step

Sub bumper1_Hit
  vpmTimer.PulseSw 43
  playsoundAt SoundFX("fx_bumper1",DOFContactors), bumper1
  layer1bumpskirt1.RotY=-3
  Bump1Step = 0
  bumper1.timerenabled = 1
End Sub

Sub bumper1_timer
  Select Case Bump1Step
    Case 3: layer1bumpskirt1.RotY=-1
    Case 4: layer1bumpskirt1.RotY=1
    Case 5: layer1bumpskirt1.RotY=-1
    Case 6: layer1bumpskirt1.RotY=1
    Case 7: layer1bumpskirt1.RotY=0:bumper1.timerenabled = 0
  End Select
  Bump1Step=Bump1Step + 1
End Sub

Sub bumper2_Hit
  vpmTimer.PulseSw 53
  DOF 101, DOFPulse
  playsoundAt SoundFX("fx_bumper2",DOFContactors), bumper1
  layer1bumpskirt2.RotY=-3
  Bump2Step = 0
  bumper2.timerenabled = 1
End Sub

Sub bumper2_timer
  Select Case Bump2Step
    Case 3: layer1bumpskirt2.RotY=-1
    Case 4: layer1bumpskirt2.RotY=1
    Case 5: layer1bumpskirt2.RotY=-1
    Case 6: layer1bumpskirt2.RotY=1
    Case 7: layer1bumpskirt2.RotY=0:bumper2.timerenabled = 0
  End Select
  Bump2Step=Bump2Step + 1
End Sub

Sub bumper3_Hit
  vpmTimer.PulseSw 53
  DOF 102, DOFPulse
  playsoundAt SoundFX("fx_bumper3",DOFContactors), bumper1
  layer1bumpskirt3.RotY=-3
  Bump3Step = 0
  bumper3.timerenabled = 1
End Sub

Sub bumper3_timer
  Select Case Bump3Step
    Case 3: layer1bumpskirt3.RotY=-1
    Case 4: layer1bumpskirt3.RotY=1
    Case 5: layer1bumpskirt3.RotY=-1
    Case 6: layer1bumpskirt3.RotY=1
    Case 7: layer1bumpskirt3.RotY=0:bumper3.timerenabled = 0
  End Select
  Bump3Step=Bump3Step + 1
End Sub

'******************************************************************************
'Animated Rubbers Code

dim cstep, t10step, t8step

Sub RubberBand_17_Hit
  crubber.Visible = 0
  crubber001.Visible = 1
  cStep = 0
  crubbertime.Enabled = 1
End Sub

Sub crubbertime_Timer
  Select Case cStep
    Case 1: crubber001.Visible = 0
        crubber002.Visible = 1
    Case 2: crubber002.Visible = 0
        crubber003.Visible = 1
    Case 3: crubber003.Visible = 0
        crubber.Visible = 1
        crubbertime.Enabled = 0
  End Select
  cStep = cStep + 1
End Sub

Sub RubberBand_8_Hit
  trubber.Visible = 0
  trubber8_001.Visible = 1
  t8Step = 0
  trubber8time.Enabled = 1
End Sub

Sub trubber8time_Timer
  Select Case t8Step
    Case 1: trubber8_001.Visible = 0
        trubber8_002.Visible = 1
    Case 2: trubber8_002.Visible = 0
        trubber8_003.Visible = 1
    Case 3: trubber8_003.Visible = 0
        trubber.Visible = 1
        trubber8time.Enabled = 0
  End Select
  t8Step = t8Step + 1
End Sub

Sub RubberBand_10_Hit
  trubber.Visible = 0
  trubber10_001.Visible = 1
  t10Step = 0
  trubber10time.Enabled = 1
End Sub

Sub trubber10time_Timer
  Select Case t10Step
    Case 1: trubber10_001.Visible = 0
        trubber10_002.Visible = 1
    Case 2: trubber10_002.Visible = 0
        trubber10_003.Visible = 1
    Case 3: trubber10_003.Visible = 0
        trubber.Visible = 1
        trubber10time.Enabled = 0
  End Select
  t10Step = t10Step + 1
End Sub

'******************************************************************************
'Primitive Flipper Code
Sub FlipperTimer_Timer
  LFlipb.rotz = LeftFlipper.currentangle
  LFlipr.rotz = LeftFlipper.currentangle
  LFlipb1.rotz = LeftFlipper1.currentangle
  LFlipr1.rotz = LeftFlipper1.currentangle
  RFlipb.rotz = RightFlipper.currentangle
  RFlipr.rotz = RightFlipper.currentangle
  RFlipb1.rotz = RightFlipper1.currentangle
  RFlipr1.rotz = RightFlipper1.currentangle
  RFlipb2.rotz = RightFlipper2.currentangle
  RFlipr2.rotz = RightFlipper2.currentangle

  batleftshadow.rotz = LeftFlipper.CurrentAngle
  batleftshadow1.rotz = LeftFlipper1.CurrentAngle
  batrightshadow.rotz  = RightFlipper.CurrentAngle
  batrightshadow1.rotz  = RightFlipper1.CurrentAngle
  batrightshadow2.rotz  = RightFlipper2.CurrentAngle

  Pgate.rotx=-Gate.currentangle*.5
  plungegate_prim.RotX = Gate1.CurrentAngle + 70
End Sub
'******************************************************************************

'**********************************************************************************************************

'Map lights to an array
'**********************************************************************************************************

  Lights(4) = Array(l4p,l4pa) ' Shoot Again Playfied and backglass
Set Lights(5) = l5
  Lights(6) = Array(l6,L6a)
  Lights(7) = Array(l7,L7a)
  Lights(8) = Array(l8,L8a)
  Lights(9) = Array(l9,L9a)
  Lights(10) = Array(l10,L10a)
  Lights(11) = Array(l11,L11a)
Set Lights(12) = l12
  Lights(13) = Array(l13,L13a)
  Lights(14) = Array(l14,L14a)
  Lights(15) = Array(l15,L15a)
  Lights(16) = Array(l16,L16a)
  Lights(17) = Array(l17,L17a)
Set Lights(18) = l18
  Lights(19) = Array(l19,L19a)
  Lights(20) = Array(l20,L20a)
  Lights(21) = Array(l21,L21a)
  Lights(22) = Array(l22,L22a)
Set Lights(23) = l23
Set Lights(24) = l24
Set Lights(25) = l25
Set Lights(26) = l26
Set Lights(27) = l27
Set Lights(28) = l28
Set Lights(29) = l29
Set Lights(30) = l30
Set Lights(31) = l31
  Lights(32) = Array(l32,l32a)
  Lights(33) = Array(l33,l33a)
  Lights(36) = Array(l36,L36a)
Set Lights(150) = Light1 'Star Trigger 1
Set Lights(151) = Light2 'Star Trigger 2
Set Lights(152) = Light3 'Star Trigger 3
Set Lights(153) = Light4 'Star Trigger 4
Set Lights(154) = Light5 'Star Trigger 5


'Backglass
Set Lights(1) = l1 'High Score
Set Lights(2) = l2 'Tilt
Set Lights(3) = l3 'High Score
'**********************************************************************************************************
' Backglass Light Displays (7 digit 7 segment displays)
Dim Digits(32)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)
Digits(6)=Array(b00,b01,b02,b03,b04,b05,b06,n,b08)
Digits(7)=Array(b10,b11,b12,b13,b14,b15,b16,n,b18)
Digits(8)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(9)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(10)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(11)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)
Digits(12)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(13)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)
Digits(14)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(15)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(16)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(17)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)
Digits(18)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(19)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)
Digits(20)=Array(d20,d21,d22,d23,d24,d25,d26,n,d28)
Digits(21)=Array(d30,d31,d32,d33,d34,d35,d36,n,d38)
Digits(22)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(23)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)
Digits(24)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(25)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(26)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(27)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)


'**********************************************************************************************************
'Hanz's FSS EM Reel & Drum manager/interpreter script start
'**********************************************************************************************************

'Digital LED Display

Dim LEDFSS(28)
LEDFSS(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6, n101, LED1x8)
LEDFSS(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6, n101, LED2x8)
LEDFSS(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6, n101, LED3x8)
LEDFSS(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6, n101, LED4x8)
LEDFSS(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6, n101, LED5x8)
LEDFSS(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6, n101, LED6x8)
'LEDFSS(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6, n101, LED7x8)

LEDFSS(6) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6, n101, LED8x8)
LEDFSS(7) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6, n101, LED9x8)
LEDFSS(8) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6, n101, LED10x8)
LEDFSS(9) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6, n101, LED11x8)
LEDFSS(10) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6, n101, LED12x8)
LEDFSS(11) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6, n101, LED13x8)
'LEDFSS(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6, n101, LED14x8)

LEDFSS(12) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006, n101, LED1x008)
LEDFSS(13) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106, n101, LED1x108)
LEDFSS(14) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206, n101, LED1x208)
LEDFSS(15) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306, n101, LED1x308)
LEDFSS(16) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406, n101, LED1x408)
LEDFSS(17) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506, n101, LED1x508)
'LEDFSS(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606, n101, LED1x608)

LEDFSS(18) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006, n101, LED2x008)
LEDFSS(19) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106, n101, LED2x108)
LEDFSS(20) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206, n101, LED2x208)
LEDFSS(21) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306, n101, LED2x308)
LEDFSS(22) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406, n101, LED2x408)
LEDFSS(23) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506, n101, LED2x508)
'LEDFSS(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606, n101, LED2x608)

LEDFSS(26) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306, n101, LEDax308)
LEDFSS(27) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406, n101, LEDbx408)

LEDFSS(24) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506, n101, LEDcx508)
LEDFSS(25) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606, n101, LEDdx608)



Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    If VROn = 1 or Table1.ShowFSS = True or DesktopMode = False Then

      For ii = 0 To UBound(chgLED)
        num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
        if (num < 28) then
          For Each obj In LEDFSS(num)
            If chg And 1 Then obj.visible = stat And 1
            chg = chg\2 : stat = stat\2
          Next
        else
        end if
      next

    elseif DesktopMode = True then ' DT

      For ii = 0 To UBound(chgLED)
        num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
        if (num < 28) then
          For Each obj In Digits(num)
            If chg And 1 Then obj.State = stat And 1
            chg = chg\2 : stat = stat\2
          Next
        else
        end if
      next

    end if
  end if ' end if ChgLED

  Flasher1.visible = l2.state
  Flasher2.visible = l1.state
  If l1.state Then Flasher3.visible = false else Flasher3.visible = true
  If l1.state Then Flasher4.visible = false else Flasher4.visible = true
  Flasher5.visible = l3.state
  Flasher6.visible = l4pa.state
End Sub

'if Full Screen turn off Backglas lights
If DesktopMode = True Or VROn = 1 Then
dim xxx
For each xxx in BL:xxx.Visible = 1: Next
lockdownbar.visible=1
layer3rails.visible=1
else
For each xxx in BL:xxx.Visible = 0: Next
lockdownbar.visible=0
layer3rails.visible=0
End if

Dim BootCount:BootCount = 0

Sub BootTable_Timer()
  If BootCount = 0 Then
    BootCount = 1
    PlaySoundAt "poweron", Plunger
  Else
    '*****GI Lights On
    dim xx
    For each xx in GI:xx.State = 1: Next
    for each xx in dtShadows: xx.visible=1: Next
    For each xx in layer1GI: xx.image = "layer1": Next
    For each xx in layer2GI: xx.image = "layer2": Next
    For each xx in layer3GI: xx.image = "layer3": Next
    bumpercap12.image="bumpercaps12"
    bumpercap3.image="bumpercap3"
    plasticlow.image="lowplastic"
    bumpbase.image="bumpbaseon"
    playfield_off.visible=0
    me.enabled = false
  End If
End Sub

'**********************************************************************************************************
 Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
    With vpmDips
      .AddForm 700,400,"System 1 (Multi-Mode sound) - DIP switches"
      .AddFrame 0,0,190,"Coin chute control",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
      .AddFrame 0,46,190,"Game mode",&H00000400,Array("extra ball",0,"replay",&H00000400)'dip 11
      .AddFrame 0,92,190,"High game to date awards",&H00200000,Array("no award",0,"3 replays",&H00200000)'dip 22
      .AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3 balls",&H00000100)'dip 9
      .AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
      .AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
      .AddFrame 205,76,190,"Sound settings",&H80000000,Array("sounds",0,"tones",&H80000000)'dip 32
      .AddFrame 205,122,190,"Attract tune",&H10000000,Array("no attract tune",0,"attract tune played every 6 minutes",&H10000000)'dip 29
      .AddChk 205,175,190,Array("Match feature",&H00000200)'dip 10
      .AddChk 205,190,190,Array("Credits displayed",&H00001000)'dip 13
      .AddChk 205,205,190,Array("Play credit button tune",&H00002000)'dip 14
      .AddChk 205,220,190,Array("Play tones when scoring",&H00080000)'dip 20
      .AddChk 205,235,190,Array("Play coin switch tune",&H00400000)'dip 23
      .AddChk 205,250,190,Array("High game to date displayed",&H00100000)'dip 21
      .AddLabel 50,280,300,20,"After hitting OK, press F3 to reset game with new settings."
      .ViewDips
    End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

'******************************************************
'           FUNCTIONS
'******************************************************

'*** PI returns the value for PI
Function PI()
  PI = 4*Atn(1)
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

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    Atn2 = Sgn(dy) * (PI - Atn(Abs(dy / dx)))
  ElseIf dy = 0 Then
    Atn2 = 0
  Else
    Atn2 = Sgn(dy) * Pi / 2
  End If
End Function

'******************************************************
'           SOUNDS
'******************************************************

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

Sub PlaySoundAtExisting(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,aVol, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,aVol, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

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
    Vol = Csng(BallSpeed(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallSpeed(ball) * 20
End Function

Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*********************************************************************
'                       Collection Sounds
'*********************************************************************

Sub aApron_Hit(idx):PlaySoundAtBallVolM "fx_apron", Vol(ActiveBall)*VolumeDial:End Sub
Sub aRubbers_Hit(idx):PlaySoundAtBallVol "fx_rubber", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aPostRubbers_Hit(idx):PlaySoundAtBallVol "fx_postrubber", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aMetals_Hit(idx):PlaySoundAtBallVol "fx_MetalHit", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBallVol "fx_PlasticHit", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aGates_Hit(idx):PlaySoundAtBallVol "fx_Gate", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aTargets_hit(idx):PlaySoundAtBallVol "target", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aWoods_Hit(idx):PlaySoundAtBallVol "fx_Woodhit", Vol(ActiveBall)*VolumeDial/10:End Sub

Sub GateTrigger1_Hit()
  If Activeball.velx < 0 Then
    PlaySoundAtBallVol "fx_Gate", Vol(ActiveBall)*VolumeDial/10
  End If
End Sub

Sub GateTrigger2_Hit()
  If Activeball.velx > 0 Then
    PlaySoundAtBallVol "fx_Gate", Vol(ActiveBall)*VolumeDial/10
  End If
End Sub


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 2
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

  '*** Rolling Sounds***
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("fx_ballrolling" & b)
    BallShadow(b).visible = 0
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)
    If BallSpeed(BOT(b) ) > 1 AND BOT(b).z < 27  Then
      rolling(b) = True
      PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b))*RollingSoundFactor, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("fx_ballrolling" & b)
        rolling(b) = False
      End If
    End If

    '***Ball Shadows***
    BallShadow(b).X = BOT(b).X + (BOT(b).X - Table1.Width/2)/20
    BallShadow(b).Y = BOT(b).Y + 15 - (abs((BOT(b).X - Table1.Width/2)/20))/4
    BallShadow(b).visible = 1

    '***Ball Drop Sounds***

    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySound "fx_ball_drop" & b, 0, (ABS(BOT(b).velz)/17)^2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If
  Next

  cor.update
End Sub


'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF, RF1)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart, FlipperEnd, LR, PartialFlipCoef
  Private Balls(20), balldata(20)

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
  Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
  Public Property Get StartPoint : StartPoint = FlipperStart : End Property
  Public Property Let EndPoint(aInput) : if IsObject(aInput) then FlipperEnd = aInput.x else FlipperEnd = aInput : end if : End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
    if not DebugOn then exit sub
    dim a1, a2 : Select Case aChooseArray
      case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
      Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
      Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
      case else :tbpl.text = "wrong string" : exit sub
    End Select
    dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    tbpl.text = str
  End Sub

  Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

  Private Sub RemoveBall(aBall)
    dim x : for x = 0 to uBound(balls)
      if TypeName(balls(x) ) = "IBall" then
        if aBall.ID = Balls(x).ID Then
          balls(x) = Empty
          Balldata(x).Reset
        End If
      End If
    Next
  End Sub

  Public Sub Fire()
    Flipper.RotateToEnd
    processballs
  End Sub

  Public Property Get Pos 'returns % position a ball. For debug stuff.
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        balldata(x).Data = balls(x)
        if DebugOn then StickL.visible = True : StickL.x = balldata(x).x    'debug TODO
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 30 Then
      PartialFlipCoef = 0
    End If
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
      dim teststr : teststr = "Cutoff"
      tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
      if tmp < 0.1 then 'if real ball position is behind flipper, exit Sub to prevent stucks  'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
        if DebugOn then TestStr = "real pos < 0.1 ( " & round(tmp,2) & ")" : tbpl.text = Teststr
        'RemoveBall aBall
        'Exit Sub
      end if

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        if DebugOn then teststr = "y velocity: " & round(aBall.vely, 3) & "exit sub" : tbpl.text = teststr
        RemoveBall aBall
        exit Sub
      end if
      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          'TB.TEXT = balldata(x).id & " " & BALLDATA(X).X & VBNEWLINE & FLIPPERSTART & " " & FLIPPEREND
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        if DebugOn then set tmp = new spoofball : tmp.data = aBall : End If
        if IsEmpty(BallData(idx).id) and aBall.VelY < -12 then 'if tip hit with no collected data, do vel correction anyway
          if PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1) > 1.1 then 'adjust plz
            VelCoef = LinearEnvelope(5, VelocityIn, VelocityOut)
            if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
            if Enabled then aBall.Velx = aBall.Velx*VelCoef'VelCoef
            if Enabled then aBall.Vely = aBall.Vely*VelCoef'VelCoef
            if DebugOn then teststr = "tip protection" & vbnewline & "velcoef: " & round(velcoef,3) & vbnewline & round(PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1),3) & vbnewline
            'debug.print teststr
          end if
        Else
     :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
          if Enabled then aBall.Velx = aBall.Velx*VelCoef
          if Enabled then aBall.Vely = aBall.Vely*VelCoef
        end if
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      'debug
      if DebugOn then
        TestStr = teststr & "%pos:" & round(BallPos,2)
        if IsEmpty(PolarityOut(0) ) then
          teststr = teststr & vbnewline & "(Polarity Disabled)" & vbnewline
        else
          teststr = teststr & "+" & round(1 *(AddX*ycoef*PartialFlipcoef),3)
          if BallPos >= PolarityOut(uBound(PolarityOut) ) then teststr = teststr & "(MAX)" & vbnewline else teststr = teststr & vbnewline end if
          if Ycoef < 1 then teststr = teststr &  "ycoef: " & ycoef & vbnewline
          if PartialFlipcoef < 1 then teststr = teststr & "PartialFlipcoef: " & round(PartialFlipcoef,4) & vbnewline
        end if

        teststr = teststr & vbnewline & "Vel: " & round(BallSpeed(tmp),2) & " -> " & round(ballspeed(aBall),2) & vbnewline
        teststr = teststr & "%" & round(ballspeed(aBall) / BallSpeed(tmp),2)
        tbpl.text = TestSTR
      end if
    Else
      'if DebugOn then tbpl.text = "td" & timedelay
    End If
    RemoveBall aBall
  End Sub
End Class

'================================
'Helper Functions


Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  if offset < 0 then offset = 0
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub


Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
      id = .ID : mass = .mass : radius = .radius
    end with
  End Property
  Public Sub Reset()
    x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
    id = Empty : mass = Empty : radius = Empty
  End Sub
End Class


Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'Clamp if on the boundry lines
  'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
  'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity
dim RF1 : Set RF1 = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF, RF1)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    x.TimeDelay = 80
  Next

  'rf.report "Polarity"
  AddPt "Polarity", 0, 0, -2.7
  AddPt "Polarity", 1, 0.16, -2.7
  AddPt "Polarity", 2, 0.33, -2.7
  AddPt "Polarity", 3, 0.37, -2.7 '4.2
  AddPt "Polarity", 4, 0.41, -2.7
  AddPt "Polarity", 5, 0.45, -2.7 '4.2
  AddPt "Polarity", 6, 0.576,-2.7
  AddPt "Polarity", 7, 0.66, -1.8'-2.1896
  AddPt "Polarity", 8, 0.743, -0.5
  AddPt "Polarity", 9, 0.81, -0.5
  AddPt "Polarity", 10, 0.88, 0

  '"Velocity" Profile
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
  RF1.Object = RightFlipper1
  RF1.EndPoint = EndPointRp1
End Sub

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub
Sub TriggerRF1_Hit() : RF1.Addball activeball : End Sub
Sub TriggerRF1_UnHit() : RF1.PolarityCorrect activeball : End Sub

Sub RDampen_Timer()
  Cor.Update
End Sub

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    if gametime > 100 then Report
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

'Tracks ball velocity for judging bounce calculations & angle
'apologies to JimmyFingers is this is what his script does. I know his tracks ball velocity too but idk how it works in particular
dim cor : set cor = New CoRTracker
cor.debugOn = False
'cor.update() - put this on a low interval timer
Class CoRTracker
  public DebugOn 'tbpIn.text
  public ballvel

  Private Sub Class_Initialize : redim ballvel(0) : End Sub
  'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs
    if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      if DebugOn then
        dim s, bs 'debug spacer, ballspeed
        bs = round(BallSpeed(b),1)
        if bs < 10 then s = " " else s = "" end if
        str = str & b.id & ": " & s & bs & vbnewline
        'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
      end if
    Next
    if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
  End Sub
End Class


'**********************************************************************************************************
'Hanz's FSS BackGlass image/flash/primative move/rotate/setup and animation sequencer/player script start
'**********************************************************************************************************
Dim xoff, yoff, zoff, xrot,yrot,zrot, xcen, ycen,zcen, zscale


Dim BGArr(1)
BGArr(0) = Array( Flasher1, Flasher2, Flasher3, Flasher4, Flasher5, Flasher6) ',_
'Flasher11, Flasher12, Flasher13, Flasher14, Flasher15, Flasher16, Flasher17, Flasher18 )

Sub center_gfx(elem)
  Dim ii, xx, yy, yfact, xfact, objs
  'exit sub
  yoff = 0
  zscale = 0.00000001
  xcen =(1226 /2) - (104 / 2)
  ycen = (1110 /2 ) + (100 /2)

  yfact = 0
  xfact = -70

  For Each objs In elem
    xx =objs.x

    objs.x = (xoff - xcen) + xx + xfact
    yy = objs.y
    objs.y =yoff

      If yy < 0 then
      yy = yy * -1
      end if


    objs.height = (zoff - ycen) + yy - (yy * zscale) + yfact

    objs.rotx = xrot
  Next

end sub

Sub center_digits()

  Dim objd,ix,xx,yy,yfact,xfact

  zscale = 0.00000001
  xcen =(1226 /2) - (104 / 2)
  ycen = (1110 /2 ) + (100 /2)

  yoff = 0
  yfact =0 'y fudge factor (ycen was wrong so fix)
  xfact =-60

  for ix =0 to 27
    For Each objd in LEDFSS(ix)
      'MsgBox "Value at: " & ix
      xx =objd.x

      objd.x = (xoff -xcen) + xx +xfact
      yy = objd.y ' get the yoffset before it is changed
      objd.y =yoff

        If(yy < 0.) then
        yy = yy * -1
        end if

      objd.height =( zoff - ycen) + yy - (yy * (zscale)) + yfact

      objd.rotx = xrot

      'obj.visible = 0
      'end if
    Next
  Next

end sub


Sub hide_elem(elem)
  Dim objx

  For Each objx In elem ' hide all elements
    objx.visible =0
  Next
end sub


Sub show_elem(elem)
  Dim objx

  For Each objx In elem ' hide all elements
    objx.visible =1
  Next
end sub


' global backglass setup routine

Sub setup_backglass()

  Dim objx, ix

  xoff =625
  yoff =0
  zoff =730
  xrot = -90

  GBG.x = xoff
  GBG.y = yoff
  GBG.height = zoff
  GBG.rotx = xrot

  GBGframe.x = xoff
  GBGframe.y = yoff
  GBGframe.height = zoff
  GBGframe.rotx = xrot

  GBGframeMask.x = xoff
  GBGframeMask.y = yoff
  GBGframeMask.height = zoff
  GBGframeMask.rotx = xrot

  GBGframeMaskFill.x = xoff
  GBGframeMaskFill.y = yoff
  GBGframeMaskFill.height = zoff
  GBGframeMaskFill.rotx = xrot

  GBGHigh1.x = xoff
  GBGHigh1.y = yoff -10
  GBGHigh1.height = zoff
  GBGHigh1.rotx = xrot

  GBGHigh3.x = xoff
  GBGHigh3.y = yoff -10
  GBGHigh3.height = zoff
  GBGHigh3.rotx = xrot


   if Table1.ShowFSS = true or VROn = 1 or DesktopMode = False then

    center_gfx(BGArr(0))

    center_digits()

    for ix =0 to 27
      hide_elem(Digits(ix))  ' hide all digits
    Next

  Else ' ShowFSS = 0

    GBG.visible =0
    GBGframe.visible =0

    if  Table1.ShowDT = true then

      for ix =0 to 27
        show_elem(Digits(ix))  ' show all digits
      Next

    else ' FS

      for ix =0 to 27
        hide_elem(Digits(ix))  ' hide all digits
      Next

    end if

    hide_elem(BGArr(0))

    for ix =0 to 27
      hide_elem(LEDFSS(ix))  ' hide all leds
    Next

  End if
' setup animation sequencer slave/master
End Sub






Option Explicit
Randomize

'FSS On option (this isn't complete but the backbox is there)
  FSSOn = 0

Const VolDiv = 8000    ' Lower numper louder ballrolling sound
Const VolCol    = 3    ' Ball collision divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 8    ' Bumpers volume.
Const VolRol    = 5    ' Rollovers volume.
Const VolGates  = 5    ' Gates volume.
Const VolMetal  = 5    ' Metals volume.
Const VolRB     = 5    ' Rubber bands volume.
Const VolRH     = 5    ' Rubber hits volume.
Const VolPo     = 5    ' Rubber posts volume.
Const VolPi     = 5    ' Rubber pins volume.
Const VolPlast  = 5    ' Plastics volume.
Const VolTarg   = 5    ' Targets volume.
Const VolWood   = 5   ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = .1  ' Spinners volume.
Const VolFlip   = 5    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName="strlt_l1"

LoadVPM "01300000","S7.VBS",3.1

Const UseSolenoids=2
Const UseLamps=0
Const UseGI=0
Const UseSync=1

' Standard Sounds
Const SSolenoidOn="fx_solenoid",SSolenoidOff="fx_solenoidoff",SCoin="fx_coin",SFlipperOn="",SFlipperOff=""

Dim bsTrough,bsLSaucer,bsRSaucer,HiddenValue, FSSOn

'*********** Desktop/Cabinet settings ************************

If Table1.ShowDT = true Then
  HiddenValue = 0
  backbox.visible = 1
  railleft.visible = 1
  railright.visible = 1
Else
  HiddenValue = 1
  backbox.visible = 0
  railleft.visible = 0
  railright.visible = 0
End If
If FSSOn = 1 Then
  backbox.visible=1
  railleft.visible = 1
  railright.visible = 1
Else
  backbox.visible=0
End If

'******************************************************
'           FLIPPERS
'******************************************************

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySound SoundFX("fx_Flipperup",DOFContactors), VolFlip
    lf.fire
  Else
    if leftflipper.currentangle < leftflipper.startangle - 5 then
      PlaySound SoundFX("fx_flipperdown",DOFContactors), VolFlip
    end if
    LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySound SoundFX("fx_Flipperup1",DOFContactors), VolFlip
    RF.fire
  Else
    if RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      PlaySound SoundFX("fx_Flipperdown1",DOFContactors), VolFlip
    End If
    RightFlipper.RotateToStart
  End If
End Sub

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricksL LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricksR RightFlipper, RFPress, RFCount, RFEndAngle, RFState
end sub

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
Const EOSTnew = 1.5 'FEOST
Const EOSAnew = 0.2
Const EOSRampup = 1.5
Const SOSRampup = 8.5
Const LiveCatch = 8

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperTricksR (Flipper, FlipperPress, FCount, FEndAngle, FState)
  If Flipper.currentangle < Flipper.startangle + 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle + 3
      Flipper.Elasticity = FElasticity
      FCount = 0
      FState = 1
    End If
  ElseIf Flipper.currentangle = Flipper.endangle and FlipperPress = 1 then
    If FState <> 2  and FState Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      FState = 2
    End If

    if FCount = 0 Then FCount = GameTime

    if GameTime - FCount < LiveCatch Then
      Flipper.Elasticity = 0.1
      If Flipper.endangle <> FEndAngle Then Flipper.endangle = FEndAngle
    Else
      Flipper.Elasticity = FElasticity
    end if
  Elseif Flipper.currentangle < Flipper.endangle - 0.01 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If
  End If
End Sub

Sub FlipperTricksL (Flipper, FlipperPress, FCount, FEndAngle, FState)
  If Flipper.currentangle > Flipper.startangle - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle + 3
      Flipper.Elasticity = FElasticity
      FCount = 0
      FState = 1
    End If
  Elseif Flipper.currentangle = Flipper.endangle and FlipperPress = 1 then
    If FState <> 2  and FState Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      FState = 2
    End If

    if FCount = 0 Then FCount = GameTime

    if GameTime - FCount < LiveCatch Then
      Flipper.Elasticity = 0.1
      If Flipper.endangle <> FEndAngle Then Flipper.endangle = FEndAngle
    Else
      Flipper.Elasticity = FElasticity
    end if
  Elseif Flipper.currentangle > Flipper.endangle + 0.01 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If
  End If
End Sub

SolCallback(1)= "bsTrough.SolIn"
SolCallback(2)= "bsTrough.SolOut"
SolCallback(3)= "SolRSaucer"
SolCallback(4)= "Flasher1"
SolCallback(11)= "PFGI"
SolCallback(15)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'SolCallback(17)="vpmSolSound ""Slingshot"","
'SolCallback(18)="vpmSolSound ""Slingshot"","
'SolCallback(19)="vpmSolSound ""Bumper"","
'SolCallback(20)="vpmSolSound ""Bumper"","
'SolCallback(21)="vpmSolSound ""Bumper"","
SolCallback(25)="vpmNudge.SolGameOn"


SolCallback(sLLFlipper)="SolLFlipper"
SolCallback(sLRFlipper)="SolRFlipper"

Sub Table1_Init
   PFGI(False)
  Flasher1(False)
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Star Light"&chr(13)&"(Williams 1984)"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .Hidden = HiddenValue
        .Run
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

  PinmameTimer.Interval=PinMameInterval
  PinmameTimer.Enabled=1

  Set bsTrough=New cvpmBallstack
  with bsTrough
    .InitSw 9,10,11,0,0,0,0,0
    .InitKick BallRelease,90,7
    .InitExitSnd Soundfx("fx2_ballrel",DOFContactors), Soundfx("fx2_solenoid",DOFContactors)
    .Balls=2
  end with

  Set bsRSaucer=New cvpmBallStack
  with bsRSaucer
  end with

  vpmNudge.TiltSwitch=1
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(sw41, sw42, sw43, LeftSlingshot, RightSlingshot)
End Sub

Dim BootCount:BootCount = 0

Sub BootTable_Timer()
  If BootCount = 0 Then
    BootCount = 1
    PlaySoundAt "poweron", Plunger
  Else
    DOF 125,1
    backbox.image = "backboxon2"
    me.enabled = false
  End If
End Sub

Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftTiltKey Then Nudge 90, 2
  If keycode = RightTiltKey Then Nudge 270, 2
  If keycode = CenterTiltKey Then Nudge 0, 2

  If keycode = LeftFlipperKey Then LFPress = 1
  If keycode = RightFlipperKey Then rfpress = 1
  If Keycode = RightMagnaSave Then Controller.Switch(54)=1

  If vpmKeyDown(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.PullBack: PlaySoundAt "fx_plungerpull", Plunger:   End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then Plunger.Fire: PlaySoundAt "fx_plunger", Plunger

  If keycode = LeftFlipperKey Then
    lfpress = 0
    leftflipper.eostorqueangle = EOSA
    leftflipper.eostorque = EOST
  End If
  If keycode = RightFlipperKey Then
    rfpress = 0
    rightflipper.eostorqueangle = EOSA
    rightflipper.eostorque = EOST
  End If
  If keycode = RightMagnaSave Then Controller.Switch(54)=0
  If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub Drain_Hit()
  PlaySoundAt "fx_drain", Drain
  RF.PolarityCorrect Activeball
  LF.PolarityCorrect Activeball
  bstrough.addball me
End Sub

'Playfield GI
Sub PFGI(Enabled)
  If Enabled Then  ' GI state inverted, enabled = OFF
    dim xx
    For each xx in GI:xx.State = 0: Next
    inserts_off.visible = 1
    inserts.visible = 0
    playfield_off.visible = 1
    For each xx in layer1col: xx.image = "layer1_off": Next
    For each xx in layer2col: xx.image = "layer2_off": Next
    For each xx in layer3col: xx.image = "layer3_off": Next
    For each xx in layer4col: xx.image = "layer3_off": Next
    railright.image="railrightoff"
    railleft.image="railleftoff"
    brackets.image="bracketsoff"
        PlaySound "fx_relay"
    DOF 125,0
    backbox.image = "backboxoff2"
  Else
    For each xx in GI:xx.State = 1: Next
    inserts_off.visible = 0
    inserts.visible = 1
    playfield_off.visible = 0
    For each xx in layer1col: xx.image = "layer1": Next
    For each xx in layer2col: xx.image = "layer2": Next
    For each xx in layer3col: xx.image = "layer3": Next
    For each xx in layer4col: xx.image = "layer4": Next
    railright.image="raillright"
    railleft.image="railleft"
    brackets.image="bracketson"
        PlaySound "fx_relay"
    DOF 125,1
    backbox.image = "backboxon2"
  End If
  ' Update lamps that may swap textures when GI changes here
End Sub

Sub Flasher1 (Enabled)
  If Enabled Then
    Flasha.visible=1
    backbox.image = "backboxflash2"
  Else
    Flasha.visible=0
    if light018.state=0 then
      backbox.image = "backboxoff2"
    else
      backbox.image = "backboxon2"
    end if
  End If
End Sub
'
'******************************************************
'           Saucer
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


Dim KickerBall

dim kickstep1, kickstep2

Sub sw52_Hit
  set KickerBall = activeball
  Controller.Switch(52) = 1
  PlaySoundAtVol "fx_kicker_enter", sw52, 3
End Sub

Sub sw52_unHit
  Controller.Switch(52) = 0
End Sub

Sub SolRSaucer(Enable)
    If Enable then
'   bsSaucer.ExitSol_On
    PlaySoundAtVol SoundFX("fx2_solenoid",DOFContactors), sw52, 3
    kickstep2 = 0
    pkickarm001.rotz=15
    If sw52.ballcntover > 0 then
      KickBall KickerBall, -140, 14, 5, 30
      PlaySoundAtVol "fx_kicker", sw52, 3
    End If
    sw52.timerenabled = 1
    End If
End Sub

Sub sw52_timer
    Select Case kickstep2
        Case 3:pkickarm001.rotz=15
        Case 4:pkickarm001.rotz=15
        Case 5:pkickarm001.rotz=15
        Case 6:pkickarm001.rotz=15
        Case 7:pkickarm001.rotz=8
        Case 8:pkickarm001.rotz=3
    Case 9:pkickarm001.rotz=0:sw52.TimerEnabled = 0: if sw52.ballcntover > 0 then SolRSaucer -1
    End Select
   kickstep2 = kickstep2 + 1
End Sub

''***********************************
'' Bumpers
''***********************************

dim bump1step, bump2step, bump3step
'
Sub sw42_Hit
  vpmTimer.PulseSw 42
  playsoundAtVol SoundFX("fx2_bumper1",DOFContactors), sw42, VolBump
  bumpskirt1.RotY=-3
  Bump1Step = 0
  sw42.timerenabled = 1
End Sub

Sub sw42_timer
  Select Case Bump1Step
    Case 3: bumpskirt1.RotY=-1
    Case 4: bumpskirt1.RotY=1
    Case 5: bumpskirt1.RotY=-1
    Case 6: bumpskirt1.RotY=1
    Case 7: bumpskirt1.RotY=0:sw42.timerenabled = 0
  End Select
  Bump1Step=Bump1Step + 1
End Sub

Sub sw41_Hit
  vpmTimer.PulseSw 41
  playsoundAtVol SoundFX("fx2_bumper2",DOFContactors), sw41, VolBump
  bumpskirt2.RotY=-5
  bump2step = 0
  sw41.timerenabled = 1
End Sub

Sub sw41_timer
  Select Case bump2step
    Case 3: bumpskirt2.RotY=-3
    Case 4: bumpskirt2.RotY=2
    Case 5: bumpskirt2.RotY=-1
    Case 6: bumpskirt2.RotY=1
    Case 7: bumpskirt2.RotY=0:sw41.timerenabled = 0
  End Select
  bump2step=bump2step + 1
End Sub

Sub sw43_Hit
  vpmTimer.PulseSw 43
  playsoundAtVol SoundFX("fx2_bumper1",DOFContactors), sw43, VolBump
  bumpskirt3.RotY=-5
  bump3step = 0
  sw43.timerenabled = 1
End Sub

Sub sw43_timer
  Select Case bump3step
    Case 3: bumpskirt3.RotY=-3
    Case 4: bumpskirt3.RotY=2
    Case 5: bumpskirt3.RotY=-1
    Case 6: bumpskirt3.RotY=1
    Case 7: bumpskirt3.RotY=0:sw43.timerenabled = 0
  End Select
  bump3step=bump3step + 1
End Sub
'
''Rollovers
Sub SW26_Hit:Controller.Switch(26)=1:star26.transz=-6:End Sub
Sub SW26_unHit:Controller.Switch(26)=0:star26.transz=0:End Sub
Sub SW27_Hit:Controller.Switch(27)=1:star27.transz=-6:End Sub
Sub SW27_unHit:Controller.Switch(27)=0:star27.transz=0:End Sub
Sub SW28_Hit:Controller.Switch(28)=1:star28.transz=-6:End Sub
Sub SW28_unHit:Controller.Switch(28)=0:star28.transz=0:End Sub
Sub SW29_Hit:Controller.Switch(29)=1:star29.transz=-6:End Sub
Sub SW29_unHit:Controller.Switch(29)=0:star29.transz=0:End Sub
Sub SW30_Hit:Controller.Switch(30)=1:star30.transz=-6:End Sub
Sub SW30_unHit:Controller.Switch(30)=0:star30.transz=0:End Sub
Sub SW31_Hit:Controller.Switch(31)=1:star31.transz=-6:End Sub
Sub SW31_unHit:Controller.Switch(31)=0:star31.transz=0:End Sub
Sub SW32_Hit:Controller.Switch(32)=1:star32.transz=-6:End Sub
Sub SW32_unHit:Controller.Switch(32)=0:star32.transz=0:End Sub
Sub SW33_Hit:Controller.Switch(33)=1:star33.transz=-6:End Sub
Sub SW33_unHit:Controller.Switch(33)=0:star33.transz=0:End Sub
Sub SW34_Hit:Controller.Switch(34)=1:star34.transz=-6:End Sub
Sub SW34_unHit:Controller.Switch(34)=0:star34.transz=0:End Sub
Sub SW35_Hit:Controller.Switch(35)=1:star35.transz=-6:End Sub
Sub SW35_unHit:Controller.Switch(35)=0:star35.transz=0:End Sub

'
'Wire Triggers
Sub SW12_Hit:Controller.Switch(12)=1:tr12.transz=-6:End Sub
Sub SW12_unHit:Controller.Switch(12)=0:tr12.transz=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1:tr23.transz=-6:End Sub
Sub SW23_unHit:Controller.Switch(23)=0:tr23.transz=0:End Sub
Sub SW36_Hit:Controller.Switch(36)=1:tr36.transz=-6: End Sub
Sub SW36_unHit:Controller.Switch(36)=0:tr36.transz=0:End Sub
Sub SW37_Hit:Controller.Switch(37)=1:tr37.transz=-6:End Sub
Sub SW37_unHit:Controller.Switch(37)=0:tr37.transz=0:End Sub
Sub SW38_Hit:Controller.Switch(38)=1:tr38.transz=-6:End Sub
Sub SW38_unHit:Controller.Switch(38)=0:tr38.transz=0:End Sub
Sub SW39_Hit:Controller.Switch(39)=1:tr39.transz=-6:End Sub
Sub SW39_unHit:Controller.Switch(39)=0:tr39.transz=0:End Sub
Sub SW40_Hit:Controller.Switch(40)=1:tr40.transz=-6:End Sub
Sub SW40_unHit:Controller.Switch(40)=0:tr40.transz=0:End Sub
Sub SW44_Hit:Controller.Switch(44)=1:tr44.transz=-6:End Sub
Sub SW44_unHit:Controller.Switch(44)=0:tr44.transz=0:End Sub
Sub SW45_Hit:Controller.Switch(45)=1:tr45.transz=-6:End Sub
Sub SW45_unHit:Controller.Switch(45)=0:tr45.transz=0:End Sub
Sub SW46_Hit:Controller.Switch(46)=1:tr46.transz=-6:End Sub
Sub SW46_unHit:Controller.Switch(46)=0:tr46.transz=0:End Sub
Sub SW47_Hit:Controller.Switch(47)=1:tr47.transz=-6:End Sub
Sub SW47_unHit:Controller.Switch(47)=0:tr47.transz=0:End Sub

'
'leaf switches
Sub sw49_Hit
  vpmTimer.PulseSw 49
End Sub

Sub sw48_Hit
  vpmTimer.PulseSw 48
End Sub

'Spinners
Sub sw25_Spin : vpmTimer.PulseSw (25) :PlaySoundAtVol "fx2_spinner", sw25, VolSpin: End Sub

'***********Rotate Spinner
Dim Angle

Sub SpinnerTimer_Timer
' Angle = (sin (sw17.CurrentAngle-180))
'    SpinnerRod.TransZ = -sin( (sw17.CurrentAngle+180) * (2*3.14/360)) * 5
'    SpinnerRod.TransX = (sin( (sw17.CurrentAngle- 90) * (2*3.14/360)) * -5)
  Pgate001.rotx = -Gate001.currentangle*0.5
  Pgate002.rotx = -Gate002.currentangle*0.5
  Pgate003.rotx = -Gate004.currentangle*0.5
  Dim SpinnerRadius: SpinnerRadius=7

  SpinnerRod.TransZ = (cos((sw25.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRod.TransY = sin((sw25.CurrentAngle) * (PI/180)) * -SpinnerRadius
End Sub
'

Dim Step13, Step14, Step15, Step16, Step17, Step18, Step19, Step20, Step21, Step22, Step22a

Sub sw13_Hit
  vpmTimer.PulseSw (13)
  sw13target.TransX=5
  sw13.timerenabled = 1
  Step13 = 0
end sub

Sub sw13_timer
  Select Case Step13
    Case 3: sw13target.TransX=-3
    Case 4: sw13target.TransX=4
    Case 5: sw13target.TransX=-2
    Case 6: sw13target.TransX=3
    Case 7: sw13target.TransX=-1
    Case 8: sw13target.TransX=2
    Case 9: sw13target.TransX=0
    Case 10: sw13target.TransX=-1: sw13.timerenabled = 0
  End Select
  Step13=Step13 + 1
End Sub

Sub sw14_Hit
  vpmTimer.PulseSw (14)
  sw14target.TransX=5
  sw14.timerenabled = 1
  Step14 = 0
end sub

Sub sw14_timer
  Select Case Step14
    Case 3: sw14target.TransX=-3
    Case 4: sw14target.TransX=4
    Case 5: sw14target.TransX=-2
    Case 6: sw14target.TransX=3
    Case 7: sw14target.TransX=-1
    Case 8: sw14target.TransX=2
    Case 9: sw14target.TransX=0
    Case 10: sw14target.TransX=-1: sw14.timerenabled = 0
  End Select
  Step14=Step14 + 1
End Sub

Sub sw15_Hit
  vpmTimer.PulseSw (15)
  sw15target.TransX=5
  sw15.timerenabled = 1
  Step15 = 0
end sub

Sub sw15_timer
  Select Case Step15
    Case 3: sw15target.TransX=-3
    Case 4: sw15target.TransX=4
    Case 5: sw15target.TransX=-2
    Case 6: sw15target.TransX=3
    Case 7: sw15target.TransX=-1
    Case 8: sw15target.TransX=-2
    Case 9: sw15target.TransX=0
    Case 10: sw15target.TransX=-1: sw15.timerenabled = 0
  End Select
  Step15=Step15 + 1
End Sub

Sub sw16_Hit
  vpmTimer.PulseSw (16)
  sw16target.TransX=5
  sw16.timerenabled = 1
  Step16 = 0
end sub

Sub sw16_timer
  Select Case Step16
    Case 3: sw16target.TransX=-3
    Case 4: sw16target.TransX=4
    Case 5: sw16target.TransX=-2
    Case 6: sw16target.TransX=3
    Case 7: sw16target.TransX=-1
    Case 8: sw16target.TransX=2
    Case 9: sw16target.TransX=0
    Case 10: sw16target.TransX=-1: sw16.timerenabled = 0
  End Select
  Step16=Step16 + 1
End Sub

Sub sw17_Hit
  vpmTimer.PulseSw (17)
  sw17target.TransX=5
  sw17.timerenabled = 1
  Step17 = 0
end sub

Sub sw17_timer
  Select Case step17
    Case 3: sw17target.TransX=-3
    Case 4: sw17target.TransX=4
    Case 5: sw17target.TransX=-2
    Case 6: sw17target.TransX=3
    Case 7: sw17target.TransX=-1
    Case 8: sw17target.TransX=2
    Case 9: sw17target.TransX=0
    Case 10: sw17target.TransX=-1: sw17.timerenabled = 0
  End Select
  Step17=Step17 + 1
End Sub

Sub sw18_Hit
  vpmTimer.PulseSw (18)
  sw18target.TransX=5
  sw18.timerenabled = 1
  Step18 = 0
end sub

Sub sw18_timer
  Select Case step18
    Case 3: sw18target.TransX=-3
    Case 4: sw18target.TransX=4
    Case 5: sw18target.TransX=-2
    Case 6: sw18target.TransX=3
    Case 7: sw18target.TransX=-1
    Case 8: sw18target.TransX=2
    Case 9: sw18target.TransX=0
    Case 10: sw18target.TransX=-1: sw18.timerenabled = 0
  End Select
  Step18=Step18 + 1
End Sub

Sub sw19_Hit
  vpmTimer.PulseSw (19)
  sw19target.TransX=5
  sw19.timerenabled = 1
  Step19 = 0
end sub

Sub sw19_timer
  Select Case step19
    Case 3: sw19target.TransX=-3
    Case 4: sw19target.TransX=4
    Case 5: sw19target.TransX=-2
    Case 6: sw19target.TransX=3
    Case 7: sw19target.TransX=-1
    Case 8: sw19target.TransX=2
    Case 9: sw19target.TransX=0
    Case 10: sw19target.TransX=-1: sw19.timerenabled = 0
  End Select
  Step19=Step19 + 1
End Sub

Sub sw20_Hit
  vpmTimer.PulseSw (20)
  sw20target.TransX=5
  sw20.timerenabled = 1
  Step20 = 0
end sub

Sub sw20_timer
  Select Case step20
    Case 3: sw20target.TransX=-3
    Case 4: sw20target.TransX=4
    Case 5: sw20target.TransX=-2
    Case 6: sw20target.TransX=3
    Case 7: sw20target.TransX=-1
    Case 8: sw20target.TransX=2
    Case 9: sw20target.TransX=0
    Case 10: sw20target.TransX=-1: sw20.timerenabled = 0
  End Select
  Step20=Step20 + 1
End Sub

Sub sw21_Hit
  vpmTimer.PulseSw (21)
  sw21target.TransX=5
  sw21.timerenabled = 1
  Step21 = 0
end sub

Sub sw21_timer
  Select Case step21
    Case 3: sw21target.TransX=-3
    Case 4: sw21target.TransX=4
    Case 5: sw21target.TransX=-2
    Case 6: sw21target.TransX=3
    Case 7: sw21target.TransX=-1
    Case 8: sw21target.TransX=2
    Case 9: sw21target.TransX=0
    Case 10: sw21target.TransX=-1: sw21.timerenabled = 0
  End Select
  Step21=Step21 + 1
End Sub
'
Sub sw22_Hit
  vpmTimer.PulseSw (22)
  sw22target.TransX=5
  sw22.timerenabled = 1
  Step22 = 0
end sub

Sub sw22_timer
  Select Case Step22
    Case 3: sw22target.TransX=-3
    Case 4: sw22target.TransX=4
    Case 5: sw22target.TransX=-2
    Case 6: sw22target.TransX=3
    Case 7: sw22target.TransX=-1
    Case 8: sw22target.TransX=2
    Case 9: sw22target.TransX=0
    Case 10: sw22target.TransX=-1: sw22.timerenabled = 0
  End Select
  Step22=Step22 + 1
End Sub

Sub sw24_Hit
  vpmTimer.PulseSw (24)
  sw22atarget.TransX=5
  sw24.timerenabled = 1
  Step22a = 0
end sub

Sub sw24_timer
  Select Case Step22a
    Case 3: sw22atarget.TransX=-3
    Case 4: sw22atarget.TransX=4
    Case 5: sw22atarget.TransX=-2
    Case 6: sw22atarget.TransX=3
    Case 7: sw22atarget.TransX=-1
    Case 8: sw22atarget.TransX=2
    Case 9: sw22atarget.TransX=0
    Case 10: sw22atarget.TransX=-1: sw24.timerenabled = 0
  End Select
  Step22a=Step22a + 1
End Sub
''

'
Sub SolKnocker(Enabled)
  If Enabled Then PlaySound SoundFX("Fx2_Knocker",DOFKnocker)
End Sub
'
''**********Sling Shot Animations
'' Rstep and Lstep  are the variables that increment the animation
''****************
Dim RStep, Lstep
'
Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx2_slingshot1",DOFContactors), sling1
    vpmtimer.PulseSw(51)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.Rotx = 22
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.Rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing3.Visible = 1:sling1.Rotx = 0
        Case 5:RSLing3.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub
'
Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx2_slingshot2",DOFContactors),sling2
  vpmtimer.pulsesw(50)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.Rotx = 22
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.Rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing3.Visible = 1:sling2.Rotx = 0
        Case 5:LSLing3.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub
'
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
'NFadeL 1, L1 'Game Over
'NFadeL 2, L2 'match
'NFadeL 3, L3 'tilt
'NFadeL 4, L4 'highscoretodate
'NFadeL 5, L5 'shoot again
'NFadeL 6, L6 'BIP
NFadeL 7, l7
NFadeL 8, l8
NFadeL 9, l9
NFadeL 10, l10
NFadeL 11, l11
NFadeL 12, l12
NFadeL 13, l13
NFadeL 14, l14
NFadeL 15, l15
NFadeL 16, l16
NFadeL 17, l17
NFadeL 18, l18
NFadeL 19, l19
NFadeL 20, l20
NFadeL 21, l21
NFadeL 22, l22
NFadeL 23, l23
NFadeL 24, l24
NFadeL 25, l25
NFadeL 26, l26
NFadeL 27, l27
NFadeL 28, l28
NFadeL 29, l29
NFadeL 30, l30
NFadeL 31, l31
NFadeL 32, l32
NFadeL 33, l33
NFadeL 34, l34
NFadeL 35, l35
NFadeL 36, l36
NFadeL 37, l37
NFadeL 38, l38
NFadeL 39, l39
NFadeL 40, l40
NFadeL 41, l41
NFadeL 42, l42
NFadeL 43, l43
NFadeL 44, l44
NFadeL 45, l45
NFadeL 46, l46
NFadeL 47, l47
NFadeL 48, l48
NFadeL 49, l49
NFadeL 50, l50
NFadeL 51, l51
NFadeL 52, l52
NFadeL 53, l53
NFadeL 54, l54
NFadeL 55, l55
NFadeL 56, l56
NFadeL 57, l57
NFadeL 58, l58
NFadeL 59, l59
NFadeL 60, l60
NFadeL 61, l61
NFadeL 62, l62
NFadeL 63, l63
NFadeL 64, l64

'Solenoid Controlled

'NFadeL 104, S104

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

'
'*****************************************
'     FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh001.RotZ = LeftFlipper.currentangle
  FlipperRSh001.RotZ = RightFlipper.currentangle
' If l59.state=1 Then
'   if light001.state=1 Then
'   apronlamp.image="apronlamp_lit"
'   end if
' Else
'   if light001.state=1 Then
'   apronlamp.image="apronlamp"
'   end if
' End If
End Sub

'*****************************************
'     BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow001,BallShadow002,BallShadow003,BallShadow004,BallShadow005)

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
            BallShadow(b).X = ((BOT(b).X) + ((BOT(b).X - (Table1.Width/2))/7))
        Else
            BallShadow(b).X = ((BOT(b).X) + ((BOT(b).X - (Table1.Width/2))/7))
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

''******************************
'' Diverse Collection Hit Sounds
''******************************
'
'Sub aBumpers_Hit (idx): PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, pan(ActiveBall): End Sub
Sub aRollovers_Hit(idx):PlaySound "fx2_sensor", 0, Vol(ActiveBall)*VolRol, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx2_MetalHit2", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Metals2_Hit(idx):PlaySound "fx2_Hit_Metal", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx2_rubber", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aTargets_Hit(idx):PlaySound "fx2_target", 0, Vol(ActiveBall)*VolTarg, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx2_Woodhit3", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_Exit:Controller.Stop:End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
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
    StopSound("fx_Rolling_Wood" & b)
    StopSound("fx_Rolling_Plastic" & b)
    StopSound("fx_Rolling_Metal" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)
    If BallVel(BOT(b) ) > 1 Then

      ' ***Ball on WOOD playfield***
      if BOT(b).z < 27 Then
        PlaySound("fx_Rolling_Wood" & b), -1, Vol(BOT(b) )/2, AudioPan(BOT(b) )/5, 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        rolling(b) = True
        StopSound("fx_Rolling_Plastic" & b)
        StopSound("fx_Rolling_Metal" & b)
      ' ***Ball on PLASTIC ramp***
      Elseif  BOT(b).z > 28 Then
        PlaySound("fx_Rolling_Plastic" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        rolling(b) = True
        StopSound("fx_Rolling_Wood" & b)
        StopSound("fx_Rolling_Metal" & b)
      Else
        if rolling(b) = true Then
          rolling(b) = False
          StopSound("fx_Rolling_Wood" & b)
          StopSound("fx_Rolling_Plastic" & b)
          StopSound("fx_Rolling_Metal" & b)
        end if
      End If
    Else
      If rolling(b) = True Then
        StopSound("fx_Rolling_Wood" & b)
        StopSound("fx_Rolling_Plastic" & b)
        StopSound("fx_Rolling_Metal" & b)
        rolling(b) = False
      End If

    End If

    '***Ball Drop Sounds***

    If BOT(b).VelZ < -1 and BOT(b).z < 50 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
    End If
  Next
End Sub

Sub PlaySoundAtBOTBall(sound, BOT)
    PlaySound sound, 0, Vol(BOT), AudioPan(BOT), 0, Pitch(BOT), 0, 1, AudioFade(BOT)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, AudioPan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

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

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

  '"Polarity" Profile
'"Polarity" Profile<br>
' AddPt "Polarity", 0, 0, 0
' AddPt "Polarity", 1, 0.1, 0
' AddPt "Polarity", 2, 0.14, -2.25
' AddPt "Polarity", 3, 0.2, -2.25
' AddPt "Polarity", 4, 0.28, -3.25
' AddPt "Polarity", 5, 0.31, -3.25
' AddPt "Polarity", 6, 0.34, -3.75
' AddPt "Polarity", 7, 0.37, -3.75
' AddPt "Polarity", 8, 0.4, -4.5
' AddPt "Polarity", 9, 0.45, -3.5
' AddPt "Polarity", 10, 0.48, -3.5
' AddPt "Polarity", 11, 0.51, -3.75
' AddPt "Polarity", 12, 0.55, -3.75
' AddPt "Polarity", 13, 0.58, -3
' AddPt "Polarity", 14, 0.6, -2.75
' AddPt "Polarity", 15, 0.62, -2.75
' AddPt "Polarity", 16, 0.65, -2.5
' AddPt "Polarity", 17, 0.8, -2
' AddPt "Polarity", 18, 0.85, -1.9
' AddPt "Polarity", 19, 1.0, -1
' AddPt "Polarity", 20, 1.2, 0

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

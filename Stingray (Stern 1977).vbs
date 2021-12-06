Option Explicit
Randomize

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"

' Arngrim 2019-11-24
' Splitted shared switches with DOF Commands for DOF Config in the configtool

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

Const cGameName="stingray"

LoadVPM "01530000","BALLY.VBS",3.2
'LoadVPM "01530000","STERN.VBS",3.2

Const UseSolenoids=2
Const UseLamps=1
Const UseGI=0

' Standard Sounds
Const SSolenoidOn="fx_solenoid",SSolenoidOff="fx_solenoidoff",SCoin="fx_coin",SFlipperOn="",SFlipperOff=""

Dim bsTrough,bsSaucer,dtL,HiddenValue

'*********** Desktop/Cabinet settings ************************

If Table1.ShowDT = true Then
  HiddenValue = 0
  rails.visible=1
Else
  HiddenValue = 1
  rails.visible=0
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

  If leftflipper.currentangle = leftflipper.endangle and LFPress = 1 then
    leftflipper.eostorqueangle = EOSAnew
    leftflipper.eostorque = EOSTnew
    LeftFlipper.rampup = EOSRampup
    if LFCount = 0 Then LFCount = GameTime
    if GameTime - LFCount < LiveCatch Then
      leftflipper.Elasticity = 0.1
      If LeftFlipper.endangle <> LFEndAngle Then leftflipper.endangle = LFEndAngle
    Else
      leftflipper.Elasticity = FElasticity
    end if
  elseif leftflipper.currentangle > leftflipper.startangle - 0.05  Then
    leftflipper.rampup = SOSRampup
    leftflipper.endangle = LFEndAngle - 3
    leftflipper.Elasticity = FElasticity
    LFCount = 0
  elseif leftflipper.currentangle > leftflipper.endangle + 0.01 Then
    leftflipper.eostorque = EOST
    leftflipper.eostorqueangle = EOSA
    LeftFlipper.rampup = Frampup
    leftflipper.Elasticity = FElasticity
  end if

  If rightflipper.currentangle = rightflipper.endangle and RFPress = 1 then
    rightflipper.eostorqueangle = EOSAnew
    rightflipper.eostorque = EOSTnew
    RightFlipper.rampup = EOSRampup
    if RFCount = 0 Then RFCount = GameTime
    if GameTime - RFCount < LiveCatch Then
      rightflipper.Elasticity = 0.1
      If RightFlipper.endangle <> RFEndAngle Then rightflipper.endangle = RFEndAngle
    Else
      rightflipper.Elasticity = FElasticity
    end if
  elseif rightflipper.currentangle < rightflipper.startangle + 0.05 Then
    rightflipper.rampup = SOSRampup
    rightflipper.endangle = RFEndAngle + 3
    rightflipper.Elasticity = FElasticity
    RFCount = 0
  elseif rightflipper.currentangle < rightflipper.endangle - 0.01 Then
    rightflipper.eostorque = EOST
    rightflipper.eostorqueangle = EOSA
    RightFlipper.rampup = Frampup
    rightflipper.Elasticity = FElasticity
  end if

end sub

dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSAnew
dim FStrength, Frampup, FElasticity, EOSRampup, SOSRampup
dim RFEndAngle, LFEndAngle, LFCount, RFCount, LiveCatch

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
FStrength = LeftFlipper.strength
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
EOSTnew = 1.5 'FEOST
EOSAnew = 0.2
EOSRampup = 1.5
SOSRampup = 8.5
LiveCatch = 12

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle


'Solenoids
'SolCallback(1)="vpmSolSound SoundFX(""bell10"",DOFChimes),"
'SolCallback(2)="vpmSolSound SoundFX(""sj_chime_10a"",DOFChimes),"
'SolCallback(3)="vpmSolSound SoundFX(""sj_chime_100a"",DOFChimes),"
'SolCallBack(4)="vpmSolSound SoundFX(""sj_chime_1000a"",DOFChimes),"
SolCallback(1)="SolChime1"
SolCallback(2)="SolChime10"
SolCallback(3)="SolChime100"
SolCallBack(4)="SolChime1000"
SolCallback(6)="Solknocker"
SolCallback(7)="bsTrough.SolOut"
SolCallback(9)="solsaucer"
SolCallback(10)="SolLeftTargetReset"
'SolCallback(11)="vpmSolSound ""jet3"","  'bumper1
'SolCallback(12)="vpmSolSound ""jet3"","  'bumper2
'SolCallback(13)="vpmSolSound ""jet3"","  'bumper3
'SolCallback(14)="vpmSolSound ""sling""," 'right sling
'SolCallback(15)="vpmSolSound ""sling""," 'left sling
SolCallback(19)="vpmNudge.SolGameOn"

SolCallback(sLLFlipper)="SolLFlipper"
SolCallback(sLRFlipper)="SolRFlipper"

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Dracula"&chr(13)&"(Stern 1979)"
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
    .InitSw 0,8,0,0,0,0,0,0
'   .InitNoTrough BallRelease,8,125,5
    .InitKick BallRelease,90,7
    .InitExitSnd Soundfx("fx2_ballrel",DOFContactors), Soundfx("fx2_solenoid",DOFContactors)
    .Balls=1
  end with

' Set bsSaucer=New cvpmBallStack
' with bsSaucer
'   .InitSaucer sw40,40,108,14
'   .InitExitSnd Soundfx("fx2_ballrel",DOFContactors), Soundfx("fx2_solenoid",DOFContactors)
'   .CreateEvents "bsSaucer", sw40
'   end with
'
  Set dtL=New cvpmDropTarget
  dtL.InitDrop Array(sw25,sw26,sw27,sw28,sw29),Array(25,26,27,28,29)
  dtL.InitSnd SoundFX("fx_droptarget",DOFContactors),SoundFX("fx2_DTReset",DOFContactors)

  vpmNudge.TiltSwitch=7
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(sw25,sw26,sw27,sw28,sw29,LeftSlingshot,RightSlingshot)
End Sub

Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftTiltKey Then Nudge 90, 2
  If keycode = RightTiltKey Then Nudge 270, 2
  If keycode = CenterTiltKey Then Nudge 0, 2

  If keycode = LeftFlipperKey Then LFPress = 1
  If keycode = RightFlipperKey Then rfpress = 1

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

  If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub Drain_Hit()
  PlaySoundAt "fx_drain", Drain
  RF.PolarityCorrect Activeball
  LF.PolarityCorrect Activeball
  bstrough.addball me
End Sub

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

'
Sub sw12b_Hit
  set KickerBall = activeball
  Controller.Switch(12) = 1
End Sub

Sub sw12b_unHit
  Controller.Switch(12) = 0
End Sub

Sub sw12a_Hit
  set KickerBall = activeball
  Controller.Switch(12) = 1
End Sub

Sub sw12a_unHit
  Controller.Switch(12) = 0
End Sub

dim kickstep1

Sub SolSaucer(Enable)
    If Enable then
    PlaySoundAtVol "saucerin", sw12a, 3
'   bsSaucer.ExitSol_On
    kickstep1 = 0
    pkickarm.objrotz=15:pkickarm001.objrotz=15
    If sw12a.ballcntover > 0 then
      PlaySoundAtVol "saucerkick",sw12a, 3
      KickBall KickerBall, 200, 14, 5, 30
    End If
    If sw12b.ballcntover > 0 then
      PlaySoundAtVol "saucerkick",sw12b, 3
      KickBall KickerBall, 190, 14, 5, 30
    End If
    sw12b.timerenabled = 1
    End If
End Sub

Sub sw12b_timer
    Select Case kickstep1
        Case 3:pkickarm.objrotz=15:pkickarm001.objrotz=15
        Case 4:pkickarm.objrotz=15:pkickarm001.objrotz=15
        Case 5:pkickarm.objrotz=15:pkickarm001.objrotz=15
        Case 6:pkickarm.objrotz=15:pkickarm001.objrotz=15
        Case 7:pkickarm.objrotz=8:pkickarm001.objrotz=8
        Case 8:pkickarm.objrotz=3:pkickarm001.objrotz=3
    Case 9:pkickarm.objrotz=0:pkickarm001.objrotz=0:sw12b.TimerEnabled = 0: if sw12b.ballcntover > 0 then SolSaucer -1
    End Select
   kickstep1 = kickstep1 + 1
End Sub

dim d25step, d26step, d27step, d28step, d29step

'Drop Targets
Sub sw25_Hit: dtl.Hit 1:playsoundat SoundFX("fx2_droptarget",DOFContactors),sw25
  drop25.transz=-25
  d25step = 0
' f25.visible=0
  sw25.timerenabled=1
  sw25.collidable=0
End Sub

Sub sw25_timer
    Select Case d25step
        Case 3:drop25.transz=-24
        Case 4:drop25.transz=-36
    Case 5:drop25.transz=-43:sw25.TimerEnabled = 0
    End Select
    d25Step = d25Step + 1
End Sub

Sub sw26_Hit: dtl.Hit 2:playsoundat SoundFX("fx2_droptarget",DOFContactors),sw26
  drop26.transz=-12
  d26step = 0
' f26.visible=0
  sw26.timerenabled=1
  sw26.collidable=0
End Sub

Sub sw26_timer
    Select Case d26step
        Case 3:drop26.transz=-24
        Case 4:drop26.transz=-36
    Case 5:drop26.transz=-43:sw26.TimerEnabled = 0
    End Select
    d26Step = d26Step + 1
End Sub

Sub sw27_Hit: dtL.Hit 3:playsoundat SoundFX("fx2_droptarget2",DOFContactors),sw27
  drop27.transz=-12
  d27step = 0
' f27.visible=0
  sw27.timerenabled=1
  sw27.collidable=0
End Sub

Sub sw27_timer
    Select Case d27step
        Case 3:drop27.transz=-24
        Case 4:drop27.transz=-36
    Case 5:drop27.transz=-43:sw27.TimerEnabled = 0
    End Select
    d27Step = d27Step + 1
End Sub

Sub sw28_Hit: dtl.Hit 4:playsoundat SoundFX("fx2_droptarget",DOFContactors),sw28
  drop28.transz=-12
  d28step = 0
' f28.visible=0
  sw28.timerenabled=1
  sw28.collidable=0
End Sub

Sub sw28_timer
    Select Case d28step
        Case 3:drop28.transz=-24
        Case 4:drop28.transz=-36
    Case 5:drop28.transz=-43:sw28.TimerEnabled = 0
    End Select
    d28Step = d28Step + 1
End Sub

Sub sw29_Hit: dtL.Hit 5:playsoundat SoundFX("fx2_droptarget",DOFContactors),sw29
  drop29.transz=-12
  d29step = 0
' f29.visible=0
  sw29.timerenabled=1
  sw29.collidable=0
End Sub

Sub sw29_timer
    Select Case d29step
        Case 3:drop29.transz=-24
        Case 4:drop29.transz=-36
    Case 5:drop29.transz=-43:sw29.TimerEnabled = 0
    End Select
    d29Step = d29Step + 1
End Sub

Sub SolLeftTargetReset(enabled)
  if enabled then
    dtL.SolDropUp enabled
    LDrop.enabled = True
  end if
End Sub

Sub LDrop_timer()
  dim xx
  For each xx in DTLeftLights: xx.visible=1:Next
  sw25.collidable=1:sw26.collidable=1:sw27.collidable=1:sw28.collidable=1:sw29.collidable=1
  drop25.transz=0:drop26.transz=0:drop27.transz=0:drop28.transz=0:drop29.transz=0
  me.enabled = false
End Sub

'***********************************
' Bumpers
'***********************************

dim bump1step, bump2step, bump3step

Sub sw38_Hit
  vpmTimer.PulseSw 38
  playsoundAtVol SoundFX("fx2_bumper1",DOFContactors), sw38, VolBump
  layer1bumpskirt1.RotY=-3
  Bump1Step = 0
  sw38.timerenabled = 1
End Sub

Sub sw38_timer
  Select Case Bump1Step
    Case 3: layer1bumpskirt1.RotY=-1
    Case 4: layer1bumpskirt1.RotY=1
    Case 5: layer1bumpskirt1.RotY=-1
    Case 6: layer1bumpskirt1.RotY=1
    Case 7: layer1bumpskirt1.RotY=0:sw38.timerenabled = 0
  End Select
  Bump1Step=Bump1Step + 1
End Sub

Sub sw39_Hit
  vpmTimer.PulseSw 39
  playsoundAtVol SoundFX("fx2_bumper2",DOFContactors), sw39, VolBump
  layer1bumpskirt2.RotY=-5
  bump2step = 0
  sw39.timerenabled = 1
End Sub

Sub sw39_timer
  Select Case bump2step
    Case 3: layer1bumpskirt2.RotY=-3
    Case 4: layer1bumpskirt2.RotY=2
    Case 5: layer1bumpskirt2.RotY=-1
    Case 6: layer1bumpskirt2.RotY=1
    Case 7: layer1bumpskirt2.RotY=0:sw39.timerenabled = 0
  End Select
  bump2step=bump2step + 1
End Sub

Sub sw40_Hit
  vpmTimer.PulseSw 40
  playsoundAtVol SoundFX("fx2_bumper1",DOFContactors), sw40, VolBump
  layer1bumpskirt3.RotY=-5
  bump3step = 0
  sw40.timerenabled = 1
End Sub

Sub sw40_timer
  Select Case bump3step
    Case 3: layer1bumpskirt3.RotY=-3
    Case 4: layer1bumpskirt3.RotY=2
    Case 5: layer1bumpskirt3.RotY=-1
    Case 6: layer1bumpskirt3.RotY=1
    Case 7: layer1bumpskirt3.RotY=0:sw40.timerenabled = 0
  End Select
  bump3step=bump3step + 1
End Sub

'Rollovers
Sub SW14_Hit:Controller.Switch(14)=1:b14.transz=-12: End Sub
Sub SW14_unHit:Controller.Switch(14)=0:b14.transz=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1:w23.transz=-12: End Sub
Sub SW23_unHit:Controller.Switch(23)=0:w23.transz=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1:w24.transz=-12: End Sub
Sub SW24_unHit:Controller.Switch(24)=0:w24.transz=0:End Sub
Sub SW30a_Hit:Controller.Switch(30)=1:b30a.transz=-12: End Sub
Sub SW30a_unHit:Controller.Switch(30)=0:b30a.transz=0:End Sub
Sub SW30b_Hit:Controller.Switch(30)=1:b30b.transz=-12: End Sub
Sub SW30b_unHit:Controller.Switch(30)=0:b30b.transz=0:End Sub
Sub SW30c_Hit:Controller.Switch(30)=1:b30c.transz=-12: End Sub
Sub SW30c_unHit:Controller.Switch(30)=0:b30c.transz=0:End Sub
Sub SW30d_Hit:Controller.Switch(30)=1:b30d.transz=-12: End Sub
Sub SW30d_unHit:Controller.Switch(30)=0:b30d.transz=0:End Sub
Sub SW30e_Hit:Controller.Switch(30)=1:b30e.transz=-12: End Sub
Sub SW30e_unHit:Controller.Switch(30)=0:b30e.transz=0:End Sub
Sub SW31_Hit:Controller.Switch(31)=1:w31.transz=-12: End Sub
Sub SW31_unHit:Controller.Switch(31)=0:w31.transz=0:End Sub
Sub SW32_Hit:Controller.Switch(32)=1:w32.transz=-12:DOF 105, DOFOn:End Sub
Sub SW32_unHit:Controller.Switch(32)=0:w32.transz=0:DOF 105, DOFOff:End Sub

'Wire Triggers
Sub SW1_Hit:Controller.Switch(1)=1 : End Sub
Sub SW1_unHit:Controller.Switch(1)=0:End Sub
Sub SW2_Hit:Controller.Switch(2)=1 : End Sub
Sub SW2_unHit:Controller.Switch(2)=0:End Sub
Sub SW3_Hit:Controller.Switch(3)=1 : End Sub
Sub SW3_unHit:Controller.Switch(3)=0:End Sub
Sub SW4_Hit:Controller.Switch(4)=1 : End Sub
Sub SW4_unHit:Controller.Switch(4)=0:End Sub

'Leaf Switches
Sub SW21_Hit:vpmTimer.PulseSw (21) : End Sub
Sub SW22b_Hit:vpmTimer.PulseSw (22) : End Sub

'Spinners
Sub sw30s_Spin : vpmTimer.PulseSw (30) :PlaySoundAtVol "fx2_spinner", sw30s, VolSpin: End Sub

'***********Rotate Spinner
Dim Angle

Sub SpinnerTimer_Timer
  DOF 103, DOFPulse
  Angle = (sin (sw30s.CurrentAngle-180))
    SpinnerRod.TransZ = -sin( (sw30s.CurrentAngle+180) * (2*3.14/360)) * 5
    SpinnerRod.TransX = (sin( (sw30s.CurrentAngle- 90) * (2*3.14/360)) * -5)
End Sub

'Targets
Dim Step5, Step18, Step22, Step23, Step32

Sub sw5_Hit
  vpmTimer.PulseSw (5)
  sw5target.TransX=-5
  Timer5.enabled = 1
  Step5 = 0
end sub

Sub Timer5_timer
  Select Case Step5
    Case 3: sw5target.TransX=3
    Case 4: sw5target.TransX=-4
    Case 5: sw5target.TransX=2
    Case 6: sw5target.TransX=-3
    Case 7: sw5target.TransX=1
    Case 8: sw5target.TransX=-2
    Case 9: sw5target.TransX=0
    Case 10: sw5target.TransX=-1: Timer5.enabled = 0
  End Select
  Step5=Step5 + 1
End Sub

Sub sw18_Hit
  vpmTimer.PulseSw (18)
  sw18target.TransX=-5
  Timer18.enabled = 1
  Step18 = 0
end sub

Sub Timer18_timer
  Select Case Step18
    Case 3: sw18target.TransX=3
    Case 4: sw18target.TransX=-4
    Case 5: sw18target.TransX=2
    Case 6: sw18target.TransX=-3
    Case 7: sw18target.TransX=1
    Case 8: sw18target.TransX=-2
    Case 9: sw18target.TransX=0
    Case 10: sw18target.TransX=-1: Timer18.enabled = 0
  End Select
  Step18=Step18 + 1
End Sub

Sub sw30ta_Hit
  DOF 101, DOFPulse
  vpmTimer.PulseSw (30)
  t30ta.Transy=-5
  Timer22.enabled = 1
  Step22 = 0
end sub

Sub Timer22_timer
  Select Case Step22
    Case 3: t30ta.transy=3
    Case 4: t30ta.transy=-4
    Case 5: t30ta.transy=2
    Case 6: t30ta.transy=-3
    Case 7: t30ta.transy=1
    Case 8: t30ta.transy=-2
    Case 9: t30ta.transy=0
    Case 10: t30ta.transy=-1: Timer22.enabled = 0
  End Select
  Step22=Step22 + 1
End Sub

Sub sw30tb_Hit
  DOF 102, DOFPulse
  vpmTimer.PulseSw (30)
  t30tb.transy=-5
  Timer23.enabled = 1
  Step23 = 0
end sub

Sub Timer23_timer
  Select Case Step23
    Case 3: t30tb.transy=3
    Case 4: t30tb.transy=-4
    Case 5: t30tb.transy=2
    Case 6: t30tb.transy=-3
    Case 7: t30tb.transy=1
    Case 8: t30tb.transy=-2
    Case 9: t30tb.transy=0
    Case 10: t30tb.transy=-1: Timer23.enabled = 0
  End Select
  Step23=Step23 + 1
End Sub

Sub sw32a_Hit
  DOF 104, DOFPulse
  vpmTimer.PulseSw (32)
  t32a.transy=-5
  Timer32.enabled = 1
  Step32 = 0
end sub

Sub Timer32_timer
  Select Case Step32
    Case 3: t32a.transy=3
    Case 4: t32a.transy=-4
    Case 5: t32a.transy=2
    Case 6: t32a.transy=-3
    Case 7: t32a.transy=1
    Case 8: t32a.transy=-2
    Case 9: t32a.transy=0
    Case 10: t32a.transy=-1: Timer32.enabled = 0
  End Select
  Step32=Step32 + 1
End Sub

Sub SolKnocker(Enabled)
  If Enabled Then PlaySound SoundFX("Fx2_Knocker",DOFKnocker)
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, AStep, BStep, LWall1Step, RWall1Step, LWall2Step

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx2_slingshot1",DOFContactors), sling1
    vpmtimer.PulseSw(36)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -31
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing3.Visible = 1:sling1.TransZ = 0
        Case 5:RSLing3.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx2_slingshot2",DOFContactors),sling2
  vpmtimer.pulsesw(37)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -32
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing3.Visible = 1:sling2.TransZ = 0
        Case 5:LSLing3.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub sw21a_hit
  vpmTimer.PulseSw (21)
    larubber.Visible = 0
    larubber1.Visible = 1
    lWall1Step = 0
    lWall1.enabled=true
End Sub

Sub lWall1_Timer
    Select Case lWall1Step
        Case 3:larubber1.Visible = 0:larubber2.Visible = 1
        Case 4:larubber2.Visible = 0:larubber3.Visible = 1
        Case 5:larubber3.Visible = 0:larubber.Visible = 1:lWall1.enabled=false
    End Select
    lWall1Step = lWall1Step + 1
End Sub

Sub sw21b_hit
  vpmTimer.PulseSw (21)
    rarubber.Visible = 0
    rarubber1.Visible = 1
    rWall1Step = 0
    rWall1.enabled=true
End Sub

Sub rWall1_Timer
    Select Case RWall1Step
        Case 3:rarubber1.Visible = 0:rarubber2.Visible = 1
        Case 4:rarubber2.Visible = 0:rarubber3.Visible = 1
        Case 5:rarubber3.Visible = 0:rarubber.Visible = 1:RWall1.enabled=false
    End Select
    rWall1Step = rWall1Step + 1
End Sub

Sub sw22a_hit
  vpmTimer.PulseSw (22)
    lbrubber.Visible = 0
    lbrubber1.Visible = 1
    lWall2Step = 0
    lWall2.enabled=true
End Sub

Sub lWall2_Timer
    Select Case lWall2Step
        Case 3:lbrubber1.Visible = 0:lbrubber2.Visible = 1
        Case 4:lbrubber2.Visible = 0:lbrubber3.Visible = 1
        Case 5:lbrubber3.Visible = 0:lbrubber.Visible = 1:lWall2.enabled=false
    End Select
    lWall2Step = lWall2Step + 1
End Sub

'Chimes

Sub solchime1(Enable)
    If Enable then
    PlaySound SoundFX("bell10",DOFChimes), 1, 1
  End If
End Sub

Sub solchime10(Enable)
    If Enable then
    PlaySound SoundFX("sj_chime_10a",DOFChimes), 1, 1
  End If
End Sub

Sub solchime100(Enable)
    If Enable then
    PlaySound SoundFX("sj_chime_100a",DOFChimes), 1, 1
  End If
End Sub

Sub solchime1000(Enable)
    If Enable then
    PlaySound SoundFX("sj_chime_1000a",DOFChimes), 1, 1
  End If
End Sub

'-------------------------------------
' Map lights into array
' Set unmapped lamps to Nothing
'-------------------------------------

Set Lights(1)=L001
Set Lights(2)=L002
Set Lights(3)=L003
Set Lights(4)=L004
  Lights(5)=Array(L005,L005a)
  Lights(9)=Array(L009,L009a)
  Lights(10)=Array(L010,L010a)
  Lights(11)=Array(L011,L011a)
  Lights(12)=Array(l012,l012a,l012b,L012c,L012d,L012e)
Set Lights(16)=L016
Set Lights(17)=L017
Set Lights(18)=L018
Set Lights(19)=L019
  Lights(20)=Array(L020,L020a)
  Lights(21)=Array(L021,L021a)
  Lights(25)=Array(L025,L025a)
Set Lights(33)=L033
Set Lights(34)=L034
Set Lights(36)=L036
Set Lights(37)=L037
  Lights(38)=Array(L038,L038a)
  Lights(41)=Array(L041,L041a)
Set Lights(44)=L044
Set Lights(48)=L048
Set Lights(49)=L049
Set Lights(50)=L050
  Lights(52)=Array(L052,L052a,L052b,L052c)
  Lights(53)=Array(L053,L053a)
Set Lights(55)=L055
  Lights(57)=Array(L057,L057a)
Set Lights(58)=L058
Set Lights(64)=L064

Set vpmShowDips = GetRef("editDips")

'*****************************************
'     FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
' FlipperLSh1.RotZ = LeftFlipper1.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  If l005.state=1 Then
    b14.image="b14on"
  Else
    b14.image="b14off"
  End If
  If l036.state=1 Then
    bumpbase12.image="bumpbase"
    bumpbase3.image="bumpbaseoff"
    bumpcap12.image="bumpcap"
    bumpcap3.image="bumpcapoff"
    layer1bumpskirt1.image="bumpskirts"
    layer1bumpskirt2.image="bumpskirts"
    layer1bumpskirt3.image="bumpskirtsoff"
  Else
    bumpbase12.image="bumpbaseoff"
    bumpbase3.image="bumpbase"
    bumpcap12.image="bumpcapoff"
    bumpcap3.image="bumpcap"
    layer1bumpskirt1.image="bumpskirtsoff"
    layer1bumpskirt2.image="bumpskirtsoff"
    layer1bumpskirt3.image="bumpskirts"
  End If
End Sub

'*****************************************
'     BALL SHADOW
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

'******************************
' Diverse Collection Hit Sounds
'******************************

'Sub aBumpers_Hit (idx): PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, pan(ActiveBall): End Sub
Sub aRollovers_Hit(idx):PlaySound "fx2_sensor", 0, Vol(ActiveBall)*VolRol, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx2_MetalHit2", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx2_rubber", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aTargets_Hit(idx):PlaySound "fx2_target", 0, Vol(ActiveBall)*VolTarg, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx2_Woodhit3", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'**********************************************************************************************************
'Stern Stingray
'added by Inkochnito

Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Stingray - DIP switches"
    .AddFrame 2,0,190,"Maximum credits",&H00070000,Array("10 credits",&H00010000,"20 credits",&H00030000,"30 credits",&H00050000,"40 credits",&H00070000)'dip 17&18&19
    .AddFrame 2,76,190,"High game to date award",&H00004000,Array("3 balls - 500K, 5 balls - 840K",0,"3 free games",&H00004000)'dip 15
    .AddFrame 2,122,190,"High score award",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
    .AddFrame 2,168,190,"Balls per game",&H00000040,Array("3 balls",0,"5 balls",&H00000040)'dip 7
    .AddChk 2,220,140,Array("Credits displayed",&H00080000)'dip 20
    .AddChk 2,235,120,Array("Match feature",&H00100000)'dip 21
    .AddChk 2,250,120,Array("4 player games",&H01000000)'dip 25
    .AddChk 2,265,200,Array("Extra Balls (custom ROM only)",&H02000000)'dip 26
    .AddFrame 205,0,190,"Special award",&HC0000000,Array("100.000 points",0,"free ball",&H40000000,"free game",&H80000000,"free ball and free game",&HC0000000)'dip 31&32
    .AddFrame 205,76,190,"Feature lites adjustment",32768,Array("all lites on",0,"lites alternating",32768)'dip 16
    .AddFrame 205,122,190,"Drop target special",&H00800000,Array("special lite on only",0,"special lite on and 1 replay",&H00800000)'dip 24
    .AddFrame 205,168,190,"Melody option",&H00000080,Array("2 tones only",0,"full melody",&H00000080)'dip 8
    .AddLabel 50,280,340,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub

Set vpmShowDips = GetRef("editDips")

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

ReDim ArchRolling(tnob)
InitArchRolling

Dim ArchHit

Sub phys_metals001_Hit
  Archhit = 1
End Sub

Sub NotOnArch_Hit
  ArchHit = 0
End Sub

Sub NotOnArch_unHit
  ArchHit = 0
End Sub

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub InitArchRolling
  Dim i
  For i = 0 to tnob
    ArchRolling(i) = False
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

    If BallVel(BOT(b) ) > 1 AND ArchHit =1 Then
      ArchRolling(b) = True
      PlaySound("ArchHitA" & b),   0, (BallVel(BOT(b))/15)^5 * 1, AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
      PlaySound("ArchRollA" & b), -1, (BallVel(BOT(b))/30)^5 * 1, AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    Else
      If ArchRolling(b) = True Then
      StopSound("ArchRollA" & b)
      ArchRolling(b) = False
      End If
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
  dim a : a = Array(LF, RF)
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
        'playsound "fx_knocker"
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

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
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
End Sub

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball: End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

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

    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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

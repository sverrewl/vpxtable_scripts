' Freedom  Bally 1976
' Allknowing2012 & Hauntfreaks
' April 2016

Option Explicit
Randomize

' 1.00 - Original Release
'
' Thanks to prior VP authors for lamp info
' Switches at http://www.pinitech.com/switch_database.php?name=Bally_Freedom
' Thanks to FP author for original Plastics
' Hauntfreaks for the rework of pf and plastic images
'
' Bug Spinner Lights are not scripted correctly
'

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
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
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Dim DesktopMode: DesktopMode = Table1.ShowDT

If Table1.ShowDT = False then ' hide the backglass lights when in FS mode
  B1.visible=False
  B2.visible=False
  B3.visible=False
  B4.visible=False
  P1.visible=False:P2.visible=False:P3.visible=False:P4.visible=False
End If

Const cGameName="freedom"   ' rom
Const UseSolenoids=1,UseLamps=True,UseGI=0,UseSync=1,SSolenoidOn="SolOn",SSolenoidOff="Soloff",SFlipperOn="FlipperUpLeft",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits="Freedom (Bally 1976)"
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01560000","Bally.vbs", 3.36


SolCallback(1)="vpmSolSound SoundFX(""Chime10"",DOFChimes),"
SolCallback(2)="vpmSolSound SoundFX(""Chime100"",DOFChimes),"
SolCallback(3)="vpmSolSound SoundFX(""Chime1000"",DOFChimes),"
SolCallback(4)="vpmSolSound SoundFX(""ChimeExtra"",DOFChimes),"
SolCallback(6)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)="SolBallRelease"
SolCallback(8)="bsSaucer.SolOut"
' bumper 1,2,3 9/10/12
SolCallback(13)="dtL_SolDropUp"
SolCallback(14)="bsSaucer2.SolOut"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Dim dtL, bstrough, bsSaucer, bsSaucer2

Sub SolLFlipper(Enabled)
  If Enabled Then
       PlaySoundAtVol SoundFX("FlipperUpLeft",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
       PlaySoundAtVol SoundFX("FlipperDown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
       PlaySoundAtVol SoundFX("FlipperUpRight",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd
  Else
       PlaySoundAtVol SoundFX("FlipperDown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
  End If
End Sub

Sub dtL_SolDropUp(enabled)
  if enabled Then
    dtLTimer.interval=500
    dtLTimer.enabled=True
  end if
End Sub

Sub dtLTimer_Timer()
    dtLTimer.enabled=False
    dtL.DropSol_On
End Sub

SolCallback(19)="GameOn"

Sub GameOn(enabled)
  vpmNudge.SolGameOn(enabled)
 If Enabled Then
    GIOn
  Else
    GIOff
  End If
End Sub


Sub solballrelease(enabled)
    bstrough.solexit ssolenoidon, ssolenoidon,enabled:playsound SoundFX("BallRelease",DOFContactors) ' TODO
End sub

Sub Table1_Init
  On Error Resume Next
  With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine=cCredits
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .Hidden=True
    .ShowTitle=0
  End With
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=swTilt
  vpmNudge.Sensitivity=3
  vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3,SW_Wall1,SW_Wall2,SW_Wall3,SW_Wall4,SW_Wall5)
  vpmNudge.SolGameOn(True)

    set bstrough= new cvpmballstack
    bstrough.initnotrough ballrelease,8,80,5


  Set bsSaucer=New cvpmBallStack
  bsSaucer.InitSaucer Kicker1,24,130+Int(Rnd*10),3
  bsSaucer.InitEntrySnd "popper_ball","solon"
  bsSaucer.InitExitSnd SoundFX("popper_ball",DOFContactors),SoundFX("popper_ball",DOFContactors)

  Set bsSaucer2=New cvpmBallStack
  bsSaucer2.InitSaucer Kicker2,24,270+Int(Rnd*10),3
  bsSaucer2.InitEntrySnd "popper_ball","solon"
  bsSaucer2.InitExitSnd SoundFX("popper_ball",DOFContactors),SoundFX("popper_ball",DOFContactors)

  Set dtL=New cvpmDropTarget
  dtL.InitDrop Array(target1, target2, target3, target4, target5),Array(1,2,3,4,5)
  dtL.InitSnd SoundFX("DROP_LEFT",DOFContactors),SoundFX("DTReset",DOFContactors)

    vpmMapLights InsertLights
End Sub

Sub Table1_Exit
  If B2SOn then Controller.Stop
End Sub

Sub Table1_KeyDown(ByVal keycode)
  If KeyCode=PlungerKey Then Plunger.Pullback:PlaySoundAtVol "PlungerPull", Plunger, 1
if keycode =82 then vpmtimer.PulseSw 24
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If KeyCode=PlungerKey Then Plunger.Fire:PlaySoundAtVol "Plunger", Plunger, 1
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub SpinnerLeft_spin():PlaySoundAtVol "fx_spinner", SpinnerLeft, VolSpin:vpmtimer.pulsesw 23:DOF 106, DOFPulse:End Sub
Sub SpinnerRight_spin():PlaySoundAtVol "fx_spinner", SpinnerRight, VolSpin:vpmtimer.pulsesw 23:DOF 107, DOFPulse:End Sub

'Targets
Sub Target1_Hit: dtL.Hit 1:End Sub
Sub Target2_Hit: dtL.Hit 2:End Sub
Sub Target3_Hit: dtL.Hit 3:End Sub
Sub Target4_Hit: dtL.Hit 4:End Sub
Sub Target5_Hit: dtL.Hit 5:End Sub

'Circle Targets
Sub sw22_Hit:vpmTimer.PulseSw 22:DOF 101, DOFPulse:End Sub


' Rubber Walls
Sub sw_wall1_Hit
  vpmTimer.PulseSw 20
  DOF 108, DOFPulse
  Rubber10.visible=False
  Rubber10A.visible=True
  sw_wall1.timerinterval=100
  sw_wall1.timerenabled=True
End Sub

Sub sw_wall1_timer()
  Rubber10.visible=True
  Rubber10A.visible=False
  sw_wall1.timerenabled=False
End Sub

Sub sw_wall2_Hit
  vpmTimer.PulseSw 20
  Rubber7.visible=False
  Rubber7A.visible=True
  sw_wall2.timerinterval=100
  sw_wall2.timerenabled=True
End Sub

Sub sw_wall2_timer()
  Rubber7.visible=True
  Rubber7A.visible=False
  sw_wall2.timerenabled=False
End Sub

Sub sw_wall3_Hit   ' behind targets - most targets down.
  vpmTimer.PulseSw 19
  SideRubber0A.visible=False
  SideRubber1A.visible=True
  sling3.TransZ = -20
  sw_wall3.timerinterval=100
  sw_wall3.timerenabled=True
End Sub

Sub sw_wall3_timer()
  SideRubber0A.visible=True
  SideRubber1A.visible=False
  sling3.TransZ = 0
  sw_wall3.timerenabled=False
End Sub

Sub sw_wall4_Hit
  vpmTimer.PulseSw 20
  DOF 110, DOFPulse
  Rubber3.visible=False
  Rubber3A.visible=True
  sw_wall4.timerinterval=100
  sw_wall4.timerenabled=True
End Sub

Sub sw_wall4_timer()
  Rubber3.visible=True
  Rubber3A.visible=False
  sw_wall4.timerenabled=False
End Sub

Sub sw_wall5_Hit
  vpmTimer.PulseSw 20
  DOF 109, DOFPulse
  Rubber1.visible=False
  Rubber1A.visible=True
  sw_wall5.timerinterval=100
  sw_wall5.timerenabled=True
End Sub

Sub sw_wall5_timer()
  Rubber1.visible=True
  Rubber1A.visible=False
  sw_wall5.timerenabled=False
End Sub

' Rollover Switches
Sub sw18a_Hit:Controller.Switch(18)=1:DOF 102, DOFOn:End Sub
Sub sw18a_unHit:Controller.Switch(18)=0:DOF 102, DOFOff:End Sub
Sub sw18b_Hit:Controller.Switch(18)=1:DOF 105, DOFOn:End Sub
Sub sw18b_unHit:Controller.Switch(18)=0:DOF 105, DOFOff:End Sub
Sub sw22a_Hit:Controller.Switch(22)=1:DOF 103, DOFOn:End Sub
Sub sw22a_unHit:Controller.Switch(22)=0:DOF 103, DOFOff:End Sub
Sub sw22b_Hit:Controller.Switch(22)=1:DOF 104, DOFOn:End Sub
Sub sw22b_unHit:Controller.Switch(22)=0:DOF 104, DOFOff:End Sub

Sub StarTrigger_Hit:Controller.Switch(21)=1:End Sub
Sub StarTrigger_unHit:Controller.Switch(21)=0:End Sub

Sub Drain_Hit():bsTrough.AddBall Me::End Sub

Sub Kicker1_hit():bsSaucer.AddBall 0:End Sub
Sub Kicker1_unhit():DOF 111, DOFPulse:End Sub
Sub Kicker2_hit():bsSaucer2.AddBall 0:End Sub
Sub Kicker2_unhit():DOF 112, DOFPulse:End Sub

Sub GIOn
  dim bulb
  for each bulb in GILights
  bulb.state = LightStateOn
  next
End Sub

Sub GIOff
  dim bulb
  for each bulb in GILights
  bulb.state = LightStateOff
  next
End Sub

Sub FlipperTimer_Timer()
   GateLP.RotZ = ABS(GateL.currentangle)
   GateRP.RotZ = ABS(GateR.currentangle)
   GatePlunger.RotX = ABS(Gate1.currentangle)+160
End Sub

Dim bump1,bump2,bump3

Sub Bumper1_Hit:vpmTimer.PulseSw 29:bump1 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper1_Timer()
  Select Case bump1
        Case 1:Ring1.Z = -30:bump1 = 2
        Case 2:Ring1.Z = -20:bump1 = 3
        Case 3:Ring1.Z = -10:bump1 = 4
        Case 4:Ring1.Z = 0:Me.TimerEnabled = 0
  End Select
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 30:bump2 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper2_Timer()
  Select Case bump2
        Case 1:Ring2.Z = -30:bump2 = 2
        Case 2:Ring2.Z = -20:bump2 = 3
        Case 3:Ring2.Z = -10:bump2 = 4
        Case 4:Ring2.Z = 0:Me.TimerEnabled = 0
  End Select
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 31:bump3 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper3_Timer()
  Select Case bump3
        Case 1:Ring3.Z = -30:bump3 = 2
        Case 2:Ring3.Z = -20:bump3 = 3
        Case 3:Ring3.Z = -10:bump3 = 4
        Case 4:Ring3.Z = 0:Me.TimerEnabled = 0
  End Select
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
 '  debug.print "rubber"
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  '  debug.print "Posts"
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Bumpers_Hit(idx)
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySound SoundFX("fx_bumper2",DOFContactors), 0, Vol(ActiveBall)*VolBump, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound SoundFX("fx_bumper2",DOFContactors), 0, Vol(ActiveBall)*VolBump, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound SoundFX("fx_bumper3",DOFContactors), 0, Vol(ActiveBall)*VolBump, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 4 : PlaySound SoundFX("fx_bumper4",DOFContactors), 0, Vol(ActiveBall)*VolBump, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Tstep

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("slingshotright",DOFContactors), sling1, 1
    vpmTimer.PulseSw 28
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("slingshotleft",DOFContactors), sling2, 1
    vpmTimer.PulseSw 32
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'Digital LED Display

Dim Digits(28)
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
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)

Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
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
end if
End Sub

'Bally Freedom
'by Inkochnito
Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,400,"Freedom - DIP switches"
    .AddFrame 2,2,190,"Maximum credits",&H00070000,Array("5 credits",0,"10 credits",&H00010000,"15 credits",&H00020000,"20 credits",&H00030000,"25 credits",&H00040000,"30 credits",&H00050000,"35 credits",&H00060000,"40 credits",&H00070000)'dip 17&18&19
    .AddFrame 2,142,190,"High game to date award",&H00004000,Array("no award",0,"3 credits",&H00004000)'dip 15
    .AddFrame 2,188,190,"High score feature",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
    .AddFrame 205,2,190,"Novelty Option",&H00200000,Array("off (proceed with next 2 steps)",0,"on (next steps not necessary)",&H00200000)'dip 22
    .AddFrame 205,48,190,"Replay Option",&H00400000,Array("extra ball",0,"replay",&H00400000)'dip 23
    .AddFrame 205,96,190,"Award Option",&H00800000,Array("conservative",0,"liberal",&H00800000)'dip 24
    .AddFrame 205,142,190,"Balls per game",32768,Array("3 balls",0,"5 balls",32768)'dip 16
    .AddFrame 205,188,190,"Freedom Wheel Arrow",&H00000040,Array("advance with points scoring",0,"advance continuously",&H00000040)'dip 7
    .AddChk 2,238,120,Array("Match feature",&H00100000)'dip 21
    .AddChk 130,238,120,Array("Credits display",&H00080000)'dip 20
    .AddChk 270,238,120,Array("Melody option",&H00000080)'dip 8
    .AddLabel 50,260,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")


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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Const tnob = 1 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    RollingTimer.interval=100
    RollingTimer.enabled=True
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub



' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


'Game Plan Sharpshooter, 1979
'VP9 DT Table by Bodydump credits:
  'Thanks to Eala for his original vp8 table and his object table which I borrowed from
  'Thanks to JimmyFingers for physics tweaks and BMPR modding and improved sound/sound routines.
  'Thanks to Rob046, Uncle Willy and Destruk for help, question answering and opinions.
'VPX conversion by HSM
Option Explicit
Randomize

' Thalamus 2018-07-24
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
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolWood   = 1    ' Woods volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
Const SCoin= "fx_coin"
Const FlippersAlwaysOn = 0 '1 for testing
LoadVPM "00990300", "GamePlan.vbs", 3.1
Dim DesktopMode: DesktopMode = Sharpshooter.ShowDT
Const cGameName="sshooter" ,UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="fx_solenoid",SSolenoidOff="fx_solenoidoff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const cCredits="Sharpshooter, GamePlan, 1979"
Const sBallRelease=8
Const sBumper1=14
Const sBumper2=12
Const sBumper3=13
Const sBumper4=10
Const sKicker=15
Const sDropA=11
Const sUSling=9
Const sEnable=16
Const sChimeD=18
Const sChimeC=19
Const sChimeB=20
Const sChimeA=21
Dim bsTrough,dtDrop,bsSaucer
Dim bump1,bump2,bump3,bump4
Dim x
SolCallback(sBallRelease)="bsTrough.solOut"
SolCallback(sKicker)="bsSaucer.SolOut"
SolCallback(sDropA)="dtDrop.SolDropUp"
SolCallback(sEnable)="vpmNudge.SolGameOn"
SolCallback(3)="vpmSolSound ""fx_Knocker"","
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Dim VarHidden, aReels
If Sharpshooter.ShowDT = true then
    VarHidden = 1
    For each x in backdropstuff
        x.Visible = 1
    Next
else
    VarHidden = 0
    For each x in backdropstuff
        x.Visible = 0
    Next
    lrail.Visible = 0
    rrail.Visible = 0
end if

Sub Sharpshooter_Init
    ' Thalamus - was missing vpminit
    vpmInit Me

     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Sharp Shooter - Game Plan 1979"&chr(13)&"VPX conversion by HSM"
         .HandleKeyboard = 0
         .ShowTitle = 1
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
         .Hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

  Controller.Dip(0) = (0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '01-08
  Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 1*128) '09-16
  Controller.Dip(2) = (0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '17-24
  Controller.Dip(3) = (1*1 + 1*2 + 1*4 + 1*8 + 0*16 + 1*32 + 1*64 + 0*128) '25-32

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  vpmNudge.TiltSwitch = swTilt
  vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,Bumper4,LeftSlingshot,RightSlingshot,ShooterSlingshot,Wall34)

  Set dtdrop = New cvpmDropTarget
      With dtdrop
      .InitDrop Array(TargetS,TargetH,TargetO1,TargetO2,TargetT,TargetE,TargetR),Array(31,32,35,36,4,10,17)
         .Initsnd "fx_droptarget", "fx2_DTReset"
       End With

  Set bsSaucer = New cvpmBallStack
             With bsSaucer
              .InitSaucer Kicker1,24, 200, 10
           .KickForceVar = 1
           .KickAngleVar = 1
              .InitExitSnd "fx2_popper_ball", "Solenoid"
             .InitAddSnd "fx2_kicker_enter_left"
             End With

  Set bsTrough = New cvpmBallStack
            With bsTrough
             .InitSw 0,11,0,0,0,0,0,0
             .InitKick BallRelease,90,8
             bsTrough.InitExitSnd "fx_ballrel", "fx_solenoid"
             .Balls = 1
            End With
End Sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' Sub Sharpshooter_KeyDown(ByVal keycode)
'   If keycode = LeftTiltKey Then Nudge 90, 2
'   If keycode = RightTiltKey Then Nudge 270, 2
'   If keycode = CenterTiltKey Then Nudge 0, 2
'
'   If vpmKeyDown(keycode) Then Exit Sub
'   If keycode = PlungerKey Then Plunger.PullBack: PlaySoundAtVol "fx_plungerpull",Plunger,1:   End If
'     If keycode = LeftFlipperKey Then LeftFlipper.RotateToEnd: PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors),LeftFlipper,VolFlip
'     If keycode = RightFlipperKey Then RightFlipper.RotateToEnd: PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors),RightFlipper,VolFlip
' End Sub

' Sub Sharpshooter_KeyUp(ByVal keycode)
'   If keycode = PlungerKey Then Plunger.Fire: PlaySoundAtVol "fx_plunger", Plunger, 1
'     If keycode = LeftFlipperKey Then LeftFlipper.RotateToStart: PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors),LeftFlipper, Volflip
'     If keycode = RightFlipperKey Then RightFlipper.RotateToStart: PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors),RightFlipper,VolFlip
'   If vpmKeyUp(keycode) Then Exit Sub
' End Sub

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors),LeftFlipper,VolFlip
    LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors),LeftFlipper,Volflip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors),rightflipper,Volflip
        RightFlipper.RotateToEnd
   Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors),rightflipper,volflip
        RightFlipper.RotateToStart
    End If
End Sub

Sub SolKnocker(Enabled)
  If Enabled Then PlaySound SoundFX("fx_Knocker",DOFKnocker)
End Sub

  Sub Sharpshooter_KeyDown(ByVal keycode)
  If keycode = LeftTiltKey Then Nudge 90, 2
  If keycode = RightTiltKey Then Nudge 270, 2
  If keycode = CenterTiltKey Then Nudge 0, 2

  If vpmKeyDown(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.PullBack: PlaySoundAtVol "fx_plungerpull",plunger,1:   End If
  If FlippersAlwaysOn =1 Then
    If keycode = LeftFlipperKey Then LeftFlipper.RotateToEnd: PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors),LeftFlipper,VolFlip
    If keycode = RightFlipperKey Then RightFlipper.RotateToEnd: PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), RightFlipper, VolFlip
  End If
End Sub

Sub Sharpshooter_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then Plunger.Fire: PlaySoundAtVol "fx_plunger", Plunger, 1
  If FlippersAlwaysOn =1 Then
    If keycode = LeftFlipperKey Then LeftFlipper.RotateToStart: PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors),LeftFlipper, VolFlip
    If keycode = RightFlipperKey Then RightFlipper.RotateToStart: PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors),RightFlipper,VolFlip
  End If
  If vpmKeyUp(keycode) Then Exit Sub

End Sub
Sub Drain_Hit:PlaysoundAtVol "fx2_drain2",drain,1:bsTrough.AddBall Me:End Sub
Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub
Sub TargetS_Dropped:dtDrop.Hit 1:End Sub
Sub TargetH_Dropped:dtDrop.Hit 2:End Sub
Sub TargetO1_Dropped:dtDrop.Hit 3:End Sub
Sub TargetO2_Dropped:dtDrop.Hit 4:End Sub
Sub TargetT_Dropped:dtDrop.Hit 5:End Sub
Sub TargetE_Dropped:dtDrop.Hit 6:End Sub
Sub TargetR_Dropped:dtDrop.Hit 7:End Sub
Sub RightSlingshot_Hit:vpmTimer.PulseSw 9:End Sub
Sub ShooterSlingshot_Hit:VpmTimer.PulseSw 9:End Sub
Sub Sling2_Hit:vpmTimer.PulseSw 9:End Sub
Sub Wall34_Hit:vpmTimer.PulseSw 9:End Sub
Sub rightwalla_Hit:vpmTimer.PulseSw 9:End Sub
Sub rightwallb_Hit:vpmTimer.PulseSw 9:End Sub
Sub rightangle_Hit:vpmTimer.PulseSw 9:End Sub
Sub Lane3_Hit:Controller.Switch(12)=1:End Sub
Sub Lane3_unHit:Controller.Switch(12)=0:End Sub
Sub Lane5_Hit:Controller.Switch(13)=1:End Sub
Sub Lane5_unHit:Controller.Switch(13)=0:End Sub
Sub Lane6_Hit:Controller.Switch(14)=1:End Sub
Sub Lane6_unHit:Controller.Switch(14)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(16)=1:End Sub
Sub LeftInlane_unHit:Controller.Switch(16)=0:End Sub
Sub LeftOutlane_Hit:Controller.Switch(18)=1:End Sub
Sub LeftOutlane_unHit:Controller.Switch(18)=0:End Sub
Sub T1_Hit:vpmTimer.PulseSw 19:PlaySoundAtVol "FX2_Target",ActiveBall,1:End Sub
Sub T2_Hit:vpmTimer.PulseSw 20:PlaySoundAtVol "FX2_Target",ActiveBall,1:End Sub


Sub Bumper1_Hit:RandomSoundBumper:vpmTimer.PulseSw 34:PlaySoundAtVol "fx_bumper",bumper1,volbump:bump1 = 1:Me.TimerEnabled = 1:End Sub
 Sub Bumper1_Timer()
 if L43.state=0 then
  L43A.state=0
  L43A1.state=0
  L43A2.state=0
  end If
 if L43.state=1 then
  L43A.state=1
  L43A1.state=1
  L43A2.state=1
  end If
  End Sub
Sub Bumper2_Hit:RandomSoundBumper:vpmTimer.PulseSw 33:PlaySoundAtVol "fx_bumper",bumper2,volbump:bump2 = 1:Me.TimerEnabled = 1:End Sub
 Sub Bumper2_Timer()
 if L39.state=0 then
  L39A.state=0
  L39A1.state=0
  L39A2.state=0
  end if
 if L39.state=1 then
  L39A.state=1
  L39A1.state=1
  L39A2.state=1
  end if
End Sub

Sub Bumper3_Hit:RandomSoundBumper:vpmTimer.PulseSw 21:PlaySoundAtVol "fx_bumper",bumper3,volbump:bump3 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper4_Hit:RandomSoundBumper:vpmTimer.PulseSw 22:PlaySoundAtVol "fx_bumper",bumper4,volbump:bump4 = 1:Me.TimerEnabled = 1:End Sub

Sub Spinner_Spin:PlaySoundAtVol "fx_spinner", spinner,volspin:vpmTimer.PulseSw 23:End Sub
Sub Roll1_Hit:Controller.Switch(25)=1:End Sub
Sub Roll1_unHit:Controller.Switch(25)=0:End Sub
Sub Roll2_Hit:Controller.Switch(27)=1:End Sub
Sub Roll2_unHit:Controller.Switch(27)=0:End Sub
Sub Roll3_Hit:Controller.Switch(28)=1:End Sub
Sub Roll3_unHit:Controller.Switch(28)=0:End Sub
Sub Roll4_Hit:Controller.Switch(29)=1:End Sub
Sub Roll4_unHit:Controller.Switch(29)=0:End Sub
Sub Roll5_Hit:Controller.Switch(30)=1:End Sub
Sub Roll5_unHit:Controller.Switch(30)=0:End Sub

Sub Lane1_Hit:Controller.Switch(37)=1:End Sub
Sub Lane1_unHit:Controller.Switch(37)=0:End Sub
Sub Lane2_Hit:Controller.Switch(38)=1:End Sub
Sub Lane2_unHit:Controller.Switch(38)=0:End Sub
Sub Lane4_Hit:Controller.Switch(39)=1:End Sub
Sub Lane4_unHit:Controller.Switch(39)=0:End Sub
Sub Lane7_Hit:Controller.Switch(40)=1:End Sub
Sub Lane7_unHit:Controller.Switch(40)=0:End Sub

  Dim LStep
Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 15
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

Sub LeftFlipper_Collide(parm)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 4 then
    RandomSoundFlipper()
  Else
    RandomSoundFlipperLowVolume()
  End If

End Sub

Sub RightFlipper_Collide(parm)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 4 then
    RandomSoundFlipper()
  Else
    RandomSoundFlipperLowVolume()
  End If
End Sub
Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx2_flip_hit_1" ' TODO
    Case 2 : PlaySound "fx2_flip_hit_2"
    Case 3 : PlaySound "fx2_flip_hit_3"
  End Select
End Sub

Sub RandomSoundFlipperLowVolume()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx2_flip_hit_1_low"
    Case 2 : PlaySound "fx2_flip_hit_2_low"
    Case 3 : PlaySound "fx2_flip_hit_3_low"
  End Select
End Sub

Sub RandomSoundBumper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx2_bumper_1"
    Case 2 : PlaySound "fx2_bumper_2"
    Case 3 : PlaySound "fx2_bumper_3"
  End Select
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************
Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
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

    If VarHidden Then
        Leds_Timer
    End If
    UpdateLamps
    RollingUpdate
End Sub

Sub UpdateLamps
  NFadeL 2, l2  '1000bonus
  NFadeL 3, l3  '2000 bonus
  NFadeL 4, l4  '3000 bonus
  NFadeL 5, l5  '4000 bonus
  NFadeL 6, l6  '5000 bonus
  NFadeL 7, l7  '6000 bonus
  NFadeL 8, l8  '7000 bonus
  NFadeL 9, l9  '8000 bonus
  NFadeL 10, l10  '9000 bonus
    NFadeL 11, l11  '10000 bonus
  NFadeL 12, l12  '2x bonus
  NFadeL 13, l13  '3x bonus
  NFadeL 14, l14  '4x bonus
  NFadeL 15, l15  '5x bonus
  NFadeL 17, l17  'Special upper
  NFadeL 18, l18  'Special lower
  NFadeL 19, l19  'Extra Ball upper
  NFadeLm 19,l19b 'Rollover5
  NFadeL 20, l20  'Extra Ball lower
  NFadeL 21, l21  'S harp
  NFadeL 22, l22  's H arp
  NFadeL 23, l23  'shoote R
  NFadeL 24, l24  'sh A rp
  NFadeL 25, l25  'sha R p
  NFadeL 26, l26  'shar P
  NFadeL 28, l28  'S hooter
  NFadeL 29, l29  'Rollover 2
  NFadeL 30, l30  'Rollover 3
  NFadeL 31, l31  'sho O ter
  NFadeL 32, l32  'Rollover 4
  NFadeL 33, l33  's H ooter
  NFadeL 35, l35  '20000 bonus
  NFadeL 36, l36  'sh O oter
  NFadeL 37, l37  'upper 2x
  NFadeL 39, L39  '1st pair bumpers
    NFadeTm 40, l40, "TILT"           'Tilt Backbox
    NFadeTm 41, l41, "High Score to Date"   'High Score Backbox
  NFadeL 42, l42  'Spinner 1000
  NFadeL 43, L43   '2nd pair bumpers
  NFadeTm 44, l44, "Game Over"        'Game Over Backbox
    NFadeL 45, l45  'shoo T er
  NFadeL 46, l46  'shoot E r
    NFadeL 47, l47  'upper 3x
    NFadeL 48, l48  'upper 5x
    NFadeL 50, l50  'upper 4x
    NFadeL 51, l51  'Rollover 1
    NFadeTm 52, l52b, "Same Player Shoots Again"'Shoot Again Backbox
  NFadeL 52, l52  'Shoot Again
    NFadeTm 53, l53b, "Ball in Play"      'Ball In Play Backbox
    NFadeL 55, l55  '25000
    NFadeT 56, l56, "Match"           'Match Backbox
  NFadeL 57, l57                'Player 1 Backbox
    NFadeL 58, l58                'Player 2 Backbox
    NFadeL 59, l59                'Player 3 Backbox
    NFadeL 60, l60                'Player 4 Backbox
    NFadeT 61, l61, "MILLION"         'Player 1 Million
  NFadeT 62, l62, "MILLION"         'Player 2 Million
    NFadeT 63, l63, "MILLION"         'Player 3 Million
  NFadeT 64, l64, "MILLION"         'Player 4 Million
End Sub

' div lamp subs
Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
        FlashRepeat(x) = 20     ' how many times the flash repeats
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
  UpdateLamps:UpdateLamps:UpdateLamps
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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
    Select Case FadingLevel(nr)
        Case 4, 5
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub FlashBlink(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr) Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr) -1
                If FlashRepeat(nr) Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr) Then FadingLevel(nr) = 4
    End Select
End Sub

' Reels
Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts
Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'************************************
'          LEDs Display
'************************************
Dim Digits(28)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 124   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 103  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 252  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5

Set Digits(6) = b0
Set Digits(7) = b1
Set Digits(8) = b2
Set Digits(9) = b3
Set Digits(10) = b4
Set Digits(11) = b5

Set Digits(12) = c0
Set Digits(13) = c1
Set Digits(14) = c2
Set Digits(15) = c3
Set Digits(16) = c4
Set Digits(17) = c5

Set Digits(18) = d0
Set Digits(19) = d1
Set Digits(20) = d2
Set Digits(21) = d3
Set Digits(22) = d4
Set Digits(23) = d5

Set Digits(24) = e0
Set Digits(25) = e1
Set Digits(26) = e2
Set Digits(27) = e3

Sub Leds_Timer()
     On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
        Next
    End IF
End Sub

Sub Sharpshooter_Paused
Controller.Pause=True
End Sub

Sub Sharpshooter_UnPaused
Controller.Pause=False
End Sub

Sub Sharpshooter_Exit
Controller.Stop
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************
Sub aBumpers_Hit (idx): PlaySound SoundFX("fx_bumper", DOFContactors), 0, VolBump, pan(ActiveBall): End Sub
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall)*VolRol, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Sharpshooter" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Sharpshooter.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Sharpshooter" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Sharpshooter.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Sharpshooter" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Sharpshooter.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Sharpshooter.height-1
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
Const tnob = 20 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
' Sub RollingTimer_Timer()
   Dim BOT, b, ballpitch
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
Sub Sharpshooter_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


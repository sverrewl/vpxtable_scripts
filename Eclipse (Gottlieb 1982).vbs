'************ Eclipse Gottlieb 1982


Option Explicit
Randomize

Const cGameName     = "eclipse"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01000100", "sys80.vbs", 2.31

'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1            '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
Dim ROMSounds: ROMSounds=1      '**********set to 0 for no rom sounds, 1 to play rom sounds.. mostly used for testing
Dim BlackRails: BlackRails=1      '********** set to 0 for normal wood table rails and white rubbers, set 1 one for black.

'************************************************
'************************************************
'************************************************
'************************************************
'************************************************

Const UseSolenoids  = 2
Const UseLamps      = 1
Const UseGI         = 0

' Standard Sounds
Const SSolenoidOn  = ""
Const SSolenoidOff = ""
Const SCoin        = "fx_coin"

Dim xx, objekt
Sub Eclipse_Init()

' Thalamus : Was missing 'vpminit me'
  vpminit me

    On Error Resume Next
      With Controller
        .GameName                               = cGameName
        .SplashInfoLine                         = "Eclipse, Gottlieb 1982"
        .HandleKeyboard                         = 0
        .ShowTitle                              = 0
        .ShowDMDOnly                            = 1
        .ShowFrame                              = 0
        .ShowTitle                              = 0
        .hidden                                 = 1
        .Games(cGameName).Settings.Value("rol") = 0
  .Games(cGameName).Settings.Value("sound") = ROMSounds
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
      End With
    On Error Goto 0

' basic pinmame timer
    PinMAMETimer.Interval   = PinMAMEInterval
    PinMAMETimer.Enabled    = 1

' Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3, Bumper4, LeftSLingShot, RightSlingShot)

  if BlackRails = 1 Then
    outers.image="outersblack_off"
    rubbers.image="rubbersblack_off"
    slinga.image="rubbersblack_off"
    slingb.image="rubbersblack_off"
    slingc.image="rubbersblack_off"
'   for each objekt in SideRails: objekt.image="sidewood_BLKh": objekt.sideimage="sidewood_BLKh": Next
'   for each objekt in a_Rubbers: objekt.image="blacksquare": Next
'   for each objekt in a_Posts: objekt.image="blacksquare": Next
    Else
'   for each objekt in SideRails: objekt.image="sidewood_h": objekt.sideimage="sidewood_h": Next
'   for each objekt in a_Rubbers: objekt.image="": Next
'   for each objekt in a_Posts: objekt.image="": Next
  end if

'Controlled Lights
  vpmMapLights CPULights

'************************Trough

  Trough2.CreateBall
  Trough1.CreateBall
  sw25.CreateBall
  Controller.Switch(25) = 1
  controller.Switch(15) = 0



    if b2son then: for each xx in backdropstuff: xx.visible=0: next
    startGame.enabled=true

  if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if

    if flippershadows=1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
       else
        FlipperLSh.visible=0
        FlipperRSh.visible=0

    end if
    FindDips        'find if match enabled, if so turn back on number to match box

End Sub

Sub Eclipse_Paused:Controller.Pause = 1:End Sub

Sub Eclipse_unPaused:Controller.Pause = 0:End Sub

Sub Eclipse_Exit
    If b2son then controller.stop
End Sub

sub startGame_timer
    dim xx
    playsound "poweron"
    LampTimer.enabled=1
    For each xx in GI:xx.State = 1: Next        '*****GI Lights On
  if blackrails=1 then outers.image="outersblack_on"
  if blackrails=1 then rubbers.image="rubbersblack_on"
  if blackrails=1 then slinga.image="rubbersblack_on"
  if blackrails=1 then slinga1.image="rubbersblack_on"
  if blackrails=1 then slinga2.image="rubbersblack_on"
  if blackrails=1 then slingb.image="rubbersblack_on"
  if blackrails=1 then slingb1.image="rubbersblack_on"
  if blackrails=1 then slingb2.image="rubbersblack_on"
  if blackrails=1 then slingc.image="rubbersblack_on"
  if blackrails=1 then slingc1.image="rubbersblack_on"
  if blackrails=1 then slingc2.image="rubbersblack_on"
  if blackrails=1 then slingd.image="rubbersblack_on"
  if blackrails=1 then slingd1.image="rubbersblack_on"
  if blackrails=1 then slingd2.image="rubbersblack_on"
  if blackrails=1 then slinge1.image="rubbersblack_on"
  if blackrails=1 then slinge2.image="rubbersblack_on"
  if blackrails=1 then lsling.image="rubbersblack_on"
  if blackrails=1 then lsling1.image="rubbersblack_on"
  if blackrails=1 then lsling2.image="rubbersblack_on"
  if blackrails=1 then rsling.image="rubbersblack_on"
  if blackrails=1 then rsling1.image="rubbersblack_on"
  if blackrails=1 then rsling2.image="rubbersblack_on"
  if blackrails=0 then outers.image="outers_on"
  if blackrails=0 then rubbers.image="rubbers_on"
  if blackrails=0 then slinga.image="rubbers_on"
  if blackrails=0 then slinga1.image="rubbers_on"
  if blackrails=0 then slinga2.image="rubbers_on"
  if blackrails=0 then slingb.image="rubbers_on"
  if blackrails=0 then slingb1.image="rubbers_on"
  if blackrails=0 then slingb2.image="rubbers_on"
  if blackrails=0 then slingc.image="rubbers_on"
  if blackrails=0 then slingc1.image="rubbers_on"
  if blackrails=0 then slingc2.image="rubbers_on"
  if blackrails=0 then slingd.image="rubbers_on"
  if blackrails=0 then slingd1.image="rubbers_on"
  if blackrails=0 then slingd2.image="rubbers_on"
  if blackrails=0 then slinge1.image="rubbers_on"
  if blackrails=0 then slinge2.image="rubbers_on"
  if blackrails=0 then lsling.image="rubbers_on"
  if blackrails=0 then lsling1.image="rubbers_on"
  if blackrails=0 then lsling2.image="rubbers_on"
  if blackrails=0 then rsling.image="rubbers_on"
  if blackrails=0 then rsling1.image="rubbers_on"
  if blackrails=0 then rsling2.image="rubbers_on"
  plastic1.image="plastic1on"
  plastic2.image="plastic2on"
  plastic3.image="plastic3on"
  plastic4.image="plastic4on"
  shadowsGIOFF.visible=0
  shadowsGION.visible=1
    me.enabled=false
end sub

Sub Eclipse_KeyDown(ByVal keycode)

    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then
        Plunger.PullBack
        PlaySoundAt "plungerpull", Plunger
    End If
    If keycode=AddCreditKey then playsound "coin": vpmTimer.pulseSW (swCoin1): end if


End Sub

Sub Eclipse_KeyUp(ByVal keycode)

    If keycode = 61 then FindDips

    If keycode = PlungerKey Then
    Plunger.Fire
    if ballhome.ballcntover>0 then
      PlaySoundAt "plungerreleaseball", Plunger 'PLAY WHEN THERE IS A BALL TO HIT
      else
      PlaySoundAt "plungerreleasefree", Plunger 'PLAY WHEN NO BALL TO PLUNGE
    end if
    End If

    If vpmKeyUp(keycode) Then Exit Sub

End Sub



'*****************************************
'Solenoids
'*****************************************
SolCallBack(1)  = "Rraised"
SolCallBack(2)  = "Lraised"
SolCallback(5)  = "Ukicker"
SolCallback(6)  = "Craised"
SolCallback(8)  = "SolKnocker"
SolCallback(9)  = "SolOutHole"


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub Rraised(enabled)
    if enabled then Rreset.enabled=True
End Sub

Sub Rreset_timer
  PlaySoundAtVol "dropbankreset", sw32, 1
    For each objekt in DTright: objekt.isdropped=0: next
'    For each light in DTRightLights: light.state=0: Next
    Rreset.enabled=false
End Sub

Sub Lraised(enabled)
    if enabled then Lreset.enabled=True
End Sub

Sub Lreset_timer
  PlaySoundAtVol "dropbankreset", sw53, 1
     For each objekt in DTleft: objekt.isdropped=0: next
'    For each light in DTLeftLights: light.state=0: Next
    Lreset.enabled=false
End Sub

Sub Craised(enabled)
    if enabled then Creset.enabled=True
End Sub

Sub Creset_timer
  PlaySoundAtVol "dropbankreset", sw4, 1
     For each objekt in DTmid: objekt.isdropped=0: next
'    For each light in DTmidLights: light.state=0: Next
    Creset.enabled=false
End Sub

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), LFLogo, 1
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), LFLogo, 1
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), RfLogo, 1
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), RfLogo, 1
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub


Sub SolKnocker(Enabled)
    If Enabled Then PlaySoundAtVol SoundFX("Knocker",DOFKnocker), Plunger, 1
End Sub

Sub LSaucer
    sw5.kick 180,6
    PlaySoundAtVol SoundFX("popper_ball",DOFContactors), sw5, 1
    sw5.uservalue=1
    sw5.timerenabled=1
  PkickarmSW5.rotz=15
End Sub

Sub sw5_timer
    select case sw5.uservalue
      case 2:
        PkickarmSW5.rotz=0
        me.timerenabled=0
    end Select
    sw5.uservalue=sw5.uservalue+1
End Sub

Sub Ukicker(Enabled)
    If Enabled Then
    KickPlunger.fire
        PlaySoundAtVol SoundFX("popper_ball",DOFContactors), slingU, 1
    slingU.rotx = 20
        sw30.uservalue=1
        sw30.timerenabled=1
  end if
End Sub

Sub sw30_Timer
    Select Case sw30.uservalue
        Case 3:slingU.rotx = 10
        Case 4:slingU.rotx = 0:sw30.TimerEnabled = 0
    End Select
    sw30.uservalue = sw30.uservalue + 1
End Sub

'*****************************************
' Saucer
'*****************************************

Sub SW5_Hit:Controller.Switch(5)=1:PlaySoundAt "fx_hole_enter", sw5:End Sub
Sub SW5_unHit:Controller.Switch(5)=0:End Sub

'*****************************************
'Wire rollover Switches
'*****************************************

Sub SW0_Hit:Controller.Switch(0)=1:End Sub
Sub SW0_unHit:Controller.Switch(0)=0:End Sub
Sub SW10_Hit:Controller.Switch(10)=1:End Sub
Sub SW10_unHit:Controller.Switch(10)=0:End Sub
Sub SW20_Hit:Controller.Switch(20)=1:End Sub
Sub SW20_unHit:Controller.Switch(20)=0:End Sub

Sub SW30_Hit:Controller.Switch(30)=1:End Sub
Sub SW30_unHit:Controller.Switch(30)=0:End Sub

Sub SW52_Hit:Controller.Switch(52)=1:End Sub
Sub SW52_unHit:Controller.Switch(52)=0:End Sub

Sub SW54_Hit:Controller.Switch(54)=1:End Sub
Sub SW54_unHit:Controller.Switch(54)=0:End Sub

'*****************************************
'Rollunder Gates
'*****************************************

Sub SW50_Hit:vpmTimer.PulseSw (50):End Sub

'*****************************************
'Spinners
'*****************************************

Sub Sw35_Spin
  vpmTimer.PulseSw (35)
  PlaySoundAt "fx_spinner", sw35
End Sub

'*****************************************
' Pop Bumpers
'*****************************************

Sub Bumper1_Hit
    vpmTimer.PulseSw 51
    PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper1, 1
    DOF 103,DOFPulse
End Sub

Sub Bumper2_Hit
    vpmTimer.PulseSw 51
    PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper2, 1
    DOF 102,DOFPulse
End Sub

Sub Bumper3_Hit
    vpmTimer.PulseSw 51
    PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper3, 1
    DOF 104,DOFPulse
End Sub

Sub Bumper4_Hit
    vpmTimer.PulseSw 51
    PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper4, 1
    DOF 101,DOFPulse
End Sub

'*****************************************
' Targets
'*****************************************

Sub SW6_Hit:vpmTimer.PulseSw 06 :End Sub      'left bank
Sub SW16_Hit:vpmTimer.PulseSw 16:End Sub
Sub SW26_Hit:vpmTimer.PulseSw 26:End Sub
Sub SW36_Hit:vpmTimer.PulseSw 36:End Sub

Sub SW1_Hit:vpmTimer.PulseSw 01:End Sub       'right bank
Sub SW11_Hit:vpmTimer.PulseSw 11:End Sub
Sub SW21_Hit:vpmTimer.PulseSw 21:End Sub
Sub SW31_Hit:vpmTimer.PulseSw 31:End Sub

Sub SW34_Slingshot                  'kicking target right
  vpmTimer.PulseSw 34
  PlaySoundAtVol SoundFXDOF("fx_solenoid",106,DOFPulse,DOFContactors), PkickHole, 1
  Psw34Kicker.rotx=-2
  Psw34Shield.rotx=-2
  me.uservalue=1
  me.timerenabled=1
End Sub

Sub sw34_timer
    Select Case me.uservalue
'        Case 3:  Psw34Kicker.rotx=11: Psw34Shield.rotx=11
'        Case 4:  Psw34Kicker.rotx=7: Psw34Shield.rotx=7
    Case 1: Psw34Kicker.rotx=1: Psw34Shield.rotx=3: me.TimerEnabled = 0
    End Select
    me.uservalue = me.uservalue + 1
End Sub

Sub Tsw34_hit
  Psw34Kicker.rotx=-2
  Psw34Shield.rotx=-2
End Sub

Sub Tsw34a_hit
  Psw34Kicker.rotx=-5
  Psw34Shield.rotx=-5
End Sub

Sub Tsw34_unhit
  if Not sw34.timerenabled then
    Psw34Kicker.rotx=1
    Psw34Shield.rotx=1
  end if
End Sub

'*****************************************
' Drop Targets
'*****************************************

Sub SW4_Hit:vpmTimer.PulseSw (4):End Sub      'center bank
Sub SW14_Hit:vpmTimer.PulseSw (14):End Sub
Sub SW24_Hit:vpmTimer.PulseSw (24):End Sub

Sub SW3_Hit:vpmTimer.PulseSw (3):End Sub      'left bank
Sub SW13_Hit:vpmTimer.PulseSw (13):End Sub
Sub SW23_Hit:vpmTimer.PulseSw (23):End Sub
Sub SW33_Hit:vpmTimer.PulseSw (33):End Sub
Sub SW53_Hit:vpmTimer.PulseSw (53):End Sub

Sub SW2_Hit:vpmTimer.PulseSw (2):End Sub      'right bank
Sub SW12_Hit:vpmTimer.PulseSw (12):End Sub
Sub SW22_Hit:vpmTimer.PulseSw (22):End Sub
Sub SW32_Hit:vpmTimer.PulseSw (32):End Sub

'****************
' Sling Shot and Rubber Animations
'****************


Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 56
  PlaySoundAtVol SoundFXDOF("fx_slingshot",106,DOFPulse,DOFContactors), slingR, 1
    RSling.Visible = 0
    RSling1.Visible = 1
  slingR.rotx = 20
    me.uservalue = 1
    me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case me.uservalue
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.rotx = 10
        Case 4: slingR.rotx = 0:RSLing2.Visible = 0:RSLing.Visible = 1:me.TimerEnabled = 0
    End Select
    me.uservalue = me.uservalue + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 56
  PlaySoundAtVol SoundFXDOF("fx_slingshot",104,DOFPulse,DOFContactors), slingL, 1
    LSling.Visible = 0
    LSling1.Visible = 1
  slingL.rotx = 20
    me.uservalue= 1
    me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case me.uservalue
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.rotx = 10
        Case 4:slingL.rotx = 0:LSLing2.Visible = 0:LSLing.Visible = 1:me.TimerEnabled = 0
    End Select
    me.uservalue = me.uservalue + 1
End Sub

sub dingwalla_hit
  vpmTimer.PulseSw 56
  SlingA.visible=0
  SlingA1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwalla_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingA1.visible=0: SlingA.visible=1
    case 2: SlingA.visible=0: SlingA2.visible=1
    Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwallb_hit
  vpmTimer.PulseSw 56
  SlingB.visible=0
  SlingB1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwallb_timer                 'default 50 timer
  select case me.uservalue
    Case 1: Slingb1.visible=0: SlingB.visible=1
    case 2: SlingB.visible=0: Slingb2.visible=1
    Case 3: Slingb2.visible=0: SlingB.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwallc_hit
  vpmTimer.PulseSw 56
  Slingc.visible=0
  Slingc1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwallc_timer                 'default 50 timer
  select case me.uservalue
    Case 1: Slingc1.visible=0: Slingc.visible=1
    case 2: Slingc.visible=0: Slingc2.visible=1
    Case 3: Slingc2.visible=0: Slingc.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwalld_hit
  SlingD.visible=0
  SlingD1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwalld_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingD1.visible=0: SlingD.visible=1
    case 2: SlingD.visible=0: SlingD2.visible=1
    Case 3: SlingD2.visible=0: SlingD.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwalle_hit
  SlingD.visible=0
  SlingE1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwalle_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingE1.visible=0: SlingD.visible=1
    case 2: SlingD.visible=0: SlingE2.visible=1
    Case 3: SlingE2.visible=0: SlingD.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

'****************
' Other Stuff :)
'****************

Sub FlipperTimer_Timer()

    dim PI:PI=3.1415926
    LFLogo.Rotz = LeftFlipper.CurrentAngle
    LFLogo1.Rotz = LeftFlipper1.CurrentAngle
    RFLogo.Rotz = RightFlipper.CurrentAngle
    RFLogo1.Rotz = RightFlipper1.CurrentAngle
    LFLogo2.Rotz = LeftFlipper.CurrentAngle
    LFLogo3.Rotz = LeftFlipper1.CurrentAngle
    RfLogo2.Rotz = RightFlipper.CurrentAngle
    RFLogo3.Rotz = RightFlipper1.CurrentAngle

    PrimGate3.Rotz = Gate3.CurrentAngle * 70/90
    PrimGate1.Rotz = Gate1.CurrentAngle * 70/90

  Dim SpinnerRadius: SpinnerRadius=7

  SpinnerRod.TransZ = (cos((sw35.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRod.TransY = sin((sw35.CurrentAngle) * (PI/180)) * -SpinnerRadius

  if FlipperShadows=1 then
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
    FlipperLSh1.RotZ = LeftFlipper1.currentangle
    FlipperRSh1.RotZ = RightFlipper1.currentangle
  end if

  'drop target shadows
  if lgi.state=1 and sw3.isdropped=0 then dropplate1.visible=1
  if lgi.state=1 and sw3.isdropped=1 then dropplate1.visible=0
  if lgi.state=1 and sw13.isdropped=0 then dropplate2.visible=1
  if lgi.state=1 and sw13.isdropped=1 then dropplate2.visible=0
  if lgi.state=1 and sw23.isdropped=0 then dropplate3.visible=1
  if lgi.state=1 and sw23.isdropped=1 then dropplate3.visible=0
  if lgi.state=1 and sw33.isdropped=0 then dropplate4.visible=1
  if lgi.state=1 and sw33.isdropped=1 then dropplate4.visible=0
  if lgi.state=1 and sw53.isdropped=0 then dropplate5.visible=1
  if lgi.state=1 and sw53.isdropped=1 then dropplate5.visible=0
  if lgi.state=1 and sw2.isdropped=0 then dropplate6.visible=1
  if lgi.state=1 and sw2.isdropped=1 then dropplate6.visible=0
  if lgi.state=1 and sw12.isdropped=0 then dropplate7.visible=1
  if lgi.state=1 and sw12.isdropped=1 then dropplate7.visible=0
  if lgi.state=1 and sw22.isdropped=0 then dropplate8.visible=1
  if lgi.state=1 and sw22.isdropped=1 then dropplate8.visible=0
  if lgi.state=1 and sw32.isdropped=0 then dropplate9.visible=1
  if lgi.state=1 and sw32.isdropped=1 then dropplate9.visible=0
  if lgi.state=1 and sw4.isdropped=0 then dropplate10.visible=1
  if lgi.state=1 and sw4.isdropped=1 then dropplate10.visible=0
  if lgi.state=1 and sw14.isdropped=0 then dropplate11.visible=1
  if lgi.state=1 and sw14.isdropped=1 then dropplate11.visible=0
  if lgi.state=1 and sw24.isdropped=0 then dropplate12.visible=1
  if lgi.state=1 and sw24.isdropped=1 then dropplate12.visible=0
End Sub




'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool, added another section to get high score award to set replay cards
Dim TheDips(32)
Sub FindDips
    Dim DipsNumber
    DipsNumber = Controller.Dip(1)
    TheDips(16) = Int(DipsNumber/128)
    If TheDips(16) = 1 then DipsNumber = DipsNumber - 128 end if
    TheDips(15) = Int(DipsNumber/64)
    If TheDips(15) = 1 then DipsNumber = DipsNumber - 64 end if
    TheDips(14) = Int(DipsNumber/32)
    If TheDips(14) = 1 then DipsNumber = DipsNumber - 32 end if
    TheDips(13) = Int(DipsNumber/16)
    If TheDips(13) = 1 then DipsNumber = DipsNumber - 16 end if
    TheDips(12) = Int(DipsNumber/8)
    If TheDips(12) = 1 then DipsNumber = DipsNumber - 8 end if
    TheDips(11) = Int(DipsNumber/4)
    If TheDips(11) = 1 then DipsNumber = DipsNumber - 4 end if
    TheDips(10) = Int(DipsNumber/2)
    If TheDips(10) = 1 then DipsNumber = DipsNumber - 2 end if
    TheDips(9) = Int(DipsNumber)
    DipsNumber = Controller.Dip(2)
    TheDips(24) = Int(DipsNumber/128)
    If TheDips(24) = 1 then DipsNumber = DipsNumber - 128 end if
    TheDips(23) = Int(DipsNumber/64)
    If TheDips(23) = 1 then DipsNumber = DipsNumber - 64 end if
    TheDips(22) = Int(DipsNumber/32)
    If TheDips(22) = 1 then DipsNumber = DipsNumber - 32 end if
    TheDips(21) = Int(DipsNumber/16)
    If TheDips(21) = 1 then DipsNumber = DipsNumber - 16 end if
    TheDips(20) = Int(DipsNumber/8)
    If TheDips(20) = 1 then DipsNumber = DipsNumber - 8 end if
    TheDips(19) = Int(DipsNumber/4)
    If TheDips(19) = 1 then DipsNumber = DipsNumber - 4 end if
    TheDips(18) = Int(DipsNumber/2)
    If TheDips(18) = 1 then DipsNumber = DipsNumber - 2 end if
    TheDips(17) = Int(DipsNumber)
    DipsTimer.Enabled=1
End Sub

 Sub DipsTimer_Timer()
  dim hsaward, BPG

   hsaward = TheDips(23)
   BPG = TheDips(17)

   If BPG = 1 then
       instcard.imageA="InstCard3Balls"
     Else
       instcard.imageA="InstCard5Balls"
   End if
   replaycard.imageA="replaycard"&hsaward
    DipsTimer.enabled=0
 End Sub

'Gottlieb Eclipse
'added by Inkochnito
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Eclipse - DIP switches"
    .AddFrame 2,10,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
    .AddFrame 2,86,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,132,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 2,178,190,"Tilt penalty",&H10000000,Array("game over",0,"ball in play only",&H10000000)'dip29
    .AddFrame 2,224,190,"Game mode",&H00100000,Array("replay",0,"extra ball",&H00100000)'dip 21
    .AddFrameExtra 2,270,190,"Sound option",&H0100,Array("continuous sound",0,"scoring sounds only",&H0100)'S-board dip 1
    .AddChk 2,320,190,Array("Background tone",&H80000000)'dip 32
    '.AddChk 2,335,400,Array("dip 31 must be off",&H40000000)'dip 31
    .AddChk 2,335,190,Array("Must remain on",&H01000000)'dip 25
    .AddChk 2,350,190,Array("Must remain on",&H02000000)'dip 26
    .AddChk 2,365,190,Array("Coin switch tune",&H04000000)'dip 27
    .AddFrame 205,10,190,"High score to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrame 205,86,190,"Balls per game",&H00010000,Array("5 balls",0,"3 balls",&H00010000)'dip 17
    .AddFrame 205,132,190,"Replay limit",&H00040000,Array("no limit",0,"one per game",&H00040000)'dip 19
    .AddFrame 205,178,190,"Novelty",&H00080000,Array("normal",0,"extra ball and replay scores 50,000",&H00080000)'dip 20
    .AddFrame 205,224,190,"center coin chute credits control",&H00001000,Array("no effect",0,"add 9",&H00001000)'dip 13
    .AddFrameExtra 205,270,190,"Attract tune",&H0200,Array("no attract tune",0,"attract tune played every 6 minutes",&H0200)'S-board dip 2
    .AddChk 205,320,190,Array("Attract features",&H20000000)'dip 30
    .AddChk 205,335,190,Array("Match feature",&H00020000)'dip 18
    .AddChk 205,350,190,Array("Credits displayed",&H08000000)'dip 28
    .AddLabel 50,380,300,20,"After hitting OK, press F3 to reset game with new settings."
  End With
  Dim extra
  extra = Controller.Dip(4) + Controller.Dip(5)*256
  extra = vpmDips.ViewDipsExtra(extra)
  Controller.Dip(4) = extra And 255
  Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")



'******************************************************
'     TROUGH BASED ON cyberpez' Black Hole
'******************************************************

Sub sw25_Hit(): Controller.Switch(25) = 1: UpdateTrough:End Sub
Sub sw25_UnHit(): Controller.Switch(25) = 0: End Sub

Sub Trough2_Hit(): UpdateTrough:End Sub
Sub Trough1_Hit(): UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  if Trough1.ballCntOver = 0 then Trough2.kick 70,5
  if Trough2.ballCntOver = 0 then sw25.kick 70,5
  Me.Enabled = 0
End Sub


'******************************************************
'       DRAIN & RELEASE
'******************************************************

Sub sw15_Hit()
  PlaySoundAt "drain", sw15
  Controller.Switch(15) = 1
End Sub


Sub SolOutHole(enabled)
  If enabled Then
    sw15.kick 70,40
    PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw15
    Controller.Switch(15) = 0
  End If
End Sub


Set LampCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
  if startGame.enabled=0 then

    if controller.lamp(12) then
    Trough1.kick 60, 7
    PlaySoundat SoundFX("ballrelease",DOFContactors), Trough1
    UpdateTrough
    end if

    if controller.lamp(13) and sw5.BallCntOver > 0 then
    LSaucer
    end if


    If Controller.Lamp(11) Then   'Game Over triggers match and BIP
        GOBox.text="GAME OVER"
      Else
        GOBox.text=""
    End If


    If Controller.Lamp(1)  Then         'Tilt
        TILTBox.text="TILT"
      Else
        TILTBox.text=""
    End If


    If Controller.Lamp(10) Then       'HIGH SCORE TO DATE
        HStoDateBox.text="HIGH SCORE TO DATE"
      Else
        HStoDateBox.text=""
    End If
  end if
End Sub


Dim Digits(32)

'Score displays

Digits(0)=Array(a1,a2,a3,a4,a5,a6,a7,n,a8)
Digits(1)=Array(a9,a10,a11,a12,a13,a14,a15,n,a16)
Digits(2)=Array(a17,a18,a19,a20,a21,a22,a23,n,a24)
Digits(3)=Array(a25,a26,a27,a28,a29,a30,a31,n,a32)
Digits(4)=Array(a33,a34,a35,a36,a37,a38,a39,n,a40)
Digits(5)=Array(a41,a42,a43,a44,a45,a46,a47,n,a48)
Digits(6)=Array(a49,a50,a51,a52,a53,a54,a55,n,a56)
Digits(7)=Array(a57,a58,a59,a60,a61,a62,a63,n,a64)
Digits(8)=Array(a65,a66,a67,a68,a69,a70,a71,n,a72)
Digits(9)=Array(a73,a74,a75,a76,a77,a78,a79,n,a80)
Digits(10)=Array(a81,a82,a83,a84,a85,a86,a87,n,a88)
Digits(11)=Array(a89,a90,a91,a92,a93,a94,a95,n,a96)
Digits(12)=Array(a97,a98,a99,a100,a101,a102,a103,n,a104)
Digits(13)=Array(a105,a106,a107,a108,a109,a110,a111,n,a112)
Digits(14)=Array(a113,a114,a115,a116,a117,a118,a119,n,a120)
Digits(15)=Array(a121,a122,a123,a124,a125,a126,a127,n,a128)
Digits(16)=Array(a129,a130,a131,a132,a133,a134,a135,n,a136)
Digits(17)=Array(a137,a138,a139,a140,a141,a142,a143,n,a144)
Digits(18)=Array(a145,a146,a147,a148,a149,a150,a151,n,a152)
Digits(19)=Array(a153,a154,a155,a156,a157,a158,a159,n,a160)
Digits(20)=Array(a161,a162,a163,a164,a165,a166,a167,n,a168)
Digits(21)=Array(a169,a170,a171,a172,a173,a174,a175,n,a176)
Digits(22)=Array(a177,a178,a179,a180,a181,a182,a183,n,a184)
Digits(23)=Array(a185,a186,a187,a188,a189,a190,a191,n,a192)

'Ball in Play and Credit displays

Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n)
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n)


Sub DisplayTimer_Timer
    Dim ChgLED,ii,num,chg,stat, obj
    ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
'        If not b2son Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
            if (num < 28 ) then
                For Each obj In Digits(num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg\2 : stat = stat\2
                Next
            else

            end if
        next
'        end if
    end if
End Sub



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

'**************************************************************************
'                 Additional Positional Sound Playback Functions by DJRobX
'**************************************************************************

Sub PlaySoundAtVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set position at table object, vol, and loops manually.

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
    PlaySound sound, Loops, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / Eclipse.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / Eclipse.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'      Ball Rolling Sounds by JP
'*****************************************

Const tnob = 15 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub LeftFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBall "flip_hit_1"
    Case 2 : PlaySoundAtBall "flip_hit_2"
    Case 3 : PlaySoundAtBall "flip_hit_3"
  End Select
End Sub

'*****************************************
'           BALL SHADOW by ninnuzu
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
  Dim maxXoffset
  maxXoffset=15
    BOT = GetBalls

  ' render the shadow for each ball
    For b = 0 to UBound(BOT)
    BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(b).X)/(Eclipse.Width/2))
    BallShadow(b).Y = BOT(b).Y + 10
    If BOT(b).Z > 0 and BOT(b).Z < 30 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

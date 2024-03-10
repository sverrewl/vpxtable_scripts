'************ Time Line Gottlieb 1980


Option Explicit
Randomize

Const cGameName = "timeline"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01130100", "sys80.vbs", 3.36

'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1            '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
Dim ROMSounds: ROMSounds=1      '**********set to 0 for no rom sounds, 1 to play rom sounds.. mostly used for testing

'************************************************
'************************************************
'************************************************
'************************************************
'************************************************

Const UseSolenoids  = 2, UseLamps = 1, UseGI = 0

' Standard Sounds
Const SCoin        = "fx_coin"

Dim xx, objekt, Light
Sub TimeLine_Init()

' Thalamus : Was missing 'vpminit me'
  vpminit me

    On Error Resume Next
      With Controller
        .GameName                               = cGameName
        .SplashInfoLine                         = "Time Line, Gottlieb 1980"
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


' Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3, LeftSLingShot, RightSlingShot, UpperSlingShot)




'Controlled Lights
  vpmMapLights CPULights

'************************Trough

  sw67.CreateSizedBallWithMass 25, 1.3
  Controller.Switch(67) = 1


    if b2son then
    for each xx in backdropstuff
      xx.visible=0
    next
    backdroptimer.enabled=0
    Else
    for each xx in backdropstuff
      xx.visible=1
    next
    backdroptimer.uservalue=0
    backdroptimer.enabled=1
  end If

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

sub backdroptimer_timer
  dim bdlight
  for each bdlight in backdropanimation:bdlight.state=0: Next
  backdropanimation(me.uservalue).state=1
  backdropanimation(me.uservalue+10).state=1
  backdropanimation(me.uservalue+20).state=1
  me.uservalue=me.uservalue+1
  if me.uservalue>9 then me.uservalue=0
end sub

Sub TimeLine_Paused:Controller.Pause = 1:End Sub

Sub TimeLine_unPaused:Controller.Pause = 0:End Sub

Sub TimeLine_Exit
    If b2son then controller.stop
End Sub


Sub TimeLine_KeyDown(ByVal keycode)

    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then
        Plunger.PullBack
        PlaySoundAt "plungerpull", Plunger
    End If



End Sub

Sub TimeLine_KeyUp(ByVal keycode)

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

SolCallBack(1)  = "Braised"   'bottom drop bank
SolCallBack(2)  = "Mraised"   'mid drop bank
SolCallback(5)  = "Lkicker"     'left ball kicker
SolCallback(6)  = "Traised"     'top drop bank
SolCallback(8)  = "SolKnocker"  'knocker
SolCallback(9)  = "SolOutHole"  'drain/outhole


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"



Sub Braised(enabled)
    if enabled then Breset.enabled=True
End Sub

Sub Breset_timer
  PlaySoundAtVol SoundFX("dropbankreset", DOFContactors), sw3, 1
    For each objekt in DTbottom: objekt.isdropped=0: next
    Breset.enabled=false
End Sub

Sub Mraised(enabled)
    if enabled then Mreset.enabled=True
End Sub

Sub Mreset_timer
  PlaySoundAtVol SoundFX("dropbankreset", DOFContactors), sw2, 1
    For each objekt in DTmid: objekt.isdropped=0: next
    Mreset.enabled=false
End Sub

Sub Traised(enabled)
    if enabled then Treset.enabled=True
End Sub

Sub Treset_timer
  PlaySoundAtVol SoundFX("dropbankreset", DOFContactors), sw4, 1
    For each objekt in DTtop: objekt.isdropped=0: next
    Treset.enabled=false
End Sub

Sub SolLFlipper(Enabled)
    If Enabled Then
    if L47.state=0 then
      PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), LFLogo, 1
      LeftFlipper.RotateToEnd
      else
      PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), LFLogo1, 1
      LeftFlipper1.RotateToEnd
    end if
      Else
    if L47.state=0 then
      PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), LFLogo, 1
      LeftFlipper.RotateToStart
      LeftFlipper1.RotateToStart
      else
      PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), LFLogo1, 1
      LeftFlipper.RotateToStart
      LeftFlipper1.RotateToStart
    end if
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    if L47.state=0 then
      PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), RfLogo, 1
      RightFlipper.RotateToEnd
      else
      PlaySoundAtVol SoundFX("fx_flipperup",DOFFlippers), RFLogo1, 1
      RightFlipper1.RotateToEnd
    end if
      Else
    if L47.state=0 then
      PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), RfLogo, 1
      RightFlipper.RotateToStart
      RightFlipper1.RotateToStart
      else
      PlaySoundAtVol SoundFX("fx_flipperdown",DOFFlippers), RfLogo1, 1
      RightFlipper.RotateToStart
      RightFlipper1.RotateToStart
    end if
    End If
End Sub


Sub SolKnocker(Enabled)
    If Enabled Then PlaySoundAtVol SoundFX("Knocker",DOFKnocker), Plunger, 1
End Sub


Sub Lkicker(Enabled)
    If Enabled Then
    KickPlunger.fire
        PlaySoundAtVol SoundFX("popper_ball",DOFContactors), slingK, 1
    slingK.rotx = 20
        KickPlunger.uservalue=1
        KickPlunger.timerenabled=1
  end if
End Sub

Sub KickPlunger_Timer
    Select Case KickPlunger.uservalue
        Case 3:slingK.rotx = 10
        Case 4:slingK.rotx = 0:sw63.TimerEnabled = 0
    End Select
    KickPlunger.uservalue = KickPlunger.uservalue + 1
End Sub


'*****************************************
'Wire rollover Switches
'*****************************************


Sub SW51_Hit:Controller.Switch(51)=1:End Sub
Sub SW51_unHit:Controller.Switch(51)=0:End Sub

Sub SW52_Hit:Controller.Switch(52)=1:End Sub
Sub SW52_unHit:Controller.Switch(52)=0:End Sub

Sub SW53_Hit:Controller.Switch(53)=1:End Sub
Sub SW53_unHit:Controller.Switch(53)=0:End Sub

Sub SW62_Hit:Controller.Switch(62)=1:End Sub
Sub SW62_unHit:Controller.Switch(62)=0:End Sub
Sub SW62a_Hit:Controller.Switch(62)=1:End Sub
Sub SW62a_unHit:Controller.Switch(62)=0:End Sub
Sub SW62b_Hit:Controller.Switch(62)=1:End Sub
Sub SW62b_unHit:Controller.Switch(62)=0:End Sub

Sub SW63_Hit
  Controller.Switch(63)=1
  me.uservalue=1
  me.timerenabled=1
End Sub

Sub SW63_timer
  if GI1.state=1 then
    for each Light in GI: Light.state=0: next
    shadowsGIOFF.visible=1
    shadowsGION.visible=0
    TILTBox.text="TILT"
    else
    for each Light in GI: Light.state=1: next
    shadowsGION.visible=1
    shadowsGIOFF.visible=0
    TILTBox.text=""
  end if
  PlaySoundAt "fx_relay", HexScrew4
  me.uservalue=me.uservalue+1
  if me.uservalue=20 then
    for each Light in GI: Light.state=1: next
    shadowsGION.visible=1
    shadowsGIOFF.visible=0
    TILTBox.text=""
    me.timerenabled=0
  end if
end sub

Sub SW63_unHit:Controller.Switch(63)=0:End Sub

Sub SW72_Hit:Controller.Switch(72)=1:End Sub
Sub SW72_unHit:Controller.Switch(72)=0:End Sub
Sub SW72a_Hit:Controller.Switch(72)=1:End Sub
Sub SW72a_unHit:Controller.Switch(72)=0:End Sub

'*****************************************
' Pop Bumpers
'*****************************************

Sub Bumper1_Hit
    vpmTimer.PulseSw 73
    PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper1, 1
    DOF 104,DOFPulse
End Sub

Sub Bumper2_Hit
    vpmTimer.PulseSw 73
    PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper2, 1
    DOF 103,DOFPulse
End Sub

Sub Bumper3_Hit
    vpmTimer.PulseSw 73
    PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper3, 1
    DOF 105,DOFPulse
End Sub


'*****************************************
' Hit Targets
'*****************************************


Sub SW1_Hit:vpmTimer.PulseSw 01:End Sub
Sub SW21_Hit:vpmTimer.PulseSw 21:End Sub
Sub SW41_Hit:vpmTimer.PulseSw 41:End Sub
Sub SW43_Hit:vpmTimer.PulseSw 43:End Sub
Sub SW61_Hit:vpmTimer.PulseSw 61:End Sub


'*****************************************
' Drop Targets
'*****************************************

Sub SW4_Dropped:vpmTimer.PulseSw (4):End Sub      'top bank
Sub SW14_Dropped:vpmTimer.PulseSw (14):End Sub
Sub SW24_Dropped:vpmTimer.PulseSw (24):End Sub
Sub SW34_Dropped:vpmTimer.PulseSw (34):End Sub
Sub SW44_Dropped:vpmTimer.PulseSw (44):End Sub

Sub SW2_Dropped:vpmTimer.PulseSw (2):End Sub      'center bank
Sub SW12_Dropped:vpmTimer.PulseSw (12):End Sub
Sub SW22_Dropped:vpmTimer.PulseSw (22):End Sub
Sub SW32_Dropped:vpmTimer.PulseSw (32):End Sub
Sub SW42_Dropped:vpmTimer.PulseSw (42):End Sub


Sub SW3_Dropped:vpmTimer.PulseSw (3):End Sub      'bottom bank
Sub SW13_Dropped:vpmTimer.PulseSw (13):End Sub
Sub SW23_Dropped:vpmTimer.PulseSw (23):End Sub
Sub SW33_Dropped:vpmTimer.PulseSw (33):End Sub

'******************************************************
' Sling Shot and Rubber Animations
'******************************************************

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 71
  PlaySoundAtVol SoundFXDOF("fx_slingshot",101,DOFPulse,DOFContactors), slingL, 1
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

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 71
  PlaySoundAtVol SoundFXDOF("fx_slingshot",102,DOFPulse,DOFContactors), slingR, 1
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

Sub UpperSlingShot_Slingshot
  vpmTimer.PulseSw 71
  PlaySoundAtVol SoundFXDOF("fx_slingshot",106,DOFPulse,DOFContactors), slingU, 1
    Usling.Visible = 0
    USling1.Visible = 1
  slingU.rotx = 20
    me.uservalue = 1
    me.TimerEnabled = 1
End Sub

Sub UpperSlingShot_Timer
    Select Case me.uservalue
        Case 3:USLing1.Visible = 0:USLing2.Visible = 1:slingU.rotx = 10
        Case 4: slingU.rotx = 0:USLing2.Visible = 0:USLing.Visible = 1:me.TimerEnabled = 0
    End Select
    me.uservalue = me.uservalue + 1
End Sub

sub dingwalla_hit
  vpmTimer.PulseSw 71
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
  vpmTimer.PulseSw 71
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
  vpmTimer.PulseSw 71
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
  vpmTimer.PulseSw 71
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
  SlingE.visible=0
  SlingE1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwalle_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingE1.visible=0: SlingE.visible=1
    case 2: SlingE.visible=0: SlingE2.visible=1
    Case 3: SlingE2.visible=0: SlingE.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwallf_hit
  SlingF.visible=0
  SlingF1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwallf_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingF1.visible=0: SlingF.visible=1
    case 2: SlingF.visible=0: SlingF2.visible=1
    Case 3: SlingF2.visible=0: SlingF.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwallg_hit
  SlingG.visible=0
  SlingG1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwallg_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingG1.visible=0: SlingG.visible=1
    case 2: SlingG.visible=0: SlingG2.visible=1
    Case 3: SlingG2.visible=0: SlingG.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwallh_hit
  vpmTimer.PulseSw 71
  SlingH.visible=0
  SlingH1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwallh_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingH1.visible=0: SlingH.visible=1
    case 2: SlingH.visible=0: SlingH2.visible=1
    Case 3: SlingH2.visible=0: SlingH.visible=1: Me.timerenabled=0
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

    PrimGate.Rotx = Gate.CurrentAngle * .5
    PreturnGate.Rotz = ReturnGate.CurrentAngle

  if FlipperShadows=1 then
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
    FlipperLSh1.RotZ = LeftFlipper1.currentangle
    FlipperRSh1.RotZ = RightFlipper1.currentangle
  end if


End Sub

Sub DropShadowTimer_Timer()
'                   LOWER BANK
  if sw3.isdropped or Lgi.state=0 Then
    droplower1.visible=0
    Else
    droplower1.visible=1
  end If
  if sw13.isdropped or Lgi.state=0 Then
    droplower2.visible=0
    Else
    droplower2.visible=1
  end If
  if sw23.isdropped or Lgi.state=0 Then
    droplower3.visible=0
    Else
    droplower3.visible=1
  end If
  if sw33.isdropped or Lgi.state=0 Then
    droplower4.visible=0
    Else
    droplower4.visible=1
  end If
'                   UPPER BANK
  if sw4.isdropped or Lgi.state=0 Then
    dropupper1.visible=0
    Else
    dropupper1.visible=1
  end If
  if sw14.isdropped or Lgi.state=0 Then
    dropupper2.visible=0
    Else
    dropupper2.visible=1
  end If
  if sw24.isdropped or Lgi.state=0 Then
    dropupper3.visible=0
    Else
    dropupper3.visible=1
  end If
  if sw34.isdropped or Lgi.state=0 Then
    dropupper4.visible=0
    Else
    dropupper4.visible=1
  end If
  if sw44.isdropped or Lgi.state=0 Then
    dropupper5.visible=0
    Else
    dropupper5.visible=1
  end If
'                   RIGHT BANK
  if sw2.isdropped or Lgi.state=0 Then
    dropright1.visible=0
    Else
    dropright1.visible=1
  end If
  if sw12.isdropped or Lgi.state=0 Then
    dropright2.visible=0
    Else
    dropright2.visible=1
  end If
  if sw22.isdropped or Lgi.state=0 Then
    dropright3.visible=0
    Else
    dropright3.visible=1
  end If
  if sw32.isdropped or Lgi.state=0 Then
    dropright4.visible=0
    Else
    dropright4.visible=1
  end If
  if sw42.isdropped or Lgi.state=0 Then
    dropright5.visible=0
    Else
    dropright5.visible=1
  end If
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

'******************************************************
'Gottlieb Time Line DIP switch settings
'by Inkochnito
'******************************************************
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Time Line - DIP switches"
    .AddChk 2,10,190,Array("Match feature",&H00020000)'dip 18
    .AddChk 2,25,190,Array("Credits displayed",&H08000000)'dip 28
    .AddChk 2,40,190,Array("Attract features",&H20000000)'dip 30
    .AddChk 205,10,190,Array("Coin switch tune",&H04000000)'dip 27
    .AddChk 205,25,190,Array("Replay button tune",&H02000000)'dip 26
    .AddChk 205,40,190,Array("Sound when scoring",&H01000000)'dip 25
    .AddFrame 2,60,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
    .AddFrame 2,136,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,182,190,"Playfield special",&H00200000,Array("awards replay",0,"awards extra ball",&H00200000)'dip 22
    .AddFrame 2,228,190,"Million points replay",&H80000000,Array("when score exceeds 1 million",0,"every time score reaches a million",&H80000000)'dip 32
    .AddFrame 2,320,190,"Balls per game",&H00010000,Array("5 balls",0,"3 balls",&H00010000)'dip 17
    .AddFrameExtra 2,274,190,"Attract sound",&H0002,Array("off",0,"play tune every 6 minutes",&H0002)'S-board dip 2
    .AddFrame 205,60,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrame 205,136,190,"Replay limit",&H00040000,Array("no replay limit",0,"one per game",&H00040000)'dip 19
    .AddFrame 205,182,190,"Novelty mode",&H00080000,Array("normal mode",0,"special && extra ball scores 50K",&H00080000)'dip 20
    .AddFrame 205,228,190,"Game mode",&H00100000,Array("replay",0,"extra ball",&H00100000)'dip 21
    .AddFrame 205,274,190,"3rd coin chute credits control",&H00001000,Array("no effect",0,"add 9",&H00001000)'dip 13
    .AddFrame 205,320,190,"Tilt penalty",&H10000000,Array("game over",0,"ball in play",&H10000000)'dip 29
    .AddLabel 50,376,300,20,"After hitting OK, press F3 to reset game with new settings."
  End With
  Dim extra
  extra = Controller.Dip(4) + Controller.Dip(5)*256
  extra = vpmDips.ViewDipsExtra(extra)
  Controller.Dip(4) = extra And 255
  Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")



'******************************************************
'       DRAIN & RELEASE
'******************************************************

Sub sw67_Hit()
  PlaySoundAt "drain", sw67
  Controller.Switch(67) = 1
End Sub


Sub SolOutHole(enabled)
  If enabled Then
    sw67.kick 70,12
    PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw67
    Controller.Switch(67) = 0
  End If
End Sub

'******************************************************
'       LAMP CONTROLLED RELAYS
'******************************************************

'Set LampCallback=GetRef("UpdateMultipleLamps")

Dim GameOverActive: GameOverActive=True

Sub RelayLampTimer_timer

  if L0.state=1 then      'Game Over Relay
    if GameOverActive=True then
      GameOverActive=False
      PlaySoundAt "fx_relay", HexScrew4
    end if
    else
    if GameOverActive=false then
      GameOverActive=True
      PlaySoundAt "fx_relay", HexScrew4
    end if
  end if

    If L1.state=1  Then         'Tilt relay
    if gi1.state=1 and NOT sw63.timerenabled then
      For each xx in GI:xx.State = 0: Next
      shadowsGIOFF.visible=1
      shadowsGION.visible=0
      PlaySoundAt "fx_relay", HexScrew4
      TILTBox.text="TILT"
    end if
      Else
    if gi1.state=0 and NOT sw63.timerenabled then
      For each xx in GI:xx.State = 1: Next
      shadowsGION.visible=1
      shadowsGIOFF.visible=0
      TILTBox.text=""
    end if
    End If

    If L45.state=1 and NOT GameOverActive Then    'L45 triggers gate to return to shooter lane
    if ReturnGate.currentangle=0 then
      ReturnGate.rotatetoend
      PlaySoundAt "fx_relay", ReturnGate
    end if
    else
    if ReturnGate.currentangle=-34 then
      ReturnGate.rotatetostart
      PlaySoundAt "fx_relay", ReturnGate
    end if
  end if

    If L11.state=1 Then   'Game Over light
    if GOBox.text="" Then
      GOBox.text="GAME OVER"
    end If
      Else
    if GOBox.text="GAME OVER" Then
      GOBox.text=""
    end If
    End If


    If L10.state=1 Then       'HIGH SCORE TO DATE
        if HStoDateBox.text="" then
      HStoDateBox.text="HIGH SCORE TO DATE"
      PlaySound "fx_relay2"
    end If
      Else
        if HStoDateBox.text<>"" then
      HStoDateBox.text=""
      PlaySound "fx_relay2"
    end If
    End If

End sub




'******************************************************
'       SCORE DISPLAYS
'******************************************************

Dim Digits(33)

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

'Ball in Play display

Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n)

'Credit display

Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n)

'Bonus Count Display
Digits(28)=Array(g00,g01,g02,g03,g04,g05,g06,n,g07)
Digits(29)=Array(g10,g11,g12,g13,g14,g15,g16,n,g17)


Sub DisplayTimer_Timer
    Dim ChgLED,ii,num,chg,stat, obj
    ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
'        If not b2son Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
            if (num < 30 ) then
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
    tmp = tableobj.y * 2 / TimeLine.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / TimeLine.width-1
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

Const tnob = 3 ' total number of balls
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
'           BALL SHADOW by ninnuzu, modded by BorgDog
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
    BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(b).X)/(TimeLine.Width/2))
    BallShadow(b).Y = BOT(b).Y + 10
    If BOT(b).Z > 0 and BOT(b).Z < 30 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

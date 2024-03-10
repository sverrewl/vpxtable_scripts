Option Explicit
Randomize

Const cGameName = "silvslug"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01560000", "GTS3.VBS", 3.26


'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1          '*********** set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1    '*********** set to 1 to turn on Flipper shadows
Dim CardColor: CardColor=1        '*********** set to 1 for black apron cards, 0 for white (stock)
Dim ROMSounds: ROMSounds=1        '*********** set to 0 for no rom sounds, 1 to play rom sounds.. mostly used for testing


'************************************************
'************************************************

Dim Ballsize,BallMass
BallSize = 50
BallMass = 1.2

'************************************************
'************************************************
'************************************************
Const UseSolenoids = 2
Const UseLamps = True
Const UseSync = False
Const UseGI = False

Const VT_Delay_Factor = .86   'used to slow down the ball when hitting the vari targets, smaller number slows down faster


' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

Const swStartButton=3

Dim objekt, xx, light



Sub SilverSlugger_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Silver Slugger (Gottlieb 1990)"&chr(13)&"1.0"
         .HandleKeyboard  = 0
         .ShowTitle     = 0
         .ShowDMDOnly     = 0
         .ShowFrame     = 0
         .HandleMechanics   = 0
     .Games(cGameName).Settings.Value("sound") = ROMSounds
     .Hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

  plungerkickback.Pullback

  If B2SOn then
    for each objekt in backdropstuff : objekt.visible = 0 : next
    Else
    for each objekt in backdropstuff : objekt.visible = 1 : next
  End If

  Intensity 'sets GI brightness depending on day/night slider settings


    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

'************************Load Trough

  sw44.CreateSizedballWithMass Ballsize/2,Ballmass
  Slot2.CreateSizedballWithMass Ballsize/2,Ballmass
  Slot3.CreateSizedballWithMass Ballsize/2,Ballmass
  Controller.Switch(44) = 1
  Controller.Switch(43) = 0

'***********************Nudging
  vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, bumper1, bumper2, bumper3, PlungerKickback)

'***********************Options setup

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

    if CardColor=1 then
        instcard.image = "instcard"
    repcard.image = "repcard"
       else
        instcard.image = "instcardw"
    repcard.image = "repcardw"
    end if


  vpmMapLights CPULights

End Sub


'************************************************
' Solenoids
'************************************************
SolCallback(6) =    "SolRLock"
SolCallback(7) =    "SolLLock"
SolCallback(8) =  "solVariReset"
SolCallback(9) =  "vpmFlasher BumperLight8,"
SolCallback(10) = "vpmFlasher BumperLight9,"
SolCallback(11) = "vpmFlasher BumperLight10,"

SolCallback(17) ="vpmFlasher array(F16a, F16b, F16c, F16d),"
SolCallback(18) ="vpmFlasher array(F17a, F17b, F17c, F17d),"
SolCallback(19) ="vpmFlasher array(F18a, F18b, F18c, F18d),"
SolCallback(20) ="vpmFlasher array(F19a, F19b, F19c, F19d),"
SolCallback(21) ="vpmFlasher array(F20a, F20b, F20c, F20d),"
SolCallback(22) ="vpmFlasher F21a,"
SolCallback(23) ="vpmFlasher F22a,"

SolCallback(24) =   "dtRaised"
SolCallback(25) = "solKickback"

SolCallback(28) =    "solballrelease"
SolCallback(29) =    "SolOuthole"
SolCallback(30) =    "solknocker"
SolCallback(31) =   "solTiltRelay"
SolCallback(32) =   "solGameOverRelay"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


Sub solTiltRelay(Enabled)
  if enabled and GI1.state=1 Then
    GIturnOFF
    Else
     if GI1.state=0 Then GIturnON
  end if
end Sub

sub solGameOverRelay(Enabled)
  if enabled and GI1.state=0 Then
    GIturnON
    Else
    if GI1.state=1 Then GIturnOFF
  end if
end Sub

Sub GIturnON
  For each objekt in GI:objekt.State = 1: Next
  PLaneGuides.image="SSmetalsGIon"
  shadowsGION.visible=1
  shadowsGIOFF.visible=0
end sub

Sub GIturnOFF
  For each objekt in GI:objekt.State = 0: Next
  PLaneGuides.image="SSmetalsGIoff"
  shadowsGIOFF.visible=1
  shadowsGION.visible=0
End sub

Sub SolLFlipper(Enabled)
     If Enabled Then
        PlaySoundAt SoundFX("fx_Flipperup",DOFFlippers), LeftFlipper:LeftFlipper.RotateToEnd
     Else
        PlaySoundAt SoundFX("fx_Flipperdown",DOFFlippers), LeftFlipper:LeftFlipper.RotateToStart
     End If
  End Sub


Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("fx_Flipperup",DOFFlippers), RightFlipper:RightFlipper.RotateToEnd
     Else
         PlaySoundAt SoundFX("fx_Flipperdown",DOFFlippers), RightFlipper:RightFlipper.RotateToStart
     End If
End Sub

Dim GILevel, DayNight

Sub Intensity
  If DayNight <= 20 Then
      GILevel = .5
  ElseIf DayNight <= 40 Then
      GILevel = .4125
  ElseIf DayNight <= 60 Then
      GILevel = .325
  ElseIf DayNight <= 80 Then
      GILevel = .2375
  Elseif DayNight <= 100  Then
      GILevel = .15
  End If

  For each xx in GI: xx.IntensityScale = xx.IntensityScale * (GILevel): Next
  For each xx in DTLights: xx.IntensityScale = xx.IntensityScale * (GILevel): Next

End Sub



Sub FlipperTimer_Timer


  LFlip.RotY = LeftFlipper.CurrentAngle-90
  RFlip.RotY = RightFlipper.CurrentAngle-90

  Dim PI: PI=3.1415926
  Dim SpinnerRadius: SpinnerRadius=7

  SpinnerRodsw22.TransZ = (cos((sw22.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRodsw22.TransY = sin((sw22.CurrentAngle) * (PI/180)) * -SpinnerRadius

  SpinnerRodsw32.TransZ = (cos((sw32.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRodsw32.TransY = sin((sw32.CurrentAngle) * (PI/180)) * -SpinnerRadius

  SpinnerRodsw42.TransZ = (cos((sw42.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRodsw42.TransY = sin((sw42.CurrentAngle) * (PI/180)) * -SpinnerRadius

  Psw22.rotx=360-sw22.currentangle
  Psw32.rotx=sw32.currentangle
  Psw42.rotx=sw42.currentangle

  FwReflect.visible=l30.state
  FaReflect.visible=l31.state
  FlReflect.visible=l32.state
  FkReflect.visible=l33.state
  FtriReflect.visible=l47.state
  Ftri2Reflect.visible=l56.state

  if FlipperShadows=1 then
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
  end if
End Sub


' Ball locks / kickers

Sub sw50_Hit:PlaySoundat "holein", sw50: Controller.Switch(50) = 1:End Sub      'Left lock
Sub sw51_Hit:PlaySoundAt "holein", sw51: Controller.Switch(51) = 1:End Sub      'Right Lock


Sub SolLLock(enabled)
  If enabled Then
    sw50.kick 110, 14, 20
    PlaySoundAt "HoleKick", sw50
    Controller.Switch(50) = 0
    LeftKickTimer.uservalue = 0
    Pkickarm.Roty = 22
    LeftKickTimer.Enabled = 1
  End If
End Sub

Sub LeftKickTimer_timer
  select case me.uservalue
    case 5:
    Pkickarm.roty=0
    me.enabled=0
  end Select
  me.uservalue = me.uservalue+1
End Sub


Sub SolRLock(enabled)
  If enabled Then
    sw51.kick 270,13, 20
    PlaySoundAt "HoleKick", sw51
    Controller.Switch(51) = 0
    RightKickTimer.uservalue = 0
    PkickarmR.Roty = 22
    RightKickTimer.Enabled = 1
  End If
End Sub

Sub RightKickTimer_timer
  select case me.uservalue
    case 5:
    PkickarmR.roty=0
    me.enabled=0
  end Select
  me.uservalue = me.uservalue+1
End Sub

sub solKickBack(Enabled)
  if enabled Then
    PlaySoundAt "fx_popper", PlungerKickback
    PlungerKickback.Fire
    Else
    plungerkickback.Pullback
  end if
end Sub

'******************************************************
'     DRAIN & TROUGH BASED ON NFOZZY'S via rothbauerw's Tee'd Off
'******************************************************

Sub Slot3_Hit():UpdateTrough:End Sub
Sub Slot3_UnHit():UpdateTrough:End Sub
Sub Slot2_UnHit():UpdateTrough:End Sub


Sub UpdateTrough()
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw44.BallCntOver = 0 Then Slot2.kick 60, 9
  If Slot2.BallCntOver = 0 Then Slot3.kick 60, 9
  Me.Enabled = 0
End Sub

Sub sw43_Hit()        'drain
  Controller.Switch(43) = 1
  PlaySoundAt "drain", sw43
  UpdateTrough
End Sub
Sub sw43_UnHit():Controller.Switch(43) = 0:End Sub

Sub sw44_Hit(): PlaySoundat "Switch", sw44: Controller.Switch(44) = 1: End Sub
Sub sw44_UnHit(): Controller.Switch(44) = 0: UpdateTrough: End Sub

Sub SolOuthole(enabled)
  If enabled Then
    sw43.kick 60,20
    PlaySoundat SoundFX("holekick",DOFContactors), sw43
  End If
End Sub

Sub solballrelease(enabled)
  If enabled Then
    PlaySoundat SoundFX("holekick",DOFContactors), sw44
    sw44.kick 60, 7
  End If
End Sub



 '**********************************************************************************************************
'Digital Display on backdrop
'**********************************************************************************************************
 Dim Digits(48)
'segment ID a,  b,  c,  d,  e,  f,  m, COM, n,  g,  h,  i,  j,  k,  l,  DP
'number ID  00, 01, 02, 03, 04, 05, 06, 07, 08, 09, 10, 11, 12, 13, 14, 15

  Digits(0)=Array(a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15)
  Digits(1)=Array(a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28, a29, a30, a31)
  Digits(2)=Array(a32, a33, a34, a35, a36, a37, a38, a39, a40, a41, a42, a43, a44, a45, a46, a47)
  Digits(3)=Array(a48, a49, a50, a51, a52, a53, a54, a55, a56, a57, a58, a59, a60, a61, a62, a63)
  Digits(4)=Array(a64, a65, a66, a67, a68, a69, a70, a71, a72, a73, a74, a75, a76, a77, a78, a79)
  Digits(5)=Array(a80, a81, a82, a83, a84, a85, a86, a87, a88, a89, a90, a91, a92, a93, a94, a95)
  Digits(6)=Array(a96, a97, a98, a99, a100, a101, a102, a103, a104, a105, a106, a107, a108, a109, a110, a111)
  Digits(7)=Array(a112, a113, a114, a115, a116, a117, a118, a119, a120, a121, a122, a123, a124, a125, a126, a127)
  Digits(8)=Array(a128, a129, a130, a131, a132, a133, a134, a135, a136, a137, a138, a139, a140, a141, a142, a143)
  Digits(9)=Array(a144, a145, a146, a147, a148, a149, a150, a151, a152, a153, a154, a155, a156, a157, a158, a159)
  Digits(10)=Array(a160, a161, a162, a163, a164, a165, a166, a167, a168, a169, a170, a171, a172, a173, a174, a175)
  Digits(11)=Array(a176, a177, a178, a179, a180, a181, a182, a183, a184, a185, a186, a187, a188, a189, a190, a191)
  Digits(12)=Array(a192, a193, a194, a195, a196, a197, a198, a199, a200, a201, a202, a203, a204, a205, a206, a207)
  Digits(13)=Array(a208, a209, a210, a211, a212, a213, a214, a215, a216, a217, a218, a219, a220, a221, a222, a223)
  Digits(14)=Array(a224, a225, a226, a227, a228, a229, a230, a231, a232, a233, a234, a235, a236, a237, a238, a239)
  Digits(15)=Array(a240, a241, a242, a243, a244, a245, a246, a247, a248, a249, a250, a251, a252, a253, a254, a255)
  Digits(16)=Array(a256, a257, a258, a259, a260, a261, a262, a263, a264, a265, a266, a267, a268, a269, a270, a271)
  Digits(17)=Array(a272, a273, a274, a275, a276, a277, a278, a279, a280, a281, a282, a283, a284, a285, a286, a287)
  Digits(18)=Array(a288, a289, a290, a291, a292, a293, a294, a295, a296, a297, a298, a299, a300, a301, a302, a303)
  Digits(19)=Array(a304, a305, a306, a307, a308, a309, a310, a311, a312, a313, a314, a315, a316, a317, a318, a319)

  Digits(20)=Array(b0, b1, b2, b3, b4, b5, b6, b7, b8, b9, b10, b11, b12, b13, b14, b15)
  Digits(21)=Array(b16, b17, b18, b19, b20, b21, b22, b23, b24, b25, b26, b27, b28, b29, b30, b31)
  Digits(22)=Array(b32, b33, b34, b35, b36, b37, b38, b39, b40, b41, b42, b43, b44, b45, b46, b47)
  Digits(23)=Array(b48, b49, b50, b51, b52, b53, b54, b55, b56, b57, b58, b59, b60, b61, b62, b63)
  Digits(24)=Array(b64, b65, b66, b67, b68, b69, b70, b71, b72, b73, b74, b75, b76, b77, b78, b79)
  Digits(25)=Array(b80, b81, b82, b83, b84, b85, b86, b87, b88, b89, b90, b91, b92, b93, b94, b95)
  Digits(26)=Array(b96, b97, b98, b99, b100, b101, b102, b103, b104, b105, b106, b107, b108, b109, b110, b111)
  Digits(27)=Array(b112, b113, b114, b115, b116, b117, b118, b119, b120, b121, b122, b123, b124, b125, b126, b127)
  Digits(28)=Array(b128, b129, b130, b131, b132, b133, b134, b135, b136, b137, b138, b139, b140, b141, b142, b143)
  Digits(29)=Array(b144, b145, b146, b147, b148, b149, b150, b151, b152, b153, b154, b155, b156, b157, b158, b159)
  Digits(30)=Array(b160, b161, b162, b163, b164, b165, b166, b167, b168, b169, b170, b171, b172, b173, b174, b175)
  Digits(31)=Array(b176, b177, b178, b179, b180, b181, b182, b183, b184, b185, b186, b187, b188, b189, b190, b191)
  Digits(32)=Array(b192, b193, b194, b195, b196, b197, b198, b199, b200, b201, b202, b203, b204, b205, b206, b207)
  Digits(33)=Array(b208, b209, b210, b211, b212, b213, b214, b215, b216, b217, b218, b219, b220, b221, b222, b223)
  Digits(34)=Array(b224, b225, b226, b227, b228, b229, b230, b231, b232, b233, b234, b235, b236, b237, b238, b239)
  Digits(35)=Array(b240, b241, b242, b243, b244, b245, b246, b247, b248, b249, b250, b251, b252, b253, b254, b255)
  Digits(36)=Array(b256, b257, b258, b259, b260, b261, b262, b263, b264, b265, b266, b267, b268, b269, b270, b271)
  Digits(37)=Array(b272, b273, b274, b275, b276, b277, b278, b279, b280, b281, b282, b283, b284, b285, b286, b287)
  Digits(38)=Array(b288, b289, b290, b291, b292, b293, b294, b295, b296, b297, b298, b299, b300, b301, b302, b303)
  Digits(39)=Array(b304, b305, b306, b307, b308, b309, b310, b311, b312, b313, b314, b315, b316, b317, b318, b319)

'
'Runs displays
  Digits(40)=Array(f0,f1,f2,f3,f4,f5,f6,n)
  Digits(41)=Array(f7,f8,f9,f10,f11,f12,f13,n)
  Digits(42)=Array(f14,f15,f16,f17,f18,f19,f20,n)
  Digits(43)=Array(f21,f22,f23,f24,f25,f26,f27,n)
  Digits(44)=Array(f28,f29,f30,f31,f32,f33,f34,n)
  Digits(45)=Array(f35,f36,f37,f38,f39,f40,f41,n)
  Digits(46)=Array(f42,f43,f44,f45,f46,f47,f48,n)
  Digits(47)=Array(f49,f50,f51,f52,f53,f54,f55,n)

 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
    If not B2SOn Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 48) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next
     end if
    End If
 End Sub






'**********************************************************************************************************
'Plunger and Advance button (initials) code
'**********************************************************************************************************

Sub SilverSlugger_KeyDown(ByVal KeyCode)

    If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt "plungerpull", Plunger

  If keycode = LeftFlipperKey then Controller.Switch(4)=1
  If keycode = RightFlipperKey then Controller.Switch(5)=1

    If KeyDownHandler(keycode) Then Exit Sub

End Sub

Sub SilverSlugger_KeyUp(ByVal KeyCode)

    If keycode = PlungerKey Then
    Plunger.Fire
    if sw52.BallCntOver>0 then
      PlaySoundAt "plungerreleaseball", Plunger
      else
      PlaySoundAt "plungerreleasefree", Plunger
    end if
  end if

  If keycode = LeftFlipperKey then Controller.Switch(4)=0
  If keycode = RightFlipperKey then Controller.Switch(5)=0

    If KeyUpHandler(keycode) Then Exit Sub

End Sub



'Bumpers

Sub bumper1_Hit : vpmTimer.PulseSw 10 : PlaySoundAt SoundFX("fx_bumper4",DOFContactors), Bumper1: End Sub
Sub bumper2_Hit : vpmTimer.PulseSw 12 : PlaySoundAt SoundFX("fx_bumper4",DOFContactors), Bumper2: End Sub
Sub bumper3_Hit : vpmTimer.PulseSw 11 : PlaySoundAt SoundFX("fx_bumper4",DOFContactors), Bumper3: End Sub


'Wire Triggers
Sub SW15_Hit:Controller.Switch(15)=1:End Sub
Sub SW15_unHit:Controller.Switch(15)=0:End Sub
Sub SW25_Hit:Controller.Switch(25)=1:End Sub
Sub SW25_unHit:Controller.Switch(25)=0:End Sub
Sub SW35_Hit:Controller.Switch(35)=1:End Sub
Sub SW35_unHit:Controller.Switch(35)=0:End Sub
Sub SW45_Hit:Controller.Switch(45)=1:End Sub
Sub SW45_unHit:Controller.Switch(45)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1:End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW33_Hit:Controller.Switch(33)=1:End Sub
Sub SW33_unHit:Controller.Switch(33)=0:End Sub
Sub SW34_Hit:Controller.Switch(34)=1:End Sub
Sub SW34_unHit:Controller.Switch(34)=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1:End Sub
Sub SW24_unHit:Controller.Switch(24)=0:End Sub
Sub SW52_Hit:Controller.Switch(52)=1:End Sub
Sub SW52_unHit:Controller.Switch(52)=0:End Sub


'Targets
Sub sw20_Hit:vpmTimer.PulseSw (20):End Sub
Sub sw30_Hit:vpmTimer.PulseSw (30):End Sub
Sub sw40_Hit:vpmTimer.PulseSw (40):End Sub
Sub sw21_Hit:vpmTimer.PulseSw (21):End Sub
Sub sw31_Hit:vpmTimer.PulseSw (31):End Sub
Sub sw41_Hit:vpmTimer.PulseSw (41):End Sub

'Drop Targets
Dim zMultiplier: zMultiplier = 2.4

Sub SW17_dropped:Controller.Switch(17)=1:End Sub
Sub SW27_dropped:Controller.Switch(27)=1:End Sub
Sub SW37_dropped:Controller.Switch(37)=1:End Sub
Sub SW47_dropped:Controller.Switch(47)=1:End Sub

Sub DT17_hit()
  sw17.isdropped = true
  DT17.collidable = false
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub DT27_hit()
  sw27.isdropped = true
  DT27.collidable = false
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub DT37_hit()
  sw37.isdropped = true
  DT37.collidable = false
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub DT47_hit()
  sw47.isdropped = true
  DT47.collidable = false
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub dtRaised(Enabled)
  if enabled then DTreset.enabled=True
End Sub

Sub DTreset_timer
  For each objekt in DropTargets: objekt.isdropped=0: Next
  For each objekt in a_DropTargets: objekt.collidable=true: Next
  Controller.Switch(17)=0
  Controller.Switch(27)=0
  Controller.Switch(37)=0
  Controller.Switch(47)=0
  For each light in DTLights: light.intensityscale=DTshadow: Next
  me.enabled=false
End Sub

'Spinners

Sub Sw22_Spin
  vpmTimer.PulseSw (22)
  PlaySoundAt "fx_spinner", sw22
End Sub

Sub Sw32_Spin
  vpmTimer.PulseSw (32)
  PlaySoundAt "fx_spinner", sw32
End Sub

Sub Sw42_Spin
  vpmTimer.PulseSw (42)
  PlaySoundat "fx_spinner", sw42
End Sub

'Knocker

Sub SolKnocker(Enabled)
    If Enabled Then PlaySoundAt SoundFX("Knock",DOFKnocker), Plunger
End Sub

'****************GLASS HIT


sub Glass_hit
  PlaySoundAtBall "fx_glass"
end sub


''***********************************************************************************
''****                  VariTarget Handling                   ****
''***********************************************************************************

Sub solVarireset(enabled)
  VTReset 1, enabled
  'debug.print enabled
End Sub

'******************************************************
'*   VARI TARGET
'******************************************************

Sub VariTargetStart_Hit: VTHit 1:End Sub

'Define a variable for each vari target
Dim VT1

'Set array with vari target objects
'
' VariTargetvar = Array(primary, prim, swtich)
'   primary:    vp target to determine target hit
'   secondary:    vp target at the end of the vari target path
'   prim:       primitive target used for visuals and animation
'           IMPORTANT!!!
'           rotx must be used to offset the target animation
'   num:      unique number to identify the vari target
'   plength:    length from the pivot point of the primitive to the hit point/center of the target
'   width:      width of the vari target
'   kspring:    Spring strength constant
' stops:      Number of notches in the vari target including start position, defines where the target will stop
'   rspeed:     return speed of the target in vp units per second
'   animate:    Arrary slot for handling the animation instrucitons, set to 0
'

dim v1dist: v1dist = Distance2Obj(VariTargetStart, VariTargetStop)

Set VT1 = (new VariTarget)(VariTargetStart, VariTargetStop, VariTargetP, 1, 335, 50, 2, 7, 600, 0)

' Index, distance from seconardy, switch number (the first switch should fire at the first Stop {number of stops - 1})
VT1.addpoint 0, v1dist/6, 16
VT1.addpoint 1, v1dist*2/5, 26
VT1.addpoint 2, v1dist*3/5, 36
VT1.addpoint 3, v1dist*4/5, 46

'Add all the Vari Target Arrays to Vari Target Animation Array
'   VTArray = Array(VT1, VT2, ....)
Dim VTArray
VTArray = Array(VT1)

Class VariTarget
  Private m_primary, m_secondary, m_prim, m_num, m_plength, m_width, m_kspring, m_stops, m_rspeed, m_animate
  Public Distances, Switches, Ball

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Num(): Num = m_num: End Property
  Public Property Let Num(input): m_num = input: End Property

  Public Property Get PLength(): PLength = m_plength: End Property
  Public Property Let PLength(input): m_plength = input: End Property

  Public Property Get Width(): Width = m_width: End Property
  Public Property Let Width(input): m_width = input: End Property

  Public Property Get KSpring(): KSpring = m_kspring: End Property
  Public Property Let KSpring(input): m_kspring = input: End Property

  Public Property Get Stops(): Stops = m_stops: End Property
  Public Property Let Stops(input): m_stops = input: End Property

  Public Property Get RSpeed(): RSpeed = m_rspeed: End Property
  Public Property Let RSpeed(input): m_rspeed = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, secondary, prim, num, plength, width, kspring, stops, rspeed, animate)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    m_num = num
    m_plength = plength
    m_width = width
    m_kspring = kspring
    m_stops = stops
    m_rspeed = rspeed
    m_animate = animate

    Set Init = Me
    redim Distances(0)
    redim Switches(0)
  End Function

  Public Sub AddPoint(aIdx, dist, sw)
    ShuffleArrays Distances, Switches, 1 : Distances(aIDX) = dist : Switches(aIDX) = sw : ShuffleArrays Distances, Switches, 0
  End Sub
End Class

'''''' VARI TARGET FUNCTIONS

Sub VTHit(num)
  Dim i
  i = VTArrayID(num)

  If VTArray(i).animate <> 2 Then
    VTArray(i).animate = 1 'STCheckHit(ActiveBall,VTArray(i).primary) 'We don't need STCheckHit because VariTarget geometry should only allow a valid hit
  End If

   Set VTArray(i).ball = Activeball

  Activeball.velx = cor.BallVelx(Activeball.id)
  Activeball.vely = cor.BallVely(Activeball.id)

  DoVTAnim
End Sub

Sub VTReset(num, enabled)
  Dim i
  i = VTArrayID(num)

  If enabled = true then
    VTArray(i).animate = 2
  Else
    VTArray(i).animate = 1
  End If

  DoVTAnim
End Sub

Function VTArrayID(num)
  Dim i
  For i = 0 To UBound(VTArray)
    If VTArray(i).num = num Then
      VTArrayID = i
      Exit Function
    End If
  Next
End Function

Sub VTAnim_Timer
  DoVTAnim
  Cor.Update
End Sub

Sub DoVTAnim()
  Dim i
  For i = 0 To UBound(VTArray)
    VTArray(i).animate = VTAnimate(VTArray(i))
  Next
End Sub

Function VTAnimate(arr)
  VTAnimate = arr.animate

  If arr.animate = 0  Then
    arr.primary.uservalue = 0
    VTAnimate = 0
    arr.primary.collidable = 1
    Exit Function
  ElseIf arr.primary.uservalue = 0 Then
    arr.primary.uservalue = GameTime
  End If

  If arr.animate <> 0 Then
    Dim animtime, length, btdist, btwidth, angle
    Dim tdist, transP, transPnew, cstop, x
    cstop = 0

    animtime = GameTime - arr.primary.uservalue
    arr.primary.uservalue = GameTime

    length = Distance2Obj(VariTargetStart, VariTargetStop)

    angle = arr.primary.orientation
    transP = dSin(arr.prim.rotx)*arr.plength 'previous distance target has moved from start
    transPnew = transP + arr.rspeed * animtime/1000

    If arr.animate = 1 then
      for x = 0 to (arr.Stops - 1)
        dim d: d = -length * x / (arr.Stops - 1) 'stops at end of path, remove  - 1 to stop short of the end of path
        If transP - 0.01 <= d and transPnew + 0.01 >= d Then
          transPnew = d
          cstop = d
          'debug.print x & " " & d
        End If
      next
    End If

    if not isEmpty(arr.ball) Then
      arr.primary.collidable = 0
      tdist = 31.31 'distance between ball and target location on hit event

      btdist = DistancePL(arr.ball.x,arr.ball.y,arr.secondary.x,arr.secondary.y,arr.secondary.x+dcos(angle),arr.secondary.y+dsin(angle))-tdist 'distance between the ball and secondary target
      btwidth = DistancePL(arr.ball.x,arr.ball.y,arr.primary.x,arr.primary.y,arr.primary.x+dcos(angle+90),arr.primary.y+dsin(angle+90)) 'distance between the ball and the parallel patch of the target

      If transPnew + length => btdist and btwidth < arr.width/2 + 25 Then
        arr.ball.velx = arr.ball.velx - (arr.kspring * dsin(angle) * abs(transP) * animtime/1000)
        arr.ball.vely = arr.ball.vely + (arr.kspring * dcos(angle) * abs(transP) * animtime/1000)
        transPnew = btdist - length
        If arr.secondary.uservalue <> 1 then:PlayVTargetSound(arr.ball):arr.secondary.uservalue = 1:End If
      End If
      If btdist > length + tdist Then
        arr.ball = Empty
        arr.primary.collidable = 1
        arr.secondary.uservalue = 0
      End If
    End If

    arr.prim.rotx = dArcSin(transPnew/arr.plength)
    VTSwitch arr, transPnew

    if arr.prim.rotx >= 0 Then
      arr.prim.rotx = 0
      VTSwitch arr, 0
      VTAnimate = 0
      Exit Function
    elseif cstop = transPnew and isEmpty(arr.ball) and arr.animate <> 2 Then
      VTAnimate = 0
      debug.print cstop & " " & Controller.Switch(16) & " " & Controller.Switch(26) & " " & Controller.Switch(36) & " " & Controller.Switch(46)
    end If
  End If
End Function

Sub VTSwitch(arr, transP)
  Dim x, count, sw
  sw = 0
  count = 0
  For each x in arr.distances
    If abs(transP) > x Then
      sw = arr.switches(Count)
      count = count + 1
    End If
  Next
  For each x in arr.switches
    If x <> 0 Then Controller.Switch(x) = 0
  Next
  If sw <> 0 Then Controller.Switch(sw) = 1
End Sub

Sub PlayVTargetSound(ball)
  PlaySound SoundFX("fx_target",DOFTargets), 0, Vol(Ball), AudioPan(Ball), 0, Pitch(Ball), 0, 0, AudioFade(Ball)
End Sub

'******************************************************
'   FLIPPER POLARITY AND RUBBER DAMPENER
'     SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset)        'Resize original array
  for x = 0 to aCount-1                'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

Function Distance2Obj(obj1, obj2)
  Distance2Obj = SQR((obj1.x - obj2.x)^2 + (obj1.y - obj2.y)^2)
End Function

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
  BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

'******************************************************
'   TRACK ALL BALL VELOCITIES
'     FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)  'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

'*****************
' Maths
'*****************
Dim PI: PI = 4*Atn(1)

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    end if
  ElseIf dx = 0 Then
    if dy = 0 Then
      Atn2 = 0
    else
      Atn2 = Sgn(dy) * pi / 2
    end if
  End If
End Function

Function ArcCos(x)
  If x = 1 Then
    ArcCos = 0/180*PI
  ElseIf x = -1 Then
    ArcCos = 180/180*PI
  Else
    ArcCos = Atn(-x/Sqr(-x * x + 1)) + 2 * Atn(1)
  End If
End Function

Function dArcSin(x)
  If X = 1 Then
    dArcSin = 90
  ElseIf x = -1 Then
    dArcSin = -90
  Else
    dArcSin = Atn(X / Sqr(-X * X + 1))*180/PI
  End If
End Function

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

Function min(a,b)
  if a > b then
    min = b
  Else
    min = a
  end if
end Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI /180
End Function

'******************************************************
'*   END VARI TARGET
'******************************************************

'**********Rubber Animations timer 50

sub dingwalla_hit
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



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Tstep, URstep


Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), slingR
'    DOF 105, DOFPulse
    vpmtimer.PulseSw(14)
    RSling.Visible = 0
    RSling1.Visible = 1
    slingR.objroty = -15
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:slingR.objroty = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), slingL
'    DOF 104, DOFPulse
    vpmtimer.pulsesw(13)
    LSling.Visible = 0
    LSling1.Visible = 1
    slingL.objroty = 15
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty=0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
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
    tmp = tableobj.y * 2 / SilverSlugger.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / SilverSlugger.width-1
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

'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping and Arch rolling Sounds)
'********************************************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

ReDim ArchRolling(tnob)
InitArchRolling

Dim ArchHit

Sub LowerArch_Hit
  Archhit = 1
End Sub

Sub NotOnArch_Hit
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
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 to UBound(BOT)

        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -2 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/400, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

'   ' play arch rolling sounds
'   If BallVel(BOT(b) ) > 1 AND ArchHit =1 Then
'     ArchRolling(b) = True
'     PlaySound("ArchHitA" & b),   0, (BallVel(BOT(b))/15)^5 * 1, AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
'     PlaySound("ArchRollA" & b), -1, (BallVel(BOT(b))/30)^5 * 1, AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
'   Else
'     If ArchRolling(b) = True Then
'     StopSound("ArchRollA" & b)
'     ArchRolling(b) = False
'     End If
'   End If

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
  PlaySound "fx_switch", 0, Vol(ActiveBall), Audiopan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
  PlaySound SoundFX("fx_target",DOFTargets), 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_DropTargets_Hit (idx)
  PlaySound "fx_drop1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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




'*****************************************
'           BALL SHADOW by ninnuzu,  BorgDog mod
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
    BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(b).X)/(SilverSlugger.Width/2))
    BallShadow(b).Y = BOT(b).Y + 10
    If BOT(b).Z > 0 and BOT(b).Z < 30 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub



Sub SilverSlugger_Paused:Controller.Pause = 1:End Sub
Sub SilverSlugger_unPaused:Controller.Pause = 0:End Sub

Sub SilverSlugger_Exit
  Controller.Games(cGameName).Settings.Value("sound")=1
  Controller.Stop
End Sub




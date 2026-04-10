'
'
'        .                                                                                                                                     .
'        :2B@B@B@B@@@B@B@@@@@@@B@B@@@B@B@B@B@B@B@@@@@B@                                           PB@B@@@B@B@B@B@@@B@@@@@B@B@B@B@@@B@B@B@B@B@Bki.
'            :2@B@B@B@B@B@B@B@@@@@B@B@@@B@B@B@B@B@B@B@Bi                v:     :v                 B@B@@@B@B@@@B@B@B@B@B@@@B@B@B@B@B@B@B@B@qi
'                iMB@B@B@B@B@@@B@B@B@B@B@B@B@B@B@B@B@B@B                @B     B@                8@@@@@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@v
'                   Y@B@B@@@B@BBBBB@BBBBBBBBB@B@B@B@B@@@Z               M@:::::@M               7@@@B@B@B@BBBBBBB@B@B@B@B@B@B@B@@@B@F
'                     i@@B@B@B@B@B@B@B@BBB@BBBBBBB@B@B@@@M.             @@@B@B@B@              5@B@BBBBB@BBBBBBB@B@B@B@B@B@B@B@@@@Y
'                       r@B@@@B@B@BBB@B@B@B@BBB@B@B@BBB@B@B@M5;,       .B@B@B@B@Br       .:JE@B@B@B@B@B@B@B@B@B@B@B@B@B@BBB@B@B@u
'                         MB@B@B@BBB@BBB@B@B@BBB@B@B@BBB@B@B@B@@@B@B@B@B@B@B@B@B@B@B@B@B@@@B@B@@@B@B@B@BBBBBBBBB@BBB@B@B@B@B@B@
'                          FB@B@B@BBB@B@B@B@BBB@B@B@B@BBB@B@B@B@B@B@B@@@B@B@B@B@B@@@B@@@B@B@B@B@B@B@B@B@B@B@B@B@B@B@BBB@B@B@BM
'                           SB@B@B@B@BBB@B@BBB@B@B@B@B@B@B@B@B@B@B@B@B@@@B@B@BBB@B@@@B@B@B@B@B@B@B@B@BBB@B@B@BBB@B@B@B@B@B@@B
'                            @B@B@B@B@B@BBB@B@B@B@B@BBBBB@B@B@B@B@B@B@B@BBB@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@BBB@BBBBB@B@B@B@B@
'                            ,@@@B@BBB@BBB@B@BBB@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@BBB@B@BBB@B@B@B@BBB@B@B@B@BBB@B@B@B@B@v
'                             B@BBB@B@BBB@B@B@B@B@B@B@B@B@B@B@BBB@BBB@B@B@B@BBB@B@BBB@BBB@B@B@BBB@BBB@B@B@BBB@B@B@B@B@B@B@@
'                             @B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@
'                             B@B@B@@@B@B@B@B@@@B@@@B@B@B@BBB@B@BBB@B@B@B@BBBBB@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@B@@@B@B@B@B
'                            :@B@B@@@@@B@B@B@B@B@@@B@@@@@@@B@B@BBB@B@B@B@BBB@B@B@B@B@B@B@B@B@@@B@B@B@B@B@@@B@B@B@@@B@B@@@B@Y
'                            @B@B@@@@@B@B@B@B@B@@B@@@B@B@B@B@B@B@B@B@B@B@BBBBB@B@B@BBB@B@B@B@B@B@B@B@B@M@B@B@@@@@B@B@B@B@B@@
'                                                     ,r18@B@B@@@B@B@B@B@B@B@B@BBB@B@B@B@B@B@B@Mk7:
'                                                           ,uB@B@B@B@B@B@B@BBB@B@B@B@B@B@BS:
'                                                               iO@B@B@B@B@B@B@B@B@B@@@Bv
'                                                                  JB@B@B@B@B@B@B@B@Bk
'                                                                    Y@B@B@B@B@B@B@F
'                                                                      OB@B@B@B@@@
'                                                                       JB@@@@@BP
'                                                                        iB@@@Bu
'                                                                         rB@B1
'                                                                          P@@
'                                                                           @
'
'
'
'
'Batman - The Dark Knight - IPDB No. 5307
'© Stern 2008
'Visual mod by Skitso
'VPX recreation by tom tower & ninuzzu
'Thanks to Lord Hiryu for the playfield texture
'Thanks to DJRobX for coding the Joker
'Thanks to the VPDev Team for the amazing VPX!
'This version has been given the nfozzy/roth physics treatment and fleep sound package by bord

Option Explicit
Randomize

'******************************************************************************************
'* CUSTOMIZABLE OPTIONS
'******************************************************************************************

'***********  Flippers Type   ('0=normal, 1=custom)
Const FlippersType = 0

'***********  Disable Cabinet Side Reflection ('0=no, 1=yes)
Const DisableCabSides = 0

'***********  Enable Custom Lighting  ('0=no, 1=yes)
Const LightingMod = 0

'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.8

'************************************************************************
'            INIT VPM
'************************************************************************

Const BallSize = 50
Const BallMass = 1

dim tablewidth: tablewidth = Table1.width
dim tableheight: tableheight = Table1.height

' Load controller.vbs to handle VPinMAME.Controller and B2S.Server
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Internal DMD in Desktop Mode, using a textbox (must be called before LoadVPM)
Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

' Standard Options
Const UseVPMModSol = 1, UseSolenoids = 1, UseLamps = 0, UseSync = 0, HandleMech = 1, SSolenoidOn = "fx_solenoid", SSolenoidOff = "", SCoin = ""

' Rom Name
Const cGameName = "bdk_294"

LoadVPM "03000000", "SAM.VBS", 3.50

'************************************************************************
'            INIT TABLE
'************************************************************************

Dim dtJoker, plungerIM, mechJoker, mechCrane

Dim xx, bmBall1, bmBall2, bmBall3, bmBall4

Sub Table1_Init
  vpmInit Me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Batman - The Dark Knight (Stern 2008)"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = DesktopMode 'Hide VPM DMD in Desktop Mode
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Nudging
    vpmNudge.TiltSwitch = -7
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(sw26,sw27,sw30,sw31,sw32)

  '************  Trough **************************
  Set bmBall4 = sw18.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set bmBall3 = sw19.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set bmBall2 = sw20.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set bmBall1 = sw21.CreateSizedballWithMass(Ballsize/2,Ballmass)

  Controller.Switch(18) = 1
  Controller.Switch(19) = 1
  Controller.Switch(20) = 1
  Controller.Switch(21) = 1

  ' Auto Plunger
  Set plungerIM = New cvpmImpulseP
  With plungerIM
    .InitImpulseP sw23, 65, 0.55
    .Switch 23
    .Random 0.05
    .CreateEvents "plungerIM"
  End With

  ' Joker Motor
  Set mechJoker = new cvpmMech
  With mechJoker
    .Mtype = vpmMechLinear + vpmMechOneDirSol + vpmMechFast
    .Sol1 = 26
    .Sol2 = 30
    .Length = 1500
    .Steps = 360
    .AddSw 52, 0, 0
    .AddSw 51, 178, 182
    .AddSw 50, 359, 359
    .Callback = GetRef("UpdateJokerMech")
    .Start
     End With

  ' Crane Motor
  Set mechCrane = new cvpmMech
  With mechCrane
    .Mtype = vpmMechLinear + vpmMechOneDirSol
    .Sol1 = 31
    .Sol2 = 28
    .Length = 110
    .Steps = 90
    .AddSw 56, 0, 1
    .AddSw 57, 30, 33
    .AddSw 58, 47, 50
    .AddSw 59, 59, 62
    .AddSw 60, 71, 74
    .AddSw 61, 89, 89
    .Callback = GetRef("UpdateCraneMech")
    .Start
     End With

  'Init Other Stuff
  InitLights:InitOptions
  Rails.visible = DesktopMode
  Dim i: For i=1 to 5:CraneHit(i).Collidable = 0:Next
' For i = 0 to 4 : dropping(i) = False : Next

  ' Fast Flips
  On Error Resume Next
  InitVpmFFlipsSAM
  If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5 rev3434"
  On Error Goto 0
End Sub

'************************************************************************
'             KEYS
'************************************************************************
dim ballinshooterlane

Sub sw23_hit
  ballinshooterlane=1
End Sub

Sub sw23_unhit
  ballinshooterlane=0
End Sub

Sub Table1_KeyDown(ByVal Keycode)
  If keycode = keyFront Then Controller.Switch(15) = 1    'tournament
  If KeyCode = PlungerKey Then Plunger.Pullback:SoundPlungerPull()
  If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()
  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25

    End Select
  End If
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
  If KeyCode = PlungerKey Then Plunger.Fire: If ballinshooterlane=1 then SoundPlungerReleaseBall() else SoundPlungerReleaseNoBall()
  If keycode = keyFront Then Controller.Switch(15) = 0    'tournament
  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_unPaused:Controller.Pause = False:End Sub
Sub Table1_exit()
  Controller.Pause = False:Controller.Stop
End Sub

'*****************************************
'         General Illumination
'*****************************************
Set GiCallBack = GetRef("UpdateGi")

Sub UpdateGi(nr,enabled)
  Dim ii
  Select Case nr
  Case 0
    If enabled Then
      DOF 103, DOFOn
      If LightingMod=1 Then
        Table1.ColorGradeImage = "ColorGradeBlue_on"
        LeftFlipperP.BlendDisableLighting = 1 : RightFlipperP.BlendDisableLighting = 1
        LeftFlipperLight.state = 1 : RightFlipperLight.state = 1
      Else
        Table1.ColorGradeImage = "ColorGradeEx_8"
      End If
      For each ii in GI:ii.state=1:Next
      For each ii in BWLamps:ii.IntensityScale=1:Next
      Reflect1.IntensityScale=1
      Reflect2.IntensityScale=1
      bulbBW.BlendDisableLighting = 8
    Else
      DOF 103, DOFOff
      If LightingMod=1 Then
        Table1.ColorGradeImage = "ColorGradeBlue_off"
        LeftFlipperP.BlendDisableLighting = 0 : RightFlipperP.BlendDisableLighting = 0
        LeftFlipperLight.state = 0 : RightFlipperLight.state = 0
      Else
        Table1.ColorGradeImage = "ColorGradeEx_1"
      End If
      For each ii in GI:ii.state=0:Next
      For each ii in BWLamps:ii.IntensityScale=0:Next
      Reflect1.IntensityScale=0
      Reflect2.IntensityScale=0
      bulbBW.BlendDisableLighting = 0
    End If
  End Select
End Sub

'*****************************************
'       Lights Mapping
'*****************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

Sub InitLights
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state
        FadingLevel(x) = 0       ' current light fading level
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
  LampTimer.Interval = -1
  LampTimer.Enabled = 1
End Sub

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps
  'Inserts
  FadeLamp 3,  l3
  FadeLamp 4,  l4
  FadeLamp 5,  l5
  FadeLamp 6,  l6
  FadeLamp 7,  l7
  FadeLamp 8,  l8
  FadeLamp 9,  l9
  FadeLamp 10,  l10
  FadeLamp 11,  l11
  FadeLamp 12,  l12
  FadeLamp 13,  l13
  FadeLamp 14,  l14
  FadeLamp 15,  l15
  FadeLamp 16,  l16
  FadeLamp 17,  l17
  FadeLamp 18,  l18
  FadeLamp 19,  l19
  FadeLamp 20,  l20
  FadeLamp 21,  l21
  FadeLamp 22,  l22
  FadeLamp 23,  l23
  FadeLamp 24,  l24
  FadeLamp 25,  l25
  FadeLamp 26,  l26
  FadeLamp 27,  l27
  FadeLamp 28,  l28
  FadeLamp 29,  l29
  FadeLamp 31,  l31
  FadeLamp 32,  l32
  FadeLamp 33,  l33
  FadeLamp 34,  l34
  FadeLamp 35,  l35
  FadeLamp 36,  l36
  FadeLamp 37,  l37
  FadeLamp 38,  l38
  FadeLamp 39,  l39
  FadeLamp 40,  l40
  FadeLamp 41,  l41
  FadeLamp 42,  l42
  FadeLamp 43,  l43
  FadeLamp 44,  l44
  FadeLamp 45,  l45
  FadeLamp 47,  l47
  FadeLamp 48,  l48
  FadeLamp 49,  l49
  FadeLamp 50,  l50
  FadeLamp 51,  l51
  FadeLamp 52,  l52
  FadeLamp 53,  l53
  FadeLamp 54,  l54
  FadeLamp 55,  l55
  FadeLamp 56,  l56
  FadeLamp 57,  l57
  FadeLamp 58,  l58
  MiniPFLamps 59,  l59    'minipf BL
  FadePrim 59,  bulbBL, 8   'minipf BL
  FadeLamp 60,  l60     'left bumper
  FadePrim 60,  bulbl60,20  'left bumper
  FadeLamp 61,  l61     'right bumper
  FadePrim 61,  bulbl61,20  'right bumper
  MiniPFLamps 62,  l62    'minipf TL
  FadePrim 62,  bulbTL, 8   'minipf TL
  FadeLamp 63,  l63
  FadeLamp 64,  l64
  FadeLamp 65,  l65
  FadeLamp 66,  l66
  FadeLamp 67,  l67
  MiniPFLamps 68,  l68    'minipf BR
  FadePrim 68,  bulbBR, 8   'minipf BR
  FadeLamp 69,  l69
  FadeLamp 70,  l70
  FadeLamp 71,  l71
  FadeLamp 72,  l72
  FadeLamp 73,  l73
  FadeLamp 74,  l74
  MiniPFLamps 75,  l75    'minipf TR
  FadePrim 75,  bulbTR, 8   'minipf TR
  FadeLamp 76,  l76
  FadeLamp 77,  l77
  FadeLamp 78,  l78a      'batmobile
  FadeLamp 78,  l78b      'batmobile
  FadeLamp 79,  l79a      'joker
  FadeLamp 79,  l79b      'joker
  FadeLamp 80,  l80a      'scarecrow
  FadeLamp 80,  l80b      'scarecrow
  FadeLamp 86,  l86
  FadeLamp 87,  l87
  FadeLamp 88,  l88

  'Flashers
  FadeLamp 179,  F19a
  FadeLamp 179,  F19b
  FadeLamp 181,  F21
  FadeLamp 182,  F22
  FadeLamp 182,  F22a
  FadeLamp 182,  F22b
  FadeLamp 182,  F22c
  FadeLamp 183,  F23
  FadeLamp 185,  F25a
  FadeLamp 185,  F25b
  FadeLamp 187,  F27
  FadeLamp 187, F27a
  FadePrim 187,  LDome, 10
  FadeLamp 189,  F29
  FadeLamp 189,  F29a
  FadePrim 189,  RDome, 10
  FadeLamp 192,  F32
  FadeLamp 192,  F32a
End Sub

' Not Modulated lights and flashers
Sub FadeLamp(nr, object)
  If TypeName(object) = "Light" Then
    Object.State = LampState(nr)
  End If
  If TypeName(object) = "Flasher" Then
    Object.IntensityScale = LampState(nr)
  End If
End Sub

Sub FadePrim(nr, object, factor)
  Object.BlendDisableLighting = factor * LampState(nr)
End Sub

Sub SetLamp(nr, enabled)
    If enabled Then
    LampState(nr) = 1
  Else
    LampState(nr) = 0
  End If
End Sub

Sub MiniPFLamps(nr, object)
  object.IntensityScale = LampState(nr)/(LampState(59) + LampState(62) + LampState(68) + LampState(75) + 0.05)
End Sub

' Modulated lights and flashers
Sub FadeModLamp(nr, object)
  Object.IntensityScale = FadingLevel(nr)/255
End Sub

Sub FadeModPrim(nr, object, factor)
  Object.BlendDisableLighting = factor * FadingLevel(nr)/255
End Sub

Sub SetModLamp(nr, value)
  FadingLevel(nr) = value
End Sub

'************************************************************************
'           SOLENOIDS MAP
'************************************************************************

SolCallBack(1) = "SolRelease"           'Trough-Up Kicker
SolCallBack(2) = "SolAutoPlungerIM"         'AutoLaunch
SolCallback(3) = "SolTEject"            'Top Right Eject
SolCallBack(4) = "SolJokerEject"          'Joker Lockup
SolCallback(5) = "JokerDropUp"            'Joker Drop Target Up
SolCallback(6) = "SolVUK"             'Scarecrow VUK
SolCallback(7) = "JokerDropDown"          'Joker Drop Target Down
SolCallback(8) = "SolShaker"            'Shaker Motor
'SolCallback(9) = ""              'Left Bumper
'SolCallback(10)= ""              'Right Bumper
'SolCallback(11)= ""              'Bottom Bumper
SolCallBack(12)= "vpmSolGate LGate, SoundFX(""ElGate"",DOFContactors)," 'Left Control Gate
SolCallBack(13)= "SolBatRamp"           'Batmobile Ramp Down
SolCallBack(14)= "vpmSolGate RGate, SoundFX(""ElGate"",DOFContactors)," 'Right Control Gate
SolCallback(15)= "SolLFlipper"            'Left Flipper
SolCallback(16)= "SolRFlipper"            'Right Flipper
'SolCallback(17)= ""              'Left Sling
'SolCallback(18)= ""              'Right Sling
SolCallBack(19)= "SetLamp 179,"           'Flasher:ScareCrow Home Insert
SolCallBack(21)= "SetLamp 181,"           'Flasher:BackPanel
SolCallBack(22)= "SetLamp 182,"           'Flasher:Joker (x3)
SolCallBack(23)= "SetLamp 183,"           'Flasher:Scarecrow
SolCallBack(24)= "solKnocker"         'Knocker
SolCallBack(25)= "SetLamp 185,"           'Flasher:Pop Bumpers  (x3)
'SolCallBack(26)= ""              'Joker Motor
SolCallBack(27)= "SetLamp 187,"           'Flasher:Left SlingShot
'SolCallBack(28)= ""              'ScareCrow Motor Relay
SolCallBack(29)= "SetLamp 189,"           'Flasher:Right SlingShot
'SolCallBack(30)= ""              'Joker Motor Relay
'SolCallBack(31)= ""              'ScareCrow Motor
SolCallback(32)= "SetLamp 192,"           'Flasher:BatMobile Crash (x2)

'******************************************************
'         KNOCKER
'******************************************************
Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub

'************************************************************************
'           FLIPPERS
'************************************************************************
'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
  Next

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -5.5
  AddPt "Polarity", 2, 0.4, -5.5
  AddPt "Polarity", 3, 0.6, -5.0
  AddPt "Polarity", 4, 0.65, -4.5
  AddPt "Polarity", 5, 0.7, -4.0
  AddPt "Polarity", 6, 0.75, -3.5
  AddPt "Polarity", 7, 0.8, -3.0
  AddPt "Polarity", 8, 0.85, -2.5
  AddPt "Polarity", 9, 0.9,-2.0
  AddPt "Polarity", 10, 0.95, -1.5
  AddPt "Polarity", 11, 1, -1.0
  AddPt "Polarity", 12, 1.05, -0.5
  AddPt "Polarity", 13, 1.1, 0
  AddPt "Polarity", 14, 1.3, 0

  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

'******************************************************
'         FLIPPERS
'******************************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    lf.fire 'LeftFlipper.RotateToEnd
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
        LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
 End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        rf.fire 'RightFlipper.RotateToEnd
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
    Else
        RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
 End Sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, these seems obsolete
' Sub LeftFlipper_Collide(parm)
'   LeftFlipperCollide parm
' End Sub
'
' Sub RightFlipper_Collide(parm)
'   RightFlipperCollide parm
' End Sub
'******************************************************
'     FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
  Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
  Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

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
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        RemoveBall aBall
        exit Sub
      end if

      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)            'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
   :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
        'playsound "knocker"
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'   FLIPPER POLARITY AND RUBBER DAMPENER
'     SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
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

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
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

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

' Used for drop targets
Function Atn2(dy, dx)
  dim pi
  pi = 4*Atn(1)

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

' Used for drop targets
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

' Used for drop targets and flipper tricks
Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

'******************************************************
'     FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
end sub


dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = .8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.855
Const EOSReturn = 0.025
Const SOSRS = 0.030

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity
  Flipper.Return = FReturn

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST*EOSReturn/FReturn


  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim BOT, b
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3*Dir
      Flipper.Elasticity = FElasticity * SOSEM
      Flipper.return = SOSRS
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If

  End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
  Dim LiveCatchBounce                             'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.5 Then            'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)  'Partial catch when catch happens a bit late
    end If

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub


'************************************************************************
'           BALL TROUGH
'************************************************************************

Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    If BsTrough.Balls Then vpmTimer.PulseSw 22
  End If
End Sub

'************************************************************************
'            AUTOPLUNGER
'************************************************************************

Sub SolAutoPlungerIM(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
    PlaySoundAt SoundFX("AutoPlunger",DOFContactors), Plunger
  End If
End Sub

'************************************************************************
'            SCARECROW VUK
'************************************************************************

Sub SolVUK(enabled)
  If Enabled Then
    sw12.kick 0, 35, 1.56
    If BallInSaucer = True Then
      PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, sw12
    Else
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, sw12
    End If
  End If
End Sub

dim ballinsaucer

sub sw12_hit
  Controller.switch(12) = 1
  SoundSaucerLock
  ballinsaucer=true
end sub

sub sw12_unhit
  Controller.switch(12) = 0
  ballinsaucer=False
end sub

'************************************************************************
'            UPPER SAUCER
'************************************************************************

Sub SolTEject(enabled)
  If Enabled Then
    sw44.kick -90, 9, 0
    If BallInSaucer2 = True Then
      PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, sw44
    Else
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, sw44
    End If
  End If
End Sub

dim ballinsaucer2

sub sw44_hit
  Controller.switch(44) = 1
  SoundSaucerLock
  ballinsaucer2=true
end sub

sub sw44_unhit
  Controller.switch(44) = 0
  ballinsaucer2=False
end sub

'************************************************************************
'            SHAKER MOTOR
'************************************************************************
Dim bou,brake,perc
Dim bou2,brake2,perc2

Sub SolShaker(enabled)
  If Enabled Then
    ShakeJoker : ShakeBall
    PlaySoundAt SoundFX("ShakerPulse",DOFShaker), Primitive12
  End If
End Sub

Sub ShakeJoker
  perc=3:ShakeJokerTimer.enabled=1
End Sub

Sub ShakeBall
  perc2=3:ShakeTimer.enabled=1
End Sub

Sub ShakeJokerTimer_timer()
  bou=bou+0.4:brake=brake+0.02
  Joker.rotx=sin(bou)*(perc-(brake*(perc/6)))
  If (perc-(brake*(perc/6)))<0 Then Me.Enabled = 0 :bou=0 :brake=0 :perc=0
End Sub

Sub ShakeTimer_timer()
  bou2=bou2+0.4:brake2=brake2+0.02
  CraneS.transX=sin(bou2)*(perc2-(brake2*(perc2/6)))
  If (perc2-(brake2*(perc2/6)))<0 Then Me.Enabled = 0 :bou2=0 :brake2=0 :perc2=0
End Sub

'************************************************************************
'           BATMOBILE RAMP
'************************************************************************

Const Pi=3.1415926535
Dim fakeBall, RampDir, RampPos
ReDim BatRampRadius(4), dropping(4)

Sub sw64_Hit: If ActiveBall.VelY<0 Then Controller.Switch(64) = 1 :End If:SoundSaucerLock:End Sub
Sub Trigger1_hit: PlaysoundAt "fx_rrturn", Trigger1: If ActiveBall.VelX<0 And ActiveBall.VelX>-5 Then ActiveBall.VelX = -8 :End If : End Sub

Sub SolBatRamp(enabled)
  If Enabled Then
    CheckBalls
    RampPos=10:RampDir=-1: MoveBatRamp.Enabled=1
    PlaysoundAt SoundFX("fx_solenoid", DOFContactors), BatMobile
  Else
    Dim b:For b = 0 to 3
      If dropping(b) = True Then
        dropping(b) = False
      End If
    Next
    RampPos=0:RampDir=1: MoveBatRamp.Enabled=1
    PlaysoundAt SoundFX("fx_solenoid", DOFContactors), BatMobile
  End If
End Sub

Sub MoveBatRamp_timer
  RampPos = RampPos + RampDir
  If RampDir = -1 And RampPos<0 Then RampPos=0  : Me.Enabled=0 : DropTheBall 1
  If RampDir = 1 And RampPos>10 Then RampPos=10 : Me.Enabled=0 : DropTheBall 0
  UpperTrack.RotX = RampPos
  LowerTrack.RotX = RampPos
  DecalDropRamp.RotX = RampPos
  DecalSide.RotX = RampPos
  BatMobile.RotX = RampPos
    Dim BOT, b
    BOT = GetBalls
  For b = 0 to 3
    If dropping(b) = True Then
      BOT(b).Y = LowerTrack.Y - BatRampRadius(b) * cos (RampPos * Pi/180)
      BOT(b).Z = 135 + BatRampRadius(b) * sin (RampPos * Pi/180) + 25
    End If
  Next
End Sub

Sub CheckBalls()
    Dim BOT, b
    BOT = GetBalls
  ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' check if there are balls on the batramp
    For b = 0 to UBound(BOT)
        If BOT(b).X< 85 and BOT(b).Z> 50 Then
      dropping(b) = True
      BatRampRadius(b) = SQR((LowerTrack.X - BOT(b).X)^2 + (LowerTrack.Y - BOT(b).Y)^2)
        End If
    Next
End Sub

Sub DropTheBall(n)
  Select Case n
    Case 0
      RampUp.collidable=1
      RampDown.collidable=0
      Set fakeBall = BallDestroy.createball : fakeBall.visible = 0 : BallDestroy.kick 0,21
      MoveBatMobile.enabled= 1
      PlaySoundAt "BallRoll_0", BatMobile
    Case 1
      Dim BOT, b
      BOT = GetBalls
      For b = 0 to 3
        If dropping(b) = True Then
          BOT(b).VelZ = 10
        End If
      Next
      Controller.Switch(64) = 0
      RampUp.collidable=0
      RampDown.collidable=1
      Set fakeBall = BallMake.createball : fakeBall.visible = 0 : BallMake.kick 180,1
      MoveBatMobile.enabled= 1
  End Select
End Sub

Sub MoveBatMobile_timer
  BatMobile. TransY = -BallMake.Y + fakeBall.Y
End Sub

Sub BallMake_hit : MoveBatMobile.Enabled=0 :  Me.DestroyBall : fakeBall = Empty : StopSound "BallRoll_0" : End Sub
Sub BallDestroy_Hit : MoveBatMobile.Enabled=0 :  Me.DestroyBall : fakeBall = Empty : vpmtimer.pulsesw 47: StopSound "BallRoll_0":End Sub

'************************************************************************
'             JOKER MOTOR
'************************************************************************

Dim InPos, OutPos, CurMechPos: InPos = 0 : OutPos = 0
Dim JokerRadius:JokerRadius = SQR((cilin.X - joker.X)^2 + (cilin.Y - joker.Y)^2)

Sub UpdateJoker_timer
  if InPos < CurMechPos then
    InPos = InPos + 1
    if InPos > CurMechPos then InPos = CurMechPos
  elseif InPos > CurMechPos Then
    InPos = InPos - 1
    if InPos < CurMechPos then InPos = CurMechPos
  else
    me.Enabled=false
  end if
  if InPos <=0 then outpos = InPos
  UpdateObjects
End Sub

Sub UpdateJokerMech(aCurrPos,aSpeed,aLastPos)
  CurMechPos = aCurrPos -180
  UpdateJoker.Enabled = 1
End Sub

Sub UpdateObjects
  cilin.RotZ = InPos
  cilout.RotZ = OutPos
  Joker.X = cilin.X - JokerRadius * sin (InPos * Pi/180)
  Joker.Y = cilin.Y + JokerRadius * cos (InPos * Pi/180)
  Joker.RotZ = InPos
End Sub

'************************************************************************
'             CRANE MOTOR
'************************************************************************

Dim CranePos, MotorSnd, CurMech1Pos : MotorSnd = 0 : CranePos = 160
Dim CraneRadius:CraneRadius = SQR((crane.X - joker.X)^2 + (crane.Y - joker.Y)^2)

Sub UpdateCrane_timer
  if CranePos < CurMech1Pos then
    CranePos = CranePos + 0.5
    if CranePos > CurMech1Pos then CranePos = CurMech1Pos
  elseif CranePos > CurMech1Pos Then
    CranePos = CranePos - 0.5
    if CranePos < CurMech1Pos then CranePos = CurMech1Pos
  else
    me.Enabled=false : ShakeBall : StopSound "IdolMotor"
  end if
  UpdateParts
End Sub

Sub UpdateCraneMech(aCurrPos,aSpeed,aLastPos)
  CurMech1Pos = 250 - aCurrPos
  UpdateCrane.Enabled = 1
  If aSpeed=0 Then
    StopSound "Motor":MotorSnd=0
  Else
    If MotorSnd=0 Then PlaySoundAtVolLoops SoundFX("Motor", DOFGear), crane, 1, -1: MotorSnd=1
  End If
End Sub

Sub UpdateParts
  If CranePos > (250 -2) Then
    CraneHit(5).Collidable = 1
  ElseIf CranePos<(250 -28) AND CranePos > (250 -35) Then
    CraneHit(4).Collidable = 1
  ElseIf CranePos<(250 -45) AND CranePos > (250 -52) Then
    CraneHit(3).Collidable = 1
  ElseIf CranePos<(250 -57) AND CranePos > (250 -64) Then
    CraneHit(2).Collidable = 1
  ElseIf CranePos<(250 -70) AND CranePos > (250 -77) Then
    CraneHit(1).Collidable = 1
  ElseIf CranePos<(250 -86) Then
    CraneHit(0).Collidable = 1
  Else
    Dim i: For each i in CraneHit:i.Collidable=0 : Next
  End If
    crane.rotz=CranePos
    craneB.rotz=CranePos
    craneM.rotz=CranePos
    craneS.rotz=CranePos
End Sub

Sub CraneHit_Hit(idx)
  vpmtimer.Pulsesw 85
  ShakeBall
  PlaySoundAt "Ball_Collide_1", craneM
End Sub

'************************************************************************
'               SWITCHES
'************************************************************************

'******************************************************
'           TROUGH
'******************************************************

Sub sw19_Hit():Controller.Switch(19) = 1:UpdateTrough:End Sub
Sub sw19_UnHit():Controller.Switch(19) = 0:UpdateTrough:End Sub
Sub sw20_Hit():Controller.Switch(20) = 1:UpdateTrough:End Sub
Sub sw20_UnHit():Controller.Switch(20) = 0:UpdateTrough:End Sub
Sub sw21_Hit():Controller.Switch(21) = 1:UpdateTrough:End Sub
Sub sw21_UnHit():Controller.Switch(21) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw21.BallCntOver = 0 Then sw20.kick 60, 9
  If sw20.BallCntOver = 0 Then sw19.kick 60, 9
  If sw19.BallCntOver = 0 Then sw18.kick 60, 20
  Me.Enabled = 0
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub sw18_Hit() 'Drain
  UpdateTrough
  Controller.Switch(18) = 1
  RandomSoundDrain(sw18)
End Sub

Sub sw18_UnHit()  'Drain
  Controller.Switch(18) = 0
End Sub

Sub SolRelease(enabled)
  If enabled Then
    If sw21.BallCntOver = 0 Then
      PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw21
    Else
      RandomSoundBallRelease(sw21)
      vpmTimer.PulseSw 22
    End If
    sw21.kick 60, 9
  End If
End Sub

Sub sw46_hit
  SoundSaucerLock
  Controller.switch(46) = True
End Sub

Sub sw46_unhit
  Controller.switch(46) = False
End Sub

Sub SolJokerEject(Enabled)
  If Enabled then
  PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, sw46
    sw46.kick -13, 45, 0
  End If
End Sub

'StandUp Targets
Sub sw1_hit:vpmtimer.PulseSw 1: End Sub
Sub sw2_hit:vpmtimer.PulseSw 2: End Sub
Sub sw3_hit:vpmtimer.PulseSw 3: End Sub
Sub sw4_hit:vpmtimer.PulseSw 4: End Sub
Sub sw5_hit:vpmtimer.PulseSw 5: End Sub
Sub sw6_hit:vpmtimer.PulseSw 6: End Sub
Sub sw11_hit:vpmtimer.PulseSw 11: End Sub
Sub sw14_hit:vpmtimer.PulseSw 14: End Sub
Sub sw36_hit:vpmtimer.PulseSw 36: End Sub
Sub sw37_hit:vpmtimer.PulseSw 37: End Sub
Sub sw38_hit:vpmtimer.PulseSw 38: End Sub
Sub sw48_hit:vpmtimer.PulseSw 48: End Sub
Sub sw49_hit:vpmtimer.PulseSw 49: End Sub

'Rollovers
Sub sw7_hit:Controller.switch(7) = 1:End Sub
Sub sw7_unhit:Controller.switch(7) = 0: End Sub

Sub sw8_hit:Controller.switch(8) = 1:End Sub
Sub sw8_unhit:Controller.switch(8) = 0: End Sub

Sub sw9_hit:Controller.switch(9) = 1:End Sub
Sub sw9_unhit:Controller.switch(9) = 0: End Sub

Sub sw24_hit:Controller.switch(24) = 1:End Sub
Sub sw24_unhit:Controller.switch(24) = 0: End Sub

Sub sw25_hit:Controller.switch(25) = 1:End Sub
Sub sw25_unhit:Controller.switch(25) = 0: End Sub

Sub sw28_hit:Controller.switch(28) = 1:End Sub
Sub sw28_unhit:Controller.switch(28) = 0: End Sub

Sub sw29_hit:Controller.switch(29) = 1:End Sub
Sub sw29_unhit:Controller.switch(29) = 0: End Sub

Sub sw33_hit:Controller.switch(33) = 1:End Sub
Sub sw33_unhit:Controller.switch(33) = 0: End Sub

'Center Ramp Enter
Sub sw13_hit:vpmtimer.pulsesw 13:End Sub

'Bumpers
Sub sw30_hit:vpmtimer.pulsesw 30:RandomSoundBumperTop sw30:End Sub
Sub sw31_hit:vpmtimer.pulsesw 31:RandomSoundBumperTop sw31:End Sub
Sub sw32_hit:vpmtimer.pulsesw 32:RandomSoundBumperMiddle sw32:End Sub

'Spinners
Sub sw34_spin:vpmtimer.pulsesw 34: PlaySoundAt "fx_spinner", sw34:End Sub
Sub sw39_spin:vpmtimer.pulsesw 39: PlaySoundAt "fx_spinner", sw39:End Sub

'Slingshots
Dim LStep, RStep

Sub sw26_slingshot
  vpmTimer.PulseSw 26
  RandomSoundSlingshotLeft sling1
  LSling.Visible = 0: LSling2.Visible = 1: sling1.TransZ = -20: LStep = 0
  Me.TimerInterval = 50: Me.TimerEnabled = 1
End Sub

Sub sw27_slingshot
  vpmTimer.PulseSw 27
  RandomSoundSlingshotRight sling2
  RSling.Visible = 0: RSling2.Visible = 1: sling2.TransZ = -20: RStep = 0
  Me.TimerInterval = 50: Me.TimerEnabled = 1
End Sub

Sub sw26_Timer
  Select Case LStep
    Case 3:LSLing2.Visible = 0:LSLing1.Visible = 1:sling1.TransZ = -10
    Case 4:LSLing1.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:Me.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub

Sub sw27_Timer
  Select Case RStep
    Case 3:RSLing2.Visible = 0:RSLing1.Visible = 1:sling2.TransZ = -10
    Case 4:RSLing1.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:Me.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub

'Left Loop
Sub sw35_hit
vpmtimer.PulseSw 35
If ActiveBall.VelY>20 Then ActiveBall.VelY = 18
End Sub

'Right Loop
Sub sw41_hit
vpmtimer.PulseSw 41
If ActiveBall.VelY>20 Then ActiveBall.VelY = 18
End Sub

'VUK Exit
Sub sw42_hit:vpmtimer.pulsesw 42:End Sub

'Drop Target
Sub sw45_Hit:DTHit 25:End Sub

Sub JokerDropUp(Enabled)
     if enabled then
'          PlaySoundAt SoundFX(DTResetSound,DOFContactors), sw45
          DTRaise 45
     end if
End Sub

Sub JokerDropDown(enabled)
     If enabled Then
          DTDrop 45
     End If
End Sub

'******************************************************
'   DROP TARGETS INITIALIZATION
'******************************************************

'Define a variable for each drop target
Dim DT45

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:      primary target wall to determine drop
' secondary:      wall used to simulate the ball striking a bent or offset target after the initial Hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             rotz must be used for orientation
'             rotx to bend the target back
'             transz to move it up and down
'             the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
' Values for annimate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target

' Joker Target
DT45 = Array(sw45, sw45y, psw45, 45, 0)

'Add all the Drop Target Arrays to Drop Target Animation Array
' DTAnimationArray = Array(DT1, DT2, ....)
Dim DTArray
DTArray = Array(DT45)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110       'in milliseconds
Const DTDropUpSpeed = 40      'in milliseconds
Const DTDropUnits = 44      'VP units primitive drops
Const DTDropUpUnits = 10      'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8         'max degrees primitive rotates when hit
Const DTDropDelay = 20      'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40     'time in milliseconds before target drops back to normal up position after the solendoid fires to raise the target
Const DTBrickVel = 0        'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 1     'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DTDrop"    'Drop Target Drop sound
Const DTResetSound = "DTReset"  'Drop Target reset sound

Const DTMass = 0.2        'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

  PlaySoundAtVol  DTHitSound, Activeball, Vol(Activeball)*22.5
  DTArray(i)(4) =  DTCheckBrick(Activeball,DTArray(i)(2))
  If DTArray(i)(4) = 1 or DTArray(i)(4) = 3 Then
    DTBallPhysics Activeball, DTArray(i)(2).rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) = -1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 to uBound(DTArray)
    If DTArray(i)(3) = switch Then DTArrayID = i:Exit Function
  Next
End Function


sub DTBallPhysics(aBall, angle, mass)
  dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
  calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

  aBall.velx = calc1 * cos(rangle) + calc2
  aBall.vely = calc1 * sin(rangle) + calc3
End Sub

'Add Timer name DTAnim to editor to handle drop target animations
DTAnim.interval = 10
DTAnim.enabled = True

Sub DTAnim_Timer()
  DoDTAnim
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)

  If perpvel <= 0 or perpvelafter >= 0 Then
    DTCheckBrick = 0
  ElseIf DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
    DTCheckBrick = 3
  Else
    DTCheckBrick = 1
  End If
End Function


Sub DoDTAnim()
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray(i)(4) = DTAnimate(DTArray(i)(0),DTArray(i)(1),DTArray(i)(2),DTArray(i)(3),DTArray(i)(4))
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch,  animate)
  dim transz
  Dim animtime, rangle

  rangle = prim.rotz * 3.1416 / 180

  DTAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If animate = 1 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    DTAnimate = 1
    Exit Function
  elseif animate = 1 and animtime > DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    animate = 2
    PlaySoundAt SoundFX(DTDropSound,DOFDropTargets),prim
  End If

  if animate = 2 Then
    transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
    if prim.transz > -DTDropUnits  Then
      prim.transz = transz
    end if

    prim.rotx = DTMaxBend * cos(rangle)/2
    prim.roty = DTMaxBend * sin(rangle)/2

    if prim.transz <= -DTDropUnits Then
      prim.transz = -DTDropUnits
      secondary.collidable = 0
      controller.Switch(Switch) = 1
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    end If
  End If

  If animate = 3 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
  elseif animate = 3 and animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  if animate = -1 Then
    transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

    If prim.transz = -DTDropUnits Then
      Dim BOT, b
      BOT = GetBalls

      For b = 0 to UBound(BOT)
        If InRect(BOT(b).x,BOT(b).y,prim.x-25,prim.y-10,prim.x+25, prim.y-10,prim.x+25,prim.y+25,prim.x -25,prim.y+25) Then
          BOT(b).velz = 20
        End If
      Next
    End If

    if prim.transz < 0 Then
      prim.transz = transz
    elseif transz > 0 then
      prim.transz = transz
    end if

    if prim.transz > DTDropUpUnits then
      DTAnimate = -2
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1
    controller.Switch(Switch) = 0

  End If

  if animate = -2 and animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
    if prim.transz < 0 then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    end If
  End If
End Function



'Mini Pf Rollovers
Sub sw53_hit:Controller.switch(53) = 1: PlaySoundAt "fx_sensor",ActiveBall: sw53p.RotX = -15 : End Sub
Sub sw53_unhit:Controller.switch(53) = 0: sw53p.RotX = 0 : End Sub

Sub sw54_hit:Controller.switch(54) = 1: PlaySoundAt "fx_sensor",ActiveBall: sw54p.RotX = -15 : End Sub
Sub sw54_unhit:Controller.switch(54) = 0: sw54p.RotX = 0 : End Sub

Sub sw55_hit:Controller.switch(55) = 1: PlaySoundAt "fx_sensor",ActiveBall: sw55p.RotX = -15 : End Sub
Sub sw55_unhit:Controller.switch(55) = 0: sw55p.RotX = 0 : End Sub

Sub sw63_hit:Controller.switch(63) = 1: PlaySoundAt "fx_sensor",ActiveBall: sw63p.RotX = -15 : End Sub
Sub sw63_unhit:Controller.switch(63) = 0: sw63p.RotX = 0 : End Sub


' *********************************************************************
'             Real Time Updates
' *********************************************************************

Sub GameTimer_timer
  RollingSoundUpdate
  BallShadowUpdate
  LeftFlipperP.ObjRotZ = LeftFlipper.currentangle
  RightFlipperP.ObjRotZ = RightFlipper.currentangle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  sw13p.rotX = -sw13.currentangle
  sw34p.rotX = -sw34.currentangle
  sw35p.rotX = sw35.currentangle
  sw39p.rotX = -sw39.currentangle
  sw41p.rotY = sw41.currentangle
  sw42p.rotX = sw42.currentangle
  Gate1P.rotY = -Gate1.currentangle
  Gate2P.rotY = -Gate2.currentangle
  Gate3P.rotY = Gate3.currentangle
  LGateP.rotY = Lgate.currentangle
  RGateP.rotY = Rgate.currentangle *2/3
End Sub

Const fakeballs = 0         ' number of balls created on table start (rolling sound and ballshadow will be skipped)
ReDim BallShadow(tnob-fakeballs-1)

'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 to tnob
    rolling(i) = False
  Next
End Sub

Sub RollingSoundUpdate()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    '***Ball Drop Sounds***
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
  Next
End Sub

' *********************************************************************
'           Ramp Sounds
' *********************************************************************

'right ramp
Sub RWREnter_hit():PlaySoundAtBall "fx_metalrolling":End Sub
Sub RWRExit_Hit:StopSound "fx_metalrolling":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub RWRExit_timer:Me.TimerEnabled=0:PlaySoundAt "fx_wireramp_exit", RWRExit::End Sub
'batramp
Sub BREnter_hit():PlaySoundAtBall "fx_metalrolling":End Sub
Sub BRExit_Hit:StopSound "fx_metalrolling":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub BRExit_timer:Me.TimerEnabled=0:PlaySoundAt "fx_wireramp_exit", BRExit::End Sub
'left ramp
Sub LWREnter_hit():PlaySoundAtBall "fx_metalrolling":End Sub
Sub LWRExit_Hit:StopSound "fx_metalrolling":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub LWRExit_timer:Me.TimerEnabled=0:PlaySoundAt "fx_wireramp_exit", LWRExit::End Sub
'right ramp
Sub RREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAtBall "fx_lrenter":PlaySoundAtBall "fx_metalrolling":End If:End Sub     'ball is going up
Sub RREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_lrenter":StopSound "fx_metalrolling":End If:End Sub   'ball is going down
Sub Gate3_hit(): StopSound "fx_metalrolling":End Sub

' **********************BALL SHADOW**********************
InitBallShadow

Sub InitBallShadow
  Dim i
  For i=0 to tnob-fakeballs-1
    ExecuteGlobal "Set BallShadow(" & i & ") = BallShadow" & (i+1) & ":"
  Next
End Sub

Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' hide shadow of deleted balls
  If UBound(BOT)<(tnob-1) Then
    For b = (UBound(BOT) + 1) to (tnob-1)
    BallShadow(b-fakeballs).visible = 0
    Next
  End If
  ' exit the Sub if no balls on the table
    If UBound(BOT) = fakeballs-1 Then Exit Sub

  ' render the shadow for each ball
    For b = fakeballs to UBound(BOT)
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b-fakeballs).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) + 10
    Else
      BallShadow(b-fakeballs).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) - 10
    End If
      BallShadow(b-fakeballs).Y = BOT(b).Y + 20
    If BOT(b).Z > 20 Then
      BallShadow(b-fakeballs).visible = 1
    Else
      BallShadow(b-fakeballs).visible = 0
    End If
  Next
End Sub

' *********************************************************************
'             Table Options
' *********************************************************************

Sub InitOptions
'flippers
  Select Case FlippersType
    Case 0
    LeftFlipperP.image = "flippers" : RightFlipperP.image = "flippers"
    Case 1
    LeftFlipperP.image = "BDK flippers" : RightFlipperP.image = "BDK flippers"
  End Select
'side reflections
  Select Case DisableCabSides
    Case 0
    Reflect1.visible = 1 : Reflect2.visible = 1
    Case 1
    Reflect1.visible = 0 : Reflect2.visible = 0
  End Select
'custom lighting
  Select Case LightingMod
    Case 0
    Table1.ColorGradeImage = "ColorGradeEx_8"
    LDome.image = "Dome blue" : RDome.image = "Dome blue"
    LeftFlipperP.BlendDisableLighting = 0 : RightFlipperP.BlendDisableLighting = 0
    LeftFlipperLight.state = 0 : RightFlipperLight.state = 0
    Reflect1.imageA = "CabL" : Reflect1.imageB = "CabL"
    Reflect2.imageA = "CabR" : Reflect2.imageB = "CabR"
    Case 1
    Dim ii: For each ii in GI : ii.color = RGB(55,0,255): ii.colorfull = RGB(120,220,255) : Next
    Table1.ColorGradeImage = "ColorGradeBlue_on"
    LDome.image = "Dome purple" : RDome.image = "Dome purple"
    LeftFlipperP.BlendDisableLighting = 1 : RightFlipperP.BlendDisableLighting = 1
    LeftFlipperLight.state = 1 : RightFlipperLight.state = 1
    Reflect1.imageA = "CabLmod" : Reflect1.imageB = "CabLmod"
    Reflect2.imageA = "CabRmod" : Reflect2.imageB = "CabRmod"
    For each ii in BWLamps : ii.color = RGB(55,0,255) : Next
    L59.color = RGB(155,0,255):L62.color = RGB(155,0,255):L68.color = RGB(155,0,255):L75.color = RGB(155,0,255)
    F27.color = RGB(155,0,255): F27.colorfull = RGB(120,220,255)
    F29.color = RGB(155,0,255): F29.colorfull = RGB(120,220,255)
    F22b.color = RGB(55,255,55): F22b.colorfull = RGB(180,255,180)
    F22c.color = RGB(55,255,55): F22c.colorfull = RGB(180,255,180)
    F23.color = RGB(255,55,55): F23.colorfull = RGB(255,180,180)
    F32.color = RGB(255,55,55): F32.colorfull = RGB(255,180,180)
  End Select
End Sub

'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                         'volume level; range [0, 1]
NudgeRightSoundLevel = 1                        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                       'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                       'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                     'volume level; range [0, 1]
PlungerPullSoundLevel = 1                       'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                      'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                  'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel               'sound helper; not configurable
SlingshotSoundLevel = 0.95                        'volume level; range [0, 1]
BumperSoundFactor = 4.25                        'volume multiplier; must not be zero
KnockerSoundLevel = 1                           'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                     'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                     'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                    'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                 'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                 'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                     'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                          'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                     'volume multiplier; must not be zero
DTSoundLevel = 0.25                           'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                       'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                           'volume level; range [0, 1]
BallReleaseSoundLevel = 1                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                         'volume multiplier; must not be zero


'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
    PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
    PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound ("Nudge_1"), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
    Case 2 : PlaySound ("Nudge_2"), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
    Case 3 : PlaySound ("Nudge_3"), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
  End Select
End Sub

Sub SoundNudgeRight()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound ("Nudge_1"), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
    Case 2 : PlaySound ("Nudge_2"), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
    Case 3 : PlaySound ("Nudge_3"), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
  End Select
End Sub

Sub SoundNudgeCenter()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound ("Nudge_1"), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
    Case 2 : PlaySound ("Nudge_2"), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
    Case 3 : PlaySound ("Nudge_3"), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
  End Select
End Sub


Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, Plunger
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtLevelStatic ("Drain_1"), DrainSoundLevel, drainswitch
    Case 2 : PlaySoundAtLevelStatic ("Drain_2"), DrainSoundLevel, drainswitch
    Case 3 : PlaySoundAtLevelStatic ("Drain_3"), DrainSoundLevel, drainswitch
    Case 4 : PlaySoundAtLevelStatic ("Drain_4"), DrainSoundLevel, drainswitch
    Case 5 : PlaySoundAtLevelStatic ("Drain_5"), DrainSoundLevel, drainswitch
    Case 6 : PlaySoundAtLevelStatic ("Drain_6"), DrainSoundLevel, drainswitch
    Case 7 : PlaySoundAtLevelStatic ("Drain_7"), DrainSoundLevel, drainswitch
    Case 8 : PlaySoundAtLevelStatic ("Drain_8"), DrainSoundLevel, drainswitch
    Case 9 : PlaySoundAtLevelStatic ("Drain_9"), DrainSoundLevel, drainswitch
    Case 10 : PlaySoundAtLevelStatic ("Drain_10"), DrainSoundLevel, drainswitch
    Case 11 : PlaySoundAtLevelStatic ("Drain_11"), DrainSoundLevel, drainswitch
  End Select
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBallRelease(drainswitch)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("BallRelease1",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 2 : PlaySoundAtLevelStatic SoundFX("BallRelease2",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 3 : PlaySoundAtLevelStatic SoundFX("BallRelease3",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 4 : PlaySoundAtLevelStatic SoundFX("BallRelease4",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 5 : PlaySoundAtLevelStatic SoundFX("BallRelease5",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 6 : PlaySoundAtLevelStatic SoundFX("BallRelease6",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 7 : PlaySoundAtLevelStatic SoundFX("BallRelease7",DOFContactors), BallReleaseSoundLevel, drainswitch
  End Select
End Sub



'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Sling_L1",DOFContactors), SlingshotSoundLevel, Sling
    Case 2 : PlaySoundAtLevelStatic SoundFX("Sling_L2",DOFContactors), SlingshotSoundLevel, Sling
    Case 3 : PlaySoundAtLevelStatic SoundFX("Sling_L3",DOFContactors), SlingshotSoundLevel, Sling
    Case 4 : PlaySoundAtLevelStatic SoundFX("Sling_L4",DOFContactors), SlingshotSoundLevel, Sling
    Case 5 : PlaySoundAtLevelStatic SoundFX("Sling_L5",DOFContactors), SlingshotSoundLevel, Sling
    Case 6 : PlaySoundAtLevelStatic SoundFX("Sling_L6",DOFContactors), SlingshotSoundLevel, Sling
    Case 7 : PlaySoundAtLevelStatic SoundFX("Sling_L7",DOFContactors), SlingshotSoundLevel, Sling
    Case 8 : PlaySoundAtLevelStatic SoundFX("Sling_L8",DOFContactors), SlingshotSoundLevel, Sling
    Case 9 : PlaySoundAtLevelStatic SoundFX("Sling_L9",DOFContactors), SlingshotSoundLevel, Sling
    Case 10 : PlaySoundAtLevelStatic SoundFX("Sling_L10",DOFContactors), SlingshotSoundLevel, Sling
  End Select
End Sub

Sub RandomSoundSlingshotRight(sling)
  Select Case Int(Rnd*8)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Sling_R1",DOFContactors), SlingshotSoundLevel, Sling
    Case 2 : PlaySoundAtLevelStatic SoundFX("Sling_R2",DOFContactors), SlingshotSoundLevel, Sling
    Case 3 : PlaySoundAtLevelStatic SoundFX("Sling_R3",DOFContactors), SlingshotSoundLevel, Sling
    Case 4 : PlaySoundAtLevelStatic SoundFX("Sling_R4",DOFContactors), SlingshotSoundLevel, Sling
    Case 5 : PlaySoundAtLevelStatic SoundFX("Sling_R5",DOFContactors), SlingshotSoundLevel, Sling
    Case 6 : PlaySoundAtLevelStatic SoundFX("Sling_R6",DOFContactors), SlingshotSoundLevel, Sling
    Case 7 : PlaySoundAtLevelStatic SoundFX("Sling_R7",DOFContactors), SlingshotSoundLevel, Sling
    Case 8 : PlaySoundAtLevelStatic SoundFX("Sling_R8",DOFContactors), SlingshotSoundLevel, Sling
  End Select
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
  End Select
End Sub

Sub RandomSoundBumperMiddle(Bump)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
  End Select
End Sub

Sub RandomSoundBumperBottom(Bump)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
  End Select
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
    PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_L01",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_L02",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_L07",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_L08",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_L09",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_L10",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_L12",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 8 : PlaySoundAtLevelStatic SoundFX("Flipper_L14",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 9 : PlaySoundAtLevelStatic SoundFX("Flipper_L18",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 10 : PlaySoundAtLevelStatic SoundFX("Flipper_L20",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 11 : PlaySoundAtLevelStatic SoundFX("Flipper_L26",DOFFlippers), FlipperLeftHitParm, Flipper
  End Select
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_R01",DOFFlippers), FlipperRightHitParm, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_R02",DOFFlippers), FlipperRightHitParm, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_R03",DOFFlippers), FlipperRightHitParm, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_R04",DOFFlippers), FlipperRightHitParm, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_R05",DOFFlippers), FlipperRightHitParm, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_R06",DOFFlippers), FlipperRightHitParm, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_R07",DOFFlippers), FlipperRightHitParm, Flipper
    Case 8 : PlaySoundAtLevelStatic SoundFX("Flipper_R08",DOFFlippers), FlipperRightHitParm, Flipper
    Case 9 : PlaySoundAtLevelStatic SoundFX("Flipper_R09",DOFFlippers), FlipperRightHitParm, Flipper
    Case 10 : PlaySoundAtLevelStatic SoundFX("Flipper_R10",DOFFlippers), FlipperRightHitParm, Flipper
    Case 11 : PlaySoundAtLevelStatic SoundFX("Flipper_R11",DOFFlippers), FlipperRightHitParm, Flipper
  End Select
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L01",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L02",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L03",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
  End Select
End Sub

Sub RandomSoundReflipUpRight(flipper)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R01",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R02",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R03",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
  End Select
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_1",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_2",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_3",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_4",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_5",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_6",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_7",DOFFlippers), FlipperDownSoundLevel, Flipper
  End Select
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  Select Case Int(Rnd*8)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_1",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_2",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_3",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_4",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_5",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_6",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_7",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 8 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_8",DOFFlippers), FlipperDownSoundLevel, Flipper
  End Select
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm/10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm/10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_1"), parm  * RubberFlipperSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_2"), parm  * RubberFlipperSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_3"), parm  * RubberFlipperSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_4"), parm  * RubberFlipperSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_5"), parm  * RubberFlipperSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_6"), parm  * RubberFlipperSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_7"), parm  * RubberFlipperSoundFactor
  End Select
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rollover_1"), RolloverSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("Rollover_2"), RolloverSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("Rollover_3"), RolloverSoundLevel
    Case 4 : PlaySoundAtLevelActiveBall ("Rollover_4"), RolloverSoundLevel
  End Select
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_1"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_2"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_5"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_6"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_7"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_8"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_9"), Vol(ActiveBall) * RubberWeakSoundFactor
  End Select
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

Sub RandomSoundWall()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
  Select Case Int(Rnd*13)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Metal_Touch_1"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Metal_Touch_2"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Metal_Touch_3"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Metal_Touch_4"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall ("Metal_Touch_5"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall ("Metal_Touch_6"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall ("Metal_Touch_7"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall ("Metal_Touch_8"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall ("Metal_Touch_9"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall ("Metal_Touch_10"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 11 : PlaySoundAtLevelActiveBall ("Metal_Touch_11"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 12 : PlaySoundAtLevelActiveBall ("Metal_Touch_12"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 13 : PlaySoundAtLevelActiveBall ("Metal_Touch_13"), Vol(ActiveBall) * MetalImpactSoundFactor
  End Select
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
  RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////
Sub RandomSoundBottomArchBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_2"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_1"), BottomArchBallGuideSoundFactor * 0.25
    Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_2"), BottomArchBallGuideSoundFactor * 0.25
    Case 3 : PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_3"), BottomArchBallGuideSoundFactor * 0.25
  End Select
End Sub


Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Medium_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_2"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*7)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Soft_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Soft_2"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall ("Apron_Soft_3"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 4 : PlaySoundAtLevelActiveBall ("Apron_Soft_4"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 5 : PlaySoundAtLevelActiveBall ("Apron_Soft_5"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 6 : PlaySoundAtLevelActiveBall ("Apron_Soft_6"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 7 : PlaySoundAtLevelActiveBall ("Apron_Soft_7"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_5",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_6",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_7",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_8",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
  End Select
End Sub

Sub RandomSoundTargetHitWeak()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_1",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_2",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_3",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_4",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
End Sub

Sub PlayTargetSound()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft Activeball
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_3"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_4"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_6"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
  End Select
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySoundAtLevelStatic ("Gate_FastTrigger_1"), GateSoundLevel, Activeball
    Case 2 : PlaySoundAtLevelStatic ("Gate_FastTrigger_2"), GateSoundLevel, Activeball
  End Select
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Arch_L1"), Vol(ActiveBall) * ArchSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Arch_L2"), Vol(ActiveBall) * ArchSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Arch_L3"), Vol(ActiveBall) * ArchSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Arch_L4"), Vol(ActiveBall) * ArchSoundFactor
  End Select
End Sub

Sub RandomSoundRightArch()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Arch_R1"), Vol(ActiveBall) * ArchSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Arch_R2"), Vol(ActiveBall) * ArchSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Arch_R3"), Vol(ActiveBall) * ArchSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Arch_R4"), Vol(ActiveBall) * ArchSoundFactor
  End Select
End Sub


Sub Arch1_hit()
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If activeball.velx < -8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If Activeball.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If activeball.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  Select Case Int(Rnd*2)+1
    Case 1: PlaySoundAtLevelStatic ("Saucer_Enter_1"), SaucerLockSoundLevel, Activeball
    Case 2: PlaySoundAtLevelStatic ("Saucer_Enter_2"), SaucerLockSoundLevel, Activeball
  End Select
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, Activeball
    Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, Activeball
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  Dim snd
  Select Case Int(Rnd*7)+1
    Case 1 : snd = "Ball_Collide_1"
    Case 2 : snd = "Ball_Collide_2"
    Case 3 : snd = "Ball_Collide_3"
    Case 4 : snd = "Ball_Collide_4"
    Case 5 : snd = "Ball_Collide_5"
    Case 6 : snd = "Ball_Collide_6"
    Case 7 : snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

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

Option explicit
Randomize

'************************************************************************
'             TABLE OPTIONS
'************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const Ballsize = 50
Const BallMass = 1

Const UseVPMModSol = True
Const UseModulatedMiniDMD = True
Dim UseVPMDMD,CustomDMD,DesktopMode
DesktopMode = Table1.ShowDT : CustomDMD = False
If Right(cGamename,1)="c" Then CustomDMD=True
If CustomDMD OR (NOT DesktopMode AND NOT CustomDMD) Then UseVPMDMD = False    'hides the internal VPMDMD when using the color ROM or when table is in Full Screen and color ROM is not in use
If DesktopMode AND NOT CustomDMD Then UseVPMDMD = True              'shows the internal VPMDMD when in desktop mode and color ROM is not in use
Scoretext.visible = NOT CustomDMD                       'hides the textbox when using the color ROM

LoadVPM "02800000", "Sam.VBS", 3.54


'**********************   General Sound Options   **********************
' VolumeDial:
' VolumeDial is the actual global volume multiplier for the mechanical sounds.
' Values smaller than 1 will decrease mechanical sounds volume.
' Recommended values should be no greater than 1.
Const VolumeDial = 0.8



'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 1

'Standard Sounds
Const SSolenoidOn = "solenoid"
Const SSolenoidOff = ""
Const SCoin = ""

'************************************************************************
'            INIT TABLE
'************************************************************************

Const cGameName = "wof_500"

Dim bsTrough, PlungerIM, DTBank

Sub Table1_Init
  vpmInit Me
    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Wheel Of Fortune (Stern 2007)"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        If NOT CustomDMD Then .Hidden = DesktopMode       'hides the external DMD when in desktop mode and color ROM is not in use
        .HandleMechanics = 1
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

'Nudging
    vpmNudge.TiltSwitch = -7
    vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

'Trough
    Set bsTrough = New cvpmTrough
    With bsTrough
    .Size = 4
    .InitSwitches Array(21, 20, 19, 18)
    .InitExit BallRelease, 70, 15
    '.InitEntrySounds "fx_drain", SoundFX(SSolenoidOn,DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
    '.InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("ballrelease",DOFContactors)
    .Balls = 4
    .CreateEvents "bsTrough", Drain
    End With

'Impulse Plunger
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP sw23, 60, 0.5
        .Switch 23
        .Random 1.5
        '.InitExitSnd SoundFX("AutoPlunger",DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
        .CreateEvents "plungerIM"
    End With

'Targets
  Set DTBank = New cvpmDropTarget
      With DTBank
      .InitDrop Array(sw39, sw40, sw41), Array(39,40,41)
        .Initsnd SoundFX("TotemDropDown", DOFDropTargets), SoundFX("TotemDropUp", DOFDropTargets)
       End With

  BallLock1.IsDropped = 1
  InitWheelLights

End Sub

'************************************************************************
'             KEYS
'************************************************************************

Sub Table1_KeyDown(ByVal Keycode)
  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
    If keycode = LeftTiltKey Then Nudge 90, 3:SoundNudgeLeft()
    If keycode = RightTiltKey Then Nudge 270, 3:SoundNudgeRight()
    If keycode = CenterTiltKey Then Nudge 0, 4:SoundNudgeCenter()
  If KeyCode = PlungerKey Then Plunger.Pullback: SoundPlungerPull()
  If keycode = StartGameKey then soundStartButton()
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
  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
  If KeyCode = PlungerKey Then Plunger.Fire: SoundPlungerReleaseBall()
    If vpmKeyUp(keycode) Then Exit Sub
' if keycode = LeftMagnaSave Then     'debug
'   On Error Resume Next
'   KickerTEST.CreateBall
'   KickerTEST.Kick 0, 45.-3
'   On Error Goto 0
' end if
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_UnPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'************************************************************************
'            SOLENOIDS MAP
'************************************************************************

SolCallBack(1) = "SolTrough"            'Trough-Up Kicker
SolCallBack(2) = "SolAutoPlungerIM"         'AutoLaunch
SolCallback(3) = ""
'4-7                        'Wheel Motor
SolCallback(8) = "DTBank.SolDropUp"         'Drop Target
SolCallback(12) = "solLonnie"           'lonnie
SolCallback(13) = "solMaria"            'Maria
SolCallback(14) = "solKeith"            'Keith
SolCallback(15)= "SolLFlipper"            'Left Flipper
SolCallback(16)= "SolRFlipper"            'Right Flipper
SolModCallback(19) = "SetLampMod 129,"        'Flasher:Red Contestant
SolModCallback(20) = "SetLampMod 130,"        'Flasher:Yellow Contestant
SolModCallback(21) = "SetLampMod 131,"        'Flasher:Blue Contestant
SolCallback(22) = "solBallLock"           'MiniRamp Down Post
SolCallback(23) = "solBallLock1"          'Left Ramp Up Post
SolCallback(24) = "SolKnocker"            'Knocker
SolModCallback(25) = "SetLampMod 135,"        'Flasher:Wheel
SolModCallBack(26) = "SetLampMod 136,"        'Flasher:Backpanel Left
SolModCallback(27) = "SetLampMod 137,"        'Flasher:Left Orbit
SolModCallback(28) = "SetLampMod 138,"        'Flasher:Right Ramp
SolModCallback(29) = "SetLampMod 139,"        'Flasher:Right Orbit
SolModCallback(30) = "SetLampMod 132,"        'Flasher:Pop Bumper
SolModCallback(31) = "SetLampMod 133,"        'Flasher:Mini Ramp
SolModCallBack(32) = "SetLampMod 134,"        'Flasher:Backpanel Right

'************************************************************************
'            SOLENOIDS SUBS
'************************************************************************

' *** BALL TROUGH
Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    If BsTrough.Balls Then vpmTimer.PulseSw 22
  End If
End Sub

' *** AUTOPLUNGER
Sub SolAutoPlungerIM(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
    SoundPlungerReleaseBall()
  End If
End Sub

' *** KNOCKER
Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub

' *** FLIPPERS
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
'        LeftFlipper.RotateToEnd
    lf.fire
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
'        RightFlipper.RotateToEnd
    rf.fire
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

' *** MINIRAMP DOWN POST
Sub SolBallLock(enabled)
  If enabled then
    BallLock.IsDropped = 1
    BallLockP.Z = 0
    PlaySoundAt SoundFX("TotemDropDown", DOFContactors), BallLockP
  else
    BallLock.IsDropped = 0
    BallLockP.Z = 35
    PlaySoundAt SoundFX("TotemDropUp", DOFContactors), BallLockP
  End If
  End Sub

' *** LEFT RAMP UP POST
Sub SolBallLock1(enabled)
  If enabled then
    BallLock1.IsDropped = 0
    BallLock1P.Z = 45
    PlaySoundAt SoundFX("TotemDropUp", DOFContactors), BallLock1P
  else
    BallLock1.IsDropped = 1
    PlaySoundAt SoundFX("TotemDropDown", DOFContactors), BallLock1P
    BallLock1P.Z = 0
  End If
End Sub

' *** LONNIE
Sub solLonnie(enabled)
  If Enabled then
    Lonnie.RotX= 0
    PlaySoundAt SoundFX(SSolenoidOn, DOFContactors), Lonnie
   Else
    Lonnie.RotX= 10
  End If
End Sub

' *** MARIA
  Sub solMaria(enabled)
  If Enabled then
    Maria.RotX= 0
    PlaySoundAt SoundFX(SSolenoidOn, DOFContactors), Maria
   Else
    Maria.RotX= 10
  End If
  End Sub

' *** KEITH
  Sub solKeith(enabled)
    If Enabled then
    Keith.RotX= 0
    PlaySoundAt SoundFX(SSolenoidOn, DOFContactors), Keith
   Else
    Keith.RotX= 10
  End If
  End Sub

'************************************************************************
'             SWITCHES
'************************************************************************

' *** SLINGSHOTS
Dim LStep, RStep

Sub LeftSlingShot_Slingshot()
  RandomSoundSlingshotLeft gi3
    LSling1.Visible = 1
    Lemk.TransZ = -20
    LStep = 0
    vpmTimer.PulseSw 26
    Me.TimerInterval = 20:Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer()
    Select Case LStep
        Case 2:LSLing1.Visible = 0:LSLing2.Visible = 1:Lemk.TransZ = -10
        Case 3:LSLing2.Visible = 0:Lemk.TransZ = 0:Me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot()
  RandomSoundSlingshotRight gi1
    RSling1.Visible = 1
    Remk.TransZ = -20
    RStep = 0
    vpmTimer.PulseSw 27
    Me.TimerInterval = 20:Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer()
    Select Case RStep
        Case 2:RSLing1.Visible = 0:RSLing2.Visible = 1:Remk.TransZ = -10
        Case 3:RSLing2.Visible = 0:Remk.TransZ = 0:Me.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' *** BUMPERS
 Sub Bumper1_Hit():vpmTimer.PulseSw 31:RandomSoundBumperTop Bumper1:End Sub
 Sub Bumper2_Hit():vpmTimer.PulseSw 30:RandomSoundBumperMiddle Bumper2:End Sub
 Sub Bumper3_Hit():vpmTimer.PulseSw 32:RandomSoundBumperBottom Bumper3:End Sub


' *** ROLLOVERS
   Sub sw1_Hit:Controller.Switch(1) = 1:End Sub
   Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub
   Sub sw2_Hit:Controller.Switch(2) = 1:End Sub
   Sub sw2_UnHit:Controller.Switch(2) = 0:End Sub
   Sub sw24_Hit:Controller.Switch(24) = 1:End Sub
   Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
   Sub sw25_Hit:Controller.Switch(25) = 1:End Sub
   Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
   Sub sw28_Hit:Controller.Switch(28) = 1:End Sub
   Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
   Sub sw29_Hit:Controller.Switch(29) = 1:End Sub
   Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
   Sub sw42_Hit:Controller.Switch(42) = 1:End Sub
   Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub
   Sub sw43_Hit:Controller.Switch(43) = 1:End Sub
   Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
   Sub sw45_Hit:Controller.Switch(45) = 1:End Sub
   Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub
   Sub sw57_Hit:Controller.Switch(57) = 1:End Sub
   Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub
   Sub sw58_Hit:Controller.Switch(58) = 1:End Sub
   Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub
   Sub sw59_Hit:Controller.Switch(59) = 1:End Sub
   Sub sw59_UnHit:Controller.Switch(59) = 0:End Sub
   Sub sw60_Hit:Controller.Switch(60) = 1:End Sub
   Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

   Sub sw33_Hit:Controller.Switch(33) = 1:End Sub
   Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

' *** OPTOS
   Sub sw6_Hit:Controller.Switch(6) = 1:End Sub
   Sub sw6_UnHit:Controller.Switch(6) = 0:End Sub
   Sub sw50_Hit:Controller.Switch(50) = 1:End Sub
   Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub
   Sub sw51_Hit:Controller.Switch(51) = 1:PlaySoundAt "fx_ramp_enter", ActiveBall:End Sub
   Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub
   Sub sw52_Hit:Controller.Switch(52) = 1:End Sub
   Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub
   Sub sw53_Hit:Controller.Switch(53) = 1:PlaySoundAt "fx_ramp_metal", ActiveBall:End Sub
   Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub


' *** STANDUP TARGETS
  Sub sw7_Hit:vpmTimer.PulseSw 7:End Sub
  Sub sw8_Hit:vpmTimer.PulseSw 8:End Sub
  Sub sw9_Hit:vpmTimer.PulseSw 9:End Sub
  Sub sw10_Hit:vpmTimer.PulseSw 10:End Sub
  Sub sw11_Hit:vpmTimer.PulseSw 11:End Sub
  Sub sw12_Hit:vpmTimer.PulseSw 12:End Sub
  Sub sw13_Hit:vpmTimer.PulseSw 13:End Sub
  Sub sw14_Hit:vpmTimer.PulseSw 14:End Sub
  Sub sw36_Hit:vpmTimer.PulseSw 36:End Sub
  Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub
  Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub
  Sub sw61_Hit:vpmTimer.PulseSw 61:End Sub

' *** DROP TARGETS
  Sub sw39_dropped:DTBank.Hit 1:End Sub
  Sub sw40_dropped:DTBank.Hit 2:End Sub
  Sub sw41_dropped:DTBank.Hit 3:End Sub

'*****************************************
'           REALTIME UPDATES
'*****************************************

Sub GameTimer_Timer()
  UpdateWheel
  UpdateMiniDMD
  UpdateFlipperShadow
  UpdateBallShadow
  RollingSoundUpdate
  UpdateGates
End Sub

' *** WHEEL
Sub UpdateWheel()
  WheelP.RotY = Controller.GetMech(0)
End Sub

' *** FLIPPER SHADOWS
sub UpdateFlipperShadow()
    batleftshadow.RotZ = LeftFlipper.currentangle
    batrightshadow.RotZ = RightFlipper.currentangle
End Sub
' *** GATES PRIMITIVES
Sub UpdateGates
  Primitive15.RotX = Gate1.currentangle
  Primitive17.RotX = Gate2.currentangle
  Primitive12.RotX = Gate3.currentangle
  Primitive10.RotX = Gate4.currentangle
  Primitive16.RotX = Gate5.currentangle
  Primitive14.RotX = Gate6.currentangle
  Primitive13.RotX = Gate7.currentangle
  Primitive18.RotX = Gate8.currentangle
End Sub

' *** ROLLING SOUND
Const tnob = 4
ReDim rolling(tnob)
InitRolling
Dim isBallOnWireRamp : isBallOnWireRamp = False

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
    StopSound "BallRoll_" & b
        StopSound "fx_plasticrolling" & b
    StopSound "fx_metalrolling" & b
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 Then
      rolling(b) = True
      If isBallOnWireRamp Then
        ' ball on wire ramp
        StopSound "BallRoll_" & b
        StopSound "fx_plasticrolling" & b
        PlaySound "fx_metalrolling" & b, -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
      ElseIf BOT(b).Z > 30 Then
        ' ball on plastic ramp
        StopSound "BallRoll_" & b
        StopSound "fx_metalrolling" & b
        PlaySound "fx_plasticrolling" & b, -1, VolPlayfieldRoll(BOT(b)) * 0.4 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
      Else
        ' ball on playfield
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
        PlaySound "BallRoll_" & b, -1, VolPlayfieldRoll(BOT(b)) * 1.2 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
      End If
    Else
      If rolling(b) = True Then
        StopSound "BallRoll_" & b
        StopSound "fx_plasticrolling" & b
        StopSound "fx_metalrolling" & b
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


' *** BALL SHADOW
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4)

Sub UpdateBallShadow()
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
        If BOT(b).X < table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (table1.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (table1.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 20
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

' *** Mini-DMD

Dim LED(69)

LED(0)=Array(D1,D2,D3,D4,D5)
LED(1)=Array(D6,D7,D8,D9,D10)
LED(2)=Array(D11,D12,D13,D14,D15)
LED(3)=Array(D16,D17,D18,D19,D20)
LED(4)=Array(D21,D22,D23,D24,D25)
LED(5)=Array(D26,D27,D28,D29,D30)
LED(6)=Array(D31,D32,D33,D34,D35)
LED(7)=Array(D36,D37,D38,D39,D40)
LED(8)=Array(D41,D42,D43,D44,D45)
LED(9)=Array(D46,D47,D48,D49,D50)
LED(10)=Array(D51,D52,D53,D54,D55)
LED(11)=Array(D56,D57,D58,D59,D60)
LED(12)=Array(D61,D62,D63,D64,D65)
LED(13)=Array(D66,D67,D68,D69,D70)
LED(14)=Array(D71,D72,D73,D74,D75)
LED(15)=Array(D76,D77,D78,D79,D80)
LED(16)=Array(D81,D82,D83,D84,D85)
LED(17)=Array(D86,D87,D88,D89,D90)
LED(18)=Array(D91,D92,D93,D94,D95)
LED(19)=Array(D96,D97,D98,D99,D100)
LED(20)=Array(D101,D102,D103,D104,D105)
LED(21)=Array(D106,D107,D108,D109,D110)
LED(22)=Array(D111,D112,D113,D114,D115)
LED(23)=Array(D116,D117,D118,D119,D120)
LED(24)=Array(D121,D122,D123,D124,D125)
LED(25)=Array(D126,D127,D128,D129,D130)
LED(26)=Array(D131,D132,D133,D134,D135)
LED(27)=Array(D136,D137,D138,D139,D140)
LED(28)=Array(D141,D142,D143,D144,D145)
LED(29)=Array(D146,D147,D148,D149,D150)
LED(30)=Array(D151,D152,D153,D154,D155)
LED(31)=Array(D156,D157,D158,D159,D160)
LED(32)=Array(D161,D162,D163,D164,D165)
LED(33)=Array(D166,D167,D168,D169,D170)
LED(34)=Array(D171,D172,D173,D174,D175)

Sub UpdateMiniDMD
    dim ii
  IF UseModulatedMiniDMD Then
    dim jj, kk, target, tobj
    for jj = 0 to 4
      for ii = 0 to 6
        for kk = 0 to 4
          target = (kk * 35) + (jj * 7) + ii + 141
          LED(jj * 7 + ii)(4-kk).IntensityScale = (FadingLevel(target)/255) ^ 1.5
          LED(jj * 7 + ii)(4-kk).State = LampState(target)
        Next
      Next
    Next
  Else
    Dim ChgLED,num,chg,stat,obj
    ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
        For Each obj In LED(num)
          If chg And 1 Then
            if stat And 1 Then
              obj.state = 1
            Else
              obj.state = 0
            End If
          End If
          chg=chg\2:stat=stat\2
        Next
      Next
    End If
    End If
End Sub


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
'     With Modulated Flasher Routines (ninuzzu)
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************
Dim LampState(400), FadingLevel(400)
Dim FlashSpeedUp(400), FlashSpeedDown(400), FlashMin(400), FlashMax(400), FlashLevel(400)

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
  UpdateLamps
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
'Inserts
' nFadeL 1, l1    'start
' nFadeL 2, l2    'tournament
  NFadeL 3, l3
  NFadeL 4, l4
  NFadeL 5, l5
  NFadeL 6, l6
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
  Flash 41, F41
  LampMod 41, bulb11
  Flash 42, F42
  LampMod 42, bulb12
  Flash 43, F43
  LampMod 43, bulb13
  Flash 44, F44
  LampMod 44, bulb14
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
  NFadeLm 60, l60
  NFadeLm 60, l60a
  NFadeL 60, l60b
  NFadeLm 61, l61
  NFadeLm 61, l61a
  NFadeL 61, l61b
  NFadeLm 62, l62
  NFadeLm 62, l62a
  NFadeL 62, l62b
  NFadeL 63, l63
  NFadeL 64, l64
  NFadeL 65, l65
  NFadeL 66, l66
  NFadeL 67, l67
  NFadeL 68, l68
  NFadeL 69, l69
  NFadeL 70, l70
  NFadeL 71, l71
  NFadeL 72, l72
  NFadeL 73, l73
  NFadeL 74, l74
  NFadeL 75, l75
  NFadeL 76, l76
  NFadeL 77, l77
  NFadeL 78, l78
  NFadeL 79, l79
  NFadeL 80, l80

'Wheel LEDs
' Flash 81,   '?????????
  Flash 82, WLedRedWedge '#1
' Flash 83, WLedRedWedge '#2
' Flash 84, WLedRedWedge '#3
  Flash 85, WLedGIL   'Left GI
  Flash 86, WLedGIR   'Right GI
' Flash 87,   '?????????
' Flash 89, WLedYellowWedge '#1
  Flash 90, WLedBlueWedge '#1
' Flash 91, WLedBlueWedge '#2
' Flash 92, WLedBlueWedge '#3

' Flash 93, WLedYellowWedge '#2
  Flash 94, WLedYellowWedge '#3
' Flash 95,   '?????????
  Flash 105, WLedRedContestant
  Flash 109, WLedYellowContestant
  Flash 110, WLedBlueContestant
' Flash 111,    '?????????

'Wheel Border LEDs
  Flashm 106,WLed1:Flash 106,WBorder1
  Flashm 100,WLed2:Flash 100,WBorder2
  Flashm 99,WLed3:Flash 99,WBorder3
  Flashm 98,WLed4:Flash 98,WBorder4
  Flashm 97,WLed5:Flash 97,WBorder5
  Flashm 101,WLed6:Flash 101,WBorder6
  Flashm 102,WLed7:Flash 102,WBorder7
  Flashm 103,WLed8:Flash 103,WBorder8
  Flashm 124,WLed9:Flash 124,WBorder9
  Flashm 123,WLed10:Flash 123,WBorder10
  Flashm 122,WLed11:Flash 122,WBorder11
  Flashm 121,WLed12:Flash 121,WBorder12
  Flashm 125,WLed13:Flash 125,WBorder13
  Flashm 126,WLed14:Flash 126,WBorder14
  Flashm 127,WLed15:Flash 127,WBorder15
  Flashm 116,WLed16:Flash 116,WBorder16
  Flashm 115,WLed17:Flash 115,WBorder17
  Flashm 114,WLed18:Flash 114,WBorder18
  Flashm 113,WLed19:Flash 113,WBorder19
  Flashm 117,WLed20:Flash 117,WBorder20
  Flashm 118,WLed21:Flash 118,WBorder21
  Flashm 119,WLed22:Flash 119,WBorder22
  Flashm 108,WLed23:Flash 108,WBorder23
  Flashm 107,WLed24:Flash 107,WBorder24

'Flashers
  LampMod 129, FL19     'Flasher:Red Contestant
  LampMod 130, FL20     'Flasher:Yellow Contestant
  LampMod 131, FL21     'Flasher:Blue Contestant
  LampMod 135, FL25a      'Flasher:Wheel #1
  LampMod 135, FL25b      'Flasher:Wheel #2
  LampMod 135, FL25c      'Flasher:Wheel #3
  LampMod 135, FL25d      'Flasher:Wheel #4
  LampMod 136, FL26     'Flasher:Backpanel Left #1
  LampMod 136, FL26a      'Flasher:Backpanel Left #2
  LampMod 136, FL26b      'Flasher:Backpanel Left #3
  LampMod 137, FL27     'Flasher:Left Orbit
  LampMod 138, FL28     'Flasher:Right Ramp
  LampMod 139, FL29     'Flasher:Right Orbit
  LampMod 132, FL30     'Flasher:Pop Bumper #1
  LampMod 132, FL30a      'Flasher:Pop Bumper #2
  LampMod 133, FL31     'Flasher:Mini Ramp
  LampMod 134, FL32     'Flasher:Backpanel Right #1
  LampMod 134, FL32a      'Flasher:Backpanel Right #2
  LampMod 134, FL32b      'Flasher:Backpanel Right #3

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
        Case 5:object.image = a:FadingLevel(nr) = 1   'on
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

' Modulated Flasher and Lights objects

Sub SetLampMod(nr, value)
    If value > 0 Then
    LampState(nr) = 1
  Else
    LampState(nr) = 0
  End If
  FadingLevel(nr) = value
End Sub

Sub LampMod(nr, object)
  If TypeName(object) = "Light" Then
    Object.IntensityScale = FadingLevel(nr)/128
    Object.State = LampState(nr)
  End If
  If TypeName(object) = "Flasher" Then
    Object.IntensityScale = FadingLevel(nr)/128
    Object.visible = LampState(nr)
  End If
  If TypeName(object) = "Primitive" Then
    Object.DisableLighting = LampState(nr)
  End If
End Sub

' *********************************************************************
'                 General Illumination
' *********************************************************************
set GICallback = GetRef("UpdateGI")

Sub UpdateGI(nr,enabled)
  Dim xx
  Select Case nr
    Case 0
    If Enabled Then
      DOF 101, DOFOn
      Table1.ColorGradeImage= "ColorGrade_8"
      For Each xx in GI:xx.State=1:Next
      For Each xx in BWGI:xx.visible=1:Next
      Playfield_GI.visible=1
      metals2.image="metals_2"
      metals1.image="metals_1"
      primitive8.image="metals_3"
      primitive3.image="wood"
      For xx = 1 to 10
        ExecuteGlobal "bulb" & xx & ".image = ""bulb_orange"":bulb" & xx & ".DisableLighting = True"
      Next
      Spot1.visible=1:Spot2.visible=1:Spot3.visible=1:Spot4.visible=1
      Playsound "fx_relay_on"
    Else
      DOF 101, DOFOff
      Table1.ColorGradeImage= "ColorGrade_1"
      For Each xx in GI:xx.State=0:Next
      For Each xx in BWGI:xx.visible=0:Next
      Playfield_GI.visible=0
      metals2.image="metals_2off"
      metals1.image="metals_1off"
      primitive8.image="metals_3off"
      primitive3.image="woodoff"
      For xx = 1 to 10
        ExecuteGlobal "bulb" & xx & ".image = ""bulb"":bulb" & xx & ".DisableLighting = False"
      Next
      Spot1.visible=0:Spot2.visible=0:Spot3.visible=0:Spot4.visible=0
      Playsound "fx_relay_off"
    End If
  End Select
End Sub



'*****************************************
'         COLLISON SOUND EFFECTS
'*****************************************

' left ramp sounds
Sub LRHit0_Hit:PlaySoundAt "fx_lr1", ActiveBall:End Sub
Sub LRHit1_Hit:PlaySoundAt "fx_lr2", ActiveBall:End Sub
Sub LRHit2_Hit:PlaySoundAt "fx_lr3", ActiveBall:End Sub
Sub LRHit3_Hit:PlaySoundAt "fx_lr4", ActiveBall:End Sub
Sub LRHit4_Hit:PlaySoundAt "fx_lr5", ActiveBall:End Sub
Sub LRHit5_Hit:PlaySoundAt "fx_lr6", ActiveBall:End Sub
Sub LRHit6_Hit:PlaySoundAt "fx_lr7", ActiveBall:End Sub

'right ramp sounds
Sub RRHit0_Hit:PlaySoundAt "fx_rr1", ActiveBall:End Sub
Sub RRHit1_Hit:PlaySoundAt "fx_rr2", ActiveBall:End Sub
Sub RRHit2_Hit:PlaySoundAt "fx_rr3", ActiveBall:End Sub
Sub RRHit3_Hit:PlaySoundAt "fx_rr4", ActiveBall:End Sub
Sub RRHit4_Hit:PlaySoundAt "fx_rr5", ActiveBall:End Sub
Sub RRHit5_Hit:PlaySoundAt "fx_rr6", ActiveBall:End Sub
Sub RRHit6_Hit:PlaySoundAt "fx_lr1", ActiveBall:End Sub
Sub RRHit7_Hit:PlaySoundAt "fx_lr1", ActiveBall:End Sub
Sub RRHit8_Hit:PlaySoundAt "fx_rr7", ActiveBall:End Sub

Sub Drain_Hit()
  RandomSoundDrain Drain
End Sub

Sub BallRelease_UnHit()
  RandomSoundBallRelease BallRelease
End Sub

'*****************************************
'       INIT WHEEL LIGHTS POSITION
'*****************************************

Const Pi=3.1415926535
Const LedRadius = 141.75

Sub InitWheelLights
dim xx
'Wheel Border LEDs #1
For xx = 1 to 24
WheelLEDs(xx-1).X = WheelP.X + LedRadius * Sin((xx-1) * 15 * Pi/180)
WheelLEDs(xx-1).Y = WheelP.Y - LedRadius * Cos((xx-1) * 15 * Pi/180) * sin (WheelP.objrotx * Pi/180) -5
WheelLEDs(xx-1).Height = WheelP.Z + LedRadius * Cos((xx-1) * 15 * Pi/180) * cos (WheelP.objrotx * Pi/180)
WheelLEDs(xx-1).RotX = WheelP.objrotx -90
Next
'Wheel Border LEDs #2
For xx = 1 to 24
WheelBorder(xx-1).X = WheelP.X + LedRadius * Sin((xx-1) * 15 * Pi/180)
WheelBorder(xx-1).Y = WheelP.Y - LedRadius * Cos((xx-1) * 15 * Pi/180) * sin (WheelP.objrotx * Pi/180)-5
WheelBorder(xx-1).Height = WheelP.Z + LedRadius * Cos((xx-1) * 15 * Pi/180) * cos (WheelP.objrotx * Pi/180)
WheelBorder(xx-1).RotX = WheelP.objrotx -90
Next
'Wheel Wedge R-Y-B LEDs
WLedRedWedge.X = WheelP.X+10:WLedRedWedge.Y=WheelP.Y +1:WLedRedWedge.height = WheelP.Z:WLedRedWedge.RotX = WheelP.objrotx -90
WLedYellowWedge.X = WheelP.X:WLedYellowWedge.Y=WheelP.Y +1:WLedYellowWedge.height = WheelP.Z:WLedYellowWedge.RotX = WheelP.objrotx -90
WLedBLueWedge.X = WheelP.X-10:WLedBLueWedge.Y=WheelP.Y +1:WLedBLueWedge.height = WheelP.Z:WLedBLueWedge.RotX = WheelP.objrotx -90
'Wheel Contenstant R-Y-B LEDs
WLedRedContestant.X = WheelP.X+10:WLedRedContestant.Y=WheelP.Y +15*cos(WheelP.objrotx*Pi/180):WLedRedContestant.height = WheelP.Z+15*(1+sin(WheelP.objrotx*Pi/180)):WLedRedContestant.RotX = WheelP.objrotx -90
WLedYellowContestant.X = WheelP.X:WLedYellowContestant.Y=WheelP.Y +15*cos(WheelP.objrotx*Pi/180):WLedYellowContestant.height = WheelP.Z+15*(1+sin(WheelP.objrotx*Pi/180)):WLedYellowContestant.RotX = WheelP.objrotx -90
WLedBlueContestant.X = WheelP.X-10:WLedBlueContestant.Y=WheelP.Y +15*cos(WheelP.objrotx*Pi/180):WLedBlueContestant.height = WheelP.Z+15*(1+sin(WheelP.objrotx*Pi/180)):WLedBlueContestant.RotX = WheelP.objrotx -90
'GI:Wheel Left
WLedGIR.X = WheelP.X + 80 * Sin(150 * Pi/180)
WLedGIR.Y = WheelP.Y - 80 * Cos(150 * Pi/180) * sin (WheelP.objrotx * Pi/180) + 5
WLedGIR.Height = WheelP.Z + 90 * Cos(150 * Pi/180) * cos (WheelP.objrotx * Pi/180)
WLedGIR.RotX = WheelP.objrotx -90
'GI:Wheel Right
WLedGIL.X = WheelP.X + 80 * Sin(210 * Pi/180)
WLedGIL.Y = WheelP.Y - 80 * Cos(210 * Pi/180) * sin (WheelP.objrotx * Pi/180) + 5
WLedGIL.Height = WheelP.Z + 90 * Cos(210 * Pi/180) * cos (WheelP.objrotx * Pi/180)
WLedGIL.RotX = WheelP.objrotx -90
'Flasher:Wheel #1
FL25a.X = WheelP.X + 90 * Sin(-90 * Pi/180)
FL25a.Y = WheelP.Y - 90 * Cos(-90 * Pi/180) * sin (WheelP.objrotx * Pi/180) + 5
FL25a.Height = WheelP.Z + 90 * Cos(-90 * Pi/180) * cos (WheelP.objrotx * Pi/180)
FL25a.RotX = WheelP.objrotx -90
'Flasher:Wheel #2
FL25b.X = WheelP.X + 90 * Sin(90 * Pi/180)
FL25b.Y = WheelP.Y - 90 * Cos(90 * Pi/180) * sin (WheelP.objrotx * Pi/180) + 5
FL25b.Height = WheelP.Z + 90 * Cos(90 * Pi/180) * cos (WheelP.objrotx * Pi/180)
FL25b.RotX = WheelP.objrotx -90
'Flasher:Wheel #3
FL25c.X = WheelP.X + 90 * Sin(0 * Pi/180)
FL25c.Y = WheelP.Y - 90 * Cos(0 * Pi/180) * sin (WheelP.objrotx * Pi/180) + 5
FL25c.Height = WheelP.Z + 90 * Cos(0 * Pi/180) * cos (WheelP.objrotx * Pi/180)
FL25c.RotX = WheelP.objrotx -90
'Flasher:Wheel #4
FL25d.X = WheelP.X + 90 * Sin(180 * Pi/180)
FL25d.Y = WheelP.Y - 90 * Cos(180 * Pi/180) * sin (WheelP.objrotx * Pi/180) + 5
FL25d.Height = WheelP.Z + 90 * Cos(180 * Pi/180) * cos (WheelP.objrotx * Pi/180)
FL25d.RotX = WheelP.objrotx -90
End Sub


dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
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

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

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
'                        FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
        dim a : a = Array(LF, RF)
        dim x : for each x in a
                x.addpoint aStr, idx, aX, aY
        Next
End Sub

Class FlipperPolarity
        Public DebugOn, Enabled
        Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
        Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
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

        Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
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
        Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

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
                                        if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
                                end if
                        Next

                        If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
                                BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
                                if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
                        End If

                        'Velocity correction
                        if not IsEmpty(VelocityIn(0) ) then
                                Dim VelCoef
         :                         VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

                                if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

                                if Enabled then aBall.Velx = aBall.Velx*VelCoef
                                if Enabled then aBall.Vely = aBall.Vely*VelCoef
                        End If

                        'Polarity Correction (optional now)
                        if not IsEmpty(PolarityIn(0) ) then
                                If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
                                dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

                                if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
                        End If
                End If
                RemoveBall aBall
        End Sub
End Class

'******************************************************
'                FLIPPER POLARITY AND RUBBER DAMPENER
'                        SUPPORTING FUNCTIONS
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

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
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
        dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
                if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
        Next
        if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
        Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

        if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
        if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

        LinearEnvelope = Y
End Function

' Used for drop targets and flipper tricks
Function Distance(ax,ay,bx,by)
        Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function


'******************************************************
'                        FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
        FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
        FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
        FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
        FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
        Dim BOT, b

        If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
                EOSNudge1 = 1
                'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
                If Flipper2.currentangle = EndAngle2 Then
                        BOT = GetBalls
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
                                        'Debug.Print "ball in flip1. exit"
                                         exit Sub
                                end If
                        Next
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
                                        BOT(b).velx = BOT(b).velx / 1.7
                                        BOT(b).vely = BOT(b).vely - 1
                                end If
                        Next
                End If
        Else
                If Flipper1.currentangle <> EndAngle1 then
                        EOSNudge1 = 0
                end if
        End If
End Sub

'*****************
' Maths
'*****************Const PI = 3.1415927

Function dSin(degrees)
        dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
        dcos = cos(degrees * Pi/180)
End Function

'*************************************************
' Check ball distance from Flipper for Rem
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

Function AnglePP(ax,ay,bx,by)
        AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
        DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
        Dim DiffAngle
        DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
        If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

        If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
                FlipperTrigger = True
        Else
                FlipperTrigger = False
        End If
End Function

' Used for drop targets and stand up targets
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

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
        FlipperPress = 1
        Flipper.Elasticity = FElasticity

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
        Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

        If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
                If FState <> 1 Then
                        Flipper.rampup = SOSRampup
                        Flipper.endangle = FEndAngle - 3*Dir
                        Flipper.Elasticity = FElasticity * SOSEM
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
        Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
        Dim CatchTime : CatchTime = GameTime - FCount

        if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
                if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
                        LiveCatchBounce = 0
                else
                        LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
                end If

                If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
                ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
                ball.angmomx= 0
                ball.angmomy= 0
                ball.angmomz= 0
        End If
End Sub


'**************************************************
'        Flipper Collision Subs
'NOTE: COpy and overwrite collision sound from original collision subs over
'RandomSoundFlipper()' below
'**************************************************'

Sub LeftFlipper_Collide(parm)
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
    LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
    RightFlipperCollide parm
End Sub


'****************************************************************************
'PHYSICS DAMPENERS
'****************************************************************************

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
        RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
        SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
        Public Print, debugOn 'tbpOut.text
        public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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

' Thalamus - patched : ' Thalamus - patched :                 aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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


        Public Sub Report()         'debug, reports all coords in tbPL.text
                if not debugOn then exit sub
                dim a1, a2 : a1 = ModIn : a2 = ModOut
                dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
                TBPout.text = str
        End Sub


End Class


'******************************************************
'                TRACK ALL BALL VELOCITIES
'                 FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
        public ballvel, ballvelx, ballvely

        Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

        Public Sub Update()        'tracks in-ball-velocity
                dim str, b, AllBalls, highestID : allBalls = getballs

                for each b in allballs
                        if b.id >= HighestID then highestID = b.id
                Next

                if uBound(ballvel) < highestID then redim ballvel(highestID)        'set bounds
                if uBound(ballvelx) < highestID then redim ballvelx(highestID)        'set bounds
                if uBound(ballvely) < highestID then redim ballvely(highestID)        'set bounds

                for each b in allballs
                        ballvel(b.id) = BallSpeed(b)
                        ballvelx(b.id) = b.velx
                        ballvely(b.id) = b.vely
                Next
        End Sub
End Class

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

Sub RDampen_Timer()
Cor.Update
End Sub





'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                                                                                                                'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                                                                                                        'volume level; range [0, 1]
NudgeRightSoundLevel = 1                                                                                                'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                                                                                                'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                                                                                                'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                                                                                        'volume level; range [0, 1]
PlungerPullSoundLevel = 1                                                                                                'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                                                           'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
SlingshotSoundLevel = 0.95                                                                                                'volume level; range [0, 1]
BumperSoundFactor = 4.25                                                                                                'volume multiplier; must not be zero
KnockerSoundLevel = 1                                                                                                         'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                                                                                        'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                                                                                'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                                                                        'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                                                                        'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                                                                                        'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                                                                                                        'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                                                      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                                                                                                                'volume level; range [0, 1]
BallReleaseSoundLevel = 1                                                                                                'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                                                                        'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                                                                                'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                                                                                                        'volume multiplier; must not be zero


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
'                     Fleep  Supporting Ball & Sound Functions
' *********************************************************************

Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

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
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
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
        PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
        PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
        PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
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
        PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
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
        PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
        PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
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
        PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
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
        PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
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
                PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
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
        PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
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
                PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
         End If
        If finalspeed < 6 Then
                PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
        End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
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
        PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
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
        PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
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
        PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
        PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
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
        PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
        Select Case scenario
                Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
                Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
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


'/////////////////////////////////////////////////////////////////
'                                        End Mechanical Sounds
'/////////////////////////////////////////////////////////////////



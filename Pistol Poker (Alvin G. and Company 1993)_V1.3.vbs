'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************
' _______   __              __                __        _______            __
'/       \ /  |            /  |              /  |      /       \          /  |
'$$$$$$$  |$$/   _______  _$$ |_     ______  $$ |      $$$$$$$  | ______  $$ |   __   ______    ______
'$$ |__$$ |/  | /       |/ $$   |   /      \ $$ |      $$ |__$$ |/      \ $$ |  /  | /      \  /      \
'$$    $$/ $$ |/$$$$$$$/ $$$$$$/   /$$$$$$  |$$ |      $$    $$//$$$$$$  |$$ |_/$$/ /$$$$$$  |/$$$$$$  |
'$$$$$$$/  $$ |$$      \   $$ | __ $$ |  $$ |$$ |      $$$$$$$/ $$ |  $$ |$$   $$<  $$    $$ |$$ |  $$/
'$$ |      $$ | $$$$$$  |  $$ |/  |$$ \__$$ |$$ |      $$ |     $$ \__$$ |$$$$$$  \ $$$$$$$$/ $$ |
'$$ |      $$ |/     $$/   $$  $$/ $$    $$/ $$ |      $$ |     $$    $$/ $$ | $$  |$$       |$$ |
'$$/       $$/ $$$$$$$/     $$$$/   $$$$$$/  $$/       $$/       $$$$$$/  $$/   $$/  $$$$$$$/ $$/

'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************

Option Explicit
Randomize


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Const BallSize = 50
Const BallMass = 1
Const cGameName="pstlpkr"
Const UseSolenoids=2,UseLamps=0,UseGI=0,SCoin="Coin3"

LoadVPM "01300000","alvinG.VBS",3.10

'Dont use Table1.width or Table1.height in script as it can effect performance
dim tablewidth: tablewidth = Table1.width
dim tableheight: tableheight = Table1.height
dim rsl,dwall,FLVol,Ramp1,Ramp2,mode
'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************

'    _____       _      _          ___          _    _
'   |_   _|__ _ | |__  | |  ___   / _ \  _ __  | |_ (_)  ___   _ __   ___
'     | | / _` || '_ \ | | / _ \ | | | || '_ \ | __|| | / _ \ | '_ \ / __|
'     | || (_| || |_) || ||  __/ | |_| || |_) || |_ | || (_) || | | |\__ \
'     |_| \__,_||_.__/ |_| \___|  \___/ | .__/  \__||_| \___/ |_| |_||___/
'                                       |_|

'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.3


'Rubbers on side lanes (real table doesnt have these, but these makes the table a bit less brutal)
'on =1
'off = 0
rsl = 1

'Debug, Stop ball going in centre kicker....for testing only.
'0=off, 1=on
dwall=0

'Flasher switch volume
FLVol=0.001



'**********************************************************************************************************
'Solenoid Call backs
'**********************************************************************************************************

'SolCallback(1)="vpmSolSound ""Jet1"","
'SolCallback(2)="vpmSolSound ""Jet2"","
'SolCallback(3)="vpmSolSound ""Sling"","
'SolCallback(4)="vpmSolSound ""Sling"","
'SolCallback(5)="vpmSolSound ""Sling"","
SolCallback(6)="bsLowerleft" 'Joker hole
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(13)=  "bsTrough.SolIn"                  'outhole
SolCallback(14)=  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'SolCallback(15)="Flashers_Relay"
SolCallback(16)="PlayGI"

'Flashers
SolCallback(17)="sol17subroutine" 'left plastics
SolCallback(18)="sol18subroutine" 'right plastics
SolCallback(19)="sol19subroutine" 'under top playfield
SolCallback(20)="sol20subroutine" 'spades
SolCallback(21)="sol21subroutine" 'top left under plastic
SolCallback(22)="sol22subroutine" 'top right under plastic
SolCallback(23)="sol23subroutine" 'under ramp entry
SolCallback(24)="sol24subroutine" 'Gun Fire

SolCallback(25)="VukTopPop" 'Gun Kick
'26 not used
SolCallback(27)="vpmNudge.SolGameOn"
'28 not used
SolCallback(29)="sol29subroutine" 'Backglass GI
SolCallback(30)=  "bsTrough.SolOut"


'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

Dim bsTrough


Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Anti"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 0
        .Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=10
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(BumperL,BumperR,LeftSlingshot,RightSlingshot)

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 17,18,0,0,0,0,0,0
  bsTrough.InitKick BallRelease,130,3
  bsTrough.Balls=3

  Dim DesktopMode: DesktopMode = Table1.ShowDT

  If DesktopMode = True Then 'Show Desktop components
    Dim BGLight, DNuts
    For each BGLight in Backglass:BGLight.Visible = 1: Next
    For each DNuts in DomeNuts:DNuts.BlendDisableLighting = 0.04: Next
    BigRamp002.image = "RampDT_off"
    ramp1="rampDT"
    ramp2="rampDT_off"
    BigRamp002.Material = "MaterialDT"
    BigRamp002.image = ramp1
    mode = 1
  Else
    For each BGLight in Backglass:BGLight.Visible = 0: Next
    For each DNuts in DomeNuts:DNuts.BlendDisableLighting = 0: Next
    ramp1="rampFS"
    ramp2="rampFS_off"
    BigRamp002.Material = "MaterialFS"
    BigRamp002.image = ramp1
    mode = 2
  End if

  'Drop targets
  If dwall = 1 Then
    THEWALL.IsDropped = 0
  Else
    THEWALL.IsDropped = 1
  End If
    Wall024.IsDropped = 0
    Shutter.IsDropped = 0

  'Load LUT
    LoadLUT

  'Side Rubbers
  If rsl = 1 Then

    Pin005.Visible = 1
    Pin006.Visible = 1
    Primitive082.collidable=1
    Primitive120.collidable=1
    Primitive208.collidable=0
    Primitive209.collidable=0

  Else

    Pin005.Visible = 0
    Pin006.Visible = 0
    Primitive082.collidable=0
    Primitive120.collidable=0
    Primitive208.collidable=1
    Primitive209.collidable=1

  End If



End Sub

'******************************************************
'         SOL SUBS
'******************************************************

sub sol17subroutine(Enabled)
    If enabled then
    SetLamp 117,1
    playsound "fx_relay", 0,FLVol,0,0.02
    else
       SetLamp 117,0
    end if
End Sub

sub sol18subroutine(Enabled)
    If enabled then
    SetLamp 118,1
    playsound "fx_relay", 0,FLVol,0,0.02
    else
       SetLamp 118,0
    end if
End Sub

sub sol19subroutine(Enabled)
    If enabled then
    SetLamp 119,1
    playsound "fx_relay", 0,FLVol,0,0.02
    else
       SetLamp 119,0
    end if
End Sub

sub sol20subroutine(Enabled)
    If enabled then
    SetLamp 120,1
    playsound "fx_relay", 0,FLVol,0,0.02
    else
       SetLamp 120,0
    end if
End Sub

sub sol21subroutine(Enabled)
    If enabled then
    SetLamp 121,1
    playsound "fx_relay", 0,FLVol,0,0.02
    else
       SetLamp 121,0
    end if
End Sub

sub sol22subroutine(Enabled)
    If enabled then
    SetLamp 122,1
    playsound "fx_relay", 0,FLVol,0,0.02
    else
       SetLamp 122,0
    end if
End Sub

sub sol23subroutine(Enabled)
    If enabled then
    SetLamp 123,1
    playsound "fx_relay", 0,FLVol,0,0.02
    else
       SetLamp 123,0
    end if
End Sub

sub sol24subroutine(Enabled)
    If enabled then
    SetLamp 124,1
    playsound "fx_relay", 0,0.3,0,0.02
    else
       SetLamp 124,0
    end if
End Sub

sub sol29subroutine(Enabled)
    If enabled then
    playsound "fx_relay", 0,0.2,0,0.02
    end if
End Sub

'******************************************************
'         FLIPPERS
'******************************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend

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
    RF.Fire 'rightflipper.rotatetoend
    Flipper1.RotateToEnd
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    Flipper1.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


Sub LeftFlipper_Collide(parm)
        LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
        RightFlipperCollide parm
End Sub

'******************************************************
'         Drain hole
'******************************************************

Sub Drain_Hit()
  RandomSoundDrain Drain
  bsTrough.addball Me
End Sub

Sub BallRelease_unHit()
  RandomSoundBallRelease BallRelease
End Sub

'******************************************************
'         TRIGGERS
'******************************************************

dim BIPL

Sub sw21_Hit:Controller.Switch(21)=1 : RandomSoundRollover:BIPL = 1: End Sub
Sub sw21_unHit:Controller.Switch(21)=0 : RandomSoundRollover:BIPL = 0:End Sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' This one has sound event - keeping it
Sub sw47_Hit:Controller.Switch(47)=1 : RandomSoundRollover: End Sub
Sub sw47_unHit:Controller.Switch(47)=0:End Sub


Sub sw62_Hit:Controller.Switch(62)=1 : RandomSoundRollover: End Sub
Sub sw62_unHit:Controller.Switch(62)=0:End Sub

Sub sw63_Hit:Controller.Switch(63)=1 : RandomSoundRollover: End Sub
Sub sw63_unHit:Controller.Switch(63)=0:End Sub

Sub sw78_Hit:Controller.Switch(78)=1 : RandomSoundRollover : End Sub
Sub sw78_unHit:Controller.Switch(78)=0:End Sub

Sub sw79_Hit:Controller.Switch(79)=1 : RandomSoundRollover : End Sub
Sub sw79_unHit:Controller.Switch(79)=0:End Sub

Sub sw23_Hit:Controller.Switch(23)=1: SetLamp 123,1 : RandomSoundRollover : End Sub
Sub sw23_unHit:Controller.Switch(23)=0: SetLamp 123,0:End Sub

Sub sw24_Hit:Controller.Switch(24)=1: SetLamp 121,1: SetLamp 122,1 : RandomSoundRollover : End Sub
Sub sw24_unHit:Controller.Switch(24)=0: SetLamp 121,0: SetLamp 122,0:End Sub

Sub sw25_Hit:Controller.Switch(25)=1 : RandomSoundRollover : End Sub
Sub sw25_unHit:Controller.Switch(25)=0:End Sub

Sub sw26_Hit:Controller.Switch(26)=1 : RandomSoundRollover : End Sub
Sub sw26_unHit:Controller.Switch(26)=0:End Sub

Sub sw27_Hit:Controller.Switch(27)=1 : RandomSoundRollover : End Sub
Sub sw27_unHit:Controller.Switch(27)=0:End Sub

Sub sw28_Hit:Controller.Switch(28)=1 : RandomSoundRollover : End Sub
Sub sw28_unHit:Controller.Switch(28)=0:End Sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub sw70_Hit:Controller.Switch(70)=1 : RandomSoundRollover : End Sub
' Sub sw70_unHit:Controller.Switch(70)=0:End Sub

Sub sw70_Hit:Controller.Switch(70)=1 : RandomSoundRollover : End Sub
Sub sw70_unHit:Controller.Switch(70)=0:End Sub

Sub sw71_Hit:Controller.Switch(71)=1 : RandomSoundRollover : End Sub
Sub sw71_unHit:Controller.Switch(71)=0:End Sub

Sub sw72_Hit:Controller.Switch(72)=1 : RandomSoundRollover : End Sub
Sub sw72_unHit:Controller.Switch(72)=0:End Sub

Sub sw56_Hit:Controller.Switch(56)=1 : RandomSoundRollover : End Sub
Sub sw56_unHit:Controller.Switch(56)=0:End Sub

Sub sw80_Hit:Controller.Switch(80)=1 : RandomSoundRollover : End Sub
Sub sw80_unHit:Controller.Switch(80)=0:End Sub

Sub sw52_Hit:Controller.Switch(52)=1 : RandomSoundRollover :Wall024.IsDropped=0: End Sub
Sub sw52_unHit:Controller.Switch(52)=0:End Sub

Sub sw53_Hit:Controller.Switch(53)=1 : RandomSoundRollover : End Sub
Sub sw53_unHit:Controller.Switch(53)=0:End Sub

Sub sw46_Hit:vpmTimer.PulseSw 46:End Sub
'Sub sw47_Hit:vpmTimer.PulseSw 47:End Sub

'Targets
Sub sw73_Hit:vpmTimer.PulseSw 73:End Sub
Sub sw74_Hit:vpmTimer.PulseSw 74:End Sub
Sub sw75_Hit:vpmTimer.PulseSw 75:End Sub
Sub sw76_Hit:vpmTimer.PulseSw 76:End Sub
Sub sw77_Hit:vpmTimer.PulseSw 77:End Sub

Sub sw45_Hit:vpmTimer.PulseSw 45:End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:End Sub

Sub sw57_Hit:vpmTimer.PulseSw 57:End Sub
Sub sw58_Hit:vpmTimer.PulseSw 58:End Sub
Sub sw59_Hit:vpmTimer.PulseSw 59:End Sub
Sub sw60_Hit:vpmTimer.PulseSw 60:End Sub
Sub sw61_Hit:vpmTimer.PulseSw 61:End Sub

Sub sw65_Hit:vpmTimer.PulseSw 65:End Sub
Sub sw66_Hit:vpmTimer.PulseSw 66:End Sub
Sub sw68_Hit:vpmTimer.PulseSw 68:End Sub
Sub sw69_Hit:vpmTimer.PulseSw 69:End Sub

Sub BumperL_Hit
  vpmTimer.PulseSw 30
  RandomSoundBumperTop BumperL
End Sub
Sub BumperR_Hit
  vpmTimer.PulseSw 31
  RandomSoundBumperMiddle BumperR
End Sub


Sub sw29_Hit:vpmTimer.PulseSw 29 : RandomSoundRollover :End Sub

'******************************************************
'         Ramp & Loop Accelerators
'******************************************************

Sub swSC01_Hit
  'debug.print "--------"
  'debug.print activeball.velY
  'debug.print activeball.velX
  'debug.print "--------"
  if activeball.velY < -10 and activeball.velX > 0 then
    'debug.print "--------"
    'debug.print "T1"
    activeball.velY=activeball.velY-7
    activeball.velX=activeball.velX+9
  end if
End Sub

Sub swSC04_Hit
  'debug.print activeball.velY
  if activeball.velY < -8 then
    'debug.print "T2"
    'debug.print "--------"
    activeball.velY=activeball.velY-4
    activeball.velX=activeball.velX+4
  end if
End Sub

Sub swSC02_Hit
  'debug.print activeball.velY
  if activeball.velY < -5 then
    activeball.velY=activeball.velY-16
  end if
End Sub

Sub swSC03_Hit
  'debug.print activeball.velY
  if activeball.velY < -5 then
    activeball.velY=activeball.velY-20
  end if
  if activeball.velY > 10 then
    activeball.velY=activeball.velY-9
  end if
End Sub

'******************************************************
'       Top playfield shutter
'******************************************************

Sub swShutter_Hit
  PlaySoundAt"kicker_enter",swShutter
  Shutter.IsDropped=0
  Ramp001.collidable=0
End Sub

Sub sw100_Hit
Shutter.IsDropped=1:Ramp001.collidable=1
End Sub
Sub sw101_Hit
Shutter.IsDropped=0:Ramp001.collidable=0
End Sub

'******************************************************
'           Joker hole
'******************************************************

Sub bsLowerleft(Enabled)
    Wall024.IsDropped=1 'Joker
End Sub

Sub Wall023_Slingshot
    RandomSoundSlingshotLeft Sling3
End Sub

'******************************************************
'           Vertical Kick to Gun
'******************************************************

'Variables used for VUK
 Dim raiseballsw, raiseball

 Sub TopVUK_Hit()
  TopVUK.Enabled=FALSE
  Controller.switch (51) = True
  PlaySound "kicker_enter_center"
  playsound "popper_ball"
 End Sub

 Sub VukTopPop(enabled)
  if(enabled and Controller.switch (51)) then
    TopVUK.DestroyBall
    Set raiseball = TopVUK.CreateBall
    playsound SoundFX("Popper",DOFContactors)
    raiseballsw = True
    TopVukraiseballtimer.Enabled = True 'Added by Rascal
    TopVUK.Enabled=TRUE
    Controller.switch (51) = False
  else

  end if
End Sub

 Sub TopVukraiseballtimer_Timer()
  If raiseballsw = True then
    raiseball.z = raiseball.z + 8   '8
    raiseball.x = raiseball.x + 0.02
    If raiseball.z > 240 then
      TopVUK.Kick 45, 16
      playsound "fx_metalrolling", 1 ,1
      Set raiseball = Nothing
      TopVukraiseballtimer.Enabled = False
      raiseballsw = False
    End If
  End If
 End Sub

'******************************************************
'             Ball Control
'******************************************************

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys


Sub Table1_KeyDown(ByVal keycode)

  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
  If KeyCode = PlungerKey Then Plunger.PullBack:SoundPlungerPull()
  If keycode = LeftTiltKey Then Nudge 90, 4 : SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 4 : SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 5 : SoundNudgeCenter()
  If keycode = LeftMagnaSave Then bLutActive = True
  If keycode = RightMagnaSave Then
    If bLutActive Then NextLUT: End If
  End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25

    End Select
  End If
  if keycode=StartGameKey then soundStartButton()
    ' Manual Ball Control
  If keycode = 46 Then          ' C Key
    If EnableBallControl = 1 Then
      EnableBallControl = 0
    Else
      EnableBallControl = 1
    End If
  End If
    If EnableBallControl = 1 Then
    If keycode = 48 Then        ' B Key
      If BCboost = 1 Then
        BCboost = BCboostmulti
      Else
        BCboost = 1
      End If
    End If
    If keycode = 203 Then BCleft = 1  ' Left Arrow
    If keycode = 200 Then BCup = 1    ' Up Arrow
    If keycode = 208 Then BCdown = 1  ' Down Arrow
    If keycode = 205 Then BCright = 1 ' Right Arrow
  End If

If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
  If KeyCode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If
  If keycode = LeftMagnaSave Then bLutActive = False


    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If

  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'*********
'   LUT
'*********

Dim bLutActive, LUTImage
Sub LoadLUT
  dim x
  bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 12: UpdateLUT: SaveLUT: End Sub

Sub UpdateLUT
Select Case LutImage
Case 0: table1.ColorGradeImage = "LUT0"
Case 1: table1.ColorGradeImage = "LUT1"
Case 2: table1.ColorGradeImage = "LUT2"
Case 3: table1.ColorGradeImage = "LUT3"
Case 4: table1.ColorGradeImage = "LUT4"
Case 5: table1.ColorGradeImage = "LUT5"
Case 6: table1.ColorGradeImage = "LUT6"
Case 7: table1.ColorGradeImage = "LUT7"
Case 8: table1.ColorGradeImage = "LUT8"
Case 9: table1.ColorGradeImage = "LUT9"
Case 10: table1.ColorGradeImage = "LUT10"
Case 11: table1.ColorGradeImage = "LUT11"
End Select
End Sub

'******************************************************
'            Gi & GI Effects
'******************************************************

dim fon,foff
dim GIState: GIState = 1

GiEffects.Enabled=1
GiMaskOff.Enabled=1:foff=1



'Playfield GI
Sub PlayGI(Enabled)
  dim xx, DNuts
  If Enabled Then
    For each xx in GI:xx.State = 0: Next
    GiMaskOn.Enabled=1:fon=1
    For each DNuts in DomeNuts:DNuts.BlendDisableLighting = 0: Next
    L117.intensity=22
    L117a.intensity=22
    L118.intensity=100
    L118a.intensity=100
    'If mode = 2 Then
      BigRamp002.image = ramp2
    'End If
    GIState = 0
  Else
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay",0,0.2,0,0.02
    GiMaskOff.Enabled=1:foff=1
    For each DNuts in DomeNuts:DNuts.BlendDisableLighting = 0.04: Next
    L117.intensity=200
    L117a.intensity=400
    L118.intensity=400
    L118a.intensity=400
    'If mode = 2 Then
      BigRamp002.image = ramp1
    'End If
    GIState = 1
  End If
End Sub

Sub GiEffects_timer()
'debug.print gi001.Statev
    FadeDisableLighting PlayfieldBEdge, 0.06, 0, 0.02, 0.02 'Playfield Edge
    FadeDisableLighting Primitive207, 0.1, 0, 0.033, 0.033 'Plastics
    FadeDisableLighting PlayfieldB, 0.1, 0, 0.033, 0.033 'Top PF
    FadeDisableLighting Primcab, 0.07, 0, 0.023, 0.023 'Cabinet
    FadeDisableLighting Primitive200, 0.06, 0, 0.02, 0.02 'Cabinet back
    FadeDisableLighting Primitive033, 0.1, 0, 0.033, 0.033 'Gun
    FadeDisableLighting Primitive009, 0.05, 0, 0.016, 0.016 'Top Plastics 1
    FadeDisableLighting Primitive030, 0.05, 0, 0.016, 0.016 'Top Plastics 2
    FadeDisableLighting Primitive039, 0.2, 0, 0.066, 0.066 'Top Plastics 3
    FadeDisableLighting WireRamp, 0.9, 0.2, 0.233, 0.233 'WireRamp
    FadeDisableLighting WireKickRamp, 0.45, 0, 0.15, 0.15 'WireKickRamp
    FadeDisableLighting LeftFlipperPrim, 0.1, 0, 0.033, 0.033 'Left Flipper
    FadeDisableLighting RightFlipperPrim, 0.1, 0, 0.033, 0.033 'Right Flipper
    FadeDisableLighting RightFlipperUpper, 0.1, 0, 0.033, 0.033 'Right Upper Flipper
    'FadeDisableLighting Primitive086, 0.1, 0, 0.033, 0.033 'Screw
    'FadeDisableLighting Primitive087, 0.1, 0, 0.033, 0.033 'Screw
    'FadeDisableLighting Primitive088, 0.1, 0, 0.033, 0.033 'Screw
    'FadeDisableLighting Primitive089, 0.1, 0, 0.033, 0.033 'Screw
    FadeDisableLighting Primitive002, 1, 0, 0.333, 0.333 'Lanes
    'FadeDisableLighting BigRamp002, 1, 0.4, 0.2, 0.2 'Big ramp
    FadeDisableLighting BigRamp001, 0.6, 0.2, 0.066, 0.066 'ramp bit
    FadeDisableLighting Primitive107, 0.35, 0, 0.116, 0.116 'Top Lane
    FadeDisableLighting Primitive15, 0.35, 0, 0.116, 0.116 'Top Lane
    FadeDisableLighting Primitive006, 0.35, 0, 0.116, 0.116  'Top Lane
    FadeDisableLighting Primitive007, 0.35, 0, 0.116, 0.116  'Top Lane
    FadeDisableLighting Primitive008, 0.35, 0, 0.116, 0.116  'Top Lane
    FadeDisableLighting Primitive003, 0.08, 0, 0.026, 0.026 'PWall
'Ramp Supports
    'FadeDisableLighting Primitive032, 0.03, 0, 0.01, 0.01
    'FadeDisableLighting Primitive037, 0.03, 0, 0.01, 0.01
    'FadeDisableLighting Primitive065, 0.03, 0, 0.01, 0.01
    'FadeDisableLighting Primitive059, 0.03, 0, 0.01, 0.01
    FadeDisableLighting Primitive015, 0.1, 0, 0.033, 0.033
    FadeDisableLighting Primitive061, 0.1, 0, 0.033, 0.033
'Top Pegs
    FadeDisableLighting PegPlasticT021, 0.2, 0, 0.066, 0.066
    FadeDisableLighting PegPlasticT001, 0.2, 0, 0.066, 0.066
    FadeDisableLighting PegPlasticT002, 0.2, 0, 0.066, 0.066
    FadeDisableLighting PegPlasticT003, 0.2, 0, 0.066, 0.066
    FadeDisableLighting PegPlasticT004, 0.2, 0, 0.066, 0.066
    FadeDisableLighting PegPlasticT005, 0.2, 0, 0.066, 0.066
    FadeDisableLighting PegPlasticT006, 0.2, 0, 0.066, 0.066
    FadeDisableLighting PegPlasticT007, 0.2, 0, 0.066, 0.066
    FadeDisableLighting PegPlasticT008, 0.2, 0, 0.066, 0.066
    FadeDisableLighting PegPlasticT009, 0.2, 0, 0.066, 0.066
'Rubber Posts
    FadeDisableLighting RubberPostT023, 0.12, 0, 0.04, 0.04
    FadeDisableLighting RubberPostT022, 0.12, 0, 0.04, 0.04
    FadeDisableLighting RubberPostT021, 0.12, 0, 0.04, 0.04
    FadeDisableLighting RubberPostT020, 0.12, 0, 0.04, 0.04
    FadeDisableLighting RubberPostT019, 0.12, 0, 0.04, 0.04
    FadeDisableLighting RubberPostT018, 0.12, 0, 0.04, 0.04
    FadeDisableLighting RubberPostT017, 0.12, 0, 0.04, 0.04

End Sub

Sub GiMaskOn_timer()

  Select Case fon
    Case 1:FlasherGI001.IntensityScale = 0.66:FlasherUPPFShadow.IntensityScale = 0.66:fon=2
    Case 2:FlasherGI001.IntensityScale = 0.33:FlasherUPPFShadow.IntensityScale = 0.88:fon=3
    Case 3:FlasherGI001.IntensityScale = 0:FlasherUPPFShadow.IntensityScale = 1:fon=4
    Case 4:GiMaskOn.Enabled=0
  End Select

End Sub

Sub GiMaskOff_timer()

  Select Case foff
    Case 1:FlasherGI001.IntensityScale = 0.33:FlasherUPPFShadow.IntensityScale = 0.88:foff=2
    Case 2:FlasherGI001.IntensityScale = 0.66:FlasherUPPFShadow.IntensityScale = 0.66:foff=3
    Case 3:FlasherGI001.IntensityScale = 1:FlasherUPPFShadow.IntensityScale = 0.44:foff=4
    Case 4:GiMaskOff.Enabled=0
  End Select

End Sub


'Fade DisableLighting (object, starting brightness, ending brightness, speed on, speed off)

Sub FadeDisableLighting(a, alvlu, alvld, spdu, spdd)

  Select Case gi001.State

    Case 1:
      a.UserValue = a.UserValue + spdu
        If a.UserValue > alvlu Then
          a.UserValue = alvlu
        end If
      a.BlendDisableLighting = a.UserValue 'On

    Case 0:
      a.UserValue = a.UserValue - spdd
        If a.UserValue < alvld Then
          a.UserValue = alvld
        end If
      a.BlendDisableLighting = a.UserValue 'Off

    End Select

End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
'       Plus Anti fading disabled lighting
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingState(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)
Dim PO199:PO199 = 1

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingState(chgLamp(ii, 0)) = chgLamp(ii, 1) + 3 'fading step
        Next
    End If
    'UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps

  Lamp 1, L1
  Lamp 2, L2
  Lamp 3, L3
  Lamp 4, L4
  Lamp 5, L5
  Lamp 6, L6
  Lamp 7, L7
  Lamp 8, L8
  Lamp 9, L9
  Lamp 10, L10
  Lamp 11, L11
  Lamp 12, L12
  Lamp 13, L13
  Lamp 14, L14
  Lamp 15, L15
  Lamp 16, L16
  Lamp 17, L17
  Lamp 18, L18
  Lamp 19, L19
  Lamp 20, L20

' Not Used
' Lamp 21, L21
' Lamp 22, L22
' Lamp 23, L23
' Lamp 24, L24
' Lamp 25, L25
'---------------
  Lamp 26, L26
  Lamp 27, L27
  Lamp 28, L28
  Lamp 29, L29
  Lamp 30, L30

' Not Used
' Lamp 31, L31
'---------------
' Lamp 32, L32 'Game Start
  Lamp 33, L33
  Lamp 34, L34
  Lamp 35, L35
  Lamp 36, L36
'---------------
    Lampm 37, L37
  FadeObjm 37, P37, "LED-Red_ON","LED-Red_2","LED-Red_1", "LED-Red"
  Flashm 37, F37
    FDL 37, P37, 0.8 , 0
'---------------
    Lampm 38, L38
  FadeObjm 38, P38, "LED-Green_ON","LED-Green_2","LED-Green_1", "LED-Green"
  Flashm 38, F38
    FDL 38, P38, 0.8 , 0
'---------------
    Lampm 39, L39
  FadeObjm 39, P39, "LED-Red_ON","LED-Red_2","LED-Red_1", "LED-Red"
  Flashm 39, F39
    FDL 39, P39, 0.8 , 0
'---------------
    Lampm 40, L40
  FadeObjm 40, P40, "LED-Yellow_ON","LED-Yellow_2","LED-Yellow_1", "LED-Yellow"
  Flashm 40, F40
    FDL 40, P40, 0.8 , 0
'---------------
    Lamp 41, L41
    Lamp 42, L42
    Lamp 43, L43
    Lamp 44, L44
    Lamp 45, L45
    Lamp 46, L46
    Lamp 47, L47
'---------------
    Lampm 48, L48
  FadeObjm 48, P48, "LED-Red_ON","LED-Red_2","LED-Red_1", "LED-Red"
    FDL 48, P48, 0.8 , 0
'---------------
    Lampm 49, L49
  FadeObjm 49, P49, "LED-Blue_ON","LED-Blue_2","LED-Blue_1", "LED-Blue"
  Flashm 49, F49
    FDL 49, P49, 0.8 , 0
'---------------
    Lampm 50, L50
  FadeObjm 50, P50, "LED-Yellow_ON","LED-Yellow_2","LED-Yellow_1", "LED-Yellow"
  Flashm 50, F50
    FDL 50, P50, 0.8 , 0
'---------------
    Lampm 51, L51
  FadeObjm 51, P51, "LED-Red_ON","LED-Red_2","LED-Red_1", "LED-Red"
  Flashm 51, F51
    FDL 51, P51, 0.8 , 0
'---------------
    Lamp 52, L52
    Lamp 53, L53
    Lamp 54, L54
'---------------
    Lampm 55, L55a
    Lampm 55, L55
  Flashm 55, F55
  FDLm 55, Primitive035, 0.15, 0
  FDLm 55, Primitive103, 1, 0
  FDL 55, Primitive7, 0.3, 0
'---------------
    Lampm 56, L56a
    Lampm 56, L56
  Flashm 56, F56
  FDLm 56, Primitive118, 0.15, 0
  FDLm 56, Primitive104, 1, 0
  FDL 56, Primitive005, 0.3, 0
'---------------
    Lamp 57, L57
    Lamp 58, L58
    Lamp 59, L59
    Lamp 60, L60
    Lamp 61, L61
    Lamp 62, L62
    Lamp 63, L63
    Lamp 64, L64
    Lamp 65, L65
    Lamp 66, L66
    Lamp 67, L67
    Lamp 68, L68
  Lamp 69, L69
  Lamp 70, L70
  Lamp 71, L71
  Lamp 72, L72
  Lamp 73, L73
  Lamp 74, L74
  Lamp 75, L75
  Lamp 76, L76
  Lamp 77, L77
  Lamp 78, L78
  Lamp 79, L79
  Lamp 80, L80
  Lamp 81, L81
  Lamp 82, L82
  Lamp 83, L83
  Lamp 84, L84
  Lamp 85, L85
  Lamp 86, L86
  Lamp 87, L87
  Lamp 88, L88
  Lamp 89, L89
  Lamp 90, L90
  Lamp 91, L91
  Lamp 92, L92
  Lamp 93, L93
  Lamp 94, L94
  Lamp 95, L95
  Lamp 96, L96
'-------------------------------
'Flasher Left Btm
  FadeObjm 117 ,Flasherbase002, "Dome-Red_a", "Dome-Red_b", "Dome-Red_c", "Dome-Red"
  FDLm 117, Flasherbase002, 1, 0
  'FDLm 117, Wall005, 0.5, 0
  Lampm 117, L117a
  Lampm 117, L117
  Flashm 117, F117b
  Flashm 117, F117a
  Flash 117, F117
'-------------------------------
'Flasher Right Btm
  FadeObjm 118 ,Primitive029, "domeearbasered_2", "domeearbasered_2", "domeearbasered_1", "domeearbasered"
  FadeObjm 118 ,Primitive028, "domeearbasered_2", "domeearbasered_2", "domeearbasered_1", "domeearbasered"
  FDLm 118 ,Primitive029, 0.7, 0.1
  FDLm 118 ,Primitive028, 0.7, 0.1
  'FDLm 118, Wall004, 0.8, 0
  Lampm 118, L118a
  Lampm 118, L118
  Flashm 118, F118d
  Flashm 118, F118c
  Flashm 118, F118b
  Flashm 118, F118a
  Flash 118, F118
'-------------------------------
'Flasher Under Upper Playfield
  Lampm 119,L119
  FDLm 119 ,sw65, 0.4, 0
  FDLm 119 ,sw66, 0.4, 0
  FDLm 119 ,sw68, 0.4, 0
  FDLm 119 ,sw69, 0.4, 0
  FDLm 119 ,Primitive069, 0.5, 0
  Flash 119, F119a
'-------------------------------
'Flasher Left Plastic Ramp
  Lampm 120, L120a
  Lampm 120, L120
  Lampm 120, L120
  Flashm 120, F120a
  Flash 120, F120
'-------------------------------
'Flasher Left Cactus
  Lampm 121, L121
  Flashm 121, F121a
  Flashm 121, F121
  FDL 121, RampDecal003, 1, 0
'-------------------------------
'Flasher Right Cactus
  Lampm 122, L122
  Flashm 122, F122a
  Flashm 122, F122
  FDL 122, RampDecal004, 1, 0
'-------------------------------
'Flasher Under Ramp
  Lampm 123, L123a
  Lampm 123, L123b
  Lampm 123, L123
  Flashm 123, F123
  FDL 123, RampDecal001, 0.05, 0
'-------------------------------
'Backbox Fire
  Lampm 124, L124



If F119a.IntensityScale > 0.33 and PO199 = 0 then

  If GIState = 1  then

    FlasherUPPFShadow.IntensityScale = 0.33
    PO199 = 1

  Else

    FlasherUPPFShadow.IntensityScale = 0.22
    PO199 = 1

  End If

End If

If F119a.IntensityScale < 0.33 and PO199 = 1 then

  If GIState = 1  then

    FlasherUPPFShadow.IntensityScale = 0.44
    PO199 = 0

  Else

    FlasherUPPFShadow.IntensityScale = 1
    PO199 = 0

  End If

End If

End Sub

'*********************************************************************

' Lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 30 ' flasher fading speed
    LampTimer.Enabled = 1
    For x = 0 to 200
        LampState(x) = 0
        FadingState(x) = 3 ' used to track the fading state
        FlashLevel(x) = 0
    Next
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingState(nr) = abs(value) + 3
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingState(nr)
        Case 4:object.state = 1:FadingState(nr) = 0
        Case 3:object.state = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:object.state = 1
        Case 3:object.state = 0
    End Select
End Sub


' Flashers: 4 is on,3,2,1 fade steps. 0 is off

Sub Flash(nr, object)
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1:FadingState(nr) = 0
        Case 3:Object.IntensityScale = 0.66:FadingState(nr) = 2
        Case 2:Object.IntensityScale = 0.33:FadingState(nr) = 1
        Case 1:Object.IntensityScale = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1
        Case 3:Object.IntensityScale = 0.66
        Case 2:Object.IntensityScale = 0.33
        Case 1:Object.IntensityScale = 0
    End Select
End Sub


'Fade DisableLighting (object, start state, end state)


Sub FDL(nr, object, ss, es)

  Select Case FadingState(nr)

    Case 4:Object.BlendDisableLighting = ss:FadingState(nr) = 0
        Case 3:Object.BlendDisableLighting = ss / 1.4:FadingState(nr) = 2
        Case 2:Object.BlendDisableLighting = ss / 2.5:FadingState(nr) = 1
        Case 1:Object.BlendDisableLighting = es:FadingState(nr) = 0

  End Select

End Sub


Sub FDLm(nr, object, ss, es)

  Select Case FadingState(nr)

    Case 4:Object.BlendDisableLighting = ss
        Case 3:Object.BlendDisableLighting = ss / 1.4
        Case 2:Object.BlendDisableLighting = ss / 2.5
        Case 1:Object.BlendDisableLighting = es

  End Select

End Sub


'Fade DisableLighting two state (object, start state, end state)

Sub NFDLm(nr, object, starts, ends)

  Select Case FadingState(nr)

    Case 4:Object.BlendDisableLighting = starts
        Case 3:Object.BlendDisableLighting = ends


  End Select

End Sub


' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub Reel(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1:FadingState(nr) = 0
        Case 3:object.SetValue 2:FadingState(nr) = 2
        Case 2:object.SetValue 3:FadingState(nr) = 1
        Case 1:object.SetValue 0:FadingState(nr) = 0
    End Select
End Sub

Sub Reelm(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1
        Case 3:object.SetValue 2
        Case 2:object.SetValue 3
        Case 1:object.SetValue 0
    End Select
End Sub

Sub NFadeReel(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1:FadingState(nr) = 1
        Case 3:object.SetValue 0:FadingState(nr) = 0
    End Select
End Sub

Sub NFadeReelm(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1
        Case 3:object.SetValue 0
    End Select
End Sub

'Texts

Sub Text(nr, object, message)
    Select Case FadingState(nr)
        Case 4:object.Text = message:FadingState(nr) = 0
        Case 3:object.Text = "":FadingState(nr) = 0
    End Select
End Sub

Sub Textm(nr, object, message)
    Select Case FadingState(nr)
        Case 4:object.Text = message
        Case 3:object.Text = ""
    End Select
End Sub

' Modulated Subs for the WPC tables

Sub SetModLamp(nr, level)
    FlashLevel(nr) = level / 150 'lights & flashers
End Sub

Sub LampMod(nr, object)          ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
    Object.State = 1             'in case it was off
End Sub

Sub FlashMod(nr, object)         'sets the flashlevel from the SolModCallback
    Object.IntensityScale = FlashLevel(nr)
End Sub

'Walls and mostly Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingState(nr)
        Case 4:object.image = a:FadingState(nr) = 0 'fading to off...
        Case 3:object.image = b:FadingState(nr) = 2
        Case 2:object.image = c:FadingState(nr) = 1
        Case 1:object.image = d:FadingState(nr) = 0
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingState(nr)
        Case 4:object.image = a
        Case 3:object.image = b
        Case 2:object.image = c
        Case 1:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingState(nr)
        Case 4:object.image = a:FadingState(nr) = 0 'off
        Case 3:object.image = b:FadingState(nr) = 0 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingState(nr)
        Case 4:object.image = a
        Case 3:object.image = b
    End Select
End Sub



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 50
  RandomSoundSlingshotRight Sling1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 49
  RandomSoundSlingshotLeft Sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub LeftTopSlingShot_Slingshot
  vpmTimer.PulseSw 32
  RandomSoundSlingshotLeft Sling3
    LTSling.Visible = 0
    LTSling1.Visible = 1
    sling3.rotx = 20
    LStep = 0
    LeftTopSlingShot.TimerEnabled = 1
End Sub

Sub LeftTopSlingShot_Timer
    Select Case LStep
        Case 3:LTSLing1.Visible = 0:LTSLing2.Visible = 1:sling3.rotx = 10
        Case 4:LTSLing2.Visible = 0:LTSLing.Visible = 1:sling3.rotx = 0:LeftTopSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1       'Do Not Change - default setting
BCvel = 4       'Controls the speed of the ball movement
BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0.0000001
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
  End If
End Sub

'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

'Rotate Flipper prims

  LeftFlipperPrim.RotZ = LeftFlipper.CurrentAngle
  RightFlipperPrim.RotZ = RightFlipper.CurrentAngle
  RightFlipperUpper.RotZ = Flipper1.CurrentAngle

End Sub

'********************************************************************
'      Rolling Sounds & Ballshadow
'********************************************************************

Const tnob = 19 ' total number of balls
Const lob = 0   'number of locked balls
Const maxvel = 54 'max ball velocity
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

Sub RollingTimer_Timer()
    Dim BOT, b, speedfactorx, speedfactory
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
                    rolling(b) = True
          PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)) + 25000, 1, 0, AudioFade(BOT(b))
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


    ' render the shadow for each ball

        If BOT(b).X < tablewidth/2 Then
            aBallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (tablewidth/2))/7)) + 6
        Else
            aBallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (tablewidth/2))/7)) - 6
        End If
        aBallShadow(b).Y = BOT(b).Y + 12


    'Shadow for top playfield
    if BOT(b).Y < 950 and BOT(b).Z > 99  then
      aBallShadow(b).Height = 100
    Else
      aBallShadow(b).Height = 0.1
    end if


        If BOT(b).Z > 400 Then
            aBallShadow(b).visible = 0
        Else
            aBallShadow(b).visible = 1
        End If




        ' jps ball speed control
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub


'--------------------------------------------------------------------------------------------------------------------


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
RollingSoundFactor = 2

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
WallImpactSoundFactor = 0.075/7                                                                                       'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/7
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers, Ramps and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel, WireRampSoundFactor, LoopSoundFactor

GateSoundLevel = 0.5/5                                                                                                        'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.25
WireRampSoundFactor = 2 * 100
LoopSoundFactor = 1 * 20


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

'--------------------------------------------------------------------------------------------------------------------

' Ramps sounds
Sub RampSound001_Hit: PlaySoundAtLevelActiveBall ("ballhit"), Vol(ActiveBall) * WireRampSoundFactor: End Sub
Sub RampSound003_Hit: PlaySoundAtLevelActiveBall ("rail"), Vol(ActiveBall) * WireRampSoundFactor: End Sub
Sub RampSound005_Hit: PlaySoundAtLevelActiveBall ("kicker_enter_center"), Vol(ActiveBall) * LoopSoundFactor: End Sub
Sub RampSound006_Hit: PlaySoundAtLevelActiveBall ("kicker_enter_center"), Vol(ActiveBall) * LoopSoundFactor: End Sub
Sub RampSound007_Hit: PlaySoundAtLevelActiveBall ("kicker_enter_center"), Vol(ActiveBall) * LoopSoundFactor: End Sub
Sub RampSound008_Hit: PlaySoundAtLevelActiveBall ("fx_ball_drop4"), Vol(ActiveBall) * LoopSoundFactor: End Sub
Sub RampSound009_Hit: PlaySound "kicker_enter_center" End Sub

' Stop Ramps Sounds
Sub RampSound002_Hit: StopSound "rail": End Sub
Sub RampSound004_Hit: StopSound "fx_metalrolling":PlaySoundAtLevelActiveBall ("fx_ball_drop4"), Vol(ActiveBall): End Sub



'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************


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
                                'playsound "fx_knocker"
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
'*****************
Const PI = 3.1415927

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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub

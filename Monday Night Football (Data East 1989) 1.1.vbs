Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

'agentyEighty6
' Add nFozzy flippers - done
' Add Flupper flashers - done (layer 8)
' Add Insert flashers - done (layer 8)
' Add dampeners - done (layer 9)
' Enhance lights and bulbs on shooter/kickoff ramp - done
' change pf inserts to bulbs to reflect light onto ball - done
' fix center ramp (was operating opposite of real Table1) - done
' correct back wall lights - done
' add flashers for top rollover light reflections - done
' Convert fall through kickers to holes on shooter/kickoff ramp - done
' Tune flashers - done
' clean up upper right ramp area - done
' Add Fleep sound package - code and sounds added - done
' Tweak physics - done
' Added misc prims (goalpost, top gates, etc.) - done
' Chg VUK to be more realistic (shoot ball up vs destroy/create) - done
' Add flasher reflection to back and side walls - done
' Fix side walls and FS/DT - done
' Add metal ramp triggers/code - done


'To DO
' clean up insert shapes - do later
' Create prims for ramps - add at a later time
'   Find Better plastics scans


' ****USER OPTIONS****
Dim DynamicFlashLevel, FlipperCoilRampupMode

DynamicFlashLevel= 0.85 'set between 0 and 1 to adjust flasher brightness (default = 0.85)

'Flipper Rampup mode: 0 = fast, 1 = medium, 2 = slow (tap passes should work)
FlipperCoilRampupMode = 1

'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.6

'************************************


Const UseSolenoids=1,UseLamps=0,UseGI=0, UseSync=1
Const debugon=False

Const BallSize = 50
Const BallMass = 1
' Standard Sounds
Const SSolenoidOn = "Jon_Relay_On_2"
Const SSolenoidOff = "Jon_Relay_Off_2"
Const SCoin = "Coin_In_2"

Const cGameName="mnfb_c27"

LoadVPM "01120100","DE.VBS",3.1


Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
  Ramp001.visible=1
  Ramp002.visible=1
  l42.visible=1
  l43.visible=1
  l44.visible=1
  l60.visible=1
  l61.visible=1
  l62.visible=1
Else
  Primitive001.size_y=150
  Ramp001.visible=0
  Ramp002.visible=0
  l42.visible=0
  l43.visible=0
  l44.visible=0
  l60.visible=0
  l61.visible=0
  l62.visible=0
End if


'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------

'FLASHERS

SolCallback(1)="Flash101"
SolCallback(2)="Flash102"
SolCallback(3)="Flash103"
SolCallback(4)="Flash104"
SolCallback(5)="Flash105"
SolCallback(6)="Flash106"
SolCallback(7)="Flash107"
SolCallback(8)="Flash108"
SolCallback(9)="Flash109"
SolCallback(10)="InsertFlasherRelay"
SolCallback(11)= "PFGI"
'SolCallback(12)="Flash112"
SolCallback(13)="Flash113"

SolCallback(14)= "bsREject.SolOut"
SolCallback(15)= "ChangeRamp"
SolCallback(16)= "bsVuk.SolOut"
SolCallback(25)= "bsTrough.SolIn"
SolCallback(26)= "bsTrough.SolOut"
SolCallback(27)= "bsKickBig.SolOut"
SolCallback(28)= "bsPit.SolOut"
SolCallback(29)= "dtC.SolDropUp"
SolCallback(30)= "dtR.SolDropUp"

SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(sLLFlipper)="SolLFlipper"

Const ReflipAngle = 20
Sub SolLFlipper(Enabled)
     If Enabled Then
    LF.fire
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
    RF.fire
        If RightFlipper.currentangle < RightFlipper.endangle + ReflipAngle Then
            RandomSoundReflipUpRight RightFlipper
        Else
            SoundFlipperUpAttackRight RightFlipper
            RandomSoundFlipperUpRight RightFlipper
        End If
     Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle < RightFlipper.startAngle - 5 Then
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


'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'Playfield GI
Sub PFGI(Enabled)
  If Enabled Then
    dim xx
    For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
  Else
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
  End If
End Sub


'--------------------------
'------  Table Init  ------
'--------------------------

Dim bsTrough,dtC,dtR,bsVuk,bsPit,bsREject,bsKickBig, RampPos, RampUp

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

  On Error Resume Next
    With Controller
      .GameName=cGameName
      If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
      .SplashInfoLine = "Monday Night Football, Data East 1989"
      .HandleMechanics=0
      .HandleKeyboard=0
      .ShowDMDOnly=1
      .ShowFrame=0
      .ShowTitle=0
      .DIP(0)=&H00
            .hidden=NOT Table1.ShowDT
      .Run
      If Err Then MsgBox Err.Description
    End With
  On Error Goto 0

    referee2.visible=True:referee1.visible=False

  vpmNudge.TiltSwitch=swTilt
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
    vpmNudge.SolGameOn True

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 10,13,12,11,0,0,0,0
  bsTrough.InitKick BallRelease, 45,6
    bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors),SoundFX("Solenoid",DOFContactors)
  bsTrough.Balls=3

  Set bsVUK=New cvpmBallStack
  bsVUK.InitSaucer Kicker2VUK,37,180,50
  bsVUK.KickZ = 1.56
  bsVUK.InitExitSnd SoundFX("kicker_out",DOFContactors), SoundFX("fx_saucerHit",DOFContactors)

  Set bsPit=New cvpmBallStack
  bsPit.InitSaucer Kicker1,26,180,15
  bsPit.InitExitSnd SoundFX("fx_saucer_exit",DOFContactors),SoundFX("fx_saucerHit",DOFContactors)

  Set bsREject=New cvpmBallStack
  bsREject.InitSaucer Kicker3,39,180,1
    bsREject.InitExitSnd SoundFX("fx_saucer_exit",DOFContactors),SoundFX("fx_saucerHit",DOFContactors)

  Set bsKickBig=New cvpmBallStack
  bsKickBig.InitSaucer Kicker12,29,0,50
  bsKickBig.InitExitSnd SoundFX("Saucer_Kick",DOFContactors),SoundFX("solenoid",DOFContactors)
  bsKickBig.KickForceVar=10

  Set dtC=New cvpmDropTarget
  dtC.InitDrop Array(target1,target2,target3),Array(35,34,33)
  dtC.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

  Set dtR=New cvpmDropTarget
  dtR.InitDrop Array(target4, target5, target6),Array(41,42,43)
  dtR.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

  dim xx
  For each xx in GI:xx.intensity = xx.intensity * 2: Next

  RampPos=0
    RampUp=False
  'GoalPost.Z=5
  CenterRamp0.visible=1
  CenterRamp0.collidable=1
  Controller.Switch(23)=0
  Controller.Switch(24)=1

  Wall25.IsDropped=1
  Wall26.IsDropped=1
  Wall27.IsDropped=1

  InsertFlasherSw = False

End Sub


Sub Table1_KeyDown(ByVal KeyCode)

  If keycode = LeftFlipperKey Then
    lfpress = 1
  End If

  If keycode = RightFlipperKey Then
    rfpress = 1
  End If

  If vpmKeyDown(KeyCode) Then Exit Sub
  If keycode = LeftMagnaSave Then bLutActive = True
  If keycode = RightMagnaSave Then
    If bLutActive Then NextLUT: End If
    End If
  If KeyCode=PlungerKey Then Plunger.Pullback:SoundPlungerPull()
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
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
  If vpmKeyUp(KeyCode) Then Exit Sub
    If keycode = LeftMagnaSave Then bLutActive = False
  If KeyCode=PlungerKey Then SoundPlungerReleaseBall():Plunger.Fire
End Sub


'------------------------------
'------  Trough Handler  ------
'------------------------------

Sub Drain_Hit:bsTrough.AddBall Me :RandomSoundDrain(Drain) : End Sub            '10/11/12/13

'---------------------------
'------  Wired Targets ------
'---------------------------

Sub sw14_Hit:Controller.Switch(14)=1:Playsound "rollover":End Sub '14
Sub sw14_unHit:Controller.Switch(14)=0:End Sub
Sub sw17_Hit:Controller.Switch(17)=1:Playsound "rollover":End Sub '17
Sub sw17_unHit:Controller.Switch(17)=0:End Sub
Sub sw18_Hit:Controller.Switch(18)=1:Playsound "rollover":End Sub         '18
Sub sw18_unHit:Controller.Switch(18)=0:referee1.visible=False:referee2.visible=True:End Sub
Sub sw19_Hit:Controller.Switch(19)=1:Playsound "rollover":End Sub         '19
Sub sw19_unHit:Controller.Switch(19)=0:End Sub
Sub sw20_Hit:Controller.Switch(20)=1:Playsound "rollover":End Sub         '20
Sub sw20_unHit:Controller.Switch(20)=0:End Sub

Sub sw29_Hit:Controller.Switch(29)=1:Playsound "rollover":End Sub
Sub sw29_unHit:Controller.Switch(29)=0:End Sub
Sub sw30_Hit:Controller.Switch(30)=1:Playsound "rollover":End Sub
Sub sw30_unHit:Controller.Switch(30)=0:End Sub
Sub sw31_Hit:Controller.Switch(31)=1:Playsound "rollover":End Sub
Sub sw31_unHit:Controller.Switch(31)=0:End Sub
Sub sw32_Hit:Controller.Switch(32)=1:Playsound "rollover":End Sub
Sub sw32_unHit:Controller.Switch(32)=0:End Sub

Sub sw25_Hit:Controller.Switch(25)=1:Playsound "rollover":End Sub         '25
Sub sw25_unHit:Controller.Switch(25)=0:End Sub

Sub sw27_Hit:Controller.Switch(27)=1:Playsound "rollover":End Sub         '27
Sub sw27_unHit:Controller.Switch(27)=0:End Sub

'---------------------------
'------  Drop Targets ------
'---------------------------

Sub target3_Hit:dtC.Hit 3:End Sub                   '33
Sub target2_Hit:dtC.Hit 2:End Sub                   '34
Sub target1_Hit:dtC.Hit 1:End Sub                   '35

Sub target4_Hit:dtR.Hit 1:End Sub                   '41
Sub target5_Hit:dtR.Hit 2:End Sub                   '42
Sub target6_Hit:dtR.Hit 3:End Sub                   '43

'---------------------------
'------  Hit Targets ------
'---------------------------

Sub sw49_Hit:vpmTimer.PulseSw(49):PlayTargetSound:End Sub               '49
Sub sw50_Hit:vpmTimer.PulseSw(50):PlayTargetSound:End Sub               '50

'---------------------------
'------  Spinners ------
'---------------------------

Sub sw38_Spin:vpmTimer.PulseSw 38:playsound "fx_spinner":End Sub    '38
Sub sw36_Spin:vpmTimer.PulseSw 36:playsound "fx_spinner":End Sub    '36

'---------------------------
'------  Kicker ------
'---------------------------

' SoundSaucerKick 1, sw15
' RandomSoundBallRelease(LockRSol)


Sub Kicker1_Hit:bsPit.AddBall 0:End Sub                 '26

Sub Kicker2VUK_Hit: bsVUK.AddBall Me : End Sub

Sub Kicker3_Hit:bsREject.AddBall 0:End Sub                '39

Sub Kicker12_Hit:bsKickBig.AddBall 0:End Sub

'---------------------------
'----- Ramps Switches ------
'---------------------------

Sub sw44_Hit:Controller.Switch(44)=1:fakeswitch1.visible=False:End Sub              '44
Sub sw44_unHit:Controller.Switch(44)=0:fakeswitch1.visible=True:End Sub             '44

Sub sw45_Hit:Controller.Switch(45)=1:End Sub              '45
Sub sw45_unHit:Controller.Switch(45)=0:End Sub              '45

Sub RefTrigger_Hit:referee2.visible=False:referee1.visible=True:End Sub

Sub sw51_Hit:Controller.Switch(51)=1:End Sub              '51
Sub sw51_unHit:Controller.Switch(51)=0:End Sub              '51

'------------------------------
'------  Gates ------
'------------------------------

Sub Gate1_Hit:playsound "Gate5":End Sub
Sub Gate2_Hit:playsound "Gate5":End Sub
Sub Gate3_Hit:playsound "Gate5":End Sub

'------------------------------
'------  Bumpers ------
'------------------------------

Sub Bumper1_Hit:vpmTimer.pulseSW 46 : PlayBumperSound : End Sub
Sub Bumper2_Hit:vpmTimer.pulseSW 48 : PlayBumperSound : End Sub
Sub Bumper3_Hit:vpmTimer.pulseSW 47 : PlayBumperSound : End Sub


Sub PlayBumperSound()
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySound SoundFX("fx_bumper1",DOFContactors)
      Case 2 : PlaySound SoundFX("fx_bumper2",DOFContactors)
      Case 3 : PlaySound SoundFX("fx_bumper3",DOFContactors)
    End Select
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 22
    RandomSoundSlingshotRight(sling1)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 21
    RandomSoundSlingshotLeft(sling2)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub


'------------------------------
'------ Center Ramp Code ------
'------------------------------

Sub ChangeRamp(Enabled)
  If Enabled Then
        RampTimer.enabled=True
  End If
End Sub

Dim CRamp, p

Sub RampTimer_Timer
  playsound "fx_spinner"
  if  RampUp Then ' go down
    RampPos=RampPos-1
  Else
    RampPos=RampPos+1
  end if

  GoalPost.z=RampPos*5

  p=0
  For each CRamp in CenterRamp
    CRamp.visible=False
  CRamp.collidable=False
    if p=RampPos then
    CRamp.visible=True
    CRamp.collidable=True
  End If
    p=p+1
  Next
  If RampPos > 15 Then
  RampPos = 15
  RampUp=True
  Controller.Switch(23)=1
  Controller.Switch(24)=0

  CenterRamp15.visible=1
  CenterRamp15.collidable=1

  'Wall25.IsDropped=0
  'Wall26.IsDropped=0
  'Wall27.IsDropped=0

  GoalPost.z=75

    RampTimer.enabled=False
  End if

  If RampPos < 1 Then
  RampPos = 0
  RampUp=False
  Controller.Switch(23)=0
  Controller.Switch(24)=1

  CenterRamp0.visible=1
  CenterRamp0.collidable=1

  'Wall25.IsDropped=1
  'Wall26.IsDropped=1
  'Wall27.IsDropped=1

  GoalPost.z=0

    RampTimer.enabled=False
  End if
End Sub


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
'Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

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

NFadeLm 1, L1
NFadeL 1, L1b
NFadeL 2, L2
NFadeL 3, L3
NFadeL 4, L4
NFadeL 5, L5
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeLm 10, L10
NFadeL 10, L10b
NFadeL 11, L11
NFadeL 12, L12
NFadeL 13, L13
NFadeL 14, L14
NFadeL 15, L15
NFadeL 16, L16
NFadeL 17, L17
NFadeL 18, L18
NFadeLm 19, L19
NFadeL 19, L19b
NFadeL 20, L20
NFadeL 21, L21
NFadeL 22, L22
NFadeL 23, L23
'NFadeL 24, L24
NFadeL 25, L25
NFadeL 26, L26
NFadeL 27, L27
NFadeLm 28, L28
NFadeL 28, L28b
NFadeLm 29, L29
NFadeL 29, L29b
NFadeL 30, L30
NFadeL 31, L31
NFadeL 32, L32
NFadeL 33, L33
NFadeL 34, L34
NFadeL 35, L35
NFadeL 36, L36
NFadeLm 37, L37
NFadeL 37, L37b
NFadeL 38, L38
NFadeL 39, L39
NFadeL 40, L40
NFadeL 41, L41
NFadeL 42, L42
NFadeL 43, L43
NFadeL 44, L44
NFadeL 45, L45
NFadeLm 46, L46
NFadeL 46, L46b
NFadeL 47, L47
NFadeL 48, L48
NFadeL 49, L49
NFadeL 50, L50
NFadeL 51, L51
NFadeL 52, L52
NFadeL 53, L53
NFadeL 54, L54
NFadeLm 55, L55
NFadeL 55, L55b
NFadeL 56, L56
NFadeL 57, L57
NFadeL 58, L58
NFadeL 59, L59
NFadeL 60, L60
NFadeL 61, L61
NFadeL 62, L62
NFadeL 63, L63
NFadeLm 64, L64
NFadeL 64, L64b

End Sub


' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
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



'*************
'   JP'S LUT
'*************

Dim bLutActive, LUTImage
Sub LoadLUT
Dim x
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 10: UpdateLUT: SaveLUT: End Sub

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
  End Select
End Sub


sub GameTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle

  Primitive100.ObjRotY = Gate1.currentangle - 15
  Primitive006.ObjRotY = Gate3.currentangle - 15

  LF25.visible=L25.state
  LF26.visible=L26.state
  LF27.visible=L27.state

  LF33.visible=L33.state
  LF34.visible=L34.state
  LF35.visible=L35.state
  LF36.visible=L36.state

  LF49.visible=L49.state
  LF50.visible=L50.state
  LF51.visible=L51.state
  LF52.visible=L52.state
  LF53.visible=L53.state
  LF54.visible=L54.state

  Flasherflash1a.visible = Flasherflash1.visible
  Flasherflash1a.opacity = Flasherflash1.opacity

  Flasherflash2a.visible = Flasherflash2.visible
  Flasherflash2a.opacity = Flasherflash2.opacity

  Flasherflash4a.visible = Flasherflash4.visible
  Flasherflash4a.opacity = Flasherflash4.opacity

  Flasherflash6a.visible = Flasherflash6.visible
  Flasherflash6a.opacity = Flasherflash6.opacity

  Flasherflash7a.visible = Flasherflash7.visible
  Flasherflash7a.opacity = Flasherflash7.opacity
  Flasherflash7b.visible = Flasherflash7.visible
  Flasherflash7b.opacity = Flasherflash7.opacity

  If RampPos <= 5 then
    Wall25.IsDropped=1
    Wall26.IsDropped=1
    Wall27.IsDropped=1
  Else
    Wall25.IsDropped=0
    Wall26.IsDropped=0
    Wall27.IsDropped=0
  End If

End Sub



'****************************************************************************
' NEW ROUTINES 2021
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

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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

Sub RDampen_Timer()
Cor.Update
End Sub


Sub MRTrigger1Start_hit()
    Dim b
    BOT = GetBalls
  For b = 0 to UBound(BOT)
    if BOT(b) is activeball then
      MRBallArray(b) = True
    end if
  next
End Sub
Sub MRTrigger1End_hit()
    Dim b
    BOT = GetBalls
  For b = 0 to UBound(BOT)
    if BOT(b) is activeball then
      MRBallArray(b) = False
    end if
  next
End Sub

'*****************************************
'  JP's Rolling Sounds  + Sliderpoint's Ball Shadows
'*****************************************

Dim BallShadow, ShadowTrigger, Ballshadows
BallShadow = Array (BallShadow0,BallShadow1,BallShadow2,BallShadow3,BallShadow4)

Dim MRBallArray : MRBallArray = Array (False,False,False,False,False)

Ballshadows = 1

Dim tnob : tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Dim BOT

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 to tnob
    rolling(i) = False
  Next
End Sub

Sub RollingTimer_Timer
    Dim b
    BOT = GetBalls

    ' hide shadow + stop ballrolling sounds of deleted balls + stop ballrolling sounds
    If UBound(BOT)<(tnob-1) Then
      For b = (UBound(BOT)+1) to (tnob-1)
        BallShadow(b).visible = 0
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    StopSound("fx_Rolling_Metal" & b)
    MRBallArray(b) = False
      Next
    End If

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then
    Exit Sub
  End If

  For b = 0 to UBound(BOT)
    if BOT(b).z < 30 Then ' Ball on playfield, stop metal rolling sound
      MRBallArray(b) = False
    End If
    'Play sound for each ball
    If BallVel(BOT(b) ) > 0 Then
      rolling(b) = True
      if MRBallArray(b) = True Then
        StopSound("fx_ballrolling" & b)
        PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b)) * 0.2 * VolumeDial, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
      Else
        StopSound("fx_Rolling_Metal" & b)
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)) * 0.2 * VolumeDial, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)) * 0.8 * VolumeDial, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+15000, 1, 0, AudioFade(BOT(b) )
        End If
      End If
    Else
      If rolling(b) = True Then
        StopSound("fx_ballrolling" & b)
        StopSound("fx_Rolling_Metal" & b)
        rolling(b) = False
      End If
    End If

    ' render the shadow for each ball
        If (Ballshadows = 1) Then
      If BOT(b).Z < 30 and BOT(b).Z > 0 Then
        BallShadow(b).visible = 1
        If BOT(b).X < Table1.Width/2 Then
          BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
          BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
      Else
        BallShadow(b).visible = 0
      End If
        Else
      BallShadow(b).visible = 0
        End If

    '***Ball Drop Sounds***
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -5 Then
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



'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor
Dim tablewidth, tableheight : tablewidth = Table1.width : tableheight = Table1.height

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
Dim SlingshotSoundLevel, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                                                           'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
SlingshotSoundLevel = 0.95                                                                                                'volume level; range [0, 1]
'BumperSoundFactor = 4.25                                                                                                'volume multiplier; must not be zero
KnockerSoundLevel = 1                                                                                                         'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.75 '0.055/5                                                                                        'volume multiplier; must not be zero
RubberWeakSoundFactor = 1 '0.075/5                                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.5 '0.075/5                                                                                'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.05                                                                        'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.05                                                                        'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
WallImpactSoundFactor = 1 '0.075                                                                                        'volume multiplier; must not be zero
MetalImpactSoundFactor = 1 '0.075/3
SaucerLockSoundLevel = 1 '0.25
SaucerKickSoundLevel = 1 '0.5

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.25                                                                                                        'volume level; range [0, 1]
TargetSoundFactor = 0.25                                                                                     'volume multiplier; must not be zer
DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.5
SpinnerSoundLevel = 1                                                                     'volume level; range [0, 1]

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

Sub PlaySoundAtLevelExistingLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
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
        Vol = Csng(BallVel(ball) ^2 / 2000)
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
'Sub RandomSoundBumperTop(Bump)
'        PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
'End Sub

'Sub RandomSoundBumperMiddle(Bump)
'        PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
'End Sub

'Sub RandomSoundBumperBottom(Bump)
'        PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
'End Sub

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
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
        RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
        FlipperRightHitParm = parm/10
        If FlipperRightHitParm > 1 Then
                FlipperRightHitParm = 1
        End If
        FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
        RandomSoundRubberFlipper(parm)
End Sub

Sub UpperFlipperCollide(parm)
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
                 RandomSoundRubberStrong 2
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If
End Sub

Sub Posts_Hit(idx)
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 5 then
                 RandomSoundRubberStrong 3
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If
End Sub

Sub Pins_Hit(idx)
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

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Thin_Hit (idx)
  PlaySoundAtLevelActiveBall ("Metalhit_Thin"), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Medium_Hit (idx)
  PlaySoundAtLevelActiveBall ("Metalhit_Medium"), Vol(ActiveBall) * MetalImpactSoundFactor
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
        'If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
    If Abs(vol(activeball) > 30) then
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

'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

'Flipper Correction Initialization late 80s to early 90s
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
        AddPt "Polarity", 1, 0.05, -5
        AddPt "Polarity", 2, 0.4, -5
        AddPt "Polarity", 3, 0.6, -4.5
        AddPt "Polarity", 4, 0.65, -4.0
        AddPt "Polarity", 5, 0.7, -3.5
        AddPt "Polarity", 6, 0.75, -3.0
        AddPt "Polarity", 7, 0.8, -2.5
        AddPt "Polarity", 8, 0.85, -2.0
        AddPt "Polarity", 9, 0.9,-1.5
        AddPt "Polarity", 10, 0.95, -1.0
        AddPt "Polarity", 11, 1, -0.5
        AddPt "Polarity", 12, 1.1, 0
        AddPt "Polarity", 13, 1.3, 0

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
        'playsound "fx_knocker"
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
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim BOT, b

  If abs(Flipper1.currentangle) < abs(Endangle1) + 3 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
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
          'debug.print "flippernudge!!"
          BOT(b).velx = BOT(b).velx /1.5
          BOT(b).vely = BOT(b).vely - 1
        end If
      Next
    End If
  Else
    If abs(Flipper1.currentangle) > abs(Endangle1) + 30 then EOSNudge1 = 0
  End If
End Sub

'*****************
' Maths
'*****************
Const Pi = 3.1415927

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
Const EOSTnew = 1 '0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 4.5
  Case 2:
    SOSRampup = 8.5
End Select
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.035 '0.025


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
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

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
Const LiveDistanceMax = 117  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
  Dim LiveCatchBounce                             'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.8 Then            'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)  'Partial catch when catch happens a bit late
    end If

    'debug.print "Live catch! Bounce: " & LiveCatchBounce

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (16 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
  End If
End Sub

'*****************************************************************************************************
'END nFOZZY FLIPPERS'


Sub Flash101(Enabled)
  L101.state=Enabled
  If Enabled Then
    Flashlevel(1) = DynamicFlashLevel : FlasherFlash1_Timer
  End If
End Sub

Sub Flash102(Enabled)
  L102.state=Enabled
  If Enabled Then
    Flashlevel(2) = DynamicFlashLevel : FlasherFlash2_Timer
  End If
End Sub

Sub Flash103(Enabled)
  L103.state=Enabled
  If Enabled Then
    Flashlevel(3) = DynamicFlashLevel : FlasherFlash3_Timer
    Flashlevel(5) = DynamicFlashLevel : FlasherFlash5_Timer
  End If
End Sub

Sub Flash104(Enabled)
  L104.state=Enabled
End Sub

Sub Flash105(Enabled)
  L105.state=Enabled
End Sub

Sub Flash106(Enabled)
  L106.state=Enabled
End Sub

Sub Flash107(Enabled)
  L107.state=Enabled
End Sub

Sub Flash108(Enabled)
  L108.state=Enabled
End Sub

Sub Flash109(Enabled)
  Flashlevel(4) = DynamicFlashLevel * 0.75 : FlasherFlash4_Timer
  Flashlevel(6) = DynamicFlashLevel * 0.75 : FlasherFlash6_Timer
End Sub


Sub Flash113(Enabled)
  Flashlevel(7) = DynamicFlashLevel * 0.75 : FlasherFlash7_Timer
End Sub

Dim InsertFlasherSw
Sub InsertFlasherRelay(Enabled)
  InsertFlasherSw = Enabled
End Sub




Dim TestFlashers, TableRef

TestFlashers = 0    ' *** set this to 1 to check position of flasher object ***
Set TableRef = Table1   ' *** change this, if your table has another name       ***
Dim FlashLevel(20), objbase(20), objlit(20), objflasher(20), objlight(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "red"
InitFlasher 2, "red"
InitFlasher 3, "red"
InitFlasher 4, "red"
InitFlasher 5, "red"
InitFlasher 6, "red"
InitFlasher 7, "red"


Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)

  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    'objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    'objflasher(nr).RotZ = objbase(nr).ObjRotZ
    objflasher(nr).height = objbase(nr).z + 50
  End If

  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub


Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlashLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlashLevel(nr)^2.5
  objbase(nr).BlendDisableLighting =  10 * FlashLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * FlashLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,FlashLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  FlashLevel(nr) = FlashLevel(nr) * 0.9 - 0.01
  If FlashLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub
Sub FlasherFlash7_Timer() : FlashFlasher(7) : End Sub



'****************************************************************************

Sub Table1_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub

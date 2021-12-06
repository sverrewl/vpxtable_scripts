'Spirit - Gottlieb 1982
'----------------------
Option Explicit
Randomize

Dim Ballsize,BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01210000", "sys80.vbs", 3.1

Const cGameName="spirit"
Const UseSolenoids=2
Const UseLamps=0
Const UseGI=0
Const SSolenoidOn="SolOn"
Const SSolenoidOff="SolOff"
Const SFlipperOn="FlipperUp"
Const SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits=""
Dim Score(1), RampBallY
Dim B2SOn


'////// FLIPPER OPTIONS //////////

Const Flip_Link = 0 ' 0 For separate magnasave mini flippers - 1 to link with main flippers.

'/////////////////////////////////


'///// GAMMA - GRAPHICS OPTIONS ////

Const Flasher_Reduction = 50   '(0 to 100) 'On some monitors the playfield and plastics may look milky - by INCREASING this number you
                              'will make the flashers more transparent which will result in a blacker overall look.
                              'A setting of 0 will not make any changes to the current flasher opacity at all.
                              'A setting of 50 will reduce the flasher brightness by half.
                              'A setting of 100 will pretty much turn off the effect altogether.
                              '*NOTE  - you can use any number between 0 and 100 to fine tune.

Const GI_Modulation = 1   'Setting this to 1 will remove the milky look from the GI lamps that are apparent on some higher gamma
                          'monitors.  Leaving this at 0 will draw the GI the way it was originally set.

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

For each plasx in PlasFlash
plasx.opacity = plasx.opacity - Flasher_Reduction
Next

If GI_Modulation = 1 Then
For each FGIX in AFGI
FGIX.BulbModulateVsAdd = 1
Next
End If

Dim VarHidden
If ZZZ.ShowDT = true then
    VarHidden = 0
    ramp70.Visible = 1
    ramp73.Visible = 1
    MiddleWheel.visible = 1
    MiddleWheel1.visible = 1
else
    VarHidden = 1
    ramp70.Visible = 0
    ramp73.Visible = 0
    MiddleWheel.visible = 0
    MiddleWheel1.visible = 0
end if


Set LampCallback = GetRef("UpdateMultipleLamps")

dim LXM

Sub ZZZ_KeyDown(ByVal keycode)

    If keycode = LeftTiltKey Then Nudge 90, 2:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 2:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 2:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25

If keycode = PlungerKey Then Plunger.Pullback:PlaySoundat "PullbackPlunger", Plunger

If keycode = LeftMagnaSave AND Flip_Link = 0 AND Controller.Solenoid(10) = True Then
Flipper1.RotateToEnd
PlaySoundat SoundFXDOF("FlipperUP",140,DOFOn,DOFFlippers), Flipper1
End If

If keycode = RightMagnaSave AND Flip_Link = 0 AND Controller.Solenoid(10) = True Then
Flipper2.RotateToEnd
PlaySoundat SoundFXDOF("FlipperUP",141,DOFOn,DOFFlippers), Flipper2
Else
End If
If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub ZZZ_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundatVol "Plunger", Plunger, 0.1
    If keycode = LeftMagnaSave AND Flip_Link = 0 AND Controller.Solenoid(10) = True Then
Flipper1.RotateToStart
PlaySoundat SoundFXDOF("FlipperDown",140,DOFOff,DOFFlippers), Flipper1
End If
    If keycode = RightMagnaSave AND Flip_Link = 0 AND Controller.Solenoid(10) = True Then
Flipper2.RotateToStart
PlaySoundat SoundFXDOF("FlipperDown",141,DOFOff,DOFFlippers), Flipper2
End If
If vpmKeyUp(keycode)Then Exit Sub
End Sub

'constants
Const sLeftReset=1
Const sLLHole=2
Const sCoin1=3'ok
Const sCoin2=4'ok
Const sdtRight=5'ok
Const sCoin3=7'ok
Const sUCHole=6
Const sKnocker=8
Const sOutHole=9

SolCallback(sLeftReset)= "PlaySoundAt SoundFX(""droptarg"",DOFDropTargets),sw11:dtLeft.SolDropUp"       '"dtLeft.SolDropUp"
SolCallback(sdtRight)="PlaySoundAt SoundFX(""droptarg"",DOFDropTargets),sw15:dtRight.SolDropUp"    '"dtRight.SolDropUp"
SolCallback(sOutHole)="bsTrough.SolIn"
SolCallback(sLLHole)="SLC"
SolCallback(sUCHole)="STC"
SolCallback(sLLFlipper)="SLeftFlipHandle"
SolCallback(sLRFlipper)="SRightFlipHandle"
SolCallback(10) = "SolGI"
SolCallback(SKnocker)="SolKnocker"

Sub SolKnocker(Enabled)
  If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

Dim FGIX

Sub SolGI(Enabled)
If enabled Then
For each FGIX in AFGI
FGIX.state = 1
Next
For each plasx in PlasFlash
plasx.opacity = plasx.opacity + 5
PlasX.color = RGB(192,192,192)
Next
ZZZ.ColorGradeImage = "ColorGradeEx_2"
PlaySoundatvol "fx_relay_On", PegPlasticT21, 0.1
DOF 160,1
Else
For each FGIX in AFGI
FGIX.state = 0
Next
For each plasx in PlasFlash
plasx.opacity = plasx.opacity - 5
PlasX.color = RGB(0,0,192)
Next
ZZZ.ColorGradeImage = "ColorGradeEx_1"
PlaySoundatVol "fx_relay_On", PegPlasticT21, 0.1
DOF 160,0
End If
End Sub

Sub SLC(enabled)
If enabled Then
PlaySoundAt "popper", sw30
bsSaucer.ExitSol_On
End If
End Sub


Sub STC(enabled)
If enabled Then
PlaySoundAt "popper", sw12
bsSaucer1.ExitSol_On
End If
End Sub

Sub SLeftFlipHandle(Enabled)
If Enabled Then
  VpmSolFlipper LeftFlipper,Flipper3,True
    PlaySoundat SoundFX("FlipperUP",DOFFlippers), PLeftFlipper
    PlaySoundat SoundFX("FlipperUP",DOFFlippers), PLeftFlipper1
  If Flip_Link = 1 Then
    Flipper1.RotateToEnd
    PlaySoundat SoundFX("FlipperUP",DOFFlippers), Flipper1
    End If
  Else
  VpmSolFlipper LeftFlipper,Flipper3,False
    PlaySoundat SoundFX("FlipperDown",DOFFlippers), PLeftFlipper
    PlaySoundat SoundFX("FlipperDown",DOFFlippers), PLeftFlipper1
  If Flip_Link = 1 Then
    Flipper1.RotateToStart
    PlaySoundat SoundFX("FlipperDown",DOFFlippers), Flipper1
    End If
End If
End Sub

Sub SRightFlipHandle(Enabled)
If Enabled Then
  VpmSolFlipper RightFlipper,Flipper4,True
    PlaySoundat SoundFX("FlipperUP",DOFFlippers),PRightFlipper
    PlaySoundat SoundFX("FlipperUP",DOFFlippers), PRightFlipper1
  If Flip_Link = 1 Then
    Flipper2.RotateToEnd
    PlaySoundat SoundFX("FlipperUP",DOFFlippers),Flipper2
    End If
  Controller.Switch(64)=1
  Else
  VpmSolFlipper RightFlipper,Flipper4,False
    PlaySoundat SoundFX("FlippeDown",DOFFlippers),PRightFlipper
    PlaySoundat SoundFX("FlipperDown",DOFFlippers), PLeftFlipper1
  If Flip_Link = 1 Then
    Flipper2.RotateToStart
    PlaySoundat SoundFX("FlipperDown",DOFFlippers),Flipper2
    End If
  Controller.Switch(64)=0
End If
End Sub

Dim dtRight,dtLeft,bsTrough,bsSaucer,bsSaucer1,bsSaucer2,bsSaucer3,dtL,dtRed ' saucers and drop target declares!
Score(1)="000 000"

Sub ZZZ_Init
  On Error Resume Next
  With Controller
        vpmInit me
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine="Spirit - Gottlieb, 1983" & vbNewLine & "Table Design: Luvthatapex"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .Games(cGameName).Settings.Value("dmd_pos_x")=1380
        .Games(cGameName).Settings.Value("dmd_pos_y")=50
        .Games(cGameName).Settings.Value("dmd_red")=0
        .Games(cGameName).Settings.Value("dmd_green")=128
        .Games(cGameName).Settings.Value("dmd_blue")=192
    .Hidden=VarHidden
    'Run
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.SolMask(0)=0
    vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
  Controller.Run

  ' Main Timer init
  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

Set bsTrough=New cvpmBallstack
With bsTrough
  .InitSw 34,0,0,44,0,0,0,0
  .InitKick DrainKicker, 90, 3
  .InitExitSnd "",""
  .Balls=3
End With

  Set bsSaucer=New cvpmBallStack 'Set up right kicker pocket
  bsSaucer.InitSaucer sw30,30,100,15
  'bsSaucer.InitExitSnd "popper","popper"

  Set bsSaucer1=New cvpmBallStack 'Set up top right kicker pocket
  bsSaucer1.InitSaucer sw12,12,290,5
  'bsSaucer1.InitExitSnd "popper","popper"

  Set bsSaucer2=New cvpmBallStack 'Sets up the Left kicker pocket
  bsSaucer2.InitSaucer sw31,31,220,15
  'bsSaucer2.InitExitSnd "popper","popper"

  Set bsSaucer3=New cvpmBallStack 'Sets up the Upper left kicker pocket
  bsSaucer3.InitSaucer sw02,02,180,15
  'bsSaucer3.InitExitSnd "popper","popper"

  Set dtLeft=New cvpmDropTarget
  dtLeft.InitDrop Array(SW01,SW11,SW21),Array(1,11,21)
  'dtLeft.initsnd "",""

  Set dtRight=New cvpmDropTarget
  dtRight.InitDrop Array(SW05,SW15,SW25),Array(5,15,25)
  'dtRight.initsnd "",""

  ' Nudging
  vpmNudge.TiltSwitch=57
  vpmNudge.Sensitivity=3
  vpmNudge.TiltObj=Array(Bumper1,LeftSlingshot,RightSlingshot, MRSlingShot, MLSlingShot, CentreSling)

For each plasx in PlasFlash
plasx.opacity = plasx.opacity - 5
PlasX.color = RGB(0,0,192)
Next

End Sub

Sub Drain_Hit
  playsoundat "drain", Drain
  bstrough.addball me
End Sub

Sub sw30_Hit:bsSaucer.AddBall 0:PlaySoundat "kicker_enter_left", sw30:End Sub 'When the right saucer is hit, add a FAKE ball to the top saucer ballstack - the ball is still visible
Sub sw31_Hit:bsSaucer2.AddBall 0:PlaySoundat "kicker_enter_left", sw31:End Sub 'When the left pocket kicker is hit, add a FAKE ball to the left saucer ballstack - the ball is still visible
Sub sw02_Hit:bsSaucer3.AddBall 0:PlaySoundat "kicker_enter_left", sw02:End Sub 'When the top pocket kicker is hit, add a FAKE ball to the left saucer ballstack - the ball is still visible
Sub sw12_Hit:bsSaucer1.AddBall 0:PlaySoundat "kicker_enter_left", sw12:End Sub 'When the top pocket kicker is hit, add a FAKE ball to the top right saucer ballstack - the ball is still visible
Sub Drainkicker_Unhit:End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0)) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()

NfadeL 0, L0
NfadeL 2, L2


NfadeL 1, L1

NfadeL 3, L3
NfadeL 4, L4
NfadeL 5, L5

NfadeL 16, L16
NfadeL 17, L17

Nfadel 18, L18
NfadeL 19, L19

NfadeL 20, L20
NfadeL 21, L21
NfadeL 22, L22
NfadeL 23, L23
NfadeL 24, L24
NfadeLm 25, LB1
NFadeL 25, LB2


NfadeL 26, L26
NfadeL 27, L27
NfadeL 28, L28


NfadeL 29, L29
NfadeL 30, L30
NfadeL 31, L31
NfadeL 32, L32

NfadeL 33, L33
NfadeL 34, L34
NfadeL 35, L35
NfadeL 36, L36
NfadeL 37, L37

NfadeL 38, L38
NfadeL 39, L39
NfadeL 40, L40
NfadeL 41, L41
NfadeL 42, L42
NfadeL 43, L43

NfadeLm 44, L44a
NfadeL 44, L44

NfadeL 45, L45
NfadeL 46, L46
NfadeL 49, L49
NfadeL 50, L50

'AUX Solenoid Drivers

NFadeL 12, L12
NFadeL 13, L13
NFadeL 14, L14
NfadeL 15, L15

End Sub

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
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr)Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

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
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr)Then
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
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr)Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr)-1
                If FlashRepeat(nr)Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr)Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr)Then FadingLevel(nr) = 4
    End Select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

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

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Dim RampDir, RampStep, NewL12, OldL12, NewL14, OldL14, NewL15, OldL15
RampStep = 0

Sub UpdateMultipleLamps


    NewL15 = Controller.Lamp(15)
    If NewL15 <> OldL15 Then
    If NewL15 Then
    bsSaucer2.ExitSol_On
    PlaySoundAt "popper", sw31
    End If
    End If
    OldL15 = NewL15


    NewL14 = Controller.Lamp(14)
    If NewL14 <> OldL14 Then
    If NewL14 Then
    bsSaucer3.ExitSol_On
    PlaySoundAt "popper", sw02
    End If
    End If
    OldL14 = NewL14

    'Trough Eject  - Controlled by Lamp 12 - Aux Driver Board.
    NewL12 = Controller.Lamp(12)
    If NewL12 <> OldL12 Then
     If NewL12 Then
     bsTrough.ExitSol_On
     PlaySoundat "ballrelease", DrainKicker
     End If
     End If
    OldL12 = NewL12

 'Star Gate - Controlled by Lamp 13 - Aux Driver Board.

 ' See Diag Timer

End Sub

Sub DelayTimer_Timer
Me.Enabled=0
End Sub

'Bumpers/Slingshots
Sub Bumper1_Hit
VpmTimer.PulseSw 42
PlaySound "jet1":DOF 100,2
PlaySoundat SoundFXDOF("Jet1",100,DOFPulse,DOFContactors), Bumper1
End Sub

Dim RStep, Lstep, Tstep, Blstep, TLstep, BRstep, TRstep, MStep, MLStep, MRStep, PlasX
Dim dtest
Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 43
    For each plasx in PlasFlash
  plasx.opacity = plasx.opacity *1.1
  Next
    LSFA.state = 1
    LSFB.state = 1
    LF.visible = 1
    PlaySoundat SoundFXDOF("sling",101,DOFPulse,DOFContactors), SLING5
    LSling.Visible = 0
    LSling1.Visible = 1
    sling3.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval  = 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling3.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling3.TransZ = 0:LSFA.state = 0:LSFB.state = 0
        Case 5:
    For each plasx in PlasFlash
  plasx.opacity = plasx.opacity /1.1
  Next
    LF.visible = 0
    LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 43
    For each plasx in PlasFlash
  plasx.opacity = plasx.opacity *1.1
  Next
    RSFA.state = 1
    RSFB.state = 1
    RF.visible = 1
    PlaySoundat SoundFXDOF("sling",102,DOFPulse,DOFContactors), SLING4
    RSling.Visible = 0
    RSling1.Visible = 1
    SLING4.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval  = 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling4.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling4.TransZ = 0:RSFA.state = 0:RSFB.state = 0
        Case 5:
    For each plasx in PlasFlash
  plasx.opacity = plasx.opacity /1.1
  Next
    RF.visible = 0
    RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub CentreSling_Slingshot
    vpmTimer.PulseSw 43
    PlaySoundat SoundFXDOF("Small_Sling",170,DOFPulse,DOFContactors), SLING1
    MSling.Visible = 0
    MSling1.Visible = 1
    SLING1.TransZ = -20
    MStep = 0
    CentreSling.TimerEnabled = 1
  CentreSling.TimerInterval  = 10
End Sub

Sub CentreSling_Timer
    Select Case MStep
        Case 3:MSLing1.Visible = 0:MSLing2.Visible = 1:SLING1.TransZ = -10
        Case 4:MSLing2.Visible = 0:MSLing.Visible = 1:sling1.TransZ = 0:me.TimerEnabled = 0
    End Select
    MStep = MStep + 1
End Sub

Sub TopSling_Hit
    vpmTimer.PulseSw 26
    TSling.Visible = 0
    TSling1.Visible = 1
    SLING2.TransZ = -20
    TStep = 0
    TopSling.TimerEnabled = 1
  TopSling.TimerInterval  = 10
End Sub

Sub TopSling_Timer
    Select Case TStep
        Case 3:TSLing1.Visible = 0:TSLing2.Visible = 1:SLING2.TransZ = -10
        Case 4:TSLing2.Visible = 0:TSLing.Visible = 1:SLING2.TransZ = 0:me.TimerEnabled = 0
    End Select
    TStep = TStep + 1
End Sub

Sub MLSlingShot_SlingShot
    vpmTimer.PulseSw 43
    PlaySoundat SoundFXDOF("small_sling",113,DOFPulse,DOFContactors), SLING5
    MslingL.Visible = 0
    MSlingL1.Visible = 1
    SLING5.TransZ = -20
    MLStep = 0
    Me.TimerEnabled = 1
  Me.TimerInterval  = 10
End Sub

Sub MLSlingShot_Timer
    Select Case MLStep
        Case 3:MSLingL1.Visible = 0:MSLingL2.Visible = 1:SLING5.TransZ = -10
        Case 4:MSLingL2.Visible = 0:MslingL.Visible = 1:SLING5.TransZ = 0:me.TimerEnabled = 0
    End Select
    MLStep = MLStep + 1
End Sub

Sub MRSlingShot_SlingShot
    vpmTimer.PulseSw 43
    PlaySoundat SoundFXDOF("small_sling",114,DOFPulse,DOFContactors), SLING6
    MSlingR.Visible = 0
    MSlingR1.Visible = 1
    SLING6.TransZ = -20
    MRStep = 0
    Me.TimerEnabled = 1
  Me.TimerInterval  = 10
End Sub

Sub MRSlingShot_Timer
    Select Case MRStep
        Case 3:MSLingR1.Visible = 0:MSLingR2.Visible = 1:SLING6.TransZ = -10
        Case 4:MSLingR2.Visible = 0:MslingR.Visible = 1:SLING6.TransZ = 0:me.TimerEnabled = 0
    End Select
    MRStep = MRStep + 1
End Sub

'drop targets
Sub sw01_Hit:PlaySoundAt SoundFX("target_drop",DOFDropTargets),sw01:dtLeft.Hit 1:End Sub
Sub sw11_Hit:PlaySoundAt SoundFX("target_drop",DOFDropTargets),sw11:dtLeft.Hit 2:End Sub
Sub sw21_Hit:PlaySoundAt SoundFX("target_drop",DOFDropTargets),sw21:dtLeft.Hit 3:End Sub
Sub sw05_Hit:PlaySoundAt SoundFX("target_drop",DOFDropTargets),sw05:dtRight.Hit 1:End Sub
Sub sw15_Hit:PlaySoundAt SoundFX("target_drop",DOFDropTargets),sw15:dtRight.Hit 2:End Sub
Sub sw25_Hit:PlaySoundAt SoundFX("target_drop",DOFDropTargets),sw25:dtRight.Hit 3:End Sub

'spot targets
Sub sw32_Hit:VpmTimer.PulseSw 32:PlaySoundat "target", sw32:End Sub ' Mid upper spot target
Sub sw33_Hit:VpmTimer.PulseSw 33:PlaySoundat "target", sw33:End Sub ' Mid rt spot target
Sub sw60_Hit:VpmTimer.PulseSw 60:PlaySoundat "target", sw60:End Sub ' Lower left spot target
Sub sw61_Hit:VpmTimer.PulseSw 61:PlaySoundat "target", sw61:End Sub ' Lower right spot target
Sub sw06_Hit:VpmTimer.PulseSw 6:PlaySoundat "target", sw06:End Sub ' upper right spot target
Sub sw3_Hit:VpmTimer.PulseSw 3:PlaySoundat "target", sw3:End Sub ' Mid  spot target'modified by scapino
Sub sw13_Hit:VpmTimer.PulseSw 13:PlaySoundat "target", sw13:End Sub ' Mid  spot target'modified by scapino
Sub sw23_Hit:VpmTimer.PulseSw 23:PlaySoundat "target", sw23:End Sub ' Mid  spot target'modified by scapino

'rollovers/rollunders
Sub sw00_Hit:Controller.Switch(00)=1:End Sub
Sub sw00_UnHit:Controller.Switch(00)=0:End Sub
Sub sw10_Hit:Controller.Switch(10)=1:End Sub
Sub sw10_UnHit:Controller.Switch(10)=0:End Sub
Sub sw16_Hit:Controller.Switch(16)=1:End Sub
Sub sw16_UnHit:Controller.Switch(16)=0:End Sub
Sub sw20_Hit:Controller.Switch(20)=1:End Sub
Sub sw20_unHit:Controller.Switch(20)=0:End Sub
Sub sw22_Hit:Controller.Switch(22)=1:End Sub
Sub sw22_unHit:Controller.Switch(22)=0:End Sub
Sub sw62_Hit:Controller.Switch(62)=1:End Sub
Sub sw62_unHit:Controller.Switch(62)=0:End Sub
Sub sw41_Hit:Controller.Switch(41)=1:End Sub
Sub sw41_unHit:Controller.Switch(41)=0:End Sub
Sub sw40_Hit:Controller.Switch(40)=1:End Sub
Sub sw40_unHit:Controller.Switch(40)=0:End Sub
Sub sw50_Hit:Controller.Switch(50)=1:End Sub
Sub sw50_unHit:Controller.Switch(50)=0:End Sub
Sub sw51_Hit:Controller.Switch(51)=1:End Sub
Sub sw51_unHit:Controller.Switch(51)=0:End Sub
Sub sw62_Hit:Controller.Switch(62)=1:End Sub
Sub sw62_unHit:Controller.Switch(62)=0:End Sub
Sub sw62_Hit:Controller.Switch(62)=1:End Sub
Sub sw62_unHit:Controller.Switch(62)=0:End Sub
Sub sw63_Hit:Controller.Switch(63)=1:End Sub
Sub sw63_unHit:Controller.Switch(63)=0:End Sub

Sub ARollovers_Hit(idx)
PlaySoundAtBall "rollover"
End Sub

Sub AGates_Hit(idx)
PlaySoundatBall "fx_gate"
End Sub

Dim msg(38)
Dim char(1024)
char(0) = " "
char(63) = "0"
char(6) = "1"
char(91) = "2"
char(79) = "3"
char(102) = "4"
char(109) = "5"
char(125) = "6"
char(7) = "7"
char(127) = "8"
char(111) = "9"
char(191) = "À" ' 0,
char(134) = "Á" ' 1,
char(219) = "Â" ' 2,
char(207) = "Ã" ' 3,
char(230) = "Ä" ' 4,
char(237) = "Å" ' 5,
char(253) = "Æ" ' 6,
char(135) = "Ç" ' 7,
char(255) = "È" ' 8,
char(239) = "É" ' 9,
char(768) = "1" ' 1  Gottlieb
char(896) = "Á" ' 1,  Gottlieb
'1 on the center line (for Gottlieb) - $CA (202)
'1 on the center with comma - $CB (203)
'6 without top line (Atari, Gottlieb) - $CC (204)
'6 without top line with comma - $CD (205)
'9 without bottom line (Atari, Gottlieb) - $CE (206)
'9 without bottom line with comma - $CF (207)
char(768) = chr(&HCA) ' 1 on Gottlieb tables
char(896) = chr(&HCB) ' 1, on Gottlieb tables
char(124) = chr(&HCC) ' 6 without top line
char(252) = chr(&HCD) ' 6, without top line
char(103) = chr(&HCE) ' 9 without bottom line
char(231) = chr(&HCF) ' 9, without bottom line

Dim Digits(38)
LXM=0

Digits(32)=Array(D241,D242,D243,D244,D245,D246,D247,LXM,D248)
Digits(33)=Array(D251,D252,D253,D254,D255,D256,D257,LXM,D258)
Digits(34)=Array(D261,D262,D263,D264,D265,D266,D267,LXM,D268)
Digits(35)=Array(D271,D272,D273,D274,D275,D276,D277,LXM,D278)
Digits(36)=Array(D281,D282,D283,D284,D285,D286,D287,LXM,D288)
Digits(37)=Array(D291,D292,D293,D294,D295,D296,D297,LXM,D298)

Sub DisplayTimer_Timer
    Dim ChgLED,ii,num,chg,stat,obj
    ChgLED = Controller.ChangedLEDs(&H3f, &Hffffffff)
    'SetFlash 170,1
    If Not IsEmpty(ChgLED) Then
         For ii = 0 To UBound(chgLED)
            Num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
             If Num>31 Then
                For Each obj In Digits(Num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg\2 : stat = stat\2
                    Flasher1.visible = 1
                    For each plasx in PlasFlash
          plasx.visible = 1
          'PlasX.color = RGB(0,0,192)
          Next
                Next
             Else
          'Flasher1.visible = 0
             End If
    Next
End If
End Sub

Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Spirit - DIP switches"
    .AddFrame 2,0,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
    .AddFrame 2,76,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,122,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 2,168,190,"Novelty",&H08000000,Array("normal game mode",0,"50K for special/extra ball",&H08000000)'dip 28
    .AddFrameExtra 2,214,190,"Attract Sound",&H000C,Array("off",0,"every 10 seconds",&H0004,"every 2 minutes",&H0008,"every 4 minutes",&H000C)'SS-board dip 3&4
    .AddFrame 205,0,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrame 205,76,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,122,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
    .AddFrame 205,168,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddChk 205,221,120,Array("Match feature",&H02000000)'dip 26
    .AddChkExtra 205,238,120,Array("Background sound",&H0010)'SS-board dip 5
    .AddLabel 50,300,300,20,"After hitting OK, press F3 to reset game with new settings."
  End With
  Dim extra
  extra = Controller.Dip(4) + Controller.Dip(5)*256
  extra = vpmDips.ViewDipsExtra(extra)
  Controller.Dip(4) = extra And 255
  Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")


Dim rampup
Dim rampdown


Sub diag_timer()
  If L13.State=1 Then
    RampDir         = 1
    LiftRampTimer.Interval  = 20
    LiftRampTimer.Enabled   = True
    PlaySoundat "gunpopper", Screw39
        PlaySoundat "stargate", Screw39
        me.enabled = 0
        diag2.enabled = 1
        MB.state = 1
End If
End Sub


Sub diag2_timer()
If L13.State = 0 Then
        RampDir         = 0
    LiftRampTimer.Interval  = 20
    LiftRampTimer.Enabled   = True
    PlaySoundat "gunpopper", Screw39
        PlaySoundat "stargate", Screw39
        me.enabled = 0
        diag.enabled = 1
        MB.State = 0
End If
End Sub


Sub LiftRampTimer_Timer()
  Select Case RampDir
    Case 0
      If RampStep <= 0 Then
        LiftRamp.Collidable = True
                Ramp1.Collidable = True
                rampdown = 1
                rampup = 0
        Me.Enabled      = False
      Else
        RampStep = RampStep - 1
      End If
    Case 1
      If RampStep >= 10 Then
        LiftRamp.Collidable = False
                Ramp1.Collidable = False
                rampup = 1
                'rampdown = 0
        Me.Enabled      = False
      Else
        RampStep = RampStep + 1
      End If
  End Select
  LiftRamp.HeightBottom   = RampStep * 10
End Sub

Sub Update_Stuff_Timer
    pminil.Roty = Flipper1.Currentangle
  pLeftFlipper.Roty = LeftFlipper.Currentangle
  PLeftFlipper1.Roty = Flipper3.Currentangle
  pRightFlipper.Roty = RightFlipper.Currentangle
  pRightFlipper1.Roty = Flipper4.Currentangle
  PminiR.Roty = Flipper2.Currentangle
  BallShadowUpdate
    RollingSoundUpdate
    If ShowDT = True Then
    MiddleWheel.RotZ = MiddleWheel.Rotz + 0.3
    End If
    If Controller.Solenoid(10) = True Then
    Light1.State = 1
    Else
    Light1.State = 0
    End If

If Controller.Solenoid(10) = False AND L1.State = 1 Then
Flipper1.RotateToStart
DOF 140,0
'PlaySoundat SoundFXDOF("FlipperDown",140,DOFOff,DOFFlippers), Flipper1
End If

If Controller.Solenoid(10) = False AND L1.State = 1 Then
Flipper2.RotateToStart
'PlaySoundat SoundFXDOF("FlipperDown",140,DOFOff,DOFFlippers), Flipper1
DOF 141,0
End If

FSL.RotZ = LeftFlipper.currentangle
FSLM.RotZ = Flipper1.currentangle
FSC.RotZ = Flipper3.currentangle

FSR.RotZ = RightFlipper.currentangle
FSRM.RotZ = Flipper2.currentangle
FST.RotZ = Flipper4.currentangle

PLGott.RotY = Flipper1.currentangle - 90
PRGott.RotY = Flipper2.currentangle - 90
PLGott1.RotY = Flipper3.currentangle - 90

If L1.State = 1 OR Controller.Solenoid(10) = 0 Then
Bumper1.force = 0
LeftSlingShot.SlingshotThreshold = 100
RightSlingShot.SlingshotThreshold = 100
MLSlingShot.SlingshotThreshold = 100
MRSlingShot.SlingshotThreshold = 100
CentreSling.SlingshotThreshold = 100
else
Bumper1.force = 10
LeftSlingShot.SlingshotThreshold = 1
RightSlingShot.SlingshotThreshold = 1
MLSlingShot.SlingshotThreshold = 1
MRSlingShot.SlingshotThreshold = 1
CentreSling.SlingshotThreshold = 1
End If

End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************
Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "ZZZ" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / ZZZ.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "ZZZ" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / ZZZ.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "ZZZ" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / ZZZ.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

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

Sub PlaySoundAtVol(soundname, tableobj, vol)
    PlaySound soundname, 1, (vol), AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

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
        StopSound("fx_ballrolling" & b)
    Next

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*********Sound Effects**************************************************************************************************
                                       'Use these for your sound effects like ball rolling, etc.

Sub aRubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub Flipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub Flipper2_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub Flipper3_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub Flipper4_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub RandomSoundHole()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySound "fx_Hole1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_Hole2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_Hole3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 4 : PlaySound "fx_Hole4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub Metals_Hit(idx)
RandomSoundMetal
End Sub

Sub RandomSoundMetal()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_metal_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_metal_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_metal_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub Woods_Hit(idx)
RandomSoundWood
End Sub

Sub RandomSoundWood()
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySound "woodhit", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "woodhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "woodhit3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

'*****************************************
' Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)


Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls

  ' render the shadow for each ball
    For b = 0 to UBound(BOT)
    If BOT(b).X < ZZZ.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (ZZZ.Width/2))/7)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (ZZZ.Width/2))/7)) - 10
    End If

      BallShadow(b).Y = BOT(b).Y + 10
      BallShadow(b).Z = 1
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

Sub ZZZ_Exit
Controller.Stop
End Sub


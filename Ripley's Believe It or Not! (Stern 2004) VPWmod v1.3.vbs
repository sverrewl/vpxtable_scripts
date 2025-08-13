'Ripley's Believe It or Not! (Stern 2004)
'https://www.ipdb.org/machine.cgi?id=4917
'
' VPW mod started 12/06/2020 (!)
'
' mcarter78 - Project lead
' iaakki - Lots of stuff
' Benji - Early Developement
' Sixtoe - All sorts of tweaking and fixing
' Apophis - Code & Tech support
' RothbauerW - Varitarget support
' Bord - Playfield mesh & physical scoop.
' Hauntfreaks - Playfield image adjustments, cabinet and backglass images.
' TastyWasps - VR Room and tweaks
' PT5K - Staged flipper support
' Ebislit - Early Insert work
' Kingdids - Stern Bumpercaps
' Kevv - Stripping and scanning his RBION for plastics
' Wylte - Getting irl references
' Passion4Pins - Additional graphics help
'
' Mod started from vpx 1.0 by JPSalas August 2017
'
' All options are available from in game menu accessed pressing F12
'
' Version Log at end of script.

Option Explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs In order To run this table, available In the vp10 package"
On Error GoTo 0

'*******************************************
'  ZCON: Constants and Global Variables
'*******************************************

Const TableVersion = "1.1"  'Table version (shown in option UI)


Const BallSize = 50
Const BallMass = 1
Const tnob = 4
Const lob = 0 'locked balls on start

' Language Roms
Const cGameName = "ripleys" 'English
' Const cGameName = "ripleysf" 'French
' Const cGameName = "ripleysg" 'German
' Const cGameName = "ripleysi" 'Italian
' Const cGameName = "ripleysl" 'Spanish

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim BIP             'Balls in play
BIP = 0
Dim BIPL            'Ball in plunger lane
BIPL = False

' User Options for In-Game Menu (Magna Keys)
Dim LightLevel, VolumeDial, BallRollVolume, RampRollVolume, DynamicBallShadowsOn, AmbientBallShadowOn, StagedFlippers

Dim VarHidden, UseVPMColoredDMD
Const UseVPMModSol = 2

If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
    lrail.Visible = 0
    rrail.Visible = 0
  lockbar.Visible = 0
End If

'*************************************************************
' VR Room Auto-Detect
'*************************************************************
Dim VR_Obj, UseVPMDMD

Dim DesktopMode: DesktopMode = Table1.ShowDT

If RenderingMode = 2 Then
  Pincab_Backglass.BlendDisableLighting = 0.25
  Pincab_Cabinet.BlendDisableLighting = 0.10
  Pincab_Backbox.BlendDisableLighting = 0.10
    lrail.Visible = 0
    lrail1.Visible = 0
    rrail.Visible = 0
    rrail1.Visible = 0
  lockbar.Visible = 0
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRRoom : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRLighting : VR_Obj.BlendDisableLighting = 0.15 : Next
  UseVPMDMD = True
Else
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRRoom : VR_Obj.Visible = 0 : Next
  'UseVPMDMD = DesktopMode
End If

LoadVPM "03060000", "SEGA.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SCoin = ""

Dim bsTrough, mIdolMag, mShrunkenMag, mUp, mDown, bsLock, bsSkill, bsVUK, plungerIM, x
Dim RBall1, RBall2, RBall3, RBall4, gBOT

'************
' Table init.
'************

Sub table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        .Games(cGameName).Settings.Value("sound") = 1 '0 = disable  rom sound
        .SplashInfoLine = "Ripleys Believe It Or Not - Stern 2004" & vbNewLine & "VPW Mod"
        .Games(cGameName).Settings.Value("rol") = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .Hidden = VarHidden
        .Switch(42) = 1
        .Switch(43) = 1
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    ' Controller.SolMask(0) = 0
        ' vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        On Error Goto 0
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, Bumper5, Bumper6, LeftSlingshot, RightSlingshot)

    Set RBall1 = sw11.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set RBall2 = sw12.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set RBall3 = sw13.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set RBall4 = sw14.CreateSizedballWithMass(Ballsize/2,Ballmass)

  gBOT = Array(RBall1, RBall2, RBall3, RBall4)

  Controller.Switch(11) = 1
  Controller.Switch(12) = 1
  Controller.Switch(13) = 1
  Controller.Switch(14) = 1

    ' Add balls to shadow dictionary
    Dim xx
  For Each xx in gBOT
    bsDict.Add xx.ID, bsNone
  Next

    Set mIdolMag = New cvpmMagnet
    With mIdolMag
        .InitMagnet IMagnet, 50
        .Solenoid = 19
        .GrabCenter = 0
    End With

    Set mShrunkenMag = New cvpmMagnet
    mShrunkenMag.InitMagnet SMagnet, 40
    mShrunkenMag.GrabCenter = false

    ' Impulse Plunger
    Const IMPowerSetting = 44 ' Plunger Power
    Const IMTime = 0.7        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("AutoPlunger",DOFContactors), SoundFX("AutoPlunger",DOFContactors)
        .CreateEvents "plungerIM"
    End With

  ' vpmMapLights InsertLamps      ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub


'****
'Keys
'****
Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X + 10
  End If

  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X - 10
      If StagedFlippers = 0 Then
        FlipperActivate RightFlipper1, URFPress
    End If
  End If

  If keycode = KeyUpperRight Then
    FlipperActivate RightFlipper1, URFPress
  End If

  If keycode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 1
    SoundNudgeLeft
  End If
  If keycode = RightTiltKey Then
    Nudge 270, 1
    SoundNudgeRight
  End If
  If keycode = CenterTiltKey Then
    Nudge 0, 1
    SoundNudgeCenter
  End If

  If keycode = MechanicalTilt Then
    SoundNudgeCenter() 'Send the Tilting command to the ROM (usually by pulsing a Switch), or run the tilting code for an orginal table
  End If

  If keycode = StartGameKey Then
    SoundStartButton
  End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
    Select Case Int(Rnd * 3)
      Case 0
        PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1
        PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2
        PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X - 10
  End If

  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X + 10
      If StagedFlippers = 0 Then
      FlipperDeActivate RightFlipper1, URFPress
    End If
  End If

  If keycode = KeyUpperRight Then
    FlipperDeActivate RightFlipper1, URFPress
  End If

    If keycode = PlungerKey Then
    Plunger.Fire
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    PinCab_Plunger.Y = -45.68289
    If BIPL Then              'If true then ball in shooter lane, else no ball is shooter lane
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If

  If KeyUpHandler(KeyCode)Then Exit Sub
End Sub


'************************
'Realtime Updates
'************************

Sub RDampen_Timer() '10ms
  Cor.Update
  UpdateLeds
  RollingUpdate
    DoSTAnim
  DoVTAnim  'handle stand up target animations
End Sub

' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
    If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
    FlipperVisualUpdate
  UpdateBallBrightness
End Sub

'*********
'Solenoids
'*********

SolCallBack(1) = "ReleaseBall"
SolCallBack(2) = "Auto_Plunger"
SolCallBack(3) = "SolVUK"
SolCallBack(4) = "vpmSolDiverter TempleDiv,1,"
SolCallBack(5) = "vpmSolDiverter LockDiverter,1,"
' 6, 7, 8, 9, 10 11 pop bumpers
SolCallBack(12) = "SolSkillScoop"
SolCallback(13) = "SolLock"
SolCallback(20) = "SolUpperMagnet"
SolCallBack(21) = "SolVReset"
SolCallBack(23) = "SolPost"
'SolCallBack(24) = "SolKnocker"

SolModCallBack(22) = "Lampz.SetLamp 102,"   'Idol Opto LED (Playfield)
SolModCallBack(25) = "Flash125"       'Flasher Dome Lower Left, Lower Pops (Playfield)
SolModCallBack(26) = "Lampz.SetLamp 106,"   'Flasher Left Spinner (Playfield)
SolModCallBack(27) = "Flash127"       'Flasher Dome Upper Left (Under Temple)
SolModCallBack(28) = "Flash128"       'Flasher Shrunken Head
SolModCallBack(29) = "Flash129"       'Flasher Dome Varitarget
SolModCallBack(30) = "Flash130"       'Flasher Dome Upper Right, Upper Pops (Playfield)
SolModCallBack(31) = "Lampz.SetLamp 111,"   'Flasher Right Spinner (Playfield)
SolModCallBack(32) = "Flash132"       'Flasher Dome Right Ramp



'******************************************************
'           TROUGH
'******************************************************

Sub sw11_Hit():Controller.Switch(11) = 1:UpdateTrough:End Sub
Sub sw11_UnHit():Controller.Switch(11) = 0:UpdateTrough:End Sub
Sub sw12_Hit():Controller.Switch(12) = 1:UpdateTrough:End Sub
Sub sw12_UnHit():Controller.Switch(12) = 0:UpdateTrough:End Sub
Sub sw13_Hit():Controller.Switch(13) = 1:UpdateTrough:End Sub
Sub sw13_UnHit():Controller.Switch(13) = 0:UpdateTrough:End Sub
Sub sw14_Hit():Controller.Switch(14) = 1:UpdateTrough:End Sub
Sub sw14_UnHit():Controller.Switch(14) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 100
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw14.BallCntOver = 0 Then sw13.kick 60, 9
  If sw13.BallCntOver = 0 Then sw12.kick 60, 9
  If sw12.BallCntOver = 0 Then sw11.kick 60, 9
  UpdateTroughTimer.Enabled = 0
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub Drain_Hit() 'Drain
  UpdateTrough
    RandomSoundDrain(Drain)
  vpmTimer.AddTimer 500, "Drain.kick 60, 9'"
End Sub

Sub ReleaseBall(enabled)
  If enabled Then
    sw14.kick 60, 12
        vpmTimer.PulseSw 15
    UpdateTrough
        RandomSoundBallRelease sw14
  End If
End Sub

Sub SolKnocker(Enabled)
  If Enabled Then
    KnockerSolenoid
  End If
End Sub

'*******************************************
' ZFLP: Flippers
'*******************************************
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"


Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
    LF.Fire
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
    RF.Fire
        If rightflipper.currentangle < rightflipper.endangle + ReflipAngle Then
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

Sub SolURFlipper(Enabled)
    If Enabled Then
        RightFlipper1.RotateToEnd
        If rightflipper1.currentangle < rightflipper1.endangle + ReflipAngle Then
            RandomSoundReflipUpRight RightFlipper1
        Else
            SoundFlipperUpAttackRight RightFlipper1
      RandomSoundFlipperUpRight RightFlipper1
        End If
    Else
        RightFlipper1.RotateToStart
        If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
    End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub FlipperVisualUpdate 'This subroutine updates the flipper shadows and visual primitives
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
  FlipperRSh001.RotZ = RightFlipper1.CurrentAngle
  LFLogo.RotZ = LeftFlipper.CurrentAngle
  RFLogo.RotZ = RightFlipper.CurrentAngle
    URFLogo.RotZ = RightFlipper1.CurrentAngle
End Sub

Sub RightFlipper1_Collide(parm)
  RandomSoundRubberFlipper(parm)
End Sub

'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, -5.5
    x.AddPt "Polarity", 2, 0.4, -5.5
    x.AddPt "Polarity", 3, 0.6, -5.0
    x.AddPt "Polarity", 4, 0.65, -4.5
    x.AddPt "Polarity", 5, 0.7, -4.0
    x.AddPt "Polarity", 6, 0.75, -3.5
    x.AddPt "Polarity", 7, 0.8, -3.0
    x.AddPt "Polarity", 8, 0.85, -2.5
    x.AddPt "Polarity", 9, 0.9,-2.0
    x.AddPt "Polarity", 10, 0.95, -1.5
    x.AddPt "Polarity", 11, 1, -1.0
    x.AddPt "Polarity", 12, 1.05, -0.5
    x.AddPt "Polarity", 13, 1.1, 0
    x.AddPt "Polarity", 14, 1.3, 0

    x.AddPt "Velocity", 0, 0,    1
    x.AddPt "Velocity", 1, 0.160, 1.06
    x.AddPt "Velocity", 2, 0.410, 1.05
    x.AddPt "Velocity", 3, 0.530, 1'0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03,  0.945
  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
  LF.SetObjects "LF", LeftFlipper, TriggerLF
  RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'******************************************************
'     FLIPPER CORRECTION FUNCTIONS
'******************************************************


Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
  Private Balls(20), balldata(20)
  Private Name

  Dim PolarityIn, PolarityOut
  Dim VelocityIn, VelocityOut
  Dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    ReDim PolarityIn(0)
    ReDim PolarityOut(0)
    ReDim VelocityIn(0)
    ReDim VelocityOut(0)
    ReDim YcoefIn(0)
    ReDim YcoefOut(0)
    Enabled = True
    TimeDelay = 50
    LR = 1
    Dim x
    For x = 0 To UBound(balls)
      balls(x) = Empty
      Set Balldata(x) = new SpoofBall
    Next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
    If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
    If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
    If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper
    FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    Dim str
    str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  ' Legacy: just no op
  Public Property Let EndPoint(aInput)

  End Property

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      Case "Polarity"
        ShuffleArrays PolarityIn, PolarityOut, 1
        PolarityIn(aIDX) = aX
        PolarityOut(aIDX) = aY
        ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity"
        ShuffleArrays VelocityIn, VelocityOut, 1
        VelocityIn(aIDX) = aX
        VelocityOut(aIDX) = aY
        ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef"
        ShuffleArrays YcoefIn, YcoefOut, 1
        YcoefIn(aIDX) = aX
        YcoefOut(aIDX) = aY
        ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
  End Sub

  Public Sub AddBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If IsEmpty(balls(x)) Then
        Set balls(x) = aBall
        Exit Sub
      End If
    Next
  End Sub

  Private Sub RemoveBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If TypeName(balls(x) ) = "IBall" Then
        If aBall.ID = Balls(x).ID Then
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
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x) ) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x) ) Then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef
      Ycoef = 1

      'y safety Exit
      If aBall.VelY > -8 Then 'ball going down
        RemoveBall aBall
        Exit Sub
      End If

      'Find balldata. BallPos = % on Flipper
      For x = 0 To UBound(Balls)
        If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
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
  Dim x, aCount
  aCount = 0
  ReDim a(UBound(aArray) )
  For x = 0 To UBound(aArray)   'Shuffle objects in a temp array
    If Not IsEmpty(aArray(x) ) Then
      If IsObject(aArray(x)) Then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  If offset < 0 Then offset = 0
  ReDim aArray(aCount-1+offset)   'Resize original array
  For x = 0 To aCount-1       'set objects back into original array
    If IsObject(a(x)) Then
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
  BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
  Dim x, y, b, m
  x = input
  m = (Y2 - Y1) / (X2 - X1)
  b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x
      y = .y
      z = .z
      velx = .velx
      vely = .vely
      velz = .velz
      id = .ID
      mass = .mass
      radius = .radius
    End With
  End Property
  Public Sub Reset()
    x = Empty
    y = Empty
    z = Empty
    velx = Empty
    vely = Empty
    velz = Empty
    id = Empty
    mass = Empty
    radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  Dim y 'Y output
  Dim L 'Line
  'find active line
  Dim ii
  For ii = 1 To UBound(xKeyFrame)
    If xInput <= xKeyFrame(ii) Then
      L = ii
      Exit For
    End If
  Next
  If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )     'Clamp lower
  If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )    'Clamp upper

  LinearEnvelope = Y
End Function

'******************************************************
'     FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
          gBOT(b).velx = gBOT(b).velx / 1.3
          gBOT(b).vely = gBOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
  if velocity < 0.7 then exit sub   'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
    coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub

'*****************
' Maths
'*****************

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
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

'*************************************************
' Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, URFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 1.5  '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0
    SOSRampup = 2.5
  Case 1
    SOSRampup = 6
  Case 2
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025  'mid 90's and later

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
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b

    For b = 0 To UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3 * Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
    If FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If
  End If
End Sub

Const LiveDistanceMin = 5  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
Const BaseDampen = 0.55

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir, LiveDist
    Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce   'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime
    CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)

    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then   'Perfect catch only when catch time happens in the beginning of the window
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)  'Partial catch when catch happens a bit late
        End If

        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0

        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        Elseif LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
    End If
End Sub

'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen ActiveBall
  TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD        'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1    'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64    'there's clamping so interpolate up to 56 at least

Dim SleevesD  'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False    'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn   'tbpOut.text
  Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
  End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then
      If BallSpeed(aBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
    "actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
    If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)

    aBall.velx = aBall.velx * coef
    aBall.vely = aBall.vely * coef
    If debugOn Then TBPout.text = str
  End Sub

  Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
      aBall.velx = aBall.velx * coef
      aBall.vely = aBall.vely * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    Dim x
    For x = 0 To UBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
    Next
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub
End Class

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7  'Level of bounces. Recommmended value of 0.7

' Sub TargetBouncer(aBall,defvalue)
'   Dim zMultiplier, vel, vratio
'   If TargetBouncerEnabled = 1 And aball.z < 30 Then
'     '   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
'     vel = BallSpeed(aBall)
'     If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
'     Select Case Int(Rnd * 6) + 1
'       Case 1
'         zMultiplier = 0.2 * defvalue
'       Case 2
'         zMultiplier = 0.25 * defvalue
'       Case 3
'         zMultiplier = 0.3 * defvalue
'       Case 4
'         zMultiplier = 0.4 * defvalue
'       Case 5
'         zMultiplier = 0.45 * defvalue
'       Case 6
'         zMultiplier = 0.5 * defvalue
'     End Select
'     aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
'     aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
'     aBall.vely = aBall.velx * vratio
'     '   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
'     '   debug.print "conservation check: " & BallSpeed(aBall)/vel
'   End If
' End Sub

' BM TargetBouncer
sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel
' vel = BallSpeed(aBall)
' debug.print "bounce"
  if aball.z < 30 then
'   debug.print "velz: " & activeball.velz
    Select Case Int(Rnd * 4) + 1
      Case 1: zMultiplier = defvalue+1.1
      Case 2: zMultiplier = defvalue+1.05
      Case 3: zMultiplier = defvalue+0.7
      Case 4: zMultiplier = defvalue+0.3
    End Select
    aBall.velz = aBall.velz * zMultiplier
'   debug.print "----> velz: " & activeball.velz
    'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1.4
End Sub

'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 4
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

Class SlingshotCorrection
  Public DebugOn, Enabled
  Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut

  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
    Enabled = True
  End Sub

  Public Property Let Object(aInput)
    Set Slingshot = aInput
  End Property

  Public Property Let EndPoint1(aInput)
    SlingX1 = aInput.x
    SlingY1 = aInput.y
  End Property

  Public Property Let EndPoint2(aInput)
    SlingX2 = aInput.x
    SlingY2 = aInput.y
  End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    Dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1
      YL = SlingY1
      XR = SlingX2
      YR = SlingY2
    Else
      XL = SlingX2
      YL = SlingY2
      XR = SlingX1
      YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If Abs(XR - XL) > Abs(YR - YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If Not IsEmpty(ModIn(0) ) Then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      '   debug.print " BallPos=" & BallPos &" Angle=" & Angle
      '   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled Then aBall.Velx = RotVxVy(0)
      If Enabled Then aBall.Vely = RotVxVy(1)
      '   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      '   debug.print " "
    End If
  End Sub
End Class


'******************************************************
'   TRACK ALL BALL VELOCITIES
'     FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
  Public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize
    ReDim ballvel(0)
    ReDim ballvelx(0)
    ReDim ballvely(0)
  End Sub

  Public Sub Update() 'tracks in-ball-velocity
    Dim str, b, AllBalls, highestID
    allBalls = GetBalls

    For Each b In allballs
      If b.id >= HighestID Then highestID = b.id
    Next

    If UBound(ballvel) < highestID Then ReDim ballvel(highestID)  'set bounds
    If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)  'set bounds
    If UBound(ballvely) < highestID Then ReDim ballvely(highestID)  'set bounds

    For Each b In allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

Function RotPoint(x,y,angle)
  Dim rx, ry
  rx = x * dCos(angle) - y * dSin(angle)
  ry = x * dSin(angle) + y * dCos(angle)
  RotPoint = Array(rx,ry)
End Function


'*****************************************************************************************************
'*******************************************************************************************************
'END nFOZZY Physics'

'******************************************************
' ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************

Class StandupTarget
    Private m_primary, m_prim, m_sw, m_animate

    Public Property Get Primary(): Set Primary = m_primary: End Property
    Public Property Let Primary(input): Set m_primary = input: End Property

    Public Property Get Prim(): Set Prim = m_prim: End Property
    Public Property Let Prim(input): Set m_prim = input: End Property

    Public Property Get Sw(): Sw = m_sw: End Property
    Public Property Let Sw(input): m_sw = input: End Property

    Public Property Get Animate(): Animate = m_animate: End Property
    Public Property Let Animate(input): m_animate = input: End Property

    Public default Function init(primary, prim, sw, animate)
      Set m_primary = primary
      Set m_prim = prim
      m_sw = sw
      m_animate = animate

      Set Init = Me
    End Function
End Class

'Define a variable for each stand-up target
Dim ST9, ST17, ST19, ST22, ST32, ST32a

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST9 = (new StandupTarget)(sw9, psw9,9, 0)
Set ST17 = (new StandupTarget)(sw17, psw17,17, 0)
Set ST19 = (new StandupTarget)(sw19, psw19,19, 0)
Set ST22 = (new StandupTarget)(sw22, psw22,22, 0)
Set ST32 = (new StandupTarget)(sw32, psw32,32, 0)
Set ST32a = (new StandupTarget)(sw32a, psw32a,132, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST9, ST17, ST19, ST22, ST32, ST32a)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
    Dim i
    i = STArrayID(switch)

    PlayTargetSound
    STArray(i).animate = STCheckHit(ActiveBall,STArray(i).primary)

    If STArray(i).animate <> 0 Then
        DTBallPhysics ActiveBall, STArray(i).primary.orientation, STMass
    End If
    DoSTAnim
End Sub

Function STArrayID(switch)
    Dim i
    For i = 0 To UBound(STArray)
        If STArray(i).sw = switch Then
            STArrayID = i
            Exit Function
        End If
    Next
End Function

Function STCheckHit(aBall, target) 'Check if target is hit on it's face
    Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
    rangle = (target.orientation - 90) * 3.1416 / 180
    bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
    bangleafter = Atn2(aBall.vely,aball.velx)

    perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
    paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

    perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
    paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

    If perpvel > 0 And  perpvelafter <= 0 Then
        STCheckHit = 1
    ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
        STCheckHit = 1
    Else
        STCheckHit = 0
    End If
End Function

Sub DoSTAnim()
    Dim i
    For i = 0 To UBound(STArray)
        STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
    Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
    Dim animtime

    STAnimate = animate

    If animate = 0  Then
        primary.uservalue = 0
        STAnimate = 0
        Exit Function
    ElseIf primary.uservalue = 0 Then
        primary.uservalue = GameTime
    End If

    animtime = GameTime - primary.uservalue

    If animate = 1 Then
        primary.collidable = 0
        prim.transy =  - STMaxOffset
    vpmTimer.PulseSw switch mod 100
        STAnimate = 2
        Exit Function
    ElseIf animate = 2 Then
        prim.transy = prim.transy + STAnimStep
        If prim.transy >= 0 Then
            prim.transy = 0
            primary.collidable = 1
            STAnimate = 0
            Exit Function
        Else
            STAnimate = 2
        End If
    End If
End Function

'******************************************************
'****   END STAND-UP TARGETS
'******************************************************

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
Dim i
i = DTArrayID(switch)

PlayTargetSound
DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
End If
DoDTAnim
End Sub

Sub DTRaise(switch)
Dim i
i = DTArrayID(switch)

DTArray(i).animate =  - 1
DoDTAnim
End Sub

Sub DTDrop(switch)
Dim i
i = DTArrayID(switch)

DTArray(i).animate = 1
DoDTAnim
End Sub

Function DTArrayID(switch)
Dim i
For i = 0 To UBound(DTArray)
    If DTArray(i).sw = switch Then
        DTArrayID = i
        Exit Function
    End If
Next
End Function

Sub DTBallPhysics(aBall, angle, mass)
Dim rangle,bangle,calc1, calc2, calc3
rangle = (angle - 90) * 3.1416 / 180
bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)

aBall.velx = calc1 * Cos(rangle) + calc2
aBall.vely = calc1 * Sin(rangle) + calc3
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
rangle = (dtprim.rotz - 90) * 3.1416 / 180
rangle2 = dtprim.rotz * 3.1416 / 180
bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
bangleafter = Atn2(aBall.vely,aball.velx)

Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)

cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

If perpvel > 0 And  perpvelafter <= 0 Then
    If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
        DTCheckBrick = 3
    Else
        DTCheckBrick = 1
    End If
ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    DTCheckBrick = 4
Else
    DTCheckBrick = 0
End If
End Function

Sub DoDTAnim()
Dim i
For i = 0 To UBound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
Dim transz, switchid
Dim animtime, rangle

switchid = switch

Dim ind
ind = DTArrayID(switchid)

rangle = prim.rotz * PI / 180

DTAnimate = animate

If animate = 0 Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
End If

animtime = GameTime - primary.uservalue

If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    DTAnimate = animate
    Exit Function
ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
End If

If animate = 2 Then
    transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
    If prim.transz >  - DTDropUnits  Then
        prim.transz = transz
    End If

    prim.rotx = DTMaxBend * Cos(rangle) / 2
    prim.roty = DTMaxBend * Sin(rangle) / 2

    If prim.transz <= - DTDropUnits Then
        prim.transz =  - DTDropUnits
        secondary.collidable = 0
        DTArray(ind).isDropped = True 'Mark target as dropped
        controller.Switch(Switchid) = 1
        primary.uservalue = 0
        DTAnimate = 0
        Exit Function
    Else
        DTAnimate = 2
        Exit Function
    End If
End If

If animate = 3 And animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
ElseIf animate = 3 And animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
End If

If animate =  - 1 Then
    transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1

    If prim.transz =  - DTDropUnits Then
        Dim b
        Dim gBOT
        gBOT = GetBalls

        For b = 0 To UBound(gBOT)
            If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
                gBOT(b).velz = 20
            End If
        Next
    End If

    If prim.transz < 0 Then
        prim.transz = transz
    ElseIf transz > 0 Then
        prim.transz = transz
    End If

    If prim.transz > DTDropUpUnits Then
        DTAnimate =  - 2
        prim.transz = DTDropUpUnits
        prim.rotx = 0
        prim.roty = 0
        primary.uservalue = GameTime
    End If
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = False 'Mark target as not dropped
    controller.Switch(Switchid) = 0
End If

If animate =  - 2 And animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
    If prim.transz < 0 Then
        prim.transz = 0
        primary.uservalue = 0
        DTAnimate = 0

        primary.collidable = 1
        secondary.collidable = 0
    End If
End If
End Function

Function DTDropped(switchid)
Dim ind
ind = DTArrayID(switchid)

DTDropped = DTArray(ind).isDropped
End Function


'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************

' Used for drop targets
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy) 'Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
Dim AB, BC, CD, DA
AB = (bx * py) - (by * px) - (ax * py) + (ay * px) + (ax * by) - (ay * bx)
BC = (cx * py) - (cy * px) - (bx * py) + (by * px) + (bx * cy) - (by * cx)
CD = (dx * py) - (dy * px) - (cx * py) + (cy * px) + (cx * dy) - (cy * dx)
DA = (ax * py) - (ay * px) - (dx * py) + (dy * px) + (dx * ay) - (dy * ax)

If (AB <= 0 And BC <= 0 And CD <= 0 And DA <= 0) Or (AB >= 0 And BC >= 0 And CD >= 0 And DA >= 0) Then
    InRect = True
Else
    InRect = False
End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
Dim rotxy
rotxy = RotPoint(ax,ay,angle)
rax = rotxy(0) + px
ray = rotxy(1) + py
rotxy = RotPoint(bx,by,angle)
rbx = rotxy(0) + px
rby = rotxy(1) + py
rotxy = RotPoint(cx,cy,angle)
rcx = rotxy(0) + px
rcy = rotxy(1) + py
rotxy = RotPoint(dx,dy,angle)
rdx = rotxy(0) + px
rdy = rotxy(1) + py

InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************


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

Set VT1 = (new VariTarget)(VariTargetStart, VariTargetStop, VariTargetp, 1, 267, 50, 0.3, 7, 600, 0)

' Index, distance from seconardy, switch number (the first switch should fire at the first Stop {number of stops - 1})
VT1.addpoint 0, v1dist/6, 42
VT1.addpoint 1, v1dist*2/4, 43
VT1.addpoint 2, v1dist*3/4, 41

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

' Sub VTAnim_Timer
'   DoVTAnim
'   Cor.Update
' End Sub

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

    length = Distance2Obj(arr.primary, arr.secondary)

    angle = arr.primary.orientation
    transP = dSin(arr.prim.roty - 15)*arr.plength 'previous distance target has moved from start
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

    arr.prim.roty = dArcSin(transPnew/arr.plength) + 15
    VTSwitch arr, transPnew

    if arr.prim.roty >= 15 Then
      arr.prim.roty = 15
      VTSwitch arr, 0
      VTAnimate = 0
      Exit Function
    elseif cstop = transPnew and isEmpty(arr.ball) and arr.animate <> 2 Then
      VTAnimate = 0
      'debug.print cstop & " " & Controller.Switch(16) & " " & Controller.Switch(26) & " " & Controller.Switch(36) & " " & Controller.Switch(46)
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
    PlaySound SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), 0, Vol(Ball) * TargetSoundFactor * VolumeDial, AudioPan(ball), 0, 0, 0, 0, AudioFade(ball)
End Sub

'**************
' Solenoid Subs
'**************

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolVReset(Enabled)
    If Enabled Then
        VTReset 1, Enabled
    End If
End Sub

Sub SolPost(Enabled)
    TopPost.IsDropped = NOT Enabled
End Sub

Sub SolTrough(Enabled)
    If Enabled Then
        ' bsTrough.ExitSol_On
        vpmTimer.PulseSw 15
    End If
End Sub

'*******************************
' Shrunken Head Magnet by iaakki
'*******************************

SMag_timer.interval = 150
Sub SMagnet_Hit
    mShrunkenMag.AddBall ActiveBall
  SMag_timer.enabled = true
' debug.print "Mag hit"
End Sub

Sub SMagnet_unHit
    mShrunkenMag.RemoveBall ActiveBall
End Sub

dim smag_counter : smag_counter = 0
sub SMag_timer_timer
    Dim ball
  smag_counter = smag_counter + 1
  if smag_counter = 1 then        'Grab it on second loop
  elseif smag_counter = 3 then      'Grab it now
    mShrunkenMag.GrabCenter = true    'make it center the ball
  elseif smag_counter = 4 then      'stop spin
    For Each ball in mShrunkenMag.Balls
            ball.VelX = 0
      ball.VelY = 0
      ball.VelZ = 0
            ball.angmomX = 0
      ball.angmomY = 0
      ball.angmomZ = 0
        Next
  elseif smag_counter > 12 then   'cooldown     '
    mShrunkenMag.GrabCenter = false
    SMag_timer.enabled = false
    smag_counter = 0
  end if
end sub


' Opto
sub SMagnetOpto_hit
    Controller.Switch(23) = 1
end sub

sub SMagnetOpto_unhit
    Controller.Switch(23) = 0
end sub

' Speed limit
sub SMagnetLimit_hit
  if activeball.vely < -15 then
    activeball.vely = -15
  end if
end sub

' Magnet solenoid
dim smag_pulse : smag_pulse = 0
Sub SolUpperMagnet(Enabled)
  dim ball
' debug.print "mag: " & Enabled
    If Enabled Then
        mShrunkenMag.MagnetOn = 1
    smag_pulse = gametime
    Else
    SMag_timer.enabled = false
    smag_counter = 0
    mShrunkenMag.GrabCenter = false     'failsafe
        mShrunkenMag.MagnetOn = 0
'   smag_pulse = gametime - smag_pulse
'   debug.print "---> pulse length: " & smag_pulse

        For Each ball in mShrunkenMag.Balls
'            debug.print "ball.VelY --> " & ball.VelY
      if ball.vely < -1 And ball.vely > -17 then        'to flick ball more upwards if pulse has failed
        ball.vely = -17
'       debug.print "Speedup ball.VelY --> " & ball.VelY
      end if

      if ball.vely > -0.05 And ball.vely < 0.5 then     'to ensure ball is released consistently from the mag
        ball.vely = 0.7
'       debug.print "Release balance ball.VelY --> " & ball.VelY
      end if
        Next
    End If
End Sub

' ************************************
' Switches, bumpers, lanes and targets
' ************************************

Sub sw9_Hit:STHit 9:End Sub
Sub sw9o_Hit:End Sub
Sub sw17_Hit:STHit 17:End Sub
Sub sw17o_Hit:End Sub
Sub sw18_Hit:Controller.Switch(18) = 1:ActiveBall.VelX = - 1:ActiveBall.VelY = 1:RandomSoundBallBouncePlayfieldSoft ActiveBall:End Sub
Sub sw18_unHit:Controller.Switch(18) = 0:End Sub
Sub sw19_Hit:STHit 19:End Sub
Sub sw19o_Hit:End Sub
Sub sw20_Spin::vpmTimer.PulseSw 20:SoundSpinner sw20:End Sub
Sub sw21_Spin:vpmTimer.PulseSw 21:SoundSpinner sw21:End Sub

'Sub sw22_Hit:debug.print "sw22 surface " & activeball.velz:STHit 22:debug.print "sw22 surface ---> " & activeball.velz:End Sub
'Sub sw22o_Hit:debug.print "sw22o " & activeball.velz:TargetBouncer ActiveBall,0.8:debug.print "sw22o ---> " & activeball.velz:End Sub

Sub sw22_Hit:debug.print "sw22 surface " & activeball.velz:STHit 22:debug.print "sw22 surface ---> " & activeball.velz:End Sub
Sub sw22o_Hit:debug.print "sw22o " & activeball.velz:End Sub

Sub IMagnet_Hit:mIdolMag.AddBall ActiveBall:Controller.Switch(24) = 1:End Sub
Sub IMagnet_unHit:mIdolMag.RemoveBall ActiveBall:Controller.Switch(24) = 0:End Sub

'right bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49:RandomSoundBumperTop(Bumper1):End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 51:RandomSoundBumperMiddle(Bumper2):End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 50:RandomSoundBumperBottom(Bumper3):End Sub

'left bumpers
Sub Bumper4_Hit:vpmTimer.PulseSw 25:RandomSoundBumperTop(Bumper4):End Sub
Sub Bumper5_Hit:vpmTimer.PulseSw 27:RandomSoundBumperMiddle(Bumper5):End Sub
Sub Bumper6_Hit:vpmTimer.PulseSw 26:RandomSoundBumperBottom(Bumper6):End Sub

Sub sw28_Hit:SoundSaucerLock:controller.switch(28)=1:End Sub
Sub sw28_unhit:controller.Switch(28)=0:End Sub
Sub sw29_Hit:SoundSaucerLock:controller.switch(29)=1:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1:End Sub
Sub sw30_unHit:Controller.Switch(30) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1:End Sub
Sub sw31_unHit:Controller.Switch(31) = 0:End Sub
Sub sw32_Hit:STHit 32:End Sub
Sub sw32o_Hit:End Sub
Sub sw32a_Hit:STHit 132:End Sub
Sub sw32ao_Hit:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:End Sub
Sub sw33_unHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:End Sub
Sub sw34_unHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:End Sub
Sub sw35_unHit:Controller.Switch(35) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:End Sub
Sub sw36_unHit:Controller.Switch(36) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:End Sub
Sub sw38_unHit:Controller.Switch(38) = 0:RandomSoundRollover:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:End Sub
Sub sw39_unHit:Controller.Switch(39) = 0:RandomSoundRollover:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:End Sub
Sub sw40_unHit:Controller.Switch(40) = 0:RandomSoundRollover:End Sub
Sub sw44_Hit:Controller.Switch(44) = 1:Debug.Print("switch 44 hit"):End Sub
Sub sw44_unHit:Controller.Switch(44) = 0:End Sub
Sub sw45_Hit:Controller.Switch(45) = 1:Debug.Print("switch 45 hit"):End Sub
Sub sw45_unHit:Controller.Switch(45) = 0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:Debug.Print("switch 46 hit"):End Sub
Sub sw46_unHit:Controller.Switch(46) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_unHit:Controller.Switch(47) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_unHit:Controller.Switch(48) = 0:End Sub



'//////////////////////////////////////////////////////////////////////
'// RAMP TRIGGERS
'//////////////////////////////////////////////////////////////////////

'Plastic_Start

Sub ramptrigger004A_hit()
  WireRampOn True 'Play Plastic Ramp Sound
    bsRampOnClear
End Sub

Sub ramptrigger005A_hit()
  WireRampOn True 'Play Plastic Ramp Sound
    bsRampOnClear
End Sub

Sub ramptrigger006A_hit()
  WireRampOn True 'Play Plastic Ramp Sound
    bsRampOnClear
End Sub

'Plastic_End

Sub REnd3_Hit()
  WireRampOff
    bsRampOff ActiveBall.ID
  ActiveBall.VelY = ActiveBall.VelY * 0.5
End Sub

Sub REnd3_UnHit()
  RandomSoundDelayedBallDropOnPlayfield(ActiveBall)
End Sub

'Wire_Start

Sub ramptrigger001A_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
    bsRampOff ActiveBall.ID
End Sub

Sub ramptrigger001A_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
    bsRampOnWire
End Sub

Sub ramptrigger002A_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
    bsRampOff ActiveBall.ID
End Sub

Sub ramptrigger002A_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
    bsRampOnWire
End Sub

Sub ramptrigger003A_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
    bsRampOff ActiveBall.ID
End Sub

Sub ramptrigger003A_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
    bsRampOnWire
End Sub


'Wire_End

Sub REnd1_Hit()
    WireRampOff
    bsRampOff ActiveBall.ID
End Sub

Sub REnd1_UnHit()
  RandomSoundRampStop REnd1
  RandomSoundDelayedBallDropOnPlayfield(ActiveBall)
End Sub

Sub REnd2_Hit()
  WireRampOff
    bsRampOff ActiveBall.ID
End Sub

Sub REnd2_UnHit()
  RandomSoundRampStop REnd2
  RandomSoundDelayedBallDropOnPlayfield(ActiveBall)
End Sub

Sub REnd4_Hit()
  WireRampOff
    bsRampOff ActiveBall.ID
 End Sub

Sub REnd4_UnHit()
  RandomSoundRampStop REnd4
    RandomSoundDelayedBallDropOnPlayfield(ActiveBall)
End Sub

Sub REnd5_Hit()
  WireRampOff
    bsRampOff ActiveBall.ID
End Sub

Sub REnd5_UnHit()
  RandomSoundRampStop REnd5
  RandomSoundDelayedBallDropOnPlayfield(ActiveBall)
End Sub

Sub GateSound2_Hit()
  If ActiveBall.velx > 1 Then
    SoundPlayfieldGate
  End If
End Sub

Sub GateSound3_Hit()
  If ActiveBall.velx > 1 Then
    SoundPlayfieldGate
  End If
End Sub

Sub GateSound5_Hit()
  If ActiveBall.velx < -1 Then
    SoundPlayfieldGate
  End If
End Sub


Sub sw42_Hit
    SoundPlayfieldGate
    If ActiveBall.VelY < 0 Then
        Controller.Switch(42) = 1
        ActiveBall.VelY = ActiveBall.VelY * 0.8
    End If
  If ActiveBall.z < 0 Then
    ActiveBall.z = 0
  End If
End Sub

Sub sw43_Hit
    If ActiveBall.VelY < 0 Then
        Controller.Switch(43) = 1
        ActiveBall.VelY = ActiveBall.VelY * 0.8
    End If
  If ActiveBall.z < 0 Then
    ActiveBall.z = 0
  End If
End Sub

Sub sw41_Hit
    If ActiveBall.VelY < 0 Then
        Controller.Switch(41) = 1
        ActiveBall.VelY = ActiveBall.VelY * 0.8
    End If
  If ActiveBall.z < 0 Then
    ActiveBall.z = 0
  End If
End Sub

Sub VariTimerDown_Timer
    VariTarget.RotY = VariTarget.RotY + 1
    If VariTarget.RotY = -5 Then Controller.Switch(41) = 0
    If VariTarget.RotY = 5 Then Controller.Switch(43) = 0
    If VariTarget.RotY > 15 Then
        VariTarget.RotY = 15
        Controller.Switch(42) = 0
        VariTimerDown.Enabled = 0
    End If
End Sub

'*********************************************************************
' Narnia Catcher
'*********************************************************************

Sub Narnia_Timer
  Dim b, gBOT
  gBOT = GetBalls
  For b = 0 to UBound(gBOT)
    'Check for narnia balls
    If gBOT(b).z < -200 Then
      'move ball to varitarget vuk
      gBOT(b).x = 674.1407 : gBOT(b).y = 504.4824 : gBOT(b).z = 26
      gBOT(b).velx = 0 : gBOT(b).vely = 0 : gBOT(b).velz = 0
    end if
  Next
End Sub

' Sub VariTargetStop_Hit()
'   ActiveBall.x = 674.1407 : ActiveBall.y = 504.4824 : ActiveBall.z = 26
'   ActiveBall.velx = 0 : ActiveBall.vely = 0 : ActiveBall.velz = 0
' End Sub

'******
' vuks / lock
'******

Sub sw52_Hit():
    Controller.Switch(52) = 1
    SoundSaucerLock
End Sub

Sub SolVUK(enabled):
    If enabled Then
        sw52.Kick 0, 36, 1.5
        SoundSaucerKick 1,sw52
        WireRampOn False
        Controller.Switch(52) = 0
    End If
End Sub

Sub Lock_Hit():
    SoundSaucerLock
End Sub

Sub SolLock(enabled):
    If enabled Then
        RandomSoundFlipperUpRight RightFlipper
        Lock.Kick 0, 36+rnd*4
    End If
End Sub

Sub SolSkillScoop(enabled):
    If enabled Then
        SoundSaucerKick 1,sw29
        sw29.Kick 177+rnd*4, 100+rnd*7
        Controller.Switch(29) = 0
    End If
End Sub


'*******
' lanes
'*******

Sub sw53_Hit:Controller.Switch(53) = 1:End Sub
Sub sw53_unHit:Controller.Switch(53) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:RandomSoundRollover:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:RandomSoundRollover:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:RandomSoundRollover:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:RandomSoundRollover:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

'************
' Slingshots
'************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    RandomSoundSlingshotLeft(PegPlasticT2)
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 59
    LeftSlingShot.TimerEnabled = 1
  LS.VelocityCorrect ActiveBall
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    RandomSoundSlingshotRight(PegPlasticT1)
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 62
    RightSlingShot.TimerEnabled = 1
  RS.VelocityCorrect ActiveBall
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'**************
' LED's display
'**************

Dim LED(14)

LED(0) = Array(LED1, LED2, LED3, LED4, LED5, LED6, LED7)
LED(1) = Array(LED8, LED9, LED10, LED11, LED12, LED13, LED14)
LED(2) = Array(LED15, LED16, LED17, LED18, LED19, LED20, LED21)
LED(3) = Array(LED22, LED23, LED24, LED25, LED26, LED27, LED28)
LED(4) = Array(LED29, LED30, LED31, LED32, LED33, LED34, LED35)

LED(5) = Array(LED36, LED37, LED38, LED39, LED40, LED41, LED42)
LED(6) = Array(LED43, LED44, LED45, LED46, LED47, LED48, LED49)
LED(7) = Array(LED50, LED51, LED52, LED53, LED54, LED55, LED56)
LED(8) = Array(LED57, LED58, LED59, LED60, LED61, LED62, LED63)
LED(9) = Array(LED64, LED65, LED66, LED67, LED68, LED69, LED70)

LED(10) = Array(LED71, LED72, LED73, LED74, LED75, LED76, LED77)
LED(11) = Array(LED78, LED79, LED80, LED81, LED82, LED83, LED84)
LED(12) = Array(LED85, LED86, LED87, LED88, LED89, LED90, LED91)
LED(13) = Array(LED92, LED93, LED94, LED95, LED96, LED97, LED98)
LED(14) = Array(LED99, LED100, LED101, LED102, LED103, LED104, LED105)

Sub UpdateLeds()
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&H00000000, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In LED(num)
                If chg And 1 Then obj.Visible = stat And 1
                chg = chg \ 4:stat = stat \ 4
            Next
        Next
    End If
End Sub

'*******************************************************'
'               Fluppers Flasher Domes
'*******************************************************'
Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1      ' *** change this, if your table has another name          ***
FlasherLightIntensity = 0.1  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)      ***
FlasherFlareIntensity = 0.1  ' *** lower this, if the flares are too bright (i.e. 0.1)        ***
FlasherBloomIntensity = 0.4  ' *** lower this, if the blooms are too bright (i.e. 0.1)        ***
FlasherOffBrightness = 0.6    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "white"
InitFlasher 2, "white"
InitFlasher 3, "white"
InitFlasher 4, "white"
InitFlasher 5, "white"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateFlasher 1,17
'RotateFlasher 2,
'   RotateFlasher 3,90
'   RotateFlasher 4,90

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr)
  Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr)
  Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)

  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbase(nr).x) / (objbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ
    objflasher(nr).height = objbase(nr).z + 40
  End If

  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0
  objlit(nr).visible = 0
  objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX
  objlit(nr).RotY = objbase(nr).RotY
  objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX
  objlit(nr).ObjRotY = objbase(nr).ObjRotY
  objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x
  objlit(nr).y = objbase(nr).y
  objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 Then
    objflasher(nr).y = objbase(nr).y + 50
    objflasher(nr).height = objbase(nr).z + 20
  Else
    objflasher(nr).y = objbase(nr).y + 20
    objflasher(nr).height = objbase(nr).z + 50
  End If
  objflasher(nr).x = objbase(nr).x

  'rothbauerw
  'Adjust the position of the light object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  objlight(nr).x = objbase(nr).x
  objlight(nr).y = objbase(nr).y
  objlight(nr).bulbhaloheight = objbase(nr).z - 10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  Dim xthird, ythird
  xthird = tablewidth / 3
  ythird = tableheight / 3
  If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  End If

  ' set the texture and color of all objects
  Select Case objbase(nr).image
    Case "dome2basewhite"
    objbase(nr).image = "dome2base" & col
    objlit(nr).image = "dome2lit" & col

    Case "ronddomebasewhite"
    objbase(nr).image = "ronddomebase" & col
    objlit(nr).image = "ronddomelit" & col

    Case "domeearbasewhite"
    objbase(nr).image = "domeearbase" & col
    objlit(nr).image = "domeearlit" & col
  End Select
  If TestFlashers = 0 Then
    objflasher(nr).imageA = "domeflashwhite"
    objflasher(nr).visible = 0
  End If
  Select Case col
    Case "blue"
    objlight(nr).color = RGB(4,120,255)
    objflasher(nr).color = RGB(200,255,255)
    objbloom(nr).color = RGB(4,120,255)
    objlight(nr).intensity = 5000

    Case "green"
    objlight(nr).color = RGB(12,255,4)
    objflasher(nr).color = RGB(12,255,4)
    objbloom(nr).color = RGB(12,255,4)

    Case "red"
    objlight(nr).color = RGB(255,32,4)
    objflasher(nr).color = RGB(255,32,4)
    objbloom(nr).color = RGB(255,32,4)

    Case "purple"
    objlight(nr).color = RGB(230,49,255)
    objflasher(nr).color = RGB(255,64,255)
    objbloom(nr).color = RGB(230,49,255)

    Case "yellow"
    objlight(nr).color = RGB(200,173,25)
    objflasher(nr).color = RGB(255,200,50)
    objbloom(nr).color = RGB(200,173,25)

    Case "white"
    objlight(nr).color = RGB(255,240,150)
    objflasher(nr).color = RGB(100,86,59)
    objbloom(nr).color = RGB(255,240,150)

    Case "orange"
    objlight(nr).color = RGB(255,70,0)
    objflasher(nr).color = RGB(255,70,0)
    objbloom(nr).color = RGB(255,70,0)
  End Select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle)
  angle = ((angle + 360 - objbase(nr).ObjRotZ) Mod 180) / 30
  objbase(nr).showframe(angle)
  objlit(nr).showframe(angle)
End Sub

Const FlashMaxLvl = 1
Dim lastFlashLvl : lastFlashlvl = 0

Sub ModFlashFlasher(nr, aValue)
  dim lvl
  lvl = min(1,aValue/FlashMaxLvl)

  If lvl < 0.01 And lvl Then
    Sound_Flash_Relay 0, Bumper1
  ElseIf lvl > 0.99 And lvl Then
    Sound_Flash_Relay 1, Bumper1
  End If


  if aValue > 0 then
    objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1
  else
    objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0
  end if
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * aValue
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * aValue
  objlight(nr).IntensityScale = 10 * FlasherLightIntensity * aValue
  objbase(nr).BlendDisableLighting = FlasherOffBrightness + 10 * aValue
  objlit(nr).BlendDisableLighting = 10 * aValue
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,aValue,RGB(255,255,255),0,0,False,True,0,0,0,0
End Sub


'******************************************************
'   END FLUPPER DOMES
'******************************************************

'******************************
'Ripley's Flasher Domes Config
'*******************************

' Lower Left Flasher
Sub Flash125(pwm)
  ModFlashFlasher 2,pwm
  f25.intensityscale = pwm
  f25a.intensityscale = pwm
  flasherlight2a.intensityscale = pwm
  pInsertOn_f25.blenddisablelighting = 10 * pwm
End Sub


' Center playfield/Varitarget Flasher'
Sub Flash129(pwm)
  ModFlashFlasher 1,pwm
End Sub


' Upper Right Flasher
Sub Flash130(pwm)
  ModFlashFlasher 4,pwm
  f30.intensityscale = pwm
  f30a.intensityscale = pwm
  f30b.intensityscale = pwm
  pInsertOn_f30.blenddisablelighting = pwm
End Sub


' Lower Right Flasher
Sub Flash132(pwm)
  ModFlashFlasher 3,pwm
  Flasherlight3a.intensityscale = pwm
End Sub


' Upper Left Flasher (Triggered by F127 Below)
Sub FlashWhite5(pwm)
  ModFlashFlasher 5,pwm
End Sub

'******************************************************
'         Other Flashers
'******************************************************

f27.intensityscale = 0
sub Flash127(pwm)
  f27.intensityscale = pwm
  FlashWhite5 pwm
end sub

f28.intensityscale = 0
sub Flash128(pwm)
  f28.intensityscale = pwm
end sub


'******************************************************
'   ZLMP:  LAMPZ by nFozzy
'******************************************************
'
' Lampz is a utility designed to manage and fade the lights and light-related objects on a table that is being driven by a ROM.
' To set up Lampz, one must populate the Lampz.MassAssign array with VPX Light objects, where the index of the MassAssign array
' corrisponds to the ROM index of the associated light. More that one Light object can be associated with a single MassAssign index (not shown in this example)
' Optionally, callbacks can be assigned for each index using the Lampz.Callback array. This is very useful for allowing 3D Insert primitives
' to be controlled by the ROM. Note, the aLvl parameter (i.e. the fading level that ranges between 0 and 1) is appended to the callback call.

Dim NullFader
Set NullFader = New NullFadingObject
Dim Lampz
Set Lampz = New VPMLampUpdater
InitLampsNF         ' Setup lamp assignments
LampTimer.Interval =  - 1
LampTimer.Enabled = 1

Sub LampTimer_Timer()
  Dim x, chglamp
  chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp) 'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)/255.0
    Next
  End If
End Sub

Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  'If Lampz.UseFunc Then aLvl = Lampz.FilterOut(aLvl) 'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

Sub DisableLighting2(pri, DLlow, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  'If Lampz.UseFunc Then aLvl = Lampz.FilterOut(aLvl) 'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * (DLintensity-DLlow) + DLlow
End Sub

Sub SetModLamp(id, val)
  Lampz.state(id) = val
End Sub

Sub InitLampsNF()

  Lampz.MassAssign(1) = l1
  Lampz.Callback(1) = "DisableLighting pInsertOn_1, 80,"


  Lampz.MassAssign(2) = l2
  Lampz.Callback(2) = "DisableLighting pInsertOn_2, 80,"

  Lampz.MassAssign(3) = l3
  Lampz.Callback(3) = "DisableLighting pInsertOn_3, 80,"

  Lampz.MassAssign(4) = l4
  Lampz.Callback(4) = "DisableLighting pInsertOn_4, 80,"


  Lampz.MassAssign(5) = l5
  Lampz.Callback(5) = "DisableLighting pInsertOn_5, 80,"
  Lampz.MassAssign(6) = l6
  Lampz.Callback(6) = "DisableLighting pInsertOn_6, 80,"
  Lampz.MassAssign(7) = l7
  Lampz.Callback(7) = "DisableLighting pInsertOn_7, 80,"
  Lampz.MassAssign(8) = l8
  Lampz.MassAssign(8) = l8a
  Lampz.Callback(8) = "DisableLighting pInsertOn_8, 80,"
  Lampz.MassAssign(9) = l9
  Lampz.MassAssign(9) = l9a
  Lampz.Callback(9) = "DisableLighting pInsertOn_9, 80,"
  Lampz.MassAssign(10) = l10
  Lampz.Callback(10) = "DisableLighting pInsertOn_10, 80,"
  Lampz.MassAssign(11) = l11
  Lampz.Callback(11) = "DisableLighting pInsertOn_11, 80,"
  Lampz.MassAssign(12) = l12
  Lampz.Callback(12) = "DisableLighting pInsertOn_12, 80,"
  Lampz.MassAssign(13) = l13
  Lampz.Callback(13) = "DisableLighting pInsertOn_13, 80,"
  Lampz.MassAssign(14) = l14
  Lampz.Callback(14) = "DisableLighting pInsertOn_14, 80,"
  Lampz.MassAssign(15) = l15
  Lampz.MassAssign(15) = l15a
  Lampz.Callback(15) = "DisableLighting pInsertOn_15, 80,"
' Lampz.MassAssign(16) = l16
  Lampz.Callback(16) = "DisableLighting arrow_bulb, 5,"
  Lampz.MassAssign(17) = l17
  Lampz.Callback(17) = "DisableLighting pInsertOn_17, 80,"
  Lampz.MassAssign(18) = l18
  Lampz.Callback(18) = "DisableLighting pInsertOn_18, 80,"
  Lampz.MassAssign(19) = l19
  Lampz.Callback(19) = "DisableLighting pInsertOn_19, 80,"
  Lampz.MassAssign(20) = l20
  Lampz.Callback(20) = "DisableLighting pInsertOn_20, 80,"
  Lampz.MassAssign(21) = l21
  Lampz.Callback(21) = "DisableLighting pInsertOn_21, 80,"
  Lampz.MassAssign(22) = l22
  Lampz.Callback(22) = "DisableLighting pInsertOn_22, 80,"
  Lampz.MassAssign(23) = l23
  Lampz.Callback(23) = "DisableLighting pInsertOn_23, 80,"
  Lampz.MassAssign(24) = l24
  Lampz.Callback(24) = "DisableLighting pInsertOn_24, 44,"
  Lampz.MassAssign(25) = l25
  Lampz.Callback(25) = "DisableLighting pInsertOn_25, 80,"
  Lampz.MassAssign(26) = l26
  Lampz.Callback(26) = "DisableLighting pInsertOn_26, 80,"
  Lampz.MassAssign(27) = l27
  Lampz.Callback(27) = "DisableLighting pInsertOn_27, 22,"
  Lampz.MassAssign(28) = l28
  Lampz.Callback(28) = "DisableLighting pInsertOn_28, 80,"
  Lampz.MassAssign(29) = l29
  Lampz.Callback(29) = "DisableLighting pInsertOn_29, 80,"
  Lampz.MassAssign(30) = l30
  Lampz.Callback(30) = "DisableLighting pInsertOn_30, 22,"
  Lampz.MassAssign(31) = l31
  Lampz.Callback(31) = "DisableLighting pInsertOn_31, 80,"
  Lampz.MassAssign(32) = l32
  Lampz.Callback(32) = "DisableLighting pInsertOn_32, 80,"



  Lampz.MassAssign(33) = bumperbiglight4
  Lampz.MassAssign(33) = bumpersmalllight4
  Lampz.MassAssign(33) = bumperhighlight4
  Lampz.Callback(33) = "DisableLighting2 bumpertop4, 1 , 20,"
  Lampz.Callback(33) = "DisableLighting2 bumperbulb4, 1 , 100,"

  Lampz.MassAssign(34) = bumperbiglight5
  Lampz.MassAssign(34) = bumpersmalllight5
  Lampz.MassAssign(34) = bumperhighlight5
  Lampz.Callback(34) = "DisableLighting2 bumpertop5, 1 , 20,"
  Lampz.Callback(34) = "DisableLighting2 bumperbulb5, 1 , 100,"

  Lampz.MassAssign(35) = bumperbiglight6
  Lampz.MassAssign(35) = bumpersmalllight6
  Lampz.MassAssign(35) = bumperhighlight6
  Lampz.Callback(35) = "DisableLighting2 bumpertop6, 1 , 20,"
  Lampz.Callback(35) = "DisableLighting2 bumperbulb6, 1 , 100,"

  Lampz.MassAssign(36) = l36
  Lampz.Callback(36) = "DisableLighting pInsertOn_36, 55,"
  Lampz.MassAssign(37) = l37
  Lampz.Callback(37) = "DisableLighting pInsertOn_37, 80,"
  Lampz.MassAssign(38) = l38
  Lampz.Callback(38) = "DisableLighting pInsertOn_38, 80,"
  Lampz.MassAssign(39) = l39
  Lampz.Callback(39) = "DisableLighting pInsertOn_39, 80,"
  Lampz.MassAssign(40) = l40
  Lampz.Callback(40) = "DisableLighting pInsertOn_40, 55,"
  Lampz.MassAssign(41) = l41
  Lampz.Callback(41) = "DisableLighting pInsertOn_41, 80,"
  Lampz.MassAssign(42) = l42
  Lampz.Callback(42) = "DisableLighting pInsertOn_42, 80,"
  Lampz.MassAssign(43) = l43
  Lampz.Callback(43) = "DisableLighting pInsertOn_43, 80,"
  Lampz.MassAssign(44) = l44
  Lampz.Callback(44) = "DisableLighting pInsertOn_44, 80,"
  Lampz.MassAssign(45) = l45
  Lampz.Callback(45) = "DisableLighting pInsertOn_45, 80,"
  Lampz.MassAssign(46) = l46
  Lampz.Callback(46) = "DisableLighting pInsertOn_46, 80,"
  Lampz.MassAssign(47) = l47
  Lampz.Callback(47) = "DisableLighting pInsertOn_47, 80,"
  Lampz.MassAssign(48) = l48
  Lampz.Callback(48) = "DisableLighting pInsertOn_48, 80,"
  Lampz.MassAssign(49) = l49
  Lampz.MassAssign(49) = l49a
  Lampz.Callback(49) = "DisableLighting pInsertOn_49, 80,"
  Lampz.MassAssign(50) = l50
  Lampz.Callback(50) = "DisableLighting pInsertOn_50, 55,"
  Lampz.MassAssign(51) = l51
  Lampz.Callback(51) = "DisableLighting pInsertOn_51, 55,"
  Lampz.MassAssign(52) = l52
  Lampz.Callback(52) = "DisableLighting pInsertOn_52, 80,"
  Lampz.MassAssign(53) = l53
  Lampz.Callback(53) = "DisableLighting pInsertOn_53, 55,"
  Lampz.MassAssign(54) = l54
  Lampz.Callback(54) = "DisableLighting pInsertOn_54, 55,"
  Lampz.MassAssign(55) = l55
  Lampz.Callback(55) = "DisableLighting pInsertOn_55, 55,"
  Lampz.MassAssign(56) = l56
  Lampz.Callback(56) = "DisableLighting pInsertOn_56, 80,"
  Lampz.MassAssign(57) = l57
  Lampz.Callback(57) = "DisableLighting pInsertOn_57, 55,"
  Lampz.MassAssign(58) = l58
  Lampz.Callback(58) = "DisableLighting pInsertOn_58, 55,"
  Lampz.MassAssign(59) = l59
  Lampz.Callback(59) = "DisableLighting pInsertOn_59, 55,"

  Lampz.MassAssign(60) = bumperbiglight1
  Lampz.MassAssign(60) = bumpersmalllight1
  Lampz.MassAssign(60) = bumperhighlight1
  Lampz.Callback(60) = "DisableLighting2 bumpertop1, 1 , 20,"
  Lampz.Callback(60) = "DisableLighting2 bumperbulb1, 1 , 200,"

  Lampz.MassAssign(61) = bumperbiglight2
  Lampz.MassAssign(61) = bumpersmalllight2
  Lampz.MassAssign(61) = bumperhighlight2
  Lampz.Callback(61) = "DisableLighting2 bumpertop2, 1 , 10,"       'this appear really bright for some reason
  Lampz.Callback(61) = "DisableLighting2 bumperbulb2, 1 , 200,"

  Lampz.MassAssign(62) = bumperbiglight3
  Lampz.MassAssign(62) = bumpersmalllight3
  Lampz.MassAssign(62) = bumperhighlight3
  Lampz.Callback(62) = "DisableLighting2 bumpertop3, 1 , 20,"
  Lampz.Callback(62) = "DisableLighting2 bumperbulb3, 1 , 200,"


  Lampz.MassAssign(63) = l63
  Lampz.MassAssign(63) = l63a
  Lampz.Callback(63) = "DisableLighting pInsertOn_63, 80,"
  Lampz.MassAssign(64) = l64
  Lampz.Callback(64) = "DisableLighting pInsertOn_64, 80,"
  Lampz.MassAssign(65) = l65
  Lampz.Callback(65) = "DisableLighting pInsertOn_65, 80,"
  Lampz.MassAssign(66) = l66
  Lampz.Callback(66) = "DisableLighting pInsertOn_66, 44,"
  Lampz.MassAssign(67) = l67
  Lampz.Callback(67) = "DisableLighting pInsertOn_67, 44,"
  Lampz.MassAssign(68) = l68
  Lampz.Callback(68) = "DisableLighting pInsertOn_68, 44,"
  Lampz.MassAssign(69) = l69
  Lampz.MassAssign(69) = l69a
  Lampz.Callback(69) = "DisableLighting pInsertOn_69, 80,"
  Lampz.MassAssign(70) = l70
  Lampz.Callback(70) = "DisableLighting pInsertOn_70, 44,"
  Lampz.MassAssign(71) = l71
  Lampz.Callback(71) = "DisableLighting pInsertOn_71, 44,"
  Lampz.MassAssign(72) = l72
  Lampz.Callback(72) = "DisableLighting pInsertOn_72, 55,"
  Lampz.MassAssign(73) = l73
  Lampz.MassAssign(73) = l73a
  Lampz.Callback(73) = "DisableLighting pInsertOn_73, 80,"
  Lampz.MassAssign(74) = l74
  Lampz.MassAssign(74) = l74a
  Lampz.Callback(74) = "DisableLighting pInsertOn_74, 80,"
  Lampz.MassAssign(75) = l75
  Lampz.MassAssign(75) = l75a
  Lampz.Callback(75) = "DisableLighting pInsertOn_75, 80,"
  Lampz.MassAssign(76) = l76
  Lampz.Callback(76) = "DisableLighting pInsertOn_76, 44,"
  Lampz.MassAssign(77) = l77
  Lampz.Callback(77) = "DisableLighting pInsertOn_77, 44,"
  Lampz.MassAssign(78) = l78
  Lampz.Callback(78) = "DisableLighting pInsertOn_78, 44,"


'**Flasher Lamps

  Lampz.MassAssign(102) = f22
  Lampz.MassAssign(102) = f22a
  Lampz.Callback(102) = "DisableLighting pInsertOn_f22, 100," 'Left Spinner Insert

' Lampz.MassAssign(125) = f25
' Lampz.MassAssign(125) = f25a
' Lampz.Callback(125) = "DisableLighting pInsertOn_f25, 100," 'Lower Pop Bumper Insert

  Lampz.MassAssign(106) = f26
  Lampz.MassAssign(106) = f26a
  Lampz.Callback(106) = "DisableLighting pInsertOn_65, 30," 'Left Spinner Insert

' Lampz.MassAssign(130) = f30
' Lampz.MassAssign(130) = f30a
' Lampz.MassAssign(130) = f30b
' Lampz.Callback(130) = "DisableLighting pInsertOn_f30, 100," 'Top Pop Bumper Insert

  Lampz.MassAssign(111) = f31
  Lampz.MassAssign(111) = f31a
  Lampz.Callback(111) = "DisableLighting pInsertOn_52, 30," 'Right Spinner Insert

' Lampz.MassAssign(105) = f125
' Lampz.Callback(105) = "DisableLighting pf125, 200,"
' Lampz.MassAssign(108) = f128
' Lampz.Callback(108) = "DisableLighting pf128, 200,"
' Lampz.MassAssign(111) = f131
' Lampz.Callback(111) = "DisableLighting pf131, 200,"
' Lampz.MassAssign(112) = f132
' Lampz.Callback(112) = "DisableLighting pf132, 200," 'Commented Out Because It Has No Primitive Insert

  'This just turns state of any lamps to 1
  Lampz.Init

  'Turn off all lamps on startup
  Dim x: For x = 0 to 150: Lampz.State(x) = 0: Next

End Sub


'******************************************************
'****  ZGIU:  GI Control
'******************************************************


dim gilvl:gilvl = 0
Dim gistuff
'**** Global variable to hold current GI light intensity. Not Mandatory, but may help in other lighting subs and timers
'**** This value is updated always when GI state is being updated
wire_left_on.opacity = 0
wire_right_on.opacity = 0
metalguides_on.opacity = 0
cramp_on.opacity = 0
rramp_on.opacity = 0
sideblades_on.opacity = 0
plastics_on.opacity = 0
brackets_on.opacity = 0


Set GICallback2 = GetRef("GIUpdate2")

Sub GIUpdate2(no, level)
  debug.print "GIUpdate2 no="&no&" level="&level
  Dim bulb
  For each bulb in aGiLights: bulb.State = level: Next
  If level >= 0.5 And gilvl < 0.5 Then
    ' if VRRoom > 0 then PinCab_Backglass.image = "Backglass_On"
    Sound_GI_Relay 1, Bumper1
  ElseIf level <= 0.4 And gilvl > 0.4 Then
    ' if VRRoom > 0 then PinCab_Backglass.image = "Backglass_Off"
    Sound_GI_Relay 0, Bumper1
  End If
  gilvl = level
  GIUpdatePrims
End Sub


Sub GIUpdatePrims()
  PinCab_Backglass.blenddisablelighting = 2 * gilvl + 0.2

  wire_left_on.opacity = gilvl * 200
  wire_right_on.opacity = gilvl * 100
  metalguides_on.opacity = gilvl * 100
  cramp_on.opacity = gilvl * 100
  rramp_on.opacity = gilvl * 100
  sideblades_on.opacity = gilvl * 120
  plastics_on.opacity = gilvl * 100
  brackets_on.opacity = gilvl * 150
  arrow_plastic.blenddisablelighting = 1.2*gilvl + 0.6
  dyk_plastic.blenddisablelighting = 0.4*gilvl + 0.2
  shrunkenHead.blenddisablelighting = 0.4*gilvl + 0.1

  FGI.opacity = 100 * gilvl
  FShadow.opacity = (1-gilvl) * 50 + 25

  For Each gistuff in backgif : gistuff.opacity = gilvl * 20 : Next
  For Each gistuff in backgip : gistuff.blenddisablelighting =  gilvl * 0.5 : Next

  For Each gistuff in GILitPosts : gistuff.blenddisablelighting =  gilvl * 0.5 : Next
End Sub

'******************************************************
'****  END GI Control
'******************************************************


'******************************************************
'   ZBBR: Ball Brightness
'******************************************************

Const BallBrightness =  1        'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
Const PLOffset = 0.5
Dim PLGain: PLGain = (1-PLOffset)/(1260-2000)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,870,2000,870,1260,930,1260,930,2000) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-2000))
    End If
    ' Assign color
    b_r = Int(d_w)
    b_g = Int(d_w)
    b_b = Int(d_w)
    If b_r > 255 Then b_r = 255
    If b_g > 255 Then b_g = 255
    If b_b > 255 Then b_b = 255
    gBOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
    'debug.print "--- ball.color level="&b_r
  Next
End Sub



'====================
'Class jungle nf
'====================

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?
Class NullFadingObject
  Public Property Let IntensityScale(input)

  End Property
End Class


Class VPMLampUpdater
  Public Name
  Public Obj(150), OnOff(150)
  Private UseCallback(150), cCallback(150)

  Sub Class_Initialize()
    Name = "VPMLampUpdater" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    Dim x : For x = 0 to uBound(OnOff)
        OnOff(x) = 0
      Set Obj(x) = NullFader
    Next
  End Sub

  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    'cCallback(idx) = String 'old execute method
    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    dim tmp : tmp = Split(cCallback(idx), "___")

    dim str, x : for x = 0 to uBound(tmp) 'build proc contents
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x)="" then str = str & tmp(x) & " aLVL:"
    Next

    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    'if idx = 132 then msgbox out 'debug
    ExecuteGlobal Out
  End Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
      end if
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    end if
  End Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub

  Public Property Let state(ByVal x, input)
    Dim xx
    OnOff(x) = input
    If IsArray(obj(x)) Then
      For Each xx In obj(x)
        xx.IntensityScale = input
        'debug.print x&"  obj.Intensityscale = " & input
      Next
    Else
      obj(x).Intensityscale = input
      'debug.print "obj("&x&").Intensityscale = " & input
    End if
    'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
    If UseCallBack(x) then Proc name & x,input
  End Property

  Public Property Get state(idx) : state = OnOff(idx) : end Property

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
          tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
        obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
      end if
    Next
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

End Class


'Helper functions
Sub Proc(String, Callback)  'proc using a string and one argument
  'On Error Resume Next
  Dim p
  Set P = GetRef(String)
  P Callback
  If err.number = 13 Then  MsgBox "Proc error! No such procedure: " & vbNewLine & String
  If err.number = 424 Then MsgBox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the End of a 1 dimensional array
  If IsArray(aInput) Then 'Input is an array...
    Dim tmp
    tmp = aArray
    If Not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else          'Append existing array with aInput array
      ReDim Preserve tmp(UBound(aArray) + UBound(aInput) + 1) 'If existing array, increase bounds by uBound of incoming array
      Dim x
      For x = 0 To UBound(aInput)
        If IsObject(aInput(x)) Then
          Set tmp(x + UBound(aArray) + 1 ) = aInput(x)
        Else
          tmp(x + UBound(aArray) + 1 ) = aInput(x)
        End If
      Next
      AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If Not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      ReDim Preserve aArray(UBound(aArray) + 1) 'If array, increase bounds by 1
      If IsObject(aInput) Then
        Set aArray(UBound(aArray)) = aInput
      Else
        aArray(UBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function


'***********************class jungle**************

Function ColtoArray(aDict)  'converts a collection to an indexed array. Indexes will come out random probably.
  redim a(999)
  dim count : count = 0
  dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
  redim preserve a(count-1) : ColtoArray = a
End Function

'******************************************************
'****  END LAMPZ
'******************************************************


'******************************************************
'   Z3DI:   3D INSERTS
'******************************************************
'
'
' Before you get started adding the inserts to your playfield in VPX, there are a few things you need to have done to prepare:
'  1. Cut out all the inserts on the playfield image so there is alpha transparency where they should be.
'   Make sure the playfield material has Opacity Active checkbox checked.
' 2. All the  insert text and/or images that lie over the insert plastic needs to be in its own file with
'    alpha transparency. Many playfields may require finding the original font and remaking the insert text.
'
' To add the inserts:
' 1. Import all the textures (images) and materials from this file that start with the word "Insert" into your Table
'   2. Copy and past the two primitves that make up the insert you want to use. One primitive is for the on state, the other for the off state.
'   3. Align the primitives with the associated insert light. Name the on and off primitives correctly.
'   4. Update the Lampz object array. Follow the example in this file.
'   5. You will need to manually tweak the disable lighting value and material parameters to achielve the effect you want.
'
'
' Quick Reference:  Laying the Inserts ( Tutorial From Iaakki)
' - Each insert consists of two primitives. On and Off primitive. Suggested naming convention is to use lamp number in the name. For example
'   is lamp number is 57, the On primitive is "p57" and the Off primitive is "p57off". This makes it easier to work on script side.
' - When starting from a new table, I'd first select to make few inserts that look quite similar. Lets say there is total of 6 small triangle
'   inserts, 4 yellow and 2 blue ones.
' - Import the insert on/off images from the image manager and the vpx materials used from the sample project first, and those should appear
'   selected properly in the primitive settings when you paste your actual insert trays in your target table . Then open up your target project
'   at same time as the sample project and use copy&paste to copy desired inserts to target project.
' - There are quite many parameters in primitive that affect a lot how they will look. I wouldn't mess too much with them. Use Size options to
'   scale the insert properly into PF hole. Some insert primitives may have incorrect pivot point, which means that changing the depth, you may
'   also need to alter the Z-position too.
' - Once you have the first insert in place, wire it up in the script (detailed in part 3 below). Then set the light bulb's intensity to zero,
'   so it won't harass the adjustment.
' - Start up the game with F6 and visually see if the On-primitive blinks properly. If it is too dim, hit D and open editor. Write:
' - p57.BlendDisableLighting = 300 and hit enter
' - -> The insert should appear differently. Find good looking brightness level. Not too bright as the light bulb is still missing. Just generic good light.
'  - If you cannot find proper light color or "mood", you can also fiddle with primitive material values. Provided material should be
'    quite ok for most of the cases.
'  - Now when you have found proper DL value (165), but that into script:
'  - Lampz.Callback(57) = " DisableLighting p57, 165,"
' - That one insert is now adjusted and you should be able to copy&paste rest of the triangle inserts in place and name them correctly. And add them
'   into script. And fine tune their brightness and color.
'
' Light bulbs and ball reflection:
'
' - This kind of lighted primitives are not giving you ball reflections. Also some more glow vould be needed to make the insert to bloom correctly.
' - Take the original lamp (l57), set the bulb mode enabled, set Halo Height to -3 (something that is inside the 2 insert primitives). I'd start with
'   falloff 100, falloff Power 2-2.5, Intensity 10, scale mesh 10, Transmit 5.
' - Start the game with F6, throw a ball on it and move the ball near the blinking insert. Visually see how the reflection looks.
' - Hit D once the reflection is the highest. Open light editor and start fine tuning the bulb values to achieve realistic look for the reflection.
' - Falloff Power value is the one that will affect reflection creatly. The higher the power value is, the brighter the reflection on the ball is.
'   This is the reason why falloff is rather large and falloff power is quite low. Change scale mesh if reflection is too small or large.
' - Transmit value can bring nice bloom for the insert, but it may also affect to other primitives nearby. Sometimes one need to set transmit to
'   zero to avoid affecting surrounding plastics. If you really need to have higher transmit value, you may set Disable Lighting From Below to 1
'   in surrounding primitive. This may remove the problem, but can make the primitive look worse too.

'******************************************************
'*****   END 3D INSERTS
'******************************************************


'******************************************************
'******  FLUPPER BUMPERS
'******************************************************
' Based on FlupperBumpers 0.145 final

' Explanation of how these bumpers work:
' There are 10 elements involved per bumper:
' - the shadow of the bumper ( a vpx flasher object)
' - the bumper skirt (primitive)
' - the bumperbase (primitive)
' - a vpx light which colors everything you can see through the bumpertop
' - the bulb (primitive)
' - another vpx light which lights up everything around the bumper
' - the bumpertop (primitive)
' - the VPX bumper object
' - the bumper screws (primitive)
' - the bulb highlight VPX flasher object
' All elements have a special name with the number of the bumper at the end, this is necessary for the fading routine and the initialisation.
' For the bulb and the bumpertop there is a unique material as well per bumpertop.
' To use these bumpers you have to first copy all 10 elements to your table.
' Also export the textures (images) with names that start with "Flbumper" and "Flhighlight" and materials with names that start with "bumper".
' Make sure that all the ten objects are aligned on center, if possible with the exact same x,y coordinates
' After that copy the script (below); also copy the BumperTimer vpx object to your table
' Every bumper needs to be initialised with the FlInitBumper command, see example below;
' Colors available are red, white, blue, orange, yellow, green, purple and blacklight.
' In a GI subroutine you can then call set the bumperlight intensity with the "FlBumperFadeTarget(nr) = value" command
' where nr is the number of the bumper, value is between 0 (off) and 1 (full on) (so you can also use 0.3 0.4 etc).

' Notes:
' - There is only one color for the disk; you can photoshop it to a different color
' - The bumpertops are angle independent up to a degree; my estimate is -45 to + 45 degrees horizontally, 0 (topview) to 70-80 degrees (frontview)
' - I built in correction for the day-night slider; this might not work perfectly, depending on your table lighting
' - These elements, textures and materials do NOT integrate with any of the lighting routines I have seen in use in many VPX tables
'   (just find the GI handling routine and insert the FlBumperFadeTarget statement)
' - If you want to use VPX native bumperdisks just copy my bumperdisk but make it invisible

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0
  DNA45 = (NightDay - 10) / 20
  DNA90 = 0
  DayNightAdjust = 0.4
Else
  DNA30 = (NightDay - 10) / 30
  DNA45 = (NightDay - 10) / 45
  DNA90 = (NightDay - 10) / 90
  DayNightAdjust = NightDay / 25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt
For cnt = 1 To 6
  FlBumperActive(cnt) = False
Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight
FlInitBumper 1, "white"
FlInitBumper 2, "white"
FlInitBumper 3, "white"
FlInitBumper 4, "red"
FlInitBumper 5, "red"
FlInitBumper 6, "red"
'FlInitBumper 2, "white"
'FlInitBumper 3, "blue"
'FlInitBumper 4, "orange"
'FlInitBumper 5, "yellow"

' ### uncomment the statement below to change the color for all bumpers ###
'   Dim ind
'   For ind = 1 To 5
'    FlInitBumper ind, "green"
'   Next

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True

  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1
  FlBumperFadeTarget(nr) = 1.1
  FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr)
  FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
  Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
' Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
' Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
' FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
  FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)

  ' set the color for the two VPX lights
  Select Case col
    Case "red"
    FlBumperSmallLight(nr).color = RGB(255,4,0)
    FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
    FlBumperBigLight(nr).color = RGB(255,32,0)
    FlBumperBigLight(nr).colorfull = RGB(255,32,0)
    FlBumperHighlight(nr).color = RGB(255,32,0)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
    FlBumperSmallLight(nr).TransmissionScale = 0

    Case "blue"
    FlBumperBigLight(nr).color = RGB(32,80,255)
    FlBumperBigLight(nr).colorfull = RGB(32,80,255)
    FlBumperSmallLight(nr).color = RGB(0,80,255)
    FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
    FlBumperSmallLight(nr).TransmissionScale = 0
    MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
    FlBumperHighlight(nr).color = RGB(255,16,8)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "green"
    FlBumperSmallLight(nr).color = RGB(8,255,8)
    FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
    FlBumperBigLight(nr).color = RGB(32,255,32)
    FlBumperBigLight(nr).colorfull = RGB(32,255,32)
    FlBumperHighlight(nr).color = RGB(255,32,255)
    MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
    FlBumperSmallLight(nr).TransmissionScale = 0.005
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "orange"
    FlBumperHighlight(nr).color = RGB(255,130,255)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    FlBumperSmallLight(nr).TransmissionScale = 0
    FlBumperSmallLight(nr).color = RGB(255,130,0)
    FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
    FlBumperBigLight(nr).color = RGB(255,190,8)
    FlBumperBigLight(nr).colorfull = RGB(255,190,8)

    Case "white"
    FlBumperBigLight(nr).color = RGB(255,230,190)
    FlBumperBigLight(nr).colorfull = RGB(255,230,190)
    FlBumperHighlight(nr).color = RGB(255,180,100)
    FlBumperSmallLight(nr).TransmissionScale = 0
    FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99

    Case "blacklight"
    FlBumperBigLight(nr).color = RGB(32,32,255)
    FlBumperBigLight(nr).colorfull = RGB(32,32,255)
    FlBumperHighlight(nr).color = RGB(48,8,255)
    FlBumperSmallLight(nr).TransmissionScale = 0
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "yellow"
    FlBumperSmallLight(nr).color = RGB(255,230,4)
    FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
    FlBumperBigLight(nr).color = RGB(255,240,50)
    FlBumperBigLight(nr).colorfull = RGB(255,240,50)
    FlBumperHighlight(nr).color = RGB(255,255,220)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    FlBumperSmallLight(nr).TransmissionScale = 0

    Case "purple"
    FlBumperBigLight(nr).color = RGB(80,32,255)
    FlBumperBigLight(nr).colorfull = RGB(80,32,255)
    FlBumperSmallLight(nr).color = RGB(80,32,255)
    FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
    FlBumperSmallLight(nr).TransmissionScale = 0
    FlBumperHighlight(nr).color = RGB(32,64,255)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1
  End Select
End Sub

'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************



'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 To tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim b
  '   Dim BOT
  '   BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz >  - 7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
        End If
      End If
    End If

    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If gBOT(b).Z > 30 Then
        BallShadowA(b).height = gBOT(b).z - BallSize / 4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height = 0.1
      End If
      BallShadowA(b).Y = gBOT(b).Y + offsetY
      BallShadowA(b).X = gBOT(b).X + offsetX
      BallShadowA(b).visible = 1
    End If
  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


'***************************************************************
' ZSHA: VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx8" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
' * Count from 0 up, with at least as many objects each as there can be balls, including locked balls.  You'll get an "eval" warning if tnob is higher
' * Warning:  If merging with another system (JP's ballrolling), you may need to check tnob math and add an extra BallShadowA# flasher (out of range error)
' Ensure you have a timer with a -1 interval that is always running
' Set plastic ramps DB to *less* than the ambient shadows (-11000) if you want to see the pf shadow through the ramp
' Place triggers at the start of each ramp *type* (solid, clear, wire) and one at the end if it doesn't return to the base pf
' * These can share duties as triggers for RampRolling sounds

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
' It's recommended that you be selective in which lights go in this collection, as there are limitations:
' 1. The shadows can "pass through" solid objects and other light sources, so be mindful of where the lights would actually able to cast shadows
' 2. If there are more than two equidistant sources, the shadows can suddenly switch on and off, so places like top and bottom lanes need attention
' 3. At this time the shadows get the light on/off from tracking gilvl, so if you have lights you want shadows for that are on at different times you will need to either:
' a) remove this restriction (shadows think lights are always On)
' b) come up with a custom solution (see TZ example in script)
' After confirming the shadows work in general, use ball control to move around and look for any weird behavior

'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Example RampShadow trigger subs:

'Sub ClearRampStart_hit()
' bsRampOnClear     'Shadow on ramp and pf below
'End Sub

'Sub SolidRampStart_hit()
' bsRampOn        'Shadow on ramp only
'End Sub

'Sub WireRampStart_hit()
' bsRampOnWire      'Shadow only on pf
'End Sub

'Sub RampEnd_hit()
' bsRampOff ActiveBall.ID 'Back to default shadow behavior
'End Sub


'' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  if a > b then
    min = b
  Else
    min = a
  end if
end Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******

' *** These define the appearance of shadows in your table  ***

'Ambient (Room light source)
Const AmbientBSFactor = 0.9  '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 5        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

'Dynamic (Table light sources)
Const DynamicBSFactor = 0.90  '0 To 1, higher is darker
Const Wideness = 20      'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness = 5        'Sets minimum as ball moves away from source

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objrtx1(5), objrtx2(5)
Dim objBallShadow(5)
Dim OnPF(5)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3)
Dim DSSources(30), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

' *** The Shadow Dictionary
Dim bsDict
Set bsDict = New cvpmDictionary
Const bsNone = "None"
Const bsWire = "Wire"
Const bsRamp = "Ramp"
Const bsRampClear = "Clear"

'Initialization
DynamicBSInit

Sub DynamicBSInit()
  Dim iii, source

  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = 1 + iii / 1000 + 0.01  'Separate z for layering without clipping
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = 1 + iii / 1000 + 0.02
    objrtx2(iii).visible = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100 * AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source In DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    '   If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1  'Adapted for TZ with GI left / GI right
    iii = iii + 1
  Next
  numberofsources = iii
End Sub

Sub BallOnPlayfieldNow(onPlayfield, ballNum)  'Whether a ball is currently on the playfield. Only update certain things once, save some cycles
  If onPlayfield Then
    OnPF(ballNum) = True
    bsRampOff gBOT(ballNum).ID
    '   debug.print "Back on PF"
    UpdateMaterial objBallShadow(ballNum).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(ballNum).size_x = 5
    objBallShadow(ballNum).size_y = 4.5
    objBallShadow(ballNum).visible = 1
    BallShadowA(ballNum).visible = 0
    BallShadowA(ballNum).Opacity = 100 * AmbientBSFactor
  Else
    OnPF(ballNum) = False
    '   debug.print "Leaving PF"
  End If
End Sub

Sub DynamicBSUpdate
  Dim falloff 'Max distance to light sources, can be changed dynamically if you have a reason
  falloff = 150
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
  Dim bsRampType
  '   Dim gBOT: gBOT=getballs 'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls

  'Hide shadow of deleted balls
  For s = UBound(gBOT) + 1 To tnob - 1
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(gBOT) < lob Then Exit Sub 'No balls in play, exit

  'The Magic happens now
  For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow
    'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your Elseif segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else (under 20)

    'Primitive shadow on playfield, flasher shadow in ramps
    If AmbientBallShadowOn = 1 Then
      '** Above the playfield
      If gBOT(s).Z > 30 Then
        If OnPF(s) Then BallOnPlayfieldNow False, s   'One-time update
        bsRampType = getBsRampType(gBOT(s).id)
        '   debug.print bsRampType

        If Not bsRampType = bsRamp Then 'Primitive visible on PF
          objBallShadow(s).visible = 1
          objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
          objBallShadow(s).Y = gBOT(s).Y + offsetY
          objBallShadow(s).size_x = 5 * ((gBOT(s).Z + BallSize) / 80) 'Shadow gets larger and more diffuse as it moves up
          objBallShadow(s).size_y = 4.5 * ((gBOT(s).Z + BallSize) / 80)
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (30 / (gBOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else 'Opaque, no primitive below
          objBallShadow(s).visible = 0
        End If

        If bsRampType = bsRampClear Or bsRampType = bsRamp Then 'Flasher visible on opaque ramp
          BallShadowA(s).visible = 1
          BallShadowA(s).X = gBOT(s).X + offsetX
          BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
          BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
          If bsRampType = bsRampClear Then BallShadowA(s).Opacity = 50 * AmbientBSFactor
        ElseIf bsRampType = bsWire Or bsRampType = bsNone Then 'Turn it off on wires or falling out of a ramp
          BallShadowA(s).visible = 0
        End If

        '** On pf, primitive only
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then
        If Not OnPF(s) Then BallOnPlayfieldNow True, s
        objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        objBallShadow(s).Y = gBOT(s).Y + offsetY
        '   objBallShadow(s).Z = gBOT(s).Z + s/1000 + 0.04    'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

        '** Under pf, flasher shadow only
      Else
        If OnPF(s) Then BallOnPlayfieldNow False, s
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If

      'Flasher shadow everywhere
    ElseIf AmbientBallShadowOn = 2 Then
      If gBOT(s).Z > 30 Then 'In a ramp
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then 'On pf
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = 1.04 + s / 1000
      Else 'Under pf
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If
    End If

    ' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If gBOT(s).Z < 30 And gBOT(s).X < 850 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
        dist1 = falloff
        dist2 = falloff
        For iii = 0 To numberofsources - 1 'Search the 2 nearest influencing lights
          LSd = Distance(gBOT(s).x, gBOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then
            '   If LSd < dist2 And ((DSGISide(iii) = 0 And Lampz.State(100)>0) Or (DSGISide(iii) = 1 And Lampz.State(104)>0)) Then  'Adapted for TZ with GI left / GI right
            dist2 = dist1
            dist1 = LSd
            src2 = src1
            src1 = iii
          End If
        Next
        ShadowOpacity1 = 0
        If dist1 < falloff Then
          objrtx1(s).visible = 1
          objrtx1(s).X = gBOT(s).X
          objrtx1(s).Y = gBOT(s).Y
          '   objrtx1(s).Z = gBOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity1 = 1 - dist1 / falloff
          objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
          UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx1(s).visible = 0
        End If
        ShadowOpacity2 = 0
        If dist2 < falloff Then
          objrtx2(s).visible = 1
          objrtx2(s).X = gBOT(s).X
          objrtx2(s).Y = gBOT(s).Y + offsetY
          '   objrtx2(s).Z = gBOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity2 = 1 - dist2 / falloff
          objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
          UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx2(s).visible = 0
        End If
        If AmbientBallShadowOn = 1 Then
          'Fades the ambient shadow (primitive only) when it's close to a light
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
        End If
      Else 'Hide dynamic shadows everywhere else, just in case
        objrtx2(s).visible = 0
        objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub

' *** Ramp type definitions

Sub bsRampOnWire()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsWire
  Else
    bsDict.Add ActiveBall.ID, bsWire
  End If
End Sub

Sub bsRampOn()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRamp
  Else
    bsDict.Add ActiveBall.ID, bsRamp
  End If
End Sub

Sub bsRampOnClear()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRampClear
  Else
    bsDict.Add ActiveBall.ID, bsRampClear
  End If
End Sub

Sub bsRampOff(idx)
  If bsDict.Exists(idx) Then
    bsDict.Item(idx) = bsNone
  End If
End Sub

Function getBsRampType(id)
  Dim retValue
  If bsDict.Exists(id) Then
    retValue = bsDict.Item(id)
  Else
    retValue = bsNone
  End If
  getBsRampType = retValue
End Function

'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************



'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'     * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'     * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'     * Create a Timer called RampRoll, that is enabled, with a interval of 100
'     * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'     * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'     * To stop tracking ball
'        * call WireRampOff
'        * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

Sub WireRampOn(input)
  Waddball ActiveBall, input
  RampRollUpdate
End Sub

Sub WireRampOff()
  WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  Dim x
  For x = 1 To UBound(RampBalls)  'Check, don't add balls twice
    If RampBalls(x, 1) = input.id Then
      If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 To UBound(RampBalls)
    If IsEmpty(RampBalls(x, 1)) Then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      Exit Sub
    End If
    If x = UBound(RampBalls) Then  'debug
      Debug.print "WireRampOn error, ball queue Is full: " & vbNewLine & _
      RampBalls(0, 0) & vbNewLine & _
      TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
      TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
      TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
      TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
      TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
      " "
    End If
  Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
  '   Debug.Print "In WRemoveBall() + Remove ball from loop array"
  Dim ballcount
  ballcount = 0
  Dim x
  For x = 1 To UBound(RampBalls)
    If ID = RampBalls(x, 1) Then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
  Next
  If BallCount = 0 Then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
  RampRollUpdate
End Sub

Sub RampRollUpdate()  'Timer update
  Dim x
  For x = 1 To UBound(RampBalls)
    If Not IsEmpty(RampBalls(x,1) ) Then
      If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
        If RampType(x) Then
          PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      End If
      If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
  Next
  If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
  "1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
  "2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
  "3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
  "4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
  "5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
  "6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
  " "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
  BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************


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
SpinnerSoundLevel = 0.5      'volume level; range [0, 1]

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
'                     Fleep  Supporting Ball & Sound Functions
' *********************************************************************

'Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^5) 'was 10
    Else
        AudioFade = Csng(-((- tmp) ^5) ) 'was 10
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^5) 'was 10
    Else
        AudioPan = Csng(-((- tmp) ^5) ) 'was 10
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
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, Bumper1
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

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
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
  FlipperLeftHitParm = parm / 10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm / 10
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
  PlaySoundAtLevelActiveBall ("Metal_Touch_"& Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
  RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////

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

Sub Apron01_Hit (idx)
  If Abs(activeball.velx) < 4 and activeball.vely > 7 then
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
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Primitive33
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  FlipperCradleCollision ball1, ball2, velocity

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

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05    'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
  End Select
End Sub


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'************************* DJROBX Bump Hits *********************************************************

Dim NextOrbitHit:NextOrbitHit = 0

Sub WireRampBumps_Hit(idx)
    if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
        RandomBump3 .3, Pitch(ActiveBall)+5
        ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
        ' Lowering these numbers allow more closely-spaced clunks.
        NextOrbitHit = Timer + .2 + (Rnd * .2)
    end if
End Sub

Sub PlasticRampBumps_Hit(idx)
    if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
        RandomBump 2, -20000
        ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
        ' Lowering these numbers allow more closely-spaced clunks.
        NextOrbitHit = Timer + .1 + (Rnd * .2)
    end if
End Sub


Sub MetalGuideBumps_Hit(idx)
    if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
        RandomBump2 2, Pitch(ActiveBall)
        ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
        ' Lowering these numbers allow more closely-spaced clunks.
        NextOrbitHit = Timer + .2 + (Rnd * .2)
    end if
End Sub

Sub MetalWallBumps_Hit(idx)
    if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
        RandomBump 2, 20000 'Increased pitch to simulate metal wall
        ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
        ' Lowering these numbers allow more closely-spaced clunks.
        NextOrbitHit = Timer + .2 + (Rnd * .2)
    end if
End Sub


'' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
    dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
        PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires metalguidebump1 to 2 in Sound Manager
Sub RandomBump2(voladj, freq)
    dim BumpSnd:BumpSnd= "metalguidebump" & CStr(Int(Rnd*2)+1)
        PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires WireRampBump1 to 5 in Sound Manager
Sub RandomBump3(voladj, freq)
    dim BumpSnd:BumpSnd= "WireRampBump" & CStr(Int(Rnd*5)+1)
        PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Stop Bump Sounds
Sub BumpSTOP1_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOP2_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOP3_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub


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



'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************


''*******************************************
''  ZOPT: User Options
''*******************************************

' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
Sub Table1_OptionEvent(ByVal eventId)

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    RampRollVolume = Table1.Option("Ramp Volume", 0, 1, 0.01, 0.5, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  DynamicBallShadowsOn = Table1.Option("Dynamic Ball Shadows", 0, 1, 1, 1, 0, Array("Off", "On"))
  AmbientBallShadowOn = Table1.Option("Ambient Ball Shadow", 0, 2, 1, 1, 0, Array("Off", "Moving Ball Shadow", "Flasher Image Shadow"))
  StagedFlippers = Table1.Option("Staged Flippers", 0, 1, 1, 0, 0, Array("Off", "On"))

End Sub


' Lava Lamp code below.  Thank you STEELY!
Dim Lspeed(7), Lbob(7), blobAng(7), blobRad(7), blobSiz(7), Blob, Bcnt   ' for VR Lava Lamp

Bcnt = 0
For Each Blob in Lspeed
  Lbob(Bcnt) = .2
  Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05
  Bcnt = Bcnt + 1
Next

Sub LavaTimer_Timer()
  Bcnt = 0
  For Each Blob in Lava
    If Blob.TransZ <= VRLavaBase.Size_Z * 1.5 Then  'Change blob direction to up
      Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05   'travel speed
      blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
      blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center, radius
      blobSiz(Bcnt) = Int((150 * Rnd) + 100)    'blob size
      Blob.Size_x = blobSiz(Bcnt):Blob.Size_y = blobSiz(Bcnt):Blob.Size_z = blobSiz(Bcnt)
      Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaBase.X  'place blob
      Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaBase.Y
    End If

    If Blob.TransZ => VRLavaBase.Size_Z*5 Then    'Change blob direction to down
      blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
      blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center,radius
      Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaBase.X  'place blob
      Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaBase.Y
      Lspeed(Bcnt) = Int((8 * Rnd) + 5) * .05:Lspeed(Bcnt) = Lspeed(Bcnt) * -1        'travel speed
    End If

    ' Make blob wobble
    If Blob.Size_x > blobSiz(Bcnt) + blobSiz(Bcnt)*.15 or Blob.Size_x < blobSiz(Bcnt) - blobSiz(Bcnt)*.15  Then Lbob(Bcnt) = Lbob(Bcnt) * -1
    Blob.Size_x = Blob.Size_x + Lbob(Bcnt)
    Blob.Size_y = Blob.Size_y + Lbob(Bcnt)
    Blob.Size_z = Blob.Size_Z - Lbob(Bcnt) * .66
    Blob.TransZ = Blob.TransZ + Lspeed(Bcnt)    'Move blob
    Bcnt = Bcnt + 1
  Next
End Sub

' Make bubbles random speed
Sub BeerTimer_Timer()

  Randomize(21)
  VRBeerBubble1.z = VRBeerBubble1.z + Rnd(1)*0.5
  if VRBeerBubble1.z > -771 then VRBeerBubble1.z = -955

  VRBeerBubble2.z = VRBeerBubble2.z + Rnd(1)*1
  if VRBeerBubble2.z > -768 then VRBeerBubble2.z = -955

  VRBeerBubble3.z = VRBeerBubble3.z + Rnd(1)*1
  if VRBeerBubble3.z > -768 then VRBeerBubble3.z = -955

  VRBeerBubble4.z = VRBeerBubble4.z + Rnd(1)*0.75
  if VRBeerBubble4.z > -774 then VRBeerBubble4.z = -955

  VRBeerBubble5.z = VRBeerBubble5.z + Rnd(1)*1
  if VRBeerBubble5.z > -771 then VRBeerBubble5.z = -955

  VRBeerBubble6.z = VRBeerBubble6.z + Rnd(1)*1
  if VRBeerBubble6.z > -774 then VRBeerBubble6.z = -955

  VRBeerBubble7.z = VRBeerBubble7.z + Rnd(1)*0.8
  if VRBeerBubble7.z > -768 then VRBeerBubble7.z = -955

  VRBeerBubble8.z = VRBeerBubble8.z + Rnd(1)*1
  if VRBeerBubble8.z > -771 then VRBeerBubble8.z = -955

End Sub

' VR Plunger Animation
Sub TimerPlunger_Timer
  If PinCab_Plunger.Y < 54.31711  then
    PinCab_Plunger.Y = PinCab_Plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  PinCab_Plunger.Y = -45.68289 + (5 * Plunger.Position) - 20
End Sub


' Culture neutral string to double conversion (handles situation where you don't know how the string was written)
Function CNCDbl(str)
    Dim strt, Sep, i
    If IsNumeric(str) Then
        CNCDbl = CDbl(str)
    Else
        Sep = Mid(CStr(0.5), 2, 1)
        Select Case Sep
        Case "."
            i = InStr(1, str, ",")
        Case ","
            i = InStr(1, str, ".")
        End Select
        If i = 0 Then
            CNCDbl = Empty
        Else
            strt = Mid(str, 1, i - 1) & Sep & Mid(str, i + 1)
            If IsNumeric(strt) Then
                CNCDbl = CDbl(strt)
            Else
                CNCDbl = Empty
            End If
        End If
    End If
End Function


'*******************************
'   VPW Revisions
'*******************************

' 001 - Multiple Authors - Many revisions prior to tracking
' 002 - iaakki - cleanup and reorder layers
' 003 - iaakki - all lamps disabled and on and off inserts moved to separate layers. Just to fix bugs in them
' 004 - iaakki - readjusted all insert material
' 005 - iaakki - bulbs back on
' 006 - iaakki - adjusted bulbs a bit and fixed bugs
' 007 - benji - LUT adjuster
' 008 - iaakki - Flasherblooms
' 009 - Benji - Fleep Sounds
' 010 - Benji - Physics Adjustments
' 011 - iaakki - fixed right flip angle, flippernudge and live catch. Hide wall54, Wall55 as they didn't fit there imo.
' 014 - Bord - New playfield mesh with new working subway for scoop
' 015 - Benji - reduced ball rolling sound, reduced flipper strength, reduced environmental emission from 5 to 1 (current lighting configuration is ideal for this dimmer 'evening' default)
' 016 - mcarter78 - updated nFozzy physics code, added endpoint primitives to slingshots
' 017 - mcarter78 - fix rubbers alignment, replace pegs & posts, code fixes for physics & sounds
' 018 - mcarter78 - fix inlane wire guides length, more code fixes, replace pegs & posts walls with prims, add missing pegs and screws, replace drop sound kickers with triggers
' 019 - mcarter78 - add missing insert text, fix right orbit return angle, add remaining triggers for ramp sounds, add in-game menu
' 020 - mcarter78 - add more missing pegs & screws, fix misaligned switches & posts, fix script typo
' 021 - TastyWasps - VR Room added
' 022 - iaakki - Lampz added, Flupper domes code updated, cor.update logic added, GI update added, Flash125 example flasher added
' 023 - mcarter78 - Added remaining flasher subs, fixed collections so sounds work, fixed surfaces for ramp sound triggers, added new arrow plastic scan
' 024 - mcarter78 - Physical trough & kickers conversion, added missing sw44, repositioned flasher31
' 025 - mcarter78 - Add arrow bracket, add missing slingshot pegs
' 026 - TastyWasps - VR Cab art and VR Backglass replaced with Hauntfreaks images
' 027 - iaakki - wireramp primitives and bakes added(needs webp convert), plastic ramps are pretty much deleted atm. Will add later. Metal walls also in works..
' 028 - iaakki - metal wall primitives and bakes
' 029 - iaakki - plastic ramp primitives and bakes
' 030 - mcarter78 - fix position for upper post rubbers
' 031 - iaakki - baked side blades
' 032 - TastyWasps - Updated in-game menu for current features.  More will be needed as features added.
' 033 - mcarter78 - Audio fixes for walls, ramps & apron.  Reposition ramp triggers
' 034 - iaakki - plastics and bezels added etc. Depth Bias somewhat fiddled..
' 035 - Sixtoe - old plastics turned off, plastics_on raised, plastic bezels turned down
' 036 - iaakki - fixed bezel DB and material, some weird sideblade walls deleted
' 037 - iaakki - bumper rework
' 038 - iaakki - arrow plastic and its bracket done. Lane guide metals baked, some lighting tuning
' 039 - iaakki - GI bake for PF, GI rework, New ramp bakes and adjustments, some lighting changes here and there
' 040 - mcarter78 - Add sounds for GI/Flasher relays, slow ball down when dropping from right ramp
' 041 - mcarter78 - replace Did You Know plastic
' 042 - mcarter78 - Fix sounds for autoplunge & knocker, add random wire ramp stop sounds, fix slingshot endpoints, add correct flupper dome prims for flasher 1 & 2
' 043 - mcarter78 - Fix sounds for knocker, missing posts, gates backsides, missing metals, spinners, adjust metal walls hit threshold
' 044 - mcarter78 - Add catcher for narnia ball in varitarget, place ball in nearby vuk if lost - fix visible rubber position below arrow
' 045 - TastyWasps - Lightened up VR Room a bit with new materials.
' 046 - mcarter78 - New arrow plastic bake
' 047 - mcarter78 - Add call to FlipperCradleCollision, Update ball rolling sounds, Add dynamic ball shadows
' 048 - mcarter78 - Transp posts/pegs material adjustment, add shadow to upper flipper, add parts library prims for outlane posts, nuts & screws, threaded bolts inside clear posts+
' 049 - mcarter78 - imported new standup prims & implemented roth standups, removed flipper endpoint prims, disabled knocker callback to prevent knock on insert coin
' 050 - mcarter78 - Extend wall by upper bumpers to prevent stuck ball, adjust plunger window position
' 051 - PT5K - Added Staged Flipper Option
' 052 - mcarter78 - fix plunger wall angle for full plunge, remove debug shot tester
' 053 - TastyWasps - Desktop left rails made visible, default Desktop POV given more incline for realism
' 054 - TastyWasps - Brightened up the VR Room one more time to an ambient balance.
' 055 - mcarter78 - Add new flipper prims, fix flipper triggers size, increase nudge sensitivity, remove manual flipper sol calls
' 056 - mcarter78 - Add wall beneath playfield for flipper issues, bumper screws material, reduce ball reflection, add lockdown bar for DT (thanks Cliffy!)
' 057 - mcarter78 - Flashers overhaul for spinners, shrunken head, temple, did you know plastic
' 058 - iaakki - reworked backpanel bulbs, top lane insert fixes, shrunken head GI lights, back panel inserts added, lock insert fix, MetalGuide hit thresholds increased 0.5 -> 1.5, varitarget primitive updated, inlane gradle corners edited
' 059 - mcarter78 - Add new varitarget code (with help from rothbauerw), increase scoop kicker strength
' 060 - mcarter78 - Increase flipper strength, EOSTnew.  Decrease scoop kicker strength, upper bumper lighting.  Keep flashers inside the cabinet. Remove extraneous collidable posts, move collidables to their own layer. Fix upper flipper shadow.
' 061 - mcarter78 - Added ramp decals
' 062 - iaakki - shot tester, digits fix, reworked shrunken head magnet, slight edit to ramp decals, flasher fixes
' 063 - mcarter78 - more flippers & scoop kicker tweaks, ramp end primitives, more missing screws & nuts, layers organization, remove more extra walls
' RC1 - mcarter78 - reduce left bumpers lamp intensity, migrate options to tweakUI
' RC2 - Sixtoe - Added Hauntfreaks playfield image mod to equalise the wood colour, Shrunk the pop bumpers as they were *far* too large, added collidable top inlane guides, realigned the bent playfield_mesh and improved the hole, numerous tweaks
' RC3 - iaakki - varitarget pf image modified, upper flip shadow depth bias fixed, some images removed, arrow bake image scaled 4k to 2k.
' RC4 - mcarter78 - new pf image, open up scoop & varitarget shots a bit, extend rubber at upper bumpers
' RC5 - mcarter78 - new screw prims for ramp starts, updated DT POV, a bit more varitarget shot tuning
' RC6 - mcarter78 - fix dSleeves hit events, remove redundant calls to DoVTAnim & Cor.Update
' RC7 - Swallowed Up
' RC8 - Sixtoe - Significant cleanups of script and table, removal of redundant objects and assets, unifying many things, changing inlane wires to prims, split scoop prim to brighten metal inside scoop, added laneguides to red ramp area, fixed broken decorative transparent pins above plastics, widened the scoop entrance as it was too narrow editing the playfield mesh and underplayfield scoop
' RC9 - Sixtoe - Rebuild wire ramp exit, more credits
' RC10 - mcarter78 - slight re-tune of physics for slings / varitarget / plunger / scoop kicker.  Add ball shadows & staged flippers to F12 menu.
' RC11 - mcarter78 - new desktop background image, courtesy of passion4pins
' RC12 - Wylte - Object positions locked, Ballshadow fix, f26a & f30a insert flashers given code, f22a & f31a created, added visible opto light under idol
' RC13 - Sixtoe - Flashers & Lamps changes, renamed everything so it now makes sense.
' RC14 - mcarter78 - fix SetLamp calls, convert f30 & f25 to insert flashers also, more varitarget tweaks, add hole to pf for idol magnet lamp and resize.
' RC15 - Sixtoe - Flasher adjustments, varitarget/ramp entrance plastic adjusted, got tiki led working properly.
' RC16 - mcarter78 - Adjust varitarget vuk size, hit height & accuracy
' RC17 - Sixtoe - Rebuilt varitarget area physically.
' RC18 - mcarter78 - move varitarget stop, adjust varitarget spring strength from .01 to .5
' RC19 - iaakki - fixed 2 domes that were not working, adjusting dome flasher light objects, added sling posts to gi lighting
' RC20 - apophis - added PWM flasher capability (Optional).
' RC21 - mcarter78 - make flasherlight2a & 3a work with PWM flashers, uncomment flasher 130 lampz assign
' RC22 - mcarter78 - fadeup/fadedown & increased intensity for PWM Flupper dome flashers
' RC23 - mcarter78 - loosen up that varitarget spring just a beeeeeeet
' RC24 - Sixtoe - Rebuilt varitarget wire ramp, removed old ramp objects (one which was partially obstructing the right orbit)
' RC25 - Sixtoe - Made new primitive collidable ramp for top middle ramp, removed duplicate old ramp.
' RC26 - Sixtoe - Rebuilt right ramp, adjusted wire arch ramp
' RC27 - iaakki - Rebuild visible right ramp and bake, rework for right side flasher, plastics, screws, spacer and something for the wirereamp still
' RC28 - Sixtoe - Adjusted plastic, fixtures and fittings, removed some duplicated prims, simplified some prims, dropped left plastic so it's under the ramp not through it.
' RC29 - mcarter78 - Disable PWM flashers, slight tweaks to shrunken head flasher fadeup/fadedown
' RC30 - mcarter78 - Re-enable PWM lights, disable lamp fading code
' RC31 - mcarter78 - Adjustments for flashers intensity & fading, inserts intensity & red bumpers highlight color
' RC32 - mcarter78 - GI callback adjustments, don't play relay sounds unless full on or full off, lower bloom a bit, fix inserts & bumper initial light states, remove wayward screw
' RC33 - Sixtoe - fixed solonoid wrong assignments
' RC34 - mcarter78 - fixed dim inserts and insert flashers lighting
' RC35 - mcarter78 - convert GI & inserts to use modsol, flashers WIP (TODO: Finish fixing reflections)
' RC36 - mcarter78 - Revert back to non-PWM lamps
' RC37 - mcarter78 - Replace TargetBouncer sub with modified code from Blood Machines
' RC38 - mcarter78 - null check on ball before playing varitarget sound
' RC39 - Sixtoe - Stripped out PWM flasher code, renamed existing flashed, fixed broken and incorrectly assigned flashers and flasher objects, reconnected opto LED, replaced flasherblooms
' RC40 - Sixtoe - Varitarget audio fix thanks to RothbauerW
' RC41 - iaakki - standup target changes (reverted)
' RC42 - mcarter78 - Revert to RC40, Fix standup targets orientation, Move pegs from dSleeves to dPosts
' RC43 - mcarter78 - Fix names and materials for pegs that were mistakenly set as sleeves
' RC44 - iaakki - removed TargetBouncer from SDT's and just changed their elasticity. fixed Flasherlight3a and Flasherlight2a
' RC45 - mcarter78 - Fix VR backglass brightness, add UpdateBallBrightness
' v1.0 Release!
'
' 1.0.1 - apophis - New PWM lamp and flasher support added (not optional). Updated CheckLiveCatch.
' 1.0.2 - Sixtoe - Changed rotation and position of gate3 as it was wrong, removed reflection from arrow_bulb, removed f28a insert light (as it doesn't have an insert!)
' 1.0.3 - apophis - Removed unnecessary UpdateLamps sub.
' 1.0.4 - apophis - Updated VPMLampUpdater class to include an Init sub.
' 1.0.5 - apophis - Added some randomness to sw29 kick out.
' 1.0.6 - apophis - Table info updated. Rules card and screenshot added.
' v1.1 Release
' 1.1.1 - Sixtoe - Altered gate 3 as it was in the wrong place, thanks to feedback from "dielated" on VPU.
' v1.2 Release
' v1.2.1 - Sixtoe - Moved multiball exit wall, rebuilt ramp end, randomised multiball kicker, removed redundant assets, edited playfield image
' v1.3 Release

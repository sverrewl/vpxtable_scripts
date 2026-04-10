Option Explicit
Randomize
'
' DDDDDDDD    EEEEEEEE    AAAAA    DDDDDDDD   LL      YY      YY       WW     WW  EEEEEEEE    AAAAA    PPPPPPPP    OOOOOOO   NN      NN
' DD     DD  EE          AA   AA   DD     DD  LL       YY    YY        WW     WW  EE         AA   AA   PP     PP  OO     OO  NNN     NN
' DD     DD  EE         AA     AA  DD     DD  LL        YY  YY         WW     WW  EE        AA     AA  PP     PP  OO     OO  NNNN    NN
' DD     DD  EE         AA     AA  DD     DD  LL         YYYY          WW     WW  EE        AA     AA  PP     PP  OO     OO  NN NN   NN
' DD     DD  EEEEEEEE   AAAAAAAAA  DD     DD  LL          YY           WW     WW  EEEEEEEE  AAAAAAAAA  PPPPPPPP   OO     OO  NN  NN  NN
' DD     DD  EE         AA     AA  DD     DD  LL          YY           WW     WW  EE        AA     AA  PP         OO     OO  NN   NN NN
' DD     DD  EE         AA     AA  DD     DD  LL          YY           WW  W  WW  EE        AA     AA  PP         OO     OO  NN    NNNN
' DD     DD  EE         AA     AA  DD     DD  LL          YY           WW  W  WW  EE        AA     AA  PP         OO     OO  NN     NNN
' DDDDDDDD    EEEEEEEE  AA     AA  DDDDDDDD   LLLLLLLL    YY            WWWWWWW   EEEEEEEE  AA     AA  PP          OOOOOOO   NN      NN
'
' by Gottlieb (Premier Technology), 1990
'
'
'
' Release Notes:
'
' 2.0
'   - Schlabber34 - Renders, materials and new Apron
'   - DGrimmreaper - VR cabinet
'   - DaRdog - VR Mega room
'   - Burger - nFozzy and Fleep - GI prim swap
'   - studly_do_right - in depth testing of fleep, VR, and a lot more
'   - Gravy - VR testing
'
'Credits from 1.0 release:
'
'   - Herweh original scripting
'   - 32assassin and BodyDump for bringing this Table to VPX. I used their table as a base in v1.0.
'   - Bord and BorgDog for images, testing and all kinds of help on the original table!
'   - BorgDog moving spinner rod
'   - RothbauerW for some code of his manuel trough in 'No Good Gofers'
'   - Last, and definetly not least the VPX devs!! Thanks guys!
'
'
' ****************************************************
' OPTIONS
' ****************************************************
Dim EnableAdditionalGI, ShowBallShadow, EnableGI
Dim GIForceMode
Dim PlasticsMode


' SHOW/HIDE the ball shadow
' 0 = ball shadow are hidden
' 1 = ball shadow are visible (default)
ShowBallShadow = 1


' ***************************************************************************************************************************************************************


'----- VR Room Auto-Detect -----
Dim VRRoomChoice : VRRoomChoice = 2         '1 - Cab Only, 2 - MEGA, 3 - Minimal
Dim VRTest : VRTest = False

Dim TableWidth  : TableWidth  = Table1.Width
Dim TableHeight : TableHeight = Table1.Height

Const Angle = 20
Const ReflipAngle = 20


' ****************************************************
' standard definitions
' ****************************************************
Const cGameName   = "deadweap"
Const UseSolenoids  = 2
Const UseLamps    = 1
Const UseGI     = 0
Const UseSync     = 0
Const SSolenoidOn = "SolOn"
Const SSolenoidOff  = "SolOff"
Const SCoin     = "coin"

Const BallSize    = 50
Const BallMass    = 1

If Version < 10600 Then
  MsgBox "This table requires Visual Pinball 10.6 Revision 3696 or newer!" & vbNewLine & "Your version: " & Replace(Version/1000,",","."), , "Deadly Weapon VPX"
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package."
On Error Goto 0

LoadVPM "01210000","GTS3.VBS",3.10

Dim DesktopMode: DesktopMode = Table1.ShowDT

' Used to enable/disable flippers based on tilt and ball status
Dim FlippersEnabled : FlippersEnabled = False

' ENABLE/DISABLE manual ball control for debugging
Dim EnableManualBallControl : EnableManualBallControl = 0

' ENABLE/DISABLE air ball mode
Dim AirBall : AirBall = 1


' ****************************************************
' table init
' ****************************************************
Sub Table1_Init
  With Controller
    On Error Resume Next
        .GameName = cGameName
        If Err Then MsgBox "Can't start game " & cGameName & vbNewLine & Err.Description : Exit Sub
    On Error Goto 0
    .SplashInfoLine = "Deadly Weapon (Gottlieb)"
    .HandleMechanics= 0
    .HandleKeyboard = 0
    .ShowDMDOnly  = 1
    .ShowFrame    = 0
    .ShowTitle    = 0
        .Hidden     = 1
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

  InitLights InsertLights
  If VRRoom = 1 Then
    setup_backglass
  End If
  InitTable
End Sub

Sub Table1_Paused()   : Controller.Pause = True  : End Sub
Sub Table1_UnPaused() : Controller.Pause = False : End Sub
Sub Table1_Exit()   : Controller.Stop      : End Sub

Sub InitTable()
  ' some timer inits
  PinMAMETimer.Interval = 1
  PinMAMETimer.Enabled  = True

  ' nudging
  vpmNudge.TiltSwitch = 151
  vpmNudge.Sensitivity= 5
  vpmNudge.TiltObj  = Array(Bumper10, Bumper11, Bumper12, Bumper13, RightSlingShot, sw14)

    ' trough (do not use the standard trough as it does not work with this Gottlieb table)
  RightSlot.CreateSizedballWithMass Ballsize/2,Ballmass
  MiddleSlot.CreateSizedballWithMass Ballsize/2,Ballmass
  LeftSlot.CreateSizedballWithMass Ballsize/2,Ballmass
  Controller.Switch(26) = True
  ' no ball in drain (Drain.CreateSizedballWithMass Ballsize/2,Ballmass)
  Controller.Switch(16) = False

  ' kickback
  LeftPlunger.Pullback

  ' side rails
  If RenderingMode <> 2 and VRTest = 0 Then
  Ramp15.Visible = DesktopMode
  Ramp16.Visible = DesktopMode
  End If
  ' maybe no GI
  If EnableGI = 0 Then SetGI False, False
BlinkGI
LastPlasticsChoice = PlasticsChoice
End Sub


' ****************************************************
' solenoids - increase by 1 because first solenoid is 0
' ****************************************************
SolCallback(5)    = "SolKickingTarget"
SolCallback(7)    = "SolKickback"
'SolCallback(7)    = "vpmSolAutoPlunger LeftPlunger,3,"
SolCallBack(8)    = "SolResetDropTargets"
SolCallback(9)    = "SolLeftKicker"
SolCallback(10)   = "SolRightKicker"

SolCallback(26)   = "SolGI"
SolCallback(28)   = "SolBallRelease"
SolCallback(29)   = "SolDrain"
SolCallback(30)   = "vpmSolSound SoundFX(""Knocker_1"",DOFKnocker),"
SolCallback(31)   = "SolTilt"
'SolCallback(32)   = "" ' do not add the callback to these solenoid as the Tilt mechanism gets crazy

' playfield flashers (11 to 18)
SolCallback(11)   = "SolFlasherMultiball"
SolCallback(12)   = "SolFlasherSpecial"
SolCallback(13)   = "SolFlasherGas"
SolCallback(14)   = "SolFlasherAlley"
SolCallback(15)   = "SolFlasherGrill"
SolCallback(16)   = "SolFlasherBank"
SolCallback(17)   = "SolFlasherMall"
SolCallback(18)   = "SolFlasherSchool"

' backboard flashers (19 - 25)
SolCallback(19)   = "SolFlasher19"
SolCallback(20)   = "SolFlasher20"
SolCallback(21)   = "SolFlasher21"
SolCallback(22)   = "SolFlasher22"
SolCallback(23)   = "SolFlasher23"
SolCallback(24)   = "SolFlasher24"


Sub SolTilt(Enabled)
  FlippersEnabled = Not Enabled
  SetGI FlippersEnabled, False
  SolLFlipper False
  SolRFlipper False
End Sub


' ******************************************************
' manual trough, based on code from nfozzy and rothbauerw
' ******************************************************
Sub LeftSlot_Hit()     : Controller.Switch(26) = True  : UpdateTrough : End Sub
Sub LeftSlot_UnHit()   : Controller.Switch(26) = False : UpdateTrough : End Sub
Sub MiddleSlot_UnHit() : UpdateTrough : End Sub
Sub RightSlot_UnHit()  : UpdateTrough : End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled  = True
End Sub

Sub UpdateTroughTimer_Timer()
  If RightSlot.BallCntOver = 0 Then MiddleSlot.Kick 60, 9
  If MiddleSlot.BallCntOver = 0 Then LeftSlot.Kick 60, 9
  UpdateTroughTimer.Enabled  = False
End Sub


' ******************************************************
' drain and ball release
' ******************************************************
Sub Drain_Hit()
  RandomSoundDrain Drain
  'PlaySoundAt "drain", Drain
  UpdateTrough
  Controller.Switch(16) = True
End Sub
Sub Drain_UnHit()
  Controller.Switch(16) = False
End Sub

Sub SolDrain(Enabled)
  If Enabled Then Drain.Kick 60,20
End Sub

Sub SolBallRelease(Enabled)
  If Enabled Then
    RandomSoundBallRelease RightSlot
    'PlaySoundAt SoundFX("Ballrelease",DOFContactors), RightSlot
    RightSlot.Kick 60, 7
    UpdateTrough
  End If
End Sub


' ****************************************************
' additional GI
' ****************************************************
Dim DrainGIOffSteps : DrainGIOffSteps = 0
Sub DrainGIOffTimer_Timer()
  If EnableAdditionalGI <> 1 Then DrainGIOffTimer.Enabled = False
  If isGIOn And DrainGIOffSteps > 1 And IsGameOver Then SetGI False, False
  DrainGIOffSteps = DrainGIOffSteps + 1
  If Not isGIOn And (DrainGIOffSteps >= 8 Or Not IsGameOver) Then
    SetGI True, False
    DrainGIOffTimer.Enabled = False
  End If
End Sub

' additional GI for attract mode
Dim AttractTimerStep : AttractTimerStep = 30 + Int((Rnd()*3+1) * 20)
Sub AttractTimer_Timer()
  If EnableAdditionalGI <> 1 Then AttractTimer.Enabled = False
  If AttractTimer.Interval = 999 Then
    If AttractTimerStep > 0 Then
      AttractTimerStep = AttractTimerStep - 1
    Else
      AttractTimerStep    = 0
      AttractTimer.Interval   = 100
    End If
  Else
    If AttractTimer.Interval = 100 Then
      AttractTimer.Interval   = IIF((AttractTimerStep>=10), 50, 1000 - AttractTimerStep * 100)
      SetGI False, False
    Else
      AttractTimer.Interval   = 100
      SetGI True, False
    End If
    AttractTimerStep = AttractTimerStep + 1
    If AttractTimerStep >= 25-Int(Rnd()*10+1) Then
      AttractTimer.Interval = 999
      AttractTimerStep = 30 + Int((Rnd()*5+1) * 20)
      SetGI True, False
    End If
  End If
End Sub

' additional GI for multiball mode
Dim MultiballTimerStep : MultiballTimerStep = 0
Sub MultiballTimer_Timer()
  If EnableAdditionalGI <> 1 Then MultiballTimer.Enabled = False
  If MultiballTimer.Interval = 1000 Then
    If MultiballTimerStep >= 3 Then
      MultiballTimerStep = 0
      MultiballTimer.Interval = 100
    End If
  Else
    If MultiballTimerStep >= 8 Then
      MultiballTimerStep = 0
      MultiballTimer.Interval = 1000
      SetGI True, False
    Else
      SetGI (Not isGIOn), False
    End If
  End If
  MultiballTimerStep = MultiballTimerStep + 1
End Sub


' ****************************************************
' keycodes
' ****************************************************

Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
'vpmTimer.PulseSwitch 26,0,0
    Plunger.Pullback
    SoundPlungerPull
    TimerVRPlunger.Enabled = True
    TimerVRPlunger2.Enabled = False
    VR_PlungerRod.Y = 0'976.1494'1415.345
    VR_PlungerKnob.Y = 0'976.1494'1415.345
  End if
    If keycode = LeftFlipperKey Then  FlipperActivate LeftFlipper, LFPress : SolLFlipper True:FlipperButtonLeft1.transx=FlipperButtonLeft1.transx + 6: FlipperButtonLeft2.transx=FlipperButtonLeft2.transx + 6
  If keycode = RightFlipperkey Then FlipperActivate RightFlipper, RFPress : SolRFlipper True:FlipperButtonRight1.transx=FlipperButtonRight1.transx - 6: FlipperButtonRight2.transx=FlipperButtonRight2.transx - 6
    If keycode = StartGameKey Then    Controller.Switch(3)  = True:VR_StartButton.transz=VR_StartButton.transz - 4:VR_StartButton1.transz=VR_StartButton1.transz - 4
  If keycode = AddCreditKey Then
'ToggleGI
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
    Controller.Switch(1)  = True
  End If
  If keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  Controller.Switch(2)  = True
  End if
  If keycode = KeySlamDoorHit Then  Controller.Switch(152)  = True
  If keycode = LeftMagnaSave Then
  'ToggleVR
  Controller.Switch(4)  = True
  End If
  If keycode = RightMagnaSave Then  Controller.Switch(5)  = True
  If IsGameOver Then
    If keycode = LeftFlipperKey Then  Controller.Switch(4)  = True
    If keycode = RightFlipperkey Then   Controller.Switch(5)  = True
  End If
  If KeyDownHandler(keycode) Then Exit Sub

  ' manual ball control
  If EnableManualBallControl = 1 Then
    If keycode = 46 Then If contball = 1 Then contball = 0 : BallControlTimer.Enabled = False Else contball = 1 : BallControlTimer.Enabled = True: End If : End If ' C Key
    If keycode = 48 Then If bcboost = 1 Then bcboost = bcboostmulti Else bcboost = 1 : End If : End If 'B Key
    If keycode = 203 Then bcleft = 1        ' Left Arrow
    If keycode = 200 Then bcup = 1          ' Up Arrow
    If keycode = 208 Then bcdown = 1        ' Down Arrow
    If keycode = 205 Then bcright = 1       ' Right Arrow
  End If

  If KeyCode = 56 Then
    VR_TopFiveLeft.transz = VR_TopFiveLeft.transz - 10
  End If

  If Keycode = 184 Then
    VR_TopFiveRight.transz = VR_TopFiveRight.transz - 10
  End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall
    TimerVRPlunger.Enabled = False
    TimerVRPlunger2.Enabled = True
    VR_PlungerRod.Y = 0'976.1494'1415.345
    VR_PlungerKnob.Y = 0'976.1494'1415.345
  End If
    If keycode = LeftFlipperKey Then  FlipperDeActivate LeftFlipper, LFPress : SolLFlipper False: FlipperButtonLeft1.transx=FlipperButtonLeft1.transx - 6: FlipperButtonLeft2.transx=FlipperButtonLeft2.transx - 6
  If keycode = RightFlipperkey Then FlipperDeActivate RightFlipper, RFPress : SolRFlipper False: FlipperButtonRight1.transx=FlipperButtonRight1.transx + 6: FlipperButtonRight2.transx=FlipperButtonRight2.transx + 6
  If keycode = StartGameKey Then    Controller.Switch(3)  = False:VR_StartButton.transz=VR_StartButton.transz + 4:VR_StartButton1.transz=VR_StartButton1.transz + 4
  If keycode = AddCreditKey Then    Controller.Switch(1)  = False
  If keycode = AddCreditKey2 Then   Controller.Switch(2)  = False
  If keycode = KeySlamDoorHit Then  Controller.Switch(152)  = False
  If keycode = LeftMagnaSave Then   Controller.Switch(4)  = False
  If keycode = RightMagnaSave Then  Controller.Switch(5)  = False
  If IsGameOver Then
    If keycode = LeftFlipperKey Then  Controller.Switch(4)  = False
    If keycode = RightFlipperkey Then   Controller.Switch(5)  = False
  End If
  If KeyUpHandler(keycode) Then Exit Sub

  ' manual ball control
  If EnableManualBallControl = 1 Then
    If keycode = 203 Then bcleft = 0        ' Left Arrow
    If keycode = 200 Then bcup = 0          ' Up Arrow
    If keycode = 208 Then bcdown = 0        ' Down Arrow
    If keycode = 205 Then bcright = 0       ' Right Arrow
  End If

  If KeyCode = 56 Then
    VR_TopFiveLeft.transz = VR_TopFiveLeft.transz + 10
  End If

  If Keycode = 184 Then
    VR_TopFiveRight.transz = VR_TopFiveRight.transz + 10
  End If
End Sub

Sub TimerVRPlunger_Timer
  If VR_PlungerRod.Y < 100 then
       VR_PlungerRod.Y = VR_PlungerRod.Y + 5
  End If
  If VR_PlungerKnob.Y < 100 then
       VR_PlungerKnob.Y = VR_PlungerKnob.Y + 5
  End If
End Sub

Sub TimerVRPlunger2_Timer
  'debug.print plunger.position
  VR_PlungerRod.Y = 0 + (5* Plunger.Position) -20'1415.345
  VR_PlungerKnob.Y = 0 + (5* Plunger.Position) -20'1415.345
End Sub

' ****************************************************
' manual ball control
' ****************************************************
Sub StartControl_Hit() : Set ControlBall = ActiveBall : contballinplay = True : End Sub
Sub StopControl_Hit() : contballinplay = false : End Sub

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1         'Do Not Change - default setting
bcvel = 4           'Controls the speed of the ball movement
bcyveloffset = -0.01    'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3      'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControlTimer_Timer()
    If Contball and ContBallInPlay then
        If bcright = 1 Then ControlBall.velx = bcvel*bcboost Else If bcleft = 1 Then ControlBall.velx = - bcvel*bcboost Else ControlBall.velx=0 : End If : End If
        If bcup = 1 Then ControlBall.vely = -bcvel*bcboost Else If bcdown = 1 Then ControlBall.vely = bcvel*bcboost Else ControlBall.vely= bcyveloffset : End If : End If
    End If
End Sub


' ****************************************************
' flippers
' ****************************************************
Sub SolLFlipper(Enabled)
  If Enabled Then
       If FlippersEnabled Then
    LF.Fire
    UpperLeftFlipper.RotateToEnd
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
        End If
  Else
    LeftFlipper.RotateToStart
    UpperLeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel

  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
        If FlippersEnabled Then
    RF.Fire
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
    End if
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub FlipperTimer_Timer()
  PLeftFlipper.RotZ = LeftFlipper.CurrentAngle - 180
  PLeftFlipperOff.RotZ = LeftFlipper.CurrentAngle - 180
  PULeftFlipper.RotZ = UpperLeftFlipper.CurrentAngle - 180
  PULeftFlipperOff.RotZ = UpperLeftFlipper.CurrentAngle - 180
  PRightFlipper.RotZ = RightFlipper.CurrentAngle - 180
  PRightFlipperOff.RotZ = RightFlipper.CurrentAngle - 180

  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperLMSh.RotZ = UpperLeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub


' ****************************************************
' left kick back
' ****************************************************

Sub SolKickBack(Enabled)
    If Enabled Then
       LeftPlunger.Fire
       'PlaySoundAt SoundFX("Popper", DOFContactors), LeftPlunger
    Else
       LeftPlunger.PullBack
    End If
End Sub

'Sub SolKickBack(Enabled)
'    If Enabled Then
'        LeftPlunger.Fire
' '       PlaySoundAt SoundFX("Popper", DOFContactors), LeftPlunger
'        LeftPlunger.PullBack
'    End If
'End Sub


' ****************************************************
' sling shots and animations
' ****************************************************
Dim RightStep

Sub RightSlingShot_Slingshot()
  RS.VelocityCorrect ActiveBall
  vpmTimer.PulseSwitch 15,0,0
  RandomSoundSlingshotRight Sling1
    'PlaySoundAt SoundFX("right_slingshot", DOFContactors), SLING1
    PSlingR1a.Visible = 0
    PSlingR1c.Visible = 1
    sling1.TransZ = -20
    RightStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub
Sub RightSlingShot_Timer()
    Select Case RightStep
        Case 3:PSlingR1c.Visible = 0:PSlingR1b.Visible = 1:sling1.TransZ = -10
        Case 4:PSlingR1b.Visible = 0:PSlingR1a.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RightStep = RightStep + 1
End Sub


' ****************************************************
' bumper
' ****************************************************
Sub Bumper10_Hit : RandomSoundBumperTop Bumper10 : MoveSkirt 10, 1, Bumper10, BumperSkirt10, ActiveBall : End Sub
Sub Bumper11_Hit : RandomSoundBumperMiddle Bumper11 : MoveSkirt 11, 2, Bumper11, BumperSkirt11, ActiveBall : End Sub
Sub Bumper12_Hit : RandomSoundBumperBottom Bumper12 : MoveSkirt 12, 3, Bumper12, BumperSkirt12, ActiveBall : End Sub
Sub Bumper13_Hit : RandomSoundBumperTop Bumper13 : MoveSkirt 13, 4, Bumper13, BumperSkirt13, ActiveBall : End Sub

Sub Bumper10_Timer() : MoveSkirt 10, 1, Bumper10, BumperSkirt10, Nothing : End Sub
Sub Bumper11_Timer() : MoveSkirt 11, 2, Bumper11, BumperSkirt11, Nothing : End Sub
Sub Bumper12_Timer() : MoveSkirt 12, 3, Bumper12, BumperSkirt12, Nothing : End Sub
Sub Bumper13_Timer() : MoveSkirt 13, 4, Bumper13, BumperSkirt13, Nothing : End Sub

ReDim skirtX(3), skirtY(3)
Sub MoveSkirt(id, soundid, bumper, skirt, ball)
  If Not ball Is Nothing Then
    bumper.TimerEnabled   = False
    'PlaySoundAt SoundFX("fx_bumper" & soundid, DOFContactors), bumper
    skirtX(id-10)       = IIF((Abs(skirt.Y-ball.Y)<20),0,IIF((skirt.Y < ball.Y), -1, 1))
    skirtY(id-10)       = IIF((Abs(skirt.X-ball.X)<20),0,IIF((skirt.X < ball.X), 1, -1))
    skirt.RotX        = skirtX(id-10)
    skirt.RotY        = skirtY(id-10)
    vpmTimer.PulseSwitch id,0,0
    bumper.TimerInterval  = 8
    bumper.TimerEnabled   = True
  Else
    ' turn the moving direction
    If Abs(skirt.RotX) >= 4 Or Abs(skirt.RotY) >= 4 Then
      skirtX(id-10)     = skirtX(id-10) * -0.5
      skirtY(id-10)     = skirtY(id-10) * -0.5
    End If
    ' move the skirt
    skirt.RotX        = skirt.RotX + skirtX(id-10) ' vertical, positive moves back down
    skirt.RotY        = skirt.RotY + skirtY(id-10) ' horizontal, positive moves right down
    skirt.TransZ      = Abs(skirt.RotX) + IIF(DesktopMode,1,0)
    ' maybe stop as all action is done
    If Abs(skirt.RotX) = 0 And Abs(skirt.RotY) = 0 Then
      skirt.TransZ    = 0
      bumper.TimerEnabled = False
    End If
  End If
End Sub


' ****************************************************
' slingshot target
' ****************************************************
Sub sw14_Slingshot(): vpmTimer.PulseSwitch 14,0,0 : SlingTarget sw14, sw14P, True : PlaysoundAt "Sling_L2", sw14P : End Sub
Sub sw14_Timer()  : SlingTarget sw14, sw14P, False : End Sub

Dim SlingSteps : SlingSteps = 0
Sub SlingTarget(sw, prim, init)
  If init Then
    sw.TimerEnabled   = False
    SlingSteps      = 0
    prim.RotX       = 0
    sw.TimerInterval  = 2
    sw.TimerEnabled   = True
  Else
    Select Case SlingSteps
      Case 0,1,2,3 : prim.RotX = prim.RotX - 6
      Case 4,5 : prim.RotX = prim.RotX * 2 / 3
      Case 50 : prim.RotX = 0 : sw.TimerEnabled = False
      Case Else : prim.RotX = prim.RotX + 0.25
    End Select
    SlingSteps = SlingSteps + 1
  End If
End Sub


' ****************************************************
' switches
' ****************************************************
' wire triggers
Sub sw22_Hit    : Controller.Switch(22) = True  : End Sub
Sub sw22_UnHit    : Controller.Switch(22) = False : End Sub
Sub sw23_Hit    : Controller.Switch(23) = True  : End Sub
Sub sw23_UnHit    : Controller.Switch(23) = False : End Sub
Sub sw25_Hit    : Controller.Switch(25) = True  : End Sub
Sub sw25_UnHit      : Controller.Switch(25) = False : End Sub
Sub sw32_Hit    : Controller.Switch(32) = True  : End Sub
Sub sw32_UnHit    : Controller.Switch(32) = False : End Sub
Sub sw33_Hit    : Controller.Switch(33) = True  : End Sub
Sub sw33_UnHit    : Controller.Switch(33) = False : End Sub
Sub sw35_Hit    : Controller.Switch(35) = True  : FlippersEnabled = True : End Sub
Sub sw35_UnHit    : Controller.Switch(35) = False : End Sub
Sub sw40_Hit    : Controller.Switch(40) = True  : End Sub
Sub sw40_UnHit    : Controller.Switch(40) = False : End Sub
Sub sw41_Hit    : Controller.Switch(41) = True  : End Sub
Sub sw41_UnHit    : Controller.Switch(41) = False : End Sub
Sub sw42_Hit    : Controller.Switch(42) = True  : End Sub
Sub sw42_UnHit    : Controller.Switch(42) = False : End Sub
Sub sw43_Hit    : Controller.Switch(43) = True  : End Sub
Sub sw43_UnHit    : Controller.Switch(43) = False : End Sub
Sub sw45_Hit    : Controller.Switch(45) = True  : End Sub
Sub sw45_UnHit    : Controller.Switch(45) = False : End Sub

' spinners
Sub Spinner1_Spin() : vpmTimer.PulseSwitch 30,0,0 : SoundSpinner Spinner1 : End Sub
Sub Spinner2_Spin() : vpmTimer.PulseSwitch 31,0,0 : SoundSpinner Spinner2 : End Sub

' drop targets
'Sub sw24_Hit()   :  SoundDropTargetDrop  Rubber2 : DropTarget 24, sw24, sw24P, sw24POff, True, ActiveBall : DoubleDrop1.Collidable = False : End Sub
'Sub sw24_Timer() : DropTarget 24, sw24, sw24P, sw24POff, False, Nothing : End Sub
'Sub DoubleDrop1_Hit : DropTarget 24, sw24, sw24P, sw24POff, True, ActiveBall : DropTarget 34, sw34, sw34P, sw34POff, False, Nothing : AvoidTripleHit 44 : DoubleDrop1.Collidable = False : DoubleDrop2.Collidable = False : End Sub
'Sub sw34_Hit()   : SoundDropTargetDrop  Rubber2 : DropTarget 34, sw34, sw34P, sw34POff, True, ActiveBall : DoubleDrop1.Collidable = False : DoubleDrop2.Collidable = False : End Sub
'Sub sw34_Timer() : DropTarget 34, sw34, sw34P, sw34POff, False, Nothing : End Sub
'Sub DoubleDrop2_Hit : DropTarget 34, sw34, sw34P, sw34POff, True, ActiveBall : DropTarget 44, sw44, sw44P, sw44POff, False, Nothing : AvoidTripleHit 24 : DoubleDrop1.Collidable = False : DoubleDrop2.Collidable = False : End Sub
'Sub sw44_Hit()   : SoundDropTargetDrop  Rubber2 : DropTarget 44, sw44, sw44P, sw44POff, True, ActiveBall : DoubleDrop2.Collidable = False : End Sub
'Sub sw44_Timer() : DropTarget 44, sw44, sw44P, sw44POff, False, Nothing : End Sub

Sub sw24_Hit()    : DropTarget 24, sw24, sw24P, sw24POff, True, ActiveBall : DoubleDrop1.Collidable = False : End Sub
Sub sw24_Timer()  : DropTarget 24, sw24, sw24P, sw24POff, False, Nothing : End Sub
Sub DoubleDrop1_Hit : DropTarget 24, sw24, sw24P, sw24POff, True, ActiveBall : DropTarget 34, sw34, sw34P, sw34POff, False, Nothing : AvoidTripleHit 44 : DoubleDrop1.Collidable = False : DoubleDrop2.Collidable = False : End Sub
Sub sw34_Hit()    : DropTarget 34, sw34, sw34P, sw34POff, True, ActiveBall : DoubleDrop1.Collidable = False : DoubleDrop2.Collidable = False : End Sub
Sub sw34_Timer()  : DropTarget 34, sw34, sw34P, sw34POff, False, Nothing : End Sub
Sub DoubleDrop2_Hit : DropTarget 34, sw34, sw34P, sw34POff, True, ActiveBall : DropTarget 44, sw44, sw44P, sw44POff, False, Nothing : AvoidTripleHit 24 : DoubleDrop1.Collidable = False : DoubleDrop2.Collidable = False : End Sub
Sub sw44_Hit()    : DropTarget 44, sw44, sw44P, sw44POff, True, ActiveBall : DoubleDrop2.Collidable = False : End Sub
Sub sw44_Timer()  : DropTarget 44, sw44, sw44P, sw44POff, False, Nothing : End Sub

Sub SolResetDropTargets(Enabled)
  If Enabled Then
    DropTargetTimer.Enabled = True
  End if
End Sub
Sub DropTargetTimer_Timer()
  DropTargetTimer.Enabled = False
  RandomSoundDropTargetReset sw42
  'PlaySoundAt SoundFX("DTReset",DOFContactors), sw42
  ResetDropTarget 24, sw24, sw24P, sw24POff
  ResetDropTarget 34, sw34, sw34P, sw34POff
  ResetDropTarget 44, sw44, sw44P, sw44POff
  DoubleDrop1.Collidable = True : DoubleDrop2.Collidable = True
End Sub
Sub AvoidTripleHit(id)
  DoubleDrop1.TimerInterval   = 500 + id
  DoubleDrop1.TimerEnabled  = True
End Sub
Sub DoubleDrop1_Timer()
  DoubleDrop1.TimerEnabled  = False
End Sub

' standup targets
Sub sw17_Hit()    : StandUpTarget 17, sw17, sw17P, True, ActiveBall : End Sub
Sub sw17_Timer()  : StandUpTarget 17, sw17, sw17P, False, Nothing : End Sub
Sub sw27_Hit()    : StandUpTarget 27, sw27, sw27P, True, ActiveBall : End Sub
Sub sw27_Timer()  : StandUpTarget 27, sw27, sw27P, False, Nothing : End Sub
Sub sw37_Hit()    : StandUpTarget 37, sw37, sw37P, True, ActiveBall : End Sub
Sub sw37_Timer()  : StandUpTarget 37, sw37, sw37P, False, Nothing : End Sub
Sub sw47_Hit()    : StandUpTarget 47, sw47, sw47P, True, ActiveBall : End Sub
Sub sw47_Timer()  : StandUpTarget 47, sw47, sw47P, False, Nothing : End Sub
Sub sw57_Hit()    : StandUpTarget 57, sw57, sw57P, True, ActiveBall : End Sub
Sub sw57_Timer()  : StandUpTarget 57, sw57, sw57P, False, Nothing : End Sub

Sub SolKickingTarget(Enabled)
  PlaySoundAt SoundFX("slingshot", DOFContactors), sw14P
End Sub

Sub DropTarget(id, sw, prim, primOff, playsnd, actBall)
  If DoubleDrop1.TimerEnabled And id = DoubleDrop1.TimerInterval-500 Then
    sw.TimerEnabled = False
    Exit Sub
  End If
  If Not sw.TimerEnabled Then
    'If Not actBall Is Nothing Then DropTargetHit playsnd, actBall
If Not actBall Is Nothing Then SoundDropTargetDrop actBall
    Controller.Switch(id) = True
    sw.Collidable = False
    sw.TimerInterval = 5
    sw.TimerEnabled = True
    prim.TransX = prim.TransX - 1
    prim.TransY = prim.TransY - 1
    primOff.TransX = primOff.TransX - 1
    primOff.TransY = primOff.TransY - 1
  End If
  prim.Z = prim.Z - 14
  primOff.Z = primOff.Z - 14
  If prim.Z < -45 Then
    sw.TimerEnabled = False
    prim.Z = -50
    primOff.Z = -50
  End If
End Sub
Sub ResetDropTarget(id, sw, prim, primOff)
  sw.TimerEnabled = False
  prim.Z = 0 : prim.TransX = 0 : prim.TransY = 0
  primOff.Z = 0 : primOff.TransX = 0 : primOff.TransY = 0
  Controller.Switch(id) = False
  sw.Collidable = True
End Sub
Sub StandUpTarget(id, sw, prim, isHit, actBall)
  If isHit And sw.TimerEnabled Then Exit Sub
  If Not sw.TimerEnabled Then
    If Not actBall Is Nothing Then StandupTargetHit actBall
    vpmTimer.PulseSwitch id,0,0
    sw.TimerInterval = 5
    sw.TimerEnabled = True
  Else
    prim.TransX = prim.TransX + 0.5
    If prim.TransX >= 2.2 Then
      prim.TransX = 0
      sw.TimerEnabled = False
    End If
  End If
End Sub


' ****************************************************
' saucer
' ****************************************************
Dim sw20Step : sw20Step = 0
Dim sw21Step : sw21Step = 0
Dim sw20Ball : Set sw20Ball = Nothing
Dim sw21Ball : Set sw21Ball = Nothing

Sub sw20_Hit() : PlaySoundAt "fx_saucer_enter", sw20 : End Sub
Sub sw21_Hit() : PlaySoundAt "fx_saucer_enter", sw21 : End Sub

Sub sw20_Unhit() : Controller.Switch(20) = False : End Sub
Sub sw21_Unhit() : Controller.Switch(21) = False : End Sub

Sub SolLeftKicker(Enabled)
  If Enabled Then
    sw20Step      = 0
    MoveHammer sw20Hammer, 0
    sw20.TimerInterval  = 11
    sw20.TimerEnabled   = True
  End If
End Sub
Sub sw20_Timer()
  Select Case sw20Step
    Case 0 :  MoveHammer sw20Hammer, -1 : PlaySoundAt SoundFX("fx_vuk_exit", DOFContactors), sw20 : NoGlassHitTimer.Enabled = True
    Case 1 :  MoveHammer sw20Hammer, 16 : If Controller.Switch(20) Then sw20Ball.VelX = 1.96+(Rnd()*0.5-0.25) : sw20Ball.VelY = 7+(Rnd()-0.5) : sw20Ball.VelZ = 12+(Rnd()*2-1) : End If
    Case 7 :  Controller.Switch(20) = False
    Case 25,26,27,28,29 : MoveHammer sw20Hammer, -2
    Case 30 :   MoveHammer sw20Hammer,  0 : sw20.TimerEnabled = False
    Case Else : ' nothing to do
  End Select
  sw20Step = sw20Step + 1
End Sub
Dim Saucer20MysteryStep : Saucer20MysteryStep = 0
Sub Saucer20MysteryTimer_Timer()
  If isGIOn And Controller.Switch(20) And Saucer20MysteryStep > 1 Then SetGI False, False
  Saucer20MysteryStep = Saucer20MysteryStep + 1
  If Not Controller.Switch(20) Then
    If Not isGIOn Then SetGI True, False
    Saucer20MysteryStep = 0
    Saucer20MysteryTimer.Enabled = False
  End If
End Sub

Sub SolRightKicker(Enabled)
  If Enabled Then
    sw21Step      = 0
    MoveHammer sw21Hammer, 0
    sw21.TimerInterval  = 13
    sw21.TimerEnabled   = True
  End If
End Sub
Sub sw21_Timer()
  Select Case sw21Step
    Case 0 :  MoveHammer sw21Hammer, -1 : PlaySoundAt SoundFX("fx_vuk_exit", DOFContactors), sw21 : NoGlassHitTimer.Enabled = True
    Case 1 :  MoveHammer sw21Hammer, 16 : If Controller.Switch(21) Then sw21Ball.VelX = -2+(Rnd()*0.5-0.25) : sw21Ball.VelY = 7+(Rnd()*0.5-0.25) : sw21Ball.VelZ = 13+(Rnd()*2-1) : End If
    Case 7 :  Controller.Switch(21) = False
    Case 25,26,27,28,29 : MoveHammer sw21Hammer, -2
    Case 30 :   MoveHammer sw21Hammer, 0 : sw21.TimerEnabled = False
    Case Else : ' nothing to do
  End Select
  sw21Step = sw21Step + 1
End Sub
Dim Saucer21MysteryStep : Saucer21MysteryStep = 0
Sub Saucer21MysteryTimer_Timer()
  If isGIOn And Controller.Switch(21) And Saucer21MysteryStep > 1 Then SetGI False, False
  Saucer21MysteryStep = Saucer21MysteryStep + 1
  If Not Controller.Switch(21) Then
    If Not isGIOn Then SetGI True, False
    Saucer21MysteryStep = 0
    Saucer21MysteryTimer.Enabled = False
  End If
End Sub

Sub MoveHammer(hammer, rotZ)
  If rotZ = 0 Then
    hammer.RotZ = 0
  Else
    hammer.RotZ = hammer.RotZ + rotZ
  End If
End Sub


' ****************************************************
' rotating spinner
' ****************************************************
Dim Angle1, Angle2

Sub SpinnerTimer1_Timer()
  Angle1 = (sin (Spinner1.CurrentAngle))
  'TextBox1.text = Spinner1.CurrentAngle
  'TextBox2.text = Angle1
  PSpinner1.RotX = Spinner1.CurrentAngle
    PSpinner1Off.RotX = Spinner1.CurrentAngle
    SpinnerRod1.TransZ = -sin( (Spinner1.CurrentAngle) * (2*3.14/360)) * 5
    SpinnerRod1.TransX = (sin( (Spinner1.CurrentAngle- 90) * (2*3.14/360)) * -5)
' SpinnerRod1.TransZ =  cos ((Spinner1.CurrentAngle)) * -2
' SpinnerRod1.TransY = Angle1 * -2
End Sub

Sub SpinnerTimer2_Timer()
  Angle2 = (sin (Spinner2.CurrentAngle))
  'TextBox3.text = Spinner2.CurrentAngle
  'TextBox4.text = Angle2
  PSpinner2.RotX = Spinner2.CurrentAngle
    PSpinner2Off.RotX = Spinner2.CurrentAngle
    SpinnerRod2.TransZ = -sin( (Spinner2.CurrentAngle) * (2*3.14/360)) * 5
    SpinnerRod2.TransX = (sin( (Spinner2.CurrentAngle- 90) * (2*3.14/360)) * -5)
' SpinnerRod2.TransZ =  cos ((Spinner2.CurrentAngle)) * -2
' SpinnerRod2.TransY = Angle2 * -2
End Sub


' *********************************************************************
' playfield insert lights
' *********************************************************************
Dim PFLights(110,3), PFLightsCount(110), isGIOn

Sub InitLights(aColl)
  ' init inserts
  Dim obj, idx
  For Each obj In aColl
    idx = obj.TimerInterval
    Set PFLights(idx, PFLightsCount(idx)) = obj
    PFLightsCount(idx) = PFLightsCount(idx) + 1
  Next
  ' init flasher
  InitFlasher
  ' init GI
  InitGI
End Sub
Sub InitFlasher()

End Sub
Sub InitGI()
  SetGI True, True
End Sub

' inserts and bumper lights
Sub LampTimer_Timer()
  Dim chgLamp, num, chg, ii, nr, xxx
  xxx = -1
    chgLamp = Controller.ChangedLamps
  If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
      Select Case chglamp(ii,0)
      Case 1
        If chgLamp(ii,1) = 1 Then aBumperValue(0) = -99 Else aBumperValue(0) = 3
      Case 70
        If chgLamp(ii,1) = 1 Then aBumperValue(1) = -99 Else aBumperValue(1) = 3
      Case 75
        If chgLamp(ii,1) = 1 Then aBumperValue(2) = -99 Else aBumperValue(2) = 3
      Case Else
        On Error Resume Next
        For nr = 1 to PFLightsCount(chgLamp(ii,0)) : PFLights(chgLamp(ii,0),nr - 1).State = chgLamp(ii,1) : Next
        On Error GoTo 0
      End Select
        Next
  End If
End Sub


' *********************************************************************
' flasher
' *********************************************************************
Sub SolFlasherMultiball(Enabled)
  FlasherMultiball.State = Abs(Enabled)
End Sub
Sub SolFlasherSpecial(Enabled)
  FlasherSpecial.State = Abs(Enabled)
End Sub
Sub SolFlasherGas(Enabled)
  FlasherGas.State = Abs(Enabled) : FlasherGasA.State = Abs(Enabled) : FlasherGasB.State = Abs(Enabled)
End Sub
Sub SolFlasherAlley(Enabled)
  FlasherAlley.State = Abs(Enabled) : FlasherAlleyA.State = Abs(Enabled) : FlasherAlleyB.State = Abs(Enabled)
End Sub
Sub SolFlasherGrill(Enabled)
  FlasherGrill.State = Abs(Enabled) : FlasherGrillA.State = Abs(Enabled)
End Sub
Sub SolFlasherBank(Enabled)
  FlasherBank.State = Abs(Enabled) : FlasherBankA.State = Abs(Enabled) : FlasherBankB.State = Abs(Enabled)
End Sub
Sub SolFlasherMall(Enabled)
  FlasherMall.State = Abs(Enabled) : FlasherMallA.State = Abs(Enabled) : FlasherMallB.State = Abs(Enabled)
End Sub
Sub SolFlasherSchool(Enabled)
  FlasherSchool.State = Abs(Enabled) : FlasherSchoolA.State = Abs(Enabled) : FlasherSchoolB.State = Abs(Enabled)
End Sub
Sub SolFlasher19(Enabled)
  Flasher19.State = Abs(Enabled)
End Sub
Sub SolFlasher20(Enabled)
  Flasher20.State = Abs(Enabled)
End Sub
Sub SolFlasher21(Enabled)
  Flasher21.State = Abs(Enabled)
End Sub
Sub SolFlasher22(Enabled)
  Flasher22.State = Abs(Enabled)
End Sub
Sub SolFlasher23(Enabled)
  Flasher23.State = Abs(Enabled)
End Sub
Sub SolFlasher24(Enabled)
  Flasher24.State = Abs(Enabled)
End Sub


Dim VR
VR=0  ' 1 for VR plastics, 0 for standard plastics


Dim GILevel : GILevel = 0
Dim HighInserts : HighInserts = False
Const HighInsertsFactor = 2.5
Sub SolGI(Enabled)
  SetGI Not Enabled, False
End Sub
Sub SetGI(Enabled, FastMode)
  If EnableGI = 0 And Not isGIOn Then Exit Sub
  If isGIOn <> Enabled Then
    Dim obj
    If Enabled Then

      ' GI goes on

'standard GiOn
Apron_Gi_On.visible=0
Standard_Upper_GiOn.visible=0
Standard_Lower_GiOn.visible=0
Decals_Lower_GiOn.visible=0
Decals_Upper_GiOn.visible=0
'standard GiOff
Apron_Gi_Off.visible=0
Standard_Upper_GiOff.visible=0
Standard_Lower_GiOff.visible=0
Decals_Lower_GiOff.visible=0
Decals_Upper_GiOff.visible=0
'VR GiOn
VR_Lower_GiOn.visible=0
VR_Upper_GiOn.visible=0
'VR GiOff
VR_Lower_GiOff.visible=0
VR_Upper_GiOff.visible=0


if PlasticsChoice=0 then
'standard GiOn
Apron_Gi_On.visible=1
Standard_Upper_GiOn.visible=1
Standard_Lower_GiOn.visible=1
Decals_Lower_GiOn.visible=1
Decals_Upper_GiOn.visible=1
'standard GiOff
Apron_Gi_Off.visible=0
Standard_Upper_GiOff.visible=0
Standard_Lower_GiOff.visible=0
Decals_Lower_GiOff.visible=0
Decals_Upper_GiOff.visible=0
'VR GiOn
VR_Lower_GiOn.visible=0
VR_Upper_GiOn.visible=0

'VR GiOff
VR_Lower_GiOff.visible=0
VR_Upper_GiOff.visible=0



    Else
        'show VR plastics
'standard GiOn
Standard_Upper_GiOn.visible=0
Standard_Lower_GiOn.visible=0
Decals_Lower_GiOn.visible=0
Decals_Upper_GiOn.visible=0
'standard GiOff
Standard_Upper_GiOff.visible=0
Standard_Lower_GiOff.visible=0
Decals_Lower_GiOff.visible=0
Decals_Upper_GiOff.visible=0
'VR GiOn
Apron_Gi_On.visible=1
VR_Lower_GiOn.visible=1
VR_Upper_GiOn.visible=1
Decals_Lower_GiOn.visible=1
Decals_Upper_GiOn.visible=1
'VR GiOff
Apron_Gi_Off.visible=0
VR_Lower_GiOff.visible=0
VR_Upper_GiOff.visible=0
Decals_Lower_GiOff.visible=0
Decals_Upper_GiOff.visible=0

End if


      For Each obj In GIBumper : obj.Image = "Bumper Base GI-ON" : Next
      For Each obj In GIBumperCaps : If obj.Image = "Bumper Caps GI-OFF" Then obj.Image = "Bumper Caps GI-ON" : End If : Next
      GILevel = IIF(FastMode,4,1)
      If HighInserts Then
        For Each obj In InsertLightsHI : obj.IntensityScale = obj.IntensityScale / HighInsertsFactor : Next
        HighInserts = False
      End If
      If Not isGIOn Then Playsound "fx_relay_on" : DOF 101, DOFOn : End If
    Else
      ' GI goes off

'standard GiOn
Apron_Gi_On.visible=0
Standard_Upper_GiOn.visible=0
Standard_Lower_GiOn.visible=0
Decals_Lower_GiOn.visible=0
Decals_Upper_GiOn.visible=0
'standard GiOff
Apron_Gi_Off.visible=0
Standard_Upper_GiOff.visible=0
Standard_Lower_GiOff.visible=0
Decals_Lower_GiOff.visible=0
Decals_Upper_GiOff.visible=0
'VR GiOn
VR_Lower_GiOn.visible=0
VR_Upper_GiOn.visible=0
'VR GiOff
VR_Lower_GiOff.visible=0
VR_Upper_GiOff.visible=0


if PlasticsChoice=0 then
'standard GiOn
Apron_Gi_On.visible=0
Standard_Upper_GiOn.visible=0
Standard_Lower_GiOn.visible=0
Decals_Lower_GiOn.visible=0
Decals_Upper_GiOn.visible=0
'standard GiOff
Apron_Gi_Off.visible=1
Standard_Upper_GiOff.visible=1
Standard_Lower_GiOff.visible=1
Decals_Lower_GiOff.visible=1
Decals_Upper_GiOff.visible=1
'VR GiOn
VR_Lower_GiOn.visible=0
VR_Upper_GiOn.visible=0
'VR GiOff
VR_Lower_GiOff.visible=0
VR_Upper_GiOff.visible=0


    Else
        'show VR plastics
'standard GiOn
Apron_Gi_On.visible=0
Standard_Upper_GiOn.visible=0
Standard_Lower_GiOn.visible=0
Decals_Lower_GiOn.visible=0
Decals_Upper_GiOn.visible=0
'standard GiOff
Apron_Gi_Off.visible=1
Standard_Upper_GiOff.visible=0
Standard_Lower_GiOff.visible=0
Decals_Lower_GiOff.visible=1
Decals_Upper_GiOff.visible=1
'VR GiOn
VR_Lower_GiOn.visible=0
VR_Upper_GiOn.visible=0

'VR GiOff
VR_Lower_GiOff.visible=1
VR_Upper_GiOff.visible=1


End if

      For Each obj In GIBumper : obj.Image = "Bumper Base GI-OFF" : Next
      For Each obj In GIBumperCaps : If obj.Image = "Bumper Caps GI-ON" Then obj.Image = "Bumper Caps GI-OFF" : End If : Next
      GILevel = IIF(FastMode,-1,-4)
      If Not HighInserts Then
        For Each obj In InsertLightsHI : obj.IntensityScale = obj.IntensityScale * HighInsertsFactor : Next
        HighInserts = True
      End If
      If isGIOn Then Playsound "fx_relay_off" : DOF 101, DOFOff : End If
    End If
    isGIOn = Enabled
    If FastMode Then
      LightsTimer.Enabled = False
      LightsTimer_Timer
      LightsTimer.Enabled = True
    End If
  End If
End Sub


Dim aBumper     : aBumper     = Array(BumperCap10, BumperCap11, BumperCap12, BumperCap13)
Dim aBumperLight  : aBumperLight  = Array(BumperLight10, BumperLight11, BumperLight12, Nothing)
Dim aBumperValue  : aBumperValue  = Array(-99, -99, -99, -99)
Dim aBumperBlend  : aBumperBlend  = Array(-1, -1, -1, -1)
Sub LightsTimer_Timer()
  Dim idx, obj, coll
  ' bumper lights
  For idx = 0 To 3
    If aBumperBlend(idx) = -1 Then
      aBumperBlend(idx) = aBumper(idx).BlendDisableLighting
    End If
    If aBumperValue(idx) > -1 Or aBumperValue(idx) = -99 Then
      If aBumperValue(idx) = -99 Then
        aBumper(idx).Image          = "Bumper Caps Lamp ON"
        aBumper(idx).BlendDisableLighting = aBumperBlend(idx)
        If Not aBumperLight(idx) Is Nothing Then aBumperLight(idx).State = LightStateOn : aBumperLight(idx).IntensityScale = 1 : End If
      ElseIf aBumperValue(idx) = 0 Then
        aBumper(idx).Image          = IIF(isGIOn, "Bumper Caps GI-ON", "Bumper Caps GI-OFF")
        aBumper(idx).BlendDisableLighting = aBumperBlend(idx)
        If Not aBumperLight(idx) Is Nothing Then aBumperLight(idx).State = LightStateOff : aBumperLight(idx).IntensityScale = 0 : End If
      ElseIf aBumperValue(idx) <= 8 Then
        aBumper(idx).BlendDisableLighting = aBumper(idx).BlendDisableLighting / 2
        If Not aBumperLight(idx) Is Nothing Then aBumperLight(idx).IntensityScale = aBumperValue(idx) / 4 : End If
      End If
      aBumperValue(idx) = aBumperValue(idx) - 1
    End If
  Next
  ' GI
  If Not PF_GI_OFF.Visible Then PF_GI_OFF.Visible = True
  If Not PF_Overlay_GI_OFF.Visible Then PF_Overlay_GI_OFF.Visible = True
  If GILevel > 0 And GILevel < 5 Then
    ' GI goes on, from 1 to 4
    For Each obj In GI : obj.IntensityScale = GILevel / 4 : Next
    ' playfield and playfield overlay
    PF_GI_OFF.TopMaterial       = "Playfield" & GILevel
    PF_Overlay_GI_OFF.TopMaterial   = "Playfield" & GILevel
    ' playfield elements
    For Each obj In GIOn : obj.Material = "Prims platt" & IIF(GILevel=4,""," " & 4-GILevel) : Next
    'For Each obj In GIOnTrans : obj.Material = "Prims platt trans" & IIF(GILevel=4,""," " & 4-GILevel) : Next
    For Each obj In GIOnShiny : obj.Material = "Prims Shiny" & IIF(GILevel=4,""," " & 4-GILevel) : Next
    If GILevel = 1 Or GILevel = 4 Then
      'For Each coll In Array(GIOn,GIOnTrans,GIOnShiny) : For Each obj In coll : obj.Visible = True : Next : Next
For Each coll In Array(GIOn,GIOnShiny) : For Each obj In coll : obj.Visible = True : Next : Next
    End If
    If GILevel = 4 Then
      'For Each coll In Array(GIOff,GIOffTrans,GIOffShiny) : For Each obj In coll : obj.Visible = False : Next : Next
For Each coll In Array(GIOff,GIOffShiny) : For Each obj In coll : obj.Visible = False : Next : Next
      For Each obj In GIFlasher : obj.Visible = True : Next
      aBumperValue(3) = -99
    End If
    GILevel = GILevel + 1
  ElseIf GILevel < 0 And GILevel > -5 Then
    ' GI goes off, from -4 to -1
    For Each obj In GI : obj.IntensityScale = (Abs(GILevel)-1) / 4 : Next
    ' playfield and playfield overlay
    PF_GI_OFF.TopMaterial       = "Playfield" & IIF(GILevel=-1,"",Abs(GILevel+1))
    PF_Overlay_GI_OFF.TopMaterial   = "Playfield" & IIF(GILevel=-1,"",Abs(GILevel+1))
    ' playfield elements
    For Each obj In GIOn : obj.Material = "Prims platt " & (GILevel + 5) : Next
    'For Each obj In GIOnTrans : obj.Material = "Prims platt trans " & (GILevel + 5) : Next
    For Each obj In GIOnShiny : obj.Material = "Prims Shiny " & (GILevel + 5) : Next
    If GILevel = -4 Or GILevel = -1 Then
      'For Each coll In Array(GIOff,GIOffTrans,GIOffShiny) : For Each obj In coll : obj.Visible = True : Next : Next
For Each coll In Array(GIOff,GIOffShiny) : For Each obj In coll : obj.Visible = True : Next : Next
      For Each obj In GIFlasher : obj.Visible = False : Next
      aBumperValue(3) = 3
    End If
    If GILevel = -1 Then
      'For Each coll In Array(GIOn,GIOnTrans,GIOnShiny) : For Each obj In coll : obj.Visible = False : Next : Next
For Each coll In Array(GIOn,GIOnShiny) : For Each obj In coll : obj.Visible = False : Next : Next
    End If
    GILevel = GILevel + 1
  End If
End Sub






'**********************************************************************************************************
' digital display
'**********************************************************************************************************
Dim Digits(48)
Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10)  = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11)  = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12)  = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13)  = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14)  = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15)  = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
Digits(16)  = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17)  = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18)  = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19)  = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20)  = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21)  = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22)  = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23)  = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24)  = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25)  = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26)  = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27)  = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28)  = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29)  = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30)  = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31)  = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)
Digits(32)  = Array(c00, c05, c0c, c0d, c08, c01, c06, c0f, c02, c03, c04, c07, c0b, c0a, c09, c0e)
Digits(33)  = Array(c10, c15, c1c, c1d, c18, c11, c16, c1f, c12, c13, c14, c17, c1b, c1a, c19, c1e)
Digits(34)  = Array(c20, c25, c2c, c2d, c28, c21, c26, c2f, c22, c23, c24, c27, c2b, c2a, c29, c2e)
Digits(35)  = Array(c30, c35, c3c, c3d, c38, c31, c36, c3f, c32, c33, c34, c37, c3b, c3a, c39, c3e)
Digits(36)  = Array(c40, c45, c4c, c4d, c48, c41, c46, c4f, c42, c43, c44, c47, c4b, c4a, c49, c4e)
Digits(37)  = Array(c50, c55, c5c, c5d, c58, c51, c56, c5f, c52, c53, c54, c57, c5b, c5a, c59, c5e)
Digits(38)  = Array(c60, c65, c6c, c6d, c68, c61, c66, c6f, c62, c63, c64, c67, c6b, c6a, c69, c6e)
Digits(39)  = Array(c70, c75, c7c, c7d, c78, c71, c76, c7f, c72, c73, c74, c77, c7b, c7a, c79, c7e)
Digits(40)  = Array(LED10, LED11, LED12, LED13, LED14, LED15, LED16)
Digits(41)  = Array(LED20, LED21, LED22, LED23, LED24, LED25, LED26)
Digits(42)  = Array(LED30, LED31, LED32, LED33, LED34, LED35, LED36)
Digits(43)  = Array(LED40, LED41, LED42, LED43, LED44, LED45, LED46)
Digits(44)  = Array(LED50, LED51, LED52, LED53, LED54, LED55, LED56)
Digits(45)  = Array(LED60, LED61, LED62, LED63, LED64, LED65, LED66)
Digits(46)  = Array(LED80, LED81, LED82, LED83, LED84, LED85, LED86)
Digits(47)  = Array(LED90, LED91, LED92, LED93, LED94, LED95, LED96)

Sub DTDisplayTimer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
    If DesktopMode Then
      For ii = 0 To UBound(chgLED)
      num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      If (num < 48) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.State=stat And 1
          chg=chg\2 : stat=stat\2
                Next
      End If
        Next
    End If
  End If
End Sub

'******************************
' Setup Backglass
'******************************

Dim xoff,yoff1,zoff,xrot,zscale, xcen,ycen

Sub setup_backglass()

  xoff = -20
  yoff1 = -5 ' this is where you adjust the forward/backward position for quarters 1-4
  zoff = 759
  xrot = -90

  center_digits()

end sub


Sub center_digits()
  Dim ix, xx, yy, yfact, xfact, xobj

  zscale = 0.0000001

  xcen = (130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

  for ix = 40 to 47
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff1

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

end sub

'********************************************
'              Display Output
'********************************************

dim DisplayColor, DisplayColorG
DisplayColor =  RGB(10,10,255)
DisplayColorG =  RGB(136,0,0)


Sub VRDisplayTimer
  Dim ii, jj, obj, b, x
  Dim ChgLED,num, chg, stat
  ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        if (num < 40) then
          For Each obj In VRDigits(num)
            If chg And 1 Then FadeDisplay obj, stat And 1
            chg=chg\2 : stat=stat\2
          Next
        Else
          For Each obj In VRDigits(num)
            If chg And 1 Then FadeDisplay1 obj, stat And 1
            chg=chg\2 : stat=stat\2
          Next
        End If
        Next
    End If
End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
  Else
    Object.Color = RGB(1,1,10)
  End If
End Sub

Sub FadeDisplay1(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColorG
    Object.Opacity = 12
  Else
    Object.Color = RGB(10,1,1)
    Object.Opacity = 6
  End If
End Sub

Dim VRDigits(48)

VRDigits(0)=Array(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15)
VRDigits(1)=Array(D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30)
VRDigits(2)=Array(D31,D32,D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45)
VRDigits(3)=Array(D46,D47,D48,D49,D50,D51,D52,D53,D54,D55,D56,D57,D58,D59,D60)
VRDigits(4)=Array(D61,D62,D63,D64,D65,D66,D67,D68,D69,D70,D71,D72,D73,D74,D75)
VRDigits(5)=Array(D76,D77,D78,D79,D80,D81,D82,D83,D84,D85,D86,D87,D88,D89,D90)
VRDigits(6)=Array(D91,D92,D93,D94,D95,D96,D97,D98,D99,D100,D101,D102,D103,D104,D105)
VRDigits(7)=Array(D106,D107,D108,D109,D110,D111,D112,D113,D114,D115,D116,D117,D118,D119,D120)
VRDigits(8)=Array(D121,D122,D123,D124,D125,D126,D127,D128,D129,D130,D131,D132,D133,D134,D135)
VRDigits(9)=Array(D136,D137,D138,D139,D140,D141,D142,D143,D144,D145,D146,D147,D148,D149,D150)
VRDigits(10)=Array(D151,D152,D153,D154,D155,D156,D157,D158,D159,D160,D161,D162,D163,D164,D165)
VRDigits(11)=Array(D166,D167,D168,D169,D170,D171,D172,D173,D174,D175,D176,D177,D178,D179,D180)
VRDigits(12)=Array(D181,D182,D183,D184,D185,D186,D187,D188,D189,D190,D191,D192,D193,D194,D195)
VRDigits(13)=Array(D196,D197,D198,D199,D200,D201,D202,D203,D204,D205,D206,D207,D208,D209,D210)
VRDigits(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218,D219,D220,D221,D222,D223,D224,D225)
VRDigits(15)=Array(D226,D227,D228,D229,D230,D231,D232,D233,D234,D235,D236,D237,D238,D239,D240)

VRDigits(16)=Array(D241,D242,D243,D244,D245,D246,D247,D248,D249,D250,D251,D252,D253,D254,D255)
VRDigits(17)=Array(D256,D257,D258,D259,D260,D261,D262,D263,D264,D265,D266,D267,D268,D269,D270)
VRDigits(18)=Array(D271,D272,D273,D274,D275,D276,D277,D278,D279,D280,D281,D282,D283,D284,D285)
VRDigits(19)=Array(D286,D287,D288,D289,D290,D291,D292,D293,D294,D295,D296,D297,D298,D299,D300)
VRDigits(20)=Array(D301,D302,D303,D304,D305,D306,D307,D308,D309,D310,D311,D312,D313,D314,D315)
VRDigits(21)=Array(D316,D317,D318,D319,D320,D321,D322,D323,D324,D325,D326,D327,D328,D329,D330)
VRDigits(22)=Array(D331,D332,D333,D334,D335,D336,D337,D338,D339,D340,D341,D342,D343,D344,D345)
VRDigits(23)=Array(D346,D347,D348,D349,D350,D351,D352,D353,D354,D355,D356,D357,D358,D359,D360)
VRDigits(24)=Array(D361,D362,D363,D364,D365,D366,D367,D368,D369,D370,D371,D372,D373,D374,D375)
VRDigits(25)=Array(D376,D377,D378,D379,D380,D381,D382,D383,D384,D385,D386,D387,D388,D389,D390)
VRDigits(26)=Array(D391,D392,D393,D394,D395,D396,D397,D398,D399,D400,D401,D402,D403,D404,D405)
VRDigits(27)=Array(D406,D407,D408,D409,D410,D411,D412,D413,D414,D415,D416,D417,D418,D419,D420)
VRDigits(28)=Array(D421,D422,D423,D424,D425,D426,D427,D428,D429,D430,D431,D432,D433,D434,D435)
VRDigits(29)=Array(D436,D437,D438,D439,D440,D441,D442,D443,D444,D445,D446,D447,D448,D449,D450)
VRDigits(30)=Array(D451,D452,D453,D454,D455,D456,D457,D458,D459,D460,D461,D462,D463,D464,D465)
VRDigits(31)=Array(D466,D467,D468,D469,D470,D471,D472,D473,D474,D475,D476,D477,D478,D479,D480)

VRDigits(32)=Array(D481,D482,D483,D484,D485,D486,D487,D488,D489,D490,D491,D492,D493,D494,D495)
VRDigits(33)=Array(D496,D497,D498,D499,D500,D501,D502,D503,D504,D505,D506,D507,D508,D509,D510)
VRDigits(34)=Array(D511,D512,D513,D514,D515,D516,D517,D518,D519,D520,D521,D522,D523,D524,D525)
VRDigits(35)=Array(D526,D527,D528,D529,D530,D531,D532,D533,D534,D535,D536,D537,D538,D539,D540)
VRDigits(36)=Array(D541,D542,D543,D544,D545,D546,D547,D548,D549,D550,D551,D552,D553,D554,D555)
VRDigits(37)=Array(D556,D557,D558,D559,D560,D561,D562,D563,D564,D565,D566,D567,D568,D569,D570)
VRDigits(38)=Array(D571,D572,D573,D574,D575,D576,D577,D578,D579,D580,D581,D582,D583,D584,D585)
VRDigits(39)=Array(D586,D587,D588,D589,D590,D591,D592,D593,D594,D595,D596,D597,D598,D599,D600)

' 1st
VRDigits(40) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6,led1x8)
VRDigits(41) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6,led2x8)
' 2nd
VRDigits(42) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6,led3x8)
VRDigits(43) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6,led4x8)
' 3rd
VRDigits(44) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6,led5x8)
VRDigits(45) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6,led6x8)
' 4th
VRDigits(46) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6,led8x8)
VRDigits(47) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6,led9x8)


InitDigits

Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each obj in VRDigits(x)
        obj.height = obj.height - 185
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub


'*****************************************
' ninuzzu's modified ball shadow
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3)

Sub BallShadowUpdate_Timer()
  If ShowBallShadow = 1 Then
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of drained balls
    If LeftSlot.BallCntOver = 1 Then BallShadow1.Visible = False
    If MiddleSlot.BallCntOver = 1 Then BallShadow2.Visible = False
    If RightSlot.BallCntOver = 1 Then BallShadow3.Visible = False
'   If UBound(BOT)<(tnob-1) Then
'     For b = (UBound(BOT) + 1) to (tnob-1)
'       BallShadow(b).Visible = False
'     Next
'   End If
    ' exit the sub if no balls on the table
    If IsGameOver Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
      If BOT(b).X < Table1.Width/2 Then
        BallShadow(b).X = ((BOT(b).X) - (Ballsize/40) + ((BOT(b).X - (Table1.Width/2))/40)) + 1
      Else
        BallShadow(b).X = ((BOT(b).X) + (Ballsize/40) + ((BOT(b).X - (Table1.Width/2))/40)) - 1
      End If
      BallShadow(b).Y = BOT(b).Y + 30
      BallShadow(b).Z = BOT(b).Z + 1
      BallShadow(b).Image = IIF((BOT(b).Z < 100), "BallShadow1", "BallShadow2")
      BallShadow(b).Visible = (BOT(b).Z > 20)
    Next
  End If
End Sub


' *********************************************************************
' some special physics behaviour
' *********************************************************************
' target or rubber post is hit so let the ball jump a bit
Sub DropTargetHit(playsnd, actBall)
  'If playsnd Then DropTargetSound actBall
  TargetHit actBall
End Sub

Sub StandupTargetHit(actBall)
  StandUpTargetSound actBall
  TargetHit actBall
End Sub

Sub TargetHit(actBall)
  If AirBall <> 1 Then Exit Sub
  Dim bVel : bVel = BallVel(actBall)
  If bVel > 9 Then
    With actBall
      .VelX = .VelX*14/15 : .VelY = .VelY*14/15 : .VelZ = (1.3+Rnd()*5)*(Sqr(bVel)*0.365)
    End With
  End If
End Sub

Sub RubberPostHit(actBall)
  If AirBall <> 1 Then Exit Sub
  Dim bVel : bVel = BallVel(actBall)
  If bVel > 9 Then
    With actBall
      .VelX = .VelX*14/15 : .VelY = .VelY*14/15 : .VelZ = (1.1+Rnd()*6)*(Sqr(bVel)*0.385)
    End With
  End If
End Sub

Sub RubberRingHit(actBall)
  If AirBall <> 1 Then Exit Sub
  Dim bVel : bVel = BallVel(actBall)
  If bVel > 9 Then
    actBall.VelZ = Sqr(bVel)*0.38
  End If
End Sub


' *********************************************************************
' sound stuff and realtime sounds
' *********************************************************************
Sub RollOverSound(actBall)
  'PlaySoundAtVolPitch SoundFX("rollover",DOFContactors), actBall, 0.02, .25
End Sub
Sub DropTargetSound(actBall)
  PlaySoundAtVolPitch SoundFX("DTDrop",DOFTargets), actBall, 2, .25
End Sub
Sub StandUpTargetSound(actBall)
  PlaySoundAtVolPitch SoundFX("target",DOFTargets), actBall, 2, .25
End Sub
Sub GateSound(actBall)
  PlaySoundAtVolPitch SoundFX("fx_gate",DOFContactors), actBall, 0.02, .25
End Sub

Sub Pins_Hit(idx)
  PlaySoundAtBall "pinhit_low"
  RubberPostHit ActiveBall
End Sub

'Sub Posts_Hit(idx)
'   Dim finalspeed
'   finalspeed = SQR(ActiveBall.VelX ^ 2 + ActiveBall.VelY ^ 2)
' If finalspeed > 16 then
'   PlaySoundAtBall "fx_rubber2"
' End if
' If finalspeed >= 6 AND finalspeed <= 16 then
'     PlaySoundAtBall "rubber_hit_" & (Int(Rnd*3)+1)
' End If
' RubberPostHit ActiveBall
'End Sub


' *********************************************************************
' Supporting Surround Sound Feedback (SSF) functions
' *********************************************************************
' set position as table object (Use object or light but NOT wall) and Vol to 1
' set position as table object and Vol manually.

' set position as table object and Vol + RndPitch manually
Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

' set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

' requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound
Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0, 0, 1, 1, AudioFade(tableobj)
End Sub


' *********************************************************************
' Supporting Ball & Sound Functions
' *********************************************************************

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function



Sub NoBallDropTimer_Timer()
  NoBallDropTimer.Enabled = False
End Sub
Sub NoGlassHitTimer_Timer()
  NoGlassHitTimer.Enabled = False
End Sub


' *********************************************************************
' some general stuff
' *********************************************************************

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order


Function InCircle(pX, pY, centerX, centerY, radius)
  Dim route
  route = Sqr((pX - centerX) ^ 2 + (pY - centerY) ^ 2)
  InCircle = (route < radius)
End Function

Const fileName = "C:\Games\Visual Pinball\User\Log.txt"
Sub WriteLine(text)
  Dim fso, fil
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set fil = fso.OpenTextFile(fileName, 8, true)
  fil.WriteLine Date & " " & Time & ": " & text
  fil.Close
  Set fso = Nothing
End Sub

Function IsGameOver()
  IsGameOver = (LeftSlot.BallCntOver = 1 And MiddleSlot.BallCntOver = 1 And RightSlot.BallCntOver = 1)
End Function

Function IsMultiball()
  IsMultiball = (LeftSlot.BallCntOver + MiddleSlot.BallCntOver + RightSlot.BallCntOver <= 1)
End Function


'**************************************************
'        Flipper Collision Subs
'NOTE: COpy and overwrite collision sound from original collision subs over
'RandomSoundFlipper()' below
'**************************************************'

Sub UpperLeftFlipper_Collide(parm)
  LeftFlipperCollide parm
End Sub

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub



'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

' Put all the Post and Pin objects in dPosts collection. Make sure dPosts fires hit events.
Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

' This collection contains the bottom sling posts. They are not in the dPosts collection so that the TargetBouncer is not applied to them, but they should still have dampening applied
' If you experience airballs with posts or targets, consider adding them to this collection
Sub NoTargetBouncer_Hit
    RubbersD.dampen ActiveBall
End Sub

' Put all the Sleeve objects in dSleeves collection. Make sure dSleeves fires hit events.
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
    'If coef > 1 Then coef = 1
    aBall.velx = aBall.velx * coef
    aBall.vely = aBall.vely * coef
    aBall.velz = aBall.velz * coef
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
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7-1.0

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

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


'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'  Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
' - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
'
' Tutorial videos by Apophis
' Audio : Adding Fleep Part 1         https://youtu.be/rG35JVHxtx4?si=zdN9W4cZWEyXbOz_
' Audio : Adding Fleep Part 2         https://youtu.be/dk110pWMxGo?si=2iGMImXXZ0SFKVCh
' Audio : Adding Fleep Part 3         https://youtu.be/ESXWGJZY_EI?si=6D20E2nUM-xAw7xy


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1        'volume level; range [0, 1]
NudgeRightSoundLevel = 1        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1        'volume level; range [0, 1]
StartButtonSoundLevel = 0.1      'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1        'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010    'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635    'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0            'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45          'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel    'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel    'sound helper; not configurable
SlingshotSoundLevel = 0.95            'volume level; range [0, 1]
BumperSoundFactor = 4.25            'volume multiplier; must not be zero
KnockerSoundLevel = 1              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
TargetSoundFactor = 0.0095 * 10  'volume multiplier; must not be zero
DTSoundLevel = 0.25        'volume multiplier; must not be zero
RolloverSoundLevel = 0.25      'volume level; range [0, 1]
SpinnerSoundLevel = 0.5      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8          'volume level; range [0, 1]
BallReleaseSoundLevel = 1        'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015  'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5      'volume multiplier; must not be zero

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
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / tableheight - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

  If tmp > 0 Then
    AudioFade = CSng(tmp ^ 10)
  Else
    AudioFade = CSng( - (( - tmp) ^ 10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / tablewidth - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

  If tmp > 0 Then
    AudioPan = CSng(tmp ^ 10)
  Else
    AudioPan = CSng( - (( - tmp) ^ 10) )
  End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
  RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
  RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall2()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, kickback
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
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), 0.1, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperLeft(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Left_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
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
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 5 Then
    RandomSoundRubberStrong 1
  End If
  If finalspeed <= 5 Then
    RandomSoundRubberWeak()
  End If
End Sub

Sub Posts_Hit(idx)
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 5 Then
    RandomSoundRubberStrong 1
  End If
  If finalspeed <= 5 Then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd * 10) + 1
    Case 1
      PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 2
      PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 3
      PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 4
      PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 5
      PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 6
      PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 7
      PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 8
      PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 9
      PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 10
      PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 5) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 4) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 3) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
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
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 10 Then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft ActiveBall
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd * 9) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd * 5) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
  If ActiveBall.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If ActiveBall.velx <  - 8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If ActiveBall.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If ActiveBall.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1
      PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)

  FlipperCradleCollision ball1, ball2, velocity

  Dim snd
  Select Case Int(Rnd * 7) + 1
    Case 1
      snd = "Ball_Collide_1"
    Case 2
      snd = "Ball_Collide_2"
    Case 3
      snd = "Ball_Collide_3"
    Case 4
      snd = "Ball_Collide_4"
    Case 5
      snd = "Ball_Collide_5"
    Case 6
      snd = "Ball_Collide_6"
    Case 7
      snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER   ////////////////////////////

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


'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************

'**********************************
'   ZMAT: General Math Functions
'**********************************
' These get used throughout the script.

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    End If
  ElseIf dx = 0 Then
    If dy = 0 Then
      Atn2 = 0
    Else
      Atn2 = Sgn(dy) * pi / 2
    End If
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

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a > b Then
    min = b
  Else
    min = a
  End If
End Function

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

Function RotPoint(x,y,angle)
  Dim rx, ry
  rx = x * dCos(angle) - y * dSin(angle)
  ry = x * dSin(angle) + y * dCos(angle)
  RotPoint = Array(rx,ry)
End Function


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim VolumeDial : VolumeDial = 1             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim PlasticsChoice
Dim LastPlasticsChoice
EnableAdditionalGI = 0
EnableGI = 1

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

    ' Sound volumes
Dim NewPlasticsChoice

If RenderingMode = 2 Or VRTest = True Then
    NewPlasticsChoice = Table1.Option( _
        "Plastics Visual Style", _
        0, 1, 1, 1, 0, _
        Array("Standard", "Optimized for VR") _
    )
Else
    NewPlasticsChoice = Table1.Option( _
        "Plastics Visual Style", _
        0, 1, 1, 0, 0, _
        Array("Standard", "Optimized for VR") _
    )
End If

' detect actual user toggle
If NewPlasticsChoice <> PlasticsChoice Then
    PlasticsChoice = NewPlasticsChoice
    BlinkGI
Else
    PlasticsChoice = NewPlasticsChoice
End If


    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 1, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

    ' VRRoom
  VRRoomChoice = Table1.Option("VR Room", 1, 3, 1, 2, 0, Array("Cab Only", "MEGA", "Minimal"))
  LoadVRRoom

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub



'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

'Dim LS: Set LS = New SlingshotCorrection
Dim RS: Set RS = New SlingshotCorrection
'Dim TS: Set TS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  'LS.Object = LeftSlingshot
  'LS.EndPoint1 = EndPoint1LS
  'LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'TS.Object = TopSlingshot
  'TS.EndPoint1 = EndPoint1TS
  'TS.EndPoint2 = EndPoint2TS

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
  a = Array(RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

' The following sub are needed, however they exist in the ZMAT maths section of the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
' dim rx, ry
' rx = x*dCos(angle) - y*dSin(angle)
' ry = x*dSin(angle) + y*dCos(angle)
' RotPoint = Array(rx,ry)
'End Function

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

Dim gBot
Sub GameTimer_Timer()
  'gBOT = GetBalls
  Cor.Update
  RollingUpdate
End Sub

Dim FrameTime, InitFrameTime
InitFrameTime = 0

Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  If VRRoom > 0 or VRTest Then
    VRDisplayTimer
  End If

  If VRRoom = 0 Then
    DTDisplayTimer
  End If
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
'CorTimer.Interval = 10
'Sub CorTimer_Timer(): Cor.Update: End Sub


'--- Add this near the top of your script ---
Function IIf(condition, truePart, falsePart)
    If condition Then
        IIf = truePart
    Else
        IIf = falsePart
    End If
End Function


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function

'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity
'Dim ULF : Set ULF = New FlipperPolarity

InitPolarity

'Late 80's early 90's

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5
    x.AddPt "Polarity", 2, 0.16, - 5
    x.AddPt "Polarity", 3, 0.22, - 0
    x.AddPt "Polarity", 4, 0.25, - 0
    x.AddPt "Polarity", 5, 0.3, - 2
    x.AddPt "Polarity", 6, 0.4, - 3
    x.AddPt "Polarity", 7, 0.5, - 4.0
    x.AddPt "Polarity", 8, 0.7, - 3.5
    x.AddPt "Polarity", 9, 0.75, - 3.0
    x.AddPt "Polarity", 10, 0.8, - 2.5
    x.AddPt "Polarity", 11, 0.85, - 2.0
    x.AddPt "Polarity", 12, 0.9, - 1.5
    x.AddPt "Polarity", 13, 0.95, - 1.0
    x.AddPt "Polarity", 14, 1, - 0.5
    x.AddPt "Polarity", 15, 1.1, 0
    x.AddPt "Polarity", 16, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.15, 0.85
    x.AddPt "Velocity", 2, 0.2, 0.9
    x.AddPt "Velocity", 3, 0.23, 0.95
    x.AddPt "Velocity", 4, 0.41, 0.95
    x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
    x.AddPt "Velocity", 6, 0.62, 1.0
    x.AddPt "Velocity", 7, 0.702, 0.968
    x.AddPt "Velocity", 8, 0.95,  0.968
    x.AddPt "Velocity", 9, 1.03,  0.945
    x.AddPt "Velocity", 10, 1.5,  0.945

  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
  LF.SetObjects "LF", LeftFlipper, TriggerLF
  RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub



'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before\
' modified 2024 by rothbauerw
' Added Reprocessballs for flipper collisions (LF.Reprocessballs Activeball and RF.Reprocessballs Activeball must be added to the flipper collide subs
' Improved handling to remove correction for backhand shots when the flipper is raised

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
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
      If Not IsEmpty(balls(x)) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        balldata(x).Data = balls(x)
      End If
    Next
    FlipStartAngle = Flipper.currentangle
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub

  Public Sub ReProcessBalls(aBall) 'save data of balls in flipper range
    If FlipperOn() Then
      Dim x
      For x = 0 To UBound(balls)
        If Not IsEmpty(balls(x)) Then
          if balls(x).ID = aBall.ID Then
            If isempty(balldata(x).ID) Then
              balldata(x).Data = balls(x)
            End If
          End If
        End If
      Next
    End If
  End Sub

  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
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
          BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
        NoCorrection = 1
      Else
        checkHit = 50 + (20 * BallPos)

        If BalltoFlip > checkHit or (PartialFlipCoef < 0.5 and BallPos > 0.22) Then
          NoCorrection = 1
        Else
          NoCorrection = 0
        End If
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        'If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled and NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
      End If
      If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
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
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines
'  - Include FlipperActivate and FlipperDeactivate in the Flipper solenoid subs

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
     Dim BOT
     BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          gBOT(b).velx = BOT(b).velx / 1.3
          gBOT(b).vely = BOT(b).vely - 0.5
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



'*************************************************
'  Check ball distance from Flipper for Rem
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

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
  DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(angle + 90)) + Flipper.y)
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
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, ULFPress, LFCount, RFCount
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
Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
'Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
   Const EOSReturn = 0.035  'mid 80's to early 90's
'    Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
    Dim BOT, b

    FlipperPress = 0
    Flipper.eostorqueangle = EOSA
    Flipper.eostorque = EOST * EOSReturn / FReturn

    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
        BOT = GetBalls   ' fresh list of valid balls on the table

        If IsArray(BOT) Then
            For b = 0 To UBound(BOT)
                If IsObject(BOT(b)) Then
                    ' check for cradle near this flipper
                    If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then
                        ' clamp downward speed a bit
                        If BOT(b).VelY >= -0.4 Then BOT(b).VelY = -0.4
                    End If
                End If
            Next
        End If
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
'****  END FLIPPER CORRECTIONS
'******************************************************

Const GATE_MIN_VELY = 2   ' tweak: 1–5
Const GATE_MIN_SPEED = 6  ' tweak: 4–10

Sub GateBackStop_Hit()
    Dim vy, spd
    vy  = ActiveBall.VelY
    spd = Sqr(ActiveBall.VelX*ActiveBall.VelX + ActiveBall.VelY*ActiveBall.VelY)

    ' Only if the ball is moving DOWN the table (toward flippers)
    If vy > GATE_MIN_VELY And spd > GATE_MIN_SPEED Then
        PlaySoundAtVol "Gate_FastTrigger_1",Gate, 0.6
    End If
End Sub

Sub GateBackStop2_Hit()
    Dim vy, spd
    vy  = ActiveBall.VelY
    spd = Sqr(ActiveBall.VelX*ActiveBall.VelX + ActiveBall.VelY*ActiveBall.VelY)

    ' Only if the ball is moving DOWN the table (toward flippers)
    If vy > GATE_MIN_VELY And spd > GATE_MIN_SPEED Then
        PlaySoundAtVol "Gate_FastTrigger_1",Gate2, 0.6
    End If
End Sub

'******************************************************
'  BALL ROLLING (boilerplate) + ROM/MYSTERY SAUCERS (moved over)
'******************************************************
Const tnob = 3 ' total number of balls
ReDim rolling(tnob)

Dim DropCount
ReDim DropCount(tnob)

InitRolling

Sub InitRolling
    Dim i
    For i = 0 To tnob
        rolling(i) = False
        DropCount(i) = 0
    Next
End Sub

' --- GLASS HIT (ball hits glass at Z=52) ---
Const GlassZ = 52
Const GlassHitCooldown = 0.25
Dim LastGlassHitTime

Sub RollingUpdate()
    Dim b
    gBOT = GetBalls

    ' --- stop the sound of deleted balls (boilerplate) ---
    For b = UBound(gBOT) + 1 To tnob - 1
        rolling(b) = False
        StopSound "BallRoll_" & b
    Next

    ' --- exit if no balls on table ---
    If UBound(gBOT) = -1 Then Exit Sub

    ' --- (ported from your original rolling timer) stop rolling if ball is basically dead ---
    ' helps prevent "stuck/hum" states when a ball index exists but isn't really moving
    For b = 0 To UBound(gBOT)
        If rolling(b) Then
            ' stop the loop when ball is essentially not moving (kills hum)
            If Abs(gBOT(b).VelX) < 0.2 And Abs(gBOT(b).VelY) < 0.2 Then
                rolling(b) = False
                StopSound "BallRoll_" & b
            End If
        End If
    Next

    ' --- additional GI timers (kept; original behavior) ---
    If EnableAdditionalGI = 1 Then
        MultiballTimer.Enabled = IsMultiball
        AttractTimer.Enabled   = IsGameOver
    End If

    ' --- optional: original code bailed when gameover (kept to preserve gameplay expectations) ---
    If IsGameOver Then
        FlippersEnabled = False
        Exit Sub
    End If

    ' --- main loop ---
    For b = 0 To UBound(gBOT)

        '==================================================
        ' GLASS HIT (real ball height check)
        '==================================================
        'If gBOT(b).Z >= GlassZ And gBOT(b).VelZ > 0 Then
        '    If Timer - LastGlassHitTime > GlassHitCooldown Then
        '        LastGlassHitTime = Timer
        '        If Int(Rnd * 2) = 0 Then PlaySound "glasshit"
        '    End If
        'End If
    If gBOT(b).Z >= GlassZ _
       And gBOT(b).VelZ > 0 _
       And gBOT(b).X <= 755 _
       And gBOT(b).Y >= 115 Then

      If Timer - LastGlassHitTime > GlassHitCooldown Then
        LastGlassHitTime = Timer
        If Int(Rnd * 2) = 0 Then PlaySound "glasshit"
      End If

    End If



        '==================================================
        ' 1) ROLLING SOUND (boilerplate)
        '==================================================
        If BallVel(gBOT(b)) > 1 And gBOT(b).Z < 30 Then
            rolling(b) = True
            PlaySound "BallRoll_" & b, -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
        Else
            If rolling(b) Then
                StopSound "BallRoll_" & b
                rolling(b) = False
            End If
        End If

        '==================================================
        ' 2) BALL DROP SOUNDS (boilerplate)
        '==================================================
        If gBOT(b).VelZ < -1 And gBOT(b).Z < 55 And gBOT(b).Z > 27 Then
            If DropCount(b) >= 5 Then
                DropCount(b) = 0
                If gBOT(b).VelZ > -7 Then
                    RandomSoundBallBouncePlayfieldSoft gBOT(b)
                Else
                    RandomSoundBallBouncePlayfieldHard gBOT(b)
                End If
            End If
        End If

        If DropCount(b) < 5 Then
            DropCount(b) = DropCount(b) + 1
        End If

        '==================================================
        ' 3) ROM / MYSTERY SAUCERS (MOVED OVER from original rolling timer)
        '    NOTE: intentionally NOT bringing over:
        '      - NoBallDropTimer.Enabled block
        '      - NoGlassHitTimer.Enabled block
        '==================================================

        ' --- top left saucer sw20 ---
        If InCircle(gBOT(b).X, gBOT(b).Y, sw20.X, sw20.Y, 20) Then
            If Abs(gBOT(b).VelX) < 2 And Abs(gBOT(b).VelY) < 2 Then
                gBOT(b).VelX = gBOT(b).VelX / 2
                gBOT(b).VelY = gBOT(b).VelY / 2

                If Not Controller.Switch(20) Then
                    Controller.Switch(20) = True
                    Set sw20Ball = gBOT(b)

                    If Not Saucer20MysteryTimer.Enabled Then
                        Saucer20MysteryStep = 0
                        Saucer20MysteryTimer.Enabled = (EnableAdditionalGI = 1)
                    End If
                End If
            End If
        End If

        ' --- middle right saucer sw21 ---
        If InCircle(gBOT(b).X, gBOT(b).Y, sw21.X, sw21.Y, 20) Then
            If Abs(gBOT(b).VelX) < 2 And Abs(gBOT(b).VelY) < 2 Then
                gBOT(b).VelX = gBOT(b).VelX / 2
                gBOT(b).VelY = gBOT(b).VelY / 2

                If Not Controller.Switch(21) Then
                    Controller.Switch(21) = True
                    Set sw21Ball = gBOT(b)

                    If Not Saucer21MysteryTimer.Enabled Then
                        Saucer21MysteryStep = 0
                        Saucer21MysteryTimer.Enabled = (EnableAdditionalGI = 1)
                    End If
                End If
            End If
        End If

    Next
End Sub


Sub ToggleGI()
    ' Flip current GI state and apply immediately
    SetGI (Not isGIOn), True
End Sub

Sub ToggleVR()
    If vr = 0 Then
        vr = 1
    Else
        vr = 0
    End If
ToggleGI
End Sub

'*************
'VR Stuff
'*************
Dim VRRoom
Dim VR_Obj

Sub LoadVRRoom
    for each VR_Obj in VRCab:VR_Obj.visible = 0:Next
    for each VR_Obj in VRMega:VR_Obj.visible = 0:Next
    for each VR_Obj in VRMinRoom:VR_Obj.visible = 0:Next
    Ramp16.visible=1
    Ramp15.visible=1
    For Each VR_Obj in VRBackglass : VR_Obj.Visible = 0 : Next

  If RenderingMode = 2 or VRTest Then
    VRRoom = VRRoomChoice
    Ramp16.visible=0
    Ramp15.visible=0
    For Each VR_Obj in VRBackglass : VR_Obj.Visible = 1 : Next
    'disable table objects that should not be visible
  Else
    VRRoom = 0
  End If
  If VRRoom = 1 Then
    for each VR_Obj in VRCab:VR_Obj.visible = 1:Next
    for each VR_Obj in VRMega:VR_Obj.visible = 0:Next
    for each VR_Obj in VRMinRoom:VR_Obj.visible = 0:Next
  End If
  If VRRoom = 2 Then
    for each VR_Obj in VRCab:VR_Obj.visible = 1:Next
    for each VR_Obj in VRMega:VR_Obj.visible = 1:Next
    for each VR_Obj in VRMinRoom:VR_Obj.visible = 0:Next
  End If
  If VRRoom = 3 Then
    for each VR_Obj in VRCab:VR_Obj.visible = 1:Next
    for each VR_Obj in VRMega:VR_Obj.visible = 0:Next
    for each VR_Obj in VRMinRoom:VR_Obj.visible = 1:Next
  End If
End Sub

Sub BlinkGI()
    ' turn GI off immediately
    SetGI False, True

    ' arm the timer to turn it back on
    GIBlinkTimer.Enabled = False
    GIBlinkTimer.Enabled = True
End Sub

Sub GIBlinkTimer_Timer()
    GIBlinkTimer.Enabled = False
    SetGI True, True
End Sub

'the below was added becasue thi call to the rom switch was present in the plunger keydown. it would
'cause suspended games if the plunger was pressed between balls. if i remove the switch we get rom stutter
' so i added it tothe ball release gate instead

sub gate1_hit
vpmTimer.PulseSwitch 26,0,0
end Sub

' the below code was added because the left kickback was not alwasy firing when light77 was on or blinking
'i added a kicker and a timer. the timer enables and disables the kicker depending on if the light if off or on/blinking
'then the trigger above the kicker registers with the riom that the switch was hit so it can add the poionts and
'fire the plunger (after the ball already launched but oh well)

Sub Kickpulsetrigger_Hit()
    If ActiveBall.VelY < 0 Then      ' ball moving UP the table
    vpmTimer.PulseSw 23
    PlaySoundAtLevelStatic ("warehousekick"), DrainSoundLevel, LeftPlunger
    End If
End Sub

'sub kicktimer_timer
'Kickback.Enabled = (Light77.State <> 0)
'end sub

'Sub KickTimer_Timer()
'    Static tLastOn  Expected statement
'    If Light77.State <> 0 Then tLastOn = GameTime
'    Kickback.Enabled = (GameTime - tLastOn) < 2500
'End Sub

Dim KickLastOn
Sub KickTimer_Timer()
    If Light77.State <> 0 Then KickLastOn = GameTime
    Kickback.Enabled = (GameTime - KickLastOn) < 2000
End Sub



sub Kickback_hit
  Kickback.kick 0,30
  PlaySoundAtLevelStatic ("warehousekick"), DrainSoundLevel, LeftPlunger
end sub

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
' for VPX by Schreibi34, 2019
'
' Special thanks to
'
' - Herweh for scripting everything on this table, including kickers, animations, airball, additional GI and SSF sounds. He also fixed stuff like the trough and the tilt script!
'   Without him this table wouldn't be half as good!!! One of the best and a great guy!! This is also your table!! Thanks!!!
'
' - 32assassin and BodyDump for bringing this Table to VPX. I used their table as a base.
'
' - Kiwi and Destruk for a great hint when 'Tilt' didn't want to work the way it should
'
' - RothbauerW for some code of his manuel trough in 'No Good Gofers'
'
' - Bord and BorgDog for images, testing and all kinds of help!
'
' - BorgDog for moving spinner rod
'
' - Flupper - cycles insert material
'
' - Thalamus for playtesting and screenshots
'
' - Ninuzzu - BallShadow
'
' - Whoever made the flipper shadow
'
' - Last, and definetly not least the VPX devs!! Thanks guys!
'
' Release Notes:
' 1.0.2: Some bugfixes
'        - fixed a graphical glitch during multiball
'        - improved the usage of 'PlaySound'
'
' 1.0.1: Some bugfixes
'        - improved the usage of 'ActiveBall'
'    - fixed flipper texture
'    - better DT image
'
' 1.0: Initial release
'
' ****************************************************
' OPTIONS
' ****************************************************
Dim EnableAdditionalGI, ShowBallShadow, EnableGI

' ENABLE/DISABLE additional GI (like dark playfield during "Mystery" or "Drain", flashing "Attract Mode" and "Multiball", etc.)
' 0 = additional GI is off (default)
' 1 = additional GI is on
EnableAdditionalGI = 1

' ENABLE/DISABLE GI (general illumination)
' 0 = GI is off
' 1 = GI is on (default)
EnableGI = 1

' SHOW/HIDE the ball shadow
' 0 = ball shadow are hidden
' 1 = ball shadow are visible (default)
ShowBallShadow = 1


' ***************************************************************************************************************************************************************


' ****************************************************
' dynamic flipper friction mod and settings
' ****************************************************
Const DynamicFlipperFriction    = True
Const DynamicFlipperFrictionResting = 0.4
Const DynamicFlipperFrictionActive  = 1.0


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
Const BallMass    = 1.3

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

' Thalamus : Was missing 'vpminit me'
  vpminit me

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

  InitTable
End Sub

Sub Table1_Paused()   : Controller.Pause = True  : End Sub
Sub Table1_UnPaused() : Controller.Pause = False : End Sub
Sub Table1_Exit()   : Controller.Stop      : End Sub

Sub InitTable()
  ' some timer inits
  PinMAMETimer.Interval = PinMAMEInterval
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
  Ramp15.Visible = DesktopMode
  Ramp16.Visible = DesktopMode

  ' maybe no GI
  If EnableGI = 0 Then SetGI False, False
End Sub


' ****************************************************
' solenoids - increase by 1 because first solenoid is 0
' ****************************************************
SolCallback(5)    = "SolKickingTarget"
SolCallback(7)    = "SolKickback"
SolCallBack(8)    = "SolResetDropTargets"
SolCallback(9)    = "SolLeftKicker"
SolCallback(10)   = "SolRightKicker"

SolCallback(26)   = "SolGI"
SolCallback(28)   = "SolBallRelease"
SolCallback(29)   = "SolDrain"
SolCallback(30)   = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
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
' right now nothing to do

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
  PlaySoundAt "drain", Drain
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
    PlaySoundAt SoundFX("Ballrelease",DOFContactors), RightSlot
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
  If keycode = PlungerKey Then    Plunger.Pullback : PlaySoundAt "plungerpull", Plunger
    If keycode = LeftFlipperKey Then  SolLFlipper True
  If keycode = RightFlipperkey Then SolRFlipper True
    If keycode = StartGameKey Then    Controller.Switch(3)  = True
  If keycode = AddCreditKey Then    Controller.Switch(1)  = True
  If keycode = AddCreditKey2 Then   Controller.Switch(2)  = True
  If keycode = KeySlamDoorHit Then  Controller.Switch(152)  = True
  If keycode = LeftMagnaSave Then   Controller.Switch(4)  = True
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
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then    Plunger.Fire : PlaySoundAt "plunger", Plunger
    If keycode = LeftFlipperKey Then  SolLFlipper False
  If keycode = RightFlipperkey Then SolRFlipper False
  If keycode = StartGameKey Then    Controller.Switch(3)  = False
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
      PlaySoundAt SoundFX("fx_Flipperup", DOFContactors), LeftFlipper
      If DynamicFlipperFriction Then LeftFlipper.Friction = DynamicFlipperFrictionActive : UpperLeftFlipper.Friction = DynamicFlipperFrictionActive : End If
      LeftFlipper.RotateToEnd
      UpperLeftFlipper.RotateToEnd
    End If
    Else
        If FlippersEnabled Then PlaySoundAt SoundFX("fx_Flipperdown", DOFContactors), LeftFlipper
    If DynamicFlipperFriction Then LeftFlipper.Friction = DynamicFlipperFrictionResting : UpperLeftFlipper.Friction = DynamicFlipperFrictionResting : End If
    LeftFlipper.RotateToStart
    UpperLeftFlipper.RotateToStart
    End If
 End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    If FlippersEnabled Then
      PlaySoundAt SoundFX("fx_Flipperup", DOFContactors), RightFlipper
      If DynamicFlipperFriction Then RightFlipper.Friction = DynamicFlipperFrictionActive
      RightFlipper.RotateToEnd
    End If
    Else
        If FlippersEnabled Then PlaySoundAt SoundFX("fx_Flipperdown", DOFContactors), RightFlipper
    If DynamicFlipperFriction Then RightFlipper.Friction = DynamicFlipperFrictionResting
    RightFlipper.RotateToStart
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

' flipper hit sounds
Sub LeftFlipper_Collide(parm)
  PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 1
End Sub
Sub UpperLeftFlipper_Collide(parm)
  PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 1
End Sub
Sub RightFlipper_Collide(parm)
  PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 1
End Sub


' ****************************************************
' left kick back
' ****************************************************
Sub SolKickBack(Enabled)
    If Enabled Then
       LeftPlunger.Fire
       PlaySoundAt SoundFX("Popper", DOFContactors), LeftPlunger
    Else
       LeftPlunger.PullBack
    End If
End Sub


' ****************************************************
' sling shots and animations
' ****************************************************
Dim RightStep

Sub RightSlingShot_Slingshot()
  vpmTimer.PulseSwitch 15,0,0
    PlaySoundAt SoundFX("right_slingshot", DOFContactors), SLING1
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
Sub Bumper10_Hit : MoveSkirt 10, 1, Bumper10, BumperSkirt10, ActiveBall : End Sub
Sub Bumper11_Hit : MoveSkirt 11, 2, Bumper11, BumperSkirt11, ActiveBall : End Sub
Sub Bumper12_Hit : MoveSkirt 12, 3, Bumper12, BumperSkirt12, ActiveBall : End Sub
Sub Bumper13_Hit : MoveSkirt 13, 4, Bumper13, BumperSkirt13, ActiveBall : End Sub

Sub Bumper10_Timer() : MoveSkirt 10, 1, Bumper10, BumperSkirt10, Nothing : End Sub
Sub Bumper11_Timer() : MoveSkirt 11, 2, Bumper11, BumperSkirt11, Nothing : End Sub
Sub Bumper12_Timer() : MoveSkirt 12, 3, Bumper12, BumperSkirt12, Nothing : End Sub
Sub Bumper13_Timer() : MoveSkirt 13, 4, Bumper13, BumperSkirt13, Nothing : End Sub

ReDim skirtX(3), skirtY(3)
Sub MoveSkirt(id, soundid, bumper, skirt, ball)
  If Not ball Is Nothing Then
    bumper.TimerEnabled   = False
    PlaySoundAt SoundFX("fx_bumper" & soundid, DOFContactors), bumper
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
Sub sw14_Slingshot(): vpmTimer.PulseSwitch 14,0,0 : SlingTarget sw14, sw14P, True : End Sub
Sub sw14_Timer()  : SlingTarget sw14, sw14P, False : End Sub

Dim SlingSteps : SlingSteps = 0
Sub SlingTarget(sw, prim, init)
  If init Then
    sw.TimerEnabled   = False
    SlingSteps      = 0
    prim.RotX       = 0
    sw.TimerInterval  = 10
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
Sub sw22_Hit    : Controller.Switch(22) = True  : RollOverSound ActiveBall : End Sub
Sub sw22_UnHit    : Controller.Switch(22) = False : End Sub
Sub sw23_Hit    : Controller.Switch(23) = True  : RollOverSound ActiveBall : End Sub
Sub sw23_UnHit    : Controller.Switch(23) = False : End Sub
Sub sw25_Hit    : Controller.Switch(25) = True  : RollOverSound ActiveBall : End Sub
Sub sw25_UnHit      : Controller.Switch(25) = False : End Sub
Sub sw32_Hit    : Controller.Switch(32) = True  : RollOverSound ActiveBall : End Sub
Sub sw32_UnHit    : Controller.Switch(32) = False : End Sub
Sub sw33_Hit    : Controller.Switch(33) = True  : RollOverSound ActiveBall : End Sub
Sub sw33_UnHit    : Controller.Switch(33) = False : End Sub
Sub sw35_Hit    : Controller.Switch(35) = True  : RollOverSound ActiveBall : FlippersEnabled = True : End Sub
Sub sw35_UnHit    : Controller.Switch(35) = False : End Sub
Sub sw40_Hit    : Controller.Switch(40) = True  : RollOverSound ActiveBall : End Sub
Sub sw40_UnHit    : Controller.Switch(40) = False : End Sub
Sub sw41_Hit    : Controller.Switch(41) = True  : RollOverSound ActiveBall : End Sub
Sub sw41_UnHit    : Controller.Switch(41) = False : End Sub
Sub sw42_Hit    : Controller.Switch(42) = True  : RollOverSound ActiveBall : End Sub
Sub sw42_UnHit    : Controller.Switch(42) = False : End Sub
Sub sw43_Hit    : Controller.Switch(43) = True  : RollOverSound ActiveBall : End Sub
Sub sw43_UnHit    : Controller.Switch(43) = False : End Sub
Sub sw45_Hit    : Controller.Switch(45) = True  : RollOverSound ActiveBall : End Sub
Sub sw45_UnHit    : Controller.Switch(45) = False : End Sub

' spinners
Sub Spinner1_Spin() : vpmTimer.PulseSwitch 30,0,0 : PlaySoundAt "fx_spinner", Spinner1 : End Sub
Sub Spinner2_Spin() : vpmTimer.PulseSwitch 31,0,0 : PlaySoundAt "fx_spinner", Spinner2 : End Sub

' drop targets
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
  PlaySoundAt SoundFX("DTReset",DOFContactors), sw42
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
    If Not actBall Is Nothing Then DropTargetHit playsnd, actBall
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
    Case 1 :  MoveHammer sw20Hammer, 16 : If Controller.Switch(20) Then sw20Ball.VelX = 2.5+(Rnd()*0.5-0.25) : sw20Ball.VelY = 7+(Rnd()-0.5) : sw20Ball.VelZ = 12+(Rnd()*2-1) : End If
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
  TextBox1.text = Spinner1.CurrentAngle
  TextBox2.text = Angle1
  PSpinner1.RotX = Spinner1.CurrentAngle
    PSpinner1Off.RotX = Spinner1.CurrentAngle
    SpinnerRod1.TransZ = -sin( (Spinner1.CurrentAngle) * (2*3.14/360)) * 5
    SpinnerRod1.TransX = (sin( (Spinner1.CurrentAngle- 90) * (2*3.14/360)) * -5)
' SpinnerRod1.TransZ =  cos ((Spinner1.CurrentAngle)) * -2
' SpinnerRod1.TransY = Angle1 * -2
End Sub

Sub SpinnerTimer2_Timer()
  Angle2 = (sin (Spinner2.CurrentAngle))
  TextBox3.text = Spinner2.CurrentAngle
  TextBox4.text = Angle2
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


' *********************************************************************
' GI (seems to be initially on and the "Enabled" from the solenoid is vice versa)
' *********************************************************************
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
    For Each obj In GIOnTrans : obj.Material = "Prims platt trans" & IIF(GILevel=4,""," " & 4-GILevel) : Next
    For Each obj In GIOnShiny : obj.Material = "Prims Shiny" & IIF(GILevel=4,""," " & 4-GILevel) : Next
    If GILevel = 1 Or GILevel = 4 Then
      For Each coll In Array(GIOn,GIOnTrans,GIOnShiny) : For Each obj In coll : obj.Visible = True : Next : Next
    End If
    If GILevel = 4 Then
      For Each coll In Array(GIOff,GIOffTrans,GIOffShiny) : For Each obj In coll : obj.Visible = False : Next : Next
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
    For Each obj In GIOnTrans : obj.Material = "Prims platt trans " & (GILevel + 5) : Next
    For Each obj In GIOnShiny : obj.Material = "Prims Shiny " & (GILevel + 5) : Next
    If GILevel = -4 Or GILevel = -1 Then
      For Each coll In Array(GIOff,GIOffTrans,GIOffShiny) : For Each obj In coll : obj.Visible = True : Next : Next
      For Each obj In GIFlasher : obj.Visible = False : Next
      aBumperValue(3) = 3
    End If
    If GILevel = -1 Then
      For Each coll In Array(GIOn,GIOnTrans,GIOnShiny) : For Each obj In coll : obj.Visible = False : Next : Next
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

Sub DisplayTimer_Timer()
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
  If playsnd Then DropTargetSound actBall
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
  PlaySoundAtVolPitch SoundFX("rollover",DOFContactors), actBall, 0.02, .25
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

Sub Targets_Hit(idx)
  PlaySoundAtBall "target"
End Sub

Sub Metals_Thin_Hit(idx)
  PlaySoundAtBall "metalhit_thin"
End Sub

Sub Metals_Medium_Hit(idx)
  PlaySoundAtBall "metalhit_medium"
End Sub

Sub Metals2_Hit(idx)
  PlaySoundAtBall "metalhit2"
End Sub

Sub Gates_Hit(idx)
  PlaySoundAtBall "fx_gate"
End Sub

Sub Rubbers_Hit(idx)
  Dim finalspeed
    finalspeed = SQR(ActiveBall.VelX ^ 2 + ActiveBall.VelY ^ 2)
  If finalspeed > 20 then
    PlaySoundAtBall "fx_rubber2"
  ElseIf finalspeed >= 4 then
    PlaySoundAtBall "rubber_hit_" & (Int(Rnd*3)+1)
  End If
  RubberRingHit ActiveBall
End Sub

Sub Posts_Hit(idx)
  Dim finalspeed
    finalspeed = SQR(ActiveBall.VelX ^ 2 + ActiveBall.VelY ^ 2)
  If finalspeed > 16 then
    PlaySoundAtBall "fx_rubber2"
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtBall "rubber_hit_" & (Int(Rnd*3)+1)
  End If
  RubberPostHit ActiveBall
End Sub


' *********************************************************************
' Supporting Surround Sound Feedback (SSF) functions
' *********************************************************************
' set position as table object (Use object or light but NOT wall) and Vol to 1
Sub PlaySoundAt(sound, tableobj)
  PlaySound sound, 1, 1, AudioPan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
End Sub
' set position as table object and Vol manually.
Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
End Sub
' set position as table object and Vol + RndPitch manually
Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

' set all as per ball position & speed.
Sub PlaySoundAtBall(sound)
  PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub
' set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub
Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

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
Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 200)
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 100)
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table.
  Dim tmp : tmp = tableobj.x * 2 / Table1.Width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table1.width-1
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

Function AudioFade(ball)
    Dim tmp : tmp = ball.y * 2 / Table1.Height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function


' *********************************************************************
' JP's VPX rolling sounds
' *********************************************************************
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

  ' stop the sound of drained balls
    For b = 0 to UBound(BOT)
    If rolling(b) And BOT(b).VelX <= 0 And BOT(b).VelY <= 0 Then
      rolling(b) = False
      StopSound "fx_ballrolling" & b
    End If
    Next

  ' maybe start additional GI for multiball and/or attract mode
  If EnableAdditionalGI = 1 Then
    MultiballTimer.Enabled = IsMultiball
    AttractTimer.Enabled = IsGameOver
  End If

  ' exit the sub if no balls on the table
  If IsGameOver Then
    FlippersEnabled = False
    Exit Sub
  End If

  ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b)) > 1 AND BOT(b).Z < 30 Then
            rolling(b) = True
            PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b)), Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0
        Else
            If rolling(b) Then
                StopSound "fx_ballrolling" & b
                rolling(b) = False
            End If
        End If

    '   ball drop sounds matching the adjusted height params
    If Not NoBallDropTimer.Enabled And BOT(b).VelZ < -3 And BOT(b).Z < 50 And BOT(b).Z > 25 then
      NoBallDropTimer.Enabled = True
      PlaySound "fx_balldrop" & Int(Rnd()*3), 0, IIF((BOT(b).VelZ>-6), ABS(BOT(b).VelZ)/100, ABS(BOT(b).VelZ)/10), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If

    If AirBall = 1 And Not NoGlassHitTimer.Enabled And BOT(b).Z > 50 Then
      NoGlassHitTimer.Enabled = True
      PlaySoundAt "fx_GlassHit6", BOT(b)
    End If

    ' top left saucer sw20
    If InCircle(BOT(b).X, BOT(b).Y, sw20.X, sw20.Y, 20) Then
'   If InRect(BOT(b).X, BOT(b).Y, 270, 30, 310, 30, 310, 70, 270, 70) Then
      If Abs(BOT(b).VelX) < 2 And Abs(BOT(b).VelY) < 2 Then
        BOT(b).VelX = BOT(b).VelX / 2
        BOT(b).VelY = BOT(b).VelY / 2
        If Not Controller.Switch(20) Then
          Controller.Switch(20)   = True
          Set sw20Ball      = BOT(b)
          If Not Saucer20MysteryTimer.Enabled Then
            Saucer20MysteryStep       = 0
            Saucer20MysteryTimer.Enabled  = (EnableAdditionalGI = 1)
          End If
        End If
      End If
    End If

    ' middle right saucer sw21
    If InCircle(BOT(b).X, BOT(b).Y, sw21.X, sw21.Y, 20) Then
'   If InRect(BOT(b).X, BOT(b).Y, 790, 275, 835, 275, 835, 310, 790, 310) Then
      If Abs(BOT(b).VelX) < 2 And Abs(BOT(b).VelY) < 2 Then
        BOT(b).VelX = BOT(b).VelX / 2
        BOT(b).VelY = BOT(b).VelY / 2
        If Not Controller.Switch(21) Then
          Controller.Switch(21)   = True
          Set sw21Ball      = BOT(b)
          If Not Saucer21MysteryTimer.Enabled Then
            Saucer21MysteryStep       = 0
            Saucer21MysteryTimer.Enabled  = (EnableAdditionalGI = 1)
          End If
        End If
      End If
    End If
  Next
End Sub

Sub NoBallDropTimer_Timer()
  NoBallDropTimer.Enabled = False
End Sub
Sub NoGlassHitTimer_Timer()
  NoGlassHitTimer.Enabled = False
End Sub


' *********************************************************************
' ball collision sound
' *********************************************************************
Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound "fx_collide", 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub


' *********************************************************************
' some general stuff
' *********************************************************************

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

Function InCircle(pX, pY, centerX, centerY, radius)
  Dim route
  route = Sqr((pX - centerX) ^ 2 + (pY - centerY) ^ 2)
  InCircle = (route < radius)
End Function

Function IIF(bool, obj1, obj2)
  If bool Then
    IIF = obj1
  Else
    IIF = obj2
  End If
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

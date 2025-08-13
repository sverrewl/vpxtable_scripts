'   ____  ___________________   ______  __________
'    / __ \/ ____/ ____/ ____/ | / / __ \/ ____/ __ \
'   / / / / __/ / /_  / __/ /  |/ / / / / __/ / /_/ /
'  / /_/ / /___/ __/ / /___/ /|  / /_/ / /___/ _, _/
' /_____/_____/_/   /_____/_/ |_/_____/_____/_/ |_|
'
'    Williams 1982
'  System 7
'
'
' DEFENDER TEAM
' -------------
' Apophis - Project lead, scripting, Roth/nFozzy physics, Fleep sound effects, Flupper 3D inserts
' Bord - 3D models and rendering
' Uncle Paulie - VR Room
' Jutsuka - Asset research and testing
' G94 - Playfield and plastics scans
' Scutters - FlexDMD Option
' Primetime5k - Staged flipper updates
' Cliffy - Desktop view updates
' Jsm174 - Standalone updates
' VPW Team - Testing
'
' Thank you to G94 for providing incredible playfield and plastic scans. Go check out his Defender playfield resotration here:
' https://pinside.com/pinball/forum/topic/defender-williams-1982-nos-playfield-


Option Explicit
Randomize


' -- FlexDMD Option --
Dim UseFlexDMD:UseFlexDMD = true        ' True = on, False = off. Use a FlexDMD for scoring etc instead of segment display when on
Dim FlexDMDStyle:FlexDMDStyle = 0       ' 0 - 128x32 4:1 DMD showing score info
                        ' 1 - 128x32 4:1 DMD showing scanner type display (use backglass scoring)
                        ' 2 - 128x72 FullDMD 16:9 ratio DMD showing score and scanner displays. For desktop or FullDMD users (but may also look ok a HD Real DMD)

'DMD Scanner key :  Humanoids (lamps) - Grey Square, Bombers (unlit bomber lanes) - Magenta square, Swarmers (lit swarmer lamps) - Red and yellow verical bars, Baiters (drop Targets) - Green square
'         Landers (drop Targets) - Yellow bar above greeen bar, Mutants (drop targets with lamp) - Green horizontal bar below colour colour changing bar, Pods (drop targets) - Colour changing square,



'**************************
'   Option Setup
'**************************
Const TestVR = false
Const DynamicBallShadowsOn=1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn=1     '0 = Static shadow under ball ("flasher" image, like JP's)
                    '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                    '2 = flasher image shadow, but it moves like ninuzzu's
'
' NOTE: DO NOT adjust these manually. Use the Tweak menu by pressing F12 in game
'
Dim VolumeDial        ' Recommended values should be no greater than 1.
Dim BallRollingVolume   ' Volume of ball rolling sound. Value ranges between 0 and 1
Dim CustomWalls         'set to 0 for Modern Minimal Walls, floor, and roof, 1 for Sixtoe's original walls and floor
Dim StagedFlippers          '0 = No Staged Flippers, 1 = Upper Left Flipper Staged
Dim WallClock           '1 Shows the clock in the VR room only
Dim topper          '0 = Off 1= On - Topper visible in VR Room only
Dim poster            '1 Shows the flyer posters in the VR room only
Dim poster2           '1 Shows the flyer posters in the VR room only

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    Dim v
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollingVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

    ' Staged Flippers
    StagedFlippers = Table1.Option("Staged Flippers", 0, 1, 1, 0, 0, Array("Normal", "Staged"))

  ' VR Options
  CustomWalls = Table1.Option("VR Custom Walls", 0, 1, 1, 0, 0, Array("Minimal", "Original"))
    WallClock   = Table1.Option("VR Wall Clock",   0, 1, 1, 1, 0, Array("Off", "On"))
  topper      = Table1.Option("VR Topper",       0, 1, 1, 1, 0, Array("Off", "On"))
  poster      = Table1.Option("VR Poster 1",     0, 1, 1, 1, 0, Array("Off", "On"))
  poster2     = Table1.Option("VR Poster 2",     0, 1, 1, 1, 0, Array("Off", "On"))

  SetupVRRoom

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub




'**********************************************************************************************************
' Load Table
'**********************************************************************************************************


Dim cab_mode, VR_Room, DesktopMode: DesktopMode = Table1.ShowDT
If RenderingMode = 2 OR TestVR Then VR_Room = 1 Else VR_Room = 0
If Not DesktopMode and VR_Room=0 Then cab_mode=1 Else cab_mode=0

Const BallMass = 1
Const BallSize = 50

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="dfndr_l4",UseSolenoids=0,UseLamps=0,UseGI=0,SSolenoidOn="",SSolenoidOff="", SCoin=""

LoadVPM "", "S7.VBS", 2.0

Dim GameRun
Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height


dim VPMversion

'**********************************************************************************************************
' Timers
'**********************************************************************************************************

Sub GameTimer_Timer
  Cor.Update
  UpdateSolenoids
  UpdateRolling
End Sub


Sub FrameTimer_Timer
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows

  Pgate1.rotx = -max(Gate1.CurrentAngle,0)
  Pgate2.rotx = -max(Gate2.CurrentAngle,0)
  gatewire.rotx = max(Gate3.CurrentAngle,0)
  diverter.roty = Flipper3.currentangle
  LeftBat1.ObjRotZ = LeftFlipper1.CurrentAngle
  LeftBat.ObjRotZ = LeftFlipper.CurrentAngle
  RightBat.ObjRotZ = RightFlipper.CurrentAngle
  batleftshadow1.ObjRotZ = LeftFlipper1.CurrentAngle
  batleftshadow.ObjRotZ = LeftFlipper.CurrentAngle
  batrightshadow.ObjRotZ = RightFlipper.CurrentAngle

  If UseFlexDMD Then Exit Sub   'flexdmd controlled by FlexTimer which calls DisplayTimer sub
DisplayTimer
' If (VR_Room = 0 AND cab_mode = 0) OR VR_Room = 1 Then
'   DisplayTimer
' End If
End Sub

'**********************************************************************************************************
' Update Solenoids
'**********************************************************************************************************


Sub UpdateSolenoids
  Dim Changed, Count, funcName, ii, sol11, solNo
  Changed = Controller.ChangedSolenoids
  If Not IsEmpty(Changed) Then
    sol11 = Controller.Solenoid(11)
    Count = UBound(Changed, 1)
    For ii = 0 To Count
      solNo = Changed(ii, CHGNO)
      ' multiplex solenoid #11 fixed in VPM 1.52beta and newer
      if VPMversion < "01519901" then
        If SolNo < 11 And sol11 Then solNo = solNo + 32
      else
        ' no need to evaluate sol 11 anymore, VPM does it
        if SolNo > 50 then solNo = solNo - 18
      end if
      funcName = SolCallback(solNo)
      If funcName <> "" Then
        'debug.print solNo
        'debug.print funcName
        Execute funcName & " CBool(" & Changed(ii, CHGSTATE) &")"
      End If
    Next
  End If
End Sub

'**********************************************************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)     = "SolDTLBankUnhit 13,"
SolCallback(2)     = "SolDTLBankUnhit 14,"
SolCallback(3)     = "SolDTLBankUnhit 15,"
SolCallback(4)     = "SolDTLBankUnhit 16,"
SolCallback(5)     = "SolDTLBankUnhit 17,"
SolCallback(6)     = "SolDTLBankDropDown"
SolCallback(7)     = "SolDTBPodDropUp"
SolCallback(8)     = "SolRelease"
SolCallback(9)     = "SolDTBait1"
SolCallback(10)    = "SolDTBait3"
SolCallback(12)    = "SolDrain"
SolCallback(13)    = "SolAPlunger"
SolCallback(14)    = "SolPFGI" 'PF GI
SolCallback(15)    = "SolKnocker"
'SolCallback(17)     = "" 'Left pop bumper
'SolCallback(18)     = "" 'Right pop bumper
SolCallback(21)    = "SolCenterFlasher"
SolCallback(22)    = "SolFlipperDiverter"
SolCallback(25)    = "SolRun"
SolCallback(33)    = "SolDTRBankUnhit 23,"
SolCallback(34)    = "SolDTRBankUnhit 24,"
SolCallback(35)    = "SolDTRBankUnhit 25,"
SolCallback(36)    = "SolDTRBankUnhit 26,"
SolCallback(37)    = "SolDTRBankUnhit 27,"
SolCallback(38)    = "SolDTRBankDropDown"
SolCallback(39)    = "SolDTTPodDropUp"
SolCallback(40)    = "SolUnlock"
SolCallback(41)    = "SolDTBait2"
SolCallback(42)    = "SolDTBaitDropDown"


'SolCallback(sLRFlipper) = ""       'Fast flip workaround, see keydown and keyup subs
'SolCallback(sLLFlipper) = ""
'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sULFlipper) = ""

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

Sub SolULFlipper(Enabled)
  If Enabled Then
    LeftFlipper1.RotateToEnd
    If leftflipper1.currentangle < leftflipper1.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper1
    Else
      SoundFlipperUpAttackLeft LeftFlipper1
      RandomSoundFlipperUpLeft LeftFlipper1
    End If
  Else
    LeftFlipper1.RotateToStart
    If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper1
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire 'rightflipper.rotatetoend

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



'**********************************************************************************************************
'Solenoid Controlled toys
'**********************************************************************************************************

dim FlippersEnabled

Sub SolRun(enabled)
  FlippersEnabled = Enabled
  if enabled then
    GameRun = True
    LeftSlingShot.Disabled = 0
    RightSlingShot.Disabled = 0
  Elseif not Enabled Then
    GameRun = false
    LeftSlingShot.Disabled = 1
    RightSlingShot.Disabled = 1
    SolLFlipper 0
    SolRFlipper 0
  end if
End Sub

Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub

dim gilvl : gilvl = 1
Sub SolPFGI(Enabled)
  If Enabled Then
    SetLamp 0,0
    Sound_GI_Relay 0,sw46
    gilvl = 0
    If UseFlexDMD Then
      Select Case FlexDMDStyle
        Case 0: Backglass.ImageA = "Defender Dark FlexDMD"
        Case 1: Backglass.ImageA = "Defender Dark"
        Case 2: Backglass.ImageA = "Defender Dark FlexDMD"
      End Select
    Else
      Backglass.ImageA = "Defender Dark"
    End if
  Else
    SetLamp 0,1
    Sound_GI_Relay 1,sw46
    gilvl = 1
    If UseFlexDMD Then
      Select Case FlexDMDStyle
        Case 0: Backglass.ImageA = "Defender Illuminated FlexDMD"
        Case 1: Backglass.ImageA = "Defender Illuminated"
        Case 2: Backglass.ImageA = "Defender Illuminated FlexDMD"
      End Select
    Else
      Backglass.ImageA = "Defender Illuminated"
    End if
  End if

End Sub

Sub SolAPlunger(enabled)
  If enabled Then
    Plunger1.Fire
    SoundPlunger1ReleaseBall
  Else
    Plunger1.PullBack
  End If
End Sub

Sub SolFlipperDiverter(enabled)
  If Enabled Then
    Flipper3.RotateToEnd
    PlungerGuideRamp4.collidable = 0
  Else
    Flipper3.RotateToStart
    PlungerGuideRamp4.collidable = 1
  End If
End Sub

Sub SolCenterFlasher(Enabled)
  If Enabled Then
    Lampz.state(100) = 1
  Else
    Lampz.state(100) = 0
  End If
End Sub




'******************************************************
' TROUGH
'******************************************************

Sub sw47_Hit   : Controller.Switch(47) = 1 : UpdateTrough : End Sub
Sub sw47_UnHit : Controller.Switch(47) = 0 : UpdateTrough : End Sub
Sub sw48_Hit   : Controller.Switch(48) = 1 : UpdateTrough : End Sub
Sub sw48_UnHit : Controller.Switch(48) = 0 : UpdateTrough : End Sub
Sub sw49_Hit   : Controller.Switch(49) = 1 : UpdateTrough : End Sub
Sub sw49_UnHit : Controller.Switch(49) = 0 : UpdateTrough : End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw47.BallCntOver = 0 Then sw48.kick 60, 20
  If sw48.BallCntOver = 0 Then sw49.kick 60, 20
  Me.Enabled = 0
End Sub

'******************************************************
' DRAIN & RELEASE
'******************************************************

Sub sw46_Hit
  RandomSoundDrain sw46
  Controller.Switch(46) = 1
End Sub

Sub SolDrain(enabled)
  If enabled Then
    sw46.kick 60, 20
    Controller.Switch(46) = 0
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw47.kick 56, 9
    RandomSoundBallRelease sw47
  End If
End Sub


'******************************************************
' LOCKUP
'******************************************************

Sub sw42_Hit   : Controller.Switch(42) = 1 : UpdateLock : End Sub
Sub sw42_UnHit : Controller.Switch(42) = 0 : UpdateLock : End Sub
Sub sw43_Hit   : Controller.Switch(43) = 1 : UpdateLock : End Sub
Sub sw43_UnHit : Controller.Switch(43) = 0 : UpdateLock : End Sub
Sub sw44_Hit   : Controller.Switch(44) = 1 : UpdateLock :  SoundSaucerLock: End Sub
Sub sw44_UnHit : Controller.Switch(44) = 0 : UpdateLock : End Sub

Sub UpdateLock
  UpdateLockTimer.Interval = 300
  UpdateLockTimer.Enabled = 1
End Sub

Sub UpdateLockTimer_Timer
  If sw42.BallCntOver = 0 Then sw43.kick 180, 10
  If sw43.BallCntOver = 0 Then sw44.kick 180, 10
  Me.Enabled = 0
End Sub

'******************************************************
' RELEASE LOCK
'******************************************************

Sub SolUnlock(enabled)
  If enabled Then
    sw42.kick 180, 1
    SoundSaucerKick 1, sw42
  End If
End Sub



'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

Dim DefenderBall1, DefenderBall2, DefenderBall3, gBOT

Sub Table1_Init

  FlexDMD_Init
  vpmInit Me
  On Error Resume Next
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Defender Williams"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 1
    If UseFlexDMD Then ExternalEnabled = Controller.Games(cGameName).Settings.Value("showpindmd")
    If UseFlexDMD Then Controller.Games(cGameName).Settings.Value("showpindmd") = 0
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With
  On Error Goto 0

  VPMversion=Controller.version

  If UseFlexDMD Then NvRam_Init

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch = swTilt
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftslingShot,RightslingShot)

  '************  Trough **************
  Set DefenderBall3 = sw49.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set DefenderBall2 = sw48.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set DefenderBall1 = sw47.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(DefenderBall1,DefenderBall2,DefenderBall3)

  Controller.Switch(49) = 1
  Controller.Switch(48) = 1
  Controller.Switch(47) = 1

  Plunger1.Pullback

  If VR_Room = 1 Then
    SetBackglass
  End If

End Sub

Sub Table1_Exit()
  Controller.Pause = False
  Controller.Stop
  If UseFlexDMD then
    If Not FlexDMD is Nothing Then
      FlexDMD.Show = False
      FlexDMD.Run = False
      FlexDMD = NULL
    End if
    Controller.Games(cGameName).Settings.Value("showpindmd") = ExternalEnabled
  End if
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
Dim BIPL : BIPL=0

Sub Table1_KeyDown(ByVal KeyCode)

  If KeyCode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull
    TimerVRPlunger.enabled = true
    TimerVRPlunger2.enabled = False
  End If

  If keycode = RightFlipperKey and GameRun Then
    Controller.Switch(swLRFlip) = True
    FlipperActivate RightFlipper, RFPress
    SolRFlipper 1       'emulating fast flips
  End If

  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = VRFlipperButtonRight.X -8
  End If

  If keycode = LeftFlipperKey and GameRun Then
    Controller.Switch(swLLFlip) = True
    FlipperActivate LeftFlipper, LFPress
    SolLFlipper 1       'emulating fast flips
    If StagedFlippers <> 1 Then
      Controller.Switch(77) = True
      FlipperActivate LeftFlipper1, ULFPress
      SolULFlipper 1       'emulating fast flips
    End If
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = VRFlipperButtonLeft.X +8
  End If

  If StagedFlippers = 1 Then
    If keycode = KeyUpperLeft and GameRun Then
      Controller.Switch(77) = True
      FlipperActivate LeftFlipper1, ULFPress
      SolULFlipper 1       'emulating fast flips
    End If
  End If

  If keycode = LeftMagnaSave Then
    Controller.Switch(61) = True
    VRFlipperButtonLeftMagna.X = VRFlipperButtonLeftMagna.X +8
  End If
  If keycode = RightMagnaSave Then
    Controller.Switch(60) = True
    VRFlipperButtonRightMagna.X = VRFlipperButtonRightMagna.X -8
    If Lampz.State(49) = 1 Then SmartBombUsed = 1
  End If
  If keycode = LockBarKey Then Controller.Switch(60) = True

  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25

    End Select
  End If
  if keycode = StartGameKey then
    soundStartButton
    startbutton.y = startbutton.y -3
  End If

  If KeyDownHandler(keycode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal KeyCode)

  If keycode = RightFlipperKey and GameRun Then
    Controller.Switch(swLRFlip) = False
    FlipperDeActivate RightFlipper, RFPress
    SolRFlipper 0       'emulating fast flips
  End If

  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = VRFlipperButtonRight.X +8
  End If

  If keycode = LeftFlipperKey and GameRun Then
    Controller.Switch(swLLFlip) = False
    FlipperDeActivate LeftFlipper, LFPress
    SolLFlipper 0       'emulating fast flips
    If StagedFlippers <> 1 Then
      Controller.Switch(77) = False
      FlipperDeActivate LeftFlipper1, ULFPress
      SolULFlipper 0       'emulating fast flips
    End If
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = VRFlipperButtonLeft.X -8
  End If

  If StagedFlippers = 1 Then
    If keycode = KeyUpperLeft and GameRun Then
      Controller.Switch(77) = False
      FlipperDeActivate LeftFlipper1, ULFPress
      SolULFlipper 0       'emulating fast flips
    End If
  End If

  If keycode = LeftMagnaSave Then
    Controller.Switch(61) = False
    VRFlipperButtonLeftMagna.X = VRFlipperButtonLeftMagna.X -8
  End If

  If keycode = RightMagnaSave Then
    Controller.Switch(60) = False
    VRFlipperButtonRightMagna.X = VRFlipperButtonRightMagna.X +8
  End If
  If keycode = LockBarKey Then Controller.Switch(60) = False

  If KeyCode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall                        'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall                        'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.enabled = false
    TimerVRPlunger2.enabled = true
  End If

  If Keycode = StartGameKey Then
    StartButton.y = StartButton.y +3
  End If

  If KeyUpHandler(keycode) Then Exit Sub

End Sub


'***************************************************************************
' VR Plunger Code
'***************************************************************************

TimerVRPlunger2.enabled = true

Sub TimerVRPlunger_Timer
  if VRPlunger.Y < 2240 then VRPlunger.Y = VRPlunger.y +6  'If the plunger is not fully extend it, then extend it by 5 coordinates in the Y,
End Sub

Sub TimerVRPlunger2_Timer
  VRPlunger.Y = 2129 + (5* Plunger.Position) -20
end sub


'***************************************************************************
'       Drop Target Controls
'***************************************************************************

Sub sw33_Hit : DTHit 33 : TargetBouncer Activeball, 1 : End Sub
Sub sw34_Hit : DTHit 34 : TargetBouncer Activeball, 1 : End Sub
Sub sw35_Hit : DTHit 35 : TargetBouncer Activeball, 1 : End Sub
Sub sw39_Hit : DTHit 39 : TargetBouncer Activeball, 1 : End Sub
Sub sw40_Hit : DTHit 40 : TargetBouncer Activeball, 1 : End Sub

Sub sw13_Hit : DTHit 13 : TargetBouncer Activeball, 1 : End Sub
Sub sw14_Hit : DTHit 14 : TargetBouncer Activeball, 1 : End Sub
Sub sw15_Hit : DTHit 15 : TargetBouncer Activeball, 1 : End Sub
Sub sw16_Hit : DTHit 16 : TargetBouncer Activeball, 1 : End Sub
Sub sw17_Hit : DTHit 17 : TargetBouncer Activeball, 1 : End Sub

Sub sw23_Hit : DTHit 23 : TargetBouncer Activeball, 1 : End Sub
Sub sw24_Hit : DTHit 24 : TargetBouncer Activeball, 1 : End Sub
Sub sw25_Hit : DTHit 25 : TargetBouncer Activeball, 1 : End Sub
Sub sw26_Hit : DTHit 26 : TargetBouncer Activeball, 1 : End Sub
Sub sw27_Hit : DTHit 27 : TargetBouncer Activeball, 1 : End Sub


' Left Center Drop Target
Sub SolDTBait1(enabled)
  If enabled Then
    DTRaise 33
  End If
End Sub

' Right Center Drop Target
Sub SolDTBait2(enabled)
  If enabled Then
    RandomSoundDropTargetReset sw33p
    DTRaise 34
  End If
End Sub

' Right Top Drop Target
Sub SolDTBait3(enabled)
  If enabled Then
    RandomSoundDropTargetReset sw35p
    DTRaise 35
  End If
End Sub

Sub SolDTBaitDropDown(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors), Light56
    DTDrop 33
    DTDrop 34
    DTDrop 35
  end if
End Sub

' Bottom Pod Top Drop Target
Sub SolDTBPodDropUp(enabled)
  If enabled Then
    RandomSoundDropTargetReset sw39p
    DTRaise 39
  End If
End Sub

' Top Pod Top Drop Target
Sub SolDTTPodDropUp(enabled)
  If enabled Then
    RandomSoundDropTargetReset sw40p
    DTRaise 40
  End If
End Sub

' Raise specific target in left bank
Sub SolDTLBankUnhit(sw,enabled)
  If enabled Then
    RandomSoundDropTargetReset sw15p
    DTRaise sw
  End If
End Sub

' Left Bank of Drop Targets
Sub SolDTLBankDropDown(enabled)
  If enabled Then
    DTDrop 13
    DTDrop 14
    DTDrop 15
    DTDrop 16
    DTDrop 17
  End If
End Sub

' Raise specific target in right bank
Sub SolDTRBankUnhit(sw,enabled)
  If enabled Then
    RandomSoundDropTargetReset sw25p
    DTRaise sw
  End If
End Sub


' Right Bank of Drop Targets
Sub SolDTRBankDropDown(enabled)
  If enabled Then
    DTDrop 23
    DTDrop 24
    DTDrop 25
    DTDrop 26
    DTDrop 27
  End If
End Sub



'**********************************************************************************************************


'Wire Triggers
Sub sw9_Hit    : Controller.Switch(9)=1 : End Sub
Sub sw9_UnHit  : Controller.Switch(9)=0 : End Sub
Sub sw10_Hit   : Controller.Switch(10)=1 : End Sub
Sub sw10_UnHit : Controller.Switch(10)=0 : End Sub
Sub sw11_Hit   : Controller.Switch(11)=1 : End Sub
Sub sw11_UnHit : Controller.Switch(11)=0 : End Sub
Sub sw12_Hit   : Controller.Switch(12)=1 : End Sub
Sub sw12_UnHit : Controller.Switch(12)=0 : End Sub
Sub sw45_Hit   : Controller.Switch(45)=1 : End Sub
Sub sw45_UnHit : Controller.Switch(45)=0 : End Sub
Sub sw50_Hit   : Controller.Switch(50)=1 : BIPL=1 : End Sub
Sub sw50_UnHit : Controller.Switch(50)=0 : BIPL=0 :End Sub
Sub sw51_Hit   : Controller.Switch(51)=1 : End Sub
Sub sw51_UnHit : Controller.Switch(51)=0 : End Sub
Sub sw52_Hit   : Controller.Switch(52)=1 : End Sub
Sub sw52_UnHit : Controller.Switch(52)=0 : End Sub
Sub sw53_Hit   : Controller.Switch(53)=1 : End Sub
Sub sw53_UnHit : Controller.Switch(53)=0 : End Sub
Sub sw54_Hit   : Controller.Switch(54)=1 : End Sub
Sub sw54_UnHit : Controller.Switch(54)=0 : End Sub

'Stand Up Targets
Sub sw18_hit : STHit 18 : End Sub
Sub sw19_hit : STHit 19 : End Sub
Sub sw20_hit : STHit 20 : End Sub
Sub sw21_hit : STHit 21 : End Sub
Sub sw22_hit : STHit 22 : End Sub
Sub sw28_hit : STHit 28 : End Sub
Sub sw29_hit : STHit 29 : End Sub
Sub sw30_hit : STHit 30 : End Sub
Sub sw31_hit : STHit 31 : End Sub
Sub sw32_hit : STHit 32 : End Sub
Sub sw36_hit : STHit 36 : End Sub
Sub sw37_hit : STHit 37 : End Sub
Sub sw41_hit : STHit 41 : End Sub

'Scoring Rubbers
Sub sw38_hit:vpmTimer.pulseSw 38 : playsound"flip_hit_3" : End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(55) : RandomSoundBumperTop Bumper1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(56) : RandomSoundBumperBottom Bumper2: End Sub




'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(32)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)

' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)


'InitDisplay
'
'Sub InitDisplay
' Dim num,obj
' For num = 0 To UBound(Digits)
'   For Each obj In Digits(num)
'     obj.State = 0
'   Next
' Next
'End Sub


Sub DisplayTimer
  If UseFlexDMD then FlexDMD.LockRenderThread
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 32) then
        If UseFlexDMD And FlexDMDStyle <> 1 then
          UpdateFlexChar num, stat
        Else
          On Error Resume Next
          If cab_mode = 0 and VR_Room = 0 Then
            For Each obj In Digits(num)
              If chg And 1 Then obj.State = stat And 1
              chg = chg\2 : stat = stat\2
            Next
          Elseif VR_Room = 1 Then

            For Each obj In VRDigits(num)
              If chg And 1 Then FadeDisplay obj, stat And 1
              chg = chg\2 : stat=stat\2
            Next
          End If
          On Error GoTo 0
        End If
      else
      end if
    next
  end if

  If UseFlexDMD then
    Dim i
    With FlexDMDScene

      If Lampz.State(2) = 1 Then  'in game

        .GetImage("Defender").Visible = False

        Select Case SmartBombUsed
        Case 0
          .GetFrame("BackColour").FillColor = vbBlack
        Case 1
          .GetFrame("BackColour").FillColor = RGB(85,85,85)
          SmartBombUsed = SmartBombUsed + 1
        Case 2
          .GetFrame("BackColour").FillColor = RGB(170,170,170)
          SmartBombUsed = SmartBombUsed + 1
        Case 3
          .GetFrame("BackColour").FillColor = vbWhite
          SmartBombUsed = SmartBombUsed + 1
        Case 4
          .GetFrame("BackColour").FillColor = RGB(170,170,170)
          SmartBombUsed = SmartBombUsed + 1
        Case Else
          .GetFrame("BackColour").FillColor = RGB(85,85,85)
          SmartBombUsed = 0
        End Select


        If FlexDMDStyle <> 1 Then
          For i = 1 To 3
            Select Case cntTimer
              Case 0
                .GetImage("Bomb" & CStr(i)).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_SmartBomb1").Bitmap
              Case 3
                .GetImage("Bomb" & CStr(i)).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_SmartBomb2").Bitmap
              Case 6
                .GetImage("Bomb" & CStr(i)).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_SmartBomb3").Bitmap
              Case 9
                .GetImage("Bomb" & CStr(i)).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_SmartBomb3").Bitmap
              Case 12
                .GetImage("Bomb" & CStr(i)).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_SmartBomb5").Bitmap
              Case 15
                .GetImage("Bomb" & CStr(i)).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_SmartBomb6").Bitmap
            End Select
          Next

          .GetImage("Bomb1").Visible = Lampz.State(51)
          .GetImage("Bomb2").Visible = Lampz.State(50)
          .GetImage("Bomb3").Visible = Lampz.State(49)

          .GetImage("CreditsWave").Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_Wave").Bitmap
          .GetImage("Credit30").Visible = True
          .GetImage("Credit31").Visible = True

          If ActivePlayer(2) = 1 Then
            For i = 14 to 20
              .GetImage("SegL" & cstr(i)).Visible = True
              .GetImage("SegL" & cstr(i + 7)).Visible = False
            Next
          ElseIf ActivePlayer(2) = 2 Then
            For i = 14 to 20
              .GetImage("SegL" & cstr(i)).Visible = False
              .GetImage("SegL" & cstr(i + 7)).Visible = True
            Next
          End If
          .GetImage("Fighter1").Visible = False
          .GetImage("Fighter2").Visible = False
          .GetImage("Fighter3").Visible = False
          .GetImage("Fighter4").Visible = False
          .GetImage("Fighter5").Visible = False

          'work out how many fighters (balls left) to display
          Dim ballsleft
          ballsleft = BallsPerGame - CInt(Join(BallNumber,"")) + CInt(Lampz.State(1))
          For i = 1 To ballsleft
            Select Case cntTimer
            Case 0,9
              .GetImage("Fighter" & CStr(i)).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_Fighter1").Bitmap
            Case 3,12
              .GetImage("Fighter" & CStr(i)).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_Fighter2").Bitmap
            Case 6,15
              .GetImage("Fighter" & CStr(i)).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_Fighter3").Bitmap
            End Select
            .GetImage("Fighter" & CStr(i)).Visible = True
            If i = 5 Then Exit For
          Next

          .GetImage("HighScore").Visible = False
          .GetImage("HighScoreOver").Visible = False
          .GetImage("GameOver").Visible = False
          .GetImage("Tilt").Visible = Lampz.State(3)
        End If


      Else  'attract
        If FlexDMDStyle <> 1 Then
          .GetImage("Bomb1").Visible = False
          .GetImage("Bomb2").Visible = False
          .GetImage("Bomb3").Visible = False

          If FreePlayEnabled = False Then
            .GetImage("CreditsWave").Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_Credits").Bitmap
            .GetImage("Credit30").Visible = True
            .GetImage("Credit31").Visible = True
          Else
            .GetImage("CreditsWave").Bitmap = FlexDMD.NewImage("", "VPX.DMD_Image_FreePlay").Bitmap
            .GetImage("Credit30").Visible = False
            .GetImage("Credit31").Visible = False
          End If

          .GetImage("Fighter1").Visible = False
          .GetImage("Fighter2").Visible = False
          .GetImage("Fighter3").Visible = False
          .GetImage("Fighter4").Visible = False
          .GetImage("Fighter5").Visible = False
          For i = 14 to 20
            .GetImage("SegL" & cstr(i)).Visible = False
            .GetImage("SegL" & cstr(i + 7)).Visible = False
            .GetImage("Seg" & cstr(i)).Visible = True
            .GetImage("Seg" & cstr(i + 7)).Visible = True
          Next

          .GetFrame("BackColour").FillColor = vbBlack

          .GetImage("HighScore").Visible = Lampz.State(6)
          .GetImage("HighScoreOver").Visible = Lampz.State(6)
          .GetImage("GameOver").Visible = True ' Lampz.State(4)
          .GetImage("Tilt").Visible = False

        End If


        If FlexDMDStyle = 1 Then .GetImage("Defender").Visible = True
      End If

      If FlexDMDStyle <> 1 Then
        If Lampz.State(5) = 1 Then
          .GetImage("BallMatch").Visible = True
          .GetImage("Ball28").Visible = True
          .GetImage("Ball29").Visible = True
        Else
          .GetImage("BallMatch").Visible = False
          .GetImage("Ball28").Visible = False
          .GetImage("Ball29").Visible = False
        End If
      End If

      If FlexDMDStyle > 0 And SmartBombUsed = 0 Then
        UpdateFlexScanner Lampz.State(2), cntTimer
      End If

    End With
    cntTimer = cntTimer + 1
    If cntTimer = 18 Then cntTimer = 0
    FlexDMD.UnlockRenderThread
  End If

End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 58
  RandomSoundSlingshotRight SLING1
  RSling.Visible = 0
  RSling1.Visible = 1
  sling1.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 2:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
    Case 3:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
  End Select
  RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 57
  RandomSoundSlingshotLeft SLING2
  LSling.Visible = 0
  LSling1.Visible = 1
  sling2.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 2:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
    Case 3:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
  End Select
  LStep = LStep + 1
End Sub




'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************
Const tnob = 3
Const lob = 0

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

Sub UpdateRolling
  Dim b

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollingVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ < -1 and gBOT(b).z < 55 and gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz > -7 Then
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
    If AmbientBallShadowOn = 0 Then
      If gBOT(b).Z > 30 Then
        BallShadowA(b).height=gBOT(b).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
      Else
        BallShadowA(b).height=gBOT(b).z - BallSize/2 + 5
      End If
      BallShadowA(b).Y = gBOT(b).Y + Ballsize/5 + fovY
      BallShadowA(b).X = gBOT(b).X
      BallShadowA(b).visible = 1
    End If
  Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************





'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

Const fovY          = 0   'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 15  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source


Dim sourcenames, currentShadowCount, DSSources(30), numberofsources, numberofsources_hold
sourcenames = Array ("","","","")
currentShadowCount = Array (0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(3), objrtx2(3)
dim objBallShadow(3)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3)

DynamicBSInit

sub DynamicBSInit
  Dim iii, source

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 0.21
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 0.22
    objrtx2(iii).visible = 0

    currentShadowCount(iii) = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = iii/1000 + 0.24
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    iii = iii + 1
  Next
  numberofsources = iii
  numberofsources_hold = iii
end sub


Sub DynamicBSUpdate
  Dim falloff:  falloff = 150     'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, currentMat, AnotherSource, iii

  'Hide shadow of deleted balls
  For s = UBound(gBOT) + 1 to tnob
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(gBOT) < lob Then Exit Sub   'No balls in play, exit

  'The Magic happens now
  For s = lob to UBound(gBOT)

  ' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If gBOT(s).Z > 30 Then              'The flasher follows the ball up ramps while the primitive is on the pf
        If gBOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((gBOT(s).X) + (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = gBOT(s).Y + BallSize/10 + fovY
        objBallShadow(s).visible = 1

        BallShadowA(s).X = gBOT(s).X
        BallShadowA(s).Y = gBOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=gBOT(s).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then  'On pf, primitive only
        objBallShadow(s).visible = 1
        If gBOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((gBOT(s).X) + (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = gBOT(s).Y + fovY
        BallShadowA(s).visible = 0
      Else                      'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If gBOT(s).Z > 30 Then              'In a ramp
        BallShadowA(s).X = gBOT(s).X
        BallShadowA(s).Y = gBOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=gBOT(s).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If gBOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          BallShadowA(s).X = ((gBOT(s).X) + (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        BallShadowA(s).Y = gBOT(s).Y + Ballsize/10 + fovY
        BallShadowA(s).height=gBOT(s).z - BallSize/2 + 5
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

  ' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If gBOT(s).Z < 30 Then 'And gBOT(s).Y < (TableHeight - 200) Then 'Or gBOT(s).Z > 105 Then   'Defining when and where (on the table) you can have dynamic shadows
        For iii = 0 to numberofsources - 1
          LSd=Distance(gBOT(s).x,gBOT(s).y,DSSources(iii)(0),DSSources(iii)(1)) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then               'If the ball is within the falloff range of a light and light is on (we will set numberofsources to 0 when GI is off)
            currentShadowCount(s) = currentShadowCount(s) + 1   'Within range of 1 or 2
            if currentShadowCount(s) = 1 Then           '1 dynamic shadow source
              sourcenames(s) = iii
              currentMat = objrtx1(s).material
              objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = gBOT(s).X : objrtx1(s).Y = gBOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01            'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), gBOT(s).X, gBOT(s).Y) + 90
              ShadowOpacity = (falloff-LSd)/falloff                 'Sets opacity/darkness of shadow by distance to light
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness           'Scales shape of shadow with distance/opacity
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^2,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-ShadowOpacity)
              End If

            Elseif currentShadowCount(s) = 2 Then
                                  'Same logic as 1 shadow, but twice
              currentMat = objrtx1(s).material
              AnotherSource = sourcenames(s)
              objrtx1(s).visible = 1 : objrtx1(s).X = gBOT(s).X : objrtx1(s).Y = gBOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(AnotherSource)(0),DSSources(AnotherSource)(1), gBOT(s).X, gBOT(s).Y) + 90
              ShadowOpacity = ( falloff - Distance(gBOT(s).x,gBOT(s).y,DSSources(AnotherSource)(0),DSSources(AnotherSource)(1)) )/falloff
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

              currentMat = objrtx2(s).material
              objrtx2(s).visible = 1 : objrtx2(s).X = gBOT(s).X : objrtx2(s).Y = gBOT(s).Y + fovY
  '           objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx2(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), gBOT(s).X, gBOT(s).Y) + 90
              ShadowOpacity2 = (falloff-LSd)/falloff
              objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2))
              End If
            end if
          Else
            currentShadowCount(s) = 0
            BallShadowA(s).Opacity = 100*AmbientBSFactor
          End If
        Next
      Else                  'Hide dynamic shadows everywhere else
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub
'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************



'**********************************************************************************************************
'**********************************************************************************************************



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
PlungerReleaseSoundLevel = 0.8 '1 wjr                 'volume level; range [0, 1]
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
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, -1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound playsoundparams, -1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
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
  Playsound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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
    AudioFade = Csng(tmp ^5) 'was 10
  Else
    AudioFade = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / tablewidth-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^5) ' was 10
  Else
    AudioPan = Csng(-((- tmp) ^5) ) ' was 10
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
  RndInt = Int(Rnd * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
  RndNum = Rnd * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub


Sub SoundPlungerPull
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlunger1ReleaseBall
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger1
End Sub

Sub SoundPlungerReleaseNoBall
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid
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
Sub RandomSoundRollover
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
    RandomSoundRubberWeak
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
Sub RandomSoundRubberWeak
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  RandomSoundWall
End Sub

Sub RandomSoundWall
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
Sub RandomSoundMetal
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
Sub RandomSoundBottomArchBallGuide
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
Sub RandomSoundBottomArchBallGuideHardHit
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
    RandomSoundBottomArchBallGuideHardHit
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide
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
Sub RandomSoundTargetHitStrong
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong
    RandomSoundBallBouncePlayfieldSoft Activeball
  Else
    RandomSoundTargetHitWeak
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

Sub SoundPlayfieldGate
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit
  If activeball.velx < -8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit
  If Activeball.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit
  If activeball.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock
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

  FlipperCradleCollision ball1, ball2, velocity
End Sub



'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315                  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05                  'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025*RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025*RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025*RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025*RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////



'******************************************************
' SLINGSHOT CORRECTION FUNCTIONS by apophis
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

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
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



'*******************************************
'     FlippersPol  late 70s to early 80s
'*******************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger


        x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, - 2.7
        x.AddPt "Polarity", 2, 0.16, - 2.7
        x.AddPt "Polarity", 3, 0.22, - 0
        x.AddPt "Polarity", 4, 0.25, - 0
        x.AddPt "Polarity", 5, 0.3, - 1
        x.AddPt "Polarity", 6, 0.4, - 2
        x.AddPt "Polarity", 7, 0.5, - 2.7
        x.AddPt "Polarity", 8, 0.65, - 1.8
        x.AddPt "Polarity", 9, 0.75, - 0.5
        x.AddPt "Polarity", 10, 0.81, - 0.5
        x.AddPt "Polarity", 11, 0.88, 0
        x.AddPt "Polarity", 12, 1.3, 0

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
  '   Dim BOT
  '   BOT = GetBalls

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

Function max(a,b)
  if a > b Then max=a Else max=b
End Function

Function min(a,b)
  if a >= b Then min=b Else min=a
End Function

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

Dim LFPress, RFPress, LFCount, RFCount, ULFPress
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
'Const EOSReturn = 0.055  'EM's
Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b', BOT
    'BOT = GetBalls

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
'****  END FLIPPER CORRECTIONS
'******************************************************



'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
'
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR



Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

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
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
      aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report         'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub

End Class



'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7 when TargetBouncerEnabled=1

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
'   DROP TARGETS INITIALIZATION
'******************************************************

Class DropTarget
  Private m_primary, m_secondary, m_prim, m_primoff, m_sw, m_animate, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get PrimOff(): Set PrimOff = m_primoff: End Property
  Public Property Let PrimOff(input): Set m_primoff = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, primoff, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    Set m_primoff = primoff
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class

'Define a variable for each drop target
Dim DT33, DT34, DT35, DT39, DT40
Dim DT13, DT14, DT15, DT16, DT17
Dim DT23, DT24, DT25, DT26, DT27

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
'
'       isDropped:  Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.

Set DT33 = (new DropTarget)(sw33, sw33a, sw33p, sw33poff, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34a, sw34p, sw34poff, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, sw35p, sw35poff, 35, 0, false)
Set DT39 = (new DropTarget)(sw39, sw39a, sw39p, sw39poff, 39, 0, false)
Set DT40 = (new DropTarget)(sw40, sw40a, sw40p, sw40poff, 40, 0, false)
Set DT13 = (new DropTarget)(sw13, sw13a, sw13p, sw13poff, 13, 0, false)
Set DT14 = (new DropTarget)(sw14, sw14a, sw14p, sw14poff, 14, 0, false)
Set DT15 = (new DropTarget)(sw15, sw15a, sw15p, sw15poff, 15, 0, false)
Set DT16 = (new DropTarget)(sw16, sw16a, sw16p, sw16poff, 16, 0, false)
Set DT17 = (new DropTarget)(sw17, sw17a, sw17p, sw17poff, 17, 0, false)
Set DT23 = (new DropTarget)(sw23, sw23a, sw23p, sw23poff, 23, 0, false)
Set DT24 = (new DropTarget)(sw24, sw24a, sw24p, sw24poff, 24, 0, false)
Set DT25 = (new DropTarget)(sw25, sw25a, sw25p, sw25poff, 25, 0, false)
Set DT26 = (new DropTarget)(sw26, sw26a, sw26p, sw26poff, 26, 0, false)
Set DT27 = (new DropTarget)(sw27, sw27a, sw27p, sw27poff, 27, 0, false)


Dim DTArray
DTArray = Array(DT33, DT34, DT35, DT39, DT40, DT13, DT14, DT15, DT16, DT17, DT23, DT24, DT25, DT26, DT27)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 50 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.3 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

  'PlaySoundAtVol  DTHitSound, Activeball, Vol(Activeball)*22.5
  PlayTargetSound
  DTArray(i).animate =  DTCheckBrick(Activeball,DTArray(i).prim)
  If DTArray(i).animate = 1 or DTArray(i).animate = 3 or DTArray(i).animate = 4 Then
    DTBallPhysics Activeball, DTArray(i).prim.rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)
  DTArray(i).animate = -1
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
  For i = 0 to uBound(DTArray)
    If DTArray(i).sw = switch Then DTArrayID = i:Exit Function
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

Sub DTAnim_Timer
  DoDTAnim
  DoSTAnim
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    If DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function


Sub DoDTAnim
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).primoff,DTArray(i).sw,DTArray(i).animate)
  Next
End Sub

Function DTAnimate(primary, secondary, prim, primoff, switch, animate)
  dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If (animate = 1 or animate = 4) and animtime < DTDropDelay Then
    primary.collidable = 0
  If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    primoff.rotx = prim.rotx
    primoff.roty = prim.roty
    DTAnimate = animate
    Exit Function
    elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    primoff.rotx = prim.rotx
    primoff.roty = prim.roty
    animate = 2
    SoundDropTargetDrop prim
  End If

  if animate = 2 Then
    transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
    if prim.transz > -DTDropUnits  Then
      prim.transz = transz
      primoff.transz = prim.transz
    end if

    prim.rotx = DTMaxBend * cos(rangle)/2
    prim.roty = DTMaxBend * sin(rangle)/2
    primoff.rotx = prim.rotx
    primoff.roty = prim.roty

    if prim.transz <= -DTDropUnits Then
      prim.transz = -DTDropUnits
      primoff.transz = prim.transz
      secondary.collidable = 0
      controller.Switch(Switchid) = 1
      DTShadowHide switchid
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
    primoff.rotx = prim.rotx
    primoff.roty = prim.roty
  elseif animate = 3 and animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primoff.rotx = prim.rotx
    primoff.roty = prim.roty
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  if animate = -1 Then
    transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

    If prim.transz = -DTDropUnits Then
      Dim b

      For b = 0 to UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and gBOT(b).z < prim.z+DTDropUnits+25 Then
          gBOT(b).velz = 20
        End If
      Next
    End If

    if prim.transz < 0 Then
      prim.transz = transz
      primoff.transz = prim.transz
    elseif transz > 0 then
      prim.transz = transz
      primoff.transz = prim.transz
    end if

    if prim.transz > DTDropUpUnits then
      DTAnimate = -2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primoff.transz = prim.transz
      primoff.rotx = prim.rotx
      primoff.roty = prim.roty
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1
    controller.Switch(Switchid) = 0
    DTShadowShow switchid
  End If

  if animate = -2 and animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
    primoff.transz = prim.transz
    if prim.transz < 0 then
      prim.transz = 0
      primoff.transz = prim.transz
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    end If
  End If
End Function

Sub DTShadowHide(switchid)
  Select Case switchid
    Case 13: ShadowDT(0).visible=False
    Case 14: ShadowDT(1).visible=False
    Case 15: ShadowDT(2).visible=False
    Case 16: ShadowDT(3).visible=False
    Case 17: ShadowDT(4).visible=False
    Case 23: ShadowDT(5).visible=False
    Case 24: ShadowDT(6).visible=False
    Case 25: ShadowDT(7).visible=False
    Case 26: ShadowDT(8).visible=False
    Case 27: ShadowDT(9).visible=False
  End Select
End Sub

Sub DTShadowShow(switchid)
  Select Case switchid
    Case 13: ShadowDT(0).visible=True
    Case 14: ShadowDT(1).visible=True
    Case 15: ShadowDT(2).visible=True
    Case 16: ShadowDT(3).visible=True
    Case 17: ShadowDT(4).visible=True
    Case 23: ShadowDT(5).visible=True
    Case 24: ShadowDT(6).visible=True
    Case 25: ShadowDT(7).visible=True
    Case 26: ShadowDT(8).visible=True
    Case 27: ShadowDT(9).visible=True
  End Select
End Sub



'******************************************************
'   DROP TARGET
'   SUPPORTING FUNCTIONS
'******************************************************


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

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
    Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
    Dim rotxy
    rotxy = RotPoint(ax,ay,angle)
    rax = rotxy(0)+px : ray = rotxy(1)+py
    rotxy = RotPoint(bx,by,angle)
    rbx = rotxy(0)+px : rby = rotxy(1)+py
    rotxy = RotPoint(cx,cy,angle)
    rcx = rotxy(0)+px : rcy = rotxy(1)+py
    rotxy = RotPoint(dx,dy,angle)
    rdx = rotxy(0)+px : rdy = rotxy(1)+py

    InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function




'******************************************************
'   STAND-UP TARGET INITIALIZATION
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
Dim ST18, ST19, ST20, ST21, ST22
Dim ST28, ST29, ST30, ST31, ST32
Dim ST36, ST37, ST41

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:      vp target to determine target hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             transy must be used to offset the target animation
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST18 = (new StandupTarget)(sw18, sw18p, 18, 0)
Set ST19 = (new StandupTarget)(sw19, sw19p, 19, 0)
Set ST20 = (new StandupTarget)(sw20, sw20p, 20, 0)
Set ST21 = (new StandupTarget)(sw21, sw21p, 21, 0)
Set ST22 = (new StandupTarget)(sw22, sw22p, 22, 0)
Set ST28 = (new StandupTarget)(sw28, sw28p, 28, 0)
Set ST29 = (new StandupTarget)(sw29, sw29p, 29, 0)
Set ST30 = (new StandupTarget)(sw30, sw30p, 30, 0)
Set ST31 = (new StandupTarget)(sw31, sw31p, 31, 0)
Set ST32 = (new StandupTarget)(sw32, sw32p, 32, 0)
Set ST36 = (new StandupTarget)(sw36, sw36p, 36, 0)
Set ST37 = (new StandupTarget)(sw37, sw37p, 37, 0)
Set ST41 = (new StandupTarget)(sw41, sw41p, 41, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST18, ST19, ST20, ST21, ST22, ST28, ST29, ST30, ST31, ST32, ST36, ST37, ST41)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5         'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  PlayTargetSound
  STArray(i).animate =  STCheckHit(Activeball,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics Activeball, STArray(i).primary.orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 to uBound(STArray)
    If STArray(i).sw = switch Then STArrayID = i:Exit Function
  Next
End Function

'Check if target is hit on it's face
Function STCheckHit(aBall, target)
  dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    STCheckHit = 1
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    STCheckHit = 1
  Else
    STCheckHit = 0
  End If
End Function

Sub DoSTAnim
  Dim i
  For i=0 to Ubound(STArray)
    STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
  Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
  Dim animtime

  STAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    STAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transy = -STMaxOffset
    vpmTimer.PulseSw switch
    STAnimate = 2
    Exit Function
  elseif animate = 2 Then
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
'   END STAND-UP TARGETS
'******************************************************




'******************************************************
'****  LAMPZ by nFozzy
'******************************************************
'
' Lampz is a utility designed to manage and fade the lights and light-related objects on a table that is being driven by a ROM.
' To set up Lampz, one must populate the Lampz.MassAssign array with VPX Light objects, where the index of the MassAssign array
' corrisponds to the ROM index of the associated light. More that one Light object can be associated with a single MassAssign index (not shown in this example)
' Optionally, callbacks can be assigned for each index using the Lampz.Callback array. This is very useful for allowing 3D Insert primitives
' to be controlled by the ROM. Note, the aLvl parameter (i.e. the fading level that ranges between 0 and 1) is appended to the callback call.


Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
InitLampsNF              ' Setup lamp assignments
LampTimer.Interval = 15
LampTimer.Enabled = 1

Sub LampTimer_Timer
  dim x, chglamp
  chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
    next
  End If
  Lampz.Update1 'update (fading logic only)
End Sub

dim FrameTime, InitFrameTime : InitFrameTime = 0
LampTimer2.Interval = -1
LampTimer2.Enabled = 1
Sub LampTimer2_Timer
  FrameTime = gametime - InitFrameTime : InitFrameTime = gametime 'Count frametime. Unused atm?
  Lampz.Update 'updates on frametime (Object updates only)
End Sub


'Fade material for green, red, yellow colored Bulb prims
Sub FadeMeshMaterial(pri, group, ByVal aLvl)  'cp's script
  ' if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  Select case FlashLevelToIndex(aLvl, 3)
    Case 0:pri.Material = group(0) 'Off
    Case 1:pri.Material = group(1) 'Fading...
    Case 2:pri.Material = group(2) 'Fading...
    Case 3:pri.Material = group(3) 'Full
  End Select
  'if tb.text <> pri.image then tb.text = pri.image : 'debug.print pri.image end If 'debug
  pri.blenddisablelighting = aLvl * 1 'Intensity Adjustment
End Sub



Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity
End Sub




Sub InitLampsNF

  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

  'Adjust fading speeds (1 / full MS fading time)
  dim x : for x = 0 to 140 : Lampz.FadeSpeedUp(x) = 1/2 : Lampz.FadeSpeedDown(x) = 1/10 : next
  Lampz.FadeSpeedUp(0) = 1/6 : Lampz.FadeSpeedDown(0) = 1/18    'GI
  Lampz.FadeSpeedUp(100) = 1/2 : Lampz.FadeSpeedDown(0) = 1/8   'Center flasher

  'Lampz Assignments
  '  In a ROM based table, the lamp ID is used to set the state of the Lampz objects

  'MassAssign is an optional way to do assignments. It'll create arrays automatically / append objects to existing arrays
  Lampz.MassAssign(7)= Light7
  Lampz.Callback(7) = "DisableLighting p9, 1.1,"
  Lampz.MassAssign(9)= Light9
  Lampz.MassAssign(9)= Light9h
  Lampz.Callback(9) = "DisableLighting p9, 75,"
  Lampz.MassAssign(10)= Light10
  Lampz.MassAssign(10)= Light10h
  Lampz.Callback(10) = "DisableLighting p10, 75,"
  Lampz.MassAssign(11)= Light11
  Lampz.MassAssign(11)= Light11h
  Lampz.Callback(11) = "DisableLighting p11, 75,"
  Lampz.MassAssign(12)= Light12
  Lampz.MassAssign(12)= Light12h
  Lampz.Callback(12) = "DisableLighting p12, 75,"
  Lampz.MassAssign(13)= Light13
  Lampz.MassAssign(13)= Light13h
  Lampz.Callback(13) = "DisableLighting p13, 75,"
  Lampz.MassAssign(14)= Light14
  Lampz.MassAssign(14)= Light14h
  Lampz.Callback(14) = "DisableLighting p14, 75,"
  Lampz.MassAssign(15)= Light15
  Lampz.MassAssign(15)= Light15h
  Lampz.Callback(15) = "DisableLighting p15, 75,"
  Lampz.MassAssign(16)= Light16
  Lampz.MassAssign(16)= Light16h
  Lampz.Callback(16) = "DisableLighting p16, 75,"
  Lampz.MassAssign(17)= Light17
  Lampz.MassAssign(17)= Light17h
  Lampz.Callback(17) = "DisableLighting p17, 75,"
  Lampz.MassAssign(18)= Light18
  Lampz.MassAssign(18)= Light18h
  Lampz.Callback(18) = "DisableLighting p18, 75,"
  Lampz.MassAssign(19)= Light19
  Lampz.MassAssign(19)= Light19h
  Lampz.Callback(19) = "DisableLighting p19, 150,"
  Lampz.MassAssign(20)= Light20
  Lampz.MassAssign(20)= Light20h
  Lampz.Callback(20) = "DisableLighting p20, 150,"
  Lampz.MassAssign(21)= Light21
  Lampz.MassAssign(21)= Light21h
  Lampz.Callback(21) = "DisableLighting p21, 150,"
  Lampz.MassAssign(22)= Light22
  Lampz.MassAssign(22)= Light22h
  Lampz.Callback(22) = "DisableLighting p22, 150,"
  Lampz.MassAssign(23)= Light23
  Lampz.MassAssign(23)= Light23h
  Lampz.Callback(23) = "DisableLighting p23, 150,"
  Lampz.MassAssign(24)= Light24
  Lampz.MassAssign(24)= Light24h
  Lampz.Callback(24) = "DisableLighting p24, 150,"
  Lampz.MassAssign(25)= Light25
  Lampz.MassAssign(25)= Light25h
  Lampz.Callback(25) = "DisableLighting p25, 150,"
  Lampz.MassAssign(26)= Light26
  Lampz.MassAssign(26)= Light26h
  Lampz.Callback(26) = "DisableLighting p26, 150,"
  Lampz.MassAssign(27)= Light27
  Lampz.MassAssign(27)= Light27h
  Lampz.Callback(27) = "DisableLighting p27, 150,"
  Lampz.MassAssign(28)= Light28
  Lampz.MassAssign(28)= Light28h
  Lampz.Callback(28) = "DisableLighting p28, 150,"
  Lampz.MassAssign(29)= Light29
  Lampz.MassAssign(29)= Light29h
  Lampz.Callback(29) = "DisableLighting p29, 300,"
  Lampz.MassAssign(30)= Light30
  Lampz.MassAssign(30)= Light30h
  Lampz.Callback(30) = "DisableLighting p30, 300,"
  Lampz.MassAssign(31)= Light31
  Lampz.MassAssign(31)= Light31h
  Lampz.Callback(31) = "DisableLighting p31, 300,"
  Lampz.MassAssign(32)= Light32
  Lampz.MassAssign(32)= Light32h
  Lampz.Callback(32) = "DisableLighting p32, 300,"
  Lampz.MassAssign(33)= Light33
  Lampz.MassAssign(33)= Light33h
  Lampz.Callback(33) = "DisableLighting p33, 300,"
  Lampz.MassAssign(34)= Light34
  Lampz.MassAssign(34)= Light34h
  Lampz.Callback(34) = "DisableLighting p34, 50,"
  Lampz.MassAssign(35)= Light35
  Lampz.MassAssign(35)= Light35h
  Lampz.Callback(35) = "DisableLighting p35, 50,"
  Lampz.MassAssign(36)= Light36
  Lampz.MassAssign(36)= Light36h
  Lampz.Callback(36) = "DisableLighting p36, 50,"
  Lampz.MassAssign(37)= Light37
  Lampz.MassAssign(37)= Light37h
  Lampz.Callback(37) = "DisableLighting p37, 50,"
  Lampz.MassAssign(38)= Light38
  Lampz.MassAssign(38)= Light38h
  Lampz.Callback(38) = "DisableLighting p38, 50,"
  Lampz.MassAssign(39)= Light39
  Lampz.MassAssign(39)= Light39h
  Lampz.MassAssign(39)= Light39f
  Lampz.Callback(39) = "DisableLighting p39, 75,"
  Lampz.MassAssign(40)= Light40
  Lampz.MassAssign(40)= Light40h
  Lampz.MassAssign(40)= Light40f
  Lampz.Callback(40) = "DisableLighting p40, 75,"
  Lampz.MassAssign(41)= Light41
  Lampz.MassAssign(41)= Light41h
  Lampz.MassAssign(41)= Light41f
  Lampz.Callback(41) = "DisableLighting p41, 75,"
  Lampz.MassAssign(42)= Light42
  Lampz.MassAssign(42)= Light42h
  Lampz.MassAssign(42)= Light42f
  Lampz.Callback(42) = "DisableLighting p42, 75,"
  Lampz.MassAssign(43)= Light43
  Lampz.MassAssign(43)= Light43h
  Lampz.Callback(43) = "DisableLighting p43, 75,"
  Lampz.MassAssign(44)= Light44
  Lampz.MassAssign(44)= Light44h
  Lampz.Callback(44) = "DisableLighting p44, 75,"
  Lampz.MassAssign(45)= Light45
  Lampz.MassAssign(45)= Light45h
  Lampz.Callback(45) = "DisableLighting p45, 75,"
  Lampz.MassAssign(46)= Light46
  Lampz.MassAssign(46)= Light46h
  Lampz.Callback(46) = "DisableLighting p46, 75,"
  Lampz.MassAssign(47)= Light47
  Lampz.MassAssign(47)= Light47h
  Lampz.Callback(47) = "DisableLighting p47, 75,"
  Lampz.MassAssign(48)= Light48
  Lampz.MassAssign(48)= Light48h
  Lampz.Callback(48) = "DisableLighting p48, 75,"
  Lampz.MassAssign(49)= Light49
  Lampz.MassAssign(49)= Light49h
  Lampz.Callback(49) = "DisableLighting p49, 75,"
  Lampz.MassAssign(50)= Light50
  Lampz.MassAssign(50)= Light50h
  Lampz.Callback(50) = "DisableLighting p50, 75,"
  Lampz.MassAssign(51)= Light51
  Lampz.MassAssign(51)= Light51h
  Lampz.Callback(51) = "DisableLighting p51, 75,"
  Lampz.MassAssign(52)= Light52
  Lampz.MassAssign(52)= Light52h
  Lampz.Callback(52) = "DisableLighting p52, 150,"
  Lampz.MassAssign(53)= Light53
  Lampz.MassAssign(53)= Light53h
  Lampz.Callback(53) = "DisableLighting p53, 150,"
  Lampz.MassAssign(54)= Light54
  Lampz.MassAssign(54)= Light54h
  Lampz.Callback(54) = "DisableLighting p54, 150,"
  Lampz.MassAssign(55)= Light55
  Lampz.MassAssign(55)= Light55h
  Lampz.MassAssign(55)= Light55f
  Lampz.Callback(55) = "DisableLighting p55, 75,"
  Lampz.MassAssign(56)= Light56
  Lampz.MassAssign(56)= Light56h
  Lampz.MassAssign(56)= Light56f
  Lampz.Callback(56) = "DisableLighting p56, 75,"
  Lampz.MassAssign(57)= Light57
  Lampz.MassAssign(57)= Light57h
  Lampz.Callback(57) = "DisableLighting p57, 75,"
  Lampz.MassAssign(58)= Light58
  Lampz.MassAssign(58)= Light58h
  Lampz.Callback(58) = "DisableLighting p58, 75,"
  Lampz.MassAssign(59)= Light59
  Lampz.MassAssign(59)= Light59h
  Lampz.Callback(59) = "DisableLighting p59, 75,"
  Lampz.MassAssign(60)= Light60
  Lampz.MassAssign(60)= Light60h
  Lampz.Callback(60) = "DisableLighting p60, 150,"
  Lampz.MassAssign(61)= Light61
  Lampz.MassAssign(61)= Light61h
  Lampz.Callback(61) = "DisableLighting p61, 150,"
  Lampz.MassAssign(62)= Light62
  Lampz.MassAssign(62)= Light62h
  Lampz.Callback(62) = "DisableLighting p62, 50,"

  'Large center PF flasher
  Lampz.MassAssign(100)= flash1
  Lampz.MassAssign(100)= flash2
  Lampz.MassAssign(100)= flash3
  Lampz.MassAssign(100)= flash4
  Lampz.MassAssign(100)= flashh
  Lampz.Callback(100) = "DisableLighting flasher_on, 50,"

  'GI
  Lampz.obj(0) = ColtoArray(GI)
  Lampz.Callback(0) = "GIUpdates"
  Lampz.state(0) = 1

  'Turn off all lamps on startup
  Lampz.Init  'This just turns state of any lamps to 1

  'Immediate update to turn on GI, turn off lamps
  Lampz.Update

End Sub


'====================
'Class jungle nf
'====================

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?

Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks

Class LampFader
  Public FadeSpeedDown(140), FadeSpeedUp(140)
  Private Lock(140), Loaded(140), OnOff(140)
  Public UseFunction
  Private cFilter
  Public UseCallback(140), cCallback(140)
  Public Lvl(140), Obj(140)
  Private Mult(140)
  Public FrameTime
  Private InitFrame
  Public Name

  Sub Class_Initialize
    InitFrame = 0
    dim x : for x = 0 to uBound(OnOff)  'Set up fade speeds
      FadeSpeedDown(x) = 1/100  'fade speed down
      FadeSpeedUp(x) = 1/80   'Fade speed up
      UseFunction = False
      lvl(x) = 0
      OnOff(x) = False
      Lock(x) = True : Loaded(x) = False
      Mult(x) = 1
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   ''debug.print Lampz.Locked(100) 'debug
  Public Property Get state(idx) : state = OnOff(idx) : end Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function
  'Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property
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
    'msgbox "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    'if idx = 132 then msgbox out 'debug
    ExecuteGlobal Out

  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    if Input <> OnOff(idx) then  'discard redundant updates
      OnOff(idx) = input
      Lock(idx) = False
      Loaded(idx) = False
    End If
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
  end Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub  'Solenoid Handler

  Public Sub TurnOnStates 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        'debugstr = debugstr & "array found at " & idx & "..."
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
    ''debug.print debugstr
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
  Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

  Public Sub Update1   'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
  End Sub

  Public Sub Update2   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
    Update
  End Sub

  Public Sub Update 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)*Mult(x)) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = Lvl(x)*Mult(x) : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(Lvl(x)*Mult(x))
          Else
            obj(x).Intensityscale = Lvl(x)
          End If
        end if
        if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,Lvl(x)*mult(x) 'Proc
        If Lock(x) Then
          if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True 'finished fading
        end if
      end if
    Next
  End Sub
End Class


'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function


'Helper functions
Sub Proc(string, Callback)  'proc using a string and one argument
  'On Error Resume Next
  dim p : Set P = GetRef(String)
  P Callback
  If err.number = 13 then  msgbox "Proc error! No such procedure: " & vbnewline & string
  if err.number = 424 then msgbox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the end of a 1 dimensional array
  if IsArray(aInput) then 'Input is an array...
    dim tmp : tmp = aArray
    If not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else            'Append existing array with aInput array
      Redim Preserve tmp(uBound(aArray) + uBound(aInput)+1) 'If existing array, increase bounds by uBound of incoming array
      dim x : for x = 0 to uBound(aInput)
        if isObject(aInput(x)) then
          Set tmp(x+uBound(aArray)+1 ) = aInput(x)
        Else
          tmp(x+uBound(aArray)+1 ) = aInput(x)
        End If
      Next
      AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      Redim Preserve aArray(uBound(aArray)+1) 'If array, increase bounds by 1
      if isObject(aInput) then
        Set aArray(uBound(aArray)) = aInput
      Else
        aArray(uBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function


Function ColtoArray(aDict)  'converts a collection to an indexed array. Indexes will come out random probably.
  redim a(999)
  dim count : count = 0
  dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
  redim preserve a(count-1) : ColtoArray = a
End Function

Sub SetLamp(aNr, aOn)
  Lampz.state(aNr) = abs(aOn)
End Sub


'GI related subs

dim giprevalvl
giprevalvl = 0

sub OnPrimsVisible(aValue)
  dim kk
  If aValue then
    For each kk in ON_Prims:kk.visible = 1:next
  Else
    For each kk in ON_Prims:kk.visible = 0:next
  end If
end Sub

sub OffPrimsVisible(aValue)
  dim kk
  If aValue then
    For each kk in OFF_Prims:kk.visible = 1:next
  Else
    For each kk in OFF_Prims:kk.visible = 0:next
  end If
end Sub


'GI callback

Sub GIUpdates(ByVal aLvl) 'argument is unused
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  'debug.print "aLvl=" & aLvl & " giprevalvl=" & giprevalvl

  if aLvl = 0 then                    'GI OFF, let's hide ON prims
    'debug.print "aLvl = 0. OnPrimsVisible False"
    OnPrimsVisible False
    If giprevalvl = 1 Then OffPrimsVisible true
  Elseif aLvl = 1 then                  'GI ON, let's hide OFF prims
    'debug.print "aLvl = 1. OffPrimsVisible False"
    OffPrimsVisible False
    If giprevalvl = 0 Then OnPrimsVisible True
  Else
    if giprevalvl = 0 Then                'GI has just changed from OFF to fading, let's show ON
      'debug.print "giprevalvl = 0. OnPrimsVisible True"
      OnPrimsVisible True
    elseif giprevalvl = 1 Then              'GI has just changed from ON to fading, let's show OFF
      'debug.print "giprevalvl = 1. OffPrimsVisible true"
      OffPrimsVisible true
    Else
      'no change
    end if
  end if

  UpdateMaterial "mesh_on", 0.25,1,0.75,0.04705882,0,1, (aLvl)^1,   RGB(200,200,200),0,0,False,True,0,0,0,0
  UpdateMaterial "Playfield", 0.50,1,0.75,0.04705882,0,1, (aLvl)^1,   RGB(200,200,200),0,0,False,True,0,0,0,0
  'UpdateMaterial "mesh_off", 0.25,1,0.75,0.04705882,0,1, (1-aLvl)^1, RGB(200,200,200),0,0,False,True,0,0,0,0

  leftbat1.blenddisablelighting = aLvl*0.7+0.3
  leftbat.blenddisablelighting = aLvl*0.7+0.3
  rightbat.blenddisablelighting = aLvl*0.7+0.3
  render5off.blenddisablelighting = aLvl*0.5+0.5

  giprevalvl = alvl
End Sub


'******************************************************
'****  END LAMPZ
'******************************************************




'********************************************
' Hybrid code for VR, Cab, and Desktop
'********************************************

Sub SetupVRRoom
  Dim VRThings

  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  for each VRThings in DTRails:VRThings.visible = 0:Next

  if VR_Room = 0 and cab_mode = 0 Then
    for each VRThings in DTRails:VRThings.visible = 1:Next
  Elseif VR_Room = 0 and cab_mode = 1 Then

  Else
    for each VRThings in VRStuff:VRThings.visible = 1:Next
    for each VRThings in VRClock:VRThings.visible = WallClock:Next
  'Custom Walls, Floor, and Roof
    if CustomWalls = 1 Then
      VR_Wall_Left.image = "VR_Wall_Left"
      VR_Wall_Right.image = "VR_Wall_Right"
      VR_Floor.image = "VR_Floor"
      VR_Roof.image = "VR_Roof"
    Else
      VR_Wall_Left.image = "wallpaper left"
      VR_Wall_Right.image = "wallpaper_right"
      VR_Floor.image = "FloorCompleteMap2"
      VR_Roof.image = "light gray"
    end if

    If topper = 1 Then
      Primary_topper.visible = 1
    Else
      Primary_topper.visible = 0
    End If

    If poster = 1 Then
      VRposter.visible = 1
    Else
      VRposter.visible = 0
    End If

    If poster2 = 1 Then
      VRposter2.visible = 1
    Else
      VRposter2.visible = 0
    End If

    If UseFlexDMD Then
      Primary_Grill.Image = "speaker UV Map FlexDMD"
      If FlexDMDStyle = 0 Then
        For each VRThings in VRDisplayOff : VRThings.visible = False : Next
        DMD.visible = true
        DMD2.visible = false
        Backglass.ImageA = "Defender Illuminated FlexDMD"
      ElseIf FlexDMDStyle = 1 Then
        For each VRThings in VRDisplayOff : VRThings.visible = True : Next
        DMD.visible = true
        DMD2.visible = false
        Backglass.ImageA = "Defender Illuminated"
      ElseIf FlexDMDStyle = 2 Then
        For each VRThings in VRDisplayOff : VRThings.visible = False : Next
        DMD.visible = false
        DMD2.visible = true
        Backglass.ImageA = "Defender Illuminated FlexDMD"
        Primary_Grill.Size_Z = 225
        Primary_Grill.x = 14
        Primary_Grill.y = -97
        Primary_Grill.z = -417
      End If
    End if
  End If
End Sub

'********************************************
' VR Clock
'********************************************

Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())
End Sub



' ***************************************************************************
'                                  LAMP CALLBACK
' ****************************************************************************

if VR_Room = 1 AND (UseFlexDMD = False OR FlexDMDStyle = 1) Then
'if VR_Room = 1 AND UseFlexDMD = False Then
  Set LampCallback = GetRef("UpdateMultipleLamps")
End If

Sub UpdateMultipleLamps()
  If Controller.Lamp(1) = 0 Then: l1.visible  =0: else: l1.visible  =1 'Shoot Again
  If Controller.Lamp(2) = 0 Then: l2.visible  =0: else: l2.visible  =1 'Ball In Play
  If Controller.Lamp(3) = 0 Then: l3.visible  =0: else: l3.visible  =1 'Tilt
  If Controller.Lamp(4) = 0 Then: l4.visible  =0: else: l4.visible  =1 'Game Over
  If Controller.Lamp(5) = 0 Then: l5.visible  =0: else: l5.visible  =1 'Match
  If Controller.Lamp(6) = 0 Then: l6.visible  =0: else: l6.visible  =1 'High Score
  If Controller.Lamp(6) = 0 Then: l6a.visible =0: else: l6a.visible =1 'High Score
  If Controller.Lamp(8) = 0 Then: l8.visible  =0: else: l8.visible  =1 'Wave In Play
End Sub


if VR_Room = 0 and cab_mode = 0 AND (UseFlexDMD = False OR FlexDMDStyle = 1) Then
  Set LampCallback = GetRef("UpdateDTLamps")
End If


Sub UpdateDTLamps()
  If Controller.Lamp(1) = 0 Then: ShootAgainReel.setValue(0):   Else: ShootAgainReel.setValue(1) 'Shoot Again
  If Controller.Lamp(2) = 0 Then: BIPReel.setValue(0):      Else: BIPReel.setValue(1) 'Ball in Play
  If Controller.Lamp(3) = 0 Then: TiltReel.setValue(0):     Else: TiltReel.setValue(1) 'Tilt
  If Controller.Lamp(4) = 0 Then: GameOverReel.setValue(0):   Else: GameOverReel.setValue(1) 'Game Over
  If Controller.Lamp(5) = 0 Then: MatchReel.setValue(0):      Else: MatchReel.setValue(1) 'Match
  If Controller.Lamp(6) = 0 Then: HighScoreReel.setValue(0):    Else: HighScoreReel.setValue(1) 'High Score
  If Controller.Lamp(8) = 0 Then: WaveInPlayReel.setValue(0):   Else: WaveInPlayReel.setValue(1) 'Wave In Play
End Sub



' *********************************************************************
' VR Alphanumeric Display
' *********************************************************************


dim DisplayColor
DisplayColor =  RGB(255,40,1)


'Sub VRDisplaytimer  'Doing this in DisplayTimer sub now
' Dim ii, jj, obj, b, x
' Dim ChgLED,num, chg, stat
' ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
'   If Not IsEmpty(ChgLED) Then
'     For ii=0 To UBound(chgLED)
'       num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
'       For Each obj In VRDigits(num)
'         If chg And 1 Then FadeDisplay obj, stat And 1
'         chg=chg\2 : stat=stat\2
'       Next
'     Next
'   End If
'End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 800
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 250
  End If
End Sub

Dim VRDigits(32)

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

VRDigits(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218)
VRDigits(15)=Array(D219,D220,D221,D222,D223,D224,D225,D226)
VRDigits(16)=Array(D227,D228,D229,D230,D231,D232,D233,D234)
VRDigits(17)=Array(D235,D236,D237,D238,D239,D240,D241,D242)
VRDigits(18)=Array(D243,D244,D245,D246,D247,D248,D249,D250)
VRDigits(19)=Array(D251,D252,D253,D254,D255,D256,D257,D258)
VRDigits(20)=Array(D259,D260,D261,D262,D263,D264,D265,D266)

VRDigits(21)=Array(D267,D268,D269,D270,D271,D272,D273,D274)
VRDigits(22)=Array(D275,D276,D277,D278,D279,D280,D281,D282)
VRDigits(23)=Array(D283,D284,D285,D286,D287,D288,D289,D290)
VRDigits(24)=Array(D291,D292,D293,D294,D295,D296,D297,D298)
VRDigits(25)=Array(D299,D300,D301,D302,D303,D304,D305,D306)
VRDigits(26)=Array(D307,D308,D309,D310,D311,D312,D313,D314)
VRDigits(27)=Array(D315,D316,D317,D318,D319,D320,D321,D322)

VRDigits(28)=Array(D323,D324,D325,D326,D327,D328,D329,D330)
VRDigits(29)=Array(D331,D332,D333,D334,D335,D336,D337,D338)
VRDigits(30)=Array(D339,D340,D341,D342,D343,D344,D345,D346)
VRDigits(31)=Array(D347,D348,D349,D350,D351,D352,D353,D354)


Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      If UseFlexDMD = False OR FlexDMDStyle = 1 Then
        For each obj in VRDigits(x)
          obj.height = obj.height + 18
          FadeDisplay obj, 0
        next
      Else
        For each obj in VRDigits(x)
          'hide segment display if using flex
          obj.Visible = False
        next
      End If
    end If
  Next
End Sub

initDigits

'**********************************************
'*******  Set Up Backglass Flashers *******
'**********************************************
' this is for lining up the backglass flashers on top of a backglass image

Sub SetBackglass()
  Dim obj

' For Each obj In BackglassLow
'   obj.x = obj.x
'   obj.height = - obj.y + -20
'   obj.y = -142 'adjusts the distance from the backglass towards the user
' Next


  For Each obj In BackglassMid
    obj.x = obj.x
    obj.height = - obj.y + 1150
    obj.y = 80 'adjusts the distance from the backglass towards the user
  Next


' For Each obj In BackglassHigh
'   obj.x = obj.x
'   obj.height = - obj.y + 40
'   obj.y = -190 'adjusts the distance from the backglass towards the user
' Next

End Sub

'**********************************************




'**********************************************************************************************************
' FlexDMD code scutters, VR DMD Mike_da_Spike
'**********************************************************************************************************
Dim FlexDMD
Dim FlexDMDScene
DIm FlexDMDDict
Dim ExternalEnabled
Dim ActivePlayer, BallNumber, SmartBombUsed, cntTimer

Dim DMDPosMap, DMDPosHumanoids(9,1), DMDPosBaiters(2,1), DMDPosLanders(9,1), DMDPosPods(1,1), DMDPosSwarmers(5,1), DMDPosBombers(3,1)


Sub FlexTimer_Timer()
  DisplayTimer
End Sub

Sub FlexDMD_Init() 'default/startup values

  If UseFlexDMD = False then
    Exit Sub
  End if

  If IsNumeric(FlexDMDStyle) Then
    Select Case FlexDMDStyle
    Case 0, 1, 2
    Case Else
      FlexDMDStyle = 0
    End Select
  Else
    FlexDMDStyle = 0
  End If

  If LCase(cGameName) <> "dfndr_l4" Then
    'nvram code used by flex only tested with dfndr_l4
    MsgBox "Wrong rom " & cGameName & ". Use dfndr_l4 if enabling FlexDMD"
    UseFlexDMD = False
    Exit Sub
  End If

  If VR_Room = 0 AND cab_mode = 0 AND FlexDMDStyle <> 1 Then
    If Table1.BackdropImage_DT = "DT_Defender3" Then
      'change DT background (if user hasn't set to None or a custom one)
      Table1.BackdropImage_DT = "DT_Defender4"
    End If
  End If


  ActivePlayer = Array(0,0,1)
  BallNumber = Array(0,0)
  SmartBombUsed = 0
  cntTimer = 0

  'pause display timer while flex
  FlexTimer.Enabled = False

  On Error Resume Next
  Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
  On Error GoTo 0

  If IsObject(FlexDMD) Then

    FlexDMD.RenderMode = 2
    FlexDMD.Width = 128
    If FlexDMDStyle <> 2 Then
      FlexDMD.Height = 32
    Else
      FlexDMD.Height = 72
    End If
    FlexDMD.Clear = True
    FlexDMD.GameName = cGameName
    FlexDMD.TableFile = Table1.Filename & ".vpx"
    FlexDMD.Run = True

    FlexDMD.LockRenderThread

    Set FlexDMDScene = FlexDMD.NewGroup("Scene")

    Dim vpmColourised, vpm33Percent, vpm66Percent, vpmOffPercent

    With FlexDMDScene

      'backcolour for when smartbombs used
      .AddActor FlexDMD.NewFrame("BackColour")
      .GetFrame("BackColour").Visible = True
      .GetFrame("BackColour").Height = FlexDMD.Height
      .GetFrame("BackColour").Width= 128
      .GetFrame("BackColour").Fill= True
      .GetFrame("BackColour").Thickness= 0
      .GetFrame("BackColour").FillColor = vbBlack

      '.AddActor FlexDMD.NewImage("BackG", "FlexDMD.Resources.dmds.black.png")
      If FlexDMDStyle <> 1 Then 'add score stuff

        '2*7 segment displays for 2 player scores
        dim i,j
        For i = 14 to 20
          'small scores P1 & P2
          .AddActor FlexDMD.NewImage("Seg" & cstr(i), "VPX.DMD_S_Blank")
          .AddActor FlexDMD.NewImage("Seg" & cstr(i + 7), "VPX.DMD_S_Blank")
          Select Case i
          Case 14
            j = 1
          Case 15,16,17
            j = 2
          Case 18,19,20
            j = 3
          End Select
          .GetImage("Seg" & cstr(i)).SetAlignedPosition ((i - 14) * 7) + j,0,0
          .GetImage("Seg" & cstr(i + 7)).SetAlignedPosition ((i - 14) * 7) + 77 + j,0,0
          'large scores
          .AddActor FlexDMD.NewImage("SegL" & cstr(i), "VPX.DMD_L_Blank")
          .AddActor FlexDMD.NewImage("SegL" & cstr(i + 7), "VPX.DMD_L_Blank")
          Select Case i
          Case 14
            j = 4
          Case 15,16,17
            j = 6
          Case 18,19,20
            j = 8
          End Select
          .GetImage("SegL" & cstr(i)).SetAlignedPosition ((i - 14) * 16) + j,9,0
          .GetImage("SegL" & cstr(i + 7)).SetAlignedPosition ((i - 14) * 16) + j,9,0
          .GetImage("SegL" & cstr(i)).Visible = False
          .GetImage("SegL" & cstr(i + 7)).Visible = False
        Next

        '3 smartbombs
        For i = 1 to 3
          .AddActor FlexDMD.NewImage("Bomb" & cstr(i), "VPX.DMD_Image_SmartBomb1")
          .GetImage("Bomb" & cstr(i)).SetAlignedPosition 120,(i + 1) * 5,0
          .GetImage("Bomb" & cstr(i)).Visible = False
        Next

        'fighters (max 5 displayed even though game can be up to 9 ball)
        For i = 1 to 5
          .AddActor FlexDMD.NewImage("Fighter" & cstr(i), "VPX.DMD_Image_Fighter1")
          .GetImage("Fighter" & cstr(i)).SetAlignedPosition (i + (i - 1) * 12) , 27, 0
          .GetImage("Fighter" & cstr(i)).Visible = False
        Next

        'credit and ball/match number
        For i = 28 to 29
          .AddActor FlexDMD.NewImage("Ball" & cstr(i), "VPX.DMD_S_Blank")
          .GetImage("Ball" & cstr(i)).SetAlignedPosition ((i - 28) * 7) + 42,25,0
          .GetImage("Ball" & cstr(i)).Visible = False
        Next
        .AddActor FlexDMD.NewImage("BallMatch", "VPX.DMD_Image_Match")
        .GetImage("BallMatch").SetAlignedPosition 1,25,0
        .GetImage("BallMatch").Visible = False
        For i = 30 to 31
          .AddActor FlexDMD.NewImage("Credit" & cstr(i), "VPX.DMD_S_Blank")
          .GetImage("Credit" & cstr(i)).SetAlignedPosition ((i - 30) * 7) + 115,25,0
          .GetImage("Credit" & cstr(i)).Visible = False
        Next
        .AddActor FlexDMD.NewImage("CreditsWave", "VPX.DMD_Image_Credits")
        .GetImage("CreditsWave").SetAlignedPosition 64,25,0
        '.GetImage("CreditsWave").Visible = False

        'game over, high score, tilt
        .AddActor FlexDMD.NewImage("HighScore", "VPX.DMD_Image_HighScore")
        .GetImage("HighScore").SetAlignedPosition 31,8,0
        .GetImage("HighScore").Visible = False
        .AddActor FlexDMD.NewImage("HighScoreOver", "VPX.DMD_Image_HighScore_Over")
        .GetImage("HighScoreOver").SetAlignedPosition 0,0,0
        .GetImage("HighScoreOver").Visible = False
        .AddActor FlexDMD.NewImage("GameOver", "VPX.DMD_Image_GameOver")
        .GetImage("GameOver").SetAlignedPosition 32,17,0
        .GetImage("GameOver").Visible = False
        .AddActor FlexDMD.NewImage("Tilt", "VPX.DMD_Image_Tilt")
        .GetImage("Tilt").SetAlignedPosition 0,9,0'51,11,0
        .GetImage("Tilt").Visible = False
      End If

      If FlexDMDStyle > 0 Then 'add scanner stuff

        FlexDMDInitialPositions

        DMDPosMap = -64
        .AddActor FlexDMD.NewImage("Map", "VPX.DMD_ScannerMap")
        If FlexDMDStyle = 2 Then
          .GetImage("Map").SetAlignedPosition DMDPosMap,32,0
        Else
          .GetImage("Map").SetAlignedPosition DMDPosMap,-8,0
        End If

        '10 humanoids
        For i = 0 To 9
          .AddActor FlexDMD.NewImage("Humanoid" & i, "VPX.DMD_Humanoid")
          .GetImage("Humanoid" & i).SetAlignedPosition DMDPosHumanoids(i,0),DMDPosHumanoids(i,1),0
        Next

        '3 baiters
        For i = 0 To 2
          .AddActor FlexDMD.NewImage("Baiter" & i, "VPX.DMD_Baiter")
          .GetImage("Baiter" & i).SetAlignedPosition DMDPosBaiters(i,0),DMDPosBaiters(i,1),0
          .GetImage("Baiter" & i).Visible = False
        Next

        '10 landers (also mutants when lamp lit)
        For i = 0 To 9
          .AddActor FlexDMD.NewImage("Lander" & i, "VPX.DMD_Lander")
          .GetImage("Lander" & i).SetAlignedPosition DMDPosLanders(i,0),DMDPosLanders(i,1),0
          .GetImage("Lander" & i).Visible = False
        Next

        '2 pods
        For i = 0 To 1
          .AddActor FlexDMD.NewImage("Pod" & i, "VPX.DMD_Pod1")
          .GetImage("Pod" & i).SetAlignedPosition DMDPosPods(i,0),DMDPosPods(i,1),0
          .GetImage("Pod" & i).Visible = False
        Next

        '6 swarmers
        For i = 0 To 5
          .AddActor FlexDMD.NewImage("Swarmer" & i, "VPX.DMD_Swarmer")
          .GetImage("Swarmer" & i).SetAlignedPosition DMDPosSwarmers(i,0),DMDPosSwarmers(i,1),0
          .GetImage("Swarmer" & i).Visible = False
        Next

        '4 bombers
        For i = 0 To 3
          .AddActor FlexDMD.NewImage("Bomber" & i, "VPX.DMD_Bomber")
          .GetImage("Bomber" & i).SetAlignedPosition DMDPosBombers(i,0),DMDPosBombers(i,1),0
          .GetImage("Bomber" & i).Visible = False
        Next

        'overlay border
        .AddActor FlexDMD.NewImage("Overlay", "VPX.DMD_ScannerOverlay")
        If FlexDMDStyle = 2 Then

          .GetImage("Overlay").SetAlignedPosition 0,32,0
        Else
          '.AddActor FlexDMD.NewFrame("Overlay")
          '.GetFrame("Overlay").Visible = True
          '.GetFrame("Overlay").Height = FlexDMD.Height
          '.GetFrame("Overlay").Width= 128
          '.GetFrame("Overlay").Fill= False
          '.GetFrame("Overlay").Thickness= 1
          '.GetFrame("Overlay").BorderColor = RGB(192,0,40)
          .GetImage("Overlay").SetAlignedPosition 0,0,0
          .GetImage("Overlay").SetBounds 0,0,128,32

        End If

      End If

      .AddActor FlexDMD.NewImage("Defender", "VPX.DMD_Image_Defender")


    End With

    FlexDMD.Stage.AddActor FlexDMDScene
    FlexDMD.UnlockRenderThread

    If FlexDMDStyle <> 1 Then FlexDictionary_Init

    FlexTimer.Enabled = True
    FlexTimer.Interval = 40 'LampTimer.Interval

  Else

    UseFlexDMD = False

  End If

End Sub

Sub FlexDictionary_Init

  'add conversion of segment charcters codes to lookup table

  Set FlexDMDDict = CreateObject("Scripting.Dictionary")

  FlexDMDDict.Add 0, "VPX.DMD_S_Blank"
  FlexDMDDict.Add 63, "VPX.DMD_S_0"
  FlexDMDDict.Add 6, "VPX.DMD_S_1"
  FlexDMDDict.Add 91, "VPX.DMD_S_2"
  FlexDMDDict.Add 79, "VPX.DMD_S_3"
  FlexDMDDict.Add 102, "VPX.DMD_S_4"
  FlexDMDDict.Add 109, "VPX.DMD_S_5"
  FlexDMDDict.Add 125, "VPX.DMD_S_6"
  FlexDMDDict.Add 7, "VPX.DMD_S_7"
  FlexDMDDict.Add 127, "VPX.DMD_S_8"
  FlexDMDDict.Add 111, "VPX.DMD_S_9"

  FlexDMDDict.Add 191, "VPX.DMD_S_0d"
  FlexDMDDict.Add 134, "VPX.DMD_S_1d"
  FlexDMDDict.Add 219, "VPX.DMD_S_2d"
  FlexDMDDict.Add 207, "VPX.DMD_S_3d"
  FlexDMDDict.Add 230, "VPX.DMD_S_4d"
  FlexDMDDict.Add 237, "VPX.DMD_S_5d"
  FlexDMDDict.Add 253, "VPX.DMD_S_6d"
  FlexDMDDict.Add 135, "VPX.DMD_S_7d"
  FlexDMDDict.Add 255, "VPX.DMD_S_8d"
  FlexDMDDict.Add 239, "VPX.DMD_S_9d"

  '+1000 values for large score font
  FlexDMDDict.Add 1000, "VPX.DMD_L_Blank"
  FlexDMDDict.Add 1063, "VPX.DMD_L_0"
  FlexDMDDict.Add 1006, "VPX.DMD_L_1"
  FlexDMDDict.Add 1091, "VPX.DMD_L_2"
  FlexDMDDict.Add 1079, "VPX.DMD_L_3"
  FlexDMDDict.Add 1102, "VPX.DMD_L_4"
  FlexDMDDict.Add 1109, "VPX.DMD_L_5"
  FlexDMDDict.Add 1125, "VPX.DMD_L_6"
  FlexDMDDict.Add 1007, "VPX.DMD_L_7"
  FlexDMDDict.Add 1127, "VPX.DMD_L_8"
  FlexDMDDict.Add 1111, "VPX.DMD_L_9"

  FlexDMDDict.Add 1191, "VPX.DMD_L_0d"
  FlexDMDDict.Add 1134, "VPX.DMD_L_1d"
  FlexDMDDict.Add 1219, "VPX.DMD_L_2d"
  FlexDMDDict.Add 1207, "VPX.DMD_L_3d"
  FlexDMDDict.Add 1230, "VPX.DMD_L_4d"
  FlexDMDDict.Add 1237, "VPX.DMD_L_5d"
  FlexDMDDict.Add 1253, "VPX.DMD_L_6d"
  FlexDMDDict.Add 1135, "VPX.DMD_L_7d"
  FlexDMDDict.Add 1255, "VPX.DMD_L_8d"
  FlexDMDDict.Add 1239, "VPX.DMD_L_9d"

  'as values for tracking ball number, +2000
  FlexDMDDict.Add 2063, "0"
  FlexDMDDict.Add 2006, "1"
  FlexDMDDict.Add 2091, "2"
  FlexDMDDict.Add 2079, "3"
  FlexDMDDict.Add 2102, "4"
  FlexDMDDict.Add 2109, "5"
  FlexDMDDict.Add 2125, "6"
  FlexDMDDict.Add 2007, "7"
  FlexDMDDict.Add 2127, "8"
  FlexDMDDict.Add 2111, "9"

End Sub

Sub UpdateFlexChar(id, value)

  If FlexDMDScene.GetImage("Defender").Visible = True Then
    If id = 27 Then FlexDMDScene.GetImage("Defender").Visible = False
  End If

  If id < 28 Then
    If id > 13 Then
      If FlexDMDDict.Exists (value) then
        FlexDMDScene.GetImage("Seg" & id).Bitmap = FlexDMD.NewImage("", FlexDMDDict.Item (value)).Bitmap
        FlexDMDScene.GetImage("SegL" & id).Bitmap = FlexDMD.NewImage("", FlexDMDDict.Item (value + 1000)).Bitmap

        If id = 27 Or id = 20 Then
          'active player score traking, looking for last digit flashing '0' ' ' '0'
          If id = 27 Then
            Select Case value
            Case 0, 63
              If ActivePlayer(1) < 126 Then ActivePlayer(1) = ActivePlayer(1) + value   'if total = 126 then it's the active player score
              ActivePlayer(0) = 0
            Case 0
            Case Else
              ActivePlayer(1) = 0
            End Select
          End If
          If id = 20 Then
            Select Case value
            Case 0, 63
              If ActivePlayer(0) < 126 Then ActivePlayer(0) = ActivePlayer(0) + value
              ActivePlayer(1) = 0
            Case 0
            Case Else
              ActivePlayer(0) = 0
            End Select
          End If
          'we now know active player
          If ActivePlayer(1) = 126 Then ActivePlayer(2) = 2
          If ActivePlayer(0) = 126 Then ActivePlayer(2) = 1

        End If
      End If
    End If

  ElseIf id < 30 Then
    If FlexDMDDict.Exists (value) Then
      FlexDMDScene.GetImage("Ball" & id).Bitmap = FlexDMD.NewImage("", FlexDMDDict.Item (value)).Bitmap
      BallNumber(id - 28) = FlexDMDDict.Item (value + 2000)
    End If
  ElseIf id < 32 Then
    If FlexDMDDict.Exists (value) Then
      FlexDMDScene.GetImage("Credit" & id).Bitmap = FlexDMD.NewImage("", FlexDMDDict.Item (value)).Bitmap
    End If
  End If

End Sub


Sub UpdateFlexScanner (InGame, TimerCount)
Dim i,x,rx,ry
Dim minX, maxX, minY, maxY

  'alien min max positions
  If FlexDMDStyle = 2 Then
    'fulldmd positions
    minX = 1
    maxX = 126
    minY = 33
    maxY = 58
  Else
    'real/slim positions
    minX = 1
    maxX = 126
    minY = 1
    maxY = 18
  End If

  With FlexDMDScene

    If InGame = 1 Then

      If TimerCount Mod 2 = 1 Then

        'direction / movement increment
        x = 1
        If Lampz.State(61) = 0 Then x = -x      'reverse lamp

        'scanner map
        DMDPosMap = DMDPosMap - x
        If DMDPosMap < -128 Then DMDPosMap = DMDPosMap + 128
        If DMDPosMap > 0 Then DMDPosMap = DMDPosMap - 128
        If FlexDMDStyle = 2 Then
          .GetImage("Map").SetAlignedPosition DMDPosMap,32,0
        Else
          .GetImage("Map").SetAlignedPosition DMDPosMap,-8,0
        End If

        'humanoids
        For i = 0 To 9
          DMDPosHumanoids(i,0) = DMDPosHumanoids(i,0) - x
          If DMDPosHumanoids(i,0) < 0 Then DMDPosHumanoids(i,0) = DMDPosHumanoids(i,0) + 128
          If DMDPosHumanoids(i,0) > 128 Then DMDPosHumanoids(i,0) = DMDPosHumanoids(i,0) - 128
          .GetImage("Humanoid" & i).SetAlignedPosition DMDPosHumanoids(i,0),DMDPosHumanoids(i,1),0
          .GetImage("Humanoid" & i).Visible = Lampz.State(i + 9)
        Next

        'baiters
        For i = 0 To 2
          If Not CBool(controller.Switch(i + 33)) = True Then

            rx = int(rnd * 6) - 3 - x   'random number -3 to 3, take into account x scroll
            ry = int(rnd * 5) - 1     'random number -2 to 2

            If minX < DMDPosBaiters(i,0) + rx And maxX > DMDPosBaiters(i,0) + rx Then
              DMDPosBaiters(i,0) = DMDPosBaiters(i,0) + rx
            ElseIf minX > DMDPosBaiters(i,0) + rx Then
              DMDPosBaiters(i,0) = DMDPosBaiters(i,0) + rx + 126
            ElseIf maxX < DMDPosBaiters(i,0) + rx Then
              DMDPosBaiters(i,0) = DMDPosBaiters(i,0) + rx - 126
            End If

            If minY < DMDPosBaiters(i,1) + ry And maxY > DMDPosBaiters(i,1) + ry Then
              DMDPosBaiters(i,1) = DMDPosBaiters(i,1) + ry
            Else
              DMDPosBaiters(i,1) = DMDPosBaiters(i,1) - ry
            End If
            .GetImage("Baiter" & i).SetAlignedPosition DMDPosBaiters(i,0),DMDPosBaiters(i,1),0
            .GetImage("Baiter" & i).Visible = True
          Else
            .GetImage("Baiter" & i).Visible = False
          End If

        Next

        'landers & mutants
        For i = 0 To 9
          rx = int(rnd * 3) - 1 - x   'random number -1 to 1, take into account x scroll
          ry = int(rnd * 3) - 1     'random number -1 to 1

          If minX < DMDPosLanders(i,0) + rx And maxX > DMDPosLanders(i,0) + rx Then
            DMDPosLanders(i,0) = DMDPosLanders(i,0) + rx
          ElseIf minX > DMDPosLanders(i,0) + rx Then
            DMDPosLanders(i,0) = DMDPosLanders(i,0) + rx + 126
          ElseIf maxX < DMDPosLanders(i,0) + rx Then
            DMDPosLanders(i,0) = DMDPosLanders(i,0) + rx - 126
          End If

          If minY < DMDPosLanders(i,1) + ry And maxY > DMDPosLanders(i,1) + ry Then
            DMDPosLanders(i,1) = DMDPosLanders(i,1) + ry
          Else
            DMDPosLanders(i,1) = DMDPosLanders(i,1) - ry
          End If
          .GetImage("Lander" & i).SetAlignedPosition DMDPosLanders(i,0),DMDPosLanders(i,1),0

          If Lampz.State(i + 19) = 0 Then
            If i < 5 Then
              .GetImage("Lander" & i).Visible = Not CBool(controller.Switch(i + 13))
            Else
              .GetImage("Lander" & i).Visible = Not CBool(controller.Switch(i + 23 - 5))
            End if
            .GetImage("Lander" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Lander").Bitmap
          Else
            Select Case TimerCount
            Case 1
              .GetImage("Lander" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Mutant1").Bitmap
            Case 3
              .GetImage("Lander" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Mutant2").Bitmap
            Case 7
              .GetImage("Lander" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Mutant3").Bitmap
            Case 9
              .GetImage("Lander" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Mutant4").Bitmap
            Case 11
              .GetImage("Lander" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Mutant5").Bitmap
            Case 15
              .GetImage("Lander" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Mutant6").Bitmap
            End Select
            .GetImage("Lander" & i).Visible = True
          End If
        Next

        'pods
        For i = 0 To 1
          If Not CBool(controller.Switch(i + 39)) = True Then
            rx = int(rnd * 3) - 1 - x   'random number -1 to 1, take into account x scroll
            ry = int(rnd * 3) - 1     'random number -1 to 1

            If minX < DMDPosPods(i,0) + rx And maxX > DMDPosPods(i,0) + rx Then
              DMDPosPods(i,0) = DMDPosPods(i,0) + rx
            ElseIf minX > DMDPosPods(i,0) + rx Then
              DMDPosPods(i,0) = DMDPosPods(i,0) + rx + 126
            ElseIf maxX < DMDPosPods(i,0) + rx Then
              DMDPosPods(i,0) = DMDPosPods(i,0) + rx - 126
            End If

            If minY < DMDPosPods(i,1) + ry And maxY > DMDPosPods(i,1) + ry Then
              DMDPosPods(i,1) = DMDPosPods(i,1) + ry
            Else
              DMDPosPods(i,1) = DMDPosPods(i,1) - ry
            End If
            .GetImage("Pod" & i).SetAlignedPosition DMDPosPods(i,0),DMDPosPods(i,1),0

            Select Case TimerCount
            Case 1
              .GetImage("Pod" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Pod1").Bitmap
            Case 3
              .GetImage("Pod" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Pod2").Bitmap
            Case 7
              .GetImage("Pod" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Pod3").Bitmap
            Case 9
              .GetImage("Pod" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Pod4").Bitmap
            Case 11
              .GetImage("Pod" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Pod5").Bitmap
            Case 15
              .GetImage("Pod" & i).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Pod6").Bitmap
            End Select
            .GetImage("Pod" & i).Visible = True
          Else
            .GetImage("Pod" & i).Visible = False
          End If

        Next

        'swarmers (count lit swarmer lamps for swarmer count)
        dim cnt
        cnt = 0
        For i = 0 to 5
          If Lampz.State(43 + i) = 1 Then cnt = cnt + 1
          .GetImage("Swarmer" & i).Visible = False
        Next
        If cnt > 0 Then
          For i = 0 to cnt - 1
            rx = int(rnd * 9) - 4 - x   'random number -4 to 4, take into account x scroll
            ry = int(rnd * 5) - 2     'random number -2 to 2

            If minX < DMDPosSwarmers(i,0) + rx And maxX > DMDPosSwarmers(i,0) + rx Then
              DMDPosSwarmers(i,0) = DMDPosSwarmers(i,0) + rx
            ElseIf minX > DMDPosSwarmers(i,0) + rx Then
              DMDPosSwarmers(i,0) = DMDPosSwarmers(i,0) + rx + 126
            ElseIf maxX < DMDPosSwarmers(i,0) + rx Then
              DMDPosSwarmers(i,0) = DMDPosSwarmers(i,0) + rx - 126
            End If
            If minY < DMDPosSwarmers(i,1) + ry And maxY > DMDPosSwarmers(i,1) + ry Then
              DMDPosSwarmers(i,1) = DMDPosSwarmers(i,1) + ry
            Else
              DMDPosSwarmers(i,1) = DMDPosSwarmers(i,1) - ry
            End If
            .GetImage("Swarmer" & i).SetAlignedPosition DMDPosSwarmers(i,0),DMDPosSwarmers(i,1),0
            .GetImage("Swarmer" & i).Visible = True
          Next
        End If

        'bombers (count unlit lanes for bomber Count)
        cnt = 0
        For i = 0 to 3
          If Lampz.State(39 + i) = 0 Then cnt = cnt + 1
          .GetImage("Bomber" & i).Visible = False
        Next
        If cnt > 0 Then
          For i = 0 to cnt - 1
            rx = int(rnd * 3) - 1 - x   'random number -1 to 1, take into account x scroll
            ry = int(rnd * 3) - 1     'random number -1 to 1

            If minX < DMDPosBombers(i,0) + rx And maxX > DMDPosBombers(i,0) + rx Then
              DMDPosBombers(i,0) = DMDPosBombers(i,0) + rx
            ElseIf minX > DMDPosBombers(i,0) + rx Then
              DMDPosBombers(i,0) = DMDPosBombers(i,0) + rx + 126
            ElseIf maxX < DMDPosBombers(i,0) + rx Then
              DMDPosBombers(i,0) = DMDPosBombers(i,0) + rx - 126
            End If
            If minY < DMDPosBombers(i,1) + ry And maxY > DMDPosBombers(i,1) + ry Then
              DMDPosBombers(i,1) = DMDPosBombers(i,1) + ry
            Else
              DMDPosBombers(i,1) = DMDPosBombers(i,1) - ry
            End If
            .GetImage("Bomber" & i).SetAlignedPosition DMDPosBombers(i,0),DMDPosBombers(i,1),0
            .GetImage("Bomber" & i).Visible = True
          Next
        End If

      End If

    Else

    End If

  End With

End Sub

Sub FlexDMDInitialPositions

  '10 humanoids
  If FlexDMDStyle = 2 Then
    'fulldmd positions
    DMDPosHumanoids(0,0) = 5
    DMDPosHumanoids(0,1) = 62
    DMDPosHumanoids(1,0) = 17
    DMDPosHumanoids(1,1) = 63
    DMDPosHumanoids(2,0) = 32
    DMDPosHumanoids(2,1) = 67
    DMDPosHumanoids(3,0) = 45
    DMDPosHumanoids(3,1) = 64
    DMDPosHumanoids(4,0) = 58
    DMDPosHumanoids(4,1) = 67
    DMDPosHumanoids(5,0) = 67
    DMDPosHumanoids(5,1) = 65
    DMDPosHumanoids(6,0) = 84
    DMDPosHumanoids(6,1) = 65
    DMDPosHumanoids(7,0) = 102
    DMDPosHumanoids(7,1) = 66
    DMDPosHumanoids(8,0) = 117
    DMDPosHumanoids(8,1) = 62
    DMDPosHumanoids(9,0) = 123
    DMDPosHumanoids(9,1) = 67
  Else
    'real/slim positions
    DMDPosHumanoids(0,0) = 5
    DMDPosHumanoids(0,1) = 22
    DMDPosHumanoids(1,0) = 17
    DMDPosHumanoids(1,1) = 23
    DMDPosHumanoids(2,0) = 32
    DMDPosHumanoids(2,1) = 27
    DMDPosHumanoids(3,0) = 45
    DMDPosHumanoids(3,1) = 24
    DMDPosHumanoids(4,0) = 58
    DMDPosHumanoids(4,1) = 27
    DMDPosHumanoids(5,0) = 67
    DMDPosHumanoids(5,1) = 25
    DMDPosHumanoids(6,0) = 84
    DMDPosHumanoids(6,1) = 25
    DMDPosHumanoids(7,0) = 102
    DMDPosHumanoids(7,1) = 26
    DMDPosHumanoids(8,0) = 117
    DMDPosHumanoids(8,1) = 22
    DMDPosHumanoids(9,0) = 123
    DMDPosHumanoids(9,1) = 27
  End If

  Dim i, rX, rY
  Dim minX, maxX, minY, maxY

  'alien min max positions
  If FlexDMDStyle = 2 Then
    'fulldmd positions
    minX = 1
    maxX = 126
    minY = 33
    maxY = 58
  Else
    'real/slim positions
    minX = 1
    maxX = 126
    minY = 1
    maxY = 18
  End If

  'get random start positions

  'baiters
  For i = 0 To 2
    rX = Int((maxX - minX + 1) * rnd + minX)
    DMDPosBaiters(i,0) = rX
    rY = Int((maxY - minY + 1) * rnd + minY)
    DMDPosBaiters(i,1) = rY
  Next

  'landers
  For i = 0 To 9
    rX = Int((maxX - minX + 1) * rnd + minX)
    DMDPosLanders(i,0) = rX
    rY = Int((maxY - minY + 1) * rnd + minY)
    DMDPosLanders(i,1) = rY
  Next

  'pods
  For i = 0 To 1
    rX = Int((maxX - minX + 1) * rnd + minX)
    DMDPosPods(i,0) = rX
    rY = Int((maxY - minY + 1) * rnd + minY)
    DMDPosPods(i,1) = rY
  Next

  'swarmers
  For i = 0 To 5
    rX = Int((maxX - minX + 1) * rnd + minX)
    DMDPosSwarmers(i,0) = rX
    rY = Int((maxY - minY + 1) * rnd + minY)
    DMDPosSwarmers(i,1) = rY
  Next

  'bombers
  For i = 0 To 3
    rX = Int((maxX - minX + 1) * rnd + minX)
    DMDPosBombers(i,0) = rX
    rY = Int((maxY - minY + 1) * rnd + minY)
    DMDPosBombers(i,1) = rY
  Next


End Sub

' nvram data variables
Dim BallsPerGame, MatchEnabled, FreePlayEnabled

Sub NvRam_Init
  ' see https://www.vpforums.org/index.php?showtopic=33611&p=346441
  Dim NVRAM
  NVRAM = Controller.NVRAM
  If Not IsEmpty(NVRAM) Then

    BallsPerGame = (NVRAM(148) - 240) '(NVRAM(147) - 240) & (NVRAM(148) - 240)

    If (NVRAM(140) - 240) Mod 2 = 0 Then
      MatchEnabled = True
    Else
      MatchEnabled = False
    End If

    If CStr((NVRAM(171) - 240) & (NVRAM(172) - 240)) = "00" Then
      FreePlayEnabled = True
    Else
      FreePlayEnabled = False
    End If

  End If
End Sub







' CHANGE LOG
' 001 - apophis - Added Fleep sound package
' 002 - apophis - Added nFozzy flippers and physics
' 003 - apophis - Added Roth drop targts for sw33,34,35,39,40. Implemented a fast flips workaround. Consolidated some timers into the GameTimer.
' 004 - apophis - Updated playfield image to an HD playfield scan. Realigned all table objects. Reformatted script. Updated ball rolling and drop target sounds.
' 005 - apophis - Fixed issue with right lane return. Updated playfield and plastic images with the scan provided by g94.
' 006 - apophis - Added Lampz and 3D inserts. Updated targetbouncer and rubberizer code.
' 007 - apophis - Adjusted some insert lighting. Adjusted slingshot strength.
' 008 - apophis - Updated PF and insert overlay images.
' 009 - apophis - Reworked insert lights and blooms...TBC
' 010 - apophis - Added physical trough and lock. No destroying the balls. Updated some DT textures.
' 011 - apophis - Reworked insert light objects. Changed rubberizer.
' 012 - apophis - Tuned kicker strength. Preliminary center flasher functionality. Updated drop target code to latest.
' 013 - apophis - Tweaked Wall5. Rotated rollovers 180 deg.
' 014 - apophis - Added plunger lane ramps. Reduced kickback strength. Removed all getball calls. Tuned some inserts.
' 015 - bord - Turned all walls and rubbers to visible=0, removed existing meshes, moved GI beneath pf, small wall tuning to match meshes, added all meshes w/ rendered textures
' 016 - apophis - Created ON_Prim and OFF_Prim collections. Hooked up ON and OFF material fading to GI
' 017 - bord - bracketed gate & wire meshes added, center flasher mesh, flashers and texture, credit light, drop target shadows, new target meshes and textures
' 018 - apophis - Roth drop targets and standup targets added/updated. Wired center flasher to lampz and tuned. Updated rubberizer and flipper nudge.
' 019 - apophis - Dynamic shadows. Drop target shadows.
' 020 - apophis - Fixed issue where playfield not fading properly during GI state changes.
' 021 - UnclePaulie - Added hybrid mode. VR environments, posters, clock, and topper. Animated Backglass for VR and desktop. Changed drop targets to drop at playfield level. Modified groove to avoid stuck ball at diverter.
' 022 - apophis - Animated gatewire. Set plunger pull speed to 0.2. Tied some prim DL values to GI state. Added instruction card zoom feature (press R button). General file cleanup.
' RC1 - apophis - Added a few more ball images, and changed default one. Adjusted wall near upper flipper. Reduced kickback strength a little. Added fire button for Bombs. Updated TimerVRPlunger interval. Updated nudge strengths. Updated DT backdrop.
' RC2 - apophis - Updated opacity of VR alphanumeric display. Added flipper shadows. Updated POV. Tuned some insert halos.
' RC3 - apophis - Adjusted height of wires below flippers.
' Release 1.0
' 1.0.1 - jsm174 - Standalone VPX compatibility updates
' 1.0.2 - Primetime5k - Staged flipper update
' 1.0.3 - apophis - Added slingshot corrections. Target physics material updated. Automated VR_Room. Fixed issue with playfield_mesh in ON_Prims collection (for 10.8 compatibility). Organized layers.
' 1.0.4 - Primetime5k - minor staged flipper fix.
' 1.0.5 - apophis - Unchecked Spherical Map on ball image mapping. Changed ball scratch image. Fix for standalone (thanks jsm174)
' 1.0.6 - apophis - Fix SolCallback(36) issue.
' Release 1.1
' Release 1.1.1 - apophis - Fix playfleid relections: deleted old playfeild_mesh and renamed playfeild_mesh_on to playfeild_mesh. Put center post in Rubbers collection. Deleted old plastics flasher.
' 1.1.2 - scutters - Added optional Flex DMD enhancements
' 1.1.3 - apophis - VolumeDial fix. Added FlipperCradleCollision.
' 1.1.4 - passion4pins - Desktop updates
' 1.1.5 - apophis - Added new "Tweak" options menu
' 1.1.6 - apophis - added LUT (thanks HauntFreaks). Updated desktop POV. Fixed some playfield reflections.
' Release 1.2
' Release 1.2.1 - apophis - Fixed VR Room options setup
' 1.2.2 - apophis - Updated flipper triggers and scripts. VPX standalone fixes (thanks jsm174). Moved rules card to Tweak menu. FlexDMD fix from Scutters. Fixed glass heights. Added Defender screenshot. Disabled "hide parts behind" for ball and flipper shadows.
' Release 1.3
' 1.3.1 - apophis - Removed old scorecard code. Fortify walls against stuck balls. Raise bumper strength a little.
' 1.3.2 - apophis - Updated DisableStaticPreRendering functionality to be consistent with VPX 10.8.1 API
' 1.3.3 - apophis - Added default 1to1 LUT to fix anaglyph issue.
' 1.3.4 - scutters - Added updated FlexDMD options: It now has a scanner like the video game (lamp and drop target states of the table for the aliens etc), also new full DMD option.
' 1.3.5 - apophis - Minor fixes to get updated flexDMD options to work with VR.
' 1.3.6 - apophis - More fixes for VR options. Fixed plunger groove ramp.
' 1.3.7 - scutters - Fix for the "match" text on FlexDMD.
' 1.3.8 - apophis - Fix VR button animations and lighting. Updated ball image.
' Release 1.4

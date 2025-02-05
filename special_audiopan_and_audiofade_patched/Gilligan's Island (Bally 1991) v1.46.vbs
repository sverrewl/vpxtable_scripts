'  ________.__.__  .__  .__                  /\         .___       .__                     .___
' /  _____/|__|  | |  | |__| _________    ___)/  ______ |   | _____|  | _____    ____    __| _/
'/   \  ___|  |  | |  | |  |/ ___\__  \  /    \ /  ___/ |   |/  ___/  | \__  \  /    \  / __ |
'\    \_\  \  |  |_|  |_|  / /_/  > __ \|   |  \\___ \  |   |\___ \|  |__/ __ \|   |  \/ /_/ |
' \______  /__|____/____/__\___  (____  /___|  /____  > |___/____  >____(____  /___|  /\____ |
'        \/               /_____/     \/     \/     \/           \/          \/     \/      \/

' Based on the original Visual Pinball mod from Dark and mfuegemann (2016)
' Further refined on rothbauerw physics and subway models
' VR Room based on Arvid's standalone VR Room

' TastyWasps - November 2023
' nFozzy/Roth Physics, Fleep Sounds, TargetBouncer, VR Room
' GI Lighting Improvements, Dynamic Ball Shadows, Multi-Ball Cradle Physics
' Modulated Flashers, LUT Selector, Gameplay Balance
' Physics - rothbauerw
' Scripting Tweaks - apophis, Sixtoe
' Playfield Upscale - redbone
' VR Upscales / B2S Backglass - Hauntfreaks
' VR Backglass - leojreimroc
' Testing - bietekwiet, bountybob, Studlygoorite, mcarter78, Primetime5k, passion4pins, PinStratsDan

Option Explicit
Randomize

'********************************************************
'           TABLE OPTIONS
'********************************************************

'----- Volume Options -----
Const VolumeDial = 0.8          ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5        ' Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5        ' Level of ramp rolling volume. Value between 0 and 1

'----- Shadow Options -----     ' If performance is suffering, turn off Dynamic Shadow options below as they are subtle.
Const DynamicBallShadowsOn = 1    ' 0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   ' 0 = Static shadow under ball ("flasher" image, like JP's)
                  ' 1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                  ' 2 = flasher image shadow, but it moves like ninuzzu's

'----- LED Light Options ----
Const GIColorLED = False      ' False = Normal GI, True = Color LEDs

'----- VPM DMD Option -----
Const ShowVPMDMD = True       ' False = Do not show native VPinMame DMD in desktop/VR, True = Show native VPinMame DMD in desktop/VR

'----- VR Options -----
Const VRRoomChoice = 1        ' 1 = Gilligan Room , 2 = Minimal Room, 3 = Pass Through Sphere Room
Const DMDReflection = 0       ' 0 = DMD Reflection off glass, 1 = DMD Reflection on glass

'----- Title Song Option -----
Const TitleSong = True        ' False = Do not play title song upon load of table, True = Play title song upon load of table

'******************************************************
'           STANDARD DEFINITIONS
'******************************************************

Dim BallMass ,BallSize
Ballsize = 50
Ballmass = 1

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 1
Const tnob = 5
Const lob = 0
Dim UseVPMModSol, ModFlashers

' Standard Sounds
Const SSolenoidOn = "":Const SSolenoidOff = "":Const SFlipperOn = "":Const SFlipperOff = "":Const SCoin = ""

Dim tablewidth, tableheight
tablewidth = Gilligan.width
tableheight = Gilligan.height

Dim DesktopMode, UseVPMDMD: DesktopMode = Gilligan.ShowDT

If DesktopMode = True Then ' Show Desktop components
  Ramp16.visible=1
  Ramp12.visible=1
  SideWood.visible=1
  SideWood1.visible=0
  If ShowVPMDMD = True Then
    UseVPMDMD = True
  End If
Else
  Ramp16.visible=0
  Ramp12.visible=0
  SideWood.visible=0
  SideWood1.visible=1
  UseVPMDMD = False
End If

' Detect if version 10.8 or less for flasher routines
If VersionMajor = "10" Then
  If VersionMinor = "8" Then
    ModFlashers = True
    UseVPMModSol = True
  Else
    ModFlashers = False
    UseVPMModSol = False
  End If
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "WPC.VBS", 3.26

If TitleSong = True Then
  PlaySound "title_song", -1
End If

Const cGameName = "gi_l9"

' VR Room Auto-Detect
Dim VR_Obj, VRRoom

If RenderingMode = 2 Then VRRoom = 1 Else VRRoom = 0
If VRRoom = 1 Then
  Ramp12.visible = 0
  Ramp16.visible = 0
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  If VRRoomChoice = 1 Then
    For Each VR_Obj in VRGilliganRoom : VR_Obj.Visible = 1 : Next
  ElseIf VRRoomChoice = 2 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
  Else
    For Each VR_Obj in VRSphereRoom : VR_Obj.Visible = 1 : Next
  End If
  Primary_Backglass.BlendDisableLighting = 0.45
  If DMDReflection = 1 Then
    Primary_DMD_Reflection.visible = 1
  Else
    Primary_DMD_Reflection.visible = 0
  End If
  For Each VR_Obj in VRBackglass : VR_Obj.Visible = 1 : Next
  Setbackglass
  If ShowVPMDMD = True Then
    UseVPMDMD = True
  End If
End If

Dim mIsland, GIBall1, GIBall2

' Table Init
Sub Gilligan_Init
  vpmInit Me
  With Controller
    .GameName = cGameName
    'If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Gilligan's Island"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = DesktopMode
    .Dip(0) = &H00
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
    .Switch(22) = 1 'close coin door
    .Switch(24) = 1 'and keep it close
  End With

  'Nudging
  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 1
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  for each obj in LightSetNormal
    obj.visible = not GIColorLED
  Next

  for each obj in LightSetColor
    obj.visible = GIColorLED
  Next

  'set initial textures
  Primitive_PlasticRamp.image = "PlasticRampMap"
  Primitive_RampHexPost1.image = "ParrotPost_off"
  Primitive_PlasticsCollection3.image = "PlasticsCollection3_off"
  Primitive_ProfessorPlastic.image = "ProfessorPlasticOFF"
  Primitive_bumpercap3.image = "BumperMap_OFF"
  Primitive_bumpercap2.image = "BumperMap_OFF"
  Primitive_bumpercap1.image = "BumperMap_OFF"
  Primitive_Kona.image = "KonaMap"
  P_BridgeRamp.image = "MechRamp_Map2"
  Primitive_PalmPostP.image = "PalmTreePlastic(sm)Map"
  Primitive_KonaFlasher.image = "dome4_red"

  '************  Trough **************************
  Set GIBall1 = sw16.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set GIBall2 = sw17.CreateSizedballWithMass(Ballsize/2,Ballmass)

  Controller.Switch(16) = 1
  Controller.Switch(17) = 1

  Set mIsland=New cvpmmech
  With mIsland
    .mtype=vpmMechOneSol+vpmMechCircle+vpmMechLinear'onedirsol
    .sol1=9
    'Sw77 is the opto - It's mostly off
    'There is one gap for each stop - and a long gap for the home position
    'The switch would be 'on' during the gaps and off otherwise
    'I think each position is just after a pulse of the opto - since the home is a long pulse
    'so it must stop when the switch turns back off

    .length=120 '2 sec to find the new Position
    .steps=34
    .AddSw 77,1,6 ' First on gap is 1 position before key 1'
    .AddSw 77,8,13 ' Second on gap is 1 position before key 2'
    .AddSw 77,15,20 ' Third on gap is 1 position before key 3'
    .AddSw 77,22,27 ' Fourth on gap is 1 position before key 4'
    .AddSw 77,29,33 ' This is the 'long' gap on for several positions before key 0'
    .callback=getref("UpdateIsland")
    .start
  End With

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  LeftAutoPlunger.Pullback
  KickBack.Pullback

  sw66wall.collidable = false
End Sub

Sub Gilligan_Paused:Controller.Pause = True:End Sub
Sub Gilligan_unPaused:Controller.Pause = False:End Sub
Sub Gilligan_Exit:SaveLUT:Controller.Pause = False:Controller.Stop:End Sub

'******************************************************
'             KEYS
'******************************************************

Sub Gilligan_KeyDown(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull()
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
    JungleActive = True
  End If

  If keycode = LeftFlipperKey Then
    LFPress = 1
    Primary_flipperbuttonleft.X = Primary_flipperbuttonleft.X + 10
  End If

  If keycode = RightFlipperKey Then
    RFPress = 1
    Primary_flipperbuttonright.X = Primary_flipperbuttonright.X - 10
  End If

  If keycode = LeftTiltKey Then
    SoundNudgeLeft()
    Nudge 90, 1
  End If

  If keycode = RightTiltKey Then
    SoundNudgeRight()
    Nudge 270, 1
  End If

  If keycode = CenterTiltKey Then
    SoundNudgeCenter()
    Nudge 0, 1.5
  End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If keycode = StartGameKey Then
    soundStartButton()
    Primary_startbutton1.y = Primary_startbutton1.y - 5
    Primary_startbutton2.y = Primary_startbutton2.y - 5
  End If

  If KeyCode = LeftMagnaSave Then
    bLutActive = True
  End If

  If KeyCode = RightMagnaSave Then
    If bLutActive = True Then
      If DisableLUTSelector = 0 Then
        LUTSet = LUTSet + 1

        If LutSet > 16 Then
          LUTSet = 0
        End If

        If LutSet <> 12 Then
          Playsound "click", 0, 1, 0.1, 0.25
        End If

        If LutSet = 12 Then
          Playsound "gun", 0, 1, 0.1, 0.25
        End If

        LutSlctr.Enabled = True

        SetLUT
        ShowLUT
      End If
    End If
  End If

  If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub Gilligan_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    Primary_plunger.Y = 1269.286
  End If

  If keycode = LeftFlipperKey Then
    LFPress = 0
    leftflipper.eostorqueangle = EOSA
    leftflipper.eostorque = EOST
    Primary_flipperbuttonleft.X = Primary_flipperbuttonleft.X - 10
  End If

  If keycode = RightFlipperKey Then
    RFPress = 0
    rightflipper.eostorqueangle = EOSA
    rightflipper.eostorque = EOST
    Primary_flipperbuttonright.X = Primary_flipperbuttonright.X + 10
  End If

  If Keycode = StartGameKey Then
    Primary_startbutton1.y = Primary_startbutton1.y + 5
    Primary_StartButton2.y = Primary_StartButton2.y + 5
    If TitleSong = True Then
      StopSound "title_song"
    End If
  End If

  ' Left Magna has to be held down with Right Magna to cycle through LUTs.
  If keycode = LeftMagnaSave Then bLutActive = False

  If vpmKeyUp(keycode) Then Exit Sub

End Sub

'******************************************************
'             SOLENOIDS
'******************************************************

SolCallback(1) = "SolLeftAutoPlunger"     'Lagoon Kicker (Left Lock)
'SolCallback(2) =                 'Island Lock
SolCallback(3) = "SolOutHole"         'Outhole
SolCallback(4) = "SolVUKPopper"         'VUK Popper
                        '5  RightSlingShot
                        '6  LeftSlingShot
SolCallback(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker)," 'Knocker
SolCallback(8) = "SolKickBack"          'sKickBack
'SolCallback(9) = IslandMotor handled by Mech
SolCallback(10) = "SolRelease"          'BallRelease
SolCallback(11) = "SolHoldLock"         'Hold Lock
                        '13 Bumper
                        '14 Bumper
                        '15 Bumper
SolCallback(16) = "SolTopKicker"        'Top Kicker
SolCallback(23) = "SolRampUp"         'RampUp
SolCallback(24) = "SolRampDown"         'RampDown
SolCallback(25) = "SetLamp 125,"        'Right Bank Flasher
'SolCallback(26) = "vpmFlasher xxx,"      'Treasure Flasher
'SolCallback(27) = "vpmFlasher xxx,"      'Title Flasher - Backbox?
'SolCallback(28) = "vpmFlasher xxx,"      'Professor Flasher - Backbox?
                        '29 Right Flipper
                        '30 Left Flipper
SolCallback(31) = "SolGameOn"         'GameOn

' Flasher Solenoids
If ModFlashers = True Then
  SolModCallback(12) = "SolMod12"         ' Island Light
  SolModCallback(17) = "SolMod17"         ' Head 1 Flasher
  SolModCallback(18) = "SolMod18"         ' Island Flasher
  SolModCallback(19) = "SolMod19"         ' Left Bank Flasher
  SolModCallback(20) = "SolMod20"         ' Left Lane Flasher
  SolModCallback(21) = "SolMod21"         ' Right Lane Flasher
  SolModCallback(22) = "SolMod22"         ' Head 2 Flasher
  SolModCallback(25) = "SolMod25"         ' Right Bank Flasher
  SolModCallback(26) = "SolMod26"         ' Insert Center title
  SolModCallback(27) = "SolMod27"         ' Insert Sides TitleSong
  SolModCallback(28) = "SolMod28"         ' Insert Professor
Else
  SolCallback(12) = "Sol12"           ' Island Light
  SolCallback(17) = "Sol17"           ' Head 1 Flasher
  SolCallback(18) = "SetLamp 118,"        ' Island Flasher
  SolCallback(19) = "SetLamp 119,"        ' Left Bank Flasher
  SolCallback(20) = "SetLamp 120,"        ' Left Lane Flasher
  SolCallback(21) = "SetLamp 121,"        ' Right Lane Flasher
  SolCallback(22) = "Sol22"           ' Head 2 Flasher
  SolCallback(25) = "SetLamp 125,"        ' Right Bank Flasher
End If

'******************************************************
'       TOP LEFT AUTOPLUNGER
'******************************************************

Sub SolLeftAutoPlunger(enabled)
  if enabled Then
    LeftAutoPlunger.Fire
    PlaySoundAt SoundFX("popper_ball",DOFContactors), LeftAutoPlunger
  Else
    LeftAutoPlunger.Pullback
  End If
End Sub

'******************************************************
'           TROUGH
'******************************************************

Sub sw16_Hit():Controller.Switch(16) = 1:UpdateTrough:End Sub
Sub sw16_UnHit():Controller.Switch(16) = 0:UpdateTrough:End Sub
Sub sw17_Hit():Controller.Switch(17) = 1:UpdateTrough:End Sub
Sub sw17_UnHit():Controller.Switch(17) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw17.BallCntOver = 0 Then sw16.kick 60, 9
  Me.Enabled = 0
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub sw18_Hit() ' Drain
  Controller.Switch(18) = 1
  RandomSoundDrain sw18
End Sub

Sub sw18_UnHit()  ' Drain
  Controller.Switch(18) = 0
End Sub

Sub SolOutHole(enabled)
  If enabled then
    If sw16.BallCntOver = 0 Then sw18.kick 60, 20
    UpdateTrough
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    If sw17.BallCntOver = 0 Then
      RandomSoundBallRelease sw17
    Else
      PlaySoundAt SoundFX("BallRelease",DOFContactors), sw17
    End If
    sw17.kick 60, 9
  End If
End Sub

'******************************************************
'         VUK POPPER
'******************************************************

Sub Vuk_Hit()
  sw66wall.collidable = true
  SoundSaucerLock
  Controller.Switch(66) = 1
End Sub

Sub SolVUKPopper(Enabled)
  If Enabled Then
    vuk.kick 155,44,1.56
    sw66wall.collidable = false
    SoundSaucerKick 1, vuk
    Controller.Switch(66) = 0
  End If
End Sub

'******************************************************
'         TOP KICKER
'******************************************************

Sub TopKicker_Hit()
  SoundSaucerLock
  Controller.Switch(67) = True
  JungleActive = True
End Sub

Sub SolTopKicker(enabled)
  If Enabled Then
    If TopKicker.BallCntOver > 0 Then
      TopKicker.kick 172 + rnd*6,15+rnd*4
      SoundSaucerKick 1, TopKicker
    Else
      PlaySoundAt SoundFX("solon",DOFContactors), TopKicker
    End If
    Primitive_BallEjectArm.RotX = 50
    TopKicker.Timerenabled = True
    Controller.Switch(67) = False
  End If
End Sub

Sub TopKicker_Timer
  Primitive_BallEjectArm.RotX = Primitive_BallEjectArm.RotX - 5
  if Primitive_BallEjectArm.RotX <= 90 Then
    TopKicker.Timerenabled = False
    Primitive_BallEjectArm.RotX = 90
  End If
End Sub

'******************************************************
'            KICK BACK
'******************************************************

Sub SolKickBack(enabled)
  If enabled Then
    KickBack.Fire
    PlaySoundAt SoundFX("popper_ball",DOFContactors), KickBack
  Else
    KickBack.Pullback
  End If
End Sub

'******************************************************
'            RAMP
'******************************************************

Sub SolRampDown(Enabled)
  If Enabled  Then
    RampDir = 1
    RampBridgeTimer.enabled = True
    Controller.Switch(62)=1
    playsound SoundFX("Soloff",DOFContactors),0,1,-0.08,0.25
  End If
End Sub

Sub SolRampUp(Enabled)
  If Enabled then
    RampDir = -1
    RampBridgeTimer.enabled = True
    Controller.Switch(62)=0
    playsound SoundFX("Soloff",DOFContactors),0,1,-0.08,0.25
  End If
End Sub


'******************************************************
'         FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const cSingleLFlip = False
Const cSingleRFlip = False

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
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  RightFlipperCollide parm
End Sub

RightFlipper.timerinterval=1
RightFlipper.timerenabled=True

Sub RightFlipper_timer()

  If leftflipper.currentangle = leftflipper.endangle and LFPress = 1 then
    leftflipper.eostorqueangle = EOSAnew
    leftflipper.eostorque = EOSTnew
    LeftFlipper.rampup = EOSRampup
    if LFCount = 0 Then LFCount = GameTime
    if GameTime - LFCount < LiveCatch Then
      leftflipper.Elasticity = 0.1
      If LeftFlipper.endangle <> LFEndAngle Then leftflipper.endangle = LFEndAngle
    Else
      leftflipper.Elasticity = FElasticity
    end if
  elseif leftflipper.currentangle > leftflipper.startangle - 0.05  Then
    leftflipper.rampup = SOSRampup
    leftflipper.endangle = LFEndAngle - 3
    leftflipper.Elasticity = FElasticity
    LFCount = 0
  elseif leftflipper.currentangle > leftflipper.endangle + 0.01 Then
    leftflipper.eostorque = EOST
    leftflipper.eostorqueangle = EOSA
    LeftFlipper.rampup = Frampup
    leftflipper.Elasticity = FElasticity
  end if

  If rightflipper.currentangle = rightflipper.endangle and RFPress = 1 then
    rightflipper.eostorqueangle = EOSAnew
    rightflipper.eostorque = EOSTnew
    RightFlipper.rampup = EOSRampup
    if RFCount = 0 Then RFCount = GameTime
    if GameTime - RFCount < LiveCatch Then
      rightflipper.Elasticity = 0.1
      If RightFlipper.endangle <> RFEndAngle Then rightflipper.endangle = RFEndAngle
    Else
      rightflipper.Elasticity = FElasticity
    end if
  elseif rightflipper.currentangle < rightflipper.startangle + 0.05 Then
    rightflipper.rampup = SOSRampup
    rightflipper.endangle = RFEndAngle + 3
    rightflipper.Elasticity = FElasticity
    RFCount = 0
  elseif rightflipper.currentangle < rightflipper.endangle - 0.01 Then
    rightflipper.eostorque = EOST
    rightflipper.eostorqueangle = EOSA
    RightFlipper.rampup = Frampup
    rightflipper.Elasticity = FElasticity
  end if

End Sub

dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSAnew
dim FStrength, Frampup, FElasticity, EOSRampup, SOSRampup
dim RFEndAngle, LFEndAngle, LFCount, RFCount, LiveCatch

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
FStrength = LeftFlipper.strength
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
EOSTnew = 1.2 'FEOST
EOSAnew = 0.2
EOSRampup = 1.5
SOSRampup = 2.5 '8.5
LiveCatch = 8

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle


Set GICallBack = GetRef("UpdateGI")

Sub SolGameOn(enabled)
  VpmNudge.SolGameOn(enabled)
End Sub

'******************************************************
'         FLASHERS
'******************************************************
Dim p

Sub Sol12(enabled)
  SetLamp 112,enabled
  UpdateKona
End Sub

Sub SolMod12(pwm)
  p = pwm / 255.0
  f112.state = p
  UpdateKona
End Sub

Sub Sol17(enabled)
  SetLamp 117,enabled
  UpdateKona
End Sub

Sub SolMod17(pwm)
  p = pwm / 255.0
  f117.state = p
  UpdateKona
End Sub

Sub SolMod18(pwm)
  p = pwm / 255.0
  f118.state = p
End Sub

Sub SolMod19(pwm)
  p = pwm / 255.0
  f119.state = p
End Sub

Sub SolMod20(pwm)
  p = pwm / 255.0
  f120.state = p
End Sub

Sub SolMod21(pwm)
  p = pwm / 255.0
  f121.state = p
End Sub

Sub Sol22(enabled)
  SetLamp 122,enabled
  UpdateKona
End Sub

Sub SolMod22(pwm)
  p = pwm / 255.0
  f122.state = p
End Sub

Sub SolMod25(pwm)
  p = pwm / 255.0
  f125.state = p
End Sub

Sub SolMod26(pwm)
  p = pwm / 255.0
  F126.state = p
End Sub

Sub SolMod27(pwm)
  p = pwm / 255.0
  F127.state = p
End Sub

Sub SolMod28(pwm)
  p = pwm / 255.0
  F128.state = p
End Sub

'******************************************************
'         GI
'******************************************************
dim obj
Sub UpdateGI(GINo,Status)

  Select Case GINo
    ' Left Inserts
    Case 0:
      If status Then
        If GIColorLED Then
          for each obj in ColorGIString1
            obj.state = lightstateon
          next
        Else
          for each obj in GIString1
            obj.state = lightstateon
          next
        End If
        GI1.state = Status
'     For each obj in VRBGGIRight : obj.visible = 1 : Next
      Else
        If GIColorLED Then
          for each obj in ColorGIString1
            obj.state = lightstateoff
          next
        Else
          for each obj in GIString1
            obj.state = lightstateoff
          next
        End If
        GI1.state = Status
'     For each obj in VRBGGIRight : obj.visible = 0 : Next
      End If
    ' Bottom Playfield
    Case 1
      If status Then
        If GIColorLED Then
          for each obj in ColorGIString2
            obj.state = lightstateon
          next
        Else
          for each obj in GIString2
            obj.state = lightstateon
          next
        End If

      Else
        If GIColorLED Then
          for each obj in ColorGIString2
            obj.state = lightstateoff
          next
        Else
          for each obj in GIString2
            obj.state = lightstateoff
          next
        End If
      End If
    ' Middle Playfield
    Case 2
      If status then
        If GIColorLED Then
          for each obj in ColorGIString3
            obj.state = lightstateon
          next
        Else
          for each obj in GIString3
            obj.state = lightstateon
          next
        End If
      Else
        If GIColorLED Then
          for each obj in ColorGIString3
            obj.state = lightstateoff
          next
        Else
          for each obj in GIString3
            obj.state = lightstateoff
          next
        End If
      End if
    ' Right Inserts
    Case 3
      If status Then
        If GIColorLED Then
          for each obj in ColorGIString4
            obj.state = lightstateon
          next
        Else
          for each obj in GIString4
            obj.state = lightstateon
          next
        End If
'       For each obj in VRBGGILeft : obj.visible = 1 : Next
      Else
        If GIColorLED Then
          for each obj in ColorGIString4
            obj.state = lightstateoff
          next
        Else
          for each obj in GIString4
            obj.state = lightstateoff
          next
        End If
'       For each obj in VRBGGILeft : obj.visible = 0 : Next
      End if
    ' Top Playfield
    case 4
      If status Then
        If GIColorLED Then
          for each obj in ColorGIString5
            obj.state = lightstateon
          next
        Else
          for each obj in GIString5
            obj.state = lightstateon
          next
        End If
        Primitive_PlasticRamp.image = "PlasticRampMapON"
        Primitive_RampHexPost1.image = "ParrotPost"
        Primitive_PlasticsCollection3.image = "PlasticsCollection3"
        Primitive_ProfessorPlastic.image = "ProfessorPlastic-Map"
        Primitive_bumpercap3.image = "BumperMap_ON"
        Primitive_bumpercap2.image = "BumperMap_ON"
        Primitive_bumpercap1.image = "BumperMap_ON"
      Else
        If GIColorLED Then
          for each obj in ColorGIString5
            obj.state = lightstateoff
          next
        Else
          for each obj in GIString5
            obj.state = lightstateoff
          next
        End If
        Primitive_PlasticRamp.image = "PlasticRampMap"
        Primitive_RampHexPost1.image = "ParrotPost_off"
        Primitive_PlasticsCollection3.image = "PlasticsCollection3_off"
        Primitive_ProfessorPlastic.image = "ProfessorPlasticOFF"
        Primitive_bumpercap3.image = "BumperMap_OFF"
        Primitive_bumpercap2.image = "BumperMap_OFF"
        Primitive_bumpercap1.image = "BumperMap_OFF"
      End if
  End Select
End Sub

Sub LeftOutlane_Hit:Controller.Switch(25) = True:End Sub
Sub LeftOutlane_Unhit:Controller.Switch(25) = False:End Sub
Sub LeftInlane_Hit:Controller.Switch(26) = True:End Sub
Sub LeftInlane_Unhit:Controller.Switch(26) = False:End Sub
Sub RightInlane_Hit:Controller.Switch(27) = True:End Sub
Sub RightInlane_Unhit:Controller.Switch(27) = False:End Sub
Sub RightOutlane_Hit:Controller.Switch(28) = True:End Sub
Sub RightOutlane_Unhit:Controller.Switch(28) = False:End Sub

Sub SW31_Hit:Controller.Switch(31) = True:End Sub
Sub SW31_Unhit:Controller.Switch(31) = False:End Sub
Sub SW33_Hit:Controller.Switch(33) = True:End Sub
Sub SW33_Unhit:Controller.Switch(33) = False:End Sub

Sub SW61a_Hit:Controller.Switch(61) = True:End Sub
Sub SW61a_Unhit:Controller.Switch(61) = False:End Sub
Sub SW61b_Hit:Controller.Switch(61) = True:End Sub
Sub SW61b_Unhit:Controller.Switch(61) = False:End Sub
Sub SW63_Hit:Controller.Switch(63) = True:End Sub
Sub SW63_Unhit:Controller.Switch(63) = False:End Sub
Sub SW64_Hit:Controller.Switch(64) = True:End Sub
Sub SW64_Unhit:Controller.Switch(64) = False:End Sub
Sub SW65_Hit:Controller.Switch(65) = True:End Sub
Sub SW65_Unhit:Controller.Switch(65) = False:End Sub
Sub SW68_Hit:Controller.Switch(68) = True:End Sub
Sub SW68_Unhit:Controller.Switch(68) = False:End Sub

Sub SW78_Hit:Controller.Switch(78) = True:End Sub
Sub SW78_Unhit:Controller.Switch(78) = False:End Sub
Sub SW75_Hit:Controller.Switch(75) = True:End Sub
Sub SW75_Unhit:Controller.Switch(75) = False:End Sub

Sub SW83_Hit:Controller.Switch(83) = True:End Sub
Sub SW83_Unhit:Controller.Switch(83) = False:End Sub
Sub SW84_Hit:Controller.Switch(84) = True:End Sub
Sub SW84_Unhit:Controller.Switch(84) = False:End Sub

Sub SW71_Hit:vpmTimer.PulseSw 71:End Sub
Sub SW72_Hit:vpmTimer.PulseSw 72:End Sub
Sub SW73_Hit:Controller.Switch(73) = True:Primitive_Switcharm73.rotx = 27:End Sub
Sub SW73_Unhit:Controller.Switch(73) = False:SW73.Timerenabled = True:End Sub
Sub SW74_Hit:Controller.Switch(74) = True:Primitive_Switcharm74.rotx = 27:End Sub
Sub SW74_Unhit:Controller.Switch(74) = False:SW74.Timerenabled = True:End Sub

Sub SW73_Timer
  Primitive_Switcharm73.rotx = Primitive_Switcharm73.rotx - 1
  if Primitive_Switcharm73.rotx <= 0 Then
    SW73.Timerenabled = False
    Primitive_Switcharm73.rotx = 0
  End If
End Sub

Sub SW74_Timer
  Primitive_Switcharm74.rotx = Primitive_Switcharm74.rotx - 1
  if Primitive_Switcharm74.rotx <= 0 Then
    SW74.Timerenabled = False
    Primitive_Switcharm74.rotx = 0
  End If
End Sub

' Rubbers
Sub SW32_Hit:vpmTimer.PulseSw 32:End Sub
Sub SW58_Hit:vpmTimer.PulseSw 58:End Sub


Sub Target35_Hit:vpmTimer.PulseSw 35:MoveTarget35:End Sub
Sub Target36_Hit:vpmTimer.PulseSw 36:MoveTarget36:End Sub
Sub Target37_Hit:vpmTimer.PulseSw 37:MoveTarget37:End Sub
Sub Target38_Hit:vpmTimer.PulseSw 38:MoveTarget38:End Sub

Sub Target46_Hit:vpmTimer.PulseSw 46:MoveTarget46:End Sub
Sub Target47_Hit:vpmTimer.PulseSw 47:MoveTarget47:End Sub
Sub Target48_Hit:vpmTimer.PulseSw 48:MoveTarget48:End Sub

Sub Target51_Hit:vpmTimer.PulseSw 51:MoveTarget51:End Sub
Sub Target52_Hit:vpmTimer.PulseSw 52:MoveTarget52:End Sub
Sub Target53_Hit:vpmTimer.PulseSw 53:MoveTarget53:End Sub
Sub Target54_Hit:vpmTimer.PulseSw 54:MoveTarget54:End Sub
Sub Target55_Hit:vpmTimer.PulseSw 55:MoveTarget55:End Sub
Sub Target56_Hit:vpmTimer.PulseSw 56:MoveTarget56:End Sub
Sub Target57_Hit:vpmTimer.PulseSw 57:MoveTarget57:End Sub

'#######################################################################
' Jungle Ramp

Sub ToggleRamp
  if not RampBridgeTimer.enabled Then
    RampDir = -RampDir
    RampBridgeTimer.enabled = True
  end If
end Sub

Dim RampDir
RampDir = -1

Const RampSpeed= 0.5
JungleRampDown.collidable = False

Sub RampBridgeTimer_Timer
  Primitive_MechArm.ObjRotX = Primitive_MechArm.ObjRotX + (RampDir * RampSpeed)
  Primitive_MechArm.ObjRotY = abs(Primitive_MechArm.ObjRotX)/10

  if Primitive_MechArm.ObjRotX >= -30 then
    JungleRampDown.collidable = True
    JungleRampUp.collidable = False
  Else
    JungleRampUp.collidable = True
    JungleRampDown.collidable = False
  End If

  if Primitive_MechArm.ObjRotX >= 4 then
    RampBridgeTimer.enabled = False
    Primitive_MechArm.ObjRotX = 4
  end if
  if Primitive_MechArm.ObjRotX <= -45 then
    RampBridgeTimer.enabled = False
    Primitive_MechArm.ObjRotX = -45
  end if
  P_BridgeRamp.Rotx = RampRot(Primitive_MechArm.ObjRotX)
End Sub

Dim ArmAngle
Const Pi=3.141592654
'P_BridgeRamp.Rotx = -RampRot(Primitive_MechArm.ObjRotX)

Function RampRot(ArmRotPar)
  ArmAngle = 122 + ArmRotPar
  RampRot = (1-Sin(ArmAngle/180*Pi))*100
  if RampRot > 14.9 Then
    RampRot = 14.9
  end If
End Function

'#############################################################################
'Jungle Turntable

'** Update Island
Dim islpos,IslandRotation

Sub UpdateIsland(anewpos,aspeed,alastpos)
  if JungleActive Then
    islpos=anewpos
    Select Case islpos
      Case 33,34,0,1:   IslandRotation = 357  'Ramp_1
      Case 6,7,8:  IslandRotation = 283   'Ramp_2
      Case 13,14,15: IslandRotation = 211   'Ramp_3
      Case 20,21,22: IslandRotation = 140   'Ramp_4
      Case 27,28,29: IslandRotation = 67    'Ramp_5
    End Select
    TurntableGate.collidable = True
    TurnTableTimer.enabled = True
    TurnTableSoundTimer.enabled = True
  end If
End Sub

'Avoid turning wheel at start of table
Dim JungleActive
JungleActive = False

'Init Paths
ToggleTurnTablePaths

Sub ToggleTurnTablePaths
  JungleRamp_1.collidable = (Primitive_TurnTable.ObjRotZ > 346) Or (Primitive_TurnTable.ObjRotZ < 8)    '357
  JungleRamp_2.collidable = (Primitive_TurnTable.ObjRotZ > 272) and (Primitive_TurnTable.ObjRotZ < 294) '283
  JungleRamp_3.collidable = (Primitive_TurnTable.ObjRotZ > 200) and (Primitive_TurnTable.ObjRotZ < 222) '211
  JungleRamp_4.collidable = (Primitive_TurnTable.ObjRotZ > 129) and (Primitive_TurnTable.ObjRotZ < 151) '140
  JungleRamp_5.collidable = (Primitive_TurnTable.ObjRotZ > 56) and (Primitive_TurnTable.ObjRotZ < 78)   ' 67
End Sub

Sub TurnTableTimer_Timer
  Primitive_TurnTable.objRotZ = Primitive_TurnTable.ObjRotZ - 1
  if Primitive_TurnTable.objRotZ <= 0 Then
    Primitive_TurnTable.objRotZ = 360
  End If
  Primitive_PalmTreePlastics.objRotZ = Primitive_TurnTable.ObjRotZ
  Primitive_TurntableScrews.objRotZ = Primitive_TurnTable.ObjRotZ
    Primitive_TurnTableCover.ObjRotZ = Primitive_TurnTable.ObjRotZ
  Primitive_TopDecals.objRotZ = Primitive_TurnTable.ObjRotZ
  if Primitive_TurnTable.objRotZ = IslandRotation Then
    TurnTableTimer.enabled = False
    TurnTableSoundTimer.enabled = False
    TurntableGate.collidable = False
  end If
  ToggleTurnTablePaths
End Sub

Sub TurnTableSoundTimer_Timer
  playsound SoundFX("motor1",DOFGear),0,1,-0.05,0.1
End Sub

Sub SolHoldLock (enabled)
  if enabled then
    Controller.switch(76)=0
  else
    Controller.switch(76)=1
  end if
end sub

'#########################################################################
' Helper Functions

'Jungle Rescue - reactivate a ball that may become trapped under the turntable
Sub RescueKicker1_Hit
  RescueKicker1.destroyball
  RescueKicker2.createball
  RescueKicker2.kick 180,1
End Sub

Sub RampHelper_Hit
  ActiveBall.VelZ = 0
End Sub

'Gates
Sub GateTimer_Timer
  Primitive_VUKGateWire.RotX = -VUKGate.Currentangle
  Primitive_GateWire.RotX = -Gate4.Currentangle
End Sub

Sub MoveTarget35
  Primitive_TargetParts35.TransZ = 5
  Primitive_Target35.TransZ = 5
  Target35.Timerenabled = False
  Target35.Timerenabled = True
End Sub
Sub Target35_Timer
  Target35.Timerenabled = False
  Primitive_TargetParts35.TransZ = 0
  Primitive_Target35.TransZ = 0
End Sub

Sub MoveTarget36
  Primitive_TargetParts36.TransZ = 5
  Primitive_Target36.TransZ = 5
  Target36.Timerenabled = False
  Target36.Timerenabled = True
End Sub
Sub Target36_Timer
  Target36.Timerenabled = False
  Primitive_TargetParts36.TransZ = 0
  Primitive_Target36.TransZ = 0
End Sub

Sub MoveTarget37
  Primitive_TargetParts37.TransZ = 5
  Primitive_Target37.TransZ = 5
  Target37.Timerenabled = False
  Target37.Timerenabled = True
End Sub
Sub Target37_Timer
  Target37.Timerenabled = False
  Primitive_TargetParts37.TransZ = 0
  Primitive_Target37.TransZ = 0
End Sub

Sub MoveTarget38
  Primitive_TargetParts38.TransZ = 5
  Primitive_Target38.TransZ = 5
  Target38.Timerenabled = False
  Target38.Timerenabled = True
End Sub
Sub Target38_Timer
  Target38.Timerenabled = False
  Primitive_TargetParts38.TransZ = 0
  Primitive_Target38.TransZ = 0
End Sub

Sub MoveTarget46
  Primitive_Target46.TransZ = 5
  Target46.Timerenabled = False
  Target46.Timerenabled = True
End Sub
Sub Target46_Timer
  Target46.Timerenabled = False
  Primitive_Target46.TransZ = 0
End Sub

Sub MoveTarget47
  Primitive_Target47.TransZ = 5
  Target47.Timerenabled = False
  Target47.Timerenabled = True
End Sub
Sub Target47_Timer
  Target47.Timerenabled = False
  Primitive_Target47.TransZ = 0
End Sub

Sub MoveTarget48
  Primitive_Target48.TransZ = 5
  Target48.Timerenabled = False
  Target48.Timerenabled = True
End Sub
Sub Target48_Timer
  Target48.Timerenabled = False
  Primitive_Target48.TransZ = 0
End Sub

Sub MoveTarget51
  Primitive_TargetParts51.TransZ = 5
  Primitive_Target51.TransZ = 5
  Target51.Timerenabled = False
  Target51.Timerenabled = True
End Sub
Sub Target51_Timer
  Target51.Timerenabled = False
  Primitive_TargetParts51.TransZ = 0
  Primitive_Target51.TransZ = 0
End Sub

Sub MoveTarget52
  Primitive_TargetParts52.TransZ = 5
  Primitive_Target52.TransZ = 5
  Target52.Timerenabled = False
  Target52.Timerenabled = True
End Sub
Sub Target52_Timer
  Target52.Timerenabled = False
  Primitive_TargetParts52.TransZ = 0
  Primitive_Target52.TransZ = 0
End Sub

Sub MoveTarget53
  Primitive_TargetParts53.TransZ = 5
  Primitive_Target53.TransZ = 5
  Target53.Timerenabled = False
  Target53.Timerenabled = True
End Sub
Sub Target53_Timer
  Target53.Timerenabled = False
  Primitive_TargetParts53.TransZ = 0
  Primitive_Target53.TransZ = 0
End Sub

Sub MoveTarget54
  Primitive_TargetParts54.TransZ = 5
  Primitive_Target54.TransZ = 5
  Target54.Timerenabled = False
  Target54.Timerenabled = True
End Sub
Sub Target54_Timer
  Target54.Timerenabled = False
  Primitive_TargetParts54.TransZ = 0
  Primitive_Target54.TransZ = 0
End Sub

Sub MoveTarget55
  Primitive_TargetParts55.TransZ = 5
  Primitive_Target55.TransZ = 5
  Target55.Timerenabled = False
  Target55.Timerenabled = True
End Sub
Sub Target55_Timer
  Target55.Timerenabled = False
  Primitive_TargetParts55.TransZ = 0
  Primitive_Target55.TransZ = 0
End Sub

Sub MoveTarget56
  Primitive_TargetParts56.TransZ = 5
  Primitive_Target56.TransZ = 5
  Target56.Timerenabled = False
  Target56.Timerenabled = True
End Sub
Sub Target56_Timer
  Target56.Timerenabled = False
  Primitive_TargetParts56.TransZ = 0
  Primitive_Target56.TransZ = 0
End Sub

Sub MoveTarget57
  Primitive_Target57.TransX = 5
  Target57.Timerenabled = False
  Target57.Timerenabled = True
End Sub
Sub Target57_Timer
  Target57.Timerenabled = False
  Primitive_Target57.TransX = 0
End Sub

'Bumper animation
Dim BumperDir1,BumperDir2,BumperDir3,BumperSpeed
BumperSpeed = 1.75 * Bumper1.Ringspeed
Const BumperLowerLimit = -30

Sub Bumper1_Hit
  vpmTimer.PulseSw 41
  RandomSoundBumperMiddle Bumper1
  MoveBumperCap1
End Sub

Sub Bumper2_Hit
  vpmTimer.PulseSw 42
  RandomSoundBumperTop Bumper2
  MoveBumperCap2
End Sub

Sub Bumper3_Hit
  vpmTimer.PulseSw 43
  RandomSoundBumperBottom Bumper3
  MoveBumperCap3
End Sub

Sub MoveBumperCap1
  BumperDir1 = -BumperSpeed
  Bumper1.Timerenabled = True
End Sub
Sub Bumper1_Timer
  Primitive_bumpercap1.TransY = Primitive_bumpercap1.TransY + BumperDir1
  Primitive_bumpercapSCREWS1.TransY = Primitive_bumpercapSCREWS1.TransY + BumperDir1
  if Primitive_bumpercap1.TransY >= 0 Then
    Bumper1.Timerenabled = False
    Primitive_bumpercap1.TransY = 0
    Primitive_bumpercapSCREWS1.TransY = 0
  end if
  if Primitive_bumpercap1.TransY <= BumperLowerLimit Then
    BumperDir1 = BumperSpeed
  end if
End Sub

Sub MoveBumperCap2
  BumperDir2 = -BumperSpeed
  Bumper2.Timerenabled = True
End Sub
Sub Bumper2_Timer
  Primitive_bumpercap2.TransY = Primitive_bumpercap2.TransY + BumperDir2
  Primitive_bumpercapSCREWS2.TransY = Primitive_bumpercapSCREWS2.TransY + BumperDir2
  if Primitive_bumpercap2.TransY >= 0 Then
    Bumper2.Timerenabled = False
    Primitive_bumpercap2.TransY = 0
    Primitive_bumpercapSCREWS2.TransY = 0
  end if
  if Primitive_bumpercap2.TransY <= BumperLowerLimit Then
    BumperDir2 = BumperSpeed
  end if
End Sub

Sub MoveBumperCap3
  BumperDir3 = -BumperSpeed
  Bumper3.Timerenabled = True
End Sub
Sub Bumper3_Timer
  Primitive_bumpercap3.TransY = Primitive_bumpercap3.TransY + BumperDir3
  Primitive_bumpercapSCREWS3.TransY = Primitive_bumpercapSCREWS3.TransY + BumperDir3
  if Primitive_bumpercap3.TransY >= 0 Then
    Bumper3.Timerenabled = False
    Primitive_bumpercap3.TransY = 0
    Primitive_bumpercapSCREWS3.TransY = 0
  end if
  if Primitive_bumpercap3.TransY <= BumperLowerLimit Then
    BumperDir3 = BumperSpeed
  end if
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  RandomSoundSlingshotLeft Sling2
  vpmTimer.PulseSw 44
  LSling.Visible = 0
  LSling1.Visible = 1
  sling2.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  RandomSoundSlingshotRight Sling1
  vpmTimer.PulseSw 45
  RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Ramp sounds
Sub JungleRampEnter_Hit
  WireRampOn True
End Sub

Sub JungleRampExit_Hit
  WireRampOff
End Sub

Sub JungleRamp2Exit_Hit
  WireRampOff
End Sub

Sub JungleRampWire1Enter_Hit
  WireRampOn False
End Sub

Sub RightRampPlasticExit_Hit
  WireRampOff
End Sub

Sub MiddleWireJunction_Hit
  WireRampOn False
End Sub

Sub KickerWireRampEnter_Hit
  WireRampOn False
End Sub

Sub TopRampPlasticEnter_Hit
  WireRampOn True
End Sub

' Timers
Sub RDampen_Timer
  Cor.Update
End Sub

Sub GameTimer_Timer
  RollingUpdate
End Sub

Sub FrameTimer_Timer()
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate
End Sub

'Kona Texture Swap - called only if necessary
Dim KonaState
Sub UpdateKona
  KonaState = 0
  if Controller.solenoid(12) then Konastate = Konastate + 1 'Spot
  if Controller.solenoid(17) then Konastate = Konastate + 10  'Red Dome
  if Controller.solenoid(22) then Konastate = Konastate + 100 'Head2

  Select Case Konastate
    case 0:     Primitive_Kona.image = "KonaMap"      'all OFF
    case 1:     Primitive_Kona.image = "KonaMap_SPOTL_ON1"  'Spot only
    case 10:  Primitive_Kona.image = "KonaMap_FLSH_ON1" 'Red Dome only
    case 11:  Primitive_Kona.image = "KonaMap_FLSH_ON3" 'Spot+Red Dome
    case 100: Primitive_Kona.image = "KonaMap_FLSH2_ON1"  'Head2 only
    case 101: Primitive_Kona.image = "KonaMap_FLSH2_ON3"  'Spot+Head2
    case 110: Primitive_Kona.image = "KonaMap_FLSHx2_ON2" 'Red Dome+Head2
    case 111: Primitive_Kona.image = "KonaMap_FLSHx2_ON3" 'all ON
  End Select
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)
Dim ccount

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

Dim chgLamp, num, chg, ii, JungleRampChanged
Sub LampTimer_Timer()
  JungleRampChanged = False
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
      FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step

      'change Jungle Ramp only if necessary
      if chgLamp(ii, CHGNO) = 67 or chgLamp(ii, CHGNO) = 68 Then
        JungleRampChanged = True
      End If
    Next
    End If
    UpdateLamps

  'Gilligan texture swap light handler
  if JungleRampChanged Then
    if Controller.Lamp(67) or Controller.Lamp(68) Then
      P_BridgeRamp.image = "mechramp_map-on"
      Primitive_PalmPostP.image = "PalmTreePlastic(sm)MapON"
    Else
      P_BridgeRamp.image = "MechRamp_Map2"
      Primitive_PalmPostP.image = "PalmTreePlastic(sm)Map"
    End If
  End If
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.3    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.15 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
  'Inserts
    Flash 11, Lamp11a
  NFadeLm 11, Lamp11b
  NFadeLm 11,  Lamp11
  NFadeLm 11, Lamp11c
  NFadeLm 11, Lamp11d
  NFadeLm 11, Lamp11e
    Flash 12, Lamp12a
  NFadeLm 12, Lamp12b
  NFadeLm 12,  Lamp12
  NFadeLm 12, Lamp12c
  NFadeLm 12, Lamp12d
  NFadeLm 12, Lamp12e
    Flash 13, Lamp13a
  NFadeLm 13, Lamp13b
  NFadeLm 13,  Lamp13
  NFadeLm 13, Lamp13c
  NFadeLm 13, Lamp13d
    Flash 14, Lamp14a
  NFadeLm 14, Lamp14b
  NFadeLm 14,  Lamp14
  NFadeLm 14, Lamp14c
  NFadeLm 14, Lamp14d
    Flash 15, Lamp15a
  NFadeLm 15, Lamp15b
  NFadeLm 15,  Lamp15
    Flash 16, Lamp16a
  NFadeLm 16,  Lamp16
  NFadeLm 16, Lamp16b
  NFadeLm 16, Lamp16c
  NFadeLm 16, Lamp16d
  NFadeLm 16, Lamp16e
  NFadeLm 16, Lamp16f
  NFadeLm 16, Lamp16g
    Flash 17, Lamp17g
  NFadeLm 17, Lamp17a
  NFadeLm 17, Lamp17b
  NFadeLm 17, Lamp17c
  NFadeLm 17, Lamp17d
  NFadeLm 17, Lamp17e
  NFadeLm 17, Lamp17f
  NFadeLm 17,  Lamp17
    Flash 18, Lamp18a
  NFadeLm 18, Lamp18b
  NFadeLm 18, Lamp18c
  NFadeLm 18,  Lamp18

    Flash 21, Lamp21a
  NFadeLm 21,  Lamp21
    Flash 22, Lamp22a
  NFadeLm 22,  Lamp22
    Flash 23, Lamp23a
  NFadeLm 23,  Lamp23
      Flash 24, Lamp24a
  NFadeLm 24,  Lamp24
      Flash 25, Lamp25a
  NFadeLm 25,  Lamp25
      Flash 26, Lamp26a
  NFadeLm 26,  Lamp26
  NFadeLm 26, Lamp26b
  NFadeLm 26, Lamp26c
      Flash 27, Lamp27a
  NFadeLm 27, Lamp27b
  NFadeLm 27,  Lamp27
      Flash 28, Lamp28a
  NFadeLm 28, Lamp28b
  NFadeLm 28,  Lamp28

      Flash 31, Lamp31a
  NFadeLm 31,  Lamp31
      Flash 32, Lamp32a
  NFadeLm 32,  Lamp32
      Flash 33, Lamp33a
  NFadeLm 33,  Lamp33
      Flash 34, Lamp34a
  NFadeLm 34,  Lamp34
      Flash 35, Lamp35a
  NFadeLm 35,  Lamp35
      Flash 36, Lamp36a
  NFadeLm 36,  Lamp36
    Flash 37, Lamp37a
  NFadeLm 37,  Lamp37
  NFadeLm 37, Lamp37b
    Flash 38, Lamp38a
  NFadeLm 38,  Lamp38

      Flash 41, Lamp41a
  NFadeLm 41,  Lamp41
      Flash 42, Lamp42a
  NFadeLm 42,  Lamp42
      Flash 43, Lamp43a
  NFadeLm 43,  Lamp43
      Flash 44, Lamp44a
  NFadeLm 44,  Lamp44
      Flash 45, Lamp45a
  NFadeLm 45,  Lamp45
      Flash 46, Lamp46a
  NFadeLm 46,  Lamp46
      Flash 47, Lamp47a
  NFadeLm 47,  Lamp47
     Flash 48, Lamp48a
  NFadeLm 48,  Lamp48

      Flash 51, Lamp51a
  NFadeLm 51,  Lamp51
      Flash 52, Lamp52a
  NFadeLm 52,  Lamp52
      Flash 53, Lamp53a
  NFadeLm 53,  Lamp53
      Flash 54, Lamp54a
  NFadeLm 54,  Lamp54
      Flash 55, Lamp55a
  NFadeLm 55,  Lamp55
      Flash 56, Lamp56a
  NFadeLm 56,  Lamp56
      Flash 57, Lamp57a
  NFadeLm 57, Lamp57b
  NFadeLm 57,  Lamp57
      Flash 58, Lamp58a
  NFadeLm 58, Lamp58b
  NFadeLm 58,  Lamp58

      Flash 61, Lamp61a
  NFadeLm 61, Lamp61b
  NFadeLm 61,  Lamp61
  NFadeLm 61, Lamp61c
  NFadeLm 61, Lamp61d
      Flash 62, Lamp62a
  NFadeLm 62, Lamp62b
  NFadeLm 62,  Lamp62
  NFadeLm 62, Lamp62c
  NFadeLm 62, Lamp62d
      Flash 63, Lamp63a
  NFadeLm 63, Lamp63b
  NFadeLm 63,  Lamp63
  NFadeLm 63, Lamp63c
      Flash 64, Lamp64a
  NFadeLm 64,  Lamp64
      Flash 65, Lamp65a
  NFadeLm 65,  Lamp65
      Flash 66, Lamp66a
  NFadeLm 66,  Lamp66
      Flash 67, Lamp67a
  NFadeLm 67,  Lamp67
      Flash 68, Lamp68a
  NFadeLm 68,  Lamp68
    Flash 71, Lamp71a
  NFadeLm 71, Lamp71b
  NFadeLm 71,  Lamp71
      Flash 72, Lamp72a
  NFadeLm 72, Lamp72b
  NFadeLm 72,  Lamp72
      Flash 73, Lamp73a
  NFadeLm 73, Lamp73b
  NFadeLm 73,  Lamp73
      Flash 74, Lamp74a
  NFadeLm 74, Lamp74b
  NFadeLm 74,  Lamp74
      Flash 76, Lamp76a
  NFadeLm 76,  Lamp76
  NFadeLm 76, Lamp76b
      Flash 77, Lamp77a
  NFadeLm 77,  Lamp77
  NFadeLm 77, Lamp77b

' 'Flashers

  NFadeL 112, F112
  NFadeObjm 117, Primitive_KonaFlasher, "dome4_red_lit", "dome4_red"  'call the Object update before the light update because of the fast status change
  NFadeLm 117, F117
  NFadeLm 117, F117a
  NFadeL 117, F117b
  NFadeL 118, F118
  NFadeL 119, F119
  NFadeL 120, F120
  NFadeL 121, F121
  NFadeLm 122, F122
  NfadeLm 122, F122b
  NFadeLm 122, F122c
  NFadeLm 122, F122d
  NFadeLm 122, F122e
  NFadeL 122, F122f
  NfadeL 125, F125
  NfadeL 126, F126
  NfadeL 127, F127
  NfadeL 128, F128
 End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

Sub SetModLamp(nr, value)
  FadingLevel(nr) = value
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

Sub NFadeLmb(nr, object) ' used for multiple lights with blinking
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 2
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
  Select Case FadingLevel(nr)
    Case 2:object.image = d:FadingLevel(nr) = 0 'Off
    Case 3:object.image = c:FadingLevel(nr) = 2 'fading...
    Case 4:object.image = b:FadingLevel(nr) = 3 'fading...
    Case 5:object.image = a:FadingLevel(nr) = 1 'ON
 End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:object.image = d
        Case 3:object.image = c
        Case 4:object.image = b
        Case 5:object.image = d
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


'******************************************************
'       FLIPPER AND RUBBER CORRECTION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    'x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 0  'don't mess with these
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    'x.DebugOn = True : stickL.visible = True : tbpl.visible = True : vpmSolFlipsTEMP.DebugOn = True
    x.TimeDelay = 60
  Next

  'rf.report "Velocity"
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.2,   1.07
  addpt "Velocity", 2, 0.41, 1.05
  addpt "Velocity", 3, 0.44, 1
  addpt "Velocity", 4, 0.65,  1.0'0.982
  addpt "Velocity", 5, 0.702, 0.968
  addpt "Velocity", 6, 0.95,  0.968
  addpt "Velocity", 7, 1.03,  0.945

  'rf.report "Polarity"
  AddPt "Polarity", 0, 0, -4.7
  AddPt "Polarity", 1, 0.16, -4.7
  AddPt "Polarity", 2, 0.33, -4.7
  AddPt "Polarity", 3, 0.37, -4.7
  AddPt "Polarity", 4, 0.41, -4.7
  AddPt "Polarity", 5, 0.45, -4.7
  AddPt "Polarity", 6, 0.576,-4.7
  AddPt "Polarity", 7, 0.66, -2.8
  AddPt "Polarity", 8, 0.743, -1.5
  AddPt "Polarity", 9, 0.81, -1.5
  AddPt "Polarity", 10, 0.88, 0



  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart, FlipperEnd, LR, PartialFlipCoef
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
  Public Property Let EndPoint(aInput) : if IsObject(aInput) then FlipperEnd = aInput.x else FlipperEnd = aInput : end if : End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property

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
        if DebugOn then StickL.visible = True : StickL.x = balldata(x).x    'debug TODO
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 20 Then
      PartialFlipCoef = 0
    End If
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
      dim teststr : teststr = "Cutoff"
      tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
      if tmp < 0.1 then 'if real ball position is behind flipper, exit Sub to prevent stucks  'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
        if DebugOn then TestStr = "real pos < 0.1 ( " & round(tmp,2) & ")" : tbpl.text = Teststr
        'RemoveBall aBall
        'Exit Sub
      end if

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        if DebugOn then teststr = "y velocity: " & round(aBall.vely, 3) & "exit sub" : tbpl.text = teststr
        RemoveBall aBall
        exit Sub
      end if
      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          'TB.TEXT = balldata(x).id & " " & BALLDATA(X).X & VBNEWLINE & FLIPPERSTART & " " & FLIPPEREND
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        if DebugOn then set tmp = new spoofball : tmp.data = aBall : End If
        if IsEmpty(BallData(idx).id) and aBall.VelY < -12 then 'if tip hit with no collected data, do vel correction anyway
          if PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1) > 1.1 then 'adjust plz
            VelCoef = LinearEnvelope(5, VelocityIn, VelocityOut)
            if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
            if Enabled then aBall.Velx = aBall.Velx*VelCoef'VelCoef
            if Enabled then aBall.Vely = aBall.Vely*VelCoef'VelCoef
            if DebugOn then teststr = "tip protection" & vbnewline & "velcoef: " & round(velcoef,3) & vbnewline & round(PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1),3) & vbnewline
            'debug.print teststr
          end if
        Else
     :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
          if Enabled then aBall.Velx = aBall.Velx*VelCoef
          if Enabled then aBall.Vely = aBall.Vely*VelCoef
        end if
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      'debug
      if DebugOn then
        TestStr = teststr & "%pos:" & round(BallPos,2)
        if IsEmpty(PolarityOut(0) ) then
          teststr = teststr & vbnewline & "(Polarity Disabled)" & vbnewline
        else
          teststr = teststr & "+" & round(1 *(AddX*ycoef*PartialFlipcoef),3)
          if BallPos >= PolarityOut(uBound(PolarityOut) ) then teststr = teststr & "(MAX)" & vbnewline else teststr = teststr & vbnewline end if
          if Ycoef < 1 then teststr = teststr &  "ycoef: " & ycoef & vbnewline
          if PartialFlipcoef < 1 then teststr = teststr & "PartialFlipcoef: " & round(PartialFlipcoef,4) & vbnewline
        end if

        teststr = teststr & vbnewline & "Vel: " & round(BallSpeed(tmp),2) & " -> " & round(ballspeed(aBall),2) & vbnewline
        teststr = teststr & "%" & round(ballspeed(aBall) / BallSpeed(tmp),2)
        tbpl.text = TestSTR
      end if
    Else
      'if DebugOn then tbpl.text = "td" & timedelay
    End If
    RemoveBall aBall
  End Sub
End Class

'================================
'Helper Functions


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

Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub


Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

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
RubbersD.addpoint 0, 0, 0.935 '0.96 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.935 '0.96
RubbersD.addpoint 2, 5.76, 0.942 '0.967 'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
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

    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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
'dim cor : set cor = New CoRTracker
'cor.debugOn = False
'cor.update() - put this on a low interval timer
'Class CoRTracker
' public DebugOn 'tbpIn.text
' public ballvel
'
' Private Sub Class_Initialize : redim ballvel(0) : End Sub
' 'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
' Public Sub Update() 'tracks in-ball-velocity
'   dim str, b, AllBalls, highestID : allBalls = getballs
'   if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
'   for each b in allballs
'     if b.id >= HighestID then highestID = b.id
'   Next
'
'   if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
'
'   for each b in allballs
'     ballvel(b.id) = BallSpeed(b)
'     if DebugOn then
'       dim s, bs 'debug spacer, ballspeed
'       bs = round(BallSpeed(b),1)
'       if bs < 10 then s = " " else s = "" end if
'       str = str & b.id & ": " & s & bs & vbnewline
'       'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
'     end if
'   Next
'   if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
' End Sub
'End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
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

Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'Clamp if on the boundry lines
  'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
  'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

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
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1       https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2       https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3       https://youtu.be/ESXWGJZY_EI


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
TargetSoundFactor = 0.0025 * 10  'volume multiplier; must not be zero
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
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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

' Thalamus, AudioFade patched
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 5) 'was 10
	Else
		AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
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

' Thalamus, AudioPan patched
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 5) 'was 10
	Else
		AudioPan = CSng( - (( - tmp) ^ 5) ) 'was 10
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
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
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

  FlipperCradleCollision ball1, ball2, velocity

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

'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
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


'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
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
  Dim gBOT
  gBOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    ' Comment the next line If you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub If no balls on the table
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

    ' Check for a stuck ball just below the Seltzer hole.  Very rare.
    If InRect(gBOT(b).x, gBOT(b).y, 618,820,647,810,650,840,645,847) and ABS(gBOT(b).velx) < 1 and ABS(gBOT(b).vely) < 1 Then
            gBOT(b).velx  = -2
            gBOT(b).vely  = -2
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
    ' Comment the next If block, If you are not implementing the Dynamic Ball Shadows
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

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order - Used above
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
'        * Otherwise, the ball will auto remove If it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  If tnob = 5, then RampBalls(6,2)
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
  ' To see If the the ball was already added to the array.
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
      If BallVel(RampBalls(x,0) ) > 1 Then ' If ball is moving, play rolling sound
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

'******************************************************
'**** END RAMP ROLLING SFX
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

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
' Example timer sub:

'Sub FrameTimer_Timer()
' If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
'End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
'Const lob = 0  'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top for easy access by players
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 '2 = flasher image shadow, but it moves like ninuzzu's

' *** The following segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' ** Change gBOT to BOT if using existing getballs code
' ** Double commented lines commonly found there included for reference:

''  ' stop the sound of deleted balls
''  For b = UBound(gBOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
''    ...rolling(b) = False
''    ...StopSound("BallRoll_" & b)
''  Next
''
'' ...rolling and drop sounds...
''
''    If DropCount(b) < 5 Then
''      DropCount(b) = DropCount(b) + 1
''    End If
''
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     BallShadowA(b).visible = 1
'     BallShadowA(b).X = gBOT(b).X + offsetX
'     If gBOT(b).Z > 30 Then
'       BallShadowA(b).height=gBOT(b).z - BallSize/4 + b/1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
'       BallShadowA(b).Y = gBOT(b).Y + offsetY + BallSize/10
'     Else
'       BallShadowA(b).height=gBOT(b).z - BallSize/2 + 1.04 + b/1000
'       BallShadowA(b).Y = gBOT(b).Y + offsetY
'     End If
'   End If

' *** Place this inside the table init, just after trough balls are added to gBOT
'
' Add balls to shadow dictionary
' For Each xx in gBOT
'   bsDict.Add xx.ID, bsNone
' Next

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


' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

'Dim PI: PI = 4*Atn(1)

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

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function Radians(Degrees)
  Radians = Degrees * PI / 180
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
Const gilvl = 1

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objrtx1(5), objrtx2(5)
Dim objBallShadow(5)
Dim OnPF(5)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4)
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
  Dim gBOT: gBOT = getballs

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
  Dim gBOT: gBOT=getballs

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

' VR Plunger Animation
Sub TimerPlunger_Timer
  If Primary_plunger.Y < 1205 Then
    Primary_plunger.Y = Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  Primary_plunger.Y = 1070 + (5* Plunger.Position) -20
End Sub


'****************************************************************
'****  LUT SELECTOR
'****************************************************************

'//////////////---- LUT (Colour Look Up Table) ----//////////////
'0 = Fleep Natural Dark 1
'1 = Fleep Natural Dark 2
'2 = Fleep Warm Dark
'3 = Fleep Warm Bright
'4 = Fleep Warm Vivid Soft
'5 = Fleep Warm Vivid Hard
'6 = Skitso Natural and Balanced
'7 = Skitso Natural High Contrast
'8 = 3rdaxis Referenced THX Standard
'9 = CalleV Punchy Brightness and Contrast
'10 = HauntFreaks Desaturated
'11 = Tomate Washed Out
'12 = VPW Original (Default)
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book
'16 = Tyson171's Skitso Mod2

Dim LUTset, DisableLUTSelector, bLutActive
DisableLUTSelector = 0  ' Disables the ability to change LUT option with magna saves in game when set to 1

LoadLUT
SetLUT

Sub SetLUT  'AXS
  Gilligan.ColorGradeImage = "LUT" & LUTset
End Sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = False
  LUTBox.Visible = 0
  VRLutdesc.Visible = 0
End Sub

' LUT Selector Timer
Sub LutSlctr_timer
  LutSlctr.Enabled = False
End Sub

Sub ShowLUT

  LUTBox.visible = 1
  VRLutdesc.visible = 1

  Select Case LUTSet
        Case 0: LUTBox.text = "Fleep Natural Dark 1": VRLUTdesc.imageA = "LUTcase0"
    Case 1: LUTBox.text = "Fleep Natural Dark 2": VRLUTdesc.imageA = "LUTcase1"
    Case 2: LUTBox.text = "Fleep Warm Dark": VRLUTdesc.imageA = "LUTcase2"
    Case 3: LUTBox.text = "Fleep Warm Bright": VRLUTdesc.imageA = "LUTcase3"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft": VRLUTdesc.imageA = "LUTcase4"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard": VRLUTdesc.imageA = "LUTcase5"
    Case 6: LUTBox.text = "Skitso Natural and Balanced": VRLUTdesc.imageA = "LUTcase6"
    Case 7: LUTBox.text = "Skitso Natural High Contrast": VRLUTdesc.imageA = "LUTcase7"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard": VRLUTdesc.imageA = "LUTcase8"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast": VRLUTdesc.imageA = "LUTcase9"
    Case 10: LUTBox.text = "HauntFreaks Desaturated" : VRLUTdesc.imageA = "LUTcase10"
      Case 11: LUTBox.text = "Tomate washed out": VRLUTdesc.imageA = "LUTcase11"
    Case 12: LUTBox.text = "VPW Default": VRLUTdesc.imageA = "LUTcase12"
    Case 13: LUTBox.text = "bassgeige": VRLUTdesc.imageA = "LUTcase13"
    Case 14: LUTBox.text = "blacklight": VRLUTdesc.imageA = "LUTcase14"
    Case 15: LUTBox.text = "B&W Comic Book": VRLUTdesc.imageA = "LUTcase15"
    Case 16: LUTBox.text = "Tyson171's Skitso Mod2": VRLUTdesc.imageA = "LUTcase16"
  End Select

  LUTBox.TimerEnabled = 1

End Sub

Sub SaveLUT
  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End If

  If LUTset = "" Then LUTset = 12 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "GilligansLUT.txt",True)
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub

Sub LoadLUT
    bLutActive = False
  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    LUTset = 12
    Exit Sub
  End If
  If Not FileObj.FileExists(UserDirectory & "GilligansLUT.txt") Then
    LUTset = 12
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "GilligansLUT.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
  If (TextStr.AtEndOfStream=True) Then
    Exit Sub
  End If
  rLine = TextStr.ReadLine
  If rLine = "" Then
    LUTset = 12
    Exit Sub
  End If
  LUTset = int (rLine)
  Set ScoreFile = Nothing
  Set FileObj = Nothing
End Sub

'****************************************************************
'****  END LUT SELECTOR
'****************************************************************


'**********************************************************
'*******Set Up VR Backglass and Backglass Flashers  *******
'**********************************************************

Sub SetBackglass()
  BGDark.visible = True
  Dim VRobj
  For Each VRobj In VRBackglass
    VRobj.x = VRobj.x
    VRobj.height = - VRobj.y + 85
    VRobj.y = 84 'adjusts the distance from the backglass towards the user
    VRobj.rotx = -88.5
  Next
End Sub

Option Explicit
Randomize

' Thalamus
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Improved directional positions
' Changed UseSolenoids=1 to 2 - and reverted
'
' Austin Powers (Stern 2001) VPin Workshop mod
' - nFozzy Physics: Benji, iaakki, Rothbauerw, Javier
' - PF rework, graphics: iaakki, EBisLit, Tomate, Brad1x
' - Inserts and lights: iaakki
' - VR work & Various Fixes: Sixtoe
' - Ramp plastics: Benji
' - Fleep's Sound package integration: Benji
' - 3D model changes: Benji, Tomate, Sixtoe
' - Various script changes by iaakki, Javier
' - Testing: people of VPin Workshop Discord, Rik Laubach

' CHANGES
' 025 - tomate - New Left Ramp Primitive, New Left Ramp LowPoly colidable object, New primitive and material for the time machine axis
' 026 - iaakki - Live catch bounce feature
' 027 - tomate - Left Ramp decals, domes, gates and plate aligned
' 028 - tomate - Left ramp mesh replacement
' 029 - benji - new right and left ramp textures, new VR right and left ramp textures
' 030 - tomate - new left ramp with fewer polygons
' 031 - Benji - new right and left ramp textures, new VR right and left ramp textures. Boosted blenddisablelighting of ramps in script to compensate
' 032 - sixtoe - Added blenddisablelighting variation for VR, manually tweaked VR ramp textures, moved ramp poles and screws to line up with ramps over slings, turned up backglass in VR, tweaaked vr plunger position
' 033 - iaakki - SSLCrystal, Diverter and plastic ramps DL lit values tuned. Backwall flasher color taken into colormod. Livecatch script updated. Target hit sounds fixed. Rolling sound volume option added.
' 034 - iaakki - fading script changed, insert flashers added for those near primitive, plastic depth bias issue
' 035 - iaakki/sixtoe - Plastics UV map fixed, lane guide insert flashers adjusted, plastic image colors changed
' 036 - iaakki - plastics colors reverted back a bit. Wall44 small tune, Dr Evil hit sound added.
' 037 - iaakki - left ramp parameters fixed, PowerFlips option added that sets alternative physics values to flips.
' 038 - iaakki - Ramp fx fix from Javier, Ball drop fx from Javier, Flip friction setup was missing in options.
' 039 - iaakki - Ball drop sounds changed, some inserts tuned, GI top area tuned, laneguides adjusted, virtucon text fixed
' 040 - iaakki - Flasherbloom31 added for bumpers
' 041 - iaakki - DrEvil bug workaround, better standup targets around TM ramp, Light54 fixed, other minor improvements
' 042 - iaakki - Flipper sizes fixed
' 043 - iaakki - Latest NF flip script, slackyflips script option added, some material tunings, testing, slingshot tweak
' 044 - benji/tomate - Mr. Bigglesworth added
' 047 - iaakki - various audio fixes, flipper tricks and slaggy flips fixed, Mr Bigglesworth tied to GI
' 048 - sixtoe - replaced pincab_bottom wall for primitive
' 049 - iaakki - fixed desktop pov a bit, adjusted slackyflips, minor fixes and some cleanup
' 052 - iaakki - Ramps fixed one more time, FlipperSolNudge option added, VUK altered, walls added near some plastics, etc...
' RC1 - iaakki - recontoured Flip Trigger areas to prevent ball stuck on trigger primitive near the flip, minor change to left orb, flipsolnudge divider 1.5 -> 1.7
' RC2 - iaakki - FlipSolfNudge renamed to FlipperNudge and multiball fix, Reduced friction on left ramp just a bit, parking cannon properly and fixed shooting
' RC3 - iaakki - Fixed reflection issue on DrEvil and removed some debug prints
' RC4 - iaakki - Flip size fine tune and FlipTrigger primitives remodeled. Assets locked
' 1.0.1 - Sixtoe - Adjusted right ramp, fixed some minor issues.
' 1.0.2 - Sixtoe - Redid all the ramp gates including prims and switches so everything both works and looks better, fixed 2 flasher switches not working.
' 1.0.2 - iaakki - added one hidden Wall003 and minor flip tweak
' 1.0.3 - iaakki - VUK fixed

'**********************************************************************************************************
'    Table options
'**********************************************************************************************************

'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.8

'Flower Side panels
FlowerSides = 1

'Cabinet mode - Will hide the rails and scale the side panels higher
CabinetMode = 0

'VR Room
Const VRRoom = 0 'Turns on VR Room and Cabinet.   VPVR must also be set for desktop play.

'Live catch window size. Higher value will make live catching easier. 8-32
Const LiveCatch = 16
'Add minor slack to flips on high collisions
Const slackyFlips = 1
'Make flipper solenoid to alter active ball behavior
Const FlipperNudge = 1
' Default or power flips,
Const PowerFlips = 1

Const BallRollVol = 0.3 'Ball Rolling Volume multiplier, 0.7 default

'GI COLOR MOD
'   White Bulbs = 0
'   Warm Bulbs = 1
'Change the value below to set option
GIColorMod = 0




On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0



Const Ballsize = 50
Const BallMass = 1.7
Const FlasherIntensity = 20  ' only for flasher27, 100 = Default, intense!   Try 20 or even less for more subdued effects.

' Thalamus - ffv2 should not be used for this table according to nFozzy
Const cGameName="austin",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="fx_SolOn",SSolenoidOff="", SCoin="fx_coin"

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Pincab_Rails.visible=1
Else
Pincab_Rails.visible=0
End if

dim UseVPMDMD:UseVPMDMD = DesktopMode

If VRRoom = 1 and DesktopMode Then 'Turns on VR Room and Cabinet
  Dim VRThings
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  ScoreText.visible = 0
  LeftRampPrimitive.image = "AP_Ramps_VR"
  RightRampPrimitive.image = "AP_Ramps_VR"
Else
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  ScoreText.visible = 1
  LeftRampPrimitive.image = "AP_Ramps"
  RightRampPrimitive.image = "AP_Ramps"
End If

LoadVPM "01200000", "SEGA.VBS", 3.10

'**********************************************************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallBack(1)   = "SolTrough"
SolCallBack(2)   = "SolAutoPlungerIM"
SolCallback(3)   = "bsScoop.SolOut"
SolCallback(4)   = "SolLaserBeam"                     'Laser Beam Eject
SolCallback(7)   = "SolAustinDance"                     'Dancing Austin
SolCallback(8)   = "bsInodoro.SolOut"                     'Toilet Post
SolCallback(19) = "mrEvil"                              ' Motor Solenoid
SolCallback(21) = "Solcannon"                             ' cannon Solenoid
SolCallback(20) = "SolTimeMachine"                      'Time Machine Motor Relay Board
SolCallback(22) = "SolDiv"                          'Laser Beam Diverter
SolCallback(23) = "vpmSolGate Gate, SoundFX(SSolenoidOn,DOFContactors),"
SolCallBack(24) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"         'Knocker
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


SolCallBack(5)   = "SetLamp 105,"                 'Time Travel Ramp
SolCallBack(6)   = "FlashGreen"                 'Right Green Dome
SolCallBack(12)  = "FlashBlue"                'Left Blue Dome
SolCallBack(13)  = "FlashRed"                 'Right Red Dome
SolCallBack(25)  = "SetLamp 125,"                 'PF light insert
SolCallBack(26)  = "SetLamp 126,"                 'PF light insert
SolCallBack(27)  = "SolFlasher27"                 'Spot Light
SolCallBack(28)  = "SetLamp 128,"                 'PF light insert
SolCallBack(29)  = "SetLamp 129,"                 'PF light insert
SolCallBack(30)  = "SetLamp 130,"                 'PF light insert
SolCallBack(31)  = "SetLamp 131,"                 'PF light insert
SolCallBack(32)  = "FlashYellow"                'Left yellow Dome



'************************************************************************
'            INIT TABLE
'************************************************************************

Dim bsTrough, bsInodoro, bsScoop, mMagnet, mLaserGun, mEvil, plungerIM

Sub Table1_Init
  TableOptions()
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Austin Powers -- Stern, 2001"&chr(13)&"Yes"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval:
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=56
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    ' Trough & Ball Release
    Set bsTrough = New cvpmtrough
  With bsTrough
    .Size = 4
    .InitSwitches Array (14, 13, 12, 11)
        .InitExit BallRelease, 90, 10
    .InitEntrySounds "fx_drain", SoundFX(SSolenoidOn,DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
    .InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_ballrel",DOFContactors)
    .Balls = 4
    .CreateEvents "bsTrough", Drain
  End With

    ' Impulse Plunger
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, 50, 0.6
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("fx_AutoPlunger",DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
        .CreateEvents "plungerIM"
    End With

    Set bsInodoro = new cvpmSaucer
  With bsInodoro
        .InitKicker Sw21, 21, 240, 8, 0
    .InitExitVariance 5, 3
    .InitSounds "popper_ball", SoundFX(SSolenoidOn,DOFContactors), SoundFX("popper",DOFContactors)
        .CreateEvents "bsInodoro", sw21
  End With

  Set bsScoop = New cvpmSaucer
  With bsScoop
      .InitKicker Sw46b, 46, 0, 120, 1.56
    .InitSounds "popper_ball", SoundFX(SSolenoidOn,DOFContactors), SoundFX("popper",DOFContactors)
    .CreateEvents "bsScoop", sw46b
    End With

  Set mMagnet= New cvpmMagnet
  With mMagnet
    .InitMagnet Magnet, 100
    .GrabCenter = 1
    .solenoid= 14
    .CreateEvents "mMagnet"
  End With

     ' LaserGun
  Set mLaserGun = new cvpmMyMech
  With mLaserGun
         .MType = vpmMechOneSol + vpmMechReverse + vpmMechLinear + vpmMechFast
         .Sol1 = 21
         .Length = 750
         .Steps = 750
         .AddSw 43, 0, 10
         .AddSw 44, 400, 750
         .Callback = GetRef("UpdateLserGun")
         .Start
  End With

    'DREvil "By JPSalas"
    Set mEvil = New cvpmMyMech
  With mEvil
        .MType=vpmMechOneSol+vpmMechReverse+vpmMechLinear + vpmMechFast
        .Sol1=19
        .Length=1300
        .Steps=500
        .AddSw 23,0,1
        .AddSw 22,498,500
        .Callback=GetRef("UpdateEvil")
        .Start
  End With

    ' Misc. Initialisation
  RampDiverter.IsDropped=0:RampDiverter2.IsDropped=1
  sw24.Isdropped = True
  Dim x:x = controller.getmech(3)
  UpdateLserGun x, 0, x
  Primitive13.visible=DesktopMode

End Sub

'****************************************
' Real Time updates
'****************************************

Set MotorCallBack = GetRef("GameTimer")

Sub GameTimer
  RollingUpdate
  BallShadowUpdate
  LFLogo.RotY = LeftFlipper.CurrentAngle
  RFlogo.RotY = RightFlipper.CurrentAngle
  AustinDancerP.RotY = -(Spinner1.currentangle)
End Sub


'******************
'   TABLE OPTIONS
'*****************
Dim FlowerSides, CabinetMode, GIColorMod
Dim xxGIColor

Sub TableOptions()
  if FlowerSides = 1 then
    sidePanels.image = "sideblades_flowers"
  Else
    sidePanels.image = "sideblades_wood"
  End if
  if CabinetMode = 1 then
    sidePanels.Size_Y = 2000
  Else
    sidePanels.Size_Y = 1000
  end If

  If GIColorMod = 0 then
    for each xxGIColor in GiLights
      xxGIColor.Color=White
      xxGIColor.ColorFull=WhiteFull
      xxGIColor.Intensity = xxGIColor.Intensity * 0.5
    next
    For each xxGIColor in GILightsBackwalls:xxGIColor.Color=rgb(220,230,255): Next  'BackWall Flashers
  End If

  If GIColorMod = 1 then
    for each xxGIColor in GiLights
      xxGIColor.Color=Warm
      xxGIColor.ColorFull=WarmFull
      'xxGIColor.Intensity = WarmI
    next
    For each xxGIColor in GILightsBackwalls:xxGIColor.Color=rgb(255,255,128): Next  'BackWall Flashers
  End If

  if PowerFlips = 1 Then
    RightFlipper.strength=3900
    RightFlipper.elasticity=0.940000
    RightFlipper.scatter=0
    RightFlipper.eosTorque=0
    RightFlipper.eosTorqueAngle=6
    RightFlipper.return=0.07
    RightFlipper.friction=7
    RightFlipper.elasticityFalloff=0.35
    RightFlipper.RampUp=0

    LeftFlipper.strength=3900
    LeftFlipper.elasticity=0.940000
    LeftFlipper.scatter=0
    LeftFlipper.eosTorque=0.5
    LeftFlipper.eosTorqueAngle=6
    LeftFlipper.return=0.07
    LeftFlipper.friction=7
    LeftFlipper.elasticityFalloff=0.35
    LeftFlipper.RampUp=0
  end if

End Sub

'**********************************************************************************************************
'**********************************************************************************************************
'***RGB OPTIONS COLOR ADJUST***
'**********************************************************************************************************
'**********************************************************************************************************

Dim White, WhiteFull, WhiteI, Warm, WarmFull, WarmI
WhiteFull = rgb(250,250,245)
White = rgb(250,250,210)
WhiteI = .15
Warm = rgb(255,128,0)
WarmFull = rgb(255,252,223)
WarmI = .25


'******************
'Keys Up and Down
'*****************
Sub Table1_KeyUp(ByVal Keycode)
  If Keycode=LeftMagnaSave Or KeyCode=RightMagnaSave or KeyCode = LockBarKey Then Controller.Switch(53)=0

  If KeyCode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
  End If


  'nFozzy Begin'
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
  'nFozzy End'

    If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub Table1_KeyDown(ByVal Keycode)
  Dim BOT, b
  If Keycode=LeftMagnaSave Or KeyCode=RightMagnaSave or KeyCode = LockBarKey Then Controller.Switch(53)=1
  If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If KeyCode = PlungerKey Then Plunger.Pullback:SoundPlungerPull()
  if keycode=StartGameKey then soundStartButton()


'nFozzy Begin'
  If keycode = LeftFlipperKey Then
        LFPress = 1
        If RightFlipper.currentangle = RFEndAngle and FlipperNudge = 1 Then
            BOT = GetBalls

            For b = 0 to Ubound(BOT)
                If FlipperTrigger(BOT(b).x, BOT(b).y, RightFlipper) Then
                    BOT(b).velx = BOT(b).velx / 1.7
                    BOT(b).vely = BOT(b).vely - 1
                end If
            Next
        End If
    end If
    If keycode = RightFlipperKey Then
        rfpress = 1
        If LeftFlipper.currentangle = LFEndAngle and FlipperNudge = 1 Then
            BOT = GetBalls

            For b = 0 to Ubound(BOT)
                If FlipperTrigger(BOT(b).x, BOT(b).y, LeftFlipper) Then
                    BOT(b).velx = BOT(b).velx / 1.7
                    BOT(b).vely = BOT(b).vely - 1
                end If
            Next
        end If
    end If
'nFozzy End'

    If vpmKeyDown(keycode) Then Exit Sub
End Sub

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

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    Atn2 = Sgn(dy) * (PI - Atn(Abs(dy / dx)))
  ElseIf dy = 0 Then
    Atn2 = 0
  Else
    Atn2 = Sgn(dy) * PI / 2
  End If
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


Sub TimerPlunger_Timer()
 'debug.print plunger.position
  VR_Primary_plunger.Y = -50 + (5* Plunger.Position) -20
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_UnPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'************************************************************************
'Solenoid Subs
'************************************************************************
'Trough
Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        If bsTrough.Balls Then vpmTimer.PulseSw 15
    End If
End Sub

'Autoplunger
Sub SolAutoPlungerIM(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

'Diverter
Sub SolDiv(Enabled)
  If enabled Then
    RampDiverter.Isdropped = 1
        PlaySoundAt SoundFX("DiverterOn",DOFContactors), Primitive_RampDiverter
    RampDiverter2.Isdropped = 0
        Primitive_RampDiverter.RotY = 25
  Else
    RampDiverter.Isdropped = 0
        PlaySoundAt SoundFX("DiverterOff",DOFContactors), Primitive_RampDiverter
    RampDiverter2.Isdropped = 1
        Primitive_RampDiverter.RotY = 0
  End If
End Sub

'Maquina del Tiempo
Sub SolTimeMachine(Enabled)
  If Enabled Then
       TimeMachineTimer.Enabled=1
     PlaySoundAtVolLoops "timemachine", TimeMachineP, 0.2, 0
  Else
       TimeMachineTimer.Enabled=0
     StopSound"timemachine"
  End If
End Sub

Sub TimeMachineTimer_Timer()
    TimeMachineP.RotX  = TimeMachineP.RotX + 15
    Me.Enabled = 1
End Sub

Sub mrEvil(Enabled)
  if Enabled then
    'Debug.print "mr evil running"
    PlaySoundAtVol "TOM_Trunk_Motor_Long", sw24p, 0.3
  Else
    'Debug.print "mr evil stopped"
    StopSound"TOM_Trunk_Motor_Long"
    PlaySoundAtVol "TOM_Trunk_Motor_Stop", sw24p, 1
    if sw24.Isdropped then
      'debug.print Sw24P.TransY & " is current. Reset to 0" & sw24.Isdropped
      Sw24P.TransY = 0
    Else
      'debug.print Sw24P.TransY & " is current. not dropped " & sw24.Isdropped
    end If
  End if
End Sub

dim CannonStatus : CannonStatus = false
Sub Solcannon(Enabled)
  if Enabled then
    'Debug.print "cannon running"
    PlaySoundAtVolLoops "CannonMotor", sw24p, 0.3, 0
  Else
    if not CannonStatus Then
      PCannon.RotY = 0
      PCannonshadow.RotY  = 0
      'Debug.print "cannon parked"
    end If
    'Debug.print "cannon stopped"
    StopSound"CannonMotor"
    'PlaySoundAtVol "TOM_Trunk_Motor_Stop", sw24p, 1
  End if
End Sub




' ### Spot light ###
Sub SolFlasher27(enabled)
  SetLamp 127, Enabled

  If enabled Then
'    Flasher27.opacity = 200
'    Flasher27m.opacity = 100
'    FlasherLight27.state = 1
     F127a.image = "bulbOn"
     'F127a.BlendDisableLighting = 20
     'Primitive196.BlendDisableLighting = 0.2
  Else
'    Flasher27.opacity = 0
'    Flasher27m.opacity = 0
'    FlasherLight27.state = 0
     F127a.image = "bulb"
     'F127a.BlendDisableLighting = 0
     'Primitive196.BlendDisableLighting = 0
  End If
End Sub


' ****************************************************
' FLIPPERS
' ****************************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
            RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
    LF.Fire
    Else
    RandomSoundFlipperDownLeft LeftFlipper
    LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    If rightflipper.currentangle < rightflipper.endangle + ReflipAngle Then
            RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
    RF.Fire
    Else
    RandomSoundFlipperDownRight RightFlipper
    RightFlipper.RotateToStart
    End If
End Sub




'*****************
' Switches
'*****************

' Slingshots
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 2.55
        Case 2: zMultiplier = 1.2
        Case 3: zMultiplier = 0.9
        Case 4: zMultiplier = 0.7
    End Select
  'debug.print "zmult: " & zMultiplier & " velz: " & activeball.velz
    activeball.velz = activeball.velz*zMultiplier

    'PlaySoundAt SoundFX("LeftSlingShot", DOFContactors), PegPlasticT4
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 59
    LeftSlingShot.TimerEnabled = 1
  RandomSoundSlingshotLeft PegPlasticT4
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
  Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 2.55
        Case 2: zMultiplier = 1.2
        Case 3: zMultiplier = 0.9
        Case 4: zMultiplier = 0.7
    End Select
  'debug.print "zmult: " & zMultiplier & " velz: " & activeball.velz
    activeball.velz = activeball.velz*zMultiplier
    'PlaySoundAt SoundFX("RightSlingShot", DOFContactors), PegPlasticT18
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 62
    RightSlingShot.TimerEnabled = 1
  RandomSoundSlingshotRight PegPlasticT18
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49 : RandomSoundBumperMiddle Bumper1 : End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50 : RandomSoundBumperBottom Bumper2 : End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51 : RandomSoundBumperTop Bumper3 : End Sub

'Stand Up Targets
' Left Banks Targets
Sub Sw25_Hit:vpmTimer.PulseSw 25 :MoveTarget25 :PlayTargetSound():End Sub
Sub Sw26_Hit:vpmTimer.PulseSw 26 :MoveTarget26 :PlayTargetSound():End Sub
Sub Sw27_Hit:vpmTimer.PulseSw 27 :MoveTarget27 :PlayTargetSound():End Sub
Sub Sw28_Hit:vpmTimer.PulseSw 28 :MoveTarget28 :PlayTargetSound():End Sub ' there was movetarget27 !!!
Sub Sw29_Hit:vpmTimer.PulseSw 29 :MoveTarget29 :PlayTargetSound():End Sub
Sub Sw30_Hit:vpmTimer.PulseSw 30 :MoveTarget30 :PlayTargetSound():End Sub

Sub MoveTarget25
  Sw25a.TransZ = 5
  Sw25b.TransZ = 5
  Sw25.Timerenabled = False
  Sw25.Timerenabled = True
End Sub
Sub Sw25_Timer
  Sw25.Timerenabled = False
  Sw25a.TransZ = 0
  Sw25b.TransZ = 0
End Sub

Sub MoveTarget26
  Sw26a.TransZ = 5
  Sw26b.TransZ = 5
  Sw26.Timerenabled = False
  Sw26.Timerenabled = True
End Sub
Sub Sw26_Timer
  Sw26.Timerenabled = False
  Sw26a.TransZ = 0
  Sw26b.TransZ = 0
End Sub

Sub MoveTarget27
  Sw27a.TransZ = 5
  Sw27b.TransZ = 5
  Sw27.Timerenabled = False
  Sw27.Timerenabled = True
End Sub
Sub Sw27_Timer
  Sw27.Timerenabled = False
  Sw27a.TransZ = 0
  Sw27b.TransZ = 0
End Sub

Sub MoveTarget28
  Sw28a.TransZ = 5
  Sw28b.TransZ = 5
  Sw28.Timerenabled = False
  Sw28.Timerenabled = True
End Sub
Sub Sw28_Timer
  Sw28.Timerenabled = False
  Sw28a.TransZ = 0
  Sw28b.TransZ = 0
End Sub


Sub MoveTarget29
  Sw29a.TransZ = 5
  Sw29b.TransZ = 5
  Sw29.Timerenabled = False
  Sw29.Timerenabled = True
End Sub
Sub Sw29_Timer
  Sw29.Timerenabled = False
  Sw29a.TransZ = 0
  Sw29b.TransZ = 0
End Sub

Sub MoveTarget30
  Sw30a.TransZ = 5
  Sw30b.TransZ = 5
  Sw30.Timerenabled = False
  Sw30.Timerenabled = True
End Sub
Sub Sw30_Timer
  Sw30.Timerenabled = False
  Sw30a.TransZ = 0
  Sw30b.TransZ = 0
End Sub


'Sub sw31_Hit:vpmTimer.PulseSw 31:End Sub
'Sub sw32_Hit:vpmTimer.PulseSw 32:End Sub

'Const VolTarg   = 1    ' Targets multiplier
Dim zMultiplier

DIM target31step
Sub sw31_Hit

    Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 1.55
        Case 2: zMultiplier = 1.2
        Case 3: zMultiplier = .9
        Case 4: zMultiplier = .7
    End Select
  'msgbox "hit 31 zmult: " & zMultiplier
    activeball.velz = activeball.velz*zMultiplier:vpmTimer.PulseSw 31:pSW31.TransY = -3:Target31Step = 1:sw31a.Enabled = 1:PlayTargetSound()
End Sub

Sub sw31a_timer()
  Select Case Target31Step

    Case 1:pSW31.TransY = -1
    Case 2:pSW31.TransY = 2
        Case 3:pSW31.TransY = 0:Me.Enabled = 0
     End Select
  Target31Step = Target31Step + 1
End Sub

DIM target32step
Sub sw32_Hit
    Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 1.55
        Case 2: zMultiplier = 1.2
        Case 3: zMultiplier = .9
        Case 4: zMultiplier = .7
    End Select
  'msgbox "hit 32 zmult: " & zMultiplier & " "
    activeball.velz = activeball.velz*zMultiplier:vpmTimer.PulseSw 32:pSW32.TransY = -3:Target32Step = 1:sw32a.Enabled = 1:PlayTargetSound()
End Sub

Sub sw32a_timer()
  Select Case Target32Step

    Case 1:pSW32.TransY = -1
    Case 2:pSW32.TransY = 2
        Case 3:pSW32.TransY = 0:Me.Enabled = 0
     End Select
  Target32Step = Target32Step + 1
End Sub



' 'Wire Triggers
' Sub sw16_Hit:Controller.Switch(16)=1 : playsoundAt "rollover", ActiveBall : End Sub
' Sub sw16_unHit:Controller.Switch(16)=0:End Sub
' Sub sw33_Hit:Controller.Switch(33)=1 : playsoundAt "rollover", ActiveBall : End Sub
' Sub sw33_unHit:Controller.Switch(33)=0:End Sub
' Sub sw39_Hit:Controller.Switch(39)=1 : playsoundAt "rollover", ActiveBall : End Sub
' Sub sw39_unHit:Controller.Switch(39)=0:End Sub
' Sub sw47_Hit:Controller.Switch(47)=1 : playsoundAt "rollover", ActiveBall : End Sub
' Sub sw47_unHit:Controller.Switch(47)=0:End Sub
' Sub sw48_Hit:Controller.Switch(48)=1 : playsoundAt "rollover", ActiveBall : End Sub
' Sub sw48_unHit:Controller.Switch(48)=0:End Sub
' Sub sw57_Hit:Controller.Switch(57)=1 : playsoundAt "rollover", ActiveBall : End Sub
' Sub sw57_unHit:Controller.Switch(57)=0:End Sub
' Sub sw58_Hit:Controller.Switch(58)=1 : playsoundAt "rollover", ActiveBall : End Sub
' Sub sw58_unHit:Controller.Switch(58)=0:End Sub
' Sub sw60_Hit:Controller.Switch(60)=1 : playsoundAt "rollover", ActiveBall : End Sub
' Sub sw60_unHit:Controller.Switch(60)=0:End Sub
' Sub sw61_Hit:Controller.Switch(61)=1 : playsoundAt "rollover", ActiveBall : End Sub
' Sub sw61_unHit:Controller.Switch(61)=0:End Sub


'Wire Triggers
Sub sw16_Hit:Controller.Switch(16)=1 : RandomSoundRollover() : End Sub
Sub sw16_unHit:Controller.Switch(16)=0:End Sub
Sub sw33_Hit:Controller.Switch(33)=1 : RandomSoundRollover() : End Sub
Sub sw33_unHit:Controller.Switch(33)=0:End Sub
Sub sw39_Hit:Controller.Switch(39)=1 : RandomSoundRollover() : End Sub
Sub sw39_unHit:Controller.Switch(39)=0:End Sub
Sub sw47_Hit:Controller.Switch(47)=1 : RandomSoundRollover() : End Sub
Sub sw47_unHit:Controller.Switch(47)=0:End Sub
Sub sw48_Hit:Controller.Switch(48)=1 : RandomSoundRollover() : End Sub
Sub sw48_unHit:Controller.Switch(48)=0:End Sub
Sub sw57_Hit:Controller.Switch(57)=1 : RandomSoundRollover() : End Sub
Sub sw57_unHit:Controller.Switch(57)=0:End Sub
Sub sw58_Hit:Controller.Switch(58)=1 : RandomSoundRollover() : End Sub
Sub sw58_unHit:Controller.Switch(58)=0:End Sub
Sub sw60_Hit:Controller.Switch(60)=1 : RandomSoundRollover() : End Sub
Sub sw60_unHit:Controller.Switch(60)=0:End Sub
Sub sw61_Hit:Controller.Switch(61)=1 : RandomSoundRollover() : End Sub
Sub sw61_unHit:Controller.Switch(61)=0:End Sub


'Ramp Gate Triggers
Sub sw18_hit:vpmTimer.pulseSw 18 : playsoundAt "rollover", ActiveBall : End Sub

Sub sw35_hit:vpmTimer.pulseSw 35 : playsoundAt "fx_metal_hit_2", ActiveBall : End Sub

'Ramp Gates

Sub sw37_hit:vpmTimer.pulseSw 37 :playsoundAt "fx_gate4", ActiveBall: End Sub
Sub sw37_Unhit:vpmTimer.pulseSw 37 : End Sub

Sub sw38_hit:vpmTimer.pulseSw 38 :playsoundAt "fx_gate4", ActiveBall: End Sub
Sub sw38_Unhit:vpmTimer.pulseSw 38 : End Sub

Sub sw41_hit:vpmTimer.pulseSw 41 :playsoundAt "fx_gate4", ActiveBall: End Sub
Sub sw41_Unhit:vpmTimer.pulseSw 41 : End Sub

Sub sw42_hit:vpmTimer.pulseSw 42 :playsoundAt "fx_gate4", ActiveBall: End Sub
Sub sw42_Unhit:vpmTimer.pulseSw 42 : End Sub

'Scoop
Sub Sw46_hit: Sw46.enabled = 0 : End Sub
Sub Sw46b_UnHit: vpmtimer.addtimer 500, "Sw46.enabled = 1 '" :End Sub

'**********************************************************************************************************
'LaserBeam
'**********************************************************************************************************

Const Pi=3.1415926535
Dim BallInGun, GPos, BallInGunRadius
BallInGunRadius = SQR((PCannon.X - Sw45.X)^2 + (PCannon.Y - Sw45.Y)^2)

 Sub SolLaserBeam(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("CannoShot", DOFContactors), sw45
    If NOT IsEmpty(BallInGun) Then
      Sw45.kick 300 + GPos, 50
      Controller.switch(45) = 0
      BallInGun = Empty
    End If
     End If
End Sub

 Sub UpdateLserGun(acurrpos, aspeed, alastpos)
  GPos = acurrpos / 10
  PCannon.RotY = GPos
  PCannonshadow.RotY  = GPos
  if GPos > 5 Then
    CannonStatus = True
  Else
    CannonStatus = False
  end If
  If Not IsEmpty(BallInGun) Then
    BallInGun.X = PCannon.X - BallInGunRadius * Cos((GPos+30)*Pi/180)
    BallInGun.Y = PCannon.Y - BallInGunRadius * Sin((GPos+30)*Pi/180)
  End If
 End Sub

Sub Sw45_hit(): StopSound "fx_ramp_enter3": PlaySoundAt "popper_ball", sw45: controller.switch(45) = 1:Set BallInGun = ActiveBall:End Sub

'**********************************************************************************************************
'Dancing Austin
'**********************************************************************************************************

Dim cBall
DanceInit

Sub DanceInit
    Set cBall = ckicker.createball
    ckicker.Kick 0, 0
End Sub


Sub SolAustinDance(Enabled)
    If Enabled Then
        AustinDance
    End If
End Sub

Sub AustinDance
    cball.velx = 47
End Sub

'**********************************************************************************************************
'Minimi Spiner
'**********************************************************************************************************
Dim SpinPos, SpinSpeed, SpinSwState
SpinPos = 0
SpinSpeed = 0
SpinSwState = 0

Sub Sw34_Spin()::End Sub

Sub Sw34Anim_Hit()
    SpinSpeed = -4.2 * (ActiveBall.VelY / 2)
    SpinTimer.Enabled = 1
End Sub

Sub Spintimer_Timer()
    SpinPos = SpinPos + SpinSpeed
  SpinSpeed = (SpinSpeed * 0.98) - (sin(SpinPos * 3.14159 / 180) * 1)
    if SpinPos>360 Then
       SpinPos = SpinPos - 360
    elseif SpinPos < 0 Then
       SpinPos = SpinPos + 360
    end if
    if SpinPos>125 and SpinPos<215 Then
    if SpinSwState =0 Then
      vpmTimer.PulseSw 34
      playsoundat "fx_spinner", sw34Anim
      SpinSwState = 1
    end if
  Else
    SpinSwState = 0
  End If
    MiniMiP.RotX = SpinPos
  if abs(SpinSpeed) < 0.0001 then SpinTimer.Enabled =0

End Sub

'**********************************************************************************************************
'Dr Evil "By JPSalas" and DJRobX
'**********************************************************************************************************
Dim Sw24Pos
Dim Sw24Strength

Sub Sw24_Hit()
  Sw24pos=0
  Sw24Strength = 25+BallVel(ActiveBall)/4:vpmTimer.PulseSw 24
  if Sw24Strength > 25 Then PlaySoundAtVol "fx_chapa", ActiveBall, Abs(Sw24Strength-25)/5 '26-28
  Me.TimerInterval=7
  Me.TimerEnabled=1
End Sub

Sub Sw24_Timer()
  Sw24p.y = 657 - sin((sw24pos + (sw24pos ^ 2)) * .3) * (10 * ((Sw24Strength-sw24pos) / 30))
  if sw24pos >= Sw24Strength then Me.TimerEnabled = 0:Sw24p.y = 657
  sw24pos = sw24pos + 1
 End Sub

Sub UpdateEvil(aNewPos,aSpeed,aLastPos)
    Sw24P.TransY = aNewPos * 3 / 10

  'debug.print "aNewPos: " & aNewPos & "  aLastPos: " & aLastPos' & "   motorStarted: " & motorStarted & "   motorStop: " & motorStop
  If aNewPos > 40 then
        sw24.IsDropped = False
    Else
        sw24.Isdropped = True
    End If

  If aNewPos >= 498 then
    'Debug.print "restart play near up position"
    StopSound "TOM_Trunk_Motor_Long"
    PlaySoundAtVol "TOM_Trunk_Motor_Stop", sw24p, 0.3
    PlaySoundAtVol "TOM_Trunk_Motor_Long", sw24p, 0.3
  End if


End Sub




'*************************************************************************************
'   GENERAL ILLUMINATION
'*************************************************************************************
set GICallback = GetRef("UpdateGI")

Sub UpdateGI(no, Enabled)
  If Enabled Then
    dim xx
    For each xx in GILights:xx.State = 1: Next
    For each xx in GILightsBackwalls:xx.visible=1: Next  'BackWall Flashers
        PlaySoundAt "fx_relay_on", Primitive6
        Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
      'Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSatDark"
    If VRRoom=1 Then
          RightRampPrimitive.BlendDisableLighting = 0.2
      LeftRampPrimitive.BlendDisableLighting = 0.2
    Else
      RightRampPrimitive.BlendDisableLighting = 0.2
      LeftRampPrimitive.BlendDisableLighting = 0.4
    End if
    PCannon.BlendDisableLighting = 0.11
    PCannon_bracket.BlendDisableLighting = 0.07
    Primitive4.BlendDisableLighting = 0.1 'fatguy
    AustinDancerP.BlendDisableLighting = 0.1
    MiniMiP.BlendDisableLighting = 0.1
    Sw24P.BlendDisableLighting = 0.1
    TimeMachineP.BlendDisableLighting = 0.1
    pLaneguide1.BlendDisableLighting = 0.9
    pLaneguide2.BlendDisableLighting = 0.9
    pLaneguide3.BlendDisableLighting = 0.9
    FlasherLight27.state = 1
    FlasherLight27.intensity = 2
    F127a.BlendDisableLighting = 35
    Primitive196.BlendDisableLighting = 0.1
    PinCab_Backglass.blenddisablelighting = 15
    SSLCrystal.BlendDisableLighting = 0.5
    Primitive_RampDiverter.blenddisablelighting = 0.06
    MrBigglesworth.blenddisablelighting = 0.3
    'BumperCap1.blenddisablelighting = 0.5
  Else
    For each xx in GILights:xx.State = 0: Next
    For each xx in GILightsBackwalls:xx.visible=0: Next
        PlaySoundAt "fx_relay_off", Primitive6
      Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSatDark"
        RightRampPrimitive.BlendDisableLighting = 0
        LeftRampPrimitive.BlendDisableLighting = 0
    PCannon.BlendDisableLighting = 0.07
    PCannon_bracket.BlendDisableLighting = 0.05
    Primitive4.BlendDisableLighting = 0 'fatguy
    AustinDancerP.BlendDisableLighting = 0
    MiniMiP.BlendDisableLighting = 0
    Sw24P.BlendDisableLighting = 0
    TimeMachineP.BlendDisableLighting = 0
    pLaneguide1.BlendDisableLighting = 0.05
    pLaneguide2.BlendDisableLighting = 0.05
    pLaneguide3.BlendDisableLighting = 0.05
    FlasherLight27.state = 0
    F127a.BlendDisableLighting = 0
    Primitive196.BlendDisableLighting = 0
    PinCab_Backglass.blenddisablelighting = 0.2
    SSLCrystal.BlendDisableLighting = 0.1
    Primitive_RampDiverter.blenddisablelighting = 0.01
    MrBigglesworth.blenddisablelighting = 0.02
    'BumperCap1.blenddisablelighting = 0.1
  End If
End Sub

' #####################################
' ###### Flupper Flasher Domes    #####
' #####################################

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.1   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 1   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.5
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20)
Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
''initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "green" : InitFlasher 2, "yellow" : InitFlasher 3, "blue" : InitFlasher 4, "red"
'' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 4,17 : RotateFlasher 5,0 : RotateFlasher 6,90
'RotateFlasher 7,0 : RotateFlasher 8,0
'RotateFlasher 9,-45 : RotateFlasher 10,90 : RotateFlasher 11,90


Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 40
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness
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
  FlasherFlash4.height = 226
  FlasherFlash3.height = 200
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub


' ###############################################
' ###### Austin Power flasher dome settings #####
' ###############################################

Sub FlashGreen(flstate)
  If Flstate Then
    Objlevel(1) = 1 : FlasherFlash1_Timer
  End If
End Sub

Sub FlashYellow(flstate)
  If Flstate Then
    Objlevel(2) = 1 : FlasherFlash2_Timer
  End If
End Sub

Sub FlashBlue(flstate)
  If Flstate Then
    Objlevel(3) = 1 : FlasherFlash3_Timer
  End If
End Sub

Sub FlashRed(flstate)
  If Flstate Then
    Objlevel(4) = 1 : FlasherFlash4_Timer
  End If
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim bulb
Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 7 'lamp fading speed
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
        FlashSpeedUp(x) = 0.2    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.05 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
  NFadeL 1, Light1
  FadeDisableLighting2 1, p1, 20
  NFadeL 2, Light2
  FadeDisableLighting2 2, p2, 20
  NFadeL 3, Light3
  FadeDisableLighting2 3, p3, 20
  NFadeL 4, Light4
  FadeDisableLighting2 4, p4, 20
  NFadeL 5, Light5
    Flash 5, Light5a
  FadeDisableLighting2 5, p5, 20
  NFadeL 6, Light6
    Flash 6, Light6a
  FadeDisableLighting2 6, p6, 20
  NFadeL 7, Light7
  FadeDisableLighting2 7, p7, 80
  NFadeL 8, Light8
  FadeDisableLighting2 8, p8, 100
  NFadeL 9, Light9
  FadeDisableLighting2 9, Sw25b, 4
  NFadeL 10, Light10
  FadeDisableLighting2 10, Sw26b, 4
  NFadeL 11, Light11
  FadeDisableLighting2 11, Sw27b, 4
  NFadeL 12, Light12
  FadeDisableLighting2 12, Sw28b, 4

  nFadeLm 13, BumperL_Flasher
  nFadeLm 13, BumperL_Flasher_a
  Flash 13, BumperL_Flasher1

  nFadeLm 14, BumperR_Flasher
  nFadeLm 14, BumperR_Flasher_a
  Flash 14, BumperR_Flasher1

  nFadeLm 15, BumperT_Flasher
  nFadeLm 15, BumperT_Flasher_a
  Flash 15, BumperT_Flasher1

  NFadeL 16, Light16
  FadeDisableLighting2 16, p16, 150
  NFadeL 17, Light17
  FadeDisableLighting2 17, p17, 20
  NFadeL 18, Light18
  FadeDisableLighting2 18, p18, 20
  NFadeL 19, Light19
  FadeDisableLighting2 19, p19, 20
  NFadeL 20, Light20
  FadeDisableLighting2 20, p20, 20
  NFadeL 21, Light21
  FadeDisableLighting2 21, p21, 20
  NFadeL 22, Light22
  FadeDisableLighting2 22, p22, 20
  NFadeL 23, Light23
  FadeDisableLighting2 23, p23, 20
  NFadeL 24, Light24
  FadeDisableLighting2 24, p24, 20
  NFadeL 25, Light25
  FadeDisableLighting2 25, p25, 20
  NFadeL 26, Light26
  FadeDisableLighting2 26, p26, 20
  NFadeL 27, Light27
  FadeDisableLighting2 27, p27, 20
  NFadeL 28, Light28
  FadeDisableLighting2 28, p28, 20
  NFadeL 29, Light29
  FadeDisableLighting2 29, p29, 20
  NFadeL 30, Light30
  FadeDisableLighting2 30, p30, 20
  NFadeL 31, Light31
  FadeDisableLighting2 31, p31, 20
  NFadeL 32, Light32
  FadeDisableLighting2 32, p32, 80
  NFadeL 33, Light33
  FadeDisableLighting2 33, Sw29b, 4
  NFadeL 34, Light34
  FadeDisableLighting2 34, Sw30b, 4
  NFadeL 35, Light35
  FadeDisableLighting2 35, p35, 150
  NFadeL 36, Light36
  FadeDisableLighting2 36, p36, 150
  NFadeL 37, Light37
  FadeDisableLighting2 37, p37, 150
  NFadeL 38, Light38
  FadeDisableLighting2 38, p38, 150
  NFadeL 39, Light39
  FadeDisableLighting2 39, p39, 150
  NFadeL 40, Light40
  FadeDisableLighting2 40, p40, 150
  NFadeL 41, Light41
  FadeDisableLighting2 41, p41, 50
  NFadeL 42, Light42
  FadeDisableLighting2 42, p42, 40
  NFadeL 43, Light43
  FadeDisableLighting2 43, p43, 40
  NFadeL 44, Light44
  FadeDisableLighting2 44, p44, 50
  NFadeL 45, Light45
    Flash 45, Light45a
  FadeDisableLighting2 45, p45, 150
  NFadeL 46, Light46
  FadeDisableLighting2 46, p46, 150
  NFadeL 47, Light47
  FadeDisableLighting2 47, p47, 150
  NFadeL 48, Light48
    Flash 48, Light48a
  FadeDisableLighting2 48, p48, 150
  NFadeL 49, Light49
  FadeDisableLighting2 49, p49, 80
  NFadeL 50, Light50
  FadeDisableLighting2 50, p50, 80
  NFadeL 51, Light51
  FadeDisableLighting2 51, p51, 80
  NFadeL 52, Light52
  FadeDisableLighting2 52, p52, 80
    NFadeObjm 53, l53, "FireButtonOn", "FireButton"
    NFadeL 53, F53 'Fire Button
  NFadeL 54, Light54 'DR Evil
  FadeDisableLighting2 54, p54, 10
  NFadeL 57, Light57
  FadeDisableLighting2 57, p57, 30
  NFadeL 58, Light58
  FadeDisableLighting2 58, p58, 20
  NFadeL 59, Light59
  FadeDisableLighting2 59, p59, 20
  NFadeL 60, Light60
  FadeDisableLighting2 60, p60, 20
  NFadeL 61, Light61
  FadeDisableLighting2 61, p61, 150
  NFadeL 62, Light62
  FadeDisableLighting2 62, p62, 150
  NFadeL 63, Light63
  FadeDisableLighting2 63, p63, 100
  NFadeL 64, Light64
  FadeDisableLighting2 64, p64, 80
  NFadeL 65, Light65
  FadeDisableLighting2 65, p65, 80
  NFadeL 66, Light66
  FadeDisableLighting2 66, p66, 80
  NFadeL 67, Light67
  FadeDisableLighting2 67, p67, 80
  NFadeL 68, Light68
  FadeDisableLighting2 68, p68, 80
  NFadeL 73, Light73
  FadeDisableLighting2 73, p73, 20
  NFadeL 74, Light74
  FadeDisableLighting2 74, p74, 20
  NFadeL 75, Light75
  FadeDisableLighting2 75, p75, 20
  NFadeL 76, Light76
  FadeDisableLighting2 76, p76, 20

'Solenoid Controlled Lamps
  NFadeLm 105, F105
  NFadeL 105, F105a
  NFadeL 125, FlasherLight25
  NFadeL 126, FlasherLight26
  NFadeL 128, FlasherLight28
  NFadeL 129, FlasherLight29
  NFadeL 130, FlasherLight30
  NFadeL 131, FlasherLight31
  Flash 131, Flasherbloom31
  Flashm 131, Flasherbloom31a
  Flashm 131, Flasherbloom31b



' Flash 106, Flasher6b
' NFadeLm 106, FlasherL6
' Flashm 106, Flasher6
' Flashm 106, Flasher6m
' Primm 106, Flasher6p
'
' Flash 112, Flasher12b
' NFadeLm 112, FlasherL12
' Flashm 112, Flasher12m
' Flashm 112, Flasher12
' Primm 112, Flasher12p
'
' Flash 113, Flasher13b
' NFadeLm 113, FlasherL13
' Flashm 113, Flasher13m
' Flashm 113, Flasher13
' Primm 113, Flasher13p

  Flash 127, Flasher27m
  'NFadeLm 127, FlasherLight27
  'I don't "get" this light - the flasher is supposed to be the right ramp, but this selectively lights up the lower PF?
    Flashm 127, Flasher27
  'FadeDisableLighting2 127, F127a, 10
  'FadeDisableLighting2 127, Primitive196, 0.2

' Flash 132, Flasher32b
' NFadeLm 132, FlasherL32
' Flashm 132, Flasher32m
' Flashm 132, Flasher32
' Primm 132, Flasher32p
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

Sub FadeDisableLighting2(nr, a, alvl)
  Select Case FadingLevel(nr)
    Case 4
      a.UserValue = a.UserValue - 0.125
      If a.UserValue < 0 Then
        a.UserValue = 0
        'FadingLevel(nr) = 0
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
    Case 5
      a.UserValue = a.UserValue + 0.250
      If a.UserValue > 1 Then
        a.UserValue = 1
        'FadingLevel(nr) = 1
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
  End Select
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0':FadingLevel(nr) = 0
        Case 5:object.state = 1':FadingLevel(nr) = 1
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
        Case 4:object.image = b:FadingLevel(nr) = 0
        Case 5:object.image = a:FadingLevel(nr) = 1
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
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
    End Select
  Flashm nr, object
End Sub

Sub Primm(nr, Object)
  Object.BlendDisableLighting = FlashLevel(nr)
end sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr) * FlasherIntensity / 100
End Sub

' RGB Leds

Sub RGBLED (object,red,green,blue)
  If TypeName(object) = "Light" Then
    object.color = RGB(0,0,0)
    object.colorfull = RGB(2.5*red,2.5*green,2.5*blue)
    object.state=1
  ElseIf TypeName(object) = "Flasher" Then
    object.color = RGB(2.5*red,2.5*green,2.5*blue)
    object.IntensityScale = 1
  End If
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
  Object.IntensityScale = FadingLevel(nr)/128
  If TypeName(object) = "Light" Then
    Object.State = LampState(nr)
  End If
  If TypeName(object) = "Flasher" Then
    Object.visible = LampState(nr)
  End If
End Sub


'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10,BallShadow11)

Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' hide shadow of deleted balls

  For b = UBound(BOT) + 1 to tnob
    BallShadow(b).visible = 0
  Next

  ' exit the Sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub
  ' render the shadow for each ball
  For b = 0 to UBound(BOT)
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
    End If
    ballShadow(b).Y = BOT(b).Y + 20

    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'******************************
'    Sound FX
'******************************

'Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
'Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):dPosts():End Sub
'Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):dPosts():End Sub
'Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
'Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
'Sub aWoods_Hit(idx):PlaySound "fx_woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub


Sub Trigger2_Hit:PlaySoundAt "fx_metal_hit_2", Trigger2 End Sub
'

' Ball Collision Sound
'Sub OnBallBallCollision(ball1, ball2, velocity)
'    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / VolDiv, Pan(ball1), 0, Pitch(ball1), 0, 0
'End Sub

Sub LREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAt "fx_ramp_enter1",LREnter:End If:End Sub     'ball is going up
Sub LREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_ramp_enter1":End If:End Sub   'ball is going down
Sub LREnter1_Hit():PlaySoundAt "fx_ramp_enter2",LREnter1:End Sub
Sub LREnter2_Hit():PlaySoundAt "fx_ramp_enter3",LREnter2:End Sub

Sub LRExit_Hit():ActiveBall.VelY=1:StopSound "fx_ramp_enter3":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub LRExit_timer():Me.TimerEnabled=0:End Sub 'RandomDropSound:End Sub

Sub RREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAt "fx_ramp_enter1", RRenter:End If:End Sub      'ball is going up
Sub RREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_ramp_enter1":End If:End Sub   'ball is going down
Sub RREnter1_Hit():PlaySoundAt "fx_ramp_enter2",RRenter1:End Sub
Sub RREnter2_Hit():PlaySoundAt "fx_ramp_enter3",RRenter2:End Sub

Sub RRExit_Hit():StopSound "fx_ramp_enter3":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub RRExit_timer():Me.TimerEnabled=0:End Sub 'RandomDropSound:End Sub

Sub TriggerLaunchBall_Hit(): vpmtimer.addtimer 350, "RandomDropSound '" End Sub

Sub RandomDropSound()
  Select Case Int(Rnd*6)
    Case 0 : PlaySound "fx_Ball_drop0"
    Case 1 : PlaySound "fx_Ball_drop1"
    Case 2 : PlaySound "fx_Ball_drop2"
    Case 3 : PlaySound "fx_Ball_drop3"
    Case 4 : PlaySound "fx_Ball_drop4"
    Case 5 : PlaySound "fx_Ball_drop5"
  End Select
End Sub

'Sub Rubbers_Hit(idx)
  'debug.print "Rubbers"
  'Select Case Int(Rnd*3)+1
  ' Case 1 : PlaySoundAtBallVol "fx_rubber_hit_1", .8
  ' Case 2 : PlaySoundAtBallVol "fx_rubber_hit_2", .8
  ' Case 3 : PlaySoundAtBallVol "fx_rubber_hit_3", .8
  'End Select
'End Sub

'Sub RandomSoundTargetHit()
' debug.print "old"
' Select Case Int(Rnd*3)+1
'   Case 1 : PlaySoundAtBallVol "TOM_Target_Hit_1", .8
'   Case 2 : PlaySoundAtBallVol "TOM_Target_Hit_2", .8
'   Case 3 : PlaySoundAtBallVol "TOM_Target_Hit_3", .8
' End Select
'End Sub


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
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, PegPlasticT4
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, PegPlasticT18
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
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*11)+1,DOFFlippers), FlipperRightHitParm, Flipper
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

Sub LeftFlipper_Collide(parm)
  FlipperLeftHitParm = parm/10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  CheckFlipperSlack Activeball, LeftFlipper, LFLogo, parm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipper_Collide(parm)
  FlipperRightHitParm = parm/10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  CheckFlipperSlack Activeball, RightFlipper, RFLogo, parm
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

'Sub Apron_Hit (idx)
' If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
'   RandomSoundBottomArchBallGuideHardHit()
' Else
'   RandomSoundBottomArchBallGuide
' End If
'End Sub

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
  PlaySoundAtLevelActiveBall ("Target_Hit_" & Int(Rnd*4)+5), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall ("Target_Hit_" & Int(Rnd*4)+1), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    'RandomSoundBallBouncePlayfieldSoft Activeball
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  'msgbox "sound"
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
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 10 ' total number of balls
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

Sub RollingUpdate()
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
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 28 Then 'height adjust for ball drop sounds
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
  addpt "Velocity", 1, 0.1,   1.07
  addpt "Velocity", 2, 0.2,   1.15
  addpt "Velocity", 3, 0.3,   1.25
  addpt "Velocity", 4, 0.41, 1.05
  addpt "Velocity", 5, 0.65,  1.0'0.982
  addpt "Velocity", 6, 0.702, 0.968
  addpt "Velocity", 7, 0.95,  0.968
  addpt "Velocity", 8, 1.03,  0.945

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -5.5
  AddPt "Polarity", 2, 0.4, -5.5
  AddPt "Polarity", 3, 0.8, -5.5
  AddPt "Polarity", 4, 0.85, -5.25
  AddPt "Polarity", 5, 0.9, -4.25
  AddPt "Polarity", 6, 0.95, -3.75
  AddPt "Polarity", 7, 1, -3.25
  AddPt "Polarity", 8, 1.05, -2.25
  AddPt "Polarity", 9, 1.1, -1.5
  AddPt "Polarity", 10, 1.15, -1
  AddPt "Polarity", 11, 1.2, -0.5
  AddPt "Polarity", 12, 1.25, 0
  AddPt "Polarity", 13, 1.3, 0

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
dim LFFlipBall, RFFlipBall : LFFlipBall = 0 : RFFlipBall = 0


Sub TriggerLF_Hit() : LF.Addball activeball : LFFlipBall = 1 : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : LFFlipBall = 0 : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : RFFlipBall = 1 : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : RFFlipBall = 0 : End Sub

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
        'debug.print BallPos & " " & AddX & " " & Ycoef & " "& PartialFlipcoef & " "& VelCoef
        'playsound "fx_knocker"
      End If
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

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
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
Const EOSTnew = 1
Const EOSAnew = 0.8
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
'Const LiveCatch = 16
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
    ball.vely = LiveCatchBounce * (16 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
    debug.print "Live catch! Time: " & CatchTime & " Bounce: " & LiveCatchBounce & " parm: " & parm
  End If

End Sub


Const FlipY = 1810      'Default flip Y position
Const Dislocation = 2   'Amount of minimum slack, affect to recovery speed too
Const SlackHitLimit = 10  '6-16 Lower value will make slack happen with lower collision force
Const HitForceDivider = 50  '10-100 Lower value will add slack. Set to roughly half of the max hit force(parm) you see when playing
Const MaxSlack = 4
Dim SlackAmount

Sub CheckFlipperSlack(ball, Flipper, Logo, parm)
  If slackyFlips = 1 Then
    SlackAmount = Dislocation + parm / HitForceDivider
    if SlackAmount > MaxSlack then SlackAmount = MaxSlack
    if parm > SlackHitLimit Then
      Flipper.Y = FlipY + SlackAmount
      Logo.Y = FlipY + SlackAmount
      slacktimer.interval = 10
      slacktimer.Enabled = 1
      'debug.print parm & " parm and SlackAmount " & SlackAmount
    end If
  end If
end Sub

Sub slacktimer_timer()
  'msgbox "location change" & RightFlipper.Y & " and " & LeftFlipper.Y
  if RightFlipper.Y > FlipY + 2 Then
    RightFlipper.Y = RightFlipper.Y - 2
    RFLogo.Y = RFLogo.Y - 2
  else
    RightFlipper.Y = FlipY
    RFLogo.Y = FlipY
  End If
  if LeftFlipper.Y > FlipY + 2 Then
    LeftFlipper.Y = LeftFlipper.Y - 2
    LFLogo.Y = LFLogo.Y - 2
  else
    LeftFlipper.Y = FlipY
    LFLogo.Y = FlipY
  End If

  if LeftFlipper.Y = FlipY and RightFlipper.Y = FlipY Then slacktimer.Enabled = 0
end Sub


'Sub slacktimer_timer()
' 'msgbox "location change" & RightFlipper.Y & " and " & LeftFlipper.Y
' if RightFlipper.Y > FlipY Then
'   RightFlipper.Y = FlipY
'   RFLogo.Y = FlipY
' End If
' if LeftFlipper.Y > FlipY Then
'   LeftFlipper.Y = FlipY
'   LFLogo.Y = FlipY
' End If
'
' if LeftFlipper.Y = FlipY and RightFlipper.Y = FlipY Then slacktimer.Enabled = 0
'end Sub

'*****************************************************************************************************
'*******************************************************************************************************
'END nFOZZY FLIPPERS'

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts()
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  PlaySoundAtBallVol "fx_rubber", .8
  SleevesD.Dampen Activeball
End Sub


Sub RDampen_Timer()
  Cor.Update
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

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    'playsound "fx_knocker"
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
    'if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
'     if DebugOn then
'       dim s, bs 'debug spacer, ballspeed
'       bs = round(BallSpeed(b),1)
'       if bs < 10 then s = " " else s = "" end if
'       str = str & b.id & ": " & s & bs & vbnewline
'       'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
'     end if
    Next
    'if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
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

' *****  cvpmMech replacement, all because the "Fast" parameter is not set on timer update!!   We can't get smooth reels without this.   Until this is addressed in Core.vbs, replacing here.

Class cvpmMyMech
  Public Sol1, Sol2, MType, Length, Steps, Acc, Ret
  Private mMechNo, mNextSw, mSw(), mLastPos, mLastSpeed, mCallback

  Private Sub Class_Initialize
    ReDim mSw(10)
    gNextMechNo = gNextMechNo + 1 : mMechNo = gNextMechNo : mNextSw = 0 : mLastPos = 0 : mLastSpeed = 0
    MType = 0 : Length = 0 : Steps = 0 : Acc = 0 : Ret = 0 : vpmTimer.addResetObj Me
  End Sub

  Public Sub AddSw(aSwNo, aStart, aEnd)
    mSw(mNextSw) = Array(aSwNo, aStart, aEnd, 0)
    mNextSw = mNextSw + 1
  End Sub

  Public Sub AddPulseSwNew(aSwNo, aInterval, aStart, aEnd)
    If Controller.Version >= "01200000" Then
      mSw(mNextSw) = Array(aSwNo, aStart, aEnd, aInterval)
    Else
      mSw(mNextSw) = Array(aSwNo, -aInterval, aEnd - aStart + 1, 0)
    End If
    mNextSw = mNextSw + 1
  End Sub

  Public Sub Start
    Dim sw, ii
    With Controller
      .Mech(1) = Sol1 : .Mech(2) = Sol2 : .Mech(3) = Length
      .Mech(4) = Steps : .Mech(5) = MType : .Mech(6) = Acc : .Mech(7) = Ret
      ii = 10
      For Each sw In mSw
        If IsArray(sw) Then
          .Mech(ii) = sw(0) : .Mech(ii+1) = sw(1)
          .Mech(ii+2) = sw(2) : .Mech(ii+3) = sw(3)
          ii = ii + 10
        End If
      Next
      .Mech(0) = mMechNo
    End With
    If IsObject(mCallback) Then mCallBack 0, 0, 0 : mLastPos = 0 : vpmTimer.EnableUpdate Me, True, True  ' <------- All for this.
  End Sub

  Public Property Get Position : Position = Controller.GetMech(mMechNo) : End Property
  Public Property Get Speed    : Speed = Controller.GetMech(-mMechNo)   : End Property
  Public Property Let Callback(aCallBack) : Set mCallback = aCallBack : End Property

  Public Sub Update
    Dim currPos, speed
    currPos = Controller.GetMech(mMechNo)
    speed = Controller.GetMech(-mMechNo)
    If currPos < 0 Or (mLastPos = currPos And mLastSpeed = speed) Then Exit Sub
    mCallBack currPos, speed, mLastPos : mLastPos = currPos : mLastSpeed = speed
  End Sub

  Public Sub Reset : Start : End Sub

End Class

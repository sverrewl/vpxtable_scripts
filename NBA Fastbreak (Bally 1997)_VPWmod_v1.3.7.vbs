'************************ VPW PRESENTS ***************************
' _   _ ______  ___   ______        _   _                    _
'| \ | || ___ \/ _ \  |  ___|      | | | |                  | |
'|  \| || |_/ / /_\ \ | |_ __ _ ___| |_| |__  _ __ ___  __ _| | __
'| . ` || ___ \  _  | |  _/ _` / __| __| '_ \| '__/ _ \/ _` | |/ /
'| |\  || |_/ / | | | | || (_| \__ \ |_| |_) | | |  __/ (_| |   <
'\_| \_/\____/\_| |_/ \_| \__,_|___/\__|_.__/|_|  \___|\__,_|_|\_\
'
'*****************************************************************
'
' NBA Fastbreak by Bally (1997)
' http://www.ipdb.org/machine.cgi?id=4023
'
'************************************
' VPin Workshop NBA MVP's
'************************************
' Project Lead - Tomate
' Graphics / 3D Work - Tomate, Sixtoe, Benji
' Scripting - iaakki, Benji, Sixtoe
' Sound - iaakki, Benji, Sixtoe
' Physics - iaakki, Benji
' Lighting - iaakki, Benji, Sixtoe
' Shadows - Sixtoe, Wylte
' VR Stuff & Tuning - Sixtoe
' VR Backglass - Rawd
'
'Heavily modified and based on;
' VP915 v1.0 by JPSalas 2013
' VPX v1.0 by MaX 2016
' based on the tables by Aurich and bmiki75
' 2018 Mod by Darth Marino. Extra special thanks to DJROBX for helping implement the Shot Clock
' Fast Flips implemented by DJRobX and CarnyPriest

Option Explicit
Randomize
SetLocale 1033      'Forces VBS to use english to stop crashes.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

Const UseVPMModSol = 2    'Set to 2 for PWM flashers, inserts, and GI. Requires VPinMame 3.6

Const BallSize = 50                 'add this here to redefine the ball size, placed before LoadVPM.'
Const BallMass = 1

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height


'*************************************
Const cGameName = "nbaf_31"
' Standard Sounds
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 1
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^

Const cSingleLFlip=0
Const cSingleRFlip=0

LoadVPM "03060000", "WPC.VBS", 3.49




Dim bsTrough, bsEject, bsSaucer1, bsSaucer2, bsSaucer3, bsSaucer4, mBallCatch, mDefender, MagnetCatch, bgball
Dim PlungerIM, x, bump1, bump2, bump3, ArenaMod, ff

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5

'***********CABINET MODE**************
Dim CabinetMode     'Cabinet mode - Will hide the rails and scale the side panels higher

'**************VR ROOM****************
Dim VRRoom

'********Dynamic Ball Shadows*********
'*** *BOTH* to 0 for simple shadow****
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's)
                  '2 = flasher image shadow, but it moves like ninuzzu's

'*****Shooter Lane Shoot Button*******
Dim ShootShooter

'*********RED SHOT CLOCK MOD**********
Dim RedClock

'********TABLE FLASHER BLOOM**********
' TURN OFF IF TABLE STUTTERS
Dim TableBloom

'*********************************************************************************************************************************

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim LiveViewVRSim : LiveViewVRSim = 0

Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 11, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 1 Then Table1.ColorGradeImage = ""
  if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100"

  ShootShooter = Table1.Option("Shooter Lane Shoot Button", 0, 1, 1, 1, 0, Array("Off", "On"))

  RedClock = Table1.Option("Red Shot Clock Mod", 0, 1, 1, 1, 0, Array("Off", "On"))

  TableBloom = Table1.Option("Flasher Bloom", 0, 1, 1, 0, 0, Array("Off", "On"))

  CabinetMode = Table1.Option("Cabinet Mode", 0, 1, 1, 0, 0, Array("False", "True"))

    VRRoom = Table1.Option("VR Room", 1, 3, 1, 3, 0, Array("Minimal", "Medium", "Court"))
  SetupVRRoom

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100

  UpdateOptions

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Sub SetupVRRoom()
  Dim VRThings
  If RenderingMode = 2 Or LiveViewVRSim > 0 Then
    ScoreText.visible = 0
    Ramp_central.image = "VR_centre_ramp"
    ramps_lateral.image = "VR_lateral_ramps"
    Ramp_central.blenddisablelighting = 0.5
    ramps_lateral.blenddisablelighting = 0.5
    'PinCab_Backglass.blenddisablelighting = 4 ' VRADDED - removed
    VR_Wall_Right.blenddisablelighting = 2
    VR_Wall_Left.blenddisablelighting = 2
    ShootButton.visible=0
    If VRRoom = 1 Then
      for each VRThings in VRCab:VRThings.visible = 1:Next
      for each VRThings in VRStuff:VRThings.visible = 1:Next
      for each VRThings in Desktop:VRThings.visible = 0:Next
      for each VRThings in VRMega:VRThings.visible = 0:Next
      PinCab_MetalsRearMega.visible = 0
    End If
    If VRRoom = 2 Then
      for each VRThings in VRCab:VRThings.visible = 0:Next
      for each VRThings in VRStuff:VRThings.visible = 0:Next
      for each VRThings in Desktop:VRThings.visible = 0:Next
      for each VRThings in VRMega:VRThings.visible = 0:Next
      PinCab_MetalsRearMega.visible = 0
      'PinCab_Backglass.visible = 1
      DMD.visible = true ' added because its part of the VRCab collection which is made invisible above.
      PinCab_Backbox.visible = 1
      PinCab_Backbox.image = "Pincab_Backbox_NBA_Min"
    End If
    If VRRoom = 3 Then
      for each VRThings in VRCab:VRThings.visible = 1:Next
      for each VRThings in VRStuff:VRThings.visible = 0:Next
      for each VRThings in Desktop:VRThings.visible = 0:Next
      for each VRThings in VRMega:VRThings.visible = 1:Next
      PinCab_MetalsRear.visible = 0
    End If
      BGDarkfront.visible = 1
      NBABackglassBack.visible = 1
  Else
      for each VRThings in VRCab:VRThings.visible = 0:Next
      for each VRThings in VRStuff:VRThings.visible = 0:Next
      for each VRThings in Desktop:VRThings.visible = 1:Next
      for each VRThings in VRBGStuff:VRThings.visible = 0:Next
      for each VRThings in VRMega:VRThings.visible = 0:Next

      If DesktopMode then ScoreText.visible = 1 else ScoreText.visible = 0 End If
      Ramp_central.image = "ramp_centre"
      ramps_lateral.image = "ramp_lateral"
  End if
End Sub


Sub UpdateOptions()
  if ShootShooter = 1 then
    ShootButton.visible = 1
  else
    ShootButton.visible = 0
  end if

  Dim Cabflash
  If CabinetMode = 1 Then
    SideWalls.visible = 1
    Pincab_Rails.visible = 0
    PinCab_Blades.visible = 0
    for each Cabflash in CabFlashers:Cabflash.height = 450:Next
    for each Cabflash in CabFlashers:Cabflash.Y = 1350:Next
  Else
    SideWalls.visible = 0
    PinCab_Rails.visible = 1
    PinCab_Blades.visible = 1
    for each Cabflash in CabFlashers:Cabflash.height = 255:Next
    for each Cabflash in CabFlashers:Cabflash.Y = 1120:Next
  End If
End Sub

'************************************
ramps_lateral.blenddisablelighting=2
ramp_central.blenddisablelighting=2
'************************************



Set GiCallback2 = GetRef("UpdateGI")

Sub UpdateGI(nr,step)
  Dim ii
  Select Case nr
  Case 0    'Bottom Playfield
    'string 1
    If step=0 Then
      For each ii in GI0:ii.state=step:Next
    Else
      For each ii in GI0:ii.state=step:Next
    End If
  Case 1    'Middle Playfield
    'string 2
    If step=0 Then
      For each ii in GI1:ii.state=step:Next
    Else
      For each ii in GI1:ii.state=step:Next
    End If
  Case 2    'Top Playfield
    'string 3
    If step=0 Then
      For each ii in GI2:ii.state=step:Next
      If RenderingMode = 2 Or LiveViewVRSim = 1 Then
        For each ii in BGGI:ii.visible = 0:Next
        For each ii in VRBGGIBulbs:ii.disablelighting = 0:Next
      End If
    Else
      For each ii in GI2:ii.state=step:Next
      If RenderingMode = 2 Or LiveViewVRSim = 1 Then
        For each ii in BGGI:ii.visible = 1:Next
        For each ii in VRBGGIBulbs:ii.disablelighting = 1:Next
      End If
    End If
  End Select
End Sub

Sub table1_Init
    vpmInit Me
    Dim ii
    With Controller
        .GameName = cGameName
        .SplashInfoLine = "NBA Fastbreak, Bally 1997" & vbNewLine & "VPW"
   '     .Games(cGameName).Settings.Value("rol") = 0 'rotated vpm display
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = 0
'      .Games(cGameName).Settings.Value("dmd_pos_x")=0
'      .Games(cGameName).Settings.Value("dmd_pos_y")=0
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
    Controller.DIP(0) = &H00
    '  Controller.Run GetPlayerHWnd
    controller.run
    Controller.Switch(22) = 1 'close coin door
    Controller.Switch(24) = 1 'and keep it close

    If ArenaMod=1 then
    ' Wall80.isdropped=false
    ' Wall123.isdropped=false

    else
    ' Wall80.isdropped=True
    ' Wall123.isdropped=True
    End If

    vpmMapLights AllLamps

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot, bumper1, bumper2, bumper3)

    ' Trough & Ball Release
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 32, 33, 34, 35, 31, 0, 0
        .InitKick BallRelease, 180, 10
        .InitExitSnd "BallRelease2", "fx_solenoid"
        .InitEntrySnd "fx_solenoid", "fx_solenoid"
        .IsTrough = True
        .Balls = 4
    End With

    ' Ball Catch Magnet
    Set MagnetCatch = New cvpmMagnet
    With MagnetCatch
        .InitMagnet BCMagnet, 7
        .Solenoid = 8
        .CreateEvents "MagnetCatch"
    End With

    ' Eject
    Set bsEject = New cvpmBallStack
    With bsEject
        .InitSaucer sw25, 25, 165, 15
        .InitExitSnd "fx_popper", "fx_Solenoid"
    End With

    ' Saucers (In the Paint)
    Set bsSaucer1 = New cvpmBallStack
    With bsSaucer1
        .InitSaucer sw68, 68, 65, 32
        .Kickz = 1.2
        .InitExitSnd "fx_popper", "fx_Solenoid"
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    Set bsSaucer2 = New cvpmBallStack
    With bsSaucer2
        .InitSaucer sw67, 67, 26, 28
        .Kickz = 1.15
        .InitExitSnd "fx_popper", "fx_Solenoid"
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    Set bsSaucer3 = New cvpmBallStack
    With bsSaucer3
        .InitSaucer sw66, 66, 337, 28
        .Kickz = 1.15
        .InitExitSnd "fx_popper", "fx_Solenoid"
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    Set bsSaucer4 = New cvpmBallStack
    With bsSaucer4
        .InitSaucer sw65, 65, 293, 31.5
        .Kickz = 1.2
        .InitExitSnd "fx_popper", "fx_Solenoid"
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    ' Defender
    Set mDefender = New cvpmMech
    With mDefender
        .MType = vpmMechOneDirSol + vpmMechStopEnd + vpmMechLinear
        .Sol1 = 37 'Enable
        .Sol2 = 38 'Direction
        .Length = 65
        .Steps = 65
        .AddSw 51, 0, 1
        .AddSw 52, 17, 18
        .AddSw 53, 31, 32
        .AddSw 54, 44, 45
        .AddSw 55, 64, 65
        .CallBack = GetRef("UpdateDefender")
        .Start
    End With
    UpdateDefender 32, 32, 32

    'Impulse Plunger
    Const IMPowerSetting = 40 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Switch 15
        .Random 0
        .InitExitSnd "fx_plunger2", "fx_plunger"
        .CreateEvents "plungerIM"
    End With

    ' Misc. Initialisation
    LeftSLing.IsDropped = 1:LeftSLing2.IsDropped = 1:LeftSLing3.IsDropped = 1
    RightSLing.IsDropped = 1:RightSLing2.IsDropped = 1:RightSLing3.IsDropped = 1
    LeftSLingH.IsDropped = 1:LeftSLingH2.IsDropped = 1:LeftSLingH3.IsDropped = 1
    RightSLingH.IsDropped = 1:RightSLingH2.IsDropped = 1:RightSLingH3.IsDropped = 1


    For each ii in ADef1:ii.Isdropped = 1:Next
    For each ii in ADef2:ii.visible = 0:ii.Collidable = 0:Next

    If Not RenderingMode = 2 And LiveViewVRSim < 1 then BackBall.CreateSizedBall(20).Image = "BasketBall"  ' VRADDED  - added VRroom switch

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    'StartShake

    Dim DesktopMode: DesktopMode = Table1.ShowDT

    If DesktopMode = True Then
        EMreel1.visible=false
    Else
        EMreel1.visible=false
    end If

    ' VRADDED *******************************************************
    If RenderingMode = 2 Or LiveViewVRSim = 1 then
        set VRBGBALL=BackBall2.createball'
        'set VRBGBALL=BackBall2.createSizedballWithMass(25,0.1)  ' makes the ball spin while seated even faster :(

        'VRBGBALL.Image = "black"
        VRBGBALL.Image = "cballb"
        VRBGBALL.FrontDecal = "BallTest6"
        VRBGBALL.DecalMode = true '- switch between using the ball decal as ball logo or ball 'scratches'
        VRBGBALL.BulbIntensityScale = 0
        'VRBGBALL.Material = "playfield"
        BackBall2.enabled = false
        BackBall2.Kick 100, 2
    end if
    ' END VRADDED *******************************************************
  SetBackglass

End Sub

Sub Table1_exit()
    Controller.Pause = False
    Controller.Stop
End Sub



'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
  If KeyCode = PlungerKey Then Controller.Switch(11) = 1
  If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25

    End Select
  End If

  if keycode=StartGameKey then soundStartButton()


  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress


    If vpmKeyDown(keycode) Then Exit Sub
End Sub


Sub table1_KeyUp(ByVal Keycode)


  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
    If keycode = PlungerKey Then Controller.Switch(11) = 0

    If vpmKeyUp(keycode) Then Exit Sub

End Sub

'*********
' Switches
'********* -20 -15 -7

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    LS.VelocityCorrect(Activeball)
  LeftSling.IsDropped = 0
  RandomSoundSlingshotLeft(pBulb003) 'primitive near left sling
  vpmTimer.PulseSw 57
  LStep = 0
  Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 0:LeftSLing.IsDropped = 0:LeftSLingH.IsDropped = 0
        Case 1: 'pause
        Case 2:LeftSLing.IsDropped = 1:LeftSLing2.IsDropped = 0:LeftSLingH.IsDropped = 1:LeftSLingH2.IsDropped = 0
        Case 3:LeftSLing2.IsDropped = 1:LeftSLing3.IsDropped = 0:LeftSLingH2.IsDropped = 1:LeftSLingH3.IsDropped = 0
        Case 4:LeftSLing3.IsDropped = 1:LeftSLingH3.IsDropped = 1:Me.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    RS.VelocityCorrect(Activeball)
  RightSling.IsDropped = 0
  RandomSoundSlingshotRight(pBulb011) 'primitive near right sling
  vpmTimer.PulseSw 58
  RStep = 0
  Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 0:RightSLing.IsDropped = 0:RightSLingH.IsDropped = 0
        Case 1: 'pause
        Case 2:RightSLing.IsDropped = 1:RightSLing2.IsDropped = 0:RightSLingH.IsDropped = 1:RightSLingH2.IsDropped = 0
        Case 3:RightSLing2.IsDropped = 1:RightSLing3.IsDropped = 0:RightSLingH2.IsDropped = 1:RightSLingH3.IsDropped = 0
        Case 4:RightSLing3.IsDropped = 1:RightSLingH3.IsDropped = 1:Me.TimerEnabled = 0
    End Select

    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 61:RandomSoundBumperBottom Bumper1:bump1 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper1_Timer()
    Select Case bump1
        Case 1:bump1 = 2: BumperRing1.Z = -30
        Case 2:bump1 = 3: BumperRing1.Z = -20
        Case 3:bump1 = 4: BumperRing1.Z = -10
        Case 4:Me.TimerEnabled = 0: BumperRing1.Z = 0
    End Select

    'Bumper1R.State = ABS(Bumper1R.State - 1) 'refresh light
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 23:RandomSoundBumperTop Bumper2:bump2 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper2_Timer()
    Select Case bump2
        Case 1:bump2 = 2 : BumperRing2.Z = -30
        Case 2:bump2 = 3 : BumperRing2.Z = -20
        Case 3:bump2 = 4 : BumperRing2.Z = -10
        Case 4:Me.TimerEnabled = 0 :  : BumperRing2.Z = 0
    End Select
 '   Bumper2R.State = ABS(Bumper2R.State - 1) 'refresh light
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 62:RandomSoundBumperMiddle Bumper3:bump3 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper3_Timer()
    Select Case bump3
        Case 1:bump3 = 2: BumperRing3.Z = -30
        Case 2:bump3 = 3: BumperRing3.Z = -20
        Case 3:bump3 = 4: BumperRing3.Z = -10
        Case 4:Me.TimerEnabled = 0: BumperRing3.Z = 0
    End Select
  '  Bumper3R.State = ABS(Bumper3R.State - 1) 'refresh light
End Sub

' Eject holes
Sub Drain_Hit:RandomSoundDrain Drain:bsTrough.AddBall Me:End Sub
Sub sw25_Hit:SoundSaucerLock:bsEject.AddBall 0:End Sub

Sub sw65_Hit
    SoundSaucerLock
    bsSaucer4.AddBall 0
End Sub

Sub sw66_Hit
    SoundSaucerLock
    bsSaucer3.AddBall 0
End Sub

Sub sw67_Hit
    SoundSaucerLock
    bsSaucer2.AddBall 0
End Sub

Sub sw68_Hit
    SoundSaucerLock
    bsSaucer1.AddBall 0
End Sub

' Rollovers
Sub sw26_Hit:Controller.Switch(26) = 1:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw56_Hit:Controller.Switch(56) = 1:End Sub
Sub sw56_UnHit:Controller.Switch(56) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

'Optos
Sub sw36_Hit:Controller.Switch(36) = 1:End Sub
Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "fx_metalrolling":ActiveBall.VelY = 10:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:End Sub
Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:End Sub '75h
Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

Sub sw115_Hit:Controller.Switch(115) = 1:End Sub
Sub sw115_UnHit:Controller.Switch(115) = 0:End Sub

Sub sw117_Hit:Controller.Switch(117) = 1:End Sub
Sub sw117_UnHit:Controller.Switch(117) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub swVR12a_Hit:Controller.Switch(12) = 1:End Sub  'VRAdded
Sub swVR12b_Hit:Controller.Switch(12) = 0:End Sub  'VRAdded

Sub sw1_Hit:BasketFlipper.RotateToStart:End Sub
Sub sw2_Hit:BasketFlipper.RotateToStart:End Sub
Sub sw3_Hit:BasketFlipper.RotateToStart:End Sub

Sub bask0_hit:EMreel1.setvalue (1):End Sub
Sub bask1_hit:EMreel1.setvalue (2):End Sub
Sub bask2_hit:EMreel1.setvalue (1):End Sub
Sub bask3_hit:EMreel1.setvalue (0):End Sub


'************Diverter2 animation*************************
dim DiverterDir
Sub SolDiverter2(Enabled)
    'debug.print "diverter2 state: " & Enabled
    If Enabled Then
        Diverter2.IsDropped=1
        DiverterDir = 2
        Diverter2.timerinterval  = 5:Diverter2.timerenabled = 1
        PlaySoundAt SoundFX("DiverterOn",DOFContactors),DiverterP2
    else
        Diverter2.IsDropped=0
        DiverterDir = -2
        Diverter2.timerinterval  = 5:Diverter2.timerenabled = 1
        PlaySoundAt SoundFX("DiverterOff",DOFContactors),DiverterP2
    End If
End Sub

Sub Diverter2_Timer()
    DiverterP2.transY=DiverterP2.transY+DiverterDir
    If DiverterP2.transY>30 AND DiverterDir=2 Then Me.timerenabled=0:DiverterP2.transY=30
    If DiverterP2.transY<0 AND DiverterDir=-2 Then Me.timerenabled=0:DiverterP2.transY=0
End Sub

'**************StandUp targets****************************


Sub sw41_Hit
  STHit 41
End Sub

Sub sw42_Hit
  STHit 42
End Sub

Sub sw43_Hit
  STHit 43
End Sub

Sub sw28_Hit
  STHit 28
End Sub

Sub sw18_Hit
  STHit 18
End Sub

'***********
' Solenoids
'***********

SolCallBack(1) = "Auto_Plunger"
'SolCallBack(2) = Not Used
'SolCallBack(3) = "vpmSolWall Diverter2,True,"
SolCallBack(3) = "SolDiverter2"
SolCallBack(4) = "vpmSolWall Diverter1,True,"
SolCallBack(5) = "bsEject.SolOut"
SolCallBack(6) = "RightGate.Open ="
SolCallBack(7) = "SolBasket"
'SolCallBack(8) ' magnet - handled in the magnet definition
SolCallBack(9) = "bsTrough.SolOut"
'SolCallBack(10)  = "vpmSolSound ""lSling"","
'SolCallBack(11)  = "vpmSolSound ""lSling"","
'SolCallBack(12)  = "vpmSolSound ""Jet1"","
'SolCallBack(13)  = "vpmSolSound ""Jet1"","
'SolCallBack(14)  = "vpmSolSound ""Jet1"","

SolCallBack(15) = "PassRight2"
SolCallBack(16) = "PassLeft2"

SolModCallBack(17) = "Flasher17"    'eject kickout flasher
SolModCallBack(18) = "Flasher18"    'left jet bumper
SolModCallBack(19) = "Flasher19"    'upper left         'BG Left
SolModCallBack(20) = "Flasher20"    'upper right        'BG Right
SolModCallBack(22) = "Flasher22"    'trophy insert
SolModCallBack(24) = "Flasher24"    'lower right left

SolCallBack(25) = "PassRight1"
SolCallBack(26) = "PassLeft3"
SolCallBack(27) = "PassRight3"
SolCallBack(28) = "PassLeft4"

SolCallBack(33) = "bsSaucer1.SolOut"
SolCallBack(34) = "bsSaucer2.SolOut"
SolCallBack(35) = "bsSaucer3.SolOut"
SolCallBack(36) = "bsSaucer4.SolOut"
'SolCallBack(37)  = Motor Enable (defender) - handled in the mech
'SolCallBack(38)  = Motor Direction (defender) - handled in the mech
'SolCallBack(39) = "ClockEnable"
'SolCallBack(40) = "ClockCount"



' Flashers
'**********

const DebugFlashers = false

Sub Flasher17(pwm)
  If DebugFlashers then debug.print "Flasher17 "&pwm
  F17.State = pwm
End Sub

Sub Flasher18(pwm)
  If DebugFlashers then debug.print "Flasher18 "&pwm
  F18.State = pwm
End Sub

Sub Flasher19(pwm)
  If DebugFlashers then debug.print "Flasher19 "&pwm
  F19.State = pwm
End Sub

Sub Flasher20(pwm)
  If DebugFlashers then debug.print "Flasher20 "&pwm
  F20.State = pwm
End Sub

Sub Flasher22(pwm)
  If DebugFlashers then debug.print "Flasher22 "&pwm
  F22.State = pwm
End Sub

Sub Flasher24(pwm)
  If DebugFlashers then debug.print "Flasher24 "&pwm
  F24a.State = pwm
  F24b.State = pwm
End Sub

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Dim VRBGBALL  'VRADDED
FlipperBumper.isdropped = true  'VRADDED - start dropped


Sub SolBasket(Enabled)
    If Enabled Then

        If RenderingMode <> 2 And LiveViewVRSim < 1 then ' VRADDED - switch
      BackBall.Kick 330, 75
      If bgball=1 then BasketFlipper.RotateToEnd
        bgball=0
        'end if

        'VRADDED **************************************************************************************************************
    else

    'VRBasketFlipper.RotateToEnd
    FlipperBumper.isdropped = false
    FlipperReturn.enabled = true

    ' Backglass FlipperPhysics
    If VRBGBALL.z < 1020 then  ' only do anything if the ball is lower than the top of the flipper.


      if VRBGBALL.X => 790 and VRBGBALL.X < 810 then
        VRBGBALL.velZ = 30 ' lower trajectory and force if hit too soon..
        VRBGBALL.velx = -10
      end If

  if VRBGBALL.X => 810 and VRBGBALL.X < 830 then
  VRBGBALL.velZ = 35 ' lower trajectory and force if hit too soon..
  VRBGBALL.velx = -10
  end If

  if VRBGBALL.X => 830 and VRBGBALL.X < 850 then
  VRBGBALL.velZ = Int(Rnd*3)+38
  end If

  if VRBGBALL.X => 850 and VRBGBALL.X < 870 then
  VRBGBALL.velZ = Int(Rnd*3)+43 ' lower trajectory and force if hit too soon..
  end If

  if VRBGBALL.X => 870 and VRBGBALL.X < 900 then
  VRBGBALL.velZ = Int(Rnd*6)+48
  end If

  if VRBGBALL.X => 900 then
  'VRBGBALL.velZ = 55
  VRBGBALL.velZ = Int(Rnd*15)+55 ' random between 55 and 70, just for different timed bounce returns

  end If

end if
End If
end if
End Sub

Dim FlipReturn: FlipReturn = 0
Sub FlipperReturn_Timer()
If FlipReturn = 0 then
VRBasketFlipper.RotateToEnd
FlipReturn =1
FlipperReturn.interval = 80
Else
VRBasketFlipper.RotateToStart
FlipReturn = 0
FlipperReturn.interval = 10
FlipperBumper.isdropped = true  ' Drop it after time..
FlipperReturn.enabled = false
end If
' END VRADDED ***********************************************************


End Sub

Sub BackBall1_Hit
    BackBall1.Destroyball
  BallReturn1.CreateSizedBall(20).Image = "BasketBall"
  BallReturn1.Kick 100, 15
End Sub
Sub BallReturn2_Hit
  BallReturn2.Destroyball
  BackBall.CreateSizedBall(20).Image = "BasketBall"
  bgball=1
End Sub

Sub UpdateDefender(aNewPos, aSpeed, aLastPos)
    ADef1(aLastPos).IsDropped = True
    ADef1(aNewPos).IsDropped = False
    ADef2(aLastPos).visible = 0:ADef2(aLastPos).collidable = 0
    ADef2(aNewPos).visible = 1:ADef2(aNewPos).collidable = 1
   'DefR.State = ABS(DefR.State -1)
End Sub

' In The Paint

Sub PassRight1(Enabled)
    If Enabled then
        bsSaucer1.Kickz = 0
        bsSaucer1.InitAltKick 110, 10
        bsSaucer1.ExitAltSol_On
        bsSaucer1.Kickz = 1.2
    End If
End Sub

Sub PassRight2(Enabled)
    If Enabled then
        bsSaucer2.Kickz = 0
        bsSaucer2.InitAltKick 70, 12
        bsSaucer2.ExitAltSol_On
        bsSaucer2.Kickz = 1.15
    End If
End Sub

Sub PassRight3(Enabled)
    If Enabled then
        bsSaucer3.Kickz = 0
        bsSaucer3.InitAltKick 45, 13
        bsSaucer3.ExitAltSol_On
        bsSaucer3.Kickz = 1.15
    End If
End Sub

Sub PassLeft2(Enabled)
    If Enabled then
        bsSaucer2.Kickz = 0
        bsSaucer2.InitAltKick 315, 13
        bsSaucer2.ExitAltSol_On
        bsSaucer2.Kickz = 1.15
    End If
End Sub

Sub PassLeft3(Enabled)
    If Enabled then
        bsSaucer3.Kickz = 0
        bsSaucer3.InitAltKick 295, 14
        bsSaucer3.ExitAltSol_On
        bsSaucer3.Kickz = 1.15
    End If
End Sub

Sub PassLeft4(Enabled)
    If Enabled then
        bsSaucer4.Kickz = 0
        bsSaucer4.InitAltKick 250, 10
        bsSaucer4.ExitAltSol_On
        bsSaucer4.Kickz = 1.2
    End If
End Sub


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

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

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub


'******************
' RealTime Updates
'******************
Dim FrameTime, InitFrameTime
InitFrameTime = 0


Sub FrameTimer_Timer
    FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime

    RollingSound
  UpdateLED
  UpdateShotClock
  DoSTAnim
  LFLogo.RotZ = LeftFlipper.CurrentAngle
  RFlogo.RotZ = RightFlipper.CurrentAngle
  LeftFlipperSh.RotZ = LeftFlipper.currentangle
  RightFlipperSh.RotZ = RightFlipper.currentangle
  'slacktimer

    batleft.objroty = VRBasketFlipper.CurrentAngle + 1   'VRADDED Backglass flipper follow
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


'********************
' Diverse Help/Sounds
'********************

'right ramp sounds
Sub RRHit0_Hit:PlaySoundAtBall "fx_rr2":End Sub
Sub RRHit1_Hit:PlaySoundAtBall "fx_rr5":End Sub
Sub RRHit2_Hit:PlaySoundAtBall "fx_rr6":End Sub

Sub WireEnter1_Hit:PlaySound "fx_metalrolling":End Sub
Sub WireExit1_Hit:StopSound "fx_metalrolling":PlaySound "fx_wirerampexit":End Sub

Sub RHelp1_Hit:PlaySound "fx_ballhit":End Sub
Sub RHelp2_Hit:PlaySound "fx_ballhit":End Sub
Sub RHelp5_Hit:PlaySound "fx_ballhit":End Sub
Sub RHelp3_Hit
    If ActiveBall.VelY < -10 Then
        ActiveBall.VelY = -10
    End If
End Sub

Sub RHelp4_Hit
    ActiveBall.VelX = -5
End Sub

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
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, swPlunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, swPlunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, swPlunger
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


'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************


'**************************************************************************************************************
' Random Ramp Bumps  - Best used to compliment the Raised Ramp RollingBall Script
' No switches are required. Make ramps have "Hit Event" and create a "Collection" for each of the following.
' Adjust Vol (eg .5) & Pitch:     RandomBump3 .5, Pitch(ActiveBall)+5
' Ramp physics also effect bump sounds. Friction will change ball speed, so sounds will react different.
' Hit Threshold will will effect how much force is required to create a bump. Low hit threshold + more bumps.
' Other adjustments avialable in the script as noted.
'**************************************************************************************************************


Dim NextOrbitHit:NextOrbitHit = 0

Sub WireRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump3 .5, Pitch(ActiveBall)+5
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub

Sub PlasticRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 2, Pitch(ActiveBall)
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



' Stop Bump Sounds - Place triggers for these at the end of the ramps, to ensure that no bump sounds are played after the ball leaves the ramp.

Sub BumpSTOP1_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOP2_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub



'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 7 ' total number of balls
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

Sub RollingSound()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("BallRoll_" & b)
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 Then
      rolling(b) = True
      if BOT(b).z < 10 Then 'Ball on playfield
        PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
      Else 'Ball on Raised Ramp
        PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 10 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b))+50000, 1, 0, AudioFade(BOT(b))
      End If

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
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

    ' "Static" Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If BOT(b).Z > 30 Then
        BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
      Else
        BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
      End If
      BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
      BallShadowA(b).X = BOT(b).X
      BallShadowA(b).visible = 1
    End If
  Next
End Sub


' ===============================================================================================
' LED's display
' ===============================================================================================
Dim Digits(2), digitState

Digits(0)   = Array(led0,led1,led2,led3,led4,led5,led6)
Digits(1)   = Array(led7,led8,led9,led10,led11,led12,led13)

Sub UpdateLED()
    Dim chgLED, ii, iii, num, chg, stat, digit
    chgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(chgLED) Then
        For ii = 0 To UBound(chgLED)
            num  = chgLED(ii, 0)
            chg  = chgLED(ii, 1)
            stat = chgLED(ii, 2)
            iii  = 0
            For Each digit In Digits(num)
                If (chg And 1) Then digit.State = (stat And 1)
                chg  = chg \ 2
                stat = stat \ 2
            Next
        Next
    End If
End Sub

Sub UpdateShotClock

  If RedClock < 1 Then
    SCleft0r.visible =false
    SCleft1r.visible =false
    SCleft2r.visible =false
    SCleft3r.visible =false
    SCleft4r.visible =false
    SCleft5r.visible =false
    SCleft6r.visible =false
    SCright0r.visible =false
    SCRight1r.visible =false
    SCRight2r.visible =false
    SCRight3r.visible =false
    SCRight4r.visible =false
    SCRight5r.visible =false
    SCRight6r.visible =false
    If LED0.State=1 then SCleft0.visible =true else SCleft0.visible =false
    If LED1.State=1 then SCleft1.visible =true else SCleft1.visible =false
    If LED2.State=1 then SCleft2.visible =true else SCleft2.visible =false
    If LED3.State=1 then SCleft3.visible =true else SCleft3.visible =false
    If LED4.State=1 then SCleft4.visible =true else SCleft4.visible =false
    If LED5.State=1 then SCleft5.visible =true else SCleft5.visible =false
    If LED6.State=1 then SCleft6.visible =true else SCleft6.visible =false

    If LED7.State=1 then SCright0.visible =true else SCright0.visible =false
    If LED8.State=1 then SCRight1.visible =true else SCRight1.visible =false
    If LED9.State=1 then SCRight2.visible =true else SCRight2.visible =false
    If LED10.State=1 then SCRight3.visible =true else SCRight3.visible =false
    If LED11.State=1 then SCRight4.visible =true else SCRight4.visible =false
    If LED12.State=1 then SCRight5.visible =true else SCRight5.visible =false
    If LED13.State=1 then SCRight6.visible =true else SCRight6.visible =false

    If l87.State=1 and Not RenderingMode = 2 And LiveViewVRSim < 1 and ShootShooter = 1 then Shoot_Flash.visible = True else Shoot_Flash.visible =false
    If l87.State=1 and RenderingMode = 2 Or LiveViewVRSim = 1 then VR_Shoot_Flash.visible = True else VR_Shoot_Flash.visible =false

    If l88.State=1 and RenderingMode = 2 Or LiveViewVRSim = 1 then VR_Start.visible =true else VR_Start.visible =false
  Else
    SCleft0.visible =false
    SCleft1.visible =false
    SCleft2.visible =false
    SCleft3.visible =false
    SCleft4.visible =false
    SCleft5.visible =false
    SCleft6.visible =false
    SCright0.visible =false
    SCRight1.visible =false
    SCRight2.visible =false
    SCRight3.visible =false
    SCRight4.visible =false
    SCRight5.visible =false
    SCRight6.visible =false

    If LED0.State=1 then SCleft0r.visible =true else SCleft0r.visible =false
    If LED1.State=1 then SCleft1r.visible =true else SCleft1r.visible =false
    If LED2.State=1 then SCleft2r.visible =true else SCleft2r.visible =false
    If LED3.State=1 then SCleft3r.visible =true else SCleft3r.visible =false
    If LED4.State=1 then SCleft4r.visible =true else SCleft4r.visible =false
    If LED5.State=1 then SCleft5r.visible =true else SCleft5r.visible =false
    If LED6.State=1 then SCleft6r.visible =true else SCleft6r.visible =false

    If LED7.State=1 then SCright0r.visible =true else SCright0r.visible =false
    If LED8.State=1 then SCRight1r.visible =true else SCRight1r.visible =false
    If LED9.State=1 then SCRight2r.visible =true else SCRight2r.visible =false
    If LED10.State=1 then SCRight3r.visible =true else SCRight3r.visible =false
    If LED11.State=1 then SCRight4r.visible =true else SCRight4r.visible =false
    If LED12.State=1 then SCRight5r.visible =true else SCRight5r.visible =false
    If LED13.State=1 then SCRight6r.visible =true else SCRight6r.visible =false

    If l87.State=1 and Not RenderingMode = 2 And LiveViewVRSim < 1 and ShootShooter = 1 then Shoot_Flash.visible = True else Shoot_Flash.visible =false
    If l87.State=1 and RenderingMode = 2 Or LiveViewVRSim = 1 then VR_Shoot_Flash.visible = True else VR_Shoot_Flash.visible =false

    If l88.State=1 and RenderingMode = 2 Or LiveViewVRSim = 1 then VR_Start.visible =true else VR_Start.visible =false
  End If
End Sub


'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

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
    x.AddPt "Polarity", 1, 0.05, - 5.5
    x.AddPt "Polarity", 2, 0.16, - 5.5
    x.AddPt "Polarity", 3, 0.20, - 0.75
    x.AddPt "Polarity", 4, 0.25, - 1.25
    x.AddPt "Polarity", 5, 0.3, - 1.75
    x.AddPt "Polarity", 6, 0.4, - 3.5
    x.AddPt "Polarity", 7, 0.5, - 5.25
    x.AddPt "Polarity", 8, 0.7, - 4.0
    x.AddPt "Polarity", 9, 0.75, - 3.5
    x.AddPt "Polarity", 10, 0.8, - 3.0
    x.AddPt "Polarity", 11, 0.85, - 2.5
    x.AddPt "Polarity", 12, 0.9, - 2.0
    x.AddPt "Polarity", 13, 0.95, - 1.5
    x.AddPt "Polarity", 14, 1, - 1.0
    x.AddPt "Polarity", 15, 1.05, -0.5
    x.AddPt "Polarity", 16, 1.1, 0
    x.AddPt "Polarity", 17, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.23, 0.85
    x.AddPt "Velocity", 2, 0.27, 1
    x.AddPt "Velocity", 3, 0.3, 1
    x.AddPt "Velocity", 4, 0.35, 1
    x.AddPt "Velocity", 5, 0.6, 1 '0.982
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
  Dim gBOT
  gBOT = GetBalls

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

Dim LFPress, RFPress, LFCount, RFCount
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
'Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
'   Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b, gBOT
    gBOT = GetBalls

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

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
' Cor.Update
'End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9  'Level of bounces. Recommmended value of 0.7-1

Sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel, vratio
  If TargetBouncerEnabled = 1 And aball.z < 30 Then
    '   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    vel = BallSpeed(aBall)
    If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
    Select Case Int(Rnd * 6) + 1
      Case 1
        zMultiplier = 0.2 * defvalue
      Case 2
        zMultiplier = 0.25 * defvalue
      Case 3
        zMultiplier = 0.3 * defvalue
      Case 4
        zMultiplier = 0.4 * defvalue
      Case 5
        zMultiplier = 0.45 * defvalue
      Case 6
        zMultiplier = 0.5 * defvalue
    End Select
    aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
    aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
    aBall.vely = aBall.velx * vratio
    '   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    '   debug.print "conservation check: " & BallSpeed(aBall)/vel
  End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1
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
' ZRST: STAND-UP TARGET INITIALIZATION
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
Dim ST18, ST28, ST41, ST42, ST43

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

Set ST18 = (new StandupTarget)(sw18, psw18, 18, 0)
Set ST28 = (new StandupTarget)(sw28, psw28, 28, 0)
Set ST41 = (new StandupTarget)(sw41, psw41, 41, 0)
Set ST42 = (new StandupTarget)(sw42, psw42, 42, 0)
Set ST43 = (new StandupTarget)(sw43, psw43, 43, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST18, ST28, ST41, ST42, ST43)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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
    prim.transx =  - STMaxOffset
    vpmTimer.PulseSw switch mod 100
    STAnimate = 2
    Exit Function
  ElseIf animate = 2 Then
    prim.transx = prim.transx + STAnimStep
    If prim.transx >= 0 Then
      prim.transx = 0
      primary.collidable = 1
      STAnimate = 0
      Exit Function
    Else
      STAnimate = 2
    End If
  End If
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


'******************************************************
'***  END STAND-UP TARGETS
'******************************************************



'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx7" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
' * with at least as many objects each as there can be balls, including locked balls
' Ensure you have a timer with a -1 interval that is always running

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
'***These must be organized in order, so that lights that intersect on the table are adjacent in the collection***
'
' This is because the code will only project two shadows if they are coming from lights that are consecutive in the collection
' The easiest way to keep track of this is to start with the group on the left slingshot and move clockwise around the table
' For example, if you use 6 lights: A & B on the left slingshot and C & D on the right, with E near A&B and F next to C&D, your collection would look like EBACDF
'
'                               E
' A    C                          B
'  B    D     your collection should look like    A   because E&B, B&A, etc. intersect; but B&D or E&F do not
'  E      F                         C
'                               D
'                               F
'
'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Timer sub
' ' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
' Sub FrameTimer_Timer()
'   If DynamicBallShadowsOn Or AmbientBallShadowOn Then
'     DynamicBSUpdate 'update ball shadows
'   Else
'     me.Enabled = False 'Remove this if you add anything else to the timer!
'   End If
' End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
Const lob = 0 'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzo's)
'                 '2 = flasher image shadow, but it moves like ninuzzo's
Const fovY          = 0   'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source


' *** This segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' ' stop the sound of deleted balls
' For b = UBound(BOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
'   rolling(b) = False
'   StopSound("BallRoll_" & b)
' Next
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If BOT(b).Z > 30 Then
'       BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
'     Else
'       BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
'     End If
'     BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
'     BallShadowA(b).X = BOT(b).X
'     BallShadowA(b).visible = 1
'   End If

' *** Required Functions, enable these if they are not already present elswhere in your table
Function DistanceFast(x, y)
  dim ratio, ax, ay
  ax = abs(x)         'Get absolute value of each vector
  ay = abs(y)
  ratio = 1 / max(ax, ay)   'Create a ratio
  ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
  if ratio > 0 then     'Quickly determine if it's worth using
    DistanceFast = 1/ratio
  Else
    DistanceFast = 0
  End if
end Function


'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******
Dim sourcenames, currentShadowCount
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(12), objrtx2(12)
dim objBallShadow(12)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9,BallShadowA10,BallShadowA11)

DynamicBSInit

sub DynamicBSInit()
  Dim iii

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 0.01
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 0.02
    objrtx2(iii).visible = 0

    currentShadowCount(iii) = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = iii/1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next
end sub


Sub DynamicBSUpdate
  Dim falloff:  falloff = 150     'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, currentMat, AnotherSource, BOT
  BOT = GetBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 to tnob
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub    'No balls in play, exit

'The Magic happens now
  For s = lob to UBound(BOT)

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If BOT(s).Z > 30 Then             'The flasher follows the ball up ramps while the primitive is on the pf
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + BallSize/10 + fovY
        objBallShadow(s).visible = 1

        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf, primitive only
        objBallShadow(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + fovY
        BallShadowA(s).visible = 0
      Else                      'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If BOT(s).Z > 30 Then             'In a ramp
        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          BallShadowA(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        BallShadowA(s).Y = BOT(s).Y + Ballsize/10 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/2 + 5
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 Then 'And BOT(s).Y < (TableHeight - 200) Then 'Or BOT(s).Z > 105 Then    'Defining when and where (on the table) you can have dynamic shadows
        For Each Source in DynamicSources
          LSd=DistanceFast((BOT(s).x-Source.x),(BOT(s).y-Source.y)) 'Calculating the Linear distance to the Source
          If LSd < falloff and Source.state=1 Then            'If the ball is within the falloff range of a light and light is on
            currentShadowCount(s) = currentShadowCount(s) + 1   'Within range of 1 or 2
            if currentShadowCount(s) = 1 Then           '1 dynamic shadow source
              sourcenames(s) = source.name
              currentMat = objrtx1(s).material
              objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01            'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(Source.x, Source.y, BOT(s).X, BOT(s).Y) + 90
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
              set AnotherSource = Eval(sourcenames(s))
              objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(AnotherSource.x, AnotherSource.y, BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-(((BOT(s).x-AnotherSource.x)^2+(BOT(s).y-AnotherSource.y)^2)^0.5))/falloff
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

              currentMat = objrtx2(s).material
              objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
  '           objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx2(s).rotz = AnglePP(Source.x, Source.y, BOT(s).X, BOT(s).Y) + 90
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







'******************************************************
'*******  Set Up Backglass Flashers *******
'******************************************************

Sub SetBackglass()
  Dim obj

  For Each obj In VRBackglassFront
    obj.x = obj.x + 20
    obj.height = - obj.y + 360
    obj.y = -50 'adjusts the distance from the backglass towards the user
    obj.rotx=-89
  Next

  For Each obj In VRBackglassBack
    obj.x = obj.x - 5
    obj.height = - obj.y + 360
    obj.y = -130 'adjusts the distance from the backglass towards the user
    obj.rotx=-87
  Next

  For Each obj In VRBackglassBulbTop
    obj.x = obj.x - 5
    obj.height = - obj.y + 360
    obj.y = -110 'adjusts the distance from the backglass towards the user
    obj.rotx=-87
  Next

  For Each obj In VRBackglassBulbMidTop
    obj.x = obj.x - 5
    obj.height = - obj.y + 360
    obj.y = -100 'adjusts the distance from the backglass towards the user
    obj.rotx=-87
  Next


  For Each obj In VRBackglassBulbMidBottom
    obj.x = obj.x - 5
    obj.height = - obj.y + 360
    obj.y = -90 'adjusts the distance from the backglass towards the user
    obj.rotx=-87
  Next

  For Each obj In VRBackglassBulbBottom
    obj.x = obj.x - 5
    obj.height = - obj.y + 360
    obj.y = -55 'adjusts the distance from the backglass towards the user
    obj.rotx=-87
  Next

End Sub


' *************************  script for LUT text display - Flupper1 *********************************
' script related to the script below is in keydown magnasave keys script and luts/lutpos script at the top of the script
' other objects used: layer 7 all Text0?? objects ("A") and the textures "32" to "96"

dim rollingtext
dim textindex : textindex = 1
dim charobj(55), glyph(201)
InitDisplayText

Sub text010_Timer()
  dim tekst
  tekst = Mid(rollingtext, textindex, 34)
  DisplayText -1, tekst
  textindex = textindex + 1
  If textindex > len(rollingtext) - 36 then textindex = 1 : end if
End Sub

Sub myChangeLut
    Table1.ColorGradeImage = "LUT" & lutpos
  DisplayText lutpos, luts(lutpos)
  vpmTimer.AddTimer 2000, "If lutpos = " & lutpos & " then for anr = 10 to 54 : charobj(anr).visible = 0 : next'"
End Sub

Sub ResetRollingTextTimer()
  Dim anr
  text010.TimerEnabled = False : textindex = 1
  For anr = 10 to 43 : charobj(anr).opacity = 4000 : charobj(anr).y = 2100 : charobj(anr).height = 150 : charobj(anr).visible = 0 : next
End Sub

Sub InitDisplayText
  Dim anr
  For anr = 10 to 54 : set charobj(anr) = eval("text0" & anr) : charobj(anr).visible = 0 : Next
  For anr = 32 to 96 : glyph(anr) = anr : next
  For anr = 0 to 31 : glyph(anr) = 32 : next
  for anr = 97 to 122 : glyph(anr)  = anr - 32 : next
  for anr = 123 to 200 : glyph(anr) = 32 : next
End Sub

Sub DisplayText(nr, luttext)
  dim tekst, anr
  for anr = 10 to 54 : charobj(anr).imageA = 32 : charobj(anr).visible = 1 : next
  If nr > -1 then
    tekst = "lutpos:" & nr
    For anr = 1 to len(tekst) : charobj(43 + anr).imageA = glyph(asc(mid(tekst, anr, 1))) : Next
  End If
  For anr = 1 to len(luttext)
    charobj(9 + anr).imageA = glyph(asc(mid(luttext, anr, 1)))
    If nr = -1 Then
      charobj(9 + anr).y = 1500 + sin(((textindex * 4 + anr)/20)*3.14) * 100
      charobj(9 + anr).height = 150 + cos(((textindex * 4 + anr)/20)*3.14) * 100
    End If
  Next
End Sub

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


'**CHANGE LOG **
' 005 - iaakki - PF and insert images, first inserts added. All lights moved to different layers
' 006 - iaakki - PF material fixed. Adding more inserts, right ramp invisible for now
' 007 - tomate - All the new geometry inside. Layer 1:Non-colidables primitives Layer 2:Colidables LowPoly primitives Layer 3:original VPX objetcs aligned. Added new rubber stuff.
' 008 - iaakki - fixed missing wall diverters
' 009 - tomate - Fixed issue with bumpers, remove collidable wall in position of diverter1, fixed plastics texture problems, added a collidable wall on shot clock, fixed geometry of the RRamp
' 010 - tomate - Added some new geometries, aligned rubbers and addded new POV
' 011 - iaakki - combined inserts and graphics from separate version
' 012 - tomate - New textures
' 013 - tomate - Improve plastic textures, shadows on PF, new apron
' 014 - tomate - More work on textures
' 015 - tomate - Work on file weight, fixed wall problem in player #3, new apron texture, some additional fixes
' 016 - Benji - Added nfozzy physics to flippers, rubber dampening is in script but not assigned to any objects yet.Physics material assigned to collidable ramp primitives
' 017 - Benji - Added darkening/contrast LUT and DL to ramps_laterl primitive (can be disabled in options below). Added collidable rubber sleeves primitives on layer 10 and assigned physics material and nfozzy dampening event
' 018 - tomate - New lighting and color correction with LUTs, new PF with color correction, shadow layer added
' 019 - tomate - Diverter added, fix on collideable walls to prevent ball from getting stuck, new flippers added, triggers fixed, some minor adjustments in textures, clock numbers fixed
' 020 - Benji - Added physics to rubber bands
' 021 - iaakki - flupper domes partially done
' 022 - tomate - Added PF mesh, values of sides passes fixed, force value for sw25 fixed, color correction to PF
' 023 - iaakki - flasher domes done
' 024 - iaakki - GI redone and connected, PF GI missing from strings 2 and 3, gInsertLevel const added to reduce insert brightness globally
' 025 - tomate - Clean script flippers williams logo, blenddisablelighting to Ramp_central added, new water level added, new transparent textures for plastic ramps added
' 026 - tomate - fix the height of the bulbs under Jam & Slam ramps, new siderails and sidewalls added, GlobalFlasher added on PF at layer10
' 027 - Benji - Added Fleep sound package and DJRobX wire/plastic/metal ramp bumps.
' 030 - iaakki - included flipper logos. Flipnudge is having some issues still
' 031 - Sixtoe - Added VR Room, fixed flipper shadows, trimmed lights, unified timmers (including removing redundant timers), removed legacy nudging system, fixed numerous VR depth bias issues, replaced flasher for red left turbo bumper flasher and hooked it up to the flasher system
' 032 - Benji - Removed extra glosiness from new flippers by creating new material for flippers. Got new Flipper Nudge functionality working
' 033 - iaakki - Diverter2, Flipper_collide sub, FlipperSlack&Checklivecatch calls fixed and slackyFlips const added -> Flip tricks works
' 034 - tomate - fixed kicker sw66 (player # 3) shooting power, adjustments on collideable walls to prevent ball jamming, adjust DL of new bats
' 035 - tomate - Primitives of collidable Ramps fixed, usless old collidable objects removed
' 036 - sixtoe/tomate - Shadow ball added, gamma correction to ramps decals
' 037.1 - Sixtoe - Updated VR cabinet, corrected ramp materials for VR, added left ramp lid, aligned 3rd basket flasher, corrected blooms for all basket flashers, added new bloom for flasher10
' 037.2 - Sixtoe - removed LowPoly_Ramps001 collision primitive as was a duplicate, hid everything visible outside of table limits, created apron "shoot" button and light, corrected apron prim so UV isn't broken, repositioned all slings and sling rubbers, fixed right sling
' 037.3 - Sixtoe - added cabinet start light, hooked up existing off cabninet start and shoot lights for desktop mode?, hooked GlobalFlasher to GI system, created sidewall "off" images and hooked up the sidewalls to the GI system for both the normal and VR version (crowd currently not implemented)
' 038 - iaakki - fixed POV and some insert light bulb values
' 039 - tomate - fixed upper right rubber, new texture for the ball, POV adjusted to show the full apron
' 040 - tomate - adjustments to apron and metals textures, tweaks DL on some objects, change ramp materials to look brighter
' 041 - iaakki - top area pf gi redone
' 042 - iaakki - Some plastic GI light heights retuned
' 043 - Sixtoe - changed lamp decal materials so that they work properly, set back to active so they're still see through, fixed some lights, fixed the start button to not show in vr, changed and adjusted the apron walls to show in VR, changed the material of the defender as it was broken in vr, hooked up rollvers to new sound code, removed old sound code and la1-6 walls, droppped playfield hole surrounds, tidied up playfield holes and made a new one for a sw43 target,
' 044 - tomate/iaakki - low res for some textures, StandUp targets separation into 3 prims and animation done, modify backwall texture
' 045 - tomate/oqq - correct apron wall height plastic under central ramp texture, add movments to bumperRings and separate prims
' 046 - tomate - add cabinet mode (working half), removed white things under the ramps, added subtle transparency to the ramp decals, so the lights below can be seen a bit, new yellow targets prims and textures
' 046.1 - tomate - yellow targets working
' 047 - Sixtoe - VR room settings fixed and tweaked, ultra minimal vr added, fixed desktop mode, fixed ball jumping into plunger lane by adding new cover wall, fixed ramp primitive have a hole in it, removed crowd sides for now, fixed desktop pov, turned flipper strength way down
' 048 - iaakki - Desktop POV adjusted, standup target bounces reduced, pf friction increased, Flasherbase1 height fixed
' 049 - Sixtoe - Fixed low poly ramp stopping ball going up left ramp, then fixed it getting stuck there!, adjusted right ramp roof again, added new wall on left by food kicker to stop ball trap, repositioned basket loop to fix POV and make sounds appear near top of table, added missing rubbers to outlane pegs, fixed blade/sidewall texture collision in cabinetmode, added wall on slam ramp to stop ball falling off wire ramp back onto slam ramp
' 050 - iaakki - Standup targets elasticity reduced. Sw64 surface fixed to 55h, l57 brightness reduced, insert paint layer color adjustment
' 051 - tomate - fix Pincab_rails mesh, change material settings and mapped. Fix POV's, now is working in FS and DT mode.
' 052 - iaakki - fine tuned flip positions, fixed lamp update functions to reset properly for all lamps.
' 053 - iaakki - sw46 and sw63 surface fixed, diverter sounds added
' 054 - tomate - Change size and textures of SideBlades, add roof to left plastic, raised the top of the LeftRamp roof to prevent ball hitting, add thickness to PF and mapping
' 054.1 - tomate - Restored the old roof of the left ramp as the new one didn´t work
' 055 - iaakki- ScSp reflections on by default, cabinetmode raises flasherblooms to 500, flasherlights adjusted and lifted to same level as the plastic.
' 056 - Sixtoe - extended collidable wall around the orbit loop to remove ball traps on the top corners, added code to adjust flashbloom height and Y position to cabinet mode, adjusted wire ramp disable lighting to stop lights coming through it in VR, removed rundant ring1/2/3 walls, switched metal post to plastic post under white flasher, adjusted switch 64 height, added bumper1 top wall,
' RC1 - tomate - make not visible wall over main bumper, set 100% sound effect Volume, set CabinetMode = 0
' RC2 - Sixtoe - Adjusted slam ramp, added kicker primitives to basket shots, added additional wall on slingshot back, tweaked a handful of other things
' RC3 - Sixtoe - Adjusted flippers, removed sling added in error, changed height of wireexit1
' RC4 - iaakki - Round standup target fixed to feel more natural, redclock script error fixed
' RC5 - tomate - red clock numbers added
' RC6 - tomate - redclocktimer enabled, add kicker primitive to middle-left hole, some tweaks in decal_ramps textures and Left ramp texture, new textures for laneguides
' RC7 - tomate - increase Cabflash.height value to 520, change POV to match cabinet view, add a wall under cnetral ramp to prevent ball jams, separate the primitives from the laneguides and made them non-collideable, add wall on the laneguides to avoid sorcery
' 1.00 - iaakki - flasherbloom default heights checked, ball counts fixed, default options set
' 1.01 - Sixtoe - Added option to disable flasherblooms', added wall under slam loop to stop trapped balls, changed colour of top right flasher to white, adjusted flasherlight's heights and turned lights down now they're both working.
' 1.02 - tomate - Corrected size of the upper domes, position of the central left dome slightly modified, added wall in the upper right plastic and sloping walls in the central ramp sector to prevent jamming, new sideblades ON/OFF textures for cabinet mode
' 1.03 - tomate - adds wall behind the basket to avoid jamming when the ball goes slow on the left ramp
' 1.04 - tomate - Adjusted the start angle of the flippers to 120/-120 and reduced the strength to 3200, lowered opacity of ShootbuttonOff and StartButtonOn so they are not visible in DT mode
' 1.05 - benji - adjusted nfozzy physics, sorted out rubbers added right physics materials ect. Sling sounds till broken. See line commmented out near line 600.
' 1.06 - iaakki - adjust rubbers, default to desktop mode, adjust flasher fade speeds, hide emreel basket in desktopmode, adjust top flasher positions, fix flasher lit sizes, reverted flip strength back to 3250
' 1.07 - Benji - nFozzy physics 'one more time'. Properly separated rubber post meshes from rubberband meshes in 3D application, re-imported non-visible collidable rubber bands + posts assigned to dPosts and rubber bands physics materials
' 1.08 - Sixtoe - Added shoot button option for it to appear in shooter / plunger lane.
' 1.09 - Sixtoe - Went over the flashers and fixed numerous issues, fixed a couple of normal lights and adjusted a few things. Hooked up slacktimer to global timer.
' 1.10 - benji - Divided rubber bands into third segment (center 'stretched' segment'), imported this new invisible collidable mesh, and applied new physics material to it
' 1.11 - iaakki - updated some NF and scripts and parameters that were old, moved triggers sw115 and sw117 slightly
' 1.12 - iaakki - Laneguides and pegs tied to GI. Laneguide material changed.
' 1.13 - iaakki - Minor improvement to lighting dynamics. N/D slider 5 -> 4 and gInsertLevel 0.5 -> 0.9
' 1.14 - tomate - make the shoot button a little smaller to fit better
' 1.15 - benji - reduced polygons on rubber bands collidable meshes
' 1.16 - iaakki - slingshot adjust
' 1.17 - iaakki - GI shape and insert adjust
' 1.18 - iaakki - Fixed one more ball stuck behind basket (wall14, sw117 reshape), "in paint" inserts dimmed a bit, improved basket front edge to work more like in real cabinet.
' 1.19 - tomate - Dashboard normals corrected
' 1.2.0 - tomate - invisible wall under defender to prevent ball stuck, rised disable lighting of decals a little bit, moved all colliadable objects to layer 2, collideable left lane protection added, roof upper right plastic added, lock all the objects
' 1.2.3 - iaakki - refactored NF flip code, postes_metalicos set non-collidable
' 1.2.4 - iaakki - Fixed slingshot thresholds, some rubber dampeners and sounds for few posts.
' 1.2.4b -  Rawd - Working VR backglass (Leojreimroc flasher lighting and code)
' 1.2.5 - Wylte - Just ball shadows (and some PoV testing for DT and FS)
' 1.2.6 - iaakki - POV adjustment, rubberizer & targetbouncer added, target code cleanup
' 1.2.7 - Rawd - Added LUT switcher and Fluppers new Display
' 1.2.8 - Sixtoe - Tweaked LUT stuff, minor script and other cleanup for release.
' 1.3 Release
' 1.3.1 - mcarter78 - full physics update, use vpm insert light mapping, use mod sol for flashers, add F12 options menu
' 1.3.2 - mcarter78 - mod sol for "basketball" flashers, finish options menu, GI tweaks, implement roth standups, some sound fixes
' 1.3.3 - mcarter78 - attempt to fix stuck ball, hook up center trophy sol flasher insert
' 1.3.4 - daRdog - Remove VR Mega Room
' 1.3.5 - daRdog - Add more performant VR Mega Room
' 1.3.6 - mcarter78 - Fix dimensions and realign objects.  Remove Flashers and replace with lights. Remove extra lights and fix assignments
' 1.3.7 - mcarter78 - Fix some GI placement, set all lights to hidden, remove a lot of visible objects, add some errantly deleted objects back in

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

'///////////////// Table Options ///////////////////

'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.8

'***********CABINET MODE**************
Const CabinetMode = 0       'Cabinet mode - Will hide the rails and scale the side panels higher

'**************VR ROOM****************
VRRoom = 0              '0 - VR Room off, 1 - Minimal Room, 2 - Ultra Minimal

'********TABLE FLASHER BLOOM**********
' TURN OFF IF TABLE STUTTERS
Const TableBloom = 1        '0 = OFF, 1 = ON

'********Dynamic Ball Shadows*********
'*** *BOTH* to 0 for simple shadow****
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's)
                  '2 = flasher image shadow, but it moves like ninuzzu's

'*******FAST FLIPS OPTION*************
Const UseSolenoids = 2        '1 = OFF     2 = ON
Const RubberizerEnabled = 1     '0 = normal flip rubber, 1 = more lively rubber for flips, 2 = more lively rubber for flips, different settings
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 1.1   'Level of bounces. 0.2 - 1.5 are probably usable values.

'*****Shooter Lane Shoot Button*******
Const ShootShooter = 1        '0 = OFF, 1 = ON

'*********RED SHOT CLOCK MOD**********
RedClock=0              '0 = OFF, 1 = ON

'*********************************************************************************************************************************


' Flupper LUT code...
Dim luts, lutpos
' ****** LUT overall contrast & brightness setting **************************************************************************************
luts = array("Fleep Natural Dark 1", "Fleep Natural Dark 2", "Fleep Warm Dark", "Fleep Warm Bright", "Fleep Warm Vivid Soft", "Fleep Warm Vivid Hard", "Skitso Natural and Balanced", "Skitso Natural High Contrast", "3rdaxis Referenced THX Standard", "CalleV Punchy Brightness&Contrast", "HauntFreaks Desaturated", "Tomate Washed Out", "VPW Original 1 to 1", "Bassgeige", "Blacklight", "B&W Comic Book", "NBA Fastbreak Original")

Const EnableMagnasave = 1   ' 1 - on; 0 - off; if on then the magnasave button let's you rotate all LUT's
'*********************************************************************************************************************************
' end Flupper LUT code...


Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50                 'add this here to redefine the ball size, placed before LoadVPM.'
Const BallMass = 1

LoadVPM "01120100", "WPC.VBS", 3.49 'minimum core.vbs version

Dim bsTrough, bsEject, bsSaucer1, bsSaucer2, bsSaucer3, bsSaucer4, mBallCatch, mDefender, MagnetCatch, bgball
Dim PlungerIM, x, bump1, bump2, bump3, ArenaMod, RedClock, FastFlips, ff

'************************************
ramps_lateral.blenddisablelighting=2
ramp_central.blenddisablelighting=2
'************************************

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

'Sub sw41_Hit:vpmTimer.PulseSw 41:End Sub

DIM target41step
Sub sw41_Hit
  vpmTimer.PulseSw 41:pSW41.TransY = -3:Target41Step = 1:sw41a.Enabled = 1
End Sub

Sub sw41a_timer()
  Select Case Target41Step
    Case 1:pSW41.TransY = -1
    Case 2:pSW41.TransY = 2
        Case 3:pSW41.TransY = 0:Me.Enabled = 0
     End Select
  Target41Step = Target41Step + 1
End Sub


DIM target42step
Sub sw42_Hit
  vpmTimer.PulseSw 42:pSW42.TransY = -3:Target42Step = 1:sw42a.Enabled = 1
End Sub

Sub sw42a_timer()
  Select Case Target42Step
    Case 1:pSW42.TransY = -1
    Case 2:pSW42.TransY = 2
        Case 3:pSW42.TransY = 0:Me.Enabled = 0
     End Select
  Target42Step = Target42Step + 1
End Sub

DIM target43step
Sub sw43_Hit
  vpmTimer.PulseSw 43:pSW43.TransY = -3:Target43Step = 1:sw43a.Enabled = 1
End Sub

Sub sw43a_timer()
  Select Case Target43Step
    Case 1:pSW43.TransY = -1
    Case 2:pSW43.TransY = 2
        Case 3:pSW43.TransY = 0:Me.Enabled = 0
     End Select
  Target43Step = Target43Step + 1
End Sub


DIM target28step
Sub sw28_Hit
  vpmTimer.PulseSw 28:pSW28.TransY = -3:Target28Step = 1:sw28a.Enabled = 1
End Sub

Sub sw28a_timer()
  Select Case Target28Step
    Case 1:pSW28.TransY = -1
    Case 2:pSW28.TransY = 2
        Case 3:pSW28.TransY = 0:Me.Enabled = 0
     End Select
  Target28Step = Target28Step + 1
End Sub

DIM target18step
Sub sw18_Hit
  vpmTimer.PulseSw 18:pSW18.TransY = -3:Target18Step = 1:sw18a.Enabled = 1
End Sub

Sub sw18a_timer()
  Select Case Target18Step
    Case 1:pSW18.TransY = -1
    Case 2:pSW18.TransY = 2
        Case 3:pSW18.TransY = 0:Me.Enabled = 0
     End Select
  Target18Step = Target18Step + 1
End Sub


'**************Bumpers animation*******************

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Sub Bumper2_Hit:vpmTimer.PulseSw 23:PlaySound "fx_bumper2":bump2 = 1:Me.TimerEnabled = 1:End Sub
' Sub Bumper2_Timer()
'     Select Case bump2
'         Case 1:Ring2.HeightTop = 15:Ring2.HeightBottom = 15:bump2 = 2
'         Case 2:Ring2.HeightTop = 25:Ring2.HeightBottom = 25:bump2 = 3
'         Case 3:Ring2.HeightTop = 35:Ring2.HeightBottom = 35:bump2 = 4
'         Case 4:Ring2.HeightTop = 45:Ring2.HeightBottom = 45:Me.TimerEnabled = 0
'     End Select
'
'     Bumper2R.State = ABS(Bumper2R.State - 1) 'refresh light
' End Sub

'*************************************

Const UseLamps = 0
'Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_SolenoidOff"
Const SFlipperOn = "fx_flipperup1"
Const SFlipperoff = "fx_flipperdown"
Const SCoin = "fx_coin"
Const cGameName = "nbaf_31"
Const cSingleLFlip=0
Const cSingleRFlip=0

Set GiCallback2 = GetRef("UpdateGI")

'sub updategi (no, aValue)

  'Eval("textbox"&no).text=enabled
  'debug.print no & " nro and aValue: " & aValue
'iaakki: ok NBA gi has few areas
'[4:35 PM] iaakki: 0 = "string 1"
'[4:36 PM] iaakki: 1 = "string 2"
'[4:36 PM] iaakki: 2 = "string 3"

'end sub

Sub UpdateGi(nr,step)
  Dim ii
  Select Case nr
  Case 0    'Bottom Playfield
    'string 1
    If step=0 Then
      For each ii in GI0:ii.state=0:Next
      Pegs.blenddisablelighting = 0
      Pegs001.blenddisablelighting = 1
    Else
      For each ii in GI0:ii.state=1:Next
      Pegs.blenddisablelighting = 0.08 * step
      Pegs001.blenddisablelighting = step + 1
    End If
    For each ii in GI0:ii.IntensityScale = 0.125 * step:Next
  Case 1    'Middle Playfield
    'string 2
    If step=0 Then
      For each ii in GI1:ii.state=0:Next
      GlobalFlasher.visible=0
      PinCab_Blades.image = "PinCab_Blades_Off"
      SideWalls.image = "sidesOFF"
    Else
      For each ii in GI1:ii.state=1:Next
      GlobalFlasher.visible=1
      PinCab_Blades.image = "PinCab_Blades_On"
      SideWalls.image = "sidesON"
    End If
    For each ii in GI1:ii.IntensityScale = 0.125 * step:Next
  Case 2    'Top Playfield
    'string 3
    If step=0 Then
      For each ii in GI2:ii.state=0:Next
      If VRRoom > 0 Then
        For each ii in BGGI:ii.visible = 0:Next
        For each ii in VRBGGIBulbs:ii.disablelighting = 0:Next
      End If
    Else
      For each ii in GI2:ii.state=1:Next
      If VRRoom > 0 Then
        For each ii in BGGI:ii.visible = 1:Next
        For each ii in VRBGGIBulbs:ii.disablelighting = 1:Next
      End If
    End If
    For each ii in GI2:ii.IntensityScale = 0.125 * step:Next
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

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot, bumper1, bumper2, bumper3)

    ' Trough & Ball Release
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 32, 33, 34, 35, 31, 0, 0
        .InitKick BallRelease, 180, 10
        .InitExitSnd "fx_Ballrel", "fx_solenoid"
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
        .KickForceVar = 3
        .KickAngleVar = 3
        .InitSaucer sw68, 68, 65, 32
        .Kickz = 1.2
        .InitExitSnd "fx_popper", "fx_Solenoid"
    End With

    Set bsSaucer2 = New cvpmBallStack
    With bsSaucer2
        .KickForceVar = 3
        .KickAngleVar = 3
        .InitSaucer sw67, 67, 26, 28
        .Kickz = 1.15
        .InitExitSnd "fx_popper", "fx_Solenoid"
    End With

    Set bsSaucer3 = New cvpmBallStack
    With bsSaucer3
        .KickForceVar = 3
        .KickAngleVar = 3
        .InitSaucer sw66, 66, 337, 28
        .Kickz = 1.15
        .InitExitSnd "fx_popper", "fx_Solenoid"
    End With

    Set bsSaucer4 = New cvpmBallStack
    With bsSaucer4
        .KickForceVar = 3
        .KickAngleVar = 3
        .InitSaucer sw65, 65, 293, 31.5
        .Kickz = 1.2
        .InitExitSnd "fx_popper", "fx_Solenoid"
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
    Const IMPowerSetting = 38 ' Plunger Power
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

If VRRoom = 0 then BackBall.CreateSizedBall(20).Image = "BasketBall"  ' VRADDED  - added VRroom switch

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
If VRroom > 0 then

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

    table1.ColorGradeImage = luts(lutpos)  ' flupper lut code..
  LoadLUT
  SetLUT

End Sub

Sub Table1_exit():SaveLUT:Controller.Pause = False:Controller.Stop:End Sub



'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
  If KeyCode = PlungerKey or KeyCode = LockBarKey Then Controller.Switch(11) = 1
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

  If keycode = RightMagnaSave and EnableMagnasave = 1 then
    ResetRollingTextTimer
    lutpos = lutpos + 1 : If lutpos > ubound(luts) Then lutpos = 0 : end if
        call myChangeLut
    playsound "LutChange"
  End if

  If keycode = LeftMagnaSave and EnableMagnasave = 1 then
    ResetRollingTextTimer
    lutpos = lutpos - 1 : If lutpos < 0 Then lutpos = ubound(luts) : end if
        call myChangeLut
    playsound "LutChange2"
    end if

  if keycode=StartGameKey then soundStartButton()

'nFozzy Begin'
  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
'nFozzy End'

    If vpmKeyDown(keycode) Then Exit Sub
End Sub


Sub table1_KeyUp(ByVal Keycode)

'nFozzy Begin'
  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
' If keycode = LeftFlipperKey Then
'   lfpress = 0
'   leftflipper.eostorqueangle = EOSA
'   leftflipper.eostorque = EOST
' End If
' If keycode = RightFlipperKey Then
'   rfpress = 0
'   rightflipper.eostorqueangle = EOSA
'   rightflipper.eostorque = EOST
' End If
'nFozzy End'

    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey or keycode = LockBarKey Then Controller.Switch(11) = 0
End Sub

'*********
' Switches
'********* -20 -15 -7

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
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
Sub Bumper1_Hit:vpmTimer.PulseSw 61:RandomSoundBumperMiddle(Bumper1):bump1 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper1_Timer()
    Select Case bump1
        Case 1:bump1 = 2: BumperRing1.Z = -30
        Case 2:bump1 = 3: BumperRing1.Z = -20
        Case 3:bump1 = 4: BumperRing1.Z = -10
        Case 4:Me.TimerEnabled = 0: BumperRing1.Z = 0
    End Select

    'Bumper1R.State = ABS(Bumper1R.State - 1) 'refresh light
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 23:RandomSoundBumperMiddle(Bumper2):bump2 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper2_Timer()
    Select Case bump2
        Case 1:bump2 = 2 : BumperRing2.Z = -30
        Case 2:bump2 = 3 : BumperRing2.Z = -20
        Case 3:bump2 = 4 : BumperRing2.Z = -10
        Case 4:Me.TimerEnabled = 0 :  : BumperRing2.Z = 0
    End Select
 '   Bumper2R.State = ABS(Bumper2R.State - 1) 'refresh light
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 62:PlaySound "fx_bumper3":bump3 = 1:Me.TimerEnabled = 1:End Sub
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
Sub Drain_Hit:ClearBallID:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub sw25_Hit:Playsound "fx_kicker_enter":bsEject.AddBall 0:End Sub

Sub sw65_Hit
    PlaySound "fx_kicker_enter"
    bsSaucer4.AddBall 0
End Sub

Sub sw66_Hit
    PlaySound "fx_kicker_enter"
    bsSaucer3.AddBall 0
End Sub

Sub sw67_Hit
    PlaySound "fx_kicker_enter"
    bsSaucer2.AddBall 0
End Sub

Sub sw68_Hit
    PlaySound "fx_kicker_enter"
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

SolCallBack(17) = "FlashSol17"    '"SetLamp 117," 'eject kickout flasher
SolCallBack(18) = "FlashSol18"    '"SetLamp 118," 'left jet bumper
SolCallBack(19) = "FlashSol19"    '"SetLamp 119," 'upper left         'BG Left
SolCallBack(20) = "FlashSol20"    '"SetLamp 120," 'upper right        'BG Right
SolCallBack(22) = "SetLamp 122,"  'trophy insert TODO
SolCallBack(24) = "FlashSol24"    '"SetLamp 124," 'lower right left

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

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Dim VRBGBALL  'VRADDED
FlipperBumper.isdropped = true  'VRADDED - start dropped


Sub SolBasket(Enabled)
    If Enabled Then

        If VRRoom = 0 then ' VRADDED - switch
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



'***************************************
'     Special JP Flippers, including:
' - tap code by Jimmifingers
' - recoil fix to enable dropcatches
' - ball hit sound
'**************************************

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


'******************
' RealTime Updates
'******************

Sub GameTimer_Timer
    RollingSound
  LampTimer
  UpdateLED
  Cor.Update
  LFLogo.RotZ = LeftFlipper.CurrentAngle
  RFlogo.RotZ = RightFlipper.CurrentAngle
  LeftFlipperSh.RotZ = LeftFlipper.currentangle
  RightFlipperSh.RotZ = RightFlipper.currentangle
  'slacktimer

   batleft.objroty = VRBasketFlipper.CurrentAngle + 1   'VRADDED Backglass flipper follow
End Sub


'*************************************************
' destruk's new vpmCreateBall for ball collision
' use it: vpmCreateBall kicker
'*************************************************

Set vpmCreateBall = GetRef("mycreateball")
Function mycreateball(aKicker)
    For cnt = 1 to ubound(ballStatus)
        If ballStatus(cnt) = 0 Then
            If Not IsEmpty(vpmBallImage) Then
                Set CurrentBall(cnt) = aKicker.CreateSizedBall(BSize).Image
            Else
                Set CurrentBall(cnt) = aKicker.CreateSizedBall(BSize)
            End If
            Set mycreateball = aKicker
            currentball(cnt).uservalue = cnt
            ballStatus(cnt) = 1
            ballStatus(0) = ballStatus(0) + 1
            If coff = False Then
                If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
            End If
            Exit For
        End If
    Next
End Function

'**************************************************************
' vpm Ball Collision based on the code by Steely & Pinball Ken
' added destruk's changes, ball size and height check by koadic
'**************************************************************

Const tnopb = 10 'max nr. of balls
Const nosf = 10  'nr. of sound files

ReDim CurrentBall(tnopb), BallStatus(tnopb)
Dim iball, cnt, coff, errMessage

XYdata.interval = 1
coff = False

For cnt = 0 to ubound(BallStatus):BallStatus(cnt) = 0:Next

' Create ball in kicker and assign a Ball ID used mostly in non-vpm tables
Sub CreateBallID(aKicker)
    For cnt = 1 to ubound(ballStatus)
        If ballStatus(cnt) = 0 Then
            Set CurrentBall(cnt) = aKicker.CreateSizedBall(Bsize)
            currentball(cnt).uservalue = cnt
            ballStatus(cnt) = 1
            ballStatus(0) = ballStatus(0) + 1
            If coff = False Then
                If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
            End If
            Exit For
        End If
    Next
End Sub

Sub ClearBallID
    On Error Resume Next
    iball = ActiveBall.uservalue
    currentball(iball).UserValue = 0
    If Err Then Msgbox Err.description & vbCrLf & iball
    ballStatus(iBall) = 0
    ballStatus(0) = ballStatus(0) -1
    On Error Goto 0
End Sub

' Ball data collection and B2B Collision detection. jpsalas: added height check
ReDim baX(tnopb, 4), baY(tnopb, 4), baZ(tnopb, 4), bVx(tnopb, 4), bVy(tnopb, 4), TotalVel(tnopb, 4)
Dim cForce, bDistance, xyTime, cFactor, id, id2, id3, B1, B2

Sub XYdata_Timer()
    xyTime = Timer + (XYdata.interval * .001)
    If id2 >= 4 Then id2 = 0
    id2 = id2 + 1
    For id = 1 to ubound(ballStatus)
        If ballStatus(id) = 1 Then
            baX(id, id2) = round(currentball(id).x, 2)
            baY(id, id2) = round(currentball(id).y, 2)
            baZ(id, id2) = round(currentball(id).z, 2)
            bVx(id, id2) = round(currentball(id).velx, 2)
            bVy(id, id2) = round(currentball(id).vely, 2)
            TotalVel(id, id2) = (bVx(id, id2) ^2 + bVy(id, id2) ^2)
            If TotalVel(id, id2) > TotalVel(0, 0) Then TotalVel(0, 0) = int(TotalVel(id, id2) )
        End If
    Next

    id3 = id2:B2 = 2:B1 = 1
    Do
        If ballStatus(B1) = 1 and ballStatus(B2) = 1 Then
            bDistance = int((TotalVel(B1, id3) + TotalVel(B2, id3) ) ^(1.04 * (CurrentBall(B1).radius + CurrentBall(B2).radius) / 50) )
            If((baX(B1, id3) - baX(B2, id3) ) ^2 + (baY(B1, id3) - baY(B2, id3) ) ^2) < (2800 * ((CurrentBall(B1).radius + CurrentBall(B2).radius) / 50) ^2) + bDistance Then
                If ABS(baZ(B1, id3) - baZ(B2, id3) ) < (CurrentBall(B1).radius + CurrentBall(B2).radius) Then collide B1, B2:Exit Sub
            End If
        End If
        B1 = B1 + 1
        If B1 = ubound(ballstatus) Then Exit Do
        If B1 >= B2 then B1 = 1:B2 = B2 + 1
    Loop

    If ballStatus(0) <= 1 Then XYdata.enabled = False

    If XYdata.interval >= 40 Then coff = True:XYdata.enabled = False
    If Timer > xyTime * 3 Then coff = True:XYdata.enabled = False
    If Timer > xyTime Then XYdata.interval = XYdata.interval + 1
End Sub

'Calculate the collision force and play sound
Dim cTime, cb1, cb2, avgBallx, cAngle, bAngle1, bAngle2

Sub Collide(cb1, cb2)
    If TotalVel(0, 0) / 1.8 > cFactor Then cFactor = int(TotalVel(0, 0) / 1.8)
    avgBallx = (bvX(cb2, 1) + bvX(cb2, 2) + bvX(cb2, 3) + bvX(cb2, 4) ) / 4
    If avgBallx < bvX(cb2, id2) + .1 and avgBallx > bvX(cb2, id2) -.1 Then
        If ABS(TotalVel(cb1, id2) - TotalVel(cb2, id2) ) < .000005 Then Exit Sub
    End If
    If Timer < cTime Then Exit Sub
    cTime = Timer + .1
    GetAngle baX(cb1, id3) - baX(cb2, id3), baY(cb1, id3) - baY(cb2, id3), cAngle
    id3 = id3 - 1:If id3 = 0 Then id3 = 4
    GetAngle bVx(cb1, id3), bVy(cb1, id3), bAngle1
    GetAngle bVx(cb2, id3), bVy(cb2, id3), bAngle2
    cForce = Cint((abs(TotalVel(cb1, id3) * Cos(cAngle-bAngle1) ) + abs(TotalVel(cb2, id3) * Cos(cAngle-bAngle2) ) ) )
    If cForce < 4 Then Exit Sub
    cForce = Cint((cForce) / (cFactor / nosf) )
    If cForce > nosf-1 Then cForce = nosf-1
    PlaySound("fx_collide" & cForce)
End Sub

' Get angle
Dim Xin, Yin, rAngle, Radit, wAngle', Pi
'Pi = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub GetAngle(Xin, Yin, wAngle)
    If Sgn(Xin) = 0 Then
        If Sgn(Yin) = 1 Then rAngle = 3 * Pi / 2 Else rAngle = Pi / 2
        If Sgn(Yin) = 0 Then rAngle = 0
        Else
            rAngle = atn(- Yin / Xin)
    End If
    If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
    If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
    wAngle = round((Radit + rAngle), 4)
End Sub

' #####################################
' ###### Flupper Flasher Domes    #####
' #####################################

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.2   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.2   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.4
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20)
Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
''initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "white" : InitFlasher 2, "blue" : InitFlasher 3, "red" : InitFlasher 4, "blue" : InitFlasher 5, "white"
InitFlasher 6, "yellow" : InitFlasher 7, "yellow" : InitFlasher 8, "yellow" : InitFlasher 9, "yellow" : InitFlasher 10, "red"
'' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 4,17 : RotateFlasher 5,0 : RotateFlasher 6,90
'RotateFlasher 7,0 : RotateFlasher 8,0
'RotateFlasher 9,-45 : RotateFlasher 10,90 : RotateFlasher 11,90
rotateflasher 6,47
rotateflasher 7,37
rotateflasher 8, -4
rotateflasher 9, -52

'Flasherlight5.BulbHaloHeight = 120
'Flasherlight4.BulbHaloHeight = 145
'Flasherlight1.BulbHaloHeight = 55
'Flasherlight2.BulbHaloHeight = 177
'Flasherlight3.BulbHaloHeight = 88

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
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(20,155,255) ': objlight(nr).intensity = 5000
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
  'FlasherFlash4.height = 226
  'FlasherFlash3.height = 200
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objlit(nr).visible = 1 : if Tablebloom = 1 then objbloom(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objbloom(nr).opacity = 50 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3' else objlight(nr).IntensityScale = 1 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.7 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub
Sub FlasherFlash7_Timer() : FlashFlasher(7) : End Sub
Sub FlasherFlash8_Timer() : FlashFlasher(8) : End Sub
Sub FlasherFlash9_Timer() : FlashFlasher(9) : End Sub
Sub FlasherFlash10_Timer() : FlashFlasher(10) : End Sub

' ###############################################
' ###### NBA Fastbreak flasher dome settings #####
' ###############################################

Sub FlashSol24(flstate)
  If Flstate Then
    Objlevel(1) = 1 : FlasherFlash1_Timer
    Objlevel(2) = 1 : FlasherFlash2_Timer
  End If
End Sub

Sub FlashSol17(flstate)
  If Flstate Then
    Objlevel(3) = 1 : FlasherFlash3_Timer
  End If
End Sub

Sub FlashSol18(flstate)
  If Flstate Then
    Objlevel(10) = 1 : FlasherFlash10_Timer
  End If
End Sub

Sub FlashSol19(flstate)
  Dim object
  If Flstate Then
    Objlevel(4) = 1 : FlasherFlash4_Timer
    If VRRoom > 0 Then
      For Each object in VRBGFL19:object.visible = 1:Next
      LEDBulbs018.disablelighting = 1
    End If
  Else
    If VRRoom > 0 Then
      For Each object in VRBGFL19:object.visible = 0:Next
      LEDBulbs018.disablelighting = 0
    End If
  End If
End Sub

Sub FlashSol20(flstate)
  Dim Object
  If Flstate Then
    Objlevel(5) = 1 : FlasherFlash5_Timer
    If VRRoom > 0 Then
      For Each object in VRBGFL20:object.visible = 1:Next
      LEDBulbs012.disablelighting = 1
    End If
  Else
    If VRRoom > 0 Then
      For Each object in VRBGFL20:object.visible = 0:Next
      LEDBulbs012.disablelighting = 0
    End If
  End If
End Sub

Sub FlashRondo(nr, flstate)
  If Flstate Then
    Select Case nr
      Case 67:Objlevel(9) = 1 : FlasherFlash9_Timer
      Case 68:Objlevel(8) = 1 : FlasherFlash8_Timer
      Case 77:Objlevel(6) = 1 : FlasherFlash6_Timer
      Case 78:Objlevel(7) = 1 : FlasherFlash7_Timer
    End Select
  End If
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

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
'LampTimer.Interval = 10 'lamp fading speed
'LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer()
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

Sub UpdateLamps
    NFadeLm 11, l11
  FadeDisableLighting2 11, p11, 122
    NfadeLm 12, l12
  FadeDisableLighting2 12, p12, 70
    NfadeLm 13, l13
  FadeDisableLighting2 13, p13, 70
    NfadeLm 14, l14
  FadeDisableLighting2 14, p14, 70
    NfadeLm 15, l15
  FadeDisableLighting2 15, p15, 122
    NfadeLm 16, l16
  FadeDisableLighting2 16, p16, 122
    NfadeLm 17, l17
  FadeDisableLighting2 17, p17, 122
    NfadeLm 18, l18
  FadeDisableLighting2 18, p18, 122
    NfadeLm 21, l21
  FadeDisableLighting2 21, p21, 122
    NfadeLm 22, l22
  FadeDisableLighting2 22, p22, 110
    NfadeLm 23, l23
  FadeDisableLighting2 23, p23, 110
    NfadeLm 24, l24
  FadeDisableLighting2 24, p24, 110
    NfadeLm 25, l25
  FadeDisableLighting2 25, p25, 120
    NfadeLm 26, l26
    NfadeLm 26, l26a
  FadeDisableLighting2 26, p26off, 15
    NfadeLm 27, l27
  FadeDisableLighting2 27, p27, 110
    NfadeLm 28, l28
  FadeDisableLighting2 28, p28, 122
    NfadeLm 31, l31
  FadeDisableLighting2 31, p31, 122
    NfadeLm 32, l32
  FadeDisableLighting2 32, p32, 122
    NfadeLm 33, l33
  FadeDisableLighting2 33, p33, 122
    NfadeLm 34, l34
  FadeDisableLighting2 34, p34, 122
    NfadeLm 35, l35
  FadeDisableLighting2 35, p35, 110
    NfadeLm 36, l36
  FadeDisableLighting2 36, p36, 110
    NfadeLm 37, l37
  NfadeLm 37, l37a
  FadeDisableLighting2 37, p37, 35
    NfadeLm 38, l38
    NfadeLm 38, l38a
  FadeDisableLighting2 38,p38, 60
    NfadeLm 41, l41
  FadeDisableLighting2 41, p41, 110
    NfadeLm 42, l42
  FadeDisableLighting2 42, p42, 110
    NfadeLm 43, l43
    NfadeLm 43, l43a
  FadeDisableLighting2 43, p43, 60
    NfadeLm 44, l44
  FadeDisableLighting2 44, p44, 110
    NfadeLm 45, l45
  FadeDisableLighting2 45, p45, 110
    NfadeLm 46, l46
  FadeDisableLighting2 46, p46, 33
    NfadeLm 47, l47
  FadeDisableLighting2 47, p47, 33
    NfadeLm 48, l48
  FadeDisableLighting2 48, p48, 70
    NfadeLm 51, l51
  NfadeLm 51, l51a
  FadeDisableLighting2 51,p51, 55
    NfadeLm 52, l52
  FadeDisableLighting2 52,p52, 155
    NfadeLm 53, l53
  NfadeLm 53, l53a
  FadeDisableLighting2 53,p53, 300
    NfadeLm 54, l54
  NfadeLm 54, l54a
  FadeDisableLighting2 54,p54, 88
    NfadeLm 55, l55
  NfadeLm 55, l55a
  FadeDisableLighting2 55,p55, 55
    NfadeLm 56, l56
  FadeDisableLighting2 56, p56, 90
    NfadeLm 57, l57
  FadeDisableLighting2 57, p57, 60
    NfadeLm 58, l58
  FadeDisableLighting2 58, p58, 70
    NFadeLm 61, l61b
    NfadeLm 61, l61
  FadeDisableLightingM 61, p61b, 90
  FadeDisableLighting2 61, p61, 90
    NfadeLm 62, l62
  FadeDisableLighting2 62, p62, 110
    NfadeLm 63, l63
  FadeDisableLighting2 63, p63, 110
    NfadeLm 64, l64
  FadeDisableLighting2 64, p64, 110
    NfadeLm 65, l65
  FadeDisableLighting2 65, p65, 110
    NfadeLm 66, l66
    NfadeLm 66, l66a
  FadeDisableLighting2 66,p66, 35

    NfadeLm 71, l71
  FadeDisableLighting2 71, p71, 110
    NfadeLm 72, l72
  FadeDisableLighting2 72, p72, 110

    NfadeLm 73, l73
    NfadeLm 73, l73a
  FadeDisableLighting2 73, p73, 35

    NfadeLm 74, l74
  FadeDisableLighting2 74, p74, 110
    NfadeLm 75, l75
  FadeDisableLighting2 75, p75, 110
    NFadeLm 76, l76
  NfadeLm 76, l76a
  FadeDisableLighting2 76, p76, 35
    NfadeLm 81, l81
  FadeDisableLighting2 81, p81, 110
    NfadeLm 82, l82
  FadeDisableLighting2 82, p82, 110
    NfadeLm 83, l83
  NfadeLm 83, l83a
  FadeDisableLighting2 83, p83, 35
    NfadeLm 84, l84
  FadeDisableLighting2 84, p84, 110
    NfadeLm 85, l85
  FadeDisableLighting2 85, p85, 110
    NfadeLm 86, l86
  FadeDisableLighting2 86, p86, 110
    NfadeL 87, l87
    NfadeL 88, l88

'    fadeobj 67, l67, "of_on", "of_a", "of_b", "empty"
'    fadeobj 68, l68, "of_on", "of_a", "of_b", "empty"
'    fadeobj 77, l77, "of_on", "of_a", "of_b", "empty"
'    fadeobj 78, l78, "of_on", "of_a", "of_b", "empty"
' flash 67,f67
' flash 68,f68
  flash2 67
  flash2 68
  flash2 77
  flash2 78
  'flash 78,f78



    ' flashers old
'   fadeobj 117, f17, "rf_on", "rf_a", "rf_b", "empty"
' NFadeLm 118, bumper1
' fadeobj 118, f18, "rf_on", "rf_a", "rf_b", "empty"
'   fadeobj 119, f19, "wf_on", "wf_a", "wf_b", "empty"
'   fadeobj 120, f20, "bf_on", "bf_a", "bf_b", "empty"
'   fadeobjm 124, f24, "wf_on", "wf_a", "wf_b", "empty"
'   fadeobj 124, f24b, "bf_on", "bf_a", "bf_b", "empty"

' flashers new
' flash 117,f17
' flash 118,f18
' flash 119,f19
' flash 120,f20
' flashm 124,f24b
' flash 124,f24
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub


const gInsertLevel = 0.9

Sub FadeDisableLightingM(nr, a, alvl)
' debug.print "fade value" & FadingLevel(nr)
  Select Case FadingLevel(nr)
    Case 4
      a.UserValue = a.UserValue - 0.200
      If a.UserValue < 0 Then
        a.UserValue = 0
        'FadingLevel(nr) = -1
      end If
      a.BlendDisableLighting = alvl * a.UserValue * gInsertLevel 'brightness
    Case 5
      a.UserValue = a.UserValue + 0.200
      If a.UserValue > 1 Then
        a.UserValue = 1
        'FadingLevel(nr) = -1
      end If
      a.BlendDisableLighting = alvl * a.UserValue * gInsertLevel 'brightness
  End Select
End Sub

Sub FadeDisableLighting2(nr, a, alvl)
  'if nr = 83 then debug.print "fade level: " & a.UserValue
  Select Case FadingLevel(nr)
    Case 4
      a.UserValue = a.UserValue - 0.200
      If a.UserValue < 0 Then
        a.UserValue = 0
        FadingLevel(nr) = 0
      end If
      a.BlendDisableLighting = alvl * a.UserValue * gInsertLevel 'brightness
    Case 5
      a.UserValue = a.UserValue + 0.200
      If a.UserValue > 1 Then
        a.UserValue = 1
        FadingLevel(nr) = 1
      end If
      a.BlendDisableLighting = alvl * a.UserValue * gInsertLevel 'brightness
  End Select
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

Sub NFadeLm(nr, object)
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

'Flashf77
Sub Flash2(nr)
  'debug.print FadingLevel(nr)
    Select Case FadingLevel(nr)
        'Case 4 'off
      'Flashf77(False)
        Case 5 ' on
      FlashRondo nr, True
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

Sub Flashm2(nr, object) 'multiple flashers, it just sets the flashlevel
    Flashf77(FlashLevel(nr))
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1              'wait
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


'********************
' Diverse Help/Sounds
'********************

'right ramp sounds
Sub RRHit0_Hit:PlaySoundAtBall "fx_rr2":End Sub
Sub RRHit1_Hit:PlaySoundAtBall "fx_rr5":End Sub
Sub RRHit2_Hit:PlaySoundAtBall "fx_rr6":End Sub

''Sub ARubbers_Hit(idx):PlaySound "fx_rubber":End Sub
'Sub APostRubbers_Hit(idx):PlaySound "fx_rubber":End Sub
'Sub AMetals_Hit(idx):PlaySound "fx_MetalHit":End Sub
'Sub AGates_Hit(idx):PlaySound "fx_Gate":End Sub
'Sub APlastics_Hit(idx):PlaySound "fx_plastichit":End Sub

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
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*11)+1,DOFFlippers), FlipperLeftHitParm, Flipper
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
  'debug.print "Rubber Sound: " & finalspeed
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
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*8)+1), Vol(ActiveBall) * RubberWeakSoundFactor
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
  TargetBouncer activeball, 1.2
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
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

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

Sub ShotClock_Timer

  If RedClock=1 then ShotClock.Enabled=False
  If RedClock=1 then RedShotClock.Enabled=True

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

  If l87.State=1 and VRRoom = 0 and ShootShooter = 1 then Shoot_Flash.visible = True else Shoot_Flash.visible =false
  If l87.State=1 and VRRoom > 0 then VR_Shoot_Flash.visible = True else VR_Shoot_Flash.visible =false

  If l88.State=1 and VRRoom > 0 then VR_Start.visible =true else VR_Start.visible =false

End Sub

Sub RedShotClock_Timer

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

  If l87.State=1 and VRRoom = 0 and ShootShooter = 1 then Shoot_Flash.visible = True else Shoot_Flash.visible =false
  If l87.State=1 and VRRoom > 0 then VR_Shoot_Flash.visible = True else VR_Shoot_Flash.visible =false

  If l88.State=1 and VRRoom > 0 then VR_Start.visible =true else VR_Start.visible =false

End Sub


'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
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

  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

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
                                        BOT(b).velx = BOT(b).velx / 1.3
                                        BOT(b).vely = BOT(b).vely - 0.5
                                end If
                        Next
                End If
        Else
                If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 15 then
                        EOSNudge1 = 0
                end if
        End If
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
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  if RubberizerEnabled = 1 then Rubberizer(parm)
  if RubberizerEnabled = 2 then Rubberizer2(parm)
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  if RubberizerEnabled = 1 then Rubberizer(parm)
  if RubberizerEnabled = 2 then Rubberizer2(parm)
  RightFlipperCollide parm
End Sub


' iaakki Rubberizer
sub Rubberizer(parm)
  if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
    'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * 1.2
    activeball.vely = activeball.vely * 1.2
    'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  Elseif parm <= 2 and parm > 0.2 Then
    'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * -1.2
    activeball.vely = activeball.vely * 1.4
    'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  end if
end sub

' apophis Rubberizer
sub Rubberizer2(parm)
  if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
    'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = -activeball.angmomz * 2
    activeball.vely = activeball.vely * 1.2
    'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  Elseif parm <= 2 and parm > 0.2 Then
    'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = -activeball.angmomz * 0.5
    activeball.vely = activeball.vely * (1.2 + rnd(1)/3 )
    'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  end if
end sub


'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 and aBall.vely > 0 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz

    'debug.print "velz: " & round(aBall.velz,2) & "   horizontal vel: " & round(sqr(abs(aBall.velx^2 + aBall.vely^2)),2) & "  angle: " & Atn(aBall.velz/sqr(abs(aBall.velx^2 + aBall.vely^2)))*180/pi

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
    'debug.print " ---> velz: " & round(aBall.velz,2) & "   horizontal vel: " & round(sqr(abs(aBall.velx^2 + aBall.vely^2)),2) & "  angle: " & Atn(aBall.velz/sqr(abs(aBall.velx^2 + aBall.vely^2)))*180/pi
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
    elseif TargetBouncerEnabled = 2 and aball.z < 30 and aBall.vely > 0 then
    'debug.print "velz: " & activeball.velz
    Select Case Int(Rnd * 4) + 1
      Case 1: zMultiplier = defvalue+1.1
      Case 2: zMultiplier = defvalue+1.05
      Case 3: zMultiplier = defvalue+0.7
      Case 4: zMultiplier = defvalue+0.3
    End Select
    aBall.velz = aBall.velz * zMultiplier * TargetBouncerFactor
    'debug.print "----> velz: " & activeball.velz
    'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  'debug.print "dposts"
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  'debug.print "dsleeves"
  SleevesD.Dampen Activeball
  TargetBouncer activeball, 1.1
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

'******************************************************
'   STAND-UP TARGET INITIALIZATION
'******************************************************

'Define a variable for each stand-up target
Dim ST12, ST15 , ST85, ST86, ST23

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:      vp target to determine target hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             transy must be used to offset the target animation
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0

'ST12 = Array(sw12, sw12p,12, 0)
'ST15 = Array(sw15, sw15p,15, 0)
'ST23 = Array(sw23, sw23p,23, 0)
'ST85 = Array(sw85, frankytargets,85, 0)
'ST86 = Array(sw86, frankytargets,86, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST12, ST15, ST85, ST86, ST23)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  0.5       'vpunits per animation step (control return to Start)
Const STMaxOffset = 6       'max vp units target moves when hit
Const STHitSound = "fx_target"  'Stand-up Target Hit sound

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  'PlayTargetSound
  STArray(i)(3) =  STCheckHit(Activeball,STArray(i)(0))

  If STArray(i)(3) <> 0 Then
    DTBallPhysics Activeball, STArray(i)(0).orientation, STMass
  End If
  DoSTAnim
End Sub

Sub STHit2(target, switch)
  'PlayTargetSound
  If STCheckHit(Activeball,target) = 1 Then
    vpmTimer.PulseSw switch
  End If
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 to uBound(STArray)
    If STArray(i)(2) = switch Then STArrayID = i:Exit Function
  Next
End Function

'Check if target is hit on it's face
Function STCheckHit(aBall, target)
  dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)

  If perpvel <= 0 or perpvelafter >= 0 Then
    STCheckHit = 0
  Else
    STCheckHit = 1
  End If
End Function


Sub DoSTAnim()
  Dim i
  For i=0 to Ubound(STArray)
    STArray(i)(3) = STAnimate(STArray(i)(0),STArray(i)(1),STArray(i)(2),STArray(i)(3))
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

' Used for drop targets and stand up targets
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
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
Sub FrameTimer_Timer()
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then
    DynamicBSUpdate 'update ball shadows
  Else
    me.Enabled = False 'Remove this if you add anything else to the timer!
  End If
End Sub

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

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

Dim PI: PI = 4*Atn(1)

'Function Atn2(dy, dx)
' If dx > 0 Then
'   Atn2 = Atn(dy / dx)
' ElseIf dx < 0 Then
'   If dy = 0 Then
'     Atn2 = pi
'   Else
'     Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
'   end if
' ElseIf dx = 0 Then
'   if dy = 0 Then
'     Atn2 = 0
'   else
'     Atn2 = Sgn(dy) * pi / 2
'   end if
' End If
'End Function
'
'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

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

'////////////////////// ***OPTIONS*** //////////////////////////////

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

DIM VRRoom, VRThings
If VRRoom > 0 Then
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
  End If
  If VRRoom = 2 Then
    for each VRThings in VRCab:VRThings.visible = 0:Next
    for each VRThings in VRStuff:VRThings.visible = 0:Next
    for each VRThings in Desktop:VRThings.visible = 0:Next
    'PinCab_Backglass.visible = 1
    DMD.visible = true ' added because its part of the VRCab collection which is made invisible above.
    PinCab_Backbox.visible = 1
    PinCab_Backbox.image = "Pincab_Backbox_NBA_Min"
  End If
    SetBackglass
    BGDarkfront.visible = 1
    NBABackglassBack.visible = 1
Else
    for each VRThings in VRCab:VRThings.visible = 0:Next
    for each VRThings in VRStuff:VRThings.visible = 0:Next
    for each VRThings in Desktop:VRThings.visible = 1:Next
        for each VRThings in VRBGStuff:VRThings.visible = 0:Next


    If DesktopMode then ScoreText.visible = 1 else ScoreText.visible = 0 End If
    Ramp_central.image = "ramp_centre"
    ramps_lateral.image = "ramp_lateral"
    if ShootShooter = 1 then ShootButton.visible = 1 else ShootButton.visible = 0
End if



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


'********************
' LUT Stuff
'********************
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
'12 = VPW Original 1 to 1
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book
'16 = NBA Original

Dim LUTset

Sub SetLUT  'AXS
  Table1.ColorGradeImage = "LUT" & lutpos
end sub

Sub SaveLUT
  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if

  if lutpos = "" then lutpos = 16 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "NBAFBLUT.txt",True)
  ScoreFile.WriteLine lutpos
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub

Sub LoadLUT
  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    lutpos=17
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "NBAFBLUT.txt") then
    lutpos=17
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "NBAFBLUT.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      lutpos=17
      Exit Sub
    End if
    lutpos = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
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

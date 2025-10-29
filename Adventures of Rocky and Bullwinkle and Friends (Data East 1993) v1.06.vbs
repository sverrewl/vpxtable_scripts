'==============================================================================================='
'                                                                                           '                                                                                           '
'           The Adventures of Rocky and Bullwinkle and Friends              '
'                                      DataEast (1993)                                    '
'                  http://www.ipdb.org/machine.cgi?id=23                          '
'                                               '
'                             Created by: cyberpez                                '
'                                               '
'==============================================================================================='
' v1.03 - nFozzy, Fleep, Dynamic Ball Shadows, 2 Hybrid VR Rooms, VR Room Assets from Arvid - TastyWasps
' v1.04 - Modifications to saw mechanism on right drop targets. - Cyberpez
'   - Fixed sw52 not popping ball up from InRect function that was removed from JP rolling sounds in prev. mod. - TastyWasps
'   - Desktop POV updated to get layback to 0 (no stretched ball)
'     - Added a script option to not show desktop simulated backglass objects for desktop tweaker peoples.
' v1.05 - Desktop rails and lockbar added by passion4pins
'   - UseVPMDMD option added - TastyWasps
' v1.06 - Adjusted desktop view backglass position slightly, gameplay tweaks - TastyWasps

'Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Dim EnableBallControl
Dim ColoredGI
Dim CLG
Dim FlipperColor
Dim WBRBulb
Dim RRBulb
Dim RedBumperCaps
Dim Instruction_Cards
Dim CustomSideWalls
Dim BallMod
Dim RubberColor

EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

'***********************************************************************************************
' TABLE OPTIONS
'***********************************************************************************************

'----- Target Bouncer Levels -----
Const TargetBouncerEnabled = 1    ' 0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   ' Level of bounces. Recommmended value of 0.7

'----- General Sound Options -----
Const VolumeDial = 0.80       ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5      ' Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      ' Level of ramp rolling volume. Value between 0 and 1

'----- Ball Shadow Options -----
Const DynamicBallShadowsOn = 1    ' 0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   ' 0 = Static shadow under ball ("flasher" image, like JP's)
'                 ' 1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 ' 2 = flasher image shadow, but it moves like ninuzzu's

'----- VR Room Options -----
Const VRRoomChoice = 1        ' 1 = Cartoon Room 2 - Minimal Room

'----- Desktop Simulated Backglass -----
Const DesktopBackglass = 1      ' 1 = Backglass simulated in top right desktop area, 0 = Turn off those desktop backglass objects

'----- VPM DMD Option -----
Const ShowVPMDMD = True       ' False = Do not show native VPinMame DMD in desktop/vr,  True = Show native VPinMame DMD in desktop/vr

Const tnob = 5
Const lob = 0
Dim tablewidth:tablewidth = Table1.width
Dim tableheight:tableheight = Table1.height
Dim gilvl: gilvl = 1

'  Ball Mod
'0 = Normal
'1 = Marbled

BallMod = 0

'  Colored Lane Guides
'0 = clear
'1 = red

CLG = 1


'  GI Color MOD
'0 = Random
'1 = White
'2 = Colored

ColoredGI = 1


'  Way Back Ramp Bulb Color
'0 = Red
'1 = Blue

WBRBulb = 1


'  Rocky Ramp Bulb Color
'0 = Red
'1 = Blue
RRBulb = 0


'  Red Bumper Caps
'0 = Normal
'1 = Red
RedBumperCaps = 0

'  Instruction Cards
'1 = Eormal Credit
'2 = Normal Free Play
'3 = Custom 1
'4 = Custom 2
Instruction_Cards = 1

'  Flipper colors
'0 = Random
'1 = Yellow Flipper Blue Rubber
'2 = Yellow flipper Blue Rubber Blue DE logo
'3 = White Flipper Black Rubber
'4 = White Flipper Black Rubber Blue DE Logo
'5 = White Flipper Black Rubber Black DE Logo
'6 = White Flipper Red Rubber
'7 = White Flipper Red Rubber Blue DE Logo
'8 = White Flipper Yellow Rubber
'9 = White Flipper Yellow Rubber Blue DE Logo

FlipperColor = 3

'SideWalls
'0=None
'1=Normal Blue
'2=Custom
CustomSideWalls = 1

'RubberColor
'0=Random
'1=White
'2=Black
'3=Blue
RubberColor = 1

' DMD rotation
Const cDMDRotation      = 0         ' 0 or 1 for a DMD rotation of 90°

' ball size
Const BallSize = 50
Const BallMass = 1

' VPinMAME ROM name

Const cGameName       = "rab_320"   '
'Const cGameName      = "rab_130"

'===============================================================================================
' General constants and variables
'===============================================================================================

Const UseSolenoids    = True
Const UseLamps      = False
Const UseGI       = True
Const UseSync       = False
Const HandleMech    = False

Const SSolenoidOn     = "SolOn"
Const SSolenoidOff    = "SolOff"
Const SCoin       = ""
Const SKnocker      = "Knocker"

Dim I, x, obj, bsTrough, bsSaucer, dtRBank

Dim UseVPMDMD
Dim DesktopMode: DesktopMode = Table1.ShowDT

'  Setup Desktop
If DesktopMode = True then
  If ShowVPMDMD = True Then
    UseVPMDMD = True
  End If
  If DesktopBackglass = 1 Then
    BBGlass.Visible = true
    BBBackground.Visible = true
    BBLion.Visible = true
    BBRhino.Visible = true
    BBArm.Visible = true
    BBRocky.Visible = true
  Else
    BBGlass.Visible = false
    BBBackground.Visible = false
    BBLion.Visible = false
    BBRhino.Visible = false
    BBArm.Visible = false
    BBRocky.Visible = false
    Ramp62.Visible = false
    Ramp63.Visible = false
    Ramp001.Visible = false
  End If
Else
  UseVPMDMD = False
  ' Hide rails and lockbar for cab mode...
  Ramp62.Visible = false
  Ramp63.Visible = false
  Ramp001.Visible = false
  BBGlass.Visible = false
  BBBackground.Visible = false
  BBLion.Visible = false
  BBRhino.Visible = false
  BBArm.Visible = false
  BBRocky.Visible = false
End If

' VR Room Auto-Detect
Dim VR_Obj, VRRoom

If RenderingMode = 2 Then
  BBGlass.Visible = 0
  BBBackground.Visible = 0
  BBLion.Visible = 0
  BBRhino.Visible = 0
  BBArm.Visible = 0
  BBRocky.Visible = 0
  pSideWalls.Visible = 0
  Ramp62.Visible = false
  Ramp63.Visible = false
  Ramp001.Visible = false
  For Each VR_Obj in VRCabinet:VR_Obj.Visible = 1:Next
  For Each VR_Obj in VRBackglass:VR_Obj.Visible = 1:Next
  If VRRoomChoice = 1 Then
    For Each VR_Obj in VRCartoonRoom:VR_Obj.Visible = 1:Next
  Else
    For Each VR_Obj in VRMinimalRoom:VR_Obj.Visible = 1:Next
  End If
  If ShowVPMDMD = True Then
    UseVPMDMD = True
  End If
End If

LoadVPM "01560000","DE.VBS",3.1

'===============================================================================================
' Solenoids
'===============================================================================================

SolCallback(sLLFlipper) = "solLFlipper"
SolCallback(sLRFlipper) = "solRFlipper"
Solcallback(1)  ="kisort"
SolCallback(2)  = "KickBallToLane"
SolCallback(3)  = "HatTrickLion"
solcallback(4)  ="KickBallUp"
Solcallback(5)  ="sw29kick"
SolCallback(6)  = "HatTrickRocky"
SolCallback(7)  = "HatTrickRhino"
SolCallback(8)  ="vpmSolSound ""Knocker"","
SolCallBack(9)  ="SolNell"
SolCallback(11) = "SolGi" 'gi
solcallback(12) ="SolDiverter2"
solcallback(13) ="SolDiverter1"
solCallback(14) ="TopGate"
solCallback(15) ="DropsUp"
SolCallBack(16) ="SolAutoPlungerIM"
'SolCallback(17) = "SetLamp 117,"             'Sol17 Left Bumper
'SolCallback(18) = "SetLamp 118,"             'Sol18 Center Bumper
'SolCallback(19) = "SetLamp 119,"             'Sol19 Right Bumper
SolCallBack(22) ="AutoPlungeLK"

'Flashers/GI
Solcallback(25) = "SetLamp 101," '1r
Solcallback(26) = "SetLamp 102," '2r
Solcallback(27) = "SetLamp 103," '3r
Solcallback(28) = "SetLamp 104," '4r
SolCallback(29) = "SetLamp 105," '5r
Solcallback(30) = "SetLamp 106," '6r
SolCallback(31) = "SetLamp 107," '7r
SolCallback(32) = "SetLamp 108," '8r

'Dim f,g,h,i,j
Dim Ball(6)
Dim InitTime
Dim TroughTime
Dim EjectTime
Dim MaxBalls
Dim TroughCount
Dim TroughBall(7)
Dim TroughEject
Dim Momentum
Dim UpperGIon
Dim Multiball
Dim BallsInPlay
Dim iBall
Dim fgBall

Dim bsLEjet, bsUpperEject, Lnell, mNell, plungerIM, dtDrop

 Sub InitVPM()
     With Controller
         .GameName = cGameName
         .SplashInfoLine = "The Adventures of Rocky and Bullwinkle and Fiends" & vbNewLine & "VPX - cyberpez"
         .HandleMechanics = 0
         .HandleKeyboard = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .ShowTitle = 0
'    .hidden = 0
''     If DesktopMode = true then .hidden = true Else .hidden = false End If
'         If Err Then MsgBox Err.Description
     End With
     On Error Goto 0
     Controller.SolMask(0) = 0
     vpmTimer.AddTimer 4000, "Controller.SolMask(0)=&Hffffffff'"
     Controller.Run
End Sub


Sub Table1_Init
  ' table initialization
  InitVPM

' ' do some controller settings
' On Error Resume Next
' With Controller
'   .GameName               = cGameName
'   .SplashInfoLine             = "Table1 "
'   .HandleKeyboard             = False
'   .ShowTitle                = False
'   .ShowDMDOnly              = True
'   .ShowFrame                = False
'   .ShowTitle                = False
'
'   .Games(cGameName).Settings.Value("rol") = cDMDRotation
'   .Run GetPlayerHWnd
'   If Err Then MsgBox Err.Description
' End With
' On Error Goto 0

  ' basic pinmame timer
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled  = True

  ' nudging
  vpmNudge.TiltSwitch   = 1
  vpmNudge.Sensitivity  = 3
' vpmNudge.TiltObj    = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

     'SuperVuk
     Set bsLEjet =new cvpmBallStack
     With  bsLEjet
           .KickForceVar = 3
           .KickAngleVar = 3
           .InitSw 0,29,0,0,0,0,0,0
           .InitKick Sw29,182,20
           .KickZ = 0.4
'        .InitExitSnd "Popper", "Solenoid"
      End With

    ' Impulse Plunger
    Const IMPowerSetting = 50 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 14
        .InitExitSnd "BallRelease1", "BallRelease1"
        .CreateEvents "plungerIM"
    End With

     ' Drop Targets
     Set dtDrop = new cvpmDropTarget
     With dtDrop
        .Initdrop Array(t17,t18,t19,t20,t21), Array(17,18,19,20,21)
      End With

  CheckOptions

  vpmInit me

  ' ball trough system
  MaxBalls=3
  InitTime=61
  EjectTime=0
  TroughEject=1
  TroughCount=0
  iBall = 3
  fgBall = false

    CreateBalls

  ' Diverters
  Diverter1a.isDropped = False
  Diverter1b.isDropped = True
  pPFDiverter.ObjRotZ = -2
  Diverter2a.isDropped = True
  Diverter2b.isDropped = False
  pRampDiverter.ObjRotZ = -2

  ' Trough Helpers
  rTroughW1.IsDropped = true
  rTroughW2.IsDropped = true

End Sub


Sub Table1_Exit:Controller.Stop:End Sub
Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_UnPaused:Controller.Pause = False:End Sub

' **** Key_Down ****
Sub Table1_KeyDown(ByVal keyCode)

  If keycode = LeftFlipperKey Then
    Controller.Switch(15) = True
    FlipperActivate LeftFlipper, LFPress
  End If
  If keycode = RightFlipperKey Then
    Controller.Switch(16) = True
    FlipperActivate RightFlipper, RFPress
  End If

  If keyCode=PlungerKey or keyCode=LockBarKey Then
    Controller.Switch(9)=True
  End If

  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 2:SoundNudgeCenter()

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If keycode=StartGameKey Then soundStartButton()

    ' Manual Ball Control
  If keycode = 46 Then          ' C Key
    If EnableBallControl = 1 Then
      EnableBallControl = 0
    Else
      EnableBallControl = 1
    End If
  End If

    If EnableBallControl = 1 Then
    If keycode = 48 Then        ' B Key
      If BCboost = 1 Then
        BCboost = BCboostmulti
      Else
        BCboost = 1
      End If
    End If
    If keycode = 203 Then BCleft = 1  ' Left Arrow
    If keycode = 200 Then BCup = 1    ' Up Arrow
    If keycode = 208 Then BCdown = 1  ' Down Arrow
    If keycode = 205 Then BCright = 1 ' Right Arrow
  End If

    If KeyDownHandler(keyCode) Then Exit Sub

End Sub

' **** Key_Up ***
Sub Table1_KeyUp(ByVal keyCode)

  If keycode = LeftFlipperKey Then
    Controller.Switch(15) = False
    FlipperDeActivate LeftFlipper, LFPress
  End If

  If keycode = RightFlipperKey Then
    Controller.Switch(16) = False
    FlipperDeActivate RightFlipper, RFPress
  End If

  If keyCode=PlungerKey or keyCode=LockBarKey Then
    Controller.Switch(9)=False
  End If

    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If

  If KeyUpHandler(keyCode) Then Exit Sub

End Sub



'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'   Options
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Dim Red, RedFull, RedI, Pink, PinkFull, PinkI, White, WhiteFull, WhiteI, Blue, BlueFull, BlueI, Yellow, YellowFull, YellowI, Green, GreenFull, GreenI, Orange, OrangeFull, OrangeI, Purple, PurpleFull, PurpleI, AmberFull, Amber, AmberI
Dim CustomSideWallsType, ColoredGIType, FlipperColorType, Instruction_CardsType, RubberColorType, xxRubbers

' colored lane guides

Sub CheckOptions()

If CLG = 1 Then

  pLaneGuideL.Material = "Clear Plastic red"
  pLaneGuideR.Material = "Clear Plastic red"

Else

  pLaneGuideL.Material = "Clear Plastic"
  pLaneGuideR.Material = "Clear Plastic"

End If


RedFull = rgb(255,0,0)
Red = rgb(255,0,0)
RedI = 5
PinkFull = rgb(255,0,128)
Pink = rgb(255,0,255)
PinkI = 5
WhiteFull = rgb(255,255,128)
White = rgb(255,255,255)
WhiteI = 7
BlueFull = rgb(0,128,255)
Blue = rgb(0,255,255)
BlueI = 10
YellowFull = rgb(255,255,128)
Yellow = rgb(255,255,0)
YellowI = 20
GreenFull = rgb(128,255,128)
Green = rgb(0,255,0)
GreenI = 20

PurpleFull = rgb(128,0,255)
Purple = rgb(64,0,128)
PurpleI = 2
OrangeFull = rgb(255,128,64)
Orange = rgb(128,128,0)
OrangeI = 20
AmberFull = rgb(255,197,143)
Amber = rgb(255,197,143)
AmberI = 10



If ColoredGI = 0 then
    ColoredGIType = Int(Rnd*2)+1
  Else
    ColoredGIType = ColoredGI
End If



If ColoredGIType = 1 Then


  gi1a.Color=Amber
  gi1a.ColorFull=AmberFull
  gi1b.Color=Amber
  gi1b.ColorFull=AmberFull
  gi1c.Color=White
  gi1c.ColorFull=WhiteFull
  gi1a.Intensity = AmberI

  gi2a.Color=Amber
  gi2a.ColorFull=AmberFull
  gi2b.Color=Amber
  gi2b.ColorFull=AmberFull
  gi2c.Color=White
  gi2c.ColorFull=WhiteFull
  gi2a.Intensity = AmberI

  gi3a.Color=Amber
  gi3a.ColorFull=AmberFull
  gi3b.Color=Amber
  gi3b.ColorFull=AmberFull
  gi3c.Color=White
  gi3c.ColorFull=WhiteFull
  gi3a.Intensity = AmberI

  gi4a.Color=Amber
  gi4a.ColorFull=AmberFull
  gi4b.Color=Amber
  gi4b.ColorFull=AmberFull
  gi4c.Color=White
  gi4c.ColorFull=WhiteFull
  gi4a.Intensity = AmberI

  gi5a.Color=Amber
  gi5a.ColorFull=AmberFull
  gi5b.Color=Amber
  gi5b.ColorFull=AmberFull
  gi5c.Color=White
  gi5c.ColorFull=WhiteFull
  gi5a.Intensity = AmberI

  gi6a.Color=Amber
  gi6a.ColorFull=AmberFull
  gi6b.Color=Amber
  gi6b.ColorFull=AmberFull
  gi6c.Color=White
  gi6c.ColorFull=WhiteFull
  gi6a.Intensity = AmberI

  gi7a.Color=Amber
  gi7a.ColorFull=AmberFull
  gi7b.Color=Amber
  gi7b.ColorFull=AmberFull
  gi7c.Color=White
  gi7c.ColorFull=WhiteFull
  gi7a.Intensity = AmberI

  gi8a.Color=YAmber
  gi8a.ColorFull=AmberFull
  gi8b.Color=Amber
  gi8b.ColorFull=AmberFull
  gi8c.Color=White
  gi8c.ColorFull=WhiteFull
  gi8a.Intensity = AmberI

  gi9a.Color=Amber
  gi9a.ColorFull=AmberFull
  gi9b.Color=Amber
  gi9b.ColorFull=AmberFull
  gi9c.Color=White
  gi9c.ColorFull=WhiteFull
  gi9a.Intensity = AmberI

  gi10a.Color=Amber
  gi10a.ColorFull=AmberFull
  gi10b.Color=Amber
  gi10b.ColorFull=AmberFull
  gi10c.Color=White
  gi10c.ColorFull=WhiteFull
  gi10a.Intensity = AmberI

  gi11a.Color=Amber
  gi11a.ColorFull=AmberFull
  gi11a1.Color=Amber
  gi11a1.ColorFull=AmberFull
  gi11a2.Color=Amber
  gi11a2.ColorFull=AmberFull
  gi11b.Color=Amber
  gi11b.ColorFull=AmberFull
  gi11c.Color=White
  gi11c.ColorFull=WhiteFull
  gi11c1.Color=White
  gi11c1.ColorFull=WhiteFull
  gi11c2.Color=White
  gi11c2.ColorFull=WhiteFull
  gi11a.Intensity = AmberI
  gi11a1.Intensity = AmberI
  gi11a2.Intensity = AmberI

  gi11t21.Color=Amber
  gi11t21.ColorFull=AmberFull
  gi11t21.Intensity = AmberI
  gi11t20.Color=Amber
  gi11t20.ColorFull=AmberFull
  gi11t20.Intensity = AmberI
  gi11t19.Color=Amber
  gi11t19.ColorFull=AmberFull
  gi11t19.Intensity = AmberI
  gi11t18.Color=Amber
  gi11t18.ColorFull=AmberFull
  gi11t18.Intensity = AmberI
  gi11t17.Color=Amber
  gi11t17.ColorFull=AmberFull
  gi11t17.Intensity = AmberI

  gi12a.Color=Amber
  gi12a.ColorFull=AmberFull
  gi12a1.Color=Amber
  gi12a1.ColorFull=AmberFull
  gi12a2.Color=Amber
  gi12a2.ColorFull=AmberFull
  gi12a3.Color=Amber
  gi12a3.ColorFull=AmberFull
  gi12b.Color=Amber
  gi12b.ColorFull=AmberFull
  gi12c.Color=White
  gi12c.ColorFull=WhiteFull
  gi12c1.Color=White
  gi12c1.ColorFull=WhiteFull
  gi12c2.Color=White
  gi12c2.ColorFull=WhiteFull
  gi12a.Intensity = AmberI
  gi12a1.Intensity = AmberI
  gi12a2.Intensity = AmberI
  gi12a3.Intensity = AmberI

  gi12t21.Color=Amber
  gi12t21.ColorFull=AmberFull
  gi12t21.Intensity = AmberI
  gi12t20.Color=Amber
  gi12t20.ColorFull=AmberFull
  gi12t20.Intensity = AmberI
  gi12t19.Color=Amber
  gi12t19.ColorFull=AmberFull
  gi12t19.Intensity = AmberI
  gi12t18.Color=Amber
  gi12t18.ColorFull=AmberFull
  gi12t18.Intensity = AmberI
  gi12t17.Color=Amber
  gi12t17.ColorFull=AmberFull
  gi12t17.Intensity = AmberI

  gi13a.Color=Amber
  gi13a.ColorFull=AmberFull
  gi13b.Color=Amber
  gi13b.ColorFull=AmberFull
  gi13c.Color=White
  gi13c.ColorFull=WhiteFull
  gi13a.Intensity = AmberI

  gi14a.Color=Amber
  gi14a.ColorFull=AmberFull
  gi14b.Color=Amber
  gi14b.ColorFull=AmberFull
  gi14c.Color=White
  gi14c.ColorFull=WhiteFull
  gi14a.Intensity = AmberI

  gi16a.Color=Amber
  gi16a.ColorFull=AmberFull
  gi16b.Color=Amber
  gi16b.ColorFull=AmberFull
  gi16c.Color=White
  gi16c.ColorFull=WhiteFull
  gi16a.Intensity = AmberI

  gi17a.Color=Amber
  gi17a.ColorFull=AmberFull
  gi17b.Color=Amber
  gi17b.ColorFull=AmberFull
  gi17c.Color=White
  gi17c.ColorFull=WhiteFull
  gi17a.Intensity = AmberI

  gi18a.Color=Amber
  gi18a.ColorFull=AmberFull
  gi18b.Color=Amber
  gi18b.ColorFull=AmberFull
  gi18c.Color=White
  gi18c.ColorFull=WhiteFull
  gi18a.Intensity = AmberI

  gi19a.Color=Amber
  gi19a.ColorFull=AmberFull
  gi19b.Color=Amber
  gi19b.ColorFull=AmberFull
  gi19c.Color=White
  gi19c.ColorFull=WhiteFull
  gi19a.Intensity = AmberI

  gi20a.Color=Amber
  gi20a.ColorFull=AmberFull
  gi20b.Color=Amber
  gi20b.ColorFull=AmberFull
  gi20c.Color=White
  gi20c.ColorFull=WhiteFull
  gi20a.Intensity = AmberI

  gi21a.Color=Amber
  gi21a.ColorFull=AmberFull
  gi21b.Color=Amber
  gi21b.ColorFull=AmberFull
  gi21c.Color=White
  gi21c.ColorFull=WhiteFull
  gi21a.Intensity = AmberI

  gi22a.Color=Amber
  gi22a.ColorFull=AmberFull
  gi22b.Color=Amber
  gi22b.ColorFull=AmberFull
  gi22c.Color=White
  gi22c.ColorFull=WhiteFull
  gi22a.Intensity = AmberI

  gi23a.Color=Amber
  gi23a.ColorFull=AmberFull
  gi23b.Color=Amber
  gi23b.ColorFull=AmberFull
  gi23c.Color=White
  gi23c.ColorFull=WhiteFull
  gi23a.Intensity = AmberI

  gi24a.Color=Amber
  gi24a.ColorFull=AmberFull
  gi24b.Color=Amber
  gi24b.ColorFull=AmberFull
  gi24c.Color=White
  gi24c.ColorFull=WhiteFull
  gi24a.Intensity = AmberI

  gi25a.Color=Amber
  gi25a.ColorFull=AmberFull
  gi25b.Color=Amber
  gi25b.ColorFull=AmberFull
  gi25c.Color=White
  gi25c.ColorFull=WhiteFull
  gi25a.Intensity = AmberI

  gi26a.Color=Amber
  gi26a.ColorFull=AmberFull
  gi26b.Color=Amber
  gi26b.ColorFull=AmberFull
  gi26c.Color=White
  gi26c.ColorFull=WhiteFull
  gi26a.Intensity = AmberI


End If

If ColoredGIType = 2 Then


  gi1a.Color=Blue
  gi1a.ColorFull=BlueFull
  gi1b.Color=Blue
  gi1b.ColorFull=BlueFull
  gi1c.Color=Blue
  gi1c.ColorFull=BlueFull
  gi1a.Intensity = BlueI

  gi2a.Color=Blue
  gi2a.ColorFull=BlueFull
  gi2b.Color=Blue
  gi2b.ColorFull=BlueFull
  gi2c.Color=Blue
  gi2c.ColorFull=BlueFull
  gi2a.Intensity = BlueI

  gi3a.Color=Blue
  gi3a.ColorFull=BlueFull
  gi3b.Color=Blue
  gi3b.ColorFull=BlueFull
  gi3c.Color=Blue
  gi3c.ColorFull=BlueFull
  gi3a.Intensity = BlueI

  gi4a.Color=Blue
  gi4a.ColorFull=BlueFull
  gi4b.Color=Blue
  gi4b.ColorFull=BlueFull
  gi4c.Color=Blue
  gi4c.ColorFull=BlueFull
  gi4a.Intensity = BlueI

  gi5a.Color=Red
  gi5a.ColorFull=RedFull
  gi5b.Color=Red
  gi5b.ColorFull=RedFull
  gi5c.Color=Red
  gi5c.ColorFull=RedFull
  gi5a.Intensity = RedI

  gi6a.Color=Red
  gi6a.ColorFull=RedFull
  gi6b.Color=Red
  gi6b.ColorFull=RedFull
  gi6c.Color=Red
  gi6c.ColorFull=RedFull
  gi6a.Intensity = RedI

  gi7a.Color=Blue
  gi7a.ColorFull=BlueFull
  gi7b.Color=Blue
  gi7b.ColorFull=BlueFull
  gi7c.Color=Blue
  gi7c.ColorFull=BlueFull
  gi7a.Intensity = BlueI

  gi8a.Color=Yellow
  gi8a.ColorFull=YellowFull
  gi8b.Color=Yellow
  gi8b.ColorFull=YellowFull
  gi8c.Color=Yellow
  gi8c.ColorFull=YellowFull
  gi8a.Intensity = YellowI

  gi9a.Color=Blue
  gi9a.ColorFull=BlueFull
  gi9b.Color=Blue
  gi9b.ColorFull=BlueFull
  gi9c.Color=Blue
  gi9c.ColorFull=BlueFull
  gi9a.Intensity = BlueI

  gi10a.Color=Blue
  gi10a.ColorFull=BlueFull
  gi10b.Color=Blue
  gi10b.ColorFull=BlueFull
  gi10c.Color=Blue
  gi10c.ColorFull=BlueFull
  gi10a.Intensity = BlueI

  gi11a.Color=Yellow
  gi11a.ColorFull=YellowFull
  gi11a1.Color=Yellow
  gi11a1.ColorFull=YellowFull
  gi11a2.Color=Yellow
  gi11a2.ColorFull=YellowFull
  gi11b.Color=Yellow
  gi11b.ColorFull=YellowFull
  gi11c.Color=Yellow
  gi11c.ColorFull=YellowFull
  gi11c1.Color=Yellow
  gi11c1.ColorFull=YellowFull
  gi11c2.Color=Yellow
  gi11c2.ColorFull=YellowFull
  gi11a.Intensity = YellowI
  gi11a1.Intensity = YellowI
  gi11a2.Intensity = YellowI

  gi11t21.Color=Yellow
  gi11t21.ColorFull=YellowFull
  gi11t21.Intensity = YellowI
  gi11t20.Color=Yellow
  gi11t20.ColorFull=YellowFull
  gi11t20.Intensity = YellowI
  gi11t19.Color=Yellow
  gi11t19.ColorFull=YellowFull
  gi11t19.Intensity = YellowI
  gi11t18.Color=Yellow
  gi11t18.ColorFull=YellowFull
  gi11t18.Intensity = YellowI
  gi11t17.Color=Yellow
  gi11t17.ColorFull=YellowFull
  gi11t17.Intensity = YellowI

  gi12a.Color=Yellow
  gi12a.ColorFull=YellowFull
  gi12a1.Color=Yellow
  gi12a1.ColorFull=YellowFull
  gi12a2.Color=Yellow
  gi12a2.ColorFull=YellowFull
  gi12a3.Color=Yellow
  gi12a3.ColorFull=YellowFull
  gi12b.Color=Yellow
  gi12b.ColorFull=YellowFull
  gi12c.Color=Yellow
  gi12c.ColorFull=YellowFull
  gi12c1.Color=Yellow
  gi12c1.ColorFull=YellowFull
  gi12c2.Color=Yellow
  gi12c2.ColorFull=YellowFull
  gi12a.Intensity = YellowI
  gi12a1.Intensity = YellowI
  gi12a2.Intensity = YellowI
  gi12a3.Intensity = YellowI

  gi12t21.Color=Yellow
  gi12t21.ColorFull=YellowFull
  gi12t21.Intensity = YellowI
  gi12t20.Color=Yellow
  gi12t20.ColorFull=YellowFull
  gi12t20.Intensity = YellowI
  gi12t19.Color=Yellow
  gi12t19.ColorFull=YellowFull
  gi12t19.Intensity = YellowI
  gi12t18.Color=Yellow
  gi12t18.ColorFull=YellowFull
  gi12t18.Intensity = YellowI
  gi12t17.Color=Yellow
  gi12t17.ColorFull=YellowFull
  gi12t17.Intensity = YellowI

  gi13a.Color=Red
  gi13a.ColorFull=RedFull
  gi13b.Color=Red
  gi13b.ColorFull=RedFull
  gi13c.Color=Red
  gi13c.ColorFull=RedFull
  gi13a.Intensity = RedI

  gi14a.Color=Red
  gi14a.ColorFull=RedFull
  gi14b.Color=Red
  gi14b.ColorFull=RedFull
  gi14c.Color=Red
  gi14c.ColorFull=RedFull
  gi14a.Intensity = RedI

  gi16a.Color=Red
  gi16a.ColorFull=RedFull
  gi16b.Color=Red
  gi16b.ColorFull=RedFull
  gi16c.Color=Red
  gi16c.ColorFull=RedFull
  gi16a.Intensity = RedI

  gi17a.Color=Green
  gi17a.ColorFull=GreenFull
  gi17b.Color=Green
  gi17b.ColorFull=GreenFull
  gi17c.Color=Green
  gi17c.ColorFull=GreenFull
  gi17a.Intensity = GreenI

  gi18a.Color=Green
  gi18a.ColorFull=GreenFull
  gi18b.Color=Green
  gi18b.ColorFull=GreenFull
  gi18c.Color=Green
  gi18c.ColorFull=GreenFull
  gi18a.Intensity = GreenI

  gi19a.Color=Green
  gi19a.ColorFull=GreenFull
  gi19b.Color=Green
  gi19b.ColorFull=GreenFull
  gi19c.Color=Green
  gi19c.ColorFull=GreenFull
  gi19a.Intensity = GreenI

  gi20a.Color=Green
  gi20a.ColorFull=GreenFull
  gi20b.Color=Green
  gi20b.ColorFull=GreenFull
  gi20c.Color=Green
  gi20c.ColorFull=GreenFull
  gi20a.Intensity = GreenI

  gi21a.Color=Green
  gi21a.ColorFull=GreenFull
  gi21b.Color=Green
  gi21b.ColorFull=GreenFull
  gi21c.Color=Green
  gi21c.ColorFull=GreenFull
  gi21a.Intensity = GreenI

  gi22a.Color=Blue
  gi22a.ColorFull=BlueFull
  gi22b.Color=Blue
  gi22b.ColorFull=BlueFull
  gi22c.Color=Blue
  gi22c.ColorFull=BlueFull
  gi22a.Intensity = BlueI

  gi23a.Color=Blue
  gi23a.ColorFull=BlueFull
  gi23b.Color=Blue
  gi23b.ColorFull=BlueFull
  gi23c.Color=Blue
  gi23c.ColorFull=BlueFull
  gi23a.Intensity = BlueI

  gi24a.Color=Blue
  gi24a.ColorFull=BlueFull
  gi24b.Color=Blue
  gi24b.ColorFull=BlueFull
  gi24c.Color=Blue
  gi24c.ColorFull=BlueFull
  gi24a.Intensity = BlueI

  gi25a.Color=Blue
  gi25a.ColorFull=BlueFull
  gi25b.Color=Blue
  gi25b.ColorFull=BlueFull
  gi25c.Color=Blue
  gi25c.ColorFull=BlueFull
  gi25a.Intensity = BlueI

  gi26a.Color=Red
  gi26a.ColorFull=RedFull
  gi26b.Color=Red
  gi26b.ColorFull=RedFull
  gi26c.Color=Red
  gi26c.ColorFull=RedFull
  gi26a.Intensity = RedI

End If

If FlipperColor = 0 then
    FlipperColorType = Int(Rnd*9)+1
  Else
    FlipperColorType = FlipperColor
End If


If FlipperColorType = 1 Then
  RightFlipper.Material = "Plastic Yellow"
  RightFlipper.RubberMaterial = "Rubber Blue"
  pRightFlipperLogo.Material = "Plastic with an image"
  pRightFlipperLogo.Image = "DataEastFlipperLogo_texture"

  LeftFlipper.Material = "Plastic Yellow"
  LeftFlipper.RubberMaterial = "Rubber Blue"
  pLeftFlipperLogo.Material = "Plastic with an image"
  pLeftFlipperLogo.Image = "DataEastFlipperLogo_texture"
End If

If FlipperColorType = 2 Then
  RightFlipper.Material = "Plastic Yellow"
  RightFlipper.RubberMaterial = "Rubber Blue"
  pRightFlipperLogo.Material = "Plastic with an image"
  pRightFlipperLogo.Image = "DataEastFlipperLogoB_texture"

  LeftFlipper.Material = "Plastic Yellow"
  LeftFlipper.RubberMaterial = "Rubber Blue"
  pLeftFlipperLogo.Material = "Plastic with an image"
  pLeftFlipperLogo.Image = "DataEastFlipperLogoB_texture"
End If

If FlipperColorType = 3 Then
  RightFlipper.Material = "Plastic White"
  RightFlipper.RubberMaterial = "Rubber Black"
  pRightFlipperLogo.Material = "Plastic with an image"
  pRightFlipperLogo.Image = "DataEastFlipperLogo_texture"

  LeftFlipper.Material = "Plastic White"
  LeftFlipper.RubberMaterial = "Rubber Black"
  pLeftFlipperLogo.Material = "Plastic with an image"
  pLeftFlipperLogo.Image = "DataEastFlipperLogo_texture"
End If

If FlipperColorType = 4 Then
  RightFlipper.Material = "Plastic White"
  RightFlipper.RubberMaterial = "Rubber Black"
  pRightFlipperLogo.Material = "Plastic with an image"
  pRightFlipperLogo.Image = "DataEastFlipperLogoB_texture"

  LeftFlipper.Material = "Plastic White"
  LeftFlipper.RubberMaterial = "Rubber Black"
  pLeftFlipperLogo.Material = "Plastic with an image"
  pLeftFlipperLogo.Image = "DataEastFlipperLogoB_texture"
End If

If FlipperColorType = 5 Then
  RightFlipper.Material = "Plastic White"
  RightFlipper.RubberMaterial = "Rubber Black"
  pRightFlipperLogo.Material = "Plastic with an image"
  pRightFlipperLogo.Image = "DataEastFlipperLogoBl_texture"

  LeftFlipper.Material = "Plastic White"
  LeftFlipper.RubberMaterial = "Rubber Black"
  pLeftFlipperLogo.Material = "Plastic with an image"
  pLeftFlipperLogo.Image = "DataEastFlipperLogoBl_texture"
End If

If FlipperColorType = 6 Then
  RightFlipper.Material = "Plastic White"
  RightFlipper.RubberMaterial = "Rubber Red"
  pRightFlipperLogo.Material = "Plastic with an image"
  pRightFlipperLogo.Image = "DataEastFlipperLogo_texture"

  LeftFlipper.Material = "Plastic White"
  LeftFlipper.RubberMaterial = "Rubber Red"
  pLeftFlipperLogo.Material = "Plastic with an image"
  pLeftFlipperLogo.Image = "DataEastFlipperLogo_texture"
End If

If FlipperColorType = 7 Then
  RightFlipper.Material = "Plastic White"
  RightFlipper.RubberMaterial = "Rubber Red"
  pRightFlipperLogo.Material = "Plastic with an image"
  pRightFlipperLogo.Image = "DataEastFlipperLogoB_texture"

  LeftFlipper.Material = "Plastic White"
  LeftFlipper.RubberMaterial = "Rubber Red"
  pLeftFlipperLogo.Material = "Plastic with an image"
  pLeftFlipperLogo.Image = "DataEastFlipperLogoB_texture"
End If

If FlipperColorType = 8 Then
  RightFlipper.Material = "Plastic White"
  RightFlipper.RubberMaterial = "Rubber Yellow"
  pRightFlipperLogo.Material = "Plastic with an image"
  pRightFlipperLogo.Image = "DataEastFlipperLogoY_texture"

  LeftFlipper.Material = "Plastic White"
  LeftFlipper.RubberMaterial = "Rubber Yellow"
  pLeftFlipperLogo.Material = "Plastic with an image"
  pLeftFlipperLogo.Image = "DataEastFlipperLogoY_texture"
End If

If FlipperColorType = 9 Then
  RightFlipper.Material = "Plastic White"
  RightFlipper.RubberMaterial = "Rubber Yellow"
  pRightFlipperLogo.Material = "Plastic with an image"
  pRightFlipperLogo.Image = "DataEastFlipperLogoB_texture"

  LeftFlipper.Material = "Plastic White"
  LeftFlipper.RubberMaterial = "Rubber Yellow"
  pLeftFlipperLogo.Material = "Plastic with an image"
  pLeftFlipperLogo.Image = "DataEastFlipperLogoB_texture"
End If



If WBRBulb = 1 then
  flasher16a.imageA = "HaloBlue"
  flasher16a.imageb = "HaloBlue"
  flasher16a.color = rgb(128,255,255)
Else
  flasher16a.imageA = "HaloRed"
  flasher16a.imageB = "HaloRed"
  flasher16a.color = rgb(255,79,79)
End If

If RRBulb = 1 then
  flasher16b.imageA = "HaloBlue"
  flasher16b.imageb = "HaloBlue"
  flasher16b.color = rgb(128,255,255)
Else
  flasher16b.imageA = "HaloRed"
  flasher16b.imageB = "HaloRed"
  flasher16b.color = rgb(255,79,79)
End If


If CustomSideWalls = 0 then
  CustomSideWallsType = 0
Else
  CustomSideWallsType = CustomSideWalls
End If

If CustomSideWallsType = 0 Then
  pSideWalls.Visible = 0
End If

If CustomSideWallsType = 1 Then
  pSideWalls.Image = "sidewalls_blue"
End If

If CustomSideWallsType = 2 Then
  pSideWalls.Image = "sidewalls_custom"
End If

If Instruction_Cards = 0 then
    Instruction_CardsType = Int(Rnd*4)+1
  Else
    Instruction_CardsType = Instruction_Cards
End If

If Instruction_CardsType = 1 then
  IC_Left.Image = "instructioncard"
  IC_Right.Image = "dataeast-coin"
End If

If Instruction_CardsType = 2 then
  IC_Left.Image = "instructioncard"
  IC_Right.Image = "dataeast-freeplay"
End If

If Instruction_CardsType = 3 then
  IC_Left.Image = "RaB_IC3_L"
  IC_Right.Image = "RaB_IC3_R"
End If

If Instruction_CardsType = 4 then
  IC_Left.Image = "RaB_IC3_L"
  IC_Right.Image = "RaB_IC4_R"
End If


If RubberColor  = 0 then
    RubberColorType = Int(Rnd*3)+1
  Else
    RubberColorType = RubberColor
End If

If RubberColorType = 1 Then
    for each xxRubbers in AllRubbers
      xxRubbers.material = "Rubber White"
      next
End If

If RubberColorType = 2 Then
    for each xxRubbers in AllRubbers
      xxRubbers.material = "Rubber Black"
      next
End If

If RubberColorType = 3 Then
    for each xxRubbers in AllRubbers
      xxRubbers.material = "Rubber Blue"
      next
End If




End Sub

'AutoPlunger
Sub SolAutoPlungerIM(Enabled)
    If Enabled Then
    PlaySoundAt "BallRelease1", Primitive23
        PlungerIM.AutoFire
    End If
End Sub

'''''''''''''''''''''''''''
''''''' Diverters
'''''''''''''''''''''''''''

' Ramp Diverter

 Sub SolDiverter2(enabled)
  if enabled then
    RampDiverterDirection = 0
    RotateRampDiverterStep = 4
    RotateRampDiverter.Enabled = 1
    Diverter2a.isdropped=false
    Diverter2b.isdropped=True
    PlaySoundAt"ramp_diverter_open", pRampDiverter
  else
    RampDiverterDirection = 1
    RotateRampDiverterStep = 2
    RotateRampDiverter.Enabled = 1
    Diverter2a.isdropped=true
    Diverter2b.isdropped=false
    PlaySoundAt"ramp_diverter_close", pRampDiverter
  end if
 end sub


Dim RotateRampDiverterStep, RampDiverterDirection

Sub RotateRampDiverter_timer()


  Select Case RotateRampDiverterStep
        Case 0: pRampDiverter.ObjRotZ = -30:RotateRampDiverter.Enabled = 0
    Case 1: pRampDiverter.ObjRotZ = -26
        Case 2: pRampDiverter.ObjRotZ = -24
        Case 3: pRampDiverter.ObjRotZ = -12
        Case 4: pRampDiverter.ObjRotZ = -2
        Case 5: pRampDiverter.ObjRotZ = 0
        Case 6: pRampDiverter.ObjRotZ = -2:RotateRampDiverter.Enabled = 0
     End Select

  If RampDiverterDirection = 1 then
    RotateRampDiverterStep = RotateRampDiverterStep + 1
  Else
    RotateRampDiverterStep = RotateRampDiverterStep - 1
  End If
End Sub


''''''Top Gate1

Dim TopGateOpen

Sub TopGate(Enabled)
    If Enabled Then
    Gate2.Collidable = false
    If TopGateOpen = 0 then PlaySoundAt "topgate_close", Gate2
    TopGateOpen = 1
      Else
    Gate2.Collidable = true
    If TopGateOpen = 1 then PlaySoundAt "topgate_open", Gate2
    TopGateOpen = 0
  End If
End Sub


'Playfield Diverter

 Sub SolDiverter1(enabled)
  if enabled then
    Diverter1a.isdropped=True
    Diverter1b.isdropped=False
    RotatePFDiverterStep = 1
    PFDiverterDirection = 2
    RotatePFDiverter.Enabled = 1
    PlaySoundAt "playfield_diverter_open", pPFDiverter
  else
    Diverter1a.isdropped=False
    Diverter1b.isdropped=True
    RotatePFDiverterStep = 4
    PFDiverterDirection = 0
    RotatePFDiverter.Enabled = 1
    PlaySoundAt "playfield_diverter_close", pPFDiverter
  end if
 end sub

Dim RotatePFDiverterStep, PFDiverterDirection

Sub RotatePFDiverter_timer()


  Select Case RotatePFDiverterStep
        Case 0: pPFDiverter.ObjRotZ = -2:RotatePFDiverter.Enabled = 0
    Case 1: pPFDiverter.ObjRotZ = -4
        Case 2: pPFDiverter.ObjRotZ = -1
        Case 3: pPFDiverter.ObjRotZ = 11
    Case 4: pPFDiverter.ObjRotZ = 26
    Case 5: pPFDiverter.ObjRotZ = 28
    Case 6: pPFDiverter.ObjRotZ = 26:RotatePFDiverter.Enabled = 0
     End Select

  If PFDiverterDirection = 1 then
    RotatePFDiverterStep = RotatePFDiverterStep + 1
  Else
    RotatePFDiverterStep = RotatePFDiverterStep - 1
  End If
End Sub


'Nell Animation
Dim NellPos, NellDir, NellMoving, NellBackDir
NellPos = 150
NellDir = -.95
NellBackDir = 10

Sub SolNell(Enabled)
    If Enabled Then
    NellMoving = 1
    MoveNell.Enabled = True
    RotateSaw.Enabled = True
    PlaySoundAt "sawmoter", pSawBlade
  Else
    If NellMoving = 1 then
      MoveNell.enabled = false
      MoveNellBack.enabled = True
    End If
  End If
End Sub

Sub MoveNell_timer()

  NellPos = NellPos + NellDir

  pLog.TransX = NellPos
  pNellPlastic.TransX = NellPos
  pNellrivit.TransX = NellPos

End Sub


Sub MoveNellBack_timer()
  NellPos = NellPos + NellBackDir

  If NellPos >= 150 Then NellPos = 150:MoveNellBack.enabled = false:RotateSaw.Enabled = False:NellMoving = 0'PlaySoundAt"sawmoter",pSawBlade

  pLog.TransX = NellPos
  pNellPlastic.TransX = NellPos
  pNellrivit.TransX = NellPos
End Sub


Dim RotateSawStep

Sub RotateSaw_timer()
  Select Case RotateSawStep
    Case 1: pSawBlade.RotY = pSawBlade.RotY + 5:
        Case 2: pSawBlade.RotY = pSawBlade.RotY + 5:
        Case 3: pSawBlade.RotY = pSawBlade.RotY + 5:
        Case 4: pSawBlade.RotY = pSawBlade.RotY + 5:RotateSawStep = 1
     End Select
  RotateSawStep = RotateSawStep + 1
End Sub



''''''''''''''''''''''''''''''''
''  Flippers
''''''''''''''''''''''''''''''''

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
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub


'''''''''''''''''''''''''''''''''''''
'''''  Rotate Primitives
'''''''''''''''''''''''''''''''''''''
Sub RotatePrimitives_Timer()
  pLeftFlipperLogo.Roty = LeftFlipper.Currentangle + 180
  pRightFlipperLogo.Roty = RightFlipper.Currentangle + 180
    p_gate3.Rotx = gate3.CurrentAngle' + 90
    p_gate4.Rotx = Gate4.CurrentAngle' + 90
    p_gate5.Rotx = Gate5.CurrentAngle' + 90
End Sub

'''''''''''''Switches

Sub swPlunger2_Hit()
  SwitchPlungerdir = 1
  SwPlungerMove = 1
  Me.TimerEnabled = true
  PlaySoundAt "sensor",swPlunger2
End Sub

Sub swPlunger2_unHit()
  SwitchPlungerdir = -1
  SwPlungerMove = 5
  Me.TimerEnabled = true
End Sub

Sub sw31_Hit()
  Switch31dir = 1
  Sw31Move = 1
  Me.TimerEnabled = true
  Controller.Switch(31) = 1
End Sub

Sub sw31_unHit()
  Switch31dir = -1
  Sw31Move = 5
  Me.TimerEnabled = true
  Controller.Switch(31) = 0
End Sub

Sub sw32_Hit()
  Switch32dir = 1
  Sw32Move = 1
  Me.TimerEnabled = true
  Controller.Switch(32) = 1
End Sub

Sub sw32_unHit()
  Switch32dir = -1
  Sw32Move = 5
  Me.TimerEnabled = true
  Controller.Switch(32) = 0
End Sub

Sub Sw33_Hit()Controller.Switch(33)=1:pLeftGate_switch.ObjRotX = -5: End Sub
Sub Sw33_UnHit()Controller.Switch(33)=0:pLeftGate_switch.ObjRotX = 0: End Sub

Sub sw34_Hit()
  Switch34dir = 1
  Sw34Move = 1
  Me.TimerEnabled = true
  Controller.Switch(34) = 1
End Sub

Sub sw34_unHit()
  Switch34dir = -1
  Sw34Move = 5
  Me.TimerEnabled = true
  Controller.Switch(34) = 0
End Sub

Sub sw35_Hit()
  Switch35dir = 1
  Sw35Move = 1
  Me.TimerEnabled = true
  Controller.Switch(35) = 1
End Sub

Sub sw35_unHit()
  Switch35dir = -1
  Sw35Move = 5
  Me.TimerEnabled = true
  Controller.Switch(35) = 0
End Sub

Sub sw36_Hit()
  Switch36dir = 1
  Sw36Move = 1
  Me.TimerEnabled = true
  Controller.Switch(36) = 1
End Sub

Sub sw36_unHit()
  Switch36dir = -1
  Sw36Move = 5
  Me.TimerEnabled = true
  Controller.Switch(36) = 0
End Sub

Sub sw39_Hit()
  Switch39dir = 1
  Sw39Move = 1
  Me.TimerEnabled = true
  Controller.Switch(39) = 1
End Sub

Sub sw39_unHit()
  Switch39dir = -1
  Sw39Move = 5
  Me.TimerEnabled = true
  Controller.Switch(39) = 0
End Sub

Sub sw40_Hit()
  Switch40dir = 1
  Sw40Move = 1
  Me.TimerEnabled = true
  Controller.Switch(40) = 1
End Sub

Sub sw40_unHit()
  Switch40dir = -1
  Sw40Move = 5
  Me.TimerEnabled = true
  Controller.Switch(40) = 0
End Sub

Sub sw41_Hit()
  Switch41dir = 1
  Sw41Move = 1
  Me.TimerEnabled = true
  Controller.Switch(41) = 1
End Sub

Sub sw41_unHit()
  Switch41dir = -1
  Sw41Move = 5
  Me.TimerEnabled = true
  Controller.Switch(41) = 0
End Sub

Sub sw42_Hit()
  Switch42dir = 1
  Sw42Move = 1
  Me.TimerEnabled = true
  Controller.Switch(42) = 1
End Sub

Sub sw42_unHit()
  Switch42dir = -1
  Sw42Move = 5
  Me.TimerEnabled = true
  Controller.Switch(42) = 0
End Sub

Sub sw43_Hit()
  Switch43dir = 1
  Sw43Move = 1
  Me.TimerEnabled = true
  Controller.Switch(43) = 1
End Sub

Sub sw43_unHit()
  Switch43dir = -1
  Sw43Move = 5
  Me.TimerEnabled = true
  Controller.Switch(43) = 0
End Sub

' Rollover Animations

Dim SwitchPlungerdir, SWPlungerMove

Sub swPlunger2_timer()
  Select case SwPlungerMove
    Case 0:me.TimerEnabled = false:pRollover11.RotX = 90
    Case 1:pRollover11.RotX = 95
    Case 2:pRollover11.RotX = 100
    Case 3:pRollover11.RotX = 105
    Case 4:pRollover11.RotX = 110
    Case 5:pRollover11.RotX = 115
    Case 6:me.TimerEnabled = false:pRollover11.RotX = 120
  End Select
  SWPlungerMove = SWPlungerMove + SwitchPlungerdir
End Sub

Dim Switch31dir, SW31Move

Sub sw31_timer()
Select case Sw31Move
  Case 0:me.TimerEnabled = false:pRollover3.RotX = 90
  Case 1:pRollover3.RotX = 95
  Case 2:pRollover3.RotX = 100
  Case 3:pRollover3.RotX = 105
  Case 4:pRollover3.RotX = 110
  Case 5:pRollover3.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover3.RotX = 120
End Select
SW31Move = SW31Move + Switch31dir
End Sub


Dim Switch32dir, SW32Move

Sub sw32_timer()
Select case Sw32Move
  Case 0:me.TimerEnabled = false:pRollover4.RotX = 90
  Case 1:pRollover4.RotX = 95
  Case 2:pRollover4.RotX = 100
  Case 3:pRollover4.RotX = 105
  Case 4:pRollover4.RotX = 110
  Case 5:pRollover4.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover4.RotX = 120
End Select
SW32Move = SW32Move + Switch32dir
End Sub

Dim Switch34dir, SW34Move

Sub sw34_timer()
Select case Sw34Move
  Case 0:me.TimerEnabled = false:pRollover5.RotX = 90
  Case 1:pRollover5.RotX = 95
  Case 2:pRollover5.RotX = 100
  Case 3:pRollover5.RotX = 105
  Case 4:pRollover5.RotX = 110
  Case 5:pRollover5.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover5.RotX = 120
End Select
SW34Move = SW34Move + Switch34dir
End Sub


Dim Switch35dir, SW35Move

Sub sw35_timer()
Select case Sw35Move
  Case 0:me.TimerEnabled = false:pRollover9.RotX = 90
  Case 1:pRollover9.RotX = 95
  Case 2:pRollover9.RotX = 100
  Case 3:pRollover9.RotX = 105
  Case 4:pRollover9.RotX = 110
  Case 5:pRollover9.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover9.RotX = 120
End Select
SW35Move = SW35Move + Switch35dir
End Sub



Dim Switch36dir, SW36Move

Sub sw36_timer()
Select case Sw36Move
  Case 0:me.TimerEnabled = false:pRollover10.RotX = 90
  Case 1:pRollover10.RotX = 95
  Case 2:pRollover10.RotX = 100
  Case 3:pRollover10.RotX = 105
  Case 4:pRollover10.RotX = 110
  Case 5:pRollover10.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover10.RotX = 120
End Select
SW36Move = SW36Move + Switch36dir
End Sub


Dim Switch39dir, SW39Move

Sub sw39_timer()
Select case Sw39Move
  Case 0:me.TimerEnabled = false:pRollover2.RotX = 90
  Case 1:pRollover2.RotX = 95
  Case 2:pRollover2.RotX = 100
  Case 3:pRollover2.RotX = 105
  Case 4:pRollover2.RotX = 110
  Case 5:pRollover2.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover2.RotX = 120
End Select
SW39Move = SW39Move + Switch39dir
End Sub


Dim Switch40dir, SW40Move

Sub sw40_timer()
Select case Sw40Move
  Case 0:me.TimerEnabled = false:pRollover1.RotX = 90
  Case 1:pRollover1.RotX = 95
  Case 2:pRollover1.RotX = 100
  Case 3:pRollover1.RotX = 105
  Case 4:pRollover1.RotX = 110
  Case 5:pRollover1.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover1.RotX = 120
End Select
SW40Move = SW40Move + Switch40dir
End Sub


Dim Switch41dir, SW41Move

Sub sw41_timer()
Select case Sw41Move
  Case 0:me.TimerEnabled = false:pRollover6.RotX = 90
  Case 1:pRollover6.RotX = 95
  Case 2:pRollover6.RotX = 100
  Case 3:pRollover6.RotX = 105
  Case 4:pRollover6.RotX = 110
  Case 5:pRollover6.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover6.RotX = 120
End Select
SW41Move = SW41Move + Switch41dir
End Sub


Dim Switch42dir, SW42Move

Sub sw42_timer()
Select case Sw42Move
  Case 0:me.TimerEnabled = false:pRollover7.RotX = 90
  Case 1:pRollover7.RotX = 95
  Case 2:pRollover7.RotX = 100
  Case 3:pRollover7.RotX = 105
  Case 4:pRollover7.RotX = 110
  Case 5:pRollover7.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover7.RotX = 120
End Select
SW42Move = SW42Move + Switch42dir
End Sub


Dim Switch43dir, SW43Move

Sub sw43_timer()
Select case Sw40Move
  Case 0:me.TimerEnabled = false:pRollover8.RotX = 90
  Case 1:pRollover8.RotX = 95
  Case 2:pRollover8.RotX = 100
  Case 3:pRollover8.RotX = 105
  Case 4:pRollover8.RotX = 110
  Case 5:pRollover8.RotX = 115
  Case 6:me.TimerEnabled = false:pRollover8.RotX = 120
End Select
SW43Move = SW43Move + Switch43dir
End Sub



''''''''''Kicker trough

Dim EnteredFromTop

Sub SW37_Hit():Controller.Switch(37) = 1:End Sub
Sub SW37_unHit():Controller.Switch(37) = 0:End Sub

Sub sw29_Hit()
  Controller.Switch(29) = 1
  rTroughW1.IsDropped = false
  If EnteredFromTop = 1 then

  Else
    PlaySoundAt "SubwayEnterLower", sw29
  End If
End Sub

Sub sw29kick(Enabled)
  sw29.timerenabled = 1
  PlaySoundAt "BallRelease1", Primitive19
  sw29.Kick 0,45,1.56 '50=Strength
  Controller.Switch(29) = 0
  EnteredFromTop = 0
  rTroughW1.IsDropped = true
End Sub

Dim sw29step

Sub sw29_timer()
  Select Case sw29step
    Case 0:
    Case 1:
    Case 2:
    Case 3:sw29.timerEnabled = 0:sw29step = 0
  End Select
  sw29step = sw29step + 1
End Sub


Sub sw37a_hit():RandomTopSubwayEnterSound:EnteredFromTop = 1:End Sub

' TroughHelpers

Sub rTroughT1_hit ()
  rTroughW1.IsDropped = false
End Sub
Sub rTroughT2_hit ()
  rTroughW2.IsDropped = false
End Sub
Sub rTroughT2_unhit ()
  rTroughW2.IsDropped = true
End Sub

' LaserKick

Dim LKStep

Sub AutoPlungeLK(Enabled)
  If Enabled Then
    LaserKick.Enabled=True
  Else
    LaserKick.Enabled=False
  End If
End Sub

Sub LaserKick_Hit:Me.Kick 0,45: PlaySoundAt "laserkick", pLaserKick:LKStep = 0:pLaserKick.TransY = 75:Me.TimerEnabled = true:End Sub  '45=strength

Sub LaserKick_timer()
  Select Case LKStep
    Case 0:pLaserKick.TransY = 50
    Case 1:pLaserKick.TransY = 25
    Case 2:pLaserKick.TransY = 0:me.TimerEnabled = false
  End Select
  LKStep = LKStep + 1
End Sub


'''''''''''''''''''''''''''
''''Hat Trick
'''''''''''''''''''''''''''
Dim HatTrickEnabled

Sub HatTrickLion(Enabled)
  If Enabled then
    HatTrickEnabled = 1
    PlaySound "ht_lion_up"
    BBTimerUp.Enabled = 1
  Else
    If HatTrickEnabled = 1 then
      PlaySound "ht_lion_down"
      BBTimerDown.Enabled = 1
    End If
  End If

End Sub

Sub HatTrickRocky(Enabled)
  If Enabled then
    HatTrickEnabled = 2
    PlaySound "ht_rocky_up"
    BBTimerUp.Enabled = 1
  Else
    If HatTrickEnabled = 2 then
      PlaySound "ht_rocky_down"
      BBTimerDown.Enabled = 1
    End If
  End If

End Sub

Sub HatTrickRhino(Enabled)
  If Enabled then
    HatTrickEnabled = 3
    PlaySound "ht_rhino_up"
    BBTimerUp.Enabled = 1
  Else
    If HatTrickEnabled = 3 then
      PlaySound "ht_rhino_down"
      BBTimerDown.Enabled = 1
    End If

  End If

End Sub

Sub BBTimerUp_Timer
  If HatTrickEnabled = 1 Then
    If BBLion.Z <143 Then
      BBLion.Z = BBLion.Z + 35.883125
    Else
      BBTimerUp.Enabled = 0
    End If
  End If
  If HatTrickEnabled = 2 Then
    If BBRocky.Z <143 Then
      BBRocky.Z = BBRocky.Z + 35.883125
    Else
      BBTimerUp.Enabled = 0
    End If
  End If
  If HatTrickEnabled = 3 Then
    If BBRhino.Z <143 Then
      BBRhino.Z = BBRhino.Z + 35.883125
    Else
      BBTimerUp.Enabled = 0
    End If
  End If
  If BBArm.RotY < 25 Then
    BBArm.RotY = BBArm.RotY + 7.5
  End If
End Sub

Sub BBTimerDown_timer
  If HatTrickEnabled = 1 Then
    If BBLion.Z > 42 Then
      BBLion.Z = BBLion.Z - 35.883125
    Else
      BBTimerDown.Enabled = 0
    End If
  End If
  If HatTrickEnabled = 2 Then
    If BBRocky.Z > 42 Then
      BBRocky.Z = BBRocky.Z - 35.883125
    Else
      BBTimerDown.Enabled = 0
    End If
  End If
  If HatTrickEnabled = 3 Then
    If BBRhino.Z > 42 Then
      BBRhino.Z = BBRhino.Z - 35.883125
    Else
      BBTimerDown.Enabled = 0
    End If
  End If
  If BBArm.RotY > -5 Then
    BBArm.RotY = BBArm.RotY - 7.5
  End If
End Sub


'Hat Trick up Kick
Sub sw52_Hit():PlaySoundAt "khattrick_kick_enter", sw52:End Sub

Sub KickBallUp(Enabled)
  sw52.timerenabled = 1
  sw52.Kick 0,50,1.56
End Sub

Dim sw52step

Sub sw52_timer()
  Select Case sw52step
    Case 0:pUpKicker.TransY = 10:PlaySoundAt "hattrickkick", pUpKicker
    Case 1:pUpKicker.TransY = 20
    Case 2:pUpKicker.TransY = 30
    Case 3:
    Case 4:
    Case 5:pUpKicker.TransY = 25
    Case 6:pUpKicker.TransY = 20
    Case 7:pUpKicker.TransY = 15: Controller.Switch(52) = 0:HatTrickEnabled = 0:sw52.enabled = false
    Case 8:pUpKicker.TransY = 10
    Case 9:pUpKicker.TransY = 5
    Case 10:pUpKicker.TransY = 0:sw52.timerEnabled = 0:sw52step = 0
  End Select
  sw52step = sw52step + 1
End Sub


'''''''''''''''''''''''''''''
''''''''' Targets
'''''''''''''''''''''''''''''

Dim Target22Step, Target23Step, Target24Step, Target25Step, Target26Step, Target27Step, Target28Step


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'   Drop Targets
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

  dim sw17Dir, sw18Dir, sw19Dir, sw20Dir, sw21Dir
  dim sw17Pos, sw18Pos, sw19Pos, sw20Pos, sw21Pos
  Dim sw17step, sw18step, sw19step, sw20step, sw21step
  Dim target17, target18, target19, target20, target21

  sw17Dir = 1:sw18Dir = 1:sw19Dir = 1:sw20Dir = 1:sw21Dir = 1
  sw17Pos = 0:sw18Pos = 0:sw19Pos = 0:sw20Pos = 0:sw21Pos = 0

'Targets Init
  t17.TimerEnabled = 1:t18.timerEnabled = 1:t19.TimerEnabled = 1:t20.timerEnabled = 1:t21.TimerEnabled = 1

  Sub DoubleDrop1_HIt:t17.timerenabled = True:t18.timerenabled = True: End Sub
  Sub DoubleDrop2_HIt:t18.timerenabled = True:t19.timerenabled = True: End Sub
  Sub DoubleDrop3_HIt:t19.timerenabled = True:t20.timerenabled = True: End Sub
  Sub DoubleDrop4_HIt:t20.timerenabled = True:t21.timerenabled = True: End Sub


  Sub t17_Hit:me.timerEnabled = 1:End Sub
  Sub t18_Hit:me.timerEnabled = 1:End Sub
  Sub t19_Hit:me.timerEnabled = 1:End Sub
  Sub t20_Hit:me.timerEnabled = 1:End Sub
  Sub t21_Hit:me.timerEnabled = 1:End Sub


'''Target animation


Sub t17_timer()
  Select Case sw17step
    Case 0:
    Case 1:pT17.RotX = 2
    Case 2:pT17.RotX = 5
    Case 3:dtDrop.Hit 1:sw17Dir = 0:t17a.Enabled = 1:DoubleDrop1.isDropped = True:t17.isDropped = 1:target17 = 1
    Case 4:pT17.RotX = 3:PlaySoundAt "drop_hit",t17
    Case 5:pT17.RotX = 0:me.timerEnabled = 0:sw17step = 0
  End Select
  sw17step = sw17step + 1
End Sub


 Sub t17a_Timer()
  Select Case sw17Pos
        Case 0: pT17.TransZ=0
         If sw17Dir = 1 then
          t17a.Enabled = 0
         else
           end if
        Case 1: pT17.TransZ=0
        Case 2: pT17.TransZ=-6
        Case 3: pT17.TransZ=-8
        Case 4: pT17.TransZ=-18
        Case 5: pT17.TransZ=-24
        Case 6: pT17.TransZ=-30
        Case 7: pT17.TransZ=-36
        Case 8: pT17.TransZ=-42
        Case 9: pT17.TransZ=-48
        Case 10: pT17.TransZ=-55
         If sw17Dir = 1 then
         else
          t17a.Enabled = 0
           end if
  End Select
  If sw17Dir = 1 then
    If sw17pos>0 then sw17pos=sw17pos-1
  else
    If sw17pos<10 then sw17pos=sw17pos+1
  end if
End Sub


Sub t18_timer()
  Select Case sw18step
    Case 0:
    Case 1:pT18.RotX = 2
    Case 2:pT18.RotX = 5
    Case 3:dtDrop.Hit 2:sw18Dir = 0:t18a.Enabled = 1:DoubleDrop1.isDropped = True:DoubleDrop2.isDropped = True:t18.isDropped = 1:target18 = 1
    Case 4:pT18.RotX = 3:PlaySoundAt "drop_hit",t18
    Case 5:pT18.RotX = 0:me.timerEnabled = 0:sw18step = 0
  End Select
  sw18step = sw18step + 1
End Sub


 Sub t18a_Timer()
  Select Case sw18Pos
        Case 0: pT18.TransZ=0
         If sw18Dir = 1 then
          t18a.Enabled = 0
         else
           end if
        Case 1: pT18.TransZ=0
        Case 2: pT18.TransZ=-6
        Case 3: pT18.TransZ=-8
        Case 4: pT18.TransZ=-18
        Case 5: pT18.TransZ=-24
        Case 6: pT18.TransZ=-30
        Case 7: pT18.TransZ=-36
        Case 8: pT18.TransZ=-42
        Case 9: pT18.TransZ=-48
        Case 10: pT18.TransZ=-55
         If sw18Dir = 1 then
         else
          t18a.Enabled = 0
           end if
  End Select
  If sw18Dir = 1 then
    If sw18pos>0 then sw18pos=sw18pos-1
  else
    If sw18pos<10 then sw18pos=sw18pos+1
  end if
End Sub

Sub t19_timer()
  Select Case sw19step
    Case 0:
    Case 1:pT19.RotX = 2
    Case 2:pT19.RotX = 5
    Case 3:dtDrop.Hit 3:sw19Dir = 0:t19a.Enabled = 1:DoubleDrop2.isDropped = True:DoubleDrop3.isDropped = True:t19.isDropped = 1:target19 = 1
    Case 4:pT19.RotX = 3:PlaySoundAt "drop_hit",t19
    Case 5:pT19.RotX = 0:me.timerEnabled = 0:sw19step = 0
  End Select
  sw19step = sw19step + 1
End Sub


 Sub t19a_Timer()
  Select Case sw19Pos
        Case 0: pT19.TransZ=0
         If sw19Dir = 1 then
          t19a.Enabled = 0
         else
           end if
        Case 1: pT19.TransZ=0
        Case 2: pT19.TransZ=-6
        Case 3: pT19.TransZ=-8
        Case 4: pT19.TransZ=-18
        Case 5: pT19.TransZ=-24
        Case 6: pT19.TransZ=-30
        Case 7: pT19.TransZ=-36
        Case 8: pT19.TransZ=-42
        Case 9: pT19.TransZ=-48
        Case 10: pT19.TransZ=-55
         If sw19Dir = 1 then
         else
          t19a.Enabled = 0
           end if
  End Select
  If sw19Dir = 1 then
    If sw19pos>0 then sw19pos=sw19pos-1
  else
    If sw19pos<10 then sw19pos=sw19pos+1
  end if
End Sub


Sub t20_timer()
  Select Case sw20step
    Case 0:
    Case 1:pT20.RotX = 2
    Case 2:pT20.RotX = 5
    Case 3:dtDrop.Hit 4:sw20Dir = 0:t20a.Enabled = 1:DoubleDrop3.isDropped = True:DoubleDrop4.isDropped = True:t20.isDropped = 1:target20 = 1
    Case 4:pT20.RotX = 3:PlaySoundAt "drop_hit",t20
    Case 5:pT20.RotX = 0:me.timerEnabled = 0:sw20step = 0
  End Select
  sw20step = sw20step + 1
End Sub


 Sub t20a_Timer()
  Select Case sw20Pos
        Case 0: pT20.TransZ=0
         If sw20Dir = 1 then
          t20a.Enabled = 0
         else
           end if
        Case 1: pT20.TransZ=0
        Case 2: pT20.TransZ=-6
        Case 3: pT20.TransZ=-8
        Case 4: pT20.TransZ=-18
        Case 5: pT20.TransZ=-24
        Case 6: pT20.TransZ=-30
        Case 7: pT20.TransZ=-36
        Case 8: pT20.TransZ=-42
        Case 9: pT20.TransZ=-48
        Case 10: pT20.TransZ=-55
         If sw20Dir = 1 then
         else
          t20a.Enabled = 0
           end if
  End Select
  If sw20Dir = 1 then
    If sw20pos>0 then sw20pos=sw20pos-1
  else
    If sw20pos<10 then sw20pos=sw20pos+1
  end if
End Sub


Sub t21_timer()
  Select Case sw21step
    Case 0:
    Case 1:pT21.RotX = 2
    Case 2:pT21.RotX = 5
    Case 3:dtDrop.Hit 5:sw21Dir = 0:t21a.Enabled = 1:DoubleDrop4.isDropped = True:t21.isDropped = 1:target21 = 1
    Case 4:pT21.RotX = 3:PlaySoundAt "drop_hit",t21
    Case 5:pT21.RotX = 0:me.timerEnabled = 0:sw21step = 0
  End Select
  sw21step = sw21step + 1
End Sub


 Sub t21a_Timer()
  Select Case sw21Pos
        Case 0: pT21.TransZ=0
         If sw21Dir = 1 then
          t21a.Enabled = 0
         else
           end if
        Case 1: pT21.TransZ=0
        Case 2: pT21.TransZ=-6
        Case 3: pT21.TransZ=-8
        Case 4: pT21.TransZ=-18
        Case 5: pT21.TransZ=-24
        Case 6: pT21.TransZ=-30
        Case 7: pT21.TransZ=-36
        Case 8: pT21.TransZ=-42
        Case 9: pT21.TransZ=-48
        Case 10: pT21.TransZ=-55
         If sw21Dir = 1 then
         else
          t21a.Enabled = 0
           end if
  End Select
  If sw21Dir = 1 then
    If sw21pos>0 then sw21pos=sw21pos-1
  else
    If sw21pos<10 then sw21pos=sw21pos+1
  end if
End Sub



'DT Subs
   Sub DropsUp(Enabled)
    If Enabled Then
      PlaySoundAt "DropsUp",t19
      sw17Dir = 1:sw18Dir = 1:sw19Dir = 1:sw20Dir = 1:sw21Dir = 1
      t17.isDropped = 0:t18.isDropped = 0:t19.isDropped = 0:t20.isDropped = 0:t21.isDropped = 0
      target17 = 0:target18 = 0:target19 = 0:target20 = 0:target21 = 0
      t17a.Enabled = 1:t18a.Enabled = 1:t19a.Enabled = 1:t20a.Enabled = 1:t21a.Enabled = 1:DoubleDrop1.isDropped = False:DoubleDrop2.isDropped = False
      dtDrop.DropSol_On
    End if
   End Sub


Sub CheckDropShadows()

' 'gion target dropped
  If LampState(200) = 1 and target17 = 1 then
    gi12t17.state = 1
    gi12ct17.state = 1
    gi11t17.state = 1
  Else
    gi12t17.state = 0
    gi12ct17.state = 0
    gi11t17.state = 0
  End If

  If LampState(200) = 1 and target18  = 1 then
    gi12t18.state = 1
    gi12ct18.state = 1
    gi11t18.state = 1
  Else
    gi12t18.state = 0
    gi12ct18.state = 0
    gi11t18.state = 0
  End If

  If LampState(200) = 1 and target19 = 1 then
    gi12t19.state = 1
    gi12ct19.state = 1
    gi11t19.state = 1
    gi11ct19.state = 1
  Else
    gi12t19.state = 0
    gi12ct19.state = 0
    gi11t19.state = 0
    gi11ct19.state = 0
  End If

  If LampState(200) = 1 and target20 = 1 then
    gi12t20.state = 1
    gi11t20.state = 1
    gi11ct20.state = 1
  Else
    gi12t20.state = 1
    gi11t20.state = 0
    gi11ct20.state = 0
  End If

  If LampState(200) = 1 and target21 = 1 then
    gi12t21.state = 1
    gi11t21.state = 1
    gi11ct21.state = 1
  Else
    gi12t21.state = 0
    gi11t21.state = 0
    gi11ct21.state = 0
  End If

End Sub

' Hat Trick Standups

Sub t22_Hit:vpmTimer.PulseSw(22):pt22a.RotY = -5:pt22b.RotY = -5:Target22Step = 0:t22t.Enabled = True:End Sub
Sub t22t_timer()
  Select Case Target22Step
    Case 1:pt22a.RotY = -3:pt22b.RotY = -3
        Case 2:pt22a.RotY = 2:pt22b.RotY = 2
        Case 3:pt22a.RotY = -1:pt22b.RotY = -1
        Case 4:pt22a.RotY = 0:pt22b.RotY = 0:t22t.Enabled = False:Target22Step = 0
     End Select
  Target22Step = Target22Step + 1
End Sub

Sub t23_Hit:vpmTimer.PulseSw(23):pt23a.RotY = -5:pt23b.RotY = -5:Target23Step = 0:t23t.Enabled = True:End Sub
Sub t23t_timer()
  Select Case Target23Step
    Case 1:pt23a.RotY = -3:pt23b.RotY = -3
        Case 2:pt23a.RotY = 2:pt23b.RotY = 2
        Case 3:pt23a.RotY = -1:pt23b.RotY = -1
        Case 4:pt23a.RotY = 0:pt23b.RotY = 0:t23t.Enabled = False:Target23Step = 0
     End Select
  Target23Step = Target23Step + 1
End Sub

Sub t24_Hit:vpmTimer.PulseSw(24):pt24a.RotY = -5:pt24b.RotY = -5:Target24Step = 0:t24t.Enabled = True:End Sub
Sub t24t_timer()
  Select Case Target24Step
    Case 1:pt24a.RotY = -3:pt24b.RotY = -3
        Case 2:pt24a.RotY = 2:pt24b.RotY = 2
        Case 3:pt24a.RotY = -1:pt24b.RotY = -1
        Case 4:pt24a.RotY = 0:pt24b.RotY = 0:t24t.Enabled = False:Target24Step = 0
     End Select
  Target24Step = Target24Step + 1
End Sub


' Bomb Standups

Sub t25_Hit:vpmTimer.PulseSw(25):pt25a.RotZ = 5:pt25b.RotZ = 5:Target25Step = 0:t25t.Enabled = True:End Sub
Sub t25t_timer()
  Select Case Target25Step
    Case 1:pt25a.RotZ = 3:pt25b.RotZ = 3
        Case 2:pt25a.RotZ = -2:pt25b.RotZ = -2
        Case 3:pt25a.RotZ = 1:pt25b.RotZ = 1
        Case 4:pt25a.RotZ = 0:pt25b.RotZ = 0:t25t.Enabled = False:Target25Step = 0
     End Select
  Target25Step = Target25Step + 1
End Sub


Sub t26_Hit:vpmTimer.PulseSw(26):pt26a.RotZ = 5:pt26b.RotZ = 5:Target26Step = 0:t26t.Enabled = True:End Sub
Sub t26t_timer()
  Select Case Target26Step
    Case 1:pt26a.RotZ = 3:pt26b.RotZ = 3
        Case 2:pt26a.RotZ = -2:pt26b.RotZ = -2
        Case 3:pt26a.RotZ = 1:pt26b.RotZ = 1
        Case 4:pt26a.RotZ = 0:pt26b.RotZ = 0:t26t.Enabled = False:Target26Step = 0
     End Select
  Target26Step = Target26Step + 1
End Sub


Sub t27_Hit:vpmTimer.PulseSw(27):pt27a.RotZ = 5:pt27b.RotZ = 5:Target27Step = 0:t27t.Enabled = True:End Sub
Sub t27t_timer()
  Select Case Target27Step
    Case 1:pt27a.RotZ = 3:pt27b.RotZ = 3
        Case 2:pt27a.RotZ = -2:pt27b.RotZ = -2
        Case 3:pt27a.RotZ = 1:pt27b.RotZ = 1
        Case 4:pt27a.RotZ = 0:pt27b.RotZ = 0:t27t.Enabled = False:Target27Step = 0
     End Select
  Target27Step = Target27Step + 1
End Sub


Sub t28_Hit:vpmTimer.PulseSw(28):pt28a.RotZ = 5:pt28b.RotZ = 5:Target28Step = 0:t28t.Enabled = True:End Sub
Sub t28t_timer()
  Select Case Target28Step
    Case 1:pt28a.RotZ = 3:pt28b.RotZ = 3
        Case 2:pt28a.RotZ = -2:pt28b.RotZ = -2
        Case 3:pt28a.RotZ = 1:pt28b.RotZ = 1
        Case 4:pt28a.RotZ = 0:pt28b.RotZ = 0:t28t.Enabled = False:Target28Step = 0
     End Select
  Target28Step = Target28Step + 1
End Sub


'''''''''''''''
''' Bumpers
'''''''''''''''

Sub Bumper1_hit:vpmTimer.PulseSw(44):RandomSoundBumperMiddle Bumper1:End Sub
Sub Bumper2_hit:vpmTimer.PulseSw(45):RandomSoundBumperTop Bumper2:End Sub
Sub Bumper3_hit:vpmTimer.PulseSw(46):RandomSoundBumperBottom Bumper3:End Sub


'''''''''''''''
''' Ramps
'''''''''''''''

Dim OnRockyRamp, OnWBRamp

Sub sw47_Hit:Controller.Switch(47) = 1:pWayBack_switch.ObjRotX = 5:OnWBRamp = 1:End Sub
Sub sw47_Unhit:Controller.Switch(47) = 0:pWayBack_switch.ObjRotX = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:pRampSwitch1.ObjRotX = -10:End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:pRampSwitch1.ObjRotX = 0:End Sub

Sub sw49_Hit:Controller.Switch(49) = 1:pRockyramp_switch.ObjRotX = -5:OnRockyRamp = 1:End Sub
Sub sw49_Unhit:Controller.Switch(49) = 0:pRockyramp_switch.ObjRotX = 0:End Sub

Sub sw50_Hit:Controller.Switch(50) = 1:pRampSwitch.ObjRotX = -10:End Sub
Sub sw50_Unhit:Controller.Switch(50) = 0:pRampSwitch.ObjRotX = 0:End Sub


''''''''''''
'  Slings
''''''''''''

Dim LeftSlingshotStep,RightSlingshotStep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  RandomSoundSlingshotLeft pSlingL
  vpmTimer.PulseSw 30
  LeftSlingA.Visible = false
  LeftSlingB.Visible = True
  pSlingL.TransZ = -10
  LeftSlingshotStep = 0
  Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LeftSlingshotStep
        Case 0:LeftSlingB.Visible = false:LeftSlingC.Visible = True:pSlingL.TransZ = -18
        Case 1:LeftSlingC.Visible = false:LeftSlingD.Visible = True:pSlingL.TransZ = -27
        Case 2:LeftSlingD.Visible = false:LeftSlingC.Visible = True:pSlingL.TransZ = -18
        Case 3:LeftSlingC.Visible = false:LeftSlingB.Visible = True:pSlingL.TransZ = -10
        Case 4:LeftSlingB.Visible = false:LeftSlingA.Visible = True:pSlingL.TransZ = 0:Me.TimerEnabled = 0 '
    End Select

    LeftSlingshotStep = LeftSlingshotStep + 1
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  RandomSoundSlingshotRight pSlingR
  vpmTimer.PulseSw 38
  RightSlingA.Visible = false
  RightSlingB.Visible = True
  pSlingR.TransZ = -10
  RightSlingshotStep = 0
  Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RightSlingshotStep
        Case 0:RightSlingB.Visible = false:RightSlingC.Visible = True:pSlingR.TransZ = -18
        Case 1:RightSlingC.Visible = false:RightSlingD.Visible = True:pSlingR.TransZ = -27
        Case 2:RightSlingD.Visible = false:RightSlingC.Visible = True:pSlingR.TransZ = -18
        Case 3:RightSlingC.Visible = false:RightSlingB.Visible = True:pSlingR.TransZ = -10
        Case 4:RightSlingB.Visible = false:RightSlingA.Visible = True::pSlingR.TransZ = 0:Me.TimerEnabled = 0 '
    End Select

    RightSlingshotStep = RightSlingshotStep + 1
End Sub

Sub Diverter2a_hit()
  PlaySoundAt "ramp_diverter_hit",pRampDiverter
End Sub


Sub GameTimer_Timer
  RollingUpdate
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''  Ball Trough system''''''''''''''''''''''''''
'''''''''''''''''''''by cyberpez''''''''''''''''''''''''''''''''
''''''''''''''''based off of EalaDubhSidhe's''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Dim BallCount
Dim cBall1, cBall2, cBall3

dim bstatus

Sub CreateBalls()
  Controller.Switch(11) = 1
  Controller.Switch(12) = 1
  Controller.Switch(13) = 1
  Set cBall1 = Kicker1.CreateBall()
  Set cBall2 = Kicker2.CreateBall()
  Set cBall3 = Kicker3.CreateBall()

  If BallMod = 1 Then
    cBall1.Image = "Chrome_Ball_29"
    cBall1.FrontDecal = "BlueYellowBall1"
    cBall2.Image = "Chrome_Ball_29"
    cBall2.FrontDecal = "BlueYellowBall2"
    cBall3.Image = "Chrome_Ball_29"
    cBall3.FrontDecal = "BlueYellowBall3"
  End If
End Sub


Sub Kicker3_Hit():Controller.Switch(11) = 1:UpdateTrough:End Sub
Sub Kicker3_UnHit():Controller.Switch(11) = 0:UpdateTrough:End Sub
Sub Kicker2_Hit():Controller.Switch(12) = 1:UpdateTrough:End Sub
Sub Kicker2_UnHit():Controller.Switch(12) = 0:UpdateTrough:End Sub
Sub Kicker1_Hit():Controller.Switch(13) = 1:UpdateTrough:End Sub
Sub Kicker1_UnHit():Controller.Switch(13) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  CheckBallStatus.Interval = 300
  CheckBallStatus.Enabled = 1
End Sub

Sub CheckBallStatus_timer()
  If Kicker1.BallCntOver = 0 Then Kicker2.kick 60, 12
  If Kicker2.BallCntOver = 0 Then Kicker3.kick 60, 9
  Me.Enabled = 0
End Sub

Dim Kicker1active, Kicker2active, Kicker3active, Kicker4active, Kicker5active, Kicker6active

dim DontKickAnyMoreBalls

Sub KickBallToLane(Enabled)
  If DontKickAnyMoreBalls = 0 then
    PlaySoundAt "BallRelease1",screw35
    Kicker1.Kick 70,12
    bstatus = 2
    Kicker1active = 0
    iBall = iBall - 1
    fgBall = false
    UpperGIon = 1
    Controller.Switch(13)=0
    DontKickAnyMoreBalls = 1
    DKTMstep = 1
    DontKickToMany.enabled = true
    BallsInPlay = BallsInPlay + 1
  End If
End Sub


Dim DKTMstep

Sub DontKickToMany_timer ()
  Select Case DKTMstep
  Case 1:
  Case 2:
  Case 3: DontKickAnyMoreBalls = 0:DontKickToMany.Enabled = False: DontKickAnyMoreBalls = 0
  End Select
  DKTMstep = DKTMstep + 1
End Sub


sub kisort(enabled)
  if fgBall then
    Drain.Kick 70,20
    iBall = iBall + 1
    fgBall = false

  end if

end sub


Sub Drain_hit()
  RandomSoundDrain Drain
  controller.switch(10) = true
  fgBall = true
  iBall = iBall + 1
  BallsInPlay = BallsInPlay - 1
End Sub

Sub Drain_UnHit()
  Controller.Switch(10) = 0
End Sub


'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1       'Do Not Change - default setting
BCvel = 4       'Controls the speed of the ball movement
BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
  End If
End Sub



'**********
' Gi Lights
'**********

dim GION

Sub SolGi(Enabled)
    Dim obj
    If Enabled Then

SetLamp 200, 0
SetLamp 111, 0
  If  GION = 1 then playsound "flasher_relay_off", 0
  GION = 0
    Else

SetLamp 200, 1
SetLamp 111, 1
  If  GION = 0 then playsound "flasher_relay_on", 0
  GION = 1
    End If
End Sub



'================Light Handling==================
'       GI, Flashers, and Lamp handling
'Based on JP's VP10 fading Lamp routine, based on PD's Fading Lights
'       Mod FrameTime and GI handling by nFozzy
'================================================
'Short installation
'Keep all non-GI lamps/Flashers in a big collection called aLampsAll
'Initialize SolModCallbacks: Const UseVPMModSol = 1 at the top of the script, before LoadVPM. vpmInit me in table1_Init()
'LUT images (optional)
'Make modifications based on era of game (setlamp / flashc for games without solmodcallback, use bonus GI subs for games with only one GI control)

Dim LampState(340), FadingLevel(340), CollapseMe
Dim FlashSpeedUp(340), FlashSpeedDown(340), FlashMin(340), FlashMax(340), FlashLevel(340)
Dim SolModValue(340)    'holds 0-255 modulated solenoid values

'These are used for fading lights and flashers brighter when the GI is darker
Dim LampsOpacity(340, 2) 'Columns: 0 = intensity / opacity, 1 = fadeup, 2 = FadeDown
Dim GIscale(4)  '5 gi strings
Dim TextureArray1: TextureArray1 = Array("Plastic with an image trans", "Plastic with an image")
Dim TextureRedBulbArray: TextureRedBulbArray = Array("Bulb Red trans", "Bulb Red")
Dim TextureBlueBulbArray: TextureBlueBulbArray = Array("Bulb Blue trans", "Bulb Blue")
Dim TextureGreenBulbArray: TextureGreenBulbArray = Array("Bulb Green trans", "Bulb Green")
Dim TextureYellowBulbArray: TextureYellowBulbArray = Array("Bulb Yellow trans", "Bulb Yellow")
Dim BulbArray1: BulbArray1 = Array("Bulb_on_texture", "Bulb_texture")
Dim RedLight: RedLight = Array("Bulb_texture_on", "Bulb_texture_66", "Bulb_texture_33", "Bulb_texture_off")
Dim BlueLight: BlueLight = Array("Bulb_Blue_texture_on", "Bulb_Blue_texture_66", "Bulb_Blue_texture_33", "Bulb_Blue_texture_off")
Dim GreenLight: GreenLight = Array("Bulb_Green_texture_on", "Bulb_Green_texture_66", "Bulb_Green_texture_33", "Bulb_Green_texture_off")
Dim YellowLight: YellowLight = Array("Bulb_Yellow_texture_on", "Bulb_Yellow_texture_66", "Bulb_Yellow_texture_33", "Bulb_Yellow_texture_off")
Dim RedBumperCap: RedBumperCap = Array("RaB_Bumpercap_redon", "RaB_Bumpercap_red66", "RaB_Bumpercap_red33", "RaB_Bumpercap")

Dim TestLight: TestLight = Array("Bulb_Yellow_texture_on", "Bulb_texture_66", "Bulb_Green_texture_33", "Bulb_Blue_texture_off")

Dim FilamentArray: FilamentArray = Array("WireDT_on", "WireDT_66", "WireDT_33", "WireDT_off")
Dim ClearBulbArray: ClearBulbArray = Array("BulbGIOn", "BulbGIOff")


InitLamps

reDim CollapseMe(1) 'Setlamps and SolModCallBacks   (Click Me to Collapse)
    Sub SetLamp(nr, value)
        If value <> LampState(nr) Then
            LampState(nr) = abs(value)
            FadingLevel(nr) = abs(value) + 4
        End If
    End Sub

    Sub SetLampm(nr, nr2, value)    'set 2 lamps
        If value <> LampState(nr) Then
            LampState(nr) = abs(value)
            FadingLevel(nr) = abs(value) + 4
        End If
        If value <> LampState(nr2) Then
            LampState(nr2) = abs(value)
            FadingLevel(nr2) = abs(value) + 4
        End If
    End Sub

    Sub SetModLamp(nr, value)
        If value <> SolModValue(nr) Then
            SolModValue(nr) = value
            if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
            FadingLevel(nr) = LampState(nr) + 4
        End If
    End Sub

    Sub SetModLampM(nr, nr2, value) 'set 2 modulated lamps
        If value <> SolModValue(nr) Then
            SolModValue(nr) = value
            if value > 0 then LampState(nr) = 1 else LampState(nr) = 0
            FadingLevel(nr) = LampState(nr) + 4
        End If
        If value <> SolModValue(nr2) Then
            SolModValue(nr2) = value
            if value > 0 then LampState(nr2) = 1 else LampState(nr2) = 0
            FadingLevel(nr2) = LampState(nr2) + 4
        End If
    End Sub

'#end section
reDim CollapseMe(2) 'InitLamps  (Click Me to Collapse)
    Sub InitLamps() 'set fading speeds and other stuff here
        GetOpacity aLampsAll    'All non-GI lamps and flashers go in this object array for compensation script!
        Dim x
        for x = 0 to uBound(LampState)
            LampState(x) = 0    ' current light state, independent of the fading level. 0 is off and 1 is on
            FadingLevel(x) = 4  ' used to track the fading state
            FlashSpeedUp(x) = 0.1   'Fading speeds in opacity per MS I think (Not used with nFadeL or nFadeLM subs!)
            FlashSpeedDown(x) = 0.1

            FlashMin(x) = 0.001         ' the minimum value when off, usually 0
            FlashMax(x) = 1             ' the minimum value when off, usually 1
            FlashLevel(x) = 0.001       ' Raw Flasher opacity value. Start this >0 to avoid initial flasher stuttering.

            SolModValue(x) = 0          ' Holds SolModCallback values

        Next

        for x = 0 to uBound(giscale)
            Giscale(x) = 1.625          ' lamp GI compensation multiplier, eg opacity x 1.625 when gi is fully off
        next

        for x = 11 to 110 'insert fading levels (only applicable for lamps that use FlashC sub)
            FlashSpeedUp(x) = 0.015
            FlashSpeedDown(x) = 0.009
        Next

        for x = 111 to 186  'Flasher fading speeds 'intensityscale(%) per 10MS
            FlashSpeedUp(x) = 1.1
            FlashSpeedDown(x) = 0.9
        next

        for x = 200 to 203      'GI relay on / off  fading speeds
            FlashSpeedUp(x) = 0.01
            FlashSpeedDown(x) = 0.008
            FlashMin(x) = 0
        Next
        for x = 300 to 303      'GI 8 step modulation fading speeds
            FlashSpeedUp(x) = 0.01
            FlashSpeedDown(x) = 0.008
            FlashMin(x) = 0
        Next

        UpdateGIon 0, 1:UpdateGIon 1, 1: UpdateGIon 2, 1 : UpdateGIon 3, 1:UpdateGIon 4, 1
        UpdateGI 0, 7:UpdateGI 1, 7:UpdateGI 2, 7 : UpdateGI 3, 7:UpdateGI 4, 7
    End Sub

    Sub GetOpacity(a)   'Keep lamp/flasher data in an array
        Dim x
        for x = 0 to (a.Count - 1)
            On Error Resume Next
            if a(x).Opacity > 0 then a(x).Uservalue = a(x).Opacity
            if a(x).Intensity > 0 then a(x).Uservalue = a(x).Intensity
            If a(x).FadeSpeedUp > 0 then LampsOpacity(x, 1) = a(x).FadeSpeedUp : LampsOpacity(x, 2) = a(x).FadeSpeedDown
        Next
        for x = 0 to (a.Count - 1) : LampsOpacity(x, 0) = a(x).UserValue : Next
    End Sub

    sub DebugLampsOn(input):Dim x: for x = 10 to 100 : setlamp x, input : next :  end sub

'#end section

reDim CollapseMe(3) 'LampTimer  (Click Me to Collapse)
    LampTimer.Interval = -1 '-1 is ideal, but it will technically work with any timer interval
    Dim FrameTime, InitFadeTime : FrameTime = 10    'Count Frametime
    Sub LampTimer_Timer()
        FrameTime = gametime - InitFadeTime
        Dim chgLamp, num, chg, ii
        chgLamp = Controller.ChangedLamps
        If Not IsEmpty(chgLamp) Then
            For ii = 0 To UBound(chgLamp)
                LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
                FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
            Next
        End If

        UpdateGIstuff
        UpdateLamps
        UpdateFlashers
    CheckDropShadows

        InitFadeTime = gametime
    End Sub
'#end section
reDim CollapseMe(4) 'ASSIGNMENTS: Lamps, GI, and Flashers (Click Me to Collapse)
    Sub UpdateGIstuff()

    End Sub

    Sub UpdateFlashers()

    End Sub

    Sub UpdateLamps()



FadeGI 200
UpdateGIobjectsSingle 200, theGicollection
GiCompensationSingle 200, aLampsAll, GIscale(0)
FadeLUTsingle 200, "LUTCont_", 28



  FadePri4m 1, s2wB1, BlueLight
  FadeDisableLighting 1, s2wB1
  FadeMaterialP 1, s2wB1, TextureArray1
  Flashc 1, s2wF1
  FadePri4m 2, s2wB2, BlueLight
  FadeDisableLighting 2, s2wB2
  FadeMaterialP 2, s2wB2, TextureArray1
  Flashc 2, s2wF2
  FadePri4m 3, s2wB3, BlueLight
  FadeDisableLighting 3, s2wB3
  FadeMaterialP 3, s2wB3, TextureArray1
  Flashc 3, s2wF3
  FadePri4m 4, s2wB4, YellowLight
  FadeDisableLighting 4, s2wB4
  FadeMaterialP 4, s2wB4, TextureArray1
  Flashc 4, s2wF4
  FadePri4m 5, s2wB5, YellowLight
  FadeDisableLighting 5, s2wB5
  FadeMaterialP 5, s2wB5, TextureArray1
  Flashc 5, s2wF5
  FadePri4m 6, s2wB6, YellowLight
  FadeDisableLighting 6, s2wB6
  FadeMaterialP 16, s2wB6, TextureArray1
  Flashc 6, s2wF6
  FadePri4m 7, s2wB7, RedLight
  FadeDisableLighting 7, s2wB7
  FadeMaterialP 7, s2wB7, TextureArray1
  Flashc 7, s2wF7
  FadePri4m 8, s2wB8, RedLight
  FadeDisableLighting 8, s2wB8
  FadeMaterialP 8, s2wB8, TextureArray1
  Flashc 8, s2wF8
  FadePri4m 9, s2wB9, RedLight
  FadeDisableLighting 9, s2wB9
  FadeMaterialP 9, s2wB9, TextureArray1
  Flashc 9, s2wF9
  FadePri4m 10, s2wB10, GreenLight
  FadeDisableLighting 10, s2wB10
  FadeMaterialP 10, s2wB10, TextureArray1
  Flashc 10, s2wF10
  FadePri4m 11, s2wB11, GreenLight
  FadeDisableLighting 11, s2wB11
  FadeMaterialP 11, s2wB11, TextureArray1
  Flashc 11, s2wF11
  FadePri4m 12, s2wB12, GreenLight
  FadeDisableLighting 12, s2wB12
  FadeMaterialP 12, s2wB12, TextureArray1
  Flashc 12, s2wF12

  FadePri4m 13, s2wB13, RedLight
  FadeDisableLighting 13, s2wB13
  FadeMaterialP 13, s2wB13, TextureArray1
  Flashc 13, s2wF13
  FadePri4m 14, s2wB14, YellowLight
  FadeDisableLighting 14, s2wB14
  FadeMaterialP 14, s2wB14, TextureArray1
  Flashc 14, s2wF14






If WBRBulb = 1 then
  FadePri4m 53, pWBbulb, BlueLight
Else
  FadePri4m 53, pWBbulb, RedLight
End If
  FadeDisableLighting 53, pWBbulb
  FadeMaterialP 53, pWBbulb, TextureArray1
  Flashc 53, flasher16a



If RRBulb = 1 then
  FadePri4m 16, pRbulb, BlueLight
Else
  FadePri4m 16, pRbulb, RedLight
End If
  FadeDisableLighting 16, pRbulb
  FadeMaterialP 16, pRbulb, TextureArray1
  Flashc 16, flasher16b

  NFadeL 17, L17
  NFadeL 18, L18
  NFadeL 19, L19
  NFadeL 20, L20
  NFadeL 21, L21
  NFadeL 22, L22
  NFadeL 23, L23
  NFadeL 24, L24
  NFadeL 25, L25
  NFadeL 26, L26
  NFadeL 27, L27
  NFadeL 28, L28
  NFadeL 29, L29
  NFadeL 30, L30
  NFadeL 31, L31
  NFadeL 32, L32
  NFadeL 33, L33
  NFadeL 34, L34
  NFadeL 35, L35
  NFadeL 36, L36
  NFadeL 37, L37
  NFadeL 38, L38
  NFadeL 39, L39
  NFadeL 40, L40
  NFadeL 41, L41
  NFadeL 42, L42
  NFadeL 43, L43
  NFadeL 44, L44
  NFadeL 45, L45
  NFadeL 46, L46
  NFadeL 47, L47
  NFadeL 49, L49
  NFadeL 50, L50
  NFadeL 51, L51
  NFadeLm 52, L52a
  NFadeL 52, L52b
  NFadeL 54, L54
  NFadeL 55, L55
  NFadeL 56, L56
  NFadeL 57, L57
  NFadeL 58, L58
  NFadeL 59, L59
  NFadeL 60, L60
  NFadeL 61, L61
  NFadeL 62, l62
  NFadeL 63, L63
  NFadeL 64, L64


' Flashers and GI

  NFadeLm 101, l1a
  NFadeLm 101, f1a
  NFadeL 101, f1b

  NFadeLm 102, l2a
  NFadeLm 102, l2b
  NFadeLm 102, l2c
  NFadeL 102, l2d

  FadeDisableLighting 103, pBulbFilament3
  FadePri4m 103, pBulbFilament3, FilamentArray
  FadeMaterialP 103, pBulbFlasher3, ClearBulbArray
  NFadeLm 103, l3a
  NFadeLm 103, l3b
  NFadeLm 103, f3a
  NFadeL 103, f3b

  NFadeLm 104, f4a
  NFadeL 104, l4a

  FadeDisableLighting 105, Primitive15
  FadeMaterialP 105, Primitive15, TextureArray1
  FadeDisableLighting 105, Primitive3
  FadeMaterialP 105, Primitive3, TextureArray1
  NFadeLm 105, f5a
  NFadeLm 105, f5b
  NFadeLm 105, fl5a
  NFadeLm 105, fl5a2
  NFadeLm 105, fl5b2
  NFadeLm 105, fl5b

  FadeDisableLighting 105, pBulbFilament1
  FadePri4m 105, pBulbFilament1, FilamentArray
  FadeMaterialP 105, pBulbFlasher1, ClearBulbArray
  NFadeLm 105, f5a
  FadeDisableLighting 105, pBulbFilament2
  FadePri4m 105, pBulbFilament2, FilamentArray
  FadeMaterialP 105, pBulbFlasher2, ClearBulbArray
  NFadeLm 105, f5b
  Flashm 105, f5c1
  Flashm 105, f5c2
  Flashm 105, f5d1
  Flashc 105, f5d2

  NFadeLm 106, l6a
  NFadeLm 106, l6b
  NFadeLm 106, f6a
  NFadeLm 106, f6b
  Flashc 106, f6c

  NFadeLm 107, l7a
  NFadeLm 107, l7b
  NFadeL 107, l7c

  NFadeLm 108, f8a
  NFadeL 108, l8a


'''GI related things
  NFadeLm 200, blight1
  NFadeLm 200, blight2
  NFadeLm 200, Lbumper3
  FadeDisableLighting1 200, pRbulb2
  FadePri4m 200, pRbulb2, RedLight
  FadeMaterialP 200, pRbulb2, TextureArray1
  FadeDisableLighting1 200, primitive1
  FadeDisableLighting1 200, primitive2
If RedBumperCaps = 1 then
  FadePri4m 200, primitive1, RedBumperCap
  FadePri4m 200, primitive2, RedBumperCap
End If


   End Sub

'#end section


''''Additions by CP


Sub FadeDisableLighting(nr, a)
    Select Case FadingLevel(nr)
    Case 2:a.BlendDisableLighting = .33
    Case 3:a.BlendDisableLighting = .66
    Case 4:a.BlendDisableLighting = 0
    Case 5:a.BlendDisableLighting = 1
    End Select
End Sub

Sub FadeDisableLighting1(nr, a)
    Select Case FadingLevel(nr)
    Case 2:a.BlendDisableLighting = .033
    Case 3:a.BlendDisableLighting = .066
    Case 4:a.BlendDisableLighting = 0
    Case 5:a.BlendDisableLighting = .1
    End Select
End Sub

'trxture swap
dim itemw, itemp, itemp2

Sub FadeMaterialW(nr, itemw, group)
    Select Case FadingLevel(nr)
        Case 4:itemw.TopMaterial = group(1):itemw.SideMaterial = group(1)
        Case 5:itemw.TopMaterial = group(0):itemw.SideMaterial = group(0)
    End Select
End Sub


Sub FadeMaterialP(nr, itemp, group)
    Select Case FadingLevel(nr)
        Case 4:itemp.Material = group(1)
        Case 5:itemp.Material = group(0)
    End Select
End Sub


Sub FadeMaterial2P(nr, itemp2, group)
    Select Case FadingLevel(nr)
        Case 4:itemp2.Material = group(1)
        Case 5:itemp2.Material = group(0)
    End Select
End Sub


'Reels

Sub FadeR(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.SetValue 3:FadingLevel(nr) = 0
        Case 3:a.SetValue 2:FadingLevel(nr) = 2
        Case 4:a.SetValue 1:FadingLevel(nr) = 3
        Case 5:a.SetValue 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeRm(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.SetValue 3
        Case 3:a.SetValue 2
        Case 4:a.SetValue 1
        Case 5:a.SetValue 1
    End Select
End Sub


Sub FadePri4m(nr, pri, group)
    Select Case FadingLevel(nr)
    Case 2:pri.image = group(1) 'Off
        Case 3:pri.image = group(3) 'Fading...
        Case 4:pri.image = group(2) 'Fading...
        Case 5:pri.image = group(0) 'ON
    End Select
End Sub

Sub FadePri4(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 2:pri.image = group(1):FadingLevel(nr) = 0 'Off
        Case 3:pri.image = group(3):FadingLevel(nr) = 2 'Fading...
        Case 4:pri.image = group(2):FadingLevel(nr) = 3 'Fading...
        Case 5:pri.image = group(0):FadingLevel(nr) = 1 'ON
    End Select
End Sub

''''End Of Additions by CP



reDim CollapseMe(5) 'Combined GI subs / functions (Click Me to Collapse)
    Set GICallback = GetRef("UpdateGIon")       'On/Off GI to NRs 200-203
    Sub UpdateGIOn(no, Enabled) : Setlamp no+200, cInt(enabled) : End Sub

    Set GICallback2 = GetRef("UpdateGI")
    Sub UpdateGI(no, step)                      '8 step Modulated GI to NRs 300-303
        Dim ii, x', i
        If step = 0 then exit sub 'only values from 1 to 8 are visible and reliable. 0 is not reliable and 7 & 8 are the same so...
        SetModLamp no+300, ScaleGI(step, 0)
        LampState((no+300)) = 0
    '   if no = 2 then tb.text = no & vbnewline & step & vbnewline & ScaleGI(step,0) & SolModValue(102)
    End Sub

    Function ScaleGI(value, scaletype)  'returns an intensityscale-friendly 0->100% value out of 1>8 'it does go to 8
        Dim i
        Select Case scaletype   'select case because bad at maths
            case 0  : i = value * (1/8) '0 to 1
            case 25 : i = (1/28)*(3*value + 4)
            case 50 : i = (value+5)/12
            case else : i = value * (1/8)   '0 to 1
    '           x = (4*value)/3 - 85    '63.75 to 255
        End Select
        ScaleGI = i
    End Function

'   Dim LSstate : LSstate = False   'fading sub handles SFX 'Uncomment to enable
    Sub FadeGI(nr) 'in On/off       'Updates nothing but flashlevel
        Select Case FadingLevel(nr)
            Case 3 : FadingLevel(nr) = 0
            Case 4 'off
    '           If Not LSstate then Playsound "FX_Relay_Off",0,LVL(0.1) : LSstate = True    'handle SFX
                FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime)
                If FlashLevel(nr) < FlashMin(nr) Then
                   FlashLevel(nr) = FlashMin(nr)
                   FadingLevel(nr) = 3 'completely off
    '               LSstate = False
                End if
            Case 5 ' on
    '           If Not LSstate then Playsound "FX_Relay_On",0,LVL(0.1) : LSstate = True 'handle SFX
                FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * FrameTime)
                If FlashLevel(nr) > FlashMax(nr) Then
                    FlashLevel(nr) = FlashMax(nr)
                    FadingLevel(nr) = 6 'completely on
    '               LSstate = False
                End if
            Case 6 : FadingLevel(nr) = 1
        End Select
    End Sub
    Sub ModGI(nr2) 'in 0->1     'Updates nothing but flashlevel 'never off
        Dim DesiredFading
        Select Case FadingLevel(nr2)
            case 3 : FadingLevel(nr2) = 0   'workaround - wait a frame to let M sub finish fading
    '       Case 4 : FadingLevel(nr2) = 3   'off -disabled off, only gicallback1 can turn off GI(?) 'experimental
            Case 5, 4 ' Fade (Dynamic)
                DesiredFading = SolModValue(nr2)
                if FlashLevel(nr2) < DesiredFading Then '+
                    FlashLevel(nr2) = FlashLevel(nr2) + (FlashSpeedUp(nr2)  * FrameTime )
                    If FlashLevel(nr2) >= DesiredFading Then FlashLevel(nr2) = DesiredFading : FadingLevel(nr2) = 1
                elseif FlashLevel(nr2) > DesiredFading Then '-
                    FlashLevel(nr2) = FlashLevel(nr2) - (FlashSpeedDown(nr2) * FrameTime    )
                    If FlashLevel(nr2) <= DesiredFading Then FlashLevel(nr2) = DesiredFading : FadingLevel(nr2) = 6
                End If
            Case 6
                FadingLevel(nr2) = 1
        End Select
    End Sub

    Sub UpdateGIobjects(nr, nr2, a) 'Just Update GI
        If FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 Then
            Dim x, Output : Output = FlashLevel(nr2) * FlashLevel(nr)
            for each x in a : x.IntensityScale = Output : next
        End If
    end Sub

    Sub GiCompensation(nr, nr2, a, GIscaleOff)  'One NR pairing only fading
    '   tbgi.text = "GI: " & SolModValue(nr) & " " & FlashLevel(nr) & " " & FadingLevel(nr) & vbnewline & _
    '               "ModGI: " & SolModValue(nr2) & " " & FlashLevel(nr2) & " " & FadingLevel(nr2) & vbnewline & _
    '               "Solmodvalue, Flashlevel, Fading step"
        if FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 Then
            Dim x, Giscaler, Output : Output = FlashLevel(nr2) * FlashLevel(nr)
            Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1    'fade GIscale the opposite direction

            for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
                On Error Resume Next
                a(x).Opacity = LampsOpacity(x, 0) * Giscaler
                a(x).Intensity = LampsOpacity(x, 0) * Giscaler
                a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
                a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
            Next
            '       tbbb.text = giscaler & " on:" & FadingLevel(nr) & vbnewline & "flash: " & output & " onmod:" & FadingLevel(nr2) & vbnewline & l37.intensity
            '       tbbb1.text = FadingLevel(nr) & vbnewline & FadingLevel(nr2)
            '   tbgi1.text = Output & " giscale:" & giscaler    'debug
        End If
        '       tbbb1.text = FLashLevel(nr) & vbnewline & FlashLevel(nr2)
    End Sub

    Sub GiCompensationAvg(nr, nr2, nr3, nr4, a, GIscaleOff) 'Two pairs of NRs averaged together
    '   tbgi.text = "GI: " & SolModValue(nr) & " " & FlashLevel(nr) & " " & FadingLevel(nr) & vbnewline & _
    '               "ModGI: " & SolModValue(nr2) & " " & FlashLevel(nr2) & " " & FadingLevel(nr2) & vbnewline & _
    '               "Solmodvalue, Flashlevel, Fading step"
        if FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 or FadingLevel(nr3) > 1 or FadingLevel(nr4) > 1 Then
            Dim x, Giscaler, Output : Output = (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4))) /2)
            Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1    'fade GIscale the opposite direction

            for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
                On Error Resume Next
                a(x).Opacity = LampsOpacity(x, 0) * Giscaler
                a(x).Intensity = LampsOpacity(x, 0) * Giscaler
                a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
                a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
            Next


        REM tbgi1.text = "Output:" & output & vbnewline & _
                    REM "GIscaler" & giscaler & vbnewline & _
                    REM "..."
        End If
        REM tbgi.text = "GI0 " & flashlevel(200) & " " & flashlevel(300) & vbnewline & _
                    REM "GI1 " & flashlevel(201) & " " & flashlevel(301) & vbnewline & _
                    REM "GI2 " & flashlevel(202) & " " & flashlevel(302) & vbnewline & _
                    REM "GI3 " & flashlevel(203) & " " & flashlevel(303) & vbnewline & _
                    REM "GI4 " & flashlevel(204) & " " & flashlevel(304) & vbnewline & _
                    REM "..."
    End Sub

    Sub GiCompensationAvgM(nr, nr2, nr3, nr4, nr5, nr6, a, GIscaleOff)  'Three pairs of NRs averaged together
    '   tbgi.text = "GI: " & SolModValue(nr) & " " & FlashLevel(nr) & " " & FadingLevel(nr) & vbnewline & _
    '               "ModGI: " & SolModValue(nr2) & " " & FlashLevel(nr2) & " " & FadingLevel(nr2) & vbnewline & _
    '               "Solmodvalue, Flashlevel, Fading step"
        if FadingLevel(nr) > 1 or FadingLevel(nr2) > 1 Then
            Dim x, Giscaler, Output
            Output = (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4)) + (FlashLevel(nr5)*FlashLevel(nr6)))/3)

            Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1    'fade GIscale the opposite direction

            for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
                On Error Resume Next
                a(x).Opacity = LampsOpacity(x, 0) * Giscaler
                a(x).Intensity = LampsOpacity(x, 0) * Giscaler
                a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
                a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
            Next
            '       tbbb.text = giscaler & " on:" & FadingLevel(nr) & vbnewline & "flash: " & output & " onmod:" & FadingLevel(nr2) & vbnewline & l37.intensity
            '       tbbb1.text = FadingLevel(nr) & vbnewline & FadingLevel(nr2)
            '   tbgi1.text = Output & " giscale:" & giscaler    'debug
        End If
        '       tbbb1.text = FLashLevel(nr) & vbnewline & FlashLevel(nr2)
    End Sub

    Sub FadeLUT(nr, nr2, LutName, LutCount) 'fade lookuptable NOTE- this is a bad idea for darkening your table as
        If FadingLevel(nr) >2 or FadingLevel(nr2) > 2 Then              '-it will strip the whites out of your image
            Dim GoLut
            GoLut = cInt(LutCount * (FlashLevel(nr)*FlashLevel(nr2) )   )
            Table1.ColorGradeImage = LutName & GoLut
    '       tbgi2.text = Table1.ColorGradeImage & vbnewline & golut 'debug
        End If
    End Sub

    Sub FadeLUTavg(nr, nr2, nr3, nr4, LutName, LutCount)    'FadeLut for two GI strings (WPC)
        If FadingLevel(nr) >2 or FadingLevel(nr2) > 2 or FadingLevel(nr3) >2 or FadingLevel(nr4) > 2 Then
            Dim GoLut
            GoLut = cInt(LutCount * (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4))) /2) )
            Table1.ColorGradeImage = LutName & GoLut
            REM tbgi2.text = Table1.ColorGradeImage & vbnewline & golut 'debug
        End If
    End Sub

    Sub FadeLUTavgM(nr, nr2, nr3, nr4, nr5, nr6, LutName, LutCount) 'FadeLut for three GI strings (WPC)
        If FadingLevel(nr) >2 or FadingLevel(nr2) > 2 or FadingLevel(nr3) >2 or FadingLevel(nr4) > 2 or _
        FadingLevel(nr5) >2 or FadingLevel(nr6) > 2 Then
            Dim GoLut
            GoLut = cInt(LutCount * (((FlashLevel(nr)*FlashLevel(nr2)) + (FlashLevel(nr3)*Flashlevel(nr4)) + (FlashLevel(nr5)*FlashLevel(nr6)))/3)  )   'what a mess
            Table1.ColorGradeImage = LutName & GoLut
    '       tbgi2.text = Table1.ColorGradeImage & vbnewline & golut 'debug
        End If
    End Sub

'#end section

reDim CollapseMe(6) 'Fading subs     (Click Me to Collapse)
    Sub nModFlash(nr, object, scaletype, offscale)  'Fading with modulated callbacks
        Dim DesiredFading
        Select Case FadingLevel(nr)
            case 3 : FadingLevel(nr) = 0    'workaround - wait a frame to let M sub finish fading
            Case 4  'off
                If Offscale = 0 then Offscale = 1
                FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime   ) * offscale
                If FlashLevel(nr) < 0 then FlashLevel(nr) = 0 : FadingLevel(nr) = 3
                Object.IntensityScale = ScaleLights(FlashLevel(nr),0 )
            Case 5 ' Fade (Dynamic)
                DesiredFading = ScaleByte(SolModValue(nr), scaletype)
                if FlashLevel(nr) < DesiredFading Then '+
                    FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * FrameTime )
                    If FlashLevel(nr) >= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 1
                elseif FlashLevel(nr) > DesiredFading Then '-
                    FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime   )
                    If FlashLevel(nr) <= DesiredFading Then FlashLevel(nr) = DesiredFading : FadingLevel(nr) = 6
                End If
                Object.Intensityscale = ScaleLights(FlashLevel(nr),0 )
            Case 6 : FadingLevel(nr) = 1
        End Select
    End Sub

    Sub nModFlashM(nr, Object)
        Select Case FadingLevel(nr)
            Case 3, 4, 5, 6 : Object.Intensityscale = ScaleLights(FlashLevel(nr),0 )
        End Select
    End Sub

    Sub Flashc(nr, object)  'FrameTime Compensated. Can work with Light Objects (make sure state is 1 though)
        Select Case FadingLevel(nr)
            Case 3 : FadingLevel(nr) = 0
            Case 4 'off
                FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * FrameTime)
                If FlashLevel(nr) < FlashMin(nr) Then
                    FlashLevel(nr) = FlashMin(nr)
                   FadingLevel(nr) = 3 'completely off
                End if
                Object.IntensityScale = FlashLevel(nr)
            Case 5 ' on
                FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * FrameTime)
                If FlashLevel(nr) > FlashMax(nr) Then
                    FlashLevel(nr) = FlashMax(nr)
                    FadingLevel(nr) = 6 'completely on
                End if
                Object.IntensityScale = FlashLevel(nr)
            Case 6 : FadingLevel(nr) = 1
        End Select
    End Sub


Sub Flash(nr, object)
    Select Case FadingLevel(nr)
    Case 3
      FadingLevel(nr) = 0
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - (1/FlashSpeedDown(nr) * FrameTime)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
               FadingLevel(nr) = 3 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + (1/FlashSpeedUp(nr) * FrameTime)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 6 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    Case 6
      FadingLevel(nr) = 1
    End Select
End Sub


    Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
        select case FadingLevel(nr)
            case 3, 4, 5, 6 : Object.IntensityScale = FlashLevel(nr)
        end select
    End Sub

    Sub NFadeL(nr, object)  'Simple VPX light fading using State
   Select Case FadingLevel(nr)
        Case 3:object.state = 0:FadingLevel(nr) = 0
        Case 4:object.state = 0:FadingLevel(nr) = 3
        Case 5:object.state = 1:FadingLevel(nr) = 6
        Case 6:object.state = 1:FadingLevel(nr) = 1
    End Select
    End Sub

    Sub NFadeLm(nr, object) ' used for multiple lights
        Select Case FadingLevel(nr)
            Case 3:object.state = 0
            Case 4:object.state = 0
            Case 5:object.state = 1
            Case 6:object.state = 1
        End Select
    End Sub

'#End Section

reDim CollapseMe(7) 'Fading Functions (Click Me to Collapse)
    Function ScaleLights(value, scaletype)  'returns an intensityscale-friendly 0->100% value out of 255
        Dim i
        Select Case scaletype   'select case because bad at maths   'TODO: Simplify these functions. B/c this is absurdly bad.
            case 0  : i = value * (1 / 255) '0 to 1
            case 6  : i = (value + 17)/272  '0.0625 to 1
            case 9  : i = (value + 25)/280  '0.089 to 1
            case 15 : i = (value / 300) + 0.15
            case 20 : i = (4 * value)/1275 + (1/5)
            case 25 : i = (value + 85) / 340
            case 37 : i = (value+153) / 408     '0.375 to 1
            case 40 : i = (value + 170) / 425
            case 50 : i = (value + 255) / 510   '0.5 to 1
            case 75 : i = (value + 765) / 1020  '0.75 to 1
            case Else : i = 10
        End Select
        ScaleLights = i
    End Function

    Function ScaleByte(value, scaletype)    'returns a number between 1 and 255
        Dim i
        Select Case scaletype
            case 0 : i = value * 1  '0 to 1
            case 9 : i = (5*(200*value + 1887))/1037 'ugh
            case 15 : i = (16*value)/17 + 15
            Case 63 : i = (3*(value + 85))/4
            case else : i = value * 1   '0 to 1
        End Select
        ScaleByte = i
    End Function

'#end section

reDim CollapseMe(8) 'Bonus GI Subs for games with only simple On/Off GI (Click Me to Collapse)
    Sub UpdateGIobjectsSingle(nr, a)    'An UpdateGI script for simple (Sys11 / Data East or whatever)
        If FadingLevel(nr) > 1 Then
            Dim x, Output : Output = FlashLevel(nr)
            for each x in a : x.IntensityScale = Output : next
        End If
    end Sub

    Sub GiCompensationSingle(nr, a, GIscaleOff) 'One NR pairing only fading
        if FadingLevel(nr) > 1 Then
            Dim x, Giscaler, Output : Output = FlashLevel(nr)
            Giscaler = ((Giscaleoff-1) * (ABS(Output-1) )  ) + 1    'fade GIscale the opposite direction

            for x = 0 to (a.Count - 1) 'Handle Compensate Flashers
                On Error Resume Next
                a(x).Opacity = LampsOpacity(x, 0) * Giscaler
                a(x).Intensity = LampsOpacity(x, 0) * Giscaler
                a(x).FadeSpeedUp = LampsOpacity(x, 1) * Giscaler
                a(x).FadeSpeedDown = LampsOpacity(x, 2) * Giscaler
            Next
        End If
        '       tbbb1.text = FLashLevel(nr) & vbnewline & FlashLevel(nr2)
    End Sub

    Sub FadeLUTsingle(nr, LutName, LutCount)    'fade lookuptable NOTE- this is a bad idea for darkening your table as
        If FadingLevel(nr) >2 Then              '-it will strip the whites out of your image
            Dim GoLut
            GoLut = cInt(LutCount * FlashLevel(nr)  )
            Table1.ColorGradeImage = LutName & GoLut
    '       tbgi2.text = Table1.ColorGradeImage & vbnewline & golut 'debug
        End If
    End Sub

'#end section

Sub theend() : End Sub


REM Troubleshooting :
REM Flashers/gi are intermittent or aren't showing up
REM Ensure flashers start visible, light objects start with state = 1

REM No lamps or no GI
REM Make sure these constants are set up this way
REM Const UseSolenoids = 1
REM Const UseLamps = 0
REM Const UseGI = 1

REM SolModCallback error
REM Ensure you have the latest scripts. Clear out any loose scripts in your tables that might be causing conflicts.

REM Table1 Error
REM Rename the table to Table1 or find/Replace table1 with whatever the table's name is

REM SolModCallbacks aren't sending anything
REM Two important things to get SolModCallbacks to initialize properly:
REM Put this at the top of the script, before LoadVPM
    REM Const UseVPMModSol = 1
REM Put this in the table1_Init() section
    REM vpmInit me

' Ramp Sounds
Sub RightRampStart_Hit
  WireRampOn True
End Sub

Sub RightRampEnd_Hit
  WireRampOff
End Sub

Sub BlueRampStart_Hit
  WireRampOn False
End Sub

Sub LeftRampStart_Hit
  WireRampOn True
End Sub

Sub LeftRampEnd_Hit
  WireRampOff
End Sub

Sub GreenRampStart_Hit
  WireRampOn False
End Sub

Sub YellowRampStart_Hit
  WireRampOn False
End Sub

Sub YellowRampEnd_Hit
  WireRampOff
End Sub

Sub LeftRampFromYellowRamp_Hit
  WireRampOn True
End Sub


' Subway sounds
Sub RandomTopSubwayEnterSound()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtBallVol "topsubwayenter1",1
    Case 2 : PlaySoundAtBallVol "topsubwayenter2",1
    Case 3 : PlaySoundAtBallVol "topsubwayenter3",1
    Case 4 : PlaySoundAtBallVol "topsubwayenter4",1
  End Select
End Sub


'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity

'*******************************************
' Early 90's and after

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
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before

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
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
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

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle)  '-1 for Right Flipper
  Dim LiveCatchBounce                                                           'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime
  CatchTime = GameTime - FCount

  If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
    If CatchTime <= LiveCatch * 0.5 Then                        'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    Else
      LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)    'Partial catch when catch happens a bit late
    End If

    If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx = 0
    ball.angmomy = 0
    ball.angmomz = 0
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
Sub RDampen_Timer
  Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

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
Function RotPoint(x,y,angle)
  dim rx, ry
  rx = x*dCos(angle) - y*dSin(angle)
  ry = x*dSin(angle) + y*dCos(angle)
  RotPoint = Array(rx,ry)
End Function

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
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

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
BumperSoundFactor = 0.025           'volume multiplier; must not be zero
KnockerSoundLevel = 1              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.375 / 5      'volume multiplier; must not be zero
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

' InRect used in legacy function below from previous versions of this mod.
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
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    ' Changed from Z value of 35 to 105 due to higher playfield - TastyWasps
    If BallVel(gBOT(b)) > 1 And gBOT(b).z < 105 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ <  - 1 And gBOT(b).z < 135 And gBOT(b).z > 115 Then 'height adjust for ball drop sounds
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

    ' Legacy code to handle the left orbit diverter area saucer in a timer.  It used to be in JP's rolling update area.
    If InRect(gBOT(b).x,gBOT(b).y,(pUpKicker.x-30),(pUpKicker.y-30),(pUpKicker.x+30),(pUpKicker.y-30),(pUpKicker.x+30),(pUpKicker.y+30),(pUpKicker.x-30),(pUpKicker.y+30)) then
      If ABS(gBOT(b).velx) < 1 and ABS(gBOT(b).vely) < 1 Then
        Controller.Switch(52) = 1
        sw52.enabled = true
        BallInKicker1 = 1
        Set BallSaucer1 = gBOT(b)
      End If
    End If


  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


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
      If RampBalls(x,0).Z < 105 And RampBalls(x, 2) > RampMinLoops Then 'if ball is on the PF, remove  it
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

Sub FrameTimer_Timer()
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate
End Sub

' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

'Function Distance(ax,ay,bx,by)
' Distance = SQR((ax - bx)^2 + (ay - by)^2)
'End Function

'Dim PI: PI = 4*Atn(1)

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

'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

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
  If onPlayfield Then
    Dim gBOT: gBOT=getballs
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
  Dim gBOT: gBOT=getballs 'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls

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
      If gBOT(s).Z > 105 Then
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
      ElseIf gBOT(s).Z <= 105 And gBOT(s).Z > 95 Then
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
      If gBOT(s).Z > 105 Then 'In a ramp
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      ElseIf gBOT(s).Z <= 105 And gBOT(s).Z > 95 Then 'On pf
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
      If gBOT(s).Z < 105 And gBOT(s).X < 850 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
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





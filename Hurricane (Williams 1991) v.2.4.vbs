' HH     HH  UU     UU  RRRRRRRR   RRRRRRRR   II   CCCCCCC    AAAAA    NN      NN   EEEEEEEE
' HH     HH  UU     UU  RR     RR  RR     RR  II  CC         AA   AA   NNN     NN  EE
' HH     HH  UU     UU  RR     RR  RR     RR  II  CC        AA     AA  NNNN    NN  EE
' HH     HH  UU     UU  RR     RR  RR     RR  II  CC        AA     AA  NN NN   NN  EE
' HHHHHHHHH  UU     UU  RRRRRRRR   RRRRRRRR   II  CC        AAAAAAAAA  NN  NN  NN  EEEEEEEE
' HH     HH  UU     UU  RR RR      RR RR      II  CC        AA     AA  NN   NN NN  EE
' HH     HH  UU     UU  RR  RR     RR  RR     II  CC        AA     AA  NN    NNNN  EE
' HH     HH  UU     UU  RR    RR   RR    RR   II  CC        AA     AA  NN     NNN  EE
' HH     HH  UUUUUUUUU  RR     RR  RR     RR  II   CCCCCCC  AA     AA  NN      NN   EEEEEEEE     by Williams, 1991
'
' Based on Williams's Hurricane / IPD No. 1257 / 1991 / 4 Players
'
' VP9 - Original version by TAB - 2009.
' VP9 - Mod version by ICPjuggla, JimmyFingers, OldSkoolGamer and Koadic - 2013.
' VPX - Original version by Herweh - 2019.
' VR Version by Retro27, UnclePaulie and Sixtoe - 2021.
' VR-Hybrid version by RobbyKingPin, Retro27, DGrimmReaper, DaRdog, tomate, fidiko, hauntfreaks, supergibson and LoadedWeapon - 2025/2026.
'
' This mod contains the following improvements:
'
' RobbyKingPin:
' - nFozzy/rothbauerw physics, fleep sound with a lot of custom sounds from VPW's Space Station and VPW Ambient Ball Shadows
' - rothbauerw Standup and Drop Target codes
' - Updated target mass and Slingshot Corrections
' - Resized flippers matching settings provided by apophis
' - Flupper Domes, Flupper Bumpers, Flupper Flipperbats, Flupper Flasherlamps and Flupper 3D Inserts
' - New Trough system that saves the balls
' - More light tweaks to more modern standards based on VPW's settings
' - Fixed some missing lights
' - New ball image and scratches based on VPW's tables
' - Updated PWM on flashers GI and Inserts and added new GI lights
' - Added refraction probes on all ramps
' - Removed a lot of timers and triggers to improve performance
' - Big uprez in the playfield and plastics using AI
' - Added a new baked minimal room and a third room for Mixed Reality
' - Added new custom sideblades and instruction cards
' - Added Rainbow RGB as an option for the GI
' - Created new primitives for the plastics
' - Added and edited a primitive and a texture for the biff bars
' - Baked textures for the backglass and topper, sidewalls, custom blades, plastics, decals, ramps and playfield shadows
' - Blended shadows with the GI
' - VPW Tweak Menu
'
' Retro27:
' - Added a new cabinet in VR
' - Created new textures for the cabinet plus options
' - Created a new topper primitive
' - Added a key lock at the top of the back box
' - Adjusted the height on a few wiregates
' - Updated the playfield, inserts and some of the plastics textures
'
' DGrimmReaper:
' - Added fully animated VR backglass using hauntfreaks and Wildman assets
' - Fixed minor issues in VR
'
' DaRdog:
' - Baked and added a very cool Mega Room with an animated rollercoaster and donated the costs for the assets
'
' tomate:
' - Added new baked textures for all the ramps and edited the Hurricane ramp primitive
'
' fidiko:
' - Created a new texture for the backpanel and added shadows on the sidewall textures
'
' hauntfreaks:
' - Provided a new backglass to match his new b2s file
'
' supergibson:
' - Created a new ramp end section that provides a better feed to the inlanes
'
' LoadedWeapon:
' - Created a custom POV to fit on Legends Cabinets
'
' Special thanks to everybody involved during development and testing: cl4ve, Tomate, Studlygoorite, Burger, Guus_8005, Dr.Nobody, Mecha_Enron, m.carter78, ZandysArcade, Smaug, PinStratsDan, fidiko, Primetime5k, MikeDASpike,
' wrd1972, hauntfreaks, supergibson, rothbauerw, bthlonewolf, apophis, Sinizin
'
' Shoutout to:
' - DaRdog for donating the assets for the VR Mega Room
' - tomate for all the help with the bakes and the ramp
' - TAB, ICPjuggla, JimmyFingers, OldSkoolGamer, Koadic, Herweh, UnclePaulie and Sixtoe for their work on previous versions
'
' If I didn't mention anybody else in this development cycle I am unaware of this and just know that I still think you are awesome and also want to say you rock!
' And basically everybody from VPW I didn't mention just for having me as part of the family, you all absolutely rock!
'
Option Explicit : Randomize
SetLocale 1033

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SCoin = ""
Const BallSize = 50        'Ball diameter in VPX units; must be 50
Const BallMass = 1          'Ball mass must be 1
Const tnob = 3            'Total number of balls the table can hold
Const lob = 0

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim BIP             ' Balls in play
BIP = 0
Dim BIPL            ' Ball in plunger lane
BIPL = False
Dim plungerpress

'************************************************************************
Const DebugFlashers = False
Const DebugGI = False

'**************************************************************
'                    USER OPTIONS
'**************************************************************

'----- VR Options -----
Dim VRRoomChoice : VRRoomChoice = 1       ' 1 = Rollercoaster 2 = Minimal Room 3 = Mixed Reality
Const VRTest = 0                ' 1 = Testing VR in Live View, 0 = Do not force VR mode.

Const cGameName = "hurr_l2"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0

Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim UseVPMModSol : UseVPMModSol = 2

Dim EnviroSound : EnviroSound = "RollercoasterLoop"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

'----- Desktop Visible Elements -----
Dim VarHidden, UseVPMDMD, x

If Table1.ShowDT = True And renderingmode <> 2 Then
  UseVPMDMD = True
  VarHidden = 1
  PinCab_Rails.Visible = 1
  PinCab_TopRail.Visible = 1
Else
  UseVPMDMD = False
    VarHidden = 0
  PinCab_Rails.Visible = 0
  PinCab_TopRail.Visible = 0
  If Table1.ShowFSS = False And LegendsCabMod = 1 Then
    Flasherbloom1.Height = 500
    Flasherbloom2.Height = 500
    Flasherbloom3.Height = 500
    Flasherbloom4.Height = 500
    Flasherbloom5.Height = 500
    Flasherbloom6.Height = 500
    Flasherbloom7.Height = 500
    Flasherlampbloom1.Height = 500
    Flasherlampbloom2.Height = 500
    Flasherlampbloom3.Height = 500
    Flasherlampbloom4.Height = 500
    Flasherlampbloom5.Height = 500
  End If
End If

if B2SOn = true then VarHidden = 1

'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************

' DMD Rotation
Const cDMDRotation  = -1      '-1 for No change, 0 - DMD rotation of 0?, 1 - DMD rotation of 90?

Dim FlippersEnabled : FlippersEnabled = False ' Used to enable/disable flippers based on tilt status
Dim myMath : Set MyMath = New clsMyMath
Dim xx, bsRightJuggler

'----- VR Room Auto-Detect -----
Dim VR_Obj, VR_Room, VRMode

If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
  VRMode = True
  UseVPMDMD = True
  VarHidden = 0
  Textbox1.visible = 0
  PinCab_Rails.Visible = 1
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  If VRRoomChoice = 1 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRMEGARoom : VR_Obj.Visible = 1 : Next
    Room360.Visible = 0
    VR_Mega061Car1.PlayAnimEndless (0.4)
    VR_Mega052Car2.PlayAnimEndless (0.4)
    VR_Mega068Car3.PlayAnimEndless (0.4)
  End If
  If VRRoomChoice = 2 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRMEGARoom : VR_Obj.Visible = 0 : Next
    If LogoPosterMod = 0 Then
      VR_Logo.Visible = False
      VR_Poster.Visible = False
    End If
    If LogoPosterMod = 1 Then
      VR_Logo.Visible = True
      VR_Poster.Visible = True
    End If
    If LogoPosterMod = 2 Then
      VR_Logo.Visible = True
      VR_Poster.Visible = False
    End If
    If LogoPosterMod = 3 Then
      VR_Logo.Visible = False
      VR_Poster.Visible = True
    End If
    Room360.Visible = 0
  End If
  If VRRoomChoice = 3 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRMEGARoom : VR_Obj.Visible = 0 : Next
    Room360.Visible = 1
  End If
Else
  VRMode = False
  Room360.Visible = 0
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMEGARoom : VR_Obj.Visible = 0 : Next
End If

'*******************************************
' ZTIM:Timers
'*******************************************

Dim FrameTime, InitFrameTime
InitFrameTime = 0

Dim BumperOneIntensity, BumperTwoIntensity, BumperThreeIntensity
Dim LastBumperOneIntensity, LastBumperTwoIntensity, LastBumperThreeIntensity
LastBumperOneIntensity = -1
LastBumperTwoIntensity = -1
LastBumperThreeIntensity = -1

Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  RollingUpdate     ' Update rolling sounds
  DoDTAnim    'handle drop target animations
  DoSTAnim    'handle stand up target animations
  If AmbientBallShadowOn = 1 Then
    BSUpdate
  End If
  UpdateBallBrightness
  UpdatePlunger
  UpdateRainbowGI
  If SidewallFlashOpt = 1 or SidewallFlashOpt = 3 Then
    Flasherflash1b.opacity = Flasherflash1.opacity
    Flasherflash1c.opacity = Flasherflash1.opacity
    Flasherflash2b.opacity = Flasherflash2.opacity
    Flasherflash2c.opacity = Flasherflash2.opacity
    Flasherflash3a.opacity = Flasherflash3.opacity
    Flasherflash3b.opacity = Flasherflash3.opacity
    Flasherflash3c.opacity = Flasherflash3.opacity
    Flasherflash3d.opacity = Flasherflash3.opacity
    Flasherflash4a.opacity = Flasherflash4.opacity
    Flasherflash4b.opacity = Flasherflash4.opacity
    Flasherflash4c.opacity = Flasherflash4.opacity
    Flasherflash4d.opacity = Flasherflash4.opacity
    Flasherflash5a.opacity = Flasherflash5.opacity
    Flasherflash5b.opacity = Flasherflash5.opacity
    Flasherflash5c.opacity = Flasherflash5.opacity
    Flasherflash5d.opacity = Flasherflash5.opacity
    Flasherflash6b.opacity = Flasherflash6.opacity
    Flasherflash6c.opacity = Flasherflash6.opacity

    Flasherflash1b.visible = Flasherflash1.visible
    Flasherflash1c.visible = Flasherflash1.visible
    Flasherflash2b.visible = Flasherflash2.visible
    Flasherflash2c.visible = Flasherflash2.visible
    Flasherflash3a.visible = Flasherflash3.visible
    Flasherflash3b.visible = Flasherflash3.visible
    Flasherflash3c.visible = Flasherflash3.visible
    Flasherflash3d.visible = Flasherflash3.visible
    Flasherflash4a.visible = Flasherflash4.visible
    Flasherflash4b.visible = Flasherflash4.visible
    Flasherflash4c.visible = Flasherflash4.visible
    Flasherflash4d.visible = Flasherflash4.visible
    Flasherflash5a.visible = Flasherflash5.visible
    Flasherflash5b.visible = Flasherflash5.visible
    Flasherflash5c.visible = Flasherflash5.visible
    Flasherflash5d.visible = Flasherflash5.visible
    Flasherflash6b.visible = Flasherflash6.visible
    Flasherflash6c.visible = Flasherflash6.visible
  Else
    Flasherflash1b.visible = False
    Flasherflash1c.visible = False
    Flasherflash2b.visible = False
    Flasherflash2c.visible = False
    Flasherflash3a.visible = False
    Flasherflash3b.visible = False
    Flasherflash3c.visible = False
    Flasherflash3d.visible = False
    Flasherflash4a.visible = False
    Flasherflash4b.visible = False
    Flasherflash4c.visible = False
    Flasherflash4d.visible = False
    Flasherflash5a.visible = False
    Flasherflash5b.visible = False
    Flasherflash5c.visible = False
    Flasherflash6b.visible = False
    Flasherflash6c.visible = False
    Flasherflash5d.visible = False
  End If
  BumperOneIntensity = l65a.GetInPlayIntensity / 15
  If BumperOneIntensity <> LastBumperOneIntensity Then
    FlFadeBumper 1,BumperOneIntensity
    LastBumperOneIntensity = BumperOneIntensity
  End If
  BumperTwoIntensity = l66a.GetInPlayIntensity / 15
  If BumperTwoIntensity <> LastBumperTwoIntensity Then
    FlFadeBumper 2,BumperTwoIntensity
    LastBumperTwoIntensity = BumperTwoIntensity
  End If
  BumperThreeIntensity = l67a.GetInPlayIntensity / 15
  If BumperThreeIntensity <> LastBumperThreeIntensity Then
    FlFadeBumper 3,BumperThreeIntensity
    LastBumperThreeIntensity = BumperThreeIntensity
  End If
  If li86.state > 0.05 Then
    Pincab_Start_Button.disablelighting = 0.5
  Else
    Pincab_Start_Button.disablelighting = 0
  End if
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

Sub WheelTimer_Timer(): WheelDrive: End Sub

LoadVPM "01560000", "WPC.VBS", 3.50

Sub UpdatePlunger()
  If plungerpress = 1 then
    If Pincab_Plunger.Y < 2410.343 then
      Pincab_Plunger.Y = Pincab_Plunger.Y + 5
    End If
  Else
    Pincab_Plunger.Y = 2310.343 + (5* Plunger.Position) -20
  End If
End Sub

' ****************************************************
' table init
' ****************************************************

Dim HurricaneBall1, HurricaneBall2, HurricaneBall3, gBOT

Sub Table1_Init()
  vpmInit Me
  With Controller
        .GameName = cGameName
        .SplashInfoLine = "Hurricane (Williams 1991)"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 1
    .Hidden = 0

    If cDMDRotation >= 0 Then .Games(cGameName).Settings.Value("rol") = cDMDRotation

    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With

    On Error Goto 0

  ' nudging
  vpmNudge.TiltSwitch  = 14
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj   = Array(Bumper51, Bumper52, Bumper53, LeftSlingShot, RightSlingShot)

  '************  Trough **************
  Set HurricaneBall1 = sw16.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set HurricaneBall2 = sw17.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set HurricaneBall3 = sw18.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(HurricaneBall1,HurricaneBall2, HurricaneBall3)

  Controller.Switch(16) = 1
  Controller.Switch(17) = 1
  Controller.Switch(18) = 1

  ' jugglers ball stack
  Set bsRightJuggler = new cvpmBallStack
  bsRightJuggler.InitSaucer RightJuggler, 57, 225, 2
  bsRightJuggler.InitExitSnd SoundFX("fx_popper_ball",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
  bsRightJuggler.KickForceVar = 3
  bsRightJuggler.KickAngleVar = 3

    ' init timers
  PinMAMETimer.Interval         = PinMAMEInterval
    PinMAMETimer.Enabled          = True

  FerrisWheelTimer.Interval       = 15
  FerrisWheelTimer.Enabled      = False
  BallWaiting4BlueWheelTimer.Interval = 10
  BallWaiting4BlueWheelTimer.Enabled  = False
  BallWaiting4RedWheelTimer.Interval  = 10
  BallWaiting4RedWheelTimer.Enabled = False
  sw33.TimerInterval          = 100
  sw33.TimerEnabled           = False

  FlFadeBumper 1,1
  FlFadeBumper 2,1
  FlFadeBumper 3,1
End Sub

Sub Table1_Paused()
  Controller.Pause = 1
End Sub
Sub Table1_UnPaused()
  Controller.Pause = 0
End Sub
Sub Table1_Exit()
  Controller.Stop
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

' ****************************************************
' keys
' ****************************************************

Sub Table1_KeyDown(ByVal keycode)
  If Keycode = StartGameKey Then
        SoundStartButton
    'StartButton.y = StartButton.y -5
    'StartButton2.y = StartButton2.y -5
  End If
    If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft
    If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight
    If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter
  If keycode = LeftFlipperKey And FlippersEnabled Then SolLFlipper(True) : Controller.Switch(12) = 1
  If keycode = RightFlipperKey And FlippersEnabled Then SolRFlipper(True) : Controller.Switch(11) = 1
  If keycode = LeftFlipperKey then PinCab_Flipper_Button_Left.X = PinCab_Flipper_Button_Left.X + 8
  If keycode = RightFlipperKey then PinCab_Flipper_Button_Right.X = PinCab_Flipper_Button_Right.X - 8
  If keycode = StartGameKey then Pincab_Start_Button.y = Pincab_Start_Button.y - 3
  If keycode = StartGameKey then Pincab_Start_Button_start.y = Pincab_Start_Button_start.y - 3
    If keycode = PlungerKey Then
    SoundPlungerPull
    Plunger.Pullback
    plungerpress = 1
  End If
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If KeyCode = PlungerKey Then
    Plunger.Fire
    plungerpress = 0
    If BIPL = 1 Then
      SoundPlungerReleaseBall()   'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall() 'Plunger release sound when there is no ball in shooter lane
    End If
  End If
  If keycode = LeftFlipperKey And FlippersEnabled Then SolLFlipper(False) : Controller.Switch(12) = 0
  If keycode = RightFlipperKey And FlippersEnabled Then SolRFlipper(False) : Controller.Switch(11) = 0
  If keycode = LeftFlipperKey then PinCab_Flipper_Button_Left.X = PinCab_Flipper_Button_Left.X - 8
  If keycode = RightFlipperKey then PinCab_Flipper_Button_Right.X = PinCab_Flipper_Button_Right.X + 8
  If keycode = StartGameKey then Pincab_Start_Button.y = Pincab_Start_Button.y + 3
  If keycode = StartGameKey then Pincab_Start_Button_Start.y = Pincab_Start_Button_start.y + 3
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 1000       'X position of punger lane left
Const PLRight = 1060      'X position of punger lane right
Const PLTop = 1225        'Y position of punger lane top
Const PLBottom = 1900       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,PLLeft,PLBottom,PLLeft,PLTop,PLRight,PLTop,PLRight,PLBottom) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-PLBottom))
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

'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 245 + 10)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next
End Sub

Dim SavedMtlColorArray
SaveMtlColors
Sub SaveMtlColors
  ReDim SavedMtlColorArray(UBound(RoomBrightnessMtlArray))
  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    SaveMaterialBaseColor RoomBrightnessMtlArray(i), i
  Next
End Sub

Sub SaveMaterialBaseColor(name, idx)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  SavedMtlColorArray(idx) = round(base,0)
End Sub


Sub ModulateMaterialBaseColor(name, idx, val)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  Dim red, green, blue, saved_base, new_base

  'First get the existing material properties
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  'Get saved color
  saved_base = SavedMtlColorArray(idx)

  'Next extract the r,g,b values from the base color
  red = saved_base And &HFF
  green = (saved_base \ &H100) And &HFF
  blue = (saved_base \ &H10000) And &HFF
  'msgbox red & " " & green & " " & blue

  'Create new color scaled down by 'val', and update the material
  new_base = RGB(red*val, green*val, blue*val)
  UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, new_base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub

' ****************************************************
' *** solenoids
' ****************************************************
SolCallback(1) = "SolMechDrive"
SolCallback(2) = "SolDTBank"
SolCallback(4) = "SolLeftJuggler"
SolCallback(5) = "SolRightJuggler"
SolCallback(6) = "SolWheelDrive"
SolCallBack(7) = "SolKnocker"
SolCallback(9) = "SolOuthole"
SolCallback(10) = "SolBallRelease"

'Solenoid Controlled Flashers

SolModCallback(17) = "SolModFlashPWM17"       'Right Side Red Flashers
SolModCallback(18) = "SolModFlashPWM18"       'Top Right White and Blue Flashers
SolModCallback(19) = "vpmFlasher fLight19,"     'All Scores 5X Flasher
SolModCallback(20) = "vpmFlasher fLight20,"     'Comet Mil Flasher
SolModCallback(21) = "vpmFlasher fLight21,"     'Jackpot Flasher
SolModCallback(22) = "vpmFlasher fLight22,"     'Ferris Wheel Flasher
SolModCallback(23) = "SolModFlashPWM23"       'Top Left Red Flashers
SolModCallback(24) = "SolModFlashPWM24"       'Bottom Left White Flasher and Backbox Topper Left Flasher
SolModCallback(25) = "SolModFlashPWM25"       'Bottom Right White Flasher and Backbox Topper Right Flasher
SolModCallback(26) = "SolModFlashPWM26"       'Jet Bumper Flasher
SolModCallback(27) = "SolModFlashPWM27"       'Dunk Dummy White Flasher
SolModCallback(28) = "SolModFlashPWM28"       'Left Side White Flasher

SolCallback(31) = "TiltSol" ' 31 for WPC

'*******************************************
' ZSOL: Other Solenoids
'*******************************************

Sub TiltSol(Enabled)
  FlippersEnabled = Enabled
  SolLFlipper(False)
  SolRFlipper(False)
End Sub

Dim WheelMech, WheelSpeed, WheelNewPos : WheelMech = 0: WheelSpeed = 0:WheelNewPos = 0

Sub SolMechDrive(Enabled)
  If enabled Then
    PlaySoundAt SoundFX("WheelMotor",DOFGear), Mystery_Wheel_Reel
    WheelMech = 1
  Else
    StopSound "WheelMotor"
    WheelMech = 0
  End If
End Sub

Sub WheelDrive()
  If WheelMech = 1 Then WheelSpeed = 168
  If WheelMech = 0 Then
    If WheelSpeed > 0 Then
      WheelSpeed = WheelSpeed - 1.75
    Else
      WheelSpeed = 0
    End If
  End If
  WheelNewPos = WheelNewPos - (WheelSpeed / 100)
  If WheelNewPos <0 Then WheelNewPos = 95
  If WheelNewPos> 95 Then WheelNewPos = 0
  Mystery_Wheel_Reel.ObjRotY = WheelNewPos * 3.75 'value is due to 96 steps for 360 degrees rotation
  WheelGI1Warm.ObjRotY = WheelNewPos * 3.75 'value is due to 96 steps for 360 degrees rotation
  WheelGI2Warm.ObjRotY = WheelNewPos * 3.75 'value is due to 96 steps for 360 degrees rotation
  WheelGI3Warm.ObjRotY = WheelNewPos * 3.75 'value is due to 96 steps for 360 degrees rotation
End Sub

Sub SolKnocker(Enabled) 'Knocker solenoid
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub

Sub CheckGameOver(isGameOver)
  Bumper51.Collidable       = Not isGameOver
  Bumper52.Collidable       = Not isGameOver
  Bumper53.Collidable       = Not isGameOver
  BlockWallGameOver.Collidable  = isGameOver
  FlippersEnabled         = Not isGameOver
End Sub

'******************************************************
'     TROUGH BASED ON FOZZY'S
'******************************************************

Sub sw16_Hit() : Controller.Switch(16) = 1 : UpdateTrough : BIP = BIP + 1  : End Sub
Sub sw16_UnHit() : Controller.Switch(16) = 0 : UpdateTrough : End Sub
Sub sw17_Hit() : Controller.Switch(17) = 1 : UpdateTrough : End Sub
Sub sw17_UnHit() : Controller.Switch(17) = 0 : UpdateTrough : End Sub
Sub sw18_Hit() : Controller.Switch(18) = 1 : UpdateTrough : End Sub
Sub sw18_UnHit() : Controller.Switch(18) = 0 : UpdateTrough : End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 150
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw16.BallCntOver = 0 Then sw17.kick 60, 10
  If sw17.BallCntOver = 0 Then sw18.kick 60, 10
  Me.Enabled = 0
End Sub

' ******************************************************
' outhole, drain and ball release
' ******************************************************

Sub SolOuthole(Enabled)
  If Enabled Then
    sw15.kick 90, 12
    RandomSoundDrainKicker sw15
    UpdateTrough
  End If
End Sub

Sub SolBallRelease(Enabled)
  If Enabled Then
    RandomSoundBallRelease sw16
    sw16.kick 60, 12
    UpdateTrough
  End If
End Sub

Sub sw15_Hit()
  Controller.Switch(15) = 1
  RandomSoundDrain sw15
  BIP = BIP - 1
End Sub

Sub sw15_UnHit()
  Controller.Switch(15) = 0
End Sub

'*******************************************
' ZFLP: Flippers
'*******************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20
Const QuickFlipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend

    If LeftFlipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
      'Play partial flip sound and stop any flip down sound
      StopAnyFlipperLowerLeftDown()
      RandomSoundFlipperLowerLeftReflip LeftFlipper
    Else
      'Play full flip sound
      If BallNearLF = 0 Then
        RandomSoundFlipperLowerLeftUpFullStroke LeftFlipper
      End If
      If BallNearLF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerLeftUpDampenedStroke LeftFlipper
          Case 2 : RandomSoundFlipperLowerLeftUpFullStroke LeftFlipper
        End Select
      End If
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      'Play flip down sound
      RandomSoundFlipperLowerLeftDown LeftFlipper
    End If
    If LeftFlipper.currentangle < LeftFlipper.startAngle + QuickFlipAngle and LeftFlipper.currentangle <> LeftFlipper.endangle Then
      'Play quick flip sound and stop any flip up sound
      StopAnyFlipperLowerLeftUp()
      RandomSoundLowerLeftQuickFlipUp()
    Else
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    End If
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire 'rightflipper.rotatetoend

    If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
      'Play partial flip sound and stop any flip down sound
      StopAnyFlipperLowerRightDown()
      RandomSoundFlipperLowerRightReflip RightFlipper
    Else
      'Play full flip sound
      If BallNearRF = 0 Then
        RandomSoundFlipperLowerRightUpFullStroke RightFlipper
      End If

      If BallNearRF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerRightUpDampenedStroke RightFlipper
          Case 2 : RandomSoundFlipperLowerRightUpFullStroke RightFlipper
        End Select
      End If
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      'Play flip down sound
      RandomSoundFlipperLowerRightDown RightFlipper
    End If
    If RightFlipper.currentangle < RightFlipper.startAngle + QuickFlipAngle and RightFlipper.currentangle <> RightFlipper.endangle Then
      'Play quick flip sound and stop any flip up sound
      StopAnyFlipperLowerRightUp()
      RandomSoundLowerRightQuickFlipUp()
    Else
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    End If
  End If
End Sub

'///////////////////////  FLIPPER BATS BALL COLLIDE SOUND  //////////////////////
dim angdamp, veldamp
angdamp = 0.2
veldamp = 0.8

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperLeftLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperLeftLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperRightLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperRightLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

  RandomSoundRubberFlipper(parm)
End Sub

'******************************************************
'  ZANI: Misc Animations
'******************************************************

Sub LeftFlipper_Animate
    dim a: a = LeftFlipper.CurrentAngle
    FlipperShadowL.ObjRotZ = a
    batleft.ObjRotZ = a
End Sub

Sub RightFlipper_Animate
    dim a: a = RightFlipper.CurrentAngle
    FlipperShadowR.ObjRotZ = a
    batright.ObjRotZ = a
End Sub

' ****************************************************
' sling shots and animations
' ****************************************************
Dim LeftStep, RightStep

Sub LeftSlingShot_Slingshot()
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 36
  RandomSoundSlingshotLeft LeftSlingHammer
    LeftStep = 0
  LeftSlingShot.TimerInterval = 20
    LeftSlingShot.TimerEnabled  = True
  LeftSlingShot_Timer
End Sub
Sub LeftSlingShot_Timer()
    Select Case LeftStep
    Case 0: LeftSling1.Visible = False : LeftSling3.Visible = True : LeftSlingHammer.TransZ = -28
        Case 1: LeftSling3.Visible = False : LeftSling2.Visible = True : LeftSlingHammer.TransZ = -10
        Case 2: LeftSling2.Visible = False : LeftSling1.Visible = True : LeftSlingHammer.TransZ = 0 : LeftSlingShot.TimerEnabled = 0
    End Select
    LeftStep = LeftStep + 1
End Sub

Sub RightSlingShot_Slingshot()
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 25
  RandomSoundSlingshotRight RightSlingHammer
    RightStep = 0
  RightSlingShot.TimerInterval = 20
    RightSlingShot.TimerEnabled  = True
  RightSlingShot_Timer
End Sub
Sub RightSlingShot_Timer()
    Select Case RightStep
    Case 0: RightSling1.Visible = False : RightSling3.Visible = True : RightSlingHammer.TransZ = -28
        Case 1: RightSling3.Visible = False : RightSling2.Visible = True : RightSlingHammer.TransZ = -10
        Case 2: RightSling2.Visible = False : RightSling1.Visible = True : RightSlingHammer.TransZ = 0 : RightSlingShot.TimerEnabled = 0
    End Select
    RightStep = RightStep + 1
End Sub

' ****************************************************
' switches
' ****************************************************
Dim Drop33 : Drop33 = 0
Dim Drop34 : Drop34 = 0
Dim Drop35 : Drop35 = 0
' drop-down targets
Sub sw33_Hit
  sw33.TimerEnabled = True
  DTHit 33
  Drop33 = 1
  TargetBouncer ActiveBall, 1
End Sub

Sub sw34_Hit
  sw33.TimerEnabled = True
  DTHit 34
  Drop34 = 1
  TargetBouncer ActiveBall, 1
End Sub

Sub sw35_Hit
  sw33.TimerEnabled = True
  DTHit 35
  Drop35 = 1
  TargetBouncer ActiveBall, 1
End Sub

Sub SolDTBank(Enabled)
  If enabled Then
    RandomSoundDropTargetBankReset sw34p
    DTRaise 33
    DTRaise 34
    DTRaise 35
    Drop33 = 0
    Drop34 = 0
    Drop35 = 0
    sw33.TimerEnabled = True
  End If
End Sub

Dim counter
Sub sw33_Timer()
  counter = counter + 1
  If counter > 10 Then counter = 0 : sw33.TimerEnabled = False
  UpdateGILightsBehindDropTargets
End Sub

' stand-up targets
Sub sw42_Hit : STHit 42 : End Sub
Sub sw43_Hit : STHit 43 : End Sub
Sub sw44_Hit : STHit 44 : End Sub
Sub sw45_Hit : STHit 45 : End Sub
Sub sw55_Hit : STHit 55 : End Sub

Sub sw42o_Hit
  TargetBouncer ActiveBall, 1
End Sub
Sub sw43o_Hit
  TargetBouncer ActiveBall, 1
End Sub
Sub sw44o_Hit
  TargetBouncer ActiveBall, 1
End Sub
Sub sw45o_Hit
  TargetBouncer ActiveBall, 1
End Sub
Sub sw55o_Hit
  TargetBouncer ActiveBall, 1
End Sub

' inlanes and outlanes
Sub sw26_Hit :   Controller.Switch(26) = 1 : End Sub
Sub sw26_Unhit : Controller.Switch(26) = 0 : End Sub
Sub sw27_Hit :   Controller.Switch(27) = 1 : End Sub
Sub sw27_Unhit : Controller.Switch(27) = 0 : End Sub
Sub sw37_Hit :   Controller.Switch(37) = 1 : End Sub
Sub sw37_Unhit : Controller.Switch(37) = 0 : End Sub
Sub sw38_Hit :   Controller.Switch(38) = 1 : End Sub
Sub sw38_Unhit : Controller.Switch(38) = 0 : End Sub

' plunger switch
Sub sw28_Hit   : Controller.Switch(28) = 1 : End Sub
Sub sw28_UnHit : Controller.Switch(28) = 0 : End Sub

' bumpers
Sub Bumper51_Hit: vpmTimer.PulseSw 51 : RandomSoundBumperTop Bumper51 : End Sub
Sub Bumper52_Hit: vpmTimer.PulseSw 52 : RandomSoundBumperMiddle Bumper52 : End Sub
Sub Bumper53_Hit: vpmTimer.PulseSw 53 : RandomSoundBumperBottom Bumper53 : End Sub

' ferris wheel switch
Sub sw31_Hit   : Controller.Switch(31) = True  : End Sub
Sub sw31_UnHit : Controller.Switch(31) = False : End Sub

' ramp triggers
Sub sw61_Hit: vpmTimer.PulseSw 61 : End Sub
Sub sw62_Hit: vpmTimer.PulseSw 62 : End Sub
Sub sw63_Hit: vpmTimer.PulseSw 63 : End Sub
Sub sw64_Hit: vpmTimer.PulseSw 64 : End Sub

Sub ramptrigger01_Hit()
  If ActiveBall.VelY < 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub
Sub ramptrigger02_Hit()
  WireRampOff
End Sub
Sub ramptrigger02_Unhit()
  WireRampOn True
End Sub
Sub ramptrigger03_Hit()
  WireRampOff
End Sub
Sub ramptrigger03_Unhit()
  WireRampOn False
End Sub
'Sub ramptrigger04_Hit()
' leftInlaneSpeedLimit
'End Sub
Sub ramptrigger04_Unhit()
  WireRampOff
End Sub
Sub ramptrigger05_Hit()
  If (ActiveBall.VelY > 0) Then
    'ball is traveling down the playfield
    RandomSoundRampFlapDown()
  ElseIf (ActiveBall.VelY < 0) Then
    RandomSoundRampFlapUp()

  End If
  WireRampOn True
End Sub
Sub ramptrigger05_unhit()
  If ActiveBall.VelY > 0 Then WireRampOff : End If
End Sub
Sub ramptrigger06_Hit()
  WireRampOff
End Sub
Sub ramptrigger06_Unhit()
  WireRampOn False
End Sub
Sub ramptrigger07_Unhit()
  WireRampOff
  rightInlaneSpeedLimit
End Sub
Sub ramptrigger08_Hit()
  If (ActiveBall.VelY > 0) Then
    'ball is traveling down the playfield
    RandomSoundRampFlapDown()
  ElseIf (ActiveBall.VelY < 0) Then
    RandomSoundRampFlapUp()
  End If
  WireRampOn True
End Sub
Sub ramptrigger08_unhit()
  If ActiveBall.VelY > 0 Then WireRampOff : End If
End Sub
Sub ramptrigger09_Hit()
  WireRampOn True
End Sub

Sub leftInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  if activeball.vely < 0 then exit sub              'don't affect upwards movement
    activeball.AngMomZ = -abs(activeball.AngMomZ) * RndNum(3,6)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

    if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


Sub rightInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  if activeball.vely < 0 then exit sub              'don't affect upwards movement
    activeball.AngMomZ = abs(activeball.AngMomZ) * RndNum(2,4)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

  if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub

' *********************************************************************
' the juggler
' *********************************************************************

Dim jugglerBall, jugglerStep
Dim LJx, LJy, RJx, RJy, jArcC(3), jSkew, jxmod, jymod
LJx = LeftJuggler.x
LJy = LeftJuggler.y
RJx = RightJuggler.x
RJy = RightJuggler.y
jSkew = myMath.dAtn2((RJx - LJx),(RJy - LJy))
jxmod = myMath.dSin(jSkew)
jymod = -myMath.dCos(jSkew)
jArcC(0) = 90
jArcC(1)=LJx+jxmod*jArcC(0)
jArcC(2)=LJy+jymod*jArcC(0)
jArcC(3)=127

Sub LeftJuggler_Hit()
  Set jugglerBall = ActiveBall
  Controller.Switch(56) = True
  PlaySoundAt "fx_kicker_enter_center", LeftJuggler
End Sub

Sub SolLeftJuggler(Enabled)
  If Enabled Then
    If Controller.Switch(56) Then
      PlaySoundAtVol SoundFX("fx_popper_ball",DOFContactors), LeftJuggler, 2
      Controller.Switch(56) = False
      jugglerStep = 1
      LeftJuggler.TimerInterval = 10
      LeftJuggler.TimerEnabled  = True
    End If
  End If
End Sub

Sub LeftJuggler_Timer()
  If jugglerStep = 28 Then
    Me.Kick 0, 1
    Me.TimerEnabled = False
    jugglerBall.X = RightJuggler.X
    jugglerBall.Y = RightJuggler.Y
    jugglerBall.Z = myMath.dSin((jugglerStep-1)*7.5)*jArcC(0) + jArcC(3)
  Else
    jugglerBall.X = jArcC(1) - myMath.dCos((jugglerStep-1)*7.5)*jArcC(0)*jxmod
    jugglerBall.Y = jArcC(2) - myMath.dCos((jugglerStep-1)*7.5)*jArcC(0)*jymod
    jugglerBall.Z = myMath.dSin((jugglerStep-1)*7.5)*jArcC(0) + jArcC(3)
    jugglerStep = jugglerStep + 1
  End If
End Sub

Sub RightJuggler_Hit()
  bsRightJuggler.AddBall Me
  PlaySoundAt "fx_kicker_enter_center", RightJuggler
End Sub
Sub SolRightJuggler(Enabled)
  If Enabled Then
    bsRightJuggler.SolOut(True)
  End If
End Sub

' *********************************************************************
' Ferris wheel mech
' *********************************************************************

' init
Dim isWheelMoving : isWheelMoving = False '
Dim ballBlueWheel(2) : Set ballBlueWheel(0) = Nothing : Set ballBlueWheel(1) = Nothing : Set ballBlueWheel(2) = Nothing
Dim ballKickerBlueWheel(2) : Set ballKickerBlueWheel(0) = KickerFerrisWheelBlue1 : Set ballKickerBlueWheel(1) = KickerFerrisWheelBlue2 : Set ballKickerBlueWheel(2) = KickerFerrisWheelBlue3
Dim ballInfoBlueWheel(2,1)
Dim ballRedWheel(2) : Set ballRedWheel(0)  = Nothing : Set ballRedWheel(1)  = Nothing : Set ballRedWheel(2)  = Nothing
Dim ballKickerRedWheel(2) : Set ballKickerRedWheel(0) = KickerFerrisWheelRed1 : Set ballKickerRedWheel(1) = KickerFerrisWheelRed2 : Set ballKickerRedWheel(2) = KickerFerrisWheelRed3
Dim ballInfoRedWheel(2,1)
myMath.WheelAngle   = 29
myMath.WheelRadius  = 58
Const ExitAngle   = 195
FerrisWheelRed.RotX = Int(Rnd*45)

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub SolWheelDrive(Enabled)
  If Enabled then
    FerrisWheelTimer.Enabled = True
    PlayLoopSoundAtVol "fx_motor", KickerFerrisWheelBlue1, 0.1
    Else
        FerrisWheelTimer.Enabled = False
    StopSound "fx_motor"
    End If
  isWheelMoving = Enabled : CheckGameOver(Not isWheelMoving)
End Sub

Sub SlowDownTimer(angle)
  Dim timer
  timer = FerrisWheelTimer.Interval
  If angle >= ExitAngle Then
    timer = 15
  ElseIf angle > 185 Then
    timer = 4
  ElseIf angle > 175 Then
    timer = timer - 2 : If timer < 4 Then timer = 4
  End If
  FerrisWheelTimer.Interval = timer
End Sub

Sub FerrisWheelTimer_Timer()
  ' turn the wheels
  FerrisWheelBlue.RotX = (FerrisWheelBlue.RotX + 1) MOD 360
  FerrisWheelRed.RotX  = (FerrisWheelRed.RotX + 1) MOD 360

  ' check whether a ball is able to get into one of the wheels
  Dim rotX
  rotX = FerrisWheelBlue.RotX MOD 90
  FerrisWheelBlueBlockingWall.Collidable = Not (rotX >= 0 And rotX <= 15)
  rotX = FerrisWheelRed.RotX MOD 90
  FerrisWheelRedBlockingWall.Collidable  = Not (rotX >= 40 And rotX <= 55)

  ' get thru all balls of the blue and red wheel (three possible balls per wheel)
  Dim ii
  For ii = 0 To 2
    ' check whether a ball can be picked up in a blue or red slot
    If Not ballBlueWheel(ii) Is Nothing And ballInfoBlueWheel(ii,0) Then
      SlowDownTimer ballInfoBlueWheel(ii,1)
      ' check whether ball is at exit angle
      If ballInfoBlueWheel(ii,1) >= ExitAngle Then
        ' get the ball out
        ballInfoBlueWheel(ii,0) = False
      End If
      ' increment ball's wheel angle according to the wheel movement
      ballInfoBlueWheel(ii,1) = ballInfoBlueWheel(ii,1) + 1
    End If
    If Not ballRedWheel(ii) Is Nothing And ballInfoRedWheel(ii,0) Then
      SlowDownTimer ballInfoRedWheel(ii,1)
      ' check whether ball is at exit angle
      If ballInfoRedWheel(ii,1) >= ExitAngle Then
        ' get the ball out
        ballInfoRedWheel(ii,0) = False
      End If
      ' increment ball's wheel angle according to the wheel movement
      ballInfoRedWheel(ii,1) = ballInfoRedWheel(ii,1) + 1
    End If

    ' move the balls in the wheels
    Dim ballX, ballY, ballZ
    If Not ballBlueWheel(ii) Is Nothing Then
      If ballInfoBlueWheel(ii,0) Then
        ' calc real ball position and move the ball
        myMath.BallPos ballInfoBlueWheel(ii,1), ballX, ballY, ballZ
        ballBlueWheel(ii).X     = FerrisWheelBlue.X + ballX
        ballBlueWheel(ii).Y     = FerrisWheelBlue.Y + ballY
        ballBlueWheel(ii).Z     = FerrisWheelBlue.Z + ballZ + ballsize/2
      Else
        ' maybe kick and reset the ball
        ballBlueWheel(ii).Z       = ballBlueWheel(ii).Z - 30
        ballKickerBlueWheel(ii).Kick 209, 4
        ballKickerBlueWheel(ii).Enabled = True
        Set ballBlueWheel(ii)       = Nothing
        PlaySoundAtVol "ferris_ball_drop", KickerFerrisWheelBlue1, 0.5
      End If
    End If
    If Not ballRedWheel(ii) Is Nothing Then
      If ballInfoRedWheel(ii,0) Then
        ' calc real ball position and move the ball
        myMath.BallPos ballInfoRedWheel(ii,1), ballX, ballY, ballZ
        ballRedWheel(ii).X    = FerrisWheelRed.X + ballX
        ballRedWheel(ii).Y    = FerrisWheelRed.Y + ballY
        ballRedWheel(ii).Z    = FerrisWheelRed.Z + ballZ + ballsize/2
      Else
        ' maybe kick and reset the ball
        ballRedWheel(ii).Z        = ballRedWheel(ii).Z - 30
        ballKickerRedWheel(ii).Kick 209, 2
        ballKickerRedWheel(ii).Enabled  = True
        Set ballRedWheel(ii)      = Nothing
        PlaySoundAtVol "ferris_ball_drop", KickerFerrisWheelRed1, 0.5
      End If
    End If
  Next
End Sub

Sub KickerFerrisWheelBlue_Hit(idx)
  PlaySoundAtVol "ferris_hit2", KickerFerrisWheelBlue1, 2
  ' disable current kicker
  ballKickerBlueWheel(idx).Enabled = False
  ' pick up the waiting ball
  If ballBlueWheel(idx) Is Nothing Then
    Set ballBlueWheel(idx)    = ActiveBall
    ballInfoBlueWheel(idx,0)  = True ' shows that a ball is in the slot
    ballInfoBlueWheel(idx,1)  = 25   ' starting angle of the ball in the wheel
  End If
End Sub
Sub KickerFerrisWheelRed_Hit(idx)
  PlaySoundAtVol "ferris_hit2", KickerFerrisWheelRed1, 2
  ' disable current kicker
  ballKickerRedWheel(idx).Enabled = False
  ' pick up the waiting ball
  If ballRedWheel(idx) Is Nothing Then
    Set ballRedWheel(idx)     = ActiveBall
    ballInfoRedWheel(idx,0)   = True ' shows that a ball is in the slot
    ballInfoRedWheel(idx,1)   = 60   ' starting angle of the ball in the wheel
  End If
End Sub

'****************************************************************
' ZGII: GI
'****************************************************************

Set GiCallback2 = GetRef("UpdateGI")

Dim currentGIOneLevel, currentGITwoLevel, currentGIBOneLevel, currentGIBTwoLevel, currentGIBThreeLevel, isGIOn

InitGI

' general illumination
Sub ResetGI()
  InitGI
  UpdateGI 1, currentGIBOneLevel
  UpdateGI 2, currentGIOneLevel
  UpdateGI 3, currentGIBTwoLevel
  UpdateGI 4, currentGITwoLevel
  UpdateGI 5, currentGIBThreeLevel
End Sub
Sub InitGI()
  isGIOn = False
  Dim coll, obj
  ' init GI lights
  For Each coll in Array(GIBottomLeft,GITopLeft)
    For Each obj In coll
      obj.IntensityScale = 0
      If GIColorMod = 1 Or GIColorMod = 5  Or GIColorMod = 6 Then
        obj.Color   = Red
        obj.ColorFull = RedFull
      ElseIf GIColorMod = 2 Then
        obj.Color   = Blue
        obj.ColorFull = BlueFull
      ElseIf GIColorMod = 3 Then
        obj.Color   = Green
        obj.ColorFull = GreenFull
      ElseIf GIColorMod = 4 Then
        obj.Color   = Orange
        obj.ColorFull = OrangeFull
      ElseIf GIColorMod = 7 Then
        obj.color = RGB(rRed, rGreen, rBlue)
        obj.colorfull = RGB(rRed, rGreen, rBlue)
      Else
        obj.Color   = White
        obj.ColorFull = WhiteFull
      End If
    Next
  Next
  For Each coll in Array(GIBottomRight,GITopRight)
    For Each obj In coll
      obj.IntensityScale = 0
      If GIColorMod = 2 Or GIColorMod = 5 Then
        obj.Color   = Blue
        obj.ColorFull = BlueFull
      ElseIf GIColorMod = 1 Then
        obj.Color   = Red
        obj.ColorFull = RedFull
      ElseIf GIColorMod = 3 Or GIColorMod = 6 Then
        obj.Color   = Green
        obj.ColorFull = GreenFull
      ElseIf GIColorMod = 4 Then
        obj.Color   = Orange
        obj.ColorFull = OrangeFull
      ElseIf GIColorMod = 7 Then
        obj.color = RGB(rRed, rGreen, rBlue)
        obj.colorfull = RGB(rRed, rGreen, rBlue)
      Else
        obj.Color   = White
        obj.ColorFull = WhiteFull
      End If
    Next
  Next
End Sub

Dim gilvl   'General Illumination light state tracked for Dynamic Ball Shadows
gilvl = 1

Sub UpdateGI(GINo, GILevel)
  ' GINo: 2 = bottom GI, 4 = top GI
  ' GILevel: value between 0 and 8
  Select Case GINo
    Case 0 : UpdateGIBackglassOne GILevel
    Case 1 : UpdateGIBackglassTwo GILevel
    Case 2 : UpdateGIBottom GILevel
    Case 3 : UpdateGIBackglassThree GILevel
    Case 4 : UpdateGITop GILevel
  End Select
  If isWheelMoving Then
    If GILevel = 0 Then
      isGIOn = False
      DOF 101, DOFOff
      gilvl = 0
    Else
      isGIOn = True
      DOF 101, DOFOn
      gilvl = 1
    End If
  End If
End Sub

Sub UpdateGIBackglassOne(level)
  Dim coll, obj
  For Each coll in Array(GIBackglassOne)
    For Each obj In coll
      obj.IntensityScale = level
    Next
  Next
  currentGIBOneLevel = level
  VRBGGI1.opacity = level * 100
  WheelGI1Warm.opacity = level * 30
  WheelGI1.opacity = level * 100
  PinCab_TopperGI1.opacity = level * 100
End Sub

Sub UpdateGIBottom(level)
  Dim coll, obj
  For Each coll in Array(GIBottomLeft,GIBottomRight)
    For Each obj In coll
      obj.IntensityScale = level
    Next
  Next
  currentGIOneLevel = level
  UpdateGILightsBehindDropTargets
  PFShadowGIBottom.opacity = level * 60
  PFShadowGIOff.opacity = 60 - (level * 60)
  If PFShadowGIBottom.opacity = 0 Then
    PFShadowGIBottom.visible = False
  Else
    PFShadowGIBottom.visible = True
  End If
  If PFShadowGIOff.opacity = 0 Then
    PFShadowGIOff.visible = False
  Else
    PFShadowGIOff.visible = True
  End If
End Sub

Sub UpdateGIBackglassTwo(level)
  Dim coll, obj
  For Each coll in Array(GIBackglassTwo)
    For Each obj In coll
      obj.IntensityScale = level
    Next
  Next
  currentGIBTwoLevel = level
  VRBGGI2.opacity = level * 100
  WheelGI2Warm.opacity = level * 30
  WheelGI2.opacity = level * 100
  PinCab_TopperGI2.opacity = level * 100
End Sub

Sub UpdateGITop(level)
  If DebugGI = True Then debug.print "GITop "& level
  Dim coll, obj
  For Each coll in Array(GITopLeft,GITopRight)
    For Each obj In coll
      obj.IntensityScale = level
    Next
  Next
  currentGITwoLevel = level
  PFShadowGITop.opacity = level * 60
  If currentGITwoLevel > 0.1 And currentGIOneLevel < 0.1 Then
    PFShadowGITop.visible = True
  Else
    PFShadowGITop.visible = False
  End If
End Sub

Sub UpdateGIBackglassThree(level)
  Dim coll, obj
  For Each coll in Array(GIBackglassThree)
    For Each obj In coll
      obj.IntensityScale = level
    Next
  Next
  currentGIBThreeLevel = level
  VRBGGI3.opacity = level * 100
  WheelGI3Warm.opacity = level * 30
  WheelGI3.opacity = level * 100
  PinCab_TopperGI3.opacity = level * 100
End Sub

Sub UpdateGILightsBehindDropTargets()
  Dim currentScale
  currentScale = currentGIOneLevel
  gibottom4_sw33_sw34_sw35.IntensityScale   = IIF(Drop33 = 1 And Drop34 = 1 And Drop35 = 1, currentScale, 0)
  gibottom4_sw34_sw35.IntensityScale      = IIF(Drop33 = 0 And Drop34 = 1 And Drop35 = 1, currentScale, 0)
  gibottom4_sw33_sw34.IntensityScale      = IIF(Drop33 = 1 And Drop34 = 1 And Drop35 = 0, currentScale, 0)
  gibottom4_sw33_sw35.IntensityScale      = IIF(Drop33 = 1 And Drop34 = 0 And Drop35 = 1, currentScale, 0)
  gibottom4_sw35.IntensityScale         = IIF(Drop33 = 0 And Drop34 = 0 And Drop35 = 1, currentScale, 0)
  gibottom4_sw34.IntensityScale         = IIF(Drop33 = 0 And Drop34 = 1 And Drop35 = 0, currentScale, 0)
  gibottom4_sw33.IntensityScale         = IIF(Drop33 = 1 And Drop34 = 0 And Drop35 = 0, currentScale, 0)
  gibottom4.IntensityScale          = IIF(Drop33 = 0 And Drop34 = 0 And Drop35 = 0, currentScale, 0)
End Sub

' *********************************************************************
' colors
' *********************************************************************

Dim White, WhiteFull, WhiteBG
WhiteFull = rgb(255,238,179)
White = rgb(255,235,155)
WhiteBG = rgb(255,255,255)

Dim Blue, BlueFull
BlueFull = rgb(75,75,255)
Blue = rgb(75,75,255)

Dim Red, RedFull
RedFull = rgb(255,75,75)
Red = rgb(255,75,75)

Dim Green, GreenFull
GreenFull = rgb(75,255,75)
Green = rgb(75,255,75)

Dim Orange, OrangeFull, OrangeBG
OrangeFull = rgb(255,128,75)
Orange = rgb(255,128,75)
OrangeBG = rgb(255,128,64)

'********************************************
'*  Led part, adapted from JP Salas script  *
'********************************************

Dim rGreen, rRed, rBlue, RGBFactor, RGBStep

Sub UpdateRainbowGI 'rainbow led light color changing
  Dim coll, obj
  RGBFactor =2
  Select Case RGBStep
    Case 0 'Green
      rGreen = rGreen + RGBFactor
      If rGreen > 255 Then
        rGreen = 255
        RGBStep = 1
      End If
    Case 1 'Red
      rRed = rRed - RGBFactor
      If rRed < 0 Then
        rRed = 0
        RGBStep = 2
      End If
    Case 2 'Blue
      rBlue = rBlue + RGBFactor
      If rBlue > 255 Then
        rBlue = 255
        RGBStep = 3
      End If
    Case 3 'Green
      rGreen = rGreen - RGBFactor
      If rGreen < 0 Then
        rGreen = 0
        RGBStep = 4
      End If
    Case 4 'Red
      rRed = rRed + RGBFactor
      If rRed > 255 Then
        rRed = 255
        RGBStep = 5
      End If
    Case 5 'Blue
      rBlue = rBlue - RGBFactor
      If rBlue < 0 Then
        rBlue = 0
        RGBStep = 0
      End If
  End Select
  If currentGITwoLevel > 0.1 Then
    RampCometPlastic2.image = "topPlastic_on"
  Else
    RampCometPlastic2.image = "topPlastic_off"
  End If
  If currentGIOneLevel > 0.1 And currentGITwoLevel < 0.1 Then
    HurricanePlastics1.image = "HurricanePlastics1GIBottom"
    HurricanePlastics2.image = "HurricanePlastics2GIBottom"
    HurricanePlastics3.image = "HurricanePlastics3GIBottom"
    JugglerBox.image = "JugglerBoxGIBottom"
    JugglerFront.image = "JugglerFrontGIBottom"
    RampCometPlastic1.image = "RampCometPlastic1GIBottom"
    RampCometPlastic2.image = "RampCometPlastic2GIBottom"
    RampHurricanePlastic1.image = "RampHurricanePlastic1GIBottom"
    RampHurricanePlastic2.image = "RampHurricanePlastic2GIBottom"
    RampScores.image = "RampscoresGIBottom"
    RampWire.image = "RampWireGIBottom"
    RampWire2.image = "RampWire2GIBottom"
    PFShadowGIBottom.imageA = "PFShadowGIBottom"
    If SideBladeMod = 0 Then
      Pincab_Blades_Left.image = "PinCab_BladesGIBottom"
      Pincab_Blades_Right.image = "PinCab_BladesGIBottom"
    End If
    If SideBladeMod = 1 Then
      Pincab_Blades_Left.image = "PinCab_Blades_Custom1GIBottom"
      Pincab_Blades_Right.image = "PinCab_Blades_Custom1GIBottom"
    End If
    If SideBladeMod = 2 Then
      Pincab_Blades_Left.image = "PinCab_Blades_Custom2GIBottom"
      Pincab_Blades_Right.image = "PinCab_Blades_Custom2GIBottom"
    End If
  ElseIf currentGIOneLevel < 0.1 And currentGITwoLevel > 0.1 Then
    HurricanePlastics1.image = "HurricanePlastics1GITop"
    HurricanePlastics2.image = "HurricanePlastics2GITop"
    HurricanePlastics3.image = "HurricanePlastics3GITop"
    JugglerBox.image = "JugglerBoxGITop"
    JugglerFront.image = "JugglerFrontGITop"
    RampCometPlastic1.image = "RampCometPlastic1GITop"
    RampCometPlastic2.image = "RampCometPlastic2GITop"
    RampHurricanePlastic1.image = "RampHurricanePlastic1GITop"
    RampHurricanePlastic2.image = "RampHurricanePlastic2GITop"
    RampScores.image = "RampscoresGITop"
    RampWire.image = "RampWireGITop"
    RampWire2.image = "RampWire2GITop"
    If SideBladeMod = 0 Then
      Pincab_Blades_Left.image = "PinCab_BladesGITop"
      Pincab_Blades_Right.image = "PinCab_BladesGITop"
    End If
    If SideBladeMod = 1 Then
      Pincab_Blades_Left.image = "PinCab_Blades_Custom1GITop"
      Pincab_Blades_Right.image = "PinCab_Blades_Custom1GITop"
    End If
    If SideBladeMod = 2 Then
      Pincab_Blades_Left.image = "PinCab_Blades_Custom2GITop"
      Pincab_Blades_Right.image = "PinCab_Blades_Custom2GITop"
    End If
  ElseIf currentGIOneLevel > 0.1 And currentGITwoLevel > 0.1 Then
    HurricanePlastics1.image = "HurricanePlastics1GIOn"
    HurricanePlastics2.image = "HurricanePlastics2GIOn"
    HurricanePlastics3.image = "HurricanePlastics3GIOn"
    JugglerBox.image = "JugglerBoxGIOn"
    JugglerFront.image = "JugglerFrontGIOn"
    RampCometPlastic1.image = "RampCometPlastic1GIOn"
    RampCometPlastic2.image = "RampCometPlastic2GIOn"
    RampHurricanePlastic1.image = "RampHurricanePlastic1GIOn"
    RampHurricanePlastic2.image = "RampHurricanePlastic2GIOn"
    RampScores.image = "RampscoresGIOn"
    RampWire.image = "RampWireGIOn"
    RampWire2.image = "RampWire2GIOn"
    PFShadowGIBottom.imageA = "PFShadowGIOn"
    If SideBladeMod = 0 Then
      Pincab_Blades_Left.image = "PinCab_BladesGIOn"
      Pincab_Blades_Right.image = "PinCab_BladesGIOn"
    End If
    If SideBladeMod = 1 Then
      Pincab_Blades_Left.image = "PinCab_Blades_Custom1GIOn"
      Pincab_Blades_Right.image = "PinCab_Blades_Custom1GIOn"
    End If
    If SideBladeMod = 2 Then
      Pincab_Blades_Left.image = "PinCab_Blades_Custom2GIOn"
      Pincab_Blades_Right.image = "PinCab_Blades_Custom2GIOn"
    End If
  ElseIf currentGIOneLevel < 0.1 And currentGITwoLevel < 0.1 Then
    HurricanePlastics1.image = "HurricanePlastics1GIOff"
    HurricanePlastics2.image = "HurricanePlastics2GIOff"
    HurricanePlastics3.image = "HurricanePlastics3GIOff"
    JugglerBox.image = "JugglerBoxGIOff"
    JugglerFront.image = "JugglerFrontGIOff"
    RampCometPlastic1.image = "RampCometPlastic1GIOff"
    RampCometPlastic2.image = "RampCometPlastic2GIOff"
    RampHurricanePlastic1.image = "RampHurricanePlastic1GIOff"
    RampHurricanePlastic2.image = "RampHurricanePlastic2GIOff"
    RampScores.image = "RampscoresGIOff"
    RampWire.image = "RampWireGIOff"
    RampWire2.image = "RampWire2GIOff"
    If SideBladeMod = 0 Then
      Pincab_Blades_Left.image = "PinCab_BladesGIOff"
      Pincab_Blades_Right.image = "PinCab_BladesGIOff"
    End If
    If SideBladeMod = 1 Then
      Pincab_Blades_Left.image = "PinCab_Blades_Custom1GIOff"
      Pincab_Blades_Right.image = "PinCab_Blades_Custom1GIOff"
    End If
    If SideBladeMod = 2 Then
      Pincab_Blades_Left.image = "PinCab_Blades_Custom2GIOff"
      Pincab_Blades_Right.image = "PinCab_Blades_Custom2GIOff"
    End If
  End If
  If currentGIBOneLevel < 0.1 And currentGIBTwoLevel < 0.1 And currentGIBThreeLevel < 0.1 Then
    Mystery_Wheel_Reel.BlendDisableLighting = 0
  ElseIf currentGIBOneLevel > 0.1 And currentGIBTwoLevel < 0.1 And currentGIBThreeLevel < 0.1 Then
    Mystery_Wheel_Reel.BlendDisableLighting = currentGIBOneLevel / 2
  ElseIf currentGIBOneLevel < 0.1 And currentGIBTwoLevel > 0.1 And currentGIBThreeLevel < 0.1 Then
    Mystery_Wheel_Reel.BlendDisableLighting = currentGIBTwoLevel / 2
  ElseIf currentGIBOneLevel < 0.1 And currentGIBTwoLevel < 0.1 And currentGIBThreeLevel > 0.1 Then
    Mystery_Wheel_Reel.BlendDisableLighting = currentGIBThreeLevel / 2
  Else
    Mystery_Wheel_Reel.BlendDisableLighting = 0.5
  End If
  If currentGITwoLevel > 0.1 Then
    If GIColorMod = 1 Or GIColorMod = 5 Or GIColorMod = 6 Then
      Pincab_Blades_Left.Color  = Red
      PFShadowGIBottom.Color  = Red
      PFShadowGITop.Color = Red
    ElseIf GIColorMod = 2 Then
      Pincab_Blades_Left.Color  = Blue
      PFShadowGIBottom.Color  = Blue
      PFShadowGITop.Color = Blue
    ElseIf GIColorMod = 3 Then
      Pincab_Blades_Left.Color  = Green
      PFShadowGIBottom.Color  = Green
      PFShadowGITop.Color = Green
    ElseIf GIColorMod = 4 Then
      Pincab_Blades_Left.Color = Orange
      Pincab_Blades_Right.Color = Orange
      PFShadowGIBottom.Color  = Orange
      PFShadowGITop.Color = Orange
    ElseIf GIColorMod = 7 Then
      Pincab_Blades_Left.color = RGB(rRed, rGreen, rBlue)
      Pincab_Blades_Right.color = RGB(rRed, rGreen, rBlue)
      PFShadowGIBottom.Color  = RGB(rRed, rGreen, rBlue)
      PFShadowGITop.Color = RGB(rRed, rGreen, rBlue)
    Else
      Pincab_Blades_Left.Color = White
      Pincab_Blades_Right.Color = White
      PFShadowGIBottom.Color  = White
      PFShadowGITop.Color = White
    End If
    If GIColorMod = 2 Or GIColorMod = 5 Then
      Pincab_Blades_Right.Color = Blue
    ElseIf GIColorMod = 1 Then
      Pincab_Blades_Right.Color = Red
    ElseIf GIColorMod = 3 Or GIColorMod = 6 Then
      Pincab_Blades_Right.Color = Green
    End If
  Else
    Pincab_Blades_Left.Color = White
    Pincab_Blades_Right.Color = White
  End If
  For Each coll in Array(BackglassGIreflections, VRTopperGI)
    For Each obj In coll
      If BackglassColorMod = 1 Then
        obj.color = WhiteBG
      Else
        obj.color = OrangeBG
      End If
    Next
  Next
  For Each coll in Array(GIBottomLeft,GITopLeft)
    For Each obj In coll
      If GIColorMod = 7 Then
        obj.color = RGB(rRed, rGreen, rBlue)
        obj.colorfull = RGB(rRed, rGreen, rBlue)
      End If
    Next
  Next
  For Each coll in Array(GIBottomRight,GITopRight)
    For Each obj In coll
      If GIColorMod = 7 Then
        obj.color = RGB(rRed, rGreen, rBlue)
        obj.colorfull = RGB(rRed, rGreen, rBlue)
      End If
    Next
  Next
End Sub

' *********************************************************************
' class for some math stuff and some more general methods
' *********************************************************************

Class clsMyMath
  Public Property Get Pi()
    Pi = CSng(4*Atn(1))
  End Property

  Public Function dCos(degrees)
    dCos = cos(degrees * Pi/180)
    if ABS(dCos) < 0.000001 Then dCos = 0
    if ABS(dCos) > 0.999999 Then dCos = 1 * sgn(dCos)
  End Function

  Public Function dSin(degrees)
    dSin = sin(degrees * Pi/180)
    if ABS(dSin) < 0.000001 Then dSin = 0
    if ABS(dSin) > 0.999999 Then dSin = 1 * sgn(dSin)
  End Function

  Public Function dAtn(x)
    dAtn = Atn(x) * 180 / Pi
  End Function

  Public Function dAtn2(X, Y)
    If X > 0 Then
      dAtn2 = dAtn(Y / X)
    ElseIf X < 0 Then
      dAtn2 = dAtn(Y / X) + 180 * Sgn(Y)
      If Y = 0 Then dAtn2 = dAtn2 + 180
      If Y < 0 Then dAtn2 = dAtn2 + 360
    Else
      dAtn2 = 90 * Sgn(Y)
    End If
    dAtn2 = dAtn2 + 90
  End Function

  ' Ferris wheel stuff
  Private pWheelRadius
  Public Property Let WheelRadius(wRadius)
    pWheelRadius = wRadius
  End Property

  Private pWheelAngle
  Public Property Let WheelAngle(wAngle)
    pWheelAngle = wAngle
  End Property

  Public Sub BallPos(wheelRot, wheelX, wheelY, wheelZ)
    Dim pSinus
    pSinus = myMath.dSin(wheelRot) * pWheelRadius
    wheelX = pSinus * myMath.dSin(pWheelAngle)
    wheelY = pSinus * myMath.dCos(pWheelAngle) * -1
    wheelZ = myMath.dCos(wheelRot) * pWheelRadius * -1
  End Sub
End Class

Function Minimum(val1, val2)
  Minimum = IIF(val1<val2,val2,val1)
End Function

Function IIF(bool, obj1, obj2)
  If bool Then
    IIF = obj1
  Else
    IIF = obj2
  End If
End Function

'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light.
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(3)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii / 1000 + 1.04
    objBallShadow(iii).visible = 0
  Next
End Sub

Sub BSUpdate
  Dim s
  'The Magic happens now
  For s = 0 To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    '** If on main pf
    If gBOT(s).Z > 20 and gBOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25
    '** Under pf, no shadow
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub

'******************************************************
' ZPHY:  GENERAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners         https://youtu.be/AXX3aen06FM?si=Xqd-rcaqTlgEd_wx
' Adding nFozzy roth physics : pt2 flipper physics          https://youtu.be/VSBFuK2RCPE?si=i8ne8Ao2co8rt7fy
' Adding nFozzy roth physics : pt3 other elements           https://youtu.be/JN8HEJapCvs?si=hvgMOk-ej1BEYjJv
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25)
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 12-15    |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 3-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |

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

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'
'        x.AddPt "Polarity", 0, 0, 0
'        x.AddPt "Polarity", 1, 0.05, - 2.7
'        x.AddPt "Polarity", 2, 0.16, - 2.7
'        x.AddPt "Polarity", 3, 0.22, - 0
'        x.AddPt "Polarity", 4, 0.25, - 0
'        x.AddPt "Polarity", 5, 0.3, - 1
'        x.AddPt "Polarity", 6, 0.4, - 2
'        x.AddPt "Polarity", 7, 0.5, - 2.7
'        x.AddPt "Polarity", 8, 0.65, - 1.8
'        x.AddPt "Polarity", 9, 0.75, - 0.5
'        x.AddPt "Polarity", 10, 0.81, - 0.5
'        x.AddPt "Polarity", 11, 0.88, 0
'        x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 3.7
'   x.AddPt "Polarity", 2, 0.16, - 3.7
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 3.7
'   x.AddPt "Polarity", 8, 0.65, - 2.3
'   x.AddPt "Polarity", 9, 0.75, - 1.5
'   x.AddPt "Polarity", 10, 0.81, - 1
'   x.AddPt "Polarity", 11, 0.88, 0
'   x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
''*******************************************
''  Late 80's early 90's
'
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

'*******************************************
' Early 90's and after
'
'Sub InitPolarity()
' Dim x, a
' a = Array(LF, RF)
' For Each x In a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5.5
'   x.AddPt "Polarity", 2, 0.16, - 5.5
'   x.AddPt "Polarity", 3, 0.20, - 0.75
'   x.AddPt "Polarity", 4, 0.25, - 1.25
'   x.AddPt "Polarity", 5, 0.3, - 1.75
'   x.AddPt "Polarity", 6, 0.4, - 3.5
'   x.AddPt "Polarity", 7, 0.5, - 5.25
'   x.AddPt "Polarity", 8, 0.7, - 4.0
'   x.AddPt "Polarity", 9, 0.75, - 3.5
'   x.AddPt "Polarity", 10, 0.8, - 3.0
'   x.AddPt "Polarity", 11, 0.85, - 2.5
'   x.AddPt "Polarity", 12, 0.9, - 2.0
'   x.AddPt "Polarity", 13, 0.95, - 1.5
'   x.AddPt "Polarity", 14, 1, - 1.0
'   x.AddPt "Polarity", 15, 1.05, -0.5
'   x.AddPt "Polarity", 16, 1.1, 0
'   x.AddPt "Polarity", 17, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.23, 0.85
'   x.AddPt "Velocity", 2, 0.27, 1
'   x.AddPt "Velocity", 3, 0.3, 1
'   x.AddPt "Velocity", 4, 0.35, 1
'   x.AddPt "Velocity", 5, 0.6, 1 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

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
' Dim BOT
' BOT = GetBalls

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
    Dim b', BOT
'   BOT = GetBalls

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
  AddSlingsPt 0, 0.00, - 3
  AddSlingsPt 1, 0.30, - 5
  AddSlingsPt 2, 0.40,  -30
  AddSlingsPt 3, 0.60,  30
  AddSlingsPt 4, 0.70,  5
  AddSlingsPt 5, 1.00,  3
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
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' The Stand Up and Drop Target solutions improve the physics for targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full target animation, switch handling and deflection on hit. For drop targets there is also a slight lift when
' the drop targets raise, bricking, and popping the ball up if it's over the drop target when it raises.
'
' Add a Timers named DTAnim and STAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
' DTAnim.interval = 10
' DTAnim.enabled = True

' Sub DTAnim_Timer
'   DoDTAnim
' DoSTAnim
' End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.
'
' For each stand up target we'll use a vp target, a laid back collidable primitive, and one primitive for visuals and animation.
' The visual primitive should should have it's pivot point centered on the x and y axis and the z should be at or just below the playfield.
' The target should animate backwards using transy.
'
' To create visual target primitives that work with the stand up and drop target code, follow the below instructions:
' (Other methods will work as well, but this is easy for even non-blender users to do)
' 1) Open a new blank table. Delete everything off the table in editor.
' 2) Copy and paste the VP target from your table into this blank table.
' 3) Place the target at x = 0, y = 0  (upper left hand corner) with an orientation of 0 (target facing the front of the table)
' 4) Under the file menu, select Export "OBJ Mesh"
' 5) Go to "https://threejs.org/editor/". Here you can modify the exported obj file. When you export, it exports your target and also
'    the playfield mesh. You need to delete the playfield mesh here. Under the file menu, chose import, and select the obj you exported
'    from VPX. In the right hand panel, find the Playfield object and click on it and delete. Then use the file menu to Export OBJ.
' 6) In VPX, you can add a primitive and use "Import Mesh" to import the exported obj from the previous step. X,Y,Z scale should be 1.
'    The primitive will use the same target texture as the VP target object.
'
' * Note, each target must have a unique switch number. If they share a same number, add 100 to additional target with that number.
' For example, three targets with switch 32 would use 32, 132, 232 for their switch numbers.
' The 100 and 200 will be removed when setting the switch value for the target.

'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

Class DropTarget
  Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class

'Define a variable for each drop target
Dim DT33, DT34, DT35

Set DT33 = (new DropTarget)(sw33, sw33a, sw33p, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34a, sw34p, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, sw35p, 35, 0, false)

Dim DTArray
DTArray = Array(DT33, DT34, DT35)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.1 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

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
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 1 'If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0 'updated by rothbauerw to account for edge case
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
      controller.Switch(Switchid mod 100) = 1
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
'     Dim BOT
'     BOT = GetBalls

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
    controller.Switch(Switchid mod 100) = 0
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
'****  END DROP TARGETS
'******************************************************

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
Dim ST42, ST43, ST44, ST45, ST55

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

Set ST42 = (new StandupTarget)(sw42, sw42p, 42, 0)
Set ST43 = (new StandupTarget)(sw43, sw43p, 43, 0)
Set ST44 = (new StandupTarget)(sw44, sw44p, 44, 0)
Set ST45 = (new StandupTarget)(sw45, sw45p, 45, 0)
Set ST55 = (new StandupTarget)(sw55, sw55p, 55, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST42, ST43, ST44, ST45, ST55)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.1        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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
'***  END STAND-UP TARGETS
'******************************************************

'******************************************************
' ZBRL: BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

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
  Dim b', BOT
  'BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound(Cartridge_Ball_Roll & "_Ball_Roll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound (Cartridge_Ball_Roll & "_Ball_Roll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound(Cartridge_Ball_Roll & "_Ball_Roll_" & b)
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
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
dim RampBalls(5,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(5)

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)  'Add ball
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  dim x : for x = 1 to uBound(RampBalls)  'Check, don't add balls twice
    if RampBalls(x, 1) = input.id then
      if Not IsEmpty(RampBalls(x,1) ) then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 to uBound(RampBalls)
    if IsEmpty(RampBalls(x, 1)) then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      exit Sub
    End If
    if x = uBound(RampBalls) then   'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
      RampBalls(0, 0) & vbnewline & _
      Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
      Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
      Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
      Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
      Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
      " "
    End If
  next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)   'Remove ball
  'Debug.Print "In WRemoveBall() + Remove ball from loop array"
  dim ballcount : ballcount = 0
  dim x : for x = 1 to Ubound(RampBalls)
    if ID = RampBalls(x, 1) then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
  next
  if BallCount = 0 then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end if
  next
  if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
  "1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
  "2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
  "3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
  "4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
  "5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
  "6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
  " "
End Sub


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

'/////////////////////////////  PLASTIC RAMPS FLAPS - SOUNDS  ////////////////////////////
Sub RandomSoundRampFlapUp()
' debug.print "flap up"
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_1"), FlapSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_2"), FlapSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_3"), FlapSoundLevel
  End Select
End Sub

Sub RandomSoundRampFlapDown()
' debug.print "flap down"
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_1"), FlapSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_2"), FlapSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_3"), FlapSoundLevel
  End Select
End Sub

'///////////////////////-----Ramp Flaps-----///////////////////////

Dim FlapSoundLevel
FlapSoundLevel = 0.8                          'volume level; range [0, 1]

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other SYS11 systems, but the most are tailored specifically for the Space Station table.

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'    GatesOneWay (one way wire gates)
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

' Updated 2025 by RobbyKingPin. This edited version contains almost exactly the same structure and subroutines as the original Fleep Sound codes but with the addition of the cartrigdes used inside the updated Fleep from 2022.

'/////////////////////////////////  CARTRIDGES  /////////////////////////////////
'
'//  Specify which mechanical sound cartridge to use for each group of elements.
'//  Mechanical sounds naming convention: <CARTRIDGE>_<Soundset_Name>
'//
'//  Cartridge name is composed using the following convention:
'//  <TABLE MANUFACTURER ABBREVIATION>_<TABLE NAME ABBREVIATION>_<SOUNDSET REVISION NUMBER>
'//
'//  General Mechanical Sounds Cartridges:

Const Cartridge_Bumpers         = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Slingshots        = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Flippers        = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Knocker         = "WS_WHD_REV02" 'Williams Whirlwind Cartridge REV02
Const Cartridge_Relays          = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Trough          = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Rollovers       = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Targets         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Gates         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Metal_Hits        = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Rubber_Hits       = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Cabinet_Sounds      = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Apron         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Ball_Roll       = "BY_TOM_REV01" 'Bally Theatre of Magic Cartridge REV01
Const Cartridge_BallBallCollision   = "BY_WDT_REV01" 'Bally WHO Dunnit Cartridge REV01
Const Cartridge_Ball_Drop_Bump      = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01

'////////////////////////////  SOUND SOURCE CREDITS  ////////////////////////////
'//  Special thanks go to the following contributors who have provided audio
'//  footage recordings:
'//
'//  Williams Whirlwind - Blackmoor, wrd1972
'//  Williams Diner - Nick Rusis
'//  Spooky Total Nuclear Annihilation - WildDogArcade, Ed and Gary
'//  Bally Theatre of Magic - CalleV, nickbuol
'//  Bally WHO Dunnit - Amazaley1
'//  Williams Pinbot - major_drain_pinball

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
'//  FLIPPER COILS:
'//  Flippers in this table: Lower Left Flipper, Lower Right Flipper, Upper Right Fliiper
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel
Dim FlipperLeftLowerHitParm, FlipperRightLowerHitParm

'//  Flipper Up Attacks initialize during playsound subs
Dim FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel

FlipperUpSoundLevel = 1
FlipperDownSoundLevel = 0.65
FlipperUpAttackMinimumSoundLevel = 0.010
FlipperUpAttackMaximumSoundLevel = 0.435

'//  Flipper Hit Param initialize with FlipperUpSoundLevel
'//  and dynamically modified calculated by ball flipper collision
FlipperLeftLowerHitParm = FlipperUpSoundLevel
FlipperRightLowerHitParm = FlipperUpSoundLevel
Dim FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

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

'////////////////////////////////  SOUND HELPERS  ///////////////////////////////
Dim SoundOn : SoundOn = 1
Dim SoundOff : SoundOff = 0
Dim Up : Up = 0
Dim Down : Down = 1
Dim RampUp : RampUp = 1
Dim RampDown : RampDown = 0
Dim RampDownSlow : RampDownSlow = 1
Dim RampDownFast : RampDownFast = 2
Dim CircuitA : CircuitA = 0
Dim CircuitC : CircuitC = 1

'//  Helper for Main (Lower) flippers dampened stroke
Dim BallNearLF : BallNearLF = 0
Dim BallNearRF : BallNearRF = 0

Sub TriggerBallNearLF_Hit()
  'Debug.Print "BallNearLF = 1"
  BallNearLF = 1
End Sub

Sub TriggerBallNearLF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearLF = 0
End Sub

Sub TriggerBallNearRF_Hit()
  'Debug.Print "BallNearRF = 1"
  BallNearRF = 1
End Sub

Sub TriggerBallNearRF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearRF = 0
End Sub

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
  PlaySound (Cartridge_Cabinet_Sounds & "_Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerPullStop()
  StopSound Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Ball_" & Int(Rnd*3)+1), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Empty"), PlungerReleaseSoundLevel, Plunger
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Knocker & "_Knocker_Coil",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic (Cartridge_Trough & "_Outhole_Drain_Hit_" & Int(Rnd*4)+1), DrainSoundLevel, drainswitch
End Sub

Sub RandomSoundDrainKicker(drainswitch)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Outhole_Kicker_" & Int(Rnd*4)+1,DOFContactors), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Shooter_Feeder_" & Int(Rnd*6)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Left_" & Int(Rnd*26)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Right_" & Int(Rnd*25)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Left_" & Int(Rnd*22)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Up_" & Int(Rnd*25)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Low_" & Int(Rnd*28)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'///////////////////////  FLIPPER BATS SOUND SUBROUTINES  ///////////////////////
'//////////////////////  FLIPPER BATS SOLENOID CORE SOUND  //////////////////////

Sub RandomSoundFlipperLowerLeftUpFullStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & RndInt(1,10),DOFFlippers), FlipperLeftLowerHitParm, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftUpDampenedStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & RndInt(1,23),DOFFlippers), FlipperLeftLowerHitParm * 1.2, Flipper
End Sub

Sub RandomSoundFlipperLowerRightUpFullStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & RndInt(1,11),DOFFlippers), FlipperRightLowerHitParm, Flipper
End Sub

Sub RandomSoundFlipperLowerRightUpDampenedStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & RndInt(1,23),DOFFlippers), FlipperLeftLowerHitParm * 1.2, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftReflip(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Reflip_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerRightReflip(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Reflip_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftDown(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & RndInt(1,10),DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerRightDown(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & RndInt(1,11),DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundLowerLeftQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, LeftFlipper
End Sub

Sub RandomSoundLowerRightQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, RightFlipper
End Sub

Sub StopAnyFlipperLowerLeftUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperLowerRightUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 11
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperLowerLeftDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & anyFullDownSound)
  Next
End Sub

Sub StopAnyFlipperLowerRightDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & anyFullDownSound)
  Next
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Flipper_Hit_" & Int(Rnd*7)+1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_All_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

'///////////////////////////  OUTLANE ROLLOVER SOUNDS  //////////////////////////

Sub RandomSoundOutlaneRollover()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_Outlane_" & Int(Rnd*4)+1), RolloverSoundLevel
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
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_10"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_" & Int(Rnd*8)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor * 0.05
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
  Select Case Int(Rnd*20)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_1"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_2"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_3"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_4"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_5"), Vol(ActiveBall) * 0.02 * MetalImpactSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_6"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_7"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_8"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_9"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_10"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 11 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_11"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 12 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_12"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 13 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_13"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 14 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_14"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 15 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_15"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 16 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_16"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 17 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_17"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 18 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_18"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 19 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_19"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 20 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_20"), Vol(ActiveBall) * MetalImpactSoundFactor
  End Select
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
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Soft_" & Int(Rnd*4)+1), BottomArchBallGuideSoundFactor
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Hard_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 3
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
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End if
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_5",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_6",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_7",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_8",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
End Sub

Sub RandomSoundTargetHitWeak()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_1",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_2",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_3",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_4",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
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
  PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_1"), GateSoundLevel * 0.5, ActiveBall
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_2"), GateSoundLevel * 0.005, ActiveBall
End Sub

Sub SoundOneWayGate() 'sound of blocked gate
  PlaySoundAtLevelStatic (Cartridge_Gates & "_Oneway_Ball_Gate_" & Int(Rnd*3)+1), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

Sub GatesOneWay_Hit(idx)
  SoundOneWayGate
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
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_1")
    Case 2
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_2")
    Case 3
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_3")
    Case 4
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_4")
    Case 5
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_5")
    Case 6
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_6")
    Case 7
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_7")
  End Select

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Targets & "_Drop_Target_1Bank_Reset_Up_" & Int(Rnd*8)+1,DOFContactors), 1, obj
End Sub

Sub RandomSoundDropTargetBankReset(obj)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Targets & "_Drop_Target_3Bank_Reset_Up_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_1Bank_Release_Down_" & Int(Rnd*2)+1), 200, obj
End Sub

Sub SoundDropTargetBankDrop(obj)
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_3Bank_Release_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.015  'volume level; range [0, 1];
Const RelayGISoundLevel = 0.45      'volume level; range [0, 1];
Const RelayBackboxSoundLevel = 0.45 'volume level; range [0, 1];
Const RelayACSelectSoundLevel = 0.3 'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_GIBackbox_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_On"), RelayBackboxSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_Off"), RelayBackboxSoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlashSoundLevel, obj
  End Select
End Sub

'//////////////////////////  SOLENOID A/C SELECT RELAY  /////////////////////////
Sub Sound_Solenoid_ACSelect_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_A"), RelayACSelectSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_C"), RelayACSelectSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'   ZFLD:  FLUPPER DOMES
'******************************************************
' Based on FlupperDomes2.2

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials domebasemat, Flashermaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "dome" and "ronddome" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names
' Copy a set of 4 objects flasherbase, flasherlit, flasherlight and flasherflash from layer 7 to your table
' If you duplicate the four objects for a new flasher dome, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherbloom flashers from layer 10 to your table. you will need to make one per flasher dome that you plan to make
' Select the correct flasherbloom texture for each flasherbloom flasher, per flasher dome
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasher in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitFlasher 1, "green"
'
' Color Options: "blue", "green", "red", "purple", "yellow", "white", and "orange"

' You can use the RotateFlasher call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasher sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasher 1, 180    'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLevel(1) = 1 : FlasherFlash1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasher below)
' - Color of the objects are set in the script, not on the primitive itself
' - Screws are optional to copy and position manually
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flashermaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlit and flasherflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlightintensity and/or flasherflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

'------ Main Dome Code ---------'

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object      ***
Set TableRef = Table1      ' *** change this, if your table has another name           ***
FlasherLightIntensity = 0.1  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)     ***
FlasherFlareIntensity = 0.3  ' *** lower this, if the flares are too bright (i.e. 0.1)       ***
FlasherBloomIntensity = 0.2  ' *** lower this, if the blooms are too bright (i.e. 0.1)       ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "red"
InitFlasher 2, "orange"
InitFlasher 3, "blue"
InitFlasher 4, "white"
InitFlasher 5, "red"
InitFlasher 6, "red"
InitFlasher 7, "red"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateFlasher 1,17
'   RotateFlasher 2,0
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

''------ Use this for PWM following domes ---------'

Sub ModFlashFlasher(nr, aValue)
  objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * aValue
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * aValue
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * aValue
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * aValue
  objlit(nr).BlendDisableLighting = 10 * aValue
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,aValue,RGB(255,255,255),0,0,False,True,0,0,0,0
End Sub

' example script for rom based tables (modulated):

Sub SolModFlashPWM17(level)       'Flasher Solonoid Name
  If DebugFlashers = True Then debug.print "SolModFlashPWM17 "& level
  ModFlashFlasher 1,level     'Flasher Number assigned in flupper script
  ModFlashFlasher 2,level     'Flasher Number assigned in flupper script
  fLight17.state = level
  WheelS17.opacity = level * 60
End Sub

Sub SolModFlashPWM18(level)       'Flasher Solonoid Name
  If DebugFlashers = True Then debug.print "SolModFlashPWM18 "& level
  ModFlashFlasher 3,level     'Flasher Number assigned in flupper script
  ModFlashFlasher 4,level     'Flasher Number assigned in flupper script
  fLight18.state = level
  WheelS18.opacity = level * 60
End Sub

Sub SolModFlashPWM23(level)       'Flasher Solonoid Name
  If DebugFlashers = True Then debug.print "SolModFlashPWM23 "& level
  ModFlashFlasher 5,level     'Flasher Number assigned in flupper script
  ModFlashFlasher 6,level     'Flasher Number assigned in flupper script
  fLight23.state = level
  WheelS23.opacity = level * 60
End Sub

Sub SolModFlashPWM26(level)       'Flasher Solonoid Name
  If DebugFlashers = True Then debug.print "SolModFlashPWM26 "& level
  ModFlashFlasher 7,level     'Flasher Number assigned in flupper script
  fLight26.state = level
  WheelS26.opacity = level * 60
End Sub

'******************************************************
'******  END FLUPPER DOMES
'******************************************************

'******************************************************
'   ZFLF:  FLUPPER FLASHERLAMPS
'******************************************************

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials flasherlampbasemat, Flasherlampmaterial0 - 20, Flasherbulbmaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "flasherbulb" and import them into your table with the same
' Export all textures (images) starting with the name "flasherlamp" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names (only needed if you don't have Flupper Domes)
' Copy a set of 7 objects flasherlampbase, flasherlamplit, flasherbulbbase, flasherbulblit, flasherlampfilament, flasherlamplight and flasherlampflash from layer 7 to your table
' If you duplicate the four objects for a new flasherlamp, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherlampbloom flashers from layer 10 to your table. you will need to make one per flasher lamp that you plan to make
' Select the correct flasherbloom texture for each flasherlampbloom flasher, per flasher lamp
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasherLamp in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitLampFlasher 1
'
' You can use the RotateFlasherLamp call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasherLamp sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasherLamp 1, 180    'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLampTargetLevel(1) = 1 : FlasherFlashLamp1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLampLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasherLamp below)
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flasherlampmaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherlampbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherlampflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlamplit and flasherlampflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlamplightintensity and/or flasherlampflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

Dim TestLampFlashers, TableLampRef, FlasherLampLightIntensity, FlasherLampFlareIntensity, FlasherLampBloomIntensity, FlasherLampOffBrightness, FlasherBulbOffBrightness
' *********************************************************************
TestLampFlashers = 0        ' *** set this to 1 to check position of flasher object      ***
Set TableLampRef = Table1      ' *** change this, if your table has another name           ***
FlasherLampLightIntensity = 0.05   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)     ***
FlasherLampFlareIntensity = 0.05   ' *** lower this, if the flares are too bright (i.e. 0.1)       ***
FlasherLampBloomIntensity = 0.2  ' *** lower this, if the blooms are too bright (i.e. 0.1)       ***
FlasherLampOffBrightness = 0.5    ' *** brightness of the flasher lamp when switched off (range 0-2)  ***
FlasherBulbOffBrightness = 0.5    ' *** brightness of the flasher bulb when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLampLevel(20), objlampbase(20), objlamplit(20), objbulbbase(20), objbulblit(20), objlampflasher(20), objlampbloom(20), objlamplight(20), ObjLampTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

InitLampFlasher 1
InitLampFlasher 2
InitLampFlasher 3
InitLampFlasher 4
InitLampFlasher 5

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateLampFlasher 1,17
'   RotateLampFlasher 2,0
'   RotateLampFlasher 3,90
'   RotateLampFlasher 4,90

Sub InitLampFlasher(nr)
  ' store all objects in an array for use in FlashLampFlasher subroutine
  Set objlampbase(nr) = Eval("Flasherlampbase" & nr)
  Set objlamplit(nr) = Eval("Flasherlamplit" & nr)
  Set objbulbbase(nr) = Eval("Flasherbulbbase" & nr)
  Set objbulblit(nr) = Eval("Flasherbulblit" & nr)
  Set objlampflasher(nr) = Eval("Flasherlampflash" & nr)
  Set objlamplight(nr) = Eval("Flasherlamplight" & nr)
  Set objlampbloom(nr) = Eval("Flasherlampbloom" & nr)

  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objlampbase(nr).RotY = 0 Then
    objlampbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objlampbase(nr).x) / (objlampbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
  End If

  If objbulbbase(nr).RotY = 0 Then
    objbulbbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbulbbase(nr).x) / (objbulbbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
    objlampflasher(nr).RotZ = objbulbbase(nr).ObjRotZ
    objlampflasher(nr).height = objbulbbase(nr).z + 40
  End If

  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlamplight(nr).IntensityScale = 0
  objlamplit(nr).visible = 0
  'objlamplit(nr).material = "Flasherlampmaterial" & nr
  objlamplit(nr).RotX = objlampbase(nr).RotX
  objlamplit(nr).RotY = objlampbase(nr).RotY
  objlamplit(nr).RotZ = objlampbase(nr).RotZ
  objlamplit(nr).ObjRotX = objlampbase(nr).ObjRotX
  objlamplit(nr).ObjRotY = objlampbase(nr).ObjRotY
  objlamplit(nr).ObjRotZ = objlampbase(nr).ObjRotZ
  objlamplit(nr).x = objlampbase(nr).x
  objlamplit(nr).y = objlampbase(nr).y
  objlamplit(nr).z = objlampbase(nr).z
  objlampbase(nr).BlendDisableLighting = FlasherLampOffBrightness

  objbulblit(nr).visible = 0
  objbulblit(nr).material = "Flasherbulbmaterial" & nr
  objbulblit(nr).RotX = objbulbbase(nr).RotX
  objbulblit(nr).RotY = objbulbbase(nr).RotY
  objbulblit(nr).RotZ = objbulbbase(nr).RotZ
  objbulblit(nr).ObjRotX = objbulbbase(nr).ObjRotX
  objbulblit(nr).ObjRotY = objbulbbase(nr).ObjRotY
  objbulblit(nr).ObjRotZ = objbulbbase(nr).ObjRotZ
  objbulblit(nr).x = objbulbbase(nr).x
  objbulblit(nr).y = objbulbbase(nr).y
  objbulblit(nr).z = objbulbbase(nr).z
  objbulbbase(nr).BlendDisableLighting = FlasherBulbOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objlampbase(nr).roty > 135 Then
    objlampflasher(nr).y = objlampbase(nr).y + 50
    objlampflasher(nr).height = objlampbase(nr).z + 20
  Else
    objlampflasher(nr).y = objlampbase(nr).y + 20
    objlampflasher(nr).height = objlampbase(nr).z + 50
  End If
  objlampflasher(nr).x = objlampbase(nr).x

  'rothbauerw
  'Adjust the position of the light object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  objlamplight(nr).x = objlampbase(nr).x
  objlamplight(nr).y = objlampbase(nr).y
  objlamplight(nr).bulbhaloheight = objlampbase(nr).z - 10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  Dim xthird, ythird
  xthird = tablewidth / 3
  ythird = tableheight / 3
  If objlampbase(nr).x >= xthird And objlampbase(nr).x <= xthird * 2 Then
    objlampbloom(nr).imageA = "flasherbloomCenter"
    objlampbloom(nr).imageB = "flasherbloomCenter"
  ElseIf objlampbase(nr).x < xthird And objlampbase(nr).y < ythird Then
    objlampbloom(nr).imageA = "flasherbloomUpperLeft"
    objlampbloom(nr).imageB = "flasherbloomUpperLeft"
  ElseIf  objlampbase(nr).x > xthird * 2 And objlampbase(nr).y < ythird Then
    objlampbloom(nr).imageA = "flasherbloomUpperRight"
    objlampbloom(nr).imageB = "flasherbloomUpperRight"
  ElseIf objlampbase(nr).x < xthird And objlampbase(nr).y < ythird * 2 Then
    objlampbloom(nr).imageA = "flasherbloomCenterLeft"
    objlampbloom(nr).imageB = "flasherbloomCenterLeft"
  ElseIf  objlampbase(nr).x > xthird * 2 And objlampbase(nr).y < ythird * 2 Then
    objlampbloom(nr).imageA = "flasherbloomCenterRight"
    objlampbloom(nr).imageB = "flasherbloomCenterRight"
  ElseIf objlampbase(nr).x < xthird And objlampbase(nr).y < ythird * 3 Then
    objlampbloom(nr).imageA = "flasherbloomLowerLeft"
    objlampbloom(nr).imageB = "flasherbloomLowerLeft"
  ElseIf  objlampbase(nr).x > xthird * 2 And objlampbase(nr).y < ythird * 3 Then
    objlampbloom(nr).imageA = "flasherbloomLowerRight"
    objlampbloom(nr).imageB = "flasherbloomLowerRight"
  End If

  ' set the texture and color of all objects
  'objlampbase(nr).image = "lamp2base"
  'objlamplit(nr).image = "lamp2lit"
  objbulbbase(nr).image = "flasherbulb"
  objbulblit(nr).image = "flasherbulblit"
  If TestLampFlashers = 0 Then
    objlampflasher(nr).imageA = "lampflashwhite"
    objlampflasher(nr).visible = 0
  End If

  objlamplight(nr).color = RGB(255,240,150)
  objlampflasher(nr).color = RGB(100,86,59)
  objlampbloom(nr).color = RGB(255,240,150)

  objlamplight(nr).colorfull = objlamplight(nr).color
  If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
    objlampflasher(nr).height = objlampflasher(nr).height - 20 * objlampflasher(nr).y / tableheight
    objlampflasher(nr).y = objlampflasher(nr).y + 10
  End If
End Sub

Sub RotateLampFlasher(nr, angle)
  angle = ((angle + 360 - objlampbase(nr).ObjRotZ) Mod 180) / 30
  objlampbase(nr).showframe(angle)
  objlamplit(nr).showframe(angle)
  objbulbbase(nr).showframe(angle)
  objbulblit(nr).showframe(angle)
End Sub

Sub ModFlashLampFlasher(nr, aValue)
    if aValue > 0 then
        objlampflasher(nr).visible = 1 : objlampbloom(nr).visible = 1 : objlamplit(nr).visible = 1 : objbulblit(nr).visible = 1
    else
        objlampflasher(nr).visible = 0 : objlampbloom(nr).visible = 0 : objlamplit(nr).visible = 0 : objbulblit(nr).visible = 0
    end if
    objlampflasher(nr).opacity = 1000 *  FlasherLampFlareIntensity * aValue
    objlampbloom(nr).opacity = 100 *  FlasherLampBloomIntensity * aValue
    objlamplight(nr).IntensityScale = 0.5 * FlasherLampLightIntensity * aValue
    objlampbase(nr).BlendDisableLighting =  FlasherLampOffBrightness + 10 * aValue
    objlamplit(nr).BlendDisableLighting = 10 * aValue
    'UpdateMaterial "Flasherlampmaterial" & nr, 0,0,0,0,0,0,Round(aValue,1),RGB(255,255,255),0,0,False,True,0,0,0,0
  objbulbbase(nr).BlendDisableLighting =  FlasherLampOffBrightness + 10 * aValue
    objbulblit(nr).BlendDisableLighting = 10 * aValue
    UpdateMaterial "Flasherbulbmaterial" & nr, 0,0,0,0,0,0,Round(aValue,1),RGB(255,255,255),0,0,False,True,0,0,0,0
End Sub

Sub SolModFlashPWM24(level)
  If DebugFlashers = True Then debug.print "SolModFlashPWM24 "& level
  ModFlashLampFlasher 1,level
  fLight24.state = level
End Sub

Sub SolModFlashPWM25(level)
  If DebugFlashers = True Then debug.print "SolModFlashPWM26 "& level
  ModFlashLampFlasher 2,level
  fLight25.state = level
End Sub

Sub SolModFlashPWM27(level)
  If DebugFlashers = True Then debug.print "SolModFlashPWM27 "& level
  ModFlashLampFlasher 3,level
  fLight27.state = level
  WheelS27.opacity = level * 60
End Sub

Sub SolModFlashPWM28(level)
  If DebugFlashers = True Then debug.print "SolModFlashPWM28 "& level
  ModFlashLampFlasher 4,level
  ModFlashLampFlasher 5,level
  fLight28.state = level
  WheelS28.opacity = level * 60
End Sub

'******************************************************
'******  END FLUPPER FLASHERLAMPS
'******************************************************

'******************************************************
'   ZFLB:  FLUPPER BUMPERS
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

FlInitBumper 1, "red"
FlInitBumper 2, "red"
FlInitBumper 3, "red"

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
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
  FlBumperBulb(nr).material = "bumperbulbmat" & nr
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
      FlBumperHighlight(nr).color = RGB(64,255,0)
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

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
  '   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  '        OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
  '        float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 ) * DayNightAdjust

  Select Case FlBumperColor(nr)
    Case "blue"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(38 - 24 * Z,130 - 98 * Z,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 500 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z + 0.97 * Z ^ 3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 10000 * (Z ^ 3) / (0.5 + DNA90)

    Case "green"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(16 + 16 * Sin(Z * 3.14),255,16 + 16 * Sin(Z * 3.14)), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
      FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 6000 * (Z ^ 3) / (1 + DNA90)

    Case "red"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 16 - 11 * Z + 16 * Sin(Z * 3.14),0), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z * 4,8 - Z * 8)

    Case "orange"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 100 - 22 * z + 16 * Sin(Z * 3.14),Z * 32), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z * 50, 0)

    Case "white"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
      FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 8 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255,255 - 20 * Z,255 - 65 * Z)
      FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20 * Z,255 - 65 * Z)
      MaterialColor "bumpertopmat" & nr, RGB(255,235 - z * 36,220 - Z * 90)

    Case "blacklight"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 1, RGB(30 - 27 * Z ^ 0.03,30 - 28 * Z ^ 0.01, 255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z ^ 3
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255 - 240 * (Z ^ 0.1),255 - 240 * (Z ^ 0.1),255)
      FlBumperSmallLight(nr).colorfull = RGB(255 - 200 * z,255 - 200 * Z,255)
      MaterialColor "bumpertopmat" & nr, RGB(255 - 190 * Z,235 - z * 180,220 + 35 * Z)

    Case "yellow"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 180 + 40 * z, 48 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

    Case "purple"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(128 - 118 * Z - 32 * Sin(Z * 3.14), 32 - 26 * Z ,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 15 + 200 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z + 0.97 * Z ^ 3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (0.5 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(128 - 60 * Z,32,255)
  End Select
End Sub

'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************

'******************************************************
'****  LAMPZ by nFozzy
'******************************************************

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz

Set Lampz = New VPMLampUpdater  'PWM inserts

InitLampsNF              ' Setup lamp assignments

LampTimer.Interval = -1   ' Using fixed value so the fading speed is same for every fps
LampTimer.Enabled = 1

Sub LampTimer_Timer()
  dim x, chglamp
  chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      If UseVPMModSol = 0 Then
        Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
      Else
        Lampz.state(chglamp(x, 0)) = chglamp(x, 1)/255.0/0.7   'PWM inserts (note the /0.7 is because the ROM does not command full brightness usually. This is likely table specific)
      End If
      'if chglamp(x, 0)=18 then debug.print "L18.state = "&Lampz.state(chglamp(x, 0))  'used for debugging inserts
      'if chglamp(x, 0)=27 then debug.print "L27.state = "&Lampz.state(chglamp(x, 0))  'used for debugging inserts
    next
  End If
  If UseVPMModSol = 0 Then Lampz.Update2  'update (fading logic only)
End Sub


Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  If UseVPMModSol = 0 Then
    If Lampz.UseFunction Then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  End If
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

Sub SetModLamp(id, val)
  Lampz.state(id) = val
  if id=118 then debug.print "SetModLamp " & val
End Sub


Sub InitLampsNF()
  Dim x, xx

  'Filtering (comment out to disable)
  If UseVPMModSol = 0 Then
    Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

    'Adjust fading speeds (max level / full MS fading time). The Modulate property must be set to 1 / max level if lamp is modulated.
    for x = 0 to 150 : Lampz.FadeSpeedUp(x) = 1/40 : Lampz.FadeSpeedDown(x) = 1/120 : next

  End If

  'Lampz Assignments
  '  In a ROM based table, the lamp ID is used to set the state of the Lampz objects

  'MassAssign is an optional way to do assignments. It'll create arrays automatically / append objects to existing arrays

  Lampz.MassAssign(11)= L11
  Lampz.MassAssign(11)= L11a
  Lampz.Callback(11) = "DisableLighting p11, 200,"

  Lampz.MassAssign(12)= L12
  Lampz.MassAssign(12)= L12a
  Lampz.Callback(12) = "DisableLighting p12, 200,"

  Lampz.MassAssign(13)= L13
  Lampz.MassAssign(13)= L13a
  Lampz.Callback(13) = "DisableLighting p13, 200,"

  Lampz.MassAssign(14)= L14
  Lampz.MassAssign(14)= L14a
  Lampz.Callback(14) = "DisableLighting p14, 200,"

  Lampz.MassAssign(15)= L15
  Lampz.MassAssign(15)= L15a
  Lampz.Callback(15) = "DisableLighting p15, 200,"

  Lampz.MassAssign(16)= L16
  Lampz.MassAssign(16)= L16a
  Lampz.Callback(16) = "DisableLighting p16, 200,"

  Lampz.MassAssign(17)= L17
  Lampz.MassAssign(17)= L17a
  Lampz.MassAssign(17)= Insert17L
  Lampz.Callback(17) = "DisableLighting p17, 200,"

  Lampz.MassAssign(18)= L18
  Lampz.MassAssign(18)= L18a
  Lampz.MassAssign(18)= Insert18L
  Lampz.Callback(18) = "DisableLighting p18, 200,"

  Lampz.MassAssign(21)= L21
  Lampz.MassAssign(21)= L21a
  Lampz.Callback(21) = "DisableLighting p21, 200,"

  Lampz.MassAssign(22)= L22
  Lampz.MassAssign(22)= L22a
  Lampz.Callback(22) = "DisableLighting p22, 200,"

  Lampz.MassAssign(23)= L23
  Lampz.MassAssign(23)= L23a
  Lampz.Callback(23) = "DisableLighting p23, 200,"

  Lampz.MassAssign(24)= L24
  Lampz.MassAssign(24)= L24a
  Lampz.Callback(24) = "DisableLighting p24, 200,"

  Lampz.MassAssign(25)= L25
  Lampz.MassAssign(25)= L25a
    Lampz.Callback(25) = "DisableLighting p25, 200,"

  Lampz.MassAssign(26)= L26
  Lampz.MassAssign(26)= L26a
  Lampz.MassAssign(26)= Insert26L
  Lampz.Callback(26) = "DisableLighting p26, 200,"

  Lampz.MassAssign(27)= L27
  Lampz.MassAssign(27)= L27a
  Lampz.MassAssign(27)= Insert27L
  Lampz.Callback(27) = "DisableLighting p27, 200,"

  Lampz.MassAssign(28)= L28
  Lampz.MassAssign(28)= L28a
  Lampz.MassAssign(28)= Insert28L
  Lampz.Callback(28) = "DisableLighting p28, 200,"

  Lampz.MassAssign(31)= L31
  Lampz.MassAssign(31)= L31a
  Lampz.Callback(31) = "DisableLighting p31, 200,"

  Lampz.MassAssign(32)= L32
  Lampz.MassAssign(32)= L32a
  Lampz.Callback(32) = "DisableLighting p32, 200,"

  Lampz.MassAssign(33)= L33
  Lampz.MassAssign(33)= L33a
  Lampz.Callback(33) = "DisableLighting p33, 200,"

  Lampz.MassAssign(34)= L34
  Lampz.MassAssign(34)= L34a
  Lampz.Callback(34) = "DisableLighting p34, 200,"

  Lampz.MassAssign(35)= L35
  Lampz.MassAssign(35)= L35a
  Lampz.Callback(35) = "DisableLighting p35, 200,"

  Lampz.MassAssign(36)= L36
  Lampz.MassAssign(36)= L36a
  Lampz.Callback(36) = "DisableLighting p36, 200,"

  Lampz.MassAssign(37)= L37
  Lampz.MassAssign(37)= L37a
  Lampz.MassAssign(37)= Insert37R
  Lampz.Callback(37) = "DisableLighting p37, 200,"

  Lampz.MassAssign(38)= L38
  Lampz.MassAssign(38)= L38a
  Lampz.MassAssign(38)= Insert38R
  Lampz.Callback(38) = "DisableLighting p38, 200,"

  Lampz.MassAssign(41)= L41
  Lampz.MassAssign(41)= L41a
  Lampz.Callback(41) = "DisableLighting p41, 200,"

  Lampz.MassAssign(42)= L42
  Lampz.MassAssign(42)= L42a
  Lampz.Callback(42) = "DisableLighting p42, 200,"

  Lampz.MassAssign(43)= L43
  Lampz.MassAssign(43)= L43a
  Lampz.Callback(43) = "DisableLighting p43, 200,"

  Lampz.MassAssign(44)= L44
  Lampz.MassAssign(44)= L44a
  Lampz.Callback(34) = "DisableLighting p34, 200,"

  Lampz.MassAssign(45)= L45
  Lampz.MassAssign(45)= L45a
  Lampz.Callback(45) = "DisableLighting p45, 200,"

  Lampz.MassAssign(46)= L46
  Lampz.MassAssign(46)= F46a
  Lampz.MassAssign(46)= F46b

  Lampz.MassAssign(47)= L47
  Lampz.MassAssign(47)= F47a
  Lampz.MassAssign(47)= F47b

  Lampz.MassAssign(48)= L48
  Lampz.MassAssign(48)= F48a
  Lampz.MassAssign(48)= F48b

  Lampz.MassAssign(51)= L51
  Lampz.MassAssign(51)= L51a
  Lampz.Callback(51) = "DisableLighting p51, 200,"

  Lampz.MassAssign(52)= L52
  Lampz.MassAssign(52)= L52a
  Lampz.Callback(52) = "DisableLighting p52, 200,"

  Lampz.MassAssign(53)= L53
  Lampz.MassAssign(53)= L53a
  Lampz.Callback(53) = "DisableLighting p53, 200,"

  Lampz.MassAssign(54)= L54
  Lampz.MassAssign(54)= L54a
  Lampz.Callback(54) = "DisableLighting p54, 200,"

  Lampz.MassAssign(55)= L55
  Lampz.MassAssign(55)= L55a
  Lampz.Callback(55) = "DisableLighting p55, 200,"

  Lampz.MassAssign(56)= L56
  Lampz.MassAssign(56)= L56a
  Lampz.MassAssign(56)= Insert56L
  Lampz.Callback(56) = "DisableLighting p56, 200,"

  Lampz.MassAssign(57)= L57
  Lampz.MassAssign(57)= L57a
  Lampz.Callback(57) = "DisableLighting p57, 200,"

  Lampz.MassAssign(58)= L58
  Lampz.MassAssign(58)= L58a
  Lampz.Callback(58) = "DisableLighting p58, 200,"

  Lampz.MassAssign(61)= L61
  Lampz.MassAssign(61)= L61a
  Lampz.Callback(61) = "DisableLighting p61, 200,"

  Lampz.MassAssign(62)= L62
  Lampz.MassAssign(62)= L62a
  Lampz.Callback(62) = "DisableLighting p62, 200,"

  Lampz.MassAssign(63)= L63
  Lampz.MassAssign(63)= L63a
  Lampz.Callback(63) = "DisableLighting p63, 200,"

  Lampz.MassAssign(64)= L64
  Lampz.MassAssign(64)= L64a
  Lampz.Callback(64) = "DisableLighting p64, 200,"

  Lampz.MassAssign(65)= L65a

  Lampz.MassAssign(66)= L66a

  Lampz.MassAssign(67)= L67a

  Lampz.MassAssign(68)= L68
  Lampz.MassAssign(68)= L68a
  Lampz.Callback(68) = "DisableLighting p68, 200,"

  Lampz.MassAssign(71)= L71
  Lampz.MassAssign(71)= L71a
  Lampz.Callback(71) = "DisableLighting p71, 200,"

  Lampz.MassAssign(72)= L72
  Lampz.MassAssign(72)= L72a
  Lampz.Callback(72) = "DisableLighting p72, 200,"

  Lampz.MassAssign(73)= L73
  Lampz.MassAssign(73)= L73a
  Lampz.Callback(73) = "DisableLighting p73, 200,"

  Lampz.MassAssign(74)= L74
  Lampz.MassAssign(74)= L74a
  Lampz.Callback(74) = "DisableLighting p74, 200,"

  Lampz.MassAssign(75)= L75
  Lampz.MassAssign(75)= L75a
  Lampz.MassAssign(75)= Insert75R
  Lampz.Callback(75) = "DisableLighting p75, 200,"

  Lampz.MassAssign(76)= L76
  Lampz.MassAssign(76)= L76a
  Lampz.MassAssign(76)= Insert76R
  Lampz.Callback(76) = "DisableLighting p76, 200,"

  Lampz.MassAssign(77)= L77
  Lampz.MassAssign(77)= L77a
  Lampz.MassAssign(77)= Insert77R
  Lampz.Callback(77) = "DisableLighting p77, 200,"

  Lampz.MassAssign(78)= L78
  Lampz.MassAssign(78)= L78a
  Lampz.MassAssign(78)= Insert78R
  Lampz.Callback(78) = "DisableLighting p78, 200,"

  Lampz.MassAssign(81)= L81
  Lampz.MassAssign(81)= F81

  Lampz.MassAssign(82)= L82
  Lampz.MassAssign(82)= F82

  Lampz.MassAssign(83)= L83
  Lampz.MassAssign(83)= F83

  Lampz.MassAssign(84)= L84
  Lampz.MassAssign(84)= F84

  Lampz.MassAssign(85)= L85L
  Lampz.MassAssign(85)= L85R
  Lampz.MassAssign(85)= L85LPlastic
  Lampz.MassAssign(85)= L85RPlastic

  Lampz.MassAssign(86)= li86

  Lampz.MassAssign(87)= L87
  Lampz.MassAssign(87)= L87Plastic
'
  Lampz.MassAssign(88)= L88
  Lampz.MassAssign(88)= L88Plastic

  For x = 0 to 150: Lampz.State(x) = 0: Next
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
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks
'Version 0.14 - Updated to support modulated signals - Niwak

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

End Class


'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function

'Helper functions
Sub Proc(String, Callback)  'proc using a string and one argument
  'On Error Resume Next
  Dim p
  Set P = GetRef(String)
  P Callback
  If err.number = 13 Then  MsgBox "Proc error! No such procedure: " & vbNewLine & String
  If err.number = 424 Then MsgBox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the end of a 1 dimensional array
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

'******************************************************
'****  END LAMPZ
'******************************************************

'******************************************************
'******  VPW TWEAK MENU
'******************************************************

Dim ColorLUT : ColorLUT = 1
Dim Anaglyph : Anaglyph = 0                 'When set to 1 in the menu this will overwrite the default LUT slot to use the VPW Original 1 on 1
Dim LightLevel : LightLevel = 0.07      'Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim VolumeDial : VolumeDial = 0.8     'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5   'Level of ramp rolling volume. Value between 0 and 1

Dim AmbientBallShadowOn : AmbientBallShadowon = 1   ' Ambient Ball Shadows. 0 = Disabled, 1 = Enabled

Dim RenderProbeOpt : RenderProbeOpt = 1         ' 0 = No Refraction Probes (best performance), 1 = Full Refraction Probes (best visual)
Dim PlayfieldReflections : PlayfieldReflections = 100 ' Defines the reflection strength of the (dynamic) table elements on the playfield (0-100) / 5 = 20 Max
Dim ReflOpt : ReflOpt = 1                 ' 0 = Reflections off, 1 = Reflections on
Dim LightReflOpt : LightReflOpt = 1               ' 0 = Light Reflections off, 1 = Light Reflections on
Dim InsertReflOpt : InsertReflOpt = 1             ' 0 = Insert Reflections off, 1 = Insert Reflections on
Dim BackglassReflOpt : BackglassReflOpt = 1       ' 0 = Backglass Reflections off, 1 = Backglass Reflections on
Dim SidewallFlashOpt : SidewallFlashOpt = 1         ' 0 = Sidewall Flashers off, 1 = Sidewall Flashers and Lights on, 2 = Sidewall Flashers on, 3 = Sidewall Lights on

Dim FlipperMod, InstMod, SideBladeMod, GlassMod, GIColorMod, BackglassColorMod, LogoPosterMod, Toppermod, CabinetMod, CabinetMetal, LegendsCabMod, Railsmod
Dim EnviroSoundsOn              ' Environment Sounds (o - off, 1 - On)
Dim EnviroSoundsLast: EnviroSoundsLast = 0


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
  'Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 27, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White", _
    "Fleep Natural Dark 1", "Fleep Natural Dark 2", "Fleep Warm Dark", "Fleep Warm Bright", "Fleep Warm Vivid Soft", "Fleep Warm Vivid Hard", "Skitso Natural & Balance", "Skitso Natural High Contrast", _
    "3rdaxis THX Standard", "Callev Brightness & Contrast", "Hauntfreaks Desaturated", "Tomate Washed out", "VPW Original 1 on 1", "Bassgeige", "Blacklight", "B&W Comic Book"))
  if ColorLUT = 1 And Anaglyph = 0 Then Table1.ColorGradeImage = "" '"Normal"
  if ColorLUT = 1 And Anaglyph = 1 Then Table1.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" '"VPW Original 1 on 1"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10" '"Desaturated 10%"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20" '"Desaturated 20%"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30" '"Desaturated 30%"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40" '"Desaturated 40%"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50" '"Desaturated 50%"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60" '"Desaturated 60%"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70" '"Desaturated 70%"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80" '"Desaturated 80%"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90" '"Desaturated 90%"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100" '"Black 'n White"
  if ColorLUT = 12 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-natural-dark-1" '"Fleep Natural Dark 1"
  if ColorLUT = 13 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-natural-dark-2" '"Fleep Natural Dark 2"
  if ColorLUT = 14 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-dark" '"Fleep Warm Dark "
  if ColorLUT = 15 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-bright" '"Fleep Warm Bright"
  if ColorLUT = 16 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-vivid-soft" '"Fleep Warm Vivid Soft"
  if ColorLUT = 17 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-vivid-hard" '"Fleep Warm Vivid Hard"
  if ColorLUT = 18 Then Table1.ColorGradeImage = "colorgradelut256x16-skitso-natural-and-balance" '"Skitso Natural & Balance"
  if ColorLUT = 19 Then Table1.ColorGradeImage = "colorgradelut256x16-skitso-natural-high-contrast" '"Skitso Natural High Contrast"
  if ColorLUT = 20 Then Table1.ColorGradeImage = "colorgradelut256x16-3rdaxis-thx-standard" '"3rdaxis THX Standard"
  if ColorLUT = 21 Then Table1.ColorGradeImage = "colorgradelut256x16-callev-brightness-and-contrast" '"Callev Brightness & Contrast"
  if ColorLUT = 22 Then Table1.ColorGradeImage = "colorgradelut256x16-hauntfreaks-desaturated" '"Hauntfreaks Desaturated"
  if ColorLUT = 23 Then Table1.ColorGradeImage = "colorgradelut256x16-tomate-washed-out" '"Tomate Washed out"
  if ColorLUT = 24 Then Table1.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" '"VPW Original 1 on 1"
  if ColorLUT = 25 Then Table1.ColorGradeImage = "colorgradelut256x16-bassgeige" '"Bassgeige"
  if ColorLUT = 26 Then Table1.ColorGradeImage = "colorgradelut256x16-blacklight" '"Blacklight"
  if ColorLUT = 27 Then Table1.ColorGradeImage = "colorgradelut256x16-b-and-w-comic-book" '"B&W Comic Book"
  'Anaglyph
  Anaglyph = Table1.Option("Anaglyph 3D/Headtracking Fix", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupAnaglyph
    'Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)
  'Ambient BallShadow
  AmbientBallShadowOn = Table1.Option("Ambient Ballshadow", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetupAmbientBallshadow
  'Reflections
  ReflOpt = Table1.Option("Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetRefl ReflOpt
  'Playfield Reflections
  PlayfieldReflections = Table1.Option("Playfield Reflections", 0, 1, 0.05, 1, 1)
  Table1.PlayfieldReflectionStrength = (PlayfieldReflections * 20)
  'Light Reflections
  LightReflOpt = Table1.Option("Light Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetLightRefl LightReflOpt
  'Insert Reflections
  InsertReflOpt = Table1.Option("Insert Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetInsertRefl InsertReflOpt
  'Backglass Reflections
  BackglassReflOpt = Table1.Option("Backglass Reflections", 0, 3, 1, 1, 0, Array("Off", "Backglass & DMD", "Backglass only", "DMD only"))
  SetBackglassRefl BackglassReflOpt
  'Sidewall Flashers
  SidewallFlashOpt = Table1.Option("Sidewall Flasher Effects", 0, 3, 1, 1, 0, Array("Off", "Lights and Flashers", "Lights Only", "Flashers Only"))
  SetSidewallFlash SidewallFlashOpt
  'Refraction Probe Setting
  RenderProbeOpt = Table1.Option("Refraction Probe Setting", 0, 1, 1, 1, 0, Array("No Refraction (best performance)", "Full Refraction (best visuals)"))
  SetRenderProbes RenderProbeOpt
  'Sideblades
  SideBladeMod = Table1.Option("Side Blades", 0, 2, 1, 0, 0, Array("Original", "Retrorefurbs 1", "Retrorefurbs 2"))
  SetupBlades
  'Instruction Cards
  InstMod = Table1.Option("Instruction Cards", 0, 5, 1, 0, 0, Array("Default Instruction Cards", "Custom Cards 1", "Custom Cards 2", "Custom Cards 3", "Custom Cards 4", "Custom Cards 5"))
  SetupInstructionCards
  'Custom Flippers
  FlipperMod = Table1.Option("Flipperbats", 0, 6, 1, 0, 0, Array("Original", "Red white", "Blue white", "Green white", "Orange white", "Red blue white", "Red green white"))
  SetupFlippers
  'GI
  GIColorMod = Table1.Option("GI", 0, 7, 1, 0, 0, Array("White bulbs and GI", "Red bulbs and GI", "Blue bulbs and GI", "Green bulbs and GI", "Orange bulbs and GI", "Red and blue GI", "Red and green GI", "Rainbow RGB"))
  ResetGI
  'BackglassColor
  BackglassColorMod = Table1.Option("Backglass Color", 0, 1, 1, 0, 0, Array("Warm Incandescent", "Bright White"))
  SetupBackglassColorMod
  'LegendsCab
  LegendsCabMod = Table1.Option("Legends Cabinet", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupLegendsCab
  'LegendsCab
  Railsmod = Table1.Option("Rails Cabinet Mode", 0, 3, 1, 0, 0, Array("Off", "Top And Sides", "Top Only", "Sides Only"))
  SetupRailsmod
  'Scratched Glass
  GlassMod = Table1.Option("Scratched Glass", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupGlass
  'Cabinet Metal
  CabinetMetal = Table1.Option("Cabinet Metal", 0, 1, 1, 0, 0, Array("Standard", "Orange"))
  SetupCabinetMetal
  'VR Topper
  TopperMod = Table1.Option("VR Topper", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetupTopper
  'VR Cabinet
  CabinetMod = Table1.Option("VR Cabinet", 0, 1, 1, 0, 0, Array("Standard Artwork", "Yellow Artwork"))
  SetupCabinet
  'Environment Sounds
  EnviroSoundsOn = Table1.Option("Environment Sounds", 0, 1, 1, 0, 0, Array("Off", "On"))
  'VR Room
    VRRoomChoice = Table1.Option("VR Room (VR-FSS)", 1, 3, 1, 1, 0, Array("Roller Coaster", "Minimal Room", "Mixed Reality"))
  SetupRoom
  'VR Minimal Room Logo & Poster (VR-FSS)
  LogoPosterMod = Table1.Option("VR Minimal Room Logo & Poster (VR-FSS)", 0, 3, 1, 1, 0, Array("Off", "Logo and Poster", "Logo", "Poster"))
  SetupLogoPoster
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel
  If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Sub SetupAnaglyph
  if ColorLUT = 1 And Anaglyph = 0 Then Table1.ColorGradeImage = "" '"Normal"
  if ColorLUT = 1 And Anaglyph = 1 Then Table1.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" '"VPW Original 1 on 1"
End Sub

Sub SetupAmbientBallshadow
  If AmbientBallShadowOn = 1 Then
    For each xx in GILights : xx.Shadows = 1 : Next
  Else
    For each xx in GILights : xx.Shadows = 0 : Next
  End If
End Sub

Sub SetRefl(Opt)
  Dim xx
  Select Case Opt
    Case 0:
      For each xx in Reflections: xx.ReflectionEnabled = False : Next
    Case 1:
      For each xx in Reflections: xx.ReflectionEnabled = True : Next
  End Select
End Sub

Sub SetLightRefl(Opt)
  Dim xx
  Select Case Opt
    Case 0:
      For each xx in LightReflections: xx.ShowReflectionOnBall = False : Next
    Case 1:
      For each xx in LightReflections: xx.ShowReflectionOnBall = True : Next
  End Select
End Sub

Sub SetInsertRefl(Opt)
  Dim xx
  Select Case Opt
    Case 0:
      For each xx in InsertReflections: xx.ShowReflectionOnBall = False : Next
    Case 1:
      For each xx in InsertReflections: xx.ShowReflectionOnBall = True : Next
  End Select
End Sub

Sub SetBackglassRefl(Opt)
  Dim xx
  Select Case Opt
    Case 0:
      For each xx in BackglassReflections: xx.ReflectionEnabled = False : Next
      For each xx in BackglassGIReflections: xx.ReflectionEnabled = False : Next
      Pincab_DMD_ReflectionDT.Visible = 0
      Pincab_DMD_Reflection.Visible = 0
      If VRMode = False Then
        For Each VR_Obj in BackglassReflections : VR_Obj.Visible = 0 : Next
        For each VR_Obj in BackglassGIReflections: VR_Obj.Visible = 0  : Next
      End If
    Case 1:
      For each xx in BackglassReflections: xx.ReflectionEnabled = True : Next
      For each xx in BackglassGIReflections: xx.ReflectionEnabled = True : Next
      For Each VR_Obj in BackglassReflections : VR_Obj.Visible = 1 : Next
      For Each VR_Obj in BackglassGIReflections : VR_Obj.Visible = 1 : Next
      If Table1.ShowDT = True And RenderingMode <> 2 Then
        Pincab_DMD_ReflectionDT.Visible = 1
        Pincab_DMD_Reflection.Visible = 0
      Else
        Pincab_DMD_ReflectionDT.Visible = 0
        Pincab_DMD_Reflection.Visible = 1
      End If
    Case 2:
      For each xx in BackglassReflections: xx.ReflectionEnabled = True : Next
      For each xx in BackglassGIReflections: xx.ReflectionEnabled = True : Next
      For Each VR_Obj in BackglassReflections : VR_Obj.Visible = 1 : Next
      For Each VR_Obj in BackglassGIReflections : VR_Obj.Visible = 1 : Next
      Pincab_DMD_ReflectionDT.Visible = 0
      Pincab_DMD_Reflection.Visible = 0
    Case 3:
      For each xx in BackglassReflections: xx.ReflectionEnabled = False : Next
      For each xx in BackglassGIReflections: xx.ReflectionEnabled = False : Next
      If VRMode = False Then
        For Each VR_Obj in BackglassReflections : VR_Obj.Visible = 0 : Next
        For each VR_Obj in BackglassGIReflections: VR_Obj.Visible = 0  : Next
      End If
      If Table1.ShowDT = True And RenderingMode <> 2 Then
        Pincab_DMD_ReflectionDT.Visible = 1
        Pincab_DMD_Reflection.Visible = 0
      Else
        Pincab_DMD_ReflectionDT.Visible = 0
        Pincab_DMD_Reflection.Visible = 1
      End If
  End Select
End Sub

Sub SetSidewallFlash(Opt)
  Select Case Opt
    Case 0:
      For each xx in SidewallFlashers: xx.visible = False : Next
      For each xx in SidewallLights: xx.visible = False : Next
    Case 1:
      For each xx in SidewallFlashers: xx.visible = True : Next
      For each xx in SidewallLights: xx.visible = True : Next
    Case 2:
      For each xx in SidewallFlashers: xx.visible = False : Next
      For each xx in SidewallLights: xx.visible = True : Next
    Case 3:
      For each xx in SidewallFlashers: xx.visible = True : Next
      For each xx in SidewallLights: xx.visible = False : Next
  End Select
End Sub

Sub SetRenderProbes(Opt)
  On Error Resume Next
    Select Case Opt
      Case 0:
        RampCometPlastic.RefractionProbe = ""
        RampCometTopPlastic.RefractionProbe = ""
        RampHurricanePlastic2.RefractionProbe = ""
        RampHurricanePlastic1.RefractionProbe = ""
      Case 1:
        RampCometPlastic.RefractionProbe = "CenterRamp"
        RampCometTopPlastic.RefractionProbe = "CenterRamp"
        RampHurricanePlastic2.RefractionProbe = "LeftRamp"
        RampHurricanePlastic1.RefractionProbe = "RightRamp"
    End Select
  On Error Goto 0
End Sub

Sub SetupInstructionCards
  If InstMod = 0 Then
    LeftCardPrim.image = "InstCard1"
    RightCardPrim.image = "InstCard2"
  End If
  If InstMod = 1 Then
    LeftCardPrim.image = "InstCard_Custom1_1"
    RightCardPrim.image = "InstCard_Custom1_2"
  End If
  If InstMod = 2 Then
    LeftCardPrim.image = "InstCard_Custom2_1"
    RightCardPrim.image = "InstCard_Custom2_2"
  End If
  If InstMod = 3 Then
    LeftCardPrim.image = "InstCard_Custom3_1"
    RightCardPrim.image = "InstCard_Custom3_2"
  End If
  If InstMod = 4 Then
    LeftCardPrim.image = "InstCard_Custom4_1"
    RightCardPrim.image = "InstCard_Custom4_2"
  End If
  If InstMod = 5 Then
    LeftCardPrim.image = "InstCard_Custom5_1"
    RightCardPrim.image = "InstCard_Custom5_2"
  End If
End Sub

Sub SetupBlades
  If SideBladeMod = 0 Then
    PinCab_Blades_Left.image = "PinCab_BladesGIOff"
    PinCab_Blades_Right.image = "PinCab_BladesGIOff"
  End If
  If SideBladeMod = 1 Then
    PinCab_Blades_Left.image = "PinCab_Blades_Custom1GIOff"
    PinCab_Blades_Right.image = "PinCab_Blades_Custom1GIOff"
  End if
  If SideBladeMod = 2 Then
    PinCab_Blades_Left.image = "PinCab_Blades_Custom2GIOff"
    PinCab_Blades_Right.image = "PinCab_Blades_Custom2GIOff"
  End if
End Sub

Sub SetupFlippers
  If FlipperMod = 0 Then
    batleft.image = "BatsYellowRed"
    batright.image = "BatsYellowRed"
  End If
  If FlipperMod = 1 Then
    batleft.image = "BatsWhiteRed"
    batright.image = "BatsWhiteRed"
  End If
  If FlipperMod = 2 Then
    batleft.image = "BatsWhiteBlue"
    batright.image = "BatsWhiteBlue"
  End If
  If FlipperMod = 3 Then
    batleft.image = "BatsWhiteGreen"
    batright.image = "BatsWhiteGreen"
  End If
  If FlipperMod = 4 Then
    batleft.image = "BatsWhiteOrange"
    batright.image = "BatsWhiteOrange"
  End If
  If FlipperMod = 5 Then
    batleft.image = "BatsWhiteRed"
    batright.image = "BatsWhiteBlue"
  End If
  If FlipperMod = 6 Then
    batleft.image = "BatsWhiteRed"
    batright.image = "BatsWhiteGreen"
  End If
End Sub

Sub SetupBackglassColorMod
  If BackglassColorMod = 0 Then
    WheelGI1.image = "HurricaneWheelGI1Warm"
    WheelGI2.image = "HurricaneWheelGI2Warm"
    WheelGI3.image = "HurricaneWheelGI3Warm"
  End If
  If BackglassColorMod = 1 Then
    WheelGI1.image = "HurricaneWheelGI1"
    WheelGI2.image = "HurricaneWheelGI2"
    WheelGI3.image = "HurricaneWheelGI3"
  End if
End Sub

Sub Setupcabinet
  If CabinetMod = 1 Then
    Pincab_Cabinet.image = "PinCab_Cabinet_Yellow"
  Else
    Pincab_Cabinet.image = "PinCab_Cabinet"
  End If
End Sub

Sub SetupcabinetMetal
  If CabinetMetal = 1 Then
    Pincab_Backbox_Trim_Left.material = "Metal_Orange_Powdercoat"
    Pincab_Backbox_Trim_Right.material = "Metal_Orange_Powdercoat"
    Pincab_Hinges.material = "Metal_Orange_Powdercoat"
    Pincab_Plunger_Housing.material = "Metal_Orange_Powdercoat"
    Pincab_Rails.material = "Metal_Orange_Powdercoat"
    Pincab_BBoxL_S1.material = "Metal_Orange_Powdercoat"
    Pincab_BBoxL_S2.material = "Metal_Orange_Powdercoat"
    Pincab_BBoxR_S1.material = "Metal_Orange_Powdercoat"
    Pincab_BBoxR_S2.material = "Metal_Orange_Powdercoat"
    Pincab_DMD_Decal_Lip.material = "Metal_Orange_Powdercoat"
    Pincab_Backbox_Trim_Left.material = "Metal_Orange_Powdercoat"
    Pincab_Backbox_Trim_Right.material = "Metal_Orange_Powdercoat"
    Pincab_Front_Left_Leg.material = "Metal_Orange_Powdercoat"
    Pincab_Front_Right_Leg.material = "Metal_Orange_Powdercoat"
    Pincab_Rear_Left_Leg.material = "Metal_Orange_Powdercoat"
    Pincab_Rear_Right_Leg.material = "Metal_Orange_Powdercoat"
    Pincab_Start_Button_Wall.material = "Metal_Orange_Powdercoat"
    Pincab_TopRail.material = "Metal_Orange_Powdercoat"
    PinCab_CoinDoor.material = "Metal_Orange_Powdercoat"
    PinCab_CoinDoorTrim.material = "Metal_Orange_Powdercoat"
  Else
    Pincab_Backbox_Trim_Left.material = "Metal_Black_Powdercoat"
    Pincab_Backbox_Trim_Right.material = "Metal_Black_Powdercoat"
    Pincab_Hinges.material = "Metal_Black_Powdercoat"
    Pincab_Plunger_Housing.material = "MetalShiny"
    Pincab_Rails.material = "MetalShiny"
    Pincab_BBoxL_S1.material = "Metal_Black_Powdercoat"
    Pincab_BBoxL_S2.material = "Metal_Black_Powdercoat"
    Pincab_BBoxR_S1.material = "Metal_Black_Powdercoat"
    Pincab_BBoxR_S2.material = "Metal_Black_Powdercoat"
    Pincab_DMD_Decal_Lip.material = "Metal_Black_Powdercoat"
    Pincab_Backbox_Trim_Left.material = "Metal_Black_Powdercoat"
    Pincab_Backbox_Trim_Right.material = "Metal_Black_Powdercoat"
    Pincab_Front_Left_Leg.material = "MetalShiny"
    Pincab_Front_Right_Leg.material = "MetalShiny"
    Pincab_Rear_Left_Leg.material = "MetalShiny"
    Pincab_Rear_Right_Leg.material = "MetalShiny"
    Pincab_Start_Button_Wall.material = "Metal_Black_Powdercoat"
    Pincab_TopRail.material = "Metal_Black_Powdercoat"
    PinCab_CoinDoor.material = "Metal_Black_Powdercoat"
    PinCab_CoinDoorTrim.material = "Metal_Black_Powdercoat"
  End If
End Sub

Sub SetupGlass
  If GlassMod = 1 Then
    PinCab_Glass_scratches.Visible = True
  Else
    PinCab_Glass_scratches.Visible = False
  End If
End Sub

Sub SetupTopper
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    VRMode = True
    If TopperMod = 1 Then
      For Each VR_Obj in VRTopper:VR_Obj.Visible = 1:Next
    Else
      For Each VR_Obj in VRTopper:VR_Obj.Visible = 0:Next
    End If
  End if
End Sub

Sub SetupLegendsCab
  If Table1.ShowDT = False And Table1.ShowFSS = False And RenderingMode <> 2 And LegendsCabMod = 1 Then
    Flasherbloom1.Height = 500
    Flasherbloom2.Height = 500
    Flasherbloom3.Height = 500
    Flasherbloom4.Height = 500
    Flasherbloom5.Height = 500
    Flasherbloom6.Height = 500
    Flasherbloom7.Height = 500
    Flasherlampbloom1.Height = 500
    Flasherlampbloom2.Height = 500
    Flasherlampbloom3.Height = 500
    Flasherlampbloom4.Height = 500
    Flasherlampbloom5.Height = 500
  Else
    Flasherbloom1.Height = 220
    Flasherbloom2.Height = 220
    Flasherbloom3.Height = 220
    Flasherbloom4.Height = 220
    Flasherbloom5.Height = 220
    Flasherbloom6.Height = 220
    Flasherbloom7.Height = 220
    Flasherlampbloom1.Height = 220
    Flasherlampbloom2.Height = 220
    Flasherlampbloom3.Height = 220
    Flasherlampbloom4.Height = 220
    Flasherlampbloom5.Height = 220
  End If
End Sub

Sub SetupRailsMod
  If Table1.ShowDT = False Then
    If RailsMod = 1 Then
      PinCab_Rails.Visible = True
      PinCab_TopRail.Visible = True
    End If
    If RailsMod = 2 Then
      PinCab_Rails.Visible = False
      PinCab_TopRail.Visible = True
    End If
    If RailsMod = 3 Then
      PinCab_Rails.Visible = True
      PinCab_TopRail.Visible = False
    End If
    If RailsMod = 0 Then
      PinCab_Rails.Visible = False
      PinCab_TopRail.Visible = False
    End If
  Else
    PinCab_Rails.Visible = True
    PinCab_TopRail.Visible = True
  End If
End Sub

Sub SetupRoom
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    VRMode = True
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
    If VRRoomChoice = 1 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VRMEGARoom : VR_Obj.Visible = 1 : Next
      Room360.Visible = 0
      VR_Mega061Car1.PlayAnimEndless (0.4)
      VR_Mega052Car2.PlayAnimEndless (0.4)
      VR_Mega068Car3.PlayAnimEndless (0.4)
      VR_Logo.Visible = False
      VR_Poster.Visible = False
      if EnviroSoundsOn = 1 and EnviroSoundsLast = 0 Then
        PlaySound EnviroSound, -1, 1.0, 0.0, 0.0, 0, 0, 0, 0.0
      ElseIf EnviroSoundsOn = 0 Then
        StopSound EnviroSound
      end If
      EnviroSoundsLast = EnviroSoundsOn
    End If
    If VRRoomChoice = 2 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
      For Each VR_Obj in VRMEGARoom : VR_Obj.Visible = 0 : Next
      If LogoPosterMod = 0 Then
        VR_Logo.Visible = False
        VR_Poster.Visible = False
      End If
      If LogoPosterMod = 1 Then
        VR_Logo.Visible = True
        VR_Poster.Visible = True
      End If
      If LogoPosterMod = 2 Then
        VR_Logo.Visible = True
        VR_Poster.Visible = False
      End If
      If LogoPosterMod = 3 Then
        VR_Logo.Visible = False
        VR_Poster.Visible = True
      End If
      Room360.Visible = 0
      VR_Mega061Car1.StopAnim()
      VR_Mega052Car2.StopAnim()
      VR_Mega068Car3.StopAnim()
      StopSound EnviroSound
      EnviroSoundsLast = 0
    End If
    If VRRoomChoice = 3 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VRMEGARoom : VR_Obj.Visible = 0 : Next
      Room360.Visible = 1
      VR_Logo.Visible = False
      VR_Poster.Visible = False
      VR_Mega061Car1.StopAnim()
      VR_Mega052Car2.StopAnim()
      VR_Mega068Car3.StopAnim()
      StopSound EnviroSound
      EnviroSoundsLast = 0
    End If
  End If
End Sub

Sub SetupLogoPoster
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    If VRRoomChoice = 2 Then
      If LogoPosterMod = 0 Then
        VR_Logo.Visible = False
        VR_Poster.Visible = False
      End If
      If LogoPosterMod = 1 Then
        VR_Logo.Visible = True
        VR_Poster.Visible = True
      End If
      If LogoPosterMod = 2 Then
        VR_Logo.Visible = True
        VR_Poster.Visible = False
      End If
      If LogoPosterMod = 3 Then
        VR_Logo.Visible = False
        VR_Poster.Visible = True
      End If
    Else
      VR_Logo.Visible = False
      VR_Poster.Visible = False
    End If
  End If
End Sub

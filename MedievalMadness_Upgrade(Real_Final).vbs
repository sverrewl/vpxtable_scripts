' Medieval Madness - IPDB No. 4032
' © Williams 1997
' Remaster by Skitso, rothbauerw and bord
' VPX Original by ninuzzu & Tom Tower
' Thanks to all the authors (JPSalas,Dozer,Pinball Ken, Jamin, Macho, Joker, PacDude) who made this table before.
' Thanks to Clark Kent for the pics and the advices
' Thanks to zany for the domes and bumpers
' Thanks to knorr for some sound effects I borrowed from his tables
' Thanks to VPDev Team for the freaking amazing VPX

Option Explicit
Randomize

'Options
Const BallBright = 1        '0 - Normal, 1 - Bright
Const TrollSpeed = 2        '0 - slow, 1 - medium, 2 - fast, 3 - really fast
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7
Const VolumeDial = 0.8
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

Dim Ballsize,BallMass
BallSize = 50
BallMass = 1
Const tnob = 5

Const AmbientBallShadowOn = 1
Const DynamicBallShadowsOn = 1
Const RubberizerEnabled = 1

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode
Const UseVPMModSol = 1

LoadVPM "01560000", "WPC.VBS", 3.50
Const cSingleLFlip = 0

Const cSingleRFlip = 0

'********************
'Standard definitions
'********************

dim x

'Const cGameName = "mm_10" 'Williams official rom
'Const cGameName = "mm_109" 'free play only
'Const cGameName="mm_109b" 'unofficial
Const cGameName="mm_109c" 'unofficial profanity rom

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = ""

'************************************************************************
'            INIT TABLE
'************************************************************************

' using table width and height in script slows down the performance

Dim MMBall1, MMBall2, MMBall3, MMBall4

Sub Table1_Init
LoadLUT
  vpmInit Me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Medieval Madness (Williams 1997)"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 1
    .Hidden = DesktopMode
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With

  On Error Goto 0

  ' Init switches
  Controller.Switch(22) = 1 'close coin door
  Controller.Switch(24) = 0 'always closed

  ' Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  ' Nudging
  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 1
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  '************  Trough **************************
  Set MMBall4 = sw35.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set MMBall3 = sw34.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set MMBall2 = sw33.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set MMBall1 = sw32.CreateSizedballWithMass(Ballsize/2,Ballmass)

  Controller.Switch(35) = 1
  Controller.Switch(34) = 1
  Controller.Switch(33) = 1
  Controller.Switch(32) = 1

  SolCastle 0
  SolMod17 0
  SolMod18 0
  SolMod19 0
  SolMod20 0
  SolMod21 0
  SolMod22 0
  SolMod23 0
  SolMod24 0
  SolMod25 0

  LockPost.IsDropped=1
  LockPostP.IsDropped=1
  TrollP1X.IsDropped = 1
  sw45.IsDropped = 1
  TrollP2X.IsDropped = 1
  LTT.collidable = False
  RTT.collidable = False
  sw46.IsDropped = 1
  BW1.isdropped = 1
  BW2.isdropped = 1
  MoatKick.collidable = False

  FlasherVisibility

  InitOptions
  InitLamps

  UpdateGI 0,0
  UpdateGI 1,0
  UpdateGI 2,0
End Sub

'//////////////////////////////////////////////////////////////////////
'// LUT
'//////////////////////////////////////////////////////////////////////

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
'12 = VPW Original 1 to 1
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book

Dim LUTset, DisableLUTSelector, LutToggleSound, bLutActive
LutToggleSound = True
LoadLUT
'LUTset = 0     ' Override saved LUT for debug
SetLUT
DisableLUTSelector = 0  ' Disables the ability to change LUT option with magna saves in game when set to 1

Sub Table1_exit:SaveLUT:Controller.Pause = False:Controller.Stop:End Sub

Sub SetLUT  'AXS
  Table1.ColorGradeImage = "LUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  Select Case LUTSet
    Case 0: LUTBox.text = "Fleep Natural Dark 1"
    Case 1: LUTBox.text = "Fleep Natural Dark 2"
    Case 2: LUTBox.text = "Fleep Warm Dark"
    Case 3: LUTBox.text = "Fleep Warm Bright"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard"
    Case 6: LUTBox.text = "Skitso Natural and Balanced"
    Case 7: LUTBox.text = "Skitso Natural High Contrast"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast"
    Case 10: LUTBox.text = "HauntFreaks Desaturated"
      Case 11: LUTBox.text = "Tomate washed out"
        Case 12: LUTBox.text = "VPW original 1on1"
        Case 13: LUTBox.text = "bassgeige"
        Case 14: LUTBox.text = "blacklight"
        Case 15: LUTBox.text = "B&W Comic Book"
    Case 16: LUTBox.text = "Skitso New ColorLut"
  End Select
  LUTBox.TimerEnabled = 1
End Sub

Sub SaveLUT
  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if

  if LUTset = "" then LUTset = 0 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "MedievalMadnessLUT.txt",True)
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub
Sub LoadLUT
    bLutActive = False
  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    LUTset=0
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "MedievalMadnessLUT.txt") then
    LUTset=0
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "MedievalMadnessLUT.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      LUTset=0
      Exit Sub
    End if
    LUTset = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

'************************************************************************
'             KEYS
'************************************************************************

Sub Table1_KeyDown(ByVal Keycode)
'LUT controls
If keycode = LeftMagnaSave Then bLutActive = True
    If keycode = RightMagnaSave Then
            If bLutActive Then
                    if DisableLUTSelector = 0 then
                LUTSet = LUTSet  - 1
                if LutSet < 0 then LUTSet = 16
                SetLUT
                ShowLUT
            End If
        End If
        End If


If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
                Select Case Int(rnd*3)
                        Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
                        Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
                        Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
End Select
        End If
  If keycode = StartGameKey Then Controller.Switch(13) = 1
  If keycode = PlungerKey Then Controller.Switch(11) = 1
  If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge",0):TwrShake
  If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge",0):TwrShake
  If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge",0):TwrShake
If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress End If

If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress End If

  'if keycode = LeftMagnaSave Then
    'BallLocation.text = "X:" & INT(MMBall1.x) & " Y:" & INT(MMBAll1.y) & " Z:" & INT(MMBall1.z) & vbnewline & _
      '"X:" & INT(MMBall2.x) & " Y:" & INT(MMBAll2.y) & " Z:" & INT(MMBall2.z) & vbnewline & _
      '"X:" & INT(MMBall3.x) & " Y:" & INT(MMBAll3.y) & " Z:" & INT(MMBall3.z) & vbnewline & _
      '"X:" & INT(MMBall4.x) & " Y:" & INT(MMBAll4.y) & " Z:" & INT(MMBall4.z)
  'End If
  'If Keycode = RightMagnaSave Then
  ' BallLocation.text = ""
  'End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
'LUT controls
If keycode = LeftMagnaSave Then bLutActive = False
  If keycode = StartGameKey Then Controller.Switch(13) = 0
  If keycode = PlungerKey Then Controller.Switch(11) = 0

If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress End If

If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress End If

  If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_unPaused:Controller.Pause = False:End Sub

'************************************************************************
'            SOLENOIDS
'************************************************************************

SolCallback(1) = "Auto_Plunger"             'AutoPlunger
SolCallback(2) = "SolBallRelease"           'Trough Eject
SolCallback(3) = "SolMoat"                'Left Popper
SolCallback(4) = "SolCastle"                'Castle Towers
SolCallback(5) = "SolCastlegatePow"           'Castle Gate Power
SolCallback(6) = "SolCastlegateHold"            'Castle Gate Hold
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(8) = "SolCatapult"              'Catapult
SolCallback(9) = "SolMerlin"                'Right Eject
'SolCallback(10)=""                   'Lsling
'SolCallback(11)=""                   'Rsling
'SolCallback(12)=""                   'LBump
'SolCallback(13)=""                   'TBump
'SolCallback(14)=""                   'RBump
SolCallback(15)= "SolTowerDivPow"           'Tower Diverter Power
SolCallback(16)= "SolTowerDivHold"            'Tower Diverter Hold
SolModcallback(17) = "SolMod17"             'Left Side Low Flasher + Insert Panel
SolModcallback(18) = "SolMod18"             'Left Ramp Flasher + Insert Panel
SolModcallback(19) = "SolMod19"             'Left Side High Flasher + Insert Panel
SolModcallback(20) = "SolMod20"             'Right Side High Flasher + Insert Panel
SolModcallback(21) = "SolMod21"             'Right Ramp Flasher + Insert Panel
SolModcallback(22) = "SolMod22"             'Castle Right Side Flasher + Backpanel
SolModcallback(23) = "SolMod23"             'Right Side Low Flashers
SolModcallback(24) = "SolMod24"             'Moat Flashers
SolModcallback(25) = "SolMod25"             'Castle Left Side Flashers + BackPanel
Solcallback(26) = "SolTowerLock"            'Tower Lock Post
SolCallback(27) = "gate3.open ="            'Right Gate
Solcallback(28) = "gate2.open ="            'Left Gate

Solcallback(33) = "SolLeftTrollPow"           'Left Troll Power
Solcallback(34) = "SolLeftTrollHold"          'Left Troll Hold
Solcallback(35) = "SolRightTrollPow"          'Right Troll Power
Solcallback(36) = "SolRightTrollHold"         'Right Troll Hold
Solcallback(37) = "SolDrawBridge"           'Drawbridge Motor

'*********************************************************
'           Functions
'*********************************************************

Function Degrees(radians)
  Degrees = 180 * radians / PI
End Function

Function ASin(val)
    ASin = 2 * Atn(val / (1 + Sqr(1 - (val * val))))
End Function

Function ACos(val)
    ACos = PI / 2 - ASin(val)
End Function

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
'       NFOZZY'S FLIPPERS
'******************************************************


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
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

' This subroutine updates the flipper shadows and visual primitives
Sub FlipperVisualUpdate
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -2.7
'        AddPt "Polarity", 2, 0.33, -2.7
'        AddPt "Polarity", 3, 0.37, -2.7
'        AddPt "Polarity", 4, 0.41, -2.7
'        AddPt "Polarity", 5, 0.45, -2.7
'        AddPt "Polarity", 6, 0.576,-2.7
'        AddPt "Polarity", 7, 0.66, -1.8
'        AddPt "Polarity", 8, 0.743, -0.5
'        AddPt "Polarity", 9, 0.81, -0.5
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -3.7
'        AddPt "Polarity", 2, 0.33, -3.7
'        AddPt "Polarity", 3, 0.37, -3.7
'        AddPt "Polarity", 4, 0.41, -3.7
'        AddPt "Polarity", 5, 0.45, -3.7
'        AddPt "Polarity", 6, 0.576,-3.7
'        AddPt "Polarity", 7, 0.66, -2.3
'        AddPt "Polarity", 8, 0.743, -1.5
'        AddPt "Polarity", 9, 0.81, -1
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'


'*******************************************
'  Late 80's early 90's

'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'   x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
' Next
'
' AddPt "Polarity", 0, 0, 0
' AddPt "Polarity", 1, 0.05, -5
' AddPt "Polarity", 2, 0.4, -5
' AddPt "Polarity", 3, 0.6, -4.5
' AddPt "Polarity", 4, 0.65, -4.0
' AddPt "Polarity", 5, 0.7, -3.5
' AddPt "Polarity", 6, 0.75, -3.0
' AddPt "Polarity", 7, 0.8, -2.5
' AddPt "Polarity", 8, 0.85, -2.0
' AddPt "Polarity", 9, 0.9,-1.5
' AddPt "Polarity", 10, 0.95, -1.0
' AddPt "Polarity", 11, 1, -0.5
' AddPt "Polarity", 12, 1.1, 0
' AddPt "Polarity", 13, 1.3, 0
'
' addpt "Velocity", 0, 0,         1
' addpt "Velocity", 1, 0.16, 1.06
' addpt "Velocity", 2, 0.41,         1.05
' addpt "Velocity", 3, 0.53,         1'0.982
' addpt "Velocity", 4, 0.702, 0.968
' addpt "Velocity", 5, 0.95,  0.968
' addpt "Velocity", 6, 1.03,         0.945
'
' LF.Object = LeftFlipper
' LF.EndPoint = EndPointLp
' RF.Object = RightFlipper
' RF.EndPoint = EndPointRp
'End Sub



'
''*******************************************
'' Early 90's and after
'
Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
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

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub


' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub




'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
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

  Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
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
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

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
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset)        'Resize original array
  for x = 0 to aCount-1                'set objects back into original array
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
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
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
  dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

  LinearEnvelope = Y
End Function


'******************************************************
'  FLIPPER TRICKS
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
  Dim b, BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
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
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
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

'*************************************************
'  Check ball distance from Flipper for Rem
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
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

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
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
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
  Flipper.eostorque = EOST*EOSReturn/FReturn


  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, BOT
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
  Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

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
  Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
    end If

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
  End If
End Sub



'************************************************************************
'            AUTOPLUNGER
'************************************************************************

Sub Auto_Plunger(Enabled)
    If Enabled Then
       Plunger.Fire
       PlaySoundAt SoundFX("Popper",DOFContactors), Plunger
  End If
End Sub


'******************************************************
'           TROUGH
'******************************************************

Sub sw34_Hit():Controller.Switch(34) = 1:UpdateTrough:End Sub
Sub sw34_UnHit():Controller.Switch(34) = 0:UpdateTrough:End Sub
Sub sw33_Hit():Controller.Switch(33) = 1:UpdateTrough:End Sub
Sub sw33_UnHit():Controller.Switch(33) = 0:UpdateTrough:End Sub
Sub sw32_Hit():Controller.Switch(32) = 1:UpdateTrough:End Sub
Sub sw32_UnHit():Controller.Switch(32) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw32.BallCntOver = 0 Then sw33.kick 60, 9
  If sw33.BallCntOver = 0 Then sw34.kick 60, 9
  If sw34.BallCntOver = 0 Then sw35.kick 60, 20
  Me.Enabled = 0
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub sw35_Hit() 'Drain
  UpdateTrough
  Controller.Switch(35) = 1
  RandomSoundDrain sw35
End Sub

Sub sw35_UnHit()  'Drain
  Controller.Switch(35) = 0
End Sub

Sub SolBallRelease(enabled)
  If enabled Then
    If sw32.BallCntOver = 0 Then
      PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw32
    Else
      RandomSoundBallRelease sw32
      vpmTimer.PulseSw 31
    End If
    sw32.kick 60, 9
  End If
End Sub


'************************************************************************
'       CASTLE TOWERS
'************************************************************************

Dim vel,per,brake,cnt4on,cnt4off
vel=0:per=0:brake=0

Sub SolCastle(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("fx_AutoPlunger",DOFContactors), Ctower
    brake=0
    cnt4off=0
    explosion.enabled=1
  Else
    cnt4on=0
    StopSound "fx_AutoPlunger"
  End If
End Sub

Sub explosion_timer()
  If Controller.Solenoid(4) then
    cnt4on=cnt4on+1
  End If

  If NOT Controller.Solenoid(4) then
    cnt4off=cnt4off+1
  End If

  If Controller.Solenoid(4) and cnt4on>15 then
    LUtower.roty=(-20)*0.75
    LDtower.roty=(-70)*0.75
    URtower.roty=(40)*0.75
    URtower.rotx=(-40)*0.75
    Ctower.roty=(-20)*0.75
    Ctower.rotx=(-20)*0.75
    brake=0.09
    vel=0.9
  End If

  If brake<1 then
    vel=vel+0.1:brake=brake+0.01
  End If

  If LUtower.roty>0 then
    per=0:vel=0
  End If

  per=(cos(vel)-brake)
  LUtower.roty=LUtower.roty-(per*2)
  LDtower.roty=LDtower.roty-(per*7)
  URtower.roty=URtower.roty+(per*4)
  URtower.rotx=URtower.rotx-(per*4)
  Ctower.roty=Ctower.roty-(per*2)
  Ctower.rotx=Ctower.rotx-(per*2)

  If NOT Controller.Solenoid(4) AND cnt4off>100 Then
    LUtower.roty=0
    LDtower.roty=0
    URtower.roty=0
    URtower.rotx=0
    Ctower.roty=0
    Ctower.rotx=0
    Me.Enabled=0
  End If

  if LUtower.roty < -3 Then
    CastleDown = 1
    FlasherVisibility
    Flasher19a.visible = True
    Flasher14a.visible = True
    Flasher19.visible = False
    Flasher14.visible = False
  Else
    CastleDown = 0
    FlasherVisibility
    Flasher19a.visible = False
    Flasher14a.visible = False
    Flasher19.visible = True
    Flasher14.visible = True
  End If
End Sub

'************************************************************************
'            CASTLE GATE
'************************************************************************

Dim IronGateDir

Sub SolCastlegatePow(Enabled)
  If Enabled Then
    CastleGateTimer.Enabled = 1
    IronGateDir = 1
    PlaySoundat SoundFX("fx_gateUp",DOFContactors), gate
    TwrShake
  End If
End Sub

Sub SolCastlegateHold(Enabled)
  If Enabled Then
    DOF 102, DOFPulse
  End If
  If NOT Enabled AND IronGateDir = 1 Then
    IronGateDir = -1
    CastleGateTimer.Enabled = 1
    PlaySoundat SoundFX("fx_gateDown",DOFContactors), gate
    DOF 102, DOFPulse
    End If
End Sub

Sub CastleGateTimer_Timer
  DOF 101, DOFOn
  gate.Z = gate.Z + IronGateDir
        If gate.Z <= 15 Then
      door2.IsDropped = 0
      DOF 101, DOFOff
            Me.Enabled = 0
        End If
        If gate.Z >= 115 Then
      door2.IsDropped = 1
      DOF 101, DOFOff
            Me.Enabled = 0
        End If
End Sub

'************************************************************************
'             Catapult
'************************************************************************

Dim catdir, CatBall: CatBall = Empty

Sub sw38_hit
  PlaySoundAtBall "VukEnter"
  Controller.switch(38)=1
  Set CatBall = Activeball
End Sub

Sub SolCatapult(enabled)
  If enabled Then
    catdir = 1
    sw38.TimerInterval = 1
    sw38.TimerEnabled = 1
    If sw38.BallCntOver = 0 Then
      PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw38
    Else
      PlaySoundAt SoundFX("KickandWire",DOFContactors), sw38
    End If
  End If
End Sub


Sub sw38_timer()
  Prim_Catapult.Rotx = Prim_Catapult.Rotx + catdir

  Dim catradius: catradius = 110

  If Not IsEmpty(CatBall) Then
    CatBall.x = Prim_Catapult.x + catradius*cos(radians(90))*cos(Radians(Prim_Catapult.Rotx-7))
    CatBall.y = Prim_Catapult.y + catradius*sin(radians(90))*cos(Radians(Prim_Catapult.Rotx-7))
    CatBall.z = 25 + catradius*sin(radians(Prim_Catapult.Rotx-7))
  End If

  If Prim_Catapult.Rotx >= 97 AND catdir=1 Then
    catdir = -1
    sw38.kick 0, 45
    CatBall = Empty
    controller.Switch(38) = 0
  End If
  If Prim_Catapult.Rotx <= 7 Then
    Me.TimerEnabled = 0
  End If
End Sub

'************************************************************************
'            TOWER DIVERTER
'************************************************************************

Dim DiverterDir

Sub SolTowerDivPow(Enabled)
  If Enabled Then
    Diverter.IsDropped=1
    Diverter.TimerEnabled = 1
    DiverterDir = 1
    PlaySoundAt SoundFX("fx_DiverterUp",DOFContactors), DiverterP
  End If
End Sub

Sub SolTowerDivHold(Enabled)
  If NOT Enabled AND DiverterDir = 1 Then
    Diverter.IsDropped=0
    DiverterDir = -1
    Diverter.TimerEnabled = 1
    PlaySoundat SoundFX("fx_DiverterDown",DOFContactors), Diverterp
    End If
End Sub

Sub Diverter_Timer
  DiverterP.Z = DiverterP.Z - 5*DiverterDir
        If DiverterP.Z <= 87 Then Me.TimerEnabled = 0
        If DiverterP.Z >= 137 Then Me.TimerEnabled = 0
End Sub

'************************************************************************
'            DRAW BRIDGE
'************************************************************************

Dim dbpos
' Keeps track of the bridge position.
Dim drawBridgeIsClosed
drawBridgeIsClosed = True

Sub SolDrawBridge(enabled)
  If Enabled AND drawBridgeIsClosed Then
    dbpos = 1
    dbridge.enabled = 1
    PlaySound SoundFX("Bridge_Move", DOFGear), -1, 0.1, AudioPan(braket) , 0, 0, 1, AudioPan(braket)
    DOF 104, DOFOn
  end If
  If Enabled AND NOT drawBridgeIsClosed Then
    dbpos = 2
    dbridge.enabled = 1
    PlaySound SoundFX("Bridge_Move", DOFGear), -1, 0.1, AudioPan(braket) , 0, 0, 1, AudioPan(braket)
    DOF 104, DOFOn
  End If
End Sub

Sub dbridge_timer()
  Select Case dbpos
    Case 1:     'bridge is going down
      drawbridgep.RotX = drawbridgep.Rotx - 1
      DBdecal.rotx=DBdecal.rotx-1
      braket.rotx=braket.rotx-1
      If drawbridgep.RotX <= -90 Then
        DOF 104, DOFOff
        drawbridgep.RotX= -90
        DBdecal.rotx=-90
        braket.rotx=-90
        Door1.isdropped = 1
        BridgeRamp.collidable = 1
        Ramp22.collidable = 0
        Ramp002.collidable = 0
        BW1.isdropped = 0
        BW2.isdropped = 0
        Me.Enabled = 0
        drawBridgeIsClosed = False
        StopSound "Bridge_Move"
        PlaySound SoundFX("Bridge_Stop", 0),0, 0.1, AudioPan(braket) , 0, 0, 1, AudioPan(braket)
      End If

    Case 2:     'bridge is going up
      drawbridgep.RotX = drawbridgep.Rotx + 1
      DBdecal.rotx=DBdecal.rotx+1
      braket.rotx=braket.rotx+1
      If drawbridgep.RotX >= 0 Then
        DOF 104, DOFOff
        drawbridgep.RotX= 0
        DBdecal.rotx=0
        braket.rotx=0
        Door1.isdropped = 0
        BridgeRamp.collidable = 0
        Ramp22.collidable = 1
        Ramp002.collidable = 1
        BW1.isdropped = 1
        BW2.isdropped = 1
        Me.Enabled = 0
        drawBridgeIsClosed = True
        StopSound "Bridge_Move"
        PlaySound SoundFX("Bridge_Stop", 0),0, 0.1, AudioPan(braket) , 0, 0, 1, AudioPan(braket)
      End If
  End Select
End Sub


'************************************************************************
'       Drawbridge and Castle Door Shake
'************************************************************************

Sub Door1_Hit()
  RandomSoundMetal
  If Controller.Switch(56) Then 'solo se sta su
    DOF 102, DOFPulse
    doorshake.Enabled = 1
    TwrShake
  End If
End Sub

Sub Wall108_hit
  TwrShake
End Sub


dim doors,doorsh,doorbrake

sub doorshake_timer()
  if doorbrake<3 then
    doorbrake=doorbrake+0.1
    doors=doors+0.5
    doorsh=sin(doors)*(3-(doorbrake))
    if gateshake=0 then drawbridgep.RotX = drawbridgep.Rotx +doorsh
    if gateshake=0 then DBdecal.rotx=DBdecal.rotx+doorsh
    if gateshake=1 then gate.transz=doorsh
   end if
'   braket.rotx=braket.rotx+50
  if doorbrake>3  then
    if gateshake=0 then drawbridgep.RotX = 0
    if gateshake=0 then DBdecal.rotx=0
    if gateshake=1 then gate.transz=0
    doors=0:doorsh=0:doorbrake=0:gateshake=0:Me.Enabled=0 :exit sub
  end if
end sub

dim gateshake

Sub Door2_Hit()
  gateshake=1
  RandomSoundMetal
  doorshake.enabled=1
  DOF 102, DOFPulse
  vpmTimer.PulseSw 37
  TwrShake
End Sub

'************************************************************************
'           TOWERS SHAKE
'************************************************************************

dim vel2,per2,brake2
per2=0:vel2=0:brake2=0

Sub TwrShake
    towersshake.enabled = 1
  If brakew > 4 or GateT.timerenabled = false Then
    rotat=0
    brakew=4
    GateT.TimerEnabled=1
  End If
End Sub

Sub towersshake_timer()

        if brake2 < 1 then vel2=vel2+0.1:brake2=brake2+0.01
        if LUtower.roty>0 then per2=0:vel2=0
        per2=cos(vel2)-brake2

  LUtower.roty=LUtower.roty-(per2*0.2)
  LDtower.roty=LDtower.roty-(per2*0.2)
  URtower.roty=URtower.roty+(per2*0.2)
  URtower.rotx=URtower.rotx-(per2*0.2)
  Ctower.roty=Ctower.roty-(per2*0.2)
  Ctower.rotx=Ctower.rotx-(per2*0.2)

   if brake2>0.9 then me.enabled=0 : per2=0:vel2=0:brake2=0:LUtower.roty=0 :LDtower.roty=0 : URtower.roty=0: Ctower.roty=0:Ctower.rotx=0 : URtower.rotx=0

End Sub

'************************************************************************
'               Tower Lock
'************************************************************************

Sub SolTowerLock(Enabled)
  StopSound "fx_Postup":PlaySoundAT SoundFX("fx_Postup",DOFContactors), sw58
  LockPost.IsDropped=NOT Enabled
  LockPostP.IsDropped=NOT Enabled
End Sub

Sub sw58_Hit: Controller.Switch(58)=1:End Sub
Sub sw58_UnHit: Controller.Switch(58)=0:End Sub

'************************************************************************
'               Merlin Saucer
'************************************************************************

Sub sw28_hit
  PlaySoundAt "VUKEnter" ,sw28
  Controller.switch(28)=1
End Sub

Sub SolMerlin(enabled)
  controller.Switch(28) = 0
  If sw28.BallCntOver = 0 Then
    PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw28
  Else
    PlaySoundAt SoundFX("VUKOut",DOFContactors), sw28
  End If
  sw28.kick 109 + (rnd*2), 14 + (rnd*2)
End Sub

'************************************************************************
'               Moat Scoop
'************************************************************************

Sub sw36_hit
  MoatKick.collidable = True
  PlaySoundAtBall "Ball_Bounce_Playfield_Hard_2"
  Controller.switch(36)=1
End Sub

Sub SolMoat(enabled)
  controller.Switch(36) = 0
  If sw36.BallCntOver = 0 Then
    PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw36
  Else
    PlaySoundAt SoundFX("Popper",DOFContactors), sw36
  End If
  sw36.kickz 215 , 15, 0, 140
  MoatKick.collidable=false
End Sub


'************************************************************************
'         Slingshots Animation
'************************************************************************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  RandomSoundSlingshotLeft sling1
  vpmTimer.PulseSw 51
  LSling.Visible = 0
  LSling1.Visible = 1
  sling1.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
  RandomSoundSlingshotRight sling2
  vpmTimer.PulseSw 52
  RSling.Visible = 0
  RSling1.Visible = 1
  sling2.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'************************************************************************
'           Bumpers Animation
'************************************************************************

Dim dirRing1:dirRing1 = -1
Dim dirRing2:dirRing2 = -1
Dim dirRing3:dirRing3 = -1

Sub Bumper1_Hit:vpmTimer.PulseSw 53:RandomSoundBumperTop Bumper1:Me.TimerEnabled = 1:me.timerinterval = 10:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 54:RandomSoundBumperMiddle Bumper2:Me.TimerEnabled = 1:me.timerinterval = 10:End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 55:RandomSoundBumperBottom Bumper3:Me.TimerEnabled = 1:me.timerinterval = 10:End Sub

Sub Bumper1_timer()
  BumperRing1.Z = BumperRing1.Z + (5 * dirRing1)
  If BumperRing1.Z <= -35 Then dirRing1 = 1
  If BumperRing1.Z >= 0 Then
    dirRing1 = -1
    BumperRing1.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  BumperRing2.Z = BumperRing2.Z + (5 * dirRing2)
  If BumperRing2.Z <= -35 Then dirRing2 = 1
  If BumperRing2.Z >= 0 Then
    dirRing2 = -1
    BumperRing2.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper3_timer()
  BumperRing3.Z = BumperRing3.Z + (5 * dirRing3)
  If BumperRing3.Z <= -35 Then dirRing3 = 1
  If BumperRing3.Z >= 0 Then
    dirRing3 = -1
    BumperRing3.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

'************************************************************************
'               Trolls
'************************************************************************

Dim LTDir, RTDir, TrollStep
dim lshake,rshake
lshake=0:rshake=0

If TrollSpeed = 0 Then
  TrollStep = 12
ElseIf TrollSpeed = 2 Then
  TrollStep = 24
ElseIf TrollSpeed = 3 Then
  TrollStep = 32
Else
  TrollStep = 16
End If

Sub SolLeftTrollPow(Enabled)
  If Enabled then
    PlaySoundAt SoundFX("fx_TrollUp",DOFContactors), ltroll
    TrollLt.enabled = false
    LTDir = 1
    TrollLt.enabled = true
  End If
  If NOT Enabled AND ltroll.Z = -8 Then
    TrollP1X.TimerEnabled=1
    TrollP1X.TimerInterval=10
  End If
End Sub

Sub SolLeftTrollHold(Enabled)
  If NOT Enabled AND ltroll.Z > -104 Then
    TrollLt.enabled = false
    LTDir = -1
    TrollLt.enabled = true
    PlaySoundAt SoundFX("fx_TrollDown",DOFContactors), ltroll
  End If
End Sub


Sub TrollLt_timer()
  ltroll.z = ltroll.z + (TrollStep * LTDir)
  liftL.Z = ltroll.Z
  if ltroll.z > -81 then
    TrollP1X.IsDropped = 0
    sw45.IsDropped = 0
    LTT.collidable = True
    Controller.Switch(74) = 1
  Else
    TrollP1X.IsDropped = 1
    sw45.IsDropped = 1
    Controller.Switch(74) = 0
    LTT.collidable = false
  end If

  Dim BOT, b
  BOT = GetBalls

  For b = 0 to UBound(BOT)
    If InRect(BOT(b).x, BOT(b).y, 291,870,373,846,395,956,312,973) and LTDir = 1 and BOT(b).z < 121 Then
      BOT(b).z = 25 + ltroll.z + 104
      BOT(b).velz = 15
    End If
  Next

  If ltroll.z >= -8 Then
    ltroll.z = -8
    liftL.Z = ltroll.Z
    me.enabled = false
  Elseif ltroll.z <= -104 then
    ltroll.z = -104
    liftL.Z = ltroll.Z
    me.enabled = false
  End If
End Sub


Sub TrollP1X_timer
  lshake=lshake+1
  If lshake<12 then
    LiftL.transz=1 + Sin(lshake)
    ltroll.transz=1 + Sin(lshake)
  Else
    me.TimerEnabled=0
    lshake=0
    LiftL.transz=0
  End If
End Sub


Sub SolRightTrollPow(Enabled)
  If Enabled then
    PlaySoundAt SoundFX("fx_TrollUp",DOFContactors), rtroll
    TrollRt.enabled = false
    RTDir = 1
    TrollRt.enabled = true
  End If
  If NOT Enabled AND rtroll.Z = -8 Then
    TrollP2X.TimerEnabled=1
    TrollP2X.TimerInterval=10
  End If
End Sub

Sub SolRightTrollHold(Enabled)
  If NOT Enabled AND rtroll.Z > -104 Then
    TrollRt.enabled = false
    RTDir = -1
    TrollRt.enabled = true
    PlaySoundAt SoundFX("fx_TrollDown",DOFContactors), rtroll
  End If
End Sub

Sub TrollRt_timer()
  rtroll.z = rtroll.z + (TrollStep * RTDir)
  liftR.Z = rtroll.Z
  if rtroll.z > -81 then
    TrollP2X.IsDropped = 0
    sw46.IsDropped = 0
    RTT.collidable = True
    Controller.Switch(75) = 1
  Else
    TrollP2X.IsDropped = 1
    sw46.IsDropped = 1
    Controller.Switch(75) = 0
    RTT.collidable = false
  end If

  Dim BOT, b
  BOT = GetBalls

  For b = 0 to UBound(BOT)
    If InRect(BOT(b).x, BOT(b).y, 487,857,571,864,552,968,470,957) and RTDir = 1 Then
      BOT(b).z = 25 + rtroll.z + 104
      BOT(b).velz = 15
    End If
  Next

  If rtroll.z >= -8 Then
    rtroll.z = -8
    liftR.Z = rtroll.Z
    me.enabled = false
  Elseif rtroll.z <= -104 then
    rtroll.z = -104
    liftR.Z = rtroll.Z
    me.enabled = false
  End If
End Sub

Sub TrollP2X_timer
  rshake=rshake+1
  if rshake<12 then
    LiftR.transz=1 + Sin(rshake)
    rtroll.transz=1 + Sin(rshake)
  Else
    me.TimerEnabled=0
    rshake=0
    LiftR.transz=0
  End If
End Sub

'Shake Trolls when hit

dim sbou,sbou2

Sub sw45_Hit():vpmTimer.PulseSw 45:Me.TimerEnabled = 1:sbou=ActiveBall.vely/4 :End Sub
Sub sw46_Hit():vpmTimer.PulseSw 46:Me.TimerEnabled = 1:sbou2=ActiveBall.vely/4 :End Sub

dim bou,brakeTl,perc
perc=1
Sub sw45_Timer()

    bou=bou+0.3:brakeTl=brakeTl+0.2
    if (perc-(brakeTl*(perc/6)))<0 then Me.TimerEnabled = 0 :bou=0 :brakeTl=0 :perc=5
    ltroll.rotx=sin(bou)*(perc-(brakeTl*(perc/6)))
    ltroll.roty=sin((bou)*sbou)*(perc-(brakeTl*(perc/6)))

    LiftL.rotx=sin(bou)*(perc-(brakeTl*(perc/6)))
    LiftL.roty=sin((bou)*sbou)*(perc-(brakeTl*(perc/6)))

end sub

dim bou2,brakeTr,perc2
perc2=1
Sub sw46_Timer()

    bou2=bou2+0.3:brakeTr=brakeTr+0.2
    if (perc2-(brakeTr*(perc2/6)))<0 then Me.TimerEnabled = 0 :bou2=0 :brakeTr=0 :perc2=5
    rtroll.rotx=sin(bou2)*(perc2-(brakeTr*(perc2/10)))
    rtroll.roty=sin((bou2)*sbou2)*(perc2-(brakeTr*(perc2/6)))

    liftR.rotx=sin(bou2)*(perc2-(brakeTr*(perc2/10)))
    liftR.roty=sin((bou2)*sbou2)*(perc2-(brakeTr*(perc2/6)))

end sub


'************************************************************************
'         Switches
'************************************************************************
' Lanes
Sub sw66_Hit:Controller.Switch(66) = 1:End Sub
Sub sw66_UnHit:Controller.Switch(66) = 0:End Sub

Sub sw67_Hit:Controller.Switch(67) = 1:End Sub
Sub sw67_UnHit:Controller.Switch(67) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw26_Hit
Controller.Switch(26) = 1
If ActiveBall.VelY => 3.5 Then
ActiveBall.VelY = 3.5
End If
If ActiveBall.VelY <= -10 Then
ActiveBall.VelY = -6
End If
End Sub

Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw17_Hit
Controller.Switch(17) = 1
If ActiveBall.VelY => 3.5 Then
ActiveBall.VelY = 3.5
End If
If ActiveBall.VelY <= -10 Then
ActiveBall.VelY = -6
End If
End Sub

Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:End Sub
Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw68_Hit:Controller.Switch(68) = 1:End Sub
Sub sw68_UnHit:Controller.Switch(68) = 0:End Sub

' Targets
Sub sw12_Hit:vpmTimer.PulseSw 12:End Sub
Sub sw15_Hit:vpmTimer.PulseSw 15:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:End Sub
Sub sw71_Hit:vpmTimer.PulseSw 71:End Sub
Sub sw72_Hit:vpmTimer.PulseSw 72:End Sub
Sub sw73_Hit:vpmTimer.PulseSw 73:End Sub

' Triggers on the ramps
Sub sw61_Hit:Controller.Switch(61) = 1:End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:sw62flip.rotatetoend:End Sub
Sub sw62_Unhit:Controller.Switch(62) = 0:sw62flip.rotatetostart:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:End Sub
Sub sw63_Unhit:Controller.Switch(63) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:sw64flip.rotatetoend:End Sub
Sub sw64_Unhit:Controller.Switch(64) = 0:sw64flip.rotatetostart:End Sub

' Opto Switches
Sub sw41_hit
  vpmtimer.pulsesw 41
End sub

Sub sw37_Hit()
  vpmTimer.PulseSw 37
End Sub

' Lock_Door
dim rotat,brakew

Sub GateT_Hit
  PlaySoundatBall "Gate1"
  rotat=0
  brakew=1
  me.TimerEnabled=0
  me.TimerEnabled=1
  me.TimerInterval=7
  if activeball.vely < -20 then TwrShake
End Sub

Sub GateT_timer()
  If brakew <7 then rotat=rotat+0.07:brakew=brakew+0.003 else Lock_Door.rotX=0:me.TimerEnabled = 0
  Lock_Door.rotX=(cos(rotat+90)*100)/brakew^3
  If Lock_Door.rotX > 20 or Lock_Door.rotX < -20 Then
    LockGateOpen = 1
  Else
    LockGateOpen = 0
  End If
End Sub

Sub sw44_Hit():vpmtimer.PulseSw 44:PlaySoundatBall "Gate1": End Sub

'************************************************************************
'           Modulated Flashers
'************************************************************************

Dim xxst, xxet, xxnt, xxtw, xxto, xxtt, xxt3, xxmf, xxtf

Sub SolMod17(value)
  If value > 0 Then
    For each xxst in BL_Flash:xxst.State = 1:next
  Else
    For each xxst in BL_Flash:xxst.State = 0:next
  End If

  For each xxst in BL_Flash:xxst.IntensityScale = value/160:next

  F_lt1.color=RGB(255,0,0)      'troll reflection
  F_lt2.color=RGB(255,0,0)      'troll reflection
  f17b.IntensityScale=value/160   'ambient reflection
  f17c.IntensityScale=value/160   'dome lit
  f17d.IntensityScale=value/160   'dome lit
  f23d005.IntensityScale=value/160  'wall reflection

  If ltroll.Z > -50 Then
    F_lt1.IntensityScale=value/160
    F_lt2.IntensityScale=value/160
  Else
    F_lt1.IntensityScale=0
    F_lt2.IntensityScale=0
  End If
End Sub

Sub SolMod18(value)
  If value > 0 Then
    For each xxet in LR_Flash:xxet.State = 1:next
  Else
    For each xxet in LR_Flash:xxet.State = 0:next
  End If

  For each xxet in LR_Flash:xxet.IntensityScale = value/160:next

  F_lt1.color=RGB(255,200,145)      'troll reflection
  F_lt2.color=RGB(255,200,145)      'troll reflection

  f18.intensityscale=value/160      'castle reflection, castle up, gate closed
' f18b.intensityscale=value/160     'castle reflection, castle up, gate open
  f18c.intensityscale=value/160     'castle reflection, castle down, gate closed
' f18d.intensityscale=value/160     'castle reflection. cast;e down, gate open

  If ltroll.Z > -50 Then
    F_lt1.IntensityScale=value/160
    F_lt2.IntensityScale=value/160
  Else
    F_lt1.IntensityScale=0
    F_lt2.IntensityScale=0
  End If
End Sub

Sub SolMod19(value)
  If value > 0 Then
    For each xxnt in TL_Flash:xxnt.State = 1:next
  Else
    For each xxnt in TL_Flash:xxnt.State = 0:next
  End If

  For each xxnt in TL_Flash:xxnt.IntensityScale = value/160:next
  For each xxnt in Castle_LF:xxnt.IntensityScale = value/160:next
End Sub

Sub SolMod20(value)
  If value > 0 Then
    For each xxtw in TR_Flash:xxtw.State = 1:next
  Else
    For each xxtw in TR_Flash:xxtw.State = 0:next
  End If

  For each xxtw in TR_Flash:xxtw.IntensityScale = value/160:next
  For each xxtw in Castle_RF:xxtw.IntensityScale = value/160:next
End Sub

Sub SolMod21(value)
  If value > 0 Then
    For each xxto in RR_Flash:xxto.State = 1:next
  Else
    For each xxto in RR_Flash:xxto.State = 0:next
  End If

  For each xxto in RR_Flash:xxto.IntensityScale = value/160:next

  F_rt1.color=RGB(255,200,145)      'troll reflection
  F_rt2.color=RGB(255,200,145)      'troll reflection
  f21.intensityscale=value/160      'castle reflection
  f21a.intensityscale=value/160     'merlin decal
  f21c.intensityscale=value/160     'dragon
  FLASHER001.intensityscale=value/160     'wall reflection

  If rtroll.Z > -50 Then
    F_rt1.IntensityScale=value/160
    F_rt2.IntensityScale=value/160
  Else
    F_rt1.IntensityScale=0
    F_rt2.IntensityScale=0
  End If
End Sub

Sub SolMod22(value)
  If value > 0 Then
    For each xxtt in CR_Flash:xxtt.State = 1:next
  Else
    For each xxtt in CR_Flash:xxtt.State = 0:next
  End If

  For each xxtt in CR_Flash:xxtt.IntensityScale = value/160:next

  F22.IntensityScale=value/160      'castle reflection
  F22a.IntensityScale=value/160     'dome lit
  F22b.IntensityScale=value/160     'dome lit
  f17d001.IntensityScale=value/160    'dome lit
End Sub

Sub SolMod23(value)
  If value > 0 Then
    For each xxt3 in BR_Flash:xxt3.State = 1:next
  Else
    For each xxt3 in BR_Flash:xxt3.State = 0:next
  End If

  For each xxt3 in BR_Flash:xxt3.IntensityScale=value/160:next

  F_rt1.color=RGB(255,0,0)        'troll reflection
  F_rt2.color=RGB(255,0,0)        'troll reflection
  f23.IntensityScale=value/160      'dragon
  f23a.IntensityScale=value/160     'dome lit
  f23b.IntensityScale=value/160     'dome lit
  f23c.IntensityScale=value/160     'dome lit
  f23d.IntensityScale=value/160     'wall reflection

  If rtroll.Z > -50 Then
    F_rt1.IntensityScale=value/160
    F_rt2.IntensityScale=value/160
  Else
    F_rt1.IntensityScale=0
    F_rt2.IntensityScale=0
  End If
End Sub

Sub SolMod24(value)
  If value > 0 Then
    For each xxmf in Moat_Flash:xxmf.State = 1:next
  Else
    For each xxmf in Moat_Flash:xxmf.State = 0:next
  End If

  For each xxmf in Moat_Flash:xxmf.IntensityScale=value/160:next

  F24.IntensityScale=value/160      'castle reflection
End Sub

Sub SolMod25(value)
  If value > 0 Then
    For each xxtf in CL_Flash:xxtf.State = 1:next
  Else
    For each xxtf in CL_Flash:xxtf.State = 0:next
  End If

  For each xxtf in CL_Flash:xxtf.IntensityScale=value/160:next

  F25.IntensityScale=value/160      'castle reflection
  F25a.IntensityScale=value/160     'dome lit
  F25b.IntensityScale=value/160     'dome lit
End Sub


Dim CastleDown, LockGateOpen
CastleDown = 0
LockGateOpen = 0

Sub FlasherVisibility()
  if CastleDown = 1 Then
    Flasher19a.visible = True
    Flasher14a.visible = True
    Flasher19.visible = False
    Flasher14.visible = False
    If LockGateOpen = 1 Then
      F18.visible = False
      'F18b.visible = False
      F18c.visible = False
      'F18d.visible = True
    Else
      F18.visible = False
      'F18b.visible = False
      F18c.visible = True
      'F18d.visible = False
    End If
  ElseIf CastleDown = 0 Then
    Flasher19a.visible = False
    Flasher14a.visible = False
    Flasher19.visible = True
    Flasher14.visible = True
    If LockGateOpen = 1 Then
      F18.visible = False
      'F18b.visible = True
      F18c.visible = False
      'F18d.visible = False
    Else
      F18.visible = True
      'F18b.visible = False
      F18c.visible = False
      'F18d.visible = False
    End If
  End If

End Sub

'**************
' Inserts
'**************

Sub InitLamps
On Error Resume Next
Dim i
For i=0 To 127
    If IsObject(eval("L" & i)) Then
        Execute "Set Lights(" & i & ")  = L" & i
    End If
Next
Lights(14)=Array(L14,L14a)
Lights(64)=Array(L64,L64a)
Lights(78)=Array(L78,L78a)
End Sub

Sub SynchFlasherObj
L11a.IntensityScale=L11.state
L12a.IntensityScale=L12.state
L13a.IntensityScale=L13.state
L15a.IntensityScale=L15.state
L31a.IntensityScale=L31.state
L55a.IntensityScale=L55.state
l53a.IntensityScale=light53.state
L56a.IntensityScale=L56.state
L57a.IntensityScale=L57.state
L58a.IntensityScale=L58.state
l31a001.IntensityScale=L78a.state
l31a002.IntensityScale=L78.state
L81a.IntensityScale=L81.state
L83a.IntensityScale=L84.state
L84a.IntensityScale=L84.state
l81a001.IntensityScale=light63.IntensityScale
l81a002.IntensityScale=light60.IntensityScale
l81a003.IntensityScale=light70.IntensityScale
l81a004.IntensityScale=light70.IntensityScale
l81a005.IntensityScale=light70.IntensityScale
l81a006.IntensityScale=light70.IntensityScale
l81a007.IntensityScale=light70.IntensityScale
l81a008.IntensityScale=light70.IntensityScale
l81a009.IntensityScale=light70.IntensityScale
Lower_bumber_ref.IntensityScale=L62a3.IntensityScale
upper_bumber_ref.IntensityScale=L62a1.IntensityScale
Bumber_castle_reflection.IntensityScale=L62a1.IntensityScale
f21d.IntensityScale=L62a1.state
End Sub

'**************
' 8-step GI
'**************

Set GiCallback2 = GetRef("UpdateGI")

Dim gistep,xx
gistep = 1/8

Sub UpdateGI(no, step)
Select Case no
Case 0 'bottom
  If step = 0 Then
    For each xx in GIB:xx.State = 0:Next
  Else
    For each xx in GIB:xx.State = 1:Next
  End If
  For each xx in GIB:xx.IntensityScale = gistep * step:next
Case 1 'middle
  If step = 0 Then
    DOF 103, DOFOff
    For each xx in GIM:xx.State = 0:Next
  Else
    DOF 103, DOFOn
    For each xx in GIM:xx.State = 1:Next
  End If
  For each xx in GIM:xx.IntensityScale = gistep * step:next
Case 2 'top
  If step = 0 Then
    For each xx in GIT:xx.State = 0:Next
    For each xx in bump1:xx.State = 0:Next
    For each xx in bump2:xx.State = 0:Next
    For each xx in bump3:xx.State = 0:Next
  Else
    For each xx in GIT:xx.State = 1:Next
    For each xx in bump1:xx.State = 1:Next
    For each xx in bump2:xx.State = 1:Next
    For each xx in bump3:xx.State = 1:Next
  End If
  If step>4 then
    Prim_Spot1.image= "spot_map (black version)on"
    Prim_Spot2.image= "spot_map (black version)on"
  Else
    Prim_Spot1.image= "spot_map (black version)"
    Prim_Spot2.image= "spot_map (black version)"
  End If
For each xx in CastleGI:xx.IntensityScale = gistep * step:next
For each xx in GIT:xx.IntensityScale = gistep * step:next
For each xx in bump1:xx.IntensityScale = gistep * step:next
For each xx in bump2:xx.IntensityScale = gistep * step:next
For each xx in bump3:xx.IntensityScale = gistep * step:next
End Select
End Sub

'******************
' RealTime Updates
'******************

' Set MotorCallback = GetRef("GameTimer")

Sub GameTimer_timer()
  cor.update
  UpdateMechs
  RollingUpdate         'update rolling sounds
  BallShadowUpdate
  SynchFlasherObj
  UpdateMechs

End Sub

Sub FrameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
End Sub

Sub UpdateMechs
  Gate2P.RotX=Gate2.currentangle
  Gate3P.RotX=Gate3.currentangle
  WireGateRSmall.RotX=Gate9.currentangle
  WireGateR.RotX=Spinner1.currentangle
  WireGateLR.RotX=Spinner2.currentangle
  WireGateRR.RotX=Spinner3.currentangle
  WireGateMerlin.RotX=Gate6.currentangle
  WireGateCatapult.RotY=Gate1.currentangle
  sw62p.rotY = sw62flip.currentangle
  sw64p.rotY = sw64flip.currentangle
  FlipperL.RotZ = LeftFlipper.CurrentAngle
  FlipperR.RotZ = RightFlipper.CurrentAngle
End Sub

'*****************************************
' Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)


Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls

  ' render the shadow for each ball
    For b = 0 to UBound(BOT)
    BallShadow(b).X = BOT(b).X
    ballShadow(b).Y = BOT(b).Y + 10

    If BOT(b).Z > 24 and BOT(b).Z < 35 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If

'   If BOT(b).X < Table1.Width/2 Then
'     BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
'   Else
'     BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
'   End If

  Next
End Sub

'**********************OPTIONS***************************

Sub InitOptions
Ramp15.visible = 0
Ramp16.visible = 0
End Sub

'//////////////////////////////////////////////////////////////////////
'// TargetBounce
'//////////////////////////////////////////////////////////////////////

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

' Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit
  TargetBouncer activeball, 1
End Sub

'//////////////////////////////////////////////////////////////////////
'// PHYSICS DAMPENERS
'//////////////////////////////////////////////////////////////////////

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


  Public Sub Report()         'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub

End Class

'//////////////////////////////////////////////////////////////////////
'// TRACK ALL BALL VELOCITIES FOR RUBBER DAMPENER AND DROP TARGETS
'//////////////////////////////////////////////////////////////////////

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

'//////////////////////////////////////////////////////////////////////
'// RAMP ROLLING SFX
'//////////////////////////////////////////////////////////////////////

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
dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)

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

'//////////////////////////////////////////////////////////////////////
'// RAMP TRIGGERS
'//////////////////////////////////////////////////////////////////////

Sub ramptrigger003_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger002_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger01_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger02_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger02_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger03_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger03_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger03
End Sub

Sub sw38_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger001_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
    WireRampOn True
End Sub

Sub ramptrigger0001_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger0002_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger0002_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger0003_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger0003_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger0003
End Sub

'//////////////////////////////////////////////////////////////////////
'// Ball Rolling
'//////////////////////////////////////////////////////////////////////

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
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
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
    ' Comment the next If block, if you are not implementing the Dyanmic Ball Shadows
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

'//////////////////////////////////////////////////////////////////////
'// Mechanic Sounds
'//////////////////////////////////////////////////////////////////////

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'   Metals (all metal objects, metal walls, metal posts, metal wire guides)
'   Apron (the apron walls and plunger wall)
'   Walls (all wood or plastic walls)
'   Rollovers (wire rollover triggers, star triggers, or button triggers)
'   Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'   Gates (plate gates)
'   GatesWire (wire gates)
'   Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Tutorial vides by Apophis
' Part 1:   https://youtu.be/PbE2kNiam3g
' Part 2:   https://youtu.be/B5cm1Y8wQsk
' Part 3:   https://youtu.be/eLhWyuYOyGg


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
SpinnerSoundLevel = 0.5                                       'volume level; range [0, 1]

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


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

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
  RandomSoundWall()
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

'//////////////////////////////////////////////////////////////////////
'// Ball
'//////////////////////////////////////////////////////////////////////

If BallBright Then
  table1.BallImage = "ball_HDR_brighter"
Else
  table1.BallImage = "ball_HDR"
End If


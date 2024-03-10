'Family Guy / IPD No. 5219 / 2006 / 4 Players

Option Explicit
Randomize

'//////////////////////////////////////////////////////////////////////
'// OPTIONS
'//////////////////////////////////////////////////////////////////////

const VR_Room = 0 ' 1 = VR Room; 0 = desktop or cab mode

'//////////////////////////////////////////////////////////////////////

Const VolumeDial = 0.8
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

'//////////////////////////////////////////////////////////////////////

' *** If using VR Room:

const CustomWalls = 0 'set to 0 for Modern Minimal Walls, floor, and roof, 1 for Sixtoe's original walls and floor
const WallClock = 1   '1 Shows the clock in the VR minimal rooms only
const topper = 1     '0 = Off 1 or 2 = On - Topper visible in VR Room only (1 = fun topper by addestratatore, 2 = Family Guy logo)
const poster = 1     '1 Shows the flyer posters in the VR room only
const poster2 = 1    '1 Shows the flyer posters in the VR room only

'//////////////////////////////////////////////////////////////////////
'// MODS
'//////////////////////////////////////////////////////////////////////

Const BallBright = 1        '0 - Normal, 1 - Bright
Const RubberColor = 0               '0 - Original (Black), 1 - White

'******************************************************
'* VPM INIT *******************************************
'******************************************************

Const BallSize = 50
Const BallMass = 1
Const tnob = 10
Const lob = 0           'Locked balls

Dim cab_mode, DesktopMode: DesktopMode = table1.ShowDT
If Not DesktopMode and VR_Room=0 Then cab_mode=1 Else cab_mode=0

Const AmbientBallShadowOn = 1
Const DynamicBallShadowsOn = 1
Const RubberizerEnabled = 1
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim UseVPMDMD:UseVPMDMD = DesktopMode

LoadVPM "01560000","sam.vbs",3.43

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOff = ""
Const SCoin = ""

'******************************************************
'* ROM VERSION ****************************************
'******************************************************

Const cGameName = "fg_1200af"

'******************************************************
'* VARIABLES ******************************************
'******************************************************

Dim bsTrough, PlungerIM, BIP, TVScoop, DrunkenClam, FARTBank, DeathTarget, CPUPos, CPDPos, MiniPF, MiniRight, MiniLeft
Dim  CastleGatePos, CastleGatePos2, CGUp, StewiePos, StewiePos2, StewiePos3, StewieDir, StewiePinball, MegPos

'******************************************************
'******************************************************
'******************************************************
'* TABLE INIT *****************************************
'******************************************************
'******************************************************
'******************************************************

Sub Table1_Init
LoadLUT

' Thalamus : Was missing 'vpminit me'
  vpminit me

'* ROM AND DMD ****************************************

With Controller
.GameName = cGameName
.SplashInfoLine = "Family Guy - STERN 2007"
.HandleMechanics = 0
.HandleKeyboard = 0
.ShowDMDOnly = 1
.ShowFrame = 0
.ShowTitle = 0
.Hidden = 0
On Error Resume Next
.Run
If Err Then MsgBox Err.Description
On Error Goto 0
End With

'* PINMAME TIMER **************************************

PinMAMETimer.Interval = PinMAMEInterval
PinMAMETimer.Enabled = 1

'* NUGDE **********************************************

vpmNudge.TiltSwitch=-7
vpmNudge.Sensitivity=5
vpmNudge.TiltObj=Array(bumper1,bumper2,bumper3,LeftSlingshot,RightSlingshot)

'* TROUGH *********************************************

Set bsTrough = New cvpmTrough
With bsTrough
.Size=4
.InitSwitches Array(21, 20, 19, 18)
.InitExit BallRelease, 90, 8
.Balls = 4
End With

'* MINI TROUGH ****************************************

Set MiniPF = New cvpmBallStack
With MiniPF
.InitSaucer sw55,55, 90, 29
.InitExitSnd SoundFX("Popper",DOFContactors), ""
.InitAddSnd "fx_kin"
.CreateEvents "MiniPF", sw55
End With

'* TV SCOOP *******************************************

Set TVScoop = New cvpmTrough
With TVScoop
.Size=3
.InitSwitches Array(13, 0, 0)
.InitExit sw13, 184.5, 20
.InitEntrySounds "VUKEnter", "", ""
.InitExitSounds "",SoundFX("VUKOut",DOFContactors)
.Balls = 0
End With

'* DRUNKEN CLAM ***************************************

Set DrunkenClam = New cvpmSaucer
With DrunkenClam
.InitKicker sw64,64, 183, 4, 0
.InitSounds "Saucer_Enter_2", "", SoundFX("Saucer_Kick",DOFContactors)
.CreateEvents "DrunkenClam", sw64
End With

''* FART TARGETS 4-BANK DROPTARGETS ********************
'
'Set FARTBank = New cvpmDropTarget
'With FARTBank
'.InitDrop Array(sw47,sw44,sw46,sw45), Array(44,47,45,46)
'.Initsnd SoundFX("",DOFContactors), SoundFX("Drop_Target_Reset_2",DOFContactors)
'End With
'
''* DEATH DROPTARGET *******************************
'
'Set DeathTarget = New cvpmDropTarget
'With DeathTarget
'.InitDrop Array(Array(sw9, sw9w)), Array(9)
'.Initsnd SoundFX("",DOFContactors), SoundFX("Drop_Target_Reset_1",DOFContactors)
'End With

'* IMPULSE PLUNGER ************************************

Const IMPowerSetting = 33
Const IMTime = 0.33
Set PlungerIM = New cvpmImpulseP
With plungerIM
.InitImpulseP BIPREG, IMPowerSetting, IMTime
.InitExitSnd "Popper", ""
.CreateEvents "PlungerIM"
End With

'* STARTUP CALLS **************************************
CPC.isdropped=1
kicker1.CreateSizedBallWithMass BallSize/2,Ballmass
kicker1.kick 180, 1
kicker1.enabled = 0
kicker2.CreateSizedBallWithMass BallSize/2,Ballmass
kicker2.kick 180, 1
kicker2.enabled = 0
kicker3.CreateSizedBallWithMass BallSize/2,Ballmass
kicker3.kick 180, 1
kicker3.enabled = 0
kicker4.CreateSizedBallWithMass BallSize/2,Ballmass
kicker4.kick 180, 1
kicker4.enabled = 0
sw55.createsizedball(14.6875)     'minipinball 5/8"
MiniPF.AddBall 0

PinCab_Backglass.blenddisablelighting = 1

InitVpmFFlipsSAM
'******************************************************
'******************************************************
'******************************************************
'* TABLE INIT END *************************************
'******************************************************
'******************************************************
'******************************************************

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// DRAIN HIT
'//////////////////////////////////////////////////////////////////////

Sub Drain_Hit:RandomSoundDrain Drain:bsTrough.AddBall Me:End Sub

'//////////////////////////////////////////////////////////////////////
'// Ball
'//////////////////////////////////////////////////////////////////////

If BallBright Then
  table1.BallImage = "old_ass_eyes_ball"
Else
  table1.BallImage = "MRBallDark2b"
End If

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
ShowLUT_Init
DisableLUTSelector = 0  ' Disables the ability to change LUT option with magna saves in game when set to 1

Sub Table1_exit:SaveLUT:Controller.Stop:End Sub

Sub SetLUT  'AXS
  Table1.ColorGradeImage = "LUT" & LUTset
  VRFlashLUT.imageA = "FlashLUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
  VRFlashLUT.opacity = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  VRFlashLUT.opacity = 100
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

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "FamilyGuyLUT.txt",True)
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
  If Not FileObj.FileExists(UserDirectory & "FamilyGuyLUT.txt") then
    LUTset=0
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "FamilyGuyLUT.txt")
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

Sub ShowLUT_Init
  LUTBox.visible = 0
  VRFlashLUT.opacity = 0
End Sub


'******************************************************
'* KEYS ***********************************************
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
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

If keycode = LeftTiltKey Then Nudge 90, 1.5:SoundNudgeLeft()
If keycode = RightTiltKey Then Nudge 270, 1.5:SoundNudgeRight()
If keycode = CenterTiltKey Then Nudge 0, 1.5:SoundNudgeCenter()

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

If keycode = PlungerKey Then
  SoundPlungerPull()
  Plunger.PullBack
  TimerVRPlunger.Enabled = True
  TimerVRPlunger1.Enabled = False
  VR_Primary_plunger.Y = -214.1667
End If


If keycode = RightFlipperKey Then
  Controller.Switch(90)=1
  Controller.Switch(82)=1
  VR_Cab_ButtonRight.transx = -8
  FlipperActivate RightFlipper, RFPress
    If MiniPF.Balls=0 Then
      PlaySound SoundFX("Stern_MiniFlipperUp2",DOFContactors)
      MiniPF_RightFlipper.RotateToEnd
      DOF 101, 1
      MiniRight=1
    End If

  vpmFFlipsSam.FlipR true

  Exit Sub

End If

If keycode = LeftFlipperKey Then
  Controller.Switch(84)=1
  VR_Cab_ButtonLeft.transx = 8
  FlipperActivate LeftFlipper, LFPress
    If MiniPF.Balls=0 Then
      PlaySound SoundFX("Stern_MiniFlipperUp1",DOFContactors)
      MiniPF_LeftFlipper.RotateToEnd
      DOF 102, 1
      MiniLeft=1
    End If

  vpmFFlipsSam.FlipL true

  Exit Sub

End If


If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal keycode)
'LUT controls
If keycode = LeftMagnaSave Then bLutActive = False

If KeyCode = PlungerKey Then
  Plunger.Fire
  TimerVRPlunger.Enabled = False
  TimerVRPlunger1.Enabled = True
  VR_Primary_plunger.Y = -214.1667

  If BIP = 1 Then
    SoundPlungerReleaseBall()                        'Plunger release sound when there is a ball in shooter lane
  Else
    SoundPlungerReleaseNoBall()                        'Plunger release sound when there is no ball in shooter lane
  End If
End If

If keycode = RightFlipperKey Then
  Controller.Switch(90)=0
  Controller.Switch(82)=0
  VR_Cab_ButtonRight.transx = 0
  FlipperDeActivate RightFlipper, RFPress
  If MiniRight=1 Then
    PlaySound SoundFX("Stern_MiniFlipperDown2",DOFContactors)
    MiniPF_RightFlipper.RotateToStart
    DOF 101, 0
    MiniRight=0
  End If

  vpmFFlipsSam.FlipR false
  Exit Sub
End If

If keycode = LeftFlipperKey Then
  Controller.Switch(84)=0
  VR_Cab_ButtonLeft.transx = 0
  FlipperDeActivate LeftFlipper, LFPress
  If MiniLeft=1 Then
    PlaySound SoundFX("Stern_MiniFlipperDown1",DOFContactors)
    MiniPF_LeftFlipper.RotateToStart
    DOF 102, 0
    MiniLeft=0
  End If

  vpmFFlipsSam.FlipL false
  Exit Sub
End If

If vpmKeyUp(keycode) Then Exit Sub


End Sub

'******************************************************
'* SOLENOID CALLS *************************************
'******************************************************

SolCallback(1)  = "solTrough"        'BallExit
SolCallback(2)  = "AutoLaunch"         'AutoLaunch
SolCallback(3)  = "SolDTDropUp"        '4-Bank Drop Tragets reset
SolCallback(4)  = "CPDSol"                   'Ball Saver Down
SolCallback(5)  = "DrunkenClam.SolOut"     'DrunkenClam Exit
SolCallback(6)  = "SolDTSweeperUp"         'Single Drop Targte reset
'SolCallback(7)  = "LeftSlingHit"            'Left Slingshot
'SolCallback(8)  = "RightSlingHit"           'Right Slingshot
'SolCallback(9)  = "Bump1Sol"                'Bottom Bumper
'SolCallback(10) = "Bump2Sol"                'Right Bumper
'SolCallback(11) = "Bump3Sol"                'Top Bumper
SolCallback(12) = "CPUSol"                   'Ball Saver Up
SolCallback(13) = "TVScoop.SolOut"                'Swamp Eject Scoop

SolCallback(16) = "SolRFlipper"
SolCallback(15) = "SolLFlipper"
SolCallback(17) = "Miniflipper_Left"
SolCallback(18) = "Miniflipper_Right"
SolCallback(19) = "CastleGuardSol"
SolCallBack(20) = "StewieMove"        'Stewie Motor Drive
SolCallBack(21) = "MiniPFSol"
SolCallBack(22) = "MegMove"         'Meg Move Solenoid
SolCallback(23) = "Flash1"      'Flash Lower Left
SolCallBack(24)= "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"          'Knocker
SolCallback(25) = "Flash3"      'Flash Backpanel Left
SolCallback(26) = "Flash4"      'Flash Backpanel Center
SolCallback(27) = "Flash5"      'Flash Backpanel Right
SolCallback(28) = "Flash6"      'Flash BeerCan
SolCallback(29) = "Flash29"     'Flash Fiona
SolCallback(30) = "Flash30"     'Flash RIght Orbit (Spinner)
SolCallback(31) = "Flash31"     'Flash Pop Bumpers
SolCallback(32) = "Flash2"      'Flash Lower Right

SolCallback(sLRFlipper) = "RFlipper"
SolCallback(sLLFlipper) = "LFlipper"

'//////////////////////////////////////////////////////////////////////

'********************************************
'  Drop Target Controls
'********************************************

Sub SolDTDropUp(Enabled)
  If Enabled Then
    debug.print "SolDTDropUp"
    DTRaise 44
    DTRaise 45
    DTRaise 46
        DTRaise 47
    RandomSoundDropTargetReset sw44p
  End If
End Sub

Sub SolDTSweeperUp(Enabled)
  If Enabled Then
    debug.print "SolDTSweeperUp"
    DTRaise 9
    RandomSoundDropTargetReset sw9p
  End If
End Sub


Sub sw9_hit
  DTHit 9
End Sub

Sub sw44_hit
  DTHit 44
End Sub

Sub sw45_hit
  DTHit 45
End Sub

Sub sw46_hit
  DTHit 46
End Sub

Sub sw47_hit
  DTHit 47
End Sub

Sub CheckDTs

    If gilvl=0 and controller.switch(44)= false then
Target4Off.visible=1
    Else
Target4Off.visible=0
    If gilvl=1 and controller.switch(44)= false then
Target4On.visible=1
    Else
Target4On.visible=0
    End if
    End If

    If gilvl=0 and controller.switch(45)= false then
Target3Off.visible=1
    Else
Target3Off.visible=0
    If gilvl=1 and controller.switch(45)= false then
Target3On.visible=1
    Else
Target3On.visible=0
    End if
    End If

    If gilvl=0 and controller.switch(46)= false then
Target2Off.visible=1
    Else
Target2Off.visible=0
    If gilvl=1 and controller.switch(46)= false then
Target2On.visible=1
    Else
Target2On.visible=0
    End if
    End If

    If gilvl=0 and controller.switch(47)= false then
Target1Off.visible=1
    Else
Target1Off.visible=0
    If gilvl=1 and controller.switch(47)= false then
Target1On.visible=1
    Else
Target1On.visible=0
    End if
    End If

    If gilvl=0 and controller.switch(9)= false then
TargetDeathOff.visible=1
    else
TargetDeathOff.visible=0
    If gilvl=1 and controller.switch(9)= false then
TargetDeathOn.visible=1
    else
TargetDeathOn.visible=0
    End if
    End If
end sub

'//////////////////////////////////////////////////////////////////////
'// FLIPPERS
'//////////////////////////////////////////////////////////////////////

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire :LeftFlipperSmall.RotateToEnd  'leftflipper.rotatetoend

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart:LeftFlipperSmall.RotateToStart
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
  FlipperLSh1.RotZ = LeftFlipperSmall.CurrentAngle
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

'******************************************************
'* CASTLE GUARD GATE **********************************
'******************************************************

Sub CastleGuardsw_Hit()
Controller.Switch(35)=0
CastleGuardUp.enabled=1
CGUp=1
CastleGuardsw.isdropped=1
End Sub

Sub CastleGuardSol (enabled)
If enabled then
Controller.Switch(35)=1
End if
End Sub

Sub CastleGuardUp_Timer()
Select Case CastleGatePos
Case 0: CastleGuardP.RotX=0:CastleGatePos=1
Case 1: CastleGuardP.RotX=0:
Case 2: CastleGuardP.RotX=-15
Case 3: CastleGuardP.RotX=-30
Case 4: CastleGuardP.RotX=-45
Case 5: CastleGuardP.RotX=-60
Case 6: CastleGuardP.RotX=-75
Case 7: CastleGuardP.RotX=-90:Me.Enabled=0:CastleGatePos=0
End Select
If CastleGatePos>0 then CastleGatePos=CastleGatePos+1
End Sub

Sub CastleGuardDown_Timer()
Select Case CastleGatePos2
Case 0: CastleGuardP.RotX=-90:CastleGatePos2=1
Case 1: CastleGuardP.RotX=-90:
Case 2: CastleGuardP.RotX=-75
Case 3: CastleGuardP.RotX=-60
Case 4: CastleGuardP.RotX=-45
Case 5: CastleGuardP.RotX=-30
Case 6: CastleGuardP.RotX=-15
Case 7: CastleGuardP.RotX=0:CastleGuardsw.isdropped=0:CGUp=0
Case 8: CastleGuardP.RotX=5
Case 9: CastleGuardP.RotX=10
Case 10: CastleGuardP.RotX=15
Case 11: CastleGuardP.RotX=20
Case 12: CastleGuardP.RotX=15
Case 13: CastleGuardP.RotX=10
Case 14: CastleGuardP.RotX=5
Case 15: CastleGuardP.RotX=0
Case 16: CastleGuardP.RotX=-5
Case 17: CastleGuardP.RotX=-10
Case 18: CastleGuardP.RotX=-15
Case 19: CastleGuardP.RotX=-10
Case 20: CastleGuardP.RotX=-5
Case 21: CastleGuardP.RotX=0
Case 22: CastleGuardP.RotX=5
Case 23: CastleGuardP.RotX=10
Case 24: CastleGuardP.RotX=15
Case 25: CastleGuardP.RotX=10
Case 26: CastleGuardP.RotX=5
Case 27: CastleGuardP.RotX=0
Case 28: CastleGuardP.RotX=-5
Case 29: CastleGuardP.RotX=-10
Case 30: CastleGuardP.RotX=-15
Case 31: CastleGuardP.RotX=-10
Case 32: CastleGuardP.RotX=-5
Case 33: CastleGuardP.RotX=0
Case 34: CastleGuardP.RotX=5
Case 35: CastleGuardP.RotX=10
Case 36: CastleGuardP.RotX=10
Case 37: CastleGuardP.RotX=5
Case 38: CastleGuardP.RotX=0
Case 39: CastleGuardP.RotX=-5
Case 40: CastleGuardP.RotX=-10
Case 41: CastleGuardP.RotX=-10
Case 42: CastleGuardP.RotX=-5
Case 43: CastleGuardP.RotX=0
Case 44: CastleGuardP.RotX=5
Case 45: CastleGuardP.RotX=10
Case 46: CastleGuardP.RotX=5
Case 47: CastleGuardP.RotX=0
Case 48: CastleGuardP.RotX=-5
Case 49: CastleGuardP.RotX=-10
Case 50: CastleGuardP.RotX=-5
Case 51: CastleGuardP.RotX=0
Case 52: CastleGuardP.RotX=5
Case 53: CastleGuardP.RotX=0
Case 54: CastleGuardP.RotX=-5
Case 55: CastleGuardP.RotX=0
Case 56: CastleGuardP.RotX=5
Case 57: CastleGuardP.RotX=0:Me.Enabled=0:CastleGatePos2=0
End Select
If CastleGatePos2>0 then CastleGatePos2=CastleGatePos2+1
End Sub

'******************************************************
'* MINI PLAYFIELD *************************************
'******************************************************

Sub Miniflipper_Right(Enabled)
If Enabled Then
PlaySound SoundFX("Stern_MiniFlipperUp1",DOFContactors):MiniPF_RightFlipper.RotateToEnd
Else
PlaySound SoundFX("Stern_MiniFlipperDown1",DOFContactors):MiniPF_RightFlipper.RotateToStart
End If
End Sub

Sub Miniflipper_Left(Enabled)
If Enabled Then
StewieSWTimer.Enabled=1
PlaySound SoundFX("Stern_MiniFlipperUp2",DOFContactors):MiniPF_LeftFlipper.RotateToEnd
Else
PlaySound SoundFX("Stern_MiniFlipperDown2",DOFContactors):MiniPF_LeftFlipper.RotateToStart
End If
End Sub

Sub MiniPFSol(enabled)
If enabled then
MiniPF.ExitSol_On
end if
End Sub

Sub StewieSWTimer_Timer()
StewieSW.enabled=1
End Sub

Sub StewieSW_hit()
StewiePinball=2
End Sub

Sub StewieTurn_Hit()
If StewiePinball=0 then
StewieTimer2.enabled=1
StewiePinball=3
StewieSWTimer.enabled=1
End If
End Sub

Sub StewieReset_Hit
If StewiePinball >0 then
StewiePinball = 0
End If
End Sub

Sub StewieDelay_Timer()
StewieTimer.Enabled = 1
StewieDelay.Enabled = 0
End Sub

'******************************************************
'* STEWIE MOVEMENT ************************************
'******************************************************

Sub StewieMove (enabled)
If enabled AND StewiePinball=0 then
StewieDelay.Enabled = 1
Else
If enabled AND StewiePinball=2 then
StewieTimer3.Enabled = 1
End If
End If
End Sub

Sub StewieTimer_Timer()
Select Case StewiePos
Case 0: StewieP.RotZ=0:StewiePos3=0:StewiePos=1
Case 1: StewieP.RotZ=+2
Case 2: StewieP.RotZ=+4
Case 3: StewieP.RotZ=+6
Case 4: StewieP.RotZ=+8
Case 5: StewieP.RotZ=+10
Case 6: StewieP.RotZ=+12
Case 7: StewieP.RotZ=+10
Case 8: StewieP.RotZ=+8
Case 9: StewieP.RotZ=+6
Case 10: StewieP.RotZ=+4
Case 11: StewieP.RotZ=+2
Case 12: StewieP.RotZ=0:StewiePos=1:StewieTimer.Enabled=0:StewiePinball=0
End Select
If StewiePos>0 then StewiePos=StewiePos+1
Sh5.RotZ = - StewieP.RotZ
End Sub

Sub StewieTimer2_Timer()
Select Case StewiePos2
Case 0: StewieP.RotZ=0:StewiePos2=1
Case 1: StewieP.RotZ=0
Case 2: StewieP.RotZ=-15
Case 3: StewieP.RotZ=-30
Case 4: StewieP.RotZ=-45
Case 5: StewieP.RotZ=-60
Case 6: StewieP.RotZ=-75
Case 7: StewieP.RotZ=-90
Case 8: StewieP.RotZ=-105
Case 9: StewieP.RotZ=-120
Case 10: StewieP.RotZ=-135
Case 11: StewieP.RotZ=-150
Case 12: StewieP.RotZ=-150:StewieTimer2.Enabled=0:StewiePos2=0:StewiePinball=3
End Select
If StewiePos2>0 then StewiePos2=StewiePos2+1
Sh5.RotZ = -StewieP.RotZ
End Sub

Sub StewieTimer3_Timer()
Select Case StewiePos3
Case 0: StewieP.RotZ=-150:StewiePos3=1
Case 1: StewieP.RotZ=-150
Case 2: StewieP.RotZ=-135
Case 3: StewieP.RotZ=-120
Case 4: StewieP.RotZ=-105
Case 5: StewieP.RotZ=-90
Case 6: StewieP.RotZ=-75
Case 7: StewieP.RotZ=-60
Case 8: StewieP.RotZ=-45
Case 9: StewieP.RotZ=-30
Case 10: StewieP.RotZ=-15
Case 11: StewieP.RotZ=0
Case 12: StewieP.RotZ=0:StewieTimer3.Enabled=0:StewiePos3=0:StewieSWTimer.enabled=0:StewiePinball=0:StewieSW.enabled=0
End Select
If StewiePos3>0 then StewiePos3=StewiePos3+1
Sh5.RotZ = -StewieP.RotZ
End Sub

'******************************************************
'* TROUGH *********************************************
'******************************************************

Sub solTrough(enabled)
If enabled then
bsTrough.ExitSol_On
vpmTimer.PulseSw 22
RandomSoundBallRelease Ballrelease
end if
End Sub

'******************************************************
'* AUTOLAUNCH *****************************************
'******************************************************

Sub AutoLaunch(Enabled)
If Enabled Then
PlungerIM.AutoFire
End If
End Sub

'******************************************************
'* TV SCOOP ****************************************
'******************************************************
Dim bBall, bZpos

Sub sw13_Hit
Set bBall = ActiveBall
bZpos = 50
Me.TimerInterval = 2
Me.TimerEnabled = 1
End Sub

Sub sw13_Timer
bBall.Z = bZpos
bZpos = bZpos-2
If bZpos <40 Then
Me.TimerEnabled = 0
Me.DestroyBall
TVScoop.AddBall Me
End If
End Sub

'******************************************************
'* LOIS SPINNER ***************************************
'******************************************************
Sub sw39_Spin:vpmTimer.PulseSw 39: SoundSpinner SW39:End Sub

'******************************************************
'* POP BUMPERS ****************************************
'******************************************************

Dim dirRing1 : dirRing1 = -1
Dim dirRing2 : dirRing2 = -1
Dim dirRing3 : dirRing3 = -1

Sub Bumper1_Hit : vpmTimer.PulseSw 32 : RandomSoundBumperTop Bumper1 : End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw 31 : RandomSoundBumperMiddle Bumper2 : End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw 30 : RandomSoundBumperBottom Bumper3 : End Sub

Sub Bumper1_timer()
  BR1.Z = BR1.Z + (5 * dirRing1)
  If BR1.Z <= -35 Then dirRing1 = 1
  If BR1.Z >= 0 Then
    dirRing1 = -1
    BR1.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  BR2.Z = BR2.Z + (5 * dirRing2)
  If BR2.Z <= -35 Then dirRing2 = 1
  If BR2.Z >= 0 Then
    dirRing2 = -1
    BR2.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper3_timer()
  BR3.Z = BR3.Z + (5 * dirRing3)
  If BR3.Z <= -35 Then dirRing3 = 1
  If BR3.Z >= 0 Then
    dirRing3 = -1
    BR3.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

'******************************************************
'* TARGETS ********************************************
'******************************************************

Sub sw3_Hit:vpmTimer.PulseSw 3:If CGUp= 1 then CastleGuardDown.enabled=1:End If:End Sub

Sub sw4_Hit:vpmTimer.PulseSw 4:End Sub

Sub sw5_Hit:vpmTimer.PulseSw 5:End Sub

Sub sw8_Hit:vpmTimer.PulseSw 8:End Sub

Sub sw10_Hit:vpmTimer.PulseSw 10:End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41:End Sub

Sub sw42_Hit:vpmTimer.PulseSw 42:End Sub

Sub sw43_Hit:vpmTimer.PulseSw 43:End Sub


'******************************************************
'* MINIPF TARGETS **************************************
'******************************************************

Sub sw50_Hit:vpmTimer.PulseSw 50:End Sub
Sub sw51_Hit:vpmTimer.PulseSw 51:End Sub

'******************************************************
'* MINI PF SWITCHES ***********************************
'******************************************************

Sub sw52_Hit:Controller.Switch(52) = 1:End Sub
Sub sw52_Unhit:Controller.Switch(52) = 0:End Sub

Sub sw53_Hit:Controller.Switch(53) = 1:End Sub
Sub sw53_Unhit:Controller.Switch(53) = 0:End Sub

Sub sw54_Hit:Controller.Switch(54) = 1:End Sub
Sub sw54_Unhit:Controller.Switch(54) = 0:End Sub

'******************************************************
'* ROLLOVER SWITCHES **********************************
'******************************************************

const zdepth = -5

Sub sw6_Hit:Controller.Switch(6) = 1:sw6p.Z=zdepth:End Sub
Sub sw6_Unhit:Controller.Switch(6) = 0:sw6p.Z=0:End Sub

Sub sw7_Hit:Controller.Switch(7) = 1:sw7p.Z=zdepth:End Sub
Sub sw7_Unhit:Controller.Switch(7) = 0:sw7p.Z=0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:sw23p.Z=zdepth:BIP=1:End Sub
Sub sw23_Unhit:Controller.Switch(23) = 0:sw23p.Z=0:BIP=0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:sw24p.Z=zdepth:End Sub
Sub sw24_Unhit:Controller.Switch(24) = 0:sw24p.Z=0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:sw25p.Z=zdepth:End Sub
Sub sw25_Unhit:Controller.Switch(25) = 0:sw25p.Z=0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:sw28p.Z=zdepth:End Sub
Sub sw28_Unhit:Controller.Switch(28) = 0:sw28p.Z=0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:sw29p.Z=zdepth:End Sub
Sub sw29_Unhit:Controller.Switch(29) = 0:sw29p.Z=0:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:sw40p.Z=zdepth:End Sub
Sub sw40_Unhit:Controller.Switch(40) = 0:sw40p.Z=0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:sw57p.Z=zdepth:End Sub
Sub sw57_Unhit:Controller.Switch(57) = 0:sw57p.Z=0:End Sub

'******************************************************
'* SLINGSHOTS *****************************************
'******************************************************
Dim LStep, RStep
Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotLeft sling1
    vpmTimer.PulseSw 26
    LSling.Visible = 0
    LSling1.Visible = 1
    sling1.TransZ = -20
    LStep = 0
   Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:Me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotRight sling2
    vpmTimer.PulseSw 27
    RSling.Visible = 0
    RSling1.Visible = 1
    sling2.TransZ = -20
    RStep = 0
   Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:Me.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'******************************************************
'* CENTER POST ****************************************
'******************************************************
'* DJRobX fixed the ball saver emulation error. Game  *
'* seems to fire both up and down to put the saver    *
'* down. Create a "latch" to hold it down'            *
'******************************************************

Dim SaverPos, SaverTime
SaverPos = 0
SaverTime = Now

Sub CPDSol (enabled)
If enabled then
SaverTime = DateAdd("s",3, Now)
CPUSol (true)
End If
End Sub

Sub CPUSol (enabled)
If enabled then
If SaverTime > Now Then
If SaverPos = 1 Then
' ** Move Saver Down **'
Controller.Switch(2) = 1
Controller.Switch(1) = 0
CenterpostDown.Enabled=1
SaverPos=0
CPC.isdropped=1
End If
Else
If SaverPos = 0 Then
' ** Move Saver Up ** '
Controller.Switch(2) = 0
Controller.Switch(1) = 1
CenterpostUp.Enabled=1
SaverPos=1
CPC.isdropped=0
End If
End If
End If
End Sub

Sub CenterpostUp_Timer()
Select Case CPUPos
Case 0: CenterPost.Z=-15
Case 1: CenterPost.Z=-05:Playsound SoundFX("Stern_Centerpost_Up",DOFContactors)
Case 2: CenterPost.Z=0
Case 3: CenterPost.Z=0:CPUPos=0:CenterpostUp.Enabled=0
End Select
If CPUPos=>0 then CPUPos=CPUPos+1
End Sub

Sub CenterpostDown_Timer()
Select Case CPDPos
Case 0: CenterPost.Z=-05
Case 1: CenterPost.Z=-15:Playsound SoundFX("Stern_Centerpost_Down",DOFContactors)
Case 2: CenterPost.Z=-26
Case 3: CenterPost.Z=-26:CPDPos=0:CenterpostDown.Enabled=0
End Select
If CPDPos=>0 then CPDPos=CPDPos+1
End Sub

'******************************************************
'* MEG MOVEMENT ***************************************
'******************************************************
Sub MegMove (enabled)
if enabled then
MegTimer.Enabled=1
End If
End Sub

Sub MegTimer_Timer()
Select Case MegPos
Case 0: MegP.Z=180
Case 1: MegP.Z=175:Playsound SoundFX("Stern_Megshake",DOFContactors)
Case 2: MegP.Z=170
Case 3: MegP.Z=165
Case 4: MegP.Z=160
Case 5: MegP.Z=170
Case 6: MegP.Z=180
Case 20: MegP.Z=180:MegPos=0:Me.Enabled=0
End Select
If MegPos=>0 then MegPos=MegPos+1
End Sub

'******************************************************
'* BRIAN BEERCAN TARGET *******************************
'******************************************************
Dim BeerCanPos:BeerCanPos=1
Dim BeerCanDir : BeerCanDir=0
Sub sw49_Hit:vpmTimer.PulseSw 49:Me.TimerEnabled = 1:End Sub

Sub sw49_Timer()
sw49P.Rotx = sw49P.Rotx - 2*BeerCanPos
BrianP.RotX= sw49P.RotX
If sw49P.Rotx < - 10 AND BeerCanPos=1 Then BeerCanPos=-1
If sw49P.Rotx > 6 AND BeerCanPos=-1 Then BeerCanDir=1 : BeerCanPos=1
If sw49P.Rotx <0 AND BeerCanDir=1 Then sw49P.Rotx =0 : BeerCanDir=0 : Me.TimerEnabled = 0
End Sub
'******************************************************
'* LEFT GATE ******************************************
'******************************************************

Sub LeftGate_hit:vpmTimer.PulseSw 33:LeftGateFl.RotateToEnd:End Sub
Sub LeftGate_Unhit:LeftGateFl.RotateToStart:End Sub



'******************************************************
'* GENERAL ILLUMINATION *******************************
'******************************************************

Set GICallBack = GetRef("UpdateGI")
dim gilvl:gilvl = 1
Sub UpdateGI(nr,enabled)
Dim x
If (Enabled) Then
For each x in GI:x.state=1:Next
gilvl = 1
Sound_GI_Relay 1, relaysound1

Backwall.image="Backwall_On"
House.image="House_On"
sideblades.image="sideblades_ON"
RenderOff.visible = 0
RenderOn.visible = 1
UpperPlayfieldOff.visible = 0
UpperPlayfieldOn.visible = 1
Primitive66.image = "3D_Monkeygate_On"
F23a2.visible = 1
Primitive65.image = "LeftRamp_Gate_On"
F23a001.visible = 1

'Bulbs

Bulb030.blenddisablelighting = 10
Bulb002.blenddisablelighting = 10
Bulb003.blenddisablelighting = 10
Bulb004.blenddisablelighting = 10
Bulb005.blenddisablelighting = 10
Bulb006.blenddisablelighting = 10
Bulb007.blenddisablelighting = 10
Bulb008.blenddisablelighting = 10
Bulb009.blenddisablelighting = 10
Bulb010.blenddisablelighting = 10
Bulb011.blenddisablelighting = 10
Bulb012.blenddisablelighting = 10
Bulb013.blenddisablelighting = 10
Bulb014.blenddisablelighting = 10
Bulb015.blenddisablelighting = 10
Bulb016.blenddisablelighting = 10
Bulb017.blenddisablelighting = 10
Bulb018.blenddisablelighting = 10
Bulb019.blenddisablelighting = 10
Bulb020.blenddisablelighting = 10
Bulb021.blenddisablelighting = 10
Bulb022.blenddisablelighting = 10
Bulb023.blenddisablelighting = 10

'Bulbs Flasher

Bulb024.blenddisablelighting = 60
Bulb025.blenddisablelighting = 60
Bulb026.blenddisablelighting = 60
Bulb027.blenddisablelighting = 60

For each x in BWGI:x.state=2:Next
gilvl = 1
Else
For each x in GI:x.state=0:Next
gilvl = 0
Sound_GI_Relay 0, relaysound1

Backwall.image="Backwall_Off"
House.image="House_Off"
sideblades.image="sideblades_OFF"
RenderOff.visible = 1
RenderOn.visible = 0
UpperPlayfieldOff.visible = 1
UpperPlayfieldOn.visible = 0
Primitive66.image = "3D_Monkeygate_Off"
F23a2.visible = 0
Primitive65.image = "LeftRamp_Gate_Off"
F23a001.visible = 0


'Bulbs

Bulb002.blenddisablelighting = 0
Bulb003.blenddisablelighting = 0
Bulb004.blenddisablelighting = 0
Bulb005.blenddisablelighting = 0
Bulb006.blenddisablelighting = 0
Bulb007.blenddisablelighting = 0
Bulb008.blenddisablelighting = 0
Bulb009.blenddisablelighting = 0
Bulb010.blenddisablelighting = 0
Bulb011.blenddisablelighting = 0
Bulb012.blenddisablelighting = 0
Bulb013.blenddisablelighting = 0
Bulb014.blenddisablelighting = 0
Bulb015.blenddisablelighting = 0
Bulb016.blenddisablelighting = 0
Bulb017.blenddisablelighting = 0
Bulb018.blenddisablelighting = 0
Bulb019.blenddisablelighting = 0
Bulb020.blenddisablelighting = 0
Bulb021.blenddisablelighting = 0
Bulb022.blenddisablelighting = 0
Bulb023.blenddisablelighting = 0
Bulb030.blenddisablelighting = 0

'Bulbs Flasher

Bulb024.blenddisablelighting = 0
Bulb025.blenddisablelighting = 0
Bulb026.blenddisablelighting = 0
Bulb027.blenddisablelighting = 0

For each x in BWGI:x.state=0:Next
gilvl = 0
End If
End Sub

'******************************************************
'* REAL TIME UPDATES *********************************
'******************************************************

'Set MotorCallback = GetRef("GameTimer")

Sub UpdateVisuals
  RightFlipperP.Rotz = RightFlipper.CurrentAngle
  LeftFlipperP.Rotz = LeftFlipper.CurrentAngle
  LeftFlipperSmallP.RotY = LeftFlipperSmall.CurrentAngle
  MiniPF_LeftFlipperP.Rotz = MiniPF_LeftFlipper.CurrentAngle
  MiniPF_RightFlipperP.Rotz = MiniPF_RightFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  Gate1p.Roty = Gate1.CurrentAngle
  LeftGatep.Rotx = LeftGate.CurrentAngle
  RightSpinnerP.Rotx = sw39.CurrentAngle
  RightSpinnerP1.Rotx = sw39.CurrentAngle
  LeftGateSwitch.Rotx = LeftGateFl.CurrentAngle
End Sub

' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
  RollingUpdate         'update rolling sounds
  DoDTAnim
  UpdateVisuals
CheckDTs
End Sub


' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
  Lampz.Update 'updates on frametime (Object updates only)
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


'******************************************************
'**** RAMP ROLLING SFX
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

'*******************************************
'  Ramp Triggers
'*******************************************

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
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
' - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
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
'***If there are more than 3 lights that overlap in a playable area, exclude the less important lights***
' This is because the code will only project two shadows if they are coming from lights that are consecutive in the collection, and more than 3 will cause "jumping" between which shadows are drawn
' The easiest way to keep track of this is to start with the group on the right slingshot and move anticlockwise around the table
' For example, if you use 6 lights: A & B on the left slingshot and C & D on the right, with E near A&B and F next to C&D, your collection would look like EBACDF
'
'G        H                     ^ E
'                             ^ B
' A    C                        ^ A
'  B    D     your collection should look like  ^ G   because E&B, B&A, etc. intersect; but B&D or E&F do not
'  E      F                       ^ H
'                             ^ C
'                             ^ D
'                             ^ F
'   When selecting them, you'd shift+click in this order^^^^^

'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
'Sub FrameTimer_Timer()
' If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
'End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
'Const lob = 0  'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 '2 = flasher image shadow, but it moves like ninuzzu's

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
'   ...rolling(b) = False
'   ...StopSound("BallRoll_" & b)
' Next
'
'...rolling and drop sounds...

'   If DropCount(b) < 5 Then
'     DropCount(b) = DropCount(b) + 1
'   End If
'
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
'
'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******
Dim sourcenames, currentShadowCount, DSSources(30), numberofsources, numberofsources_hold
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(12), objrtx2(12)
dim objBallShadow(12)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9,BallShadowA10,BallShadowA11)

DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

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
  Dim s, Source, LSd, currentMat, AnotherSource, BOT, iii
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
        For iii = 0 to numberofsources - 1
          LSd=DistanceFast((BOT(s).x-DSSources(iii)(0)),(BOT(s).y-DSSources(iii)(1))) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then               'If the ball is within the falloff range of a light and light is on (we will set numberofsources to 0 when GI is off)
            currentShadowCount(s) = currentShadowCount(s) + 1   'Within range of 1 or 2
            if currentShadowCount(s) = 1 Then           '1 dynamic shadow source
              sourcenames(s) = iii
              currentMat = objrtx1(s).material
              objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01            'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
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
              objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(AnotherSource)(0),DSSources(AnotherSource)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-DistanceFast((BOT(s).x-DSSources(AnotherSource)(0)),(BOT(s).y-DSSources(AnotherSource)(1))))/falloff
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

              currentMat = objrtx2(s).material
              objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
  '           objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx2(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
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


'******************************************************
'*****   FLUPPER DOMES
'******************************************************
' Based on FlupperDoms2.2

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
' Example: RotateFlasher 1, 180     'where 1 is the flasher number and 180 is the angle of Z rotation

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

' example script for rom based tables (non modulated):

' SolCallback(25)="FlashRed"
'
' Sub FlashRed(flstate)
' If Flstate Then
'   Objlevel(1) = 1 : FlasherFlash1_Timer
' End If
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
' Objlevel(1) = level/255 : FlasherFlash1_Timer
' End Sub


 Sub Flash1(Enabled)
  If Enabled Then
    ObjTargetLevel(1) = 1
  Else
    ObjTargetLevel(1) = 0
  End If
  FlasherFlash1_Timer
  Sound_Flash_Relay enabled, Flasherbase1
 End Sub

  Sub Flash2(Enabled)
  If Enabled Then
    ObjTargetLevel(2) = 1
  Else
    ObjTargetLevel(2) = 0
  End If
  FlasherFlash2_Timer
  Sound_Flash_Relay enabled, Flasherbase2
 End Sub

 Sub Flash3(Enabled)
  If Enabled Then
    ObjTargetLevel(3) = 1
  Else
    ObjTargetLevel(3) = 0
  End If
  FlasherFlash3_Timer
  Sound_Flash_Relay enabled, Flasherbase3
 End Sub

  Sub Flash4(Enabled)
  If Enabled Then
    ObjTargetLevel(4) = 1
  Else
    ObjTargetLevel(4) = 0
  End If
  FlasherFlash4_Timer
  Sound_Flash_Relay enabled, Flasherbase4
 End Sub

 Sub Flash5(Enabled)
  If Enabled Then
    ObjTargetLevel(5) = 1
  Else
    ObjTargetLevel(5) = 0
  End If
  FlasherFlash5_Timer
  Sound_Flash_Relay enabled, Flasherbase5
 End Sub

 Sub Flash6(Enabled)
  If Enabled Then
    ObjTargetLevel(6) = 1
  Else
    ObjTargetLevel(6) = 0
  End If
  FlasherFlash6_Timer
  Sound_Flash_Relay enabled, Flasherbase6
 End Sub



Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.1   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.3   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.2   ' *** lower this, if the blooms are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"

InitFlasher 1, "yellow"
InitFlasher 2, "yellow"
InitFlasher 3, "green"
InitFlasher 4, "yellow"
InitFlasher 5, "white"
InitFlasher 6, "white"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 1,17 : RotateFlasher 2,0 : RotateFlasher 3,90 : RotateFlasher 4,90
RotateFlasher 1,180
RotateFlasher 2,180
RotateFlasher 3,180
RotateFlasher 4,180
RotateFlasher 5,180


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
  if nr=6 then objbase(nr).BlendDisableLighting = 0

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 then
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
  objlight(nr).bulbhaloheight = objbase(nr).z -10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  dim xthird, ythird
  xthird = tablewidth/3
  ythird = tableheight/3

  If objbase(nr).x >= xthird and objbase(nr).x <= xthird*2 then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  end if

  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objbloom(nr).color = RGB(4,120,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4) : objbloom(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4) : objbloom(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255) : objbloom(nr).color = RGB(230,49,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50) : objbloom(nr).color = RGB(200,173,25)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59) : objbloom(nr).color = RGB(255,240,150)
    Case "orange" :  objlight(nr).color = RGB(255,70,0) : objflasher(nr).color = RGB(255,70,0) : objbloom(nr).color = RGB(255,70,0)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  if nr=6 then
    objbase(nr).BlendDisableLighting =  1 * ObjLevel(nr)^3
  else
    objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  end if
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  if round(ObjTargetLevel(nr),1) > round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) + 0.3
    if ObjLevel(nr) > 1 then ObjLevel(nr) = 1
  Elseif round(ObjTargetLevel(nr),1) < round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
    if ObjLevel(nr) < 0 then ObjLevel(nr) = 0
  Else
    ObjLevel(nr) = round(ObjTargetLevel(nr),1)
    objflasher(nr).TimerEnabled = False
  end if
  'ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub


'******************************************************
'******  END FLUPPER DOMES
'******************************************************

'******************************************************
'****  DROP TARGETS by Rothbauerw
'******************************************************
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.
'
'Add a Timer named DTAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
'DTAnim.interval = 10
'DTAnim.enabled = True

'Sub DTAnim_Timer
' DoDTAnim
' DoSTAnim
'End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.


'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

'Define a variable for each drop target
Dim DT9, DT47,DT46,DT45,DT44

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


DT9 = Array(sw9, sw9a, sw9p, 9, 0)
DT47 = Array(sw47, sw47a, sw47p, 47, 0)
DT46 = Array(sw46, sw46a, sw46p, 46, 0)
DT45 = Array(sw45, sw45a, sw45p, 45, 0)
DT44 = Array(sw44, sw44a, sw44p, 44, 0)


Dim DTArray
DTArray = Array(DT9,DT47,DT46,DT45,DT44)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
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

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i)(4) =  DTCheckBrick(Activeball,DTArray(i)(2))
  If DTArray(i)(4) = 1 or DTArray(i)(4) = 3 or DTArray(i)(4) = 4 Then
    DTBallPhysics Activeball, DTArray(i)(2).rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) = -1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 to uBound(DTArray)
    If DTArray(i)(3) = switch Then DTArrayID = i:Exit Function
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


Sub DoDTAnim()
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray(i)(4) = DTAnimate(DTArray(i)(0),DTArray(i)(1),DTArray(i)(2),DTArray(i)(3),DTArray(i)(4))
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch,  animate)
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
    DTAnimate = animate
    Exit Function
    elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  if animate = 2 Then
    transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
    if prim.transz > -DTDropUnits  Then
      prim.transz = transz
    end if

    prim.rotx = DTMaxBend * cos(rangle)/2
    prim.roty = DTMaxBend * sin(rangle)/2

    if prim.transz <= -DTDropUnits Then
      prim.transz = -DTDropUnits
      secondary.collidable = 0
      controller.Switch(Switchid) = 1
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
  elseif animate = 3 and animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  if animate = -1 Then
    transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

    If prim.transz = -DTDropUnits Then
      Dim BOT, b
      BOT = GetBalls

      For b = 0 to UBound(BOT)
        If InRotRect(BOT(b).x,BOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and BOT(b).z < prim.z+DTDropUnits+25 Then
          BOT(b).velz = 20
        End If
      Next
    End If

    if prim.transz < 0 Then
      prim.transz = transz
    elseif transz > 0 then
      prim.transz = transz
    end if

    if prim.transz > DTDropUpUnits then
      DTAnimate = -2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1
    controller.Switch(Switchid) = 0

  End If

  if animate = -2 and animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
    if prim.transz < 0 then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    end If
  End If
End Function


'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
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
'****  END DROP TARGETS
'******************************************************



'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************

dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection

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
  AddSlingsPt 0, 0.00,  -4
  AddSlingsPt 1, 0.45,  -7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4

End Sub


Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LS, RS)
  dim x : for each x in a
    x.addpoint idx, aX, aY
  Next
End Sub

Class SlingshotCorrection
  Public DebugOn, Enabled
  private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): Enabled = True : End Sub

  Public Property let Object(aInput) : Set Slingshot = aInput : End Property
  Public Property Let EndPoint1(aInput) : SlingX1 = aInput.x: SlingY1 = aInput.y: End Property
  Public Property Let EndPoint2(aInput) : SlingX2 = aInput.x: SlingY2 = aInput.y: End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 then Report
  End Sub

  Public Sub Report()         'debug, reports all coords in tbPL.text
    If not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1 : YL = SlingY1 : XR = SlingX2 : YR = SlingY2
    Else
      XL = SlingX2 : YL = SlingY2 : XR = SlingX1 : YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If ABS(XR-XL) > ABS(YR-YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If not IsEmpty(ModIn(0) ) then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      'debug.print " BallPos=" & BallPos &" Angle=" & Angle
      'debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled then aBall.Velx = RotVxVy(0)
      If Enabled then aBall.Vely = RotVxVy(1)
      'debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      'debug.print " "
    End If
  End Sub

End Class


'******************************************************
'****  Solenoid Called Lamps
'******************************************************

sub Flash29(Enabled)
  If Enabled Then
    Lampz.state(199) = 1
  Else
    Lampz.state(199) = 0
  End If
End Sub

sub Flash30(Enabled)
  If Enabled Then
    Lampz.state(190) = 1
  Else
    Lampz.state(190) = 0
  End If
End Sub

sub Flash31(Enabled)
  If Enabled Then
    Lampz.state(191) = 1
  Else
    Lampz.state(191) = 0
  End If
End Sub


'******************************************************
'****  LAMPZ by nFozzy
'******************************************************

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
InitLampsNF               ' Setup lamp assignments
LampTimer.Interval = 16   ' Using fixed value so the fading speed is same for every fps
LampTimer.Enabled = 1

Sub LampTimer_Timer()
  dim x, chglamp
  chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
    next
  End If
  Lampz.Update1 'update (fading logic only)
End Sub



Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

const insert_dl_on_white = 120
const insert_dl_on_white_tri = 180
const insert_dl_on_red = 100
const insert_dl_on_orange = 60
const insert_dl_on_blue = 200
const insert_dl_on_yellow = 40
const insert_dl_on_yellow_tri = 180
const boost = 60


Sub InitLampsNF()

  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

  'Adjust fading speeds (1 / full MS fading time)
  dim x : for x = 0 to 200 : Lampz.FadeSpeedUp(x) = 5/10 : Lampz.FadeSpeedDown(x) = 35/100 : next 'Updated to match original fading..

  Lampz.MassAssign(3) = l3
  Lampz.MassAssign(3) = l3a
  Lampz.Callback(3) = "DisableLighting p3, insert_dl_on_orange,"
  Lampz.MassAssign(4) = l4
  Lampz.MassAssign(4) = l4a
  Lampz.Callback(4) = "DisableLighting p4, insert_dl_on_orange,"
  Lampz.MassAssign(5) = l5
  Lampz.MassAssign(5) = l5a
  Lampz.Callback(5) = "DisableLighting p5, insert_dl_on_orange,"
  Lampz.MassAssign(6) = l6
  Lampz.MassAssign(6) = l6a
  Lampz.Callback(6) = "DisableLighting p6, insert_dl_on_orange,"
  Lampz.MassAssign(7) = l7
  Lampz.MassAssign(7) = l7a
  Lampz.Callback(7) = "DisableLighting p7, insert_dl_on_orange,"
  Lampz.MassAssign(8) = l8
  Lampz.MassAssign(8) = l8a
  Lampz.Callback(8) = "DisableLighting p8, insert_dl_on_orange,"
  Lampz.MassAssign(9) = l9
  Lampz.MassAssign(9) = l9a
  Lampz.Callback(9) = "DisableLighting p9, insert_dl_on_orange,"
  Lampz.MassAssign(10) = l10
  Lampz.MassAssign(10) = l10a
  Lampz.Callback(10) = "DisableLighting p10, insert_dl_on_orange,"
  Lampz.MassAssign(11) = l11
  Lampz.MassAssign(11) = l11a
  Lampz.Callback(11) = "DisableLighting p11, insert_dl_on_orange,"
  Lampz.MassAssign(12) = l12
  Lampz.MassAssign(12) = l12a
  Lampz.Callback(12) = "DisableLighting p12, insert_dl_on_orange,"
  Lampz.MassAssign(13) = l13
  Lampz.MassAssign(13) = l13a
  Lampz.Callback(13) = "DisableLighting p13, insert_dl_on_orange,"
  Lampz.MassAssign(14) = l14
  Lampz.MassAssign(14) = l14a
  Lampz.Callback(14) = "DisableLighting p14, insert_dl_on_orange,"
  Lampz.MassAssign(15) = l15
  Lampz.MassAssign(15) = l15a
  Lampz.Callback(15) = "DisableLighting p15, insert_dl_on_orange,"
  Lampz.MassAssign(16) = l16
  Lampz.MassAssign(16) = l16a
  Lampz.Callback(16) = "DisableLighting p16, insert_dl_on_orange,"
  Lampz.MassAssign(17) = l17
  Lampz.MassAssign(17) = l17a
  Lampz.Callback(17) = "DisableLighting p17, insert_dl_on_white,"
  Lampz.MassAssign(18) = l18
  Lampz.MassAssign(18) = l18a
  Lampz.Callback(18) = "DisableLighting p18, insert_dl_on_white,"
  Lampz.MassAssign(19) = l19
  Lampz.MassAssign(19) = l19a
  Lampz.Callback(19) = "DisableLighting p19, insert_dl_on_white,"
  Lampz.MassAssign(20) = l20
  Lampz.MassAssign(20) = l20a
  Lampz.Callback(20) = "DisableLighting p20, insert_dl_on_orange,"
  Lampz.MassAssign(21) = l21
  Lampz.MassAssign(21) = l21a
  Lampz.Callback(21) = "DisableLighting p21, insert_dl_on_orange,"
  Lampz.MassAssign(22) = l22
  Lampz.MassAssign(22) = l22a
  Lampz.Callback(22) = "DisableLighting p22, insert_dl_on_red,"
  Lampz.MassAssign(23) = l23
  Lampz.MassAssign(23) = l23a
  Lampz.Callback(23) = "DisableLighting p23, insert_dl_on_red,"
  Lampz.MassAssign(24) = l24
  Lampz.MassAssign(24) = l24a
  Lampz.Callback(24) = "DisableLighting p24, insert_dl_on_blue,"
  Lampz.MassAssign(25) = l25
  Lampz.MassAssign(25) = l25a
  Lampz.Callback(25) = "DisableLighting p25, insert_dl_on_white,"
  Lampz.MassAssign(26) = l26
  Lampz.MassAssign(26) = l26a
  Lampz.Callback(26) = "DisableLighting p26, insert_dl_on_red,"
  Lampz.MassAssign(27) = l27
  Lampz.MassAssign(27) = l27a
  Lampz.Callback(27) = "DisableLighting p27, insert_dl_on_yellow,"
  Lampz.MassAssign(28) = l28
  Lampz.MassAssign(28) = l28a
  Lampz.Callback(28) = "DisableLighting p28, insert_dl_on_red,"
  Lampz.MassAssign(29) = l29
  Lampz.MassAssign(29) = l29a
  Lampz.Callback(29) = "DisableLighting p29, insert_dl_on_white_tri,"
  Lampz.MassAssign(30) = l30
  Lampz.MassAssign(30) = l30a
  Lampz.Callback(30) = "DisableLighting p30, insert_dl_on_yellow_tri,"
  Lampz.MassAssign(31) = l31
  Lampz.MassAssign(31) = l31a
  Lampz.Callback(31) = "DisableLighting p31, insert_dl_on_yellow_tri,"
  Lampz.MassAssign(32) = l32
  Lampz.MassAssign(32) = l32a
  Lampz.Callback(32) = "DisableLighting p32, insert_dl_on_white_tri,"
  Lampz.MassAssign(33) = l33
  Lampz.MassAssign(33) = l33a
  Lampz.Callback(33) = "DisableLighting p33, insert_dl_on_orange + boost,"
  Lampz.MassAssign(34) = l34
  Lampz.MassAssign(34) = l34a
  Lampz.Callback(34) = "DisableLighting p34, insert_dl_on_white_tri,"
  Lampz.MassAssign(35) = l35
  Lampz.MassAssign(35) = l35a
  Lampz.Callback(35) = "DisableLighting p35, insert_dl_on_yellow,"
  Lampz.MassAssign(36) = l36
  Lampz.MassAssign(36) = l36a
  Lampz.Callback(36) = "DisableLighting p36, insert_dl_on_yellow,"
  Lampz.MassAssign(37) = l37
  Lampz.MassAssign(37) = l37a
  Lampz.Callback(37) = "DisableLighting p37, insert_dl_on_yellow,"
  Lampz.MassAssign(38) = l38
  Lampz.MassAssign(38) = l38a
  Lampz.Callback(38) = "DisableLighting p38, insert_dl_on_yellow,"
  Lampz.MassAssign(39) = l39
  Lampz.MassAssign(39) = l39a
  Lampz.Callback(39) = "DisableLighting p39, insert_dl_on_white_tri,"
  Lampz.MassAssign(40) = l40
  Lampz.MassAssign(40) = l40a
  Lampz.Callback(40) = "DisableLighting p40, insert_dl_on_yellow_tri,"
  Lampz.MassAssign(41) = l41
  Lampz.MassAssign(41) = l41a
  Lampz.Callback(41) = "DisableLighting p41, insert_dl_on_yellow,"
  Lampz.MassAssign(42) = l42
  Lampz.MassAssign(42) = l42a
  Lampz.Callback(42) = "DisableLighting p42, insert_dl_on_white_tri + boost,"
  Lampz.MassAssign(43) = l43
  Lampz.MassAssign(43) = l43a
  Lampz.Callback(43) = "DisableLighting p43, insert_dl_on_white + boost,"
  Lampz.MassAssign(44) = l44
  Lampz.MassAssign(44) = l44a
  Lampz.Callback(44) = "DisableLighting p44, insert_dl_on_orange + boost,"
  Lampz.MassAssign(45) = l45
  Lampz.MassAssign(45) = l45a
  Lampz.Callback(45) = "DisableLighting p45, insert_dl_on_orange + boost,"
  Lampz.MassAssign(46) = l46
  Lampz.MassAssign(46) = l46a
  Lampz.Callback(46) = "DisableLighting p46, insert_dl_on_red + boost,"
  Lampz.MassAssign(47) = l47
  Lampz.MassAssign(47) = l47a
  Lampz.Callback(47) = "DisableLighting p47, insert_dl_on_red + boost,"
  Lampz.MassAssign(48) = l48
  Lampz.MassAssign(48) = l48a
  Lampz.Callback(48) = "DisableLighting p48, insert_dl_on_white + boost,"
  Lampz.MassAssign(49) = l49
  Lampz.MassAssign(49) = l49a
  Lampz.Callback(49) = "DisableLighting p49, insert_dl_on_orange + boost,"
  Lampz.MassAssign(50) = l50
  Lampz.MassAssign(50) = l50a
  Lampz.Callback(50) = "DisableLighting p50, insert_dl_on_red + boost,"
  Lampz.MassAssign(51) = l51
  Lampz.MassAssign(51) = l51a
  Lampz.Callback(51) = "DisableLighting p51, insert_dl_on_red + boost,"
  Lampz.MassAssign(52) = l52
  Lampz.MassAssign(52) = l52a
  Lampz.Callback(52) = "DisableLighting p52, insert_dl_on_red + boost,"
  Lampz.MassAssign(53) = l53
  Lampz.Callback(53) = "DisableLighting p53, insert_dl_on_white_tri + boost,"
  Lampz.MassAssign(54) = l54
  Lampz.MassAssign(54) = l54a
  Lampz.Callback(54) = "DisableLighting p54, insert_dl_on_white_tri + boost,"
  Lampz.MassAssign(55) = l55
  Lampz.MassAssign(55) = l55a
  Lampz.Callback(55) = "DisableLighting p55, insert_dl_on_yellow_tri,"
  Lampz.MassAssign(56) = l56
  Lampz.MassAssign(56) = l56a
  Lampz.Callback(56) = "DisableLighting p56, insert_dl_on_orange,"
  Lampz.MassAssign(57) = l57
  Lampz.MassAssign(57) = l57a
  Lampz.Callback(57) = "DisableLighting p57, insert_dl_on_orange,"
  Lampz.MassAssign(58) = l58
  Lampz.MassAssign(58) = l58a
  Lampz.Callback(58) = "DisableLighting p58, insert_dl_on_orange,"
  Lampz.MassAssign(61) = l61
  Lampz.MassAssign(61) = l61a
  Lampz.MassAssign(62) = l62
  Lampz.MassAssign(62) = l62a
  Lampz.Callback(62) = "DisableLighting p62, insert_dl_on_white,"
  Lampz.MassAssign(63) = l63
  Lampz.MassAssign(63) = l63a
  Lampz.Callback(63) = "DisableLighting p63, insert_dl_on_white,"
  Lampz.MassAssign(64) = l64
  Lampz.MassAssign(64) = l64a
  Lampz.Callback(64) = "DisableLighting p64, insert_dl_on_red,"
  Lampz.MassAssign(65) = l65
  Lampz.MassAssign(65) = l65a
  Lampz.Callback(65) = "DisableLighting p65, insert_dl_on_white_tri,"
  Lampz.MassAssign(66) = l66
  Lampz.MassAssign(66) = l66a
  Lampz.Callback(66) = "DisableLighting p66, insert_dl_on_white_tri,"
  Lampz.MassAssign(67) = l67
  Lampz.MassAssign(67) = l67a
  Lampz.Callback(67) = "DisableLighting p67, insert_dl_on_white,"
  Lampz.MassAssign(68) = l68
  Lampz.MassAssign(68) = l68a
  Lampz.Callback(68) = "DisableLighting p68, insert_dl_on_red,"
  Lampz.MassAssign(69) = l69
  Lampz.MassAssign(70) = l70
  Lampz.MassAssign(70) = l70a
  Lampz.MassAssign(70) = l70a1


'Mini Playfield LED Inserts
  Lampz.MassAssign(125) = LED1       'B
  Lampz.MassAssign(125) = LED1A
  Lampz.Callback(125) = "DisableLighting pLED1, insert_dl_on_orange,"
  Lampz.MassAssign(124) = LED2       'R
  Lampz.MassAssign(124) = LED2A
  Lampz.Callback(124) = "DisableLighting pLED2, insert_dl_on_orange,"
  Lampz.MassAssign(123) = LED3       'I
  Lampz.MassAssign(123) = LED3A
  Lampz.Callback(123) = "DisableLighting pLED3, insert_dl_on_orange,"
  Lampz.MassAssign(122) = LED4       'A
  Lampz.MassAssign(122) = LED4A
  Lampz.Callback(122) = "DisableLighting pLED4, insert_dl_on_orange,"
  Lampz.MassAssign(121) = LED5       'N
  Lampz.MassAssign(121) = LED5A
  Lampz.Callback(121) = "DisableLighting pLED5, insert_dl_on_orange,"

  Lampz.MassAssign(98) = LED6        'M
  Lampz.MassAssign(98) = LED6A
  Lampz.Callback(98) = "DisableLighting pLED6, insert_dl_on_yellow,"
  Lampz.MassAssign(99) = LED7        'E
  Lampz.MassAssign(99) = LED7A
  Lampz.Callback(99) = "DisableLighting pLED7, insert_dl_on_yellow,"
  Lampz.MassAssign(100) = LED8       'G
  Lampz.MassAssign(100) = LED8A
  Lampz.Callback(100) = "DisableLighting pLED8, insert_dl_on_yellow,"

  Lampz.MassAssign(114) = LED9       'P
  Lampz.MassAssign(114) = LED9A
  Lampz.Callback(114) = "DisableLighting pLED9, insert_dl_on_white,"
  Lampz.MassAssign(89) = LED10       'E
  Lampz.MassAssign(89) = LED10A
  Lampz.Callback(89) = "DisableLighting pLED10, insert_dl_on_white,"
  Lampz.MassAssign(90) = LED11       'T
  Lampz.MassAssign(90) = LED11A
  Lampz.Callback(90) = "DisableLighting pLED11, insert_dl_on_white,"
  Lampz.MassAssign(91) = LED12       'E
  Lampz.MassAssign(91) = LED12A
  Lampz.Callback(91) = "DisableLighting pLED12, insert_dl_on_white,"
  Lampz.MassAssign(92) = LED13       'R
  Lampz.MassAssign(92) = LED13A
  Lampz.Callback(92) = "DisableLighting pLED13, insert_dl_on_white,"

  Lampz.MassAssign(108) = LED14      'L
  Lampz.MassAssign(108) = LED14A
  Lampz.Callback(108) = "DisableLighting pLED14, insert_dl_on_yellow,"
  Lampz.MassAssign(107) = LED15      'O
  Lampz.MassAssign(107) = LED15A
  Lampz.Callback(107) = "DisableLighting pLED15, insert_dl_on_yellow,"
  Lampz.MassAssign(106) = LED16      'I
  Lampz.MassAssign(106) = LED16A
  Lampz.Callback(106) = "DisableLighting pLED16, insert_dl_on_yellow,"
  Lampz.MassAssign(105) = LED17      'S
  Lampz.MassAssign(105) = LED17A
  Lampz.Callback(105) = "DisableLighting pLED17, insert_dl_on_yellow,"

  Lampz.MassAssign(84) = LED18       'C
  Lampz.MassAssign(84) = LED18A
  Lampz.Callback(84) = "DisableLighting pLED18, insert_dl_on_orange,"
  Lampz.MassAssign(83) = LED19       'H
  Lampz.MassAssign(83) = LED19A
  Lampz.Callback(83) = "DisableLighting pLED19, insert_dl_on_orange,"
  Lampz.MassAssign(82) = LED20       'R
  Lampz.MassAssign(82) = LED20A
  Lampz.Callback(82) = "DisableLighting pLED20, insert_dl_on_orange,"
  Lampz.MassAssign(81) = LED21       'I
  Lampz.MassAssign(81) = LED21A
  Lampz.Callback(81) = "DisableLighting pLED21, insert_dl_on_orange,"
  Lampz.MassAssign(97) = LED22       'S
  Lampz.MassAssign(97) = LED22A
  Lampz.Callback(97) = "DisableLighting pLED22, insert_dl_on_orange,"

  Lampz.MassAssign(199) = F29
  Lampz.MassAssign(199) = f29a
  Lampz.Callback(199) = "DisableLighting pf29, insert_dl_on_white,"
  Lampz.MassAssign(190) = F30
  Lampz.MassAssign(191) = F31
  Lampz.MassAssign(191) = f31a
  Lampz.Callback(191) = "DisableLighting pf31, insert_dl_on_white,"


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
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks

Class LampFader
  Public FadeSpeedDown(200), FadeSpeedUp(200)
  Private Lock(200), Loaded(200), OnOff(200)
  Public UseFunction
  Private cFilter
  Public UseCallback(200), cCallback(200)
  Public Lvl(200), Obj(200)
  Private Mult(200)
  Public FrameTime
  Private InitFrame
  Public Name

  Sub Class_Initialize()
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

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
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

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
  Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

  Public Sub Update1()   'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
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

  Public Sub Update2()   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
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

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
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

'******************************************************
'****  END LAMPZ
'******************************************************


'***************************************************************************************
' Hybrid code for VR, Cab, and Desktop
'***************************************************************************************

'******************************************************
'               VR Plunger Code
'******************************************************

Sub TimerVRPlunger_Timer
  If VR_Primary_plunger.Y < -90 then
       VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VR_Primary_plunger.Y = -214.1667 + (5* Plunger.Position) -20
End Sub

Dim VRThings

if VR_Room = 0 and cab_mode = 0 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in DTRails:VRThings.visible = 1:Next
Elseif VR_Room = 0 and cab_mode = 1 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in DTRails:VRThings.visible = 0:Next
Else
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  for each VRThings in VRClock:VRThings.visible = WallClock:Next
  for each VRThings in DTRails:VRThings.visible = 0:Next
'Custom Walls, Floor, and Roof
  if CustomWalls = 1 Then
    VR_Wall_Left.image = "VR_Wall_Left"
    VR_Wall_Right.image = "VR_Wall_Right"
    VR_Floor.image = "VR_Floor"
    VR_Roof.image = "VR_Roof"
  end if


  If topper = 1 Then
    Primary_topper.visible = 1
    Primary_topper.imageA = "FGTopper_ addestratore"
  Elseif topper = 2 Then
    Primary_topper.visible = 1
    Primary_topper.imageA = "FG_Logo"
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

End If


'***************************************************************************************
' CODE BELOW IS FOR THE VR CLOCK.  This was originally taken from Rascal VP9 clock table
'***************************************************************************************

Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()

    'ClockHands Below *********************************************************
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

End Sub

'***************************************************************************************
' RUBBER COLOR
'***************************************************************************************

If RubberColor = 0 Then

Rubber15.material = "RubberBlack"
Rubber013.material = "RubberBlack"
Rubber020.material = "RubberBlack"
Rubber017.material = "RubberBlack"
Rubber018.material = "RubberBlack"
Rubber011.material = "RubberBlack"
Rubber19.material = "RubberBlack"
Rubber003.material = "RubberBlack"
Rubber001.material = "RubberBlack"
Rubber16.material = "RubberBlack"
Rubber006.material = "RubberBlack"
Rubber007.material = "RubberBlack"
Rubber016.material = "RubberBlack"
Rubber005.material = "RubberBlack"
Rubber004.material = "RubberBlack"
Rubber7.material = "RubberBlack"
Rubber002.material = "RubberBlack"
Rubber008.material = "RubberBlack"
Rubber24.material = "RubberBlack"
Rubber010.material = "RubberBlack"
Rubber009.material = "RubberBlack"
Rubber012.material = "RubberBlack"
Rubber17.material = "RubberBlack"
Rubber015.material = "RubberBlack"
Rubber014.material = "RubberBlack"
LSling2.material = "RubberBlack"
LSling.material = "RubberBlack"
LSling1.material = "RubberBlack"
RSling.material = "RubberBlack"
RSling1.material = "RubberBlack"
RSling2.material = "RubberBlack"

Else

Rubber15.material = "RubberWhite"
Rubber013.material = "RubberWhite"
Rubber020.material = "RubberWhite"
Rubber017.material = "RubberWhite"
Rubber018.material = "RubberWhite"
Rubber011.material = "RubberWhite"
Rubber19.material = "RubberWhite"
Rubber003.material = "RubberWhite"
Rubber001.material = "RubberWhite"
Rubber16.material = "RubberWhite"
Rubber006.material = "RubberWhite"
Rubber007.material = "RubberWhite"
Rubber016.material = "RubberWhite"
Rubber005.material = "RubberWhite"
Rubber004.material = "RubberWhite"
Rubber7.material = "RubberWhite"
Rubber002.material = "RubberWhite"
Rubber008.material = "RubberWhite"
Rubber24.material = "RubberWhite"
Rubber010.material = "RubberWhite"
Rubber009.material = "RubberWhite"
Rubber012.material = "RubberWhite"
Rubber17.material = "RubberWhite"
Rubber015.material = "RubberWhite"
Rubber014.material = "RubberWhite"
LSling2.material = "RubberWhite"
LSling.material = "RubberWhite"
LSling1.material = "RubberWhite"
RSling.material = "RubberWhite"
RSling1.material = "RubberWhite"
RSling2.material = "RubberWhite"

End If


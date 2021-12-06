Option Explicit
Randomize

'Dungeons & Dragons (Bally 1987)
'
' Development by agentEighty6 2019 for Visual Pinball X
' Version 1.1.0

' I would like to send a very special thanks to the following.
' - mfuegemann : For being able to freely use his VP9 table.
' - Herweh : Your Atlantis table and techniques you used were invaluable in me learning VPX.
' - nFozzy : For coming up with those brilliant flipper physics subroutines
' - sheltemke : For playtesting the table and advice on tweeks and improvements.
' - brandonlaw : For help in cleaning up and improving the payfield, plastics, and ball graphics.
' - I could go on forever thanking many many more of you guys for the hard work bringing this hobby up to such an awesome level!
' - VPX development team: For making and keeping VP up to date and one of the best pinball simulators available
' - To anyone I may have forgotten, please let me know and I'll add you.

' Release notes:
' Version 1.1.2 - Controller.vbs implementation
' Version 1.1.1 - Small tweeks to improve performance.
' Version 1.1.0 - Final release (for now) with a few modifications to lighting, and table images (thanks Brandon).
'       - Left ramp physical walls (not visible) increased due to some reports of balls launching off of the table.
' Version 1.0.2 - Minor update to improve lighting and physics.
' Version 1.0.1 - Minor update to straighten out some missing testures and extraneous table elements.
' Version 1.0.0 - New VPX version built loosely from the V9 version.

'Table Components
' Layer1 - Table Mechanics
' Layer2 - Table Lamps
' Layer3 - Plastics
' Layer4 - Plastics Level 2 and Covers
' Layer5 - Walls and apron
' Layer6 - Misc, timers, testing tools, etc.
' Layer7 - Export to Blender
' Layer8 - Primitives
' Layer9 - GI Lights
' Layer10 - Flashers
' Layer11 - GI Lights for plastics

' ****************************************************
' OPTIONS
' ****************************************************

' Volume devided by - lower gets higher sound
Const VolDiv = 800    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Dim DynamicLampIntensity, DynamicFlasherIntensity, DynamicGIIntensity, LetTheBallJump, ShowBallShadow, SubstMagnaSaveButtons

' LET THE BALL JUMP A BIT
' 0 = off
' 1 to 6 = ball jump intensity (4 = default)
LetTheBallJump = 4

' FLIPPERS ALSO ACTIVATE MAGNASAVE (MAGIC GATE)
' Set to True if Your cabinet has no Magnasave Buttons. (default=False)
SubstMagnaSaveButtons = False

' DYNAMIC LAMP INTENSITY MULTIPLIER
' Numeric value to dynamically multiply every playfield lamp's intensity (decimal, default=1)
DynamicLampIntensity = 1

' DYNAMIC GI INTENSITY MULTIPLIER
' Numeric value to dynamically multiply every GI lamp's intensity (decimal, default=1)
DynamicGIIntensity = 1

' DYNAMIC FLASHER INTENSITY MULTIPLIER
' Numeric value to dynamically multiply every playfield flasher's intensity (decimal, default=1)
DynamicFlasherIntensity = 1

' SHOW BALL SHADOWS
' 0 = no ball shadows
' 1 = ball shadows are visible (default)
ShowBallShadow = 1

Dim bsTrough,bsLeftTeleporter,bsRightTeleporter,DropTargetBank

' ****************************************************
' standard definitions
' ****************************************************

Const cGameName="dungdrag"
Const UseSolenoids  = 2
Const UseLamps    = 1
Const UseSync     = 1
Const HandleMech  = 0
Const UseGI     = 0

'Standard Sounds
Const SSolenoidOn="solon"
Const SSolenoidOff="soloff"
Const SFlipperOn="fx_flipperup"
Const SFlipperOff="fx_flipperdown"
Const sCoin="fx_coin"

Const BallSize = 50
Const BallMass = 1.2

If Version < 10600 Then
  MsgBox "This table requires Visual Pinball 10.6 or newer!" & vbNewLine & "Your version: " & Replace(Version/1000,",","."), , "Dungeons & Dragons VPX"
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package."
On Error Goto 0

LoadVPM "03020000","6803.VBS",3.2

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------
SolCallback(1) = "vpmSolSound SoundFX(""bumper"",DOFContactors),"     'Top Bumper
SolCallback(2) = "vpmSolSound SoundFX(""bumper"",DOFContactors),"     'Left Bumper
SolCallback(3) = "vpmSolSound SoundFX(""bumper"",DOFContactors),"     'Right Bumper
SolCallback(4) = "vpmSolSound SoundFX(""bumper"",DOFContactors),"     'Bottom Bumper
SolCallback(5) = "vpmSolSound SoundFX(""lSling"",DOFContactors),"     'Left Slingshot
SolCallback(6) = "vpmSolSound SoundFX(""lSling"",DOFContactors),"     'Right Slingshot
SolCallback(7) = "DropTargetBank.SolDropUp"     '7 Reset Drop Targets
SolCallback(8) = "SolLeftKicker"          'Kicker Left
SolCallback(9) = "SolRightKicker"         'Kicker Right

SolCallback(10) = "SolLeftTeleporter"       'Teleporter Left
SolCallback(11) = "SolRightTeleporter"        'Teleporter Right

SolCallback(12)  = "bsTrough.SolOut"        'Kick to Playfield
'13 reserved for German use
SolCallback(14) = "bsTrough.SolIn"          'OutHole
SolCallback(15)  = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"   'Knocker

SolCallback(17) = "SolGi"             'GI Light control

SolCallback(18) = "SolFlexsaveRight"        'Flexsave Right (Right Gate)
SolCallback(20) = "SolFlexsaveLeft"         'Flexsave Left (Left Gate)

SolCallback(19) = "vpmNudge.SolGameOn"

SolCallback(sLRFlipper)="SolRightFlipper"
SolCallback(sLLFlipper)="SolLeftFlipper"


If LetTheBallJump > 6 Then LetTheBallJump = 6 : If LetTheBallJump < 0 Then LetTheBallJump = 0


Dim DesktopMode: DesktopMode = DnD.ShowDT
If DesktopMode = True Then
  'Rail1.visible=1
  'Rail2.visible=1
  TextBox001.visible=True
  TextBox002.visible=True
  Wall095.sideVisible = 0
  Wall096.sideVisible = 0
  Wall095.visible = 0
  Wall096.visible = 0
Else
  'Rail1.visible=0
  'Rail2.visible=0
  TextBox001.visible=False
  TextBox002.visible=False
  Wall095.sideVisible = 1
  Wall096.sideVisible = 1
  Wall095.visible = 1
  Wall096.visible = 1
End if


Sub DnD_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine="Dungeons & Dragons, Bally 1987"
    .HandleMechanics=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden= 1
    On Error Resume Next
    .Run
    If Err Then MsgBox Err.Description
    On Error Goto 0
    End With
  On Error Goto 0

  DisplayTimer.Enabled = DesktopMode

  vpmNudge.TiltSwitch=15
  vpmNudge.Sensitivity=3
  vpmNudge.TiltObj=Array(Bumper17,Bumper18,Bumper19,Bumper20,LeftSlingshot,RightSlingshot)

    Set Lights(1) = Lamp1
  Set Lights(17) = Lamp2
  Set Lights(33) = Lamp3

  Set Lights(2) = Lamp4
  Set Lights(18) = Lamp5
  Set Lights(34) = Lamp6

  'Set Lights(3) = Lamp
  Set Lights(19) = Lamp7
  Set Lights(35) = Lamp8

  Set Lights(4) = Lamp9
  Set Lights(20) = Lamp10
  Set Lights(36) = Lamp11

  Set Lights(5) = Lamp12
  Set Lights(21) = Lamp13
  Set Lights(37) = Lamp14

  Set Lights(6) = Lamp15
  Set Lights(22) = Lamp16
  Set Lights(38) = Lamp17

  Set Lights(7) = Lamp18
  Set Lights(23) = Lamp19
  Set Lights(39) = Lamp20

  Set Lights(8) = Lamp21
  Set Lights(24) = Lamp22
  Set Lights(40) = Lamp23

  Set Lights(9) = Lamp24
  Set Lights(25) = Lamp25
  Set Lights(41) = Lamp26

  Set Lights(10) = Lamp27
  Set Lights(26) = Lamp28
  Set Lights(42) = Lamp29

  Set Lights(11) = Lamp30
  Set Lights(27) = Lamp31
  'Set Lights(43) = Lamp ' Box Top Lights (not used)

  Set Lights(12) = Lamp33 ' Right Sling (to control flasher)
  Set Lights(28) = Lamp34 ' Sword targets (to control flasher)
  'Set Lights(44) = Lamp ' Box Top Lights (not used)

  Set Lights(13) = Lamp36 ' Backwall Lamp4 (to control flasher)
  Set Lights(29) = Lamp37 ' Shield targets (to control flasher)
  Set Lights(45) = Lamp38 ' Skillshot 2 (to control flasher)

  Set Lights(14) = Lamp39 ' Qualify Million (to control flasher)
  Set Lights(30) = Lamp40 ' Left Outlane
  Set Lights(46) = Lamp41 ' Right Outlane

  Set Lights(15) = Lamp42
  Set Lights(31) = Lamp43
  'Set Lights(47) = Lamp


  Set Lights(49) = Lamp44
  Set Lights(65) = Lamp45
  Set Lights(81) = Lamp46

  Set Lights(50) = Lamp47
  Set Lights(66) = Lamp48
  Set Lights(82) = Lamp49

  Set Lights(51) = Lamp50
  Set Lights(67) = Lamp51
  Set Lights(83) = Lamp52

  Set Lights(52) = Lamp53
  Set Lights(68) = Lamp54
  Set Lights(84) = Lamp55

  Set Lights(53) = Lamp56
  Set Lights(69) = Lamp57
  Set Lights(85) = Lamp58

  Set Lights(54) = Lamp59
  Set Lights(70) = Lamp60 ' Bumper1 (Top)
  Set Lights(86) = Lamp61 ' Bumper2 (Left)

  Set Lights(55) = Lamp62 ' Bumper3 (Right)
  Set Lights(71) = Lamp63 ' Bumper20 (bottom)
  Set Lights(87) = Lamp64 ' Backwall Lamp2 (to control flasher)

  Set Lights(56) = Lamp65 ' Backwall Lamp5 (to control flasher)

  Set Lights(91)= Lamp66 ' Left Sling (to control flasher)

  'Set Lights(60) = Lamp ' Box Top Lights (not used)
  'Set Lights(76) = Lamp ' Box Top Lights (not used)
  'Set Lights(92) = Lamp ' Box Top Lights (not used)

  Set Lights(61) = Lamp70 ' Dust targets (to control flasher)
  Set Lights(77) = Lamp71 ' Skillshot 1 (to control flasher)
  Set Lights(93) = Lamp72 ' Backwall Lamp6 (to control flasher)

  Set Lights(62) = Lamp73 ' Skillshot 3 (to control flasher)
  Set Lights(78) = Lamp74 ' Backwall Lamp1 (far Left) (to control flasher)
  Set Lights(94) = Lamp75 ' Backwall Lamp3 (to control flasher)

  Set Lights(63) = Lamp76 ' Restore

' Unused lamps
  'Set Lights(57) = Laptest002
  'Set Lights(58) = Laptest003
  'Set Lights(59) = Laptest004
  'Set Lights(72) = Laptest009
  'Set Lights(73) = Laptest010
  'Set Lights(74) = Laptest011
  'Set Lights(75) = Laptest012
  'Set Lights(79) = Laptest016
  'Set Lights(88) = Laptest018
  'Set Lights(89) = Laptest019
  'Set Lights(90) = Laptest020

  SetDynamicLampIntensity
  SetDynamicFlasherIntensity

  Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,48,47,46,0,0,0,0
    bsTrough.InitKick BallRelease,90,5
    bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors),SoundFX("solenoid",DOFContactors)
    bsTrough.Balls=3

  Set bsLeftTeleporter=New cvpmBallStack
    bsLeftTeleporter.InitSw 0,32,0,0,0,0,0,0
    bsLeftTeleporter.InitKick LeftTeleporter,220,5
    bsLeftTeleporter.InitExitSnd SoundFX("ballrel",DOFContactors),SoundFX("solenoid",DOFContactors)

  Set bsRightTeleporter=New cvpmBallStack
    bsRightTeleporter.InitSw 0,40,0,0,0,0,0,0
    bsRightTeleporter.InitKick RightTeleporter,180,5
    bsRightTeleporter.InitExitSnd SoundFX("ballrel",DOFContactors),SoundFX("solenoid",DOFContactors)

  set DropTargetBank = new cvpmDropTarget
    DropTargetBank.InitDrop Array(DropTargetTop,DropTargetMiddle,DropTargetBottom), Array(35,34,33)
    DropTargetBank.InitSnd SoundFX("droptarget",DOFDropTargets),SoundFX("droptargetreset",DOFContactors)

  pMillionRamp.rotx = -24
  RampIsUp = True         'Ramp starts up
  RampTimer.enabled = True
  MillionRamp.collidable = False

  LeftTeleporter.enabled = True   'Start teleporter blocks up
  LeftTeleporterBlockTimer.enabled = True

  RightTeleporter.enabled = True
  RightTeleporterBlockTimer.enabled = True

End Sub


' ****************************************************
' keys
' ****************************************************
Sub DnD_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then Plunger.PullBack
  if SubstMagnaSaveButtons = False then
    If keycode = LeftMagnaSave Then Controller.Switch(5) = True  ' LeftMagnaSave
    If keycode = RightMagnaSave Then Controller.Switch(7) = True  ' RightMagnaSave
  else
    If keycode = LeftFlipperKey Then Controller.Switch(5) = True  ' LeftMagnaSave
    If keycode = RightFlipperKey Then Controller.Switch(7) = True  ' RightMagnaSave
  end if
  If keycode = LeftFlipperKey Then lfpress = 1
  If keycode = RightFlipperKey Then rfpress = 1

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub DnD_KeyUp(ByVal keycode)
    If keycode = PlungerKey Then Plunger.Fire
  if SubstMagnaSaveButtons = False then
    If keycode = LeftMagnaSave Then Controller.Switch(5) = False  ' LeftMagnaSave
    If keycode = RightMagnaSave Then Controller.Switch(7) = False  ' RightMagnaSave
  else
    If keycode = LeftFlipperKey Then Controller.Switch(5) = False  ' LeftMagnaSave
    If keycode = RightFlipperKey Then Controller.Switch(7) = False  ' RightMagnaSave
  end if
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
    If vpmKeyUp(keycode) Then Exit Sub
End Sub


' ****************************************************
' flipper subs
' ****************************************************
Sub SolLeftFlipper(Enabled)
  If Enabled Then
    LF.fire
    PlaySound SoundFX("fx_flipperup",DOFFlippers)
    Else
    PlaySound SoundFX("fx_flipperdown",DOFFlippers)
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRightFlipper(Enabled)
  If Enabled Then
    RF.fire
    PlaySound SoundFX("fx_flipperup",DOFFlippers)
    URightFlipper.RotateToEnd
    Else
    PlaySound SoundFX("fx_flipperdown",DOFFlippers)
        RightFlipper.RotateToStart
    URightFlipper.RotateToStart
    End If
End Sub


Sub Drain_Hit()
  vpmtimer.pulsesw 8
  bsTrough.AddBall Me
  playsound "drain5"
End Sub


Sub SolLeftKicker(enabled)
  if enabled then
    LeftKicker.kick 0,55
  end if
End Sub

Sub SolRightKicker(enabled)
  if enabled then
    RightKicker.kick 0,55
  end if
End Sub

Sub SolLeftTeleporter(enabled)
  if enabled then
    LeftTeleporter.enabled = True
    LeftTeleporterBlockTimer.enabled = True
    LeftGate.Collidable = True
    bsLeftTeleporter.ExitSol_on
  end if
End Sub

Sub SolRightTeleporter(enabled)
  if enabled then
    MillionRamp.collidable = False
    RampIsUp = True
    RampTimer.enabled = True
    RightTeleporter.enabled = True
    RightTeleporterBlockTimer.enabled = True
    bsRightTeleporter.ExitSol_On
  end if
End Sub

Sub SolFlexsaveRight(enabled)
  if enabled then
    RightMagic.rotatetoend
  else
    RightMagic.rotatetostart
  end if
End Sub

Sub SolFlexsaveLeft(enabled)
  if enabled then
    LeftMagic.rotatetoend
  else
    LeftMagic.rotatetostart
  end if
End Sub


'---------------------------------
'------  Switch Assignment  ------
'---------------------------------
sub Dust1_hit:StandupTargetHit:vpmtimer.pulsesw 1:End Sub     'Dust Target 1
sub Dust2_hit:StandupTargetHit:vpmtimer.pulsesw 2:End Sub     'Dust Target 2
sub Dust3_hit:StandupTargetHit:vpmtimer.pulsesw 3:End Sub     'Dust Target 3
sub Dust4_hit:StandupTargetHit:vpmtimer.pulsesw 4:End Sub     'Dust Target 4
'5  Cabinet Left - handled elsewhere
'6  Cabinet Credit
'7  Cabinet Right - handled elsewhere
'8  Outhole - handled elsewhere
'9  Coins Right (Door) - handled elsewhere
'10 Coins Left (Door) - handled elsewhere
'11 Coins Middle (Door) - handled elsewhere
Sub LeftInlane_Hit:controller.switch(12) = True:End Sub     'Left Lane 12a
Sub LeftInlane_Unhit:controller.switch(12) = False:End Sub
Sub LeftOutlane_Hit:controller.switch(12) = True:End Sub      'Left Lane 12b
Sub LeftOutlane_Unhit:controller.switch(12) = False:End Sub
Sub RightInlane_Hit:controller.switch(13) = True:End Sub      'Right Lane 13a
Sub RightInlane_Unhit:controller.switch(13) = False:End Sub
Sub RightOutlane_Hit:controller.switch(13) = True:End Sub     'Right Lane 13b
Sub RightOutlane_Unhit:controller.switch(13) = False:End Sub
'14 Slam Tilt
'15 Tilt - handled elsewhere
Sub Rebound_Hit:controller.switch(16) = True:End Sub      'Rebound 16
Sub Rebound_Unhit:controller.switch(16) = False:End Sub
Sub Bumper17_Hit:vpmtimer.pulsesw 17:End Sub      'Top Bumper
Sub Bumper18_Hit:vpmtimer.pulsesw 18:End Sub      'Left Bumper
Sub Bumper19_Hit:vpmtimer.pulsesw 19:End Sub      'Right Bumper
Sub Bumper20_Hit:vpmtimer.pulsesw 20:End Sub      'Bottom Bumper
Sub LeftSlingshot_Slingshot:vpmtimer.pulsesw 21:End Sub     'Left Slingshot
Sub RightSlingshot_Slingshot:vpmtimer.pulsesw 22:End Sub      'Right Slingshot
Sub Trigger23_Hit:controller.switch(23) = True:End Sub    'Dragon Lair Left
Sub Trigger23_Unhit:controller.switch(23) = False:End Sub
Sub Trigger24_Hit:controller.switch(24) = True:End Sub    'Dragon Lair Right
Sub Trigger24_Unhit:controller.switch(24) = False:End Sub
Sub Shield25_hit:StandupTargetHit:vpmtimer.pulsesw 25:End Sub     'Shield Target 1
Sub Shield26_hit:StandupTargetHit:vpmtimer.pulsesw 26:End Sub     'Shield Target 2
Sub Shield27_hit:StandupTargetHit:vpmtimer.pulsesw 27:End Sub     'Shield Target 3
Sub Sword28_hit:StandupTargetHit:vpmtimer.pulsesw 28:End Sub      'Sword Target 1
Sub Sword29_hit:StandupTargetHit:vpmtimer.pulsesw 29:End Sub      'Sword Target 2
Sub Sword30_hit:StandupTargetHit:vpmtimer.pulsesw 30:End Sub      'Sword Target 3
                          '31 Teleporter Left Empty
Sub LeftTeleporter_Hit                '32 Teleporter Left Loaded
  LeftTeleporter.enabled = False
  LeftTeleporterBlocktimer.enabled = True
  bsLeftTeleporter.addball me
End Sub

' May need to combine these
Sub DropTargetBottom_Hit:vpmTimer.pulseSwitch(33),0,"":End Sub      'Drop Target Bottom
Sub DropTargetMiddle_Hit:vpmTimer.pulseSwitch(34),0,"":End Sub      'Drop Target Middle
Sub DropTargetTop_Hit:vpmTimer.pulseSwitch(35),0,"":End Sub       'Drop Target Top
Sub DropTargetBottom_Hit:DropTargetHit:DropTargetBank.Hit 3:DropTargetBottom.isdropped=true:End Sub       'Drop Target Bottom
Sub DropTargetMiddle_Hit:DropTargetHit:DropTargetBank.Hit 2:DropTargetMiddle.isdropped=true:End Sub       'Drop Target Middle
Sub DropTargetTop_Hit:DropTargetHit:DropTargetBank.Hit 1:DropTargetTop.isdropped=true:End Sub       'Drop Target Top

Sub Skill36_Hit:controller.switch(36) = True:End Sub      'Skill 1 Bottom
Sub Skill36_Unhit:controller.switch(36) = False:End Sub
Sub Skill37_Hit:controller.switch(37) = True:End Sub      'Skill 2 Middle
Sub Skill37_Unhit:controller.switch(37) = False:End Sub
Sub Skill38_Hit:controller.switch(38) = True:End Sub      'Skill 3 Top
Sub Skill38_Unhit:controller.switch(38) = False:End Sub
                          '39 Teleporter Right Empty
Sub RightTeleporter_Hit               '40 Teleporter Right Loaded
  MillionRamp.collidable = True
  RampIsUp = False
  RampTimer.enabled = True
  RightTeleporter.enabled = False
  RightTeleporterBlockTimer.enabled = True
  bsRightTeleporter.addball me
End Sub
Sub Trigger41_Hit:controller.switch(41) = True:End Sub      'Left Return Lane
Sub Trigger41_Unhit:controller.switch(41) = False:End Sub
Sub Trigger42_Hit:controller.switch(42) = True:End Sub      'Level Switch
Sub Trigger42_Unhit:controller.switch(42) = False:End Sub
Sub Trigger43_Hit:controller.switch(43) = True:End Sub      'Million Switch
Sub Trigger43_Unhit:controller.switch(43) = False:End Sub
Sub Trigger44_Hit:controller.switch(44) = True:End Sub      'Restore Weapons
Sub Trigger44_Unhit:controller.switch(44) = False:End Sub
'45 not used
'46 Outhole 1 Left - handled elsewhere
'47 Outhole 2 Middle - handled elsewhere
'48 Outhole 3 Right - handled elsewhere


Dim RampIsUp
Sub RampTimer_Timer
  if RampIsUp then      'up to 0
    pMillionRamp.rotx = pMillionRamp.rotx + 1
    if pMillionRamp.rotx >= 0 then
      RampTimer.enabled = False
      pMillionRamp.rotx = 0
    end if
  else            'down to -23
    pMillionRamp.rotx = pMillionRamp.rotx - 1
    if pMillionRamp.rotx <= -23 then
      RampTimer.enabled = False
      pMillionRamp.rotx = -23
    end if
  end if
End Sub


Sub RightTeleporterBlockTimer_Timer
  if RightTeleporter.Enabled then     'up to 0
    RightTeleporterBlock.TransZ = RightTeleporterBlock.TransZ + 1
    if RightTeleporterBlock.TransZ >= 0 then
      RightTeleporterBlockTimer.enabled = False
      RightTeleporterBlock.TransZ = 0
    end if
  else            'down to -60
    RightTeleporterBlock.TransZ = RightTeleporterBlock.TransZ - 1
    if RightTeleporterBlock.TransZ <= -60 then
      RightTeleporterBlockTimer.enabled = False
      RightTeleporterBlock.TransZ = -60
    end if
  end if
End Sub


Sub LeftTeleporterBlockTimer_Timer
  if LeftTeleporter.Enabled then      'up to 0
    LeftTeleporterBlock.TransZ = LeftTeleporterBlock.TransZ + 1
    if LeftTeleporterBlock.TransZ >= 0 then
      LeftTeleporterBlockTimer.enabled = False
      LeftGate.Collidable = True
      LeftTeleporterBlock.TransZ = 0
    end if
  else            'down to -60
    LeftTeleporterBlock.TransZ = LeftTeleporterBlock.TransZ - 1
    if LeftTeleporterBlock.TransZ <= -60 then
      LeftTeleporterBlockTimer.enabled = False
      LeftGate.Collidable = False
      LeftTeleporterBlock.TransZ = -60
    end if
  end if
End Sub


' *********************************************************************
' digital display
' *********************************************************************

Dim Digits(28)
Digits(0) = Array(a00,a01,a02,a03,a04,a05,a06,a07,a08)
Digits(1) = Array(a10,a11,a12,a13,a14,a15,a16,a17,a18)
Digits(2) = Array(a20,a21,a22,a23,a24,a25,a26,a27,a28)
Digits(3) = Array(a30,a31,a32,a33,a34,a35,a36,a37,a38)
Digits(4) = Array(a40,a41,a42,a43,a44,a45,a46,a47,a48)
Digits(5) = Array(a50,a51,a52,a53,a54,a55,a56,a57,a58)
Digits(6) = Array(a60,a61,a62,a63,a64,a65,a66,a67,a68)

Digits(7) = Array(b00,b01,b02,b03,b04,b05,b06,b07,b08)
Digits(8) = Array(b10,b11,b12,b13,b14,b15,b16,b17,b18)
Digits(9) = Array(b20,b21,b22,b23,b24,b25,b26,b27,b28)
Digits(10)  = Array(b30,b31,b32,b33,b34,b35,b36,b37,b38)
Digits(11)  = Array(b40,b41,b42,b43,b44,b45,b46,b47,b48)
Digits(12)  = Array(b50,b51,b52,b53,b54,b55,b56,b57,b58)
Digits(13)  = Array(b60,b61,b62,b63,b64,b65,b66,b67,b68)

Digits(14)  = Array(c00,c01,c02,c03,c04,c05,c06,c07,c08)
Digits(15)  = Array(c10,c11,c12,c13,c14,c15,c16,c17,c18)
Digits(16)  = Array(c20,c21,c22,c23,c24,c25,c26,c27,c28)
Digits(17)  = Array(c30,c31,c32,c33,c34,c35,c36,c37,c38)
Digits(18)  = Array(c40,c41,c42,c43,c44,c45,c46,c47,c48)
Digits(19)  = Array(c50,c51,c52,c53,c54,c55,c56,c57,c58)
Digits(20)  = Array(c60,c61,c62,c63,c64,c65,c66,c67,c68)

Digits(21)  = Array(d00,d01,d02,d03,d04,d05,d06,d07,d08)
Digits(22)  = Array(d10,d11,d12,d13,d14,d15,d16,d17,d18)
Digits(23)  = Array(d20,d21,d22,d23,d24,d25,d26,d27,d28)
Digits(24)  = Array(d30,d31,d32,d33,d34,d35,d36,d37,d38)
Digits(25)  = Array(d40,d41,d42,d43,d44,d45,d46,d47,d48)
Digits(26)  = Array(d50,d51,d52,d53,d54,d55,d56,d57,d58)
Digits(27)  = Array(d60,d61,d62,d63,d64,d65,d66,d67,d68)

Sub DisplayTimer_Timer()
    Dim chgLED, ii, num, chg, stat, obj
  chgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(chgLED) Then
    If DesktopMode Then
      For ii = 0 To UBound(chgLED)
        num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
        if (num < 32) then
          For Each obj In Digits(num)
            If chg And 1 Then obj.State = stat And 1
            chg = chg\2 : stat = stat\2
          Next
        End If
      Next
    End If
    End If
End Sub


Sub SolGi(enabled)
    If enabled Then
        GiON
    Else
        GiOFF
    End If
End Sub

Sub GiON
  dim xx
    For each xx in GILights
        xx.State = LightStateOn
    xx.intensity = xx.intensity * DynamicGIIntensity * 1.5
    Next
End Sub

Sub GiOFF
  dim xx
    For each xx in GILights
        xx.State = LightStateOff
    Next
End Sub


Sub SetDynamicLampIntensity
  dim xx
  For each xx in AllLights
    xx.intensity = xx.intensity * DynamicLampIntensity
  Next
End Sub


Sub SetDynamicFlasherIntensity
  dim xx
  For each xx in Flashers
    xx.Opacity = xx.Opacity * DynamicFlasherIntensity * 1.5
  Next
End Sub


' Flasher routine controlled by Lights
Sub LSampleTimer_Timer()
  'Bright playfield lights
  Flasher33.visible = (lamp33.state = LightStateOn)
  Flasher34.visible = (lamp34.state = LightStateOn)
  Flasher37.visible = (lamp37.state = LightStateOn)
  Flasher38.visible = (lamp38.state = LightStateOn)
  Flasher39.visible = (lamp39.state = LightStateOn)
  Flasher66.visible = (lamp66.state = LightStateOn)
  Flasher70.visible = (lamp70.state = LightStateOn)
  Flasher71.visible = (lamp71.state = LightStateOn)
  Flasher73.visible = (lamp73.state = LightStateOn)


  'Back panel flashers
  Lamp36a.state = Lamp36.state
  Flasher36.visible = (Lamp36.state = LightStateOn)
  Flasher36a.visible = (Lamp36.state = LightStateOn)
  Flasher36b.visible = (Lamp36.state = LightStateOn)
  MysticalFlasherDomeLit.Visible = (Lamp36.state = LightStateOn)
  MysticalFlasherDome.Visible = (Lamp36.state <> LightStateOn)

  Flasher64.visible = (lamp64.state = LightStateOn)
  Flasher64a.visible = (lamp64.state = LightStateOn)
  Flasher64b.visible = (lamp64.state = LightStateOn)
  LeftDragonFlasherDomeLit.Visible = (lamp64.state = LightStateOn)
  LeftDragonFlasherDome.Visible = (lamp64.state <> LightStateOn)

  Flasher65.visible = (lamp65.state = LightStateOn)
  Flasher65a.visible = (lamp65.state = LightStateOn)
  Flasher65b.visible = (lamp65.state = LightStateOn)
  RightDragonFlasherDomeLit.Visible = (Lamp65.state = LightStateOn)
  RightDragonFlasherDome.Visible = (lamp65.state <> LightStateOn)

  Flasher72.visible = (lamp72.state = LightStateOn)
  Flasher74.visible = (lamp74.state = LightStateOn)
  Flasher75.visible = (lamp75.state = LightStateOn)


  'Bumper cap and center lights
  Lamp60a.state = Lamp60.state
  Lamp61a.state = Lamp61.state
  Lamp62a.state = Lamp62.state
  Lamp63a.state = Lamp63.state

  'Outlane lights
  Lamp40a.state = Lamp40.state
  Lamp41a.state = Lamp41.state
End Sub


'******************************************************
'     STEPS 2-4 (FLIPPER POLARITY SETUP
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

'InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    x.TimeDelay = 44
  Next

  'rf.report "Velocity"  1990s
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.2,   1.07
  addpt "Velocity", 2, 0.41, 1.05
  addpt "Velocity", 3, 0.44, 1
  addpt "Velocity", 4, 0.65,  1.0'0.982
  addpt "Velocity", 5, 0.702, 0.968
  addpt "Velocity", 6, 0.95,  0.968
  addpt "Velocity", 7, 1.03,  0.945

  'rf.report "Polarity" 1990s
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
  LF.EndPoint = EndPointLp
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub



'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() :  LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() :  RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub


'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

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
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 30 Then
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


rightFlipper.timerinterval=1
rightflipper.timerenabled=True

sub RightFlipper_timer()

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

end sub

dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSAnew
dim FStrength, Frampup, FElasticity, EOSRampup, SOSRampup
dim RFEndAngle, LFEndAngle, LFCount, RFCount, LiveCatch

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
FStrength = LeftFlipper.strength
Frampup = leftFlipper.rampup
FElasticity = leftFlipper.elasticity
EOSTnew = 1.0 'FEOST
EOSAnew = 0.2
EOSRampup = 1.5
SOSRampup = 8.5
LiveCatch = 8

LFEndAngle = leftflipper.endangle
RFEndAngle = rightflipper.endangle


'******************************************************
'   HELPER FUNCTIONS
'******************************************************


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

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
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



' *********************************************************************
' some special physics behaviour
' *********************************************************************
' target or rubber post is hit so let the ball jump a bit
Sub DropTargetHit()
  DropTargetSound
  TargetHit
End Sub

Sub StandupTargetHit()
  StandUpTargetSound
  TargetHit
End Sub

Sub TargetHit()
    ActiveBall.VelZ = ActiveBall.VelZ * (0.8 + (Rnd()*LetTheBallJump + Rnd()*LetTheBallJump + 1) / 6)
End Sub

Sub RubberPostHit()
  ActiveBall.VelZ = ActiveBall.VelZ * (0.9 + (Rnd()*(LetTheBallJump-1) + Rnd()*(LetTheBallJump-1) + 1) / 6)
End Sub

Sub RubberRingHit()
  ActiveBall.VelZ = ActiveBall.VelZ * (0.8 + (Rnd()*(LetTheBallJump-1) + 1) / 6)
End Sub


' *********************************************************************
' more realtime sounds
' *********************************************************************
' ball collision sound
Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound "fx_collide", 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' rubber hit sounds
Sub RubberWalls_Hit(idx)
    PlaySoundAtBallVol "fx_rubber_hit_" & Int(Rnd*3)+1, 10
  RubberRingHit
End Sub

Sub RubberPosts_Hit(idx)
    PlaySoundAtBallVol "fx_rubber_hit_" & Int(Rnd*3)+1, 10
  RubberPostHit
End Sub

' metal hit sounds
Sub MetalWalls_Hit(idx)
  PlaySoundAtBallAbsVol "fx_metalhit" & Int(Rnd*3), Minimum(Vol(ActiveBall),0.5)
End Sub

' plastics hit sounds
Sub Plastics_Hit(idx)
  PlaySoundAtBallAbsVol "fx_ball_hitting_plastic", Minimum(Vol(ActiveBall),0.5)
End Sub

' gates sound
Sub Gates_Hit(idx)
  GateSound
End Sub

' sound at ramp rubber at the diverter
Sub RampRubber_Hit()
  PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 10
End Sub


'
' *********************************************************************
' sound stuff
' *********************************************************************
Sub RollOverSound()
  PlaySoundAtVolPitch "fx_rollover", ActiveBall, 0.02, .25
End Sub
Sub DropTargetSound()
  PlaySoundAtVolPitch SoundFX("fx_droptarget",DOFTargets), ActiveBall, 2, .25
End Sub
Sub StandUpTargetSound()
  PlaySoundAtVolPitch SoundFX("fx_target",DOFTargets), ActiveBall, 2, .25
End Sub
Sub GateSound()
  PlaySoundAtVolPitch SoundFX("fx_gate",DOFContactors), ActiveBall, 0.02, .25
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


'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

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

  ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

'***Ball Drop Sounds***

    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySound "ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If

    Next
End Sub


'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
' ninuzzu's FLIPPER COVERS AND SHADOWS
'*****************************************

Dim BallShadow : BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)

sub GraphicsTimer_Timer()
  Dim ii

  ' maybe show ball shadows
  If ShowBallShadow <> 0 Then
    Dim BOT
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT) < tnob - 1 Then
      For ii = UBound(BOT) + 1 To tnob - 1
        If BallShadow(ii).Visible Then BallShadow(ii).Visible = False
      Next
    End If
    ' render the shadow for each ball
    For ii = 0 to UBound(BOT)
      If BOT(ii).X < DnD.Width/2 Then
        BallShadow(ii).X = ((BOT(ii).X) - (Ballsize/6) + ((BOT(ii).X - (DnD.Width/2))/7)) + 6
      Else
        BallShadow(ii).X = ((BOT(ii).X) + (Ballsize/6) + ((BOT(ii).X - (DnD.Width/2))/7)) - 6
      End If
      BallShadow(ii).Y = BOT(ii).Y + 12
'     If TroughBalls <= 0 Then
        BallShadow(ii).Visible = True '(BOT(ii).Z > 20)
'     Else
'       BallShadow(ii).Visible = (BOT(ii).Z > 20 And ii >= TroughBalls And ii <> lockedBallID)
'     End If
    Next
  End If

  ' Flipper Shadows
  LeftFlipperShadow.RotY = LeftFlipper.CurrentAngle - 90
  RightFlipperShadow.RotY = RightFlipper.CurrentAngle - 90
  URightFlipperShadow.RotY = URightFlipper.CurrentAngle - 90

  ' Flipper Covers
  LFLogo.RotY = LeftFlipper.CurrentAngle + 240
  RFLogo.RotY = RightFlipper.CurrentAngle + 120
  RF2Logo.RotY = URightFlipper.CurrentAngle + 150

End Sub


Function GetBallID(actBall)
  Dim b, BOT, ret
  ret = -1
  BOT = GetBalls
  For b = 0 to UBound(BOT)
    If actBall Is BOT(b) Then
      ret = b : Exit For
    End If
  Next
  GetBallId = ret
End Function

''''''''''''''''''''''''
''' Test Kicker
''''''''''''''''''''''''
Sub TestKickerIn_Hit
  TestKickerIn.DestroyBall
  TestKickerOut.CreateBall
  'TestKickerOut.Kick Angle,velocity
  TestKickerOut.Kick 40, 40
End Sub
Sub TestKickerLoop_Hit
  TestKickerLoop.DestroyBall
  TestKickerOut.CreateBall
  'TestKickerOut.Kick Angle,velocity
  TestKickerOut.Kick 40, 40
End Sub
Sub TestKickerOut_Hit
  TestKickerOut.Kick 40, 40
End Sub

''''''''''''''''''''''''''


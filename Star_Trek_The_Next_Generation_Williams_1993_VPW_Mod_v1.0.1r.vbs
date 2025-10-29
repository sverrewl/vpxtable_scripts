'*********************************************************************
'   Star Trek - The Next Generation (1993 Williams)
'*********************************************************************
'Concept by:    Steve Ritchie
'Design by:     Steve Ritchie, Dwight Sullivan, Greg Freres
'Art by:      Greg Freres
'Dots/Animation by: Scott Slomiany, Eugene Geer
'Mechanics by:    Carl Biagi
'Music & Sound by:  Dan Forden
'Software by:   Dwight Sullivan, Matt Coriale
'*********************************************************************
'Recreated for Visual Pinball by Knorr and Clark Kent
'*********************************************************************

'VPW Mod v1.0
'============
'Project Lead - fluffhead35
'VR Holodeck - Rawd/Steely
'VR Work/Backglass - Sixtoe/Leojreimroc
'Tweaks - ClarkKent, Apophis, Sixtoe

'Changelog at bottom

'v1.0.1 - Adjusted the neutral zone target, changed the vr wall heights to be correct
'v1.0.1a - disabled borgramp1a and set detarampendborg to collidable.
'v1.0.1b - adjusted the neutralzone sounds so that they will be heard. Moved sw48 on borgramp1 so that it see hit value, adjusted cannonlBall.z and cannonRBall.z to fix ball hitting habbit trail.
'v1.0.1c - added option to show cab rails in cabnet mode.  Added in staged flippers mod
'v1.0.1d - moved ramp3 to right a little to line up better with mission start scope.  Replaced the UpdateLamps subroutine with nFozzys versios to fix kickback light
'v1.0.1e - removed upper flipper from keyup and keydown as conflicted with selenoid calls.
'v1.0.1f - Changed the Elaciticy of netural zone target to .2 and elacticiy falloff to 0 for all targets
'v1.0.1g -  set T27P2 to be collidable so that ball hits target and gets into neutral zone
'v1.0.1h - scoopleft2 got moved to wrong position.  Fixed this issue.
'v1.0.1i - Added leftramp triger to start ramp sound.  Adjusting the ramp exit volumes some to make it better.  Added new logic for rampRoll to equalize sound and do a playsoundatvolumelevel call logic.
'        - Added left and right wire ramp gun ramp sounds.
'V1.0.1j - Replaced spinners that were used for gate animations with actual gates.
'v1.0.1k - fixed nuturalzone sound, changed sw27 to star and realligned, fixed issue with ramproll sound with defualt volume level not working right.  set some triggers to hidden.
'v1.0.1l - Added in code to try and help resolve vr dmd issue.
'v1.0.1m - nfozzy borg lamp fix, rawd vr reflection amount tweak, six klingon bird of prey tweak
'v1.0.1n - adjusted physics on lower sling posts.  Adjusted sounds lever for middle wire ramp. Put wirerampoff when entering borg ship.  Adjusted wireramp sound for vpw option.
'v1.0.1o - Addeed Lockdown Firebutton. Removed unused sounds  Commented out Debug statements.  Made OriginalSounds the default.
'v1.0.1p - fixed missing variable and addede spacebackground.
'v1.0.1q - fixed center Borg light

Option Explicit
Randomize

Const BallSize = 50         'Ball size must be 50
Const BallMass = 1          'Ball mass must be 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode
DesktopMode = Table1.ShowDT
Dim UseVPMDMD
If RenderingMode = 2 Then
  UseVPMDMD = True
Else
  UseVPMDMD = DesktopMode
End If
Const UseVPMModSol = 1

'VR room Dims..
dim VRRoom
dim VRRoomChoice
dim ReflectDMD
dim GlassScratches
dim HDprim
Dim DoorsOpen
DoorsOpen = False
Dim HoloPosition
HoloPosition = 0
DIM VRThings
Dim DoorMoving
DoorMoving = false
HD_HallLites.disablelighting = 4
'End VR room Dims...

' VR Options **************

VRRoomChoice = 1      ' 1 = Holodeck   2 = Minimal Room  3 = Ultra Minimal room
ReflectDMD = 1      ' 0 = DMD Reflection OFF  1 = DMD Reflection ON  (This will only run if VRRoom = 1 or 2)
GlassScratches = 1    ' 0 = GlassScratches OFF  1 = GlassScratches ON  (This will only run if VRRoom = 1 or 2)

'End VR Options **********

'----- General Sound Options -----
Const VolumeDial = 0.8        'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

Const OriginalWireRampSounds = 1    'Set to 1 to Use the Original Wire Ramp Sounds from ClarkKents Machine else set to 0 for VPW Wire Loop Sounds
Const LowerBallLaunchSound = 0      'If Ball Launch has too much bass for ssf then set this to 1.  Only works if OriginalWireRampSounds = 0

Const ShowCabRails = 0              ' If set to 1 then cabrails will be shown in cabinet mode

Const StagedFlipperMod = 0        '0 = not staged, 1 - staged (dual leaf switches)

'**********
'***Mods***
'**********

'1 for ON, 0 for OFF

InlanesMod = 1        'with this mod the inlanes are about 1cm/0,39inch longer
BorgMod = 0         'changes the Original Model to a Custom Model

RGBMod = 0          'RGB
  RGBBumpers = 1
  RGBArrows = 0

LaserMod = 0        'Activate Lasers on the Cannons
  LaserColor = 1      '1 = RED, 2 = GREEN, 3 = BLUE
  LaserType = 1     '1 = Fast, turns off after the ball is shoot, 2 = Slow, turns off after the cannon is back initial_pos

'********************
'Standard definitions
'********************

LoadVPM "01560000", "wpc.VBS", 3.36

'***ROM***ROM***ROM***ROM***
Const cGameName = "sttng_l7"
'***************************

Const UseSolenoids = 2
Const UseLamps = 0
Const HandleMech = 0
Const SSolenoidOn = "SolOn"
Const SSolenoidOff = "SolOff"
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "Coin4"
Dim bsTrough,TopDrop, LeftCannonMech, RightCannonMech, BorgLock
Dim KB, CP, UnderDiverterTop1, UnderDiverterBottom1
Dim InlanesMod, BorgMod, FlipperMod, LaserMod, LaserColor, LaserType, RGBMod, RGBBumpers, RGBArrows


Set GiCallback2 = GetRef("UpdateGI")
Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

' VR Init...
If RenderingMode = 2 Then VRRoom = VRRoomChoice Else VRRoom = 0

If VRRoom = 1 then
  for each HDPrim in Holodeck: HDPrim.visible = true: Next
  for each HDPrim in VRCab: HDPrim.visible = true: Next
  SetBackglass
  LCARSTimer.enabled = true
  RedAlertTimer.enabled = true
  if ReflectDMD = 1 then DMDReflection.visible = true
  if GlassScratches = 1 then GlassImpurities.visible = true
  Lockbar.visible = 0
  LeftSideRail.visible = 0
  RightSideRail.visible = 0
  VRRight.visible = True
  VRLeft.visible = true
  VRSignsTimer.enabled = true
  table1.PlayfieldReflectionStrength = 7
End If

if VRRoom = 2 then
  for each HDPrim in VRMinimal: HDPrim.visible = true: Next
  for each HDPrim in VRCab: HDPrim.visible = true: Next
  SetBackglass
  if ReflectDMD = 1 then DMDReflection.visible = true
  if GlassScratches = 1 then GlassImpurities.visible = true
  Lockbar.visible = 0
  LeftSideRail.visible = 0
  RightSideRail.visible = 0
  table1.PlayfieldReflectionStrength = 7
End If

if VRRoom = 3 then
  PinCab_Backbox.visible = true
  BackGlassNew.visible = true
  Pincab_Metals.visible = true
  DMD1.visible = true
  Lockbar.visible = 0
  LeftSideRail.visible = 0
  RightSideRail.visible = 0
  table1.PlayfieldReflectionStrength = 7
End If

'************
' Table Init
'************

Sub table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = DesktopMode
        '.SetDisplayPosition 0, 0, GetPlayerHWnd 'uncomment this line if you don't see the DMD
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 0 'and keep it close
        PinMAMETimer.Interval = PinMAMEInterval
        PinMAMETimer.Enabled = true
        vpmNudge.TiltSwitch = 14
        vpmNudge.Sensitivity = 2
    vpmNudge.TiltObj = Array(LeftJetBumper, RightJetBumper, BottomJetBumper, SlingShotLeft, SlingShotRight)
    End With

    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 66, 65, 64, 63, 62, 61, 0
        .InitKick BallRelease, 80, 10
        .InitExitSnd SoundFX("BallRelease", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 3 ' total = 6 - three start in other troughs
    End With

     Set TopDrop = New cvpmDropTarget
     With TopDrop
        .InitDrop sw57, 57
    .InitSnd SoundFX("Dropdown",DOFDropTargets), SoundFX("DropUp",DOFDropTargets)
     End With

     Set BorgLock = New cvpmBallStack
     With BorgLock
         .InitSw 0, 31, 0, 0, 0, 0, 0, 0
         .InitKick BorgKicker, 180, 24
         .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
         .KickForceVar = 3
         .KickAngleVar = 3
'        .Balls = 0
     End With

  BallSearch() ' update various trough switches if a ball is in them

  '**********************
  '***DesktopMode Init***
  '**********************

  If table1.ShowDt = False then
    if RenderingMode = 2 Then
      Korpus.visible = 1
      Korpus1.visible = 0
      LeftSideRail.visible = 0
      RightSideRail.visible = 0
    Else
      Korpus.visible = 0
      Korpus1.visible = 1
      if ShowCabRails Then
        LeftSideRail.visible = 1
        RightSideRail.visible = 1
      Else
        LeftSideRail.visible = 0
        RightSideRail.visible = 0
      End If
    End If
   Else
    if RenderingMode = 2 Then
      Korpus.visible = 1
      Korpus1.visible = 0
      LeftSideRail.visible = 0
      RightSideRail.visible = 0
    Else
      Korpus.visible = 1
      Korpus1.visible = 0
      LeftSideRail.visible = 1
      RightSideRail.visible = 1
    End If
  End if

  '****************************
  '***Switches/Diverter Init***
  '****************************


  Controller.Switch(127) = 1
  Controller.Switch(122) = 1
  Controller.Switch(125) = 1
  Controller.Switch(126) = 1
  DiverterFRG.isDropped = 1
  DiverterFLG.isDropped = 0
  sw57.isDropped = 1
  sw41dr.isDropped = 1
  sw35dr.isDropped = 1
  sw36dr.isDropped = 1
  sw37dr.isDropped = 1
  sw37dr1.isDropped = 0
  sw42dr.isDropped = 1
  sw43dr.isDropped = 1
  sw47dr.isDropped = 1

  '**************
  '***Mods Init***
  '**************

  If InlanesMod = 1 then
    leftinlanewall.collidable = 0
    rightinlanewall.collidable = 0
    inlanemetal.visible = 0
    leftinlanewallmod.collidable = 1
    rightinlanewallmod.collidable = 1
    inlanemetalmod.visible = 1
  Else
    leftinlanewall.collidable = 1
    rightinlanewall.collidable = 1
    inlanemetal.visible = 1
    leftinlanewallmod.collidable = 0
    rightinlanewallmod.collidable = 0
    inlanemetalmod.visible = 0
  End if


  If BorgMod = 1 then
    borgshipcustom.visible = 1
    borgshipcustomledoff.visible = 1
    borgshipcustomwiresplines.visible = 1
    deltarampendborgcustom.visible = 1

    l78a.Intensity = 30
    l78b.Intensity = 30
    l78c.Intensity = 6
    l78d.Intensity = 30
    l78e.Intensity = 10
    l78a1.State = 1
    l78b1.State = 1
    l78c1.State = 1
    l78d1.State = 1
    l78e1.State = 1

    borgshiporiginal.visible = 0
    borgledoriginal.visible = 0
    deltarampendborg.visible = 0
    l78borga.Intensity = 0
    l78borgb.Intensity = 0
    l78borgc.Intensity = 0
    l78borgd.Intensity = 0
    l78borge.Intensity = 0
    l78borga1.State = 0
    l78borgb1.State = 0
    l78borgc1.State = 0
    l78borgd1.State = 0
    l78borge1.State = 0
  Else
    borgshipcustom.visible = 0
    borgshipcustomledoff.visible = 0
    borgshipcustomwiresplines.visible = 0
    deltarampendborgcustom.visible = 0

    l78a.Intensity = 0
    l78b.Intensity = 0
    l78c.Intensity = 0
    l78d.Intensity = 0
    l78e.Intensity = 0
    l78a1.State = 0
    l78b1.State = 0
    l78c1.State = 0
    l78d1.State = 0
    l78e1.State = 0

    borgshiporiginal.visible = 1
    borgledoriginal.visible = 1
    deltarampendborg.visible = 1
    l78borga.Intensity = 15
    l78borgb.Intensity = 50
    l78borgc.Intensity = 60
    l78borgd.Intensity = 30
    l78borge.Intensity = 50
    l78borga1.State = 1
    l78borgb1.State = 1
    l78borgc1.State = 1
    l78borgd1.State = 1
    l78borge1.State = 1
  End if

  If RGBMod = 1 then RGBTimer.Enabled = 1
    If RGBBumpers = 1 And RGBMod = 1 then
      BumperCap1.Image = "bumpero_weis"
      BumperCap2.Image = "bumpero_weis"
      BumperCap3.Image = "bumpero_weis_half"
    End if


  If LaserColor = 1 then LaserL.Material = "LaserRed"
  If LaserColor = 1 then LaserL1.Material = "LaserRed1"
  If LaserColor = 1 then LaserR.Material = "LaserRed"
  If LaserColor = 1 then LaserR1.Material = "LaserRed1"

  If LaserColor = 2 then LaserL.Material = "LaserGreen"
  If LaserColor = 2 then LaserL1.Material = "LaserGreen1"
  If LaserColor = 2 then LaserR.Material = "LaserGreen"
  If LaserColor = 2 then LaserR1.Material = "LaserGreen1"

  If LaserColor = 3 then LaserL.Material = "LaserBlue"
  If LaserColor = 3 then LaserL1.Material = "LaserBlue1"
  If LaserColor = 3 then LaserR.Material = "LaserBlue"
  If LaserColor = 3 then LaserR1.Material = "LaserBlue1"

  If StagedFlipperMod = 1 Then
    keyStagedFlipperR = KeyUpperRight ' change to RightMagnaSave if you want to use MagnaSave instead
  End If
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "LeftCannonKicker"
SolCallback(2) = "RightCannonKicker"
SolCallback(3) = "UnderLeftGun"
SolCallback(4) = "UnderRightGun"
SolCallback(5) = "LeftLock"
SolCallback(6) = "AutoPlunge"
SolCallback(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallBack(8) = "Kickback"
SolCallBack(11) = "SolRelease"
SolCallBack(15) = "TopDiverter"
SolCallBack(16) = "BorgLock.SolOut"
SolCallback(17) = "LeftCannonMotor"
SolCallback(18) = "RightCannonMotor"
SolCallBack(51) = "UnderDiverterTop"
SolCallBack(52) = "UnderDiverterBottom"
SolCallBack(53) = "TopDrop.SolDropUp"
SolCallBack(54) = "TopDrop.SolDropDown"

'***Flasher***
SolModCallBack(20) = "Flash120"     'JetsFlasher
SolModCallBack(21) = "Flash121"     'RightPopperFlasher
SolModCallBack(22) = "Flash122"     'MiddleRampFlasher
SolModCallBack(23) = "Flash123"     'ShieldsFlasher
SolCallBack(24) = "SetLamp 124,"    'AutofireFlasher
SolModCallBack(25) = "Flash125"     'ExitUnGndFlasher
SolModCallBack(26) = "Flash126"     'RightBorgFlasher
SolModCallBack(27) = "Flash127"     'LeftBorgFlasher
SolModCallBack(28) = "Flash128"     'CenterBorgFlasher
SolModCallBack(55) = "Flash141"     'RomulanFlasher
SolModCallBack(56) = "Flash142"     'RightRampFlasher


Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls> 0 Then
        vpmTimer.PulseSw 67
        bsTrough.ExitSol_On
    End If
End Sub

Sub Drain_Hit() 'nf destroy excess blals
  if bsTrough.balls + BorgLock.balls + uBound(getballs)+1 > 6 then me.destroyball : Exit Sub : end if
    RandomSoundDrain Me
    bsTrough.AddBall Me
End Sub

SubwayStart
Sub SubwayStart() 'nf fast boot
  sw36.CreateSizedBallWithMass Ballsize/2, BallMass   'left gun
  sw37.CreateSizedBallWithMass Ballsize/2, BallMass 'right gun
  sw41.CreateSizedBallWithMass Ballsize/2, BallMass 'lock - this still sometimes gets kicked out on startup

  dim x : for each x in Array(sw41,sw35,sw42,sw43, sw36,sw32,sw33,sw37) ' put associated switch number in uservalue
    x.UserValue = cInt(mid(x.name, 3, 2))
  Next
  AutoPlunger.UserValue = 68 ' sw68 - autoplunger

End Sub

Sub BallSearch() ' check all these triggers and kickers and update switches accordingly
  dim x : for each x in Array(sw41,sw35,sw42,sw43, sw36,sw32,sw33,sw37,AutoPlunger)
    if x.ballcntover then controller.Switch(x.uservalue) = True
  Next
  bsTrough.Update
End Sub



Sub AutoPlunge(Enabled)
  if enabled then
    AutoPlunger.Kick 0, 280
    'PlaySoundAt SoundFX("PlungerCatapult",DOFContactors), Catapult
    SoundSaucerKick 1, Catapult
    CP = True
    Else
    Controller.Switch(68) = 0
    CP = False
  End if
End Sub
Sub AutoPlunger_Hit:Controller.Switch(68) = 1:End Sub


Sub TopDiverter(enabled)
  If Enabled Then
    diverter.rotatetoend
    topdiverterwall.isdropped = 0
    PlaySoundAt SoundFX("TopdiverterOn",DOFContactors), diverterp
  Else
    diverter.rotatetostart
    topdiverterwall.isdropped = 1
    PlaySoundAt SoundFX("TopdiverterOff",DOFContactors), diverterp
  End If
End Sub

Sub UnderDiverterTop(enabled)
  If enabled Then
    DiverterFRG.isDropped = 0
    sw37dr1.isDropped = 1
  Else
    DiverterFRG.isDropped = 1
    sw37dr1.isDropped = 0
  End If
End Sub

Sub UnderDiverterBottom(enabled)
  If enabled Then
    DiverterFLG.isDropped = 1
  Else
    DiverterFLG.isDropped = 0
  End If
End Sub


Sub BorgKicker1_Hit()
  borglock.Addball Me
End Sub

Sub KickBack(Enabled)
  If enabled Then
    KickbackPlunger.Fire
    KB = True
'   KickBackTimer.Enabled = 1
    PlaysoundAt SoundFX("Kickback", DOFContactors), Kickbackplunger
    Else
  End if
End Sub

Sub LeftCannonMotor(enabled)
  If enabled Then
    CML = True
    CannonLTimer.Enabled = 1
    PlaySoundAt SoundFX("LGunMotor",DOFGear), CannonBaseL
  Else
    CannonLTimer.Enabled = 0
    StopSound "LGunMotor"
  End If
End Sub

Sub RightCannonMotor(enabled)
  If enabled Then
    CMR = True
    CannonRTimer.Enabled = 1
    PlaySoundAt SoundFX("RGunMotor",DOFGear),CannonBaseR
  Else
    CannonRTimer.Enabled = 0
    StopSound "RGunMotor"
  End If
End Sub

Sub UnderLeftGun(Enabled)
  If enabled then
  sw36.kick 161, 50, 1.4
  PlaySoundAt SoundFX("LGunPopper",DOFContactors),sw41
  Controller.Switch(36) = 0
  Else
' Controller.Switch(36) = 0
  vpmtimer.addtimer 1000, "sw36dr.isDropped = 1'"
  End if
End Sub

Sub UnderRightGun(Enabled)
  If enabled then
  sw37.kick 199, 50, 1.4
  PlaySoundAt SoundFX("RGunPopper",DOFContactors),sw37
  Controller.Switch(37) = 0
  Else
' Controller.Switch(37) = 0
  vpmtimer.addtimer 1000, "sw37dr.isDropped = 1'"
  End if
End Sub

Sub LeftLock(Enabled)
  If enabled then
  sw41.kick 255, 50, 1.4
  PlaySoundAt SoundFX("LeftPopper",DOFContactors),sw41
  PlaySoundAt "LeftPopperOut", sw41
  Controller.Switch(41) = 0
  Else
  vpmtimer.addtimer 250, "sw41dr.isDropped = 1'"
  End if
End Sub


'*********
' Bumper
'*********

Sub LeftJetBumper_hit:vpmTimer.pulseSw 71:RandomSoundBumperTop LeftJetBumper:Me.TimerEnabled = 1:End Sub
Sub LeftJetBumper_Timer:Me.Timerenabled = 0:End Sub

Sub RightJetBumper_hit:vpmTimer.pulseSw 72:RandomSoundBumperMiddle RightJetBumper:Me.TimerEnabled = 1:End Sub
Sub RightJetBumper_Timer:Me.Timerenabled = 0:End Sub

Sub BottomJetBumper_hit:vpmTimer.pulseSw 73:RandomSoundBumperBottom BottomJetBumper:Me.TimerEnabled = 1:End Sub
Sub BottomJetBumper_Timer:Me.Timerenabled = 0:End Sub


'*********
' Switches
'*********

Sub sw15_Hit:Controller.Switch(15) = 1:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw117spinner_Spin():vpmTimer.PulseSw 117:SoundSpinner sw117spinner:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub
Sub sw48_Hit: Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw57_Hit:PlaySoundAt SoundFX ("target", DOFContactors), sw45:TopDrop.hit 1:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw76_Hit:Controller.Switch(76) = 1:End Sub
Sub sw76_UnHit:Controller.Switch(76) = 0:End Sub

Sub sw77_Hit:Controller.Switch(77) = 1:End Sub
Sub sw77_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw78_Hit:Controller.Switch(78) = 1:End Sub
Sub sw78_UnHit:Controller.Switch(78) = 0:End Sub


Sub sw25_Hit:vpmTimer.PulseSw 25: WireRampOn True : SoundPlayfieldGate :  End Sub'PlaySoundAt "gate4",sw25: End Sub 'right Ramp
Sub sw87_Hit:vpmTimer.PulseSw 87: SoundPlayfieldGate : End Sub 'PlaySoundAt "gate4", sw87: End Sub

Sub sw88_Hit:vpmTimer.PulseSw 88: SoundPlayfieldGate : End Sub 'PlaySoundAt "gate4", sw88: End Sub 'left Ramp
Sub sw83_Hit:vpmTimer.PulseSw 83: sw83wire.RotX = 110:Me.TimerEnabled = 1:End Sub
Sub sw83_Timer: sw83wire.RotX = 90: Me.TimerEnabled = 0: End Sub

Sub sw23_Hit:vpmTimer.PulseSw 23: sw23wire.RotX = 110:Me.TimerEnabled = 1:End Sub 'middle Ramp
Sub sw23_Timer: sw23wire.RotX = 90: Me.TimerEnabled = 0: End Sub


'GunSwitches

Sub sw32_Hit:Controller.Switch(32) = 1:End Sub
'Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:Controller.Switch(32) = 0:sw36dr.isDropped = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:End Sub
'Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw37_Hit:Controller.Switch(37) = 1:Controller.Switch(33) = 0:sw37dr.isDropped = 0:End Sub

'LeftLock

Sub sw41_Hit:Controller.Switch(41) = 1:sw41dr.isDropped = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:sw35dr.isDropped = 0:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:sw35dr.isDropped = 1:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:sw42dr.isDropped = 0:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:sw42dr.isDropped = 1:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:sw43dr.isDropped = 0:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:sw43dr.isDropped = 1:End Sub

'***Holes***

Sub sw27_Hit:  vpmTimer.PulseSw 27:T27P.RotX = 100:T27P2.RotX = 100:T27.TimerEnabled = 1: PlayNeutralZoneSubwaySnd = True : NeutralZoneSWTimer.Enabled = 1 : PlayTargetSound : End Sub ' SoundNeutralZone T27P2:End Sub
Dim PlayNeutralZoneSubwaySnd : PlayNeutralZoneSubwaySnd = False
Sub NeutrualZoneSubWayTrigger_Hit
  If PlayNeutralZoneSubwaySnd Then
    'Debug.Print "PlayNeutralZoneSubwaySnd_hit"
    PlaySoundAt "NeutralZoneUnderramp", NeutrualZoneSubWayTrigger
    PlayNeutralZoneSubwaySnd = False
    NeutralZoneSWTimer.Enabled = 0
  End If
End Sub
Sub NeutralZoneSWTimer_Timer
  'Debug.Print "NZTimer"
  PlayNeutralZoneSubwaySnd = False
  NeutralZoneSWTimer.Enabled = 0
End Sub
Sub sw45_Hit: SoundNeutralZone sw45:vpmTimer.PulseSw 45:End Sub
'Sub sw45t_Hit: Controller.Switch(45) = 1:End Sub
'Sub sw45t_UnHit: Controller.Switch(45) = 0:End Sub
'Sub sw46t_Hit: Controller.Switch(46) = 1:PlaySoundAt "CommandD2", sw46t:End Sub
Sub sw46t_Hit: Controller.Switch(46) = 1:End Sub
Sub sw46_Hit: PlaySound "CommandD2", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
Sub sw46t_UnHit: Controller.Switch(46) = 0:End Sub
Sub sw47t_Hit: Controller.Switch(47) = 1:sw47dr.isDropped = 0:End Sub
Sub sw47_Hit:PlaySoundAt "StartMission1", sw47:End Sub
Sub sw47t_UnHit: Controller.Switch(47) = 0:sw47dr.isDropped = 1:End Sub



'***Targets***
'T27 only pulse with sw27(Hole) to prevent double hits
Sub T26_Hit:vpmTimer.PulseSw 26:End Sub
'Sub T27_Hit:T27P.RotX = 100:T27P2.RotX = 100:PlaySound "target":T27.TimerEnabled = 1:End Sub
Sub T27_Timer:T27P.RotX = 90:T27P2.RotX = 90:End Sub
Sub T28_Hit:vpmTimer.PulseSw 28:End Sub

Sub T51_Hit:vpmTimer.PulseSw 51:End Sub
Sub T52_Hit:vpmTimer.PulseSw 52:End Sub
Sub T53_Hit:vpmTimer.PulseSw 53:End Sub

Sub T54_Hit:vpmTimer.PulseSw 54:End Sub
Sub T55_Hit:vpmTimer.PulseSw 55:End Sub
Sub T56_Hit:vpmTimer.PulseSw 56:End Sub

Sub T81_Hit:vpmTimer.PulseSw 81:End Sub
Sub T82_Hit:vpmTimer.PulseSw 82:End Sub

Sub T84_Hit:vpmTimer.PulseSw 84:End Sub

Sub T85_Hit:vpmTimer.PulseSw 85:End Sub
Sub T86_Hit:vpmTimer.PulseSw 86:End Sub


'KickerAnimation


'***************
' ToysAnimation
'***************


'***Cannons***

Dim CannonSpeed, CannonSteps, CannonAnimationSteps, CannonLBall, CannonRBall, RadiusL, RadiusR, CannonAngleL, CannonAngleR
Dim RadiusRL, RadiusLL, CML, CMR, CPL, CPR
Dim ForceL, ForceR, LaserLON, LaserRON

CannonSpeed = 600 'Timer with 1 = 1000/Second
CannonSteps = 82

Const Pi = 3.14159265358979
RadiusL = CannonBaseL.Y - (kicker1.Y-6)
RadiusR = CannonBaseR.Y - (kicker2.Y-6)
RadiusLL = CannonBaseL.Y - l52.Y
RadiusRL = CannonBaseR.Y - l82.Y


Sub Kicker1_Hit()
' me.destroyball
' Set CannonLBall = kicker1.createball
  Set CannonLBall = ActiveBall
  Controller.Switch(38) = 1
  me.Enabled = False
  If LaserMod = 1 then
    LaserL.Size_Z = 0.3
    LaserL1.Size_Z = 0.3
    LaserL.visible = 1
    LaserL1.visible = 1
    LaserZTimer.Enabled = 1
    LaserLON = 1
  End if
End Sub

Sub Kicker2_Hit()
' me.destroyball
' Set CannonRBall = kicker2.createball
  Set CannonRBall = ActiveBall
  Controller.Switch(34) = 1
  me.Enabled = False
  If LaserMod = 1 then
    LaserR.Size_Z = 0.5
    LaserR1.Size_Z = 0.5
    LaserR.visible = 1
    LaserR1.visible = 1
    LaserZTimer1.Enabled = 1
    LaserRON = 1
  End if
End Sub




Sub CannonLTimer_Timer()
    If CML = True and CannonBaseL.ObjRotZ < 64 Then
    CannonBaseL.ObjRotZ = CannonBaseL.ObjRotZ +(CannonSteps*2)/CannonSpeed
    End if
    If CML = False and CannonBaseL.ObjRotZ > -19 then
    CannonBaseL.ObjRotZ = CannonBaseL.ObjRotZ -(CannonSteps*2)/CannonSpeed
    End if
    If CannonBaseL.ObjRotZ >= 64 then CML = False
    If CannonBaseL.ObjRotZ >= -20 And CannonBaseL.ObjRotZ <= -17 then Controller.Switch(127) = 1: Else Controller.Switch(127) = 0
    If CannonBaseL.ObjRotZ >= -20 And CannonBaseL.ObjRotZ <= 9 then Controller.Switch(122) = 1:ForceL = 25: Else Controller.Switch(122) = 0: ForceL = 50

    If Not IsEmpty(CannonLBall) Then
    CannonLBall.X= CannonBaseL.X - RadiusL*sin(-CannonBaseL.ObjRotZ*(Pi/180))
    CannonLBall.Y= CannonBaseL.Y - RadiusL*cos(-CannonBaseL.ObjRotZ*(Pi/180))
    'CannonLBall.Z=105 + BallSize
    CannonLBall.Z=86 + BallSize/2
    CannonAngleL = CannonBaseL.ObjRotZ
    End If
    CannonL.ObjRotZ = CannonBaseL.ObjRotZ
    CannonDomeL.ObjRotZ = CannonBaseL.ObjRotZ
    CannonPinL.ObjRotZ = CannonBaseL.ObjRotZ
    l52.X = CannonBaseL.X - RadiusRL*sin(-CannonBaseL.ObjRotZ*(Pi/180))
    l52.Y = CannonBaseL.Y - RadiusRL*cos(-CannonBaseL.ObjRotZ*(Pi/180))
    LaserL.ObjRotZ = CannonBaseL.ObjRotZ
    LaserL1.ObjRotZ = CannonBaseL.ObjRotZ
End Sub



Sub CannonRTimer_Timer()
    If CMR = True and CannonBaseR.ObjRotZ < 64 Then
    CannonBaseR.ObjRotZ = CannonBaseR.ObjRotZ +(CannonSteps*2)/CannonSpeed
    End if
    If CMR = False and CannonBaseR.ObjRotZ > -19 then
    CannonBaseR.ObjRotZ = CannonBaseR.ObjRotZ -(CannonSteps*2)/CannonSpeed
    End if
    If CannonBaseR.ObjRotZ >= 64 then CMR = False
    If CannonBaseR.ObjRotZ >= -20 And CannonBaseR.ObjRotZ <= -17 then Controller.Switch(125) = 1: Else Controller.Switch(125) = 0
    If CannonBaseR.ObjRotZ >= -20 And CannonBaseR.ObjRotZ <= 9 then Controller.Switch(126) = 1:ForceR = 25: Else Controller.Switch(126) = 0: ForceR = 50

    If Not IsEmpty(CannonRBall) Then
    CannonRBall.X= CannonBaseR.X - RadiusR*sin(+CannonBaseR.ObjRotZ*(Pi/180))
    CannonRBall.Y= CannonBaseR.Y - RadiusR*cos(+CannonBaseR.ObjRotZ*(Pi/180))
    'CannonRBall.Z=105 + BallSize
    CannonRBall.Z=86 + BallSize/2
    CannonAngleR = CannonBaseR.ObjRotZ- (CannonBaseR.ObjRotZ*2)
    End If
    CannonR.ObjRotZ = CannonBaseR.ObjRotZ
    CannonDomeR.ObjRotZ = CannonBaseR.ObjRotZ
    CannonPinR.ObjRotZ = CannonBaseR.ObjRotZ
    l82.X = CannonBaseR.X - RadiusRL*sin(+CannonBaseR.ObjRotZ*(Pi/180))
    l82.Y = CannonBaseR.Y - RadiusRL*cos(+CannonBaseR.ObjRotZ*(Pi/180))
    LaserR.ObjRotZ = CannonBaseR.ObjRotZ
    LaserR1.ObjRotZ = CannonBaseR.ObjRotZ
End Sub

'***CannonShoot***

Sub LeftCannonKicker(Enabled)
  If Enabled then
    CPL = True
    If Not IsEmpty(CannonLBall) Then
      Kicker1.kick CannonAngleL, ForceL
      'PlaysoundAt SoundFX("LGunKicker", DOFContactors), CannonBaseL
      RandomSoundBallRelease CannonBaseL
      Controller.Switch(38) = 0
      CannonLBall = Empty
      Kicker1.Enabled = True
      LaserLON = 0
      If LaserType = 1 then
        LaserL.Visible = 0
        LaserL1.Visible = 0
        LaserZTimer.Enabled = 0
      End if
    End if
  End if
End Sub


Sub RightCannonKicker(Enabled)
  If Enabled then
    CPR = True
    If Not IsEmpty(CannonRBall) Then
      Kicker2.kick CannonAngleR, ForceR
      'PlaysoundAt SoundFX("RGunKicker", DOFContactors), CannonBaseR
      RandomSoundBallRelease CannonBaseR
      Controller.Switch(34) = 0
      CannonRBall = Empty
      Kicker2.Enabled = True
      LaserRON = 0
      If LaserType = 1 then
        LaserR.Visible = 0
        LaserR1.Visible = 0
        LaserZTimer1.Enabled = 0
      End if
    End if
  End if
End Sub


'***LaserZTimer***

Sub LaserZTimer_Timer()
  If CannonBaseL.ObjRotZ >= -20 And CannonBaseL.ObjRotZ <= -17 then LaserL.Size_Z = 0.3: LaserL1.Size_Z = 0.3
  If CannonBaseL.ObjRotZ >= -16.99 And CannonBaseL.ObjRotZ <= -8 then LaserL.Size_Z = 0.25: LaserL1.Size_Z = 0.25
  If CannonBaseL.ObjRotZ >= -7.99 And CannonBaseL.ObjRotZ <= -5 then LaserL.Size_Z = 0.3: LaserL1.Size_Z = 0.3
  If CannonBaseL.ObjRotZ >= -4.99 And CannonBaseL.ObjRotZ <= -1 then LaserL.Size_Z = 0.67: LaserL1.Size_Z = 0.67
  If CannonBaseL.ObjRotZ >= 0.99 And CannonBaseL.ObjRotZ <= 5.99 then LaserL.Size_Z = 1: LaserL1.Size_Z = 1

  If BorgMod = 0 then
    If CannonBaseL.ObjRotZ >= 6 And CannonBaseL.ObjRotZ <= 21.99 then LaserL.Size_Z = 0.75: LaserL1.Size_Z = 0.75
  End if
  If BorgMod = 1 then
    If CannonBaseL.ObjRotZ >= 6 And CannonBaseL.ObjRotZ <= 21.99 then LaserL.Size_Z = 1: LaserL1.Size_Z = 1
  End if

  If CannonBaseL.ObjRotZ >= 22 And CannonBaseL.ObjRotZ <= 56.99 then LaserL.Size_Z = 1: LaserL1.Size_Z = 1
  If CannonBaseL.ObjRotZ >= 57 And CannonBaseL.ObjRotZ <= 62.99 then LaserL.Size_Z = 0.48: LaserL1.Size_Z = 0.48
  If CannonBaseL.ObjRotZ >= 63 then LaserL.Size_Z = 1:LaserL1.Size_Z = 1

  If CannonBaseL.ObjRotZ <= -19 And LaserLON = 0 Then
    LaserL.Visible = 0
    LaserL1.Visible = 0
    LaserZTimer.Enabled = 0
  End if
End Sub

Sub LaserZTimer1_Timer()
  If CannonBaseR.ObjRotZ >= -20 And CannonBaseR.ObjRotZ <= -18 then LaserR.Size_Z = 0.5: LaserR1.Size_Z = 0.5
  If CannonBaseR.ObjRotZ >= -17.99 And CannonBaseR.ObjRotZ <= -8 then LaserR.Size_Z = 0.235: LaserR1.Size_Z = 0.235
  If CannonBaseR.ObjRotZ >= -7.99 And CannonBaseR.ObjRotZ <= 8.99 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1

  If BorgMod = 0 then
    If CannonBaseR.ObjRotZ >= 9 And CannonBaseR.ObjRotZ <= 22.99 then LaserR.Size_Z = 0.76: LaserR1.Size_Z = 0.76
  End if
  If BorgMod = 1 then
    If CannonBaseR.ObjRotZ >= 9 And CannonBaseR.ObjRotZ <= 22.99 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1
  End if

  If CannonBaseR.ObjRotZ >= 23 And CannonBaseR.ObjRotZ <= 28.99 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1
  If CannonBaseR.ObjRotZ >= 29 And CannonBaseR.ObjRotZ <= 32.99 then LaserR.Size_Z = 0.785: LaserR1.Size_Z = 0.785
  If CannonBaseR.ObjRotZ >= 33 And CannonBaseR.ObjRotZ <= 50.99 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1
  If CannonBaseR.ObjRotZ >= 51 And CannonBaseR.ObjRotZ <= 57.99 then LaserR.Size_Z = 0.51: LaserR1.Size_Z = 0.51
  If CannonBaseR.ObjRotZ >= 58 And CannonBaseR.ObjRotZ <= 62.99 then LaserR.Size_Z = 0.49: LaserR1.Size_Z = 0.49
  If CannonBaseR.ObjRotZ >= 63 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1

  If CannonBaseR.ObjRotZ <= -18 And LaserRON = 0 Then
    LaserR.Visible = 0
    LaserR1.Visible = 0
    LaserZTimer1.Enabled = 0
  End if
End Sub

If Lasermod = 1 Then
  LaserBarrier1.visible = 1
  LaserBarrier2.visible = 1
Else
  LaserBarrier1.visible = 0
  LaserBarrier2.visible = 0
End If


'***CannonPin Animation***

Sub CannonPinLTimer
  If CPL = True and CannonPinL.TransZ < 40 then CannonPinL.TransZ = CannonPinL.TransZ + 5
  If CPL = False and CannonPinL.TransZ > 0 then CannonPinL.TransZ = CannonPinL.TransZ - 5
  If CannonPinL.TransZ >= 35 then CPL = False
End Sub


Sub CannonPinRTimer
  If CPR = True and CannonPinR.TransZ < 40 then CannonPinR.TransZ = CannonPinR.TransZ + 5
  If CPR = False and CannonPinR.TransZ > 0 then CannonPinR.TransZ = CannonPinR.TransZ - 5
  If CannonPinR.TransZ >= 35 then CPR = False
End Sub

'***Kickback Animation***


Sub KickBackTimer()
  if KB = True and KickBackP.TransZ < 60 then KickBackP.TransZ = KickBackP.TransZ +5
  if KB = False and KickBackP.TransZ > 0 then KickBackP.TransZ = KickBackP.TransZ -5
  if KickBackP.TransZ >= 60 then KB = False
End Sub

'***Catapult Animation***


Sub CatapultTimer()
  if CP = True and Catapult.RotX < 110 then Catapult.RotX = Catapult.RotX +10
  if CP = False and Catapult.RotX > 90 then Catapult.RotX = Catapult.RotX -10
' if Catapult.RotX >= 107 then CP = False
End Sub

'************
' SlingShots
'************


Dim RStep, Lstep

Sub SlingShotRight_Slingshot
  RS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotRight sling1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    SlingShotRight.TimerEnabled = 1
    vpmTimer.PulseSw 74
End Sub

Sub SlingShotRight_Timer
    Select Case RStep
        Case 2:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 3:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:SlingShotRight.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub SlingShotLeft_Slingshot
  LS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    SlingShotLeft.TimerEnabled = 1
    vpmTimer.PulseSw 75
End Sub

Sub SlingShotLeft_Timer
    Select Case LStep
        Case 2:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 3:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:SlingShotLeft.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = MechanicalTilt Then
        vpmTimer.PulseSw vpmNudge.TiltSwitch
        Exit Sub
    End if

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

  If keycode = LockBarKey Then Controller.Switch(12) = 1

    If keycode = PlungerKey Then
    Controller.Switch(12) = 1
    If VRRoom > 0 and VRRoom < 3 then STNNGPlungerTrigger.rotX = 40
  End if
  If keycode = keyFront Then Controller.Switch(11) = 1
  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    If VRRoom > 0 and VRRoom < 3 then
      FlipperButtonLeft.x = FlipperButtonLeft.x + 5
    End If
  End If
  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress

    'FlipperActivate FlipperR1, RF1Press
    If VRRoom > 0 and VRRoom < 3 then
      FlipperButtonRight.x = FlipperButtonRight.x - 5
    End If
  End If


  If keycode = LeftMagnasave and VRRoom = 1 Then
    HolodeckMoveLeftTimer.enabled = true
  End If
  If keycode = RightMagnasave and VRRoom = 1 Then
    HolodeckMoveRightTimer.enabled = true
  End If


  if keycode=StartGameKey then soundStartButton() : BallSearch ' check trough switches (catch hard pinmame reset)

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey or keycode = LockBarKey Then
    Controller.Switch(12) = 0
    If VRRoom > 0 and VRRoom < 3 then STNNGPlungerTrigger.rotX = 6
  End if
  If keycode = keyFront Then Controller.Switch(11) = 0

  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    If VRRoom > 0 and VRRoom < 3 then
      FlipperButtonLeft.x = FlipperButtonLeft.x - 5
    End If
  End If
  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    'FlipperDeActivate FlipperR1, RF1Press
    If VRRoom > 0 and VRRoom < 3 then
      FlipperButtonRight.x = FlipperButtonRight.x + 5
    end If
  End If


  If keycode = LeftMagnasave and VRRoom = 1 Then
    HolodeckMoveLeftTimer.enabled = false
  End If
  If keycode = RightMagnasave and VRRoom = 1 Then
    HolodeckMoveRightTimer.enabled = false
  End If

  If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub table1_Paused:Controller.Pause = True:End Sub
Sub table1_unPaused:Controller.Pause = False:End Sub
Sub table1_exit():Controller.Pause = False:Controller.Stop:End Sub



'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"

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
    'RF1.Fire
    If StagedFlipperMod <> 1 Then FlipperR1.RotateToEnd

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    If StagedFlipperMod <> 1 Then FlipperR1.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolURFlipper(Enabled)
  If Enabled Then
    FlipperR1.RotateToEnd
    If FlipperR1.currentangle > FlipperR1.endangle - ReflipAngle Then
      RandomSoundReflipUpRight FlipperR1
    Else
      SoundFlipperUpAttackRight FlipperR1
      RandomSoundFlipperUpRight FlipperR1
    End If
  Else
    FlipperR1.RotateToStart
    If FlipperR1.currentangle > FlipperR1.startAngle + 5 Then
      RandomSoundFlipperDownRight FlipperR1
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  'Debug.Print "LFCollide"
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  'Debug.Print "RFCollide"
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub FlipperR1_Collide(parm)
  'CheckLiveCatch ActiveBall, Flipperr1, RF1Count, parm
  RightFlipperCollide parm
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
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
  NFadeLm 11,  l11
  NFadeLm 11,  l11b
  Flash 11, f11

  NFadeLm 12,  l12
  NFadeLm 12,  l12b
  Flash 12, f12

  NFadeLm 13,  l13
  NFadeL 13,  l13b
  NFadeLm 14,  l14
  NFadeL 14,  l14b

  NFadeLm 15,  l15
  NFadeLm 15,  l15b
  Flash 15, f15

  NFadeLm 16,  l16
  NFadeL 16,  l16b
  NFadeLm 17,  l17
  NFadeL 17,  l17b
  NFadeLm 18,  l18
  NFadeL 18,  l18b

  NFadeLm 21,  l21
  NFadeL 21,  l21b
  NFadeLm 22,  l22
  NFadeL 22,  l22b
  NFadeLm 23,  l23
  NFadeL 23,  l23b

  NFadeLm 24,  l24
  NFadeLm 24,  l24b
  Flash 24, f24

  NFadeLm 25,  l25
  NFadeLm 25,  l25b
  Flash 25, f25

  Flashm 26, f26s
  Flashm 26, f26s1
  NFadeObjm 26, l26blue, "bulbcover_blueOn", "bulbcover_blue"
  Flash 26, f26

  NFadeLm 26,  l26
  NFadeL 26,  l26b
  NFadeLm 27,  l27
  NFadeL 27,  l27b
  NFadeLm 28,  l28
  NFadeL 28,  l28b
  Flash 28, f28

  NFadeLm 31,  l31
  NFadeL 31,  l31b
  NFadeLm 32,  l32
  NFadeL 32,  l32b
  NFadeLm 33,  l33
  NFadeL 33,  l33b
  NFadeLm 34,  l34
  NFadeLm 34,  l34b
  Flash 34, f34

  NFadeLm 35,  l35
  NFadeL 35,  l35b
  NFadeLm 36,  l36
  NFadeL 36,  l36b
  NFadeLm 37,  l37
  NFadeL 37,  l37b

  NFadeLm 38,  l38
  NFadeLm 38,  l38b
  Flash 38, f38

  NFadeLm 41,  l41
  NFadeLm 41,  l41b
  Flash 41, f41

  NFadeLm 42,  l42
  NFadeL 42,  l42b
  NFadeLm 43,  l43
  NFadeL 43,  l43b
  NFadeLm 44,  l44
  NFadeL 44,  l44b
  NFadeLm 45,  l45
  NFadeLm 45,  l45b
  Flash 45, f45

  NFadeLm 46,  l46
  NFadeL 46,  l46b
  NFadeLm 47,  l47
  NFadeL 47,  l47b
  NFadeLm 48,  l48
  NFadeL 48,  l48b
  Flash 48, f48

  NFadeLm 51,  l51
  NFadeLm 51,  l51b
  Flashm 52, f52s
  Flashm 52, f52s1
  NFadeLm 52,  l52
  Flash 52, f52

  Flashm 53, f53s
  NFadeObjm 53, l53yellow, "bulbcover_yellowOn", "bulbcover_yellow"
  Flash 53, f53

  NFadeLm 54,  l54
  NFadeL 54,  l54b
  NFadeLm 55,  l55
  NFadeL 55,  l55b
  NFadeLm 56,  l56
  NFadeL 56,  l56b
  NFadeLm 57,  l57
  NFadeL 57,  l57b
  NFadeLm 58,  l58
  NFadeL 58,  l58b

  NFadeLm 61,  l61
  NFadeL 61,  l61b
  NFadeLm 62,  l62
  NFadeL 62,  l62b
  NFadeLm 63,  l63
  NFadeL 63,  l63b
  NFadeLm 64,  l64
  NFadeL 64,  l64b
  NFadeLm 65,  l65
  NFadeL 65,  l65b
  NFadeLm 66,  l66
  NFadeL 66,  l66b

  NFadeLm 67,  l67
  NFadeLm 67,  l67b
  Flash 67, f67

  NFadeLm 68,  l68
  NFadeLm 68,  l68b
  Flash 68, f68

  NFadeLm 71,  l71
  NFadeL 71,  l71b
  NFadeLm 72,  l72
  NFadeL 72,  l72b
  NFadeLm 73,  l73
  NFadeL 73,  l73b
  NFadeLm 74,  l74
  NFadeL 74,  l74b
  NFadeLm 75,  l75
  NFadeL 75,  l75b
  NFadeLm 76,  l76
  NFadeL 76,  l76b
  NFadeLm 77,  l77
  NFadeL 77,  l77b

  NFadeLm 78,  l78a
  NFadeLm 78,  l78b
  NFadeLm 78,  l78c
  NFadeLm 78,  l78d
  NFadeLm 78,  l78e

  NFadeLm 78,  l78borga
  NFadeLm 78,  l78borgb
  NFadeLm 78,  l78borgc
  NFadeLm 78,  l78borgd
  NFadeL 78,  l78borge

  NFadeLm 81,  l81
  NFadeL 81,  l81b

  Flashm 82,  f82s
  Flashm 82,  f82s1
  NFadeLm 82,  l82
  Flash 82,  f82

  NFadeLm 84,  l84
  NFadeLm 84,  l84b
  Flash 84, f84

  Flashm 85, f85s
  NFadeObjm 85, l85green, "bulbcover_greenOn", "bulbcover_green"
  Flash 85, f85

  Flashm 86, f86s
  Flashm 86, f86s1
  NFadeObjm 86, l86red, "bulbcover_redOn", "bulbcover_red"
  Flash 86, f86

  NFadeLm 124, l124
  NFadeL 124, l124b
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


'***SolModFlasher***


dim FlashTargetLevel20, FlashLevel20, FlashTargetLevel21, FlashLevel21, FlashTargetLevel22, FlashLevel22, FlashTargetLevel23, FlashLevel23,  FlashTargetLevel25, FlashLevel25
dim FlashTargetLevel26, FlashLevel26, FlashTargetLevel27, FlashLevel27, FlashTargetLevel28, FlashLevel28,  FlashTargetLevel41, FlashLevel41, FlashTargetLevel42, FlashLevel42

Sub Flash120(aLvl) 'RightBorgFlasher
  FlashTargetLevel20 = aLvl/255
  l120_Timer
End Sub

l120.state=1
l83.state=1
l120.IntensityScale = 0
l83.IntensityScale = 0

sub l120_Timer()
  If not l120.TimerEnabled Then
    l120.TimerEnabled = True
  End If
    l120.IntensityScale = 6 * FlashLevel20^3
    l83.IntensityScale = 5 * FlashLevel20^2
  if round(FlashTargetLevel20,1) > round(FlashLevel20,1) Then
    FlashLevel20 = FlashLevel20 + 0.3
    if FlashLevel20 > 1 then FlashLevel20 = 1
  Elseif round(FlashTargetLevel20,1) < round(FlashLevel20,1) Then
    FlashLevel20= FlashLevel20 * 0.8 - 0.01
    if FlashLevel20 < 0 then FlashLevel20 = 0
  Else
    FlashLevel20 = round(FlashTargetLevel20,1)
    l120.TimerEnabled = False
  end if

  If FlashLevel20 <= 0 Then
    l120.TimerEnabled = False
    l120.IntensityScale = 0
    l83.IntensityScale = 0
  End If
end sub


Sub Flash121(aLvl) 'Right Popper
  FlashTargetLevel21 = aLvl/255
  f121_Timer
End Sub

f121.visible=0
l121.IntensityScale = 0
l121.state=1

sub F121_Timer()
  If not F121.TimerEnabled Then
    F121.TimerEnabled = True
    F121.visible=1
    If VRRoom > 0 Then
      VRBGFL21_1.visible = true
      VRBGFL21_2.visible = true
      VRBGFL21_3.visible = true
      VRBGFL21_4.visible = true
      VRBGFL21_5.visible = true
    End If
  End If
    f121.opacity = 1000 * FlashLevel21^2
    l121.IntensityScale = 10 * FlashLevel21
  If VRRoom > 0 Then
    VRBGFL21_1.opacity = 100 * FlashLevel21^1.5
    VRBGFL21_2.opacity = 100 * FlashLevel21^1.5
    VRBGFL21_3.opacity = 100 * FlashLevel21^1.5
    VRBGFL21_4.opacity = 100 * FlashLevel21^2
    VRBGFL21_5.opacity = 100 * FlashLevel21^3
  End If

  if round(FlashTargetLevel21,1) > round(FlashLevel21,1) Then
    FlashLevel21 = FlashLevel21 + 0.3
    if FlashLevel21 > 1 then FlashLevel21 = 1
  Elseif round(FlashTargetLevel21,1) < round(FlashLevel21,1) Then
    FlashLevel21 = FlashLevel21 * 0.8 - 0.01
    if FlashLevel21 < 0 then FlashLevel21 = 0
  Else
    FlashLevel21 = round(FlashTargetLevel21,1)
'   debug.print "stop timer"
    F121.TimerEnabled = False
  end if

  If FlashLevel21 <= 0 Then
    F121.TimerEnabled = False
    F121.visible=0
    l121.IntensityScale = 0
    If VRRoom > 0 Then
      VRBGFL21_1.visible = False
      VRBGFL21_2.visible = False
      VRBGFL21_3.visible = False
      VRBGFL21_4.visible = False
      VRBGFL21_5.visible = False
    End If
  End If
end sub

Sub Flash122(aLvl) 'Middle Ramp
  FlashTargetLevel22 = aLvl/255
  f122_Timer
End Sub

f122.visible=0
f122b.visible=0
f122s.visible=0

sub F122_Timer()
  If not f122.TimerEnabled Then
    f122.TimerEnabled = True
    f122.visible=1
    f122b.visible=1
    f122s.visible=1
    If VRRoom > 0 Then
      VRBGFL22_1.visible = true
      VRBGFL22_2.visible = true
      VRBGFL22_3.visible = true
      VRBGFL22_4.visible = true
      VRBGFL22_5.visible = true
    End If
  End If
    f122.opacity = 1000 * FlashLevel22^2
    f122b.opacity = 500 * FlashLevel22^2
    f122s.opacity = 100 * FlashLevel22^2

  If VRRoom > 0 Then
    VRBGFL22_1.opacity = 100 * FlashLevel22^1.5
    VRBGFL22_2.opacity = 100 * FlashLevel22^1.5
    VRBGFL22_3.opacity = 100 * FlashLevel22^1.5
    VRBGFL22_4.opacity = 100 * FlashLevel22^2
    VRBGFL22_5.opacity = 100 * FlashLevel22^3
  End If

  if round(FlashTargetLevel22,1) > round(FlashLevel22,1) Then
    FlashLevel22 = FlashLevel22 + 0.3
    if FlashLevel22 > 1 then FlashLevel22 = 1
  Elseif round(FlashTargetLevel22,1) < round(FlashLevel22,1) Then
    FlashLevel22 = FlashLevel22 * 0.8 - 0.01
    if FlashLevel22 < 0 then FlashLevel22 = 0
  Else
    FlashLevel22 = round(FlashTargetLevel22,1)
'   debug.print "stop timer"
    F122.TimerEnabled = False
  end if

  If FlashLevel22 <= 0 Then
    f122.TimerEnabled = False
    f122.visible=0
    f122b.visible=0
    f122s.visible=0
    If VRRoom > 0 Then
      VRBGFL22_1.visible = False
      VRBGFL22_2.visible = False
      VRBGFL22_3.visible = False
      VRBGFL22_4.visible = False
      VRBGFL22_5.visible = False
    End If
  End If
end sub


Sub Flash123(aLvl) 'Shield Flasher
  FlashTargetLevel23 = aLvl/255
  f123_Timer
End Sub

f123.visible=0
f123a.visible=0
ShieldGiBig7.IntensityScale = 0
ShieldGiBig8.IntensityScale = 0
ShieldGiBig10.IntensityScale = 0


sub F123_Timer()
  If not f123.TimerEnabled Then
    f123.TimerEnabled = True
    f123.visible=1
    f123a.visible=1
    If VRRoom > 0 Then
      VRBGFL23_1.visible = true
      VRBGFL23_2.visible = true
      VRBGFL23_3.visible = true
      VRBGFL23_4.visible = true
      VRBGFL23_5.visible = true
    End If
  End If
    f123.opacity = 20 * FlashLevel23^2
    f123a.opacity = 20 * FlashLevel23^2
    ShieldGiBig7.IntensityScale = 30 * FlashLevel23
    ShieldGiBig8.IntensityScale = 30 * FlashLevel23
    ShieldGiBig10.IntensityScale = 30 * FlashLevel23

  If VRRoom > 0 Then
    VRBGFL23_1.opacity = 100 * FlashLevel23^1.5
    VRBGFL23_2.opacity = 100 * FlashLevel23^1.5
    VRBGFL23_3.opacity = 100 * FlashLevel23^1.5
    VRBGFL23_4.opacity = 100 * FlashLevel23^2
    VRBGFL23_5.opacity = 100 * FlashLevel23^3
  End If

  if round(FlashTargetLevel23,1) > round(FlashLevel23,1) Then
    FlashLevel23 = FlashLevel23 + 0.3
    if FlashLevel23 > 1 then FlashLevel23 = 1
  Elseif round(FlashTargetLevel23,1) < round(FlashLevel23,1) Then
    FlashLevel23 = FlashLevel23 * 0.8 - 0.01
    if FlashLevel23 < 0 then FlashLevel23 = 0
  Else
    FlashLevel23 = round(FlashTargetLevel23,1)
'   debug.print "stop timer"
    F123.TimerEnabled = False
  end if

  If FlashLevel23 <= 0 Then
    f123.TimerEnabled = False
    f123.visible=0
    f123a.visible=0
    ShieldGiBig7.IntensityScale = 0
    ShieldGiBig8.IntensityScale = 0
    ShieldGiBig10.IntensityScale = 0
    If VRRoom > 0 Then
      VRBGFL23_1.visible = False
      VRBGFL23_2.visible = False
      VRBGFL23_3.visible = False
      VRBGFL23_4.visible = False
      VRBGFL23_5.visible = False
    End If
  End If
end sub

Sub Flash125(aLvl) 'LeftPopper
  FlashTargetLevel25 = aLvl/255
  f125_Timer
End Sub

f125.visible=0
l125.IntensityScale = 0
l125.state=1

sub F125_Timer()
  If not f125.TimerEnabled Then
    f125.TimerEnabled = True
    f125.visible=1
    If VRRoom > 0 Then
      VRBGFL25_1.visible = true
      VRBGFL25_2.visible = true
      VRBGFL25_3.visible = true
      VRBGFL25_4.visible = true
      VRBGFL25_5.visible = true
    End If
  End If
    f125.opacity = 1000 * FlashLevel25^2
    l125.IntensityScale = 10 * FlashLevel25
  If VRRoom > 0 Then
    VRBGFL25_1.opacity = 100 * FlashLevel25^1.5
    VRBGFL25_2.opacity = 100 * FlashLevel25^1.5
    VRBGFL25_3.opacity = 100 * FlashLevel25^1.5
    VRBGFL25_4.opacity = 100 * FlashLevel25^2
    VRBGFL25_5.opacity = 100 * FlashLevel25^3
  End If

  if round(FlashTargetLevel25,1) > round(FlashLevel25,1) Then
    FlashLevel25 = FlashLevel25 + 0.3
    if FlashLevel25 > 1 then FlashLevel25 = 1
  Elseif round(FlashTargetLevel25,1) < round(FlashLevel25,1) Then
    FlashLevel25 = FlashLevel25 * 0.8 - 0.01
    if FlashLevel25 < 0 then FlashLevel25 = 0
  Else
    FlashLevel25 = round(FlashTargetLevel25,1)
'   debug.print "stop timer"
    f125.TimerEnabled = False
  end if

  If FlashLevel25 <= 0 Then
    f125.TimerEnabled = False
    f125.visible=0
    l125.IntensityScale = 0
    If VRRoom > 0 Then
      VRBGFL25_1.visible = False
      VRBGFL25_2.visible = False
      VRBGFL25_3.visible = False
      VRBGFL25_4.visible = False
      VRBGFL25_5.visible = False
    End If
  End If
end sub

Sub Flash126(aLvl) 'RightBorgFlasher
  FlashTargetLevel26 = aLvl/255
  f126_Timer
End Sub

sub F126_Timer()
  If not f126.TimerEnabled Then
    f126.TimerEnabled = True
    f126.visible=1
    f126s.visible=1
  End If
    f126.opacity = 500 * FlashLevel26^2.5
    f126s.opacity = 275 * FlashLevel26^2.5
'   l78d1.IntensityScale = 100 * FlashLevel26^1.5
'   l78e1.IntensityScale = 100 * FlashLevel26^1.5
    l78borgb1.IntensityScale = 100 * FlashLevel26^2
    l78borga1.IntensityScale = 100 * FlashLevel26^2
  if round(FlashTargetLevel26,1) > round(FlashLevel26,1) Then
    FlashLevel26 = FlashLevel26 + 0.3
    if FlashLevel26 > 1 then FlashLevel26 = 1
  Elseif round(FlashTargetLevel26,1) < round(FlashLevel26,1) Then
    FlashLevel26 = FlashLevel26 * 0.8 - 0.01
    if FlashLevel26 < 0 then FlashLevel26 = 0
  Else
    FlashLevel26 = round(FlashTargetLevel26,1)
'   debug.print "stop timer"
    f126.TimerEnabled = False
  end if

  If FlashLevel26 <= 0 Then
    f126.TimerEnabled = False
    f126s.visible = 0
    f126.visible = 0
'   l78d1.IntensityScale = 0
'   l78e1.IntensityScale = 0
    l78borgb1.IntensityScale = 0
    l78borga1.IntensityScale = 0
  End If
end sub

const SolDrawSpeed = 1.3
sub Flash127(lvl) 'left borg flasher
  dim alvl : aLvl = lvl / 255

  if aLvl < FlashLevel27 Then
    FlashLevel27 = Flashlevel27/SolDrawSpeed
  Else
    Flashlevel27 = aLvl
  end If
  f127_timer
End Sub


sub f127_timer()
  If not f127.TimerEnabled Then
    f127.TimerEnabled = True
    f127.visible=1
    f127s.visible=1
    If VRRoom > 0 Then
      VRBGFL27_1.visible = true
      VRBGFL27_2.visible = true
      VRBGFL27_3.visible = true
      VRBGFL27_4.visible = true
      VRBGFL27_5.visible = true
    End If
  End If

  f127.opacity = 450 * FlashLevel27^1.5
  f127s.opacity = 400 * FlashLevel27^1.5
  l78b1.IntensityScale = 50 * FlashLevel27^1
  l78c1.IntensityScale = 50 * FlashLevel27^1
  l78borgd1.IntensityScale = 50 * FlashLevel27^1
  l78borge1.IntensityScale = 50 * FlashLevel27^1
  If VRRoom > 0 Then
    VRBGFL27_1.opacity = 100 * FlashLevel27^1.5
    VRBGFL27_2.opacity = 100 * FlashLevel27^1.5
    VRBGFL27_3.opacity = 100 * FlashLevel27^1.5
    VRBGFL27_4.opacity = 100 * FlashLevel27^2
    VRBGFL27_5.opacity = 100 * FlashLevel27^3
  End If

  FlashLevel27 = FlashLevel27 * 0.95 - 0.01

  If FlashLevel27 <= 0 Then
    f127.TimerEnabled = False
    f127s.visible = 0
    f127.visible = 0
    l78b1.IntensityScale = 0
    l78c1.IntensityScale = 0
    l78borgd1.IntensityScale = 0
    l78borge1.IntensityScale = 0
    If VRRoom > 0 Then
      VRBGFL27_1.visible = False
      VRBGFL27_2.visible = False
      VRBGFL27_3.visible = False
      VRBGFL27_4.visible = False
      VRBGFL27_5.visible = False
    End If
  End If

end sub

Sub Flash128(aLvl) 'CenterBorgFlasher
  FlashTargetLevel28 = aLvl/255
  f128_Timer
End Sub

sub F128_Timer()
  If not f128.TimerEnabled Then
    f128.TimerEnabled = True
    f128.visible=1
    f128s.visible=1
    GiBigGreen.visible=1
    If VRRoom > 0 Then
      VRBGFL28_1.visible = true
      VRBGFL28_2.visible = true
      VRBGFL28_3.visible = true
      VRBGFL28_4.visible = true
      VRBGFL28_5.visible = true
    End If
  End If
  f128.opacity = 400 * FlashLevel28^2.5
  f128s.opacity = 250 * FlashLevel28^2.5
  GiBigGreen.opacity = 150 * FlashLevel28^2.5
  l78a1.IntensityScale = 100 * FlashLevel28^2
  l78borgc1.IntensityScale = 100 * FlashLevel28^2
  If VRRoom > 0 Then
    VRBGFL28_1.opacity = 100 * FlashLevel28^1.5
    VRBGFL28_2.opacity = 100 * FlashLevel28^1.5
    VRBGFL28_3.opacity = 100 * FlashLevel28^1.5
    VRBGFL28_4.opacity = 100 * FlashLevel28^2
    VRBGFL28_5.opacity = 100 * FlashLevel28^3
  End If
  if round(FlashTargetLevel28,1) > round(FlashLevel28,1) Then
    FlashLevel28 = FlashLevel28 + 0.3
    if FlashLevel28 > 1 then FlashLevel28 = 1
  Elseif round(FlashTargetLevel28,1) < round(FlashLevel28,1) Then
    FlashLevel28 = FlashLevel28 * 0.8 - 0.01
    if FlashLevel28 < 0 then FlashLevel28 = 0
  Else
    FlashLevel28 = round(FlashTargetLevel28,1)
'   debug.print "stop timer"
    f128.TimerEnabled = False
  end if

  If FlashLevel28 <= 0 Then
    f128.TimerEnabled = False
    f128s.visible = 0
    f128.visible = 0
'   GiBigGreen.visible=0
    l78a1.IntensityScale = 0
    l78borgc1.IntensityScale = 0
    If VRRoom > 0 Then
      VRBGFL28_1.visible = False
      VRBGFL28_2.visible = False
      VRBGFL28_3.visible = False
      VRBGFL28_4.visible = False
      VRBGFL28_5.visible = False
    End If
  End If
end sub

Sub Flash141(aLvl) 'RomulanFlasher
  FlashTargetLevel41 = aLvl/255
  f141_Timer
End Sub

f141.visible=0
f141s.visible=0
f141s1.visible=0
GiBigGreen.visible=0
l141.state=1
l141a.state=1
l141.IntensityScale = 0
l141a.IntensityScale = 0
FlasherCapGreen.blenddisablelighting = 0.1

sub F141_Timer()
  If not f141.TimerEnabled Then
    f141.TimerEnabled = True
    f141.visible=1
    f141s.visible=1
    f141s1.visible=1
    GiBigGreen.visible=1
    If VRRoom > 0 Then
      VRBGFL41_1.visible = true
      VRBGFL41_2.visible = true
      VRBGFL41_3.visible = true
      VRBGFL41_4.visible = true
      VRBGFL41_5.visible = true
    End If
  End If
    f141.opacity = 1000 * FlashLevel41^2.5
    f141s.opacity = 250 * FlashLevel41^2.5
    f141s1.opacity = 200 * FlashLevel41^2.5
    GiBigGreen.opacity = 250 * FlashLevel41^2.5
    l141.IntensityScale = 0.5 * FlashLevel41^2
    l141a.IntensityScale = 10 * FlashLevel41^2
    FlasherCapGreen.blenddisablelighting = 2 * FlashLevel41^1.5 + 0.1
  If VRRoom > 0 Then
    VRBGFL41_1.opacity = 100 * FlashLevel41^1.5
    VRBGFL41_2.opacity = 100 * FlashLevel41^1.5
    VRBGFL41_3.opacity = 100 * FlashLevel41^1.5
    VRBGFL41_4.opacity = 100 * FlashLevel41^2
    VRBGFL41_5.opacity = 100 * FlashLevel41^3
  End If
  if round(FlashTargetLevel41,1) > round(FlashLevel41,1) Then
    FlashLevel41 = FlashLevel41 + 0.3
    if FlashLevel41 > 1 then FlashLevel41 = 1
  Elseif round(FlashTargetLevel41,1) < round(FlashLevel41,1) Then
    FlashLevel41 = FlashLevel41 * 0.8 - 0.01
    if FlashLevel41 < 0 then FlashLevel41 = 0
  Else
    FlashLevel41 = round(FlashTargetLevel41,1)
'   debug.print "stop timer"
    f141.TimerEnabled = False
  end if

  If FlashLevel41 <= 0 Then
    f141.TimerEnabled = False
    f141.visible = 0
    f141s.visible = 0
    f141s1.visible=0
'   GiBigGreen.visible=0
    l141.IntensityScale = 0
    l141a.IntensityScale = 0
    FlasherCapGreen.blenddisablelighting = 0.1
    If VRRoom > 0 Then
      VRBGFL41_1.visible = False
      VRBGFL41_2.visible = False
      VRBGFL41_3.visible = False
      VRBGFL41_4.visible = False
      VRBGFL41_5.visible = False
    End If
  End If
end sub

Sub Flash142(aLvl) 'RightRampFlasher
  FlashTargetLevel42 = aLvl/255
  f142_Timer
End Sub

f142.visible=0
f142s.visible=0
f142s1.visible=0
GiBigRed.visible=0
GiBigRed1.visible=0
l142.state=1
l142.IntensityScale = 0
FlasherCapRed.blenddisablelighting = 0.1

sub F142_Timer()
  If not f142.TimerEnabled Then
    f142.TimerEnabled = True
    f142.visible=1
    f142s.visible=1
    f142s1.visible=1
    GiBigRed.visible=1
    GiBigRed1.visible=1
    If VRRoom > 0 Then
      VRBGFL42_1.visible = true
      VRBGFL42_2.visible = true
      VRBGFL42_3.visible = true
      VRBGFL42_4.visible = true
      VRBGFL42_5.visible = true
    End If
  End If
  f142.opacity = 1000 * FlashLevel42^2.5
  f142s.opacity = 250 * FlashLevel42^2.5
  f142s1.opacity = 200 * FlashLevel42^2.5
  GiBigRed.opacity = 100 * FlashLevel42^3.5
  GiBigRed1.opacity = 150 * FlashLevel42^3
  l142.IntensityScale = 10 * FlashLevel42^2
  FlasherCapRed.blenddisablelighting = 3 * FlashLevel42^2 + 0.1
  If VRRoom > 0 Then
    VRBGFL42_1.opacity = 100 * FlashLevel42^1.5
    VRBGFL42_2.opacity = 100 * FlashLevel42^1.5
    VRBGFL42_3.opacity = 100 * FlashLevel42^1.5
    VRBGFL42_4.opacity = 100 * FlashLevel42^2
    VRBGFL42_5.opacity = 100 * FlashLevel42^3
  End If
  if round(FlashTargetLevel42,1) > round(FlashLevel42,1) Then
    FlashLevel42 = FlashLevel42 + 0.3
    if FlashLevel42 > 1 then FlashLevel42 = 1
  Elseif round(FlashTargetLevel42,1) < round(FlashLevel42,1) Then
    FlashLevel42 = FlashLevel42 * 0.8 - 0.01
    if FlashLevel42 < 0 then FlashLevel42 = 0
  Else
    FlashLevel42 = round(FlashTargetLevel42,1)
'   debug.print "stop timer"
    f142.TimerEnabled = False
  end if

  If FlashLevel42 <= 0 Then
    f142.TimerEnabled = False
    f142.visible = 0
    f142s.visible = 0
    f142s1.visible=0
    GiBigRed.visible=0
    GiBigRed1.visible=0
    l142.IntensityScale = 0
    FlasherCapRed.blenddisablelighting = 0.1
    If VRRoom > 0 Then
      VRBGFL42_1.visible = false
      VRBGFL42_2.visible = false
      VRBGFL42_3.visible = false
      VRBGFL42_4.visible = false
      VRBGFL42_5.visible = false
    End If
  End If
end sub


'BorgLampTimer - turns OFF the blue lamps while the green flasher are ON

'Sub BorgLampTimer_Timer()
Sub BorgLampTimer
  If BorgMod = 1 then
    If l78a1.Intensity >= 10 Then l78a.Intensity = 0:Else:l78a.Intensity = 30:End if
    If l78b1.Intensity >= 10 Then l78b.Intensity = 0:Else:l78b.Intensity = 30:End if
    If l78c1.Intensity >= 10 Then l78c.Intensity = 0:Else:l78c.Intensity = 6:End if
    If l78d1.Intensity >= 10 Then l78d.Intensity = 0:Else:l78d.Intensity = 30:End if
    If l78e1.Intensity >= 10 Then l78e.Intensity = 0:Else:l78e.Intensity = 10:End if
  End if
  If BorgMod = 0 then
    If l78borga1.Intensity >= 10 Then l78borga.Intensity = 0:Else:l78borga.Intensity = 15:End if
    If l78borgb1.Intensity >= 10 Then l78borgb.Intensity = 0:Else:l78borgb.Intensity = 50:End if
    If l78borgc1.Intensity >= 10 Then l78borgc.Intensity = 0:Else:l78borgc.Intensity = 60:End if
    If l78borgd1.Intensity >= 10 Then l78borgd.Intensity = 0:Else:l78borgd.Intensity = 30:End if
    If l78borge1.Intensity >= 10 Then l78borge.Intensity = 0:Else:l78borge.Intensity = 50:End if
  End if
End Sub

'**************
'***RGB Mode***
'**************

Dim RGBStep, RGBFactor, Red, Green, Blue

RGBStep = 0
RGBFactor = 1
Red = 255
Green = 0
Blue = 0

Sub RGBTimer_timer 'rainbow light color changing
    Select Case RGBStep
        Case 0 'Green
            Green = Green + RGBFactor
            If Green > 255 then
                Green = 255
                RGBStep = 1
            End If
        Case 1 'Red
            Red = Red - RGBFactor
            If Red < 0 then
                Red = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            Blue = Blue + RGBFactor
            If Blue > 255 then
                Blue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            Green = Green - RGBFactor
            If Green < 0 then
                Green = 0
                RGBStep = 4
            End If
        Case 4 'Red
            Red = Red + RGBFactor
            If Red > 255 then
                Red = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            Blue = Blue - RGBFactor
            If Blue < 0 then
                Blue = 0
                RGBStep = 0
            End If
    End Select
    'Light1.color = RGB(Red\10, Green\10, Blue\10)
'    light1.colorfull = RGB(Red, Green, Blue)

  If RGBBumpers = 1 then
    lbumperr.colorfull = RGB(Red, Green, Blue)
    lbumperr1.colorfull = RGB(Red, Green, Blue)
    lbumperr2.colorfull = RGB(Red, Green, Blue)
    lbumperr3.colorfull = RGB(Red, Green, Blue)
    lbumperr4.colorfull = RGB(Red, Green, Blue)
    lbumperr5.colorfull = RGB(Red, Green, Blue)
  End if
  If RGBArrows = 1 then
    l42.colorfull = RGB(Red, Green, Blue)
    l42b.colorfull = RGB(Red, Green, Blue)
    l46.colorfull = RGB(Red, Green, Blue)
    l46b.colorfull = RGB(Red, Green, Blue)
    l54.colorfull = RGB(Red, Green, Blue)
    l54b.colorfull = RGB(Red, Green, Blue)
    l61.colorfull = RGB(Red, Green, Blue)
    l61b.colorfull = RGB(Red, Green, Blue)
    l64.colorfull = RGB(Red, Green, Blue)
    l64b.colorfull = RGB(Red, Green, Blue)
    l71.colorfull = RGB(Red, Green, Blue)
    l71b.colorfull = RGB(Red, Green, Blue)
    l76.colorfull = RGB(Red, Green, Blue)
    l76b.colorfull = RGB(Red, Green, Blue)
  End if
  'Gi
' gi4.colorfull = RGB(Red, Green, Blue)
' gis7.colorfull = RGB(Red, Green, Blue)
' gi8.colorfull = RGB(Red, Green, Blue)
' gis8.colorfull = RGB(Red, Green, Blue)

End Sub


'***********
' Update GI
'***********

Dim gistep, xx, obj
   gistep = 1 / 8

Sub UpdateGI(no, step)
  If step > 1 Then
    DOF 200, DOFOn
  Else
    DOF 200, DOFOff
  End If
    Select Case no
        Case 0
            For each xx in St1Shields:xx.IntensityScale = gistep*step:next
        Case 1
            For each xx in St2Gi1:xx.opacity = gistep*step*50:next
      Table1.ColorGradeImage = "grade_" & step
        Case 2
            For each xx in St3Gi2:xx.opacity = gistep*step*50:next
      Table1.ColorGradeImage = "grade_" & step
        Case 3
            For each xx in St4PFGI:xx.IntensityScale = gistep*step:next
      Table1.ColorGradeImage = "grade_" & step
      For each xx in St4PFGI: if xx.IntensityScale = 0 then Table1.ColorGradeImage = "grade_1":End if:next
        Case 4
            For each xx in St5ReLa:xx.IntensityScale = gistep*step:next
      Table1.ColorGradeImage = "grade_" & step
      For each xx in St5ReLa: if xx.IntensityScale = 0 then Table1.ColorGradeImage = "grade_1":End if:next
    End Select
End Sub

'****************************
' Timers
'****************************

Sub FrameTimer_Timer()
    'flippers,gates
    FlipperLP.RotY = LeftFlipper.CurrentAngle
    FlipperLShadow.RotZ = LeftFlipper.CurrentAngle
    FlipperRP.RotY = RightFlipper.CurrentAngle
    FlipperRShadow.RotZ = RightFlipper.CurrentAngle
  FlipperRP1.RotY = FlipperR1.CurrentAngle
    FlipperR1Shadow.RotZ = FlipperR1.CurrentAngle
  diverterp.RotY = diverter.CurrentAngle
  sw25spinnerp.RotX = sw25Gate.CurrentAngle + 90 'sw25spinner.CurrentAngle +90
  sw87spinnerp.RotX = -sw87Gate.CurrentAngle + 90
  sw88spinnerp.RotX = sw88Gate.CurrentAngle + 90'sw88spinner.CurrentAngle +90

  CannonDomeL.blenddisablelighting = f52.intensityscale
  CannonDomeR.blenddisablelighting = f82.intensityscale

    ' other stuff
    'rollingupdate
  BallShadowUpdate
  BorgLampTimer
  CatapultTimer
  KickbackTimer
  CannonPinRTimer
  CannonPinLTimer
  'Cor.Update
End Sub

Sub GameTimer_Timer()
  Cor.Update            'update ball tracking (this sometimes goes in the RDampen_Timer sub)
  RollingUpdate         'update rolling sounds
  'DoDTAnim             'handle drop target animations
  'DoSTAnim           'handle stand up target animations
End Sub


'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************
'

Const tnob = 6 ' total number of balls
Const lob = 0   'number of locked balls

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
    'If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND (BOT(b).z < 30 AND BOT(b).z > 0) Then
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
  Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


'*******************
' WireRampDropSounds
'*******************

Dim WireSound

Sub LeftRampStart_Hit()
  'Debug.Print "LeftRampStart Hit"
  WireRampOn True
End Sub

Sub LeftPopperTriggerStart_Hit()
  'Debug.Print "LeftPoperTriggerStart Hit"
  WireSound = True
  'WireRampOn False
  If Not OriginalWireRampSounds Then
    'Debug.Print "Play WireRampHit " & VolPlayfieldRoll(ActiveBall) * volumedial * 5
    PlaySoundAtVol "WireRamp_Hit", LeftPopperTriggerStart, VolPlayfieldRoll(ActiveBall) * volumedial * 5
  End If
  WireRampOnPlayAtVolume False, 0.8
End Sub

Sub LRStart_Hit
  WireRampOn True
End Sub

Sub LeftWireTriggerStart_Hit()
  WireSound = True
  'Debug.Print "LeftWireTriggerStart_hit"
  WireRampOff
  'WireRampOn False
  If Not OriginalWireRampSounds Then
    'Debug.Print "Play WireRampHit " & VolPlayfieldRoll(ActiveBall) * volumedial * 5
    PlaySoundAtVol "WireRamp_Hit", LeftWireTriggerStart, VolPlayfieldRoll(ActiveBall) * volumedial * 5
  End If
  WireRampOnPlayAtVolume False, 0.8
End Sub

Sub LeftWireTriggerEnd_Hit()
  WireSound = False
  WireRampOff
  RandomSoundWireRampStop LeftWireTriggerEnd
End Sub

Sub LeftWireRampGunStart_Hit()
  If Not OriginalWireRampSounds Then
    'Debug.Print "Play WireRampHit " & VolPlayfieldRoll(ActiveBall) * volumedial * 5
    PlaySoundAtVol "WireRamp_Hit", LeftWireRampGunStart, VolPlayfieldRoll(ActiveBall) * volumedial * 5
  End If
  WireRampOn False
End Sub

Sub LeftWireRampGunEnd_Hit()
  WireRampOff
  'RandomSoundWireRampStop LeftWireRampGunEnd
End Sub

Sub RightWireRampGunStart_Hit()
  If Not OriginalWireRampSounds Then
    'Debug.Print "Play WireRampHit " & VolPlayfieldRoll(ActiveBall) * volumedial * 5
    PlaySoundAtVol "WireRamp_Hit", RightWireRampGunStart, VolPlayfieldRoll(ActiveBall) * volumedial * 5
  End If
  WireRampOn False
End Sub

Sub RightWireRampGunEnd_Hit()
  WireRampOff
End Sub

Sub RightWireTriggerStart_Hit()
  WireSound = True
  WireRampOff
  If Not OriginalWireRampSounds Then
    'Debug.Print "Play WireRampHit " & VolPlayfieldRoll(ActiveBall) * volumedial * 5
    PlaySoundAtVol "WireRamp_Hit", RightWireTriggerStart, VolPlayfieldRoll(ActiveBall) * volumedial * 5
  End If
  WireRampOn False
End Sub

Sub RightWireTriggerMiddle_Hit()
  If Not OriginalWireRampSounds Then
    'Debug.Print "Play WireRampHit " & VolPlayfieldRoll(ActiveBall)* volumedial * 5
    PlaySoundAtVol "WireRamp_Hit", RightWireTriggerMiddle, VolPlayfieldRoll(ActiveBall) * volumedial * 5
  End If
  WireRampOn False
End Sub

Sub RightWireTriggerEnd_Hit()
  WireSound = False
  WireRampOff
  RandomSoundWireRampStop RightWireTriggerEnd
End Sub

Sub MiddleWireTriggerStart_Hit()
  WireSound = True
  WireRampOff
  'WireRampOn False
  If Not OriginalWireRampSounds Then
    'Debug.Print "Play WireRampHit " & VolPlayfieldRoll(ActiveBall)* volumedial * 5
    PlaySoundAtVol "WireRamp_Hit", MiddleWireTriggerStart, VolPlayfieldRoll(ActiveBall) * volumedial * 5
  End If
  WireRampOnPlayAtVolume False, 0.8
End Sub

Sub MiddleWireTriggerEnd_Hit()
  WireSound = False
  WireRampOff
  RandomSoundWireRampStop MiddleWireTriggerEnd
End Sub

Sub BorgRampEnd_Hit()
  WireRampOff
  RandomSoundRampStop BorgRampEnd
End Sub

Sub PlungerWireTriggerStart_Hit()
  WireSound = True
  RandomSoundBallRelease PlungerWireTriggerStart
  'WireRampOn False
  If LowerBallLaunchSound Then
    LaunchRampOn
  Else
    WireRampOn False
  End If
End Sub

Sub PlungerWireTriggerEnd_Hit()
  WireSound = False
  WireRampOff
  RandomSoundWireRampStop PlungerWireTriggerEnd
  'PlaySoundAT "fx_balldrop6", PlungerWireTriggerEnd
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub SoundNeutralZone(tableObj)
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySoundAt "NeutralZone1", tableobj
        Case 2:PlaySoundAt "NeutralZone2", tableobj
        Case 3:PlaySoundAt "Lock", tableobj
    End Select
End Sub

'*****************************************
' Ball Shadow by Ninuzzo
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3,Ballshadow4,Ballshadow5,Ballshadow6)

Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' hide shadow of deleted balls
  If UBound(BOT)<(tnob-1) Then
    For b = (UBound(BOT) + 1) to (tnob-1)
      BallShadow(b).visible = 0
    Next
  End If
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

'******************************************************
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level well need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
' 4. and, special scripting
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
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity
'dim RF1 : Set RF1 = New FlipperPolarity

InitPolarity



'*******************************************
'  Late 80's early 90's


Sub InitPolarity()
'        dim x, a : a = Array(LF, RF, RF1)
        dim x, a : a = Array(LF, RF)
        for each x in a

      x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 ' disabled
      x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
      x.enabled = True
      x.TimeDelay = 60
      x.DebugOn = True ' prints some info in debugger

      x.AddPt "Polarity", 0, 0, 0
      x.AddPt "Polarity", 1, 0.05, -5.5
      x.AddPt "Polarity", 2, 0.40, -5.5
      x.AddPt "Polarity", 3, 0.60, -5.0
      x.AddPt "Polarity", 4, 0.65, -4.5
      x.AddPt "Polarity", 5, 0.70, -4.0
      x.AddPt "Polarity", 6, 0.75, -3.5
      x.AddPt "Polarity", 7, 0.80, -3.0
      x.AddPt "Polarity", 8, 0.85, -2.5
      x.AddPt "Polarity", 9, 0.90, -2.0
      x.AddPt "Polarity", 10,0.95, -1.5
      x.AddPt "Polarity", 11,1.00, -1.0
      x.AddPt "Polarity", 12,1.05, -0.5
      x.AddPt "Polarity", 13,1.10, 0
      x.AddPt "Polarity", 14,1.30, 0

      x.AddPt "Velocity", 0, 0.000, 1
      x.AddPt "Velocity", 1, 0.160, 1.06
      x.AddPt "Velocity", 2, 0.410, 1.05
      x.AddPt "Velocity", 3, 0.530, 1 ' 0.982
      x.AddPt "Velocity", 4, 0.702, 0.968
      x.AddPt "Velocity", 5, 0.950, 0.968
      x.AddPt "Velocity", 6, 1.030, 0.945

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
' Otherwise it should function exactly the same as before

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt        ' Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay        ' delay before trigger turns off and polarity is disabled
  private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
  Private Balls(20), balldata(20)
  private Name

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    if typename(aName) <> "String" then msgbox "FlipperPolarity: .SetObjects error: first argument must be a string (and name of Object). Found:" & typename(aName) end if
    if typename(aFlipper) <> "Flipper" then msgbox "FlipperPolarity: .SetObjects error: second argument must be a flipper. Found:" & typename(aFlipper) end if
    if typename(aTrigger) <> "Trigger" then msgbox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & typename(aTrigger) end if
    if aFlipper.EndAngle > aFlipper.StartAngle then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper : FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    dim str : str = "sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'" ' automatically create hit / unhit events if uncommented
    ExecuteGlobal(str)
    str = "sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  Public Property Let EndPoint(aInput) :  : End Property ' Legacy: just no op

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) ' Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
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

  Public Sub ProcessBalls() ' save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        balldata(x).Data = balls(x)
        exit for ' nf 2023
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle)) ' might div0
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        ' Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

      ' y safety Exit
      if aBall.VelY > -8 then 'ball going down
        RemoveBall aBall
        exit Sub
      end if

      ' Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                ' find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then ' no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                ' find safety coefficient 'ycoef' data
      End If

      ' Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      ' Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      if DebugOn then debug.print "PolarityCorrect" & " " & Name & " @ " & gametime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class




'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************
' To add these slingshot corrections:
'   - On the table, add the endpoint primitives that define the two ends of the Slingshot
' - Initialize the SlingshotCorrection objects in InitSlingCorrection
'   - Call the .VelocityCorrect methods from the respective _Slingshot event sub


dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection

  LS.Object = SlingShotLeft
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = SlingshotRight
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

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function

Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function

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
'  FLIPPER POLARITY. RUBBER DAMPENER, AND SLINGSHOT CORRECTION SUPPORTING FUNCTIONS
'******************************************************


Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  'dim a : a = Array(LF, RF, RF1)
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub


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
  'FlipperTricks FlipperR1, RF1Press, RF1Count, RF1EndAngle, RF1State
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b, gBOT
  gBOT = GetBalls

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
          gBOT(b).velx = gBOT(b).velx / 1.3
          gBOT(b).vely = gBOT(b).vely - 0.5
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
'Dim PI: PI = 4*Atn(1)

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

dim LFPress, RFPress, RF1Press, LFCount, RFCount, RF1Count
dim LFState, RFState, RF1State
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, RF1EndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
'RF1State = 1
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
'RF1EndAngle = FlipperR1.endangle

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
    Dim b, gBOT
    gBOT = GetBalls

    For b = 0 to UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= -0.4 Then gBOT(b).vely = -0.4
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
  'TargetBouncer Activeball, 0.7
End Sub

' Collection to do target bouner for sleeves
Sub tbSleeves_Hit(idx)
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



'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
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

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
' Cor.Update
'End Sub



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

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
Sub TargetBounce_Hit(idx)
  TargetBouncer activeball, 1
End Sub






'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

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
'RubberFlipperSoundFactor = 0.075/5                   'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.08                   'volume multiplier; must not be zero
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
  'TargetBouncer activeball, 1
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

Sub RandomSoundWireRampStop(obj)
  Select Case Int(rnd*8)
    Case 0: PlaySoundAtVol "wireramp_stop", obj, 0.2 * VolPlayfieldRoll(ActiveBall) * volumedial * 2
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2 * VolPlayfieldRoll(ActiveBall) * volumedial * 2
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2 * VolPlayfieldRoll(ActiveBall) * volumedial * 2
    Case 3: PlaySoundAtVol "wireramp_stop", obj, 0.1 * VolPlayfieldRoll(ActiveBall) * volumedial * 2
    Case 4: PlaySoundAtVol "wireramp_stop2", obj, 0.1 * VolPlayfieldRoll(ActiveBall) * volumedial * 2
    Case 5: PlaySoundAtVol "wireramp_stop3", obj, 0.1 * VolPlayfieldRoll(ActiveBall) * volumedial * 2
    Case 6: PlaySoundAtVol "WireRamp_Hit", obj, 0.1 * VolPlayfieldRoll(ActiveBall) * volumedial * 2
    Case 7: PlaySoundAtVol "WireRamp_Hit", obj, 0.2 * VolPlayfieldRoll(ActiveBall) * volumedial * 2
  End Select
End Sub

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), VolPlayfieldRoll(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), VolPlayfieldRoll(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), VolPlayfieldRoll(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

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
dim RampBalls(7,4)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(7)
dim LaunchRamp : LaunchRamp = False
dim RampPlayAtVolume : RampPlayAtVolume = 0

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub
Sub LaunchRampOn() : LaunchRamp = True : Waddball ActiveBall, False : RampRollUpdate : End Sub
Sub WireRampOnPlayAtVolume(input, wrvol) :  RampPlayAtVolume = wrvol : Waddball ActiveBall, input :RampRollUpdate : End Sub


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
      If LaunchRamp Then
        If OriginalWireRampSounds Then
          RampBalls(x, 3) = False
        Else
          RampBalls(x, 3) = True
        End If
        LaunchRamp = False
      Else
        RampBalls(x, 3) = False
        LaunchRamp = False
      End If
      If RampPlayAtVolume > 0 Then
        'Debug.Print "Set RampPlayAtVolume to " & RampPlayAtVolume & " for RampBalls("&x&",4)"
        RampBalls(x, 4) = RampPlayAtVolume
        RampPlayAtVolume = 0
      Else
        'Debug.Print "RampPlayAtVolume = 0 for RampBalls("&x&",4)"
        RampBalls(x,4) = 0
      End If
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
      RampBalls(x, 3) = False
      RampBalls(x, 4) = 0
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
      StopSound("WireLoopBallLaunch")
      StopSound("WireRamp" & x)
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
        'Debug.Print "RampRole Volume = " & VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * 10
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * 10, AudioPan(RampBalls(x,0))*3, 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
          StopSound("WireRamp" & x)
        Else
          StopSound("RampLoop" & x)
          if RampBalls(x,3) Then
            PlaySound("WireLoopBallLaunch"), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          Else
            If OriginalWireRampSounds Then
              'Debug.print "Play wireRamp Volume = " & VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial * 4
              If (RampBalls(x,4) > 0) Then
                'Debug.Print "RampBalls("&x&",4) = " & RampBalls(x,4)
                PlaySound("WireRamp" & x), -1, RampBalls(x,4), AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
              Else
                'Debug.Print "RampBalls = 0 for ball x = " & x
                PlaySound("WireRamp" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial * 4, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
              End If
            Else
'             If (RampBalls(x,4) > 0) Then0
'               PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampBalls(x,4), AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
'             Else
                PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * 10, AudioPan(RampBalls(x,0))*3, 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
                '' PlaySound("wireloop" & x), -1, Vol(RampBalls(x,0) )*10, Pan(RampBalls(x,0) )*3, 0, BallPitch(RampBalls(x,0) ), 1, 0,Fade(RampBalls(x,0) )
              'End If
            End If
          End If
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        StopSound("WireLoopBallLaunch")
        StopSound("WireRamp" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        'if RampBalls(x, 3) Then
          StopSound("WireLoopBallLaunch")
        StopSound("WireRamp" & x)
        'End If
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
      'if RampBalls(x, 3) Then
        StopSound("WireLoopBallLaunch")
      StopSound("WireRamp" & x)
      'End If
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



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************




'**********************************************************
'*******Set Up VR Backglass and Backglass Flashers  *******
'**********************************************************

Sub SetBackglass()

BGDark.visible = True

Dim VRobj

  For Each VRobj In VRBackglass
    VRobj.x = VRobj.x
    VRobj.height = - VRobj.y + 250
    VRobj.y = -60 'adjusts the distance from the backglass towards the user
    VRobj.rotx = -86.5
  Next

End Sub


'VRRoom Stuff below... **************************************************************************************************************************

'VRtimers...

Sub DoorsOpenTimer_Timer()
if DoorMoving = False and DoorsOpen = False then
DoorMoving = True
Playsound "HDOpen"
end if
If HD_DoorRT.y <= 2100 Then HD_DoorRT.y = HD_DoorRT.y + 10:HD_DoorLT.y = HD_DoorLT.y - 10:HD_DoorRT.z = HD_DoorRT.z + 1:HD_DoorLT.z = HD_DoorLT.z - 1
if HD_DoorRT.y > 2100 Then DoorsOpenTimer.enabled = false: DoorsOpen = True :DoorMoving = False ' turn itself off
End sub

Sub DoorsCloseTimer_Timer()
if DoorMoving = False and DoorsOpen = True then
DoorMoving = True
Playsound "HDClose"
end if
If HD_DoorRT.y => -470 Then HD_DoorRT.y = HD_DoorRT.y - 10:HD_DoorLT.y = HD_DoorLT.y + 10:HD_DoorRT.z = HD_DoorRT.z - 1:HD_DoorLT.z = HD_DoorLT.z + 1
if HD_DoorRT.y < -470 Then DoorsCloseTimer.enabled = false: DoorsOpen = false :DoorMoving = False  ' turn itself off
End sub

Sub RedAlertTimer_Timer()
If HD_LiteRed.disablelighting = 0 then
HD_LiteRed.disablelighting = 2
Else
HD_LiteRed.disablelighting = 0
End if
End Sub

Sub LcarsTimer_Timer()
if HD_Lcars1.image = "LCARSHolo1" Then
HD_Lcars1.image = "LCars3"
else
HD_Lcars1.image = "LCARSHolo1"
end If
End Sub

Sub HolodeckMoveLeftTimer_Timer()
if HoloPosition > -9500 then
For Each HDprim in HoloDeck
HDprim.x = HDprim.x + 10
Next
HoloPosition = HoloPosition - 10
if HoloPosition =< 3400 then
DoorsOpenTimer.enabled = false
DoorsCloseTimer.enabled = true
end if
end if
End Sub


Sub HolodeckMoveRightTimer_Timer()
if HoloPosition < 5390 then
For Each HDprim in HoloDeck
HDprim.x = HDprim.x - 10
Next
HoloPosition = HoloPosition + 10
if HoloPosition => 3500 then
DoorsOpenTimer.enabled = true
DoorsCloseTimer.enabled = false
end if
end if
End Sub


Sub VRSignsTimer_Timer
VRSignsTimer.enabled = False
VRSignsTimer2.enabled = True
End Sub

Sub VRSignsTimer2_Timer
VRRight.opacity = VRright.opacity - 0.51
VRLeft.opacity = VRLeft.opacity - 0.51
If VRRight.opacity =< 0 then VRSignsTimer2.enabled = false: VRRight.visible = False:  VRLeft.visible = False
End Sub




'Knorr & Clark Kent
'V1.0       First Release for Visual Pinball 10.2
'V1.1       Quick Fix for the SpiralRamp
'V1.2       Added more lights
'         updated physics
'         updated ballshadow and ballrolling script from Ninuzzu --- Thanks!
'         reduced polycount
'         added new sounds
'V1.3       Added SurroundSound
'         New Playfield Image
'         minor bug fixes
'         changed HDR enviroment
'         changed "useSolenoids" for fastflips, however... the flippers will move while video mode which does not happen on the real table.
'V1.4       new ramp decals
'         fixed some ball flying issues
'         changed some SoundFX stuff
'VPW Mod
'v01 - fluffhead35 - Added Flipper Triggers, Rubbers and Posts, Bumpers, slings, flipper, and table physics corrections.  Adding in fliper and physic damperner code.
'                  - Adding in materials for all code
'v02 - fluffhead35 - Added Fleep sound
'v03 - fluffhead35 - finished fleep sound and fixing upper flipper, chanign flipper polarity to early 90's and later. Added Sling correction.  Increased plunger strength.
'                  - Added logic to stop ball rolling sounds in subways.  Set FlipperCoilRampupMode to 2
'v05 - leojreimroc - Imported Rawd's Holodeck VR Room.  Adjusted Laser barriers for VR.
'v06 - fluffhead35 - fixed flipper triggers.  Rubber thickness fixed on upper flipper.  Increased flipper hit sound level.  Updated sling rubbers to use bottom corner for post pass.
'v07 - clarkkent   - deleted zCol_Rubber_Sleeve008, RubberPost4,  SubwayScoop_Prim
'v08 - fluffhead35 - Made Posts for post pass collidable, adjusted flipers to be in line with guide.  Adjust playfield friction to .15.  Adjusted sling posts physics for post pass.
'v09 - apophis     - Added new playfield mesh. Increased speed of slingshot animations.
'v10 - Rawd        - Added ClarkKent Cabinet artwork, some VR fixes and tweaks
'v11 - leojreimroc - Implemented Iaaki's flasher code to all solenoid flashers.  Fixed a few flasher positions that were past the cabinet.  Implemented VR Backglass lighting.
'v13 - fluffhead35 - Added new ball launch sound to help prevent high bass sounds.  Adjusted size of sleave and physic materials of start mission scoop.  Removed ball dampening triggers.
'v14 - fluffhead35 - updated the sling posts for post pass based on clark kents suggestions.  Changed Table difficulty to 56. Added option to switch if ball launch should have bass lowered.
'v15 - fluffhead35 - Set default wire Ramp sounds to be from clark kents samples.  Added option to use the other samples if user wants.  Changed ball out to brighter ball.  Truned off flipper corrections on upper right Flipper.
'                  - Changed environmnet emission image to be shinyenvironment3blur4.  Adjusted gi arround inlnaes and inline insert lightings.
'15d - set the falloff of gi lights to 5. Set the day night to 1, set envionrment lighting to black, adjusted warbird material, adjusted ligting arround upper lanes.  Adjusted some random insert ligting.  Adjusted ramp friction based on clark kents suggestions.
'15e - changed elacticty and scatter of sw57 to give ball more random bounces and goes down inlines based on BountyBobs testing
'16  - fluffhead35 - merged in nfozzy changed to init throughts at game start as well as using new flipper physics.  Removed target bouncer from sleeve in start mission scoop.  Lowered bumpercap3 from 25 to 20. Added TargetBouncer to Targets_hit
'17  - Sixtoe    - Added VR logo back in, hooked up cannon light primitives to lighting system.

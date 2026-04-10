' Bally's Radical / IPD No. 1904 / September, 1990 / 4 Players
' This tabls is based on the Radical prototype.
' The prototype playfield included an additional slingshot and a vertical up-kicker not present on production games.
' VPX version by JPSalas, 2018, version 1.1.1
' VPin Workshop MOD, 2020

'xxx - benji - New plastics, fleeps sounds, physics, various versions
'009 - iaakki - new cutout pf and insert text images, most of the insert types in, many inserts needs adjust, renaming and linking
'010 - iaakki - some more inserts...
'011 - iaakki - rest of the inserts layed, lights needs tuning, pf and insert images updated one more time, need to add additional lights for some inserts
'012 - iaakki - insert work finished
'013 - iaakki - Flupper domees... RampDecals to non-static
'014 - iaakki - Flupper Bumpers
'017 - iaakki - Desktop POV fixed, dropsounds checked, side blade reflections checked, some djent to those standup targets, DisableUpperSling option added
'RC1 - Sixtoe - Fixed the backglass cabinet lamps, tons of small corrections and tweaks to primitives, new holes in playfield, tweaked the vr room, fixed some depth bias issues,
'changed the drop target material so they pop a little more, changed table lighting settings
'RC2 - iaakki - Added DL to sideblades and ramps for GI update. Minor tweak to plunger lane. Flasher options fine tune. Reflections on for bats.
'RC3 - iaakki - bumper tune up, debug cleanup..
'v100 - iaakki - release
'v101 - iaakki - CabinetMode off by default

Option Explicit
Randomize

' Thalamus - added more randomness to kickers

'Disables the upper sling making the table more like the retail version instead of prototype
Const DisableUpperSling = 1

'///////////////////////-----VR Room-----///////////////////////
VRRoom = 0 ' 0 - VR Room off, 1 - Minimal Room, 2 - Ultra Minimal

'SIDE RAILS
'   Hide Side Rails = 0
'   Show Side Rails = 1
SideRails = 1

'Cabinet mode - Will hide the rails and scale the side panels higher
' Cabinet Mode Off = 0
' Cabinet Mode On = 1
CabinetMode = 0

'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.8

'Make flipper solenoid to alter active ball behavior
Const FlipSolNudge = 1

Const slackyflips = 1

'Live catch window size. Higher value will make live catching easier. 8-32
Const LiveCatch = 16

Const VolDiv = 100    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolFlip   = 1    ' Flipper volume.
Const VolBall   = 1.5    ' Ball volume.

'*** Ramps Lighting Correction *** ---> handled in GI now
'Primitive83.blenddisablelighting = .25
'Primitive82.blenddisablelighting = .25
'Primitive80.blenddisablelighting = .25
'***************

Const BallSize = 50
Const BallMass = 1.4

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, dtBankM, dtBankT, bsLock, bsTP, plungerIM, x

Const cGameName = "radcl_l1" 'Latest rom
'Const cGameName = "radcl_g1" 'German
'Const cGameName = "radcl_p3"  'Prototype home

Const UseSolenoids = 2 'Fastflips enabled
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden, UseVPMDMD
If Table1.ShowDT = true then
    UseVPMDMD = True
    VarHidden = 1
'    Next
else
    UseVPMDMD = False
    VarHidden = 0
'    Next
end if

if B2SOn = true then VarHidden = 1

LoadVPM "01550000", "S11.vbs", 3.26

Dim DesktopMode: DesktopMode = Table1.ShowDT

' Standard Sounds
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Radical - Bally 1990" & vbNewLine & "VPX table by jpsalas. VPW Mod"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 1
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 1000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot, RightSlingshot1)
    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 10, 11, 12, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        '.InitEntrySnd "fx_drain", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 2
    End With

    ' Visible Lock
    Set bsLock = new cvpmVLock
    With bsLock
        .InitVLock Array(sw40, sw39), Array(sw40k, sw39k), Array(40, 39)
        .InitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .ExitDir = 0
        .ExitForce = 21
        .CreateEvents "bsLock"
    End With

    ' Top Eject Hole 'Not available in the latest ROMs, we activate it manually
    Set bsTP = New cvpmBallStack
    With bsTP
    .InitSw 0, 16, 0, 0, 0, 0, 0, 0
    .InitKick sw16, 0, 55
    .KickZ = 1.5
    .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    .KickForceVar = 3
    .KickAngleVar = 3
    End With

    ' Drop Middle targets
    set dtbankM = new cvpmdroptarget
    With dtbankM
        .initdrop array(sw22, sw23, sw24), array(22, 23, 24)
        .initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtbankM"
    End With

    ' Drop Top targets
    set dtbankT = new cvpmdroptarget
    With dtbankT
        .initdrop array(sw30, sw31, sw32), array(30, 31, 32)
        .initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtbankT"
    End With

    ' Impulse Plunger used for kickback
    Const IMPowerSetting = 42 'Plunger Power
    Const IMTime = 0.3        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP sw51, IMPowerSetting, IMTime
        .Random 0.3
        .switch 51
        .InitExitSnd "fx_Solenoid", "fx_Solenoid"
        .CreateEvents "plungerIM"
    End With

    ' Init diverters
    DiverterDown.IsDropped = 1
    RightDiverter1.IsDropped = 1

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    vpmTimer.AddTimer 2000, "GiOn '"
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********
' Keys
'**********


'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback

'nFozzy Physics
  If keycode = LeftFlipperKey Then LFPress = 1
  If keycode = RightFlipperKey Then rfpress = 1
'end nFozzy Physics

    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)

'nFozzy physics
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
' end nFozzy Physics


    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep, Rstep1

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 55
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSling4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 56
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub RightSlingShot1_Slingshot
    if DisableUpperSling = 0 Then
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    r6.Visible = 1
    Remk1.RotX = 26
    RStep1 = 0
    vpmTimer.PulseSw 49
    RightSlingShot1.TimerEnabled = 1
  end if
End Sub

Sub RightSlingShot1_Timer
    Select Case RStep1
        Case 1:r6.Visible = 0:r4.Visible = 1:Remk1.RotX = 14
        Case 2:r4.Visible = 0:r3.Visible = 1:Remk1.RotX = 2
        Case 3:r3.Visible = 0:Remk1.RotX = -20:RightSlingshot1.TimerEnabled = 0
    End Select
    RStep1 = RStep1 + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 52:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.05:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.02:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 54:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.05:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 53:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.02:End Sub

' Drain & Saucers
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub sw16a_Hit:PlaySound "fx_hole_enter", 0, 1, -0.02:bsTP.AddBall Me:vpmTimer.AddTimer 2500, "EjectBall":End Sub
Sub EjectBall(swno):PlaySound "fx_popper", 0, 1, -0.05:bsTP.ExitSol_On:End Sub
Sub RHelp_Hit:ActiveBall.Velx = 10:End Sub 'to ensure the ball always lands on the ramp

' Rollovers
Sub sw51_Hit:Controller.Switch(51) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub

' Spinners
Sub Spinner1_Spin:vpmTimer.PulseSw 29:PlaySound "fx_spinner", 0, 1, 0.05:End Sub
Sub Spinner2_Spin:vpmTimer.PulseSw 17:PlaySound "fx_spinner", 0, 1, -0.05:End Sub

'Targets
dim zMultiplier
Sub sw28_Hit
  Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 2.55
        Case 2: zMultiplier = 2.2
        Case 3: zMultiplier = 1.9
        Case 4: zMultiplier = 1.7
    End Select
    activeball.velz = activeball.velz*zMultiplier
  vpmTimer.PulseSw 28
  PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall)
End Sub

Sub sw43_Hit
  Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 2.55
        Case 2: zMultiplier = 2.2
        Case 3: zMultiplier = 1.9
        Case 4: zMultiplier = 1.7
    End Select
  vpmTimer.PulseSw 43
  PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall)
End Sub

Sub sw44_Hit
  Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 2.55
        Case 2: zMultiplier = 2.2
        Case 3: zMultiplier = 1.9
        Case 4: zMultiplier = 1.7
    End Select
  vpmTimer.PulseSw 44
  PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall)
End Sub

'Droptargets (only sound effect)

Sub sw22_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw23_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw24_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw30_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw31_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw32_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "bsTrough.SolIn"
SolCallback(2) = "bsTrough.SolOut"
SolCallback(3) = "bsTP.SolOut" 'Not used in the latest ROMs
SolCallback(4) = "dtBankT.SolDropUp"
SolCallback(6) = "dtBankM.SolDropUp"
SolCallback(5) = "VpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(7) = "SolRampDiv"
SolCallback(8) = "bsLock.SolExit" 'vlock
SolCallback(10) = "SolGI"
SolCallback(13) = "vpmSolToggleWall RightDiverter2,RightDiverter1,""fx_diverter"","
SolCallback(14) = "SolKickBack"
SolCallback(23) = "vpmNudge.SolGameOn"

' flashers
SolCallback(9) = "FlashPurpleMiddleRight" '"SetLamp 90,"
SolCallback(16) = "SetLamp 116,"
SolCallback(25) = "FlashPurpleLeftBottom"'"SetLamp 125,"
SolCallback(26) = "FlashGreenLeftBottom"'126
SolCallback(27) = "FlashPurpleMiddleLeft"'"SetLamp 127,"
SolCallback(28) = "FlashGreenLeftTop" '"SetLamp 128,"
SolCallback(29) = "FlashWhiteBack" '"SetLamp 129,"
SolCallback(30) = "FlashPurpleMiddleTop"'"SetLamp 130,"
SolCallback(31) = "FlashGreenTop" '"SetLamp 131,"
SolCallback(32) = "FlashGreenRight" '"SetLamp 132,"


Sub FlashGreenRight(flstate)
  If Flstate Then
    Objlevel(1) = 1 : FlasherFlash1_Timer
    Objlevel(2) = 1 : FlasherFlash2_Timer
  End If
End Sub

Sub FlashGreenTop(flstate)
  If Flstate Then
    Objlevel(3) = 1 : FlasherFlash3_Timer
  End If
End Sub

Sub FlashGreenLeftTop(flstate)
  If Flstate Then
    Objlevel(4) = 1 : FlasherFlash4_Timer
  End If
End Sub

Sub FlashGreenLeftBottom(flstate)
  If Flstate Then
    Objlevel(5) = 1 : FlasherFlash5_Timer
  End If
End Sub

Sub FlashPurpleLeftBottom(flstate)
  If Flstate Then
    Objlevel(6) = 1 : FlasherFlash6_Timer
  End If
End Sub

Sub FlashPurpleMiddleLeft(flstate)
  If Flstate Then
    Objlevel(7) = 1 : FlasherFlash7_Timer
  End If
End Sub

Sub FlashPurpleMiddleTop(flstate)
  If Flstate Then
    Objlevel(8) = 1 : FlasherFlash8_Timer
  End If
End Sub

Sub FlashPurpleMiddleRight(flstate)
  If Flstate Then
    Objlevel(9) = 1 : FlasherFlash9_Timer
  End If
End Sub

Sub FlashWhiteBack(flstate)
  If Flstate Then
    Objlevel(10) = 1 : FlasherFlash10_Timer
  End If
End Sub

'Solenoid Subs

Sub SolRampDiv(Enabled)
    If Enabled Then
        If DiverterDown.IsDropped Then
            Controller.Switch(20) = 1
            DiverterDown.IsDropped = 0
            DiverterUp.IsDropped = 1
        Else
            Controller.Switch(20) = 0
            DiverterDown.IsDropped = 1
            DiverterUp.IsDropped = 0
        End If
    End If
End Sub

Sub SolKickBack(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolGi(enabled)
    If enabled Then
        GiOff
    Else
        GiOn
    End If
End Sub

Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
  FlBumperFadeTarget(1) = 0.7
  FlBumperFadeTarget(2) = 0.7
  FlBumperFadeTarget(3) = 0.7
  FlBumperFadeTarget(4) = 0.7
  Primitive80.blenddisablelighting = 0.25
  Primitive82.blenddisablelighting = 0.25
  Primitive83.blenddisablelighting = 0.25
  PinCab_Blades.blenddisablelighting = 0.3
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
  FlBumperFadeTarget(1) = 0
  FlBumperFadeTarget(2) = 0
  FlBumperFadeTarget(3) = 0
  FlBumperFadeTarget(4) = 0
  Primitive80.blenddisablelighting = 0.02
  Primitive82.blenddisablelighting = 0.02
  Primitive83.blenddisablelighting = 0.02
  PinCab_Blades.blenddisablelighting = 0.05

End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
            PlaySoundAtVol SoundFX("TOM_Calle_ReFlip_L0" & Int(Rnd*3)+1, DOFFlippers), LeftFlipper, VolFlip
    Else
      PlaySoundAtVol SoundFX("TOM_Calle_Flipper_Attack-L01", DOFFlippers), LeftFlipper, VolFlip
            PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L0" & Int(Rnd*9)+1, DOFFlippers), LeftFlipper, VolFlip
    End If
    LeftFlipper1.RotateToEnd
    LF.Fire
    Else
        PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_" & Int(Rnd*7)+1, DOFFlippers), LeftFlipper, VolFlip
    LeftFlipper.RotateToStart
    LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    If rightflipper.currentangle < rightflipper.endangle + ReflipAngle Then
            PlaySoundAtVol SoundFX("TOM_Calle_ReFlip_R0" & Int(Rnd*3)+1, DOFFlippers), RightFlipper, VolFlip
    Else
      PlaySoundAtVol SoundFX("TOM_Calle_Flipper_Attack-R01", DOFFlippers), RightFlipper, VolFlip
            PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R0" & Int(Rnd*11)+1, DOFFlippers), RightFlipper, VolFlip
    End If
    RightFlipper1.RotateToEnd
    RF.Fire
    Else
        PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_" & Int(Rnd*8)+1, DOFFlippers), RightFlipper, VolFlip
    RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
    End If
End Sub

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

Sub RightFlipper1_Collide(parm)
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
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

'******************** Slacky Flips **************************'
Dim FlipY

FlipY = rightflipper.y

'Const FlipY = 1810            'Default flip Y position
Const Dislocation = 2        'Amount of minimum slack, affect to recovery speed too
Const SlackHitLimit = 10    '6-16    Lower value will make slack happen with lower collision force
Const HitForceDivider = 50    '10-100    Lower value will add slack. Set to roughly half of the max hit force(parm) you see when playing
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

'******************************************************
'     FLIPPER TRICKS
'******************************************************

Const PI = 3.14159

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  If FlipSolNudge = 1 Then
    FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
    FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
  End If
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim BOT, b

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    If Flipper2.currentangle = EndAngle2 Then
      BOT = GetBalls
        For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.7
          BOT(b).vely = BOT(b).vely - 1
        end If
      Next
    End If
  Else
    If Flipper1.currentangle <> EndAngle1 then EOSNudge1 = 0
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
  AnglePP = Atn2((by - ay),(bx - ax)) * 180/PI
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

'*****************************************************************************************************
'*******************************************************************************************************
'END nFOZZY FLIPPERS'

' #####################################
' ###### Flupper Flasher Domes    #####
' #####################################

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.9   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.2   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.4    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objlight(20)
Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "green"
InitFlasher 2, "green"
InitFlasher 3, "green"
InitFlasher 4, "green"
InitFlasher 5, "green"
InitFlasher 6, "purple"
InitFlasher 7, "purple"
InitFlasher 8, "purple"
InitFlasher 9, "purple"
InitFlasher 10, "white"

' : InitFlasher 2, "red" : InitFlasher 3, "white"
'InitFlasher 4, "green" : InitFlasher 5, "red" : InitFlasher 6, "white"
'InitFlasher 7, "green" : InitFlasher 8, "red"
'InitFlasher 9, "green" : InitFlasher 10, "red" : InitFlasher 11, "white"
' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
RotateFlasher 1,30
RotateFlasher 2,45
RotateFlasher 3,-70
RotateFlasher 4,30
RotateFlasher 5,0
RotateFlasher 6,-60
RotateFlasher 7,-130
RotateFlasher 8,-70
RotateFlasher 9,30


Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 60
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
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objlit(nr).visible = 0 : End If
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
Sub FlasherFlash11_Timer() : FlashFlasher(11) : End Sub

' #####################################
' ###### Flupper Bumpers      #####
' #####################################

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0 : DNA45 = (NightDay-10)/20 : DNA90 = 0 : DayNightAdjust = 0.4
Else
  DNA30 = (NightDay-10)/30 : DNA45 = (NightDay-10)/45 : DNA90 = (NightDay-10)/90 : DayNightAdjust = NightDay/25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt : For cnt = 1 to 6 : FlBumperActive(cnt) = False : Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight

FlInitBumper 1, "red"
FlInitBumper 2, "red"
FlInitBumper 3, "red"
FlInitBumper 4, "red"

' ### uncomment the statement below to change the color for all bumpers ###
' Dim ind : For ind = 1 to 5 : FlInitBumper ind, "green" : next

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True
  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1 : FlBumperFadeTarget(nr) = 1.1: FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr) : FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr) : Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr) : Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr) : FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr): FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
  ' set the color for the two VPX lights
  select case col
    Case "red"
      FlBumperSmallLight(nr).color = RGB(255,4,0) : FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
      FlBumperBigLight(nr).color = RGB(255,32,0) : FlBumperBigLight(nr).colorfull = RGB(255,32,0)
      FlBumperHighlight(nr).color = RGB(64,255,0)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
      FlBumperSmallLight(nr).TransmissionScale = 0
    Case "blue"
      FlBumperBigLight(nr).color = RGB(32,80,255) : FlBumperBigLight(nr).colorfull = RGB(32,80,255)
      FlBumperSmallLight(nr).color = RGB(0,80,255) : FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
      FlBumperSmallLight(nr).TransmissionScale = 0 : MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
      FlBumperHighlight(nr).color = RGB(255,16,8)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    Case "green"
      FlBumperSmallLight(nr).color = RGB(8,255,8) : FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
      FlBumperBigLight(nr).color = RGB(32,255,32) : FlBumperBigLight(nr).colorfull = RGB(32,255,32)
      FlBumperHighlight(nr).color = RGB(255,32,255) : MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
      FlBumperSmallLight(nr).TransmissionScale = 0.005
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    Case "orange"
      FlBumperHighlight(nr).color = RGB(255,130,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).color = RGB(255,130,0) : FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
      FlBumperBigLight(nr).color = RGB(255,190,8) : FlBumperBigLight(nr).colorfull = RGB(255,190,8)
    Case "white"
      FlBumperBigLight(nr).color = RGB(255,230,190) : FlBumperBigLight(nr).colorfull = RGB(255,230,190)
      FlBumperHighlight(nr).color = RGB(255,180,100) :
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
    Case "blacklight"
      FlBumperBigLight(nr).color = RGB(32,32,255) : FlBumperBigLight(nr).colorfull = RGB(32,32,255)
      FlBumperHighlight(nr).color = RGB(48,8,255) :
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    Case "yellow"
      FlBumperSmallLight(nr).color = RGB(255,230,4) : FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
      FlBumperBigLight(nr).color = RGB(255,240,50) : FlBumperBigLight(nr).colorfull = RGB(255,240,50)
      FlBumperHighlight(nr).color = RGB(255,255,220)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
    Case "purple"
      FlBumperBigLight(nr).color = RGB(80,32,255) : FlBumperBigLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).color = RGB(80,32,255) : FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).TransmissionScale = 0 :
      FlBumperHighlight(nr).color = RGB(32,64,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
  end select
End Sub

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
' UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
'               OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
'               float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 )* DayNightAdjust

  select case FlBumperColor(nr)

    Case "blue" :
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(38-24*Z,130 - 98*Z,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20  + 500 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z +0.97 * Z^3)
      Flbumperbiglight(nr).intensity = 45 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 10000 * (Z^3) / (0.5 + DNA90)

    Case "green"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(16 + 16 * sin(Z*3.14),255,16 + 16 * sin(Z*3.14)), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
      FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 6000 * (Z^3) / (1 + DNA90)

    Case "red"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 16 - 11*Z + 16 * sin(Z*3.14),0), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45) '20 -> 10. was too bright
      FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z*4,8-Z*8)

    Case "orange"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 100 - 22*z  + 16 * sin(Z*3.14),Z*32), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z*50, 0)

    Case "white"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
      FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 14 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255,255 - 20*Z,255-65*Z) : FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20*Z,255-65*Z)
      MaterialColor "bumpertopmat" & nr, RGB(255,235 - z*36,220 - Z*90)

    Case "blacklight"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 1, RGB(30-27*Z^0.03,30-28*Z^0.01, 255), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z^3
      Flbumperbiglight(nr).intensity = 40 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255-240*(Z^0.1),255 - 240*(Z^0.1),255) : FlBumperSmallLight(nr).colorfull = RGB(255-200*z,255 - 200*Z,255)
      MaterialColor "bumpertopmat" & nr, RGB(255-190*Z,235 - z*180,220 + 35*Z)

    Case "yellow"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 180 + 40*z, 48* Z), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

    Case "purple" :
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(128-118*Z - 32 * sin(Z*3.14), 32-26*Z ,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 15  + 200 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z +0.97 * Z^3)
      Flbumperbiglight(nr).intensity = 50 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (0.5 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(128-60*Z,32,255)


  end select
End Sub

Sub BumperTimer_Timer
  dim nr
  For nr = 1 to 6
    If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
      If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
    If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
      If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
  next
End Sub

'' ###################################

if DisableUpperSling = 1 Then 'Disabling upper slingshot
  RightSlingshot1.SlingshotThreshold = 1000
end if

'**********************************************************
'     JP's Flasher Fading for VPX and Vpinmame v 1.0
'       (Based on Pacdude's Fading Light System)
' This is a fast fading for the Flashers in vpinmame tables
'**********************************************************

Dim FadingState(200)

InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            FadingState(chgLamp(ii, 0)) = chgLamp(ii, 1) + 3 'fading step
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps
    ' playfield lights
    Lampm 1, li1
  FadeDisableLighting2 1, p1, 20
    Lampm 2, li2
  FadeDisableLighting2 2, p2, 20
    Lampm 3, li3
  FadeDisableLighting2 3, p3, 20
    Lampm 4, li4
  FadeDisableLighting2 4, p4, 20
    Lampm 5, li5
  FadeDisableLighting2 5, p5, 20
    Lampm 6, li6
  FadeDisableLighting2 6, p6, 20
    Lampm 7, li7
  FadeDisableLighting2 7, p7, 20
    Lampm 8, li8
  FadeDisableLighting2 8, p8, 20
    Lampm 9, li9
    Lampm 9, li9a
  FadeDisableLighting2 9, p9, 20
    Lampm 10, li10
    Lampm 10, li10a
  FadeDisableLighting2 10, p10, 20
    Lampm 11, li11
  Lampm 11, li11a
  FadeDisableLighting2 11, p11, 20
    Lampm 12, li12
  Lampm 12, li12a
  FadeDisableLighting2 12, p12, 20
    Lampm 13, li13
    Lampm 13, li13a
  FadeDisableLighting2 13, p13, 20
    Lampm 14, li14
  Lampm 14, li14a
  FadeDisableLighting2 14, p14, 20
    Lampm 15, li15
  Lampm 15, li15a
  FadeDisableLighting2 15, p15, 20
    Lampm 16, li16
  FadeDisableLighting2 16, p16, 20
    Lampm 17, li17
    Lampm 17, li17a
  FadeDisableLighting2 17, p17, 60
    Lampm 18, li18
    Lampm 18, li18a
  FadeDisableLighting2 18, p18, 60
    Lampm 19, li19
    Lampm 19, li19a
  FadeDisableLighting2 19, p19, 60
    Lampm 20, li20
    Lampm 20, li20a
  FadeDisableLighting2 20, p20, 60
    Lampm 21, li21
    Lampm 21, li21a
  FadeDisableLighting2 21, p21, 60
    Lampm 22, li22
    Lampm 22, li22a
  FadeDisableLighting2 22, p22, 60
    Lampm 23, li23
    Lampm 23, li23a
  FadeDisableLighting2 23, p23, 60
    Lampm 24, li24
  FadeDisableLighting2 24, p24, 60
    Lampm 25, li25
  FadeDisableLighting2 25, p25, 40
    Lampm 26, li26
  FadeDisableLighting2 26, p26, 40
    Lampm 27, li27
  FadeDisableLighting2 27, p27, 40
    Lampm 28, li28
  FadeDisableLighting2 28, p28, 40
    Lampm 29, li29
  FadeDisableLighting2 29, p29, 40
    Lampm 30, li30
  FadeDisableLighting2 30, p30, 60
    Lampm 31, li31
  FadeDisableLighting2 31, p31, 60
    Lampm 32, li32
  FadeDisableLighting2 32, p32, 60
    Lampm 33, li33
  FadeDisableLighting2 33, p33, 20
    Lampm 34, li34
  FadeDisableLighting2 34, p34, 20
    Lampm 35, li35
  FadeDisableLighting2 35, p35, 20
    Lampm 36, li36
    Lampm 36, li36a
  FadeDisableLighting2 36, p36, 20
    Lampm 37, li37
  FadeDisableLighting2 37, p37, 70
    Lampm 38, li38
  FadeDisableLighting2 38, p38, 60
    Lampm 39, li39
  FadeDisableLighting2 39, p39, 60
    Lampm 40, li40
  FadeDisableLighting2 40, p40, 70
    Lampm 41, li41
  FadeDisableLighting2 41, p41, 60
    Lampm 42, li42
  FadeDisableLighting2 42, p42, 20
    Lampm 43, li43
  FadeDisableLighting2 43, p43, 50
    Lampm 44, li44
  FadeDisableLighting2 44, p44, 60
    Lampm 45, li45
  FadeDisableLighting2 45, p45, 60
    Lampm 46, li46
  FadeDisableLighting2 46, p46, 15
    Lampm 47, li47
  FadeDisableLighting2 47, p47, 20
    Lampm 48, li48
  FadeDisableLighting2 48, p48, 20
    Lampm 49, li49
  FadeDisableLighting2 49, p49, 60
    Lampm 50, li50
  FadeDisableLighting2 50, p50, 60
    Lampm 51, li51
  FadeDisableLighting2 51, p51, 60
    Lampm 52, li52
  FadeDisableLighting2 52, p52, 60
    Lampm 53, li53
  FadeDisableLighting2 53, p53, 60
    Lampm 54, li54
  FadeDisableLighting2 54, p54, 60
    Lampm 55, li55
  FadeDisableLighting2 55, p55, 60
    Lampm 56, li56
  FadeDisableLighting2 56, p56, 60
    if desktopmode then Lampm 57, li57
    Flash 57, bg1
    if desktopmode then Lampm 58, li58
    Flash 58, bg2
    if desktopmode then Lampm 59, li59
    Flash 59, bg3
    if desktopmode then Lampm 60, li60
    Flash 60, bg4
    if desktopmode then Lampm 61, li61
    Flash 61, bg5
    if desktopmode then Lampm 62, li62
    Flash 62, bg6
    if desktopmode then Lampm 63, li63
  FadeDisableLighting2 63, p63, 20
    Lampm 64, li64
  FadeDisableLighting2 64, p64, 20

    'Flashers
    Lamp 116, f16 'ramp flasher on front left

End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 25 ' flasher fading speed
    LampTimer.Enabled = 1
    For x = 0 to 200
        FadingState(x) = 3 ' used to track the fading state
    Next
End Sub

Sub FadeDisableLighting2(nr, a, alvl)
  Select Case FadingState(nr)
    Case 3
      a.UserValue = a.UserValue - 0.25
      If a.UserValue < 0 Then
        a.UserValue = 0
        FadingState(nr) = 0
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
    Case 4
      a.UserValue = a.UserValue + 0.5
      If a.UserValue > 1 Then
        a.UserValue = 1
        FadingState(nr) = 0
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
  End Select
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingState(nr) = abs(value) + 3
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingState(nr)
        Case 4:object.state = 1:FadingState(nr) = 0
        Case 3:object.state = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:object.state = 1
        Case 3:object.state = 0
    End Select
End Sub

' Flashers: 4 is on,3,2,1 fade steps. 0 is off

Sub Flash(nr, object)
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1:FadingState(nr) = 0
        Case 3:Object.IntensityScale = 0.66:FadingState(nr) = 2
        Case 2:Object.IntensityScale = 0.33:FadingState(nr) = 1
        Case 1:Object.IntensityScale = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1
        Case 3:Object.IntensityScale = 0.66
        Case 2:Object.IntensityScale = 0.33
        Case 1:Object.IntensityScale = 0
    End Select
End Sub

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(32)
Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10) = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11) = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12) = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13) = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14) = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15) = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

Digits(16) = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17) = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18) = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19) = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20) = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21) = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22) = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23) = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24) = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25) = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26) = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27) = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28) = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29) = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30) = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31) = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

'Sub UpdateLeds
'    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
'    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
'    If Not IsEmpty(ChgLED)Then
'        For ii = 0 To UBound(chgLED)
'            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
'            if(num <32)then
'                For Each obj In Digits(num)
'                    If chg And 1 Then obj.State = stat And 1
'                    chg = chg \ 2:stat = stat \ 2
'                Next
'            Else
'            end if
'        Next
'    End If
'End Sub



'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):dPostsHit():End Sub
Sub aRubber_Posts_Hit(idx):dPostsHit():End Sub
Sub aRubber_Pins_Hit(idx):dPostsHit():End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'Sub REnd1_Hit:PlaySound "fx5_ballrampdrop", 0, 1, 0.05, -0.05:End Sub
'Sub REnd2_Hit:PlaySound "fx_ballrampdrop", 0, 1, 0.05, 0.05:End Sub
'Sub REnd3_Hit:PlaySound "fx_balldrop", 0, 1, 0.05, 0.05:End Sub
'Sub REnd4_Hit:PlaySound "fx_ballrampdrop", 0, 1, 0.05, -0.05:End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp> 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 4 ' total number of balls
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
    If (BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27) or (BOT(b).VelZ < -1 and BOT(b).z < 180 and BOT(b).z > 170) Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        'debug.print "drop sound velz: " & BOT(b).velz & " z: " & BOT(b).z
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

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'******************
' RealTime Updates
'******************



Sub FrameTimer_Timer
    RollingSound
  Displaytimer
  RDampen
    LFLogo.RotZ = LeftFlipper.CurrentAngle
    LF1Logo.RotZ = LeftFlipper1.CurrentAngle
    RFLogo.RotZ = RightFlipper.CurrentAngle
    RF1Logo.RotZ = RightFlipper1.CurrentAngle
  PinCab_Shooter.Y = -140 + (5* Plunger.Position) -20
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
'                     Fleep  Supporting Ball & Sound Functions
' *********************************************************************
'Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

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

'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub RubbersHit()
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

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR
Sub RDampen()
  Cor.Update
End Sub

Sub dPostsHit()
    RubbersD.dampen Activeball
  RubbersHit()
End Sub

Sub dSleevesHit()
    SleevesD.Dampen Activeball
End Sub

dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False    'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.935 '0.96 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.935 '0.96
RubbersD.addpoint 2, 5.76, 0.942 '0.967 'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64   'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False    'shows info in textbox "TBPout"
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

' Thalamus - patched : ' Thalamus - patched :         aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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


    Public Sub Report()     'debug, reports all coords in tbPL.text
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

'***************************************************************


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


dim DisplayColor, DisplayColorG
DisplayColor =  RGB(255,40,1)

Sub DisplayTimer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 32) then
        if VRRoom > 0 Then
          For Each obj In VRDigits(num)
          If chg And 1 Then FadeDisplay obj, stat And 1
          chg=chg\2 : stat=stat\2
          Next
        end If
        if DesktopMode = True and VRRoom < 1 Then
                For Each obj In Digits(num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg \ 2:stat = stat \ 2
                Next
        end If
        else
      end if
    Next
    End If
End Sub


Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
  Else
    Object.Color = RGB(1,1,1)
  End If
End Sub

Dim VRDigits(32)

VRDigits(0)=Array(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15)
VRDigits(1)=Array(D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30)
VRDigits(2)=Array(D31,D32,D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45)
VRDigits(3)=Array(D46,D47,D48,D49,D50,D51,D52,D53,D54,D55,D56,D57,D58,D59,D60)
VRDigits(4)=Array(D61,D62,D63,D64,D65,D66,D67,D68,D69,D70,D71,D72,D73,D74,D75)
VRDigits(5)=Array(D76,D77,D78,D79,D80,D81,D82,D83,D84,D85,D86,D87,D88,D89,D90)
VRDigits(6)=Array(D91,D92,D93,D94,D95,D96,D97,D98,D99,D100,D101,D102,D103,D104,D105)
VRDigits(7)=Array(D106,D107,D108,D109,D110,D111,D112,D113,D114,D115,D116,D117,D118,D119,D120)
VRDigits(8)=Array(D121,D122,D123,D124,D125,D126,D127,D128,D129,D130,D131,D132,D133,D134,D135)
VRDigits(9)=Array(D136,D137,D138,D139,D140,D141,D142,D143,D144,D145,D146,D147,D148,D149,D150)
VRDigits(10)=Array(D151,D152,D153,D154,D155,D156,D157,D158,D159,D160,D161,D162,D163,D164,D165)
VRDigits(11)=Array(D166,D167,D168,D169,D170,D171,D172,D173,D174,D175,D176,D177,D178,D179,D180)
VRDigits(12)=Array(D181,D182,D183,D184,D185,D186,D187,D188,D189,D190,D191,D192,D193,D194,D195)
VRDigits(13)=Array(D196,D197,D198,D199,D200,D201,D202,D203,D204,D205,D206,D207,D208,D209,D210)
VRDigits(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218,D219,D220,D221,D222,D223,D224,D225)
VRDigits(15)=Array(D226,D227,D228,D229,D230,D231,D232,D233,D234,D235,D236,D237,D238,D239,D240)

VRDigits(16)=Array(D241,D242,D243,D244,D245,D246,D247,D248,D249,D250,D251,D252,D253,D254,D255)
VRDigits(17)=Array(D256,D257,D258,D259,D260,D261,D262,D263,D264,D265,D266,D267,D268,D269,D270)
VRDigits(18)=Array(D271,D272,D273,D274,D275,D276,D277,D278,D279,D280,D281,D282,D283,D284,D285)
VRDigits(19)=Array(D286,D287,D288,D289,D290,D291,D292,D293,D294,D295,D296,D297,D298,D299,D300)
VRDigits(20)=Array(D301,D302,D303,D304,D305,D306,D307,D308,D309,D310,D311,D312,D313,D314,D315)
VRDigits(21)=Array(D316,D317,D318,D319,D320,D321,D322,D323,D324,D325,D326,D327,D328,D329,D330)
VRDigits(22)=Array(D331,D332,D333,D334,D335,D336,D337,D338,D339,D340,D341,D342,D343,D344,D345)
VRDigits(23)=Array(D346,D347,D348,D349,D350,D351,D352,D353,D354,D355,D356,D357,D358,D359,D360)
VRDigits(24)=Array(D361,D362,D363,D364,D365,D366,D367,D368,D369,D370,D371,D372,D373,D374,D375)
VRDigits(25)=Array(D376,D377,D378,D379,D380,D381,D382,D383,D384,D385,D386,D387,D388,D389,D390)
VRDigits(26)=Array(D391,D392,D393,D394,D395,D396,D397,D398,D399,D400,D401,D402,D403,D404,D405)
VRDigits(27)=Array(D406,D407,D408,D409,D410,D411,D412,D413,D414,D415,D416,D417,D418,D419,D420)
VRDigits(28)=Array(D421,D422,D423,D424,D425,D426,D427,D428,D429,D430,D431,D432,D433,D434,D435)
VRDigits(29)=Array(D436,D437,D438,D439,D440,D441,D442,D443,D444,D445,D446,D447,D448,D449,D450)
VRDigits(30)=Array(D451,D452,D453,D454,D455,D456,D457,D458,D459,D460,D461,D462,D463,D464,D465)
VRDigits(31)=Array(D466,D467,D468,D469,D470,D471,D472,D473,D474,D475,D476,D477,D478,D479,D480)


Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each obj in VRDigits(x)
        obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub

InitDigits

'*********************
'Show Side Rails
'*********************
DIM SideRails
If SideRails =1 and VRRoom < 1 Then
    PinCab_Metal_Rails.visible = 1
Else
    PinCab_Metal_Rails.visible = 0
End if

'*********************
'Cabinet Mode
'*********************

DIM Cabinetmode, xx
if CabinetMode = 1 and VRRoom < 1 then
    PinCab_Blades.Size_y=2000
    PinCab_Metal_Rails.visible = 0
    For each xx in aReels
        xx.Visible = 0
    Next
Else
    PinCab_Blades.Size_y=1000
    PinCab_Metal_Rails.visible = 1
    For each xx in aReels
        xx.Visible = 1
    Next
end If

'*********************
'Hide Desktop LCD
'*********************
DIM VRRoom, VRThings
If VRRoom > 0 Then
  If VRRoom = 1 Then
    for each VRThings in aReels:VRThings.visible = 0:Next
    for each VRThings in VR_Stuff:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 1:Next
    PinCab_Metal_Rails.visible = 1
  End If
  If VRRoom = 2 Then
    for each VRThings in aReels:VRThings.visible = 0:Next
    for each VRThings in VR_Stuff:VRThings.visible = 0:Next
    for each VRThings in VR_Min:VRThings.visible = 1:Next
    PinCab_Metal_Rails.visible = 1
    PinCab_Backbox.image = "color_dark"
  End If
Else
    for each VRThings in aReels:VRThings.visible = 1:Next
    for each VRThings in VR_Stuff:VRThings.visible = 0:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
End if

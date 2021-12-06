'
'                    **************************************************************************
'                    *                                                                        *
'                    *                          STERN PINBALL, INC.                           *
'                    *             (C) COPYRIGHT 2011 - 2012 STERN PINBALL, INC.              *
'                    *                                                                        *
'                    **************************************************************************
'                    *                                                                        *
'                    *                               AC/DC Premium                          *
'                    *                                                                        *
'                    **************************************************************************
'
'             *                      *                                       *********                      *
'            ***                    ***                                     ***********                    ***
'           *****                  *****                                   *************                  *****
'          *******                *******                                 ***************                *******
'         *********              *********                               *****************              *********
'        ***********            ***********                             *******************            ***********
'       *************          *************                            ********************          *************
'      ***************        ***************                            ********************        ***************
'     *****************      *****************                            ******** ***********      *****************
'    ****** ************    ******* ***********             ********      ********  **********     ******* ***********
'   *******  ***********    *******  **********             ********      ********   *********     *******  **********
'   *******   **********    *******   ********             ********       ********    ********     *******   ********
'   *******    *********    *******    ******              ********       ********     *******     *******    ******
'   *******     ********    *******     ****              *********       ********     *******     *******     ****
'   *******      *******    *******      **               ********        ********     *******     *******      **
'   *******      *******    *******                      *********        ********     *******     *******
'   *******      *******    *******                      *********        ********     *******     *******
'   *******      *******    *******                      ********         ********     *******     *******
'   *******      *******    *******                     *********         ********     *******     *******
'   *******      *******    *******                     ********          ********     *******     *******
'   *******      *******    *******                    *********          ********     *******     *******
'   *******      *******    *******                    ***************    ********     *******     *******
'   ********************    *******                   ***************     ********     *******     *******
'   ********************    *******                   **************      ********     *******     *******
'   ********************    *******                   **************      ********     *******     *******
'   ********************    *******                  **************       ********     *******     *******
'   ********************    *******                  *************        ********     *******     *******
'   ********************    *******                 *************         ********     *******     *******
'   ********************    *******                       *******         ********     *******     *******
'   ********************    *******       *               ******          ********     *******     *******      **
'   *******      *******    *******      ***             ******           ********     *******     *******     ****
'   *******      *******    *******     *****            ******           ********     *******     *******    ******
'   *******      *******    *******    *******           *****            ********     *******     *******   ********
'   *******      *******    ********  *********         *****             ********    ********     *******  **********
'   *******      *******    *******************         ****              ********   *********     *******************
'  ********     ********    ******************          ****              ********  **********     ******************
' **********   **********    ****************          ****               ********************      ****************
'************ ************    **************           ***               ********************        **************
' *********** ************     ************            ***              ********************          ************
'  *********   **********       **********            ***               *******************            **********
'   *******     ********         ********             **                 *****************              ********
'    *****       ******           ******              *                   ***************                ******
'     ***         ****             ****              **                    *************                  ****
'      *           **               **               *                      ***********                    **
'                                                   *
'VPX recreation by ninuzzu
'thanks to javier for let me continue this table, wich is based on his PRO version
'Big thanks to DJRobX: without him, this table wouldn't exist.
'I also would like to thank:
'dark for the bell base model;
'zany for the flasher domes;
'tom tower for the apron;
'knorr for the laser mod
'rysr, Peter J and javier (again :D ) for the help with the pics.
'rom and zedonius for some resources I used from their FP table;
'the VPDevTeam for the freaking amazing VPX

'Surround Sound MOD by RustyCardores
' This mod includeds a shaker pulse and if ned be it can easily be reomoved in the SOLENOIDS section of this script.

' Thalamus 2018-07-24
' No special SSF tweaks yet.
' Added InitVpmFFlipsSAM


Option explicit
Randomize

'************************************************************************
'             TABLE OPTIONS
'************************************************************************

Const ACDCLuci = 0          'Show Luci Mini Playfield (0= no, 1= yes)
Const InstrCardType = 1       'Instruction Cards Type (0= white, 1= yellow, 2 = custom)
Const TrainHornsLit = 1       'Light the Train Horns (0= no, 1= yes)
Const LaserMod = 0          'Shows a laser beam when Cannon is moving (0= no, 1= yes)
Const BellMod = 0         'Shows a Wooden Bell Log instead of the Metallic one (0= no, 1= yes)
Const RampsDecals = 1       'Shows Custom Ramps Decals (0= no, 1= yes)
Const FlippersDecals = 1      'Shows Custom Flippers Decals (0= no, 1= yes)
Const SideCabDecals = 1       'Shows Custom Cabinet Decals (0= no, 1= yes)
Const FIREButtonLight = 1     'Light the apron decal when FIRE button is on, useful for pincabs (0= no, 1= yes)

'************************************************************************
'           END OF TABLE OPTIONS
'************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const Ballsize = 51
Const BallMass = 1.1

Const UseVPMModSol = True
Dim UseVPMDMD,CustomDMD,DesktopMode
DesktopMode = ACDC.ShowDT : CustomDMD = False
If Right(cGamename,1)="c" Then CustomDMD=True
If CustomDMD OR (NOT DesktopMode AND NOT CustomDMD) Then UseVPMDMD = False    'hides the internal VPMDMD when using the color ROM or when table is in Full Screen and color ROM is not in use
If DesktopMode AND NOT CustomDMD Then UseVPMDMD = True              'shows the internal VPMDMD when in desktop mode and color ROM is not in use
Scoretext.visible = NOT CustomDMD                       'hides the textbox when using the color ROM

LoadVPM "02800000", "Sam.VBS", 3.54

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "fx_Solon"
Const SSolenoidOff = ""
Const SCoin = "fx_coin"

'************************************************************************
'            INIT TABLE
'************************************************************************

'Const cGameName = "acd_168h"
Const cGameName = "acd_168hc" 'Colour DMD

Dim bsTrough, bsTEject, bsBellEject, bsLowEject, PlungerIM, ACDCBank, TNTBank

Sub ACDC_Init
  vpmInit Me
    With Controller
        .GameName = cGameName
        .SplashInfoLine = "AC/DC Premium (Stern 2012)"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        If NOT CustomDMD Then .Hidden = DesktopMode       'hides the external DMD when in desktop mode and color ROM is not in use
        .HandleMechanics = 0
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
  InitVpmFFlipsSAM
    End With

'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

'Nudging
    vpmNudge.TiltSwitch = -7
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

'Trough
    Set bsTrough = New cvpmTrough
    With bsTrough
    .Size = 4
    .InitSwitches Array(21, 20, 19, 18)
    .InitExit BallRelease, 70, 15
    .InitEntrySounds "fx_drain", SoundFX(SSolenoidOn,DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
    .InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_ballrel",DOFContactors)
    .Balls = 4
    .CreateEvents "bsTrough", Drain
    End With

'Top Eject
    Set bsTEject = new cvpmSaucer
    With bsTEject
    .InitKicker Sw37, 37, 90, 20, 0
    .InitExitVariance 0, 1
        .InitSounds "fx_kicker_enter", SoundFX(SSolenoidOn,DOFContactors), SoundFX("ExitSandman",DOFContactors)
        .CreateEvents "bsTEject", Sw37
    End With

'Bell Eject
    Set bsBellEject = new cvpmSaucer
    With bsBellEject
        .InitKicker Sw36, 36, 190, 10, 0
    .InitExitVariance 0, 2
        .InitSounds "fx_kicker_enter", SoundFX(SSolenoidOn,DOFContactors), SoundFX("ExitSandman",DOFContactors)
        .CreateEvents "bsBellEject", Sw36
    End With

'Lower PF Eject
    Set bsLowEject = new cvpmTrough
    With bsLowEject
    .Size = 1
    .InitSwitches Array(49)
    .InitExit Sw49, 32, 35
    .InitExitVariance 0, 2
    .InitEntrySounds "fx_kin", "", ""
    .InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("Popper",DOFContactors)
    .Balls = 1
    .CreateEvents "bsLowEject", Sw49
    End With

'Impulse Plunger
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP sw23, 65, 0.5
        .Switch 23
        .Random 1.5
        .InitExitSnd SoundFX("fx_AutoPlunger",DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
        .CreateEvents "plungerIM"
    End With

'5-bank Drop Target
  Set ACDCBank = New cvpmDropTarget
  With ACDCBank
    .InitDrop Array(sw1,sw2,sw3,sw4,sw5), Array(1,2,3,4,5)
    .Initsnd SoundFX("fx_DropTargetDownL",DOFDropTargets), SoundFX("fx_DropTargetUpL",DOFDropTargets)
  End With

'3-bank Drop Target
  Set TNTBank = New cvpmDropTarget
  With TNTBank
    .InitDrop Array(sw10,sw11,sw12), Array(10,11,12)
    .Initsnd SoundFX("fx_DropTargetDownTNT",DOFDropTargets), SoundFX("fx_DropTargetUpTNT",DOFDropTargets)
  End With

'Other Suff
  InitOptions:InitBell:InitCannon:InitDiverters
End Sub

'************************************************************************
'             KEYS
'************************************************************************

Sub ACDC_KeyDown(ByVal Keycode)
  If keycode = LeftFlipperKey  Then Controller.Switch(88) = 1     'Mini PF flipper switch
  If keycode = RightFlipperKey  Then Controller.Switch(86) = 1    'Mini PF flipper switch
    If keycode = LeftTiltKey Then Nudge 90, 3:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 3:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 4:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
  If KeyCode = PlungerKey Then Plunger.Pullback: PlaysoundAt "fx_PlungerPull",Plunger
  If KeyCode = RightMagnaSave OR KeyCode = LockbarKey Then Controller.Switch(64)=1        'FIRE Button, mapped to the right magna save
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub ACDC_KeyUp(ByVal Keycode)
  If keycode = LeftFlipperKey  Then Controller.Switch(88) = 0     'Mini PF flipper switch
  If keycode = RightFlipperKey  Then Controller.Switch(86) = 0    'Mini PF flipper switch
  If KeyCode = PlungerKey Then Plunger.Fire: StopSound "fx_PlungerPull":PlaySoundAt "fx_Plunger",Plunger
  If KeyCode = RightMagnaSave OR KeyCode = LockbarKey Then Controller.Switch(64)=0        'FIRE Button, mapped to the right magna save
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub ACDC_Paused:Controller.Pause = 1:End Sub
Sub ACDC_UnPaused:Controller.Pause = 0:End Sub
Sub ACDC_Exit:Controller.Stop:End Sub

'************************************************************************
'            SOLENOIDS
'************************************************************************

SolCallBack(1) = "SolTrough"            'Trough-Up Kicker
SolCallBack(2) = "SolAutoPlungerIM"         'AutoLaunch
SolCallback(3) = "bsLowEject.SolOut"        'Lower PF Eject
SolCallback(4) = "SolLFlipperMini"          'Lower PF Left Flipper
SolCallback(5) = "SolRFlipperMini"          'Lower PF Right Flipper
SolCallback(6) = "ACDCBank.SolDropUp"       '5-bank Drop Target
SolCallback(7) = "TNTBank.SolDropUp"        '3-bank Drop Target
SolCallback(8) = "vpmSolSound SoundFX(""ShakerPulse"",DOFShaker),"  'Shaker Motor
'SolCallback(8) = "vpmSolSound SoundFX("""",DOFShaker),"  'Shaker Motor - Swap this line with above to remove shaker audio.
'SolCallback(9) = ""                'Left Bumper
'SolCallback(10)= ""                'Right Bumper
'SolCallback(11)= ""                'Top Bumper
SolCallback(12)= "SolTopEject"            'Top Eject
'SolCallback(13)= ""                'Left Sling
'SolCallback(14)= ""                'Right Sling
SolCallback(15)= "SolLFlipper"            'Left Flipper
SolCallback(16)= "SolRFlipper"            'Right Flipper
SolModCallBack(17)= "SetLampMod 177,"       'Flasher: Train
SolCallback(18)= "SolDetonator"           'Detonator
SolModCallBack(19)= "SetLampMod 179,"       'Flasher:Bottom Arch (x2)
SolModCallBack(20)= "SetLampMod 180,"       'Flasher: Left Ramp
SolModCallBack(21)= "SetLampMod 181,"       'Flasher:Left Side
SolModCallBack(22)= "SetLampMod 182,"       'Flasher:BackPanel
SolModCallBack(23)= "SetLampMod 183,"       'Flasher: Top Eject
'SolCallBack(24)= "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"         'Knocker - Uncomment for Knocker on coin
SolModCallBack(25)= "SetLampMod 185,"       'Flasher:Pop Bumpers (x3)
SolModCallBack(26)= "SetLampMod 186,"       'Flasher: Bell Arrow
SolModCallBack(27)= "SetLampMod 187,"       'Flasher:Left Ramp Left
SolModCallBack(28)= "SetLampMod 188,"       'Flasher:Left Ramp Right
SolModCallBack(29)= "SetLampMod 189,"       'Flasher:Right Ramp Right
SolModCallBack(30)= "SetLampMod 190,"       'Flasher:Right Ramp
SolModCallBack(31)= "SetLampMod 191,"       'Flasher:Right Side
SolCallback(32)= "SolCannonMotor"         'Cannon Motor

'aux coils 41-48
SolCallBack(51)  = "SolBandMembers"         'Band Members Mech
SolCallBack(52)  = "bsBellEject.SolOut"       'Bell Eject
SolCallBack(53)  = "SolCannon"            'Cannon Eject
SolCallBack(54)  = "SolBellMove"          'Bell Magnet
SolCallBack(55)  = "SolCannonDiv"         'Cannon Diverter
SolCallBack(56)  = "vpmSolGate Gate, SoundFX(SSolenoidOn,DOFContactors)," 'Right Control Gate
SolCallBack(57)  = "SolRampDiv"           'Left Ramp Diverter
'SolCallBack(58)  = ""                'Cabinet Topper (???)

'************************************************************************
'               FLIPPERS
'************************************************************************
Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers),LeftFlipper
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers),LeftFlipper
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers),RightFlipper
        RightFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers),RightFlipper
        RightFlipper.RotateToStart
    End If
End Sub

Sub SolLFlipperMini(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers),LeftFlipperMini
        LeftFlipperMini.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers),LeftFlipperMini
        LeftFlipperMini.RotateToStart
    End If
End Sub

Sub SolRFlipperMini(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers),RightFlipperMini
        RightFlipperMini.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers),RightFlipperMini
        RightFlipperMini.RotateToStart
    End If
End Sub

'************************************************************************
'            BALL TROUGH
'************************************************************************
Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    If BsTrough.Balls Then vpmTimer.PulseSw 22
  End If
End Sub

'************************************************************************
'            AUTOPLUNGER
'************************************************************************
Sub SolAutoPlungerIM(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
End Sub

'************************************************************************
'            EXIT JUKEBOX
'************************************************************************
Sub SolTopEject(enabled)
  If Enabled Then
    bsTEject.ExitSol_On
    Gate.elasticity = RndNum(6,10)/10     ' Sets a random elasticity between 0.6 and 1
  End If
End Sub

'************************************************************************
'             RAMP DIVERTERS
'************************************************************************
Dim Div1Hit,Div3Hit

Sub InitDiverters
  Div1.Isdropped = 1:Div3.Isdropped = 1
  Div1Hit=0: Div3Hit=0
End Sub

Sub SolCannonDiv(Enabled)
    If Enabled Then
        Div1.Isdropped = 0
        Div2.Isdropped = 1
        DivP.Z = 142
        PlaySoundAt SoundFX("DiverterOn", DOFContactors),light41
    Else
        Div1.Isdropped = 1
        Div2.Isdropped = 0
        DivP.Z = 172.5
        PlaySoundAt SoundFX("DiverterOff", DOFContactors),light41
    End If
End Sub

Sub SolRampDiv(Enabled)
    If Enabled Then
        Div3.Isdropped = 0
        Div4.Isdropped = 1
        Div1P.Z = 150
        PlaySoundAt SoundFX("DiverterOn", DOFContactors),light5
    Else
        Div3.Isdropped = 1
        Div4.Isdropped = 0
        Div1P.Z = 180.5
        PlaySoundAt SoundFX("DiverterOff", DOFContactors),light5
    End If
End Sub

'************************************************************************
'             CANNON ENTER
'************************************************************************
Const Pi=3.1415926535
Dim GPos, GDir, BallInGun, BallInGunRadius, LaserON
BallInGunRadius = SQR((Cannon_assyM.X - Sw45.X)^2 + (Cannon_assyM.Y - Sw45.Y)^2)
GPos = 110: GDir = -1

Sub Sw45_hit()
  RandomSoundMetal
  Controller.switch(45) = 1
  Set BallInGun = ActiveBall
End Sub

'************************************************************************
'             CANNON MOTOR
'************************************************************************
Sub InitCannon
    Controller.switch(61) = 1
    Controller.switch(62) = 0
End Sub

Sub SolCannonMotor(Enabled)
    If Enabled Then
    GDir = -1
    UpdateCannon.Enabled = 1
    PlaySoundAt SoundFX("CannonMotor", DOFGear),light41
    If LaserMod = 1 then
      LaserR.Size_Z = 0.5
      LaserR1.Size_Z = 0.5
      LaserR.visible = 1
      LaserR1.visible = 1
      LaserZTimer.Enabled = 1
      LaserON = 1
    End if
     Else
    UpdateCannon.Enabled = 0
    StopSound "CannonMotor"
    LaserON = 0
  End If
End Sub

 Sub UpdateCannon_Timer()
  GPos = GPos + GDir
  If GPos < 110 AND GDir=-1 Then Controller.switch(61) = 0
  If GPos <= 80 Then Controller.switch(62) = 1
  If GPos <= 20 Then GDir = 1
  If GPos > 80 AND GPos < 110 AND GDir=1 Then Controller.switch(61) = 0: Controller.switch(62) = 0
  If GPos >= 110 Then GPos=110:GDir = -1: Controller.switch(61) = 1: Controller.switch(62) = 0

  Cannon_assyM.RotZ = GPos
  Cannon_decalM.RotZ = GPos
  Cannon_assyP.RotZ = GPos
  Cannon_decalP.RotZ = GPos
  Cannon_Plunger.RotZ = GPos
  Cannon_shadow.RotZ = GPos
  LaserR.objRotZ = -(GPos-90)
  LaserR1.objRotZ = -(Gpos -90)
  If Not IsEmpty(BallInGun) Then
    BallInGun.X = Cannon_assyM.X - BallInGunRadius * Cos(GPos*Pi/180)
    BallInGun.Y = Cannon_assyM.Y - BallInGunRadius * Sin(GPos*Pi/180)
    BallInGun.Z = CannonRamp.HeightTop + Ballsize
  End If
 End Sub

'***LaserZTimer***

Sub LaserZTimer_Timer()
  If GPos >= 102 And GPos <= 110 then LaserR.Size_Z = 0.5: LaserR1.Size_Z = 0.5
  If GPos <= -101.99 And GPos >= 97 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1
  If GPos <= -96.99 And GPos >= 89.01 then LaserR.Size_Z = 0.55: LaserR1.Size_Z = 0.55
  If GPos <= 89 And GPos >= 81.01 then LaserR.Size_Z = 0.48: LaserR1.Size_Z = 0.48
  If GPos <= 81 And GPos >= 73.01 then LaserR.Size_Z = 0.78: LaserR1.Size_Z = 0.78
  If GPos <= 73 And GPos >= 65.01 then LaserR.Size_Z = 0.7: LaserR1.Size_Z = 0.7
  If GPos <= 65 And GPos >= 45.01 then LaserR.Size_Z = 0.6: LaserR1.Size_Z = 0.6
  If GPos <= 45 And GPos >= 42.01 then LaserR.Size_Z = 0.43: LaserR1.Size_Z = 0.43
  If GPos <= 42 And GPos >= 40.01 then LaserR.Size_Z = 0.425: LaserR1.Size_Z = 0.425
  If GPos <= 40 And GPos >= 37.01 then LaserR.Size_Z = 0.41: LaserR1.Size_Z = 0.41
  If GPos <= 37 And GPos >= 30.01 then LaserR.Size_Z = 0.39: LaserR1.Size_Z = 0.39
  If GPos <= 30 And GPos >= 22.01 then LaserR.Size_Z = 0.385: LaserR1.Size_Z = 0.385
  If GPos <= 22  then LaserR.Size_Z = 0.385: LaserR1.Size_Z = 0.375
  If GPos >= 110 And LaserON = 0 Then
    LaserR.Visible = 0
    LaserR1.Visible = 0
    LaserZTimer.Enabled = 0
  End if
End Sub

'************************************************************************
'             CANNON KICKOUT
'************************************************************************
 Sub SolCannon(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("CannoShot", DOFContactors),light41
    Cannon_Plunger.PlayAnim 0, 1
    If NOT IsEmpty(BallInGun) AND Gpos < 100 Then
      Sw45.kick GPos-90, 45
      Controller.switch(45) = 0
      BallInGun = Empty
    End If
  End If
End Sub

'************************************************************************
'           BAND MEMBERS MECH
'************************************************************************
Dim LegRotRadius,LegCase,LegShake
LegRotRadius = SQR((Figure_AngusLeg.X - Supp_Angus.X)^2 + (Figure_AngusLeg.Z - Supp_Angus.Z)^2)
LegShake = 0

Sub SolBandMembers(enabled)
  If Enabled Then
    Flipper1.TimerInterval = 10
    Flipper1.TimerEnabled = 1
    Flipper1.RotateToEnd
    Flipper2.RotateToEnd
    LegShake = 0: PlaySoundAtVol SoundFX("fx_solon", DOFContactors),light31,2
  End If
End Sub

Sub Flipper1_timer
  SolPlung.TransX = Flipper2.currentangle
  Supp_Angus.RotY = Flipper1.currentangle
  Supp_Malcolm.RotY = Flipper1.currentangle
  Supp_Brian.RotY = Flipper1.currentangle
  Supp_Cliff.RotY = Flipper1.currentangle
  Figure_Angus.RotY = Flipper1.currentangle
  Figure_AngusLeg.X = Supp_Angus.X + LegRotRadius * Sin((Supp_Angus.RotY)*Pi/180)
  Figure_AngusLeg.Z = Supp_Angus.Z + LegRotRadius * Cos((Supp_Angus.RotY)*Pi/180)
  Figure_Malcolm.RotY = Flipper1.currentangle
  Figure_Brian.RotY = Flipper1.currentangle
  Figure_Cliff.RotY = Flipper1.currentangle
  Figure_DrumL.RotY = Flipper2.currentangle
  Figure_DrumR.RotY = -Flipper2.currentangle
  If Supp_Angus.RotY > 10 Then LegCase = 1: Flipper2.TimerInterval= 10: Flipper2.TimerEnabled= 1
  If Supp_Angus.RotY = 35 Then Flipper1.RotateToStart: Flipper2.RotateToStart
  If Supp_Angus.RotY = 0 Then Me.TimerEnabled = 0
End Sub

Sub Flipper2_timer
  Select Case LegCase
    Case 1
      Figure_AngusLeg.RotY = Figure_AngusLeg.RotY - 10
      If Figure_AngusLeg.RotY < -50 Then LegCase = 2
    Case 2
      LegShake = LegShake + 0.05
      Figure_AngusLeg.RotY = Figure_AngusLeg.RotY + 3 * Sin(LegShake)
      If LegShake > 1.25 Then Me.TimerEnabled = 0: Figure_AngusLeg.RotY = 0: LegShake = 0
  End Select
End Sub

'************************************************************************
'             T.N.T. DETONATOR
'************************************************************************
Dim detDir

Sub SolDetonator(enabled)
If Enabled Then
  PlaySoundAt SoundFX(SSolenoidOn,DOFContactors),light7
  DetonatorTimer.Enabled=1
  detDir=1
End If
End Sub

Sub DetonatorTimer_Timer
Detonator.Z= Detonator.Z - 2 * detDir
If Detonator.Z <= 120 AND detDir=1 Then detDir=-1
If Detonator.Z >= 150 AND detDir=-1 Then Me.Enabled= 0
End Sub

'************************************************************************
'           BELL SWINGING ANIMATION
'************************************************************************
Dim NewtonBall,BellHit,BellExit,ccc,rotat,brakew
Const NewtonBallMass = 0.7

Sub InitBell
  BellMove.Enabled=1:BellHit=False:BellExit=False
  Set NewtonBall = Bellk.CreateSizedBallWithMass (Ballsize/2, NewtonBallMass)
  NewtonBall.visible=false
End Sub

Sub BellMove_Timer
  Bell.RotX = Spinner5.currentangle:Bell_Support.RotX = Bell.RotX: Bell_Log.RotX = Bell.RotX
  NewtonBall.X = Bell.X + (Bell.Z-50) * sin (Bell.RotX*Pi/180) * sin (5 * Pi/180)
  NewtonBall.Y = Bell.Y + (Bell.Z-50) * sin (Bell.RotX*Pi/180) * cos (5 * Pi/180)
  NewtonBall.Z = Bell.Z - (Bell.Z-50) * cos (Bell.RotX*Pi/180)
End Sub

Sub Sw47_Spin():vpmTimer.PulseSw 47:End Sub

Sub SolBellMove(Enabled)
If Enabled then
  BellMove.Enabled=0
  rotat=0:brakew=1:ccc=0
  BellK.TimerEnabled=0
  BellK.TimerEnabled=1
  BellK.TimerInterval=10
End If
End Sub

Sub Sw36_UnHit                    'fix after Hells Bells lower pf
  If Bell.rotX<1 And Bell.rotX>-1 And Not BellHit Then
    SolBellMove True : BellExit = True
  End If
  Me.TimerInterval=1000:Me.TimerEnabled=1
End Sub

Sub sw36_timer:Me.TimerEnabled=0:BellExit=False:End Sub

Sub CheckBell_Hit:BellHit=True:Me.TimerInterval=2000:Me.TimerEnabled=0:Me.TimerEnabled=1:End Sub
Sub CheckBell_timer:Me.TimerEnabled=0:BellHit=False:End Sub

Sub BellK_timer
  ccc=ccc+1
  If brakew <7 then rotat=rotat+0.085:brakew=brakew+0.005
  Bell.rotX=-(cos(rotat+90)*100)/brakew^3:Bell_Support.RotX = Bell.RotX: Bell_Log.RotX = Bell.RotX
  NewtonBall.X = Bell.X + (Bell.Z-50) * sin (Bell.RotX*Pi/180) * sin (5 * Pi/180)
  NewtonBall.Y = Bell.Y + (Bell.Z-50) * sin (Bell.RotX*Pi/180) * cos (5 * Pi/180)
  NewtonBall.Z = Bell.Z - (Bell.Z-50) * cos (Bell.RotX*Pi/180)
  If (Bell.RotX < 16 AND Bell.RotX > 8) OR (Bell.RotX < -8 AND Bell.RotX > -16) Then Controller.Switch(47) = 1 : Else Controller.Switch(47) = 0 : End If
  If (Bell.RotX < 1 AND Bell.RotX>-1 AND ccc>500) OR (BellHit AND NOT BellExit) Then Me.TimerEnabled=0:ccc=0:BellMove.Enabled=1
End Sub

'************************************************************************
'           SWITCHES
'************************************************************************

'Targets AC/DC
Sub sw1_dropped:ACDCBank.Hit 1:End Sub
Sub sw2_dropped:ACDCBank.Hit 2:End Sub
Sub sw3_dropped:ACDCBank.Hit 3:End Sub
Sub sw4_dropped:ACDCBank.Hit 4:End Sub
Sub sw5_dropped:ACDCBank.Hit 5:End Sub

'Targets ROCK
Sub Sw6_Hit():vpmTimer.PulseSw 6: PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub
Sub Sw7_Hit():vpmTimer.PulseSw 7: PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub
Sub Sw8_Hit():vpmTimer.PulseSw 8: PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub
Sub Sw9_Hit():vpmTimer.PulseSw 9: PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub

'Targets TNT
Sub sw10_dropped:TNTBank.Hit 1:End Sub
Sub sw11_dropped:TNTBank.Hit 2:End Sub
Sub sw12_dropped:TNTBank.Hit 3:End Sub

'Left Ramp
Sub Sw13_Hit:Controller.Switch(13)=1: End Sub
Sub Sw13_UnHit:Controller.Switch(13)=0: End Sub

Sub Sw14_Hit:Controller.Switch(14) = 1:End Sub
Sub Sw14_UnHit():Controller.Switch(14) = 0:End Sub

'Bottom Lane Rollovers
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "rollover",Sw24:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "rollover",Sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "rollover",Sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "rollover",Sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

' Slingshots
Dim LStep, RStep

Sub LeftSlingShot_Slingshot()
    PlaySoundAt SoundFX("LeftSlingShot", DOFContactors),lemk
    LeftSling4.Visible = 1
    Lemk.TransZ = -20
    LStep = 0
    vpmTimer.PulseSw 26
    Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer()
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.TransZ = -12
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.TransZ = -5
        Case 3:LeftSLing2.Visible = 0:Lemk.TransZ = 0:Me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot()
    PlaySoundAt SoundFX("RightSlingShot", DOFContactors),remk
    RightSling4.Visible = 1
    Remk.TransZ = -20
    RStep = 0
    vpmTimer.PulseSw 27
    Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer()
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.TransZ = -12
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.TransZ = -5
        Case 3:RightSLing2.Visible = 0:Remk.TransZ = 0:Me.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit():vpmTimer.PulseSw 30:PlaySoundAtBumperVol SoundFX("LeftBumper_Hit", DOFContactors),Bumper1,2:End Sub
Sub Bumper2_Hit():vpmTimer.PulseSw 31:PlaySoundAtBumperVol SoundFX("RightBumper_Hit", DOFContactors),Bumper2,2:End Sub
Sub Bumper3_Hit():vpmTimer.PulseSw 32:PlaySoundAtBumperVol SoundFX("TopBumper_Hit", DOFContactors),Bumper3,2:End Sub

' Spinner
Sub Sw33_Spin():vpmTimer.PulseSw 33:PlaySoundAt "fx_spinner",sw34 : End Sub

'Targets Thunder
Sub Sw34_Hit():vpmTimer.PulseSw 34:PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub
Sub Sw35_Hit():vpmTimer.PulseSw 35:PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub
Sub Sw42_Hit():vpmTimer.PulseSw 42:PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub

'Top Lane Rollovers
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "rollover",Sw38:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAt "rollover",Sw39:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:PlaySoundAt "rollover",Sw40:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub

'Right Ramp
Sub Sw41_Hit:Controller.Switch(41) = 1:End Sub
Sub Sw41_UnHit:Controller.Switch(41) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43)=1: End Sub
Sub sw43_UnHit:Controller.Switch(43)=0: End Sub

'Left Orbit
Sub sw44_Hit:Controller.Switch(44) = 1:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:If Activeball.velX>20 Then Activeball.velX = Activeball.VelX*0.8:End If:End Sub

'Plunger Lane
Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "rollover",Sw48:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

'Right Orbit
Sub sw59_Hit:Controller.Switch(59) = 1:End Sub
Sub sw59_UnHit:Controller.Switch(59) = 0:End Sub

'Detonator Target
Sub sw46_hit():vpmTimer.PulseSw 46: PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub

'Lower PF target left
Sub sw50_hit():vpmTimer.PulseSw 50: PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub

'Lower PF target center
Sub sw51_hit():vpmTimer.PulseSw 51: PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub

'Lower PF target right
Sub sw52_hit():vpmTimer.PulseSw 52: PlaySoundAtBallVol SoundFX("fx_target",DOFTargets),2: End Sub

'Lower PF rollover left
Sub sw53_hit:Controller.Switch(53) = 1:PlaySoundAt "rollover",sw53:End Sub
Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

'Lower PF rollover right
Sub sw54_hit:Controller.Switch(54) = 1:PlaySoundAt "rollover",sw54:End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
'     With Modulated Flasher Routines (ninuzzu)
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim bulb
Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
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
  UpdateGIs
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
'LED Inserts
    nFadeL 1, l41
    nFadeL 2, l42
    nFadeL 3, l43
    nFadeL 4, l44
    nFadeL 5, l45
    nFadeL 6, l46
    nFadeL 7, l47
    nFadeL 8, l48
    nFadeL 9, l51
    nFadeL 10, l50
    nFadeL 11, l49
    nFadeL 12, l3
    nFadeL 13, l4
    nFadeL 14, l5

    nFadeL 17, l9
    nFadeL 18, l10
    nFadeL 19, l11
    nFadeL 20, l12
    nFadeL 21, l13
    nFadeL 22, l19
    nFadeL 23, l21
    nFadeL 24, l25
    nFadeL 25, l1

  NFadeObjm 32, TrainFrontLight,"bulbBlueOn", "bulbBlue"
    Flash 32, f17c
    nFadeL 33, l33
    nFadeL 34, l32
    nFadeL 35, l31
    nFadeL 36, l30
    nFadeL 37, l2
    nFadeL 38, l34

    nFadeL 40, l27
    nFadeL 41, l64
    nFadeL 42, l6
    nFadeL 43, l7
    nFadeL 44, l8

    nFadeL 49, l22
    nFadeL 50, l23
    nFadeL 51, l24

'hornet left
  NFadeObjm 53, l57p, "Horns_on", "Horns_off"
  Flashm 53, f17a   'Train Horn
  Flashm 53, l57
  Flash 53, l57a
'hornet right
  NFadeObjm 54, l58p, "Horns_on", "Horns_off"
  Flashm 54, f17b   'Train Horn
  Flashm 54, l58
  Flash 54, l58a
'Bumpers
    nFadeLm 62, l60
    nFadeLm 62, l60a
    nFadeL 62, l60b
    nFadeLm 63, l61
    nFadeLm 63, l61a
    nFadeL 63, l61b
    nFadeLm 64, l62
    nFadeLm 64, l62a
    nFadeL 64, l62b

'Tracks
    Flash 65, T_YouShookMe        'YouShookMeAllNightLong
    Flash 66, T_HighwaytoHell     'HighwayToHell
    Flash 67, T_RockNRollTrain      'RockNRollTrain
    Flash 68, T_WholeLottaRosie     'WholeLottaRosie
    Flash 69, T_HellsBells        'HellsBells
    Flash 70, T_ThunderStruck     'ThunderStruck
    Flash 76, T_BackInBlack       'BackInBlack
    Flash 75, T_WarMachine        'WarMachine
    Flash 74, T_ForThoseAbout     'ForThoseAboutToRock
    Flash 73, T_TNT           'TNT
    Flash 72, T_HellAintBad       'HellAintBadPlaceToBe
    Flash 71, T_LetThereBeRock      'LetThereBeRock

'LED Flames Tunnel
    LampMod 151, l151
    LampMod 152, l152
    LampMod 153, l153
    LampMod 154, l154
    LampMod 155, l155
    LampMod 156, l156
    LampMod 157, l157
    LampMod 158, l158

'FIRE button
l63.colorfull = RGB(255*LampState(59),255*LampState(60),255*LampState(61))
l63a.colorfull = RGB(255*LampState(59),255*LampState(60),255*LampState(61))
If LampState(59) OR LampState(60) OR LampState(61) Then l63.state=1:l63a.state=1:Else l63.state=0:l63a.state=0:End If

'LED RGB Inserts
'VPM returns an 0-100 range value
  RGBLED l17, LampState(83), LampState(82), LampState(81)     'face mouth
  RGBLED l35, LampState(86), LampState(85), LampState(84)     'bell arrow (top)
  RGBLED l38, LampState(89), LampState(88), LampState(87)     'center top lane
  RGBLED l40, LampState(92), LampState(91), LampState(90)     'tunes n stuff
  RGBLED l40a,LampState(92), LampState(91), LampState(90)     'tunes n stuff (metal reflection)
  RGBLED l28, LampState(95), LampState(94), LampState(93)     'right loop arrow (mid)
  RGBLED l36, LampState(98), LampState(97), LampState(96)     'bell arrow (bot)
  RGBLED l37, LampState(101), LampState(100), LampState(99)   'left top lane
  RGBLED l26, LampState(104), LampState(103), LampState(102)    'right ramp arrow
  RGBLED l18, LampState(107), LampState(106), LampState(105)    'left loop arrow (bot)
  RGBLED l29, LampState(110), LampState(109), LampState(108)    'right loop arrow (bot)
  RGBLED l52, LampState(113), LampState(112), LampState(111)    'left loop arrow (top)
  RGBLED l15, LampState(119), LampState(118), LampState(117)    'face right eye
  RGBLED l14, LampState(122), LampState(121), LampState(120)    'face left eye
  RGBLED l20, LampState(125), LampState(124), LampState(123)    'left ramp arrow
  RGBLED l39, LampState(128), LampState(127), LampState(126)    'right top lane
  RGBLED l39a,LampState(128), LampState(127), LampState(126)    'right top lane (metal reflection)

'Flashers
'VPM returns an 0-255 range value
  LampMod 177, f17    'Train
  LampMod 177, f17r   'Train
  LampMod 179, f19    'Bottom Arch
  LampMod 179, f19a   'Bottom Arch
  LampMod 180, f20    'Left Ramp
  LampMod 181, f21    'Yellow Dome
  LampMod 181, f21a   'Yellow Dome
  LampMod 181, f21c   'Yellow Dome
  LampMod 181, f21b   'Yellow Dome
  LampMod 181, f21bDT   'Yellow Dome
  LampMod 181, f21r   'Yellow Dome
  LampMod 182, f22    'Backbox
  LampMod 182, f22r     'Backbox
  LampMod 183, f23    'Top Eject
  LampMod 183, f23a     'Top Eject
  LampMod 185, f25a     'Bumper Flash
  LampMod 185, f25    'Bumper Flash
  LampMod 186, f26    'Bell Arrow
  LampMod 187, f27    'Thunder Left
  LampMod 187, f27a   'Thunder Left
  LampMod 187, f27b   'Thunder Left
  LampMod 187, f27c   'Thunder Left
  LampMod 188, f28    'Thunder Center
  LampMod 188, f28a   'Thunder Center
  LampMod 188, f28b   'Thunder Center
  LampMod 188, f28c   'Thunder Center
  LampMod 189, f29    'Thunder Right
  LampMod 189, f29a   'Thunder Right
  LampMod 189, f29b   'Thunder Right
  LampMod 189, f29c   'Thunder Right
  LampMod 190, f30    'Right Ramp
  LampMod 191, f31    'Red Dome
  LampMod 191, f31a   'Red Dome
  LampMod 191, f31b   'Red Dome
  LampMod 191, f31c   'Red Dome
  LampMod 191, f31bDT   'Red Dome
  LampMod 191, f31bDT1  'Red Dome
  LampMod 191, f31d   'Red Dome
  LampMod 191, f31r   'Red Dome
End Sub

Sub UpdateGIs
  'VPM returns an 0-50 range value
  '130 gi red
  For each bulb in GI_Red: bulb.Intensity = LampState (130)*1.5: Next
  For Each bulb in BLRLights: bulb.IntensityScale=LampState (130)/50:Next
  '132 gi blue
  For each bulb in GI_Blue: bulb.Intensity = LampState (132)*1.5: Next
  For Each bulb in BLBLights: bulb.IntensityScale=LampState (132)/50:Next
  '134 gi low pf
  light16.intensityScale = LampState (134)/50
  light19.intensityScale = LampState (134)/50
  light44.intensityScale = LampState (134)/50
  light45.intensityScale = LampState (134)/50
  For each bulb in GI_LowPF: bulb.Intensity = LampState (134)*3: Next
  '136 gi white
  For each bulb in GI_White: bulb.IntensityScale = LampState (136)/50: Next
  'change colorgrade when gi is off
  If Lampstate (130) = 0 AND Lampstate (132) = 0 AND Lampstate (136) = 0 Then
    ACDC.ColorGradeImage="ColorGrade_off":For Each bulb in BandMembersPoster: bulb.IntensityScale=0:Next
  ElseIf LampState (130) > 0 AND LampState (132) = 0 AND Lampstate (134) > 0  AND LampState (136) = 0 Then ACDC.ColorGradeImage="ColorGrade_red"
  ElseIf LampState (130) = 0 AND LampState (132) > 0 AND Lampstate (134) = 0  AND LampState (136) = 0 Then ACDC.ColorGradeImage="ColorGrade_blue"
  Else
    ACDC.ColorGradeImage="ColorGrade_on":For Each bulb in BandMembersPoster: bulb.IntensityScale=1:Next
  End If
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

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table
    Dim tmp
    tmp = ball.x * 2 / ACDC.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / ACDC.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function RndNum(min,max)
  RndNum = Int(Rnd()*(max-min+1))+min     ' Sets a random number between min AND max
End Function

' *********************************************************************
'           Ball Drop & Ramp Sounds
' *********************************************************************

'Sub ShooterStart_Hit():StopSound "fx_launchball":If ActiveBall.VelY < 0 Then PlaySoundAt "fx_launchball",Plunger:End If:End Sub  'ball is going up

Sub ShooterEnd_Hit:If ActiveBall.Z > 30  Then Me.TimerInterval=100:Me.TimerEnabled=1:End If:End Sub           'ball is flying
Sub ShooterEnd_Timer(): Me.TimerEnabled=0 : PlaySound "fx_balldrop",0,1,.2,0,0,0,0,-.8 : End Sub

Sub LREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAtBallVol "fx_ramp_enter1",.2:End If:End Sub     'ball is going up
Sub LREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_ramp_enter1":End If:End Sub   'ball is going down

Sub RREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAtBallVol "fx_ramp_enter1",.2:End If:End Sub     'ball is going up
Sub RREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_ramp_enter1":End If:End Sub   'ball is going down

Sub LREnter1_Hit():PlaySoundAtBallVol "fx_ramp_enter2",.2:End Sub
Sub RREnter1_Hit():PlaySoundAtBallVol "fx_ramp_enter2",.2:End Sub
Sub LREnter2_Hit():PlaySoundAtBallVol "fx_ramp_enter3",.2:End Sub
Sub RREnter2_Hit():PlaySoundAtBallVol "fx_ramp_enter3",.2:End Sub

Sub LRExit_Hit():ActiveBall.VelY=1:StopSound "fx_ramp_enter3":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub LRExit_timer():Me.TimerEnabled=0:PlaySound "fx_BallDrop2",0,3,-.4,0,0,0,0,.7:End Sub

Sub RRExit_Hit():StopSound "fx_ramp_enter3":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub RRExit_timer():Me.TimerEnabled=0:PlaySound "fx_BallDrop2",0,3,.4,0,0,0,0,.7:End Sub

Sub Div1_Hit() : StopSound "fx_ramp_enter3" : If Div1Hit=0 Then PlaySoundAtBall "fx_ramp_turn": End If: Div1Hit=1: Me.TimerInterval=3000: Me.TimerEnabled=1: End Sub
Sub Div1_Timer(): Me.TimerEnabled=0 : Div1Hit = 0 : End Sub

Sub Div3_Hit() : StopSound "fx_ramp_enter3" : If Div3Hit=0 Then PlaySoundAtBall "fx_ramp_turn" : End If: Div3Hit=1: Me.TimerInterval=3000: Me.TimerEnabled=1: End Sub
Sub Div3_Timer(): Me.TimerEnabled=0 : Div3Hit = 0 : End Sub

Sub CREnter_Hit() : PlaySoundAtBall "fx_ramp_metal" : End Sub

Sub CRExit_Hit() : StopSound "fx_ramp_metal" : PlaySoundAtBallVol "fx_ramp_enter3",.2: End Sub

Sub Helper_hit:ActiveBall.VelX=0.9*ActiveBall.VelX:ActiveBall.VelY=0.9*ActiveBall.VelY:End Sub

' *********************************************************************
'             Other Sound FX
' *********************************************************************

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 3
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 3
End Sub

Sub LeftFlipperMini_Collide(parm)
    PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 3
End Sub

Sub RightFlipperMini_Collide(parm)
    PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, 2
End Sub

Sub RandomSoundMetal()
    PlaySoundAtBallVol "fx_metal_hit_" & Int(Rnd*3)+1, 1
End Sub

Sub cRubbers_Hit(idx)
    PlaySoundAtBallVol "fx_rubber_hit_" & Int(Rnd*3)+1, 3
End Sub

Sub aGates_Hit (idx)
  PlaySoundAt "fx_gate4",ActiveBall
End Sub

Dim NextOrbitHit:NextOrbitHit = 0

'Sub WireRampBumps_Hit(idx)
' if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
'   RandomBump3 .5, Pitch(ActiveBall)+5
'   ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
'   ' Lowering these numbers allow more closely-spaced clunks.
'   NextOrbitHit = Timer + .2 + (Rnd * .2)
' end if
'End Sub

'Sub MetalGuideBumps_Hit(idx)
' if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
'   RandomBump2 2, Pitch(ActiveBall)
'   ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
'   ' Lowering these numbers allow more closely-spaced clunks.
'   NextOrbitHit = Timer + .2 + (Rnd * .2)
' end if
'End Sub

Sub PlasticRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 2, Pitch(ActiveBall)
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .1 + (Rnd * .2)
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
    PlaySound BumpSnd, 0, Vol(ActiveBall)+voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

'' Requires metalguidebump1 to 2 in Sound Manager
'Sub RandomBump2(voladj, freq)
' dim BumpSnd:BumpSnd= "metalguidebump" & CStr(Int(Rnd*2)+1)
'   PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
'End Sub
'
'' Requires WireRampBump1 to 5 in Sound Manager
'Sub RandomBump3(voladj, freq)
' dim BumpSnd:BumpSnd= "WireRampBump" & CStr(Int(Rnd*5)+1)
'   PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
'End Sub



' Stop Bump Sounds
Sub BumpSTOPplastic1_Hit ()
dim i:for i=1 to 7:StopSound "rampbump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOPplastic2_Hit ()
dim i:for i=1 to 7:StopSound "rampbump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOPplastic3_Hit ()
dim i:for i=1 to 7:StopSound "rampbump" & i:next
NextOrbitHit = Timer + 1
End Sub


' *********************************************************************
'             RealTime Updates
' *********************************************************************
Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
  RollingSoundUpdate
  BallShadowUpdate
  GateR2.RotX = - Spinner1.currentangle
  GateR.RotX = - Spinner2.currentangle
  GateL.RotX = - Spinner3.currentangle
  GateL2.RotX = - Spinner4.currentangle
  SpinnerT1.RotX = - Sw33.currentangle
  GateR1.RotX = - Spinner6.currentangle
  GateL1.RotX = - Spinner7.currentangle
  RightGate.RotY = Gate.currentangle
  LeftGate.RotX = -Gate2.currentangle
End Sub

'*********** ROLLING SOUND *********************************
Const tnob = 6            ' total number of balls : 4 (trough) + 1 (minipf trough) + 1 (Newton Ball)
Const fakeballs = 1         ' number of balls created on table start (rolling sound will be skipped)
ReDim rolling(tnob)
InitRolling

Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' stop the sound of deleted balls
  If UBound(BOT)<(tnob - 1) Then
    For b = (UBound(BOT) + 1) to (tnob-1)
      rolling(b) = False
      StopSound("fx_ballrolling" & b)
    Next
  End If
  ' exit the Sub if no balls on the table
    If UBound(BOT) = fakeballs-1 Then Exit Sub

'**********************************************************
' Raised Ramp RollingBall Sounds by RustyCardores & DJRobX
'**********************************************************

       ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 Then
      rolling(b) = True
      if BOT(b).z < 30 Then ' Ball on playfield
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.8, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
      Else ' Ball on raised ramp
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.4, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
    Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'**************************************************************************
'                 Positional Sound Playback Functions by DJRobX
'**************************************************************************

'Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(sound, tableobj)
    PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'Set all as per ball position & speed.

Sub PlaySoundAtBall(sound)
    PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub


'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub


'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub



'*********** BALL SHADOW *********************************
Dim BallShadow:BallShadow = Array (BallShadow1, BallShadow2, BallShadow3,Ballshadow4,Ballshadow5,Ballshadow6)

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
    If BOT(b).X < ACDC.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (ACDC.Width/2))/7)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (ACDC.Width/2))/7)) - 10
    End If
      ballShadow(b).Y = BOT(b).Y + 20
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

Sub InitOptions
  leftrail.visible = DesktopMode:rightrail.visible = DesktopMode
  f21b.visible = NOT DesktopMode : f21bDT.visible = DesktopMode
  f31b.visible = NOT DesktopMode : f31bDT.visible = DesktopMode : f31bDT1.visible = DesktopMode
  f17a.visible = TrainHornsLit
  f17b.visible = TrainHornsLit
  If ACDCLuci=0 Then
    Lower_PF.image = "low-pf"
  Else
    Lower_PF.image = "low-pf_Luci"
  End If
  If FlippersDecals=0 Then
    LeftFlipper.image = "LeftFlipper" : RightFlipper.image = "RightFlipper"
  Else
    LeftFlipper.image = "LeftFlipper_c" : RightFlipper.image = "RightFlipper_c"
  End If
  If SideCabDecals=0 Then
    LeftCab.image = "sidecabL":RightCab.image = "sidecabR"
  Else
    LeftCab.image = "sidecabL_c":RightCab.image = "sidecabR_c"
  End If
  Select Case InstrCardType
    Case 0
      CardLeft.image= "ACDC-CardLeft" : CardRight.image= "ACDC-CardRight"
    Case 1
      CardLeft.image= "ACDC-CardLeft_y" : CardRight.image= "ACDC-CardRight_y"
    Case 2
      CardLeft.image= "ACDCLuci-CardL" : CardRight.image= "ACDCLuci-CardR"
  End Select
  Bell_Log.visible = BellMod
  LeftRampDecal.visible = RampsDecals
  RightRampDecal.visible = RampsDecals
  CrossOverDecal.visible = RampsDecals
  l63a.visible = FIREButtonLight
End Sub


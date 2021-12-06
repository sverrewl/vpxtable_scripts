'  +              `s
' yNhssyyyydddhhyhhNo
'+/`      oMN      `+:                    ....`                  ```  ::  .//+/:
'         hMm        : `:/++.    ```.`    :mMd`    :ddo-    `/yyooooyhmm   +MM`  -sNNh/ `          `
'         hMN `/odmh:`   :NMy   .Nm:..     hMd`     yM/   `oNd.       -s+  /MM`    NM: `NhhyyhhssshN.
'         dMd    oMm`     hMM-  .N:`:hmh`  hMh      sM+  `dMm.          .  /MM`    NM. ss`  oMh    :+
'         NMd     dMh`   -MMMo `d+   mMs   hMs      sMo  yMM-           ```:MM+//++MM. `    yMy
'         mMd     `mMh   ds+MN`sy    mMy   hMs      oMo `NMm      -.:yNMN:`/MM/-.`.MM.      yMy
'         NMm      `hMy oh  NMdd     hMo   dMo      yM: `MMd         `NMd  /MM.   `MM`      hMs
'         MMd       `hMdN.  +Mm`     dMs   dMo     .ss+` dMN.         NMd  -Mm`   .MM`      mMo
'        `MMo        `dM:    d-      mMs   NMs   `.:/    .dMN/      `+Ns.  :MM`   -MM.      NM+
'        .MMy         `:            :MMs  :yyssyyhdm.      /ymmsooooo-    -dNd:` .+oo+/    `MM+`.
'        oMMy                       ..``                                                  ./+/:--
'      -+oo+/:-     ``
'                  -o.                  -
'                 /Nmdddddysysoooo+yNNNy.   `-+s+:-  .-//+/     -+sss:           +`
'                +s/:-.`         :dMMs.  `+ho:-:/yMNo`  /MMy      mM.   +hdysossyms
'               `              /dMm+`   oMy`      .dMm. `MMMm`    oN    .NM:      .:
'                           `oNMd/     yMh         `MMy `N+hMN:   +N    .MM-       `
'                         `oNMh:      `NM:          mMy  N+ oMMo  :M    `NMdhhhhdmmy
'                       .smMh-        -MM/         `NN.  m+  :NMh`-M`   `NM+``   `:
'                     :hMMy.           yMd-       `hy.   N+   .hMm:M.   `NM.
'                   :dMMs.          .+/ +dMdo:-:/oo-    `M/     oMMM-   `NM:       -
'                 :hMMd///++osyhhdddMy    .:++/-`       oNy+:.   -mM:   `MM- .-/+ym.
'                /o+/::--.`````     o                  -/-.`       s/  -ohysoo+/:/-

' Twilight Zone - IPDB No. 2684
' © Bally/Midway 1993
' VPX recreation by ninuzzu
' This table started as a FP conversion by coindropper. Clark Kent continued improving it at graphics and physics level
' Then he asked me an help with lighting and other things and it evolved into something different.
' We have completely rebuild the table, so I would call it a "from scratch" build, rather than a conversion.
' Flupper and Tom Tower joined the team later and enhanced the visuals with some stunning models. You guys rock!
' This table is maybe the most modded ever, lots of options in the script. Lots of fun!

' Credits/Thanks
' coindropper early development
' nFozzy for scripting the trough, the lock, the magnets and a lot of other things
' JPSalas for the help with the inserts and the script (you know who is THE MAN!)
' Tom Tower for the mini clock, pyramid, piano, camera, gumballmachine models, wich I retextured.
' Hauntfreaks for the Mystic Seer Toy, wich I retextured
' Flupper for retexturing and remeshing the plastic ramp
' Zany for the domes, flippers and bumpers models
' rom for "Robbie the Robot" model, wich I edited and retextured
' knorr and pacdude for some sound effects I borrowed from their tables
'

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' Wob 2018-08-09
' Added vpmInit Me to table init and cSingleLFlip
' DJRobX 2018-08-18
' Remove CSingleLFlip - there is an ULFlip sub
' Add more SSF and VPMModSol code.

Option Explicit
Randomize

Dim Romset, PowerballStart, CabinetSide, FlipperType, RampShadow, ScoopLight, GumballRGB, SlotMachineMod, PyramidMod, PyramidTopLight, MiniClockMod
Dim InvaderMod, MysticSeerMod, TargetMod, TVMod, LampMod, LampLightColor, TownSquarePostMod, SpiralMod, ExtraMagnet, BumperPostsMod

'******************************************************************************************
'* TABLE OPTIONS **************************************************************************
'******************************************************************************************

'***********  Choose the ROM    *********************************************************

Romset = 0      'arcade rom with credits,  1 = home rom (free play)

'***********  Set the Powerball Starting Location   ************************************

PowerBallStart = 7    '1-3 = gumball machine, 4-6 = trough, 0 = random, 7 = random gumball

'***********  Show cabinet sidewalls custom artwork ************************************

CabinetSide = 1     '0 = no , 1 = yes

'***********  Set the Flippers Type ***************************************************

FlipperType = 0     '0 = yellow , 1 = spiral , 2 = random

'***********  Render the Wireramp shadow on the playfield *******************************

RampShadow = 0      '0 = no , 1 = yes

'***********  Red Light under the scoop     *******************************************

ScoopLight = 1      '0 = no , 1 = yes

'***********  Enable RGB LEDS inside the Gumball Machine ********************************

GumballRGB = 1      '0 = no , 1 = yes

'***********  Render the Slot Machine Toy   *******************************************

SlotMachineMod = 1    '0 = no , 1 = yes

'***********  Render the Pyramid Toy with customizable light  ***************************

PyramidMod = 1
PyramidTopLight = 4   '0 = red, 1 = green, 2 = blue, 3 = yellow, 4 = purple, 5= cyan

'***********  Render the Mini Clock at ramps entrance  **********************************

MiniClockMod = 1      '0 = no , 1 = yes

'***********  Render the Invader Toy  ***************************************************

InvaderMod = 1      '0 = no , 1 = yes

'***********  Render the Mystic Seer  Toy   *******************************************

MysticSeerMod = 1       '0 = no , 1 = yes

'***********  Enable TargetMod (Targets with themed decals)   ***************************

TargetMod = 0       '0 = disabled , 1 = enabled

'***********  Enable TV Mod   ***********************************************************

TVMod = 1       '0 = disabled , 1 = enabled

'***********  Enable Lamp Mod (new lamp model with customizable color   ***************

LampMod = 1       '1 = no, 1 = silver metal, 2 = gold metal
LampLightColor = 1    '0 = yellow , 1= blue

'***********  Enable Town Square Mod  ***************************************************

TownSquarePostMod = 1   '0 = disabled , 1 = enabled

'***********  Enable the Spiral Animated Cover Mod  ***********************************

SpiralMod = 1       '0 = disabled , 1 = enabled

'***********  Enable Bumper Posts Mod ***************************************************

'the game will be harder if you enable this

BumperPostsMod = 1    '0 = disabled , 1 = enabled

'***********  Enable 3rd magnet *******************************************************

ExtraMagnet = 1     '0 = disabled , 1 = enabled

'******************************************************************************************
'* END OF TABLE OPTIONS *******************************************************************
'******************************************************************************************

Dim Ballsize,BallMass
BallSize = 51
BallMass = (Ballsize^3)/125000

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode
Const UseVPMModSol = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "02000000", "WPC.VBS", 3.49

'******************************
'******************************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 1
Const HandleMech = 0

Const SSolenoidOn = "fx_solon"
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "fx_coin"

Set MotorCallback = GetRef("GameTimer")
Set GiCallback2 = GetRef("UpdateGI")

'********************************************
'**********   Keys   ************************
'********************************************

Dim BIPL        'balls in plunger

Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 3:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If Keycode = KeyFront Then Controller.Switch(23) = 1
    If KeyCode = PlungerKey Then Plunger.Pullback:PlaySoundAt "fx_plungerpull", Plunger
    If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyCode = KeyFront Then Controller.Switch(23) = 0
    If KeyCode = PlungerKey Then
        Plunger.Fire
        If BIPL = 1 Then
            Playsound "fx_launch"
        Else
            Playsound "fx_plunger"
        End If
    End If
    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'************************************************
'*********   BallInit ***************************
'************************************************

Dim PowerBall, powerballnf, PowerballLocation

If PowerballStart = 0 then PowerballStart = RndNum(1, 6) end if
If PowerballStart = 7 then PowerballStart = RndNum(1, 3) end if

Sub Scramble
  If powerballstart = 1 Then Set PowerBallLocation = Gumball1
  If powerballstart = 2 Then Set PowerBallLocation = Gumball2
  If powerballstart = 3 Then Set PowerBallLocation = Gumball3
  If powerballstart = 4 Then Set PowerBallLocation = sw17
  If powerballstart = 5 Then Set PowerBallLocation = sw16
  If powerballstart = 6 Then Set PowerBallLocation = sw15
  PowerBallLocation.Destroyball
  Set PowerBall = PowerBallLocation.CreateSizedBall(Ballsize/2)
  With PowerBall
    .image = "powerball"
    .color = RGB(255,255,255)
    .id = 666
    .Mass = 0.8*Ballmass
  End With
End sub

'**************************************************************
'************   Table Init   **********************************
'**************************************************************

Dim bsSlot, bsAutoPlunger, bsRocket, mLeftMini, mRightMini, mslot, mLeftMagnet, mLowerRightMagnet, mUpperRightMagnet

Sub Table1_Init
  vpmInit Me
    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Twilight Zone (Bally 1992)"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .Hidden = DesktopMode
        .HandleMechanics = 1 + 2 'Gumball + Clock
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    '************  Main Timer init  ********************

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

    '************   Nudging   **************************

  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 4
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Leftslingshot, Rightslingshot)

    '************   AutoPlunger   ******************

    Set bsAutoPlunger = new cvpmSaucer
    With bsAutoPlunger
        .InitKicker AutoPlungerKicker, 72, 0, 40, 0
    .InitExitVariance 0, 8
        .InitSounds "", SoundFX("fx_AutoPlunger",DOFContactors), SoundFX("fx_launch",DOFContactors)
    End With

    '************   SlotMachine   ******************

  Set bsSlot = new cvpmSaucer
  With bsSlot
    .InitKicker sw58, 58, 4.8, 45, 0
    .InitSounds "fx_kicker_catch", SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_SlotM_exit",DOFContactors)
    .createevents "bsSlot", sw58
  End With

    Set mSlot = New cvpmMagnet      'Helps two balls get out, see sub slotmachinekickout
    With mSlot
        .InitMagnet Slothelper, -75
        .CreateEvents "mSlot"
        .GrabCenter = false
    End With

    '*******************   Rocket   ****************

  Set bsRocket = new cvpmSaucer
    With bsRocket
        .InitKicker RocketKicker, 28, 280, 55, 0
    .InitExitVariance 3, 10
        .InitSounds "fx_rocket_enter",  SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_rocket_exit",DOFContactors)
    End With

    '****************   Magnets   ******************

    Set mLeftMini = New cvpmMagnet
    With mLeftMini
        .InitMagnet TLMiniFlip, 70    ' Left Powerfield Real Magnet Strength (adjust to taste)
        .GrabCenter = False: .Size = 195
        .CreateEvents "mLeftMini"
    End With

    Set mRightMini = New cvpmMagnet
    With mRightMini
        .InitMagnet TRMiniFlip, 70    ' Right Powerfield Real Magnet Strength (adjust to taste)
        .GrabCenter = False: .Size = 195
        .CreateEvents "mRightMini"
    End With

    Set mLeftMagnet = New cvpmMagnet
    With mLeftMagnet
        .InitMagnet LeftMagnet, 67
        .CreateEvents "mLeftMagnet"
        .GrabCenter = True
    End With

    Set mUpperRightMagnet = New cvpmMagnet
    With mUpperRightMagnet
        .InitMagnet UpperRightMagnet, 63
        .CreateEvents "mUpperRightMagnet"
        .GrabCenter = True
    End With

    Set mLowerRightMagnet = New cvpmMagnet
    With mLowerRightMagnet
        .InitMagnet LowerRightMagnet, 63
        .CreateEvents "mLowerRightMagnet"
        .GrabCenter = True
    End With

    '****************   Trough   ******************

  sw17.CreateSizedBallWithMass Ballsize/2, BallMass
  sw16.CreateSizedBallWithMass Ballsize/2, BallMass
  sw15.CreateSizedBallWithMass Ballsize/2, BallMass

    '****************   Gumball   ******************

  gumball1.CreateSizedBallWithMass Ballsize/2, BallMass
  gumball2.CreateSizedBallWithMass Ballsize/2, BallMass
  gumball3.CreateSizedBallWithMass Ballsize/2, BallMass
  Controller.Switch(55) = 0

    '****************   Init GI   ******************

    UpdateGI 0, 1:UpdateGI 1, 1:UpdateGI 2, 1:UpdateGI 3, 1:UpdateGI 4, 1
End Sub

'**************************************************************
' SOLENOIDS
'**************************************************************

'   (*) - only in prototype, supported by rom 9.4
'  (**) - the additional GUM and BALL flashers were removed ro reduce cost
' (***) - Gumball and Clock Mechanics are handled by vpm classes

'standard coils
SolCallback(1) = "SlotMachineKickout"                 '(01) Slot Kickout
SolCallback(2) = "bsRocket.SolOut"                    '(02) Rocket Kicker
SolCallback(3) = "bsAutoPlunger.SolOut"                 '(03) Auto-Fire Kicker
SolCallback(4) = "SolGumballPopper"                   '(04) Gumball Popper
SolCallback(5) = "SolRightRampDiverter"                 '(05) Right Ramp Diverter
SolCallback(6) = "SolGumballDiverter"                 '(06) Gumball Diverter
SolCallback(7) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"    '(07) Knocker
SolCallback(8) = "SolOuthole"                     '(08) Outhole
SolCallback(9) = "SolBallRelease"                   '(09) Ball Release
'SolCallback(10) = "SolRightSling"                    '(10) Right Slingshot
'SolCallback(11) = "SolLeftSling"                   '(11) Left Slingshot
'SolCallback(12) = "SolLowerBumper"                   '(12) Lower Jet Bumper
'SolCallback(13) = "SolLeftBumper"                    '(13) Left Jet Bumper
'SolCallback(14) = "SolRightBumper"                   '(14) Right Jet Bumper
SolCallback(15) = "LockKickout"                     '(15) Lock Release nf
SolCallback(16) = "SolShootDiverter"                  '(16) Shooter Diverter
SolModCallback(17) = "SetModLamp 117,"                    '(17) Flasher bumpers x2
SolModCallback(18) = "SetModLamp 118,"                    '(18) Flasher Power Payoff x2
SolModCallback(19) = "SetModLamp 119,"                    '(19) Flasher Mini-Playfield x2
SolModCallback(20) = "SetModLamp 120,"                    '(20) Flasher Upper Left Ramp x2 (**)
SolCallback(21) = "SolLeftMagnet"                   '(21) Left Magnet
SolCallback(22) = "SolUpperRightMagnet"                 '(22) Upper Right Magnet (*)
SolCallback(23) = "SolLowerRightMagnet"                 '(23) Lower Right Magnet
SolCallback(24) = "SolGumballMotor"                   '(24) Gumball Motor
SolCallback(25) = "SolMiniMagnet mLeftMini,"              '(25) Left Mini-Playfield Magnet
SolCallback(26) = "SolMiniMagnet mRightMini,"             '(26) Right Mini-Playfield Magnet
SolCallback(27) = "SolLeftRampDiverter"                 '(27) Left Ramp Diverter
SolCallback(28) = "setlamp 128,"                    '(28) Flasher Inside Ramp
'aux board coils
SolModCallback(51) = "SetModLamp 137,"                    '(37) Flasher Upper Right Flipper
SolModCallback(52) = "SetModLamp 138,"                    '(38) Flasher Gumball Machine Higher
SolModCallback(53) = "SetModLamp 139,"                    '(39) Flasher Gumball Machine Middle
SolModCallback(54) = "SetModLamp 140,"                    '(40) Flasher Gumball Machine Lower
SolModCallback(55) = "SetModLamp 141,"                    '(41) Flasher Upper Right Ramp x2 (**)
'SolCallback(56) = ""                         '(42) Clock Reverse (***)
'SolCallback(57) = ""                         '(43) Clock Forward (***)
'SolCallback(58) = ""                         '(44) Clock Switch Strobe (***)
SolCallback(59) = "SolGumRelease"                   '(??) Gumball Release (***)
'fliptronic board
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"
SolCallback(sULFlipper) = "SolULFlipper"


'***********   Rocket   ********************************

Sub RocketKicker_Hit
    bsRocket.AddBall 0
    Light10.State = 1
End Sub

Sub RocketKicker_UnHit
    Light10.state = 0
End Sub

'***********   Autoplunger   ********************************

Sub AutoPlungerKicker_Hit
    bsAutoPlunger.AddBall 0
End Sub

'***********   Gumball Popper   ******************************

Dim BIK:BIK=0   'ball in kicker

sub GumballPopper_Hit()
  Controller.Switch(74) = 1
  BIK=BIK+1
  Playsound "fx_kicker_catch"
end sub

sub GumballPopper_UnHit()
  Controller.Switch(74) = 0
  BIK=BIK-1
end sub

Sub GumballPopperHole_Hit
  Playsoundat "fx_Hole", GumballPopperHole
    vpmTimer.PulseSw 51
End Sub

Sub SolGumballPopper(enabled) 'VUK
    If enabled Then
    GumballPopper.Kick 0, 65, 1.5
        If BIK = 0 Then
            PlaySound SoundFX(SSolenoidOn,DOFContactors)
        Else
            PlaySound SoundFX("fx_GumPop",DOFContactors)
        End If
    End If
End Sub

'**************   Drain  and Release  ************************************

Dim BIP:BIP = 0       'Balls In Play

sub sw18_hit()
  Controller.Switch(18) = 1:Playsound "fx_drain":BIP = BIP - 1
    If TVMod = 1 then TVTimer.enabled = 0:Frame.imageA = "tv_gameover"
end sub

sub sw18_unhit():controller.Switch(18) = 0:end sub

Sub SolOuthole(enabled)
  If Enabled Then
    sw18.kick 85, 11
    Updatetrough
    Playsoundat SoundFX(SSolenoidOn,DOFContactors), sw18
  End If
end sub

Sub SolBallrelease(enabled)
  If Enabled Then
    sw15.kick 58, 10
    Updatetrough          'this is important to reset trough intervals
    BIP = BIP + 1
  End If
End sub

'******************************************************************
'** NFOZZY'S BALLTROUGH AND LOCK  *******************************
'******************************************************************

sub boottrough_timer()        'starts up the trough
  controller.Switch(17) = 1
  controller.Switch(16) = 1
  controller.Switch(15) = 1
  if powerballstart <> 6 then
    controller.Switch(26) = 1 'Trough sensor, needs to start on unless the powerball spawns in the forward trough position
  end if
  me.enabled = 0
  updatetrough
  Scramble  'Powerball insertion
end sub

sub sw25_hit():controller.Switch(25) = 1:end sub
sub sw25_unhit():controller.Switch(25) = 0:updatetrough:end sub

sub sw17_hit():controller.Switch(17) = 1:end sub
sub sw17_unhit():controller.Switch(17) = 0:updatetrough:end sub

sub sw16_hit():controller.Switch(16) = 1:end sub
sub sw16_unhit():controller.Switch(16) = 0:updatetrough:end sub

sub sw15_hit()
if activeball.id = 666 then   'opto handler
  controller.Switch(26) = 0 'if powerball
Else
  controller.Switch(26) = 1 'if regular ball
end if
controller.Switch(15) = 1
end sub
sub sw15_unhit():controller.Switch(15) = 0:controller.switch(26) = 0:playsoundat SoundFX("fx_ballrel",DOFContactors), sw15:updatetrough:end sub 'if empty

sub sw26_hit
  if activeball.id = 666 then
    set powerballnf = activeball
  end if
end Sub

'order-
'25, 17, 16, 15

sub Updatetrough()
  updatetroughTimer.interval = 300
  updatetroughTimer.enabled = 1
end sub

sub updatetroughTimer_timer()
  if sw15.BallCntOver = 0 then sw16.kick 58, 8 end If
  if sw16.BallCntOver = 0 then sw17.kick 58, 8 end If
  if sw17.BallCntOver = 0 then sw25.Kick 58, 8 end If
  me.enabled = 0
end sub

'*******************   Lock   ******************

sub sw85_hit():controller.Switch(85) = 1:PlaysoundAt "fx_Lock_enter", sw85:updatelock:end sub
sub sw85_unhit():controller.Switch(85) = 0:end sub
sub sw84_hit():controller.Switch(84) = 1:updatelock:end sub
sub sw84_unhit():controller.Switch(84) = 0:end sub
sub sw88_hit():controller.Switch(88) = 1:updatelock:end sub
sub sw88_unhit():controller.Switch(88) = 0:end sub

sub updatelock
  updatelocktimer.interval = 32
  updatelocktimer.enabled = 1
end sub

sub updatelocktimer_timer()
  if sw88.BallCntOver = 0 then sw84.kick 180, 2 end If
  if sw84.BallCntOver = 0 then sw85.kick 180, 2 end if
  me.enabled = 0
end sub

sub LockKickout(enabled)
  If enabled then
  sw88.kick 88, 10
  Playsoundat SoundFX("fx_Lock_exit",DOFContactors), sw88
  updatelock
  End If
end sub

'******************************************************************
'** SUBWAY, SHOOTER LANE, SLOTMACHINE; CAMERA, PIANO, DEAD END  ***
'******************************************************************

Sub SubwaySound(dummy)
  PlaySound "fx_subway"
End sub

'******** Slot Machine  ****************************************

Sub SlotMachine_Hit()
    PlaySound "fx_SlotM_enter"
End Sub

Sub sw57_Hit()                      'submarine switch, Tslot proximity
if activeball.id = 666 then
  set powerballnf = activeball
else
controller.Switch(57) = 1
end if
end Sub

Sub sw57_UnHit():controller.Switch(57) = 0:end Sub

sub slotmachinekickout(enabled)
  if enabled Then
    bsSlot.SolOut enabled
    slotMachine.enabled = 0             'let's disable and reanable it again so the ball can pass
    vpmtimer.addtimer 250, "slotmachine.enabled = 1'"
    if SlotBall2.BallCntOver > 0 then
      mslot.MagnetOn=1
    end if
  Else
    mslot.MagnetOn=0
  end if
end sub

'******** Shooter Lane  ***************************************

Sub ShooterLaneKicker_Hit
    PlaySoundAt "fx_hole", ShooterLaneKicker
  vpmtimer.addtimer 100, "SubwaySound"
End Sub

'********   Dead End   ********************************************

Sub DeadEnd_Hit
    vpmTimer.PulseSw 41
  PlaySoundAt "fx_DeadEnd", DeadEnd
  vpmtimer.addtimer 100, "SubwaySound"
End Sub

'********   Camera   ***********************************************

Sub CameraKicker_Hit
  PlaySoundAt "fx_hole", CameraKicker
  vpmtimer.addtimer 100, "SubwaySound"
End Sub

Sub sw42_Hit():controller.Switch(42) = 1:end Sub    'submarine switch, camera / upper playfield
Sub sw42_UnHit():controller.Switch(42) = 0:end Sub

'********  Piano   *************************************************

Sub Piano_Hit()
    PlaySound "fx_Piano"
End Sub

Sub sw43_Hit():controller.Switch(43) = 1:end Sub    'submarine switch, piano
Sub sw43_UnHit():controller.Switch(43) = 0:end Sub

''************************************************************************************
''*****************       SLINGSHOTS                      ****************************
''************************************************************************************

Dim RStep, LStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshotL",DOFContactors)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    Me.TimerEnabled = 1
    vpmTimer.PulseSw 34
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:Me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  PlaySound SoundFX("fx_slingshotR",DOFContactors)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    Me.TimerEnabled = 1
    vpmTimer.PulseSw 35
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:Me.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

''************************************************************************************
''*****************               Bumpers                 ****************************
''************************************************************************************

Dim bump1, bump2, bump3

Sub Bumper1_Hit
    vpmTimer.PulseSw 31
    PlaySound SoundFX("fx_BumperLeft", DOFContactors)
    bump1 = 1:Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer()
    Select Case bump1
        Case 1:BR1.Z = 15:bump1 = 2
        Case 2:BR1.Z = 25:bump1 = 3
        Case 3:BR1.Z = 35:bump1 = 4
        Case 4:BR1.Z = 45:Me.TimerEnabled = 0
    End Select
End Sub

Sub Bumper2_Hit
    vpmTimer.PulseSw 32
    PlaySound SoundFX("fx_BumperRight", DOFContactors)
    bump2 = 1:Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer()
    Select Case bump2
        Case 1:BR2.Z = 15:bump2 = 2
        Case 2:BR2.Z = 25:bump2 = 3
        Case 3:BR2.Z = 35:bump2 = 4
        Case 4:BR2.Z = 45:Me.TimerEnabled = 0
    End Select
End Sub

Sub Bumper3_Hit
    vpmTimer.PulseSw 33
    PlaySound SoundFX("fx_BumperMiddle", DOFContactors)
    bump3 = 1:Me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer()
    Select Case bump3
        Case 1:BR3.Z = 15:bump3 = 2
        Case 2:BR3.Z = 25:bump3 = 3
        Case 3:BR3.Z = 35:bump3 = 4
        Case 4:BR3.Z = 45:Me.TimerEnabled = 0
    End Select
End Sub

'*******************************************************************
'****************   Diverters   ************************************
'*******************************************************************

Sub solShootDiverter(Enabled)
    If Enabled Then
        shooterdiverter.rotatetoend : Playsound SoundFX("fx_DivSS",DOFContactors)
    Else
        shooterdiverter.rotatetostart
  End If
End Sub

Sub SolGumballDiverter (enabled)
    If Enabled Then
    GumballDiverter.rotatetoend : Playsound SoundFX("fx_DivGM",DOFContactors)
  Else
    GumballDiverter.rotatetostart
  End If
End Sub

'********  Left Ramp Diverter   **********************

Sub SolLeftRampDiverter(enabled)
    If Enabled Then
    Playsound SoundFX("fx_DivLR",DOFContactors)
    RampDivWall.IsDropped=1
        RampDiverter.RotatetoEnd
    Plunger1.pullback
    Else
    RampDivWall.IsDropped=0
        RampDiverter.Rotatetostart
    Plunger1.fire
  End If
End Sub

'********  Right Ramp Diverter   **********************

Dim divDir, divPos
divPos = 0

Sub divKicker_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_metalHit"
End Sub

Sub SolRightRampDiverter(enabled)
    If enabled Then
    Playsound SoundFX("fx_DivRR",DOFContactors)
        DivKicker.Kick 0, 10, 34
        divDir = 9
    Else
        divDir = -9
    End If
    DiverterTimer.Enabled = 1
    DivKicker.Enabled = Not enabled
End Sub

Sub diverterTimer_Timer()
    divPos = divPos + divDir
    If divPos > 90 Then
        divPos = 90
        DiverterTimer.Enabled = 0
    End If
    If divPos < 0 Then
        divPos = 0
        diverterTimer.Enabled = 0
    End If
    RDiv.RotX = divPos
    SpiralToy.RotX = divPos
End Sub

'*******************************************************************
'*************************       Targets        ********************
'*******************************************************************

Sub sw47_Hit:vpmTimer.PulseSw 47:PlaySound SoundFX("fx_target",DOFContactors):End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:PlaySound SoundFX("fx_target",DOFContactors):End Sub
Sub sw64_Hit:vpmTimer.PulseSw 64:PlaySound SoundFX("fx_target",DOFContactors):End Sub
Sub sw65_Hit:vpmTimer.PulseSw 65:PlaySound SoundFX("fx_target",DOFContactors):End Sub
Sub sw65a_Hit:vpmTimer.PulseSw 65:PlaySound SoundFX("fx_target",DOFContactors):End Sub
Sub sw66_Hit:vpmTimer.PulseSw 66:PlaySound SoundFX("fx_target",DOFContactors):End Sub
Sub sw67_Hit:vpmTimer.PulseSw 67:PlaySound SoundFX("fx_target",DOFContactors):End Sub
Sub sw68_Hit:vpmTimer.PulseSw 68:PlaySound SoundFX("fx_target",DOFContactors):End Sub
Sub sw77_Hit:vpmTimer.PulseSw 77:PlaySound SoundFX("fx_target",DOFContactors):End Sub
Sub sw78_Hit:vpmTimer.PulseSw 78:PlaySound SoundFX("fx_target",DOFContactors):End Sub

'******************************************************
'***************   Mini PF Switches *******************
'******************************************************

Sub sw44_Hit:Controller.Switch(44) = 1 : End Sub
Sub sw44_Unhit:Controller.Switch(44) = 0 : End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45 : End Sub
Sub sw45a_Hit:vpmTimer.PulseSw 45 : End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46 : End Sub
Sub sw46a_Hit:vpmTimer.PulseSw 46 : End Sub
Sub sw75_Hit:Controller.Switch(75) = 1 : End Sub
Sub sw75_UnHit:Controller.Switch(75) = 0 : End Sub
Sub sw76_Hit:Controller.Switch(76) = 1: vpmtimer.addtimer 300, "BallDropSound": End Sub
Sub sw76_Unhit:Controller.Switch(76) = 0 : End Sub

'******************************************************
'***************  Ramps Switches **********************
'******************************************************

Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySound "fx_Gate" : End Sub
Sub sw54_Hit:Controller.Switch(54) = 1 : PlaySound "fx_Gate" : LRampSw.rotatetoend : End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0 : LRampSw.rotatetostart : End Sub
Sub sw73_Hit:vpmTimer.PulseSw 73: Playsound "fx_metalrolling" : End Sub

'**************************************************************
'***************  Rollover Switches   *************************
'**************************************************************

Sub sw11_Hit:Controller.Switch(11)= 1:PlaySound "fx_sensor" : End Sub
Sub sw11_Unhit:Controller.Switch(11) = 0:End Sub
Sub sw12_Hit:Controller.Switch(12)= 1:PlaySound "fx_sensor" : End Sub
Sub sw12_Unhit:Controller.Switch(12) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36)= 1:PlaySound "fx_sensor" : End Sub
Sub sw36_Unhit:Controller.Switch(36)= 0:End Sub
Sub sw37_Hit:Controller.Switch(37)= 1:PlaySound "fx_sensor" : End Sub
Sub sw37_Unhit:Controller.Switch(37)= 0:End Sub
Sub sw38_Hit:Controller.Switch(38)= 1:PlaySound "fx_sensor" : End Sub
Sub sw38_Unhit:Controller.Switch(38)= 0:End Sub
Sub sw52_Hit: Controller.Switch(52) = 1:PlaySound "fx_sensor" : End Sub
Sub sw52_Unhit:Controller.Switch(52)= 0:End Sub
Sub sw56_Hit: Controller.Switch(56) = 1:PlaySound "fx_sensor" : End Sub
Sub sw56_Unhit:Controller.Switch(56)= 0:End Sub
Sub sw61_Hit: Controller.Switch(61) = 1:PlaySound "fx_sensor" : End Sub
Sub sw61_Unhit:Controller.Switch(61)= 0:End Sub
Sub sw62_Hit: Controller.Switch(62) = 1:PlaySound "fx_sensor" : End Sub
Sub sw62_Unhit:Controller.Switch(62)= 0:End Sub
Sub sw63_Hit: Controller.Switch(63) = 1:PlaySound "fx_sensor" : End Sub
Sub sw63_Unhit:Controller.Switch(63)= 0:End Sub

'**************************************************************
'***************  Opto Switches   *********************
'**************************************************************

Sub sw81_Hit:Controller.Switch(81) = 1:End Sub
Sub sw81_UnHit:Controller.Switch(81) = 0:End Sub
Sub sw82_Hit:Controller.Switch(82) = 1:End Sub
Sub sw82_UnHit:Controller.Switch(82) = 0:End Sub
Sub sw83_Hit:Controller.Switch(83) = 1:End Sub
Sub sw83_UnHit:Controller.Switch(83) = 0:End Sub

'Clock Pass Opto (only in prototypes,supported by rom version 9.4)
Sub sw86_Hit:Controller.Switch(86) = 1:End Sub
Sub sw86_UnHit:Controller.Switch(86) = 0:End Sub
'Gumball entry opto
sub sw87_hit():controller.switch(87) = 1:end Sub
sub sw87_unhit():controller.switch(87) = 0:end Sub
''Autoplunger 2nd Opto (only in prototypes,supported by rom version 9.4)
'Sub sw71_Hit:Controller.Switch(71) = 1:End Sub
'Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

'*************************************************************
'***************   Shooting lane   ***************************
'*************************************************************

Sub sw27_Hit():controller.switch(27) = 1:BIPL = 1
If TVMod = 1 then TVTimer.enabled = 1:Frame.imageA = "tv_1"
End Sub

Sub sw27_UnHit():controller.switch(27) = 0:BIPL = 0:End Sub

''************************************************************
''*****************   Flippers    ****************************
''************************************************************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperUp",DOFContactors)
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors)
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperUp",DOFContactors)
        RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors)
        RightFlipper.RotateToStart
    End If
End Sub

Sub SolULFlipper(Enabled)
    If Enabled Then
        LeftMiniFlipper.RotateToEnd
    Else
        LeftMiniFlipper.RotateToStart
    End If
End Sub

Sub SolURFlipper(Enabled)
    If Enabled Then
        RightminiFlipper.RotateToEnd
    Else
        RightMiniFlipper.RotateToStart
    End If
End Sub

'********************************************************************************
'*******************    REATIME UPDATES   *********************************
'********************************************************************************

Sub GameTimer
    UpdateClock
    RollingSoundUpdate
    UpdateMechs
End Sub

Sub UpdateMechs
  FlipperL.RotZ = LeftFlipper.CurrentAngle
  FlipperR.RotZ = RightFlipper.CurrentAngle
  FlipperL1.RotY = LeftMiniFlipper.CurrentAngle
  FlipperR1.RotZ = RightMiniFlipper.CurrentAngle
  ShooterDiv.RotY = ShooterDiverter.CurrentAngle
  DiverterP.RotZ = RampDiverter.CurrentAngle
  DiverterP1.RotZ = GumballDiverter.CurrentAngle
  sw53p.RotX = LRampG.CurrentAngle
  sw54p.RotX = LRampSw.CurrentAngle
End Sub

'********************************************************************************
'******************  NFOZZY'S GUMBALL MACHINE   *********************************
'********************************************************************************

sub Gumball0_hit():updategumballs:end Sub
sub Gumball1_hit():updategumballs:end Sub
sub Gumball2_hit():updategumballs:end Sub
sub Gumball3_hit():updategumballs:end Sub
sub Gumball3_Unhit():updategumballs:end Sub

sub updategumballs
  gumballupdate.interval = 50
  gumballupdate.enabled = 1
End sub

sub Gumballupdate_timer()
  if gumball3.BallCntOver = 0 then gumball2.kick 170, 1 end If
  if gumball2.BallCntOver = 0 then gumball1.kick 170, 1 end If
  if gumball1.BallCntOver = 0 then gumball0.kick 170, 1 end If
  me.enabled = 0    'Might cause problems
end sub

Sub SolGumRelease(enabled)
    If enabled Then
    gumball3.kick 170, 1
    updategumballs
    vpmTimer.AddTimer 100, "GumKnobTimer.enabled = 1'"
    vpmtimer.PulseSw 55
    End If
End Sub

Sub SolGumballMotor(enabled)
  If enabled Then
    PlaySound SoundFX("fx_GumMachine",DOFGear)
  End If
End Sub

Sub GumKnobTimer_Timer()
  GumballMachineKnob.RotY= GumballMachineKnob.RotY + 20
  If GumballMachineKnob.RotY >  360 then GumKnobTimer.enabled = 0 : GumballMachineKnob.RotY = 0
End Sub

'********************************************************************
'*************************   CLOCK   ********************************
'********************************************************************
Dim LastTime
LastTime = 0

Sub UpdateClock
    Dim Time, Min, Hour, temp
    Time = CInt(Controller.GetMech(0) )
    If Time <> LastTime Then
    Min = (Time Mod 60)
    Hour = Int(Time / 2)
    ClockShort.RotY = Hour - 45
    ClockLarge.RotY = min * 6
    Clock_mech.RotY = min * 6
    LastTime = Time
    PlaySound SoundFXDOF("fx_motor",101,DOFPulse,DOFGear), -1, 1, 0.05, 0, 0, 1, 0
  Else Stopsound "fx_motor"
    End If
End Sub

'**********************************************************************
'**************   POWER FIELD MAGNETS *********************************
'**********************************************************************

Sub SolMiniMagnet(aMag, enabled)
    If enabled Then
    PlaySound SoundFX("fx_magnet",DOFShaker)
        With aMag
      .removeball powerballnf
            .MagnetOn = True
            .Update
            .MagnetOn = False
        End With
    End If
End Sub

'**********************************************************************
' SPECIAL CODE BY NFOZZY TO HANDLE MAGNET TRIGGERS
' based on the code by KIEFERSKUNK/DORSOLA
' Method: on extra triggers unhit, kill the velocity of the
' ball if the magnet is on, helping the magnet catch the ball.
'**********************************************************************

Sub sw81_help_unhit
  If activeball.vely > 28 then  activeball.vely = RndNum (26,27)      '-ninuzzu- Let's slow down the ball a bit so the magnets can
  If activeball.vely < - 28 then  activeball.vely = - RndNum (26,27)    'catch the ball
  If mLowerRightMagnet.MagnetOn = 1 and activeball.id <> 666 then
    activeball.vely = activeball.vely * -0.2
    activeball.velx = activeball.velx * -0.2
  End If
End Sub

Sub sw82_help_unhit
  If mUpperRightMagnet.MagnetOn = 1 and activeball.id <> 666 then
    activeball.vely = activeball.vely * -0.2
    activeball.velx = activeball.velx * -0.2
    End If
End Sub

Sub sw83_help_unhit
  If mLeftMagnet.MagnetOn = 1 and activeball.id <> 666 then
    activeball.vely = activeball.vely * -0.2
    activeball.velx = activeball.velx * -0.2
  End If
End Sub

Sub SolLeftMagnet(enabled)
  If enabled Then
    mLeftMagnet.MagnetOn = 1
    mleftmagnet.removeball powerballnf
    PlaySound SoundFX("fx_magnet_catch",DOFShaker)
  Else
    mLeftMagnet.MagnetOn = 0
  End If
End Sub

Sub SolUpperRightMagnet(enabled)
  If enabled Then
    mUpperRightMagnet.MagnetOn = 1
    mUpperRightMagnet.removeball powerballnf
    PlaySound SoundFX("fx_magnet_catch",DOFShaker)
  Else
    mUpperRightMagnet.MagnetOn = 0
  End If
End Sub

Sub SolLowerRightMagnet(enabled)
  If enabled Then
    mLowerRightMagnet.MagnetOn = 1
    mLowerRightMagnet.removeball powerballnf
    PlaySound SoundFX("fx_magnet_catch",DOFShaker)
  Else
    mLowerRightMagnet.MagnetOn = 0
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
LampTimer.Interval = 20 'lamp fading speed
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
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeLm 13, l13
    NFadeLm 13, l13a  'Robbie lights
    NFadeL 13, l13b   'Robbie lights
    NFadeL 14, l14
    NFadeL 15, l15
    NFadeL 16, l16
    NFadeLm 17, l17
    NFadeLm 17, l17a    'TownSquarePost lights
    NFadeLm 17, l17a1   'TownSquarePost lights
    NFadeLm 17, l17a2   'TownSquarePost lights
    Flash 17, l17r    'TownSquarePost lights
    NFadeL 18, l18

    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeLm 27, l27
    NFadeLm 27, l27a  'SlotMachine Lights
    Flash 27, l27b    'SlotMachine Lights
    NFadeLm 28, l28
    NFadeLm 28, l28a  'Robbie lights
    NFadeL 28, l28b   'Robbie lights

    NFadeL 31, l31
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeLm 36, l36
    NFadeLm 36, l36a  'Robbie lights
    NFadeL 36, l36b   'Robbie lights
    NFadeL 37, l37
    NFadeL 38, l38

    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
    NFadeL 46, l46
    NFadeL 47, l47
    NFadeL 48, l48

  NFadeL 51, l51
  NFadeL 52, l52
  NFadeL 53, l53
  NFadeL 54, l54
  NFadeLm 55, l55
  NFadeLm 55, l55a  'Camera lights
  NFadeLm 55, l55b  'Camera lights
  Flash 55, l55c    'Camera lights
  NFadeLm 56, l56
  Flashm 56, l56a   'Pyramid lights
  Flash 56, l56b    'Pyramid lights
    NFadeLm 57, l57   'Robbie lights
    Flashm 57, l57a   'Robbie lights
    Flash 57, l57b
    NFadeL 58, l58

    NFadeLm 61, l61
    NFadeLm 61, l61a
    NFadeL 61, l61b
    NFadeLm 62, l62
    NFadeLm 62, l62a
    NFadeL 62, l62b
    NFadeLm 63, l63
    NFadeLm 63, l63a
    NFadeL 63, l63b
    NFadeL 64, l64
    NFadeL 65, l65
    NFadeL 66, l66
    NFadeL 67, l67
    NFadeL 68, l68

    NFadeL 71, l71
    NFadeL 72, l72
    NFadeL 73, l73
    NFadeL 74, l74
    NFadeL 75, l75
    NFadeL 76, l76
    NFadeL 77, l77
    NFadeL 78, l78

    NFadeLm 81, l81
    Flash 81, l81a
    NFadeLm 82, l82
    NFadeLm 82, l82a  'Clock Toy Mod lights
    Flash 82, l82r    'Clock Toy Mod lights
    NFadeLm 83, l83
    NFadeL 83, l83a
    NFadeLm 84, l84
    NFadeL 84, l84a
    NFadeLm 85, l85
    Flashm 85, l85a   'SlotMachine Lights
    Flash 85, l85b    'SlotMachine Lights
    NFadeLm 86, l86
    Flash 86, l86r

    'Flashers
  NFadeLModm 117, f17
  NFadeLMod 117, f17a
  NFadeLModm 118, f18
  NFadeLModm 118, f18a
  NFadeLModm 118, f18b
  NFadeLModm 118, f18c
  NFadeLModm 118, f18d
  Flash 118, f18e
  NFadeLModm 119, f19
  NFadeLModm 119, f19a
  Flash 119, f19b     'Pyramid lights
  NFadeLModm 120, f20
  NFadeLModm 120, f20a
  NFadeLModm 120, f20b
  NFadeLModm 120, f20c
  NFadeLModm 120, f20d
  Flashm 120, f20e
  Flash 120, f20r
  NFadeLModm 128, f28
  NFadeLModm 128, f28a
  NFadeLModm 128, f28b
  Flashm 128, f28c
  Flash 128, f28r
  NFadeLModm 137, f37
  NFadeLModm 137, f37a    'Rocket lights
  NFadeLMod 137, f37b   'Rocket lights
  NFadeLModm 138, f38
  NFadeLModm 138, f38a
  Flashm 138, f38b
  Flash 138, f38r
  NFadeLModm 139, f39
  NFadeLModm 139, f39a
  Flashm 139, f39b
  Flash 139, f39r
  NFadeLModm 140, f40
  NFadeLModm 140, f40a
  Flashm 140, f40b
  Flash 140, f40r
  NFadeLModm 141, f41
  NFadeLModm 141, f41a
  NFadeLModm 141, f41b
  NFadeLModm 141, f41c
  NFadeLModm 141, f41d
  Flashm 141, f41e
    Flash 141, f41r

    If BIP > 0 AND SpiralMod = 1 Then SpiralMove.enabled = LampState(68)
End Sub


Sub SetModLamp(nr, level)
  debug.print nr & ": " & level
  FadingLevel(nr) = level
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeLMod(nr, object) '3 lights
    object.Intensity = FadingLevel(nr) / 10
  object.state = 1
    object.visible = 1
End Sub

Sub NFadeLModm(nr, object) ' used for multiple lights
  NFadeLMod nr, Object
End Sub

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

'*********
'Update GI
'*********

Dim gistep

Sub UpdateGI(no, step)
    Dim xx
    If step = 0 then exit sub 'only values from 1 to 8 are visible and reliable. 0 is not reliable and 7 & 8 are the same so...
    gistep = (step-1) / 7
    Select Case no
        Case 0
            For each xx in GILeft:xx.IntensityScale = gistep:next   'Playfield Left
        Case 1
            For each xx in GIMinipf:xx.IntensityScale = gistep:next 'Mini Playfield + Insert
        Case 2
            For each xx in GIClock:xx.IntensityScale = gistep:next  'Clock + Insert
        Case 3
            For each xx in GImain:xx.IntensityScale = gistep:next   'Insert Main
        Case 4
            For each xx in GIRight:xx.IntensityScale = gistep:next  'Playfield Right
    End Select

    ' change the intensity of the flasher depending on the gi to compensate for the gi lights being off
    For xx = 0 to 200
        FlashMax(xx) = 6 - gistep * 3 ' the maximum value of the flashers
    Next
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function Vol(ball)                             ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 500)
End Function

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / table1.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'    JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = FALSE
    Next
End Sub

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = FALSE
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = TRUE
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = TRUE Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = FALSE
            End If
        End If
    Next
End Sub

'*****************************************
' JP's VP10 Ball Collision Sound
'*****************************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 500, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
' JF's Sound Routines
'*****************************************

Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub Posts_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
        PlaySound "fx_rubber", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "rubber_hit_1", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2:PlaySound "rubber_hit_2", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3:PlaySound "rubber_hit_3", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub LeftMiniFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RightMiniFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "flip_hit_1", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2:PlaySound "flip_hit_2", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3:PlaySound "flip_hit_3", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

'**********************
' Balldrop & Ramp Sound
'**********************

Sub BallDropSound(dummy)
    PlaySound "fx_ball_drop"
End Sub

Sub Balldrop1_Hit()
    vpmtimer.addtimer 300, "BallDropSound"
    StopSound "fx_metalrolling"
End Sub

Sub Balldrop2_Hit()
    vpmtimer.addtimer 100, "BallDropSound"
End Sub

Sub Balldrop3_Hit()
    vpmtimer.addtimer 300, "BallDropSound"
    StopSound "fx_metalrolling"
End Sub

Sub Balldrop4_Hit()
    PlaySound "fx_power"
End Sub

Sub WirerampSound1_Hit()
    PlaySound "fx_metalrolling"
End Sub

Sub WirerampSound2_Hit()
    PlaySound "fx_metalrolling"
End Sub

Sub WirerampSoundStop_Hit()
    StopSound "fx_metalrolling"
End Sub

Sub LREnter_Hit()
    If ActiveBall.VelY < 0 Then PlaySound "fx_rlenter"
End Sub

Sub RREnter_Hit()
    If ActiveBall.VelY < 0 Then PlaySound "fx_rrenter"
End Sub

'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************

Dim TVPic, SlotPic, RGBStep, RGBFactor, Red, Green, Blue, cGameName

TVPic = 0
SlotPic = 0
RGBStep = 0
RGBFactor = 1
Red = 255
Green = 0
Blue = 0

LeftCab.visible = CabinetSide
RightCab.visible = CabinetSide
led6.visible = NOT DesktopMode
led7.visible = DesktopMode
RGBLed1.visible = DesktopMode
RailL.visible = DesktopMode
RailR.visible = DesktopMode
pfshadow.visible = RampShadow
Scoop1L.visible = ScoopLight
Scoop2L.visible = ScoopLight
RGBLed3.visible = GumballRGB
RGBLed4.visible = GumballRGB
RGBLed5.visible = GumballRGB
RGBLed6.visible = GumballRGB
RGBLed7.visible = GumballRGB
RGBLed8.visible = GumballRGB
RGBLed9.visible = GumballRGB
RGBLed10.visible = GumballRGB
SlotMachineToy.visible = SlotMachineMod
swSlotReel.enabled = SlotMachineMod
SlotReel.visible = SlotMachineMod
SlotReelLight1.visible = SlotMachineMod
SlotReelLight2.visible = SlotMachineMod
l27a.visible = SlotMachineMod
l27b.visible = SlotMachineMod
l85a.visible = SlotMachineMod
l85b.visible = SlotMachineMod
Pyramid.visible = PyramidMod
PyramidCap.visible = PyramidMod
GI_Pyr1.visible = PyramidMod
GI_Pyr2.visible = PyramidMod
GI_Pyr3.visible = PyramidMod
l56a.visible = PyramidMod
l56b.visible = PyramidMod
f19b.visible = PyramidMod
ClockToy.visible = MiniClockMod
MysticSeerToy.visible = MysticSeerMod
InvaderToy.visible = InvaderMod
l00.visible = InvaderMod
led3.visible = InvaderMod
led4.visible = InvaderMod
led5.visible = InvaderMod
l82a.visible = MiniClockMod
l82r.visible = MiniClockMod
TVToy.visible = TVMod
Frame.visible = TVMod
SpiralToy.visible = SpiralMod
URMagnetP.visible = ExtraMagnet
sw82.enabled = ExtraMagnet
sw82_help.enabled = ExtraMagnet

If FlipperType = 2 then FlipperType = RndNum(0, 1)

If TargetMod = 1 Then
    sw47.image = "target-clock"
    sw48.image = "target-greed"
    sw64.image = "target-greed"
    sw65.image = "target-power"
    sw65a.image = "target-power"
    sw66.image = "target-greed"
    sw67.image = "target-greed"
    sw68.image = "target-coins"
    sw77.image = "target-greed"
    sw78.image = "target-greed"
End If

If TownSquarePostMod = 1 Then
    Rubber12.visible = 0
    TownSquarePost.visible = 1
    TownSquarePostW.visible = 1
    TownSquarePostBulb.visible = 1
    l17a.visible = 1
    l17a1.visible = 1
    l17a2.visible = 1
    l17r.visible = 1
Else
    l17a.visible = 0
    l17a1.visible = 0
    l17a2.visible = 0
    l17r.visible = 0
End If

Select Case Romset
      Case 0: cGameName = "tz_94ch"
      Case 1: cGameName = "tz_94h"
End Select

Select Case FlipperType
    Case 0:FlipperL.visible = 1
        FlipperR.visible = 1
        FlipperL1.visible = 1
        FlipperR1.visible = 1
    Case 1:LeftFlipper.visible = 1
        RightFlipper.visible = 1
        LeftMiniFlipper.visible = 1
        RightMiniFlipper.visible = 1
        LogoL.visible = 1
        LogoR.visible = 1
        LogoL1.visible = 1
        LogoR1.visible = 1
End Select

Select Case PyramidTopLight
  'Red
    Case 0 :  l56a.color = RGB (255,0,0):l56a.colorfull = RGB (255,170,170)
          l56b.color = RGB (255,0,0):f19b.color = RGB (255,0,0)
  'Green
    Case 1 :  l56a.color = RGB (0,255,0):l56a.colorfull = RGB (170,255,170)
          l56b.color = RGB (0,255,0):f19b.color = RGB (0,255,0)
  'Blue
    Case 2 :  l56a.color = RGB (0,0,255):l56a.colorfull = RGB (170,170,255)
          l56b.color = RGB (0,0,255):f19b.color = RGB (0,0,255)
  'Yellow
    Case 3 :  l56a.color = RGB (255,255,0):l56a.colorfull = RGB (255,255,170)
          l56b.color = RGB (255,255,0):f19b.color = RGB (255,255,0)
  'Magenta
    Case 4 :  l56a.color = RGB (255,0,255):l56a.colorfull = RGB (255,170,255)
          l56b.color = RGB (255,0,255):f19b.color = RGB (255,0,255)
  'Cyan
    Case 5 :  l56a.color = RGB (0,255,255):l56a.colorfull = RGB (170,255,255)
          l56b.color = RGB (0,255,255):f19b.color = RGB (0,255,255)
End Select

Select Case LampMod
    Case 0 :  TopPFLamp.visible=1
          TopPFLampMod.visible=0
          light46.visible=0
          light47.visible=0
    Case 1 :  TopPFLamp.visible=0
          TopPFLampMod.visible=1
          light46.visible=1
          light47.visible=1
    Case 2 :  TopPFLamp.visible=0
          TopPFLampMod.visible=1
          TopPFLampMod.image="gold"
          light46.visible=1
          light47.visible=1
End Select

If LampLightColor = 1 then
          light68.color = RGB (0,0,255):light68.colorfull = RGB (85,85,255)
          light19.color = RGB (0,0,255):light19.colorfull = RGB (85,85,255)
          light17.color = RGB (0,0,255):light17.colorfull = RGB (85,85,255)
          light44.color = RGB (0,0,255):light44.colorfull = RGB (85,85,255)
          light45.color = RGB (0,0,255):light45.colorfull = RGB (85,85,255)
          light46.color = RGB (0,0,255)
          light47.color = RGB (0,0,255)
          light48.color = RGB (0,0,255):light45.colorfull = RGB (85,85,255)
End If

Sub swSlotReel_Hit:SlotReelTimer.enabled = 1:End Sub

Sub SlotReelTimer_Timer()
    SlotPic = SlotPic + 1
    if SlotPic > 10 Then SlotReelTimer.enabled = 0:ResetSlot.enabled = 1
    SlotReel.imageA = "slot_" & SlotPic
End Sub

Sub ResetSlot_Timer()
    SlotReel.imageA = "slot_10":vpmtimer.addtimer 150, "SlotReel.imageA = ""slot_0"" '":SlotPic = 0
    If SlotPic = 0 then ResetSlot.enabled = 0
End Sub

Sub TVTimer_Timer()
    TVPic = TVPic + 1
    if TVPic = 34 Then TVPic = 2
    Frame.imageA = "tv_" & TVPic
End Sub

Sub SpiralMove_Timer()
    SpiralToy.rotz = SpiralToy.rotz + 10
End Sub

Select Case BumperPostsMod
    Case 0:   PegPlastic7.visible = 1
          PegPlastic8.visible = 1
          Rubber9.visible = 1
          Rubber29.visible = 1
          Rubber9.collidable = 1
          Rubber29.collidable = 1
    Case 1:   PegPlastic7.visible = 0
          PegPlastic8.visible = 0
          Rubber9.visible = 0
          Rubber29.visible = 0
          Rubber9.collidable = 0
          Rubber29.collidable = 0
End Select


Sub RGBTimer_timer 'rainbow light color changing
    Select Case RGBStep
        Case 0 'Green
            Green = Green + RGBFactor
            If Green > 255 then
                Green = 255
                RGBStep = 1
        led1.state=0
        led2.state=0
            End If
        Case 1 'Red
            Red = Red - RGBFactor
            If Red < 0 then
                Red = 0
                RGBStep = 2
        led1.state=1
        led2.state=1
            End If
        Case 2 'Blue
            Blue = Blue + RGBFactor
            If Blue > 255 then
                Blue = 255
                RGBStep = 3
        led1.state=0
        led2.state=0
            End If
        Case 3 'Green
            Green = Green - RGBFactor
            If Green < 0 then
                Green = 0
                RGBStep = 4
        led1.state=1
        led2.state=1
            End If
        Case 4 'Red
            Red = Red + RGBFactor
            If Red > 255 then
                Red = 255
                RGBStep = 5
        led1.state=0
        led2.state=0
            End If
        Case 5 'Blue
            Blue = Blue - RGBFactor
            If Blue < 0 then
                Blue = 0
                RGBStep = 0
        led1.state=1
        led2.state=1
            End If
    End Select
    'RGBLed.color = RGB(Red\10, Green\10, Blue\10)
    RGBLed.colorfull = RGB(Red, Green, Blue)
  RGBLed1.colorfull = RGB(Red, Green, Blue)
  RGBLed2.colorfull = RGB(Red, Green, Blue)
  RGBLed3.colorfull = RGB(Red, Green, Blue)
  RGBLed4.colorfull = RGB(Red, Green, Blue)
  RGBLed5.colorfull = RGB(Red, Green, Blue)
  RGBLed6.colorfull = RGB(Red, Green, Blue)
  RGBLed7.colorfull = RGB(Red, Green, Blue)
  RGBLed8.colorfull = RGB(Red, Green, Blue)
  RGBLed9.colorfull = RGB(Red, Green, Blue)
  RGBLed10.colorfull = RGB(Red, Green, Blue)
End Sub

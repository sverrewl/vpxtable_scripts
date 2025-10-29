' _______ __________________ _______  _______  _
'(  ___  )\__   __/\__   __/(  ___  )(  ____ \| \    /\
'| (   ) |   ) (      ) (   | (   ) || (    \/|  \  / /
'| (___) |   | |      | |   | (___) || |      |  (_/ /
'|  ___  |   | |      | |   |  ___  || |      |   _ (
'| (   ) |   | |      | |   | (   ) || |      |  ( \ \
'| )   ( |   | |      | |   | )   ( || (____/\|  /  \ \
'|/     \|   )_(      )_(   |/     \|(_______/|_/    \/
'
' _______  _______  _______  _______    _______  _______  _______  _______
'(  ____ \(  ____ )(  ___  )(       )  (       )(  ___  )(  ____ )(  ____ \
'| (    \/| (    )|| (   ) || () () |  | () () || (   ) || (    )|| (    \/
'| (__    | (____)|| |   | || || || |  | || || || (___) || (____)|| (_____
'|  __)   |     __)| |   | || |(_)| |  | |(_)| ||  ___  ||     __)(_____  )
'| (      | (\ (   | |   | || |   | |  | |   | || (   ) || (\ (         ) |
'| )      | ) \ \__| (___) || )   ( |  | )   ( || )   ( || ) \ \__/\____) |
'|/       |/   \__/(_______)|/     \|  |/     \||/     \||/   \__/\_______)
'
'
'
'
'  ______________               ___ __  _
'  / ____/ ____/ /__   ___  ____/ (_) /_(_)___  ____
' / / __/___ \/ //_/  / _ \/ __  / / __/ / __ \/ __ \
'/ /_/ /___/ / ,<    /  __/ /_/ / / /_/ / /_/ / / / /
'\____/_____/_/|_|   \___/\__,_/_/\__/_/\____/_/ /_/
'


'//////////////////////////////////////////////////////////////////////
' Update VPINMAME.DLL to support WPC-95 FASTFLIPS (SAMBuild r4550 or newer)
'http://vpuniverse.com/forums/topic/3461-sambuild31-beta-thread/
'//////////////////////////////////////////////////////////////////////


' Attack from Mars - IPDB No. 3781
' © Bally Williams 1995
' Design by Brian Eddy
' VPX recreation by g5k
' Thanks to:
' Tom Tower: Ramp, Main Saucer and Alien primitives
' DJRobX: Fastflips WPC95, SSF and sound sweeten, develop assistance
' JimmyFingers for passing on lots of helpful knowledge to do with GI
' Previous AFM VPX authors who's work helped guide the build on this one and to the components that are carried over
' To those in the Monster Bash Pincab and VP Junkies community that helped contributed artwork, photos and the banter
' The VPDevTeam and authors who have contributed to the projects and components that get passed on from table to table
' I'm sure there are primitives, scripts, sounds and so on that have been by others not mentioned above, thank you for your contributions!
' Thanks to Monsterbash Pincab for distribution and providing such a great resource for the community


' AFM Gameplay rules
' http://pinball.org/rules/attackfrommars.html


'//////////////////////////////////////////////////////////////////////
'// LAYER 01 = STANDARD VP OBJECTS
'// LAYER 02 = MAIN SAUCER + DIVERTER
'// LAYER 03 = INSERT LIGHTS
'// LAYER 04 = RED FLASHERS
'// LAYER 05 = BLUE/GREEN FLASHERS
'// LAYER 06 = G5k 3D
'// LAYER 07 = GI - Lower Zone
'// LAYER 08 = GI - Mid Zone
'// LAYER 09 = GI - Upper Zone
'// LAYER 10 = INSERTs Spill to other objects (Cheap SSR)
'// LAYER 11 = INSERTs LIGHT GLOW/BALL REFLECT
'//////////////////////////////////////////////////////////////////////


' 1.3 - fluffhead35 - Added Flipper Logic and Triggers.  Changed ballmass to 1. Changed Flipper Physics. Changed Ballspeed to Ballspeed1
'                   - Updated playfield physics, min/max slope, and gameplay difficulty. Setup rubber Dampeners.
' 1.3.1  - Added Fleep Collections, Changed RUBBERS to RubbersPrim. Changed Materials on objects on table and added to collections.  Updating slings, bumpers, gates
' 1.3.2. - Adding Fleep Sounds and Logic, added wire ramp exit sound, adjusted ramproll logic so you can provide a volume multiplyer to it.  Added New Rubberizer and fixed livecatch
'        - lowered target bouncer for sleeves.  Added target bouncer to TargetsBounce_Hit subroutine and created new TargetsBounce collection
' 1.3.3  - Added SleevesBounce collection and hit for Sleeves Bounce
' 1.3.4  - cleaned up unused sounds
' 1.3.5  - Adding Saucer Sounds and knocker and new wireramp_stop sounds,
' 1.3.6  - All the collidable ramps that existed in layer 1, 2 and 3 were removed and replaced by new primitives with gaps at the ends (layer 1). Fall through kickers disabled.
' 1.3.7  - Added Wire Ramp Off sounds to end of wire ramps
' 1.3.8  - Changed Scoop Kicker power, adjusted Playfied Friction to .20 on both playfield mesh and playfield.  Changed zCol_Ramp material friction to .12
' 1.3.9  - Imported Playfield Mesh from Tomate, Set Rubbers to not be collidable.
' 1.3.10 - Immported new Ramp from Tomate.
' 1.3.11 - Updated RampRoll and BallRoll to use max amplification. Updated Ramp physics to be inline with guide.

'NOTE for 120hz users, select the object on the table, "LAMPTIMER" on layer 1 and change the value to 16 which will force the speed to match how it looks at 60hz
'-1 (this is default and correct for 60hz)



Option Explicit
Randomize


'///////////////////////////////////////////////////////////////////////////////////////////////////
' _       ______  ______     ____  ______   _________   ___________   ________    ________  _____
'| |     / / __ \/ ____/    / __ \/ ____/  / ____/   | / ___/_  __/  / ____/ /   /  _/ __ \/ ___/
'| | /| / / /_/ / /  ______/ /_/ /___ \   / /_  / /| | \__ \ / /    / /_  / /    / // /_/ /\__ \
'| |/ |/ / ____/ /__/_____/\__, /___/ /  / __/ / ___ |___/ // /    / __/ / /____/ // ____/___/ /
'|__/|__/_/    \____/     /____/_____/  /_/   /_/  |_/____//_/    /_/   /_____/___/_/    /____/
'
'
'///////////////////////////////////////////////////////////////////////////////////////////////////
'// FAST FLIPS (DJRobX) - needs vpinmame.dll
'// Available in SAMupdate 4550 or newer
'// http://vpuniverse.com/forums/topic/3461-sambuild31-beta-thread/
'// If you don't want to update change UseSolendoids below = 1
'//////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////
'// UseSolenoids = [1 or 2] 1 = Standard flipper latency, 2 = FastFlips

Const UseSolenoids = 2

'//////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////
'Options to disable if your system is not coping
'1=YES/ON     0=NO/OFF

'If you are experiencing performance issues, the first thing you should do is turn off playfield reflections in table options/playfield graphics.
'Next I recommend changing STEELWALLS primitive to static rendering (large texture, but mostly obscured by other things on table)

'PRELOAD ALL THE FLASHERS ON STARTUP
'NOTE: If saucer hesitates on first hit after lowering the bank
'you can use this and might help performance overall slightly.   Causes the table to do a flash sequence at startup.
Const PreloadMe = 1

'OPTIONAL COOL STUFF
'
'If you don't like things that are cool disable flasher effects below
'Might improve performance too :)
' 1 = ON (Default awesome), 0 = OFF (Lame)
' Additional performance tweaks:
'
Const AlienFlashers = 1 'New to v1.1
Const SideWallFlashers = 1 'These add a lot visually but are also taxing
Const ChromeRailFlashers = 1 'New to v1.1.  May be better to change CHROMERAILS primitive to static rendering, will give more performance and improve AA quality (Note: If static though you do not get flashing effects on the rails)
Const UFOFlashers = 1
Const PlasticsTopFlasher = 1 'New to v1.1
Const ApronGIProper = 1 'This would be the first thing to lose if you are trying to add performance as difference is minor.  May also change APRON primitive to static rendering.

'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.8
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

'----- Phsyics Mods -----
Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets, 2 = orig TargetBouncer
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7 when TargetBouncerEnabled=1, and 1.1 when TargetBouncerEnabled=2


'//////////////////////////////////////////////////////////////////////

Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0
Const cSingleLFlip = 0
Const cSingleRFlip = 0
Const SCoin = "fx_Coin"
Const UseVPMModSol = 1


'//////////////////////////////////////////////////////////////////////
'// DOF CONTROLLER INIT
'//////////////////////////////////////////////////////////////////////

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'//////////////////////////////////////////////////////////////////////
'// VARIABLES
'//////////////////////////////////////////////////////////////////////

Dim bsTrough, plungerIM, Ballspeed1, bsL, bsR, aBall, bBall, dtDrop, currGITop, currGIMid, currGIBot
Dim DesktopMode: DesktopMode = AFM.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode     'shows VPX internal DMD


'//////////////////////////////////////////////////////////////////////
'// VPM INIT
'//////////////////////////////////////////////////////////////////////

Const BallSize = 50
Const Ballmass = 1

LoadVPM "01560000", "WPC.VBS", 3.46

Sub LoadVPM(VPMver, VBSfile, VBSver)
On Error Resume Next
If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
ExecuteGlobal GetTextFile(VBSfile)
If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
Set Controller = CreateObject("VPinMAME.Controller")
If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
On Error Goto 0

End Sub



Set GiCallback2 = GetRef("UpdateGI")

'//////////////////////////////////////////////////////////////////////
'// GAME NAME
'//////////////////////////////////////////////////////////////////////

Const cGameName = "afm_113b"


'//////////////////////////////////////////////////////////////////////
'// STARTUP SETTINGS
'//////////////////////////////////////////////////////////////////////

SW45.isdropped=1:SW46.isdropped=1:SW47.isdropped=1

RightSlingRubber_A.visible=0
LeftSlingRubber_A.visible=0
SLINGL.visible=0
SLINGR.visible=0





If ApronGIProper = 1 then APRON_GIoff.visible = 1 else APRON_GIoff.visible = 0

'Show Desktop components and adjust sidewalls
  If DesktopMode Then
  Ufo.Z=20
  LED1.Z=20
LED10.Z=20
LED11.Z=20
LED12.Z=20
LED13.Z=20
LED14.Z=20
LED15.Z=20
LED16.Z=20
LED2.Z=20
LED3.Z=20
LED4.Z=20
LED5.Z=20
LED6.Z=20
LED7.Z=20
LED8.Z=20
LED9.Z=20
METALBRUSHED.Z=20
'   SIDERAILS_B.visible=1
'   CABINETSIDES.SIZE_Y=21
'   CABINETSIDES.X=0
'   CABINETSIDES.Z=0
'   CABINETSIDES_GIOFF.SIZE_Y=21
'   CABINETSIDES_GIOFF.X=0
'   CABINETSIDES_GIOFF.Z=0

  Else
    SIDERAILS_B.visible=0
  End if

CENTREBANK.Z=-76:SW45P.Z=-76:SW46P.Z=-76:SW47P.Z=-76

Diverter8.isdropped=1:Diverter7.isdropped=1:Diverter6.isdropped=1:Diverter5.isdropped=1

Diverter4.isdropped=1:Diverter3.isdropped=1:Diverter2.isdropped=1:Diverter1.isdropped=1

DivPos=0



'//////////////////////////////////////////////////////////////////////
'// TABLE INIT
'//////////////////////////////////////////////////////////////////////


Sub AFM_Init
vpmInit Me
With Controller
.GameName = cGameName
If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
.SplashInfoLine = "Attack from Mars" & vbNewLine & "g5k Edition"
.Games(cGameName).Settings.Value("rol") = 0
.HandleKeyboard = 0
.ShowTitle = 0
.ShowDMDOnly = 1
.ShowFrame = 0
.HandleMechanics = 0
.Hidden = DesktopMode
On Error Resume Next
.Run GetPlayerHWnd
If Err Then MsgBox Err.Description
On Error Goto 0
.Switch(22) = 1 'close coin door
.Switch(24) = 1 'and keep it close
End With

'//////////////////////////////////////////////////////////////////////
'// PINMAME TIMER
'//////////////////////////////////////////////////////////////////////

PinMAMETimer.Interval = PinMAMEInterval
PinMAMETimer.Enabled = 1

'//////////////////////////////////////////////////////////////////////
'// BALL TROUGH
'//////////////////////////////////////////////////////////////////////

Set bsTrough = New cvpmTrough
    With bsTrough
        .size = 4
        .initSwitches Array(32, 33, 34, 35)
        .Initexit BallRelease, 90, 4
        .InitExitSounds SoundFX("BallRelease",DOFContactors), SoundFX("BallRelease",DOFContactors)
        .Balls = 4
    End With


'//////////////////////////////////////////////////////////////////////
'// IMPULSE PLUNGER
'//////////////////////////////////////////////////////////////////////

Const IMPowerSetting = 60 'Plunger Power
Const IMTime = 0.6        ' Time in seconds for Full Plunge
Set plungerIM = New cvpmImpulseP
With plungerIM
.InitImpulseP swplunger, IMPowerSetting, IMTime
.Random 3
.switch 18
.InitExitSnd SoundFX("ShooterLane",DOFContactors), ""
.CreateEvents "plungerIM"
End With
TBPos=28:TBTimer.Enabled=0:TBDown=1:Controller.Switch(66) = 1:Controller.Switch(67) = 0


'//////////////////////////////////////////////////////////////////////
'// VUK
'//////////////////////////////////////////////////////////////////////

  ' Left hole
    Set bsL = New cvpmTrough
    With bsL
        .size = 1
        .initSwitches Array(36)
        .Initexit sw36, 180, 0
        .InitExitSounds "", SoundFX("VUKOut",DOFContactors)
        .InitExitVariance 3, 2
    End With

' Set bsL = New cvpmBallStack
' With bsL
' .InitSw 0, 36, 0, 0, 0, 0, 0, 0
' .InitKick sw36, 180, 0
' .InitExitSnd SoundFX("VUKOut",DOFContactors), ""
' .KickForceVar = 3
' End With

'//////////////////////////////////////////////////////////////////////
'// SCOOP
'//////////////////////////////////////////////////////////////////////

 Set bsR = New cvpmTrough
    With bsR
        .size = 4
        .initSwitches Array(37)
        .Initexit sw37, 198, 25
        .InitExitSounds "", SoundFX("ScoopExit",DOFContactors)
        .InitExitVariance 2, 2
        .MaxBallsPerKick = 1
    '.KickZ = 0.4
    '.KickAngleVar = 1
    End With

' Set bsR = New cvpmBallStack
' With bsR
' .InitSw 0, 37, 0, 0, 0, 0, 0, 0
' .InitKick SW37, 198, 30
' .KickZ = 0.4
' .InitExitSnd SoundFx("ScoopExit",DOFContactors), ""
' .KickAngleVar = 1
' .KickBalls = 1
' End With

'//////////////////////////////////////////////////////////////////////
'// DROPTARGET
'//////////////////////////////////////////////////////////////////////
 Set dtDrop = New cvpmDropTarget
 With dtDrop
 .InitDrop sw77, 77
 .initsnd "", ""
 End With

'//////////////////////////////////////////////////////////////////////
'// NUDGING
'//////////////////////////////////////////////////////////////////////

vpmNudge.TiltSwitch = 14
vpmNudge.Sensitivity = 5.35
vpmNudge.TiltObj = Array(LeftBumper, BottomBumper, RightBumper, LeftSlingshot, RightSlingshot)

'//////////////////////////////////////////////////////////////////////
'// UPDATE GI
'//////////////////////////////////////////////////////////////////////

UpdateGI 0, 1:UpdateGI 1, 1:UpdateGI 2, 1



'//////////////////////////////////////////////////////////////////////
'// END OF TABLE INIT
'//////////////////////////////////////////////////////////////////////
End Sub




'//////////////////////////////////////////////////////////////////////
'// CONTROLLER PAUSE
'//////////////////////////////////////////////////////////////////////

Sub AFM_Paused:Controller.Pause = 1:End Sub
Sub AFM_unPaused:Controller.Pause = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// DRAIN HIT
'//////////////////////////////////////////////////////////////////////

Sub Drain_Hit:RandomSoundDrain Drain:bsTrough.AddBall Me:End Sub
Sub Drain1_Hit:RandomSoundDrain Drain:bsTrough.AddBall Me:End Sub
Sub Drain2_Hit:RandomSoundDrain Drain:bsTrough.AddBall Me:End Sub

'//////////////////////////////////////////////////////////////////////
'// VUK HIT
'//////////////////////////////////////////////////////////////////////

'Sub sw36a_Hit:PlaySoundAtVol "VUKEnter",sw36a,1:bsL.AddBall Me:End Sub
Sub sw36a_Hit:SoundSaucerLock:bsL.AddBall Me:End Sub

'//////////////////////////////////////////////////////////////////////
'// SCOOP HIT
'//////////////////////////////////////////////////////////////////////

'Sub sw37a_Hit:PlaySoundAtVol "ScoopBack",sw37a,1:bsR.AddBall Me:End Sub
Sub sw37a_Hit:SoundSaucerLock:bsR.AddBall Me:End Sub

Sub sw37_Hit
'PlaySoundAtVol "ScoopFront",sw37,1
SoundSaucerLock
Set bBall = ActiveBall:Me.TimerEnabled =1
bsR.AddBall 0
End Sub

Sub sw37_Timer
Do While bBall.Z >0
bBall.Z = bBall.Z -5
Exit Sub
Loop
Me.DestroyBall
Me.TimerEnabled = 0
End Sub

'//////////////////////////////////////////////////////////////////////
'// SAUCER DROPHOLE HIT
'//////////////////////////////////////////////////////////////////////

'Sub sw78_Hit:PlaySoundAtVol "Drophole",sw78,1
Sub sw78_Hit: SoundSaucerLock
  Set aBall = ActiveBall
  Me.TimerEnabled =1
  vpmTimer.PulseSwitch(78), 150, "bsl.addball 0 '"
End Sub

Sub sw78_Timer
  Do While aBall.Z >0
    aBall.Z = aBall.Z -5
    Exit Sub
  Loop
  Me.DestroyBall
  Me.TimerEnabled = 0
End Sub

'//////////////////////////////////////////////////////////////////////
'// SAUCER TARGET HIT
'//////////////////////////////////////////////////////////////////////

Sub sw77_Hit:dtDrop.Hit 1:End Sub

'//////////////////////////////////////////////////////////////////////
'// KEYS
'//////////////////////////////////////////////////////////////////////

Sub AFM_KeyDown(ByVal Keycode)
  If keycode = PlungerKey Then Controller.Switch(11) = 1:End If
  If keycode = LockbarKey Then Controller.Switch(11) = 1:End If

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

  if keycode=StartGameKey then soundStartButton()

  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
  End If

  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
  End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub AFM_KeyUp(ByVal Keycode)
  If keycode = PlungerKey Then Controller.Switch(11) = 0
  If keycode = LockbarKey Then Controller.Switch(11) = 0
  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
  End If

  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
  End If
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'//////////////////////////////////////////////////////////////////////
'// SOLENOID CALLS
'//////////////////////////////////////////////////////////////////////

SolCallback(1) = "Auto_Plunger"
SolCallback(2) = "SolRelease"
SolCallback(3) = "bsL.SolOut"
SolCallback(4) = "bsR.SolOut"
SolCallBack(5) = "Alien1Move"
SolCallBack(6) = "Alien2Move"
SolCallback(7) = "SolKnocker"
SolCallBack(8) = "Alien3Move"
SolCallBack(14) = "Alien4Move"
SolCallBack(15) = "SolUfoShake"
SolCallback(16) = "ResetUFODrop"
SolModCallback(17) = "setmodlamp 117,"
SolModCallback(18) = "setmodlamp 118,"
SolModCallback(19) = "setmodlamp 119,"
SolModCallBack(20) = "setmodlamp 120,"
SolModCallback(21) = "setmodlamp 121,"
SolModCallback(22) = "setmodlamp 122,"
SolModCallback(23) = "setmodlamp 123,"
SolCallBack(24) = "TBMove"
SolModCallBack(25) = "setmodlamp 125,"
SolModCallBack(26) = "setmodlamp 126,"
SolModCallBack(27) = "setmodlamp 127,"
SolModCallBack(28) = "setmodlamp 128,"
SolCallback(33) = "vpmSolGate RGate,false,"
SolCallback(34) = "vpmSolGate LGate,false,"
SolCallback(36) = "DivMove"
SolModCallback(39) = "setmodlamp 123,"
SolModCallback(43) = "setmodlamp 130,"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'//////////////////////////////////////////////////////////////////////
'// FLIPPERS
'//////////////////////////////////////////////////////////////////////

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If UseSolenoids = 2 then Controller.Switch(swULFlip)=Enabled
  If Enabled Then
    LF.Fire 'LeftFlipper.RotateToEnd

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
  if UseSolenoids = 2 then Controller.Switch(swURFlip)=Enabled
  If Enabled Then
    RF.Fire 'RightFlipper.RotateToEnd

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

'//////////////////////////////////////////////////////////////////////
'// SOLENOID SUBS
'//////////////////////////////////////////////////////////////////////

Sub Auto_Plunger(Enabled)
  If Enabled Then
  PlungerIM.AutoFire
  End If
End Sub

Sub SolRelease(Enabled)
  If Enabled And bsTrough.Balls> 0 Then
    vpmTimer.PulseSw 31
    bsTrough.ExitSol_On
  End If
End Sub


Sub SolKnocker(Enabled)
    If enabled Then
        KnockerSolenoid
    End If
End Sub

'//////////////////////////////////////////////////////////////////////
'// POP BUMPERS
'//////////////////////////////////////////////////////////////////////

Sub LeftBumper_Hit:vpmTimer.PulseSw 53:RandomSoundBumperTop LeftBumper
End Sub

Sub BottomBumper_Hit:vpmTimer.PulseSw 54:RandomSoundBumperMiddle BottomBumper
End Sub

Sub RightBumper_Hit:vpmTimer.PulseSw 55:RandomSoundBumperBottom RightBumper
End Sub

'//////////////////////////////////////////////////////////////////////
'// SLINGSHOTS
'//////////////////////////////////////////////////////////////////////

Sub LeftSlingShot_Slingshot:LeftSlingRubber_A.visible = 1:SLINGL.visible = 1:RandomSoundSlingshotLeft SLINGL:vpmTimer.PulseSw 51:LeftSlingshot.timerenabled = 1:End Sub
Sub LeftSlingshot_timer:LeftSlingRubber_A.visible = 0:SLINGL.visible = 0:LeftSlingshot.timerenabled= 0:End Sub

Sub RightSlingShot_Slingshot:RightSlingRubber_A.visible = 1:SLINGR.visible = 1:RandomSoundSlingshotRight SLINGR:vpmTimer.PulseSw 52:Rightslingshot.timerenabled = 1:End Sub
Sub Rightslingshot_timer:RightSlingRubber_A.visible = 0:SLINGR.visible = 0:Rightslingshot.timerenabled= 0:End Sub



'//////////////////////////////////////////////////////////////////////
'// MARTIAN TARGETS
'//////////////////////////////////////////////////////////////////////


'// "M"ARTIAN TARGET //////////////////////////////////////////////////

Sub SW56_Hit
  If SW56P.ObjRotY = 0 Then
    vpmTimer.PulseSw 56:SW56P.ObjRotY=1:SW56P1.ObjRotY=1:Me.TimerEnabled=1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
End Sub

Sub SW56_Timer:SW56P.ObjRotY=0:SW56P1.ObjRotY=0:Me.TimerEnabled=0:End Sub


'// M"A"RTIAN TARGET //////////////////////////////////////////////////

Sub SW57_Hit
  If SW57P.ObjRotY = 0 Then
    vpmTimer.PulseSw 57:SW57P.ObjRotY=1:SW57P1.ObjRotY=1:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
End Sub

Sub SW57_Timer:SW57P.ObjRotY=0:SW57P1.ObjRotY=0:Me.TimerEnabled=0:End Sub

'// MA"R"TIAN TARGET //////////////////////////////////////////////////

Sub SW58_Hit
  If SW58P.ObjRotY = 0 Then
    vpmTimer.PulseSw 58:SW58P.ObjRotY=-1:SW58P1.ObjRotY=-1:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
End Sub

Sub SW58_Timer:SW58P.ObjRotY=0:SW58P1.ObjRotY=0:Me.TimerEnabled=0:End Sub


'// MAR"T"IAN TARGET //////////////////////////////////////////////////


Sub SW43_Hit
  If SW43P.ObjRotX = 0 Then
    vpmTimer.PulseSw 43:SW43P.ObjRotX=0.35:SW43P1.ObjRotX=0.35:Me.TimerEnabled=1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
End Sub

Sub SW43_Timer:SW43P.ObjRotX=0:SW43P1.ObjRotX=0:Me.TimerEnabled=0:End Sub


'// MART"I"AN TARGET //////////////////////////////////////////////////

Sub SW44_Hit
  If SW44P.ObjRotY = 0 Then
    vpmTimer.PulseSw 44:SW44P.ObjRotX = 0.35:SW44P1.ObjRotX = 0.35:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
End Sub

Sub SW44_Timer:SW44P.ObjRotX = 0:SW44P1.ObjRotX = 0:Me.TimerEnabled = 0:End Sub



'// MARTI"A"N TARGET //////////////////////////////////////////////////

Sub SW41_Hit
  If SW41P.ObjRotY = 0 Then
    vpmTimer.PulseSw 41:SW41P.ObjRotY = 1:SW41P1.ObjRotY = 1:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
End Sub

Sub SW41_Timer:SW41P.ObjRotY = 0:SW41P1.ObjRotY = 0:Me.TimerEnabled = 0:End Sub



'// MARTIA"N" TARGET //////////////////////////////////////////////////

Sub SW42_Hit
  If SW42P.ObjRotY = 0 Then
    vpmTimer.PulseSw 42:SW42P.ObjRotY = 1:SW42P1.ObjRotY = 1:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
End Sub

Sub SW42_Timer:SW42P.ObjRotY = 0:SW42P1.ObjRotY = 0:Me.TimerEnabled = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// SAUCER TARGETS
'//////////////////////////////////////////////////////////////////////

Sub SW75_Hit
  If SW75P.ObjRotY = 0 Then
    vpmTimer.PulseSw 75:SW75P.ObjRotY = 1:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
End Sub

Sub SW75_Timer:SW75P.ObjRotY = 0:Me.TimerEnabled = 0:End Sub

Sub SW76_Hit
  If SW76P.ObjRotY = 0 Then
    vpmTimer.PulseSw 76:SW76P.ObjRotY = 1:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
End Sub

Sub SW76_Timer:SW76P.ObjRotY = 0:Me.TimerEnabled = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// TARGETBANK TARGETS
'//////////////////////////////////////////////////////////////////////

Sub SW45_Hit
  If SW45P.ObjRotX = 0 Then
    vpmTimer.PulseSw 45:SW45P.ObjRotX = 0.35:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
  If Ballresting = True Then
    DPBall.VelY = ActiveBall.VelY * 3
  End If
End Sub

Sub SW45_Timer:SW45P.ObjRotX = 0:Me.TimerEnabled = 0:End Sub

Sub SW46_Hit
  If SW46P.ObjRotX = 0 Then
    vpmTimer.PulseSw 46:SW46P.ObjRotX = 0.35:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
  If Ballresting = True Then
    DPBall.VelY = ActiveBall.VelY * 3
  End If
End Sub

Sub SW46_Timer:SW46P.ObjRotX = 0:Me.TimerEnabled = 0:End Sub

Sub SW47_Hit
  If SW47P.ObjRotX = 0 Then
    vpmTimer.PulseSw 47:SW47P.ObjRotX = 0.35:Me.TimerEnabled = 1
    'PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  End If
  If Ballresting = True Then
    DPBall.VelY = ActiveBall.VelY * 3
  End If
End Sub

Sub SW47_Timer:SW47P.ObjRotX = 0:Me.TimerEnabled = 0:End Sub


Sub sw77_Dropped:dtDrop.Hit 1:End Sub

Sub ResetUFODrop(enabled)
  dtDrop.SolDropUp enabled
  if Enabled then sw77.visible = True
End Sub



'//////////////////////////////////////////////////////////////////////
'// RAMP SWITCHES
'//////////////////////////////////////////////////////////////////////

Sub SW61_Hit:Controller.Switch(61) = 1:WireRampOn True, "LeftRamp Plastic":End Sub
Sub SW61_UnHit:Controller.Switch(61) = 0:End Sub

Sub SW62_Hit:Controller.Switch(62) = 1:WireRampOn True, "MiddleRamp Plastic":End Sub
Sub SW62_UnHit:Controller.Switch(62) = 0:End Sub

Sub SW63_Hit:Controller.Switch(63) = 1:WireRampOn True, "RightRamp Plastic":End Sub
Sub SW63_UnHit:Controller.Switch(63) = 0:End Sub

Sub SW64_Hit:Controller.Switch(64) = 1:End Sub
Sub SW64_UnHit:Controller.Switch(64) = 0:End Sub

Sub SW65_Hit
  Controller.Switch(65) = 1
  WireRampOff
  WireRampOnSndX False, 1.5, "RightRamp Wire"
  'PlaySound "fx_metalrolling",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
End Sub
Sub SW65_UnHit:Controller.Switch(65) = 0:End Sub


'//////////////////////////////////////////////////////////////////////
'// ROLLOVER SWITCHES
'//////////////////////////////////////////////////////////////////////

Sub SW16_Hit:Controller.Switch(16) = 1:End Sub
Sub SW16_UnHit:Controller.Switch(16) = 0:End Sub

Sub SW17_Hit:Controller.Switch(17) = 1:End Sub
Sub SW17_UnHit:Controller.Switch(17) = 0:End Sub

Sub SW26_Hit:Controller.Switch(26) = 1:End Sub
Sub SW26_UnHit:Controller.Switch(26) = 0:End Sub

Sub SW27_Hit:Controller.Switch(27) = 1:End Sub
Sub SW27_UnHit:Controller.Switch(27) = 0:End Sub


Sub SW38_Hit:Controller.Switch(38) = 1:End Sub
Sub SW38_UnHit:Controller.Switch(38) = 0:End Sub

Sub SW48_Hit:Controller.Switch(48) = 1:End Sub
Sub SW48_UnHit:Controller.Switch(48) = 0:End Sub

Sub SW71_Hit:Controller.Switch(71) = 1:End Sub
Sub SW71_UnHit:Controller.Switch(71) = 0:End Sub

Sub SW72_Hit:Controller.Switch(72) = 1:End Sub
Sub SW72_UnHit:Controller.Switch(72) = 0:End Sub

Sub SW73_Hit:Controller.Switch(73) = 1:End Sub
Sub SW73_UnHit:Controller.Switch(73) = 0:End Sub

Sub SW74_Hit:Controller.Switch(74) = 1:End Sub
Sub SW74_UnHit:Controller.Switch(74) = 0:End Sub

'//////////////////////////////////////////////////////////////////////
'// SAUCER MOVEMENT
'//////////////////////////////////////////////////////////////////////
Dim cBall, mMagnet

Sub WobbleMagnet_Init
   Set mMagnet = new cvpmMagnet
   With mMagnet
    .InitMagnet WobbleMagnet, 2
    .Size = 100
    .CreateEvents mMagnet
    .MagnetOn = True
   End With
  Set cBall = ckicker.createball:ckicker.Kick 0,0:mMagnet.addball cball

End Sub

Sub SolUfoShake(Enabled)
  If Enabled Then
    BigUfoShake
    PlaySound SoundFX("SaucerShake",DOFShaker),0,2,0,0,0,0,1,-.3
  End If
End Sub

Sub BigUfoShake
cball.velx = 4
cball.vely = -18
End Sub

Sub UfoShaker_Timer
Dim a, b, c
a = 90-((ckicker.y - cball.y) / 5)
b = (ckicker.y - cball.y) / 10
c = (cball.x - ckicker.x) / 5




Ufo.rotx = a:Ufo.transz = b:Ufo.rotz = c
'LED1.rotx = a:LED1.transz = b:LED1.rotz = c
'LED2.rotx = a:LED2.transz = b:LED2.rotz = c
'LED3.rotx = a:LED3.transz = b:LED3.rotz = c
LED4.rotx = a:LED4.transz = b:LED4.rotz = c
LED5.rotx = a:LED5.transz = b:LED5.rotz = c
LED6.rotx = a:LED6.transz = b:LED6.rotz = c
LED7.rotx = a:LED7.transz = b:LED7.rotz = c
LED8.rotx = a:LED8.transz = b:LED8.rotz = c
LED9.rotx = a:LED9.transz = b:LED9.rotz = c
LED10.rotx = a:LED10.transz = b:LED10.rotz = c
LED11.rotx = a:LED11.transz = b:LED11.rotz = c
LED12.rotx = a:LED12.transz = b:LED12.rotz = c
LED13.rotx = a:LED13.transz = b:LED13.rotz = c
'LED14.rotx = a:LED14.transz = b:LED14.rotz = c
'LED15.rotx = a:LED15.transz = b:LED15.rotz = c
'LED16.rotx = a:LED16.transz = b:LED16.rotz = c
End Sub

'//////////////////////////////////////////////////////////////////////
'// ALIEN MOVEMENTS
'//////////////////////////////////////////////////////////////////////

Dim Alien1Pos

Sub Alien1Move (enabled)
if enabled then
Alien1Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake1",DOFShaker),Licht14,2
End If
End Sub

Sub Alien1Timer_Timer()
Select Case Alien1Pos
Case 0: Alien_LF.Z=0:AlienPost1.Z=-60
Case 1: Alien_LF.Z=5:AlienPost1.Z=-55
Case 2: Alien_LF.Z=10:AlienPost1.Z=-50
Case 3: Alien_LF.Z=15:AlienPost1.Z=-45
Case 4: Alien_LF.Z=20:AlienPost1.Z=-40
Case 5: Alien_LF.Z=25:AlienPost1.Z=-35
Case 6: Alien_LF.Z=30:AlienPost1.Z=-30
Case 7: Alien_LF.Z=35:AlienPost1.Z=-25
Case 8: Alien_LF.Z=40:AlienPost1.Z=-20
Case 9: Alien_LF.Z=35:AlienPost1.Z=-25
Case 10: Alien_LF.Z=30:AlienPost1.Z=-30
Case 11: Alien_LF.Z=25:AlienPost1.Z=-35
Case 12: Alien_LF.Z=20:AlienPost1.Z=-40
Case 13: Alien_LF.Z=15:AlienPost1.Z=-45
Case 14: Alien_LF.Z=10:AlienPost1.Z=-50
Case 15: Alien_LF.Z=5:AlienPost1.Z=-55
Case 16: Alien_LF.Z=0:AlienPost1.Z=-60:Alien1Pos=0:Alien1Timer.Enabled=0
End Select

If Alien1Pos=>0 then Alien1Pos=Alien1Pos+1
End Sub


Dim Alien2Pos

Sub Alien2Move (enabled)
if enabled then
Alien2Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake2",DOFShaker),f28a,2
End If
End Sub

Sub Alien2Timer_Timer()
Select Case Alien2Pos
Case 0: Alien_LR.Z=0:AlienPost2.Z=-60
Case 1: Alien_LR.Z=5:AlienPost2.Z=-55
Case 2: Alien_LR.Z=10:AlienPost2.Z=-50
Case 3: Alien_LR.Z=15:AlienPost2.Z=-45
Case 4: Alien_LR.Z=20:AlienPost2.Z=-40
Case 5: Alien_LR.Z=25:AlienPost2.Z=-35
Case 6: Alien_LR.Z=30:AlienPost2.Z=-30
Case 7: Alien_LR.Z=35:AlienPost2.Z=-25
Case 8: Alien_LR.Z=40:AlienPost2.Z=-20
Case 9: Alien_LR.Z=35:AlienPost2.Z=-25
Case 10: Alien_LR.Z=30:AlienPost2.Z=-30
Case 11: Alien_LR.Z=25:AlienPost2.Z=-35
Case 12: Alien_LR.Z=20:AlienPost2.Z=-40
Case 13: Alien_LR.Z=15:AlienPost2.Z=-45
Case 14: Alien_LR.Z=10:AlienPost2.Z=-50
Case 15: Alien_LR.Z=5:AlienPost2.Z=-55
case 16: Alien_LR.Z=0:AlienPost2.Z=-60:Alien2Pos=1:Alien2Timer.Enabled=0
End Select

If Alien2Pos=>0 then Alien2Pos=Alien2Pos+1
End Sub

Dim Alien3Pos

Sub Alien3Move (enabled)
if enabled then
Alien3Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake3",DOFShaker),f28a,2
End If
End Sub

Sub Alien3Timer_Timer()
Select Case Alien3Pos
Case 0: Alien_RR.Z=0:AlienPost3.Z=-60
Case 1: Alien_RR.Z=5:AlienPost3.Z=-55
Case 2: Alien_RR.Z=10:AlienPost3.Z=-50
Case 3: Alien_RR.Z=15:AlienPost3.Z=-45
Case 4: Alien_RR.Z=20:AlienPost3.Z=-40
Case 5: Alien_RR.Z=25:AlienPost3.Z=-35
Case 6: Alien_RR.Z=30:AlienPost3.Z=-30
Case 7: Alien_RR.Z=35:AlienPost3.Z=-25
Case 8: Alien_RR.Z=40:AlienPost3.Z=-20
Case 9: Alien_RR.Z=35:AlienPost3.Z=-25
Case 10: Alien_RR.Z=30:AlienPost3.Z=-30
Case 11: Alien_RR.Z=25:AlienPost3.Z=-35
Case 12: Alien_RR.Z=20:AlienPost3.Z=-40
Case 13: Alien_RR.Z=15:AlienPost3.Z=-45
Case 14: Alien_RR.Z=10:AlienPost3.Z=-50
Case 15: Alien_RR.Z=5:AlienPost3.Z=-55
Case 16: Alien_RR.Z=0:AlienPost3.Z=-60:Alien3Pos=1:Alien3Timer.Enabled=0

End Select


If Alien3Pos=>0 then Alien3Pos=Alien3Pos+1
End Sub


Dim Alien4Pos

Sub Alien4Move (enabled)
if enabled then
Alien4Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake4",DOFShaker),Licht9,2
End If
End Sub

Sub Alien4Timer_Timer()
Select Case Alien4Pos
Case 0: Alien_RF.Z=0:AlienPost4.Z=-60
Case 1: Alien_RF.Z=5:AlienPost4.Z=-55
Case 2: Alien_RF.Z=10:AlienPost4.Z=-50
Case 3: Alien_RF.Z=15:AlienPost4.Z=-45
Case 4: Alien_RF.Z=20:AlienPost4.Z=-40
Case 5: Alien_RF.Z=25:AlienPost4.Z=-35
Case 6: Alien_RF.Z=30:AlienPost4.Z=-30
Case 7: Alien_RF.Z=35:AlienPost4.Z=-25
Case 8: Alien_RF.Z=40:AlienPost4.Z=-20
Case 9: Alien_RF.Z=35:AlienPost4.Z=-25
Case 10: Alien_RF.Z=30:AlienPost4.Z=-30
Case 11: Alien_RF.Z=25:AlienPost4.Z=-35
Case 12: Alien_RF.Z=20:AlienPost4.Z=-40
Case 13: Alien_RF.Z=15:AlienPost4.Z=-45
Case 14: Alien_RF.Z=10:AlienPost4.Z=-50
Case 15: Alien_RF.Z=5:AlienPost4.Z=-55
Case 16: Alien_RF.Z=0:AlienPost4.Z=-60:Alien4Pos=1:Alien4Timer.Enabled=0

End Select

If Alien4Pos=>0 then Alien4Pos=Alien4Pos+1
End Sub

'//////////////////////////////////////////////////////////////////////
'// TARGETBANK MOVEMENT
'//////////////////////////////////////////////////////////////////////

Dim TBPos, TBDown

Sub TBMove (enabled)
  TBTimer.Enabled = Enabled
  if enabled then PlaySoundAt SoundFX("TargetBank",DOFContactors),DPTrigger
End Sub

Sub TBTimer_Timer()
Select Case TBPos
Case 0: CENTREBANK.Z=-20:SW45P.Z=-20:SW46P.Z=-20:SW47P.Z=-20:TBPos=0:Controller.Switch(66) = 0:Controller.Switch(67) = 1::SW45.isdropped=0:SW46.isdropped=0:SW47.isdropped=0:DPWall.isdropped=0:DPWall1.isdropped=1:DPRAMP.collidable=1
Case 1: CENTREBANK.Z=-22:SW45P.Z=-22:SW46P.Z=-22:SW47P.Z=-22:Controller.Switch(67) = 0
Case 2: CENTREBANK.Z=-24:SW45P.Z=-24:SW46P.Z=-24:SW47P.Z=-24
Case 3: CENTREBANK.Z=-26:SW45P.Z=-26:SW46P.Z=-26:SW47P.Z=-26
Case 4: CENTREBANK.Z=-28:SW45P.Z=-28:SW46P.Z=-28:SW47P.Z=-28
Case 5: CENTREBANK.Z=-30:SW45P.Z=-30:SW46P.Z=-30:SW47P.Z=-30
Case 6: CENTREBANK.Z=-32:SW45P.Z=-32:SW46P.Z=-32:SW47P.Z=-32
Case 7: CENTREBANK.Z=-34:SW45P.Z=-34:SW46P.Z=-34:SW47P.Z=-34
Case 8: CENTREBANK.Z=-36:SW45P.Z=-36:SW46P.Z=-36:SW47P.Z=-36
Case 9: CENTREBANK.Z=-38:SW45P.Z=-38:SW46P.Z=-38:SW47P.Z=-38
Case 10: CENTREBANK.Z=-40:SW45P.Z=-40:SW46P.Z=-40:SW47P.Z=-40
Case 11: CENTREBANK.Z=-42:SW45P.Z=-42:SW46P.Z=-42:SW47P.Z=-42
Case 12: CENTREBANK.Z=-44:SW45P.Z=-44:SW46P.Z=-44:SW47P.Z=-44:
Case 13: CENTREBANK.Z=-46:SW45P.Z=-46:SW46P.Z=-46:SW47P.Z=-46:
Case 14: CENTREBANK.Z=-48:SW45P.Z=-48:SW46P.Z=-48:SW47P.Z=-48
Case 15: CENTREBANK.Z=-50:SW45P.Z=-50:SW46P.Z=-50:SW47P.Z=-50
Case 16: CENTREBANK.Z=-52:SW45P.Z=-52:SW46P.Z=-52:SW47P.Z=-52
Case 17: CENTREBANK.Z=-54:SW45P.Z=-54:SW46P.Z=-54:SW47P.Z=-54
Case 18: CENTREBANK.Z=-56:SW45P.Z=-56:SW46P.Z=-56:SW47P.Z=-56
Case 19: CENTREBANK.Z=-58:SW45P.Z=-58:SW46P.Z=-58:SW47P.Z=-58
Case 20: CENTREBANK.Z=-60:SW45P.Z=-60:SW46P.Z=-60:SW47P.Z=-60
Case 21: CENTREBANK.Z=-62:SW45P.Z=-62:SW46P.Z=-62:SW47P.Z=-62
Case 22: CENTREBANK.Z=-64:SW45P.Z=-64:SW46P.Z=-64:SW47P.Z=-64
Case 23: CENTREBANK.Z=-66:SW45P.Z=-66:SW46P.Z=-66:SW47P.Z=-66
Case 24: CENTREBANK.Z=-68:SW45P.Z=-68:SW46P.Z=-68:SW47P.Z=-68
Case 25: CENTREBANK.Z=-70:SW45P.Z=-70:SW46P.Z=-70:SW47P.Z=-70
Case 26: CENTREBANK.Z=-72:SW45P.Z=-72:SW46P.Z=-72:SW47P.Z=-72
Case 27: CENTREBANK.Z=-74:SW45P.Z=-74:SW46P.Z=-74:SW47P.Z=-74
Case 28: CENTREBANK.Z=-76:SW45P.Z=-76:SW46P.Z=-76:SW47P.Z=-76:SW45.isdropped=1:SW46.isdropped=1:SW47.isdropped=1:DPWALL.isdropped=1:DPRAMP.collidable=0:Controller.Switch(66) = 0
Case 29: Controller.Switch(66) = 1:Controller.Switch(67) = 0
End Select

If TBDown=0 then TBPos=TBPos+1
If TBDown=1 then TBPos=TBPos-1
if TBPos < -2 then TBDown = 0 else if TBPos > 31 then TBDown = 1
End Sub

'//////////////////////////////////////////////////////////////////////
'// DIRTY POOL HANDLE
'//////////////////////////////////////////////////////////////////////

Dim DPBall, Ballresting

DPBall = ""
Ballresting = False

Sub DPTrigger_Hit
Ballresting = True
Set DPBall = ActiveBall
End Sub

Sub DPTrigger_UnHit
Ballresting = False
DPRAMP.collidable=0
End Sub

'//////////////////////////////////////////////////////////////////////
'// DIVERTER MOVEMENT
'//////////////////////////////////////////////////////////////////////

Dim DivPos, DivClosed, DiverterDir

Sub DivMove(Enabled)
If Enabled Then
PlaySound SoundFX("Diverter",DOFContactors)
DiverterDir = 1
Else
DiverterDir = -1
End If

DivTimer.Enabled = 0
If DivPos <1 Then DivPos = 1
If DivPos > 8 Then DivPos = 8

DivTimer.Enabled = 1
End Sub

Sub DivTimer_Timer()
Select Case DivPos
Case 0:Diverter.ObjRotZ=-15:Diverter9.IsDropped = 0:Diverter8.IsDropped = 1:DivTimer.Enabled = 0
Case 1:Diverter.ObjRotZ=-8:Diverter8.IsDropped = 0:Diverter9.IsDropped = 1:Diverter7.IsDropped = 1
Case 2:Diverter.ObjRotZ=-3:Diverter7.IsDropped = 0:Diverter8.IsDropped = 1:Diverter6.IsDropped = 1
Case 3:Diverter.ObjRotZ=2:Diverter6.IsDropped = 0:Diverter7.IsDropped = 1:Diverter5.IsDropped = 1
Case 4:Diverter.ObjRotZ=7:Diverter5.IsDropped = 0:Diverter6.IsDropped = 1:Diverter4.IsDropped = 1
Case 5:Diverter.ObjRotZ=12:Diverter4.IsDropped = 0:Diverter5.IsDropped = 1:Diverter3.IsDropped = 1
Case 6:Diverter.ObjRotZ=17:Diverter3.IsDropped = 0:Diverter4.IsDropped = 1:Diverter2.IsDropped = 1
Case 7:Diverter.ObjRotZ=22:Diverter2.IsDropped = 0:Diverter3.IsDropped = 1:Diverter1.IsDropped = 1
Case 8:Diverter.ObjRotZ=27:Diverter1.IsDropped = 0:Diverter2.isdropped = 0
Case 9:DivTimer.Enabled = 0
End Select

DivPos = DivPos + DiverterDir
End Sub


Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub RRTrigger_Hit
  WireRampOff
  RandomSoundRampStop RRTrigger
  RandomSoundBallBouncePlayfieldSoft ActiveBall
End Sub

Sub RDrop_Hit()
  ActiveBall.VelZ = -2
  ActiveBall.VelY = 0
  ActiveBall.VelX = 0
  'StopSound "fx_metalrolling"
  WireRampOff
  RandomSoundRampStop RDrop
  RandomSoundBallBouncePlayfieldSoft ActiveBall
  'PlaySound "Balldrop",0,1,.4,0,0,0,1,.8
End Sub

Sub LRTrigger_Hit
  Debug.Print "LRTrigger"
  WireRampOff
  RandomSoundRampStop LRTrigger
  RandomSoundBallBouncePlayfieldSoft ActiveBall
End Sub

Sub LDrop_Hit()
  ActiveBall.VelZ = -2
  ActiveBall.VelY = 0
  ActiveBall.VelX = 0
  WireRampOff
  RandomSoundRampStop LDrop
  RandomSoundBallBouncePlayfieldSoft ActiveBall
  'StopSound "fx_metalrolling"
  'PlaySound "Balldrop",0,1,-.4,0,0,0,1,.8
End Sub

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*6)
    Case 0: PlaySoundAtVol "wireramp_stop", obj, 0.2*volumedial
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*volumedial
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*volumedial
    Case 3: PlaySoundAtVol "wireramp_stop", obj, 0.1*volumedial
    Case 4: PlaySoundAtVol "wireramp_stop2", obj, 0.1*volumedial
    Case 5: PlaySoundAtVol "wireramp_stop3", obj, 0.1*volumedial
  End Select
End Sub

Sub LRoll_Hit()
  WireRampOff
  WireRampOnSndX False, 1.5, "LeftRamp Wire"
  'PlaySound "fx_metalrolling",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
End Sub

'//////////////////////////////////////////////////////////////////////
'// BALL SOUND FUNCTIONS
'//////////////////////////////////////////////////////////////////////
'
' TODO - Remove Later
Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "AFM" is the name of the table
  Dim tmp
  tmp = ball.x * 2 / AFM.width-1
  If tmp > 0 Then
  Pan = Csng(tmp ^10)
  Else
  Pan = Csng(-((- tmp) ^10) )
  End If
End Function


'//////////////////////////////////////////////////////////////////////
'// VPX ROLLING SOUNDS
'//////////////////////////////////////////////////////////////////////

Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing
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
  Next
End Sub


'//////////////////////////////////////////////////////////////////////
'// GI CONTROL
'//////////////////////////////////////////////////////////////////////


Sub UpdateGI(no, step)
Dim gistep, ii, a
gistep = step / 8
Select Case no
Case 0
if Round(8*gistep) < 1.25 then PLASTICS_Bottom.visible = 0 else PLASTICS_Bottom.visible = 1 end If
PLASTICS_Bottom.material = "GIFade" & Round(8*gistep)
STARPOSTS.material = "GIFade" & Round(8*gistep)

If ApronGIProper = 1 Then
  APRON.material = "GIFade" & Round(8*gistep)
Else
  APRON.material = "GIFadeNoOpacity" & Round(8*gistep)
End If

batleft.material = "FlipperPlastic" & Round(8*gistep)
batright.material = "FlipperPlastic" & Round(8*gistep)
For each ii in GIBottom
ii.IntensityScale = gistep
currGIBot = 8*gistep
Next
Case 1
if Round(8*gistep) < 1.25 then PLASTICS_MidL.visible = 0 else PLASTICS_MidL.visible = 1 end If
if Round(8*gistep) < 1.25 then PLASTICS_MidR.visible = 0 else PLASTICS_MidR.visible = 1 end If
PLASTICS_MidL.material = "GIFade" & Round(8*gistep)
PLASTICS_MidR.material = "GIFade" & Round(8*gistep)
SIDEWALL.material = "GIFade" & Round(8*gistep)
if Round(8*gistep) < 1.25 then SIDEWALL.visible = 0 else SIDEWALL.visible = 1 end If
if Round(8*gistep) >= 7 then SIDEWALL_GIOFF.visible = 0 else SIDEWALL_GIOFF.visible = 1 end If
LEFTPOPPERSCOOP.material = "GIFade" & Round(8*gistep)
SW56P.material = "GIFade" & Round(8*gistep)
SW57P.material = "GIFade" & Round(8*gistep)
SW58P.material = "GIFade" & Round(8*gistep)
SW42P.material = "GIFade" & Round(8*gistep)
SW41P.material = "GIFade" & Round(8*gistep)
SAUCERSML.material = "GIFade" &  Round(8*gistep)
STEELWALLS.material = "GIFadeNoOpacity" & Round(8*gistep)
Diverter.material = "GIFadeNoOpacity" & Round(8*gistep)
For each ii in GIMiddle
ii.IntensityScale = gistep
currGIMid = 8*gistep
Next
Case 2
if Round(8*gistep) < 1.25 then PLASTICS_Top.visible = 0 else PLASTICS_Top.visible = 1 end If
PLASTICS_Top.material = "GIFade" & Round(8*gistep)
CLEARAMPS.material = "ClearRamp1" & Round(8*gistep)

if Round(8*gistep) < 1.1 then
  CLEARAMPS.image = "CLEARAMPS_A"
  Alien_LF.image = "ALIEN_LF_A"
  Alien_LR.image = "ALIEN_LR_A"
  Alien_RF.image = "ALIEN_RF_A"
  Alien_RR.image = "ALIEN_RR_A"
  Ufo.image = "UFO_GIOff"
    else
  CLEARAMPS.image = "CLEARAMPS_B"
  Alien_LF.image = "ALIEN_LF_A"
  Alien_LR.image = "ALIEN_LR_A"
  Alien_RF.image = "ALIEN_RF_A"
  Alien_RR.image = "ALIEN_RR_A"
  Ufo.image = "UFO_RG"
end If

'REARWALL.material = "GIFadeNoOpacity" & Round(8*gistep)
REDPLASTICS.material = "GIFade" & Round(8*gistep)
SW43P.material = "GIFade" & Round(8*gistep)
SW44P.material = "GIFade" & Round(8*gistep)
For each ii in GITop
ii.IntensityScale = gistep
currGITop = 8*gistep
Next
End Select
End Sub


'//////////////////////////////////////////////////////////////////////
'// REALTIME UPDATES
'//////////////////////////////////////////////////////////////////////

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
  RollingUpdate
End Sub

'//////////////////////////////////////////////////////////////////////
'// JP FADING LIGHT SYSTEM
'//////////////////////////////////////////////////////////////////////

Dim LampState(200), FadingLevel(200), ModLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = -1 'lamp fading speed
LampTimer.Enabled = 1



' Lamp & Flasher Timers


'//PRELOADING LIGHTS ON  START HERE////////////////////////////////////

Dim GIInit: GIInit=90

Sub LampTimer_Timer()

If PreloadMe = 1 Then
  if GIInit > 0 Then
    GIInit = GIInit -1
    if GIInit <= 90 and GIInit > 82 then
      UpdateGI 0, GIInit-82
      UpdateGI 1, GIInit-82
      UpdateGI 2, GIInit-82
    Else
      if GIInit < 40 then
        UpdateGI 0, 8:UpdateGI 1, 8:UpdateGI 2, 8
      elseif GIInit >= 40 and GIInit <=82 then
        UpdateGI 0, 0:UpdateGI 1, 0:UpdateGI 2, 0
      end if
    end if
    if (GIInit mod 40) > 1 then SetLamp 90 + (GIInit mod 40) +1, 1
    SetLamp 90 + (GIInit mod 40) + 2, 0
  end if
End If


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

'//////////////////////////////////////////////////////////////////////
'// INSERT LIGHTS
'//////////////////////////////////////////////////////////////////////

Sub UpdateLamps

NFadeLm 11, l11, l11b
NFadeLm 12, l12, l12b
NFadeLm 13, l13, l13b
NFadeLm 14, l14, l14b
NFadeLmm 15, l15l, l15r, l15a
NFadeLm 16, l16, l16b
NFadeLm 17, l17, l17b
NFadeLmm 18, l18, l18b, l18a
NFadeLm 21, l21, l21b
NFadeLm 22, l22, l22b
NFadeLm 23, l23, l23b
NFadeLm 24, l24, l24b
NFadeLmm 25, l25, l25b, l25a
NFadeLm 26, l26, l26b
NFadeLm 27, l27, l27b
NFadeLm 28, l28, l28b
NFadeLm 31, l31, l31b
NFadeLm 32, l32, l32b
NFadeLm 33, l33, l33b
NFadeLm 34, l34, l34b
NFadeLm 35, l35, l35b
NFadeLm 36, l36, l36b
NFadeLm 37, l37, l37b
NFadeLm 38, l38, l38b
NFadeLmm 41, l41, l41b, l41a
NFadeLmm 42, l42, l42b, l42a
NFadeLmm 43, l43, l43b, l43a
NFadeLmm 44, l44, l44b, l44a
NFadeLmm 45, l45, l45b, l45a
NFadeLmm 46, l46, l46b, l46a
NFadeLmm 47, l47, l47b, l47a
NFadeLmm 48, l48, l48b, l48a
NFadeLm 51, l51, l51b
NFadeLm 52, l52, l52b
NFadeLm 53, l53, l53b
NFadeLm 54, l54, l54b
NFadeLm 55, l55, l55b
NFadeLm 56, l56, l56b
NFadeLm 57, l57, l57b
NFadeLm 58, l58, l58b
NFadeLm 61, l61, l61b
NFadeLm 62, l62, l62b
NFadeLm 63, l63, l63b
NFadeLm 64, l64, l64b
NFadeLm 65, l65, l65b
NFadeLm 66, l66, l66b
NFadeLm 67, l67, l67b
NFadeLmm 68, l68, l68b, l68a
NFadeLmm 71, l71, l71b, l71a
NFadeLm 72, l72, l72b
NFadeLm 73, l73, l73b
NFadeLm 74, l74, l74b
NFadeLmm 75, l75, l75b, l75a
NFadeLm 76, l76, l76b
NFadeLm 77, l77, l77b
NFadeLm 78, l78, l78b
NFadeLmm 81, l81, l81b, l81a
NFadeLmm 82, l82, l82b, l82a
NFadeLmm 83, l83, l83b, l83a
NFadeLmm 84, l84, l84b, l84a
NFadeLmm 85, l85, l85b, l85a

'86 - Launch button LED
'88 - Start button LED

'NFadeL 86, l86
'NFadeL 88, l88

'//////////////////////////////////////////////////////////////////////
'// FLASHER
'//////////////////////////////////////////////////////////////////////

NFadeLmf 117, F17dr, F17flare
NFadeLmf 118, F18dr, F18flare
NFadeLmf 119, F19dr, F19flare
NFadeLmf 125, f25dr, F25flare
NFadeLmf 126, F26dr, F26flare
NFadeLmf 127, f27dr, F27flare





'NFadeLmm 117, F17, F17b, F17c
'NFadeLmm 118, F18, F18b, F18c
'NFadeLmm 119, F19, F19b, F19c


NFadeLmf 120, f20, f20a
NFadeLm 121, F21, F21a
NFadeLm 122, f22, f22b
NFadeLmf 123, f23, f23c
'NFadeLmm 125, F25, F25b, F25c
'NFadeLmm 126, F26, F26b, F26c
'NFadeLmm 127, F27, F27b, F27c
NFadeLmf 128, f28, f28a




'//////////////////////////////////////////////////////////////////////
'// SAUCER LED
'//////////////////////////////////////////////////////////////////////


FadeObj 94, Led4, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 95, Led5, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 96, Led6, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 97, Led7, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 98, Led8, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 101, Led9, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 102, Led10, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 103, Led11, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 104, Led12, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
FadeObj 105, Led13, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"


'// These LEDS are round back so not visibile - Off for performance ///////
'FadeObj 91, Led1, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
'FadeObj 92, Led2, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
'FadeObj 93, Led3, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"

'FadeObj 106, Led14, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
'FadeObj 107, Led15, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"
'FadeObj 108, Led16, "UFOLED_A", "UFOLED_B", "UFOLED_C", "UFOLED_D"




'///////////////////////////////////////////////////////////////////////////
'// FLASHERS - 20 = LeftFlasher 23 = Saucer Flasher 28 = Right Flasher 39 = Strobe
'///////////////////////////////////////////////////////////////////////////


If UFOFlashers =1 Then
  NFadeObjm 120, Ufo, "UFO_RF", "UFO_RG"
  NFadeObjm 123, Ufo, "UFO_Flasher_D", "UFO_RG"
  NFadeObjm 128, Ufo, "UFO_LF", "UFO_RG"
End If

If ChromeRailFlashers = 1 Then

' FLASHERS / STROBE - switchback to GI off texture
  NFadeObjm 120, CHROMERAILS, "CHROMERAILS_D", "CHROMERAILS_A"
  NFadeObjm 123, CHROMERAILS, "CHROMERAILS_C", "CHROMERAILS_A"
  NFadeObjm 128, CHROMERAILS, "CHROMERAILS_E", "CHROMERAILS_A"

' REDFLASHERS - switchback to GI on texture
  NFadeObjm 125, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
  NFadeObjm 126, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
  NFadeObjm 127, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
  NFadeObjm 117, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
  NFadeObjm 119, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"
  NFadeObjm 118, CHROMERAILS, "CHROMERAILS_F", "CHROMERAILS_B"


  NFadeObjm 120, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_C", "LEFTPOPPERSCOOP_A"
  NFadeObjm 123, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_C", "LEFTPOPPERSCOOP_A"
  NFadeObjm 128, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_C", "LEFTPOPPERSCOOP_A"
  NFadeObjm 125, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
  NFadeObjm 126, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
  NFadeObjm 127, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
  NFadeObjm 117, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
  NFadeObjm 119, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"
  NFadeObjm 118, LEFTPOPPERSCOOP_GIOFF, "LEFTPOPPERSCOOP_F", "LEFTPOPPERSCOOP_A"

  NFadeObjm 120, METALBRUSHED, "METALBRUSHED_C", "METALBRUSHED_B"
  NFadeObjm 123, METALBRUSHED, "METALBRUSHED_C", "METALBRUSHED_B"
  NFadeObjm 128, METALBRUSHED, "METALBRUSHED_C", "METALBRUSHED_B"
  NFadeObjm 125, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
  NFadeObjm 126, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
  NFadeObjm 127, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
  NFadeObjm 117, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
  NFadeObjm 119, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"
  NFadeObjm 118, METALBRUSHED, "METALBRUSHED_F", "METALBRUSHED_B"

  NFadeObjm 120, BLACK_PAINTED, "BLACK_PAINTED_C", "BLACK_PAINTED_A"
  NFadeObjm 123, BLACK_PAINTED, "BLACK_PAINTED_C", "BLACK_PAINTED_A"
  NFadeObjm 128, BLACK_PAINTED, "BLACK_PAINTED_C", "BLACK_PAINTED_A"
  NFadeObjm 125, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
  NFadeObjm 126, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
  NFadeObjm 127, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
  NFadeObjm 117, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
  NFadeObjm 119, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"
  NFadeObjm 118, BLACK_PAINTED, "BLACK_PAINTED_F", "BLACK_PAINTED_A"


End If

If PlasticsTopFlasher = 1 Then
  NFadeObjm 123, PLASTICS_Top, "PLASTICS_C", "PLASTICS_B"
  NFadeObjm 123, PLASTICS_Top_GIOFF, "PLASTICS_C", "PLASTICS_A"
  NFadeObjm 123, REDPLASTICS_GIOFF, "RED_PLASTICS_FLASHERS", "REDPLASTICS_A"


If currGITop <4 Then
NFadeObjm 123, CLEARAMPS, "CLEARAMPS_C", "CLEARAMPS_A"
Else
NFadeObjm 123, CLEARAMPS, "CLEARAMPS_C", "CLEARAMPS_B"
End If

End If

If AlienFlashers = 1 Then
  NFadeObjm 120, Alien_LF, "ALIEN_LF_D", "ALIEN_LF_A"
  NFadeObjm 120, Alien_RF, "ALIEN_RF_D", "ALIEN_RF_A"
  NFadeObjm 120, Alien_RR, "ALIEN_RR_D", "ALIEN_RR_A"
  NFadeObjm 120, Alien_LR, "ALIEN_LR_D", "ALIEN_LR_A"

  NFadeObjm 123, Alien_LF, "ALIEN_LF_C", "ALIEN_LF_A"
  NFadeObjm 123, Alien_RF, "ALIEN_RF_C", "ALIEN_RF_A"
  NFadeObjm 123, Alien_RR, "ALIEN_RR_C", "ALIEN_RR_A"
  NFadeObjm 123, Alien_LR, "ALIEN_LR_C", "ALIEN_LR_A"

  NFadeObjm 128, Alien_LF, "ALIEN_LF_E", "ALIEN_LF_A"
  NFadeObjm 128, Alien_RF, "ALIEN_RF_E", "ALIEN_RF_A"
  NFadeObjm 128, Alien_RR, "ALIEN_RR_E", "ALIEN_RR_A"
  NFadeObjm 128, Alien_LR, "ALIEN_LR_E", "ALIEN_LR_A"
End If

If SideWallFlashers = 1 Then
  NFadeObjmm 120, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_DF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_D"
  NFadeObjmm 123, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_CF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_C"
  NFadeObjmm 128, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_EF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_E"
  NFadeObjmm 117, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
  NFadeObjmm 118, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
  NFadeObjmm 119, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
  NFadeObjmm 125, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
  NFadeObjmm 126, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
  NFadeObjmm 127, SIDEWALL, SIDEWALL_GIOFF, "SIDEWALL_FF", "SIDEWALL_B", "SIDEWALL_A", "SIDEWALL_F"
End If







End Sub

'//////////////////////////////////////////////////////////////////////
'// LAMP SUBS
'//////////////////////////////////////////////////////////////////////

Sub InitLamps()
Dim x
For x = 0 to 200
LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
FadingLevel(x) = 4       ' used to track the fading state
ModLevel(x) = 0
FlashSpeedUp(x) = 0.2    ' faster speed when turning on the flasher
FlashSpeedDown(x) = 0.01  ' slower speed when turning off the flasher
FlashMax(x) = 1          ' the maximum value when on, usually 1
FlashMin(x) = 0          ' the minimum value when off, usually 0
FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
Next
End Sub

Sub AllLampsOff
Dim x
For x = 0 to 200
SetLamp x, 0
Next
End Sub

Sub SetLamp(nr, value)
  If value <> LampState(nr) Then
    LampState(nr) = abs(value)
    FadingLevel(nr) = abs(value) + 4
  End If
End Sub

Sub SetModLamp(nr, value)
  dim TurnedOn:TurnedOn = 0
  if value > 0 then TurnedOn = 1
  If TurnedOn <> LampState(nr) Then
    LampState(nr) = abs(TurnedOn)
    FadingLevel(nr) = abs(TurnedOn) + 4
  end if
  ModLevel(nr) = value
End Sub

'//////////////////////////////////////////////////////////////////////
'// VPX STANDARD LIGHTS
'//////////////////////////////////////////////////////////////////////

Sub NFadeL(nr, object)
Select Case FadingLevel(nr)
Case 4:object.state = 0:FadingLevel(nr) = 0
Case 5:object.state = 1:FadingLevel(nr) = 1
End Select
if (ModLevel(nr) > 0) Then
  object.IntensityScale = ModLevel(nr) / 255
end if

End Sub

Sub NFadeLm(nr, object, object2) ' used for multiple lights
Select Case FadingLevel(nr)
Case 4:object.state = 0 : object2.state=0
Case 5:object.state = 1: object2.state=1
End Select
if (ModLevel(nr) > 0) Then
  object.IntensityScale = ModLevel(nr) / 128 : object2.IntensityScale=ModLevel(nr) / 128
end if
End Sub


Sub NFadeLmf(nr, object, object2) ' used for multiple lights
Select Case FadingLevel(nr)
Case 4:object.state = 0 : object2.visible=0
Case 5:object.state = 1: object2.visible=1
End Select
End Sub


Sub NFadeLf(nr, object) ' used for multiple lights
Select Case FadingLevel(nr)
Case 4:object.visible=0
Case 5:object.visible=1
End Select
End Sub



Sub NFadeLmm(nr, object, object2, object3) ' used for multiple lights
Select Case FadingLevel(nr)
Case 4:object.state = 0 : object2.state=0 : object3.state=0
Case 5:object.state = 1: object2.state=1 : object3.state=1
End Select
if (ModLevel(nr) > 0) Then
  object.IntensityScale = ModLevel(nr) / 128 : object2.IntensityScale=ModLevel(nr) / 128 : object3.IntensityScale =ModLevel(nr) / 128
end if
End Sub

'//////////////////////////////////////////////////////////////////////
'// VPX RAMP & PRIMITIVE LIGHTS
'//////////////////////////////////////////////////////////////////////

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


Sub NFadeObjmm(nr, object, object2, a, b, c, d)
Select Case FadingLevel(nr)
Case 4:object.image = b:FadingLevel(nr) = 0: object2.image = c:FadingLevel(nr) = 0 'off
Case 5:object.image = a:FadingLevel(nr) = 1: object2.image = d:FadingLevel(nr) = 1 'on
End Select
End Sub



Sub NFadeObjm(nr, object, a, b)
Select Case FadingLevel(nr)
Case 4:object.image = b
Case 5:object.image = a
End Select
End Sub

'//////////////////////////////////////////////////////////////////////
'// VPX FLASHER OBJECTS
'//////////////////////////////////////////////////////////////////////

Sub SetFlash(nr, stat)
FadingLevel(nr) = ABS(stat)
End Sub

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

'//////////////////////////////////////////////////////////////////////
'// VPX REELS & TEXT
'//////////////////////////////////////////////////////////////////////

Sub FadeR(nr, object)
Select Case FadingLevel(nr)
Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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



'//////////////////////////////////////////////////////////////////////
' FLIPPER SHADOWS
'//////////////////////////////////////////////////////////////////////

Sub GraphicsTimer_Timer()
    batleft.objrotz = LeftFlipper.CurrentAngle + 1
    batleftshadow.objrotz = batleft.objrotz
    batright.objrotz = RightFlipper.CurrentAngle - 1
    batrightshadow.objrotz  = batright.objrotz
End Sub

'//////////////////////////////////////////////////////////////////////
' BALL SHADOWS
'//////////////////////////////////////////////////////////////////////


Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
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
  If UBound(BOT) > UBound (BallShadow) Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
      if not BOT(b) is Cball then
        If BOT(b).X < 1024/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (1024/2))/7)) + 5
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (1024/2))/7)) - 5
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
      end if
    Next
End Sub



'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

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
    elseif TargetBouncerEnabled = 2 and aball.z < 30 then
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

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity


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
    debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          Debug.Print "ball in flip1. exit"
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
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then
      EOSNudge1 = 0
    end if
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
End Sub

Sub SleevesBounce_Hit(idx)
  TargetBouncer Activeball, 0.5
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
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
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0000001)
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


Sub RDampen_Timer()
  Cor.Update
End Sub


'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


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

Dim tablewidth, tableheight : tablewidth = AFM.width : tableheight = AFM.height

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
  PlayTargetSound
End Sub

Sub TargetsBounce_Hit (idx)
  TargetBouncer Activeball, 1
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


'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "

Const KeepLogs = False

Class DebugLogFile

    Private Filename
    Private TxtFileStream

    Private Function LZ(ByVal Number, ByVal Places)
        Dim Zeros
        Zeros = String(CInt(Places), "0")
        LZ = Right(Zeros & CStr(Number), Places)
    End Function

    Private Function GetTimeStamp
        Dim CurrTime, Elapsed, MilliSecs
        CurrTime = Now()
        Elapsed = Timer()
        MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
        GetTimeStamp = _
            LZ(Year(CurrTime),   4) & "-" _
            & LZ(Month(CurrTime),  2) & "-" _
            & LZ(Day(CurrTime),    2) & " " _
            & LZ(Hour(CurrTime),   2) & ":" _
            & LZ(Minute(CurrTime), 2) & ":" _
            & LZ(Second(CurrTime), 2) & ":" _
            & LZ(MilliSecs, 4)
    End Function

' *** Debug.Print the time with milliseconds, and a message of your choice
    Public Sub WriteToLog(label, message, code)
        Dim FormattedMsg, Timestamp
        'Filename = UserDirectory + "\" + cGameName + "_debug_log.txt"
    Filename = cGameName + "_debug_log.txt"

        Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, code, True)
        Timestamp  = GetTimeStamp
        FormattedMsg = GetTimeStamp + " : " + label + " : " + message
        TxtFileStream.WriteLine FormattedMsg
        TxtFileStream.Close
    debug.print label & " : " & message
  End Sub

End Class

Sub WriteToLog(label, message)
  if KeepLogs Then
    Dim LogFileObj
    Set LogFileObj = New DebugLogFile
    LogFileObj.WriteToLog label, message, 8
  end if
End Sub

Sub NewLog()
  if KeepLogs Then
    Dim LogFileObj
    Set LogFileObj = New DebugLogFile
    LogFileObj.WriteToLog "NEW LOG", " ", 2
  end if
End Sub

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
'- To stop tracking ball, use WireRampoff
'-- Otherwise, the ball will auto remove if it's below 30 vp units
'
'Example:
'Sub RampSoundPlunge1_hit() : WireRampOn  False, "LeftRamp" : End Sub     ' Wire Ramp Sound
'Sub RampSoundPlunge2_hit() : WireRampOn True, "CenterRamp" : End Sub     ' Plastic Ramp Sound
'Sub RampSoundPlunge3_hit() : WireRampOnSndX False, 2, "RightRamp" : End Sub    ' Wire Ramp Sound volume multiplication of 2
'Sub RampEntry_Hit() : If activeball.vely < -10 then WireRampOn True : End Sub  'Ramp enterance

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'                    Set BallSndMultiplyer and RampID to be same as RampBalls or tnob + 1
dim RampBalls(6,2)
dim BallSndMultiplyer(6)
dim RampID(6)

'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)

Sub WireRampOnSndX(input, sndMultiplier, rmp)  : Waddball ActiveBall, input, sndMultiplier, rmp : RampRollUpdate: End Sub
Sub WireRampOn(input, rmp)  : Waddball ActiveBall, input, 1, rmp : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub

' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput, sndMultiplier, rampname) 'Add ball
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
      BallSndMultiplyer(x) = sndMultiplier
      RampID(x) = rampname
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
      BallSndMultiplyer(x) = 1
      RampID(x) = Empty
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
        WriteToLog "1 Ramp " & RampID(x), "Volume = " & VolPlayfieldRoll(RampBalls(x,0)) * 1.1 * VolumeDial
        WriteToLog "2 Ramp " & RampID(x), "Volume = " & VolPlayfieldRoll(RampBalls(x,0)) * 1.1 * VolumeDial * BallSndMultiplyer(x)
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial * BallSndMultiplyer(x), AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial * BallSndMultiplyer(x), AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
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

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************


'   _________    ____  ________  __   _       ________    __       ____  ______
'   / ____/   |  / __ \/_  __/ / / /  | |     / /  _/ /   / /      / __ )/ ____/
'  / __/ / /| | / /_/ / / / / /_/ /   | | /| / // // /   / /      / __  / __/
' / /___/ ___ |/ _, _/ / / / __  /    | |/ |/ // // /___/ /___   / /_/ / /___
'/_____/_/  |_/_/ |_| /_/ /_/ /_/     |__/|__/___/_____/_____/  /_____/_____/
'
'   ____  __  ______  _____
'  / __ \/ / / / __ \/ ___/
' / / / / / / / /_/ /\__ \
'/ /_/ / /_/ / _, _/___/ /
'\____/\____/_/ |_|/____/
'

' Thalamus : Exit in a clean and proper way
Sub AFM_exit
  Controller.Pause = False
  Controller.Stop
End Sub

'  _____ _           __        __    _ _    _               ____                 _   ____                     _                    ___     _____
' |_   _| |__   ___  \ \      / /_ _| | | _(_)_ __   __ _  |  _ \  ___  __ _  __| | |  _ \ _ __ ___ _ __ ___ (_)_   _ _ __ ___    / / |   | ____|
'   | | | '_ \ / _ \  \ \ /\ / / _` | | |/ / | '_ \ / _` | | | | |/ _ \/ _` |/ _` | | |_) | '__/ _ \ '_ ` _ \| | | | | '_ ` _ \  / /| |   |  _|
'   | | | | | |  __/   \ V  V / (_| | |   <| | | | | (_| | | |_| |  __/ (_| | (_| | |  __/| | |  __/ | | | | | | |_| | | | | | |/ / | |___| |___
'   |_| |_| |_|\___|    \_/\_/ \__,_|_|_|\_\_|_| |_|\__, | |____/ \___|\__,_|\__,_| |_|   |_|  \___|_| |_| |_|_|\__,_|_| |_| |_/_/  |_____|_____|
'                                                   |___/
' Authors:
' Graphics coding, textures, inserts, 3d objects, zombie models, playfield image: Flupper1
' Nfozzy, Fleep addition: Robby King Pin
' Physics tuning : Rothbauerw
' Plastics scans: HiRez00
' Decals apron - ClarkKent
' Apron cards - Carny Priest
' Testing - VPW with special thanks to Iaakki for several fixes and Primetime5k for physics tuning
' VR room - DaRdog
' special thanks to Vbousquet for the Blender toolkit, Nfozzy for the physics scripts, Fleep for the sounds
' former authors; reused bits and pieces
' ICPJUGGLA, FRENETICAMNESIC, GTXJOE, RANDR, DARK, ZANY, Dozer, PinBill

Option Explicit : Randomize : Setlocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const UseVPMModSol = 2

'*******************************************
'  ZOPT: User Options
'*******************************************

' >>>>>
' >>>>> All user options are accessible from the F12 menu, not in script
' >>>>> Please don't change the tonemapping to anything else then AgX, table is optimized & rendered for AgX
' >>>>> Really, no more user options below this point
' >>>>>

'*******************************************
'  ZCON: Constants and Global Variables
'*******************************************

Dim visbl
For Each visbl In physicalramps : visbl.visible = false :If TypeName(visbl) = "Wall" Then visbl.SideVisible  = false : End If: Next
For Each visbl In physicalwalls : visbl.visible = false :If TypeName(visbl) = "Wall" Then visbl.SideVisible  = false : End If: Next
For Each visbl In physicaltargets : visbl.visible = false : If TypeName(visbl) = "Wall" Then visbl.SideVisible = false : End If: Next
For Each visbl In physicalbumpers : visbl.BaseVisible = false : visbl.SkirtVisible = false : visbl.RingVisible = false : Next
For Each visbl In physicalmovables : visbl.visible = false : If TypeName(visbl) = "Wall" Then visbl.SideVisible  = false : End If : Next
For Each visbl In rubbers : visbl.visible = false :Next

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim FSSMode:FSSMode = Table1.ShowFSS

Dim VRTest : VRTest = False

Const BallShadowOn = True
Const BallSize = 50        'Ball diameter in VPX units; must be 50
Const BallMass = 1         'Ball mass must be 1
Const tnob = 4           'Total number of balls the table can hold
Const lob = 0          'Locked balls
Const cGameName = "twd_160h"   'The unique alphanumeric name for this table

Dim tablewidth : tablewidth = Table1.width
Dim tableheight : tableheight = Table1.height

Sub table1_Paused
  Controller.Pause = True
  Dim x : For x = 1 To UBound(RampBalls) : StopSound("RampLoop" & x) : StopSound("wireloop" & x) : Next
End Sub
Sub table1_unPaused:Controller.Pause = False : end sub

'*******************************************
' ZTIM:Timers
'*******************************************

Dim VarHidden, UseVPMDMD
If Desktopmode or RenderingMode = 2 Then
  UseVPMDMD = True : VarHidden = 1
Else
  UseVPMDMD = False : VarHidden = 0
End If


LoadVPM "01560000", "sam.VBS", 3.10

Sub TimerPlunger_Timer

  If VRCab_PlungerHead.Y < 794.303 then
      VRCab_PlungerHead.Y = VRCab_PlungerHead.Y + 5
  End If
  If VRCab_Plunger.Y < -69.95 then
      VRCab_Plunger.Y = VRCab_Plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
    VRCab_PlungerHead.Y = 694.303 + (5* Plunger.Position) -20
    VRCab_Plunger.Y = -169.95 + (5* Plunger.Position) -20
End Sub

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = "coin"

'Const AllInsertsAction = 2 '0=normal, 1=all off, 2 = all on

'************
' Table init.
'************

Dim xx
Dim Bump1,Bump2,Bump3,Mech3bank,bsVUK,visibleLock,bsTEject,bsceject
Dim PMag, DMag, WMag
Dim TWDBall1, TWDBall2, TWDBall3, TWDBall4, gBOT

Sub Table1_Init
  With Controller
    vpmInit Me
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "The Walking Dead Premium/LE"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 1
    .Hidden = VarHidden
    .Games(cGameName).Settings.Value("sound") = 1
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With

    On Error Goto 0

  'Ball initializations need for physical trough
  Set TWDBall1 = sw21.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  Set TWDBall2 = sw20.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  Set TWDBall3 = sw19.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  Set TWDBall4 = sw18.CreateSizedballWithMass(Ballsize / 2,Ballmass)

  Controller.switch(21) = 1
  Controller.switch(20) = 1
  Controller.switch(19) = 1
  Controller.switch(18) = 1

  '*** Use gBOT in the script wherever BOT is normally used. Then there is no need for GetBalls calls ***
  gBOT = Array(TWDBall1, TWDBall2, TWDBall3, TWDBall4)

  Set PMag = New cvpmMagnet : PMag.InitMagnet PrisonMag,25 : PMag.GrabCenter = False '42 25
  Set WMag = New cvpmMagnet : WMag.InitMagnet WWMag,25 : WMag.GrabCenter = False '42
  Set DMag = New cvpmMagnet : DMag.InitMagnet Divmagnet,25 : DMag.GrabCenter = False '100 25

  '**Nudging
  vpmNudge.TiltSwitch = -7
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  '**Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  'Fast Flips
  On Error Resume Next
  InitVpmFFlipsSAM
  If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available from vp10.5 and higher"
  On Error Goto 0

  InitLights InsertOn, InsertLights, InsertOff, InsertFlasher
  InitVLM
  If version<10800 or Controller.version <03060000 or PlatformBits = "32" Then warning.visible = 1
End Sub

Sub Table1_KeyDown(ByVal Keycode)
  ' DebugShotTableKeyDownCheck keycode
  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    If VRRoom > 0 Then VRCab_FlipperLeft.x = VRCab_FlipperLeft.x + 5
  End If
  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    If VRRoom > 0 Then VRCab_FlipperRight.x = VRCab_FlipperRight.x - 5
  End If
  If keycode = PlungerKey Then
    Plunger.Pullback
    If NOT Controller.Switch(23) = True Then
    Controller.Switch(71) = 1
    End If
    End If

  If keycode = PlungerKey Then
    If VRRoom > 0 then
      TimerPlunger.Enabled = True
      TimerPlunger2.Enabled = False
    End If
  End If
    If keycode = StartGameKey Then
    Controller.Switch(16) = 1
    If VRRoom > 0 Then
      StartButtonon.y = StartButtonon.y - 2
      StartButtonoff.y = StartButtonoff.y - 2
    End If
  End If

    If keycode = LeftTiltKey Then Nudge 90, 2:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 2:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()
    If keycode = RightMagnaSave or keycode = LockBarKey Then Controller.Switch(71) = 1 : firebuttonon.Z = 51 : firebuttonoff.Z = 51
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3  Then
    Select Case Int(rnd*3)
      Case 0: PlaySoundAtLevelStatic ("Coin_In_1"), CoinSoundLevel, crossshadow
      Case 1: PlaySoundAtLevelStatic ("Coin_In_2"), CoinSoundLevel, crossshadow
      Case 2: PlaySoundAtLevelStatic ("Coin_In_3"), CoinSoundLevel, crossshadow
    End Select
  End If

    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  'DebugShotTableKeyUpCheck keycode
    If keycode = StartGameKey Then
    Controller.Switch(16) = 0
    If VRRoom > 0 Then
      StartButtonon.y = StartButtonon.y + 2
      StartButtonoff.y = StartButtonoff.y + 2
    End If
  End If
  If vpmKeyUp(keycode) Then Exit Sub
  If keycode = PlungerKey Then dmag.MagnetOn = False
  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    If VRRoom > 0 Then VRCab_FlipperLeft.x = 2117.037
  End If
  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    If VRRoom > 0 Then VRCab_FlipperRight.x = 979.6727
  End If

    If keycode = PlungerKey Then
    If VRRoom > 0 then
      TimerPlunger.Enabled = False
      TimerPlunger2.Enabled = True
      VRCab_PlungerHead.Y = 694.303
      VRCab_Plunger.Y = -169.95
    End if
  End If

  If keycode = RightMagnaSave or keycode = LockBarKey Then Controller.Switch(71) = 0 : : firebuttonon.Z = 55 : firebuttonoff.Z = 55
  If keycode = PlungerKey Then Plunger.Fire : PlaySound "plunger" : Controller.Switch(71) = 0 : End If
End Sub

'Solenoids
SolCallback(1)     = "solTrough"
SolCallback(2)     = "solAutofire"
SolCallback(3)     = "PrisonDoorsPower"
SolCallback(4)     = "PrisonDoorsHold"
SolCallback(5)     = "DmagHandled"
SolCallback(6)     = "WmagHandled"
SolCallback(7)     = "PmagHandled"
'SolCallback(9)     = "SolLeftPop" 'left pop bumper
'SolCallback(10)    = "SolRightPop" 'right pop bumper
'SolCallback(11)    = "SolTopPop" 'top pop bumper
SolCallback(12)    = "SolLeftBank"
'SolCallback(13)    = "???" 'left slingshot
'SolCallback(14)    = "???" 'right slingshot
SolCallback(15)    = "SolLFlipper"
SolCallback(16)    = "SolRFlipper"
'SolModCallback(17) = "Sol17" 'empty in manual; Topper head 1 flasher
'SolModCallback(18) = "Sol18" 'empty in manual; Topper head 2 flasher
SolModCallback(19) = "WellWalkerFlash" 'wellwalker flasher
SolModCallback(20) = "Sol20" 'right spinner
SolCallback(21)    = "Dazz_Motor" 'horde flasher
'SolModCallback(22) = "Sol22" 'empty in manual; Topper head 3 flasher
'SolModCallback(23) = "Sol23" 'empty in manual; Topper illumination flasher
'SolCallback(24)    = "optional e.g. coin meter" (from manual)
SolModCallback(25) = "FlashPops" 'flash pops
SolModCallback(26) = "Prison_Top" 'prison top flasher
SolModCallback(27) = "Prison_Front" 'prison l/r flasher x2
SolModCallback(28) = "LeftDome" 'left dome flasher
SolModCallback(29) = "RightDome" 'right dome flasher
'SolModCallback(30) = "???" 'empty in manual
SolModCallback(31) = "Sol31" 'left loop flasher
'SolModCallback(32) = "Sol32" 'center lane flasher *** not coded ***
SolCallback(51)    = "Shoot_Cannon"
SolCallback(52)    = "DT_Down"
SolCallback(53)    = "DT_UP"
'SolCallback(55)    = "????" ' bicycle ramp PWR/open, only goes enabled when the ramp is lifting for a short time
SolCallback(56)    = "Corpse_Ramp" 'bicycle ramp hold

'*******************************************
' ZDRN: Drain, Trough, and Ball Release
'*******************************************
' It is best practice to never destroy balls. This leads to more stable and accurate pinball game simulations.
' The following code supports a "physical trough" where balls are not destroyed.
' To use this,
'   - The trough geometry needs to be modeled with walls, and a set of kickers needs to be added to
'  the trough. The number of kickers depends on the number of physical balls on the table.
'   - A timer called "UpdateTroughTimer" needs to be added to the table. It should have an interval of 300 and be initially disabled.
'   - The balls need to be created within the Table1_Init sub. A global ball array (gBOT) can be created and used throughout the script


'TROUGH
Sub sw22_Hit:Controller.Switch(22) = 1:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw21_Hit:UpdateTrough:Controller.Switch(21) = 1:End Sub
Sub sw21_UnHit:UpdateTrough:Controller.Switch(21) = 0:End Sub

Sub sw20_Hit:UpdateTrough:Controller.Switch(20) = 1:End Sub
Sub sw20_UnHit:UpdateTrough:Controller.Switch(20) = 0:End Sub

Sub sw19_Hit:UpdateTrough:Controller.Switch(19) = 1:End Sub
Sub sw19_UnHit:UpdateTrough:Controller.Switch(19) = 0:End Sub

Sub sw18_Hit:UpdateTrough:Controller.Switch(18) = 1:End Sub
Sub sw18_UnHit:UpdateTrough:Controller.Switch(18) = 0:End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 150
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw21.BallCntOver = 0 Then sw20.kick 61, 20
  If sw20.BallCntOver = 0 Then sw19.kick 61, 20
  If sw19.BallCntOver = 0 Then sw18.kick 61, 20
  If sw18.BallCntOver = 0 Then drain.kick 61, 20
  Me.Enabled = 0
End Sub

Sub Drain_Hit
  RandomSoundDrain Drain
  UpdateTrough
End Sub

Sub solTrough(enabled)
  If enabled Then
    sw21.kick 61, 10
    RandomSoundBallRelease sw21
  End If
End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    Plunger1.Fire
    PlaySound SoundFX("fx_solenoid",DOFContactors)
  End If
End Sub

'******************************************************
'   ZRDT:  DROP TARGETS by Rothbauerw
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


Sub sw9_hit:DTHit 9:End Sub
Sub sw10_hit:DTHit 10:End Sub
Sub sw11_hit:DTHit 11:End Sub

Sub sw48_hit:DTHit 48:End Sub

Sub DT_Down(Enabled)
  If Enabled Then
    DTDrop 48
  End If
End Sub

Sub DT_Up(Enabled)
  If Enabled Then
    DTRaise 48
    RandomSoundDropTargetReset psw48
  End If
End Sub

Sub SolLeftBank(Enabled)
  If Enabled Then
    DTRaise 9
    DTRaise 10
    DTRaise 11
    RandomSoundDropTargetReset psw10
  End If
End Sub

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
Dim DT9, DT10, DT11, DT48

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:  primary target wall to determine drop
'   secondary:  wall used to simulate the ball striking a bent or offset target after the initial Hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          rotz must be used for orientation
'          rotx to bend the target back
'          transz to move it up and down
'          the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'   switch:  ROM switch number
'   animate:  Array slot for handling the animation instrucitons, set to 0
'          Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:  Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.
'         Use the function DTDropped(switchid) to check a target's drop status.

Set DT9 = (new DropTarget)(sw9, sw9a, psw9, 9, 0, False)
Set DT10 = (new DropTarget)(sw10, sw10a, psw10, 10, 0, False)
Set DT11 = (new DropTarget)(sw11, sw11a, psw11, 11, 0, False)
Set DT48 = (new DropTarget)(sw48, sw48a, psw48, 48, 0, False)

Dim DTArray
DTArray = Array(DT9, DT10, DT11, DT48)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 50 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

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
  Dim animtime, rangle, nrx

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
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  If animate = 2 Then
    transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
    If prim.transz >  - DTDropUnits  Then
      prim.transz = transz
      Select Case switchid
        Case 9 : For Each nrx In BP_psw9: nrx.transz = transz:Next
        Case 10 : For Each nrx In BP_psw10: nrx.transz = transz:Next
        Case 11 : For Each nrx In BP_psw11: nrx.transz = transz:Next
        Case 48 : For Each nrx In BP_psw48: nrx.transz = transz:Next
      End Select
    End If

    prim.rotx = DTMaxBend * Cos(rangle) / 2
    prim.roty = DTMaxBend * Sin(rangle) / 2

    If prim.transz <= - DTDropUnits Then
      prim.transz =  - DTDropUnits
      secondary.collidable = 0
      DTArray(ind).isDropped = True 'Mark target as dropped
      controller.Switch(Switchid) = 1
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

      For b = 0 To UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
          gBOT(b).velz = 20
        End If
      Next
    End If

    If prim.transz < 0 Then
      prim.transz = transz
      Select Case switchid
        Case 9 : For Each nrx In BP_psw9: nrx.transz = transz:Next
        Case 10 : For Each nrx In BP_psw10: nrx.transz = transz:Next
        Case 11 : For Each nrx In BP_psw11: nrx.transz = transz:Next
        Case 48 : For Each nrx In BP_psw48: nrx.transz = transz:Next
      End Select
    ElseIf transz > 0 Then
      prim.transz = transz
      Select Case switchid
        Case 9 : For Each nrx In BP_psw9: nrx.transz = transz:Next
        Case 10 : For Each nrx In BP_psw10: nrx.transz = transz:Next
        Case 11 : For Each nrx In BP_psw11: nrx.transz = transz:Next
        Case 48 : For Each nrx In BP_psw48: nrx.transz = transz:Next
      End Select
    End If

    If prim.transz > DTDropUpUnits Then
      DTAnimate =  - 2
      prim.transz = DTDropUpUnits
      Select Case switchid
        Case 9 : For Each nrx In BP_psw9: nrx.transz = DTDropUpUnits:Next
        Case 10 : For Each nrx In BP_psw10: nrx.transz = DTDropUpUnits:Next
        Case 11 : For Each nrx In BP_psw11: nrx.transz = DTDropUpUnits:Next
        Case 48 : For Each nrx In BP_psw48: nrx.transz = DTDropUpUnits:Next
      End Select
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = GameTime
    End If
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = False 'Mark target as not dropped
    controller.Switch(Switchid) = 0
  End If

  If animate =  - 2 And animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
    If prim.transz < 0 Then
      prim.transz = 0
      Select Case switchid
        Case 9 : For Each nrx In BP_psw9: nrx.transz = 0:Next
        Case 10 : For Each nrx In BP_psw10: nrx.transz = 0:Next
        Case 11 : For Each nrx In BP_psw11: nrx.transz = 0:Next
        Case 48 : For Each nrx In BP_psw48: nrx.transz = 0:Next
      End Select
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
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************

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

'Function RotPoint(x,y,angle)
' Dim rx, ry
' rx = x * dCos(angle) - y * dSin(angle)
' ry = x * dSin(angle) + y * dCos(angle)
' RotPoint = Array(rx,ry)
'End Function



'******************************************************
'****  END DROP TARGETS
'******************************************************


'******************************************************
' ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************

Sub sw44_Hit: STHit 44 : End Sub 'l prison standup
Sub sw45_Hit: STHit 45 : End Sub 'r prison standup

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
Dim ST44, ST45

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST44 = (new StandupTarget)(sw44, SU1, 44, 0)
Set ST45 = (new StandupTarget)(sw45, SU2, 45, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST44, ST45)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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
    'prim.transy =  - STMaxOffset
    prim.transx =  - STMaxOffset
    vpmTimer.PulseSw switch
    STAnimate = 2
    Exit Function
  ElseIf animate = 2 Then
    'prim.transy = prim.transy + STAnimStep
    prim.transx = prim.transx + STAnimStep
    If prim.transx >= 0 Then
      prim.transx = 0
'   If prim.transy >= 0 Then
'     prim.transy = 0
      primary.collidable = 1
      STAnimate = 0
      Exit Function
    Else
      STAnimate = 2
    End If
  End If
End Function

'******************************************************
'****   END STAND-UP TARGETS
'******************************************************

'*******************************************
'  ZKTA: SW38 TARGET
'*******************************************

Sub sw38_Hit:KTHit 38:vpmTimer.PulseSw 38:End Sub

Class KickingTarget
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

'Define a variable for each kicking target
Dim KT38

'Set array with kicking target objects
'
'KickingTargetvar = Array(primary, prim, swtich)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'


Set KT38 = (new KickingTarget)(sw38, pSW38, 38, 0)

'Add all the Kicking Target Arrays to Kicking Target Animation Array
'   KTAnimationArray = Array(KT1, KT2, ....)
Dim KTArray
KTArray = Array(KT38)

' kSpring   - strength of the target spring (non-kicking)
' kKickVel  - velocity imparted on the ball with the target solenoid fires
' kDist   - distance the target must be displaced before the switch registers and the solenoid fires
' kWidth  - total width of the target
Dim kSpring, kKickVel, kDist, kWidth
kSpring = 1
kKickvel = 30
kDist = 20
kWidth = 66

'''''' KICKING TARGETS FUNCTIONS
Dim KTBall

Sub KTHit(switch)
  Dim i
  i = KTArrayID(switch)

  PlayTargetSound
  KTArray(i).animate = STCheckHit(ActiveBall,KTArray(i).primary)

  Set KTBall = Activeball

  Activeball.velx = cor.BallVelx(Activeball.id)
  Activeball.vely = cor.BallVely(Activeball.id)

  DoKTAnim
End Sub

Sub KTKick(switch)
  Dim i
  i = KTArrayID(switch)

  KTArray(i).animate = 2

  Activeball.velx = cor.BallVelx(Activeball.id)
  Activeball.vely = cor.BallVely(Activeball.id)

  DoKTAnim
End Sub


Function KTArrayID(switch)
  Dim i
  For i = 0 To UBound(KTArray)
    If KTArray(i).sw = switch Then
      KTArrayID = i
      Exit Function
    End If
  Next
End Function

Sub DoKTAnim()
  Dim i
  For i = 0 To UBound(KTArray)
    KTArray(i).animate = KTAnimate(KTArray(i).primary,KTArray(i).prim,KTArray(i).sw,KTArray(i).animate)
  Next
End Sub

Function KTAnimate(primary, prim, switch,  animate)
  KTAnimate = animate

  If animate = 0  Then
    primary.uservalue = 0
    KTAnimate = 0
    primary.collidable = 1
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  Dim animtime, btdist, btwidth, btangle, bangle, angle, nposx, nposy, tdist, kparavel, kballvel
  tdist = 31.31

  angle = primary.orientation

  animtime = GameTime - primary.uservalue
  primary.uservalue = GameTime

  nposx = primary.x + prim.transy * dCos(angle + 90)
  nposy = primary.y + prim.transy * dSin(angle + 90)

  btwidth = DistancePL(KTBall.x,KTBall.y,primary.x,primary.y,primary.x+dcos(angle+90),primary.y+dsin(angle+90))
  btdist = DistancePL(KTBall.x,KTBall.y,nposx,nposy,nposx+dcos(angle),nposy+dsin(angle))

  kballvel = sqr(KTBall.velx^2 + KTBall.vely^2)

  if kballvel <> 0 Then
    bangle = arcCos(KTball.velx/kballvel)*180/PI
  Else
    bangle = 0
  End If

  If KTBall.vely < 0 Then
    bangle = bangle * -1
  End If

  btangle = bangle - angle

  prim.transy = prim.transy + btdist - tdist

  'debug.print btangle & " " & btwidth & " " & btdist &  " " & prim.transy

  If animate = 1 Then
    primary.collidable = 0

    If btdist < tdist and btwidth < kWidth/2 + 25 Then
      if abs(prim.transy) >= kDist Then
        vpmTimer.PulseSw switch
        'fire Solenoid
        RandomSoundSlingshotLeft primary
        kparavel = Sqr(KTBall.velx^2 + KTBall.vely^2)*dCos(btangle)
        KTBall.velx = dcos(angle)*kparavel - (kKickVel * dsin(angle))
        KTBall.vely = dsin(angle)*kparavel + (kKickVel * dcos(angle))
        KTAnimate = 3
        'debug.print KTBall.velx & " fire " & KTBall.vely & " kpara " & kparavel
      Else
        KTBall.velx = KTBall.velx - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000)
        KTBall.vely = KTBall.vely + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)
        'debug.print KTBall.velx & " nofire " & KTBall.vely
        'debug.print - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000) & " delta " & + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)
      End If
    Elseif btdist > tdist then
      KTBall.velx = KTBall.velx - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000)
      KTBall.vely = KTBall.vely + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)
      'debug.print KTBall.velx & " nofire " & KTBall.vely
      'debug.print - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000) & " delta " & + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)

      if prim.transy >= 0 Then
        prim.transy = 0
        KTAnimate = 0
        Exit Function
      end If
    Else
      prim.transy = prim.transy + 1
      if prim.transy >= 0 Then
        prim.transy = 0
        KTAnimate = 0
        Exit Function
      end If
    End If
  Elseif animate = 2 Then
    vpmTimer.PulseSw switch
    'fire Solenoid
    RandomSoundSlingshotLeft primary
    kparavel = Sqr(KTBall.velx^2 + KTBall.vely^2)*dCos(btangle)
    KTBall.velx = dcos(angle)*kparavel - (kKickVel * dsin(angle))
    KTBall.vely = dsin(angle)*kparavel + (kKickVel * dcos(angle))
    KTAnimate = 3
    'debug.print KTBall.velx & " fire2 " & KTBall.vely & " kpara " & kparavel
  Elseif animate = 3 Then
    'debug.print KTBall.velx & " fire3 " & KTBall.vely & " tang " & dCos(btangle) * kballvel

    if prim.transy >= 0 Then
      prim.transy = 0
      KTAnimate = 0
      Exit Function
    end If
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

'******************************************************
'*   END SW38 TARGET
'******************************************************

Sub Prison_Top(Enabled) : f26.state = Enabled : End Sub
Sub Prison_Front(Enabled) : f27.state = Enabled : End Sub

Sub Sol31(Enabled) ' *** flash left loop ***
  f31.state = Enabled : 'f31c.state =Enabled
  f31off.BlendDisableLightingFromBelow = -1 + 2 * Enabled
  f31off.BlendDisableLighting = 1 + Enabled * 18
  If Enabled > 0 Then
    f31off.material = "insertround2con" : f31off.image = "insertround2on"
    f31on.visible = 1 : f31on.BlendDisableLighting = 3000 * Enabled
  Else
    f31off.material = "insertround2coff" : f31off.image = "insertround2off"
    f31on.visible = 0 : f31on.BlendDisableLighting = 0
  End If

End Sub

'Sub Sol32(Enabled) ' *** flash center loop ***
' not coded
'End Sub

Sub Sol20(Enabled) ' *** right spinner ***
  f20.state = Enabled
  f20off.BlendDisableLightingFromBelow = -1 + 2 * Enabled
  f20off.BlendDisableLighting = 1 + Enabled * 18
  If Enabled > 0 Then
    f20off.material = "insertround2con" : f20off.image = "insertround2on"
    f20on.visible = 1 : f20on.BlendDisableLighting = 3000 * Enabled
  Else
    f20off.material = "insertround2coff" : f20off.image = "insertround2off"
    f20on.visible = 0 : f20on.BlendDisableLighting = 0
  End If
End Sub


'Sub Sol23(Enabled) 'empty in manual
'End Sub
'
'Sub Sol17(Enabled) 'empty in manual
'End Sub
'
'Sub Sol18(Enabled) 'empty in manual
'End Sub
'
'Sub Sol22(Enabled) 'empty in manual
'End Sub

Sub WellWalkerFlash(Enabled) : f19.state = Enabled : End Sub

Sub FlashPops(Enabled)
  dim x
  If Enabled > 0 Then
    f25.state = Enabled
    flasher25prim.visible = true
    flasher25prim.BlendDisableLighting = 2000 * Enabled '500
  Else
    flasher25prim.visible = false : f25.state = Enabled
  End If
End Sub

Sub LeftDome(Enabled) : f28.state = Enabled : f28b.state = Enabled : End Sub
Sub RightDome(Enabled) : f29.state = Enabled : f29b.state = Enabled : End Sub


'Sub SolTopPop(Enabled)
' 'handled by lamptimer light 60
'End Sub
'
'Sub SolRightPop(Enabled)
' 'handled by lamptimer light 61
'End Sub
'
'Sub SolLeftPop(Enabled)
' 'handled by lamptimer light 62
'End Sub

'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
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

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
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
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub LeftFlipper_animate
  Dim nrx : batleftshadow.ObjRotZ = LeftFlipper.currentAngle
  For Each nrx in BP_BatLeft : nrx.RotZ = LeftFlipper.currentAngle : next
End Sub

Sub RightFlipper_animate
  Dim nrx : batrightshadow.ObjRotZ = RightFlipper.currentAngle
  For Each nrx in BP_BatRight : nrx.RotZ = RightFlipper.currentAngle : next
End Sub

Sub WZ_Init : Controller.Switch (2) = 1 : End Sub

Dim WZOpen, WZrot : WZOpen = False : WZrot = 0
wzf.endangle = 0

Sub WZT_Timer()
  Dim nrx
  If WZOpen = True Then wzf.endangle = WZMax + 2 : wzf.rotatetoend
  If WZOpen = False Then wzf.rotatetostart
  If WZrot >= WZMax Then WZOpen = False: Controller.Switch (2) = 1
  If WZrot = 0 and WZOpen = False Then me.enabled = False : wzf.endangle = 0
  WZrot = wzf.CurrentAngle * 1.5
  For Each nrx in BP_wellzom : nrx.RotX = WZrot : nrx.TransY = - WZrot/3 : next
  For Each nrx in BP_wellzombase : nrx.TransY = - WZrot/3 : next
End Sub

Dim WZMax

Sub WZ_Hit
  Dim nrx
  WZMax = sqr(ActiveBall.VelX^2 + ActiveBall.VelY^2 + ActiveBall.VelZ^2): If WZMax > 33 Then WZMax = 33
  If WZMax > 6 then
    Controller.Switch(2) = 0:WZOpen = True: Playsound "Metal_Touch_9" : WZT.enabled = True
  Else
    WZOpen = True: Playsound "Metal_Touch_9" : WZT.enabled = True
  End If
  ActiveBall.VelX = ActiveBall.VelX * 0.25
  ActiveBall.VelY = ActiveBall.VelY * 0.25
  ActiveBall.VelZ = ActiveBall.VelZ * 0.25
End Sub

Dim prisonstate : prisonstate = False
Dim doorrotz : doorrotz = array(0, 0.40, 0.8, 1.6, 3, 8, 48, 88,  92, 94.4, 95.2, 95.6, 96)
Dim dooractz : dooractz = 0

Sub PrisonDoorsPower(Enabled) : If Enabled Then prisonstate = False : PlaySound SoundFX("fx_solenoid",DOFContactors) : PrisonT.enabled = True : End If : End Sub
Sub PrisonDoorsHold(Enabled) : If not Enabled Then prisonstate = True : PlaySound SoundFX("fx_solenoid",DOFContactors) : PrisonT.enabled = True : End If : End Sub

Sub PrisonT_Timer()
  Dim nrx
  If prisonstate = False Then 'Opening
    If dooractz < 12 Then
      dooractz = dooractz + 1
      For Each nrx in BP_doorleft : nrx.ObjRotZ = doorrotz(dooractz) : Next
      For Each nrx in BP_doorright : nrx.ObjRotZ = -doorrotz(dooractz) : Next
    Else
      PrisonT.Enabled = False : Controller.Switch(4) = 1 : For Each nrx in BL_diode : nrx.visible = 0 : next : LM_GI1_priszom.visible = 1 : LM_GI0_priszom.visible = 1
    End If
    sw4.isdropped = True
  Else  'Closing
    If dooractz > 0 Then
      dooractz = dooractz - 1
      For Each nrx in BP_doorleft : nrx.ObjRotZ = doorrotz(dooractz) : Next
      For Each nrx in BP_doorright : nrx.ObjRotZ = -doorrotz(dooractz) : Next
    Else
      PrisonT.Enabled = False : Controller.Switch(4) = 0 : For Each nrx in BL_diode : nrx.visible = 1 : next : LM_GI1_priszom.visible = 0 : LM_GI0_priszom.visible = 0
    End If
    sw4.isdropped = False
  End If
End Sub

'*******************************************
' ZCAN: Crossbow
'*******************************************

Sub Dazz_Motor(enabled)
  Dim xlss
  If enabled Then
    cbdir = 1 : Darryl.enabled = 1
  Else
    Darryl.enabled = 0
  End If
End Sub

Dim cbdir, ckdir, ccount : cbdir = 1
Dim CBDist: CBDist = Distance(Cannon_Load.x, Cannon_Load.y, BM_crossbow.x, BM_crossbow.y)
Dim CBDrop, CBrotZ: CBDrop = 0 : CBRotZ = 55 + 125

Sub Darryl_Timer()
  Dim nrx
  PlaySound SoundFX("bridge2",DOFGear), , 0.001
  ckdir = CBRotZ - 180
  Select Case cbdir
  Case 1:
    If CBRotZ => 127 Then
      Controller.Switch(50) = 0 : Controller.Switch(52) = 0
      If CBDrop < 85 Then CBDrop = CBDrop + 2.5
    End If
    If CBRotZ => 160 Then Controller.Switch(51) = 1
    If CBRotZ => 198 Then cbdir = 3
    CBRotZ = CBRotZ + 0.2 + Rnd(1) * 0.3
    For Each nrx in BP_crossbow : nrx.RotZ = CBRotZ - 180.5: Next : crossshadow.RotZ = CBRotZ - 180.5
  Case 2:
    If CBRotZ <= 127 Then
      Controller.Switch(50) = 1
      If Not isEmpty(CBBall) Then Controller.Switch(52) = 1
    End If
    If CBRotZ <= 160 Then Controller.Switch(51) = 0
    If CBRotZ <= 124.8 Then
      CBRotZ = 125
      For Each nrx in BP_crossbow : nrx.RotZ = CBRotZ - 180.5: Next : crossshadow.RotZ = CBRotZ - 180.5
      me.enabled = 0
    End If
    CBRotZ = CBRotZ - 0.2 - Rnd(1) * 0.3
    For Each nrx in BP_crossbow : nrx.RotZ = CBRotZ - 180.5: Next : crossshadow.RotZ = CBRotZ - 180.5
  Case 3:
    If ccount = 20 Then cbdir = 2 : ccount = -1
    ccount = ccount + 1
  End Select
  If not isEmpty(CBBall) Then
    CBBall.x = (CBDist - CBDrop) * dCos(CBRotZ + 90) + BM_crossbow.x
    CBBall.y = (CBDist - CBDrop) * dSin(CBRotZ + 90) + BM_crossbow.y
  Else
    CBDrop = 0
  End If
End Sub

Dim CBBall

Sub Cannon_Load_Hit()
  WireRampOff  'Turn off the Plastic Ramp Sound
  SoundSaucerLock
  Set CBBall = Activeball
  Controller.Switch(52) = 1
End Sub

Sub Shoot_Cannon(Enabled)
  If Enabled Then
    dim kang
    If cbdir = 1 Then kang = 0.8 Else kang = -0.8
    kang = kang + CBRotZ - 180
    SoundSaucerKick 1, BM_crossbow
    Cannon_Load.kick kang, 50, 0.05
    CBBall = Empty
  End If
End Sub

Sub Left_Spinner_Spin() : vpmTimer.PulseSw 40 : SoundSpinner Left_Spinner :End Sub

Sub Left_Spinner_Animate
  Dim nr, ang
  ang = Left_spinner.CurrentAngle : If ang > 140 and ang < 320 Then ang = ang - 180 '135 240
  For Each nr in BP_Spin2 : nr.RotX =  ang: Next
End Sub

Sub Right_Spinner_Spin() : vpmTimer.PulseSw 1 : SoundSpinner Right_Spinner : End Sub

Sub Right_Spinner_Animate
  Dim nr, ang
  ang = Right_spinner.CurrentAngle : If ang > 30 and ang < 210 Then ang = ang - 180 '135 270
  For Each nr in BP_Spin1 : nr.RotX =  ang : Next
End Sub

' bicycle girl zombie hit
Dim girlzomRotX : girlzomRotX = 0
Sub sw49_Hit()
  dim nr
  vpmTimer.PulseSw 49
  playsound SoundFX("Metal_Touch_5",DOFTargets)
  girlzomRotX = 0
  Me.timerenabled = 1
End Sub

Dim girlrotx : girlrotx = array(0, 4, 8, 12, 16,19, 20, 21, 20, 19, 16, 12, 8, 4,0)

Sub sw49_timer()
  dim nr
  If girlzomRotX < 14 Then
    girlzomRotX = girlzomRotX + 1
    For Each nr in BP_biczom : nr.RotX = girlrotx(girlzomRotX) : next
  Else
    girlzomRotX = 0 : sw49.timerenabled = 0
  End If
End Sub

Sub Corpse_Ramp(Enabled)
  If Enabled Then
    brpos = 1 : Bottom_Ramp.enabled = 1 ': Trigger1.enabled = 0
  Else
    brpos = 2 : Bottom_Ramp.Enabled = 1
  End If
End Sub


Dim brpos, botramprotX : botramprotX = 0
Dim ramprotx : ramprotx = array(0, 0.05, 0.1, 0.2, 0.5, 1, 6, 11,  11.5, 11.8, 11.9, 11.95, 12)

Sub Bottom_Ramp_Timer()
  dim nr
  Select Case brpos
    Case 1:
      botramprotX = botramprotX + 1
      If botramprotX >= 12 Then
        botramprotX = 12
        BotRamp1.collidable = False : PlaySound SoundFX("Wall_Hit_2",DOFContactors)
        For Each nr in BL_insrt78 : nr.color = RGB(255,0,0) : next
        Me.Enabled = 0
      End If
      For Each nr in BP_BotRamp : nr.RotX = ramprotx(botramprotX) : next
    Case 2:
      botramprotX = botramprotX - 1
      If botramprotX <= 0 Then
        botramprotX = 0
        BotRamp1.collidable = True : PlaySound SoundFX("Wall_Hit_2",DOFContactors)
        For Each nr in BL_insrt78 : nr.color = RGB(0,0,0) : next
        Me.Enabled = 0
      End If
      For Each nr in BP_BotRamp : nr.RotX = ramprotx(botramprotX) : next
  End Select
End Sub

Sub Trigger1_Hit()
  If BotRamp1.collidable Then WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub Trigger2_Hit() : WireRampOn True : End Sub 'Play Plastic Ramp Sound

brpos = 2 : Bottom_Ramp.Enabled = 1

' Prison zombie hit
Dim priszomRotX : priszomRotX = 0
Sub sw3_Hit()
  dim nr
  vpmTimer.PulseSw 3
  playsound SoundFX("Metal_Touch_5",DOFTargets)
  priszomRotX = 0
  Me.timerenabled = 1
End Sub

Dim prisrotx : prisrotx = array(0, 4, 8, 12, 16,19, 20, 21, 20, 19, 16, 12, 8, 4,0)

Sub sw3_timer()
  dim nr
  If priszomRotX < 14 Then
    priszomRotX = priszomRotX + 1
    For Each nr in BP_priszom : nr.RotX = prisrotx(priszomRotX) : next
  Else
    priszomRotX = 0 : sw3.timerenabled = 0
  End If
End Sub

'Sub sw3_UnHit: Controller.Switch(3) = 0: End Sub 'not needed
Sub sw4_Hit  : playsound SoundFX("Metal_Touch_5",DOFShaker): End Sub ' prison door closed
'Sub sw4_UnHit: Controller.Switch(4) = 0: End Sub 'not needed


Sub sw72_Hit  : Controller.Switch(72) = 1: switch80.Z = -5 : End Sub ' star top
Sub sw72_UnHit: Controller.Switch(72) = 0: switch80.Z = 2 : End Sub
Sub sw70_Hit  : Controller.Switch(70) = 1: switch79.Z = -5 : End Sub ' star bottom
Sub sw70_UnHit: Controller.Switch(70) = 0: switch79.Z = 2 : End Sub
Sub sw23_Hit  : Controller.Switch(23) = 1: End Sub ' shooter lane
Sub sw23_UnHit: Controller.Switch(23) = 0: End Sub
Sub sw24_Hit  : Controller.Switch(24) = 1: End Sub ' left outlane
Sub sw24_UnHit: Controller.Switch(24) = 0: End Sub
Sub sw25_Hit  : Controller.Switch(25) = 1: ActiveBall.VelY = ActiveBall.VelY * Rnd * 0.5 : End Sub ' left inlane
Sub sw25_UnHit: Controller.Switch(25) = 0: End Sub
Sub sw28_Hit  : Controller.Switch(28) = 1: ActiveBall.VelY = ActiveBall.VelY * Rnd * 0.5 : End Sub ' right inlane
Sub sw28_UnHit: Controller.Switch(28) = 0: End Sub
Sub sw29_Hit  : Controller.Switch(29) = 1: End Sub ' right outlane
Sub sw29_UnHit: Controller.Switch(29) = 0: End Sub
Sub sw33_Hit  : Controller.Switch(33) = 1: End Sub ' upper shooter lane
Sub sw33_UnHit: Controller.Switch(33) = 0: End Sub
Sub sw34_Hit  : Controller.Switch(34) = 1: End Sub ' upper shooter lane
Sub sw34_UnHit: Controller.Switch(34) = 0: End Sub
Sub sw35_Hit  : vpmTimer.PulseSw 35: End Sub ' l ramp exit
Sub sw36_Hit  : Controller.Switch(36) = 1: End Sub 'left top lane
Sub sw36_UnHit: Controller.Switch(36) = 0: End Sub
Sub sw37_Hit  : Controller.Switch(37) = 1: End Sub 'right top lane
Sub sw37_UnHit: Controller.Switch(37) = 0: End Sub
'Sub sw39_UnHit: Controller.Switch(39) = 0: End Sub ' not needed because of PulseSw
'Sub sw41_UnHit: Controller.Switch(41) = 0: End Sub ' not needed because of PulseSw
Sub sw39_Hit 'right loop
  vpmTimer.PulseSw 39
  If ActiveBall.VelY >= 0 Then
    ActiveBall.ReflectionEnabled = True
  Else
    ActiveBall.ReflectionEnabled = False
  End If
End Sub
Sub sw41_Hit 'left loop
  vpmTimer.PulseSw 41
  If ActiveBall.VelY < 0 Then
    ActiveBall.ReflectionEnabled = False
  Else
    ActiveBall.ReflectionEnabled = True
  End If
End Sub
Sub sw42_Hit  : vpmTimer.PulseSw 42: End Sub ' R ramp exit
Sub sw43_Hit  : vpmTimer.PulseSw 43: End Sub ' L ramp entrance

Sub door_hit_hit()
  dim nrx
  Controller.Switch(46) = 1
  PlaySound "Metal_Touch_11"
  For Each nrx in BP_doorleft : nrx.ObjRotZ =  nrx.ObjRotZ - 2 : Next
  For Each nrx in BP_doorright : nrx.ObjRotZ = nrx.ObjRotZ + 2 : Next
  Me.timerenabled = 1
End Sub

Sub door_hit_timer()
  dim nrx
  For Each nrx in BP_doorleft : nrx.ObjRotZ =  nrx.ObjRotZ + 2 : Next
  For Each nrx in BP_doorright : nrx.ObjRotZ = nrx.ObjRotZ - 2 : Next
  Me.timerenabled = 0
End Sub

Sub door_hit_Unhit()
  Controller.Switch(46) = 0
End Sub

Sub LHD_Hit()
  WireRampOff  'Turn off the Plastic Ramp Sound
  ActiveBall.VelX = 0 : ActiveBall.VelY = 0 : ActiveBall.VelZ = 8
  ActiveBall.AngMomX = 0 : ActiveBall.AngMomY = 0 : ActiveBall.AngMomZ = 0
End Sub

Sub RHD_Hit()
  WireRampOff  'Turn off the Plastic Ramp Sound
  ActiveBall.VelX = 0 : ActiveBall.VelY = 0: ActiveBall.VelZ = 8
  ActiveBall.AngMomX = 0 : ActiveBall.AngMomY = 0 : ActiveBall.AngMomZ = 0
End Sub

Sub sw47_Hit  : Controller.Switch(47) = 1: End Sub 'center lane
Sub sw47_UnHit: Controller.Switch(47) = 0: End Sub
Sub PrisonMag_Hit:PMag.AddBall ActiveBall:End Sub
Sub PrisonMag_UnHit:PMag.RemoveBall ActiveBall:End Sub
Sub DivMagnet_Hit:DMag.AddBall ActiveBall:End Sub
Sub DivMagnet_UnHit:DMag.RemoveBall ActiveBall:End Sub
Sub WWMag_Hit:WMag.AddBall ActiveBall:End Sub
Sub WWMag_UnHit:WMag.RemoveBall ActiveBall:End Sub

'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' - also allows to increase/decrease the frequency of an already playing samples, and in general to apply all the settings to an already playing sample, and to choose if to restart this playing sample from the beginning or not
'   loopcound chooses the amount of loops
'   volume is in 0..1
'   pan ranges from -1.0 (left) over 0.0 (both) to 1.0 (right)
'   randompitch ranges from 0.0 (no randomization) to 1.0 (vary between half speed to double speed)
'   pitch can be positive or negative and directly adds onto the standard sample frequency
'   useexisting is 0 or 1 (if no existing/playing copy of the sound is found, then a new one is created)
'   restart is 0 or 1 (only useful if useexisting is 1)
'   front_rear_fade is similar to pan but fades between the front and rear speakers
' - so f.e. PlaySound "FlipperUp",0,0.5,-1.0,1.0,1,1,0.0 would play the sound not looped (0), at half the volume (0.5), only on left speaker (-1.0), with varying pitch (1.0), it would reuse the same channel if it is already playing
' (1), restarts the sample (1), and plays it on the front speakers only (0.0)

sub PmagHandled(enabled)
  Dim obj
    PMag.MagnetOn=enabled
    If enabled Then
    PlaySound SoundFX("fx_magnet",DOFContactors), 20, 0.5, 0, 0.2
    Else
    ' randomizes the ball release a bit to prevent SDTM
    For Each obj In Pmag.Balls
      If Rnd > 0.5 Then
        obj.Velx = obj.Velx + RndInt(1,4)
      Else
        obj.Velx = obj.Velx + RndInt(-1,-4)
      End If
      obj.VelY = obj.VelY + RndInt(10,20)
      obj.angmomx = rndint(-7,7)
      obj.angmomy = rndint(-7,7)
    Next
    StopSound SoundFX("fx_magnet",DOFContactors)
    end if
end sub

sub WmagHandled(enabled)
    WMag.MagnetOn=enabled
  If enabled Then
    PlaySound SoundFX("fx_magnet",DOFContactors), 20, 0.5, 0, 0.2
    Else
    StopSound SoundFX("fx_magnet",DOFContactors)
    end if
end sub

sub DmagHandled(enabled)
    DMag.MagnetOn=enabled
  If enabled Then
    PlaySound SoundFX("fx_magnet",DOFContactors), 20, 0.5, 0, 0.2
    Else
    StopSound SoundFX("fx_magnet",DOFContactors)
    end if
end sub


'***********************************************
'*********         Slings            ***********
'***********************************************

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw(26)
  RandomSoundSlingshotLeft Left3
  leftsl.PlayAnim 0,0.4 : leftslkick.PlayAnim 0,0.4
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw(27)
  RandomSoundSlingshotRight Right3
  rightsl.PlayAnim 0,0.4 : rightslkick.PlayAnim 0,0.4
  End Sub

'***********************************************
'*********         Bumpers           ***********
'***********************************************

Sub Bumper1_Hit:vpmTimer.PulseSw 30:RandomSoundBumperTop(Bumper1): End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 31:RandomSoundBumperMiddle(Bumper2): End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 32:RandomSoundBumperBottom(Bumper3): End Sub

'bumper3.EnableSkirtAnimation = True # does not seem to do anything
bumper3.SkirtVisible = False : bumper2.SkirtVisible = False : bumper1.SkirtVisible = False ' if not visible at start they are not animated

Sub bumper3_animate
  dim nrx
  For Each nrx in BP_Bumper3_Ring : nrx.z = bumper3.CurrentRingOffset : Next
  For Each nrx in BP_Bumper3_Socket : nrx.RotX = bumper3.Rotx : nrx.RotY = bumper3.RotY : Next
End Sub

Sub bumper2_animate
  dim nrx
  For Each nrx in BP_Bumper2_Ring : nrx.z = bumper2.CurrentRingOffset : Next
  For Each nrx in BP_Bumper2_Socket : nrx.RotX = bumper2.Rotx : nrx.RotY = bumper2.RotY : Next
End Sub

Sub bumper1_animate
  dim nrx
  For Each nrx in BP_Bumper1_Ring : nrx.z = bumper1.CurrentRingOffset : Next
  For Each nrx in BP_Bumper1_Socket : nrx.RotX = bumper1.Rotx : nrx.RotY = bumper1.RotY : Next
End Sub

Sub BallstopL_Hit: PlaySound "drop_left":End Sub
Sub BallstopR_Hit: PlaySound "drop_Right": End Sub
Sub Gate3_Hit: PlaySound "Gate": End Sub
Sub Gate2_Hit: PlaySound "Gate": End Sub

'' *** Ramp type definitions
'
'Sub bsRampOnWire()
' If bsDict.Exists(ActiveBall.ID) Then
'   bsDict.Item(ActiveBall.ID) = bsWire
' Else
'   bsDict.Add ActiveBall.ID, bsWire
' End If
'End Sub
'
'Sub bsRampOn()
' If bsDict.Exists(ActiveBall.ID) Then
'   bsDict.Item(ActiveBall.ID) = bsRamp
' Else
'   bsDict.Add ActiveBall.ID, bsRamp
' End If
'End Sub
'
'Sub bsRampOnClear()
' If bsDict.Exists(ActiveBall.ID) Then
'   bsDict.Item(ActiveBall.ID) = bsRampClear
' Else
'   bsDict.Add ActiveBall.ID, bsRampClear
' End If
'End Sub
'
'Sub bsRampOff(idx)
' If bsDict.Exists(idx) Then
'   bsDict.Item(idx) = bsNone
' End If
'End Sub
'
'Function getBsRampType(id)
' Dim retValue
' If bsDict.Exists(id) Then
'   retValue = bsDict.Item(id)
' Else
'   retValue = bsNone
' End If
' getBsRampType = retValue
'End Function

'******************************************************
' ZPHY:  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners         https://youtu.be/AXX3aen06FM
' Adding nFozzy roth physics : pt2 flipper physics          https://youtu.be/VSBFuK2RCPE
' Adding nFozzy roth physics : pt3 other elements           https://youtu.be/JN8HEJapCvs
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
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
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

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

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
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 2.7
'   x.AddPt "Polarity", 2, 0.33, - 2.7
'   x.AddPt "Polarity", 3, 0.37, - 2.7
'   x.AddPt "Polarity", 4, 0.41, - 2.7
'   x.AddPt "Polarity", 5, 0.45, - 2.7
'   x.AddPt "Polarity", 6, 0.576, - 2.7
'   x.AddPt "Polarity", 7, 0.66, - 1.8
'   x.AddPt "Polarity", 8, 0.743, - 0.5
'   x.AddPt "Polarity", 9, 0.81, - 0.5
'   x.AddPt "Polarity", 10, 0.88, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03, 0.945
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
'   x.AddPt "Polarity", 2, 0.33, - 3.7
'   x.AddPt "Polarity", 3, 0.37, - 3.7
'   x.AddPt "Polarity", 4, 0.41, - 3.7
'   x.AddPt "Polarity", 5, 0.45, - 3.7
'   x.AddPt "Polarity", 6, 0.576,- 3.7
'   x.AddPt "Polarity", 7, 0.66, - 2.3
'   x.AddPt "Polarity", 8, 0.743, - 1.5
'   x.AddPt "Polarity", 9, 0.81, - 1
'   x.AddPt "Polarity", 10, 0.88, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03, 0.945
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
'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5
'   x.AddPt "Polarity", 2, 0.4, - 5
'   x.AddPt "Polarity", 3, 0.6, - 4.5
'   x.AddPt "Polarity", 4, 0.65, - 4.0
'   x.AddPt "Polarity", 5, 0.7, - 3.5
'   x.AddPt "Polarity", 6, 0.75, - 3.0
'   x.AddPt "Polarity", 7, 0.8, - 2.5
'   x.AddPt "Polarity", 8, 0.85, - 2.0
'   x.AddPt "Polarity", 9, 0.9, - 1.5
'   x.AddPt "Polarity", 10, 0.95, - 1.0
'   x.AddPt "Polarity", 11, 1, - 0.5
'   x.AddPt "Polarity", 12, 1.1, 0
'   x.AddPt "Polarity", 13, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03,  0.945
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

'   copied from Metallica
     x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, - 5.5
        x.AddPt "Polarity", 2, 0.16, - 5.5
        x.AddPt "Polarity", 3, 0.20, - 0.75
        x.AddPt "Polarity", 4, 0.25, - 1.25
        x.AddPt "Polarity", 5, 0.3, - 1.75
        x.AddPt "Polarity", 6, 0.4, - 3.5
        x.AddPt "Polarity", 7, 0.5, - 5.25
        x.AddPt "Polarity", 8, 0.7, - 4.0
        x.AddPt "Polarity", 9, 0.75, - 3.5
        x.AddPt "Polarity", 10, 0.8, - 3.0
        x.AddPt "Polarity", 11, 0.85, - 2.5
        x.AddPt "Polarity", 12, 0.9, - 2.0
        x.AddPt "Polarity", 13, 0.95, - 1.5
        x.AddPt "Polarity", 14, 1, - 1.0
        x.AddPt "Polarity", 15, 1.05, -0.5
        x.AddPt "Polarity", 16, 1.1, 0
        x.AddPt "Polarity", 17, 1.3, 0

        x.AddPt "Velocity", 0, 0, 0.85
        x.AddPt "Velocity", 1, 0.23, 0.85
        x.AddPt "Velocity", 2, 0.35, 0.88
        x.AddPt "Velocity", 3, 0.45, 1
        x.AddPt "Velocity", 4, 0.6, 1 '0.982
        x.AddPt "Velocity", 5, 0.62, 1.0
        x.AddPt "Velocity", 6, 0.702, 0.968
        x.AddPt "Velocity", 7, 0.95,  0.968
        x.AddPt "Velocity", 8, 1.03,  0.945
        x.AddPt "Velocity", 9, 1.5,  0.945

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
  Dim BOT
  BOT = gBOT

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
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
Const EOSTnew = 1 '90's and later
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
' Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.045

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
    Dim b, BOT
    BOT = gBOT

    For b = 0 To UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
      End If
    Next
  End If
End Sub


' adapted by Rothbauerw
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
    ElseIf Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 20 and FlipperPress = 0 Then
            Flipper.endangle = FEndAngle - 3 * Dir
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
  Dim LiveCatchBounce                 'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime
  CatchTime = GameTime - FCount

  If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
    If CatchTime <= LiveCatch * 0.5 Then      'Perfect catch only when catch time happens in the beginning of the window
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
Sub RDampen_Timer : Cor.Update : End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************

'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7  'Level of bounces. Recommmended value of 0.7

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
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
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
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4)

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

Dim RememberBallb : RememberBallb = -1

Sub RollingUpdate()
  Dim b
  Dim BOT
  BOT = gBOT

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 To tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(BOT)
    If BallVel(BOT(b)) > 1 And BOT(b).z < 30 and NOT Controller.Pause Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ <  - 1 And BOT(b).z < 55 And BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz >  - 7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If

    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If

    '***Ball Shadows and ball color ***
    If BallShadowOn Then
      BallShadow(b).X = BOT(b).X : ballShadow(b).Y = BOT(b).Y + 10
      If gi1br > gi0br And BOT(b).Y < 1000 Then
        BOT(b).Color = RGB(128 + gi1br * 255 + gi0br * 128, 128 + gi0br * 128, 128 + gi0br * 128)
      Else
        If Bot(b).x > 820 and Bot(b).Y >  1850 Then
          BOT(b).Color = RGB(25 + gi0br * 50, 25 + gi0br * 50, 25 + gi0br * 50)
        Else
          BOT(b).Color = RGB(100 + gi0br * 200, 100 + gi0br * 200, 100 + gi0br * 200)
        End If
      End If
      If BOT(b).Z > 24 and BOT(b).Z < 35 then
        BallShadow(b).visible = 1
      Else
        BallShadow(b).visible = 0
      End If
      If Not IsEmpty(CBBall) Then
        If CBBall.ID = Bot(b).ID Then RememberBallb = b : Bot(b).color = RGB(50,50,50)
      Else
        If RememberBallb > -1 Then RememberBallb = -1
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

Sub RampRoll_Timer() : RampRollUpdate : End Sub

Sub RampRollUpdate()  'Timer update
  Dim x
  For x = 1 To UBound(RampBalls)
    If Not IsEmpty(RampBalls(x,1) ) Then 'and NOT controller.pause Then
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
      If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then  'if ball is on the PF, remove  it
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

'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'  Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
' - Nudging, plunger, coin-in, start button sounds will be added to the  and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
'
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1       https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2       https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3       https://youtu.be/ESXWGJZY_EI


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
SpinnerSoundLevel = 0.2      'volume level; range [0, 1] #### was 0.5

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

  FlipperCradleCollision ball1, ball2, velocity

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

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

'******************************************************
'   ZTST:  Debug Shot Tester
'******************************************************

'****************************************************************
' Section; Debug Shot Tester v3.2
'
' 1.  Raise/Lower outlanes and drain posts by pressing 2 key
' 2.  Capture and Launch ball, Press and hold one of the buttons (W, E, R, Y, U, I, P, A) below to capture ball by flipper.  Release key to shoot ball
' 3.  To change the test shot angles, press and hold a key and use Flipper keys to adjust the shot angle.  Shot angles are saved into the User direction as cgamename.txt
' 4.  Set DebugShotMode = 0 to disable debug shot test code.
'
' HOW TO INSTALL: Copy all debug* objects from Layer 2 to table and adjust. Copy the Debug Shot Tester code section to the script.
' Add "DebugShotTableKeyDownCheck keycode" to top of Table1_KeyDown sub and add "DebugShotTableKeyUpCheck keycode" to top of Table1_KeyUp sub
'****************************************************************
Const DebugShotMode = 1 'Set to 0 to disable.  1 to enable
Dim DebugKickerForce
DebugKickerForce = 55

' Enable Disable Outlane and Drain Blocker Wall for debug testing
Dim DebugBLState
debug_BLW1.IsDropped = 1
debug_BLP1.Visible = 0
debug_BLR1.Visible = 0
debug_BLW2.IsDropped = 1
debug_BLP2.Visible = 0
debug_BLR2.Visible = 0
'debug_BLW3.IsDropped = 1
debug_BLP3.Visible = 0
debug_BLR3.Visible = 0

Sub BlockerWalls
  DebugBLState = (DebugBLState + 1) Mod 4
  ' debug.print "BlockerWalls"
  PlaySound ("Start_Button")

  Select Case DebugBLState
    Case 0
      debug_BLW1.IsDropped = 1
      debug_BLP1.Visible = 0
      debug_BLR1.Visible = 0
      debug_BLW2.IsDropped = 1
      debug_BLP2.Visible = 0
      debug_BLR2.Visible = 0
      debug_BLW3.IsDropped = 1
      debug_BLP3.Visible = 0
      debug_BLR3.Visible = 0

    Case 1
      debug_BLW1.IsDropped = 0
      debug_BLP1.Visible = 1
      debug_BLR1.Visible = 1
      debug_BLW2.IsDropped = 0
      debug_BLP2.Visible = 1
      debug_BLR2.Visible = 1
      debug_BLW3.IsDropped = 0
      debug_BLP3.Visible = 1
      debug_BLR3.Visible = 1

    Case 2
      debug_BLW1.IsDropped = 0
      debug_BLP1.Visible = 1
      debug_BLR1.Visible = 1
      debug_BLW2.IsDropped = 0
      debug_BLP2.Visible = 1
      debug_BLR2.Visible = 1
      debug_BLW3.IsDropped = 1
      debug_BLP3.Visible = 0
      debug_BLR3.Visible = 0

    Case 3
      debug_BLW1.IsDropped = 1
      debug_BLP1.Visible = 0
      debug_BLR1.Visible = 0
      debug_BLW2.IsDropped = 1
      debug_BLP2.Visible = 0
      debug_BLR2.Visible = 0
      debug_BLW3.IsDropped = 0
      debug_BLP3.Visible = 1
      debug_BLR3.Visible = 1
  End Select
End Sub

Sub DebugShotTableKeyDownCheck (Keycode)
  'Cycle through Outlane/Centerlane blocking posts
  '-----------------------------------------------
  If Keycode = 3 Then
    BlockerWalls
  End If

  If DebugShotMode = 1 Then
    'Capture and launch ball:
    ' Press and hold one of the buttons (W, E, R, T, Y, U, I, P) below to capture ball by flipper.  Release key to shoot ball
    ' To change the test shot angles, press and hold a key and use Flipper keys to adjust the shot angle.
    '--------------------------------------------------------------------------------------------
    If keycode = 17 Then 'W key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleW
    End If
    If keycode = 18 Then 'E key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleE
    End If
    If keycode = 19 Then 'R key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleR
    End If
    If keycode = 21 Then 'Y key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleY
    End If
    If keycode = 22 Then 'U key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleU
    End If
    If keycode = 23 Then 'I key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleI
    End If
    If keycode = 25 Then 'P key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleP
    End If
    If keycode = 30 Then 'A key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleA
    End If
    If keycode = 31 Then 'S key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleS
    End If
    If keycode = 33 Then 'F key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleF
    End If
    If keycode = 34 Then 'G key
      debugKicker.enabled = True
      TestKickerVar = TestKickAngleG
    End If

    If debugKicker.enabled = True Then    'Use Flippers to adjust angle while holding key
      If keycode = LeftFlipperKey Then
        debugKickAim.Visible = True
        TestKickerVar = TestKickerVar - 1
        Debug.print TestKickerVar
      ElseIf keycode = RightFlipperKey Then
        debugKickAim.Visible = True
        TestKickerVar = TestKickerVar + 1
        Debug.print TestKickerVar
      End If
      debugKickAim.ObjRotz = TestKickerVar
    End If
  End If
End Sub


Sub DebugShotTableKeyUpCheck (Keycode)
  ' Capture and launch ball:
  ' Release to shoot ball. Set up angle and force as needed for each shot.
  '--------------------------------------------------------------------------------------------
  If DebugShotMode = 1 Then
    If keycode = 17 Then 'W key
      TestKickAngleW = TestKickerVar
      debugKicker.kick TestKickAngleW, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 18 Then 'E key
      TestKickAngleE = TestKickerVar
      debugKicker.kick TestKickAngleE, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 19 Then 'R key
      TestKickAngleR = TestKickerVar
      debugKicker.kick TestKickAngleR, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 21 Then 'Y key
      TestKickAngleY = TestKickerVar
      debugKicker.kick TestKickAngleY, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 22 Then 'U key
      TestKickAngleU = TestKickerVar
      debugKicker.kick TestKickAngleU, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 23 Then 'I key
      TestKickAngleI = TestKickerVar
      debugKicker.kick TestKickAngleI, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 25 Then 'P key
      TestKickAngleP = TestKickerVar
      debugKicker.kick TestKickAngleP, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 30 Then 'A key
      TestKickAngleA = TestKickerVar
      debugKicker.kick TestKickAngleA, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 31 Then 'S key
      TestKickAngleS = TestKickerVar
      debugKicker.kick TestKickAngleS, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 33 Then 'F key
      TestKickAngleF = TestKickerVar
      debugKicker.kick TestKickAngleF, DebugKickerForce
      debugKicker.enabled = False
    End If
    If keycode = 34 Then 'G key
      TestKickAngleG = TestKickerVar
      debugKicker.kick TestKickAngleG, DebugKickerForce
      debugKicker.enabled = False
    End If

    '   EXAMPLE CODE to set up key to cycle through 3 predefined shots
    '   If keycode = 17 Then   'Cycle through all left target shots
    '     If TestKickerAngle = -28 then
    '       TestKickerAngle = -24
    '     ElseIf TestKickerAngle = -24 Then
    '       TestKickerAngle = -19
    '     Else
    '       TestKickerAngle = -28
    '     End If
    '     debugKicker.kick TestKickerAngle, DebugKickerForce: debugKicker.enabled = false      'W key
    '   End If

  End If

  If (debugKicker.enabled = False And debugKickAim.Visible = True) Then 'Save Angle changes
    debugKickAim.Visible = False
    SaveTestKickAngles
  End If
End Sub

Dim TestKickerAngle, TestKickerAngle2, TestKickerVar, TeskKickKey, TestKickForce
Dim TestKickAngleWDefault, TestKickAngleEDefault, TestKickAngleRDefault, TestKickAngleYDefault, TestKickAngleUDefault, TestKickAngleIDefault
Dim TestKickAnglePDefault, TestKickAngleADefault, TestKickAngleSDefault, TestKickAngleFDefault, TestKickAngleGDefault
Dim TestKickAngleW, TestKickAngleE, TestKickAngleR, TestKickAngleY, TestKickAngleU, TestKickAngleI
Dim TestKickAngleP, TestKickAngleA, TestKickAngleS, TestKickAngleF, TestKickAngleG
TestKickAngleWDefault =  - 27
TestKickAngleEDefault =  - 20
TestKickAngleRDefault =  - 14
TestKickAngleYDefault =  - 8
TestKickAngleUDefault =  - 3
TestKickAngleIDefault = 1
TestKickAnglePDefault = 5
TestKickAngleADefault = 11
TestKickAngleSDefault = 17
TestKickAngleFDefault = 19
TestKickAngleGDefault = 5
If DebugShotMode = 1 Then LoadTestKickAngles

Sub SaveTestKickAngles
  Dim FileObj, OutFile
  Set FileObj = CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then Exit Sub
  Set OutFile = FileObj.CreateTextFile(UserDirectory & cGameName & ".txt", True)

  OutFile.WriteLine TestKickAngleW
  OutFile.WriteLine TestKickAngleE
  OutFile.WriteLine TestKickAngleR
  OutFile.WriteLine TestKickAngleY
  OutFile.WriteLine TestKickAngleU
  OutFile.WriteLine TestKickAngleI
  OutFile.WriteLine TestKickAngleP
  OutFile.WriteLine TestKickAngleA
  OutFile.WriteLine TestKickAngleS
  OutFile.WriteLine TestKickAngleF
  OutFile.WriteLine TestKickAngleG
  OutFile.Close

  Set OutFile = Nothing
  Set FileObj = Nothing
End Sub

Sub LoadTestKickAngles
  Dim FileObj, OutFile, TextStr

  Set FileObj = CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    MsgBox "User directory missing"
    Exit Sub
  End If

  If FileObj.FileExists(UserDirectory & cGameName & ".txt") Then
    Set OutFile = FileObj.GetFile(UserDirectory & cGameName & ".txt")
    Set TextStr = OutFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream = True) Then
      Exit Sub
    End If

    TestKickAngleW = TextStr.ReadLine
    TestKickAngleE = TextStr.ReadLine
    TestKickAngleR = TextStr.ReadLine
    TestKickAngleY = TextStr.ReadLine
    TestKickAngleU = TextStr.ReadLine
    TestKickAngleI = TextStr.ReadLine
    TestKickAngleP = TextStr.ReadLine
    TestKickAngleA = TextStr.ReadLine
    TestKickAngleS = TextStr.ReadLine
    TestKickAngleF = TextStr.ReadLine
    TestKickAngleG = TextStr.ReadLine
    TextStr.Close
  Else
    'create file
    TestKickAngleW = TestKickAngleWDefault
    TestKickAngleE = TestKickAngleEDefault
    TestKickAngleR = TestKickAngleRDefault
    TestKickAngleY = TestKickAngleYDefault
    TestKickAngleU = TestKickAngleUDefault
    TestKickAngleI = TestKickAngleIDefault
    TestKickAngleP = TestKickAnglePDefault
    TestKickAngleA = TestKickAngleADefault
    TestKickAngleS = TestKickAngleSDefault
    TestKickAngleF = TestKickAngleFDefault
    TestKickAngleG = TestKickAngleGDefault
    SaveTestKickAngles
  End If

  Set OutFile = Nothing
  Set FileObj = Nothing

End Sub
'****************************************************************
' End of Section; Debug Shot Tester 3.2
'****************************************************************


' *****************************************
' *** insert lights                   *****
' *****************************************

Dim PFLights(220,5), PFLCount(220), PFInsertOn(220), PFInsertOnX(220), PFFlasher(220), PFInsertState(220), PFInsertOff(220), PFInsertOffX(220), PFsw(220)
Dim PFInsertOffMat(220,2), PfInsertOffTex(220,2), tr(220) : tr(152)=152:tr(153)=152:tr(154)=152:tr(168)=168:tr(169)=168:tr(170)=168:tr(187)=187:tr(188)=187
tr(189)=187:tr(195)=195:tr(196)=195:tr(197)=195:tr(203)=203:tr(204)=203:tr(205)=203:tr(133)=133:tr(134)=133:tr(135)=133:tr(136)=136:tr(137)=136:tr(138)=136
tr(120)=120:tr(121)=120:tr(122)=120

Sub InitLights(aColl, aColl2, aColl3, aColl4) 'InsertOn, InsertLights, InsertOff, InsertFlasher
  Dim obj, idx
  For Each obj In aColl2 : idx = left(obj.name, 3) : Set PFLights(idx, PFLCount(idx)) = obj : PFLCount(idx) = PFLCount(idx) + 1 : Next
  For Each obj In aColl4 : idx = left(obj.name, 3) : Set PFFlasher(idx) = obj : Next
  For Each obj In aColl ' *** InsertOn ***
    idx = left(obj.name, 3) : PfInsertOnX(idx) = obj.BlendDisableLightingFromBelow : obj.BlendDisableLightingFromBelow = 1 : Set PFInsertOn(idx) = obj
    PFInsertOn(idx).BlendDisableLighting = PFInsertOnX(idx)
  Next
  For Each obj In aColl3 '*** InsertOff ***
    idx = left(obj.name, 3) : PfInsertOffX(idx) = obj.BlendDisableLightingFromBelow : obj.BlendDisableLightingFromBelow = -1 : Set PFInsertOff(idx) = obj
    PFInsertOffMat(idx,0) = obj.material : PFInsertOffMat(idx,1) = Left(obj.material, Len(obj.material) - 3) & "on"
    PfInsertOffTex(idx,0) = obj.image : PfInsertOffTex(idx,1) = Left(obj.image, Len(obj.image) - 3) & "on"
  Next
  set PFsw(136)= switch79 : set PFsw(133)= switch80 ' *** rollover switches ***
  PFLights(106,0).state = 0.5 : PFLights(107,0).state = 0.5
End Sub

dim gi0br, gi1br, gi0brfull, gi1brfull : gi0br = 0.5 : gi1br = 0.5 : gi0brfull = 1 : gi1brfull = 1 : rollOversGI
dim lockbarlighton : lockbarlighton = False

Sub LampTimer_Timer()
  Dim ColR2, ColG2, ColB2
  Dim chgLamp, ii, nr, cnt, R, G, B, newcol, onoff,nrx : chgLamp = Controller.ChangedLamps
  If Not IsEmpty(chgLamp) Then
    For cnt = 0 To UBound(chgLamp)
      ii = chglamp(cnt,0) : PFInsertState(ii) = chglamp(cnt,1) / 255
      'ii = 44 : PFInsertState(ii) = rnd(1) ' ### testcode ###
      'If AllInsertsAction > 0 Then PFInsertState(ii) = AllInsertsAction - 1 : onoff = 1 ' ### testcode ###
      Select case ii
        case 152, 153, 154, 168, 169, 170, 187, 188, 189, 195, 196, 197, 203, 204, 205 '*** RGB triangles ***
          ii = tr(ii) : R = PFInsertState(ii) * 255: G = PFInsertState(ii + 1) * 255: B = PFInsertState(ii + 2) * 255: newcol = RGB(R, G, B)
          If R + G + B < 3 Then
            PFInsertOff(ii).material = "inserttriangle3off" : PFFlasher(ii).color = 0 : PFInsertOn(ii).visible = false
            PFInsertOff(ii).BlendDisableLighting = 1 : PFInsertOff(ii).color = RGB(255,255,255) : PFInsertOff(ii).image = "inserttriangle3off"
            PFLights(ii,0).color = 0 : PFLights(ii,0).colorfull = 0 : PFLights(ii,1).color = 0 : PFLights(ii,1).colorfull = 0
          Else
            PFInsertOff(ii).image = "inserttriangle3on" : PFInsertOff(ii).material = "inserttriangle3on" : : PFInsertOn(ii).visible = true
            PFInsertOn(ii).BlendDisableLighting = 800 : PFInsertOn(ii).color = newcol
            PFInsertOff(ii).BlendDisableLighting = 130 : PFInsertOff(ii).color = newcol : PFFlasher(ii).color = newcol
            PFLights(ii,0).color = newcol : PFLights(ii,0).colorfull = newcol : PFLights(ii,1).color = newcol : PFLights(ii,1).colorfull = newcol
          End If
          Select case ii
              case 168 : For Each nr in BL_insrt_168 : nr.color = newcol : next
              case 152 : For Each nr in BL_insrt_152 : nr.color = newcol : next
              case 203 : For Each nr in BL_insrt_203 : nr.color = newcol : next
              case 195 : For Each nr in BL_insrt_195 : nr.color = newcol : next
              case 187 : For Each nr in BL_insrt_187 : nr.color = newcol : next
            End Select
        case 133, 134, 135, 136, 137, 138 ' *** switch rollovers ***
          ii = tr(ii) : R = PFInsertState(ii + 2) * 255 : G = PFInsertState(ii + 1) * 255 : B = PFInsertState(ii) * 255: newcol = RGB(R, G, B)
          If R + G + B < 3 Then
            PFInsertOff(ii).material = "insertrolloveroff" : PFInsertOn(ii).visible = false : PFLights(ii,0).color = 0 : PFLights(ii,0).colorfull = 0
            PFInsertOff(ii).BlendDisableLighting = 7 : PFInsertOff(ii).image = "insertrolloveroff" : PFInsertOff(ii).color = RGB(255,255,255)
            PFInsertOff(ii).BlendDisableLightingFromBelow = -1 : PFsw(ii).image = "switchoff" : PFsw(ii).material = "insertswitchoff"
          Else
            PFInsertOff(ii).image = "insertrolloveron" : PFInsertOff(ii).material = "insertrolloveron" : PFInsertOn(ii).visible = true
            PFInsertOn(ii).BlendDisableLighting = 500 : PFInsertOn(ii).color = newcol
            PFInsertOff(ii).BlendDisableLighting = 3 : PFInsertOff(ii).color = newcol : PFInsertOff(ii).BlendDisableLightingFromBelow = 0.5
            PFLights(ii,0).color = newcol : PFLights(ii,0).colorfull = newcol : PFsw(ii).image = "switchon" : PFsw(ii).material = "insertswitchon"
          End If
          Select case ii
            case 133 : For Each nr in BL_insrt_133 : nr.color = newcol : next
            case 136 : For Each nr in BL_insrt_136 : nr.color = newcol : next
          End Select
        case 106 ' *** white GI ***
          gi0brfull = PFInsertState(ii) : gi0br = gi0brfull ' minimum is 0,0627
          If gi0br > 0.5 and gi1br > 0.5 Then gi0br = 0.5 : gi1br = 0.5
          PFLights(106,0).state = gi0br ^ 0.8 : PFLights(107,0).state = gi1br ^ 0.5
          RolloversGI
        case 107 ' *** red GI ***
          gi1brfull = PFInsertState(ii) : : gi1br = gi1brfull
          If gi0br > 0.5 and gi1br > 0.5 Then gi0br = 0.5 : gi1br = 0.5
          PFLights(106,0).state = gi0br ^ 0.8 : PFLights(107,0).state = gi1br ^ 0.5
          RolloversGI
        case 3,4,15,16,17,18,19,20,21,23,31,32,34,35,38,39,43,44,46,47,48,49,50,52,53,58,59,64,65,66,67,68,69,72
          If PFInsertState(ii) < 0.01 then onoff = 0 else onoff = 1
          For nr = 1 to PFLCount(ii) : PFLights(ii,nr - 1).state = PFInsertState(ii) : Next
          PFInsertOff(ii).BlendDisableLightingFromBelow = -1 + 2 * PFInsertState(ii)
          PFInsertOff(ii).BlendDisableLighting = 1 + PFInsertState(ii) * (PFInsertOffX(ii) -1)
          PFInsertOff(ii).material = PFInsertOffMat(ii,onoff) : PFInsertOff(ii).image = PFInsertOffTex(ii,onoff)
          PFInsertOn(ii).visible = onoff : PFInsertOn(ii).BlendDisableLighting = PFInsertOnX(ii) * PFInsertState(ii)
          If onoff = 0 Then
            PFFlasher(ii).imageA = "insertshadows" : PFFlasher(ii).AddBlend = False : PFFlasher(ii).height = -0.2 : PFFlasher(ii).Opacity = 100
          Else
            PFFlasher(ii).imageA = "flashertex" : PFFlasher(ii).AddBlend = True : PFFlasher(ii).height = 0.2 : PFFlasher(ii).Opacity = 1000 * PFInsertState(ii)
          End If
        case 55,56
          If PFInsertState(ii) < 0.01 then onoff = 0 else onoff = 1
          For nr = 1 to PFLCount(ii) : PFLights(ii,nr - 1).state = PFInsertState(ii) : Next
          PFInsertOff(ii).BlendDisableLightingFromBelow = -1 + 2 * PFInsertState(ii)
          PFInsertOff(ii).BlendDisableLighting = 1 + PFInsertState(ii) * (PFInsertOffX(ii) -1)
          PFInsertOff(ii).material = PFInsertOffMat(ii,onoff) : PFInsertOff(ii).image = PFInsertOffTex(ii,onoff)
          PFInsertOn(ii).visible = onoff : PFInsertOn(ii).BlendDisableLighting = PFInsertOnX(ii) * PFInsertState(ii)
        case 5,6,7,8,9,10,11,12,13,14,22,25,26,27,28,29,30,37,40,42,45,51,54,63,76,77
          If PFInsertState(ii) < 0.01 then onoff = 0 else onoff = 1
          For nr = 1 to PFLCount(ii) : PFLights(ii,nr - 1).state = PFInsertState(ii) : Next
          PFInsertOn(ii).visible = onoff : PFInsertOn(ii).BlendDisableLighting = PFInsertOnX(ii) * PFInsertState(ii)
          If onoff = 0 Then
            PFFlasher(ii).imageA = "insertshadows" : PFFlasher(ii).AddBlend = False : PFFlasher(ii).height = -0.2 : PFFlasher(ii).Opacity = 100
          Else
            PFFlasher(ii).imageA = "flashertex" : PFFlasher(ii).AddBlend = True : PFFlasher(ii).height = 0.2 : PFFlasher(ii).Opacity = 1000 * PFInsertState(ii)
          End If
        case 60,61,62,70,71,73,74,75,78
          PFLights(ii,0).state = PFInsertState(ii)
        case 120, 121, 122 ' *** lockbar fire button ***
          ii = tr(ii) : B = PFInsertState(ii) * 255: G = PFInsertState(ii + 1) * 255: R = PFInsertState(ii + 2) * 255 : newcol = RGB(R, G, B)
          If R+G+B > 0 then firebuttonon.opacity = 765 * 200 / (R+G+B) : firebuttonon001.opacity = 765 * 40 / (R+G+B)
          If SpareFireButton Then
            If R + G + B < 3 Then
              firebuttonon001.visible = 0
            Else
              firebuttonon001.visible = 1 : firebuttonon001.color = newcol : lockbarlighton = True
            End If
          Elseif firebuttonon001.visible Then firebuttonon001.visible = 0 : lockbarlighton = False
          End If
          If RailsLockbar.visible Then
            If R + G + B < 3 Then
              firebuttonon.visible = 0 : firebuttonlight.state = 0
            Else
              firebuttonon.visible = 1 : firebuttonlight.state = 1
              firebuttonon.color = newcol : firebuttonlight.color = newcol
            End If
                    End If
        Case 1 ' *** start button ***
          startbuttonon.opacity = PFInsertState(ii) * 400
        Case 2 ' *** tournament button ***
          tourbuttonon.opacity = PFInsertState(ii) * 300
'       case else
      End Select
    Next
  End If
  RollingUpdate
  DoDTAnim    'handle drop target animations
  DoSTAnim    'handle stand up target animations
  DoKTAnim    'sw38 animation
End Sub

Sub RolloversGI()
  dim colR, ColG, ColB, ColR2, ColG2, ColB2, colRGB
  If gi0br > 0 Then
    If gi1br >= gi0br Then ColR = gi1br : ColG = gi0br * 0.5 : ColB = gi0br * 0.5 else ColR = gi0br : ColG = gi0br : ColB = gi0br
  Else
    If gi1br > 0 Then colR = gi1br : ColG = 0 : ColB = 0 else ColR = 0.10 : colG = 0.10 : colB = 0.10
  End If
  MaterialColor "rollovers", RGB(10 + colR * 245, 10 + ColG * 245, 10 + ColB * 245)
  MaterialColor "rolloverstop", RGB(128 + colR * 128, 10 + ColG * 245, 10 + ColB * 245)
  MaterialColor "rolloversalt", RGB(8 + colR * 48, 8 + ColG * 48, 8 + ColB * 48)
  MaterialColor "slings", RGB(64 + ColR * 255, 64 + ColG * 192, 64 + ColB * 192)
  If gi0brfull >= gi1brfull Then
    GiGlow.color = RGB(gi0brfull * 100,gi0brfull * 100, gi0brfull * 100)
  Else
    If gi1brfull^0.5 < 0.3 Then
      GiGlow.color = RGB(0,0,0)
    Else
      GiGlow.color = RGB(gi1brfull^0.5 * 255 ,0,0)
    End If
  End If
  If gi0brfull^0.8 > gi1brfull^0.5 - 0.1 Then
    MaterialColor "insertswitchoff", RGB(7 + 180 * gi0brfull, 7 + 200 * gi0brfull, 7 + 248 * gi0brfull)
    MaterialColor "insertrolloveroff", RGB(7 + 90 * gi0brfull, 7 + 100 * gi0brfull, 7 + 128 * gi0brfull)
  Else
    MaterialColor "insertswitchoff", RGB(7 + 100 * gi1brfull, 7 + 100 * gi0brfull, 7 + 100 * gi0brfull)
    MaterialColor "insertrolloveroff", RGB(7 + 100 * gi1brfull, 7 + 100 * gi0brfull, 7 + 100 * gi0brfull)
  End If
End Sub

Sub initVLM()
  dim nrx
  BM_wellzom.color = RGB(192,192,192)
  For Each nrx in BL_GI0 : nrx.color = RGB(150,180,255) : nrx.opacity = 200 : next : LM_GI0_Wellzom.opacity = 150 : LM_GI0_Playfield.opacity = 250
  For Each nrx in BL_GI1 : nrx.color = RGB(255,0,0) : nrx.opacity = 100 : next : LM_GI1_Wellzom.opacity = 75 : LM_GI1_Playfield.opacity = 120 : LM_GI1_Wellzom.color = RGB(255,32,32)
  LM_GI1_armp0.color = RGB(255,32,32) : LM_GI1_armp1.color = RGB(255,32,32) : LM_GI1_armp2.color = RGB(255,32,32): LM_GI1_armp3.color = RGB(255,32,32)
  LM_GI1_priszom.visible = 0 : LM_GI0_priszom.visible = 0
  'LM_diode_Parts.opacity = 200
  For Each nrx in BL_flsh19a : nrx.color = RGB(255,0,0) : nrx.opacity = 200 : next : LM_flsh19a_wellzom.opacity = 150
  For Each nrx in BL_flsh20 : nrx.opacity = 600 * FlasherIntensity : nrx.color = RGB(180,200,255) : next
  For Each nrx in BL_flsh25 : nrx.opacity = 200 * FlasherIntensity : next
  For Each nrx in BL_flsh26 : nrx.color = RGB(100,160,255) : nrx.opacity = 2800 * FlasherIntensity : next
  For Each nrx in BL_flsh27 : nrx.color = RGB(255,0,0) : nrx.opacity = 800 * FlasherIntensity : next : LM_flsh27_Playfield.opacity = 400 * FlasherIntensity
  For Each nrx in BL_flsh28 : nrx.color = RGB(255,0,0) : nrx.opacity = 300 * FlasherIntensity : next : LM_flsh28_Playfield.opacity = 95 * FlasherIntensity : LM_flsh28_armp0.opacity = 150 * FlasherIntensity : LM_flsh28_BotRamp.opacity = 100 * FlasherIntensity : LM_flsh28_wellzom.opacity = 100 * FlasherIntensity : LM_flsh28_wellzombase.opacity = 100 * FlasherIntensity
  For Each nrx in BL_flsh29 : nrx.color = RGB(255,0,0) : nrx.opacity = 300 * FlasherIntensity : next : LM_flsh29_Playfield.opacity = 95 * FlasherIntensity : LM_flsh29_armp3.opacity = 150 * FlasherIntensity : LM_flsh29_wellzom.opacity = 100 * FlasherIntensity : LM_flsh29_doorleft.opacity = 600 * FlasherIntensity : LM_flsh29_doorright.opacity = 600 * FlasherIntensity : LM_flsh29_psw48.opacity = 50 * FlasherIntensity
  For Each nrx in BL_flsh31 : nrx.opacity = 400 * FlasherIntensity : nrx.color = RGB(180,200,255) : next
  For each nrx in BL_insrt_003: nrx.opacity = 150 : Next
  LM_insrt_022_Parts.opacity = 100 : LM_insrt_022_armp0.opacity = 400
  LM_insrt_023_armp0.opacity = 300
  LM_insrt_030_Parts.opacity = 300
  LM_insrt_031_Parts.opacity = 300 : LM_insrt_031_armp0.opacity = 400 : LM_insrt_031_Spin2.opacity = 300
  LM_insrt_032_armp0.opacity = 300 : LM_insrt_032_Spin2.opacity = 500
  LM_insrt_037_armp3.opacity = 300
  LM_insrt_038_armp3.opacity = 300
  LM_insrt_039_Parts.opacity = 200 : LM_insrt_039_Spin1.opacity = 200
  LM_insrt_040_armp3.opacity = 400 : LM_insrt_040_Parts.opacity = 200
  LM_insrt_050_wellzom.opacity = 150
  LM_insrt_051_Parts.opacity = 150 : LM_insrt_051_wellzom.opacity = 200 : LM_insrt_051_wellzombase.opacity = 200
  LM_insrt_054_wellzom.opacity = 500 : LM_insrt_054_wellzombase.opacity = 500
  LM_insrt_055_Parts.opacity = 300 : LM_insrt_055_target1.opacity = 300
  LM_insrt_056_target2.opacity = 300
  LM_insrt_058_Parts.opacity = 300
  LM_insrt_059_Parts.opacity = 500 :  LM_insrt_059_target2.opacity = 300
  For Each nrx in BL_insrt60 : nrx.opacity = 1200 : next : LM_insrt60_armp2.opacity = 2000 : LM_insrt60_doorright.opacity = 2000
  For Each nrx in BL_insrt61 : nrx.opacity = 1200 : next : LM_insrt61_armp2.opacity = 3000 :  LM_insrt61_armp3.opacity = 3000
  For Each nrx in BL_insrt62 : nrx.opacity = 1200 : next
  LM_insrt_063_Parts.opacity = 500
  For each nrx in BL_insrt_070: nrx.opacity = 300 : Next
  For each nrx in BL_insrt_071: nrx.opacity = 300 : Next
  LM_insrt_071_wellzom.opacity = 500 : LM_insrt_071_armp3.opacity = 600
  LM_insrt_072_Parts.opacity = 300
  For each nrx in BL_insrt_073: nrx.opacity = 150 : Next
  For each nrx in BL_insrt_074: nrx.opacity = 150 : Next
  For each nrx in BL_insrt_075: nrx.opacity = 150 : Next
  LM_insrt_076_armp2.opacity = 200 : LM_insrt_076_Parts.opacity = 200 : LM_insrt_076_rampscrw.opacity = 5000
  LM_insrt_077_Parts.opacity = 200 : LM_insrt_077_armp2.opacity = 200 : LM_insrt_077_target3.opacity = 200
  LM_insrt_078_armp0.opacity = 700 : LM_insrt_078_Parts.opacity = 200
  For Each nrx in BL_insrt78 : nrx.color = RGB(0,0,0) : nrx.opacity = 1200 * FlasherIntensity : next : LM_insrt78_biczom.opacity = 900 * FlasherIntensity
  For Each nrx in BL_insrt_133 : nrx.color = 0 : next : LM_insrt_133_Parts.opacity = 100:  LM_insrt_133_Playfield.opacity = 600
  For Each nrx in BL_insrt_136 : nrx.color = 0 : next : LM_insrt_136_Parts.opacity = 100:  LM_insrt_136_Playfield.opacity = 600
  For Each nrx in BL_insrt_152 : nrx.color = 0 : next : LM_insrt_152_non_opaque.opacity = 300
  For Each nrx in BL_insrt_168 : nrx.color = 0 : nrx.opacity = 200 : next
  For Each nrx in BL_insrt_187 : nrx.color = 0 : nrx.opacity = 200 : next
  For Each nrx in BL_insrt_195 : nrx.color = 0 : nrx.opacity = 200 : next
  For Each nrx in BL_insrt_203 : nrx.color = 0 : nrx.opacity = 200 : next
  For Each nrx in BP_BotRamp : nrx.DepthBias = -100 : next
  For Each nrx in BP_biczom : nrx.ObjRotZ = -11 : nrx.RotZ = 0 : Next
  For Each nrx in BP_BatLeft : nrx.ObjRotZ = 0 : next
  For Each nrx in BP_BatRight : nrx.ObjRotZ = 0 : next
  LM_GI1_crossbow.opacity = 60 : LM_GI0_crossbow.opacity = 60
  LM_GI1_crossbowbott.opacity = 60 : LM_GI0_crossbowbott.opacity = 60
  For Each nrx in BP_armp0 : nrx.depthbias = -1500 :  next
  For Each nrx in BP_armp1 : nrx.depthbias = -1500 :  next
  For Each nrx in BP_armp2 : nrx.depthbias = -1500 :  next
  For Each nrx in BP_armp3 : nrx.depthbias = -1500 :  next
  'BM_parts.ReflectionEnabled = True
  'BM_Playfield.ReflectionProbe = "Playfield Reflections"
  For Each nrx in BP_rampscrw : nrx.depthbias = -10000 : next

  ' #### for testing only ####
' For Each nrx in BL_GI1 : nrx.color = RGB(0,0,0) : nrx.opacity = 0 : next
' For Each nrx in BL_GI0 : nrx.color = RGB(0,0,0) : nrx.opacity = 0 : next
' For Each nrx in BL_DIODE : nrx.color = RGB(0,0,0) : nrx.opacity = 0 : next

End Sub


' ************************
' *** f12 options menu ***
' ************************

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

Dim VRRoomChoice : VRRoomChoice = 3       ' 1 - Cab Only, 2 - Minimal Room, 3 - MEGA room
Dim dspTriggered : dspTriggered = False

Dim EnableRaytracedBallShadows : EnableRaytracedBallShadows = 1
Dim FlasherIntensity : FlasherIntensity = 1

Dim VolumeDial : VolumeDial = 0.8     'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5   'Level of ramp rolling volume. Value between 0 and 1
Dim SpareFireButton : SpareFireButton = 0

Sub Table1_OptionEvent(ByVal eventId)

  If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  Dim nrx, lightlevel
  LightLevel = NightDay/100
  If LightLevel < 0.37 Then LightLevel = 0.37
  nrx = (LightLevel - 0.37) * 235 + 20
  MaterialColor "VLM.Bake.Active", RGB(nrx, nrx, nrx)
  MaterialColor "VLM.Bake.Solid", RGB(nrx, nrx, nrx)
  MaterialColor "VLM.Bake.Ramps", RGB(nrx, nrx, nrx)
  MaterialColor "VLM.Bake.crossbow", RGB(nrx + 22, nrx + 22, nrx + 22)
  GiGlow.opacity = 500 * (0.5 - LightLevel) * 7.7

  ' VRRoom
  VRRoomChoice = Table1.Option("VR Room", 0, 3, 1, 3, 0, Array("Off", "CabOnly", "Minimal", "MEGA"))
  LoadVRRoom

  ' VRCab Art Selection
  Dim VRCabArtOpt : VRCabArtOpt = 1
  VRCabArtOpt = Table1.Option("VRCabArtOpt", 1, 2, 1, 2, 0, Array("Limited Edition", "Premium"))
  Select Case VRCabArtOpt
    Case 1:
      VRCab_Cabinet.image = "VR_CabinetNEW_1"
      VRCab_Backbox.image = "VR_BackboxNew_1"
    Case 2:
      VRCab_Cabinet.image = "VRCab_CabinetPremium"
      VRCab_Backbox.image = "VRCab_Backbox_Premium"
  End Select

  ' VRBackglass Art Selection
  Dim VRBackglassArtOpt : VRBackglassArtOpt = 1
  VRBackglassArtOpt = Table1.Option("VRBackglassArtOpt", 1, 3, 1, 2, 0, Array("Limited Edition", "Premium", "Cartoon"))
  Select Case VRBackglassArtOpt
    Case 1:
      VRCab_Backglass.image = "BackglassImage"
    Case 2:
      VRCab_Backglass.image = "BackglassimagePremium"
    Case 3:
      VRCab_Backglass.image = "BackglassImageCartoon"
  End Select

  Dim EnableRailsLockbar : EnableRailsLockbar = 1
  EnableRailsLockbar = Table1.Option("Enable rails & lockbar for fullscreen", 1, 2, 1, 2, 0, Array("Enabled", "Disabled"))
    Select Case EnableRailsLockbar
    Case 1:
      RailsLockbar.visible = 1
    Case 2:
      If not DesktopMode and not FSSMode and RenderingMode<>2 Then
        RailsLockbar.visible = 0 : firebuttonon.visible = 0 : firebuttonlight.state = 0 : firebuttonoff.visible = 0
      End If
  End Select

  Dim GlassChoice : GlassChoice = 1
  GlassChoice = Table1.Option("Playfield glass type", 1,2, 1, 1, 0, Array("normal", "bloody"))
  If GlassChoice = 1 Then
    Primitive001.image = "glassnormal"
  Else
    Primitive001.image = "glassblood"
  End If

  Dim GlassVisible : GlassVisible = 1
  GlassVisible = Table1.Option("Playfield glass visibility", 1,6, 1, 2, 0, Array("100%", "off", "20%", "40%", "60%", "80%"))
  Primitive001.opacity = 110
  ' Select case GlassVisible
'   Case 1 : Primitive001.color = RGB(255,255,255) : Primitive001.visible = 1
'   Case 2 : Primitive001.visible = 0
'   Case 3 : Primitive001.color = RGB(50,50,50) : Primitive001.visible = 1
'   Case 4 : Primitive001.color = RGB(100,100,100) : Primitive001.visible = 1
'   Case 5 : Primitive001.color = RGB(150,150,150) : Primitive001.visible = 1
'   Case 6 : Primitive001.color = RGB(200,200,200) : Primitive001.visible = 1
' End Select
  Select case GlassVisible
    Case 1 : Primitive001.blenddisablelighting = 0.25 : Primitive001.visible = 1
    Case 2 : Primitive001.visible = 0
    Case 3 : Primitive001.blenddisablelighting = 0.02 : Primitive001.visible = 1
    Case 4 : Primitive001.blenddisablelighting = 0.05 : Primitive001.visible = 1
    Case 5 : Primitive001.blenddisablelighting = 0.10 : Primitive001.visible = 1
    Case 6 : Primitive001.blenddisablelighting = 0.20 : Primitive001.visible = 1
  End Select

  Dim LutChoice : LutChoice = 1
  LutChoice = Table1.Option("LUT",1,5,1,1,0, Array("Daryl (default)", "Carl", "Rick", "Michonne", "Eugene"))
  Select case LutChoice
    Case 1 : table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.7vibr100"
    Case 2 : table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.85vibr100"
    Case 3 : table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50"
    Case 4 : table1.ColorGradeImage = "ColorGradeLUT256x16_1to1"
    Case 5 : table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gamma0.70"
  End Select

  EnableRaytracedBallShadows = Table1.Option("Ball raytraced shadows", 1, 2, 1, 1, 0, Array("Enabled", "Disabled"))
  If EnableRaytracedBallShadows  = 2 Then
    Light004.state = 0 : Light001.state = 0 : Light005.state = 0 : Light006.state = 0
  Else
    Light004.state = 1 : Light001.state = 1 : Light005.state = 1 : Light006.state = 1
  End If

  Dim FlashOld : FlashOld = FlasherIntensity
  FlasherIntensity = Table1.Option("Flasher intensity (default=100%)", 0.3, 2, 0.1, 1, 1)', Array("Enabled", "Disabled"))
  If FlashOld <> FlasherIntensity Then
    initVLM : flashintensity = 1 : FlashersOff.enabled = true
  End If

  SpareFireButton = Table1.Option("Show fire/lockbar button on apron", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))
  If lockbarlighton Then
    firebuttonon001.visible = SpareFireButton
  End If

  VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Dim flashintensity
Sub FlashersOff_Timer()
  f29.state  = flashintensity : f28.state = flashintensity : f27.state = flashintensity
  flashintensity = flashintensity - 0.1
  If flashintensity < -0.01 Then FlashersOff.enabled = False
End Sub


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BatLeft: BP_BatLeft=Array(BM_BatLeft, LM_GI1_BatLeft, LM_GI0_BatLeft, LM_flsh28_BatLeft, LM_flsh29_BatLeft)
Dim BP_BatRight: BP_BatRight=Array(BM_BatRight, LM_GI1_BatRight, LM_GI0_BatRight, LM_flsh28_BatRight)
Dim BP_BotRamp: BP_BotRamp=Array(BM_BotRamp, LM_insrt_032_BotRamp, LM_insrt78_BotRamp, LM_insrt_203_BotRamp, LM_GI1_BotRamp, LM_GI0_BotRamp, LM_flsh27_BotRamp, LM_flsh28_BotRamp, LM_flsh29_BotRamp)
Dim BP_Bumper1_Ring: BP_Bumper1_Ring=Array(BM_Bumper1_Ring, LM_flsh19a_Bumper1_Ring, LM_insrt60_Bumper1_Ring, LM_insrt61_Bumper1_Ring, LM_insrt62_Bumper1_Ring, LM_insrt_168_Bumper1_Ring, LM_GI1_Bumper1_Ring, LM_GI0_Bumper1_Ring, LM_diode_Bumper1_Ring, LM_flsh25_Bumper1_Ring, LM_flsh26_Bumper1_Ring, LM_flsh27_Bumper1_Ring, LM_flsh28_Bumper1_Ring, LM_flsh29_Bumper1_Ring)
Dim BP_Bumper1_Socket: BP_Bumper1_Socket=Array(BM_Bumper1_Socket, LM_insrt61_Bumper1_Socket, LM_insrt62_Bumper1_Socket, LM_insrt_168_Bumper1_Socket, LM_GI1_Bumper1_Socket, LM_GI0_Bumper1_Socket, LM_flsh25_Bumper1_Socket, LM_flsh26_Bumper1_Socket, LM_flsh27_Bumper1_Socket, LM_flsh28_Bumper1_Socket, LM_flsh29_Bumper1_Socket)
Dim BP_Bumper2_Ring: BP_Bumper2_Ring=Array(BM_Bumper2_Ring, LM_insrt60_Bumper2_Ring, LM_insrt61_Bumper2_Ring, LM_insrt62_Bumper2_Ring, LM_GI1_Bumper2_Ring, LM_GI0_Bumper2_Ring, LM_flsh25_Bumper2_Ring, LM_flsh27_Bumper2_Ring, LM_flsh28_Bumper2_Ring, LM_flsh29_Bumper2_Ring)
Dim BP_Bumper2_Socket: BP_Bumper2_Socket=Array(BM_Bumper2_Socket, LM_insrt60_Bumper2_Socket, LM_insrt61_Bumper2_Socket, LM_GI1_Bumper2_Socket, LM_GI0_Bumper2_Socket, LM_flsh25_Bumper2_Socket, LM_flsh27_Bumper2_Socket, LM_flsh28_Bumper2_Socket, LM_flsh29_Bumper2_Socket)
Dim BP_Bumper3_Ring: BP_Bumper3_Ring=Array(BM_Bumper3_Ring, LM_insrt60_Bumper3_Ring, LM_insrt61_Bumper3_Ring, LM_insrt62_Bumper3_Ring, LM_GI1_Bumper3_Ring, LM_GI0_Bumper3_Ring, LM_flsh25_Bumper3_Ring, LM_flsh26_Bumper3_Ring, LM_flsh27_Bumper3_Ring, LM_flsh28_Bumper3_Ring, LM_flsh29_Bumper3_Ring)
Dim BP_Bumper3_Socket: BP_Bumper3_Socket=Array(BM_Bumper3_Socket, LM_insrt60_Bumper3_Socket, LM_insrt61_Bumper3_Socket, LM_insrt62_Bumper3_Socket, LM_GI1_Bumper3_Socket, LM_GI0_Bumper3_Socket, LM_flsh25_Bumper3_Socket, LM_flsh26_Bumper3_Socket, LM_flsh27_Bumper3_Socket, LM_flsh28_Bumper3_Socket, LM_flsh29_Bumper3_Socket)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_flsh19a_Parts, LM_insrt_022_Parts, LM_insrt_023_Parts, LM_insrt_030_Parts, LM_insrt_031_Parts, LM_insrt_032_Parts, LM_insrt_037_Parts, LM_insrt_038_Parts, LM_insrt_039_Parts, LM_insrt_040_Parts, LM_insrt_051_Parts, LM_insrt_052_Parts, LM_insrt_053_Parts, LM_insrt_055_Parts, LM_insrt_058_Parts, LM_insrt_059_Parts, LM_insrt60_Parts, LM_insrt61_Parts, LM_insrt62_Parts, LM_insrt_063_Parts, LM_insrt_070_Parts, LM_insrt_071_Parts, LM_insrt_072_Parts, LM_insrt_073_Parts, LM_insrt_074_Parts, LM_flsh20_Parts, LM_flsh31_Parts, LM_insrt_075_Parts, LM_insrt_076_Parts, LM_insrt_077_Parts, LM_insrt78_Parts, LM_insrt_078_Parts, LM_insrt_133_Parts, LM_insrt_136_Parts, LM_insrt_152_Parts, LM_insrt_168_Parts, LM_insrt_187_Parts, LM_insrt_195_Parts, LM_insrt_203_Parts, LM_GI1_Parts, LM_GI0_Parts, LM_diode_Parts, LM_flsh25_Parts, LM_flsh26_Parts, LM_flsh27_Parts, LM_flsh28_Parts, LM_flsh29_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_flsh19a_Playfield, LM_insrt60_Playfield, LM_insrt61_Playfield, LM_insrt62_Playfield, LM_flsh20_Playfield, LM_flsh31_Playfield, LM_insrt78_Playfield, LM_insrt_133_Playfield, LM_insrt_136_Playfield, LM_GI1_Playfield, LM_GI0_Playfield, LM_diode_Playfield, LM_flsh25_Playfield, LM_flsh26_Playfield, LM_flsh27_Playfield, LM_flsh28_Playfield, LM_flsh29_Playfield)
Dim BP_Spin1: BP_Spin1=Array(BM_Spin1, LM_insrt_039_Spin1, LM_insrt_040_Spin1, LM_insrt_053_Spin1, LM_flsh20_Spin1, LM_insrt_152_Spin1, LM_GI1_Spin1, LM_GI0_Spin1, LM_flsh28_Spin1, LM_flsh29_Spin1)
Dim BP_Spin2: BP_Spin2=Array(BM_Spin2, LM_insrt_031_Spin2, LM_insrt_032_Spin2, LM_flsh31_Spin2, LM_insrt78_Spin2, LM_insrt_195_Spin2, LM_insrt_203_Spin2, LM_GI1_Spin2, LM_GI0_Spin2, LM_flsh27_Spin2, LM_flsh28_Spin2, LM_flsh29_Spin2)
Dim BP_armp0: BP_armp0=Array(BM_armp0, LM_insrt_022_armp0, LM_insrt_023_armp0, LM_insrt_031_armp0, LM_insrt_032_armp0, LM_flsh31_armp0, LM_insrt78_armp0, LM_insrt_078_armp0, LM_GI1_armp0, LM_GI0_armp0, LM_flsh26_armp0, LM_flsh27_armp0, LM_flsh28_armp0, LM_flsh29_armp0)
Dim BP_armp1: BP_armp1=Array(BM_armp1, LM_GI1_armp1, LM_GI0_armp1, LM_flsh26_armp1, LM_flsh27_armp1, LM_flsh28_armp1, LM_flsh29_armp1)
Dim BP_armp2: BP_armp2=Array(BM_armp2, LM_insrt60_armp2, LM_insrt61_armp2, LM_insrt_076_armp2, LM_insrt_077_armp2, LM_GI1_armp2, LM_GI0_armp2, LM_flsh25_armp2, LM_flsh26_armp2, LM_flsh27_armp2, LM_flsh28_armp2, LM_flsh29_armp2)
Dim BP_armp3: BP_armp3=Array(BM_armp3, LM_insrt_037_armp3, LM_insrt_038_armp3, LM_insrt_040_armp3, LM_insrt61_armp3, LM_insrt_071_armp3, LM_flsh20_armp3, LM_GI1_armp3, LM_GI0_armp3, LM_flsh25_armp3, LM_flsh26_armp3, LM_flsh27_armp3, LM_flsh28_armp3, LM_flsh29_armp3)
Dim BP_biczom: BP_biczom=Array(BM_biczom, LM_insrt78_biczom, LM_insrt_078_biczom, LM_GI1_biczom, LM_GI0_biczom, LM_flsh27_biczom, LM_flsh28_biczom)
Dim BP_crossbow: BP_crossbow=Array(BM_crossbow, LM_GI1_crossbow, LM_GI0_crossbow, LM_flsh28_crossbow, LM_flsh29_crossbow)
Dim BP_crossbowbott: BP_crossbowbott=Array(BM_crossbowbott, LM_GI1_crossbowbott, LM_GI0_crossbowbott, LM_flsh27_crossbowbott, LM_flsh28_crossbowbott, LM_flsh29_crossbowbott)
Dim BP_doorleft: BP_doorleft=Array(BM_doorleft, LM_GI1_doorleft, LM_GI0_doorleft, LM_diode_doorleft, LM_flsh26_doorleft, LM_flsh27_doorleft, LM_flsh28_doorleft, LM_flsh29_doorleft)
Dim BP_doorright: BP_doorright=Array(BM_doorright, LM_insrt60_doorright, LM_GI1_doorright, LM_GI0_doorright, LM_diode_doorright, LM_flsh26_doorright, LM_flsh27_doorright, LM_flsh28_doorright, LM_flsh29_doorright)
Dim BP_non_opaque: BP_non_opaque=Array(BM_non_opaque, LM_insrt_053_non_opaque, LM_insrt_073_non_opaque, LM_insrt_074_non_opaque, LM_flsh20_non_opaque, LM_insrt_075_non_opaque, LM_insrt_152_non_opaque, LM_GI1_non_opaque, LM_GI0_non_opaque, LM_flsh25_non_opaque, LM_flsh27_non_opaque, LM_flsh28_non_opaque, LM_flsh29_non_opaque)
Dim BP_priszom: BP_priszom=Array(BM_priszom, LM_GI1_priszom, LM_GI0_priszom, LM_diode_priszom, LM_flsh26_priszom, LM_flsh27_priszom, LM_flsh28_priszom, LM_flsh29_priszom)
Dim BP_psw10: BP_psw10=Array(BM_psw10, LM_GI1_psw10, LM_GI0_psw10, LM_flsh28_psw10, LM_flsh29_psw10)
Dim BP_psw11: BP_psw11=Array(BM_psw11, LM_GI1_psw11, LM_GI0_psw11, LM_flsh28_psw11, LM_flsh29_psw11)
Dim BP_psw48: BP_psw48=Array(BM_psw48, LM_GI1_psw48, LM_GI0_psw48, LM_flsh28_psw48, LM_flsh29_psw48)
Dim BP_psw9: BP_psw9=Array(BM_psw9, LM_GI1_psw9, LM_GI0_psw9, LM_flsh28_psw9, LM_flsh29_psw9)
Dim BP_rampscrw: BP_rampscrw=Array(BM_rampscrw, LM_insrt_038_rampscrw, LM_insrt_063_rampscrw, LM_insrt_076_rampscrw, LM_insrt78_rampscrw, LM_insrt_078_rampscrw, LM_GI1_rampscrw, LM_GI0_rampscrw, LM_flsh25_rampscrw, LM_flsh26_rampscrw, LM_flsh27_rampscrw, LM_flsh28_rampscrw, LM_flsh29_rampscrw)
Dim BP_target1: BP_target1=Array(BM_target1, LM_insrt_055_target1, LM_GI1_target1, LM_GI0_target1, LM_flsh27_target1, LM_flsh28_target1)
Dim BP_target2: BP_target2=Array(BM_target2, LM_insrt_056_target2, LM_insrt_059_target2, LM_insrt_063_target2, LM_insrt78_target2, LM_GI1_target2, LM_GI0_target2, LM_flsh26_target2, LM_flsh27_target2, LM_flsh28_target2, LM_flsh29_target2)
Dim BP_target3: BP_target3=Array(BM_target3, LM_insrt_077_target3, LM_GI1_target3, LM_GI0_target3, LM_flsh27_target3, LM_flsh28_target3)
Dim BP_wellzom: BP_wellzom=Array(BM_wellzom, LM_insrt_003_wellzom, LM_flsh19a_wellzom, LM_insrt_050_wellzom, LM_insrt_051_wellzom, LM_insrt_052_wellzom, LM_insrt_053_wellzom, LM_insrt_054_wellzom, LM_insrt_070_wellzom, LM_insrt_071_wellzom, LM_insrt78_wellzom, LM_insrt_168_wellzom, LM_GI1_wellzom, LM_GI0_wellzom, LM_flsh27_wellzom, LM_flsh28_wellzom, LM_flsh29_wellzom)
Dim BP_wellzombase: BP_wellzombase=Array(BM_wellzombase, LM_flsh19a_wellzombase, LM_insrt_051_wellzombase, LM_insrt_052_wellzombase, LM_insrt_053_wellzombase, LM_insrt_054_wellzombase, LM_insrt_168_wellzombase, LM_GI1_wellzombase, LM_GI0_wellzombase, LM_flsh27_wellzombase, LM_flsh28_wellzombase, LM_flsh29_wellzombase)
' Arrays per lighting scenario
Dim BL_GI0: BL_GI0=Array(LM_GI0_BatLeft, LM_GI0_BatRight, LM_GI0_BotRamp, LM_GI0_Bumper1_Ring, LM_GI0_Bumper1_Socket, LM_GI0_Bumper2_Ring, LM_GI0_Bumper2_Socket, LM_GI0_Bumper3_Ring, LM_GI0_Bumper3_Socket, LM_GI0_Parts, LM_GI0_Playfield, LM_GI0_Spin1, LM_GI0_Spin2, LM_GI0_armp0, LM_GI0_armp1, LM_GI0_armp2, LM_GI0_armp3, LM_GI0_biczom, LM_GI0_crossbow, LM_GI0_crossbowbott, LM_GI0_doorleft, LM_GI0_doorright, LM_GI0_non_opaque, LM_GI0_priszom, LM_GI0_psw10, LM_GI0_psw11, LM_GI0_psw48, LM_GI0_psw9, LM_GI0_rampscrw, LM_GI0_target1, LM_GI0_target2, LM_GI0_target3, LM_GI0_wellzom, LM_GI0_wellzombase)
Dim BL_GI1: BL_GI1=Array(LM_GI1_BatLeft, LM_GI1_BatRight, LM_GI1_BotRamp, LM_GI1_Bumper1_Ring, LM_GI1_Bumper1_Socket, LM_GI1_Bumper2_Ring, LM_GI1_Bumper2_Socket, LM_GI1_Bumper3_Ring, LM_GI1_Bumper3_Socket, LM_GI1_Parts, LM_GI1_Playfield, LM_GI1_Spin1, LM_GI1_Spin2, LM_GI1_armp0, LM_GI1_armp1, LM_GI1_armp2, LM_GI1_armp3, LM_GI1_biczom, LM_GI1_crossbow, LM_GI1_crossbowbott, LM_GI1_doorleft, LM_GI1_doorright, LM_GI1_non_opaque, LM_GI1_priszom, LM_GI1_psw10, LM_GI1_psw11, LM_GI1_psw48, LM_GI1_psw9, LM_GI1_rampscrw, LM_GI1_target1, LM_GI1_target2, LM_GI1_target3, LM_GI1_wellzom, LM_GI1_wellzombase)
Dim BL_World: BL_World=Array(BM_BatLeft, BM_BatRight, BM_BotRamp, BM_Bumper1_Ring, BM_Bumper1_Socket, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_Bumper3_Ring, BM_Bumper3_Socket, BM_Parts, BM_Playfield, BM_Spin1, BM_Spin2, BM_armp0, BM_armp1, BM_armp2, BM_armp3, BM_biczom, BM_crossbow, BM_crossbowbott, BM_doorleft, BM_doorright, BM_non_opaque, BM_priszom, BM_psw10, BM_psw11, BM_psw48, BM_psw9, BM_rampscrw, BM_target1, BM_target2, BM_target3, BM_wellzom, BM_wellzombase)
Dim BL_diode: BL_diode=Array(LM_diode_Bumper1_Ring, LM_diode_Parts, LM_diode_Playfield, LM_diode_doorleft, LM_diode_doorright, LM_diode_priszom)
Dim BL_flsh19a: BL_flsh19a=Array(LM_flsh19a_Bumper1_Ring, LM_flsh19a_Parts, LM_flsh19a_Playfield, LM_flsh19a_wellzom, LM_flsh19a_wellzombase)
Dim BL_flsh20: BL_flsh20=Array(LM_flsh20_Parts, LM_flsh20_Playfield, LM_flsh20_Spin1, LM_flsh20_armp3, LM_flsh20_non_opaque)
Dim BL_flsh25: BL_flsh25=Array(LM_flsh25_Bumper1_Ring, LM_flsh25_Bumper1_Socket, LM_flsh25_Bumper2_Ring, LM_flsh25_Bumper2_Socket, LM_flsh25_Bumper3_Ring, LM_flsh25_Bumper3_Socket, LM_flsh25_Parts, LM_flsh25_Playfield, LM_flsh25_armp2, LM_flsh25_armp3, LM_flsh25_non_opaque, LM_flsh25_rampscrw)
Dim BL_flsh26: BL_flsh26=Array(LM_flsh26_Bumper1_Ring, LM_flsh26_Bumper1_Socket, LM_flsh26_Bumper3_Ring, LM_flsh26_Bumper3_Socket, LM_flsh26_Parts, LM_flsh26_Playfield, LM_flsh26_armp0, LM_flsh26_armp1, LM_flsh26_armp2, LM_flsh26_armp3, LM_flsh26_doorleft, LM_flsh26_doorright, LM_flsh26_priszom, LM_flsh26_rampscrw, LM_flsh26_target2)
Dim BL_flsh27: BL_flsh27=Array(LM_flsh27_BotRamp, LM_flsh27_Bumper1_Ring, LM_flsh27_Bumper1_Socket, LM_flsh27_Bumper2_Ring, LM_flsh27_Bumper2_Socket, LM_flsh27_Bumper3_Ring, LM_flsh27_Bumper3_Socket, LM_flsh27_Parts, LM_flsh27_Playfield, LM_flsh27_Spin2, LM_flsh27_armp0, LM_flsh27_armp1, LM_flsh27_armp2, LM_flsh27_armp3, LM_flsh27_biczom, LM_flsh27_crossbowbott, LM_flsh27_doorleft, LM_flsh27_doorright, LM_flsh27_non_opaque, LM_flsh27_priszom, LM_flsh27_rampscrw, LM_flsh27_target1, LM_flsh27_target2, LM_flsh27_target3, LM_flsh27_wellzom, LM_flsh27_wellzombase)
Dim BL_flsh28: BL_flsh28=Array(LM_flsh28_BatLeft, LM_flsh28_BatRight, LM_flsh28_BotRamp, LM_flsh28_Bumper1_Ring, LM_flsh28_Bumper1_Socket, LM_flsh28_Bumper2_Ring, LM_flsh28_Bumper2_Socket, LM_flsh28_Bumper3_Ring, LM_flsh28_Bumper3_Socket, LM_flsh28_Parts, LM_flsh28_Playfield, LM_flsh28_Spin1, LM_flsh28_Spin2, LM_flsh28_armp0, LM_flsh28_armp1, LM_flsh28_armp2, LM_flsh28_armp3, LM_flsh28_biczom, LM_flsh28_crossbow, LM_flsh28_crossbowbott, LM_flsh28_doorleft, LM_flsh28_doorright, LM_flsh28_non_opaque, LM_flsh28_priszom, LM_flsh28_psw10, LM_flsh28_psw11, LM_flsh28_psw48, LM_flsh28_psw9, LM_flsh28_rampscrw, LM_flsh28_target1, LM_flsh28_target2, LM_flsh28_target3, LM_flsh28_wellzom, LM_flsh28_wellzombase)
Dim BL_flsh29: BL_flsh29=Array(LM_flsh29_BatLeft, LM_flsh29_BotRamp, LM_flsh29_Bumper1_Ring, LM_flsh29_Bumper1_Socket, LM_flsh29_Bumper2_Ring, LM_flsh29_Bumper2_Socket, LM_flsh29_Bumper3_Ring, LM_flsh29_Bumper3_Socket, LM_flsh29_Parts, LM_flsh29_Playfield, LM_flsh29_Spin1, LM_flsh29_Spin2, LM_flsh29_armp0, LM_flsh29_armp1, LM_flsh29_armp2, LM_flsh29_armp3, LM_flsh29_crossbow, LM_flsh29_crossbowbott, LM_flsh29_doorleft, LM_flsh29_doorright, LM_flsh29_non_opaque, LM_flsh29_priszom, LM_flsh29_psw10, LM_flsh29_psw11, LM_flsh29_psw48, LM_flsh29_psw9, LM_flsh29_rampscrw, LM_flsh29_target2, LM_flsh29_wellzom, LM_flsh29_wellzombase)
Dim BL_flsh31: BL_flsh31=Array(LM_flsh31_Parts, LM_flsh31_Playfield, LM_flsh31_Spin2, LM_flsh31_armp0)
Dim BL_insrt_003: BL_insrt_003=Array(LM_insrt_003_wellzom)
Dim BL_insrt_022: BL_insrt_022=Array(LM_insrt_022_Parts, LM_insrt_022_armp0)
Dim BL_insrt_023: BL_insrt_023=Array(LM_insrt_023_Parts, LM_insrt_023_armp0)
Dim BL_insrt_030: BL_insrt_030=Array(LM_insrt_030_Parts)
Dim BL_insrt_031: BL_insrt_031=Array(LM_insrt_031_Parts, LM_insrt_031_Spin2, LM_insrt_031_armp0)
Dim BL_insrt_032: BL_insrt_032=Array(LM_insrt_032_BotRamp, LM_insrt_032_Parts, LM_insrt_032_Spin2, LM_insrt_032_armp0)
Dim BL_insrt_037: BL_insrt_037=Array(LM_insrt_037_Parts, LM_insrt_037_armp3)
Dim BL_insrt_038: BL_insrt_038=Array(LM_insrt_038_Parts, LM_insrt_038_armp3, LM_insrt_038_rampscrw)
Dim BL_insrt_039: BL_insrt_039=Array(LM_insrt_039_Parts, LM_insrt_039_Spin1)
Dim BL_insrt_040: BL_insrt_040=Array(LM_insrt_040_Parts, LM_insrt_040_Spin1, LM_insrt_040_armp3)
Dim BL_insrt_050: BL_insrt_050=Array(LM_insrt_050_wellzom)
Dim BL_insrt_051: BL_insrt_051=Array(LM_insrt_051_Parts, LM_insrt_051_wellzom, LM_insrt_051_wellzombase)
Dim BL_insrt_052: BL_insrt_052=Array(LM_insrt_052_Parts, LM_insrt_052_wellzom, LM_insrt_052_wellzombase)
Dim BL_insrt_053: BL_insrt_053=Array(LM_insrt_053_Parts, LM_insrt_053_Spin1, LM_insrt_053_non_opaque, LM_insrt_053_wellzom, LM_insrt_053_wellzombase)
Dim BL_insrt_054: BL_insrt_054=Array(LM_insrt_054_wellzom, LM_insrt_054_wellzombase)
Dim BL_insrt_055: BL_insrt_055=Array(LM_insrt_055_Parts, LM_insrt_055_target1)
Dim BL_insrt_056: BL_insrt_056=Array(LM_insrt_056_target2)
Dim BL_insrt_058: BL_insrt_058=Array(LM_insrt_058_Parts)
Dim BL_insrt_059: BL_insrt_059=Array(LM_insrt_059_Parts, LM_insrt_059_target2)
Dim BL_insrt_063: BL_insrt_063=Array(LM_insrt_063_Parts, LM_insrt_063_rampscrw, LM_insrt_063_target2)
Dim BL_insrt_070: BL_insrt_070=Array(LM_insrt_070_Parts, LM_insrt_070_wellzom)
Dim BL_insrt_071: BL_insrt_071=Array(LM_insrt_071_Parts, LM_insrt_071_armp3, LM_insrt_071_wellzom)
Dim BL_insrt_072: BL_insrt_072=Array(LM_insrt_072_Parts)
Dim BL_insrt_073: BL_insrt_073=Array(LM_insrt_073_Parts, LM_insrt_073_non_opaque)
Dim BL_insrt_074: BL_insrt_074=Array(LM_insrt_074_Parts, LM_insrt_074_non_opaque)
Dim BL_insrt_075: BL_insrt_075=Array(LM_insrt_075_Parts, LM_insrt_075_non_opaque)
Dim BL_insrt_076: BL_insrt_076=Array(LM_insrt_076_Parts, LM_insrt_076_armp2, LM_insrt_076_rampscrw)
Dim BL_insrt_077: BL_insrt_077=Array(LM_insrt_077_Parts, LM_insrt_077_armp2, LM_insrt_077_target3)
Dim BL_insrt_078: BL_insrt_078=Array(LM_insrt_078_Parts, LM_insrt_078_armp0, LM_insrt_078_biczom, LM_insrt_078_rampscrw)
Dim BL_insrt_133: BL_insrt_133=Array(LM_insrt_133_Parts, LM_insrt_133_Playfield)
Dim BL_insrt_136: BL_insrt_136=Array(LM_insrt_136_Parts, LM_insrt_136_Playfield)
Dim BL_insrt_152: BL_insrt_152=Array(LM_insrt_152_Parts, LM_insrt_152_Spin1, LM_insrt_152_non_opaque)
Dim BL_insrt_168: BL_insrt_168=Array(LM_insrt_168_Bumper1_Ring, LM_insrt_168_Bumper1_Socket, LM_insrt_168_Parts, LM_insrt_168_wellzom, LM_insrt_168_wellzombase)
Dim BL_insrt_187: BL_insrt_187=Array(LM_insrt_187_Parts)
Dim BL_insrt_195: BL_insrt_195=Array(LM_insrt_195_Parts, LM_insrt_195_Spin2)
Dim BL_insrt_203: BL_insrt_203=Array(LM_insrt_203_BotRamp, LM_insrt_203_Parts, LM_insrt_203_Spin2)
Dim BL_insrt60: BL_insrt60=Array(LM_insrt60_Bumper1_Ring, LM_insrt60_Bumper2_Ring, LM_insrt60_Bumper2_Socket, LM_insrt60_Bumper3_Ring, LM_insrt60_Bumper3_Socket, LM_insrt60_Parts, LM_insrt60_Playfield, LM_insrt60_armp2, LM_insrt60_doorright)
Dim BL_insrt61: BL_insrt61=Array(LM_insrt61_Bumper1_Ring, LM_insrt61_Bumper1_Socket, LM_insrt61_Bumper2_Ring, LM_insrt61_Bumper2_Socket, LM_insrt61_Bumper3_Ring, LM_insrt61_Bumper3_Socket, LM_insrt61_Parts, LM_insrt61_Playfield, LM_insrt61_armp2, LM_insrt61_armp3)
Dim BL_insrt62: BL_insrt62=Array(LM_insrt62_Bumper1_Ring, LM_insrt62_Bumper1_Socket, LM_insrt62_Bumper2_Ring, LM_insrt62_Bumper3_Ring, LM_insrt62_Bumper3_Socket, LM_insrt62_Parts, LM_insrt62_Playfield)
Dim BL_insrt78: BL_insrt78=Array(LM_insrt78_BotRamp, LM_insrt78_Parts, LM_insrt78_Playfield, LM_insrt78_Spin2, LM_insrt78_armp0, LM_insrt78_biczom, LM_insrt78_rampscrw, LM_insrt78_target2, LM_insrt78_wellzom)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BatLeft, BM_BatRight, BM_BotRamp, BM_Bumper1_Ring, BM_Bumper1_Socket, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_Bumper3_Ring, BM_Bumper3_Socket, BM_Parts, BM_Playfield, BM_Spin1, BM_Spin2, BM_armp0, BM_armp1, BM_armp2, BM_armp3, BM_biczom, BM_crossbow, BM_crossbowbott, BM_doorleft, BM_doorright, BM_non_opaque, BM_priszom, BM_psw10, BM_psw11, BM_psw48, BM_psw9, BM_rampscrw, BM_target1, BM_target2, BM_target3, BM_wellzom, BM_wellzombase)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI0_BatLeft, LM_GI0_BatRight, LM_GI0_BotRamp, LM_GI0_Bumper1_Ring, LM_GI0_Bumper1_Socket, LM_GI0_Bumper2_Ring, LM_GI0_Bumper2_Socket, LM_GI0_Bumper3_Ring, LM_GI0_Bumper3_Socket, LM_GI0_Parts, LM_GI0_Playfield, LM_GI0_Spin1, LM_GI0_Spin2, LM_GI0_armp0, LM_GI0_armp1, LM_GI0_armp2, LM_GI0_armp3, LM_GI0_biczom, LM_GI0_crossbow, LM_GI0_crossbowbott, LM_GI0_doorleft, LM_GI0_doorright, LM_GI0_non_opaque, LM_GI0_priszom, LM_GI0_psw10, LM_GI0_psw11, LM_GI0_psw48, LM_GI0_psw9, LM_GI0_rampscrw, LM_GI0_target1, LM_GI0_target2, LM_GI0_target3, LM_GI0_wellzom, LM_GI0_wellzombase, LM_GI1_BatLeft, LM_GI1_BatRight, LM_GI1_BotRamp, LM_GI1_Bumper1_Ring, LM_GI1_Bumper1_Socket, LM_GI1_Bumper2_Ring, LM_GI1_Bumper2_Socket, LM_GI1_Bumper3_Ring, LM_GI1_Bumper3_Socket, LM_GI1_Parts, LM_GI1_Playfield, LM_GI1_Spin1, LM_GI1_Spin2, LM_GI1_armp0, LM_GI1_armp1, LM_GI1_armp2, LM_GI1_armp3, LM_GI1_biczom, LM_GI1_crossbow, LM_GI1_crossbowbott, LM_GI1_doorleft, LM_GI1_doorright, LM_GI1_non_opaque, _
  LM_GI1_priszom, LM_GI1_psw10, LM_GI1_psw11, LM_GI1_psw48, LM_GI1_psw9, LM_GI1_rampscrw, LM_GI1_target1, LM_GI1_target2, LM_GI1_target3, LM_GI1_wellzom, LM_GI1_wellzombase, LM_diode_Bumper1_Ring, LM_diode_Parts, LM_diode_Playfield, LM_diode_doorleft, LM_diode_doorright, LM_diode_priszom, LM_flsh19a_Bumper1_Ring, LM_flsh19a_Parts, LM_flsh19a_Playfield, LM_flsh19a_wellzom, LM_flsh19a_wellzombase, LM_flsh20_Parts, LM_flsh20_Playfield, LM_flsh20_Spin1, LM_flsh20_armp3, LM_flsh20_non_opaque, LM_flsh25_Bumper1_Ring, LM_flsh25_Bumper1_Socket, LM_flsh25_Bumper2_Ring, LM_flsh25_Bumper2_Socket, LM_flsh25_Bumper3_Ring, LM_flsh25_Bumper3_Socket, LM_flsh25_Parts, LM_flsh25_Playfield, LM_flsh25_armp2, LM_flsh25_armp3, LM_flsh25_non_opaque, LM_flsh25_rampscrw, LM_flsh26_Bumper1_Ring, LM_flsh26_Bumper1_Socket, LM_flsh26_Bumper3_Ring, LM_flsh26_Bumper3_Socket, LM_flsh26_Parts, LM_flsh26_Playfield, LM_flsh26_armp0, LM_flsh26_armp1, LM_flsh26_armp2, LM_flsh26_armp3, LM_flsh26_doorleft, LM_flsh26_doorright, LM_flsh26_priszom, _
  LM_flsh26_rampscrw, LM_flsh26_target2, LM_flsh27_BotRamp, LM_flsh27_Bumper1_Ring, LM_flsh27_Bumper1_Socket, LM_flsh27_Bumper2_Ring, LM_flsh27_Bumper2_Socket, LM_flsh27_Bumper3_Ring, LM_flsh27_Bumper3_Socket, LM_flsh27_Parts, LM_flsh27_Playfield, LM_flsh27_Spin2, LM_flsh27_armp0, LM_flsh27_armp1, LM_flsh27_armp2, LM_flsh27_armp3, LM_flsh27_biczom, LM_flsh27_crossbowbott, LM_flsh27_doorleft, LM_flsh27_doorright, LM_flsh27_non_opaque, LM_flsh27_priszom, LM_flsh27_rampscrw, LM_flsh27_target1, LM_flsh27_target2, LM_flsh27_target3, LM_flsh27_wellzom, LM_flsh27_wellzombase, LM_flsh28_BatLeft, LM_flsh28_BatRight, LM_flsh28_BotRamp, LM_flsh28_Bumper1_Ring, LM_flsh28_Bumper1_Socket, LM_flsh28_Bumper2_Ring, LM_flsh28_Bumper2_Socket, LM_flsh28_Bumper3_Ring, LM_flsh28_Bumper3_Socket, LM_flsh28_Parts, LM_flsh28_Playfield, LM_flsh28_Spin1, LM_flsh28_Spin2, LM_flsh28_armp0, LM_flsh28_armp1, LM_flsh28_armp2, LM_flsh28_armp3, LM_flsh28_biczom, LM_flsh28_crossbow, LM_flsh28_crossbowbott, LM_flsh28_doorleft, LM_flsh28_doorright, _
  LM_flsh28_non_opaque, LM_flsh28_priszom, LM_flsh28_psw10, LM_flsh28_psw11, LM_flsh28_psw48, LM_flsh28_psw9, LM_flsh28_rampscrw, LM_flsh28_target1, LM_flsh28_target2, LM_flsh28_target3, LM_flsh28_wellzom, LM_flsh28_wellzombase, LM_flsh29_BatLeft, LM_flsh29_BotRamp, LM_flsh29_Bumper1_Ring, LM_flsh29_Bumper1_Socket, LM_flsh29_Bumper2_Ring, LM_flsh29_Bumper2_Socket, LM_flsh29_Bumper3_Ring, LM_flsh29_Bumper3_Socket, LM_flsh29_Parts, LM_flsh29_Playfield, LM_flsh29_Spin1, LM_flsh29_Spin2, LM_flsh29_armp0, LM_flsh29_armp1, LM_flsh29_armp2, LM_flsh29_armp3, LM_flsh29_crossbow, LM_flsh29_crossbowbott, LM_flsh29_doorleft, LM_flsh29_doorright, LM_flsh29_non_opaque, LM_flsh29_priszom, LM_flsh29_psw10, LM_flsh29_psw11, LM_flsh29_psw48, LM_flsh29_psw9, LM_flsh29_rampscrw, LM_flsh29_target2, LM_flsh29_wellzom, LM_flsh29_wellzombase, LM_flsh31_Parts, LM_flsh31_Playfield, LM_flsh31_Spin2, LM_flsh31_armp0, LM_insrt_003_wellzom, LM_insrt_022_Parts, LM_insrt_022_armp0, LM_insrt_023_Parts, LM_insrt_023_armp0, LM_insrt_030_Parts, _
  LM_insrt_031_Parts, LM_insrt_031_Spin2, LM_insrt_031_armp0, LM_insrt_032_BotRamp, LM_insrt_032_Parts, LM_insrt_032_Spin2, LM_insrt_032_armp0, LM_insrt_037_Parts, LM_insrt_037_armp3, LM_insrt_038_Parts, LM_insrt_038_armp3, LM_insrt_038_rampscrw, LM_insrt_039_Parts, LM_insrt_039_Spin1, LM_insrt_040_Parts, LM_insrt_040_Spin1, LM_insrt_040_armp3, LM_insrt_050_wellzom, LM_insrt_051_Parts, LM_insrt_051_wellzom, LM_insrt_051_wellzombase, LM_insrt_052_Parts, LM_insrt_052_wellzom, LM_insrt_052_wellzombase, LM_insrt_053_Parts, LM_insrt_053_Spin1, LM_insrt_053_non_opaque, LM_insrt_053_wellzom, LM_insrt_053_wellzombase, LM_insrt_054_wellzom, LM_insrt_054_wellzombase, LM_insrt_055_Parts, LM_insrt_055_target1, LM_insrt_056_target2, LM_insrt_058_Parts, LM_insrt_059_Parts, LM_insrt_059_target2, LM_insrt_063_Parts, _
  LM_insrt_063_rampscrw, LM_insrt_063_target2, LM_insrt_070_Parts, LM_insrt_070_wellzom, LM_insrt_071_Parts, LM_insrt_071_armp3, LM_insrt_071_wellzom, LM_insrt_072_Parts, LM_insrt_073_Parts, LM_insrt_073_non_opaque, LM_insrt_074_Parts, LM_insrt_074_non_opaque, LM_insrt_075_Parts, LM_insrt_075_non_opaque, LM_insrt_076_Parts, LM_insrt_076_armp2, LM_insrt_076_rampscrw, LM_insrt_077_Parts, LM_insrt_077_armp2, LM_insrt_077_target3, LM_insrt_078_Parts, LM_insrt_078_armp0, LM_insrt_078_biczom, LM_insrt_078_rampscrw, LM_insrt_133_Parts, LM_insrt_133_Playfield, LM_insrt_136_Parts, LM_insrt_136_Playfield, LM_insrt_152_Parts, LM_insrt_152_Spin1, _
  LM_insrt_152_non_opaque, LM_insrt_168_Bumper1_Ring, LM_insrt_168_Bumper1_Socket, LM_insrt_168_Parts, LM_insrt_168_wellzom, LM_insrt_168_wellzombase, LM_insrt_187_Parts, LM_insrt_195_Parts, LM_insrt_195_Spin2, LM_insrt_203_BotRamp, LM_insrt_203_Parts, LM_insrt_203_Spin2, LM_insrt60_Bumper1_Ring, LM_insrt60_Bumper2_Ring, LM_insrt60_Bumper2_Socket, LM_insrt60_Bumper3_Ring, LM_insrt60_Bumper3_Socket, LM_insrt60_Parts, LM_insrt60_Playfield, LM_insrt60_armp2, LM_insrt60_doorright, LM_insrt61_Bumper1_Ring, LM_insrt61_Bumper1_Socket, LM_insrt61_Bumper2_Ring, LM_insrt61_Bumper2_Socket, LM_insrt61_Bumper3_Ring, LM_insrt61_Bumper3_Socket, LM_insrt61_Parts, LM_insrt61_Playfield, LM_insrt61_armp2, LM_insrt61_armp3, LM_insrt62_Bumper1_Ring, LM_insrt62_Bumper1_Socket, LM_insrt62_Bumper2_Ring, LM_insrt62_Bumper3_Ring, LM_insrt62_Bumper3_Socket, LM_insrt62_Parts, LM_insrt62_Playfield, LM_insrt78_BotRamp, _
  LM_insrt78_Parts, LM_insrt78_Playfield, LM_insrt78_Spin2, LM_insrt78_armp0, LM_insrt78_biczom, LM_insrt78_rampscrw, LM_insrt78_target2, LM_insrt78_wellzom)
Dim BG_All: BG_All=Array(BM_BatLeft, BM_BatRight, BM_BotRamp, BM_Bumper1_Ring, BM_Bumper1_Socket, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_Bumper3_Ring, BM_Bumper3_Socket, BM_Parts, BM_Playfield, BM_Spin1, BM_Spin2, BM_armp0, BM_armp1, BM_armp2, BM_armp3, BM_biczom, BM_crossbow, BM_crossbowbott, BM_doorleft, BM_doorright, BM_non_opaque, BM_priszom, BM_psw10, BM_psw11, BM_psw48, BM_psw9, BM_rampscrw, BM_target1, BM_target2, BM_target3, BM_wellzom, BM_wellzombase, LM_GI0_BatLeft, LM_GI0_BatRight, LM_GI0_BotRamp, LM_GI0_Bumper1_Ring, LM_GI0_Bumper1_Socket, LM_GI0_Bumper2_Ring, LM_GI0_Bumper2_Socket, LM_GI0_Bumper3_Ring, LM_GI0_Bumper3_Socket, LM_GI0_Parts, LM_GI0_Playfield, LM_GI0_Spin1, LM_GI0_Spin2, LM_GI0_armp0, LM_GI0_armp1, LM_GI0_armp2, LM_GI0_armp3, LM_GI0_biczom, LM_GI0_crossbow, LM_GI0_crossbowbott, LM_GI0_doorleft, LM_GI0_doorright, LM_GI0_non_opaque, LM_GI0_priszom, LM_GI0_psw10, LM_GI0_psw11, LM_GI0_psw48, LM_GI0_psw9, LM_GI0_rampscrw, LM_GI0_target1, LM_GI0_target2, LM_GI0_target3, LM_GI0_wellzom, _
  LM_GI0_wellzombase, LM_GI1_BatLeft, LM_GI1_BatRight, LM_GI1_BotRamp, LM_GI1_Bumper1_Ring, LM_GI1_Bumper1_Socket, LM_GI1_Bumper2_Ring, LM_GI1_Bumper2_Socket, LM_GI1_Bumper3_Ring, LM_GI1_Bumper3_Socket, LM_GI1_Parts, LM_GI1_Playfield, LM_GI1_Spin1, LM_GI1_Spin2, LM_GI1_armp0, LM_GI1_armp1, LM_GI1_armp2, LM_GI1_armp3, LM_GI1_biczom, LM_GI1_crossbow, LM_GI1_crossbowbott, LM_GI1_doorleft, LM_GI1_doorright, LM_GI1_non_opaque, LM_GI1_priszom, LM_GI1_psw10, LM_GI1_psw11, LM_GI1_psw48, LM_GI1_psw9, LM_GI1_rampscrw, LM_GI1_target1, LM_GI1_target2, LM_GI1_target3, LM_GI1_wellzom, LM_GI1_wellzombase, LM_diode_Bumper1_Ring, LM_diode_Parts, LM_diode_Playfield, LM_diode_doorleft, LM_diode_doorright, LM_diode_priszom, LM_flsh19a_Bumper1_Ring, LM_flsh19a_Parts, LM_flsh19a_Playfield, LM_flsh19a_wellzom, LM_flsh19a_wellzombase, LM_flsh20_Parts, LM_flsh20_Playfield, LM_flsh20_Spin1, LM_flsh20_armp3, LM_flsh20_non_opaque, LM_flsh25_Bumper1_Ring, LM_flsh25_Bumper1_Socket, LM_flsh25_Bumper2_Ring, LM_flsh25_Bumper2_Socket, _
  LM_flsh25_Bumper3_Ring, LM_flsh25_Bumper3_Socket, LM_flsh25_Parts, LM_flsh25_Playfield, LM_flsh25_armp2, LM_flsh25_armp3, LM_flsh25_non_opaque, LM_flsh25_rampscrw, LM_flsh26_Bumper1_Ring, LM_flsh26_Bumper1_Socket, LM_flsh26_Bumper3_Ring, LM_flsh26_Bumper3_Socket, LM_flsh26_Parts, LM_flsh26_Playfield, LM_flsh26_armp0, LM_flsh26_armp1, LM_flsh26_armp2, LM_flsh26_armp3, LM_flsh26_doorleft, LM_flsh26_doorright, LM_flsh26_priszom, LM_flsh26_rampscrw, LM_flsh26_target2, LM_flsh27_BotRamp, LM_flsh27_Bumper1_Ring, LM_flsh27_Bumper1_Socket, LM_flsh27_Bumper2_Ring, LM_flsh27_Bumper2_Socket, LM_flsh27_Bumper3_Ring, LM_flsh27_Bumper3_Socket, LM_flsh27_Parts, LM_flsh27_Playfield, LM_flsh27_Spin2, LM_flsh27_armp0, LM_flsh27_armp1, LM_flsh27_armp2, LM_flsh27_armp3, LM_flsh27_biczom, LM_flsh27_crossbowbott, LM_flsh27_doorleft, LM_flsh27_doorright, LM_flsh27_non_opaque, LM_flsh27_priszom, LM_flsh27_rampscrw, LM_flsh27_target1, LM_flsh27_target2, LM_flsh27_target3, LM_flsh27_wellzom, LM_flsh27_wellzombase, LM_flsh28_BatLeft, _
  LM_flsh28_BatRight, LM_flsh28_BotRamp, LM_flsh28_Bumper1_Ring, LM_flsh28_Bumper1_Socket, LM_flsh28_Bumper2_Ring, LM_flsh28_Bumper2_Socket, LM_flsh28_Bumper3_Ring, LM_flsh28_Bumper3_Socket, LM_flsh28_Parts, LM_flsh28_Playfield, LM_flsh28_Spin1, LM_flsh28_Spin2, LM_flsh28_armp0, LM_flsh28_armp1, LM_flsh28_armp2, LM_flsh28_armp3, LM_flsh28_biczom, LM_flsh28_crossbow, LM_flsh28_crossbowbott, LM_flsh28_doorleft, LM_flsh28_doorright, LM_flsh28_non_opaque, LM_flsh28_priszom, LM_flsh28_psw10, LM_flsh28_psw11, LM_flsh28_psw48, LM_flsh28_psw9, LM_flsh28_rampscrw, LM_flsh28_target1, LM_flsh28_target2, LM_flsh28_target3, LM_flsh28_wellzom, LM_flsh28_wellzombase, LM_flsh29_BatLeft, LM_flsh29_BotRamp, LM_flsh29_Bumper1_Ring, LM_flsh29_Bumper1_Socket, LM_flsh29_Bumper2_Ring, LM_flsh29_Bumper2_Socket, LM_flsh29_Bumper3_Ring, LM_flsh29_Bumper3_Socket, LM_flsh29_Parts, LM_flsh29_Playfield, LM_flsh29_Spin1, LM_flsh29_Spin2, LM_flsh29_armp0, LM_flsh29_armp1, LM_flsh29_armp2, LM_flsh29_armp3, LM_flsh29_crossbow, _
  LM_flsh29_crossbowbott, LM_flsh29_doorleft, LM_flsh29_doorright, LM_flsh29_non_opaque, LM_flsh29_priszom, LM_flsh29_psw10, LM_flsh29_psw11, LM_flsh29_psw48, LM_flsh29_psw9, LM_flsh29_rampscrw, LM_flsh29_target2, LM_flsh29_wellzom, LM_flsh29_wellzombase, LM_flsh31_Parts, LM_flsh31_Playfield, LM_flsh31_Spin2, LM_flsh31_armp0, LM_insrt_003_wellzom, LM_insrt_022_Parts, LM_insrt_022_armp0, LM_insrt_023_Parts, LM_insrt_023_armp0, LM_insrt_030_Parts, LM_insrt_031_Parts, LM_insrt_031_Spin2, LM_insrt_031_armp0, LM_insrt_032_BotRamp, LM_insrt_032_Parts, LM_insrt_032_Spin2, LM_insrt_032_armp0, LM_insrt_037_Parts, LM_insrt_037_armp3, LM_insrt_038_Parts, LM_insrt_038_armp3, LM_insrt_038_rampscrw, LM_insrt_039_Parts, LM_insrt_039_Spin1, LM_insrt_040_Parts, LM_insrt_040_Spin1, LM_insrt_040_armp3, LM_insrt_050_wellzom, LM_insrt_051_Parts, LM_insrt_051_wellzom, _
  LM_insrt_051_wellzombase, LM_insrt_052_Parts, LM_insrt_052_wellzom, LM_insrt_052_wellzombase, LM_insrt_053_Parts, LM_insrt_053_Spin1, LM_insrt_053_non_opaque, LM_insrt_053_wellzom, LM_insrt_053_wellzombase, LM_insrt_054_wellzom, LM_insrt_054_wellzombase, LM_insrt_055_Parts, LM_insrt_055_target1, LM_insrt_056_target2, LM_insrt_058_Parts, LM_insrt_059_Parts, LM_insrt_059_target2, LM_insrt_063_Parts, LM_insrt_063_rampscrw, LM_insrt_063_target2, LM_insrt_070_Parts, LM_insrt_070_wellzom, LM_insrt_071_Parts, LM_insrt_071_armp3, LM_insrt_071_wellzom, LM_insrt_072_Parts, LM_insrt_073_Parts, LM_insrt_073_non_opaque, LM_insrt_074_Parts, LM_insrt_074_non_opaque, LM_insrt_075_Parts, LM_insrt_075_non_opaque, LM_insrt_076_Parts, LM_insrt_076_armp2, LM_insrt_076_rampscrw, LM_insrt_077_Parts, LM_insrt_077_armp2, _
  LM_insrt_077_target3, LM_insrt_078_Parts, LM_insrt_078_armp0, LM_insrt_078_biczom, LM_insrt_078_rampscrw, LM_insrt_133_Parts, LM_insrt_133_Playfield, LM_insrt_136_Parts, LM_insrt_136_Playfield, LM_insrt_152_Parts, LM_insrt_152_Spin1, LM_insrt_152_non_opaque, LM_insrt_168_Bumper1_Ring, LM_insrt_168_Bumper1_Socket, LM_insrt_168_Parts, LM_insrt_168_wellzom, LM_insrt_168_wellzombase, LM_insrt_187_Parts, LM_insrt_195_Parts, LM_insrt_195_Spin2, LM_insrt_203_BotRamp, LM_insrt_203_Parts, LM_insrt_203_Spin2, LM_insrt60_Bumper1_Ring, LM_insrt60_Bumper2_Ring, LM_insrt60_Bumper2_Socket, LM_insrt60_Bumper3_Ring, LM_insrt60_Bumper3_Socket, LM_insrt60_Parts, LM_insrt60_Playfield, LM_insrt60_armp2, LM_insrt60_doorright, LM_insrt61_Bumper1_Ring, LM_insrt61_Bumper1_Socket, _
  LM_insrt61_Bumper2_Ring, LM_insrt61_Bumper2_Socket, LM_insrt61_Bumper3_Ring, LM_insrt61_Bumper3_Socket, LM_insrt61_Parts, LM_insrt61_Playfield, LM_insrt61_armp2, LM_insrt61_armp3, LM_insrt62_Bumper1_Ring, LM_insrt62_Bumper1_Socket, LM_insrt62_Bumper2_Ring, LM_insrt62_Bumper3_Ring, LM_insrt62_Bumper3_Socket, LM_insrt62_Parts, LM_insrt62_Playfield, LM_insrt78_BotRamp, LM_insrt78_Parts, LM_insrt78_Playfield, LM_insrt78_Spin2, LM_insrt78_armp0, LM_insrt78_biczom, LM_insrt78_rampscrw, LM_insrt78_target2, LM_insrt78_wellzom)
' VLM  Arrays - End

'*************
'VR Stuff
'*************
Dim VRRoom
DIM VRThings

Sub LoadVRRoom
  for each VRThings in VR_Cab:VRThings.visible = 0:Next
  for each VRThings in VR_Min:VRThings.visible = 0:Next
  for each VRThings in VR_Mega:VRThings.visible = 0:Next

  If RenderingMode = 2 or VRTest Then
    VRRoom = VRRoomChoice
    TimerPlunger2.Enabled = True
    'disable table objects that should not be visible
  Else
    VRRoom = 0
    RailsLockbar.NormalMap = ""
  End If
  If VRRoom = 1 Then
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
    for each VRThings in VR_Mega:VRThings.visible = 0:Next
  End If
  If VRRoom = 2 Then
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 1:Next
    for each VRThings in VR_Mega:VRThings.visible = 0:Next
  End If
  If VRRoom = 3 Then
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
    for each VRThings in VR_Mega:VRThings.visible = 1:Next
  End If
  If FSSMode Then
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    TextBox.x = -500
  End If
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


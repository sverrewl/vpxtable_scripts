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

Const VolumeDial = 10       'Change volume of hit events
Const RollingSoundFactor = 0.3    'Change volume of rolling sounds
Const TrollSpeed = 2        '0 - slow, 1 - medium, 2 - fast, 3 - really fast

Dim Ballsize,BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode
Const UseVPMModSol = 1

LoadVPM "01560000", "WPC.VBS", 3.50

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
Const SCoin = "fx_Coin"

'************************************************************************
'            INIT TABLE
'************************************************************************

' using table width and height in script slows down the performance
dim tablewidth: tablewidth = Table1.width
dim tableheight: tableheight = Table1.height

Dim MMBall1, MMBall2, MMBall3, MMBall4

Sub Table1_Init
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

'************************************************************************
'             KEYS
'************************************************************************

Sub Table1_KeyDown(ByVal Keycode)
  If keycode = StartGameKey Then Controller.Switch(13) = 1
  If keycode = PlungerKey Then Controller.Switch(11) = 1
  If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge",0):TwrShake
  If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge",0):TwrShake
  If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge",0):TwrShake
  If keycode = LeftFlipperKey Then LFPress = 1
  If keycode = RightFlipperKey Then rfpress = 1
  If keycode = RightMagnaSave Then Controller.Switch(11) = 1
  If keycode = LockbarKey Then Controller.Switch(11) = 1
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
  If keycode = StartGameKey Then Controller.Switch(13) = 0
  If keycode = PlungerKey Then Controller.Switch(11) = 0
  If keycode = RightMagnaSave Then Controller.Switch(11) = 0
  If keycode = LockbarKey Then Controller.Switch(11) = 0

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

Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_unPaused:Controller.Pause = False:End Sub
Sub Table1_exit():Controller.Pause = False:Controller.Stop:End Sub

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

'******************************************************
'       NFOZZY'S FLIPPERS
'******************************************************


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const cSingleLFlip = False
Const cSingleRFlip = False

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("fx_flipperupL",DOFFlippers), LeftFlipper
    LF.fire  'LeftFlipper.RotateToEnd
  Else
  if leftflipper.currentangle < leftflipper.startangle - 5 then
    PlaySoundAt SoundFX("fx_flipperdownL",DOFFlippers), LeftFlipper
  end if
    LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
   If Enabled Then
    PlaySoundAt SoundFX("fx_flipperupR",DOFFlippers), RightFlipper
    RF.fire  'RightFlipper.RotateToEnd
  Else
  if RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
    PlaySoundAt SoundFX("fx_flipperdownR",DOFFlippers), RightFlipper
  end if
    RightFlipper.RotateToStart
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub RightFlipper_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

rightflipper.timerenabled=True
RightFlipper.timerinterval=1

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
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
EOSTnew = 1.0 'FEOST
EOSAnew = 0.2
EOSRampup = 1.5
SOSRampup = 8.5
LiveCatch = 8

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

'************************************************************************
'            AUTOPLUNGER
'************************************************************************

Sub Auto_Plunger(Enabled)
    If Enabled Then
       Plunger.Fire
       PlaySoundAt SoundFX("fx_AutoPlunger",DOFContactors), Plunger
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
  PlaySoundAtVol "fx_drain", 0.3, sw35
End Sub

Sub sw35_UnHit()  'Drain
  Controller.Switch(35) = 0
End Sub

Sub SolBallRelease(enabled)
  If enabled Then
    If sw32.BallCntOver = 0 Then
      PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw32
    Else
      PlaySoundAtVol SoundFX("fx_ballrel",DOFContactors), 0.5, sw32
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
  PlaySoundAtBall "fx_kicker-enter"
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
      PlaySoundAt SoundFX("fx_scoop_out",DOFContactors), sw38
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

'*********************************************************
'           Functions
'*********************************************************

Function PI()
  PI = 4*Atn(1)
End Function

Function Radians(angle)
  Radians = PI * angle / 180
End Function

Function Degrees(radians)
  Degrees = 180 * radians / PI
End Function

Function ASin(val)
    ASin = 2 * Atn(val / (1 + Sqr(1 - (val * val))))
End Function

Function ACos(val)
    ACos = PI / 2 - ASin(val)
End Function

Function Distance(ax,ay,az,bx,by,bz)
  Distance = SQR((ax - bx)^2 + (ay - by)^2 + (az - bz)^2)
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

Sub SolDrawBridge(enabled)
  If Enabled AND Controller.GetMech(0)/16 <= 15 then
    dbpos = 1
    dbridge.enabled = 1
    PlaySound SoundFX("Bridge_Move", DOFGear), -1, 0.1, AudioPan(braket) , 0, 0, 1, AudioPan(braket)
    DOF 104, DOFOn
  end If
  If Enabled AND Controller.GetMech(0)/16 > 15 then
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
  PlaySoundAtBall "fx_kicker-enter"
  Controller.switch(28)=1
End Sub

Sub SolMerlin(enabled)
  controller.Switch(28) = 0
  If sw28.BallCntOver = 0 Then
    PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw28
  Else
    PlaySoundAt SoundFX("fx_saucer_exit",DOFContactors), sw28
  End If
  sw28.kick 109 + (rnd*2), 14 + (rnd*2)
End Sub

'************************************************************************
'               Moat Scoop
'************************************************************************

Sub sw36_hit
  MoatKick.collidable = True
  PlaySoundAtBall "fx_Moat_enter"
  Controller.switch(36)=1
End Sub

Sub SolMoat(enabled)
  controller.Switch(36) = 0
  If sw36.BallCntOver = 0 Then
    PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw36
  Else
    PlaySoundAt SoundFX("fx_popper",DOFContactors), sw36
  End If
  sw36.kickz 215 , 15, 0, 140
  MoatKick.collidable=false
End Sub


'************************************************************************
'         Slingshots Animation
'************************************************************************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  PlaySoundAt SoundFX("LeftSlingshot",DOFContactors), sling1
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
  PlaySoundat SoundFX("RightSlingshot",DOFContactors), sling2
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

Sub Bumper1_Hit:vpmTimer.PulseSw 53:PlaySoundat SoundFX("LeftBumper_Hit",DOFContactors), Bumper1:Me.TimerEnabled = 1:me.timerinterval = 10:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 54:PlaySoundat SoundFX("RightBumper_Hit",DOFContactors), Bumper2:Me.TimerEnabled = 1:me.timerinterval = 10:End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 55:PlaySoundat SoundFX("TopBumper_Hit",DOFContactors), Bumper3:Me.TimerEnabled = 1:me.timerinterval = 10:End Sub

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
Sub sw66_Hit:Controller.Switch(66) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw66_UnHit:Controller.Switch(66) = 0:End Sub

Sub sw67_Hit:Controller.Switch(67) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw67_UnHit:Controller.Switch(67) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw26_Hit
Controller.Switch(26) = 1
If ActiveBall.VelY => 3.5 Then
ActiveBall.VelY = 3.5
End If
If ActiveBall.VelY <= -10 Then
ActiveBall.VelY = -6
End If
PlaySoundatball "fx_sensor"
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
PlaySoundatball "fx_sensor"
End Sub

Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw68_Hit:Controller.Switch(68) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw68_UnHit:Controller.Switch(68) = 0:End Sub

' Targets
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySoundAtBallVol SoundFX("fx_target",DOFTargets), Vol(ActiveBall)*VolumeDial:End Sub
Sub sw15_Hit:vpmTimer.PulseSw 15:PlaySoundAtBallVol SoundFX("fx_target",DOFTargets), Vol(ActiveBall)*VolumeDial:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySoundAtBallVol SoundFX("fx_target",DOFTargets), Vol(ActiveBall)*VolumeDial:End Sub
Sub sw71_Hit:vpmTimer.PulseSw 71:PlaySoundAtBallVol SoundFX("fx_target",DOFTargets), Vol(ActiveBall)*VolumeDial:End Sub
Sub sw72_Hit:vpmTimer.PulseSw 72:PlaySoundAtBallVol SoundFX("fx_target",DOFTargets), Vol(ActiveBall)*VolumeDial:End Sub
Sub sw73_Hit:vpmTimer.PulseSw 73:PlaySoundAtBallVol SoundFX("fx_target",DOFTargets), Vol(ActiveBall)*VolumeDial:End Sub

' Triggers on the ramps
Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAtBall "fx_gate4":End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:sw62flip.rotatetoend:PlaySoundAtBall "fx_gate4":End Sub
Sub sw62_Unhit:Controller.Switch(62) = 0:sw62flip.rotatetostart:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:End Sub
Sub sw63_Unhit:Controller.Switch(63) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:sw64flip.rotatetoend:PlaySoundAtBall "fx_gate4":End Sub
Sub sw64_Unhit:Controller.Switch(64) = 0:sw64flip.rotatetostart:End Sub

' Gates
Sub Gate1_Hit():PlaySoundAtBall "fx_gate":End Sub
Sub Gate2_Hit():PlaySoundAtBall "fx_gate":End Sub
Sub Gate3_Hit():PlaySoundAtBall "fx_gate":End Sub
Sub Gate9_Hit():PlaySoundAtBall "fx_gate":End Sub
Sub Gate6_Hit():PlaySoundAtBall "fx_gate":End Sub

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
  PlaySoundatBall "fx_gate"
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

Sub sw44_Hit():vpmtimer.PulseSw 44:PlaySoundatBall "fx_gate": End Sub

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
For i=0 To 127: Execute "Set Lights(" & i & ")  = L" & i: Next
Lights(14)=Array(L14,L14a)
Lights(31)=Array(l31,l31a0,l31a1,l31a2)
Lights(32)=Array(l32,l32a0,l32a1,l32a2)
Lights(33)=Array(l33,l33a0,l33a1,l33a2)
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
  RollingSoundUpdate
  BallShadowUpdate
  SynchFlasherObj
  UpdateMechs

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
  batleftshadow.rotz = LeftFlipper.CurrentAngle
  batrightshadow.rotz  = RightFlipper.CurrentAngle

End Sub

'**************
' Ramp Sounds
'**************

Sub LHP1_Hit()
  If ActiveBall.velY < 0  Then    'ball is going up
    PlaySoundAtBall "fx_rrenter"
  Else
    StopSound "fx_rrenter"
  End If
End Sub

Sub LHP1_unHit()
  If ActiveBall.velY > 0  Then
    StopSound "fx_rrenter"
  End If
End Sub

Sub RHP1_Hit()
  If ActiveBall.velY < 0  Then    'ball is going up
    PlaySoundAtBall "fx_rrenter"
  Else
    StopSound "fx_rrenter"
  End If
End Sub

Sub RHP1_unHit()
  If ActiveBall.velY > 0  Then
    StopSound "fx_rrenter"
  End If
End Sub

Sub LHP2_Hit()
  PlaySoundAtBall "fx_lr2"
End Sub

Sub RHP2_Hit()
  PlaySoundAtBall "fx_lr2"
End Sub

Sub LHM_Hit()
  PlaySoundAtBall "fx_metalrolling"
  StopSound "fx_rrenter"
End Sub

Sub RHM_Hit()
  PlaySoundAtBall "fx_metalrolling"
  StopSound "fx_rrenter"
End Sub

Sub CHM_Hit()
  PlaySoundAtBall "fx_metalrolling"
End Sub

Sub LHD_Hit()
  StopSound "fx_metalrolling"
  'vpmtimer.addtimer 200, "BallDropSound"
End Sub

Sub RHD_Hit()
  StopSound "fx_metalrolling"
  'vpmtimer.addtimer 200, "BallDropSound"
End Sub

Sub Trigger8_Hit(): StopSound "fx_metalrolling":End Sub

Sub Trigger4_Hit():RandomSoundMetal:End Sub



'******************************************************
'           SOUNDS
'******************************************************

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLoop(soundname, tableobj)
    PlaySound soundname, -1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, aVol, tableobj)
    PlaySound soundname, 1, aVol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


Sub PlaySoundAtExisting(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,aVol, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,aVol, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallSpeed(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallSpeed(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / tablewidth-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingSoundUpdate()

  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("fx_ballrolling" & b)
  Next

  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)
    If BallSpeed(BOT(b) ) > 1 AND BOT(b).z < 27 and BOT(b).radius > 23  Then
      rolling(b) = True
      PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b))*RollingSoundFactor, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("fx_ballrolling" & b)
        rolling(b) = False
      End If
    End If

    '***Ball Drop Sounds***
    If BOT(b).VelZ < -1 and ((BOT(b).z < 55 and BOT(b).z > 27) or (InRect(BOT(b).x,BOT(b).y,25,25,400,25,100,300,25,300) and BOT(b).z <205 and BOT(b).z > 185)) and not InRect(BOT(b).x, BOT(b).y,135,436,537,376,553,500,155,559) Then 'height adjust for ball drop sounds
      PlaySound "fx_ball_drop" & b, 0, (ABS(BOT(b).velz)/17)^2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    Elseif  BOT(b).VelZ < -1 and BOT(b).z < -30 and BOT(b).z > -60 Then
      PlaySound "fx_hole1", 0, (ABS(BOT(b).velz)/17)^2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    Elseif BOT(b).z < -70 and BOT(b).x > 130 and BOT(b).x < 180 Then
      BOT(b).velx = BOT(b).velx - 1
    Elseif BOT(b).z < 170 and BOT(b).z > 120 and BOT(b).velz < -1 and BOT(b).y < 125 and BOT(b).x > 815 Then
      PlaySound "fx_damseltower", 0, (ABS(BOT(b).velz)/17)^2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If
  Next

  Dim nudgeX, nudgeY, nTilt
  NudgeTiltStatus nudgeX, nudgeY, nTilt

  If nTilt > 3 Then TwrShake

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

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 500, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'**********************
' Other Sound FX
'**********************

Sub MoatK_Hit(idx)
  RandomSoundHole()
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub RandomSoundHole()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySound "fx_Hole1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_Hole2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_Hole3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 4 : PlaySound "fx_Hole4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub RandomSoundMetal()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBall "fx_metal_hit_1"
    Case 2 : PlaySoundAtBall "fx_metal_hit_2"
    Case 3 : PlaySoundAtBall "fx_metal_hit_3"
  End Select
End Sub

'*********************************************************************
'                       Collection Sounds
'*********************************************************************

Sub aApron_Hit(idx):PlaySoundAtBallVolM "fx_apron", Vol(ActiveBall)*VolumeDial:End Sub
Sub aRubbers_Hit(idx):PlaySoundAtBallVol "fx_rubber", Vol(ActiveBall)*VolumeDial:End Sub
Sub aPostRubbers_Hit(idx):PlaySoundAtBallVol "fx_postrubber", Vol(ActiveBall)*VolumeDial:End Sub
Sub aMetals_Hit(idx):PlaySoundAtBallVolM "fx_MetalHit", Vol(ActiveBall)*VolumeDial/100:End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBallVol "fx_PlasticHit", Vol(ActiveBall)*VolumeDial*5:End Sub
Sub aGates_Hit(idx):PlaySoundAtBallVol "fx_Gate", Vol(ActiveBall)*VolumeDial:End Sub
Sub aWoods_Hit(idx):PlaySoundAtBallVolM "fx_Woodhit", Vol(ActiveBall)*VolumeDial/5:End Sub
Sub aSkillShot_Hit(idx):PlaySoundAtBallVol "fx_MetalHit", Vol(ActiveBall)*VolumeDial/100:End Sub


'**********************OPTIONS***************************

Sub InitOptions
Ramp15.visible = 1
Ramp16.visible = 1
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
    x.TimeDelay = 44
  Next

  'rf.report "Velocity"
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.2,   1.07
  addpt "Velocity", 2, 0.41, 1.05
  addpt "Velocity", 3, 0.44, 1
  addpt "Velocity", 4, 0.65,  1.0'0.982
  addpt "Velocity", 5, 0.702, 0.968
  addpt "Velocity", 6, 0.95,  0.968
  addpt "Velocity", 7, 1.03,  0.945

  'rf.report "Polarity"
  AddPt "Polarity", 0, 0, -4.7
  AddPt "Polarity", 1, 0.16, -4.7
  AddPt "Polarity", 2, 0.33, -4.7
  AddPt "Polarity", 3, 0.37, -4.7 '4.2
  AddPt "Polarity", 4, 0.41, -4.7
  AddPt "Polarity", 5, 0.45, -4.7 '4.2
  AddPt "Polarity", 6, 0.576,-4.7
  AddPt "Polarity", 7, 0.66, -2.8'-2.1896
  AddPt "Polarity", 8, 0.743, -1.5
  AddPt "Polarity", 9, 0.81, -1.5
  AddPt "Polarity", 10, 0.88, 0

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

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

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
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 20 Then
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

      'Velocity correction
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
        'playsound "fx_knocker"
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


'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub

dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.935 '0.96 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.935 '0.96
RubbersD.addpoint 2, 5.76, 0.942 '0.967 'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
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

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

'Tracks ball velocity for judging bounce calculations & angle
'apologies to JimmyFingers is this is what his script does. I know his tracks ball velocity too but idk how it works in particular
dim cor : set cor = New CoRTracker
cor.debugOn = False
'cor.update() - put this on a low interval timer
Class CoRTracker
  public DebugOn 'tbpIn.text
  public ballvel

  Private Sub Class_Initialize : redim ballvel(0) : End Sub
  'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs
    'if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
'     if DebugOn then
'       dim s, bs 'debug spacer, ballspeed
'       bs = round(BallSpeed(b),1)
'       if bs < 10 then s = " " else s = "" end if
'       str = str & b.id & ": " & s & bs & vbnewline
'       'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
'     end if
    Next
    'if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
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

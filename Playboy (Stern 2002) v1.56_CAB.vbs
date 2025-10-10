'..................................................................................................................................................................
'.............................................................................................:x$:....................................................................
'.............................................................:&$+...........................X&&$.....................................................................
'..............................................................$&&&&x:.....................+&&&&$.....................................................................
'..............................................................;&&&&&&&+..................$&&&&&$.....................................................................
'...............................................................X&&&&&&&&x..............:$&&&&&&X.....................................................................
'...............................................................:$&&&&&&&&&x............X&&&&&&&X.....................................................................
'.................................................................$&&&&&&&&&&x.........x&&&&&&&&+.....................................................................
'..................................................................x&&&&&&&&&&&:......:&&&&&&&&&:.....................................................................
'...................................................................:$&&&&&&&&&&x.....+&&&&&&&&X......................................................................
'.....................................................................;$&&&&&&&&&$:...x&&&&&&&&:......................................................................
'........................................................................x&&&&&&&&&+..x&&&&&&&;.......................................................................
'...........................................................................;x&&&&&&$:x&&&&&&;........................................................................
'................................................................................x&&&&$&&&&$:.........................................................................
'.........................................................................:xX&&&&&&&&&&&&&+...........................................................................
'.......................................................................+&&&&&&&&&&&&&&&&&&+..........................................................................
'.....................................................................+&&&+..;&&&&&&&&&&&&&&x.........................................................................
'....................................................................x&&&&+..;&&&&&&&&&&&&&&&+........................................................................
'...................................................................;&&&&&&&&&&&&&&&&&&&&&&&&$........................................................................
'...................................................................x&&&&&&&&&&&&&&&&&&&&&&&&$........................................................................
'...................................................................:&&&&&&&&&&&&&&&&&&&&&&&&$........................................................................
'....................................................................:X&&&&&&&&&&&&&&&&&&&&&&x........................................................................
'.......................................................................+&&&&&&&&&&&&&&&&&&&&:........................................................................
'..........................................................................:;$&&&&&&&&&&&&&&+.........................................................................
'...........................................................................;&&&&&$&&&&&&$+...........................................................................
'........................................................................:...&$;:..X&x;:..;+..........................................................................
'........................................................................x&&X;x&&+.:;+$&$Xx:..........................................................................
'........................................................................;&&X;$&&x.+;:................................................................................
'.........................................................................:.....xx....................................................................................
'.....................................................................................................................................................................
'.....................................................................................................................................................................
'.............................................................................................................:::.....................................................
'.................................$&&&&&&&&&$x.:$&&&&&&X.........;&&&&&&+.x&&&&&&X.$&&&&X;&&&&&&&&&&&&+...:$&&&&&&&X:.+&&&&&&&:$&&&&$.................................
'.................................x&&&&&$$&&&&&:x$&&&&X+.........$&&&&&&$.+$&&&&&+.X&&&$+:X&&&&&$$&&&&$..x&&&&&X$&&&&x;X&&&&&X:x&&&&x.................................
'..................................x&&&&:.X&&&&+.x&&&&:.........+&&$;&&&&;..x&&&&X:$&&x....$&&&$..&&&&$.+&&&&&:..$&&&&+.:&&&&&+X&&$...................................
'..................................x&&&&$$&&&&&:.x&&&&:.........$&&+.X&&&$...;&&&&&&&X.....X&&&&&&&&$x..X&&&&&...X&&&&X..:$&&&&&&X....................................
'..................................x&&&&&&&&&X:..x&&&&:........+&&&&&&&&&&x...:&&&&&x......X&&&&&&&&&&X.X&&&&&...X&&&&X....X&&&&$.....................................
'..................................x&&&&;........x&&&&:..:&&:.:&&&&&&&&&&&&:...+&&&&:......X&&&$..$&&&&+;&&&&&:.:$&&&&+....:&&&&x.....................................
'.................................X&&&&&$x......x&&&&&&&&&&&:X&&&&x..;$&&&&&+:X&&&&&$x...;$&&&&&&&&&&&&+.x&&&&&&&&&&&X....x$&&&&&x....................................
'.................................$&&&&&&$.....:$&&&&&&&&&&&:$&&&&X..+&&&&&&x:$&&&&&&X...;&&&&&&&&&&&&x...:X&&&&&&&X:.....X&&&&&&$....................................
'.....................................................................................................................................................................
'.....................................................................................................................................................................
'.....................................................................................................................................................................
'PLAYBOY (STERN 2002) v1.56: by CHEESE3075.  A mod of RASCAL AND HIREZ00's v1.1 table.

Option Explicit


'GAME OPTIONS in f12 'tweak' menu


Dim CurrTime,objIEDebugWindow
CurrTime = Timer

' Thalamus 2020 April : Improved directional sounds
' Options
' Volume devided by - lower gets higher sound
Const VolDiv = 5000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Const VolBump   = 2    ' Bumpers volume.
Const VolFlip   = 1    ' Flipper volume.


'****** PuP Variables ******

Dim usePUP: Dim cPuPPack: Dim PuPlayer: Dim PUPStatus: PUPStatus=false ' dont edit this line!!!

'*************************** PuP Settings for this table ********************************

usePUP   = False         ' enable Pinup Player functions for this table
cPuPPack = "playboys"    ' name of the PuP-Pack / PuPVideos folder for this table

'//////////////////// PINUP PLAYER: STARTUP & CONTROL SECTION //////////////////////////

' This is used for the startup and control of Pinup Player

Sub PuPStart(cPuPPack)
    If PUPStatus=true then Exit Sub
    If usePUP=true then
        Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
        If PuPlayer is Nothing Then
            usePUP=false
            PUPStatus=false
        Else
            PuPlayer.B2SInit "",cPuPPack 'start the Pup-Pack
            PUPStatus=true
        End If
    End If
End Sub

Sub pupevent(EventNum)
    if (usePUP=false or PUPStatus=false) then Exit Sub
    PuPlayer.B2SData "E"&EventNum,1  'send event to Pup-Pack
End Sub

'************ PuP-Pack Startup **************

PuPStart(cPuPPack) 'Check for PuP - If found, then start Pinup Player / PuP-Pack
'pupevent xxx


'**********************************
'   ZMAT: General Math Functions
'**********************************
' These get used throughout the script.

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

Function ArcCos(x)
  If x = 1 Then
    ArcCos = 0/180*PI
  ElseIf x = -1 Then
    ArcCos = 180/180*PI
  Else
    ArcCos = Atn(-x/Sqr(-x * x + 1)) + 2 * Atn(1)
  End If
End Function

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a > b Then
    min = b
  Else
    min = a
  End If
End Function

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

Function RotPoint(x,y,angle)
  Dim rx, ry
  rx = x * dCos(angle) - y * dSin(angle)
  ry = x * dSin(angle) + y * dCos(angle)
  RotPoint = Array(rx,ry)
End Function

'******************************************************
'  ZANI: Misc Animations
'******************************************************

Sub LeftFlipper_Animate
  dim a: a = LeftFlipper.CurrentAngle
  PLeftFlipper.RotY = a
  'Add any left flipper related animations here
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  PRightFlipper.RotY = a
  'Add any right flipper related animations here
End Sub

'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************
'Const BallSize = 50    'Ball size must be 50
'Const BallMass = 1   'Ball mass must be 1
'Const tnob = 7     'Total number of balls on the playfield including captive balls.
'Const lob = 1      'Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Dim BIP              'Balls in play
BIP = 0
Dim BIPL              'Ball in plunger lane
BIPL = False

'  Standard definitions
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
'Const UseLamps = 1       '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
'Const UseSync = 0
'Const HandleMech = 0
'Const SSolenoidOn = ""     'Sound sample used for this, obsolete.
'Const SSolenoidOff = ""      ' ^
'Const SFlipperOn = ""      ' ^
'Const SFlipperOff = ""     ' ^
'Const SCoin = ""       ' ^


'Const UseVPMModSol = 1   'Old PWM method. Don't use this
'Const UseVPMModSol = 2     'Set to 2 for PWM flashers, inserts, and GI. Requires VPinMame 3.6

'NOTES on UseVPMModSol = 2:
'  - Only supported for S9/S11/DataEast/WPC/Capcom/Whitestar (Sega & Stern)/SAM
'  - All lights on the table must have their Fader model set tp "LED (None)" to get the correct fading effects
'  - When not supported VPM outputs only 0 or 1. Therefore, use VPX "Incandescent" fader for lights

'LoadVPM "03060000", "de.vbs", 3.02  'The "03060000" argument forces user to have VPinMame 3.6

'******************************************************
'  ZTIM: Timers
'******************************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime
  'Add animation stuff here
  'BSUpdate
  'UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  'DoSTAnim         'Standup target animations
  'DoDTAnim         'Drop target animations
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


Dim bulkrnd
bulkrnd=int(rnd*5)+1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


Dim UseVPMDMD
UseVPMDMD=True
LoadVPM "01560000", "SEGA.VBS", 3.26

Const FlashTime = 50 ' ms used for flashing overlay icons

Dim walldrop, dmdhide
If Table1.ShowDT = True Then
  dmdhide=1
  DMD.Visible=True
  DMD1.Visible=False
  For Each walldrop in cabwalls
    walldrop.Visible=True
  Next
  For Each walldrop in dtdisplay
    walldrop.Visible=True
  Next
Else
  dmdhide=1
  DMD.Visible=False
  DMD1.Visible=True

  For Each walldrop in cabwalls
    walldrop.Visible=False
  Next
  For Each walldrop in dtdisplay
    walldrop.Visible=False
  Next
End If

If B2SOn = True then
  dmdhide=0
  DMD.Visible=False
  DMD1.Visible=False
End If

'Const UseSolenoids = True
Const UseLamps    = True
Const UseSync   = True
Const UseGI     = False     'Only WPC games have special GI circuit.
Const SSolenoidOn = "SolOn"       'Solenoid activates
Const SSolenoidOff  = "SolOff"      'Solenoid deactivates
Const SFlipperOn  = "FlipperUp"   'Flipper activated
Const SFlipperOff = "FlipperDown" 'Flipper deactivated
Const SCoin     = "coin3"     'Coin inserted

Dim bubble
Set bubble=Kicker003.CreateSizedBall(6)
bubble.ReflectionEnabled=False
Kicker003.Kick 0, 0
Kicker003.Enabled=False
bubble.Image="bubble"

Sub table1_KeyDown(ByVal keycode)
  'If keycode = LeftMagnaSave then maturepictures=True:mature        'mirrorflash true  'Used for testing
  'If keycode = RightMagnaSave then maturepictures=False:mature       'mirrorflash false  'Used for testing

  If keycode = PlungerKey Then Plunger.Pullback:SoundPlungerPull()
  If keycode = LeftTiltKey Then Nudge 90, 3 : SoundNudgeLeft()    ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 3 : SoundNudgeRight()   ' ^
  If keycode = CenterTiltKey Then Nudge 0, 3 : SoundNudgeCenter()   ' ^
  If keycode = StartGameKey then SoundStartButton
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall
    Else
      SoundPlungerReleaseNoBall
    End If
  End If
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'Controller.SolMask(0)=0
SolCallBack(1)      = "SolTrough"
SolCallBack(2)    = "AutoPlunger"
SolCallBack(3)    = "leftramplockpost"
SolCallback(4)    = "LeftOrbitPost"
SolCallback(5)    = "dtDrop.SolDropUp"
SolCallback(6)    = "SolCLane"
SolCallback(7)    = "SolPeekLeft"'7=Bead Screen Left (peekaboo)
SolCallback(8)    = "SolPeekRight"'8=Bead Screen Right (peekaboo)
SolCallBack(12)   = "bsGrotto.SolOut"
SolCallback(13)   = "bsVUK.SolOut"
SolCallback(14)     = "magopenclose" '14=Magazine Post
'SolCallback(19)    = "SolTease1"'19=Drop Screen Stepper 1
'SolCallback(20)    = "SolTease2"'20=Drop Screen Stepper 2
'SolCallback(21)    = "SolTease3"'21=Drop Screen Stepper 3
 SolCallback(22)  = "splashmotor"
'SolCallback(23)    = "SolTease4"'23=Drop Screen Stepper 4
solcallback(25)="splashflash"
SolCallBack(26) ="mirrorflash"
SolCallback(27)="RearFlashers"
SolCallback(28) = "LeftSlingFlasher"
SolCallBack(29) = "RightSlingFlasher"
SolCallBack(30) = "TripleJackpot"
'31=Centerfold On/Off
'32=Centerfold Open/Close

'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"
'Controller.SolMask(0)=-1

'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire 'rightflipper.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
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
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub

Sub AutoPlunger(Enabled)
  If Enabled Then
    Plunger1.Fire
    SoundPlungerReleaseBall
  End If
End Sub

Set GICallback = GetRef("GIUpdate")
Dim xx
Sub GIUpdate(no, Enabled)
    For each xx in GiLights
        xx.State = ABS(Enabled)
    Next
'    For each xx in GiFlashers
'        xx.visible = Enabled
'    Next
  striplightstimer.enabled = Light010.state
End Sub

Sub leftramplockpost(Enabled)
  If Enabled then
    balllockpost.IsDropped=True:PlaysoundAtVol SoundFX("solenoid",DOFContactors), balllockpost, 1
  Else
    balllockpost.IsDropped=False:PlaysoundAtVol SoundFX("solenoid",DOFContactors), balllockpost, 1
  End If
End Sub

Dim mirrorrnd
Sub mirrorflash(Enabled)
  If Enabled then
    If maturepictures=True Then
      Randomize
      mirrorrnd=int(rnd*5)+1
      Pmirror.Image="mirror-nude-"&mirrorrnd&"a"
      vpmTimer.AddTimer 120,"FadeMirror1"
      vpmTimer.AddTimer 240,"FadeMirror2"
      vpmTimer.AddTimer 360,"FadeMirror3"
      vpmTimer.AddTimer 480,"FadeMirror4"
      vpmTimer.AddTimer 600,"FadeMirror5"
    Else
      Pmirror.Image="mirror-no-nude_a"
      vpmTimer.AddTimer 120,"FadeMirror1"
      vpmTimer.AddTimer 240,"FadeMirror2"
      vpmTimer.AddTimer 360,"FadeMirror3"
      vpmTimer.AddTimer 480,"FadeMirror4"
      vpmTimer.AddTimer 600,"FadeMirror5"
    End If
  Else
    Pmirror.Image="mirrornotlit"
  End If
End Sub

sub FadeMirror1(dummy)
  If maturepictures=True Then
    Pmirror.Image="mirror-nude-"&mirrorrnd&"b"
  Else
    Pmirror.Image="mirror-no-nude_b"
  End If
End Sub
sub FadeMirror2(dummy)
  If maturepictures=True Then
    Pmirror.Image="mirror-nude-"&mirrorrnd&"c"
  Else
    Pmirror.Image="mirror-no-nude_c"
  End If
End Sub
sub FadeMirror3(dummy)
  If maturepictures=True Then
    Pmirror.Image="mirror-nude-"&mirrorrnd&"d"
  Else
    Pmirror.Image="mirror-no-nude_d"
  End If
End Sub
sub FadeMirror4(dummy)
  If maturepictures=True Then
    Pmirror.Image="mirror-nude-"&mirrorrnd&"dd"
  Else
    Pmirror.Image="mirror-no-nude_dd"
  End If
End Sub
sub FadeMirror5(dummy)
  If maturepictures=True Then
    Pmirror.Image="mirror-nude-"&mirrorrnd&"on"
  Else
    Pmirror.Image="mirror-no-nude"
  End If
End Sub

mature
'Dim bulkrnd
Sub mature
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    PCFTop.Image="centerfold-top-nude-"&bulkrnd:PCFMiddle.Image="centerfold-middle-nude-"&bulkrnd:PCFBottom.Image="centerfold-bottom-nude-"&bulkrnd
  Else
    PCFTop.Image="centerfold-top-non-nude":PCFMiddle.Image="centerfold-middle-non-nude":PCFBottom.Image="centerfold-bottom-no-nude"
  End If
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    Pmaginside.Image="maginside-nude-"&bulkrnd
  Else
    Pmaginside.Image="maginside-non-nude"
  End If
  If maturepictures=True Then
    bg1.Image="pb-background-1-off":bg2.Image="pb-background-1-off"
    pf_Clothes.visible = 0
    Apron_Clothes.visible = 0
  Else
    bg1.Image="pb-background-2-off":bg2.Image="pb-background-2-off"
    pf_Clothes.visible = 1
    Apron_Clothes.visible = 1
  End If
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    Ppeekaboo.Image="peek-nude-"&bulkrnd
  Else
    Ppeekaboo.Image="peek-no-nude-1"
  End If
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    Pteasepic.Image="tease-nude-"&bulkrnd
  Else
    Pteasepic.Image="tease-non-nude"
  End If
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    Psplashtriangle001.Image="splashtriangle-nude-"&bulkrnd
  Else
    Psplashtriangle001.Image="splashtriangle-non-nude"
  End If
  If maturepictures=True Then
    Randomize
    bulkrnd=int(rnd*5)+1
    Pmirror.Image="mirror-nude-"&bulkrnd
    Pmirror.Image="mirrornotlit"
  End If
End Sub

Dim splashdir
Sub splashmotor(Enabled)
  If Enabled then splashtimer.Enabled=True:splashdir=1 Else splashtimer.Enabled=True:splashdir=-1
End Sub

Sub splashtimer_Timer()
  Psplashtriangle001.RotY=Psplashtriangle001.RotY+splashdir
  If Psplashtriangle001.RotY=360 then Psplashtriangle001.RotY=0
  If Psplashtriangle001.RotY=200 or Psplashtriangle001.RotY=-160 then splashtimer.Enabled=False
End Sub

Dim magmov
Sub magopenclose(Enabled)
  If Enabled then magmov=1 Else magmov=-1
  magmove.Enabled=True
End Sub

Sub magmove_Timer()
  Pmagcover.RotZ=Pmagcover.RotZ+magmov
  If Pmagcover.RotZ = 0 or Pmagcover.RotZ = 100 then magmove.Enabled=False
End Sub

Sub RearFlashers(Enabled)
  If Enabled then f27a.Visible=True:f27b.Visible=True:f27c.Visible=True:Primitive033.DisableLighting=1:Primitive034.DisableLighting=1:f27d.State=False Else f27a.Visible=False:f27b.Visible=False:f27c.Visible=False:Primitive033.DisableLighting=0.2:Primitive034.DisableLighting=0.2:f27d.State=True
End Sub

Sub LeftSlingFlasher(Enabled)
  If Enabled then f28.Visible=True:Primitive007.DisableLighting=1:f28a.State=False Else f28.Visible=False:Primitive007.DisableLighting=0.2:f28a.State=True
End Sub

Sub RightSlingFlasher(Enabled)
  If Enabled then f29.Visible=True:Primitive011.DisableLighting=1:f29a.State=False:f29b.State=False:f29c.State=False:f29d.State=1 Else f29.Visible=False:Primitive011.DisableLighting=0.6:f29a.State=True:f29b.State=True:f29c.State=True:f29d.State=0
End Sub

Sub splashflash(Enabled)
  If Enabled Then Flasher25.Visible=True Else Flasher25.Visible=False
End Sub

Sub TripleJackpot(Enabled)
    If Enabled then Flasher30.Visible=True Else Flasher30.Visible=False
End Sub

Dim chainleft
Sub SolPeekLeft(Enabled)
  'If Enabled Then PeekReel.SetValue 0
  If Enabled Then
    For Each chainleft in chainopenleft
      chainleft.Visible=True
    Next
    For Each chainleft in chainclosedleft
      chainleft.Visible=False
    Next
  Else
    For Each chainleft in chainopenleft
      chainleft.Visible=False
    Next
    For Each chainleft in chainclosedleft
      chainleft.Visible=True
    Next
    cball.VelX=3:cball.VelY=3
  End If
End Sub

Dim chainright
Sub SolPeekRight(Enabled)
  'If Enabled Then PeekReel.SetValue 1
  If Enabled Then
    For Each chainright in chainopenright
      chainright.Visible=True
    Next
    For Each chainright in chainclosedright
      chainright.Visible=False
    Next
  Else
    For Each chainright in chainopenright
      chainright.Visible=False
    Next
    For Each chainright in chainclosedright
      chainright.Visible=True
    Next
    cball.VelX=3:cball.VelY=3
  End If
End Sub

Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    BIP = BIP + 1
    RandomSoundBallRelease BallRelease
    vpmTimer.PulseSw 15
  End If
End Sub

Sub Drain_Hit                         '11
  BIP = BIP - 1
  bsTrough.AddBall Me
  RandomSoundDrain Drain
End Sub

Sub LeftOrbitPost(Enabled)
  If Enabled Then
    LO.IsDropped=0
  Else
    LO.IsDropped=1
  End If
End Sub

Sub SolCLane(Enabled)
  If Enabled Then
    CLane.IsDropped=0
  Else
    CLane.IsDropped=1
  End If
End Sub

Dim bsTrough,bsVUK,dtDrop,bsGrotto,mTease,mCenterfold
Const cRegistryName = "Playboy"
Const cOptionsName = "Options"
Dim vpmDips
Dim PlayboyOptions
Dim OptRom
Dim OptYield
Dim OptSkip
Sub PlayboyShowDips
  If Not IsObject(vpmDips) Then
    Set vpmDips = New cvpmDips
    With vpmDips
      .AddForm 100, 100, "Playboy Game Settings"
      .AddFrameExtra 0, 0, 250, "ROM Version", &Hf0, Array("International Display (default)", &H00, "France Display", &H10, "German Display", &H20, "Italy Display", &H40, "Spain Display", &H80)
        .AddChkExtra 7,100,250,Array("Enable YieldTime (for slow computers)", &H100)
      .AddChkExtra 7,120,250,Array("Do not show this menu again at startup", &H200)
      .AddLabel 7,140,250,20,"Press F6 during play to bring up this menu."
      .AddLabel 7,160,250,20,"Quit and restart game for changes to take effect."
    End With
  End If
  PlayboyOptions = vpmDips.ViewDipsExtra(PlayboyOptions)
  SaveValue cRegistryName,cOptionsName,PlayboyOptions
  PlayboySetOptions
End Sub

Set vpmShowDips = GetRef("PlayboyShowDips")
Sub PlayboySetOptions
  OptRom = "playboys"
  OptYield = False
  OptSkip = False
  If PlayboyOptions And &H10 Then OptRom = "playboyf"
  If PlayboyOptions And &H20 Then OptRom = "playboyg"
  If PlayboyOptions And &H40 Then OptRom = "playboyi"
  If PlayboyOptions And &H80 Then OptRom = "playboyl"
  If PlayboyOptions And &H100 Then OptYield = True
  If PlayboyOptions And &H200 Then OptSkip = True
End Sub

Const tnob = 5 ' total number of balls

'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

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

Sub RollingUpdate()
  Dim b
  Dim BOT
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 To tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(BOT)
    If BallVel(BOT(b)) > 1 And BOT(b).z < 30 Then
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
  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

Sub table1_Init

 '  SaveValue cRegistryName,cOptionsName,0  ' This will clear your Registry settings if uncommented

  PlayboyOptions = CInt("0" & LoadValue(cRegistryName,cOptionsName))

  PlayboySetOptions

  If optSkip = False Then vpmShowDips

  PlayboySetOptions
  With Controller
    If optYield = True Then : If .Version >= "01500000" Then Table1.YieldTime = 2 : End If
    .GameName=OptRom
    Plunger1.PullBack
    table1.Yieldtime=1
    .SplashInfoLine = "Playboy (Stern 2002)" & vbNewLine & "VPX version by Cheese3075, HiRez00 and Rascal"
    .ShowTitle = False
    '.ShowDMDOnly = 1
    '.Hidden = dmdhide
    .ShowFrame = False
    .HandleMechanics=0
    '.SetDisplayPosition 13,229
  .DIP(0)=&H00
  On Error Resume Next
  .Run
    If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch = 56
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
  vpmMapLights clights

    Set bsTrough = New cvpmBallStack
  bsTrough.InitSw 0,14,13,12,11,0,0,0
  bsTrough.InitKick BallRelease,90,3
  'bsTrough.InitEntrySnd "Solenoid", "Solenoid"
  'bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors), SoundFX("Solon",DOFContactors)
  bsTrough.Balls=4

  ' Thalamus, more randomness pls
  Set bsVUK=New cvpmBallStack
  bsVUK.InitSw 0,35,0,0,0,0,0,0
  bsVUK.InitKick Kicker2,270,8
  bsVUK.KickForceVar = 3
  bsVUK.KickAngleVar = 3
  'bsVUK.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("solon",DOFContactors)

  Set bsGrotto=New cvpmBallStack
  bsGrotto.InitSw 0,34,0,0,0,0,0,0
  bsGrotto.InitKick Kicker4,40,6
  bsGrotto.KickForceVar = 3
  bsGrotto.KickAngleVar = 3
  'bsGrotto.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("solon",DOFContactors)

  set dtDrop=new cvpmDropTarget
  dtDrop.InitDrop wall5,26
  dtDrop.InitSnd SoundFX("flapclos",DOFDropTargets),SoundFX("flapopen",DOFDropTargets)

  Set mTease=New cvpmMech
  mTease.MType=vpmMechStepSol+vpmMechReverse+vpmMechLinear
  mTease.Sol1=19
  mTease.Sol2=21
  mTease.Length=90
  mTease.Steps=120
  mTease.AddSw 52,0,0
  mTease.Callback=GetRef("UpdateTease")
  mTease.Start

  Set mCenterfold=New cvpmMech
  mCenterfold.MType=vpmMechStepSol+vpmMechReverse+vpmMechLinear
  mCenterfold.Sol1=31
  mCenterfold.Sol2=32
  mCenterfold.Length=165
  mCenterfold.Steps=170
  mCenterfold.AddSw 43,0,0
  mCenterfold.Callback=GetRef("UpdateCenterfold")
  mCenterfold.Start

 End with
End Sub

Sub UpdateTease(aNewPos,aSpeed,aLastPos)
  If aNewPos>-30 And aNewPos<131 Then
    Pteasedoor.Z= (150 - aNewPos) - 60
  End If
End Sub

 Sub UpdateCenterfold(aNewPos,aSpeed,aLastPos)
  If aNewPos>-1 And aNewPos<170 Then
    If aNewPos<=5 then aNewPos=0
    If aNewPos<85 then PCFTop.ObjRotX=aNewPos*2
    If aNewPos>=85 then PCFBottom.ObjRotX=-(aNewPos-85)*2 Else PCFBottom.ObjRotX=0
  End If
End Sub

Sub Trigger13_Hit:Controller.Switch(9)=1:End Sub        '9
Sub Trigger13_unHit:Controller.Switch(9)=0:End Sub
Sub Trigger4_Hit:Controller.Switch(10)=1:End Sub        '10
Sub Trigger4_unHit:Controller.Switch(10)=0:End Sub
Sub Trigger6_Hit:Controller.Switch(16)=1:BIPL=1:End Sub       '16
Sub Trigger6_unHit:Controller.Switch(16)=0:BIPL=0:End Sub
Sub Trigger7_Hit:Controller.Switch(17)=1:End Sub        '17
Sub Trigger7_unHit:Controller.Switch(17)=0:End Sub
Sub Trigger8_Hit:Controller.Switch(18)=1:End Sub        '18
Sub Trigger8_unHit:Controller.Switch(18)=0:End Sub
Sub Trigger14_Hit:Controller.Switch(19)=1:End Sub       '19
Sub Trigger14_unHit:Controller.Switch(19)=0:End Sub
Sub Trigger9_Hit:Controller.Switch(21)=1:End Sub        '21
Sub Trigger9_unHit:Controller.Switch(21)=0:End Sub
Sub Trigger1_Hit:Controller.Switch(22)=1:End Sub            '22
Sub Trigger1_unHit:Controller.Switch(22)=0:End Sub
Sub Trigger2_Hit:Controller.Switch(23)=1:End Sub '23
Sub Trigger2_unHit:Controller.Switch(23)=0:End Sub
Sub Trigger3_Hit:Controller.Switch(24)=1:End Sub            '24
Sub Trigger3_unHit:Controller.Switch(24)=0:End Sub
Sub Trigger12_Hit:Controller.Switch(25)=1:End Sub       '25
Sub Trigger12_unHit:Controller.Switch(25)=0:End Sub
Sub Wall5_Hit:dtDrop.Hit 1:End Sub                '26
Sub Trigger5_Hit:Controller.Switch(27)=1:End Sub        '27
Sub Trigger5_unHit:Controller.Switch(27)=0:End Sub
'29=Triangle Mech 1 (Right) on right ramp
'30=Triangle Mech 2 (Left) on right ramp
Sub T32_Hit:vpmTimer.PulseSw 32:End Sub             '32
Sub Target001_Hit:vpmTimer.PulseSw 33:End Sub         '33
Sub Kicker4_Hit:bsGrotto.AddBall Me:SoundSaucerLock:End Sub   '34
Sub Kicker4_unHit:SoundSaucerKick 1, Kicker4:End Sub      '34
Sub Kicker1_Hit:bsVUK.AddBall Me:End Sub            '35
Sub Kicker1_unHit:SoundSaucerKick 1, Kicker1:End Sub            '35
Sub Trigger17_Hit:Controller.Switch(38)=1 : WireRampOn False: End Sub       '38
Sub Trigger17_unHit:Controller.Switch(38)=0:End Sub
Sub Trigger15_Hit:Controller.Switch(39)=1:End Sub       '39
Sub Trigger15_unHit:Controller.Switch(39)=0:End Sub
Sub Trigger16_Hit:Controller.Switch(40)=1:End Sub       '40
Sub Trigger16_unHit:Controller.Switch(40)=0:End Sub

Dim MyBall
Sub Trigger10_Hit
  Set MyBall=ActiveBall
  'If MyBall.VelY>0 Then MyBall.VelX=MyBall.VelX+2
  Controller.Switch(41)=1
End Sub       '41

Sub Trigger10_unHit:Controller.Switch(41)=0:End Sub
Sub Trigger11_Hit:Controller.Switch(42)=1:End Sub       '42
Sub Trigger11_unHit:Controller.Switch(42)=0:End Sub
'43=Centerfold 1 (Closed)
'44=Centerfold 2 (Open)
Sub Bumper1_Hit:vpmTimer.PulseSw 49:bumpershake:RandomSoundBumperTop Bumper1:End Sub      '49
Sub Bumper2_Hit:vpmTimer.PulseSw 50:bumpershake:RandomSoundBumperMiddle Bumper2:End Sub     '50
Sub Bumper3_Hit:vpmTimer.PulseSw 51:bumpershake:RandomSoundBumperBottom Bumper3:End Sub     '51

Sub LeftSlingshot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 59
  RandomSoundSlingshotLeft Sling2
End Sub

Sub RightSlingshot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 62
  RandomSoundSlingshotRight Sling1
End Sub

Dim chainmovx:chainmovx=1
Sub bumpershake()
  cball.VelY=-.5
  If chainmovx=1 then cball.VelX=-.5:chainmovx=2 Else cball.VelX=.5:chainmovx=1
End Sub
'52=Tease Screw Limit on backsplash board
'53=Tournament Button
                                '56 Tilt
Sub LeftOutlane_Hit:Controller.Switch(57)=1:End Sub       '57
Sub LeftOutlane_unHit:Controller.Switch(57)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(58)=1:End Sub        '58
Sub LeftInlane_unHit:Controller.Switch(58)=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(60)=1:End Sub      '60
Sub RightOutlane_unHit:Controller.Switch(60)=0:End Sub
Sub RightInlane_Hit:Controller.Switch(61)=1:End Sub       '61
Sub RightInlane_unHit:Controller.Switch(61)=0:End Sub
Sub Kicker7_Hit:WireRampOff : Me.DestroyBall : Set mball = Kicker004.CreateBall : Kicker8.TimerEnabled = True : End Sub
Sub Kicker8_Timer()
  Kicker004.DestroyBall
  Set mball=Kicker8.CreateBall
  Kicker8.Kick 180, 0
  Kicker8.TimerEnabled=False
End Sub
Sub Kicker8_unHit:WireRampOn False:End Sub
Sub Trigger18_Hit : ActiveBall.VelY = 2 : WireRampOff : End Sub
Sub Trigger19_Hit : ActiveBall.VelY = 2 : WireRampOff : End Sub

Sub Trigger001_Hit()
  ActiveBall.VelX=0
End Sub

Sub Kicker001_Hit()
  Kicker002.Enabled=True
  Kicker001.TimerEnabled=True
End Sub

Sub Kicker001_Timer()
  Kicker001.DestroyBall
  Set mball=Kicker002.CreateBall
  Kicker002.Kick 180, 3
  Kicker002.Enabled=False
  Kicker001.TimerEnabled=False
End Sub

Sub Timer002_Timer()
  If Light74.State = 1 then Flasher002.Visible=True Else Flasher002.Visible=False
  If Light75.State = 1 then Flasher001.Visible=True Else Flasher001.Visible=False
End Sub

Dim slswitch:slswitch=0
Dim slitems
Sub striplightstimer_Timer()
  If slswitch = 0 Then
    slswitch=1
    For Each slitems in striplights1
      slitems.Visible=True
    Next
    For Each slitems in striplights2
      slitems.Visible=False
    Next
  Else
    slswitch=0
    For Each slitems in striplights1
      slitems.Visible=False
    Next
    For Each slitems in striplights2
      slitems.Visible=True
    Next
  End If
End Sub

Sub mirrorballtrigger_Hit()
  mirrorballtimer.Enabled=True
End Sub

Sub mirrorballtrigger_Unhit()
  mirrorballtimer.Enabled=False
End Sub

Sub mirrorballtimer_Timer()
  On Error Resume Next
  Pmirrorball.Z=mball.Y-20
  Pmirrorball.X=mball.X
End Sub

Dim mball
Sub mirrorballcreate_Hit()
  mirrorballcreate.DestroyBall
  Set mball=mirrorballcreate.CreateBall
  mirrorballcreate.kick 0, 0
  mirrorballcreate.Enabled = False
End Sub

Sub Gate1_Hit()
  mirrorballcreate.Enabled = True
End Sub

' *** Nudge
Dim mMagnet, cBall

Sub WobbleMagnet_Init
   Set mMagnet = new cvpmMagnet
   With mMagnet
    .InitMagnet WobbleMagnet, .9
    .Size = 150
    .CreateEvents mMagnet
    .MagnetOn = True
    .GrabCenter = False
   End With
  Set cBall = ckicker.CreateBall
' Set cBall = ckicker.CreateSizedBallWithMass(25, 1)
  ckicker.Kick 0,0:mMagnet.addball cball
End Sub

Dim chain
Sub chaintimer_Timer()
  For Each chain in chainsclosed
    chain.RotX=cBall.X-ckicker.X:chain.RotZ=cBall.Y-ckicker.Y
  Next
End Sub

Sub light45timer_Timer()
  If Light45.State=1 then grottolight.Visible=True Else grottolight.Visible=False
  If Light67.State=1 then L67a.Visible=True:L67b.Visible=True Else L67a.Visible=False:L67b.Visible=False
  If Light68.State=1 then L68.Visible=True Else L68.Visible=False
  If Light69.State=1 then L69.Visible=True Else L69.Visible=False
End Sub

Sub ballhitchains_Hit()
  cball.VelX=-4:cball.VelY=-4
  PlaySoundAtVol "chain-01a", ActiveBall, 1
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
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

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
dim RampBalls(5,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(5)

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


Function BallPitch(BOT) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(BOT), 1, -1000, 60, 10000)
End Function

Function BallPitchV(BOT) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(BOT), 1, -4000, 60, 7000)
End Function

'Ramp triggers
Sub ramptrigger01_hit()
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub ramptrigger02_hit()
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub ramptrigger03_hit()
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTriggerL_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTriggerLM_Hit
  WireRampOff
  WireRampOn False
End Sub

Sub RampTriggerLE_Hit
    WireRampOff
  WireRampOn False
End Sub

'***

Sub RampTriggerM_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTriggerME_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

'***

Sub RampTrigger5_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger6_Hit
  WireRampOff
End Sub

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

    '////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
    '//  This part in the script is an entire block that is dedicated to the physics sound system.
    '//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.


    '///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
    Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
    Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

    CoinSoundLevel = 1                                                                                                                'volume level; range [0, 1]
    NudgeLeftSoundLevel = 1                                                                                                        'volume level; range [0, 1]
    NudgeRightSoundLevel = 1                                                                                                'volume level; range [0, 1]
    NudgeCenterSoundLevel = 1                                                                                                'volume level; range [0, 1]
    StartButtonSoundLevel = 0.1                                                                                                'volume level; range [0, 1]
    PlungerReleaseSoundLevel = 0.8 '1 wjr                                                                                        'volume level; range [0, 1]
    PlungerPullSoundLevel = 1                                                                                                'volume level; range [0, 1]
    RollingSoundFactor = 1.1/5

    '///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
    Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
    Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
    Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

    FlipperUpAttackMinimumSoundLevel = 0.010                                                           'volume level; range [0, 1]
    FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
    FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
    FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
    FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
    FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
    SlingshotSoundLevel = 0.95                                                                                                'volume level; range [0, 1]
    BumperSoundFactor = 4.25                                                                                                'volume multiplier; must not be zero
    KnockerSoundLevel = 1                                                                                                         'volume level; range [0, 1]

    '///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
    Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
    Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
    Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
    Dim SaucerLockSoundLevel, SaucerKickSoundLevel

    BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
    RubberStrongSoundFactor = 0.055/5                                                                                        'volume multiplier; must not be zero
    RubberWeakSoundFactor = 0.075/5                                                                                        'volume multiplier; must not be zero
    RubberFlipperSoundFactor = 0.075/5                                                                                'volume multiplier; must not be zero
    BallBouncePlayfieldSoftFactor = 0.025                                                                        'volume multiplier; must not be zero
    BallBouncePlayfieldHardFactor = 0.025                                                                        'volume multiplier; must not be zero
    DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
    WallImpactSoundFactor = 0.075                                                                                        'volume multiplier; must not be zero
    MetalImpactSoundFactor = 0.075/3
    SaucerLockSoundLevel = 0.8
    SaucerKickSoundLevel = 0.8

    '///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

    Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

    GateSoundLevel = 0.5/5                                                                                                        'volume level; range [0, 1]
    TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
    DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
    RolloverSoundLevel = 0.25                                                                      'volume level; range [0, 1]
    SpinnerSoundLevel = 0.5                                                                      'volume level; range [0, 1]

    '///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
    Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

    DrainSoundLevel = 0.8                                                                                                                'volume level; range [0, 1]
    BallReleaseSoundLevel = 1                                                                                                'volume level; range [0, 1]
    BottomArchBallGuideSoundFactor = 0.2                                                                        'volume multiplier; must not be zero
    FlipperBallGuideSoundFactor = 0.015                                                                                'volume multiplier; must not be zero

    '///////////////////////-----Loops and Lanes-----///////////////////////
    Dim ArchSoundFactor
    ArchSoundFactor = 0.025/5                                                                                                        'volume multiplier; must not be zero


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
        tmp = 2100 * 2 / tableheight-1

            if tmp > 7000 Then
                    tmp = 7000
            elseif tmp < -7000 Then                     'cheese uncomment
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
        tmp = 1060 * 2 / tablewidth-1

            if tmp > 7000 Then
                    tmp = 7000
            elseif tmp < -7000 Then                     'cheese uncomment
                    tmp = -7000
            end if

        If tmp > 0 Then
            AudioPan = Csng(tmp ^10)
        Else
            AudioPan = Csng(-((- tmp) ^10) )
        End If
    End Function

    Function Vol(BOT) ' Calculates the volume of the sound based on the ball speed
            Vol = Csng(BallVel(BOT) ^2)
    End Function

    Function Volz(BOT) ' Calculates the volume of the sound based on the ball speed
            Volz = Csng((BOT.velz) ^2)
    End Function

    Function Pitch(BOT) ' Calculates the pitch of the sound based on the ball speed
            Pitch = BallVel(BOT) * 20
    End Function

    Function BallVel(BOT) 'Calculates the ball speed
            BallVel = INT(SQR((BOT.VelX ^2) + (BOT.VelY ^2) ) )
    End Function

    Function VolPlayfieldRoll(BOT) ' Calculates the roll volume of the sound based on the ball speed
            VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(BOT) ^3)
    End Function

    Function PitchPlayfieldRoll(BOT) ' Calculates the roll pitch of the sound based on the ball speed
            PitchPlayfieldRoll = BallVel(BOT) ^2 * 15
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
    Sub RandomSoundBallBouncePlayfieldSoft(ActiveBall)
            Select Case Int(Rnd*9)+1
                    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(ActiveBall) * BallBouncePlayfieldSoftFactor, ActiveBall
                    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.5, ActiveBall
                    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.8, ActiveBall
                    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.5, ActiveBall
                    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(ActiveBall) * BallBouncePlayfieldSoftFactor, ActiveBall
                    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.2, ActiveBall
                    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.2, ActiveBall
                    Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.2, ActiveBall
                    Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(ActiveBall) * BallBouncePlayfieldSoftFactor * 0.3, ActiveBall
            End Select
    End Sub

    Sub RandomSoundBallBouncePlayfieldHard(ActiveBall)
            PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(ActiveBall) * BallBouncePlayfieldHardFactor, ActiveBall
    End Sub

    '/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
    Sub RandomSoundDelayedBallDropOnPlayfield(ActiveBall)
            Select Case Int(Rnd*5)+1
                    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, ActiveBall
                    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, ActiveBall
                    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, ActiveBall
                    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, ActiveBall
                    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, ActiveBall
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

    Const RelayFlashSoundLevel = 0.315                                                                        'volume level; range [0, 1];
    Const RelayGISoundLevel = 1.05                                                                        'volume level; range [0, 1];

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
    '                                        End Mechanical Sounds
    '/////////////////////////////////////////////////////////////////


'******************************************************
' ZPHY:  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners         https://youtu.be/AXX3aen06FM?si=Xqd-rcaqTlgEd_wx
' Adding nFozzy roth physics : pt2 flipper physics          https://youtu.be/VSBFuK2RCPE?si=i8ne8Ao2co8rt7fy
' Adding nFozzy roth physics : pt3 other elements           https://youtu.be/JN8HEJapCvs?si=hvgMOk-ej1BEYjJv
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
' | Force         | 12-15    |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 3-5  |
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
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
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
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

'*******************************************
' Early 90's and after

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

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
    x.AddPt "Velocity", 2, 0.27, 1
    x.AddPt "Velocity", 3, 0.3, 1
    x.AddPt "Velocity", 4, 0.35, 1
    x.AddPt "Velocity", 5, 0.6, 1 '0.982
    x.AddPt "Velocity", 6, 0.62, 1.0
    x.AddPt "Velocity", 7, 0.702, 0.968
    x.AddPt "Velocity", 8, 0.95,  0.968
    x.AddPt "Velocity", 9, 1.03,  0.945
    x.AddPt "Velocity", 10, 1.5,  0.945
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
'   Otherwise it should function exactly the same as before\
' modified 2024 by rothbauerw
' Added Reprocessballs for flipper collisions (LF.Reprocessballs Activeball and RF.Reprocessballs Activeball must be added to the flipper collide subs
' Improved handling to remove correction for backhand shots when the flipper is raised

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
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

  Public Sub AddBall(ActiveBall)
    Dim x
    For x = 0 To UBound(balls)
      If IsEmpty(balls(x)) Then
        Set balls(x) = ActiveBall
        Exit Sub
      End If
    Next
  End Sub

  Private Sub RemoveBall(ActiveBall)
    Dim x
    For x = 0 To UBound(balls)
      If TypeName(balls(x) ) = "IBall" Then
        If ActiveBall.ID = Balls(x).ID Then
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
      If Not IsEmpty(balls(x)) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        balldata(x).Data = balls(x)
      End If
    Next
    FlipStartAngle = Flipper.currentangle
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub

  Public Sub ReProcessBalls(ActiveBall) 'save data of balls in flipper range
    If FlipperOn() Then
      Dim x
      For x = 0 To UBound(balls)
        If Not IsEmpty(balls(x)) Then
          if balls(x).ID = ActiveBall.ID Then
            If isempty(balldata(x).ID) Then
              balldata(x).Data = balls(x)
            End If
          End If
        End If
      Next
    End If
  End Sub

  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(ActiveBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
      Ycoef = 1

      'y safety Exit
      If ActiveBall.VelY > -8 Then 'ball going down
        RemoveBall ActiveBall
        Exit Sub
      End If

      'Find balldata. BallPos = % on Flipper
      For x = 0 To UBound(Balls)
        If ActiveBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(ActiveBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(ActiveBall.Y, YcoefIn, YcoefOut)                       'find safety coefficient 'ycoef' data
        NoCorrection = 1
      Else
        checkHit = 50 + (20 * BallPos)

        If BalltoFlip > checkHit or (PartialFlipCoef < 0.5 and BallPos > 0.22) Then
          NoCorrection = 1
        Else
          NoCorrection = 0
        End If
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        'If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then ActiveBall.Velx = ActiveBall.Velx*VelCoef
        If Enabled Then ActiveBall.Vely = ActiveBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled and NoCorrection = 0 Then ActiveBall.VelX = ActiveBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
      End If
      If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall ActiveBall
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
Function BallSpeed(BOT) 'Calculates the ball speed
  BallSpeed = Sqr(BOT.VelX^2 + BOT.VelY^2 + BOT.VelZ^2)
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
  Public Property Let Data(ActiveBall)
    With ActiveBall
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
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines
'  - Include FlipperActivate and FlipperDeactivate in the Flipper solenoid subs

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
     BOT = GetBalls

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

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
  DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(angle + 90)) + Flipper.y)
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

Dim LFPress, RFPress, ULFPress, LFCount, RFCount
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
'Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
'   Const EOSReturn = 0.035  'mid 80's to early 90's
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
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, BOT
    BOT = GetBalls

    For b = 0 To UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
      End If
    Next
  End If
End Sub

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

Const LiveDistanceMin = 5  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
Const BaseDampen = 0.55

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir, LiveDist
    Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce   'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime
    CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)

    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then   'Perfect catch only when catch time happens in the beginning of the window
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)  'Partial catch when catch happens a bit late
        End If

        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0

        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        Elseif LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
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

' Put all the Post and Pin objects in dPosts collection. Make sure dPosts fires hit events.
Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

' This collection contains the bottom sling posts. They are not in the dPosts collection so that the TargetBouncer is not applied to them, but they should still have dampening applied
' If you experience airballs with posts or targets, consider adding them to this collection
Sub NoTargetBouncer_Hit
    RubbersD.dampen ActiveBall
End Sub

' Put all the Sleeve objects in dSleeves collection. Make sure dSleeves fires hit events.
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

  Public Sub Dampen(ActiveBall)
    If threshold Then
      If BallSpeed(ActiveBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(ActiveBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(ActiveBall) / (cor.ballvel(ActiveBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " In vel:" & Round(cor.ballvel(ActiveBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
    "actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
    If Print Then Debug.print Round(cor.ballvel(ActiveBall.id),2) & ", " & Round(desiredcor,3)

    ActiveBall.velx = ActiveBall.velx * coef
    ActiveBall.vely = ActiveBall.vely * coef
    ActiveBall.velz = ActiveBall.velz * coef
    If debugOn Then TBPout.text = str
  End Sub

  Public Sub Dampenf(ActiveBall, parm) 'Rubberizer is handle here
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(ActiveBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(ActiveBall) / (cor.ballvel(ActiveBall.id) + 0.0001)
    coef = desiredcor / realcor
    If Abs(ActiveBall.velx) < 2 And ActiveBall.vely < 0 And ActiveBall.vely >  - 3.75 Then
      ActiveBall.velx = ActiveBall.velx * coef
      ActiveBall.vely = ActiveBall.vely * coef
      ActiveBall.velz = ActiveBall.velz * coef
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
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7-1.0

sub TargetBouncer(ActiveBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and ActiveBall.z < 30 then
        'debug.print "velx: " & ActiveBall.velx & " vely: " & ActiveBall.vely & " velz: " & ActiveBall.velz
        vel = BallSpeed(activeBall)
        if ActiveBall.velx = 0 then vratio = 1 else vratio = ActiveBall.vely/ActiveBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        ActiveBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        ActiveBall.velx = sgn(ActiveBall.velx) * sqr(abs((vel^2 - ActiveBall.velz^2)/(1+vratio^2)))
        ActiveBall.vely = ActiveBall.velx * vratio
        'debug.print "---> velx: " & ActiveBall.velx & " vely: " & ActiveBall.vely & " velz: " & ActiveBall.velz
        'debug.print "conservation check: " & BallSpeed(ActiveBall)/vel
  end if
end sub


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



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************

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
'Function RotPoint(x,y,angle)
' dim rx, ry
' rx = x*dCos(angle) - y*dSin(angle)
' ry = x*dSin(angle) + y*dCos(angle)
' RotPoint = Array(rx,ry)
'End Function

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

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, girllights, pfgirllights, gl, maturepictures, alternateflippers
Dim dspTriggered : dspTriggered = False

'*******************************************
'  ZOPT: User Options
'*******************************************
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim StagedFlippers                              ' Staged Flippers. 0 = Disabled, 1 = Enabled


Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 12, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Black Light") )
    UpdateOptions

  ' Nude mode
    x = Table1.Option("Nude mode", 0, 1, 1, 0, 0, Array("Off", "On") )
    If x = 1 Then
    maturepictures=True
    mature
  Else
    maturepictures=False
    mature
  End If

    ' Calendar girl lights
    x = Table1.Option("Calendar girl lights", 0, 1, 1, 0, 0, Array("On (Default)", "Off") )
    If x = 1 Then girllights = False Else girllights = True
    UpdateOptions

    ' Alternate flippers decals
    x = Table1.Option("Alternate flippers decals", 0, 1, 1, 0, 0, Array("Off", "On") )
    If x = 1 Then alternateflippers = True Else alternateflippers = False
  UpdateOptions

  'Playfield girl lights
    x = Table1.Option("Playfield girl lights", 0, 1, 1, 0, 0, Array("On (Default)", "Off") )
    If x = 1 Then pfgirllights = False Else pfgirllights = True
  UpdateOptions

    ' Staged Flippers 'NOT USED
    'x = Table1.Option("Staged Flippers", 0, 1, 1, 0, 0, Array("Disabled", "Enabled") )
  '    If x = 1 Then StagedFlippers = 1 Else StagedFlippers = 0

    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If




'GAME OPTIONS BELOW
'SET OPTIONS TO TRUE OR FALSE
'-------------------------------------------------------------------------------------------------------------
'alternateflippers=True '<--- True = Alternative Flipper Decals, False = Factory Default Flipper Decals.

End Sub

Sub UpdateOptions
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT Warm 0"
        Case 7:table1.ColorGradeImage = "LUT Warm 1"
        Case 8:table1.ColorGradeImage = "LUT Warm 2"
        Case 9:table1.ColorGradeImage = "LUT Warm 3"
        Case 10:table1.ColorGradeImage = "LUT Warm 4"
        Case 11:table1.ColorGradeImage = "LUT Warm 5"
        Case 12:table1.ColorGradeImage = "LUT Black Light"
    End Select

  If alternateflippers=True Then
    PLeftFlipper.Image="HiRez00_left-flipper-alt"
    PRightFlipper.Image="HiRez00_right-flipper-alt"
  Else
    PLeftFlipper.Image="lflipper AP"
    PRightFlipper.Image="rflipper AP"
  End If

  If girllights=True then
    For Each gl in monthgirls
      gl.Visible=True
    Next
  Else
    For Each gl in monthgirls
      gl.Visible=False
    Next
  End If

  If pfgirllights=True then
    Light051.state=1
    Light052.state=1
    Light053.state=1
  Else
    Light051.state=0
    Light052.state=0
    Light053.state=0
  End If
End Sub

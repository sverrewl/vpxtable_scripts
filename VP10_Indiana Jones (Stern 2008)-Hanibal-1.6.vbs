

Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Table used modified JP ball rolling routine
' Added InitVpmFFlipsSAM
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0



'******************* Options *********************

Const GIOnDuringAttractMode   = 0         '1 - GI on during attract, 0 - GI off during attract

dim DivValue:DivValue = 2 '**** Change Value to 4 if LED array does not display properly on your system

LoadVPM "01560000", "sam.VBS", 3.10



'********************
'Standard definitions
'********************

  Const cGameName = "ij4_210" 'change the romname here

     Const UseSolenoids = 1
     Const UseLamps = 0
     Const UseSync = 1
     Const HandleMech = 0

     'Standard Sounds
     Const SSolenoidOn = "Solenoid"
     Const SSolenoidOff = ""
     Const SCoin = "coin"



'************
' Table init.
'************
   'Variables
    Dim xx
    Dim Bump1,Bump2,Bump3,Mech3bank,bsTrough,bsVUK,visibleLock,bsTEject,bsMapVUK,bsRScoop
  Dim dtUDrop,dtLDropLower,dtLDropUpper,dtRDrop
  Dim PlungerIM
  Dim PMag
    Dim mLockMagnet
    Dim TBall1, TBALL2, PulserBall
    Dim Lift, Frei
    Dim DesktopMode:DesktopMode = Table.ShowDT

'
  Sub Table_Init
    tablelaunch.visible = 1 'Turn On Ball Info
    Controller.Switch(51) = 0 'ark position unknown
  Controller.Switch(51) = 0 'ark position unknown
  Sw39_Init 'ark enter opto
  Controller.Switch(40) = 1 'ark hit opto

  'Ark Init
  'TOD1.CreateBall
  'TOD2.CreateBall
  MAP1.CreateBall
  MAP2.CreateBall
  'MAP3.CreateBall
  'MAP4.CreateBall
  'TOD1.Kick 0,0
  'TOD2.Kick 0,0
  MAP1.Kick 0,0
  MAP2.Kick 0,0
  'MAP3.Kick 0,0
' MAP4.Kick 0,0
     Set PulserBall = MAP3.Createsizedball(25):PulserBall.image = "blank":MAP3.Kick 0,0





  'TOD1.Enabled = false
  'TOD2.Enabled = false
  MAP1.Enabled = false
  MAP2.Enabled = false
  MAP3.Enabled = false
  'MAP4.Enabled = false
  ArkLoading.Enabled = 1
'*****
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Indiana Jones (Stern 2008) by Hanibal"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 1
    .Hidden = 0
        .Games(cGameName).Settings.Value("sound") = 1
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With


    On Error Goto 0

' Set GICallback = GetRef("UpdateGI")

'**Trough
    Set bsTrough = New cvpm8BallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 17, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd "ballrelease", "Solenoid"
    bsTrough.Balls = 8

  Set PMag=New cvpmMagnet
  PMag.InitMagnet ArkMag,20
  PMag.GrabCenter=True

    ' Lock Magnet
    Set mLockMagnet = New cvpmMagnet
    With mLockMagnet
 '       .InitMagnet TempelMag4, 50
        .InitMagnet Trigger3, 300
        .GrabCenter = 1
 '       .Size = 300
 '       .CreateEvents "mLockMagnet"
     End With

  mLockMagnet.addball PulserBall

     Set TBall1 = MAP4.Createsizedball(27):TBall1.image = "blank":'MAP4.Kick 0,0':mLockMagnet.addball TBall1

     Set TBall2 = MAP5.Createsizedball(27):TBall2.image = "blank":'MAP4.Kick 0,0':mLockMagnet.addball TBALL2

     Frei = 1






' Set bsSVUK=New cvpmBallStack
' bsSVUK.InitSw 0,3,0,0,0,0,0,0
' bsSVUK.InitKick TopLaneKicker,0,20
' bsSVUK.InitExitSnd SoundFX("scoopexit"), SoundFX("rail")

  Set bsVUK=New cvpmBallStack
  bsVUK.InitSw 0,11,0,0,0,0,0,0
  bsVUK.InitKick sw11,165,20
  bsVUK.InitExitSnd "Ball-im-Target", "rail"


  Set bsMapVUK=New cvpmBallStack
  bsMapVUK.InitSw 0,45,0,0,0,0,0,0
  bsMapVUK.InitKick sw45,250,10
  bsMapVUK.InitExitSnd "Ball-im-Target", "rail"

'    Set bsRScoop=New cvpmBallStack
' bsRScoop.InitSw 0,49,0,0,0,0,0,0
' bsRScoop.InitKick sw49,270,42
' bsRScoop.KickZ=3.1415926/2
' bsRScoop.InitExitSnd SoundFX("scoopexit"), SoundFX("rail")

'**Nudging
      vpmNudge.TiltSwitch=-7
    vpmNudge.Sensitivity=1
    vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,Bumper4b,LeftSlingshot,RightSlingshot)


      '**Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  If GIOnDuringAttractMode = 1 Then GI_AllOn
  InitVpmFFlipsSAM
  End Sub

'MAP Newballid assignment
Sub trigger1_hit: NewBallID :trigger1.enabled = 0:End sub
Sub trigger2_hit: NewBallID :trigger2.enabled = 0:End sub
Sub trigger3_hit: NewBallID :trigger3.enabled = 0:End sub

Dim ArkLoadingCnt
Sub ArkLoading_Timer
  if controller.switch(17) = False then
    ArkLoadingCnt = ArkLoadingCnt + 1
  elseif controller.switch(17) = True then
    ArkLoadingCnt = 0
  end if

  if ArkLoadingCnt = 4 Then
    'MsgBox "Play Ball!"
    'tablelaunch.visible = False
    ArkLoading.Enabled = 0
  End if
End Sub

'*****Keys
Sub Table_KeyDown(ByVal Keycode)
' If keycode = 30 then sw60p1.Transz = -48
' If keycode = 31 then sw60p1.Transz = 0

  If keycode = 3 Then
    drainwall.isdropped = True
  End If

  If Keycode = LeftFlipperKey then
    'SolLFlipper true
    'SolULFlipper true
  End If
  If Keycode = RightFlipperKey then
    'SolRFlipper true
    'SolURFlipper true
  End If
'    If keycode = PlungerKey Then Plunger.Pullback:PNewPos = 0:PTime.Enabled = 1
    If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAtVol "PlungerPull", plunger, 1
    If keycode = LeftTiltKey Then LeftNudge 80, 1, 20:nudgebobble(keycode):End If
   If keycode = RightTiltKey Then RightNudge 280, 1, 20:nudgebobble(keycode):End If
   If keycode = CenterTiltKey Then CenterNudge 0, 1, 25 End If
   If vpmKeyDown(keycode) Then Exit Sub
End Sub


Sub table_KeyUp(ByVal Keycode)
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "Plunger", plunger, 1
    If vpmKeyUp(keycode) Then Exit Sub
End Sub



Sub Table_KeyUp(ByVal keycode)
  If vpmKeyUp(keycode) Then Exit Sub
  If Keycode = LeftFlipperKey then
    SolLFlipper false
    'SolULFlipper false
  End If
  If Keycode = RightFlipperKey then
    SolRFlipper False
    'SolURFlipper False
  End If
  If Keycode = StartGameKey Then Controller.Switch(16) = 0: tablelaunch.visible = 0
    If keycode = PlungerKey Then
    PTime.Enabled = 0:PTime2.Enabled = 1:Plunger.Fire
        If(BallinPlunger = 1) then 'the ball is in the plunger lane
            PlaySoundAtVol "Plunger2", plunger, 1
        else
            PlaySoundAtVol "Plunger", plunger, 1
        end if
  End If
End Sub


   'Solenoids

SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "PMag.MagnetOn="
SolCallback(4) = "bSVUK.SolOut"
SolCallback(5) = "RampGateSol"
SolCallback(6) = "TempleMotor"
SolCallback(7) = "solArkDiverter"

SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"

SolCallback(19) = "Flasher_Temple"
SolCallback(20) = "SetLamp 120,"
SolCallBack(21) = "SlingshotFlash"
SolCallback(22) = "Sankara_Flash"
SolCallback(23) = "SetLamp 123,"

SolCallback(25) = "RampFlash"
SolCallback(26) = "bSMapVUK.SolOut"  'MAP EJECT
SolCallback(27) = "solArkMotor"
SolCallback(28) = "FlashBackPanel" 'back panel 3
SolCallback(29) = "ArkFlash" 'ark front flash
SolCallback(30) = "SwordsmanMotor"
SolCallback(31) = "BumperFlash"
SolCallback(32) = "FlashCrusade"
SolCallback(36) = "ArkTopFlash"
SolCallback(37) = "SwordsmanFlash"
SolCallback(38) = "SchaedelFlash"


' Physics Module by Hanibal

Sub Physik_Timer
'Table.playfield.friction = (0.0020+(0.001*Rnd))


'If frei = 1 Then


     TBall1.x = templeball1.x  + 39 -(Lift/6)
     TBall1.y = templeball1.y  + (130*dCos(Lift))

     TBALL2.x = templeball2.x  + 24 -(Lift/6)
     TBALL2.y = templeball2.y  + (75*dCos(Lift))

     TBall1.velx=0
     TBall1.vely=0
     TBall1.velz=0
     TBALL2.velx=0
     TBALL2.vely=0
     TBALL2.velz=0
'End if

     TBall1.z = templeball1.z  + 40 +(135*dSin(Lift)) 'templeball1
     TBALL2.z = templeball1.z  + 40 +(79*dSin(Lift))'templeball2

If frei = 0 Then
frei = 1
End If





' Pulser Ball limitieren
'Dim MaxImpulse

'Maximpulse = 20

'If PulserBall.velx > Maximpulse Then
'PulserBall.velx=Maximpulse
'End If

'If PulserBall.velx < -Maximpulse Then
'PulserBall.velx= -Maximpulse
'End If

'If PulserBall.vely > Maximpulse Then
'PulserBall.vely=Maximpulse
'End If

'If PulserBall.vely < -Maximpulse Then
'PulserBall.vely=-Maximpulse
'End If


End Sub

'** Extra math to make my life easier **
Function dCos(degrees)
  Dim Pi:Pi = CSng(4*Atn(1))
  dcos = cos(degrees * Pi/180)
  if ABS(dCos) < 0.000001 Then dCos = 0
  if ABS(dCos) > 0.999999 Then dCos = 1 * sgn(dCos)
End Function

Function dSin(degrees)
  Dim Pi:Pi = CSng(4*Atn(1))
  dsin = sin(degrees * Pi/180)
  if ABS(dSin) < 0.000001 Then dSin = 0
  if ABS(dSin) > 0.999999 Then dSin = 1 * sgn(dSin)
End Function


 Sub Augen_Timer

' ******Hanibals Random Lights Script

Flasher25a.Intensity = (70+(40*Rnd))
Flasher25a1.Intensity = (200+(60*Rnd))
Flasher25b.Intensity = (70+(40*Rnd))
Flasher25b1.Intensity = (200+(60*Rnd))
Flasher29.Intensity = (50+(30*Rnd))
Flasher38.Intensity = (90+(20*Rnd))
Flasher38a.Intensity = (15+(2*Rnd))
Flasher38b.Intensity = (15+(2*Rnd))
Flasher21a.Intensity = (55+(10*Rnd))
Flasher21b.Intensity = (55+(10*Rnd))
Flasher21c.Intensity = (55+(10*Rnd))
Flasher21d.Intensity = (55+(10*Rnd))
Flasher32.Intensity = (30+(10*Rnd))
Flasher19.Intensity = (20+(7*Rnd))
Flasher22.Intensity = (20+(7*Rnd))
Flasher36a.Intensity = (90+(20*Rnd))
Flasher36b.Intensity = (90+(20*Rnd))
Flasher37.Intensity = (60+(10*Rnd))
Flasher31.Intensity = (55+(10*Rnd))
Flasher28.Intensity = (100+(10*Rnd))
Flasher28b.Intensity = (100+(10*Rnd))
Flasher28c.Intensity = (100+(10*Rnd))
l15.Intensity = (15+(5*Rnd))
FlasherAkator.Intensity = (70+(40*Rnd))

 End Sub





RampGateWall.IsDropped = 1

Sub RampGateSol(Enabled)
  If Enabled then
    'RampGate.collidable = 0
    RampGateWall.IsDropped = 0
    FlasherAkator.State = 1
  Else
    'RampGate.collidable = 1
    RampGateWall.IsDropped = 1
    FlasherAkator.State = 0
  End If
End Sub

Sub FlashBackPanel(Enabled)
  If Enabled Then
        If DesktopMode = False Then
        Flasher28.state = 1
        End if
        If DesktopMode = True Then
        Flasher28b.state = 1
        Flasher28c.state = 1
        End if
  Else
        Flasher28.state = 0
        Flasher28b.state = 0
        Flasher28c.state = 0
  End If
End Sub

Sub SwordsmanFlash(Enabled)
  If Enabled Then
        Flasher37.state = 1
  Else
        Flasher37.state = 0
  End If
End Sub

Sub SlingshotFlash(Enabled)
  If Enabled Then
        Flasher21a.state = 1
        Flasher21a1.state = 1
        Flasher21b.state = 1
        Flasher21b1.state = 1
  Else
        Flasher21a.state = 0
        Flasher21a1.state = 0
        Flasher21b.state = 0
        Flasher21b1.state = 0
  End If
End Sub

Sub ArkTopFlash(Enabled)
  If Enabled Then
        Flasher36a.state = 1
        Flasher36b.state = 1

  Else
        Flasher36a.state = 0
        Flasher36b.state = 0
  End If
End Sub


Sub Sankara_Flash(Enabled)
  If Enabled Then
        Flasher22.state = 1
  Else
        Flasher22.state = 0
  End If
End Sub


Sub Flasher_Temple(Enabled)
  If Enabled Then
        Flasher19.state = 1
  Else
        Flasher19.state = 0
  End If
End Sub


Sub FlashCrusade(Enabled)
  If Enabled Then
        Flasher32.state = 1
  Else
        Flasher32.state = 0
  End If
End Sub

Sub RampFlash(Enabled)
  If Enabled Then
        Flasher25a.state = 1
        Flasher25a1.state = 1
        Flasher25b.state = 1
        Flasher25b1.state = 1
  Else
        Flasher25a.state = 0
        Flasher25a1.state = 0
        Flasher25b.state = 0
        Flasher25b1.state = 0
  End If
End Sub

Sub BumperFlash(Enabled)
  If Enabled Then
        Flasher31.state = 1
  Else
        Flasher31.state = 0
  End If
End Sub


Sub SchaedelFlash(Enabled)
  If Enabled Then
        Flasher38.state = 1
        Flasher38a.state = 1
        Flasher38b.state = 1

  Else
        Flasher38.state = 0
        Flasher38a.state = 0
        Flasher38b.state = 0

  End If
End Sub



Sub ArkFlash(Enabled)
  If Enabled Then
        Flasher29.state = 1

  Else
    Flasher29.state = 0

  End If
End Sub


Sub TempleMotor(Enabled)
  If Enabled Then
    DT.Enabled = 1
'        PlaySound "E_Motorlift2", 0, 40 / (15*Rnd), -0.05, 0.15
    'Debug.print Timer & "DT Enabled"

  Else
    DT.Enabled = 0
 '       StopSound "E_Motorlift2"
    'Debug.print Timer & "DT Disabled"

  End If
End Sub

Sub SwordsmanMotor(Enabled)
  If Enabled Then
    SM.Enabled = 1
        PlaySound "E_Motorlift3", 0, 10 / (25*Rnd), -0.15, 0.15
    'Debug.print Timer & "SM Enabled"
  Else
    SM.Enabled = 0
        StopSound "E_Motorlift3"
    'Debug.print Timer & "SM Disabled"
  End If
End Sub

Sub solTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  End If
 End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    Plunger.Pullback
      vpmTimer.AddTimer 200, PlungerIM.AutoFire : Plunger.Fire
  End If
End Sub

Sub solArkDiverter(Enabled)
  if enabled Then
    diverter.rotatetoend
    'arkdiv.isdropped = 0
  else
    diverter.rotatetostart
    'arkdiv.isdropped = 1
  end if
End Sub

Sub solArkMotor(Enabled)
  if enabled then
    arkmotor.enabled = 1
        playsound "E_Motorlift", 0, 40 / (25*Rnd), 0.05, 0.15
    'debug.print "ArkMotor Enabled"
  else
    arkmotor.enabled = 0
    'debug.print "ArkMotor Disabled"
        StopSound "E_Motorlift"
  end if
End Sub

'primitive flippers!
dim MotorCallback
Set MotorCallback = GetRef("GameTimer")
Sub GameTimer
    UpdateFlipperLogos
End Sub

Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
    sw6p.RotX = -(sw6.currentangle) +90
    sw14p.RotX = -(sw14.currentangle) +90
    Primitive134.RotX = -(spinner1.currentangle)
    Primitive157.RotX = -(spinner2.currentangle)
End Sub




Sub SolLFlipper(Enabled)
     If Enabled Then
         LeftFlipper.RotateToEnd: PlaySoundAtVol "Flipper-oben-Links", LeftFlipper,VolFlip
     Else
         LeftFlipper.RotateToStart: PlaySoundAtVol "Flipper-unten-Links", LeftFlipper,VolFlip
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         RightFlipper.RotateToEnd : PlaySoundAtVol "Flipper-oben-Rechts", RightFlipper, VolFlip
     Else
         RightFlipper.RotateToStart : PlaySoundAtVol "Flipper-unten-Rechts", RightFlipper, VolFlip
     End If
End Sub





 'Drains and Kickers
Sub Drain_Hit
  PlaySoundAtVol "Drain", drain, 1
' ClearBallID
  bsTrough.AddBall Me
  Drain.TimerInterval = 200
  Drain.TimerEnabled = 1
End Sub

'Dim Ballcount:Ballcount = 8
Sub Drain_Timer
  GI_TroughCheck
  If GI_TroughCheck = 8 Then
    GI_AllOff 0
  End If
  Drain.TimerEnabled = False
End Sub

Sub BallRelease_UnHit()
  If arkballcnt = 4 then
    GI_AllOn
  Else
  End If
  NewBallId
End Sub

'Sub BallCountTimer_Timer()
' If Ballcount <= 3 then
'   GI_AllOn
' Else
'   GI_AllOff 0
' End If
'End Sub



'''*****************************************************************************************
'''*freneticamnesic level nudge script, based on rascals nudge bobble with help from gtxjoe*
'''*     add timers and "Nudgebobble(keycode)" to left and right tilt keys to activate     *
'''*****************************************************************************************
Dim bgcharctr:bgcharctr = 2
Dim centerlocation:centerlocation = 90
Dim bgdegree:bgdegree = 7 'move +/- 7 degrees
Dim bgdurationctr:bgdurationctr = 0

Sub LevelT_Timer()
  Dim loopctr
  Level.RotAndTra7 = Level.RotAndTra7 + bgcharctr  'change rotation value by bgcharctr
  'debug.print "Degrees: " & Level.RotAndTra7 & " Max degree offset: " & bgdegree & " Cycle count: " & bgdurationctr ''debug print
  If Level.RotAndTra7 >= bgdegree + centerlocation then bgcharctr = -1:bgdurationctr = bgdurationctr + 1   'if level moves past max degrees, change direction and increate durationctr
  If Level.RotAndTra7 <= -bgdegree + centerlocation then bgcharctr = 1  'if level moves past min location, change direction
  If bgdurationctr = 4 then bgdegree = bgdegree - 2:bgdurationctr = 0 'if level has moved back and forth 5 times, decrease amount of movement by -2 and repeat by resetting durationctr
  If bgdegree <= 0 then LevelT.Enabled = False:bgdegree = 7 'if amount of movement is 0, turn off LevelT timer and reset movement back to max 7 degrees
End Sub


Sub Nudgebobble(keycode)
  If keycode = LeftTiltKey then bgcharctr = -1  'if nudge left, move in - direction
  If keycode = RightTiltKey then bgcharctr = 1  'if nudge left, move in + direction
  If keycode = CenterTiltKey then     'if nudge center, generate random number 1 or 2.  If 1 change it to -2.  use this number for initial direction
    Dim randombobble:randombobble = Int(2 * Rnd + 1)
    If randombobble = 1 then randombobble = -2
    bgcharctr = randombobble
  End If
  LevelT.Enabled = True:bgdurationctr = 0:bgdegree = 7
End Sub

Sub bobblesome_Timer()  'This looks like a free running timer that 1 out of ten times will start movement
  Dim chance
  chance = Int(10*Rnd+1)
  If chance = 5 then Nudgebobble(CenterTiltKey)
End Sub




'Sub LaneKicker_Hit:
' bsSVUK.AddBall Me:
'End Sub
'
Sub sw11d_Hit:
  'GI_AllOff 1000
' ClearBallID
  bsVUK.AddBall Me
  PlaySound "scoopenter", 0, (8 +Vol(ActiveBall)),  Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
End Sub
Sub sw11_UnHit: NewBallId: End Sub


Sub TOD_Hit: bsVUK.AddBall Me:vpmTimer.PulseSw 12:  PlaySound "scoopenter", 0, (8 +Vol(ActiveBall)),  Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall):End Sub 'ClearBallID

Sub sw45_Hit:
  'GI_AllOff 1000
' ClearBallID
  bsMapVUK.AddBall Me:
    PlaySound "kicker_enter_center", 0, (8 +Vol(ActiveBall)),  Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
End Sub
Sub sw45_UnHit: NewBallId: End Sub



Sub TempleWall2_Hit  : vpmTimer.PulseSw 59:Me.TimerEnabled = 1:templeball2.TransZ = 15:sw59p.TransZ = 4: playsound "fx_chapa", 0, (8 +Vol(ActiveBall)),  Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall): Frei = 0: End Sub
Sub TempleWall2_Timer:Me.TimerEnabled = 0:templeball2.TransZ = 0:sw59p.TransZ = 0:End Sub

' Hanibals Pulser Script for cinectic impuls transfer
Sub Pulser_Hit :Pulsercinetic :End Sub 'mLockMagnet.Magneton = 1
'Sub Pulser_UnHit :mLockMagnet.Magneton = 0:End Sub

Sub Pulsercinetic

If Not isNull(PulserBall.velx) and Not isNull(ActiveBall.velx) Then
PulserBall.velx = (ActiveBall.velx * 0.95)
PulserBall.vely = (ActiveBall.vely * 0.95)
End If

End Sub
'-------------------------------------------------------------



Sub ArkMag_Hit:PMag.AddBall ActiveBall:End Sub
Sub ArkMag_UnHit:PMag.RemoveBall ActiveBall:End Sub

Sub TestMag_Hit:mLockMagnet.Magneton = 1:End Sub
Sub TestMag_UnHit:mLockMagnet.Magneton = 0:End Sub


Sub sw54_Hit  : vpmTimer.PulseSw(54): PlaysoundAtVol "rollover" ,ActiveBall, 1:End Sub

Sub sw38_Hit  : vpmTimer.PulseSw(38): PlaysoundAtVol "rollover", ActiveBall, 1 :End Sub

Sub sw58_Hit  : vpmTimer.PulseSw(58): PlaysoundAtVol "rollover", ActiveBall, 1 :End Sub

Sub sw13_Hit  : vpmTimer.PulseSw(13): PlaysoundAtVol "rollover", ActiveBall, 1 :End Sub


Sub sw24_Hit  : vpmTimer.PulseSw(24): PlaysoundAtVol "rollover", ActiveBall, 1 :sw24p.TransZ = -18: End Sub ' left outlane
Sub sw24_UnHit: sw24p.TransZ = 0: End Sub
Sub sw25_Hit  : vpmTimer.PulseSw(25): PlaysoundAtVol "rollover", ActiveBall, 1 :sw25p.TransZ = -18: End Sub ' left inlane
Sub sw25_UnHit: sw25p.TransZ = 0: End Sub

Sub sw26_Hit  : vpmTimer.PulseSw(26): Flasher21c.state = 1 : End Sub
Sub sw26_UnHit  : Flasher21c.state = 0 : End Sub
Sub sw27_Hit  : vpmTimer.PulseSw(27): Flasher21d.state = 1 : End Sub
Sub sw27_UnHit  : Flasher21d.state = 0 : End Sub

Sub sw28_Hit  : vpmTimer.PulseSw(28): PlaysoundAtVol "rollover", ActiveBall, 1 :sw28p.TransZ = -18: End Sub ' right inlane
Sub sw28_UnHit: sw28p.TransZ = 0: End Sub
Sub sw29_Hit  : vpmTimer.PulseSw(29): PlaysoundAtVol "rollover", ActiveBall, 1 :sw29p.TransZ = -18: End Sub ' right outlane
Sub sw29_UnHit: sw29p.TransZ = 0: End Sub

Sub sw61_Hit  : vpmTimer.PulseSw(61): End Sub

Sub sw62_Hit  : vpmTimer.PulseSw(62): End Sub


Sub sw48_Hit  : vpmTimer.PulseSw(48): End Sub

' Sub sw60_Hit  : Controller.Switch(60) = 1: End Sub
' Sub sw60_UnHit: Controller.Switch(60) = 0: End Sub
Sub TriggerSW60_Hit: vpmTimer.PulseSw 60: FlasherAkator2.State =1  End Sub
Sub TriggerSW60_UnHit: FlasherAkator2.State =0  End Sub

Sub sw40_Hit  : Controller.Switch(40) = 1:  PlaySoundAtVol "metalhit_medium",sw40,1: End Sub 'ark hit opto
Sub sw40_UnHit: Controller.Switch(40) = 0: End Sub

Sub sw35_Hit  : vpmTimer.PulseSw 35:Me.TimerEnabled = 1:sw35p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw35_Timer:Me.TimerEnabled = 0:sw35p.TransX = 0:End Sub
Sub sw36_Hit  : vpmTimer.PulseSw 36:Me.TimerEnabled = 1:sw36p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw36_Timer:Me.TimerEnabled = 0:sw36p.TransX = 0:End Sub
Sub sw37_Hit  : vpmTimer.PulseSw 37:Me.TimerEnabled = 1:sw37p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw37_Timer:Me.TimerEnabled = 0:sw37p.TransX = 0:End Sub
Sub sw41_Hit  : vpmTimer.PulseSw 41:Me.TimerEnabled = 1:sw41p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw41_Timer:Me.TimerEnabled = 0:sw41p.TransX = 0:End Sub
Sub sw42_Hit  : vpmTimer.PulseSw 42:Me.TimerEnabled = 1:sw42p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw42_Timer:Me.TimerEnabled = 0:sw42p.TransX = 0:End Sub
Sub sw43_Hit  : vpmTimer.PulseSw 43:Me.TimerEnabled = 1:sw43p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw43_Timer:Me.TimerEnabled = 0:sw43p.TransX = 0:End Sub
Sub sw44_Hit  : vpmTimer.PulseSw 44:Me.TimerEnabled = 1:sw44p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw44_Timer:Me.TimerEnabled = 0:sw44p.TransX = 0:End Sub
Sub sw55_Hit  : vpmTimer.PulseSw 55:Me.TimerEnabled = 1:sw55p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw55_Timer:Me.TimerEnabled = 0:sw55p.TransX = 0:End Sub
Sub sw56_Hit  : vpmTimer.PulseSw 56:Me.TimerEnabled = 1:sw56p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw56_Timer:Me.TimerEnabled = 0:sw56p.TransX = 0:End Sub
Sub sw57_Hit  : vpmTimer.PulseSw 57:Me.TimerEnabled = 1:sw57p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw57_Timer:Me.TimerEnabled = 0:sw57p.TransX = 0:End Sub
Sub sw1_Hit  : vpmTimer.PulseSw 1:Me.TimerEnabled = 1:sw1p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw1_Timer:Me.TimerEnabled = 0:sw1p.TransX = 0:End Sub
Sub sw2_Hit  : vpmTimer.PulseSw 2:Me.TimerEnabled = 1:sw2p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw2_Timer:Me.TimerEnabled = 0:sw2p.TransX = 0:End Sub
Sub sw3_Hit  : vpmTimer.PulseSw 3:Me.TimerEnabled = 1:sw3p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw3_Timer:Me.TimerEnabled = 0:sw3p.TransX = 0:End Sub
Sub sw4_Hit  : vpmTimer.PulseSw 4:Me.TimerEnabled = 1:sw4p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw4_Timer:Me.TimerEnabled = 0:sw4p.TransX = 0:End Sub
Sub sw5_Hit  : vpmTimer.PulseSw 5:Me.TimerEnabled = 1:sw5p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw5_Timer:Me.TimerEnabled = 0:sw5p.TransX = 0:End Sub
Sub sw7_Hit  : vpmTimer.PulseSw 7:Me.TimerEnabled = 1:sw7p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw7_Timer:Me.TimerEnabled = 0:sw7p.TransX = 0:End Sub
Sub sw8_Hit  : vpmTimer.PulseSw 8:Me.TimerEnabled = 1:sw8p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw8_Timer:Me.TimerEnabled = 0:sw8p.TransX = 0:End Sub
Sub sw9_Hit  : vpmTimer.PulseSw 9:Me.TimerEnabled = 1:sw9p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw9_Timer:Me.TimerEnabled = 0:sw9p.TransX = 0:End Sub
Sub sw10_Hit  : vpmTimer.PulseSw 10:Me.TimerEnabled = 1:sw10p.TransX = -4: playsoundAtVol "target", ActiveBall, VolTarg : End Sub
Sub sw10_Timer:Me.TimerEnabled = 0:sw10p.TransX = 0:End Sub

Sub sw6_Spin:vpmTimer.PulseSw 6: playsoundAtVol "spinner" , sw6,VolSpin: End Sub
Sub sw14_Spin:vpmTimer.PulseSw 14:playsoundAtVol "spinner" , sw14,VolSpin:End Sub




Sub sw48s_Spin:Me.TimerEnabled = 1:playsoundAtVol "spinner" , sw48s,VolSpin:End Sub
'Sub sw48s_Spin:Me.TimerEnabled = 1:SkullFlash.Alpha = 255:vpmTimer.PulseSw 48:End Sub
Sub sw48s_Timer:Me.TimerEnabled = 0:End Sub
Sub Spinner1_Hit:Test.State = 1:End Sub
Sub Spinner1_Spin:Me.TimerEnabled = 1:playsoundAtVol "spinner" , Spinner1,VolSpin:End Sub
'Sub Spinner1_Spin:Me.TimerEnabled = 1:SkullFlash.Alpha = 255:vpmTimer.PulseSw 60:End Sub
Sub Spinner1_Timer:Me.TimerEnabled = 0:End Sub

Sub Spinner2_Hit: playsoundAtVol "spinner" , Spinner2,VolSpin: End Sub
Sub Spinner2_Spin: playsoundAtVol "spinner" , Spinner2,VolSpin: End Sub




'***Slings and rubbers
 ' Slings
 'Dim LStep, RStep

 Sub LeftSlingShot_Slingshot
  'For each xx in LHammerA:xx.IsDropped = 0:Next
  Leftsling1 = True
  Leftsling2 = True
  Leftsling3 = True
  Controller.Switch(26) = 1
  PlaySoundAtVol "slingshot",ActiveBall, 1:LeftSlingshot.TimerEnabled = 1
  End Sub

Dim Leftsling1:Leftsling1 = False

Sub LS1_Timer()
  If Leftsling1 = True and Left1.ObjRotZ < -7 then Left1.ObjRotZ = Left1.ObjRotZ + 2
  If Leftsling1 = False and Left1.ObjRotZ > -20 then Left1.ObjRotZ = Left1.ObjRotZ - 2
  If Left1.ObjRotZ >= -7 then Leftsling1 = False
End Sub

Dim Leftsling2:Leftsling2 = False

Sub LS2_Timer()
  If Leftsling2 = True and Left2.ObjRotZ > -212.5 then Left2.ObjRotZ = Left2.ObjRotZ - 2
  If Leftsling2 = False and Left2.ObjRotZ < -199.5 then Left2.ObjRotZ = Left2.ObjRotZ + 2
  If Left2.ObjRotZ <= -212.5 then Leftsling2 = False
End Sub

Dim Leftsling3:Leftsling3 = False

Sub LS3_Timer()
  If Leftsling3 = True and Left3.TransZ > -23 then Left3.TransZ = Left3.TransZ - 4
  If Leftsling3 = False and Left3.TransZ < -0 then Left3.TransZ = Left3.TransZ + 4
  If Left3.TransZ <= -23 then Leftsling3 = False
End Sub

 Dim Rightsling1:Rightsling1 = False

Sub RS1_Timer()
  If Rightsling1 = True and Right1.ObjRotZ > 7 then Right1.ObjRotZ = Right1.ObjRotZ - 2
  If Rightsling1 = False and Right1.ObjRotZ < 20 then Right1.ObjRotZ = Right1.ObjRotZ + 2
  If Right1.ObjRotZ <= 7 then Rightsling1 = False
End Sub

Dim Rightsling2:Rightsling2 = False

Sub RS2_Timer()
  If Rightsling2 = True and Right2.ObjRotZ < 212.5 then Right2.ObjRotZ = Right2.ObjRotZ + 2
  If Rightsling2 = False and Right2.ObjRotZ > 199.5 then Right2.ObjRotZ = Right2.ObjRotZ - 2
  If Right2.ObjRotZ >= 212.5 then Rightsling2 = False
End Sub

Dim Rightsling3:Rightsling3 = False

Sub RS3_Timer()
  If Rightsling3 = True and Right3.TransZ > -23 then Right3.TransZ = Right3.TransZ - 4
  If Rightsling3 = False and Right3.TransZ < -0 then Right3.TransZ = Right3.TransZ + 4
  If Right3.TransZ <= -23 then Rightsling3 = False
End Sub

 Sub LeftSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(26) = 0:End Sub

 Sub RightSlingShot_Slingshot
  Rightsling1 = True
  Rightsling2 = True
  Rightsling3 = True
  Controller.Switch(27) = 1
  PlaySoundAtVol "slingshot" ,ActiveBall, 1:RightSlingshot.TimerEnabled = 1
  End Sub

 Sub RightSlingShot_Timer:Me.TimerEnabled = 0:Controller.Switch(27) = 0:End Sub

 Sub sw46_Slingshot
  Controller.Switch(46) = 1
  PlaySoundAtVol "slingshot", ActiveBall, 1
  sw46.TimerEnabled = 1
  End Sub

Sub sw46_Timer:Me.TimerEnabled = 0:Controller.Switch(46) = 0:End Sub

 Sub sw47_Slingshot
  Controller.Switch(47) = 1
  PlaySoundAtVol "slingshot", ActiveBall, 1
  sw47.TimerEnabled = 1
  End Sub

Sub sw47_Timer:Me.TimerEnabled = 0:Controller.Switch(47) = 0:End Sub

     ' Impulse Plunger
    Const IMPowerSetting = 50
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
    .Switch 23
        .Random 1.5
        .InitExitSnd "plunger2", "plunger"
        .CreateEvents "plungerIM"
    End With

Sub sw23_unhit :GI_TroughCheck: End Sub
'
'Hanibals Bumper Script with Light
      Sub Bumper1b_Hit:vpmTimer.PulseSw 30: BumperL1.State = 1 :PlaySoundAtVol "bumper", Bumper1b,VolBump :B1Drop = True:End Sub

      Sub Bumper2b_Hit:vpmTimer.PulseSw 31: BumperL2.State = 1 :PlaySoundAtVol "bumper", Bumper2b,VolBump:B2Drop = True: End Sub

      Sub Bumper3b_Hit:vpmTimer.PulseSw 32: BumperL3.State = 1 :PlaySoundAtVol "bumper", Bumper3b,VolBump:B3Drop = True:End Sub

      Sub Bumper4b_Hit:vpmTimer.PulseSw 33: BumperL4.State = 1 :PlaySoundAtVol "bumper", Bumper4b,VolBump:B4Drop = True:End Sub

Dim B1Drop:B1Drop = False

Sub B1T_Timer()
  If B1Drop = True and Br1.z > -65 then Br1.z = Br1.z - 7
  If B1Drop = False and Br1.z < -25 then Br1.z = Br1.z + 7
  If Br1.z <= -65 then B1Drop = False : BumperL1.State = 0
End Sub

Dim B2Drop:B2Drop = False

Sub B2T_Timer()
  If B2Drop = True and Br2.z > -65 then Br2.z = Br2.z - 7
  If B2Drop = False and Br2.z < -25 then Br2.z = Br2.z + 7
  If Br2.z <= -65 then B2Drop = False : BumperL2.State = 0
End Sub

Dim B3Drop:B3Drop = False

Sub B3T_Timer()
  If B3Drop = True and Br3.z > -65 then Br3.z = Br3.z - 7
  If B3Drop = False and Br3.z < -25 then Br3.z = Br3.z + 7
  If Br3.z <= -65 then B3Drop = False : BumperL3.State = 0
End Sub

Dim B4Drop:B4Drop = False

Sub B4T_Timer()
  If B4Drop = True and Br4.z > -65 then Br4.z = Br4.z - 7
  If B4Drop = False and Br4.z < -25 then Br4.z = Br4.z + 7
  If Br4.z <= -65 then B4Drop = False : BumperL4.State = 0
End Sub


 '****************************************
 ' SetLamp 0 is Off
 ' SetLamp 1 is On
 ' LampState(x) current state
 '****************************************

Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim x


AllLampsOff()
LampTimer.Interval = 40 'lamp fading speed
LampTimer.Enabled = 1
'
FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1
'
'' Lamp & Flasher Timers
'
Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
      FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
End Sub
Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub



'
 Sub UpdateLamps()

  nFadeL 7, l7, "indypflightson", "indypf" ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 8, l8, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 9, l9, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 10, l10, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 11, l11, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 12, l12, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 13, l13, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 14, l14, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 15, l15, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 16, l16, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 19, l19, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 20, l20, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 21, l21, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 22, l22, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 23, l23, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 24, l24, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 25, l25, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 26, l26, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 27, l27, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 28, l28, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 29, l29, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 30, l30, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 31, l31, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 32, l32, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 33, l33, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 34, l34, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 35, l35, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 36, l36, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 37, l37, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 38, l38, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 39, l39, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 41, l41, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 42, l42, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 43, l43, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 45, l45, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 46, l46, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 47, l47, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 48, l48, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 49, l49, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 50, l50, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 51, l51, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 52, l52, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 53, l53, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 54, l54, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 55, l55, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 56, l56, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 64, l64, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 65, l65, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 66, l66, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 67, l67, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 68, l68, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 69, l69, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 70, l70, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 71, l71, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 72, l72, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
    NFadeLD 73, l73,l73b1, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 77, l77, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeL 79, l79, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
  nFadeLD 80, l80,l80b, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"
' nFadeL 80, l80b, "indypflightson", "indypf"  ', "indypflightsb", "indypflightsa", "indypflightsoff"


End Sub








Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 1
        Case 4:pri.image = b:FadingLevel(nr) = 2
        Case 5:pri.image = a:FadingLevel(nr) = 3
    End Select
End Sub


''Lights

Sub FadeL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeL(nr, Light, a,b)   'NFadeL(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.image = b:FadingLevel(nr) = 0 : Light.state = 0 'Off
        Case 3:Light.image = b:FadingLevel(nr) = 2 : Light.state = 1'fading...
        Case 4:Light.image = b:FadingLevel(nr) = 3 : Light.state = 1
        Case 5:Light.image = b:FadingLevel(nr) = 1 : Light.state = 1
    End Select
End Sub

'Parallel Script
Sub NFadeLD(nr, Light,Light2, a,b)   'NFadeL(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.image = b:FadingLevel(nr) = 0 : Light.state = 0 : Light2.state = 0   'Off
        Case 3:Light.image = b:FadingLevel(nr) = 2 : Light.state = 1 : Light2.state = 1   'fading...
        Case 4:Light.image = b:FadingLevel(nr) = 3 : Light.state = 1 : Light2.state = 1
        Case 5:Light.image = b:FadingLevel(nr) = 1 : Light.state = 1 : Light2.state = 1
    End Select
End Sub







' Flasher objects
' Uses own faster timer

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
 '           Object.alpha = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 255 Then
                FlashLevel(nr) = 255
                FlashState(nr) = -2 'completely on
            End if
'            Object.alpha = FlashLevel(nr)
    End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub


Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
'    LampState(nr) = abs(value) + 4
'FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub


Sub FlashAR(nr, ramp, a, b, c, r)                                                          'used for reflections when there is no off ramp
    Select Case FadingState(nr)
        Case 2:'ramp.intensity = 0:r.State = ABS(r.state -1):FadingState(nr) = 0                'Off
        Case 3:ramp.image = c:r.State = ABS(r.state -1):FadingState(nr) = 2                'fading...
        Case 4:ramp.image = b:r.State = ABS(r.state -1):FadingState(nr) = 3                'fading...
        Case 5:ramp.image = a:'ramp.intensity = 1:r.State = ABS(r.state -1):FadingState(nr) = 1 'ON
    End Select
End Sub


Sub FlashARm(nr, ramp, a, b, c, r)
    Select Case FadingState(nr)
        Case 2:'ramp.intensity = 0:r.State = ABS(r.state -1)
        Case 3:ramp.image = c:r.State = ABS(r.state -1)
        Case 4:ramp.image = b:r.State = ABS(r.state -1)
        Case 5:ramp.image = a:'ramp.intensity = 1:r.State = ABS(r.state -1)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
 '           Object.alpha = FlashLevel(nr)
        Case 1         ' on
 '           Object.alpha = FlashLevel(nr)
    End Select
End Sub

Sub FlasherTimer_Timer()


 End Sub

'SOUNDS

Sub BallstopL_Hit: PlaySoundAtVol "drop_left", ActiveBall,1:End Sub
Sub arkdrop_Hit: PlaySoundAtVol "drop_right", ActiveBall,1:End Sub
'Sub BallstopR_Hit: PlaySound "drop_Right": End Sub
'Sub Gate3_Hit: PlaySound "Gate": End Sub



'***********************
'Hanibal Ramp Plunger
'***********************

 Dim BallinPlunger


Sub swPlunger_Hit:BallinPlunger = 1:End Sub                            'in this sub you may add a switch, for example Controller.Switch(14) = 1

Sub swPlunger_UnHit:BallinPlunger = 0:End Sub                          'in this sub you may add a switch, for example Controller.Switch(14) = 0

Sub Plungin_Hit:Plungerlight.state = 1 :End Sub

Sub Plungin_UnHit: Plungerlight.state = 0    :End Sub


   '*************************************
  '          Nudge System
  '*************************************

  Dim LeftNudgeEffect, RightNudgeEffect, NudgeEffect

  Sub LeftNudge(angle, strength, delay)
      vpmNudge.DoNudge angle, (strength * (delay-LeftNudgeEffect) / delay) + RightNudgeEffect / delay
      LeftNudgeEffect = delay
      RightNudgeEffect = 0
      RightNudgeTimer.Enabled = 0
      LeftNudgeTimer.Interval = delay
      LeftNudgeTimer.Enabled = 1
  End Sub

  Sub RightNudge(angle, strength, delay)
      vpmNudge.DoNudge angle, (strength * (delay-RightNudgeEffect) / delay) + LeftNudgeEffect / delay
      RightNudgeEffect = delay
      LeftNudgeEffect = 0
      LeftNudgeTimer.Enabled = 0
      RightNudgeTimer.Interval = delay
      RightNudgeTimer.Enabled = 1
  End Sub

  Sub CenterNudge(angle, strength, delay)
      vpmNudge.DoNudge angle, strength * (delay-NudgeEffect) / delay
      NudgeEffect = delay
      NudgeTimer.Interval = delay
      NudgeTimer.Enabled = 1
  End Sub

  Sub LeftNudgeTimer_Timer()
      LeftNudgeEffect = LeftNudgeEffect-1
      If LeftNudgeEffect = 0 then LeftNudgeTimer.Enabled = 0
  End Sub

  Sub RightNudgeTimer_Timer()
      RightNudgeEffect = RightNudgeEffect-1
      If RightNudgeEffect = 0 then RightNudgeTimer.Enabled = 0
  End Sub

  Sub NudgeTimer_Timer()
      NudgeEffect = NudgeEffect-1
      If NudgeEffect = 0 then NudgeTimer.Enabled = 0
  End Sub


 Dim tnopb, nosf
 '
 tnopb = 16   ' <<<<< SET to the "Total Number Of Possible Balls" in play at any one time
 nosf = 9 ' <<<<< SET to the "Number Of Sound Files" used / B2B collision volume levels

 Dim currentball(16), ballStatus(16)
 Dim iball, cnt, coff, errMessage

 XYdata.interval = 1      ' Timer interval starts at 1 for the highest ball data sample rate
 coff = False       ' Collision off set to false

 For cnt = 0 to ubound(ballStatus)  ' Initialize/clear all ball stats, 1 = active, 0 = non-existant
  ballStatus(cnt) = 0
 Next
'
 '======================================================
 ' <<<<<<<<<<<<<< Ball Identification >>>>>>>>>>>>>>
 '======================================================
 ' Call this sub from every kicker(or plunger) that creates a ball.
 Sub NewBallID            ' Assign new ball object and give it ID for tracking
  For cnt = 1 to ubound(ballStatus)   ' Loop through all possible ball IDs
      If ballStatus(cnt) = 0 Then     ' If ball ID is available...
      Set currentball(cnt) = ActiveBall     ' Set ball object with the first available ID
      currentball(cnt).uservalue = cnt      ' Assign the ball's uservalue to it's new ID
      ballStatus(cnt) = 1       ' Mark this ball status active
      ballStatus(0) = ballStatus(0)+1     ' Increment ballStatus(0), the number of active balls
  If coff = False Then        ' If collision off, overrides auto-turn on collision detection
              ' If more than one ball active, start collision detection process
  If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
  End If
  Exit For          ' New ball ID assigned, exit loop
      End If
      Next
 '    Debugger          ' For demo only, display stats
 End Sub

 ' Call this sub from every kicker that destroys a ball, before the ball is destroyed.
 'Sub ClearBallID
 '    On Error Resume Next        ' Error handling for debugging purposes
 '    iball = ActiveBall.uservalue      ' Get the ball ID to be cleared
 '    currentball(iball).UserValue = 0      ' Clear the ball ID
 '    If Err Then Msgbox Err.description & vbCrLf & iball
 '      ballStatus(iBall) = 0         ' Clear the ball status
 '    ballStatus(0) = ballStatus(0)-1     ' Subtract 1 ball from the # of balls in play
 '    On Error Goto 0
 'End Sub
'
 '=====================================================
 ' <<<<<<<<<<<<<<<<< XYdata_Timer >>>>>>>>>>>>>>>>>
 '=====================================================
 ' Ball data collection and B2B Collision detection.
 ReDim baX(tnopb,4), baY(tnopb,4), bVx(tnopb,4), bVy(tnopb,4), TotalVel(tnopb,4)
 Dim cForce, bDistance, xyTime, cFactor, id, id2, id3, B1, B2

 Sub XYdata_Timer()
  ' xyTime... Timers will not loop or start over 'til it's code is finished executing. To maximize
  ' performance, at the end of this timer, if the timer's interval is shorter than the individual
  ' computer can handle this timer's interval will increment by 1 millisecond.
     xyTime = Timer+(XYdata.interval*.001)  ' xyTime is the system timer plus the current interval time
  ' Ball Data... When a collision occurs a ball's velocity is often less than it's velocity before the
  ' collision, if not zero. So the ball data is sampled and saved for four timer cycles.
      If id2 >= 4 Then id2 = 0            ' Loop four times and start over

  '      tablelaunch.visible = 0                'Turn off Ball sorting Info
      id2 = id2+1               ' Increment the ball sampler ID

      For id = 1 to ubound(ballStatus)          ' Loop once for each possible ball
 '    If ballStatus(id) = 1 Then            ' If ball is active...
 '       If BallStatus(b)=0 Then Exit Sub
 '           tablelaunch.Visible = True                 'Turn on Ball sorting Info
 '      baX(id,id2) = round(currentball(id).x,2)        ' Sample x-coord
 '      baY(id,id2) = round(currentball(id).y,2)        ' Sample y-coord
 '      bVx(id,id2) = round(currentball(id).velx,2)       ' Sample x-velocity
 '      bVy(id,id2) = round(currentball(id).vely,2)       ' Sample y-velocity
 '      TotalVel(id,id2) = (bVx(id,id2)^2+bVy(id,id2)^2)    ' Calculate total velocity
 '    If TotalVel(id,id2) > TotalVel(0,0) Then TotalVel(0,0) = int(TotalVel(id,id2))
 '    End If

      Next
  ' Collision Detection Loop - check all possible ball combinations for a collision.
  ' bDistance automatically sets the distance between two colliding balls. Zero milimeters between
  ' balls would be perfect, but because of timing issues with ball velocity, fast-traveling balls
  ' prevent a low setting from always working, so bDistance becomes more of a sensitivity setting,
  ' which is automated with calculations using the balls' velocities.
  ' Ball x/y-coords plus the bDistance determines B2B proximity and triggers a collision.
  id3 = id2 : B2 = 2 : B1 = 1           ' Set up the counters for looping
  Do
  If ballStatus(B1) = 1 and ballStatus(B2) = 1 Then     ' If both balls are active...
    bDistance = int((TotalVel(B1,id3)+TotalVel(B2,id3))^1.04)
    If ((baX(B1,id3)-baX(B2,id3))^2+(baY(B1,id3)-baY(B2,id3))^2)<2800+bDistance Then collide B1,B2 : Exit Sub
    End If
    B1 = B1+1             ' Increment ball1

    If B1 >= ballStatus(0) Then Exit Do       ' Exit loop if all ball combinations checked
    If B1 >= B2 then B1 = 1:B2 = B2+1       ' If ball1 >= reset ball1 and increment ball2

  Loop

    If ballStatus(0) <= 1 Then XYdata.enabled = False       ' Turn off timer if one ball or less

  If XYdata.interval >= 40 Then coff = True : XYdata.enabled = False  ' Auto-shut off
  If Timer > xyTime * 3 Then coff = True : XYdata.enabled = False   ' Auto-shut off
      If Timer > xyTime Then XYdata.interval = XYdata.interval+1    ' Increment interval if needed
 End Sub

 '=========================================================
 ' <<<<<<<<<<< Collide(ball id1, ball id2) >>>>>>>>>>>
 '=========================================================
 'Calculate the collision force and play sound accordingly.
 Dim cTime, cb1,cb2, avgBallx, cAngle, bAngle1, bAngle2

 Sub Collide(cb1,cb2)
 ' The Collision Factor(cFactor) uses the maximum total ball velocity and automates the cForce calculation, maximizing the
 ' use of all sound files/volume levels. So all the available B2B sound levels are automatically used by adjusting to a
 ' player's style and the table's characteristics.
    If TotalVel(0,0)/1.8 > cFactor Then cFactor = int(TotalVel(0,0)/1.8)
 ' The following six lines limit repeated collisions if the balls are close together for any period of time
    avgBallx = (bvX(cb2,1)+bvX(cb2,2)+bvX(cb2,3)+bvX(cb2,4))/4
    If avgBallx < bvX(cb2,id2)+.1 and avgBallx > bvX(cb2,id2)-.1 Then
    If ABS(TotalVel(cb1,id2)-TotalVel(cb2,id2)) < .000005 Then Exit Sub
    End If
    If Timer < cTime Then Exit Sub
    cTime = Timer+.1        ' Limits collisions to .1 seconds apart
' GetAngle(x-value, y-value, the angle name) calculates any x/y-coords or x/y-velocities and returns named angle in radians
'   GetAngle baX(cb1,id3)-baX(cb2,id3), baY(cb1,id3)-baY(cb2,id3),cAngle  ' Collision angle via x/y-coordinates
  id3 = id3 - 1 : If id3 = 0 Then id3 = 4   ' Step back one xyData sampling for a good velocity reading
'   GetAngle bVx(cb1,id3), bVy(cb1,id3), bAngle1  ' ball 1 travel direction, via velocity
'   GetAngle bVx(cb2,id3), bVy(cb2,id3), bAngle2  ' ball 2 travel direction, via velocity
 ' The main cForce formula, calculating the strength of a collision
  cForce = Cint((abs(TotalVel(cb1,id3)*Cos(cAngle-bAngle1))+abs(TotalVel(cb2,id3)*Cos(cAngle-bAngle2))))
      If cForce < 4 Then Exit Sub     ' Another collision limiter
      cForce = Cint((cForce)/(cFactor/nosf))    ' Divides up cForce for the proper sound selection.
    If cForce > nosf-1 Then cForce = nosf-1   ' First sound file 0(zero) minus one from number of sound files
      PlaySound("collide" & cForce)     ' Combines "collide" with the calculated sound level and play sound
 End Sub

 '=================================================
 ' <<<<<<<< GetAngle(X, Y, Anglename) >>>>>>>>
 '=================================================
 ' A repeated function which takes any set of coordinates or velocities and calculates an angle in radians.
 Dim Xin,Yin,rAngle,Radit,wAngle,Pi
 Pi = Round(4*Atn(1),6)         '3.1415926535897932384626433832795

 Sub GetAngle(Xin, Yin, wAngle)
    If Sgn(Xin) = 0 Then
      If Sgn(Yin) = 1 Then rAngle = 3 * Pi/2 Else rAngle = Pi/2
      If Sgn(Yin) = 0 Then rAngle = 0
    Else
      rAngle = atn(-Yin/Xin)      ' Calculates angle in radians before quadrant data
    End If
    If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
    If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
    wAngle = round((Radit + rAngle),4)    ' Calculates angle in radians with quadrant data
  '"wAngle = round((180/Pi) * (Radit + rAngle),4)" ' Will convert radian measurements to degrees - to be used in future
 End Sub


Dim gistep

gistep = 255 / 8

Sub UpdateGI(no, step)
    Select Case no
        Case 0 'All

    Case 1 'PF Spotlights

    End Select
End Sub


Sub FlippersOn
  'LFLogo.image = "flipper-l2"
  'LFLogo1.image = "flipper-l2"
  'RFLogo.image = "flipper-r2"
  'RFLogo1.image = "flipper-r2"
End Sub

Sub FlippersOff
  'LFLogo.image = "flipper-l2off"
  'LFLogo1.image = "flipper-l2off"
  'RFLogo.image = "flipper-r2off"
  'RFLogo1.image = "flipper-r2off"
End Sub

Sub FlippersRedOn
  'LFLogo.image = "flipper-l2red"
  'LFLogo1.image = "flipper-l2red"
  'RFLogo.image = "flipper-r2red"
  'RFLogo1.image = "flipper-r2red"
End Sub

Dim Jaildown, RPDown, LPUp

Sub Table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub



Sub GI_AllOff (time) 'Turn GI Off
  'debug.print "GI OFF " & time
  UpdateGI 0,0
    Backlight1.state = 0
    Backlight2.state = 0
    Backlight3.state = 0
    Backlight4.state = 0
    Backlight5.state = 0
    Backlight6.state = 0
    Backlight7.state = 0
    Backlight8.state = 0
    Backlight9.state = 0
    Backlight10.state = 0
    Backlight11.state = 0
  GI_1.state = 0
  GI_2.state = 0
  GI_3.state = 0
  GI_4.state = 0
  GI_5.state = 0
  GI_6.state = 0
  GI_7.state = 0
  GI_8.state = 0
  GI_9.state = 0
  GI_10.state = 0
  GI_11.state = 0
  GI_12.state = 0
  GI_13.state = 0
  GI_14.state = 0
  GI_15.state = 0
  GI_16.state = 0
  GI_17.state = 0
  GI_18.state = 0
  GI_19.state = 0
  GI_20.state = 0
  GI_21.state = 0
  GI_22.state = 0
  GI_23.state = 0
  GI_24.state = 0
  GI_25.state = 0
  GI_26.state = 0
  GI_27.state = 0
  GI_28.state = 0
  GI_29.state = 0
  GI_30.state = 0
  GI_31.state = 0
  GI_33.state = 0
  GI_34.state = 0
  Groundlight.state = 0


  FlippersOff
  If time > 0 Then
    GI_AllOnT.Interval = time
    GI_AllOnT.Enabled = 0
    GI_AllOnT.Enabled = 1
  End If
End Sub

Sub GI_AllOn 'Turn GI On
  UpdateGI 0,8
    Backlight1.state = 1
    Backlight2.state = 1
    Backlight3.state = 1
    Backlight4.state = 1
    Backlight5.state = 1
    Backlight6.state = 1
    Backlight7.state = 1
    Backlight8.state = 1
    Backlight9.state = 1
    Backlight10.state = 1
    Backlight11.state = 1
  GI_1.state = 1
  GI_2.state = 1
  GI_3.state = 1
  GI_4.state = 1
  GI_5.state = 1
  GI_6.state = 1
  GI_7.state = 1
  GI_8.state = 1
  GI_9.state = 1
  GI_10.state = 1
  GI_11.state = 1
  GI_12.state = 1
  GI_13.state = 1
  GI_14.state = 1
  GI_15.state = 1
  GI_16.state = 1
  GI_17.state = 1
  GI_18.state = 1
  GI_19.state = 1
  GI_20.state = 1
  GI_21.state = 1
  GI_22.state = 1
  GI_23.state = 1
  GI_24.state = 1
  GI_25.state = 1
  GI_26.state = 1
  GI_27.state = 1
  GI_28.state = 1
  GI_29.state = 1
  GI_30.state = 1
  GI_31.state = 1
  GI_33.state = 1
  GI_34.state = 1
  Groundlight.state = 1










  FlippersOn
End Sub

Sub GI_AllOnT_Timer 'Turn GI On timer
  UpdateGI 0,8

  FlippersOn
  GI_AllOnT.Enabled = 0
End Sub

Dim MultiballFlag
Function GI_TroughCheck
  Dim Ballcount:  Ballcount = 0
  If Controller.Switch(18) = TRUE then Ballcount = Ballcount + 1':debug.print "18x"
    If Controller.Switch(19) = TRUE then Ballcount = Ballcount + 1':debug.print "19x"
    If Controller.Switch(20) = TRUE then Ballcount = Ballcount + 1':debug.print "20x"
    If Controller.Switch(21) = TRUE then Ballcount = Ballcount + 1':debug.print "21x"
  Ballcount = Ballcount + arkballcnt

  If Ballcount < 7 Then 'Keep track of multiball mode
    MultiballFlag = 1
  Else
    MultiballFlag = 0
  End If

  GI_TroughCheck = Ballcount

  debug.print "Troughcheck " & ballcount & " Multiball " & MultiballFlag

  If ballcount = 8 then
    GameOverTimerCheck.Enabled = 1 'no ball in play
    'debug.print timer & "Game Over?"
  Else
    GameOverTimerCheck.Enabled = 0 'ball in play
    'debug.print timer & "Game Not Over"
  End If

End Function

GameOverTimerCheck.Interval = 30000
Sub GameOverTimerCheck_Timer
  'debug.print timer & "Game Over!"
  If GIOnDuringAttractMode = 1 Then GI_AllOn
  GameOverTimerCheck.Enabled = 0
End Sub

sub solcheck(value,enabled) 'gtxjoe romtest script
  dim x
  dprint timer & " solenoid " & value &"="&enabled

  'Add required table solenoid actions here
  Select Case value
    Case 1:
      If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 22
        'debug.print "BALLRELEASE"
      End If
    Case 2: SolAutoFire enabled
    Case 7: 'ARK DIVERTER
      'debug.print timer & " diverter " & enabled
      if enabled Then
        'diverter.rotatetoend
        arkdiv.isdropped = 0
      else
        'diverter.rotatetostart
        arkdiv.isdropped = 1
      end if
'   Case  25: If enabled then bsTrough.EntrySol_On
'   Case  26: If enabled then bsTrough.ExitSol_On
    case 27: 'ARK MOTOR
      if enabled then
        controller.switch(50) = 0
        controller.switch(51) = 0
        arkmotor.interval = 2000
        arkmotor.enabled = 1
      else
        arkmotor.enabled = 0
      end if
  End Select
End Sub


Const ArkHeightMax = 208
Const ArkLidMax = 35
Const ArkSpeed = .15
arkmotor.interval = 2
dim Arkdirection: ArkDirection = 1
Dim Arkheight, ArkPos, ArkLidAngle
Dim ArkTempPos: ArkTempPos = 0

Sub ArkMotor_Timer

  'Change height
  ArkTempPos = ArkTempPos + ArkDirection*ArkSpeed
  ArkPos = ArkTempPos
  'debug.print arktemppos
  ArkLidAngle = ArkTempPos * ArkLidMax/ArkHeightMax

  'Set Switches
  If ArkTempPos >= ArkHeightMax Then  'UP
    Controller.Switch(50) = 1
    ArkPos = ArkHeightMax
  Else
    Controller.Switch(50) = 0
  End If
  If ArkTempPos <= 0 Then         'DOWN
    Controller.Switch(51) = 1
    ArkPos = 0
  Else
    Controller.Switch(51) = 0
  End If

  'Move Primitives and balls
  if Not isNull(arkballs(0)) then Arkballs(0).z = ArkPos + 25:
  if Not isNull(arkballs(1)) then Arkballs(1).z = ArkPos + 25:
  if Not isNull(arkballs(2)) then Arkballs(2).z = ArkPos + 25:
  if Not isNull(arkballs(3)) then Arkballs(3).z = ArkPos + 25:
  Primitive107.TransZ = ArkPos - 15 'Platform
  PrimArkLid.RotZ = ArkLidAngle

  'Release Balls if on top of ark
  if Not isNull(arkballs(0)) Then
    If Arkballs(0).z = ArkHeightMax + 25 then
      ArkPlatformTrigger0.Enabled = 1
      ArkKicker0.kick 0,0
      arkballs(0) = NULL
      arkballcnt = 0
    End If
  End If
  if Not isNull(arkballs(1)) Then
    If Arkballs(1).z = ArkHeightMax + 25 then
      ArkPlatformTrigger1.Enabled = 1
      ArkKicker1.kick 0,0
      arkballs(1) = NULL
    End If
  End If
  if Not isNull(arkballs(2)) Then
    If Arkballs(2).z = ArkHeightMax + 25 then
      ArkPlatformTrigger2.Enabled = 1
      ArkKicker2.kick 0,0
      arkballs(2) = NULL
    End If
  End If
  if Not isNull(arkballs(3)) Then
    If Arkballs(3).z = ArkHeightMax + 25 then
      ArkPlatformTrigger3.Enabled = 1
      ArkKicker3.kick 0,0
      arkballs(3) = NULL
    End If
  End If
' if Not isNull(arkballs(1)) and Arkballs(1).z = ArkHeightMax + 25 then ArkKicker0.kick 0,0: arkballs(1) = NULL
' if Not isNull(arkballs(2)) and Arkballs(2).z = ArkHeightMax + 25 then ArkKicker0.kick 0,0: arkballs(2) = NULL
' if Not isNull(arkballs(3)) and Arkballs(3).z = ArkHeightMax + 25 then ArkKicker0.kick 0,0: arkballs(3) = NULL

  'Check Motor Direction
  If ArkTempPos >= ArkHeightMax +3 Then
    Arkdirection = -1
  ElseIf ArkTempPos <= 0 -3 Then
    Arkdirection = 1
  End If


End Sub

Sub ArkPlatformTrigger0_UnHit:NewBallID: me.enabled=0: me.timerinterval = 500: me.timerenabled = 1:End Sub
Sub ArkPlatformTrigger0_Timer: GI_TroughCheck: me.timerenabled = 0: End Sub
Sub ArkPlatformTrigger1_UnHit:NewBallID: me.enabled=0:End Sub
Sub ArkPlatformTrigger2_UnHit:NewBallID: me.enabled=0:End Sub
Sub ArkPlatformTrigger3_UnHit:NewBallID: me.enabled=0:End Sub


' Ark Ball Count

Dim arkballs(4), arkballcnt
arkballs(0) = NULL:arkballs(1) = NULL:arkballs(2) = NULL:arkballs(3) = NULL

Sub sw39_Init: Controller.Switch(39) = 0 :End Sub 'enter ark
Sub sw39_Hit: 'Ark Opto
  Controller.Switch(39) = 1
  me.TimerInterval = 300  '500
  me.TimerEnabled = 1

' 'Track all four balls as they enter ark
  Set arkballs(arkballcnt) = ActiveBall
' debug.print "sw39 hit" & arkballcnt & ":" & arkballs(arkballcnt).x
' arkballcnt = (arkballcnt+1) mod 4
End Sub
Sub sw39_UnHit
  Controller.Switch(39) = 0:
' ClearballId
  Activeball.DestroyBall
  'Track all four balls as they enter ark
  Select Case arkballcnt:
    Case 0:
      Set arkballs(0) = arkkicker0.CreateBall
      'debug.print "Arkball0"
    Case 1:
      Set arkballs(1) = arkkicker1.CreateBall
      'debug.print "Arkball1"
    Case 2:
      Set arkballs(2) = arkkicker2.CreateBall
      'debug.print "Arkball2"
    Case 3:
      Set arkballs(3) = arkkicker3.CreateBall
      'debug.print "Arkball3"
            tablelaunch.visible = 0  'Turn Off Ball Info
  End Select

' Set arkballs(arkballcnt) = ActiveBall
' debug.print "sw39 hit" & arkballcnt & ":" & arkballs(arkballcnt).x
  arkballcnt = (arkballcnt+1) mod 5
End Sub
Sub sw39_Timer
  GI_TroughCheck 'update ball count after ball enters ark
  me.TimerEnabled = 0
End Sub

Sub sw49_Hit:Controller.Switch(49) = 1:End Sub 'enter launch ramp
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub



Const DTMaxRotate = 30
Dim TempleSpeed:TempleSpeed = .25
Dim TempleDir: TempleDir = 1
Dim TPos,PrimPos
Dim TempMotorUp, TempMotorDown, TempSoundTrigger1, TempSoundTrigger2

DT.Interval = 30
Sub DT_Timer
  'Change angle
  TPos = TPos + TempleDir*TempleSpeed
  PrimPos = TPos



  'Set Switches and Walls
  If TPos >= DTMaxRotate Then  'If up all the way
    Controller.Switch(52) = 1
    PrimPos = DTMaxRotate
    sw59.IsDropped = True
    TempleWall1.IsDropped = True
    TempleWall2.IsDropped = True
    templebase.isdropped = False

  Else
    Controller.Switch(52) = 0
  End If
  If TPos <= 0 Then           'If down all the way
    Controller.Switch(53) = 1
    PrimPos = 0
    sw59.IsDropped = False
    TempleWall1.IsDropped = False
    TempleWall2.IsDropped = False
    templebase.isdropped = True

  Else
    Controller.Switch(53) = 0
  End If

  ' Sound Aktivation

   If TPos = 1 AND TempleDir = 1 Then
  TempSoundTrigger1 = 1
    TempMotorUp = 1
    End If


   If TPos = 1 AND TempleDir = -1 Then
    StopSound "E_Motorlift4"
    TempMotorDown = 0
    End If


    If TPos = DTMaxRotate-1 AND TempleDir = -1 Then
  TempSoundTrigger2 = 1
    TempMotorDown = 1
    End If


    If TPos = DTMaxRotate-1 AND TempleDir = 1 Then
    StopSound "E_Motorlift2"
    TempMotorUp = 0
    End If


  'Check Motor Direction
  If TPos >= DTMaxRotate + 1 Then 'If up all the way + padding
    TempleDir = -1
  ElseIf TPos <= 0 - 1 Then     'If down all the way + padding
    TempleDir = 1
  End If


' Hanibals Sound Routine for Temple moving



If TempMotorUp = 1 AND TempSoundTrigger1 = 1 Then
        PlaySound "E_Motorlift2", 0, 40 / (25*Rnd), -0.05, 0.15 ' TODO
        TempSoundTrigger1 = 0

End If

If TempMotorDown = 1 AND TempSoundTrigger2 = 1  Then
        PlaySound "E_Motorlift4", 0, 40 / (25*Rnd), -0.05, 0.15
    TempSoundTrigger2 = 0
End If





  'Move Primitives
  sw59p.Rotx = PrimPos
  temple1.Rotx = PrimPos
  temple2.Rotx = PrimPos
  temple3.Rotx = PrimPos
  temple4.Rotx = PrimPos
  temple5.Rotx = PrimPos
  temple6.Rotx = PrimPos
  temple7.Rotx = PrimPos
  temple8.Rotx = PrimPos
  templeball1.Rotx = PrimPos
  templeball2.Rotx = PrimPos
    Lift = PrimPos


End Sub





' Hanibals Sound Routine for Temple moving

Sub TempMotorSound

If TempMotorUp = 1 Then
        PlaySound "E_Motorlift2", 0, 40 / (25*Rnd), -0.05, 0.15
    Else
    StopSound "E_Motorlift2"
End If

If TempMotorDown = 1 Then
        PlaySound "E_Motorlift4", 0, 40 / (25*Rnd), -0.05, 0.15
    Else
    StopSound "E_Motorlift4"
End If

End Sub


Const SMMaxRotate = 80
Const SMMinRotate = -10
Dim SwordsManSpeed:SwordsManSpeed = 1.5
Dim SwordsManDir: SwordsManDir = 1
Dim SMPos,SMPrimPos
SM.Interval = 30
SMPos = SMMinRotate
Sub SM_Timer
  'Change angle
  SMPos = SMPos + SwordsManDir*SwordsManSpeed
  SMPrimPos = SMPos

  'Set Switches
  If SMPos >= SMMaxRotate Then
    Controller.Switch(64) = 1
    SMPrimPos = SMMaxRotate
  Else
    Controller.Switch(64) = 0
  End If
  If SMPos <= SMMinRotate Then
    Controller.Switch(63) = 1
    SMPrimPos = SMMinRotate
  Else
    Controller.Switch(63) = 0
  End If

  'Check Motor Direction
  If SMPos >= SMMaxRotate + 1 Then
    SwordsManDir = -1
  ElseIf SMPos <= SMMinRotate - 1 Then
    SwordsManDir = 1
  End If

  'Move Primitives
  swordsmanprim.ObjRotZ = SMPrimPos:'debug.print SMPrimPos
End Sub


'----- Debug Print
Dim dString
Const debugOn = True
Sub dprint (dString)
  If debugOn = True Then
    debug.print dString
  End If
End Sub


'--------------------
'     8 BallStack
'--------------------
Private Const con8StackSw    = 8  ' Stack switches
Class cvpm8BallStack

  Private mSw(), mEntrySw, mBalls, mBallIn, mBallPos(), mSaucer, mBallsMoving
  Private mInitKicker, mExitKicker, mExitDir, mExitForce
  Private mExitDir2, mExitForce2
  Private mEntrySnd, mEntrySndBall, mExitSnd, mExitSndBall, mAddSnd
  Public KickZ, KickBalls, KickForceVar, KickAngleVar

  Private Sub Class_Initialize
    ReDim mSw(con8StackSw), mBallPos(conMaxBalls)
    mBallIn = 0 : mBalls = 0 : mExitKicker = 0 : mInitKicker = 0 : mBallsMoving = False
    KickBalls = 1 : mSaucer = False : mExitDir = 0 : mExitForce = 0
    mExitDir2 = 0 : mExitForce2 = 0 : KickZ = 0 : KickForceVar = 0 : KickAngleVar = 0
    mAddSnd = 0 : mEntrySnd = 0 : mEntrySndBall = 0 : mExitSnd = 0 : mExitSndBall = 0
    vpmTimer.AddResetObj Me
  End Sub

  Private Property Let NeedUpdate(aEnabled) : vpmTimer.EnableUpdate Me, False, aEnabled : End Property

  Private Function SetSw(aNo, aStatus)
    SetSw = False : If HasSw(aNo) Then Controller.Switch(mSw(aNo)) = aStatus : SetSw = True
  End Function

  Private Function HasSw(aNo)
    HasSw = False : If aNo <= con8StackSw Then If mSw(aNo) Then HasSw = True
  End Function

  Public Sub Reset
    Dim ii : If mBalls Then For ii = 1 to mBalls : SetSw mBallPos(ii), True : Next
        If mEntrySw And mBallIn > 0 Then Controller.Switch(mEntrySw) = True
  End Sub

  Public Sub Update
    Dim BallQue, ii
    NeedUpdate = False : BallQue = 1
    For ii = 1 To mBalls
      If mBallpos(ii) > BallQue Then ' next slot available
        NeedUpdate = True
        If HasSw(mBallPos(ii)) Then ' has switch
          If Controller.Switch(mSw(mBallPos(ii))) Then
            SetSw mBallPos(ii), False
          Else
            mBallPos(ii) = mBallPos(ii) - 1
            SetSw mBallPos(ii), True
          End If
        Else ' no switch. Move ball to first switch or occupied slot
          Do
            mBallPos(ii) = mBallPos(ii) - 1
          Loop Until SetSw(mBallPos(ii), True) Or mBallPos(ii) = BallQue
        End If
      End If
      BallQue = mBallPos(ii) + 1
    Next
  End Sub

  Public Sub AddBall(aKicker)
    If isObject(aKicker) Then
      If mSaucer Then
        If aKicker Is mExitKicker Then
          mExitKicker.Enabled = False : mInitKicker = 0
        Else
          aKicker.Enabled = False : Set mInitKicker = aKicker
        End If
      Else
        aKicker.DestroyBall
      End If
    ElseIf mSaucer Then
      mExitKicker.Enabled = False : mInitKicker = 0
    End If
    If mEntrySw Then
      Controller.Switch(mEntrySw) = True : mBallIn = mBallIn + 1
    Else
      mBalls = mBalls + 1 : mBallPos(mBalls) = con8StackSw + 1 : NeedUpdate = True
    End If
    PlaySound mAddSnd
  End Sub

  ' A bug in the script engine forces the "End If" at the end
  Public Sub SolIn(aEnabled)     : If aEnabled Then KickIn        : End If : End Sub
  Public Sub SolOut(aEnabled)    : If aEnabled Then KickOut False : End If : End Sub
  Public Sub SolOutAlt(aEnabled) : If aEnabled Then KickOut True  : End If : End Sub
  Public Sub EntrySol_On   : KickIn        : End Sub
  Public Sub ExitSol_On    : KickOut False : End Sub
  Public Sub ExitAltSol_On : KickOut True  : End Sub

  Private Sub KickIn
    If mBallIn Then PlaySound mEntrySndBall Else PlaySound mEntrySnd : Exit Sub
    mBalls = mBalls + 1 : mBallIn = mBallIn - 1 : mBallPos(mBalls) = con8StackSw + 1 : NeedUpdate = True
    If mEntrySw And mBallIn = 0 Then Controller.Switch(mEntrySw) = False
  End Sub

  Private Sub KickOut(aAltSol)
    Dim ii,jj, kForce, kDir, kBaseDir
    If mBalls Then PlaySound mExitSndBall Else PlaySound mExitSnd : Exit Sub
    If aAltSol Then kForce = mExitForce2 : kBaseDir = mExitDir2 Else kForce = mExitForce : kBaseDir = mExitDir
    kForce = kForce + (Rnd - 0.5)*KickForceVar
    If mSaucer Then
      SetSw 1, False : mBalls = 0 : kDir = kBaseDir + (Rnd - 0.5)*KickAngleVar
      If isObject(mInitKicker) Then
        vpmCreateBall mExitKicker : mInitKicker.Destroyball : mInitKicker.Enabled = True
      Else
        mExitKicker.Enabled = True
      End If
      mExitKicker.Kick kDir, kForce, KickZ
    Else
      For ii = 1 To kickballs
        If mBalls = 0 Or mBallPos(1) <> ii Then Exit For ' No more balls
        For jj = 2 To mBalls ' Move balls in array
          mBallPos(jj-1) = mBallPos(jj)
        Next
        mBallPos(mBalls) = 0 : mBalls = mBalls - 1 : NeedUpdate = True
        SetSw ii, False
        If isObject(mExitKicker) Then
          If kForce < 1 Then kForce = 1
          kDir = kBaseDir + (Rnd - 0.5)*KickAngleVar
          vpmTimer.AddTimer (ii-1)*200, "vpmCreateBall(" & mExitKicker.Name & ").Kick " &_
            CInt(kDir) & "," & Replace(kForce,",",".") & "," & Replace(KickZ,",",".") & " '"
        End If
        kForce = kForce * 0.8
      Next
    End If
  End Sub

  Public Sub InitSaucer(aKicker, aSw, aDir, aPower)
    InitKick aKicker, aDir, aPower : mSaucer = True
    If aSw Then mSw(1) = aSw Else mSw(1) = aKicker.TimerInterval
  End Sub

  Public Sub InitNoTrough(aKicker, aSw, aDir, aPower)
    InitKick aKicker, aDir, aPower : Balls = 1
    If aSw Then mSw(1) = aSw Else mSw(1) = aKicker.TimerInterval
    If Not IsObject(vpmTrough) Then Set vpmTrough = Me
  End Sub

  Public Sub InitSw(aEntry, aSw1, aSw2, aSw3, aSw4, aSw5, aSw6, aSw7, aSw8)
    mEntrySw = aEntry : mSw(1) = aSw1 : mSw(2) = aSw2 : mSw(3) = aSw3 : mSw(4) = aSw4
    mSw(5) = aSw5 : mSw(6) = aSw6 : mSw(7) = aSw7: mSw(8) =aSw8
    If Not IsObject(vpmTrough) Then Set vpmTrough = Me
  End Sub

  Public Sub InitKick(aKicker, aDir, aForce)
    Set mExitKicker = aKicker : mExitDir = aDir : mExitForce = aForce
  End Sub

  Public Sub CreateEvents(aName, aKicker)
    Dim obj, tmp
    If Not vpmCheckEvent(aName, Me) Then Exit Sub
    vpmSetArray tmp, aKicker
    For Each obj In tmp
      If isObject(obj) Then
        vpmBuildEvent obj, "Hit", aName & ".AddBall Me"
      Else
        vpmBuildEvent mExitKicker, "Hit", aName & ".AddBall Me"
      End If
    Next
  End Sub

  Public Property Let IsTrough(aIsTrough)
    If aIsTrough Then
      Set vpmTrough = Me
    ElseIf IsObject(vpmTrough) Then
      If vpmTrough Is Me Then vpmTrough = 0
    End If
  End Property

  Public Property Get IsTrough : IsTrough = vpmTrough Is Me : End Property

  Public Sub InitAltKick(aDir, aForce)
    mExitDir2 = aDir : mExitForce2 = aForce
  End Sub

  Public Sub InitEntrySnd(aBall, aNoBall) : mEntrySndBall = aBall : mEntrySnd = aNoBall : End Sub
  Public Sub InitExitSnd(aBall, aNoBall)  : mExitSndBall = aBall  : mExitSnd = aNoBall  : End Sub
  Public Sub InitAddSnd(aSnd) : mAddSnd = aSnd : End Sub

  Public Property Let Balls(aBalls)
    Dim ii
    For ii = 1 To con8StackSw
      SetSw ii, False : mBallPos(ii) = con8StackSw + 1
    Next
    If mSaucer And aBalls > 0 And mBalls = 0 Then vpmCreateBall mExitKicker
    mBalls = aBalls : NeedUpdate = True
  End Property

  Public Default Property Get Balls : Balls = mBalls         : End Property
  Public Property Get BallsPending  : BallsPending = mBallIn : End Property

  ' Obsolete stuff
  Public Sub SolEntry(aSnd1, aSnd2, aEnabled)
    If aEnabled Then mEntrySndBall = aSnd1 : mEntrySnd = aSnd2 : KickIn
  End Sub
  Public Sub SolExit(aSnd1, aSnd2, aEnabled)
    If aEnabled Then mExitSndBall = aSnd1 : mExitSnd = aSnd2 : KickOut False
  End Sub
  Public Sub InitProxy(aProxyPos, aSwNo) : End Sub
  Public TempBallColour, TempBallImage, BallColour
  Public Property Let BallImage(aImage) : vpmBallImage = aImage : End Property

End Class

'General Sound Handler

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / (15*Rnd), -0.05, 0.15
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / (15*Rnd), 0.05, 0.15
End Sub

Sub Rubbers_Hit(idx)
  PlaySound "fx_rubber2", 0, (20 + Vol(ActiveBall)), (20+Pan(ActiveBall)), Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
End Sub

Sub Pins_Hit (idx)
  PlaySound "metalhit_medium", 0, (8 +Vol(ActiveBall)),  Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
    'Test1.State=1:
End Sub

Sub Primitive_Hit (idx)
  PlaySound "metalhit_medium", 0, (8 +Vol(ActiveBall)),  Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
'    Test1.State=1:
End Sub


Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, (8 +Vol(ActiveBall)), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, (8 +Vol(ActiveBall)), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, (8 +Vol(ActiveBall)), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate", 0, (8 +Vol(ActiveBall)), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
End Sub

Sub Posts_Hit(idx)
    PlaySound "fx_postrubber", 0, (8 +Vol(ActiveBall)), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall)
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Table" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Table.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Table" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Table.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Table" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
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

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 14 ' total number of balls
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
    Dim Rampy
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
        If BallVel(BOT(b) ) > 1  Then 'AND BOT(b).z < 30
            If BOT(b).z > 30 Then
            Rampy = 2
            Else
      Rampy = 1
      End if

            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b))*Rampy, 1, 0
      '***********************************************************
      '      Hanibal's VP10 Automated Collision Detection Sounds
      '***********************************************************

      IF Pan(BOT(b) ) > 8 Then
      PlaySound("fx_rubber2"), 0, Vol(BOT(b)), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
      End IF


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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub


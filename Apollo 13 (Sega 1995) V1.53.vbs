'==================================================================================================================================================='
'                                                                                                   '
'                                  APOLLO 13                                                                                      '
'                                 Sega Pinball (1995)                                                                                         '
'               http://www.ipdb.org/machine.cgi?id=3592                                                                   '
'                                                                                                   '
'                     Created by ICPjuggla and Herweh vp9                                                                   '
'                    Balater vpx                                                                          '
'               Contribution  Thalamus, JP's, DJRobX, Rothbauerw, Amgrim, Bigus1, 32assassin,ClarKent, Chucky87                   '
'==================================================================================================================================================='

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const UseVPMModSol = 1

Dim  varhidden, UseVPMColoredDMD, bgsko
  bgsko=0
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
  varhidden=1
Else
    UseVPMColoredDMD = False
  varhidden=0
End If




' ***********************************************************************************************
' OPTIONS
' ***********************************************************************************************


' VPinMAME ROM name
Const cGameName       = "apollo13"    ' enter string of valid ROM



' ball size
Const BallSize        = 50        ' sets the ball size


' ***********************************************************************************************
' OPTIONS END
' ***********************************************************************************************
' Thalamus 2020 January : Improved directional sounds/
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 0.2  ' Bumpers volume.
Const VolRol    = 0.1    ' Rollovers volume.
Const VolGates  = 0.1  ' Gates volume.
Const VolMetal  = 0.1    ' Metals volume.
Const VolRB     = 0.1   ' Rubber bands volume.
Const VolRH     = 0.1    ' Rubber hits volume.
Const VolPo     = 0.1   ' Rubber posts volume.
Const VolPi     = 0.1    ' Rubber pins volume.
Const VolPlast  = 0.1    ' Plastics volume.
Const VolTarg   = 0.1    ' Targets volume.
Const VolWood   = 0.1   ' Woods volume.
Const VolKick   = 0.1    ' Kicker volume.
Const VolSpin   = 0.15  ' Spinners volume.
Const VolFlip   = 0.1    ' Flipper volume.


' ===============================================================================================
' some general constants and variables
' ===============================================================================================

Const UseLamps      = False
Const UseGI       = False
Const UseSync       = True
Const HandleMech    = False

Const SSolenoidOn     = "SolenoidOn"
Const SSolenoidOff    = "SolenoidOff"
Const SCoin       = "Coin"
Const SKnocker      = "Knocker"

Dim I, x, obj, bsTrough, bsTroughVUK, bsTroughSuperVUK, bsRocket, bsTrough8BL, nbb, tempo, sap, drp, rl, rfd, lfd, flipp, deb18
Dim mUpperSaucer, mRightSaucer, bsUpperSaucer, bsRightSaucer, mmoon,bsball003, bsBall002, ball8, noball8, noball8b, bkgg

Const SFlipperOn="Flipper2",SFlipperOff="Flipper"

' ===============================================================================================
' load game controller
' ===============================================================================================

deb18=0 '1 visualization of the loading of the balls before the first game
flipp=0 '1 flipper works continuously
bkgg=1  '1 GI Animation on Backglass

LoadVPM  "01560000", "SEGA2.VBS", 3.56



' ===============================================================================================
' solenoids
' ===============================================================================================

SolCallback(1)      = "SolTroughVUKOut"
SolCallback(2)      = "SolAPlunger"
SolCallback(3)      = "SolRightEject"
SolCallback(4)      = "SolTopEject"
SolCallback(5)      = "SolRocketEject"
SolCallback(7)      = "ShakeRocket"
SolCallback(8)      = "SolKnocker"
SolCallback(14)         = "SolTroughSuperVUKOut"
SolCallback(15)         = "SolLFlipperd"
SolCallback(16)         = "SolRFlipperd"
SolCallback(17)     = "SolTroughOut"
SolCallback(18)     = "SolRampLift"
SolCallback(19)     = "SolmoonLift"
SolCallback(20)     = "SolRocketLift"
SolCallback(21)     = "SolTrapDoor"
SolCallback(22)     = "Sol8BLockPlunger"
SolCallback(23)     = "SolDrainPost"
SolCallback(34)     = "SolMoon"


' flasher
SolCallback(25)     = "SolFLamp1"
SolCallback(26)     = "SolFLamp2"
SolCallback(27)     = "SolFLamp3"
SolCallback(28)     = "SolFLamp4"
SolCallback(29)     = "SolFLamp5"
SolCallback(30)     = "SolFLamp6"
SolCallback(31)     = "SolFLamp7"
SolCallback(32)     = "SolFLamp8"


' =====================================================================================================
' table events
' =====================================================================================================
Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
  If bgsko Then
    If enabled Then
      Controller.b2ssetdata 79,1
    Else
      Controller.b2ssetdata 79,0
    End If
  End If
gi1.state=Enabled
gi2.state=Enabled
gi3.state=Enabled
gi4.state=Enabled
gi5.state=Enabled
gi6.state=Enabled
gi7.state=Enabled
gi8.state=Enabled
gi9.state=Enabled
gi10.state=Enabled
gi11.state=Enabled
gi12.state=Enabled
gi13.state=Enabled
gi14.state=Enabled
gi15.state=Enabled
gi16.state=Enabled
gi17.state=Enabled
gi18.state=Enabled
gi19.state=Enabled
gi20.state=Enabled
gi21.state=Enabled
gi22.state=Enabled
gi23.state=Enabled
GI24.state=Enabled
GI25.state=Enabled
GI26.state=Enabled
GI27.state=Enabled
GI28.state=Enabled
GI29.state=Enabled
GI30.state=Enabled
GI31.state=Enabled
GI32.state=Enabled
GI33.state=Enabled
GI34.state=Enabled
GI35.state=Enabled
GI36.state=Enabled
GI37.state=Enabled
GI38.state=Enabled
GI39.state=Enabled
GI40.state=Enabled
GI41.state=Enabled
GI42.state=Enabled
GI43.state=Enabled
gi44.state=Enabled
gi45.state=Enabled
gi46.state=Enabled
Flasher001.visible = Enabled
Flasher002.visible = Enabled
Flasher003.visible = Enabled
Flasher004.visible = Enabled
GI47.state=Enabled
GI48.state=Enabled
GI49.state=Enabled
GI50.state=Enabled
GI51.state=Enabled
GI52.state=Enabled
GI53.state=Enabled
GI54.state=Enabled
GI55.state=Enabled
GI56.state=Enabled
GI57.state=Enabled
GI58.state=Enabled
GI59.state=Enabled
GI60.state=Enabled
GI61.state=Enabled
GI62.state=Enabled
GI63.state=Enabled

End Sub

Sub Table1_Init
  On Error Resume Next
  With Controller
    .GameName=cGameName
    .SplashInfoLine             ="Apollo 13, Sega, 1995"
    .HandleKeyboard             = False
    .ShowTitle                = False
    .ShowDMDOnly              = True
    .ShowFrame                = False
    .Hidden                 = VarHidden
    .Games(cGameName).Settings.Value("rol") = 0
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
  If b2son Then bgsko = bkgg
Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
  ' Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  ' basic pinmame timer
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled  = True

  ' nudging
  vpmNudge.TiltSwitch   = 1
  vpmNudge.Sensitivity  = 3
  vpmNudge.TiltObj    = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)



  Set bsTroughSuperVUK = New cvpmBallStack
    bsTroughSuperVUK.InitSw 0, 48, 0, 0, 0, 0, 0, 0
    bsTroughSuperVUK.InitKick BallReleaseSuperVUK, 110, 0.1
    bsTroughSuperVUK.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("SolOn",DOFContactors)
  Set bsball002 = New cvpmBallStack
    bsball002.InitSw 0, 15, 0, 0, 0, 0, 0, 0
    bsball002.InitKick Ball002, 0, 0
    bsball002.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("SolOn",DOFContactors)
  Set bsball003 = New cvpmBallStack
    bsball003.InitKick Ball003, 45, 10

     ' magnet
  Set mmoon = New cvpmMagnet
    With mmoon
      .InitMagnet mMoonLock, 90
      .GrabCenter = True
      .MagnetOn = 0
      .CreateEvents "mmoon"
    End With

    wall173.collidable=1
    releaseout.collidable=1
    moonballstop.collidable=0
    moonballstopb.collidable=0
    mmoon.MagnetOn = 0
    moonlock.Enabled=0
    Ramp139.collidable=1
    Ramp029.collidable=0
    BallReleaseSuperVUK.CreateSizedBall(25)
    BallReleaseSuperVUK.Kick 180,0.1
    BallReleaseSuperVUK001.CreateSizedBall(25)
    BallReleaseSuperVUK001.Kick 180,0.1
    BallReleaseSuperVUK001.enabled=false
    BallReleaseSuperVUK002.CreateSizedBall(25)
    BallReleaseSuperVUK002.Kick 180,0.1
    BallReleaseSuperVUK002.enabled=false
    BallReleaseSuperVUK003.CreateSizedBall(25)
    BallReleaseSuperVUK003.Kick 1800,0.1
    BallReleaseSuperVUK003.enabled=false
    BallReleaseSuperVUK004.CreateSizedBall(25)
    BallReleaseSuperVUK004.Kick 180,0.1
    BallReleaseSuperVUK004.enabled=false
    If deb18 Then
      noball8b=1
      noball8=1
      Kicker001.enabled=1
      for i = 0 to 7
        Kicker001.CreateSizedBall(25)
        Kicker001.Kick 110,0.1
      Next
      Kicker001.enabled=false
    Else
      sw24b.CreateSizedBall(25)
      sw23b.CreateSizedBall(25)
      sw22b.CreateSizedBall(25)
      sw21b.CreateSizedBall(25)
      sw20b.CreateSizedBall(25)
      sw19b.CreateSizedBall(25)
      sw18b.CreateSizedBall(25)
      sw17b.CreateSizedBall(25)
      Controller.Switch(17) = 1
      Controller.Switch(18) = 1
      Controller.Switch(19) = 1
      Controller.Switch(20) = 1
      Controller.Switch(21) = 1
      Controller.Switch(22) = 1
      Controller.Switch(23) = 1
      Controller.Switch(24) = 1
    End If
    Wallpost001.collidable=0
    InitMoon
    InitRocket
End Sub

Sub Table1_Exit()
  Controller.Stop
End Sub

Sub Table1_Paused
  Controller.Pause = True
End Sub
Sub Table1_UnPaused
  Controller.Pause = False
End Sub

Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftFlipperKey Then SolLFlipper 1
  If keycode = RightFlipperKey Then SolRFlipper 1
  If vpmKeyDown(keycode) Then Exit Sub
  If keycode = PlungerKey Then Controller.Switch(9) = True
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then SolLFlipper 0
  If keycode = RightFlipperKey Then SolRFlipper 0
  If vpmKeyUp(KeyCode) Then Exit Sub
  If keycode = PlungerKey Then Controller.Switch(9) = False
End Sub


' ===============================================================================================
' realtime updates
' ===============================================================================================

'Set MotorCallback = GetRef("RealTimeUpdates")



' ===============================================================================================
' ball draining, release and auto plunger
' ===============================================================================================

Sub SolTroughOut(Enabled)
  If Enabled Then
  if sap= 0 Then
  If bsball003.Balls <= 0 Then
  releaseout.collidable= 0
  End If
  End If
End If
End Sub

Sub SolTroughVUKOut(Enabled)
  If Enabled Then
    If Controller.Switch(16) =0 Then
      bsball003.ExitSol_On
    End If
  End If
End Sub

Sub ball002_Hit()
  releaseout.collidable=1
  If bsball003.Balls <= 0 Then
    me.destroyball
    bsball003.AddBall Me
  End If
End Sub


Sub sw10_Hit()
  Controller.Switch(10)   = True
End Sub
Sub sw10_Unhit()
  Controller.Switch(10)   = False
End Sub
Sub sw11_Hit()
  Controller.Switch(11)   = True
End Sub
Sub sw11_Unhit()
  Controller.Switch(11)   = False
End Sub
Sub sw12_Hit()
  Controller.Switch(12)   = True
End Sub
Sub sw12_Unhit()
  Controller.Switch(12)   = False
End Sub
Sub sw13_Hit()
  Controller.Switch(13)   = True
End Sub
Sub sw13_Unhit()
  Controller.Switch(13)   = False
End Sub
Sub sw14_Hit()
  Controller.Switch(14)   = True
End Sub
Sub sw14_Unhit()
  Controller.Switch(14)   = False
End Sub
Sub sw151_Hit()
  Controller.Switch(15)   = True
  releaseout.collidable=1
End Sub
Sub sw151_Unhit()
  Controller.Switch(15)   = False
End Sub
Sub sw16_Hit()
  Controller.Switch(16)   = True
  if sap Then
    AutoPlunger.Kick 0, 35+Int(rnd(30))
    PlaySoundAtVol SoundFX("SolenoidOn",DOFContactors), sw16, 1
  End If
End Sub
sub sw16t_hit()
  sap=0
End Sub
sub sw16b_hit()
  If noball8b=1 then
    MultiBallPost.IsDropped = 0
    wall173.collidable= 1
    sw24b.enabled=1
    sw23b.enabled=0
    sw22b.enabled=0
    sw21b.enabled=0
    sw20b.enabled=0
    sw19b.enabled=0
    sw18b.enabled=0
    sw17b.enabled=0
    sw24.enabled=1
    sw23.enabled=0
    sw22.enabled=0
    sw21.enabled=0
    sw20.enabled=0
    sw19.enabled=0
    sw18.enabled=0
    sw17.enabled=0
    noball8b=0
  End If
  if drp Then
    Ramp139.collidable=1
    ramp029.collidable=0
    PlaySoundAtVol SoundFX("SolenoidOn",DOFContactors), sw16b, 1
  End If
End Sub
Sub sw16_Unhit()
  Controller.Switch(16)   = False
End Sub
Sub sw17b_Hit()
  Controller.Switch(17)   = True
  If noball8 Then
    noball8=0
  End If
End Sub
Sub sw17b_Unhit()
  Controller.Switch(17)   = False
End Sub
Sub sw18b_Hit()
  Controller.Switch(18)   = True
  If noball8 Then
    sw17.enabled=1
    sw17b.enabled=1
  End If
End Sub
Sub sw18b_Unhit()
  Controller.Switch(18)   = False
End Sub
Sub sw19b_Hit()
  Controller.Switch(19)   = True
  If noball8 Then
    sw18.enabled=1
    sw18b.enabled=1
  End If
End Sub
Sub sw19b_Unhit()
  Controller.Switch(19)   = False
End Sub
Sub sw20b_Hit()
  Controller.Switch(20)   = True
  If noball8 Then
    sw19.enabled=1
    sw19b.enabled=1
  End If
End Sub
Sub sw20b_Unhit()
  Controller.Switch(20)   = False
End Sub
Sub sw21b_Hit()
  Controller.Switch(21)   = True
  If noball8 Then
    sw20.enabled=1
    sw20b.enabled=1
  End If
End Sub
Sub sw21b_Unhit()
  Controller.Switch(21)   = False
End Sub
Sub sw22b_Hit()
  Controller.Switch(22)   = True
  If noball8 Then
    sw21.enabled=1
    sw21b.enabled=1
  End If
End Sub
Sub sw22_Unhit()
  Controller.Switch(22)   = False
End Sub
Sub sw23b_Hit()
  Controller.Switch(23)   = True
  If noball8 Then
    sw22.enabled=1
    sw22b.enabled=1
  End If
End Sub
Sub sw23b_Unhit()
  Controller.Switch(23)   = False
End Sub
Sub sw24b_Hit()
  Controller.Switch(24)   = True
  If noball8 Then
    sw23.enabled=1
    sw23b.enabled=1
  End If
End Sub
Sub sw24b_Unhit()
  Controller.Switch(24)   = False
End Sub
Sub sw32_hit()
vpmTimer.PulseSw 32
End Sub
Sub sw31_hit()
vpmTimer.PulseSw 31
End Sub
Sub sw30_hit()
vpmTimer.PulseSw 30
End Sub
Sub sw29_hit()
vpmTimer.PulseSw 29
End Sub
Sub sw28_hit()
vpmTimer.PulseSw 28
End Sub
Sub sw27_hit()
vpmTimer.PulseSw 27
End Sub
Sub sw26_hit()
vpmTimer.PulseSw 26
End Sub
Sub sw25_hit()
vpmTimer.PulseSw 25
End Sub
Sub sw38_hit()
vpmTimer.PulseSw 38
End Sub
Sub sw40_hit()
vpmTimer.PulseSw 40
End Sub
Sub sw49a_hit()
vpmTimer.PulseSw 49
End Sub
Sub sw49b_hit()
vpmTimer.PulseSw 49
End Sub
Sub sw49c_hit()
vpmTimer.PulseSw 49
End Sub


Sub SolAPlunger(Enabled)
  if controller.Switch(16) then
    AutoPlunger.Kick 0, 35-int(Rnd(30))

'   PlaySound "SolenoidOn"
    PlaySoundAtVol SoundFX("SolenoidOn",DOFContactors), AutoPlunger, 1
    sap=1
  end If
End Sub
'flipper
Sub sollFlipperd(Enabled)
     lfd=Enabled
 End Sub
Sub solrFlipperd(Enabled)
     rfd=Enabled
End Sub
Sub sollFlipper(Enabled)
  if lfd  or rfd or flipp Then
     If Enabled Then
         PlaySoundAtVol SoundFXDOF("fx_Flipperup",101,DOFOn,DOFFlippers),LeftFlipper,1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFXDOF("fx_Flipperdown",101,DOFOff,DOFFlippers),LeftFlipper,1:LeftFlipper.RotateToStart
     End If
  Else
    LeftFlipper.RotateToStart
  End If
  End Sub

Sub solrflipper(Enabled)
  if rfd or lfd or flipp Then
     If Enabled Then
         PlaySoundAtVol SoundFXDOF("fx_Flipperup",102,DOFOn,DOFFlippers), RightFlipper, 1:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFXDOF("fx_Flipperdown",102,DOFOff,DOFFlippers), RightFlipper, 1:RightFlipper.RotateToStart
     End If
  Else
    RightFlipper.RotateToStart
  End If
End Sub


'Flippers()


Sub LeftFlipper_Collide(parm)
  CollideWithFlipper
End Sub
Sub RightFlipper_Collide(parm)
  CollideWithFlipper
End Sub

Sub CollideWithFlipper()
  If BallSpeed > 4 then
    RandomSoundFlipper()
  Else
    RandomSoundFlipperLowVolume()
  End If
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtVol "flip_hit_1", ActiveBall, 1
    Case 2 : PlaySoundAtVol "flip_hit_2", ActiveBall, 1
    Case 3 : PlaySoundAtVol "flip_hit_3", ActiveBall, 1
  End Select
End Sub

Sub RandomSoundFlipperLowVolume()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtVol "flip_hit_1_low", ActiveBall, 1
    Case 2 : PlaySoundAtVol "flip_hit_2_low", ActiveBall, 1
    Case 3 : PlaySoundAtVol "flip_hit_3_low", ActiveBall, 1
  End Select
End Sub




' ===============================================================================================
' slingshots events and slingshot animation scripting from JPSalas
' ===============================================================================================






' ===============================================================================================
' bumpers
' ===============================================================================================
Sub Bumper1_Hit()
vpmTimer.PulseSw 35
PlayBumperSound
End Sub

Sub Bumper2_Hit()
vpmTimer.PulseSw 33
PlayBumperSound
End Sub

Sub Bumper3_Hit()
vpmTimer.PulseSw 34
PlayBumperSound
End Sub

Sub PlayBumperSound()
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtVol SoundFX("bumper_1",DOFContactors), ActiveBall, 1
      Case 2 : PlaySoundAtVol SoundFX("bumper_2",DOFContactors), ActiveBall, 1
      Case 3 : PlaySoundAtVol SoundFX("bumper_3",DOFContactors), ActiveBall, 1
    End Select
End Sub


' ===============================================================================================
' knocker
' ===============================================================================================

Sub SolKnocker(Enabled)
  If Enabled Then PlaySound SoundFX("knocker",DOFKnocker)
End Sub

' ===============================================================================================
' ramp diverter
' ===============================================================================================

Dim diverterAdd

InitDiverter

Sub InitDiverter()
  Diverter2.collidable = 0
  Diverter1.collidable = 1
  diverterAdd     = 1
End Sub

Sub DiverterTriggerLeft_Hit()
  ActiveBall.VelY       = 0
  Diverter1.collidable    = 1
  Diverter2.collidable    = 0
  diverterAdd         = -3
  RampDiverterTimer.Enabled   = True
End Sub
Sub DiverterTriggerRight_Hit()
  ActiveBall.VelY       = 0
  Diverter1.collidable    = 0
  Diverter2.collidable    = 1
  diverterAdd         = 3
  RampDiverterTimer.Enabled   = True
End Sub

Sub RampDiverterTimer_Timer()
  RampDiverter.RotY = RampDiverter.RotY + diverterAdd
  If RampDiverter.RotY <= 0 Or RampDiverter.RotY => 30 Then
    RampDiverterTimer.Enabled = False
  End If
End Sub


' ===============================================================================================
' spinner
' ===============================================================================================

Sub sw41_Spin()
  vpmTimer.PulseSw 41
  PlaysoundAtVol "Gate2_low", sw41, 1
End Sub

Sub sw55_Hit()
  vpmTimer.PulseSw 55
  Playsound "Gate2_low"
End Sub


' ===============================================================================================
' drain post
' ===============================================================================================

Sub SolDrainPost(Enabled)
  If enabled Then
    DrainPost.TransZ = 30
    wallpost.HeightTop = 31
    Wallpost001.collidable=1
    PlaySound SoundFX("SolenoidOn",DOFContactors)
  Else
    DrainPost.TransZ = 0
    wallpost.HeightTop = 1
    Wallpost001.collidable=0
    PlaySound SoundFX("SolenoidOff",DOFContactors)
  end If
End Sub



' ===============================================================================================
' trap door to 8-ball lock
' ===============================================================================================



Sub SolTrapDoor(Enabled)
  If noball8=0 Then
    Ramp139.collidable=1
    ramp029.collidable=0
  Else
    if enabled Then
      Ramp139.collidable=0
      ramp029.collidable=1
      if sap then drp=1
      PlaySound SoundFX("SolenoidOff",DOFContactors)
    Else
      if drp=0 Then
        Ramp139.collidable=1
        ramp029.collidable=0
      End If
    end If
  End If
End Sub


Sub Sol8BLockPlunger(Enabled)
    MultiBallPost.IsDropped = 1
    wall173.collidable= 0
    sw24b.Kick 180,0.1
    sw24b.enabled=false
    sw23b.Kick 180,0.1
    sw23b.enabled=false
    sw22b.Kick 180,0.1
    sw22b.enabled=false
    sw21b.Kick 180,0.1
    sw21b.enabled=false
    sw20b.Kick 180,0.1
    sw20b.enabled=false
    sw19b.Kick 180,0.1
    sw19b.enabled=false
    sw18b.Kick 180,0.1
    sw18b.enabled=false
    sw17b.Kick 180,0.1
    sw17b.enabled=false
    Controller.Switch(17) = 0
    Controller.Switch(18) = 0
    Controller.Switch(19) = 0
    Controller.Switch(20) = 0
    Controller.Switch(21) = 0
    Controller.Switch(22) = 0
    Controller.Switch(23) = 0
    Controller.Switch(24) = 0
    sw24.enabled=0
    sw23.enabled=0
    sw22.enabled=0
    sw21.enabled=0
    sw20.enabled=0
    sw19.enabled=0
    sw18.enabled=0
    sw17.enabled=0
    noball8=1
    noball8b=1
End Sub

' ===============================================================================================
' ramp lift
' ===============================================================================================

Dim rampIsUp, rampCountDown
rampIsUp    = True
rampCountDown = 0

Sub SolRampLift(Enabled)
  if enabled then
    RampLift.HeightBottom = 0
    RampLiftr.HeightBottom = 0
  Else
    RampLift.HeightBottom = 70
    RampLiftr.HeightBottom = 70
  end If
End Sub


' ===============================================================================================
' rocket
' ===============================================================================================
Sub sw43k_Hit()
  Controller.Switch(43)   = True
End Sub
Sub sw43k_Unhit()
  Controller.Switch(43)   = False
End Sub


Sub InitRocket()
  Controller.Switch(50) = True
  Controller.Switch(51) = False
End Sub




Sub SolRocketEject(Enabled)
  If Enabled Then
    sw43k.kick 155+int(rnd(10)), 0.1
    PlaySound SoundFX("SolenoidOn",DOFContactors)
  End If
End Sub

Sub SolRocketLift(Enabled)
  If Enabled Then
    RocketTimer.Interval = 30
    RocketTimer.Enabled  = True
    PlaySound SoundFX("SolenoidOn",DOFContactors)
  Else
    RocketTimer.Interval = 30
    RocketTimer.Enabled  = True
    PlaySound SoundFX("SolenoidOn",DOFContactors)
  End If
End Sub

Sub RocketTimer_Timer()
  Controller.Switch(50) = False
  Controller.Switch(51) = False
  if Rocket.rotx< -16 Then
    RocketTrough.Enabled  = True
  Else
    RocketTrough.Enabled  = 0
  End If
  If controller.Solenoid(20) Then
    If Rocket.RotX <= -18 Then
      PlaySound SoundFX("SolenoidOff",DOFContactors)
      RocketTimer.Enabled   = False
      Controller.Switch(51)   = True
      Exit Sub
    End If
    Rocket.RotX = Rocket.RotX - 0.1
  Else
    If Rocket.RotX >= 0 Then
      PlaySound SoundFX("SolenoidOff",DOFContactors)
      RocketTimer.Enabled   = False
      Controller.Switch(50)   = True
      Exit Sub
    End If
    Rocket.RotX = Rocket.RotX + 0.1
  End If
End Sub


Sub RocketTrough_Hit()
  PlaysoundAtVol "ball_bounce_low", RocketTrough, 1
  me.destroyball
  PlaysoundAtVol "jf_rollingfaster", RocketTrough, 1
  bsTroughSuperVUK.AddBall Me
End Sub

'Sub ShakeRocket(direction)
Sub ShakeRocket (Enabled)
  If Not RocketTimer.Enabled Then
    RocketShakeTimer.Interval = 5+Int(Rnd(5))
    RocketShakeTimer.Enabled  = True
  End If
End Sub

Sub RocketShakeTimer_Timer()
  ' nudging
  Select Case RocketShakeTimer.Interval
  Case 0
    Rocket.RotX = Rocket.RotY - 1
  Case 1, 2
    Rocket.RotX = Rocket.RotX + 1
  Case 3
    Rocket.RotX = Rocket.RotY - 1
  Case 5
    Rocket.RotY = Rocket.RotY - 1
  Case 6, 7
    Rocket.RotY = Rocket.RotY + 1
  Case 8
    Rocket.RotY = Rocket.RotY - 1
  Case 10
    Rocket.RotY = Rocket.RotY + 1
  Case 11, 12
    Rocket.RotY = Rocket.RotY - 1
  Case 13
    Rocket.RotY = Rocket.RotY + 1
  Case Else
    RocketShakeTimer.Enabled = False
  End Select
  RocketShakeTimer.Interval = RocketShakeTimer.Interval + 1
End Sub


' ===============================================================================================
' moon
' ===============================================================================================

Dim MoonBall, radians, angle, radius,angleb, balx ,baly, balyb ,balz, rada, radab

Sub InitMoon()
  ' rotation stuff
  radians         = 3.1415926 / 180
  angle           = 0
  angleb          =180-Moon.objrotz
  radius          = 126

  ' moon is at home
  Controller.Switch(36)   = True
  Controller.Switch(37)   = False
End Sub


Sub SolMoon(Enabled)
  if enabled Then
    if Controller.Switch(36) Then
      mmoon.MagnetOn = 1
      MoonLock.Enabled = 1
    End If
  Else
    mmoon.MagnetOn = 0
    MoonLock.Enabled = 0
    moonballstop.collidable=0
    moonballstopb.collidable=0
    if Controller.Switch(36) Then
      If Controller.Switch(19) = 0 Then
        actvb=0
        set MoonBall = Nothing
      End If
    End If
  End If
End Sub

Dim actvb


Sub MoonLock_Hit()
  If Controller.Switch(36) Then
    moonballstop.collidable=1
    moonballstopb.collidable=1
    actvb=1
    PlaySoundAtVol "collide1", MoonLock, 1
    Set MoonBall  = ActiveBall
    MoonBall.VelX = 0
    MoonBall.VelY = 0
    MoonBall.VelZ = 0
  End If
    mmoon.MagnetOn = 0
    MoonLock.Enabled = 0
End Sub


Sub MoonTrough_Hit()
  me.destroyball
  PlaySoundAtVol "jf_rollingfaster", MoonTrough, 1
  bsTroughSuperVUK.AddBall Me
End Sub

sub SolmoonLift(Enabled)
  if Enabled Then
  MoonTimer.Interval  = 20
  MoonTimer.Enabled   = True
  PlaySound "SolenoidOn"
  End If
End Sub

Sub MoonTimer_Timer()
  mmoon.MagnetOn = 0
  moonballstop.collidable=0
  moonballstopb.collidable=0
  moonlock.enabled=0
  if controller.Solenoid(19) Then
    If Moon.RotX <= -180 Then
        Controller.Switch(37)   = True
      If Controller.Solenoid(34)=0 Then
        set MoonBall = Nothing
        If actvb = 1 Then
          PlaySoundAtVol "ball_bounce_low", MoonLock, 1
          actvb=0
        End If
        Exit Sub
      End If
    Else
      Controller.Switch(36) = False
      Moon.RotX   = Moon.RotX - 1
    End If
  Else
    If Moon.RotX >= Moon.ObjRotX Then
        Controller.Switch(36)   = True
      If Controller.Solenoid(34)=0 Then
        MoonTimer.Enabled     = False
        set MoonBall = Nothing
        PlaySoundAtVol "ball_bounce_low", MoonLock, 1
        actvb=0
        Exit Sub
      End If
    Else
      Controller.Switch(37) = False
      Moon.RotX   = Moon.RotX +2 -actvb
    End If
  End If
      If actvb Then
        angle= abs(moon.rotX)
        rada=radians*angle
        radab=radians*angleb
        MoonBall.VelX = 0
        MoonBall.VelY = 0
        MoonBall.VelZ = 0
        MoonBall.z= sin (rada)*radius+moon.z
        balyb=cos( rada)*radius
        MoonBall.y= balyb*cos(radab)+Moon.y
        MoonBall.x=Sin(radab)*balyb+Moon.x
      End If
End Sub


' ===============================================================================================
' super VUK
' ===============================================================================================

Sub SolTroughSuperVUKOut(Enabled)
  If Enabled Then
    If bsTroughSuperVUK.Balls > 0 Then
      bsTroughSuperVUK.ExitSol_On
    End If
  End If
End Sub



' ===============================================================================================
' top saucer
' ===============================================================================================

Sub sw54k_Hit()
  Controller.Switch(54)   = True
  PlaySoundAtVol "kicker_enter", sw54k, 1

End Sub
Sub sw54k_Unhit()
  Controller.Switch(54)   = False
End Sub

Sub SolTopEject(Enabled)
    If Enabled Then
    sw54k.Kick 220,5+ Int(Rnd()*5)
    PlaySound SoundFX("SolenoidOn",DOFContactors)
  End If
End Sub

' ===============================================================================================
' right saucer
' ===============================================================================================

Sub sw47k_Hit()
  Controller.Switch(47)   = True
  PlaySoundAtVol "kicker_enter", sw47k, 1
End Sub
Sub sw47k_Unhit()
  Controller.Switch(47)   = False
End Sub

Sub SolRightEject(Enabled)
    If Enabled Then
    sw47k.Kick 200,5+ Int(Rnd()*5)
    PlaySound SoundFX("SolenoidOn",DOFContactors)
  End If
End Sub

' ===============================================================================================
' switches
' ===============================================================================================

' top lanes
Sub sw52_Hit()
  controller.Switch(52)=1
End Sub
Sub sw52_Unhit()
  controller.Switch(52)=0
End Sub
Sub sw53_Hit()
  controller.Switch(53)=1
End Sub
Sub sw53_Unhit()
  controller.Switch(53)=0
End Sub

' orbit
Sub sw45_Hit()
  controller.Switch(45)=1
End Sub
Sub sw45_Unhit()
  controller.Switch(45)=0
End Sub
Sub sw46_Hit()
  controller.Switch(46)=1
End Sub
Sub sw46_Unhit()
  controller.Switch(46)=0
End Sub

' return lanes
Sub sw59_Hit()
  controller.Switch(59)=1
End Sub
Sub sw59_Unhit()
  controller.Switch(59)=0
End Sub
Sub sw60_Hit()
  controller.Switch(60)=1
End Sub
Sub sw60_Unhit()
  controller.Switch(60)=0
End Sub

' out lanes
Sub sw57_Hit()
  controller.Switch(57)=1
End Sub
Sub sw57_Unhit()
  controller.Switch(57)=0
End Sub
Sub sw58_Hit()
  controller.Switch(58)=1
End Sub
Sub sw58_Unhit()
  controller.Switch(58)=0
End Sub

' right ramp
Sub sw44_Hit()
  controller.Switch(44)=1
End Sub
Sub sw44_Unhit()
  controller.Switch(44)=0
End Sub

Sub sw56_Hit()
  controller.Switch(56)=1
End Sub
Sub sw56_Unhit()
  controller.Switch(56)=0
End Sub





' ===============================================================================================
' sound effects and some physics for gates, metals, plastics, woods and rubbers
' ===============================================================================================

Dim MetalsSoundActive
Dim WoodsSoundActive

Sub FallingBall_Hit(idx)
  PlaySoundAtVol "ballhit4", FallingBall, 1
End Sub

Sub Gate1_Hit()
  PlaySoundEffect 1, "gate2_low", 8, "gate2", 0, "", 1, 0
End Sub
Sub Gate2_Hit()
  PlaySoundEffect 1, "gate2_low", 8, "gate2", 0, "", 1, 0
End Sub
Sub Gate2Back_Hit()
  PlaySoundEffect 2, "gateback_low", 6, "gateback", 0, "", -1, 0
End Sub

' targets
 Sub Targets_Hit(idx)
  PlaySoundEffect 2, "target_low", 8, "target", 0, "", 0, 0
End Sub

' plastics
Sub Plastics_Hit()
  PlaySoundAtVol "PlasticHit", ActiveBall, 1
End Sub

' rubbers
Sub Rubbers_Hit(idx)
  PlaySoundEffect 1, RandomSoundRubberLowVolume(), 4, RandomSoundRubber(), 14, "bump", 0, 0
End Sub
Sub PostRubbers_Hit(idx)
  PlaySoundEffect 1, RandomSoundRubberLowVolume(), 4, RandomSoundRubber(), 14, "bump", 0, 0
End Sub
Sub MiscRubbers_Hit(idx)
  PlaySoundEffect 1, RandomSoundRubberLowVolume(), 4, RandomSoundRubber(), 14, "bump", 0, 0
End Sub

Function RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : RandomSoundRubber = "rubber_hit_1"
    Case 2 : RandomSoundRubber = "rubber_hit_2"
    Case 3 : RandomSoundRubber = "rubber_hit_3"
  End Select
End Function
Function RandomSoundRubberLowVolume()
  Select Case Int(Rnd*3)+1
    Case 1 : RandomSoundRubberLowVolume = "rubber_hit_1_low"
    Case 2 : RandomSoundRubberLowVolume = "rubber_hit_2_low"
    Case 3 : RandomSoundRubberLowVolume = "rubber_hit_3_low"
  End Select
End Function

' metal rails
Sub MetalRailsStart_Hit(idx)
  PlaySoundAtVol "metalrolling", ActiveBall, 1
End Sub
Sub MetalRailsStop_Hit(idx)
  StopSound "metalrolling"
End Sub

' metal hits
Sub Metals_Hit(idx)
  If MetalsSoundActive Then Exit Sub
  PlaySoundEffect 2, "metalhit_medium_low", 6, "metalhit_medium", 0, "", 0, 0
  MetalsSoundActive   = True
  MetalSoundTimer.Enabled = True
End Sub
Sub MetalSoundTimer_Timer(idx)
  Me.Enabled      = False
  MetalsSoundActive   = False
End Sub

Sub MetalEnds_Hit(idx)
  PlaySoundEffect 1.5, "metalhit", 6, "metalhit", 0, "", 0, 0
End Sub

' woods
Sub Woods_Hit(idx)
  If WoodsSoundActive Then Exit Sub
  PlaySoundEffect 1.5, "woodhit_low", 5, "woodhit", 0, "", 0, 0
  WoodsSoundTimer.Enabled = True
  WoodsSoundActive    = True
End Sub
Sub WoodsSoundTimer_Timer()
  Me.Enabled      = False
  WoodsSoundActive  = False
End Sub


Sub SolFLamp1(Enabled)
  lf1a.state=Enabled
  lf1b.state=Enabled
  lf1c.state=Enabled
End Sub
Sub SolFLamp2(Enabled)
  lf2a.state=Enabled
  lf2b.state=Enabled
  lf2c.state=Enabled
End Sub
Sub SolFLamp3(Enabled)
  lf3a.state=Enabled
  lf3b.state=Enabled
  lf3c.state=Enabled
End Sub
Sub SolFLamp4(Enabled)
  lf4a.state=Enabled
  lf4b.state=Enabled
  lf4c.state=Enabled
  f4.visible = Enabled
End Sub
Sub SolFLamp5(Enabled)
  lf5.state=Enabled
  lf5a.state=Enabled
  lf5b.state=Enabled
  lf5c.state=Enabled
  f5.visible = Enabled
  f5b.visible = Enabled
  End Sub
Sub SolFLamp6(Enabled)
  lf6a.state=Enabled
  lf6b.state=Enabled
  lf6c.state=Enabled
  f6c.visible = Enabled
End Sub
Sub SolFLamp7(Enabled)
  lf7.state= Enabled
  lf7a.state=Enabled
  f7.visible = Enabled
End Sub
Sub SolFLamp8(Enabled)
  lf8a.state= Enabled
  lf8b.state= Enabled
End Sub



' ===============================================================================================
' some general methods
' ===============================================================================================

Function BallSpeed()
  On Error Resume Next
  BallSpeed = SQR(ActiveBall.VelX ^ 2 + ActiveBall.VelY ^ 2)
End Function

Function Sign(number)
  If number < 0 Then
    Sign = -1
  ElseIf number > 0 Then
    Sign = 1
  Else
    Sign = 0
  End If
End Function

Function PlaySoundEffect(minspeed, lowsound, lowspeed, sound, highspeed, highsound, velX, velY)
  Dim currentspeed
  currentspeed = BallSpeed
  If currentspeed >= minspeed Then
    If (velX = 0 Or Sign(ActiveBall.VelX) = Sign(velX)) And (velY = 0 Or Sign(ActiveBall.VelY) = Sign(velY)) Then
      If highspeed > 0 And currentspeed > highspeed Then
        PlaySoundAt highsound, ActiveBall
      ElseIf currentspeed < lowspeed Then
        PlaySoundAt lowsound, ActiveBall
      Else
        PlaySoundAt sound, ActiveBall
      End If
    End If
  End If
End Function

Function IIF(ifcond, iftrue, iffalse)
  If ifcond Then
    IIF = iftrue
  Else
    IIF = iffalse
  End If
End Function



' ===============================================================================================

Set LampCallback = GetRef("Lamps") ' individual lampcallbacks instead of using vpmMapLights (no reason, just because)
Sub Lamps
  pUpdateLED ()
    L1.State = Controller.Lamp(1)
    L2.State = Controller.Lamp(2)
    L3.State = Controller.Lamp(3)
    L4.State = Controller.Lamp(4)
    L5.State = Controller.Lamp(5)
    L6.State = Controller.Lamp(6)
    L7.State = Controller.Lamp(7)
    L8.State = Controller.Lamp(8)
    L9.State = Controller.Lamp(9)
    L10.State = Controller.Lamp(10)
    L11.State = Controller.Lamp(11)
    L12.State = Controller.Lamp(12)
    L13.State = Controller.Lamp(13)
    'L14.State = Controller.Lamp(14)
    L15.State = Controller.Lamp(15)
    L16.State = Controller.Lamp(16)
    L17.State = Controller.Lamp(17)
    L18.State = Controller.Lamp(18)
    L19.State = Controller.Lamp(19)
    L20.State = Controller.Lamp(20)
    L21.State = Controller.Lamp(21)
    L22.State = Controller.Lamp(22)
    L23.State = Controller.Lamp(23)
    L24.State = Controller.Lamp(24)
    L25.State = Controller.Lamp(25)
    L26.State = Controller.Lamp(26)
    L27.State = Controller.Lamp(27)
    L28.State = Controller.Lamp(28)
    L29.State = Controller.Lamp(29)
    L30.State = Controller.Lamp(30)
    L31.State = Controller.Lamp(31)
    L32.State = Controller.Lamp(32)
    L33.State = Controller.Lamp(33)
    L34.State = Controller.Lamp(34)
    L35.State = Controller.Lamp(35)
    L36.State = Controller.Lamp(36)
    L37.State = Controller.Lamp(37)
    L38.State = Controller.Lamp(38)
    L39.State = Controller.Lamp(39)
    L40.State = Controller.Lamp(40)
    L41.State = Controller.Lamp(41)
    L42.State = Controller.Lamp(42)
    L43.State = Controller.Lamp(43)
    L44.State = Controller.Lamp(44)
    L45.State = Controller.Lamp(45)
    L46.State = Controller.Lamp(46)
    L47.State = Controller.Lamp(47)
    L48.State = Controller.Lamp(48)
    L49.State = Controller.Lamp(49)
    f49.visible = Controller.lamp(49)
    L50.State = Controller.Lamp(50)
    L51.State = Controller.Lamp(51)
    L52.State = Controller.Lamp(52)
    L53.State = Controller.Lamp(53)
    L54.State = Controller.Lamp(54)
    L55.State = Controller.Lamp(55)
    L56.State = Controller.Lamp(56)
    L60.State = Controller.Lamp(60)
    L61.State = Controller.Lamp(61)
    L62.State = Controller.Lamp(62)
    l63.State = Controller.Lamp(63)
    f63.visible = Controller.lamp(63)
    L64.State = Controller.Lamp(64)
    f64.visible = Controller.lamp(64)
  If varhidden=1 Then
    f65.visible = Controller.lamp(65)
    f66.visible = Controller.Lamp(66)
    f67.visible = Controller.Lamp(67)
    f68.visible = Controller.Lamp(68)
    f69.visible = Controller.Lamp(69)
    f70.visible = Controller.Lamp(70)
    f71.visible = Controller.Lamp(71)
    f72.visible = Controller.Lamp(72)
  End If
sw41Prim.Rotx = sw41.CurrentAngle-90
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle


End Sub




' ===============================================================================================
' LED's display
' ===============================================================================================

Dim Digits(0), digitState, nummb,mb0, mb1, mb2, mb3, mb4, mb5, mb6, mb7, mb8, mb9, mbl

Digits(0)   = Array(led0,led1,led2,led3,led4,led5,led6)
digitState  = Array(0,0,0,0,0,0,0)

Sub pUpdateLED()
  Dim chgLED, ii, iii, num, chg, stat, digit
  chgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(chgLED) Then
    For ii = 0 To UBound(chgLED)
      num  = chgLED(ii, 0)
      chg  = chgLED(ii, 1)
      stat = chgLED(ii, 2)
      iii  = 0
      For Each digit In Digits(num)
        If Num = 0 Then
          If (chg And 1) Then digit.State = (stat And 1)
          digitState(iii) = digit.State
          chg  = chg \ 2
          stat = stat \ 2
        End If
        iii = iii + 1
      Next
    Next
  End If
  nummb=led0.state+led1.state*2+led2.state*2+led3.state*3+led4.state*4+led5.state+led6.state-3
  mbl=1
  Select Case nummb
       Case 1:mb.image= "pf1 apollo13"
        Case 8:mb.image= "pf2 apollo13"
        Case 6:mb.image= "pf3 apollo13"
        Case 3:mb.image= "pf4 apollo13"
        Case 5:mb.image= "pf5 apollo13"
        Case 9:mb.image= "pf6 apollo13"
        Case 2:mb.image= "pf7 apollo13"
        Case 11:mb.image= "pf8 apollo13"
        Case 7:mb.image= "pf9 apollo13"
        Case 10:mb.image= "pf0 apollo13"
    Case Else  : mbl=0: mb.image="pf00 apollo13"
  End Select
  mb.state=mbl
End Sub






'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep,LStep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw(62)
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    DOF 121, DOFPulse
  RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub


Sub leftSlingShot_Slingshot
  vpmTimer.PulseSw(61)
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
  DOF 122, DOFPulse
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    leftSlingShot.TimerEnabled = 1
End Sub

Sub leftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:leftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

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

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
   BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
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


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 13 ' total number of balls
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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
      'debug.print BOT(b).velz
    End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.

Sub rollingmetal_Hit (idx)
'PlaySound SoundFX("Wire ramp",DOFContactors), 0, 1, 0.05, 0.05
PlaySound "Wire Ramp", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub rollingplastic_Hit (idx)
  PlaySound "fx_ballrolling", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub


Sub Switches_Hit (idx)
  PlaySound "switch_hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub



Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub


Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub



Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub



Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundbumper()
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtBallVol SoundFX("fx_Bumper",DOFContactors), 1
    Case 2 : PlaySoundAtBallVol SoundFX("fx_Bumper1",DOFContactors), 1
    Case 3 : PlaySoundAtBallVol SoundFX("fx_Bumper2",DOFContactors), 1
    Case 4 : PlaySoundAtBallVol SoundFX("fx_bumper3",DOFContactors), 1
    Case 5 : PlaySoundAtBallVol SoundFX("fx_bumper4",DOFContactors), 1
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub
Sub LeftFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub
Sub RightFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub





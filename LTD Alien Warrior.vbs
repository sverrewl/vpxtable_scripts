Option Explicit
Randomize


'*** Backglass effects volume
' Use left and right MagnaSave keys to raise or lower backglass volume

Const cGameName="alienwar"
Const VPMver = "03040000"
Const UseSolenoids    = 1
Const UseLamps      = True
Const UseGI       = False
Const UseSync       = True
Const SCoin       = "Coin"
Const SSolenoidOn     = "SolOn"
Const SSolenoidOff  = "SolOff"
Const SKnocker    = "Knocker"
Const BallSize = 50
Const BallMass = 1.7
Const MaxCheckLoops = 10 'If You are experiencing short GameOver phases during normal gameplay, increase the MaxCheckLoops variable to 15 or 20
Const VPXFastFlips = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

LoadVPM "01560000","LTD.vbs",3.2


' **** VARIABLES ****
Dim TableWidth : TableWidth = Table1.width
Dim TableHeight : TableHeight = Table1.height
Dim Desktop : Desktop = Table1.ShowDT
Dim LUTImage
Dim BgVolume



' ********** Solenoids ************
SolCallback(33) = "SolLFlipper"
SolCallback(34) = "SolRFlipper"
SolCallback(35) = "SolDivert"
'*************** END FLIPPERS ********


SolCallback(1) = "BallTrough"
'SolCallback(6) = "b6_sol"
'SolCallback(7) = "LSLINGSHOT"
SolCallback(8) = "DropReset"
'SolCallback(9) = "RIGHTBUMPER"
'SolCallback(10) = "CENTERBUMPER"
'SolCallback(11) = "LEFTBUMPER"

SolCallback(2) = "CoinDrop"


Dim GameIsActive, InactiveCount, Flipperactive
Flipperactive = False

Sub BallReleaseHelper_Hit
Playsound "BG_BallRelease", 0, BgVolume
  if VPXFastFlips Then
    CheckGameOnTimer.enabled = False
    GameOnTimer.enabled = False
    InactiveCount = 0
    Sol_GameOn(True)
  end If
End Sub

Sub GameOnTimer_Timer
  GameOnTimer.enabled = False
  if VPXFastFlips Then
    GameIsActive = False
    CheckGameOnTimer.enabled = False
    CheckGameOnTimer.enabled = True
    vpmTimer.PulseSw(1)
  end if
End Sub

Sub CheckGameOnTimer_Timer
  if VPXFastFlips Then
    if not GameIsActive Then
      StopSound "BG_Background"
      InactiveCount = InactiveCount + 1
      vpmTimer.PulseSw(84)
      if InactiveCount > MaxCheckLoops Then
        Sol_GameOn(False)
        CheckDrainTimer.enabled = False
        InactiveCount = 0
        CheckGameOnTimer.enabled = False
      end If
    Else
      InactiveCount = 0
      Sol_GameOn(True)
      CheckGameOnTimer.enabled = False
    end If
  end If
End Sub

Dim RunCount
Sub CheckDrainTimer_Timer
  RunCount = RunCount + 1
  GameOnTimer.enabled = False
  GameOnTimer.enabled = True
  if RunCount > 2 then
    CheckDrainTimer.enabled = False
  end if
End Sub

Sub Sol_GameOn(enabled)
dim obj
  VpmNudge.SolGameOn enabled
  Flipperactive = enabled
  if not Flipperactive Then
attract.Enabled=1
StopSound "BG_Background"
SW19.Threshold=100
SW20.Threshold=100
SW21.Threshold=100
RightSlingshot.SlingshotThreshold=100
LeftSlingshot.SlingshotThreshold=100
Else
attract.Enabled=0
Stopsound "BG_Attract"
PlaySound "BG_Background",  -1, BgVolume
SW19.Threshold=2
SW20.Threshold=2
SW21.Threshold=2
RightSlingshot.SlingshotThreshold=2
LeftSlingshot.SlingshotThreshold=2

  end If

For each obj in FlashGI
obj.visible=Flipperactive
next

for each obj in GiLights
 obj.state=Flipperactive
next
end sub

Sub CoinDrop(Enabled)
If Enabled Then
PlaySoundAt "CoinDrop", Plunger
End If
End Sub

Sub attract_timer()
Playsound "BG_Attract"
Me.Enabled=0
End Sub


Dim bsTrough,MyBall

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

   vpmMapLights AllLights


  On Error Resume Next
  With Controller
    .GameName=cGameName
    .Games(cGameName).Settings.value("sound") = 0 ' this rom has broken sound
    .Games(cGameName).Settings.value("dmd_compact") = 1
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine="ALIEN WARRIOR LTD Do BRAZIL 1983"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=0
    .ShowFrame=1
    .ShowTitle=0
    .Hidden=1
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.Run




  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

    vpmNudge.TiltSwitch=18
    vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(SW19,SW20,SW21,LeftSlingshot,RightSlingshot)

  L44.Visible=0
  DRAIN.createball
  DRAIN.kick 270, 2

End Sub




Sub Table1_Exit
SaveBGV
Controller.Games(cGameName).Settings.value("volume") = 0
Controller.Games(cGameName).Settings.value("sound") = 1
Controller.Stop
End Sub

Sub Table1_KeyDown(ByVal keycode)

  If keycode = KeyRules Then RuleCard.visible=1

  if keycode=LockbarKey Then
  NextLUT
  end if

  If Keycode= AddCreditKey then  vpmTimer.AddTimer 750,"vpmTimer.PulseSw 14'" : PlaySound "BG_Coin"
  If keycode = LeftTiltKey Then
Nudge 90, 6
  GameOnTimer.enabled = False
  GameOnTimer.enabled = True
End If
  If keycode = RightTiltKey Then
Nudge 270, 6
  GameOnTimer.enabled = False
  GameOnTimer.enabled = True
End If
  If keycode = CenterTiltKey Then
Nudge 0, 6
  GameOnTimer.enabled = False
  GameOnTimer.enabled = True
End If
  If keycode = PlungerKey Then PlungerPull

' *** VPX Fast Flippers
If Keycode = RightFlipperKey then
        If VPXFastFlips=1 Then
                  if flipperactive then
      RightFlipper.TimerEnabled = True
      RIGHTFLIPPER.EOSTorque = 0.75:RIGHTFLIPPER.RotateToEnd:PlaySoundAt SoundFX("FX_FlipperUpR",DOFFlippers),RIGHTFLIPPER
      PlaySound "buzzR", -1, 0.003, Pan(RIGHTFLIPPER),AudioFade(RIGHTFLIPPER)
          End If
        End If

End If


If Keycode = LeftFlipperKey then
        If VPXFastFlips=1 Then
                  if flipperactive then
        LeftFlipper.TimerEnabled = True
      LEFTFLIPPER.EOSTorque = 0.75:LEFTFLIPPER.RotateToEnd:PlaySoundAt SoundFX("FX_FlipperUpL",DOFFlippers),LEFTFLIPPER
      PlaySound "buzzL", -1, 0.003, Pan(LEFTFLIPPER),AudioFade(LEFTFLIPPER)
          End If
        End If
End If

If vpmKeyDown(KeyCode) Then Exit Sub

End Sub


Sub PlungerPull
  Plunger.Pullback
  PlaySoundAt "plungerpull", Plunger
End Sub


Sub Table1_KeyUp(ByVal keycode)
  If keycode = KeyRules Then RuleCard.visible=0
  If Keycode = StartGameKey then playsound "BG_PlayerUP"
  If keycode = PlungerKey Then PlungerRelease
  if keycode = LockBarKey then

  End If

If Keycode = LeftMagnaSave then
  If BgVolume > 0 Then
    BgVolume = (BgVolume - 0.05)
    Playsound "BG_VolumeDown", 0, BgVolume
  End If
End If

If Keycode = RightMagnaSave then
  If BgVolume < 1 Then
    BgVolume = (BgVolume + 0.05)
    Playsound "BG_VolumeUp", 0, BgVolume
  End If
End If

If Keycode = RightFlipperKey then

        If VPXFastFlips = 1 Then
                  if flipperactive then
    RIGHTFLIPPER.EOSTorque = 0.1
    RIGHTFLIPPER.RotateToStart:PlaySoundAt SoundFX("FX_FlipperDown",DOFFlippers),RIGHTFLIPPER
    Stopsound "buzzR"
        Else
    Stopsound "buzzR"
    RIGHTFLIPPER.EOSTorque = 0.1
    RIGHTFLIPPER.RotateToStart
      End If
  End If
End If


If Keycode = LeftFlipperKey then
        If VPXFastFlips = 1 Then
                  if flipperactive then
    LEFTFLIPPER.EOSTorque = 0.1
    LEFTFLIPPER.RotateToStart:PlaySoundAt SoundFX("FX_FlipperDown",DOFFlippers),LEFTFLIPPER
    Stopsound "buzzL"
      Else
    LEFTFLIPPER.EOSTorque = 0.1
    LEFTFLIPPER.RotateToStart
    Stopsound "buzzL"
      End If
  End If
End If



  If vpmKeyUp(KeyCode) Then Exit Sub

End Sub

Sub PlungerRelease
  Plunger.Fire
  PlaySoundAt "plunger", Plunger
End Sub

Sub SolDivert(Enabled)
If Enabled Then
Playsound "BG_Diverter", 0, BgVolume
Diverter.RotateToEnd
L44.visible=1
PlaySoundAt SoundFX("FX_Diverter",DOFFlippers), Diverter
Else
Diverter.RotateToStart
L44.Visible=0
PlaySoundAt SoundFX("FX_Diverter",DOFFlippers), Diverter
End If
End Sub


'*** Ball Trough
Sub DRAIN_HIT             'SWITCH 6
playsound "BG_DrainHit", 0, BgVolume
PlaySoundAt "fx_hole_enter", DRAIN
Controller.Switch(6)=1
  if VPXFastFlips then
    GameOnTimer.enabled = False
    GameOnTimer.enabled = True
    CheckDrainTimer.enabled = False
    CheckDrainTimer.enabled = True
    RunCount = 0
  end if
End Sub
Sub DRAIN_unhit : Controller.Switch(6)=0 : End Sub

Sub BallTrough(Enabled)
If Enabled Then
DRAIN.kick 65, 15
PlaySoundAt "Solenoid", DRAIN
PlaysoundAt SoundFX("fx_kicker", DOFContactors), BallReleaseHelper
End If
End Sub
'***


' *** in/outlanes
Sub SW43_HIT:Controller.Switch(43)=1
If controller.lamp(12) Then
Playsound "BG_OutlaneBonus", 0, BgVolume
Else
Playsound "BG_LeftOutlane", 0, BgVolume
End If              'switch 43
PlaySoundAt "switch1", SW43
End Sub
Sub SW43_UNHIT:Controller.Switch(43)=0:End Sub            'switch 43

Sub SW42_HIT:Controller.Switch(42)=1              'switch 42
PlaySoundAt "switch1", SW42
End Sub
Sub SW42_UNHIT:Controller.Switch(42)=0:End Sub            'switch 42

Sub SW41_HIT:Controller.Switch(41)=1
If controller.lamp(11) Then
Playsound "BG_OutlaneBonus", 0, BgVolume
Else
Playsound "BG_LeftOutlane", 0, BgVolume
End If                    'switch 41
PlaySoundAt "switch1", SW41
End Sub
Sub SW41_UNHIT:Controller.Switch(41)=0:End Sub            'switch 41
' ***

'*** side lanes
Sub SW49_HIT:Controller.Switch(49)=1                'switch 49
Playsound "BG_LeftTunnel", 0, BgVolume
PlaySoundAt "switch1", SW49
End Sub
Sub SW49_UNHIT:Controller.Switch(49)=0:End Sub              'switch 49

Sub SW50_HIT:Controller.Switch(50)=1                'switch 50
If controller.lamp(27) Then
Playsound "BG_LeftTunnelBonus", 0, BgVolume
Else
Playsound "BG_LeftTunnel", 0, BgVolume
End If
PlaySoundAt "switch1", SW50
End Sub
Sub SW50_UNHIT:Controller.Switch(50)=0:End Sub              'switch 50
'***




'*** Spinner
Sub SW38_Spin
vpmTimer.PulseSw 38
Playsound "BG_Spinner", 0, BgVolume
PlaysoundAt "spinnerclicking", SW38
End Sub
'***              'switch 31

'*** top lanes
Sub SW44_Hit:Controller.Switch(44)=1                'switch 44
playsound "BG_ToplaneA", 0, BgVolume
PlaySoundAt "switch1", SW44
End Sub
Sub SW44_unHit:Controller.Switch(44)=0:End Sub              'switch 44

Sub SW45_Hit:Controller.Switch(45)=1              'switch 45
playsound "BG_ToplaneB", 0, BgVolume
PlaySoundAt "switch1", SW45
End Sub
Sub SW45_unHit:Controller.Switch(45)=0:End Sub              'switch 45

Sub SW46_Hit:Controller.Switch(46)=1              'switch 46
playsound "BG_ToplaneC", 0, BgVolume
PlaySoundAt "switch1", SW46
End Sub
Sub SW46_unHit:Controller.Switch(46)=0:End Sub              'switch 46
'***


'*** Bumpers
Sub SW19_Hit      'switch 19
vpmTimer.PulseSw 19
Playsound "BG_Bumper1", 0, BgVolume
PlaySoundAt SoundFX("fx_bumper1", DOFContactors), SW19
End Sub

Sub SW20_Hit          'switch 20
vpmTimer.PulseSw 20
Playsound "BG_Bumper2", 0, BgVolume
PlaySoundAt SoundFX("fx_bumper3", DOFContactors), SW20
End Sub

Sub SW21_Hit      'switch 21
vpmTimer.PulseSw 21
Playsound "BG_Bumper3", 0, BgVolume
PlaySoundAt SoundFX("fx_bumper4", DOFContactors), SW21
End Sub
'***

'*** Hit Targets
Sub SW34_Hit : Playsound "BG_HT1", 0, BgVolume : PlaysoundAt "FX_HitTarget",SW34 :vpmTimer.PulseSw 34:End Sub
Sub SW35_Hit : Playsound "BG_HT2", 0, BgVolume : PlaysoundAt "FX_HitTarget",SW35 :vpmTimer.PulseSw 35:End Sub
Sub SW36_Hit : Playsound "BG_HT3", 0, BgVolume : PlaysoundAt "FX_HitTarget",SW36 :vpmTimer.PulseSw 36:End Sub
'***



'*** Hit target in drop target chain
Sub SW33_Hit
Playsound "BG_HTSpecial", 0, BgVolume
PlaysoundAt "FX_HitTarget",SW33
vpmtimer.pulsesw 33
End Sub
'***

'*** Drop Targets
Sub SW27_DROPPED
Controller.Switch(27)=1
playsoundat "fx_DropTargetDown", SW27
playsound "BG_DT1", 0, BgVolume
End Sub

Sub SW28_DROPPED
Controller.Switch(28)=1
playsoundat "fx_DropTargetDown", SW28
playsound "BG_DT2", 0, BgVolume
End Sub

Sub SW29_DROPPED
Controller.Switch(29)=1
playsoundat "fx_DropTargetDown", SW29
playsound "BG_DT3", 0, BgVolume
End Sub


Sub SW30_DROPPED
Controller.Switch(30)=1
playsoundat "fx_DropTargetDown", SW30
playsound "BG_DT4", 0, BgVolume
End Sub
'***


'*** Drop Target Reset Solenoid
Sub DropReset(Enabled)
If Enabled Then
PlaySoundAt SoundFX("FX_DropTargetUp", DOFContactors), SW33
SW27.IsDropped=0:Controller.Switch(27)=0
SW28.IsDropped=0:Controller.Switch(28)=0
SW29.IsDropped=0:Controller.Switch(29)=0
SW30.IsDropped=0:Controller.Switch(30)=0
End If
End Sub
'***

'*** Rebound Targets
Sub SW22a_Hit : playsound "BG_ReboundA", 0, BgVolume :VpmTimer.PulseSw 22: playsoundat "fx_rubber_band", S62a : End Sub     'switch 22
Sub SW22b_Hit : playsound "BG_ReboundB", 0, BgVolume :VpmTimer.PulseSw 22: playsoundat "fx_rubber_band", S62b : End Sub     'switch 22
Sub SW22c_Hit : playsound "BG_ReboundC", 0, BgVolume :VpmTimer.PulseSw 22: playsoundat "fx_rubber_band", S62c : End Sub     'switch 22
Sub SW22d_Hit : playsound "BG_ReboundD", 0, BgVolume :VpmTimer.PulseSw 22: playsoundat "fx_rubber_band", S62d : End Sub     'switch 22
Sub SW22e_Hit : playsound "BG_ReboundE", 0, BgVolume :VpmTimer.PulseSw 22: playsoundat "fx_rubber_band", S62e : End Sub     'switch 22
Sub SW22f_Hit : playsound "BG_ReboundA", 0, BgVolume :VpmTimer.PulseSw 22: playsoundat "fx_rubber_band", S62f : End Sub     'switch 22
'***

Dim RStep, LStep


Sub LeftSlingshot_Slingshot   'switch 26
playsound"BG_LSlingshot", 0, BgVolume
PlaySoundAt SoundFX("FX_Slingshot", DOFContactors), Lsling
VpmTimer.PulseSw 26
    LSling1.Visible = 0
    LSling2.Visible = 1
    lsling.rotx = 10
    LStep = 0
    LeftSlingshot.TimerEnabled = 1
End Sub

Sub LeftSlingshot_Timer
    Select Case LStep
        Case 1:LSLing2.Visible = 0:LSLing3.Visible = 1:lsling.rotx = 20
        Case 2:LSLing3.Visible = 0:LSLing2.Visible = 1:lsling.rotx = 10
        Case 3:LSLing2.Visible = 0:LSLing1.Visible = 1:lsling.rotx = 0:LeftSlingshot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub RightSlingshot_Slingshot  'switch 25
playsound "BG_RSlingshot", 0, BgVolume
PlaySoundAt SoundFX("FX_Slingshot", DOFContactors), Rsling
VpmTimer.PulseSw 25
    RSling1.Visible = 0
    RSling2.Visible = 1
    Rsling.rotx = 10
    RStep = 0
    RightSlingshot.TimerEnabled = 1
End Sub

Sub RightSlingshot_Timer
    Select Case RStep
        Case 1:RSLing2.Visible = 0:RSLing3.Visible = 1:Rsling.rotx = 20
        Case 2:RSLing3.Visible = 0:RSLing2.Visible = 1:Rsling.rotx = 10
        Case 3:RSLing2.Visible = 0:RSLing1.Visible = 1:Rsling.rotx = 0:RightSlingshot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub



Sub Gate3_hit ' - SW C15 PinMAME 37
PlaysoundAT "fx_gate", Gate3
vpmtimer.pulseSW 37
End Sub

Sub GateTimer_Timer()
   Gate2Flap.RotZ = ABS(Gate2.currentangle)
   Gate3Flap.RotZ = ABS(Gate1.currentangle)
End Sub




'*******************
' Flipper Subs v3.0
'*******************


Sub SolLFlipper(Enabled)
  If Enabled Then
If VPXFastFlips = 1 Then
      GameIsActive = True

  End If
End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
If VPXFastFlips = 1 Then
      GameIsActive = True

  End If
End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub





'************************************
' Diverse Collection Hit Sounds v3.0
'************************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubberbandShort_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubberbandLong_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aPosts_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aPlastic_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


sub g1rebound_hit : PlaysoundAT "fx_gate_rebound", g1rebound : end Sub
sub g2rebound_hit : PlaysoundAT "fx_gate_rebound", g2rebound : end Sub
Sub Gate1_hit : PlaysoundAT "fx_gate_open", Gate1 : end sub
Sub Gate2_hit : PlaysoundAT "fx_gate_open", Gate2 : end sub



'***************************************************************
'             Supporting Ball & Sound Functions v3.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************



Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of a fast object, like bumpers, flippers and other solenoids
    PlaySound soundname, 0, 1, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.4, 0, 0, 0, AudioFade(ActiveBall)
End Sub






'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v3.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 1   'total number of balls, 20 balls, from 0 to 19
Const lob = 0     'number of locked balls
Const maxvel = 25 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate_Timer()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls and hide the shadow
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
        aBallShadow(b).Y = 3000
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y

        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 10
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed control
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next

'** Wylte's shadow stretcher
    Dim LightStrength:    LightStrength    = 12             'Multiplier affecting how quickly the shadow moves away as the ball approaches a light source
    Dim LimD:             LimD            = 25            'The limit for how far the shadow can get from the ball (>25 is not touching)
    Dim RtxBScnt:        RtxBScnt        = 10            'Number of lights in your RtxBS collection
    Dim fovY:            fovY            = 2                'Offset y axis to account for layback or inclination

    Dim BWS, s, Source, LSx, LSy, LSd, denom, CntDwn
    BWS = GetBalls
    If UBound(BWS) = lob - 1 Then Exit Sub                'No balls in play exit, same as JP's

    For s = lob to UBound(BWS)
        For Each Source in GiLights                            'Rename this to match your collection name
            LSx=BWS(s).x-Source.x: LSy=BWS(s).y-Source.y: LSd=(LSx^2+LSy^2)^0.5            'Calculating the Linear distance to the Source
            If LSd < Source.falloff Then                                                'If the ball is within the falloff range of a light
                CntDwn = RtxBScnt
                denom = 100000/(Source.falloff-LSd)

                aMovingShadow(s).X=BWS(s).X+Source.intensity*LightStrength*(BWS(s).x-Source.x)/denom
                If aMovingShadow(s).X-BWS(s).X>LimD Then aMovingShadow(s).X=BWS(s).X+LimD
                If aMovingShadow(s).X-BWS(s).X<-LimD Then aMovingShadow(s).X=BWS(s).X-LimD

                aMovingShadow(s).Y=BWS(s).Y-fovY+Source.intensity*LightStrength*(BWS(s).y-Source.y)/denom
                If aMovingShadow(s).Y-BWS(s).Y>LimD-fovY Then aMovingShadow(s).Y=BWS(s).Y+LimD-fovY
                If aMovingShadow(s).Y-BWS(s).Y<-LimD-fovY Then aMovingShadow(s).Y=BWS(s).Y-LimD-fovY

                aMovingShadow(s).height=BWS(s).Z-24
            Else
                CntDwn = CntDwn - 1                'If ball is not in range of any light, this will hit 0
            End If
        Next
        If CntDwn <= 0 Then
            aMovingShadow(s).X=BWS(s).X            'Center shadow on ball when CntDwn hits 0
            aMovingShadow(s).Y=BWS(s).Y-fovY
            aMovingShadow(s).height=BWS(s).Z-24
        End If
    Next

End Sub



    LoadLut

Sub LoadLUT
Dim x
  x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 20: UpdateLUT: SaveLUT: End Sub



Sub UpdateLUT
Select Case LutImage
Case 0: Table1.ColorGradeImage = "_ColorGrade_0"
Case 1: Table1.ColorGradeImage = "_ColorGrade_1"
Case 2: Table1.ColorGradeImage = "_ColorGrade_2"
Case 3: Table1.ColorGradeImage = "_ColorGrade_3"
Case 4: Table1.ColorGradeImage = "_ColorGrade_4"
Case 5: Table1.ColorGradeImage = "_ColorGrade_5"
Case 6: Table1.ColorGradeImage = "_ColorGrade_6"
Case 7: Table1.ColorGradeImage = "_ColorGrade_7"
Case 8: Table1.ColorGradeImage = "_ColorGrade_8"
Case 9: Table1.ColorGradeImage = "_ColorGrade_9"
Case 10: Table1.ColorGradeImage = "_ColorGrade_9_Lite"
Case 11: Table1.ColorGradeImage = "_ColorGrade_8_Lite"
Case 12: Table1.ColorGradeImage = "_ColorGrade_7_Lite"
Case 13: Table1.ColorGradeImage = "_ColorGrade_6_Lite"
Case 14: Table1.ColorGradeImage = "_ColorGrade_5_Lite"
Case 15: Table1.ColorGradeImage = "_ColorGrade_4_Lite"
Case 16: Table1.ColorGradeImage = "_ColorGrade_3_Lite"
Case 17: Table1.ColorGradeImage = "_ColorGrade_2_Lite"
Case 18: Table1.ColorGradeImage = "_ColorGrade_1_Lite"
Case 19: Table1.ColorGradeImage = "_ColorGrade_0_Lite"
End Select
End Sub

'*** Flipper shadows
Sub LeftFlipper_Init()
    LeftFlipper.TimerInterval = 10
End Sub

Sub RightFlipper_Init()
    RightFlipper.TimerInterval = 10
End Sub

Sub LeftFlipper_Timer()
    FlipperLSh.RotZ = LeftFlipper.CurrentAngle
    If LeftFlipper.CurrentAngle = LeftFlipper.StartAngle Then
        LeftFlipper.TimerEnabled = False
    End If
End Sub

Sub RightFlipper_Timer()
    FlipperRSh.RotZ = RightFlipper.CurrentAngle
    If RightFlipper.CurrentAngle = RightFlipper.StartAngle Then
        RightFlipper.TimerEnabled = False
    End If
End Sub
'***

'*** Load and Savee BG Volume settings
    LoadBGV

Sub LoadBGV
Dim x
  x = LoadValue(cGameName, "BgVolume")
    If(x <> "") Then BgVolume = x Else BgVolume = 1
  UpdateLUT
End Sub

Sub SaveBGV
    SaveValue cGameName, "BgVolume", BgVolume
End Sub
'***





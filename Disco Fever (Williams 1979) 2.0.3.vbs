'************************************************************
'************************************************************
'
'  Disco Fever / IPD No. 686  / June, 1979 / 4 Players
'
'
'************************************************************
'************************************************************
'
Option Explicit
Randomize

'******************************************************
'             OPTIONS
'******************************************************
Dim Flippertype
dim B2SOff, form1off, controller, desktopmode


FlipperType = 2           ' 1 for normal, 2 for banana (use left magnasave before game starts)
Const VolumeDial = 10       ' Change volume of hit events
Const RollingSoundFactor = 3    ' Change volume of rolling sounds
Const VRRoom = false        ' set to false to show desktop/cabinet
form1off=0              ' 1 will stops the b2s being displayed. set to 0 to show directb2s
B2Soff = 0              ' set to 1 to turn off and 0 to display b2s
desktopmode = 0           ' set to 1 for desktop mode and 0 for cabinet

'example1: Cabinetmode, directb2s On: VRRoom = false , form1off=0 , B2Soff = 0, desktopmode = 0 (backglass check off "Test Desktop", mode fullscreen)
'example2: Desktopmode, directb2s Off: VRRoom = false , form1off=1 , B2Soff = 1, desktopmode = 1 (backglass check on "Test Desktop", mode desktop)

'********************************************** Volume Chimes

Const VolChimes = 0.5

'******************************************************
'           STANDARD DEFINITIONS
'******************************************************

Dim BallMass ,BallSize
'Ballsize = 50
Ballmass = 1.0

Const UseSolenoids=2
Const UseLamps=1
Const UseSync=1
Const UseGI=0


'Dim DesktopMode: DesktopMode = discofever.ShowDT

' Standard Sounds

Const SSolenoidOn = "SolOn"       'Solenoid activates
Const SSolenoidOff  = "SolOff"      'Solenoid deactivates
Const SCoin     = "CoinIn"      'Coin inserted
Const SKnocker    = "Knocker"
Dim object

'********************  Flasher position  *******************

Sub SetBackglass()



For Each object In Backglass_flashers

object.x = object.x - 10

object.height = - object.y - 120

object.y = -40   'adjusts the distance from the backglass towards the user

Next

For Each object In LEDS

object.x = object.x - 10

'object.height =  object.height

object.y = -40   'adjusts the distance from the backglass towards the user

Next


Go1.y = go1.y +3
hs1.y = hs1.y +3
tilt1.y = tilt1.y +3
BIP1.y = BIP1.y +3
SPSA.y = SPSA.y +3
match1.y = match1.y +3

'for each Object in Backglass_flashers : object.visible = 1 : next
for each Object in leds : object.visible = 0 : next
for each Object in Backglass_bulbs : object.visible = 1 : next

'Backglass.visible=0

End Sub


'******************************************************
'           TABLE INIT
'******************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01300000","S6.VBS",3.1
Const cGameName = "disco_l1"

if form1off=1 Then
  Set Controller = CreateObject("b2s.server")
  controller.launchbackglass=0
  b2soff=1
end if

If B2Soff=0 then Set Controller = CreateObject("b2s.server") end if
If B2Soff=1 then Set Controller = CreateObject("VPinMAME.Controller") end if

Dim dt3Target, dt2Target
Dim bsTrough, BallInPlay, oRing, xx

Sub discofever_Init

  SetBackglass

  vpmInit Me

  With Controller
    .GameName="disco_l1"
    .SplashInfoLine="Disco Fever, Williams 1979"
    .HandleKeyboard=0
        .HandleMechanics = 0
    .ShowTitle=0
    .ShowDMDOnly=0
    .ShowFrame=0
    .Hidden = 1

    On Error Resume Next
    .SolMask(0) = 0
    vpmTimer.AddTimer 1000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 1 seconds
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
    End With


  vpmMapLights CPULights
  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = 1

  vpmNudge.TiltSwitch = 30
  vpmNudge.Sensitivity = 1
  vpmNudge.TiltObj=Array(LeftSling,RightSling, Bumper19, Bumper20)

  Set bsTrough=New cvpmBallStack
  With bsTrough
        .InitSw 0, 22, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
'   .InitNoTrough BallRelease,9,75,5
        .InitEntrySnd "ballrel","solon"
    .InitExitSnd "ballrel","solon"
        .Balls = 1
  End With

    ' 3 Drop targets
  Set dt3Target=New cvpmDropTarget
  With dt3Target
    .InitDrop Array(sw11,sw20,sw29),Array(11,20,29)
    .InitSnd SoundFX("DropTarget",DOFContactors),SoundFX("FiveBankDropTargetReset",DOFContactors)
    .AllDownSw=18
    .CreateEvents "dt3Target"
  End With

    ' 2 Drop targets
  Set dt2Target=New cvpmDropTarget
    With dt2Target
    .InitDrop Array(sw38,sw39),Array(38,39)
    .InitSnd SoundFX("DropTarget",DOFContactors),SoundFX("ThreeBankTargetReset",DOFContactors)
    .AllDownSw=19
        .CreateEvents "dt2Target"
    End With




  If FlipperType = 1 Then
    LeftFlipper.enabled = 1
    RightFlipper.enabled = 1
    LeftFlipper.visible = 1
    RightFlipper.visible = 1

    LeftFlipper1.enabled = 0
    RightFlipper1.enabled = 0
    LeftFlipper2.enabled = 0
    RightFlipper2.enabled = 0
    LeftFlipper3.enabled = 0
    RightFlipper3.enabled = 0

    lflip.visible = 0
    rflip.visible = 0

    notCollidable
  Else
    LeftFlipper.enabled = 0
    RightFlipper.enabled = 0
    LeftFlipper.visible = 0
    RightFlipper.visible = 0

    LeftFlipper1.enabled = 1
    RightFlipper1.enabled = 1
    LeftFlipper2.enabled = 1
    RightFlipper2.enabled = 1
    LeftFlipper3.enabled = 1
    RightFlipper3.enabled = 1

    batleftshadow.imagea = "flippers_shadowb1"
    batrightshadow.imagea= "flippers_shadowb2"

    lflip.visible = 1
    rflip.visible = 1
  End If
End Sub

If VRroom then
    For each xx in DT:xx.Visible = false: Next
  For each xx in VRstuff:xx.Visible = true: Next
  For each xx in CPUlights:xx.visible = false: Next
end If

If VRroom=false and desktopmode=1 then
    For each xx in DT:xx.Visible = true: Next
  For each xx in VRstuff:xx.Visible = false: Next
  For each xx in CPUlights:xx.visible = true: Next
  for each xx in Backglass_flashers : xx.visible = 0 : next
  Licht.enabled=0
End If

If VRroom=false and desktopmode=0 then
    For each xx in DT:xx.Visible = false: Next
  For each xx in VRstuff:xx.Visible = false: Next
  For each xx in CPUlights:xx.visible = false: Next
  for each xx in Backglass_flashers : xx.visible = 0 : next
  Licht.enabled=0
End If

Sub discofever_Exit():Controller.Stop:End Sub
Sub discofever_Paused:Controller.Pause = 1:End Sub
Sub discofever_unPaused:Controller.Pause = 0:End Sub

TimerVRPlunger2.enabled = true   '  This sits outside of a sub, and tells the timer2 to be enabled at table load

Sub TimerVRPlunger_Timer
if VRPlunger.Y < 2180 then VRPlunger.Y = VRPlunger.y +5  'If the plunger is not fully extend it, then extend it by 5 coordinates in the Y,
End Sub

Sub TimerVRPlunger2_Timer
VRPlunger.Y = 2120 + (5* Plunger.Position) -20  ' This follows our dummy plunger position for analog plunger hardware users.
end sub




'******************************************************
'             KEYS
'******************************************************

Sub discofever_KeyDown(ByVal KeyCode)
  If keycode = PlungerKey Then
    Plunger.Pullback
    PlaySound "fx_PlungerPull"
        TimerVRPlunger.enabled = true  ' We enable the plunger timer below..  look for the TimerVPlunger Sub.
        TimerVRPlunger2.enabled = False' We disaable the plunger2 timer below..  look for the TimerVPlunger2 Sub.
  end if

  If vpmKeyDown(KeyCode) Then Exit Sub
  If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("nudge_left",0)
  If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("nudge_right",0)
  If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("nudge_forward",0)


  if keycode=leftmagnasave and Controller.Lamp(62) Then
    If flippertype = 2 then flippertype = 1 else flippertype = 2 end if

    If FlipperType = 1 Then
      LeftFlipper.enabled = 1
      RightFlipper.enabled = 1
      LeftFlipper.visible = 1
      RightFlipper.visible = 1

      LeftFlipper1.enabled = 0
      RightFlipper1.enabled = 0
      LeftFlipper2.enabled = 0
      RightFlipper2.enabled = 0
      LeftFlipper3.enabled = 0
      RightFlipper3.enabled = 0

      lflip.visible = 0
      rflip.visible = 0

      notCollidable
    Else
      LeftFlipper.enabled = 0
      RightFlipper.enabled = 0
      LeftFlipper.visible = 0
      RightFlipper.visible = 0

      LeftFlipper1.enabled = 1
      RightFlipper1.enabled = 1
      LeftFlipper2.enabled = 1
      RightFlipper2.enabled = 1
      LeftFlipper3.enabled = 1
      RightFlipper3.enabled = 1

      batleftshadow.imagea = "flippers_shadowb1"
      batrightshadow.imagea= "flippers_shadowb2"

      lflip.visible = 1
      rflip.visible = 1
    End If



  end if
End Sub

Sub discofever_KeyUp(ByVal KeyCode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "plunger"
    TimerVRPlunger.enabled = false  'Disabling the plungertimer.. this is used to animate the plunger
    TimerVRPlunger2.enabled = true  'enabling the plungertimer.. this is used to animate the plunger
    VRPlunger.Y = 2120 ' Putting the plunger back to start position when we let go of it with a button. change number to the plunger position
  end if

  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'******************************************************
'             SOLENOIDS
'******************************************************

SolCallback(1)  = "bsTrough.SolOut"
SolCallback(3) = "dt2Target.SolDropUp"
SolCallback(2) = "dt3Target.SolDropUp"
SolCallback(14)  = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"


Solcallback(9) = "Chime10Sound"
Solcallback(10) = "Chime100Sound"
Solcallback(11) = "Chime1000Sound"
Solcallback(12) = "Chime10000Sound"
' Solcallback(12) =  sound alternator???

Solcallback(15) = "BuzzerSound"

Sub Chime10Sound(Enabled)
  If Enabled Then
  PlaySound SoundFX("NoNota1a",DOFChimes),0,VolChimes,0,0,0,0,0
  End If
End Sub

Sub Chime100Sound(Enabled)
  If Enabled Then
  PlaySound SoundFX("NoNota3a",DOFChimes),0,VolChimes,0,0,0,0,0
  End If
End Sub

Sub Chime1000Sound(Enabled)
  If Enabled Then
  PlaySound SoundFX("NoNota",DOFChimes),0,VolChimes,0,0,0,0,0
  End If
End Sub

Sub Chime10000Sound(Enabled)
  If Enabled Then
  PlaySound SoundFX("NoNota2a",DOFChimes),0,VolChimes,0,0,-1,0,0
  End If
End Sub

Sub BuzzerSound(Enabled)
  If Enabled Then
  PlaySound "Buzzer",0,VolChimes,0,0,-1,0,0
  End If
End Sub

'******************************************************
'         FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

 Sub SolLFlipper(Enabled)
 If Enabled Then
   PlaySoundAt SoundFX("LeftFlipperUp",DOFContactors), LeftFlipper
    If FlipperType = 1 Then
      LF.Fire 'LeftFlipper.RotateToEnd
    Else
      LeftFlipper1.RotateToEnd
      LeftFlipper2.RotateToEnd
      LeftFlipper3.RotateToEnd
    End If
   Else
    PlaySoundAt SoundFX("LeftFlipperDown",DOFContactors), LeftFlipper
    If FlipperType = 1 Then
      LeftFlipper.RotateToStart
    Else
      LeftFlipper1.RotateToStart
      LeftFlipper2.RotateToStart
      LeftFlipper3.RotateToStart
    End If
   End If
 End Sub

 Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("RightFlipperUp",DOFContactors), RightFlipper
   If FlipperType = 1 Then
      RF.Fire 'RightFlipper.RotateToEnd
    Else
      RightFlipper1.RotateToEnd
      RightFlipper2.RotateToEnd
      RightFlipper3.RotateToEnd
    End If
   Else
    PlaySoundAt SoundFX("RightFlipperDown",DOFContactors), RightFlipper
   If FlipperType = 1 Then
      RightFlipper.RotateToStart
    Else
      RightFlipper1.RotateToStart
      RightFlipper2.RotateToStart
      RightFlipper3.RotateToStart
    End If
   End If
 End Sub
 
Sub LeftFlipper_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub RightFlipper_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub LeftFlipper1_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub RightFlipper1_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub LeftFlipper2_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub RightFlipper2_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub LeftFlipper3_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub RightFlipper3_Collide(parm)
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub notCollidable()
  Dim xx
  for each xx in bananaL: xx.collidable = 0: Next
  for each xx in bananaR: xx.collidable = 0: Next
End Sub

'******************************************************
'         SLINGSHOTS
'******************************************************

Dim LStep, RStep

Sub LeftSling_Slingshot
  PlaySoundAt SoundFx("LeftSlingshot",DOFContactors), Sling2
  LSling.Visible = 0
  LSling1.Visible = 1
  Sling2.TransY = -18
  LStep = 0
  vpmTimer.PulseSw 23
  LeftSling.TimerEnabled = 1
End Sub

Sub LeftSling_Timer
  Select Case LStep
    Case 1: LSLing1.Visible = 0
        LSLing2.Visible = 1
        Sling2.TransY = -9
    Case 2: LSLing2.Visible = 0
        LSLing.Visible = 1
        Sling2.TransY = 0
        LeftSling.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub

Sub RightSling_Slingshot
  PlaySoundAt SoundFx("RightSlingshot",DOFContactors), Sling1
  RSling.Visible = 0
  RSling1.Visible = 1
  Sling1.TransY = -18
  RStep = 0
  vpmTimer.PulseSw 21
  RightSling.TimerEnabled = 1
End Sub

Sub RightSling_Timer
  Select Case RStep
    Case 1: RSLing1.Visible = 0
        RSLing2.Visible = 1
        Sling1.TransY  = -9
    Case 2: RSLing2.Visible = 0
        RSLing.Visible = 1
        Sling1.TransY = 0
        RightSling.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub Drain_Hit:PlaySoundAtBall "Drain":bsTrough.AddBall Me:End Sub

Sub BallRelease_UnHit:Set BallInPlay=ActiveBall:PlaySoundAtBall "BallEject":End Sub

'******************************************************
'     SWITCHES, KICKERS, & BUMPERS
'******************************************************




Sub Bumper19_Hit
  vpmTimer.PulseSwitch 33, 0, 0
  PlaySoundAT SoundFX("Bumper19",DOFContactors), Bumper19
End Sub

Sub Bumper20_Hit
  vpmTimer.PulseSwitch 17, 0, 0
  PlaySoundAt SoundFX("Bumper20",DOFContactors), Bumper20
End Sub



' Rollover Switches
Sub sw10_Hit:Controller.Switch(10)=1:PlaySoundAt "fx_sensor", sw10:End Sub
Sub sw10_UnHit:Controller.Switch(10)=0:End Sub

Sub sw27_Hit:Controller.Switch(27)=1:PlaySoundAt "fx_sensor", sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27)=0:End Sub

Sub sw26_Hit:Controller.Switch(26)=1:PlaySoundAt "fx_sensor", sw26:End Sub
Sub sw26_UnHit:Controller.Switch(26)=0:End Sub

Sub sw25_Hit:Controller.Switch(25)=1:PlaySoundAt "fx_sensor", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25)=0:End Sub

Sub sw19_Hit:Controller.Switch(19)=1:PlaySoundAt "fx_sensor", sw19:End Sub
Sub sw19_UnHit:Controller.Switch(19)=0:End Sub

Sub sw24_Hit:Controller.Switch(24)=1:PlaySoundAt "fx_sensor", sw24:End Sub
Sub sw24_UnHit:Controller.Switch(24)=0:End Sub

Sub sw28_hit:Controller.Switch(28)=1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28)=0:End Sub

Sub sw35_hit:Controller.Switch(35)=1:PlaySoundAt "fx_sensor", sw35:End Sub
Sub sw35_UnHit:Controller.Switch(35)=0:End Sub


'Targets
Sub sw13_Hit: vpmTimer.PulseSwitch (13), 0, "" :end sub
Sub sw14_Hit: vpmTimer.PulseSwitch (14), 0, "" :end sub
Sub sw15_Hit: vpmTimer.PulseSwitch (15), 0, "" :end sub
Sub sw34_Hit: vpmTimer.PulseSwitch (34), 0, "" :end sub

Sub Gate_Hit:PlaySound "Gate":End Sub
Sub Gate2_Hit:PlaySound "Gate":End Sub

'Scoring rubbers

sub sw9_hit()
  vpmTimer.PulseSw 9
End Sub

sub sw12_hit()
  vpmTimer.PulseSw 12
End Sub

Sub sw16_hit()
  vpmTimer.PulseSw 16
End Sub

Sub sw18_hit()
  vpmTimer.PulseSw 18
End Sub




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

Sub PlaySoundAtExisting(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,aVol, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / discofever.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / discofever.width-1
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

Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / discofever.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*********************************************************************
'                       Collection Sounds
'*********************************************************************

Sub aRubbers_Hit(idx):PlaySoundAtBallVol "fx_rubber", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aPostRubbers_Hit(idx):PlaySoundAtBallVol "fx_postrubber", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aMetals_Hit(idx):PlaySoundAtBallVol "fx_MetalHit", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBallVol "fx_PlasticHit", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aGates_Hit(idx):PlaySoundAtBallVol "fx_Gate", Vol(ActiveBall)*VolumeDial/10:End Sub
Sub aWoods_Hit(idx):PlaySoundAtBallVol "fx_Woodhit", Vol(ActiveBall)*VolumeDial/10:End Sub


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 2
ReDim rolling(tnob)
InitRolling

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3)

Sub InitRolling
  Dim i
  For i = 0 to tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate_Timer()
  If FlipperType = 1 Then
    batleftshadow.rotz = LeftFlipper.CurrentAngle
    batrightshadow.rotz  = RightFlipper.CurrentAngle
  Else
    batleftshadow.rotz = LeftFlipper1.CurrentAngle - 19
    batrightshadow.rotz  = RightFlipper1.CurrentAngle + 19

    LFlip.RotZ = LeftFlipper1.currentangle -19
    RFlip.RotZ = RightFlipper1.currentangle + 19

    notCollidable

    If abs(140-LeftFlipper1.currentangle) < 5 Then
      bananaL0.collidable=1
    ElseIf abs(140-LeftFlipper1.currentangle) < 10 Then
      bananaL5.collidable=1
    ElseIf abs(140-LeftFlipper1.currentangle) < 15 Then
      bananaL10.collidable=1
    ElseIf abs(140-LeftFlipper1.currentangle) < 20 Then
      bananaL15.collidable=1
    ElseIf abs(140-LeftFlipper1.currentangle) < 25 Then
      bananaL20.collidable=1
    ElseIf abs(140-LeftFlipper1.currentangle) < 30 Then
      bananaL25.collidable=1
    ElseIf abs(140-LeftFlipper1.currentangle) < 35 Then
      bananaL30.collidable=1
    ElseIf abs(140-LeftFlipper1.currentangle) < 40 Then
      bananaL35.collidable=1
    ElseIf abs(140-LeftFlipper1.currentangle) < 45 Then
      bananaL40.collidable=1
    ElseIf abs(140-LeftFlipper1.currentangle) < 50 Then
      bananaL45.collidable=1
    Else
      bananaL50.collidable=1
    End If

    If abs(-140-RightFlipper1.currentangle) < 5 Then
      bananaR0.collidable=1
    ElseIf abs(-140-RightFlipper1.currentangle) < 10 Then
      bananaR5.collidable=1
    ElseIf abs(-140-RightFlipper1.currentangle) < 15 Then
      bananaR10.collidable=1
    ElseIf abs(-140-RightFlipper1.currentangle) < 20 Then
      bananaR15.collidable=1
    ElseIf abs(-140-RightFlipper1.currentangle) < 25 Then
      bananaR20.collidable=1
    ElseIf abs(-140-RightFlipper1.currentangle) < 30 Then
      bananaR25.collidable=1
    ElseIf abs(-140-RightFlipper1.currentangle) < 35 Then
      bananaR30.collidable=1
    ElseIf abs(-140-RightFlipper1.currentangle) < 40 Then
      bananaR35.collidable=1
    ElseIf abs(-140-RightFlipper1.currentangle) < 45 Then
      bananaR40.collidable=1
    ElseIf abs(-140-RightFlipper1.currentangle) < 50 Then
      bananaR45.collidable=1
    Else
      bananaR50.collidable=1
    End If
  End If

  Dim BOT, b
  BOT = GetBalls

  '*** Rolling Sounds***
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("fx_ballrolling" & b)
    BallShadow(b).visible = 0
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

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

    '***Ball Shadows***
    BallShadow(b).X = BOT(b).X + (BOT(b).X - discofever.Width/2)/12
    BallShadow(b).Y = BOT(b).Y + 30 - (abs((BOT(b).X - discofever.Width/2)/12))/2
    BallShadow(b).visible = 1

    '***Ball Drop Sounds***

    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySound "fx_ball_drop" & b, 0, (ABS(BOT(b).velz)/17)^2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If
  Next

  cor.update
End Sub


' **************************** Insert flipper correction here



'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

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
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 30 Then
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
        'playsound "knocker"
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
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
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
    'playsound "knocker"
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
    if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      if DebugOn then
        dim s, bs 'debug spacer, ballspeed
        bs = round(BallSpeed(b),1)
        if bs < 10 then s = " " else s = "" end if
        str = str & b.id & ": " & s & bs & vbnewline
        'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
      end if
    Next
    if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
  End Sub
End Class

Sub RDampen_Timer()
Cor.Update
End Sub

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    x.TimeDelay = 44
  Next
  '"Polarity" Profile
  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.11, -2.5
  AddPt "Polarity", 2, 0.14, -2.5
  AddPt "Polarity", 3, 0.25, -4
  AddPt "Polarity", 4, 0.28, -4
  AddPt "Polarity", 5, 0.31, -3.5
  AddPt "Polarity", 6, 0.34, -4.5
  AddPt "Polarity", 7, 0.37, -4.25
  AddPt "Polarity", 8, 0.4, -4.25
  AddPt "Polarity", 9, 0.45, -3.75
  AddPt "Polarity", 10, 0.48, -3.5
  AddPt "Polarity", 11, 0.51, -3.9
  AddPt "Polarity", 12, 0.55, -3
  AddPt "Polarity", 13, 0.58, -3.5
  AddPt "Polarity", 14, 0.6, -2.25
  AddPt "Polarity", 15, 0.62, -2.5
  AddPt "Polarity", 16, 0.65, -2.25
  AddPt "Polarity", 17, 0.8, -2
  AddPt "Polarity", 18, 0.85, -1.9
  AddPt "Polarity", 19, 1.0, -1
  AddPt "Polarity", 20, 1.2, 0

  '"Velocity" Profile
  addpt "Velocity", 0, 0, 1
  addpt "Velocity", 1, 0.15, 1
  addpt "Velocity", 2, 0.25, 1.1
  addpt "Velocity", 3, 0.28, 1.1
  addpt "Velocity", 4, 0.41, 1.05
  addpt "Velocity", 5, 0.53, 1 '0.982
  addpt "Velocity", 6, 0.702, 0.968
  addpt "Velocity", 7, 0.95, 0.968
  addpt "Velocity", 8, 1.03, 0.945


  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub



Sub TriggerLF_Hit() : If FlipperType = 1 Then LF.Addball activeball End If: End Sub
Sub TriggerLF_UnHit() : If FlipperType = 1 Then LF.PolarityCorrect activeball End If : End Sub
Sub TriggerRF_Hit() : If FlipperType = 1 Then RF.Addball activeball End If : End Sub
Sub TriggerRF_UnHit() : If FlipperType = 1 Then RF.PolarityCorrect activeball End If : End Sub


'*****************UpdateMultipleLamps

Sub lamptimer_timer


If Controller.Lamp(5) Then l5.state = 1 else l5.state=0 'AXS
If Controller.Lamp(6) Then l6.state = 1 else l6.state=0
If Controller.Lamp(7) Then l7.state = 1 else l7.state=0
If Controller.Lamp(8) Then l8.state = 1 else l8.state=0
If Controller.Lamp(9) Then l9.state = 1 else l9.state=0
If Controller.Lamp(10) Then l10.state = 1 else l10.state=0
If Controller.Lamp(11) Then l11.state = 1 else l11.state=0
If Controller.Lamp(12) Then l12.state = 1 else l12.state=0
If Controller.Lamp(13) Then l13.state = 1 else l13.state=0
If Controller.Lamp(14) Then l14.state = 1 else l14.state=0
If Controller.Lamp(15) Then l15.state = 1 else l15.state=0
If Controller.Lamp(16) Then l16.state = 1 else l16.state=0
If Controller.Lamp(17) Then l17.state = 1 else l17.state=0
If Controller.Lamp(18) Then l18.state = 1 else l18.state=0
If Controller.Lamp(19) Then l19.state = 1 else l19.state=0
If Controller.Lamp(20) Then l20.state = 1 else l20.state=0
If Controller.Lamp(21) Then l21.state = 1 else l21.state=0
If Controller.Lamp(22) Then l22.state = 1 else l22.state=0
If Controller.Lamp(23) Then l23.state = 1 else l23.state=0
If Controller.Lamp(24) Then l24.state = 1 else l24.state=0
If Controller.Lamp(25) Then l25.state = 1 else l25.state=0
If Controller.Lamp(26) Then l26.state = 1 else l26.state=0
If Controller.Lamp(27) Then l27.state = 1 else l27.state=0
If Controller.Lamp(28) Then l28.state = 1 else l28.state=0
If Controller.Lamp(29) Then l29.state = 1 else l29.state=0
If Controller.Lamp(30) Then l30.state = 1 else l30.state=0
If Controller.Lamp(31) Then l31.state = 1 else l31.state=0
If Controller.Lamp(32) Then l32.state = 1 else l32.state=0
If Controller.Lamp(33) Then l33.state = 1 else l33.state=0
If Controller.Lamp(34) Then l34.state = 1 else l34.state=0
If Controller.Lamp(35) Then l35.state = 1 else l35.state=0
If Controller.Lamp(36) Then l36.state = 1 else l36.state=0
If Controller.Lamp(37) Then l37.state = 1 else l37.state=0
If Controller.Lamp(38) Then l38.state = 1 else l38.state=0
If Controller.Lamp(39) Then l39.state = 1 else l39.state=0
If Controller.Lamp(40) Then l40.state = 1 else l40.state=0
If Controller.Lamp(41) Then l41.state = 1 else l41.state=0
If Controller.Lamp(42) Then l42.state = 1 else l42.state=0
If Controller.Lamp(43) Then l43.state = 1 else l43.state=0
If Controller.Lamp(44) Then l44.state = 1 else l44.state=0
If Controller.Lamp(45) Then l45.state = 1 else l45.state=0
If Controller.Lamp(46) Then l46.state = 1 else l46.state=0
If Controller.Lamp(47) Then l47.state = 1 else l47.state=0
If Controller.Lamp(48) Then l48.state = 1 else l48.state=0
If Controller.Lamp(56) then L56.state = 1 else L56.state=0

If VRroom then

 If Controller.Lamp(50) Then
  FlPLR1.visible = 1: FlPLR1A.visible = 1:
else
  FlPLR1.visible = 0: FlPLR1A.visible = 0
end if

 If Controller.Lamp(51) Then
  FlPLR2.visible = 1: FlPLR2A.visible = 1:
else
  FlPLR2.visible = 0: FlPLR2A.visible = 0
end if


 If Controller.Lamp(52) Then
  FlPLR3.visible = 1: FlPLR3A.visible = 1:
else
  FlPLR3.visible = 0: FlPLR3A.visible = 0
end if


 If Controller.Lamp(53) Then
  FlPLR4.visible = 1: FlPLR4A.visible = 1:
else
  FlPLR4.visible = 0: FlPLR4A.visible = 0
end if


 If Controller.Lamp(54) Then          '   MATCH
  for each object in colmatch : object.visible = 0 : next
  match1.visible=1
else
  for each object in colmatch : object.visible = 0 : next
  match1.visible=0
end if



 If Controller.Lamp(55) Then          '   BIP
  for each object in COLBALLINPLAY : object.visible = 0 : next
  bip1.visible=1
else
  for each object in colballinplay : object.visible = 0 : next
  bip1.visible=0
end if


 If Controller.Lamp(57) Then          '   1st up
  FlP1.visible = 1: FlP1A.visible = 1: FlP1B.visible = 1: FlP1C.visible = 1: FlP1D.visible = 1
else
  FlP1.visible = 0: FlP1A.visible = 0: FlP1B.visible = 0: FlP1C.visible = 0: FlP1D.visible = 0
end if



 If Controller.Lamp(58) Then          '   2nd up
  FlP2.visible = 1: FlP2A.visible = 1: FlP2B.visible = 1: FlP2C.visible = 1: FlP2D.visible = 1
else
  FlP2.visible = 0: FlP2A.visible = 0: FlP2B.visible = 0: FlP2C.visible = 0: FlP2D.visible = 0
end if


 If Controller.Lamp(59) Then          '     3rd up
  FlP3.visible = 1: FlP3A.visible = 1: FlP3B.visible = 1: FlP3C.visible = 1: FlP3D.visible = 1
else
  FlP3.visible = 0: FlP3A.visible = 0: FlP3B.visible = 0: FlP3C.visible = 0: FlP3D.visible = 0
end if


 If Controller.Lamp(60) Then          '   4th up
  FlP4.visible = 1: FlP4A.visible = 1: FlP4B.visible = 1: FlP4C.visible = 1: FlP4D.visible = 1
else
  FlP4.visible = 0: FlP4A.visible = 0: FlP4B.visible = 0: FlP4C.visible = 0: FlP4D.visible = 0
end if


 If Controller.Lamp(61) Then          '   Tilt
  for each object in coltilt : object.visible = 0 : next
  tilt1.visible=1
else
  for each object in coltilt : object.visible = 0 : next
  tilt1.visible=0
end if


 If Controller.Lamp(62) Then          '   Game over
  for each object in colgameover : object.visible = 0 : next
  GO1.visible=1
else
  for each object in colgameover : object.visible = 0 : next
  GO1.visible=0
end if


 If Controller.Lamp(63) Then          '   SPSA
  for each object in colspsa : object.visible = 0 : next
  spsa.visible=1
else
  for each object in colspsa : object.visible = 0 : next
  spsa.visible=0
end if


 If Controller.Lamp(64) Then          '   HSTD
  for each object in colHSTD : object.visible = 0 : next
  HS1.visible=1
else
  for each object in colHSTD : object.visible = 0 : next
  HS1.visible=0
end if

end if

If VRroom then
  UpdateLedsF
else
  DisplayTimer
end if

end sub



' ********* backglass lights ************



Dim DigitsF(28)
DigitsF(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
DigitsF(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
DigitsF(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
DigitsF(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
DigitsF(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
DigitsF(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)

DigitsF(6) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
DigitsF(7) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
DigitsF(8) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
DigitsF(9) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
DigitsF(10) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
DigitsF(11) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)

DigitsF(12) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006)
DigitsF(13) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
DigitsF(14) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
DigitsF(15) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306)
DigitsF(16) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
DigitsF(17) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)

DigitsF(18) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006)
DigitsF(19) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
DigitsF(20) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
DigitsF(21) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306)
DigitsF(22) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
DigitsF(23) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)


DigitsF(24) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
DigitsF(25) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)
DigitsF(26) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
DigitsF(27) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)


'**** flashing backGlassOn

DIM min
Dim max
Dim disc1, disc2, disc3, disc4, disc5, disc6, disc7
Dim zeit


Sub licht_Timer

min =0
max =6
disc1=(Int((max-min+1)*Rnd+min))
min =0
max =6
disc2=(Int((max-min+1)*Rnd+min))
min =0
max =6
disc3=(Int((max-min+1)*Rnd+min))
min =0
max =6
disc4=(Int((max-min+1)*Rnd+min))
min =0
max =6
disc5=(Int((max-min+1)*Rnd+min))
min =0
max =6
disc6=(Int((max-min+1)*Rnd+min))
min =0
max =1
disc7=(Int((max-min+1)*Rnd+min))




If disc1 >0 then fbb19.visible=1 end if
If disc1 =0 then fbb19.visible=0 end if
If disc2 >0 then fbb22.visible=1 end if
If disc2 =0 then fbb22.visible=0 end if
If disc3 >0 then fbb27.visible=1 end if
If disc3 =0 then fbb27.visible=0 end if
If disc4 >0 then fbb28.visible=1 end if
If disc4 =0 then fbb28.visible=0 end if
If disc5 >0 then fbb33.visible=1 end if
If disc5 =0 then fbb33.visible=0 end if
If disc6 >0 then fbb36.visible=1 end if
If disc6 =0 then fbb36.visible=0 end if
If disc7 >0 then fbb3.visible=1 : fbb11.visible=1: end if
If disc7 =0 then fbb3.visible=0 : fbb11.visible=0: end if


min =250
max =500
zeit=(Int((max-min+1)*Rnd+min))
licht.Interval =zeit
End Sub


'=========================================================
'                    LED Handling
'=========================================================
'Modified version of Scapino's LED code for Fathom
'
Dim SixDigitOutput(32)
Dim DisplayPatterns(11)
Dim DigStorage(32)


'Binary/Hex Pattern Recognition Array
DisplayPatterns(0) = 0    '0000000 Blank
DisplayPatterns(1) = 63   '0111111 zero
DisplayPatterns(2) = 6    '0000110 one
DisplayPatterns(3) = 91   '1011011 two
DisplayPatterns(4) = 79   '1001111 three
DisplayPatterns(5) = 102  '1100110 four
DisplayPatterns(6) = 109  '1101101 five
DisplayPatterns(7) = 125  '1111101 six
DisplayPatterns(8) = 7    '0000111 seven
DisplayPatterns(9) = 127  '1111111 eight
DisplayPatterns(10)= 111  '1101111 nine

'Assign 6-Digit output to reels
Set SixDigitOutput(0)  = P1D6
Set SixDigitOutput(1)  = P1D5
Set SixDigitOutput(2)  = P1D4
Set SixDigitOutput(3)  = P1D3
Set SixDigitOutput(4)  = P1D2
Set SixDigitOutput(5)  = P1D1

Set SixDigitOutput(6)  = P2D6
Set SixDigitOutput(7)  = P2D5
Set SixDigitOutput(8)  = P2D4
Set SixDigitOutput(9)  = P2D3
Set SixDigitOutput(10) = P2D2
Set SixDigitOutput(11) = P2D1

Set SixDigitOutput(12) = P3d6
Set SixDigitOutput(13) = P3d5
Set SixDigitOutput(14) = P3d4
Set SixDigitOutput(15) = P3d3
Set SixDigitOutput(16) = P3d2
Set SixDigitOutput(17) = P3d1

Set SixDigitOutput(18) = P4D6
Set SixDigitOutput(19) = P4D5
Set SixDigitOutput(20) = P4D4
Set SixDigitOutput(21) = P4D3
Set SixDigitOutput(22) = P4D2
Set SixDigitOutput(23) = P4D1

Set SixDigitOutput(24) = CrD2
Set SixDigitOutput(25) = CrD1
Set SixDigitOutput(26) = BaD2
Set SixDigitOutput(27) = BaD1

Sub DisplayTimer' 6-Digit output
  On Error Resume Next
  Dim chgLED,ii,chg,stat,obj,TempCount,temptext,adj

  chgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)

  If Not IsEmpty(chgLED) Then
    For ii = 0 To UBound(chgLED)
      chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      For TempCount = 0 to 10
        If stat = DisplayPatterns(TempCount) then
          SixDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
          DigStorage(chgLED(ii, 0)) = TempCount
        End If
        If stat = (DisplayPatterns(TempCount) + 128) then
          SixDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
          DigStorage(chgLED(ii, 0)) = TempCount
        End If
      Next
    Next
  End IF
End Sub

Sub UpdateLedsF
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x

    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)

    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 28) then
              For Each obj In DigitsF(num)
                   If chg And 1 Then obj.visible=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
      end if
        Next
  End if

 End Sub

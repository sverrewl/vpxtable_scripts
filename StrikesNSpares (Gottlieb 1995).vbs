Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
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
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


Const cGameName="snspares",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000","gts3.vbs",3.2

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------
Const SpareBalls = 3      'Spare Balls per side (0-3)
Const ShowBeerTexture = 1   'after a Strike or a Spare the topper reads "Beer N' Beer" for a short time


Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)= "SolLeftKicker"               'left hole kicker
SolCallback(2)= "SolRightKicker"              'right hole kicker
SolCallback(3)= "bsCenterKicker.solout"           'center kicker between Flippers
'SolCallback(4)= "SolLeftRampGate"              'left plunger gate  - deactivated for manual Queue management
'SolCallback(5)= "SolRightRampGate"             'right plunger gate - deactivated for manual Queue management
SolCallback(6)= "bsTrough.solout"             'Top VUK
SolCallback(7)= "SolPin23"                  'pin 1 trip
SolCallback(8)= "SolPin24"                  'pin 2 trip
SolCallback(9)= "SolPin25"                  'pin 3 trip
SolCallback(10)= "SolPin26"                 'pin 4 trip
SolCallback(11)= "SolPin27"                 'pin 5 trip
SolCallback(12)= "SolPin33"                 'pin 6 trip
SolCallback(13)= "SolPin34"                 'pin 7 trip
SolCallback(14)= "SolPin35"                 'pin 8 trip
SolCallback(15)= "SolPin36"                 'pin 9 trip
SolCallback(16)= "SolPin37"                 'pin 10 trip
SolCallback(17)= "SetLamp 117,"             'left bottom flasher
SolCallback(18)= "SetLamp 118,"             'left center flasher
SolCallback(19)= "SetLamp 119,"             'left top flasher
SolCallback(20)= "SetLamp 120,"             'right top flasher
SolCallback(21)= "SetLamp 121,"             'right center flasher
SolCallback(22)= "SetLamp 122,"             'right bottom flasher
'SolCallback(24)=                     'coin meter
'SolCallback(27)=                     'ticket dispenser
SolCallback(28)= "ResetPins"                'pin reset motor
SolCallback(32)= "SolGameOver"                'game over relay (Q)

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolLeftKicker(enabled)
  if enabled and bsLeftKicker.balls then
    AddToGateQueue(-1)
    bsLeftKicker.ExitSol_on
  end if
End Sub

Sub SolRightKicker(enabled)
  if enabled and bsRightKicker.balls then
    AddToGateQueue(1)
    bsRightKicker.ExitSol_on
  end if
End Sub

Sub SolLeftRampGate(enabled)
  LeftRampLock.isdropped = not enabled
  PlaySound SoundFX("fx_Flipperdown",DOFContactors)
End Sub

Sub SolRightRampGate(enabled)
  RightRampLock.isdropped = not enabled
  PlaySound SoundFX("fx_Flipperdown",DOFContactors):
End Sub

Sub SolGameOver(enabled)

end sub

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsRightKicker, bsLeftKicker, bsCenterKicker

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Strikes n' Spares"&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch=151
    vpmNudge.Sensitivity=2
  vpmNudge.TiltObj = Array(LeftFlipper,RightFlipper)

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 0,10,0,0,0,0,0,0
        bsTrough.InitKick TopKicker,180,5
        bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)

    Set bsLeftKicker=New cvpmBallStack
        bsLeftKicker.InitSaucer LeftKicker,11,100,2+rnd(5)
        bsLeftKicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    Set bsRightKicker=New cvpmBallStack
        bsRightKicker.InitSaucer RightKicker,12,260,2+rnd(5)
        bsRightKicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    Set bsCenterKicker=New cvpmBallStack
        bsCenterKicker.InitSaucer CenterKicker,13,0,50  '15,15
        bsCenterKicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  InitBalls
  Controller.Switch(14) = 1   'Motor intit

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Const swStartButton=4   'thanks to Destruk/TAB

Sub Table1_KeyDown(ByVal keycode)

  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)

  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub


'**********************************************************************************************************

'Nonsense
Sub ShowBeer
  if ShowBeerTexture then
    P_Topper.image = "sns_topper_beer"
    BeerTextureTimer.enabled = True
  end if
End Sub

Sub BeerTextureTimer_Timer
  BeerTextureTimer.enabled = False
  P_Topper.image = "sns_topper"
End Sub

' Drain hole and kickers
Sub BallDrop_Hit()
  BallDrop.destroyball
  Troughtop.createball
  Troughtop.kick 90,3
  playsound "Drain5"
End Sub

Sub TroughBottom_Hit
  bsTrough.AddBall Me
End Sub

Sub LeftKicker_Hit:bsLeftKicker.addball me : playsoundAtVol "popper_ball", ActiveBall, VolKick: End Sub
Sub RightKicker_Hit:bsRightKicker.addball me : playsoundAtVol "popper_ball", ActiveBall, VolKick: End Sub
Sub CenterKicker_Hit:bsCenterKicker.addball me : playsoundAtVol "popper_ball", ActiveBall, VolKick: End Sub

Sub LeftBallFeed_Hit:Controller.Switch(20)=1:End Sub
Sub LeftBallFeed_unHit:Controller.Switch(20)=0:End Sub
Sub RightBallFeed_Hit:Controller.Switch(30)=1:End Sub
Sub RightBallFeed_unHit:Controller.Switch(30)=0:End Sub


'Rollover Switches
Const unhitState = 1
sub Trigger23_hit:Controller.Switch(23)=1:P_T23.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger23_unhit:Controller.Switch(23)=unhitState:Trigger23.timerenabled=True:End Sub
sub Trigger24_hit:Controller.Switch(24)=1:P_T24.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger24_unhit:Controller.Switch(24)=unhitState:Trigger24.timerenabled=True:End Sub
sub Trigger25_hit:Controller.Switch(25)=1:P_T25.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger25_unhit:Controller.Switch(25)=unhitState:Trigger25.timerenabled=True:End Sub
sub Trigger26_hit:Controller.Switch(26)=1:P_T26.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger26_unhit:Controller.Switch(26)=unhitState:Trigger26.timerenabled=True:End Sub
sub Trigger27_hit:Controller.Switch(27)=1:P_T27.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger27_unhit:Controller.Switch(27)=unhitState:Trigger27.timerenabled=True:End Sub

sub Trigger33_hit:Controller.Switch(33)=1:P_T33.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger33_unhit:Controller.Switch(33)=unhitState:Trigger33.timerenabled=True:End Sub
sub Trigger34_hit:Controller.Switch(34)=1:P_T34.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger34_unhit:Controller.Switch(34)=unhitState:Trigger34.timerenabled=True:End Sub
sub Trigger35_hit:Controller.Switch(35)=1:P_T35.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger35_unhit:Controller.Switch(35)=unhitState:Trigger35.timerenabled=True:End Sub
sub Trigger36_hit:Controller.Switch(36)=1:P_T36.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger36_unhit:Controller.Switch(36)=unhitState:Trigger36.timerenabled=True:End Sub
sub Trigger37_hit:Controller.Switch(37)=1:P_T37.RotX=30 : playsoundAtVol"FlipperDown_Akiles" , ActiveBall, 1: End Sub
sub Trigger37_unhit:Controller.Switch(37)=unhitState:Trigger37.timerenabled=True:End Sub


sub Trigger23_Timer
  Trigger23.timerenabled = False
  P_T23.RotX = 0
End Sub
sub Trigger24_Timer
  Trigger24.timerenabled = False
  P_T24.RotX = 0
End Sub
sub Trigger25_Timer
  Trigger25.timerenabled = False
  P_T25.RotX = 0
End Sub
sub Trigger26_Timer
  Trigger26.timerenabled = False
  P_T26.RotX = 0
End Sub
sub Trigger27_Timer
  Trigger27.timerenabled = False
  P_T27.RotX = 0
End Sub

sub Trigger33_Timer
  Trigger33.timerenabled = False
  P_T33.RotX = 0
End Sub
sub Trigger34_Timer
  Trigger34.timerenabled = False
  P_T34.RotX = 0
End Sub
sub Trigger35_Timer
  Trigger35.timerenabled = False
  P_T35.RotX = 0
End Sub
sub Trigger36_Timer
  Trigger36.timerenabled = False
  P_T36.RotX = 0
End Sub
sub Trigger37_Timer
  Trigger37.timerenabled = False
  P_T37.RotX = 0
End Sub



'**********************************************************************************************************
'------  Sized Balls  ------
'**********************************************************************************************************

Dim GateQueue,GQ0,GQ1,GQ2,GQ3,GQ4,GQ5,GQ6,GQ7,GQ8,GQ9
GateQueue = Array(GQ0,GQ1,GQ2,GQ3,GQ4,GQ5,GQ6,GQ7,GQ8,GQ9)
InitGateQueue

Dim n
Sub AddToGateQueue(DirPar)
  n = 0
  while GateQueue(n) <> 0
    n = n + 1
  wend
  GateQueue(n) = DirPar
End Sub

Dim o
Sub ShiftGateQueue
  for o = 0 to 8
    GateQueue(o) = GateQueue(o+1)
  next
  GateQueue(9) = 0
End Sub

Sub InitGateQueue
  for o = 0 to 9
    GateQueue(o) = 0
  next
End Sub

Const BowlingBallSize = 29.7    '1 1/4 inch = 29.41
Sub BallLaunch_Hit
  BallLaunch.destroyball

  SolLeftRampGate(GateQueue(0) = 1)
  SolRightRampGate(GateQueue(0) = -1)
  ShiftGateQueue

  BallRelease.createsizedball (BowlingBallSize)
  BallRelease.kick 180,5
end sub

Dim LoopCount,LoopCount2
Sub InitBalls
  LoopCount = 0
  LoopCount2 = 0
  SpareBallInitTimer.enabled = True
  BallInitTimer.enabled = True
end sub

Sub SpareBallInitTimer_Timer
  LoopCount = LoopCount + 1

  if LoopCount >= SpareBalls then
    SpareBallInitTimer.enabled = False
  end if

  if LoopCount <= SpareBalls then
    select case Loopcount
      case 1: Kicker1.createsizedball (BowlingBallSize)
          Kicker1.kick 180,5
          Kicker2.createsizedball (BowlingBallSize)
          Kicker2.kick 180,5
      case 2: Kicker3.createsizedball (BowlingBallSize)
          Kicker3.kick 180,5
          Kicker4.createsizedball (BowlingBallSize)
          Kicker4.kick 180,5
      case 3: Kicker5.createsizedball (BowlingBallSize)
          Kicker5.kick 180,5
          Kicker6.createsizedball (BowlingBallSize)
          Kicker6.kick 180,5
    end select
  end if
End Sub

Sub BallInitTimer_Timer
  LoopCount2 = LoopCount2 + 1
  PlaySound SoundFX("ballrelease",DOFContactors)
  select case LoopCount2
    case 1,3,5,7
          Kicker7.createsizedball (29.41)
          Kicker7.kick 180,5
    case 2,4,6,8
          Kicker8.createsizedball (29.41)
          Kicker8.kick 180,5
  end select

  if LoopCount2 >= 8 then
    BallInitTimer.enabled = False
  end if
End Sub


'**********************************************************************************************************
'------  Pin animation  ------
'**********************************************************************************************************

Dim Pins,PinDir,Pin23Dir,Pin24Dir,Pin25Dir,Pin26Dir,Pin27Dir,Pin33Dir,Pin34Dir,Pin35Dir,Pin36Dir,Pin37Dir
PinDir = Array(Pin23Dir,Pin24Dir,Pin25Dir,Pin26Dir,Pin27Dir,Pin33Dir,Pin34Dir,Pin35Dir,Pin36Dir,Pin37Dir)
Pins = Array(Pin23,Pin24,Pin25,Pin26,Pin27,Pin33,Pin34,Pin35,Pin36,Pin37)
Const PinAngle = -60

Sub SolPin23(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(0) = -1
  end if
End Sub
Sub SolPin24(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(1) = -1
  end if
End Sub
Sub SolPin25(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(2) = -1
  end if
End Sub
Sub SolPin26(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(3) = -1
  end if
End Sub
Sub SolPin27(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(4) = -1
  end if
End Sub

Sub SolPin33(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(5) = -1
  end if
End Sub
Sub SolPin34(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(6) = -1
  end if
End Sub
Sub SolPin35(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(7) = -1
  end if
End Sub
Sub SolPin36(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(8) = -1
  end if
End Sub
Sub SolPin37(Enabled)
  If Enabled Then
    Playsound "soloff"
    PinDir(9) = -1
  end if
End Sub

Sub ResetPins(Enabled)
  If Enabled Then
    PinDir(0) = 1
    PinDir(1) = 1
    PinDir(2) = 1
    PinDir(3) = 1
    PinDir(4) = 1
    PinDir(5) = 1
    PinDir(6) = 1
    PinDir(7) = 1
    PinDir(8) = 1
    PinDir(9) = 1
    MotorTimer.enabled = True
  Else
    Controller.Switch(14)=0
  End If
End Sub

Dim AllPinsReset,m
Sub MotorTimer_Timer
  PlaySound SoundFX("Motor1",DOFContactors)
  AllPinsReset = True
  for m = 0 to 9
    if Pins(m).RotX <> 0 then
      AllPinsReset = False
    end if
  next
  if AllPinsReset then
    MotorTimer.enabled = False
    Controller.Switch(23) = 0
    Controller.Switch(24) = 0
    Controller.Switch(25) = 0
    Controller.Switch(26) = 0
    Controller.Switch(27) = 0
    Controller.Switch(33) = 0
    Controller.Switch(34) = 0
    Controller.Switch(35) = 0
    Controller.Switch(36) = 0
    Controller.Switch(37) = 0
    Controller.Switch(14)=1
  end if
End Sub

Dim j
Sub PinFlipTimer_Timer
  for j = 0 to 9
    if PinDir(j) < 0 then
      Pins(j).RotX = Pins(j).RotX - 7     'Flip speed
      if Pins(j).RotX <= PinAngle then
        Pins(j).RotX = Pinangle
        PinDir(j) = 0
      end if
    end if
  next
End Sub

Dim k
Sub PinResetTimer_Timer
  for k = 0 to 9
    if PinDir(k) > 0 then
      Pins(k).RotX = Pins(k).RotX + 1     'Reset speed
      if Pins(k).RotX >= 0 then
        Pins(k).RotX = 0
        PinDir(k) = 0
      end if
    end if
  next
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
LampTimer.Interval = 5 'lamp fading speed
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


Sub UpdateLamps()
NfadeL 0,  L0
'NfadeL 1,  L1 'Credit Button
NfadeL 3,  L3
NfadeL 4,  L4
NfadeL 5,  L5
NfadeL 6,  L6
NfadeL 7,  L7
NfadeL 10,  L10
NfadeL 11,  L11
NfadeL 12,  L12
NfadeL 13,  L13
NfadeL 14,  L14
NfadeL 15,  L15
NfadeL 16,  L16
NfadeL 17,  L17
NfadeL 21,  L21
NfadeL 22,  L22
NfadeL 23,  L23
NfadeL 24,  L24
NfadeL 25,  L25
NfadeL 26,  L26
NfadeL 27,  L27

'NfadeL 30,  L30 'Lightbox #1
'NfadeL 31,  L31 'Lightbox #2
'NfadeL 32,  L32 'Lightbox #3
'NfadeL 33,  L33 'Lightbox #4
'NfadeL 34,  L34 'Lightbox #5
'NfadeL 35,  L35 'Lightbox #6
'NfadeL 36,  L36 'Lightbox #7
'NfadeL 40,  L40 'Lightbox #8
'NfadeL 41,  L41 'Lightbox #9
'NfadeL 42,  L42 'Lightbox #10
'NfadeL 43,  L43 'Lightbox #11
'NfadeL 44,  L44 'Lightbox #12
'NfadeL 45,  L45 'Lightbox #13
'NfadeL 46,  L46 'Lightbox #14

'Solenoid Controlled Lamps

    NFadeObjm 117, P117, "dome3_redON", "dome3_red"  'Dome
    NfadeL 117,  F117

    NFadeObjm 118, P118, "dome3_yellowON", "dome3_yellow"  'Dome
    NfadeL 118,  F118

    NFadeObjm 119, P119, "dome3_clearON", "dome3_clear"  'Dome
    NfadeL 119,  F119

    NFadeObjm 120, P120, "dome3_clearON", "dome3_clear"  'Dome
    NfadeL 120,  F120

    NFadeObjm 121, P121, "dome3_yellowON", "dome3_yellow"  'Dome
    NfadeL 121,  F121

    NFadeObjm 122, P122, "dome3_redON", "dome3_red"  'Dome
    NfadeL 122,  F122

End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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
    If TypeName(object) = "Light" Then
        Object.IntensityScale = FadingLevel(nr)/128
        Object.State = LampState(nr)
    End If
    If TypeName(object) = "Flasher" Then
        Object.IntensityScale = FadingLevel(nr)/128
        Object.visible = LampState(nr)
    End If
    If TypeName(object) = "Primitive" Then
        Object.DisableLighting = LampState(nr)
    End If
End Sub


 'Reels
Sub FadeReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
    End Select
End Sub

 'Inverted Reels
Sub FadeIReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
    End Select
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
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
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

  FlipperT3RS.roty=leftFlipper.CurrentAngle +240
  FlipperT3RS1.roty=rightFlipper.CurrentAngle +120

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

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


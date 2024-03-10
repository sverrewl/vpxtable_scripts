Option Explicit
Randomize

Const cGameName = "pnkpnthr"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01210000", "sys80.VBS", 3.1

'*********** Desktop/Cabinet settings ************************
Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim hiddenvalue

If DesktopMode = True Then 'Show Desktop components
hiddenvalue=0
Else
hiddenvalue=1
End if

'****************************** Script for VR Rooms  ********************
' VR Option
Dim VRRoom, Object, UseVPMDMD, InfoBox

'VR Room
VRRoom = 0      ' 0 = desktop/fullscreen, 1 = pink room, 2 = minimal VR Room
Infobox = 0     ' 0 = huge infobox in rooms, 1 = no infobox

If VRRoom = 1 Or VRRoom = 2 then UseVPMDMD = true Else UseVPMDMD = DesktopMode

If VRRoom = 0 Then
  for each Object in ColVR : object.visible = 0 : next
End If

If VRRoom = 1 Then  ' pink room
  for each Object in ColRoom1 : object.visible = 1 : next
  for each Object in ColRoomMinimal : object.visible = 0 : next
  for each Object in ColBackdrop : object.visible = 0 : next
    If InfoBox = 0 Then for each Object in ColInfobox : object.visible = 1 : next
    If InfoBox = 1 Then for each Object in ColInfobox : object.visible = 0 : next
End If

If VRRoom = 2 Then  ' minimal room
  for each Object in ColRoom1 : object.visible = 0 : next
  for each Object in ColRoomMinimal : object.visible = 1 : next
  for each Object in ColBackdrop : object.visible = 0 : next
    If InfoBox = 0 Then for each Object in ColInfobox : object.visible = 1 : next
    If InfoBox = 1 Then for each Object in ColInfobox : object.visible = 0 : next
End If

'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1          '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

' Block shadow from playfield in plungerlane
Dim ControlBall, contballinplay
Sub StartControl_Hit()
  Set ControlBall = ActiveBall
  contballinplay = True
  VR_ShadowBlocker.visible = False
End Sub

Sub ShadowOn_Hit()
  VR_ShadowBlocker.visible = True
End Sub

'*************
'//////////////---- LUT (Colour Look Up Table) ----//////////////
'*************
'0 = Fleep Natural Dark 1
'1 = Fleep Natural Dark 2
'2 = Fleep Warm Dark
'3 = Fleep Warm Bright
'4 = Fleep Warm Vivid Soft
'5 = Fleep Warm Vivid Hard
'6 = Skitso Natural and Balanced
'7 = Skitso Natural High Contrast
'8 = 3rdaxis Referenced THX Standard
'9 = CalleV Punchy Brightness and Contrast
'10 = HauntFreaks Desaturated
'11 = Tomate Washed Out
'12 = VPW Original 1 to 1
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book

Dim LUTset, LutToggleSound
LutToggleSound = True
LoadLUT
SetLUT

'*************
'//////////////---- Flyer ----//////////////
'*************

Dim FlyerSet, DisableFlyerSelector
DisableFlyerSelector = 1

'************************************************
'************************************************
'************************************************
'************************************************
'************************************************
Const UseSolenoids = True
Const UseLamps = True
Const UseSync = False
Const UseGI = False

' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

Dim bsTrough, bskicker, bslLock, bsrLock, dttop, dtleft, FastFlips

DisplayTimer.Enabled = true

Sub FlipperTimer_Timer
    'Add flipper, gate and spinner rotations here
  FlipperLB.Rotz = LeftFlipper.CurrentAngle
  FlipperLR.Rotz = LeftFlipper.CurrentAngle
  FlipperRB1.Rotz = RightFlipper.CurrentAngle
  FlipperRR1.Rotz = RightFlipper.CurrentAngle
  FlipperRB.Rotz = RightFlipper.CurrentAngle
  FlipperRR.Rotz = RightFlipper.CurrentAngle
  FlipperLSh.rotz = LeftFlipper.currentangle '+ 45
  FlipperRSh.rotz = RightFlipper1.currentangle '+ 45
  FlipperRSh1.rotz = RightFlipper1.currentangle '+ 45

  BGDisplayTimer

End Sub

'*****************************************************************************************************
' VR PLUNGER ANIMATION
'
' Code needed to animate the plunger. If you pull the plunger it will move in VR.
' IMPORTANT: there are two numeric values in the code that define the postion of the plunger and the
' range in which it can move. The first numeric value is the actual y position of the plunger primitive
' and the second is the actual y position + 135 to determine the range in which it can move.
'
' You need to to select the Primary_plunger primitive you copied from the
' template and copy the value of the Y position
' (e.g. 2114) into the code. The value that determines the range of the plunger is always the y
' position + 135 (e.g. 2249).
'
'*****************************************************************************************************

Sub TimerPlunger_Timer
  If VR_Primary_plunger.Y < 2485 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
 'debug.print plunger.position
  VR_Primary_plunger.Y = 2350 + (5* Plunger.Position) -20
End Sub

'*****************************************************
' Table Init
'*****************************************************

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Pink Panther (Gottlieb 1981)"&chr(13)&"1.0"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = False
    .Hidden = HiddenValue
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

'Trough
   Set bsTrough=New cvpmBallStack
    with bsTrough
        .InitSw 63,73,0,0,0,0,0,0
        .InitKick ballrelease, 110, 12
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
        .Balls=3
    end with

'Kickers
    Set bskicker=New cvpmBallStack
    with bskicker
        .InitSaucer sw13,13,255,12
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
        .kickanglevar=2
    end with

    Set bslLock=New cvpmBallStack
    with bslLock
        .InitSaucer sw43,43,-170,10
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
    end with

    Set bsrLock=New cvpmBallStack
    with bsrLock
        .InitSaucer sw53,53,320,11
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
    end with

' Nudging
    vpmNudge.TiltSwitch = 57
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, bumper1, bumper2, bumper3, bumper4)

    Set dtTop=New cvpmDropTarget
        dtTop.InitDrop Array(sw51,sw61,sw71),Array(51,61,71)
        dtTop.InitSnd SoundFX("fx_droptarget",DOFDropTargets),SoundFX("fx_DTReset",DOFDropTargets)

    Set dtLeft=New cvpmDropTarget
        dtLeft.InitDrop Array(sw40,sw50,sw60,sw70),Array(40,50,60,70)
        dtLeft.InitSnd SoundFX("fx_droptarget",DOFDropTargets),SoundFX("fx_DTReset",DOFDropTargets)

if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if

    if flippershadows=1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
        FlipperRSh1.visible=1
      else
        FlipperLSh.visible=0
        FlipperRSh.visible=0
        FlipperRSh1.visible=0
    end if

    Set FastFlips = new cFastFlips
    with FastFlips
        .CallBackL = "SolLflipper"  'Point these to flipper subs
        .CallBackR = "SolRflipper"  '...
    '   .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
    '   .CallBackUR = "SolURflipper"'...
        .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
    '   .DebugOn = False        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
    end with

  If VRRoom = 1 Or VRRoom = 2  Then
    setup_backglass()
  End If

End Sub

'************************************************
' Solenoids
'************************************************
SolCallback(1) =    "SolKicker"
SolCallback(2) =    "SolTopTargetReset"
'SolCallback(3) =   coin1
'SolCallback(4) =   coin2
SolCallback(5) =    "SolLeftTargetReset"
SolCallback(6) =    "delayedSolOut"
'SolCallback(7) =   coin3
SolCallback(8) =    "solknocker"
SolCallback(9) =    "bstrough.SolIn"
SolCallback(10) = "FastFlips.TiltSol"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors)
    'LeftFlipper.RotateToEnd
    lf.fire
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors)
    'RightFlipper.RotateToEnd
    'RightFlipper1.RotateToEnd
    RF.fire
    RF1.fire
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub

'***** EM kicker animation

Dim RightEMPos

Sub SolKicker(enabled)
  If enabled Then
    bskicker.ExitSol_On
    RightEMpos = 0
    PkickarmR.RotZ = 2
    RightEMTimer.Enabled = 1
  End If
End Sub

Sub RightEMTimer_Timer
    Select Case RightEMpos
        Case 1:PkickarmR.Rotz = 15
        Case 2:PkickarmR.Rotz = 15
        Case 3:PkickarmR.Rotz = 15
        Case 4:PkickarmR.Rotz = 8
        Case 5:PkickarmR.Rotz = 4
        Case 6:PkickarmR.Rotz = 2
        Case 7:PkickarmR.Rotz = 0:RightEMTimer.Enabled = 0
    End Select
    RightEMpos = RightEMpos + 1
End Sub

Dim timer

Sub delayedSolOut(enabled)
    if not timer then bsTrough.SolOut enabled
    if not enabled then timer = True : ballrelease.TimerEnabled = True
End Sub

Sub ballrelease_Timer()
    ballrelease.TimerEnabled = False
    timer = False
End Sub


'*****Drop Lights Off
   dim xx
    For each xx in dtLeftLights: xx.state=0:Next

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

  KeyDownHandler(KeyCode)

  If Keycode = LeftMagnaSave Then             ' Change LUT with magna-save buttons
        LUTSet = LUTSet  + 1
    if LutSet > 15 then LUTSet = 0
        lutsetsounddir = 1
      If LutToggleSound then
        If lutsetsounddir = 1 And LutSet <> 15 Then
          Playsound "click", 0, 1, 0.1, 0.25
        End If
      If lutsetsounddir = -1 And LutSet <> 15 Then
          Playsound "click", 0, 1, 0.1, 0.25
        End If
      If LutSet = 15 Then
          Playsound "gun", 0, 1, 0.1, 0.25
        End If
        LutSlctr.enabled = true
      SetLUT
      ShowLUT
    End If
  End If

  If keycode = LeftFlipperKey Then
    FastFlips.FlipL True
    FastFlips.FlipUL True
  End If

  If KeyCode = RightFlipperKey then
    FastFlips.FlipR True
    FastFlips.FlipUR True
  End If

  If VRRoom = 1 Or VRRoom = 2  Then
    If keycode = LeftFlipperKey Then
      VR_CabFlipperLeft.X = VR_CabFlipperLeft.X +10 ' Animate VR Left flipperbutton
      If DisableFlyerSelector = 0 Then
        FlyerSet = Flyerset - 1
        if FlyerSet < 0 then FlyerSet = 3
        End If
      End If
    If keycode = RightFlipperKey Then
      VR_CabFlipperRight.X = VR_CabFlipperRight.X - 10  ' Animate VR Right flipperbutton
      If DisableFlyerSelector = 0 Then
        FlyerSet = FlyerSet + 1
        if FlyerSet > 3 then FlyerSet = 0
        End If
      End If
  End If


  If Keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5    ' Animate VR Startbutton
    DisableFlyerSelector = 1
  End If

  If keycode = PlungerKey Then
    Plunger.Pullback
    playsound"plungerpull"
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

  If keycode=AddCreditKey then
    PlaySound "fx_coin", 0, 1, 0.1, 0.25
  end if

End Sub

Sub Table1_KeyUp(ByVal KeyCode)

  KeyUpHandler(KeyCode)


  If KeyCode = LeftFlipperKey then
    FastFlips.FlipL False
    FastFlips.FlipUL False
  End If

  If KeyCode = RightFlipperKey then
    FastFlips.FlipR False
    FastFlips.FlipUR False
  End If

  If VRRoom = 1 Or VRRoom = 2  Then
    If keycode = LeftFlipperKey Then
      VR_CabFlipperLeft.X = VR_CabFlipperLeft.X -10 ' Animate VR Left flipper
    End If
    If keycode = RightFlipperKey Then
      VR_CabFlipperRight.X = VR_CabFlipperRight.X +10 ' Animate VR Right flipperbutton
    End If
    If Keycode = StartGameKey Then
      VR_Cab_StartButton.y = VR_Cab_StartButton.y +5  ' Animate VR Startbutton
    End If
  End If

  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "plunger", 0, 1, 0.1, 0.25
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    VR_Primary_plunger.Y = 2350
  end if

  If Keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y +5    ' Animate VR Startbutton
  End If

End Sub

'**********************************************************************************************************
'LUT selector timer
'**********************************************************************************************************

dim lutsetsounddir
sub LutSlctr_timer
  If lutsetsounddir = 1 And LutSet <> 15 Then
    Playsound "click", 0, 1, 0.1, 0.25
  End If
  If lutsetsounddir = -1 And LutSet <> 15 Then
    Playsound "click", 0, 1, 0.1, 0.25
  End If
  If LutSet = 15 Then
    Playsound "gun", 0, 1, 0.1, 0.25
  End If
  LutSlctr.enabled = False
end sub

'**********************************************************************************************************

Sub Drain_Hit():PlaySound "fx_drain":bstrough.addball me:End Sub
Sub sw13_Hit:PlaySound "kicker_enter_center":bskicker.AddBall 0:End Sub
Sub sw43_Hit:PlaySound "kicker_enter_center":bslLock.AddBall 0:End Sub
Sub sw53_Hit:PlaySound "kicker_enter_center":bsrLock.AddBall 0:End Sub

'Drop Targets
Sub Sw51_Dropped:dtTop.Hit 1 : End Sub
 Sub Sw61_Dropped:dtTop.Hit 2 : End Sub
 Sub Sw71_Dropped:dtTop.Hit 3 : End Sub

 Sub Sw40_Dropped:dtLeft.Hit 1 : D1L1.state=1 : End Sub
 Sub Sw50_Dropped:dtLeft.Hit 2 : D2L1.state=1 : D2L2.state=1 : End Sub
 Sub Sw60_Dropped:dtLeft.Hit 3 : D3L1.state=1 : D3L2.state=1 : End Sub
 Sub Sw70_Dropped:dtLeft.Hit 4 : D4L1.state=1 : D4L2.state=1 : End Sub

Sub SolTopTargetReset(enabled)
    dim xx
    if enabled then
        dtTop.SolDropUp enabled
    end if
End Sub

Sub SolLeftTargetReset(enabled)
    dim xx
    if enabled then
        dtLeft.SolDropUp enabled
        For each xx in DTLeftLights: xx.state=0:Next
    end if
End Sub

'Bumpers

Sub bumper1_Hit : vpmTimer.PulseSw 23 : playsound SoundFX("fx_bumper1",DOFContactors): DOF 206, DOFPulse:End Sub
Sub bumper2_Hit : vpmTimer.PulseSw 23 : playsound SoundFX("fx_bumper2",DOFContactors): DOF 209, DOFPulse:End Sub
Sub bumper3_Hit : vpmTimer.PulseSw 23 : playsound SoundFX("fx_bumper3",DOFContactors): DOF 208, DOFPulse:End Sub
Sub bumper4_Hit : vpmTimer.PulseSw 23 : playsound SoundFX("fx_bumper4",DOFContactors): DOF 207, DOFPulse:End Sub

'Wire Triggers
Sub SW02_Hit:Controller.Switch(02)=1:PlaySound"rollover":End Sub    'P
Sub SW02_unHit:Controller.Switch(02)=0:End Sub
Sub SW12_Hit:Controller.Switch(12)=1:PlaySound"rollover":End Sub    'I
Sub SW12_unHit:Controller.Switch(12)=0:End Sub
Sub SW22_Hit:Controller.Switch(22)=1:PlaySound"rollover":End Sub    'N
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW32_Hit:Controller.Switch(32)=1:PlaySound"rollover":End Sub    'K
Sub SW32_unHit:Controller.Switch(32)=0:End Sub
Sub SW03_Hit:Controller.Switch(03)=1:PlaySound"rollover":End Sub    'right inlane
Sub SW03_unHit:Controller.Switch(03)=0:End Sub
Sub SW42_Hit:Controller.Switch(42)=1:PlaySound"rollover":End Sub    'left inlane
Sub SW42_unHit:Controller.Switch(42)=0:End Sub
Sub SW62_Hit:Controller.Switch(62)=1:PlaySound"rollover":End Sub    'left outlane
Sub SW62_unHit:Controller.Switch(62)=0:End Sub
Sub SW72_Hit:Controller.Switch(72)=1:PlaySound"rollover":End Sub    'right outlane
Sub SW72_unHit:Controller.Switch(72)=0:End Sub

'Targets
Sub sw00_Hit:vpmTimer.PulseSw (00):playsound"target":End Sub
Sub sw10_Hit:vpmTimer.PulseSw (10):playsound"target":End Sub
Sub sw20_Hit:vpmTimer.PulseSw (20):playsound"target":End Sub
Sub sw30_Hit:vpmTimer.PulseSw (30):playsound"target":End Sub
Sub sw01_Hit:vpmTimer.PulseSw (01):playsound"target":End Sub
Sub sw11_Hit:vpmTimer.PulseSw (11):playsound"target":End Sub
Sub sw21_Hit:vpmTimer.PulseSw (21):playsound"target":End Sub
Sub sw31_Hit:vpmTimer.PulseSw (31):playsound"target":End Sub
Sub sw41_Hit:vpmTimer.PulseSw (41):playsound"target":End Sub

Sub SolKnocker(Enabled)
    If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************

Dim RStep, Lstep, D1Step, D2Step, D3Step, D4Step, D5Step, D6Step

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    DOF 202, DOFPulse
    vpmtimer.PulseSw(33)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing3.Visible = 1:sling1.TransZ = 0
        Case 5:RSLing3.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors),0,1,-0.05,0.05
    DOF 201, DOFPulse
    vpmtimer.pulsesw(33)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing3.Visible = 1:sling2.TransZ = 0
        Case 5:LSLing3.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Dingwall1_Slingshot
    vpmtimer.pulsesw(33)
    Rubber7a.Visible = 0
    Rubber7a1.Visible = 1
    D1Step = 0
    Dingwall1.TimerEnabled = 1
End Sub

Sub Dingwall1_Timer
    Select Case D1Step
        Case 3:Rubber7a1.Visible = 0:Rubber7a2.Visible = 1
        Case 4:Rubber7a2.Visible = 0:Rubber7a.Visible = 1:Dingwall1.TimerEnabled = 0
    End Select
    D1Step = D1Step + 1
End Sub

Sub Dingwall2_Slingshot
    Rubber7a.Visible = 0
    Rubber7a3.Visible = 1
    D2Step = 0
    Dingwall2.TimerEnabled = 1
End Sub

Sub Dingwall2_Timer
    Select Case D2Step
        Case 3:Rubber7a3.Visible = 0:Rubber7a4.Visible = 1
        Case 4:Rubber7a4.Visible = 0:Rubber7a.Visible = 1:Dingwall2.TimerEnabled = 0
    End Select
    D2Step = D2Step + 1
End Sub

Sub Dingwall3_Slingshot
    Rubberc.Visible = 0
    Rubberc1.Visible = 1
    D3Step = 0
    Dingwall3.TimerEnabled = 1
End Sub

Sub Dingwall3_Timer
    Select Case D3Step
        Case 3:Rubberc1.Visible = 0:Rubberc2.Visible = 1
        Case 4:Rubberc2.Visible = 0:Rubberc.Visible = 1:Dingwall3.TimerEnabled = 0
    End Select
    D3Step = D3Step + 1
End Sub

Sub Dingwall4_Slingshot
    Rubberd.Visible = 0
    Rubberd1.Visible = 1
    D4Step = 0
    Dingwall4.TimerEnabled = 1
End Sub

Sub Dingwall4_Timer
    Select Case D4Step
        Case 3:Rubberd1.Visible = 0:Rubberd2.Visible = 1
        Case 4:Rubberd2.Visible = 0:Rubberd.Visible = 1:Dingwall4.TimerEnabled = 0
    End Select
    D4Step = D4Step + 1
End Sub

Sub Dingwall5_Slingshot
    vpmtimer.pulsesw(33)
    Rubbere.Visible = 0
    Rubbere1.Visible = 1
    D5Step = 0
    Dingwall5.TimerEnabled = 1
End Sub

Sub Dingwall5_Timer
    Select Case D5Step
        Case 3:Rubbere1.Visible = 0:Rubbere2.Visible = 1
        Case 4:Rubbere2.Visible = 0:Rubbere.Visible = 1:Dingwall5.TimerEnabled = 0
    End Select
    D5Step = D5Step + 1
End Sub

Sub Dingwall6_Slingshot
    vpmtimer.pulsesw(33)
    Rubberf.Visible = 0
    Rubberf1.Visible = 1
    D6Step = 0
    Dingwall6.TimerEnabled = 1
End Sub

Sub Dingwall6_Timer
    Select Case D6Step
        Case 3:Rubberf1.Visible = 0:Rubberf2.Visible = 1
        Case 4:Rubberf2.Visible = 0:Rubberf.Visible = 1:Dingwall6.TimerEnabled = 0
    End Select
    D6Step = D6Step + 1
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed, decrease the default 2000 to hear a louder rolling ball sound
  Vol = Csng(BallVel(ball) ^2 / 1000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the Table. "Table" is the name of the table
  Dim tmp
    tmp = ball.x * 2 / Table1.width-1
    If tmp> 0 Then
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

Const tnob = 3 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************
'Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aTargets_Hit(idx):PlaySound "fx2_target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

Sub aRubber_Pins_Hit (idx)
    PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub aTargets_Hit (idx)
    PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub aMetals_Hit (idx)
    PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub aGates_Hit (idx)
    PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub aRubber_Bands_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub aRubber_Posts_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
        PlaySound "fx_postrubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
    Select Case Int(Rnd*3)+1
        Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
        Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End Select
End Sub

'GI uselamps workaround
dim GIlamps : set GIlamps = New GIcatcherobject
Class GIcatcherObject   'object that disguises itself as a light. (UseLamps workaround for System80 GI circuit)
    Public Property Let State(input)
        dim x
        if input = 1 then 'If GI switch is engaged, turn off GI.
            for each x in gi : x.state = 0 : next
        elseif input = 0 then
            for each x in gi : x.state = 1 : next
        end if
        'tb.text = "gitcatcher.state = " & input    'debug
    End Property
End Class

'-------------------------------------
' Map lights into array
' Set unmapped lamps to Nothing
'-------------------------------------

Set Lights(0)  = l0 'ball in play
'Set Lights(1)  = l1 'tilt
set Lights(1) = GIlamps 'GI circuit
'Set Lights(2)  = l2    coin lockout coil
Set Lights(3)  = l3    'shoot again
Set Lights(8)  = l8 'left capture release
Set Lights(9)  = l9 'right capture release
'Set Lights(10) = l10   high score
'Set Lights(11) = l11   game over
Set Lights(12) = l12
Set Lights(13) = l13
Set Lights(14) = l14
Set Lights(15) = l15
Set Lights(16) = l16
Set Lights(17) = l17
Set Lights(18) = l18
Set Lights(19) = l19
Set Lights(20) = l20
Set Lights(21) = l21
Set Lights(22) = l22
Set Lights(23) = l23
Set Lights(24) = l24
Set Lights(25) = l25
Set Lights(26) = l26
Set Lights(27) = l27
Set Lights(28) = l28
Set Lights(29) = l29
Set Lights(30) = l30
Set Lights(31) = l31
Set Lights(32) = l32
Set Lights(33) = l33
Set Lights(34) = l34
Set Lights(35) = l35
Set Lights(36) = l36
Set Lights(37) = l37
Set Lights(38) = l38
Set Lights(39) = l39
Set Lights(40) = l40
Set Lights(41) = l41
Set Lights(42) = l42
Set Lights(43) = l43
Set Lights(44) = l44
Set Lights(45) = l45
Set Lights(46) = l46
Set Lights(47) = l47
Set Lights(48) = l48
Set Lights(49) = l49
Set Lights(50) = l50
Set Lights(51) = l51

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'           LL      EEEEEE  DDDD        ,,   SSSSS
'           LL      EE      DD  DD      ,,  SS
'           LL      EE      DD   DD      ,   SS
'           LL      EEEE    DD   DD            SS
'           LL      EE      DD  DD              SS
'           LLLLLL  EEEEEE  DDDD            SSSSS
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'       7 Digit Array
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Dim LED7(36)
LED7(0)=Array()
LED7(1)=Array()
LED7(2)=Array()
LED7(3)=Array()
LED7(4)=Array()
LED7(5)=Array()
LED7(6)=Array()
LED7(7)=Array()
LED7(8)=Array()
LED7(9)=Array()
LED7(10)=Array()
LED7(11)=Array()
LED7(12)=Array()
LED7(13)=Array()
LED7(14)=Array()
LED7(15)=Array()
LED7(16)=Array()
LED7(17)=Array()
LED7(18)=Array()
LED7(19)=Array()
LED7(20)=Array()
LED7(21)=Array()
LED7(22)=Array()
LED7(23)=Array()

LED7(24)=Array()
LED7(25)=Array()
LED7(26)=Array()
LED7(27)=Array()

LED7(28)=Array()
LED7(29)=Array()
LED7(30)=Array()
LED7(31)=Array()
LED7(32)=Array()
LED7(33)=Array()
LED7(34)=Array()
LED7(35)=Array()

Sub DisplayTimer_Timer
    Dim ChgLED, II, Num, Chg, Stat, Obj
    ChgLED = Controller.ChangedLEDs(0, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        For II = 0 To UBound(ChgLED)
            Num = ChgLED(II, 0):Chg = ChgLED(II, 1):Stat = ChgLED(II, 2)
            If Num > 35 Then
                For Each Obj In LED7(Num)
                    If Chg And 1 Then Obj.State = Stat And 1

                    Chg = Chg \ 2:Stat = Stat \ 2
                Next
            End If
        Next
    End If
End Sub

'*****************************************
'           BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3)

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
    ' render the shadow for each ball
   For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' + 13
       Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' - 13
       End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'Gottlieb Pink Panther
'added by Inkochnito
'Added coins chute by Mike da Spike
Sub editDips
    Dim vpmDips : Set vpmDips = New cvpmDips
    With vpmDips
       .AddForm 700,400,"Pink Panther - DIP switches"
       .AddFrame 2,10,190,"Coin Chute 1 (Coins/Credit)",&H0000000F,Array("2/1",&H00000008,"1/1",&H00000000,"1/2",&H00000009) 'Dip 1-4
       .AddFrame 2,70,190,"Coin Chute 2 (Coins/Credit)",&H000000F0,Array("2/1",&H00000080,"1/1",&H00000000,"1/2",&H00000090) 'Dip 5-8
       .AddFrame 2,130,190,"Coin Chute 3 (Coins/Credit)",&H00000F00,Array("2/1",&H00000800,"1/1",&H00000000,"1/2",&H00000900) 'Dip 9-12
       .AddFrame 2,190,190,"Coin Chute 3 extra credits",&H00001000,Array("no effect",0,"add 9",&H00001000)'dip 13
       .AddFrame 207,10,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
       .AddFrame 207,86,190,"Coin chute 1 and 2 control",&H00002000,Array("Seperate",0,"Same",&H00002000)'dip 14
       .AddFrame 207,132,190,"Playfield special",&H00200000,Array("Replay",0,"Extra Ball",&H00200000)'dip 22
       .AddFrame 207,178,190,"Maximum blue diamond total",&H80000000,Array("maximum 40",0,"maximum 50",&H80000000)'dip32
       .AddFrame 207,224,190,"High score to date awards",&H00C00000,Array("Not displayed and no award",0,"Displayed and no award",&H00800000,"Displayed and 2 replays",&H00400000,"Displayed and 3 replays",&H00C00000)'dip 23&24
       .AddChk 2,300,190,Array("Sound when scoring?",&H01000000)'dip 25
       .AddChk 2,315,190,Array("Replay button tune?",&H02000000)'dip 26
       .AddChk 2,330,190,Array("Coin switch tune?",&H04000000)'dip 27
       .AddChk 2,345,190,Array("Credits displayed?",&H08000000)'dip 28
       .AddChk 2,360,190,Array("Match feature",&H00020000)'dip 18
       .AddChk 2,375,190,Array("Attract features",&H20000000)'dip 30
       .AddFrameExtra 412,10,190,"Attract tune",&H0200,Array("No attract tune",0,"attract tune played every 5 minutes",&H0200)'S-board dip 2
       .AddFrame 412,56,190,"Balls per game",&H00010000,Array("5 balls",0,"3 balls",&H00010000)'dip 17
       .AddFrame 412,102,190,"Replay limit",&H00040000,Array("No limit",0,"One per ball",&H00040000)'dip 19
       .AddFrame 412,148,190,"Novelty",&H00080000,Array("Normal",0,"Points",&H0080000)'dip 20
       .AddFrame 412,194,190,"Game mode",&H00100000,Array("Replay",0,"Extra ball",&H00100000)'dip 21
       .AddFrame 412,240,190,"Tilt penalty",&H10000000,Array("Game over",0,"Ball in play",&H10000000)'dip 29
       .AddFrame 412,286,190,"Playfield special adjust",&H40000000,Array("On 20% longer than conservative",0,"Conservative",&H40000000)'dip 31
       .AddLabel 50,400,300,20,"After hitting OK, press F3 to reset game with new settings."
    End With
    Dim extra
    extra = Controller.Dip(4) + Controller.Dip(5)*256
    extra = vpmDips.ViewDipsExtra(extra)
    Controller.Dip(4) = extra And 255
    Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")

  Sub table1_Paused:Controller.Pause = 1:End Sub
  Sub table1_unPaused:Controller.Pause = 0:End Sub
  Sub table1_Exit:SaveLUT:Controller.Stop:End Sub

'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 1.1 beta2 (More proper behaviour, extra safety against script errors)
'*************************************************
Function NullFunction(aEnabled):End Function    '1 argument null function placeholder
Class cFastFlips
    Public TiltObjects, DebugOn, hi
    Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name, FlipState(3)

    Private Sub Class_Initialize()
        Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
        Set SubL = GetRef("NullFunction"): Set SubR = GetRef("NullFunction") : Set SubUL = GetRef("NullFunction"): Set SubUR = GetRef("NullFunction")
    End Sub

    'set callbacks
    Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : Decouple sLLFlipper, aInput: End Property
    Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
    Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : Decouple sLRFlipper, aInput:  End Property
    Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
    Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub   'Create Delay
    'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
    Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub

    'call callbacks
    Public Sub FlipL(aEnabled)
        FlipState(0) = aEnabled 'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
        If not FlippersEnabled and not DebugOn then Exit Sub
        subL aEnabled
    End Sub

    Public Sub FlipR(aEnabled)
        FlipState(1) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subR aEnabled
    End Sub

    Public Sub FlipUL(aEnabled)
        FlipState(2) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subUL aEnabled
    End Sub

    Public Sub FlipUR(aEnabled)
        FlipState(3) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subUR aEnabled
    End Sub

    Public Sub TiltSol(aEnabled)    'Handle solenoid / Delay (if delayinit)
        If delay > 0 and not aEnabled then  'handle delay
            vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
            LagCompensation = True
        else
            If Delay > 0 then LagCompensation = False
            EnableFlippers(aEnabled)
        end If
    End Sub

    Sub FireDelay() : If LagCompensation then EnableFlippers False End If : End Sub

    Private Sub EnableFlippers(aEnabled)
        If aEnabled then SubL FlipState(0) : SubR FlipState(1) : subUL FlipState(2) : subUR FlipState(3)
        FlippersEnabled = aEnabled
        If TiltObjects then vpmnudge.solgameon aEnabled
        If Not aEnabled then
            subL False
            subR False
            If not IsEmpty(subUL) then subUL False
            If not IsEmpty(subUR) then subUR False
        End If
    End Sub


    End Class

'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF, RF1)
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

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity
dim RF1 : Set RF1 = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF, RF1)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    x.TimeDelay = 44
  Next

  '"Polarity" Profile
'"Polarity" Profile<br>
  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.1, 0
  AddPt "Polarity", 2, 0.14, -2.25
  AddPt "Polarity", 3, 0.2, -2.25
  AddPt "Polarity", 4, 0.28, -3.25
  AddPt "Polarity", 5, 0.31, -3.25
  AddPt "Polarity", 6, 0.34, -3.75
  AddPt "Polarity", 7, 0.37, -3.75
  AddPt "Polarity", 8, 0.4, -4.5
  AddPt "Polarity", 9, 0.45, -3.5
  AddPt "Polarity", 10, 0.48, -3.5
  AddPt "Polarity", 11, 0.51, -3.75
  AddPt "Polarity", 12, 0.55, -3.75
  AddPt "Polarity", 13, 0.58, -3
  AddPt "Polarity", 14, 0.6, -2.75
  AddPt "Polarity", 15, 0.62, -2.75
  AddPt "Polarity", 16, 0.65, -2.5
  AddPt "Polarity", 17, 0.8, -2
  AddPt "Polarity", 18, 0.85, -1.9
  AddPt "Polarity", 19, 1.0, -1
  AddPt "Polarity", 20, 1.2, 0


  '"Velocity" Profile
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
  RF1.Object = RightFlipper1
  RF1.EndPoint = EndPointRp1
End Sub

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub
Sub TriggerRF1_Hit() : RF1.Addball activeball : End Sub
Sub TriggerRF1_UnHit() : RF1.PolarityCorrect activeball : End Sub

Sub RDampen_Timer()
Cor.Update
End Sub

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

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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

'******************************************************************************************
' CODE BELOW IS FOR THE VR CLOCK.  This was originally taken from Rascal VP9 clock table
'******************************************************************************************

Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()

    'ClockHands Below *********************************************************
  VRClockMinutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  VRClockHours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  VRClockSeconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

  VRClockMinutes001.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  VRClockHours001.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  VRClockSeconds001.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

End Sub

'******************************
' Setup Backglass
'******************************

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5, zoff,xrot,zscale, xcen,ycen

Sub setup_backglass()

  xoff = -20
  yoff1 = 60 ' this is where you adjust the forward/backward position for player 1 score
  yoff2 = 53 ' this is where you adjust the forward/backward position for player 2 score
  yoff3 = 47 ' this is where you adjust the forward/backward position for player 3 score
  yoff4 = 41 ' this is where you adjust the forward/backward position for player 4 score
  yoff5 = 60 ' this is where you adjust the forward/backward position for the bonus score
  zoff = 699
  xrot = -90

  center_digits()

end sub

Sub center_digits()
  Dim ix, xx, yy, yfact, xfact, xobj

  zscale = 0.0000001

  xcen = (130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

  for ix = 0 to 5
    For Each xobj In Digits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff1

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 6 to 11
    For Each xobj In Digits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff2

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 12 to 17
    For Each xobj In Digits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff3

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 18 to 23
    For Each xobj In Digits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff4

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 28 to 31
    For Each xobj In Digits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff5

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

end sub

'********************************************
'              Display Output
'********************************************

Dim Digits(32)
' Player 1 to 4
Digits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6,led1x7,led1x8)
Digits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6,led2x7,led2x8)
Digits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6,led3x7,led3x8)
Digits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6,led4x7,led4x8)
Digits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6,led5x7,led5x8)
Digits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6,led6x7,led6x8)

Digits(6) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6,led8x7,led8x8)
Digits(7) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6,led9x7,led9x8)
Digits(8) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6,led10x7,led10x8)
Digits(9) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6,led11x7,led11x8)
Digits(10) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6,led12x7,led12x8)
Digits(11) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6,led13x7,led13x8)

Digits(12) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006,LED1x007,LED1x008)
Digits(13) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106,LED1x107,LED1x108)
Digits(14) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206,LED1x207,LED1x208)
Digits(15) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306,LED1x307,LED1x308)
Digits(16) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406,LED1x407,LED1x408)
Digits(17) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506,LED1x507,LED1x508)

Digits(18) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006,led2x007,led2x008)
Digits(19) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106,led2x107,led2x108)
Digits(20) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206,led2x207,led2x208)
Digits(21) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306,led2x307,led2x308)
Digits(22) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406,led2x407,led2x408)
Digits(23) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506,led2x507,led2x508)

' LED on apron (credit -- Ball In Play -- Match)
Digits(24) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
Digits(25) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)
Digits(26) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
Digits(27) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)

' Player's Total -- Total's To Beat
Digits(28)= Array(LED3x000,LED3x001,LED3x002,LED3x003,LED3x004,LED3x005,LED3x006,led3x007,led3x008)
Digits(29)= Array(LED4x000,LED4x001,LED4x002,LED4x003,LED4x004,LED4x005,LED4x006,led4x007,led4x008)
Digits(30)= Array(LED5x000,LED5x001,LED5x002,LED5x003,LED5x004,LED5x005,LED5x006,LED5x007,LED5x008)
Digits(31)= Array(LED6x000,LED6x001,LED6x002,LED6x003,LED6x004,LED6x005,LED6x006,LED6x007,LED6x008)

dim DisplayColor
DisplayColor =  RGB(1,40,255)

Sub BGDisplayTimer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      If num <32 Then
              For Each obj In Digits(num)
'              If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
         If chg And 1 Then FadeDisplay obj, stat And 1
                   chg=chg\2 : stat=stat\2
              Next
      End If
        Next
    End If
 End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 12
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 6
  End If
End Sub


Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(Digits)
    if IsArray(Digits(x) ) then
      For each obj in Digits(x)
        obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub

InitDigits

' #####  COMMENTS  #####
' This the section to add if you have lamps called out on the backglass
' ######################

' ******************************************************************************************
'      LAMP CALLBACK for the backglass flasher lamps (not the solenoid conrolled ones)
' ******************************************************************************************


Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps()

    vpmNudge.solGameOn (Controller.Lamp(0) and not Controller.Lamp(1))
    if l8.state=1 then bsLLock.SolOut True
    if l9.state=1 then bsRLock.SolOut True
    L20a.State=L20.State
    L21a.State=L21.State
    L22a.State=L22.State

  If Controller.Lamp(1) = 0 Then: f_Tilt.visible=0: else: f_Tilt.visible=1  'Tilt
  If Controller.Lamp(3) = 0 Then: f_SA.visible=0: else: f_SA.visible=1  'Shoot Again
  If Controller.Lamp(10) = 0 Then: f_HS.visible=0: else: f_HS.visible=1 'Highscore
  If Controller.Lamp(11) = 0 Then: f_GO.visible=0: else: f_GO.visible=1: DisableFlyerSelector = 0 'Game Over

End Sub

'******************************************************
'           LUT
'******************************************************

Sub SetLUT
  Table1.ColorGradeImage = "LUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
  LUTBack.visible = 0
  VRLutdesc.visible = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  LUTBack.visible = 1
  VRLutdesc.visible = 1

  Select Case LUTSet
        Case 0: LUTBox.text = "Fleep Natural Dark 1": VRLUTdesc.imageA = "LUTcase0"
    Case 1: LUTBox.text = "Fleep Natural Dark 2": VRLUTdesc.imageA = "LUTcase1"
    Case 2: LUTBox.text = "Fleep Warm Dark": VRLUTdesc.imageA = "LUTcase2"
    Case 3: LUTBox.text = "Fleep Warm Bright": VRLUTdesc.imageA = "LUTcase3"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft": VRLUTdesc.imageA = "LUTcase4"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard": VRLUTdesc.imageA = "LUTcase5"
    Case 6: LUTBox.text = "Skitso Natural and Balanced": VRLUTdesc.imageA = "LUTcase6"
    Case 7: LUTBox.text = "Skitso Natural High Contrast": VRLUTdesc.imageA = "LUTcase7"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard": VRLUTdesc.imageA = "LUTcase8"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast": VRLUTdesc.imageA = "LUTcase9"
    Case 10: LUTBox.text = "HauntFreaks Desaturated" : VRLUTdesc.imageA = "LUTcase10"
      Case 11: LUTBox.text = "Tomate washed out": VRLUTdesc.imageA = "LUTcase11"
    Case 12: LUTBox.text = "VPW original 1on1": VRLUTdesc.imageA = "LUTcase12"
    Case 13: LUTBox.text = "bassgeige": VRLUTdesc.imageA = "LUTcase13"
    Case 14: LUTBox.text = "blacklight": VRLUTdesc.imageA = "LUTcase14"
    Case 15: LUTBox.text = "B&W Comic Book": VRLUTdesc.imageA = "LUTcase15"
  End Select

  LUTBox.TimerEnabled = 1

End Sub

Sub SaveLUT

  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if

  if LUTset = "" then LUTset = 12 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "PPLUT.txt",True)
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing

End Sub

Sub LoadLUT

  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    LUTset=12
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "PPLUT.txt") then
    LUTset=12
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "PPLUT.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      LUTset=12
      Exit Sub
    End if
    LUTset = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

Sub Flyer_Timer_Timer()
  If FlyerSet = 0 Then
      VR_Room_Info_sheet.Image = "Flyer1"
  End If
  If FlyerSet = 1 Then
      VR_Room_Info_sheet.Image = "Flyer2"
  End If
  If FlyerSet = 2 Then
      VR_Room_Info_sheet.Image = "Flyer3"
  End If
  If FlyerSet = 3 Then
      VR_Room_Info_sheet.Image = "Flyer4"
  End If
  If FlyerSet > 3 Then
    FlyerSet = 0
  End If
  If FlyerSet < 0 Then
    FlyerSet = 3
  End If
End Sub

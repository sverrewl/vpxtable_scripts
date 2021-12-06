'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######                               FirePower                                     ########
'#######                            (Williams 1980)                                  ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.1 WED21 2017

Option Explicit
Randomize

'
' Thanks To:
' GTXJoe for the Primative Collection
' Walamab for script help from his FirePower
' UncleReamus and Noah Fentz for some images from their VP9 FirePower
' Hauntfreaks for help in the planet redraw!
' Flupper for the physics starting point
' Every other author for their amazing work!

'Version 1.1 Changes
'Changed Top Left Plastic to uncolidable to fix ball getting stuck on it issue (not sure how that slipped by)
'Used invisible ramp on curve to help guide ball to top kickout
'Adjusted inserts lighting to match playfield better
'Made Ball Release Gate visible
'Added Ninuzzo's Ball Shadow

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2, reverted - didn't work
' Followed suggestion from Wob, changed to system7 and re-applied useSolenoid=2
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
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="frpwr_c7",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"
' Thal : Added because of useSolenoid=2
Const cSingleLFlip = 0
Const cSingleRFlip = 0

Const BallSize = 52
Const Ballmass = 2

LoadVPM "01560000", "S7.VBS", 3.26

Dim T1Step,T2Step,T3Step,T4Step,T5Step,T6Step,topPWRStep,midPWRStep,BtmPWRStep,CTStep
Dim trTrough

'---------------------------------------------------
' Set up consts with easy names for solenoid numbers
Const cBallRelease = 1
Const cLeftEjectHole = 4
Const cRightEjectHole = 5
Const cUpperEjectHole = 6
Const cBallSaveKick = 7
Const cBallRampThrower = 8
Const cCreditKnocker = 14
Const cFlashLamps = 15
Const cTopLeftBumper = 17
Const cBottomLeftBumper = 18
Const cTopRightBumper = 19
Const cBottomRightBumper = 20
Const cRightSlingShot = 21
Const cLeftSlingShot = 22
'---------------------------------------------------

'---------------------------------------------------
' Setup Consts with easy names for switch numbers
Const cOutHoleSW = 9
Const cLeftOutsideRolloverSW = 10
Const cLeftInsideRolloverSW = 11
Const cLeftKickerSW = 12
Const cLeftEjectHoleSW = 13
Const cUpperMiddleLeftStandupSW = 14
Const cSpinnerSW = 15
Const cTopLeftStandupSW = 16
Const cTarget1SW = 17
Const cTarget2SW = 18
Const cTarget3SW = 19
Const cTarget4SW = 21
Const cTarget5SW = 22
Const cTarget6SW = 23
Const cBottomLeftBumperSW = 25
Const cTopLeftBumperSW = 26
Const cTopRightBumperSW = 27
Const cBottomRightBumperSW = 28
Const cTopCenterTargetSW = 29
Const cRightEjectHoleSW = 30
Const cUpperTopRightStandupSW = 31
Const cFRolloverSW = 32
Const cIRolloverSW = 33
Const cRRolloverSW = 34
Const cERolloverSW = 35
Const cUpperRightEjectHoleSW = 36
Const cLowerTopRightStnadupSW = 37
Const cMIddleRightStandupSW = 38
Const cTopPOWERTargetSW = 39
Const cMiddlePOWERTargetSW = 40
Const cBottomPOWERTargetSW = 41
Const cRightKickerSW = 42
Const cRightInsideRolloverSW = 43
Const cRightOutsideRolloverSW = 44
Const cRightFlipperSW = 45
Const cBallShooterSW = 46
Const cPlayfieldTiltSW = 47
Const cLowerRightStandupSW = 48
Const cCenterMiddleLeftStandupSW = 49
Const cLowerMiddleLeftStuandupSW = 50
Const cLeftBallRampSW = 51
Const cLeftEjectRolloverSW = 53
Const cRightEjectRolloverSW = 54
Const cRightBallRampSW = 57
Const cCenterBallRampSW = 58
'--------------------------------------------------



'Solenoids Setup
SolCallback(cBallRelease) = "trTrough.SolIn"
SolCallback(cLeftEjectHole) = "LeftEjectHole"
SolCallback(cRightEjectHole) = "RightEjectHole"
SolCallback(cUpperEjectHole) = "UpperEjectHole"
SolCallback(cBallSaveKick) = "BallSaveKick"
SolCallback(cBallRampThrower) = "trTrough.SolOut"
SolCallback(cFlashLamps) = "Flashers"
SolCallback(cCreditKnocker) = "CreditKnocker"
SolCallback(cRightSlingShot) = "sRightSlingShot"
SolCallback(cLeftSlingShot) = "sLeftSlingShot"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"



'------------------------------------------------------------------
' Setup Solenoid Subs
Sub LeftEjectHole(enabled)
  if enabled Then
    KLeftEjectHole.Kick 180, 15
    controller.Switch(cLeftEjectHoleSW)=0
    PlaysoundAtVol "Popper_ball", KLeftEjectHole, VolKick
  end If
End Sub

Sub RightEjectHole(enabled)
  if enabled Then
    KRightEjectHole.Kick 180,15
    Controller.Switch(cRightEjectHoleSW)=0
    PlaysoundAtVol "Popper_ball", KRightEjectHole, VolKick
  end If
End Sub

Sub UpperEjectHole(enabled)
  if enabled Then
    KUpperEjectHole.Kick -90,15
    Controller.Switch(cUpperRightEjectHoleSW)=0
  PlaysoundAtVol "Popper_ball", KUpperEjectHole, VolKick
  end if
End Sub

Sub CreditKnocker(enabled)
  if enabled then
  Playsound "knocker",0, 1, 0.1, 0.5
  'DOF 14,2
  end if
end Sub



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot()
  vpmTimer.PulseSW cRightKickerSW
End Sub

Sub LeftSlingShot_Slingshot()
  vpmTimer.PulseSW cLeftKickerSW
End Sub

Sub sRightSlingShot(enabled)
  If enabled Then
    PlaySoundAtVol "slingshotRight", sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  End If
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub sLeftSlingShot(enabled)
  If enabled Then
    PlaySoundAtVol "slingshotLeft", sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub



'-----------------------------------------------------------------------
' Flipper Solenoid Handlers

sub SolRFlipper(enabled)
  if enabled Then
    PlaySoundAtVol "Fx_FlipperUp", RightFlipper, VolFlip
    RightFlipper.RotateToEnd

  Else
    PlaySoundAtVol "Fx_FlipperDown", RightFlipper, VolFlip
    RightFlipper.RotateToStart
  end If
end Sub


sub SolLFlipper(enabled)
  if enabled Then
    PlaySoundAtVol "Fx_FlipperUp", LeftFlipper, VolFlip
    LeftFlipper.RotateToEnd
  Else
    PlaySoundAtVol "Fx_FlipperDown", LeftFlipper, VolFlip
    LeftFlipper.RotateToStart
  end If
end Sub
'-----------------------------------------------------------------------
' Handle eject hole hit and unhit events and set corresponding switches
Sub KLeftEjectHole_Hit()
  controller.Switch(cLeftEjectHoleSW)=1
End Sub

Sub KLeftEjectHole_UnHit()
  controller.Switch(cLeftEjectHoleSW)=0
End Sub

Sub KRightEjectHole_Hit()
  controller.Switch(cRightEjectHoleSW)=1
End Sub

Sub KRightEjectHole_UnHit()
  controller.Switch(cRightEjectHoleSW)=0
End Sub

Sub KUpperEjectHole_Hit()
  controller.Switch(cUpperRightEjectHoleSW)=1
End Sub

Sub KUpperEjectHole_Unhit()
  controller.Switch(cUpperRightEjectHoleSW)=0
End Sub

Sub BallSaveKick(enabled)
  if enabled Then
    Playsound "slingshotleft", 0, .67, -0.05, 0.05 ' TODO
  End If

End Sub

Sub Flashers(enabled)
  If enabled Then
    LFire1.State = LightStateOn
    LFire2.State = LightStateOn
    LPower1.State = LightStateOn
    LPower2.State = LightStateOn
  Else
    LFire1.State = LightStateOff
    LFire2.State = LightStateOff
    LPower1.State = LightStateOff
    LPower2.State = LightStateOff
  End If
End Sub

'---------------------------------------------------------------------------------
'Initialize Table

Sub Table1_Init()
  vpmInit Me
  With Controller
      .GameName = cGameName
          If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
          .SplashInfoLine = "Firepower Williams 1980" & vbNewLine & "Created for VPX by WED21"
          .HandleKeyboard = 0
          .ShowTitle = 0
          .ShowFrame = 0
          .HandleMechanics = 0
      .ShowDMDOnly = 0
      .Hidden = 0
     .dip(0)=&h00  'Set to usa
          On Error Resume Next
         ' .Run GetPlayerHWnd
          If Err Then MsgBox Err.Description
          On Error Goto 0
      End With
    Controller.Run
PinMameTimer.enabled = 1
vpmMapLights AllLights


'----------------------------------------------
'Setup Ball Trough Object
Set trTrough = new cvpmTrough
trTrough.CreateEvents "trTrough", Array(Drain, BallRelease)
trTrough.balls = 3
trTrough.size = 3
trTrough.EntrySw = cOutHoleSW
trTrough.InitEntrySounds "drain", "drain", "drain"
trTrough.InitExitSounds "BallRelease", "BallRelease"
trTrough.addsw 2,cLeftBallRampSW
trTrough.addsw 1, cCenterBallRampSW
trTrough.addsw 0, cRightBallRampSW
trTrough.initexit BallRelease, 90, 10
trTrough.StackExitBalls = 1
trTrough.MaxBallsPerKick = 1
trTrough.Reset

'----------------------------------------------

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
  RailLeft.visible=1
  RailRight.visible=1
  DisplayTimer7.enabled = True
  P1D1.visible = 1
  P1D2.visible = 1
  P1D3.visible = 1
  P1D4.visible = 1
  P1D5.visible = 1
  P1D6.visible = 1
  P1D7.visible = 1
  P2D1.visible = 1
  P2D2.visible = 1
  P2D3.visible = 1
  P2D4.visible = 1
  P2D5.visible = 1
  P2D6.visible = 1
  P2D7.visible = 1
  P3D1.visible = 1
  P3D2.visible = 1
  P3D3.visible = 1
  P3D4.visible = 1
  P3D5.visible = 1
  P3D6.visible = 1
  P3D7.visible = 1
  P4D1.visible = 1
  P4D2.visible = 1
  P4D3.visible = 1
  P4D4.visible = 1
  P4D5.visible = 1
  P4D6.visible = 1
  P4D7.visible = 1
  BaD1.visible = 1
  BaD2.visible = 1
  CrD1.visible = 1
  CrD2.visible = 1

Else
  RailLeft.visible=0
  RailRight.visible=0
  DisplayTimer7.enabled = False
  P1D1.visible = 0
  P1D2.visible = 0
  P1D3.visible = 0
  P1D4.visible = 0
  P1D5.visible = 0
  P1D6.visible = 0
  P1D7.visible = 0
  P2D1.visible = 0
  P2D2.visible = 0
  P2D3.visible = 0
  P2D4.visible = 0
  P2D5.visible = 0
  P2D6.visible = 0
  P2D7.visible = 0
  P3D1.visible = 0
  P3D2.visible = 0
  P3D3.visible = 0
  P3D4.visible = 0
  P3D5.visible = 0
  P3D6.visible = 0
  P3D7.visible = 0
  P4D1.visible = 0
  P4D2.visible = 0
  P4D3.visible = 0
  P4D4.visible = 0
  P4D5.visible = 0
  P4D6.visible = 0
  P4D7.visible = 0
  BaD1.visible = 0
  BaD2.visible = 0
  CrD1.visible = 0
  CrD2.visible = 0
end If


end Sub


'-----------------------------------------------------------------
' Setup Switch Subs

Sub LeftOutsideRollover_hit()
  vpmtimer.pulsesw cLeftOutsideRolloverSW
  if LShieldOn.state = lightstateon Then
    KBallSaveKicker.enabled = 1
  Else
    KballsaveKicker.enabled = 0
  end if
End Sub


Sub LeftInsideRollover_hit:vpmtimer.pulsesw cLeftInsideRolloverSW:End Sub
Sub LeftInsideRollover_unhit:Controller.Switch(cLeftInsideRolloverSW)=0:End Sub



Sub UpperMiddleLeftStandup_hit:vpmtimer.pulsesw cUpperMiddleLeftStandupSW:PlaysoundAtVol "rubber_hit_3", UpperMiddleLeftStandup, 1:End Sub
Sub UpperMiddleLeftStandup_unhit:Controller.Switch(cUpperMiddleLeftStandupSW)=0:End Sub

Sub Spinner_Spin:vpmTimer.PulseSw cSpinnerSW:PlaySoundAtVol "fx_spinner",Spinner, VolSpin:End Sub

Sub TopLeftStandup_hit:vpmtimer.pulsesw cTopLeftStandupSW:PlaysoundAtVol "rubber_hit_3", cTopLeftStandupSW, 1:End Sub
Sub TopLeftStandup_unhit:Controller.Switch(cTopLeftStandupSW)=0:End Sub


'----------------
'TargetSubs
Sub Target1_hit()
  vpmTimer.PulseSw cTarget1SW
  PlaysoundAtVol "Target", Target1, VolTarg
End Sub


Sub Target2_hit()
  vpmTimer.PulseSw cTarget2SW
  PlaysoundAtVol "Target", Target2, VolTarg
End Sub

Sub Target3_hit()
  vpmTimer.PulseSw cTarget3SW
  PlaysoundAtVol "Target", Target3, VolTarg
End Sub


Sub Target4_hit()
  vpmTimer.PulseSw cTarget4SW
  PlaysoundAtVol "Target", Target4, VolTarg
End Sub


Sub Target5_hit()
  vpmTimer.PulseSw cTarget5SW
  PlaysoundAtVol "Target", Target5, VolTarg
End Sub

Sub Target6_hit()
  vpmTimer.PulseSw cTarget6SW
  PlaysoundAtVol "Target", Target6, VolTarg
End Sub


Sub TopPowerTarget_hit()
  vpmTimer.PulseSw cTopPOWERTargetSW
  PlaysoundAtVol "Target", TopPowerTarget, VolTarg
End Sub


Sub MIddlePOWERTarget_hit()
  vpmTimer.PulseSw cMiddlePOWERTargetSW
  PlaysoundAtVol "Target", MiddlePOWERTarget, VolTarg
End Sub


Sub BottomPOWERTarget_hit()
  vpmTimer.PulseSw cBottomPOWERTargetSW
  PlaysoundAtVol "Target", BottomPOWERTarget, VolTarg
End Sub


Sub TopCenterTarget_hit()
  vpmTimer.PulseSw  cTopCenterTargetSW
  PlaysoundAtVol "Target", TopCenterTarget, VolTarg
End Sub

Sub Bumper4_hit:vpmTimer.PulseSw cBottomLeftBumperSW:PlaysoundAtVol "Fx_Bumper1", Bumper4, VolBump:End Sub

Sub Bumper1_hit:vpmTimer.PulseSw cTopLeftBumperSW:PlaysoundAtVol "Fx_Bumper2", Bumper1, VolBump:End Sub

Sub Bumper2_hit:vpmTimer.PulseSw cTopRightBumperSW:PlaysoundAtVol "Fx_Bumper3", Bumper2, VolBump:End Sub

Sub Bumper3_hit:vpmTimer.PulseSw cBottomRightBumperSW:PlaysoundAtVol "Fx_Bumper4",Bumper3, VolBump:End Sub




Sub FRollover_hit:vpmtimer.pulsesw cFRolloverSW:End Sub
Sub FRollover_unhit:Controller.Switch(cFRolloverSW)=0:End Sub

Sub IRollover_hit:vpmtimer.pulsesw cIRolloverSW:End Sub
Sub IRollover_unhit:Controller.Switch(cIRolloverSW)=0:End Sub

Sub RRollover_hit:vpmtimer.pulsesw cRRolloverSW:End Sub
Sub RRollover_unhit:Controller.Switch(cRRolloverSW)=0:End Sub

Sub ERollover_hit:vpmtimer.pulsesw cERolloverSW:End Sub
Sub ERollover_unhit:Controller.Switch(cERolloverSW)=0:End Sub


Sub LowerTopRightStandup_hit:vpmtimer.pulsesw cLowerTopRightStnadupSW:PlaysoundAtVol "rubber_hit_3", LowerTopRightStandup, 1:End Sub
Sub LowerTopRightStandup_unhit:Controller.Switch(cLowerTopRightStnadupSW)=0:End Sub

Sub MiddleRightStandup_hit:vpmtimer.pulsesw cMIddleRightStandupSW:PlaysoundAtVol "rubber_hit_3", cMiddleRightStandupSW, 1:End Sub
Sub MiddleRightStandup_unhit:Controller.Switch(cMIddleRightStandupSW)=0:End Sub


Sub RightInsideRollover_hit:vpmtimer.pulsesw cRightInsideRolloverSw:End Sub
Sub RightInsideRollover_unhit:Controller.Switch(cRightInsideRolloverSW)=0:End Sub

Sub RightOutsideRollover_hit:vpmtimer.pulsesw cRightOutsideRolloverSW:End Sub
Sub RightOutsideRollover_unhit:Controller.Switch(cRightOutsideRolloverSW)=0:End Sub

Sub BallShooter_hit:Controller.Switch(cBallShooterSW)=1:End Sub
Sub BallShooter_unhit:Controller.Switch(cBallShooterSW)=0:End Sub

Sub LowerRightStandup_hit:vpmtimer.pulsesw cLowerRightStandupSW:PlaysoundAtVol "rubber_hit_3", cLowerRigthStandupSW, 1:End Sub
Sub LowerRightStandup_unhit:Controller.Switch(cLowerRightStandupSW)=0:End Sub


Sub CenterMIddleLeftStandup_hit:vpmtimer.pulsesw cCenterMiddleLeftStandupSW:PlaysoundAtVol "rubber_hit_3", cCenterMiddleLeftStandupSW, 1:End Sub
Sub CenterMIddleLeftStandup_unhit:Controller.Switch(cCenterMiddleLeftStandupSW)=0:End Sub

Sub LowerMIddleLeftStandup_hit:vpmtimer.pulsesw cLowerMiddleLeftStuandupSW:PlaysoundAtVol "rubber_hit_3", cLowerMiddleLeftStuandupSW, 1:End Sub
Sub LowerMIddleLeftStandup_unhit:Controller.Switch(cLowerMiddleLeftStuandupSW)=0:End Sub

Sub LeftEjectRollover_hit:vpmtimer.pulsesw cLeftEjectRolloverSW:End Sub
Sub LeftEjectRollover_unhit:Controller.Switch(cLeftEjectRolloverSW)=0:End Sub

Sub RightEjectRollover_hit:vpmtimer.pulsesw cRightEjectRolloverSW:End Sub
Sub RightEjectRollover_unhit:Controller.Switch(cRightEjectRolloverSW)=0:End Sub



'------------------------------------------------------------------
'Handle Keyboard Inputs

Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaySoundAtVol "plungerpull", Plunger, 1
  End If

  If keycode = 4 or keycode = 5 or keycode = 6 Then
    PlaysoundAtVol "fx_coin", drain, 1
  end if
    vpmKeyDown(keycode)
End Sub

Sub Table1_KeyUp(ByVal keycode)
  vpmKeyUp(keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAtVol "plunger", Plunger, 1
  End If

End Sub


'--------------------------------------------------------------
'*********BALLKICKER************

Sub KBallSaveKicker_Hit()

    Me.kick 0, 35
    Playsound "SlingshotLeft"

End Sub


'--------------------------------------------------------------
' Play sounds when gates are Hit
Sub Gate3_Hit()
  PlaysoundAtVol "Gate5", gate3, VolGates
End Sub

Sub Gate1_Hit()
  PlaysoundAtVol "Gate5", Gate1, VolGates
End Sub

Sub Gate2_Hit()
  PlaysoundAtVol "Gate5", Gate2, VolGates
End Sub

Sub BallReleaseGate_Hit()
  PlaysoundAtVol "Gate5", BallReleaseGate, 1
End Sub

Sub MetalPost12_Hit()
  PlaysoundAtVol "metalhit_medium", MetalPost12, 1
End Sub

Sub Ramp2_Hit()
  PlaysoundAtVol "metalhit_medium", Ramp2, 1
End Sub



 '=========================================================
'                    LED Handling
'=========================================================
'Modified version of Scapino's LED code for Fathom
'and borrowed from Uncle Willy's VP9 Firepower
'
Dim SixDigitOutput(32)
Dim SevenDigitOutput(32)
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

'Assign 7-digit output to reels
Set SevenDigitOutput(0)  = P1D7
Set SevenDigitOutput(1)  = P1D6
Set SevenDigitOutput(2)  = P1D5
Set SevenDigitOutput(3)  = P1D4
Set SevenDigitOutput(4)  = P1D3
Set SevenDigitOutput(5)  = P1D2
Set SevenDigitOutput(6)  = P1D1

Set SevenDigitOutput(7)  = P2D7
Set SevenDigitOutput(8)  = P2D6
Set SevenDigitOutput(9)  = P2D5
Set SevenDigitOutput(10) = P2D4
Set SevenDigitOutput(11) = P2D3
Set SevenDigitOutput(12) = P2D2
Set SevenDigitOutput(13) = P2D1

Set SevenDigitOutput(14) = P3D7
Set SevenDigitOutput(15) = P3D6
Set SevenDigitOutput(16) = P3D5
Set SevenDigitOutput(17) = P3D4
Set SevenDigitOutput(18) = P3D3
Set SevenDigitOutput(19) = P3D2
Set SevenDigitOutput(20) = P3D1

Set SevenDigitOutput(21) = P4D7
Set SevenDigitOutput(22) = P4D6
Set SevenDigitOutput(23) = P4D5
Set SevenDigitOutput(24) = P4D4
Set SevenDigitOutput(25) = P4D3
Set SevenDigitOutput(26) = P4D2
Set SevenDigitOutput(27) = P4D1

Set SevenDigitOutput(28) = CrD2
Set SevenDigitOutput(29) = CrD1
Set SevenDigitOutput(30) = BaD2
Set SevenDigitOutput(31) = BaD1

Sub DisplayTimer7_Timer ' 7-Digit output
  On Error Resume Next
  Dim ChgLED,ii,chg,stat,obj,TempCount,temptext,adj

  ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)

  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(ChgLED)
      chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      For TempCount = 0 to 10
        If stat = DisplayPatterns(TempCount) then
          If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
          DigStorage(chgLED(ii, 0)) = TempCount
        End If
        If stat = (DisplayPatterns(TempCount) + 128) then
          If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
          DigStorage(chgLED(ii, 0)) = TempCount
        End If
      Next
    Next
  End IF
End Sub


'*DOF method for non rom controller tables by Arngrim****************
'*******Use DOF 1**, 1 to activate a ledwiz output*******************
'*******Use DOF 1**, 0 to deactivate a ledwiz output*****************
'*******Use DOF 1**, 2 to pulse a ledwiz output**********************
Sub DOF(dofevent, dofstate)
  If B2SOn=True Then
    If dofstate = 2 Then
      Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
    Else
      Controller.B2SSetData dofevent, dofstate
    End If
  End If
End Sub
'********************************************************************
'*************SUPPORTING SOUNDS************************

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Switches_Hit (idx)
  PlaySound "metalhit_thin", 0, .25, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 20
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Const tnob = 4 ' total number of balls
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

